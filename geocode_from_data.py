# -*- coding: utf-8 -*-
"""
파주 안심밥상 지도 — data.json 직접 지오코딩 스크립트
================================================================
이미 정규화가 끝난 data.json의 '정제주소'를 읽어 좌표(lat/lng)를 채웁니다.
원본 엑셀이 없어도 동작하며, update_data.py와 같은 캐시(geocode_cache.json)를
공유하므로 두 스크립트를 섞어 써도 됩니다.

동작:
  1) data.json 로드 (restaurants[].정제주소 사용)
  2) geocode_cache.json 조회 → '캐시에 없는 신규 주소만' 지오코딩
  3) V-World Geocoder 2.0 (도로명→지번 순), 실패 시 선택적으로 카카오 2차 시도
  4) 좌표가 채워진 data.json 덮어쓰기 + 캐시 갱신 + meta.좌표없음 재계산
  5) 리포트 출력 (캐시적중/신규/카카오보정/실패 + 실패 목록)

사용법:
  1) 아래 VWORLD_KEY 확인 (기본값은 이 프로젝트 키)
     (선택) KAKAO_REST_KEY 입력 → V-World 실패 건 2차 시도
  2) python geocode_from_data.py
  3) 갱신된 data.json을 웹앱과 같은 폴더에 배포(교체)

주의:
  - V-World 지오코딩 API(api.vworld.kr)에 아웃바운드 접속이 가능한 환경에서
    실행해야 합니다. (회사/클라우드 방화벽이 막으면 로컬 PC에서 실행하세요.)
  - 실패한 주소는 캐시에 lat=null로 기록됩니다. geocode_cache.json에서 해당
    주소의 lat/lng를 직접 채우고 재실행하면 반영됩니다.

필요 패키지:  pip install requests
"""
import json
import os
import time
import sys

try:
    import requests
except ImportError:
    print("[오류] requests 미설치 → pip install requests 후 재실행하세요.")
    sys.exit(1)

# ==================== CONFIG ====================
VWORLD_KEY     = "962880AD-3D17-3B2C-ABBA-BEED02A87F2C"  # 필수
KAKAO_REST_KEY = ""            # 선택: 실패 건 카카오 2차 시도 (없으면 빈 문자열)

DATA   = "data.json"
CACHE  = "geocode_cache.json"
REQUEST_DELAY = 0.15           # API 호출 간 대기(초)
# ================================================


def geocode_vworld(addr):
    """V-World Geocoder 2.0. 도로명(ROAD) 우선, 실패 시 지번(PARCEL)."""
    base = "https://api.vworld.kr/req/address"
    for typ in ("ROAD", "PARCEL"):
        params = {
            "service": "address", "request": "getcoord", "version": "2.0",
            "crs": "epsg:4326", "type": typ, "address": addr,
            "format": "json", "key": VWORLD_KEY,
        }
        try:
            r = requests.get(base, params=params, timeout=10)
            j = r.json()
            if j.get("response", {}).get("status") == "OK":
                p = j["response"]["result"]["point"]
                return float(p["y"]), float(p["x"])  # lat, lng
        except Exception:
            pass
    return None


def geocode_kakao(addr):
    """카카오 로컬 주소검색 (실패 건 2차 시도용)."""
    if not KAKAO_REST_KEY:
        return None
    try:
        r = requests.get(
            "https://dapi.kakao.com/v2/local/search/address.json",
            params={"query": addr},
            headers={"Authorization": f"KakaoAK {KAKAO_REST_KEY}"}, timeout=10)
        docs = r.json().get("documents", [])
        if docs:
            return float(docs[0]["y"]), float(docs[0]["x"])
    except Exception:
        pass
    return None


def run():
    if not VWORLD_KEY or "입력" in VWORLD_KEY:
        print("[오류] VWORLD_KEY를 설정하세요.")
        sys.exit(1)

    if not os.path.exists(DATA):
        print(f"[오류] {DATA} 를 찾을 수 없습니다. 웹앱과 같은 폴더에서 실행하세요.")
        sys.exit(1)

    with open(DATA, encoding="utf-8") as f:
        payload = json.load(f)
    records = payload.get("restaurants", [])
    meta = payload.get("meta", {})

    cache = {}
    if os.path.exists(CACHE):
        with open(CACHE, encoding="utf-8") as f:
            cache = json.load(f)

    stat = {"cache": 0, "new": 0, "fail": 0, "kakao": 0, "skip": 0}
    fails = []

    for x in records:
        # 이미 좌표가 있으면 건너뜀 (재실행 시 안전)
        if x.get("lat") is not None and x.get("lng") is not None:
            stat["skip"] += 1
            continue

        addr = x.get("정제주소") or x.get("주소")
        if not addr:
            stat["fail"] += 1
            fails.append((x.get("업소명", "?"), "(주소 없음)"))
            continue

        if addr in cache:                          # 캐시 적중
            c = cache[addr]
            x["lat"], x["lng"] = c.get("lat"), c.get("lng")
            if x["lat"] is not None:
                stat["cache"] += 1
            else:
                stat["fail"] += 1
                fails.append((x.get("업소명", "?"), addr))
            continue

        coord = geocode_vworld(addr)               # 신규만 V-World 호출
        src = "vworld"
        if coord is None:                          # 실패 → 카카오 2차
            coord = geocode_kakao(addr)
            if coord is not None:
                src = "kakao"
                stat["kakao"] += 1
        time.sleep(REQUEST_DELAY)

        if coord is None:
            stat["fail"] += 1
            fails.append((x.get("업소명", "?"), addr))
            cache[addr] = {"lat": None, "lng": None, "src": "fail"}
        else:
            x["lat"], x["lng"] = coord
            stat["new"] += 1
            cache[addr] = {"lat": coord[0], "lng": coord[1], "src": src}

    # meta.좌표없음 재계산
    meta["좌표없음"] = sum(1 for x in records if x.get("lat") is None)
    payload["meta"] = meta

    with open(DATA, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    with open(CACHE, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)

    # 리포트
    print("=" * 48)
    print(f"총 {len(records)}개소")
    print(f"이미좌표 {stat['skip']}  캐시적중 {stat['cache']}  "
          f"신규지오코딩 {stat['new']}  카카오보정 {stat['kakao']}  실패 {stat['fail']}")
    print(f"좌표없음(지도 미표시) {meta['좌표없음']}개소")
    if fails:
        print("\n--- 지오코딩 실패 목록 (수동 보정 필요) ---")
        for n, a in fails:
            print(f"  {n}  |  {a}")
        print("\n[보정법] geocode_cache.json에서 해당 주소의 lat/lng를 직접 채우고 재실행하세요.")
    print("=" * 48)
    print(f"갱신: {DATA}, {CACHE}")


if __name__ == "__main__":
    run()
