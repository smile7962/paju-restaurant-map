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

REQUEST_DELAY = 0.15           # API 호출 간 대기(초)
# ================================================

# 스크립트가 어디서 실행되든(예: 스마트폰 Pydroid 3) 스크립트와 같은
# 폴더의 data.json / geocode_cache.json 을 찾도록 절대경로로 고정한다.
HERE   = os.path.dirname(os.path.abspath(__file__))
DATA   = os.path.join(HERE, "data.json")
CACHE  = os.path.join(HERE, "geocode_cache.json")


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


def diagnose(addr):
    """첫 주소로 V-World를 한 번 호출해 '전부 실패'의 원인을 구분해준다.
    - 네트워크 자체가 막히면(행정망/방화벽) → 예외 발생
    - 접속은 되는데 키 문제면 → status != OK + 원문 에러 메시지"""
    print("-" * 48)
    print(f"[진단] V-World 연결 확인 중 …  ({addr})")
    base = "https://api.vworld.kr/req/address"
    params = {
        "service": "address", "request": "getcoord", "version": "2.0",
        "crs": "epsg:4326", "type": "ROAD", "address": addr,
        "format": "json", "key": VWORLD_KEY,
    }
    try:
        r = requests.get(base, params=params, timeout=10)
    except Exception as e:
        print(f"  ✗ 네트워크 오류: {repr(e)[:120]}")
        print("  → 인터넷 접속 자체가 막혀 있습니다 (행정망/사내 방화벽).")
        print("    개인 인터넷(집 와이파이·개인 노트북·휴대폰 테더링)에서 실행하세요.")
        print("-" * 48)
        return False
    try:
        j = r.json()
    except Exception:
        print(f"  ✗ 응답이 JSON이 아님 (HTTP {r.status_code}). 프록시 차단 페이지일 수 있습니다.")
        print("-" * 48)
        return False
    st = j.get("response", {}).get("status")
    if st == "OK":
        print("  ✓ 연결·인증 정상. 지오코딩을 시작합니다.")
        print("-" * 48)
        return True
    print(f"  ✗ 접속은 됐지만 V-World가 거부: status={st}")
    print(f"    응답 원문: {json.dumps(j, ensure_ascii=False)[:400]}")
    print("  → 네트워크가 아니라 '인증키' 문제입니다. V-World 마이페이지에서")
    print("    이 키에 '주소검색(getcoord)' API가 신청돼 있는지, 등록 도메인을 확인하세요.")
    print("-" * 48)
    return False


def run():
    if not VWORLD_KEY or "입력" in VWORLD_KEY:
        print("[오류] VWORLD_KEY를 설정하세요.")
        sys.exit(1)

    if not os.path.exists(DATA):
        print(f"[오류] data.json 을 찾을 수 없습니다. 이 스크립트와 같은 폴더에 data.json이 있어야 합니다.\n       (찾은 위치: {DATA})")
        sys.exit(1)

    with open(DATA, encoding="utf-8") as f:
        payload = json.load(f)
    records = payload.get("restaurants", [])
    meta = payload.get("meta", {})

    cache = {}
    if os.path.exists(CACHE):
        with open(CACHE, encoding="utf-8") as f:
            cache = json.load(f)

    # 지오코딩이 필요한 첫 주소로 사전 진단 (전부 실패 원인 구분 + 캐시 오염 방지)
    need_geo = [x for x in records
                if x.get("lat") is None
                and (x.get("정제주소") or x.get("주소"))
                and (x.get("정제주소") or x.get("주소")) not in cache]
    if need_geo:
        probe = need_geo[0].get("정제주소") or need_geo[0].get("주소")
        if not diagnose(probe):
            print("사전 진단 실패 → 지오코딩을 중단합니다 (data.json·캐시는 그대로 둡니다).")
            sys.exit(1)

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
