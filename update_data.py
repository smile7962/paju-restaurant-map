# -*- coding: utf-8 -*-
"""
파주시 식품안심업소 웹앱 — 월별 데이터 갱신 스크립트
================================================================
매달 새 안심업소 엑셀을 넣고 실행하면:
  1) 통합 스키마로 정규화 + 주소 정제
     (위탁급식소·집단급식소는 일반인 이용 대상이 아니므로 자동 제외)
  2) 좌표 캐시(geocode_cache.json)를 조회해 '신규 주소만' 지오코딩
  3) V-World로 좌표 변환 (실패 시 선택적으로 카카오로 2차 시도)
  4) 좌표가 채워진 data.json 생성 + 캐시 갱신
  5) 리포트 출력 (신규/캐시적중/실패 건수, 실패 목록)

사용법:
  1) 아래 CONFIG의 VWORLD_KEY 확인
     (선택) KAKAO_REST_KEY에 카카오 REST 키 입력 → 실패 건 2차 시도
  2) 새 엑셀 파일명을 EXCEL_안심에 맞춤
  3) python update_data.py 실행
  4) 생성된 data.json을 웹앱과 같은 폴더에 배포(교체)

참고: 모범음식점 지정사업은 2026-07 관련 부서 요청으로 한시적 유예되어
      이 스크립트와 웹앱에서 제거되었습니다. (과거 처리 로직은 git 이력 참조)

필요 패키지:  pip install pandas openpyxl requests
"""
import pandas as pd
import re, json, os, time, sys
from datetime import datetime

try:
    import requests
except ImportError:
    requests = None  # 지오코딩 없이 정규화만 할 때는 없어도 됨

# ==================== CONFIG ====================
VWORLD_KEY   = "962880AD-3D17-3B2C-ABBA-BEED02A87F2C"   # 필수
KAKAO_REST_KEY = ""                          # 선택: 실패 건 카카오 2차 시도 (없으면 빈 문자열)

EXCEL_안심 = "식품안심업소_현황.xlsx"          # 매달 새 파일명으로 교체

기준일_안심 = "2026-06-30"                     # 엑셀 기준일 (헤더 표시용)

OUT_DATA   = "data.json"
OUT_CACHE  = "geocode_cache.json"
REQUEST_DELAY = 0.15   # API 호출 간 대기(초). 대량 처리 시 예의상 간격

# 일반인 이용 대상이 아니어서 제외할 업종/유형 키워드
EXCLUDE_업종 = ['위탁급식']
EXCLUDE_유형 = ['집단급식소']
# ================================================

# 파주시 권역 매핑 (읍면동 → 권역)
REGION = {
 '운정·교하권': ['동패동','와동동','목동동','야당동','다율동','문발동','산남동','상지석동',
             '교하동','당하동','오도동','서패동','연다산동','하지석동','송촌동','검산동','맥금동'],
 '금촌권': ['금촌동','아동동','금릉동'],
 '문산·파주권': ['문산읍','파주읍'],
 '탄현·조리·월롱권': ['탄현면','조리읍','월롱면'],
 '북부·기타권': ['광탄면','법원읍','적성면','파평면','군내면','장단면','진동면','진서면'],
}
EMD2REGION = {emd: reg for reg, emds in REGION.items() for emd in emds}

# 도로명 → 법정동 힌트 (도로명만 있고 동 표기가 없는 주소 보정용)
# 새로운 미상 주소가 나오면 여기에 한 줄씩 추가하세요.
ROAD_HINT = {
    '교하로': '교하동',
}

def extract_emd(addr):
    """주소에서 읍면동 추출. 읍/면 우선, 없으면 괄호 안 법정동(2글자+동)."""
    addr = str(addr)
    m = re.search(r'파주시\s+([가-힣]+[읍면])(?![가-힣])', addr)
    if m:
        return m.group(1)
    m = re.search(r'([가-힣]{2}동)(?![가-힣])', addr)
    if m:
        return m.group(1)
    # 도로명 힌트로 보정
    for road, emd in ROAD_HINT.items():
        if road in addr:
            return emd
    return '미상'

def clean_addr(addr):
    """괄호·층수·호수 정보를 제거해 지오코딩용 기본 주소만 남김."""
    a = str(addr).strip()
    cut = len(a)
    for ch in ['(', ',']:
        i = a.find(ch)
        if i != -1:
            cut = min(cut, i)
    return a[:cut].strip().rstrip('.')

def fmt_date(d):
    try:
        return pd.to_datetime(d).strftime('%Y-%m-%d')
    except Exception:
        return str(d)

# ---------- 1) 정규화 ----------
def normalize():
    df = pd.read_excel(EXCEL_안심)

    records, excluded = [], []
    for _, r in df.iterrows():
        업종 = str(r['업종']).strip()
        유형 = str(r['프랜차이즈/개별/집단급식소']).strip()
        업소명 = str(r['업소명(상호)']).strip()
        # 위탁급식소·집단급식소 제외 (일반인 이용 대상 아님)
        if any(k in 업종 for k in EXCLUDE_업종) or any(k in 유형 for k in EXCLUDE_유형):
            excluded.append(업소명)
            continue
        addr = str(r['영업장소재지']).strip()
        emd = extract_emd(addr)
        records.append({
            'id': f"A{int(r['연번']):03d}", '업소명': 업소명,
            '구분': '안심', '카테고리': 업종, '업태': None,
            '업종': 업종, '유형': 유형,
            '주메뉴': None, '전화번호': None, '주소': addr, '정제주소': clean_addr(addr),
            '권역': EMD2REGION.get(emd, '북부·기타권'), '읍면동': emd,
            '지정일자': fmt_date(r['지정일자']), '비고': None,
            'lat': None, 'lng': None, 'naver_query': f"{업소명} 파주",
        })
    return records, excluded

# ---------- 2) 지오코딩 ----------
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
    if requests is None:
        print("[경고] requests 미설치 → 지오코딩 없이 정규화만 수행합니다.")
    # 캐시 로드
    cache = {}
    if os.path.exists(OUT_CACHE):
        with open(OUT_CACHE, encoding='utf-8') as f:
            cache = json.load(f)

    records, excluded = normalize()

    stat = {'cache': 0, 'new': 0, 'fail': 0, 'kakao': 0}
    fails = []
    do_geo = requests is not None and VWORLD_KEY and "입력" not in VWORLD_KEY

    for x in records:
        addr = x['정제주소']
        if addr in cache:                      # 캐시 적중
            x['lat'], x['lng'] = cache[addr]['lat'], cache[addr]['lng']
            if x['lat'] is not None:
                stat['cache'] += 1
            continue
        if not do_geo:
            continue
        coord = geocode_vworld(addr)           # 신규만 V-World 호출
        src = 'vworld'
        if coord is None:                      # 실패 → 카카오 2차
            coord = geocode_kakao(addr)
            if coord is not None:
                src = 'kakao'; stat['kakao'] += 1
        time.sleep(REQUEST_DELAY)
        if coord is None:
            stat['fail'] += 1
            fails.append((x['업소명'], addr))
            cache[addr] = {'lat': None, 'lng': None, 'src': 'fail'}
        else:
            x['lat'], x['lng'] = coord
            stat['new'] += 1
            cache[addr] = {'lat': coord[0], 'lng': coord[1], 'src': src}

    # 저장
    meta = {
        '기준일_안심': 기준일_안심,
        '생성일': datetime.now().strftime('%Y-%m-%d'),
        '총업소수': len(records),
        '안심': len(records),
        '제외_급식소': len(excluded),
        '좌표없음': sum(1 for x in records if x['lat'] is None),
    }
    with open(OUT_DATA, 'w', encoding='utf-8') as f:
        json.dump({'meta': meta, 'restaurants': records}, f, ensure_ascii=False, indent=2)
    with open(OUT_CACHE, 'w', encoding='utf-8') as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)

    # 리포트
    print("=" * 48)
    print(f"총 {meta['총업소수']}개소  (급식소 제외 {meta['제외_급식소']}건)")
    if excluded:
        for n in excluded:
            print(f"  [제외] {n}")
    print(f"캐시적중 {stat['cache']}  신규지오코딩 {stat['new']}  카카오보정 {stat['kakao']}  실패 {stat['fail']}")
    print(f"좌표없음(지도 미표시) {meta['좌표없음']}개소")
    if fails:
        print("\n--- 지오코딩 실패 목록 (수동 보정 필요) ---")
        for n, a in fails:
            print(f"  {n}  |  {a}")
        print("\n[보정법] geocode_cache.json에서 해당 주소의 lat/lng를 직접 채우고 재실행하세요.")
    print("=" * 48)
    print(f"생성: {OUT_DATA}, {OUT_CACHE}")

if __name__ == "__main__":
    run()
