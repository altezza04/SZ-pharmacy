#!/usr/bin/env python3
"""Rebuild pharmacies.json from official Excel with ArcGIS geocoding."""
import json
import re
import time
from pathlib import Path

import pandas as pd
from geopy.geocoders import ArcGIS

ROOT = Path(__file__).resolve().parents[1]
XLSX = ROOT / 'data' / 'pharmacies.xlsx'
OUT = ROOT / 'data' / 'pharmacies.json'

# 深圳市重點醫院（官方名稱 + 地址，供 geocode）
HOSPITALS = [
    {'id': 'h1', 'name': '北京大學深圳醫院', 'name_sc': '北京大学深圳医院', 'district': '福田區', 'address': '深圳市福田区莲花路1120号'},
    {'id': 'h2', 'name': '深圳市人民醫院（留醫部）', 'name_sc': '深圳市人民医院（留医部）', 'district': '羅湖區', 'address': '深圳市罗湖区翠竹路1017号'},
    {'id': 'h3', 'name': '深圳市第二人民醫院', 'name_sc': '深圳市第二人民医院', 'district': '福田區', 'address': '深圳市福田区笋岗西路3002号'},
    {'id': 'h4', 'name': '深圳市中醫院', 'name_sc': '深圳市中医院', 'district': '福田區', 'address': '深圳市福田区福华路1号'},
    {'id': 'h5', 'name': '香港大學深圳醫院', 'name_sc': '香港大学深圳医院', 'district': '南山區', 'address': '深圳市南山区海园一路1号'},
    {'id': 'h6', 'name': '深圳市兒童醫院', 'name_sc': '深圳市儿童医院', 'district': '福田區', 'address': '深圳市福田区益田路7019号'},
    {'id': 'h7', 'name': '中國醫學科學院腫瘤醫院深圳醫院', 'name_sc': '中国医学科学院肿瘤医院深圳医院', 'district': '龍崗區', 'address': '深圳市龙岗区宝荷路113号'},
    {'id': 'h8', 'name': '深圳市第三人民醫院', 'name_sc': '深圳市第三人民医院', 'district': '龍崗區', 'address': '深圳市龙岗区布澜路29号'},
    {'id': 'h9', 'name': '深圳大學總醫院', 'name_sc': '深圳大学总医院', 'district': '南山區', 'address': '深圳市南山区学苑大道1098号'},
    {'id': 'h10', 'name': '寶安區人民醫院', 'name_sc': '宝安区人民医院', 'district': '寶安區', 'address': '深圳市宝安区龙井二路118号'},
    {'id': 'h11', 'name': '龍崗區中心醫院', 'name_sc': '龙岗区中心医院', 'district': '龍崗區', 'address': '深圳市龙岗区龙岗大道6082号'},
    {'id': 'h12', 'name': '中山大學附屬第七醫院', 'name_sc': '中山大学附属第七医院', 'district': '光明區', 'address': '深圳市光明区圳园路628号'},
]

SUBDISTRICT_RE = re.compile(r'([\u4e00-\u9fff]+街道)')
LANDMARK_RE = re.compile(
    r'([\u4e00-\u9fffA-Za-z0-9]+(?:大厦|大楼|广场|中心|花园|公寓|苑|城|酒店|商场|写字楼|会所|大厦))'
)
VILLAGE_RE = re.compile(r'([\u4e00-\u9fff]+村)')
ROAD_NUM_RE = re.compile(r'([\u4e00-\u9fff]{1,6}[路街道巷]\d+[-\d]*号?)')

def extract_subdistrict(address: str):
    match = SUBDISTRICT_RE.search(address)
    return match.group(1) if match else None


geocoder = ArcGIS(timeout=20)
_cache = {}

# 深圳市大致邊界（拒絕誤 geocode 到外省）
SZ_LAT = (22.40, 22.95)
SZ_LNG = (113.70, 114.65)

# 各區大致範圍（過濾跨區誤判，如寶安地址落到龍崗）
DISTRICT_BBOX = {
    '福田区': (22.50, 22.59, 113.98, 114.11),
    '罗湖区': (22.52, 22.61, 114.08, 114.20),
    '南山区': (22.48, 22.61, 113.88, 114.05),
    '宝安区': (22.50, 22.82, 113.78, 114.06),
    '龙岗区': (22.58, 22.88, 114.05, 114.45),
    '龙华区': (22.60, 22.80, 113.98, 114.10),
    '坪山区': (22.62, 22.78, 114.28, 114.42),
    '光明区': (22.70, 22.85, 113.88, 114.05),
}


def in_shenzhen(lat: float, lng: float) -> bool:
    return SZ_LAT[0] <= lat <= SZ_LAT[1] and SZ_LNG[0] <= lng <= SZ_LNG[1]


def in_district(lat: float, lng: float, district: str) -> bool:
    box = DISTRICT_BBOX.get(district)
    if not box:
        return True
    lat_min, lat_max, lng_min, lng_max = box
    return lat_min <= lat <= lat_max and lng_min <= lng <= lng_max


def geocode(query: str):
    q = query.strip()
    if not q:
        return None
    if q in _cache:
        return _cache[q]
    time.sleep(1.05)
    loc = geocoder.geocode(q + ', 中国')
    if not loc:
        loc = geocoder.geocode(q)
    if loc:
        _cache[q] = (loc.latitude, loc.longitude, loc.address or '')
        return _cache[q]
    return None


def extract_landmarks(address: str, name: str):
    text = f'{address} {name}'
    found = []
    for pattern in (LANDMARK_RE, VILLAGE_RE):
        for match in pattern.findall(text):
            token = match.strip()
            if len(token) >= 2 and token not in found:
                found.append(token)
    return found


def extract_road_number(address: str):
    match = ROAD_NUM_RE.search(address.replace('､', ''))
    return match.group(1) if match else None


def score_result(query: str, result_addr: str, landmarks, district: str):
    """Higher is better; penalise generic or cross-district hits."""
    addr = result_addr or ''
    score = 0
    if district in addr:
        score += 10
    else:
        for other in DISTRICT_BBOX:
            if other != district and other in addr:
                score -= 12
    if re.search(r'\d+号', addr):
        score += 4
    sub = extract_subdistrict(query)
    if sub and sub in addr:
        score += 4
    for lm in landmarks:
        if lm in addr:
            score += 5
        elif lm in query and len(lm) >= 3:
            score += 1
    if '街道' in addr and not re.search(r'\d+号', addr) and not any(lm in addr for lm in landmarks):
        score -= 3
    if addr.count('区') >= 2 and '号' not in addr:
        score -= 2
    return score


def geocode_address(name: str, address: str, district: str):
    """Try progressively more specific queries; prefer building/street-level hits."""
    addr = address.replace('广东省', '').strip()
    landmarks = extract_landmarks(addr, name)
    road_num = extract_road_number(addr)
    subdistrict = extract_subdistrict(addr)

    candidates = []
    if subdistrict and road_num:
        candidates.append(f'{subdistrict}{road_num} 深圳市{district}')
        candidates.append(f'{subdistrict}{road_num} {district}')
    for lm in landmarks:
        candidates.append(f'{lm} {addr}')
        candidates.append(f'{lm} 深圳市{district}')
        candidates.append(f'{lm} 深圳')
    if road_num:
        candidates.append(f'深圳市{district}{road_num}')
        candidates.append(f'{road_num} 深圳市{district}')
        candidates.append(f'{road_num} 深圳')
    candidates.extend([
        addr,
        f'{name} {addr}',
        addr.replace('深圳市', f'深圳市{district}'),
        name,
    ])

    seen = set()
    best = None
    best_score = -999
    fallback = None
    fallback_score = -999
    for q in candidates:
        q = q.strip()
        if not q or q in seen:
            continue
        seen.add(q)
        hit = geocode(q)
        if not hit:
            continue
        lat, lng, result_addr = hit
        score = score_result(q, result_addr, landmarks, district)
        if not in_shenzhen(lat, lng) or not in_district(lat, lng, district):
            continue
        if score > best_score:
            best_score = score
            best = (lat, lng, score >= 4)
        if score > fallback_score:
            fallback_score = score
            fallback = (lat, lng, score < 4)
    if best:
        return best[0], best[1], best[2]
    if fallback:
        return fallback[0], fallback[1], True
    return None


def haversine_km(lat1, lon1, lat2, lon2):
    from math import asin, cos, radians, sin, sqrt
    r = 6371
    dlat = radians(lat2 - lat1)
    dlon = radians(lon2 - lon1)
    a = sin(dlat / 2) ** 2 + cos(radians(lat1)) * cos(radians(lat2)) * sin(dlon / 2) ** 2
    return 2 * r * asin(sqrt(a))


def attach_hospital_ref(item, hospitals_out):
    nearest = None
    nearest_km = 9999
    for h in hospitals_out:
        d = haversine_km(item['lat'], item['lng'], h['lat'], h['lng'])
        if d < nearest_km:
            nearest_km = d
            nearest = h
    item.pop('isNearHospital', None)
    item.pop('hospitalRef', None)
    if nearest and nearest_km <= 0.8:
        item['isNearHospital'] = True
        item['hospitalRef'] = nearest['name']


def fix_district_mismatches():
    pharmacies = json.loads(OUT.read_text(encoding='utf-8'))
    hospitals_out = json.loads((ROOT / 'data' / 'hospitals.json').read_text(encoding='utf-8'))
    fixed = 0
    for item in pharmacies:
        if in_district(item['lat'], item['lng'], item['district']):
            continue
        result = geocode_address(item['name'], item['address'], item['district'])
        if not result:
            print('STILL FAILED', item['code'], item['name'][:30])
            continue
        item['lat'] = round(result[0], 6)
        item['lng'] = round(result[1], 6)
        if result[2]:
            item['approx'] = True
        else:
            item.pop('approx', None)
        attach_hospital_ref(item, hospitals_out)
        fixed += 1
        print('FIXED', item['code'], item['lat'], item['lng'], item['name'][:30])
    OUT.write_text(json.dumps(pharmacies, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'Fixed {fixed} pharmacies')


def rebuild_pharmacies_only():
    hospitals_out = json.loads((ROOT / 'data' / 'hospitals.json').read_text(encoding='utf-8'))
    df = pd.read_excel(XLSX, sheet_name=0, header=1)
    df.columns = ['id', 'district', 'code', 'name', 'address']
    df = df.dropna(subset=['name'])

    pharmacies = []
    failed = []
    for _, row in df.iterrows():
        code = str(row['code']).strip()
        name = str(row['name']).strip()
        address = str(row['address']).strip()
        district = str(row['district']).strip()
        pid = str(int(row['id'])) if str(row['id']).replace('.', '', 1).isdigit() else str(row['id'])

        result = geocode_address(name, address, district)
        if not result:
            failed.append(code)
            continue
        item = {
            'id': pid,
            'district': district,
            'code': code,
            'name': name,
            'address': address,
            'lat': round(result[0], 6),
            'lng': round(result[1], 6),
        }
        if result[2]:
            item['approx'] = True
        attach_hospital_ref(item, hospitals_out)
        pharmacies.append(item)
        print(f"OK {pid:>3} {code} -> {name[:28]}")

    if failed:
        print('FAILED', failed)
    OUT.write_text(json.dumps(pharmacies, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'Wrote {len(pharmacies)} pharmacies -> {OUT}')


def main():
    df = pd.read_excel(XLSX, sheet_name=0, header=1)
    df.columns = ['id', 'district', 'code', 'name', 'address']
    df = df.dropna(subset=['name'])

    hospitals_out = []
    for h in HOSPITALS:
        pt = geocode(h['address']) or geocode(h['name_sc'])
        if not pt or not in_shenzhen(pt[0], pt[1]):
            raise SystemExit(f"Failed to geocode hospital: {h['name']}")
        hospitals_out.append({
            **h,
            'lat': round(pt[0], 6),
            'lng': round(pt[1], 6),
            'desc': h['address'].replace('深圳市', '').replace('广东省', ''),
        })

    pharmacies = []
    failed = []
    for _, row in df.iterrows():
        code = str(row['code']).strip()
        name = str(row['name']).strip()
        address = str(row['address']).strip()
        district = str(row['district']).strip()
        pid = str(int(row['id'])) if str(row['id']).replace('.', '', 1).isdigit() else str(row['id'])

        result = geocode_address(name, address, district)
        if not result:
            failed.append(code)
            continue
        pt = (result[0], result[1])
        approx = result[2] if len(result) > 2 else False

        item = {
            'id': pid,
            'district': district,
            'code': code,
            'name': name,
            'address': address,
            'lat': round(pt[0], 6),
            'lng': round(pt[1], 6),
        }
        if approx:
            item['approx'] = True

        nearest = None
        nearest_km = 9999
        for h in hospitals_out:
            d = haversine_km(pt[0], pt[1], h['lat'], h['lng'])
            if d < nearest_km:
                nearest_km = d
                nearest = h
        if nearest and nearest_km <= 0.8:
            item['isNearHospital'] = True
            item['hospitalRef'] = nearest['name']

        pharmacies.append(item)
        print(f"OK {pid:>3} {code} {nearest_km*1000:.0f}m -> {name[:28]}")

    if failed:
        print('FAILED', failed)

    OUT.write_text(json.dumps(pharmacies, ensure_ascii=False, indent=2), encoding='utf-8')
    hosp_path = ROOT / 'data' / 'hospitals.json'
    hosp_path.write_text(json.dumps(hospitals_out, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f"\nWrote {len(pharmacies)} pharmacies -> {OUT}")
    print(f"Wrote {len(hospitals_out)} hospitals -> {hosp_path}")


if __name__ == '__main__':
    import sys
    if '--fix-district' in sys.argv:
        fix_district_mismatches()
    elif '--pharmacies-only' in sys.argv:
        rebuild_pharmacies_only()
    else:
        main()
