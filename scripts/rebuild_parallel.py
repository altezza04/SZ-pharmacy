#!/usr/bin/env python3
"""Parallel rebuild of pharmacies.json from official Excel using 3 ArcGIS geocoders."""
import hashlib
import json
import math
import re
import threading
import time
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import pandas as pd
from geopy.geocoders import ArcGIS

ROOT = Path(__file__).resolve().parents[1]
XLSX = ROOT / 'data' / 'pharmacies.xlsx'
OUT = ROOT / 'data' / 'pharmacies.json'
HOSP_OUT = ROOT / 'data' / 'hospitals.json'

# 官方地址 + 經核對的 POI 座標（優先醫院名稱 geocode / OSM，避免街道中心點誤判）
HOSPITALS = [
    {'id': 'h1', 'name': '北京大學深圳醫院', 'name_sc': '北京大学深圳医院', 'district': '福田區', 'address': '深圳市福田区莲花路1120号', 'lat': 22.559265, 'lng': 114.044095},
    {'id': 'h2', 'name': '深圳市人民醫院（留醫部）', 'name_sc': '深圳市人民医院（留医部）', 'district': '羅湖區', 'address': '深圳市罗湖区东门北路1017号', 'lat': 22.559836, 'lng': 114.122751},
    {'id': 'h3', 'name': '深圳市第二人民醫院', 'name_sc': '深圳市第二人民医院', 'district': '福田區', 'address': '深圳市福田区笋岗西路3002号', 'lat': 22.559792, 'lng': 114.080573},
    {'id': 'h4', 'name': '深圳市中醫院', 'name_sc': '深圳市中医院', 'district': '福田區', 'address': '深圳市福田区福华路1号', 'lat': 22.539755, 'lng': 114.080333},
    {'id': 'h5', 'name': '香港大學深圳醫院', 'name_sc': '香港大学深圳医院', 'district': '福田區', 'address': '深圳市福田区海园一路1号', 'lat': 22.528516, 'lng': 113.991192},
    {'id': 'h6', 'name': '深圳市兒童醫院', 'name_sc': '深圳市儿童医院', 'district': '福田區', 'address': '深圳市福田区益田路7019号', 'lat': 22.550092, 'lng': 114.049301},
    {'id': 'h7', 'name': '中國醫學科學院腫瘤醫院深圳醫院', 'name_sc': '中国医学科学院肿瘤医院深圳医院', 'district': '龍崗區', 'address': '深圳市龙岗区宝荷路113号', 'lat': 22.697156, 'lng': 114.246156},
    {'id': 'h8', 'name': '深圳市第三人民醫院', 'name_sc': '深圳市第三人民医院', 'district': '龍崗區', 'address': '深圳市龙岗区布澜路29号', 'lat': 22.638232, 'lng': 114.123374},
    {'id': 'h9', 'name': '深圳大學總醫院', 'name_sc': '深圳大学总医院', 'district': '南山區', 'address': '深圳市南山区学苑大道1098号', 'lat': 22.597126, 'lng': 113.982434},
    {'id': 'h10', 'name': '寶安區人民醫院', 'name_sc': '宝安区人民医院', 'district': '寶安區', 'address': '深圳市宝安区龙井二路118号', 'lat': 22.564596, 'lng': 113.909713},
    {'id': 'h11', 'name': '龍崗區中心醫院', 'name_sc': '龙岗区中心医院', 'district': '龍崗區', 'address': '深圳市龙岗区龙岗大道6082号', 'lat': 22.723037, 'lng': 114.242656},
    {'id': 'h12', 'name': '中山大學附屬第七醫院', 'name_sc': '中山大学附属第七医院', 'district': '光明區', 'address': '深圳市光明区圳园路628号', 'lat': 22.789493, 'lng': 113.947443},
]

N_WORKERS = 4
RATE_S = 1.1

SUBDISTRICT_RE = re.compile(r'([\u4e00-\u9fff]+街道)')
LANDMARK_RE = re.compile(r'([\u4e00-\u9fffA-Za-z0-9]+(?:大厦|大楼|广场|中心|花园|公寓|苑|城|酒店|商场|写字楼|会所))')
VILLAGE_RE = re.compile(r'([\u4e00-\u9fff]+村)')
ROAD_NUM_RE = re.compile(r'([\u4e00-\u9fff]{1,6}[路街道巷]\d+[-\d]*号?)')
BRANCH_RE = re.compile(r'([\u4e00-\u9fffA-Za-z0-9·]{2,24}(?:健康药房|大药房|药房|药店|分店|店))$')
COMPANY_SPLIT_RE = re.compile(r'(?:有限公司|连锁有限公司|医药连锁|药业连锁|连锁)')

SZ_LAT = (22.40, 22.95)
SZ_LNG = (113.70, 114.65)
DISTRICT_BBOX = {
    '福田区': (22.50, 22.59, 113.98, 114.11),
    '罗湖区': (22.52, 22.61, 114.08, 114.20),
    '南山区': (22.48, 22.61, 113.88, 114.05),
    '宝安区': (22.50, 22.82, 113.78, 114.06),
    '龙岗区': (22.58, 22.88, 114.05, 114.45),
    '龙华区': (22.60, 22.80, 113.98, 114.10),
    '坪山区': (22.62, 22.78, 114.28, 114.42),
    '光明区': (22.70, 22.85, 113.88, 114.05),
    '盐田区': (22.52, 22.66, 114.18, 114.35),
    '大鹏新区': (22.43, 22.65, 114.28, 114.65),
}

_worker_geos = [ArcGIS(timeout=20) for _ in range(N_WORKERS)]
_worker_locks = [threading.Lock() for _ in range(N_WORKERS)]
_worker_last = [0.0] * N_WORKERS
_shared_cache: dict = {}
_cache_lock = threading.Lock()


def in_shenzhen(lat, lng):
    return SZ_LAT[0] <= lat <= SZ_LAT[1] and SZ_LNG[0] <= lng <= SZ_LNG[1]


def in_district(lat, lng, district):
    box = DISTRICT_BBOX.get(district)
    if not box:
        return True
    return box[0] <= lat <= box[1] and box[2] <= lng <= box[3]


def raw_geocode(q: str, worker_idx: int):
    with _cache_lock:
        if q in _shared_cache:
            return _shared_cache[q]
    geo = _worker_geos[worker_idx]
    with _worker_locks[worker_idx]:
        elapsed = time.time() - _worker_last[worker_idx]
        if elapsed < RATE_S:
            time.sleep(RATE_S - elapsed)
        _worker_last[worker_idx] = time.time()
        loc = geo.geocode(q + ', 中国') or geo.geocode(q)
    if loc:
        result = (loc.latitude, loc.longitude, loc.address or '')
        with _cache_lock:
            _shared_cache[q] = result
        return result
    return None


def extract_landmarks(address, name):
    text = f'{address} {name}'
    found = []
    for pattern in (LANDMARK_RE, VILLAGE_RE):
        for m in pattern.findall(text):
            if len(m) >= 2 and m not in found:
                found.append(m)
    return found


def extract_shop_pois(name, address):
    pois = []
    if name:
        pois.append(name.strip())
        branch = BRANCH_RE.search(name.strip())
        if branch:
            full_branch = branch.group(1)
            if full_branch not in pois:
                pois.append(full_branch)
            short = re.sub(r'(健康药房|大药房|药房|药店|分店|店)$', '', full_branch)
            if len(short) >= 2:
                pois.append(short)
        parts = COMPANY_SPLIT_RE.split(name)
        if len(parts) > 1:
            tail = parts[-1].strip('（）() ')
            if len(tail) >= 2 and tail not in pois:
                pois.append(tail)
    for lm in extract_landmarks(address, name):
        if lm not in pois:
            pois.append(lm)
    return pois


def extract_subdistrict(address):
    m = SUBDISTRICT_RE.search(address)
    return m.group(1) if m else None


def extract_road_number(address):
    m = ROAD_NUM_RE.search(address.replace('､', ''))
    return m.group(1) if m else None


def score_result(query, result_addr, landmarks, district, pois=None):
    addr = result_addr or ''
    score = 0
    pois = pois or []
    if district in addr:
        score += 10
    else:
        for other in DISTRICT_BBOX:
            if other != district and other in addr:
                score -= 12
    for poi in pois:
        if len(poi) >= 2 and poi in addr:
            score += 12 if len(poi) >= 4 else 6
    if any(k in addr for k in ('药房', '药店', '医院', '大厦', '花园', '广场', '中心')):
        score += 6
    if re.search(r'\d+号', addr):
        score += 3
    sub = extract_subdistrict(query)
    if sub and sub in addr:
        score += 3
    for lm in landmarks:
        if lm in addr:
            score += 5
        elif lm in query and len(lm) >= 3:
            score += 1
    if '街道' in addr and not re.search(r'\d+号', addr) and not any(p in addr for p in pois):
        score -= 5
    if addr.count('区') >= 2 and '号' not in addr and not any(p in addr for p in pois[:3]):
        score -= 3
    return score


def geocode_address(name, address, district, worker_idx):
    addr = address.replace('广东省', '').strip()
    landmarks = extract_landmarks(addr, name)
    road_num = extract_road_number(addr)
    subdistrict = extract_subdistrict(addr)
    pois = extract_shop_pois(name, addr)

    candidates = []
    for poi in pois:
        candidates.append(f'{poi} 深圳市{district}')
        candidates.append(f'{poi} {district} 深圳')
        candidates.append(f'{poi} 深圳')
        if subdistrict:
            candidates.append(f'{poi} {subdistrict}')
    for poi in pois[:4]:
        candidates.append(f'{poi} {addr}')
    if subdistrict and road_num:
        candidates.append(f'{subdistrict}{road_num} 深圳市{district}')
    if road_num:
        candidates.append(f'深圳市{district}{road_num}')
        candidates.append(f'{road_num} 深圳市{district}')
    candidates.extend([addr, f'{name} {addr}'])

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
        hit = raw_geocode(q, worker_idx)
        if not hit:
            continue
        lat, lng, result_addr = hit
        score = score_result(q, result_addr, landmarks, district, pois)
        if not in_shenzhen(lat, lng) or not in_district(lat, lng, district):
            continue
        if score > best_score:
            best_score = score
            best = (lat, lng, score < 8)
        if score > fallback_score:
            fallback_score = score
            fallback = (lat, lng, True)

    return best or fallback


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


def spread_duplicate_pins(pharmacies, precision=5, min_radius_m=40, max_radius_m=180):
    groups = defaultdict(list)
    for p in pharmacies:
        key = (round(p['lat'], precision), round(p['lng'], precision))
        groups[key].append(p)
    spread_count = 0
    for group in groups.values():
        if len(group) <= 1:
            continue
        group = sorted(group, key=lambda x: (x['code'], x['id']))
        base_lat, base_lng = group[0]['lat'], group[0]['lng']
        n = len(group)
        for i, p in enumerate(group):
            digest = hashlib.sha256(f"{p['code']}|{p['address']}".encode()).hexdigest()
            h = int(digest[:8], 16)
            angle = (h % 360) * math.pi / 180 + (i * 2 * math.pi / n)
            radius_m = min_radius_m + (h % (max_radius_m - min_radius_m + 1))
            d_lat = (radius_m / 111_320) * math.sin(angle)
            d_lng = (radius_m / (111_320 * math.cos(math.radians(base_lat)))) * math.cos(angle)
            p['lat'] = round(base_lat + d_lat, 6)
            p['lng'] = round(base_lng + d_lng, 6)
            p['pinApprox'] = True
            spread_count += 1
    return spread_count


def geocode_row(args):
    idx, pid, code, name, address, district, worker_idx = args
    result = geocode_address(name, address, district, worker_idx)
    return idx, pid, code, name, address, district, result


def main():
    df = pd.read_excel(XLSX, sheet_name=0, header=1)
    df.columns = ['id', 'district', 'code', 'name', 'address']
    df = df.dropna(subset=['name'])
    print(f'Loaded {len(df)} pharmacies from Excel')

    # Keep researched hospital coordinates from hospitals.json
    hospitals_out = json.loads(HOSP_OUT.read_text(encoding='utf-8'))
    for h in hospitals_out:
        print(f"Hospital {h['id']}: {h['name']} -> {h['lat']:.5f},{h['lng']:.5f}")

    # Build task list
    tasks = []
    for i, (_, row) in enumerate(df.iterrows()):
        code = str(row['code']).strip()
        name = str(row['name']).strip()
        address = str(row['address']).strip()
        district = str(row['district']).strip()
        pid = str(int(row['id'])) if str(row['id']).replace('.', '', 1).isdigit() else str(row['id'])
        tasks.append((i, pid, code, name, address, district, i % N_WORKERS))

    print(f'\nGeocoding {len(tasks)} pharmacies with {N_WORKERS} workers...')
    results_map = {}
    failed = []
    done_count = [0]

    with ThreadPoolExecutor(max_workers=N_WORKERS) as ex:
        futs = {ex.submit(geocode_row, t): t for t in tasks}
        for f in as_completed(futs):
            idx, pid, code, name, address, district, result = f.result()
            done_count[0] += 1
            if done_count[0] % 100 == 0:
                print(f'  {done_count[0]}/{len(tasks)} done...')
            if result:
                results_map[idx] = (pid, code, name, address, district, result)
            else:
                failed.append(code)

    print(f'\nGeocoded {len(results_map)} pharmacies, failed {len(failed)}')
    if failed:
        print('Failed codes:', failed[:20])

    pharmacies = []
    for i, (_, row) in enumerate(df.iterrows()):
        if i not in results_map:
            continue
        pid, code, name, address, district, result = results_map[i]
        item = {
            'id': pid,
            'district': district,
            'code': code,
            'name': name,
            'address': address,
            'lat': round(result[0], 6),
            'lng': round(result[1], 6),
        }
        if len(result) > 2 and result[2]:
            item['pinApprox'] = True
        attach_hospital_ref(item, hospitals_out)
        pharmacies.append(item)

    spread = spread_duplicate_pins(pharmacies)
    print(f'Spread {spread} duplicate pins')

    OUT.write_text(json.dumps(pharmacies, ensure_ascii=False, indent=2), encoding='utf-8')
    HOSP_OUT.write_text(json.dumps(hospitals_out, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'\nWrote {len(pharmacies)} pharmacies -> {OUT}')
    print(f'Wrote {len(hospitals_out)} hospitals -> {HOSP_OUT}')


if __name__ == '__main__':
    main()
