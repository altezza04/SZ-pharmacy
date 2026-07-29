#!/usr/bin/env python3
"""Geocode pharmacies + hospitals via Baidu Map POI → GCJ-02 (Amap-aligned).

Coordinates are stored as GCJ-02 so they line up with 高德 tiles / navigation.
"""
from __future__ import annotations

import json
import math
import re
import time
import urllib.parse
import urllib.request
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
XLSX = ROOT / 'data' / 'pharmacies.xlsx'
OUT = ROOT / 'data' / 'pharmacies.json'
HOSP_OUT = ROOT / 'data' / 'hospitals.json'

UA = (
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
    '(KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36'
)
BRANCH_RE = re.compile(r'([\u4e00-\u9fffA-Za-z0-9·]{2,24}(?:健康药房|大药房|药房|药店|分店|店))$')
LANDMARK_RE = re.compile(
    r'([\u4e00-\u9fffA-Za-z0-9]+(?:大厦|大楼|广场|中心|花园|公寓|苑|城|商场|会所|酒店|村|园))'
)
PHARM_NAME_RE = re.compile(r'药房|药店|大药房|医药')
NON_PHARM_RE = re.compile(r'酒店|邮局|邮政|停车场|肠粉|彩票|公交|地铁|银行|餐厅|咖啡|超市(?!.*药)')

SZ_LAT = (22.40, 22.95)
SZ_LNG = (113.70, 114.65)

# gcoord BD09MC tables
MCBAND = [12890594.86, 8362377.87, 5591021, 3481989.83, 1678043.12, 0]
MC2LL = [
    [1.410526172116255e-8, 0.00000898305509648872, -1.9939833816331, 200.9824383106796,
     -187.2403703815547, 91.6087516669843, -23.38765649603339, 2.57121317296198,
     -0.03801003308653, 17337981.2],
    [-7.435856389565537e-9, 0.000008983055097726239, -0.78625201886289, 96.32687599759846,
     -1.85204757529826, -59.36935905485877, 47.40033549296737, -16.50741931063887,
     2.28786674699375, 10260144.86],
    [-3.030883460898826e-8, 0.00000898305509983578, 0.30071316287616, 59.74293618442277,
     7.357984074871, -25.38371002664745, 13.45380521110908, -3.29883767235584,
     0.32710905363475, 6856817.37],
    [-1.981981304930552e-8, 0.000008983055099779535, 0.03278182852591, 40.31678527705744,
     0.65659298677277, -4.44255534477492, 0.85341911805263, 0.12923347998204,
     -0.04625736007561, 4482777.06],
    [3.09191371068437e-9, 0.000008983055096812155, 0.00006995724062, 23.10934304144901,
     -0.00023663490511, -0.6321817810242, -0.00663494467273, 0.03430082397953,
     -0.00466043876332, 2555164.4],
    [2.890871144776878e-9, 0.000008983055095805407, -3.068298e-8, 7.47137025468032,
     -0.00000353937994, -0.02145144861037, -0.00001234426596, 0.00010322952773,
     -0.00000323890364, 826088.5],
]
X_PI = math.pi * 3000.0 / 180.0


def get_json(url: str) -> dict:
    req = urllib.request.Request(url, headers={'User-Agent': UA, 'Referer': 'https://map.baidu.com/'})
    raw = urllib.request.urlopen(req, timeout=25).read().decode('utf-8', 'replace')
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return json.loads(re.sub(r'\\(?!["\\/bfnrtu])', r'\\\\', raw))


def transform_lat(x, y):
    ret = -100.0 + 2.0 * x + 3.0 * y + 0.2 * y * y + 0.1 * x * y + 0.2 * math.sqrt(abs(x))
    ret += (20.0 * math.sin(6.0 * x * math.pi) + 20.0 * math.sin(2.0 * x * math.pi)) * 2.0 / 3.0
    ret += (20.0 * math.sin(y * math.pi) + 40.0 * math.sin(y / 3.0 * math.pi)) * 2.0 / 3.0
    ret += (160.0 * math.sin(y / 12.0 * math.pi) + 320 * math.sin(y * math.pi / 30.0)) * 2.0 / 3.0
    return ret


def transform_lng(x, y):
    ret = 300.0 + x + 2.0 * y + 0.1 * x * x + 0.1 * x * y + 0.1 * math.sqrt(abs(x))
    ret += (20.0 * math.sin(6.0 * x * math.pi) + 20.0 * math.sin(2.0 * x * math.pi)) * 2.0 / 3.0
    ret += (20.0 * math.sin(x * math.pi) + 40.0 * math.sin(x / 3.0 * math.pi)) * 2.0 / 3.0
    ret += (150.0 * math.sin(x / 12.0 * math.pi) + 300.0 * math.sin(x / 30.0 * math.pi)) * 2.0 / 3.0
    return ret


def wgs_to_gcj(lat, lng):
    dlat = transform_lat(lng - 105.0, lat - 35.0)
    dlng = transform_lng(lng - 105.0, lat - 35.0)
    radlat = lat / 180.0 * math.pi
    magic = 1 - 0.00669342162296594323 * (math.sin(radlat) ** 2)
    sqrtmagic = math.sqrt(magic)
    dlat = (dlat * 180.0) / ((6335552.717000426 * magic) / sqrtmagic * math.pi)
    dlng = (dlng * 180.0) / (6378245.0 / sqrtmagic * math.cos(radlat) * math.pi)
    return lat + dlat, lng + dlng


def bd09_to_gcj(bd_lat, bd_lng):
    x = bd_lng - 0.0065
    y = bd_lat - 0.006
    z = math.sqrt(x * x + y * y) - 0.00002 * math.sin(y * X_PI)
    theta = math.atan2(y, x) - 0.000003 * math.cos(x * X_PI)
    return z * math.sin(theta), z * math.cos(theta)


def bd09mc_to_bd09(x, y):
    x, y = float(x), float(y)
    if abs(x) > 1e7:
        x /= 100.0
        y /= 100.0
    factors = MC2LL[-1]
    for i, band in enumerate(MCBAND):
        if abs(y) >= band:
            factors = MC2LL[i]
            break
    cc = abs(y) / factors[9]
    lng = factors[0] + factors[1] * abs(x)
    lat = (
        factors[2] + factors[3] * cc + factors[4] * cc ** 2 + factors[5] * cc ** 3
        + factors[6] * cc ** 4 + factors[7] * cc ** 5 + factors[8] * cc ** 6
    )
    if x < 0:
        lng = -lng
    if y < 0:
        lat = -lat
    return lat, lng


def bd09mc_to_gcj(x, y):
    bd_lat, bd_lng = bd09mc_to_bd09(x, y)
    return bd09_to_gcj(bd_lat, bd_lng)


def in_shenzhen(lat, lng):
    return SZ_LAT[0] <= lat <= SZ_LAT[1] and SZ_LNG[0] <= lng <= SZ_LNG[1]


def haversine_km(lat1, lon1, lat2, lon2):
    r = 6371
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = (
        math.sin(dlat / 2) ** 2
        + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dlon / 2) ** 2
    )
    return 2 * r * math.asin(math.sqrt(a))


def parse_lnglat_pair(text: str):
    part = text.split(';')[0].strip()
    bits = [b.strip() for b in part.split(',')]
    if len(bits) < 2:
        return None
    try:
        a, b = float(bits[0]), float(bits[1])
    except ValueError:
        return None
    if 113.5 <= a <= 114.8 and 22.3 <= b <= 23.0:
        return b, a
    if 113.5 <= b <= 114.8 and 22.3 <= a <= 23.0:
        return a, b
    return None


DISTRICT_PREFIX_RE = re.compile(
    r'^(?:深圳市|广东|福田|罗湖|南山|宝安|龙岗|龙华|盐田|坪山|光明|大鹏|'
    r'海王星辰|国药控股|叮当智慧药房|圆心友和|立丰|万泽|大参林|民心|广药联康|健一生|国大)+'
)


def branch_key(name: str) -> str:
    m = re.search(r'(?:有限公司|连锁有限公司|连锁)(.+)$', name)
    tail = (m.group(1) if m else name).strip()
    tail = re.sub(r'(健康药房|大药房|药房|药店|分店|店)$', '', tail).strip()
    tail = DISTRICT_PREFIX_RE.sub('', tail)
    return tail.strip('（）() ')


def branch_keys(name: str) -> list[str]:
    """Primary branch key plus shorter landmark variants (宝莲大厦 from 福田宝莲大厦)."""
    key = branch_key(name)
    keys = []
    if key:
        keys.append(key)
    # trailing landmark token
    m = re.search(r'([\u4e00-\u9fff]{2,12}(?:大厦|花园|广场|苑|城|中心|公寓))$', key or '')
    if m and m.group(1) not in keys:
        keys.append(m.group(1))
    return keys


def shop_queries(name: str, address: str) -> list[str]:
    qs = []
    name = name.strip()
    address = address.strip()
    keys = branch_keys(name)
    for key in keys:
        if len(key) < 2:
            continue
        if '海王星辰' in name:
            qs += [f'海王星辰({key}分店)', f'海王星辰健康药房({key}店)', f'海王星辰健康药房({key}分店)']
        if '国大药房' in name or '国药控股国大' in name:
            qs += [f'国大药房({key}分店)', f'国大药房({key}店)']
        if '圆心友和' in name:
            qs += [f'圆心友和({key})', f'圆心友和医药({key})']
        if '叮当' in name:
            qs += [f'叮当智慧药房({key})', f'叮当({key})']
        if '大参林' in name:
            qs += [f'大参林({key})', f'大参林大药房({key})']
        qs.append(key)
    branch = BRANCH_RE.search(name)
    if branch:
        qs.append(branch.group(1))
    m = re.search(r'([^\d号路街巷村社区\s]{2,12}(?:花园|大厦|广场|苑|城|中心))\s*\d+栋', address)
    if m:
        qs.insert(0, m.group(0))
    qs.append(name)
    out, seen = [], set()
    for q in qs:
        q = q.strip()
        if q and len(q) >= 2 and q not in seen:
            seen.add(q)
            out.append(q)
    return out[:8]


def extract_gcj(item: dict):
    """GCJ-02 from route_end_list / dest_lnglat, else Baidu mercator x/y."""
    blob = json.dumps(item, ensure_ascii=False)
    for key in ('route_end_list', 'dest_lnglat'):
        m = re.search(rf'"{key}"\s*:\s*"([^"]+)"', blob)
        if m:
            hit = parse_lnglat_pair(m.group(1))
            if hit and in_shenzhen(*hit):
                return hit, 'route_end'

    # navi meters (preferred) or diPoint / x,y (*100)
    for xk, yk in (('navi_x', 'navi_y'), ('diPointX', 'diPointY'), ('x', 'y')):
        if item.get(xk) is not None and item.get(yk) is not None:
            try:
                lat, lng = bd09mc_to_gcj(item[xk], item[yk])
                if in_shenzhen(lat, lng):
                    return (lat, lng), 'bd09mc'
            except Exception:
                pass
    # nested point
    point = ((item.get('ext') or {}).get('detail_info') or {}).get('point') or {}
    if point.get('x') and point.get('y'):
        try:
            lat, lng = bd09mc_to_gcj(point['x'], point['y'])
            if in_shenzhen(lat, lng):
                return (lat, lng), 'bd09mc_point'
        except Exception:
            pass
    return None, None


def address_score(candidate_addr: str, target_addr: str, candidate_name: str, target_name: str, district: str) -> int:
    score = 0
    ca = candidate_addr or ''
    ta = target_addr or ''
    cn = candidate_name or ''
    tn = target_name or ''
    keys = branch_keys(tn)

    if NON_PHARM_RE.search(cn) and not PHARM_NAME_RE.search(cn):
        return -100

    is_pharm = bool(PHARM_NAME_RE.search(cn))
    if not is_pharm:
        # building-only hit: allow only with very strong address identity later
        score -= 8

    dist_token = district.replace('区', '') if district else ''
    if dist_token and dist_token not in ca and dist_token not in cn:
        landmarks = LANDMARK_RE.findall(ta)
        short_lms = re.findall(r'[\u4e00-\u9fff]{2,10}(?:大厦|花园|广场|苑|城|中心)', ta)
        if not any(lm in ca or lm in cn for lm in landmarks + short_lms):
            return -100

    key_hit = False
    for key in keys:
        if len(key) >= 2 and key in cn:
            score += 18
            key_hit = True
            break
        if len(key) >= 2 and key in ca:
            score += 12
            key_hit = True
            break
    if not key_hit:
        if any(b in tn for b in ('海王星辰', '国大药房', '圆心友和', '大参林', '叮当')):
            if is_pharm:
                score -= 12
            else:
                score -= 25

    for token in LANDMARK_RE.findall(ta) + re.findall(r'[\u4e00-\u9fff]{2,10}(?:大厦|花园|广场|苑|城|中心)', ta):
        if token in ca or token in cn:
            score += 10
            break
    for token in re.findall(r'\d+栋', ta):
        if token in ca:
            score += 8
    for token in re.findall(r'[\u4e00-\u9fff]{2,8}路\d*', ta):
        if token in ca:
            score += 5

    if is_pharm:
        score += 6

    compact_t = re.sub(r'\s+', '', ta)
    compact_c = re.sub(r'\s+', '', ca)
    if len(compact_t) >= 16 and compact_t[-18:] in compact_c:
        score += 20
    elif '碧海云天' in ta and '碧海云天' in ca:
        score += 14

    # reject non-pharmacy unless address identity is excellent
    if not is_pharm and score < 28:
        return -50
    return score


def baidu_search(query: str) -> list[dict]:
    url = 'https://api.map.baidu.com/?' + urllib.parse.urlencode({
        'qt': 's', 'c': '340', 'wd': query, 'rn': '8', 'ie': 'utf-8',
        'oue': '1', 'fromproduct': 'jsapi', 'res': 'api',
    })
    data = get_json(url)
    content = data.get('content') or []
    return content if isinstance(content, list) else []


def geocode_pharmacy(name: str, address: str, district: str):
    best = None
    best_score = -1
    for q in shop_queries(name, address):
        time.sleep(0.22)
        try:
            items = baidu_search(q)
        except Exception as e:
            print('  search fail', q, e)
            continue
        for item in items:
            score = address_score(item.get('addr', ''), address, item.get('name', ''), name, district)
            gcj, src = extract_gcj(item)
            if not gcj or score < 0:
                continue
            lat, lng = gcj
            if not in_shenzhen(lat, lng):
                continue
            if score > best_score:
                best_score = score
                best = {
                    'lat': round(lat, 6),
                    'lng': round(lng, 6),
                    'poi_name': item.get('name'),
                    'poi_addr': item.get('addr'),
                    'score': score,
                    'query': q,
                    'pinApprox': score < 20,
                    'geoSource': f'baidu_poi_gcj02:{src}',
                }
        if best_score >= 30:
            break
    if best and best['score'] >= 14:
        return best
    return None


def attach_hospital_ref(item, hospitals):
    nearest = None
    nearest_km = 9999
    for h in hospitals:
        d = haversine_km(item['lat'], item['lng'], h['lat'], h['lng'])
        if d < nearest_km:
            nearest_km = d
            nearest = h
    item.pop('isNearHospital', None)
    item.pop('hospitalRef', None)
    if nearest and nearest_km <= 0.8:
        item['isNearHospital'] = True
        item['hospitalRef'] = nearest['name']


def hospital_queries(h):
    sc = h['name_sc']
    qs = [sc, sc.replace('（', '(').replace('）', ')'), h['address']]
    if '人民医院' in sc and '留医' in sc:
        qs = [
            '深圳市人民医院-东门(2号门)',
            '深圳市人民医院东门北路1017号',
            '深圳市人民医院(留医部)',
            '深圳市人民医院留医部',
            h['address'],
        ]
    elif '第三人民医院' in sc:
        qs = ['深圳市第三人民医院-2号门', '深圳市第三人民医院布澜路29号', '深圳市第三人民医院体检中心', sc]
    elif '香港大学' in sc:
        qs = ['香港大学深圳医院', '香港大学深圳医院门诊部', h['address']]
    elif '肿瘤医院' in sc:
        qs = [sc, '中国医学科学院肿瘤医院深圳医院', h['address']]
    return [q for q in qs if q]


def hospital_score(item, h, baseline) -> int:
    nm = item.get('name') or ''
    addr = item.get('addr') or ''
    sc = h['name_sc']
    score = 0

    if any(bad in nm for bad in ('社康', '国际医疗', '药房', '药店')) and sc not in nm:
        return -40
    # avoid wrong district hospitals with similar names
    if '深圳市人民医院' in sc and '人民医院' in nm and '深圳市人民医院' not in nm:
        return -50
    if '第三人民医院' in sc and '第三人民医院' not in nm and '三院' not in nm:
        return -40

    core = sc.replace('（留医部）', '').replace('（', '').replace('）', '')
    if nm == sc or nm == core or nm.replace('(', '（').replace(')', '）') == sc:
        score += 16
    if core in nm:
        score += 10
    if any(g in nm for g in ('东门', '2号门', '门诊', '门诊部', '门诊楼')):
        score += 12

    addr_key = (h.get('address') or '').replace('深圳市', '').replace('广东省', '')
    for token in re.findall(r'[\u4e00-\u9fff]{2,8}(?:路|街)\d*', addr_key):
        if token in addr or token in nm:
            score += 10
    for token in re.findall(r'\d+号', addr_key):
        if token in addr:
            score += 6

    dist = (h.get('district') or '').replace('區', '区')
    if dist and (dist in addr or dist.replace('区', '') in addr):
        score += 4

    # distance vs baseline GCJ — Baidu sometimes returns wrong campus for main POI
    gcj, _ = extract_gcj(item)
    if gcj and baseline:
        d = haversine_km(gcj[0], gcj[1], baseline[0], baseline[1])
        if d > 2.5:
            # only keep if address road matches strongly
            roads = re.findall(r'[\u4e00-\u9fff]{2,8}路', addr_key)
            if not roads or not any(r in addr for r in roads):
                return -30
            score -= 8
        elif d < 0.8:
            score += 6
    return score


def geocode_hospital(h, baseline_gcj):
    best = None
    best_score = -1
    for q in hospital_queries(h):
        time.sleep(0.3)
        for item in baidu_search(q):
            score = hospital_score(item, h, baseline_gcj)
            gcj, src = extract_gcj(item)
            if not gcj or score < 10:
                continue
            lat, lng = gcj
            if not in_shenzhen(lat, lng):
                continue
            if baseline_gcj and haversine_km(lat, lng, *baseline_gcj) > 4.0 and score < 30:
                continue
            if score > best_score:
                best_score = score
                best = (round(lat, 6), round(lng, 6), item.get('name'), item.get('addr'), score, src)
        if best_score >= 30:
            break
    return best


def main():
    hospitals = json.loads(HOSP_OUT.read_text(encoding='utf-8'))
    print('Updating hospital POIs (GCJ-02, Amap-aligned)...')
    for h in hospitals:
        if h.get('geoSource', '').startswith('baidu_poi_gcj02'):
            baseline = (h['lat'], h['lng'])
        else:
            glat, glng = wgs_to_gcj(h['lat'], h['lng'])
            h['lat'], h['lng'] = round(glat, 6), round(glng, 6)
            baseline = (h['lat'], h['lng'])
            h['geoSource'] = 'wgs84_to_gcj02'

        hit = geocode_hospital(h, baseline)
        if hit and hit[4] >= 18:
            old = (h['lat'], h['lng'])
            # reject wild jumps unless score is strong
            jump = haversine_km(hit[0], hit[1], old[0], old[1])
            if jump > 3.0 and hit[4] < 28:
                print(f"  SKIP jump {jump:.1f}km {h['id']} {h['name']} keep {old} (cand {hit[2]} score={hit[4]})")
            else:
                h['lat'], h['lng'] = hit[0], hit[1]
                h['geoSource'] = f'baidu_poi_gcj02:{hit[5]}'
                print(f"  {h['id']} {h['name'][:16]} {old} -> {hit[0], hit[1]} | {hit[2]} score={hit[4]}")
        else:
            print(f"  KEEP {h['id']} {h['name']} {h['lat']},{h['lng']} (score={hit[4] if hit else 'n/a'})")

    existing = {p['code']: p for p in json.loads(OUT.read_text(encoding='utf-8'))}
    df = pd.read_excel(XLSX, sheet_name=0, header=1)
    df.columns = ['id', 'district', 'code', 'name', 'address']
    df = df.dropna(subset=['name'])
    print(f'\nGeocoding {len(df)} pharmacies to GCJ-02...')

    pharmacies = []
    failed = []
    ok_n = 0
    for _, row in df.iterrows():
        code = str(row['code']).strip()
        name = str(row['name']).strip()
        address = str(row['address']).strip()
        district = str(row['district']).strip()
        pid = str(int(row['id'])) if str(row['id']).replace('.', '', 1).isdigit() else str(row['id'])
        hit = geocode_pharmacy(name, address, district)
        if not hit:
            prev = existing.get(code)
            if prev and prev.get('lat') and prev.get('lng'):
                plat, plng = prev['lat'], prev['lng']
                if not str(prev.get('geoSource', '')).startswith('baidu_poi_gcj02'):
                    plat, plng = wgs_to_gcj(plat, plng)
                item = {
                    'id': pid,
                    'district': district,
                    'code': code,
                    'name': name,
                    'address': address,
                    'lat': round(plat, 6),
                    'lng': round(plng, 6),
                    'pinApprox': True,
                    'geoSource': 'fallback_gcj02',
                }
                pharmacies.append(item)
                print(f'FALLBACK {pid:>3} keep previous→GCJ {name[:28]}')
            else:
                failed.append(code)
                print(f'FAIL {pid} {code} {name[:28]}')
            continue
        item = {
            'id': pid,
            'district': district,
            'code': code,
            'name': name,
            'address': address,
            'lat': hit['lat'],
            'lng': hit['lng'],
            'geoSource': hit['geoSource'],
        }
        if hit.get('pinApprox'):
            item['pinApprox'] = True
        pharmacies.append(item)
        ok_n += 1
        print(
            f"OK {pid:>3} score={hit['score']:<2} {hit['lat']},{hit['lng']} "
            f"<- {hit['poi_name']} ({hit['query'][:28]})"
        )

    for item in pharmacies:
        attach_hospital_ref(item, hospitals)

    OUT.write_text(json.dumps(pharmacies, ensure_ascii=False, indent=2), encoding='utf-8')
    HOSP_OUT.write_text(json.dumps(hospitals, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'\nWrote {len(pharmacies)} pharmacies (POI OK={ok_n}), failed={len(failed)}')
    if failed:
        print('Failed:', failed)

    for p in pharmacies:
        if '红树东方' in p['name']:
            print('\nSPOT 红树东方:', p['lat'], p['lng'], p.get('geoSource'), p.get('hospitalRef'))
            for h in hospitals:
                if '香港大學' in h['name']:
                    d = haversine_km(p['lat'], p['lng'], h['lat'], h['lng'])
                    print(f"  vs {h['name']}: {int(d * 1000)} m @ {h['lat']},{h['lng']}")


if __name__ == '__main__':
    main()
