# /// script
# requires-python = ">=3.11"
# ///
"""Fetch Maharashtra district boundaries and build the district run plan.

One Nominatim request per district (cached in maharashtra/_raw/), then writes:

    maharashtra/polygons/<slug>.geojson   boundary geometry for --polygon
    maharashtra/districts.json            name, bbox, area, cell counts

Run once, then use run_maharashtra_districts.ps1.

    python fetch_maharashtra_districts.py
    python fetch_maharashtra_districts.py --force
    python fetch_maharashtra_districts.py --only kolhapur,pune
"""

import argparse
import json
import math
import sys
import time
import urllib.parse
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

from grid import (bbox_size_km, cell_touches_polygons, make_cells,
                  parse_polygon_geojson, polygon_area_km2)

ROOT = Path(__file__).resolve().parent
OUT_DIR = ROOT / 'maharashtra'
POLY_DIR = OUT_DIR / 'polygons'
RAW_DIR = OUT_DIR / '_raw'
PLAN_FILE = OUT_DIR / 'districts.json'

NOMINATIM = 'https://nominatim.openstreetmap.org/search'
USER_AGENT = ('MaharashtraMobileShopMapper/1.0 '
              '(local research; boundaries cached locally)')

# Official 36 districts. Aliases cover the 2023 renames and Mumbai's split.
# Ahmednagar -> Ahilyanagar, Aurangabad -> Chhatrapati Sambhajinagar,
# Osmanabad -> Dharashiv. Old names stay as aliases for Nominatim.
DISTRICTS = [
    ('Ahilyanagar', ['Ahilyanagar district', 'Ahmednagar district']),
    ('Akola', ['Akola district']),
    ('Amravati', ['Amravati district']),
    ('Beed', ['Beed district']),
    ('Bhandara', ['Bhandara district']),
    ('Buldhana', ['Buldhana district']),
    ('Chandrapur', ['Chandrapur district']),
    ('Chhatrapati Sambhajinagar', ['Chhatrapati Sambhajinagar district',
                                   'Aurangabad district']),
    ('Dharashiv', ['Dharashiv district', 'Osmanabad district']),
    ('Dhule', ['Dhule district']),
    ('Gadchiroli', ['Gadchiroli district']),
    ('Gondia', ['Gondiya district', 'Gondia district']),
    ('Hingoli', ['Hingoli district']),
    ('Jalgaon', ['Jalgaon district']),
    ('Jalna', ['Jalna district']),
    ('Kolhapur', ['Kolhapur district']),
    ('Latur', ['Latur district']),
    ('Mumbai City', ['Mumbai City district', 'Mumbai district']),
    ('Mumbai Suburban', ['Mumbai Suburban district']),
    ('Nagpur', ['Nagpur district']),
    ('Nanded', ['Nanded district']),
    ('Nandurbar', ['Nandurbar district']),
    ('Nashik', ['Nashik district']),
    ('Palghar', ['Palghar district']),
    ('Parbhani', ['Parbhani district']),
    ('Pune', ['Pune district']),
    ('Raigad', ['Raigad district']),
    ('Ratnagiri', ['Ratnagiri district']),
    ('Sangli', ['Sangli district']),
    ('Satara', ['Satara district']),
    ('Sindhudurg', ['Sindhudurg district']),
    ('Solapur', ['Solapur district']),
    ('Thane', ['Thane district']),
    ('Wardha', ['Wardha district']),
    ('Washim', ['Washim district']),
    ('Yavatmal', ['Yavatmal district']),
]

# Old district names accepted by --only after the 2023 renames.
ONLY_ALIASES = {
    'ahmednagar': 'ahilyanagar',
    'aurangabad': 'chhatrapati_sambhajinagar',
    'osmanabad': 'dharashiv',
}


def slugify(text: str) -> str:
    return '_'.join(text.lower().replace('-', ' ').split())


def fetch(query: str, retries: int = 3) -> list[dict]:
    params = urllib.parse.urlencode({
        'q': query,
        'format': 'json',
        'limit': 8,
        'polygon_geojson': 1,
        'addressdetails': 1,
        'countrycodes': 'in',
    })
    request = urllib.request.Request(
        f'{NOMINATIM}?{params}', headers={'User-Agent': USER_AGENT})
    last_error = None
    for attempt in range(1, retries + 1):
        try:
            with urllib.request.urlopen(request, timeout=45) as response:
                return json.loads(response.read().decode('utf-8'))
        except Exception as exc:  # network hiccup, rate limit, etc.
            last_error = exc
            if attempt < retries:
                time.sleep(5 * attempt)
    raise RuntimeError(f'Nominatim failed for {query!r}: {last_error}')


def pick_boundary(results: list[dict], name: str):
    """Choose the best administrative polygon among Nominatim results."""
    scored = []
    for item in results:
        geometry = item.get('geojson')
        if not geometry:
            continue
        display = (item.get('display_name') or '').lower()
        if 'maharashtra' not in display:
            continue
        try:
            shapes = parse_polygon_geojson(geometry)
            area = polygon_area_km2(shapes)
        except Exception:
            continue
        is_admin = (item.get('class') == 'boundary'
                    and item.get('type') == 'administrative')
        # Districts are 40-18,000 km2. Reject city nodes and state boundaries.
        if not (20 <= area <= 60_000):
            continue
        scored.append((is_admin, area, item, shapes))
    if not scored:
        return None, None, None
    # Administrative boundary first, then largest area (a district polygon is
    # always bigger than a municipal corporation polygon).
    scored.sort(key=lambda row: (row[0], row[1]), reverse=True)
    _, area, item, shapes = scored[0]
    return item, shapes, area


def build_district(name: str, aliases: list[str], force: bool,
                   step: float = 0.05) -> dict | None:
    slug = slugify(name)
    poly_file = POLY_DIR / f'{slug}.geojson'
    raw_file = RAW_DIR / f'{slug}.json'

    if poly_file.exists() and not force:
        data = json.loads(poly_file.read_text(encoding='utf-8'))
        shapes = parse_polygon_geojson(data)
        area = polygon_area_km2(shapes)
        props = data.get('properties') or {}
        osm = (props.get('osm_type'), props.get('osm_id'))
    else:
        results = None
        chosen = None
        for alias in aliases:
            query = f'{alias}, Maharashtra, India'
            print(f'  query: {query}')
            results = fetch(query)
            item, shapes, area = pick_boundary(results, name)
            if item is not None:
                chosen = (alias, item, shapes, area)
                break
            time.sleep(1.2)
        if chosen is None:
            print(f'  [FAIL] no district boundary found for {name}')
            return None
        alias, item, shapes, area = chosen
        RAW_DIR.mkdir(parents=True, exist_ok=True)
        raw_file.write_text(json.dumps(results, ensure_ascii=False),
                            encoding='utf-8')
        feature = {
            'type': 'Feature',
            'properties': {
                'name': name,
                'display_name': item.get('display_name'),
                'osm_type': item.get('osm_type'),
                'osm_id': item.get('osm_id'),
                'query': alias,
                'area_km2': round(area, 1),
            },
            'geometry': item['geojson'],
        }
        poly_file.parent.mkdir(parents=True, exist_ok=True)
        poly_file.write_text(json.dumps(feature, ensure_ascii=False),
                             encoding='utf-8')
        osm = (item.get('osm_type'), item.get('osm_id'))

    # Recompute bbox from the saved geometry so the grid always matches it.
    lats = [lat for shape in shapes for ring in shape.rings for _, lat in ring]
    lons = [lon for shape in shapes for ring in shape.rings for lon, _ in ring]
    bbox = (min(lats), min(lons), max(lats), max(lons))
    height_km, width_km = bbox_size_km(bbox)

    raw_cells = make_cells(bbox, step)
    cells = [c for c in raw_cells if cell_touches_polygons(c, shapes)]

    print(f'  boundary: {len(shapes)} polygon(s), '
          f'{len(raw_cells)} -> {len(cells)} cells '
          f'({height_km:.0f} x {width_km:.0f} km, {area:.0f} km2)')
    return {
        'name': name,
        'slug': slug,
        'bbox': [round(v, 6) for v in bbox],
        'polygon': f'polygons/{slug}.geojson',
        'area_km2': round(area, 1),
        'cells_total': len(raw_cells),
        'cells_clipped': len(cells),
        'height_km': round(height_km, 1),
        'width_km': round(width_km, 1),
        'osm_type': osm[0],
        'osm_id': osm[1],
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument('--force', action='store_true',
                        help='ignore cached polygons and refetch')
    parser.add_argument('--only', default='',
                        help='comma-separated district names or slugs')
    parser.add_argument('--step', type=float, default=0.05,
                        help='grid cell size in degrees (default 0.05; '
                             'coarser loses shops, finer costs searches)')
    parser.add_argument('--min-step', type=float, default=0.00125,
                        help='cap-split threshold in degrees (default 0.00125; '
                             'actual smallest cell is about half this, ~175 m)')
    parser.add_argument('--max-zoom', type=int, default=15,
                        help='default cap for splits at this Maps zoom '
                             '(default 15, ~2.8 km). The six metros are '
                             'automatically raised to 16 (~1.4 km).')
    args = parser.parse_args()

    only = {ONLY_ALIASES.get(s.strip().lower().replace(' ', '_'),
                             s.strip().lower().replace(' ', '_'))
            for s in args.only.split(',') if s.strip()}

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    POLY_DIR.mkdir(parents=True, exist_ok=True)

    selected = []
    for name, aliases in DISTRICTS:
        slug = slugify(name)
        if only and slug not in only and name.lower() not in only:
            continue
        selected.append((name, aliases))

    if not selected:
        print('No districts matched --only.')
        sys.exit(1)

    print(f'Fetching {len(selected)} district boundaries from Nominatim ...')
    districts = []
    failures = []
    for index, (name, aliases) in enumerate(selected, 1):
        print(f'\n[{index}/{len(selected)}] {name}')
        record = build_district(name, aliases, args.force, args.step)
        if record:
            districts.append(record)
        else:
            failures.append(name)
        time.sleep(1.2)  # Nominatim usage policy: max 1 request/second

    if failures:
        print(f'\n[WARN] failed: {", ".join(failures)}')

    # Keep previously fetched districts when --only was used.
    existing = {}
    if PLAN_FILE.exists():
        try:
            old = json.loads(PLAN_FILE.read_text(encoding='utf-8'))
            existing = {d['slug']: d for d in old.get('districts', [])}
        except Exception:
            existing = {}
    for record in districts:
        existing[record['slug']] = record

    ordered = []
    for name, _ in DISTRICTS:
        slug = slugify(name)
        if slug in existing:
            ordered.append(existing[slug])

    # The six big city markets keep one more split level. Everywhere else
    # z15 already captures ~91% of split-derived leads for a fraction of the
    # searches, so the statewide default stays at z15.
    metro_slugs = {'mumbai_city', 'mumbai_suburban', 'thane',
                   'pune', 'nagpur', 'nashik'}
    for record in ordered:
        record['metro'] = record['slug'] in metro_slugs
        record['max_zoom'] = 16 if record['metro'] else args.max_zoom

    total_cells = sum(d['cells_clipped'] for d in ordered)
    plan = {
        'state': 'Maharashtra',
        'generated': datetime.now(timezone.utc).isoformat(timespec='seconds'),
        'keywords': ['cell phone store', 'mobile phone repair shop'],
        'optional_keyword': 'smartphone accessories',
        'step': args.step,
        'min_step': args.min_step,
        'max_zoom': args.max_zoom,
        'metro_max_zoom': 16,
        'metro_slugs': sorted(metro_slugs),
        'districts': ordered,
        'totals': {
            'districts': len(ordered),
            'cells_total': sum(d['cells_total'] for d in ordered),
            'cells_clipped': total_cells,
            'queries_2_keywords': total_cells * 2,
            'queries_3_keywords': total_cells * 3,
        },
    }
    PLAN_FILE.write_text(json.dumps(plan, indent=2), encoding='utf-8')

    print(f'\nWrote {PLAN_FILE}')
    print(f'  districts: {len(ordered)}')
    print(f'  cells (rectangles): {plan["totals"]["cells_total"]:,}')
    print(f'  cells (clipped):    {total_cells:,}')
    print(f'  queries @2 keywords: {total_cells * 2:,}')
    print(f'  queries @3 keywords: {total_cells * 3:,}')


if __name__ == '__main__':
    main()
