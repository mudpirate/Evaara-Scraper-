# /// script
# requires-python = ">=3.11"
# ///
"""GPS grid helpers for scrape.py (standard library only).

Google Maps caps a single search at ~120 results, so a broad
"cell phone store in <district>" query misses small shops. Instead, cover
the district with a checkerboard of small cells and search once per cell.

This module builds those cells and can look up a place's bounding box from
a name (OpenStreetMap Nominatim, no API key).
"""

import json
import math
import urllib.parse
import urllib.request
from dataclasses import dataclass
from pathlib import Path

VIEWPORT_WIDTH_PX = 1400
VIEWPORT_HEIGHT_PX = 900
METERS_PER_PIXEL_Z0 = 156543.03392
METERS_PER_DEGREE = 111_320.0

NOMINATIM_URL = "https://nominatim.openstreetmap.org/search"


@dataclass(frozen=True)
class GridCell:
    """One search cell: a lat/lon rectangle plus the zoom to search it at."""

    south: float
    west: float
    north: float
    east: float
    zoom: int

    @property
    def center(self) -> tuple[float, float]:
        return ((self.south + self.north) / 2,
                (self.west + self.east) / 2)

    @property
    def lat_span(self) -> float:
        return self.north - self.south

    @property
    def lon_span(self) -> float:
        return self.east - self.west

    def can_split(self, min_step: float, max_zoom: int = 15) -> bool:
        """True when a 120-result cap can be relieved by splitting again."""
        return (self.zoom < max_zoom
                and min(self.lat_span, self.lon_span) >= 2 * min_step - 1e-12)

    def split(self) -> list["GridCell"]:
        """Four quadrants, one zoom level closer."""
        mid_lat = (self.south + self.north) / 2
        mid_lon = (self.west + self.east) / 2
        z = min(self.zoom + 1, 20)
        return [
            GridCell(self.south, self.west, mid_lat, mid_lon, z),
            GridCell(self.south, mid_lon, mid_lat, self.east, z),
            GridCell(mid_lat, self.west, self.north, mid_lon, z),
            GridCell(mid_lat, mid_lon, self.north, self.east, z),
        ]


def parse_bbox(text: str) -> tuple[float, float, float, float]:
    """Parse 'lat1,lon1,lat2,lon2' (any two opposite corners).

    Returns (lat_min, lon_min, lat_max, lon_max). Normalizing means callers
    can paste top-left and bottom-right in either order.
    """
    parts = [p.strip() for p in text.replace(';', ',').split(',') if p.strip()]
    if len(parts) != 4:
        raise ValueError(
            "bounding box needs 4 numbers: lat1,lon1,lat2,lon2 "
            "(two opposite corners from Google Maps)"
        )
    try:
        lat1, lon1, lat2, lon2 = (float(p) for p in parts)
    except ValueError:
        raise ValueError(f"bounding box is not numeric: {text!r}") from None
    if not (-85 <= lat1 <= 85 and -85 <= lat2 <= 85):
        raise ValueError("latitude must be between -85 and 85")
    if not (-180 <= lon1 <= 180 and -180 <= lon2 <= 180):
        raise ValueError("longitude must be between -180 and 180")
    lat_min, lat_max = sorted((lat1, lat2))
    lon_min, lon_max = sorted((lon1, lon2))
    if lat_max - lat_min <= 0 or lon_max - lon_min <= 0:
        raise ValueError("bounding box has zero width or height")
    return lat_min, lon_min, lat_max, lon_max


def grid_shape(bbox: tuple[float, float, float, float],
               step: float) -> tuple[int, int]:
    """Number of rows (lat) and columns (lon) for a step in degrees."""
    if step <= 0:
        raise ValueError("step must be greater than 0")
    lat_min, lon_min, lat_max, lon_max = bbox
    rows = max(1, math.ceil((lat_max - lat_min) / step - 1e-9))
    cols = max(1, math.ceil((lon_max - lon_min) / step - 1e-9))
    return rows, cols


def auto_zoom(lat_span: float, lon_span: float, lat: float) -> int:
    """Zoom whose viewport still covers one cell (so no gaps between cells).

    Google Maps search results follow the visible map area. A viewport at
    zoom z is 1400x900 px; this picks the largest zoom that keeps the cell
    inside that viewport. Smaller cells therefore get a closer zoom.
    """
    lat_span = max(lat_span, 1e-6)
    lon_span = max(lon_span, 1e-6)
    lat_m = lat_span * METERS_PER_DEGREE
    lon_m = lon_span * METERS_PER_DEGREE * math.cos(math.radians(lat))
    need_m = max(lat_m, lon_m * VIEWPORT_HEIGHT_PX / VIEWPORT_WIDTH_PX)
    world_m = (VIEWPORT_HEIGHT_PX * METERS_PER_PIXEL_Z0
               * math.cos(math.radians(lat)))
    z = math.floor(math.log2(max(world_m / need_m, 1.0)))
    return max(3, min(20, z))


def make_cells(bbox: tuple[float, float, float, float], step: float,
               zoom: int | None = None) -> list[GridCell]:
    """Split a bounding box into cells no larger than `step` degrees."""
    rows, cols = grid_shape(bbox, step)
    lat_min, lon_min, lat_max, lon_max = bbox
    lat_size = (lat_max - lat_min) / rows
    lon_size = (lon_max - lon_min) / cols
    cells: list[GridCell] = []
    for i in range(rows):
        south = lat_min + i * lat_size
        north = south + lat_size
        for j in range(cols):
            west = lon_min + j * lon_size
            east = west + lon_size
            z = zoom if zoom is not None else auto_zoom(lat_size, lon_size,
                                                        (south + north) / 2)
            cells.append(GridCell(south, west, north, east, z))
    return cells


def bbox_size_km(bbox: tuple[float, float, float, float]) -> tuple[float, float]:
    """Return (height_km, width_km) of a bounding box."""
    lat_min, lon_min, lat_max, lon_max = bbox
    mid_lat = (lat_min + lat_max) / 2
    height = (lat_max - lat_min) * METERS_PER_DEGREE / 1000
    width = ((lon_max - lon_min) * METERS_PER_DEGREE
             * math.cos(math.radians(mid_lat)) / 1000)
    return height, width


def lookup_area_bbox(area: str,
                     timeout: int = 25) -> tuple[tuple[float, float, float, float], str]:
    """Look up a place name with OpenStreetMap Nominatim.

    Returns ((lat_min, lon_min, lat_max, lon_max), display_name).
    One request per run; Nominatim asks for a descriptive User-Agent.
    """
    params = urllib.parse.urlencode({
        'q': area,
        'format': 'json',
        'limit': 1,
        'addressdetails': 0,
    })
    request = urllib.request.Request(
        f"{NOMINATIM_URL}?{params}",
        headers={'User-Agent': 'google-maps-scraper/1.0 (local use)'},
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        data = json.loads(response.read().decode('utf-8'))
    if not data:
        raise LookupError(f"no match for '{area}'")
    item = data[0]
    south, north, west, east = (float(v) for v in item['boundingbox'])
    bbox = (min(south, north), min(west, east),
            max(south, north), max(west, east))
    return bbox, item.get('display_name') or area


# ---------------------------------------------------------------------------
# Polygon clipping
#
# A taluka/district bounding box is a rectangle, but the real boundary is an
# irregular polygon. Covering the whole rectangle wastes roughly half the
# search cells in the empty corners (measured ~2x on Solapur talukas). These
# helpers let grid mode skip cells that cannot contain any part of the area.
#
# Coordinates in GeoJSON are (longitude, latitude); all public helpers below
# take (lat, lon) like the rest of this module.
# ---------------------------------------------------------------------------

Point = tuple[float, float]


@dataclass(frozen=True)
class PolygonShape:
    """One GeoJSON polygon: an outer ring followed by optional holes."""

    rings: tuple[tuple[Point, ...], ...]
    bbox: tuple[float, float, float, float]  # lat_min, lon_min, lat_max, lon_max

    def contains(self, lat: float, lon: float) -> bool:
        """True when the point is inside the outer ring and outside all holes."""
        lat_min, lon_min, lat_max, lon_max = self.bbox
        if not (lat_min <= lat <= lat_max and lon_min <= lon <= lon_max):
            return False
        if not _ring_contains(self.rings[0], lon, lat):
            return False
        return not any(_ring_contains(ring, lon, lat) for ring in self.rings[1:])


def _ring_contains(ring: tuple[Point, ...], lon: float, lat: float) -> bool:
    """Ray-casting point-in-ring test. Ring points are (lon, lat)."""
    inside = False
    j = len(ring) - 1
    for i in range(len(ring)):
        xi, yi = ring[i]
        xj, yj = ring[j]
        if (yi > lat) != (yj > lat):
            x_cross = (xj - xi) * (lat - yi) / (yj - yi) + xi
            if lon < x_cross:
                inside = not inside
        j = i
    return inside


def _iter_geojson_polygons(geometry: dict):
    """Yield raw polygon coordinate arrays from any GeoJSON geometry."""
    kind = geometry.get('type')
    if kind == 'Polygon':
        yield geometry.get('coordinates') or []
    elif kind == 'MultiPolygon':
        for polygon in geometry.get('coordinates') or []:
            yield polygon
    elif kind == 'GeometryCollection':
        for child in geometry.get('geometries') or []:
            yield from _iter_geojson_polygons(child)


def load_polygon_geojson(path: str | Path) -> list[PolygonShape]:
    """Load Polygon/MultiPolygon geometry from a GeoJSON file."""
    data = json.loads(Path(path).read_text(encoding='utf-8-sig'))
    try:
        return parse_polygon_geojson(data)
    except ValueError as exc:
        raise ValueError(f'{path}: {exc}') from None


def parse_polygon_geojson(data: dict) -> list[PolygonShape]:
    """Parse a GeoJSON geometry/Feature/FeatureCollection into shapes.

    Returns one PolygonShape per polygon so callers can test several
    disjoint pieces.
    """
    if not isinstance(data, dict):
        raise ValueError('not a GeoJSON object')

    if data.get('type') == 'FeatureCollection':
        geometries = [f.get('geometry') for f in data.get('features') or []]
    elif data.get('type') == 'Feature':
        geometries = [data.get('geometry')]
    else:
        geometries = [data]

    shapes: list[PolygonShape] = []
    for geometry in geometries:
        if not geometry:
            continue
        for polygon in _iter_geojson_polygons(geometry):
            rings: list[tuple[Point, ...]] = []
            for ring in polygon:
                points = tuple((float(p[0]), float(p[1])) for p in ring)
                if len(points) >= 3:
                    rings.append(points)
            if not rings:
                continue
            xs = [p[0] for ring in rings for p in ring]
            ys = [p[1] for ring in rings for p in ring]
            shapes.append(PolygonShape(
                rings=tuple(rings),
                bbox=(min(ys), min(xs), max(ys), max(xs)),
            ))
    if not shapes:
        raise ValueError('no Polygon/MultiPolygon geometry found')
    return shapes


def polygon_contains_any(lat: float, lon: float,
                         polygons: list[PolygonShape]) -> bool:
    """True when (lat, lon) lies inside any of the supplied polygons."""
    return any(poly.contains(lat, lon) for poly in polygons)


def cell_touches_polygons(cell: 'GridCell',
                          polygons: list[PolygonShape]) -> bool:
    """True when a cell is likely to overlap the polygon.

    Tests the centre, four corners and four edge midpoints. A cell that
    straddles the boundary is kept, so shops just inside the edge are not
    lost. A boundary that crosses a cell as a very thin sliver with all nine
    points outside could still be skipped; at 0.05 deg cells this is a
    sub-kilometre edge case and the neighbouring district covers the other
    side.
    """
    lat0, lon0 = cell.center
    half_lat = cell.lat_span / 2
    half_lon = cell.lon_span / 2
    probes = [
        (lat0, lon0),
        (cell.south, cell.west), (cell.south, cell.east),
        (cell.north, cell.west), (cell.north, cell.east),
        (cell.south, lon0), (cell.north, lon0),
        (lat0, cell.west), (lat0, cell.east),
    ]
    # local vars keep the inner loop cheap on big district polygons
    for lat, lon in probes:
        for poly in polygons:
            if poly.contains(lat, lon):
                return True
    return False


def clipped_cells(bbox: tuple[float, float, float, float], step: float,
                  polygons: list[PolygonShape],
                  zoom: int | None = None) -> list['GridCell']:
    """Build the grid and keep only cells that touch the polygon."""
    return [cell for cell in make_cells(bbox, step, zoom=zoom)
            if cell_touches_polygons(cell, polygons)]


def polygon_area_km2(polygons: list[PolygonShape]) -> float:
    """Approximate polygon area in square kilometres (outer minus holes)."""
    total = 0.0
    for poly in polygons:
        lat0 = math.radians((poly.bbox[0] + poly.bbox[2]) / 2)
        for index, ring in enumerate(poly.rings):
            if len(ring) < 3:
                continue
            area = 0.0
            j = len(ring) - 1
            for i in range(len(ring)):
                xi, yi = ring[i]
                xj, yj = ring[j]
                area += (xj + xi) * (yj - yi)
                j = i
            area = abs(area) / 2
            area *= (METERS_PER_DEGREE / 1000) ** 2 * math.cos(lat0)
            total += area if index == 0 else -area
    return total
