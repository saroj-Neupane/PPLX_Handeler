import math
from pathlib import Path

import shapefile  # pyshp
from pyproj import CRS, Transformer


def load_crs_from_prj(prj_path: Path) -> CRS:
    wkt = prj_path.read_text()
    return CRS.from_wkt(wkt)


def transformer_wgs84_to_layer(prj_path: Path) -> Transformer:
    layer_crs = load_crs_from_prj(prj_path)
    wgs84 = CRS.from_epsg(4326)
    return Transformer.from_crs(wgs84, layer_crs, always_xy=True)


def point_to_segment_dist2(px: float, py: float, x1: float, y1: float, x2: float, y2: float) -> float:
    """Squared distance from point to line segment."""
    dx = x2 - x1
    dy = y2 - y1
    if dx == 0 and dy == 0:
        return (px - x1) ** 2 + (py - y1) ** 2
    t = ((px - x1) * dx + (py - y1) * dy) / (dx * dx + dy * dy)
    if t < 0:
        return (px - x1) ** 2 + (py - y1) ** 2
    if t > 1:
        return (px - x2) ** 2 + (py - y2) ** 2
    proj_x = x1 + t * dx
    proj_y = y1 + t * dy
    return (px - proj_x) ** 2 + (py - proj_y) ** 2


def point_to_polyline_dist2(px: float, py: float, points) -> float:
    """Squared distance from point to a polyline (list of (x, y))."""
    best = float("inf")
    if len(points) == 1:
        x, y = points[0]
        return (px - x) ** 2 + (py - y) ** 2
    for (x1, y1), (x2, y2) in zip(points, points[1:]):
        d2 = point_to_segment_dist2(px, py, x1, y1, x2, y2)
        if d2 < best:
            best = d2
    return best


def main():
    # Input coordinates in (lat, lon)
    lat1, lon1 = 41.2535750, -96.1102139
    lat2, lon2 = 41.2540500, -96.1102222

    base = Path(__file__).parent
    prj_path = base / "data" / "OPPD" / "ElectricLine selection.prj"
    shp_path = base / "data" / "OPPD" / "ElectricLine selection.shp"

    transformer = transformer_wgs84_to_layer(prj_path)

    # Transformer with always_xy=True expects (lon, lat)
    x1, y1 = transformer.transform(lon1, lat1)
    x2, y2 = transformer.transform(lon2, lat2)

    print("Transformed coordinates (layer CRS, feet):")
    print(f"  P1: ({x1:.3f}, {y1:.3f})")
    print(f"  P2: ({x2:.3f}, {y2:.3f})")

    r = shapefile.Reader(str(shp_path))
    fields = [f[0] for f in r.fields[1:]]
    try:
        idx_master = fields.index("d_masterma")
    except ValueError:
        idx_master = None
    try:
        idx_neutral = fields.index("d_neutralm")
    except ValueError:
        idx_neutral = None
    try:
        idx_orient = fields.index("d_orientat")
    except ValueError:
        idx_orient = None
    try:
        idx_runtype = fields.index("d_runtype")
    except ValueError:
        idx_runtype = None

    best_i = None
    best_score = float("inf")
    best_d1 = None
    best_d2 = None

    for i, (shape_rec, rec) in enumerate(zip(r.shapes(), r.records())):
        pts = shape_rec.points
        if not pts:
            continue
        d1_2 = point_to_polyline_dist2(x1, y1, pts)
        d2_2 = point_to_polyline_dist2(x2, y2, pts)

        # Use max distance as score: we want a line close to both points.
        score = max(d1_2, d2_2)

        if score < best_score:
            best_score = score
            best_i = i
            best_d1 = math.sqrt(d1_2)
            best_d2 = math.sqrt(d2_2)

    if best_i is None:
        print("No candidate line found.")
        return

    rec = r.record(best_i)
    master = rec[idx_master] if idx_master is not None else None
    neutral = rec[idx_neutral] if idx_neutral is not None else None
    orient = rec[idx_orient] if idx_orient is not None else None
    runtype = rec[idx_runtype] if idx_runtype is not None else None

    print("\nBest matching line index:", best_i)
    print(f"  Distance P1 -> line: {best_d1:.2f} ft")
    print(f"  Distance P2 -> line: {best_d2:.2f} ft")
    print("  Attributes:")
    print(f"    d_masterma (primary?): {master}")
    print(f"    d_neutralm:            {neutral}")
    print(f"    d_orientat:            {orient}")
    print(f"    d_runtype:             {runtype}")


if __name__ == "__main__":
    main()

