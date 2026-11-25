"""Combine local daily QC, station metadata and GLO-30 samples in a review workbook."""
import argparse
import calendar
from contextlib import ExitStack
from datetime import date
from hashlib import sha256
import math
import os
from pathlib import Path, PureWindowsPath
import sys
from tempfile import mkstemp

from openpyxl import load_workbook
from xlsxwriter import Workbook


MISSING = {"", "na", "nan", "-9999", "-9999.0"}
FOREST = "WSL_LWF:isone_bestand_12"
OPEN = "WSL_LWF:isone_freiland_13"
ARCHIVE = "https://doi.org/10.1594/PANGAEA.868407"


def daily_path(dataset, relative_path):
    path = PureWindowsPath(str(relative_path))
    if (path.anchor or len(path.parts) != 2 or path.parts[0].casefold() != "daily"
            or path.name in {".", ".."} or ":" in path.name):
        raise ValueError(f"Invalid Daily file path: {relative_path!r}")
    result = dataset.joinpath(*path.parts)
    if result.resolve().parent != dataset / "Daily":
        raise ValueError(f"Daily file resolves outside the Daily folder: {relative_path!r}")
    return result


def daily_grid(blob):
    values = {}
    previous = (0, 0)
    for number, line in enumerate(blob.decode("utf-8-sig").splitlines(), start=1):
        if not line.strip():
            continue
        fields = line.split("\t")
        try:
            if len(fields) != 33:
                raise ValueError(f"expected 33 tab-separated fields, found {len(fields)}")
            year, month = map(int, fields[:2])
            if not (1 <= year <= 9999 and 1 <= month <= 12):
                raise ValueError("invalid calendar month")
            if (year, month) <= previous:
                raise ValueError("duplicate or out-of-order month")
            previous = year, month
            n_days = calendar.monthrange(year, month)[1]
            for day, token in enumerate(fields[2:], start=1):
                token = token.strip()
                value = None if token.lower() in MISSING else float(token)
                if value is not None and (not math.isfinite(value) or value < 0):
                    raise ValueError(f"invalid precipitation {token!r} on day {day}")
                if day > n_days:
                    if value is not None:
                        raise ValueError(f"numeric precipitation on nonexistent day {day}")
                    continue
                values[date(year, month, day).toordinal()] = value
        except ValueError as exc:
            raise ValueError(f"line {number}: {exc}") from exc
    if not values:
        raise ValueError("empty daily grid")
    return values


def write_sheet(workbook, title, headers, rows, widths):
    ws = workbook.add_worksheet(title)
    header = workbook.add_format({"bold": True, "font_color": "white", "bg_color": "#21485E",
                                  "text_wrap": True, "valign": "vcenter"})
    ws.write_row(0, 0, headers, header)
    ws.set_row(0, 30)
    ws.hide_gridlines(2)
    ws.freeze_panes(1, 1)
    for i, width in enumerate(widths):
        ws.set_column(i, i, width)
    i = 0
    for i, row in enumerate(rows, 1):
        if len(row) != len(headers):
            raise ValueError(f"Report column count mismatch in {title}")
        ws.write_row(i, 0, row)
    ws.autofilter(0, 0, i, len(headers) - 1)


def read_rows(book, name):
    data = iter(book[name].values)
    columns = next(data)
    return [dict(zip(columns, row)) for row in data if any(v is not None for v in row)]


def indexed(rows, column):
    result = {}
    for row in rows:
        key = row[column]
        if not key or key in result:
            raise ValueError(f"Missing or duplicate {column}: {key}")
        result[key] = row
    return result


def numeric(value):
    if value is None or value == "":
        return None
    number = float(value)
    return number if math.isfinite(number) else None


def tile_tag(lon, lat):
    y, x = math.floor(lat), math.floor(lon)
    return f"{'N' if y >= 0 else 'S'}{abs(y):02d}_00_{'E' if x >= 0 else 'W'}{abs(x):03d}_00"


def dem_samples(tiles, stations, cutoff):
    import rasterio
    from rasterio.crs import CRS
    from rasterio.windows import Window

    files = sorted(tiles.glob("Copernicus_DSM_COG_10_*_DEM.tif"))
    if not files:
        raise ValueError(f"No GLO-30 GeoTIFFs in {tiles}")
    result = {}
    source_rows = []
    with ExitStack() as stack:
        rasters = {}
        for path in files:
            ds = stack.enter_context(rasterio.open(path))
            if ds.crs != CRS.from_epsg(4326) or ds.count != 1:
                raise ValueError(f"Expected a single-band WGS84 tile: {path.name}")
            rasters[path.stem] = ds
            with path.open("rb") as handle:
                digest = sha256()
                for data in iter(lambda: handle.read(1024 * 1024), b""):
                    digest.update(data)
            source_rows.append([path.name, path.stat().st_size, digest.hexdigest(), "EPSG:4326"])
        for sid, station in stations.items():
            lon, lat, elev = (numeric(station[k]) for k in ("longitude", "latitude", "elevation_m"))
            coords = station["coordinate_system"]
            label = str(coords or "").strip().upper()
            assumed = label != "EPSG:4326"
            out = {"dem": None, "delta": None, "tile": None, "local_min": None, "local_max": None,
                   "status": "Missing or invalid coordinates", "provisional": assumed}
            if lon is None or lat is None or not (-180 <= lon <= 180 and -90 <= lat <= 90):
                result[sid] = out
                continue
            if label not in {"", "EPSG:4326", "NOT EXPLICITLY VERIFIED IN ARCHIVE METADATA"}:
                out["status"] = "Coordinate CRS needs transformation"
                result[sid] = out
                continue
            tag = f"Copernicus_DSM_COG_10_{tile_tag(lon,lat)}_DEM"
            order = [tag, *[k for k in rasters if k != tag]] if tag in rasters else list(rasters)
            out["status"] = "No valid DEM sample"
            for name in order:
                ds = rasters[name]
                row, col = ds.index(lon, lat)
                if not (0 <= row < ds.height and 0 <= col < ds.width):
                    continue
                # Read the containing pixel, plus available neighbours from the same tile.
                col0, row0 = max(col-1, 0), max(row-1, 0)
                width, height = min(col+2, ds.width)-col0, min(row+2, ds.height)-row0
                values = ds.read(1, window=Window(col0, row0, width, height), masked=True)
                centre = values[row-row0, col-col0]
                if bool(getattr(centre, "mask", False)):
                    continue
                dem = numeric(centre)
                if dem is None:
                    continue
                d = elev-dem if elev is not None else None
                context = [float(v) for v in values.compressed() if math.isfinite(float(v))]
                status = ("Registered elevation missing" if d is None else
                          "Review elevation difference" if abs(d) >= cutoff else "Below screening threshold")
                out.update(dem=round(dem,2), delta=round(d,2) if d is not None else None,
                           tile=name+".tif", local_min=round(min(context),2), local_max=round(max(context),2),
                           status=status)
                break
            result[sid] = out
    return result, source_rows


def nearby_open_gauge(dataset, coverage, extremes):
    if OPEN not in coverage or not any(row["Station ID"] == FOREST for row in extremes):
        return {}
    station = coverage[OPEN]
    path = daily_path(dataset, station["Final file"])
    data = path.read_bytes()
    if sha256(data).hexdigest() != station["SHA-256"]:
        raise ValueError(f"Isone open gauge differs from daily QC snapshot: {path}")
    grid = daily_grid(data)
    return {row["Date"]: grid.get(date.fromisoformat(row["Date"]).toordinal())
            for row in extremes if row["Station ID"] == FOREST}


def build_report(project, qc_path, tiles, cutoff):
    dataset = (project / "03_Daily_dataset").resolve()
    register = load_workbook(dataset / "Station_Register.xlsx", read_only=True, data_only=True)
    try:
        stations = indexed(read_rows(register, "Stations"), "dataset_station_id")
        daily_products = indexed((r for r in read_rows(register, "Products") if r["resolution"] == "daily"),
                                 "dataset_station_id")
    finally:
        register.close()
    qc = load_workbook(qc_path, read_only=True, data_only=True)
    try:
        coverage = indexed(read_rows(qc, "Station coverage"), "Station ID")
        extremes = read_rows(qc, "Extreme days")
        neighbours = read_rows(qc, "Event neighbours")
        missing, zeros = (read_rows(qc, k) for k in ("Missing runs", "Zero runs"))
        period = indexed(read_rows(qc, "Summary"), "Item")["Period"]["Result"]
    finally:
        qc.close()
    if not stations or stations.keys() != coverage.keys() or stations.keys() != daily_products.keys():
        raise ValueError("Station IDs differ between the register and daily QC workbook")
    field_pairs = (("longitude", "Longitude"), ("latitude", "Latitude"), ("elevation_m", "Elevation (m)"))
    for sid, st in stations.items():
        old = coverage[sid]
        for a, b in field_pairs:
            left, right = numeric(st[a]), numeric(old[b])
            if left != right and (None in (left, right) or not math.isclose(left, right, abs_tol=1e-8, rel_tol=0)):
                raise ValueError(f"Registration coordinates or elevation changed since daily QC: {sid}")
        if str(daily_products[sid]["sha256"]).lower() != str(old["SHA-256"]).lower():
            raise ValueError(f"The register records a different daily file from QC: {sid}")

    dem, sources = dem_samples(tiles, stations, cutoff)
    open_values = nearby_open_gauge(dataset, coverage, extremes)
    open_window = daily_products.get(OPEN, {}).get("accumulation_interval")
    events, station_rows = [], []
    refs = {}
    for row in neighbours:
        refs.setdefault((row["Target station ID"], row["Target date"]), []).append(row)
    for row in extremes:
        sid, day = row["Station ID"], row["Date"]
        nearby = refs.get((sid, day), [])
        same = [r["D (mm)"] for r in nearby if r["D (mm)"] is not None
                and r["Window comparison"] == "Same documented UTC window"]
        pool = [r["D (mm)"] for r in nearby if r["D (mm)"] is not None]
        status = dem[sid]
        events.append([sid, row["Station"], day, row["Rainfall (mm)"], row["UTC window"],
                       status["dem"], status["delta"], max(pool, default=None), len(same), max(same, default=None),
                       open_values.get(day) if sid == FOREST else None, open_window if sid == FOREST else None,
                       row["Source file"], ARCHIVE if sid == FOREST else None,
                       "Verify hourly source and accumulation window"])
    for sid, st in sorted(stations.items()):
        d = dem[sid]
        obs = coverage[sid]
        station_rows.append([sid, st["station_name"], st["network"], st["longitude"], st["latitude"],
                             st["coordinate_system"], d["provisional"], st["elevation_m"], d["dem"],
                             d["delta"], abs(d["delta"]) if d["delta"] is not None else None,
                             d["local_min"], d["local_max"], d["tile"], d["status"], obs["Observed days"],
                             obs["Extreme days"], obs["Missing runs"], obs["Zero runs"]])
    output = project / "04_Analysis_Review"
    output.mkdir(exist_ok=True)
    fd, name = mkstemp(prefix="06_Station_DEM_QC_", suffix=".part.xlsx", dir=output)
    os.close(fd)
    temporary = Path(name)
    final = temporary.with_name(temporary.name.replace(".part.xlsx", ".xlsx"))
    try:
        with Workbook(temporary, {"constant_memory": True, "strings_to_formulas": False,
                                  "strings_to_urls": False}) as wb:
            metrics = [
                ["Period of copied daily QC", period], ["Station records", len(stations)],
                ["GLO-30 tiles read", len(sources)], ["Valid DEM samples", sum(d["dem"] is not None for d in dem.values())],
                ["Review elevation differences", sum(d["status"] == "Review elevation difference" for d in dem.values())],
                ["Provisional coordinate CRS", sum(d["provisional"] for d in dem.values())],
                ["Extreme rainfall cases pending", len(events)], ["Missing runs", len(missing)], ["Recorded zero runs", len(zeros)],
                ["Station register", str(dataset / "Station_Register.xlsx")], ["Daily QC input", str(qc_path)],
                ["Flag threshold (m)", cutoff], ["Difference", "Registered elevation minus GLO-30 elevation"],
                ["Historical sites", "Uses the register location for each ID. Historical site segments require dated metadata."],
                ["DEM sampling", "Containing pixel. Local min/max use valid pixels in the surrounding 3×3 window from that tile."],
                ["Provisional CRS", "Unverified archive coordinates are interpreted provisionally as WGS84 longitude/latitude."],
                ["Rainfall evidence", "Neighbour comparisons are context. Source hourly values must be checked before corrections."],
            ]
            write_sheet(wb, "Summary", ["Item", "Value"], metrics, [37, 108])
            write_sheet(wb, "Extreme review", ["Station ID","Station","Date","Rainfall (mm)","UTC window","DEM (m)",
                "Elevation difference (m)","Neighbour maximum (mm)","Same-window neighbours","Same-window maximum (mm)",
                "Isone open gauge (mm)","Isone open window","Daily source","Hourly archive","Next check"],
                events,[44,22,14,16,15,15,17,18,17,19,17,32,69,49,45])
            detail = ["Target station ID","Target date","Target rainfall (mm)","Neighbour station ID","Neighbour",
                      "Distance (km)","D−1 (mm)","D (mm)","D+1 (mm)","Window comparison","Neighbour source"]
            write_sheet(wb, "Event neighbours", detail, ([row[h] for h in detail] for row in neighbours),
                        [44,14,17,44,21,15,14,14,14,32,65])
            write_sheet(wb, "Station DEM", ["Station ID","Station","Network","Longitude","Latitude","Coordinate CRS",
                "CRS provisional","Registered elevation (m)","DEM elevation (m)","Register minus DEM (m)","Absolute difference (m)",
                "Nearby DEM minimum (m)","Nearby DEM maximum (m)","Tile","Elevation review","Observed days",
                "Extreme days","Missing runs","Zero runs"], station_rows,
                [45,23,16,14,14,40,15,18,17,17,18,17,17,52,29,15,15,15,15])
            for title, rows in (("Missing runs", missing), ("Recorded zero runs", zeros)):
                headers = ["Station ID","Station","Network","First date","Last date","Consecutive days","Touches start","Touches end"]
                write_sheet(wb,title,headers,([r[h] for h in headers] for r in rows),[45,24,16,15,15,18,15,15])
            write_sheet(wb,"DEM sources",["Tile","Bytes","SHA-256","Horizontal CRS"],sources,[57,16,69,18])
        temporary.rename(final)
    finally:
        temporary.unlink(missing_ok=True)
    print(f"Sampled {sum(d['dem'] is not None for d in dem.values())}/{len(stations)} stations")
    print(f"Workbook: {final}")
    return final


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project", type=Path, required=True)
    parser.add_argument("--qc", type=Path, help="Daily QC workbook; defaults to newest 05_Daily_QC_*.xlsx")
    parser.add_argument("--tiles", type=Path, help="Local GLO-30 tiles directory")
    parser.add_argument("--flag-m", type=float, default=100)
    args = parser.parse_args()
    if not math.isfinite(args.flag_m) or args.flag_m <= 0:
        parser.error("--flag-m must be a positive number")
    project = args.project.resolve()
    review = project / "04_Analysis_Review"
    qc = args.qc or next(iter(sorted(review.glob("05_Daily_QC_*.xlsx"), reverse=True)), None)
    if qc is None:
        raise ValueError("No daily QC report in 04_Analysis_Review. Use --qc to specify its path.")
    build_report(project, qc.resolve(), (args.tiles or review / "06_DEM" / "tiles").resolve(), args.flag_m)


if __name__ == "__main__":
    try:
        main()
    except (OSError, ValueError, KeyError) as exc:
        sys.exit(f"DEM review stopped: {exc}")
