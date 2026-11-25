"""Build a local coordinate, DEM, and daily precipitation QC workbook."""
import argparse
from collections import Counter, defaultdict
from contextlib import ExitStack
import csv
from datetime import date, datetime
import gzip
from hashlib import sha256
from importlib.util import find_spec
from io import TextIOWrapper
import math
import os
from pathlib import Path, PureWindowsPath
import sys
from tempfile import mkstemp
from zipfile import BadZipFile, ZipFile

from openpyxl import load_workbook
from xlsxwriter import Workbook


RELEASE_NAME = "IHGAR_MeteoSwiss_Station_Update"
RELEASE_SHA = "c4eb5b2ae66df1a6cfb6e1a50077cb5bc4de99259c5984f130d60f9b5d69a838"
UPDATE_SHEET = "IH-GAR update coordinates"


def daily_path(dataset, relative_path):
    path = PureWindowsPath(str(relative_path))
    if (path.anchor or len(path.parts) != 2 or path.parts[0].casefold() != "daily"
            or path.name in {".", ".."} or ":" in path.name):
        raise ValueError(f"Invalid Daily file path: {relative_path!r}")
    result = dataset.joinpath(*path.parts)
    if result.resolve().parent != dataset / "Daily":
        raise ValueError(f"Daily file resolves outside the Daily folder: {relative_path!r}")
    return result


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
            source_rows.append([path.name, path.stat().st_size, digest(path), "EPSG:4326"])
        for sid, station in stations.items():
            lon, lat, elev = (numeric(station[k]) for k in ("longitude", "latitude", "elevation_m"))
            coords = station["coordinate_system"]
            label = str(coords or "").strip().upper()
            assumed = label != "EPSG:4326"
            out = {"dem": None, "delta": None, "tile": None, "local_min": None, "local_max": None,
                   "pixel_center": (None, None),
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
                           status=status, pixel_center=ds.xy(row, col))
                break
            result[sid] = out
    return result, source_rows


def digest(path):
    checksum = sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            checksum.update(block)
    return checksum.hexdigest()


def book_tables(path, names):
    with ExitStack() as stack:
        book = load_workbook(path, read_only=True, data_only=True)
        stack.callback(book.close)
        return {name: read_rows(book, name) for name in names}


def same_number(a, b):
    a, b = numeric(a), numeric(b)
    return a == b or (a is not None and b is not None and
                      math.isclose(a, b, abs_tol=1e-8, rel_tol=0))


def verify_snapshot(dataset, stations, products, coverage, station_dem):
    if not stations or not (stations.keys() == products.keys() == coverage.keys() == station_dem.keys()):
        raise ValueError("Station IDs differ between the register, daily QC and previous DEM report.")
    for sid, station in stations.items():
        qc, prior = coverage[sid], station_dem[sid]
        for key, qc_key, dem_key in (("longitude", "Longitude", "Longitude"),
                                     ("latitude", "Latitude", "Latitude"),
                                     ("elevation_m", "Elevation (m)", "Registered elevation (m)")):
            if not (same_number(station[key], qc[qc_key]) and same_number(station[key], prior[dem_key])):
                raise ValueError(f"Metadata changed for {sid}. Refresh the daily QC and DEM inputs.")
        if (station["network"] != qc["Network"] or station["network"] != prior["Network"] or
                station["coordinate_system"] != qc["Coordinate CRS"] or
                station["coordinate_system"] != prior["Coordinate CRS"]):
            raise ValueError(f"Network or coordinate CRS differs for {sid}.")
        product = products[sid]
        expected = str(qc["SHA-256"]).lower()
        if str(product["sha256"]).lower() != expected or digest(daily_path(dataset, product["relative_path"])) != expected:
            raise ValueError(f"Daily file version differs from QC for {sid}. Refresh the daily QC report.")


def coordinate_pair(row):
    lon, lat = numeric(row.get("longitude")), numeric(row.get("latitude"))
    if lon is None or lat is None or not (-180 <= lon <= 180 and -90 <= lat <= 90):
        return None
    return lon, lat


def distance(a, b):
    lon_a, lat_a = map(math.radians, a)
    lon_b, lat_b = map(math.radians, b)
    h = (math.sin((lat_b-lat_a)/2)**2 +
         math.cos(lat_a)*math.cos(lat_b)*math.sin((lon_b-lon_a)/2)**2)
    return 6371008.8 * 2 * math.asin(math.sqrt(min(1, h)))


def code(value):
    return str(value or "").strip().upper()


def position_crs(value):
    label = str(value or "").strip().upper()
    if label in {"WGS84 (EXPLICIT COLUMN LABEL)", "WGS84 (EXPLICIT SOURCE COLUMN LABEL)",
                 "WGS84", "WGS 84", "EPSG:4326"}:
        return "EPSG:4326"
    return value


def comparable(row):
    return position_crs(row.get("coordinate_system")) in {
        "EPSG:4326", "Not explicitly verified in archive metadata"
    } and coordinate_pair(row) is not None


def find_update_release(project, explicit):
    if explicit:
        candidates = [explicit]
    else:
        pattern = RELEASE_NAME + "*.zip"
        candidates = [
            *(project / "03_Daily_dataset" / "Release_candidates").rglob(pattern),
            *(project / "02_Data_extraction").rglob(pattern),
            *(Path.home() / "Downloads").glob(pattern),
        ]
    for path in candidates:
        if path.is_file() and digest(path) == RELEASE_SHA:
            return path.resolve()
    raise ValueError("Verified IH-GAR update archive not found. Use --release with the original "
                     f"{RELEASE_NAME}.zip used to build the final dataset.")


def applied_update_locations(release, stations):
    ihgar = indexed((s for s in stations.values() if s["network"] == "IHGAR"), "provider_id")
    with ZipFile(release) as archive:
        with archive.open("replay/stations.csv") as member:
            replay = indexed(csv.DictReader(TextIOWrapper(member, encoding="utf-8-sig")), "target_id")
        file_codes = defaultdict(set)
        with archive.open("replay/source_files.csv") as member:
            for row in csv.DictReader(TextIOWrapper(member, encoding="utf-8-sig")):
                path = f"{row['output_folder']}/{row['output_filename']}"
                file_codes[(path, row["source_parameter"])].add(code(row["source_station_code"]))
        actions = defaultdict(Counter)
        products = defaultdict(set)
        seen = set()
        with archive.open("Applied_Daily_Changes.csv.gz") as member:
            with gzip.open(member, "rt", encoding="utf-8-sig", newline="") as stream:
                for row in csv.DictReader(stream):
                    target, kind = row["IH-GAR ID"], row["Action"]
                    key = (target, row["Date"])
                    source = file_codes.get((row["Source path"], row["Product"]))
                    if (target not in ihgar or target not in replay or key in seen or
                            kind not in {"Addition", "Correction"} or
                            source != {code(replay[target]["code"])}):
                        raise ValueError(f"Unresolved or duplicate update ledger record: {key}")
                    seen.add(key)
                    actions[target][kind] += 1
                    products[target].add(row["Product"])
    expected = {target for target, s in ihgar.items()
                if numeric(s["added_days"]) or numeric(s["corrected_days"])}
    if not expected or expected != actions.keys():
        raise ValueError("The applied update ledger and station register list different updated IH-GAR records.")

    results = {}
    for target, counts in actions.items():
        reg, source = ihgar[target], replay[target]
        meteo = {
            "provider_id": source["code"], "station_name": source["provider_name"],
            "longitude": numeric(source["source_longitude"]),
            "latitude": numeric(source["source_latitude"]),
            "elevation_m": numeric(source["source_elevation_m"]),
            "coordinate_system": "EPSG:4326",
        }
        originals = (("longitude", "Longitudine"), ("latitude", "Latitudine"), ("elevation_m", "Elevazione"))
        if (code(reg["matched_meteoswiss_code"]) != code(source["code"]) or
                any(not same_number(reg[field], source[original]) for field, original in originals) or
                not comparable(reg) or not comparable(meteo) or
                reg["elevation_m"] is None or meteo["elevation_m"] is None or
                counts["Addition"] != (numeric(reg["added_days"]) or 0) or
                counts["Correction"] != (numeric(reg["corrected_days"]) or 0)):
            raise ValueError(f"Station register does not agree with the applied update replay for {target}.")
        results[reg["dataset_station_id"]] = (meteo, sorted(products[target]), counts)
    return results


def write_update_locations(wb, updates, stations, samples, cutoff):
    def signed_gap(left, right):
        return round(left - right, 2) if None not in (left, right) else None

    rows = []
    for sid in sorted(updates, key=lambda key: (stations[key]["station_name"].casefold(), key)):
        reg = stations[sid]
        meteo, products, counts = updates[sid]
        local_dem, donor_dem = samples[sid], samples[("update", sid)]
        original_m, meteo_m = numeric(reg["elevation_m"]), meteo["elevation_m"]
        dems = local_dem["dem"], donor_dem["dem"]
        gaps = (signed_gap(meteo_m, original_m), signed_gap(original_m, dems[0]),
                signed_gap(meteo_m, dems[1]))
        screen = ("Missing DEM sample" if None in dems else
                  "Review" if any(abs(gap) >= cutoff for gap in gaps) else None)
        rows.append([
            sid, reg["station_name"], meteo["provider_id"], meteo["station_name"],
            *coordinate_pair(reg), original_m, dems[0],
            *coordinate_pair(meteo), meteo_m, dems[1],
            round(distance(coordinate_pair(reg), coordinate_pair(meteo)), 1), *gaps,
            signed_gap(dems[1], dems[0]), screen, ", ".join(products),
            counts["Addition"], counts["Correction"], reg["coordinate_system"],
            *(round(value, 8) if value is not None else None for value in local_dem["pixel_center"]),
            *(round(value, 8) if value is not None else None for value in donor_dem["pixel_center"]),
            local_dem["tile"], donor_dem["tile"], None,
        ])
    table(wb, UPDATE_SHEET, [
        "IH-GAR station ID", "IH-GAR station", "MeteoSwiss code", "MeteoSwiss station",
        "IH-GAR longitude (°)", "IH-GAR latitude (°)", "IH-GAR elevation (m)", "DEM at IH-GAR (m)",
        "MeteoSwiss longitude (°)", "MeteoSwiss latitude (°)", "MeteoSwiss elevation (m)", "DEM at MeteoSwiss (m)",
        "Horizontal distance (m)", "MeteoSwiss minus IH-GAR (m)", "IH-GAR minus its DEM (m)",
        "MeteoSwiss minus its DEM (m)", "DEM at MeteoSwiss minus DEM at IH-GAR (m)", "Elevation review",
        "Applied products", "Added days in release", "Corrected days in release", "IH-GAR coordinate CRS",
        "IH-GAR DEM pixel centre lon (°)", "IH-GAR DEM pixel centre lat (°)",
        "MeteoSwiss DEM pixel centre lon (°)", "MeteoSwiss DEM pixel centre lat (°)",
        "DEM tile at IH-GAR", "DEM tile at MeteoSwiss", "Reviewer notes",
    ], rows, [57, 25, 16, 25, 19, 19, 19, 19, 20, 20, 22, 20, 19, 24, 23, 24, 30,
              20, 24, 20, 22, 41, 25, 25, 26, 26, 56, 56, 30])


def candidate_rows(stations, meteo_path, supplied_path, radius_km):
    rows = []
    meteo = defaultdict(list)
    if meteo_path:
        for row in book_tables(meteo_path, ("Stations",))["Stations"]:
            if code(row["network"]) == "METEOSWISS" and coordinate_pair(row) is not None:
                meteo[code(row["provider_station_id"])].append(row)

    for sid, st in sorted(stations.items()):
        link = code(st["matched_meteoswiss_code"])
        is_ihgar = st["network"] == "IHGAR"
        provider = {**st, "longitude": st["provider_longitude"], "latitude": st["provider_latitude"],
                    "elevation_m": st["provider_elevation_m"]}
        provider_changed = any(not same_number(st[key], provider[key])
                               for key in ("longitude", "latitude", "elevation_m"))
        if is_ihgar and (link or coordinate_pair(provider) is not None):
            rows.append((sid, {**provider, "network": "MeteoSwiss",
                               "provider_id": link or None, "station_name": st["provider_station_name"],
                               "coordinate_system": "Not explicitly verified in archive metadata",
                               "location_source": "Station_Register.xlsx / Stations / provider coordinate fields",
                               "location_years": None},
                         f"Register link: {st['identity_status'] or 'unclassified'}"))
        if not is_ihgar and provider_changed and coordinate_pair(provider) is not None:
            rows.append((sid, {**provider, "location_source": "Station_Register.xlsx / Stations / provider coordinate fields"},
                         "Provider coordinates in register"))
        lookup = link if is_ihgar else code(st["provider_id"]) if st["network"] == "MeteoSwiss" else ""
        for item in meteo.get(lookup, ()):
            rows.append((sid, {**item, "network": "MeteoSwiss", "provider_id": item["provider_station_id"],
                               "station_name": item["original_station_name"],
                               "coordinate_system": position_crs(item["coordinate_system"]),
                               "location_source": " | ".join(str(s) for s in (
                                   f"{meteo_path.name} / Stations", item.get("metadata_file"),
                                   item.get("source_record"), item.get("metadata_sha256")
                               ) if s)},
                         "Exact provider code. Check site and dates"))

    if supplied_path:
        required = {"station_id", "network", "provider_id", "longitude", "latitude",
                    "elevation_m", "coordinate_system", "valid_from", "valid_to", "source"}
        with supplied_path.open(encoding="utf-8-sig", newline="") as stream:
            reader = csv.DictReader(stream)
            if not required.issubset(reader.fieldnames or []):
                raise ValueError(f"Candidate CSV needs columns: {', '.join(sorted(required))}")
            for item in reader:
                sid = item["station_id"]
                if sid not in stations or not all(item[key] for key in ("source", "provider_id", "network")):
                    raise ValueError(f"Unresolved station ID or source fields on candidate CSV line {reader.line_num}")
                if not coordinate_pair(item) or not item["coordinate_system"]:
                    raise ValueError(f"Invalid coordinates or missing CRS on candidate CSV line {reader.line_num}")
                start, end = [date.fromisoformat(item[k]) if item[k] else None
                              for k in ("valid_from", "valid_to")]
                if start and end and start > end:
                    raise ValueError(f"Reversed dates on candidate CSV line {reader.line_num}")
                rows.append((sid, {**item, "coordinate_system": position_crs(item["coordinate_system"]),
                                   "location_source": item["source"]}, "Documented candidate. Review identity"))

    points = {sid: coordinate_pair(st) for sid, st in stations.items() if comparable(st)}
    for sid, here in sorted(points.items()):
        st = stations[sid]
        nearest_by_network = {}
        for other, there in points.items():
            ref = stations[other]
            network = ref["network"]
            if network == st["network"]:
                continue
            km = distance(here, there) / 1000
            candidate = (km, other)
            if km <= radius_km and candidate < nearest_by_network.get(network, (math.inf, "")):
                nearest_by_network[network] = candidate
        for _, other in sorted(nearest_by_network.values()):
            rows.append((sid, stations[other], "Nearby gauge. Identity unverified"))
    return rows


def table(wb, title, headers, rows, widths=None):
    def excel_values():
        for row in rows:
            yield [v.date().isoformat() if isinstance(v, datetime) and v.time() == datetime.min.time()
                   else v.isoformat() if isinstance(v, (date, datetime)) else v for v in row]
    if widths is None:
        widths = [min(48, max(16, len(h) + 4, 44 if "station id" in h.lower() else 0)) for h in headers]
    write_sheet(wb, title, headers, excel_values(), widths)


def copy_table(wb, title, rows, headers):
    headers = list(rows[0]) if rows else headers
    table(wb, title, headers, ([r[h] for h in headers] for r in rows))


def candidate_details(rows, stations, samples):
    result = []
    for index, (sid, ref, evidence) in enumerate(rows):
        st, base, alt = stations[sid], samples[sid], samples[("candidate", index)]
        pair = coordinate_pair(ref)
        lr = coordinate_pair(st)
        sep = (round(distance(lr, pair), 1) if comparable(ref) and comparable(st) else None)
        result.append([
            sid, st["station_name"], st["network"], evidence, ref["network"],
            ref.get("dataset_station_id") if evidence == "Nearby gauge. Identity unverified" else None,
            ref.get("provider_id"), ref.get("station_name"), st["longitude"], st["latitude"],
            numeric(ref.get("longitude")), numeric(ref.get("latitude")), ref.get("coordinate_system"), sep,
            numeric(st["elevation_m"]), numeric(ref.get("elevation_m")), base["dem"], alt["dem"],
            base["delta"], alt["delta"], alt["status"], ref.get("valid_from"), ref.get("valid_to"),
            ref.get("location_years"), ref.get("location_source") or ref.get("metadata_source"), "Pending", None,
        ])
    return result


def sheet_rows(qc, previous, daily_summary):
    def index_events(tab, id_col, date_col):
        return indexed(({"event": (row[id_col], row[date_col]), **row} for row in tab), "event")

    events = index_events(qc["Extreme days"], "Station ID", "Date")
    prior_events = index_events(previous["Extreme review"], "Station ID", "Date")
    if events.keys() != prior_events.keys() or any(
        not same_number(row["Rainfall (mm)"], prior_events[key]["Rainfall (mm)"])
        for key, row in events.items()
    ):
        raise ValueError("The prior DEM and current daily QC reports list different extreme events.")
    for qc_name, previous_name in (("Missing runs", "Missing runs"), ("Zero runs", "Recorded zero runs")):
        columns = ("Station ID", "First date", "Last date", "Consecutive days", "Touches start", "Touches end")
        current, prior = ([tuple(row[c] for c in columns) for row in book[name]]
                          for book, name in ((qc, qc_name), (previous, previous_name)))
        if Counter(current) != Counter(prior):
            raise ValueError(f"The prior DEM and current daily QC reports disagree on {qc_name}.")
    if daily_summary["Period"]["Result"] != indexed(previous["Summary"], "Item")["Period of copied daily QC"]["Value"]:
        raise ValueError("The daily QC and prior DEM reports describe different date ranges.")
    return events, prior_events


def build(args):
    project = args.project.resolve()
    review = project / "04_Analysis_Review"
    dataset = (project / "03_Daily_dataset").resolve()
    register_path = dataset / "Station_Register.xlsx"
    qc_path = args.qc or newest(review, "05_Daily_QC_*.xlsx")
    dem_path = args.dem_qc or newest(review, "06_Station_DEM_QC_*.xlsx")
    tiles = args.tiles or review / "06_DEM" / "tiles"
    provider_path = args.meteoswiss_metadata
    default_meteo = project / "01_Data_downloaded" / "02_MeteoSwiss" / "MeteoSwiss_Metadata.xlsx"
    if provider_path is None and default_meteo.is_file():
        provider_path = default_meteo

    register = book_tables(register_path, ("Stations", "Products"))
    qc = book_tables(qc_path, ("Summary", "Station coverage", "Extreme days", "Missing runs",
                               "Zero runs", "Event neighbours"))
    previous = book_tables(dem_path, ("Summary", "Station DEM", "Extreme review", "DEM sources",
                                     "Missing runs", "Recorded zero runs"))
    stations = indexed(register["Stations"], "dataset_station_id")
    release = find_update_release(project, args.release)
    updates = applied_update_locations(release, stations)
    products = indexed((r for r in register["Products"] if r["resolution"] == "daily"), "dataset_station_id")
    coverage = indexed(qc["Station coverage"], "Station ID")
    prior_dem = indexed(previous["Station DEM"], "Station ID")
    verify_snapshot(dataset, stations, products, coverage, prior_dem)
    summary = indexed(qc["Summary"], "Item")
    events, prior_events = sheet_rows(qc, previous, summary)
    if any((row["Target station ID"], row["Target date"]) not in events
           or row["Neighbour station ID"] not in stations for row in qc["Event neighbours"]):
        raise ValueError("An event neighbour is absent from the station or extreme-event index.")
    print(f"Verified {len(stations)} daily files against the QC workbook.")

    alternatives = candidate_rows(stations, provider_path, args.candidates, args.nearby_km)
    positions = dict(stations)
    positions.update({("candidate", i): row for i, (_, row, _) in enumerate(alternatives)})
    positions.update({("update", sid): entry[0] for sid, entry in updates.items()})
    samples, tiles_read = dem_samples(tiles, positions, args.flag_m)
    if {r[0]: r[2].lower() for r in tiles_read} != {
        r["Tile"]: r["SHA-256"].lower() for r in previous["DEM sources"]
    }:
        raise ValueError("Local DEM tile names or checksums differ from the previous report.")
    for sid, old in prior_dem.items():
        if not same_number(samples[sid]["dem"], old["DEM elevation (m)"]):
            raise ValueError(f"DEM sample changed unexpectedly for {sid}.")
    station_rows, followup = [], []
    candidate_counts = Counter(sid for sid, _, _ in alternatives)
    for sid, st in sorted(stations.items()):
        d, obs = samples[sid], coverage[sid]
        flags = []
        if d["provisional"]:
            flags.append("Verify CRS")
        if d["status"] != "Below screening threshold":
            flags.append(d["status"])
        station_rows.append([sid, st["station_name"], st["network"], st["provider_id"],
                             st["matched_meteoswiss_code"], st["identity_status"], st["longitude"],
                             st["latitude"], st["elevation_m"], st["coordinate_system"], d["provisional"],
                             st["coordinate_basis"], st["location_source"] or st["metadata_source"],
                             st["location_years"], st["location_verification"], d["dem"], d["delta"],
                             abs(d["delta"]) if d["delta"] is not None else None,
                             d["local_min"], d["local_max"], d["tile"], d["status"], candidate_counts[sid],
                             obs["Observed days"], obs["Extreme days"], obs["Missing runs"], obs["Zero runs"]])
        if flags:
            followup.append([sid, st["station_name"], st["network"], " | ".join(flags),
                             d["delta"], candidate_counts[sid], "Pending", None])
    followup.sort(key=lambda row: (-(abs(row[4]) if row[4] is not None else -1), row[0]))

    fd, temporary_name = mkstemp(prefix=f"06_Station_QC_Combined_{datetime.now():%Y%m%d_%H%M%S}_",
                                  suffix=".part.xlsx", dir=review)
    os.close(fd)
    temporary = Path(temporary_name)
    final = temporary.with_name(temporary.name.replace(".part.xlsx", ".xlsx"))
    try:
        with Workbook(temporary, {"constant_memory": True, "strings_to_formulas": False,
                                   "strings_to_urls": False}) as wb:
            overview = [
                ("Daily QC period", summary["Period"]["Result"]),
                ("Daily files verified by SHA-256", len(stations)),
                *[(f"Station records — {n}", count) for n, count in sorted(
                    Counter(st["network"] for st in stations.values()).items())],
                ("Extreme cutoff, strict > mm/day", summary["Extreme cutoff, strict > mm/day"]["Result"]),
                ("Extreme days", len(events)),
                ("Minimum missing run days", summary["Minimum missing run days"]["Result"]),
                ("Missing runs", len(qc["Missing runs"])),
                ("Minimum recorded zero run days", summary["Minimum recorded zero run days"]["Result"]),
                ("Recorded zero runs", len(qc["Zero runs"])),
                ("Valid registered DEM samples", sum(samples[sid]["dem"] is not None for sid in stations)),
                ("Absolute elevation review threshold (m)", args.flag_m),
                ("Station elevation differences at or above threshold", sum(
                    samples[sid]["status"] == "Review elevation difference" for sid in stations)),
                ("Registered stations with provisional CRS", sum(samples[sid]["provisional"] for sid in stations)),
                ("IH-GAR records with an applied update", len(updates)),
                ("IH-GAR update coordinates", "MeteoSwiss codes, coordinates and update-day counts come from the verified release ledger and replay tables. Counts describe that release."),
                ("DEM coordinates", "Pixel-centre coordinates are in EPSG:4326. Samples use the pixel containing each station point. IH-GAR distances and DEM comparisons assume WGS84 where the archive CRS is unverified."),
                ("Coordinate evidence rows", len(alternatives)),
                ("Nearby search radius (km)", args.nearby_km),
                ("Coordinate evidence", "Nearest registered gauge per other network within radius, plus available provider and documented candidates."),
                ("Coordinate decisions", "Confirm the sensor, site and effective dates from station metadata. A product link or proximity alone leaves identity unresolved."),
                ("Elevation difference", "Reported station elevation minus sampled DEM. Review datum, coordinate precision, slope, and site history before approving any change."),
                ("DEM sampling", "Containing pixel, with valid 3 × 3 local elevation range on the registered site. Unverified archive CRS samples and distances are provisional."),
                ("Gaps and zeros", "Missing runs are inside first-to-last observed coverage. Zero runs are observed zero precipitation. Use the daily QC summary for edge rules."),
                ("Extreme review", "Compare documented accumulation windows and the source observations. Blank values mean no reported evidence for that comparison."),
            ]
            table(wb, "Overview", ["Item", "Value"], overview, [49, 132])
            write_update_locations(wb, updates, stations, samples, args.flag_m)
            table(wb, "Station DEM", [
                "Station ID", "Station", "Network", "Provider ID", "MeteoSwiss link", "Identity status",
                "Longitude", "Latitude", "Registered elevation (m)", "Registered CRS", "CRS provisional",
                "Coordinate basis", "Location source", "Location years", "Location verification",
                "DEM elevation (m)", "Register minus DEM (m)", "Absolute difference (m)", "Local DEM minimum (m)",
                "Local DEM maximum (m)", "Tile", "DEM screen", "Candidate rows", "Observed days",
                "Extreme days", "Missing runs", "Zero runs",
            ], station_rows)
            table(wb, "Coordinate evidence", [
                "Station ID", "Station", "Network", "Evidence", "Candidate network", "Candidate station ID",
                "Candidate provider ID", "Candidate station", "Registered longitude", "Registered latitude",
                "Candidate longitude", "Candidate latitude", "Candidate CRS", "Distance (m)",
                "Registered elevation (m)", "Candidate elevation (m)", "Registered DEM (m)", "Candidate DEM (m)",
                "Register minus DEM (m)", "Candidate minus DEM (m)", "Candidate DEM screen",
                "Valid from", "Valid to", "Location years", "Evidence source", "Decision", "Reviewer notes",
            ], candidate_details(alternatives, stations, samples))
            event_headers = ["Station ID", "Station", "Date", "Rainfall (mm)", "Network", "Product",
                             "UTC window", "UTC start", "UTC end", "DEM (m)", "Same-window neighbours",
                             "Same-window maximum (mm)", "Isone open gauge (mm)", "Isone open window",
                             "Source file", "Hourly archive", "Decision", "Reviewer notes"]
            table(wb, "Extreme review", event_headers, (
                [row[c] for c in event_headers[:9]] + [samples[key[0]]["dem"]] +
                [prior_events[key][c] for c in event_headers[10:14]] +
                [row["Source file"], prior_events[key]["Hourly archive"], "Pending source review", None]
                for key, row in events.items()
            ))
            neighbour_headers = ["Target station ID", "Target date", "Target rainfall (mm)", "Rank",
                                 "Neighbour station ID", "Neighbour", "Network", "Distance (km)", "Elevation (m)",
                                 "D−1 (mm)", "D (mm)", "D+1 (mm)", "UTC window", "Window comparison", "UTC start",
                                 "UTC end", "Neighbour source", "Source evidence", "Coordinate CRS"]
            copy_table(wb, "Event neighbours", qc["Event neighbours"], neighbour_headers)
            run_headers = ["Station ID", "Station", "Network", "First date", "Last date", "Consecutive days",
                           "Touches start", "Touches end", "Final file"]
            copy_table(wb, "Missing runs", qc["Missing runs"], run_headers)
            copy_table(wb, "Recorded zero runs", qc["Zero runs"], run_headers)
            table(wb, "Coordinate review", ["Station ID", "Station", "Network", "Review reason",
                                             "Register minus DEM (m)", "Candidate rows", "Decision", "Reviewer notes"],
                  followup)
            table(wb, "DEM sources", ["Tile", "Bytes", "SHA-256", "Horizontal CRS"], tiles_read, [60, 16, 69, 22])
            inputs = [("Station register", register_path), ("Daily QC", qc_path), ("Prior DEM review", dem_path),
                      ("MeteoSwiss metadata", provider_path), ("Documented candidates", args.candidates),
                      ("Verified IH-GAR update archive", release)]
            table(wb, "Inputs", ["Input", "Path", "SHA-256"],
                  ([label, str(path.resolve()) if path else "Unavailable or optional input omitted",
                    digest(path) if path else None] for label, path in inputs), [28, 100, 70])

        # Check the workbook can be opened before exposing the final filename.
        verified = load_workbook(temporary, read_only=True, data_only=True)
        try:
            expected = {UPDATE_SHEET: len(updates), "Station DEM": len(stations), "Coordinate evidence": len(alternatives),
                        "Extreme review": len(events), "Missing runs": len(qc["Missing runs"]),
                        "Recorded zero runs": len(qc["Zero runs"])}
            for title, count in expected.items():
                if verified[title].max_row != count + 1:
                    raise ValueError(f"Unexpected output row count in {title}")
        finally:
            verified.close()
        temporary.rename(final)
    finally:
        temporary.unlink(missing_ok=True)
    print(f"Workbook: {final}")
    return final


def newest(folder, pattern):
    paths = sorted((p for p in folder.glob(pattern) if p.is_file() and not p.name.endswith(".part.xlsx")),
                   key=lambda p: (p.stat().st_mtime_ns, p.name), reverse=True)
    if not paths:
        raise ValueError(f"No report matching {pattern} in {folder}. Supply the report path explicitly.")
    return paths[0]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project", type=Path, required=True)
    parser.add_argument("--qc", type=Path, help="Daily QC XLSX. Default: latest 05_Daily_QC_*.xlsx.")
    parser.add_argument("--dem-qc", type=Path, help="Prior DEM QC XLSX. Default: latest 06_Station_DEM_QC_*.xlsx.")
    parser.add_argument("--tiles", type=Path, help="GLO-30 tiles. Default: 04_Analysis_Review/06_DEM/tiles.")
    parser.add_argument("--meteoswiss-metadata", type=Path, help="Override the MeteoSwiss metadata XLSX path.")
    parser.add_argument("--candidates", type=Path, help="Optional CSV of documented candidate locations.")
    parser.add_argument("--release", type=Path, help="Verified IH-GAR update archive, if outside the project or Downloads.")
    parser.add_argument("--flag-m", type=float, default=100, help="Absolute elevation difference in metres for review.")
    parser.add_argument("--nearby-km", type=float, default=2, help="Radius for candidate gauges from other networks.")
    args = parser.parse_args()
    if any(not math.isfinite(v) or v <= 0 for v in (args.flag_m, args.nearby_km)):
        parser.error("--flag-m and --nearby-km must both be positive finite values.")
    package = find_spec("rasterio")
    if package is None:
        parser.error("Install dependencies with: py -3 -m pip install rasterio openpyxl XlsxWriter")
    bundled_proj = Path(package.origin).parent / "proj_data"
    if (bundled_proj / "proj.db").is_file():
        # Keep Rasterio on the PROJ database bundled with its Python package.
        os.environ["PROJ_DATA"] = os.environ["PROJ_LIB"] = str(bundled_proj)
    build(args)


if __name__ == "__main__":
    try:
        main()
    except (OSError, ValueError, KeyError, BadZipFile) as exc:
        sys.exit(f"Combined station QC stopped: {exc}")
