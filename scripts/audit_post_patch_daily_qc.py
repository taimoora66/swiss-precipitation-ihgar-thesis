"""Report daily extremes, internal missing runs, zero runs and nearby observations."""
import argparse
from array import array
import calendar
from collections import defaultdict
import csv
from dataclasses import dataclass
from datetime import date, datetime
import gzip
from hashlib import sha256
import io
import json
import math
from pathlib import Path, PureWindowsPath
import sys
from tempfile import mkstemp
from zipfile import BadZipFile, ZipFile
import os

from openpyxl import load_workbook
from xlsxwriter import Workbook


MISSING = {"", "na", "nan", "-9999", "-9999.0"}
RELEASE_NAME = "IHGAR_MeteoSwiss_Station_Update.zip"
RELEASE_SHA = "c4eb5b2ae66df1a6cfb6e1a50077cb5bc4de99259c5984f130d60f9b5d69a838"
DAY_SOURCE = {"rre150d0": "06 UTC", "rka150d0": "00 UTC"}
WINDOW_LABEL = {"06:00 UTC to following 06:00 UTC": "06 UTC",
                "00:00-24:00 UTC calendar day": "00 UTC"}


def tabular_rows(book, sheet, fields):
    rows = book[sheet].values
    headers = next(rows, ())
    if not set(fields).issubset(headers):
        raise ValueError(f"{sheet} is missing columns: {', '.join(sorted(set(fields) - set(headers)))}")
    return [dict(zip(headers, row)) for row in rows if any(v is not None for v in row)]


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


def iso(day):
    return date.fromordinal(day).isoformat() if day is not None else None


def finite_number(value):
    if value is None or value == "":
        return None
    try:
        number = float(value)
    except (ValueError, TypeError):
        return None
    return number if math.isfinite(number) else None


@dataclass
class Station:
    info: dict
    product: dict
    values: array
    start: int
    first: int | None
    last: int | None
    lon: float | None
    lat: float | None
    elev: float | None

    @property
    def id(self):
        return str(self.info["dataset_station_id"])

    @property
    def name(self):
        return self.info["station_name"]

    @property
    def network(self):
        return str(self.info["network"])

    @property
    def meteo_code(self):
        code = (self.info["provider_id"] if self.network == "MeteoSwiss"
                else self.info["matched_meteoswiss_code"] if self.network == "IHGAR" else None)
        return str(code).strip().upper() if code else None

    def value(self, day):
        i = day - self.start
        if 0 <= i < len(self.values):
            value = self.values[i]
            return value if not math.isnan(value) else None
        return None


def unique(rows, key, label):
    mapping = {}
    for row in rows:
        v = row[key]
        if not v or v in mapping:
            raise ValueError(f"Missing or duplicated {label}: {v!r}")
        mapping[v] = row
    return mapping


def applied_patch(project, dataset):
    parent = project / "working" / "ihgar_checked_exclusions"
    for path in sorted(parent.glob("*/Manifest.json"), reverse=True):
        manifest = json.loads(path.read_text(encoding="utf-8-sig"))
        if manifest.get("status") != "APPLIED_AND_VERIFIED":
            continue
        if Path(manifest["dataset"]).resolve() != dataset:
            raise ValueError(f"Applied patch manifest points to a different dataset: {path}")
        daily = {k: v["after"] for k, v in manifest["files"].items() if k.startswith("Daily/")}
        if not daily or len(daily) != manifest["stations"]:
            raise ValueError("Applied manifest has inconsistent daily file counts")
        return path, daily
    raise ValueError("No completed Checked IH-GAR patch found in working/ihgar_checked_exclusions")


def read_stations(dataset, patch_hashes, start, end):
    wb = load_workbook(dataset / "Station_Register.xlsx", read_only=True, data_only=True)
    try:
        infos = tabular_rows(wb, "Stations", ["dataset_station_id", "station_name", "network",
                  "provider_id", "matched_meteoswiss_code", "longitude", "latitude", "elevation_m",
                  "coordinate_system", "daily_file"])
        products = unique((r for r in tabular_rows(wb, "Products", ["dataset_station_id",
                          "source_network", "resolution", "relative_path", "sha256", "valid_observations",
                          "units", "source_path", "accumulation_interval"]) if r["resolution"] == "daily"),
                          "dataset_station_id", "daily product")
    finally:
        wb.close()
    metadata = unique(infos, "dataset_station_id", "station")
    if not products or metadata.keys() != products.keys():
        raise ValueError("Station and daily product IDs do not match one-to-one")
    stations, paths = [], set()
    for n, (sid, row) in enumerate(sorted(metadata.items()), 1):
        product = products[sid]
        path = daily_path(dataset, product["relative_path"])
        if daily_path(dataset, row["daily_file"]) != path:
            raise ValueError(f"Station and product path differ: {sid}")
        key = path.relative_to(dataset).as_posix()
        if key in paths or str(product["units"]).strip().lower() != "mm" or row["network"] != product["source_network"]:
            raise ValueError(f"Duplicate file, inconsistent network, or units other than mm: {sid}")
        paths.add(key)
        data = path.read_bytes()
        calculated = sha256(data).hexdigest()
        if calculated != str(product["sha256"]).strip().lower():
            raise ValueError(f"Daily file differs from Station_Register.xlsx: {path}")
        if key in patch_hashes and calculated != patch_hashes[key]:
            raise ValueError(f"Final daily file differs from verified patch: {path}")
        try:
            grid = daily_grid(data)
        except ValueError as exc:
            raise ValueError(f"{sid} ({path.name}): {exc}") from exc
        observed = [day for day, value in grid.items() if value is not None]
        if len(observed) != product["valid_observations"]:
            raise ValueError(f"Daily observation count differs from register: {sid}")
        # Keep one day of context on either side of the selected period.
        buffer_start = max(1, start - 1)
        values = array("d", [math.nan]) * (min(date.max.toordinal(), end + 1) - buffer_start + 1)
        for day in observed:
            if 0 <= day - buffer_start < len(values):
                values[day - buffer_start] = grid[day]
        lon, lat, elev = (finite_number(row[k]) for k in ("longitude", "latitude", "elevation_m"))
        if lon is not None and lat is not None and not (-180 <= lon <= 180 and -90 <= lat <= 90):
            lon = lat = None
        stations.append(Station(row, product, values, buffer_start, min(observed, default=None),
                                max(observed, default=None), lon, lat, elev))
        if n % 50 == 0 or n == len(metadata):
            print(f"Read and verified {n}/{len(metadata)} final daily files", flush=True)
    if patch_hashes.keys() - paths:
        raise ValueError("One or more patched files are missing from the register")
    return stations


def scan(station, start, end, extreme_mm, minimum_missing, minimum_zero):
    first = max(start, station.first) if station.first is not None else end + 1
    last = min(end, station.last) if station.last is not None else start - 1
    summary = {"observed": 0, "zero": 0, "missing": 0, "outside_span": end - start + 1,
               "highest": None}
    runs = {"Missing": [], "Zero": []}
    extremes = []
    active, run_start = None, None
    for day in range(first, last + 2) if first <= last else ():
        state = None
        if day <= last:
            value = station.value(day)
            state = "Missing" if value is None else "Zero" if value == 0 else None
            if value is None:
                summary["missing"] += 1
            else:
                summary["observed"] += 1
                summary["zero"] += value == 0
                high = summary["highest"]
                summary["highest"] = value if high is None else max(value, high)
                if value > extreme_mm:
                    extremes.append((day, value))
        if state != active:
            if active is not None:
                length = day - run_start
                cutoff = minimum_missing if active == "Missing" else minimum_zero
                if length >= cutoff:
                    runs[active].append((run_start, day - 1, length, run_start == start, day - 1 == end))
            active, run_start = state, day
    summary["outside_span"] -= summary["observed"] + summary["missing"]
    return summary, extremes, runs


def release_ledger(dataset, explicit, wanted):
    options = [explicit.resolve()] if explicit else list((dataset / "Release_candidates").rglob(RELEASE_NAME))
    if not explicit:
        options += list((Path.home() / "Downloads").glob("IHGAR_MeteoSwiss_Station_Update*.zip"))
    release = next((p for p in options if p.is_file() and sha256(p.read_bytes()).hexdigest() == RELEASE_SHA), None)
    if explicit and release is None:
        raise ValueError("--release does not match the verified IH-GAR update release")
    if release is None:
        return None, {}
    result = {}
    with ZipFile(release) as archive, archive.open("Applied_Daily_Changes.csv.gz") as raw:
        with gzip.GzipFile(fileobj=raw) as compressed, io.TextIOWrapper(compressed, encoding="utf-8-sig") as stream:
            for row in csv.DictReader(stream):
                if row["Date"] not in wanted:
                    continue
                key = row["IH-GAR ID"], row["Date"]
                if key in result:
                    raise ValueError(f"Duplicate source ledger entry: {key}")
                result[key] = row
    return release, result


def day_source(station, day, ledger):
    row = ledger.get((str(station.info["provider_id"]), iso(day))) if station.network == "IHGAR" else None
    registered = station.product["accumulation_interval"]
    path = station.product["source_path"]
    if row:
        logged = finite_number(row["Added token"])
        if logged == station.value(day):
            return (row["Source path"], row["Product"], DAY_SOURCE.get(row["Product"]),
                    row["UTC start"], row["UTC end"], "Exact date in applied source ledger")
        return path, None, None, None, None, "Applied source ledger amount differs from final file"
    known = WINDOW_LABEL.get(registered)
    return (path, None, known, None, None,
            "Registered product window" if known else "Day-specific window unresolved")


def distance_km(a, b):
    if None in (a.lon, a.lat, b.lon, b.lat):
        return None
    p1, p2 = math.radians(a.lat), math.radians(b.lat)
    x = math.sin((p2 - p1) / 2)**2 + math.cos(p1) * math.cos(p2) * math.sin(math.radians(b.lon - a.lon) / 2)**2
    return 6371.0088 * 2 * math.asin(math.sqrt(min(1, max(0, x))))


def neighbours(target, stations, day, limit, radius, ledger):
    ranked = []
    for ref in stations:
        if ref.id == target.id or ref.value(day) is None:
            continue
        km = distance_km(target, ref)
        if km is None or not 0.25 <= km <= radius:
            continue
        if target.meteo_code and target.meteo_code == ref.meteo_code:
            continue
        ranked.append((km, ref.id, ref))
    result = []
    target_window = day_source(target, day, ledger)[2]
    for km, _, ref in sorted(ranked)[:limit]:
        ref_source, _, window, ref_start, ref_end, quality = day_source(ref, day, ledger)
        label = ("Window unresolved" if not target_window or not window else
                 "Same documented UTC window" if target_window == window else "Different UTC windows")
        result.append((ref.id, ref.name, ref.network, round(km, 2), ref.elev,
                       ref.value(day - 1), ref.value(day), ref.value(day + 1), window, label,
                       ref_start, ref_end, ref_source, quality, ref.info["coordinate_system"]))
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


def report(args, project, dataset, manifest, stations):
    first, last = args.start.toordinal(), args.end.toordinal()
    coverage, events = [], []
    all_runs = defaultdict(list)
    for station in stations:
        stats, extremes, runs = scan(station, first, last, args.extreme_mm, args.missing_days, args.zero_days)
        coverage.append([station.id, station.name, station.network, iso(station.first), iso(station.last),
                         stats["observed"], stats["missing"], stats["outside_span"], stats["zero"],
                         stats["highest"], len(extremes), len(runs["Missing"]), len(runs["Zero"]),
                         station.lon, station.lat, station.elev, station.info["coordinate_system"],
                         station.product["relative_path"], station.product["sha256"]])
        events.extend((day, amount, station) for day, amount in extremes)
        for category, group in runs.items():
            all_runs[category].extend([station.id, station.name, station.network, iso(a), iso(b),
                                       n, left, right, station.product["relative_path"]]
                                      for a, b, n, left, right in group)
    dates = {iso(day) for day, _, _ in events}
    release, ledger = release_ledger(dataset, args.release, dates) if dates else (None, {})
    event_rows, neighbor_rows = [], []
    for day, amount, st in sorted(events, key=lambda item: (-item[1], item[0], item[2].id)):
        source, prod, win, utc_start, utc_end, evidence = day_source(st, day, ledger)
        close = neighbours(st, stations, day, args.neighbours, args.radius_km, ledger)
        event_rows.append([st.id, st.name, st.network, iso(day), amount, st.value(day - 1), st.value(day + 1),
                           len(close), sum(row[9] == "Same documented UTC window" for row in close),
                           prod, win, utc_start, utc_end, evidence, source, st.product["relative_path"],
                           st.lon, st.lat, st.elev, st.info["coordinate_system"], "Pending manual review"])
        neighbor_rows.extend([st.id, iso(day), amount, rank, *row] for rank, row in enumerate(close, 1))
    output = project / "04_Analysis_Review"
    output.mkdir(parents=True, exist_ok=True)
    if output.resolve().is_relative_to(dataset):
        raise ValueError("The review folder resolves inside the input dataset")
    fd, temp_name = mkstemp(prefix="05_Daily_QC_", suffix=".part.xlsx", dir=output)
    os.close(fd)
    temp = Path(temp_name)
    result = output / f"05_Daily_QC_{datetime.now():%Y%m%d_%H%M%S}_{temp.stem.rsplit('_',1)[-1].split('.')[0]}.xlsx"
    try:
        with Workbook(temp, {"constant_memory": True, "strings_to_formulas": False, "strings_to_urls": False}) as wb:
            metrics = [
                ["Period", f"{args.start.isoformat()} to {args.end.isoformat()}"],
                ["Final daily stations checked", len(stations)], ["Extreme days", len(event_rows)],
                ["Stations with an extreme", len({st.id for _, _, st in events})],
                ["Missing runs", len(all_runs["Missing"])], ["Recorded zero runs", len(all_runs["Zero"])],
                ["Nearby observations listed", len(neighbor_rows)], ["Extreme cutoff, strict > mm/day", args.extreme_mm],
                ["Minimum missing run days", args.missing_days], ["Minimum recorded zero run days", args.zero_days],
                ["Maximum neighbours per event", args.neighbours], ["Neighbour radius (km)", args.radius_km],
                ["Final dataset", str(dataset)], ["Applied patch manifest", str(manifest)],
                ["Exact-day source ledger", str(release) if release else
                    ("Unavailable; register metadata used" if events else "No extreme events required a source lookup")],
                ["Missing run rule", "Calendar dates without values inside the first-to-last observed span, clipped to the selected period."],
                ["Outside observed span", "Report as coverage only. Leading or trailing dates outside observed span are not internal gaps."],
                ["Recorded zeros", "Observed values equal to zero. A missing day or a positive observation ends a recorded zero run."],
                ["Period edge", "Run touches a reporting date boundary; displayed duration is within the selected period."],
                ["Nearby observations", "Up to the selected number of closest stations with a same-label observation inside the radius."],
                ["Reference safeguards", "Shared MeteoSwiss codes and sites closer than 250 m are omitted. Candidate coordinates may still need verification."],
                ["Accumulation windows", "Different or unresolved UTC windows need source review. D−1/D+1 give date-label context."],
                ["Decision", "Every extreme remains pending. Observations are never edited by this report."],
            ]
            write_sheet(wb, "Summary", ["Item", "Result"], metrics, [42, 117])
            write_sheet(wb, "Extreme days", ["Station ID","Station","Network","Date","Rainfall (mm)",
                        "D−1 (mm)","D+1 (mm)","Neighbours","Same-window neighbours","Product","UTC window",
                        "UTC start","UTC end","Source evidence","Source file","Final file","Longitude","Latitude",
                        "Elevation (m)","Coordinate CRS","Decision"], event_rows, [45,24,17,15,15,13,13,12,15,14,15,26,26,47,70,60,15,15,15,43,25])
            for kind, name in (("Missing", "Missing runs"), ("Zero", "Zero runs")):
                write_sheet(wb, name, ["Station ID","Station","Network","First date","Last date",
                            "Consecutive days","Touches start","Touches end","Final file"],
                            sorted(all_runs[kind], key=lambda row: (-row[5], row[0], row[3])),
                            [45,25,17,15,15,17,15,15,62])
            write_sheet(wb, "Event neighbours", ["Target station ID","Target date","Target rainfall (mm)","Rank",
                        "Neighbour station ID","Neighbour","Network","Distance (km)","Elevation (m)",
                        "D−1 (mm)","D (mm)","D+1 (mm)","UTC window","Window comparison","UTC start","UTC end",
                        "Neighbour source","Source evidence","Coordinate CRS"], neighbor_rows,
                        [45,15,18,10,50,25,16,15,15,14,14,14,15,31,26,26,65,44,40])
            write_sheet(wb, "Station coverage", ["Station ID","Station","Network","First observed",
                        "Last observed","Observed days","Internal missing days","Outside observed span days","Zero days",
                        "Maximum (mm)","Extreme days","Missing runs","Zero runs","Longitude","Latitude",
                        "Elevation (m)","Coordinate CRS","Final file","SHA-256"], coverage,
                        [45,25,17,16,16,15,17,20,15,15,15,15,15,15,15,16,40,62,68])
        temp.rename(result)
    finally:
        temp.unlink(missing_ok=True)
    print(f"Extreme days: {len(event_rows)} | Missing runs: {len(all_runs['Missing'])} | Zero runs: {len(all_runs['Zero'])}")
    print(f"Workbook: {result}")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project", type=Path, required=True)
    parser.add_argument("--start", type=date.fromisoformat, default=date(1951, 1, 1))
    parser.add_argument("--end", type=date.fromisoformat, default=date(2023, 12, 31))
    parser.add_argument("--extreme-mm", type=float, default=600)
    parser.add_argument("--missing-days", type=int, default=90)
    parser.add_argument("--zero-days", type=int, default=90)
    parser.add_argument("--neighbours", type=int, default=15)
    parser.add_argument("--radius-km", type=float, default=100)
    parser.add_argument("--release", type=Path, help="Optional verified update ZIP for exact-day provenance")
    args = parser.parse_args()
    if (args.start > args.end or min(args.missing_days, args.zero_days, args.neighbours) < 1
            or not all(math.isfinite(x) and x > 0 for x in (args.extreme_mm, args.radius_km))):
        parser.error("Use ordered dates and positive finite thresholds")
    project = args.project.resolve()
    dataset = (project / "03_Daily_dataset").resolve()
    if (dataset / ".ihgar_checked_patch_lock").exists():
        raise ValueError("Daily patch lock exists. Finish patch verification first.")
    path, hashes = applied_patch(project, dataset)
    stations = read_stations(dataset, hashes, args.start.toordinal(), args.end.toordinal())
    report(args, project, dataset, path, stations)


if __name__ == "__main__":
    try:
        main()
    except (OSError, ValueError, KeyError, BadZipFile) as exc:
        sys.exit(f"Daily QC stopped: {exc}")
