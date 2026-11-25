"""Apply verified Checked IH-GAR exclusions to final daily files. Default: dry run."""

import argparse
from array import array
import calendar
from collections import Counter
from contextlib import contextmanager
import csv
from datetime import date, datetime
import hashlib
import io
import json
import math
import os
import shutil
from pathlib import Path, PureWindowsPath
import sys
import tempfile
from zipfile import ZipFile

from openpyxl import load_workbook

def number(token):
    token = token.strip()
    if token.upper() in ("", "NA", "NAN", "-9999"):
        return math.nan
    value = float(token)
    if not math.isfinite(value) or value < 0:
        raise ValueError(f"Unexpected precipitation token {token!r}")
    return value


@contextmanager
def checked_stream(path):
    if path.suffix.lower() == ".zip":
        with ZipFile(path) as archive:
            members = [n for n in archive.namelist() if n.lower().endswith(".csv")]
            if len(members) != 1:
                raise ValueError("The checked ZIP must contain one CSV. Pass the extracted CSV if it contains several.")
            with io.TextIOWrapper(archive.open(members[0]), encoding="utf-8-sig", newline="") as stream:
                yield stream
    else:
        with path.open(encoding="utf-8-sig", newline="") as stream:
            yield stream


def checked_values(path):
    with checked_stream(path) as stream:
        reader = csv.reader(stream, delimiter=";")
        header = next(reader)
        ids = [name.strip() for name in header[3:]]
        if header[:3] != ["Anno", "Mese", "Giorno"] or not ids or any(not x.startswith("PCPD_") for x in ids) or len(ids) != len(set(ids)):
            raise ValueError("Expected semicolon-delimited Anno;Mese;Giorno and unique PCPD station IDs.")
        values = [array("d") for _ in ids]
        first = last = None
        for row in reader:
            if not row:
                continue
            if len(row) != len(header):
                raise ValueError(f"Checked CSV line {reader.line_num} has the wrong number of columns.")
            day = date(*map(int, row[:3])).toordinal()
            if last is not None and day != last + 1:
                raise ValueError(f"Checked CSV calendar has a gap, duplicate or out-of-order date at line {reader.line_num}.")
            if first is None:
                first = day
            last = day
            for column, token in zip(values, row[3:]):
                column.append(number(token))
    if first is None:
        raise ValueError("The checked CSV has no data rows.")
    return first, last, dict(zip(ids, values))


@contextmanager
def original_archive(path):
    if path.is_dir():
        files = [p for p in path.rglob("PCPD_*") if p.is_file()]
        entries = {p.name: p for p in files}
        load = lambda member: member.read_bytes()
    else:
        with ZipFile(path) as archive:
            files = [n for n in archive.namelist() if not n.endswith("/") and Path(n).name.startswith("PCPD_")]
            entries = {Path(n).name: n for n in files}
            if len(entries) != len(files) or not entries:
                raise ValueError("Original archive must contain uniquely named PCPD daily files.")
            yield entries, archive.read
        return
    if len(entries) != len(files) or not entries:
        raise ValueError("Original folder must contain uniquely named PCPD daily files.")
    yield entries, load


def station_values(blob, station):
    result = {}
    months = set()
    for line_no, line in enumerate(blob.decode("utf-8-sig").splitlines(), 1):
        if not line.strip():
            continue
        fields = line.split("\t")
        if len(fields) != 33:
            raise ValueError(f"{station}, line {line_no}: expected year, month and 31 tab-separated fields.")
        year, month = map(int, fields[:2])
        if (year, month) in months:
            raise ValueError(f"{station}: duplicate month {year}-{month}.")
        months.add((year, month))
        n_days = calendar.monthrange(year, month)[1]
        start = date(year, month, 1).toordinal()
        for offset, token in enumerate(fields[2:]):
            value = number(token)
            if math.isnan(value):
                continue
            if offset >= n_days:
                raise ValueError(f"{station}, line {line_no}: numeric value on a nonexistent date.")
            result[start + offset] = value
    return result


STUDY_FIRST = date(1951, 1, 1).toordinal()
STUDY_LAST = date(2023, 12, 31).toordinal()
TOLERANCE = 1e-8
CHECKED_NAME = "Swiss_series_IH_GAR_checked 1(in).zip"


def sha256(data):
    return hashlib.sha256(data).hexdigest()


def iso(day):
    return date.fromordinal(day).isoformat() if day else None


def first_last(values):
    return [iso(min(values, default=0)), iso(max(values, default=0)), len(values)]


def worksheet_rows(book, name, required):
    rows = iter(book[name].values)
    header = next(rows)
    if not required <= set(header) or len(header) != len(set(header)):
        raise ValueError(f"Invalid or incomplete headers in {name}")
    return [dict(zip(header, row)) for row in rows if any(x is not None for x in row)]


def get_register(path):
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        stations = [r for r in worksheet_rows(book, "Stations", {
            "network", "provider_id", "station_name", "dataset_station_id", "update_status"
        }) if r["network"] == "IHGAR"]
        products = [r for r in worksheet_rows(book, "Products", {
            "resolution", "source_network", "dataset_station_id", "relative_path",
            "source_path", "sha256", "valid_observations"
        }) if (r["resolution"], r["source_network"]) == ("daily", "IHGAR")]
    finally:
        book.close()
    by_id = {r["dataset_station_id"]: r for r in products}
    ids = {s["dataset_station_id"] for s in stations}
    if not stations or len(ids) != len(stations) or len(by_id) != len(products) or ids != by_id.keys():
        raise ValueError("Expected one primary daily file for each IH-GAR register identity")
    return sorted(stations, key=lambda s: s["provider_id"]), by_id


def relative_path(value):
    path = PureWindowsPath(value)
    if not path.parts or path.anchor or ".." in path.parts:
        raise ValueError(f"Unsafe relative path in the register: {value}")
    return Path(*path.parts)


def read_updated(root, station, product, original):
    assembled = root / "03_Daily_dataset" / relative_path(product["relative_path"])
    if assembled.is_file():
        data, source = assembled.read_bytes(), "Assembled Daily"
    else:
        parts = PureWindowsPath(product["source_path"]).parts
        if "02_Data_extraction" in parts:
            rel = relative_path(str(PureWindowsPath(*parts[parts.index("02_Data_extraction"):])))
            extracted = root / rel
        else:
            extracted = None
        if extracted is not None and extracted.is_file():
            data, source = extracted.read_bytes(), "Extracted fallback"
        elif station["update_status"] == "ORIGINAL_UNCHANGED":
            data, source = original, "Original archive fallback"
        else:
            raise FileNotFoundError(f"Missing registered daily file for {station['provider_id']}: {assembled}")
    if sha256(data) != str(product["sha256"]).lower():
        raise ValueError(f"Registered SHA-256 mismatch for {station['provider_id']} ({source}). Audit stopped.")
    return data, source


def classify(original, checked, updated, checked_has_observations):
    if original is not None and checked is None:
        if checked_has_observations:
            return "Checked exclusion", "Propose removal" if updated is not None else "Already missing in Updated"
        return "Checked column has no observations", "Review"
    if checked is not None:
        if original is None:
            finding = "Checked-only addition" if updated is None else "Original missing; Checked and Updated observed"
        elif abs(original - checked) > TOLERANCE:
            finding = "Original and Checked numeric difference"
        elif updated is None:
            finding = "Checked value missing in Updated"
        elif abs(checked - updated) > TOLERANCE:
            finding = "Checked and Updated numeric difference"
        else:
            return "Three versions agree", None
        return finding, "Review"
    return "Updated addition on Checked missing date", "Review"


def rows_digest(rows):
    payload = json.dumps(sorted(rows, key=lambda row: (row[0], row[1])),
                         ensure_ascii=False, separators=(",", ":"), allow_nan=False)
    return sha256(payload.encode("utf-8"))


def compare(root, archive, checked_path, register):
    start, end, checked = checked_values(checked_path)
    stations, products = get_register(register)
    full, study, sources = Counter(), Counter(), Counter()
    coverage, exclusions, differences, exclusions_key, differences_key = [], [], [], [], []
    ids = {s["provider_id"] for s in stations}
    if len(ids) != len(stations):
        raise ValueError("Original archive IDs are duplicated in the register")
    with original_archive(archive) as (members, load):
        if ids != members.keys():
            raise ValueError("Original archive station IDs differ from the IH-GAR register")
        for index, station in enumerate(stations, start=1):
            sid = station["provider_id"]
            name = station["station_name"]
            product = products[station["dataset_station_id"]]
            old_bytes = load(members[sid])
            upd_bytes, source = read_updated(root, station, product, old_bytes)
            sources[source] += 1
            old = station_values(old_bytes, sid)
            upd = old if upd_bytes == old_bytes else station_values(upd_bytes, sid)
            if len(upd) != product["valid_observations"]:
                raise ValueError(f"Valid-observation count differs from the register for {sid}")
            col = checked.get(sid)
            obs = {} if col is None else {start+i: v for i, v in enumerate(col) if not math.isnan(v)}
            first, last = min(obs, default=0), max(obs, default=0)
            counts = Counter()
            after, before = [], 0
            for day in sorted(old.keys() | obs.keys() | upd.keys()):
                o, c, u = old.get(day), obs.get(day), upd.get(day)
                in_study = STUDY_FIRST <= day <= STUDY_LAST
                if col is None or not start <= day <= end:
                    finding = "Checked station absent" if col is None else "Outside Checked calendar"
                    if day > end and u is not None:
                        after.append(day)
                        if o is None:
                            counts["Updated additions after Checked"] += 1
                            full["Updated additions after Checked"] += 1
                    before += day < start and u is not None
                    decision = None
                else:
                    finding, decision = classify(o, c, u, bool(obs))
                counts[finding] += 1
                full[finding] += 1
                if in_study:
                    study[finding] += 1
                if decision is None:
                    continue
                date_text = iso(day)
                period = "1951–2023" if in_study else "Outside 1951–2023"
                values = [sid, name, date_text, period, o, c, u]
                if finding == "Checked exclusion":
                    counts[decision] += 1
                    full[decision] += 1
                    if in_study:
                        study[decision] += 1
                    edge = "Within valid span" if first <= day <= last else "Outside valid span"
                    exclusions.append(values + [edge, decision])
                    exclusions_key.append([sid,date_text,o,c,u,decision])
                else:
                    differences.append(values + [finding])
                    differences_key.append([sid,date_text,o,c,u,finding])
            match = "Absent from Checked" if col is None else "Exact station ID" if obs else "Checked column all missing"
            coverage.append([sid,name,station["update_status"],match,source,*first_last(old),
                             *first_last(obs),*first_last(upd),counts["Propose removal"],
                             counts["Already missing in Updated"],counts["Outside Checked calendar"],
                             counts["Checked station absent"],len(after),counts["Updated additions after Checked"],
                             iso(min(after, default=0)),iso(max(after, default=0)),before,product["sha256"]])
            if index % 50 == 0:
                print(f"Audited {index}/{len(stations)} IH-GAR records", flush=True)
    for sid in sorted(checked.keys() - ids):
        obs = [start+i for i,v in enumerate(checked[sid]) if not math.isnan(v)]
        coverage.append([sid,None,None,"Checked-only station",None,None,None,0,
                         *first_last(obs),None,None,0,0,0,0,0,0,0,None,None,0,None])
    removals = [r for r in exclusions if r[8] == "Propose removal"]
    full["Stations with proposed removals"] = len({r[0] for r in removals})
    study["Stations with proposed removals"] = len({r[0] for r in removals if r[3]=="1951–2023"})
    result = dict(
        ihgar_records=len(stations), matched_checked_stations=len(ids & checked.keys()),
        absent_checked_station_ids=sorted(ids-checked.keys()), checked_only_ids=sorted(checked.keys()-ids),
        checked_calendar=[iso(start),iso(end)], tolerance_mm=TOLERANCE,
        input_sha256={k:sha256(p.read_bytes()) for k,p in {
            "original":archive,"checked":checked_path,"register":register
        }.items()}, updated_backing=sources, full_archive=full, study_1951_2023=study,
        exclusion_rows_sha256=rows_digest(exclusions_key),
        other_difference_rows_sha256=rows_digest(differences_key)
    )
    return result, coverage, exclusions, differences


def write_csv(path, header, rows):
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(header)
        writer.writerows(rows)


def unique_input(explicit, default, candidates, flag):
    if explicit is not None:
        if not explicit.is_file():
            raise ValueError(f"{flag} does not point to a file: {explicit}")
        return explicit.resolve()
    if default.is_file():
        return default
    found = sorted({p.resolve() for p in candidates if p.is_file()})
    if len(found) != 1:
        raise ValueError(f"Found {len(found)} candidate files. Provide {flag} with the exact file path.")
    return found[0]


def compare_reference(result):
    measured = {
        "ihgar_records":result["ihgar_records"],
        "matched_checked_stations":result["matched_checked_stations"],
        "input_sha256":result["input_sha256"],
        "full_proposed_removals":result["full_archive"]["Propose removal"],
        "study_proposed_removals":result["study_1951_2023"]["Propose removal"],
        "exclusion_rows_sha256":result["exclusion_rows_sha256"],
        "other_difference_rows_sha256":result["other_difference_rows_sha256"],
    }
    checks = {k:{"expected":v,"local":measured[k],"matches":v==measured[k]} for k,v in REFERENCE.items()}
    return {"status":"MATCH" if all(x["matches"] for x in checks.values()) else "DIFFERENCES", "checks":checks}

REFERENCE = {'ihgar_records': 217,
 'matched_checked_stations': 213,
 'input_sha256': {'original': '8521e488b1361cfb89235fa47894c2a6d3f83ee2d0748bee865a681e4657733e',
                  'checked': '8b416cfe9194fcadab7f51b8899f032807f4429372b11818b2454fd666169c89',
                  'register': '7161b8d5722731fc8aee08ded5c5c93ced63f58c4a132a138f11d3c567e6150a'},
 'full_proposed_removals': 1893,
 'study_proposed_removals': 1824,
 'exclusion_rows_sha256': 'ca78724c31cd3dac5d7133336f339c9cf1dd5eca07e8a283fa623ad6340b03b8',
 'other_difference_rows_sha256': '2dc1339b3c6553c3fe81e30ec531fe024278288168e6ee545d1ea5a1e464d9b7'}


DATASET = "03_Daily_dataset"
REGISTER = "Station_Register.xlsx"
EXPECTED_CHANGES = 1535
EXPECTED_STATIONS = 61
MONTHLY_STATS = ("complete_daily_months_compared", "within_0_05mm",
                 "outside_0_05mm", "max_absolute_sum_difference_mm")


def patch_station(data, station, proposed):
    """Change only the selected tokens and preserve the file's byte layout."""
    before = station_values(data, station)
    lines = data.splitlines(keepends=True)
    edits = {}
    for index, line in enumerate(lines):
        if not line.strip():
            continue
        body = line.rstrip(b"\r\n")
        ending = line[len(body):]
        cells = body.split(b"\t")
        year, month = int(cells[0].removeprefix(b"\xef\xbb\xbf")), int(cells[1])
        for day in range(1, calendar.monthrange(year, month)[1] + 1):
            ordinal = date(year, month, day).toordinal()
            if ordinal not in proposed:
                continue
            original, current = proposed[ordinal]
            if ordinal not in before or before[ordinal] != current or original != current:
                raise ValueError(f"{station} {iso(ordinal)}: precondition failed")
            edits[ordinal] = cells[day + 1].decode("ascii")
            cells[day + 1] = b"NA"
        lines[index] = b"\t".join(cells) + ending
    patched = b"".join(lines)
    after = station_values(patched, station)
    if edits.keys() != proposed.keys() or after != {
        day: value for day, value in before.items() if day not in proposed
    }:
        raise ValueError(f"{station}: patched series failed verification")
    return patched, after, edits


def table_rows(sheet):
    columns = {cell.value: cell.column for cell in sheet[1] if cell.value is not None}
    for row in range(2, sheet.max_row + 1):
        yield {key: sheet.cell(row, col) for key, col in columns.items()}


def set_cells(row, values):
    for field, value in values.items():
        row[field].value = value


def staged_register(source, destination, updates):
    book = load_workbook(source)
    seen = set()
    try:
        for row in table_rows(book["Stations"]):
            identity = row["dataset_station_id"].value
            if identity in updates:
                seen.add(identity)
                info = updates[identity]
                set_cells(row, {"daily_valid_observations": info["valid_observations"],
                                "daily_first_valid": info["first_valid"],
                                "daily_last_valid": info["last_valid"]})
        if seen != updates.keys():
            raise ValueError("Missing station rows in register")
        seen.clear()
        for row in table_rows(book["Products"]):
            identity = row["dataset_station_id"].value
            if identity not in updates:
                continue
            if row["resolution"].value == "daily" and row["source_network"].value == "IHGAR":
                seen.add(identity)
                set_cells(row, updates[identity])
            if row["resolution"].value == "monthly" and any(row[s].value is not None for s in MONTHLY_STATS):
                set_cells(row, dict.fromkeys(MONTHLY_STATS))
                row["comparison_interpretation"].value = (
                    "Daily QC changed the paired series. Recompute the daily/monthly diagnostic."
                )
        if seen != updates.keys():
            raise ValueError("Missing daily product rows in register")
        book.save(destination)
    finally:
        book.close()
    stations, products = get_register(destination)
    for station in stations:
        key = station["dataset_station_id"]
        if key in updates and any(products[key][k] != v for k, v in updates[key].items()):
            raise ValueError(f"Staged register does not match patched file for {key}")


def commit(dataset, run, manifest):
    """Replace staged files, restoring applied files on a handled failure."""
    for rel, hashes in manifest.items():
        if sha256((dataset / rel).read_bytes()) != hashes["before"]:
            raise ValueError(f"Final file changed during staging: {rel}")
        if sha256((run / "staged" / rel).read_bytes()) != hashes["after"]:
            raise ValueError(f"Staged file changed during staging: {rel}")
        if sha256((run / "backup" / rel).read_bytes()) != hashes["before"]:
            raise ValueError(f"Backup hash mismatch: {rel}")
    applied = []
    try:
        for rel in manifest:
            os.replace(run / "staged" / rel, dataset / rel)
            applied.append(rel)
        for rel, hashes in manifest.items():
            if sha256((dataset / rel).read_bytes()) != hashes["after"]:
                raise ValueError(f"Final verification failed: {rel}")
    except (OSError, ValueError, KeyboardInterrupt):
        failed_restore = []
        for rel in reversed(applied):
            try:
                restore = run / "staged" / rel
                shutil.copyfile(run / "backup" / rel, restore)
                os.replace(restore, dataset / rel)
            except OSError:
                failed_restore.append(rel)
        if failed_restore:
            raise RuntimeError(f"Restore these paths from {run / 'backup'}: {failed_restore}")
        raise


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project", required=True, type=Path)
    parser.add_argument("--original", type=Path, help="Override path to Original IH_GAR_Data.zip")
    parser.add_argument("--checked", type=Path, help="Override path to the verified Checked ZIP")
    parser.add_argument("--apply", action="store_true", help="Write the staged changes to the final dataset")
    args = parser.parse_args()
    root = args.project.resolve()
    dataset = root / DATASET
    if dataset.is_symlink() or dataset.resolve() != dataset or (dataset / "Daily").resolve() != dataset / "Daily":
        raise ValueError("Final dataset and Daily paths must point inside the project without redirection")
    register = dataset / REGISTER
    if register.is_symlink() or register.resolve() != register:
        raise ValueError("Final register must point inside the final dataset without redirection")
    raw = root / "01_Data_downloaded"
    archived = root / "working/ihgar_station_update/offline_v3_20261109_205739_593/inputs/IH_GAR_Data.zip"
    original = unique_input(args.original, archived if archived.is_file() else raw / "01_IHGAR/IH_GAR_Data.zip",
                            (root / "working/ihgar_station_update").rglob("IH_GAR_Data.zip"), "--original")
    checked = unique_input(args.checked, raw / CHECKED_NAME,
                           raw.glob("*IH_GAR_checked*.zip"), "--checked")
    for name, path in (("original", original), ("checked", checked), ("register", register)):
        if sha256(path.read_bytes()) != REFERENCE["input_sha256"][name]:
            raise ValueError(f"{name} SHA-256 differs from the verified audit: {path}")

    audit, _, exclusions, _ = compare(root, original, checked, register)
    if compare_reference(audit)["status"] != "MATCH" or audit["updated_backing"] != {"Assembled Daily": 217}:
        raise ValueError("The full audit must match the reference using final daily files exclusively")
    groups = {}
    for sid, _, day, period, old, check, current, span, decision in exclusions:
        if (period, span, decision) == ("1951–2023", "Within valid span", "Propose removal"):
            if check is not None or old is None or old != current:
                raise ValueError(f"Unexpected exclusion evidence for {sid} {day}")
            groups.setdefault(sid, {})[date.fromisoformat(day).toordinal()] = (old, current)
    if len(groups) != EXPECTED_STATIONS or sum(map(len, groups.values())) != EXPECTED_CHANGES:
        raise ValueError("Unexpected approved batch size")

    parent = root / "working/ihgar_checked_exclusions"
    parent.mkdir(parents=True, exist_ok=True)
    run = Path(tempfile.mkdtemp(prefix=f"{datetime.now():%Y%m%d_%H%M%S}_", dir=parent))
    for sub in ("staged", "backup"):
        (run / sub / "Daily").mkdir(parents=True)
    manifest, updates, log = {}, {}, []
    stations, products = get_register(register)
    for station in stations:
        sid, identity = station["provider_id"], station["dataset_station_id"]
        if sid not in groups:
            continue
        product = products[identity]
        rel = relative_path(product["relative_path"])
        path = dataset / rel
        if len(rel.parts) != 2 or rel.parts[0] != "Daily" or path.resolve() != path:
            raise ValueError(f"Unexpected final path for {sid}: {path}")
        if rel.as_posix() in manifest:
            raise ValueError(f"Final path repeated: {rel}")
        data = path.read_bytes()
        digest = sha256(data)
        if digest != product["sha256"]:
            raise ValueError(f"Final file hash changed: {path}")
        after, values, edits = patch_station(data, sid, groups[sid])
        new_hash = sha256(after)
        (run / "backup" / rel).write_bytes(data)
        (run / "staged" / rel).write_bytes(after)
        manifest[rel.as_posix()] = {"before": digest, "after": new_hash}
        first, last, valid = first_last(values)
        updates[identity] = {
            "sha256": new_hash, "first_valid": first, "last_valid": last,
            "valid_observations": valid, "zero_observations": sum(v == 0 for v in values.values()),
            "missing_explicit_calendar_cells": product["missing_explicit_calendar_cells"] + len(edits),
            "destination_verification": "CHECKED_IHGAR_EXCLUSIONS_VERIFIED",
        }
        log.extend((sid, iso(day), token, "NA", rel.as_posix()) for day, token in sorted(edits.items()))
    if len(log) != EXPECTED_CHANGES or len(updates) != EXPECTED_STATIONS:
        raise ValueError("Staged edits do not match the approved batch")
    reg_before = register.read_bytes()
    if sha256(reg_before) != REFERENCE["input_sha256"]["register"]:
        raise ValueError("Register changed during staging")
    (run / "backup" / REGISTER).write_bytes(reg_before)
    staged_register(run / "backup" / REGISTER, run / "staged" / REGISTER, updates)
    manifest[REGISTER] = {"before": sha256(reg_before),
                          "after": sha256((run / "staged" / REGISTER).read_bytes())}
    write_csv(run / "Change_log.csv", ["IH-GAR station ID", "Date", "Before", "After", "Final file"], log)
    report = run / "Manifest.json"
    summary = {"status": "READY_TO_APPLY" if args.apply else "DRY_RUN", "changes": len(log),
               "stations": len(updates), "dataset": str(dataset), "files": manifest}
    report.write_text(json.dumps(summary, indent=2) + "\n", encoding="utf-8")
    print(f"Change log, manifest and backups: {run}", flush=True)
    if args.apply:
        lock = dataset / ".ihgar_checked_patch_lock"
        try:
            lock.mkdir()
        except FileExistsError as exc:
            raise ValueError(f"Another apply may be in progress. Check the patch lock: {lock}") from exc
        try:
            commit(dataset, run, manifest)
            summary["status"] = "APPLIED_AND_VERIFIED"
            report.write_text(json.dumps(summary, indent=2) + "\n", encoding="utf-8")
        finally:
            lock.rmdir()
    print(f"{summary['status']}: {len(log):,} station-days across {len(updates)} stations")


if __name__ == "__main__":
    try:
        main()
    except (OSError, ValueError, KeyError, RuntimeError) as exc:
        sys.exit(f"Patch stopped: {exc}")
