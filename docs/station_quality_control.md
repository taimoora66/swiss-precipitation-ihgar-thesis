# Swiss station quality control

The workflow starts with the assembled daily files and `Station_Register.xlsx` in `03_Daily_dataset`. See the [assembly guide](station_assembly.md) for the preceding step. Provider observations, source archives, DEM tiles and review workbooks are local inputs and outputs. Obtain the inputs under their applicable terms.

## Checked IH-GAR reconciliation

The reconciliation compared the original, Checked and assembled IH-GAR daily series by station ID and date. The verified batch set 1,535 station-days to missing across 61 records within 1951–2023. The remaining observations were checked against the pre-patch baseline. Coverage-edge dates and unresolved differences were retained.

[`apply_ihgar_checked_exclusions.py`](../scripts/apply_ihgar_checked_exclusions.py) replays this specific batch against the exact original and Checked archives and the pre-patch register. It verifies the input hashes and defaults to a dry run. The applied run retains backups, a change log and a manifest under `working/ihgar_checked_exclusions`. [`verify_ihgar_checked_alignment.py`](../scripts/verify_ihgar_checked_alignment.py) independently verifies the applied batch against the retained inputs. Use each script's `--help` for paths. An already-patched dataset starts at daily QC.

## Daily and coordinate review

Run commands from the repository checkout. Replace `/path/to/project` with the project containing `03_Daily_dataset`. Install the Python dependencies into your environment.

```sh
python -m pip install numpy openpyxl XlsxWriter rasterio
python scripts/audit_post_patch_daily_qc.py --project "/path/to/project"
python scripts/build_station_dem_review.py --project "/path/to/project"
python scripts/build_combined_station_qc.py --project "/path/to/project"
```

- **Daily QC** reads the post-patch register and its applied manifest. The default daily period is 1951–2023. Review thresholds are precipitation greater than 600 mm and internal missing or recorded-zero runs of at least 90 days. Missing values and recorded zero amounts are distinct. Event comparisons retain the recorded accumulation-window information.
- **DEM review** reads the daily QC workbook and local Copernicus GLO-30 tiles under `04_Analysis_Review/06_DEM/tiles` by default. It compares registered elevations with terrain sampled at the recorded locations.
- **Combined QC** uses the daily and DEM reports and the verified IH-GAR update archive. It reports horizontal separation, reported elevation difference and elevation discrepancies against each location's DEM sample. The default 100 m threshold is an elevation review flag. Source metadata and dated site histories are needed to resolve a potential relocation or coordinate discrepancy.

These reports are written under `04_Analysis_Review`. To reproduce a particular review, pass the exact inputs using the options shown by `--help`.
