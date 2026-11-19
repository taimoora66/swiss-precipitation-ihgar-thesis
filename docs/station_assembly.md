# Reproduce the November 2026 station assembly

This workflow reproduces the November 2026 assembly of a mixed-window station collection. Use Python 3.10 or later, install `openpyxl`, and run the commands from this repository with a separate local project directory for the inputs and outputs.

| Input | Required version or layout |
| --- | --- |
| `Network_station_export.zip` | Audited numbered export v1.1.1, SHA-256 `5c868b12987990b59fb6959f135de9b05070decf413bbbf00ebfd4fbac22314c`. |
| `IH_GAR_Data.zip` | Original 217-entry IH-GAR archive, SHA-256 `8521e488b1361cfb89235fa47894c2a6d3f83ee2d0748bee865a681e4657733e`. |
| IH-GAR MeteoSwiss update release | The reviewed release with an `input_hashes.json` that matches both archives above and the accompanying update ledger. |
| Prepared local inputs | Provider metadata under `01_Data_downloaded` and the reviewed second-stage products under `02_Data_extraction`, including WSL daily products for both UTC windows. Obtain these separately under the applicable data terms. |

Set the local input paths in PowerShell. Keep the prepared project directory separate from the Git checkout.

```powershell
$project = 'C:/path/to/local/project'
$export = 'C:/path/to/Network_station_export.zip'
$original = 'C:/path/to/IH_GAR_Data.zip'
$release = 'C:/path/to/IHGAR_MeteoSwiss_Station_Update.zip'
py -3 -m pip install openpyxl
```

1. [Audit candidates](../scripts/audit_dataset_candidates.py). Check `working/dataset_assembly_audit/Dataset_Register_Audit.xlsx` and `Dataset_Assembly_Summary.json` for the proposed selection and review items. The reference run verified 1,067 source products.

   ```powershell
   py -3 scripts/audit_dataset_candidates.py --root $project --export $export --ihgar-original $original --ihgar-release $release
   ```

2. [Assemble the reviewed collection](../scripts/assemble_station_dataset.py). Use a local project directory with no conflicting files under `03_Daily_dataset`. Run the check first and proceed only after the audit has zero blocking issues and the check succeeds. The reference run produced a station register, 435 daily files and 343 monthly files.

   ```powershell
   py -3 scripts/assemble_station_dataset.py --root $project --check-only
   if ($LASTEXITCODE -ne 0) { throw 'Assembly check failed.' }
   py -3 scripts/assemble_station_dataset.py --root $project
   ```

3. [Switch the WSL daily products to 00 UTC](../scripts/switch_wsl_dataset_to_00utc.py). Confirm the check succeeds before applying the switch. The reference run replaced 13 daily files and found zero differences above 0.05 mm across 2,732 complete WSL station-month comparisons. The script backs up affected files and writes `working/dataset_assembly/WSL_00UTC_Verification.json`.

   ```powershell
   py -3 scripts/switch_wsl_dataset_to_00utc.py --root $project --export $export --check-only
   if ($LASTEXITCODE -ne 0) { throw 'WSL check failed.' }
   py -3 scripts/switch_wsl_dataset_to_00utc.py --root $project --export $export
   ```

These counts describe the pinned project release. The WSL comparison checks arithmetic consistency between products that share an hourly source. The assembled collection retains different accumulation windows, and the verification report records `common_window_analysis_ready = false`.

Compare the resulting register and verification reports with the [project summary](../README.md#november-2026-station-collection).
