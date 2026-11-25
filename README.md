# Swiss station precipitation for IH-GAR

This MSc thesis at the University of Milan examines whether additional Swiss gauge records improve daily precipitation estimates for Alpine catchments draining into northern Italy. The reference is the 1951–2023 reconstruction for the Italian Hydrological part of the Greater Alpine Region (IH-GAR), distributed at 30 arc-second resolution. The thesis is supervised by Prof. Maurizio Maugeri and Dr. Veronica Manara. [Manara et al. (2026)](https://doi.org/10.1016/j.ejrh.2026.103776)

The current result is a Swiss station collection assembled from the original IH-GAR archive, MeteoSwiss, Agrometeo, the SUPSI Ticino hydrological yearbooks, and the WSL Long-term Forest Ecosystem Research network (LWF). The collection retains the time windows under which precipitation was accumulated. Its verification report records that it is still pending screening for common-window analysis.

## November 2026 station collection

The audit verified 1,067 source products before transfer. The assembled collection contains 435 daily files and 343 monthly files. Counts refer to register records and their files. One physical gauge can appear in more than one source archive, and a station name can span changes in location or instrumentation.

| Register source | Daily files | Monthly files |
| --- | ---: | ---: |
| IH-GAR archive and reviewed MeteoSwiss updates | 217 | 0 |
| Additional MeteoSwiss records | 83 | 208 |
| Agrometeo | 82 | 82 |
| SUPSI | 40 | 40 |
| WSL LWF | 13 | 13 |
| **Total** | **435** | **343** |

The register kept the original coordinates and elevations for all 217 IH-GAR entries. The reviewed archive release changed 124 entries, of which 123 received additions and one received corrections only. Where a MeteoSwiss record had a confirmed match in that release, it was linked to the corresponding IH-GAR entry. Ambiguous links were retained for review. The MeteoSwiss monthly inventory has different coverage from the selected daily inventory.

## What the checks established

- **WSL accumulation window.** After assembly, 13 WSL daily files were switched to the 00 UTC product. Their daily sums agreed with the corresponding monthly products within 0.05 mm in all 2,732 complete station-month comparisons. Both products draw on the same hourly source, so this check establishes arithmetic consistency.
- **SUPSI month eligibility.** The processing replay covered 40 register entries. It applied the provider’s month-level eligibility and dash rules to saved transcriptions, retained missing days, and flagged 163 monthly discrepancies for review. The replay verified how the saved transcription was processed. The [SUPSI extraction guide](docs/supsi_extraction.md) describes the earlier transcription stage.
- **Daily and monthly agreement across networks.** The read-only review found 5,696 differences above 0.05 mm among 209,170 complete comparisons with a monthly reference. Of those, 613 exceeded 10 mm in absolute value. Comparisons with incomplete daily coverage were recorded separately. The later review approved zero corrections. These figures are counts of comparisons, and a flag can arise from different accumulation windows as well as an observation issue.

## Checked IH-GAR reconciliation and station quality control

The verified reconciliation with Checked IH-GAR set 1,535 approved station-days to missing across 61 records. The remaining values were checked against the pre-patch series. The subsequent daily review screened precipitation extremes, missing runs and recorded zero runs. The coordinate review compared registered positions and elevations with provider metadata and Copernicus GLO-30 terrain samples. Coordinate and elevation discrepancies are flagged for station-history review. [Station quality control workflow](docs/station_quality_control.md)

## Code and inputs

| Task | Entry point |
| --- | --- |
| Retrieve MeteoSwiss station metadata and daily and monthly products | [MeteoSwiss acquisition script](scripts/meteoswiss_download.py) |
| Archive Agrometeo source responses and organize precipitation records | [Acquisition](docs/agrometeo_acquisition.md) and [organization](docs/agrometeo_organization.md) |
| Transcribe the SUPSI yearbooks and track coordinates and elevations printed in successive editions | [Extraction](docs/supsi_extraction.md), [organization](docs/supsi_station_organization.md), and [site history](docs/supsi_site_history.md) |
| Archive and screen WSL LWF hourly files | [Acquisition](docs/wsl_lwf_acquisition.md), [screening](docs/wsl_lwf_screening.md), and [API workbook](docs/wsl_lwf_api_only_excel.md) |
| Audit the prepared products, assemble the reference collection, and apply the WSL window selection | [Station assembly instructions](docs/station_assembly.md) |
| Verify the Checked IH-GAR reconciliation and review daily and coordinate flags | [Station quality control workflow](docs/station_quality_control.md) |

The repository distributes code and documentation. Obtain the observations and archive packages under their applicable terms. The assembly instructions identify the local inputs, checksums and verification files needed to repeat the reference run. Attribute MeteoSwiss observations according to the [provider’s open-data documentation](https://opendatadocs.meteoswiss.ch/).

## From station records to a gridded comparison

The next scientific step is to resolve station identity and accumulation-window flags, then check record homogeneity where station history and overlapping neighbours provide sufficient evidence. Catchment boundaries define the target area. Stations outside those boundaries, including parts of Bern and Vaud, may still support interpolation.

The thesis will compare baseline and revised precipitation grids under the same reconstruction settings, assessing daily amounts and extremes at withheld stations. The Ticino system is the focus of the planned hydrological comparison, subject to suitable discharge, evapotranspiration, storage and transfer records. Whether the revised inputs improve the grid or the catchment water balance remains an open research question.
