# Batch Workbook Year Selection Design

## Purpose

Extend the command-line finance file generator so a user can generate files for either one Excel workbook or every Excel workbook in a folder. Each workbook can be run for its own explicit list of financial years, and those years do not need to be consecutive.

## Interaction

The program asks for an input mode:

- `single` processes one workbook path entered by the user.
- `multiple` asks for a folder path, then discovers Excel workbooks directly inside that folder.

In both modes, years are entered as a comma-separated list such as `2024,2026,2028`. Duplicate years are rejected to avoid overwriting the same output set twice in one run.

For multiple mode, the program asks for a separate year list for each discovered workbook. Workbooks are processed in sorted filename order. Supported file types are `.xlsx` and `.xlsm`; temporary Excel lock files beginning with `~$` are ignored.

Directory ID, Advisor ID, and base output directory remain shared for all requested runs.

## Generation Flow

The existing `Config` object and `generate_finance_files(config)` function remain the single-run contract. The new interactive layer converts the user's selected workbooks and year lists into one `Config` per workbook/year pair, then loops through those configs.

This keeps the workbook parsing, filtering, aggregation, template transformation, and output writing behavior unchanged for each individual run.

## Output File Names

Generated files remain under `<base output directory>/<policy number>/<financial year>/`, but the file names are generated from the file type, financial year, record sequence, and current timestamp. The timestamp uses `YYYYMMDDHHMM` and does not include seconds.

- Finance: `RMM1ITCFIN<YEAR>001<YYYYMMDDHHMM>`
- Member: `RMM1ITCMEM<YEAR>002<YYYYMMDDHHMM>`
- Advisor: `RMM1ITCADV<YEAR>003<YYYYMMDDHHMM>`

## Errors

Input validation remains fail-fast. Invalid mode, missing folder, empty folder, invalid year list, duplicate years, invalid identifiers, missing templates, or missing input files stop the run with a clear exception before generation proceeds.

During generation, any workbook/year data error uses the existing validation behavior and stops the batch.

## Verification

Tests cover:

- parsing comma-separated, non-consecutive year lists;
- rejecting duplicate years;
- accepting numbered and text mode choices;
- building multiple configs from single mode;
- discovering `.xlsx` and `.xlsm` files in folder mode while ignoring other files and Excel lock files; and
- generating timestamped finance, member, and advisor file names without seconds; and
- preserving the existing end-to-end generation behavior for the sample workbook.
