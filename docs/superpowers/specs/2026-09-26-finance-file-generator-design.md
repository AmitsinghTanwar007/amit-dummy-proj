# Finance File Generator Design

## Purpose

Build a Python program that reads a source Excel workbook, filters and aggregates qualifying transactions for a specified financial year, applies the calculated values to the finance, member, and advisor templates, and writes three generated files beneath a policy/year directory structure.

## Configuration

When the module is executed, it interactively prompts for:

- input Excel path;
- financial year;
- 15-digit Directory ID;
- five-digit Advisor ID; and
- base output directory.

Relative paths are resolved from the current working directory, and absolute paths are accepted unchanged. The three template paths remain fixed relative to the Python module. Input values are converted into the existing immutable `Config` object, so the transformation functions remain independent of the interactive interface.

`default_config()` remains available for automated tests and programmatic use. `interactive_config()` owns prompting and input conversion. Invalid input produces a clear error and no output files are generated.

## Required source columns

The program reads these columns by exact header name and ignores every other column:

- `Policy Number`
- `System Date`
- `Transaction Code`
- `Reversal File Code`
- `Transaction Amount`

## Filtering

Rows are filtered only in memory. The source workbook remains unchanged.

A row qualifies when both conditions are true:

1. `Transaction Code` is the text `NA` after trimming surrounding whitespace.
2. `Reversal File Code` is blank or the exact text `00000` after trimming surrounding whitespace.

A numeric reversal value of `0` does not qualify.

For financial year `Y`, the applicable date range is March 1 of `Y - 1` through the final day of February in `Y`, inclusive. Therefore, financial year `2026` covers `2025-03-01` through `2026-02-28`.

## Validation

The program stops with a clear error when:

- a required column is missing;
- the configured financial year is invalid;
- the Directory ID is not exactly 15 digits;
- the Advisor ID is not exactly five digits;
- a retained policy number is not exactly 15 digits;
- retained records contain more than one policy number;
- a date or amount cannot be parsed; or
- no rows remain after applying the code, reversal, and financial-year filters.

The workbook's `System Date` values are parsed using the strict `YYYY-MM-DD` pattern. The source workbook created for this project stores them as text in that format.

## Aggregation

The program groups retained transactions into the twelve finance months in this order:

`MARCH`, `APRIL`, `MAY`, `JUNE`, `JULY`, `AUGUST`, `SEPTEMBER`, `OCTOBER`, `NOVEMBER`, `DECEMBER`, `JANUARY`, `FEBRUARY`.

Multiple qualifying amounts in one month are summed. A month without a qualifying transaction has an amount of zero. Decimal arithmetic is used for financial calculations. The annual current amount is the sum of all twelve monthly values.

## Template transformation

The program reads `templates/financefile.txt`. Because the supplied file has physical line wrapping inside its header and detail records, the parser reconstructs the three logical records: header, detail, and footer.

It locates detail fields using the reconstructed header rather than assuming fixed indexes, then updates:

- `TAX YEAR` with the configured financial year;
- `CLIENT ID` with the configured Directory ID;
- `POLICY ID` with the single policy number from the retained Excel rows;
- `MARCH` through `FEBRUARY` with the monthly totals;
- `CURRENT AMOUNT/INCOME PROTECTION AMOUNT` with the annual total; and
- `ARREARS AMOUNT` with zero.

The footer contains one detail-record count, the twelve monthly totals, the annual total, and a final zero arrears amount. Since this version permits one constant policy only, the record count remains one.

All output amounts use the fixed-width format demonstrated by the template, such as `0000000000000336.17` and `0000000000000000.00`.

## Output

The program generates three files beneath the same policy/year directory:

- `<base output directory>/<policy number>/<financial year>/financial.txt`
- `<base output directory>/<policy number>/<financial year>/member.txt`
- `<base output directory>/<policy number>/<financial year>/advisor.txt`

The program creates missing policy and financial-year directories. It does not alter the source Excel workbook or source template.

## Member file

The member generator reads `templates/memberfile.txt`, reconstructs its wrapped header and detail records, and locates fields by normalized header name. It updates:

- `TAX YEAR` with the configured financial year;
- `CLIENT ID` with the configured 15-digit Directory ID;
- `POLICY ID` with the constant 15-digit policy number from the filtered Excel data;
- `POLICY ON DATE` with the earliest qualifying System Date in `YYYY/MM/DD` template format; and
- `POLICY OFF DATE` with a blank value.

The earliest date is selected after applying the transaction-code, reversal-code, and financial-year filters and sorting qualifying transactions by System Date. The member footer retains its template record type `0002` and a detail-record count of one.

## Advisor file

The advisor generator reads `templates/advisorfile.txt`, reconstructs its wrapped header and detail records, and locates fields by normalized header name. It updates:

- `TAX YEAR` with the configured financial year;
- `CLIENT ID` with the configured Directory ID;
- `POLICY ID` with the policy number from Excel; and
- `ADVISOR ID` with the configured value `87689`.

Advisor ID must contain exactly five digits. Other advisor template values remain unchanged. The advisor footer retains its template record type `0003` and a detail-record count of one.

## Functional structure

The implementation follows a functional style with small, single-purpose functions. Separate functions handle configuration validation, workbook loading, column selection, code/reversal filtering, strict date parsing, financial-year filtering, policy validation, monthly aggregation, amount formatting, logical template parsing, detail transformation, footer transformation, output-path construction, and file writing.

Transformation functions accept inputs and return new values rather than modifying shared global state. A short orchestration function connects the shared filtering pipeline to the three template transformers and reports the generated paths and calculated totals.

## Verification

Automated checks will run the generator against `sample_data/Finance_Dummy_Input.xlsx` for financial year `2026`. They verify all three output paths, record structures, updated identifiers, earliest on-date, blank off-date, footer types, and finance monthly and annual totals.

Tests also mock the five interactive responses and verify that relative paths, year, Directory ID, Advisor ID, and output directory are mapped to the expected configuration fields.
