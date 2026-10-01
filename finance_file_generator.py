"""Generate a finance text file from an Excel transaction workbook."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Callable, Iterable, Mapping, Sequence

from openpyxl import load_workbook


BASE_DIRECTORY = Path(__file__).resolve().parent

# First-version configuration. These values can become command-line arguments
# later without changing the transformation functions below.
INPUT_EXCEL = BASE_DIRECTORY / "sample_data" / "Finance_Dummy_Input.xlsx"
FINANCIAL_YEAR = 2026
DIRECTORY_ID = "000000023353484"
TEMPLATE_FILE = BASE_DIRECTORY / "templates" / "financefile.txt"
MEMBER_TEMPLATE_FILE = BASE_DIRECTORY / "templates" / "memberfile.txt"
ADVISOR_TEMPLATE_FILE = BASE_DIRECTORY / "templates" / "advisorfile.txt"
ADVISOR_ID = "87689"
OUTPUT_BASE_DIRECTORY = BASE_DIRECTORY / "output"

REQUIRED_COLUMNS = (
    "Policy Number",
    "System Date",
    "Transaction Code",
    "Reversal File Code",
    "Transaction Amount",
)

MONTH_NAMES = (
    "MARCH",
    "APRIL",
    "MAY",
    "JUNE",
    "JULY",
    "AUGUST",
    "SEPTEMBER",
    "OCTOBER",
    "NOVEMBER",
    "DECEMBER",
    "JANUARY",
    "FEBRUARY",
)

ANNUAL_TOTAL_COLUMN = "CURRENT AMOUNT/INCOME PROTECTION AMOUNT"
ARREARS_COLUMN = "ARREARS AMOUNT"
ZERO = Decimal("0.00")
MAX_FINANCE_AMOUNT = Decimal("9999999999999999.99")
EXCEL_FILE_SUFFIXES = (".xlsx", ".xlsm")
FINANCE_FILE_PREFIX = "FIN"
MEMBER_FILE_PREFIX = "MEM"
ADVISOR_FILE_PREFIX = "ADV"
FINANCE_FILE_SEQUENCE = "001"
MEMBER_FILE_SEQUENCE = "002"
ADVISOR_FILE_SEQUENCE = "003"


@dataclass(frozen=True)
class Config:
    input_excel: Path
    financial_year: int
    directory_id: str
    template_file: Path
    member_template_file: Path
    advisor_template_file: Path
    advisor_id: str
    output_base_directory: Path


@dataclass(frozen=True)
class SourceRow:
    policy_number: object
    system_date: object
    transaction_code: object
    reversal_file_code: object
    transaction_amount: object


@dataclass(frozen=True)
class Transaction:
    policy_number: str
    system_date: date
    amount: Decimal


@dataclass(frozen=True)
class FinanceSummary:
    policy_number: str
    monthly_amounts: tuple[Decimal, ...]
    annual_total: Decimal
    earliest_system_date: date


@dataclass(frozen=True)
class GeneratedFiles:
    financial: Path
    member: Path
    advisor: Path


def default_config() -> Config:
    """Return the hardcoded configuration used by this first version."""

    return Config(
        input_excel=INPUT_EXCEL,
        financial_year=FINANCIAL_YEAR,
        directory_id=DIRECTORY_ID,
        template_file=TEMPLATE_FILE,
        member_template_file=MEMBER_TEMPLATE_FILE,
        advisor_template_file=ADVISOR_TEMPLATE_FILE,
        advisor_id=ADVISOR_ID,
        output_base_directory=OUTPUT_BASE_DIRECTORY,
    )


def require_input(value: str, field_name: str) -> str:
    """Return trimmed interactive input or reject an empty response."""

    normalized = value.strip()
    if not normalized:
        raise ValueError(f"{field_name} is required.")
    return normalized


def resolve_input_path(value: str, field_name: str) -> Path:
    """Resolve a required relative or absolute path entered by the user."""

    return Path(require_input(value, field_name)).expanduser().resolve()


def parse_financial_year_input(value: str) -> int:
    """Convert an interactive financial-year response to an integer."""

    normalized = require_input(value, "Financial year")
    try:
        financial_year = int(normalized)
    except ValueError as error:
        raise ValueError("Financial year must be a four-digit integer.") from error
    if financial_year < 1000 or financial_year > 9999:
        raise ValueError("Financial year must be a four-digit integer.")
    return financial_year


def parse_financial_years_input(value: str) -> tuple[int, ...]:
    """Convert a comma-separated financial-year response to integers."""

    normalized = require_input(value, "Financial years")
    parts = [part.strip() for part in normalized.split(",")]
    if any(not part for part in parts):
        raise ValueError("Financial years must be comma-separated four-digit integers.")
    years = tuple(parse_financial_year_input(part) for part in parts)
    if len(set(years)) != len(years):
        raise ValueError("Financial years must not contain duplicates.")
    return years


def parse_run_mode_input(value: str) -> str:
    """Return the selected workbook input mode."""

    normalized = require_input(value, "Input mode").strip().lower()
    if normalized in ("1", "single", "s"):
        return "single"
    if normalized in ("2", "multiple", "m"):
        return "multiple"
    raise ValueError("Input mode must be single or multiple.")


def discover_excel_files(folder_path: Path) -> tuple[Path, ...]:
    """Return supported Excel workbooks directly inside a folder."""

    if not folder_path.is_dir():
        raise NotADirectoryError(f"Excel folder not found: {folder_path}")
    excel_files = tuple(
        sorted(
            path
            for path in folder_path.iterdir()
            if path.is_file()
            and path.suffix.lower() in EXCEL_FILE_SUFFIXES
            and not path.name.startswith("~$")
        )
    )
    if not excel_files:
        raise FileNotFoundError(f"No Excel files found in folder: {folder_path}")
    return excel_files


def build_config(
    input_excel: Path,
    financial_year: int,
    directory_id: str,
    advisor_id: str,
    output_directory: Path,
) -> Config:
    """Create and validate one generator configuration."""

    return validate_config(
        Config(
            input_excel=input_excel,
            financial_year=financial_year,
            directory_id=directory_id,
            template_file=TEMPLATE_FILE,
            member_template_file=MEMBER_TEMPLATE_FILE,
            advisor_template_file=ADVISOR_TEMPLATE_FILE,
            advisor_id=advisor_id,
            output_base_directory=output_directory,
        )
    )


def build_configs_for_workbook_years(
    workbook_years: Sequence[tuple[Path, Sequence[int]]],
    directory_id: str,
    advisor_id: str,
    output_directory: Path,
) -> tuple[Config, ...]:
    """Build one configuration for each workbook and requested year."""

    return tuple(
        build_config(workbook, year, directory_id, advisor_id, output_directory)
        for workbook, years in workbook_years
        for year in years
    )


def interactive_config(input_function: Callable[[str], str] = input) -> Config:
    """Prompt for runtime values and return a validated immutable configuration."""

    input_excel = resolve_input_path(
        input_function("Excel file path: "),
        "Excel file path",
    )
    financial_year = parse_financial_year_input(
        input_function("Financial year (for example, 2026): ")
    )
    directory_id = require_input(
        input_function("15-digit Directory/Client ID: "),
        "Directory ID",
    )
    advisor_id = require_input(
        input_function("5-digit Advisor ID: "),
        "Advisor ID",
    )
    output_directory = resolve_input_path(
        input_function("Base output directory: "),
        "Base output directory",
    )
    return validate_config(
        Config(
            input_excel=input_excel,
            financial_year=financial_year,
            directory_id=directory_id,
            template_file=TEMPLATE_FILE,
            member_template_file=MEMBER_TEMPLATE_FILE,
            advisor_template_file=ADVISOR_TEMPLATE_FILE,
            advisor_id=advisor_id,
            output_base_directory=output_directory,
        )
    )


def interactive_configs(input_function: Callable[[str], str] = input) -> tuple[Config, ...]:
    """Prompt for one or more workbooks and return validated run configurations."""

    mode = parse_run_mode_input(input_function("Mode (single/multiple): "))
    if mode == "single":
        input_excel = resolve_input_path(
            input_function("Excel file path: "),
            "Excel file path",
        )
        years = parse_financial_years_input(
            input_function("Financial years (comma-separated, for example 2024,2026): ")
        )
        workbook_years = ((input_excel, years),)
    else:
        excel_folder = resolve_input_path(
            input_function("Excel folder path: "),
            "Excel folder path",
        )
        workbook_years = tuple(
            (
                excel_file,
                parse_financial_years_input(
                    input_function(
                        f"Financial years for {excel_file.name} "
                        "(comma-separated, for example 2024,2026): "
                    )
                ),
            )
            for excel_file in discover_excel_files(excel_folder)
        )

    directory_id = require_input(
        input_function("15-digit Directory/Client ID: "),
        "Directory ID",
    )
    advisor_id = require_input(
        input_function("5-digit Advisor ID: "),
        "Advisor ID",
    )
    output_directory = resolve_input_path(
        input_function("Base output directory: "),
        "Base output directory",
    )
    return build_configs_for_workbook_years(
        workbook_years,
        directory_id,
        advisor_id,
        output_directory,
    )


def validate_fifteen_digit_id(value: str, field_name: str) -> str:
    """Return a valid 15-digit identifier or raise a descriptive error."""

    normalized = value.strip()
    if len(normalized) != 15 or not normalized.isdigit():
        raise ValueError(f"{field_name} must contain exactly 15 digits.")
    return normalized


def validate_advisor_id(value: str) -> str:
    """Return a valid five-digit Advisor ID or raise a descriptive error."""

    normalized = value.strip()
    if len(normalized) != 5 or not normalized.isdigit():
        raise ValueError("Advisor ID must contain exactly five digits.")
    return normalized


def validate_config(config: Config) -> Config:
    """Validate configuration values and input file availability."""

    if isinstance(config.financial_year, bool) or not isinstance(config.financial_year, int):
        raise ValueError("Financial year must be a four-digit integer.")
    if config.financial_year < 1000 or config.financial_year > 9999:
        raise ValueError("Financial year must be a four-digit integer.")
    validate_fifteen_digit_id(config.directory_id, "Directory ID")
    validate_advisor_id(config.advisor_id)
    if not config.input_excel.is_file():
        raise FileNotFoundError(f"Input Excel file not found: {config.input_excel}")
    templates = {
        "Finance": config.template_file,
        "Member": config.member_template_file,
        "Advisor": config.advisor_template_file,
    }
    for template_name, template_path in templates.items():
        if not template_path.is_file():
            raise FileNotFoundError(f"{template_name} template not found: {template_path}")
    return config


def build_header_map(headers: Sequence[object]) -> dict[str, int]:
    """Map required Excel header names to their zero-based column positions."""

    header_map: dict[str, int] = {}
    for index, value in enumerate(headers):
        if isinstance(value, str):
            name = value.strip()
            if name in REQUIRED_COLUMNS:
                if name in header_map:
                    raise ValueError(f"Duplicate required Excel column: {name}")
                header_map[name] = index

    missing = [name for name in REQUIRED_COLUMNS if name not in header_map]
    if missing:
        raise ValueError("Missing required Excel columns: " + ", ".join(missing))
    return header_map


def source_row_from_values(values: Sequence[object], positions: Mapping[str, int]) -> SourceRow:
    """Select the five required values from one Excel row."""

    return SourceRow(
        policy_number=values[positions["Policy Number"]],
        system_date=values[positions["System Date"]],
        transaction_code=values[positions["Transaction Code"]],
        reversal_file_code=values[positions["Reversal File Code"]],
        transaction_amount=values[positions["Transaction Amount"]],
    )


def read_source_rows(excel_path: Path) -> tuple[SourceRow, ...]:
    """Read required values from the active worksheet without changing it."""

    workbook = load_workbook(excel_path, read_only=True, data_only=True)
    try:
        worksheet = workbook.active
        values = worksheet.iter_rows(values_only=True)
        headers = next(values, None)
        if headers is None:
            raise ValueError("The input Excel worksheet is empty.")
        positions = build_header_map(headers)
        return tuple(source_row_from_values(row, positions) for row in values)
    finally:
        workbook.close()


def has_qualifying_transaction_code(value: object) -> bool:
    """Return whether a Transaction Code is the required text NA."""

    return isinstance(value, str) and value.strip() == "NA"


def has_qualifying_reversal_code(value: object) -> bool:
    """Return whether a Reversal File Code is blank or exact text 00000."""

    return value is None or (isinstance(value, str) and value.strip() in ("", "00000"))


def filter_by_codes(rows: Iterable[SourceRow]) -> tuple[SourceRow, ...]:
    """Keep rows satisfying both transaction and reversal code rules."""

    return tuple(
        row
        for row in rows
        if has_qualifying_transaction_code(row.transaction_code)
        and has_qualifying_reversal_code(row.reversal_file_code)
    )


def parse_system_date(value: object, excel_row_number: int) -> date:
    """Parse one System Date using the required YYYY-MM-DD representation."""

    if isinstance(value, datetime):
        value = value.date()
    if isinstance(value, date):
        return value
    if not isinstance(value, str):
        raise ValueError(
            f"Invalid System Date at Excel row {excel_row_number}: expected YYYY-MM-DD."
        )
    try:
        return datetime.strptime(value.strip(), "%Y-%m-%d").date()
    except ValueError as error:
        raise ValueError(
            f"Invalid System Date at Excel row {excel_row_number}: {value!r}; "
            "expected YYYY-MM-DD."
        ) from error


def normalize_policy_number(value: object, excel_row_number: int) -> str:
    """Normalize and validate one policy number as exactly 15 digits."""

    if isinstance(value, bool):
        normalized = str(value)
    elif isinstance(value, int):
        normalized = str(value)
    elif isinstance(value, float) and value.is_integer():
        normalized = str(int(value))
    else:
        normalized = str(value).strip() if value is not None else ""
    return validate_fifteen_digit_id(normalized, f"Policy Number at Excel row {excel_row_number}")


def parse_amount(value: object, excel_row_number: int) -> Decimal:
    """Convert one transaction amount to Decimal."""

    if value is None or isinstance(value, bool):
        raise ValueError(f"Invalid Transaction Amount at Excel row {excel_row_number}: {value!r}")
    try:
        amount = Decimal(str(value).strip())
    except (InvalidOperation, ValueError) as error:
        raise ValueError(
            f"Invalid Transaction Amount at Excel row {excel_row_number}: {value!r}"
        ) from error
    if not amount.is_finite():
        raise ValueError(f"Invalid Transaction Amount at Excel row {excel_row_number}: {value!r}")
    return amount


def financial_year_bounds(financial_year: int) -> tuple[date, date]:
    """Return the inclusive March-through-February bounds for a financial year."""

    return date(financial_year - 1, 3, 1), date(financial_year, 2, 29 if is_leap_year(financial_year) else 28)


def is_leap_year(year: int) -> bool:
    """Return whether a calendar year is a leap year."""

    return year % 4 == 0 and (year % 100 != 0 or year % 400 == 0)


def filter_and_parse_financial_year(
    rows: Sequence[SourceRow], financial_year: int
) -> tuple[Transaction, ...]:
    """Parse qualifying rows and keep those inside the requested financial year."""

    start_date, end_date = financial_year_bounds(financial_year)
    transactions: list[Transaction] = []
    for offset, row in enumerate(rows, start=2):
        system_date = parse_system_date(row.system_date, offset)
        if start_date <= system_date <= end_date:
            transactions.append(
                Transaction(
                    policy_number=normalize_policy_number(row.policy_number, offset),
                    system_date=system_date,
                    amount=parse_amount(row.transaction_amount, offset),
                )
            )
    return tuple(transactions)


def require_transactions(transactions: Sequence[Transaction]) -> tuple[Transaction, ...]:
    """Reject an empty post-filter transaction collection."""

    if not transactions:
        raise ValueError("No qualifying rows remain for the requested financial year.")
    return tuple(transactions)


def require_single_policy(transactions: Sequence[Transaction]) -> str:
    """Return the sole policy number or reject mixed-policy input."""

    policies = sorted({transaction.policy_number for transaction in transactions})
    if len(policies) != 1:
        raise ValueError(
            "Expected one constant Policy Number after filtering; found: " + ", ".join(policies)
        )
    return policies[0]


def finance_month_index(system_date: date) -> int:
    """Map March..February to indexes 0..11."""

    return system_date.month - 3 if system_date.month >= 3 else system_date.month + 9


def aggregate_transactions(transactions: Sequence[Transaction]) -> FinanceSummary:
    """Sum transactions by finance month and calculate the annual total."""

    policy_number = require_single_policy(transactions)
    monthly = [ZERO for _ in MONTH_NAMES]
    for transaction in transactions:
        index = finance_month_index(transaction.system_date)
        monthly[index] += transaction.amount
    monthly_amounts = tuple(monthly)
    return FinanceSummary(
        policy_number=policy_number,
        monthly_amounts=monthly_amounts,
        annual_total=sum(monthly_amounts, ZERO),
        earliest_system_date=min(transaction.system_date for transaction in transactions),
    )


def format_amount(amount: Decimal) -> str:
    """Format a value as 16 integer digits, a decimal point, and two decimals."""

    rounded = amount.quantize(Decimal("0.01"))
    if abs(rounded) > MAX_FINANCE_AMOUNT:
        raise ValueError(f"Amount is too large for the finance format: {amount}")
    formatted = f"{rounded:019.2f}"
    if len(formatted) != 19:
        raise ValueError(f"Amount does not fit the finance format: {amount}")
    return formatted


def reconstruct_template_records(template_text: str) -> tuple[str, str, str]:
    """Join wrapped physical lines into header, detail, and footer records."""

    records: dict[str, list[str]] = {"HEADER": [], "DETAIL": [], "FOOTER": []}
    current = "HEADER"
    for raw_line in template_text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        upper = line.upper()
        if upper.startswith("DETAIL|"):
            current = "DETAIL"
        elif upper.startswith("FOOTER|"):
            current = "FOOTER"
        records[current].append(line)

    joined = tuple("".join(records[name]) for name in ("HEADER", "DETAIL", "FOOTER"))
    if any(not record for record in joined):
        raise ValueError("Template must contain one header, DETAIL record, and FOOTER record.")
    header, detail, footer = joined
    return repair_wrapped_header(header), detail, footer


def repair_wrapped_header(header_record: str) -> str:
    """Restore spaces lost where the supplied header wraps between words."""

    repairs = {
        "FUND ENTITYCODE": "FUND ENTITY CODE",
        "TAX SOURCECODE": "TAX CODE",
        "TAXSOURCECODE": "TAX CODE",
        "TAXCODE": "TAX CODE",
        "NATURE OFPERSON": "NATURE OF PERSON",
        "PASSPORT COUNTRY OFISSUE": "PASSPORT COUNTRY OF ISSUE",
        "PHYSICAL ADDRESS UNITNUMBER": "PHYSICAL ADDRESS UNIT NUMBER",
        "PHYSICAL ADDRESS STREETNAME": "PHYSICAL ADDRESS STREET NAME",
        "POSTALSAME AS RESIDENTIAL": "POSTAL SAME AS RESIDENTIAL",
        "POSTALADDRESS LINE4": "POSTAL ADDRESS LINE4",
    }
    repaired = header_record
    for broken, correct in repairs.items():
        repaired = repaired.replace(broken, correct)
    return repaired


def split_record(record: str) -> list[str]:
    """Split one pipe-delimited record, preserving a meaningful final blank."""

    return record.split("|")


def align_detail_fields(
    header_fields: Sequence[str], detail_record: str, template_name: str
) -> list[str]:
    """Align detail fields, removing only a proven extra trailing delimiter."""

    detail_fields = split_record(detail_record)
    if len(detail_fields) == len(header_fields) + 1 and detail_fields[-1] == "":
        detail_fields = detail_fields[:-1]
    if len(header_fields) != len(detail_fields):
        raise ValueError(
            f"{template_name} template header and DETAIL field counts do not match."
        )
    return detail_fields


def align_short_advisor_detail(
    header_fields: Sequence[str], detail_record: str
) -> list[str]:
    """Expand omitted advisor blanks and keep the postal code correctly aligned."""

    detail_fields = split_record(detail_record)
    if len(detail_fields) == len(header_fields) + 1 and detail_fields[-1] == "":
        detail_fields = detail_fields[:-1]
    if len(detail_fields) == len(header_fields):
        return detail_fields

    positions = build_required_positions(
        header_fields,
        ("ADVISOR ID", "POSTAL ADDRESS POST CODE", "LICENCE NUMBER"),
        "Advisor",
    )
    advisor_position = positions["ADVISOR ID"]
    if len(detail_fields) <= advisor_position:
        raise ValueError("Advisor template DETAIL ends before the ADVISOR ID field.")

    populated_suffix = [
        str(value).strip()
        for value in detail_fields[advisor_position + 1 :]
        if str(value).strip()
    ]
    if populated_suffix != ["703"]:
        raise ValueError(
            "Short Advisor DETAIL may omit blank fields only and must contain "
            "703 as its sole populated value after ADVISOR ID."
        )

    aligned = [""] * len(header_fields)
    aligned[: advisor_position + 1] = detail_fields[: advisor_position + 1]
    aligned[positions["POSTAL ADDRESS POST CODE"]] = "703"
    aligned[positions["LICENCE NUMBER"]] = ""
    return aligned


def build_template_positions(header_fields: Sequence[str]) -> dict[str, int]:
    """Map normalized finance header names to their positions."""

    positions: dict[str, int] = {}
    for index, value in enumerate(header_fields):
        name = value.strip().upper()
        if name in positions:
            raise ValueError(f"Duplicate finance template column: {name}")
        positions[name] = index

    required = ("TAX YEAR", "CLIENT ID", "POLICY ID", *MONTH_NAMES, ANNUAL_TOTAL_COLUMN, ARREARS_COLUMN)
    missing = [name for name in required if name not in positions]
    if missing:
        raise ValueError("Missing finance template columns: " + ", ".join(missing))
    return positions


def build_required_positions(
    header_fields: Sequence[str], required_fields: Sequence[str], template_name: str
) -> dict[str, int]:
    """Map and validate fields required by a non-finance template."""

    positions = {value.strip().upper(): index for index, value in enumerate(header_fields)}
    missing = [field for field in required_fields if field not in positions]
    if missing:
        raise ValueError(
            f"Missing {template_name} template columns: " + ", ".join(missing)
        )
    return positions


def update_detail_record(
    header_record: str,
    detail_record: str,
    summary: FinanceSummary,
    financial_year: int,
    directory_id: str,
) -> str:
    """Return a new DETAIL record containing identifiers and calculated amounts."""

    header_fields = split_record(header_record)
    detail_fields = align_detail_fields(header_fields, detail_record, "Finance")
    positions = build_template_positions(header_fields)
    updated = list(detail_fields)
    updated[positions["TAX YEAR"]] = str(financial_year)
    updated[positions["CLIENT ID"]] = directory_id
    updated[positions["POLICY ID"]] = summary.policy_number
    for month_name, amount in zip(MONTH_NAMES, summary.monthly_amounts):
        updated[positions[month_name]] = format_amount(amount)
    updated[positions[ANNUAL_TOTAL_COLUMN]] = format_amount(summary.annual_total)
    updated[positions[ARREARS_COLUMN]] = format_amount(ZERO)
    return "|".join(updated)


def update_footer_record(footer_record: str, summary: FinanceSummary) -> str:
    """Return a one-record FOOTER containing monthly and annual totals."""

    footer_fields = split_record(footer_record)
    expected_fields = 3 + len(MONTH_NAMES) + 2
    if len(footer_fields) != expected_fields:
        raise ValueError(
            f"Finance template FOOTER has {len(footer_fields)} fields; expected {expected_fields}."
        )
    prefix = [footer_fields[0], f"{1:018d}", f"{1:04d}"]
    amounts = [format_amount(amount) for amount in summary.monthly_amounts]
    return "|".join(prefix + amounts + [format_amount(summary.annual_total), format_amount(ZERO)])


def transform_template(
    template_text: str,
    summary: FinanceSummary,
    financial_year: int,
    directory_id: str,
) -> str:
    """Return the complete three-record finance file text."""

    header, detail, footer = reconstruct_template_records(template_text)
    updated_detail = update_detail_record(header, detail, summary, financial_year, directory_id)
    updated_footer = update_footer_record(footer, summary)
    return "\n".join((header, updated_detail, updated_footer)) + "\n"


def update_member_detail_record(
    header_record: str,
    detail_record: str,
    summary: FinanceSummary,
    financial_year: int,
    directory_id: str,
) -> str:
    """Return a MEMBER detail with identifiers and policy dates updated."""

    header_fields = split_record(header_record)
    detail_fields = align_detail_fields(header_fields, detail_record, "Member")
    required = ("TAX YEAR", "CLIENT ID", "POLICY ID", "POLICY ON DATE", "POLICY OFF DATE")
    positions = build_required_positions(header_fields, required, "Member")
    updated = list(detail_fields)
    updated[positions["TAX YEAR"]] = str(financial_year)
    updated[positions["CLIENT ID"]] = directory_id
    updated[positions["POLICY ID"]] = summary.policy_number
    updated[positions["POLICY ON DATE"]] = summary.earliest_system_date.strftime("%Y/%m/%d")
    updated[positions["POLICY OFF DATE"]] = ""
    return "|".join(updated)


def update_advisor_detail_record(
    header_record: str,
    detail_record: str,
    summary: FinanceSummary,
    financial_year: int,
    directory_id: str,
    advisor_id: str,
) -> str:
    """Return an ADVISOR detail with its identifiers updated."""

    header_fields = split_record(header_record)
    detail_fields = align_short_advisor_detail(header_fields, detail_record)
    required = ("TAX YEAR", "CLIENT ID", "POLICY ID", "ADVISOR ID")
    positions = build_required_positions(header_fields, required, "Advisor")
    updated = list(detail_fields)
    updated[positions["TAX YEAR"]] = str(financial_year)
    updated[positions["CLIENT ID"]] = directory_id
    updated[positions["POLICY ID"]] = summary.policy_number
    updated[positions["ADVISOR ID"]] = advisor_id
    return "|".join(updated)


def update_simple_footer(footer_record: str, expected_type: str, template_name: str) -> str:
    """Return a one-detail footer while preserving its member/advisor type."""

    fields = split_record(footer_record)
    if len(fields) != 3 or fields[0].strip().upper() != "FOOTER":
        raise ValueError(f"{template_name} template FOOTER must contain three fields.")
    if fields[2].strip() != expected_type:
        raise ValueError(
            f"{template_name} template FOOTER type must be {expected_type}; found {fields[2]!r}."
        )
    return "|".join((fields[0], f"{1:018d}", expected_type))


def transform_member_template(
    template_text: str,
    summary: FinanceSummary,
    financial_year: int,
    directory_id: str,
) -> str:
    """Return the complete member file text."""

    header, detail, footer = reconstruct_template_records(template_text)
    updated_detail = update_member_detail_record(
        header, detail, summary, financial_year, directory_id
    )
    updated_footer = update_simple_footer(footer, "0002", "Member")
    return "\n".join((header, updated_detail, updated_footer)) + "\n"


def transform_advisor_template(
    template_text: str,
    summary: FinanceSummary,
    financial_year: int,
    directory_id: str,
    advisor_id: str,
) -> str:
    """Return the complete advisor file text."""

    header, detail, footer = reconstruct_template_records(template_text)
    updated_detail = update_advisor_detail_record(
        header, detail, summary, financial_year, directory_id, advisor_id
    )
    updated_footer = update_simple_footer(footer, "0003", "Advisor")
    return "\n".join((header, updated_detail, updated_footer)) + "\n"


def build_output_path(
    base_directory: Path, policy_number: str, financial_year: int, file_name: str
) -> Path:
    """Build one required policy/year output path."""

    return base_directory / policy_number / str(financial_year) / file_name


def format_output_timestamp(generated_at: datetime) -> str:
    """Return the output timestamp segment as YYYYMMDDHHMM."""

    return generated_at.strftime("%Y%m%d%H%M")


def build_output_file_name(
    file_prefix: str, financial_year: int, sequence_number: str, generated_at: datetime
) -> str:
    """Build one generated output file name without an extension."""

    return (
        f"RMM1ITC{file_prefix}{financial_year}"
        f"{sequence_number}{format_output_timestamp(generated_at)}"
    )


def write_output_file(output_path: Path, contents: str) -> Path:
    """Create output directories and write one generated text file."""

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(contents, encoding="utf-8", newline="\n")
    return output_path


def generate_finance_files(
    config: Config, generated_at: datetime | None = None
) -> tuple[GeneratedFiles, FinanceSummary]:
    """Run the shared pipeline and write finance, member, and advisor files."""

    checked = validate_config(config)
    output_timestamp = generated_at or datetime.now()
    source_rows = read_source_rows(checked.input_excel)
    code_filtered = filter_by_codes(source_rows)
    transactions = filter_and_parse_financial_year(code_filtered, checked.financial_year)
    summary = aggregate_transactions(require_transactions(transactions))
    directory_id = validate_fifteen_digit_id(checked.directory_id, "Directory ID")
    advisor_id = validate_advisor_id(checked.advisor_id)
    finance_text = transform_template(
        checked.template_file.read_text(encoding="utf-8-sig"),
        summary,
        checked.financial_year,
        directory_id,
    )
    member_text = transform_member_template(
        checked.member_template_file.read_text(encoding="utf-8-sig"),
        summary,
        checked.financial_year,
        directory_id,
    )
    advisor_text = transform_advisor_template(
        checked.advisor_template_file.read_text(encoding="utf-8-sig"),
        summary,
        checked.financial_year,
        directory_id,
        advisor_id,
    )
    paths = GeneratedFiles(
        financial=build_output_path(
            checked.output_base_directory,
            summary.policy_number,
            checked.financial_year,
            build_output_file_name(
                FINANCE_FILE_PREFIX,
                checked.financial_year,
                FINANCE_FILE_SEQUENCE,
                output_timestamp,
            ),
        ),
        member=build_output_path(
            checked.output_base_directory,
            summary.policy_number,
            checked.financial_year,
            build_output_file_name(
                MEMBER_FILE_PREFIX,
                checked.financial_year,
                MEMBER_FILE_SEQUENCE,
                output_timestamp,
            ),
        ),
        advisor=build_output_path(
            checked.output_base_directory,
            summary.policy_number,
            checked.financial_year,
            build_output_file_name(
                ADVISOR_FILE_PREFIX,
                checked.financial_year,
                ADVISOR_FILE_SEQUENCE,
                output_timestamp,
            ),
        ),
    )
    written = GeneratedFiles(
        financial=write_output_file(paths.financial, finance_text),
        member=write_output_file(paths.member, member_text),
        advisor=write_output_file(paths.advisor, advisor_text),
    )
    return written, summary


def generate_finance_files_for_configs(
    configs: Sequence[Config], generated_at: datetime | None = None
) -> tuple[tuple[GeneratedFiles, FinanceSummary], ...]:
    """Generate output files for each configured workbook/year run."""

    if not configs:
        raise ValueError("At least one generation configuration is required.")
    output_timestamp = generated_at or datetime.now()
    return tuple(generate_finance_files(config, output_timestamp) for config in configs)


def main() -> None:
    """Prompt for runtime settings and generate all three files."""

    results = generate_finance_files_for_configs(interactive_configs())
    for index, (output_files, summary) in enumerate(results, start=1):
        if len(results) > 1:
            print(f"Run {index} of {len(results)}")
        print(f"Generated finance file: {output_files.financial}")
        print(f"Generated member file: {output_files.member}")
        print(f"Generated advisor file: {output_files.advisor}")
        print(f"Policy Number: {summary.policy_number}")
        print(f"Annual Total: {format_amount(summary.annual_total)}")


if __name__ == "__main__":
    main()
