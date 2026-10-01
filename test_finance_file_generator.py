"""Tests for the finance file generator."""

from __future__ import annotations

import tempfile
import unittest
from datetime import datetime
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook

from finance_file_generator import (
    Config,
    MONTH_NAMES,
    SourceRow,
    default_config,
    discover_excel_files,
    filter_by_codes,
    format_amount,
    generate_finance_files,
    generate_finance_files_for_configs,
    interactive_config,
    interactive_configs,
    parse_financial_years_input,
    parse_run_mode_input,
    reconstruct_template_records,
    split_record,
)


EXCEL_HEADERS = (
    "Policy Number",
    "System Date",
    "Transaction Code",
    "Reversal File Code",
    "Transaction Amount",
)


def write_dummy_workbook(path: Path, rows: tuple[tuple[object, ...], ...]) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.append(EXCEL_HEADERS)
    for row in rows:
        worksheet.append(row)
    workbook.save(path)


class FilteringTests(unittest.TestCase):
    def test_only_blank_and_text_zero_reversal_codes_qualify(self) -> None:
        rows = (
            SourceRow("1", "2025-03-01", "NA", None, 1),
            SourceRow("1", "2025-03-01", "NA", "", 1),
            SourceRow("1", "2025-03-01", "NA", "00000", 1),
            SourceRow("1", "2025-03-01", "NA", 0, 1),
            SourceRow("1", "2025-03-01", "NB", "00000", 1),
        )
        self.assertEqual(len(filter_by_codes(rows)), 3)


class FormattingTests(unittest.TestCase):
    def test_amount_format(self) -> None:
        self.assertEqual(format_amount(Decimal("336.17")), "0000000000000336.17")
        self.assertEqual(format_amount(Decimal("0")), "0000000000000000.00")


class InteractiveConfigTests(unittest.TestCase):
    def test_five_responses_build_expected_config(self) -> None:
        project_config = default_config()
        responses = iter(
            (
                str(project_config.input_excel),
                "2026",
                "000000023353484",
                "87689",
                "not-yet-created-output",
            )
        )

        config = interactive_config(lambda _prompt: next(responses))

        self.assertEqual(config.input_excel, project_config.input_excel.resolve())
        self.assertEqual(config.financial_year, 2026)
        self.assertEqual(config.directory_id, "000000023353484")
        self.assertEqual(config.advisor_id, "87689")
        self.assertEqual(
            config.output_base_directory,
            (Path.cwd() / "not-yet-created-output").resolve(),
        )

    def test_comma_separated_years_are_parsed_in_given_order(self) -> None:
        self.assertEqual(parse_financial_years_input("2026, 2024,2028"), (2026, 2024, 2028))

    def test_duplicate_years_are_rejected(self) -> None:
        with self.assertRaisesRegex(ValueError, "must not contain duplicates"):
            parse_financial_years_input("2026,2026")

    def test_mode_accepts_numbered_and_text_choices(self) -> None:
        self.assertEqual(parse_run_mode_input("1"), "single")
        self.assertEqual(parse_run_mode_input("multiple"), "multiple")

    def test_single_mode_builds_one_config_per_requested_year(self) -> None:
        project_config = default_config()
        responses = iter(
            (
                "single",
                str(project_config.input_excel),
                "2026,2024",
                "000000023353484",
                "87689",
                "not-yet-created-output",
            )
        )

        configs = interactive_configs(lambda _prompt: next(responses))

        self.assertEqual([config.input_excel for config in configs], [project_config.input_excel.resolve()] * 2)
        self.assertEqual([config.financial_year for config in configs], [2026, 2024])
        self.assertEqual(
            [config.output_base_directory for config in configs],
            [(Path.cwd() / "not-yet-created-output").resolve()] * 2,
        )

    def test_multiple_mode_discovers_folder_files_and_prompts_each_year_list(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            folder = Path(temporary_directory)
            first = folder / "a.xlsx"
            second = folder / "b.xlsm"
            ignored_text = folder / "notes.txt"
            ignored_temporary = folder / "~$open.xlsx"
            for path in (first, second, ignored_text, ignored_temporary):
                path.write_text("", encoding="utf-8")

            responses = iter(
                (
                    "multiple",
                    str(folder),
                    "2026",
                    "2024,2028",
                    "000000023353484",
                    "87689",
                    "not-yet-created-output",
                )
            )

            configs = interactive_configs(lambda _prompt: next(responses))
            self.assertEqual(discover_excel_files(folder), (first, second))

        self.assertEqual(
            [(config.input_excel, config.financial_year) for config in configs],
            [(first, 2026), (second, 2024), (second, 2028)],
        )


class IntegrationTests(unittest.TestCase):
    def test_multiple_dummy_workbooks_generate_each_requested_year(self) -> None:
        project_config = default_config()
        with tempfile.TemporaryDirectory() as temporary_directory:
            base = Path(temporary_directory)
            input_folder = base / "input"
            output_folder = base / "output"
            input_folder.mkdir()

            first = input_folder / "first.xlsx"
            second = input_folder / "second.xlsx"
            write_dummy_workbook(
                first,
                (
                    ("111111111111111", "2023-03-10", "NA", "00000", "100.00"),
                    ("111111111111111", "2025-04-20", "NA", "", "200.00"),
                    ("111111111111111", "2025-05-01", "NB", "", "999.00"),
                ),
            )
            write_dummy_workbook(
                second,
                (
                    ("222222222222222", "2026-12-15", "NA", None, "300.00"),
                    ("222222222222222", "2027-01-05", "NA", "00000", "400.00"),
                ),
            )

            responses = iter(
                (
                    "multiple",
                    str(input_folder),
                    "2024,2026",
                    "2027",
                    "000000023353484",
                    "87689",
                    str(output_folder),
                )
            )

            configs = interactive_configs(lambda _prompt: next(responses))
            results = generate_finance_files_for_configs(
                configs, datetime(2026, 10, 1, 9, 45, 59)
            )

            self.assertEqual(
                [(summary.policy_number, summary.annual_total) for _files, summary in results],
                [
                    ("111111111111111", Decimal("100.00")),
                    ("111111111111111", Decimal("200.00")),
                    ("222222222222222", Decimal("700.00")),
                ],
            )
            expected_outputs = (
                output_folder
                / "111111111111111"
                / "2024"
                / "RMM1ITCFIN2024001202610010945",
                output_folder
                / "111111111111111"
                / "2026"
                / "RMM1ITCFIN2026001202610010945",
                output_folder
                / "222222222222222"
                / "2027"
                / "RMM1ITCFIN2027001202610010945",
            )
            self.assertTrue(all(path.is_file() for path in expected_outputs))
            self.assertTrue(all(files.member.is_file() and files.advisor.is_file() for files, _summary in results))

    def test_dummy_workbook_generates_expected_finance_file(self) -> None:
        project_config = default_config()
        with tempfile.TemporaryDirectory() as temporary_directory:
            config = Config(
                input_excel=project_config.input_excel,
                financial_year=2026,
                directory_id="000000023353484",
                template_file=project_config.template_file,
                member_template_file=project_config.member_template_file,
                advisor_template_file=project_config.advisor_template_file,
                advisor_id="87689",
                output_base_directory=Path(temporary_directory),
            )
            generated_at = datetime(2026, 10, 1, 9, 45, 59)
            output_files, summary = generate_finance_files(config, generated_at)

            self.assertEqual(
                output_files.financial,
                Path(temporary_directory)
                / "563410000070095"
                / "2026"
                / "RMM1ITCFIN2026001202610010945",
            )
            self.assertEqual(
                output_files.member,
                Path(temporary_directory)
                / "563410000070095"
                / "2026"
                / "RMM1ITCMEM2026002202610010945",
            )
            self.assertEqual(
                output_files.advisor,
                Path(temporary_directory)
                / "563410000070095"
                / "2026"
                / "RMM1ITCADV2026003202610010945",
            )
            self.assertTrue(all(path.is_file() for path in (
                output_files.financial, output_files.member, output_files.advisor
            )))
            self.assertEqual(summary.policy_number, "563410000070095")
            self.assertEqual(len(summary.monthly_amounts), 12)
            self.assertEqual(
                summary.monthly_amounts,
                (
                    Decimal("1961.67"),
                    Decimal("836.42"),
                    Decimal("436.17"),
                    Decimal("336.17"),
                    Decimal("349.61"),
                    Decimal("349.61"),
                    Decimal("1449.71"),
                    Decimal("349.61"),
                    Decimal("349.61"),
                    Decimal("1175.01"),
                    Decimal("349.61"),
                    Decimal("990.60"),
                ),
            )
            self.assertEqual(summary.annual_total, Decimal("8933.80"))
            self.assertEqual(summary.earliest_system_date.isoformat(), "2025-03-03")

            header, detail, footer = reconstruct_template_records(output_files.financial.read_text())
            header_fields = split_record(header)
            detail_fields = split_record(detail)
            positions = {name.strip().upper(): index for index, name in enumerate(header_fields)}

            self.assertIn("FUND ENTITY CODE", positions)
            self.assertIn("TAX CODE", positions)
            self.assertEqual(detail_fields[positions["TAX YEAR"]], "2026")
            self.assertEqual(detail_fields[positions["CLIENT ID"]], "000000023353484")
            self.assertEqual(detail_fields[positions["POLICY ID"]], "563410000070095")
            self.assertEqual(detail_fields[positions["MARCH"]], "0000000000001961.67")
            self.assertEqual(
                detail_fields[positions["CURRENT AMOUNT/INCOME PROTECTION AMOUNT"]],
                "0000000000008933.80",
            )
            self.assertEqual(detail_fields[positions["ARREARS AMOUNT"]], "0000000000000000.00")

            footer_fields = split_record(footer)
            self.assertEqual(footer_fields[1], "000000000000000001")
            self.assertEqual(footer_fields[2], "0001")
            self.assertEqual(footer_fields[3:15], [format_amount(value) for value in summary.monthly_amounts])
            self.assertEqual(footer_fields[15], "0000000000008933.80")
            self.assertEqual(footer_fields[16], "0000000000000000.00")
            self.assertEqual(len(MONTH_NAMES), 12)

            member_header, member_detail, member_footer = reconstruct_template_records(
                output_files.member.read_text()
            )
            member_positions = {
                name.strip().upper(): index
                for index, name in enumerate(split_record(member_header))
            }
            member_fields = split_record(member_detail)
            self.assertEqual(member_fields[member_positions["TAX YEAR"]], "2026")
            self.assertEqual(member_fields[member_positions["CLIENT ID"]], "000000023353484")
            self.assertEqual(member_fields[member_positions["POLICY ID"]], "563410000070095")
            self.assertEqual(member_fields[member_positions["POLICY ON DATE"]], "2025/03/03")
            self.assertEqual(member_fields[member_positions["POLICY OFF DATE"]], "")
            self.assertEqual(split_record(member_footer), ["FOOTER", "000000000000000001", "0002"])

            advisor_header, advisor_detail, advisor_footer = reconstruct_template_records(
                output_files.advisor.read_text()
            )
            advisor_positions = {
                name.strip().upper(): index
                for index, name in enumerate(split_record(advisor_header))
            }
            advisor_fields = split_record(advisor_detail)
            self.assertEqual(advisor_fields[advisor_positions["TAX YEAR"]], "2026")
            self.assertEqual(advisor_fields[advisor_positions["CLIENT ID"]], "000000023353484")
            self.assertEqual(advisor_fields[advisor_positions["POLICY ID"]], "563410000070095")
            self.assertEqual(advisor_fields[advisor_positions["ADVISOR ID"]], "87689")
            self.assertEqual(advisor_fields[advisor_positions["POSTAL ADDRESS POST CODE"]], "703")
            self.assertEqual(advisor_fields[advisor_positions["LICENCE NUMBER"]], "")
            self.assertEqual(len(advisor_fields), len(split_record(advisor_header)))
            self.assertEqual(split_record(advisor_footer), ["FOOTER", "000000000000000001", "0003"])


if __name__ == "__main__":
    unittest.main()
