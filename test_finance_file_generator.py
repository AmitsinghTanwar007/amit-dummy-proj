"""Tests for the finance file generator."""

from __future__ import annotations

import tempfile
import unittest
from decimal import Decimal
from pathlib import Path

from finance_file_generator import (
    Config,
    MONTH_NAMES,
    SourceRow,
    default_config,
    filter_by_codes,
    format_amount,
    generate_finance_files,
    interactive_config,
    reconstruct_template_records,
    split_record,
)


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


class IntegrationTests(unittest.TestCase):
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
            output_files, summary = generate_finance_files(config)

            self.assertEqual(
                output_files.financial,
                Path(temporary_directory) / "563410000070095" / "2026" / "financial.txt",
            )
            self.assertEqual(
                output_files.member,
                Path(temporary_directory) / "563410000070095" / "2026" / "member.txt",
            )
            self.assertEqual(
                output_files.advisor,
                Path(temporary_directory) / "563410000070095" / "2026" / "advisor.txt",
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
