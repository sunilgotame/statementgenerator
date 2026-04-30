from __future__ import annotations

import unittest
from collections import Counter
from datetime import date
from statistics import mean

from .generator import StatementConfig, build_quarter_schedule, generate_statement
from .utils import next_business_day


class GeneratorTests(unittest.TestCase):
    def build_config(self) -> StatementConfig:
        return StatementConfig(
            bank_name="Nepal Bank Limited",
            branch_name="Main Branch",
            customer_name="Rubi Test",
            customer_address="Kathmandu, Nepal",
            account_number="1234567890",
            account_type="Saving Account",
            member_id="",
            currency="NPR",
            reference_no="",
            opening_date=date(2022, 5, 10),
            start_date=date(2024, 8, 8),
            end_date=date(2025, 9, 21),
            opening_balance=1_350_000,
            target_closing_balance=2_250_000,
            interest_rate=8.0,
            tax_rate=6.0,
            cheque_start=6_247_362,
            deposit_text="Cash Deposit",
            withdrawal_text="Cheque Withdrawal",
            interest_text="Interest",
            tax_text="Tax",
            first_date_description="Opening Balance",
            last_date_description="Balance C/F",
            deposit_names=["Self", "Kamala Pandey", "Santosh Thapa"],
            withdrawal_names=["Self", "Kabita Thapa"],
            holiday_dates={date(2025, 9, 22)},
            seed=123456,
        )

    def test_generation_reaches_target_closely(self) -> None:
        result = generate_statement(self.build_config())
        self.assertLessEqual(abs(result.final_balance - 2_250_000), 3_000)
        self.assertGreaterEqual(len(result.rows), 40)
        self.assertLessEqual(len(result.rows), 55)

    def test_deposit_amounts_are_not_monotonic(self) -> None:
        result = generate_statement(self.build_config())
        deposits = [event.amount for event in result.events if event.event_type == "deposit"]
        self.assertGreater(len(deposits), 3)
        self.assertNotEqual(deposits, sorted(deposits))

    def test_issue_date_matches_next_working_day(self) -> None:
        result = generate_statement(self.build_config())
        expected = next_business_day(result.last_transaction_date, self.build_config().holiday_dates, include_self=False)
        self.assertEqual(result.issue_date, expected)
        self.assertNotEqual(result.issue_date.weekday(), 5)
        self.assertNotIn(result.issue_date, self.build_config().holiday_dates)

    def test_january_2025_quarter_date_uses_13th(self) -> None:
        schedule = build_quarter_schedule(date(2025, 1, 1), date(2025, 1, 31), set())
        self.assertIn((date(2025, 1, 13), date(2025, 1, 13)), schedule)
        self.assertNotIn((date(2025, 1, 14), date(2025, 1, 14)), schedule)

    def test_interest_and_tax_stay_on_blocked_quarter_date_without_user_rows(self) -> None:
        config = self.build_config()
        config.holiday_dates = {date(2025, 7, 16), date(2025, 9, 22)}
        result = generate_statement(config)
        rows_on_quarter_day = [row for row in result.rows if row.date == date(2025, 7, 16)]
        categories = [row.category for row in rows_on_quarter_day]
        self.assertIn("interest", categories)
        self.assertIn("tax", categories)
        self.assertFalse(any(row.category in {"deposit", "withdrawal"} for row in rows_on_quarter_day))

    def test_first_and_last_date_descriptions_are_applied(self) -> None:
        config = self.build_config()
        config.first_date_description = "Balance B/F"
        config.last_date_description = "Closing Balance"
        result = generate_statement(config)
        opening_row = next(row for row in result.rows if row.category == "opening")
        closing_row = next(row for row in result.rows if row.category == "closing")
        self.assertEqual(opening_row.description, "Balance B/F")
        self.assertEqual(closing_row.description, "Closing Balance")

    def test_monthly_transaction_mix_varies_with_seed(self) -> None:
        config_a = self.build_config()
        config_b = self.build_config()
        config_b.seed = 123457
        result_a = generate_statement(config_a)
        result_b = generate_statement(config_b)
        month_counts_a = Counter((row.date.year, row.date.month) for row in result_a.rows if row.category in {"deposit", "withdrawal"})
        month_counts_b = Counter((row.date.year, row.date.month) for row in result_b.rows if row.category in {"deposit", "withdrawal"})
        self.assertNotEqual(dict(month_counts_a), dict(month_counts_b))

    def test_deposit_count_stays_above_withdrawals(self) -> None:
        for seed in (123456, 123457, 123458, 123459):
            config = self.build_config()
            config.seed = seed
            result = generate_statement(config)
            deposits = result.summary.deposit_count
            withdrawals = result.summary.withdrawal_count
            self.assertGreater(deposits, withdrawals)
            self.assertLessEqual(deposits - withdrawals, max(12, withdrawals))

    def test_consecutive_deposit_runs_stay_within_three(self) -> None:
        for seed in (123450, 123451, 123452, 123453, 123454, 123455, 123456, 123457, 123458, 123459):
            config = self.build_config()
            config.seed = seed
            result = generate_statement(config)
            deposit_runs: list[int] = []
            current_run = 0
            longest_run = 0
            for event in result.events:
                if event.event_type == "deposit":
                    current_run += 1
                    longest_run = max(longest_run, current_run)
                else:
                    if current_run:
                        deposit_runs.append(current_run)
                    current_run = 0
            if current_run:
                deposit_runs.append(current_run)
            self.assertLessEqual(longest_run, 3)
            self.assertTrue(any(run_length != 2 for run_length in deposit_runs))

    def test_consecutive_withdrawal_runs_stay_within_two(self) -> None:
        for seed in (123450, 123451, 123452, 123453, 123454, 123455, 123456, 123457, 123458, 123459):
            config = self.build_config()
            config.seed = seed
            result = generate_statement(config)
            withdrawal_runs: list[int] = []
            current_run = 0
            longest_run = 0
            for event in result.events:
                if event.event_type == "withdrawal":
                    current_run += 1
                    longest_run = max(longest_run, current_run)
                else:
                    if current_run:
                        withdrawal_runs.append(current_run)
                    current_run = 0
            if current_run:
                withdrawal_runs.append(current_run)
            self.assertLessEqual(longest_run, 2)
            self.assertLessEqual(sum(1 for run_length in withdrawal_runs if run_length == 2), 3)

    def test_deposit_run_lengths_vary_across_seed_sample(self) -> None:
        saw_single = False
        saw_triple = False
        for seed in (123450, 123451, 123452, 123453, 123454, 123455, 123456, 123457, 123458, 123459):
            config = self.build_config()
            config.seed = seed
            result = generate_statement(config)
            current_run = 0
            for event in result.events:
                if event.event_type == "deposit":
                    current_run += 1
                else:
                    if current_run == 1:
                        saw_single = True
                    elif current_run == 3:
                        saw_triple = True
                    current_run = 0
            if current_run == 1:
                saw_single = True
            elif current_run == 3:
                saw_triple = True
        self.assertTrue(saw_single)
        self.assertTrue(saw_triple)

    def test_run_mix_stays_near_requested_percentages(self) -> None:
        single_event_ratios: list[float] = []
        double_event_ratios: list[float] = []
        triple_event_ratios: list[float] = []
        withdrawal_double_event_ratios: list[float] = []
        for seed in (123450, 123451, 123452, 123453, 123454, 123455, 123456, 123457, 123458, 123459):
            config = self.build_config()
            config.seed = seed
            result = generate_statement(config)
            deposit_runs: list[int] = []
            withdrawal_runs: list[int] = []
            current_run = 0
            current_type = ""
            for event in result.events:
                if event.event_type == current_type:
                    current_run += 1
                else:
                    if current_type == "deposit":
                        deposit_runs.append(current_run)
                    elif current_type == "withdrawal":
                        withdrawal_runs.append(current_run)
                    current_type = event.event_type
                    current_run = 1
            if current_type == "deposit":
                deposit_runs.append(current_run)
            elif current_type == "withdrawal":
                withdrawal_runs.append(current_run)

            deposit_total = sum(deposit_runs)
            withdrawal_total = sum(withdrawal_runs)
            single_event_ratios.append(sum(run for run in deposit_runs if run == 1) / deposit_total)
            double_event_ratios.append(sum(run for run in deposit_runs if run == 2) / deposit_total)
            triple_event_ratios.append(sum(run for run in deposit_runs if run == 3) / deposit_total)
            withdrawal_double_event_ratios.append(sum(run for run in withdrawal_runs if run == 2) / withdrawal_total)

        self.assertGreaterEqual(mean(single_event_ratios), 0.15)
        self.assertLessEqual(mean(single_event_ratios), 0.35)
        self.assertGreaterEqual(mean(double_event_ratios), 0.50)
        self.assertLessEqual(mean(double_event_ratios), 0.70)
        self.assertGreaterEqual(mean(triple_event_ratios), 0.10)
        self.assertLessEqual(mean(triple_event_ratios), 0.25)
        self.assertGreaterEqual(mean(withdrawal_double_event_ratios), 0.20)
        self.assertLessEqual(mean(withdrawal_double_event_ratios), 0.40)

    def test_deposit_amounts_limit_repeats_and_keep_rounding_mix(self) -> None:
        result = generate_statement(self.build_config())
        deposits = [int(event.amount) for event in result.events if event.event_type == "deposit"]
        self.assertGreater(len(deposits), 6)
        counts = Counter(deposits)
        duplicate_groups = [value for value, count in counts.items() if count > 1]
        self.assertLessEqual(max(counts.values()), 2)
        self.assertLessEqual(len(duplicate_groups), 2)
        self.assertGreater(max(deposits), 50_000)
        self.assertTrue(any(value <= 25_000 for value in deposits))
        hundred_only_ratio = sum(1 for value in deposits if value % 500 != 0) / len(deposits)
        self.assertGreaterEqual(hundred_only_ratio, 0.10)
        self.assertLessEqual(hundred_only_ratio, 0.30)


def run_tests() -> unittest.result.TestResult:
    suite = unittest.defaultTestLoader.loadTestsFromTestCase(GeneratorTests)
    return unittest.TextTestRunner(verbosity=2).run(suite)


if __name__ == "__main__":
    run_tests()
