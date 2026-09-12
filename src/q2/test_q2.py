from __future__ import annotations

import sys
import unittest
from datetime import date, timedelta
from pathlib import Path

import numpy as np

sys.path.insert(0, str(Path(__file__).resolve().parent))
import q2  # noqa: E402


class TimeMappingTests(unittest.TestCase):
    def test_physical_interval_labels(self) -> None:
        labels = q2.physical_interval_labels()
        self.assertEqual(len(labels), 144)
        self.assertEqual(labels[0], "0:00-0:10")
        self.assertEqual(labels[60], "10:00-10:10")
        self.assertEqual(labels[-1], "23:50-00:00+1")

    def test_cross_day_plan_mapping(self) -> None:
        current = np.arange(1, 145, dtype=float)
        following = np.arange(1001, 1145, dtype=float)
        displayed = np.concatenate((current[1:], following[:1]))
        self.assertEqual(len(displayed), 144)
        self.assertEqual(displayed[0], 2.0)
        self.assertEqual(displayed[-2], 144.0)
        self.assertEqual(displayed[-1], 1001.0)


class ClassificationTests(unittest.TestCase):
    def setUp(self) -> None:
        self.january_dates = tuple(date(2025, 1, 1) + timedelta(days=i) for i in range(31))
        # Friday/Saturday are deliberately lower; every row has 144 kWh total per unit value.
        self.january_load = np.stack(
            [
                np.full(q2.T, 50.0 if day.weekday() in q2.LOW_LOAD_WEEKDAYS else 100.0)
                for day in self.january_dates
            ]
        )

    def test_january_only_classification_passes(self) -> None:
        rows, summary = q2.validate_january_load_classes(
            self.january_dates, self.january_load
        )
        self.assertTrue(summary["passed"])
        self.assertFalse(summary["uses_february_to_december"])
        self.assertEqual(len(rows), 9)
        self.assertEqual(
            set(summary["weekday_mean_order_low_to_high"][:2]),
            {"周五", "周六"},
        )

    def test_classification_rejects_non_january_data(self) -> None:
        with self.assertRaises(ValueError):
            q2.validate_january_load_classes(
                self.january_dates + (date(2025, 2, 1),),
                np.vstack((self.january_load, np.full((1, q2.T), 1.0))),
            )


class ForecastTests(unittest.TestCase):
    @staticmethod
    def history(days: int = 10) -> q2.ForecastHistory:
        history = q2.ForecastHistory.empty()
        first = date(2025, 1, 1)
        for i in range(days):
            target = first + timedelta(days=i)
            draft = q2.forecast_one_day(target, history)
            load = np.full(q2.T, 100.0 + i)
            pv = np.zeros(q2.T)
            pv[40:100] = 20.0 + i
            q2.append_observation(history, target, load, pv, draft)
        return history

    def test_inverse_distance_weights(self) -> None:
        dates = [date(2025, 1, i) for i in (7, 8, 9)]
        selected = q2.inverse_distance_selection(
            date(2025, 1, 10), dates, range(3)
        )
        self.assertEqual([item[1] for item in selected], [1, 2, 3])
        np.testing.assert_allclose(
            [item[2] for item in selected], [6 / 11, 3 / 11, 2 / 11]
        )

    def test_forecast_rejects_target_or_future_in_history(self) -> None:
        history = self.history()
        with self.assertRaises(ValueError):
            q2.forecast_one_day(history.dates[-1], history)

    def test_paired_scenario_dates_and_weights_are_causal(self) -> None:
        history = self.history(15)
        target = date(2025, 1, 16)
        draft = q2.forecast_one_day(target, history)
        self.assertTrue(all(day < target for day in draft.scenario_dates))
        self.assertEqual(
            draft.scenario_load_residuals.shape,
            draft.scenario_pv_residuals.shape,
        )
        bundle = q2.require_forecast_bundle(draft)
        audit = q2.forecast_audit_rows(bundle, "test")
        scenario_rows = [
            row for row in audit if row["record_type"] == "paired_final_residual_scenario"
        ]
        self.assertEqual(len(scenario_rows), len(bundle.scenario_dates))
        self.assertAlmostEqual(sum(row["weight"] for row in scenario_rows), 1.0)

    def test_future_observation_cannot_change_existing_forecast(self) -> None:
        history = self.history(10)
        target = date(2025, 1, 11)
        first = q2.forecast_one_day(target, history)
        copied = q2.ForecastHistory(
            dates=list(history.dates),
            load_actual=[value.copy() for value in history.load_actual],
            pv_actual=[value.copy() for value in history.pv_actual],
            load_primary_residual=[None if v is None else v.copy() for v in history.load_primary_residual],
            pv_primary_residual=[None if v is None else v.copy() for v in history.pv_primary_residual],
            load_final_residual=[None if v is None else v.copy() for v in history.load_final_residual],
            pv_final_residual=[None if v is None else v.copy() for v in history.pv_final_residual],
        )
        second = q2.forecast_one_day(target, copied)
        np.testing.assert_array_equal(first.load_forecast, second.load_forecast)
        np.testing.assert_array_equal(first.pv_forecast, second.pv_forecast)


class OptimizationAndExecutorTests(unittest.TestCase):
    def test_small_scenario_lp(self) -> None:
        prices = np.full(q2.T, 0.5)
        scenario_net = np.vstack((np.full(q2.T, 500.0), np.full(q2.T, 600.0)))
        result = q2.solve_day_ahead(
            date(2025, 2, 1),
            prices,
            scenario_net,
            q2.INITIAL_SOC_KWH,
            q2.ETA_DISCHARGE * float(np.min(prices)),
        )
        self.assertIn("Optimal", result.status)
        self.assertLessEqual(result.max_balance_residual_kwh, q2.CHECK_TOLERANCE)
        self.assertLessEqual(result.max_soc_residual_kwh, q2.CHECK_TOLERANCE)
        self.assertLessEqual(
            result.max_simultaneous_product_kwh2,
            q2.SIMULTANEOUS_PRODUCT_TOLERANCE,
        )

    def test_analytical_executor_constraints(self) -> None:
        prices = np.linspace(0.4, 1.4, q2.T)
        plan = np.full(q2.T, 500.0)
        actual_net = np.concatenate((np.full(72, 300.0), np.full(72, 900.0)))
        result = q2.execute_analytical(
            prices, actual_net, plan, q2.INITIAL_SOC_KWH
        )
        self.assertLessEqual(result.max_balance_residual_kwh, q2.CHECK_TOLERANCE)
        self.assertGreaterEqual(result.min_soc_kwh, q2.SOC_MIN_KWH)
        self.assertLessEqual(result.max_soc_kwh, q2.SOC_MAX_KWH)

    def test_dp_executor_constraints(self) -> None:
        prices = np.linspace(0.4, 1.4, q2.T)
        plan = np.full(q2.T, 500.0)
        actual_net = np.full(q2.T, 550.0)
        scenario_net = np.vstack((np.full(q2.T, 540.0), np.full(q2.T, 560.0)))
        result = q2.execute_dp(
            prices,
            actual_net,
            plan,
            q2.INITIAL_SOC_KWH,
            scenario_net,
            q2.ETA_DISCHARGE * float(np.min(prices)),
            10.0,
        )
        self.assertLessEqual(result.max_balance_residual_kwh, q2.CHECK_TOLERANCE)
        self.assertGreaterEqual(result.min_soc_kwh, q2.SOC_MIN_KWH)
        self.assertLessEqual(result.max_soc_kwh, q2.SOC_MAX_KWH)

    def test_emergency_interval_merging(self) -> None:
        emergency = np.zeros(q2.T)
        emergency[0:3] = [1.0, 2.0, 3.0]
        emergency[60:62] = [4.0, 5.0]
        emergency[-1] = 6.0
        intervals = q2.merge_emergency_intervals(emergency)
        self.assertEqual(
            [item["physical_interval"] for item in intervals],
            ["0:00-0:30", "10:00-10:20", "23:50-00:00+1"],
        )
        self.assertEqual(
            [item["emergency_energy_kwh"] for item in intervals],
            [6.0, 9.0, 6.0],
        )


if __name__ == "__main__":
    unittest.main(verbosity=2)
