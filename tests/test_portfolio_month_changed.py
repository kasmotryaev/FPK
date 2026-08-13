"""Regression tests for simultaneous portfolio and month changes."""
import unittest

from app.importer import _collapse_portfolio_changes


def _event(event_type, row_id, month, portfolio, amount, contract="C-1"):
    return {
        "fp_row_id": row_id,
        "event_type": event_type,
        "field_label": None,
        "old_value": None,
        "new_value": None,
        "amount_before": amount if event_type == "deactivated" else None,
        "amount_after": amount if event_type == "new" else 0.0,
        "month": month,
        "client_name": "BANK",
        "project_num": "P-1",
        "project_name": "Same project",
        "section": "Support",
        "portfolio": portfolio,
        "contract_num": contract,
    }


class PortfolioAndMonthChangedTest(unittest.TestCase):
    def test_forecast_to_fact_is_collapsed_when_month_also_changes(self):
        events = [
            _event("new", 20, "August", "\u0424\u0430\u043a\u0442", 37318.40),
            _event("deactivated", 10, "July", "0-100", 37318.40),
        ]

        collapsed = _collapse_portfolio_changes(events)

        self.assertEqual(len(collapsed), 1)
        event = collapsed[0]
        self.assertEqual(event["event_type"], "portfolio_changed")
        self.assertEqual(event["old_value"], "0-100")
        self.assertEqual(event["new_value"], "\u0424\u0430\u043a\u0442")
        self.assertEqual(event["month"], "August")
        self.assertEqual(event["amount_after"], 37318.40)

    def test_different_project_is_not_collapsed(self):
        old_event = _event("deactivated", 10, "July", "0-100", 1000)
        new_event = _event("new", 20, "August", "\u0424\u0430\u043a\u0442", 1000)
        new_event["project_num"] = "P-2"

        collapsed = _collapse_portfolio_changes([old_event, new_event])

        self.assertEqual(len(collapsed), 2)

    def test_amount_difference_over_one_rouble_is_not_collapsed(self):
        events = [
            _event("deactivated", 10, "July", "0-100", 1000),
            _event("new", 20, "August", "\u0424\u0430\u043a\u0442", 1001.01),
        ]

        collapsed = _collapse_portfolio_changes(events)

        self.assertEqual(len(collapsed), 2)


if __name__ == "__main__":
    unittest.main()
