"""Regression tests for project-transfer change events."""
import unittest

from app.importer import _collapse_project_moves


def _event(event_type, row_id, project_num, project_name, amount, contract="C-1"):
    return {
        "fp_row_id": row_id,
        "event_type": event_type,
        "field_label": None,
        "old_value": None,
        "new_value": None,
        "amount_before": amount if event_type == "deactivated" else None,
        "amount_after": amount if event_type == "new" else 0.0,
        "month": "August",
        "client_name": "BANK",
        "project_num": project_num,
        "project_name": project_name,
        "section": "Custom",
        "portfolio": "Fact",
        "contract_num": contract,
    }


class ProjectChangedTest(unittest.TestCase):
    def test_same_bank_and_amount_are_shown_as_project_change(self):
        events = [
            _event("new", 20, "2978393", "New project", 358064.64),
            _event("deactivated", 10, "2909922", "Old project", 358064.64),
        ]

        collapsed = _collapse_project_moves(events)

        self.assertEqual(len(collapsed), 1)
        event = collapsed[0]
        self.assertEqual(event["event_type"], "field_changed")
        self.assertEqual(event["field_label"], "\u041f\u0440\u043e\u0435\u043a\u0442 \u0438\u0437\u043c\u0435\u043d\u0438\u043b\u0441\u044f")
        self.assertEqual(event["old_value"], "2909922 - Old project")
        self.assertEqual(event["new_value"], "2978393 - New project")
        self.assertEqual(event["amount_after"], 358064.64)

    def test_different_bank_is_not_collapsed(self):
        old_event = _event("deactivated", 10, "P1", "Old project", 1000)
        new_event = _event("new", 20, "P2", "New project", 1000)
        new_event["client_name"] = "OTHER BANK"

        collapsed = _collapse_project_moves([old_event, new_event])

        self.assertEqual(len(collapsed), 2)
        self.assertEqual({event["event_type"] for event in collapsed}, {"new", "deactivated"})

    def test_amount_difference_over_one_rouble_is_not_collapsed(self):
        events = [
            _event("deactivated", 10, "P1", "Old project", 1000),
            _event("new", 20, "P2", "New project", 1001.01),
        ]

        collapsed = _collapse_project_moves(events)

        self.assertEqual(len(collapsed), 2)

    def test_exact_contract_is_preferred_for_equal_amounts(self):
        old = _event("deactivated", 10, "P1", "Old", 1000, contract="MATCH")
        wrong = _event("new", 20, "P2", "Wrong", 1000, contract="OTHER")
        right = _event("new", 30, "P3", "Right", 1000, contract="MATCH")

        collapsed = _collapse_project_moves([old, wrong, right])

        changed = next(event for event in collapsed if event["event_type"] == "field_changed")
        self.assertEqual(changed["new_value"], "P3 - Right")


if __name__ == "__main__":
    unittest.main()
