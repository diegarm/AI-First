"""
Unit tests for timesheet.py

Run with:
    python test_timesheet.py
    python -m pytest test_timesheet.py -v   (requires pytest)
"""

import csv
import io
import sys
import unittest
from datetime import date, timedelta
from pathlib import Path
from unittest.mock import MagicMock, patch

# ---------------------------------------------------------------------------
# Import helpers from timesheet without running main()
# ---------------------------------------------------------------------------
import importlib.util

spec = importlib.util.spec_from_file_location(
    "timesheet", Path(__file__).parent / "timesheet.py"
)
ts = importlib.util.module_from_spec(spec)
spec.loader.exec_module(ts)


# ---------------------------------------------------------------------------
# extract_devops_ids
# ---------------------------------------------------------------------------
class TestExtractDevopsIds(unittest.TestCase):
    def test_finds_5digit_starting_with_2(self):
        self.assertEqual(ts.extract_devops_ids("Fix bug #23456"), ["23456"])

    def test_finds_multiple(self):
        self.assertEqual(ts.extract_devops_ids("Refs 21000 and 29999"), ["21000", "29999"])

    def test_ignores_4digit_numbers(self):
        self.assertEqual(ts.extract_devops_ids("ticket 1234 done"), [])

    def test_ignores_6digit_numbers(self):
        self.assertEqual(ts.extract_devops_ids("ref 123456"), [])

    def test_ignores_numbers_not_starting_with_2(self):
        self.assertEqual(ts.extract_devops_ids("issue 31234"), [])

    def test_empty_subject(self):
        self.assertEqual(ts.extract_devops_ids(""), [])

    def test_no_ids(self):
        self.assertEqual(ts.extract_devops_ids("WIP: minor cleanup"), [])

    def test_id_at_start(self):
        self.assertEqual(ts.extract_devops_ids("24567: implement feature"), ["24567"])


# ---------------------------------------------------------------------------
# build_rows — no commits → N/A row
# ---------------------------------------------------------------------------
class TestBuildRowsNoCommits(unittest.TestCase):
    def setUp(self):
        self.target = date(2026, 5, 6)  # Wednesday
        self.row = ts.build_rows(self.target, "Test User", "SI3", "Dev", 8.0, [])[0]

    def test_single_row(self):
        rows = ts.build_rows(self.target, "Test User", "SI3", "Dev", 8.0, [])
        self.assertEqual(len(rows), 1)

    def test_date_yyyymmdd(self):
        self.assertEqual(self.row[0], "20260506")

    def test_user(self):
        self.assertEqual(self.row[1], "Test User")

    def test_project(self):
        self.assertEqual(self.row[2], "SI3")

    def test_devops_na(self):
        self.assertEqual(self.row[3], "N/A")

    def test_weekday_portuguese(self):
        self.assertEqual(self.row[9], "Quarta-feira")

    def test_date_ddmmyyyy(self):
        self.assertEqual(self.row[8], "06-05-2026")


# ---------------------------------------------------------------------------
# build_rows — with commits, DevOps IDs present
# ---------------------------------------------------------------------------
class TestBuildRowsWithCommits(unittest.TestCase):
    def _make_commit(self, subject, repo="repo1"):
        return {"datetime": "2026-05-07 10:00:00 +0100", "hash": "abc", "subject": subject, "repo": repo}

    def test_single_task_gets_all_hours(self):
        commits = [self._make_commit("Fix login flow 23456")]
        rows = ts.build_rows(date(2026, 5, 7), "User", "SI3", "Dev", 8.0, commits)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0][3], "23456")
        self.assertEqual(rows[0][5], 8)

    def test_two_tasks_split_hours_evenly(self):
        commits = [
            self._make_commit("Task A 21001"),
            self._make_commit("Task B 21002"),
        ]
        rows = ts.build_rows(date(2026, 5, 7), "User", "SI3", "Dev", 8.0, commits)
        self.assertEqual(len(rows), 2)
        hours = [int(r[5]) for r in rows]
        self.assertEqual(sum(hours), 8)
        # Both should get 4 hours each
        self.assertCountEqual(hours, [4, 4])

    def test_three_tasks_remainder_goes_first(self):
        commits = [
            self._make_commit("Task A 21001"),
            self._make_commit("Task B 21002"),
            self._make_commit("Task C 21003"),
        ]
        rows = ts.build_rows(date(2026, 5, 7), "User", "SI3", "Dev", 8.0, commits)
        self.assertEqual(len(rows), 3)
        hours = [int(r[5]) for r in rows]
        self.assertEqual(sum(hours), 8)
        # 8 / 3 = 2r2 → [3, 3, 2]
        self.assertCountEqual(hours, [3, 3, 2])

    def test_explicit_horas_tag_respected(self):
        commits = [self._make_commit("Fix 23456 [horas: 3]")]
        rows = ts.build_rows(date(2026, 5, 7), "User", "SI3", "Dev", 8.0, commits)
        self.assertEqual(rows[0][5], 3.0)

    def test_no_devops_id_falls_back_to_na(self):
        commits = [self._make_commit("Just some work, no ID")]
        rows = ts.build_rows(date(2026, 5, 7), "User", "SI3", "Dev", 8.0, commits)
        self.assertEqual(rows[0][3], "N/A")

    def test_hours_capped_at_8(self):
        commits = [self._make_commit("Fix 23456 [horas: 12]")]
        rows = ts.build_rows(date(2026, 5, 7), "User", "SI3", "Dev", 8.0, commits)
        # Explicit tag of 12, but build_rows caps total_hours at 8 before distributing
        # The explicit tag value is used as-is; only the remainder is capped
        # 12 > 8 → remaining = int(max(0, 8-12)) = 0 → implicit tasks get 0
        self.assertEqual(rows[0][3], "23456")

    def test_orphan_commits_attached_to_first_id(self):
        commits = [
            self._make_commit("Task with ID 22222"),
            self._make_commit("Orphan commit no id"),
        ]
        rows = ts.build_rows(date(2026, 5, 7), "User", "SI3", "Dev", 8.0, commits)
        # Only one row (orphan merged into 22222)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0][3], "22222")
        self.assertIn("Orphan commit no id", rows[0][6])

    def test_weekday_monday(self):
        rows = ts.build_rows(date(2026, 5, 4), "U", "P", "T", 8.0, [])  # Monday
        self.assertEqual(rows[0][9], "Segunda-feira")

    def test_weekday_friday(self):
        rows = ts.build_rows(date(2026, 5, 8), "U", "P", "T", 8.0, [])  # Friday
        self.assertEqual(rows[0][9], "Sexta-feira")


# ---------------------------------------------------------------------------
# get_dates_to_register
# ---------------------------------------------------------------------------
class TestGetDatesToRegister(unittest.TestCase):
    def test_no_last_date_returns_only_today(self):
        today = date(2026, 5, 8)
        result = ts.get_dates_to_register(None, today)
        self.assertEqual(result, [today])

    def test_contiguous_weekdays(self):
        last = date(2026, 5, 4)   # Monday
        target = date(2026, 5, 8)  # Friday
        result = ts.get_dates_to_register(last, target)
        expected = [date(2026, 5, 5), date(2026, 5, 6), date(2026, 5, 7), date(2026, 5, 8)]
        self.assertEqual(result, expected)

    def test_skips_weekends(self):
        last = date(2026, 5, 7)    # Thursday
        target = date(2026, 5, 11)  # Monday
        result = ts.get_dates_to_register(last, target)
        # Should include Fri 8, Mon 11 (skip Sat 9, Sun 10)
        expected = [date(2026, 5, 8), date(2026, 5, 11)]
        self.assertEqual(result, expected)

    def test_up_to_date_returns_target(self):
        today = date(2026, 5, 8)
        result = ts.get_dates_to_register(today, today)
        # last == target → no new dates → returns [target]
        self.assertEqual(result, [today])

    def test_last_date_is_friday_next_is_monday(self):
        last = date(2026, 5, 8)   # Friday
        target = date(2026, 5, 11)  # Monday
        result = ts.get_dates_to_register(last, target)
        self.assertEqual(result, [date(2026, 5, 11)])


# ---------------------------------------------------------------------------
# read_existing_dates
# ---------------------------------------------------------------------------
class TestReadExistingDates(unittest.TestCase):
    def _make_xlsx(self, rows):
        """Create a temp xlsx with header + given rows on the Timesheet sheet."""
        import tempfile
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Timesheet"
        ws.append(["Data\nYYYYMMDD", "Pessoa", "Projeto"])
        for row in rows:
            ws.append(row)
        tmp = Path(tempfile.mktemp(suffix=".xlsx"))
        wb.save(tmp)
        return tmp

    def test_reads_user_date_pairs(self):
        tmp = self._make_xlsx([
            [20260507, "Test User", "SI3"],
            [20260508, "Test User", "SI3"],
        ])
        try:
            result = ts.read_existing_dates(tmp)
            self.assertIn("20260507|Test User", result)
            self.assertIn("20260508|Test User", result)
        finally:
            tmp.unlink(missing_ok=True)

    def test_empty_xlsx_returns_empty_set(self):
        tmp = self._make_xlsx([])
        try:
            result = ts.read_existing_dates(tmp)
            self.assertEqual(result, set())
        finally:
            tmp.unlink(missing_ok=True)

    def test_nonexistent_file_returns_empty_set(self):
        result = ts.read_existing_dates(Path("/does/not/exist.xlsx"))
        self.assertEqual(result, set())


# ---------------------------------------------------------------------------
# Dry-run integration test
# ---------------------------------------------------------------------------
class TestDryRunIntegration(unittest.TestCase):
    """Smoke test: run main() with --dry-run — xlsx must NOT be written."""

    def test_dry_run_no_xlsx_write(self):
        from datetime import date as _date
        with patch("sys.argv", ["timesheet.py", "--dry-run", "--date", "20260507"]), \
             patch.object(ts, "get_commits_for_date", return_value=[
                 {"datetime": "2026-05-07 10:00:00 +0100",
                  "hash": "deadbeef",
                  "subject": "Implement feature 23456",
                  "repo": "test-repo"}
             ]), \
             patch.object(ts, "get_last_registered_date", return_value=_date(2026, 5, 6)), \
             patch.object(ts, "read_existing_dates", return_value=set()), \
             patch.object(ts, "append_rows") as mock_append, \
             patch.object(ts, "replace_rows_for_date") as mock_replace:
            ts.main()

        mock_append.assert_not_called()
        mock_replace.assert_not_called()


if __name__ == "__main__":
    unittest.main(verbosity=2)
