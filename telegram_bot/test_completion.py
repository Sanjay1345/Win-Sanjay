"""Tests for completion.wait_for_completion using fake snapshots.

Run: python -m unittest test_completion  (from the telegram_bot folder)
"""

import asyncio
import unittest

from completion import (
    Poll,
    cell_not_ready,
    freeze,
    not_ready_cells,
    to_grid,
    wait_for_completion,
)

NA = -2146826246
NAME = -2146826259
F = [["EPS", "=VA(1)"], ["DPS", "=VA(2)"]]  # column A labels, column B formulas


class FakeExcel:
    """Returns the scripted snapshots one per poll, repeating the last one."""

    def __init__(self, snapshots):
        self.snapshots = snapshots
        self.i = 0
        self.now = 0.0

    async def poll(self):
        s = self.snapshots[min(self.i, len(self.snapshots) - 1)]
        self.i += 1
        return s

    async def sleep(self, dt):
        self.now += dt

    def clock(self):
        return self.now


def run(fake, timeout=5.0, previous=None):
    return asyncio.run(
        wait_for_completion(
            fake.poll, timeout=timeout, previous=previous,
            sleep=fake.sleep, clock=fake.clock,
        )
    )


def snap(eps, dps, calc_done=True, formulas=F):
    return Poll(calc_done, [["EPS", eps], ["DPS", dps]], formulas)


class CellChecks(unittest.TestCase):
    def test_not_ready_values(self):
        for v in (NA, NAME, -2146826273, "#GETTING_DATA", "Loading...", "Retrieving",
                  "BUSY", "#N/A Requesting Data"):
            self.assertTrue(cell_not_ready(v), v)
        for v in (12.5, 0, "EPS", None, True, 2042):
            self.assertFalse(cell_not_ready(v), v)

    def test_only_formula_cells_count(self):
        values = [["Loading", NA]]
        self.assertEqual(not_ready_cells(values, [["Loading", "=VA()"]]), [(0, 1)])
        self.assertEqual(not_ready_cells(values, [["Loading", ""]]), [])

    def test_to_grid(self):
        self.assertEqual(to_grid(5), [[5]])
        self.assertEqual(to_grid((1, 2)), [[1, 2]])
        self.assertEqual(to_grid(((1, 2), (3, 4))), [[1, 2], [3, 4]])


class WaitForCompletion(unittest.TestCase):
    def test_done_after_loading_then_two_stable_polls(self):
        fake = FakeExcel([
            snap("Loading", NA, calc_done=False),
            snap(1.2, NA),
            snap(1.2, 0.8),
            snap(1.2, 0.8),
        ])
        r = run(fake)
        self.assertTrue(r.done)
        self.assertFalse(r.stale)
        self.assertEqual(r.values[1][1], 0.8)
        self.assertEqual(fake.i, 4)

    def test_single_clean_poll_is_not_enough(self):
        fake = FakeExcel([snap(1.2, 0.8), snap(1.3, 0.8), snap(1.3, 0.8)])
        r = run(fake)
        self.assertTrue(r.done)
        self.assertEqual(r.values[0][1], 1.3)
        self.assertEqual(fake.i, 3)

    def test_calc_state_must_be_done(self):
        fake = FakeExcel([snap(1.2, 0.8, calc_done=False)])
        r = run(fake, timeout=2)
        self.assertFalse(r.done)
        self.assertEqual(r.not_ready_rows, [])

    def test_timeout_reports_not_ready_rows(self):
        fake = FakeExcel([snap(1.2, "#GETTING_DATA")])
        r = run(fake, timeout=3)
        self.assertFalse(r.done)
        self.assertEqual(r.not_ready_rows, [1])
        self.assertGreaterEqual(r.waited, 3)

    def test_error_in_non_formula_cell_ignored(self):
        formulas = [["EPS", "=VA(1)"], ["DPS", NA]]  # DPS typed in, not a formula
        fake = FakeExcel([snap(1.2, NA, formulas=formulas)])
        self.assertTrue(run(fake).done)

    def test_values_still_changing_times_out(self):
        fake = FakeExcel([snap(float(i), 1.0) for i in range(100)])
        r = run(fake, timeout=3)
        self.assertFalse(r.done)

    def test_stale_snapshot_waits_then_accepts_new_values(self):
        old = freeze([["EPS", 9.9], ["DPS", 9.9]])
        # Add-in hasn't started: old numbers for 3s, then new numbers arrive
        # (inside the extra cycle, which would otherwise end at ~5.5s).
        polls = [snap(9.9, 9.9)] * 6 + [snap(1.2, 0.8)] * 3
        fake = FakeExcel(polls)
        r = run(fake, timeout=5, previous=old)
        self.assertTrue(r.done)
        self.assertFalse(r.stale)
        self.assertEqual(r.values[0][1], 1.2)

    def test_stale_snapshot_accepted_after_extra_cycle(self):
        old = freeze([["EPS", 9.9], ["DPS", 9.9]])
        fake = FakeExcel([snap(9.9, 9.9)])
        r = run(fake, timeout=5, previous=old)
        self.assertTrue(r.done)
        self.assertTrue(r.stale)
        self.assertGreaterEqual(r.waited, 5)  # waited at least one extra cycle


if __name__ == "__main__":
    unittest.main()
