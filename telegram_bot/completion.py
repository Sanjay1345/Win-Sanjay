"""Decide when Visible Alpha formulas in the output range have finished loading.

Pure logic with no Excel or COM imports, so it can be tested with fake snapshots.
"""

import asyncio
import time
from dataclasses import dataclass, field
from typing import Awaitable, Callable

# Excel error values as pywin32 returns them from Range.Value.
XL_ERRORS = {
    -2146826288: "#NULL!",
    -2146826281: "#DIV/0!",
    -2146826273: "#VALUE!",
    -2146826265: "#REF!",
    -2146826259: "#NAME?",
    -2146826252: "#NUM!",
    -2146826246: "#N/A",
    -2146826245: "#GETTING_DATA",
}

# Text that add-ins show in a cell while data is still loading.
LOADING_MARKERS = ("#GETTING_DATA", "LOADING", "#N/A REQUESTING", "RETRIEVING", "#BUSY")
LOADING_SUBSTRINGS = ("loading", "retrieving", "#getting", "busy")

XL_CALC_DONE = 0


def is_excel_error(v) -> bool:
    return isinstance(v, int) and not isinstance(v, bool) and v in XL_ERRORS


def cell_not_ready(v) -> bool:
    if is_excel_error(v):
        return True
    if isinstance(v, str):
        low = v.lower()
        return any(m.lower() in low for m in LOADING_MARKERS) or any(
            s in low for s in LOADING_SUBSTRINGS
        )
    return False


def to_grid(v) -> list[list]:
    """Normalise a COM Range.Value / Range.Formula result to a list of rows."""
    if isinstance(v, (tuple, list)):
        if v and isinstance(v[0], (tuple, list)):
            return [list(r) for r in v]
        return [list(v)]
    return [[v]]


def freeze(grid: list[list]) -> tuple:
    return tuple(tuple(r) for r in grid)


def not_ready_cells(values: list[list], formulas: list[list]) -> list[tuple[int, int]]:
    """(row, col) of formula cells whose value is an error or loading text."""
    bad = []
    for r, (vrow, frow) in enumerate(zip(values, formulas)):
        for c, (v, f) in enumerate(zip(vrow, frow)):
            if isinstance(f, str) and f.startswith("=") and cell_not_ready(v):
                bad.append((r, c))
    return bad


@dataclass
class Poll:
    calc_done: bool
    values: list[list]
    formulas: list[list]


@dataclass
class Result:
    done: bool                 # False means partial (timed out)
    values: list[list]
    waited: float
    not_ready_rows: list[int] = field(default_factory=list)  # 0-based within range
    stale: bool = False        # identical to the previous ticker's numbers


async def wait_for_completion(
    poll: Callable[[], Awaitable[Poll]],
    *,
    timeout: float,
    interval: float = 0.5,
    previous: tuple | None = None,
    sleep: Callable[[float], Awaitable[None]] = asyncio.sleep,
    clock: Callable[[], float] = time.monotonic,
) -> Result:
    """Poll until calc is done, no formula cell is loading/erroring, and two
    consecutive snapshots match.

    If ``previous`` (the snapshot from before the ticker changed) is given and
    the settled snapshot equals it, keep waiting one more full ``timeout`` in
    case the add-in has not started fetching yet; accept it after that with
    ``stale=True``.
    """
    start = clock()
    deadline = start + timeout
    extended = False
    last = None
    while True:
        p = await poll()
        snap = freeze(p.values)
        bad = not_ready_cells(p.values, p.formulas)
        stable = last is not None and snap == last
        last = snap
        settled = p.calc_done and not bad and stable
        same_as_previous = previous is not None and snap == previous

        if settled and not same_as_previous:
            return Result(True, p.values, clock() - start)
        if settled and same_as_previous and not extended:
            extended = True
            deadline = clock() + timeout

        if clock() >= deadline:
            if settled:  # only reachable when same_as_previous after the extra cycle
                return Result(True, p.values, clock() - start, stale=True)
            rows = sorted({r for r, _ in bad})
            return Result(False, p.values, clock() - start, rows)
        await sleep(interval)
