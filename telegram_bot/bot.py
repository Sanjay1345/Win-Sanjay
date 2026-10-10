"""Telegram bot that reads live Visible Alpha data from Excel on your Windows PC.

Runs on the PC where Excel and the Visible Alpha add-in are installed and
logged in. You message the bot from your phone; it drives its own hidden Excel
instance over COM, recalculates, waits for the add-in, and sends the numbers back.

Commands
  /q <TICKER>        put TICKER into the template's input cell, recalc, send results
  /range <Sheet!A1:D10>  send any range from the workbook
  /refresh           full recalculation of the workbook
  /sheets            list sheets in the workbook
  /id                show your Telegram chat id
"""

import asyncio
import logging
import os
import time
from concurrent.futures import ThreadPoolExecutor
from contextlib import asynccontextmanager

import pythoncom
import win32com.client
from dotenv import load_dotenv
from telegram import Update
from telegram.constants import ParseMode
from telegram.ext import Application, CommandHandler, ContextTypes

from completion import (
    XL_CALC_DONE,
    XL_ERRORS,
    Poll,
    freeze,
    is_excel_error,
    to_grid,
    wait_for_completion,
)

# --- configuration (environment variables, see .env.example) -----------------
HERE = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(HERE, ".env"))
TOKEN = os.environ["TELEGRAM_BOT_TOKEN"]
ALLOWED_CHAT_IDS = {
    int(x) for x in os.environ.get("ALLOWED_CHAT_IDS", "").split(",") if x.strip()
}
WORKBOOK_PATH = os.path.abspath(os.environ["WORKBOOK_PATH"])
TEMPLATE_SHEET = os.environ.get("TEMPLATE_SHEET", "Bot")
INPUT_CELL = os.environ.get("INPUT_CELL", "B1")
OUTPUT_RANGE = os.environ.get("OUTPUT_RANGE", "A3:F40")
CALC_TIMEOUT = int(os.environ.get("CALC_TIMEOUT", "60"))

POLL_INTERVAL = 0.5
QUEUE_NOTICE_AFTER = 5
VA_ADDIN_NAME = "visible alpha"

logging.basicConfig(
    format="%(asctime)s %(levelname)s %(message)s",
    level=logging.INFO,
    handlers=[
        logging.FileHandler(os.path.join(HERE, "bot.log"), encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
log = logging.getLogger("va-bot")
# httpx logs every request URL, which contains the bot token; keep it out of bot.log.
logging.getLogger("httpx").setLevel(logging.WARNING)


# --- Excel, all COM calls on one dedicated thread ------------------------------
class Excel:
    """Owns the Excel COM objects. Every method except ``call`` runs on the
    single worker thread, so COM objects never cross threads."""

    def __init__(self):
        self.pool = ThreadPoolExecutor(
            max_workers=1, thread_name_prefix="excel", initializer=pythoncom.CoInitialize
        )
        self.app = None
        self.wb = None
        self.owns_app = False

    async def call(self, fn, *args):
        return await asyncio.get_running_loop().run_in_executor(self.pool, fn, *args)

    # ---- worker-thread methods ----
    def open(self):
        if self.wb is not None:
            return self.wb
        app = None
        try:
            app = win32com.client.DispatchEx("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            if not self._load_va_addin(app):
                raise RuntimeError("Visible Alpha add-in not loaded in hidden instance")
            self.wb = app.Workbooks.Open(WORKBOOK_PATH)
            self.app, self.owns_app = app, True
            log.info("Opened %s in dedicated hidden Excel instance", WORKBOOK_PATH)
        except Exception as e:
            log.warning("Hidden Excel instance unusable (%s); falling back to "
                        "GetObject on the open workbook - this drives your own Excel", e)
            if app is not None:
                try:
                    app.Quit()
                except Exception:
                    pass
            self.wb = win32com.client.GetObject(WORKBOOK_PATH)
            self.app, self.owns_app = self.wb.Application, False
        return self.wb

    def _load_va_addin(self, app) -> bool:
        com_addins = app.COMAddIns
        loaded = False
        for i in range(1, com_addins.Count + 1):
            a = com_addins.Item(i)
            log.info("COM add-in: %s | %s | connected=%s", a.Description, a.ProgId, a.Connect)
            if VA_ADDIN_NAME in (a.Description or "").lower():
                try:
                    a.Connect = True
                except Exception as e:
                    log.warning("Could not connect %s: %s", a.Description, e)
                loaded = loaded or bool(a.Connect)
        for i in range(1, app.AddIns.Count + 1):
            a = app.AddIns.Item(i)
            log.info("Excel add-in: %s | installed=%s", a.Name, a.Installed)
        log.info("Visible Alpha add-in loaded in hidden instance: %s", loaded)
        return loaded

    def reset(self):
        """Drop handles after a COM failure so the next call reopens."""
        self.close()

    def close(self):
        if self.owns_app and self.app is not None:
            try:
                self.wb.Close(SaveChanges=False)
            except Exception:
                pass
            try:
                self.app.Quit()
            except Exception:
                pass
        self.app = self.wb = None
        self.owns_app = False

    def _sheet(self):
        return self.open().Worksheets(TEMPLATE_SHEET)

    def prepare(self, ticker: str):
        """Snapshot output, write ticker, start full recalc.
        Returns (previous B1, previous snapshot, first row number of output range)."""
        sheet = self._sheet()
        out = sheet.Range(OUTPUT_RANGE)
        before = freeze(to_grid(out.Value))
        prev = sheet.Range(INPUT_CELL).Value
        sheet.Range(INPUT_CELL).Value = ticker
        self.app.CalculateFull()
        return prev, before, out.Row

    def recalc(self):
        self.open()
        self.app.CalculateFull()

    def poll(self) -> Poll:
        out = self._sheet().Range(OUTPUT_RANGE)
        return Poll(
            calc_done=self.app.CalculationState == XL_CALC_DONE,
            values=to_grid(out.Value),
            formulas=to_grid(out.Formula),
        )

    def read_input(self):
        return self._sheet().Range(INPUT_CELL).Value

    def read_range(self, sheet_name: str, addr: str):
        return to_grid(self.open().Worksheets(sheet_name).Range(addr).Value)

    def sheet_names(self):
        wb = self.open()
        return [wb.Worksheets(i).Name for i in range(1, wb.Worksheets.Count + 1)]


excel = Excel()
EXCEL_LOCK = asyncio.Lock()


@asynccontextmanager
async def excel_turn(update: Update):
    """Serialise Excel work; tell the user if they're waiting behind someone."""
    acquire = asyncio.ensure_future(EXCEL_LOCK.acquire())
    done, _ = await asyncio.wait({acquire}, timeout=QUEUE_NOTICE_AFTER)
    if not done:
        await update.message.reply_text("Queued behind another request…")
    await acquire
    try:
        yield
    finally:
        EXCEL_LOCK.release()


# --- formatting ----------------------------------------------------------------
def fmt(v) -> str:
    if v is None:
        return ""
    if is_excel_error(v):
        return XL_ERRORS[v]
    if isinstance(v, float):
        if abs(v) >= 1000:
            return f"{v:,.0f}"
        return f"{v:,.2f}"
    if hasattr(v, "strftime"):
        return v.strftime("%Y-%m-%d")
    return str(v)


def table_text(rows_in: list[list]) -> str:
    """Render a 2-D range as a monospaced table, dropping empty rows/columns."""
    rows = [[fmt(v) for v in r] for r in rows_in]
    rows = [r for r in rows if any(c for c in r)]
    if not rows:
        return "(empty)"
    keep = [i for i in range(len(rows[0])) if any(r[i] for r in rows)]
    rows = [[r[i] for i in keep] for r in rows]
    widths = [min(max(len(r[i]) for r in rows), 22) for i in range(len(keep))]
    lines = ["  ".join(c[:22].ljust(w) for c, w in zip(r, widths)).rstrip() for r in rows]
    return "\n".join(lines)


def escape(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


async def send_table(update: Update, title: str, body: str) -> None:
    # Telegram messages are capped at 4096 chars; split long tables.
    chunk, chunks = [], []
    for line in body.splitlines():
        if sum(len(l) + 1 for l in chunk) + len(line) > 3500:
            chunks.append("\n".join(chunk))
            chunk = []
        chunk.append(line)
    chunks.append("\n".join(chunk))
    for i, c in enumerate(chunks):
        head = f"{title}\n" if i == 0 else ""
        await update.message.reply_text(
            f"{head}<pre>{escape(c)}</pre>", parse_mode=ParseMode.HTML
        )


def row_label(values: list[list], r: int, first_row: int) -> str:
    label = values[r][0] if values[r] else None
    if isinstance(label, str) and label.strip():
        return label.strip()
    return f"row {first_row + r}"


# --- access control ------------------------------------------------------------
def allowed(update: Update) -> bool:
    return update.effective_chat.id in ALLOWED_CHAT_IDS


async def deny(update: Update) -> None:
    await update.message.reply_text(
        "Not authorised. Add this chat id to ALLOWED_CHAT_IDS on the PC: "
        f"{update.effective_chat.id}"
    )


# --- command handlers ----------------------------------------------------------
async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not allowed(update):
        return await deny(update)
    await update.message.reply_text(__doc__.split("Commands", 1)[1].strip())


async def cmd_id(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(f"Your chat id: {update.effective_chat.id}")


async def cmd_q(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not allowed(update):
        return await deny(update)
    if not context.args:
        return await update.message.reply_text("Usage: /q SGRO")
    ticker = context.args[0].upper()
    async with excel_turn(update):
        await update.message.reply_text(f"Pulling {ticker} from Visible Alpha…")
        t0 = time.monotonic()
        try:
            prev, before, first_row = await excel.call(excel.prepare, ticker)
            # Same ticker again legitimately gives identical numbers: no stale guard.
            previous = before if str(prev or "").strip().upper() != ticker else None
            res = await wait_for_completion(
                lambda: excel.call(excel.poll),
                timeout=CALC_TIMEOUT, interval=POLL_INTERVAL, previous=previous,
            )
            in_b1 = str(await excel.call(excel.read_input) or "").strip()
        except Exception as e:
            log.exception("q ticker=%s wait=%.1fs outcome=error", ticker, time.monotonic() - t0)
            await excel.call(excel.reset)
            return await update.message.reply_text(f"Error: {e}")

    notes = []
    if in_b1.upper() != ticker:
        notes.append(f"⚠️ {INPUT_CELL} shows {escape(in_b1)}, you asked for {escape(ticker)}")
    if res.stale:
        notes.append("⚠️ Numbers unchanged from the previous ticker — may be stale")
    if res.done:
        outcome = "stale" if res.stale else "done"
    else:
        outcome = "partial"
        what = ", ".join(row_label(res.values, r, first_row) for r in res.not_ready_rows)
        notes.insert(0, "⚠️ Partial — still loading: "
                     + escape(what or "Excel still calculating / values still changing"))
    log.info("q ticker=%s b1=%s wait=%.1fs outcome=%s not_ready=%s",
             ticker, in_b1, res.waited, outcome, res.not_ready_rows)
    title = "\n".join(notes + [f"<b>{escape(in_b1)}</b> ({res.waited:.0f}s)"])
    await send_table(update, title, table_text(res.values))


async def cmd_range(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not allowed(update):
        return await deny(update)
    if not context.args or "!" not in context.args[0]:
        return await update.message.reply_text("Usage: /range Sheet1!A1:D20")
    ref = " ".join(context.args)
    sheet_name, addr = ref.rsplit("!", 1)
    async with excel_turn(update):
        try:
            values = await excel.call(excel.read_range, sheet_name.strip("'"), addr)
        except Exception as e:
            log.exception("range ref=%s outcome=error", ref)
            await excel.call(excel.reset)
            return await update.message.reply_text(f"Error: {e}")
    log.info("range ref=%s outcome=done", ref)
    await send_table(update, f"<b>{escape(ref)}</b>", table_text(values))


async def cmd_refresh(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not allowed(update):
        return await deny(update)
    async with excel_turn(update):
        await update.message.reply_text("Recalculating…")
        t0 = time.monotonic()
        try:
            await excel.call(excel.recalc)
            res = await wait_for_completion(
                lambda: excel.call(excel.poll),
                timeout=CALC_TIMEOUT, interval=POLL_INTERVAL,
            )
        except Exception as e:
            log.exception("refresh wait=%.1fs outcome=error", time.monotonic() - t0)
            await excel.call(excel.reset)
            return await update.message.reply_text(f"Error: {e}")
    outcome = "done" if res.done else "partial"
    log.info("refresh wait=%.1fs outcome=%s", res.waited, outcome)
    await update.message.reply_text(
        "Done." if res.done else "⚠️ Timed out — some cells still loading."
    )


async def cmd_sheets(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not allowed(update):
        return await deny(update)
    async with excel_turn(update):
        try:
            names = await excel.call(excel.sheet_names)
        except Exception as e:
            await excel.call(excel.reset)
            return await update.message.reply_text(f"Error: {e}")
    await update.message.reply_text("\n".join(names))


async def on_startup(app: Application) -> None:
    # Open Excel now so the add-in list is logged at startup.
    try:
        await excel.call(excel.open)
    except Exception:
        log.exception("Could not open workbook at startup; will retry on first command")


async def on_shutdown(app: Application) -> None:
    await excel.call(excel.close)
    excel.pool.shutdown(wait=True)


def main() -> None:
    if not ALLOWED_CHAT_IDS:
        log.warning("ALLOWED_CHAT_IDS is empty: the bot will refuse everyone. "
                    "Message it /id to get your chat id.")
    app = (
        Application.builder()
        .token(TOKEN)
        .concurrent_updates(True)  # handlers run concurrently; EXCEL_LOCK serialises Excel
        .post_init(on_startup)
        .post_shutdown(on_shutdown)
        .build()
    )
    app.add_handler(CommandHandler(["start", "help"], cmd_start))
    app.add_handler(CommandHandler("id", cmd_id))
    app.add_handler(CommandHandler("q", cmd_q))
    app.add_handler(CommandHandler("range", cmd_range))
    app.add_handler(CommandHandler("refresh", cmd_refresh))
    app.add_handler(CommandHandler("sheets", cmd_sheets))
    log.info("Bot running. Press Ctrl+C to stop.")
    app.run_polling()


if __name__ == "__main__":
    main()
