"""Telegram bot that reads live Visible Alpha data from Excel on your Windows PC.

Runs on the PC where Excel and the Visible Alpha add-in are installed and
logged in. You message the bot from your phone; it drives Excel through
xlwings, recalculates, and sends the numbers back.

Commands
  /q <TICKER>        put TICKER into the template's input cell, recalc, send results
  /range <Sheet!A1:D10>  send any range from the workbook
  /refresh           full recalculation of the workbook
  /sheets            list sheets in the workbook
  /id                show your Telegram chat id
"""

import logging
import os
import time

import xlwings as xw
from dotenv import load_dotenv
from telegram import Update
from telegram.constants import ParseMode
from telegram.ext import Application, CommandHandler, ContextTypes

# --- configuration (environment variables, see .env.example) -----------------
load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))
TOKEN = os.environ["TELEGRAM_BOT_TOKEN"]
ALLOWED_CHAT_IDS = {
    int(x) for x in os.environ.get("ALLOWED_CHAT_IDS", "").split(",") if x.strip()
}
WORKBOOK_PATH = os.environ["WORKBOOK_PATH"]
TEMPLATE_SHEET = os.environ.get("TEMPLATE_SHEET", "Bot")
INPUT_CELL = os.environ.get("INPUT_CELL", "B1")
OUTPUT_RANGE = os.environ.get("OUTPUT_RANGE", "A3:F40")
CALC_TIMEOUT = int(os.environ.get("CALC_TIMEOUT", "60"))

# Text that add-ins show in a cell while data is still loading.
LOADING_MARKERS = ("#GETTING_DATA", "LOADING", "#N/A REQUESTING", "RETRIEVING", "#BUSY")
XL_CALC_DONE = 0

logging.basicConfig(format="%(asctime)s %(levelname)s %(message)s", level=logging.INFO)
log = logging.getLogger("va-bot")


# --- Excel helpers -------------------------------------------------------------
def get_book() -> xw.Book:
    """Return the workbook, attaching to it if already open in Excel."""
    name = os.path.basename(WORKBOOK_PATH)
    for app in xw.apps:
        for book in app.books:
            if book.name.lower() == name.lower():
                return book
    return xw.Book(WORKBOOK_PATH)


def is_loading(values) -> bool:
    flat = values if isinstance(values, list) else [values]
    for row in flat:
        for v in row if isinstance(row, list) else [row]:
            if isinstance(v, str) and any(m in v.upper() for m in LOADING_MARKERS):
                return True
    return False


def recalc_and_wait(book: xw.Book, watch: xw.Range | None = None) -> bool:
    """Full recalc, then wait until Excel and the add-in have finished."""
    app = book.app
    app.api.CalculateFull()
    deadline = time.time() + CALC_TIMEOUT
    while time.time() < deadline:
        done = app.api.CalculationState == XL_CALC_DONE
        if done and (watch is None or not is_loading(watch.value)):
            return True
        time.sleep(1)
    return False


def fmt(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float):
        if abs(v) >= 1000:
            return f"{v:,.0f}"
        return f"{v:,.2f}"
    return str(v)


def table_text(values) -> str:
    """Render a 2-D range as a monospaced table, dropping empty rows/columns."""
    if not isinstance(values, list):
        values = [[values]]
    elif values and not isinstance(values[0], list):
        values = [values]
    rows = [[fmt(v) for v in r] for r in values]
    rows = [r for r in rows if any(c for c in r)]
    if not rows:
        return "(empty)"
    keep = [i for i in range(len(rows[0])) if any(r[i] for r in rows)]
    rows = [[r[i] for i in keep] for r in rows]
    widths = [min(max(len(r[i]) for r in rows), 22) for i in range(len(keep))]
    lines = ["  ".join(c[:22].ljust(w) for c, w in zip(r, widths)).rstrip() for r in rows]
    return "\n".join(lines)


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
        head = f"<b>{title}</b>\n" if i == 0 else ""
        await update.message.reply_text(
            f"{head}<pre>{escape(c)}</pre>", parse_mode=ParseMode.HTML
        )


def escape(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


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
    await update.message.reply_text(f"Pulling {ticker} from Visible Alpha…")
    try:
        book = get_book()
        sheet = book.sheets[TEMPLATE_SHEET]
        sheet.range(INPUT_CELL).value = ticker
        out = sheet.range(OUTPUT_RANGE)
        ok = recalc_and_wait(book, out)
        body = table_text(out.value)
        title = ticker if ok else f"{ticker} (timed out – some cells may still be loading)"
        await send_table(update, title, body)
    except Exception as e:
        log.exception("/q failed")
        await update.message.reply_text(f"Error: {e}")


async def cmd_range(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not allowed(update):
        return await deny(update)
    if not context.args or "!" not in context.args[0]:
        return await update.message.reply_text("Usage: /range Sheet1!A1:D20")
    ref = " ".join(context.args)
    sheet_name, addr = ref.rsplit("!", 1)
    try:
        book = get_book()
        rng = book.sheets[sheet_name.strip("'")].range(addr)
        await send_table(update, ref, table_text(rng.value))
    except Exception as e:
        log.exception("/range failed")
        await update.message.reply_text(f"Error: {e}")


async def cmd_refresh(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not allowed(update):
        return await deny(update)
    await update.message.reply_text("Recalculating…")
    try:
        ok = recalc_and_wait(get_book())
        await update.message.reply_text("Done." if ok else "Timed out waiting for Excel.")
    except Exception as e:
        log.exception("/refresh failed")
        await update.message.reply_text(f"Error: {e}")


async def cmd_sheets(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not allowed(update):
        return await deny(update)
    try:
        names = [s.name for s in get_book().sheets]
        await update.message.reply_text("\n".join(names))
    except Exception as e:
        await update.message.reply_text(f"Error: {e}")


def main() -> None:
    if not ALLOWED_CHAT_IDS:
        log.warning("ALLOWED_CHAT_IDS is empty: the bot will refuse everyone. "
                    "Message it /id to get your chat id.")
    app = Application.builder().token(TOKEN).build()
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
