# Visible Alpha → Telegram bot

Pull live Visible Alpha numbers from Excel on your Windows PC by messaging a
Telegram bot from your phone.

```
phone (Telegram) ──► bot.py on your PC ──► Excel + Visible Alpha add-in
                 ◄── table of numbers  ◄──
```

The PC must be **on, logged in, with Excel and the Visible Alpha add-in signed in**.

## 1. Create the bot (on your phone, 2 min)
1. In Telegram, open **@BotFather** → send `/newbot` → pick a name and username.
2. Copy the token it gives you (`123456789:AA...`). Treat it like a password.

## 2. Build the template sheet in Excel
Make a workbook (e.g. `C:\Users\you\Documents\VA_Bot.xlsx`) with a sheet called **Bot**:

| | A | B | C | D |
|---|---|---|---|---|
| 1 | Ticker | `SGRO` *(the bot writes here)* | | |
| 2 | | | | |
| 3 | Metric | FY25 | FY26 | FY27 |
| 4 | EPS | `=VA formula using $B$1` | … | … |
| 5 | DPS | … | … | … |
| … | | | | |

Use the Visible Alpha add-in to insert the formulas for the metrics you want,
and make them reference `$B$1` for the ticker. Whatever sits in `A3:F40`
gets sent to your phone. Save the workbook.

## 3. Install on the PC
1. Install Python 3.11+ from python.org (tick **Add to PATH**).
2. Copy this `telegram_bot` folder to the PC.
3. Copy `.env.example` to `.env` and fill in `TELEGRAM_BOT_TOKEN` and `WORKBOOK_PATH`.
4. Double-click **`start_bot.bat`**. First run installs everything.

## 4. Lock it to you
1. On your phone, message your bot `/id`. It replies with your chat id.
2. Put it in `.env` as `ALLOWED_CHAT_IDS=123456789`, close the window, double-click `start_bot.bat` again.

Anyone else who finds the bot gets "Not authorised".

## 5. Use it
| Command | What it does |
|---|---|
| `/q SGRO` | Puts SGRO in `Bot!B1`, recalculates, waits for VA, sends `A3:F40` |
| `/range Model!A1:H30` | Sends any range from the workbook |
| `/refresh` | Full recalc of the workbook |
| `/sheets` | Lists sheet names |
| `/help` | Command list |

## How it decides the data is ready
After writing the ticker it runs a full recalc and checks every 0.5s. It accepts
when Excel says calculation is finished, no formula cell in the output range
shows an error or loading text, and two checks in a row match. Requests are
handled one at a time; a second request waits ("Queued behind another request…").
Every request is logged to `bot.log` with ticker, wait time and outcome.

Tests for this logic (no Excel needed): `python -m unittest test_completion`

## Keep it running
- Windows power settings: **Sleep → Never** while plugged in.
- To start automatically: press `Win+R`, type `shell:startup`, put a shortcut to `start_bot.bat` there.

## Troubleshooting
- **"⚠️ Partial — still loading: …":** those rows still showed an Excel error or
  loading text when `CALC_TIMEOUT` ran out. Raise `CALC_TIMEOUT`, or if a metric
  is genuinely `#N/A` for that company, remove it from the template. If your add-in
  shows a different loading message, add it to `LOADING_MARKERS` in `completion.py`.
- **"⚠️ Numbers unchanged from the previous ticker":** the add-in may not have
  fetched the new ticker. Run `/q` again.
- **Hidden Excel vs your Excel:** the bot opens its own hidden Excel and connects
  the Visible Alpha add-in there, so your own window is left alone. If the add-in
  can't load there, it falls back to your open workbook (see `bot.log`, which also
  lists every add-in found at startup).
- **Old numbers:** make sure the VA add-in is signed in on the PC; try `/refresh`.
- **Errors about the workbook:** check `WORKBOOK_PATH` and that the sheet is named `Bot`.

Check your firm's Visible Alpha licence terms before sending its data off the
workstation; the bot only goes to chat ids you allow.
