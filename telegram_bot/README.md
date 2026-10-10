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

## Keep it running
- Windows power settings: **Sleep → Never** while plugged in.
- To start automatically: press `Win+R`, type `shell:startup`, put a shortcut to `start_bot.bat` there.

## Troubleshooting
- **Cells come back as loading text / timeout:** raise `CALC_TIMEOUT` in `.env`.
  If your add-in shows a different loading message, add it to `LOADING_MARKERS` in `bot.py`.
- **Old numbers:** make sure the VA add-in is signed in on the PC; try `/refresh`.
- **Errors about the workbook:** check `WORKBOOK_PATH` and that the sheet is named `Bot`.

Check your firm's Visible Alpha licence terms before sending its data off the
workstation; the bot only goes to chat ids you allow.
