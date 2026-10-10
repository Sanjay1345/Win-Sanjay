@echo off
cd /d "%~dp0"
if not exist .venv (
  py -m venv .venv
  .venv\Scripts\pip install -r requirements.txt
)
.venv\Scripts\python bot.py
pause
