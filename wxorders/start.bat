@echo off
setlocal enabledelayedexpansion

REM Switch to script directory
cd /d %~dp0

if not exist .venv (
	python -m venv .venv
)

call .venv\Scripts\activate

pip install --upgrade pip >nul 2>&1
pip install -r requirements.txt
if exist requirements.win.txt (
	pip install -r requirements.win.txt
)

set FLASK_APP=app.py
start "" http://127.0.0.1:5000/
python app.py