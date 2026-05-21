@echo off
cd /d "%~dp0"
python -m pip install requests openpyxl --quiet
python star_query_tool.py
pause
