@echo off
rem Usage: daily VOCAB IDIOMS OWS ITEM DATE [extra options]
rem   daily 38 41-50 1-10 8 24/9/2026
rem   daily 38 41-50 1-10 8 24/9/2026 --only quiz
rem   daily 38 41-50 1-10 8 24/9/2026 --print
cd /d "%~dp0"
"%~dp0venv\Scripts\python.exe" daily.py --vocab %1 --idioms %2 --ows %3 --item %4 --date %5 %6 %7 %8 %9
