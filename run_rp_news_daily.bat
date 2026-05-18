@echo off
setlocal

cd /d "%~dp0"

where py >nul 2>nul
if %errorlevel%==0 (
    py -3 update_rp_news.py
) else (
    python update_rp_news.py
)

exit /b %errorlevel%
