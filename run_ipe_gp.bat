@echo off
title IPE/GP Portal - LJ University
cd /d "%~dp0"
echo.
echo ============================================
echo   IPE/GP Portal (NOT the conference app)
echo   Open: http://127.0.0.1:8088/
echo   (Port 8000/8001 may be used by conference app)
echo ============================================
echo.
python manage.py runserver 8088
pause
