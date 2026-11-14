@echo off
chcp 65001 >nul
title BewerbungsTerminal - Portable Edition
cd /d "%~dp0"

echo ====================================================================
echo   BEWERBUNGS-TERMINAL (PORTABLE EDITION) WIRD GESTARTET...
echo ====================================================================
echo.
echo   Alle Dateien werden lokal in diesem Ordner gespeichert:
echo   - Muster-Word-Layouts (.docx) : %~dp0layouts
echo   - Fertige Bewerbungen (.docx) : %~dp0output
echo   - Gespeicherte Daten & Links  : %~dp0data
echo.

:: Starte die App im modernen App-Fenster (Microsoft Edge App-Mode falls vorhanden, sonst Standardbrowser)
start "" cmd /c "timeout /t 1 /nobreak >nul & start msedge --app=http://127.0.0.1:8765 2>nul || start http://127.0.0.1:8765"

python app.py --no-browser
pause
