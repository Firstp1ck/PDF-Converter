@echo off
:: Überprüfen, ob die Pakete bereits installiert sind

python -m pip list | findstr /C:"pdf2docx" > nul

:: Wenn das Paket nicht gefunden wurde, führe die Installation aus
if errorlevel 1 (
    echo Paket 'pdf2docx' wird installiert...

    python -m pip install --trusted-host pypi.org --trusted-host files.pythonhosted.org pdf2docx

) else (
    echo Pakete sind bereits installiert.
)

:: Führen Sie Ihr Python-Skript aus. Ersetzen Sie 'Main.py' durch den Namen Ihres Skripts.
python Main.py

echo.
echo Skript wurde ausgefuehrt. Fenster schliessen oder weitere Befehle eingeben.
pause

