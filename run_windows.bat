@echo off
REM Αυτόνομη εκκίνηση της γραφικής εφαρμογής ΓΕΜΗ στα Windows.
REM Διπλό κλικ σε αυτό το αρχείο.

cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo Δεν βρέθηκε Python. Εγκατέστησε το Python 3.11+ από https://www.python.org/downloads/
    echo (Στην εγκατάσταση, τσέκαρε "Add Python to PATH".)
    pause
    exit /b 1
)

echo Εγκατάσταση/έλεγχος βιβλιοθηκών...
python -m pip install --quiet --disable-pip-version-check -r requirements.txt

echo Εκκίνηση εφαρμογής...
python gemi_app.py

pause
