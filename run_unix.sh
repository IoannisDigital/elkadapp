#!/usr/bin/env bash
# Αυτόνομη εκκίνηση της γραφικής εφαρμογής ΓΕΜΗ σε macOS / Linux.
# Τρέξε:  ./run_unix.sh   (ή διπλό κλικ αν το επιτρέπει το σύστημα)
set -e
cd "$(dirname "$0")"

PY="${PYTHON:-python3}"
if ! command -v "$PY" >/dev/null 2>&1; then
    echo "Δεν βρέθηκε python3. Εγκατέστησε Python 3.11+."
    exit 1
fi

echo "Εγκατάσταση/έλεγχος βιβλιοθηκών..."
"$PY" -m pip install --quiet --disable-pip-version-check -r requirements.txt

echo "Εκκίνηση εφαρμογής..."
exec "$PY" gemi_app.py
