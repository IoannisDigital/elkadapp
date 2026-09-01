# -*- coding: utf-8 -*-
"""
ΓΕΜΗ OpenData -> Excel ανά ΚΑΔ (ένα φύλλο/tab ανά Νομό) — CLI.

Για γραφική εφαρμογή (χωρίς επεξεργασία κώδικα) τρέξε:  python gemi_app.py

Χρήση CLI:
    python gemi_pull.py search ΚΟΜΜΩΤΗΡ    # βρες ΚΑΔ με λέξη-κλειδί
    python gemi_pull.py                     # τρέξε την εξαγωγή (ρυθμίσεις πιο κάτω)
"""
import os
import sys

import gemi_core as core

# ================== ΡΥΘΜΙΣΕΙΣ ΠΟΥ ΑΛΛΑΖΕΙΣ ΕΣΥ ==================
# 1) Ποιους ΚΑΔ θέλεις (βάλε όσους θες — παράγεται 1 xlsx ανά ΚΑΔ):
KADS = ["47730000", "47750000"]  # π.χ. φαρμακεία, καλλυντικά

# 2) Ποιοι Νομοί; "ALL" για όλη την Ελλάδα, ή λίστα με ονόματα/κωδικούς:
PREFECTURES = "ALL"  # π.χ. ["ΘΕΣΣΑΛΟΝΙΚΗΣ", "ΑΤΤΙΚΗΣ"] ή ["19", "5"]

# 3) Κράτα μόνο όσες έχουν τον ΚΑΔ ως ΚΥΡΙΑ δραστηριότητα (καθαρή λίστα):
PRIMARY_ONLY = True

# 4) Αλυσίδες ενός brand που ΔΕΝ θέλεις (case-insensitive, ελέγχει "περιέχει"):
EXCLUDE = list(core.DEFAULT_EXCLUDE)
# ================================================================


def search_kad(keyword):
    hits = core.search_activities(keyword)
    print(f"Βρέθηκαν {len(hits)} ΚΑΔ για '{keyword}':")
    for a in hits:
        print(f"  {a.get('id')}  {a.get('descr')}")


def main():
    if len(sys.argv) >= 3 and sys.argv[1] == "search":
        search_kad(" ".join(sys.argv[2:]))
        return

    os.makedirs("output", exist_ok=True)
    for kad in KADS:
        print("== ΚΑΔ", kad, "| Νομοί:", PREFECTURES)
        path, rows = core.export_kad(
            kad, prefectures=PREFECTURES, out_dir="output",
            primary_only=PRIMARY_ONLY, exclude=EXCLUDE, progress=print)
        print("  -> αποθηκεύτηκε:", path, "με", len(rows), "εγγραφές")


if __name__ == "__main__":
    main()
