# -*- coding: utf-8 -*-
"""
ΓΕΜΗ OpenData — Γραφική Εφαρμογή (αυτόνομη).

Άνοιγμα:  python gemi_app.py   (ή διπλό κλικ στα run scripts)

Δυνατότητες:
  • Αναζήτηση ΚΑΔ με λέξη-κλειδί και προσθήκη στη λίστα εξαγωγής.
  • Επιλογή Νομών: ΟΛΗ η Ελλάδα ή συγκεκριμένοι (πολλαπλή επιλογή).
  • Ρυθμίσεις καθαρισμού (μόνο Κύρια δραστηριότητα, λίστα αποκλεισμού brand).
  • Εξαγωγή: ένα .xlsx ανά ΚΑΔ στον φάκελο output/, με Σύνοψη + tab ανά Νομό.
  • Ζωντανή πρόοδος και δυνατότητα ακύρωσης.

Δεν χρειάζεται να πειράξεις καθόλου κώδικα — όλα γίνονται από το παράθυρο.
"""
import os
import queue
import threading
import webbrowser

try:
    import tkinter as tk
    from tkinter import ttk, messagebox
except ModuleNotFoundError:
    import sys
    sys.stderr.write(
        "\nΤο tkinter (γραφικό περιβάλλον) δεν είναι εγκατεστημένο.\n"
        "  • Windows/macOS: επανεγκατέστησε Python από python.org "
        "(περιλαμβάνει tkinter).\n"
        "  • Linux (Debian/Ubuntu): sudo apt install python3-tk\n"
        "  • Linux (Fedora): sudo dnf install python3-tkinter\n\n"
        "Εναλλακτικά, χρησιμοποίησε το CLI:  python gemi_pull.py\n\n")
    sys.exit(1)

import gemi_core as core


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("ΓΕΜΗ OpenData — Εξαγωγή σε Excel ανά ΚΑΔ")
        self.geometry("1000x680")
        self.minsize(880, 600)

        self.msg_q = queue.Queue()
        self.worker = None
        self.stop_flag = threading.Event()
        self.selected_kads = {}  # id -> descr

        self._build_ui()
        self.after(120, self._drain_queue)

    # ---------------- UI ----------------
    def _build_ui(self):
        pad = dict(padx=8, pady=6)
        top = ttk.Frame(self)
        top.pack(fill="x", **pad)

        ttk.Label(top, text="API key:").pack(side="left")
        self.api_var = tk.StringVar(value=core.get_api_key())
        ttk.Entry(top, textvariable=self.api_var, width=42, show="•").pack(side="left", padx=6)

        self.primary_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(top, text="Μόνο Κύρια δραστηριότητα (καθαρή λίστα)",
                        variable=self.primary_var).pack(side="left", padx=12)

        body = ttk.Panedwindow(self, orient="horizontal")
        body.pack(fill="both", expand=True, **pad)

        # ----- Αριστερά: αναζήτηση & επιλογή ΚΑΔ -----
        left = ttk.Labelframe(body, text="1) ΚΑΔ")
        body.add(left, weight=3)

        srow = ttk.Frame(left)
        srow.pack(fill="x", padx=6, pady=6)
        ttk.Label(srow, text="Αναζήτηση:").pack(side="left")
        self.search_var = tk.StringVar()
        ent = ttk.Entry(srow, textvariable=self.search_var)
        ent.pack(side="left", fill="x", expand=True, padx=6)
        ent.bind("<Return>", lambda e: self.do_search())
        ttk.Button(srow, text="Ψάξε", command=self.do_search).pack(side="left")

        ttk.Label(left, text="Αποτελέσματα (διπλό κλικ ή «Προσθήκη» για επιλογή):").pack(
            anchor="w", padx=6)
        res_wrap = ttk.Frame(left)
        res_wrap.pack(fill="both", expand=True, padx=6, pady=(0, 6))
        self.results = ttk.Treeview(res_wrap, columns=("id", "descr"), show="headings",
                                    height=10, selectmode="extended")
        self.results.heading("id", text="ΚΑΔ")
        self.results.heading("descr", text="Περιγραφή")
        self.results.column("id", width=100, anchor="w", stretch=False)
        self.results.column("descr", width=380, anchor="w")
        rsb = ttk.Scrollbar(res_wrap, orient="vertical", command=self.results.yview)
        self.results.configure(yscrollcommand=rsb.set)
        self.results.pack(side="left", fill="both", expand=True)
        rsb.pack(side="right", fill="y")
        self.results.bind("<Double-1>", lambda e: self.add_selected())

        arow = ttk.Frame(left)
        arow.pack(fill="x", padx=6, pady=(0, 6))
        ttk.Button(arow, text="➕ Προσθήκη στην εξαγωγή", command=self.add_selected).pack(
            side="left")
        ttk.Button(arow, text="Προσθήκη με κωδικό…", command=self.add_manual).pack(
            side="left", padx=6)

        ttk.Label(left, text="Επιλεγμένοι ΚΑΔ προς εξαγωγή (1 Excel ανά ΚΑΔ):").pack(
            anchor="w", padx=6)
        sel_wrap = ttk.Frame(left)
        sel_wrap.pack(fill="both", expand=True, padx=6, pady=(0, 6))
        self.chosen = ttk.Treeview(sel_wrap, columns=("id", "descr"), show="headings",
                                  height=6, selectmode="extended")
        self.chosen.heading("id", text="ΚΑΔ")
        self.chosen.heading("descr", text="Περιγραφή")
        self.chosen.column("id", width=100, anchor="w", stretch=False)
        self.chosen.column("descr", width=380, anchor="w")
        csb = ttk.Scrollbar(sel_wrap, orient="vertical", command=self.chosen.yview)
        self.chosen.configure(yscrollcommand=csb.set)
        self.chosen.pack(side="left", fill="both", expand=True)
        csb.pack(side="right", fill="y")
        ttk.Button(left, text="➖ Αφαίρεση επιλεγμένου", command=self.remove_selected).pack(
            anchor="w", padx=6, pady=(0, 6))

        # ----- Δεξιά: νομοί -----
        right = ttk.Labelframe(body, text="2) Νομοί")
        body.add(right, weight=2)
        self.all_pref_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(right, text="ΟΛΗ η Ελλάδα (όλοι οι νομοί)",
                        variable=self.all_pref_var, command=self._toggle_all_pref).pack(
            anchor="w", padx=8, pady=6)
        ttk.Label(right, text="ή διάλεξε συγκεκριμένους (Ctrl/Shift για πολλαπλή):").pack(
            anchor="w", padx=8)
        pf_wrap = ttk.Frame(right)
        pf_wrap.pack(fill="both", expand=True, padx=8, pady=(0, 6))
        self.pref_list = tk.Listbox(pf_wrap, selectmode="extended", exportselection=False)
        for name in core.PREF_NAMES:
            self.pref_list.insert("end", name)
        psb = ttk.Scrollbar(pf_wrap, orient="vertical", command=self.pref_list.yview)
        self.pref_list.configure(yscrollcommand=psb.set)
        self.pref_list.pack(side="left", fill="both", expand=True)
        psb.pack(side="right", fill="y")
        self.pref_list.configure(state="disabled")

        # ----- Κάτω: ενέργειες & log -----
        bottom = ttk.Frame(self)
        bottom.pack(fill="both", expand=False, **pad)

        brow = ttk.Frame(bottom)
        brow.pack(fill="x")
        self.run_btn = ttk.Button(brow, text="▶ Εξαγωγή σε Excel", command=self.start_export)
        self.run_btn.pack(side="left")
        self.stop_btn = ttk.Button(brow, text="■ Ακύρωση", command=self.cancel_export,
                                  state="disabled")
        self.stop_btn.pack(side="left", padx=6)
        ttk.Button(brow, text="📂 Άνοιγμα φακέλου output", command=self.open_output).pack(
            side="left", padx=6)
        self.status_var = tk.StringVar(value="Έτοιμο.")
        ttk.Label(brow, textvariable=self.status_var).pack(side="right")

        self.log = tk.Text(bottom, height=11, wrap="word", state="disabled")
        self.log.pack(fill="both", expand=True, pady=(6, 0))

    # ---------------- ενέργειες ----------------
    def _toggle_all_pref(self):
        self.pref_list.configure(state="disabled" if self.all_pref_var.get() else "normal")

    def log_msg(self, msg):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def do_search(self):
        kw = self.search_var.get().strip()
        if not kw:
            return
        self.status_var.set(f"Αναζήτηση «{kw}»…")
        self.results.delete(*self.results.get_children())

        def work():
            try:
                hits = core.search_activities(kw, api_key=self.api_var.get().strip())
                self.msg_q.put(("search_done", hits))
            except Exception as e:
                self.msg_q.put(("error", f"Σφάλμα αναζήτησης: {e}"))

        threading.Thread(target=work, daemon=True).start()

    def _fill_results(self, hits):
        for a in hits:
            self.results.insert("", "end", values=(a.get("id"), a.get("descr")))
        self.status_var.set(f"Βρέθηκαν {len(hits)} ΚΑΔ.")
        if not hits:
            self.log_msg("Δεν βρέθηκαν ΚΑΔ για αυτή τη λέξη.")

    def add_selected(self):
        for item in self.results.selection():
            kid, descr = self.results.item(item, "values")
            self._add_kad(str(kid), descr)

    def add_manual(self):
        win = tk.Toplevel(self)
        win.title("Προσθήκη ΚΑΔ με κωδικό")
        win.transient(self)
        win.grab_set()
        ttk.Label(win, text="8ψήφιος ΚΑΔ:").pack(padx=12, pady=(12, 4))
        var = tk.StringVar()
        e = ttk.Entry(win, textvariable=var, width=20)
        e.pack(padx=12)
        e.focus_set()

        def ok():
            code = var.get().strip()
            if code:
                self._add_kad(code, "(χειροκίνητη προσθήκη)")
            win.destroy()

        ttk.Button(win, text="Προσθήκη", command=ok).pack(pady=12)
        win.bind("<Return>", lambda ev: ok())

    def _add_kad(self, kid, descr):
        if kid in self.selected_kads:
            return
        self.selected_kads[kid] = descr
        self.chosen.insert("", "end", iid=kid, values=(kid, descr))

    def remove_selected(self):
        for item in self.chosen.selection():
            self.selected_kads.pop(item, None)
            self.chosen.delete(item)

    def _chosen_prefectures(self):
        if self.all_pref_var.get():
            return "ALL"
        sel = [self.pref_list.get(i) for i in self.pref_list.curselection()]
        return sel or "ALL"

    def start_export(self):
        if self.worker and self.worker.is_alive():
            return
        kads = list(self.selected_kads.keys())
        if not kads:
            messagebox.showwarning("Προσοχή", "Δεν έχεις επιλέξει κανέναν ΚΑΔ.")
            return
        prefs = self._chosen_prefectures()
        api_key = self.api_var.get().strip()
        primary = self.primary_var.get()

        self.stop_flag.clear()
        self.run_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.log_msg("=" * 60)
        self.log_msg(f"Έναρξη εξαγωγής: {len(kads)} ΚΑΔ | Νομοί: "
                     f"{'ΟΛΗ η Ελλάδα' if prefs == 'ALL' else ', '.join(prefs)}")
        self.log_msg("Μεγάλοι ΚΑΔ σε όλη την Ελλάδα μπορεί να θέλουν 40–90 λεπτά "
                     "(όριο ~8 αιτήματα/λεπτό). Άφησέ το να τρέξει.")

        def prog(msg):
            self.msg_q.put(("log", msg))

        def should_stop():
            return self.stop_flag.is_set()

        def work():
            done = []
            for kad in kads:
                if self.stop_flag.is_set():
                    break
                self.msg_q.put(("status", f"ΚΑΔ {kad}…"))
                self.msg_q.put(("log", f"\n== ΚΑΔ {kad} =="))
                try:
                    path, rows = core.export_kad(
                        kad, prefectures=prefs, out_dir="output", api_key=api_key,
                        primary_only=primary, progress=prog, should_stop=should_stop)
                    self.msg_q.put(("kad_done", (kad, path, rows)))
                    done.append((kad, path, rows))
                except Exception as e:
                    self.msg_q.put(("log", f"  ΣΦΑΛΜΑ στον ΚΑΔ {kad}: {e}"))
            self.msg_q.put(("all_done", done))

        self.worker = threading.Thread(target=work, daemon=True)
        self.worker.start()

    def cancel_export(self):
        self.stop_flag.set()
        self.status_var.set("Ακύρωση…")

    def open_output(self):
        path = os.path.abspath("output")
        os.makedirs(path, exist_ok=True)
        try:
            webbrowser.open(f"file://{path}")
        except Exception:
            messagebox.showinfo("output", path)

    # ---------------- ουρά μηνυμάτων από threads ----------------
    def _drain_queue(self):
        try:
            while True:
                kind, payload = self.msg_q.get_nowait()
                if kind == "search_done":
                    self._fill_results(payload)
                elif kind == "log":
                    self.log_msg(payload)
                elif kind == "status":
                    self.status_var.set(payload)
                elif kind == "error":
                    self.log_msg(payload)
                    self.status_var.set("Σφάλμα.")
                elif kind == "kad_done":
                    kad, path, rows = payload
                    self.log_msg(f"  -> Αποθηκεύτηκε: {path}  ({len(rows)} εγγραφές)")
                    self.log_msg("  Σύνοψη ανά Νομό:")
                    for name, cnt in core.summary(rows):
                        self.log_msg(f"     {name}: {cnt}")
                elif kind == "all_done":
                    self.run_btn.configure(state="normal")
                    self.stop_btn.configure(state="disabled")
                    n = len(payload)
                    self.status_var.set(f"Ολοκληρώθηκε: {n} αρχείο(α) στο output/.")
                    self.log_msg(f"\nΟλοκληρώθηκε. Δημιουργήθηκαν {n} αρχείο(α) στον "
                                 f"φάκελο output/.")
        except queue.Empty:
            pass
        self.after(120, self._drain_queue)


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
