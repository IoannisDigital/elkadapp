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
import sys
import queue
import threading
import webbrowser

try:
    import tkinter as tk
    from tkinter import ttk, messagebox
except ModuleNotFoundError:
    sys.stderr.write(
        "\nΤο tkinter (γραφικό περιβάλλον) δεν είναι εγκατεστημένο.\n"
        "  • Windows/macOS: επανεγκατέστησε Python από python.org "
        "(περιλαμβάνει tkinter).\n"
        "  • Linux (Debian/Ubuntu): sudo apt install python3-tk\n"
        "  • Linux (Fedora): sudo dnf install python3-tkinter\n\n"
        "Εναλλακτικά, χρησιμοποίησε το CLI:  python gemi_pull.py\n\n")
    sys.exit(1)

import gemi_core as core


def resource_path(rel):
    """Διαδρομή πόρου που δουλεύει και σε PyInstaller (.exe) και σε πηγαίο κώδικα."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


# ---------------- Οπτική ταυτότητα (DMS Hellas) ----------------
BRAND = "DMS Hellas"
COMPANY = "Digital Marketing Systems Hellas Ltd"
APP_NAME = "ΓΕΜΗ Data Extractor"
MOTTO = "Δημόσια δεδομένα ΓΕΜΗ, έτοιμα για B2B συνεργασίες"

# Χρώματα από το λογότυπο DMS (μπλε / γκρι)
BLUE = "#1E63D6"      # κύριο μπλε (brand)
BLUE_DK = "#12356F"   # σκούρο μπλε
GREY = "#3A3F47"      # ανθρακί (brand)
NAVY = "#1F3864"      # (επικεφαλίδες Excel — παραμένει)
GOLD = "#1E63D6"      # τόνος = brand μπλε (μπάρα προόδου)
GREEN = "#2E7D32"     # κουμπί έναρξης
GREEN_DK = "#1B5E20"
RED = "#B3261E"       # κουμπί ακύρωσης
RED_DK = "#7F1710"
BG = "#EEF2F8"        # φόντο εφαρμογής
CARD = "#FFFFFF"      # φόντο πινάκων
INK = "#1B2430"       # κείμενο
MUTED = "#5B6472"     # δευτερεύον κείμενο


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{BRAND} — {APP_NAME}")
        self.geometry("1080x860")
        self.minsize(980, 780)
        self.configure(bg=BG)

        self.msg_q = queue.Queue()
        self.worker = None
        self.stop_flag = threading.Event()
        self.selected_kads = {}  # id -> descr
        self.all_activities = []  # πλήρης κατάλογος ΚΑΔ (cache)
        self._catalog_loading = False
        self.out_dir = self._resolve_out_dir()

        self._setup_style()
        self._build_ui()
        self.after(120, self._drain_queue)
        self.after(300, self.load_catalog)  # αυτόματη φόρτωση καταλόγου ΚΑΔ

    # ---------------- Θέμα / στυλ ----------------
    def _setup_style(self):
        st = ttk.Style(self)
        try:
            st.theme_use("clam")
        except Exception:
            pass
        st.configure(".", background=BG, foreground=INK, font=("Segoe UI", 10))
        st.configure("TFrame", background=BG)
        st.configure("Card.TFrame", background=CARD)
        st.configure("TLabel", background=BG, foreground=INK)
        st.configure("Muted.TLabel", background=BG, foreground=MUTED)
        st.configure("TCheckbutton", background=BG, foreground=INK)
        st.configure("Card.TCheckbutton", background=CARD, foreground=INK)
        st.map("Card.TCheckbutton", background=[("active", CARD)])
        st.configure("Card.TLabelframe", background=CARD, borderwidth=1, relief="solid")
        st.configure("Card.TLabelframe.Label", background=CARD, foreground=NAVY,
                     font=("Segoe UI", 10, "bold"))
        st.configure("Treeview", background=CARD, fieldbackground=CARD, foreground=INK,
                     rowheight=22, borderwidth=0)
        st.configure("Treeview.Heading", background=NAVY, foreground="white",
                     font=("Segoe UI", 9, "bold"))
        st.map("Treeview.Heading", background=[("active", NAVY)])
        st.map("Treeview", background=[("selected", NAVY)], foreground=[("selected", "white")])

        # Κουμπιά
        st.configure("Start.TButton", font=("Segoe UI", 12, "bold"), padding=(16, 10),
                     background=GREEN, foreground="white", borderwidth=0)
        st.map("Start.TButton", background=[("active", GREEN_DK), ("disabled", "#9AA6B2")])
        st.configure("Stop.TButton", font=("Segoe UI", 10, "bold"), padding=(12, 8),
                     background=RED, foreground="white", borderwidth=0)
        st.map("Stop.TButton", background=[("active", RED_DK), ("disabled", "#C9B7B5")])
        st.configure("Ghost.TButton", font=("Segoe UI", 10), padding=(12, 8),
                     background="#D7DEEA", foreground=NAVY, borderwidth=0)
        st.map("Ghost.TButton", background=[("active", "#C3CEE0")])
        st.configure("Accent.TButton", font=("Segoe UI", 9, "bold"), padding=(10, 5),
                     background=NAVY, foreground="white", borderwidth=0)
        st.map("Accent.TButton", background=[("active", BLUE_DK)])

        # Μπάρα προόδου
        st.configure("Brand.Horizontal.TProgressbar", troughcolor="#D7DEEA",
                     background=GOLD, thickness=22, borderwidth=0)

    # ---------------- Logo έμβλημα: κίονας + βέλος ανάπτυξης (brand) ----------------
    def _draw_logo(self, cv, x, y, s):
        blue, grey = BLUE, GREY
        # Βέλος ανάπτυξης (πίσω, μπλε) — από κάτω-αριστερά προς πάνω-δεξιά
        cv.create_line(x + s * 0.16, y + s * 0.90, x + s * 0.90, y + s * 0.16,
                       fill=blue, width=max(3, int(s * 0.09)), capstyle="round")
        ah = s * 0.16
        cv.create_polygon(
            x + s * 0.90, y + s * 0.16,
            x + s * 0.90 - ah, y + s * 0.20,
            x + s * 0.86, y + s * 0.16 + ah,
            fill=blue, outline="")
        # Κίονας (μπροστά, ανθρακί)
        cx = x + s * 0.44
        cap_w = s * 0.46
        # κιονόκρανο (capital)
        cv.create_rectangle(cx - cap_w / 2, y + s * 0.18,
                            cx + cap_w / 2, y + s * 0.27, fill=grey, outline="")
        cv.create_oval(cx - cap_w / 2 - s * 0.02, y + s * 0.17,
                       cx - cap_w / 2 + s * 0.10, y + s * 0.27, fill=grey, outline="")
        cv.create_oval(cx + cap_w / 2 - s * 0.10, y + s * 0.17,
                       cx + cap_w / 2 + s * 0.02, y + s * 0.27, fill=grey, outline="")
        # κορμός με ραβδώσεις (shaft)
        sh_w = s * 0.34
        cv.create_rectangle(cx - sh_w / 2, y + s * 0.28,
                            cx + sh_w / 2, y + s * 0.74, fill=grey, outline="")
        for k in range(1, 4):
            fx = cx - sh_w / 2 + sh_w * k / 4
            cv.create_line(fx, y + s * 0.30, fx, y + s * 0.72,
                           fill="#525863", width=max(1, int(s * 0.015)))
        # βάση (base)
        cv.create_rectangle(cx - cap_w / 2, y + s * 0.74,
                            cx + cap_w / 2, y + s * 0.82, fill=grey, outline="")

    def _load_logo_image(self, target_h):
        """
        Φορτώνει το λογότυπο από assets/dms_logo.png (ή το πρώτο διαθέσιμο PNG
        στο assets/) και το κλιμακώνει στο ζητούμενο ύψος. Επιστρέφει None αν
        δεν υπάρχει/δεν διαβάζεται.
        """
        candidates = [resource_path(os.path.join("assets", "dms_logo.png"))]
        adir = resource_path("assets")
        try:
            if os.path.isdir(adir):
                for fn in sorted(os.listdir(adir)):
                    if fn.lower().endswith(".png") and fn.lower() != "dms_logo.png":
                        candidates.append(os.path.join(adir, fn))
        except Exception:
            pass
        for path in candidates:
            if not os.path.exists(path):
                continue
            try:
                img = tk.PhotoImage(file=path)
                factor = max(1, round(img.height() / target_h))
                if factor > 1:
                    img = img.subsample(factor, factor)
                return img
            except Exception:
                continue
        return None

    def _build_header(self):
        h = 96
        header = tk.Frame(self, bg=CARD, height=h)
        header.pack(side="top", fill="x")
        header.pack_propagate(False)

        self._logo_img = self._load_logo_image(64)
        if self._logo_img is not None:
            try:
                self.iconphoto(True, self._logo_img)
            except Exception:
                pass
            # Το λογότυπο περιέχει ήδη την επωνυμία — δεν την επαναλαμβάνουμε.
            tk.Label(header, image=self._logo_img, bg=CARD).pack(
                side="left", padx=(18, 16), pady=12)
            txt = tk.Frame(header, bg=CARD)
            txt.pack(side="left", pady=16, anchor="w")
            tk.Label(txt, text=APP_NAME, bg=CARD, fg=GREY,
                     font=("Segoe UI", 15, "bold")).pack(anchor="w")
            tk.Label(txt, text=MOTTO, bg=CARD, fg=MUTED,
                     font=("Segoe UI", 9, "italic")).pack(anchor="w")
        else:
            cv = tk.Canvas(header, width=68, height=68, bg=CARD, highlightthickness=0)
            cv.pack(side="left", padx=(18, 14), pady=12)
            self._draw_logo(cv, 2, 2, 64)
            txt = tk.Frame(header, bg=CARD)
            txt.pack(side="left", pady=14, anchor="w")
            row = tk.Frame(txt, bg=CARD)
            row.pack(anchor="w")
            tk.Label(row, text="DMS", bg=CARD, fg=BLUE,
                     font=("Segoe UI", 20, "bold")).pack(side="left")
            tk.Label(row, text=" Hellas", bg=CARD, fg=GREY,
                     font=("Segoe UI", 20, "bold")).pack(side="left")
            tk.Label(txt, text=COMPANY, bg=CARD, fg=GREY,
                     font=("Segoe UI", 9)).pack(anchor="w")
            tk.Label(txt, text=f"{APP_NAME} · {MOTTO}", bg=CARD, fg=MUTED,
                     font=("Segoe UI", 9, "italic")).pack(anchor="w")

        # μπλε γραμμή-τόνος κάτω από το header
        tk.Frame(self, bg=BLUE, height=3).pack(side="top", fill="x")

    # ---------------- UI ----------------
    def _build_ui(self):
        pad = dict(padx=10, pady=6)

        self._build_header()

        top = ttk.Frame(self)
        top.pack(side="top", fill="x", **pad)

        ttk.Label(top, text="API key:").pack(side="left")
        self.api_var = tk.StringVar(value=core.get_api_key())
        ttk.Entry(top, textvariable=self.api_var, width=42, show="•").pack(side="left", padx=6)

        self.primary_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(top, text="Μόνο Κύρια δραστηριότητα (καθαρή λίστα)",
                        variable=self.primary_var).pack(side="left", padx=12)

        # ----- Κάτω μπάρα ενεργειών + προόδου (καρφωμένη κάτω, πάντα ορατή) -----
        actions = ttk.Frame(self)
        actions.pack(side="bottom", fill="x", **pad)

        brow = ttk.Frame(actions)
        brow.pack(fill="x")
        self.run_btn = ttk.Button(brow, text="▶  ΕΝΑΡΞΗ ΕΞΑΓΩΓΗΣ",
                                  style="Start.TButton", command=self.start_export)
        self.run_btn.pack(side="left")
        self.stop_btn = ttk.Button(brow, text="■ Ακύρωση", command=self.cancel_export,
                                  state="disabled")
        self.stop_btn.pack(side="left", padx=6)
        ttk.Button(brow, text="📂 Άνοιγμα φακέλου output",
                   command=self.open_output).pack(side="left", padx=6)

        prow = ttk.Frame(actions)
        prow.pack(fill="x", pady=(6, 0))
        self.progress = ttk.Progressbar(prow, mode="determinate", maximum=100,
                                        style="Brand.Horizontal.TProgressbar")
        self.progress.pack(side="left", fill="x", expand=True)
        self.pct_var = tk.StringVar(value="0%")
        ttk.Label(prow, textvariable=self.pct_var, width=6, anchor="e",
                  font=("Segoe UI", 10, "bold")).pack(side="left", padx=(8, 0))
        self.status_var = tk.StringVar(value="Έτοιμο.")
        ttk.Label(actions, textvariable=self.status_var, anchor="w",
                  style="Muted.TLabel").pack(fill="x", pady=(4, 0))

        self.log = tk.Text(actions, height=7, wrap="word", state="disabled",
                           bg=CARD, fg=INK, relief="flat", borderwidth=1,
                           highlightthickness=1, highlightbackground="#C7D0DE",
                           font=("Consolas", 9))
        self.log.pack(fill="x", pady=(6, 0))

        # Footer branding
        tk.Frame(actions, bg="#C7D0DE", height=1).pack(fill="x", pady=(8, 0))
        tk.Label(actions, text=f"© {COMPANY}", bg=BG, fg=MUTED,
                 font=("Segoe UI", 8)).pack(anchor="e", pady=(3, 0))

        body = ttk.Panedwindow(self, orient="horizontal")
        body.pack(side="top", fill="both", expand=True, **pad)

        # ----- Αριστερά: αναζήτηση & επιλογή ΚΑΔ -----
        left = ttk.Labelframe(body, text="1) ΚΑΔ", style="Card.TLabelframe")
        body.add(left, weight=3)

        # Grid layout: μόνο η λίστα αποτελεσμάτων (row 2) μεγαλώνει· οι
        # επιλεγμένοι ΚΑΔ φαίνονται σε μία γραμμή (row 4), χωρίς δεύτερη λίστα.
        left.columnconfigure(0, weight=1)
        left.rowconfigure(2, weight=1)

        # row 0: αναζήτηση
        srow = ttk.Frame(left, style="Card.TFrame")
        srow.grid(row=0, column=0, sticky="ew", padx=6, pady=6)
        ttk.Label(srow, text="Αναζήτηση:", background=CARD).pack(side="left")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *a: self._filter_local())
        ent = ttk.Entry(srow, textvariable=self.search_var)
        ent.pack(side="left", fill="x", expand=True, padx=6)
        ent.bind("<Return>", lambda e: self.do_search())
        ttk.Button(srow, text="Ψάξε", style="Accent.TButton",
                   command=self.do_search).pack(side="left")

        # row 1: ετικέτα αποτελεσμάτων
        ttk.Label(left, text="Αποτελέσματα (κλικ/Ctrl για πολλαπλή, μετά «➕ Προσθήκη» ή «Έναρξη»):",
                  background=CARD, foreground=MUTED).grid(row=1, column=0, sticky="w",
                                                          padx=6)
        # row 2: λίστα αποτελεσμάτων (μεγαλώνει — το κύριο στοιχείο)
        res_wrap = ttk.Frame(left)
        res_wrap.grid(row=2, column=0, sticky="nsew", padx=6, pady=(0, 6))
        self.results = ttk.Treeview(res_wrap, columns=("id", "descr"), show="headings",
                                    height=12, selectmode="extended")
        self.results.heading("id", text="ΚΑΔ")
        self.results.heading("descr", text="Περιγραφή")
        self.results.column("id", width=100, anchor="w", stretch=False)
        self.results.column("descr", width=380, anchor="w")
        rsb = ttk.Scrollbar(res_wrap, orient="vertical", command=self.results.yview)
        self.results.configure(yscrollcommand=rsb.set)
        self.results.pack(side="left", fill="both", expand=True)
        rsb.pack(side="right", fill="y")
        self.results.bind("<Double-1>", lambda e: self.add_selected())

        # row 3: κουμπιά προσθήκης
        arow = ttk.Frame(left, style="Card.TFrame")
        arow.grid(row=3, column=0, sticky="ew", padx=6, pady=(0, 6))
        ttk.Button(arow, text="➕ Προσθήκη στην εξαγωγή", style="Accent.TButton",
                   command=self.add_selected).pack(side="left")
        ttk.Button(arow, text="Προσθήκη με κωδικό…", style="Ghost.TButton",
                   command=self.add_manual).pack(side="left", padx=6)
        ttk.Button(arow, text="🗑 Καθαρισμός", style="Ghost.TButton",
                   command=self.clear_selected).pack(side="left", padx=6)

        # row 4: επιλεγμένοι ΚΑΔ σε μία γραμμή (χωρίς δεύτερη λίστα)
        self.chosen_var = tk.StringVar(value="Επιλεγμένοι ΚΑΔ: (κανένας)")
        ttk.Label(left, textvariable=self.chosen_var, background=CARD, foreground=NAVY,
                  font=("Segoe UI", 9, "bold"), wraplength=520, justify="left").grid(
            row=4, column=0, sticky="w", padx=6, pady=(0, 8))

        # ----- Δεξιά: νομοί -----
        right = ttk.Labelframe(body, text="2) Νομοί", style="Card.TLabelframe")
        body.add(right, weight=2)
        self.all_pref_var = tk.BooleanVar(value=True)
        cb = ttk.Checkbutton(right, text="ΟΛΗ η Ελλάδα (όλοι οι νομοί)",
                             variable=self.all_pref_var, command=self._toggle_all_pref)
        cb.configure(style="Card.TCheckbutton")
        cb.pack(anchor="w", padx=8, pady=6)
        ttk.Label(right, text="ή διάλεξε συγκεκριμένους (Ctrl/Shift για πολλαπλή):",
                  background=CARD, foreground=MUTED).pack(anchor="w", padx=8)
        pf_wrap = ttk.Frame(right, style="Card.TFrame")
        pf_wrap.pack(fill="both", expand=True, padx=8, pady=(0, 6))
        self.pref_list = tk.Listbox(pf_wrap, selectmode="extended", exportselection=False,
                                    bg=CARD, fg=INK, selectbackground=NAVY,
                                    selectforeground="white", highlightthickness=1,
                                    highlightbackground="#C7D0DE", relief="flat",
                                    disabledforeground="#9AA6B2", activestyle="none")
        for name in core.PREF_NAMES:
            self.pref_list.insert("end", name)
        psb = ttk.Scrollbar(pf_wrap, orient="vertical", command=self.pref_list.yview)
        self.pref_list.configure(yscrollcommand=psb.set)
        self.pref_list.pack(side="left", fill="both", expand=True)
        psb.pack(side="right", fill="y")
        self.pref_list.configure(state="disabled")

    # ---------------- ενέργειες ----------------
    def _resolve_out_dir(self):
        """
        Επιστρέφει έναν ΣΙΓΟΥΡΑ εγγράψιμο φάκελο εξόδου (τα αρχεία δίπλα στο
        .exe μπορεί να αποτυγχάνουν με «Access denied» σε προστατευμένες θέσεις).
        Προτίμηση: Έγγραφα/Documents → home → temp.
        """
        import tempfile
        home = os.path.expanduser("~")
        candidates = []
        docs = os.path.join(home, "Documents")
        if os.path.isdir(docs):
            candidates.append(os.path.join(docs, "DMS_GEMH_output"))
        candidates.append(os.path.join(home, "DMS_GEMH_output"))
        candidates.append(os.path.join(tempfile.gettempdir(), "DMS_GEMH_output"))
        for d in candidates:
            try:
                os.makedirs(d, exist_ok=True)
                test = os.path.join(d, ".write_test")
                with open(test, "w") as fh:
                    fh.write("ok")
                os.remove(test)
                return d
            except Exception:
                continue
        return os.path.abspath("output")

    def _toggle_all_pref(self):
        self.pref_list.configure(state="disabled" if self.all_pref_var.get() else "normal")

    def log_msg(self, msg):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def load_catalog(self):
        """Φορτώνει ΟΛΟΚΛΗΡΟ τον κατάλογο ΚΑΔ (τρέχουσα ταξινομία) στο άνοιγμα."""
        if self._catalog_loading:
            return
        self._catalog_loading = True
        self.status_var.set("Φόρτωση καταλόγου ΚΑΔ…")

        def work():
            try:
                acts = core.list_activities(api_key=self.api_var.get().strip())
                self.msg_q.put(("catalog", acts))
            except Exception as e:
                self.msg_q.put(("catalog_error", str(e)))

        threading.Thread(target=work, daemon=True).start()

    def do_search(self):
        # Αν έχει φορτωθεί ο κατάλογος, φιλτράρουμε τοπικά (άμεσα).
        if self.all_activities:
            self._filter_local()
        else:
            self.load_catalog()

    def _filter_local(self):
        """Φιλτράρει τον τοπικό κατάλογο ΚΑΔ (χωρίς τόνους). Κενό = όλοι."""
        if not self.all_activities:
            return
        kw = core.strip_acc(self.search_var.get())
        hits = [a for a in self.all_activities
                if kw in core.strip_acc(a.get("descr", ""))]
        self._fill_results(hits)

    def _fill_results(self, hits):
        self.results.delete(*self.results.get_children())
        for a in hits:
            self.results.insert("", "end", values=(a.get("id"), a.get("descr")))
        total = len(self.all_activities)
        if total and len(hits) != total:
            self.status_var.set(f"{len(hits)} από {total} ΚΑΔ (φίλτρο).")
        else:
            self.status_var.set(f"{len(hits)} ΚΑΔ στον κατάλογο.")

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
        self._update_chosen_label()

    def _update_chosen_label(self):
        codes = list(self.selected_kads.keys())
        if codes:
            self.chosen_var.set(f"Επιλεγμένοι ΚΑΔ ({len(codes)}): " + ", ".join(codes))
        else:
            self.chosen_var.set("Επιλεγμένοι ΚΑΔ: (κανένας)")

    def clear_selected(self):
        self.selected_kads.clear()
        self._update_chosen_label()

    def _chosen_prefectures(self):
        if self.all_pref_var.get():
            return "ALL"
        sel = [self.pref_list.get(i) for i in self.pref_list.curselection()]
        return sel or "ALL"

    def start_export(self):
        if self.worker and self.worker.is_alive():
            return
        # Αν δεν έχει προστεθεί τίποτα στην κάτω λίστα αλλά υπάρχουν
        # μαρκαρισμένες γραμμές στα αποτελέσματα, πρόσθεσέ τες αυτόματα.
        if not self.selected_kads and self.results.selection():
            self.add_selected()
        kads = list(self.selected_kads.keys())
        if not kads:
            messagebox.showwarning(
                "Προσοχή",
                "Δεν έχεις επιλέξει κανέναν ΚΑΔ.\n\n"
                "Διάλεξε έναν ή περισσότερους ΚΑΔ από τη λίστα «Αποτελέσματα» "
                "(κλικ πάνω τους) και πάτα «➕ Προσθήκη στην εξαγωγή» — "
                "ή απλώς μαρκάρισέ τους και ξαναπάτα «Έναρξη».")
            return
        prefs = self._chosen_prefectures()
        api_key = self.api_var.get().strip()
        primary = self.primary_var.get()

        self.stop_flag.clear()
        self.run_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.msg_q.put(("percent", 0.0))
        self.log_msg("=" * 60)
        self.log_msg(f"Έναρξη εξαγωγής: {len(kads)} ΚΑΔ | Νομοί: "
                     f"{'ΟΛΗ η Ελλάδα' if prefs == 'ALL' else ', '.join(prefs)}")
        self.log_msg(f"Φάκελος εξόδου: {self.out_dir}")
        self.log_msg("Μεγάλοι ΚΑΔ σε όλη την Ελλάδα μπορεί να θέλουν 40–90 λεπτά "
                     "(όριο ~8 αιτήματα/λεπτό). Άφησέ το να τρέξει.")

        def prog(msg):
            self.msg_q.put(("log", msg))

        def should_stop():
            return self.stop_flag.is_set()

        n = len(kads)

        def work():
            done = []
            for idx, kad in enumerate(kads):
                if self.stop_flag.is_set():
                    break
                self.msg_q.put(("status", f"ΚΑΔ {kad}  ({idx + 1}/{n})…"))
                self.msg_q.put(("log", f"\n== ΚΑΔ {kad} =="))

                def on_prog(scanned, total, _idx=idx):
                    frac = (scanned / total) if total else 0.0
                    overall = (_idx + min(frac, 1.0)) / n * 100.0
                    self.msg_q.put(("percent", overall))

                try:
                    path, rows = core.export_kad(
                        kad, prefectures=prefs, out_dir=self.out_dir, api_key=api_key,
                        primary_only=primary, progress=prog, should_stop=should_stop,
                        on_progress=on_prog)
                    self.msg_q.put(("kad_done", (kad, path, rows)))
                    self.msg_q.put(("percent", (idx + 1) / n * 100.0))
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
        path = self.out_dir
        try:
            os.makedirs(path, exist_ok=True)
        except Exception:
            pass
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)  # noqa: S606 (Windows Explorer)
            elif sys.platform == "darwin":
                import subprocess
                subprocess.Popen(["open", path])
            else:
                webbrowser.open(f"file://{path}")
        except Exception:
            messagebox.showinfo("Φάκελος εξόδου", path)

    # ---------------- ουρά μηνυμάτων από threads ----------------
    def _drain_queue(self):
        try:
            while True:
                kind, payload = self.msg_q.get_nowait()
                if kind == "search_done":
                    self._fill_results(payload)
                elif kind == "catalog":
                    self._catalog_loading = False
                    self.all_activities = payload
                    self._filter_local()
                    self.log_msg(f"Φορτώθηκε ο κατάλογος: {len(payload)} ΚΑΔ "
                                 f"(τρέχουσα ταξινομία).")
                elif kind == "catalog_error":
                    self._catalog_loading = False
                    self.status_var.set("Αποτυχία φόρτωσης καταλόγου ΚΑΔ.")
                    self.log_msg(f"Σφάλμα φόρτωσης καταλόγου ΚΑΔ: {payload}")
                    messagebox.showerror(
                        "Πρόβλημα σύνδεσης",
                        "Δεν φορτώθηκε ο κατάλογος ΚΑΔ.\n\n"
                        "Έλεγξε τη σύνδεση στο διαδίκτυο και το API key, "
                        "και πάτα ξανά «Ψάξε».\n\n"
                        f"Λεπτομέρεια: {payload}")
                elif kind == "log":
                    self.log_msg(payload)
                elif kind == "status":
                    self.status_var.set(payload)
                elif kind == "percent":
                    pct = max(0.0, min(100.0, float(payload)))
                    self.progress["value"] = pct
                    self.pct_var.set(f"{pct:.0f}%")
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
                    if not self.stop_flag.is_set():
                        self.progress["value"] = 100
                        self.pct_var.set("100%")
                    self.status_var.set(f"Ολοκληρώθηκε: {n} αρχείο(α) στον φάκελο εξόδου.")
                    self.log_msg(f"\nΟλοκληρώθηκε. Δημιουργήθηκαν {n} αρχείο(α) στον "
                                 f"φάκελο:\n  {self.out_dir}\n"
                                 f"(κουμπί «📂 Άνοιγμα φακέλου output» για να τα δεις)")
        except queue.Empty:
            pass
        self.after(120, self._drain_queue)


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
