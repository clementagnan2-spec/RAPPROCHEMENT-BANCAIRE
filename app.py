"""Rapprochement bancaire - interface Tkinter (couleurs et formules de l'état Excel d'origine)."""
import os, sys, tkinter as tk
from datetime import date
from tkinter import ttk, filedialog, messagebox
import core

ORANGE, YELLOW, BLUE, RED, GREEN, PURPLE, GREY = "#FFC000", "#FFFF00", "#5B9BD5", "#FF0000", "#92D050", "#7030A0", "#BFBFBF"
PT_BG, SUSP_BG = "#E2F0D9", "#FFF2CC"
FONT, BOLD = ("Arial", 10), ("Arial", 10, "bold")
ETAPES = ["1- Opérer à toutes les saisies", "2- Imprimer les comptes aux mêmes dates (relevés + grand livre)",
          "3- Pointage minutieux", "4- Surligner les suspens", "5- Pointer les suspens des rapprochements antérieurs",
          "6- Rapport de rapprochement (saisir les soldes et les opérations en suspens, mois courant et mois passés)",
          "7- Analyse des suspens du rapprochement (d'abord pour l'entreprise et ensuite pour la banque)",
          "8- Les opérations expliquées sont à comptabiliser",
          "10- Imprimer et signer le rapprochement, imprimer le grand livre, le relevé, les classer par périodes"]


def sort_key(d):
    p = d.split("/")
    return (p[2], p[1], p[0]) if len(p) == 3 else ("", "", "")


class Table(tk.Frame):
    """Tableau 'Opérations / Débit / Crédit' + sous-totaux, solde rapproché, totaux (feuille d'accueil)."""
    W = (85, 235, 105, 105)

    def __init__(self, master, title, bg, sol, tot, tot_fg):
        super().__init__(master)
        tk.Label(self, text=title, bg=bg, font=BOLD, pady=3).pack(fill="x")
        box = tk.Frame(self); box.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(box, columns=("d", "o", "db", "cr"), show="headings", height=9, selectmode="none")
        for c, t, w, a in zip(("d", "o", "db", "cr"), ("Date", "Opérations", "Débit", "Crédit"), self.W, ("center", "w", "e", "e")):
            self.tree.heading(c, text=t); self.tree.column(c, width=w, anchor=a, stretch=(c == "o"))
        sb = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True); sb.pack(side="right", fill="y")
        self.tree.tag_configure("report", font=("Arial", 10, "bold"))
        foot = tk.Frame(self, bd=1, relief="solid"); foot.pack(fill="x")
        foot.grid_columnconfigure(0, weight=1)
        for i, w in ((1, self.W[2]), (2, self.W[3])):
            foot.grid_columnconfigure(i, minsize=w)
        self.val = {}
        for r, (name, lab) in enumerate((("sub", "SOUS TOTAUX"), ("sol", "SOLDE RAPPROCHE"), ("tot", "TOTAUX"))):
            tk.Label(foot, text=lab, font=BOLD, anchor="w").grid(row=r, column=0, sticky="ew", padx=4)
            for k in (0, 1):
                kw = {}
                if name == "sol":
                    kw = dict(bg=sol[k], fg="white")
                if name == "tot":
                    kw = dict(bg=tot, fg=tot_fg)
                l = tk.Label(foot, text="", font=BOLD, anchor="e", padx=4, **kw)
                l.grid(row=r, column=1 + k, sticky="ew", padx=1, pady=1)
                self.val[(name, k)] = l

    def fill(self, rows, sub, sol, tot):
        self.tree.delete(*self.tree.get_children())
        for i, (d, lib, deb, cre) in enumerate(rows):
            self.tree.insert("", "end", values=(d, lib, core.fmt_num(deb), core.fmt_num(cre)), tags=("report",) if i == 0 else ())
        for name, v in (("sub", sub), ("sol", sol), ("tot", tot)):
            for k in (0, 1):
                self.val[(name, k)].config(text=core.fmt_num(v[k], blank_zero=(name == "sol")))


class Home(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master, padding=8)
        self.app, self._job, self._loading = app, None, False
        self.title_lbl = tk.Label(self, font=("Arial", 14, "bold")); self.title_lbl.pack(pady=(0, 6))
        info = ttk.LabelFrame(self, text=" Informations ", padding=6); info.pack(fill="x")
        self.vars = {}
        fields = (("societe", "Société", 24), ("banque", "Code (banque)", 20), ("compte", "Compte", 18), ("numero", "N° de compte", 20),
                  ("mois", "Mois", 12), ("au", "Au (jj/mm/aaaa)", 12))
        for i, (k, lab, w) in enumerate(fields):
            r, cc = divmod(i, 4)
            ttk.Label(info, text=lab).grid(row=r, column=2 * cc, sticky="e", padx=(8, 2), pady=2)
            v = tk.StringVar(); self.vars[k] = v
            ttk.Entry(info, textvariable=v, width=w).grid(row=r, column=2 * cc + 1, pady=2)
            v.trace_add("write", lambda *_, k=k: self.edit_meta(k))
        ttk.Label(info, text="Solde comptabilité").grid(row=1, column=4, sticky="e", padx=(8, 2), pady=2)
        v = tk.StringVar(); self.vars["solde_compta"] = v
        ttk.Combobox(info, textvariable=v, values=("DEBITEUR", "CREDITEUR"), width=12, state="readonly").grid(row=1, column=5, sticky="w", pady=2)
        v.trace_add("write", lambda *_: self.edit_meta("solde_compta"))
        rep = ttk.LabelFrame(self, text=" Report des soldes (1re ligne de chaque tableau) ", padding=6); rep.pack(fill="x", pady=6)
        self.rep = {}
        for i, (k, lab) in enumerate((("compta", "Solde dans nos livres"), ("banque", "Solde dans les livres de la banque"))):
            ttk.Label(rep, text=lab).grid(row=0, column=3 * i, sticky="e", padx=(8, 2))
            m, s = tk.StringVar(), tk.StringVar(); self.rep[k] = (m, s)
            e = ttk.Entry(rep, textvariable=m, width=18, justify="right"); e.grid(row=0, column=3 * i + 1)
            e.bind("<FocusOut>", lambda _e, k=k: self.fmt_rep(k))
            ttk.Combobox(rep, textvariable=s, values=("Débit", "Crédit"), width=7, state="readonly").grid(row=0, column=3 * i + 2, padx=4)
            m.trace_add("write", lambda *_, k=k: self.edit_rep(k)); s.trace_add("write", lambda *_, k=k: self.edit_rep(k))
        mid = tk.Frame(self); mid.pack(fill="both", expand=True)
        self.left = Table(mid, "COMPTE BANQUE DANS NOS LIVRES", ORANGE, (BLUE, RED), GREEN, "black")
        self.right = Table(mid, "COMPTE DANS LES LIVRES DE LA BANQUE", YELLOW, (RED, BLUE), PURPLE, "white")
        self.left.pack(side="left", fill="both", expand=True, padx=(0, 4)); self.right.pack(side="left", fill="both", expand=True)
        bot = tk.Frame(self); bot.pack(fill="x", pady=(8, 0))
        ctl = tk.Frame(bot, bd=1, relief="solid"); ctl.pack(side="left", anchor="n")
        tk.Label(ctl, text="CONTROLES OBLIGATOIRES", bg=GREY, font=("Algerian", 13, "bold"), padx=8).grid(row=0, column=0, columnspan=2, sticky="ew")
        self.ctl = {}
        for r, (key, col, lab) in enumerate((("E54", RED, "Solde rapproché Crédit (nos livres) − Débit (banque)"),
                                             ("E55", BLUE, "Solde rapproché Débit (nos livres) − Crédit (banque)"),
                                             ("E56", "#5E9E2A", "Totaux nos livres : Débit − Crédit"),
                                             ("E57", PURPLE, "Totaux banque : Débit − Crédit")), 1):
            tk.Label(ctl, text=lab, font=("Arial", 9), anchor="w").grid(row=r, column=0, sticky="w", padx=6)
            l = tk.Label(ctl, font=("Arial", 11, "bold"), fg=col, width=16, anchor="e"); l.grid(row=r, column=1, padx=6)
            self.ctl[key] = l
        st = tk.Frame(bot, bd=1, relief="solid"); st.pack(side="left", fill="both", expand=True, padx=(12, 0))
        tk.Label(st, text="LES ETAPES DE RAPPROCHEMENT BANCAIRES", bg=ORANGE, font=BOLD).pack(fill="x")
        for t in ETAPES:
            tk.Label(st, text=t, font=("Arial", 9), anchor="w").pack(fill="x", padx=6)

    # -- édition
    def edit_meta(self, k):
        if not self._loading:
            self.app.state["meta"][k] = self.vars[k].get(); self.app.touch()

    def edit_rep(self, k):
        if not self._loading:
            m, s = self.rep[k]
            self.app.state["report"][k] = {"montant": core.parse_num(m.get()), "sens": s.get() or "Crédit"}; self.app.touch()

    def fmt_rep(self, k):
        m = self.rep[k][0]; self._loading = True
        m.set(core.fmt_num(core.parse_num(m.get()))); self._loading = False

    def load(self):
        st = self.app.state; self._loading = True
        for k, v in self.vars.items():
            v.set(st["meta"].get(k, ""))
        for k, (m, s) in self.rep.items():
            m.set(core.fmt_num(st["report"][k]["montant"])); s.set(st["report"][k]["sens"])
        self._loading = False; self.refresh()

    def refresh(self):
        if self._job:
            self.after_cancel(self._job)
        self._job = self.after(60, self._refresh)

    def _refresh(self):
        self._job = None
        st = self.app.state; m = st["meta"]
        self.title_lbl.config(text=f"RAPPROCHEMENT BANCAIRE AU …{m['au']}…")
        L, R = core.left_rows(st), core.right_rows(st); c = core.calc(L, R)
        self.right.winfo_children()[0].config(text=f"COMPTE {m['compte'].upper()} DANS LES LIVRES DE LA BANQUE")
        self.left.fill(L, (c["D51"], c["E51"]), (c["D52"], c["E52"]), (c["D53"], c["E53"]))
        self.right.fill(R, (c["H51"], c["I51"]), (c["H52"], c["I52"]), (c["H53"], c["I53"]))
        for k, l in self.ctl.items():
            v = c[k]; l.config(text=v if v == "OK" else core.fmt_num(v, False))


class Ledger(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master, padding=8)
        self.app = app
        bar = ttk.Frame(self); bar.pack(fill="x")
        ttk.Button(bar, text="Importer le grand livre…", command=self.do_import).pack(side="left")
        ttk.Button(bar, text="Pointer la sélection", command=lambda: self.toggle(self.tree.selection())).pack(side="left", padx=(12, 2))
        ttk.Button(bar, text="Tout pointer", command=lambda: self.set_all(True)).pack(side="left", padx=2)
        ttk.Button(bar, text="Tout dépointer", command=lambda: self.set_all(False)).pack(side="left", padx=2)
        ttk.Button(bar, text="Vider", command=self.clear).pack(side="left", padx=2)
        self.q = tk.StringVar(); self.flt = tk.StringVar(value="Tous")
        ttk.Combobox(bar, textvariable=self.flt, values=("Tous", "Non pointés", "Pointés"), width=11, state="readonly").pack(side="right")
        ttk.Entry(bar, textvariable=self.q, width=22).pack(side="right", padx=6); ttk.Label(bar, text="Rechercher :").pack(side="right")
        self.q.trace_add("write", lambda *_: self.fill()); self.flt.trace_add("write", lambda *_: self.fill())
        ttk.Label(self, text="Cliquez sur ☐ (ou double-clic / barre d'espace) pour pointer une écriture retrouvée sur le relevé. "
                  "Les écritures NON pointées alimentent « Compte banque dans nos livres ».", foreground="#555").pack(anchor="w", pady=4)
        box = ttk.Frame(self); box.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(box, columns=("p", "d", "o", "db", "cr"), show="headings", selectmode="extended")
        for c, t, w, a in (("p", "Pointé", 60, "center"), ("d", "Date", 90, "center"), ("o", "Libellé / Opération", 480, "w"), ("db", "Débit", 120, "e"), ("cr", "Crédit", 120, "e")):
            self.tree.heading(c, text=t); self.tree.column(c, width=w, anchor=a, stretch=(c == "o"))
        sb = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview); self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True); sb.pack(side="right", fill="y")
        self.tree.tag_configure("ok", background=PT_BG); self.tree.tag_configure("susp", background=SUSP_BG)
        self.tree.bind("<Button-1>", self.on_click); self.tree.bind("<Double-1>", self.on_dbl); self.tree.bind("<space>", lambda e: self.toggle(self.tree.selection()))
        self.info = ttk.Label(self, font=BOLD); self.info.pack(anchor="w", pady=(6, 0))

    def on_click(self, e):
        if self.tree.identify_region(e.x, e.y) == "cell" and self.tree.identify_column(e.x) == "#1":
            iid = self.tree.identify_row(e.y)
            if iid:
                self.toggle([iid]); return "break"

    def on_dbl(self, e):
        iid = self.tree.identify_row(e.y)
        if iid and self.tree.identify_column(e.x) != "#1":
            self.toggle([iid])

    def toggle(self, iids):
        gl = self.app.state["gl"]; ids = [int(i) for i in iids]
        if not ids:
            return
        new = not all(gl[i]["pointe"] for i in ids)
        for i in ids:
            gl[i]["pointe"] = new
        self.app.touch(); self.fill(keep=iids)

    def set_all(self, v):
        for g in self.app.state["gl"]:
            g["pointe"] = v
        self.app.touch(); self.fill()

    def clear(self):
        if self.app.state["gl"] and messagebox.askyesno("Vider", "Supprimer toutes les écritures importées ?"):
            self.app.state["gl"] = []; self.app.touch(); self.fill()

    def do_import(self):
        p = filedialog.askopenfilename(title="Grand livre de la banque", filetypes=[("Excel / CSV", "*.xlsx *.xlsm *.csv *.txt"), ("Tous", "*.*")])
        if not p:
            return
        try:
            rows = core.import_grand_livre(p)
        except Exception as ex:
            messagebox.showerror("Import impossible", str(ex)); return
        if not rows:
            messagebox.showwarning("Import", "Aucune écriture trouvée. Le fichier doit contenir des colonnes Date, Libellé, Débit, Crédit."); return
        if self.app.state["gl"] and not messagebox.askyesno("Import", "Remplacer le grand livre actuel (les pointages seront perdus) ?"):
            return
        self.app.state["gl"] = rows; self.app.touch(); self.fill()

    def fill(self, keep=()):
        gl, flt, q = self.app.state["gl"], self.flt.get(), self.q.get().lower()
        y = self.tree.yview()[0]; self.tree.delete(*self.tree.get_children())
        for i, g in enumerate(gl):
            if (flt == "Pointés" and not g["pointe"]) or (flt == "Non pointés" and g["pointe"]) or (q and q not in g["lib"].lower() and q not in g["date"]):
                continue
            self.tree.insert("", "end", iid=str(i), values=("☑" if g["pointe"] else "☐", g["date"], g["lib"], core.fmt_num(g["debit"]), core.fmt_num(g["credit"])),
                             tags=("ok" if g["pointe"] else "susp",))
        self.tree.yview_moveto(y)
        sel = [i for i in keep if self.tree.exists(i)]
        if sel:
            self.tree.selection_set(sel)
        p = [g for g in gl if g["pointe"]]; n = [g for g in gl if not g["pointe"]]
        s = lambda L, k: core.fmt_num(sum(g[k] for g in L), False)
        self.info.config(text=f"{len(gl)} écritures — pointées : {len(p)} (D {s(p, 'debit')} / C {s(p, 'credit')})   |   "
                              f"en suspens : {len(n)} (D {s(n, 'debit')} / C {s(n, 'credit')})")


class Bank(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master, padding=8)
        self.app, self.edit = app, None
        ttk.Label(self, text="Saisissez ligne par ligne les opérations du relevé bancaire qui ne figurent pas dans le grand livre. "
                  "Elles alimentent « Compte dans les livres de la banque ».", foreground="#555").pack(anchor="w")
        f = ttk.LabelFrame(self, text=" Nouvelle ligne du relevé ", padding=8); f.pack(fill="x", pady=8)
        self.v = {k: tk.StringVar() for k in ("date", "lib", "debit", "credit")}
        for i, (k, lab, w) in enumerate((("date", "Date", 12), ("lib", "Opération", 36), ("debit", "Débit", 16), ("credit", "Crédit", 16))):
            ttk.Label(f, text=lab).grid(row=0, column=2 * i, padx=(8, 2))
            e = ttk.Entry(f, textvariable=self.v[k], width=w, justify="right" if k in ("debit", "credit") else "left"); e.grid(row=0, column=2 * i + 1)
            e.bind("<Return>", lambda _e: self.save())
            if k == "date": self.e_date = e
        self.btn = ttk.Button(f, text="Ajouter", command=self.save); self.btn.grid(row=0, column=8, padx=(12, 2))
        ttk.Button(f, text="Nouveau", command=self.reset).grid(row=0, column=9)
        box = ttk.Frame(self); box.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(box, columns=("d", "o", "db", "cr"), show="headings", selectmode="browse")
        for c, t, w, a in (("d", "Date", 90, "center"), ("o", "Opération", 520, "w"), ("db", "Débit", 130, "e"), ("cr", "Crédit", 130, "e")):
            self.tree.heading(c, text=t); self.tree.column(c, width=w, anchor=a, stretch=(c == "o"))
        sb = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview); self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True); sb.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self.on_select); self.tree.bind("<Delete>", lambda e: self.delete())
        bar = ttk.Frame(self); bar.pack(fill="x", pady=(6, 0))
        ttk.Button(bar, text="Supprimer la ligne", command=self.delete).pack(side="left")
        self.info = ttk.Label(bar, font=BOLD); self.info.pack(side="right")

    def reset(self):
        self.edit = None; self.btn.config(text="Ajouter")
        for k, v in self.v.items():
            v.set(date.today().strftime("%d/%m/%Y") if k == "date" else "")
        self.tree.selection_set(()); self.e_date.focus_set()

    def save(self):
        d = core.parse_date(self.v["date"].get()); lib = self.v["lib"].get().strip()
        deb, cre = core.parse_num(self.v["debit"].get()), core.parse_num(self.v["credit"].get())
        if not d:
            messagebox.showwarning("Saisie", "Date invalide (format jj/mm/aaaa)."); return
        if deb == 0 and cre == 0:
            messagebox.showwarning("Saisie", "Saisissez un montant au débit ou au crédit."); return
        if deb and cre:
            messagebox.showwarning("Saisie", "Un seul montant par ligne : débit OU crédit."); return
        row = {"date": d, "lib": lib, "debit": deb, "credit": cre}
        L = self.app.state["bank"]
        if self.edit is not None:
            L[self.edit] = row
        else:
            L.append(row)
        self.app.touch(); self.fill(); self.reset()

    def on_select(self, _e):
        s = self.tree.selection()
        if s:
            self.edit = int(s[0]); r = self.app.state["bank"][self.edit]
            for k in self.v:
                self.v[k].set(core.fmt_num(r[k]) if k in ("debit", "credit") else r[k])
            self.btn.config(text="Mettre à jour")

    def delete(self):
        s = self.tree.selection()
        if s:
            del self.app.state["bank"][int(s[0])]; self.app.touch(); self.fill(); self.reset()

    def fill(self):
        L = self.app.state["bank"]; self.tree.delete(*self.tree.get_children())
        for i, r in enumerate(L):
            self.tree.insert("", "end", iid=str(i), values=(r["date"], r["lib"], core.fmt_num(r["debit"]), core.fmt_num(r["credit"])))
        self.info.config(text=f"{len(L)} lignes — Débit {core.fmt_num(sum(r['debit'] for r in L), False)}   Crédit {core.fmt_num(sum(r['credit'] for r in L), False)}")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.geometry("1280x820"); self.minsize(1100, 700)
        st = ttk.Style(self); st.theme_use("clam")
        st.configure("Treeview", font=FONT, rowheight=22); st.configure("Treeview.Heading", font=("Arial", 11, "bold"))
        st.configure("TNotebook.Tab", font=("Arial", 10, "bold"), padding=(14, 6)); st.configure("TLabel", font=FONT)
        st.map("TNotebook.Tab", background=[("selected", ORANGE)])
        self.state, self.path, self.dirty = core.new_state(), None, False
        self.nb = ttk.Notebook(self); self.nb.pack(fill="both", expand=True)
        self.home, self.ledger, self.bank = Home(self.nb, self), Ledger(self.nb, self), Bank(self.nb, self)
        self.nb.add(self.home, text="Page d'accueil"); self.nb.add(self.ledger, text="Import du grand livre banque"); self.nb.add(self.bank, text="Saisie relevé bancaire non pointé")
        self.nb.bind("<<NotebookTabChanged>>", lambda e: self.home.refresh())
        m = tk.Menu(self); f = tk.Menu(m, tearoff=0); self.config(menu=m); m.add_cascade(label="Fichier", menu=f)
        for lab, cmd, acc in (("Nouveau", self.new, "<Control-n>"), ("Ouvrir…", self.open, "<Control-o>"), ("Enregistrer", self.save, "<Control-s>"),
                              ("Enregistrer sous…", lambda: self.save(True), None), ("Exporter l'état vers Excel…", self.export, "<Control-e>")):
            f.add_command(label=lab, command=cmd, accelerator=acc.replace("<Control-", "Ctrl+").strip(">").upper() if acc else None)
            if acc: self.bind(acc, lambda e, c=cmd: c())
        f.add_separator(); f.add_command(label="Quitter", command=self.on_close)
        self.protocol("WM_DELETE_WINDOW", self.on_close); self.reload()

    def touch(self):
        self.dirty = True; self.update_title(); self.home.refresh()

    def update_title(self):
        self.title(f"Rapprochement bancaire — {os.path.basename(self.path) if self.path else 'Sans titre'}{' *' if self.dirty else ''}")

    def reload(self):
        self.home.load(); self.ledger.fill(); self.bank.fill(); self.bank.reset(); self.dirty = False; self.update_title()

    def confirm(self):
        if not self.dirty:
            return True
        r = messagebox.askyesnocancel("Rapprochement", "Enregistrer les modifications ?")
        return False if r is None else (self.save() if r else True)

    def new(self):
        if self.confirm():
            self.state, self.path = core.new_state(), None; self.reload()

    def open(self):
        p = filedialog.askopenfilename(filetypes=[("Rapprochement", "*.rapp"), ("Tous", "*.*")])
        if p and self.confirm():
            try:
                self.state, self.path = core.load_project(p), p; self.reload()
            except Exception as ex:
                messagebox.showerror("Ouverture impossible", str(ex))

    def save(self, as_new=False):
        p = self.path
        if as_new or not p:
            p = filedialog.asksaveasfilename(defaultextension=".rapp", filetypes=[("Rapprochement", "*.rapp")])
        if not p:
            return False
        core.save_project(self.state, p); self.path, self.dirty = p, False; self.update_title(); return True

    def export(self):
        p = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")], initialfile="ETAT_DE_RAPPROCHEMENT_BANCAIRE.xlsx")
        if p:
            try:
                core.export_xlsx(self.state, p); messagebox.showinfo("Export", "État exporté (avec les formules) :\n" + p)
            except Exception as ex:
                messagebox.showerror("Export impossible", str(ex))

    def on_close(self):
        if self.confirm():
            self.destroy()


if __name__ == "__main__":
    App().mainloop()
