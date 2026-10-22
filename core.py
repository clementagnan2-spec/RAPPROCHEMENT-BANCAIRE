"""Moteur de calcul, import du grand livre et export Excel (mêmes formules que le fichier d'origine)."""
import csv, json, re, unicodedata
from datetime import datetime, date

MOIS = ["JAN", "FEV", "MARS", "AVR", "MAI", "JUIN", "JUIL", "AOUT", "SEPT", "OCT", "NOV", "DEC"]


# ---------------------------------------------------------------- utilitaires
def parse_num(v):
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[\s\u00a0\u202f]", "", str(v))
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()").replace("FCFA", "").replace("XOF", "")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", "") if re.fullmatch(r"-?\d{1,3}(,\d{3})+", s) else s.replace(",", ".")
    try:
        x = float(s)
    except ValueError:
        return 0.0
    return -x if neg else x


def parse_date(v):
    """Retourne 'jj/mm/aaaa' ou '' si non reconnue."""
    if isinstance(v, (datetime, date)):
        return v.strftime("%d/%m/%Y")
    s = str(v or "").strip()[:10]
    for f in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y", "%d.%m.%Y"):
        try:
            return datetime.strptime(s, f).strftime("%d/%m/%Y")
        except ValueError:
            pass
    return ""


def fmt_num(x, blank_zero=True):
    if blank_zero and abs(x) < 0.005:
        return ""
    return f"{x:,.0f}".replace(",", "\u202f")


def _norm(s):
    s = unicodedata.normalize("NFD", str(s or "")).encode("ascii", "ignore").decode()
    return s.lower().strip()


# ---------------------------------------------------------------- état
def new_state():
    t = date.today()
    return {
        "meta": {"societe": "", "banque": "", "compte": "", "numero": "",
                 "mois": f"{MOIS[t.month - 1]} {t.year}", "au": t.strftime("%d/%m/%Y"),
                 "solde_compta": "DEBITEUR"},
        "report": {"compta": {"montant": 0.0, "sens": "Crédit"}, "banque": {"montant": 0.0, "sens": "Crédit"}},
        "gl": [],    # {date, lib, debit, credit, pointe}
        "bank": [],  # {date, lib, debit, credit}
    }


def _report_row(state, key, label):
    r = state["report"][key]
    m = parse_num(r["montant"])
    return ("", f"{label} {state['meta']['au']}", m if r["sens"] == "Débit" else 0.0, m if r["sens"] == "Crédit" else 0.0)


def left_rows(state):
    """Compte banque dans nos livres = report + écritures du grand livre NON pointées."""
    rows = [_report_row(state, "compta", "Report solde au")]
    rows += [(g["date"], g["lib"], g["debit"], g["credit"]) for g in state["gl"] if not g.get("pointe")]
    return rows


def right_rows(state):
    """Compte dans les livres de la banque = report + lignes du relevé saisies."""
    rows = [_report_row(state, "banque", "Report solde au")]
    rows += [(b["date"], b["lib"], b["debit"], b["credit"]) for b in state["bank"]]
    return rows


def calc(left, right):
    """Reproduit les formules de la feuille (lignes 51 à 57)."""
    D51, E51 = sum(r[2] for r in left), sum(r[3] for r in left)
    H51, I51 = sum(r[2] for r in right), sum(r[3] for r in right)
    D52 = E51 - D51 if D51 - E51 < 0 else 0      # =IF(D51-E51<0,E51-D51,0)
    E52 = D51 - E51 if D51 - E51 > 0 else 0      # =IF(D51-E51>0,D51-E51,0)
    H52 = I51 - H51 if H51 - I51 < 0 else 0      # =IF(H51-I51<0,I51-H51,0)
    I52 = H51 - I51 if H51 - I51 > 0 else 0      # =IF(H51-I51>0,H51-I51,0)
    D53, E53, H53, I53 = D51 + D52, E51 + E52, H51 + H52, I51 + I52
    ok = lambda x: "OK" if abs(x) < 0.005 else x
    return dict(D51=D51, E51=E51, H51=H51, I51=I51, D52=D52, E52=E52, H52=H52, I52=I52,
                D53=D53, E53=E53, H53=H53, I53=I53,
                E54=ok(E52 - H52), E55=ok(D52 - I52), E56=ok(D53 - E53), E57=ok(H53 - I53))


# ---------------------------------------------------------------- import grand livre
def _read_rows(path):
    if path.lower().endswith((".xlsx", ".xlsm")):
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True, data_only=True)
        return [list(r) for r in wb.worksheets[0].iter_rows(values_only=True)]
    with open(path, "r", encoding="utf-8-sig", errors="replace", newline="") as f:
        txt = f.read()
    try:
        dialect = csv.Sniffer().sniff(txt[:4000], delimiters=";,\t|")
    except csv.Error:
        dialect = csv.excel
        dialect.delimiter = ";"
    return list(csv.reader(txt.splitlines(), dialect))


def import_grand_livre(path):
    """Lit un grand livre (xlsx/csv). Détecte les en-têtes Date / Libellé / Débit / Crédit."""
    data = _read_rows(path)
    kw = {"date": ("date",), "debit": ("debit",), "credit": ("credit",),
          "lib": ("libelle", "operation", "designation", "description", "intitule", "piece", "reference", "tiers")}
    hdr, cols = None, {}
    for i, row in enumerate(data[:40]):
        found = {k: [j for j, c in enumerate(row) if any(w in _norm(c) for w in ws)] for k, ws in kw.items()}
        if found["date"] and found["debit"] and found["credit"]:
            hdr, cols = i, found
            break
    if hdr is None:  # repli : colonnes A..D = date, libellé, débit, crédit
        hdr, cols = -1, {"date": [0], "lib": [1], "debit": [2], "credit": [3]}
    out = []
    for row in data[hdr + 1:]:
        g = lambda j: row[j] if j < len(row) else None
        lib = " / ".join(str(g(j)).strip() for j in cols["lib"] if g(j) not in (None, "")) if cols["lib"] else ""
        d, c = parse_num(g(cols["debit"][0])), parse_num(g(cols["credit"][0]))
        dt = parse_date(g(cols["date"][0]))
        if (d == 0 and c == 0) or (not dt and (not lib or _norm(lib).startswith(("total", "solde", "report")))):
            continue
        out.append({"date": dt, "lib": lib, "debit": d, "credit": c, "pointe": False})
    return out


# ---------------------------------------------------------------- fichiers projet
def save_project(state, path):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=1)


def load_project(path):
    with open(path, "r", encoding="utf-8") as f:
        s = json.load(f)
    base = new_state()
    base["meta"].update(s.get("meta", {}))
    for k in ("compta", "banque"):
        base["report"][k].update(s.get("report", {}).get(k, {}))
    base["gl"], base["bank"] = s.get("gl", []), s.get("bank", [])
    return base


# ---------------------------------------------------------------- export Excel (avec formules)
def export_xlsx(state, path):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    wb = Workbook()
    ws = wb.active
    ws.title = "Etat de rapprochement"
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEFGHI", (0.9, 14.7, 44.6, 15.1, 19.3, 14.7, 39.1, 19, 18.3)):
        ws.column_dimensions[col].width = w
    F = lambda **k: Font(name="Arial", **k)
    fill = lambda c: PatternFill("solid", fgColor=c)
    thin, dbl = Side(style="thin"), Side(style="double")
    m = state["meta"]
    ws.merge_cells("B2:I2")
    ws["B2"] = f"RAPPROCHEMENT BANCAIRE AU {m['au']}"
    ws["B2"].font, ws["B2"].alignment = F(bold=True, size=12), Alignment(horizontal="center")
    for cell, txt in (("B4", f"SOCIETE: {m['societe']}"), ("D4", f"CODE: {m['banque']}"), ("F4", f"COMPTE: {m['compte']}"),
                      ("H4", "AU"), ("B5", f"MOIS: {m['mois']}"), ("D5", f"COMPTE N°: {m['numero']}"),
                      ("G5", f"SOLDE COMPTABILITE: {m['solde_compta']}")):
        ws[cell], ws[cell].font = txt, F(bold=True, size=10)
    ws["I4"], ws["I4"].number_format, ws["I4"].font = datetime.strptime(m["au"], "%d/%m/%Y") if parse_date(m["au"]) else m["au"], "dd/mm/yyyy", F(bold=True, size=10)
    ws.merge_cells("B7:E7"); ws.merge_cells("F7:I7")
    ws["B7"], ws["F7"] = "COMPTE BANQUE DANS NOS LIVRES", f"COMPTE {m['compte'].upper()} DANS LES LIVRES DE LA BANQUE"
    for c, col in (("B7", "FFC000"), ("F7", "FFFF00")):
        ws[c].fill, ws[c].font, ws[c].alignment = fill(col), F(bold=True, size=10), Alignment(horizontal="center")
    ws.merge_cells("B8:C8"); ws.merge_cells("F8:G8")
    for c, t in (("B8", "Opérations"), ("D8", "Débit"), ("E8", "Crédit"), ("F8", "Opérations"), ("H8", "Débit"), ("I8", "Crédit")):
        ws[c], ws[c].font, ws[c].alignment = t, F(bold=True, size=12), Alignment(horizontal="center")
    L, R = left_rows(state), right_rows(state)
    last = 8 + max(len(L), len(R))
    for rows, cols in ((L, "BCDE"), (R, "FGHI")):
        for i, (d, lib, deb, cre) in enumerate(rows):
            r = 9 + i
            ws[f"{cols[0]}{r}"] = datetime.strptime(d, "%d/%m/%Y") if parse_date(d) else d
            ws[f"{cols[0]}{r}"].number_format = "dd/mm/yyyy"
            ws[f"{cols[1]}{r}"] = lib
            for c, v in ((cols[2], deb), (cols[3], cre)):
                if v:
                    ws[f"{c}{r}"] = v
            for c in cols:
                ws[f"{c}{r}"].font = F(size=11, bold=(i == 0))
    for r in range(9, last + 1):
        for c in "DEHI":
            ws[f"{c}{r}"].number_format = "#,##0"
    s, so, t = last + 2, last + 3, last + 4  # SOUS TOTAUX / SOLDE RAPPROCHE / TOTAUX
    for r, lab in ((s, "SOUS TOTAUX"), (so, "SOLDE RAPPROCHE"), (t, "TOTAUX")):
        for c in "BF":
            ws[f"{c}{r}"], ws[f"{c}{r}"].font = lab, F(bold=True, size=11)
    for c in "DEHI":
        ws[f"{c}{s}"] = f"=SUM({c}9:{c}{last + 1})"
    ws[f"D{so}"] = f"=IF(D{s}-E{s}<0,E{s}-D{s},0)"
    ws[f"E{so}"] = f"=IF(D{s}-E{s}>0,D{s}-E{s},0)"
    ws[f"H{so}"] = f"=IF(H{s}-I{s}<0,I{s}-H{s},0)"
    ws[f"I{so}"] = f"=IF(H{s}-I{s}>0,H{s}-I{s},0)"
    for c in "DEHI":
        ws[f"{c}{t}"] = f"=SUM({c}{s}:{c}{so})"
    colors = {"D": ("5B9BD5", "5B9BD5"), "E": ("FF0000",) * 2, "H": ("FF0000",) * 2, "I": ("5B9BD5",) * 2}
    for c in "DEHI":
        ws[f"{c}{s}"].font = F(bold=True, size=10)
        ws[f"{c}{so}"].fill, ws[f"{c}{so}"].font = fill(colors[c][0]), F(bold=True, size=10, color="FFFFFF")
        ws[f"{c}{t}"].fill = fill("92D050" if c in "DE" else "7030A0")
        ws[f"{c}{t}"].font = F(bold=True, size=10, color="000000" if c in "DE" else "FFFFFF")
        for r in (s, so, t):
            ws[f"{c}{r}"].number_format = "#,##0"
            ws[f"{c}{r}"].border = Border(top=thin, bottom=thin, left=thin, right=thin)
    ws[f"E{t + 1}"] = f'=IF(E{so}-H{so}=0,"OK",E{so}-H{so})'
    ws[f"E{t + 2}"] = f'=IF(D{so}-I{so}=0,"OK",D{so}-I{so})'
    ws[f"E{t + 3}"] = f'=IF(D{t}-E{t}=0,"OK",D{t}-E{t})'
    ws[f"E{t + 4}"] = f'=IF(H{t}-I{t}=0,"OK",H{t}-I{t})'
    for k, col in enumerate(("FF0000", "5B9BD5", "92D050", "7030A0")):
        ws[f"E{t + 1 + k}"].font, ws[f"E{t + 1 + k}"].number_format = F(bold=True, size=11, color=col), "#,##0"
    ws.merge_cells(f"C{t + 2}:C{t + 3}")
    ws[f"C{t + 2}"], ws[f"C{t + 2}"].font, ws[f"C{t + 2}"].fill = "CONTROLES OBLIGATOIRES", F(bold=True, size=12), fill("BFBFBF")
    ws[f"C{t + 2}"].alignment = Alignment(horizontal="center", vertical="center")
    ws.page_setup.orientation, ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = "landscape", 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save(path)
