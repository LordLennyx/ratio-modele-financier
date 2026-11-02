"""Génère le classeur Excel du modèle financier (formules actives)."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference

wb = Workbook()
h = wb.active; h.title = "Hypothèses"
VERT = "134E4A"; F_IN = PatternFill("solid", fgColor="E9F3F1"); F_T = PatternFill("solid", fgColor="F3F5F4")
bold = Font(name="Calibri", bold=True); bleu = Font(name="Calibri", color="1F4FB4")
h["A1"] = "Jus Mbanga · hypothèses du modèle"; h["A1"].font = Font(size=14, bold=True, color=VERT)
h["A2"] = "Les cellules en bleu sur fond vert clair sont modifiables ; tout le reste du classeur se recalcule."
H = [("Volume de l'an 1 (bouteilles)", 180000, "#,##0"), ("Croissance des volumes", .25, "0%"),
     ("Prix de vente HT (€)", 2.4, "0.00"), ("Hausse annuelle des prix", .03, "0.0%"),
     ("Coût matière et emballage (% CA)", .38, "0%"), ("Charges fixes de l'an 1 (k€)", 168, "0.0"),
     ("Inflation des charges fixes", .04, "0.0%"), ("Investissement initial (k€)", 340, "0.0"),
     ("Durée d'amortissement (ans)", 5, "0"), ("Emprunt bancaire (k€)", 200, "0.0"),
     ("Taux de l'emprunt", .045, "0.00%"), ("Durée de l'emprunt (ans)", 5, "0"),
     ("BFR (% du CA)", .12, "0%"), ("Impôt sur les sociétés", .25, "0.0%"), ("Taux d'actualisation", .09, "0.0%")]
for i, (n, v, f) in enumerate(H, start=4):
    h.cell(i, 1, n); c = h.cell(i, 2, v); c.number_format = f; c.font = bleu; c.fill = F_IN
h.column_dimensions["A"].width = 38; h.column_dimensions["B"].width = 14
N = {n: f"Hypothèses!$B${i}" for i, (n, _, _) in enumerate(H, start=4)}
k = lambda s: next(v for n, v in N.items() if n.startswith(s))
VOL, CRV, PX, HP, CM, CF, INF, INV, DUR, EMP, TX, DE, BFR, IS, ACT = [k(s) for s in [
    "Volume", "Croissance", "Prix", "Hausse", "Coût", "Charges", "Inflation", "Investissement",
    "Durée d'amort", "Emprunt", "Taux de l'", "Durée de l'emp", "BFR", "Impôt", "Taux d'act"]]

m = wb.create_sheet("Modèle")
m["A1"] = "Jus Mbanga · modèle sur 5 ans (k€)"; m["A1"].font = Font(size=14, bold=True, color=VERT)
cols = "BCDEFG"  # B = an 0, C..G = an 1..5
m["A3"] = "Année"; m.cell(3, 2, 0)
for j in range(1, 6): m.cell(3, 2 + j, j)
for c in m[3]: c.font = bold; c.fill = F_T
L = {}
def ligne(r, nom, f, tot=False, an0=None):
    m.cell(r, 1, nom); L[nom] = r
    if an0 is not None: m.cell(r, 2, an0)
    for j in range(1, 6):
        col = cols[j]; prev = cols[j - 1]
        m.cell(r, 2 + j, "=" + f(col, prev, j)).number_format = '#,##0.0;[Red]-#,##0.0'
    if an0 is not None: m.cell(r, 2).number_format = '#,##0.0;[Red]-#,##0.0'
    if tot:
        for c in m[r]: c.font = bold; c.fill = F_T
r = 4
ligne(4, "Volume (bouteilles)", lambda c, p, j: f"{VOL}*(1+{CRV})^({c}$3-1)")
for j in range(1, 6): m.cell(4, 2 + j).number_format = "#,##0"
ligne(5, "Prix unitaire (€)", lambda c, p, j: f"{PX}*(1+{HP})^({c}$3-1)")
for j in range(1, 6): m.cell(5, 2 + j).number_format = "0.00"
ligne(6, "Chiffre d'affaires", lambda c, p, j: f"{c}4*{c}5/1000", True)
ligne(7, "Coût matière et emballage", lambda c, p, j: f"-{c}6*{CM}")
ligne(8, "Marge brute", lambda c, p, j: f"{c}6+{c}7", True)
ligne(9, "Charges fixes", lambda c, p, j: f"-{CF}*(1+{INF})^({c}$3-1)")
ligne(10, "Excédent brut d'exploitation", lambda c, p, j: f"{c}8+{c}9", True)
ligne(11, "Dotations aux amortissements", lambda c, p, j: f"-IF({c}$3<={DUR},{INV}/{DUR},0)")
ligne(12, "Résultat d'exploitation", lambda c, p, j: f"{c}10+{c}11", True)
ligne(13, "Charges d'intérêts", lambda c, p, j: f"IF({EMP}=0,0,IPMT({TX},{c}$3,{DE},{EMP}))")
ligne(14, "Résultat avant impôt", lambda c, p, j: f"{c}12+{c}13")
ligne(15, "Impôt sur les sociétés", lambda c, p, j: f"-MAX(0,{c}14*{IS})")
ligne(16, "Résultat net", lambda c, p, j: f"{c}14+{c}15", True)
m["A18"] = "Plan de trésorerie"; m["A18"].font = Font(bold=True, color=VERT)
ligne(19, "Capacité d'autofinancement", lambda c, p, j: f"{c}16-{c}11")
ligne(20, "BFR", lambda c, p, j: f"{c}6*{BFR}")
m["B20"] = 0
ligne(21, "Variation du BFR", lambda c, p, j: f"-({c}20-{p}20)")
ligne(22, "Remboursement d'emprunt", lambda c, p, j: f"IF({EMP}=0,0,PPMT({TX},{c}$3,{DE},{EMP}))")
m["B23"] = f"=MAX(0,{INV}-{EMP})+{EMP}-{INV}"; m["B23"].number_format = '#,##0.0;[Red]-#,##0.0'
ligne(23, "Trésorerie de clôture", lambda c, p, j: f"{p}23+{c}19+{c}21+{c}22", True)
m["A25"] = "Rentabilité du projet"; m["A25"].font = Font(bold=True, color=VERT)
ligne(26, "Flux de trésorerie disponible", lambda c, p, j: f"{c}12*(1-{IS})-{c}11+{c}21" + ("+G20" if j == 5 else ""), True)
m["B26"] = f"=-{INV}"; m["B26"].number_format = '#,##0.0;[Red]-#,##0.0'
res = [("VAN (k€)", f"=B26+NPV({ACT},C26:G26)", '#,##0.0;[Red]-#,##0.0'),
       ("TRI", "=IRR(B26:G26)", "0.0%"),
       ("Point mort de l'an 1 (bouteilles)", f"=ROUNDUP({CF}*1000/({PX}*(1-{CM})),0)", "#,##0"),
       ("Résultat net cumulé (k€)", "=SUM(C16:G16)", '#,##0.0;[Red]-#,##0.0')]
for i, (n, f, nf) in enumerate(res, start=28):
    m.cell(i, 1, n).font = bold; c = m.cell(i, 2, f); c.number_format = nf; c.font = Font(bold=True, color=VERT)
m.column_dimensions["A"].width = 34
for c in "BCDEFG": m.column_dimensions[c].width = 13
m.freeze_panes = "B4"
ch = BarChart(); ch.title = "Chiffre d'affaires et résultat net"; ch.y_axis.title = "k€"
ch.add_data(Reference(m, min_col=3, max_col=7, min_row=6), from_rows=True, titles_from_data=False)
ch.add_data(Reference(m, min_col=3, max_col=7, min_row=16), from_rows=True, titles_from_data=False)
ch.set_categories(Reference(m, min_col=3, max_col=7, min_row=3)); ch.height = 7; ch.width = 16
ch.series[0].tx = None; m.add_chart(ch, "I3")
wb.save("modele-financier-jus-mbanga.xlsx"); print("ok")
