"""Écriture du classeur de reporting mis en forme (openpyxl, graphiques Excel natifs)."""
from __future__ import annotations

from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from .cleaning import LABELS

# ---------------------------------------------------------------------------
# Charte graphique (à adapter aux couleurs du client)
# ---------------------------------------------------------------------------
PRIMARY = "1F4E79"
LIGHT = "DDEBF7"
GREY = "595959"
TABLE_STYLE = "TableStyleMedium2"
EUR = '#,##0" €"'
EUR2 = '#,##0.00" €"'
PCT = '+0.0%;-0.0%;0.0%'
SHARE = "0.0%"
MONTH = "mmm yyyy"

TITLE = Font(size=18, bold=True, color=PRIMARY)
SUBTITLE = Font(size=10, italic=True, color=GREY)
H2 = Font(size=12, bold=True, color=PRIMARY)
KPI_LABEL = Font(size=10, color=GREY)
KPI_VALUE = Font(size=20, bold=True, color=PRIMARY)
KPI_FILL = PatternFill("solid", fgColor=LIGHT)
THIN = Side(style="thin", color="BFBFBF")
GREEN = Font(color="2E7D32")
RED = Font(color="C62828")


def _autosize(ws, min_width=10, max_width=45):
    widths = {}
    for row in ws.iter_rows():
        for c in row:
            if c.value is not None:
                n = len(c.value.strftime("%d/%m/%Y")) if isinstance(c.value, datetime) else len(str(c.value))
                widths[c.column_letter] = max(widths.get(c.column_letter, 0), n)
    for col, w in widths.items():
        ws.column_dimensions[col].width = max(min_width, min(max_width, w + 3))


def _write_table(ws, df: pd.DataFrame, name: str, top: int = 1, left: int = 1, formats=None,
                 variation_cols=()):
    """Écrit un DataFrame sous forme de Tableau Excel (filtres, lignes alternées)."""
    formats = formats or {}
    for j, col in enumerate(df.columns):
        ws.cell(top, left + j, col)
    for i, row in enumerate(df.itertuples(index=False), start=1):
        for j, v in enumerate(row):
            if v is not None and not (isinstance(v, float) and pd.isna(v)):
                if isinstance(v, pd.Timestamp):
                    v = v.to_pydatetime()
                ws.cell(top + i, left + j, v)
    last_row = top + max(len(df), 1)
    last_col = left + len(df.columns) - 1
    ref = f"{get_column_letter(left)}{top}:{get_column_letter(last_col)}{last_row}"
    t = Table(displayName=name, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name=TABLE_STYLE, showRowStripes=True)
    ws.add_table(t)
    for j, col in enumerate(df.columns):
        letter = get_column_letter(left + j)
        if col in formats:
            for r in range(top + 1, last_row + 1):
                ws[f"{letter}{r}"].number_format = formats[col]
        if col in variation_cols and len(df):
            rng = f"{letter}{top + 1}:{letter}{last_row}"
            ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0"], font=RED))
            ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["0"], font=GREEN))
    return ref


def _bar(title, ws, cats_col, vals_col, top, n, horizontal=True, y_title="CA HT (€)"):
    ch = BarChart()
    ch.type = "bar" if horizontal else "col"
    ch.title = title
    ch.style = 2
    ch.legend = None
    ch.y_axis.title = y_title
    ch.y_axis.numFmt = '#,##0'
    ch.y_axis.majorGridlines = None
    data = Reference(ws, min_col=vals_col, min_row=top, max_row=top + n)
    cats = Reference(ws, min_col=cats_col, min_row=top + 1, max_row=top + n)
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    if horizontal:
        ch.x_axis.scaling.orientation = "maxMin"  # le plus gros en haut
    ch.series[0].graphicalProperties.solidFill = PRIMARY
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    return ch


def _line(ws, n, top=1):
    ch = LineChart()
    ch.title = "Chiffre d'affaires HT mensuel"
    ch.style = 2
    ch.legend = None
    ch.height, ch.width = 7.5, 22
    ch.y_axis.title = "CA HT (€)"
    ch.y_axis.numFmt = '#,##0'
    ch.x_axis.number_format = "mmm yy"
    ch.x_axis.majorTimeUnit = "months"
    data = Reference(ws, min_col=2, min_row=top, max_row=top + n)
    cats = Reference(ws, min_col=1, min_row=top + 1, max_row=top + n)
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    s = ch.series[0]
    s.graphicalProperties.line.solidFill = PRIMARY
    s.graphicalProperties.line.width = 28000
    s.smooth = False
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    return ch


def _kpi(ws, row, col, label, value, fmt, width=2):
    if value is not None and pd.isna(value):
        value = None
    ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + width - 1)
    ws.merge_cells(start_row=row + 1, start_column=col, end_row=row + 1, end_column=col + width - 1)
    a, b = ws.cell(row, col, label), ws.cell(row + 1, col, value if value is not None else "n.d.")
    a.font, b.font = KPI_LABEL, KPI_VALUE
    b.number_format = fmt
    for r in (row, row + 1):
        for c in range(col, col + width):
            cell = ws.cell(r, c)
            cell.fill = KPI_FILL
            cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row + 1].height = 32
    if fmt == PCT and value is not None:
        b.font = Font(size=20, bold=True, color=GREEN.color if value >= 0 else RED.color)


def build_workbook(clean_df, quality, monthly_df, region_df, clients_df, products_df, k,
                   source_name: str, company: str = "Société exemple") -> Workbook:
    wb = Workbook()
    syn = wb.active
    syn.title = "Synthèse"

    # --- Onglets de détail (écrits d'abord : les graphiques de la synthèse s'y réfèrent) ---
    ws_m = wb.create_sheet("CA mensuel")
    m = monthly_df.rename(columns={"mois": "Mois", "ca_ht": "CA HT", "nb_factures": "Nb factures",
                                   "var_m1": "Évol. M-1", "var_n1": "Évol. N-1"})
    _write_table(ws_m, m, "CA_mensuel", formats={"Mois": MONTH, "CA HT": EUR, "Évol. M-1": PCT,
                                                  "Évol. N-1": PCT}, variation_cols=("Évol. M-1", "Évol. N-1"))
    ws_m.freeze_panes = "A2"
    _autosize(ws_m)
    ws_m.add_chart(_line(ws_m, len(m)), "G2")

    ws_r = wb.create_sheet("Régions")
    r = region_df.rename(columns={"region": "Région", "ca_ht": "CA HT", "nb_clients": "Nb clients",
                                  "part": "Part du CA"})
    _write_table(ws_r, r, "Regions", formats={"CA HT": EUR, "Part du CA": SHARE})
    _autosize(ws_r)
    ws_r.add_chart(_bar("CA HT par région", ws_r, 1, 2, 1, len(r)), "F2")

    ws_c = wb.create_sheet("Top clients")
    c = clients_df.rename(columns={"client": "Client", "ca_ht": "CA HT", "nb_factures": "Nb factures",
                                   "part": "Part du CA"})
    _write_table(ws_c, c, "Top_clients", formats={"CA HT": EUR, "Part du CA": SHARE})
    _autosize(ws_c)
    ws_c.add_chart(_bar(f"Top {len(c)} clients (CA HT)", ws_c, 1, 2, 1, len(c)), "F2")

    ws_p = wb.create_sheet("Top produits")
    p = products_df.rename(columns={"produit": "Produit", "ca_ht": "CA HT", "quantite": "Quantité",
                                    "part": "Part du CA"})
    _write_table(ws_p, p, "Top_produits", formats={"CA HT": EUR, "Part du CA": SHARE})
    _autosize(ws_p)
    ws_p.add_chart(_bar(f"Top {len(p)} produits (CA HT)", ws_p, 1, 2, 1, len(p)), "F2")

    # --- Anomalies ---
    ws_a = wb.create_sheet("Anomalies")
    ws_a["A1"], ws_a["A1"].font = "Contrôle qualité de l'export", H2
    ws_a["A2"] = f"Lignes lues : {quality.rows_read} — lignes retenues : {quality.rows_kept}"
    row = 4
    ws_a.cell(row, 1, "Corrections automatiques").font = Font(bold=True)
    for label, n in quality.corrections.items():
        row += 1
        ws_a.cell(row, 1, label)
        ws_a.cell(row, 2, n)
    row += 2
    ws_a.cell(row, 1, "Lignes à vérifier ou exclues").font = Font(bold=True)
    if len(quality.anomalies):
        _write_table(ws_a, quality.anomalies, "Anomalies", top=row + 1)
    else:
        ws_a.cell(row + 1, 1, "Aucune anomalie détectée.")
    _autosize(ws_a, max_width=50)
    ws_a.column_dimensions["A"].width = 34

    # --- Données nettoyées ---
    ws_d = wb.create_sheet("Données nettoyées")
    cols = [c for c in ["ligne", "date", "mois", "facture", "client", "produit", "region", "commercial",
                        "quantite", "montant_ht", "tva", "montant_ttc"] if c in clean_df]
    d = clean_df[cols].rename(columns={**LABELS, "ligne": "Ligne source"})
    _write_table(ws_d, d, "Donnees", formats={"Date": "dd/mm/yyyy", "Mois": MONTH, "Montant HT": EUR2,
                                              "TVA": EUR2, "Montant TTC": EUR2})
    ws_d.freeze_panes = "A2"
    _autosize(ws_d)

    # --- Synthèse ---
    syn.sheet_view.showGridLines = False
    for col in range(1, 13):
        syn.column_dimensions[get_column_letter(col)].width = 13
    syn["A1"], syn["A1"].font = f"Reporting des ventes — {company}", TITLE
    syn["A2"] = (f"Mois de référence : {k.mois_ref:%m/%Y}  •  Source : {source_name}  •  "
                 f"Généré le {datetime.now():%d/%m/%Y à %H:%M}")
    syn["A2"].font = SUBTITLE
    _kpi(syn, 4, 1, f"CA HT {k.mois_ref:%m/%Y}", k.ca_mois, EUR)
    _kpi(syn, 4, 3, "Évolution vs mois précédent", k.var_m1, PCT)
    _kpi(syn, 4, 5, "Évolution vs même mois N-1", k.var_n1, PCT)
    _kpi(syn, 4, 7, f"CA HT cumulé {k.mois_ref.year}", k.ca_cumul, EUR)
    _kpi(syn, 4, 9, f"Évol. cumul vs {k.mois_ref.year - 1}", k.var_cumul, PCT)
    _kpi(syn, 4, 11, "Panier moyen du mois", k.panier_moyen_mois, EUR)
    _kpi(syn, 7, 1, "Factures du mois", k.nb_factures_mois, "0")
    _kpi(syn, 7, 3, "Clients actifs (12 mois)", k.clients_actifs_12m, "0")
    _kpi(syn, 7, 5, "Lignes à vérifier", quality.n_anomalies, "0")
    note = syn.cell(7, 7, "→ détail dans l'onglet « Anomalies »")
    note.font = SUBTITLE
    note.alignment = Alignment(vertical="center")
    syn.merge_cells("G7:J8")

    syn.add_chart(_line(ws_m, len(m)), "A11")
    ch_r = _bar("CA HT par région", ws_r, 1, 2, 1, len(r))
    ch_r.height, ch_r.width = 7.5, 11
    syn.add_chart(ch_r, "A27")
    ch_c = _bar("Top 5 clients", ws_c, 1, 2, 1, min(5, len(c)))
    ch_c.height, ch_c.width = 7.5, 11
    syn.add_chart(ch_c, "G27")
    syn.page_setup.orientation = "landscape"
    syn.page_setup.fitToWidth = 1
    syn.sheet_properties.pageSetUpPr.fitToPage = True
    syn.sheet_properties.tabColor = PRIMARY
    wb.active = 0
    return wb
