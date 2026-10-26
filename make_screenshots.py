"""Génère les images du README à partir des mêmes calculs que le reporting.

- docs/img/ca_mensuel.png      : CA HT mensuel (matplotlib)
- docs/img/apercu_synthese.png : aperçu de l'onglet « Synthèse » (reconstitution matplotlib,
                                 pas une capture d'Excel)
- docs/img/avant_apres.png     : extrait de l'export brut et des mêmes lignes nettoyées

Usage : python make_screenshots.py [data/export_ventes_brut.xlsx]
"""
from __future__ import annotations

import sys
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.dates as mdates  # noqa: E402
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib.ticker import FuncFormatter  # noqa: E402

from reporting.aggregates import by_region, kpis, monthly, top_clients  # noqa: E402
from reporting.cleaning import clean, read_export  # noqa: E402

PRIMARY, LIGHT, GREY, GREEN, RED = "#1F4E79", "#DDEBF7", "#595959", "#2E7D32", "#C62828"
OUT = Path("docs/img")
EUR = FuncFormatter(lambda x, _: f"{x:,.0f} €".replace(",", " "))
MOIS = ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc."]


def fr_month(d):
    return f"{MOIS[d.month - 1]} {d:%y}"


def fmt_eur(x):
    return f"{x:,.0f} €".replace(",", " ")


def fmt_pct(x):
    return "n.d." if x is None or x != x else f"{x:+.1%}".replace(".", ",")


def style(ax):
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.spines["left"].set_color("#BFBFBF")
    ax.spines["bottom"].set_color("#BFBFBF")
    ax.tick_params(colors=GREY, labelsize=9)
    ax.grid(axis="y", color="#E5E5E5", linewidth=0.8)
    ax.set_axisbelow(True)


def monthly_chart(ax, m):
    ax.plot(m["mois"], m["ca_ht"], color=PRIMARY, linewidth=2.2, marker="o", markersize=3.5)
    ax.yaxis.set_major_formatter(EUR)
    ax.xaxis.set_major_locator(mdates.MonthLocator(bymonth=[1, 4, 7, 10]))
    ax.xaxis.set_major_formatter(FuncFormatter(lambda x, _: fr_month(mdates.num2date(x))))
    ax.set_ylim(bottom=0)
    last = m.iloc[-1]
    ax.annotate(fmt_eur(last["ca_ht"]), (last["mois"], last["ca_ht"]), xytext=(6, 4),
                textcoords="offset points", fontsize=9, color=PRIMARY, fontweight="bold")
    style(ax)


def chart_png(m):
    fig, ax = plt.subplots(figsize=(10, 4), dpi=150)
    monthly_chart(ax, m)
    ax.set_title("Chiffre d'affaires HT mensuel (données fictives)", loc="left", fontsize=13,
                 color=PRIMARY, fontweight="bold")
    fig.tight_layout()
    fig.savefig(OUT / "ca_mensuel.png")
    plt.close(fig)


def synthese_png(df, q):
    k = kpis(df)
    m = monthly(df)
    reg = by_region(df)
    top = top_clients(df, 5)
    fig = plt.figure(figsize=(12, 8.2), dpi=130)
    fig.patch.set_facecolor("white")
    fig.text(0.03, 0.965, "Reporting des ventes — Société exemple", fontsize=18, fontweight="bold",
             color=PRIMARY, va="top")
    fig.text(0.03, 0.925, f"Mois de référence : {k.mois_ref:%m/%Y}  •  Source : export_ventes_brut.xlsx  •  "
             "aperçu reconstitué de l'onglet Synthèse", fontsize=9, style="italic", color=GREY, va="top")
    tiles = [
        (f"CA HT {k.mois_ref:%m/%Y}", fmt_eur(k.ca_mois), PRIMARY),
        ("Évol. vs mois précédent", fmt_pct(k.var_m1), GREEN if (k.var_m1 or 0) >= 0 else RED),
        ("Évol. vs même mois N-1", fmt_pct(k.var_n1), GREEN if (k.var_n1 or 0) >= 0 else RED),
        (f"CA HT cumulé {k.mois_ref.year}", fmt_eur(k.ca_cumul), PRIMARY),
        (f"Évol. cumul vs {k.mois_ref.year - 1}", fmt_pct(k.var_cumul), GREEN if (k.var_cumul or 0) >= 0 else RED),
        ("Panier moyen du mois", fmt_eur(k.panier_moyen_mois), PRIMARY),
        ("Factures du mois", str(k.nb_factures_mois), PRIMARY),
        ("Clients actifs (12 mois)", str(k.clients_actifs_12m), PRIMARY),
        ("Lignes à vérifier", str(q.n_anomalies), PRIMARY),
    ]
    w, h, gap = 0.148, 0.075, 0.008
    for i, (label, value, color) in enumerate(tiles):
        row, col = divmod(i, 6)
        x, y = 0.03 + col * (w + gap), 0.80 - row * (h + gap)
        fig.patches.append(plt.Rectangle((x, y), w, h, transform=fig.transFigure, facecolor=LIGHT,
                                         edgecolor="none"))
        fig.text(x + w / 2, y + h * 0.75, label, ha="center", va="center", fontsize=8.5, color=GREY)
        fig.text(x + w / 2, y + h * 0.32, value, ha="center", va="center", fontsize=15,
                 fontweight="bold", color=color)
    ax1 = fig.add_axes([0.08, 0.40, 0.88, 0.27])
    monthly_chart(ax1, m)
    ax1.set_title("Chiffre d'affaires HT mensuel", loc="left", fontsize=11, color=PRIMARY)
    ax2 = fig.add_axes([0.17, 0.05, 0.30, 0.27])
    ax2.barh(reg["region"], reg["ca_ht"], color=PRIMARY)
    ax2.invert_yaxis()
    ax2.xaxis.set_major_formatter(FuncFormatter(lambda x, _: f"{x / 1000:.0f} k€"))
    ax2.set_title("CA HT par région", loc="left", fontsize=11, color=PRIMARY)
    style(ax2)
    ax2.grid(axis="y", visible=False)
    ax2.grid(axis="x", color="#E5E5E5")
    ax3 = fig.add_axes([0.66, 0.05, 0.30, 0.27])
    ax3.barh(top["client"], top["ca_ht"], color=PRIMARY)
    ax3.invert_yaxis()
    ax3.xaxis.set_major_formatter(FuncFormatter(lambda x, _: f"{x / 1000:.0f} k€"))
    ax3.set_title("Top 5 clients", loc="left", fontsize=11, color=PRIMARY)
    style(ax3)
    ax3.grid(axis="y", visible=False)
    ax3.grid(axis="x", color="#E5E5E5")
    fig.savefig(OUT / "apercu_synthese.png")
    plt.close(fig)


def before_after_png(raw, df):
    lines = list(range(174, 184))  # lignes Excel 174 à 183 de l'export d'exemple (dont une ligne vide)
    before = raw.loc[[l - 2 for l in lines], ["Date", "Client", "Produit", "Région", "Montant HT"]]
    after = df[df["ligne"].isin(lines)][["ligne", "date", "client", "produit", "region", "montant_ht"]]

    def show(v):
        if v is None or (isinstance(v, float) and v != v):
            return "(vide)"
        if hasattr(v, "strftime"):
            return v.strftime("%d/%m/%Y") + " (date)"
        if isinstance(v, str):
            return f"«{v}»"
        return str(v)

    fig, (a, b) = plt.subplots(2, 1, figsize=(12, 6.2), dpi=130)
    for ax in (a, b):
        ax.axis("off")
    t1 = a.table(cellText=[[show(v) for v in r] for r in before.itertuples(index=False)],
                 colLabels=list(before.columns), loc="center", cellLoc="left")
    a.set_title("Avant : export brut (espaces, casse, dates et montants en texte, ligne vide)", loc="left",
                fontsize=11, color=RED, fontweight="bold")
    rows = [[f"{r.date:%d/%m/%Y}", r.client, r.produit, r.region, f"{r.montant_ht:.2f}".replace(".", ",")]
            for r in after.itertuples()]
    t2 = b.table(cellText=rows, colLabels=["Date", "Client", "Produit", "Région", "Montant HT"],
                 loc="center", cellLoc="left")
    b.set_title("Après : données nettoyées (chaque correction est tracée dans l’onglet Anomalies)", loc="left",
                fontsize=11, color=GREEN, fontweight="bold")
    for t in (t1, t2):
        t.auto_set_font_size(False)
        t.set_fontsize(8.5)
        t.scale(1, 1.25)
        for (r, _), cell in t.get_celld().items():
            cell.set_edgecolor("#D9D9D9")
            if r == 0:
                cell.set_facecolor(PRIMARY)
                cell.set_text_props(color="white", fontweight="bold")
    fig.tight_layout()
    fig.savefig(OUT / "avant_apres.png")
    plt.close(fig)


def main(path="data/export_ventes_brut.xlsx"):
    OUT.mkdir(parents=True, exist_ok=True)
    raw = read_export(path)
    df, q = clean(raw)
    chart_png(monthly(df))
    synthese_png(df, q)
    before_after_png(raw, df)
    print(f"Images écrites dans {OUT}/")


if __name__ == "__main__":
    main(*sys.argv[1:])
