"""Transforme un export de ventes Excel brut en classeur de reporting mis en forme.

Usage :
    python report.py data/export_ventes_brut.xlsx -o reporting.xlsx
    python report.py export.xlsx -o reporting.xlsx --mois 2026-06 --top 15 --societe "Ma PME"
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd

from reporting.aggregates import by_region, kpis, monthly, top_clients, top_products
from reporting.cleaning import clean, read_export
from reporting.excel_report import build_workbook


def run(input_path, output_path, mois=None, top=10, sheet=0, company="Société exemple"):
    raw = read_export(input_path, sheet)
    df, quality = clean(raw)
    if df.empty:
        raise ValueError("Aucune ligne exploitable après nettoyage : vérifiez le fichier et l'onglet lu.")
    if mois:
        ref = pd.Timestamp(mois).to_period("M").to_timestamp()
        df = df[df["mois"] <= ref]
    k = kpis(df, mois)
    wb = build_workbook(df, quality, monthly(df), by_region(df), top_clients(df, top),
                        top_products(df, top), k, Path(input_path).name, company)
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return df, quality, k


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("input", help="export Excel brut (.xlsx)")
    p.add_argument("-o", "--output", default="reporting.xlsx", help="classeur produit (défaut : reporting.xlsx)")
    p.add_argument("--mois", help="mois de référence AAAA-MM (défaut : dernier mois des données)")
    p.add_argument("--top", type=int, default=10, help="taille des tops clients / produits (défaut : 10)")
    p.add_argument("--onglet", default=0, help="nom ou index de l'onglet à lire (défaut : le premier)")
    p.add_argument("--societe", default="Société exemple", help="nom affiché dans le titre")
    a = p.parse_args(argv)
    sheet = int(a.onglet) if str(a.onglet).isdigit() else a.onglet
    try:
        df, q, k = run(a.input, a.output, a.mois, a.top, sheet, a.societe)
    except (ValueError, FileNotFoundError) as e:
        print(f"Erreur : {e}", file=sys.stderr)
        return 1
    print(f"Lignes lues : {q.rows_read} | retenues : {q.rows_kept} | à vérifier : {q.n_anomalies}")
    for label, n in q.corrections.items():
        print(f"  - {label} : {n}")
    print(f"Mois de référence {k.mois_ref:%m/%Y} : CA HT {k.ca_mois:,.0f} €".replace(",", " "))
    print(f"Reporting écrit dans {a.output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
