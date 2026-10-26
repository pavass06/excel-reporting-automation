from datetime import date, datetime

import pandas as pd
import pytest

from reporting.cleaning import (UNKNOWN_CLIENT, canonical_forms, clean, map_columns, normalize_spaces,
                                parse_amount, parse_date)


@pytest.mark.parametrize("value, expected", [
    (datetime(2025, 3, 15, 10, 30), date(2025, 3, 15)),
    (date(2025, 3, 15), date(2025, 3, 15)),
    (pd.Timestamp("2025-03-15"), date(2025, 3, 15)),
    ("15/03/2025", date(2025, 3, 15)),
    ("2025-03-15", date(2025, 3, 15)),
    ("15-03-25", date(2025, 3, 15)),
    ("15.03.2025", date(2025, 3, 15)),
    (" 15/03/2025 00:00:00", date(2025, 3, 15)),
    (45731, date(2025, 3, 15)),          # numéro de série Excel
    ("31/02/2025", None),                # date impossible
    ("bientôt", None),
    ("", None),
    (None, None),
    (float("nan"), None),
])
def test_parse_date(value, expected):
    assert parse_date(value) == expected


def test_parse_date_is_day_first():
    # 03/04/2025 est le 3 avril en France, pas le 4 mars
    assert parse_date("03/04/2025") == date(2025, 4, 3)


@pytest.mark.parametrize("value, expected", [
    (1234.5, 1234.5),
    (12, 12.0),
    ("1234.50", 1234.5),
    ("1234,50", 1234.5),
    ("1 234,50 €", 1234.5),
    ("1 234,50 €", 1234.5),   # espaces insécables (copier-coller depuis un PDF ou Excel FR)
    ("1.234,50", 1234.5),
    ("1,234.50", 1234.5),
    ("(120,00)", -120.0),
    ("-45,10 €", -45.1),
    ("à voir", None),
    ("", None),
    (None, None),
    (float("nan"), None),
])
def test_parse_amount(value, expected):
    assert parse_amount(value) == expected


def test_normalize_spaces():
    assert normalize_spaces("  Atelier   Moreau  ") == "Atelier Moreau"
    assert normalize_spaces("   ") is None
    assert normalize_spaces(None) is None


def test_canonical_forms_prefers_most_frequent_then_mixed_case():
    forms = canonical_forms(["Île-de-France"] * 3 + ["ÎLE-DE-FRANCE", "île-de-france", "Ile-de-France"])
    assert forms["ile-de-france"] == "Île-de-France"
    assert canonical_forms(["BRETAGNE", "Bretagne"])["bretagne"] == "Bretagne"


def test_map_columns_accepts_aliases_and_reports_missing():
    df = pd.DataFrame(columns=["Date facture", "Raison sociale", "Désignation", "Total HT"])
    assert map_columns(df) == {"date": "Date facture", "client": "Raison sociale",
                               "produit": "Désignation", "montant_ht": "Total HT"}
    with pytest.raises(ValueError, match="montant_ht"):
        map_columns(pd.DataFrame(columns=["Date", "Client", "Produit"]))


def raw_export():
    cols = ["Date", "N° facture", "Client", "Produit", "Région", "Commercial", "Quantité",
            "Montant HT", "TVA", "Montant TTC"]
    rows = [
        [datetime(2025, 1, 10), "F1", "Atelier Moreau", "Lampe", "Bretagne", "Julie", 1, 100.0, 20.0, 120.0],
        ["10/01/2025", "F1", "  ATELIER MOREAU ", "Lampe", "bretagne", "Julie", 1, "100,00 €", "20,00", 120.0],
        [None] * 10,
        ["31/02/2025", "F2", "Atelier Moreau", "Lampe", "Bretagne", "Julie", 1, 50.0, 10.0, 60.0],
        ["2025-02-03", "F3", "   ", "Bureau", "Bretagne", "Julie", 2, "1 000,00", 200.0, 1200.0],
        ["2025-02-04", "F4", "Cave Roux", "Bureau", "Occitanie", "Karim", 1, "à voir", 0.0, 0.0],
        ["2025-02-05", "F5", "Cave Roux", "Bureau", "Occitanie", "Karim", 1, 500.0, 100.0, 999.0],
        ["2025-02-05", "F6", "Atelier Moreau", "Bureau", "Bretagne", "Julie", 1, 500.0, 100.0, 600.0],
    ]
    return pd.DataFrame(rows, columns=cols, dtype=object)


def test_clean_end_to_end_on_small_export():
    df, q = clean(raw_export())
    types = dict(zip(q.anomalies["Ligne source"], q.anomalies["Type"]))
    # numéros de ligne Excel : en-tête en ligne 1, première donnée en ligne 2
    assert types[3] == "Doublon"              # identique à la ligne 2 une fois nettoyée
    assert types[4] == "Ligne vide"
    assert types[5] == "Date illisible ou impossible"
    assert types[6] == "Client manquant"
    assert types[7] == "Montant HT illisible"
    assert types[8] == "Incohérence HT + TVA ≠ TTC"
    assert q.rows_read == 8 and q.rows_kept == 4
    assert list(df["ligne"]) == [2, 6, 8, 9]
    assert df["montant_ht"].tolist() == [100.0, 1000.0, 500.0, 500.0]
    assert df.loc[df["ligne"] == 6, "client"].item() == UNKNOWN_CLIENT
    assert set(df["region"]) == {"Bretagne", "Occitanie"}
    assert df["mois"].dt.strftime("%Y-%m").tolist() == ["2025-01", "2025-02", "2025-02", "2025-02"]


def test_clean_counts_automatic_corrections():
    _, q = clean(raw_export())
    assert q.corrections["Montants convertis depuis du texte"] >= 3
    assert q.corrections["Dates converties depuis du texte"] >= 4
    assert q.corrections["Espaces superflus supprimés"] >= 1
