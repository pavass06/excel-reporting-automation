import pandas as pd
import pytest

from reporting.aggregates import by_region, kpis, monthly, top_clients, top_products
from reporting.cleaning import UNKNOWN_CLIENT


def sales():
    rows = [
        # date, facture, client, produit, région, quantité, HT
        ("2024-09-10", "F1", "A", "P1", "Nord", 1, 100.0),
        ("2025-01-15", "F2", "A", "P1", "Nord", 2, 200.0),
        ("2025-08-05", "F3", "B", "P2", "Sud", 1, 400.0),
        ("2025-08-20", "F3", "B", "P1", "Sud", 1, 100.0),   # même facture, deux lignes
        ("2025-09-02", "F4", "C", "P2", "Sud", 1, 150.0),
        ("2025-09-03", "F5", UNKNOWN_CLIENT, "P2", "Nord", 1, 50.0),
    ]
    df = pd.DataFrame(rows, columns=["date", "facture", "client", "produit", "region", "quantite", "montant_ht"])
    df["date"] = pd.to_datetime(df["date"])
    df["mois"] = df["date"].dt.to_period("M").dt.to_timestamp()
    return df


def test_monthly_fills_missing_months_and_computes_variations():
    m = monthly(sales()).set_index("mois")
    assert len(m) == 13                                   # sept. 2024 -> sept. 2025, mois vides inclus
    assert m.loc["2024-10-01", "ca_ht"] == 0
    assert m.loc["2025-08-01", "ca_ht"] == 500
    assert m.loc["2025-08-01", "nb_factures"] == 1        # factures distinctes, pas lignes
    assert m.loc["2025-09-01", "var_m1"] == pytest.approx(200 / 500 - 1)
    assert m.loc["2025-09-01", "var_n1"] == pytest.approx(200 / 100 - 1)
    assert pd.isna(m.loc["2024-11-01", "var_m1"])          # pas de variation depuis un mois à 0


def test_region_shares_sum_to_one():
    r = by_region(sales())
    assert r["part"].sum() == pytest.approx(1)
    assert r.iloc[0]["region"] == "Sud" and r.iloc[0]["ca_ht"] == 650


def test_top_clients_excludes_unknown_client_but_keeps_share_of_total():
    t = top_clients(sales(), 2)
    assert t["client"].tolist() == ["B", "A"]
    assert UNKNOWN_CLIENT not in t["client"].tolist()
    assert t.iloc[0]["part"] == pytest.approx(500 / 1000)


def test_top_products():
    t = top_products(sales(), 5)
    assert t.set_index("produit").loc["P2", "ca_ht"] == 600
    assert t.set_index("produit").loc["P1", "quantite"] == 4


def test_kpis_reference_month():
    k = kpis(sales())
    assert k.mois_ref == pd.Timestamp("2025-09-01")
    assert k.ca_mois == 200
    assert k.ca_cumul == 200 + 500 + 200                   # janv. -> sept. 2025
    assert k.ca_cumul_n1 == 100                            # janv. -> sept. 2024
    assert k.var_cumul == pytest.approx(900 / 100 - 1)
    assert k.nb_factures_mois == 2 and k.panier_moyen_mois == 100
    assert k.clients_actifs_12m == 3                       # A, B, C (client inconnu exclu)


def test_kpis_rejects_month_outside_data():
    with pytest.raises(ValueError, match="absent"):
        kpis(sales(), "2023-01")
