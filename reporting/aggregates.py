"""Agrégats du reporting : CA mensuel, régions, tops, indicateurs clés."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

import pandas as pd

from .cleaning import UNKNOWN_CLIENT


def _pct(new, old):
    if old is None or pd.isna(old) or old == 0 or new is None or pd.isna(new):
        return None
    return new / old - 1


def monthly(df: pd.DataFrame) -> pd.DataFrame:
    """CA HT par mois (mois sans vente inclus, à 0), nombre de factures, variations M-1 et N-1."""
    months = pd.date_range(df["mois"].min(), df["mois"].max(), freq="MS")
    g = df.groupby("mois").agg(ca_ht=("montant_ht", "sum"), nb_factures=("facture", "nunique"))
    g = g.reindex(months, fill_value=0)
    g.index.name = "mois"
    prev_m = g["ca_ht"].shift(1)
    prev_y = g["ca_ht"].shift(12)
    g["var_m1"] = [_pct(a, b) for a, b in zip(g["ca_ht"], prev_m)]
    g["var_n1"] = [_pct(a, b) for a, b in zip(g["ca_ht"], prev_y)]
    return g.reset_index()


def _share(df: pd.DataFrame, by: str, n: Optional[int] = None, extra=None) -> pd.DataFrame:
    agg = {"ca_ht": ("montant_ht", "sum")}
    if extra:
        agg.update(extra)
    g = df.groupby(by).agg(**agg).sort_values("ca_ht", ascending=False)
    g["part"] = g["ca_ht"] / df["montant_ht"].sum()
    if n:
        g = g.head(n)
    return g.reset_index()


def by_region(df: pd.DataFrame) -> pd.DataFrame:
    return _share(df, "region", extra={"nb_clients": ("client", "nunique")})


def top_clients(df: pd.DataFrame, n: int = 10) -> pd.DataFrame:
    known = df[df["client"] != UNKNOWN_CLIENT]
    out = _share(known, "client", n, extra={"nb_factures": ("facture", "nunique")})
    out["part"] = out["ca_ht"] / df["montant_ht"].sum()  # part du CA total, inconnus compris
    return out


def top_products(df: pd.DataFrame, n: int = 10) -> pd.DataFrame:
    return _share(df, "produit", n, extra={"quantite": ("quantite", "sum")})


@dataclass
class Kpis:
    mois_ref: pd.Timestamp
    ca_mois: float
    var_m1: Optional[float]
    var_n1: Optional[float]
    ca_cumul: float          # cumul depuis le 1er janvier jusqu'au mois de référence inclus
    ca_cumul_n1: float       # même période, année précédente
    var_cumul: Optional[float]
    nb_factures_mois: int
    panier_moyen_mois: Optional[float]
    clients_actifs_12m: int


def kpis(df: pd.DataFrame, mois_ref: Optional[pd.Timestamp] = None) -> Kpis:
    """Indicateurs du mois de référence (par défaut : le dernier mois présent dans les données)."""
    m = monthly(df).set_index("mois")
    ref = pd.Timestamp(mois_ref).to_period("M").to_timestamp() if mois_ref is not None else m.index.max()
    if ref not in m.index:
        raise ValueError(f"Mois de référence {ref:%m/%Y} absent des données "
                         f"({m.index.min():%m/%Y} à {m.index.max():%m/%Y}).")
    ca = float(m.loc[ref, "ca_ht"])
    ref_n1 = ref - pd.DateOffset(years=1)
    ytd = df[(df["mois"] >= pd.Timestamp(ref.year, 1, 1)) & (df["mois"] <= ref)]["montant_ht"].sum()
    ytd_n1 = df[(df["mois"] >= pd.Timestamp(ref.year - 1, 1, 1)) & (df["mois"] <= ref_n1)]["montant_ht"].sum()
    nb = int(m.loc[ref, "nb_factures"])
    last12 = df[(df["mois"] > ref - pd.DateOffset(months=12)) & (df["mois"] <= ref)]
    return Kpis(
        mois_ref=ref,
        ca_mois=ca,
        var_m1=m.loc[ref, "var_m1"],
        var_n1=m.loc[ref, "var_n1"],
        ca_cumul=float(ytd),
        ca_cumul_n1=float(ytd_n1),
        var_cumul=_pct(ytd, ytd_n1),
        nb_factures_mois=nb,
        panier_moyen_mois=ca / nb if nb else None,
        clients_actifs_12m=int(last12.loc[last12["client"] != UNKNOWN_CLIENT, "client"].nunique()),
    )
