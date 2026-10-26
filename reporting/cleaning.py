"""Nettoyage d'un export de ventes brut et contrôle qualité.

Principe : on ne supprime jamais une ligne en silence. Chaque correction ou exclusion
est tracée dans le rapport d'anomalies, avec le numéro de ligne du fichier d'origine.
"""
from __future__ import annotations

import re
import unicodedata
from collections import Counter
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta

import numpy as np
import pandas as pd

# ---------------------------------------------------------------------------
# Paramètres à adapter au fichier du client
# ---------------------------------------------------------------------------
# Colonnes attendues : nom interne -> intitulés acceptés dans l'export (comparés sans
# accents, sans casse, sans espaces). Ajouter ici les intitulés propres à chaque ERP.
COLUMNS = {
    "date": ["date", "date commande", "date facture", "date de vente"],
    "facture": ["n° facture", "no facture", "numero facture", "facture", "n facture"],
    "client": ["client", "raison sociale", "nom client"],
    "produit": ["produit", "article", "designation", "libelle article"],
    "region": ["region", "zone", "secteur"],
    "commercial": ["commercial", "vendeur", "representant"],
    "quantite": ["quantite", "qte", "qty"],
    "montant_ht": ["montant ht", "total ht", "ca ht"],
    "tva": ["tva", "montant tva"],
    "montant_ttc": ["montant ttc", "total ttc"],
}
REQUIRED = ["date", "client", "produit", "montant_ht"]          # indispensables au nettoyage
REPORT_REQUIRED = REQUIRED + ["facture", "region", "quantite"]  # indispensables au reporting complet
TEXT_COLUMNS = ["facture", "client", "produit", "region", "commercial"]
AMOUNT_COLUMNS = ["montant_ht", "tva", "montant_ttc"]
DATE_FORMATS = ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d-%m-%y", "%d.%m.%Y", "%d/%m/%y", "%Y/%m/%d"]
TTC_TOLERANCE = 0.05  # écart toléré (en euros) entre HT + TVA et TTC
UNKNOWN_CLIENT = "(client non renseigné)"

# Intitulés des colonnes dans le classeur produit
LABELS = {
    "date": "Date", "mois": "Mois", "facture": "N° facture", "client": "Client",
    "produit": "Produit", "region": "Région", "commercial": "Commercial",
    "quantite": "Quantité", "montant_ht": "Montant HT", "tva": "TVA",
    "montant_ttc": "Montant TTC",
}


# ---------------------------------------------------------------------------
# Fonctions élémentaires (testées unitairement)
# ---------------------------------------------------------------------------
def strip_accents(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))


def key(s) -> str:
    """Clé de comparaison : sans accents, sans casse, espaces réduits."""
    return re.sub(r"\s+", " ", strip_accents(str(s))).strip().casefold()


def normalize_spaces(v):
    """Supprime les espaces en début/fin et réduit les espaces multiples. None si vide."""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return None
    s = re.sub(r"\s+", " ", str(v).replace(" ", " ")).strip()
    return s or None


def parse_date(v):
    """Convertit une date Excel, un numéro de série ou un texte (formats FR et ISO) en `date`.

    Renvoie None si la valeur est vide ou illisible (ex. « 31/02/2025 »).
    """
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return None
    if isinstance(v, pd.Timestamp):
        return None if pd.isna(v) else v.date()
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, (int, float, np.integer, np.floating)):
        if 20000 <= float(v) <= 80000:  # numéro de série Excel (1954 - 2119)
            return (datetime(1899, 12, 30) + timedelta(days=float(v))).date()
        return None
    s = str(v).strip()
    if not s:
        return None
    s = s.split(" ")[0]  # « 15/03/2025 00:00:00 » -> « 15/03/2025 »
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def parse_amount(v):
    """Convertit un montant numérique ou texte (« 1 234,50 € », « 1234.5 ») en float.

    Renvoie None si la valeur est vide ou illisible.
    """
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float, np.integer, np.floating)):
        return None if np.isnan(v) else float(v)
    s = str(v)
    for ch in ("€", "EUR", " ", " ", " ", "'"):
        s = s.replace(ch, "")
    s = s.strip()
    if not s:
        return None
    if "," in s and "." in s:  # le dernier séparateur est le séparateur décimal
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    else:
        s = s.replace(",", ".")
    if s.startswith("(") and s.endswith(")"):  # convention comptable (123,45) = négatif
        s = "-" + s[1:-1]
    try:
        return float(s)
    except ValueError:
        return None


def canonical_forms(values) -> dict:
    """Pour chaque groupe de variantes (« ATELIER MOREAU », « atelier moreau »...),
    choisit la graphie la plus fréquente, en préférant une casse mixte en cas d'égalité."""
    groups: dict = {}
    for v, n in Counter(v for v in values if v).items():
        groups.setdefault(key(v), Counter())[v] += n

    def score(item):
        v, n = item
        mixed = not (v.isupper() or v.islower())
        return (n, mixed)

    return {k: max(c.items(), key=score)[0] for k, c in groups.items()}


def map_columns(df: pd.DataFrame) -> dict:
    """Associe les colonnes de l'export aux noms internes. Lève une erreur claire si une
    colonne obligatoire manque."""
    found = {}
    for col in df.columns:
        k = key(col)
        for name, aliases in COLUMNS.items():
            if name not in found and k in {key(a) for a in aliases}:
                found[name] = col
    missing = [c for c in REQUIRED if c not in found]
    if missing:
        raise ValueError(
            f"Colonnes obligatoires introuvables : {missing}. Colonnes lues : {list(df.columns)}. "
            "Ajoutez l'intitulé correspondant dans COLUMNS (reporting/cleaning.py)."
        )
    return found


# ---------------------------------------------------------------------------
# Nettoyage complet
# ---------------------------------------------------------------------------
@dataclass
class QualityReport:
    anomalies: pd.DataFrame
    corrections: dict = field(default_factory=dict)
    rows_read: int = 0
    rows_kept: int = 0

    @property
    def n_anomalies(self) -> int:
        return len(self.anomalies)


def clean(raw: pd.DataFrame) -> tuple:
    """Nettoie un export brut. Renvoie (données propres, QualityReport)."""
    mapping = map_columns(raw)
    df = pd.DataFrame({name: raw[col] for name, col in mapping.items()})
    df.insert(0, "ligne", raw.index + 2)  # numéro de ligne Excel (en-tête en ligne 1)
    anomalies = []
    corr = Counter()

    def flag(rows, kind, column, values, action):
        for r, v in zip(rows, values):
            anomalies.append({"Ligne source": int(r), "Type": kind, "Colonne": column,
                              "Valeur d'origine": "" if v is None else str(v), "Action": action})

    rows_read = len(df)

    # 1. Lignes entièrement vides
    empty = df.drop(columns="ligne").isna().all(axis=1) | df.drop(columns="ligne").apply(
        lambda r: all(normalize_spaces(v) is None for v in r), axis=1)
    flag(df.loc[empty, "ligne"], "Ligne vide", "", [None] * empty.sum(), "Supprimée")
    df = df[~empty].copy()

    # 2. Texte : espaces parasites puis casse harmonisée
    for c in [c for c in TEXT_COLUMNS if c in df]:
        original = df[c]
        stripped = original.map(normalize_spaces)
        corr["Espaces superflus supprimés"] += sum(
            isinstance(o, str) and o != (s or "") for o, s in zip(original, stripped))
        forms = canonical_forms(stripped)
        harmonised = stripped.map(lambda v: forms.get(key(v)) if v else None)
        corr["Casse / orthographe harmonisée"] += int((harmonised.fillna("") != stripped.fillna("")).sum())
        df[c] = harmonised

    # 3. Dates
    parsed = df["date"].map(parse_date)
    corr["Dates converties depuis du texte"] += int(
        (df["date"].map(lambda v: isinstance(v, str)) & parsed.notna()).sum())
    bad = parsed.isna()
    flag(df.loc[bad, "ligne"], "Date illisible ou impossible", LABELS["date"], df.loc[bad, "date"],
         "Ligne exclue des calculs")
    df["date"] = parsed
    df = df[~bad].copy()

    # 4. Montants
    for c in [c for c in AMOUNT_COLUMNS if c in df]:
        parsed = df[c].map(parse_amount)
        corr["Montants convertis depuis du texte"] += int(
            (df[c].map(lambda v: isinstance(v, str)) & parsed.notna()).sum())
        bad = parsed.isna() & df[c].notna()
        if c == "montant_ht":
            bad = parsed.isna()
            flag(df.loc[bad, "ligne"], "Montant HT illisible", LABELS[c], df.loc[bad, c],
                 "Ligne exclue des calculs")
            df[c] = parsed
            df = df[~bad].copy()
        else:
            flag(df.loc[bad, "ligne"], f"{LABELS[c]} illisible", LABELS[c], df.loc[bad, c],
                 "Laissé vide, à vérifier")
            df[c] = parsed
        df[c] = df[c].astype(float)
    if "quantite" in df:
        df["quantite"] = pd.to_numeric(df["quantite"].map(parse_amount), errors="coerce")

    # 5. Client manquant : la vente compte dans le CA, mais pas dans le top clients
    missing = df["client"].isna()
    flag(df.loc[missing, "ligne"], "Client manquant", LABELS["client"], [None] * missing.sum(),
         f"Conservée sous « {UNKNOWN_CLIENT} », à compléter")
    df.loc[missing, "client"] = UNKNOWN_CLIENT

    # 6. Cohérence HT + TVA = TTC
    if {"tva", "montant_ttc"} <= set(df.columns):
        gap = (df["montant_ht"] + df["tva"] - df["montant_ttc"]).abs()
        bad = gap > TTC_TOLERANCE
        flag(df.loc[bad, "ligne"], "Incohérence HT + TVA ≠ TTC", LABELS["montant_ttc"],
             df.loc[bad, "montant_ttc"], "Conservée (calculs sur le HT), à vérifier")

    # 7. Doublons (après nettoyage, pour attraper aussi ceux qui ne diffèrent que par la forme)
    business = [c for c in df.columns if c != "ligne"]
    dup = df.duplicated(subset=business, keep="first")
    first_line = df.groupby(business, dropna=False)["ligne"].transform("first")
    flag(df.loc[dup, "ligne"], "Doublon", "", [f"identique à la ligne {n}" for n in first_line[dup]],
         "Supprimée")
    df = df[~dup].copy()

    df["mois"] = pd.to_datetime(df["date"]).dt.to_period("M").dt.to_timestamp()
    df["date"] = pd.to_datetime(df["date"])
    df = df.sort_values(["date", "ligne"]).reset_index(drop=True)

    anomalies_df = pd.DataFrame(anomalies, columns=["Ligne source", "Type", "Colonne",
                                                    "Valeur d'origine", "Action"])
    anomalies_df = anomalies_df.sort_values(["Ligne source", "Type"]).reset_index(drop=True)
    report = QualityReport(anomalies_df, {k: v for k, v in corr.items() if v}, rows_read, len(df))
    return df, report


def read_export(path, sheet=0) -> pd.DataFrame:
    """Lit l'export sans laisser pandas deviner les types (on veut voir les valeurs brutes)."""
    return pd.read_excel(path, sheet_name=sheet, dtype=object)
