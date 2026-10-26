"""Génère un export de ventes FICTIF, volontairement « sale », pour la démonstration.

Aucune donnée réelle : clients, produits et commerciaux sont inventés.
Les défauts injectés sont ceux que l'on rencontre dans les exports d'ERP / logiciels
de caisse des PME : formats de date mélangés, doublons, espaces parasites, casse
incohérente, lignes vides, montants saisis en texte.

Usage :
    python generate_sample.py                      # -> data/export_ventes_brut.xlsx
    python generate_sample.py -o autre.xlsx --seed 7
"""
from __future__ import annotations

import argparse
import random
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook

CLIENTS = [
    "Boulangerie Martin", "Garage des Tilleuls", "Cabinet Leroy", "Hôtel du Parc",
    "Pharmacie Centrale", "Menuiserie Dubois", "Atelier Moreau", "Bistrot Le Comptoir",
    "Clinique Saint-Roch", "Transports Garnier", "Librairie Lefèvre", "Camping Les Pins",
    "Fromagerie Bernard", "Agence Horizon", "Imprimerie Petit", "Domaine Roux",
    "Salon Élégance", "Ferme Fontaine", "Cave Chevalier", "Opticien Girard",
]
PRODUITS = {  # produit -> prix unitaire HT
    "Ramette papier A4": 4.90, "Cartouche encre noire": 32.00, "Cartouche encre couleur": 41.50,
    "Fauteuil ergonomique": 289.00, "Bureau assis-debout": 549.00, "Lampe LED bureau": 39.90,
    "Classeur à levier": 3.20, "Destructeur de documents": 159.00, "Écran 27 pouces": 229.00,
    "Clavier sans fil": 34.90, "Casque audio pro": 79.00, "Tableau blanc 120x90": 119.00,
}
REGIONS = ["Île-de-France", "Auvergne-Rhône-Alpes", "Nouvelle-Aquitaine", "Occitanie",
           "Hauts-de-France", "Grand Est", "Bretagne"]
COMMERCIAUX = ["Sophie Laurent", "Karim Benali", "Julie Mercier", "Thomas Faure"]
TVA = 0.20


def _messy_date(d: date, rng: random.Random):
    """Renvoie la date sous un format choisi au hasard (comme un export mal maîtrisé)."""
    r = rng.random()
    if r < 0.55:
        return d  # vraie date Excel
    if r < 0.75:
        return d.strftime("%d/%m/%Y")
    if r < 0.85:
        return d.strftime("%Y-%m-%d")
    if r < 0.93:
        return d.strftime("%d-%m-%y")
    return d.strftime("%d.%m.%Y")


def _messy_text(s: str, rng: random.Random) -> str:
    r = rng.random()
    if r < 0.10:
        s = s.upper()
    elif r < 0.18:
        s = s.lower()
    if rng.random() < 0.12:
        s = "  " + s
    if rng.random() < 0.12:
        s = s + "   "
    if rng.random() < 0.05:
        s = s.replace(" ", "  ", 1)
    return s


def _messy_amount(x: float, rng: random.Random):
    r = rng.random()
    if r < 0.75:
        return round(x, 2)
    if r < 0.85:  # format français avec séparateur de milliers et symbole
        txt = f"{x:,.2f}".replace(",", " ").replace(".", ",")
        return txt + " €"
    if r < 0.93:
        return f"{x:.2f}".replace(".", ",")
    return f"{x:.2f}"


def build_rows(seed: int = 42, start=date(2024, 1, 1), end=date(2026, 9, 30)):
    rng = random.Random(seed)
    # chaque client a une région et un commercial attitrés, et un « poids » (gros / petits clients)
    profils = {
        c: (rng.choice(REGIONS), rng.choice(COMMERCIAUX), rng.uniform(0.6, 2.2) ** 2)
        for c in CLIENTS
    }
    clients, poids = zip(*[(c, p[2]) for c, p in profils.items()])
    rows = []
    facture = 24000
    d = start
    while d <= end:
        if d.weekday() < 5:
            # légère croissance dans le temps + saisonnalité (creux en août, pic en septembre)
            months = (d.year - start.year) * 12 + d.month - 1
            saison = {8: 0.35, 9: 1.35, 12: 0.8}.get(d.month, 1.0)
            n = int(rng.gauss(2.2 * saison * (1 + 0.012 * months), 1))
            for _ in range(max(n, 0)):
                facture += 1
                client = rng.choices(clients, weights=poids)[0]
                region, commercial, _ = profils[client]
                produit = rng.choice(list(PRODUITS))
                qte = rng.choice([1, 1, 1, 2, 2, 3, 5, 10, 20]) if PRODUITS[produit] < 50 else rng.choice([1, 1, 2, 3])
                ht = round(PRODUITS[produit] * qte, 2)
                tva = round(ht * TVA, 2)
                rows.append([
                    _messy_date(d, rng), f"F{facture}", _messy_text(client, rng),
                    _messy_text(produit, rng), _messy_text(region, rng), _messy_text(commercial, rng),
                    qte, _messy_amount(ht, rng), _messy_amount(tva, rng), round(ht + tva, 2),
                ])
        d += timedelta(days=1)

    # quelques erreurs de saisie franches, qui doivent ressortir dans le rapport d'anomalies
    rows[15][0] = "31/02/2025"          # date impossible
    rows[40][7] = "à voir"              # montant illisible
    rows[75][2] = "   "                 # client manquant
    rows[120][9] = 999.99               # TTC incohérent avec HT + TVA
    # doublons : une ligne exportée deux fois (fréquent lors d'extractions successives)
    for i in rng.sample(range(len(rows)), k=max(len(rows) // 60, 1)):
        rows.insert(i + 1, list(rows[i]))
    # lignes vides au milieu du fichier
    for i in rng.sample(range(len(rows)), k=6):
        rows.insert(i, [None] * 10)
    return rows


HEADERS = ["Date", "N° facture", "Client", "Produit", "Région", "Commercial",
           "Quantité", "Montant HT", "TVA", "Montant TTC"]


def write_xlsx(rows, path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Export"
    ws.append(HEADERS)
    for r in rows:
        ws.append(r)
    for cell in ws["A"][1:]:
        if not isinstance(cell.value, str):
            cell.number_format = "DD/MM/YYYY"
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("-o", "--output", default="data/export_ventes_brut.xlsx")
    p.add_argument("--seed", type=int, default=42)
    a = p.parse_args(argv)
    rows = build_rows(a.seed)
    write_xlsx(rows, Path(a.output))
    print(f"{len(rows)} lignes écrites dans {a.output}")


if __name__ == "__main__":
    main()
