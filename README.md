# Classement Tarot - Agrégateur de scores

Ce petit outil lit un fichier Excel contenant plusieurs feuilles (une feuille = un tournoi), agrège les scores par joueur, extrait leurs 15 meilleurs scores, calcule un total, détermine un classement, puis exporte le résultat en CSV et/ou en PDF (paysage).

## Prérequis

- Python 3.13 (ou compatible avec votre environnement `env/` déjà présent)
- Dépendances Python listées dans `requirements.txt`

Installez les dépendances dans votre environnement virtuel:

```bash
env\Scripts\pip.exe install -r requirements.txt
```

## Utilisation

```bash
env\Scripts\python.exe tarot_rankings.py <chemin_vers_excel.xlsx> --out sorties --pdf --csv
```

Options:
- `--out` : dossier de sortie (défaut: `.`)
- `--pdf` : génère un PDF paysage
- `--csv` : génère un CSV

Le fichier Excel doit contenir toutes les feuilles de tournois partageant la même structure. Pour chaque feuille, les lignes utiles sont 4 à 100 (incluses), avec:
- colonne C: nom de famille
- colonne D: prénom
- colonne I: score de la journée

Les règles appliquées:
- Un joueur peut être absent sur certaines feuilles
- Un joueur sans participation n'apparaît pas dans la sortie
- Si un joueur a moins de 15 participations, on remplit juste avec moins de colonnes renseignées

## Sorties

- CSV: `classement_tarot.csv`
- PDF: `classement_tarot.pdf` (format paysage, tableau multi‑pages si nécessaire)

Les colonnes sont, dans l'ordre: `Nom`, `Prénom`, 15 colonnes de meilleurs scores, `Total`, `Classement`.

# classement-tarot