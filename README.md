# Transactions — application Android hors ligne

Lit votre classeur Excel (.xlsx / .xlsm), filtre chaque colonne indépendamment, calcule des totaux regroupables et affiche un histogramme. Aucune permission, aucun accès Internet.

## Fichiers

| Fichier | Rôle |
|---|---|
| `main.py` | Interface Kivy (onglets Transactions / Totaux / Graphique) |
| `logic.py` | Lecture Excel, filtres, regroupements (sans Kivy) |
| `test_logic.py` | Tests de la logique : `python test_logic.py` |
| `buildozer.spec` | Compilation, variante A (Python 3.11 imposé) |
| `buildozer_b.spec` | Compilation, variante B (secours : python-for-android figé à janvier 2024) |
| `.github/workflows/build-apk.yml` | Compile les deux variantes sur GitHub |

## Compiler

1. Envoyez ces fichiers à la racine du dépôt GitHub (Add file → Upload files ; les fichiers de même nom sont remplacés).
2. Ouvrez `.github/workflows/build-apk.yml` (crayon) et remplacez son contenu par celui fourni.
3. Onglet **Actions** : deux compilations démarrent en parallèle, « Variante A » et « Variante B » (20 à 40 min).
4. Téléchargez l'artefact `transactions-apk-variante-A` (ou B si A échoue) : un zip contenant le `.apk`.
5. En cas d'échec : ouvrez la variante, l'étape **« Résumé de l'erreur »** affiche l'essentiel à copier.

## Utilisation

- **Ouvrir Excel** : choisissez votre fichier. L'appli en garde une copie privée et la recharge au lancement suivant. Pour mettre à jour les données, rouvrez la dernière version du fichier.
- **Feuille** : l'onglet `2014_HB_KBC` est ouvert automatiquement ; la liste permet d'en changer. Les en-têtes doivent être sur la première ligne.
- **Colonnes** : choisissez celles à afficher (mémorisé). Le détail complet d'une ligne s'ouvre en la touchant.
- **Filtres** : sous chaque en-tête, un champ de saisie et un bouton ▼ (liste de valeurs à cocher, avec recherche). Tous les filtres se combinent (ET).

| Colonne | Saisie libre |
|---|---|
| Texte | `cafe` (contient) · `salaire\|loyer` (ou) · `!salaire` (exclut) |
| Nombre | `>50` · `<=-10` · `!=0` · `10..20` · `..100` |
| Date | `2026` · `2026-03` · `>=2026-01-15` · `2026-01-01..2026-03-31` |

- **Période rapide** : Ce mois, Mois dernier, Cette année, Tout (sur la colonne date choisie dans Totaux).
- **Totaux** : trois niveaux de regroupement au choix (Période, Catégorie, Sous Catégorie, ou n'importe quelle colonne). Touchez une ligne pour déplier ou replier ; « Tout déplier / Tout replier » ; mettez un niveau sur « — » pour dégrouper. Chaque ligne donne Nb, Net, Entrées, Sorties. Les totaux portent sur les lignes **filtrées**.
- **Graphique** : entrées (vert) et sorties (rouge) par période ; touchez une barre pour le détail.

## Limites

- Formats `.xlsx` / `.xlsm` seulement. Les macros ne sont pas exécutées.
- Les formules sont lues avec leur dernier résultat enregistré par Excel (enregistrez le fichier dans Excel avant de le copier).
- APK de debug, non signé pour le Play Store : suffisant pour un usage personnel.
- Si l'appli rencontre une erreur, elle l'affiche à l'écran (faites une capture) et l'écrit dans `error.log`.
