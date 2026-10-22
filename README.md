# Rapprochement bancaire

Logiciel Windows (.exe) basé sur le classeur *ETAT DE RAPPROCHEMENT BANCAIRE.xlsx* : mêmes couleurs, mêmes formules.

## Pages
1. **Page d'accueil** : état de rapprochement (nos livres / livres de la banque), sous-totaux, solde rapproché, totaux et contrôles obligatoires. Saisie de l'en-tête et des soldes reportés.
2. **Import du grand livre banque** : importer un grand livre (.xlsx / .csv avec colonnes Date, Libellé, Débit, Crédit) et pointer chaque écriture. Les écritures **non pointées** alimentent « Compte banque dans nos livres ».
3. **Saisie relevé bancaire non pointé** : saisie ligne par ligne des opérations du relevé absentes du grand livre. Elles alimentent « Compte dans les livres de la banque ».

Menu *Fichier* : enregistrer / ouvrir un dossier (.rapp) et **exporter l'état vers Excel** (avec les formules d'origine) pour l'imprimer et le signer.

## Obtenir le .exe (GitHub)
1. Créez un dépôt GitHub et poussez ces fichiers (branche `main`).
2. Onglet **Actions** → workflow *Build Windows EXE* → à la fin, téléchargez l'artefact `RapprochementBancaire-windows`.
3. Pour une release : `git tag v1.0 && git push --tags` (le .exe est joint à la release).

## Lancer en local
```
pip install -r requirements.txt
python app.py
```
