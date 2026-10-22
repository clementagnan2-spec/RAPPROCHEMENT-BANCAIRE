# Rapprochement bancaire

Logiciel Windows (.exe) basé sur le classeur *ETAT DE RAPPROCHEMENT BANCAIRE.xlsx* : mêmes couleurs, mêmes formules.

## Pages
1. **Page d'accueil** : état de rapprochement (nos livres / livres de la banque), sous-totaux, solde rapproché, totaux et contrôles obligatoires. Saisie de l'en-tête et des soldes reportés.
2. **Import du grand livre banque** : importer un grand livre (.xlsx / .csv avec colonnes Date, Libellé, Débit, Crédit) et pointer chaque écriture. Les écritures **non pointées** alimentent « Compte banque dans nos livres ».
3. **Saisie relevé bancaire non pointé** : saisie ligne par ligne des opérations du relevé absentes du grand livre. Elles alimentent « Compte dans les livres de la banque ».

## Import cumulatif et modification des écritures
- Chaque import **s'ajoute à la suite** des précédents (janvier, puis février…) : rien n'est écrasé et les pointages déjà faits sont conservés. Une écriture strictement identique (date, libellé, montants) déjà présente est ignorée, pour éviter les doublons en cas de ré-import d'un même fichier.
- **Modifier une écriture** (date, libellé, montant) : double-clic sur la ligne ou bouton « Modifier la ligne… ». Les lignes modifiées sont marquées « (modifiée) » et l'ancienne valeur est conservée dans le journal. **Supprimer la sélection** (ou touche Suppr) retire des lignes. Ces actions sont bloquées quand la période est validée.

## Plusieurs comptes bancaires
- Barre **Banque / compte** en haut de la fenêtre : choisissez la banque dont vous éditez l'état de rapprochement. Chaque banque a **son propre** état : en-tête (code, compte, N°), soldes reportés, grand livre importé, pointages, relevé saisi, validation et journal.
- **+ Ajouter une banque** (ou menu *Rapprochement*) : crée un nouveau compte (société, période et mot de passe repris du compte courant). Renseignez ensuite le compte, le N° de compte et les soldes reportés sur la page d'accueil. **Supprimer cette banque** retire le compte affiché (impossible s'il est validé ou s'il est le seul).
- **Récapitulatif…** : une ligne par banque (équilibré ou non, validé ou non).
- La **société** est commune à toutes les banques ; le **mot de passe** de déverrouillage aussi.
- **Imprimer toutes les banques** : une page A4 par banque. **Exporter vers Excel** : un classeur avec une feuille par banque.
- **Nouveau (rapprochement suivant)** crée le mois suivant pour **toutes** les banques (suspens et lignes de relevé reportés). Il exige que chaque banque soit validée ; sinon la liste des banques à valider s'affiche.
- Les anciens fichiers `.rapp` (un seul compte) s'ouvrent normalement et sont convertis automatiquement.

## Impression, modèle d'import, validation
- **Imprimer l'état** (bouton de la page d'accueil, ou Ctrl+P) : ouvre l'état en couleurs, au format A4 paysage avec zones de signature, et lance la boîte d'impression du navigateur (choisissez aussi « Enregistrer au format PDF »).
- **Modèle d'importation** (`modele_import_grand_livre.xlsx`, aussi via le bouton « Modèle d'import… ») : feuille *Grand livre* à remplir, feuille *Exemple* et feuille *Instructions*.
- **Valider le rapprochement** : possible uniquement si les 4 contrôles obligatoires sont à « OK ». La validation **verrouille** tous les pointages, les saisies du relevé, l'import et l'en-tête de la période. Le premier mot de passe est défini à la première validation ; le **déverrouillage** (menu *Rapprochement*) demande ce mot de passe. Validations et déverrouillages sont tracés dans le *Journal*.
- **Nouveau (rapprochement suivant)** (bouton de la page d'accueil) : si le rapprochement en cours n'est pas validé, propose de le valider ; puis crée le dossier du mois suivant (société, banque, mot de passe conservés ; date proposée = fin du mois suivant) en reportant les écritures non pointées et les lignes du relevé saisies.
- Page *Saisie relevé bancaire* : **Supprimer la sélection** (plusieurs lignes avec Ctrl/Maj + clic) et **Tout effacer** pour vider les lignes reportées.
- Les soldes reportés se placent en sens miroir : solde comptable en **Débit** ⇔ solde du relevé en **Crédit** (et inversement).

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
