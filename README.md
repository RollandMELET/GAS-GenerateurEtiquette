<!--
# Version: 2.0.0
# Date: 2025-06-20 22:40
# Author: Rolland MELET & AI Senior Coder
# Description: Refonte majeure du README pour intégrer la génération de QR-Codes dynamiques et la nouvelle gestion des templates via une feuille de calcul dédiée.
-->

# 🏭 Générateur d'Étiquettes Duhalde (v2)

Ce projet Google Apps Script permet de générer des planches d'étiquettes prêtes à imprimer en format PDF, à partir de données saisies dans un Google Sheet et d'une liste de templates pré-configurés.

## ✨ Fonctionnalités
*   Génération d'étiquettes en série avec numérotation automatique.
*   **Nouvelle fonctionnalité :** Génération d'étiquettes avec **QR-Codes uniques** liés à la plateforme 360SmartConnect.
*   **Nouvelle configuration :** Gestion centralisée des templates dans une feuille dédiée, permettant de choisir facilement entre différents types de templates (avec ou sans QR-Code, Docs ou Slides).
*   Support de templates **Google Docs** ET **Google Slides**.
*   Export direct au format PDF dans votre Google Drive.
*   Menu intégré dans Google Sheets pour un accès facile.
*   Outils de test de configuration et de création de template d'exemple.

## 🚀 Installation & Configuration Avancée

La configuration se fait maintenant en 3 étapes clés pour une flexibilité maximale.

### Étape 1 : Créer le Script
1.  Ouvrez le Google Sheet qui servira de panneau de contrôle.
2.  Allez dans `Extensions > Apps Script`.
3.  Nommez votre projet (ex: "Générateur Étiquettes Duhalde").
4.  Supprimez le contenu du fichier `Code.gs` et collez-y l'intégralité du script fourni.
5.  Cliquez sur l'icône de sauvegarde 💾.

### Étape 2 : Configurer les Feuilles de Calcul

#### A. Feuille `ListeTemplates` (Le catalogue de vos modèles)

Cette feuille est le cœur de la configuration. Elle vous permet de définir tous vos templates disponibles.

1.  Dans votre classeur, créez une nouvelle feuille et nommez-la **exactement** `ListeTemplates`.
2.  Configurez les colonnes comme dans l'exemple ci-dessous.

| | A | B | C | D | E |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **1** | **NomTemplate** | **TemplateID** | **TypeTemplate** | **QR-Code** | **EnvironnementAPI** |
| **2** | Étiquettes Simples (Doc) | `1Abc...xyz_templateDoc` | Google Doc | Non | |
| **3** | Étiquettes QR-Code (TEST) | `2Def...uvw_templateSlide` | Google Slide | Oui | TEST |
| **4** | Étiquettes QR-Code (PROD) | `2Def...uvw_templateSlide` | Google Slide | Oui | PROD |

*   **`NomTemplate`**: Le nom descriptif que vous choisirez dans le menu déroulant.
*   **`TemplateID`**: L'ID de votre fichier Google Doc/Slide.
*   **`TypeTemplate`**: `Google Doc` ou `Google Slide`.
*   **`QR-Code`**: Mettez `Oui` si ce template doit générer des QR-Codes, sinon `Non`.
*   **`EnvironnementAPI`**: Si `QR-Code` est `Oui`, spécifiez l'environnement (`TEST` ou `PROD`) que l'API doit utiliser.

#### B. Feuille Principale (Le panneau de contrôle)

La feuille principale sert de "panneau de contrôle" pour chaque génération.

1.  **En `B5` (Nom du Template) :** Créez un menu déroulant pointant vers vos templates.
    *   Sélectionnez `B5` > `Données` > `Validation des données` > `Ajouter une règle`.
    *   Critères : `Liste à partir d'une plage`.
    *   Plage : `=ListeTemplates!A2:A`.

2.  **Configuration Générale :** Remplissez les paramètres de génération. Les champs marqués "(Auto)" se rempliront seuls grâce aux formules ci-dessous.

| Cellule  | Description | Exemple de valeur | Formule (si applicable) |
| :------- | :--- | :--- | :--- |
| **A2** | `Série` | `ENVELOPPE` | |
| **B2** | Nom de la série (ENVELOPPE, TOIT, DALLE) | | |
| **A3** | `N° de départ` | `1` | |
| **B3** | Le premier numéro de la série | | |
| **A4** | `Nombre de pages` | `10` | |
| **B4** | Le nombre de planches à générer | | |
| **A5** | `Nom du Template` | `Étiquettes QR-Code (TEST)` | (Menu déroulant) |
| **B5** | **Choisissez le template à utiliser ici.** | | |
| **A8** | `ID du Template (Auto)` | | `=SIERREUR(RECHERCHEV(B5; ListeTemplates!A:E; 2; FAUX); "")` |
| **A9** | `Type de Template (Auto)` | | `=SIERREUR(RECHERCHEV(B5; ListeTemplates!A:E; 3; FAUX); "")` |
| **A10** | `ID Dossier Destination (Optionnel)` | `1x2y3z_FOLDER_ID...` | |
| **A11** | `Utilise QR-Code (Auto)` | | `=SIERREUR(RECHERCHEV(B5; ListeTemplates!A:E; 4; FAUX); "Non")` |
| **A12** | `Environnement API (Auto)` | | `=SIERREUR(RECHERCHEV(B5; ListeTemplates!A:E; 5; FAUX); "")` |

### Étape 3 : Créer vos Templates

> #### Comment obtenir l'ID d'un fichier ?
> Ouvrez votre fichier Google Doc ou Slide. L'URL ressemble à ceci :<br>
> `https://docs.google.com/document/d/ID_DU_FICHIER/edit`<br>
> Copiez la longue chaîne de caractères qui constitue l'`ID_DU_FICHIER`.

#### A. Template Standard (Google Doc ou Slide)
*   Utilisez les placeholders : `{{SERIE}}`, `{{DateJour}}`, `{{NUMERO1}}` à `{{NUMERO5}}`.

#### B. Template avec QR-Code (Google Slide Uniquement)
*   Utilisez les placeholders texte habituels (`{{SERIE}}`, etc.).
*   Pour l'emplacement des QR-Codes, **n'utilisez pas de texte**. À la place, **insérez une forme géométrique** (ex: un carré).
*   Nommez chaque forme de manière unique via le panneau "Texte alternatif" (Clic-droit > `Texte alternatif...` > Champ `Titre`) :
    *   `QR_CODE_PLACEHOLDER_1`
    *   `QR_CODE_PLACEHOLDER_2`
    *   ... jusqu'à `QR_CODE_PLACEHOLDER_5`.

## 📖 Utilisation
Après avoir rechargé votre Google Sheet, un nouveau menu "🏭 DUHALDE ÉTIQUETTES" apparaît. Son utilisation reste inchangée. Le script s'adaptera automatiquement en fonction du template que vous avez sélectionné en `B5`.

*   **Générer Étiquettes :** Lance la génération du PDF en utilisant la configuration actuelle.
*   **Tester Configuration :** Vérifie que le template sélectionné est accessible.
*   **Créer Template Exemple :** Crée un nouveau Google Doc dans votre Drive avec des instructions.

## 🧑‍💻 Pour les Développeurs : Tests Internes
Le script inclut deux fonctions de test qui peuvent être lancées directement depuis l'éditeur Apps Script.
1.  Ouvrez l'éditeur (`Extensions > Apps Script`).
2.  Dans la barre de menu, sélectionnez la fonction à exécuter (ex: `_test_generation_doc`).
3.  **Important :** Vous devrez modifier ces fonctions pour y insérer des ID de templates de test valides qui vous appartiennent.
4.  Cliquez sur `Exécuter`.