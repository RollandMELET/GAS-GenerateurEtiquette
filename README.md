<!-- START OF FILE: README.md -->
<!--
# Version: 2.1.0
# Date: 2025-06-21 18:34
# Author: Rolland MELET & AI Senior Coder
# Description: Version finale post-implémentation des QR-Codes. Inclut un guide détaillé des placeholders.
-->

# 🏭 Générateur d'Étiquettes Duhalde (v2.1)

Ce projet Google Apps Script permet de générer des planches d'étiquettes prêtes à imprimer en format PDF, à partir de données saisies dans un Google Sheet et d'une liste de templates pré-configurés.

## ✨ Fonctionnalités
*   Génération d'étiquettes en série avec numérotation automatique.
*   **Génération de QR-Codes uniques** par étiquette, liés à la plateforme 360SmartConnect.
*   Gestion centralisée et flexible des templates via une feuille de calcul dédiée.
*   Support de templates **Google Docs** (standard) et **Google Slides** (standard & QR-Code).
*   Export direct au format PDF dans votre Google Drive.
*   Menu intégré dans Google Sheets pour un accès facile.

## 🚀 Installation & Configuration

### Étape 1 : Créer une Feuille `ListeTemplates`
Cette feuille est le cœur de la configuration. Elle vous permet de définir tous vos templates disponibles.

1.  Dans votre classeur, créez une nouvelle feuille et nommez-la **exactement** `ListeTemplates`.
2.  Configurez les colonnes comme dans l'exemple ci-dessous.

| | A | B | C | D | E |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **1** | **NomTemplate** | **TemplateID** | **TypeTemplate** | **QR-Code** | **EnvironnementAPI** |
| **2** | Étiquettes Simples | `ID_DU_TEMPLATE_DOC` | Google Doc | Non | |
| **3** | Étiquettes QR-Code TEST | `ID_DU_TEMPLATE_SLIDE_QR` | Google Slide | Oui | TEST |
| **4** | Étiquettes QR-Code PROD | `ID_DU_TEMPLATE_SLIDE_QR` | Google Slide | Oui | PROD |

*   **`NomTemplate`**: Le nom que vous choisirez dans le menu déroulant.
*   **`TemplateID`**: L'ID de votre fichier Google Doc/Slide.
*   **`TypeTemplate`**: `Google Doc` ou `Google Slide`.
*   **`QR-Code`**: `Oui` si ce template doit générer des QR-Codes, sinon `Non`.
*   **`EnvironnementAPI`**: Si `QR-Code` est `Oui`, spécifiez l'environnement (`TEST` ou `PROD`) que l'API doit utiliser.

### Étape 2 : Configurer la Feuille Principale
La feuille principale sert de "panneau de contrôle" pour chaque génération.

1.  **En `B5` (Nom du Template) :** Créez un menu déroulant pointant vers vos templates via `Données` > `Validation des données` > `Liste à partir d'une plage` et utilisez la plage `=ListeTemplates!A2:A`.

2.  **Configuration Générale :** Remplissez les paramètres. Les champs marqués "(Auto)" se rempliront seuls grâce aux formules.

| Cellule  | Description | Formule (à coller dans la cellule) |
| :------- | :--- | :--- |
| **A2** | `Série` | |
| **A3** | `N° de départ` | |
| **A4** | `Nb pages` | |
| **A5** | `Nom du template` | *(Menu déroulant)* |
| **A8** | `ID du Template (Auto)` | `=SIERREUR(RECHERCHEV(B5; ListeTemplates!A:E; 2; FAUX); "")` |
| **A9** | `Type de Template (Auto)` | `=SIERREUR(RECHERCHEV(B5; ListeTemplates!A:E; 3; FAUX); "")` |
| **A10** | `ID Dossier Destination (Optionnel)` | |
| **A11** | `Utilise QR-Code (Auto)` | `=SIERREUR(RECHERCHEV(B5; ListeTemplates!A:E; 4; FAUX); "Non")` |
| **A12** | `Environnement API (Auto)` | `=SIERREUR(RECHERCHEV(B5; ListeTemplates!A:E; 5; FAUX); "")` |


### Étape 3 : Créer vos Templates

#### Guide Détaillé des Placeholders
Le bon fonctionnement du script repose sur l'utilisation de **placeholders** (balises de remplacement) nommés de manière très précise.

##### A. Placeholders de Texte
Ces placeholders fonctionnent dans **Google Docs** et **Google Slides**.

| Placeholder | Rôle |
| :--- | :--- |
| `{{SERIE}}` | Affiche le nom de la série (ex: "ENVELOPPE"). |
| `{{DateJour}}`| Affiche la date de la génération (ex: "21/06/2025"). |
| `{{NUMERO1}}`| Affiche le numéro formaté de la première étiquette. |
| ... | ... |
| `{{NUMERO5}}`| Affiche le numéro de la cinquième étiquette. |

##### B. Placeholders d'Image (pour QR-Codes)
Ces placeholders fonctionnent **uniquement dans les Google Slides** et sont des **formes géométriques nommées**. Le script peut remplacer plusieurs formes portant le même nom.

**Méthode de création (Impératif) :**

1.  **Insérer une Forme :** Insérez une forme (ex: un carré) via `Insertion` > `Forme`. Dimensionnez et positionnez-la.
2.  **Nommer la Forme :**
    *   Faites un **clic-droit** sur la forme > **`Texte alternatif...`**.
    *   Dépliez **`Options avancées`**.
    *   Dans le champ **`Titre`**, entrez le nom exact du placeholder. **Ce champ est le seul que le script lit.**

**Noms des Placeholders d'Image à Utiliser :**

| Placeholder (Nom du `Titre`) | Doit être utilisé pour l'étiquette... |
| :--- | :--- |
| `QR_CODE_PLACEHOLDER_1` | ... qui utilise `{{NUMERO1}}`. |
| `QR_CODE_PLACEHOLDER_2` | ... qui utilise `{{NUMERO2}}`. |
| ... | ... |
| `QR_CODE_PLACEHOLDER_5` | ... qui utilise `{{NUMERO5}}`. |

> **⚠️ AVERTISSEMENT :** Une erreur de frappe, l'utilisation d'un tiret au lieu d'un underscore (`_`), ou le fait de mettre le nom dans le champ `Description` au lieu de `Titre` entraînera l'échec de l'insertion du QR-Code.

<!-- END OF FILE: README.md -->