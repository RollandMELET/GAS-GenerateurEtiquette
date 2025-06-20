<!--
# Version: 1.0.0
# Date: 2025-06-20 11:45
# Author: AI Senior Coder
# Description: README pour le Générateur d'Étiquettes Duhalde.
-->

# 🏭 Générateur d'Étiquettes Duhalde

Ce projet Google Apps Script permet de générer des planches d'étiquettes prêtes à imprimer en format PDF, à partir de données saisies dans un Google Sheet et d'un template Google Workspace.

## ✨ Fonctionnalités
*   Génération d'étiquettes en série avec numérotation automatique.
*   Configuration simple et rapide via un Google Sheet.
*   Support de templates **Google Docs** ET **Google Slides**.
*   Export direct au format PDF dans votre Google Drive.
*   Menu intégré dans Google Sheets pour un accès facile.
*   Outils de test de configuration et de création de template d'exemple.

## 🚀 Installation & Configuration

### Étape 1 : Créer le Script
1.  Ouvrez le Google Sheet qui servira de panneau de contrôle.
2.  Allez dans `Extensions > Apps Script`.
3.  Nommez votre projet (ex: "Générateur Étiquettes Duhalde").
4.  Supprimez le contenu du fichier `Code.gs` et collez-y l'intégralité du script fourni.
5.  Cliquez sur l'icône de sauvegarde 💾.

### Étape 2 : Configurer le Google Sheet
Le script lit sa configuration depuis la feuille active. Préparez votre feuille comme suit :

| Cellule  | Description                                                                                                | Exemple de valeur             |
| :------- | :--------------------------------------------------------------------------------------------------------- | :---------------------------- |
| **A2**   | `Série`                                                                                                    | `ENVELOPPE`                   |
| **B2**   | Nom de la série (ENVELOPPE, TOIT, DALLE)                                                                   |                               |
| **A3**   | `N° de départ`                                                                                             | `1`                           |
| **B3**   | Le premier numéro de la série                                                                              |                               |
| **A4**   | `Nombre de pages`                                                                                          | `10`                          |
| **B4**   | Le nombre de planches à générer                                                                            |                               |
| **A5**   | `Nom du Template`                                                                                          | `Mon Template Étiquettes V1`  |
| **B5**   | Nom descriptif de votre template (informatif)                                                              |                               |
| <!-- A6/B6 et A7/B7 peuvent être laissées vides ou utilisées pour autre chose si besoin -->                            |                                                                                                            |                               |
| **A8**   | `ID du Template`                                                                                           | `1a2b3c4d...`                 |
| **B8**   | L'ID de votre fichier template Google Doc/Slide                                                            |                               |
| **A9**   | `Type de Template`                                                                                         | `Google Doc`                  |
| **B9**   | **Liste déroulante** avec `Google Doc` et `Google Slide`. <br>*(À créer via `Données > Validation des données > Liste d'éléments` : `Google Doc,Google Slide`)* |                               |
| **A10**  | `ID Dossier Destination (Optionnel)`                                                                       | `1x2y3z_FOLDER_ID...`         |
| **B10**  | L'ID du dossier Google Drive où le PDF sera sauvegardé. Si laissé vide, le dossier du template sera utilisé. |                               |

### Étape 3 : Créer votre Template
> #### Comment obtenir l'ID d'un fichier ?
> Ouvrez votre fichier Google Doc ou Slide. L'URL ressemble à ceci :<br>
> `https://docs.google.com/document/d/ID_DU_FICHIER/edit`<br>
> Copiez la longue chaîne de caractères qui constitue l'`ID_DU_FICHIER`.

#### Option A : Template Google Doc
Le template doit être une page A4 contenant les placeholders pour **une** page d'étiquettes. Le script dupliquera cette page.
**Placeholders à utiliser :** `{{SERIE}}`, `{{DateJour}}`, `{{NUMERO1}}`, `{{NUMERO2}}`, `{{NUMERO3}}`, `{{NUMERO4}}`, `{{NUMERO5}}`

#### Option B : Template Google Slide (Recommandé)
C'est l'approche recommandée pour un positionnement précis des éléments. Le template doit être une présentation contenant **une seule diapositive**. Cette diapositive représente votre planche d'étiquettes.
**Placeholders à utiliser :** Identiques à ceux pour Google Docs. Utilisez des zones de texte (TextBox) pour placer chaque placeholder exactement où vous le souhaitez.

## 📖 Utilisation
Après avoir rechargé votre Google Sheet, un nouveau menu "🏭 DUHALDE ÉTIQUETTES" apparaît.
*   **Générer Étiquettes :** Lance la génération du PDF.
*   **Tester Configuration :** Vérifie que l'ID du template est correct et accessible.
*   **Créer Template Exemple :** Crée un nouveau Google Doc dans votre Drive avec des instructions.

## 🧑‍💻 Pour les Développeurs : Tests Internes
Le script inclut deux fonctions de test qui peuvent être lancées directement depuis l'éditeur Apps Script.
1.  Ouvrez l'éditeur (`Extensions > Apps Script`).
2.  Dans la barre de menu, sélectionnez la fonction à exécuter (ex: `_test_generation_doc`).
3.  **Important :** Vous devrez modifier ces fonctions pour y insérer des ID de templates de test valides qui vous appartiennent.
4.  Cliquez sur `Exécuter`.
