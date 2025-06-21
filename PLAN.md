<!-- START OF FILE: plan.md -->
# FILENAME: plan.md
# Version: 2.1.0
# Date: 2025-06-20 22:40
# Author: AI Senior Coder (pour Rolland MELET)
# Description: Mise à jour du plan d'action. L'étape 2.1 (ajout de la bibliothèque) est terminée. Précisions ajoutées pour les étapes suivantes.

# Plan d'Action Détaillé : Intégration de QR-Codes Dynamiques

## 1. Objectif et Stratégie
(... Contenu inchangé ...)

## 2. Prérequis Indispensables
(... Contenu inchangé ...)

## 3. Plan d'Exécution par Phase

---

### **Phase 1 : Publication du Projet `GAS-CreationObjet360sc-API` en Bibliothèque**
(... Contenu inchangé ...)

---

### **Phase 2 : Préparation du Projet `GAS-GenerateurEtiquette`**

*Objectif : Configurer le projet pour qu'il puisse utiliser la bibliothèque et la nouvelle gestion de templates.*

*   **Tâche 2.1 : Inclure la Bibliothèque `Api360sc`**
    *   **STATUT :** ✅ **TERMINÉ**

*   **Tâche 2.2 : Créer et Configurer la feuille `ListeTemplates`**
    *   **Action requise :** Dans le classeur Google Sheets, créer une nouvelle feuille et la nommer **exactement** `ListeTemplates`.
    *   **Action requise :** Configurer les colonnes de cette feuille en suivant précisément ce modèle. Remplacez les `TemplateID` par les vôtres.

| | A (Titre: `NomTemplate`) | B (`TemplateID`) | C (`TypeTemplate`) | D (`QR-Code`) | E (`EnvironnementAPI`) |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **1** | **NomTemplate** | **TemplateID** | **TypeTemplate** | **QR-Code** | **EnvironnementAPI** |
| **2** | Étiquettes Simples (Doc) | `1Abc...xyz_templateDoc` | Google Doc | Non | |
| **3** | Étiquettes QR-Code (TEST) | `2Def...uvw_templateSlide` | Google Slide | Oui | TEST |
| **4** | Étiquettes QR-Code (PROD) | `2Def...uvw_templateSlide` | Google Slide | Oui | PROD |

*   **Tâche 2.3 : Adapter la feuille principale de configuration**
    1.  **Menu déroulant :**
        *   Sélectionner la cellule `B5`.
        *   Aller dans le menu `Données` > `Validation des données` > `Ajouter une règle`.
        *   Dans "Critères", choisir `Liste à partir d'une plage`.
        *   **Action requise :** Pour la plage, entrer la formule `=ListeTemplates!A2:A`. Cliquer sur `OK`.
    2.  **Formules de lookup :**
        *   Renommer les libellés `A8` à `A12` pour correspondre à cette nouvelle structure :
            *   `A8`: `ID du Template (Auto)`
            *   `A9`: `Type de Template (Auto)`
            *   `A10`: `ID Dossier Destination (Optionnel)`
            *   `A11`: `Utilise QR-Code (Auto)`
            *   `A12`: `Environnement API (Auto)`
        *   **Action requise (Cellule B8) :** `=SIERREUR(RECHERCHEV(B5; ListeTemplates!A:E; 2; FAUX); "")`
        *   **Action requise (Cellule B9) :** `=SIERREUR(RECHERCHEV(B5; ListeTemplates!A:E; 3; FAUX); "")`
        *   **Action requise (Cellule B11) :** `=SIERREUR(RECHERCHEV(B5; ListeTemplates!A:E; 4; FAUX); "Non")`
        *   **Action requise (Cellule B12) :** `=SIERREUR(RECHERCHEV(B5; ListeTemplates!A:E; 5; FAUX); "")`

*   **Tâche 2.4 : Créer le Template Google Slides avec Placeholders**
    *   **Action requise :** Créez ou dupliquez une présentation Google Slides.
    *   Dans la diapositive, pour chaque étiquette, insérez les placeholders texte habituels (`{{SERIE}}`, `{{NUMERO1}}`, etc.).
    *   À l'endroit où le QR-Code doit apparaître, **insérez une forme géométrique** (un carré est idéal).
    *   **Action requise :** Pour chaque forme, assignez un nom unique via le panneau "Texte alternatif" (Clic-droit sur la forme -> `Texte alternatif...` -> Remplir le champ `Titre`).
        *   Pour la 1ère étiquette : `QR_CODE_PLACEHOLDER_1`
        *   Pour la 2ème étiquette : `QR_CODE_PLACEHOLDER_2`
        *   ...et ainsi de suite jusqu'à `QR_CODE_PLACEHOLDER_5`.

---

### **Phase 3 : Implémentation de la Nouvelle Logique de Génération**

*Objectif : Modifier `Code.gs` dans `GAS-GenerateurEtiquette` pour intégrer la nouvelle fonctionnalité.*

*   **Tâche 3.1 : Modifier la fonction principale `genererEtiquettes`**
    *   **Action requise :** Mettez à jour l'en-tête de versioning du fichier `Code.gs` à `2.1.0`.
    *   Dans la fonction, après avoir lu tous les paramètres existants, lisez les nouvelles cellules :
        ```javascript
        // ... lecture des autres paramètres ...
        const useQrCode = sheet.getRange('B11').getValue() === 'Oui';
        const environnementApi = sheet.getRange('B12').getValue().toString().trim();
        // ...
        parametres.useQrCode = useQrCode;
        parametres.environnementApi = environnementApi;
        ```
    *   **Action requise :** Adaptez le `dispatcher` pour appeler la nouvelle fonction :
        ```javascript
        if (parametres.useQrCode && parametres.typeTemplate === 'Google Slide') {
          fichiersGeneres = _genererEtiquettesAvecQRCode(parametres);
        } else if (parametres.typeTemplate === 'Google Doc') {
          // ... code existant
        } // ... etc.
        ```

*   **Tâche 3.2 : Créer la fonction `_genererEtiquettesAvecQRCode(parametres)`**
    *   **Rappel bonne pratique :** Cette fonction pourrait devenir longue. Si elle dépasse 100-150 lignes, nous la décomposerons en fonctions utilitaires plus petites (ex: `_createQrCodeObject`, `_replaceShapeWithImage`).
    *   La logique de la fonction sera un mélange de `_genererEtiquettesDepuisSlide` et de nouveaux appels API. Voici les étapes clés :
    ```javascript
    function _genererEtiquettesAvecQRCode(parametres) {
      // 1. Initialisation (date, dossier, nom de fichier...) - similaire à la fonction existante.

      // 2. Création de la présentation temporaire.

      // 3. Boucle sur le nombre de pages.
      for (let page = 0; page < parametres.nbPages; page++) {
        const nouvelleSlide = presentation.insertSlide(/*...*/);

        // 4. Boucle sur les 5 étiquettes de la diapositive.
        for (let i = 1; i <= 5; i++) {
          const numeroEtiquette = formatNumero(baseNum + i - 1);
          const nomObjetBase = `${parametres.serie}-${numeroEtiquette}`;

          // 5. Appel à la bibliothèque pour créer l'objet et obtenir la mcUrl.
          const resultatApi = Api360sc.creerObjetUnique360sc(nomObjetBase, parametres.environnementApi, "MOULE", "Autre");
          const objetApi = JSON.parse(resultatApi);

          if (!objetApi.success || !objetApi.mcUrl) {
            throw new Error(`Échec création objet API pour ${nomObjetBase}: ${objetApi.error}`);
          }

          // 6. Génération de l'URL de l'image du QR-Code.
          const qrCodeUrl = 'https://chart.googleapis.com/chart?cht=qr&chs=150x150&chl=' + encodeURIComponent(objetApi.mcUrl);

          // 7. Remplacement du placeholder par l'image.
          const placeholderName = `QR_CODE_PLACEHOLDER_${i}`;
          // ... code pour trouver la forme, obtenir ses dimensions/positions, insérer l'image, et supprimer la forme.
          
          // 8. Remplacement des placeholders texte.
          nouvelleSlide.replaceAllText(`{{NUMERO${i}}}`, numeroEtiquette);
        }
        // ... Remplacement des placeholders communs à la slide (SERIE, DateJour).
      }

      // 9. Nettoyage (suppression slide template) et conversion en PDF.
    }
    ```

---

### **Phase 4 : Tests et Validation**

*Objectif : Garantir la robustesse et le bon fonctionnement de la nouvelle fonctionnalité.*

*   **Tâche 4.1 : Test de Bout-en-Bout**
    1.  Remplissez la feuille de configuration principale avec des données de test.
    2.  **Action requise :** Dans le menu déroulant en `B5`, sélectionnez le template configuré pour les QR-Codes (ex: "Étiquettes QR-Code (TEST)").
    3.  Vérifiez que `B11` affiche "Oui" et `B12` affiche "TEST".
    4.  Lancez la génération depuis le menu `🏭 DUHALDE ÉTIQUETTES`.
    5.  **Action requise :** Ouvrez le PDF généré.
        *   Vérifiez que les QR-Codes sont présents et nets.
        *   Vérifiez que les autres informations (série, numéros) sont correctes.
        *   Utilisez un smartphone pour scanner un des QR-Codes et assurez-vous que l'URL décodée correspond à une `mcUrl` de la plateforme 360sc.

---

### **Phase 5 : Finalisation et Application des Bonnes Pratiques**

*Objectif : Nettoyer, documenter et versionner correctement le travail accompli.*

*   **Tâche 5.1 : Validation des Bonnes Pratiques**
    *   [ ] **Versioning des Fichiers :** Assurez-vous que tous les fichiers `.gs` modifiés ont un en-tête de versioning mis à jour (`Version: 2.1.0`, date, description).
    *   [ ] **Limite de 500 Lignes :** Vérifiez qu'aucun fichier `.gs` ne dépasse cette limite. Si `Code.gs` s'en approche, planifiez une refactorisation.
    *   [ ] **Lisibilité :** Relisez le code ajouté pour garantir qu'il est clair et bien commenté.

*   **Tâche 5.2 : Mise à jour de la Documentation**
    *   **Action requise :** Mettez à jour le fichier `README.md` du projet `GAS-GenerateurEtiquette`. Expliquez la nouvelle fonctionnalité, la nécessité de la feuille `ListeTemplates` et montrez un exemple de sa configuration.

*   **Tâche 5.3 : Commit Git Final**
    *   **Action requise :** Effectuez un commit propre et descriptif en utilisant les "Conventional Commits".
    ```bash
    # 1. Ajouter les fichiers modifiés au suivi
    git add Code.gs README.md plan.md

    # 2. Créer le commit avec un message explicite
    git commit -m "feat(qrcode): Ajoute la génération d'étiquettes avec QR-Codes dynamiques" -m "Intègre GAS-CreationObjet360sc-API en tant que bibliothèque. La configuration se fait via une feuille 'ListeTemplates' pour une sélection flexible des templates (avec ou sans QR-Code)."

    # 3. Pousser les changements vers le dépôt distant
    git push origin main
    ```

<!-- END OF FILE: plan.md -->