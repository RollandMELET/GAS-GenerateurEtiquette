<!-- START OF FILE: plan.md -->
# FILENAME: plan.md
# Version: 2.0.0
# Date: 2025-06-20 22:40
# Author: AI Senior Coder (pour Rolland MELET)
# Description: Plan d'action détaillé V2 pour l'intégration de QR-Codes dynamiques. Inclut les prérequis, des exemples concrets, et le rappel des bonnes pratiques de développement.

# Plan d'Action Détaillé : Intégration de QR-Codes Dynamiques

## 1. Objectif et Stratégie

### 1.1. Objectif Final

Modifier le projet `GAS-GenerateurEtiquette` pour qu'il puisse générer des étiquettes sur un template Google Slides contenant un QR-Code unique. Ce QR-Code correspondra à l'URL (`mcUrl`) d'un objet "Avatar" créé en temps réel sur la plateforme 360SmartConnect, rendant chaque étiquette physiquement et numériquement unique.

### 1.2. Stratégie Technique Adoptée

Nous utiliserons le projet `GAS-CreationObjet360sc-API` en tant que **Bibliothèque Google Apps Script**. Cette architecture est la plus propre et la plus maintenable :
*   **Centralisation :** La logique de communication avec l'API 360sc reste dans un seul et unique projet.
*   **Réutilisabilité :** La logique est disponible pour `GAS-GenerateurEtiquette` sans duplication de code.
*   **Séparation des Rôles :** Chaque projet conserve sa spécialité.

La configuration des templates (support QR-Code, environnement API, etc.) sera gérée dans une feuille dédiée `ListeTemplates` pour une flexibilité maximale.

---

## 2. Prérequis Indispensables

Avant de commencer le développement, il est impératif de s'assurer que les éléments suivants sont en place.

*   **Accès aux Projets Google Apps Script :**
    *   [ ] Accès en **éditeur** au projet `GAS-GenerateurEtiquette`.
    *   [ ] Accès en **éditeur** au projet `GAS-CreationObjet360sc-API`.

*   **Permissions de Déploiement :**
    *   [ ] Vous devez avoir les droits suffisants sur votre compte Google pour créer des **déploiements** de type "Bibliothèque" depuis le projet `GAS-CreationObjet360sc-API`.

*   **Configuration de l'Environnement API :**
    *   [ ] Les **identifiants secrets** pour l'environnement 360sc que vous souhaitez cibler (ex: `TEST`) doivent être correctement configurés dans les `Propriétés du script` de `GAS-CreationObjet360sc-API`.
        *   Clés requises : `API_USERNAME_TEST` et `API_PASSWORD_TEST`.
        *   **Rappel :** La bibliothèque utilisera ses propres propriétés, et non celles du script qui l'appelle.

*   **Connaissance des Bonnes Pratiques :**
    *   [ ] Familiarité avec les règles définies dans le guide `BONNES_PRATIQUES_VERSIONNEMENT.md`, notamment le versionnement sémantique, le format des commits, et les en-têtes de fichiers.

---

## 3. Plan d'Exécution par Phase

### **Phase 1 : Publication du Projet `GAS-CreationObjet360sc-API` en Bibliothèque**

*Objectif : Rendre la logique API disponible pour d'autres projets.*

*   **Tâche 1.1 : Validation de la Configuration API**
    *   Ouvrez le projet `GAS-CreationObjet360sc-API`.
    *   **Action requise :** Exécutez la fonction `checkStoredApiCredentials` et vérifiez dans les journaux que les identifiants pour l'environnement `TEST` sont bien définis. S'ils ne le sont pas, configurez-les via les "Propriétés du Script".
    *   **Raison :** La bibliothèque s'exécutera avec son propre contexte et ses propres secrets. Cette étape garantit qu'elle pourra s'authentifier correctement lorsqu'elle sera appelée.

*   **Tâche 1.2 : Créer un Déploiement de type "Bibliothèque"**
    1.  Dans l'éditeur Apps Script de `GAS-CreationObjet360sc-API`, cliquez sur le bouton bleu `Déployer` > `Nouveau déploiement`.
    2.  Une nouvelle fenêtre s'ouvre. Cliquez sur l'icône engrenage (⚙️) "Sélectionner le type" et choisissez `Bibliothèque`.
    3.  Dans le champ "Description", entrez un nom explicite. **Exemple :** `Bibliothèque API 360sc v1.0 - Création d'objets`.
    4.  Cliquez sur `Déployer`.
    5.  **Action CRUCIALE :** Dans la section "Détails de la bibliothèque", copiez et conservez précieusement l'`ID de script`. Il sera utilisé à la phase suivante.

---

### **Phase 2 : Configuration du Projet `GAS-GenerateurEtiquette`**

*Objectif : Adapter le projet pour qu'il puisse consommer la bibliothèque et gérer les nouveaux templates.*

*   **Tâche 2.1 : Inclure la Bibliothèque `Api360sc`**
    1.  Ouvrez le projet `GAS-GenerateurEtiquette` dans l'éditeur Apps Script.
    2.  Dans le menu de gauche, cliquez sur le `+` à côté de "Bibliothèques".
    3.  Collez l'`ID de script` de la bibliothèque obtenu à l'étape 1.2, puis cliquez sur `Rechercher`.
    4.  Laissez la version sur "HEAD (Développement)" pour le moment afin de bénéficier des dernières modifications sans avoir à redéployer la bibliothèque à chaque fois. Nous la figerons à une version stable plus tard.
    5.  **Action requise :** Changez l'**Identifiant** par défaut pour `Api360sc`. C'est cet alias qui sera utilisé dans le code.
        *   **Exemple d'appel :** `Api360sc.creerObjetUnique360sc(...)`.
    6.  Cliquez sur `Ajouter`.

*   **Tâche 2.2 : Configurer la Feuille `ListeTemplates`**
    *   **Action requise :** Dans votre classeur Google Sheets, créez une nouvelle feuille et nommez-la exactement `ListeTemplates`.
    *   **Action requise :** Reproduisez la structure suivante dans cette feuille. Remplacez les IDs par les vôtres.

| | A (Titre: `NomTemplate`) | B (`TemplateID`) | C (`TypeTemplate`) | D (`QR-Code`) | E (`EnvironnementAPI`) |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **1** | **NomTemplate** | **TemplateID** | **TypeTemplate** | **QR-Code** | **EnvironnementAPI** |
| **2** | Étiquettes Simples (Doc) | `1Abc...xyz_templateDoc` | Google Doc | Non | |
| **3** | Étiquettes QR-Code (TEST) | `2Def...uvw_templateSlide` | Google Slide | Oui | TEST |
| **4** | Étiquettes QR-Code (PROD) | `2Def...uvw_templateSlide` | Google Slide | Oui | PROD |

*   **Tâche 2.3 : Adapter la Feuille de Configuration Principale**
    *   Dans votre feuille principale (là où se trouvent les paramètres en colonne B) :
    1.  **Configurer le menu déroulant :**
        *   Sélectionnez la cellule `B5`.
        *   Allez dans le menu `Données` > `Validation des données`.
        *   Ajoutez une règle. Critères : `Liste à partir d'une plage`.
        *   **Action requise :** Pour la plage, entrez la formule `=ListeTemplates!A2:A`.
    2.  **Configurer les champs auto-remplis :**
        *   Renommez les libellés en `A11` et `A12` pour plus de clarté.
        *   **Action requise (Cellule B11) :** Collez la formule `=SIERREUR(RECHERCHEV(B5; ListeTemplates!A:E; 4; FAUX); "Non")`. Cette formule recherche le nom du template (de B5) dans la feuille `ListeTemplates` et retourne la valeur de la 4ème colonne ("QR-Code").
        *   **Action requise (Cellule B12) :** Collez la formule `=SIERREUR(RECHERCHEV(B5; ListeTemplates!A:E; 5; FAUX); "")`. Cette formule fait de même pour la 5ème colonne ("EnvironnementAPI").

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