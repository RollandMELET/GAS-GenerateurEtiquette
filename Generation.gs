// <!-- START OF FILE: Generation.gs -->
// FILENAME: Generation.gs
// Version: 1.0.0
// Date: 2025-06-21 18:34
// Author: Rolland MELET & AI Senior Coder
// Description: Contient toute la logique de génération de documents (Docs, Slides, et Slides avec QR-Codes).

// ========================================
// LOGIQUE DE GÉNÉRATION SPÉCIFIQUE
// ========================================

/**
 * [NOUVEAU] Génère un PDF à partir d'un template Slides en y insérant des QR-Codes dynamiques.
 * @param {object} parametres - L'objet contenant la configuration complète.
 * @returns {Array<object>} Un tableau d'objets représentant les fichiers générés.
 * @private
 */
function _genererEtiquettesAvecQRCode(parametres) {
  console.log("-> Exécution de la logique de génération pour Google Slides avec QR-Codes...");

  // 1. Initialisation
  const maintenant = new Date();
  const dateFormatee = formatDateFrancais(maintenant);
  const dossier = creerDossierSiNecessaire(parametres);
  const premierNum = formatNumero(parametres.numeroDebut);
  const dernierNum = formatNumero(parametres.numeroDebut + (parametres.nbPages * 5) - 1);
  const nomFichier = `Etiquettes_QR_${parametres.serie}_${premierNum}-${dernierNum}_${dateFormatee.replace(/\//g, '-')}`;

  // 2. Création de la présentation temporaire à partir du template
  const templateSlideFile = DriveApp.getFileById(parametres.templateId);
  const presentationCopieFile = templateSlideFile.makeCopy(nomFichier + '_temp', dossier);
  const presentation = SlidesApp.openById(presentationCopieFile.getId());

  const slideTemplate = presentation.getSlides()[0];
  if (!slideTemplate) throw new Error("Le template Google Slides ne contient aucune diapositive !");

  // 3. Boucle sur le nombre de pages à générer
  for (let page = 0; page < parametres.nbPages; page++) {
    console.log(`--- Création page (diapositive) ${page + 1}/${parametres.nbPages} ---`);
    const baseNum = parametres.numeroDebut + (page * 5);
    
    // Dupliquer la diapositive du template pour chaque nouvelle page
    const nouvelleSlide = presentation.insertSlide(presentation.getSlides().length, slideTemplate);

    // Remplacer les placeholders communs à la diapositive
    nouvelleSlide.replaceAllText('{{SERIE}}', parametres.serie);
    nouvelleSlide.replaceAllText('{{DateJour}}', dateFormatee);

    // 4. Boucle sur les 5 étiquettes de la diapositive
    for (let i = 1; i <= 5; i++) {
      const numeroEtiquette = formatNumero(baseNum + i - 1);
      console.log(`-- Traitement étiquette ${i}/5, Numéro: ${numeroEtiquette}`);
      
      // A. Remplacer le placeholder de numéro
      nouvelleSlide.replaceAllText(`{{NUMERO${i}}}`, numeroEtiquette);
      
      // B. Créer l'objet via l'API et générer le QR-Code
      const nomObjetBase = `${parametres.serie}-${numeroEtiquette}`;
      const placeholderName = `QR_CODE_PLACEHOLDER_${i}`;
      
      try {
        // APPEL À LA BIBLIOTHÈQUE
        console.log(`Appel API pour créer l'objet: ${nomObjetBase} sur env: ${parametres.environnementApi}`);
        const resultatApiString = Api360sc.creerObjetUnique360sc(nomObjetBase, parametres.environnementApi, "MOULE", "Autre");
        const objetApi = JSON.parse(resultatApiString);

        if (!objetApi.success || !objetApi.mcUrl) {
          throw new Error(`Échec création objet API pour ${nomObjetBase}: ${objetApi.error || 'mcUrl manquante'}`);
        }
        console.log(`Objet créé avec succès. mcUrl: ${objetApi.mcUrl}`);

        // Générer l'URL de l'image du QR-Code via l'API Google Charts
        const qrCodeUrl = 'https://chart.googleapis.com/chart?cht=qr&chs=150x150&chl=' + encodeURIComponent(objetApi.mcUrl);

        // Trouver la forme placeholder par son titre
        const shapePlaceholder = nouvelleSlide.getPageElements().find(el => 
          el.getPageElementType() === SlidesApp.PageElementType.SHAPE && el.asShape().getTitle() === placeholderName
        );

        if (shapePlaceholder) {
          // Insérer l'image du QR-Code aux mêmes dimensions et position, puis supprimer le placeholder
          const left = shapePlaceholder.getLeft();
          const top = shapePlaceholder.getTop();
          const width = shapePlaceholder.getWidth();
          const height = shapePlaceholder.getHeight();
          
          nouvelleSlide.insertImage(qrCodeUrl, left, top, width, height);
          shapePlaceholder.remove();
          console.log(`Placeholder ${placeholderName} remplacé par un QR-Code.`);
        } else {
          console.warn(`Avertissement: Placeholder de forme '${placeholderName}' non trouvé sur la diapositive.`);
        }
      } catch(e) {
        console.error(`Erreur critique lors du traitement de l'étiquette ${numeroEtiquette}: ${e.toString()}`);
        // En cas d'erreur API, on affiche un message d'erreur à la place du QR-Code pour ne pas bloquer toute la génération
        const shapePlaceholder = nouvelleSlide.getPageElements().find(el => el.asShape() && el.asShape().getTitle() === placeholderName);
        if(shapePlaceholder) shapePlaceholder.asShape().getText().setText(`ERREUR API\n${e.message.substring(0, 50)}...`);
      }
    }
  }

  // 5. Nettoyage et finalisation
  presentation.getSlides()[0].remove();
  presentation.saveAndClose();
  
  const pdfBlob = presentationCopieFile.getAs('application/pdf').setName(nomFichier + '.pdf');
  const pdfFile = dossier.createFile(pdfBlob);
  console.log(`PDF multi-pages avec QR-Codes créé: ${pdfFile.getUrl()}`);

  presentationCopieFile.setTrashed(true);

  return [{ url: pdfFile.getUrl(), nom: pdfFile.getName() }];
}


/**
 * Génère un PDF multi-pages à partir d'un template Google Docs.
 * @param {object} parametres - L'objet contenant la configuration.
 * @returns {Array<object>} Un tableau d'objets représentant les fichiers générés.
 * @private
 */
function _genererEtiquettesDepuisDoc(parametres) {
  console.log("-> Exécution de la logique de génération pour Google Docs...");
  const maintenant = new Date();
  const dateFormatee = formatDateFrancais(maintenant);
  const dossier = creerDossierSiNecessaire(parametres);
  const premierNum = formatNumero(parametres.numeroDebut);
  const dernierNum = formatNumero(parametres.numeroDebut + (parametres.nbPages * 5) - 1);
  const nomFichier = `Etiquettes_${parametres.serie}_${premierNum}-${dernierNum}_${dateFormatee.replace(/\//g, '-')}`;
  
  const documentsPages = [];
  
  for (let page = 0; page < parametres.nbPages; page++) {
    const baseNum = parametres.numeroDebut + (page * 5);
    const numeros = [formatNumero(baseNum), formatNumero(baseNum + 1), formatNumero(baseNum + 2), formatNumero(baseNum + 3), formatNumero(baseNum + 4)];
    const templateDoc = DriveApp.getFileById(parametres.templateId);
    const docPageCopie = templateDoc.makeCopy(`${nomFichier}_page${page + 1}_temp`, dossier);
    const docPage = DocumentApp.openById(docPageCopie.getId());
    const bodyPage = docPage.getBody();
    
    bodyPage.replaceText('{{SERIE}}', parametres.serie);
    numeros.forEach((num, index) => bodyPage.replaceText(`{{NUMERO${index + 1}}}`, num));
    bodyPage.replaceText('{{DateJour}}', dateFormatee);
    
    docPage.saveAndClose();
    documentsPages.push(docPageCopie);
  }
  
  console.log("--- Fusion des pages ---");
  const docFinal = DocumentApp.openById(documentsPages[0].getId());
  const bodyFinal = docFinal.getBody();
  
  for (let i = 1; i < documentsPages.length; i++) {
    bodyFinal.appendPageBreak();
    const bodySource = DocumentApp.openById(documentsPages[i].getId()).getBody();
    for (let j = 0; j < bodySource.getNumChildren(); j++) {
      const element = bodySource.getChild(j).copy();
      const type = element.getType();
      if (type === DocumentApp.ElementType.TABLE) bodyFinal.appendTable(element);
      else if (type === DocumentApp.ElementType.PARAGRAPH) bodyFinal.appendParagraph(element);
    }
  }
  
  docFinal.setName(nomFichier + '_temp');
  docFinal.saveAndClose();
  
  const finalFileId = docFinal.getId();
  let pdfFile;
  const MAX_PDF_RETRIES = 7, PDF_RETRY_INITIAL_DELAY_MS = 10000, PDF_RETRY_SUBSEQUENT_DELAY_MS = 10000;

  Utilities.sleep(PDF_RETRY_INITIAL_DELAY_MS);

  for (let attempt = 1; attempt <= MAX_PDF_RETRIES; attempt++) {
    try {
      const fileToConvert = DriveApp.getFileById(finalFileId);
      const pdfBlob = fileToConvert.getAs('application/pdf').setName(nomFichier + '.pdf');
      pdfFile = dossier.createFile(pdfBlob);
      console.log(`PDF multi-pages créé: ${pdfFile.getUrl()}`);
      break;
    } catch (e) {
      console.warn(`Erreur tentative ${attempt} PDF: ${e.toString()}`);
      if (e.message.includes("Service unavailable") && attempt < MAX_PDF_RETRIES) {
        Utilities.sleep(PDF_RETRY_SUBSEQUENT_DELAY_MS);
      } else {
        throw e;
      }
    }
  }
  if (!pdfFile) throw new Error("Impossible de générer le PDF après " + MAX_PDF_RETRIES + " tentatives.");
  
  documentsPages.forEach(doc => DriveApp.getFileById(doc.getId()).setTrashed(true));
  
  return [{ url: pdfFile.getUrl(), nom: pdfFile.getName() }];
}

/**
 * Génère un PDF multi-pages à partir d'un template Google Slides (sans QR-Code).
 * @param {object} parametres - L'objet contenant la configuration.
 * @returns {Array<object>} Un tableau d'objets représentant les fichiers générés.
 * @private
 */
function _genererEtiquettesDepuisSlide(parametres) {
  console.log("-> Exécution de la logique de génération pour Google Slides...");
  
  const maintenant = new Date();
  const dateFormatee = formatDateFrancais(maintenant);
  const dossier = creerDossierSiNecessaire(parametres);
  const premierNum = formatNumero(parametres.numeroDebut);
  const dernierNum = formatNumero(parametres.numeroDebut + (parametres.nbPages * 5) - 1);
  const nomFichier = `Etiquettes_${parametres.serie}_${premierNum}-${dernierNum}_${dateFormatee.replace(/\//g, '-')}`;
  
  const templateSlideFile = DriveApp.getFileById(parametres.templateId);
  const presentationCopieFile = templateSlideFile.makeCopy(nomFichier + '_temp', dossier);
  const presentation = SlidesApp.openById(presentationCopieFile.getId());
  
  const slideTemplate = presentation.getSlides()[0];
  if (!slideTemplate) throw new Error("Le template Google Slides ne contient aucune diapositive !");

  for (let page = 0; page < parametres.nbPages; page++) {
    const baseNum = parametres.numeroDebut + (page * 5);
    const numeros = [formatNumero(baseNum), formatNumero(baseNum + 1), formatNumero(baseNum + 2), formatNumero(baseNum + 3), formatNumero(baseNum + 4)];
    const nouvelleSlide = presentation.insertSlide(presentation.getSlides().length, slideTemplate);
    
    nouvelleSlide.replaceAllText('{{SERIE}}', parametres.serie);
    numeros.forEach((num, index) => nouvelleSlide.replaceAllText(`{{NUMERO${index + 1}}}`, num));
    nouvelleSlide.replaceAllText('{{DateJour}}', dateFormatee);
  }
  
  presentation.getSlides()[0].remove();
  
  presentation.saveAndClose();
  const pdfBlob = presentationCopieFile.getAs('application/pdf').setName(nomFichier + '.pdf');
  const pdfFile = dossier.createFile(pdfBlob);
  console.log(`PDF multi-pages créé: ${pdfFile.getUrl()}`);

  presentationCopieFile.setTrashed(true);
  
  return [{ url: pdfFile.getUrl(), nom: pdfFile.getName() }];
}

// <!-- END OF FILE: Generation.gs -->