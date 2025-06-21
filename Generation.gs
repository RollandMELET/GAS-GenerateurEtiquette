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
  // Cette fonction sera implémentée à la prochaine étape.
  // Pour l'instant, elle lève une erreur pour indiquer qu'elle n'est pas prête.
  console.log("-> Exécution de la logique de génération pour Google Slides avec QR-Codes...");
  throw new Error("La fonction de génération de QR-Codes n'est pas encore implémentée.");
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
      if (element.getType() === DocumentApp.ElementType.TABLE) bodyFinal.appendTable(element);
      else if (element.getType() === DocumentApp.ElementType.PARAGRAPH) bodyFinal.appendParagraph(element);
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