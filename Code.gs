<!-- START OF FILE: Code.gs -->
// Version: 1.2.5
// Date: 2025-06-20 21:15 
// Author: Rolland MELET & AI Senior Coder
// Description: Ajout de la possibilité de spécifier un dossier de destination pour les PDF générés (cellule B10).

// ========================================
// MENU PERSONNALISÉ GOOGLE SHEETS
// ========================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🏭 DUHALDE ÉTIQUETTES')
    .addItem('📋 Générer Étiquettes', 'genererEtiquettes')
    .addSeparator()
    .addItem('⚙️ Tester Configuration', 'testerConfiguration')
    .addItem('📝 Créer Template Exemple', 'creerTemplateExemple')
    .addSeparator()
    .addItem('🔧 À propos', 'afficherAPropos')
    .addToUi();
}

function afficherAPropos() {
  const message = `🏭 GÉNÉRATEUR D'ÉTIQUETTES DUHALDE

Version: 1.0
Développé pour: 360SmartConnect

Fonctionnalités:
✅ Génération automatique d'étiquettes AGIPA 119013
✅ Support 4 séries: ENVELOPPE, TOIT, DALLE, TEST
✅ Export PDF prêt à imprimer
✅ Sauvegarde dans le même dossier que le template

Support: Vérifiez que l'ID du template est correct dans le script.`;

  SpreadsheetApp.getUi().alert(message);
}

// ========================================
// FONCTION PRINCIPALE (DISPATCHER)
// ========================================

function genererEtiquettes() {
  console.log("=== DÉBUT GÉNÉRATION ÉTIQUETTES ===");
  
  try {
    // 1. Lire les paramètres depuis le Google Sheet
    console.log("1. Lecture des paramètres...");
    const sheet = SpreadsheetApp.getActiveSheet();
    const serie = sheet.getRange('B2').getValue() || 'ENVELOPPE';
    const numeroDebut = parseInt(sheet.getRange('B3').getValue()) || 1;
    const nbPages = parseInt(sheet.getRange('B4').getValue()) || 1;
    const nomTemplate = sheet.getRange('B5').getValue().toString().trim(); // Nouveau: Nom du template
    const templateId = sheet.getRange('B8').getValue().toString().trim();
    const typeTemplate = sheet.getRange('B9').getValue().toString().trim(); // Nouveau: Type de template
    const dossierId = sheet.getRange('B10').getValue().toString().trim(); // Nouveau: ID du dossier de destination
    
    // 2. Validation
    if (!['ENVELOPPE', 'TOIT', 'DALLE', 'TEST'].includes(serie)) {
      throw new Error('Série doit être : ENVELOPPE, TOIT, DALLE ou TEST');
    }
    if (numeroDebut < 1 || nbPages < 1) {
      throw new Error('N° début et Nb pages doivent être >= 1');
    }
    if (!templateId) {
      throw new Error('ID du template manquant dans la cellule B8');
    }
    if (!typeTemplate) {
      throw new Error('Type de template manquant dans la cellule B9. Doit être "Google Doc" ou "Google Slide".');
    }
    if (typeTemplate !== 'Google Doc' && typeTemplate !== 'Google Slide') {
      throw new Error('Type de template invalide en B9. Doit être "Google Doc" ou "Google Slide".');
    }

    // Regrouper les paramètres dans un objet pour les passer aux fonctions de génération
    const parametres = {
      serie: serie,
      numeroDebut: numeroDebut,
      nbPages: nbPages,
      templateId: templateId,
      nomTemplate: nomTemplate, // Ajout du nom du template aux paramètres
      typeTemplate: typeTemplate, // Ajout du type de template aux paramètres
      dossierId: dossierId // Ajout de l'ID du dossier de destination
    };
    
    console.log(`Paramètres: série=${parametres.serie}, numeroDebut=${parametres.numeroDebut}, nbPages=${parametres.nbPages}, nomTemplate=${parametres.nomTemplate}, typeTemplate=${parametres.typeTemplate}, templateId=${parametres.templateId}, dossierId=${parametres.dossierId}`);

    let fichiersGeneres;
    if (parametres.typeTemplate === 'Google Doc') {
      fichiersGeneres = _genererEtiquettesDepuisDoc(parametres);
    } else if (parametres.typeTemplate === 'Google Slide') {
      fichiersGeneres = _genererEtiquettesDepuisSlide(parametres);
    } else {
      // Cette erreur est déjà gérée par la validation plus haut, mais par sécurité :
      throw new Error('Type de template non supporté : ' + parametres.typeTemplate);
    }

    // 6. Message de succès
    const message = `✅ PDF d'étiquettes généré !\n\nSérie: ${parametres.serie}\nNuméros: ${formatNumero(parametres.numeroDebut)} à ${formatNumero(parametres.numeroDebut + (parametres.nbPages * 5) - 1)}\nPages: ${parametres.nbPages}\n\n📁 Fichier sauvé dans le même dossier que votre template\n\n🖨️ Prêt à imprimer directement !`;
    
    console.log("Affichage du message de succès");
    SpreadsheetApp.getUi().alert(message);
    
    console.log("=== FIN GÉNÉRATION RÉUSSIE ===");
    return fichiersGeneres;
    
  } catch (error) {
    console.error("ERREUR:", error.toString());
    console.error("Stack:", error.stack);
    
    SpreadsheetApp.getUi().alert(`❌ Erreur lors de la génération:\n\n${error.toString()}`);
    throw error;
  }
}

// ========================================
// LOGIQUE DE GÉNÉRATION SPÉCIFIQUE
// ========================================

/**
 * Génère un PDF multi-pages à partir d'un template Google Docs.
 * @param {object} parametres - L'objet contenant la configuration (serie, numeroDebut, nbPages, templateId).
 * @returns {Array<object>} Un tableau d'objets représentant les fichiers générés.
 * @private
 */
function _genererEtiquettesDepuisDoc(parametres) {
  console.log("-> Exécution de la logique de génération pour Google Docs...");
  
  const maintenant = new Date();
  const dateFormatee = formatDateFrancais(maintenant);
  console.log(`Date: ${dateFormatee}`);
    
  const dossier = creerDossierSiNecessaire(parametres);
  
  console.log("--- Création du document multi-pages ---");
  
  const premierNum = formatNumero(parametres.numeroDebut);
  const dernierNum = formatNumero(parametres.numeroDebut + (parametres.nbPages * 5) - 1);
  const nomFichier = `Etiquettes_${parametres.serie}_${premierNum}-${dernierNum}_${dateFormatee.replace(/\//g, '-')}`;
  
  const documentsPages = [];
  
  for (let page = 0; page < parametres.nbPages; page++) {
    console.log(`--- Création page ${page + 1}/${parametres.nbPages} ---`);
    const baseNum = parametres.numeroDebut + (page * 5);
    const numeros = [
      formatNumero(baseNum), formatNumero(baseNum + 1), formatNumero(baseNum + 2),
      formatNumero(baseNum + 3), formatNumero(baseNum + 4)
    ];
    
    const templateDoc = DriveApp.getFileById(parametres.templateId);
    const docPageCopie = templateDoc.makeCopy(`${nomFichier}_page${page + 1}_temp`, dossier);
    const docPage = DocumentApp.openById(docPageCopie.getId());
    const bodyPage = docPage.getBody();
    
    bodyPage.replaceText('{{SERIE}}', parametres.serie);
    bodyPage.replaceText('{{NUMERO1}}', numeros[0]);
    bodyPage.replaceText('{{NUMERO2}}', numeros[1]);
    bodyPage.replaceText('{{NUMERO3}}', numeros[2]);
    bodyPage.replaceText('{{NUMERO4}}', numeros[3]);
    bodyPage.replaceText('{{NUMERO5}}', numeros[4]);
    bodyPage.replaceText('{{DateJour}}', dateFormatee);
    
    docPage.saveAndClose();
    documentsPages.push(docPageCopie);
  }
  
  console.log("--- Fusion des pages ---");
  const docFinal = DocumentApp.openById(documentsPages[0].getId());
  const bodyFinal = docFinal.getBody();
  
  for (let i = 1; i < documentsPages.length; i++) {
    bodyFinal.appendPageBreak();
    const docSource = DocumentApp.openById(documentsPages[i].getId());
    const bodySource = docSource.getBody();
    const numElements = bodySource.getNumChildren();
    for (let j = 0; j < numElements; j++) {
      const element = bodySource.getChild(j).copy();
      const type = element.getType();
      if (type === DocumentApp.ElementType.TABLE) bodyFinal.appendTable(element);
      else if (type === DocumentApp.ElementType.PARAGRAPH) bodyFinal.appendParagraph(element);
    }
  }
  
  docFinal.setName(nomFichier + '_temp');
  docFinal.saveAndClose();
  
  const finalFileId = docFinal.getId(); // Obtenir l'ID avant la boucle de tentatives
  let pdfBlob;
  let pdfFile;
  const MAX_PDF_RETRIES = 7; // Augmentation du nombre de tentatives
  const PDF_RETRY_INITIAL_DELAY_MS = 10000; // Délai initial avant la première tentative (10 secondes)
  const PDF_RETRY_SUBSEQUENT_DELAY_MS = 10000; // Délai entre les tentatives suivantes (10 secondes)

  console.log(`Attente de ${PDF_RETRY_INITIAL_DELAY_MS / 1000} secondes avant la première tentative de conversion PDF du fichier ID: ${finalFileId}...`);
  Utilities.sleep(PDF_RETRY_INITIAL_DELAY_MS); // Pause initiale plus longue après saveAndClose

  for (let attempt = 1; attempt <= MAX_PDF_RETRIES; attempt++) {
    try {
      // Re-récupérer l'objet fichier à chaque tentative pour avoir la référence la plus fraîche
      const fileToConvert = DriveApp.getFileById(finalFileId);
      console.log(`Conversion en PDF (tentative ${attempt}/${MAX_PDF_RETRIES})...`);
      pdfBlob = fileToConvert.getAs('application/pdf');
      pdfBlob.setName(nomFichier + '.pdf');
      pdfFile = dossier.createFile(pdfBlob);
      console.log(`PDF multi-pages créé: ${pdfFile.getUrl()}`);
      break; // Si succès, sortir de la boucle
    } catch (e) {
      console.warn(`Erreur lors de la tentative ${attempt} de conversion PDF: ${e.toString()}`);
      // Vérifier spécifiquement le message d'erreur et si d'autres tentatives sont possibles
      if (e.message && e.message.includes("Service unavailable") && attempt < MAX_PDF_RETRIES) { // Vérifier "Service unavailable" plus généralement
        console.log(`Nouvel essai de conversion PDF dans ${PDF_RETRY_SUBSEQUENT_DELAY_MS / 1000} secondes...`);
        Utilities.sleep(PDF_RETRY_SUBSEQUENT_DELAY_MS);
      } else {
        console.error("Échec final de la conversion PDF après plusieurs tentatives ou erreur non récupérable.");
        // Loguer plus d'informations si possible
        console.error(`Détails du fichier avant échec: ID=${finalFileId}, Nom (temporaire)=${docFinal.getName()}`); // docFinal pourrait ne plus être accessible ici, mais l'ID est sûr.
        throw e; // Relancer l'erreur si ce n'est pas une erreur de service récupérable ou si c'est la dernière tentative
      }
    }
  }
  if (!pdfFile) { // Vérification supplémentaire au cas où la boucle se terminerait sans succès ni erreur relancée
    throw new Error("Impossible de générer le fichier PDF après " + MAX_PDF_RETRIES + " tentatives.");
  }
  console.log("Suppression des documents temporaires...");
  documentsPages.forEach(doc => {
    DriveApp.getFileById(doc.getId()).setTrashed(true);
  });
  
  return [{ url: pdfFile.getUrl(), nom: pdfFile.getName() }];
}

/**
 * Génère un PDF multi-pages à partir d'un template Google Slides.
 * @param {object} parametres - L'objet contenant la configuration (serie, numeroDebut, nbPages, templateId).
 * @returns {Array<object>} Un tableau d'objets représentant les fichiers générés.
 * @private
 */
function _genererEtiquettesDepuisSlide(parametres) {
  console.log("-> Exécution de la logique de génération pour Google Slides...");
  
  const maintenant = new Date();
  const dateFormatee = formatDateFrancais(maintenant);
  console.log(`Date: ${dateFormatee}`);
  
  const dossier = creerDossierSiNecessaire(parametres);
  
  const premierNum = formatNumero(parametres.numeroDebut);
  const dernierNum = formatNumero(parametres.numeroDebut + (parametres.nbPages * 5) - 1);
  const nomFichier = `Etiquettes_${parametres.serie}_${premierNum}-${dernierNum}_${dateFormatee.replace(/\//g, '-')}`;
  
  // 1. Copier le template de présentation
  const templateSlideFile = DriveApp.getFileById(parametres.templateId);
  const presentationCopieFile = templateSlideFile.makeCopy(nomFichier + '_temp', dossier);
  const presentation = SlidesApp.openById(presentationCopieFile.getId());
  
  // 2. Le template est supposé n'avoir qu'une seule diapositive
  const slideTemplate = presentation.getSlides()[0];
  if (!slideTemplate) {
    throw new Error("Le template Google Slides ne contient aucune diapositive !");
  }

  // 3. Générer les pages en dupliquant la diapositive du template
  for (let page = 0; page < parametres.nbPages; page++) {
    console.log(`--- Création page (diapositive) ${page + 1}/${parametres.nbPages} ---`);
    const baseNum = parametres.numeroDebut + (page * 5);
    const numeros = [
      formatNumero(baseNum), formatNumero(baseNum + 1), formatNumero(baseNum + 2),
      formatNumero(baseNum + 3), formatNumero(baseNum + 4)
    ];
    
    // Dupliquer la diapositive du template pour chaque nouvelle page
    const nouvelleSlide = presentation.insertSlide(presentation.getSlides().length, slideTemplate);
    
    // Remplacer les placeholders sur la NOUVELLE diapositive
    nouvelleSlide.replaceAllText('{{SERIE}}', parametres.serie);
    nouvelleSlide.replaceAllText('{{NUMERO1}}', numeros[0]);
    nouvelleSlide.replaceAllText('{{NUMERO2}}', numeros[1]);
    nouvelleSlide.replaceAllText('{{NUMERO3}}', numeros[2]);
    nouvelleSlide.replaceAllText('{{NUMERO4}}', numeros[3]);
    nouvelleSlide.replaceAllText('{{NUMERO5}}', numeros[4]);
    nouvelleSlide.replaceAllText('{{DateJour}}', dateFormatee);
  }
  
  // 4. Supprimer la diapositive originale du template (la première)
  presentation.getSlides()[0].remove();
  
  // 5. Sauvegarder la présentation et la convertir en PDF
  presentation.saveAndClose();
  console.log("Conversion en PDF...");
  const pdfBlob = presentationCopieFile.getAs('application/pdf');
  pdfBlob.setName(nomFichier + '.pdf');
  const pdfFile = dossier.createFile(pdfBlob);
  console.log(`PDF multi-pages créé: ${pdfFile.getUrl()}`);

  // 6. Supprimer la présentation temporaire
  console.log("Suppression de la présentation temporaire...");
  presentationCopieFile.setTrashed(true);
  
  return [{ url: pdfFile.getUrl(), nom: pdfFile.getName() }];
}


// ========================================
// FONCTIONS UTILITAIRES ET DE TEST
// ========================================

function formatNumero(num) {
  return num.toString().padStart(5, '0');
}

function formatDateFrancais(date) {
  const jour = date.getDate().toString().padStart(2, '0');
  const mois = (date.getMonth() + 1).toString().padStart(2, '0');
  const annee = date.getFullYear();
  return `${jour}/${mois}/${annee}`;
}

function creerDossierSiNecessaire(parametres) {
  if (parametres.dossierId) {
    try {
      const dossier = DriveApp.getFolderById(parametres.dossierId);
      console.log(`Dossier de destination spécifié: ${dossier.getName()}`);
      return dossier;
    } catch (e) {
      console.error(`ID du dossier de destination invalide ou accès refusé: ${parametres.dossierId}`);
      throw new Error(`L'ID du dossier de destination en B10 est invalide ou vous n'y avez pas accès.`);
    }
  } else {
    // Comportement par défaut : utiliser le dossier du template
    console.log("Aucun dossier de destination spécifié, déduction à partir du template...");
    const templateFile = DriveApp.getFileById(parametres.templateId);
    const dossiersParents = templateFile.getParents();
    if (dossiersParents.hasNext()) {
      const dossierParent = dossiersParents.next();
      console.log(`Dossier de destination déduit du template: ${dossierParent.getName()}`);
      return dossierParent;
    } else {
      console.log("Template dans la racine du Drive. Utilisation de la racine.");
      return DriveApp.getRootFolder();
    }
  }
}

function testerConfiguration() {
  console.log("=== TEST DE CONFIGURATION ===");
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const templateId = sheet.getRange('B8').getValue().toString().trim();
    const typeTemplate = sheet.getRange('B9').getValue().toString().trim();
    const dossierId = sheet.getRange('B10').getValue().toString().trim(); // Nouveau

    if (!templateId) {
      throw new Error('ID du template manquant dans la cellule B8');
    }
    if (!typeTemplate) {
      throw new Error('Type de template manquant dans la cellule B9. Doit être "Google Doc" ou "Google Slide".');
    }

    let templateFile;
    let alertMessage = '✅ Configuration OK.';

    // Test du template
    if (typeTemplate === 'Google Doc') {
      templateFile = DriveApp.getFileById(templateId);
      DocumentApp.openById(templateId); // Tente d'ouvrir pour vérifier l'accès et le type
      console.log(`Template Google Doc trouvé: ${templateFile.getName()}`);
    } else if (typeTemplate === 'Google Slide') {
      templateFile = DriveApp.getFileById(templateId);
      SlidesApp.openById(templateId); // Tente d'ouvrir pour vérifier l'accès et le type
      console.log(`Template Google Slide trouvé: ${templateFile.getName()}`);
    } else {
      throw new Error('Type de template invalide en B9. Doit être "Google Doc" ou "Google Slide".');
    }
    alertMessage += `\nType: ${typeTemplate}\nTemplate: ${templateFile.getName()}\nID: ${templateId}`;

    // Test du dossier de destination (si fourni)
    if (dossierId) {
      const dossierDestination = DriveApp.getFolderById(dossierId);
      console.log(`Dossier de destination trouvé: ${dossierDestination.getName()}`);
      alertMessage += `\nDossier Destination: ${dossierDestination.getName()}`;
    } else {
      alertMessage += `\n\nDossier Destination: (dossier du template sera utilisé)`;
    }

    alertMessage += `\n\nLe script peut accéder à tous les éléments.`;
    SpreadsheetApp.getUi().alert(alertMessage);
  } catch (error) {
    console.error("❌ Erreur de configuration:", error.toString());
    SpreadsheetApp.getUi().alert(`❌ Problème détecté:\n\n${error.toString()}`);
  }
}

function creerTemplateExemple() {
  console.log("Création d'un template exemple...");
  const doc = DocumentApp.create('Template Étiquettes DUHALDE - EXEMPLE');
  const body = doc.getBody();
  body.setPageWidth(595.28).setPageHeight(841.89); // A4
  body.setMarginTop(0).setMarginBottom(0).setMarginLeft(0).setMarginRight(0);
  body.appendParagraph('TEMPLATE ÉTIQUETTES DUHALDE').setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph('Placeholders à utiliser: {{SERIE}}, {{NUMERO1}} à {{NUMERO5}}, {{DateJour}}');
  
  const url = doc.getUrl();
  console.log(`Template exemple créé: ${url}`);
  SpreadsheetApp.getUi().alert(`Template exemple créé:\n\n${url}`);
}

// ========================================
// FONCTIONS DE TEST INTERNES (POUR DÉVELOPPEMENT)
// ========================================

/**
 * Fonction de test pour la génération d'étiquettes depuis un template Google Docs.
 * REMPLACEZ 'ID_VOTRE_TEMPLATE_DOC_ICI' par un ID de template Google Doc valide.
 * Cette fonction peut être exécutée directement depuis l'éditeur Apps Script.
 */
function _test_generation_doc() {
  console.log("=== DÉBUT TEST GÉNÉRATION DOCS ===");
  try {
    const parametresTest = {
      serie: "TEST_DOC",
      numeroDebut: 101,
      nbPages: 2,
      templateId: "1U5QmMzr2Q0Sf4KVOnLjA1FdK2S4Y7WnV-O77H6xx6vM", // ⚠️ REMPLACEZ PAR UN VRAI ID DE TEMPLATE DOC
      nomTemplate: "Template Test Doc",
      typeTemplate: "Google Doc", // Important pour la cohérence
      dossierId: null // Mettre un ID de dossier ici pour tester, ou laisser null/vide pour utiliser le dossier du template
    };

    if (parametresTest.templateId === "ID_VOTRE_TEMPLATE_DOC_ICI") {
      console.warn("⚠️ ATTENTION: Veuillez remplacer 'ID_VOTRE_TEMPLATE_DOC_ICI' par un ID de template Google Doc valide dans la fonction _test_generation_doc().");
      SpreadsheetApp.getUi().alert("⚠️ Test non exécuté : Veuillez configurer un ID de template Google Doc valide dans la fonction _test_generation_doc() du script.");
      return;
    }

    console.log("Paramètres de test:", JSON.stringify(parametresTest, null, 2));
    const resultat = _genererEtiquettesDepuisDoc(parametresTest);
    console.log("Résultat de la génération (Docs):", JSON.stringify(resultat, null, 2));
    console.log("✅ TEST GÉNÉRATION DOCS TERMINÉ AVEC SUCCÈS ===");
  } catch (error) {
    console.error("❌ ERREUR LORS DU TEST GÉNÉRATION DOCS:", error.toString(), error.stack);
    SpreadsheetApp.getUi().alert(`❌ Erreur lors du test de génération Docs:\n\n${error.toString()}`);
  }
}

<!-- END OF FILE: Code.gs -->