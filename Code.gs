<!-- START OF FILE: Code.gs -->

// Version: 1.1.0
// Date: 2025-06-20 12:00
// Author: Rolland MELET
// Description: Refactorisation pour isoler la logique de génération Docs dans une fonction dédiée (_genererEtiquettesDepuisDoc) 
// et transformer genererEtiquettes en dispatcher. Prépare l'ajout du support pour Google Slides.

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
✅ Support 3 séries: ENVELOPPE, TOIT, DALLE
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
    const templateId = sheet.getRange('B8').getValue();
    
    // 2. Validation
    if (!['ENVELOPPE', 'TOIT', 'DALLE'].includes(serie)) {
      throw new Error('Série doit être : ENVELOPPE, TOIT ou DALLE');
    }
    if (numeroDebut < 1 || nbPages < 1) {
      throw new Error('N° début et Nb pages doivent être >= 1');
    }
    if (!templateId || templateId.toString().trim() === '') {
      throw new Error('ID du template manquant dans la cellule B8');
    }

    // Regrouper les paramètres dans un objet pour les passer aux fonctions de génération
    const parametres = {
      serie: serie,
      numeroDebut: numeroDebut,
      nbPages: nbPages,
      templateId: templateId
    };
    
    console.log(`Paramètres: série=${parametres.serie}, numeroDebut=${parametres.numeroDebut}, nbPages=${parametres.nbPages}`);
    console.log(`Template ID: ${parametres.templateId}`);

    // Pour l'instant, on appelle directement la fonction pour Google Docs.
    // Plus tard, nous ajouterons une condition ici pour appeler la fonction pour Slides.
    const fichiersGeneres = _genererEtiquettesDepuisDoc(parametres);

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
  
  // 3. Date au format français
  const maintenant = new Date();
  const dateFormatee = formatDateFrancais(maintenant);
  console.log(`Date: ${dateFormatee}`);
    
  // 4. Créer le dossier de destination (même que le template)
  const dossier = creerDossierSiNecessaire(parametres.templateId);
  
  // 5. Générer un seul document multi-pages
  console.log("--- Création du document multi-pages ---");
  
  const premierNum = formatNumero(parametres.numeroDebut);
  const dernierNum = formatNumero(parametres.numeroDebut + (parametres.nbPages * 5) - 1);
  const nomFichier = `Etiquettes_${parametres.serie}_${premierNum}-${dernierNum}_${dateFormatee.replace(/\//g, '-')}`;
  
  const documentsPages = [];
  
  for (let page = 0; page < parametres.nbPages; page++) {
    console.log(`--- Création page ${page + 1}/${parametres.nbPages} ---`);
    const baseNum = parametres.numeroDebut + (page * 5);
    const numeros = [
      formatNumero(baseNum),
      formatNumero(baseNum + 1),
      formatNumero(baseNum + 2),
      formatNumero(baseNum + 3),
      formatNumero(baseNum + 4)
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
  
  console.log("Conversion en PDF...");
  const pdfBlob = documentsPages[0].getAs('application/pdf');
  pdfBlob.setName(nomFichier + '.pdf');
  const pdfFile = dossier.createFile(pdfBlob);
  console.log(`PDF multi-pages créé: ${pdfFile.getUrl()}`);
  
  console.log("Suppression des documents temporaires...");
  documentsPages.forEach(doc => {
    DriveApp.getFileById(doc.getId()).setTrashed(true);
  });
  
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

function creerDossierSiNecessaire(templateId) {
  const templateFile = DriveApp.getFileById(templateId);
  const dossiersParents = templateFile.getParents();
  
  if (dossiersParents.hasNext()) {
    const dossierParent = dossiersParents.next();
    console.log(`Dossier de destination: ${dossierParent.getName()}`);
    return dossierParent;
  } else {
    console.log("Template dans la racine du Drive. Utilisation de la racine.");
    return DriveApp.getRootFolder();
  }
}

function testerConfiguration() {
  console.log("=== TEST DE CONFIGURATION ===");
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const templateId = sheet.getRange('B8').getValue();
    if (!templateId || templateId.toString().trim() === '') {
      throw new Error('ID du template manquant dans la cellule B8');
    }
    // Pour l'instant, on ne teste que pour Doc. Sera adapté en phase 3.
    const templateDoc = DriveApp.getFileById(templateId);
    console.log(`Template trouvé: ${templateDoc.getName()}`);
    console.log(`Template ID: ${templateId}`);
    
    SpreadsheetApp.getUi().alert(`✅ Configuration OK - Le script peut accéder au template.`);
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

<!-- END OF FILE: Code.gs -->