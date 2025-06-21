// <!-- START OF FILE: Utils.gs -->
// FILENAME: Utils.gs
// Version: 1.1.0
// Date: 2025-06-21 18:34
// Author: Rolland MELET & AI Senior Coder
// Description: Ajout d'une fonction de formatage date-heure pour les noms de fichiers.

// ========================================
// FONCTIONS UTILITAIRES
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

/**
 * [NOUVEAU] Formate une date en chaîne AAAA-MM-JJ_HH-mm-ss pour les noms de fichiers.
 * @param {Date} date - L'objet Date à formater.
 * @returns {string} La date formatée.
 */
function formatDateTimePourNomFichier(date) {
  const annee = date.getFullYear();
  const mois = (date.getMonth() + 1).toString().padStart(2, '0');
  const jour = date.getDate().toString().padStart(2, '0');
  const heure = date.getHours().toString().padStart(2, '0');
  const minute = date.getMinutes().toString().padStart(2, '0');
  const seconde = date.getSeconds().toString().padStart(2, '0');
  return `${annee}-${mois}-${jour}_${heure}-${minute}-${seconde}`;
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
    const nomTemplate = sheet.getRange('B5').getValue().toString().trim();
    const templateId = sheet.getRange('B8').getValue().toString().trim();
    const typeTemplate = sheet.getRange('B9').getValue().toString().trim();
    const dossierId = sheet.getRange('B10').getValue().toString().trim();

    if (!templateId) {
      throw new Error('ID du template manquant en B8. Avez-vous choisi un template valide en B5 ?');
    }

    let templateFile;
    let alertMessage = `✅ Configuration OK pour "${nomTemplate}".`;

    if (typeTemplate === 'Google Doc') {
      templateFile = DriveApp.getFileById(templateId);
      DocumentApp.openById(templateId);
    } else if (typeTemplate === 'Google Slide') {
      templateFile = DriveApp.getFileById(templateId);
      SlidesApp.openById(templateId);
    } else {
      throw new Error(`Type de template invalide en B9. Il doit être "Google Doc" ou "Google Slide".`);
    }
    alertMessage += `\nType: ${typeTemplate}\nTemplate: ${templateFile.getName()}`;

    if (dossierId) {
      const dossierDestination = DriveApp.getFolderById(dossierId);
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
  body.appendParagraph('Placeholders texte à utiliser: {{SERIE}}, {{NUMERO1}} à {{NUMERO5}}, {{DateJour}}');
  body.appendParagraph("\nPour les QR-Codes (sur un template Google Slide), insérez une forme carrée et nommez-la (Texte alternatif > Titre) 'QR_CODE_PLACEHOLDER_1', etc.");
  
  const url = doc.getUrl();
  SpreadsheetApp.getUi().alert(`Template exemple créé:\n\n${url}`);
}
// <!-- END OF FILE: Utils.gs -->