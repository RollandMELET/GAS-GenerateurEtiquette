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
// CONFIGURATION
// ========================================

// L'ID du template sera lu depuis la cellule B8 du Google Sheets

function genererEtiquettes() {
  console.log("=== DÉBUT GÉNÉRATION ÉTIQUETTES AVEC GOOGLE DOCS ===");
  
  try {
    // 1. Lire les paramètres depuis le Google Sheet
    console.log("1. Lecture des paramètres...");
    const sheet = SpreadsheetApp.getActiveSheet();
    const serie = sheet.getRange('B2').getValue() || 'ENVELOPPE';
    const numeroDebut = parseInt(sheet.getRange('B3').getValue()) || 1;
    const nbPages = parseInt(sheet.getRange('B4').getValue()) || 1;
    const templateId = sheet.getRange('B8').getValue();
    
    console.log(`Paramètres: série=${serie}, numeroDebut=${numeroDebut}, nbPages=${nbPages}`);
    console.log(`Template ID: ${templateId}`);
    
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
    
    // 3. Date au format français
    const maintenant = new Date();
    const dateFormatee = formatDateFrancais(maintenant);
    console.log(`Date: ${dateFormatee}`);
    
    // 4. Créer le dossier de destination (même que le template)
    const dossier = creerDossierSiNecessaire(templateId);
    
    // 5. Générer un seul document multi-pages
    console.log("--- Création du document multi-pages ---");
    
    const premierNum = formatNumero(numeroDebut);
    const dernierNum = formatNumero(numeroDebut + (nbPages * 5) - 1);
    const nomFichier = `Etiquettes_${serie}_${premierNum}-${dernierNum}_${dateFormatee.replace(/\//g, '-')}`;
    
    // Créer des documents séparés pour chaque page
    const documentsPages = [];
    
    for (let page = 0; page < nbPages; page++) {
      console.log(`--- Création page ${page + 1}/${nbPages} ---`);
      
      // Calculer les numéros pour cette page
      const baseNum = numeroDebut + (page * 5);
      const numeros = [
        formatNumero(baseNum),
        formatNumero(baseNum + 1),
        formatNumero(baseNum + 2),
        formatNumero(baseNum + 3),
        formatNumero(baseNum + 4)
      ];
      
      console.log(`Numéros page ${page + 1}: ${numeros.join(', ')}`);
      
      // Dupliquer le template pour cette page
      const templateDoc = DriveApp.getFileById(templateId);
      const docPageCopie = templateDoc.makeCopy(`${nomFichier}_page${page + 1}_temp`, dossier);
      
      // Ouvrir et modifier le document
      const docPage = DocumentApp.openById(docPageCopie.getId());
      const bodyPage = docPage.getBody();
      
      // Remplacer les placeholders
      bodyPage.replaceText('{{SERIE}}', serie);
      bodyPage.replaceText('{{NUMERO1}}', numeros[0]);
      bodyPage.replaceText('{{NUMERO2}}', numeros[1]);
      bodyPage.replaceText('{{NUMERO3}}', numeros[2]);
      bodyPage.replaceText('{{NUMERO4}}', numeros[3]);
      bodyPage.replaceText('{{NUMERO5}}', numeros[4]);
      bodyPage.replaceText('{{DateJour}}', dateFormatee);
      
      // Sauvegarder
      docPage.saveAndClose();
      
      documentsPages.push(docPageCopie);
    }
    
    // Fusionner tous les documents en un seul
    console.log("--- Fusion des pages ---");
    const docFinal = DocumentApp.openById(documentsPages[0].getId());
    const bodyFinal = docFinal.getBody();
    
    // Ajouter les pages suivantes
    for (let i = 1; i < documentsPages.length; i++) {
      console.log(`Ajout de la page ${i + 1}...`);
      
      // Ajouter un saut de page
      bodyFinal.appendPageBreak();
      
      // Ouvrir le document à fusionner
      const docSource = DocumentApp.openById(documentsPages[i].getId());
      const bodySource = docSource.getBody();
      
      // Copier tous les éléments
      const numElements = bodySource.getNumChildren();
      for (let j = 0; j < numElements; j++) {
        const element = bodySource.getChild(j);
        const elementType = element.getType();
        
        if (elementType === DocumentApp.ElementType.TABLE) {
          bodyFinal.appendTable(element.asTable().copy());
        } else if (elementType === DocumentApp.ElementType.PARAGRAPH) {
          const para = element.asParagraph();
          const newPara = bodyFinal.appendParagraph(para.getText());
          // Copier le formatage
          newPara.setAttributes(para.getAttributes());
        }
      }
      
      docSource.saveAndClose();
    }
    
    // Renommer le document final
    docFinal.setName(nomFichier + '_temp');
    docFinal.saveAndClose();
    console.log("Document multi-pages créé et sauvegardé");
    
    // Convertir en PDF
    console.log("Conversion en PDF...");
    const pdfBlob = documentsPages[0].getAs('application/pdf');
    pdfBlob.setName(nomFichier + '.pdf');
    
    // Sauvegarder le PDF
    const pdfFile = dossier.createFile(pdfBlob);
    console.log(`PDF multi-pages créé: ${pdfFile.getUrl()}`);
    
    // Supprimer tous les documents temporaires
    console.log("Suppression des documents temporaires...");
    documentsPages.forEach(doc => {
      DriveApp.getFileById(doc.getId()).setTrashed(true);
    });
    
    const fichiersGeneres = [{
      page: 'Multi-pages',
      url: pdfFile.getUrl(),
      nom: pdfFile.getName()
    }];
    
    // 6. Message de succès
    const message = `✅ PDF d'étiquettes généré !\n\nSérie: ${serie}\nNuméros: ${formatNumero(numeroDebut)} à ${formatNumero(numeroDebut + (nbPages * 5) - 1)}\nDate: ${dateFormatee}\nPages: ${nbPages}\n\n📁 Fichier sauvé dans le même dossier que votre template\n\n🖨️ Prêt à imprimer directement !`;
    
    console.log("Affichage du message de succès");
    SpreadsheetApp.getUi().alert(message);
    
    console.log("=== FIN GÉNÉRATION RÉUSSIE ===");
    return fichiersGeneres;
    
  } catch (error) {
    console.error("ERREUR:", error.toString());
    console.error("Stack:", error.stack);
    
    SpreadsheetApp.getUi().alert(`❌ Erreur lors de la génération:\n\n${error.toString()}${error.toString().includes('TEMPLATE_DOC_ID') ? '\n\n⚠️ Vérifiez que l\'ID du template est correctement configuré dans le script.' : ''}`);
    throw error;
  }
}

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
  // Obtenir le dossier parent du template
  const templateDoc = DriveApp.getFileById(templateId);
  const dossiersParents = templateDoc.getParents();
  
  if (dossiersParents.hasNext()) {
    const dossierParent = dossiersParents.next();
    console.log(`PDFs seront sauvés dans le même dossier que le template: ${dossierParent.getName()}`);
    return dossierParent;
  } else {
    // Fallback: utiliser la racine du Drive
    console.log("Template dans la racine - utilisation de la racine du Drive");
    return DriveApp.getRootFolder();
  }
}

// Fonction de test pour vérifier la configuration
function testerConfiguration() {
  console.log("=== TEST DE CONFIGURATION ===");
  
  try {
    // Lire l'ID du template depuis B8
    const sheet = SpreadsheetApp.getActiveSheet();
    const templateId = sheet.getRange('B8').getValue();
    
    if (!templateId || templateId.toString().trim() === '') {
      throw new Error('ID du template manquant dans la cellule B8');
    }
    
    // Vérifier l'accès au template
    const templateDoc = DriveApp.getFileById(templateId);
    console.log(`Template trouvé: ${templateDoc.getName()}`);
    console.log(`Template ID: ${templateId}`);
    
    // Vérifier l'accès au sheet
    const serie = sheet.getRange('B2').getValue();
    console.log(`Série lue: ${serie}`);
    
    console.log("✅ Configuration OK - Le script devrait fonctionner");
    SpreadsheetApp.getUi().alert(`✅ Configuration OK - Vous pouvez générer les étiquettes !`);
    
  } catch (error) {
    console.error("❌ Erreur de configuration:", error.toString());
    SpreadsheetApp.getUi().alert(`❌ Problème détecté:\n\n${error.toString()}\n\nVérifiez l'ID du template dans le script.`);
  }
}

// Bouton pour créer un template exemple
function creerTemplateExemple() {
  console.log("Création d'un template exemple...");
  
  // Créer un nouveau document
  const doc = DocumentApp.create('Template Étiquettes DUHALDE - EXEMPLE');
  
  // Configuration de la page
  const body = doc.getBody();
  body.setPageWidth(595.28).setPageHeight(841.89); // A4 en points
  body.setMarginTop(0).setMarginBottom(0).setMarginLeft(0).setMarginRight(0);
  
  // Ajouter le contenu du template
  body.appendParagraph('TEMPLATE ÉTIQUETTES DUHALDE').setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph('');
  body.appendParagraph('Instructions:');
  body.appendParagraph('1. Créez un tableau 2 colonnes x 5 lignes');
  body.appendParagraph('2. Taille des cellules: 105mm x 57mm');
  body.appendParagraph('3. Dans chaque cellule, mettez:');
  body.appendParagraph('   DUHALDE INDUSTRIES');
  body.appendParagraph('   {{SERIE}}');
  body.appendParagraph('   {{NUMERO1}} (ou NUMERO2, NUMERO3, etc.)');
  body.appendParagraph('   Pour Fiche Suiveuse (ou Pour Pièces BETON)');
  body.appendParagraph('   {{DateJour}}');
  body.appendParagraph('');
  body.appendParagraph('Une fois créé, copiez l\'ID du document et mettez-le dans le script à la ligne TEMPLATE_DOC_ID');
  
  const url = doc.getUrl();
  console.log(`Template exemple créé: ${url}`);
  
  SpreadsheetApp.getUi().alert(`Template exemple créé:\n\n${url}\n\nUtilisez-le comme référence pour créer votre template.`);
  
  return url;
}