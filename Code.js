<!-- START OF FILE: Code.gs -->

// FILENAME: Code.gs
// Version: 2.1.0
// Date: 2025-06-21 18:34
// Author: Rolland MELET & AI Senior Coder
// Description: Ajout de la logique de génération d'étiquettes avec QR-Codes dynamiques via la bibliothèque Api360sc. Agit comme aiguilleur principal.

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

Version: 2.1.0
Développé pour: 360SmartConnect

Fonctionnalités:
✅ Génération d'étiquettes avec QR-Codes dynamiques
✅ Gestion centralisée des templates
✅ Support de templates Google Docs et Google Slides
✅ Support 4 séries: ENVELOPPE, TOIT, DALLE, TEST
✅ Export PDF prêt à imprimer
✅ Sauvegarde dans le dossier du template ou un dossier spécifié
✅ Message de succès interactif avec liens directs

Support: Vérifiez la configuration dans votre Google Sheet (Feuille 'ListeTemplates').`;

  SpreadsheetApp.getUi().alert(message);
}

// ========================================
// FONCTION PRINCIPALE (DISPATCHER)
// ========================================

function genererEtiquettes() {
  console.log("=== DÉBUT GÉNÉRATION ÉTIQUETTES ===");
  const ui = SpreadsheetApp.getUi();
  
  try {
    // 1. Lire les paramètres depuis le Google Sheet
    console.log("1. Lecture des paramètres...");
    const sheet = SpreadsheetApp.getActiveSheet();
    const parametres = {
      serie: sheet.getRange('B2').getValue() || 'ENVELOPPE',
      numeroDebut: parseInt(sheet.getRange('B3').getValue()) || 1,
      nbPages: parseInt(sheet.getRange('B4').getValue()) || 1,
      nomTemplate: sheet.getRange('B5').getValue().toString().trim(),
      templateId: sheet.getRange('B8').getValue().toString().trim(),
      typeTemplate: sheet.getRange('B9').getValue().toString().trim(),
      dossierId: sheet.getRange('B10').getValue().toString().trim(),
      useQrCode: sheet.getRange('B11').getValue() === 'Oui',
      environnementApi: sheet.getRange('B12').getValue().toString().trim()
    };
    
    // 2. Validation des paramètres
    console.log("2. Validation des paramètres...");
    if (!['ENVELOPPE', 'TOIT', 'DALLE', 'TEST'].includes(parametres.serie)) {
      throw new Error('Série doit être : ENVELOPPE, TOIT, DALLE ou TEST');
    }
    if (parametres.numeroDebut < 1 || parametres.nbPages < 1) {
      throw new Error('N° début et Nb pages doivent être >= 1');
    }
    if (!parametres.templateId) {
      throw new Error('ID du template manquant en B8. Avez-vous choisi un template valide en B5 ?');
    }
    if (parametres.useQrCode && !parametres.environnementApi) {
      throw new Error("Le template est configuré pour utiliser des QR-Codes, mais l'environnement API n'est pas spécifié en B12.");
    }
    if (parametres.useQrCode && parametres.typeTemplate !== 'Google Slide') {
      throw new Error("La génération de QR-Codes est uniquement supportée pour les templates de type 'Google Slide'.");
    }
    console.log(`Paramètres lus et validés: ${JSON.stringify(parametres, null, 2)}`);

    // 3. Aiguillage vers la bonne fonction de génération
    console.log("3. Aiguillage vers la fonction de génération appropriée...");
    let fichiersGeneres;
    if (parametres.useQrCode) {
      fichiersGeneres = _genererEtiquettesAvecQRCode(parametres);
    } else if (parametres.typeTemplate === 'Google Doc') {
      fichiersGeneres = _genererEtiquettesDepuisDoc(parametres);
    } else if (parametres.typeTemplate === 'Google Slide') {
      fichiersGeneres = _genererEtiquettesDepuisSlide(parametres);
    } else {
      throw new Error('Type de template non supporté : ' + parametres.typeTemplate);
    }

    // 4. Message de succès
    console.log("4. Préparation du message de succès...");
    const pdfUrl = fichiersGeneres[0].url;
    const pdfName = fichiersGeneres[0].nom;
    const dossier = creerDossierSiNecessaire(parametres);
    const dossierUrl = dossier.getUrl();
    const dossierName = dossier.getName();

    const htmlMessage = `
      <style>
        body { font-family: Arial, sans-serif; font-size: 14px; }
        a { color: #1a73e8; text-decoration: none; }
        a:hover { text-decoration: underline; }
        .button { display: inline-block; padding: 10px 20px; font-size: 16px; font-weight: bold; color: #fff; background-color: #4CAF50; border: none; border-radius: 5px; text-align: center; cursor: pointer; }
        .button:hover { background-color: #45a049; }
      </style>
      <h2>✅ PDF d'étiquettes généré !</h2>
      <p>
        <b>Série:</b> ${parametres.serie}<br>
        <b>Numéros:</b> ${formatNumero(parametres.numeroDebut)} à ${formatNumero(parametres.numeroDebut + (parametres.nbPages * 5) - 1)}<br>
        <b>Pages:</b> ${parametres.nbPages}
      </p>
      <p>Le fichier <b><a href="${pdfUrl}" target="_blank" title="Cliquez pour ouvrir le PDF">${pdfName}</a></b> a été sauvegardé dans le dossier :<br><b><a href="${dossierUrl}" target="_blank" title="Cliquez pour ouvrir le dossier">${dossierName}</a></b></p>
      <br><a href="${pdfUrl}" target="_blank" class="button">Ouvrir le PDF</a>`;

    const htmlOutput = HtmlService.createHtmlOutput(htmlMessage).setWidth(450).setHeight(300);
    console.log("Affichage du message de succès avec lien cliquable.");
    ui.showModalDialog(htmlOutput, "Génération Réussie");

    console.log("=== FIN GÉNÉRATION RÉUSSIE ===");
    return fichiersGeneres;
    
  } catch (error) {
    console.error("ERREUR:", error.toString());
    console.error("Stack:", error.stack);
    ui.alert(`❌ Erreur lors de la génération:\n\n${error.toString()}`);
    throw error;
  }
}
<!-- END OF FILE: Code.gs -->