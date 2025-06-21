// <!-- START OF FILE: Tests.gs -->
// FILENAME: Tests.gs
// Version: 1.0.0
// Date: 2025-06-21 18:34
// Author: Rolland MELET & AI Senior Coder
// Description: Fonctions de test dédiées pour la génération d'étiquettes, isolées du code de production.

// ========================================
// FONCTIONS DE TEST INTERNES (POUR DÉVELOPPEMENT)
// ========================================

/**
 * Fonction de test pour la génération d'étiquettes depuis un template Google Docs.
 */
function _test_generation_doc() {
  console.log("=== DÉBUT TEST GÉNÉRATION DOCS ===");
  try {
    const parametresTest = {
      serie: "TEST_DOC",
      numeroDebut: 101,
      nbPages: 2,
      templateId: "ID_VOTRE_TEMPLATE_DOC_ICI", // ⚠️ REMPLACEZ PAR UN VRAI ID DE TEMPLATE DOC
      nomTemplate: "Template Test Doc",
      typeTemplate: "Google Doc",
      dossierId: null 
    };

    if (parametresTest.templateId.includes("VOTRE_TEMPLATE")) {
      SpreadsheetApp.getUi().alert("⚠️ Test non exécuté : Veuillez configurer un ID de template Google Doc valide dans la fonction _test_generation_doc() du script.");
      return;
    }
    _genererEtiquettesDepuisDoc(parametresTest);
    console.log("✅ TEST GÉNÉRATION DOCS TERMINÉ AVEC SUCCÈS ===");
  } catch (error) {
    console.error("❌ ERREUR LORS DU TEST GÉNÉRATION DOCS:", error.toString(), error.stack);
  }
}

/**
 * Fonction de test pour la génération d'étiquettes depuis un template Google Slides.
 */
function _test_generation_slide() {
  console.log("=== DÉBUT TEST GÉNÉRATION SLIDES ===");
  try {
    const parametresTest = {
      serie: "TEST_SLIDE",
      numeroDebut: 201,
      nbPages: 3,
      templateId: "ID_VOTRE_TEMPLATE_SLIDE_ICI", // ⚠️ REMPLACEZ PAR UN VRAI ID DE TEMPLATE SLIDE
      nomTemplate: "Template Test Slide",
      typeTemplate: "Google Slide",
      dossierId: null
    };

    if (parametresTest.templateId.includes("VOTRE_TEMPLATE")) {
       SpreadsheetApp.getUi().alert("⚠️ Test non exécuté : Veuillez configurer un ID de template Google Slide valide dans la fonction _test_generation_slide() du script.");
      return;
    }
    _genererEtiquettesDepuisSlide(parametresTest);
    console.log("✅ TEST GÉNÉRATION SLIDES TERMINÉ AVEC SUCCÈS ===");
  } catch (error) {
    console.error("❌ ERREUR LORS DU TEST GÉNÉRATION SLIDES:", error.toString(), error.stack);
  }
}


// ========================================
// FONCTION DE DIAGNOSTIC DE TEMPLATE
// ========================================

/**
 * [NOUVEAU] Analyse la première diapositive d'un template Google Slides
 * et liste tous ses éléments, leur type, et leur titre.
 * Très utile pour déboguer les problèmes de placeholders non trouvés.
 */
function _diagnostiquerTemplateSlide() {
  const ui = SpreadsheetApp.getUi();
  
  // 1. Demander l'ID du template à l'utilisateur
  const response = ui.prompt(
    'Diagnostic de Template Slide',
    'Veuillez coller l\'ID de votre présentation Google Slides :',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK || !response.getResponseText()) {
    ui.alert('Opération annulée.');
    return;
  }
  
  const templateId = response.getResponseText().trim();
  console.log(`Lancement du diagnostic pour le template ID: ${templateId}`);

  try {
    // 2. Ouvrir la présentation et accéder à la première diapositive
    const presentation = SlidesApp.openById(templateId);
    const slide = presentation.getSlides()[0];
    
    if (!slide) {
      console.error("Aucune diapositive trouvée dans cette présentation.");
      ui.alert("Aucune diapositive trouvée dans cette présentation.");
      return;
    }

    // 3. Parcourir tous les éléments de la diapositive
    const pageElements = slide.getPageElements();
    console.log(`Analyse de ${pageElements.length} éléments sur la diapositive...`);

    if (pageElements.length === 0) {
      ui.alert("La diapositive est vide !");
      return;
    }

    pageElements.forEach((element, index) => {
      const elementType = element.getPageElementType();
      let title = "(aucun titre)";
      
      // Essayer de récupérer le titre, si l'élément en a un
      if (typeof element.getTitle === 'function') {
        title = element.getTitle() || "(titre vide)";
      }
      
      console.log(`Élément ${index + 1}: Type = ${elementType}, Titre = "${title}"`);
    });

    console.log("--- Diagnostic Terminé ---");
    ui.alert("Diagnostic terminé. Veuillez consulter les journaux (Afficher > Journaux) pour voir la liste des éléments.");

  } catch (e) {
    console.error(`Erreur lors du diagnostic: ${e.toString()}`);
    ui.alert(`Erreur lors du diagnostic:\n\n${e.toString()}\n\nVérifiez que l'ID est correct et que vous avez accès au fichier.`);
  }
}

// <!-- END OF FILE: Tests.gs -->