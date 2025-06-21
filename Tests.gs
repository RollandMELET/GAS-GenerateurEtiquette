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
// <!-- END OF FILE: Tests.gs -->