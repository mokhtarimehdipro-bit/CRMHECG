// =============================================================================
//  BACHELOR CDSP — TESTS D'ADMISSION (sujet + corrigé), mise en forme HECG (Google Docs)
//  QCM (20 questions, /8) + exercice rédactionnel (/6) + exercice calculatoire (/6) = /20
//  Les taux de l'exercice de calcul sont DONNÉS dans l'énoncé (taux simplifiés) : le sujet
//  reste juste quelle que soit l'évolution des taux réels.
// =============================================================================
var FP2_QCM = [
  ["Le bulletin de paie doit être remis au salarié :", ["Une fois par an", "À chaque paiement du salaire", "Uniquement à sa demande", "À la fin du contrat"], 1],
  ["La DSN (Déclaration Sociale Nominative) est transmise :", ["Une fois par an", "Chaque trimestre", "Chaque mois", "Uniquement à l'embauche"], 2],
  ["Le salaire brut diminué des cotisations salariales donne :", ["Le coût employeur", "Le net avant impôt sur le revenu", "Le salaire de base", "Le net imposable majoré"], 1],
  ["La durée légale du travail pour un salarié à temps plein est de :", ["32 heures par semaine", "35 heures par semaine", "39 heures par semaine", "40 heures par semaine"], 1],
  ["Sans accord collectif, les 8 premières heures supplémentaires de la semaine sont majorées de :", ["10 %", "25 %", "50 %", "100 %"], 1],
  ["Un salarié acquiert, par mois de travail effectif, des congés payés à hauteur de :", ["1,5 jour ouvrable", "2 jours ouvrables", "2,5 jours ouvrables", "3 jours ouvrables"], 2],
  ["Le prélèvement à la source de l'impôt sur le revenu s'applique sur :", ["Le salaire brut", "Le net imposable", "Le coût employeur", "Le salaire de base"], 1],
  ["La déclaration préalable à l'embauche (DPAE) doit être faite :", ["Avant l'embauche", "Dans le mois qui suit l'embauche", "À la fin de la période d'essai", "Uniquement pour les CDD"], 0],
  ["Pour un employé en CDI, la durée maximale initiale de la période d'essai prévue par la loi est de :", ["1 mois", "2 mois", "3 mois", "4 mois"], 1],
  ["Le SMIC est revalorisé au minimum :", ["Tous les 1er janvier", "Tous les 5 ans", "Uniquement par décision de l'entreprise", "Tous les 1er septembre"], 0],
  ["Une cotisation patronale est :", ["Payée par le salarié", "À la charge de l'employeur", "Remboursée au salarié", "Payée par l'État"], 1],
  ["Le registre unique du personnel mentionne :", ["Uniquement les cadres", "Tous les salariés, dans l'ordre des embauches", "Uniquement les salariés en CDD", "Les candidats non retenus"], 1],
  ["En cas d'arrêt maladie, le salarié doit transmettre l'arrêt à son employeur dans un délai de :", ["24 heures", "48 heures", "8 jours", "15 jours"], 1],
  ["Le reçu pour solde de tout compte :", ["Récapitule les sommes versées à la fin du contrat", "Remplace le contrat de travail", "Est remis à l'embauche", "Est établi par France Travail"], 0],
  ["Les données de paie d'un salarié :", ["Peuvent être consultées par tous les collègues", "Sont des données personnelles à accès restreint (RGPD)", "Sont publiques", "Doivent être affichées dans l'entreprise"], 1],
  ["Pour la complémentaire santé d'entreprise, l'employeur finance au minimum :", ["25 % de la cotisation", "50 % de la cotisation", "75 % de la cotisation", "100 % de la cotisation"], 1],
  ["Un avantage en nature (ex. véhicule de fonction utilisé à titre privé) :", ["N'apparaît pas sur le bulletin", "Est intégré au brut et soumis à cotisations", "Est versé en net uniquement", "Est interdit"], 1],
  ["L'indemnité de congés payés est calculée selon :", ["Le maintien de salaire uniquement", "La règle du dixième uniquement", "La méthode la plus favorable au salarié entre le maintien et le dixième", "Un forfait fixé par l'employeur"], 2],
  ["Un salarié à temps partiel travaille :", ["Plus de 35 heures", "Moins que la durée légale ou conventionnelle", "Uniquement le week-end", "Sans contrat écrit"], 1],
  ["Le taux horaire d'un salarié mensualisé à 35 h se calcule en divisant le salaire mensuel par :", ["140", "151,67", "160", "169"], 1]
];

function _fp2DocEntete_(body, titre, sousTitre) {
  var t = body.appendParagraph("HECG — École de commerce et de gestion");
  t.setFontSize(10).setForegroundColor("#F26522").setBold(true);
  var h = body.appendParagraph(titre); h.setHeading(DocumentApp.ParagraphHeading.TITLE);
  h.editAsText().setFontSize(22).setForegroundColor("#1A4E8A").setBold(true);
  var s = body.appendParagraph(sousTitre); s.editAsText().setFontSize(11).setForegroundColor("#3b4859").setBold(false);
  body.appendHorizontalRule();
}
function _fp2DocTitre_(body, txt) { var p = body.appendParagraph(txt); p.setHeading(DocumentApp.ParagraphHeading.HEADING1); p.editAsText().setForegroundColor("#1A4E8A").setFontSize(15).setBold(true); return p; }
function _fp2DocTexte_(body, txt, gras) { var p = body.appendParagraph(txt); p.editAsText().setFontSize(11).setForegroundColor("#0e1a2b").setBold(!!gras); return p; }
function _fp2DocTable_(body, rows) {
  var t = body.appendTable(rows);
  t.setBorderColor("#cdd8e6");
  var e = t.getRow(0); for (var c = 0; c < e.getNumCells(); c++) { e.getCell(c).setBackgroundColor("#1A4E8A"); e.getCell(c).editAsText().setForegroundColor("#ffffff").setBold(true).setFontSize(10); }
  for (var r = 1; r < t.getNumRows(); r++) for (var c2 = 0; c2 < t.getRow(r).getNumCells(); c2++) t.getRow(r).getCell(c2).editAsText().setFontSize(10).setForegroundColor("#0e1a2b").setBold(false);
  return t;
}

function _fp2TestsBachelor_() {
  var props = PropertiesService.getScriptProperties();
  var parent; try { var it = DriveApp.getFoldersByName("Bachelor CDSP — Tests d'admission"); parent = it.hasNext() ? it.next() : DriveApp.createFolder("Bachelor CDSP — Tests d'admission"); } catch (e) { parent = DriveApp.getRootFolder(); }
  var L = "ABCD";
  // ---------- SUJET ----------
  var ds = DocumentApp.create("Bachelor CDSP — Test d'admission — Sujet"), b = ds.getBody();
  b.setMarginTop(42).setMarginBottom(42).setMarginLeft(50).setMarginRight(50);
  _fp2DocEntete_(b, "Test d'admission — Bachelor CDSP", "Ressources humaines et paie · Durée : 1 h 30 · Calculatrice autorisée · Aucun document");
  _fp2DocTable_(b, [["Nom", "Prénom", "Date", "Note /20"], ["", "", "", ""]]);
  _fp2DocTexte_(b, "Le test comporte trois parties : un QCM (8 points), un exercice rédactionnel (6 points) et un exercice de calcul de paie (6 points).");
  _fp2DocTitre_(b, "Partie 1 — QCM (8 points)");
  _fp2DocTexte_(b, "Pour chaque question, entourez une seule réponse. Bonne réponse : 0,4 point ; pas de point négatif.");
  FP2_QCM.forEach(function (q, i) {
    _fp2DocTexte_(b, (i + 1) + ". " + q[0], true);
    q[1].forEach(function (rep, k) { var li = b.appendListItem(L[k] + ") " + rep); li.setGlyphType(DocumentApp.GlyphType.HOLLOW_BULLET); li.editAsText().setFontSize(10.5).setBold(false); });
  });
  b.appendPageBreak();
  _fp2DocTitre_(b, "Partie 2 — Exercice rédactionnel (6 points)");
  _fp2DocTexte_(b, "Situation : vous êtes assistant(e) paie dans un cabinet. M. Martin, salarié d'une entreprise cliente, vous écrit qu'il a effectué 6 heures supplémentaires en mars et qu'elles n'apparaissent pas sur son bulletin de paie. Il est mécontent et demande une régularisation immédiate.");
  _fp2DocTexte_(b, "Rédigez la réponse par courriel (15 à 20 lignes) : accusé de réception, vérifications que vous allez mener, pièces que vous lui demandez, délai et modalités de régularisation, formule de politesse. Soignez le ton, la clarté et la confidentialité.");
  for (var l = 0; l < 18; l++) b.appendParagraph("_____________________________________________________________________________").editAsText().setForegroundColor("#c3cbd6");
  b.appendPageBreak();
  _fp2DocTitre_(b, "Partie 3 — Exercice de calcul de paie (6 points)");
  _fp2DocTexte_(b, "Mme Durand est employée à temps plein (35 h, soit 151,67 h par mois). Données de mars :");
  _fp2DocTable_(b, [["Élément", "Valeur"], ["Salaire de base mensuel brut", "2 000,00 €"], ["Heures supplémentaires effectuées", "6 heures, majorées à 25 %"], ["Prime d'objectif", "150,00 €"],
    ["Taux global (simplifié) des cotisations salariales", "22 % du brut"], ["Taux global (simplifié) des cotisations patronales", "42 % du brut"], ["Taux de prélèvement à la source", "4,5 % (appliqué au net avant impôt, simplification)"],
    ["Ancienneté au 31 mars", "7 mois de travail effectif"]]);
  _fp2DocTexte_(b, "Questions (arrondir au centime) :", true);
  ["1. Calculez le taux horaire brut (au centime).", "2. Calculez le montant des heures supplémentaires.", "3. Calculez le salaire brut du mois.", "4. Calculez les cotisations salariales et le net avant impôt.",
    "5. Calculez le prélèvement à la source et le net payé.", "6. Calculez le coût total employeur.", "7. Combien de jours ouvrables de congés payés Mme Durand a-t-elle acquis à fin mars ?"].forEach(function (q) { _fp2DocTexte_(b, q); });
  ds.saveAndClose();

  // ---------- CORRIGÉ ----------
  var dc = DocumentApp.create("Bachelor CDSP — Test d'admission — Corrigé (jury)"), c = dc.getBody();
  c.setMarginTop(42).setMarginBottom(42).setMarginLeft(50).setMarginRight(50);
  _fp2DocEntete_(c, "Corrigé — Test d'admission Bachelor CDSP", "Document réservé au jury · Barème : QCM /8 · Rédaction /6 · Calcul /6");
  _fp2DocTitre_(c, "Partie 1 — QCM : réponses (0,4 point par bonne réponse)");
  _fp2DocTable_(c, [["Question", "Réponse", "Question", "Réponse"]].concat([0, 1, 2, 3, 4, 5, 6, 7, 8, 9].map(function (i) {
    return [String(i + 1), L[FP2_QCM[i][2]] + ") " + FP2_QCM[i][1][FP2_QCM[i][2]], String(i + 11), L[FP2_QCM[i + 10][2]] + ") " + FP2_QCM[i + 10][1][FP2_QCM[i + 10][2]]]; })));
  _fp2DocTitre_(c, "Partie 2 — Rédaction : grille de notation (/6)");
  _fp2DocTable_(c, [["Critère", "Attendu", "Points"],
    ["Structure", "Objet clair, accusé de réception, corps organisé, formule de politesse", "1"],
    ["Démarche", "Vérification des relevés d'heures et de la validation par le manager ; contrôle du bulletin", "1,5"],
    ["Pièces demandées", "Relevés d'heures, planning ou tout justificatif des 6 heures", "1"],
    ["Régularisation", "Rappel de salaire sur le bulletin suivant (ou acompte), délai annoncé", "1"],
    ["Ton et confidentialité", "Courtois, factuel, rassurant ; pas de mise en cause ; données personnelles protégées", "1"],
    ["Langue", "Orthographe et syntaxe correctes", "0,5"]]);
  _fp2DocTitre_(c, "Partie 3 — Calcul : corrigé (/6)");
  _fp2DocTable_(c, [["Question", "Calcul", "Résultat", "Points"],
    ["1. Taux horaire", "2 000 / 151,67", "13,19 €", "0,5"],
    ["2. Heures supplémentaires", "6 × 13,19 × 1,25", "98,93 €", "1"],
    ["3. Salaire brut", "2 000 + 98,93 + 150", "2 248,93 €", "1"],
    ["4. Cotisations salariales / net avant impôt", "2 248,93 × 22 % = 494,76 ; 2 248,93 − 494,76", "494,76 € / 1 754,17 €", "1"],
    ["5. Prélèvement à la source / net payé", "1 754,17 × 4,5 % = 78,94 ; 1 754,17 − 78,94", "78,94 € / 1 675,23 €", "1"],
    ["6. Coût employeur", "2 248,93 × 42 % = 944,55 ; 2 248,93 + 944,55", "3 193,48 €", "1"],
    ["7. Congés payés acquis", "7 × 2,5 = 17,5 → arrondi au nombre entier supérieur", "18 jours ouvrables", "0,5"]]);
  _fp2DocTexte_(c, "Note du test (/20) à reporter dans le premier critère de l'onglet « Critères d'admission Bachelor » de la fiche de positionnement.", true);
  dc.saveAndClose();

  [ds.getId(), dc.getId()].forEach(function (id) { try { DriveApp.getFileById(id).moveTo(parent); } catch (e) {} });
  var T = { sujet: ds.getUrl(), corrige: dc.getUrl(), dossier: parent.getUrl(), le: Utilities.formatDate(new Date(), "Europe/Paris", "dd/MM/yyyy HH:mm") };
  props.setProperty("FP2_BACH_TESTS", JSON.stringify(T));
  return T;
}
