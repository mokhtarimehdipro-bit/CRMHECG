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

// Google Doc créé depuis du HTML par l'API Drive (conversion) : pas besoin de l'autorisation « documents »
function _fp2DocHtml_(nom, html, dossierId) {
  var bound = "fp2" + Date.now(), meta = { name: nom, mimeType: "application/vnd.google-apps.document", parents: [dossierId] };
  var corps = "--" + bound + "\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n" + JSON.stringify(meta) + "\r\n--" + bound + "\r\nContent-Type: text/html; charset=UTF-8\r\n\r\n" + html + "\r\n--" + bound + "--";
  var r = UrlFetchApp.fetch("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true&fields=id,webViewLink", {
    method: "post", contentType: "multipart/related; boundary=" + bound, payload: Utilities.newBlob(corps).getBytes(), muteHttpExceptions: true,
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() } });
  if (r.getResponseCode() !== 200) throw new Error("Création du document impossible : " + r.getContentText().slice(0, 200));
  var j = JSON.parse(r.getContentText()); return { id: j.id, url: j.webViewLink };
}
function _fp2E_(s) { return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); }
var FP2_CSS = "body{font-family:Arial,sans-serif;color:#0e1a2b;font-size:11pt}h1{color:#1A4E8A;font-size:22pt;margin:0}h2{color:#1A4E8A;font-size:14pt;margin-top:18pt;border-bottom:2px solid #F26522;padding-bottom:3pt}"
  + ".k{color:#F26522;font-weight:bold;font-size:9pt;letter-spacing:1px}.s{color:#3b4859;font-size:10pt}table{border-collapse:collapse;width:100%}th{background:#1A4E8A;color:#ffffff;font-size:9.5pt;padding:5pt;border:1px solid #1A4E8A;text-align:left}"
  + "td{border:1px solid #cdd8e6;padding:5pt;font-size:10pt;vertical-align:top}.q{font-weight:bold;margin:9pt 0 2pt}.r{margin:0 0 0 18pt}.l{color:#c3cbd6}";
function _fp2Entete_(titre, sous) { return "<p class='k'>HECG — ÉCOLE DE COMMERCE ET DE GESTION</p><h1>" + _fp2E_(titre) + "</h1><p class='s'>" + _fp2E_(sous) + "</p><hr>"; }
function _fp2Table_(rows) { return "<table><tr>" + rows[0].map(function (h) { return "<th>" + _fp2E_(h) + "</th>"; }).join("") + "</tr>" + rows.slice(1).map(function (r) { return "<tr>" + r.map(function (c) { return "<td>" + _fp2E_(c) + "</td>"; }).join("") + "</tr>"; }).join("") + "</table>"; }

function _fp2TestsBachelor_() {
  var props = PropertiesService.getScriptProperties(), L = "ABCD";
  var parent; try { var it = DriveApp.getFoldersByName("Bachelor CDSP — Tests d'admission"); parent = it.hasNext() ? it.next() : DriveApp.createFolder("Bachelor CDSP — Tests d'admission"); } catch (e) { parent = DriveApp.getRootFolder(); }
  var calc = [["Élément", "Valeur"], ["Salaire de base mensuel brut", "2 000,00 €"], ["Heures supplémentaires effectuées", "6 heures, majorées à 25 %"], ["Prime d'objectif", "150,00 €"],
    ["Taux global (simplifié) des cotisations salariales", "22 % du brut"], ["Taux global (simplifié) des cotisations patronales", "42 % du brut"],
    ["Taux de prélèvement à la source", "4,5 % (appliqué au net avant impôt, simplification)"], ["Ancienneté au 31 mars", "7 mois de travail effectif"]];
  // ---------- SUJET ----------
  var h = "<html><head><meta charset='utf-8'><style>" + FP2_CSS + "</style></head><body>"
    + _fp2Entete_("Test d'admission — Bachelor CDSP", "Ressources humaines et paie · Durée : 1 h 30 · Calculatrice autorisée · Aucun document")
    + _fp2Table_([["Nom", "Prénom", "Date", "Note /20"], ["", "", "", ""]])
    + "<p>Le test comporte trois parties : un QCM (8 points), un exercice rédactionnel (6 points) et un exercice de calcul de paie (6 points).</p>"
    + "<h2>Partie 1 — QCM (8 points)</h2><p class='s'>Pour chaque question, cochez une seule réponse. Bonne réponse : 0,4 point ; pas de point négatif.</p>"
    + FP2_QCM.map(function (q, i) { return "<p class='q'>" + (i + 1) + ". " + _fp2E_(q[0]) + "</p>" + q[1].map(function (r, k) { return "<p class='r'>☐ " + L[k] + ") " + _fp2E_(r) + "</p>"; }).join(""); }).join("")
    + "<h2 style='page-break-before:always'>Partie 2 — Exercice rédactionnel (6 points)</h2>"
    + "<p><b>Situation :</b> vous êtes assistant(e) paie dans un cabinet. M. Martin, salarié d'une entreprise cliente, vous écrit qu'il a effectué 6 heures supplémentaires en mars et qu'elles n'apparaissent pas sur son bulletin de paie. Il est mécontent et demande une régularisation immédiate.</p>"
    + "<p>Rédigez la réponse par courriel (15 à 20 lignes) : accusé de réception, vérifications que vous allez mener, pièces que vous lui demandez, délai et modalités de régularisation, formule de politesse. Soignez le ton, la clarté et la confidentialité.</p>"
    + new Array(19).join("<p class='l'>____________________________________________________________________________________</p>")
    + "<h2 style='page-break-before:always'>Partie 3 — Exercice de calcul de paie (6 points)</h2>"
    + "<p>Mme Durand est employée à temps plein (35 h, soit 151,67 h par mois). Données de mars :</p>" + _fp2Table_(calc)
    + "<p class='q'>Questions (arrondir au centime) :</p>"
    + ["1. Calculez le taux horaire brut (au centime).", "2. Calculez le montant des heures supplémentaires.", "3. Calculez le salaire brut du mois.", "4. Calculez les cotisations salariales et le net avant impôt.",
       "5. Calculez le prélèvement à la source et le net payé.", "6. Calculez le coût total employeur.", "7. Combien de jours ouvrables de congés payés Mme Durand a-t-elle acquis à fin mars ?"].map(function (q) { return "<p>" + _fp2E_(q) + "</p>"; }).join("")
    + "</body></html>";
  var ds = _fp2DocHtml_("Bachelor CDSP — Test d'admission — Sujet", h, parent.getId());
  // ---------- CORRIGÉ ----------
  var c = "<html><head><meta charset='utf-8'><style>" + FP2_CSS + "</style></head><body>"
    + _fp2Entete_("Corrigé — Test d'admission Bachelor CDSP", "Document réservé au jury · Barème : QCM /8 · Rédaction /6 · Calcul /6")
    + "<h2>Partie 1 — QCM : réponses (0,4 point par bonne réponse)</h2>"
    + _fp2Table_([["Question", "Réponse", "Question", "Réponse"]].concat([0, 1, 2, 3, 4, 5, 6, 7, 8, 9].map(function (i) {
        return [String(i + 1), L[FP2_QCM[i][2]] + ") " + FP2_QCM[i][1][FP2_QCM[i][2]], String(i + 11), L[FP2_QCM[i + 10][2]] + ") " + FP2_QCM[i + 10][1][FP2_QCM[i + 10][2]]]; })))
    + "<h2>Partie 2 — Rédaction : grille de notation (/6)</h2>"
    + _fp2Table_([["Critère", "Attendu", "Points"],
        ["Structure", "Objet clair, accusé de réception, corps organisé, formule de politesse", "1"],
        ["Démarche", "Vérification des relevés d'heures et de la validation par le manager ; contrôle du bulletin", "1,5"],
        ["Pièces demandées", "Relevés d'heures, planning ou tout justificatif des 6 heures", "1"],
        ["Régularisation", "Rappel de salaire sur le bulletin suivant (ou acompte), délai annoncé", "1"],
        ["Ton et confidentialité", "Courtois, factuel, rassurant ; pas de mise en cause ; données personnelles protégées", "1"],
        ["Langue", "Orthographe et syntaxe correctes", "0,5"]])
    + "<h2>Partie 3 — Calcul : corrigé (/6)</h2>"
    + _fp2Table_([["Question", "Calcul", "Résultat", "Points"],
        ["1. Taux horaire", "2 000 / 151,67", "13,19 €", "0,5"],
        ["2. Heures supplémentaires", "6 × 13,19 × 1,25", "98,93 €", "1"],
        ["3. Salaire brut", "2 000 + 98,93 + 150", "2 248,93 €", "1"],
        ["4. Cotisations salariales / net avant impôt", "2 248,93 × 22 % = 494,76 ; 2 248,93 − 494,76", "494,76 € / 1 754,17 €", "1"],
        ["5. Prélèvement à la source / net payé", "1 754,17 × 4,5 % = 78,94 ; 1 754,17 − 78,94", "78,94 € / 1 675,23 €", "1"],
        ["6. Coût employeur", "2 248,93 × 42 % = 944,55 ; 2 248,93 + 944,55", "3 193,48 €", "1"],
        ["7. Congés payés acquis", "7 × 2,5 = 17,5 → arrondi au nombre entier supérieur", "18 jours ouvrables", "0,5"]])
    + "<p><b>Note du test (/20)</b> à reporter dans le premier critère de l'onglet « Critères d'admission Bachelor » de la fiche de positionnement.</p></body></html>";
  var dc = _fp2DocHtml_("Bachelor CDSP — Test d'admission — Corrigé (jury)", c, parent.getId());
  var T = { sujet: ds.url, corrige: dc.url, dossier: parent.getUrl(), le: Utilities.formatDate(new Date(), "Europe/Paris", "dd/MM/yyyy HH:mm") };
  props.setProperty("FP2_BACH_TESTS", JSON.stringify(T));
  return T;
}
