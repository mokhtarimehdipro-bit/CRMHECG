// =============================================================================
//  VIE SCOLAIRE — TICKETS DE MAINTENANCE & CENTRE DES FORMULAIRES (demande 2026-11-09)
//
//  1. Tickets de maintenance : les réponses du formulaire Google (classeur MAINT_SHEET_ID) sont
//     lues telles quelles ; le suivi (statut, responsable, commentaire, résolution) est conservé
//     dans la feuille MAINTENANCE_SUIVI du CRM — le classeur des réponses n'est jamais modifié.
//  2. Formulaires : création dans l'application (feuille FORMULAIRES), réponses publiques via
//     formulaire.html (feuille FORMULAIRES_REPONSES), et formulaires Google existants
//     référencés (type « externe ») pour tout centraliser au même endroit.
// =============================================================================

var MAINT_SHEET_ID = "1lcRRUQ84nC-SVjPfe2sKTbU6KNBZRvTrMh28eYCo_fs";
var SATISF_ENT_SHEET_ID = "1vfzvdf5S8-ZyQiAa5QRaOsSL4vpsMfvm4Z1nhNMv8kg";
var FORM_PAGE_URL = "https://mokhtarimehdipro-bit.github.io/CRMHECG/formulaire.html";
var F_ENTETES = ["ID", "Titre", "Description", "Statut", "Type", "Questions_JSON", "Theme_JSON", "URL_externe", "Sheet_externe", "Créé par", "Créé le", "Modifié le", "Indicateur Qualiopi", "Message de fin"];
var F_TYPES = { texte: 1, paragraphe: 1, choix: 1, cases: 1, liste: 1, echelle: 1, etoiles: 1, nps: 1, oui_non: 1, date: 1, email: 1, tel: 1, nombre: 1, section: 1 };

function _fRole_(p) { return String(p.vieScRole || "").trim(); }
function _fPeutGerer_(p) { var r = _fRole_(p); return r === "Super-admin" || r === "Pédagogie" || r === "Administratif" || String(p.emailAuteur || "").toLowerCase() === "m.mokhtari@hecg.fr"; }
function _fDate_() { return Utilities.formatDate(new Date(), "Europe/Paris", "yyyy-MM-dd HH:mm"); }

// ── 1. TICKETS DE MAINTENANCE ────────────────────────────────────────────────
function _maintSuiviFeuille_(ss) {
  var ws = ss.getSheetByName("MAINTENANCE_SUIVI");
  if (!ws) {
    ws = ss.insertSheet("MAINTENANCE_SUIVI");
    ws.getRange(1, 1, 1, 8).setValues([["Clé ticket", "Statut", "Responsable", "Commentaire", "Priorité", "Résolu le", "Mis à jour le", "Par"]]).setFontWeight("bold");
    ws.setFrozenRows(1);
  }
  return ws;
}
// Repère les colonnes utiles d'après l'intitulé des questions du formulaire
function _maintColonnes_(h) {
  var n = h.map(function (x) { return String(x || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase(); });
  function f(re, sauf) { for (var i = 0; i < n.length; i++) if (re.test(n[i]) && !(sauf && sauf.test(n[i]))) return i; return -1; }
  return {
    date: f(/horodat|timestamp|date/), email: f(/e-?mail|courriel/), nom: f(/\bnom\b|prenom|demandeur|qui /),
    campus: f(/campus|site|etablissement/), lieu: f(/salle|lieu|localisation|emplacement|ou se/), categorie: f(/categorie|type|nature|domaine|objet/),
    urgence: f(/urgen|priorit|gravit|criticit/), description: f(/descri|probleme|detail|precis|commentaire|message/, /photo/), photo: f(/photo|image|piece|fichier|capture/)
  };
}
function handleGetTicketsMaintenance(p, ss) {
  if (!_fPeutGerer_(p)) return createJsonResponse({ success: false, message: "Accès réservé à l'équipe Vie scolaire." });
  var ext;
  try { ext = SpreadsheetApp.openById(MAINT_SHEET_ID); }
  catch (e) {
    return createJsonResponse({ success: false, acces: false, message: "Le CRM n'a pas accès au classeur des tickets. Partagez-le (lecture) avec m.mokhtari@hecg.fr puis rechargez l'onglet.", lien: "https://docs.google.com/spreadsheets/d/" + MAINT_SHEET_ID + "/edit" });
  }
  try {
    var ws = ext.getSheets()[0], data = ws.getDataRange().getDisplayValues();
    var h = data[0] || [], C = _maintColonnes_(h), formUrl = "";
    try { formUrl = ws.getFormUrl() || ext.getFormUrl() || ""; } catch (eF) {}
    var suivi = {}, dS = _maintSuiviFeuille_(ss).getDataRange().getValues();
    for (var k = 1; k < dS.length; k++) if (dS[k][0]) suivi[String(dS[k][0])] = { statut: String(dS[k][1] || ""), responsable: String(dS[k][2] || ""), commentaire: String(dS[k][3] || ""),
      priorite: String(dS[k][4] || ""), resolu: dS[k][5] ? Utilities.formatDate(new Date(dS[k][5]), "Europe/Paris", "dd/MM/yyyy") : "", maj: String(dS[k][6] || ""), par: String(dS[k][7] || "") };
    var tickets = [];
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      if (!r.join("").trim()) continue;
      var cle = "L" + (i + 1) + "|" + String(C.date >= 0 ? r[C.date] : "").trim();
      var v = function (c) { return c >= 0 ? String(r[c] || "").trim() : ""; };
      var s = suivi[cle] || {};
      tickets.push({ cle: cle, ligne: i + 1, date: v(C.date), email: v(C.email), nom: v(C.nom), campus: v(C.campus), lieu: v(C.lieu), categorie: v(C.categorie),
        urgence: v(C.urgence), description: v(C.description), photo: v(C.photo), champs: h.map(function (t, j) { return [String(t), String(r[j] || "")]; }).filter(function (x) { return x[1]; }),
        statut: s.statut || "Nouveau", responsable: s.responsable || "", commentaire: s.commentaire || "", priorite: s.priorite || "", resolu: s.resolu || "", maj: s.maj || "", par: s.par || "" });
    }
    tickets.reverse();   // plus récents d'abord
    return createJsonResponse({ success: true, tickets: tickets, formUrl: formUrl, sheetUrl: ext.getUrl(), titre: ext.getName(), colonnes: h });
  } catch (e2) { return createJsonResponse({ success: false, message: e2.toString() }); }
}
function handleUpdateTicketMaintenance(p, ss) {
  if (!_fPeutGerer_(p)) return createJsonResponse({ success: false, message: "Accès réservé à l'équipe Vie scolaire." });
  try {
    var ws = _maintSuiviFeuille_(ss), data = ws.getDataRange().getValues(), cle = String(p.cle || "");
    if (!cle) return createJsonResponse({ success: false, message: "Ticket inconnu." });
    var statut = String(p.statut || "Nouveau"), par = p.nomAuteur || p.emailAuteur || "";
    var ligne = [cle, statut, String(p.responsable || ""), String(p.commentaire || "").slice(0, 2000), String(p.priorite || ""),
      statut === "Résolu" ? new Date() : "", _fDate_(), par];
    var avant = "";
    for (var i = 1; i < data.length; i++) if (String(data[i][0]) === cle) {
      avant = String(data[i][1] || "");
      if (statut === "Résolu" && data[i][5]) ligne[5] = data[i][5];
      ws.getRange(i + 1, 1, 1, 8).setValues([ligne]);
      break;
    }
    if (i >= data.length) ws.appendRow(ligne);
    logAction(par, p.vieScRole || "", "Modification", "Vie scolaire — Maintenance", cle,
      "Ticket « " + String(p.resume || "").slice(0, 80) + " » : " + (avant || "Nouveau") + " → " + statut + (p.responsable ? " (responsable : " + p.responsable + ")" : ""));
    return createJsonResponse({ success: true });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// ── 2. FORMULAIRES ───────────────────────────────────────────────────────────
function _fFeuille_(ss) {
  var ws = ss.getSheetByName("FORMULAIRES");
  if (!ws) {
    ws = ss.insertSheet("FORMULAIRES");
    ws.getRange(1, 1, 1, F_ENTETES.length).setValues([F_ENTETES]).setFontWeight("bold");
    ws.setFrozenRows(1);
    var d = _fDate_();
    ws.getRange(2, 1, 3, F_ENTETES.length).setValues([
      ["F-SATISF-ENT", "Questionnaire de satisfaction entreprise", "Formulaire Google existant (indicateur 30) : appréciation des entreprises d'accueil.", "publié", "externe", "[]",
        JSON.stringify({ couleur: "#5B3F86", accent: "#F26522" }), "https://docs.google.com/forms/d/1RBV2Gfg6OxkA8zf3mqLtlnF41NtqXm_LWSUx_B-FM_M/viewform", SATISF_ENT_SHEET_ID, "Import", d, d, "30", ""],
      ["F-MAINT", "Tickets de maintenance", "Formulaire Google de signalement des problèmes techniques et matériels (suivi dans l'onglet Maintenance).", "publié", "externe", "[]",
        JSON.stringify({ couleur: "#0F766E", accent: "#F59E0B" }), "", MAINT_SHEET_ID, "Import", d, d, "31", ""],
      ["F-IND33", "Évaluation des contenus et des enseignements", "Votre avis sur les cours de cette période : il est anonyme et sert à améliorer les enseignements (indicateur Qualiopi 33).", "brouillon", "interne",
        JSON.stringify(F_MODELE_IND33), JSON.stringify({ couleur: "#1A4E8A", accent: "#F26522", fond: "degrade" }), "", "", "Modèle", d, d, "33",
        "Merci ! Vos réponses sont analysées avec l'équipe pédagogique et donnent lieu à des actions d'amélioration."]
    ]);
  }
  return ws;
}
var F_MODELE_IND33 = [
  { id: "q1", type: "liste", label: "Votre formation", obligatoire: true, options: ["BTS CG", "BTS GPME", "DCG", "DSCG", "Bachelor CDSP"] },
  { id: "q2", type: "texte", label: "Matière / module évalué", obligatoire: true, aide: "Ex. : UE 9 Comptabilité, CEJM, Droit social…" },
  { id: "s1", type: "section", label: "Les contenus" },
  { id: "q3", type: "etoiles", label: "Clarté des objectifs du module", obligatoire: true },
  { id: "q4", type: "etoiles", label: "Adéquation des contenus avec l'examen / la certification", obligatoire: true },
  { id: "q5", type: "etoiles", label: "Utilité pour vos missions en entreprise", obligatoire: true },
  { id: "s2", type: "section", label: "Les enseignements" },
  { id: "q6", type: "echelle", label: "Rythme de progression", min: 1, max: 5, minLabel: "Trop lent", maxLabel: "Trop rapide", obligatoire: true },
  { id: "q7", type: "etoiles", label: "Qualité des supports et ressources", obligatoire: true },
  { id: "q8", type: "etoiles", label: "Qualité des évaluations (DST, examens blancs, corrections)", obligatoire: true },
  { id: "q9", type: "nps", label: "Recommanderiez-vous ce module à un autre apprenant ?", obligatoire: false },
  { id: "q10", type: "paragraphe", label: "Ce qui fonctionne bien", obligatoire: false },
  { id: "q11", type: "paragraphe", label: "Ce que vous proposez d'améliorer", obligatoire: false }
];
function _fReponsesFeuille_(ss) {
  var ws = ss.getSheetByName("FORMULAIRES_REPONSES");
  if (!ws) { ws = ss.insertSheet("FORMULAIRES_REPONSES"); ws.getRange(1, 1, 1, 5).setValues([["Formulaire", "Horodatage", "Réponses_JSON", "Durée (s)", "Réf"]]).setFontWeight("bold"); ws.setFrozenRows(1); }
  return ws;
}
function _fLire_(ss) {
  var data = _fFeuille_(ss).getDataRange().getValues(), out = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[0]) continue;
    var q = [], th = {};
    try { q = JSON.parse(r[5] || "[]"); } catch (e) {}
    try { th = JSON.parse(r[6] || "{}"); } catch (e2) {}
    out.push({ ligne: i + 1, id: String(r[0]), titre: String(r[1] || ""), description: String(r[2] || ""), statut: String(r[3] || "brouillon"), type: String(r[4] || "interne"),
      questions: q, theme: th, urlExterne: String(r[7] || ""), sheetExterne: String(r[8] || ""), creePar: String(r[9] || ""), creeLe: String(r[10] || ""), modifieLe: String(r[11] || ""),
      indicateur: String(r[12] || ""), messageFin: String(r[13] || "") });
  }
  return out;
}
function handleGetFormulaires(p, ss) {
  if (!_fPeutGerer_(p)) return createJsonResponse({ success: false, message: "Accès réservé à l'équipe Vie scolaire." });
  try {
    var forms = _fLire_(ss).filter(function (f) { return f.statut !== "archivé"; });
    var nb = {}, dR = _fReponsesFeuille_(ss).getDataRange().getValues(), der = {};
    for (var i = 1; i < dR.length; i++) { var k = String(dR[i][0]); nb[k] = (nb[k] || 0) + 1; der[k] = dR[i][1]; }
    forms.forEach(function (f) {
      f.lienPublic = f.type === "interne" ? FORM_PAGE_URL + "?f=" + encodeURIComponent(f.id) : f.urlExterne;
      if (f.type === "interne") { f.nbReponses = nb[f.id] || 0; f.derniere = der[f.id] ? Utilities.formatDate(new Date(der[f.id]), "Europe/Paris", "dd/MM/yyyy HH:mm") : ""; }
      else if (f.sheetExterne) {
        try { var w = SpreadsheetApp.openById(f.sheetExterne).getSheets()[0]; f.nbReponses = Math.max(0, w.getLastRow() - 1); f.sheetUrl = "https://docs.google.com/spreadsheets/d/" + f.sheetExterne + "/edit";
          if (!f.urlExterne) { try { f.urlExterne = f.lienPublic = w.getFormUrl() || ""; } catch (eU) {} } }
        catch (eS) { f.nbReponses = null; f.accesRefuse = true; }
      }
    });
    return createJsonResponse({ success: true, data: forms, pageUrl: FORM_PAGE_URL });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}
function _fNettoyerQuestions_(qs) {
  if (!Array.isArray(qs)) return [];
  return qs.slice(0, 80).map(function (q, i) {
    var t = F_TYPES[q.type] ? q.type : "texte";
    var o = { id: String(q.id || ("q" + (i + 1))).replace(/[^\w-]/g, "").slice(0, 20) || ("q" + (i + 1)), type: t, label: String(q.label || "").slice(0, 300),
      aide: String(q.aide || "").slice(0, 400), obligatoire: !!q.obligatoire && t !== "section" };
    if (t === "choix" || t === "cases" || t === "liste") o.options = (q.options || []).map(function (x) { return String(x).slice(0, 200); }).filter(Boolean).slice(0, 40);
    if (t === "echelle") { o.min = Math.max(0, Math.min(1, Number(q.min) || 1)); o.max = Math.max(3, Math.min(10, Number(q.max) || 5)); o.minLabel = String(q.minLabel || "").slice(0, 60); o.maxLabel = String(q.maxLabel || "").slice(0, 60); }
    if (t === "cases" && q.autre) o.autre = true;
    if (t === "choix" && q.autre) o.autre = true;
    return o;
  });
}
function handleSaveFormulaire(p, ss) {
  if (!_fPeutGerer_(p)) return createJsonResponse({ success: false, message: "Accès réservé à l'équipe Vie scolaire." });
  try {
    var f = typeof p.formulaire === "string" ? JSON.parse(p.formulaire) : (p.formulaire || {});
    var titre = String(f.titre || "").trim().slice(0, 200);
    if (!titre) return createJsonResponse({ success: false, message: "Donnez un titre au formulaire." });
    var type = f.type === "externe" ? "externe" : "interne";
    var statut = { "brouillon": 1, "publié": 1, "fermé": 1 }[f.statut] ? f.statut : "brouillon";
    var qs = _fNettoyerQuestions_(f.questions);
    if (type === "interne" && statut === "publié" && !qs.filter(function (q) { return q.type !== "section"; }).length) return createJsonResponse({ success: false, message: "Ajoutez au moins une question avant de publier." });
    var sheetExt = String(f.sheetExterne || "").match(/[-\w]{25,}/);
    var th = f.theme || {};
    var theme = { couleur: /^#[0-9a-f]{6}$/i.test(th.couleur) ? th.couleur : "#1A4E8A", accent: /^#[0-9a-f]{6}$/i.test(th.accent) ? th.accent : "#F26522",
      fond: { degrade: 1, clair: 1, sombre: 1, aurore: 1 }[th.fond] ? th.fond : "degrade", image: /^https:\/\//.test(th.image || "") ? String(th.image).slice(0, 500) : "", police: { moderne: 1, elegante: 1, ronde: 1 }[th.police] ? th.police : "moderne" };
    var ws = _fFeuille_(ss), forms = _fLire_(ss), d = _fDate_(), par = p.nomAuteur || p.emailAuteur || "";
    var exist = f.id ? forms.filter(function (x) { return x.id === f.id; })[0] : null;
    var id = exist ? exist.id : "F-" + Date.now().toString(36).toUpperCase();
    var ligne = [id, titre, String(f.description || "").slice(0, 2000), statut, type, JSON.stringify(qs), JSON.stringify(theme),
      String(f.urlExterne || "").slice(0, 500), sheetExt ? sheetExt[0] : "", exist ? exist.creePar : par, exist ? exist.creeLe : d, d,
      String(f.indicateur || "").slice(0, 10), String(f.messageFin || "").slice(0, 600)];
    if (exist) ws.getRange(exist.ligne, 1, 1, F_ENTETES.length).setValues([ligne]); else ws.appendRow(ligne);
    logAction(par, p.vieScRole || "", exist ? "Modification" : "Création", "Vie scolaire — Formulaires", id,
      (exist ? "Formulaire modifié : " : "Formulaire créé : ") + titre + " (" + type + ", " + statut + ", " + qs.length + " question(s))");
    return createJsonResponse({ success: true, id: id, lienPublic: type === "interne" ? FORM_PAGE_URL + "?f=" + encodeURIComponent(id) : ligne[7] });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}
function handleDeleteFormulaire(p, ss) {
  if (!_fPeutGerer_(p)) return createJsonResponse({ success: false, message: "Accès réservé à l'équipe Vie scolaire." });
  var f = _fLire_(ss).filter(function (x) { return x.id === String(p.id); })[0];
  if (!f) return createJsonResponse({ success: false, message: "Formulaire introuvable." });
  _fFeuille_(ss).getRange(f.ligne, 4).setValue("archivé");   // les réponses sont conservées
  logAction(p.nomAuteur || p.emailAuteur || "", p.vieScRole || "", "Suppression", "Vie scolaire — Formulaires", f.id, "Formulaire archivé : " + f.titre);
  return createJsonResponse({ success: true });
}
function handleGetFormulaireReponses(p, ss) {
  if (!_fPeutGerer_(p)) return createJsonResponse({ success: false, message: "Accès réservé à l'équipe Vie scolaire." });
  try {
    var f = _fLire_(ss).filter(function (x) { return x.id === String(p.id); })[0];
    if (!f) return createJsonResponse({ success: false, message: "Formulaire introuvable." });
    if (f.type === "externe") {
      if (!f.sheetExterne) return createJsonResponse({ success: true, externe: true, colonnes: [], lignes: [] });
      var w;
      try { w = SpreadsheetApp.openById(f.sheetExterne).getSheets()[0]; }
      catch (eA) { return createJsonResponse({ success: false, message: "Le CRM n'a pas accès au classeur des réponses : partagez-le avec m.mokhtari@hecg.fr." }); }
      var d = w.getDataRange().getDisplayValues();
      return createJsonResponse({ success: true, externe: true, colonnes: d[0] || [], lignes: d.slice(1).filter(function (r) { return r.join("").trim(); }) });
    }
    var dR = _fReponsesFeuille_(ss).getDataRange().getValues(), out = [];
    for (var i = 1; i < dR.length; i++) if (String(dR[i][0]) === f.id) {
      var rep = {};
      try { rep = JSON.parse(dR[i][2] || "{}"); } catch (e) {}
      out.push({ date: Utilities.formatDate(new Date(dR[i][1]), "Europe/Paris", "dd/MM/yyyy HH:mm"), duree: dR[i][3], reponses: rep });
    }
    return createJsonResponse({ success: true, externe: false, questions: f.questions, reponses: out });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// ── Accès public (formulaire.html, sans connexion) ───────────────────────────
function handleFormPublicGet(params, ss) {
  var f = _fLire_(ss).filter(function (x) { return x.id === String(params.id || ""); })[0];
  if (!f || f.type !== "interne" || f.statut === "archivé" || f.statut === "brouillon") return createJsonResponse({ success: false, message: "Ce formulaire n'existe pas ou n'est pas encore publié." });
  if (f.statut === "fermé") return createJsonResponse({ success: false, ferme: true, titre: f.titre, theme: f.theme, message: "Ce formulaire est clôturé. Merci de votre intérêt !" });
  return createJsonResponse({ success: true, data: { id: f.id, titre: f.titre, description: f.description, questions: f.questions, theme: f.theme, messageFin: f.messageFin } });
}
function handleFormPublicSubmit(params, ss) {
  try {
    if (params.hp) return createJsonResponse({ success: true });            // pot de miel : robot ignoré silencieusement
    var f = _fLire_(ss).filter(function (x) { return x.id === String(params.id || ""); })[0];
    if (!f || f.type !== "interne" || f.statut !== "publié") return createJsonResponse({ success: false, message: "Ce formulaire n'accepte plus de réponses." });
    var duree = Math.round(Number(params.duree) || 0);
    if (duree < 3) return createJsonResponse({ success: false, message: "Réponse trop rapide, merci de réessayer." });
    var cache = CacheService.getScriptCache(), cle = "fsub_" + f.id, n = Number(cache.get(cle) || 0);
    if (n > 60) return createJsonResponse({ success: false, message: "Trop de réponses en peu de temps, réessayez dans une minute." });
    cache.put(cle, String(n + 1), 60);
    var brut = typeof params.reponses === "string" ? JSON.parse(params.reponses) : (params.reponses || {}), propre = {};
    var manquantes = [];
    f.questions.forEach(function (q) {
      if (q.type === "section") return;
      var v = brut[q.id];
      if (Array.isArray(v)) v = v.map(function (x) { return String(x).slice(0, 300); }).slice(0, 40);
      else if (v !== undefined && v !== null) v = String(v).slice(0, 5000);
      var vide = v === undefined || v === null || v === "" || (Array.isArray(v) && !v.length);
      if (q.obligatoire && vide) manquantes.push(q.label);
      if (!vide) {
        if (q.type === "email" && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(v)) manquantes.push(q.label + " (e-mail invalide)");
        if ((q.type === "etoiles" || q.type === "echelle" || q.type === "nps" || q.type === "nombre") && isNaN(Number(v))) manquantes.push(q.label);
        propre[q.id] = v;
      }
    });
    if (manquantes.length) return createJsonResponse({ success: false, message: "Merci de compléter : " + manquantes.slice(0, 5).join(", ") });
    _fReponsesFeuille_(ss).appendRow([f.id, new Date(), JSON.stringify(propre), duree, "R-" + Date.now().toString(36)]);
    return createJsonResponse({ success: true, messageFin: f.messageFin || "Merci, votre réponse a bien été enregistrée !" });
  } catch (e) { return createJsonResponse({ success: false, message: "Envoi impossible : " + e }); }
}
