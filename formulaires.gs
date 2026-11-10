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
var MAINT_ENTETES = ["Clé ticket", "Statut", "Responsable", "Commentaire", "Priorité", "Résolu le", "Mis à jour le", "Par",
  "Archivé", "Journal_JSON", "Corrections_JSON", "Relances", "Dernière relance", "Responsable — mail", "N° ticket"];
function _maintNumLib_(n) { n = Number(n || 0); return n ? "T-" + ("000" + n).slice(-4) : ""; }
var MAINT_FERMES = ["Résolu", "Clôturé", "Annulé", "Sans suite"];
function _maintSuiviFeuille_(ss) {
  var ws = ss.getSheetByName("MAINTENANCE_SUIVI");
  if (!ws) { ws = ss.insertSheet("MAINTENANCE_SUIVI"); ws.setFrozenRows(1); }
  var h = ws.getRange(1, 1, 1, MAINT_ENTETES.length).getValues()[0];
  if (String(h[MAINT_ENTETES.length - 1] || "") !== MAINT_ENTETES[MAINT_ENTETES.length - 1]) ws.getRange(1, 1, 1, MAINT_ENTETES.length).setValues([MAINT_ENTETES]).setFontWeight("bold");
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
function _maintJson_(v, def) { try { var x = JSON.parse(String(v || "")); return x == null ? def : x; } catch (e) { return def; } }
// Équipe pouvant être désignée responsable (nom + mail) — sert à la relance par mail
function _maintEquipe_(ss) {
  var out = [];
  try { var d = getPersonnelDataCached_(ss); for (var i = 1; i < d.length; i++) { var nom = (String(d[i][2] || "") + " " + String(d[i][1] || "").toUpperCase()).trim(), mail = String(d[i][6] || "").trim(); if (nom) out.push({ nom: nom, mail: mail }); } } catch (e) {}
  return out;
}
// Contacts externes (prestataires, propriétaire, informaticien…) créés depuis les tickets
function _maintContactsFeuille_(ss) {
  var ws = ss.getSheetByName("MAINTENANCE_CONTACTS");
  if (!ws) { ws = ss.insertSheet("MAINTENANCE_CONTACTS"); ws.getRange(1, 1, 1, 6).setValues([["Nom", "Mail", "Téléphone", "Société / rôle", "Créé le", "Par"]]).setFontWeight("bold"); ws.setFrozenRows(1); }
  return ws;
}
function _maintContacts_(ss) {
  var d = _maintContactsFeuille_(ss).getDataRange().getValues(), out = [];
  for (var i = 1; i < d.length; i++) if (String(d[i][0] || "").trim()) out.push({ nom: String(d[i][0]).trim(), mail: String(d[i][1] || "").trim(), tel: String(d[i][2] || "").trim(), societe: String(d[i][3] || "").trim() });
  return out;
}
// Onglet des réponses : celui relié au formulaire (sinon celui qui a le plus de lignes) — pas forcément le 1er onglet
function _maintOngletReponses_(ext) {
  var tous = ext.getSheets(), meilleur = tous[0];
  for (var i = 0; i < tous.length; i++) { try { if (tous[i].getFormUrl()) return tous[i]; } catch (e) {} }
  tous.forEach(function (w) { if (w.getLastRow() > meilleur.getLastRow()) meilleur = w; });
  return meilleur;
}
// Clé stable d'un ticket : horodatage + mail (ne bouge pas si le Sheet est trié) ; l'ancienne clé « L<ligne>|date » reste reconnue
function _maintCle_(date, email, ligne) { return "T|" + String(date || "").trim() + "|" + String(email || "").trim().toLowerCase() + (date ? "" : "|L" + ligne); }
// Lecture des réponses du formulaire → lignes avec leur clé stable (ordre d'arrivée = ordre des lignes)
function _maintLireReponses_(ext) {
  var ws = _maintOngletReponses_(ext), data = ws.getDataRange().getDisplayValues();
  var h = data[0] || [], C = _maintColonnes_(h), lignes = [], vus = {};
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r.join("").trim()) continue;
    var d = C.date >= 0 ? String(r[C.date] || "").trim() : "", m = C.email >= 0 ? String(r[C.email] || "").trim() : "";
    var cle = _maintCle_(d, m, i + 1);
    if (vus[cle]) cle += "|L" + (i + 1);   // deux réponses à la même seconde par la même personne
    vus[cle] = 1;
    lignes.push({ i: i, r: r, cle: cle, ancienneCle: "L" + (i + 1) + "|" + d });
  }
  return { ws: ws, data: data, h: h, C: C, lignes: lignes };
}
// Attribue un numéro « T-0001 » à chaque nouvelle demande (ordre d'arrivée), une fois pour toutes.
// Une ligne de suivi « Nouveau » est créée pour chaque nouveau ticket. Renvoie le nombre de numéros attribués.
function _maintNumeroter_(ss, rep) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return 0;
  try {
    var ws = _maintSuiviFeuille_(ss), NC = MAINT_ENTETES.length, nL = Math.max(ws.getLastRow(), 1);
    var d = ws.getRange(1, 1, nL, NC).getValues(), index = {}, max = 0, modif = false, ajouts = [];
    for (var k = 1; k < d.length; k++) { if (d[k][0]) index[String(d[k][0])] = k; max = Math.max(max, Number(d[k][NC - 1] || 0)); }
    rep.lignes.forEach(function (l) {
      var k = index[l.cle];
      if (k == null && index[l.ancienneCle] != null) { k = index[l.ancienneCle]; d[k][0] = l.cle; index[l.cle] = k; modif = true; }   // migration ancienne clé
      if (k != null) { if (!Number(d[k][NC - 1] || 0)) { d[k][NC - 1] = ++max; modif = true; } return; }
      var n = [l.cle, "Nouveau", "", "", "", "", _fDate_(), "Numérotation auto", false, JSON.stringify([{ d: _fDate_(), par: "Système", evt: "Ticket reçu — n° " + _maintNumLib_(max + 1) }]), "{}", 0, "", "", ++max];
      ajouts.push(n); index[l.cle] = -1;
    });
    if (modif && d.length > 1) ws.getRange(1, 1, d.length, NC).setValues(d);
    if (ajouts.length) ws.getRange(ws.getLastRow() + 1, 1, ajouts.length, NC).setValues(ajouts);
    if (ajouts.length) logAction("Système", "", "Création", "Vie scolaire — Maintenance — numérotation", "MAINTENANCE_SUIVI", ajouts.length + " nouveau(x) ticket(s) numéroté(s) : " + ajouts.map(function (x) { return _maintNumLib_(x[NC - 1]); }).join(", "));
    return ajouts.length;
  } finally { lock.releaseLock(); }
}
// Passage automatique (auditAdressesTick) toutes les 15 min : les nouvelles demandes reçoivent leur numéro sans attendre l'ouverture de l'onglet
function _maintAuto_() {
  var props = PropertiesService.getScriptProperties(), der = Number(props.getProperty("MAINT_NUM_TS") || 0);
  if (Date.now() - der < 15 * 60 * 1000) return;
  props.setProperty("MAINT_NUM_TS", String(Date.now()));
  try {
    var ss = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ"), rep = _maintLireReponses_(SpreadsheetApp.openById(MAINT_SHEET_ID));
    _maintNumeroter_(ss, rep);
    _maintMiroir_(ss, rep);
  } catch (e) { console.warn("numérotation tickets : " + e); }
}
// Copie en lecture seule des réponses (feuille MAINTENANCE_COPIE du CRM) avec le n° de ticket — sert à l'analyse et aux contrôles
function _maintMiroir_(ss, rep) {
  var NC = MAINT_ENTETES.length, wsS = _maintSuiviFeuille_(ss), dS = wsS.getRange(1, 1, Math.max(wsS.getLastRow(), 1), NC).getValues(), num = {};
  for (var k = 1; k < dS.length; k++) if (dS[k][0]) num[String(dS[k][0])] = _maintNumLib_(dS[k][NC - 1]);
  var w = rep.h.length, rows = [["N° ticket", "Clé ticket"].concat(rep.h)];
  rep.lignes.forEach(function (l) { var r = l.r.slice(0, w); while (r.length < w) r.push(""); rows.push([num[l.cle] || "", l.cle].concat(r)); });
  var ws = ss.getSheetByName("MAINTENANCE_COPIE") || ss.insertSheet("MAINTENANCE_COPIE");
  ws.clearContents();
  ws.getRange(1, 1, rows.length, rows[0].length).setNumberFormat("@").setValues(rows);
  ws.getRange(1, 1, 1, rows[0].length).setFontWeight("bold"); ws.setFrozenRows(1);
}
// ── Tickets ouverts directement dans l'application (admin maître) : feuille MAINTENANCE_INTERNES ──
var MAINT_INT_ENTETES = ["ID", "Horodatage", "Mail demandeur", "Demandeur", "Campus", "Lieu", "Catégorie", "Urgence", "Description", "Créé par"];
function _maintMaitre_(p) { return String(p.vieScRole || "") === "Super-admin" || String(p.emailAuteur || "").toLowerCase() === "m.mokhtari@hecg.fr"; }
function _maintInternesFeuille_(ss) {
  var ws = ss.getSheetByName("MAINTENANCE_INTERNES");
  if (!ws) { ws = ss.insertSheet("MAINTENANCE_INTERNES"); ws.getRange(1, 1, 1, MAINT_INT_ENTETES.length).setValues([MAINT_INT_ENTETES]).setFontWeight("bold"); ws.setFrozenRows(1); }
  return ws;
}
function _maintInternes_(ss) {
  var ws = ss.getSheetByName("MAINTENANCE_INTERNES"); if (!ws || ws.getLastRow() < 2) return [];
  return ws.getRange(2, 1, ws.getLastRow() - 1, MAINT_INT_ENTETES.length).getDisplayValues().filter(function (r) { return r[0]; }).map(function (r) {
    return { cle: "I|" + r[0], date: r[1], email: r[2], nom: r[3], campus: r[4], lieu: r[5], categorie: r[6], urgence: r[7], description: r[8], par: r[9],
      champs: MAINT_INT_ENTETES.map(function (h, j) { return [h, r[j]]; }).filter(function (x) { return x[1] && x[0] !== "ID"; }) };
  });
}
// Lien public du formulaire (…/viewform), à partir de l'URL renvoyée par le Sheet
function _maintLienPublic_(u) { var m = String(u || "").match(/\/forms\/d\/(e\/)?([A-Za-z0-9_-]{20,})/); return m ? "https://docs.google.com/forms/d/" + (m[1] || "") + m[2] + "/viewform" : String(u || ""); }
function handleGetTicketsMaintenance(p, ss) {
  if (!_fPeutGerer_(p)) return createJsonResponse({ success: false, message: "Accès réservé à l'équipe Vie scolaire." });
  var ext;
  try { ext = SpreadsheetApp.openById(MAINT_SHEET_ID); }
  catch (e) {
    return createJsonResponse({ success: false, acces: false, message: "Le CRM n'a pas accès au classeur des tickets. Partagez-le (lecture) avec m.mokhtari@hecg.fr puis rechargez l'onglet.", lien: "https://docs.google.com/spreadsheets/d/" + MAINT_SHEET_ID + "/edit" });
  }
  try {
    var rep = _maintLireReponses_(ext), ws = rep.ws, h = rep.h, C = rep.C, formUrl = "";
    try { formUrl = ws.getFormUrl() || ext.getFormUrl() || ""; } catch (eF) {}
    var numerotes = 0; try { numerotes = _maintNumeroter_(ss, rep); } catch (eN) { console.warn(eN); }
    var NC = MAINT_ENTETES.length, wsS = _maintSuiviFeuille_(ss), dS = wsS.getRange(1, 1, Math.max(wsS.getLastRow(), 1), NC).getValues(), suivi = {};
    for (var k = 1; k < dS.length; k++) if (dS[k][0]) suivi[String(dS[k][0])] = { statut: String(dS[k][1] || ""), responsable: String(dS[k][2] || ""), commentaire: String(dS[k][3] || ""),
      priorite: String(dS[k][4] || ""), resolu: dS[k][5] ? Utilities.formatDate(new Date(dS[k][5]), "Europe/Paris", "dd/MM/yyyy") : "", maj: String(dS[k][6] || ""), par: String(dS[k][7] || ""),
      archive: dS[k][8] === true || String(dS[k][8]).toUpperCase() === "TRUE", journal: _maintJson_(dS[k][9], []), corr: _maintJson_(dS[k][10], {}),
      relances: Number(dS[k][11] || 0), derniereRelance: String(dS[k][12] || ""), responsableMail: String(dS[k][13] || ""), num: Number(dS[k][NC - 1] || 0) };
    var tickets = [];
    rep.lignes.forEach(function (l) {
      var r = l.r, v = function (c) { return c >= 0 ? String(r[c] || "").trim() : ""; };
      var s = suivi[l.cle] || suivi[l.ancienneCle] || {}, corr = s.corr || {};
      var orig = { campus: v(C.campus), lieu: v(C.lieu), categorie: v(C.categorie), description: v(C.description) };
      tickets.push({ cle: l.cle, ancienneCle: suivi[l.cle] ? "" : (suivi[l.ancienneCle] ? l.ancienneCle : ""), ligne: l.i + 1, numero: _maintNumLib_(s.num), num: s.num || 0,
        date: v(C.date), email: v(C.email), nom: v(C.nom),
        campus: corr.campus != null ? corr.campus : orig.campus, lieu: corr.lieu != null ? corr.lieu : orig.lieu,
        categorie: corr.categorie != null ? corr.categorie : orig.categorie, description: corr.description != null ? corr.description : orig.description,
        original: orig, corrige: Object.keys(corr).length > 0,
        urgence: v(C.urgence), photo: v(C.photo), champs: h.map(function (t, j) { return [String(t), String(r[j] || "")]; }).filter(function (x) { return x[1]; }),
        statut: s.statut || "Nouveau", responsable: s.responsable || "", responsableMail: s.responsableMail || "", commentaire: s.commentaire || "", priorite: s.priorite || "", resolu: s.resolu || "", maj: s.maj || "", par: s.par || "",
        archive: !!s.archive, journal: s.journal || [], relances: s.relances || 0, derniereRelance: s.derniereRelance || "" });
    });
    var nbSheet = tickets.length;
    _maintInternes_(ss).forEach(function (x) {
      var s = suivi[x.cle] || {}, corr = s.corr || {}, orig = { campus: x.campus, lieu: x.lieu, categorie: x.categorie, description: x.description };
      tickets.push({ cle: x.cle, ancienneCle: "", ligne: 0, numero: _maintNumLib_(s.num), num: s.num || 0, source: "application", date: x.date, email: x.email, nom: x.nom,
        campus: corr.campus != null ? corr.campus : orig.campus, lieu: corr.lieu != null ? corr.lieu : orig.lieu,
        categorie: corr.categorie != null ? corr.categorie : orig.categorie, description: corr.description != null ? corr.description : orig.description,
        original: orig, corrige: Object.keys(corr).length > 0, urgence: x.urgence, photo: "", champs: x.champs.concat([["Source", "Ouvert depuis l'application"]]),
        statut: s.statut || "Nouveau", responsable: s.responsable || "", responsableMail: s.responsableMail || "", commentaire: s.commentaire || "", priorite: s.priorite || "", resolu: s.resolu || "", maj: s.maj || "", par: s.par || "",
        archive: !!s.archive, journal: s.journal || [], relances: s.relances || 0, derniereRelance: s.derniereRelance || "" });
    });
    tickets.sort(function (x, y) { return (y.num || 0) - (x.num || 0) || y.ligne - x.ligne; });   // plus récents d'abord
    return createJsonResponse({ success: true, tickets: tickets, formUrl: formUrl, formPublic: _maintLienPublic_(formUrl), peutCreer: _maintMaitre_(p), sheetUrl: ext.getUrl(), titre: ext.getName(), colonnes: h, equipe: _maintEquipe_(ss), contacts: _maintContacts_(ss),
      controle: { onglet: ws.getName(), lignesSheet: rep.lignes.length, affiches: nbSheet, internes: tickets.length - nbSheet, numerotes: numerotes, colonnesReperees: C, lu: _fDate_(), dernier: tickets.length ? tickets[0].date : "" } });
  } catch (e2) { return createJsonResponse({ success: false, message: e2.toString() }); }
}

// Mise à jour d'un ou plusieurs tickets. p.op : "maj" (défaut, formulaire complet), "cloturer", "annuler",
// "rouvrir", "archiver", "desarchiver", "relancer". p.cle (un ticket) ou p.cles (action groupée).
// Le classeur des réponses n'est jamais modifié : les corrections de contenu vont dans Corrections_JSON.
function handleUpdateTicketMaintenance(p, ss) {
  if (!_fPeutGerer_(p)) return createJsonResponse({ success: false, message: "Accès réservé à l'équipe Vie scolaire." });
  if (String(p.op || "") === "logqr") { logAction(p.nomAuteur || p.emailAuteur || "", p.vieScRole || "", "Export", "Vie scolaire — Maintenance — QR code", "Formulaire", "QR code du formulaire de maintenance téléchargé"); return createJsonResponse({ success: true }); }
  if (String(p.op || "") === "creer") {   // ticket ouvert depuis l'application — admin maître
    if (!_maintMaitre_(p)) return createJsonResponse({ success: false, message: "Réservé à l'admin maître." });
    var desc = String(p.description || "").trim().slice(0, 3000);
    if (desc.length < 5) return createJsonResponse({ success: false, message: "Décrivez le problème." });
    var lockC = LockService.getScriptLock();
    if (!lockC.tryLock(20000)) return createJsonResponse({ success: false, message: "Serveur occupé, réessayez dans quelques secondes." });
    try {
      var wsI = _maintInternesFeuille_(ss), idI = "INT-" + Utilities.formatDate(new Date(), "Europe/Paris", "yyMMddHHmmss") + "-" + Math.floor(Math.random() * 90 + 10);
      var horo = Utilities.formatDate(new Date(), "Europe/Paris", "dd/MM/yyyy HH:mm:ss"), par0 = p.nomAuteur || p.emailAuteur || "";
      var ligneI = [idI, horo, String(p.mailDemandeur || p.emailAuteur || "").trim(), String(p.demandeur || par0).trim(), String(p.campus || "").trim(), String(p.lieu || "").trim(),
        String(p.categorie || "").trim(), String(p.priorite || "").trim(), desc, par0];
      wsI.getRange(wsI.getLastRow() + 1, 1, 1, ligneI.length).setNumberFormat("@").setValues([ligneI]);
      var wsS0 = _maintSuiviFeuille_(ss), NC0 = MAINT_ENTETES.length, dN = wsS0.getLastRow() > 1 ? wsS0.getRange(2, NC0, wsS0.getLastRow() - 1, 1).getValues() : [], max0 = 0;
      dN.forEach(function (r) { max0 = Math.max(max0, Number(r[0] || 0)); });
      var num0 = max0 + 1, maint0 = _fDate_();
      var ligneS = ["I|" + idI, p.responsable ? "Pris en charge" : "Nouveau", String(p.responsable || ""), "", String(p.priorite || ""), "", maint0, par0, false,
        JSON.stringify([{ d: maint0, par: par0, evt: "Ticket ouvert depuis l'application — n° " + _maintNumLib_(num0) }]), "{}", 0, "", String(p.responsableMail || ""), num0];
      wsS0.getRange(wsS0.getLastRow() + 1, 1, 1, NC0).setValues([ligneS]);
      logAction(par0, p.vieScRole || "", "Création", "Vie scolaire — Maintenance — ticket", _maintNumLib_(num0), "Ticket ouvert depuis l'application : " + desc.slice(0, 80));
      return createJsonResponse({ success: true, numero: _maintNumLib_(num0) });
    } finally { lockC.releaseLock(); }
  }
  if (String(p.op || "") === "contact") {   // création d'un contact externe
    var nomC = String(p.nom || "").trim().slice(0, 120), mailC = String(p.mail || "").trim().slice(0, 200);
    if (!nomC) return createJsonResponse({ success: false, message: "Nom du contact obligatoire." });
    if (mailC && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(mailC)) return createJsonResponse({ success: false, message: "Adresse mail invalide." });
    if (_maintContacts_(ss).some(function (c) { return c.nom.toLowerCase() === nomC.toLowerCase(); })) return createJsonResponse({ success: false, message: "Un contact porte déjà ce nom." });
    var cTel = String(p.tel || "").trim().slice(0, 40), cSoc = String(p.societe || "").trim().slice(0, 120);
    _maintContactsFeuille_(ss).appendRow([nomC, mailC, cTel, cSoc, _fDate_(), p.nomAuteur || p.emailAuteur || ""]);
    logAction(p.nomAuteur || p.emailAuteur || "", p.vieScRole || "", "Création", "Vie scolaire — Maintenance — contacts", nomC, "Contact externe créé" + (cSoc ? " (" + cSoc + ")" : ""));
    return createJsonResponse({ success: true, contact: { nom: nomC, mail: mailC, tel: cTel, societe: cSoc } });
  }
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return createJsonResponse({ success: false, message: "Serveur occupé, réessayez dans quelques secondes." });
  try {
    var ws = _maintSuiviFeuille_(ss), data = ws.getRange(1, 1, Math.max(ws.getLastRow(), 1), MAINT_ENTETES.length).getValues();
    var cles = Array.isArray(p.cles) && p.cles.length ? p.cles.map(String) : (p.cle ? [String(p.cle)] : []);
    if (!cles.length) return createJsonResponse({ success: false, message: "Ticket inconnu." });
    var op = String(p.op || "maj"), par = p.nomAuteur || p.emailAuteur || "", maintenant = _fDate_();
    var equipe = op === "relancer" ? _maintEquipe_(ss) : [], resumes = p.resumes || {}, resultats = [], relancesEnvoyees = 0, erreurs = [];
    var index = {}; for (var i = 1; i < data.length; i++) if (data[i][0]) index[String(data[i][0])] = i;
    var aAjouter = [];

    cles.forEach(function (cle) {
      var anc = p.anciennesCles && p.anciennesCles[cle];
      if (index[cle] == null && anc && index[anc] != null && index[anc] >= 0) { index[cle] = index[anc]; data[index[cle]][0] = cle; }
      var iL = index[cle], ligne = iL != null ? data[iL].slice() : [cle, "Nouveau", "", "", "", "", "", "", false, "[]", "{}", 0, "", "", ""];
      while (ligne.length < MAINT_ENTETES.length) ligne.push("");
      var ancienCom = String(ligne[3] || ""), ancienResp = String(ligne[2] || "");
      var avant = String(ligne[1] || "Nouveau"), journal = _maintJson_(ligne[9], []), corr = _maintJson_(ligne[10], {}), evt = "";
      var resume = String(resumes[cle] || p.resume || "").slice(0, 80), numLib = _maintNumLib_(ligne[MAINT_ENTETES.length - 1]);
      if (numLib) resume = numLib + " — " + resume;
      if (op === "maj") {
        var statut = String(p.statut || avant);
        ligne[1] = statut; ligne[2] = String(p.responsable || ""); ligne[13] = String(p.responsableMail || "").trim(); ligne[3] = String(p.commentaire || "").slice(0, 2000); ligne[4] = String(p.priorite || "");
        if (p.corrections && typeof p.corrections === "object") {
          ["campus", "lieu", "categorie", "description"].forEach(function (k) {
            if (p.corrections[k] === undefined) return;
            var val = String(p.corrections[k] || "").slice(0, 2000), orig = p.originaux ? String(p.originaux[k] || "") : null;
            if (orig !== null && val === orig) delete corr[k]; else corr[k] = val;
          });
        }
        evt = statut !== avant ? avant + " → " + statut : "Modifié";
      } else if (op === "cloturer") { ligne[1] = "Clôturé"; evt = avant + " → Clôturé"; }
      else if (op === "annuler") { ligne[1] = "Annulé"; evt = avant + " → Annulé"; }
      else if (op === "rouvrir") { ligne[1] = "Pris en charge"; ligne[8] = false; evt = avant + " → rouvert"; }
      else if (op === "archiver") { ligne[8] = true; evt = "Archivé"; }
      else if (op === "desarchiver") { ligne[8] = false; evt = "Désarchivé"; }
      else if (op === "relancer") {
        var resp = String(ligne[2] || p.responsable || "").trim(), mail = String(ligne[13] || "").trim();
        if (mail.indexOf("@") < 1) mail = "";
        if (!mail && resp.indexOf("@") > 0) mail = resp;
        else { var n = resp.toLowerCase(); equipe.forEach(function (x) { if (!mail && x.mail && n && (x.nom.toLowerCase() === n || x.nom.toLowerCase().indexOf(n) !== -1 || n.indexOf(x.nom.toLowerCase()) !== -1)) mail = x.mail; }); }
        if (!mail) { erreurs.push((resume || cle) + " : aucun responsable avec adresse mail"); return; }
        if (MAINT_FERMES.indexOf(avant) !== -1) ligne[1] = "Pris en charge";
        var esc = function (x) { return String(x == null ? "" : x).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); };
        var msg = String(p.message || "").trim();
        var corps = '<div style="font-family:Arial;padding:20px"><h2 style="color:#1A4E8A;margin:0 0 10px">Relance — ticket de maintenance</h2><p>Bonjour,</p><p>Le ticket ci-dessous, dont vous êtes responsable, est toujours en attente de traitement' +
          (Number(ligne[11] || 0) ? " (relance n° " + (Number(ligne[11] || 0) + 1) + ")" : "") + ".</p>" + (msg ? '<p style="background:#fff7ed;border-left:4px solid #F26522;padding:10px">' + esc(msg).replace(/\n/g, "<br>") + "</p>" : "") +
          '<div style="background:#f8fafc;border-left:4px solid #1A4E8A;padding:12px;margin:14px 0"><b>' + esc(resume || "Signalement") + "</b><br><span style=\"color:#64748b;font-size:13px\">Statut : " + esc(ligne[1]) +
          (ligne[4] ? " · Priorité : " + esc(ligne[4]) : "") + (ligne[3] ? "<br>Suivi : " + esc(ligne[3]) : "") + "</span></div><p>Merci de mettre à jour le ticket dans Vie scolaire › Tickets maintenance.</p><p style=\"color:#64748b;font-size:13px\">Relance envoyée par " + esc(par) + ".</p></div>";
        try {
          var opt = { htmlBody: corps + getSignatureHTML() };
          if (mail.toLowerCase() !== "m.mokhtari@hecg.fr") opt.cc = "m.mokhtari@hecg.fr";
          if (p.emailAuteur && String(p.emailAuteur).indexOf("@") > 0) opt.replyTo = String(p.emailAuteur);
          GmailApp.sendEmail(mail, "Relance — ticket de maintenance : " + (resume || "signalement"), "", opt);
          relancesEnvoyees++;
        } catch (eM) { erreurs.push((resume || cle) + " : " + eM.message); return; }
        ligne[11] = Number(ligne[11] || 0) + 1; ligne[12] = maintenant; evt = "Relance envoyée à " + resp + " (" + mail + ")";
      } else return;
      var statutFinal = String(ligne[1]);
      if (statutFinal === "Résolu" || statutFinal === "Clôturé") { if (!ligne[5]) ligne[5] = new Date(); } else ligne[5] = "";
      ligne[6] = maintenant; ligne[7] = par;
      journal.push({ d: maintenant, par: par, evt: evt + (op === "maj" && String(p.commentaire || "") !== ancienCom ? " — suivi mis à jour" : "") + (op === "maj" && String(ligne[2]) !== ancienResp ? " — responsable : " + (ligne[2] || "aucun") : "") });
      if (journal.length > 60) journal = journal.slice(-60);
      ligne[9] = JSON.stringify(journal); ligne[10] = JSON.stringify(corr);
      if (iL != null) data[iL] = ligne; else { aAjouter.push(ligne); index[cle] = -1; }
      resultats.push({ cle: cle, statut: statutFinal, archive: ligne[8] === true, journal: journal, relances: Number(ligne[11] || 0), derniereRelance: String(ligne[12] || ""), responsable: String(ligne[2] || ""), responsableMail: String(ligne[13] || ""),
        resolu: ligne[5] ? Utilities.formatDate(new Date(ligne[5]), "Europe/Paris", "dd/MM/yyyy") : "", maj: maintenant, par: par, corr: corr });
      logAction(par, p.vieScRole || "", "Modification", "Vie scolaire — Maintenance" + (op !== "maj" ? " — " + op : ""), cle, "Ticket « " + resume + " » : " + evt);
    });
    if (data.length > 1) ws.getRange(1, 1, data.length, MAINT_ENTETES.length).setValues(data.map(function (r) { while (r.length < MAINT_ENTETES.length) r.push(""); return r.slice(0, MAINT_ENTETES.length); }));
    if (aAjouter.length) ws.getRange(ws.getLastRow() + 1, 1, aAjouter.length, MAINT_ENTETES.length).setValues(aAjouter);
    return createJsonResponse({ success: resultats.length > 0 || !erreurs.length, resultats: resultats, relances: relancesEnvoyees, erreurs: erreurs,
      message: erreurs.length ? erreurs.join(" ; ") : "" });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
  finally { lock.releaseLock(); }
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
