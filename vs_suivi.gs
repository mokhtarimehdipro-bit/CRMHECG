// =============================================================================
//  VIE SCOLAIRE — SUIVI DES RDV EN ENTREPRISE · COMPARATIF DIRECTEURS · EXAMENS OFFICIELS
//  (demande direction 2026-11-12) — front : vs_modules.js (VX.rdv*, VX.dir*, VX.ex*)
// =============================================================================

// ── Socle : étudiants du CRM (lecture unique de ETUDIANTS) ───────────────────
function _vsuEtudiants_(ss) {
  var d = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues(), out = [];
  for (var i = 1; i < d.length; i++) {
    var r = d[i], id = String(r[0] || "").trim();
    if (!id || id.toUpperCase() === "ETU000X") continue;
    var typo = String(r[20] || "").trim();
    var tn = typo.toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[^a-z]/g, "");
    out.push({ id: id, nom: String(r[2] || "").trim(), prenom: String(r[3] || "").trim(), mail: String(r[26] || "").trim() || String(r[4] || "").trim(),
      mailPerso: String(r[4] || "").trim().toLowerCase(), mailHecg: String(r[26] || "").trim().toLowerCase(),
      idEnt: String(r[11] || "").trim(), statut: String(r[13] || "").trim(), classe: String(r[14] || "").trim(), nomEnt: String(r[21] || "").trim(),
      typo: typo, inscrit: tn === "inscrit" || tn === "preinscrit", campusCode: String(r[22] || "").trim(), campus: _vsxCampusLib_(r[22]),
      idEleve: String(r[23] || "").trim() });
  }
  return out;
}
function _vsuAlternant_(e) { return /altern|apprenti|contrat/i.test(e.statut) || !!e.idEnt; }
function _vsuDiplome_(classe) { return _ckDiplome_(classe) || "Autre"; }
function _vsuNorm_(s) { return String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim(); }
function _vsuIso_(d) { return d instanceof Date && !isNaN(d) ? Utilities.formatDate(d, "Europe/Paris", "yyyy-MM-dd") : String(d || ""); }
function _vsuAnneeDe_(iso) {   // année scolaire (sept. → août) d'une date ISO
  var m = String(iso || "").match(/^(\d{4})-(\d{2})/); if (!m) return _ckAnneeCourante_();
  var y = +m[1]; return +m[2] >= 9 ? y + "-" + (y + 1) : (y - 1) + "-" + y;
}
function _vsuAdmin_(p) { return _vsxIsSuper_(p) || String(p.vieScRole || "") === "Administratif"; }
// Entreprises : nom + adresse de référence (une adresse par entreprise, adresses.gs)
function _vsuEntreprises_(ss) {
  var d = getSheetSafe(ss, "PARTENARIAT").getDataRange().getValues();
  var C = _adrColsLecture_(d[0]), g = {};
  for (var i = 1; i < d.length; i++) {
    var id = String(d[i][1] || "").trim().toUpperCase(); if (!id) continue;
    g[id] = g[id] || { id: String(d[i][1]).trim(), nom: String(d[i][2] || "").trim(), lignes: [] };
    g[id].lignes.push(_adrDeLigne_(d[i], C));
  }
  var out = {};
  Object.keys(g).forEach(function (k) { var a = _adrReference_(g[k].lignes); out[k] = { id: g[k].id, nom: g[k].nom, lieu: [a.adresse, a.complement, a.cp, a.ville].filter(Boolean).join(" ") }; });
  return out;
}

// ═══════════════════════════ 1. SUIVI DES RDV EN ENTREPRISE ═══════════════════════════
// RDV_VISITES : une ligne par rendez-vous (on peut visiter plusieurs fois un étudiant dans l'année).
// L'entreprise et le lieu sont FIGÉS à la création : l'historique dit « tel jour, untel a visité tel
// étudiant dans telle structure », même si l'étudiant change d'entreprise ensuite.
// « Date RDV pris » = date initiale ; « Date RDV » = date en vigueur (après report).
var RDV_HEADERS = ["ID", "Année scolaire", "ID étudiant", "Étudiant", "Classe", "ID entreprise", "Entreprise", "Lieu", "Visiteur", "Email visiteur",
  "Date RDV pris", "Date RDV", "Statut", "Historique", "Commentaire", "Source", "Créé le", "Créé par", "Mis à jour le", "Archivé"];
var RDV_REF_HEADERS = ["Année scolaire", "ID étudiant", "Responsable visite", "Email", "Source", "Mis à jour le", "Par"];
var RDV_STATUTS = ["À planifier", "Planifié", "Honoré", "Reporté", "Annulé", "Autre"];

function _rdvSheet_(ss) { var ws = _vsxSheet_(ss, "RDV_VISITES", RDV_HEADERS); return ws; }
function _rdvLire_(ss) {
  var d = _rdvSheet_(ss).getDataRange().getValues(), out = [];
  for (var i = 1; i < d.length; i++) {
    var r = d[i]; if (!r[0] || String(r[19]) === "Oui") continue;
    var h = []; try { h = JSON.parse(String(r[13] || "[]")); } catch (e) { h = []; }
    out.push({ id: String(r[0]), annee: String(r[1]), etuId: String(r[2]), etudiant: String(r[3]), classe: String(r[4]), entId: String(r[5]), entreprise: String(r[6]),
      lieu: String(r[7]), visiteur: String(r[8]), visiteurEmail: String(r[9]), datePris: _vsuIso_(r[10]), date: _vsuIso_(r[11]), statut: String(r[12] || "Planifié"),
      historique: h, commentaire: String(r[14] || ""), source: String(r[15] || ""), creeLe: _vsuIso_(r[16]), creePar: String(r[17] || ""), row: i + 1 });
  }
  return out;
}
function _rdvReferents_(ss, annee) {
  var d = _vsxSheet_(ss, "RDV_REFERENTS", RDV_REF_HEADERS).getDataRange().getValues(), out = {};
  for (var i = 1; i < d.length; i++) if (String(d[i][0]) === annee && d[i][1]) out[String(d[i][1])] = { nom: String(d[i][2] || ""), email: String(d[i][3] || ""), source: String(d[i][4] || ""), row: i + 1 };
  return out;
}

function handleGetRdv(p, ss) {
  if (!_vsxCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  try {
    var annee = String(p.annee || _ckAnneeCourante_());
    var etus = _vsuEtudiants_(ss), ents = _vsuEntreprises_(ss), visites = _rdvLire_(ss), refs = _rdvReferents_(ss, annee);
    var avecVisite = {}; visites.forEach(function (v) { avecVisite[v.etuId] = true; });
    var liste = [];
    etus.forEach(function (e) {
      if (!e.inscrit && !(avecVisite[e.id] && visites.some(function (v) { return v.etuId === e.id && v.annee === annee; }))) return;
      var ent = e.idEnt ? ents[e.idEnt.toUpperCase()] : null;
      liste.push({ id: e.id, nom: (e.prenom + " " + e.nom).trim(), classe: e.classe, diplome: _vsuDiplome_(e.classe), campus: e.campus, typo: e.typo,
        alternant: _vsuAlternant_(e), statut: e.statut, entId: ent ? ent.id : "", entreprise: ent ? ent.nom : e.nomEnt, lieu: ent ? ent.lieu : "",
        referent: refs[e.id] ? refs[e.id].nom : "", referentEmail: refs[e.id] ? refs[e.id].email : "", referentSource: refs[e.id] ? refs[e.id].source : "" });
    });
    var annees = {}; annees[_ckAnneeCourante_()] = 1; visites.forEach(function (v) { if (v.annee) annees[v.annee] = 1; });
    var equipe = _vsxRef_(ss).personnel.filter(function (x) { return x.email; }).map(function (x) { return { nom: x.nom, email: x.email, role: x.role }; });
    visites.forEach(function (v) { delete v.row; });
    return createJsonResponse({ success: true, annee: annee, annees: Object.keys(annees).sort().reverse(), etudiants: liste, visites: visites,
      equipe: equipe, peutSuivre: _vsuAdmin_(p), envois: _rdvEnvois_(ss, annee),
      liens: _rdvResponsables_(ss).filter(function (r) { return r.lien; }).map(function (r) { return r.email || r.nom; }) });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

/** Crée / modifie / retire un RDV. */
function handleSaveRdv(p, ss) {
  if (!_vsxCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var ws = _rdvSheet_(ss), v = p.rdv || {}, now = new Date(), par = String(p.nomAuteur || p.emailAuteur || "");
    var date = /^\d{4}-\d{2}-\d{2}$/.test(String(v.date || "")) ? String(v.date) : "";
    if (v.id) {
      var d = ws.getDataRange().getValues();
      for (var i = 1; i < d.length; i++) if (String(d[i][0]) === String(v.id)) {
        if (v.archiver) {
          ws.getRange(i + 1, 19, 1, 2).setValues([[now, "Oui"]]);
          logAction(par, p.vieScRole || "", "Suppression", "Vie scolaire — Suivi des RDV", String(v.id), "RDV retiré : " + d[i][3] + " (" + d[i][6] + ", " + _vsuIso_(d[i][11]) + ")");
          return createJsonResponse({ success: true });
        }
        var row = d[i].slice(0, RDV_HEADERS.length);
        if (v.visiteur !== undefined) { row[8] = String(v.visiteur || "").trim(); row[9] = String(v.visiteurEmail || "").trim(); }
        if (v.commentaire !== undefined) row[14] = String(v.commentaire || "");
        if (date && String(row[12]) !== "Honoré") {
          if (!row[10]) row[10] = date;
          if (["Planifié", "À planifier"].indexOf(String(row[12])) !== -1) { row[10] = date; row[12] = "Planifié"; }
          row[11] = date;
        }
        row[18] = now;
        ws.getRange(i + 1, 1, 1, RDV_HEADERS.length).setValues([row]);
        ws.getRange(i + 1, 11, 1, 2).setNumberFormat("@").setValues([[_vsuIso_(row[10]), _vsuIso_(row[11])]]);
        logAction(par, p.vieScRole || "", "Modification", "Vie scolaire — Suivi des RDV", String(v.id), "RDV modifié : " + row[3] + " — " + (row[8] || "visiteur ?") + " le " + _vsuIso_(row[11]));
        return createJsonResponse({ success: true, id: String(v.id) });
      }
      return createJsonResponse({ success: false, message: "RDV introuvable." });
    }
    if (!v.etuId) return createJsonResponse({ success: false, message: "Étudiant manquant." });
    var e = _vsuEtudiants_(ss).filter(function (x) { return x.id === String(v.etuId); })[0];
    if (!e) return createJsonResponse({ success: false, message: "Étudiant introuvable." });
    var ents = _vsuEntreprises_(ss), ent = e.idEnt ? ents[e.idEnt.toUpperCase()] : null;
    if (v.entId && ents[String(v.entId).toUpperCase()]) ent = ents[String(v.entId).toUpperCase()];
    var id = "RDV" + now.getTime();
    var annee = String(v.annee || (date ? _vsuAnneeDe_(date) : _ckAnneeCourante_()));
    var hist = [{ le: Utilities.formatDate(now, "Europe/Paris", "yyyy-MM-dd HH:mm"), statut: date ? "Planifié" : "À planifier", date: date, par: par }];
    ws.appendRow([id, annee, e.id, (e.prenom + " " + e.nom).trim(), e.classe, ent ? ent.id : "", ent ? ent.nom : (String(v.entreprise || "") || e.nomEnt), ent ? ent.lieu : String(v.lieu || ""),
      String(v.visiteur || "").trim(), String(v.visiteurEmail || "").trim(), "", "", date ? "Planifié" : "À planifier", JSON.stringify(hist), String(v.commentaire || ""), "Saisie", now, par, now, ""]);
    ws.getRange(ws.getLastRow(), 11, 1, 2).setNumberFormat("@").setValues([[date, date]]);
    logAction(par, p.vieScRole || "", "Création", "Vie scolaire — Suivi des RDV", id, "RDV " + (date ? "pris le " + date : "à planifier") + " : " + (e.prenom + " " + e.nom).trim() + " chez " + (ent ? ent.nom : "?") + " — visiteur " + (v.visiteur || "?"));
    return createJsonResponse({ success: true, id: id });
  } catch (err) { return createJsonResponse({ success: false, message: err.toString() }); }
  finally { try { lock.releaseLock(); } catch (eL) {} }
}

/** Suivi (profils Super-admin / Administratif) : Honoré, Reporté (avec date), Annulé, Autre, Planifié. */
function handleStatutRdv(p, ss) {
  if (!_vsuAdmin_(p)) return createJsonResponse({ success: false, message: "Réservé aux profils Super-admin et Administratif (suivi des RDV)." });
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var statut = String(p.statut || "");
    if (["Planifié", "Honoré", "Reporté", "Annulé", "Autre"].indexOf(statut) === -1) return createJsonResponse({ success: false, message: "Statut inconnu." });
    var dateReport = /^\d{4}-\d{2}-\d{2}$/.test(String(p.dateReport || "")) ? String(p.dateReport) : "";
    if (statut === "Reporté" && !dateReport) return createJsonResponse({ success: false, message: "Indiquez la nouvelle date du RDV." });
    if (statut === "Autre" && !String(p.commentaire || "").trim()) return createJsonResponse({ success: false, message: "Précisez ce qui s'est passé." });
    var ws = _rdvSheet_(ss), d = ws.getDataRange().getValues(), now = new Date(), par = String(p.nomAuteur || p.emailAuteur || "");
    for (var i = 1; i < d.length; i++) if (String(d[i][0]) === String(p.id)) {
      var h = []; try { h = JSON.parse(String(d[i][13] || "[]")); } catch (e) { h = []; }
      var avant = _vsuIso_(d[i][11]);
      h.push({ le: Utilities.formatDate(now, "Europe/Paris", "yyyy-MM-dd HH:mm"), statut: statut, date: avant, dateReport: dateReport, commentaire: String(p.commentaire || ""), par: par });
      var nouvelleDate = statut === "Reporté" ? dateReport : avant;
      ws.getRange(i + 1, 12, 1, 3).setValues([[nouvelleDate, statut, JSON.stringify(h)]]);
      ws.getRange(i + 1, 12).setNumberFormat("@").setValue(nouvelleDate);
      if (p.commentaire) ws.getRange(i + 1, 15).setValue(String(d[i][14] ? d[i][14] + "\n" : "") + Utilities.formatDate(now, "Europe/Paris", "dd/MM/yyyy") + " — " + statut + " : " + p.commentaire);
      ws.getRange(i + 1, 19).setValue(now);
      logAction(par, p.vieScRole || "", "Modification", "Vie scolaire — Suivi des RDV", String(p.id),
        "RDV " + statut.toLowerCase() + (statut === "Reporté" ? " au " + dateReport : "") + " : " + d[i][3] + " chez " + d[i][6] + " (prévu le " + avant + ")" + (p.commentaire ? " — " + p.commentaire : ""));
      return createJsonResponse({ success: true });
    }
    return createJsonResponse({ success: false, message: "RDV introuvable." });
  } catch (err) { return createJsonResponse({ success: false, message: err.toString() }); }
  finally { try { lock.releaseLock(); } catch (eL) {} }
}

/** Responsable des visites d'un étudiant pour l'année (saisi ou repris de Charlemagne). */
function handleSaveReferentRdv(p, ss) {
  if (!_vsxCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  try {
    var annee = String(p.annee || _ckAnneeCourante_()), ids = Array.isArray(p.etuIds) ? p.etuIds.map(String) : [String(p.etuId || "")];
    var n = _rdvEcrireReferents_(ss, annee, ids.map(function (id) { return { etuId: id, nom: String(p.nom || "").trim(), email: String(p.email || "").trim() }; }), "Saisie", String(p.nomAuteur || p.emailAuteur || ""));
    logAction(p.nomAuteur || p.emailAuteur || "", p.vieScRole || "", "Modification", "Vie scolaire — Suivi des RDV", ids.length === 1 ? ids[0] : "RDV_REFERENTS",
      "Responsable visite " + annee + " : " + (p.nom || "(retiré)") + " pour " + ids.length + " étudiant(s)");
    return createJsonResponse({ success: true, n: n });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}
function _rdvEcrireReferents_(ss, annee, items, source, par) {
  var ws = _vsxSheet_(ss, "RDV_REFERENTS", RDV_REF_HEADERS), refs = _rdvReferents_(ss, annee), now = new Date(), n = 0, ajout = [];
  items.forEach(function (it) {
    if (!it.etuId) return;
    var r = refs[it.etuId];
    if (r) {
      if (source === "Charlemagne" && r.source === "Saisie") return;   // une saisie manuelle n'est jamais écrasée par l'import
      ws.getRange(r.row, 3, 1, 5).setValues([[it.nom, it.email, source, now, par]]); n++;
    } else if (it.nom) { ajout.push([annee, it.etuId, it.nom, it.email, source, now, par]); n++; }
  });
  if (ajout.length) ws.getRange(ws.getLastRow() + 1, 1, ajout.length, RDV_REF_HEADERS.length).setValues(ajout);
  return n;
}

// Import Charlemagne : colonnes dont l'en-tête contient « visite » (ex. « Responsable visite »).
// Le responsable devient le responsable des visites de l'année ; chaque date trouvée (dans une colonne
// « date … visite » ou dans le texte) crée une visite (passée = Honoré, future = Planifié), sans doublon.
function _rdvImporterCharl_(ss, lignes, par) {
  var etus = _vsuEtudiants_(ss), parId = {}, parMail = {}, parNom = {};
  etus.forEach(function (e) {
    if (e.idEleve) parId[e.idEleve] = e;
    [e.mailPerso, e.mailHecg].forEach(function (m) { if (m) parMail[m] = e; });
    parNom[_vsuNorm_(e.nom + " " + e.prenom)] = parNom[_vsuNorm_(e.nom + " " + e.prenom)] || e;
  });
  var equipe = _vsxRef_(ss).personnel, ents = _vsuEntreprises_(ss);
  var mailDe = function (nom) {
    var n = _vsuNorm_(nom); if (!n) return "";
    var t = n.split(" ");
    var x = equipe.filter(function (q) { var qn = _vsuNorm_(q.nom); return t.every(function (w) { return w.length < 2 || qn.indexOf(w) !== -1; }); })[0];
    return x ? x.email : "";
  };
  var existants = {}; _rdvLire_(ss).forEach(function (v) { existants[v.etuId + "|" + v.date] = 1; });
  var annee = _ckAnneeCourante_(), refs = [], nouvelles = [], now = new Date(), auj = Utilities.formatDate(now, "Europe/Paris", "yyyy-MM-dd");
  var stats = { lignesAvecVisite: 0, responsables: 0, visites: 0, nonRapproches: 0 };
  lignes.forEach(function (l) {
    var vis = []; try { vis = JSON.parse(l.visites || "[]"); } catch (e) { vis = []; }
    vis = vis.filter(function (x) { return x && String(x.v || "").trim(); });
    if (!vis.length) return;
    stats.lignesAvecVisite++;
    var e = (l.idEleve && parId[l.idEleve]) || parMail[String(l.email || "").toLowerCase()] || parMail[String(l.mailPerso || "").toLowerCase()] || parNom[_vsuNorm_(l.nom + " " + l.prenom)];
    if (!e) { stats.nonRapproches++; return; }
    var noms = [], dates = [];
    vis.forEach(function (x) {
      var h = _vsuNorm_(x.h), v = String(x.v).trim();
      var reDate = /(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})/g, m, trouve = false;
      while ((m = reDate.exec(v))) { var y = +m[3] < 100 ? 2000 + +m[3] : +m[3]; dates.push(y + "-" + ("0" + m[2]).slice(-2) + "-" + ("0" + m[1]).slice(-2)); trouve = true; }
      if (/^\d{4}-\d{2}-\d{2}/.test(v)) { dates.push(v.slice(0, 10)); trouve = true; }
      var reste = v.replace(/(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})/g, " ").replace(/\b(le|du|au|visite|rdv)\b/gi, " ").replace(/\s+/g, " ").trim();
      if (h.indexOf("date") === -1 && reste && /[a-z]{2}/i.test(reste)) noms.push(reste);
    });
    var nom = noms[0] || "";
    if (nom) { refs.push({ etuId: e.id, nom: nom, email: mailDe(nom) }); stats.responsables++; }
    var ent = e.idEnt ? ents[e.idEnt.toUpperCase()] : null;
    dates.filter(function (x, i, a) { return a.indexOf(x) === i; }).forEach(function (dt) {
      if (existants[e.id + "|" + dt]) return;
      existants[e.id + "|" + dt] = 1;
      var st = dt <= auj ? "Honoré" : "Planifié";
      nouvelles.push(["RDV" + now.getTime() + "C" + nouvelles.length, _vsuAnneeDe_(dt), e.id, (e.prenom + " " + e.nom).trim(), e.classe, ent ? ent.id : "", ent ? ent.nom : e.nomEnt, ent ? ent.lieu : "",
        nom, mailDe(nom), dt, dt, st, JSON.stringify([{ le: Utilities.formatDate(now, "Europe/Paris", "yyyy-MM-dd HH:mm"), statut: st, date: dt, par: "Import Charlemagne", commentaire: vis.map(function (x) { return x.h + " : " + x.v; }).join(" | ") }]),
        "Repris de Charlemagne : " + vis.map(function (x) { return x.h + " = " + x.v; }).join(" ; "), "Charlemagne", now, par, now, ""]);
    });
  });
  if (refs.length) _rdvEcrireReferents_(ss, annee, refs, "Charlemagne", par);
  if (nouvelles.length) {
    var ws = _rdvSheet_(ss), r0 = ws.getLastRow() + 1;
    ws.getRange(r0, 11, nouvelles.length, 2).setNumberFormat("@");
    ws.getRange(r0, 1, nouvelles.length, RDV_HEADERS.length).setValues(nouvelles);
  }
  stats.visites = nouvelles.length;
  return stats;
}
/** Re-synchronise depuis le dernier import Charlemagne (feuille CHARLEMAGNE_IMPORT). */
function handleImporterRdvCharlemagne(p, ss) {
  if (!_vsuAdmin_(p)) return createJsonResponse({ success: false, message: "Réservé aux profils Super-admin et Administratif." });
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var ws = ss.getSheetByName("CHARLEMAGNE_IMPORT");
    if (!ws || ws.getLastRow() < 3) return createJsonResponse({ success: false, message: "Aucun import Charlemagne : importez d'abord le fichier dans « Contrôle Charlemagne »." });
    var d = ws.getDataRange().getValues(), h = d[1].map(String);
    if (h.indexOf("visites") === -1) return createJsonResponse({ success: false, message: "Le dernier import Charlemagne date d'avant le suivi des RDV : réimportez le fichier dans « Contrôle Charlemagne »." });
    var lignes = d.slice(2).map(function (r) { var o = {}; h.forEach(function (k, i) { o[k] = String(r[i] == null ? "" : r[i]); }); return o; });
    var S = _rdvImporterCharl_(ss, lignes, String(p.nomAuteur || p.emailAuteur || ""));
    logAction(p.nomAuteur || p.emailAuteur || "", p.vieScRole || "", "Action", "Vie scolaire — Suivi des RDV", "RDV_VISITES",
      "Reprise Charlemagne : " + S.lignesAvecVisite + " étudiant(s) avec une colonne visite, " + S.responsables + " responsable(s), " + S.visites + " visite(s) créée(s), " + S.nonRapproches + " non rapproché(s)");
    return createJsonResponse({ success: true, stats: S });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
  finally { try { lock.releaseLock(); } catch (eL) {} }
}

// ═══════════════════════════ 2. COMPARATIF AVEC LE SUIVI DES DIRECTEURS ═══════════════════════════
// LECTURE SEULE du classeur « Suivi inscription étudiants » (aucune écriture, ni dans le classeur ni
// dans l'application). Onglets en colonnes (une colonne par classe), une case = un étudiant :
// fond ROUGE = plus dans l'établissement, fond VERT = en alternance.
var DIR_SHEET_ID_DEFAUT = "1YUKel1_uROkxg9BmHafZ3z-40g2GhrarhDiXP5FURrE";
var DIR_CLASSE_RE = /(BTS|DSCG|DCG|GPME|GMPE|BACHELOR|MASTER|PREPA|BAC\s?PRO|LICENCE)/i;   // « DCG1A », « DSCG1B ?? » : pas de \b (chiffre collé)

function _dirCouleur_(hex) {
  var m = String(hex || "").match(/^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})/i);
  if (!m) return "blanc";
  var r = parseInt(m[1], 16) / 255, g = parseInt(m[2], 16) / 255, b = parseInt(m[3], 16) / 255;
  if (r > 0.9 && g > 0.9 && b > 0.9) return "blanc";
  if (r >= 0.75 && g <= 0.4 && b <= 0.4) return "rouge";
  if (g >= 0.45 && g - r >= 0.12 && g - b >= 0.12) return "vert";
  return "autre";
}
function _dirNomAffiche_(t) {
  var s = String(t || "").replace(/^\s*x\s+/i, "").split(/\(|\s-\s|\s–\s|:/)[0].trim();
  var mots = s.split(/\s+/), out = [];
  for (var i = 0; i < mots.length && out.length < 5; i++) {
    var w = mots[i];
    if (out.length >= 2 && (/^[a-zà-ÿ]/.test(w) || (w.length <= 2 && /^[A-Z][a-z]$/.test(w)))) break;   // « ct envoyé… », « rompu… », « Nv contrat »
    if (/\d/.test(w)) break;
    out.push(w);
  }
  return out.join(" ") || s.slice(0, 40);
}
function _dirTokens_(t) {
  return _vsuNorm_(String(t || "").replace(/^\s*x\s+/i, "")).split(" ").filter(function (w) { return w.length >= 2; });
}

function handleGetComparatifDirecteurs(p, ss) {
  if (!_vsxCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  try {
    var sid = PropertiesService.getScriptProperties().getProperty("DIR_SHEET_ID") || DIR_SHEET_ID_DEFAUT;
    var book;
    try { book = SpreadsheetApp.openById(sid); }
    catch (eO) { return createJsonResponse({ success: false, message: "Impossible d'ouvrir le classeur des directeurs (" + eO.message + "). Partagez-le en lecture avec m.mokhtari@hecg.fr." }); }
    var annee = _ckAnneeCourante_();
    var tous = book.getSheets().map(function (s) { return s.getName(); });
    var choix = Array.isArray(p.onglets) && p.onglets.length ? p.onglets.map(String)
      : tous.filter(function (n) { return n.indexOf(annee) !== -1 && !/initiaux/i.test(n); });
    if (!choix.length) return createJsonResponse({ success: true, onglets: tous, choisis: [], message: "Aucun onglet pour " + annee + " : choisissez-les dans la liste." });

    // 1. Lecture du classeur (valeurs affichées + couleurs de fond)
    var cases = [], couleursAutres = {};
    choix.forEach(function (nomOnglet) {
      var sh = book.getSheetByName(nomOnglet); if (!sh) return;
      var rg = sh.getDataRange(), vals = rg.getDisplayValues(), bgs = rg.getBackgrounds();
      var hr = -1;
      for (var i = 0; i < Math.min(vals.length, 15) && hr < 0; i++) if (vals[i].filter(function (c) { return DIR_CLASSE_RE.test(c); }).length >= 2) hr = i;
      if (hr < 0) return;
      var campus = /aix/i.test(nomOnglet) ? "Aix-en-Provence" : /mrs|marseille/i.test(nomOnglet) ? "Marseille" : "";
      vals[hr].forEach(function (h, c) {
        if (!DIR_CLASSE_RE.test(h)) return;
        var classe = String(h).replace(/\(.*?\)/g, " ").replace(/\?+/g, " ").replace(/\s+/g, " ").trim();
        for (var r = hr + 1; r < vals.length; r++) {
          var t = String(vals[r][c] || "").trim();
          if (!t || /^\d+$/.test(t) || _dirTokens_(t).length < 2 || DIR_CLASSE_RE.test(t) && t.length < 25 && !/[a-z]{3}\s+[A-Z]{3}/.test(t)) continue;   // case vide, numéro, « x », en-tête
          var coul = _dirCouleur_(bgs[r][c]);
          if (coul === "autre") couleursAutres[bgs[r][c]] = (couleursAutres[bgs[r][c]] || 0) + 1;
          cases.push({ onglet: nomOnglet, campus: campus, classeDir: classe, diplome: _vsuDiplome_(classe), texte: t, nom: _dirNomAffiche_(t),
            couleur: bgs[r][c], etat: coul === "rouge" ? "absent" : "present", vert: coul === "vert", autreCouleur: coul === "autre", ref: nomOnglet + "!" + String.fromCharCode(65 + c) + (r + 1) });
        }
      });
    });

    // 2. Rapprochement avec le CRM (tous les noms de la fiche doivent figurer dans la case)
    var etus = _vsuEtudiants_(ss), idx = {};
    etus.forEach(function (e) {
      e.tNom = _dirTokens_(e.nom); e.tPre = _dirTokens_(e.prenom);
      if (!e.tNom.length || !e.tPre.length) return;
      e.tNom.forEach(function (w) { (idx[w] = idx[w] || []).push(e); });
    });
    var pris = {};
    cases.forEach(function (k) {
      var T = {}; _dirTokens_(k.texte).forEach(function (w) { T[w] = 1; });
      var vus = {}, cands = [];
      Object.keys(T).forEach(function (w) { (idx[w] || []).forEach(function (e) { if (!vus[e.id]) { vus[e.id] = 1; cands.push(e); } }); });
      var best = null, bs = -1;
      cands.forEach(function (e) {
        if (!e.tNom.every(function (w) { return T[w]; }) || !T[e.tPre[0]]) return;
        var s = e.tNom.length * 2 + e.tPre.filter(function (w) { return T[w]; }).length
          + (_vsuDiplome_(e.classe) === k.diplome ? 3 : 0) + (e.inscrit ? 2 : 0) + (!k.campus || e.campus === k.campus ? 1 : 0);
        if (s > bs) { bs = s; best = e; }
      });
      if (best) { k.etu = best; k.score = bs; if (!pris[best.id] || pris[best.id].score < bs) pris[best.id] = k; }
    });
    cases.forEach(function (k) { if (k.etu && pris[k.etu.id] !== k) { k.doublon = true; } });

    // 3. Double lecture
    var campusCouverts = {}; cases.forEach(function (k) { if (k.campus) campusCouverts[k.campus] = 1; });
    var nbCampus = Object.keys(campusCouverts).length;
    var crmInscrits = etus.filter(function (e) { return e.inscrit && (!nbCampus || !e.campus || campusCouverts[e.campus]); });
    var L1 = [], L2 = [];
    var fiche = function (e) { return e ? { id: e.id, nom: (e.prenom + " " + e.nom).trim(), classe: e.classe, typo: e.typo, statut: e.statut, campus: e.campus, alternant: _vsuAlternant_(e), entreprise: e.nomEnt } : null; };
    cases.forEach(function (k) {
      var e = k.etu && !k.doublon ? k.etu : null;
      var base = { dir: { nom: k.nom, texte: k.texte, classe: k.classeDir, campus: k.campus, onglet: k.onglet, ref: k.ref, couleur: k.couleur, vert: k.vert, absent: k.etat === "absent", autreCouleur: k.autreCouleur }, crm: fiche(e) };
      var v;
      if (k.doublon) v = "Doublon dans le fichier directeur";
      else if (k.etat === "absent") v = e && e.inscrit ? "Rouge chez le directeur, inscrit dans le CRM" : "Sorti des deux côtés";
      else if (!e) v = "Absent du CRM";
      else if (!e.inscrit) v = "Non inscrit dans le CRM";
      else v = _vsuDiplome_(e.classe) === k.diplome ? "Présent des deux côtés" : "Présent, diplôme différent";
      L1.push(Object.assign({ verdict: v }, base));
      if (k.etat !== "absent" && e && e.inscrit && !k.doublon) {
        var a = _vsuAlternant_(e);
        L2.push(Object.assign({ verdict: k.vert && a ? "Alternance des deux côtés" : k.vert ? "Vert chez le directeur, pas alternant dans le CRM" : a ? "Alternant dans le CRM, pas vert chez le directeur" : "Sans alternance des deux côtés" }, base));
      }
    });
    crmInscrits.forEach(function (e) {
      if (pris[e.id] && !pris[e.id].doublon) return;
      L1.push({ verdict: "Absent du fichier directeur", dir: null, crm: fiche(e) });
    });
    var nb = function (L, v) { return L.filter(function (x) { return x.verdict === v; }).length; };
    var dirPresents = cases.filter(function (k) { return k.etat !== "absent" && !k.doublon; });
    var stats = {
      dirInscrits: dirPresents.length, dirRouges: cases.filter(function (k) { return k.etat === "absent"; }).length,
      crmInscrits: crmInscrits.length, dirAlternants: dirPresents.filter(function (k) { return k.vert; }).length,
      crmAlternants: crmInscrits.filter(_vsuAlternant_).length,
      altAccord: nb(L2, "Alternance des deux côtés"), altDirSeul: nb(L2, "Vert chez le directeur, pas alternant dans le CRM"), altCrmSeul: nb(L2, "Alternant dans le CRM, pas vert chez le directeur")
    };
    // Effectifs par classe du directeur
    var parClasse = {};
    dirPresents.forEach(function (k) { var c = (k.campus ? k.campus.slice(0, 3).toUpperCase() + " · " : "") + k.classeDir; parClasse[c] = parClasse[c] || { classe: c, dir: 0, dirVert: 0, crm: 0, crmAlt: 0 }; parClasse[c].dir++; if (k.vert) parClasse[c].dirVert++;
      if (k.etu && k.etu.inscrit) { parClasse[c].crm++; if (_vsuAlternant_(k.etu)) parClasse[c].crmAlt++; } });
    return createJsonResponse({ success: true, onglets: tous, choisis: choix, annee: annee, stats: stats, lecture1: L1, lecture2: L2,
      parClasse: Object.keys(parClasse).sort().map(function (c) { return parClasse[c]; }),
      couleursAutres: couleursAutres, url: "https://docs.google.com/spreadsheets/d/" + sid + "/edit" });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// ═══════════════════════════ 3. INSCRIPTIONS AUX EXAMENS OFFICIELS ═══════════════════════════
// Par session (année scolaire), diplôme et classe : état « Inscrit » (bouton, saisi à la main) et
// état « Justificatif OK » (automatique dès qu'un justificatif est déposé). Épreuves choisies dans le
// référentiel des matières du diplôme (feuille Referentiel_matieres). Export Excel côté page.
var EXAM_HEADERS = ["Session", "ID étudiant", "Étudiant", "Diplôme", "Classe", "Inscrit", "Inscrit le", "Inscrit par", "Épreuves", "N° candidat", "Commentaire", "Mis à jour le", "Par"];
var EXAMDOC_HEADERS = ["ID", "Session", "ID étudiant", "Nom du fichier", "ID fichier", "URL", "Commentaire", "Déposé par", "Déposé le", "Archivé"];

function _exRows_(ss, session) {
  var ws = _vsxSheet_(ss, "EXAMENS_INSCRIPTIONS", EXAM_HEADERS), d = ws.getDataRange().getValues(), rows = {}, sessions = {};
  for (var i = 1; i < d.length; i++) {
    if (!d[i][0]) continue; sessions[String(d[i][0])] = 1;
    if (String(d[i][0]) === session) rows[String(d[i][1])] = { row: i + 1, v: d[i] };
  }
  return { ws: ws, rows: rows, sessions: Object.keys(sessions) };
}
function _exDocs_(ss, session) {
  var ws = ss.getSheetByName("EXAMENS_DOCS"), out = {};
  if (!ws) return out;
  var d = ws.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) {
    if (!d[i][0] || String(d[i][9]) === "Oui" || String(d[i][1]) !== session) continue;
    (out[String(d[i][2])] = out[String(d[i][2])] || []).push({ id: String(d[i][0]), nom: String(d[i][3] || ""), url: String(d[i][5] || ""), commentaire: String(d[i][6] || ""), par: String(d[i][7] || ""), le: _vsuIso_(d[i][8]) });
  }
  return out;
}

function handleGetExamens(p, ss) {
  if (!_vsxCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  try {
    var courante = _ckAnneeCourante_(), session = String(p.session || courante);
    var X = _exRows_(ss, session), docs = _exDocs_(ss, session), etus = _vsuEtudiants_(ss), byId = {};
    etus.forEach(function (e) { byId[e.id] = e; });
    var ids = {};
    if (session === courante) etus.forEach(function (e) { if (e.inscrit && _ckDiplome_(e.classe)) ids[e.id] = 1; });
    Object.keys(X.rows).forEach(function (id) { ids[id] = 1; });
    var liste = Object.keys(ids).map(function (id) {
      var e = byId[id] || {}, r = X.rows[id] ? X.rows[id].v : null, ep = [];
      if (r) try { ep = JSON.parse(String(r[8] || "[]")); } catch (eJ) { ep = []; }
      var ds = docs[id] || [];
      return { id: id, nom: e.id ? (e.prenom + " " + e.nom).trim() : String(r ? r[2] : id), classe: e.classe || String(r ? r[4] : ""),
        diplome: (r && r[3]) ? String(r[3]) : _vsuDiplome_(e.classe), campus: e.campus || "", mail: e.mail || "", typo: e.typo || "", sorti: !!e.id && !e.inscrit,
        inscrit: !!r && String(r[5]) === "Oui", inscritLe: r ? _vsuIso_(r[6]) : "", inscritPar: r ? String(r[7] || "") : "",
        epreuves: ep, numCandidat: r ? String(r[9] || "") : "", commentaire: r ? String(r[10] || "") : "", docs: ds, justifOk: ds.length > 0 };
    });
    var ref = {};
    try {
      var dR = getSheetSafe(ss, "Referentiel_matieres").getDataRange().getValues();
      for (var i = 1; i < dR.length; i++) { var dip = String(dR[i][0] || "").trim(), mat = String(dR[i][1] || "").trim(); if (dip && mat) (ref[dip] = ref[dip] || []).push(mat); }
    } catch (eR) {}
    var sessions = X.sessions.concat([courante]).filter(function (a, i, arr) { return arr.indexOf(a) === i; }).sort().reverse();
    return createJsonResponse({ success: true, session: session, sessions: sessions, etudiants: liste, epreuvesRef: ref });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

/** Met à jour l'inscription d'un ou plusieurs étudiants : inscrit, épreuves, n° candidat, commentaire. */
function handleSaveExamen(p, ss) {
  if (!_vsxCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var session = String(p.session || _ckAnneeCourante_()), X = _exRows_(ss, session), now = new Date(), par = String(p.nomAuteur || p.emailAuteur || "");
    var etus = {}; _vsuEtudiants_(ss).forEach(function (e) { etus[e.id] = e; });
    var ids = Array.isArray(p.etuIds) ? p.etuIds.map(String) : [String(p.etuId || "")];
    var ch = p.champs || {}, noms = [];
    ids.forEach(function (id) {
      if (!id) return;
      var e = etus[id] || {}, r = X.rows[id];
      var v = r ? r.v.slice(0, EXAM_HEADERS.length) : [session, id, ((e.prenom || "") + " " + (e.nom || "")).trim(), _vsuDiplome_(e.classe), e.classe || "", "", "", "", "[]", "", "", "", ""];
      while (v.length < EXAM_HEADERS.length) v.push("");
      if (ch.inscrit !== undefined) { v[5] = ch.inscrit ? "Oui" : ""; v[6] = ch.inscrit ? now : ""; v[7] = ch.inscrit ? par : ""; }
      if (ch.epreuves !== undefined) v[8] = JSON.stringify((Array.isArray(ch.epreuves) ? ch.epreuves : []).map(String));
      if (ch.numCandidat !== undefined) v[9] = String(ch.numCandidat || "").trim();
      if (ch.commentaire !== undefined) v[10] = String(ch.commentaire || "");
      v[11] = now; v[12] = par;
      if (r) X.ws.getRange(r.row, 1, 1, EXAM_HEADERS.length).setValues([v]);
      else { X.ws.appendRow(v); X.rows[id] = { row: X.ws.getLastRow(), v: v }; }
      X.ws.getRange(X.rows[id].row, 10).setNumberFormat("@").setValue(v[9]);
      noms.push(v[2]);
    });
    var quoi = ch.inscrit !== undefined ? (ch.inscrit ? "inscription enregistrée" : "inscription annulée") : ch.epreuves !== undefined ? "épreuves : " + (ch.epreuves || []).join(", ") : "fiche mise à jour";
    logAction(par, p.vieScRole || "", "Modification", "Vie scolaire — Examens officiels", ids.length === 1 ? ids[0] : "EXAMENS " + session,
      "Session " + session + " — " + quoi + " — " + (noms.length > 6 ? noms.length + " étudiants" : noms.join(", ")));
    return createJsonResponse({ success: true });
  } catch (err) { return createJsonResponse({ success: false, message: err.toString() }); }
  finally { try { lock.releaseLock(); } catch (eL) {} }
}

function _exDossier_(session, etudiant) {
  var props = PropertiesService.getScriptProperties(), rootId = props.getProperty("EXAMENS_FOLDER_ID"), root = null;
  if (rootId) try { root = DriveApp.getFolderById(rootId); } catch (e) { root = null; }
  if (!root) {
    var parent; try { parent = DriveApp.getFolderById(MEMOIRES_PARENT_FOLDER_ID); } catch (eP) { parent = DriveApp.getRootFolder(); }
    root = parent.createFolder("Inscriptions examens officiels");
    props.setProperty("EXAMENS_FOLDER_ID", root.getId());
  }
  var sub = function (f, nom) { var it = f.getFoldersByName(nom); return it.hasNext() ? it.next() : f.createFolder(nom); };
  return sub(sub(root, "Session " + session), String(etudiant || "Sans nom").replace(/[\\/:*?"<>|]/g, " ").trim());
}
/** Dépose un justificatif d'inscription : l'état « Justificatif OK » devient automatique. */
function handleUploadExamDoc(p, ss) {
  if (!_vsxCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    if (!p.etuId || !p.data) return createJsonResponse({ success: false, message: "Fichier ou étudiant manquant." });
    var session = String(p.session || _ckAnneeCourante_());
    var e = _vsuEtudiants_(ss).filter(function (x) { return x.id === String(p.etuId); })[0] || { prenom: "", nom: String(p.etuId) };
    var nomEtu = (e.prenom + " " + e.nom).trim();
    var orig = String(p.fileName || "justificatif"), ext = orig.indexOf(".") !== -1 ? orig.slice(orig.lastIndexOf(".")) : "";
    var nomFichier = nomEtu + " — Inscription examen " + session + " — " + orig.replace(/\.[^.]+$/, "") + ext;
    var file = _exDossier_(session, nomEtu).createFile(Utilities.newBlob(Utilities.base64Decode(String(p.data)), String(p.mimeType || "application/octet-stream"), nomFichier));
    var ws = _vsxSheet_(ss, "EXAMENS_DOCS", EXAMDOC_HEADERS), id = "EXD" + Date.now(), par = String(p.nomAuteur || p.emailAuteur || "");
    ws.appendRow([id, session, String(p.etuId), nomFichier, file.getId(), file.getUrl(), String(p.commentaire || ""), par, new Date(), ""]);
    logAction(par, p.vieScRole || "", "Création", "Vie scolaire — Examens officiels", String(p.etuId), "Justificatif d'inscription " + session + " déposé pour " + nomEtu + " : " + orig);
    return createJsonResponse({ success: true, doc: { id: id, nom: nomFichier, url: file.getUrl(), commentaire: String(p.commentaire || ""), par: par, le: _vsuIso_(new Date()) } });
  } catch (err) { return createJsonResponse({ success: false, message: "Dépôt impossible : " + err.toString() }); }
  finally { try { lock.releaseLock(); } catch (eL) {} }
}
function handleArchiveExamDoc(p, ss) {
  if (!_vsxCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  try {
    var ws = _vsxSheet_(ss, "EXAMENS_DOCS", EXAMDOC_HEADERS), d = ws.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) if (String(d[i][0]) === String(p.id)) {
      ws.getRange(i + 1, 10).setValue("Oui");
      logAction(p.nomAuteur || p.emailAuteur || "", p.vieScRole || "", "Suppression", "Vie scolaire — Examens officiels", String(d[i][2]), "Justificatif retiré : " + d[i][3]);
      return createJsonResponse({ success: true });
    }
    return createJsonResponse({ success: false, message: "Document introuvable." });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// ═══════════════════════════ 1 bis. RDV : PARAMÈTRES, ENVOI DU LIEN DE PRISE DE RDV, ANALYSE ═══════════════════════════
// RDV_RESPONSABLES : lien de calendrier (prise de RDV) de chaque responsable de visite.
// Modèle de mail (propriété RDV_MAIL_MODELE) avec variables {{…}}. Envoi : un mail par étudiant au
// maître d'apprentissage (col. M d'ETUDIANTS = contact tuteur ; sinon contacts « MA = Oui » de
// l'entreprise ; sinon 1er contact avec mail), ou à tous les contacts de l'entreprise ; le responsable
// est en copie et en adresse de réponse. Aperçu obligatoire avant envoi (previewMailsRdv).
var RDV_RESP_HEADERS = ["Email", "Nom", "Lien calendrier", "Mis à jour le", "Par"];
var RDV_ENVOI_HEADERS = ["Date", "Année scolaire", "ID étudiant", "Étudiant", "Destinataires", "Responsable", "Objet", "Par"];
var RDV_MODELE_DEFAUT = {
  objet: "Visite en entreprise — {{etudiant}} — prise de rendez-vous",
  corps: "Bonjour {{contact}},\n\nDans le cadre du suivi en alternance de notre étudiant(e) {{etudiant}} ({{classe}}), {{responsable}} de l'école HECG souhaite vous rencontrer au sein de {{entreprise}} pour faire le point sur sa formation.\n\nMerci de choisir le créneau qui vous convient en cliquant sur ce lien :\n{{lien_rdv}}\n\nNous vous remercions pour votre accompagnement.\n\nBien cordialement,\n{{responsable}}\nHECG"
};
var RDV_VARIABLES = ["contact", "etudiant", "prenom_etudiant", "classe", "entreprise", "responsable", "lien_rdv", "annee"];
var RDV_SS_ID = "1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ";

function _rdvResponsables_(ss) {
  var d = _vsxSheet_(ss, "RDV_RESPONSABLES", RDV_RESP_HEADERS).getDataRange().getValues(), out = [];
  for (var i = 1; i < d.length; i++) if (d[i][0] || d[i][1]) out.push({ email: String(d[i][0] || "").trim().toLowerCase(), nom: String(d[i][1] || "").trim(), lien: String(d[i][2] || "").trim() });
  return out;
}
function _rdvModele_() {
  try { var m = JSON.parse(PropertiesService.getScriptProperties().getProperty("RDV_MAIL_MODELE") || "null"); if (m && m.objet && m.corps) return m; } catch (e) {}
  return RDV_MODELE_DEFAUT;
}
function _rdvEnvois_(ss, annee) {
  var ws = ss.getSheetByName("RDV_ENVOIS"), out = {};
  if (!ws) return out;
  var d = ws.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) if (String(d[i][1]) === annee && d[i][2]) {
    var k = String(d[i][2]); out[k] = out[k] || [];
    out[k].push({ le: _vsuIso_(d[i][0]), dest: String(d[i][4] || ""), resp: String(d[i][5] || ""), par: String(d[i][7] || "") });
  }
  return out;
}

function handleGetRdvParams(p, ss) {
  if (!_vsxCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  try {
    var equipe = _vsxRef_(ss).personnel.filter(function (x) { return x.email; });
    var liens = {}; _rdvResponsables_(ss).forEach(function (r) { liens[r.email || _vsuNorm_(r.nom)] = r.lien; });
    var quota = null; try { quota = MailApp.getRemainingDailyQuota(); } catch (eQ) {}
    return createJsonResponse({ success: true, modele: _rdvModele_(), modeleDefaut: RDV_MODELE_DEFAUT, variables: RDV_VARIABLES, quota: quota,
      responsables: equipe.map(function (x) { return { email: x.email, nom: x.nom, role: x.role, lien: liens[x.email] || liens[_vsuNorm_(x.nom)] || "" }; }) });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}
function handleSaveRdvParams(p, ss) {
  if (!_vsuAdmin_(p)) return createJsonResponse({ success: false, message: "Réservé aux profils Super-admin et Administratif." });
  try {
    var par = String(p.nomAuteur || p.emailAuteur || ""), now = new Date(), modifs = [];
    if (Array.isArray(p.responsables)) {
      var ws = _vsxSheet_(ss, "RDV_RESPONSABLES", RDV_RESP_HEADERS);
      var avant = {}; _rdvResponsables_(ss).forEach(function (r) { avant[r.email] = r.lien; });
      var rows = p.responsables.filter(function (r) { return r && r.email; }).map(function (r) {
        var lien = String(r.lien || "").trim();
        if (lien && !/^https?:\/\//i.test(lien)) lien = "https://" + lien;
        if ((avant[String(r.email).toLowerCase()] || "") !== lien) modifs.push(r.nom + (lien ? " : lien mis à jour" : " : lien retiré"));
        return [String(r.email).trim().toLowerCase(), String(r.nom || ""), lien, now, par];
      });
      if (ws.getLastRow() > 1) ws.getRange(2, 1, ws.getLastRow() - 1, RDV_RESP_HEADERS.length).clearContent();
      if (rows.length) ws.getRange(2, 1, rows.length, RDV_RESP_HEADERS.length).setValues(rows);
    }
    if (p.modele && p.modele.objet && p.modele.corps) {
      var avantM = _rdvModele_();
      if (avantM.objet !== String(p.modele.objet) || avantM.corps !== String(p.modele.corps)) {
        PropertiesService.getScriptProperties().setProperty("RDV_MAIL_MODELE", JSON.stringify({ objet: String(p.modele.objet), corps: String(p.modele.corps) }));
        modifs.push("modèle de mail modifié");
      }
    }
    if (modifs.length) logAction(par, p.vieScRole || "", "Modification", "Vie scolaire — Suivi des RDV", "RDV_RESPONSABLES", "Paramètres RDV : " + modifs.join(" ; "));
    return createJsonResponse({ success: true });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// Prépare les mails (aperçu et envoi utilisent exactement le même calcul)
function _rdvPreparerMails_(ss, annee, etuIds, cible) {
  var etus = {}; _vsuEtudiants_(ss).forEach(function (e) { etus[e.id] = e; });
  var dE = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues(), tuteur = {};
  for (var i = 1; i < dE.length; i++) if (dE[i][0]) tuteur[String(dE[i][0]).trim()] = String(dE[i][12] || "").trim();
  var dP = getSheetSafe(ss, "PARTENARIAT").getDataRange().getValues(), contacts = {}, parIdContact = {};
  for (var j = 1; j < dP.length; j++) {
    var idEnt = String(dP[j][1] || "").trim().toUpperCase(); if (!idEnt) continue;
    var c = { id: String(dP[j][0] || "").trim(), nom: (String(dP[j][4] || "") + " " + String(dP[j][3] || "")).trim(), email: String(dP[j][6] || "").trim() || String(dP[j][8] || "").trim(),
      ma: /^oui$/i.test(String(dP[j][14] || "").trim()), entNom: String(dP[j][2] || "").trim() };
    (contacts[idEnt] = contacts[idEnt] || []).push(c);
    if (c.id) parIdContact[c.id.toUpperCase()] = c;
  }
  var refs = _rdvReferents_(ss, annee), liens = {};
  _rdvResponsables_(ss).forEach(function (r) { if (r.email) liens[r.email] = r.lien; if (r.nom) liens["n:" + _vsuNorm_(r.nom)] = r.lien; });
  var M = _rdvModele_(), out = [];
  var valide = function (m) { return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(m); };
  etuIds.forEach(function (id) {
    var e = etus[id]; if (!e) return;
    var it = { etuId: id, etudiant: (e.prenom + " " + e.nom).trim(), classe: e.classe, entreprise: e.nomEnt, destinataires: [], responsable: "", responsableEmail: "", bloquant: "" };
    var ref = refs[id];
    if (!ref || !ref.nom) { it.bloquant = "Aucun responsable visite"; out.push(it); return; }
    it.responsable = ref.nom; it.responsableEmail = ref.email;
    var lien = (ref.email && liens[String(ref.email).toLowerCase()]) || liens["n:" + _vsuNorm_(ref.nom)] || "";
    if (!lien) { it.bloquant = "Pas de lien de calendrier pour " + ref.nom; out.push(it); return; }
    if (!e.idEnt) { it.bloquant = "Aucune entreprise dans le CRM"; out.push(it); return; }
    var lst = contacts[e.idEnt.toUpperCase()] || [];
    if (lst[0]) it.entreprise = lst[0].entNom || it.entreprise;
    var t = tuteur[id] ? parIdContact[tuteur[id].toUpperCase()] : null, dest = [];
    if (cible === "tous") dest = lst.filter(function (c) { return valide(c.email); });
    else if (t && valide(t.email)) dest = [t];
    else { dest = lst.filter(function (c) { return c.ma && valide(c.email); }); if (!dest.length) dest = lst.filter(function (c) { return valide(c.email); }).slice(0, 1); }
    var vus = {}; dest = dest.filter(function (c) { var k = c.email.toLowerCase(); if (vus[k]) return false; vus[k] = 1; return true; });
    if (!dest.length) { it.bloquant = "Aucun contact avec e-mail dans l'entreprise"; out.push(it); return; }
    it.destinataires = dest.map(function (c) { return { nom: c.nom, email: c.email, ma: !!(c.ma || (t && c.id === t.id)) }; });
    var vars = { contact: dest.length === 1 ? dest[0].nom : "Madame, Monsieur", etudiant: it.etudiant, prenom_etudiant: e.prenom, classe: e.classe, entreprise: it.entreprise,
      responsable: ref.nom, lien_rdv: lien, annee: annee };
    var rep = function (s) { return String(s).replace(/\{\{\s*([a-z_]+)\s*\}\}/gi, function (m, k) { return vars[k.toLowerCase()] !== undefined ? vars[k.toLowerCase()] : m; }); };
    it.objet = rep(M.objet); it.corps = rep(M.corps); it.lien = lien;
    out.push(it);
  });
  return out;
}
function _rdvHtml_(txt) {
  var h = String(txt).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  h = h.replace(/(https?:\/\/[^\s<]+)/g, '<a href="$1" style="color:#1A4E8A;font-weight:bold">$1</a>');
  return '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;line-height:1.55;color:#1f2937">' + h.replace(/\n/g, "<br>") + "</div>";
}
function handlePreviewMailsRdv(p, ss) {
  if (!_vsuAdmin_(p)) return createJsonResponse({ success: false, message: "Réservé aux profils Super-admin et Administratif." });
  try {
    var annee = String(p.annee || _ckAnneeCourante_());
    var items = _rdvPreparerMails_(ss, annee, (p.etuIds || []).map(String), String(p.cible || "ma"));
    var quota = null; try { quota = MailApp.getRemainingDailyQuota(); } catch (eQ) {}
    return createJsonResponse({ success: true, items: items, quota: quota });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}
function handleSendMailsRdv(p, ss) {
  if (!_vsuAdmin_(p)) return createJsonResponse({ success: false, message: "Réservé aux profils Super-admin et Administratif." });
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var annee = String(p.annee || _ckAnneeCourante_()), par = String(p.nomAuteur || p.emailAuteur || "");
    var items = _rdvPreparerMails_(ss, annee, (p.etuIds || []).map(String), String(p.cible || "ma")).filter(function (x) { return !x.bloquant; });
    if (!items.length) return createJsonResponse({ success: false, message: "Aucun mail envoyable dans la sélection." });
    var quota = MailApp.getRemainingDailyQuota();
    var nbDest = items.reduce(function (s, x) { return s + x.destinataires.length + (x.responsableEmail ? 1 : 0); }, 0);
    if (nbDest > quota) return createJsonResponse({ success: false, message: "Quota d'envoi du jour insuffisant (" + quota + " restants pour " + nbDest + " destinataires). Réduisez la sélection." });
    var ws = _vsxSheet_(ss, "RDV_ENVOIS", RDV_ENVOI_HEADERS), lignes = [], echecs = [], now = new Date();
    items.forEach(function (x) {
      try {
        var opt = { to: x.destinataires.map(function (c) { return c.email; }).join(","), subject: x.objet, body: x.corps, htmlBody: _rdvHtml_(x.corps), name: "HECG — " + x.responsable };
        if (x.responsableEmail) { opt.cc = x.responsableEmail; opt.replyTo = x.responsableEmail; }
        MailApp.sendEmail(opt);
        lignes.push([now, annee, x.etuId, x.etudiant, opt.to, x.responsable, x.objet, par]);
      } catch (eS) { echecs.push(x.etudiant + " : " + eS.message); }
    });
    if (lignes.length) ws.getRange(ws.getLastRow() + 1, 1, lignes.length, RDV_ENVOI_HEADERS.length).setValues(lignes);
    logAction(par, p.vieScRole || "", "Action", "Vie scolaire — Suivi des RDV", "RDV_ENVOIS",
      "Envoi du lien de prise de RDV (" + (p.cible === "tous" ? "tous les contacts" : "maîtres d'apprentissage") + ") : " + lignes.length + " mail(s) envoyé(s)" + (echecs.length ? ", " + echecs.length + " échec(s)" : "")
      + " — " + lignes.slice(0, 30).map(function (l) { return l[3] + " → " + l[4]; }).join(" ; "));
    return createJsonResponse({ success: true, envoyes: lignes.length, echecs: echecs });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
  finally { try { lock.releaseLock(); } catch (eL) {} }
}

// ── Mouvements d'alternance : relevé quotidien de l'entreprise de chaque étudiant inscrit ──
// ALTERNANCE_ETAT = dernier état connu ; ALTERNANCE_MOUVEMENTS = gains / pertes / changements datés.
// Le 1er relevé reprend aussi les départs déjà tracés dans ANCIENS_ALTERNANTS (changements, pertes).
var ALT_ETAT_HEADERS = ["ID étudiant", "ID entreprise", "Entreprise", "Relevé le"];
var ALT_MVT_HEADERS = ["Date", "ID étudiant", "Étudiant", "Classe", "Mouvement", "Ancienne entreprise", "Nouvelle entreprise", "Source"];
function _altReleve_(ss) {
  var wsE = _vsxSheet_(ss, "ALTERNANCE_ETAT", ALT_ETAT_HEADERS), wsM = _vsxSheet_(ss, "ALTERNANCE_MOUVEMENTS", ALT_MVT_HEADERS);
  var etus = _vsuEtudiants_(ss).filter(function (e) { return e.inscrit; }), now = new Date();
  var dEt = wsE.getDataRange().getValues(), avant = {}, premier = dEt.length < 2;
  for (var i = 1; i < dEt.length; i++) if (dEt[i][0]) avant[String(dEt[i][0])] = { id: String(dEt[i][1] || ""), nom: String(dEt[i][2] || "") };
  var mvts = [];
  if (premier) {
    var wsA = ss.getSheetByName("ANCIENS_ALTERNANTS");
    if (wsA && wsA.getLastRow() > 1) {
      var byId = {}; etus.forEach(function (e) { byId[e.id] = e; });
      wsA.getDataRange().getValues().slice(1).forEach(function (r) {
        if (!(r[5] instanceof Date) || String(r[6]) === "Erreur") return;
        var e = byId[String(r[2]).trim()]; if (!e) return;
        var type = e.idEnt && e.idEnt.toUpperCase() !== String(r[1]).trim().toUpperCase() ? "Changement" : !e.idEnt ? "Perte" : "";
        if (type) mvts.push([r[5], e.id, (e.prenom + " " + e.nom).trim(), e.classe, type, String(r[7] || ""), type === "Changement" ? e.nomEnt : "", "ANCIENS_ALTERNANTS"]);
      });
    }
  } else {
    etus.forEach(function (e) {
      var a = avant[e.id]; if (!a) return;   // nouvel inscrit : pas de mouvement
      var x = a.id.toUpperCase(), y = e.idEnt.toUpperCase();
      if (x === y) return;
      var type = !x ? "Gain" : !y ? "Perte" : "Changement";
      mvts.push([now, e.id, (e.prenom + " " + e.nom).trim(), e.classe, type, a.nom, e.nomEnt, "Relevé quotidien"]);
    });
  }
  if (mvts.length) wsM.getRange(wsM.getLastRow() + 1, 1, mvts.length, ALT_MVT_HEADERS.length).setValues(mvts);
  var rows = etus.map(function (e) { return [e.id, e.idEnt, e.nomEnt, now]; });
  if (wsE.getLastRow() > 1) wsE.getRange(2, 1, wsE.getLastRow() - 1, ALT_ETAT_HEADERS.length).clearContent();
  if (rows.length) wsE.getRange(2, 1, rows.length, ALT_ETAT_HEADERS.length).setValues(rows);
  PropertiesService.getScriptProperties().setProperty("ALT_RELEVE_JOUR", Utilities.formatDate(now, "Europe/Paris", "yyyy-MM-dd"));
  if (premier) PropertiesService.getScriptProperties().setProperty("ALT_SUIVI_DEPUIS", Utilities.formatDate(now, "Europe/Paris", "yyyy-MM-dd"));
  return mvts.length;
}
// Appelé par auditAdressesTick (adresses.gs) : un relevé par jour, à partir de 6 h
function _altAuto_() {
  var now = new Date(), jour = Utilities.formatDate(now, "Europe/Paris", "yyyy-MM-dd");
  if (+Utilities.formatDate(now, "Europe/Paris", "H") < 6) return;
  if (PropertiesService.getScriptProperties().getProperty("ALT_RELEVE_JOUR") === jour) return;
  var lock = LockService.getScriptLock(); if (!lock.tryLock(5000)) return;
  try { _altReleve_(SpreadsheetApp.openById(RDV_SS_ID)); } finally { lock.releaseLock(); }
}

function handleGetRdvAnalyse(p, ss) {
  if (!_vsxCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  try {
    var annee = String(p.annee || _ckAnneeCourante_());
    var props = PropertiesService.getScriptProperties();
    if (!props.getProperty("ALT_SUIVI_DEPUIS")) { var lock = LockService.getScriptLock(); if (lock.tryLock(10000)) try { _altReleve_(ss); } finally { lock.releaseLock(); } }
    var y0 = +annee.split("-")[0], debut = new Date(y0, 8, 1), fin = new Date(y0 + 1, 7, 31, 23, 59);
    var mv = [], wsM = ss.getSheetByName("ALTERNANCE_MOUVEMENTS");
    if (wsM && wsM.getLastRow() > 1) wsM.getDataRange().getValues().slice(1).forEach(function (r) {
      if (r[0] instanceof Date && r[0] >= debut && r[0] <= fin) mv.push({ date: _vsuIso_(r[0]), etuId: String(r[1]), etudiant: String(r[2]), classe: String(r[3]), type: String(r[4]), avant: String(r[5]), apres: String(r[6]), source: String(r[7]) });
    });
    return createJsonResponse({ success: true, annee: annee, mouvements: mv, suiviDepuis: props.getProperty("ALT_SUIVI_DEPUIS") || "", dernierReleve: props.getProperty("ALT_RELEVE_JOUR") || "" });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}
