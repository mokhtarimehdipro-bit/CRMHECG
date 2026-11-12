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
  Object.keys(g).forEach(function (k) { var a = _adrReference_(g[k].lignes); out[k] = { id: g[k].id, nom: g[k].nom, lieu: [a.adresse, a.complement, a.cp, a.ville].filter(Boolean).join(" "), cp: String(a.cp || ""), ville: String(a.ville || "") }; });
  return out;
}

// ═══════════════════════════ 1. SUIVI DES RDV EN ENTREPRISE ═══════════════════════════
// RDV_VISITES : une ligne par rendez-vous (on peut visiter plusieurs fois un étudiant dans l'année).
// L'entreprise et le lieu sont FIGÉS à la création : l'historique dit « tel jour, untel a visité tel
// étudiant dans telle structure », même si l'étudiant change d'entreprise ensuite.
// « Date RDV pris » = date initiale ; « Date RDV » = date en vigueur (après report).
var RDV_HEADERS = ["ID", "Année scolaire", "ID étudiant", "Étudiant", "Classe", "ID entreprise", "Entreprise", "Lieu", "Visiteur", "Email visiteur",
  "Date RDV pris", "Date RDV", "Statut", "Historique", "Commentaire", "Source", "Créé le", "Créé par", "Mis à jour le", "Archivé", "Précision"];
var RDV_REF_HEADERS = ["Année scolaire", "ID étudiant", "Responsable visite", "Email", "Source", "Mis à jour le", "Par", "Précision", "Attribué le"];
var RDV_STATUTS = ["À planifier", "Planifié", "Honoré", "Reporté", "Annulé", "Autre"];

function _rdvSheet_(ss) { return _vsuEntete_(_vsxSheet_(ss, "RDV_VISITES", RDV_HEADERS), RDV_HEADERS); }
function _rdvLire_(ss) {
  var d = _rdvSheet_(ss).getDataRange().getValues(), out = [];
  for (var i = 1; i < d.length; i++) {
    var r = d[i]; if (!r[0] || String(r[19]) === "Oui") continue;
    var h = []; try { h = JSON.parse(String(r[13] || "[]")); } catch (e) { h = []; }
    out.push({ id: String(r[0]), annee: String(r[1]), etuId: String(r[2]), etudiant: String(r[3]), classe: String(r[4]), entId: String(r[5]), entreprise: String(r[6]),
      lieu: String(r[7]), visiteur: String(r[8]), visiteurEmail: String(r[9]), datePris: _vsuIso_(r[10]), date: _vsuIso_(r[11]), statut: String(r[12] || "Planifié"),
      historique: h, commentaire: String(r[14] || ""), source: String(r[15] || ""), creeLe: _vsuIso_(r[16]), creePar: String(r[17] || ""), precision: String(r[20] || ""), row: i + 1 });
  }
  return out;
}
function _rdvReferents_(ss, annee) {
  var d = _vsuEntete_(_vsxSheet_(ss, "RDV_REFERENTS", RDV_REF_HEADERS), RDV_REF_HEADERS).getDataRange().getValues(), out = {};
  for (var i = 1; i < d.length; i++) if (String(d[i][0]) === annee && d[i][1]) out[String(d[i][1])] = { nom: String(d[i][2] || ""), email: String(d[i][3] || ""), source: String(d[i][4] || ""), precision: String(d[i][7] || ""), attribueLe: _vsuIso_(d[i][8] || d[i][5]), row: i + 1 };
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
        alternant: _vsuAlternant_(e), statut: e.statut, entId: ent ? ent.id : "", entreprise: ent ? ent.nom : e.nomEnt, lieu: ent ? ent.lieu : "", cp: ent ? ent.cp : "", ville: ent ? ent.ville : "",
        referent: refs[e.id] ? refs[e.id].nom : "", referentEmail: refs[e.id] ? refs[e.id].email : "", referentSource: refs[e.id] ? refs[e.id].source : "",
        referentPrecision: refs[e.id] ? refs[e.id].precision : "", referentDepuis: refs[e.id] && refs[e.id].nom ? refs[e.id].attribueLe : "" });
    });
    var drap = _rdvDrapeaux_(ss, annee);
    try { _rdvAutoDrapeaux_(ss, annee, liste, visites, drap); } catch (eD) { console.error("drapeaux auto : " + eD); }
    liste.forEach(function (x) { var dr = drap[x.id]; x.drapeau = dr && dr.niveau ? dr.niveau : ""; x.drapeauCom = dr ? dr.commentaire : ""; x.drapeauLe = dr ? dr.poseLe : ""; x.drapeauPar = dr ? dr.par : ""; x.drapeauLeve = dr && dr.leveLe ? dr.leveLe + " par " + dr.levePar : ""; });
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
        while (row.length < RDV_HEADERS.length) row.push("");
        if (v.visiteur !== undefined) { var cvv = _rdvCanon_(v.visiteur, _rdvCtx_(ss)); row[8] = cvv.nom; row[9] = String(v.visiteurEmail || "").trim() || cvv.email; if (v.precision === undefined && cvv.precision) row[20] = _rdvJoindre_(row[20], cvv.precision); }
        if (v.precision !== undefined) row[20] = String(v.precision || "").trim();
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
    var cv = _rdvCanon_(v.visiteur || "", _rdvCtx_(ss));
    ws.appendRow([id, annee, e.id, (e.prenom + " " + e.nom).trim(), e.classe, ent ? ent.id : "", ent ? ent.nom : (String(v.entreprise || "") || e.nomEnt), ent ? ent.lieu : String(v.lieu || ""),
      cv.nom, String(v.visiteurEmail || "").trim() || cv.email, "", "", date ? "Planifié" : "À planifier", JSON.stringify(hist), String(v.commentaire || ""), "Saisie", now, par, now, "",
      v.precision !== undefined ? String(v.precision || "").trim() : cv.precision]);
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
    var c = _rdvCanon_(p.nom || "", _rdvCtx_(ss)), prec = p.precision !== undefined ? String(p.precision || "").trim() : c.precision;
    var n = _rdvEcrireReferents_(ss, annee, ids.map(function (id) { return { etuId: id, nom: c.nom, email: String(p.email || "").trim() || c.email, precision: prec }; }), "Saisie", String(p.nomAuteur || p.emailAuteur || ""));
    logAction(p.nomAuteur || p.emailAuteur || "", p.vieScRole || "", "Modification", "Vie scolaire — Suivi des RDV", ids.length === 1 ? ids[0] : "RDV_REFERENTS",
      "Responsable visite " + annee + " : " + (p.nom || "(retiré)") + " pour " + ids.length + " étudiant(s)");
    return createJsonResponse({ success: true, n: n });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}
function _rdvEcrireReferents_(ss, annee, items, source, par) {
  var refs = _rdvReferents_(ss, annee), ws = _vsxSheet_(ss, "RDV_REFERENTS", RDV_REF_HEADERS), now = new Date(), n = 0, ajout = [];
  items.forEach(function (it) {
    if (!it.etuId) return;
    var r = refs[it.etuId];
    if (r) {
      if (source === "Charlemagne" && r.source === "Saisie") return;   // une saisie manuelle n'est jamais écrasée par l'import
      var nouveau = _vsuNorm_(it.nom) !== _vsuNorm_(r.nom);   // date d'attribution = changement de responsable
      ws.getRange(r.row, 3, 1, 7).setValues([[it.nom, it.email, source, now, par, it.precision !== undefined ? it.precision : r.precision, nouveau || !r.attribueLe ? (it.nom ? now : "") : new Date(r.attribueLe + "T12:00:00")]]); n++;
    } else if (it.nom) { ajout.push([annee, it.etuId, it.nom, it.email, source, now, par, it.precision || "", now]); n++; }
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
  var ctxN = _rdvCtx_(ss);
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
    var cn = _rdvCanon_(noms[0] || "", ctxN), nom = cn.nom;
    if (nom) { refs.push({ etuId: e.id, nom: nom, email: cn.email || mailDe(nom), precision: cn.precision }); stats.responsables++; }
    var ent = e.idEnt ? ents[e.idEnt.toUpperCase()] : null;
    dates.filter(function (x, i, a) { return a.indexOf(x) === i; }).forEach(function (dt) {
      if (existants[e.id + "|" + dt]) return;
      existants[e.id + "|" + dt] = 1;
      var st = dt <= auj ? "Honoré" : "Planifié";
      nouvelles.push(["RDV" + now.getTime() + "C" + nouvelles.length, _vsuAnneeDe_(dt), e.id, (e.prenom + " " + e.nom).trim(), e.classe, ent ? ent.id : "", ent ? ent.nom : e.nomEnt, ent ? ent.lieu : "",
        nom, cn.email || mailDe(nom), dt, dt, st, JSON.stringify([{ le: Utilities.formatDate(now, "Europe/Paris", "yyyy-MM-dd HH:mm"), statut: st, date: dt, par: "Import Charlemagne", commentaire: vis.map(function (x) { return x.h + " : " + x.v; }).join(" | ") }]),
        "Repris de Charlemagne : " + vis.map(function (x) { return x.h + " = " + x.v; }).join(" ; "), "Charlemagne", now, par, now, "", cn.precision]);
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
var RDV_MODELE_DEFAUT = {   // impersonnel (12/11/2026) : « Bonjour », aucun nom de visiteur
  objet: "Visite en entreprise — {{etudiant}} — prise de rendez-vous",
  corps: "Bonjour,\n\nDans le cadre du suivi de sa formation en alternance, {{etudiant}} ({{classe}}) doit recevoir la visite d'un membre de l'équipe pédagogique de l'école HECG au sein de {{entreprise}}.\n\nMerci de choisir le créneau qui vous convient en cliquant sur ce lien :\n{{lien_rdv}}\n\nNous vous remercions pour votre accompagnement.\n\nBien cordialement,\nL'équipe pédagogique HECG"
};
var RDV_EXPEDITEUR = "HECG — Équipe pédagogique";
var RDV_VARIABLES = ["contact", "etudiant", "prenom_etudiant", "classe", "entreprise", "responsable", "lien_rdv", "annee"];
var RDV_SS_ID = "1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ";
var RDV_CC_POLE = "polepedagogique@hecg.fr";   // en copie de TOUS les envois du lien de RDV (demande direction 12/11/2026)

function _rdvResponsables_(ss) {
  var d = _vsxSheet_(ss, "RDV_RESPONSABLES", RDV_RESP_HEADERS).getDataRange().getValues(), out = [];
  for (var i = 1; i < d.length; i++) if (d[i][0] || d[i][1]) out.push({ email: String(d[i][0] || "").trim().toLowerCase(), nom: String(d[i][1] || "").trim(), lien: String(d[i][2] || "").trim() });
  return out;
}
function _rdvModele_() {
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty("RDV_MAIL_IMPERSONNEL")) {   // migration unique : ancien modèle nominatif → modèle impersonnel
    try { var a = JSON.parse(props.getProperty("RDV_MAIL_MODELE") || "null"); if (a && /\{\{\s*(contact|responsable)\s*\}\}/i.test(a.objet + a.corps)) props.deleteProperty("RDV_MAIL_MODELE"); } catch (e0) {}
    props.setProperty("RDV_MAIL_IMPERSONNEL", "1");
  }
  try { var m = JSON.parse(props.getProperty("RDV_MAIL_MODELE") || "null"); if (m && m.objet && m.corps) return m; } catch (e) {}
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
    it.cc = [ref.email, RDV_CC_POLE].filter(function (m, k, a) { return m && a.indexOf(m) === k && !dest.some(function (c) { return c.email.toLowerCase() === String(m).toLowerCase(); }); });
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
    if (p.test) {   // mail de test : le 1er mail de la sélection, envoyé uniquement au pôle pédagogique
      var t = items[0], dest = "polepedagogique@hecg.fr";
      var avert = "[MAIL DE TEST — destinataire(s) réel(s) : " + t.destinataires.map(function (c) { return c.nom + " <" + c.email + ">"; }).join(", ") + (t.responsableEmail ? " ; copie : " + t.responsableEmail : "") + "]\n\n";
      MailApp.sendEmail({ to: dest, subject: "[TEST] " + t.objet, body: avert + t.corps, htmlBody: _rdvHtml_(avert + t.corps), name: RDV_EXPEDITEUR });
      logAction(par, p.vieScRole || "", "Action", "Vie scolaire — Suivi des RDV", "RDV_ENVOIS", "Mail de test du lien de RDV envoyé à " + dest + " (exemple : " + t.etudiant + ", responsable " + t.responsable + ")");
      return createJsonResponse({ success: true, test: true, a: dest, exemple: t.etudiant });
    }
    var quota = MailApp.getRemainingDailyQuota();
    var nbDest = items.reduce(function (s, x) { return s + x.destinataires.length + (x.cc || []).length; }, 0);
    if (nbDest > quota) return createJsonResponse({ success: false, message: "Quota d'envoi du jour insuffisant (" + quota + " restants pour " + nbDest + " destinataires). Réduisez la sélection." });
    var ws = _vsxSheet_(ss, "RDV_ENVOIS", RDV_ENVOI_HEADERS), lignes = [], echecs = [], now = new Date();
    items.forEach(function (x) {
      try {
        var opt = { to: x.destinataires.map(function (c) { return c.email; }).join(","), subject: x.objet, body: x.corps, htmlBody: _rdvHtml_(x.corps), name: RDV_EXPEDITEUR };
        if ((x.cc || []).length) opt.cc = x.cc.join(",");
        if (x.responsableEmail) opt.replyTo = x.responsableEmail;
        MailApp.sendEmail(opt);
        lignes.push([now, annee, x.etuId, x.etudiant, opt.to + (opt.cc ? " (cc " + opt.cc + ")" : ""), x.responsable, x.objet, par]);
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

// ── Harmonisation des noms de visiteurs (« nico », « Nico », « Nico à sa demande » → un seul nom) ──
// Le texte libre (« à sa demande », parenthèses, « - … ») part dans la colonne « Précision ».
// Reconnaissance : alias mémorisés (propriété RDV_ALIAS : nom normalisé → e-mail), puis membre de
// l'équipe dont chaque mot du nom saisi commence un mot du nom (« nico » → « Nicolas DURAND »), sinon
// le nom est simplement mis en forme (« nico » → « Nico ») pour regrouper les variantes.
function _vsuEntete_(ws, headers) {
  var n = headers.length;
  if (ws.getLastColumn() < n || String(ws.getRange(1, n).getValue()) !== headers[n - 1]) {
    if (ws.getMaxColumns() < n) ws.insertColumnsAfter(ws.getMaxColumns(), n - ws.getMaxColumns());
    ws.getRange(1, 1, 1, n).setValues([headers]).setFontWeight("bold").setBackground("#1A4E8A").setFontColor("#ffffff");
  }
  return ws;
}
function _rdvCtx_(ss) {
  var alias = {}; try { alias = JSON.parse(PropertiesService.getScriptProperties().getProperty("RDV_ALIAS") || "{}"); } catch (e) { alias = {}; }
  return { equipe: _vsxRef_(ss).personnel.filter(function (x) { return x.email; }), alias: alias };
}
function _rdvSplit_(raw) {
  var s = String(raw == null ? "" : raw).replace(/\s+/g, " ").trim(), prec = [];
  s = s.replace(/\(([^)]*)\)/g, function (m, x) { if (x.trim()) prec.push(x.trim()); return " "; });
  var m = s.match(/\s*(?:[-–:,;]\s*)?(?:^|\s)(?:à|a)\s+(?:sa|la|leur|ma|notre|votre)\s+demande\b.*$/i);
  if (m) { prec.unshift(m[0].replace(/^[\s\-–:,;]+/, "").trim()); s = s.slice(0, m.index); }
  var m2 = s.match(/\s+[-–:]\s+(.*)$/);
  if (m2) { prec.push(m2[1].trim()); s = s.slice(0, m2.index); }
  return { nom: s.replace(/\s+/g, " ").trim(), precision: prec.filter(Boolean).join(" ; ") };
}
function _rdvCasse_(s) {
  return String(s).toLowerCase().split(" ").map(function (w) { return w.split("-").map(function (x) { return x ? x.charAt(0).toUpperCase() + x.slice(1) : x; }).join("-"); }).join(" ");
}
function _rdvCanon_(raw, ctx) {
  var sp = _rdvSplit_(raw);
  if (!sp.nom) return { nom: "", email: "", precision: sp.precision, reconnu: false };
  var k = _vsuNorm_(sp.nom);
  var parMail = function (m) { m = String(m || "").toLowerCase(); return ctx.equipe.filter(function (q) { return q.email.toLowerCase() === m; })[0]; };
  var a = ctx.alias[k] && parMail(ctx.alias[k]);
  if (a) return { nom: a.nom, email: a.email, precision: sp.precision, reconnu: true };
  if (sp.nom.indexOf("@") > 0) { var pm = parMail(sp.nom); if (pm) return { nom: pm.nom, email: pm.email, precision: sp.precision, reconnu: true }; }
  var toks = k.split(" ").filter(function (w) { return w.length >= 2; });
  var c = toks.length ? ctx.equipe.filter(function (q) {
    var qt = _vsuNorm_(q.nom).split(" ");
    return toks.every(function (w) { return qt.some(function (x) { return x.indexOf(w) === 0; }); });
  }) : [];
  if (c.length === 1) return { nom: c[0].nom, email: c[0].email, precision: sp.precision, reconnu: true };
  return { nom: _rdvCasse_(sp.nom), email: "", precision: sp.precision, reconnu: false, ambigu: c.length > 1 };
}
function _rdvJoindre_(a, b) { a = String(a || "").trim(); b = String(b || "").trim(); if (!a) return b; if (!b || a.toLowerCase().indexOf(b.toLowerCase()) !== -1) return a; return a + " ; " + b; }

/** Aperçu (appliquer absent) ou application de l'harmonisation des noms de visiteurs / responsables. */
function handleHarmoniserVisiteurs(p, ss) {
  if (!_vsuAdmin_(p)) return createJsonResponse({ success: false, message: "Réservé aux profils Super-admin et Administratif." });
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var props = PropertiesService.getScriptProperties();
    if (p.appliquer && p.alias && typeof p.alias === "object") {
      var al = {}; try { al = JSON.parse(props.getProperty("RDV_ALIAS") || "{}"); } catch (e) { al = {}; }
      Object.keys(p.alias).forEach(function (k) { if (p.alias[k]) al[_vsuNorm_(k)] = String(p.alias[k]).toLowerCase(); });
      props.setProperty("RDV_ALIAS", JSON.stringify(al));
    }
    var ctx = _rdvCtx_(ss), groupes = {};
    var wsV = _vsuEntete_(_rdvSheet_(ss), RDV_HEADERS), dV = wsV.getDataRange().getValues();
    var wsR = _vsuEntete_(_vsxSheet_(ss, "RDV_REFERENTS", RDV_REF_HEADERS), RDV_REF_HEADERS), dR = wsR.getDataRange().getValues();
    var noter = function (brut) {
      var b = String(brut || "").trim(); if (!b) return null;
      var c = _rdvCanon_(b, ctx); groupes[b] = groupes[b] || { brut: b, nom: c.nom, email: c.email, precision: c.precision, reconnu: c.reconnu, ambigu: !!c.ambigu, n: 0 };
      groupes[b].n++; return c;
    };
    var modV = 0, modR = 0;
    for (var i = 1; i < dV.length; i++) {
      if (!dV[i][0]) continue;
      var c = noter(dV[i][8]); if (!c) continue;
      var prec = _rdvJoindre_(dV[i][20], c.precision);
      if (p.appliquer && (String(dV[i][8]) !== c.nom || String(dV[i][9] || "") !== (c.email || String(dV[i][9] || "")) || String(dV[i][20] || "") !== prec)) {
        wsV.getRange(i + 1, 9, 1, 2).setValues([[c.nom, c.email || String(dV[i][9] || "")]]); wsV.getRange(i + 1, 21).setValue(prec); modV++;
      }
    }
    for (var j = 1; j < dR.length; j++) {
      if (!dR[j][1]) continue;
      var c2 = noter(dR[j][2]); if (!c2) continue;
      var prec2 = _rdvJoindre_(dR[j][7], c2.precision);
      if (p.appliquer && (String(dR[j][2]) !== c2.nom || String(dR[j][3] || "") !== (c2.email || String(dR[j][3] || "")) || String(dR[j][7] || "") !== prec2)) {
        wsR.getRange(j + 1, 3, 1, 2).setValues([[c2.nom, c2.email || String(dR[j][3] || "")]]); wsR.getRange(j + 1, 8).setValue(prec2); modR++;
      }
    }
    var liste = Object.keys(groupes).map(function (k) { return groupes[k]; }).sort(function (a, b) { return a.nom.localeCompare(b.nom) || a.brut.localeCompare(b.brut); });
    if (p.appliquer) logAction(p.nomAuteur || p.emailAuteur || "", p.vieScRole || "", "Modification", "Vie scolaire — Suivi des RDV", "RDV_VISITES",
      "Harmonisation des noms de visiteurs : " + modV + " RDV et " + modR + " responsable(s) mis à jour — " + liste.filter(function (g) { return g.brut !== g.nom; }).slice(0, 40).map(function (g) { return "« " + g.brut + " » → " + g.nom + (g.precision ? " [" + g.precision + "]" : ""); }).join(" ; "));
    return createJsonResponse({ success: true, groupes: liste, equipe: ctx.equipe.map(function (q) { return { nom: q.nom, email: q.email }; }), modifies: { rdv: modV, responsables: modR } });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
  finally { try { lock.releaseLock(); } catch (eL) {} }
}

// ── Import d'un fichier Excel (export retravaillé) : aperçu puis application ──
// Colonnes reconnues (export « Suivi des RDV ») : ID RDV, ID étudiant, Responsable visite, Précision
// responsable, Date RDV pris, Date RDV (après report), Visiteur, Précision, Statut, Commentaire.
// Ligne avec ID RDV = mise à jour de ce RDV ; sans ID RDV mais avec une date = nouveau RDV.
function _rdvIsoImport_(v) {
  var s = String(v == null ? "" : v).trim(); if (!s) return "";
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/); if (m) return m[0];
  m = s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})$/);
  if (m) { var y = +m[3] < 100 ? 2000 + +m[3] : +m[3]; return y + "-" + ("0" + m[2]).slice(-2) + "-" + ("0" + m[1]).slice(-2); }
  return "?";
}
function handleImporterRdvExcel(p, ss) {
  if (!_vsuAdmin_(p)) return createJsonResponse({ success: false, message: "Réservé aux profils Super-admin et Administratif." });
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    var lignes = []; try { lignes = typeof p.lignes === "string" ? JSON.parse(p.lignes) : (p.lignes || []); } catch (e) { lignes = []; }
    if (!lignes.length) return createJsonResponse({ success: false, message: "Le fichier ne contient aucune ligne exploitable." });
    var annee = String(p.annee || _ckAnneeCourante_()), ctx = _rdvCtx_(ss), par = String(p.nomAuteur || p.emailAuteur || ""), now = new Date();
    var ws = _vsuEntete_(_rdvSheet_(ss), RDV_HEADERS), d = ws.getDataRange().getValues(), parId = {};
    for (var i = 1; i < d.length; i++) if (d[i][0] && String(d[i][19]) !== "Oui") parId[String(d[i][0])] = { row: i + 1, v: d[i] };
    var etus = {}; _vsuEtudiants_(ss).forEach(function (e) { etus[e.id] = e; });
    var refs = _rdvReferents_(ss, annee), dejaDate = {};
    Object.keys(parId).forEach(function (k) { var v = parId[k].v; dejaDate[String(v[2]) + "|" + _vsuIso_(v[10])] = 1; dejaDate[String(v[2]) + "|" + _vsuIso_(v[11])] = 1; });
    var STATUTS = ["À planifier", "Planifié", "Honoré", "Reporté", "Annulé", "Autre"];
    var statutOk = function (s) { s = String(s || "").trim(); if (!s) return ""; var x = STATUTS.filter(function (t) { return _vsuNorm_(t) === _vsuNorm_(s); })[0]; return x || "?"; };
    var modifs = [], creations = [], refModifs = [], erreurs = [], refVus = {};
    lignes.forEach(function (l, n) {
      var ligne = n + 2, idR = String(l.idRdv || "").trim(), etuId = String(l.etuId || "").trim();
      // Responsable visite de l'étudiant (une fois par étudiant)
      if (etuId && etus[etuId] && String(l.responsable || "").trim() && !refVus[etuId]) {
        refVus[etuId] = 1;
        var c = _rdvCanon_(l.responsable, ctx), prec = l.precisionResp !== undefined ? String(l.precisionResp || "").trim() : c.precision;
        var r = refs[etuId];
        if (!r || r.nom !== c.nom || (r.precision || "") !== prec) refModifs.push({ etuId: etuId, etudiant: (etus[etuId].prenom + " " + etus[etuId].nom).trim(), avant: r ? r.nom + (r.precision ? " (" + r.precision + ")" : "") : "—", apres: c.nom + (prec ? " (" + prec + ")" : ""), nom: c.nom, email: c.email, precision: prec });
      }
      var dp = _rdvIsoImport_(l.datePris), dr = _rdvIsoImport_(l.date), st = statutOk(l.statut);
      if (dp === "?" || dr === "?") { erreurs.push("Ligne " + ligne + " : date illisible (format attendu jj/mm/aaaa)"); return; }
      if (st === "?") { erreurs.push("Ligne " + ligne + " : statut « " + l.statut + " » inconnu (" + STATUTS.join(", ") + ")"); return; }
      if (idR) {
        var x = parId[idR]; if (!x) { erreurs.push("Ligne " + ligne + " : RDV " + idR + " introuvable (supprimé ?)"); return; }
        var v = x.v, ch = [], nv = v.slice(0, RDV_HEADERS.length);
        while (nv.length < RDV_HEADERS.length) nv.push("");
        var vis = String(l.visiteur || "").trim();
        if (vis) { var cv = _rdvCanon_(vis, ctx); if (cv.nom !== String(v[8])) { ch.push(["Visiteur", v[8], cv.nom]); nv[8] = cv.nom; nv[9] = cv.email || ""; }
          var pv = l.precision !== undefined ? String(l.precision || "").trim() : _rdvJoindre_(v[20], cv.precision); if (pv !== String(v[20] || "")) { ch.push(["Précision", v[20], pv]); nv[20] = pv; } }
        else if (l.precision !== undefined && String(l.precision || "").trim() !== String(v[20] || "")) { ch.push(["Précision", v[20], l.precision]); nv[20] = String(l.precision || "").trim(); }
        if (dp && dp !== _vsuIso_(v[10])) { ch.push(["Date RDV pris", _vsuIso_(v[10]), dp]); nv[10] = dp; }
        if (dr && dr !== _vsuIso_(v[11])) { ch.push(["Date RDV", _vsuIso_(v[11]), dr]); nv[11] = dr; }
        if (st && st !== String(v[12])) { ch.push(["Statut", v[12], st]); nv[12] = st; }
        if (l.commentaire !== undefined && String(l.commentaire || "").trim() !== String(v[14] || "").trim()) { ch.push(["Commentaire", v[14], l.commentaire]); nv[14] = String(l.commentaire || "").trim(); }
        if (ch.length) modifs.push({ id: idR, row: x.row, etudiant: String(v[3]), changements: ch.map(function (c) { return { champ: c[0], avant: String(c[1] || ""), apres: String(c[2] || "") }; }), nv: nv, statutChange: st && st !== String(v[12]) });
        return;
      }
      if (!dp && !dr) return;   // ligne « Aucun RDV » ou étudiant sans date : rien à créer
      if (!etuId || !etus[etuId]) { erreurs.push("Ligne " + ligne + " : nouvel RDV sans ID étudiant valide"); return; }
      var dd = dp || dr;
      if (dejaDate[etuId + "|" + dd]) return;
      dejaDate[etuId + "|" + dd] = 1;
      var cn = _rdvCanon_(l.visiteur || "", ctx);
      creations.push({ etuId: etuId, etudiant: (etus[etuId].prenom + " " + etus[etuId].nom).trim(), datePris: dp || dr, date: dr || dp, statut: st || (dd <= Utilities.formatDate(now, "Europe/Paris", "yyyy-MM-dd") ? "Planifié" : "Planifié"),
        visiteur: cn.nom, email: cn.email, precision: l.precision !== undefined ? String(l.precision || "").trim() : cn.precision, commentaire: String(l.commentaire || "").trim() });
    });
    if (!p.appliquer) {
      return createJsonResponse({ success: true, apercu: true, modifs: modifs.map(function (m) { return { id: m.id, etudiant: m.etudiant, changements: m.changements }; }), creations: creations, refModifs: refModifs, erreurs: erreurs });
    }
    var horo = Utilities.formatDate(now, "Europe/Paris", "yyyy-MM-dd HH:mm");
    modifs.forEach(function (m) {
      var h = []; try { h = JSON.parse(String(m.nv[13] || "[]")); } catch (e) { h = []; }
      h.push({ le: horo, statut: m.nv[12], date: _vsuIso_(m.nv[11]), par: par + " (import Excel)", commentaire: m.changements.map(function (c) { return c.champ + " : " + (c.avant || "—") + " → " + (c.apres || "—"); }).join(" | ") });
      m.nv[13] = JSON.stringify(h); m.nv[18] = now;
      ws.getRange(m.row, 1, 1, RDV_HEADERS.length).setValues([m.nv]);
      ws.getRange(m.row, 11, 1, 2).setNumberFormat("@").setValues([[_vsuIso_(m.nv[10]), _vsuIso_(m.nv[11])]]);
    });
    if (creations.length) {
      var ents = _vsuEntreprises_(ss), rows = creations.map(function (c, k) {
        var e = etus[c.etuId], ent = e.idEnt ? ents[e.idEnt.toUpperCase()] : null;
        return ["RDV" + now.getTime() + "X" + k, _vsuAnneeDe_(c.datePris), e.id, c.etudiant, e.classe, ent ? ent.id : "", ent ? ent.nom : e.nomEnt, ent ? ent.lieu : "",
          c.visiteur, c.email, c.datePris, c.date, c.statut, JSON.stringify([{ le: horo, statut: c.statut, date: c.date, par: par + " (import Excel)" }]), c.commentaire, "Import Excel", now, par, now, "", c.precision];
      });
      var r0 = ws.getLastRow() + 1;
      ws.getRange(r0, 11, rows.length, 2).setNumberFormat("@");
      ws.getRange(r0, 1, rows.length, RDV_HEADERS.length).setValues(rows);
    }
    if (refModifs.length) _rdvEcrireReferents_(ss, annee, refModifs, "Saisie", par + " (import Excel)");
    logAction(par, p.vieScRole || "", "Modification", "Vie scolaire — Suivi des RDV", "RDV_VISITES",
      "Import Excel « " + String(p.fichier || "").slice(0, 80) + " » : " + modifs.length + " RDV modifié(s), " + creations.length + " créé(s), " + refModifs.length + " responsable(s) mis à jour"
      + (erreurs.length ? ", " + erreurs.length + " ligne(s) ignorée(s)" : "") + " — " + modifs.slice(0, 20).map(function (m) { return m.etudiant + " (" + m.changements.map(function (c) { return c.champ + " → " + c.apres; }).join(", ") + ")"; }).join(" ; "));
    return createJsonResponse({ success: true, modifies: modifs.length, crees: creations.length, responsables: refModifs.length, erreurs: erreurs });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
  finally { try { lock.releaseLock(); } catch (eL) {} }
}

// ── Drapeaux de priorité (RDV_DRAPEAUX) : persistants, levés uniquement à la main ──
// Drapeau ROUGE automatique (règle du 12/11/2026) : l'étudiant avait HISTORIQUEMENT un visiteur — responsable
// repris de Charlemagne, ou attribué au plus tard le 11/11/2026 — et aucune visite n'a jamais été honorée.
// Les visiteurs attribués après le 11/11/2026 ne déclenchent pas de drapeau. Une fois posé, le drapeau reste
// jusqu'à ce qu'on le lève à la main. Niveaux : Rouge (urgent), Orange, Jaune ; commentaire libre.
var RDV_DRAP_DATE_HISTO = "2026-11-11";
var RDV_DRAP_AUTO_COM = "Visiteur attribué historiquement, visite jamais honorée";
var RDV_DRAP_HEADERS = ["Année scolaire", "ID étudiant", "Niveau", "Commentaire", "Posé le", "Posé par", "Levé le", "Levé par", "Mis à jour le"];
var RDV_NIVEAUX = ["Rouge", "Orange", "Jaune"];
function _rdvDrapeaux_(ss, annee) {
  var d = _vsxSheet_(ss, "RDV_DRAPEAUX", RDV_DRAP_HEADERS).getDataRange().getValues(), out = {};
  for (var i = 1; i < d.length; i++) if (String(d[i][0]) === annee && d[i][1]) out[String(d[i][1])] = {
    niveau: d[i][6] ? "" : String(d[i][2] || ""), commentaire: String(d[i][3] || ""), poseLe: _vsuIso_(d[i][4]), par: String(d[i][5] || ""),
    leveLe: _vsuIso_(d[i][6]), levePar: String(d[i][7] || ""), row: i + 1 };
  return out;
}
function _rdvAutoDrapeaux_(ss, annee, liste, visites, drap) {
  var honore = {}; visites.forEach(function (v) { if (v.statut === "Honoré") honore[v.etuId] = 1; });
  var regle = function (e) {
    var histo = e.referentSource === "Charlemagne" || (e.referentDepuis && e.referentDepuis <= RDV_DRAP_DATE_HISTO);
    return !!e.referent && histo && !honore[e.id];
  };
  var ws = _vsxSheet_(ss, "RDV_DRAPEAUX", RDV_DRAP_HEADERS), now = new Date(), ajout = [], aSupprimer = [];
  // Nettoyage : drapeaux posés automatiquement par l'ancienne règle et qui ne la respectent plus (jamais touchés à la main)
  liste.forEach(function (e) {
    var d = drap[e.id];
    if (d && d.par === "Automatique" && d.niveau === "Rouge" && !d.leveLe && !regle(e)) { aSupprimer.push(d.row); delete drap[e.id]; }
  });
  aSupprimer.sort(function (a, b) { return b - a; }).forEach(function (r) { ws.deleteRow(r); });
  liste.forEach(function (e) {
    if (drap[e.id] || !regle(e)) return;   // déjà un drapeau (même levé) : on ne le repose pas
    ajout.push([annee, e.id, "Rouge", RDV_DRAP_AUTO_COM, now, "Automatique", "", "", now]);
    drap[e.id] = { niveau: "Rouge", commentaire: RDV_DRAP_AUTO_COM, poseLe: _vsuIso_(now), par: "Automatique", leveLe: "", levePar: "" };
  });
  if (ajout.length) ws.getRange(ws.getLastRow() + 1, 1, ajout.length, RDV_DRAP_HEADERS.length).setValues(ajout);
  if (aSupprimer.length || ajout.length) console.log("Drapeaux auto : " + aSupprimer.length + " retiré(s) (ancienne règle), " + ajout.length + " posé(s)");
}
/** Pose / modifie / lève un drapeau (un ou plusieurs étudiants) ; niveau vide = lever ; commentaire seul possible. */
function handleSaveDrapeauRdv(p, ss) {
  if (!_vsxCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var annee = String(p.annee || _ckAnneeCourante_()), ids = (Array.isArray(p.etuIds) ? p.etuIds : [p.etuId]).map(String).filter(Boolean);
    var niveau = String(p.niveau || ""), lever = niveau === "" || niveau === "lever";
    if (!lever && RDV_NIVEAUX.indexOf(niveau) === -1) return createJsonResponse({ success: false, message: "Niveau inconnu." });
    var ws = _vsxSheet_(ss, "RDV_DRAPEAUX", RDV_DRAP_HEADERS), drap = _rdvDrapeaux_(ss, annee), now = new Date(), par = String(p.nomAuteur || p.emailAuteur || ""), ajout = [];
    var com = p.commentaire !== undefined ? String(p.commentaire || "").trim() : null;
    ids.forEach(function (id) {
      var x = drap[id];
      if (x) {
        var r = ws.getRange(x.row, 3, 1, 7).getValues()[0];
        if (lever) { r[4] = now; r[5] = par; } else { r[0] = niveau; r[4] = ""; r[5] = ""; if (x.leveLe || !r[2]) { r[2] = now; r[3] = par; } }
        if (com !== null) r[1] = com;
        r[6] = now;
        ws.getRange(x.row, 3, 1, 7).setValues([r]);
      } else if (!lever) ajout.push([annee, id, niveau, com || "", now, par, "", "", now]);
    });
    if (ajout.length) ws.getRange(ws.getLastRow() + 1, 1, ajout.length, RDV_DRAP_HEADERS.length).setValues(ajout);
    logAction(par, p.vieScRole || "", "Modification", "Vie scolaire — Suivi des RDV", ids.length === 1 ? ids[0] : "RDV_DRAPEAUX",
      (lever ? "Drapeau levé" : "Drapeau " + niveau.toLowerCase()) + " pour " + ids.length + " étudiant(s)" + (com ? " — " + com : ""));
    return createJsonResponse({ success: true });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
  finally { try { lock.releaseLock(); } catch (eL) {} }
}
