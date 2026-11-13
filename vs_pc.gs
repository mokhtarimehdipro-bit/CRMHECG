// =============================================================================
//  VIE SCOLAIRE — SUIVI DES PC (demande direction 13/11/2026) — front : vs_modules.js (VX.pc*)
//  Chaque année des PC sont attribués aux étudiants. Statuts : « PC attribué », « PC restitué »,
//  « PC donné » (PC laissé aux étudiants diplômés chez nous).
//  Import Charlemagne : colonne AH = numéro du PC (notation exponentielle Excel harmonisée),
//  colonne AQ = commentaires (seules les phrases qui parlent du PC quittent l'ordinateur).
//  Modèle : n° commençant par une lettre (PF…) = Lenovo, par 5 (ou NX) = Acer.
//  Une saisie manuelle n'est jamais écrasée par un import (l'écart est signalé « à vérifier »).
// =============================================================================
var PC_HEADERS = ["ID", "ID étudiant", "ID élève Charlemagne", "Nom", "Prénom", "Classe (Charlemagne)", "N° de série", "N° brut Charlemagne",
  "Modèle", "Statut", "Année d'attribution", "Date de restitution", "Commentaire", "Commentaire Charlemagne", "Source", "À vérifier",
  "Saisie manuelle", "Mis à jour le", "Par", "Historique", "À récupérer (choix manuel)", "Motif (récupération)"];
// Alerte « PC à récupérer » (13/11/2026) : étudiant sorti, avec un n° de PC, statut « PC attribué » (ni restitué ni donné
// = pas de diplôme obtenu). Ajustable au cas par cas : « Oui » force l'alerte, « Non » l'écarte (avec un motif).
var PC_STATUTS = ["PC attribué", "PC restitué", "PC donné"];
var PC_ANNEES_DEFAUT = ["Années précédentes", "2026-2027", "2027-2028"];
var PC_CAT = "Vie scolaire — Suivi des PC";

function _pcSheet_(ss) {
  var ws = _vsuEntete_(_vsxSheet_(ss, "PC_SUIVI", PC_HEADERS), PC_HEADERS);
  ws.getRange(1, 7, ws.getMaxRows(), 2).setNumberFormat("@");    // n° de série : jamais de notation exponentielle
  ws.getRange(1, 11, ws.getMaxRows(), 2).setNumberFormat("@");   // année / date de restitution en texte
  return ws;
}
function _pcLire_(ss) {
  var ws = ss.getSheetByName("PC_SUIVI"); if (!ws || ws.getLastRow() < 2) return [];
  var d = ws.getRange(2, 1, ws.getLastRow() - 1, PC_HEADERS.length).getValues(), out = [];
  d.forEach(function (r, i) {
    if (!r[0]) return;
    var h = []; try { h = JSON.parse(String(r[19] || "[]")); } catch (e) {}
    out.push({ row: i + 2, pcId: String(r[0]), etuId: String(r[1] || ""), idEleve: String(r[2] || ""), nom: String(r[3] || ""), prenom: String(r[4] || ""),
      classeCharl: String(r[5] || ""), serie: String(r[6] || ""), brut: String(r[7] || ""), modele: String(r[8] || ""), statut: String(r[9] || ""),
      annee: String(r[10] || ""), dateRestit: r[11] instanceof Date ? _vsuIso_(r[11]) : String(r[11] || ""), commentaire: String(r[12] || ""),
      comCharl: String(r[13] || ""), source: String(r[14] || ""), verifier: String(r[15] || ""), manuel: String(r[16]) === "Oui",
      maj: _vsuIso_(r[17]), par: String(r[18] || ""), historique: h, recupForce: String(r[20] || ""), recupMotif: String(r[21] || "") });
  });
  return out;
}
function _pcVersLigne_(x) {
  return [x.pcId, x.etuId, x.idEleve, x.nom, x.prenom, x.classeCharl, x.serie, x.brut, x.modele, x.statut, x.annee, x.dateRestit,
    x.commentaire, x.comCharl, x.source, x.verifier, x.manuel ? "Oui" : "", x.majDate || new Date(), x.par, JSON.stringify(x.historique || []), x.recupForce || "", x.recupMotif || ""];
}
function _pcAnnees_() {
  var l = PC_ANNEES_DEFAUT.slice();
  try { (JSON.parse(PropertiesService.getScriptProperties().getProperty("PC_ANNEES") || "[]") || []).forEach(function (a) { if (l.indexOf(a) === -1) l.push(a); }); } catch (e) {}
  return [l[0]].concat(l.slice(1).sort());
}
function _pcModele_(serie) {
  var s = String(serie || "").toUpperCase();
  if (!s) return "";
  if (/^5/.test(s) || /^NX/.test(s)) return "Acer";
  if (/^[A-Z]/.test(s)) return "Lenovo";
  return "À préciser";
}
function _pcAnneeDeEntree_(entree) {   // « 05/09/2026 » → 2026-2027 ; avant 2026-2027 → « Années précédentes »
  var m = String(entree || "").match(/(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})/), iso = m ? m[3] + "-" + ("0" + m[2]).slice(-2) : "";
  if (!iso) { m = String(entree || "").match(/^(\d{4})-(\d{2})/); if (m) iso = m[1] + "-" + m[2]; }
  if (!iso) return PC_ANNEES_DEFAUT[0];
  var a = _vsuAnneeDe_(iso + "-01");
  return a < "2026-2027" ? PC_ANNEES_DEFAUT[0] : a;
}

// ── Lecture de la colonne AH : n° de PC (harmonisation de la notation exponentielle) + mentions ──
function _pcAnalyser_(ah, com, motif) {
  var t = String(ah || ""), notes = [], series = [];
  t = t.replace(/(\d+)(?:[.,](\d+))?\s*E\s*\+?\s*(\d{1,2})\b/gi, function (all, ent, dec, ex) {   // 5,00123E+11 → 500123000000
    dec = dec || ""; var len = ent.length + (+ex), dig = ent + dec;
    if (dig.length > len) return all;
    if (dig.length < len) notes.push("n° « " + all + " » en notation exponentielle Excel : les derniers chiffres sont à vérifier sur l'étiquette");
    return dig + new Array(len - dig.length + 1).join("0");
  });
  t.split(/[\s;\/|,+]+|\bet\b/i).forEach(function (w) {
    w = w.replace(/^n[°o]\s*/i, "").replace(/[^A-Za-z0-9\-]/g, "").replace(/-/g, "").toUpperCase();
    if (w.length >= 6 && /\d/.test(w) && !/^\d{6}$/.test(w) && !/^(19|20)\d{2}\d{0,4}$/.test(w) && series.indexOf(w) === -1) series.push(w);
  });
  var n = _vsuNorm_(t.replace(/\b[A-Z0-9]{6,}\b/g, " ") + " " + (com || "")), nm = _vsuNorm_(motif || "");
  var nonRestit = /\b(pas|non|jamais|ni)\s+(encore\s+)?(ete\s+)?(restitue|restituee|rendu|rendue|recupere|recuperee|retourne)\b/.test(n);
  var restit = !nonRestit && /\b(restitue|restituee|restitues|restitution|rendu|rendue|recupere|recuperee|retourne|retournee)\b/.test(n);
  var donne = /\bdiplome\w*\s+(obtenu|a demander)\b/.test(n) || /^\s*(obtenu|a demander)\b/.test(_vsuNorm_(ah)) || /\bpc\s+(donne|offert)\b/.test(n)
    || /\b(obtention\s+(du\s+)?diplome|diplome\w*\s+obtenu)\b/.test(nm);
  var statut = restit ? "PC restitué" : donne ? "PC donné" : series.length ? "PC attribué" : "";
  if (restit && donne) notes.push("Charlemagne mentionne à la fois une restitution et un diplôme obtenu : statut « PC restitué » retenu, à confirmer");
  var dateRestit = "";
  if (restit) {
    var phr = (String(ah || "") + " | " + String(com || "")).split(/\||\n|(?:[.!;]\s)/).filter(function (s) { return /restitu|rendu|r[eé]cup[eé]r|retourn/i.test(s); });
    phr.some(function (s) { var m = s.match(/(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})/); if (!m) return false; var y = +m[3] < 100 ? 2000 + +m[3] : +m[3]; dateRestit = y + "-" + ("0" + m[2]).slice(-2) + "-" + ("0" + m[1]).slice(-2); return true; });
  }
  return { series: series, statut: statut, dateRestit: dateRestit, notes: notes };
}

// ── Import depuis l'export Charlemagne (appelé par le contrôle Charlemagne ou depuis la page PC) ──
function _pcImporterCharl_(ss, lignes, par) {
  var S = { etudiants: 0, crees: 0, maj: 0, restitues: 0, donnes: 0, nonRapproches: 0, aVerifier: 0, manuelsPreserves: 0 };
  var avecPc = lignes.filter(function (l) { return String(l.pc || "").trim() || String(l.pcCom || "").trim(); });
  if (!avecPc.length) return S;
  var etus = _vsuEtudiants_(ss), parId = {}, parMail = {}, parNom = {};
  etus.forEach(function (e) {
    if (e.idEleve) parId[e.idEleve] = e;
    [e.mailPerso, e.mailHecg].forEach(function (m) { if (m) parMail[m] = e; });
    var k = _vsuNorm_(e.nom + " " + e.prenom); parNom[k] = parNom[k] || e;
  });
  var ws = _pcSheet_(ss), exist = _pcLire_(ss), now = new Date(), le = Utilities.formatDate(now, "Europe/Paris", "yyyy-MM-dd HH:mm");
  var cle = function (x) { return x.etuId || ("CH:" + x.idEleve); };
  var ajouts = [], modifs = {};
  avecPc.forEach(function (l) {
    var A = _pcAnalyser_(l.pc, l.pcCom, l.motif);
    if (!A.series.length && !A.statut) return;   // commentaire sans information exploitable
    S.etudiants++;
    var e = (l.idEleve && parId[l.idEleve]) || parMail[String(l.email || "").toLowerCase()] || parMail[String(l.mailPerso || "").toLowerCase()] || parNom[_vsuNorm_(l.nom + " " + l.prenom)];
    if (!e) S.nonRapproches++;
    var base = { etuId: e ? e.id : "", idEleve: String(l.idEleve || (e ? e.idEleve : "")), nom: e ? e.nom : String(l.nom || ""), prenom: e ? e.prenom : String(l.prenom || ""), classeCharl: String(l.classe || "") };
    var k = base.etuId || ("CH:" + base.idEleve);
    var series = A.series.length ? A.series : [""];
    series.forEach(function (serie) {
      var x = exist.filter(function (y) { return cle(y) === k && y.serie === serie; })[0]
        || (serie ? exist.filter(function (y) { return cle(y) === k && !y.serie; })[0] : exist.filter(function (y) { return cle(y) === k; }).length === 1 ? exist.filter(function (y) { return cle(y) === k; })[0] : null);
      var notes = A.notes.slice(); if (!e) notes.push("étudiant introuvable dans le CRM (rapproché par l'identifiant élève Charlemagne)");
      var statut = A.statut || "PC attribué";
      if (!x) {
        x = { pcId: "PC" + now.getTime() + "-" + (ajouts.length + 1), etuId: base.etuId, idEleve: base.idEleve, nom: base.nom, prenom: base.prenom, classeCharl: base.classeCharl,
          serie: serie, brut: String(l.pc || ""), modele: _pcModele_(serie), statut: statut, annee: _pcAnneeDeEntree_(l.entree), dateRestit: statut === "PC restitué" ? A.dateRestit : "",
          commentaire: "", comCharl: String(l.pcCom || ""), source: "Charlemagne", verifier: notes.join(" ; "), manuel: false, par: "Import Charlemagne",
          historique: [{ le: le, par: par || "Import Charlemagne", quoi: "Créé depuis Charlemagne : " + statut + (serie ? " — n° " + serie : "") }] };
        ajouts.push(x); exist.push(x); S.crees++;
      } else {
        var avant = x.statut;
        x.idEleve = x.idEleve || base.idEleve; x.classeCharl = base.classeCharl || x.classeCharl; x.brut = String(l.pc || ""); x.comCharl = String(l.pcCom || "");
        if (!x.etuId && base.etuId) { x.etuId = base.etuId; x.nom = base.nom; x.prenom = base.prenom; }
        if (x.manuel) {
          if (A.statut && A.statut !== x.statut) notes.push("Charlemagne indique « " + A.statut + " » (saisie manuelle « " + x.statut + " » conservée)");
          S.manuelsPreserves++;
        } else {
          if (serie && !x.serie) { x.serie = serie; x.modele = x.modele || _pcModele_(serie); }
          if (A.statut) x.statut = A.statut;
          if (x.statut === "PC restitué" && A.dateRestit) x.dateRestit = A.dateRestit;
        }
        x.verifier = notes.join(" ; ");
        if (avant !== x.statut) x.historique.push({ le: le, par: par || "Import Charlemagne", quoi: "Statut " + avant + " → " + x.statut + " (Charlemagne)" });
        if (x.row) modifs[x.row] = x;
        S.maj++;
      }
      if (x.statut === "PC restitué") S.restitues++;
      if (x.statut === "PC donné") S.donnes++;
    });
  });
  // un même n° de PC sur deux étudiants : à vérifier
  var parSerie = {};
  exist.forEach(function (x) { if (x.serie) (parSerie[x.serie] = parSerie[x.serie] || []).push(x); });
  Object.keys(parSerie).forEach(function (s) {
    var g = parSerie[s]; if (g.length < 2 || g.every(function (x) { return cle(x) === cle(g[0]); })) return;
    g.forEach(function (x) {
      var msg = "n° " + s + " aussi attribué à " + g.filter(function (y) { return y !== x; }).map(function (y) { return (y.prenom + " " + y.nom).trim(); }).join(", ");
      if (x.verifier.indexOf(msg) === -1) { x.verifier = (x.verifier ? x.verifier + " ; " : "") + msg; if (x.row) modifs[x.row] = x; }
    });
  });
  Object.keys(modifs).forEach(function (r) { var x = modifs[r]; x.majDate = now; x.par = x.manuel ? x.par : "Import Charlemagne"; ws.getRange(+r, 1, 1, PC_HEADERS.length).setValues([_pcVersLigne_(x)]); });
  if (ajouts.length) ws.getRange(ws.getLastRow() + 1, 1, ajouts.length, PC_HEADERS.length).setValues(ajouts.map(function (x) { x.majDate = now; return _pcVersLigne_(x); }));
  S.aVerifier = exist.filter(function (x) { return x.verifier; }).length;
  logAction(par || "Import Charlemagne", "", "Action", PC_CAT, "PC_SUIVI",
    "Import Charlemagne : " + S.etudiants + " étudiant(s) avec un PC — " + S.crees + " PC créé(s), " + S.maj + " mis à jour, " + S.restitues + " restitué(s), " + S.donnes + " donné(s), "
    + S.nonRapproches + " hors CRM, " + S.manuelsPreserves + " saisie(s) manuelle(s) conservée(s)");
  return S;
}

function handleGetPc(p, ss) {
  if (!_vsxCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  try {
    var etus = _vsuEtudiants_(ss), byId = {}; etus.forEach(function (e) { byId[e.id] = e; });
    var pcs = _pcLire_(ss), avec = {}, out = [];
    var etat = function (e) { return e ? (e.typo || "Non renseigné") : "Hors CRM"; };
    var regime = function (e) { return e ? (_vsuAlternant_(e) ? "Alternance" : "Initial") : "—"; };   // alternance = statut ou entreprise renseignée
    pcs.forEach(function (x) {
      var e = x.etuId ? byId[x.etuId] : null; if (x.etuId) avec[x.etuId] = 1;
      var classe = e && e.classe ? e.classe : x.classeCharl;
      var recupAuto = !!(e && /sorti/i.test(e.typo) && x.serie && x.statut === "PC attribué");
      out.push({ pcId: x.pcId, etuId: x.etuId, idEleve: x.idEleve || (e ? e.idEleve : ""), nom: e ? e.nom : x.nom, prenom: e ? e.prenom : x.prenom, classe: classe, diplome: _vsuDiplome_(classe),
        campus: e ? e.campus : "", etat: etat(e), regime: regime(e), serie: x.serie, brut: x.brut, modele: x.modele, statut: x.statut, annee: x.annee, dateRestit: x.dateRestit,
        commentaire: x.commentaire, comCharl: x.comCharl, source: x.source, verifier: x.verifier, manuel: x.manuel, maj: x.maj, par: x.par, historique: x.historique,
        recupAuto: recupAuto, recupForce: x.recupForce, recupMotif: x.recupMotif, recup: x.recupForce === "Oui" ? true : x.recupForce === "Non" ? false : recupAuto });
    });
    etus.forEach(function (e) {   // étudiants inscrits / préinscrits sans PC enregistré
      if (!e.inscrit || avec[e.id]) return;
      out.push({ pcId: "", etuId: e.id, idEleve: e.idEleve, nom: e.nom, prenom: e.prenom, classe: e.classe, diplome: _vsuDiplome_(e.classe), campus: e.campus, etat: etat(e), regime: regime(e),
        serie: "", brut: "", modele: "", statut: "Aucun PC", annee: "", dateRestit: "", commentaire: "", comCharl: "", source: "", verifier: "", manuel: false, maj: "", par: "", historique: [],
        recupAuto: false, recupForce: "", recupMotif: "", recup: false });
    });
    var dern = ""; try { var ci = ss.getSheetByName("CHARLEMAGNE_IMPORT"); if (ci) dern = String(ci.getRange(1, 1).getValue() || ""); } catch (eC) {}
    return createJsonResponse({ success: true, lignes: out, statuts: PC_STATUTS, annees: _pcAnnees_(), dernierImport: dern, inventaire: _pcInvAutorise_(p) });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

/** Crée / modifie un ou plusieurs PC. items : [{pcId?, etuId, champs:{serie, modele, statut, annee, dateRestit, commentaire}}] */
function handleSavePc(p, ss) {
  if (!_vsxCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var items = Array.isArray(p.items) ? p.items : [], par = String(p.nomAuteur || p.emailAuteur || ""), now = new Date(), le = Utilities.formatDate(now, "Europe/Paris", "yyyy-MM-dd HH:mm");
    var ws = _pcSheet_(ss), pcs = _pcLire_(ss), byPc = {}; pcs.forEach(function (x) { byPc[x.pcId] = x; });
    var etus = {}; _vsuEtudiants_(ss).forEach(function (e) { etus[e.id] = e; });
    var annees = _pcAnnees_(), ajouts = [], noms = [], res = [];
    items.forEach(function (it, n) {
      var c = it.champs || {}, x = it.pcId ? byPc[String(it.pcId)] : null;
      if (it.pcId && !x) return;
      if (c.statut !== undefined && PC_STATUTS.indexOf(c.statut) === -1) throw new Error("Statut inconnu : " + c.statut);
      if (c.annee && annees.indexOf(c.annee) === -1) throw new Error("Année d'attribution inconnue : " + c.annee);
      if (!x) {
        var e = etus[String(it.etuId || "")]; if (!e) throw new Error("Étudiant introuvable : " + it.etuId);
        x = { pcId: "PC" + now.getTime() + "-M" + n, etuId: e.id, idEleve: e.idEleve, nom: e.nom, prenom: e.prenom, classeCharl: "", serie: "", brut: "", modele: "", statut: "PC attribué",
          annee: "", dateRestit: "", commentaire: "", comCharl: "", source: "Saisie", verifier: "", manuel: true, historique: [] };
        ajouts.push(x);
      }
      var ch = [];
      if (c.serie !== undefined) { var s = String(c.serie || "").replace(/\s+/g, "").toUpperCase(); if (s !== x.serie) { ch.push("n° " + (x.serie || "—") + " → " + (s || "—")); x.serie = s; if (c.modele === undefined && !x.modele) x.modele = _pcModele_(s); } }
      if (c.recupForce !== undefined && ["", "Oui", "Non"].indexOf(String(c.recupForce)) === -1) throw new Error("Choix de récupération inconnu");
      ["modele", "statut", "annee", "dateRestit", "commentaire", "recupForce", "recupMotif"].forEach(function (f) {
        if (c[f] === undefined) return; var v = String(c[f] || "").trim(); if (v === (x[f] || "")) return;
        ch.push({ modele: "modèle", statut: "statut", annee: "année", dateRestit: "date de restitution", commentaire: "commentaire", recupForce: "alerte « à récupérer »", recupMotif: "motif récupération" }[f] + " : "
          + (f === "commentaire" || f === "recupMotif" ? (f === "recupMotif" ? v || "—" : "modifié") : f === "recupForce" ? ({ "": "automatique", Oui: "forcée", Non: "écartée" }[x[f] || ""] + " → " + { "": "automatique", Oui: "forcée", Non: "écartée" }[v]) : (x[f] || "—") + " → " + (v || "—")));
        x[f] = v;
      });
      if (c.statut === "PC restitué" && !x.dateRestit && c.dateRestit === undefined) x.dateRestit = Utilities.formatDate(now, "Europe/Paris", "yyyy-MM-dd");
      if (!ch.length && x.row) return;
      x.manuel = true; x.par = par; x.majDate = now;
      if (/notation exponentielle|Charlemagne indique/.test(x.verifier) && (c.serie !== undefined || c.statut !== undefined)) x.verifier = "";   // vérifié à la main
      x.historique.push({ le: le, par: par, quoi: x.row ? ch.join(" ; ") : "PC attribué manuellement" + (ch.length ? " — " + ch.join(" ; ") : "") });
      if (x.row) ws.getRange(x.row, 1, 1, PC_HEADERS.length).setValues([_pcVersLigne_(x)]);
      noms.push((x.prenom + " " + x.nom).trim() + (ch.length ? " (" + ch.join(" ; ") + ")" : ""));
      res.push({ pcId: x.pcId, etuId: x.etuId });
    });
    if (ajouts.length) ws.getRange(ws.getLastRow() + 1, 1, ajouts.length, PC_HEADERS.length).setValues(ajouts.map(_pcVersLigne_));
    if (noms.length) logAction(par, p.vieScRole || "", ajouts.length ? "Création" : "Modification", PC_CAT, items.length === 1 ? String(items[0].pcId || items[0].etuId || "") : "PC_SUIVI",
      (noms.length > 8 ? noms.length + " PC mis à jour : " + noms.slice(0, 8).join(", ") + "…" : noms.join(" · ")));
    return createJsonResponse({ success: true, maj: res });
  } catch (err) { return createJsonResponse({ success: false, message: err.message || err.toString() }); }
  finally { try { lock.releaseLock(); } catch (eL) {} }
}

function handleDeletePc(p, ss) {
  if (!_vsuAdmin_(p)) return createJsonResponse({ success: false, message: "Réservé aux profils Super-admin et Administratif." });
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var x = _pcLire_(ss).filter(function (y) { return y.pcId === String(p.pcId || ""); })[0];
    if (!x) return createJsonResponse({ success: false, message: "PC introuvable." });
    ss.getSheetByName("PC_SUIVI").deleteRow(x.row);
    logAction(p.nomAuteur || p.emailAuteur || "", p.vieScRole || "", "Suppression", PC_CAT, x.pcId, "Ligne PC supprimée : " + (x.prenom + " " + x.nom).trim() + " — n° " + (x.serie || "—") + " — " + x.statut);
    return createJsonResponse({ success: true });
  } catch (err) { return createJsonResponse({ success: false, message: err.toString() }); }
  finally { try { lock.releaseLock(); } catch (eL) {} }
}

function handleSavePcAnnee(p, ss) {
  if (!_vsxCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  var a = String(p.annee || "").trim();
  if (!/^\d{4}-\d{4}$/.test(a) || +a.slice(5) !== +a.slice(0, 4) + 1) return createJsonResponse({ success: false, message: "Format attendu : 2028-2029" });
  var props = PropertiesService.getScriptProperties(), l = [];
  try { l = JSON.parse(props.getProperty("PC_ANNEES") || "[]") || []; } catch (e) {}
  if (l.indexOf(a) === -1 && PC_ANNEES_DEFAUT.indexOf(a) === -1) { l.push(a); props.setProperty("PC_ANNEES", JSON.stringify(l)); logAction(p.nomAuteur || p.emailAuteur || "", p.vieScRole || "", "Création", PC_CAT, "PC_ANNEES", "Année scolaire ajoutée : " + a); }
  return createJsonResponse({ success: true, annees: _pcAnnees_() });
}

/** Import direct depuis la page « Suivi des PC » (lignes lues dans le navigateur), ou reprise du dernier import Charlemagne. */
function handleImporterPc(p, ss) {
  if (!_vsxCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var lignes = [];
    if (p.lignes) { try { lignes = typeof p.lignes === "string" ? JSON.parse(p.lignes) : p.lignes; } catch (e) { lignes = []; } }
    else {
      var ws = ss.getSheetByName("CHARLEMAGNE_IMPORT");
      if (!ws || ws.getLastRow() < 3) return createJsonResponse({ success: false, message: "Aucun import Charlemagne enregistré : importez le fichier." });
      var d = ws.getDataRange().getValues(), h = d[1].map(String);
      if (h.indexOf("pc") === -1) return createJsonResponse({ success: false, message: "Le dernier import Charlemagne ne contient pas les colonnes PC : importez à nouveau le fichier." });
      lignes = d.slice(2).map(function (r) { var o = {}; h.forEach(function (k, i) { o[k] = String(r[i] == null ? "" : r[i]); }); return o; });
    }
    var champs = ["idEleve", "nom", "prenom", "email", "mailPerso", "classe", "entree", "motif", "pc", "pcCom"];
    lignes = (lignes || []).filter(function (l) { return l && (l.nom || l.idEleve); }).map(function (l) { var o = {}; champs.forEach(function (c) { o[c] = l[c] == null ? "" : String(l[c]).trim(); }); return o; });
    lignes = _pcFusionner_(lignes);
    var S = _pcImporterCharl_(ss, lignes, String(p.nomAuteur || p.emailAuteur || ""));
    return createJsonResponse({ success: true, stats: S });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
  finally { try { lock.releaseLock(); } catch (eL) {} }
}
// L'export répète l'élève sur plusieurs lignes : une entrée par identifiant, n° et commentaires cumulés
function _pcFusionner_(lignes) {
  var par = {}, out = [];
  lignes.forEach(function (l) {
    var k = l.idEleve || ("?" + _vsuNorm_(l.nom + " " + l.prenom)), m = par[k];
    if (!m) { m = par[k] = {}; Object.keys(l).forEach(function (c) { m[c] = l[c]; }); out.push(m); return; }
    Object.keys(l).forEach(function (c) {
      if (c === "pc" || c === "pcCom") { if (l[c] && String(m[c] || "").indexOf(l[c]) === -1) m[c] = m[c] ? m[c] + " | " + l[c] : l[c]; }
      else if (!m[c] && l[c]) m[c] = l[c];
    });
  });
  return out;
}

// ═══════════════════════════ INVENTAIRE DES PC (réservé m.mokhtari@hecg.fr — 13/11/2026) ═══════════════════════════
// PC_INVENTAIRE : les PC de l'établissement et leur emplacement de rangement. Un PC est « Disponible » sauf si son
// n° figure dans PC_SUIVI avec le statut « PC attribué » ou « PC donné » (un PC restitué redevient disponible).
var PCINV_HEADERS = ["ID", "N° de série", "Modèle", "Emplacement", "Commentaire", "Ajouté le", "Ajouté par", "Mis à jour le", "Par"];
var PCINV_EMAIL = "m.mokhtari@hecg.fr";
function _pcInvAutorise_(p) { return String(p.emailAuteur || "").trim().toLowerCase() === PCINV_EMAIL; }
function _pcInvSheet_(ss) {
  var ws = _vsuEntete_(_vsxSheet_(ss, "PC_INVENTAIRE", PCINV_HEADERS), PCINV_HEADERS);
  ws.getRange(1, 2, ws.getMaxRows(), 1).setNumberFormat("@");
  return ws;
}
function _pcInvLire_(ss) {
  var ws = ss.getSheetByName("PC_INVENTAIRE"); if (!ws || ws.getLastRow() < 2) return [];
  return ws.getRange(2, 1, ws.getLastRow() - 1, PCINV_HEADERS.length).getValues().map(function (r, i) {
    return { row: i + 2, id: String(r[0] || ""), serie: String(r[1] || ""), modele: String(r[2] || ""), emplacement: String(r[3] || ""), commentaire: String(r[4] || ""),
      ajouteLe: _vsuIso_(r[5]), ajoutePar: String(r[6] || ""), maj: _vsuIso_(r[7]), par: String(r[8] || "") };
  }).filter(function (x) { return x.id; });
}
function _pcNormSerie_(s) { return String(s || "").replace(/\s+/g, "").toUpperCase(); }

function handleGetPcInventaire(p, ss) {
  if (!_pcInvAutorise_(p)) return createJsonResponse({ success: false, message: "Inventaire réservé à la direction." });
  try {
    var etus = {}; _vsuEtudiants_(ss).forEach(function (e) { etus[e.id] = e; });
    var parSerie = {};
    _pcLire_(ss).forEach(function (x) {
      if (!x.serie) return;
      var e = etus[x.etuId], o = { statut: x.statut, etudiant: e ? (e.prenom + " " + e.nom).trim() : (x.prenom + " " + x.nom).trim(), classe: e ? e.classe : x.classeCharl, etat: e ? e.typo : "Hors CRM", dateRestit: x.dateRestit };
      var c = parSerie[x.serie];   // priorité : attribué > donné > restitué
      var rang = { "PC attribué": 3, "PC donné": 2, "PC restitué": 1 };
      if (!c || (rang[o.statut] || 0) > (rang[c.statut] || 0)) parSerie[x.serie] = o;
    });
    var inv = _pcInvLire_(ss), dansInv = {};
    var items = inv.map(function (x) {
      dansInv[x.serie] = 1;
      var s = parSerie[x.serie], dispo = !s || s.statut === "PC restitué";
      return { id: x.id, serie: x.serie, modele: x.modele || _pcModele_(x.serie), emplacement: x.emplacement, commentaire: x.commentaire, ajouteLe: x.ajouteLe, ajoutePar: x.ajoutePar, maj: x.maj, par: x.par,
        etat: dispo ? "Disponible" : s.statut === "PC donné" ? "Donné" : "Attribué", chez: dispo ? (s ? "restitué par " + s.etudiant + (s.dateRestit ? " le " + s.dateRestit.split("-").reverse().join("/") : "") : "") : s.etudiant + (s.classe ? " (" + s.classe + ")" : "") };
    });
    var restitues = Object.keys(parSerie).filter(function (k) { return parSerie[k].statut === "PC restitué" && !dansInv[k]; }).map(function (k) {
      var s = parSerie[k]; return { serie: k, modele: _pcModele_(k), etudiant: s.etudiant, classe: s.classe, dateRestit: s.dateRestit };
    });
    return createJsonResponse({ success: true, items: items, restitues: restitues });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

/** ajouts : [{serie, modele, emplacement, commentaire}] (n° déjà présent = mise à jour), maj : [{id, champs}], suppr : [id] */
function handleSavePcInventaire(p, ss) {
  if (!_pcInvAutorise_(p)) return createJsonResponse({ success: false, message: "Inventaire réservé à la direction." });
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var ws = _pcInvSheet_(ss), inv = _pcInvLire_(ss), par = String(p.nomAuteur || p.emailAuteur || ""), now = new Date();
    var bySerie = {}, byId = {}; inv.forEach(function (x) { bySerie[x.serie] = x; byId[x.id] = x; });
    var nouv = [], nAj = 0, nMaj = 0, nSup = 0, det = [];
    var maj = function (x, c) {
      var ch = [];
      if (c.serie !== undefined) { var s = _pcNormSerie_(c.serie); if (s && s !== x.serie) { if (bySerie[s] && bySerie[s] !== x) throw new Error("Le n° " + s + " est déjà dans l'inventaire."); ch.push("n° " + x.serie + " → " + s); delete bySerie[x.serie]; x.serie = s; bySerie[s] = x; } }
      ["modele", "emplacement", "commentaire"].forEach(function (f) { if (c[f] === undefined) return; var v = String(c[f] || "").trim(); if (v !== x[f]) { ch.push(f + " : " + (x[f] || "—") + " → " + (v || "—")); x[f] = v; } });
      if (ch.length && x.row) { ws.getRange(x.row, 2, 1, 8).setValues([[x.serie, x.modele, x.emplacement, x.commentaire, x.ajouteLe ? new Date(x.ajouteLe + "T12:00:00") : now, x.ajoutePar, now, par]]); nMaj++; det.push(x.serie + " (" + ch.join(", ") + ")"); }
    };
    (Array.isArray(p.ajouts) ? p.ajouts : []).forEach(function (a, i) {
      var s = _pcNormSerie_(a.serie); if (!s) return;
      if (bySerie[s]) { maj(bySerie[s], { modele: a.modele || undefined, emplacement: a.emplacement, commentaire: a.commentaire || undefined }); return; }
      var x = { id: "INV" + now.getTime() + "-" + i, serie: s, modele: String(a.modele || "") || _pcModele_(s), emplacement: String(a.emplacement || "").trim(), commentaire: String(a.commentaire || "").trim() };
      bySerie[s] = x; nouv.push([x.id, x.serie, x.modele, x.emplacement, x.commentaire, now, par, now, par]); nAj++;
    });
    (Array.isArray(p.maj) ? p.maj : []).forEach(function (m) { var x = byId[String(m.id)]; if (x) maj(x, m.champs || {}); });
    var sup = (Array.isArray(p.suppr) ? p.suppr : []).map(function (id) { return byId[String(id)]; }).filter(Boolean).sort(function (a, b) { return b.row - a.row; });
    sup.forEach(function (x) { ws.deleteRow(x.row); nSup++; det.push(x.serie + " supprimé"); });
    if (nouv.length) ws.getRange(ws.getLastRow() + 1, 1, nouv.length, PCINV_HEADERS.length).setValues(nouv);
    if (nAj || nMaj || nSup) logAction(par, p.vieScRole || "", nSup ? "Suppression" : nAj ? "Création" : "Modification", PC_CAT + " — Inventaire", "PC_INVENTAIRE",
      nAj + " PC ajouté(s)" + (nouv.length ? " (" + nouv.slice(0, 10).map(function (r) { return r[1] + (r[3] ? " → " + r[3] : ""); }).join(", ") + (nouv.length > 10 ? "…" : "") + ")" : "") + ", " + nMaj + " modifié(s), " + nSup + " supprimé(s)" + (det.length ? " — " + det.slice(0, 10).join(" ; ") : ""));
    return createJsonResponse({ success: true, ajoutes: nAj, modifies: nMaj, supprimes: nSup });
  } catch (err) { return createJsonResponse({ success: false, message: err.message || err.toString() }); }
  finally { try { lock.releaseLock(); } catch (eL) {} }
}
