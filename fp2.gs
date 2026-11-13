// =============================================================================
//  FICHES DE POSITIONNEMENT — REFONTE (demande direction 2026-11-13) — compte maître uniquement
//
//  • Inventaire : toutes les fiches historiques du Drive (faites par l'équipe, toutes mises en
//    forme) rattachées aux étudiants (dossier GED, puis nom/prénom dans le nom ou le chemin).
//  • Un classeur Google par étudiant dans son dossier GED : 1 onglet par fiche historique (copie
//    fidèle, l'original n'est jamais modifié) + les onglets « nouvelle trame » (modèle HECG) :
//      – préinscrit : 1 seul onglet ;
//      – inscrit : 1 onglet par diplôme (diplômes des fiches historiques + diplôme actuel) ;
//      – refusé (sortis choisis par la direction) : 1 onglet « fiche de refus », note < 10 + motif.
//  • Grille d'évaluation : reprise de la fiche historique du même diplôme, sinon laissée vide
//    (jamais inventée) et signalée dans l'état des mentions manquantes.
//  • Données : CRM (classe, campus, statut, entreprise), export Charlemagne (date de naissance),
//    fiches historiques (civilité, ville, nationalité, langues, cursus…), pièces d'identité /
//    dossiers d'inscription du dossier GED lus par OCR (valeurs signalées « à vérifier »).
//  • Bachelor : onglet « Critères d'admission » pondérés (paramétrables) + tests d'admission.
//  • Moteur en arrière-plan (déclencheur fp2Tick chaque minute, bail de 5 min) avec curseur et
//    journal : reprise automatique après coupure, suivi en temps réel dans positionnement.html.
//  Les fiches générées le 10/11/2026 (positionnement.gs) partent à la corbeille une fois remplacées.
// =============================================================================
var FP2_SS = "1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ";
var FP2_MAITRE = "m.mokhtari@hecg.fr";
var FP2_H = ["ID étudiant", "Nom", "Prénom", "Typo", "Classe", "Diplôme actuel", "Historiques", "Décision", "Lien redirection",
  "Note refus", "Motif refus", "Classeur ID", "Classeur URL", "État", "Manquants", "Message", "MAJ"];
var FP2_C = { id: 0, nom: 1, prenom: 2, typo: 3, classe: 4, diplome: 5, histo: 6, decision: 7, lien: 8, noteRefus: 9, motifRefus: 10,
  classeurId: 11, classeurUrl: 12, etat: 13, manquants: 14, message: 15, maj: 16 };
var FP2_HH = ["ID fichier", "Nom fichier", "Type", "Créé le", "Lien", "ID converti", "Formation", "Diplôme", "Données", "Extrait le", "Erreur"];
var FP2_DIPLOMES = ["BTS CG", "BTS GPME", "DCG", "DSCG", "BACHELOR"];

// ── Accès, feuilles, état ────────────────────────────────────────────────────
function _fp2Autorise_(p) { return String(p.emailAuteur || "").trim().toLowerCase() === FP2_MAITRE; }
function _fp2Feuille_(ss, nom, H) {
  var ws = ss.getSheetByName(nom);
  if (!ws) { ws = ss.insertSheet(nom); ws.getRange(1, 1, 1, H.length).setValues([H]).setFontWeight("bold").setBackground("#1A4E8A").setFontColor("#ffffff"); ws.setFrozenRows(1); }
  return ws;
}
function _fp2Etu_(ss) { return _fp2Feuille_(ss, "FP2_ETUDIANTS", FP2_H); }
function _fp2Histo_(ss) { return _fp2Feuille_(ss, "FP2_HISTORIQUES", FP2_HH); }
function _fp2Job_() { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty("FP2_JOB") || "{}"); } catch (e) { return {}; } }
function _fp2SetJob_(j) { j.maj = Date.now(); PropertiesService.getScriptProperties().setProperty("FP2_JOB", JSON.stringify(j)); }
function _fp2N_(s) { return String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim(); }
function _fp2Fmt_(d) { return d ? Utilities.formatDate(d, "Europe/Paris", "dd/MM/yyyy") : ""; }

// Journal en temps réel (500 dernières lignes)
function _fp2Log_(ss, type, msg) {
  try {
    var ws = _fp2Feuille_(ss, "FP2_JOURNAL", ["Date", "Type", "Message"]);
    ws.appendRow([Utilities.formatDate(new Date(), "Europe/Paris", "dd/MM HH:mm:ss"), type, String(msg).slice(0, 500)]);
    if (ws.getLastRow() > 700) ws.deleteRows(2, 200);
    SpreadsheetApp.flush();   // visible tout de suite dans la page de suivi
  } catch (e) {}
}

function _fp2Diplome_(formation) {
  var f = String(formation || "").toUpperCase().replace(/^AIX_?/, "");
  if (/BACHELOR/.test(f)) return "BACHELOR";
  if (/GPME/.test(f)) return "BTS GPME";
  if (/DSCG/.test(f)) return "DSCG";
  if (/DCG/.test(f)) return "DCG";
  if (/BTS|CG ?[12]/.test(f)) return "BTS CG";
  return "";
}

// ── Étudiants du CRM ─────────────────────────────────────────────────────────
function _fp2Etudiants_(ss) {
  var d = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues(), out = {};
  var h = d[0].map(function (x) { return String(x || "").trim().toLowerCase(); }), colCrea = h.indexOf(String(DATE_CREATION_HEADER).toLowerCase());
  for (var i = 1; i < d.length; i++) {
    var r = d[i], id = String(r[0] || "").trim(); if (!id || id.toUpperCase() === "ETU000X") continue;
    var typo = String(r[20] || "").trim() || String(calculerTypoEtu(r[17], r[18], r[14]));
    out[id] = { id: id, ligne: i + 1, nom: String(r[2] || "").trim(), prenom: String(r[3] || "").trim(), typo: typo,
      typoN: _fp2N_(typo).replace(/ /g, ""), classe: String(r[14] || "").trim(), campus: String(r[22] || "").trim(),
      statut: String(r[13] || "").trim(), entreprise: String(r[21] || "").trim(), ref: String(r[7] || "").trim(),
      idEleve: String(r[23] || "").trim(), mail: String(r[4] || "").trim(), entree: r[17], sortie: r[18], motifSortie: String(r[19] || "").trim(),
      crea: colCrea >= 0 ? r[colCrea] : "" };
  }
  return out;
}
function _fp2EstCible_(e) { return (e.typoN === "inscrit" || e.typoN === "preinscrit") && e.classe; }

// ── 1) INVENTAIRE ────────────────────────────────────────────────────────────
function _fp2DriveListe_(q) {
  var tok = ScriptApp.getOAuthToken(), H = { headers: { Authorization: "Bearer " + tok }, muteHttpExceptions: true }, out = [], pt = "";
  do {
    var url = "https://www.googleapis.com/drive/v3/files?pageSize=1000&corpora=allDrives&includeItemsFromAllDrives=true&supportsAllDrives=true&q=" + encodeURIComponent(q)
      + "&fields=" + encodeURIComponent("nextPageToken,files(id,name,mimeType,createdTime,parents,owners(emailAddress))") + (pt ? "&pageToken=" + pt : "");
    var r = JSON.parse(UrlFetchApp.fetch(url, H).getContentText());
    if (r.error) throw new Error("Drive : " + JSON.stringify(r.error).slice(0, 200));
    (r.files || []).forEach(function (f) { out.push(f); });
    pt = r.nextPageToken || "";
  } while (pt);
  return out;
}
function _fp2Dossier_(cache, id) {
  if (!id) return null;
  if (cache[id]) return cache[id];
  var sc = CacheService.getScriptCache(), k = "fp2d_" + id, c = sc.get(k);
  if (c) { try { return (cache[id] = JSON.parse(c)); } catch (e) {} }
  var r = {}; try { r = JSON.parse(UrlFetchApp.fetch("https://www.googleapis.com/drive/v3/files/" + id + "?supportsAllDrives=true&fields=name,parents",
    { headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }, muteHttpExceptions: true }).getContentText()); } catch (e) {}
  cache[id] = { id: id, nom: r.name || "", parent: (r.parents || [])[0] || "" };
  try { sc.put(k, JSON.stringify(cache[id]), 21600); } catch (e) {}
  return cache[id];
}

function _fp2Inventaire_(ss) {
  var etus = _fp2Etudiants_(ss), job = _fp2Job_();
  // fichiers à exclure : fiches générées le 10/11, classeurs FP2, modèles et documents génériques
  var exclus = {};
  try { var reg = ss.getSheetByName("FICHES_POSITIONNEMENT"); if (reg && reg.getLastRow() > 1) reg.getRange(2, 10, reg.getLastRow() - 1, 1).getValues().forEach(function (r) { if (r[0]) exclus[String(r[0]).trim()] = 1; }); } catch (e) {}
  var wsE = _fp2Etu_(ss), dE = wsE.getDataRange().getValues();
  for (var i = 1; i < dE.length; i++) if (dE[i][FP2_C.classeurId]) exclus[String(dE[i][FP2_C.classeurId])] = 1;
  var props = PropertiesService.getScriptProperties();
  [props.getProperty("FICHE_POS_MODELE_ID"), FP_MODELE_XLSX].forEach(function (x) { if (x) exclus[x] = 1; });

  var fichiers = _fp2DriveListe_("(name contains 'positionnement' or name contains 'positionement' or name contains 'positionnment' or (fullText contains 'POSITIONNEMENT CANDIDAT' and mimeType contains 'spreadsheet')) and trashed = false");
  fichiers = fichiers.filter(function (f) {
    if (exclus[f.id]) return false;
    if (/folder|shortcut|image\//.test(f.mimeType)) return false;
    if (/mod[eè]le|^mcp|conseil (de )?positionnement|capture|proc[ée]dure|^TMP lecture|^MODÈLE CRM/i.test(f.name)) return false;
    if (/^Fiches? de positionnement - /.test(f.name) && String(f.createdTime) >= "2026-11-10") return false;   // générées par le CRM
    return true;
  });
  _fp2Log_(ss, "Inventaire", fichiers.length + " fichier(s) candidat(s) trouvé(s) dans le Drive");

  // dossiers GED → étudiant
  var parDossier = {};
  try { getOrCreateGedEtuSheet(ss).getDataRange().getValues().slice(1).forEach(function (r) { if (r[0] && r[2]) parDossier[String(r[2]).trim()] = String(r[0]).trim(); }); } catch (e) {}
  // index nom → étudiants
  var liste = Object.keys(etus).map(function (k) { return etus[k]; });
  liste.forEach(function (e) { e._nom = _fp2N_(e.nom).split(" ").filter(function (t) { return t.length > 1; }); e._pre = _fp2N_(e.prenom).split(" ").filter(Boolean); });

  var cache = {}, rattach = {}, ambigus = 0, orphelins = 0, libres = [];
  fichiers.forEach(function (f) {
    var chaine = [], p = _fp2Dossier_(cache, (f.parents || [])[0]);
    for (var k = 0; k < 4 && p; k++) { chaine.push(p); p = p.parent ? _fp2Dossier_(cache, p.parent) : null; }
    var idEtu = "";
    chaine.some(function (d) { if (parDossier[d.id]) { idEtu = parDossier[d.id]; return true; } return false; });
    if (!idEtu) {
      var txt = " " + _fp2N_([f.name].concat(chaine.slice(0, 3).map(function (d) { return d.nom; })).join(" ")) + " ";
      var cands = liste.filter(function (e) { return e._nom.length && e._pre.length && e._nom.every(function (t) { return txt.indexOf(" " + t + " ") >= 0; }) && txt.indexOf(" " + e._pre[0] + " ") >= 0; });
      if (cands.length > 1) { var full = cands.filter(function (e) { return e._pre.every(function (t) { return txt.indexOf(" " + t + " ") >= 0; }); }); if (full.length) cands = full; }
      if (cands.length > 1) { var act = cands.filter(_fp2EstCible_); if (act.length === 1) cands = act; }
      if (cands.length === 1) idEtu = cands[0].id; else if (cands.length > 1) ambigus++; else orphelins++;
    }
    if (!idEtu) { if (!/^Fiches? de positionnement - /.test(f.name) || String(f.createdTime) < "2026-11-10") libres.push([f.id, f.name, String(f.createdTime).slice(0, 10), chaine.slice(0, 3).map(function (d) { return d.nom; }).reverse().join(" / ")]); return; }
    (rattach[idEtu] = rattach[idEtu] || []).push({ id: f.id, nom: f.name, mime: f.mimeType, cree: String(f.createdTime).slice(0, 10) });
  });

  // fichiers non rattachés : base des suggestions de redirection
  var wsL = _fp2Feuille_(ss, "FP2_NON_RATTACHES", ["ID fichier", "Nom", "Créé le", "Chemin"]);
  if (wsL.getLastRow() > 1) wsL.getRange(2, 1, wsL.getLastRow() - 1, 4).clearContent();
  if (libres.length) wsL.getRange(2, 1, libres.length, 4).setValues(libres);

  // registre FP2_HISTORIQUES (une ligne par fichier)
  var wsH = _fp2Histo_(ss), dH = wsH.getDataRange().getValues(), connus = {}, nouv = [];
  for (var j = 1; j < dH.length; j++) connus[String(dH[j][0])] = 1;
  Object.keys(rattach).forEach(function (k) { rattach[k].forEach(function (f) { if (!connus[f.id]) { connus[f.id] = 1; nouv.push([f.id, f.nom, f.mime, f.cree, "https://drive.google.com/file/d/" + f.id, "", "", "", "", "", ""]); } }); });
  if (nouv.length) wsH.getRange(wsH.getLastRow() + 1, 1, nouv.length, FP2_HH.length).setValues(nouv);

  // registre FP2_ETUDIANTS : cibles = inscrits + préinscrits (+ refus déjà décidés) ; décisions conservées
  var exist = {}; for (var i2 = 1; i2 < dE.length; i2++) if (dE[i2][0]) exist[String(dE[i2][0])] = { ligne: i2 + 1, r: dE[i2] };
  var maj = [], ajouts = [], st = { inscrits: 0, preinscrits: 0, avecHisto: 0, sansHisto: 0 };
  liste.forEach(function (e) {
    var x = exist[e.id], cible = _fp2EstCible_(e), refus = x && String(x.r[FP2_C.decision]) === "refus";
    if (!cible && !refus) return;
    var h = (rattach[e.id] || []).sort(function (a, b) { return a.cree < b.cree ? -1 : 1; });
    if (x && x.r[FP2_C.lien]) { try { var idL = String(x.r[FP2_C.lien]).match(/[-\w]{25,}/)[0]; if (!h.some(function (y) { return y.id === idL; })) h.push({ id: idL, nom: "Fiche indiquée (redirection)", mime: "", cree: "", redir: 1 }); } catch (eL) {} }
    var pre = e.typoN === "preinscrit";
    if (cible) { if (pre) st.preinscrits++; else st.inscrits++; if (h.length) st.avecHisto++; else st.sansHisto++; }
    var dec = x ? String(x.r[FP2_C.decision] || "") : "";
    if (!dec || dec === "attente") dec = (pre || h.length) ? "auto" : "attente";
    if (dec === "redirection" && !h.length) dec = "attente";
    var etat = x ? String(x.r[FP2_C.etat] || "") : "";
    if (etat !== "fait" && etat !== "en cours") etat = dec === "attente" ? "attente décision" : "à faire";
    var row = [e.id, e.nom, e.prenom, e.typo, e.classe, _fp2Diplome_(_fpFormation_(e.classe)), JSON.stringify(h.map(function (y) { return { id: y.id, nom: y.nom, cree: y.cree }; })),
      dec, x ? x.r[FP2_C.lien] : "", x ? x.r[FP2_C.noteRefus] : "", x ? x.r[FP2_C.motifRefus] : "", x ? x.r[FP2_C.classeurId] : "", x ? x.r[FP2_C.classeurUrl] : "",
      etat, x ? x.r[FP2_C.manquants] : "", x ? x.r[FP2_C.message] : "", new Date()];
    if (x) maj.push({ ligne: x.ligne, row: row }); else ajouts.push(row);
  });
  maj.forEach(function (m) { wsE.getRange(m.ligne, 1, 1, FP2_H.length).setValues([m.row]); });
  if (ajouts.length) wsE.getRange(wsE.getLastRow() + 1, 1, ajouts.length, FP2_H.length).setValues(ajouts);
  _fp2Log_(ss, "Inventaire", "Inscrits : " + st.inscrits + " · préinscrits : " + st.preinscrits + " · avec historique : " + st.avecHisto
    + " · sans historique : " + st.sansHisto + " · fichiers non rattachés : " + orphelins + " · ambigus (homonymes) : " + ambigus);
  return st;
}

// ── 2) EXTRACTION DES FICHES HISTORIQUES ─────────────────────────────────────
function _fp2DossierTravail_() {
  var props = PropertiesService.getScriptProperties(), id = props.getProperty("FP2_TRAVAIL");
  if (id) { try { var f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {} }
  var d = DriveApp.createFolder("CRM — Fiches de positionnement (copies de travail)");
  props.setProperty("FP2_TRAVAIL", d.getId());
  return d;
}
function _fp2Convertir_(idFichier, mime) {
  if (mime === MimeType.GOOGLE_SHEETS) return idFichier;
  if (!/spreadsheetml|ms-excel|opendocument\.spreadsheet/.test(mime)) return "";
  var r = UrlFetchApp.fetch("https://www.googleapis.com/drive/v3/files/" + idFichier + "/copy?supportsAllDrives=true", {
    method: "post", contentType: "application/json", muteHttpExceptions: true, headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    payload: JSON.stringify({ name: "Copie de travail — " + idFichier, mimeType: "application/vnd.google-apps.spreadsheet", parents: [_fp2DossierTravail_().getId()] }) });
  if (r.getResponseCode() !== 200) throw new Error("Conversion impossible (HTTP " + r.getResponseCode() + ")");
  return JSON.parse(r.getContentText()).id;
}

// Lecture par libellés (les anciennes fiches n'ont pas toutes la même mise en page)
function _fp2Lire_(sheetId) {
  var sh = SpreadsheetApp.openById(sheetId).getSheets()[0], v = sh.getRange(1, 1, Math.min(60, sh.getMaxRows()), Math.min(14, sh.getMaxColumns())).getDisplayValues();
  var D = { champs: {}, grille: [], cursus: {} }, N = _fp2N_, titre = false;
  var cell = function (r, c) { return (v[r] && v[r][c] != null) ? String(v[r][c]).trim() : ""; };
  // valeur à droite d'un libellé : zone identité = colonnes C-D seulement ; zone administration = 3 colonnes
  var droite = function (r, c, large) { var fin = large ? Math.min(c + 4, 14) : Math.min(4, 14); for (var k = c + 1; k < fin; k++) if (cell(r, k)) return cell(r, k); return ""; };
  var LIB = { "civilite": "civilite", "nom": "nom", "prenom": "prenom", "ville": "ville", "ville gare de depart": "ville", "age": "age", "date de naissance": "naissance",
    "nationalite": "nationalite", "autorisation de travail": "autorisation", "date fin titre de sejour": "titreFin", "benefice rqth": "rqth", "source": "source",
    "permis b": "permis", "vehicule": "vehicule", "langues parlees": "langues", "entreprise": "entreprise", "nom de la societe": "societe" };
  var CRIT = [/preparation/, /attitude/, /expression/, /argumentation|projet/, /adequation/, /bases? theorique|theorique/, /experience|pratique/, /resultat|^bulletins? scolaires/, /appreciation/];   // ancienne trame : une seule ligne « Bulletins scolaires » = résultats
  for (var r = 0; r < v.length; r++) for (var c = 0; c < v[r].length; c++) {
    var t = N(cell(r, c)); if (!t) continue;
    if (/fiche de positionn?e?ment/.test(t)) titre = true;
    if (LIB[t] && c <= 3 && D.champs[LIB[t]] == null) D.champs[LIB[t]] = droite(r, c);
    if (/intitule de la formation/.test(t)) { for (var r2 = r + 1; r2 < r + 4; r2++) if (cell(r2, c)) { D.formation = cell(r2, c); break; } }
    if (t === "statut" && !D.statut) D.statut = cell(r + 1, c) || droite(r, c);
    if (t === "campus" && !D.campus) { for (var r3 = r + 1; r3 < r + 4; r3++) if (cell(r3, c)) { D.campus = cell(r3, c); break; } }
    if (t === "dates" && !D.date) D.date = droite(r, c, 1);
    if (t === "interlocuteur" && !D.interlocuteur) D.interlocuteur = droite(r, c, 1);
    if (t === "resultat" && !D.note) { D.note = droite(r, c, 1); D.resultat = cell(r + 1, c + 1) || cell(r + 1, c); }
    if (t === "etudes en cours") D.cursus.encours = [cell(r, c + 1), cell(r, c + 2)];
    if (/^autre s diplome/.test(t)) D.cursus.autres = [cell(r, c + 1), cell(r, c + 2)];
    if (t === "points forts") D.forts = cell(r + 1, c);
    if (t === "axes d amelioration") D.axes = cell(r + 1, c);
    if (/^synthese/.test(t)) D.synthese = cell(r + 1, c);
    if (c >= 3 && c <= 5) for (var k = 0; k < 9; k++) if (CRIT[k].test(t) && !D.grille[k]) {
      var niv = -1; for (var q = 0; q < 4; q++) if (/x/i.test(cell(r, c + 1 + q))) niv = q;
      if (niv >= 0) D.grille[k] = { niveau: niv, obs: cell(r, c + 5) };
    }
  }
  if (!D.date) D.date = "";
  D.nbCriteres = D.grille.filter(Boolean).length;
  D.estFiche = titre || D.nbCriteres > 0;
  D.complete = D.nbCriteres === 9;
  return D;
}

function _fp2Extraire_(ss, deadline) {
  var ws = _fp2Histo_(ss), d = ws.getDataRange().getValues(), faits = 0;
  for (var i = 1; i < d.length && Date.now() < deadline; i++) {
    if (d[i][9] || d[i][10]) continue;   // déjà extrait ou en erreur
    var t0 = Date.now(), id = String(d[i][0]), mime = String(d[i][2] || "");
    try {
      if (!mime) { mime = DriveApp.getFileById(id).getMimeType(); ws.getRange(i + 1, 3).setValue(mime); }
      var conv = String(d[i][5] || "") || _fp2Convertir_(id, mime);
      if (!conv) { ws.getRange(i + 1, 10, 1, 2).setValues([[new Date(), "Document non tableur (" + mime.split(/[./]/).pop() + ") : gardé en lien"]]); continue; }
      var D = _fp2Lire_(conv);
      var dip = _fp2Diplome_(D.formation);
      ws.getRange(i + 1, 6, 1, 5).setValues([[conv, D.formation || "", dip, JSON.stringify(D), new Date()]]);
      _fp2Log_(ss, "Extraction", d[i][1] + " → " + (D.formation || "formation ?") + " · grille " + D.nbCriteres + "/9" + " (" + Math.round((Date.now() - t0) / 100) / 10 + " s)");
      faits++;
    } catch (e) { ws.getRange(i + 1, 10, 1, 2).setValues([[new Date(), String(e).slice(0, 200)]]); _fp2Log_(ss, "Erreur", "Extraction " + d[i][1] + " : " + e); }
    _fp2Mesure_("extraction", Date.now() - t0);
  }
  return faits;
}

function _fp2EstFiche_(D) { if (!D) return false; return D.estFiche != null ? !!D.estFiche : ((D.grille || []).filter(Boolean).length > 0 || !!D.formation); }

// ── Sources complémentaires ──────────────────────────────────────────────────
function _fp2Charlemagne_(ss) {
  var out = { parId: {}, parNom: {} };
  try {
    var ws = ss.getSheetByName("CHARLEMAGNE_IMPORT"); if (!ws || ws.getLastRow() < 3) return out;
    var d = ws.getDataRange().getValues(), h = d[1].map(String);
    for (var i = 2; i < d.length; i++) {
      var o = {}; h.forEach(function (k, j) { o[k] = d[i][j]; });
      if (o.idEleve) out.parId[String(o.idEleve).trim()] = o;
      out.parNom[_fp2N_(o.nom + " " + o.prenom)] = o;
    }
  } catch (e) {}
  return out;
}

// OCR (Drive) des pièces d'identité / dossiers d'inscription du dossier GED
function _fp2Ocr_(dossier) {
  var res = {}, n = 0;
  try {
    var it = dossier.getFiles(), cands = [];
    while (it.hasNext()) { var f = it.next(), nm = _fp2N_(f.getName()), mt = f.getMimeType();
      if (!/pdf|image\/(jpe?g|png)/.test(mt)) continue;
      if (/\bcni\b|carte d identite|identite|passeport|titre de sejour|piece|inscription|dossier d admission|candidature/.test(nm)) cands.push(f); }
    var subs = dossier.getFolders();
    while (subs.hasNext() && cands.length < 6) { var s = subs.next(); if (!/administratif|admin|inscription/i.test(s.getName())) continue;
      var it2 = s.getFiles(); while (it2.hasNext()) { var f2 = it2.next(), nm2 = _fp2N_(f2.getName());
        if (/pdf|image\/(jpe?g|png)/.test(f2.getMimeType()) && /\bcni\b|carte d identite|identite|passeport|titre de sejour|piece|inscription/.test(nm2)) cands.push(f2); } }
    cands.slice(0, 3).forEach(function (f) {
      if (res.naissance && res.nationalite && res.ville) return;
      var r = UrlFetchApp.fetch("https://www.googleapis.com/drive/v3/files/" + f.getId() + "/copy?supportsAllDrives=true&ocrLanguage=fr", {
        method: "post", contentType: "application/json", muteHttpExceptions: true, headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
        payload: JSON.stringify({ name: "OCR temporaire", mimeType: "application/vnd.google-apps.document", parents: [_fp2DossierTravail_().getId()] }) });
      if (r.getResponseCode() !== 200) return;
      var docId = JSON.parse(r.getContentText()).id, txt = "";
      try { txt = UrlFetchApp.fetch("https://www.googleapis.com/drive/v3/files/" + docId + "/export?mimeType=text/plain", { headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }, muteHttpExceptions: true }).getContentText(); } catch (e) {}
      try { DriveApp.getFileById(docId).setTrashed(true); } catch (e) {}
      n++;
      var m;
      if (!res.naissance && (m = txt.match(/n[ée]+\s*\(?e?\)?\s*le\s*[:\s]*(\d{1,2})[\s.\/-](\d{1,2})[\s.\/-](\d{4})/i) || txt.match(/date de naiss\w*\s*[:\s\/]*(\d{1,2})[\s.\/-](\d{1,2})[\s.\/-](\d{4})/i)))
        res.naissance = { v: ("0" + m[1]).slice(-2) + "/" + ("0" + m[2]).slice(-2) + "/" + m[3], src: f.getName() };
      if (!res.nationalite && (m = txt.match(/nationalit[ée]\s*[:\/]?\s*([A-Za-zÀ-ÿ\-]{4,20})/i))) res.nationalite = { v: m[1].charAt(0).toUpperCase() + m[1].slice(1).toLowerCase(), src: f.getName() };
      if (!res.nationalite && /REPUBLIQUE FRAN[CÇ]AISE|CARTE NATIONALE D.IDENTIT/i.test(txt)) res.nationalite = { v: "Française", src: f.getName() };
      if (!res.ville && (m = txt.match(/\b(\d{5})\s+([A-ZÀ-Ÿ][A-ZÀ-Ÿa-zà-ÿ' \-]{2,30})/))) res.ville = { v: m[2].trim().replace(/\s+(CEDEX.*)$/i, ""), src: f.getName() };
    });
  } catch (e) {}
  res._lus = n;
  return res;
}

// ── 3) GÉNÉRATION ────────────────────────────────────────────────────────────
function _fp2Mesure_(phase, ms) {
  var j = _fp2Job_(); j.durees = j.durees || {}; var x = j.durees[phase] || [0, 0];
  j.durees[phase] = [x[0] + ms, x[1] + 1]; j.fait = (j.fait || 0) + 1; _fp2SetJob_(j);
}

function _fp2Generer_(ss, row, ctx) {
  var e = ctx.etus[String(row[FP2_C.id])];
  if (!e) throw new Error("Étudiant introuvable dans ETUDIANTS");
  var dec = String(row[FP2_C.decision]), refus = dec === "refus", pre = e.typoN === "preinscrit";
  var histos = []; try { histos = JSON.parse(row[FP2_C.histo] || "[]"); } catch (x) {}
  histos = histos.map(function (h) { var H = ctx.histo[h.id] || {}; return { id: h.id, nom: h.nom, cree: h.cree, conv: H.conv, D: H.D, dip: H.dip, formation: H.formation, doc: H.doc, lien: "https://drive.google.com/file/d/" + h.id }; })
    .filter(function (h) { return _fp2EstFiche_(h.D) || (h.doc && /positionn?e?ment/i.test(h.nom)); });   // documents qui ne sont pas des fiches : écartés
  histos.sort(function (a, b) { return String(b.cree).localeCompare(String(a.cree)); });   // plus récent d'abord
  var recent = histos.filter(function (h) { return h.D; })[0];
  var champ = function (k) { for (var i = 0; i < histos.length; i++) { var D = histos[i].D; if (D && D.champs && D.champs[k]) return D.champs[k]; } return ""; };

  var dossier = _fpDossierGed_(ss, e.id, e.nom, e.prenom);
  var ocr = null, notes = {};
  var charl = ctx.charl.parId[e.idEleve] || ctx.charl.parNom[_fp2N_(e.nom + " " + e.prenom)] || {};
  var naissance = charl.naissance ? _fp2Fmt_(_charlDate_(charl.naissance)) : (champ("naissance") || "");
  var nationalite = champ("nationalite"), ville = champ("ville");
  if (!naissance || !nationalite || !ville) {
    ocr = _fp2Ocr_(dossier);
    if (!naissance && ocr.naissance) { naissance = ocr.naissance.v; notes.naissance = "Lu par OCR sur « " + ocr.naissance.src + " » — à vérifier"; }
    if (!nationalite && ocr.nationalite) { nationalite = ocr.nationalite.v; notes.nationalite = "Lu par OCR sur « " + ocr.nationalite.src + " » — à vérifier"; }
    if (!ville && ocr.ville) { ville = ocr.ville.v; notes.ville = "Lu par OCR sur « " + ocr.ville.src + " » — à vérifier"; }
  }
  var civ = champ("civilite");
  var base = { civilite: civ, nom: e.nom.toUpperCase(), prenom: e.prenom, ville: ville, naissance: naissance, nationalite: nationalite,
    autorisation: champ("autorisation") || (/fran[cç]ais/i.test(nationalite) ? "Oui" : ""), permis: champ("permis"), vehicule: champ("vehicule"), langues: champ("langues") };

  // onglets nouvelle trame
  var dipActuel = _fp2Diplome_(_fpFormation_(e.classe)) || "AUTRE";
  var onglets = [];
  if (refus) onglets.push({ dip: dipActuel, actuel: true, refus: true });
  else if (pre) onglets.push({ dip: dipActuel, actuel: true });
  else {
    var vus = {};
    histos.slice().reverse().forEach(function (h) { if (h.dip && h.dip !== dipActuel && !vus[h.dip]) { vus[h.dip] = 1; onglets.push({ dip: h.dip, actuel: false }); } });
    onglets.push({ dip: dipActuel, actuel: true });
  }

  var nomF = "Fiche de positionnement - " + e.nom.toUpperCase() + " " + e.prenom;
  var classeur = SpreadsheetApp.create(nomF), fichier = DriveApp.getFileById(classeur.getId());
  fichier.moveTo(dossier);
  var premier = classeur.getSheets()[0], manq = [], creees = [];
  try {
    // copies fidèles des fiches historiques (onglets « base »)
    var docs = [];
    if (!pre || refus) histos.slice().reverse().forEach(function (h, k) {
      if (h.conv) {
        try { var sh = SpreadsheetApp.openById(h.conv).getSheets()[0].copyTo(classeur);
          sh.setName(("Historique " + (h.cree || "").slice(0, 4) + " — " + (h.formation || h.dip || "fiche")).slice(0, 90) + (k ? " (" + (k + 1) + ")" : ""));
          sh.getRange("A1").setNote("Copie fidèle de la fiche historique « " + h.nom + " » (" + h.lien + "). L'original reste inchangé.");
        } catch (eC) { docs.push(h); }
      } else docs.push(h);
    });
    if (docs.length) {
      var sd = classeur.insertSheet("Historique — documents");
      sd.getRange(1, 1, 1, 3).setValues([["Fiche historique (document non tableur)", "Créée le", "Lien"]]).setFontWeight("bold").setBackground("#1A4E8A").setFontColor("#fff");
      sd.getRange(2, 1, docs.length, 3).setValues(docs.map(function (h) { return [h.nom, h.cree, h.lien]; }));
      sd.autoResizeColumns(1, 3);
    }
    // nouvelle trame
    onglets.forEach(function (o) {
      var src = o.actuel ? null : histos.filter(function (h) { return h.dip === o.dip && h.D; })[0];
      if (o.actuel) src = histos.filter(function (h) { return h.dip === o.dip && h.D; })[0] || null;
      var srcDouteuse = false;
      if (o.actuel && !src) { src = histos.filter(function (h) { return !h.dip && h.D && _fp2EstFiche_(h.D); })[0] || null; srcDouteuse = !!src; }   // fiche historique sans formation lisible
      var r = _fp2Onglet_(classeur, ctx, e, o, src, base, notes);
      r.manquants.forEach(function (m) { manq.push(o.dip + " : " + m); });
      if (srcDouteuse) manq.push(o.dip + " : Grille reprise d'une fiche historique sans formation indiquée — à vérifier");
      creees.push(r.sheet);
      if (o.dip === "BACHELOR" && !o.refus) { var rb = _fp2OngletBachelor_(classeur, ctx); creees.push(rb); manq.push("BACHELOR : notes des critères d'admission pondérés"); }
    });
    classeur.deleteSheet(premier);
    // ordre : historiques d'abord, nouvelle trame ensuite, onglet actif = fiche du diplôme actuel
    classeur.setActiveSheet(creees[creees.length - 1] && /Critères/.test(creees[creees.length - 1].getName()) ? creees[creees.length - 2] : creees[creees.length - 1]);
    SpreadsheetApp.flush();
  } catch (eG) { try { fichier.setTrashed(true); } catch (x) {} throw eG; }

  // fiche du 10/11 remplacée → corbeille
  try {
    var reg = ss.getSheetByName("FICHES_POSITIONNEMENT");
    if (reg && reg.getLastRow() > 1) {
      var dr = reg.getRange(2, 1, reg.getLastRow() - 1, 12).getValues();
      dr.forEach(function (x, k) {
        if (String(x[0]).trim() === e.id && x[9] && !/remplacée/.test(String(x[11]))) {
          try { DriveApp.getFileById(String(x[9])).setTrashed(true); } catch (eT) {}
          reg.getRange(k + 2, 12).setValue("Remplacée le " + _fp2Fmt_(new Date()) + " (refonte — corbeille Drive)");
        }
      });
    }
  } catch (eR) {}
  if (refus) {
    try { getSheetSafe(ss, "ETUDIANTS").getRange(e.ligne, 21).setValue("REFUSÉ"); _invalidateStudentDetailsCache_(e.id); } catch (eU) {}
  }
  return { id: classeur.getId(), url: classeur.getUrl(), manquants: manq, onglets: onglets.length, historiques: histos.length };
}

// Un onglet « nouvelle trame » (copie du modèle HECG)
function _fp2Onglet_(classeur, ctx, e, o, src, base, notes) {
  var M = ctx.modele, sh = SpreadsheetApp.openById(M.id).getSheets()[0].copyTo(classeur);
  var annee = new Date().getMonth() >= 7 ? new Date().getFullYear() : new Date().getFullYear() - 1;
  var nomO = o.refus ? "Fiche de refus" : (o.dip === "AUTRE" ? "Fiche" : o.dip) + " — nouvelle trame";
  sh.setName(nomO);
  var D = src ? src.D : null, V = M.valeurs.map(function (l) { return l.slice(); }), manq = [];
  var set = function (a1, v) { var col = a1.charCodeAt(0) - 66, row = Number(a1.slice(1)) - 2; V[row][col] = v; };
  var formation = o.actuel ? _fpFormation_(e.classe) : (src && src.formation) || o.dip;
  var aix = o.actuel ? (e.campus.toUpperCase() === "HECG_AIX" || /AIX/i.test(e.classe)) : /aix/i.test(D && D.campus || "");
  var alternance = o.actuel ? /altern|contrat/i.test(e.statut) : /contrat|altern|apprent/i.test(D && D.statut || "");
  var refId = _referentPourClasse_(e.classe, e.campus) || e.ref;
  set("B4", "Date de rédaction : " + _fp2Fmt_(new Date()));
  set("B7", formation); set("E7", aix ? "Aix" : "MARSEILLE");
  set("D8", alternance ? "Contrat d'apprentissage" : "Initiale");
  set("H7", D && D.date ? D.date : ""); set("H8", D && D.interlocuteur ? D.interlocuteur : "");
  set("L8", o.refus ? "Refus" : "Admis");
  set("C11", base.civilite); set("C12", base.nom); set("C13", base.prenom); set("C14", base.ville); set("C15", base.naissance);
  set("C16", ""); set("C17", base.nationalite); set("C18", base.autorisation); set("C19", "");
  set("C22", base.permis); set("C23", base.vehicule); set("C24", base.langues);
  var ent = o.actuel ? e.entreprise : (D && D.champs && D.champs.societe) || "";
  set("C25", alternance || ent ? "Oui" : "Non"); set("C26", alternance ? ent : "");
  var enc = D && D.cursus && D.cursus.encours, aut = D && D.cursus && D.cursus.autres;
  if (o.actuel) { set("C29", formation); set("D29", annee + "-" + (annee + 1)); } else { set("C29", enc ? enc[0] : ""); set("D29", enc ? enc[1] : ""); }
  set("C30", aut ? aut[0] : (o.actuel && !o.refus && src && enc ? enc[0] : "")); set("D30", aut ? aut[1] : (o.actuel && src && enc ? enc[1] : ""));
  for (var k = 0; k < 9; k++) {
    var g = D && D.grille && D.grille[k];
    ["F", "G", "H", "I"].forEach(function (c, lv) { set(c + (13 + k), g && g.niveau === lv ? "X" : ""); });
    set("J" + (13 + k), g ? g.obs : "");
  }
  set("E26", D && D.forts ? D.forts : ""); set("H26", D && D.axes ? D.axes : "");
  var synth = D && D.synthese ? D.synthese : "";
  if (o.refus) synth = "Candidature non retenue." + (e._motifRefus ? " Motif : " + e._motifRefus + "." : "") + (synth ? " " + synth : "");
  set("B32", synth);
  // note : calculée seulement quand la grille est complète (sinon vide, pas 0)
  var fNote = String(M.formuleNote || "");
  sh.getRange(2, 2, V.length, V[0].length).setValues(V);
  // séparateurs d'arguments selon la langue du classeur (« ; » en français)
  if (fNote) sh.getRange("F22").setFormula(_fp2F_(classeur, "=IF(COUNTIF(F13:I21,\"X\")=0,\"\"," + fNote.replace(/^=/, "") + ")"));
  sh.setRowHeight(4, 20);
  if (o.refus) {
    var n = Number(String(e._noteRefus).replace(",", "."));
    sh.getRange("L8").setValue("Refus").setFontColor("#c62828").setFontWeight("bold");   // valeur de la liste du modèle (Admis / Refus)
    sh.getRange("L7").setValue(isNaN(n) ? "" : n).setNumberFormat("0.00");
    sh.getRange("F22").setValue(isNaN(n) ? "" : n).setNumberFormat("0.00");
    sh.getRange("E23").setValue("Note inférieure à 10/20 : candidature non retenue");
  }
  sh.getRange("B4").setFontWeight("bold").setFontColor("#1A4E8A").setHorizontalAlignment("right");
  var cn = { naissance: "C15", nationalite: "C17", ville: "C14" };
  Object.keys(notes).forEach(function (k) { sh.getRange(cn[k]).setNote(notes[k]).setBackground("#fff4e5"); });
  // mentions manquantes
  var L = [["C11", "Civilité"], ["C14", "Ville"], ["C15", "Date de naissance"], ["C17", "Nationalité"], ["C18", "Autorisation de travail"],
    ["C22", "Permis B"], ["C23", "Véhicule"], ["C24", "Langues parlées"], ["H7", "Date de l'entretien"], ["H8", "Interlocuteur"],
    ["E26", "Points forts"], ["H26", "Axes d'amélioration"], ["B32", "Synthèse de l'entretien"]];
  var get = function (a1) { var col = a1.charCodeAt(0) - 66, row = Number(a1.slice(1)) - 2; return String(V[row][col] || "").trim(); };
  L.forEach(function (x) { if (!get(x[0])) manq.push(x[1]); });
  var nbG = D && D.grille ? D.grille.filter(Boolean).length : 0;
  if (!o.refus && nbG === 0) manq.push("Grille d'évaluation (9 critères)");
  else if (!o.refus && nbG < 9) {
    var NOMS = ["Préparation de l'entretien", "Attitude générale", "Expression verbale", "Argumentation métier/projet", "Adéquation du profil", "Connaissances théoriques", "Connaissances pratiques", "Résultats scolaires", "Appréciations sur bulletins"];
    var vides = NOMS.filter(function (n, k) { return !(D.grille[k]); });
    manq.push("Grille : " + vides.join(", ") + (vides.length === 1 && !D.grille[8] ? " (absent de l'ancienne trame)" : ""));
  }
  if (o.refus && !(D && D.grille && D.grille.filter(Boolean).length)) manq.push("Grille d'évaluation (justifiant le refus)");
  Object.keys(notes).forEach(function (k) { manq.push({ naissance: "Date de naissance", nationalite: "Nationalité", ville: "Ville" }[k] + " (lue par OCR, à vérifier)"); });
  return { sheet: sh, manquants: manq };
}

// Formule écrite en syntaxe anglaise → séparateurs « ; » si le classeur est dans une langue à virgule décimale
function _fp2F_(classeur, f) {
  var loc = ""; try { loc = String(classeur.getSpreadsheetLocale() || ""); } catch (e) {}
  if (!/^(fr|de|es|it|pt|nl|ru|pl|tr|sv|da|fi|nb|no|cs|el|hu|ro|uk|id|vi)/i.test(loc)) return f;
  var out = "", q = false;
  for (var i = 0; i < f.length; i++) { var ch = f.charAt(i); if (ch === '"') q = !q; out += (!q && ch === ",") ? ";" : ch; }
  return out;
}

// ── BACHELOR CDSP : critères pondérés ───────────────────────────────────────
var FP2_BACH_DEFAUT = [
  ["Résultat au test d'admission (QCM, rédaction, calcul)", "Tous blocs", 20],
  ["Bases de la paie (bulletin, cotisations, net à payer)", "Réaliser la paie et les déclarations sociales", 15],
  ["Droit du travail et administration du personnel", "Gérer l'administration du personnel", 15],
  ["Rigueur et aisance calculatoire", "Contrôler et fiabiliser la paie", 10],
  ["Outils : Excel, logiciel de paie, SIRH, DSN", "Utiliser les outils SIRH et de paie", 10],
  ["Expression écrite professionnelle", "Communiquer avec les salariés et les organismes", 10],
  ["Communication orale, posture RH et confidentialité", "Accompagner et conseiller les salariés", 10],
  ["Projet professionnel en RH / paie et motivation", "Projet", 5],
  ["Expérience en entreprise (RH, paie, cabinet)", "Projet", 5]
];
function _fp2BachCriteres_() { try { var x = JSON.parse(PropertiesService.getScriptProperties().getProperty("FP2_BACH_CRITERES") || "null"); if (x && x.length) return x; } catch (e) {} return FP2_BACH_DEFAUT; }
function _fp2OngletBachelor_(classeur, ctx) {
  var C = _fp2BachCriteres_(), sh = classeur.insertSheet("Critères d'admission Bachelor"), n = C.length;
  sh.getRange("B2").setValue("BACHELOR CDSP — CRITÈRES D'ADMISSION PONDÉRÉS").setFontSize(15).setFontWeight("bold").setFontColor("#1A4E8A");
  sh.getRange("B3").setValue("Chaque critère est noté sur 20 par le jury ; la note finale est la moyenne pondérée. Admission si la note finale est supérieure ou égale à 10/20.").setFontColor("#3b4859");
  sh.getRange(5, 2, 1, 5).setValues([["Critère", "Compétence visée (Bachelor CDSP — RH et paie)", "Poids (%)", "Note /20", "Note pondérée"]])
    .setFontWeight("bold").setBackground("#1A4E8A").setFontColor("#ffffff");
  sh.getRange(6, 2, n, 4).setValues(C.map(function (c) { return [c[0], c[1], c[2], ""]; }));
  sh.getRange(6, 6, n, 1).setFormulas(C.map(function (c, i) { return [_fp2F_(classeur, "=IF(E" + (6 + i) + "=\"\",\"\",E" + (6 + i) + "*D" + (6 + i) + "/100)")]; }));
  sh.getRange(6, 5, n, 1).setBackground("#fff8e6");
  var t = 6 + n;
  sh.getRange(t, 2, 1, 5).setFontWeight("bold").setBackground("#e8eff8");
  sh.getRange(t, 2).setValue("NOTE FINALE /20");
  sh.getRange(t, 4).setFormula("=SUM(D6:D" + (t - 1) + ")");
  sh.getRange(t, 6).setFormula(_fp2F_(classeur, "=IF(COUNT(E6:E" + (t - 1) + ")<" + n + ",\"\",SUM(F6:F" + (t - 1) + ")*100/D" + t + ")"));
  sh.getRange(t + 1, 2, 1, 2).setFontWeight("bold");
  sh.getRange(t + 1, 2).setValue("Décision");
  sh.getRange(t + 1, 3).setFormula(_fp2F_(classeur, "=IF(F" + t + "=\"\",\"À compléter\",IF(F" + t + ">=10,\"Admis\",\"Refusé\"))"));
  var tests = PropertiesService.getScriptProperties().getProperty("FP2_BACH_TESTS");
  if (tests) { try { var T = JSON.parse(tests); sh.getRange(t + 3, 2, 3, 2).setValues([["Test d'admission — sujet", T.sujet], ["Test d'admission — corrigé (jury uniquement)", T.corrige], ["Grille de notation du test", "QCM /8 · Rédaction /6 · Calcul /6 = /20 → reporter la note dans le 1er critère"]]); } catch (e) {} }
  sh.setColumnWidth(2, 360); sh.setColumnWidth(3, 330); sh.setColumnWidths(4, 3, 110);
  sh.getRange(6, 6, n + 1, 1).setNumberFormat("0.00"); sh.getRange(t, 6).setNumberFormat("0.00");
  sh.setHiddenGridlines(true);
  return sh;
}

// ── Moteur ───────────────────────────────────────────────────────────────────
function _fp2Ctx_(ss) {
  var M = _fpModeleCharge_(ss), shM = SpreadsheetApp.openById(M.id).getSheets()[0], f = shM.getRange("F22").getFormula();
  M.formuleNote = f;
  var hd = _fp2Histo_(ss).getDataRange().getValues(), histo = {};
  for (var i = 1; i < hd.length; i++) { var D = null; try { D = hd[i][8] ? JSON.parse(hd[i][8]) : null; } catch (e) {}
    histo[String(hd[i][0])] = { conv: String(hd[i][5] || ""), formation: String(hd[i][6] || ""), dip: String(hd[i][7] || ""), D: D, doc: /non tableur/.test(String(hd[i][10] || "")) }; }
  return { modele: M, etus: _fp2Etudiants_(ss), charl: _fp2Charlemagne_(ss), histo: histo };
}

function _fp2GenererLot_(ss, deadline) {
  var ws = _fp2Etu_(ss), d = ws.getDataRange().getValues(), ctx = null, faits = 0;
  for (var i = 1; i < d.length && Date.now() < deadline - 25000; i++) {
    if (String(d[i][FP2_C.etat]) !== "à faire" && String(d[i][FP2_C.etat]) !== "en cours") continue;
    if (_fp2Job_().pause) break;
    ctx = ctx || _fp2Ctx_(ss);
    var t0 = Date.now(), id = String(d[i][0]);
    // un passage précédent s'est arrêté au milieu de cet étudiant : 2 tentatives maximum
    if (String(d[i][FP2_C.etat]) === "en cours") {
      var pE = PropertiesService.getScriptProperties(), ess = {}; try { ess = JSON.parse(pE.getProperty("FP2_ESSAIS") || "{}"); } catch (x) {}
      ess[id] = (ess[id] || 0) + 1; pE.setProperty("FP2_ESSAIS", JSON.stringify(ess));
      if (ess[id] >= 2) {
        ws.getRange(i + 1, FP2_C.etat + 1, 1, 4).setValues([["erreur", d[i][FP2_C.manquants], "Génération interrompue deux fois (dossier trop volumineux ?) — relancez avec « Refaire »", new Date()]]);
        _fp2Log_(ss, "Erreur", d[i][1] + " " + d[i][2] + " : génération interrompue deux fois, étudiant mis de côté"); continue;
      }
    }
    ws.getRange(i + 1, FP2_C.etat + 1).setValue("en cours"); SpreadsheetApp.flush();
    try {
      var e = ctx.etus[id]; if (e) { e._noteRefus = d[i][FP2_C.noteRefus]; e._motifRefus = d[i][FP2_C.motifRefus]; }
      if (e && String(d[i][FP2_C.decision]) === "auto" && e.typoN === "inscrit") {
        var hv = []; try { hv = JSON.parse(d[i][FP2_C.histo] || "[]"); } catch (x) {}
        var valides = hv.filter(function (h) { var H = ctx.histo[h.id] || {}; return _fp2EstFiche_(H.D) || (H.doc && /positionn?e?ment/i.test(h.nom)); });
        if (!valides.length) {
          ws.getRange(i + 1, FP2_C.decision + 1).setValue("attente"); ws.getRange(i + 1, FP2_C.etat + 1).setValue("attente décision");
          ws.getRange(i + 1, FP2_C.message + 1).setValue("Documents trouvés mais aucun n'est une fiche de positionnement");
          _fp2Log_(ss, "Décision", d[i][1] + " " + d[i][2] + " : aucun des documents rattachés n'est une fiche → à décider");
          continue;
        }
      }
      // ancien classeur FP2 (régénération) → corbeille
      if (d[i][FP2_C.classeurId]) { try { DriveApp.getFileById(String(d[i][FP2_C.classeurId])).setTrashed(true); } catch (x) {} }
      var r = _fp2Generer_(ss, d[i], ctx);
      ws.getRange(i + 1, FP2_C.classeurId + 1, 1, 6).setValues([[r.id, r.url, "fait", JSON.stringify(r.manquants), r.historiques + " historique(s), " + r.onglets + " onglet(s) nouvelle trame", new Date()]]);
      var jp = _fp2Job_(), propsP = PropertiesService.getScriptProperties();
      if (propsP.getProperty("FP2_PILOTE") === "1" || jp.pilote) { propsP.deleteProperty("FP2_PILOTE"); jp.pilote = false; jp.pause = true; jp.piloteUrl = r.url; jp.piloteNom = d[i][1] + " " + d[i][2]; _fp2SetJob_(jp);
        _fp2Log_(ss, "Pilote", "Fiche pilote créée pour " + jp.piloteNom + " — vérifiez-la puis cliquez sur « Reprendre »"); }
      _fp2Log_(ss, "Fiche", d[i][1] + " " + d[i][2] + " — " + r.historiques + " historique(s), " + r.onglets + " onglet(s), " + r.manquants.length + " mention(s) manquante(s) (" + Math.round((Date.now() - t0) / 100) / 10 + " s)");
      faits++;
    } catch (eG) {
      ws.getRange(i + 1, FP2_C.etat + 1, 1, 4).setValues([["erreur", d[i][FP2_C.manquants], String(eG).slice(0, 300), new Date()]]);
      _fp2Log_(ss, "Erreur", d[i][1] + " " + d[i][2] + " : " + eG);
    }
    _fp2Mesure_("generation", Date.now() - t0);
  }
  return faits;
}

// Passage de travail : appelé par auditAdressesTick (toutes les 5 min) et, tant que la page de suivi
// est ouverte, relancé par elle dès que le précédent est fini. Un seul passage à la fois (bail de 5 min 30).
function fp2Tick(budget) {
  var t0 = Date.now(), props = PropertiesService.getScriptProperties();
  var lock = LockService.getScriptLock(); if (!lock.tryLock(3000)) return;
  try {
    var bail = Number(props.getProperty("FP2_BAIL") || 0);
    if (bail && t0 - bail < 330000) return;
    props.setProperty("FP2_BAIL", String(t0));
  } finally { lock.releaseLock(); }
  var ss = SpreadsheetApp.openById(FP2_SS), job = _fp2Job_();
  try {
    if (!job.actif || job.pause) return;
    var deadline = t0 + Math.min(280000, budget || 270000);
    job.dernierPassage = t0; _fp2SetJob_(job);
    if (job.phase === "inventaire") {
      _fp2Inventaire_(ss); job = _fp2Job_(); job.phase = "extraction"; job.phaseDebut = Date.now(); _fp2SetJob_(job);
    }
    if (job.phase === "extraction" && Date.now() < deadline) {
      _fp2Extraire_(ss, deadline);
      var reste = _fp2Histo_(ss).getDataRange().getValues().slice(1).filter(function (r) { return !r[9] && !r[10]; }).length;
      if (!reste) { job = _fp2Job_(); job.phase = "generation"; job.phaseDebut = Date.now(); job.fait = 0; _fp2SetJob_(job); _fp2Log_(ss, "Phase", "Extraction terminée → génération des fiches"); }
    }
    job = _fp2Job_();
    if (job.phase === "generation" && Date.now() < deadline) {
      _fp2GenererLot_(ss, deadline);
      var rest = _fp2Etu_(ss).getDataRange().getValues().slice(1).filter(function (r) { return r[FP2_C.etat] === "à faire" || r[FP2_C.etat] === "en cours"; }).length;
      if (!rest) {
        job = _fp2Job_(); job.phase = "termine"; job.actif = false; job.fin = Date.now(); _fp2SetJob_(job);
        _fp2Log_(ss, "Phase", "Génération terminée ✓");
        logAction(FP2_MAITRE, "", "Création", "Étudiants — Fiches de positionnement (refonte)", "FP2_ETUDIANTS", "Génération terminée");
        _fp2Declencheur_(false);
      }
    }
  } catch (e) { _fp2Log_(ss, "Erreur", "Passage : " + e); }
  finally { props.deleteProperty("FP2_BAIL"); }
}
// Pas de déclencheur dédié (limite Google de 20 atteinte) : le moteur vit dans auditAdressesTick.
function _fp2Declencheur_(on) {}
function _fp2Relancer_(phase) {
  var job = _fp2Job_();
  if (phase === "inventaire" || !job.phase || job.phase === "termine") { job.phase = phase || "inventaire"; job.debut = Date.now(); job.phaseDebut = Date.now(); job.fait = 0; }
  if (job.piloteFait !== true) { job.pilote = true; job.piloteFait = true; }   // 1re génération : une seule fiche, puis pause pour contrôle
  else if (phase && phase !== job.phase && !(phase === "generation" && job.phase !== "termine")) { job.phase = phase; job.phaseDebut = Date.now(); job.fait = 0; }
  job.actif = true; job.pause = false; _fp2SetJob_(job);
  PropertiesService.getScriptProperties().setProperty("FICHES_POS_ACTIVE", "0");   // l'ancien générateur (10/11) est coupé
  _fp2Declencheur_(true);
}

// ── API (compte maître) ──────────────────────────────────────────────────────
function handleFp2(p, ss) {
  if (!_fp2Autorise_(p)) return createJsonResponse({ success: false, message: "Espace réservé au compte maître." });
  try {
    var op = String(p.op || "etat"), ws = _fp2Etu_(ss);
    if (op === "etat") {
      var d = ws.getDataRange().getValues().slice(1), cpt = {}, dec = {}, manq = 0;
      d.forEach(function (r) { cpt[r[FP2_C.etat]] = (cpt[r[FP2_C.etat]] || 0) + 1; dec[r[FP2_C.decision]] = (dec[r[FP2_C.decision]] || 0) + 1; if (r[FP2_C.etat] === "fait") { try { if (JSON.parse(r[FP2_C.manquants] || "[]").length) manq++; } catch (e) {} } });
      var hd = _fp2Histo_(ss).getDataRange().getValues().slice(1), hx = { total: hd.length, faits: hd.filter(function (r) { return r[9] && !r[10]; }).length, erreurs: hd.filter(function (r) { return r[10]; }).length };
      var jr = ss.getSheetByName("FP2_JOURNAL"), journal = [];
      if (jr && jr.getLastRow() > 1) { var n = Math.min(80, jr.getLastRow() - 1); journal = jr.getRange(jr.getLastRow() - n + 1, 1, n, 3).getDisplayValues().reverse(); }
      var job = _fp2Job_(), trig = ScriptApp.getProjectTriggers().some(function (x) { return x.getHandlerFunction() === "auditAdressesTick"; });
      return createJsonResponse({ success: true, job: job, etats: cpt, decisions: dec, histo: hx, avecManquants: manq, journal: journal, declencheur: trig,
        bail: Number(PropertiesService.getScriptProperties().getProperty("FP2_BAIL") || 0), maintenant: Date.now() });
    }
    if (op === "liste") {
      var libresL = []; try { var wl = ss.getSheetByName("FP2_NON_RATTACHES"); if (wl && wl.getLastRow() > 1) libresL = wl.getRange(2, 1, wl.getLastRow() - 1, 4).getValues().map(function (x) { return { id: x[0], nom: String(x[1]), cree: x[2] instanceof Date ? Utilities.formatDate(x[2], "Europe/Paris", "yyyy-MM-dd") : String(x[2]), chemin: String(x[3]), t: " " + _fp2N_(x[1] + " " + x[3]) + " " }; }); } catch (eL) {}
      var sugg = function (nom, prenom) {
        var toks = _fp2N_(nom).split(" ").filter(function (t) { return t.length >= 3; }), pre = _fp2N_(prenom).split(" ")[0] || "";
        if (!toks.length) return [];
        return libresL.map(function (f) {
          var sc = toks.filter(function (t) { return f.t.indexOf(" " + t + " ") >= 0 || f.t.indexOf(t) >= 0 && t.length >= 5; }).length;
          if (!sc) return null; if (pre && f.t.indexOf(" " + pre + " ") >= 0) sc += 2; if (/positionn?e?ment/i.test(f.nom)) sc += 1;
          return { id: f.id, nom: f.nom, cree: f.cree, chemin: f.chemin, sc: sc };
        }).filter(Boolean).sort(function (a, b) { return b.sc - a.sc; }).slice(0, 4);
      };
      var L = ws.getDataRange().getValues().slice(1).map(function (r) {
        var m = []; try { m = JSON.parse(r[FP2_C.manquants] || "[]"); } catch (e) {}
        var h = []; try { h = JSON.parse(r[FP2_C.histo] || "[]"); } catch (e) {}
        return { id: r[0], nom: r[1], prenom: r[2], typo: r[3], classe: r[4], diplome: r[5], histo: h, decision: r[7], lien: r[8], noteRefus: r[9], motifRefus: r[10],
          url: r[12], etat: r[13], manquants: m, message: r[15], suggestions: r[7] === "attente" ? sugg(r[1], r[2]) : [] };
      });
      return createJsonResponse({ success: true, liste: L });
    }
    if (op === "demarrer") { _fp2Relancer_(p.phase || "inventaire"); _fp2Log_(ss, "Commande", "Lancement (" + (p.phase || "inventaire") + ") par la direction"); logAction(p.emailAuteur, "", "Action", "Étudiants — Fiches de positionnement (refonte)", "FP2", "Lancement : " + (p.phase || "inventaire")); return createJsonResponse({ success: true }); }
    if (op === "pause" || op === "reprendre") { var j = _fp2Job_(); j.pause = op === "pause"; if (op === "reprendre") { j.actif = true; _fp2Declencheur_(true); } _fp2SetJob_(j); _fp2Log_(ss, "Commande", op === "pause" ? "Mise en pause" : "Reprise"); return createJsonResponse({ success: true }); }
    if (op === "decision") {
      var dd = ws.getDataRange().getValues();
      for (var i = 1; i < dd.length; i++) if (String(dd[i][0]) === String(p.id)) {
        var decI = String(p.decision), lien = String(p.lien || "").trim();
        if (decI === "redirection" && !/[-\w]{25,}/.test(lien)) return createJsonResponse({ success: false, message: "Lien de fiche invalide" });
        var hh = []; try { hh = JSON.parse(dd[i][FP2_C.histo] || "[]"); } catch (e) {}
        if (decI === "redirection") {
          var idL = lien.match(/[-\w]{25,}/)[0], f = DriveApp.getFileById(idL);
          if (!hh.some(function (x) { return x.id === idL; })) hh.push({ id: idL, nom: f.getName(), cree: Utilities.formatDate(f.getDateCreated(), "Europe/Paris", "yyyy-MM-dd") });
          var wsH = _fp2Histo_(ss); if (!wsH.getDataRange().getValues().some(function (r) { return String(r[0]) === idL; })) wsH.appendRow([idL, f.getName(), f.getMimeType(), Utilities.formatDate(f.getDateCreated(), "Europe/Paris", "yyyy-MM-dd"), f.getUrl(), "", "", "", "", "", ""]);
        }
        ws.getRange(i + 1, FP2_C.histo + 1, 1, 3).setValues([[JSON.stringify(hh), decI, lien]]);
        ws.getRange(i + 1, FP2_C.etat + 1).setValue(decI === "attente" ? "attente décision" : "à faire");
        _fp2Log_(ss, "Décision", dd[i][1] + " " + dd[i][2] + " → " + (decI === "redirection" ? "redirection vers une fiche" : decI === "generer" ? "nouvelle fiche" : decI));
        if (decI !== "attente") _fp2Relancer_(decI === "redirection" ? "extraction" : "generation");
        return createJsonResponse({ success: true });
      }
      return createJsonResponse({ success: false, message: "Étudiant absent du registre" });
    }
    if (op === "regenerer") {
      var dg = ws.getDataRange().getValues(), n2 = 0;
      for (var g = 1; g < dg.length; g++) if ((p.ids || []).indexOf(String(dg[g][0])) >= 0 && dg[g][FP2_C.decision] !== "attente") { ws.getRange(g + 1, FP2_C.etat + 1).setValue("à faire"); n2++; }
      _fp2Relancer_("generation");
      if (p.pilote) PropertiesService.getScriptProperties().setProperty("FP2_PILOTE", "1");   // propriété à part : jamais écrasée par un passage en cours
      return createJsonResponse({ success: true, n: n2 });
    }
    if (op === "sortis") {
      var etus = _fp2Etudiants_(ss), deja = {}; ws.getDataRange().getValues().slice(1).forEach(function (r) { deja[String(r[0])] = r[FP2_C.decision]; });
      var S = Object.keys(etus).map(function (k) { return etus[k]; }).filter(function (e) { return e.typoN === "sortie" || e.typoN === "refuse"; })
        .map(function (e) { return { id: e.id, nom: e.nom, prenom: e.prenom, classe: e.classe, typo: e.typo, sortie: e.sortie instanceof Date ? _fp2Fmt_(e.sortie) : String(e.sortie || ""), motif: e.motifSortie, decision: deja[e.id] || "" }; });
      return createJsonResponse({ success: true, sortis: S });
    }
    if (op === "refuser") {
      var note = Number(String(p.note).replace(",", "."));
      if (isNaN(note) || note >= 10 || note < 0) return createJsonResponse({ success: false, message: "La note doit être comprise entre 0 et 9,99" });
      if (!String(p.motif || "").trim()) return createJsonResponse({ success: false, message: "Motif de refus obligatoire" });
      var et = _fp2Etudiants_(ss), dr = ws.getDataRange().getValues(), idx = {}, nb = 0;
      for (var a = 1; a < dr.length; a++) idx[String(dr[a][0])] = a + 1;
      (p.ids || []).forEach(function (id) {
        var e = et[id]; if (!e) return;
        var row = [e.id, e.nom, e.prenom, e.typo, e.classe, _fp2Diplome_(_fpFormation_(e.classe)), "[]", "refus", "", note, String(p.motif).trim(), "", "", "à faire", "", "", new Date()];
        if (idx[id]) { var old = dr[idx[id] - 1]; row[FP2_C.histo] = old[FP2_C.histo] || "[]"; row[FP2_C.classeurId] = old[FP2_C.classeurId]; ws.getRange(idx[id], 1, 1, FP2_H.length).setValues([row]); }
        else ws.appendRow(row);
        nb++;
      });
      _fp2Log_(ss, "Décision", nb + " étudiant(s) sorti(s) à refuser (note " + note + ", motif : " + p.motif + ")");
      logAction(p.emailAuteur, "", "Modification", "Étudiants — Fiches de positionnement (refonte)", "Refus", nb + " fiche(s) de refus demandée(s) — " + p.motif);
      _fp2Relancer_("inventaire");   // l'inventaire rattache leurs fiches historiques avant la génération
      return createJsonResponse({ success: true, n: nb });
    }
    if (op === "bachelor") {
      if (p.criteres) {
        var tot = p.criteres.reduce(function (s, c) { return s + Number(c[2] || 0); }, 0);
        if (tot !== 100) return createJsonResponse({ success: false, message: "La somme des poids doit faire 100 % (actuellement " + tot + " %)" });
        PropertiesService.getScriptProperties().setProperty("FP2_BACH_CRITERES", JSON.stringify(p.criteres));
        _fp2Log_(ss, "Paramètres", "Critères Bachelor mis à jour");
      }
      var T = null; try { T = JSON.parse(PropertiesService.getScriptProperties().getProperty("FP2_BACH_TESTS") || "null"); } catch (e) {}
      return createJsonResponse({ success: true, criteres: _fp2BachCriteres_(), tests: T });
    }
    if (op === "lire") { var conv = _fp2Convertir_(p.id, DriveApp.getFileById(p.id).getMimeType()); var Dl = _fp2Lire_(conv); if (conv !== p.id) DriveApp.getFileById(conv).setTrashed(true); return createJsonResponse({ success: true, d: Dl }); }
    if (op === "reextraire") {
      var wh = _fp2Histo_(ss); if (wh.getLastRow() > 1) wh.getRange(2, 9, wh.getLastRow() - 1, 3).clearContent();
      var jx = _fp2Job_(); jx.phase = "extraction"; jx.phaseDebut = Date.now(); jx.actif = true; jx.pause = false; _fp2SetJob_(jx);
      _fp2Log_(ss, "Commande", "Relecture de toutes les fiches historiques"); return createJsonResponse({ success: true });
    }
    if (op === "formule") { var shF = SpreadsheetApp.openById(p.id).getSheetByName(p.onglet) || SpreadsheetApp.openById(p.id).getSheets()[0]; return createJsonResponse({ success: true, f: shF.getRange(p.a1).getFormula(), v: shF.getRange(p.a1).getDisplayValue(), modele: SpreadsheetApp.openById(_fpModele_()).getSheets()[0].getRange(p.a1).getFormula() }); }
    if (op === "kick") { fp2Tick(280000); return createJsonResponse({ success: true }); }
    if (op === "declencheurs") return createJsonResponse({ success: true, liste: ScriptApp.getProjectTriggers().map(function (t) { return { f: t.getHandlerFunction(), type: String(t.getEventType()), src: String(t.getTriggerSource()) }; }) });
    if (op === "tests") { var R = _fp2TestsBachelor_(); _fp2Log_(ss, "Bachelor", "Tests d'admission générés"); logAction(p.emailAuteur, "", "Création", "Étudiants — Fiches de positionnement (refonte)", "Bachelor", "Tests d'admission (sujet + corrigé)"); return createJsonResponse({ success: true, tests: R }); }
    return createJsonResponse({ success: false, message: "Opération inconnue" });
  } catch (e) { return createJsonResponse({ success: false, message: String(e) }); }
}
