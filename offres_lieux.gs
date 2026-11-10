// =============================================================================
//  OFFRES — LIEUX HARMONISÉS + COHÉRENCE DES PARTAGES (demande direction 2026-11-10)
//
//  Lieux (OFFRES col. F) : les aspirateurs et la saisie produisent des formats hétérogènes
//  (« Marseille, France », « Marseille 8e - 13 », « G&#xE9;menos - 13 », « 13008, Marseille, France »…).
//  Comme pour les entreprises (adresses.gs), chaque lieu est rapproché de la Base Adresse Nationale
//  (communes) et réécrit au format unique « Commune (département) » — ex. « Aix-en-Provence (13) »,
//  « Marseille 8e (13) ». Colonnes ajoutées (repérées par leur en-tête) : code postal, latitude,
//  longitude, statut, lieu d'origine (texte brut conservé). Statuts : « BAN », « Département »,
//  « Non précisé » (« Alternance », « CDI »… ne sont pas des lieux), « À vérifier ».
//  Passe complète au déploiement, puis toutes les nuits (nouvelles offres) via auditAdressesTick.
//
//  Partages : une offre partagée doit être « Vérifiée » (col. J) et « Partagée » (col. N).
//  Réparation unique des offres déjà partagées (sauf celles marquées « Obsolète »).
// =============================================================================
var OF_LIEU_HEADERS = { cp: "Lieu — code postal", lat: "Lieu — latitude", lon: "Lieu — longitude", statut: "Lieu statut", origine: "Lieu d'origine" };
var OF_DEPTS = { "bouches du rhone": "13", "var": "83", "vaucluse": "84", "alpes de haute provence": "04", "hautes alpes": "05", "alpes maritimes": "06", "gard": "30", "herault": "34", "provence alpes cote d azur": "", "ile de france": "", "paris": "75" };
// Communes à plusieurs codes postaux : la BAN renvoie un code arbitraire (13001, 13080…) → code principal ou vide
var OF_CP_MULTI = { "marseille": "", "aix en provence": "13100", "toulon": "", "martigues": "13500", "la seyne sur mer": "83500" };
var OF_PAS_LIEU = /^(alternance|cdi|cdd|stage|interim|teletravail|tele travail|remote|france|a distance|hybride|non precise|n a|na|nan|undefined|null|-)$/;

function _ofCols_(ws) {
  var lastCol = Math.max(ws.getLastColumn(), 1);
  var header = ws.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h || "").trim().toLowerCase(); });
  var res = {}, toAdd = [];
  Object.keys(OF_LIEU_HEADERS).forEach(function (k) { var i = header.indexOf(OF_LIEU_HEADERS[k].toLowerCase()); if (i === -1) toAdd.push(k); else res[k] = i; });
  if (toAdd.length) {
    if (ws.getMaxColumns() < lastCol + toAdd.length) ws.insertColumnsAfter(ws.getMaxColumns(), lastCol + toAdd.length - ws.getMaxColumns());
    ws.getRange(1, lastCol + 1, 1, toAdd.length).setValues([toAdd.map(function (k) { return OF_LIEU_HEADERS[k]; })]).setFontWeight("bold");
    toAdd.forEach(function (k, j) { res[k] = lastCol + j; });
  }
  return res;
}

function _ofDecode_(s) {
  return String(s || "").replace(/&#x([0-9a-f]+);/gi, function (m, h) { return String.fromCharCode(parseInt(h, 16)); })
    .replace(/&#(\d+);/g, function (m, d) { return String.fromCharCode(+d); })
    .replace(/&amp;/g, "&").replace(/&apos;|&#39;/g, "'").replace(/&quot;/g, '"').replace(/\s+/g, " ").trim();
}

// Analyse locale du texte brut → { ville, cp, dep, arr, pasLieu, departement }
function _ofAnalyser_(brut) {
  var s = _ofDecode_(brut);
  var n = _adrNorm_(s);
  if (!n || OF_PAS_LIEU.test(n)) return { pasLieu: true };
  var parts = s.split(",").map(function (x) { return x.trim(); }).filter(function (x) { return x && !/^france$/i.test(x); });
  if (!parts.length) return { pasLieu: true };
  var cp = "", dep = "", arr = 0;
  var mcp = s.match(/\b(\d{5})\b/); if (mcp) cp = mcp[1];
  // « Ville - 13 » / « Ville (13) »
  var cand = parts.length > 1 ? parts.filter(function (x) { return !/^\d{5}$/.test(x) && !/\d+[ ,]/.test(x.slice(0, 4)); }).pop() || parts[parts.length - 1] : parts[0];
  var md = cand.match(/^(.*?)\s*(?:-|\()\s*(\d{2}|2A|2B)\s*\)?$/i); if (md) { cand = md[1]; dep = md[2]; }
  cand = cand.replace(/\b\d{5}\b/, "").replace(/\s+France$/i, "").split(/\s+-\s+/)[0].replace(/^arrondissement d(?:e |')/i, "").trim();   // « A - B » : premier lieu
  var nc = _adrNorm_(cand);
  if (OF_DEPTS.hasOwnProperty(nc)) return { departement: true, ville: _ofDecode_(cand), dep: OF_DEPTS[nc] };
  var mr = nc.match(/^(?:arrondissement de )?marseille(?: (\d{1,2})(?: ?(?:e|er|eme))?(?: arrondissement)?)?$/);
  if (mr) { arr = mr[1] ? +mr[1] : 0; if (!arr && cp && /^130(0[1-9]|1[0-6])$/.test(cp)) arr = +cp.slice(3); return { ville: "Marseille", cp: arr ? String(13000 + arr) : (cp || ""), dep: "13", arr: arr }; }
  var ma = nc.match(/^aix en provence/); if (ma) cand = "Aix-en-Provence";
  if (!dep && cp) dep = cp.slice(0, 2);
  return { ville: cand, cp: cp, dep: dep };
}

function _ofAffichage_(a) {
  if (a.arr) return "Marseille " + a.arr + (a.arr === 1 ? "er" : "e") + " (13)";
  return a.ville + (a.dep ? " (" + a.dep + ")" : "");
}

// Harmonise les lieux ; forcer = repasse aussi les lignes déjà traitées
function _ofHarmoniserLieux_(ss, forcer, limiteMs) {
  var t0 = Date.now(), ws = getSheetSafe(ss, "OFFRES"), C = _ofCols_(ws);
  var d = ws.getDataRange().getValues(), aFaire = [];
  for (var i = 1; i < d.length; i++) {
    if (!d[i][0]) continue;
    var f = String(d[i][5] || "").trim(), st = String(d[i][C.statut] || "").trim();
    if (!forcer && st && /^.+ \((\d{2}|2A|2B)\)$/.test(f)) continue;     // déjà au format, non modifié depuis
    if (!forcer && st === "Non précisé" && !f) continue;
    aFaire.push(i);
  }
  if (!aFaire.length) return { traitees: 0, reste: 0 };
  var analyses = {}, communes = {};
  aFaire.forEach(function (i) {
    var brut = forcer && String(d[i][C.origine] || "").trim() ? String(d[i][C.origine]) : String(d[i][5] || "");
    var a = _ofAnalyser_(brut); a.brut = brut; analyses[i] = a;
    if (!a.pasLieu && !a.departement && !a.arr && a.ville) communes[_adrNorm_(a.ville) + "|" + (a.dep || "")] = { ville: a.ville, cp: a.cp, dep: a.dep };
  });
  var cles = Object.keys(communes), items = cles.map(function (k) { return communes[k]; });
  var cacheC = CacheService.getScriptCache(), res = {}, manquants = [];
  cles.forEach(function (k, j) { var c = cacheC.get("ofc_" + Utilities.base64EncodeWebSafe(k).slice(0, 200)); if (c) res[k] = JSON.parse(c); else manquants.push(j); });
  if (manquants.length) {
    var trouves = _adrCommuneLot_(manquants.map(function (j) { return items[j]; }));
    manquants.forEach(function (j, q) {
      var k = cles[j], f = trouves[q];
      res[k] = f ? { ville: f.city, cp: f.postcode, lat: f.lat, lon: f.lon } : null;
      try { cacheC.put("ofc_" + Utilities.base64EncodeWebSafe(k).slice(0, 200), JSON.stringify(res[k]), 21600); } catch (e) {}
    });
  }
  var nb = { ban: 0, dep: 0, vide: 0, verif: 0 }, exemples = [];
  var c0 = Math.min(C.cp, C.lat, C.lon, C.statut, C.origine), c1 = Math.max(C.cp, C.lat, C.lon, C.statut, C.origine);
  var largeur = Math.max(d[0].length, c1 + 1);
  d.forEach(function (row) { while (row.length < largeur) row.push(""); });
  aFaire.forEach(function (i) {
    if (Date.now() - t0 > limiteMs) return;
    var a = analyses[i], out = { f: "", cp: "", lat: "", lon: "", statut: "" };
    if (a.pasLieu) { out.statut = "Non précisé"; nb.vide++; }
    else if (a.departement) { out.f = a.ville + (a.dep ? " (" + a.dep + ")" : ""); out.statut = "Département"; nb.dep++; }
    else if (a.arr) { out.f = _ofAffichage_(a); out.cp = a.cp; out.statut = "BAN"; nb.ban++;
      var m = res["marseille|13"]; if (m) { out.lat = m.lat; out.lon = m.lon; } }
    else {
      var r = res[_adrNorm_(a.ville) + "|" + (a.dep || "")];
      if (r) { out.f = r.ville + " (" + String(r.cp).slice(0, 2) + ")"; out.cp = a.cp || (OF_CP_MULTI[_adrNorm_(r.ville)] !== undefined ? OF_CP_MULTI[_adrNorm_(r.ville)] : r.cp); out.lat = r.lat; out.lon = r.lon; out.statut = "BAN"; nb.ban++; }
      else { out.f = _ofAffichage_(a); out.cp = a.cp || ""; out.statut = "À vérifier"; nb.verif++; }
    }
    if (exemples.length < 15 && out.f !== a.brut) exemples.push("« " + a.brut + " » → « " + (out.f || "(vide)") + " »");
    var ligne = d[i];
    ligne[5] = out.f; ligne[C.cp] = out.cp; ligne[C.lat] = out.lat; ligne[C.lon] = out.lon; ligne[C.statut] = out.statut;
    if (!String(ligne[C.origine] || "").trim()) ligne[C.origine] = a.brut;
  });
  // Écriture groupée : colonne F + colonnes de lieu (au lieu d'une écriture par ligne)
  var n = d.length - 1;
  if (n > 0) {
    ws.getRange(2, 6, n, 1).setNumberFormat("@").setValues(d.slice(1).map(function (r) { return [r[5]]; }));
    ws.getRange(2, c0 + 1, n, c1 - c0 + 1).setValues(d.slice(1).map(function (r) { return r.slice(c0, c1 + 1); }));
  }
  var traitees = nb.ban + nb.dep + nb.vide + nb.verif;
  _invalidateOffresCache_();
  return { traitees: traitees, reste: aFaire.length - traitees, nb: nb, exemples: exemples };
}

// Offres présentes dans PARTAGES : qualité « Vérifiée » (sauf Obsolète) et état « Partagée »
function _ofReparerPartages_(ss) {
  var ws = getSheetSafe(ss, "OFFRES"), d = ws.getDataRange().getValues(), part = {}, nbQ = 0, nbE = 0;
  getSheetSafe(ss, "PARTAGES").getDataRange().getValues().forEach(function (r, i) { if (i) { var k = String(r[1] || "").trim(); if (k) part[k] = 1; } });
  for (var i = 1; i < d.length; i++) {
    if (!part[String(d[i][0] || "").trim()]) continue;
    var q = String(d[i][9] || "").trim(), e = String(d[i][13] || "").trim();
    if (q !== "Vérifiée" && q !== "Obsolète") { ws.getRange(i + 1, 10).setValue("Vérifiée"); nbQ++; }
    if (e.indexOf("Partagée") === -1) { ws.getRange(i + 1, 14).setValue("Partagée"); nbE++; }
  }
  _invalidateOffresCache_();
  logAction("Réparation automatique", "", "Modification", "Offres — partages", "OFFRES", nbQ + " offre(s) partagée(s) passée(s) en « Vérifiée », " + nbE + " remise(s) en « Partagée »");
  return { qualite: nbQ, etat: nbE };
}

// Appelé par auditAdressesTick : réparation unique des partages, passe complète des lieux, puis 1×/nuit
function _ofAuto_(t0) {
  var props = PropertiesService.getScriptProperties();
  var ss = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");
  if (props.getProperty("OFFRES_PARTAGES_REPARE_V1") !== "1") {
    var rp = _ofReparerPartages_(ss);
    props.setProperty("OFFRES_PARTAGES_REPARE_V1", "1");
    props.setProperty("OFFRES_PARTAGES_REPARE_RESUME", JSON.stringify(rp));
  }
  var h = Number(Utilities.formatDate(new Date(), "Europe/Paris", "H")), auj = Utilities.formatDate(new Date(), "Europe/Paris", "yyyy-MM-dd");
  var enCours = props.getProperty("OFFRES_LIEUX_SUITE") === "1" || props.getProperty("OFFRES_LIEUX_V1") !== "1" || props.getProperty("OFFRES_LIEUX_V2") !== "1";
  if (!enCours && !(h === 2 && props.getProperty("OFFRES_LIEUX_JOUR") !== auj)) return;
  var budget = Math.min(120000, 200000 - (Date.now() - (t0 || Date.now())));
  if (budget < 30000) return;
  props.setProperty("OFFRES_LIEUX_JOUR", auj);
  var forcer = props.getProperty("OFFRES_LIEUX_V2") !== "1";   // V2 : codes postaux des communes multi-CP corrigés
  var R = _ofHarmoniserLieux_(ss, forcer, budget);
  props.setProperty("OFFRES_LIEUX_V1", "1");
  if (forcer && R.reste === 0) props.setProperty("OFFRES_LIEUX_V2", "1");
  if (R.reste > 0) props.setProperty("OFFRES_LIEUX_SUITE", "1"); else props.deleteProperty("OFFRES_LIEUX_SUITE");
  if (R.traitees) {
    logAction("Harmonisation automatique", "", "Modification", "Offres — lieux", "OFFRES",
      R.traitees + " lieu(x) d'offre harmonisé(s) : " + R.nb.ban + " commune(s) BAN, " + R.nb.dep + " département(s), " + R.nb.vide + " non précisé(s), " + R.nb.verif + " à vérifier" + (R.reste ? " — reste " + R.reste : ""));
    var prev = props.getProperty("OFFRES_LIEUX_RESUME"), acc = prev ? JSON.parse(prev) : { ban: 0, dep: 0, vide: 0, verif: 0, ex: [] };
    ["ban", "dep", "vide", "verif"].forEach(function (k) { acc[k] += R.nb[k]; });
    acc.ex = (acc.ex || []).concat(R.exemples).slice(0, 20);
    props.setProperty("OFFRES_LIEUX_RESUME", JSON.stringify(acc));
  }
}
