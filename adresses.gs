/* =============================================================================
   ADRESSES ENTREPRISES — adressage normalisé (novembre 2026)
   Référentiels officiels, gratuits et sans clé :
   • Base Adresse Nationale (BAN, IGN / data.gouv) via la Géoplateforme :
       https://data.geopf.fr/geocodage/search   (+ /search/csv pour le traitement en lot)
   • Répertoire SIRENE (INSEE) via https://recherche-entreprises.api.gouv.fr (côté page).
   Règle : UNE entreprise = UNE adresse. La feuille PARTENARIAT a une ligne par contact :
   l'adresse (K numéro + voie, L CP, M ville) et les colonnes ci-dessous sont recopiées
   à l'identique sur toutes les lignes de l'entreprise.
   Colonnes ajoutées au besoin (repérées par leur en-tête) : voir ADR_HEADERS.
   Statuts : « BAN » (adresse au numéro), « BAN voie » (rue sans numéro), « SIRENE »,
   « Commune » (position approximative au centre de la commune), « À vérifier »,
   « Saisie libre » (adresse forcée par l'utilisateur, non trouvée dans la BAN).
   Front : adresse-fr.js (fenêtre d'adresse avec autocomplétion BAN + recherche SIRENE).
   ========================================================================== */

var ADR_HEADERS = {
  complement: "Complément adresse", lat: "Latitude", lon: "Longitude",
  statut: "Adresse statut", siret: "SIRET", origine: "Adresse d'origine"
};
var ADR_BAN_URL = "https://data.geopf.fr/geocodage/search";
var ADR_STATUTS_OK = { "BAN": 1, "BAN voie": 1, "SIRENE": 1 };
var _ADR_COLS_ = null;

// Index (0-based) des colonnes d'adresse étendues ; crée les en-têtes manquants.
function _adrCols_(ws) {
  if (_ADR_COLS_) return _ADR_COLS_;
  var lastCol = Math.max(ws.getLastColumn(), 1);
  var header = ws.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h || "").trim().toLowerCase(); });
  var res = {}, toAdd = [];
  Object.keys(ADR_HEADERS).forEach(function (k) {
    var i = header.indexOf(ADR_HEADERS[k].toLowerCase());
    if (i === -1) toAdd.push(k); else res[k] = i;
  });
  if (toAdd.length) {
    if (ws.getMaxColumns() < lastCol + toAdd.length) ws.insertColumnsAfter(ws.getMaxColumns(), lastCol + toAdd.length - ws.getMaxColumns());
    ws.getRange(1, lastCol + 1, 1, toAdd.length).setValues([toAdd.map(function (k) { return ADR_HEADERS[k]; })]).setFontWeight("bold");
    toAdd.forEach(function (k, j) { res[k] = lastCol + j; });
  }
  _ADR_COLS_ = res;
  return res;
}

// Lecture (sans création de colonnes) pour les fonctions de lecture
function _adrColsLecture_(header) {
  var h = (header || []).map(function (x) { return String(x || "").trim().toLowerCase(); });
  var res = {};
  Object.keys(ADR_HEADERS).forEach(function (k) { res[k] = h.indexOf(ADR_HEADERS[k].toLowerCase()); });
  return res;
}

// Adresse complète d'une ligne PARTENARIAT
function _adrDeLigne_(row, C) {
  function g(k) { return C[k] >= 0 ? row[C[k]] : ""; }
  var lat = Number(g("lat")), lon = Number(g("lon"));
  return {
    adresse: _adrVide_(row[10]), cp: _adrVide_(row[11]), ville: _adrVide_(row[12]),
    complement: _adrVide_(g("complement")),
    lat: isFinite(lat) && lat ? lat : "", lon: isFinite(lon) && lon ? lon : "",
    statut: String(g("statut") || "").trim(), siret: String(g("siret") || "").replace(/\s/g, "")
  };
}

// « x », « . », « - » ou espaces = vide
function _adrVide_(s) {
  s = String(s == null ? "" : s).replace(/\s+/g, " ").trim();
  return !s || /^[x.\-\s]+$/i.test(s) ? "" : s;
}
function _adrNorm_(s) {
  return String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase()
    .replace(/[’']/g, " ").replace(/[^a-z0-9]+/g, " ").trim();
}
var ADR_PETITS_MOTS = { de: 1, du: 1, des: 1, la: 1, le: 1, les: 1, et: 1, "à": 1, au: 1, aux: 1, sur: 1, en: 1, sous: 1, "lès": 1 };
// « Avenue du prado » → « Avenue du Prado » (la BAN renvoie parfois les voies en minuscules)
function _adrCasse_(s) {
  var mots = String(s || "").toLowerCase().split(/(\s+|-)/);
  return mots.map(function (w, i) {
    if (!w.trim() || w === "-") return w;
    if (i > 0 && ADR_PETITS_MOTS[w]) return w;
    var m = w.match(/^([dl])(['’])(.+)$/);
    if (m) return (i > 0 ? m[1] : m[1].toUpperCase()) + "'" + m[3].charAt(0).toUpperCase() + m[3].slice(1);
    return w.charAt(0).toUpperCase() + w.slice(1);
  }).join("");
}

var ADR_MOTS_VOIE = {
  rue: 1, r: 1, avenue: 1, av: 1, ave: 1, boulevard: 1, bd: 1, bld: 1, blvd: 1, chemin: 1, chem: 1, ch: 1, route: 1, rte: 1,
  place: 1, pl: 1, allee: 1, all: 1, impasse: 1, imp: 1, traverse: 1, trav: 1, quai: 1, cours: 1, crs: 1, montee: 1,
  square: 1, passage: 1, voie: 1, lotissement: 1, residence: 1, zac: 1, za: 1, zi: 1, parc: 1, rond: 1, point: 1,
  de: 1, du: 1, des: 1, la: 1, le: 1, les: 1, d: 1, l: 1, et: 1, saint: 1, st: 1, sainte: 1, ste: 1, bis: 1, ter: 1,
  docteur: 1, dr: 1, general: 1, gal: 1, marechal: 1, mal: 1, president: 1, pdt: 1, commandant: 1, cdt: 1,
  lieutenant: 1, lt: 1, capitaine: 1, cap: 1, professeur: 1, pr: 1, colonel: 1, col: 1, abbe: 1
};
// Types de voie (saisie abrégée → forme BAN) : la BAN doit proposer un type présent dans la saisie
var ADR_TYPES_VOIE = {
  rue: "rue", r: "rue", avenue: "avenue", av: "avenue", ave: "avenue", boulevard: "boulevard", bd: "boulevard", bld: "boulevard",
  blvd: "boulevard", bvd: "boulevard", chemin: "chemin", chem: "chemin", ch: "chemin", route: "route", rte: "route", place: "place",
  pl: "place", allee: "allee", allees: "allee", impasse: "impasse", imp: "impasse", traverse: "traverse", trav: "traverse",
  quai: "quai", cours: "cours", crs: "cours", montee: "montee", square: "square", sq: "square", passage: "passage", rond: "rond"
};
function _adrTypesVoie_(txt) {
  var t = {};
  _adrNorm_(txt).split(" ").forEach(function (w) { if (ADR_TYPES_VOIE[w]) t[ADR_TYPES_VOIE[w]] = 1; });
  return t;
}
function _adrMotsRue_(street) {
  return _adrNorm_(street).split(" ").filter(function (w) { return w.length >= 3 && !ADR_MOTS_VOIE[w] && !/^\d+$/.test(w); });
}
// La voie proposée par la BAN doit se retrouver dans le texte saisi (tolère une petite faute)
function _adrRueCompatible_(street, original) {
  var o = " " + _adrNorm_(original) + " ";
  var mots = _adrMotsRue_(street);
  if (!mots.length) return false;
  var tSaisie = _adrTypesVoie_(original), kB = Object.keys(_adrTypesVoie_(String(street || "").split(/\s+/)[0]))[0];
  if (Object.keys(tSaisie).length && kB && !tSaisie[kB]) return false;   // « Avenue … » ≠ « Chemin … »
  var ok = mots.filter(function (w) {
    return o.indexOf(" " + w + " ") !== -1 || (w.length >= 5 && o.indexOf(" " + w.slice(0, 5)) !== -1);
  }).length;
  return ok / mots.length >= 0.6;
}
function _adrVilleCompatible_(villeSaisie, villeBan) {
  var a = _adrNorm_(villeSaisie), b = _adrNorm_(villeBan);
  if (!a || !b) return false;
  var a1 = a.split(" ")[0], b1 = b.split(" ").filter(function (w) { return w.length > 2; })[0] || b.split(" ")[0];
  return a.indexOf(b1) !== -1 || b.indexOf(a1) !== -1;
}

// Retire « 13006 Marseille » recopié dans la ligne d'adresse, « Adresse : », espaces doubles
function _adrNettoyer_(adresse, cp) {
  var s = String(adresse || "").replace(/^\s*adresse\s*:\s*/i, "").replace(/\s+/g, " ").trim();
  // CP (+ ville / CEDEX) en fin de ligne — mais pas « CS 70167 » ni « BP 90020 »
  s = s.replace(/[,\s]*(?<!\b(?:cs|bp|box|tsa)\s?)\b\d{5}\b(?:\s+[A-Za-zÀ-ÿ][\wÀ-ÿ'’ .-]*)?\s*$/i, "").replace(/[,;\s]+$/, "").trim();
  return s;
}

// Complément (bâtiment, résidence…) = ce qui, dans le texte saisi, n'est ni le numéro ni la voie
function _adrComplement_(original, housenumber, street, ville) {
  var s = _adrNettoyer_(original, "");
  var motsRue = _adrMotsRue_(street);
  var villeN = _adrNorm_(ville);
  var segs = s.split(/\s*(?:,|\s[-–]\s|\.\s)\s*/).map(function (x) { return x.trim(); }).filter(Boolean);
  var out = [];
  segs.forEach(function (seg) {
    var toks = seg.split(/\s+/);
    var tn = toks.map(function (t) { var n = _adrNorm_(t).split(" "); return n[n.length - 1]; });
    var idxRue = -1, idxNum = -1;
    tn.forEach(function (t, i) {
      if (motsRue.some(function (w) { return t === w || (w.length >= 5 && t.indexOf(w.slice(0, 5)) === 0); })) idxRue = i;
      if (idxNum === -1 && housenumber && (t === _adrNorm_(housenumber) || t.indexOf(_adrNorm_(housenumber)) === 0 && /^\d/.test(t))) idxNum = i;
    });
    if (idxRue === -1) { // segment sans la voie → complément (sauf s'il ne contient que la ville)
      if (_adrNorm_(seg) && _adrNorm_(seg) !== villeN) out.push(seg);
      return;
    }
    if (idxNum > 0) out.push(toks.slice(0, idxNum).join(" "));
    var rueToks = " " + _adrNorm_(street) + " ", deb = idxRue + 1;
    while (deb < toks.length && rueToks.indexOf(" " + tn[deb] + " ") !== -1) deb++;   // « 8 Mai 1945 » : 1945 fait partie de la voie
    var tail = toks.slice(deb).join(" ").trim();
    if (tail && _adrNorm_(tail) !== villeN) out.push(tail);
  });
  return out.join(", ").replace(/^[,\s]+|[,\s]+$/g, "").slice(0, 120);
}

// ---------------------------------------------------------------------------
// Évaluation d'un résultat BAN par rapport à la saisie → adresse retenue ou null
// f = { type, score, housenumber, street, name, postcode, city, lat, lon }
function _adrEvaluer_(f, saisie) {
  if (!f || !f.type || !isFinite(f.lat)) return null;
  var cpOk = saisie.cp && f.postcode === saisie.cp;
  var villeOk = saisie.ville && _adrVilleCompatible_(saisie.ville, f.city);
  var depOk = !saisie.cp || String(f.postcode).slice(0, 2) === String(saisie.cp).slice(0, 2);
  // Décalage « léger » seulement : même ville (le CP peut être corrigé dans la même ville / le
  // même département). Une ville différente (ex. « Marseille » saisi, adresse trouvée à Aix) = trop
  // complexe → pas de correction automatique.
  if (saisie.ville ? !(villeOk && depOk) : !cpOk) return null;
  var orig = saisie.adresse || "";
  if ((f.type === "housenumber" || f.type === "street" || f.type === "locality") && orig) {
    if (f.score < 0.4 || !_adrRueCompatible_(f.street || f.name, orig)) return null;
    var origN = " " + _adrNorm_(orig) + " ";
    var hn = f.type === "housenumber" ? String(f.housenumber || "") : "";
    if (hn && origN.indexOf(" " + _adrNorm_(hn)) === -1) return null;  // numéro différent
    var voie = _adrCasse_(f.street || f.name).replace(/\s+/g, " ").trim();
    var numSaisi = "";
    if (!hn) {   // numéro saisi juste avant le type de voie (« Le Noilly Prat 167 rue Paradis »)
      var m = orig.match(/(?:^|[\s,])(\d{1,4}\s?(?:bis|ter|[a-d])?)\s*,?\s+(?:rue|r|av|ave|avenue|bd|bld|blvd|boulevard|chemin|chem|ch|route|rte|place|pl|all[ée]e?s?|impasse|imp|traverse|trav|quai|cours|crs|mont[ée]e|square|passage)\b/i);
      if (!m) m = orig.match(/^\s*(\d{1,4}\s?(?:bis|ter)?)(?=[\s,a-zà]|$)/i);   // « 68 Aevnue… », « 2à Avenue… »
      if (m) numSaisi = m[1].replace(/\s+/g, " ").trim();
    }
    return {
      adresse: (hn ? hn + " " : numSaisi ? numSaisi + " " : "") + voie,
      cp: String(f.postcode), ville: f.city,
      complement: _adrComplement_(orig, hn || numSaisi, f.street || f.name, saisie.ville || f.city),
      lat: Number(f.lat), lon: Number(f.lon), statut: hn ? "BAN" : "BAN voie"
    };
  }
  return null;
}

// Commune seule (position approximative) : f = résultat BAN de type municipality
function _adrEvaluerCommune_(f, saisie) {
  if (!f || f.type !== "municipality" || !isFinite(f.lat) || f.score < 0.45) return null;
  if (saisie.cp && String(f.postcode).slice(0, 2) !== String(saisie.cp).slice(0, 2)) return null;
  if (saisie.ville && !_adrVilleCompatible_(saisie.ville, f.city)) return null;
  return { cp: saisie.cp || String(f.postcode), ville: f.city, lat: Number(f.lat), lon: Number(f.lon), statut: "Commune" };
}

// Géocodage unitaire (création / modification sans choix dans les propositions)
function _adrGeocoderUne_(saisie) {
  try {
    var q = saisie.adresse ? [saisie.adresse, saisie.cp, saisie.ville].filter(Boolean).join(" ")
      : [saisie.ville, saisie.cp ? saisie.cp.slice(0, 2) : ""].filter(Boolean).join(" ");
    if (!q || q.length < 3) return null;
    var url = ADR_BAN_URL + "?limit=1&q=" + encodeURIComponent(q.slice(0, 200)) + (saisie.adresse ? "" : "&type=municipality");
    var r = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) return null;
    var ft = (JSON.parse(r.getContentText()).features || [])[0];
    if (!ft) return null;
    var pr = ft.properties || {};
    var f = { type: pr.type, score: pr.score, housenumber: pr.housenumber, street: pr.street, name: pr.name,
      postcode: pr.postcode, city: pr.city, lon: ft.geometry.coordinates[0], lat: ft.geometry.coordinates[1] };
    return saisie.adresse ? _adrEvaluer_(f, saisie) : _adrEvaluerCommune_(f, saisie);
  } catch (e) { return null; }
}

// Géocodage en lot (BAN /search/csv) : items = [{ q, t? }] → tableau de résultats f (ou null)
function _adrGeocoderLot_(items, typeCommune) {
  var out = [];
  for (var start = 0; start < items.length; start += 400) {
    if (!_adrTempsOk_()) { _adrDiag_("temps limite atteint : " + (items.length - start) + " requête(s) non faite(s)"); while (out.length < items.length) out.push(null); break; }
    var part = items.slice(start, start + 400);
    var csv = "id,q,t\n" + part.map(function (it, i) {
      return i + ',"' + String(it.q || "").replace(/"/g, '""').replace(/[\r\n]+/g, " ") + '",' + (typeCommune ? "municipality" : "");
    }).join("\n");
    var payload = { data: Utilities.newBlob(csv, "text/csv", "adresses.csv"), columns: "q" };
    if (typeCommune) payload.type = "t";
    var res = null;
    for (var essai = 0; essai < 3 && !res && !_ADR_CSV_KO_; essai++) {
      try {
        var r = UrlFetchApp.fetch(ADR_BAN_URL + "/csv", { method: "post", payload: payload, muteHttpExceptions: true });
        if (r.getResponseCode() === 200) res = r.getContentText("UTF-8");
        else { _adrDiag_("BAN csv HTTP " + r.getResponseCode() + " : " + r.getContentText().slice(0, 150)); if (r.getResponseCode() >= 500) _ADR_CSV_KO_ = true; else Utilities.sleep(1500); }
      } catch (e) { _adrDiag_("BAN csv : " + e); Utilities.sleep(1500); }
    }
    var parsed = {}, nb = 0;
    try {
      var rows = res ? Utilities.parseCsv(res.replace(/^﻿/, "")) : [];
      var h = rows.length ? rows[0].map(function (x) { return String(x).trim(); }) : [];
      var ix = function (n) { return h.indexOf(n); };
      for (var k = 1; k < rows.length; k++) {
        var row = rows[k];
        if (!row[ix("result_type")]) continue;
        nb++;
        parsed[String(row[0]).trim()] = {
          type: row[ix("result_type")], score: Number(row[ix("result_score")]) || 0,
          housenumber: row[ix("result_housenumber")], street: row[ix("result_street")], name: row[ix("result_name")],
          postcode: row[ix("result_postcode")], city: row[ix("result_city")],
          lat: Number(row[ix("latitude")]), lon: Number(row[ix("longitude")])
        };
      }
      if (res && !nb) _adrDiag_("BAN csv sans résultat : " + res.slice(0, 150));
    } catch (e3) { _adrDiag_("BAN csv illisible : " + e3); nb = 0; }
    if (!nb && part.length) {   // repli : une requête par adresse (en parallèle)
      var unit = _adrGeocoderUnitaires_(part, typeCommune);
      part.forEach(function (it, i) { out.push(unit[i]); });
    } else part.forEach(function (it, i) { out.push(parsed[String(i)] || null); });
  }
  return out;
}

var _ADR_DIAG_ = [], _ADR_T0_ = 0;
var _ADR_LIMITE_ = 0, _ADR_CSV_KO_ = false;   // _ADR_CSV_KO_ : géocodage en lot refusé → requêtes unitaires directement
// fetchAll tolérant : une seule URL en échec (« Adresse non disponible ») ne fait plus échouer tout le paquet
function _adrFetchAll_(reqs) {
  try { return UrlFetchApp.fetchAll(reqs); } catch (e) {
    _adrDiag_("paquet en échec, reprise une par une : " + String(e).slice(0, 120));
    return reqs.map(function (q) { try { return UrlFetchApp.fetch(q.url, { muteHttpExceptions: true }); } catch (e2) { return null; } });
  }
}   // échéance absolue de l'exécution en cours (posée par auditAdressesTick)
function _adrTempsOk_() { return (!_ADR_T0_ || Date.now() - _ADR_T0_ < 250000) && (!_ADR_LIMITE_ || Date.now() < _ADR_LIMITE_); }   // marge sous la limite de 6 min
// Sonde au début de l'audit : ce que répondent réellement les services depuis les serveurs Google
function _adrSonde_() {
  try {
    var t = Date.now();
    var r1 = UrlFetchApp.fetch(ADR_BAN_URL + "/csv", { method: "post", muteHttpExceptions: true,
      payload: { data: Utilities.newBlob('id,q\n0,"146 rue Paradis 13006 Marseille"', "text/csv", "a.csv"), columns: "q" } });
    if (r1.getResponseCode() >= 500) _ADR_CSV_KO_ = true;
    _adrDiag_("sonde BAN lot HTTP " + r1.getResponseCode() + " (" + (Date.now() - t) + " ms) : " + r1.getContentText().replace(/\s+/g, " ").slice(0, 160));
    t = Date.now();
    var r2 = UrlFetchApp.fetch(ADR_BAN_URL + "?limit=1&q=" + encodeURIComponent("146 rue Paradis 13006 Marseille"), { muteHttpExceptions: true });
    _adrDiag_("sonde BAN unitaire HTTP " + r2.getResponseCode() + " (" + (Date.now() - t) + " ms)");
    t = Date.now();
    var r3 = UrlFetchApp.fetch(ADR_SIRENE_URL + "?per_page=1&q=secafi", { muteHttpExceptions: true });
    _adrDiag_("sonde SIRENE HTTP " + r3.getResponseCode() + " (" + (Date.now() - t) + " ms)");
  } catch (e) { _adrDiag_("sonde : " + e); }
}
function _adrDiag_(m) { if (_ADR_DIAG_.length < 12) _ADR_DIAG_.push(m); console.warn("adresses : " + m); }

function _adrGeocoderUnitaires_(items, typeCommune) {
  var out = [];
  for (var s = 0; s < items.length; s += 25) {
    if (!_adrTempsOk_()) { _adrDiag_("temps limite atteint : " + (items.length - s) + " requête(s) non faite(s)"); while (out.length < items.length) out.push(null); break; }
    var part = items.slice(s, s + 25);
    var reqs = part.map(function (it) {
      return { url: ADR_BAN_URL + "?limit=1&q=" + encodeURIComponent(String(it.q || "x").slice(0, 200)) + (typeCommune ? "&type=municipality" : ""), muteHttpExceptions: true };
    });
    var rep = [];
    rep = _adrFetchAll_(reqs);
    part.forEach(function (it, k) {
      var f = null;
      try {
        var r = rep[k];
        if (r && r.getResponseCode() === 200) {
          var ft = (JSON.parse(r.getContentText()).features || [])[0];
          if (ft) {
            var pr = ft.properties || {};
            f = { type: pr.type, score: pr.score, housenumber: pr.housenumber, street: pr.street, name: pr.name,
              postcode: pr.postcode, city: pr.city, lon: ft.geometry.coordinates[0], lat: ft.geometry.coordinates[1] };
          }
        } else if (r) _adrDiag_("BAN unitaire HTTP " + r.getResponseCode());
      } catch (e2) {}
      out.push(f);
    });
    Utilities.sleep(600);
  }
  return out;
}

// ---------------------------------------------------------------------------
// Adresse de référence d'une entreprise (lignes déjà lues) : une ligne vérifiée
// l'emporte, sinon la plus complète / la plus fréquente.
function _adrReference_(lignes) {
  var verif = lignes.filter(function (a) { return ADR_STATUTS_OK[a.statut] && a.lat; })[0];
  if (verif) return verif;
  var freq = {};
  lignes.forEach(function (a) { var k = _adrNorm_(a.adresse + "|" + a.cp + "|" + a.ville); freq[k] = (freq[k] || 0) + 1; });
  var best = null, bestScore = -1;
  lignes.forEach(function (a) {
    var s = (a.adresse ? 4 : 0) + (a.cp ? 2 : 0) + (a.ville ? 1 : 0) + (a.lat ? 0.5 : 0) + freq[_adrNorm_(a.adresse + "|" + a.cp + "|" + a.ville)] * 0.1;
    if (s > bestScore) { bestScore = s; best = a; }
  });
  return best || { adresse: "", cp: "", ville: "", complement: "", lat: "", lon: "", statut: "", siret: "" };
}

// Adresse actuelle d'une entreprise (pour un nouveau contact : il hérite de cette adresse)
function _adrEntreprise_(ss, idEnt) {
  var ws = getSheetSafe(ss, "PARTENARIAT");
  var data = ws.getDataRange().getValues();
  var C = _adrColsLecture_(data[0]);
  var id = String(idEnt || "").trim().toUpperCase(), lignes = [];
  for (var i = 1; i < data.length; i++) if (String(data[i][1]).trim().toUpperCase() === id) lignes.push(_adrDeLigne_(data[i], C));
  return lignes.length ? _adrReference_(lignes) : null;
}

// Adresse fournie par la page (création / modification) → adresse à enregistrer.
// Statut BAN/SIRENE + coordonnées envoyés par la page = choisie dans le référentiel.
// Sinon tentative de géocodage serveur ; à défaut « Saisie libre ».
function _adrDepuisSaisie_(p) {
  var a = {
    adresse: _adrVide_(p.adresse), cp: _adrVide_(p.cp), ville: _adrVide_(p.ville),
    complement: _adrVide_(p.complement), lat: Number(p.lat) || "", lon: Number(p.lon) || "",
    statut: String(p.adrStatut || "").trim(), siret: String(p.siret || "").replace(/\D/g, "").slice(0, 14)
  };
  if (!a.adresse && !a.cp && !a.ville) return { adresse: "", cp: "", ville: "", complement: "", lat: "", lon: "", statut: "", siret: a.siret };
  if (ADR_STATUTS_OK[a.statut] && a.lat && a.lon) return a;
  var g = _adrGeocoderUne_(a);
  if (g) {
    var r = { adresse: g.adresse || a.adresse, cp: g.cp, ville: g.ville, complement: a.complement || g.complement || "",
      lat: g.lat, lon: g.lon, statut: g.statut, siret: a.siret };
    if (g.statut === "Commune") { r.adresse = a.adresse; r.statut = a.adresse ? "Saisie libre" : "Commune"; }
    return r;
  }
  a.statut = "Saisie libre"; a.lat = ""; a.lon = "";
  return a;
}

// Écrit l'adresse « a » dans une ligne (tableau) — K, L, M + colonnes étendues
function _adrDansLigne_(row, C, a) {
  var need = Math.max(13, C.complement + 1, C.lat + 1, C.lon + 1, C.statut + 1, C.siret + 1, C.origine + 1);
  while (row.length < need) row.push("");
  row[10] = a.adresse || ""; row[11] = a.cp || ""; row[12] = a.ville || "";
  row[C.complement] = a.complement || ""; row[C.lat] = a.lat || ""; row[C.lon] = a.lon || "";
  row[C.statut] = a.statut || ""; row[C.siret] = a.siret ? "'" + a.siret : "";
  return row;
}

function _adrTexte_(a) {
  return a ? [a.adresse, a.complement, [a.cp, a.ville].filter(Boolean).join(" ")].filter(Boolean).join(", ") : "";
}

// ---------------------------------------------------------------------------
// Postes des contacts : liste type + tous les postes déjà saisis (dédoublonnés)
var POSTES_TYPES = [
  "Gérant", "Dirigeant", "Directeur général", "Directeur administratif et financier", "Expert comptable",
  "Commissaire aux comptes", "Chef de mission", "Collaborateur comptable", "Responsable comptable",
  "Comptable", "Responsable paie", "Gestionnaire de paie", "RH", "Responsable RH", "Chargé de recrutement",
  "Assistant de direction", "Office manager", "Standard"
];
function handleGetPostesContacts(p, ss) {
  return getCachedJsonResponse_("postes_contacts_v1", 300, function () {
    var data = getSheetSafe(ss, "PARTENARIAT").getDataRange().getValues();
    var vus = {}, liste = [];
    function ajouter(v, n) {
      v = String(v || "").replace(/\s+/g, " ").trim();
      if (!v || /^[x.\-]+$/i.test(v)) return;
      var k = _adrNorm_(v);
      if (!vus[k]) { vus[k] = { poste: v.charAt(0).toUpperCase() + v.slice(1), nb: 0 }; liste.push(vus[k]); }
      vus[k].nb += n;
    }
    POSTES_TYPES.forEach(function (v) { ajouter(v, 0); });
    for (var i = 1; i < data.length; i++) ajouter(data[i][13], 1);
    liste.sort(function (a, b) { return b.nb - a.nb || a.poste.localeCompare(b.poste, "fr"); });
    return createJsonResponse({ success: true, data: liste });
  });
}

// ---------------------------------------------------------------------------
// NORMALISATION DE TOUTE LA BASE (bouton administrateur de partenaires.html,
// ou exécuter harmoniserAdressesEntreprises() depuis l'éditeur).
//  1. une seule adresse par entreprise (recopiée sur tous ses contacts) ;
//  2. géocodage BAN en lot : adresse officielle au numéro (statut BAN), à la voie,
//     sinon position de la commune (statut Commune) ou « À vérifier » ;
//  3. le texte saisi avant normalisation est conservé dans « Adresse d'origine ».
function harmoniserAdressesEntreprises() {   // depuis l'éditeur Apps Script
  var ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");
  _adrAuditDemarrer_("Éditeur Apps Script", "Administrateur", false);
}

function _adrDistanceM_(a, b) {
  var R = 6371000, t = Math.PI / 180;
  var dLat = (b.lat - a.lat) * t, dLon = (b.lon - a.lon) * t;
  var x = Math.sin(dLat / 2) * Math.sin(dLat / 2) + Math.cos(a.lat * t) * Math.cos(b.lat * t) * Math.sin(dLon / 2) * Math.sin(dLon / 2);
  return 2 * R * Math.asin(Math.sqrt(x));
}

// AUDIT + HARMONISATION. Règle (demande direction, 2026-11-09) :
//  • décalage LÉGER (orthographe, abréviation, CP dans la même ville, adresse absente sur
//    certains contacts) → adresse officielle BAN appliquée à tous les contacts ;
//  • cas COMPLEXE (voie introuvable, ville différente, adresses contradictoires entre contacts)
//    → l'adresse saisie n'est PAS modifiée, seul le statut d'alerte est posé
//    (« À vérifier » ou « Adresses multiples »), affiché sur l'application.
// Planification pure (testable hors Apps Script) : groupes → décisions
function _adrDecider_(groupes, ordre, geocoderLot, forcer) {
  var items1 = [], items2 = [];
  ordre.forEach(function (id) {
    var g = groupes[id], vus = {};
    g.cands = [];
    g.lignes.forEach(function (a) {
      if (!a.adresse && !a.cp && !a.ville) return;
      var cp = /^\d{5}$/.test(a.cp) ? a.cp : (/^\d{4}$/.test(a.cp) ? "0" + a.cp : "");
      var sa = { adresse: _adrNettoyer_(a.adresse, cp), cp: cp, ville: a.ville, brut: a };
      var k = _adrNorm_(sa.adresse + "|" + sa.cp + "|" + sa.ville);
      if (vus[k]) { vus[k].nb++; return; }
      sa.nb = 1; vus[k] = sa; g.cands.push(sa);
    });
    var verif = g.lignes.filter(function (a) { return ADR_STATUTS_OK[a.statut] && a.lat; })[0];
    if (verif && !forcer) { g.decision = { type: "ok", a: verif }; return; }   // déjà validée (fenêtre d'adresse)
    if (!g.cands.length) { g.decision = { type: "vide" }; return; }
    g.cands.forEach(function (c) {
      if (c.adresse) items1.push({ c: c, q: [c.adresse, c.cp, c.ville].filter(Boolean).join(" ") });
      items2.push({ c: c, q: [c.ville, c.cp ? c.cp.slice(0, 2) : ""].filter(Boolean).join(" ") || c.cp });
    });
  });
  var r1 = geocoderLot(items1, false);
  items1.forEach(function (it, k) { it.c.res = _adrEvaluer_(r1[k], it.c); it.c.ban = r1[k]; });
  items2 = items2.filter(function (it) { return !it.c.res; });   // commune utile seulement si la rue n'est pas validée
  var r2 = geocoderLot(items2, true);
  items2.forEach(function (it, k) { it.c.commune = _adrEvaluerCommune_(r2[k], it.c); });

  ordre.forEach(function (id) {
    var g = groupes[id];
    if (g.decision) return;
    var avecRue = g.cands.filter(function (c) { return c.adresse; });
    var sansRue = g.cands.filter(function (c) { return !c.adresse; });
    var compl = (g.lignes.filter(function (a) { return a.complement; })[0] || {}).complement || "";
    var siret = (g.lignes.filter(function (a) { return a.siret; })[0] || {}).siret || "";
    if (avecRue.length) {
      var ok = avecRue.filter(function (c) { return c.res; });
      var memeLieu = ok.length === avecRue.length && ok.every(function (c) { return _adrDistanceM_(c.res, ok[0].res) <= 150; });
      var villeRef = ok.length ? ok[0].res.ville : avecRue[0].ville;
      var villesOk = sansRue.every(function (c) { return !c.ville || _adrVilleCompatible_(c.ville, villeRef); });
      if (ok.length && memeLieu && villesOk) {
        ok.sort(function (a, b) { return (b.res.statut === "BAN") - (a.res.statut === "BAN") || b.nb - a.nb; });
        var r = ok[0].res;
        g.decision = { type: "harmonise", a: { adresse: r.adresse, cp: r.cp, ville: r.ville, complement: compl || r.complement, lat: r.lat, lon: r.lon, statut: r.statut, siret: siret } };
      } else if (avecRue.length === 1 && villesOk) {
        var cm = avecRue[0].commune, b = avecRue[0].ban;
        g.decision = { type: "alerte", statut: "À vérifier", lat: cm ? cm.lat : "", lon: cm ? cm.lon : "",
          raison: b && b.type !== "municipality" ? "écart trop important avec la Base Adresse Nationale (proposition : " + _adrCasse_(String((b.housenumber ? b.housenumber + " " : "") + (b.street || b.name)).replace(/\s+/g, " ")) + ", " + b.postcode + " " + b.city + ")" : "voie introuvable dans la Base Adresse Nationale" };
      } else {
        g.decision = { type: "alerte", statut: "Adresses multiples", lat: "", lon: "",
          raison: g.cands.length + " adresses différentes selon les contacts" };
      }
    } else {
      var cms = sansRue.filter(function (c) { return c.commune; });
      var memeVille = cms.length === sansRue.length && cms.every(function (c) { return _adrNorm_(c.commune.ville) === _adrNorm_(cms[0].commune.ville); });
      if (memeVille) {
        var c0 = cms[0].commune;
        g.decision = { type: "harmonise", a: { adresse: "", cp: c0.cp, ville: c0.ville, complement: compl, lat: c0.lat, lon: c0.lon, statut: "Commune", siret: siret } };
      } else if (sansRue.length === 1) {
        g.decision = { type: "alerte", statut: "À vérifier", lat: "", lon: "", raison: "commune introuvable" };
      } else {
        g.decision = { type: "alerte", statut: "Adresses multiples", lat: "", lon: "", raison: sansRue.length + " communes différentes selon les contacts" };
      }
    }
  });
}

// ---------------------------------------------------------------------------
// RECALAGE AUTOMATIQUE des adresses « à vérifier » (demande direction, 2026-11-09 bis) :
//  1. SIRENE : on cherche l'entreprise par son NOM (dans le département) ; un établissement
//     actif est retenu si sa rue correspond à la saisie, ou si sa commune correspond et
//     qu'il est le seul candidat → adresse officielle INSEE + SIRET (statut SIRENE) ;
//  2. sinon la proposition de la Base Adresse Nationale est acceptée même si la ville saisie
//     était fausse (même rue, même numéro, même département) → statut BAN / BAN voie ;
//  3. commune mal orthographiée / mauvais CP → position de la commune (statut Commune) ;
//  4. « Adresses multiples » : recalée seulement si SIRENE désigne l'une des adresses saisies
//     (sinon alerte conservée, mais positionnée sur la carte à l'adresse la plus fréquente).
// Le texte d'origine reste dans « Adresse d'origine » et chaque recalage figure dans l'onglet
// « Audit adresses ».
var ADR_SIRENE_URL = "https://recherche-entreprises.api.gouv.fr/search";
var ADR_NOM_VIDES = { sas: 1, sasu: 1, sarl: 1, eurl: 1, sa: 1, sci: 1, scp: 1, selarl: 1, selas: 1, selafa: 1, snc: 1,
  cabinet: 1, groupe: 1, societe: 1, ste: 1, et: 1, de: 1, du: 1, des: 1, la: 1, le: 1, les: 1, l: 1, d: 1, en: 1,
  expert: 1, experts: 1, expertise: 1, comptable: 1, comptables: 1, comptabilite: 1, conseil: 1, conseils: 1,
  audit: 1, associes: 1, associe: 1, cie: 1, agence: 1, france: 1 };

function _adrMotsNom_(nom) {
  var t = _adrNorm_(String(nom || "").replace(/\(.*?\)/g, " ")).split(/[^a-z0-9]+/).filter(function (w) { return w.length >= 2; });
  var s = t.filter(function (w) { return !ADR_NOM_VIDES[w]; });
  return s.length ? s : t;
}
function _adrNomCorrespond_(nom, r) {
  var mots = _adrMotsNom_(nom);
  if (!mots.length) return false;
  var cible = " " + _adrNorm_([r.nom_complet, r.nom_raison_sociale, r.sigle,
    (r.siege || {}).nom_commercial, ((r.siege || {}).liste_enseignes || []).join(" ")]
    .concat((r.matching_etablissements || []).map(function (e) { return [e.nom_commercial, (e.liste_enseignes || []).join(" ")].join(" "); }))
    .filter(Boolean).join(" ")).replace(/[^a-z0-9]+/g, " ") + " ";
  var n = mots.filter(function (w) { return cible.indexOf(" " + w + " ") !== -1 || (w.length >= 5 && cible.indexOf(" " + w.slice(0, 5)) !== -1); }).length;
  return n / mots.length >= 0.6;
}
// Établissements actifs d'un résultat SIRENE → [{ adresse (texte brut), voie, cp, ville, lat, lon, siret }]
function _adrEtablissements_(r) {
  var liste = (r.matching_etablissements || []).slice();
  if (r.siege) liste.push(r.siege);
  var vus = {}, out = [];
  liste.forEach(function (e) {
    if (!e || e.etat_administratif === "F" || !e.adresse || vus[e.siret]) return;
    vus[e.siret] = 1;
    var cp = String(e.code_postal || "");
    if (!/^\d{5}$/.test(cp) || /DIFFUSIBLE/i.test(e.adresse)) return;   // établissement non diffusible
    var voie = String(e.adresse).replace(new RegExp("\\s+" + cp + "\\b.*$"), "").trim();
    out.push({ brut: String(e.adresse), voie: voie, cp: cp, ville: String(e.libelle_commune || ""),
      lat: Number(e.latitude), lon: Number(e.longitude), siret: String(e.siret || "") });
  });
  return out;
}

// Recherche SIRENE en lot (7 requêtes/s max côté API) : items = [{ q, dep }] → [résultats[]]
function _adrSireneLot_(items) {
  var out = [];
  for (var s = 0; s < items.length; s += 6) {
    if (!_adrTempsOk_()) { _adrDiag_("temps limite atteint : " + (items.length - s) + " requête(s) non faite(s)"); while (out.length < items.length) out.push([]); break; }
    var part = items.slice(s, s + 6);
    var reqs = part.map(function (it) {
      return { url: ADR_SIRENE_URL + "?per_page=5&q=" + encodeURIComponent(String(it.q).slice(0, 80)) + (it.dep ? "&departement=" + it.dep : ""), muteHttpExceptions: true };
    });
    var rep = [];
    rep = _adrFetchAll_(reqs);
    part.forEach(function (it, k) {
      var res = [];
      try {
        var r = rep[k];
        if (r && (r.getResponseCode() === 429 || r.getResponseCode() >= 500)) { Utilities.sleep(1500); try { r = UrlFetchApp.fetch(reqs[k].url, { muteHttpExceptions: true }); } catch (eR) {} }
        if (r && r.getResponseCode() === 200) res = JSON.parse(r.getContentText()).results || [];
        else if (r) _adrDiag_("SIRENE HTTP " + r.getResponseCode());
      } catch (e2) {}
      out.push(res);
    });
    Utilities.sleep(1100);
  }
  return out;
}

// Commune dans le bon département parmi les propositions BAN (features GeoJSON)
function _adrChoisirCommune_(features, it) {
  var mot = _adrNorm_(String(it.ville || "").split(/[\s\-']+/)[0] || "");
  for (var i = 0; i < (features || []).length; i++) {
    var ft = features[i], pr = ft.properties || {};
    if (pr.type !== "municipality") continue;
    if (it.dep && String(pr.postcode).slice(0, 2) !== it.dep) continue;
    if (mot.length >= 3 && _adrNorm_(pr.city).indexOf(mot.slice(0, Math.min(mot.length, 5))) !== 0
      && !_adrVilleCompatible_(it.ville, pr.city)) continue;
    return { city: pr.city, postcode: String(pr.postcode), lon: ft.geometry.coordinates[0], lat: ft.geometry.coordinates[1] };
  }
  return null;
}
function _adrCommuneLot_(items) {
  var out = [];
  for (var s = 0; s < items.length; s += 20) {
    if (!_adrTempsOk_()) { _adrDiag_("temps limite atteint : " + (items.length - s) + " requête(s) non faite(s)"); while (out.length < items.length) out.push(null); break; }
    var part = items.slice(s, s + 20);
    var reqs = part.map(function (it) {
      var q = (it.ville || it.cp || "").slice(0, 80);
      return { url: ADR_BAN_URL + "?type=municipality&limit=10&q=" + encodeURIComponent(q.length >= 3 ? q : it.cp), muteHttpExceptions: true };
    });
    var rep = [];
    rep = _adrFetchAll_(reqs);
    part.forEach(function (it, k) {
      var f = null;
      try { if (rep[k] && rep[k].getResponseCode() === 200) f = _adrChoisirCommune_(JSON.parse(rep[k].getContentText()).features, it); } catch (e2) {}
      out.push(f);
    });
    Utilities.sleep(300);
  }
  return out;
}

function _adrRecaler_(groupes, ordre, sireneLot, geocoderLot, communeLot, inclureCommune) {
  var cibles = ordre.filter(function (id) { var d = groupes[id].decision;
    return d && (d.type === "alerte" || (inclureCommune && d.type === "harmonise" && d.a && d.a.statut === "Commune")); });
  if (!cibles.length) return;

  // 1. SIRENE par nom d'entreprise
  var reqS = [];
  cibles.forEach(function (id) {
    var g = groupes[id];
    var cpRef = (g.cands.filter(function (c) { return c.cp; })[0] || {}).cp || "";
    var dep = cpRef ? cpRef.slice(0, 2) : "";
    if (dep === "20") dep = "";
    if (_adrMotsNom_(g.nom).length) reqS.push({ g: g, q: g.nom.replace(/\(.*?\)/g, " ").replace(/\s+/g, " ").trim(), dep: dep });
  });
  var resS = reqS.length ? sireneLot(reqS) : [];
  var aGeocoder = [];
  reqS.forEach(function (it, k) {
    var g = it.g, choix = [];
    (resS[k] || []).forEach(function (r) {
      if (!_adrNomCorrespond_(g.nom, r)) return;
      _adrEtablissements_(r).forEach(function (e) {
        if (it.dep && e.cp.slice(0, 2) !== it.dep) return;
        var rue = g.cands.some(function (c) { return c.adresse && _adrRueCompatible_(e.voie.replace(/^[\d\s\-,]+(bis|ter)?\s*/i, ""), c.adresse); });
        var ville = g.cands.some(function (c) { return (c.cp && c.cp === e.cp) || (c.ville && _adrVilleCompatible_(c.ville, e.ville)); });
        if (rue || ville) choix.push({ e: e, rue: rue, siren: r.siren });
      });
    });
    var parRue = choix.filter(function (c) { return c.rue; });
    var retenu = null;
    if (parRue.length) retenu = parRue[0];
    else if (choix.length && g.decision.statut !== "Adresses multiples") {   // seule la commune concorde : un seul établissement possible
      var lieux = choix.filter(function (c, i) { return isFinite(c.e.lat) && choix.every(function (d) { return !isFinite(d.e.lat) || _adrDistanceM_(c.e, d.e) <= 150; }); });
      if (lieux.length === choix.length && choix.every(function (c) { return c.siren === choix[0].siren; })) retenu = choix[0];
    }
    if (retenu) { g.sirene = retenu.e; aGeocoder.push({ g: g, q: retenu.e.brut }); }
  });
  var resG = aGeocoder.length ? geocoderLot(aGeocoder, false) : [];
  aGeocoder.forEach(function (it, k) {
    var g = it.g, e = g.sirene;
    var compl = (g.lignes.filter(function (a) { return a.complement; })[0] || {}).complement || "";
    var ev = _adrEvaluer_(resG[k], { adresse: e.voie, cp: e.cp, ville: e.ville });
    var a = ev ? { adresse: ev.adresse, cp: ev.cp, ville: ev.ville, complement: compl || ev.complement, lat: ev.lat, lon: ev.lon }
      : isFinite(e.lat) ? (function () {   // « ARTEPARC BAT D RUE DE … » → complément + voie
        var m = e.voie.match(/(\d{1,4}\s?(?:bis|ter|[a-d])?\s+)?\b(?:rue|av|avenue|bd|boulevard|chemin|route|place|all[ée]e?s?|impasse|traverse|quai|cours|mont[ée]e|square|passage|esplanade)\b.*$/i);
        var voie = m ? m[0] : e.voie, avant = m ? e.voie.slice(0, m.index).trim() : "";
        return { adresse: _adrCasse_(voie), cp: e.cp, ville: _adrCasse_(e.ville), complement: compl || _adrCasse_(avant), lat: e.lat, lon: e.lon };
      })() : null;
    if (!a) return;
    a.statut = "SIRENE"; a.siret = e.siret;
    g.decision = { type: "recale", source: "SIRENE", a: a, raisonAvant: g.decision.raison };
  });

  // 2-3. Base Adresse Nationale, ville saisie ignorée (même département)
  function infos(g) {
    return { compl: (g.lignes.filter(function (a) { return a.complement; })[0] || {}).complement || "",
      siret: (g.lignes.filter(function (a) { return a.siret; })[0] || {}).siret || "" };
  }
  function tenter(g, f, c, source) {   // résultat BAN f accepté si même rue / numéro / département
    if (!f || f.type === "municipality" || !(f.score >= 0.4)) return false;
    if (c.cp && String(f.postcode).slice(0, 2) !== c.cp.slice(0, 2)) return false;
    var ev = _adrEvaluer_(f, { adresse: c.adresse, cp: "", ville: f.city });
    if (!ev) return false;
    var x = infos(g);
    g.decision = { type: "recale", source: source, raisonAvant: g.decision.raison,
      a: { adresse: ev.adresse, cp: ev.cp, ville: ev.ville, complement: x.compl || ev.complement, lat: ev.lat, lon: ev.lon, statut: ev.statut, siret: x.siret } };
    return true;
  }
  var epurees = [], communes = [];
  cibles.forEach(function (id) {
    var g = groupes[id], d = g.decision;
    if (d.type !== "alerte" || d.statut === "Adresses multiples") return;
    var c = g.cands[0];
    if (!c) return;
    if (c.adresse && tenter(g, c.ban, c, "BAN")) return;
    if (c.adresse) {   // 2e essai : seulement « numéro + voie » (sans ZAC, lieu-dit, BP, bâtiment, ville…)
      var m = c.adresse.match(/(\d{1,4}\s?(?:bis|ter|[a-d])?\s*,?\s+)?\b(?:rue|av|ave|avenue|bd|bld|blvd|boulevard|chemin|chem|ch|route|rte|place|pl|all[ée]e?s?|impasse|imp|traverse|trav|quai|cours|crs|mont[ée]e|square|passage|esplanade|parvis|rond[- ]point|voie|zone)\b[^,;]*/i);
      if (m) {
        var v = m[0].replace(/\s+\b(zac|za|zi|zone|lieu[- ]?dit|ld|lot|lotissement|bp|cs|tsa|box|b[aâ]t(iment)?|imm(euble)?|r[ée]sidence|r[ée]s|parc|centre|cedex)\b.*$/i, "").replace(/\s+/g, " ").trim();
        if (v.length >= 6 && _adrNorm_(v) !== _adrNorm_(c.adresse)) { epurees.push({ g: g, c: c, q: v + " " + (c.cp || c.ville || "") }); return; }
      }
    }
    communes.push({ g: g, c: c });
  });
  var resE = epurees.length ? geocoderLot(epurees, false) : [];
  epurees.forEach(function (it, k) { if (!tenter(it.g, resE[k], it.c, "BAN")) communes.push(it); });

  // Commune : recherche dans le département (« Salon » 13000 → Salon-de-Provence 13300)
  communes = communes.filter(function (it) { return it.c.ville || it.c.cp; });
  var resC = communes.length ? communeLot(communes.map(function (it) {
    return { ville: it.c.ville, cp: it.c.cp, dep: it.c.cp ? it.c.cp.slice(0, 2) : "" };
  })) : [];
  communes.forEach(function (it, k) {
    var f = resC[k], g = it.g, c = it.c, d = g.decision;
    if (!f) return;
    if (c.adresse) { d.lat = f.lat; d.lon = f.lon; return; }   // rue douteuse : alerte conservée, placée sur la bonne commune
    var x = infos(g);
    g.decision = { type: "recale", source: "BAN commune", raisonAvant: d.raison,
      a: { adresse: "", cp: String(f.postcode), ville: f.city, complement: x.compl, lat: f.lat, lon: f.lon, statut: "Commune", siret: x.siret } };
  });

  // 4. Adresses multiples non tranchées : position de l'adresse la plus fréquente pour la carte
  cibles.forEach(function (id) {
    var g = groupes[id], d = g.decision;
    if (d.type !== "alerte" || d.lat) return;
    var tri = g.cands.slice().sort(function (a, b) { return b.nb - a.nb; });
    for (var i = 0; i < tri.length; i++) {
      var p = tri[i].res || tri[i].commune;
      if (p && isFinite(p.lat)) { d.lat = p.lat; d.lon = p.lon; break; }
    }
  });
}

function _adrHarmoniser_(ss, forcer, silencieux, plage) {
  var ws = getSheetSafe(ss, "PARTENARIAT");
  var C = _adrCols_(ws);
  _ADR_T0_ = Date.now(); _ADR_DIAG_ = [];
  if (plage && plage.csvKo) _ADR_CSV_KO_ = true;
  if (!plage || !plage.debut) _adrSonde_();
  var data = ws.getDataRange().getValues();
  var nCols = data[0].length;
  var colsMin = Math.min(C.complement, C.lat, C.lon, C.statut, C.siret, C.origine);
  var colsMax = Math.max(C.complement, C.lat, C.lon, C.statut, C.siret, C.origine);
  for (var r0 = 1; r0 < data.length; r0++) while (data[r0].length < Math.max(nCols, colsMax + 1)) data[r0].push("");

  var groupes = {}, ordre = [];
  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][1] || "").trim().toUpperCase();
    if (!id) continue;
    if (!groupes[id]) { groupes[id] = { id: id, nom: String(data[i][2] || "").trim(), rows: [], lignes: [] }; ordre.push(id); }
    groupes[id].rows.push(i);
    groupes[id].lignes.push(_adrDeLigne_(data[i], C));
  }
  var total = ordre.length;
  if (plage && plage.ids) {   // recalage ciblé (résultats SIRENE fournis par le navigateur)
    var voulu = {};
    plage.ids.forEach(function (x) { voulu[String(x).trim().toUpperCase()] = 1; });
    ordre = ordre.filter(function (id) { return voulu[id]; });
  } else if (plage) ordre = ordre.slice(plage.debut, plage.debut + plage.taille);   // traitement par lots (limite 6 min)
  _adrDecider_(groupes, ordre, _adrGeocoderLot_, forcer);
  var sireneLot = plage && plage.sirene
    ? function (items) { return items.map(function (it) { return plage.sirene[it.g.id] || []; }); }
    : _adrSireneLot_;
  _adrRecaler_(groupes, ordre, sireneLot, _adrGeocoderLot_, _adrCommuneLot_, !!(plage && plage.sirene));
  var rapport = { total: total, suite: !!plage && plage.debut + plage.taille < total, entreprises: ordre.length, dejaOk: 0, harmonisees: 0, ban: 0, voie: 0, commune: 0, aVerifier: 0, multiples: 0,
    sansAdresse: 0, lignesModifiees: 0, recalees: 0, recalSirene: 0, recalBan: 0, recalCommune: 0, exemples: [], recalages: [], alertes: [] };

  ordre.forEach(function (id) {
    var g = groupes[id], d = g.decision;
    var avantTxt = (g.cands || []).map(function (c) { return _adrTexte_(c.brut); }).join(" / ");
    var modifie = false;
    g.rows.forEach(function (ri) {
      var row = data[ri];
      var avant = row.slice(10, 13).concat([row[C.complement], row[C.lat], row[C.lon], row[C.statut], row[C.origine]]).map(String).join("|");
      if (d.type === "harmonise" || d.type === "ok" || d.type === "recale") {
        var a = d.a;
        var origine = row[C.origine];
        var txtLigne = [row[10], row[11], row[12]].map(function (x) { return String(x || "").trim(); }).filter(Boolean).join(", ");
        if (!origine && txtLigne && _adrNorm_(txtLigne) !== _adrNorm_(_adrTexte_(a))) origine = txtLigne;
        _adrDansLigne_(row, C, a);
        row[C.origine] = origine || "";
      } else if (d.type === "alerte") {   // adresse saisie conservée telle quelle
        row[C.statut] = d.statut; row[C.lat] = d.lat || ""; row[C.lon] = d.lon || "";
      }
      var apres = row.slice(10, 13).concat([row[C.complement], row[C.lat], row[C.lon], row[C.statut], row[C.origine]]).map(String).join("|");
      if (avant !== apres) { rapport.lignesModifiees++; modifie = true; }
    });
    if (d.type === "vide") rapport.sansAdresse++;
    else if (d.type === "ok") rapport.dejaOk++;
    else if (d.type === "harmonise") {
      rapport.harmonisees++;
      if (d.a.statut === "BAN") rapport.ban++; else if (d.a.statut === "BAN voie") rapport.voie++; else rapport.commune++;
      if (modifie && _adrNorm_(avantTxt) !== _adrNorm_(_adrTexte_(d.a)) && rapport.exemples.length < 300)
        rapport.exemples.push({ id: id, nom: g.nom, avant: avantTxt || "—", apres: _adrTexte_(d.a), statut: d.a.statut });
    } else if (d.type === "recale") {
      rapport.recalees++;
      if (d.source === "SIRENE") rapport.recalSirene++; else if (d.source === "BAN") rapport.recalBan++; else rapport.recalCommune++;
      if (rapport.recalages.length < 500)
        rapport.recalages.push({ id: id, nom: g.nom, avant: avantTxt || "—", apres: _adrTexte_(d.a) + (d.a.siret ? " (SIRET " + d.a.siret + ")" : ""), source: d.source, raison: d.raisonAvant || "" });
    } else {
      if (d.statut === "Adresses multiples") rapport.multiples++; else rapport.aVerifier++;
      if (rapport.alertes.length < 500) rapport.alertes.push({ id: id, nom: g.nom, adresse: avantTxt, statut: d.statut, raison: d.raison });
    }
  });

  var rMin = data.length, rMax = 0;   // écriture limitée aux lignes des entreprises du lot
  ordre.forEach(function (id) { groupes[id].rows.forEach(function (ri) { if (ri < rMin) rMin = ri; if (ri > rMax) rMax = ri; }); });
  var bloc1 = [], bloc2 = [];
  for (var r = rMin; r <= rMax; r++) {
    bloc1.push([data[r][10], data[r][11], data[r][12]]);
    bloc2.push(data[r].slice(colsMin, colsMax + 1));
  }
  if (bloc1.length) {
    ws.getRange(rMin + 1, 11, bloc1.length, 3).setValues(bloc1);
    ws.getRange(rMin + 1, colsMin + 1, bloc2.length, colsMax - colsMin + 1).setValues(bloc2);
  }
  rapport.diag = _ADR_DIAG_.slice();
  if (!plage && (!silencieux || rapport.lignesModifiees)) _adrRapportFeuille_(ss, rapport);
  return rapport;
}

// Rapport lisible dans l'onglet « Audit adresses » (recréé à chaque audit)
function _adrRapportFeuille_(ss, R) {
  try {
    var ws = ss.getSheetByName("Audit adresses") || ss.insertSheet("Audit adresses");
    ws.clear();
    var rows = [["Audit des adresses entreprises — " + Utilities.formatDate(new Date(), "Europe/Paris", "dd/MM/yyyy HH:mm"), "", "", "", ""],
      [R.entreprises + " entreprises · " + R.harmonisees + " harmonisées · " + (R.recalees || 0) + " recalées automatiquement · " + R.dejaOk + " déjà conformes · " + R.aVerifier + " à vérifier · " + R.multiples + " à adresses multiples · " + R.sansAdresse + " sans adresse", "", "", "", ""],
      [_ADR_DIAG_.length ? "Incidents techniques : " + _ADR_DIAG_.join(" | ") : "", "", "", "", ""], ["Type", "ID entreprise", "Entreprise", "Adresse avant", "Adresse après / motif de l'alerte"]];
    R.alertes.forEach(function (a) { rows.push(["⚠️ " + a.statut, a.id, a.nom, a.adresse, a.raison]); });
    (R.recalages || []).forEach(function (e) { rows.push(["🔁 Recalée (" + e.source + ")", e.id, e.nom, e.avant, e.apres + (e.raison ? " — était : " + e.raison : "")]); });
    R.exemples.forEach(function (e) { rows.push(["✅ Harmonisée (" + e.statut + ")", e.id, e.nom, e.avant, e.apres]); });
    ws.getRange(1, 1, rows.length, 5).setValues(rows);
    ws.getRange(1, 1).setFontWeight("bold").setFontSize(13);
    ws.getRange(5, 1, 1, 5).setFontWeight("bold").setBackground("#1A4E8A").setFontColor("#ffffff");
    ws.setFrozenRows(5);
    ws.setColumnWidths(1, 1, 190); ws.setColumnWidths(2, 1, 120); ws.setColumnWidths(3, 1, 240); ws.setColumnWidths(4, 2, 380);
  } catch (e) { console.error("rapport audit : " + e); }
}

function handleHarmoniserAdresses(p, ss) {
  var isAdmin = String(p.roleAuteur || "") === "Administrateur" || PARAM_MASTER_EMAILS.indexOf(String(p.emailAuteur || "").toLowerCase()) !== -1;
  if (!isAdmin) return createJsonResponse({ success: false, message: "Réservé aux administrateurs." });
  var etat = _adrEtat_();
  if (etat && etat.enCours && Date.now() - etat.maj < 15 * 60 * 1000)
    return createJsonResponse({ success: true, enCours: true, message: "Un audit est déjà en cours (" + etat.debut + " / " + etat.total + " entreprises traitées)." });
  _adrAuditDemarrer_(p.nomAuteur || p.idAuteur || "Administrateur", p.roleAuteur || "Administrateur", false);
  return createJsonResponse({ success: true, enCours: true, message: "Audit lancé en arrière-plan : traitement par lots de " + ADR_LOT + " entreprises (environ 1 minute par lot). Résultat dans l'onglet « Audit adresses » et dans l'Historique." });
}

// ---------------------------------------------------------------------------
// AUDIT PAR LOTS : chaque lot (ADR_LOT entreprises) tourne dans sa propre exécution
// (déclencheur ponctuel), pour ne jamais dépasser la limite de 6 min d'Apps Script.
// L'état cumulé est conservé dans la propriété ADR_AUDIT_ETAT.
var ADR_LOT = 150;
function _adrEtat_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty("ADR_AUDIT_ETAT") || "null"); } catch (e) { return null; }
}
function _adrEtatSauver_(e) {
  e.maj = Date.now();
  PropertiesService.getScriptProperties().setProperty("ADR_AUDIT_ETAT", JSON.stringify(e));
}
function _adrAuditDemarrer_(auteur, role, silencieux) {
  _adrEtatSauver_({ enCours: true, debut: 0, total: 0, auteur: auteur, role: role, silencieux: !!silencieux, lance: Date.now(),
    c: { dejaOk: 0, harmonisees: 0, ban: 0, voie: 0, commune: 0, aVerifier: 0, multiples: 0, sansAdresse: 0,
      lignesModifiees: 0, recalees: 0, recalSirene: 0, recalBan: 0, recalCommune: 0 }, ex: [], rc: [], diag: [] });
  _adrTickInstaller_();
}
// UN SEUL déclencheur récurrent (toutes les 5 min) pour tout le module adresses : le projet est
// à la limite de 20 déclencheurs par script, les déclencheurs ponctuels échouaient sans bruit.
function _adrTickInstaller_() {
  try {
    var tick = false;
    ScriptApp.getProjectTriggers().forEach(function (t) {
      var h = t.getHandlerFunction();
      if (h === "auditAdressesLot" || h === "auditAdressesAuto" || h === "auditAdressesNocturne") ScriptApp.deleteTrigger(t);
      if (h === "auditAdressesTick") tick = true;
    });
    if (!tick) ScriptApp.newTrigger("auditAdressesTick").timeBased().everyMinutes(5).create();
  } catch (e) {
    var msg = "Impossible de programmer l'audit des adresses : " + e;
    console.error(msg);
    PropertiesService.getScriptProperties().setProperty("ADR_AUDIT_ERREUR", new Date().toISOString() + " " + msg);
    try {
      var ss = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");
      var ws = ss.getSheetByName("Audit adresses") || ss.insertSheet("Audit adresses");
      ws.getRange(3, 1).setValue("⛔ " + msg + " (" + ScriptApp.getProjectTriggers().length + " déclencheurs dans le projet)");
    } catch (e2) {}
  }
}
function auditAdressesTick() {
  var t0 = Date.now();
  try { _qualiopiPackAuto_(); } catch (eQ) { console.error('pack Qualiopi : ' + eQ); }   // import unique (qualiopi.gs)
  try { if (typeof _initSecrets_ === 'function') _initSecrets_(); } catch (eS) {}   // paramètres sensibles posés une fois (fichier hors dépôt)
  try { _rcAuto_(); } catch (eR) { console.error('audit recruteurs : ' + eR); }   // audit espace recruteurs, chaque nuit (recruteurs.gs)
  try { if (typeof _frisesAuto_ === 'function') _frisesAuto_(); } catch (eF) { console.error('rappels frises : ' + eF); }   // rappels d'échéances des frises (vs_extras.gs)
  try { _relCandAuto_(); } catch (eC) { console.error('relance candidatures : ' + eC); }   // relance quotidienne 9 h, jours ouvrés (relances.gs)
  try { if (typeof _ofAuto_ === 'function') _ofAuto_(t0); } catch (eO) { console.error('offres (lieux / partages) : ' + eO); }   // offres_lieux.gs
  try { if (typeof _fpAuto_ === 'function') _fpAuto_(t0); } catch (eP) { console.error('fiches de positionnement : ' + eP); }   // positionnement.gs
  _ADR_LIMITE_ = t0 + 320000;
  var etat = _adrEtat_();
  var h = Number(Utilities.formatDate(new Date(), "Europe/Paris", "H"));
  if ((!etat || !etat.enCours) && h === 3 && (!etat || Date.now() - (etat.lance || 0) > 20 * 3600 * 1000)) {
    _adrAuditDemarrer_("Recalage nocturne des adresses", "Administrateur", true);   // chaque nuit
    etat = _adrEtat_();
  }
  if (!etat || !etat.enCours) return;
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;   // un autre lot tourne encore : prochain passage
  try {
    var ss = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");
    while (etat.enCours && Date.now() - t0 < 150000) {   // plusieurs lots par passage, sous la limite de 6 min
      var R;
      try {
        R = _adrHarmoniser_(ss, false, etat.silencieux, { debut: etat.debut, taille: ADR_LOT, csvKo: !!etat.csvKo });
        if (_ADR_CSV_KO_) etat.csvKo = true;
      } catch (e) {
        etat.diag.push("lot " + etat.debut + " : " + e);
        etat.erreurs = (etat.erreurs || 0) + 1;
        if (etat.erreurs > 3) { etat.debut += ADR_LOT; etat.erreurs = 0; }   // lot ignoré après 3 échecs
        _adrEtatSauver_(etat);
        break;
      }
      etat.total = R.total;
      Object.keys(etat.c).forEach(function (k) { etat.c[k] += R[k] || 0; });
      R.exemples.forEach(function (e) { if (etat.ex.length < 25) etat.ex.push(e.nom + " : « " + e.avant + " » → « " + e.apres + " » [" + e.statut + "]"); });
      R.recalages.forEach(function (e) { if (etat.rc.length < 25) etat.rc.push(e.nom + " : « " + e.avant + " » → « " + e.apres + " » [" + e.source + "]"); });
      (R.diag || []).forEach(function (d) { if (etat.diag.length < 10) etat.diag.push(d); });
      _adrRapportLot_(ss, etat, R, etat.debut === 0);
      etat.debut += ADR_LOT;
      etat.erreurs = 0;
      if (R.suite) _adrEtatSauver_(etat); else _adrAuditTerminer_(ss, etat);
      _invalidatePartnersCache_();
    }
  } finally { lock.releaseLock(); }
}
function _adrResume_(c, total) {
  return total + " entreprises — " + c.harmonisees + " harmonisées (écart léger), " + c.recalees + " recalées automatiquement ("
    + c.recalSirene + " via SIRENE, " + c.recalBan + " via BAN, " + c.recalCommune + " à la commune), " + c.dejaOk + " déjà conformes, ALERTES : "
    + c.aVerifier + " à vérifier + " + c.multiples + " à adresses multiples, " + c.sansAdresse + " sans adresse, " + c.lignesModifiees + " ligne(s) modifiée(s)";
}
function _adrAuditTerminer_(ss, etat) {
  etat.enCours = false; etat.fini = Date.now();
  _adrEtatSauver_(etat);
  var resume = _adrResume_(etat.c, etat.total) + ". Détail : onglet « Audit adresses »";
  try {
    var ws = ss.getSheetByName("Audit adresses");
    ws.getRange(1, 1).setValue("Audit des adresses entreprises — terminé le " + Utilities.formatDate(new Date(), "Europe/Paris", "dd/MM/yyyy HH:mm"));
    ws.getRange(2, 1).setValue(resume);
  } catch (e) {}
  if (!etat.silencieux || etat.c.lignesModifiees)
    logAction(etat.auteur || "Système", etat.role || "", "Modification", "Entreprises — normalisation des adresses", "PARTENARIAT",
      resume + (etat.rc.length ? " | Recalages : " + etat.rc.join(" ; ") : "") + (etat.ex.length ? " | Exemples : " + etat.ex.join(" ; ") : ""));
}
// Onglet « Audit adresses » : (re)créé au 1er lot, puis lignes ajoutées à chaque lot
function _adrRapportLot_(ss, etat, R, premier) {
  var ws = ss.getSheetByName("Audit adresses") || ss.insertSheet("Audit adresses");
  if (premier) {
    ws.clear();
    ws.getRange(1, 1, 5, 5).setValues([["Audit des adresses entreprises — en cours (lancé le " + Utilities.formatDate(new Date(etat.lance), "Europe/Paris", "dd/MM/yyyy HH:mm") + ")", "", "", "", ""],
      ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""],
      ["Type", "ID entreprise", "Entreprise", "Adresse avant", "Adresse après / motif de l'alerte"]]);
    ws.getRange(1, 1).setFontWeight("bold").setFontSize(13);
    ws.getRange(5, 1, 1, 5).setFontWeight("bold").setBackground("#1A4E8A").setFontColor("#ffffff");
    ws.setFrozenRows(5);
    ws.setColumnWidths(1, 1, 190); ws.setColumnWidths(2, 1, 120); ws.setColumnWidths(3, 1, 240); ws.setColumnWidths(4, 2, 380);
  }
  ws.getRange(2, 1).setValue("Progression : " + Math.min(etat.debut + ADR_LOT, R.total) + " / " + R.total + " entreprises — " + _adrResume_(etat.c, R.total));
  ws.getRange(3, 1).setValue(etat.diag.length ? "Incidents techniques : " + etat.diag.join(" | ") : "");
  var rows = [];
  R.alertes.forEach(function (a) { rows.push(["⚠️ " + a.statut, a.id, a.nom, a.adresse, a.raison]); });
  R.recalages.forEach(function (e) { rows.push(["🔁 Recalée (" + e.source + ")", e.id, e.nom, e.avant, e.apres + (e.raison ? " — était : " + e.raison : "")]); });
  R.exemples.forEach(function (e) { rows.push(["✅ Harmonisée (" + e.statut + ")", e.id, e.nom, e.avant, e.apres]); });
  if (rows.length) ws.getRange(Math.max(ws.getLastRow(), 5) + 1, 1, rows.length, 5).setValues(rows);
}

// Audit automatique UNE fois après mise en ligne : le premier appel au serveur programme
// le lancement (voir doPost), puis chaque nuit à 3 h.
function _adrAuditAutoProgrammer_() {
  try {
    var props = PropertiesService.getScriptProperties();
    if (props.getProperty("ADR_AUDIT_V6")) return;
    props.setProperty("ADR_AUDIT_V6", "lancé " + new Date().toISOString());
    _adrAuditDemarrer_("Audit automatique des adresses", "Administrateur", false);
  } catch (e) { console.error("audit adresses : " + e); }
}
function auditAdressesAuto() { _adrAuditDemarrer_("Audit automatique des adresses", "Administrateur", false); }   // anciens déclencheurs
function auditAdressesLot() { auditAdressesTick(); }
function auditAdressesNocturne() { _adrAuditDemarrer_("Recalage nocturne des adresses", "Administrateur", true); }

// ---------------------------------------------------------------------------
// CARTE DES ENTREPRISES PARTENAIRES (espace étudiant, dossier.html — demande 2026-11-09)
// Données volontairement réduites : nom de l'ENTREPRISE (jamais les contacts), adresse,
// position et NOMBRE d'alternants HECG en cours (jamais leurs noms).
function handleGetCarteEntreprises(p, ss) {
  try {
    logAction(p.idAuteur || p.nomAuteur || "Inconnu", p.roleAuteur || "", "Consultation", "Dossier étudiant — carte des entreprises",
      String(p.targetId || p.id_etu || ""), "Ouverture de la carte des entreprises partenaires");
  } catch (eL) {}
  return getCachedJsonResponse_("carte_entreprises_v1", 600, function () {
    var dP = getSheetSafe(ss, "Partenariat").getDataRange().getValues();
    var dE = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues();
    var C = _adrColsLecture_(dP[0]);
    var ents = {}, ordre = [];
    for (var i = 1; i < dP.length; i++) {
      var id = String(dP[i][1] || "").trim().toUpperCase();
      if (!id) continue;
      if (!ents[id]) { ents[id] = { nom: String(dP[i][2] || "").trim(), lignes: [], alt: 0 }; ordre.push(id); }
      if (!ents[id].nom) ents[id].nom = String(dP[i][2] || "").trim();
      ents[id].lignes.push(_adrDeLigne_(dP[i], C));
    }
    for (var j = 1; j < dE.length; j++) {   // même règle que la base entreprises : statut alternant / apprenti
      var idE = String(dE[j][11] || "").trim().toUpperCase();
      var st = String(dE[j][13] || "").toLowerCase();
      if (ents[idE] && (st.indexOf("altern") !== -1 || st.indexOf("apprenti") !== -1)) ents[idE].alt++;
    }
    var out = [];
    ordre.forEach(function (id) {
      var e = ents[id], a = _adrReference_(e.lignes);
      if (!e.nom || (!a.adresse && !a.cp && !a.ville)) return;
      out.push({ nom: e.nom, adresse: a.adresse || "", complement: a.complement || "", cp: a.cp || "", ville: a.ville || "",
        lat: a.lat || "", lon: a.lon || "", precise: !!ADR_STATUTS_OK[a.statut], alternants: e.alt });
    });
    out.sort(function (x, y) { return x.nom.localeCompare(y.nom, "fr"); });
    return createJsonResponse({ success: true, data: out });
  });
}

// ---------------------------------------------------------------------------
// RECALAGE SIRENE DEPUIS LE NAVIGATEUR (bouton 🏛️ de partenaires.html) : l'annuaire
// recherche-entreprises refuse les serveurs Google (HTTP 502), le navigateur l'interroge donc
// et transmet les résultats bruts ; les contrôles (nom, rue / commune, département) restent ici.
function handleRecalerAdressesSirene(p, ss) {
  var isAdmin = String(p.roleAuteur || "") === "Administrateur" || PARAM_MASTER_EMAILS.indexOf(String(p.emailAuteur || "").toLowerCase()) !== -1;
  if (!isAdmin) return createJsonResponse({ success: false, message: "Réservé aux administrateurs." });
  var sirene = {}, ids = [];
  try { sirene = typeof p.resultats === "string" ? JSON.parse(p.resultats) : (p.resultats || {}); } catch (e) { sirene = {}; }
  Object.keys(sirene).forEach(function (k) { var K = String(k).trim().toUpperCase(); ids.push(K); if (K !== k) sirene[K] = sirene[k]; });
  if (!ids.length) return createJsonResponse({ success: false, message: "Aucune entreprise à recaler." });
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return createJsonResponse({ success: false, message: "Un audit des adresses est en cours, réessayez dans quelques minutes." });
  try {
    var R = _adrHarmoniser_(ss, false, false, { ids: ids, sirene: sirene });
    var resume = ids.length + " entreprises à recaler soumises à SIRENE — " + R.recalSirene + " recalées via SIRENE (adresse officielle INSEE + SIRET), "
      + (R.recalees - R.recalSirene) + " via la Base Adresse Nationale, " + R.harmonisees + " harmonisées, ALERTES restantes : " + R.aVerifier + " à vérifier + " + R.multiples + " adresses multiples";
    var rc = R.recalages.slice(0, 30).map(function (e) { return e.nom + " : « " + e.avant + " » → « " + e.apres + " » [" + e.source + "]"; });
    logAction(p.nomAuteur || p.idAuteur || "Administrateur", p.roleAuteur || "", "Modification", "Entreprises — recalage SIRENE des adresses", "PARTENARIAT",
      resume + (rc.length ? " | " + rc.join(" ; ") : ""));
    _invalidatePartnersCache_();
    return createJsonResponse({ success: true, resume: resume, rapport: { soumises: ids.length, sirene: R.recalSirene, autres: R.recalees - R.recalSirene,
      harmonisees: R.harmonisees, aVerifier: R.aVerifier, multiples: R.multiples, recalages: R.recalages.slice(0, 200) } });
  } catch (e) {
    return createJsonResponse({ success: false, message: e.toString() });
  } finally { lock.releaseLock(); }
}
