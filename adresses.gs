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
    var tail = toks.slice(idxRue + 1).join(" ").trim();
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
  if (!(cpOk || (villeOk && depOk))) return null;
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
    var part = items.slice(start, start + 400);
    var csv = "id,q,t\n" + part.map(function (it, i) {
      return i + ',"' + String(it.q || "").replace(/"/g, '""').replace(/[\r\n]+/g, " ") + '",' + (typeCommune ? "municipality" : "");
    }).join("\n");
    var payload = { data: Utilities.newBlob(csv, "text/csv", "adresses.csv"), columns: "q" };
    if (typeCommune) payload.type = "t";
    var res = null;
    for (var essai = 0; essai < 3 && !res; essai++) {
      try {
        var r = UrlFetchApp.fetch(ADR_BAN_URL + "/csv", { method: "post", payload: payload, muteHttpExceptions: true });
        if (r.getResponseCode() === 200) res = r.getContentText("UTF-8"); else Utilities.sleep(1500);
      } catch (e) { Utilities.sleep(1500); }
    }
    var rows = res ? Utilities.parseCsv(res) : [];
    var h = rows.length ? rows[0] : [];
    function ix(n) { return h.indexOf(n); }
    var parsed = {};
    for (var k = 1; k < rows.length; k++) {
      var row = rows[k];
      parsed[row[0]] = row[ix("result_type")] ? {
        type: row[ix("result_type")], score: Number(row[ix("result_score")]) || 0,
        housenumber: row[ix("result_housenumber")], street: row[ix("result_street")], name: row[ix("result_name")],
        postcode: row[ix("result_postcode")], city: row[ix("result_city")],
        lat: Number(row[ix("latitude")]), lon: Number(row[ix("longitude")])
      } : null;
    }
    part.forEach(function (it, i) { out.push(parsed[String(i)] || null); });
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
function handleHarmoniserAdresses(p, ss) {
  var isAdmin = String(p.roleAuteur || "") === "Administrateur" || PARAM_MASTER_EMAILS.indexOf(String(p.emailAuteur || "").toLowerCase()) !== -1;
  if (!isAdmin) return createJsonResponse({ success: false, message: "Réservé aux administrateurs." });
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return createJsonResponse({ success: false, message: "Une normalisation est déjà en cours." });
  try {
    var rapport = _adrHarmoniser_(ss, !!p.forcer);
    var resume = rapport.entreprises + " entreprises — " + rapport.ban + " adresse(s) au numéro, " + rapport.voie + " à la voie, "
      + rapport.commune + " à la commune, " + rapport.aVerifier + " à vérifier, " + rapport.sansAdresse + " sans adresse, "
      + rapport.unifiees + " entreprise(s) à adresses multiples unifiées, " + rapport.lignesModifiees + " ligne(s) modifiée(s)";
    var ex = rapport.exemples.slice(0, 25).map(function (e) { return e.nom + " : « " + e.avant + " » → « " + e.apres + " » [" + e.statut + "]"; });
    logAction(p.nomAuteur || p.idAuteur || "Système", p.roleAuteur || "", "Modification", "Entreprises — normalisation des adresses",
      "PARTENARIAT", resume + (ex.length ? " | Exemples : " + ex.join(" ; ") : ""));
    _invalidatePartnersCache_();
    return createJsonResponse({ success: true, rapport: rapport, resume: resume });
  } catch (e) {
    return createJsonResponse({ success: false, message: e.toString() });
  } finally { lock.releaseLock(); }
}

function harmoniserAdressesEntreprises() {   // depuis l'éditeur Apps Script
  var ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");
  var out = handleHarmoniserAdresses({ roleAuteur: "Administrateur", nomAuteur: "Éditeur Apps Script" }, ss);
  Logger.log(out.getContent());
}

function _adrHarmoniser_(ss, forcer) {
  var ws = getSheetSafe(ss, "PARTENARIAT");
  var C = _adrCols_(ws);
  var data = ws.getDataRange().getValues();
  var nCols = data[0].length;
  var groupes = {}, ordre = [];
  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][1] || "").trim().toUpperCase();
    if (!id) continue;
    if (!groupes[id]) { groupes[id] = { id: id, nom: String(data[i][2] || "").trim(), rows: [], lignes: [] }; ordre.push(id); }
    groupes[id].rows.push(i);
    groupes[id].lignes.push(_adrDeLigne_(data[i], C));
  }
  var rapport = { entreprises: ordre.length, ban: 0, voie: 0, commune: 0, aVerifier: 0, sansAdresse: 0, dejaOk: 0, unifiees: 0, lignesModifiees: 0, exemples: [], aVerifierListe: [] };

  // 1. Adresse de référence + besoin de géocodage
  var aGeocoder = [];
  ordre.forEach(function (id) {
    var g = groupes[id];
    var cles = {};
    g.lignes.forEach(function (a) { if (a.adresse || a.cp || a.ville) cles[_adrNorm_(a.adresse + "|" + a.cp + "|" + a.ville)] = 1; });
    g.multiples = Object.keys(cles).length > 1;
    g.ref = _adrReference_(g.lignes);
    g.avant = _adrTexte_(g.ref);
    g.final = null;
    if (!g.ref.adresse && !g.ref.cp && !g.ref.ville) { g.final = g.ref; return; }
    var dejaTraitee = g.ref.statut && (ADR_STATUTS_OK[g.ref.statut] || (!forcer && (g.ref.statut === "Saisie libre")));
    if (dejaTraitee) { g.final = g.ref; return; }
    g.saisie = { adresse: _adrNettoyer_(g.ref.adresse, g.ref.cp), cp: /^\d{5}$/.test(g.ref.cp) ? g.ref.cp : (/^\d{4}$/.test(g.ref.cp) ? "0" + g.ref.cp : ""), ville: g.ref.ville };
    aGeocoder.push(g);
  });

  // 2. Passe 1 : adresses complètes
  var avecRue = aGeocoder.filter(function (g) { return g.saisie.adresse; });
  var res1 = _adrGeocoderLot_(avecRue.map(function (g) { return { q: [g.saisie.adresse, g.saisie.cp, g.saisie.ville].filter(Boolean).join(" ") }; }), false);
  avecRue.forEach(function (g, k) {
    var r = _adrEvaluer_(res1[k], g.saisie);
    if (r) g.final = { adresse: r.adresse, cp: r.cp, ville: r.ville, complement: g.ref.complement || r.complement, lat: r.lat, lon: r.lon, statut: r.statut, siret: g.ref.siret };
  });
  // 3. Passe 2 : commune (adresse introuvable ou absente)
  var sansRes = aGeocoder.filter(function (g) { return !g.final && (g.saisie.ville || g.saisie.cp); });
  var res2 = _adrGeocoderLot_(sansRes.map(function (g) { return { q: [g.saisie.ville, g.saisie.cp ? g.saisie.cp.slice(0, 2) : ""].filter(Boolean).join(" ") || g.saisie.cp }; }), true);
  sansRes.forEach(function (g, k) {
    var r = _adrEvaluerCommune_(res2[k], g.saisie);
    g.final = {
      adresse: g.saisie.adresse, cp: r ? r.cp : g.saisie.cp || g.ref.cp, ville: r ? r.ville : g.ref.ville, complement: g.ref.complement,
      lat: r ? r.lat : "", lon: r ? r.lon : "", statut: g.saisie.adresse ? "À vérifier" : (r ? "Commune" : "À vérifier"), siret: g.ref.siret
    };
  });
  aGeocoder.forEach(function (g) {
    if (!g.final) g.final = { adresse: g.saisie.adresse, cp: g.ref.cp, ville: g.ref.ville, complement: g.ref.complement, lat: "", lon: "", statut: "À vérifier", siret: g.ref.siret };
  });

  // 4. Application en mémoire sur toutes les lignes de chaque entreprise
  var colsMin = Math.min(C.complement, C.lat, C.lon, C.statut, C.siret, C.origine);
  var colsMax = Math.max(C.complement, C.lat, C.lon, C.statut, C.siret, C.origine);
  var bloc1 = [], bloc2 = [];   // K:M et colonnes étendues
  for (var r0 = 1; r0 < data.length; r0++) {
    while (data[r0].length < Math.max(nCols, colsMax + 1)) data[r0].push("");
  }
  ordre.forEach(function (id) {
    var g = groupes[id], a = g.final;
    var st = a.statut;
    if (!a.adresse && !a.cp && !a.ville) rapport.sansAdresse++;
    else if (st === "BAN" || st === "SIRENE") rapport.ban++;
    else if (st === "BAN voie") rapport.voie++;
    else if (st === "Commune") rapport.commune++;
    else { rapport.aVerifier++; if (rapport.aVerifierListe.length < 400) rapport.aVerifierListe.push({ id: id, nom: g.nom, adresse: _adrTexte_(a), statut: st }); }
    var modifie = false;
    g.rows.forEach(function (ri) {
      var row = data[ri];
      var avantLigne = [row[10], row[11], row[12], row[C.complement], row[C.lat], row[C.lon], row[C.statut], String(row[C.siret]).replace(/^'/, "")].map(String).join("|");
      var origine = row[C.origine];
      if (!origine && g.avant && g.avant !== _adrTexte_(a)) origine = [row[10], row[11], row[12]].map(function (x) { return String(x || "").trim(); }).filter(Boolean).join(", ");
      _adrDansLigne_(row, C, a);
      row[C.origine] = origine || "";
      var apresLigne = [row[10], row[11], row[12], row[C.complement], row[C.lat], row[C.lon], row[C.statut], a.siret || ""].map(String).join("|");
      if (avantLigne !== apresLigne) { rapport.lignesModifiees++; modifie = true; }
    });
    if (g.multiples) rapport.unifiees++;
    if (modifie && g.avant !== _adrTexte_(a) && rapport.exemples.length < 300) rapport.exemples.push({ id: id, nom: g.nom, avant: g.avant || "—", apres: _adrTexte_(a) || "—", statut: st || "—" });
  });

  // 5. Écriture groupée (2 plages)
  for (var r1 = 1; r1 < data.length; r1++) {
    bloc1.push([data[r1][10], data[r1][11], data[r1][12]]);
    bloc2.push(data[r1].slice(colsMin, colsMax + 1));
  }
  if (bloc1.length) {
    ws.getRange(2, 11, bloc1.length, 3).setValues(bloc1);
    ws.getRange(2, colsMin + 1, bloc2.length, colsMax - colsMin + 1).setValues(bloc2);
  }
  return rapport;
}
