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
    var resume = rapport.entreprises + " entreprises — " + rapport.harmonisees + " harmonisées (" + rapport.ban + " au numéro, " + rapport.voie + " à la voie, "
      + rapport.commune + " à la commune), " + rapport.dejaOk + " déjà conformes, ALERTES : " + rapport.aVerifier + " à vérifier + " + rapport.multiples
      + " à adresses multiples (non modifiées), " + rapport.sansAdresse + " sans adresse, " + rapport.lignesModifiees + " ligne(s) modifiée(s). Détail : onglet « Audit adresses »";
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

function _adrHarmoniser_(ss, forcer) {
  var ws = getSheetSafe(ss, "PARTENARIAT");
  var C = _adrCols_(ws);
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
  _adrDecider_(groupes, ordre, _adrGeocoderLot_, forcer);
  var rapport = { entreprises: ordre.length, dejaOk: 0, harmonisees: 0, ban: 0, voie: 0, commune: 0, aVerifier: 0, multiples: 0,
    sansAdresse: 0, lignesModifiees: 0, exemples: [], alertes: [] };

  ordre.forEach(function (id) {
    var g = groupes[id], d = g.decision;
    var avantTxt = (g.cands || []).map(function (c) { return _adrTexte_(c.brut); }).join(" / ");
    var modifie = false;
    g.rows.forEach(function (ri) {
      var row = data[ri];
      var avant = row.slice(10, 13).concat([row[C.complement], row[C.lat], row[C.lon], row[C.statut], row[C.origine]]).map(String).join("|");
      if (d.type === "harmonise" || d.type === "ok") {
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
    } else {
      if (d.statut === "Adresses multiples") rapport.multiples++; else rapport.aVerifier++;
      if (rapport.alertes.length < 500) rapport.alertes.push({ id: id, nom: g.nom, adresse: avantTxt, statut: d.statut, raison: d.raison });
    }
  });

  var bloc1 = [], bloc2 = [];
  for (var r = 1; r < data.length; r++) {
    bloc1.push([data[r][10], data[r][11], data[r][12]]);
    bloc2.push(data[r].slice(colsMin, colsMax + 1));
  }
  if (bloc1.length) {
    ws.getRange(2, 11, bloc1.length, 3).setValues(bloc1);
    ws.getRange(2, colsMin + 1, bloc2.length, colsMax - colsMin + 1).setValues(bloc2);
  }
  _adrRapportFeuille_(ss, rapport);
  return rapport;
}

// Rapport lisible dans l'onglet « Audit adresses » (recréé à chaque audit)
function _adrRapportFeuille_(ss, R) {
  try {
    var ws = ss.getSheetByName("Audit adresses") || ss.insertSheet("Audit adresses");
    ws.clear();
    var rows = [["Audit des adresses entreprises — " + Utilities.formatDate(new Date(), "Europe/Paris", "dd/MM/yyyy HH:mm"), "", "", "", ""],
      [R.entreprises + " entreprises · " + R.harmonisees + " harmonisées · " + R.dejaOk + " déjà conformes · " + R.aVerifier + " à vérifier · " + R.multiples + " à adresses multiples · " + R.sansAdresse + " sans adresse", "", "", "", ""],
      ["", "", "", "", ""], ["Type", "ID entreprise", "Entreprise", "Adresse avant", "Adresse après / motif de l'alerte"]];
    R.alertes.forEach(function (a) { rows.push(["⚠️ " + a.statut, a.id, a.nom, a.adresse, a.raison]); });
    R.exemples.forEach(function (e) { rows.push(["✅ Harmonisée (" + e.statut + ")", e.id, e.nom, e.avant, e.apres]); });
    ws.getRange(1, 1, rows.length, 5).setValues(rows);
    ws.getRange(1, 1).setFontWeight("bold").setFontSize(13);
    ws.getRange(4, 1, 1, 5).setFontWeight("bold").setBackground("#1A4E8A").setFontColor("#ffffff");
    ws.setFrozenRows(4);
    ws.setColumnWidths(1, 1, 190); ws.setColumnWidths(2, 1, 120); ws.setColumnWidths(3, 1, 240); ws.setColumnWidths(4, 2, 380);
  } catch (e) { console.error("rapport audit : " + e); }
}

// Audit automatique UNE fois après mise en ligne : le premier appel au serveur programme
// l'exécution par un déclencheur ponctuel (hors requête utilisateur) — voir doPost.
function _adrAuditAutoProgrammer_() {
  try {
    var props = PropertiesService.getScriptProperties();
    if (props.getProperty("ADR_AUDIT_V1")) return;
    props.setProperty("ADR_AUDIT_V1", "programmé " + new Date().toISOString());
    ScriptApp.newTrigger("auditAdressesAuto").timeBased().after(60 * 1000).create();
  } catch (e) { console.error("audit adresses : " + e); }
}
function auditAdressesAuto() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === "auditAdressesAuto") ScriptApp.deleteTrigger(t); });
  var ss = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");
  var out = handleHarmoniserAdresses({ roleAuteur: "Administrateur", nomAuteur: "Audit automatique des adresses" }, ss);
  PropertiesService.getScriptProperties().setProperty("ADR_AUDIT_V1", "fait " + new Date().toISOString() + " — " + String(out.getContent()).slice(0, 300));
}
