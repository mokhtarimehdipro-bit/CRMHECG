// ==========================================
// CUBE — Analyse croisée dynamique (Vie scolaire → Tableaux de suivis)
// ==========================================
// Construit un jeu de données « à plat » à partir d'une ou plusieurs sources du CRM
// (réutilise _tabSuiviLoadCrm_ de tableaux_suivi.gs), avec :
//   - union de plusieurs sources (colonne « Source ») ;
//   - enrichissement par la fiche entreprise (étudiants, contacts, offres) ;
//   - dimensions dérivées : année / mois / trimestre des dates, département (CP).
// Le front (cube.js) calcule le tableau croisé dans le navigateur.
// L'export crée dans un tableau de suivi Google Sheets :
//   - un onglet de données masqué « CUBE·data … » ;
//   - un onglet avec un VRAI tableau croisé dynamique Google Sheets (même configuration),
//     reconstruit à chaque synchronisation (bouton 🔄 + trigger horaire existant).

var CUBE_SOURCES = {
  etudiants:   "Étudiants",
  entreprises: "Entreprises",
  contacts:    "Contacts entreprises",
  offres:      "Offres",
  campagnes:   "Campagnes",
  activite:    "Activité CRM"
};
// Champs non pertinents pour l'analyse (coordonnées personnelles, liens) : exclus du cube
var CUBE_EXCLUDE = { "Mail": 1, "Téléphone": 1, "Mail HECG": 1, "Téléphone 2": 1, "Mail 2": 1, "Lien": 1, "Adresse": 1, "Alternants actifs": 1 };
var CUBE_NUMERIC = { "Nb contacts": 1, "Nb alternants": 1, "Nb partages": 1, "Entreprise › Nb contacts": 1, "Entreprise › Nb alternants": 1 };
var CUBE_JOIN_FIELDS = ["Ville", "CP", "Nb contacts", "Nb alternants", "Dernier échange"];
var CUBE_MAX_ACTIVITE = 30000;

function _cubeCanSeeActivite_(p) {
  return String(p.vieScRole || "") === "Super-admin" || String(p.roleAuteurValide || "") === "Administrateur";
}

function _cubeIsDateField_(name) {
  return /^(Date|Dernier échange|Dernier envoi)/.test(String(name).replace(/^Entreprise › /, ""));
}

// "dd/MM/yyyy" → [année, "yyyy-MM", "T1 2026"]
function _cubeDateParts_(s) {
  const m = String(s || "").match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (!m) return ["", "", ""];
  const mo = ("0" + m[2]).slice(-2);
  return [m[3], m[3] + "-" + mo, "T" + Math.ceil(Number(m[2]) / 3) + " " + m[3]];
}

function _cubeLoadActivite_(ss, cache) {
  const tz = Session.getScriptTimeZone();
  const d = cache.HISTORIQUE || (cache.HISTORIQUE = getSheetSafe(ss, "HISTORIQUE").getDataRange().getValues());
  const jours = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi"];
  const out = [];
  const start = Math.max(1, d.length - CUBE_MAX_ACTIVITE);
  for (let i = start; i < d.length; i++) {
    const r = d[i];
    const dt = r[1] instanceof Date ? r[1] : new Date(r[1]);
    if (isNaN(dt.getTime())) continue;
    out.push({ id: "H" + i, values: [
      Utilities.formatDate(dt, tz, "dd/MM/yyyy"), String(r[2] || ""), String(r[3] || ""), String(r[4] || ""),
      String(r[5] || ""), jours[dt.getDay()], Utilities.formatDate(dt, tz, "HH") + "h"
    ]});
  }
  return out;
}
var CUBE_ACTIVITE_FIELDS = ["Date", "Auteur", "Rôle", "Type d'action", "Catégorie", "Jour", "Heure"];

/**
 * Construit le jeu de données du cube.
 * @return {{fields:Array<{n:string,t:string}>, rows:Array<Array>}}
 */
function _cubeBuild_(ss, sources, join, cache, allowActivite) {
  cache = cache || {};
  sources = (sources || []).filter(function(s) { return CUBE_SOURCES[s] && (s !== "activite" || allowActivite); });
  if (!sources.length) throw new Error("Choisissez au moins une source de données.");
  const multi = sources.length > 1;

  // Référentiel entreprises (pour les jointures)
  let entById = null, entByName = null;
  const entFields = TS_FIELDS.entreprises;
  const loadEnt = function() {
    if (entById) return;
    entById = {}; entByName = {};
    _tabSuiviLoadCrm_(ss, "entreprises", {}, cache).forEach(function(r) {
      entById[String(r.id).toUpperCase()] = r.values;
      const nm = String(r.values[0] || "").trim().toLowerCase();
      if (nm && !entByName[nm]) entByName[nm] = r.values;
    });
  };
  const joinVals = function(vals) {
    return CUBE_JOIN_FIELDS.map(function(f) { const k = entFields.indexOf(f); return vals && k !== -1 ? vals[k] : ""; });
  };

  // 1. Jeux par source : { names:[], rows:[[...]] }
  const blocks = sources.map(function(src) {
    let baseNames, items;
    if (src === "activite") { baseNames = CUBE_ACTIVITE_FIELDS; items = _cubeLoadActivite_(ss, cache); }
    else { baseNames = TS_FIELDS[src]; items = _tabSuiviLoadCrm_(ss, src, {}, cache); }
    const keep = [];
    baseNames.forEach(function(n, k) { if (!CUBE_EXCLUDE[n]) keep.push(k); });
    const names = keep.map(function(k) { return baseNames[k]; });

    // Jointure entreprise
    let joinKeyFn = null;
    if (join && (src === "etudiants" || src === "contacts" || src === "offres")) {
      loadEnt();
      if (src === "etudiants") {
        const dE = cache.ETUDIANTS || (cache.ETUDIANTS = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues());
        const idEntByEtu = {};
        for (let j = 1; j < dE.length; j++) idEntByEtu[String(dE[j][0] || "").trim()] = String(dE[j][11] || "").trim().toUpperCase();
        joinKeyFn = function(it) { return entById[idEntByEtu[String(it.id).trim()]] || null; };
      } else if (src === "contacts") {
        const dP = cache.PARTENARIAT || (cache.PARTENARIAT = getSheetSafe(ss, "PARTENARIAT").getDataRange().getValues());
        const idEntByContact = {};
        for (let j = 1; j < dP.length; j++) idEntByContact[String(dP[j][0] || "").trim()] = String(dP[j][1] || "").trim().toUpperCase();
        joinKeyFn = function(it) { return entById[idEntByContact[String(it.id).trim()]] || null; };
      } else {
        const kEnt = TS_FIELDS.offres.indexOf("Entreprise");
        joinKeyFn = function(it) { return entByName[String(it.values[kEnt] || "").trim().toLowerCase()] || null; };
      }
      CUBE_JOIN_FIELDS.forEach(function(f) { names.push("Entreprise › " + f); });
    }

    const rows = items.map(function(it) {
      const line = keep.map(function(k) { return it.values[k]; });
      if (joinKeyFn) {
        const ev = joinKeyFn(it);
        joinVals(ev).forEach(function(v) { line.push(v); });
      }
      return line;
    });
    return { src: src, names: names, rows: rows };
  });

  // 2. Union des champs + dimensions dérivées
  const fieldOrder = [];
  const seen = {};
  const addField = function(n) { if (!seen[n]) { seen[n] = true; fieldOrder.push(n); } };
  if (multi) addField("Source");
  blocks.forEach(function(b) { b.names.forEach(addField); });
  const derived = [];
  fieldOrder.slice().forEach(function(n) {
    if (_cubeIsDateField_(n)) { derived.push({ from: n, kind: "y", n: n + " · Année" }, { from: n, kind: "q", n: n + " · Trimestre" }, { from: n, kind: "m", n: n + " · Mois" }); }
    if (n === "CP" || n === "Entreprise › CP") derived.push({ from: n, kind: "dep", n: n === "CP" ? "Département" : "Entreprise › Département" });
  });
  derived.forEach(function(d) { addField(d.n); });
  const idx = {}; fieldOrder.forEach(function(n, i) { idx[n] = i; });

  const rows = [];
  blocks.forEach(function(b) {
    const map = b.names.map(function(n) { return idx[n]; });
    b.rows.forEach(function(line) {
      const out = new Array(fieldOrder.length).fill("");
      if (multi) out[0] = CUBE_SOURCES[b.src];
      for (let k = 0; k < line.length; k++) {
        let v = line[k];
        const name = b.names[k];
        if (CUBE_NUMERIC[name]) { const num = Number(v); v = isNaN(num) ? "" : num; }
        out[map[k]] = v === null || v === undefined ? "" : v;
      }
      derived.forEach(function(d) {
        const src = out[idx[d.from]];
        if (d.kind === "dep") { const cp = String(src || "").replace(/\s/g, ""); out[idx[d.n]] = /^\d{5}$/.test(cp) ? (cp.indexOf("97") === 0 ? cp.slice(0, 3) : cp.slice(0, 2)) : ""; }
        else { const parts = _cubeDateParts_(src); out[idx[d.n]] = d.kind === "y" ? parts[0] : (d.kind === "m" ? parts[1] : parts[2]); }
      });
      rows.push(out);
    });
  });

  const fields = fieldOrder.map(function(n) {
    return { n: n, t: CUBE_NUMERIC[n] ? "num" : (_cubeIsDateField_(n) && n.indexOf(" · ") === -1 ? "date" : "txt") };
  });
  return { fields: fields, rows: rows };
}

// ── Lecture (front) ───────────────────────────────────────────────
function handleGetCubeData(p, ss) {
  if (!_tabSuiviCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  try {
    const data = _cubeBuild_(ss, p.sources, p.join === true, {}, _cubeCanSeeActivite_(p));
    return createJsonResponse({ success: true, fields: data.fields, rows: data.rows, activiteAutorisee: _cubeCanSeeActivite_(p) });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// ── Export → vrai tableau croisé dynamique Google Sheets ─────────
var CUBE_AGG_SHEETS = { count: "COUNTA", countd: "COUNTUNIQUE", sum: "SUM", avg: "AVERAGE", min: "MIN", max: "MAX" };
var CUBE_AGG_LABEL = { count: "Nombre", countd: "Nombre distinct", sum: "Somme", avg: "Moyenne", min: "Min", max: "Max" };

function _cubeResume_(cfg) {
  const s = (cfg.sources || []).map(function(x) { return CUBE_SOURCES[x] || x; }).join(" + ");
  const parts = [];
  if ((cfg.rows || []).length) parts.push("Lignes : " + cfg.rows.join(", "));
  if ((cfg.cols || []).length) parts.push("Colonnes : " + cfg.cols.join(", "));
  if ((cfg.vals || []).length) parts.push("Valeurs : " + cfg.vals.map(function(v) { return (CUBE_AGG_LABEL[v.agg] || v.agg) + (v.f ? " de " + v.f : ""); }).join(", "));
  return "🧊 Cube — " + s + (cfg.join ? " (+ fiche entreprise)" : "") + (parts.length ? " — " + parts.join(" · ") : "");
}

function _cubeWriteSheets_(ss, file, entry, cache, allowActivite) {
  const cfg = entry.config || {};
  const data = _cubeBuild_(ss, cfg.sources, cfg.join === true, cache, allowActivite);
  // Colonne technique « Compteur » = 1 par ligne : le « Nombre » du pivot est une SOMME de cette
  // colonne (COUNTA sur une colonne métier se tromperait dès qu'une cellule est vide).
  const names = data.fields.map(function(f) { return f.n; }).concat(["Compteur"]);
  data.rows.forEach(function(r) { r.push(1); });

  // Onglet de données (masqué)
  let ds = file.getSheetByName(entry.dataTab);
  if (!ds) ds = file.insertSheet(entry.dataTab);
  ds.clear();
  const nbRows = data.rows.length + 1;
  if (ds.getMaxRows() < nbRows) ds.insertRowsAfter(ds.getMaxRows(), nbRows - ds.getMaxRows());
  if (ds.getMaxColumns() < names.length) ds.insertColumnsAfter(ds.getMaxColumns(), names.length - ds.getMaxColumns());
  ds.getRange(1, 1, 1, names.length).setValues([names]).setFontWeight("bold").setBackground("#1A4E8A").setFontColor("#ffffff");
  if (data.rows.length) ds.getRange(2, 1, data.rows.length, names.length).setValues(data.rows.map(function(r) { return r.map(function(v) { return v === "" ? "" : v; }); }));
  ds.setFrozenRows(1);
  try { ds.hideSheet(); } catch (eH) {}

  // Onglet tableau croisé
  let ps = file.getSheetByName(entry.tab);
  if (!ps) ps = file.insertSheet(entry.tab);
  ps.getPivotTables().forEach(function(pt) { pt.remove(); });
  ps.clear();
  const tz = Session.getScriptTimeZone();
  ps.getRange("A1").setValue("🧊 " + (cfg.nom || entry.tab)).setFontWeight("bold").setFontSize(14).setFontColor("#1A4E8A");
  ps.getRange("A2").setValue(_cubeResume_(cfg).replace(/^🧊 Cube — /, "") + "  ·  " + data.rows.length + " lignes  ·  mis à jour le " + Utilities.formatDate(new Date(), tz, "dd/MM/yyyy à HH:mm"))
    .setFontColor("#64748b").setFontSize(9);
  ps.getRange("A3").setValue("Tableau croisé dynamique Google Sheets : modifiable via l'éditeur à droite. Il est reconstruit à chaque synchronisation CRM.")
    .setFontColor("#F26522").setFontSize(9);

  const src = ds.getRange(1, 1, Math.max(nbRows, 2), names.length);
  const pt = ps.getRange("A5").createPivotTable(src);
  const col = function(n) { return names.indexOf(n) + 1; };
  (cfg.rows || []).forEach(function(n) { if (col(n) > 0) { const g = pt.addRowGroup(col(n)); try { g.showTotals(true); } catch (e) {} } });
  (cfg.cols || []).forEach(function(n) { if (col(n) > 0) pt.addColumnGroup(col(n)); });
  const vals = (cfg.vals && cfg.vals.length) ? cfg.vals : [{ f: "", agg: "count" }];
  vals.forEach(function(v) {
    const isCount = !v.f || v.agg === "count";
    const c = isCount ? names.length : col(v.f);
    if (c <= 0) return;
    const fn = SpreadsheetApp.PivotTableSummarizeFunction[isCount ? "SUM" : (CUBE_AGG_SHEETS[v.agg] || "COUNTA")];
    const pv = pt.addPivotValue(c, fn);
    try { pv.setDisplayName((CUBE_AGG_LABEL[v.agg] || "Nombre") + (v.f ? " · " + v.f : "")); } catch (e) {}
    try {
      if (cfg.show === "pctTotal") pv.showAs(SpreadsheetApp.PivotValueDisplayType.PERCENT_OF_GRAND_TOTAL);
      else if (cfg.show === "pctRow") pv.showAs(SpreadsheetApp.PivotValueDisplayType.PERCENT_OF_ROW_TOTAL);
      else if (cfg.show === "pctCol") pv.showAs(SpreadsheetApp.PivotValueDisplayType.PERCENT_OF_COLUMN_TOTAL);
    } catch (e) {}
  });
  const filters = cfg.filters || {};
  Object.keys(filters).forEach(function(n) {
    const allowed = filters[n];
    if (col(n) > 0 && Array.isArray(allowed) && allowed.length) {
      try { pt.addFilter(col(n), SpreadsheetApp.newFilterCriteria().setVisibleValues(allowed.map(function(v) { return v === "(vide)" ? "" : String(v); })).build()); } catch (eF) { Logger.log("Filtre cube " + n + " : " + eF); }
    }
  });
  ps.setFrozenRows(4);
  return { lignes: data.rows.length };
}

function handleExportCubeToSheet(p, ss) {
  if (!_tabSuiviCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    const reg = _tabSuiviGetRegistre_(ss);
    const f = _tabSuiviFindRow_(reg, p.id);
    if (!f) return createJsonResponse({ success: false, message: "Tableau introuvable." });
    if (!_tabSuiviIsOwnerOrAdmin_(p, f.row)) return createJsonResponse({ success: false, message: "Seul l'auteur (ou un Super-admin) peut ajouter un cube dans ce tableau." });
    const cfg = p.config || {};
    if ((cfg.sources || []).indexOf("activite") !== -1 && !_cubeCanSeeActivite_(p)) return createJsonResponse({ success: false, message: "Source « Activité CRM » réservée aux administrateurs." });

    const file = SpreadsheetApp.openById(String(f.row[TS_COL.FILE_ID]));
    let base = ("🧊 " + (String(cfg.nom || "").trim() || "Cube")).replace(/[\[\]\*\?\/\\:]/g, " ").substring(0, 80);
    let tab = base, n = 2;
    while (file.getSheetByName(tab)) tab = base + " (" + (n++) + ")";
    const entry = { type: "cube", tab: tab, dataTab: ("CUBE·data " + tab.replace(/^🧊 /, "")).substring(0, 95), config: cfg,
                    createdAt: new Date().toISOString(), resume: _cubeResume_(cfg), allowActivite: _cubeCanSeeActivite_(p) };
    const res = _cubeWriteSheets_(ss, file, entry, {}, entry.allowActivite);
    const accueil = file.getSheetByName("Accueil");
    if (accueil && file.getSheets().length > 1) { try { file.deleteSheet(accueil); } catch (eA) {} }
    try { file.setActiveSheet(file.getSheetByName(tab)); file.moveActiveSheet(1); } catch (eM) {}

    const imports = _tabSuiviParseImports_(f.row[TS_COL.IMPORTS]);
    imports.push(entry);
    reg.getRange(f.index, TS_COL.IMPORTS + 1).setValue(JSON.stringify(imports));
    reg.getRange(f.index, TS_COL.SYNCHRO + 1).setValue(new Date());
    logAction(p.nomAuteur || p.emailAuteur, p.roleAuteur || p.vieScRole || "", "Modification", "Tableau de suivi", p.id, "Cube « " + tab + " » (" + res.lignes + " lignes) ajouté à « " + f.row[TS_COL.NOM] + " »");
    return createJsonResponse({ success: true, onglet: tab, lignes: res.lignes, url: String(f.row[TS_COL.URL] || "") });
  } catch (e) {
    return createJsonResponse({ success: false, message: "Erreur export cube : " + e.toString() });
  } finally { try { lock.releaseLock(); } catch (eL) {} }
}

// Appelé par la synchronisation des tableaux de suivi pour les entrées de type « cube »
function _cubeSyncEntry_(ss, file, entry, cache) {
  if (!file.getSheetByName(entry.tab)) return { ajoutes: 0, maj: 0, absent: true };
  const r = _cubeWriteSheets_(ss, file, entry, cache, entry.allowActivite === true);
  return { ajoutes: 0, maj: r.lignes };
}
