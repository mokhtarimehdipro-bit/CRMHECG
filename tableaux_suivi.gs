// ==========================================
// TABLEAUX DE SUIVI (Vie scolaire → onglet « Tableaux de suivis Excel »)
// ==========================================
// Chaque tableau de suivi est un Google Sheet créé dans le dossier Drive TS_FOLDER_ID.
// Le registre (onglet TABLEAUX_SUIVI du classeur CRM) garde : nom, auteur, date, fichier,
// état masqué et la configuration des imports (type + filtres, en JSON).
//
// Chaque import = un onglet dans le Google Sheet du tableau :
//   - colonnes « CRM » (en-têtes bleus) : réécrites à chaque synchronisation, repérées par
//     leur en-tête (l'utilisateur peut trier les lignes ou déplacer les colonnes) ;
//   - colonne « Statut CRM » : tag dynamique (Actif / SORTIE / Hors filtre / Supprimé du CRM / Nouveau) ;
//   - colonnes libres (en-têtes orange) : jamais touchées par le CRM.
// Les lignes sont rattachées au CRM par la colonne « ID CRM » ; une ligne n'est jamais
// supprimée par la synchro, elle est seulement taguée.
//
// Synchro : bouton « Synchroniser » côté CRM + trigger horaire (creerTriggerTableauxSuivi,
// à exécuter UNE fois depuis l'éditeur Apps Script).

var TS_FOLDER_ID = "1XuAnUzZDftjKP_yswZHEO_e9PA58PoBI";
var TS_REGISTRE = "TABLEAUX_SUIVI";
var TS_COL = { ID:0, NOM:1, AUTEUR:2, EMAIL:3, DATE:4, FILE_ID:5, URL:6, MASQUE:7, IMPORTS:8, SYNCHRO:9 };
var TS_HEADERS = ["ID", "Nom", "Auteur", "Email auteur", "Date création", "ID fichier", "URL", "Masqué", "Imports (JSON)", "Dernière synchro"];
var TS_ID_HEADER = "ID CRM";
var TS_TAG_HEADER = "Statut CRM";
var TS_FREE_HEADERS = ["📝 Commentaire", "📌 Action à mener"];
var TS_TYPES_LABEL = { etudiants: "Étudiants", entreprises: "Entreprises", contacts: "Contacts entreprises", offres: "Offres", campagnes: "Campagnes" };

// Colonnes CRM par type (la 1re colonne est toujours l'ID CRM, ajoutée automatiquement)
var TS_FIELDS = {
  etudiants:   ["Nom", "Prénom", "Classe", "Campus", "Typologie", "Statut", "Entreprise", "Mail", "Téléphone", "Mail HECG", "Référent", "Atelier CV", "Atelier LinkedIn", "CVthèque", "Date entrée", "Date sortie", "Motif sortie", "Dernier échange"],
  entreprises: ["Entreprise", "Adresse", "CP", "Ville", "Nb contacts", "Contact principal", "Poste", "Téléphone", "Mail", "Alternants actifs", "Nb alternants", "Dernier échange"],
  contacts:    ["Entreprise", "Nom", "Prénom", "Poste", "Maître d'apprentissage", "Téléphone", "Mail", "Téléphone 2", "Mail 2", "Ville", "Alternants actifs"],
  offres:      ["Poste", "Entreprise", "Ville", "Plateforme", "Date", "Qualité", "État", "Lien", "Nb partages"],
  campagnes:   ["Nom", "Objet", "Type", "Statut", "Campus", "Date événement", "Dernier envoi"]
};

// ── Accès ─────────────────────────────────────────────────────────
function _tabSuiviCanUse_(p) { return !!String(p.vieScRole || "").trim(); }
function _tabSuiviIsOwnerOrAdmin_(p, row) {
  if (String(p.vieScRole || "") === "Super-admin") return true;
  const me = String(p.emailAuteur || "").trim().toLowerCase();
  return !!me && me === String(row[TS_COL.EMAIL] || "").trim().toLowerCase();
}

function _tabSuiviGetRegistre_(ss) {
  let ws = ss.getSheetByName(TS_REGISTRE);
  if (!ws) {
    ws = ss.insertSheet(TS_REGISTRE);
    ws.appendRow(TS_HEADERS);
    ws.setFrozenRows(1);
    ws.getRange(1, 1, 1, TS_HEADERS.length).setFontWeight("bold").setBackground("#1A4E8A").setFontColor("#ffffff");
    ws.getRange(2, TS_COL.DATE + 1, ws.getMaxRows() - 1, 1).setNumberFormat("dd/MM/yyyy HH:mm:ss");
  }
  return ws;
}

function _tabSuiviFindRow_(ws, id) {
  const d = ws.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) if (String(d[i][TS_COL.ID]).trim() === String(id).trim()) return { index: i + 1, row: d[i] };
  return null;
}

function _tabSuiviFmtDate_(v, withTime) {
  if (!v) return "";
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), withTime ? "dd/MM/yyyy HH:mm" : "dd/MM/yyyy");
  return String(v);
}

function _tabSuiviParseImports_(v) {
  try { const a = JSON.parse(String(v || "[]")); return Array.isArray(a) ? a : []; } catch (e) { return []; }
}

// ── Liste ─────────────────────────────────────────────────────────
function handleGetTableauxSuivi(p, ss) {
  if (!_tabSuiviCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  try {
    const d = _tabSuiviGetRegistre_(ss).getDataRange().getValues();
    const me = String(p.emailAuteur || "").trim().toLowerCase();
    const list = [];
    for (let i = 1; i < d.length; i++) {
      if (!d[i][TS_COL.ID]) continue;
      const imports = _tabSuiviParseImports_(d[i][TS_COL.IMPORTS]);
      list.push({
        id: String(d[i][TS_COL.ID]), nom: String(d[i][TS_COL.NOM] || ""),
        auteur: String(d[i][TS_COL.AUTEUR] || ""), emailAuteur: String(d[i][TS_COL.EMAIL] || ""),
        dateCreation: _tabSuiviFmtDate_(d[i][TS_COL.DATE], true),
        dateTri: d[i][TS_COL.DATE] instanceof Date ? d[i][TS_COL.DATE].getTime() : 0,
        url: String(d[i][TS_COL.URL] || ""),
        masque: String(d[i][TS_COL.MASQUE] || "").trim().toLowerCase() === "oui",
        imports: imports.map(function(im) { return { tab: im.tab, type: im.type, resume: im.resume || "" }; }),
        derniereSynchro: _tabSuiviFmtDate_(d[i][TS_COL.SYNCHRO], true),
        estAuteur: !!me && me === String(d[i][TS_COL.EMAIL] || "").trim().toLowerCase()
      });
    }
    list.sort(function(a, b) { return b.dateTri - a.dateTri; });
    return createJsonResponse({ success: true, data: list });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// ── Création ──────────────────────────────────────────────────────
function handleCreateTableauSuivi(p, ss) {
  if (!_tabSuiviCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  const nom = String(p.nom || "").trim();
  if (!nom) return createJsonResponse({ success: false, message: "Le nom du tableau est obligatoire." });
  try {
    const auteur = String(p.nomAuteur || p.emailAuteur || "Inconnu");
    const email = String(p.emailAuteur || "").trim();
    const now = new Date();
    const file = SpreadsheetApp.create(nom);
    const driveFile = DriveApp.getFileById(file.getId());
    driveFile.moveTo(DriveApp.getFolderById(TS_FOLDER_ID));
    if (email && email.indexOf("@") > 0) { try { driveFile.addEditor(email); } catch (eShare) { Logger.log("Partage impossible : " + eShare); } }
    // Onglet d'accueil (supprimé automatiquement au premier import)
    const accueil = file.getSheets()[0];
    accueil.setName("Accueil");
    accueil.getRange("A1").setValue("📋 " + nom).setFontWeight("bold").setFontSize(14);
    accueil.getRange("A2").setValue("Tableau de suivi créé par " + auteur + " le " + Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy à HH:mm"));
    accueil.getRange("A4").setValue("Utilisez « Importer des données » dans le CRM (Vie scolaire → Tableaux de suivis) pour ajouter des listes synchronisées.");

    const id = "TS" + now.getTime();
    const ws = _tabSuiviGetRegistre_(ss);
    ws.appendRow([id, nom, auteur, email, now, file.getId(), file.getUrl(), "Non", "[]", ""]);
    logAction(auteur, p.roleAuteur || p.vieScRole || "", "Création", "Tableau de suivi", id, "A créé le tableau de suivi « " + nom + " »");
    return createJsonResponse({ success: true, id: id, url: file.getUrl() });
  } catch (e) { return createJsonResponse({ success: false, message: "Erreur création : " + e.toString() }); }
}

// ── Masquer / réafficher (jamais de suppression) ──────────────────
function handleMasquerTableauSuivi(p, ss) {
  if (!_tabSuiviCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  try {
    const ws = _tabSuiviGetRegistre_(ss);
    const f = _tabSuiviFindRow_(ws, p.id);
    if (!f) return createJsonResponse({ success: false, message: "Tableau introuvable." });
    if (!_tabSuiviIsOwnerOrAdmin_(p, f.row)) return createJsonResponse({ success: false, message: "Seul l'auteur (ou un Super-admin) peut masquer ce tableau." });
    const masquer = p.masquer !== false;
    ws.getRange(f.index, TS_COL.MASQUE + 1).setValue(masquer ? "Oui" : "Non");
    logAction(p.nomAuteur || p.emailAuteur, p.roleAuteur || p.vieScRole || "", "Modification", "Tableau de suivi", p.id, (masquer ? "A masqué" : "A réaffiché") + " le tableau « " + f.row[TS_COL.NOM] + " »");
    return createJsonResponse({ success: true });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// ── Options des filtres (listes déroulantes de la fenêtre d'import) ──
function handleGetSuiviImportOptions(p, ss) {
  if (!_tabSuiviCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  try {
    const uniq = function(arr) { return Array.from(new Set(arr.map(function(v) { return String(v || "").trim(); }).filter(Boolean))).sort(); };
    const dE = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues().slice(1);
    const dP = getPersonnelDataCached_(ss);
    const referents = [];
    for (let r = 1; r < dP.length; r++) if (dP[r][0]) referents.push({ ref: String(dP[r][0]), nom: (dP[r][2] + " " + dP[r][1]).trim() });
    const dPart = getSheetSafe(ss, "PARTENARIAT").getDataRange().getValues().slice(1);
    const dO = getSheetSafe(ss, "OFFRES").getDataRange().getValues().slice(1);
    let dC = [];
    try { dC = getSheetSafe(ss, "CAMPAGNES").getDataRange().getValues().slice(1); } catch (eC) {}
    return createJsonResponse({ success: true, data: {
      classes: uniq(dE.map(function(r) { return r[14]; })),
      referents: referents,
      villes: uniq(dPart.map(function(r) { return r[12]; })),
      offreQualites: uniq(dO.map(function(r) { return r[9]; })),
      offreEtats: uniq(dO.map(function(r) { return r[13]; })),
      campagneStatuts: uniq(dC.map(function(r) { return r[7]; })),
      campagneTypes: uniq(dC.map(function(r) { return r[5]; }))
    }});
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// ── Lecture CRM + filtres (mêmes règles que le dashboard pour les étudiants) ──
// Retourne { rows: [{ id, values:[...], match:bool, tag:"" }] } pour TOUS les éléments du type,
// avec match = correspond aux filtres. La synchro en a besoin pour les lignes devenues hors filtre.
function _tabSuiviLoadCrm_(ss, type, f, cache) {
  f = f || {};
  cache = cache || {};
  const tz = Session.getScriptTimeZone();
  const fmt = function(v) { return v instanceof Date ? Utilities.formatDate(v, tz, "dd/MM/yyyy") : String(v == null ? "" : v).replace(/^'/, ""); };
  const lower = function(v) { return String(v || "").trim().toLowerCase(); };
  const inList = function(list, v) { return !list || !list.length || list.indexOf(String(v || "").trim()) !== -1; };
  const get = function(name) { return cache[name] || (cache[name] = getSheetSafe(ss, name).getDataRange().getValues()); };
  const out = [];

  if (type === "etudiants") {
    const dE = get("ETUDIANTS");
    const dC = get("Carnet_route");
    const dernier = {};
    for (let i = 1; i < dC.length; i++) {
      const k = String(dC[i][0]).trim(); const t = new Date(dC[i][2]);
      if (!isNaN(t) && (!dernier[k] || t > dernier[k])) dernier[k] = t;
    }
    const dP = getPersonnelDataCached_(ss);
    const refNom = {};
    for (let r = 1; r < dP.length; r++) if (dP[r][0]) refNom[String(dP[r][0]).trim()] = (dP[r][2] + " " + dP[r][1]).trim();
    const now = new Date();
    for (let k = 1; k < dE.length; k++) {
      const r = dE[k]; const id = String(r[0] || "").trim(); if (!id) continue;
      const typo = lower(r[20]), statut = lower(r[13]), motif = lower(r[19]);
      const ech = dernier[id] || null;
      let match;
      if (f.alumni) {
        match = motif.indexOf("dipl") !== -1 && motif.indexOf("obtenu") !== -1;
      } else {
        match = (!f.typos || !f.typos.length || f.typos.indexOf(typo) !== -1)
          && inList(f.classes, r[14])
          && (!f.statuts || !f.statuts.length || f.statuts.some(function(s) { return statut.indexOf(s) !== -1; }));
        const cv = lower(r[15]), li = lower(r[16]), cvt = lower(r[24]);
        if (f.atelier === "CV") match = match && cv === "oui";
        else if (f.atelier === "CV_NON") match = match && cv !== "oui";
        else if (f.atelier === "LI") match = match && li === "oui";
        else if (f.atelier === "LI_NON") match = match && li !== "oui";
        else if (f.atelier === "BOTH") match = match && cv === "oui" && li === "oui";
        if (f.cvTech === "OUI") match = match && cvt === "oui";
        else if (f.cvTech === "NON") match = match && cvt !== "oui";
        if (f.periode) {
          if (f.periode === "NEVER") match = match && !ech;
          else if (!ech) match = false;
          else {
            const diff = (now - ech) / 86400000;
            if (f.periode === "7") match = match && diff <= 7;
            if (f.periode === "7-30") match = match && diff > 7 && diff < 30;
            if (f.periode === "30") match = match && diff >= 30;
          }
        }
      }
      match = match && inList(f.campus, r[22]) && (!f.referent || String(r[7]).trim() === String(f.referent).trim());
      const campusLib = r[22] === "HECG_AIX" ? "Aix-en-Provence" : (r[22] === "HECG" ? "Marseille" : String(r[22] || ""));
      out.push({ id: id, match: match, tag: typo === "sortie" ? "❌ SORTIE" : "", values: [
        String(r[2] || ""), String(r[3] || ""), String(r[14] || ""), campusLib, String(r[20] || ""), String(r[13] || ""),
        String(r[21] || ""), String(r[4] || ""), String(r[5] || ""), String(r[26] || ""), refNom[String(r[7]).trim()] || String(r[7] || ""),
        String(r[15] || ""), String(r[16] || ""), String(r[24] || ""), fmt(r[17]), fmt(r[18]), String(r[19] || ""), ech ? fmt(ech) : ""
      ]});
    }
    return out;
  }

  if (type === "entreprises" || type === "contacts") {
    const dP = get("PARTENARIAT");
    const dE = get("ETUDIANTS");
    const alt = {};
    for (let j = 1; j < dE.length; j++) {
      const idEnt = String(dE[j][11] || "").trim().toUpperCase(); if (!idEnt) continue;
      const st = lower(dE[j][13]);
      if (st.indexOf("altern") !== -1 || st.indexOf("apprenti") !== -1) (alt[idEnt] = alt[idEnt] || []).push((dE[j][3] + " " + dE[j][2]).trim());
    }
    let suivi = [];
    try { suivi = get("SUIVI_PARTENARIAT"); } catch (eS) {}
    const lastEch = {};
    for (let s = 1; s < suivi.length; s++) {
      const k = String(suivi[s][0] || "").trim().toUpperCase(); const t = new Date(suivi[s][2]);
      if (k && !isNaN(t) && (!lastEch[k] || t > lastEch[k])) lastEch[k] = t;
    }
    const fVille = lower(f.ville);
    const matchEnt = function(idEnt, ville) {
      const nbAlt = (alt[idEnt] || []).length;
      let m = !fVille || lower(ville).indexOf(fVille) !== -1;
      if (f.alternants === "AVEC") m = m && nbAlt > 0;
      if (f.alternants === "SANS") m = m && nbAlt === 0;
      return m;
    };
    if (type === "contacts") {
      for (let i = 1; i < dP.length; i++) {
        const r = dP[i]; const id = String(r[0] || "").trim(); if (!id) continue;
        const idEnt = String(r[1] || "").trim().toUpperCase();
        let match = matchEnt(idEnt, r[12]);
        if (f.ma === "OUI") match = match && lower(r[14]) === "oui";
        out.push({ id: id, match: match, tag: "", values: [
          String(r[2] || ""), String(r[3] || ""), String(r[4] || ""), String(r[13] || ""), String(r[14] || "Non"),
          String(r[5] || ""), String(r[6] || ""), String(r[9] || ""), String(r[8] || ""), String(r[12] || ""), (alt[idEnt] || []).join(", ")
        ]});
      }
      return out;
    }
    const ents = {}, order = [];
    for (let i = 1; i < dP.length; i++) {
      const r = dP[i]; const idEnt = String(r[1] || "").trim().toUpperCase(); if (!idEnt) continue;
      if (!ents[idEnt]) { ents[idEnt] = { first: r, nb: 0 }; order.push(idEnt); }
      ents[idEnt].nb++;
    }
    order.forEach(function(idEnt) {
      const r = ents[idEnt].first; const a = alt[idEnt] || [];
      out.push({ id: idEnt, match: matchEnt(idEnt, r[12]), tag: "", values: [
        String(r[2] || ""), String(r[10] || ""), String(r[11] || ""), String(r[12] || ""), ents[idEnt].nb,
        (String(r[4] || "") + " " + String(r[3] || "")).trim(), String(r[13] || ""), String(r[5] || ""), String(r[6] || ""),
        a.join(", "), a.length, lastEch[idEnt] ? fmt(lastEch[idEnt]) : ""
      ]});
    });
    return out;
  }

  if (type === "offres") {
    const dO = get("OFFRES");
    const dPa = get("PARTAGES");
    const nbPart = {};
    for (let s = 1; s < dPa.length; s++) { const k = String(dPa[s][1] || "").trim(); if (k) nbPart[k] = (nbPart[k] || 0) + 1; }
    const fEnt = lower(f.entreprise);
    for (let i = 1; i < dO.length; i++) {
      const r = dO[i]; const id = String(r[0] || "").trim(); if (!id) continue;
      const match = inList(f.qualites, r[9]) && inList(f.etats, r[13]) && (!fEnt || lower(r[4]).indexOf(fEnt) !== -1);
      out.push({ id: id, match: match, tag: lower(r[9]) === "obsolète" ? "⚠️ Obsolète" : "", values: [
        String(r[14] || ""), String(r[4] || ""), String(r[5] || ""), String(r[3] || ""), fmt(r[2]),
        String(r[9] || ""), String(r[13] || ""), String(r[1] || ""), nbPart[id] || 0
      ]});
    }
    return out;
  }

  if (type === "campagnes") {
    const dC = get("CAMPAGNES");
    for (let i = 1; i < dC.length; i++) {
      const r = dC[i]; const id = String(r[0] || "").trim(); if (!id) continue;
      const match = inList(f.statuts, r[7]) && inList(f.types, r[5]);
      out.push({ id: id, match: match, tag: "", values: [
        String(r[1] || ""), String(r[2] || ""), String(r[5] || ""), String(r[7] || ""), String(r[11] || ""), fmt(r[12]), fmt(r[6])
      ]});
    }
    return out;
  }
  throw new Error("Type d'import inconnu : " + type);
}

// Résumé lisible des filtres (affiché dans le CRM)
function _tabSuiviResumeFiltres_(type, f) {
  f = f || {};
  const parts = [];
  const add = function(lib, v) { if (v && (!Array.isArray(v) || v.length)) parts.push(lib + " : " + (Array.isArray(v) ? v.join(", ") : v)); };
  if (type === "etudiants") {
    if (f.alumni) parts.push("Alumni");
    add("Typo", f.typos); add("Classes", f.classes); add("Statuts", f.statuts); add("Campus", f.campus);
    add("Atelier", f.atelier); add("CVthèque", f.cvTech); add("Dernier échange", f.periode); add("Référent", f.referentNom || f.referent);
  } else if (type === "entreprises" || type === "contacts") {
    add("Ville", f.ville); add("Alternants", f.alternants); if (f.ma === "OUI") parts.push("MA uniquement");
  } else if (type === "offres") {
    add("Qualité", f.qualites); add("État", f.etats); add("Entreprise", f.entreprise);
  } else if (type === "campagnes") {
    add("Statut", f.statuts); add("Type", f.types);
  }
  return parts.join(" · ") || "Sans filtre";
}

// ── Aperçu avant import ───────────────────────────────────────────
function handlePreviewSuiviImport(p, ss) {
  if (!_tabSuiviCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  try {
    const rows = _tabSuiviLoadCrm_(ss, p.type, p.filtres || {}).filter(function(r) { return r.match; });
    return createJsonResponse({ success: true, count: rows.length, apercu: rows.slice(0, 15).map(function(r) { return r.values.slice(0, 3).join(" · "); }) });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// ── Import dans un tableau ────────────────────────────────────────
function handleImportSuiviData(p, ss) {
  if (!_tabSuiviCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const type = String(p.type || "");
    if (!TS_FIELDS[type]) return createJsonResponse({ success: false, message: "Type d'import inconnu." });
    const reg = _tabSuiviGetRegistre_(ss);
    const f = _tabSuiviFindRow_(reg, p.id);
    if (!f) return createJsonResponse({ success: false, message: "Tableau introuvable." });
    if (!_tabSuiviIsOwnerOrAdmin_(p, f.row)) return createJsonResponse({ success: false, message: "Seul l'auteur (ou un Super-admin) peut importer dans ce tableau." });

    const file = SpreadsheetApp.openById(String(f.row[TS_COL.FILE_ID]));
    const imports = _tabSuiviParseImports_(f.row[TS_COL.IMPORTS]);
    // Nom d'onglet unique
    let base = String(p.nomOnglet || "").trim() || TS_TYPES_LABEL[type];
    base = base.replace(/[\[\]\*\?\/\\:]/g, " ").substring(0, 90);
    let tab = base, n = 2;
    while (file.getSheetByName(tab)) tab = base + " (" + (n++) + ")";

    const sh = file.insertSheet(tab);
    const crmHeaders = [TS_ID_HEADER].concat(TS_FIELDS[type]).concat([TS_TAG_HEADER]);
    const headers = crmHeaders.concat(TS_FREE_HEADERS);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setFontColor("#ffffff").setWrap(true);
    sh.getRange(1, 1, 1, crmHeaders.length).setBackground("#1A4E8A");
    sh.getRange(1, crmHeaders.length + 1, 1, TS_FREE_HEADERS.length).setBackground("#F26522");
    sh.setFrozenRows(1);
    sh.setFrozenColumns(Math.min(3, crmHeaders.length));
    sh.getRange(1, 1).setNote("Colonnes bleues = synchronisées automatiquement avec le CRM (ne pas les modifier, elles seraient écrasées).\nColonnes orange = libres : ajoutez-en autant que vous voulez à droite.\nVous pouvez trier les lignes : le lien se fait par l'ID CRM.");
    // Format texte : garde le 0 initial des téléphones et évite la conversion des dates/ID
    sh.getRange(2, 1, sh.getMaxRows() - 1, crmHeaders.length).setNumberFormat("@");
    try { sh.getRange(1, 1, sh.getMaxRows(), crmHeaders.length).protect().setDescription("Colonnes synchronisées CRM").setWarningOnly(true); } catch (eProt) {}

    const now = new Date();
    const entry = { tab: tab, type: type, filtres: p.filtres || {}, autoAjout: p.autoAjout !== false, createdAt: now.toISOString(), resume: TS_TYPES_LABEL[type] + " — " + _tabSuiviResumeFiltres_(type, p.filtres) };
    const res = _tabSuiviSyncOnglet_(ss, file, entry, {}, true);

    // Retire l'onglet d'accueil vide après le premier import
    const accueil = file.getSheetByName("Accueil");
    if (accueil && file.getSheets().length > 1) { try { file.deleteSheet(accueil); } catch (eA) {} }

    imports.push(entry);
    reg.getRange(f.index, TS_COL.IMPORTS + 1).setValue(JSON.stringify(imports));
    reg.getRange(f.index, TS_COL.SYNCHRO + 1).setValue(now);
    logAction(p.nomAuteur || p.emailAuteur, p.roleAuteur || p.vieScRole || "", "Modification", "Tableau de suivi", p.id, "Import « " + entry.resume + " » (" + res.ajoutes + " ligne(s)) dans « " + f.row[TS_COL.NOM] + " »");
    return createJsonResponse({ success: true, ajoutes: res.ajoutes, onglet: tab });
  } catch (e) {
    return createJsonResponse({ success: false, message: "Erreur import : " + e.toString() });
  } finally { try { lock.releaseLock(); } catch (eL) {} }
}

// ── Synchronisation d'un onglet ───────────────────────────────────
// Met à jour les colonnes CRM des lignes existantes (repérées par ID CRM), tague les
// lignes, et ajoute en bas les nouveaux éléments correspondant aux filtres si autoAjout.
function _tabSuiviSyncOnglet_(ss, file, entry, cache, isFirst) {
  const sh = file.getSheetByName(entry.tab);
  if (!sh) return { ajoutes: 0, maj: 0, absent: true };
  const fields = TS_FIELDS[entry.type];
  const crm = _tabSuiviLoadCrm_(ss, entry.type, entry.filtres, cache);
  const byId = {};
  crm.forEach(function(r) { byId[String(r.id).toUpperCase()] = r; });

  const lastCol = Math.max(sh.getLastColumn(), 1);
  const lastRow = sh.getLastRow();
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h || "").trim(); });
  const colId = header.indexOf(TS_ID_HEADER);
  const colTag = header.indexOf(TS_TAG_HEADER);
  if (colId === -1 || colTag === -1) return { ajoutes: 0, maj: 0, erreur: "En-têtes ID CRM / Statut CRM introuvables" };
  const colFields = fields.map(function(h) { return header.indexOf(h); });

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
  const tagFor = function(r, existing, oldTag) {
    if (!r) return "🗑️ Supprimé du CRM";
    if (r.tag) return r.tag;
    if (!r.match) return "⚠️ Hors filtre";
    // Conserve le tag « Nouveau » posé lors d'un ajout automatique jusqu'à la synchro suivante du lendemain
    if (existing && /^🆕/.test(oldTag) && oldTag.indexOf(today) !== -1) return oldTag;
    return "✅ Actif";
  };
  const colorFor = function(tag) {
    if (/SORTIE|Supprimé/.test(tag)) return "#fee2e2";
    if (/Hors filtre|Obsolète/.test(tag)) return "#fef9c3";
    if (/^🆕/.test(tag)) return "#dbeafe";
    return "#ffffff";
  };

  const seen = {};
  let maj = 0;
  if (lastRow > 1) {
    const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const bg = sh.getRange(2, 1, lastRow - 1, lastCol).getBackgrounds();
    let changed = false;
    for (let i = 0; i < data.length; i++) {
      const id = String(data[i][colId] || "").trim();
      if (!id) continue; // ligne ajoutée à la main : on n'y touche pas
      seen[id.toUpperCase()] = true;
      const r = byId[id.toUpperCase()];
      if (r) colFields.forEach(function(c, k) { if (c !== -1 && String(data[i][c]) !== String(r.values[k])) { data[i][c] = String(r.values[k]); changed = true; } });
      const tag = tagFor(r, true, String(data[i][colTag] || ""));
      if (data[i][colTag] !== tag) { data[i][colTag] = tag; changed = true; }
      const color = colorFor(tag);
      [colId, colTag].concat(colFields).forEach(function(c) { if (c !== -1) bg[i][c] = color; });
      maj++;
    }
    // Réécrit uniquement les colonnes CRM (les colonnes libres ne sont jamais réécrites)
    const crmCols = [colId, colTag].concat(colFields).filter(function(c) { return c !== -1; });
    if (changed) crmCols.forEach(function(c) {
      sh.getRange(2, c + 1, data.length, 1).setValues(data.map(function(row) { return [row[c]]; }))
        .setBackgrounds(bg.map(function(row) { return [row[c]]; }));
    });
  }

  let ajoutes = 0;
  if (isFirst || entry.autoAjout !== false) {
    const nouveaux = crm.filter(function(r) { return r.match && !seen[String(r.id).toUpperCase()]; });
    if (nouveaux.length) {
      const rows = nouveaux.map(function(r) {
        const line = new Array(lastCol).fill("");
        line[colId] = r.id;
        colFields.forEach(function(c, k) { if (c !== -1) line[c] = String(r.values[k]); });
        line[colTag] = r.tag || (isFirst ? "✅ Actif" : "🆕 Ajouté auto le " + today);
        return line;
      });
      const start = Math.max(sh.getLastRow(), 1) + 1;
      const besoin = start + rows.length - 1 - sh.getMaxRows();
      if (besoin > 0) sh.insertRowsAfter(sh.getMaxRows(), besoin);
      [colId, colTag].concat(colFields).forEach(function(c) { if (c !== -1) sh.getRange(start, c + 1, rows.length, 1).setNumberFormat("@"); });
      sh.getRange(start, 1, rows.length, lastCol).setValues(rows);
      const bgs = rows.map(function(line) { const c = colorFor(line[colTag]); return line.map(function(_, idx) { return (idx === colId || idx === colTag || colFields.indexOf(idx) !== -1) ? c : null; }); });
      sh.getRange(start, 1, rows.length, lastCol).setBackgrounds(bgs);
      ajoutes = rows.length;
    }
  }
  if (isFirst) { try { sh.autoResizeColumns(1, Math.min(lastCol, 25)); } catch (eR) {} }
  return { ajoutes: ajoutes, maj: maj };
}

function _tabSuiviSyncTableau_(ss, regRow, cache) {
  const file = SpreadsheetApp.openById(String(regRow[TS_COL.FILE_ID]));
  const imports = _tabSuiviParseImports_(regRow[TS_COL.IMPORTS]);
  let ajoutes = 0, maj = 0;
  imports.forEach(function(entry) {
    const r = entry.type === "cube" ? _cubeSyncEntry_(ss, file, entry, cache) : _tabSuiviSyncOnglet_(ss, file, entry, cache, false);
    ajoutes += r.ajoutes || 0; maj += r.maj || 0;
  });
  return { ajoutes: ajoutes, maj: maj };
}

// ── Synchro manuelle (bouton CRM) ─────────────────────────────────
function handleSyncTableauSuivi(p, ss) {
  if (!_tabSuiviCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const reg = _tabSuiviGetRegistre_(ss);
    const f = _tabSuiviFindRow_(reg, p.id);
    if (!f) return createJsonResponse({ success: false, message: "Tableau introuvable." });
    const res = _tabSuiviSyncTableau_(ss, f.row, {});
    reg.getRange(f.index, TS_COL.SYNCHRO + 1).setValue(new Date());
    return createJsonResponse({ success: true, ajoutes: res.ajoutes, maj: res.maj });
  } catch (e) {
    return createJsonResponse({ success: false, message: "Erreur synchro : " + e.toString() });
  } finally { try { lock.releaseLock(); } catch (eL) {} }
}

// ── Synchro automatique (trigger horaire) ─────────────────────────
function synchroniserTableauxSuivi() {
  const ss = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");
  const reg = _tabSuiviGetRegistre_(ss);
  const d = reg.getDataRange().getValues();
  const cache = {}; // feuilles CRM lues une seule fois pour tous les tableaux
  const t0 = Date.now();
  for (let i = 1; i < d.length; i++) {
    if (Date.now() - t0 > 4.5 * 60 * 1000) { Logger.log("Arrêt anticipé (limite de temps) — reprise au prochain passage."); break; }
    if (!d[i][TS_COL.ID] || String(d[i][TS_COL.MASQUE]).toLowerCase() === "oui") continue;
    if (!_tabSuiviParseImports_(d[i][TS_COL.IMPORTS]).length) continue;
    try {
      _tabSuiviSyncTableau_(ss, d[i], cache);
      reg.getRange(i + 1, TS_COL.SYNCHRO + 1).setValue(new Date());
    } catch (e) { Logger.log("Synchro " + d[i][TS_COL.ID] + " : " + e); }
  }
}

/**
 * À exécuter UNE SEULE FOIS depuis l'éditeur Apps Script (▶ Exécuter) :
 * installe la synchronisation automatique horaire des tableaux de suivi.
 */
function creerTriggerTableauxSuivi() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "synchroniserTableauxSuivi") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("synchroniserTableauxSuivi").timeBased().everyHours(1).create();
}
