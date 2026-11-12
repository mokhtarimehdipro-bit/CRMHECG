// ==========================================
// VIE SCOLAIRE — SUIVI DES MÉMOIRES · FRISES CHRONOLOGIQUES · PROSPECTS SALON
// (ajouté nov. 2026)
// ==========================================

function _vsxCanUse_(p) { return !!String(p.vieScRole || "").trim() || String(p.roleAuteurValide || p.roleAuteur || "") === "Administrateur"; }
function _vsxIsSuper_(p) { return String(p.vieScRole || "") === "Super-admin" || String(p.emailAuteur || "").toLowerCase() === "m.mokhtari@hecg.fr"; }
function _vsxSheet_(ss, name, headers) {
  let ws = ss.getSheetByName(name);
  if (!ws) {
    ws = ss.insertSheet(name);
    ws.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#1A4E8A").setFontColor("#ffffff");
    ws.setFrozenRows(1);
  }
  return ws;
}
function _vsxCampusLib_(c) { return c === "HECG_AIX" ? "Aix-en-Provence" : (c === "HECG" ? "Marseille" : String(c || "")); }

// Référentiel étudiants actifs + personnel (formateurs) — partagé par les deux modules
function _vsxRef_(ss) {
  const dE = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues();
  const etudiants = [];
  for (let i = 1; i < dE.length; i++) {
    const id = String(dE[i][0] || "").trim(); if (!id) continue;
    const typo = String(dE[i][20] || "").toLowerCase();
    if (typo === "prospect") continue;
    etudiants.push({ id: id, nom: String(dE[i][2] || ""), prenom: String(dE[i][3] || ""), classe: String(dE[i][14] || ""),
                     campus: _vsxCampusLib_(dE[i][22]), sortie: typo === "sortie", typo: typo,
                     mail: String(dE[i][26] || "").trim() || String(dE[i][4] || "").trim(),
                     mailHecg: String(dE[i][26] || "").trim(), mailPerso: String(dE[i][4] || "").trim(), entId: String(dE[i][11] || "").trim() });
  }
  const dP = getPersonnelDataCached_(ss);
  const personnel = [];
  for (let r = 1; r < dP.length; r++) if (dP[r][0]) personnel.push({ ref: String(dP[r][0]), nom: (String(dP[r][2] || "") + " " + String(dP[r][1] || "")).trim(), role: String(dP[r][5] || ""), email: String(dP[r][6] || "").trim().toLowerCase() });
  // Formateurs actifs : feuille « Adultes » (ou FORMATEURS), état formateur = OUI, pas sorti
  const formateurs = [];
  try {
    let wa; try { wa = getSheetSafe(ss, "Adultes"); } catch (eA) { wa = getSheetSafe(ss, "FORMATEURS"); }
    const dA = wa.getDataRange().getValues();
    const now = new Date();
    for (let r = 1; r < dA.length; r++) {
      if (!dA[r][0]) continue;
      if (String(dA[r][12] || "NON").trim().toUpperCase() !== "OUI") continue;
      const ds = dA[r][10];
      if (ds instanceof Date && ds < now) continue;
      formateurs.push({ ref: "A:" + String(dA[r][0]), nom: (String(dA[r][2] || "") + " " + String(dA[r][1] || "").toUpperCase()).trim(), matieres: String(dA[r][8] || ""),
                        email: String(dA[r][4] || "").trim() || String(dA[r][3] || "").trim() });
    }
  } catch (eF) { Logger.log("Formateurs : " + eF); }
  formateurs.sort(function(a, b) { return a.nom.localeCompare(b.nom); });
  return { etudiants: etudiants, personnel: personnel, formateurs: formateurs };
}

// ═══════════════════════════ MÉMOIRES ═══════════════════════════
var MEM_HEADERS = ["ID", "ID étudiant", "Étudiant", "Classe", "Campus", "Sujet", "Mots-clés", "Réf. formateur", "Formateur référent", "Statut", "Notes", "Créé le", "Créé par", "Mis à jour le", "Archivé", "ID entreprise", "Entreprise"];
// Colonnes P/Q ajoutées le 12/11/2026 : on complète l'en-tête des feuilles existantes
function _memSheet_(ss) {
  const ws = _vsxSheet_(ss, "MEMOIRES", MEM_HEADERS);
  if (ws.getLastColumn() < MEM_HEADERS.length || String(ws.getRange(1, MEM_HEADERS.length).getValue()) !== MEM_HEADERS[MEM_HEADERS.length - 1])
    ws.getRange(1, 16, 1, 2).setValues([MEM_HEADERS.slice(15)]).setFontWeight("bold").setBackground("#1A4E8A").setFontColor("#ffffff");
  return ws;
}
// Référentiel entreprises (feuille Partenariat : col. B = ID entreprise, col. C = nom)
function _memEntreprises_(ss) {
  const d = getSheetSafe(ss, "Partenariat").getDataRange().getValues();
  const vu = {}, out = [];
  for (let i = 1; i < d.length; i++) {
    const id = String(d[i][1] || "").trim(), nom = String(d[i][2] || "").trim();
    if (!id || !nom || vu[id.toUpperCase()]) continue;
    vu[id.toUpperCase()] = true; out.push({ id: id, nom: nom });
  }
  return out.sort(function(a, b) { return a.nom.localeCompare(b.nom); });
}

function handleGetMemoires(p, ss) {
  if (!_vsxCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  return getCachedJsonResponse_('memoires_v1', 120, function() {
    try {
      const ref = _vsxRef_(ss);
      const byId = {}; ref.etudiants.forEach(function(e) { byId[e.id] = e; });
      const d = _memSheet_(ss).getDataRange().getValues();
      const list = [];
      for (let i = 1; i < d.length; i++) {
        if (!d[i][0] || String(d[i][14]) === "Oui") continue;
        const e = byId[String(d[i][1])];
        list.push({ id: String(d[i][0]), etuId: String(d[i][1]), etudiant: e ? (e.prenom + " " + e.nom).trim() : String(d[i][2] || ""),
          classe: e ? e.classe : String(d[i][3] || ""), campus: e ? e.campus : String(d[i][4] || ""), sujet: String(d[i][5] || ""),
          motsCles: String(d[i][6] || "").split(",").map(function(s) { return s.trim(); }).filter(Boolean),
          formateurRef: String(d[i][7] || ""), formateur: String(d[i][8] || ""), statut: String(d[i][9] || "En cours"), notes: String(d[i][10] || ""),
          creeLe: d[i][11] instanceof Date ? Utilities.formatDate(d[i][11], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(d[i][11] || ""),
          majLe: d[i][13] instanceof Date ? Utilities.formatDate(d[i][13], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(d[i][13] || ""),
          entrepriseId: String(d[i][15] || ""), entreprise: String(d[i][16] || "") });
      }
      const docs = _memDocsByMemoire_(ss);
      const motsClesRef = _memMotsCles_(ss);
      const mailEtu = {}; ref.etudiants.forEach(function(e) { mailEtu[e.id] = e.mail; });
      const mailForm = {}, mailFormNom = {};
      ref.formateurs.concat(ref.personnel).forEach(function(x) { if (x.email) { mailForm[x.ref] = x.email; mailFormNom[String(x.nom).toLowerCase().replace(/\s+/g, " ")] = x.email; } });
      list.forEach(function(m) {
        m.etuMail = mailEtu[m.etuId] || "";
        m.formateurMail = mailForm[m.formateurRef] || mailFormNom[String(m.formateur).toLowerCase().replace(/\s+/g, " ")] || "";
      });
      const ech = _memEchangesByMemoire_(ss);
      list.forEach(function(m) { m.docs = docs[m.id] || []; m.echanges = ech[m.id] || []; });
      let entreprises = []; try { entreprises = _memEntreprises_(ss); } catch (eE) { Logger.log("Entreprises mémoires : " + eE); }
      return createJsonResponse({ success: true, data: list, etudiants: ref.etudiants.filter(function(e) { return !e.sortie; }), personnel: ref.personnel, formateurs: ref.formateurs, motsClesRef: motsClesRef, entreprises: entreprises });
    } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
  });
}

function handleSaveMemoire(p, ss) {
  if (!_vsxCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  try {
    const ws = _memSheet_(ss);
    const m = p.memoire || {};
    if (!m.etuId) return createJsonResponse({ success: false, message: "Choisissez un étudiant." });
    const ref = _vsxRef_(ss);
    const e = ref.etudiants.filter(function(x) { return x.id === String(m.etuId); })[0] || {};
    const f = ref.formateurs.concat(ref.personnel).filter(function(x) { return x.ref === String(m.formateurRef); })[0] || {};
    const mots = (Array.isArray(m.motsCles) ? m.motsCles : String(m.motsCles || "").split(",")).map(function(s) { return String(s).trim(); }).filter(Boolean).join(", ");
    const now = new Date();
    const rowVals = [m.id || ("MEM" + now.getTime()), String(m.etuId), ((e.prenom || "") + " " + (e.nom || "")).trim(), e.classe || "", e.campus || "",
      String(m.sujet || ""), mots, f.nom ? String(m.formateurRef || "") : "", f.nom || String(m.formateur || "").trim(), String(m.statut || "En cours"), String(m.notes || ""), now,
      String(p.nomAuteur || p.emailAuteur || ""), now, "", String(m.entrepriseId || "").trim(), String(m.entreprise || "").trim()];
    if (m.id) {
      const d = ws.getDataRange().getValues();
      for (let i = 1; i < d.length; i++) if (String(d[i][0]) === String(m.id)) {
        rowVals[11] = d[i][11]; rowVals[12] = d[i][12];
        if (m.archiver) { rowVals[14] = "Oui"; rowVals[15] = d[i][15] || ""; rowVals[16] = d[i][16] || ""; }
        ws.getRange(i + 1, 1, 1, MEM_HEADERS.length).setValues([rowVals]);
        logAction(p.nomAuteur || p.emailAuteur, p.vieScRole || "", m.archiver ? "Suppression" : "Modification", "Vie scolaire", m.id, (m.archiver ? "Mémoire retiré : " : "Mémoire modifié : ") + rowVals[2]);
        return createJsonResponse({ success: true, id: m.id });
      }
      return createJsonResponse({ success: false, message: "Mémoire introuvable." });
    }
    ws.appendRow(rowVals);
    logAction(p.nomAuteur || p.emailAuteur, p.vieScRole || "", "Création", "Vie scolaire", rowVals[0], "Suivi de mémoire : " + rowVals[2] + " — " + String(m.sujet || "").substring(0, 60));
    return createJsonResponse({ success: true, id: rowVals[0] });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// ═══════════════════════════ FRISES CHRONOLOGIQUES ═══════════════════════════
var FRISE_HEADERS = ["ID", "Titre", "Description", "Couleur", "Destinataires", "Rappels (jours avant)", "Créé par", "Email créateur", "Créé le", "Archivée", "Participants"];
var ETAPE_HEADERS = ["ID", "ID frise", "Date", "Titre", "À rendre", "Statut", "Rappels envoyés", "Rendu le"];

function _friseVisible_(p, row) {
  if (_vsxIsSuper_(p)) return true;
  const me = String(p.emailAuteur || "").toLowerCase();
  return String(row[7] || "").toLowerCase() === me || String(row[4] || "").toLowerCase().split(/[,;\s]+/).indexOf(me) !== -1;
}

function handleGetFrises(p, ss) {
  if (!_vsxCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  try {
    const tz = Session.getScriptTimeZone();
    const fd = _vsxSheet_(ss, "FRISES", FRISE_HEADERS).getDataRange().getValues();
    const ed = _vsxSheet_(ss, "FRISES_ETAPES", ETAPE_HEADERS).getDataRange().getValues();
    const etapes = {};
    for (let i = 1; i < ed.length; i++) {
      if (!ed[i][0]) continue;
      (etapes[String(ed[i][1])] = etapes[String(ed[i][1])] || []).push({ id: String(ed[i][0]),
        date: ed[i][2] instanceof Date ? Utilities.formatDate(ed[i][2], tz, "yyyy-MM-dd") : String(ed[i][2] || ""),
        titre: String(ed[i][3] || ""), aRendre: String(ed[i][4] || ""), statut: String(ed[i][5] || "À faire"),
        renduLe: ed[i][7] instanceof Date ? Utilities.formatDate(ed[i][7], tz, "yyyy-MM-dd") : String(ed[i][7] || "") });
    }
    const list = [];
    for (let i = 1; i < fd.length; i++) {
      if (!fd[i][0] || String(fd[i][9]) === "Oui" || !_friseVisible_(p, fd[i])) continue;
      const et = (etapes[String(fd[i][0])] || []).sort(function(a, b) { return a.date < b.date ? -1 : 1; });
      list.push({ id: String(fd[i][0]), titre: String(fd[i][1] || ""), description: String(fd[i][2] || ""), couleur: String(fd[i][3] || "#1A4E8A"),
        destinataires: String(fd[i][4] || ""), rappels: String(fd[i][5] || "7,2,0"), createur: String(fd[i][6] || ""), emailCreateur: String(fd[i][7] || ""),
        creeLe: fd[i][8] instanceof Date ? Utilities.formatDate(fd[i][8], tz, "yyyy-MM-dd") : "", etapes: et,
        participants: _friseParticipants_(fd[i][10]) });
    }
    // Référentiel du sélecteur de participants : étudiants actifs, équipe, formateurs (avec mail)
    const ref = _vsxRef_(ss), personnes = [];
    ref.etudiants.forEach(function(e) { if (!e.sortie && e.mail.indexOf("@") > 0) personnes.push({ t: "etu", id: e.id, nom: (e.prenom + " " + e.nom).trim(), info: e.classe, email: e.mail.toLowerCase() }); });
    ref.personnel.forEach(function(x) { if (x.email.indexOf("@") > 0) personnes.push({ t: "equipe", id: x.ref, nom: x.nom, info: x.role, email: x.email }); });
    ref.formateurs.forEach(function(x) { if (x.email.indexOf("@") > 0) personnes.push({ t: "form", id: x.ref, nom: x.nom, info: x.matieres, email: x.email.toLowerCase() }); });
    return createJsonResponse({ success: true, data: list, personnes: personnes });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleSaveFrise(p, ss) {
  if (!_vsxCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const f = p.frise || {};
    const fws = _vsxSheet_(ss, "FRISES", FRISE_HEADERS), ews = _vsxSheet_(ss, "FRISES_ETAPES", ETAPE_HEADERS);
    const now = new Date();
    const me = String(p.emailAuteur || "").toLowerCase();
    const rappels = String(f.rappels || "7,2,0").split(/[,;\s]+/).map(Number).filter(function(n) { return !isNaN(n) && n >= 0 && n <= 60; }).join(",") || "7,2,0";
    const parts = _friseParticipants_(f.participants);
    const emails = String(f.destinataires || (parts.length ? "" : me)).toLowerCase().split(/[,;\s]+/).filter(function(s) { return s.indexOf("@") > 0; });
    parts.forEach(function(x) { if (emails.indexOf(x.email) === -1) emails.push(x.email); });
    const dest = emails.join(", "), partsJson = JSON.stringify(parts);
    if (!String(fws.getRange(1, 11).getValue() || "")) fws.getRange(1, 11).setValue("Participants").setFontWeight("bold").setBackground("#1A4E8A").setFontColor("#ffffff");
    let id = f.id;
    if (id) {
      const d = fws.getDataRange().getValues(); let found = false;
      for (let i = 1; i < d.length; i++) if (String(d[i][0]) === String(id)) {
        if (!_vsxIsSuper_(p) && String(d[i][7]).toLowerCase() !== me) return createJsonResponse({ success: false, message: "Seul le créateur peut modifier cette frise." });
        fws.getRange(i + 1, 2, 1, 5).setValues([[String(f.titre || ""), String(f.description || ""), String(f.couleur || "#1A4E8A"), dest, rappels]]);
        if (f.archiver) fws.getRange(i + 1, 10).setValue("Oui");
        else if (f.participants !== undefined) fws.getRange(i + 1, 11).setValue(partsJson);
        found = true; break;
      }
      if (!found) return createJsonResponse({ success: false, message: "Frise introuvable." });
    } else {
      if (!String(f.titre || "").trim()) return createJsonResponse({ success: false, message: "Donnez un titre à la frise." });
      id = "FR" + now.getTime();
      fws.appendRow([id, String(f.titre), String(f.description || ""), String(f.couleur || "#1A4E8A"), dest, rappels, String(p.nomAuteur || me), me, now, "", partsJson]);
    }
    // Étapes : remplacement de la liste (on conserve l'historique des rappels envoyés par ID)
    if (Array.isArray(f.etapes)) {
      const ed = ews.getDataRange().getValues();
      const old = {};
      for (let i = ed.length - 1; i >= 1; i--) if (String(ed[i][1]) === String(id)) { old[String(ed[i][0])] = ed[i]; ews.deleteRow(i + 1); }
      const rows = f.etapes.filter(function(e) { return e.date && String(e.titre || "").trim(); }).map(function(e, k) {
        const o = old[String(e.id)] || [];
        const eid = e.id && old[String(e.id)] ? String(e.id) : ("ET" + now.getTime() + k);
        const sameDate = o.length && (o[2] instanceof Date ? Utilities.formatDate(o[2], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(o[2])) === String(e.date);
        return [eid, id, String(e.date), String(e.titre).trim(), String(e.aRendre || ""), e.statut === "Rendu" ? "Rendu" : "À faire",
                sameDate ? String(o[6] || "") : "", e.statut === "Rendu" ? (o[7] || now) : ""];
      });
      if (rows.length) {
        const start = ews.getLastRow() + 1;
        ews.getRange(start, 3, rows.length, 1).setNumberFormat("@");
        ews.getRange(start, 1, rows.length, ETAPE_HEADERS.length).setValues(rows);
      }
    }
    _ensureFriseTrigger_();
    logAction(p.nomAuteur || me, p.vieScRole || "", f.id ? (f.archiver ? "Suppression" : "Modification") : "Création", "Vie scolaire", id, "Frise : " + String(f.titre || "") + (parts.length ? " — " + parts.length + " participant(s) aux rappels" : ""));
    return createJsonResponse({ success: true, id: id });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
  finally { try { lock.releaseLock(); } catch (eL) {} }
}

function handleToggleEtapeFrise(p, ss) {
  if (!_vsxCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  try {
    const ews = _vsxSheet_(ss, "FRISES_ETAPES", ETAPE_HEADERS);
    const d = ews.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) if (String(d[i][0]) === String(p.id)) {
      ews.getRange(i + 1, 6, 1, 3).setValues([[p.rendu ? "Rendu" : "À faire", d[i][6], p.rendu ? new Date() : ""]]);
      return createJsonResponse({ success: true });
    }
    return createJsonResponse({ success: false, message: "Étape introuvable." });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// Participants d'une frise : [{t: "etu"|"equipe"|"form", id, nom, email}] (colonne K, JSON)
function _friseParticipants_(v) {
  let a = v;
  if (typeof a === "string") { try { a = a ? JSON.parse(a) : []; } catch (e) { a = []; } }
  if (!Array.isArray(a)) return [];
  const vu = {};
  return a.map(function(x) { return { t: ["etu", "equipe", "form"].indexOf(x && x.t) !== -1 ? x.t : "equipe", id: String(x && x.id || ""), nom: String(x && x.nom || ""), email: String(x && x.email || "").trim().toLowerCase() }; })
    .filter(function(x) { if (x.email.indexOf("@") < 1 || vu[x.email]) return false; vu[x.email] = 1; return true; });
}

// Filet de sécurité : la limite de 20 déclencheurs peut empêcher la création de « rappelsFrises ».
// auditAdressesTick (toutes les 5 min) lance donc aussi les rappels une fois par jour à partir de 7 h.
// rappelsFrises mémorise par étape les rappels déjà envoyés : pas de doublon si le déclencheur existe aussi.
function _frisesAuto_() {
  const props = PropertiesService.getScriptProperties(), now = new Date();
  if (Number(Utilities.formatDate(now, "Europe/Paris", "H")) < 7) return;
  const auj = Utilities.formatDate(now, "Europe/Paris", "yyyy-MM-dd");
  if (props.getProperty("FRISES_RAPPEL_JOUR") === auj) return;
  props.setProperty("FRISES_RAPPEL_JOUR", auj);
  rappelsFrises();
}

function _ensureFriseTrigger_() {
  try {
    const exists = ScriptApp.getProjectTriggers().some(function(t) { return t.getHandlerFunction() === "rappelsFrises"; });
    if (!exists) ScriptApp.newTrigger("rappelsFrises").timeBased().everyDays(1).atHour(7).create();
  } catch (e) { Logger.log("Trigger frises : " + e); }
}

/** Déclencheur quotidien (7h) : mails de rappel avant chaque échéance de frise. */
function rappelsFrises() {
  const ss = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");
  const tz = Session.getScriptTimeZone();
  const fws = _vsxSheet_(ss, "FRISES", FRISE_HEADERS), ews = _vsxSheet_(ss, "FRISES_ETAPES", ETAPE_HEADERS);
  const fd = fws.getDataRange().getValues(), ed = ews.getDataRange().getValues();
  const frises = {};
  for (let i = 1; i < fd.length; i++) if (fd[i][0] && String(fd[i][9]) !== "Oui") frises[String(fd[i][0])] = fd[i];
  const today = new Date(Utilities.formatDate(new Date(), tz, "yyyy-MM-dd") + "T12:00:00");
  const parDest = {}; // email → [{frise, étape, jours}]
  const qui = {};     // email → participant (prénom, type) pour personnaliser le mail
  Object.keys(frises).forEach(function(k) { _friseParticipants_(frises[k][10]).forEach(function(x) { if (!qui[x.email] || x.t === "etu") qui[x.email] = x; }); });
  for (let i = 1; i < ed.length; i++) {
    const fr = frises[String(ed[i][1])];
    if (!fr || String(ed[i][5]) === "Rendu") continue;
    const ds = ed[i][2] instanceof Date ? Utilities.formatDate(ed[i][2], tz, "yyyy-MM-dd") : String(ed[i][2]);
    const dEt = new Date(ds + "T12:00:00"); if (isNaN(dEt)) continue;
    const jours = Math.round((dEt - today) / 86400000);
    const offsets = String(fr[5] || "7,2,0").split(",").map(Number);
    const deja = String(ed[i][6] || "").split(",").filter(Boolean);
    // On rappelle le jour J-n prévu ; si un rappel a été manqué (trigger absent), le plus proche à venir part quand même.
    const due = offsets.filter(function(n) { return jours <= n && jours >= 0 && deja.indexOf(String(n)) === -1; });
    if (!due.length) continue;
    const n = Math.min.apply(null, due);
    String(fr[4] || fr[7] || "").split(/[,;\s]+/).filter(function(s) { return s.indexOf("@") > 0; }).forEach(function(em) {
      (parDest[em] = parDest[em] || []).push({ fr: fr, et: ed[i], ds: ds, jours: jours });
    });
    ews.getRange(i + 1, 7).setValue(deja.concat(due.map(String)).join(","));
  }
  let envoyes = 0;
  Object.keys(parDest).forEach(function(em) {
    const items = parDest[em].sort(function(a, b) { return a.jours - b.jours; });
    const esc = function(s) { return String(s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); };
    const quand = function(j) { return j === 0 ? "Aujourd'hui" : (j === 1 ? "Demain" : "Dans " + j + " jours"); };
    const lignes = items.map(function(it) {
      const col = String(it.fr[3] || "#1A4E8A");
      return '<tr><td style="padding:14px 0;border-bottom:1px solid #eef2f7;vertical-align:top;width:110px"><span style="display:inline-block;background:' + (it.jours <= 1 ? '#fde8e8;color:#c62828' : '#fff4ed;color:#c2410c') + ';font-weight:700;font-size:12px;padding:4px 10px;border-radius:999px">' + quand(it.jours) + '</span></td>'
        + '<td style="padding:14px 0 14px 12px;border-bottom:1px solid #eef2f7"><div style="font-size:12px;color:' + col + ';font-weight:700;text-transform:uppercase;letter-spacing:.06em">' + esc(it.fr[1]) + '</div>'
        + '<div style="font-size:16px;font-weight:700;color:#0f1b2d;margin-top:2px">' + esc(it.et[3]) + '</div>'
        + (it.et[4] ? '<div style="font-size:13.5px;color:#475569;margin-top:4px">À rendre : ' + esc(it.et[4]) + '</div>' : '')
        + '<div style="font-size:12.5px;color:#94a3b8;margin-top:4px">Échéance : ' + Utilities.formatDate(new Date(it.ds + "T12:00:00"), tz, "EEEE d MMMM yyyy") + '</div></td></tr>';
    }).join("");
    const pq = qui[em], etu = !!(pq && pq.t === "etu");
    const prenom = pq && pq.nom ? (etu ? pq.nom.split(/\s+/)[0] : pq.nom) : "";
    // Les étudiants n'ont pas accès à Vie scolaire : pas de lien vers les frises pour eux.
    const html = '<div style="font-family:Segoe UI,Arial,sans-serif;background:#f5f7fb;padding:28px"><div style="max-width:580px;margin:0 auto;background:#fff;border-radius:16px;border:1px solid #e6ebf2;overflow:hidden">'
      + '<div style="background:#1A4E8A;color:#fff;padding:20px 26px;font-size:18px;font-weight:700">⏰ Prochaines échéances</div>'
      + '<div style="padding:8px 26px 24px">'
      + (prenom ? '<p style="font-size:14.5px;color:#1f2937;margin:16px 0 4px">Bonjour ' + esc(prenom) + ',</p><p style="font-size:14px;color:#475569;margin:0 0 6px">' + (items.length > 1 ? 'voici les échéances qui approchent :' : "voici l'échéance qui approche :") + '</p>' : '')
      + '<table style="width:100%;border-collapse:collapse">' + lignes + '</table>'
      + (etu ? '<p style="font-size:13px;color:#64748b;margin-top:18px">Une question ? Contactez le pôle pédagogique : polepedagogique@hecg.fr</p>'
             : '<a href="https://mokhtarimehdipro-bit.github.io/CRMHECG/vie_scolaire.html" style="display:inline-block;margin-top:20px;background:#F26522;color:#fff;text-decoration:none;font-weight:700;padding:12px 22px;border-radius:10px">Ouvrir les frises</a>')
      + '</div></div></div>';
    try { MailApp.sendEmail({ to: em, subject: "[HECG] ⏰ " + (items[0].jours === 0 ? "Échéance aujourd'hui" : "Échéance proche") + " : " + items[0].et[3], htmlBody: html, name: "HECG — Rappel d'échéance" }); envoyes++; }
    catch (e) { Logger.log("Rappel frise " + em + " : " + e); }
  });
  if (envoyes) { try { logAction("Rappel automatique", "", "Action", "Vie scolaire — frises (rappels)", "FRISES", envoyes + " rappel(s) d'échéance envoyé(s) par mail"); } catch (eL) {} }
}

// ═══════════════════════════ PROSPECTS SALON (salon.html) ═══════════════════════════
/**
 * Reçoit un lot de prospects saisis (éventuellement hors ligne) et les ajoute dans ETUDIANTS
 * en typologie PROSPECT. Idempotent : chaque fiche porte un identifiant client (cid) —
 * un renvoi après coupure réseau ne crée jamais de doublon. Un prospect déjà présent
 * (même mail ou même téléphone) n'est pas recréé.
 */
function handleAddProspectsSalon(p, ss) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    const items = Array.isArray(p.prospects) ? p.prospects.slice(0, 200) : [];
    if (!items.length) return createJsonResponse({ success: true, resultats: [] });
    const ws = getSheetSafe(ss, "ETUDIANTS");
    const d = ws.getDataRange().getValues();
    const normTel = function(t) { return String(t || "").replace(/\D/g, "").replace(/^33/, "0").slice(-9); };
    const byMail = {}, byTel = {};
    for (let i = 1; i < d.length; i++) {
      const m = String(d[i][4] || "").trim().toLowerCase(); if (m) byMail[m] = String(d[i][0]);
      const t = normTel(d[i][5]); if (t.length === 9) byTel[t] = String(d[i][0]);
    }
    const cache = CacheService.getScriptCache();
    const res = [], rows = [], notes = [];
    const auteur = String(p.nomAuteur || p.idAuteur || p.emailAuteur || "Salon");
    const colEvt = _salonColByHeader_(ws, SALON_EVT_COL_HEADER);
    const evtSheet = _salonEvtSheet_(ss);
    const evtLib = function(it) { return it.salon ? String(it.salon).trim() + (it.salonDate ? " (" + _salonFmtDateFr_(it.salonDate) + ")" : "") : ""; };
    const majEvtExistant = {}; // ID étudiant existant → libellés d'événements à ajouter
    let base = Date.now();
    items.forEach(function(it) {
      const cid = String(it.cid || "");
      const seen = cid ? cache.get("salon_cid_" + cid) : null;
      if (seen) { res.push({ cid: cid, statut: "deja", id: seen }); return; }
      const nom = String(it.nom || "").trim().toUpperCase(), prenom = String(it.prenom || "").trim();
      if (!nom && !prenom) { res.push({ cid: cid, statut: "erreur", message: "Nom manquant" }); return; }
      if (it.salon) _salonEvtUpsert_(evtSheet, it.salon, it.salonDate, it.salonLieu, auteur);
      const lib = evtLib(it);
      const mail = String(it.mail || "").trim().toLowerCase(), tel = String(it.tel || "").trim();
      const dup = (mail && byMail[mail]) || (normTel(tel).length === 9 && byTel[normTel(tel)]);
      if (dup) {
        if (cid) cache.put("salon_cid_" + cid, dup, 21600);
        if (lib) { (majEvtExistant[dup] = majEvtExistant[dup] || []).push(lib); notes.push([dup, auteur, new Date(), "📍 Rencontré de nouveau lors de l'événement « " + lib + " »", false]); }
        res.push({ cid: cid, statut: "existant", id: dup }); return;
      }
      const id = "ETU" + String(base++).slice(-6);
      const row = [id, "P@ss123", nom, prenom, mail, tel, "", "", "", "", "", "", "", "", String(it.classe || ""), "Non", "Non",
                   "", "", "", "PROSPECT", "", String(it.campus || ""), "", "", "", ""];
      _withDateCreation_(ws, row);
      while (row.length < colEvt) row.push("");
      row[colEvt - 1] = lib;
      rows.push(row);
      if (lib) notes.push([id, auteur, new Date(), "📍 Rencontré lors de l'événement « " + lib + " »", false]);
      if (mail) byMail[mail] = id; if (normTel(tel).length === 9) byTel[normTel(tel)] = id;
      if (cid) cache.put("salon_cid_" + cid, id, 21600);
      res.push({ cid: cid, statut: "ajoute", id: id });
      logAction(auteur, p.roleAuteur || "", "Création", "Dossier étudiant", id,
        "Prospect salon" + (lib ? " « " + lib + " »" : " (sans événement)") + " : " + prenom + " " + nom);
    });
    if (rows.length) {
      const width = Math.max.apply(null, rows.map(function(r) { return r.length; }));
      const norm = rows.map(function(r) { while (r.length < width) r.push(""); return r; });
      ws.getRange(ws.getLastRow() + 1, 1, norm.length, width).setValues(norm);
    }
    // Prospects déjà connus : l'événement est ajouté à leur fiche (sans doublon)
    Object.keys(majEvtExistant).forEach(function(idEtu) {
      for (let i = 1; i < d.length; i++) if (String(d[i][0]) === idEtu) {
        const cell = ws.getRange(i + 1, colEvt);
        const actuels = String(cell.getValue() || "").split(" | ").map(function(s) { return s.trim(); }).filter(Boolean);
        majEvtExistant[idEtu].forEach(function(l) { if (actuels.indexOf(l) === -1) actuels.push(l); });
        cell.setValue(actuels.join(" | "));
        try { _invalidateStudentDetailsCache_(idEtu); } catch (eC) {}
        logAction(auteur, p.roleAuteur || "", "Modification", "Dossier étudiant", idEtu, "Événement ajouté (déjà dans le CRM) : " + majEvtExistant[idEtu].join(", "));
        break;
      }
    });
    if (notes.length) { const wc = getSheetSafe(ss, "Carnet_route"); wc.getRange(wc.getLastRow() + 1, 1, notes.length, 5).setValues(notes); }
    return createJsonResponse({ success: true, resultats: res });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
  finally { try { lock.releaseLock(); } catch (eL) {} }
}

function handleGetSalonRef(p, ss) {
  return getCachedJsonResponse_('salon_ref_v2', 600, function() {
    try {
      const d = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues();
      const cl = {};
      for (let i = 1; i < d.length; i++) { const c = String(d[i][14] || "").trim(); if (c) cl[c] = true; }
      return createJsonResponse({ success: true, classes: Object.keys(cl).sort(), evenements: _salonEvtList_(ss) });
    } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
  });
}

// ── Événements (salons, JPO, forums…) : feuille SALON_EVENEMENTS ──
// Créés depuis salon.html (y compris hors ligne : l'événement est créé côté serveur au premier
// envoi d'un prospect qui le porte). Le libellé « Nom (jj/mm/aaaa) » est écrit dans la colonne
// « Événements » de la fiche étudiant (plusieurs séparés par « | ») + note dans le carnet de route.
var SALON_EVT_HEADERS = ["ID", "Nom", "Date", "Lieu", "Créé par", "Créé le", "Archivé"];
var SALON_EVT_COL_HEADER = "Événements";

function _salonFmtDateFr_(iso) { const m = String(iso || "").match(/^(\d{4})-(\d{2})-(\d{2})/); return m ? m[3] + "/" + m[2] + "/" + m[1] : String(iso || ""); }
function _salonNorm_(s) { return String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/\s+/g, " ").trim(); }
function _salonEvtSheet_(ss) { return _vsxSheet_(ss, "SALON_EVENEMENTS", SALON_EVT_HEADERS); }

// Colonne repérée par son en-tête dans ETUDIANTS (créée en fin de feuille si absente)
function _salonColByHeader_(ws, header) {
  const lastCol = Math.max(ws.getLastColumn(), 1);
  const h = ws.getRange(1, 1, 1, lastCol).getValues()[0];
  for (let c = 0; c < h.length; c++) if (_salonNorm_(h[c]) === _salonNorm_(header)) return c + 1;
  if (ws.getMaxColumns() < lastCol + 1) ws.insertColumnsAfter(ws.getMaxColumns(), 1);
  ws.getRange(1, lastCol + 1).setValue(header).setFontWeight("bold");
  return lastCol + 1;
}

function _salonEvtList_(ss) {
  const d = _salonEvtSheet_(ss).getDataRange().getValues();
  const tz = Session.getScriptTimeZone(), out = [];
  for (let i = 1; i < d.length; i++) {
    if (!d[i][0] || String(d[i][6]) === "Oui") continue;
    out.push({ id: String(d[i][0]), nom: String(d[i][1] || ""), date: d[i][2] instanceof Date ? Utilities.formatDate(d[i][2], tz, "yyyy-MM-dd") : String(d[i][2] || ""), lieu: String(d[i][3] || "") });
  }
  return out.sort(function(a, b) { return (b.date || "").localeCompare(a.date || ""); });
}

// Crée l'événement s'il n'existe pas déjà (même nom + même date) ; renvoie son ID
function _salonEvtUpsert_(ws, nom, date, lieu, auteur) {
  nom = String(nom || "").trim(); if (!nom) return "";
  date = String(date || "").trim();
  const d = ws.getDataRange().getValues(), tz = Session.getScriptTimeZone();
  for (let i = 1; i < d.length; i++) {
    const di = d[i][2] instanceof Date ? Utilities.formatDate(d[i][2], tz, "yyyy-MM-dd") : String(d[i][2] || "");
    if (_salonNorm_(d[i][1]) === _salonNorm_(nom) && (!date || !di || di === date)) return String(d[i][0]);
  }
  const id = "EVT" + Date.now() + Math.floor(Math.random() * 10);
  ws.appendRow([id, nom, "", String(lieu || ""), auteur, new Date(), ""]);
  ws.getRange(ws.getLastRow(), 3).setNumberFormat("@").setValue(date);
  try { logAction(auteur, "", "Création", "Salon — Événements", id, "Événement créé : " + nom + (date ? " (" + _salonFmtDateFr_(date) + ")" : "") + (lieu ? " — " + lieu : "")); } catch (eL) {}
  return id;
}

function handleSaveSalonEvenement(p, ss) {
  try {
    const nom = String(p.nom || "").trim();
    if (!nom) return createJsonResponse({ success: false, message: "Le nom de l'événement est obligatoire." });
    const ws = _salonEvtSheet_(ss);
    const id = _salonEvtUpsert_(ws, nom, p.date, p.lieu, String(p.nomAuteur || p.emailAuteur || ""));
    return createJsonResponse({ success: true, id: id, evenements: _salonEvtList_(ss) });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}


// ── Documents de mémoire (versions V1, V2… datées, un dossier Drive par étudiant) ──
var MEMDOC_HEADERS = ["ID", "ID mémoire", "ID étudiant", "Version", "Date", "Nom du fichier", "ID fichier", "URL", "Commentaire", "Déposé par", "Déposé le", "Archivé"];

// Dossier parent imposé par la direction (nov. 2026) : le dossier « CRM HECG — Mémoires étudiants »
// y est créé, ou y est déplacé s'il existait ailleurs (les liens des fichiers restent valides).
var MEMOIRES_PARENT_FOLDER_ID = "1kv1iYmmc0XLc8B1pl4saGxfUqzkP49Xp";

function _memDocsRoot_() {
  const props = PropertiesService.getScriptProperties();
  let rootId = props.getProperty("MEMOIRES_FOLDER_ID"), root = null;
  if (rootId) { try { root = DriveApp.getFolderById(rootId); if (root.isTrashed()) root = null; } catch (e) { root = null; } }
  const parent = DriveApp.getFolderById(MEMOIRES_PARENT_FOLDER_ID);
  if (!root) {
    root = parent.createFolder("CRM HECG — Mémoires étudiants");
    props.setProperty("MEMOIRES_FOLDER_ID", root.getId());
  } else if (props.getProperty("MEMOIRES_FOLDER_PARENT") !== MEMOIRES_PARENT_FOLDER_ID) {
    let dejaDedans = false;
    const it = root.getParents(); while (it.hasNext()) if (it.next().getId() === MEMOIRES_PARENT_FOLDER_ID) dejaDedans = true;
    if (!dejaDedans) {
      root.moveTo(parent);
      try { logAction("Système", "Auto", "Modification", "Vie scolaire — Mémoires", root.getId(), "Dossier Drive des mémoires déplacé dans le dossier " + MEMOIRES_PARENT_FOLDER_ID); } catch (eL) {}
    }
  }
  props.setProperty("MEMOIRES_FOLDER_PARENT", MEMOIRES_PARENT_FOLDER_ID);
  return root;
}

/** À exécuter une fois depuis l'éditeur pour déplacer le dossier tout de suite (sinon : au prochain dépôt). */
function deplacerDossierMemoires() { const r = _memDocsRoot_(); Logger.log("Dossier mémoires : " + r.getUrl()); return r.getUrl(); }

function _memDocsFolder_(etudiantNom) {
  const root = _memDocsRoot_();
  const nom = String(etudiantNom || "Sans nom").replace(/[\\/:*?"<>|]/g, " ").trim();
  const it = root.getFoldersByName(nom);
  return it.hasNext() ? it.next() : root.createFolder(nom);
}

function _memDocsByMemoire_(ss) {
  const out = {};
  const ws = ss.getSheetByName("MEMOIRES_DOCS");
  if (!ws) return out;
  const tz = Session.getScriptTimeZone();
  const d = ws.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (!d[i][0] || String(d[i][11]) === "Oui") continue;
    (out[String(d[i][1])] = out[String(d[i][1])] || []).push({ id: String(d[i][0]), version: String(d[i][3] || ""),
      date: d[i][4] instanceof Date ? Utilities.formatDate(d[i][4], tz, "yyyy-MM-dd") : String(d[i][4] || ""),
      nom: String(d[i][5] || ""), url: String(d[i][7] || ""), commentaire: String(d[i][8] || ""), par: String(d[i][9] || "") });
  }
  const num = function(v) { return parseInt(String(v).replace(/\D/g, ""), 10) || 0; };
  Object.keys(out).forEach(function(k) { out[k].sort(function(a, b) { return num(a.version) - num(b.version) || (a.date < b.date ? -1 : 1); }); });
  return out;
}

/** Dépose un document de mémoire : version auto (V1, V2…) ou saisie, date choisie, renommé proprement. */
function handleUploadMemoireDoc(p, ss) {
  if (!_vsxCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    if (!p.memoireId || !p.data) return createJsonResponse({ success: false, message: "Fichier ou mémoire manquant." });
    const md = _vsxSheet_(ss, "MEMOIRES", MEM_HEADERS).getDataRange().getValues();
    let mem = null;
    for (let i = 1; i < md.length; i++) if (String(md[i][0]) === String(p.memoireId)) { mem = md[i]; break; }
    if (!mem) return createJsonResponse({ success: false, message: "Mémoire introuvable." });
    const ws = _vsxSheet_(ss, "MEMOIRES_DOCS", MEMDOC_HEADERS);
    const existants = (_memDocsByMemoire_(ss)[String(p.memoireId)] || []);
    let version = String(p.version || "").trim().toUpperCase();
    if (!version) {
      const max = existants.reduce(function(m, d) { return Math.max(m, parseInt(String(d.version).replace(/\D/g, ""), 10) || 0); }, 0);
      version = "V" + (max + 1);
    }
    const date = String(p.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"));
    const orig = String(p.fileName || "document");
    const ext = orig.indexOf(".") !== -1 ? orig.slice(orig.lastIndexOf(".")) : "";
    const etu = String(mem[2] || "Étudiant");
    const nomFichier = etu + " — Mémoire " + version + " — " + date + ext;
    const blob = Utilities.newBlob(Utilities.base64Decode(String(p.data)), String(p.mimeType || "application/octet-stream"), nomFichier);
    const file = _memDocsFolder_(etu).createFile(blob);
    const id = "MDOC" + Date.now();
    ws.appendRow([id, String(p.memoireId), String(mem[1] || ""), version, "", nomFichier, file.getId(), file.getUrl(), String(p.commentaire || ""), String(p.nomAuteur || p.emailAuteur || ""), new Date(), ""]);
    ws.getRange(ws.getLastRow(), 5).setNumberFormat("@").setValue(date);
    logAction(p.nomAuteur || p.emailAuteur, p.vieScRole || "", "Création", "Vie scolaire", String(p.memoireId), "Document mémoire " + version + " déposé pour " + etu);
    return createJsonResponse({ success: true, doc: { id: id, version: version, date: date, nom: nomFichier, url: file.getUrl(), commentaire: String(p.commentaire || "") } });
  } catch (e) { return createJsonResponse({ success: false, message: "Dépôt impossible : " + e.toString() }); }
  finally { try { lock.releaseLock(); } catch (eL) {} }
}

/** Retire un document de la liste (le fichier reste dans Drive). */
function handleArchiveMemoireDoc(p, ss) {
  if (!_vsxCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  try {
    const ws = _vsxSheet_(ss, "MEMOIRES_DOCS", MEMDOC_HEADERS);
    const d = ws.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) if (String(d[i][0]) === String(p.id)) { ws.getRange(i + 1, 12).setValue("Oui"); return createJsonResponse({ success: true }); }
    return createJsonResponse({ success: false, message: "Document introuvable." });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}


// ── Carnet d'échanges par mémoire (feuille MEMOIRES_ECHANGES) ──
var MEMECH_HEADERS = ["ID", "ID mémoire", "ID étudiant", "Date", "Type", "Message", "Auteur", "Email auteur", "Créé le", "Archivé"];
var MEMECH_TYPES = ["Échange", "Rendez-vous", "Mail", "Appel", "Relance", "Note"];

function _memEchangesByMemoire_(ss) {
  const out = {};
  const ws = ss.getSheetByName("MEMOIRES_ECHANGES");
  if (!ws) return out;
  const tz = Session.getScriptTimeZone();
  const d = ws.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (!d[i][0] || String(d[i][9]) === "Oui") continue;
    (out[String(d[i][1])] = out[String(d[i][1])] || []).push({ id: String(d[i][0]),
      date: d[i][3] instanceof Date ? Utilities.formatDate(d[i][3], tz, "yyyy-MM-dd") : String(d[i][3] || ""),
      type: String(d[i][4] || "Échange"), texte: String(d[i][5] || ""), auteur: String(d[i][6] || ""), email: String(d[i][7] || ""),
      creeLe: d[i][8] instanceof Date ? Utilities.formatDate(d[i][8], tz, "yyyy-MM-dd HH:mm") : String(d[i][8] || "") });
  }
  // Plus récent en premier
  Object.keys(out).forEach(function(k) { out[k].sort(function(a, b) { return (b.date + b.creeLe).localeCompare(a.date + a.creeLe); }); });
  return out;
}

/** Ajoute une entrée au carnet d'échanges d'un mémoire, ou la retire (p.retirer = id). */
function handleSaveEchangeMemoire(p, ss) {
  if (!_vsxCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const ws = _vsxSheet_(ss, "MEMOIRES_ECHANGES", MEMECH_HEADERS);
    const par = String(p.nomAuteur || p.emailAuteur || "");
    if (p.retirer) {
      const d = ws.getDataRange().getValues();
      for (let i = 1; i < d.length; i++) if (String(d[i][0]) === String(p.retirer)) {
        ws.getRange(i + 1, 10).setValue("Oui");
        logAction(par, p.vieScRole || "", "Suppression", "Vie scolaire — Mémoires", String(d[i][1]), "Carnet d'échanges : entrée du " + (d[i][3] instanceof Date ? Utilities.formatDate(d[i][3], Session.getScriptTimeZone(), "dd/MM/yyyy") : String(d[i][3])) + " retirée");
        return createJsonResponse({ success: true });
      }
      return createJsonResponse({ success: false, message: "Entrée introuvable." });
    }
    const e = p.echange || {};
    const texte = String(e.texte || "").trim();
    if (!p.memoireId || !texte) return createJsonResponse({ success: false, message: "Message vide." });
    const md = _vsxSheet_(ss, "MEMOIRES", MEM_HEADERS).getDataRange().getValues();
    let mem = null;
    for (let i = 1; i < md.length; i++) if (String(md[i][0]) === String(p.memoireId)) { mem = md[i]; break; }
    if (!mem) return createJsonResponse({ success: false, message: "Mémoire introuvable." });
    const tz = Session.getScriptTimeZone(), now = new Date();
    const date = /^\d{4}-\d{2}-\d{2}$/.test(String(e.date || "")) ? String(e.date) : Utilities.formatDate(now, tz, "yyyy-MM-dd");
    const type = MEMECH_TYPES.indexOf(String(e.type)) !== -1 ? String(e.type) : "Échange";
    const id = "MECH" + now.getTime();
    ws.appendRow([id, String(p.memoireId), String(mem[1] || ""), "", type, texte, par, String(p.emailAuteur || ""), now, ""]);
    ws.getRange(ws.getLastRow(), 4).setNumberFormat("@").setValue(date);
    logAction(par, p.vieScRole || "", "Création", "Vie scolaire — Mémoires", String(p.memoireId), "Carnet d'échanges (" + type + ") — " + String(mem[2] || "") + " : " + texte.substring(0, 80));
    return createJsonResponse({ success: true, echange: { id: id, date: date, type: type, texte: texte, auteur: par, email: String(p.emailAuteur || ""), creeLe: Utilities.formatDate(now, tz, "yyyy-MM-dd HH:mm") } });
  } catch (err) { return createJsonResponse({ success: false, message: err.toString() }); }
  finally { try { lock.releaseLock(); } catch (eL) {} }
}


// ── Référentiel des mots-clés de mémoire (feuille MEMOIRES_MOTSCLES) ──
var MEM_KW_DEFAUT = ["Fiscalité", "CDG", "IA", "Dématérialisation", "Comptabilité", "Audit", "Finance", "Paie", "Droit social", "RH", "Marketing", "RSE", "Gestion de trésorerie", "Consolidation", "Digitalisation"];

function _memMotsClesSheet_(ss) {
  let ws = ss.getSheetByName("MEMOIRES_MOTSCLES");
  if (!ws) {
    ws = _vsxSheet_(ss, "MEMOIRES_MOTSCLES", ["Mot-clé", "Créé par", "Créé le", "Retiré"]);
    const now = new Date();
    ws.getRange(2, 1, MEM_KW_DEFAUT.length, 4).setValues(MEM_KW_DEFAUT.map(function(k) { return [k, "Liste initiale", now, ""]; }));
  }
  return ws;
}
function _memMotsCles_(ss) {
  const d = _memMotsClesSheet_(ss).getDataRange().getValues();
  const out = [];
  for (let i = 1; i < d.length; i++) if (d[i][0] && String(d[i][3]) !== "Oui") out.push(String(d[i][0]).trim());
  return out;
}

/** Ajoute / retire des mots-clés du référentiel (retirer = masquer, jamais supprimer). */
function handleSaveMotsClesMemoire(p, ss) {
  if (!_vsxCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  try {
    const ws = _memMotsClesSheet_(ss);
    const d = ws.getDataRange().getValues();
    const norm = function(s) { return String(s || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().toLowerCase(); };
    const idx = {}; for (let i = 1; i < d.length; i++) idx[norm(d[i][0])] = i + 1;
    const ajoutes = [];
    (Array.isArray(p.ajouter) ? p.ajouter : []).forEach(function(k) {
      k = String(k || "").trim().replace(/\s+/g, " ").slice(0, 40); if (!k) return;
      const row = idx[norm(k)];
      if (row) { ws.getRange(row, 4).setValue(""); return; }   // réactive un mot-clé retiré
      ws.appendRow([k, String(p.nomAuteur || p.emailAuteur || ""), new Date(), ""]);
      idx[norm(k)] = ws.getLastRow(); ajoutes.push(k);
    });
    (Array.isArray(p.retirer) ? p.retirer : []).forEach(function(k) { const row = idx[norm(k)]; if (row) ws.getRange(row, 4).setValue("Oui"); });
    if (ajoutes.length) logAction(p.nomAuteur || p.emailAuteur, p.vieScRole || "", "Création", "Vie scolaire", "MOTSCLES", "Mot(s)-clé(s) mémoire : " + ajoutes.join(", "));
    return createJsonResponse({ success: true, motsCles: _memMotsCles_(ss) });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}


// ═══════════════════════════ CHECKLIST MÉMOIRES PAR DIPLÔME ═══════════════════════════
// Par année scolaire (sept. → août) et par diplôme (BTS CG, BTS GPME, DCG, DSCG) :
// on coche les étudiants concernés ; 4 contrôles par étudiant :
//   sujet (mémoire), mémoire inséré (au moins un document), prof référent, dépôt Cyclade (coché à la main).
// Année en cours : liste calculée en direct depuis ETUDIANTS (inscrits / préinscrits).
// Année terminée : figée (instantané JSON) puis consultable en lecture seule.
var CK_HEADERS = ["Année scolaire", "Diplôme", "ID étudiant", "Concerné", "Dépôt Cyclade", "Date dépôt Cyclade", "Instantané (archive)", "Mis à jour le", "Par"];
var CK_DIPLOMES = ["BTS CG", "BTS GPME", "DCG", "DSCG"];

function _ckDiplome_(classe) {
  // Classes compactes du CRM : BTSCG1A, AIX_BTS1GPME, DCG2A, DSCG1A… → on compare les lettres seules
  const c = String(classe || "").toUpperCase().normalize("NFD").replace(/[^A-Z]/g, "");
  if (c.indexOf("DSCG") !== -1) return "DSCG";
  if (c.indexOf("DCG") !== -1) return "DCG";
  if (c.indexOf("GPME") !== -1) return "BTS GPME";
  if (c.indexOf("BTSCG") !== -1 || (c.indexOf("BTS") !== -1 && c.indexOf("CG") !== -1)) return "BTS CG";
  return "";
}
function _ckAnneeCourante_() {
  const d = new Date(), y = d.getFullYear();
  return d.getMonth() >= 8 ? y + "-" + (y + 1) : (y - 1) + "-" + y;
}
function _ckEstTerminee_(annee) {
  const fin = parseInt(String(annee).split("-")[1], 10);
  return !isNaN(fin) && new Date() > new Date(fin, 7, 31, 23, 59);
}

function _ckRows_(ss, annee) {
  const ws = _vsxSheet_(ss, "MEMOIRES_CHECKLIST", CK_HEADERS);
  const d = ws.getDataRange().getValues();
  const rows = {}, annees = {};
  for (let i = 1; i < d.length; i++) {
    if (!d[i][0]) continue;
    annees[String(d[i][0])] = true;
    if (String(d[i][0]) === annee) rows[String(d[i][2])] = { row: i + 1, v: d[i] };
  }
  return { ws: ws, rows: rows, annees: Object.keys(annees) };
}

function _ckBuild_(ss, annee, ck) {
  const tz = Session.getScriptTimeZone();
  const ref = _vsxRef_(ss);
  // Rapprochement avec les mémoires déjà saisis : par ID étudiant, sinon par nom
  // (mémoire saisi sous un ancien ID, ou fiche étudiant recréée) — le plus récent l'emporte.
  const nomKey = function(s) { return String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/[^a-z]+/g, " ").trim().split(" ").sort().join(" "); };
  const md = _vsxSheet_(ss, "MEMOIRES", MEM_HEADERS).getDataRange().getValues();
  const mem = {}, memNom = {};
  const garder = function(map, k, o) { if (k && (!map[k] || o.maj >= map[k].maj)) map[k] = o; };
  for (let i = 1; i < md.length; i++) {
    if (!md[i][0] || String(md[i][14]) === "Oui") continue;
    const maj = md[i][13] instanceof Date ? md[i][13].getTime() : 0;
    const o = { id: String(md[i][0]), sujet: String(md[i][5] || ""), formateurRef: String(md[i][7] || ""), formateur: String(md[i][8] || ""), statut: String(md[i][9] || ""), maj: maj };
    garder(mem, String(md[i][1]), o);
    garder(memNom, nomKey(md[i][2]), o);
  }
  const docs = _memDocsByMemoire_(ss);
  const ech = _memEchangesByMemoire_(ss);
  const byId = {}; ref.etudiants.forEach(function(e) { byId[e.id] = e; });
  // Mails des formateurs / personnel (par réf., sinon par nom)
  const mailRef = {}, mailNom = {};
  ref.formateurs.concat(ref.personnel).forEach(function(x) { if (x.email) { mailRef[x.ref] = x.email; mailNom[nomKey(x.nom)] = mailNom[nomKey(x.nom)] || x.email; } });
  const ids = {};
  ref.etudiants.forEach(function(e) {
    if (e.sortie || (e.typo !== "inscrit" && e.typo !== "préinscrit" && e.typo !== "preinscrit")) return;
    if (_ckDiplome_(e.classe)) ids[e.id] = true;
  });
  Object.keys(ck.rows).forEach(function(id) { if (String(ck.rows[id].v[3]) === "Oui") ids[id] = true; }); // concernés devenus sortis
  const out = {}; CK_DIPLOMES.forEach(function(d) { out[d] = []; });
  Object.keys(ids).forEach(function(id) {
    const e = byId[id] || {};
    const r = ck.rows[id] ? ck.rows[id].v : null;
    const dip = _ckDiplome_(e.classe) || (r ? String(r[1]) : "");
    if (!out[dip]) return;
    let m = mem[id], rapprochement = m ? "id" : "";
    if (!m && (e.prenom || e.nom)) { m = memNom[nomKey((e.prenom || "") + " " + (e.nom || ""))]; if (m) rapprochement = "nom"; }
    const ds = m ? (docs[m.id] || []) : [];
    const last = ds.length ? ds[ds.length - 1] : null;
    const dc = r && r[5] ? (r[5] instanceof Date ? Utilities.formatDate(r[5], tz, "yyyy-MM-dd") : String(r[5])) : "";
    out[dip].push({ etuId: id, etudiant: ((e.prenom || "") + " " + (e.nom || "")).trim() || id, classe: e.classe || "", campus: e.campus || "",
      sorti: !!e.sortie || !byId[id], concerne: !!r && String(r[3]) === "Oui", memoireId: m ? m.id : "", rapprochement: rapprochement,
      mailHecg: e.mailHecg || "", mailPerso: e.mailPerso || "",
      sujet: m ? m.sujet : "", referent: m ? m.formateur : "", statut: m ? m.statut : "",
      referentMail: m ? (mailRef[m.formateurRef] || mailNom[nomKey(m.formateur)] || "") : "", nbDocs: ds.length,
      dernierDoc: last ? { version: last.version, date: last.date, url: last.url } : null,
      nbEchanges: m ? (ech[m.id] || []).length : 0, dernierEchange: m && (ech[m.id] || []).length ? ech[m.id][0].date : "",
      cyclade: !!r && String(r[4]) === "Oui", cycladeDate: dc });
  });
  CK_DIPLOMES.forEach(function(d) { out[d].sort(function(a, b) { return (a.classe + a.etudiant).localeCompare(b.classe + b.etudiant); }); });
  return out;
}

function handleGetMemoiresChecklist(p, ss) {
  if (!_vsxCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  try {
    const courante = _ckAnneeCourante_();
    const annee = String(p.annee || courante);
    const ck = _ckRows_(ss, annee);
    const terminee = _ckEstTerminee_(annee);
    let diplomes;
    const ids = Object.keys(ck.rows);
    const snap = ids.length && ids.every(function(id) { return ck.rows[id].v[6] || String(ck.rows[id].v[3]) !== "Oui"; }) && ids.some(function(id) { return ck.rows[id].v[6]; });
    if (terminee && snap) {
      // Archive : lecture des instantanés figés
      diplomes = {}; CK_DIPLOMES.forEach(function(d) { diplomes[d] = []; });
      ids.forEach(function(id) { const v = ck.rows[id].v; if (!v[6]) return; try { const o = JSON.parse(String(v[6])); (diplomes[o.diplome] = diplomes[o.diplome] || []).push(o); } catch (e) {} });
    } else {
      diplomes = _ckBuild_(ss, annee, ck);
      if (terminee) {
        // Fin d'année : on fige les étudiants concernés (archive consultable)
        const now = new Date();
        Object.keys(diplomes).forEach(function(dip) {
          diplomes[dip].forEach(function(o) {
            if (!o.concerne || !ck.rows[o.etuId]) return;
            o.diplome = dip;
            ck.ws.getRange(ck.rows[o.etuId].row, 7, 1, 2).setValues([[JSON.stringify(o), now]]);
          });
          diplomes[dip] = diplomes[dip].filter(function(o) { return o.concerne; });
        });
      }
    }
    const annees = ck.annees.concat([courante]).filter(function(a, i, arr) { return arr.indexOf(a) === i; }).sort().reverse();
    return createJsonResponse({ success: true, annee: annee, anneeCourante: courante, archivee: terminee, annees: annees, diplomes: diplomes });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

/** Coche / décoche « concerné » (plusieurs étudiants) ou « dépôt Cyclade » (un étudiant). */
function handleSaveMemoiresChecklist(p, ss) {
  if (!_vsxCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const annee = String(p.annee || _ckAnneeCourante_());
    if (_ckEstTerminee_(annee)) return createJsonResponse({ success: false, message: "Année " + annee + " archivée : consultation seule." });
    const ck = _ckRows_(ss, annee);
    const now = new Date(), par = String(p.nomAuteur || p.emailAuteur || "");
    const dE = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues();
    const classe = {}; for (let i = 1; i < dE.length; i++) classe[String(dE[i][0])] = String(dE[i][14] || "");
    const ensure = function(id) {
      if (ck.rows[id]) return ck.rows[id];
      ck.ws.appendRow([annee, _ckDiplome_(classe[id]), id, "", "", "", "", now, par]);
      const r = { row: ck.ws.getLastRow(), v: [annee, _ckDiplome_(classe[id]), id, "", "", "", "", now, par] };
      ck.rows[id] = r; return r;
    };
    (Array.isArray(p.etuIds) ? p.etuIds : []).forEach(function(id) {
      const r = ensure(String(id));
      ck.ws.getRange(r.row, 4).setValue(p.concerne ? "Oui" : "");
      ck.ws.getRange(r.row, 8, 1, 2).setValues([[now, par]]);
    });
    if (p.cycladeEtuId) {
      const r = ensure(String(p.cycladeEtuId));
      ck.ws.getRange(r.row, 5, 1, 2).setValues([[p.cyclade ? "Oui" : "", p.cyclade ? String(p.cycladeDate || Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd")) : ""]]);
      ck.ws.getRange(r.row, 8, 1, 2).setValues([[now, par]]);
    }
    // Traçabilité (Historique)
    const nomEtu = function(id) { for (let i = 1; i < dE.length; i++) if (String(dE[i][0]) === String(id)) return (String(dE[i][3] || "") + " " + String(dE[i][2] || "")).trim(); return String(id); };
    const ids = Array.isArray(p.etuIds) ? p.etuIds : [];
    if (ids.length) logAction(par, p.vieScRole || "", p.concerne ? "Création" : "Suppression", "Vie scolaire — Mémoires", ids.length === 1 ? String(ids[0]) : "Checklist " + annee,
      (p.concerne ? "Checklist mémoires " + annee + " : ajout de " : "Checklist mémoires " + annee + " : retrait de ") + (ids.length > 5 ? ids.length + " étudiants" : ids.map(nomEtu).join(", ")));
    if (p.cycladeEtuId) logAction(par, p.vieScRole || "", "Modification", "Vie scolaire — Mémoires", String(p.cycladeEtuId),
      "Dépôt Cyclade " + (p.cyclade ? "enregistré" : "annulé") + " pour " + nomEtu(p.cycladeEtuId) + " (" + annee + ")");
    return createJsonResponse({ success: true });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
  finally { try { lock.releaseLock(); } catch (eL) {} }
}
