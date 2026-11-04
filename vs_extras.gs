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
                     campus: _vsxCampusLib_(dE[i][22]), sortie: typo === "sortie" });
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
      formateurs.push({ ref: "A:" + String(dA[r][0]), nom: (String(dA[r][2] || "") + " " + String(dA[r][1] || "").toUpperCase()).trim(), matieres: String(dA[r][8] || "") });
    }
  } catch (eF) { Logger.log("Formateurs : " + eF); }
  formateurs.sort(function(a, b) { return a.nom.localeCompare(b.nom); });
  return { etudiants: etudiants, personnel: personnel, formateurs: formateurs };
}

// ═══════════════════════════ MÉMOIRES ═══════════════════════════
var MEM_HEADERS = ["ID", "ID étudiant", "Étudiant", "Classe", "Campus", "Sujet", "Mots-clés", "Réf. formateur", "Formateur référent", "Statut", "Notes", "Créé le", "Créé par", "Mis à jour le", "Archivé"];

function handleGetMemoires(p, ss) {
  if (!_vsxCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  return getCachedJsonResponse_('memoires_v1', 120, function() {
    try {
      const ref = _vsxRef_(ss);
      const byId = {}; ref.etudiants.forEach(function(e) { byId[e.id] = e; });
      const d = _vsxSheet_(ss, "MEMOIRES", MEM_HEADERS).getDataRange().getValues();
      const list = [];
      for (let i = 1; i < d.length; i++) {
        if (!d[i][0] || String(d[i][14]) === "Oui") continue;
        const e = byId[String(d[i][1])];
        list.push({ id: String(d[i][0]), etuId: String(d[i][1]), etudiant: e ? (e.prenom + " " + e.nom).trim() : String(d[i][2] || ""),
          classe: e ? e.classe : String(d[i][3] || ""), campus: e ? e.campus : String(d[i][4] || ""), sujet: String(d[i][5] || ""),
          motsCles: String(d[i][6] || "").split(",").map(function(s) { return s.trim(); }).filter(Boolean),
          formateurRef: String(d[i][7] || ""), formateur: String(d[i][8] || ""), statut: String(d[i][9] || "En cours"), notes: String(d[i][10] || ""),
          creeLe: d[i][11] instanceof Date ? Utilities.formatDate(d[i][11], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(d[i][11] || ""),
          majLe: d[i][13] instanceof Date ? Utilities.formatDate(d[i][13], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(d[i][13] || "") });
      }
      const docs = _memDocsByMemoire_(ss);
      list.forEach(function(m) { m.docs = docs[m.id] || []; });
      return createJsonResponse({ success: true, data: list, etudiants: ref.etudiants.filter(function(e) { return !e.sortie; }), personnel: ref.personnel, formateurs: ref.formateurs });
    } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
  });
}

function handleSaveMemoire(p, ss) {
  if (!_vsxCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  try {
    const ws = _vsxSheet_(ss, "MEMOIRES", MEM_HEADERS);
    const m = p.memoire || {};
    if (!m.etuId) return createJsonResponse({ success: false, message: "Choisissez un étudiant." });
    const ref = _vsxRef_(ss);
    const e = ref.etudiants.filter(function(x) { return x.id === String(m.etuId); })[0] || {};
    const f = ref.formateurs.concat(ref.personnel).filter(function(x) { return x.ref === String(m.formateurRef); })[0] || {};
    const mots = (Array.isArray(m.motsCles) ? m.motsCles : String(m.motsCles || "").split(",")).map(function(s) { return String(s).trim(); }).filter(Boolean).join(", ");
    const now = new Date();
    const rowVals = [m.id || ("MEM" + now.getTime()), String(m.etuId), ((e.prenom || "") + " " + (e.nom || "")).trim(), e.classe || "", e.campus || "",
      String(m.sujet || ""), mots, f.nom ? String(m.formateurRef || "") : "", f.nom || String(m.formateur || "").trim(), String(m.statut || "En cours"), String(m.notes || ""), now,
      String(p.nomAuteur || p.emailAuteur || ""), now, ""];
    if (m.id) {
      const d = ws.getDataRange().getValues();
      for (let i = 1; i < d.length; i++) if (String(d[i][0]) === String(m.id)) {
        rowVals[11] = d[i][11]; rowVals[12] = d[i][12];
        if (m.archiver) rowVals[14] = "Oui";
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
var FRISE_HEADERS = ["ID", "Titre", "Description", "Couleur", "Destinataires", "Rappels (jours avant)", "Créé par", "Email créateur", "Créé le", "Archivée"];
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
        creeLe: fd[i][8] instanceof Date ? Utilities.formatDate(fd[i][8], tz, "yyyy-MM-dd") : "", etapes: et });
    }
    return createJsonResponse({ success: true, data: list });
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
    const dest = String(f.destinataires || me).toLowerCase().split(/[,;\s]+/).filter(function(s) { return s.indexOf("@") > 0; }).join(", ");
    let id = f.id;
    if (id) {
      const d = fws.getDataRange().getValues(); let found = false;
      for (let i = 1; i < d.length; i++) if (String(d[i][0]) === String(id)) {
        if (!_vsxIsSuper_(p) && String(d[i][7]).toLowerCase() !== me) return createJsonResponse({ success: false, message: "Seul le créateur peut modifier cette frise." });
        fws.getRange(i + 1, 2, 1, 5).setValues([[String(f.titre || ""), String(f.description || ""), String(f.couleur || "#1A4E8A"), dest, rappels]]);
        if (f.archiver) fws.getRange(i + 1, 10).setValue("Oui");
        found = true; break;
      }
      if (!found) return createJsonResponse({ success: false, message: "Frise introuvable." });
    } else {
      if (!String(f.titre || "").trim()) return createJsonResponse({ success: false, message: "Donnez un titre à la frise." });
      id = "FR" + now.getTime();
      fws.appendRow([id, String(f.titre), String(f.description || ""), String(f.couleur || "#1A4E8A"), dest, rappels, String(p.nomAuteur || me), me, now, ""]);
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
    logAction(p.nomAuteur || me, p.vieScRole || "", f.id ? (f.archiver ? "Suppression" : "Modification") : "Création", "Vie scolaire", id, "Frise : " + String(f.titre || ""));
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
    const html = '<div style="font-family:Segoe UI,Arial,sans-serif;background:#f5f7fb;padding:28px"><div style="max-width:580px;margin:0 auto;background:#fff;border-radius:16px;border:1px solid #e6ebf2;overflow:hidden">'
      + '<div style="background:#1A4E8A;color:#fff;padding:20px 26px;font-size:18px;font-weight:700">⏰ Prochaines échéances</div>'
      + '<div style="padding:8px 26px 24px"><table style="width:100%;border-collapse:collapse">' + lignes + '</table>'
      + '<a href="https://mokhtarimehdipro-bit.github.io/CRMHECG/vie_scolaire.html" style="display:inline-block;margin-top:20px;background:#F26522;color:#fff;text-decoration:none;font-weight:700;padding:12px 22px;border-radius:10px">Ouvrir les frises</a></div></div></div>';
    try { MailApp.sendEmail({ to: em, subject: "[HECG] ⏰ " + (items[0].jours === 0 ? "Échéance aujourd'hui" : "Échéance proche") + " : " + items[0].et[3], htmlBody: html, name: "CRM HECG" }); }
    catch (e) { Logger.log("Rappel frise " + em + " : " + e); }
  });
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
    const res = [], rows = [];
    let base = Date.now();
    items.forEach(function(it) {
      const cid = String(it.cid || "");
      const seen = cid ? cache.get("salon_cid_" + cid) : null;
      if (seen) { res.push({ cid: cid, statut: "deja", id: seen }); return; }
      const nom = String(it.nom || "").trim().toUpperCase(), prenom = String(it.prenom || "").trim();
      if (!nom && !prenom) { res.push({ cid: cid, statut: "erreur", message: "Nom manquant" }); return; }
      const mail = String(it.mail || "").trim().toLowerCase(), tel = String(it.tel || "").trim();
      const dup = (mail && byMail[mail]) || (normTel(tel).length === 9 && byTel[normTel(tel)]);
      if (dup) { if (cid) cache.put("salon_cid_" + cid, dup, 21600); res.push({ cid: cid, statut: "existant", id: dup }); return; }
      const id = "ETU" + String(base++).slice(-6);
      const row = [id, "P@ss123", nom, prenom, mail, tel, "", "", "", "", "", "", "", "", String(it.classe || ""), "Non", "Non",
                   "", "", "", "PROSPECT", "", String(it.campus || ""), "", "", "", ""];
      rows.push(_withDateCreation_(ws, row));
      if (mail) byMail[mail] = id; if (normTel(tel).length === 9) byTel[normTel(tel)] = id;
      if (cid) cache.put("salon_cid_" + cid, id, 21600);
      res.push({ cid: cid, statut: "ajoute", id: id });
      logAction(p.nomAuteur || p.idAuteur || p.emailAuteur, p.roleAuteur || "", "Création", "Dossier étudiant", id,
        "Prospect salon" + (it.salon ? " « " + String(it.salon) + " »" : "") + " : " + prenom + " " + nom);
    });
    if (rows.length) {
      const width = Math.max.apply(null, rows.map(function(r) { return r.length; }));
      const norm = rows.map(function(r) { while (r.length < width) r.push(""); return r; });
      ws.getRange(ws.getLastRow() + 1, 1, norm.length, width).setValues(norm);
    }
    return createJsonResponse({ success: true, resultats: res });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
  finally { try { lock.releaseLock(); } catch (eL) {} }
}

function handleGetSalonRef(p, ss) {
  return getCachedJsonResponse_('salon_ref_v1', 600, function() {
    try {
      const d = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues();
      const cl = {};
      for (let i = 1; i < d.length; i++) { const c = String(d[i][14] || "").trim(); if (c) cl[c] = true; }
      return createJsonResponse({ success: true, classes: Object.keys(cl).sort() });
    } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
  });
}


// ── Documents de mémoire (versions V1, V2… datées, un dossier Drive par étudiant) ──
var MEMDOC_HEADERS = ["ID", "ID mémoire", "ID étudiant", "Version", "Date", "Nom du fichier", "ID fichier", "URL", "Commentaire", "Déposé par", "Déposé le", "Archivé"];

function _memDocsFolder_(etudiantNom) {
  const props = PropertiesService.getScriptProperties();
  let rootId = props.getProperty("MEMOIRES_FOLDER_ID"), root = null;
  if (rootId) { try { root = DriveApp.getFolderById(rootId); } catch (e) { root = null; } }
  if (!root) { root = DriveApp.createFolder("CRM HECG — Mémoires étudiants"); props.setProperty("MEMOIRES_FOLDER_ID", root.getId()); }
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
