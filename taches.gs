// ==========================================
// TÂCHES DE L'ÉQUIPE (page missions.html — refonte nov. 2026)
// ==========================================
// Les tâches sont les lignes de la feuille MISSIONS (missions Gmail, missions manuelles,
// missions hebdo du planning) : rien n'est dupliqué, les trophées « Missions » continuent
// de compter les tâches clôturées. Colonnes ajoutées à droite :
//   O Échéance (yyyy-MM-dd) · P Urgent (Oui/…) · Q Catégorie (ID) · R Type mandat
//   S ID mandat · T Libellé mandat · U Commentaire · V Terminée par · W Terminée le
// Feuilles annexes : TACHES_CATEGORIES (ID, Nom, Couleur, Ordre) et TACHES_PARAMS (Clé, Valeur).

var TACHE_COL = { ID:0, DATE:1, HEURE:2, EMAIL:3, COLLAB:4, CC:5, OBJET:6, CORPS:7, STATUT:8, NOTES:9, MAJ:10, PRIO:11, RESUME:12,
                  CREATEUR:13, ECHEANCE:14, URGENT:15, CATEGORIE:16, MANDAT_TYPE:17, MANDAT_ID:18, MANDAT_LABEL:19, COMMENTAIRE:20,
                  FAIT_PAR:21, FAIT_LE:22, ARCHIVE:23 };
var TACHE_EXTRA_HEADERS = ["Échéance", "Urgent", "Catégorie", "Type mandat", "ID mandat", "Mandat lié", "Commentaire", "Terminée par", "Terminée le", "Archivée le"];
var TACHE_DEFAULT_CATS = [
  ["CAT1", "Suivi étudiant", "#1A4E8A", 1], ["CAT2", "Entreprises", "#6d28d9", 2], ["CAT3", "Administratif", "#0f766e", 3],
  ["CAT4", "Communication", "#F26522", 4], ["CAT5", "Pédagogie", "#b45309", 5]
];

function _tacheIsAdmin_(p) {
  return String(p.roleAuteurValide || p.roleAuteur || "") === "Administrateur" || String(p.emailAuteur || "").toLowerCase() === "m.mokhtari@hecg.fr";
}

function _tacheSheet_(ss) {
  const ws = getSheetSafe(ss, "MISSIONS");
  // En-têtes des nouvelles colonnes (une seule fois)
  if (ws.getMaxColumns() < 24) ws.insertColumnsAfter(ws.getMaxColumns(), 24 - ws.getMaxColumns());
  const h = ws.getRange(1, 15, 1, TACHE_EXTRA_HEADERS.length).getValues()[0];
  if (String(h[0] || "") !== TACHE_EXTRA_HEADERS[0] || String(h[TACHE_EXTRA_HEADERS.length - 1] || "") !== TACHE_EXTRA_HEADERS[TACHE_EXTRA_HEADERS.length - 1])
    ws.getRange(1, 15, 1, TACHE_EXTRA_HEADERS.length).setValues([TACHE_EXTRA_HEADERS]).setFontWeight("bold");
  return ws;
}

function _tacheCatSheet_(ss) {
  let ws = ss.getSheetByName("TACHES_CATEGORIES");
  if (!ws) {
    ws = ss.insertSheet("TACHES_CATEGORIES");
    ws.appendRow(["ID", "Nom", "Couleur", "Ordre"]);
    ws.getRange(2, 1, TACHE_DEFAULT_CATS.length, 4).setValues(TACHE_DEFAULT_CATS);
    ws.setFrozenRows(1);
  }
  return ws;
}

function _tacheParams_(ss) {
  let ws = ss.getSheetByName("TACHES_PARAMS");
  if (!ws) { ws = ss.insertSheet("TACHES_PARAMS"); ws.appendRow(["Clé", "Valeur"]); ws.appendRow(["notifMail", "Oui"]); }
  const d = ws.getDataRange().getValues();
  const out = { ws: ws, values: {} };
  for (let i = 1; i < d.length; i++) if (d[i][0]) out.values[String(d[i][0])] = String(d[i][1] || "");
  return out;
}

function _tacheFmtDate_(v) {
  if (!v) return "";
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  const s = String(v).trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  return m ? m[3] + "-" + ("0" + m[2]).slice(-2) + "-" + ("0" + m[1]).slice(-2) : s.slice(0, 10);
}

function _tacheRowToObj_(r) {
  const C = TACHE_COL;
  return {
    id: String(r[C.ID]), date: _tacheFmtDate_(r[C.DATE]), email: String(r[C.EMAIL] || "").toLowerCase(), collab: String(r[C.COLLAB] || ""),
    titre: String(r[C.OBJET] || ""), detail: String(r[C.CORPS] || ""), statut: String(r[C.STATUT] || "Attribué"),
    notes: String(r[C.NOTES] || ""), priorite: Number(r[C.PRIO] || 3), createur: String(r[C.CREATEUR] || "").toLowerCase(),
    echeance: _tacheFmtDate_(r[C.ECHEANCE]), urgent: String(r[C.URGENT] || "") === "Oui",
    categorie: String(r[C.CATEGORIE] || ""), mandatType: String(r[C.MANDAT_TYPE] || ""), mandatId: String(r[C.MANDAT_ID] || ""),
    mandatLabel: String(r[C.MANDAT_LABEL] || ""), commentaire: String(r[C.COMMENTAIRE] || ""),
    faitPar: String(r[C.FAIT_PAR] || ""), faitLe: _tacheFmtDate_(r[C.FAIT_LE])
  };
}

// ── Lecture ───────────────────────────────────────────────────────
function handleGetTaches(p, ss) {
  const isAdmin = _tacheIsAdmin_(p);
  const me = String(p.emailAuteur || "").toLowerCase();
  return getEntityCachedJsonResponse_('taches_v1', isAdmin ? 'admin' : me, 60, function() {
    try {
      const ws = _tacheSheet_(ss);
      const d = ws.getDataRange().getValues();
      const list = [];
      for (let i = 1; i < d.length; i++) {
        if (!d[i][0]) continue;
        const st = String(d[i][TACHE_COL.STATUT] || "");
        if (st === "Annulé" || d[i][TACHE_COL.ARCHIVE]) continue;
        if (!isAdmin && String(d[i][TACHE_COL.EMAIL] || "").toLowerCase() !== me) continue;
        list.push(_tacheRowToObj_(d[i]));
      }
      return createJsonResponse({ success: true, data: list, isAdmin: isAdmin });
    } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
  });
}

// Référentiels : équipe, catégories, paramètres + index des « mandats » (admin)
function handleGetTachesRef(p, ss) {
  const isAdmin = _tacheIsAdmin_(p);
  return getCachedJsonResponse_('taches_ref_v1_' + (isAdmin ? 'a' : 'u'), 300, function() {
    try {
      const dP = getPersonnelDataCached_(ss);
      const equipe = [];
      for (let r = 1; r < dP.length; r++) {
        const email = String(dP[r][6] || "").trim().toLowerCase();
        if (!dP[r][0]) continue;
        equipe.push({ ref: String(dP[r][0]), nom: (String(dP[r][2] || "") + " " + String(dP[r][1] || "")).trim(), email: email || ("ref:" + String(dP[r][0]).toLowerCase()), role: String(dP[r][5] || ""), sansEmail: !email });
      }
      const cd = _tacheCatSheet_(ss).getDataRange().getValues();
      const cats = [];
      for (let i = 1; i < cd.length; i++) if (cd[i][0]) cats.push({ id: String(cd[i][0]), nom: String(cd[i][1] || ""), couleur: String(cd[i][2] || "#64748b"), ordre: Number(cd[i][3] || i) });
      cats.sort(function(a, b) { return a.ordre - b.ordre; });
      const params = _tacheParams_(ss).values;

      let mandats = [];
      if (isAdmin) {
        // Format compact : [type, id, libellé, précision]
        const dE = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues();
        const classes = {};
        for (let k = 1; k < dE.length; k++) {
          const id = String(dE[k][0] || "").trim(); if (!id) continue;
          const lib = (String(dE[k][3] || "") + " " + String(dE[k][2] || "")).trim();
          const cl = String(dE[k][14] || "").trim();
          if (cl) classes[cl] = true;
          mandats.push(["etudiant", id, lib, cl + (String(dE[k][20] || "").toLowerCase() === "sortie" ? " · sortie" : "")]);
        }
        Object.keys(classes).sort().forEach(function(c) { mandats.push(["classe", c, c, "Classe"]); });
        const dPa = getSheetSafe(ss, "PARTENARIAT").getDataRange().getValues();
        const ents = {};
        for (let k = 1; k < dPa.length; k++) {
          const idE = String(dPa[k][1] || "").trim().toUpperCase(), nomE = String(dPa[k][2] || "").trim();
          if (idE && nomE && !ents[idE]) { ents[idE] = true; mandats.push(["entreprise", idE, nomE, String(dPa[k][12] || "")]); }
          const idC = String(dPa[k][0] || "").trim();
          const libC = (String(dPa[k][4] || "") + " " + String(dPa[k][3] || "")).trim();
          if (idC && libC) mandats.push(["contact", idC, libC, nomE]);
        }
        equipe.forEach(function(e) { mandats.push(["personnel", e.ref, e.nom, e.role]); });
        try {
          const dO = getSheetSafe(ss, "OFFRES").getDataRange().getValues();
          for (let k = Math.max(1, dO.length - 400); k < dO.length; k++) {
            if (String(dO[k][9] || "").toLowerCase() === "obsolète") continue;
            if (dO[k][0]) mandats.push(["offre", String(dO[k][0]), String(dO[k][14] || "Offre"), String(dO[k][4] || "")]);
          }
        } catch (eO) {}
      }
      return createJsonResponse({ success: true, equipe: equipe, categories: cats, params: params, mandats: mandats, isAdmin: isAdmin });
    } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
  });
}

// ── Écriture ──────────────────────────────────────────────────────
function _tacheInvalidate_(email) {
  bumpEntityCacheVersion_('taches_v1', 'admin');
  if (email) bumpEntityCacheVersion_('taches_v1', String(email).toLowerCase());
  try { _invalidateMissionsCache_(email); } catch (e) {}
}

/**
 * Création (admin) ou mise à jour partielle d'une tâche.
 * p.id vide → création ; sinon seuls les champs présents dans p.champs sont modifiés.
 * Titulaire : peut cocher « terminée » et écrire un commentaire. Admin : tout.
 */
function handleSaveTache(p, ss) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const isAdmin = _tacheIsAdmin_(p);
    const me = String(p.emailAuteur || "").toLowerCase();
    const ws = _tacheSheet_(ss);
    const tz = Session.getScriptTimeZone();
    const now = new Date();
    const c = p.champs || {};

    if (!p.id) {
      if (!isAdmin) return createJsonResponse({ success: false, message: "Seul l'administrateur attribue des tâches." });
      const titre = String(c.titre || "").trim();
      if (!titre) return createJsonResponse({ success: false, message: "Le nom de la tâche est obligatoire." });
      const email = String(c.email || me).trim().toLowerCase();
      const id = "T" + now.getTime();
      const row = new Array(23).fill("");
      row[TACHE_COL.ID] = id;
      row[TACHE_COL.DATE] = Utilities.formatDate(now, tz, "dd/MM/yyyy");
      row[TACHE_COL.HEURE] = Utilities.formatDate(now, tz, "HH:mm");
      row[TACHE_COL.EMAIL] = email;
      row[TACHE_COL.COLLAB] = String(c.collab || "");
      row[TACHE_COL.OBJET] = titre;
      row[TACHE_COL.CORPS] = String(c.detail || "");
      row[TACHE_COL.STATUT] = "Attribué";
      row[TACHE_COL.MAJ] = now;
      row[TACHE_COL.PRIO] = c.urgent ? 5 : 3;
      row[TACHE_COL.CREATEUR] = me;
      row[TACHE_COL.ECHEANCE] = String(c.echeance || "");
      row[TACHE_COL.URGENT] = c.urgent ? "Oui" : "";
      row[TACHE_COL.CATEGORIE] = String(c.categorie || "");
      row[TACHE_COL.MANDAT_TYPE] = String(c.mandatType || "");
      row[TACHE_COL.MANDAT_ID] = String(c.mandatId || "");
      row[TACHE_COL.MANDAT_LABEL] = String(c.mandatLabel || "");
      row[TACHE_COL.COMMENTAIRE] = String(c.commentaire || "");
      const nextRow = ws.getLastRow() + 1;
      ws.getRange(nextRow, TACHE_COL.ECHEANCE + 1).setNumberFormat("@");
      ws.getRange(nextRow, 1, 1, 23).setValues([row]);
      logAction(p.idAuteur || p.nomAuteur, p.roleAuteur, "Création", "Missions", id, "Tâche : " + titre.substring(0, 80) + " | Pour : " + (c.collab || email));
      _tacheInvalidate_(email);
      _ensureRelanceTachesTrigger_();
      let mail = false;
      if (_tacheParams_(ss).values.notifMail !== "Non" && email && email.indexOf("@") > 0 && email !== me) mail = _tacheNotifier_(email, row, ss);
      return createJsonResponse({ success: true, id: id, tache: _tacheRowToObj_(row), mailEnvoye: mail });
    }

    const d = ws.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][0]) !== String(p.id)) continue;
      const r = d[i].slice(); while (r.length < 23) r.push("");
      const titulaire = String(r[TACHE_COL.EMAIL] || "").toLowerCase();
      const isTitulaire = titulaire === me;
      if (!isAdmin && !isTitulaire) return createJsonResponse({ success: false, message: "Accès refusé." });
      const allowed = isAdmin ? ["titre", "detail", "email", "collab", "echeance", "urgent", "categorie", "mandatType", "mandatId", "mandatLabel", "commentaire", "fait", "supprimer"]
                              : ["fait", "commentaire"];
      const ancienEmail = titulaire;
      let reassigne = false;
      Object.keys(c).forEach(function(k) {
        if (allowed.indexOf(k) === -1) return;
        const v = c[k];
        if (k === "titre") r[TACHE_COL.OBJET] = String(v);
        else if (k === "detail") r[TACHE_COL.CORPS] = String(v);
        else if (k === "email") { const ne = String(v).toLowerCase(); if (ne && ne !== ancienEmail) { r[TACHE_COL.EMAIL] = ne; reassigne = true; } }
        else if (k === "collab") r[TACHE_COL.COLLAB] = String(v);
        else if (k === "echeance") r[TACHE_COL.ECHEANCE] = String(v || "");
        else if (k === "urgent") { r[TACHE_COL.URGENT] = v ? "Oui" : ""; r[TACHE_COL.PRIO] = v ? 5 : 3; }
        else if (k === "categorie") r[TACHE_COL.CATEGORIE] = String(v || "");
        else if (k === "mandatType") r[TACHE_COL.MANDAT_TYPE] = String(v || "");
        else if (k === "mandatId") r[TACHE_COL.MANDAT_ID] = String(v || "");
        else if (k === "mandatLabel") r[TACHE_COL.MANDAT_LABEL] = String(v || "");
        else if (k === "commentaire") r[TACHE_COL.COMMENTAIRE] = String(v || "");
        else if (k === "fait") {
          r[TACHE_COL.STATUT] = v ? "Clôturé" : "Attribué";
          r[TACHE_COL.FAIT_PAR] = v ? String(p.nomAuteur || p.idAuteur || me) : "";
          r[TACHE_COL.FAIT_LE] = v ? Utilities.formatDate(now, tz, "yyyy-MM-dd") : "";
        }
        else if (k === "supprimer" && v) r[TACHE_COL.STATUT] = "Annulé";
      });
      r[TACHE_COL.MAJ] = now;
      ws.getRange(i + 1, TACHE_COL.ECHEANCE + 1).setNumberFormat("@");
      ws.getRange(i + 1, 1, 1, 23).setValues([r.slice(0, 23)]);
      const resume = Object.keys(c).filter(function(k) { return allowed.indexOf(k) !== -1; }).join(", ");
      logAction(p.idAuteur || p.nomAuteur, p.roleAuteur, "Modification", "Missions", p.id,
        (c.fait === true ? "Tâche terminée (clôturé) : " : (c.supprimer ? "Tâche supprimée : " : "Tâche modifiée (" + resume + ") : ")) + String(r[TACHE_COL.OBJET]).substring(0, 70));
      // Sync planning : tâche terminée → tâche planning liée terminée
      if (c.fait === true) {
        try {
          const planWs = ss.getSheetByName("PLANNING");
          if (planWs) {
            const pd = planWs.getDataRange().getValues();
            for (let k = 1; k < pd.length; k++) if (String(pd[k][12] || "") === String(p.id)) { planWs.getRange(k + 1, 6).setValue("Terminé"); break; }
          }
        } catch (eP) {}
      }
      _tacheInvalidate_(ancienEmail);
      if (reassigne) _tacheInvalidate_(r[TACHE_COL.EMAIL]);
      let mail = false;
      if (reassigne && _tacheParams_(ss).values.notifMail !== "Non" && r[TACHE_COL.EMAIL] !== me) mail = _tacheNotifier_(r[TACHE_COL.EMAIL], r, ss);
      return createJsonResponse({ success: true, tache: _tacheRowToObj_(r), mailEnvoye: mail });
    }
    return createJsonResponse({ success: false, message: "Tâche introuvable." });
  } catch (e) {
    return createJsonResponse({ success: false, message: e.toString() });
  } finally { try { lock.releaseLock(); } catch (eL) {} }
}

function _tacheNotifier_(email, r, ss) {
  try {
    if (!email || email.indexOf("@") === -1) return false;
    const tz = Session.getScriptTimeZone();
    let catNom = "";
    try {
      const cd = _tacheCatSheet_(ss).getDataRange().getValues();
      for (let i = 1; i < cd.length; i++) if (String(cd[i][0]) === String(r[TACHE_COL.CATEGORIE])) catNom = String(cd[i][1]);
    } catch (e) {}
    const ech = r[TACHE_COL.ECHEANCE] ? Utilities.formatDate(new Date(String(r[TACHE_COL.ECHEANCE]) + "T12:00:00"), tz, "EEEE d MMMM yyyy") : "";
    const esc = function(s) { return String(s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); };
    const line = function(lib, val) { return val ? '<tr><td style="padding:6px 0;color:#64748b;font-size:13px;width:120px">' + lib + '</td><td style="padding:6px 0;color:#0f1b2d;font-size:14px;font-weight:600">' + val + '</td></tr>' : ''; };
    const url = "https://mokhtarimehdipro-bit.github.io/CRMHECG/missions.html";
    const html = '<div style="font-family:Segoe UI,Arial,sans-serif;background:#f5f7fb;padding:28px">'
      + '<div style="max-width:560px;margin:0 auto;background:#fff;border-radius:16px;overflow:hidden;border:1px solid #e6ebf2">'
      + '<div style="background:#1A4E8A;padding:20px 26px;color:#fff;font-size:18px;font-weight:700">' + (r[TACHE_COL.URGENT] === "Oui" ? "🚩 Nouvelle tâche urgente" : "Nouvelle tâche") + '</div>'
      + '<div style="padding:24px 26px"><p style="margin:0 0 6px;color:#64748b;font-size:13px">Bonjour ' + esc(r[TACHE_COL.COLLAB]) + ', une tâche vous a été attribuée :</p>'
      + '<p style="margin:0 0 16px;font-size:19px;font-weight:700;color:#0f1b2d">' + esc(r[TACHE_COL.OBJET]) + '</p>'
      + (r[TACHE_COL.CORPS] ? '<p style="margin:0 0 16px;color:#334155;font-size:14px;line-height:1.55;white-space:pre-line">' + esc(r[TACHE_COL.CORPS]) + '</p>' : '')
      + '<table style="border-collapse:collapse;width:100%">' + line("Échéance", esc(ech)) + line("Catégorie", esc(catNom)) + line("Mandat lié", esc(r[TACHE_COL.MANDAT_LABEL])) + line("Commentaire", esc(r[TACHE_COL.COMMENTAIRE])) + '</table>'
      + '<a href="' + url + '" style="display:inline-block;margin-top:22px;background:#F26522;color:#fff;text-decoration:none;font-weight:700;padding:12px 22px;border-radius:10px">Voir mes tâches</a>'
      + '</div></div></div>';
    const mailOpts = { to: email, subject: "[HECG] " + (r[TACHE_COL.URGENT] === "Oui" ? "🚩 " : "") + "Nouvelle tâche : " + String(r[TACHE_COL.OBJET]).substring(0, 90), htmlBody: html, name: "CRM HECG" };
    if (email.toLowerCase() !== TACHE_SUPERVISEUR) mailOpts.cc = TACHE_SUPERVISEUR; // direction toujours en copie
    MailApp.sendEmail(mailOpts);
    _ensureRelanceTachesTrigger_();
    return true;
  } catch (e) { Logger.log("Notif tâche : " + e); return false; }
}

// ── Relance automatique : toutes les 2 h, de 7 h à 18 h, du lundi au vendredi ──
// Chaque membre ayant des tâches du jour ou en retard (ou urgentes sans échéance) reçoit
// un récapitulatif, direction en copie. Désactivable : TACHES_PARAMS « relanceAuto » = Non.
var TACHE_SUPERVISEUR = "m.mokhtari@hecg.fr";

function _ensureRelanceTachesTrigger_() {
  try {
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty("RELANCE_TACHES_TRIGGER") === "1") return;
    const exists = ScriptApp.getProjectTriggers().some(function(t) { return t.getHandlerFunction() === "relancerTachesEquipe"; });
    if (!exists) ScriptApp.newTrigger("relancerTachesEquipe").timeBased().everyHours(2).create();
    props.setProperty("RELANCE_TACHES_TRIGGER", "1");
  } catch (e) { Logger.log("Trigger relance tâches : " + e); }
}
/** À exécuter une fois depuis l'éditeur si besoin (le trigger est aussi créé automatiquement). */
function creerTriggerRelanceTaches() {
  ScriptApp.getProjectTriggers().forEach(function(t) { if (t.getHandlerFunction() === "relancerTachesEquipe") ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("relancerTachesEquipe").timeBased().everyHours(2).create();
  PropertiesService.getScriptProperties().setProperty("RELANCE_TACHES_TRIGGER", "1");
}

function relancerTachesEquipe(force) {
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const heure = parseInt(Utilities.formatDate(now, tz, "H"), 10);
  const jourSem = parseInt(Utilities.formatDate(now, tz, "u"), 10); // 1 = lundi … 7 = dimanche
  if (force !== true && (heure < 7 || heure > 18 || jourSem > 5)) return;
  const ss = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");
  if (_tacheParams_(ss).values.relanceAuto === "Non") return;
  const today = Utilities.formatDate(now, tz, "yyyy-MM-dd");
  const d = _tacheSheet_(ss).getDataRange().getValues();
  const parPersonne = {};
  for (let i = 1; i < d.length; i++) {
    if (!d[i][0] || d[i][TACHE_COL.ARCHIVE]) continue;
    const st = String(d[i][TACHE_COL.STATUT] || "");
    if (st === "Clôturé" || st === "Annulé") continue;
    const t = _tacheRowToObj_(d[i]);
    if (!t.email || t.email.indexOf("@") === -1) continue;
    const etat = t.echeance ? (t.echeance < today ? "retard" : (t.echeance === today ? "jour" : "")) : (t.urgent ? "urgent" : "");
    if (!etat) continue;
    t.etat = etat;
    t.joursRetard = etat === "retard" ? Math.round((new Date(today + "T12:00:00") - new Date(t.echeance + "T12:00:00")) / 86400000) : 0;
    (parPersonne[t.email] = parPersonne[t.email] || []).push(t);
  }
  const esc = function(s) { return String(s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); };
  const url = "https://mokhtarimehdipro-bit.github.io/CRMHECG/missions.html";
  let nbMails = 0;
  Object.keys(parPersonne).forEach(function(email) {
    const list = parPersonne[email].sort(function(a, b) { return (b.joursRetard - a.joursRetard) || ((b.urgent ? 1 : 0) - (a.urgent ? 1 : 0)); });
    const retard = list.filter(function(t) { return t.etat === "retard"; });
    const nom = list[0].collab || email;
    const ligne = function(t) {
      const badge = t.etat === "retard" ? '<span style="background:#dc2626;color:#fff;padding:2px 8px;border-radius:99px;font-size:11px;font-weight:700">⚠️ EN RETARD de ' + t.joursRetard + ' j</span>'
        : (t.etat === "jour" ? '<span style="background:#F26522;color:#fff;padding:2px 8px;border-radius:99px;font-size:11px;font-weight:700">Aujourd\'hui</span>'
        : '<span style="background:#b91c1c;color:#fff;padding:2px 8px;border-radius:99px;font-size:11px;font-weight:700">🚩 Urgent</span>');
      return '<tr><td style="padding:10px 12px;border-bottom:1px solid #eef2f7;background:' + (t.etat === "retard" ? "#fef2f2" : "#fff") + '">'
        + '<div style="font-weight:700;color:#0f1b2d;font-size:14px">' + (t.urgent ? "🚩 " : "") + esc(t.titre) + '</div>'
        + (t.mandatLabel ? '<div style="color:#64748b;font-size:12px;margin-top:2px">Concerne : ' + esc(t.mandatLabel) + '</div>' : '')
        + '</td><td style="padding:10px 12px;border-bottom:1px solid #eef2f7;text-align:right;white-space:nowrap;background:' + (t.etat === "retard" ? "#fef2f2" : "#fff") + '">' + badge
        + (t.echeance ? '<div style="color:#64748b;font-size:11px;margin-top:4px">Échéance : ' + Utilities.formatDate(new Date(t.echeance + "T12:00:00"), tz, "dd/MM/yyyy") + '</div>' : '') + '</td></tr>';
    };
    const html = '<div style="font-family:Segoe UI,Arial,sans-serif;background:#f5f7fb;padding:24px"><div style="max-width:620px;margin:0 auto;background:#fff;border-radius:16px;overflow:hidden;border:1px solid #e6ebf2">'
      + '<div style="background:' + (retard.length ? "#b91c1c" : "#1A4E8A") + ';padding:18px 24px;color:#fff;font-size:18px;font-weight:700">'
      + (retard.length ? "⚠️ " + retard.length + " tâche" + (retard.length > 1 ? "s" : "") + " en retard" : "⏰ Rappel de vos tâches du jour") + '</div>'
      + '<div style="padding:20px 24px"><p style="margin:0 0 14px;color:#334155;font-size:14px">Bonjour ' + esc(nom) + ', voici vos tâches à traiter (' + list.length + ') :</p>'
      + '<table style="width:100%;border-collapse:collapse;border:1px solid #eef2f7;border-radius:10px">' + list.map(ligne).join("") + '</table>'
      + '<a href="' + url + '" style="display:inline-block;margin-top:18px;background:#F26522;color:#fff;text-decoration:none;font-weight:700;padding:11px 20px;border-radius:10px">Ouvrir mes tâches</a>'
      + '<p style="margin:18px 0 0;padding:12px 14px;background:#fff7ed;border-left:4px solid #F26522;color:#7c2d12;font-size:13px">En cas d\'impossibilité de réaliser une tâche, envoyez un mail à '
      + '<a href="mailto:' + TACHE_SUPERVISEUR + '?subject=' + encodeURIComponent("Tâche impossible à réaliser") + '" style="color:#c2410c;font-weight:700">' + TACHE_SUPERVISEUR + '</a> en expliquant pourquoi.</p>'
      + '<p style="margin:14px 0 0;color:#94a3b8;font-size:11px">Rappel automatique toutes les 2 h (7 h – 18 h, du lundi au vendredi) tant que des tâches du jour ou en retard restent ouvertes.</p></div></div></div>';
    const sujet = "[HECG] " + (retard.length ? "⚠️ " + retard.length + " tâche(s) EN RETARD — " : "⏰ ") + list.length + " tâche(s) à traiter aujourd'hui";
    try {
      const opts = { to: email, subject: sujet, htmlBody: html, name: "CRM HECG" };
      if (email !== TACHE_SUPERVISEUR) opts.cc = TACHE_SUPERVISEUR;
      MailApp.sendEmail(opts); nbMails++;
    } catch (eM) { Logger.log("Relance tâches " + email + " : " + eM); }
  });
  if (nbMails) logAction("Système", "Auto", "Action", "Missions — relances", "Relance auto", nbMails + " mail(s) de rappel envoyé(s) (tâches du jour / en retard), copie " + TACHE_SUPERVISEUR);
}

// Paramétrage (admin) : liste des catégories + notifications
function handleSaveTachesParams(p, ss) {
  if (!_tacheIsAdmin_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  try {
    if (Array.isArray(p.categories)) {
      const ws = _tacheCatSheet_(ss);
      const rows = p.categories.filter(function(c) { return String(c.nom || "").trim(); }).map(function(c, i) {
        return [String(c.id || ("CAT" + Date.now() + i)), String(c.nom).trim(), String(c.couleur || "#64748b"), i + 1];
      });
      if (ws.getLastRow() > 1) ws.getRange(2, 1, ws.getLastRow() - 1, 4).clearContent();
      if (rows.length) ws.getRange(2, 1, rows.length, 4).setValues(rows);
    }
    const setParam = function(cle, on) {
      const pr = _tacheParams_(ss);
      const d = pr.ws.getDataRange().getValues();
      let found = false;
      for (let i = 1; i < d.length; i++) if (String(d[i][0]) === cle) { pr.ws.getRange(i + 1, 2).setValue(on ? "Oui" : "Non"); found = true; }
      if (!found) pr.ws.appendRow([cle, on ? "Oui" : "Non"]);
    };
    if (p.notifMail !== undefined) setParam("notifMail", p.notifMail);
    if (p.relanceAuto !== undefined) { setParam("relanceAuto", p.relanceAuto); if (p.relanceAuto) _ensureRelanceTachesTrigger_(); }
    logAction(p.idAuteur || p.nomAuteur, p.roleAuteur, "Modification", "Missions", "PARAMS", "Paramètres des tâches mis à jour"
      + (p.notifMail !== undefined ? " — notifications mail : " + (p.notifMail ? "activées" : "désactivées") : "")
      + (p.relanceAuto !== undefined ? " — relances 2 h : " + (p.relanceAuto ? "activées" : "désactivées") : ""));
    return createJsonResponse({ success: true });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}


/**
 * Purge en masse (admin) : archive les tâches choisies sans les supprimer.
 * La date d'archivage est écrite en colonne X ; le statut est conservé
 * (une tâche terminée continue de compter pour les trophées).
 */
function handlePurgeTaches(p, ss) {
  if (!_tacheIsAdmin_(p)) return createJsonResponse({ success: false, message: "Seul l'administrateur peut purger des tâches." });
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const ids = {}; (Array.isArray(p.ids) ? p.ids : []).forEach(function(id) { ids[String(id)] = true; });
    const n = Object.keys(ids).length;
    if (!n) return createJsonResponse({ success: false, message: "Aucune tâche sélectionnée." });
    const ws = _tacheSheet_(ss);
    const d = ws.getDataRange().getValues();
    const col = ws.getRange(2, TACHE_COL.ARCHIVE + 1, Math.max(1, d.length - 1), 1).getValues();
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    const emails = {};
    let count = 0;
    for (let i = 1; i < d.length; i++) {
      if (!ids[String(d[i][0])] || col[i - 1][0]) continue;
      col[i - 1][0] = now; count++;
      emails[String(d[i][TACHE_COL.EMAIL] || "").toLowerCase()] = true;
    }
    if (count) ws.getRange(2, TACHE_COL.ARCHIVE + 1, col.length, 1).setValues(col);
    Object.keys(emails).forEach(function(e) { _tacheInvalidate_(e); });
    logAction(p.idAuteur || p.nomAuteur, p.roleAuteur, "Modification", "Missions", "PURGE", count + " tâche(s) archivée(s) (purge en masse)");
    return createJsonResponse({ success: true, archivees: count });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
  finally { try { lock.releaseLock(); } catch (eL) {} }
}
