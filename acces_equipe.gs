// ==========================================
// PARAMÈTRES — ACCÈS DE L'ÉQUIPE (page parametres.html, compte maître uniquement) — fichier acces_equipe.gs
// ==========================================
// Onglet « Personnel » : A Ref_Perso · B NOM · C PRÉNOM · D FONCTION · E MDP · F ROLE (CRM)
//   · G Mail · H Role_Vie_Scolaire · I-K préférences trophées (non touchées ici)
// Colonnes ajoutées (repérées par leur en-tête, créées à la 1re utilisation) :
//   « Accès » (vide = actif, « Révoqué » = connexion refusée) · « Dernière connexion » · « Accès modifié le »
// Révoquer = connexion refusée + sessions ouvertes supprimées (effet immédiat).

var PARAM_MASTER_EMAILS = ["m.mokhtari@hecg.fr"];
var PARAM_COL_ACCES = "Accès", PARAM_COL_LASTLOGIN = "Dernière connexion", PARAM_COL_MAJ = "Accès modifié le";
var PARAM_ROLES_CRM = ["Administrateur", "Superviseur", "Communication", "Commercial", "Pédagogie"];
var PARAM_ROLES_VS = ["", "Super-admin", "Pédagogie", "Administratif"];

function _paramIsMaster_(ss, p) {
  const sess = _validateSession(ss, p.sessionToken);
  return !!sess && PARAM_MASTER_EMAILS.indexOf(String(sess.email || "").toLowerCase()) !== -1 ? sess : null;
}

function _paramNorm_(s) { return String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").trim().toLowerCase(); }

// Index (0-based) d'une colonne par en-tête dans un tableau de valeurs du Personnel
function _paramColIdx_(header, nom) {
  for (let c = 0; c < header.length; c++) if (_paramNorm_(header[c]) === _paramNorm_(nom)) return c;
  return -1;
}

// Garantit les colonnes ajoutées ; renvoie { ws, data, idx:{acces, last, maj} }
function _paramPersonnel_(ss) {
  const ws = getSheetSafe(ss, "Personnel");
  let lastCol = Math.max(ws.getLastColumn(), 11);
  let header = ws.getRange(1, 1, 1, lastCol).getValues()[0];
  [PARAM_COL_ACCES, PARAM_COL_LASTLOGIN, PARAM_COL_MAJ].forEach(function(h) {
    if (_paramColIdx_(header, h) === -1) {
      lastCol++;
      if (ws.getMaxColumns() < lastCol) ws.insertColumnsAfter(ws.getMaxColumns(), 1);
      ws.getRange(1, lastCol).setValue(h).setFontWeight("bold");
      header = ws.getRange(1, 1, 1, lastCol).getValues()[0];
    }
  });
  const data = ws.getRange(1, 1, Math.max(ws.getLastRow(), 1), lastCol).getValues();
  return { ws: ws, data: data, idx: { acces: _paramColIdx_(header, PARAM_COL_ACCES), last: _paramColIdx_(header, PARAM_COL_LASTLOGIN), maj: _paramColIdx_(header, PARAM_COL_MAJ) } };
}

function _paramInvalidatePersonnel_() {
  try { CacheService.getScriptCache().remove('personnel_data_v1'); } catch (e) {}
  try { _invalidatePersonnelExtCache_(); } catch (e2) {}
}

/** Utilisé par handleLogin et par le repli d'authentification de doPost. */
function _persoRevoque_(dPerso, i) {
  const c = _paramColIdx_(dPerso[0] || [], PARAM_COL_ACCES);
  return c !== -1 && _paramNorm_(dPerso[i][c]) === "revoque";
}

/** Note la date de dernière connexion (sans bloquer la connexion en cas d'erreur). */
function _paramNoteConnexion_(ss, ref) {
  try {
    const P = _paramPersonnel_(ss);
    for (let i = 1; i < P.data.length; i++) if (String(P.data[i][0]) === String(ref)) { P.ws.getRange(i + 1, P.idx.last + 1).setValue(new Date()); break; }
  } catch (e) { Logger.log("Dernière connexion : " + e); }
}

// Supprime les sessions d'un membre (Sessions + cache des jetons)
function _paramFermerSessions_(ss, ref) {
  const ws = _getSessionSheet(ss);
  const d = ws.getDataRange().getValues();
  const cache = CacheService.getScriptCache();
  let n = 0;
  for (let i = d.length - 1; i >= 1; i--) {
    if (String(d[i][1]) !== String(ref)) continue;
    try { cache.remove('sess_v1_' + d[i][0]); } catch (e) {}
    ws.deleteRow(i + 1); n++;
  }
  return n;
}

function _paramMembre_(row, idx, tz) {
  const fmt = function(v) { return v instanceof Date ? Utilities.formatDate(v, tz, "dd/MM/yyyy HH:mm") : String(v || ""); };
  return { ref: String(row[0] || ""), nom: String(row[1] || ""), prenom: String(row[2] || ""), fonction: String(row[3] || ""),
    aMdp: !!String(row[4] || ""), role: String(row[5] || ""), email: String(row[6] || ""), vieScRole: String(row[7] || ""),
    revoque: _paramNorm_(row[idx.acces]) === "revoque", derniereConnexion: fmt(row[idx.last]), accesModifie: fmt(row[idx.maj]) };
}

function handleGetAccesEquipe(p, ss) {
  if (!_paramIsMaster_(ss, p)) return createJsonResponse({ success: false, message: "Réservé au compte maître." });
  try {
    const P = _paramPersonnel_(ss), tz = Session.getScriptTimeZone();
    const list = [];
    for (let i = 1; i < P.data.length; i++) if (P.data[i][0]) list.push(_paramMembre_(P.data[i], P.idx, tz));
    const roles = PARAM_ROLES_CRM.slice();
    list.forEach(function(m) { if (m.role && roles.indexOf(m.role) === -1) roles.push(m.role); });
    return createJsonResponse({ success: true, data: list, rolesCrm: roles, rolesVs: PARAM_ROLES_VS, masters: PARAM_MASTER_EMAILS });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function _paramGenMdp_() {
  const c = "ABCDEFGHJKMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789";
  let s = ""; for (let i = 0; i < 10; i++) s += c.charAt(Math.floor(Math.random() * c.length));
  return s;
}

function _paramMailMdp_(m, mdp, auteur) {
  const html = '<div style="font-family:Segoe UI,Arial,sans-serif;background:#f5f7fb;padding:24px"><div style="max-width:520px;margin:0 auto;background:#fff;border-radius:16px;border:1px solid #e6ebf2;overflow:hidden">'
    + '<div style="background:#1A4E8A;color:#fff;padding:18px 24px;font-size:18px;font-weight:700">Vos accès au CRM HECG</div>'
    + '<div style="padding:22px 24px;color:#334155;font-size:14px;line-height:1.6">Bonjour ' + m.prenom + ',<br>voici vos identifiants de connexion au CRM HECG :'
    + '<div style="margin:16px 0;padding:14px 16px;background:#f5f8fc;border-radius:12px;border-left:4px solid #F26522">Identifiant : <b>' + (m.email || m.ref) + '</b><br>Mot de passe : <b style="font-family:monospace;font-size:15px">' + mdp + '</b></div>'
    + '<a href="https://mokhtarimehdipro-bit.github.io/CRMHECG/" style="display:inline-block;background:#F26522;color:#fff;text-decoration:none;font-weight:700;padding:11px 20px;border-radius:10px">Se connecter</a>'
    + '<p style="color:#94a3b8;font-size:12px;margin-top:18px">Message envoyé par ' + auteur + '. Ne partagez jamais ce mot de passe.</p></div></div></div>';
  MailApp.sendEmail({ to: m.email, subject: "[HECG] Vos accès au CRM", htmlBody: html, name: "CRM HECG", cc: PARAM_MASTER_EMAILS[0] });
}

/**
 * Création (p.membre.ref vide) ou modification d'un membre.
 * Champs : nom, prenom, fonction, email, role, vieScRole, mdp (création : obligatoire ou généré).
 */
function handleSaveAccesMembre(p, ss) {
  const sess = _paramIsMaster_(ss, p);
  if (!sess) return createJsonResponse({ success: false, message: "Réservé au compte maître." });
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const m = p.membre || {};
    const nom = String(m.nom || "").trim().toUpperCase(), prenom = String(m.prenom || "").trim();
    const email = String(m.email || "").trim().toLowerCase();
    if (!nom || !prenom) return createJsonResponse({ success: false, message: "Nom et prénom obligatoires." });
    if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return createJsonResponse({ success: false, message: "Mail invalide." });
    const P = _paramPersonnel_(ss), tz = Session.getScriptTimeZone(), now = new Date();
    for (let i = 1; i < P.data.length; i++) {
      if (email && String(P.data[i][6] || "").trim().toLowerCase() === email && String(P.data[i][0]) !== String(m.ref || ""))
        return createJsonResponse({ success: false, message: "Ce mail est déjà utilisé par " + P.data[i][2] + " " + P.data[i][1] + "." });
    }
    const auteur = sess.nom || sess.email;
    if (!m.ref) {
      // Création : référence PERS-n suivante, mot de passe saisi ou généré
      let max = 0;
      for (let i = 1; i < P.data.length; i++) { const n = parseInt(String(P.data[i][0]).replace(/\D/g, ""), 10); if (!isNaN(n) && n > max) max = n; }
      const ref = "PERS" + String(max + 1).padStart(3, "0");
      const mdp = String(m.mdp || "").trim() || _paramGenMdp_();
      const row = new Array(P.data[0].length).fill("");
      row[0] = ref; row[1] = nom; row[2] = prenom; row[3] = String(m.fonction || ""); row[4] = mdp; row[5] = String(m.role || "");
      row[6] = email; row[7] = String(m.vieScRole || ""); row[P.idx.maj] = now;
      P.ws.getRange(P.data.length + 1, 1, 1, row.length).setValues([row]);
      _paramInvalidatePersonnel_();
      let mailEnvoye = false;
      if (m.envoyerMail && email) { try { _paramMailMdp_({ prenom: prenom, email: email, ref: ref }, mdp, auteur); mailEnvoye = true; } catch (eM) { Logger.log(eM); } }
      logAction(auteur, sess.role, "Création", "Paramètres — Accès", ref, "Accès créé pour " + prenom + " " + nom + " (rôle CRM : " + (m.role || "—") + ", Vie scolaire : " + (m.vieScRole || "aucun") + ")" + (mailEnvoye ? " — identifiants envoyés par mail" : ""));
      return createJsonResponse({ success: true, ref: ref, mdp: mailEnvoye ? "" : mdp, mailEnvoye: mailEnvoye });
    }
    for (let i = 1; i < P.data.length; i++) {
      if (String(P.data[i][0]) !== String(m.ref)) continue;
      const old = P.data[i];
      if (PARAM_MASTER_EMAILS.indexOf(String(old[6] || "").toLowerCase()) !== -1 && String(m.role || old[5]) !== "Administrateur")
        return createJsonResponse({ success: false, message: "Le compte maître doit rester Administrateur." });
      const changes = [];
      const set = function(col, val, lib) { if (String(old[col] || "") !== String(val)) { changes.push(lib + " : " + (old[col] || "—") + " → " + (val || "—")); P.ws.getRange(i + 1, col + 1).setValue(val); } };
      set(1, nom, "Nom"); set(2, prenom, "Prénom"); set(3, String(m.fonction || ""), "Fonction"); set(5, String(m.role || ""), "Rôle CRM");
      set(6, email, "Mail"); set(7, String(m.vieScRole || ""), "Rôle Vie scolaire");
      if (!changes.length) return createJsonResponse({ success: true, inchange: true });
      P.ws.getRange(i + 1, P.idx.maj + 1).setValue(now);
      _paramInvalidatePersonnel_();
      // Un changement de rôle doit s'appliquer tout de suite : on ferme les sessions ouvertes
      const roleChange = changes.some(function(c) { return c.indexOf("Rôle") === 0 || c.indexOf("Mail") === 0; });
      if (roleChange) _paramFermerSessions_(ss, m.ref);
      logAction(auteur, sess.role, "Modification", "Paramètres — Accès", m.ref, prenom + " " + nom + " — " + changes.join(" | ") + (roleChange ? " (reconnexion demandée)" : ""));
      return createJsonResponse({ success: true, reconnexion: roleChange });
    }
    return createJsonResponse({ success: false, message: "Membre introuvable." });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
  finally { try { lock.releaseLock(); } catch (eL) {} }
}

/** Révoquer (p.revoquer = true) ou rétablir l'accès d'un membre. */
function handleSetAccesMembre(p, ss) {
  const sess = _paramIsMaster_(ss, p);
  if (!sess) return createJsonResponse({ success: false, message: "Réservé au compte maître." });
  try {
    const P = _paramPersonnel_(ss);
    for (let i = 1; i < P.data.length; i++) {
      if (String(P.data[i][0]) !== String(p.ref)) continue;
      if (PARAM_MASTER_EMAILS.indexOf(String(P.data[i][6] || "").toLowerCase()) !== -1) return createJsonResponse({ success: false, message: "Impossible de révoquer le compte maître." });
      const revoquer = p.revoquer === true;
      P.ws.getRange(i + 1, P.idx.acces + 1).setValue(revoquer ? "Révoqué" : "");
      P.ws.getRange(i + 1, P.idx.maj + 1).setValue(new Date());
      _paramInvalidatePersonnel_();
      const n = revoquer ? _paramFermerSessions_(ss, p.ref) : 0;
      const qui = P.data[i][2] + " " + P.data[i][1];
      logAction(sess.nom || sess.email, sess.role, revoquer ? "Suppression" : "Modification", "Paramètres — Accès", p.ref,
        (revoquer ? "Accès RÉVOQUÉ pour " + qui + " (" + n + " session(s) fermée(s))" : "Accès rétabli pour " + qui) + (p.motif ? " — Motif : " + p.motif : ""));
      return createJsonResponse({ success: true, sessionsFermees: n });
    }
    return createJsonResponse({ success: false, message: "Membre introuvable." });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

/** Change le mot de passe : p.mdp saisi, sinon généré ; p.envoyerMail pour l'envoyer au membre. */
function handleResetMdpMembre(p, ss) {
  const sess = _paramIsMaster_(ss, p);
  if (!sess) return createJsonResponse({ success: false, message: "Réservé au compte maître." });
  try {
    const mdp = String(p.mdp || "").trim() || _paramGenMdp_();
    if (mdp.length < 6) return createJsonResponse({ success: false, message: "6 caractères minimum." });
    const P = _paramPersonnel_(ss);
    for (let i = 1; i < P.data.length; i++) {
      if (String(P.data[i][0]) !== String(p.ref)) continue;
      P.ws.getRange(i + 1, 5).setValue(mdp);
      P.ws.getRange(i + 1, P.idx.maj + 1).setValue(new Date());
      _paramInvalidatePersonnel_();
      const n = _paramFermerSessions_(ss, p.ref);
      const m = { prenom: String(P.data[i][2] || ""), email: String(P.data[i][6] || "").trim(), ref: String(P.data[i][0]) };
      let mailEnvoye = false;
      if (p.envoyerMail && m.email) { try { _paramMailMdp_(m, mdp, sess.nom || sess.email); mailEnvoye = true; } catch (eM) { Logger.log(eM); } }
      logAction(sess.nom || sess.email, sess.role, "Modification", "Paramètres — Accès", p.ref, "Mot de passe changé pour " + m.prenom + " " + P.data[i][1] + (mailEnvoye ? " — envoyé par mail" : "") + " (" + n + " session(s) fermée(s))");
      return createJsonResponse({ success: true, mdp: mailEnvoye ? "" : mdp, mailEnvoye: mailEnvoye });
    }
    return createJsonResponse({ success: false, message: "Membre introuvable." });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}
