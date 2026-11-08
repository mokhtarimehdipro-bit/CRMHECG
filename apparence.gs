/* =============================================================================
   APPARENCE — thèmes personnalisés du CRM (novembre 2026)
   Réservé aux membres de niveau ≥ 14 (« Mythe », 35 000 XP, cf. NIVEAUX de trophees.html)
   et au compte maître (PARAM_MASTER_EMAILS, acces_equipe.gs).
   Le thème choisi est enregistré dans la feuille Personnel, colonne « Apparence »
   (créée au besoin) → il suit l'utilisateur d'un poste à l'autre.
   Front : hecg-skins.js + hecg-skins.css.
   ========================================================================== */

var APPARENCE_COL = "Apparence";
var APPARENCE_NIVEAU_MIN = 14;
// Seuils d'XP des 15 niveaux — doit rester aligné avec NIVEAUX dans trophees.html
var APPARENCE_SEUILS_XP = [0, 100, 250, 500, 900, 1500, 2500, 4000, 6000, 9000, 13000, 18000, 25000, 35000, 50000];
var APPARENCE_THEMES = {
  classique: "Classique HECG", aurore: "Aurore", ocean: "Océan", emeraude: "Émeraude",
  solaire: "Solaire", saphir: "Saphir & or", nuit: "Nuit néon"
};

function _apparenceNiveau_(xp) {
  var n = 1;
  for (var i = 1; i < APPARENCE_SEUILS_XP.length; i++) if (xp >= APPARENCE_SEUILS_XP[i]) n = i + 1;
  return n;
}

// XP de l'utilisateur : mémorisée par handleGetTropheeData (propriété XP_<email>) ;
// recalculée si absente ou sous le seuil (l'utilisateur a pu progresser depuis).
function _apparenceXp_(ss, email, nom) {
  var xp = -1;
  try { var v = PropertiesService.getScriptProperties().getProperty("XP_" + email); if (v !== null) xp = Number(v) || 0; } catch (e) {}
  if (xp >= APPARENCE_SEUILS_XP[APPARENCE_NIVEAU_MIN - 1]) return xp;
  try {
    var out = handleGetTropheeData({ emailUser: email, nomUser: nom || "" }, ss);
    var j = JSON.parse(out.getContent());
    if (j && j.success !== false && j.xpTotal !== undefined) xp = Number(j.xpTotal) || 0;
  } catch (e2) {}
  return Math.max(xp, 0);
}

function _apparenceDroits_(ss, p) {
  var email = String(p.emailAuteur || "").trim().toLowerCase();   // issu du jeton (doPost)
  var master = PARAM_MASTER_EMAILS.indexOf(email) !== -1;
  var xp = master ? null : _apparenceXp_(ss, email, p.nomAuteur);
  var niveau = master ? null : _apparenceNiveau_(xp);
  return { email: email, master: master, xp: xp, niveau: niveau, eligible: master || niveau >= APPARENCE_NIVEAU_MIN };
}

// Colonne « Apparence » du Personnel (créée au besoin) → { ws, data, col (0-based) }
function _apparenceSheet_(ss) {
  var ws = getSheetSafe(ss, "Personnel");
  var lastCol = Math.max(ws.getLastColumn(), 1);
  var header = ws.getRange(1, 1, 1, lastCol).getValues()[0];
  var col = _paramColIdx_(header, APPARENCE_COL);
  if (col === -1) {
    if (ws.getMaxColumns() <= lastCol) ws.insertColumnsAfter(ws.getMaxColumns(), 1);
    ws.getRange(1, lastCol + 1).setValue(APPARENCE_COL).setFontWeight("bold");
    col = lastCol; lastCol++;
  }
  var data = ws.getRange(1, 1, Math.max(ws.getLastRow(), 1), lastCol).getValues();
  return { ws: ws, data: data, col: col };
}

function _apparenceLigne_(S, email) {
  for (var i = 1; i < S.data.length; i++) if (String(S.data[i][6] || "").trim().toLowerCase() === email) return i;
  return -1;
}

function handleGetApparence(p, ss) {
  try {
    var d = _apparenceDroits_(ss, p);
    var S = _apparenceSheet_(ss), i = _apparenceLigne_(S, d.email);
    var theme = i > 0 ? String(S.data[i][S.col] || "").trim() : "";
    if (!APPARENCE_THEMES[theme] || !d.eligible) theme = "classique";
    return createJsonResponse({ success: true, eligible: d.eligible, master: d.master, niveau: d.niveau, xp: d.xp,
      niveauMin: APPARENCE_NIVEAU_MIN, xpMin: APPARENCE_SEUILS_XP[APPARENCE_NIVEAU_MIN - 1], theme: theme });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleSaveApparence(p, ss) {
  try {
    var theme = String(p.theme || "").trim();
    if (!APPARENCE_THEMES[theme]) return createJsonResponse({ success: false, message: "Thème inconnu." });
    var d = _apparenceDroits_(ss, p);
    if (!d.eligible && theme !== "classique")
      return createJsonResponse({ success: false, message: "Les thèmes se débloquent au niveau " + APPARENCE_NIVEAU_MIN + "." });
    var S = _apparenceSheet_(ss), i = _apparenceLigne_(S, d.email);
    if (i < 1) return createJsonResponse({ success: false, message: "Utilisateur introuvable dans le Personnel." });
    var avant = String(S.data[i][S.col] || "").trim() || "classique";
    S.ws.getRange(i + 1, S.col + 1).setValue(theme);
    try { _paramInvalidatePersonnel_(); } catch (eC) {}
    if (avant !== theme) {
      logAction(p.nomAuteur || d.email, p.roleAuteur || "", "Modification", "Paramètres — Apparence", d.email,
        "Thème « " + APPARENCE_THEMES[theme] + " » appliqué (avant : « " + (APPARENCE_THEMES[avant] || avant) + " »)"
        + (d.master ? " — compte maître" : " — niveau " + d.niveau));
    }
    return createJsonResponse({ success: true, theme: theme });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}
