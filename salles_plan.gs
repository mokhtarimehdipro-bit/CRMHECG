// ==========================================
// VIE SCOLAIRE — PLANS DE SALLE (éditeur salles_plan.js dans vie_scolaire.html)
// ==========================================
// Feuille SALLES_PLANS : une ligne par salle (ID = « campus||numéro », même clé que getSalles).
// Plan (JSON) : { largeur, hauteur, bureaux:[{ id, type:"eleve"|"prof", x, y, w, h, places:[nom, …], label }] }
// Comptage des étudiants par classe = même règle que _countStudentsByClasse (hors SORTIE / PROSPECT).

var SPLAN_HEADERS = ["ID salle", "Campus", "Salle", "Plan (JSON)", "Classe simulée", "Mis à jour le", "Par"];

function _splanSheet_(ss) { return _vsxSheet_(ss, "SALLES_PLANS", SPLAN_HEADERS); }

function _splanEtudiantsParClasse_(ss) {
  const d = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues();
  const map = {};
  for (let i = 1; i < d.length; i++) {
    const classe = String(d[i][14] || "").trim(), typo = String(d[i][20] || "").trim().toUpperCase();
    if (!classe || /^SORTIE/.test(typo) || typo === "PROSPECT") continue;
    (map[classe] = map[classe] || []).push({ id: String(d[i][0]), nom: (String(d[i][3] || "").trim() + " " + String(d[i][2] || "").trim().toUpperCase()).trim() });
  }
  Object.keys(map).forEach(function(c) { map[c].sort(function(a, b) { return a.nom.localeCompare(b.nom, "fr"); }); });
  return map;
}

function handleGetSallePlan(p, ss) {
  if (!_vsxCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  try {
    const d = _splanSheet_(ss).getDataRange().getValues();
    let plan = null, classe = "", maj = "", par = "";
    for (let i = 1; i < d.length; i++) if (String(d[i][0]) === String(p.salleId)) {
      try { plan = JSON.parse(String(d[i][3] || "null")); } catch (e) { plan = null; }
      classe = String(d[i][4] || "");
      maj = d[i][5] instanceof Date ? Utilities.formatDate(d[i][5], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : String(d[i][5] || "");
      par = String(d[i][6] || "");
      break;
    }
    const parClasse = _splanEtudiantsParClasse_(ss);
    const classes = Object.keys(parClasse).sort().map(function(c) { return { classe: c, effectif: parClasse[c].length }; });
    return createJsonResponse({ success: true, plan: plan, classe: classe, maj: maj, par: par, classes: classes, etudiantsParClasse: parClasse });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleSaveSallePlan(p, ss) {
  if (["Super-admin", "Pédagogie"].indexOf(String(p.vieScRole || "")) === -1) return createJsonResponse({ success: false, message: "Accès refusé." });
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const plan = p.plan || {};
    const bureaux = Array.isArray(plan.bureaux) ? plan.bureaux.slice(0, 300) : [];
    const json = JSON.stringify({ largeur: Number(plan.largeur) || 900, hauteur: Number(plan.hauteur) || 600, bureaux: bureaux });
    if (json.length > 45000) return createJsonResponse({ success: false, message: "Plan trop volumineux." });
    const ws = _splanSheet_(ss), d = ws.getDataRange().getValues();
    const parts = String(p.salleId || "").split("||");
    const row = [String(p.salleId), parts[0] || "", parts[1] || "", json, String(p.classe || ""), new Date(), String(p.nomAuteur || p.emailAuteur || "")];
    let found = false;
    for (let i = 1; i < d.length; i++) if (String(d[i][0]) === String(p.salleId)) { ws.getRange(i + 1, 1, 1, row.length).setValues([row]); found = true; break; }
    if (!found) ws.appendRow(row);
    const nbPlaces = bureaux.filter(function(b) { return b.type !== "prof"; }).reduce(function(s, b) { return s + (Array.isArray(b.places) ? b.places.length : 0); }, 0);
    const nbNommees = bureaux.filter(function(b) { return b.type !== "prof"; }).reduce(function(s, b) { return s + (Array.isArray(b.places) ? b.places.filter(function(x) { return String(x || "").trim(); }).length : 0); }, 0);
    logAction(p.nomAuteur || p.emailAuteur, p.vieScRole || "", found ? "Modification" : "Création", "Vie scolaire — Salles", String(p.salleId),
      "Plan de salle " + (parts[1] || "") + " (" + (parts[0] || "") + ") : " + bureaux.length + " bureau(x), " + nbPlaces + " place(s)" + (p.classe ? ", simulation " + p.classe + " (" + nbNommees + " place(s) nommée(s))" : ""));
    return createJsonResponse({ success: true });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
  finally { try { lock.releaseLock(); } catch (eL) {} }
}
