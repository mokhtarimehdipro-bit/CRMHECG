// ==========================================
// VIE SCOLAIRE — SIMULATIONS D'EXAMENS BLANCS (refonte nov. 2026)
// Front : exsim.js (overlay plein écran dans Vie scolaire › Salles & Simulations › Simulations)
// ==========================================
// Feuille SIMULATIONS_EXAMENS : une ligne par simulation, le JSON est découpé en tranches
// de 45 000 caractères (colonnes H et suivantes) pour dépasser la limite de 50 000 / cellule.
// Modèle (JSON) :
// { titre, lundi:"yyyy-MM-dd", nbJours:5|6, hDebut:"08:00", hFin:"19:00", statut:"brouillon"|"terminee",
//   surveillants:[{id, nom, source, dispo:{ "0":[["08:00","12:00"],…] | absent → toute la journée }}],
//   classes:[{classe, effectif}], salles:[{id:"campus||num", campus, num, capacite}],
//   epreuves:[{id, classe, jour, debut, fin, matiere, note, salles:[{id, surv:[idSurv…]}]}],
//   optim:{ params, groupes:[…], placements… } }
// Règles dures (rappelées ici et revérifiées pour passer au statut « terminée ») :
//  - une salle n'accueille jamais deux classes différentes en même temps ;
//  - un surveillant n'est jamais dans deux salles en même temps ;
//  - un surveillant n'est affecté que sur un créneau couvert par ses disponibilités.

var EXSIM_HEADERS = ["ID", "Titre", "Lundi", "Statut", "Mis à jour le", "Par", "Version", "JSON (tranches →)"];
var EXSIM_TRANCHE = 45000;

function _exsimSheet_(ss) { return _vsxSheet_(ss, "SIMULATIONS_EXAMENS", EXSIM_HEADERS); }
function _exsimPeutEcrire_(p) { return ["Super-admin", "Pédagogie"].indexOf(String(p.vieScRole || "")) !== -1 || String(p.emailAuteur || "").toLowerCase() === "m.mokhtari@hecg.fr"; }
function _exsimMin_(h) { var m = String(h || "").match(/^(\d{1,2}):(\d{2})$/); return m ? (+m[1]) * 60 + (+m[2]) : NaN; }
function _exsimFmt_(d) { return d instanceof Date ? Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(d || ""); }

function _exsimLireLigne_(r) {
  var json = "";
  for (var c = 7; c < r.length; c++) json += String(r[c] || "");
  var o = {}; try { o = JSON.parse(json || "{}"); } catch (e) { o = {}; }
  return o;
}

// Contrôle des règles dures — renvoie la liste des conflits (texte)
function _exsimConflits_(S) {
  var out = [], ep = S.epreuves || [], surv = {};
  (S.surveillants || []).forEach(function (s) { surv[s.id] = s; });
  for (var i = 0; i < ep.length; i++) {
    var a = ep[i], a0 = _exsimMin_(a.debut), a1 = _exsimMin_(a.fin);
    if (!(a1 > a0)) out.push(a.classe + " — " + (a.matiere || "épreuve") + " : horaires invalides");
    (a.salles || []).forEach(function (sa) {
      (sa.surv || []).forEach(function (sid) {
        var s = surv[sid]; if (!s) { out.push("Surveillant inconnu sur " + a.classe); return; }
        var pl = s.dispo && s.dispo[String(a.jour)];
        if (pl && !pl.some(function (p) { return _exsimMin_(p[0]) <= a0 && _exsimMin_(p[1]) >= a1; }))
          out.push(s.nom + " n'est pas disponible le jour " + (a.jour + 1) + " de " + a.debut + " à " + a.fin);
      });
    });
    for (var j = i + 1; j < ep.length; j++) {
      var b = ep[j];
      if (b.jour !== a.jour || !(_exsimMin_(b.debut) < a1 && _exsimMin_(b.fin) > a0)) continue;
      (a.salles || []).forEach(function (sa) {
        (b.salles || []).forEach(function (sb) {
          if (sa.id === sb.id && a.classe !== b.classe) out.push("Salle " + sa.id.split("||")[1] + " : " + a.classe + " et " + b.classe + " en même temps");
          (sa.surv || []).forEach(function (x) { if (sa.id !== sb.id && (sb.surv || []).indexOf(x) !== -1) out.push((surv[x] ? surv[x].nom : x) + " dans deux salles en même temps (" + a.classe + " / " + b.classe + ")"); });
        });
      });
    }
  }
  return out;
}

// Données de référence : classes + étudiants (nominatif), salles, plans de salle, surveillants possibles, matières EXB
function _exsimRef_(ss) {
  var parClasse = _splanEtudiantsParClasse_(ss);
  var classes = Object.keys(parClasse).sort().map(function (c) { return { classe: c, effectif: parClasse[c].length }; });
  var salles = [], d = _getSallesSheet(ss).getDataRange().getValues();
  for (var i = 1; i < d.length; i++) {
    var campus = String(d[i][0] || "").trim(), num = String(d[i][1] || "").trim();
    if (!campus && !num) continue;
    salles.push({ id: campus + "||" + num, campus: campus, num: num, capacite: parseInt(d[i][2]) || 0 });
  }
  var plans = {};
  try {
    var pl = _splanSheet_(ss).getDataRange().getValues();
    for (var k = 1; k < pl.length; k++) { try { var o = JSON.parse(String(pl[k][3] || "null")); if (o && o.bureaux && o.bureaux.length) plans[String(pl[k][0])] = o; } catch (e) {} }
  } catch (e) {}
  var personnes = [];
  try {
    var pd = getPersonnelDataCached_(ss);
    for (var a = 1; a < pd.length; a++) if (pd[a][0]) personnes.push({ id: "P-" + pd[a][0], nom: (String(pd[a][2] || "") + " " + String(pd[a][1] || "")).trim(), source: "Personnel" + (pd[a][5] ? " · " + pd[a][5] : "") });
  } catch (e) {}
  try {
    var ws; try { ws = getSheetSafe(ss, "Adultes"); } catch (e1) { ws = getSheetSafe(ss, "FORMATEURS"); }
    var fd = ws.getDataRange().getValues();
    for (var f = 1; f < fd.length; f++) if (fd[f][0] && !fd[f][10]) personnes.push({ id: "F-" + fd[f][0], nom: (String(fd[f][2] || "") + " " + String(fd[f][1] || "")).trim(), source: "Formateur" });
  } catch (e) {}
  var exb = {};
  try {
    var ed = _getSuiviEXBSheet(ss).getDataRange().getValues();
    for (var e = 1; e < ed.length; e++) {
      var cl = String(ed[e][1] || "").trim(), mt = String(ed[e][2] || "").trim();
      if (cl && mt) { exb[cl] = exb[cl] || []; if (exb[cl].indexOf(mt) === -1) exb[cl].push(mt); }
    }
  } catch (e) {}
  return { classes: classes, etudiants: parClasse, salles: salles, plans: plans, personnes: personnes, matieresExb: exb };
}

function handleExsim(p, ss) {
  if (!_vsxCanUse_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  var op = String(p.op || "");
  var auteur = p.nomAuteur || p.emailAuteur || "";
  try {
    if (op === "ref") { var R = _exsimRef_(ss); R.success = true; return createJsonResponse(R); }

    var ws = _exsimSheet_(ss), d = ws.getDataRange().getValues();
    if (op === "liste") {
      var L = [];
      for (var i = 1; i < d.length; i++) {
        if (!d[i][0]) continue;
        var o = _exsimLireLigne_(d[i]);
        L.push({ id: String(d[i][0]), titre: String(d[i][1] || ""), lundi: _exsimFmt_(d[i][2]), statut: String(d[i][3] || "brouillon"),
          maj: d[i][4] instanceof Date ? Utilities.formatDate(d[i][4], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : String(d[i][4] || ""),
          par: String(d[i][5] || ""), nbClasses: (o.classes || []).length, nbSalles: (o.salles || []).length,
          nbSurv: (o.surveillants || []).length, nbEpreuves: (o.epreuves || []).length, optim: !!(o.optim && o.optim.groupes) });
      }
      L.reverse();
      return createJsonResponse({ success: true, data: L, ecriture: _exsimPeutEcrire_(p) });
    }
    if (op === "get") {
      for (var g = 1; g < d.length; g++) if (String(d[g][0]) === String(p.id))
        return createJsonResponse({ success: true, id: String(d[g][0]), version: String(d[g][6] || ""), sim: _exsimLireLigne_(d[g]), ecriture: _exsimPeutEcrire_(p) });
      return createJsonResponse({ success: false, message: "Simulation introuvable." });
    }

    if (!_exsimPeutEcrire_(p)) return createJsonResponse({ success: false, message: "Accès refusé (Super-admin ou Pédagogie)." });
    var lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      d = ws.getDataRange().getValues();
      if (op === "supprimer") {
        for (var s = d.length - 1; s >= 1; s--) if (String(d[s][0]) === String(p.id)) {
          ws.deleteRow(s + 1);
          logAction(auteur, p.vieScRole || "", "Suppression", "Vie scolaire — Simulations d'examens", String(p.id), "Simulation « " + d[s][1] + " » supprimée");
          return createJsonResponse({ success: true });
        }
        return createJsonResponse({ success: false, message: "Simulation introuvable." });
      }
      if (op === "enregistrer") {
        var S = typeof p.sim === "string" ? JSON.parse(p.sim) : (p.sim || {});
        if (!String(S.titre || "").trim()) return createJsonResponse({ success: false, message: "Titre obligatoire." });
        if (S.statut === "terminee") {
          var cf = _exsimConflits_(S);
          if (cf.length) return createJsonResponse({ success: false, message: "Impossible de terminer la semaine : " + cf.length + " conflit(s). " + cf.slice(0, 5).join(" ; "), conflits: cf });
        }
        var json = JSON.stringify(S);
        var tranches = []; for (var t = 0; t < json.length; t += EXSIM_TRANCHE) tranches.push(json.slice(t, t + EXSIM_TRANCHE));
        if (tranches.length > 40) return createJsonResponse({ success: false, message: "Simulation trop volumineuse." });
        var id = String(p.id || ""), ligne = -1, ancienne = null;
        if (id) for (var x = 1; x < d.length; x++) if (String(d[x][0]) === id) { ligne = x + 1; ancienne = d[x]; break; }
        if (ligne > 0 && String(p.version || "") !== String(ancienne[6] || ""))
          return createJsonResponse({ success: false, conflitVersion: true, message: "Cette simulation a été modifiée entre-temps par " + (ancienne[5] || "un autre utilisateur") + ". Rechargez-la avant d'enregistrer." });
        if (ligne < 0) { id = "EXS-" + Date.now(); ligne = ws.getLastRow() + 1; }
        var version = String(Date.now());
        var row = [id, S.titre, S.lundi || "", S.statut || "brouillon", new Date(), auteur, version].concat(tranches);
        var largeur = Math.max(ws.getLastColumn(), row.length);
        while (row.length < largeur) row.push("");
        if (ws.getMaxColumns() < largeur) ws.insertColumnsAfter(ws.getMaxColumns(), largeur - ws.getMaxColumns());
        ws.getRange(ligne, 1, 1, largeur).setNumberFormat("@").setValues([row]);
        var avant = ancienne ? _exsimLireLigne_(ancienne) : null;
        var det = "« " + S.titre + " » (semaine du " + (S.lundi || "?") + ") : " + (S.classes || []).length + " classe(s), " + (S.salles || []).length + " salle(s), "
          + (S.surveillants || []).length + " surveillant(s), " + (S.epreuves || []).length + " épreuve(s), statut " + (S.statut || "brouillon")
          + (avant ? " — avant : " + (avant.epreuves || []).length + " épreuve(s), statut " + (avant.statut || "brouillon") : "")
          + (S.optim && S.optim.groupes ? " — optimisation : " + S.optim.groupes.length + " salle(s)·créneau(x)" : "");
        logAction(auteur, p.vieScRole || "", ancienne ? "Modification" : "Création", "Vie scolaire — Simulations d'examens", id, det);
        return createJsonResponse({ success: true, id: id, version: version });
      }
    } finally { try { lock.releaseLock(); } catch (eL) {} }
    return createJsonResponse({ success: false, message: "Opération inconnue : " + op });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}
