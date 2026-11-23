// ==========================================
// SIMULATIONS D'EXAMENS — SUIVI DES SUJETS (remplace « Suivi EXB », nov. 2026)
// Front : exsim.js (onglet « Suivi & sujets » + formulaire d'épreuve)
// ==========================================
// - Formateurs de l'épreuve : dans le JSON de la simulation (epreuve.formateurs = [{id, nom, email}]).
// - Sujet / corrigé / impression / historique des relances : feuille EXSIM_SUIVI, une ligne par épreuve
//   (clé = id simulation | id épreuve) — indépendant du verrou « semaine terminée ».
// - Fichiers : GED › Examens blancs › <titre de la session> › <classe> › « <UE> - Sujet » / « <UE> - Corrigé ».
// - Relances automatiques (via auditAdressesTick, une fois par jour à partir de 8 h) :
//   J-30, J-14 puis chaque jour jusqu'à la veille (« Pour rappel »), tant que le sujet n'est pas déposé.
//   Un seul mail par épreuve, tous les formateurs de l'épreuve en destinataires.

var EXSUIVI_HEADERS = ["Clé", "Simulation", "Épreuve", "Classe", "Matière", "Sujet ID", "Sujet URL", "Sujet nom",
  "Corrigé ID", "Corrigé URL", "Corrigé nom", "Imprimé", "Relances (JSON)", "Mis à jour le", "Par"];
var EXSUIVI_EXPEDITEUR = "HECG — Équipe pédagogique";
var EXSUIVI_CAT = "Vie scolaire — Simulations d'examens";

function _exsuiviSheet_(ss) { return _vsxSheet_(ss, "EXSIM_SUIVI", EXSUIVI_HEADERS); }

function _exsuiviLire_(ss, simId) {
  var d = _exsuiviSheet_(ss).getDataRange().getValues(), out = {};
  for (var i = 1; i < d.length; i++) {
    var k = String(d[i][0] || ""); if (!k) continue;
    if (simId && String(d[i][1]) !== String(simId)) continue;
    var rel = []; try { rel = JSON.parse(String(d[i][12] || "[]")) || []; } catch (e) {}
    out[k] = { row: i + 1, simId: String(d[i][1]), epId: String(d[i][2]),
      sujet: d[i][5] ? { id: String(d[i][5]), url: String(d[i][6]), nom: String(d[i][7]) } : null,
      corrige: d[i][8] ? { id: String(d[i][8]), url: String(d[i][9]), nom: String(d[i][10]) } : null,
      imprime: d[i][11] === true || String(d[i][11]).toUpperCase() === "TRUE" || String(d[i][11]) === "1",
      relances: rel };
  }
  return out;
}
// Écrit (crée si besoin) la ligne de l'épreuve ; f = fonction qui modifie l'objet
function _exsuiviMaj_(ss, sim, simId, ep, f, auteur) {
  var ws = _exsuiviSheet_(ss), cle = simId + "|" + ep.id, all = _exsuiviLire_(ss, simId);
  var o = all[cle] || { sujet: null, corrige: null, imprime: false, relances: [] };
  f(o);
  var row = [cle, simId, ep.id, ep.classe || "", ep.matiere || "",
    o.sujet ? o.sujet.id : "", o.sujet ? o.sujet.url : "", o.sujet ? o.sujet.nom : "",
    o.corrige ? o.corrige.id : "", o.corrige ? o.corrige.url : "", o.corrige ? o.corrige.nom : "",
    o.imprime ? "TRUE" : "", JSON.stringify((o.relances || []).slice(-60)), new Date(), auteur || ""];
  if (o.row) ws.getRange(o.row, 1, 1, row.length).setValues([row]);
  else ws.appendRow(row);
  return o;
}
function _exsuiviSim_(ss, simId) {
  var d = _exsimSheet_(ss).getDataRange().getValues();
  for (var i = 1; i < d.length; i++) if (String(d[i][0]) === String(simId)) return _exsimLireLigne_(d[i]);
  return null;
}
function _exsuiviDossier_(sim, classe) {
  var f = DriveApp.getFolderById(GED_ROOT_ID);
  [ "Examens blancs", String(sim.titre || "Session sans titre").replace(/[\\/:*?"<>|]+/g, " ").trim(), classe || "Sans classe" ].forEach(function (n) {
    var it = f.getFoldersByName(n); f = it.hasNext() ? it.next() : f.createFolder(n);
  });
  return f;
}
// Formateurs actifs (feuille Adultes / FORMATEURS) : id → {nom, prenom, email}
function _exsuiviFormateurs_(ss) {
  var out = {};
  try {
    var ws; try { ws = getSheetSafe(ss, "Adultes"); } catch (e1) { ws = getSheetSafe(ss, "FORMATEURS"); }
    var fd = ws.getDataRange().getValues();
    for (var f = 1; f < fd.length; f++) {
      if (!fd[f][0]) continue;
      var mail = String(fd[f][4] || "").trim() || String(fd[f][3] || "").trim();
      out["F-" + fd[f][0]] = { id: "F-" + fd[f][0], nom: (String(fd[f][2] || "") + " " + String(fd[f][1] || "")).trim(), prenom: String(fd[f][2] || "").trim(),
        email: mail.indexOf("@") > 0 ? mail : "", matieres: String(fd[f][8] || ""), classes: String(fd[f][14] || ""), sorti: !!fd[f][10] };
    }
  } catch (e) {}
  return out;
}
function _exsuiviDateEp_(sim, ep) {
  if (!sim.lundi) return null;
  var d = new Date(String(sim.lundi) + "T12:00:00"); if (isNaN(d)) return null;
  d.setDate(d.getDate() + (+ep.jour || 0)); return d;
}
function _exsuiviDateLongue_(d) {
  var J = ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"], M = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre"];
  return J[d.getDay()] + " " + d.getDate() + (d.getDate() === 1 ? "er" : "") + " " + M[d.getMonth()] + " " + d.getFullYear();
}
function _exsuiviH_(h) { return String(h || "").replace(":", "h"); }
function _exsuiviEsc_(s) { return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); }

// Destinataires de l'épreuve (e-mail rafraîchi depuis la fiche formateur)
function _exsuiviDest_(ep, F) {
  var L = [];
  (ep.formateurs || []).forEach(function (x) {
    var f = F[x.id] || {}, mail = f.email || x.email || "";
    if (mail && L.every(function (y) { return y.email.toLowerCase() !== mail.toLowerCase(); }))
      L.push({ nom: f.nom || x.nom || "", prenom: f.prenom || String(x.nom || "").split(" ")[0], email: mail });
  });
  return L;
}

// Mail de relance (HTML soigné + texte brut)
function _exsuiviMail_(sim, ep, dest, dateEp, rappel, manque, jours) {
  var prenoms = dest.map(function (d) { return d.prenom || d.nom; }).filter(Boolean);
  var salut = "Bonjour" + (prenoms.length ? " " + (prenoms.length > 1 ? prenoms.slice(0, -1).join(", ") + " et " + prenoms[prenoms.length - 1] : prenoms[0]) : "") + ",";
  var docs = manque.sujet && manque.corrige ? "le sujet et le corrigé" : manque.sujet ? "le sujet" : "le corrigé";
  var vous = dest.length > 1 ? "vous" : "vous";
  var mat = ep.matiere || "l'épreuve", session = sim.titre || "d'examens blancs";
  var quand = _exsuiviDateLongue_(dateEp), horaire = "de " + _exsuiviH_(ep.debut) + " à " + _exsuiviH_(ep.fin);
  var delai = jours === 1 ? "demain" : "dans " + jours + " jours";
  var objet = (rappel ? "Pour rappel — " : "") + "Examen blanc " + mat + " (" + ep.classe + ") du " + quand + " : " + docs + " à nous transmettre";
  var intro = (rappel ? "Pour rappel, l'épreuve" : "L'épreuve") + " de " + mat + " de la session d'examens « " + session + " » pour la classe de " + ep.classe
    + " est planifiée le " + quand + ", " + horaire + (rappel ? " (" + delai + ")" : "") + ".";
  var demande = "À ce propos, merci de bien vouloir nous envoyer par mail " + docs + " de cette épreuve. Il " + vous + " suffit de répondre à ce message en y joignant "
    + (manque.sujet && manque.corrige ? "les documents" : "le document") + ".";
  var pourquoi = "Les recevoir suffisamment tôt nous permet de préparer sereinement l'impression des copies et l'organisation de la session.";
  var merci = "Un grand merci pour votre implication et votre précieuse collaboration.";
  var corps = salut + "\n\n" + intro + "\n\n" + demande + "\n\n" + pourquoi + "\n\n" + merci + "\n\nBien cordialement,\nL'équipe pédagogique HECG";

  var ligne = function (k, v) { return '<tr><td style="padding:7px 14px 7px 0;color:#64748b;font-size:13px;white-space:nowrap;vertical-align:top">' + k + '</td><td style="padding:7px 0;color:#0e1a2b;font-size:14px;font-weight:bold">' + _exsuiviEsc_(v) + '</td></tr>'; };
  var html = '<div style="background:#f3f6fb;padding:28px 12px;font-family:Arial,Helvetica,sans-serif">'
    + '<div style="max-width:600px;margin:0 auto;background:#ffffff;border-radius:16px;overflow:hidden;box-shadow:0 10px 30px -18px rgba(14,26,43,.45)">'
    + '<div style="background:#1A4E8A;padding:22px 28px"><div style="color:#ffffff;font-size:22px;font-weight:800;letter-spacing:.5px">HECG<span style="color:#F26522">.</span></div>'
    + '<div style="color:#dce6f4;font-size:13px;margin-top:4px">' + (rappel ? "Pour rappel · " : "") + 'Examens blancs — sujet et corrigé</div></div>'
    + '<div style="height:4px;background:linear-gradient(90deg,#F26522 0,#F26522 35%,#1A4E8A 35%)"></div>'
    + '<div style="padding:28px;color:#1f2937;font-size:15px;line-height:1.6">'
    + '<p style="margin:0 0 16px">' + _exsuiviEsc_(salut) + '</p>'
    + '<p style="margin:0 0 18px">' + (rappel ? "<b>Pour rappel</b>, l'épreuve" : "L'épreuve") + ' de <b>' + _exsuiviEsc_(mat) + '</b> de la session d\'examens « ' + _exsuiviEsc_(session) + ' » pour la classe de <b>' + _exsuiviEsc_(ep.classe) + '</b> est planifiée le <b>' + _exsuiviEsc_(quand) + '</b>' + (rappel ? ' <span style="color:#F26522;font-weight:bold">(' + delai + ')</span>' : '') + '.</p>'
    + '<div style="background:#f8fafd;border:1px solid #e6ecf4;border-left:4px solid #F26522;border-radius:12px;padding:12px 18px;margin:0 0 20px"><table style="border-collapse:collapse">'
    + ligne("Épreuve", mat) + ligne("Classe", ep.classe) + ligne("Session", session) + ligne("Date", quand) + ligne("Horaires", horaire) + '</table></div>'
    + '<p style="margin:0 0 14px">À ce propos, merci de bien vouloir nous envoyer par mail <b>' + docs + '</b> de cette épreuve. Il vous suffit de <b>répondre à ce message</b> en y joignant ' + (manque.sujet && manque.corrige ? "les documents" : "le document") + '.</p>'
    + '<p style="margin:0 0 14px;color:#475569">' + pourquoi + '</p>'
    + '<p style="margin:0 0 22px">' + merci + '</p>'
    + '<p style="margin:0">Bien cordialement,<br><b style="color:#1A4E8A">L\'équipe pédagogique HECG</b></p></div>'
    + '<div style="background:#f8fafd;border-top:1px solid #e6ecf4;padding:14px 28px;color:#94a3b8;font-size:11.5px">Message automatique — vous ne recevrez plus de rappel dès que le sujet aura été enregistré.</div>'
    + '</div></div>';
  return { objet: objet, corps: corps, html: html };
}

// Envoi effectif pour une épreuve ; renvoie un texte d'erreur ou "" si envoyé
function _exsuiviEnvoyer_(ss, simId, sim, ep, F, rappel, par) {
  var dest = _exsuiviDest_(ep, F);
  if (!dest.length) return "aucun formateur avec adresse e-mail";
  var dEp = _exsuiviDateEp_(sim, ep); if (!dEp) return "date de l'épreuve inconnue (semaine non renseignée)";
  var tz = "Europe/Paris", auj = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  var jours = Math.round((new Date(Utilities.formatDate(dEp, tz, "yyyy-MM-dd") + "T12:00:00") - new Date(auj + "T12:00:00")) / 86400000);
  var suivi = _exsuiviLire_(ss, simId)[simId + "|" + ep.id] || {};
  var m = _exsuiviMail_(sim, ep, dest, dEp, rappel, { sujet: !suivi.sujet, corrige: !suivi.corrige }, jours);
  MailApp.sendEmail({ to: dest.map(function (d) { return d.email; }).join(","), subject: m.objet, body: m.corps, htmlBody: m.html, name: EXSUIVI_EXPEDITEUR });
  _exsuiviMaj_(ss, sim, simId, ep, function (o) {
    o.relances = (o.relances || []).concat([{ le: Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm"), jour: auj, j: jours, rappel: !!rappel, a: dest.map(function (d) { return d.email; }).join(", "), par: par || "Automatique" }]);
  }, par || "Automatique");
  return "";
}

// Relances quotidiennes (appelé par auditAdressesTick, adresses.gs)
function _exsuiviAuto_() {
  var now = new Date(), tz = "Europe/Paris", jour = Utilities.formatDate(now, tz, "yyyy-MM-dd"), props = PropertiesService.getScriptProperties();
  if (+Utilities.formatDate(now, tz, "H") < 8 || props.getProperty("EXSUIVI_RELANCE_JOUR") === jour) return;
  props.setProperty("EXSUIVI_RELANCE_JOUR", jour);
  _exsuiviRelancer_(SpreadsheetApp.openById(RDV_SS_ID));
}
function _exsuiviRelancer_(ss) {
  var tz = "Europe/Paris", auj = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  var d = _exsimSheet_(ss).getDataRange().getValues(), F = null, suivi = _exsuiviLire_(ss, ""), envoyes = 0, ko = [];
  for (var i = 1; i < d.length; i++) {
    var simId = String(d[i][0] || ""); if (!simId) continue;
    var sim = _exsimLireLigne_(d[i]);
    if (sim.relances === false || !sim.lundi) continue;
    (sim.epreuves || []).forEach(function (ep) {
      if (!(ep.formateurs || []).length) return;
      var dEp = _exsuiviDateEp_(sim, ep); if (!dEp) return;
      var jours = Math.round((new Date(Utilities.formatDate(dEp, tz, "yyyy-MM-dd") + "T12:00:00") - new Date(auj + "T12:00:00")) / 86400000);
      if (jours < 1 || jours > 30) return;
      var s = suivi[simId + "|" + ep.id] || { relances: [] };
      if (s.sujet) return;                                                         // sujet déposé : fin des relances
      var R = s.relances || [];
      if (R.some(function (r) { return r.jour === auj; })) return;                  // déjà relancé aujourd'hui
      if (!(jours === 30 || jours === 14 || jours <= 13 || !R.length)) return;      // J-30, J-14, puis chaque jour (ou rattrapage du 1er envoi)
      F = F || _exsuiviFormateurs_(ss);
      try {
        var err = _exsuiviEnvoyer_(ss, simId, sim, ep, F, R.length > 0, "Automatique");
        if (err) ko.push(ep.classe + " · " + (ep.matiere || "épreuve") + " : " + err); else envoyes++;
      } catch (e) { ko.push(ep.classe + " · " + (ep.matiere || "épreuve") + " : " + e.message); }
    });
  }
  if (envoyes || ko.length) logAction("Système", "Auto", "Action", EXSUIVI_CAT + " — Relances sujets", "EXSIM_SUIVI",
    "Relances formateurs (sujets/corrigés) : " + envoyes + " envoyée(s)" + (ko.length ? " — non envoyées : " + ko.join(" ; ") : ""));
  return { envoyes: envoyes, nonEnvoyes: ko };
}

// Opérations front (appelées depuis handleExsim)
function _exsuiviOp_(p, ss, op, auteur) {
  var simId = String(p.id || "");
  if (op === "suivi") {
    var all = _exsuiviLire_(ss, simId), out = {};
    Object.keys(all).forEach(function (k) { var o = all[k]; out[o.epId] = { sujet: o.sujet, corrige: o.corrige, imprime: o.imprime, relances: o.relances }; });
    return { success: true, data: out };
  }
  if (!_exsimPeutEcrire_(p)) return { success: false, message: "Accès refusé (réservé au personnel)." };
  var sim = _exsuiviSim_(ss, simId); if (!sim) return { success: false, message: "Enregistrez d'abord la simulation." };
  var ep = (sim.epreuves || []).filter(function (e) { return String(e.id) === String(p.epId); })[0];
  if (!ep) return { success: false, message: "Épreuve introuvable : enregistrez d'abord la simulation." };
  var lib = ep.classe + " · " + (ep.matiere || "épreuve");
  if (op === "docDeposer") {
    var type = p.type === "corrige" ? "corrige" : "sujet";
    if (!p.base64Data) return { success: false, message: "Fichier manquant." };
    var ext = (String(p.fileName || "").match(/\.[A-Za-z0-9]{1,6}$/) || [""])[0].toLowerCase();
    var nom = String(ep.matiere || "Épreuve").replace(/[\\/:*?"<>|]+/g, " ").trim() + " - " + (type === "sujet" ? "Sujet" : "Corrigé") + ext;
    var dossier = _exsuiviDossier_(sim, ep.classe);
    var fichier = dossier.createFile(Utilities.newBlob(Utilities.base64Decode(p.base64Data), p.mimeType || "application/octet-stream", nom));
    var o = _exsuiviMaj_(ss, sim, simId, ep, function (x) {
      if (x[type] && x[type].id) { try { DriveApp.getFileById(x[type].id).setTrashed(true); } catch (eT) {} }
      x[type] = { id: fichier.getId(), url: fichier.getUrl(), nom: nom };
    }, auteur);
    logAction(auteur, p.vieScRole || "", "Ajout", EXSUIVI_CAT + " — Sujets", simId, (type === "sujet" ? "Sujet" : "Corrigé") + " déposé : " + lib + " (« " + sim.titre + " ») → GED « " + nom + " »");
    return { success: true, data: { sujet: o.sujet, corrige: o.corrige, imprime: o.imprime, relances: o.relances } };
  }
  if (op === "docRetirer") {
    var t2 = p.type === "corrige" ? "corrige" : "sujet";
    var o2 = _exsuiviMaj_(ss, sim, simId, ep, function (x) { if (x[t2] && x[t2].id) { try { DriveApp.getFileById(x[t2].id).setTrashed(true); } catch (eT) {} } x[t2] = null; }, auteur);
    logAction(auteur, p.vieScRole || "", "Suppression", EXSUIVI_CAT + " — Sujets", simId, (t2 === "sujet" ? "Sujet" : "Corrigé") + " retiré : " + lib);
    return { success: true, data: { sujet: o2.sujet, corrige: o2.corrige, imprime: o2.imprime, relances: o2.relances } };
  }
  if (op === "imprime") {
    var o3 = _exsuiviMaj_(ss, sim, simId, ep, function (x) { x.imprime = String(p.valeur) === "true" || p.valeur === true; }, auteur);
    logAction(auteur, p.vieScRole || "", "Modification", EXSUIVI_CAT + " — Sujets", simId, "Impression " + (o3.imprime ? "faite" : "à faire") + " : " + lib);
    return { success: true, data: { sujet: o3.sujet, corrige: o3.corrige, imprime: o3.imprime, relances: o3.relances } };
  }
  if (op === "relancer") {
    var prev = _exsuiviLire_(ss, simId)[simId + "|" + ep.id] || { relances: [] };
    var err = _exsuiviEnvoyer_(ss, simId, sim, ep, _exsuiviFormateurs_(ss), (prev.relances || []).length > 0, auteur || "Manuel");
    if (err) return { success: false, message: "Relance non envoyée : " + err + "." };
    logAction(auteur, p.vieScRole || "", "Envoi", EXSUIVI_CAT + " — Relances sujets", simId, "Relance manuelle aux formateurs : " + lib);
    var o4 = _exsuiviLire_(ss, simId)[simId + "|" + ep.id];
    return { success: true, data: { sujet: o4.sujet, corrige: o4.corrige, imprime: o4.imprime, relances: o4.relances } };
  }
  return { success: false, message: "Opération inconnue : " + op };
}
