// ════════════════════════════════════════════════════════════════════════════
//  COMPTES D'UN NOUVEL ÉTUDIANT (inscrit / préinscrit) — nov. 2026
//  Deux opérations SÉPARÉES (droits différents) depuis la fiche étudiant (dossier.html) :
//  1) Charlemagne : pas d'API. Le CRM met l'étudiant dans une file d'attente (feuille CHARL_FILE) ;
//     l'agent installé sur le PC où Charlemagne est ouvert (agent_charlemagne.ps1, généré par le CRM)
//     la lit, saisit la fiche par pilotage d'écran puis renvoie le résultat. Reprise possible a posteriori.
//     Ouvert au personnel ayant la Vie scolaire (ou Administrateur).
//  2) Google Workspace : création directe par l'API Admin SDK (compte du CRM, administrateur) :
//     prénom, nom, adresse @hecg.fr pré-saisie dans la fiche (col. AA), mot de passe provisoire affiché
//     une seule fois, changement obligatoire à la 1re connexion. Réservé aux profils Super-admin.
// ════════════════════════════════════════════════════════════════════════════
var CHARL_FILE_HEADERS = ["ID", "Demandé le", "ID étudiant", "Nom", "Prénom", "Classe", "Statut", "Id Charlemagne",
  "Message", "Demandé par", "Traité le", "Poste", "Données (JSON)", "Photo"];
var COMPTES_CAT = "Comptes étudiants";
var CHARL_PAYS_FRANCE = "";   // né en France : champ pays laissé vide (Charlemagne = France par défaut)

function _cptSheet_(ss) { return _vsxSheet_(ss, "CHARL_FILE", CHARL_FILE_HEADERS); }
function _cptPeutCharl_(p) { return _vsxCanUse_(p); }
function _cptPeutWs_(p) { return _vsxIsSuper_(p); }
function _cptNow_() { return Utilities.formatDate(new Date(), "Europe/Paris", "dd/MM/yyyy HH:mm"); }

function _cptEtudiant_(ss, etuId) {
  var ws = getSheetSafe(ss, "ETUDIANTS"), d = ws.getDataRange().getValues(), h = d[0].map(function (x) { return String(x || "").trim(); });
  var cCrea = h.indexOf("Date création");
  for (var i = 1; i < d.length; i++) {
    var r = d[i]; if (String(r[0]).trim() !== String(etuId).trim()) continue;
    var typo = String(r[20] || "").trim(), tn = typo.toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[^a-z]/g, "");
    var crea = cCrea >= 0 ? r[cCrea] : "";
    return { row: i + 1, id: String(r[0]).trim(), nom: String(r[2] || "").trim(), prenom: String(r[3] || "").trim(), mail: String(r[4] || "").trim(),
      tel: String(r[5] || "").replace(/\s+/g, ""), idEnt: String(r[11] || "").trim(), idCont: String(r[12] || "").trim(), statut: String(r[13] || "").trim(),
      classe: String(r[14] || "").trim(), typo: typo, inscrit: tn === "inscrit" || tn === "preinscrit", entreprise: String(r[21] || "").trim(),
      campus: String(r[22] || "").trim(), idEleve: String(r[23] || "").trim(), mailHecg: String(r[26] || "").trim().toLowerCase(),
      dateCrea: crea instanceof Date ? Utilities.formatDate(crea, "Europe/Paris", "dd/MM/yyyy") : String(crea || "").slice(0, 10) };
  }
  return null;
}
// Maître d'apprentissage (contact de l'entreprise rattaché à la fiche)
function _cptMa_(ss, e) {
  if (!e.idCont) return null;
  try {
    var d = getSheetSafe(ss, "PARTENARIAT").getDataRange().getValues();
    for (var i = 1; i < d.length; i++) if (String(d[i][0]).trim() === e.idCont)
      return { nom: String(d[i][3] || "").trim(), prenom: String(d[i][4] || "").trim(), entreprise: String(d[i][2] || "").trim() };
  } catch (x) {}
  return null;
}
function _cptTel_(t) { t = String(t || "").replace(/[^\d+]/g, ""); if (/^\d{9}$/.test(t)) t = "0" + t; return t; }

// Date / lieu / pays de naissance lus dans le « Dossier de candidature » du dossier Drive (OCR Google)
function _cptNaissance_(ss, e) {
  var out = { date: "", lieu: "", pays: "", source: "" };
  try {
    var g = getOrCreateGedEtuSheet(ss).getDataRange().getValues(), fid = "";
    for (var i = 1; i < g.length; i++) if (String(g[i][0]) === e.id) { fid = String(g[i][2] || "").trim(); break; }
    if (!fid) { out.source = "aucun dossier Drive lié à la fiche"; return out; }
    var dossier = DriveApp.getFolderById(fid), cand = null, chercher = function (f, prof) {
      var it = f.getFiles(); while (it.hasNext() && !cand) { var x = it.next(); if (/candidature/i.test(x.getName()) && /pdf|image/i.test(x.getMimeType())) cand = x; }
      if (!cand && prof < 2) { var s = f.getFolders(); while (s.hasNext() && !cand) chercher(s.next(), prof + 1); }
    };
    chercher(dossier, 0);
    if (!cand) { out.source = "« Dossier de candidature » introuvable dans le dossier Drive"; return out; }
    var r = UrlFetchApp.fetch("https://www.googleapis.com/drive/v2/files/" + cand.getId() + "/copy?ocr=true&ocrLanguage=fr&convert=true", {
      method: "post", contentType: "application/json", payload: JSON.stringify({ title: "tmp-ocr-" + e.id }),
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    if (r.getResponseCode() >= 300) { out.source = "lecture du dossier de candidature impossible"; return out; }
    var docId = JSON.parse(r.getContentText()).id, txt = "";
    try { txt = DocumentApp.openById(docId).getBody().getText(); } finally { try { DriveApp.getFileById(docId).setTrashed(true); } catch (eT) {} }
    var m = txt.match(/N[ée]\s*\(?e?\)?\s*le\s*:?\s*(\d{1,2})\s*[\/.\-]\s*(\d{1,2})\s*[\/.\-]\s*(\d{2,4})/i);
    if (m) { var y = m[3].length === 2 ? (+m[3] > 30 ? "19" : "20") + m[3] : m[3]; out.date = ("0" + m[1]).slice(-2) + "/" + ("0" + m[2]).slice(-2) + "/" + y; }
    var l = txt.match(/Ville\s*\/\s*Pays\s*:?\s*([^\n,]+)/i) || txt.match(/(?:Lieu|Ville) de naissance\s*:?\s*([^\n,]+)/i);
    if (l) out.lieu = l[1].replace(/[.\s]+$/, "").trim().toUpperCase();
    // pays : n° de sécurité sociale (département 99 = né à l'étranger, puis code pays INSEE) — le numéro n'est pas conservé
    var s = txt.match(/S[ée]curit[ée]\s+sociale\s*:?\s*([\d\s]{13,25})/i);
    if (s) { var n = s[1].replace(/\s+/g, ""); if (n.length >= 10) { out.pays = n.substr(5, 2) === "99" ? n.substr(7, 3) : CHARL_PAYS_FRANCE; out.paysDeduit = true; } }
    out.source = "lu dans « " + cand.getName() + " » — à vérifier";
  } catch (x) { out.source = "lecture impossible : " + x.message; }
  return out;
}

function _cptFile_(ss) {
  var d = _cptSheet_(ss).getDataRange().getValues(), L = [];
  for (var i = 1; i < d.length; i++) if (d[i][0]) L.push({ row: i + 1, id: String(d[i][0]), le: String(d[i][1] instanceof Date ? Utilities.formatDate(d[i][1], "Europe/Paris", "dd/MM/yyyy HH:mm") : d[i][1]),
    etuId: String(d[i][2]), nom: String(d[i][3]), prenom: String(d[i][4]), classe: String(d[i][5]), statut: String(d[i][6]), idCharl: String(d[i][7] || ""),
    message: String(d[i][8] || ""), par: String(d[i][9] || ""), traite: String(d[i][10] instanceof Date ? Utilities.formatDate(d[i][10], "Europe/Paris", "dd/MM/yyyy HH:mm") : d[i][10] || ""),
    poste: String(d[i][11] || ""), donnees: String(d[i][12] || ""), photo: String(d[i][13] || "") });
  return L;
}
function _cptEtatCharl_(ss, etuId) {
  var L = _cptFile_(ss).filter(function (x) { return x.etuId === etuId; }), j = L[L.length - 1] || null;
  if (j && j.statut === "En attente") j.position = _cptFile_(ss).filter(function (x) { return x.statut === "En attente" && x.row <= j.row; }).length;
  return j ? { id: j.id, statut: j.statut, idCharl: j.idCharl, message: j.message, le: j.le, par: j.par, traite: j.traite, poste: j.poste, photo: j.photo, position: j.position || 0 } : null;
}

// ── Workspace ──────────────────────────────────────────────────────────────
function _cptMdp_() {
  var A = "ABCDEFGHJKLMNPQRSTUVWXYZ", a = "abcdefghijkmnopqrstuvwxyz", n = "23456789", s = "!?@#%*", all = A + a + n, out = "";
  var pick = function (x) { return x.charAt(Math.floor(Math.random() * x.length)); };
  out = pick(A) + pick(a) + pick(a) + pick(n) + pick(s); for (var i = 0; i < 7; i++) out += pick(all);
  return out.split("").sort(function () { return Math.random() - 0.5; }).join("");
}
// Unité d'organisation des étudiants : la plus fréquente parmi les comptes des fiches inscrites (mémorisée 1 jour)
function _cptOuEtudiants_(ss) {
  var c = CacheService.getScriptCache(), v = c.get("cpt_ou_etu"); if (v) return v;
  var comptes = {}, cnt = {};
  _wsComptes_().forEach(function (x) { comptes[x.email] = x; });
  _vsuEtudiants_(ss).forEach(function (e) { var x = e.inscrit && e.mailHecg && comptes[e.mailHecg]; if (x && !x.admin && !x.suspendu) cnt[x.ou] = (cnt[x.ou] || 0) + 1; });
  var best = "/", n = 0; Object.keys(cnt).forEach(function (k) { if (cnt[k] > n) { n = cnt[k]; best = k; } });
  c.put("cpt_ou_etu", best, 86400); return best;
}
function _cptEtatWs_(e) {
  if (!e.mailHecg) return { statut: "sans adresse" };
  try { var u = _wsFetch_("get", "/users/" + encodeURIComponent(e.mailHecg) + "?projection=basic");
    return { statut: u.suspended ? "suspendu" : "existe", ou: u.orgUnitPath || "/", cree: String(u.creationTime || "").slice(0, 10) };
  } catch (x) { return /404|not found|Resource Not Found/i.test(x.message) ? { statut: "absent" } : { statut: "erreur", message: x.message }; }
}

// ── Point d'entrée front ───────────────────────────────────────────────────
function handleComptesEtudiant(p, ss) {
  var op = String(p.op || ""), par = String(p.nomAuteur || p.emailAuteur || ""), role = p.roleAuteur || p.vieScRole || "";
  try {
    if (String(p.roleAuteurValide || "") === "Etudiant") return createJsonResponse({ success: false, message: "Accès refusé." });
    var e = p.etuId ? _cptEtudiant_(ss, p.etuId) : null;
    if (op === "etat") {
      if (!e) return createJsonResponse({ success: false, message: "Étudiant introuvable." });
      var R = { success: true, inscrit: e.inscrit, typo: e.typo, idEleve: e.idEleve, mailHecg: e.mailHecg, peutCharl: _cptPeutCharl_(p), peutWs: _cptPeutWs_(p) };
      if (R.peutCharl) R.charl = _cptEtatCharl_(ss, e.id);
      if (R.peutWs && p.avecWs) R.ws = _cptEtatWs_(e);
      return createJsonResponse(R);
    }
    if (op === "charlPreparer") {
      if (!_cptPeutCharl_(p)) return createJsonResponse({ success: false, message: "Accès réservé au personnel ayant Charlemagne." });
      if (!e) return createJsonResponse({ success: false, message: "Étudiant introuvable." });
      if (!e.inscrit) return createJsonResponse({ success: false, message: "Seuls les étudiants inscrits ou préinscrits peuvent être créés dans Charlemagne (typologie actuelle : " + (e.typo || "—") + ")." });
      if (e.idEleve) return createJsonResponse({ success: false, message: "Déjà présent dans Charlemagne (Identifiant Elève " + e.idEleve + ")." });
      var ma = e.idEnt ? _cptMa_(ss, e) : null, nai = _cptNaissance_(ss, e), doublon = _cptDoublonCharl_(ss, e);
      return createJsonResponse({ success: true, donnees: {
        nom: e.nom.toUpperCase(), prenom: e.prenom, tel: _cptTel_(e.tel), mail: e.mail, classe: e.classe, dateEntree: e.dateCrea || Utilities.formatDate(new Date(), "Europe/Paris", "dd/MM/yyyy"),
        naissanceDate: nai.date, naissanceLieu: nai.lieu, naissancePays: nai.pays, naissanceSource: nai.source,
        alternant: !!e.idEnt, entreprise: e.entreprise, famille: e.idEnt ? (ma ? { nom: ma.nom.toUpperCase(), prenom: ma.prenom } : null) : { nom: "SANS", prenom: "ENTREPRISE" },
        photo: e.nom.toUpperCase() + " " + e.prenom }, doublon: doublon });
    }
    if (op === "charlDemander") {
      if (!_cptPeutCharl_(p)) return createJsonResponse({ success: false, message: "Accès réservé au personnel ayant Charlemagne." });
      if (!e || !e.inscrit) return createJsonResponse({ success: false, message: "Étudiant introuvable ou non inscrit / préinscrit." });
      var D = typeof p.donnees === "string" ? JSON.parse(p.donnees) : (p.donnees || {});
      if (!D.nom || !D.prenom || !D.classe) return createJsonResponse({ success: false, message: "Nom, prénom et classe sont obligatoires." });
      if (!/^\d{2}\/\d{2}\/\d{4}$/.test(String(D.naissanceDate || ""))) return createJsonResponse({ success: false, message: "Date de naissance obligatoire (jj/mm/aaaa)." });
      if (D.alternant && !(D.famille && D.famille.nom)) return createJsonResponse({ success: false, message: "Alternant : indiquez le nom du maître d'apprentissage (fiche famille)." });
      var lock = LockService.getScriptLock(); lock.waitLock(15000);
      try {
        var enCours = _cptFile_(ss).filter(function (x) { return x.etuId === e.id && (x.statut === "En attente" || x.statut === "En cours"); });
        if (enCours.length) return createJsonResponse({ success: false, message: "Une création est déjà en attente pour cet étudiant." });
        delete D.naissanceSource;
        var id = "CH-" + Date.now();
        _cptSheet_(ss).appendRow([id, new Date(), e.id, D.nom, D.prenom, D.classe, "En attente", "", "", par, "", "", JSON.stringify(D), ""]);
        logAction(par, role, "Création", COMPTES_CAT + " — Charlemagne", e.id, "Création Charlemagne demandée pour " + D.prenom + " " + D.nom + " (" + D.classe + ", " + (D.alternant ? "alternant · famille " + D.famille.nom : "SANS ENTREPRISE") + ")");
      } finally { lock.releaseLock(); }
      return createJsonResponse({ success: true, charl: _cptEtatCharl_(ss, e.id) });
    }
    if (op === "charlAnnuler") {
      if (!_cptPeutCharl_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
      var j = _cptFile_(ss).filter(function (x) { return x.id === String(p.jobId); })[0];
      if (!j || j.statut !== "En attente") return createJsonResponse({ success: false, message: "Demande introuvable ou déjà prise en charge par l'agent." });
      _cptSheet_(ss).getRange(j.row, 7, 1, 3).setValues([["Annulé", "", "Annulé par " + par]]);
      logAction(par, role, "Suppression", COMPTES_CAT + " — Charlemagne", j.etuId, "Demande de création Charlemagne annulée");
      return createJsonResponse({ success: true, charl: _cptEtatCharl_(ss, j.etuId) });
    }
    if (op === "charlIdManuel") {   // création faite à la main : on renseigne l'identifiant Charlemagne
      if (!_cptPeutCharl_(p) || !e) return createJsonResponse({ success: false, message: "Accès refusé." });
      var idm = String(p.idCharl || "").replace(/\D/g, ""); if (!idm) return createJsonResponse({ success: false, message: "Identifiant Charlemagne invalide." });
      getSheetSafe(ss, "ETUDIANTS").getRange(e.row, 24).setValue(idm);
      try { _invalidateStudentDetailsCache_(e.id); } catch (eC) {}
      logAction(par, role, "Modification", COMPTES_CAT + " — Charlemagne", e.id, "Identifiant Charlemagne saisi à la main : " + idm);
      return createJsonResponse({ success: true });
    }
    if (op === "agentTelecharger") {
      if (!_cptPeutCharl_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
      return createJsonResponse({ success: true, cle: _cptCleAgent_(), poste: "" });
    }
    if (op === "wsEtat" || op === "wsCreer") {
      if (!_cptPeutWs_(p)) return createJsonResponse({ success: false, message: "Création des comptes Google Workspace réservée aux profils Super-admin." });
      if (!e) return createJsonResponse({ success: false, message: "Étudiant introuvable." });
      if (op === "wsEtat") return createJsonResponse({ success: true, ws: _cptEtatWs_(e), mailHecg: e.mailHecg });
      if (!e.inscrit) return createJsonResponse({ success: false, message: "Seuls les étudiants inscrits ou préinscrits peuvent recevoir un compte." });
      if (!/^[a-z0-9._%+-]+@hecg\.fr$/.test(e.mailHecg)) return createJsonResponse({ success: false, message: "Renseignez d'abord l'adresse mail domaine (@hecg.fr) dans la fiche." });
      var et = _cptEtatWs_(e);
      if (et.statut === "existe" || et.statut === "suspendu") return createJsonResponse({ success: false, message: "Le compte " + e.mailHecg + " existe déjà" + (et.statut === "suspendu" ? " (suspendu)" : "") + "." });
      if (et.statut === "erreur") return createJsonResponse({ success: false, message: "Google Workspace : " + et.message });
      var ou = _cptOuEtudiants_(ss), mdp = _cptMdp_();
      _wsFetch_("post", "/users", { primaryEmail: e.mailHecg, name: { givenName: e.prenom, familyName: e.nom.toUpperCase() }, password: mdp, changePasswordAtNextLogin: true, orgUnitPath: ou });
      try { _wsJournal_(ss, [[new Date(), e.mailHecg, e.prenom + " " + e.nom, "Création", "", ou, e.id, par, "OK"]]); } catch (eJ) {}
      logAction(par, role, "Création", COMPTES_CAT + " — Google Workspace", e.id, "Compte " + e.mailHecg + " créé pour " + e.prenom + " " + e.nom + " (unité " + ou + ", mot de passe provisoire à changer à la 1re connexion)");
      return createJsonResponse({ success: true, email: e.mailHecg, motDePasse: mdp, ou: ou });
    }
    return createJsonResponse({ success: false, message: "Opération inconnue : " + op });
  } catch (x) { return createJsonResponse({ success: false, message: x.message || String(x) }); }
}

// ── Agent Charlemagne (PC local) : action publique protégée par une clé ────
function _cptCleAgent_() {
  var pr = PropertiesService.getScriptProperties(), k = pr.getProperty("CHARL_AGENT_CLE");
  if (!k) { k = Utilities.getUuid().replace(/-/g, "") + Utilities.getUuid().replace(/-/g, ""); pr.setProperty("CHARL_AGENT_CLE", k); }
  return k;
}
function _cptCleOk_(k) {
  var ref = PropertiesService.getScriptProperties().getProperty("CHARL_AGENT_CLE") || "";
  if (!ref || !k || String(k).length !== ref.length) return false;
  var diff = 0; for (var i = 0; i < ref.length; i++) diff |= ref.charCodeAt(i) ^ String(k).charCodeAt(i);
  return diff === 0;
}
function handleCharlAgent(p, ss) {
  if (!_cptCleOk_(p.cle)) return createJsonResponse({ success: false, message: "Clé d'agent invalide." });
  var op = String(p.op || ""), poste = String(p.poste || "PC").slice(0, 60), ws = _cptSheet_(ss);
  var lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    var L = _cptFile_(ss), maintenant = new Date();
    if (op === "suivant") {
      // une demande « En cours » depuis plus de 20 min (agent arrêté) repasse en attente
      L.forEach(function (x) { if (x.statut === "En cours") { var t = ws.getRange(x.row, 11).getValue(); if (t instanceof Date && maintenant - t > 20 * 60000) ws.getRange(x.row, 7, 1, 3).setValues([["En attente", "", "Reprise après interruption de l'agent"]]); } });
      L = _cptFile_(ss);
      var j = L.filter(function (x) { return x.statut === "En attente"; })[0];
      if (!j) return createJsonResponse({ success: true, job: null });
      ws.getRange(j.row, 7).setValue("En cours"); ws.getRange(j.row, 11, 1, 2).setValues([[maintenant, poste]]);
      return createJsonResponse({ success: true, job: { id: j.id, etuId: j.etuId, donnees: JSON.parse(j.donnees || "{}") } });
    }
    if (op === "resultat") {
      var r = L.filter(function (x) { return x.id === String(p.jobId); })[0];
      if (!r) return createJsonResponse({ success: false, message: "Demande introuvable." });
      var statut = p.ok === true || String(p.ok) === "true" ? "Créé" : (String(p.statut || "") === "Existe déjà" ? "Existe déjà" : "Échec");
      var idc = String(p.idCharl || "").replace(/\D/g, "");
      ws.getRange(r.row, 7, 1, 3).setValues([[statut, idc, String(p.message || "").slice(0, 500)]]);
      ws.getRange(r.row, 11, 1, 2).setValues([[maintenant, poste]]); ws.getRange(r.row, 14).setValue(String(p.photo || ""));
      if (idc && (statut === "Créé" || statut === "Existe déjà")) {
        var e = _cptEtudiant_(ss, r.etuId); if (e && !e.idEleve) { getSheetSafe(ss, "ETUDIANTS").getRange(e.row, 24).setValue(idc); try { _invalidateStudentDetailsCache_(e.id); } catch (eC) {} }
      }
      logAction("Agent Charlemagne (" + poste + ")", "Auto", statut === "Échec" ? "Erreur" : "Création", COMPTES_CAT + " — Charlemagne", r.etuId,
        r.prenom + " " + r.nom + " : " + statut + (idc ? " (Id " + idc + ")" : "") + (p.message ? " — " + p.message : "") + (p.photo ? " · photo : " + p.photo : ""));
      return createJsonResponse({ success: true });
    }
    if (op === "ping") return createJsonResponse({ success: true, enAttente: L.filter(function (x) { return x.statut === "En attente"; }).length });
    return createJsonResponse({ success: false, message: "Opération inconnue." });
  } finally { try { lock.releaseLock(); } catch (eL) {} }
}

// Doublon possible : même nom + prénom dans le dernier export Charlemagne importé (Vie scolaire › Contrôle Charlemagne)
function _cptDoublonCharl_(ss, e) {
  try {
    var ws = ss.getSheetByName("CHARLEMAGNE_IMPORT"); if (!ws) return "";
    var d = ws.getDataRange().getValues(), n = _vsuNorm_(e.nom), pr = _vsuNorm_(e.prenom);
    for (var i = 1; i < d.length; i++) {
      var cells = d[i].map(function (x) { return _vsuNorm_(x); });
      if (cells.indexOf(n) !== -1 && cells.indexOf(pr) !== -1) return "Un élève « " + e.nom.toUpperCase() + " " + e.prenom + " » figure déjà dans le dernier export Charlemagne importé : vérifiez avant de créer.";
    }
  } catch (x) {}
  return "";
}
