// ════════════════════════════════════════════════════════════════════════════
//  COMPTES WORKSPACE DES ÉTUDIANTS SORTIS (compte maître uniquement)
//  Rapprochement CRM (onglet ETUDIANTS, typologie « SORTI… ») ↔ annuaire Google Workspace :
//  1) par l'adresse @hecg.fr de la fiche (colonne AA) — sûr ; 2) à défaut par le nom, si un seul compte
//  étudiant porte ce nom — à vérifier. Un compte relié AUSSI à une fiche inscrite / préinscrite est bloqué.
//  Actions : suspendre (réversible, + déplacement dans l'unité « Anciens étudiants »), réactiver,
//  supprimer (définitif, uniquement un compte déjà suspendu). Tout est journalisé (WORKSPACE_ACTIONS + mouchard).
//  Droits utilisés : admin.directory.user et admin.directory.orgunit du compte du CRM (administrateur).
// ════════════════════════════════════════════════════════════════════════════
var WS_OU_ANCIENS = "/Anciens étudiants";
var WS_ACT_HEADERS = ["Date", "Compte", "Nom", "Action", "Unité avant", "Unité après", "Fiche CRM", "Par", "Résultat"];
var WS_API = "https://admin.googleapis.com/admin/directory/v1";

function _wsFetch_(methode, chemin, corps) {
  var o = { method: methode, headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }, muteHttpExceptions: true };
  if (corps) { o.contentType = "application/json"; o.payload = JSON.stringify(corps); }
  var r = UrlFetchApp.fetch(WS_API + chemin, o), c = r.getResponseCode();
  if (c >= 300) throw new Error("Google Workspace (" + c + ") : " + (function () { try { return JSON.parse(r.getContentText()).error.message; } catch (e) { return r.getContentText().slice(0, 200); } })());
  return r.getContentText() ? JSON.parse(r.getContentText()) : {};
}
function _wsComptes_() {
  var out = [], page = "";
  do {
    var j = _wsFetch_("get", "/users?customer=my_customer&maxResults=500&projection=basic&orderBy=email" + (page ? "&pageToken=" + page : ""));
    (j.users || []).forEach(function (u) {
      out.push({ email: String(u.primaryEmail).toLowerCase(), nom: (u.name && u.name.fullName) || "", prenomWs: (u.name && u.name.givenName) || "", nomWs: (u.name && u.name.familyName) || "",
        ou: u.orgUnitPath || "/", suspendu: !!u.suspended, admin: !!u.isAdmin, derniere: u.lastLoginTime && !/^1970/.test(u.lastLoginTime) ? String(u.lastLoginTime).slice(0, 10) : "",
        alias: (u.aliases || []).map(function (a) { return String(a).toLowerCase(); }) });
    });
    page = j.nextPageToken || "";
  } while (page);
  return out;
}
function _wsEstEtudiantOu_(ou) { return /etudiant|étudiant|anciens/i.test(String(ou || "")); }
function _wsJournal_(ss, lignes) {
  if (!lignes.length) return;
  var ws = _vsxSheet_(ss, "WORKSPACE_ACTIONS", WS_ACT_HEADERS);
  ws.getRange(ws.getLastRow() + 1, 1, lignes.length, WS_ACT_HEADERS.length).setValues(lignes);
}
function _wsOuAnciens_() {
  try { _wsFetch_("get", "/customer/my_customer/orgunits" + encodeURI(WS_OU_ANCIENS).replace(/^\/?/, "/")); return; } catch (e) {}
  _wsFetch_("post", "/customer/my_customer/orgunits", { name: WS_OU_ANCIENS.replace(/^\//, ""), parentOrgUnitPath: "/", description: "Comptes suspendus des étudiants sortis (CRM HECG)" });
}

function _wsApercu_(ss) {
  var comptes = _wsComptes_(), parMail = {}, parNom = {};
  comptes.forEach(function (c) {
    parMail[c.email] = c; c.alias.forEach(function (a) { parMail[a] = c; });
    if (!_wsEstEtudiantOu_(c.ou) || c.admin) return;
    [_vsuNorm_(c.prenomWs + " " + c.nomWs), _vsuNorm_(c.nomWs + " " + c.prenomWs), _vsuNorm_(c.nom)].forEach(function (k) { if (k) { parNom[k] = parNom[k] || []; if (parNom[k].indexOf(c) === -1) parNom[k].push(c); } });
  });
  var etus = _vsuEtudiants_(ss), actifs = {}, vus = {};
  // comptes rattachés à une fiche inscrite / préinscrite : jamais touchés
  etus.forEach(function (e) {
    if (!e.inscrit) return;
    if (e.mailHecg && parMail[e.mailHecg]) actifs[parMail[e.mailHecg].email] = e.prenom + " " + e.nom + " (" + e.typo + ")";
    var l = parNom[_vsuNorm_(e.prenom + " " + e.nom)] || [];
    if (l.length === 1 && !actifs[l[0].email]) actifs[l[0].email] = e.prenom + " " + e.nom + " (" + e.typo + ", même nom)";
  });
  var items = [];
  etus.forEach(function (e) {
    if (!/^sorti/i.test(e.typo)) return;
    var c = null, rappr = "", doute = "";
    if (e.mailHecg && parMail[e.mailHecg]) { c = parMail[e.mailHecg]; rappr = "adresse"; }
    else {
      var l = parNom[_vsuNorm_(e.prenom + " " + e.nom)] || [];
      if (l.length === 1) { c = l[0]; rappr = "nom"; doute = "rapproché par le nom seulement" + (e.mailHecg ? " (l'adresse de la fiche, " + e.mailHecg + ", n'existe pas dans Workspace)" : ""); }
      else if (l.length > 1) { items.push({ email: "", etuId: e.id, nomCrm: e.prenom + " " + e.nom, typo: e.typo, classe: e.classe, rappr: "homonymes", doute: l.length + " comptes portent ce nom : " + l.map(function (x) { return x.email; }).join(", "), bloque: true }); return; }
    }
    if (!c) return;   // pas de compte Workspace : rien à faire
    if (c.admin) return;
    var it = vus[c.email];
    if (it) { it.fiches.push(e.id); return; }   // plusieurs fiches sorties pour le même compte
    it = vus[c.email] = { email: c.email, nomWs: c.nom, ou: c.ou, suspendu: c.suspendu, derniere: c.derniere, etuId: e.id, fiches: [e.id], nomCrm: e.prenom + " " + e.nom, typo: e.typo, classe: e.classe,
      rappr: rappr, doute: doute, bloque: false };
    if (actifs[c.email]) { it.bloque = true; it.doute = "compte aussi relié à une fiche ACTIVE : " + actifs[c.email]; }
    else if (!_wsEstEtudiantOu_(c.ou)) { it.bloque = true; it.doute = "compte hors des unités étudiantes (" + c.ou + ") : traitement manuel"; }
    items.push(it);
  });
  // comptes étudiants actifs sans aucune fiche dans le CRM (information)
  var connus = {}; etus.forEach(function (e) { if (e.mailHecg) connus[e.mailHecg] = 1; connus["n:" + _vsuNorm_(e.prenom + " " + e.nom)] = 1; });
  var orphelins = comptes.filter(function (c) { return _wsEstEtudiantOu_(c.ou) && c.ou !== WS_OU_ANCIENS && !c.suspendu && !c.admin && !connus[c.email]
    && !connus["n:" + _vsuNorm_(c.prenomWs + " " + c.nomWs)] && !connus["n:" + _vsuNorm_(c.nom)]; })
    .map(function (c) { return { email: c.email, nomWs: c.nom, ou: c.ou, derniere: c.derniere }; });
  return { items: items, orphelins: orphelins, total: comptes.length };
}

function handleWorkspaceComptes(p, ss) {
  if (!_fp2Autorise_(p) && !_pilAutorise_(p)) return createJsonResponse({ success: false, message: "Module réservé à m.mokhtari@hecg.fr et a.mhoudini@hecg.fr." });
  var par = String(p.nomAuteur || p.emailAuteur || ""), op = String(p.op || "");
  try {
    if (op === "apercu") {
      var A = _wsApercu_(ss);
      logAction(par, p.roleAuteur || "", "Consultation", "Comptes Workspace — étudiants sortis", "annuaire",
        "Rapprochement : " + A.items.length + " sorti(s) avec compte (" + A.items.filter(function (x) { return !x.bloque && !x.suspendu; }).length + " actif(s) traitable(s)), " + A.orphelins.length + " compte(s) étudiant(s) sans fiche CRM");
      return createJsonResponse({ success: true, items: A.items, orphelins: A.orphelins, total: A.total, ouAnciens: WS_OU_ANCIENS });
    }
    if (op === "suspendre" || op === "reactiver" || op === "supprimer") {
      var emails = (p.emails || []).map(function (x) { return String(x).trim().toLowerCase(); }).filter(Boolean).slice(0, 15);
      if (op === "supprimer" && String(p.confirmation || "") !== "SUPPRIMER") return createJsonResponse({ success: false, message: "Confirmation manquante." });
      // garde-fou : on revérifie le rapprochement côté serveur (jamais un compte bloqué, jamais un compte hors liste)
      var A2 = _wsApercu_(ss), ok = {}; A2.items.forEach(function (x) { if (x.email && !x.bloque) ok[x.email] = x; });
      if (op === "reactiver") A2.items.forEach(function (x) { if (x.email) ok[x.email] = x; });
      if (op === "suspendre" && p.deplacer !== false) _wsOuAnciens_();
      var res = [], lignes = [];
      emails.forEach(function (em) {
        var x = ok[em];
        if (!x) { res.push({ email: em, statut: "refusé", detail: "compte non relié à une fiche sortie (ou bloqué)" }); return; }
        try {
          var u = _wsFetch_("get", "/users/" + encodeURIComponent(em) + "?projection=basic"), avant = u.orgUnitPath || "/", apres = avant;
          if (op === "suspendre") {
            var corps = { suspended: true }; if (p.deplacer !== false) { corps.orgUnitPath = WS_OU_ANCIENS; apres = WS_OU_ANCIENS; }
            _wsFetch_("put", "/users/" + encodeURIComponent(em), corps);
            res.push({ email: em, statut: "ok", detail: "suspendu" + (apres !== avant ? " · déplacé dans " + apres : "") });
          } else if (op === "reactiver") {
            // retour dans l'unité d'origine (dernière suspension journalisée)
            var wsJ = ss.getSheetByName("WORKSPACE_ACTIONS"), orig = "";
            if (wsJ && wsJ.getLastRow() > 1) { var dJ = wsJ.getRange(2, 1, wsJ.getLastRow() - 1, 9).getValues(); for (var k = dJ.length - 1; k >= 0; k--) if (String(dJ[k][1]).toLowerCase() === em && dJ[k][3] === "Suspension") { orig = String(dJ[k][4] || ""); break; } }
            var c2 = { suspended: false }; if (orig && orig !== avant) { c2.orgUnitPath = orig; apres = orig; }
            _wsFetch_("put", "/users/" + encodeURIComponent(em), c2);
            res.push({ email: em, statut: "ok", detail: "réactivé" + (apres !== avant ? " · remis dans " + apres : "") });
          } else {
            if (!u.suspended) { res.push({ email: em, statut: "refusé", detail: "à suspendre d'abord (la suppression ne vise que des comptes déjà suspendus)" }); return; }
            _wsFetch_("delete", "/users/" + encodeURIComponent(em));
            apres = "(supprimé — récupérable 20 jours dans la console d'administration)";
            res.push({ email: em, statut: "ok", detail: "supprimé" });
          }
          lignes.push([new Date(), em, x.nomCrm || x.nomWs || "", op === "suspendre" ? "Suspension" : op === "reactiver" ? "Réactivation" : "Suppression", avant, apres, (x.fiches || [x.etuId]).join(", "), par, "OK"]);
        } catch (eU) { res.push({ email: em, statut: "erreur", detail: String(eU.message || eU).slice(0, 200) }); lignes.push([new Date(), em, x.nomCrm || "", op, "", "", x.etuId || "", par, "ÉCHEC : " + String(eU.message || eU).slice(0, 200)]); }
      });
      _wsJournal_(ss, lignes);
      logAction(par, p.roleAuteur || "", op === "supprimer" ? "Suppression" : "Modification", "Comptes Workspace — étudiants sortis", res.length + " compte(s)",
        ({ suspendre: "Suspension", reactiver: "Réactivation", supprimer: "SUPPRESSION DÉFINITIVE" })[op] + " — " + res.map(function (r) { return r.email + " : " + r.statut + (r.statut === "ok" ? " (" + r.detail + ")" : " (" + r.detail + ")"); }).join(" ; "));
      return createJsonResponse({ success: true, resultats: res });
    }
    return createJsonResponse({ success: false, message: "Opération inconnue : " + op });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}
// À lancer UNE fois dans l'éditeur : autorise la gestion des comptes et des unités (droits admin.directory.user / orgunit)
function autoriserGestionComptesWorkspace() { Logger.log(_wsComptes_().length + " comptes lus ; unités : " + (_wsFetch_("get", "/customer/my_customer/orgunits?type=all").organizationUnits || []).length); }
