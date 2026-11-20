// ════════════════════════════════════════════════════════════════════════════
//  COMMUNICATION — SIGNATURE HECG (bannière sous les signatures Gmail)
//  Étape 1 (actuelle) : APERÇU sur la signature de m.mokhtari@hecg.fr, lue via le compte
//  de service robot-signature-hecg (clé dans _secrets.gs). Aucune écriture dans Gmail à ce stade.
//  La bannière est ajoutée SOUS la signature existante, cliquable vers https://hecg.fr ;
//  une bannière HECG déjà présente (ancien script) est remplacée, rien d'autre n'est touché.
// ════════════════════════════════════════════════════════════════════════════
var SIG_LIEN_SITE = "https://hecg.fr/";
var SIG_LARGEUR_DEFAUT = 500;

// Lien collé (Drive « partager », Drive « open?id= », lien direct) → URL d'image affichable dans Gmail
function _sigUrlImage_(lien) {
  var s = String(lien || "").trim(), m = s.match(/drive\.google\.com\/(?:file\/d\/|open\?id=|uc\?(?:export=\w+&)?id=|thumbnail\?id=)([\w-]{20,})/) || s.match(/[?&]id=([\w-]{20,})/);
  if (m && /google\.com/.test(s)) return { url: "https://drive.google.com/thumbnail?id=" + m[1] + "&sz=w1000", driveId: m[1] };
  return { url: s, driveId: "" };
}
// Bannières HECG déjà présentes : <a href="hecg.fr"><img alt="HECG"|"Bannière HECG"|src Drive></a>, avec son <div> d'enveloppe s'il ne contient qu'elle
var SIG_RE_BANNIERE = /(<div[^>]*>\s*)?<a[^>]*href=["']https?:\/\/(?:www\.)?hecg\.fr\/?["'][^>]*>\s*<img(?=[^>]*(?:alt=["'](?:HECG|Bannière HECG)["']|drive\.google\.com\/thumbnail))[^>]*>\s*<\/a>(\s*<\/div>)?/gi;

function _sigAnalyser_(html) {
  var trouvees = [], largeurs = [];
  String(html).replace(SIG_RE_BANNIERE, function (m) { trouvees.push(m); return m; });
  var sans = String(html).replace(SIG_RE_BANNIERE, function (m, d, f) { return d && !f ? d : (!d && f ? f : ""); });
  // largeur de référence : la plus grande image de la signature (icônes < 120 px ignorées), sinon les blocs
  sans.replace(/<img[^>]*>/gi, function (t) { var w = t.match(/\bwidth=["']?(\d+)/i) || t.match(/[;"'\s]width:\s*(\d+)px/i); if (w && +w[1] >= 120) largeurs.push(+w[1]); return t; });
  if (!largeurs.length) sans.replace(/max-width:\s*(\d+)px/gi, function (t, w) { if (+w >= 120) largeurs.push(+w); return t; });   // à défaut d'image : largeur des blocs
  var larg = largeurs.length ? Math.min(Math.max.apply(null, largeurs), 650) : SIG_LARGEUR_DEFAUT;
  return { anciennes: trouvees, sans: sans.replace(/(\s*<br\s*\/?>\s*)+$/i, ""), largeur: larg, largeurs: largeurs };
}
function _sigBloc_(url, larg) {
  return '<div style="margin-top:12px"><a href="' + SIG_LIEN_SITE + '" target="_blank"><img src="' + url + '" alt="Bannière HECG" width="' + larg
    + '" style="display:block;width:' + larg + 'px;max-width:100%;height:auto;border:0"></a></div>';
}
// Jeton au nom d'un membre du domaine (compte de service robot-signature-hecg, délégation domaine)
var SIG_SCOPE_GMAIL = "https://www.googleapis.com/auth/gmail.settings.basic https://www.googleapis.com/auth/gmail.settings.sharing";
function _sigJeton_(email, scope) {
  var b64 = function (x) { return Utilities.base64EncodeWebSafe(x).replace(/=+$/, ""); }, now = Math.floor(Date.now() / 1000);
  var entree = b64(JSON.stringify({ alg: "RS256", typ: "JWT" })) + "." + b64(JSON.stringify({ iss: SIG_SA_EMAIL, sub: email, aud: "https://oauth2.googleapis.com/token",
    scope: scope || SIG_SCOPE_GMAIL, iat: now, exp: now + 3600 }));
  var r = UrlFetchApp.fetch("https://oauth2.googleapis.com/token", { method: "post", muteHttpExceptions: true,
    payload: { grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer", assertion: entree + "." + b64(Utilities.computeRsaSha256Signature(entree, SIG_SA_KEY)) } });
  var j = JSON.parse(r.getContentText() || "{}");
  if (!j.access_token) throw new Error("Autorisation refusée pour " + email + " : " + (j.error_description || j.error || r.getResponseCode()));
  return j.access_token;
}
// Adresse d'envoi par défaut d'un membre (+ jeton pour la réécrire)
function _sigLire_(email) {
  var tok = _sigJeton_(email);
  var r = UrlFetchApp.fetch("https://gmail.googleapis.com/gmail/v1/users/me/settings/sendAs", { headers: { Authorization: "Bearer " + tok }, muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error("Lecture de la signature impossible (" + r.getResponseCode() + ") : " + r.getContentText().slice(0, 200));
  var l = JSON.parse(r.getContentText()).sendAs || [], d = l.filter(function (a) { return a.isDefault; })[0] || l[0];
  if (!d) throw new Error("Aucune adresse d'envoi trouvée.");
  return { email: d.sendAsEmail, signature: d.signature || "", jeton: tok };
}
function _sigLireMienne_() { return _sigLire_(FP2_MAITRE); }
function _sigEcrire_(S, html) {
  var r = UrlFetchApp.fetch("https://gmail.googleapis.com/gmail/v1/users/me/settings/sendAs/" + encodeURIComponent(S.email), { method: "patch", contentType: "application/json",
    headers: { Authorization: "Bearer " + S.jeton }, payload: JSON.stringify({ signature: html }), muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error("Écriture refusée (" + r.getResponseCode() + ") : " + r.getContentText().slice(0, 200));
}
// Membres du domaine : annuaire Workspace (délégation) ; à défaut, les adresses @hecg.fr de l'onglet PERSONNEL
function _sigMembres_(ss) {
  try {
    var tok = _sigJeton_(FP2_MAITRE, "https://www.googleapis.com/auth/admin.directory.user.readonly"), out = [], page = "";
    do {
      var r = UrlFetchApp.fetch("https://admin.googleapis.com/admin/directory/v1/users?customer=my_customer&maxResults=200&orderBy=email" + (page ? "&pageToken=" + page : ""),
        { headers: { Authorization: "Bearer " + tok }, muteHttpExceptions: true });
      if (r.getResponseCode() !== 200) throw new Error(r.getContentText().slice(0, 150));
      var j = JSON.parse(r.getContentText());
      (j.users || []).forEach(function (u) { out.push({ email: String(u.primaryEmail).toLowerCase(), nom: (u.name && u.name.fullName) || "", suspendu: !!(u.suspended || u.archived) }); });
      page = j.nextPageToken || "";
    } while (page);
    return { source: "Annuaire Google Workspace", membres: out };
  } catch (e) {
    var dP = getPersonnelDataCached_(ss), vus = {}, l = [];
    for (var i = 1; i < dP.length; i++) dP[i].forEach(function (v) { var m = String(v || "").trim().toLowerCase(); if (/^[^@\s]+@hecg\.fr$/.test(m) && !vus[m]) { vus[m] = 1; l.push({ email: m, nom: (String(dP[i][2] || "") + " " + String(dP[i][1] || "")).trim(), suspendu: false }); } });
    if (!vus[FP2_MAITRE]) l.push({ email: FP2_MAITRE, nom: "", suspendu: false });
    return { source: "Onglet PERSONNEL (annuaire Workspace non accessible : " + String(e.message || e).slice(0, 120) + ")", membres: l.sort(function (a, b) { return a.email < b.email ? -1 : 1; }) };
  }
}
var SIG_SAUV_HEADERS = ["Date", "Membre", "Adresse d'envoi", "Signature avant", "Signature après", "Bannière", "Par", "Action"];
function _sigDate_(d) { try { return Utilities.formatDate(new Date(d), "Europe/Paris", "dd/MM/yyyy HH:mm"); } catch (e) { return String(d); } }

function handleSignatureHecg(p, ss) {
  if (!_fp2Autorise_(p)) return createJsonResponse({ success: false, message: "Module réservé à m.mokhtari@hecg.fr." });
  try {
    var op = String(p.op || ""), props = PropertiesService.getScriptProperties(), par = String(p.nomAuteur || p.emailAuteur || "");
    if (op === "etat") return createJsonResponse({ success: true, banniere: props.getProperty("SIG_BANNIERE") || "" });
    if (op === "apercu") {
      var lien = String(p.lien || "").trim();
      if (!/^https?:\/\//i.test(lien)) return createJsonResponse({ success: false, message: "Collez le lien complet de l'image (https://…)." });
      var img = _sigUrlImage_(lien), alertes = [];
      if (img.driveId) {
        try { var f = DriveApp.getFileById(img.driveId), acc = f.getSharingAccess();
          if (!/^image\//.test(f.getMimeType())) alertes.push("Le fichier Drive n'est pas une image (" + f.getMimeType() + ").");
          if (acc !== DriveApp.Access.ANYONE && acc !== DriveApp.Access.ANYONE_WITH_LINK) alertes.push("L'image n'est pas partagée « Tous les utilisateurs disposant du lien » : vos destinataires extérieurs ne la verraient pas.");
        } catch (eD) { alertes.push("Fichier Drive non accessible avec le compte du CRM : vérifiez le lien et son partage."); }
      }
      var moi = _sigLireMienne_(), A = _sigAnalyser_(moi.signature);
      var nouvelle = A.sans + _sigBloc_(img.url, A.largeur);
      props.setProperty("SIG_BANNIERE", lien);
      logAction(par, p.roleAuteur || "", "Consultation", "Communication — Signature HECG", moi.email,
        "Aperçu de la bannière sur la signature de " + moi.email + " (aucune modification Gmail) — image : " + lien + " ; largeur " + A.largeur + " px ; " + A.anciennes.length + " ancienne(s) bannière(s) remplacée(s)");
      return createJsonResponse({ success: true, email: moi.email, actuelle: moi.signature, nouvelle: nouvelle, imageUrl: img.url, largeur: A.largeur,
        anciennes: A.anciennes.length, alertes: alertes, vide: !moi.signature });
    }
    if (op === "membres") {
      var M = _sigMembres_(ss), wsS = ss.getSheetByName("SIGNATURES_SAUVEGARDE"), sauv = {};
      if (wsS && wsS.getLastRow() > 1) wsS.getRange(2, 1, wsS.getLastRow() - 1, 8).getValues().forEach(function (r) { if (r[1]) sauv[String(r[1]).toLowerCase()] = { le: _sigDate_(r[0]), action: String(r[7] || "") }; });
      M.membres.forEach(function (m) { m.dernier = sauv[m.email] || null; });
      return createJsonResponse({ success: true, source: M.source, membres: M.membres, banniere: props.getProperty("SIG_BANNIERE") || "" });
    }
    if (op === "appliquer") {   // par lots de quelques membres (limite de 6 min d'Apps Script)
      var lienA = String(p.lien || props.getProperty("SIG_BANNIERE") || "").trim();
      if (!/^https?:\/\//i.test(lienA)) return createJsonResponse({ success: false, message: "Aucune bannière définie : faites d'abord l'aperçu." });
      var imgA = _sigUrlImage_(lienA), ws = _vsxSheet_(ss, "SIGNATURES_SAUVEGARDE", SIG_SAUV_HEADERS), res = [], lignes = [];
      (p.emails || []).slice(0, 8).forEach(function (em) {
        em = String(em).trim().toLowerCase();
        try {
          var S = _sigLire_(em);
          if (!String(S.signature).trim()) { res.push({ email: em, statut: "ignoré", detail: "signature vide — rien ajouté" }); return; }
          var B = _sigAnalyser_(S.signature), nv = B.sans + _sigBloc_(imgA.url, B.largeur);
          if (nv === S.signature) { res.push({ email: em, statut: "inchangé", detail: "bannière déjà à jour" }); return; }
          _sigEcrire_(S, nv);
          lignes.push([new Date(), em, S.email, S.signature, nv, lienA, par, "Bannière appliquée"]);
          res.push({ email: em, statut: "ok", detail: B.largeur + " px" + (B.anciennes.length ? " · ancienne bannière remplacée" : " · bannière ajoutée") });
        } catch (eU) { res.push({ email: em, statut: "erreur", detail: String(eU.message || eU).slice(0, 200) }); }
        Utilities.sleep(250);
      });
      if (lignes.length) ws.getRange(ws.getLastRow() + 1, 1, lignes.length, SIG_SAUV_HEADERS.length).setValues(lignes);
      logAction(par, p.roleAuteur || "", "Modification", "Communication — Signature HECG", lignes.length + " signature(s)",
        "Bannière " + lienA + " — " + res.map(function (x) { return x.email + " : " + x.statut + (x.statut === "ok" ? "" : " (" + x.detail + ")"); }).join(" ; "));
      return createJsonResponse({ success: true, resultats: res });
    }
    if (op === "restaurer") {   // remet la signature d'avant la DERNIÈRE application de la bannière
      var em2 = String(p.email || "").trim().toLowerCase(), ws2 = ss.getSheetByName("SIGNATURES_SAUVEGARDE");
      if (!ws2 || ws2.getLastRow() < 2) return createJsonResponse({ success: false, message: "Aucune sauvegarde." });
      var d = ws2.getRange(2, 1, ws2.getLastRow() - 1, 8).getValues(), k = -1;
      for (var i = d.length - 1; i >= 0; i--) if (String(d[i][1]).toLowerCase() === em2 && d[i][7] === "Bannière appliquée") { k = i; break; }
      if (k < 0) return createJsonResponse({ success: false, message: "Aucune sauvegarde pour " + em2 + "." });
      var S2 = _sigLire_(em2);
      _sigEcrire_(S2, String(d[k][3]));
      ws2.appendRow([new Date(), em2, S2.email, S2.signature, String(d[k][3]), "", par, "Restaurée (sauvegarde du " + _sigDate_(d[k][0]) + ")"]);
      logAction(par, p.roleAuteur || "", "Modification", "Communication — Signature HECG", em2, "Signature restaurée telle qu'avant la bannière (sauvegarde du " + _sigDate_(d[k][0]) + ")");
      return createJsonResponse({ success: true });
    }
    return createJsonResponse({ success: false, message: "Opération inconnue : " + op });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}
