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
  // largeur de référence : images et blocs du reste de la signature (icônes < 120 px ignorées)
  sans.replace(/<img[^>]*>/gi, function (t) { var w = t.match(/\bwidth=["']?(\d+)/i) || t.match(/[;"'\s]width:\s*(\d+)px/i); if (w && +w[1] >= 120) largeurs.push(+w[1]); return t; });
  sans.replace(/max-width:\s*(\d+)px/gi, function (t, w) { if (+w >= 120) largeurs.push(+w); return t; });
  var larg = largeurs.length ? Math.min(Math.max.apply(null, largeurs), 650) : SIG_LARGEUR_DEFAUT;
  return { anciennes: trouvees, sans: sans.replace(/(\s*<br\s*\/?>\s*)+$/i, ""), largeur: larg, largeurs: largeurs };
}
function _sigBloc_(url, larg) {
  return '<div style="margin-top:12px"><a href="' + SIG_LIEN_SITE + '" target="_blank"><img src="' + url + '" alt="Bannière HECG" width="' + larg
    + '" style="display:block;width:' + larg + 'px;max-width:100%;height:auto;border:0"></a></div>';
}
// Jeton Gmail « paramètres » au nom d'un membre du domaine (compte de service robot-signature-hecg, délégation domaine)
function _sigJeton_(email) {
  var b64 = function (x) { return Utilities.base64EncodeWebSafe(x).replace(/=+$/, ""); }, now = Math.floor(Date.now() / 1000);
  var entree = b64(JSON.stringify({ alg: "RS256", typ: "JWT" })) + "." + b64(JSON.stringify({ iss: SIG_SA_EMAIL, sub: email, aud: "https://oauth2.googleapis.com/token",
    scope: "https://www.googleapis.com/auth/gmail.settings.basic https://www.googleapis.com/auth/gmail.settings.sharing", iat: now, exp: now + 3600 }));
  var r = UrlFetchApp.fetch("https://oauth2.googleapis.com/token", { method: "post", muteHttpExceptions: true,
    payload: { grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer", assertion: entree + "." + b64(Utilities.computeRsaSha256Signature(entree, SIG_SA_KEY)) } });
  var j = JSON.parse(r.getContentText() || "{}");
  if (!j.access_token) throw new Error("Autorisation Gmail refusée pour " + email + " : " + (j.error_description || j.error || r.getResponseCode()));
  return j.access_token;
}
function _sigLireMienne_() {
  var r = UrlFetchApp.fetch("https://gmail.googleapis.com/gmail/v1/users/me/settings/sendAs", { headers: { Authorization: "Bearer " + _sigJeton_(FP2_MAITRE) }, muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error("Lecture de la signature impossible (" + r.getResponseCode() + ") : " + r.getContentText().slice(0, 200));
  var l = JSON.parse(r.getContentText()).sendAs || [], p = l.filter(function (a) { return a.isDefault; })[0] || l[0];
  if (!p) throw new Error("Aucune adresse d'envoi trouvée.");
  return { email: p.sendAsEmail, signature: p.signature || "" };
}

function handleSignatureHecg(p, ss) {
  if (!_fp2Autorise_(p)) return createJsonResponse({ success: false, message: "Module réservé à m.mokhtari@hecg.fr." });
  try {
    var op = String(p.op || ""), props = PropertiesService.getScriptProperties();
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
      logAction(p.nomAuteur || p.emailAuteur, p.roleAuteur || "", "Consultation", "Communication — Signature HECG", moi.email,
        "Aperçu de la bannière sur la signature de " + moi.email + " (aucune modification Gmail) — image : " + lien + " ; largeur " + A.largeur + " px ; " + A.anciennes.length + " ancienne(s) bannière(s) remplacée(s)");
      return createJsonResponse({ success: true, email: moi.email, actuelle: moi.signature, nouvelle: nouvelle, imageUrl: img.url, largeur: A.largeur,
        anciennes: A.anciennes.length, alertes: alertes, vide: !moi.signature });
    }
    return createJsonResponse({ success: false, message: "Opération inconnue : " + op });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}
