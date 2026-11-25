// =============================================================================
//  AUDIT ESPACE RECRUTEURS (hecg.fr/espace-recruteurs) ↔ CRM — demande direction 2026-11-09
//  • Les étudiants « En recherche » (actifs) du CRM doivent figurer dans la CVthèque du site ;
//    les étudiants qui n'y sont plus attendus (en alternance, sortis) doivent en être retirés.
//  • Chaque CV du site est copié dans le Drive (dossier « CRM HECG — CV espace recruteurs »),
//    sa « Date d'édition » (bas du CV) est lue par conversion OCR Google Drive, et l'ensemble est
//    affiché dans le dossier étudiant (vert < 6 mois, orange 6-12 mois, rouge > 1 an).
//  • Mot de passe de la page : propriété de script RECRUTEURS_MDP (jamais dans le code, le dépôt
//    est public) — saisi une fois depuis Vie scolaire › Espace recruteurs.
//  • Lancement : bouton manuel + chaque nuit (6 h) via auditAdressesTick ; les CV restant à lire
//    sont traités aux passages suivants (5 min).
// =============================================================================
var RC_PAGE = "https://hecg.fr/espace-recruteurs";
var RC_COLS = { lien: "CV site — lien", drive: "CV site — copie Drive", date: "CV site — date d'édition", verifie: "CV site — vérifié le", ecart: "CV site — écart", dispo: "CV site — disponibilité" };
// « Indisponible » sur le site = profil déjà choisi (en alternance…) laissé visible volontairement : argument
// commercial auprès des recruteurs (demande direction 2026-11-10). Information réservée à l'équipe.
function _rcIndispo_(txt) { return /indispon|non\s*dispon|pas\s*dispon|alternance|plac[ée]|pourvu|recrut[ée]/i.test(String(txt || "")); }
function _rcLibDispo_(txt) { var t = String(txt || "").trim(); if (!t) return ""; return _rcIndispo_(t) ? "Indisponible" : "Disponible"; }

function _rcPeut_(p) { var r = String(p.vieScRole || ""); return r === "Super-admin" || r === "Administratif" || r === "Pédagogie" || String(p.emailAuteur || "").toLowerCase() === "m.mokhtari@hecg.fr"; }
function _rcNorm_(s) { return String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/[^a-z]+/g, " ").trim(); }
function _rcIso_(d) { return Utilities.formatDate(d, "Europe/Paris", "yyyy-MM-dd"); }

// Colonnes ajoutées à ETUDIANTS (repérées par en-tête, créées au besoin)
function _rcColonnes_(ws) {
  var lastCol = ws.getLastColumn(), h = ws.getRange(1, 1, 1, lastCol).getValues()[0].map(function (x) { return String(x || "").trim(); });
  var C = {};
  Object.keys(RC_COLS).forEach(function (k) {
    var i = h.indexOf(RC_COLS[k]);
    if (i === -1) { lastCol++; ws.getRange(1, lastCol).setValue(RC_COLS[k]).setFontWeight("bold"); h.push(RC_COLS[k]); i = lastCol - 1; }
    C[k] = i;
  });
  return C;
}
// Lecture pour le dossier étudiant (sans création de colonnes)
function _rcInfoEtudiant_(header, row) {
  var h = (header || []).map(function (x) { return String(x || "").trim(); }), o = {};
  Object.keys(RC_COLS).forEach(function (k) { var i = h.indexOf(RC_COLS[k]); o[k] = i >= 0 ? row[i] : ""; });
  if (!o.lien && !o.ecart) return null;
  var d = o.date instanceof Date ? _rcIso_(o.date) : String(o.date || "");
  var v = o.verifie instanceof Date ? Utilities.formatDate(o.verifie, "Europe/Paris", "dd/MM/yyyy") : String(o.verifie || "");
  return { lien: String(o.lien || ""), drive: String(o.drive || ""), dateEdition: d, verifieLe: v, ecart: String(o.ecart || ""), dispo: String(o.dispo || "") };
}

// ── Accès à la page protégée ─────────────────────────────────────────────────
function _rcPage_() {
  var mdp = PropertiesService.getScriptProperties().getProperty("RECRUTEURS_MDP");
  if (!mdp) throw new Error("MDP_MANQUANT");
  var r = UrlFetchApp.fetch("https://hecg.fr/wp-login.php?action=postpass", { method: "post", followRedirects: false, muteHttpExceptions: true,
    payload: { post_password: mdp, Submit: "Valider", redirect_to: RC_PAGE } });
  var h = r.getAllHeaders(), sc = h["Set-Cookie"] || h["set-cookie"] || [];
  if (!Array.isArray(sc)) sc = [sc];
  var cookie = sc.map(function (c) { return String(c).split(";")[0]; }).filter(function (c) { return /^wp-postpass_/.test(c); }).join("; ");
  if (!cookie) throw new Error("Connexion à l'espace recruteurs refusée (HTTP " + r.getResponseCode() + ").");
  var page = UrlFetchApp.fetch(RC_PAGE + "?nc=" + Date.now(), { headers: { Cookie: cookie }, muteHttpExceptions: true }).getContentText();
  if (/post-password-form/.test(page)) throw new Error("MDP_INCORRECT");
  return page;
}
// Fiches de la CVthèque : nom, diplôme, campus, statut, lien du CV.
// Statut réel d'un profil (demande direction 2026-11-25) : bouton « Lien CV » (fond bleu) = en recherche ;
// mention « PLUS DISPONIBLE » (fond gris) = plus disponible. La ligne « Disponibilité : Immédiate » n'est PAS le statut.
function _rcFiches_(html) {
  var propre = String(html || "").replace(/<(script|style|svg|noscript)[\s\S]*?<\/\1>/gi, " ")
    .replace(/<a\b[^>]*href="(https:\/\/hecg\.fr\/wp-content\/uploads\/[^"]+\.pdf)"[^>]*>/gi, "|@@PDF $1@@|");
  var seg = propre.replace(/<[^>]+>/g, "|").split("|").map(function (x) {
    return x.replace(/&nbsp;/g, " ").replace(/&#8217;|&rsquo;/g, "'").replace(/&amp;/g, "&").replace(/&#039;/g, "'").replace(/\s+/g, " ").trim();
  }).filter(Boolean);
  var idx = [];
  seg.forEach(function (s, k) { if (/^dipl[oô]me\s*:/i.test(s) && k > 0) idx.push(k); });
  var out = [], vus = {};
  idx.forEach(function (iD, n) {
    var finBloc = n + 1 < idx.length ? idx[n + 1] - 1 : Math.min(seg.length, iD + 25), bloc = seg.slice(iD, finBloc);
    var get = function (re) {
      for (var j = 0; j < bloc.length; j++) if (re.test(bloc[j])) {
        var v = bloc[j].replace(re, "").trim();
        if (!v && bloc[j + 1] && !/:\s*$/.test(bloc[j + 1]) && !/^(campus|disponibilit|@@PDF)/i.test(bloc[j + 1])) v = bloc[j + 1];
        return v;
      }
      return "";
    };
    var url = "", plus = false;
    bloc.forEach(function (s) {
      var m = s.match(/^@@PDF (\S+)@@$/);
      if (m && !url && !/dossier-de-candidature|plaquette|certificat/i.test(m[1])) url = m[1];
      if (/plus\s+disponible/i.test(s)) plus = true;
    });
    var nom = seg[iD - 1].replace(/^@@PDF \S+@@$/, "");
    if (!nom || /^@@/.test(nom)) return;
    var cle = _rcNorm_(nom) + "|" + url;
    if (vus[cle]) return; vus[cle] = 1;
    out.push({ nom: nom, diplome: get(/^dipl[oô]me\s*:\s*/i), campus: get(/^campus\s*:\s*/i), dispoTexte: get(/^disponibilit[ée]\s*:\s*/i),
      plusDispo: plus, statutSite: plus ? "PLUS DISPONIBLE" : (url ? "En recherche (lien CV)" : "Sans CV"), dispo: plus ? "Plus disponible" : "", url: url });
  });
  return out;
}

// ── Lecture de la date d'édition (OCR Google Drive) ─────────────────────────
function _rcDossier_() {
  var props = PropertiesService.getScriptProperties(), id = props.getProperty("RC_DOSSIER");
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var f = DriveApp.getFolderById(GED_ROOT_ID).createFolder("CRM HECG — CV espace recruteurs");   // dans la GED, jamais à la racine
  props.setProperty("RC_DOSSIER", f.getId());
  return f;
}
function _rcCacheFeuille_(ss) {
  var ws = ss.getSheetByName("RECRUTEURS_CV");
  if (!ws) { ws = ss.insertSheet("RECRUTEURS_CV"); ws.getRange(1, 1, 1, 6).setValues([["URL du CV (site)", "Nom affiché", "Date d'édition", "Origine de la date", "Copie Drive", "Lu le"]]).setFontWeight("bold"); ws.setFrozenRows(1); }
  return ws;
}
function _rcDateDepuisTexte_(t, url) {
  var m = String(t || "").match(/date\s*d['’]?\s*[ée]dition\s*:?\s*(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})/i);
  if (m) return { date: _rcIso_(new Date(+(m[3].length === 2 ? "20" + m[3] : m[3]), +m[2] - 1, +m[1])), origine: "Date d'édition du CV" };
  var f = decodeURIComponent(url).match(/(20\d\d)-(\d{2})-(\d{2})/);
  if (f) return { date: f[1] + "-" + f[2] + "-" + f[3], origine: "Date du nom de fichier" };
  var tous = String(t || "").match(/(\d{1,2})\/(\d{1,2})\/(20\d\d)/g), limite = Date.now() - 3 * 365 * 86400000;
  if (tous) for (var q = tous.length - 1; q >= 0; q--) {   // une date ancienne (naissance, diplôme…) n'est pas une date d'édition
    var x = tous[q].split("/"), dq = new Date(+x[2], +x[1] - 1, +x[0]);
    if (dq.getTime() > limite && dq.getTime() <= Date.now() + 86400000) return { date: _rcIso_(dq), origine: "Dernière date récente lue sur le CV" };
  }
  var u = url.match(/uploads\/(20\d\d)\/(\d{2})\//);
  if (u) return { date: u[1] + "-" + u[2] + "-01", origine: "Mois de mise en ligne (date non lisible sur le CV)" };
  return { date: "", origine: "Date non lisible" };
}
function _rcLireCv_(fiche, dossier) {
  var blob = UrlFetchApp.fetch(fiche.url, { muteHttpExceptions: true }).getBlob();
  var nomFichier = (fiche.nom || "CV") + " — CV espace recruteurs.pdf";
  var copie = dossier.createFile(blob.setName(nomFichier));
  copie.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
  var texte = "", tok = ScriptApp.getOAuthToken();
  try {
    var c = UrlFetchApp.fetch("https://www.googleapis.com/drive/v3/files/" + copie.getId() + "/copy", { method: "post", contentType: "application/json", muteHttpExceptions: true,
      headers: { Authorization: "Bearer " + tok }, payload: JSON.stringify({ name: "tmp-ocr-cv", mimeType: "application/vnd.google-apps.document" }) });
    var idDoc = JSON.parse(c.getContentText()).id;
    if (idDoc) {
      texte = UrlFetchApp.fetch("https://www.googleapis.com/drive/v3/files/" + idDoc + "/export?mimeType=text/plain", { headers: { Authorization: "Bearer " + tok }, muteHttpExceptions: true }).getContentText();
      try { DriveApp.getFileById(idDoc).setTrashed(true); } catch (eT) {}
    }
  } catch (eO) {}
  var d = _rcDateDepuisTexte_(texte, fiche.url);
  return { date: d.date, origine: d.origine, drive: copie.getUrl() };
}

// CVthèque à jour ? (date d'édition < 6 mois) — vide si la date n'est pas encore lue
function _rcEtatCvtheque_(iso) {
  var m = String(iso || "").match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return "";
  var mois = (Date.now() - new Date(+m[1], +m[2] - 1, +m[3]).getTime()) / (30.44 * 86400000);
  return mois < 6 ? "Oui" : "À actualiser";
}

// ── Audit ────────────────────────────────────────────────────────────────────
function _rcAudit_(ss, auteur, budgetMs) {
  var t0 = Date.now(), fin = t0 + (budgetMs || 240000);
  var fiches = _rcFiches_(_rcPage_());
  // Cache des CV déjà lus
  var wsC = _rcCacheFeuille_(ss), dC = wsC.getDataRange().getValues(), cache = {};
  for (var i = 1; i < dC.length; i++) if (dC[i][0]) cache[String(dC[i][0])] = { date: dC[i][2] instanceof Date ? _rcIso_(dC[i][2]) : String(dC[i][2] || ""), origine: String(dC[i][3] || ""), drive: String(dC[i][4] || "") };
  var aLire = fiches.filter(function (f) { return f.url && !cache[f.url]; }), dossier = null, lus = 0;
  aLire.forEach(function (f) {
    if (Date.now() > fin) return;
    try {
      dossier = dossier || _rcDossier_();
      var r = _rcLireCv_(f, dossier);
      cache[f.url] = r; lus++;
      wsC.appendRow([f.url, f.nom, r.date, r.origine, r.drive, new Date()]);
    } catch (e) { console.warn("CV " + f.url + " : " + e); }
  });
  var enAttente = fiches.filter(function (f) { return f.url && !cache[f.url]; }).length;

  // Étudiants du CRM
  var wsE = getSheetSafe(ss, "ETUDIANTS"), C = _rcColonnes_(wsE), dE = wsE.getDataRange().getValues();
  var etus = [];
  for (var j = 1; j < dE.length; j++) {
    var row = dE[j];
    if (!row[0] || String(row[0]).toUpperCase() === "ETU000X") continue;
    var toks = _rcNorm_(row[2] + " " + row[3]).split(" ").filter(function (w) { return w.length > 1; });
    etus.push({ ligne: j + 1, id: String(row[0]), nom: (String(row[3] || "") + " " + String(row[2] || "")).trim(), classe: String(row[14] || ""), statut: String(row[13] || ""),
      actif: !_etuMotifRefus_(row), recherche: /recherche/i.test(String(row[13] || "")), toks: toks, ancienEcart: String(row[C.ecart] || "") });
  }
  function trouver(nom) {
    var t = _rcNorm_(nom).split(" ").filter(function (w) { return w.length > 1; });
    if (t.length < 2) return [];
    return etus.filter(function (e) {
      if (!e.toks.length) return false;
      var a = e.toks.every(function (w) { return t.indexOf(w) !== -1; }), b = t.every(function (w) { return e.toks.indexOf(w) !== -1; });
      return a || b;
    });
  }
  var ecarts = [], surSite = {}, maj = {}, nbIndispo = 0;
  fiches.forEach(function (f) {
    var c = trouver(f.nom), actifs = c.filter(function (e) { return e.actif; });
    var e = actifs.length === 1 ? actifs[0] : (c.length === 1 ? c[0] : null);
    if (!e) {
      ecarts.push({ statutCrm: "—", surSite: true, statutSite: f.statutSite, type: c.length > 1 ? "Homonymes" : "Inconnu du CRM", gravite: "À vérifier", nom: f.nom, id: "", detail: c.length > 1 ? "Plusieurs fiches CRM possibles : " + c.map(function (x) { return x.id; }).join(", ") : "Profil présent sur l'espace recruteurs sans fiche correspondante dans le CRM", site: f.diplome + (f.campus ? " · " + f.campus : ""), url: f.url });
      return;
    }
    surSite[e.id] = f;
    var info = cache[f.url] || {};
    // Statut sur le site : « Lien CV » (bleu) = en recherche ; « PLUS DISPONIBLE » (gris) = plus disponible
    var note = "", type = "", indispo = f.plusDispo, dispo = indispo ? "Plus disponible" : (f.url ? "En recherche" : "Sans CV");
    if (indispo) {
      if (e.actif && e.recherche) { type = "À repasser en recherche"; note = "Affiché « PLUS DISPONIBLE » aux recruteurs alors qu'en recherche d'alternance dans le CRM — remettre le « Lien CV » sur le site"; }
      else nbIndispo++;   // profil déjà choisi (ou sorti), laissé en vitrine « PLUS DISPONIBLE » : conforme
    } else if (!e.actif) { type = "Sorti mais disponible"; note = "Étudiant sorti dans le CRM mais présenté en recherche (lien CV) aux recruteurs — les sortis ne doivent pas être disponibles : passer en « PLUS DISPONIBLE » ou retirer du site"; }
    else if (!e.recherche) { type = "À passer PLUS DISPONIBLE"; note = "Statut CRM « " + (e.statut || "vide") + " » mais présenté en recherche (lien CV) aux recruteurs — passer en « PLUS DISPONIBLE » sur le site (ou corriger le statut CRM)"; }
    else if (!f.url) { type = "Sans lien CV"; note = "En recherche dans le CRM, présent sur le site mais sans « Lien CV » — ajouter le CV"; }
    if (note) ecarts.push({ statutCrm: e.statut || (e.actif ? "—" : "Sorti"), surSite: true, statutSite: f.statutSite, type: type, gravite: "Erreur", nom: e.nom, id: e.id, detail: note, site: f.diplome, url: f.url });
    maj[e.ligne] = { lien: f.url, drive: info.drive || "", date: info.date || "", ecart: note, dispo: dispo, cvtheque: indispo ? "Plus disponible" : _rcEtatCvtheque_(info.date) };
  });
  etus.forEach(function (e) {
    if (!e.actif || !e.recherche || surSite[e.id]) return;
    var note = "En recherche d'alternance dans le CRM mais absent de l'espace recruteurs du site";
    ecarts.push({ statutCrm: e.statut, surSite: false, statutSite: "", type: "Absent du site", gravite: "Erreur", nom: e.nom, id: e.id, detail: note, site: "—", url: "" });
    maj[e.ligne] = { lien: "", drive: "", date: "", ecart: note, dispo: "", cvtheque: "Non" };
  });
  // Les étudiants qui ne sont plus concernés : on efface l'ancien écart
  etus.forEach(function (e) { if (!maj[e.ligne] && e.ancienEcart) maj[e.ligne] = { lien: null, drive: null, date: null, dispo: null, ecart: "" }; });

  // Écriture dans ETUDIANTS + note au carnet de route à l'apparition d'un nouvel écart
  var wsN = getSheetSafe(ss, "Carnet_route"), maintenant = new Date(), notes = 0, nouvellesNotes = [];
  var nL = dE.length - 1, cols = {};
  Object.keys(C).forEach(function (k) { cols[k] = dE.slice(1).map(function (r) { return [r[C[k]] === undefined ? "" : r[C[k]]]; }); });   // écriture groupée par colonne
  Object.keys(maj).forEach(function (l) {
    var m = maj[l], e = etus.filter(function (x) { return x.ligne === Number(l); })[0], k = Number(l) - 2;
    if (m.lien !== null) { cols.lien[k][0] = m.lien; cols.drive[k][0] = m.drive; cols.date[k][0] = m.date; cols.dispo[k][0] = m.dispo; }
    cols.verifie[k][0] = maintenant;
    cols.ecart[k][0] = m.ecart;
    if (m.ecart && m.ecart !== e.ancienEcart) {
      nouvellesNotes.push([e.id, "Audit espace recruteurs", maintenant, "⚠️ Espace recruteurs hecg.fr : " + m.ecart + ".", false]);
      notes++;
    }
    try { _invalidateStudentDetailsCache_(e.id); } catch (eI) {}
  });
  if (nL > 0) Object.keys(C).forEach(function (k) { wsE.getRange(2, C[k] + 1, nL, 1).setValues(cols[k]); });
  // Colonne Y « CVthèque » (ex-CV Tech) : Oui si CV à jour (< 6 mois) sur le site, « À actualiser » sinon, Non si absent
  var colY = dE.slice(1).map(function (r) { return [r[24] === undefined ? "" : r[24]]; }), cvOk = 0;
  Object.keys(maj).forEach(function (l) { var v = maj[l].cvtheque; if (v) { colY[Number(l) - 2][0] = v; if (v === "Oui") cvOk++; } });
  if (nL > 0) wsE.getRange(2, 25, nL, 1).setValues(colY);
  if (nouvellesNotes.length) wsN.getRange(wsN.getLastRow() + 1, 1, nouvellesNotes.length, 5).setValues(nouvellesNotes);

  // Rapport
  var R = { date: Utilities.formatDate(maintenant, "Europe/Paris", "dd/MM/yyyy HH:mm"), auteur: auteur, profilsSite: fiches.length,
    enRecherche: etus.filter(function (e) { return e.actif && e.recherche; }).length, cvthequeOk: cvOk, presents: Object.keys(surSite).length, indisponibles: nbIndispo, cvLus: lus, cvEnAttente: enAttente, notes: notes, ecarts: ecarts };
  var wsR = ss.getSheetByName("AUDIT_RECRUTEURS") || ss.insertSheet("AUDIT_RECRUTEURS");
  wsR.clear();
  var rows = [["Audit espace recruteurs ↔ CRM — " + R.date + " — " + auteur, "", "", "", "", ""],
    [R.profilsSite + " profils sur le site · " + R.enRecherche + " étudiants « En recherche » dans le CRM · " + R.presents + " rapprochés (dont " + nbIndispo + " affichés « Indisponible » — déjà choisis) · " + ecarts.length + " écart(s) · " + R.cvEnAttente + " CV en attente de lecture", "", "", "", "", ""],
    ["", "", "", "", "", ""], ["Gravité", "Type", "Étudiant", "ID CRM", "Détail", "CV sur le site"]];
  ecarts.forEach(function (x) { rows.push([x.gravite, x.type, x.nom, x.id, x.detail, x.url]); });
  wsR.getRange(1, 1, rows.length, 6).setValues(rows);
  wsR.getRange(4, 1, 1, 6).setFontWeight("bold").setBackground("#1A4E8A").setFontColor("#fff");
  wsR.setFrozenRows(4);
  var meta = JSON.parse(JSON.stringify(R)); meta.ecarts = ecarts.length;
  PropertiesService.getScriptProperties().setProperty("RC_DERNIER", JSON.stringify(meta));
  if (enAttente) PropertiesService.getScriptProperties().setProperty("RC_SUITE", "1"); else PropertiesService.getScriptProperties().deleteProperty("RC_SUITE");
  R.ecartsDetail = ecarts;
  logAction(auteur, "", "Action", "Étudiants — audit espace recruteurs", "AUDIT_RECRUTEURS",
    R.profilsSite + " profils sur le site (" + nbIndispo + " « PLUS DISPONIBLE » conformes), " + R.enRecherche + " en recherche dans le CRM, " + R.presents + " rapprochés (" + nbIndispo + " indisponibles), " + ecarts.length + " écart(s), " + notes + " note(s) ajoutée(s) aux dossiers, " + lus + " CV lus" + (enAttente ? " (" + enAttente + " en attente)" : ""));
  return R;
}

// ── Actions de l'application ─────────────────────────────────────────────────
function handleGetAuditRecruteurs(p, ss) {
  if (!_rcPeut_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  var props = PropertiesService.getScriptProperties(), meta = props.getProperty("RC_DERNIER"), ws = ss.getSheetByName("AUDIT_RECRUTEURS"), ecarts = [];
  if (ws && ws.getLastRow() > 4) ws.getRange(5, 1, ws.getLastRow() - 4, 6).getValues().forEach(function (r) { ecarts.push({ gravite: r[0], type: r[1], nom: r[2], id: r[3], detail: r[4], url: r[5] }); });
  return createJsonResponse({ success: true, mdpConfigure: !!props.getProperty("RECRUTEURS_MDP"), dernier: meta ? JSON.parse(meta) : null, ecarts: ecarts });
}
function handleLancerAuditRecruteurs(p, ss) {
  if (!_rcPeut_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return createJsonResponse({ success: false, message: "Un autre traitement est en cours, réessayez dans une minute." });
  try {
    var R = _rcAudit_(ss, p.nomAuteur || p.emailAuteur || "Manuel", 200000);
    if (p.envoyerMail === true || p.envoyerMail === "true") _rcMailEcarts_(R, R.ecartsDetail);
    delete R.ecartsDetail;
    return createJsonResponse({ success: true, rapport: R });
  } catch (e) {
    var m = String(e.message || e);
    if (m === "MDP_MANQUANT") return createJsonResponse({ success: false, mdp: true, message: "Le mot de passe de l'espace recruteurs n'est pas encore enregistré." });
    if (m === "MDP_INCORRECT") return createJsonResponse({ success: false, mdp: true, message: "Mot de passe de l'espace recruteurs refusé par le site." });
    return createJsonResponse({ success: false, message: m });
  } finally { lock.releaseLock(); }
}
function handleSaveMdpRecruteurs(p, ss) {
  if (String(p.vieScRole || "") !== "Super-admin" && String(p.emailAuteur || "").toLowerCase() !== "m.mokhtari@hecg.fr") return createJsonResponse({ success: false, message: "Réservé au Super-admin." });
  var mdp = String(p.mdp || "").trim();
  if (!mdp || mdp.length > 100) return createJsonResponse({ success: false, message: "Mot de passe invalide." });
  PropertiesService.getScriptProperties().setProperty("RECRUTEURS_MDP", mdp);
  logAction(p.nomAuteur || p.emailAuteur || "", p.vieScRole || "", "Modification", "Étudiants — audit espace recruteurs", "Paramètre", "Mot de passe de l'espace recruteurs enregistré (non journalisé)");
  return createJsonResponse({ success: true });
}
// Passage automatique (appelé par auditAdressesTick) : chaque nuit à 6 h, puis suite des CV à lire
function _rcAuto_() {
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty("RECRUTEURS_MDP")) return;
  var h = Number(Utilities.formatDate(new Date(), "Europe/Paris", "H")), dernier = Number(props.getProperty("RC_AUTO_TS") || 0);
  var suite = props.getProperty("RC_SUITE"), nuit = h === 6 && Date.now() - dernier > 20 * 3600 * 1000;
  if (!suite && !nuit) return;
  props.setProperty("RC_AUTO_TS", String(Date.now()));
  try {
    var R = _rcAudit_(SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ"), "Audit automatique (nuit)", 150000);
    // Rapport des écarts par mail tous les 2 jours (une fois tous les CV lus, ou au plus tard au passage de nuit suivant)
    var dernierMail = Number(props.getProperty("RC_MAIL_TS") || 0);
    if (Date.now() - dernierMail > 44 * 3600 * 1000 && (!props.getProperty("RC_SUITE") || nuit)) {
      _rcMailEcarts_(R, R.ecartsDetail);
      props.setProperty("RC_MAIL_TS", String(Date.now()));
    }
  }
  catch (e) { console.error("audit recruteurs : " + e); }
}

// ── Mail des écarts, tous les 2 jours (demande direction 2026-11-25) ─────────
var RC_MAIL_DEST = ["m.mokhtari@hecg.fr", "a.mhoudini@hecg.fr"];
var RC_APP = "https://mokhtarimehdipro-bit.github.io/CRMHECG/";
function _rcMailEcarts_(R, ecarts) {
  var esc = function (v) { return String(v == null ? "" : v).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); };
  var th = 'style="background:#1A4E8A;color:#fff;padding:7px 8px;font-size:12px;text-align:left"', td = 'style="padding:6px 8px;border-bottom:1px solid #e5e7eb;font-size:12px;vertical-align:top"';
  var lignes = ecarts.map(function (x) {
    var coul = x.gravite === "Erreur" ? "#b91c1c" : "#b45309";
    return "<tr><td " + td + "><b>" + esc(x.nom) + "</b>" + (x.id ? '<br><span style="color:#94a3b8">' + esc(x.id) + "</span>" : "") + "</td>" +
      "<td " + td + ">" + esc(x.statutCrm || "—") + "</td>" +
      "<td " + td + ">" + (x.surSite ? "Oui" : "<b>Non</b>") + "</td>" +
      "<td " + td + ">" + (x.statutSite === "PLUS DISPONIBLE" ? '<span style="background:#6b7280;color:#fff;border-radius:4px;padding:1px 6px;font-weight:bold">PLUS DISPONIBLE</span>' : x.statutSite ? '<span style="background:#2563eb;color:#fff;border-radius:4px;padding:1px 6px;font-weight:bold">' + esc(x.statutSite) + "</span>" : "—") + "</td>" +
      '<td ' + td + '><span style="color:' + coul + ';font-weight:bold">' + esc(x.type) + "</span><br>" + esc(x.detail) + "</td>" +
      "<td " + td + ">" + (x.id ? '<a href="' + RC_APP + "dossier.html?id=" + encodeURIComponent(x.id) + '">Espace CRM</a>' : "—") + "</td>" +
      "<td " + td + '><a href="' + esc(x.url || RC_PAGE) + '">' + (x.url ? "CV sur l'espace recruteurs" : "Espace recruteurs") + "</a></td></tr>";
  }).join("");
  var html = '<div style="font-family:Arial,sans-serif;max-width:1000px">' +
    '<h2 style="color:#1A4E8A;margin:0 0 6px">Audit espace recruteurs ↔ CRM — ' + esc(R.date) + "</h2>" +
    '<p style="font-size:13px;color:#374151;margin:0 0 12px">' + R.profilsSite + " profils sur l'espace recruteurs · " + R.enRecherche + " étudiants « En recherche » dans le CRM · " +
    R.presents + " rapprochés (dont " + R.indisponibles + " « PLUS DISPONIBLE ») · <b>" + ecarts.length + " écart(s)</b></p>" +
    (ecarts.length ? '<table style="border-collapse:collapse;width:100%"><tr><th ' + th + ">Étudiant</th><th " + th + ">Statut CRM</th><th " + th + ">Présent CVthèque</th><th " + th + ">Statut CVthèque</th><th " + th + ">Écart</th><th " + th + ">CRM</th><th " + th + ">Espace recruteurs</th></tr>" + lignes + "</table>"
      : '<p style="font-size:14px;color:#047857;font-weight:bold">✓ Aucun écart : la CVthèque est conforme au CRM.</p>') +
    '<p style="font-size:11px;color:#94a3b8;margin-top:14px">Règles : « Lien CV » (fond bleu) = en recherche ; « PLUS DISPONIBLE » (fond gris) = plus disponible. Un étudiant sorti ne doit jamais être présenté en recherche. ' +
    "Détail complet : CRM › Vie scolaire › Espace recruteurs. Envoi automatique tous les 2 jours.</p></div>";
  MailApp.sendEmail({ to: RC_MAIL_DEST.join(","), subject: "🔎 Audit espace recruteurs — " + ecarts.length + " écart(s) — " + R.date.slice(0, 10), htmlBody: html, name: "CRM HECG" });
  logAction("Système", "", "Envoi mail", "Étudiants — audit espace recruteurs", "AUDIT_RECRUTEURS", "Rapport des écarts (" + ecarts.length + ") envoyé à " + RC_MAIL_DEST.join(", "));
}
