// =============================================================================
//  RELANCE QUOTIDIENNE DES ÉTUDIANTS EN RECHERCHE (demande direction 2026-11-09)
//  Chaque jour ouvré (lundi-vendredi, 9 h), un mail est envoyé aux étudiants INSCRITS ou PRÉINSCRITS « En recherche »
//  qui n'ont ajouté AUCUNE candidature dans leur suivi (feuille TSUIVI) depuis plus de 24 h,
//  pour leur rappeler de tenir l'application à jour. polepedagogique@hecg.fr est en copie de
//  toutes les relances étudiants (celle-ci et les relances hebdomadaires de code.gs).
//  Réglages (onglet Vie scolaire › Espace recruteurs) : activation, aperçu, envoi immédiat.
// =============================================================================
var POLE_PEDA_EMAIL = "polepedagogique@hecg.fr";
var RELANCE_CAND_HEURE = 9;

function _relNorm_(v) { return String(v || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toUpperCase().replace(/[^A-Z]/g, ""); }

function _relCandCibles_(ss) {
  var dE = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues();
  var dT = getSheetSafe(ss, "TSUIVI").getDataRange().getValues(), derniere = {}, total = {};
  for (var i = 1; i < dT.length; i++) {
    var id = String(dT[i][0] || "").trim().toUpperCase();
    if (!id) continue;
    var d = dT[i][12] instanceof Date ? dT[i][12] : (dT[i][13] instanceof Date ? dT[i][13] : null);
    total[id] = (total[id] || 0) + 1;
    if (d && (!derniere[id] || d > derniere[id])) derniere[id] = d;
  }
  var seuil = Date.now() - 24 * 3600 * 1000, out = [];
  for (var j = 1; j < dE.length; j++) {
    var r = dE[j], id2 = String(r[0] || "").trim();
    if (!id2 || id2.toUpperCase() === "ETU000X") continue;
    if (!/recherche/i.test(String(r[13] || ""))) continue;
    var typo = _relNorm_(r[20]) || _relNorm_(calculerTypoEtu(r[17], r[18], r[14]));   // colonne U « Typo » (sinon recalculée)
    if (typo !== "INSCRIT" && typo !== "PREINSCRIT") continue;            // uniquement inscrits et préinscrits
    if (_etuMotifRefus_(r)) continue;                                   // sortis et prospects exclus
    var mail = String(r[4] || "").trim();
    if (mail.indexOf("@") < 1) continue;
    var der = derniere[id2.toUpperCase()] || null;
    if (der && der.getTime() > seuil) continue;                          // candidature ajoutée depuis moins de 24 h
    out.push({ id: id2, nom: String(r[2] || "").trim(), prenom: String(r[3] || "").trim(), mail: mail, classe: String(r[14] || ""),
      nb: total[id2.toUpperCase()] || 0, derniere: der ? Utilities.formatDate(der, "Europe/Paris", "dd/MM/yyyy") : "" });
  }
  return out;
}

function _relCandHtml_(c) {
  var jours = "";
  if (c.derniere) {
    var m = c.derniere.split("/"), n = Math.floor((Date.now() - new Date(+m[2], +m[1] - 1, +m[0]).getTime()) / 86400000);
    jours = n >= 1 ? n + " jour" + (n > 1 ? "s" : "") : "moins d'un jour";
  }
  var app = "https://mokhtarimehdipro-bit.github.io/CRMHECG/";
  return '<div style="font-family:Arial,Helvetica,sans-serif;max-width:580px;margin:auto;border:1px solid #eef2f7;border-radius:16px;overflow:hidden">'
    + '<div style="background:linear-gradient(120deg,#1A4E8A,#F26522);padding:22px 26px;color:#fff"><p style="margin:0;font-size:12px;opacity:.85;letter-spacing:.08em;text-transform:uppercase">HECG · Suivi de votre recherche d\'alternance</p>'
    + '<h2 style="margin:6px 0 0;font-size:21px">Bonjour ' + c.prenom + ', où en est votre recherche ?</h2></div>'
    + '<div style="padding:22px 26px;color:#1f2937;font-size:14.5px;line-height:1.6">'
    + (c.nb ? '<p>Vous avez enregistré <b>' + c.nb + ' candidature' + (c.nb > 1 ? 's' : '') + '</b> dans votre espace ; la dernière date du <b>' + c.derniere + '</b> (' + jours + ').</p>'
            : '<p>Vous n\'avez encore enregistré <b>aucune candidature</b> dans votre espace étudiant.</p>')
    + '<p>Pour que l\'équipe pédagogique puisse vous accompagner efficacement, pensez à mettre à jour l\'application :</p>'
    + '<ul style="padding-left:18px;margin:8px 0 14px"><li>ajoutez <b>chaque nouvelle candidature</b> (entreprise, date, contact) ;</li>'
    + '<li>notez les <b>relances, entretiens et réponses</b> reçues ;</li>'
    + '<li>consultez les <b>offres partagées</b> par l\'école ;</li>'
    + '<li>vous avez trouvé votre alternance ? <b>Prévenez-nous</b> pour mettre à jour votre statut.</li></ul>'
    + '<p style="text-align:center;margin:22px 0"><a href="' + app + '" style="background:#F26522;color:#fff;text-decoration:none;font-weight:bold;padding:13px 28px;border-radius:12px;display:inline-block">Mettre à jour mon suivi</a></p>'
    + '<p style="font-size:12.5px;color:#6b7280">Identifiant de connexion : <b>' + c.id + '</b>. En cas de difficulté, répondez à ce message : le pôle pédagogique est en copie.</p></div>'
    + '<div style="background:#f8fafc;padding:12px 26px;font-size:11px;color:#9ca3af">Rappel automatique envoyé les jours ouvrés aux étudiants en recherche sans nouvelle candidature depuis 24 h.</div></div>';
}

function _relCandEnvoyer_(ss, auteur) {
  var cibles = _relCandCibles_(ss), envoyes = 0, erreurs = [];
  cibles.forEach(function (c) {
    var sujet = "HECG — Pensez à mettre à jour votre suivi de candidatures";
    try {
      GmailApp.sendEmail(c.mail, sujet, "", { htmlBody: _relCandHtml_(c), from: SENDER_EMAIL, cc: POLE_PEDA_EMAIL, name: "HECG — Pôle pédagogique" });
      envoyes++;
      try { logMailDetail(ss, { auteur: auteur, categorie: "Relance auto", sousCat: "Relance quotidienne candidatures", objet: sujet, destinataires: [c.mail, POLE_PEDA_EMAIL],
        corps: c.prenom + " " + c.nom + " — " + c.nb + " candidature(s), dernière : " + (c.derniere || "aucune"), sourceId: "RELCAND-" + c.id + "-" + Utilities.formatDate(new Date(), "Europe/Paris", "yyyy-MM-dd") }); } catch (eL) {}
    } catch (e) { erreurs.push(c.mail + " : " + e); }
  });
  logAction(auteur, "", "Action", "Étudiants — relance candidatures", "TSUIVI",
    envoyes + " relance(s) envoyée(s) aux étudiants en recherche sans candidature depuis 24 h (copie " + POLE_PEDA_EMAIL + ")" + (erreurs.length ? " — " + erreurs.length + " erreur(s)" : ""));
  PropertiesService.getScriptProperties().setProperty("RELANCE_CAND_DERNIER", JSON.stringify({ date: Utilities.formatDate(new Date(), "Europe/Paris", "dd/MM/yyyy HH:mm"), envoyes: envoyes, erreurs: erreurs.length, auteur: auteur }));
  return { envoyes: envoyes, erreurs: erreurs };
}

// Déclencheur : appelé à chaque passage de auditAdressesTick (toutes les 5 min)
function _relCandAuto_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty("RELANCE_CAND_ACTIVE") === "0") return;
  var now = new Date(), jour = Number(Utilities.formatDate(now, "Europe/Paris", "u")), h = Number(Utilities.formatDate(now, "Europe/Paris", "H"));
  if (jour > 5 || h < RELANCE_CAND_HEURE || h >= 12) return;              // jours ouvrés, envoi le matin (9 h-12 h) uniquement
  var auj = Utilities.formatDate(now, "Europe/Paris", "yyyy-MM-dd");
  if (props.getProperty("RELANCE_CAND_JOUR") === auj) return;               // une fois par jour
  props.setProperty("RELANCE_CAND_JOUR", auj);
  _relCandEnvoyer_(SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ"), "Relance automatique (9 h)");
}

// ── Actions de l'application ──
function handleGetRelanceCandidatures(p, ss) {
  if (!_rcPeut_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  var props = PropertiesService.getScriptProperties(), d = props.getProperty("RELANCE_CAND_DERNIER");
  return createJsonResponse({ success: true, active: props.getProperty("RELANCE_CAND_ACTIVE") !== "0", heure: RELANCE_CAND_HEURE, copie: POLE_PEDA_EMAIL,
    dernier: d ? JSON.parse(d) : null, envoyeAujourdhui: props.getProperty("RELANCE_CAND_JOUR") === Utilities.formatDate(new Date(), "Europe/Paris", "yyyy-MM-dd"),
    cibles: _relCandCibles_(ss) });
}
function handleSaveRelanceCandidatures(p, ss) {
  if (!_rcPeut_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
  if (p.active !== undefined) {
    PropertiesService.getScriptProperties().setProperty("RELANCE_CAND_ACTIVE", p.active ? "1" : "0");
    logAction(p.nomAuteur || p.emailAuteur || "", p.vieScRole || "", "Modification", "Étudiants — relance candidatures", "Paramètre", "Relance quotidienne " + (p.active ? "activée" : "désactivée"));
  }
  if (p.envoyerMaintenant) {
    var r = _relCandEnvoyer_(ss, p.nomAuteur || p.emailAuteur || "Envoi manuel");
    PropertiesService.getScriptProperties().setProperty("RELANCE_CAND_JOUR", Utilities.formatDate(new Date(), "Europe/Paris", "yyyy-MM-dd"));
    return createJsonResponse({ success: true, envoyes: r.envoyes, erreurs: r.erreurs });
  }
  return createJsonResponse({ success: true });
}
