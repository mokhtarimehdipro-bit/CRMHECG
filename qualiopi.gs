// =============================================================================
//  QUALIOPI — preuves datées, obsolescence, plan d'actions (Vie scolaire › Qualiopi)
//  Référentiel : décret n° 2026-728 du 1er août 2026 (33 indicateurs, en vigueur au 1/11/2026).
//  Feuille QUALIOPI_PREUVES (une ligne par preuve) :
//   A Indicateur | B Libellé | C URL (fichier Drive ou lien web) | D Date de dépôt | E Type (document|lien|capture)
//   F Commentaire | G Date du document | H Validité (mois, 0 = permanente) | I Capture (URL) | J Source | K ID
//  Feuille QUALIOPI_ACTIONS : points à traiter avant l'audit (cochables).
// =============================================================================

var Q_ENTETES = ["Indicateur_Numero", "Nom_Preuve", "URL", "Date_Depot", "Type", "Commentaire", "Date_Document", "Validite_Mois", "Capture_URL", "Source", "ID"];
var Q_ACTIONS_ENTETES = ["ID", "Indicateur", "Priorité", "Action", "Statut", "Mis à jour le", "Par"];

function _qEstSuper_(p) { return String(p.vieScRole || "") === "Super-admin" || String(p.emailAuteur || "").toLowerCase() === "m.mokhtari@hecg.fr"; }
function _qPeutLire_(p) { return _qEstSuper_(p) || String(p.vieScRole || "") === "Pédagogie"; }

function _qFeuille_(ss) {
  var ws = getSheetSafe(ss, "QUALIOPI_PREUVES");
  var h = ws.getRange(1, 1, 1, Math.max(ws.getLastColumn(), Q_ENTETES.length)).getValues()[0];
  if (String(h[4] || "") !== "Type" || String(h[10] || "") !== "ID") ws.getRange(1, 1, 1, Q_ENTETES.length).setValues([Q_ENTETES]).setFontWeight("bold");
  return ws;
}
function _qDateIso_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return Utilities.formatDate(v, "Europe/Paris", "yyyy-MM-dd");
  var t = String(v || "").trim(), m;
  if ((m = t.match(/^(\d{4})-(\d{2})-(\d{2})/))) return m[0];
  if ((m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/))) return m[3] + "-" + ("0" + m[2]).slice(-2) + "-" + ("0" + m[1]).slice(-2);
  return "";
}
// Validité par défaut (mois) selon le type et le libellé
function _qValiditeDefaut_(type, nom) {
  if (type === "lien" || type === "capture") return 12;
  if (/article l|code du travail|bo programme|r[ée]f[ée]rentiel|d[ée]cret|guide de lecture/i.test(nom || "")) return 36;
  return 24;
}
function _qId_() { return "QP-" + Date.now().toString(36) + "-" + Math.floor(Math.random() * 1e4).toString(36); }

function handleGetQualiopiPreuves(p, ss) {
  try {
    var ws = _qFeuille_(ss), data = ws.getDataRange().getValues(), rows = [], aCompleter = [];
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      if (!r[0]) continue;
      var id = String(r[10] || "");
      if (!id) { id = _qId_(); aCompleter.push([i + 1, id]); }
      rows.push({ indicateur: String(r[0]), nomPreuve: String(r[1] || ""), urlFichier: String(r[2] || ""), dateDepot: _qDateIso_(r[3]) || String(r[3] || ""),
        type: String(r[4] || "") || (/drive\.google|docs\.google/.test(String(r[2])) ? "document" : "lien"), commentaire: String(r[5] || ""),
        dateDocument: _qDateIso_(r[6]) || _qDateIso_(r[3]), validite: r[7] === "" || r[7] === null ? "" : Number(r[7]),
        captureUrl: String(r[8] || ""), source: String(r[9] || ""), id: id });
    }
    aCompleter.forEach(function (x) { ws.getRange(x[0], 11).setValue(x[1]); });   // identifiants des anciennes lignes
    return createJsonResponse({ success: true, data: rows });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function _qLigne_(p, url, type) {
  var date = Utilities.formatDate(new Date(), "Europe/Paris", "yyyy-MM-dd");
  var val = p.validite === "" || p.validite === undefined || p.validite === null ? _qValiditeDefaut_(type, p.nomPreuve) : Number(p.validite) || 0;
  return [String(p.indicateur || ""), String(p.nomPreuve || "").slice(0, 300), url, date, type, String(p.commentaire || "").slice(0, 1000),
    _qDateIso_(p.dateDocument) || date, val, String(p.captureUrl || ""), String(p.source || "Ajout manuel"), _qId_()];
}

// Lien web (ou fichier déjà présent dans Drive)
function handleSaveQualiopiPreuve(p, ss) {
  try {
    if (!_qEstSuper_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
    var url = String(p.urlFichier || "").trim();
    if (!/^https?:\/\//i.test(url)) return createJsonResponse({ success: false, message: "Lien invalide (il doit commencer par https://)." });
    var type = p.type === "capture" ? "capture" : (/drive\.google|docs\.google/.test(url) ? "document" : "lien");
    var ligne = _qLigne_(p, url, type);
    _qFeuille_(ss).appendRow(ligne);
    logAction(p.nomAuteur || p.emailAuteur || "?", p.vieScRole || "", "Création", "Vie scolaire — Qualiopi", "Indicateur " + p.indicateur,
      "Preuve (" + type + ") : " + p.nomPreuve + " — " + url);
    return createJsonResponse({ success: true, id: ligne[10] });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// Dépôt d'un fichier (sous-dossier CRITERE_INDICATEUR du dossier Qualiopi)
function handleUploadQualiopiFile(p, ss) {
  try {
    if (!_qEstSuper_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
    if (!p.base64Data || !p.fileName || !p.mimeType) return createJsonResponse({ success: false, message: "Paramètres manquants." });
    var folder = _qDossier_(p.critere, p.indicateur);
    var file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(p.base64Data), p.mimeType, p.fileName));
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var ligne = _qLigne_({ indicateur: p.indicateur, nomPreuve: p.nomPreuve || p.fileName, commentaire: p.commentaire, dateDocument: p.dateDocument || p.dateDepot,
      validite: p.validite, source: "Dépôt dans l'application" }, file.getUrl(), "document");
    _qFeuille_(ss).appendRow(ligne);
    logAction(p.nomAuteur || p.emailAuteur || "?", p.vieScRole || "", "Création", "Vie scolaire — Qualiopi", "Indicateur " + p.indicateur, "Preuve déposée : " + p.fileName);
    return createJsonResponse({ success: true, id: ligne[10], url: file.getUrl(), folderUrl: folder.getUrl() });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}
function _qDossier_(critere, indicateur) {
  var nom = String(critere || "0") + "_" + String(indicateur || "0");
  var root = DriveApp.getFolderById(QUALIOPI_ROOT_ID), it = root.getFoldersByName(nom);
  return it.hasNext() ? it.next() : root.createFolder(nom);
}
function handleGetQualiopiDossierUrl(p, ss) {
  try {
    if (!_qEstSuper_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
    var root = DriveApp.getFolderById(QUALIOPI_ROOT_ID), it = root.getFoldersByName(String(p.critere || "0") + "_" + String(p.indicateur || "0"));
    return createJsonResponse({ success: true, url: it.hasNext() ? it.next().getUrl() : root.getUrl() });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// Modification : date du document, validité, commentaire, libellé, indicateur
function handleUpdateQualiopiPreuve(p, ss) {
  try {
    if (!_qEstSuper_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
    var ws = _qFeuille_(ss), data = ws.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][10]) !== String(p.id)) continue;
      var r = data[i].slice(0, 11);
      while (r.length < 11) r.push("");
      if (p.indicateur) r[0] = String(p.indicateur);
      if (p.nomPreuve) r[1] = String(p.nomPreuve).slice(0, 300);
      if (p.commentaire !== undefined) r[5] = String(p.commentaire).slice(0, 1000);
      if (p.dateDocument) r[6] = _qDateIso_(p.dateDocument);
      if (p.validite !== undefined && p.validite !== "") r[7] = Number(p.validite) || 0;
      ws.getRange(i + 1, 1, 1, 11).setValues([r]);
      logAction(p.nomAuteur || p.emailAuteur || "?", p.vieScRole || "", "Modification", "Vie scolaire — Qualiopi", "Indicateur " + r[0],
        "Preuve modifiée : " + r[1] + " (document du " + r[6] + ", validité " + r[7] + " mois)");
      return createJsonResponse({ success: true });
    }
    return createJsonResponse({ success: false, message: "Preuve introuvable." });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleDeleteQualiopiPreuve(p, ss) {
  try {
    if (!_qEstSuper_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
    var ws = _qFeuille_(ss), data = ws.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var ok = p.id ? String(data[i][10]) === String(p.id)
        : (String(data[i][0]) === String(p.indicateur) && String(data[i][1]) === String(p.nomPreuve) && String(data[i][2]) === String(p.urlFichier));
      if (!ok) continue;
      ws.deleteRow(i + 1);
      logAction(p.nomAuteur || p.emailAuteur || "?", p.vieScRole || "", "Suppression", "Vie scolaire — Qualiopi", "Indicateur " + data[i][0], "Preuve retirée : " + data[i][1]);
      return createJsonResponse({ success: true });
    }
    return createJsonResponse({ success: false, message: "Preuve introuvable." });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// ── Plan d'actions avant audit ───────────────────────────────────────────────
function _qActionsFeuille_(ss) {
  var ws = ss.getSheetByName("QUALIOPI_ACTIONS");
  if (!ws) {
    ws = ss.insertSheet("QUALIOPI_ACTIONS");
    ws.getRange(1, 1, 1, Q_ACTIONS_ENTETES.length).setValues([Q_ACTIONS_ENTETES]).setFontWeight("bold");
    ws.setFrozenRows(1);
    var d = "2026-11-09", lignes = Q_ACTIONS_INITIALES.map(function (a, i) { return ["QA-" + (i + 1), a[0], a[1], a[2], "À faire", d, "Analyse préparatoire (site, Drive, audits 2024-2025, décret 2026-728)"]; });
    ws.getRange(2, 1, lignes.length, 7).setValues(lignes);
  }
  return ws;
}
// Écarts relevés le 09/11/2026 (site hecg.fr, Drive « Dossier pour Audit », rapports d'audit 2024 et 2025, décret 2026-728)
var Q_ACTIONS_INITIALES = [
  ["1", "Haute", "Site : supprimer ou justifier les affirmations « 100 % des étudiants sont placés par l'école » / « placer l'intégralité de ses étudiants » (nouvel ind. 1 : aucune mention de nature à induire en erreur)."],
  ["1", "Haute", "Site, DCG/DSCG : « aucun reste à charge » est inexact pour les contrats d'apprentissage niveau 6-7 signés depuis le 1/7/2025 (participation employeur de 750 €) — corriger et ajouter les montants de prise en charge OPCO (remarque audit 2025)."],
  ["1", "Haute", "Bachelor CDSP : publier tarifs, n° RNCP, certificateur et convention d'habilitation ; retirer « diplôme d'État » (titre RNCP) ; harmoniser CDSP / RPSP."],
  ["1", "Haute", "Liens « Mentions légales », « Politique de confidentialité », « Mission d'un CFA » et contact pointent vers l'ancien domaine brown-squid-812410.hostingersite.com (morts) — republier sur hecg.fr."],
  ["1", "Moyenne", "Ajouter la mention d'accessibilité handicap + contact référente sur CHAQUE fiche formation ; mettre une date de mise à jour sur les fiches."],
  ["1", "Moyenne", "Dossier de candidature et Instagram mentionnent le DEC et une « DSCG année 3 » non proposés sur le site ; formulaire d'inscription : ajouter BTS GPME et CDSP, remplacer les séries de bac L/ES/S."],
  ["1", "Basse", "BTS CG : volume horaire incohérent (1362 h / « TOTAL 1350 H ») ; page contrat d'apprentissage : SMIC au 01/01/2025 à actualiser."],
  ["2", "Haute", "Actualiser les taux (page « Nos diplômes » datée du 30/06/2025, « DCG disponible en juillet 2025 ») avec les résultats 2026 et publier la MÉTHODE DE CALCUL de chaque taux (nouvel ind. 2) ; distinguer apprentis / formation continue."],
  ["3", "Haute", "Mettre le lien InserJeunes (taux d'obtention, poursuite d'études, insertion, rupture) sur le site ; ajouter « chaque UE peut être validée indépendamment » et les passerelles / poursuites d'études par diplôme (remarques audit 2025, ind. 3 modifié)."],
  ["4", "Moyenne", "Formaliser une trame d'entretien amont « analyse du besoin entreprise » datée 2026 ; dater les CV ; conserver lettres de motivation école + entreprise (remarque audit 2025)."],
  ["6", "Moyenne", "Progressions pédagogiques 2026-2027 (les pièces actuelles datent de 2024-2025)."],
  ["7", "Haute", "Mettre à jour les tableaux croisés DCG/DSCG (2024) ; prouver la capacité à assurer la certification (inscriptions SIEC / convention d'habilitation CDSP) — ind. 7 modifié."],
  ["8", "Moyenne", "Actualiser la fiche de positionnement (2022) et joindre des positionnements 2026 renseignés."],
  ["9", "Moyenne", "Livret d'accueil et règlement intérieur 2026-2027 datés, avec preuve de remise (émargement / mail) ; emplois du temps 2026-2027."],
  ["12", "Haute", "NOUVEAU : procédure écrite de prévention ET de traitement des violences (VSS), du harcèlement et des discriminations, référent identifié, circuit de signalement ; finaliser l'affiche harcèlement ; registre des situations traitées."],
  ["13", "Haute", "NC mineure 2024 : livret de suivi en entreprise RENSEIGNÉ et signé + CR de visites, preuve de l'articulation et de la progressivité centre / entreprise pour l'échantillon."],
  ["14", "Haute", "NC majeure 2024 : actions d'information sur la mixité des métiers et l'égalité femmes-hommes (affiches, posts, interventions) + procédure de traitement sans délai des ruptures liées à des difficultés, violences ou discriminations (ind. 14 modifié) ; actualiser aides permis / 1er équipement (2022)."],
  ["15", "Haute", "Ind. 15 modifié : information des apprentis (y compris mineurs) sur VSS/harcèlement/discriminations, médiateur de l'apprentissage et signalement à l'inspection du travail — absent du site ; mettre à jour le document SST (2022)."],
  ["16", "Moyenne", "Joindre les preuves d'inscription aux examens et convocations 2026 (DCG/DSCG/BTS)."],
  ["17", "Moyenne", "DUERP 2026 ; bail et plan du campus d'Aix (extension 290 m²) ; registre public d'accessibilité ; liste des équipements."],
  ["18", "Moyenne", "Liste des formateurs 2026-2027 et comptes rendus de réunions pédagogiques."],
  ["19", "Moyenne", "Ind. 19 modifié : désigner un référent pédagogique par formation ; si modules à distance, tracer l'effectivité du suivi (connexions)."],
  ["20", "Haute", "Nommer et publier le référent mobilité nationale / internationale ; CR du conseil de perfectionnement 2025-2026 ; gouvernance participative (apprentis, formateurs, entreprises) — ind. 20 modifié."],
  ["21", "Moyenne", "Matrice des compétences formateurs 2026 et CV datés ; évaluations annuelles des formateurs."],
  ["22", "Basse", "Plan de développement des compétences des salariés 2026 + attestations."],
  ["23", "Moyenne", "Veille légale récente (les pièces datent de 2022) : décret 2026-728, réforme du financement de l'apprentissage — et trace de son exploitation (CR de direction)."],
  ["25", "Moyenne", "Veille sur les innovations pédagogiques : une seule pièce ; documenter les outils testés (IA générative, plateformes, classes inversées)."],
  ["26", "Basse", "Actualiser les ressources handicap (AGEFIPH, Handicontacts, livret d'accueil handicap 2022)."],
  ["27", "Basse", "Sous-traitance : déclarer « non applicable » si aucune ; sinon contrats avec clause de conformité Qualiopi tracée (ind. 27 modifié)."],
  ["30", "Haute", "Lancer les enquêtes 2025-2026 : apprentis, entreprises, formateurs ET financeurs (OPCO) — ind. 30 modifié ; les pièces actuelles portent sur 2023-2024."],
  ["32", "Haute", "Plan d'amélioration continue 2026 issu de l'analyse des appréciations et réclamations + ANALYSE DES RISQUES sur la qualité des formations (ind. 32 modifié) — dossier 32 vide."],
  ["33", "Haute", "NOUVEAU (CFA) : dispositif d'évaluation des contenus et des enseignements par les apprenants, distinct de la satisfaction, résultats partagés avec les formateurs, démarche d'amélioration et mesure de son efficacité — à créer (le module Formulaires de Vie scolaire peut le porter)."]
];
function handleGetQualiopiActions(p, ss) {
  try {
    if (!_qPeutLire_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
    var data = _qActionsFeuille_(ss).getDataRange().getValues(), out = [];
    for (var i = 1; i < data.length; i++) if (data[i][0])
      out.push({ id: String(data[i][0]), indicateur: String(data[i][1]), priorite: String(data[i][2]), action: String(data[i][3]), statut: String(data[i][4] || "À faire"),
        maj: _qDateIso_(data[i][5]), par: String(data[i][6] || "") });
    return createJsonResponse({ success: true, data: out });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}
function handleSaveQualiopiAction(p, ss) {
  try {
    if (!_qEstSuper_(p)) return createJsonResponse({ success: false, message: "Accès refusé." });
    var ws = _qActionsFeuille_(ss), data = ws.getDataRange().getValues();
    var date = Utilities.formatDate(new Date(), "Europe/Paris", "yyyy-MM-dd"), par = p.nomAuteur || p.emailAuteur || "";
    for (var i = 1; i < data.length; i++) if (String(data[i][0]) === String(p.id)) {
      if (p.statut) ws.getRange(i + 1, 5).setValue(p.statut);
      if (p.action) ws.getRange(i + 1, 4).setValue(String(p.action).slice(0, 1000));
      ws.getRange(i + 1, 6, 1, 2).setValues([[date, par]]);
      logAction(par, p.vieScRole || "", "Modification", "Vie scolaire — Qualiopi", "Action " + p.id, "Indicateur " + data[i][1] + " : « " + String(data[i][3]).slice(0, 80) + " » → " + (p.statut || "modifiée"));
      return createJsonResponse({ success: true });
    }
    var id = "QA-" + Date.now().toString(36);
    ws.appendRow([id, String(p.indicateur || ""), String(p.priorite || "Moyenne"), String(p.action || "").slice(0, 1000), "À faire", date, par]);
    logAction(par, p.vieScRole || "", "Création", "Vie scolaire — Qualiopi", "Action " + id, "Indicateur " + p.indicateur + " : " + String(p.action).slice(0, 120));
    return createJsonResponse({ success: true, id: id });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// ── Import unique des captures d'écran du site (pack préparé le 09/11/2026) ──
// Les images (site public hecg.fr) sont publiées avec le front : qualiopi/captures/<id>.png
// + qualiopi/captures.json. Exécuté une fois par le déclencheur récurrent (auditAdressesTick).
var Q_PACK_URL = "https://mokhtarimehdipro-bit.github.io/CRMHECG/qualiopi/";
function _qualiopiPackAuto_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty("QUALIOPI_PACK_V1")) return;
  props.setProperty("QUALIOPI_PACK_V1", "en cours " + new Date().toISOString());
  var ss = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");
  var bilan = { captures: 0, erreurs: [] };
  try {
    var r = UrlFetchApp.fetch(Q_PACK_URL + "captures.json?v=" + Date.now(), { muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) throw new Error("manifeste HTTP " + r.getResponseCode());
    var liste = JSON.parse(r.getContentText()), ws = _qFeuille_(ss), lignes = [];
    var CRIT = function (n) { n = Number(n); return n <= 3 ? 1 : n <= 8 ? 2 : n <= 16 ? 3 : n <= 20 ? 4 : n <= 22 ? 5 : n <= 29 ? 6 : 7; };
    liste.forEach(function (c) {
      try {
        var url = c.fichier;
        if (c.png) {
          var img = UrlFetchApp.fetch(Q_PACK_URL + "captures/" + c.png, { muteHttpExceptions: true });
          if (img.getResponseCode() !== 200) throw new Error("image HTTP " + img.getResponseCode());
          var f = _qDossier_(CRIT(c.ind), c.ind).createFile(img.getBlob().setName("Capture " + c.date + " — " + c.label + ".png"));
          f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          url = f.getUrl();
          bilan.captures++;
        }
        lignes.push([c.ind, c.label, c.url, Utilities.formatDate(new Date(), "Europe/Paris", "yyyy-MM-dd"), "lien", c.comment, c.date, 12, url || "", "Site hecg.fr (capture du " + c.date + ")", _qId_()]);
      } catch (e1) { bilan.erreurs.push(c.id + " : " + e1); }
    });
    if (lignes.length) ws.getRange(ws.getLastRow() + 1, 1, lignes.length, 11).setValues(lignes);
    props.setProperty("QUALIOPI_PACK_V1", "fait " + new Date().toISOString() + " — " + bilan.captures + " captures" + (bilan.erreurs.length ? " — erreurs : " + bilan.erreurs.join(" | ").slice(0, 500) : ""));
    logAction("Préparation Qualiopi (automatique)", "Super-admin", "Création", "Vie scolaire — Qualiopi", "QUALIOPI_PREUVES",
      lignes.length + " preuves issues du site hecg.fr ajoutées (" + bilan.captures + " captures d'écran surlignées enregistrées dans le Drive Qualiopi)" + (bilan.erreurs.length ? " — " + bilan.erreurs.length + " erreur(s)" : ""));
  } catch (e) {
    props.deleteProperty("QUALIOPI_PACK_V1");   // nouvel essai au prochain passage (5 min)
    console.error("pack Qualiopi : " + e);
  }
}
