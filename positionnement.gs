// =============================================================================
//  RÉFÉRENTS PÉDAGOGIQUES PAR CLASSE + FICHES DE POSITIONNEMENT (demande direction 2026-11-10)
//
//  1) Référent pédagogique (ETUDIANTS col. H « Ref_Perso ») fixé selon la classe :
//       BTS CG 1 → NMAS · BTS CG 2 → AMHO · DCG 1 à 3 → RBEN · BTS GPME → LKEB
//       DSCG 1 et 2 → MMOK · Bachelor → DBEN · tout étudiant aixois (campus Aix ou classe AIX_…) → VVEL
//     Appliqué à la création d'un étudiant et à chaque changement de classe / campus
//     (code.gs handleAddStudent / handleUpdateStudent), plus une passe de contrôle sur tous
//     les étudiants non sortis (rapport : feuille CONTROLE_REFERENTS).
//
//  2) Fiche de positionnement (modèle Excel « ETUDIANT MODELE - Fiche de positionnement - HECG »,
//     converti une fois en Google Sheet) créée dans les Documents (GED) de chaque étudiant
//     INSCRIT ou PRÉINSCRIT. Date de rédaction = date de création de l'étudiant (en haut de la fiche).
//     Statut : alternance (contrat d'apprentissage) ou formation initiale. Aucune donnée titre de
//     séjour / RQTH. Grille d'évaluation tirée au hasard (graine = ID étudiant, donc stable) mais
//     cohérente, observations assorties, note toujours > 10/20. Interlocuteur = référent pédagogique.
//     Registre : feuille FICHES_POSITIONNEMENT. Traitement par lots via auditAdressesTick.
// =============================================================================
var FP_MODELE_XLSX = "1Vi2RtLQgrdWLcCLnqaNPCay7-ZaZe_FL";
var FP_HEADERS = ["ID étudiant", "Nom", "Prénom", "Classe", "Statut fiche", "Note /20", "Référent", "Date de rédaction", "Lien", "ID fichier", "Créée le", "Remarque"];

// ── 1) Référent selon la classe ──────────────────────────────────────────────
function _referentPourClasse_(classe, campus) {
  var c = String(classe || "").toUpperCase().replace(/[\s_\-]/g, "");
  var aix = String(campus || "").toUpperCase() === "HECG_AIX" || c.indexOf("AIX") !== -1;
  if (!c) return aix ? "VVEL" : "";
  if (aix) return "VVEL";
  if (/BACHELOR/.test(c)) return "DBEN";
  if (/DSCG/.test(c)) return "MMOK";
  if (/DCG[123]/.test(c)) return "RBEN";
  if (/GPME/.test(c)) return "LKEB";
  if (/BTSCG1|CG1/.test(c)) return "NMAS";
  if (/BTSCG2|CG2/.test(c)) return "AMHO";
  return "";
}

// Passe de contrôle sur tous les étudiants non sortis : corrige la col. H et trace les écarts.
function _referentsPasse_(ss, auteur) {
  var ws = getSheetSafe(ss, "ETUDIANTS"), d = ws.getDataRange().getValues();
  var lignes = [], corriges = 0, sansRegle = 0, conformes = 0;
  for (var i = 1; i < d.length; i++) {
    var id = String(d[i][0] || "").trim(); if (!id) continue;
    var typo = String(d[i][20] || "").toUpperCase();
    if (/SORTI|REFUS/.test(typo) || (d[i][18] && String(d[i][18]).trim())) continue;
    var cible = _referentPourClasse_(d[i][14], d[i][22]);
    var actuel = String(d[i][7] || "").trim();
    if (!cible) { if (String(d[i][14] || "").trim()) { sansRegle++; lignes.push([id, d[i][2], d[i][3], d[i][14], d[i][22], actuel, "", "Classe sans règle : à vérifier"]); } continue; }
    if (actuel === cible) { conformes++; continue; }
    ws.getRange(i + 1, 8).setValue(cible);
    corriges++;
    lignes.push([id, d[i][2], d[i][3], d[i][14], d[i][22], actuel || "(vide)", cible, "Corrigé"]);
    try { _invalidateStudentDetailsCache_(id); } catch (e) {}
  }
  var wr = ss.getSheetByName("CONTROLE_REFERENTS") || ss.insertSheet("CONTROLE_REFERENTS");
  wr.clear();
  wr.getRange(1, 1, 1, 8).setValues([["ID", "Nom", "Prénom", "Classe", "Campus", "Ancien référent", "Nouveau référent", "Résultat"]]).setFontWeight("bold").setBackground("#1A4E8A").setFontColor("#ffffff");
  if (lignes.length) wr.getRange(2, 1, lignes.length, 8).setValues(lignes);
  wr.setFrozenRows(1);
  var resume = corriges + " référent(s) corrigé(s), " + conformes + " déjà conforme(s), " + sansRegle + " classe(s) sans règle";
  logAction(auteur || "Contrôle automatique", "", "Modification", "Étudiants — référents pédagogiques", "ETUDIANTS", "Passe de contrôle des référents par classe : " + resume);
  PropertiesService.getScriptProperties().setProperty("REFERENTS_PASSE_RESUME", Utilities.formatDate(new Date(), "Europe/Paris", "dd/MM/yyyy HH:mm") + " — " + resume);
  return { corriges: corriges, conformes: conformes, sansRegle: sansRegle };
}

// ── 2) Fiche de positionnement ───────────────────────────────────────────────
function _fpRegistre_(ss) {
  var ws = ss.getSheetByName("FICHES_POSITIONNEMENT");
  if (!ws) {
    ws = ss.insertSheet("FICHES_POSITIONNEMENT");
    ws.getRange(1, 1, 1, FP_HEADERS.length).setValues([FP_HEADERS]).setFontWeight("bold").setBackground("#1A4E8A").setFontColor("#ffffff");
    ws.setFrozenRows(1);
  }
  return ws;
}

// Modèle Excel → Google Sheet (une seule fois), rangé à côté du modèle
function _fpModele_() {
  var props = PropertiesService.getScriptProperties(), id = props.getProperty("FICHE_POS_MODELE_ID");
  if (id) { try { if (!DriveApp.getFileById(id).isTrashed()) return id; } catch (e) {} }
  var tok = ScriptApp.getOAuthToken();
  var r = UrlFetchApp.fetch("https://www.googleapis.com/drive/v3/files/" + FP_MODELE_XLSX + "/copy?supportsAllDrives=true", {
    method: "post", contentType: "application/json", muteHttpExceptions: true, headers: { Authorization: "Bearer " + tok },
    payload: JSON.stringify({ name: "MODÈLE CRM — Fiche de positionnement (Google Sheet)", mimeType: "application/vnd.google-apps.spreadsheet" }) });
  if (r.getResponseCode() !== 200) throw new Error("Conversion du modèle impossible : HTTP " + r.getResponseCode() + " " + r.getContentText().slice(0, 200));
  id = JSON.parse(r.getContentText()).id;
  props.setProperty("FICHE_POS_MODELE_ID", id);
  return id;
}

function _fpRng_(graine) {   // mulberry32 : tirage reproductible par étudiant
  var h = 1779033703 ^ graine.length;
  for (var i = 0; i < graine.length; i++) { h = Math.imul(h ^ graine.charCodeAt(i), 3432918353); h = (h << 13) | (h >>> 19); }
  var a = h >>> 0;
  return function () { a |= 0; a = (a + 0x6D2B79F5) | 0; var t = Math.imul(a ^ (a >>> 15), 1 | a); t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t; return ((t ^ (t >>> 14)) >>> 0) / 4294967296; };
}

function _fpFormation_(classe) {
  var c = String(classe || "").toUpperCase().replace(/^AIX_?/, "").replace(/AIX$/, "");
  if (/BACHELOR/.test(c)) return "BACHELOR";
  var m;
  if ((m = c.match(/^BTS ?(\d)?GPME(\d)?/))) return "BTS GPME " + (m[1] || m[2] || "1");
  if ((m = c.match(/^BTSCG(\d)([A-Z]?)$/))) return "BTS CG " + m[1] + (m[2] || "");
  if ((m = c.match(/^(DSCG|DCG)(\d)([A-Z]?)$/))) return m[1] + " " + m[2] + (m[3] || "");
  return String(classe || "");
}
function _fpFiliere_(formation) { var m = String(formation).match(/^(BTS CG|BTS GPME|DSCG|DCG|BACHELOR)/); return m ? (m[1] === "BACHELOR" ? "Bachelor" : m[1]) : formation; }

// Observations par critère et par niveau (0 = ++, 1 = +, 2 = +/-, 3 = -)
var FP_CRITERES = [
  ["Entretien très bien préparé, dossier complet.", "Entretien bien préparé, dossier complet.", "Préparation correcte, quelques pièces restaient à fournir.", "Préparation insuffisante, dossier incomplet."],
  ["Attitude très professionnelle, candidat(e) motivé(e).", "Attitude sérieuse et positive.", "Attitude correcte mais réservée.", "Attitude à travailler."],
  ["S'exprime avec aisance et clarté.", "S'exprime clairement.", "Expression correcte, parfois hésitante.", "Expression orale à améliorer."],
  ["Projet professionnel clair et très bien argumenté.", "Projet cohérent avec la formation {F}.", "Projet professionnel encore à préciser.", "Projet professionnel peu défini."],
  ["Profil en parfaite adéquation avec la filière {F}.", "Le profil est en adéquation avec la filière {F}.", "Adéquation partielle avec la filière {F}.", "Adéquation avec la filière {F} à confirmer."],
  ["Très bonnes bases théoriques.", "Bonnes bases théoriques.", "Bases correctes, quelques fragilités.", "Des fragilités dans certaines disciplines."],
  ["Expérience professionnelle significative.", "Première expérience professionnelle pertinente.", "Expérience professionnelle limitée.", "Pas encore d'expérience professionnelle."],
  ["Très bons résultats scolaires.", "Bons résultats scolaires.", "Résultats corrects mais irréguliers.", "Résultats scolaires fragiles."],
  ["Appréciations très positives des enseignants.", "Appréciations positives des enseignants.", "Appréciations encourageantes.", "Appréciations mitigées."]
];
var FP_FORTS = ["Bonne préparation de l'entretien", "Attitude sérieuse et motivée", "Aisance à l'oral", "Projet professionnel construit", "Profil adapté à la filière", "Solides bases théoriques", "Expérience professionnelle", "Bons résultats scolaires", "Appréciations positives"];
var FP_AXES = ["Mieux préparer ses démarches et son dossier", "Gagner en assurance", "Travailler l'expression orale", "Préciser son projet professionnel", "Consolider sa connaissance de la filière", "Renforcer les bases théoriques", "Acquérir de l'expérience en entreprise", "Fournir un travail personnel plus régulier", "Approfondir son travail personnel"];

// Grille cohérente : profil de base + petites variations, paires liées (attitude/expression,
// résultats/appréciations) à un niveau d'écart au plus, note = formule du modèle, 11 ≤ note ≤ 17,5.
function _fpGrille_(id, alternance) {
  var rnd = _fpRng_(String(id)), pts = [20, 15, 10, 5];
  for (var essai = 0; essai < 200; essai++) {
    var base = rnd() < 0.3 ? 0 : (rnd() < 0.75 ? 1 : 2), n = [];
    for (var k = 0; k < 9; k++) { var dlt = [-1, 0, 0, 0, 1][Math.floor(rnd() * 5)]; n.push(Math.max(0, Math.min(3, base + dlt))); }
    if (alternance && n[6] > 1) n[6] = 1;                           // en contrat : expérience pratique au moins « + »
    if (Math.abs(n[1] - n[2]) > 1) n[2] = n[1] + (n[2] > n[1] ? 1 : -1);
    if (Math.abs(n[7] - n[8]) > 1) n[8] = n[7] + (n[8] > n[7] ? 1 : -1);
    if (n[4] > 2) n[4] = 2;                                          // admis : adéquation au moins « +/- »
    var note = n.reduce(function (s, v) { return s + pts[v]; }, 0) / 8;   // formule de la cellule F22 du modèle
    if (note >= 11 && note <= 17.5) return { niveaux: n, note: note };
  }
  return { niveaux: [1, 1, 1, 1, 1, 2, 2, 2, 2], note: 13.75 };
}

function _fpDate_(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  var s = String(v || "").trim(), m;
  if ((m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/))) return new Date(+m[3], +m[2] - 1, +m[1]);
  if ((m = s.match(/^(\d{4})-(\d{2})-(\d{2})/))) return new Date(+m[1], +m[2] - 1, +m[3]);
  return null;
}

// Dossier GED de l'étudiant (lié, retrouvé par nom, ou créé et lié)
function _fpDossierGed_(ss, id, nom, prenom) {
  var ws = getOrCreateGedEtuSheet(ss), d = ws.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) if (String(d[i][0]).trim() === id && String(d[i][2] || "").trim()) {
    try { return DriveApp.getFolderById(String(d[i][2]).trim()); } catch (e) { break; }
  }
  var root = DriveApp.getFolderById(GED_ETU_ROOT_ID), libelles = [String(nom).toUpperCase() + " " + prenom, prenom + " " + nom, nom + " " + prenom], f = null;
  for (var k = 0; k < libelles.length && !f; k++) { var it = root.getFoldersByName(libelles[k].trim()); if (it.hasNext()) f = it.next(); }
  if (!f) f = root.createFolder(libelles[0].trim());
  ws.appendRow([id, (prenom + " " + nom).trim(), f.getId(), f.getUrl()]);
  return f;
}

function _fpGenerer_(ss, r, persoMap, modele) {
  var id = String(r[0]).trim(), nom = String(r[2] || "").trim(), prenom = String(r[3] || "").trim();
  var classe = String(r[14] || ""), campus = String(r[22] || ""), statut = String(r[13] || "");
  var alternance = /altern|contrat/i.test(statut);
  var entreprise = String(r[21] || "").trim();
  var refId = _referentPourClasse_(classe, campus) || String(r[7] || "").trim();
  var ref = persoMap[refId] || "";
  var dCrea = _fpDate_(r[modele.colCrea]) || _fpDate_(r[17]) || new Date();
  var formation = _fpFormation_(classe), filiere = _fpFiliere_(formation);
  var aix = campus.toUpperCase() === "HECG_AIX" || /AIX/i.test(classe);
  var g = _fpGrille_(id, alternance), n = g.niveaux;
  var annee = dCrea.getMonth() >= 7 ? dCrea.getFullYear() + "-" + (dCrea.getFullYear() + 1) : (dCrea.getFullYear() - 1) + "-" + dCrea.getFullYear();

  // Valeurs de la zone B2:L37 du modèle (formules conservées), complétées
  var V = modele.valeurs.map(function (l) { return l.slice(); });
  function set(a1, v) { var col = a1.charCodeAt(0) - 66, row = Number(a1.slice(1)) - 2; V[row][col] = v; }
  set("B4", "Date de rédaction : " + Utilities.formatDate(dCrea, "Europe/Paris", "dd/MM/yyyy"));
  set("B7", formation); set("E7", aix ? "AIX-EN-PROVENCE" : "MARSEILLE");
  set("D8", alternance ? "Alternance (contrat d'apprentissage)" : "Formation initiale");
  set("H7", dCrea); set("H8", ref); set("L8", "Admis");
  ["C11", "C14", "C15", "C16", "C17", "C18", "C19", "C22", "C23", "C24", "C30", "D30"].forEach(function (a) { set(a, ""); });   // ni titre de séjour ni RQTH
  set("C12", nom.toUpperCase()); set("C13", prenom);
  set("C25", alternance ? "Oui" : "Non"); set("C26", alternance ? entreprise : "");
  set("C29", formation); set("D29", annee);
  for (var k = 0; k < 9; k++) {
    ["F", "G", "H", "I"].forEach(function (c, lv) { set(c + (13 + k), lv === n[k] ? "X" : ""); });
    set("J" + (13 + k), FP_CRITERES[k][n[k]].replace("{F}", filiere));
  }
  var ordre = n.map(function (v, k) { return { k: k, v: v }; });
  var forts = ordre.filter(function (o) { return o.v <= 1; }).sort(function (a, b) { return a.v - b.v; }).slice(0, 3).map(function (o) { return FP_FORTS[o.k]; });
  var axes = ordre.filter(function (o) { return o.v >= 2; }).sort(function (a, b) { return b.v - a.v; }).slice(0, 2).map(function (o) { return FP_AXES[o.k]; });
  set("E26", (forts.length ? forts : ["Motivation pour la formation"]).join("\n"));
  set("H26", (axes.length ? axes : ["Maintenir un travail personnel régulier"]).join("\n"));
  var synth = g.note >= 15 ? "Profil solide qui convient pleinement à la formation " + formation + "."
    : (g.note >= 12.5 ? "Profil qui convient à la formation " + formation + ", avec une bonne marge de progression."
    : "Profil qui convient à la formation " + formation + ", sous réserve d'un travail personnel régulier.");
  synth += alternance ? " Formation suivie en alternance" + (entreprise ? " (contrat d'apprentissage chez " + entreprise + ")." : " (contrat d'apprentissage).")
                      : " Formation suivie en initial ; accompagnement à la recherche d'une entreprise d'accueil.";
  set("B32", synth);

  var dossier = _fpDossierGed_(ss, id, nom, prenom);
  var fichier = DriveApp.getFileById(modele.id).makeCopy("Fiche de positionnement - " + nom.toUpperCase() + " " + prenom, dossier);
  var sh = SpreadsheetApp.openById(fichier.getId()).getSheets()[0];
  sh.getRange(2, 2, V.length, V[0].length).setValues(V);
  sh.getRange("B4").setFontWeight("bold").setFontColor("#1A4E8A").setHorizontalAlignment("right");
  SpreadsheetApp.flush();
  return { fichier: fichier, note: g.note, ref: ref, dCrea: dCrea, alternance: alternance };
}

function _fpModeleCharge_(ss) {
  var id = _fpModele_(), sh = SpreadsheetApp.openById(id).getSheets()[0];
  var rg = sh.getRange("B2:L37"), v = rg.getValues(), f = rg.getFormulas();
  for (var i = 0; i < v.length; i++) for (var j = 0; j < v[i].length; j++) if (f[i][j]) v[i][j] = f[i][j];
  var hE = getSheetSafe(ss, "ETUDIANTS").getRange(1, 1, 1, 40).getValues()[0];
  var colCrea = hE.map(function (h) { return String(h || "").trim().toLowerCase(); }).indexOf(DATE_CREATION_HEADER.toLowerCase());
  return { id: id, valeurs: v, colCrea: colCrea };
}

function _fpEstCible_(r) {
  var typo = String(r[20] || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toUpperCase().replace(/[^A-Z]/g, "")
    || String(calculerTypoEtu(r[17], r[18], r[14])).normalize("NFD").replace(/[̀-ͯ]/g, "").toUpperCase().replace(/[^A-Z]/g, "");
  return (typo === "INSCRIT" || typo === "PREINSCRIT") && String(r[14] || "").trim() !== "";
}

// Lot : crée les fiches manquantes jusqu'à la limite de temps ; renvoie le nombre restant
function _fpLot_(ss, limiteMs) {
  var t0 = Date.now(), reg = _fpRegistre_(ss), dR = reg.getDataRange().getValues(), deja = {};
  // Garde-fou : une première fiche « test » est créée seule ; la génération complète ne démarre
  // qu'une fois « OK » écrit dans sa colonne Remarque (contrôle visuel du rendu).
  var test = dR.length < 2;
  if (!test && String(dR[1][11] || "").indexOf("OK") === -1 && String(dR[1][11] || "").indexOf("TEST") !== -1) return 1;
  for (var i = 1; i < dR.length; i++) if (dR[i][0]) deja[String(dR[i][0]).trim()] = 1;
  var dE = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues();
  var cibles = dE.slice(1).filter(function (r) { var id = String(r[0] || "").trim(); return id && id.toUpperCase() !== "ETU000X" && !deja[id] && _fpEstCible_(r); });
  if (!cibles.length) return 0;
  var dP = getPersonnelDataCached_(ss), persoMap = {};
  for (var p = 1; p < dP.length; p++) if (dP[p][0]) persoMap[String(dP[p][0]).trim()] = (String(dP[p][2] || "") + " " + String(dP[p][1] || "").toUpperCase()).trim();
  var modele = _fpModeleCharge_(ss), faits = 0, erreurs = 0;
  if (test) cibles = cibles.filter(function (r) { return /altern|contrat/i.test(String(r[13])); }).slice(0, 1).concat(cibles).slice(0, 1);
  for (var c = 0; c < cibles.length && Date.now() - t0 < limiteMs; c++) {
    var r = cibles[c], id = String(r[0]).trim();
    try {
      var res = _fpGenerer_(ss, r, persoMap, modele);
      reg.appendRow([id, r[2], r[3], r[14], res.alternance ? "Alternance" : "Initial", res.note, res.ref, Utilities.formatDate(res.dCrea, "Europe/Paris", "dd/MM/yyyy"), res.fichier.getUrl(), res.fichier.getId(), new Date(), test ? "FICHE TEST — écrire OK ici pour lancer toutes les fiches" : ""]);
      faits++;
    } catch (e) {
      erreurs++;
      reg.appendRow([id, r[2], r[3], r[14], "ERREUR", "", "", "", "", "", new Date(), String(e).slice(0, 300)]);
    }
  }
  if (faits || erreurs) logAction("Génération automatique", "", "Création", "Étudiants — fiches de positionnement", "GED_ETU",
    faits + " fiche(s) de positionnement créée(s) dans les documents étudiants" + (erreurs ? ", " + erreurs + " erreur(s) (voir FICHES_POSITIONNEMENT)" : "") + " — reste " + Math.max(0, cibles.length - faits - erreurs));
  return Math.max(0, cibles.length - faits - erreurs);
}

// Appelé par auditAdressesTick (toutes les 5 min) : passe des référents une fois, puis fiches par lots ;
// ensuite un contrôle par jour pour les nouveaux inscrits / préinscrits.
function _fpAuto_(t0) {
  var budget = Math.min(150000, 230000 - (Date.now() - (t0 || Date.now())));   // reste sous la limite de 6 min du passage
  if (budget < 40000) return;
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty("FICHES_POS_ACTIVE") === "0") return;
  var ss = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");
  if (props.getProperty("REFERENTS_PASSE_V1") !== "1") { _referentsPasse_(ss, "Contrôle des référents (mise en place)"); props.setProperty("REFERENTS_PASSE_V1", "1"); return; }
  var auj = Utilities.formatDate(new Date(), "Europe/Paris", "yyyy-MM-dd");
  if (props.getProperty("FICHES_POS_SUITE") !== "1" && props.getProperty("FICHES_POS_JOUR") === auj) return;
  props.setProperty("FICHES_POS_JOUR", auj);
  var reste = _fpLot_(ss, budget);
  if (reste > 0) props.setProperty("FICHES_POS_SUITE", "1"); else props.deleteProperty("FICHES_POS_SUITE");
}
