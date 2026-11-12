// =============================================================================
//  CONTRÔLE CHARLEMAGNE ↔ CRM (Vie scolaire › onglet « Contrôle Charlemagne »)
//  Demande direction 2026-11-09. La base Charlemagne (HFSQL chiffrée, éditeur Aplon) n'est pas
//  lisible directement : on importe l'édition « Tableau de bord COMPLET MMOK » exportée en Excel.
//  La page ne transmet QUE les colonnes utiles (jamais n° Sécu, RQTH, handicap, titres de séjour).
//  Rien n'est modifié dans ETUDIANTS : le contrôle produit un rapport (feuille CONTROLE_CHARLEMAGNE).
//
//  Règles :
//   • rapprochement étudiant : Identifiant Elève (col. X d'ETUDIANTS) → e-mail → nom + prénom ;
//   • « SANS ENTREPRISE » (ou vide) dans Société = pas d'alternance ;
//   • alternant : la Société Charlemagne est rapprochée de l'entreprise du CRM avec une
//     comparaison approchante (formes juridiques, accents, « cabinet », sigles ignorés) ;
//   • classes : les deux logiciels n'ont pas les mêmes libellés → la correspondance dominante
//     (classe Charlemagne → classe CRM) est apprise, seuls les écarts à cette règle sont signalés ;
//   • date de sortie Charlemagne absente du CRM = ERREUR (l'étudiant garde l'accès à son espace).
//   • (12/11/2026) seuls les étudiants INSCRITS / PRÉINSCRITS du CRM (typologie, col. U) sont comparés ;
//     l'effectif doit être identique à celui des étudiants actifs de Charlemagne (sans sortie passée) ;
//     un comparatif « une ligne par étudiant » indique l'entreprise de chaque côté (feuille CHARLEMAGNE_COMPARATIF).
// =============================================================================

var CHARL_ROLES = { "Super-admin": 1, "Administratif": 1 };
var CHARL_FEUILLE = "CONTROLE_CHARLEMAGNE";
var CHARL_IMPORT = "CHARLEMAGNE_IMPORT";
var CHARL_COMPARATIF = "CHARLEMAGNE_COMPARATIF";
var CHARL_COMP_COLS = ["verdict", "idEtu", "idEleve", "nom", "typo", "classeCrm", "classeCharl", "entCharl", "entCrm", "statutCrm", "ressemblance"];
var CHARL_COMP_TITRES = ["Concordance entreprise", "ID CRM", "Identifiant Charlemagne", "Étudiant", "Typologie CRM", "Classe CRM", "Classe Charlemagne", "Entreprise Charlemagne", "Entreprise CRM", "Statut CRM", "Ressemblance %"];
var CHARL_VERDICTS = ["Différente", "Charlemagne seulement", "CRM seulement", "Absent du CRM", "Absent de Charlemagne", "Sorti dans Charlemagne", "Non inscrit dans le CRM", "Doublon Charlemagne", "Doublon CRM", "Approchant", "Identique", "Aucune des deux"];
function _charlInscrit_(typo) { var t = _charlNorm_(typo); return t === "inscrit" || t === "preinscrit" || t === "pre inscrit"; }
var CHARL_CHAMPS = ["idEleve", "nom", "prenom", "naissance", "email", "mailPerso", "classe", "codeClasse", "niveau",
  "entree", "sortie", "motif", "etatPreinscription", "societe", "nomResp", "prenomResp", "emailResp", "professionResp",
  "visites"];   // visites = JSON [{h: en-tête, v: valeur}] des colonnes « … visite … » (suivi des RDV, vs_suivi.gs)

function _charlAutorise_(p) {
  return !!CHARL_ROLES[String(p.vieScRole || "")] || String(p.emailAuteur || "").toLowerCase() === "m.mokhtari@hecg.fr";
}
function _charlNorm_(s) {
  return String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase()
    .replace(/[’']/g, " ").replace(/[^a-z0-9]+/g, " ").trim();
}
function _charlDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  var t = String(v || "").trim(), m;
  if (!t) return null;
  if ((m = t.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})/))) return new Date(+m[3], +m[2] - 1, +m[1]);
  if ((m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/))) return new Date(+m[1], +m[2] - 1, +m[3]);
  if (/^\d{5}(\.\d+)?$/.test(t)) { var d = new Date(Math.round((+t - 25569) * 86400000)); return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()); }   // n° de série Excel
  return null;
}
function _charlFmt_(d) { return d ? Utilities.formatDate(d, "Europe/Paris", "dd/MM/yyyy") : ""; }

// ── Entreprises : comparaison approchante ───────────────────────────────────
var CHARL_MOTS_ENT = { sas: 1, sasu: 1, sarl: 1, eurl: 1, sa: 1, sci: 1, scp: 1, selarl: 1, selas: 1, selafa: 1, snc: 1, sca: 1,
  cabinet: 1, groupe: 1, group: 1, societe: 1, ste: 1, et: 1, de: 1, du: 1, des: 1, la: 1, le: 1, les: 1, l: 1, d: 1, en: 1,
  expert: 1, experts: 1, expertise: 1, comptable: 1, comptables: 1, comptabilite: 1, conseil: 1, conseils: 1, audit: 1,
  associes: 1, associe: 1, cie: 1, france: 1, agence: 1, office: 1, ec: 1, cac: 1, and: 1, co: 1,
  // lieux (« Exco Provence » = « Exco Marseille ») : n'entrent pas dans la comparaison
  provence: 1, marseille: 1, aix: 1, paca: 1, sud: 1, mediterranee: 1, region: 1, alpes: 1, cote: 1, azur: 1, est: 1, ouest: 1, nord: 1,
  paris: 1, lyon: 1, nice: 1, toulon: 1, avignon: 1, vitrolles: 1, aubagne: 1, bouches: 1, rhone: 1, var: 1 };
function _charlMotsEnt_(s) {
  var t = _charlNorm_(String(s || "").replace(/\(.*?\)/g, " ")).split(" ").filter(Boolean);
  var u = t.filter(function (w) { return !CHARL_MOTS_ENT[w]; });
  return u.length ? u : t;
}
function _charlBigrammes_(s) { var o = {}; for (var i = 0; i < s.length - 1; i++) { var b = s.substr(i, 2); o[b] = (o[b] || 0) + 1; } return o; }
// Similarité 0..1 entre deux noms d'entreprise
function _charlSimEnt_(a, b) {
  var ma = _charlMotsEnt_(a), mb = _charlMotsEnt_(b);
  if (!ma.length || !mb.length) return 0;
  var ja = ma.join(""), jb = mb.join("");
  if (ja === jb) return 1;
  var sa = {}, sb = {};
  ma.forEach(function (w) { sa[w] = 1; }); mb.forEach(function (w) { sb[w] = 1; });
  var communs = ma.filter(function (w) { return sb[w] && w.length >= 3; }).length;
  if (communs && (communs === Object.keys(sa).length || communs === Object.keys(sb).length)) return 0.9;   // l'un contient l'autre
  if (ja.length >= 4 && jb.length >= 4 && (ja.indexOf(jb) !== -1 || jb.indexOf(ja) !== -1)) return 0.85;
  var ba = _charlBigrammes_(ja), bb = _charlBigrammes_(jb), inter = 0, na = 0, nb = 0, k;
  for (k in ba) { na += ba[k]; if (bb[k]) inter += Math.min(ba[k], bb[k]); }
  for (k in bb) nb += bb[k];
  var dice = (na + nb) ? 2 * inter / (na + nb) : 0;
  return Math.max(dice, communs ? 0.5 + 0.1 * communs : 0);
}
var CHARL_SEUIL_ENT = 0.62;
function _charlSansEntreprise_(s) {
  var n = _charlNorm_(s);
  return !n || n === "sans entreprise" || n === "aucune" || n === "aucun" || n === "neant" || n === "x";
}

// ── Regroupement par Identifiant Elève ──────────────────────────────────────
// L'export Charlemagne répète un même élève sur plusieurs lignes (même Identifiant Elève : une ligne par
// responsable, par exemple). On raisonne par identifiant : les lignes sont fusionnées (1re valeur non vide
// de chaque champ ; pour la société, la 1re qui n'est pas « SANS ENTREPRISE » ; colonnes « visite » cumulées).
// Une ligne sans identifiant reste une entrée à part.
function _charlFusionner_(lignes) {
  var parId = {}, out = [];
  lignes.forEach(function (l) {
    var id = String(l.idEleve || "").trim();
    if (!id) { out.push(l); return; }
    var m = parId[id];
    if (!m) { m = parId[id] = {}; CHARL_CHAMPS.forEach(function (c) { m[c] = l[c] || ""; }); m._n = 1; out.push(m); return; }
    m._n++;
    CHARL_CHAMPS.forEach(function (c) {
      if (c === "visites" || c === "societe") return;
      if (!m[c] && l[c]) m[c] = l[c];
    });
    if ((!m.societe || _charlSansEntreprise_(m.societe)) && l.societe && !_charlSansEntreprise_(l.societe)) m.societe = l.societe;
    else if (!m.societe && l.societe) m.societe = l.societe;
    if (l.visites) {
      var a = [], b = [];
      try { a = JSON.parse(m.visites || "[]"); } catch (e) {}
      try { b = JSON.parse(l.visites || "[]"); } catch (e) {}
      var vus = {}; a.forEach(function (x) { vus[x.h + "|" + x.v] = 1; });
      b.forEach(function (x) { if (!vus[x.h + "|" + x.v]) { vus[x.h + "|" + x.v] = 1; a.push(x); } });
      m.visites = a.length ? JSON.stringify(a) : "";
    }
  });
  out.forEach(function (m) { delete m._n; });
  return out;
}

// ── Point d'entrée : import + contrôle ──────────────────────────────────────
function handleControleCharlemagne(p, ss) {
  if (!_charlAutorise_(p)) return createJsonResponse({ success: false, message: "Réservé aux profils Super-admin et Administratif." });
  var lignes = [];
  try { lignes = typeof p.lignes === "string" ? JSON.parse(p.lignes) : (p.lignes || []); } catch (e) { lignes = []; }
  lignes = lignes.filter(function (l) { return l && (l.nom || l.idEleve); }).map(function (l) {
    var o = {}; CHARL_CHAMPS.forEach(function (c) { o[c] = l[c] == null ? "" : String(l[c]).trim(); }); return o;   // liste blanche des champs
  });
  if (!lignes.length) return createJsonResponse({ success: false, message: "Le fichier ne contient aucun étudiant exploitable." });
  var nbLignes = lignes.length;
  lignes = _charlFusionner_(lignes);   // une entrée par Identifiant Elève (l'export répète l'élève sur plusieurs lignes)
  try {
    var R = _charlControler_(ss, lignes);
    R.stats.lignesFichier = nbLignes; R.stats.lignesFusionnees = nbLignes - lignes.length;
    R.resume = R.resume.replace(/^/, nbLignes + " lignes dans le fichier = " + lignes.length + " identifiants élève — ");
    R.fichier = String(p.fichier || "").slice(0, 120);
    R.importePar = p.nomAuteur || p.emailAuteur || "";
    R.date = Utilities.formatDate(new Date(), "Europe/Paris", "dd/MM/yyyy HH:mm");
    _charlEcrireImport_(ss, lignes, R);
    _charlEcrireRapport_(ss, R);
    _charlEcrireComparatif_(ss, R);
    try { R.rdv = _rdvImporterCharl_(ss, lignes, R.importePar); } catch (eV) { R.rdv = { erreur: String(eV) }; }
    logAction(p.nomAuteur || p.emailAuteur || "", p.vieScRole || "", "Action", "Vie scolaire — Contrôle Charlemagne", CHARL_FEUILLE,
      "Import « " + R.fichier + " » (" + lignes.length + " étudiants Charlemagne) — " + R.resume);
    return createJsonResponse({ success: true, rapport: R });
  } catch (e) {
    return createJsonResponse({ success: false, message: e.toString() });
  }
}

// Dernier contrôle (réouverture de l'onglet)
function handleGetControleCharlemagne(p, ss) {
  if (!_charlAutorise_(p)) return createJsonResponse({ success: false, message: "Réservé aux profils Super-admin et Administratif." });
  try {
    var json = PropertiesService.getScriptProperties().getProperty("CHARL_DERNIER");
    if (!json) return createJsonResponse({ success: true, rapport: null });
    var meta = JSON.parse(json);
    var ws = ss.getSheetByName(CHARL_FEUILLE);
    var anomalies = [];
    if (ws && ws.getLastRow() > 4) {
      ws.getRange(5, 1, ws.getLastRow() - 4, 10).getValues().forEach(function (r) {
        anomalies.push({ gravite: r[0], type: r[1], idEtu: r[2], idEleve: r[3], nom: r[4], classe: r[5], charlemagne: r[6], crm: r[7], detail: r[8], suggestion: r[9] });
      });
    }
    meta.anomalies = anomalies;
    var wc = ss.getSheetByName(CHARL_COMPARATIF), comp = [];
    if (wc && wc.getLastRow() > 1) {
      wc.getRange(2, 1, wc.getLastRow() - 1, CHARL_COMP_COLS.length).getValues().forEach(function (r) {
        var o = {}; CHARL_COMP_COLS.forEach(function (c, i) { o[c] = String(r[i] == null ? "" : r[i]); }); comp.push(o);
      });
    }
    meta.comparatif = comp;
    return createJsonResponse({ success: true, rapport: meta });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function _charlControler_(ss, lignes) {
  var dE = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues();
  var dP = getSheetSafe(ss, "Partenariat").getDataRange().getValues();
  var nomsEnt = {}, listeEnt = [];
  for (var i = 1; i < dP.length; i++) {
    var idE = String(dP[i][1] || "").trim().toUpperCase(), nE = String(dP[i][2] || "").trim();
    if (idE && nE && !nomsEnt[idE]) { nomsEnt[idE] = nE; listeEnt.push({ id: idE, nom: nE }); }
  }
  // Index CRM
  var crm = [], parId = {}, parMail = {}, parNom = {};
  for (var j = 1; j < dE.length; j++) {
    var r = dE[j];
    if (!r[0] || String(r[0]).toUpperCase() === "ETU000X") continue;
    var e = { ligne: j + 1, id: String(r[0]).trim(), nom: String(r[2] || "").trim(), prenom: String(r[3] || "").trim(),
      mail: String(r[4] || "").trim().toLowerCase(), mailHecg: String(r[26] || "").trim().toLowerCase(),
      idEnt: String(r[11] || "").trim().toUpperCase(), statut: String(r[13] || "").trim(), classe: String(r[14] || "").trim(),
      entree: _charlDate_(r[17]), sortie: _charlDate_(r[18]), sortieBrute: String(r[18] || "").trim(),
      nomEnt: String(r[21] || "").trim(), idEleve: String(r[23] || "").trim(), typo: String(r[20] || "").trim(), vu: false };
    e.inscrit = _charlInscrit_(e.typo);
    e.entreprise = (e.idEnt && nomsEnt[e.idEnt]) || e.nomEnt;
    crm.push(e);
    if (e.idEleve) parId[e.idEleve] = e;
    [e.mail, e.mailHecg].forEach(function (m) { if (m && m.indexOf("@") > 0) parMail[m] = e; });
    var kn = _charlNorm_(e.nom + " " + e.prenom);
    if (kn) (parNom[kn] = parNom[kn] || []).push(e);
    var kn2 = _charlNorm_(e.prenom + " " + e.nom);
    if (kn2 && kn2 !== kn) (parNom[kn2] = parNom[kn2] || []).push(e);
  }

  var A = [], paires = [], C = [], stats = { charlemagne: lignes.length, rapproches: 0, parId: 0, parMail: 0, parNom: 0, alternantsCharl: 0, sansEntCharl: 0, entOk: 0, entApprochant: 0,
    inscritsCrm: 0, preinscritsCrm: 0, actifsCharl: 0, sortisCharl: 0, absentsCrm: 0, absentsCharl: 0, nonInscritsCrm: 0 };
  crm.forEach(function (e) { if (e.inscrit) { stats.inscritsCrm++; if (_charlNorm_(e.typo) !== "inscrit") stats.preinscritsCrm++; } });
  var auj = new Date();
  var altR = function (e) { return /altern|contrat|apprenti/i.test(e.statut) || !!e.idEnt; };
  stats.alt = { deux: 0, charlSeul: 0, crmSeul: 0, charlAbsentCrm: 0, charlNonInscrit: 0, crmAbsentCharl: 0, crmSortiCharl: 0 };
  // Doublons CRM : plusieurs fiches inscrites / préinscrites pour la même personne (même Identifiant Elève ou même nom)
  var grp = {}, cleGrp = {}, nGrp = 0;
  crm.forEach(function (e) {
    if (!e.inscrit) return;
    var cles = ["n:" + _charlNorm_(e.nom + " " + e.prenom)].concat(e.idEleve ? ["i:" + e.idEleve] : []);
    var g = null; cles.forEach(function (k) { if (!g && cleGrp[k]) g = cleGrp[k]; });
    if (!g) { g = "G" + (nGrp++); grp[g] = []; }
    cles.forEach(function (k) { cleGrp[k] = cleGrp[k] || g; });
    grp[g].push(e); e.grp = g;
  });
  stats.doublonsCrm = 0;
  Object.keys(grp).forEach(function (g) { if (grp[g].length > 1) stats.doublonsCrm += grp[g].length - 1; });
  stats.inscritsCrmBrut = stats.inscritsCrm;
  stats.inscritsCrm = stats.inscritsCrmBrut - stats.doublonsCrm;   // personnes distinctes
  stats.altCrm = Object.keys(grp).filter(function (g) { return grp[g].some(altR); }).length;
  stats.doublonsCharl = 0;
  var vusCharl = {};
  // Une ligne par étudiant : entreprise de chaque côté
  function ligneComp(verdict, c, e, sim) {
    C.push({ verdict: verdict, idEtu: e ? e.id : "", idEleve: c ? c.idEleve : (e ? e.idEleve : ""),
      nom: e ? (e.nom + " " + e.prenom).trim() : (c.nom + " " + c.prenom).trim(), typo: e ? e.typo : "",
      classeCrm: e ? e.classe : "", classeCharl: c ? (c.classe || c.codeClasse) : "",
      entCharl: c && !_charlSansEntreprise_(c.societe) ? c.societe : "", entCrm: e ? (e.entreprise || "") : "",
      statutCrm: e ? e.statut : "", ressemblance: sim == null ? "" : String(Math.round(sim * 100)) });
  }
  function anomalie(g, type, c, e, valC, valR, detail, sugg) {
    A.push({ gravite: g, type: type, idEtu: e ? e.id : "", idEleve: c ? c.idEleve : (e ? e.idEleve : ""),
      nom: c ? (c.nom + " " + c.prenom).trim() : (e.nom + " " + e.prenom).trim(), classe: c ? (c.classe || c.codeClasse) : (e ? e.classe : ""),
      charlemagne: valC || "", crm: valR || "", detail: detail || "", suggestion: sugg || "" });
  }

  // 1. Rapprochement des étudiants
  lignes.forEach(function (c) {
    var e = null, via = "";
    if (c.idEleve && parId[c.idEleve]) { e = parId[c.idEleve]; via = "id"; }
    if (!e) [c.email, c.mailPerso].some(function (m) { m = String(m || "").toLowerCase(); if (m && parMail[m]) { e = parMail[m]; via = "mail"; return true; } return false; });
    if (!e) {
      var cands = (parNom[_charlNorm_(c.nom + " " + c.prenom)] || []).filter(function (x) { return !x.vu; });
      if (cands.length === 1) { e = cands[0]; via = "nom"; }
      else if (cands.length > 1) {
        var inscrits = cands.filter(function (x) { return x.inscrit; });
        var actifs = inscrits.length ? inscrits : cands.filter(function (x) { return !x.sortie; });
        if (actifs.length === 1) { e = actifs[0]; via = "nom"; }
      }
    }
    var cSortie = _charlDate_(c.sortie);
    var actifC = !cSortie || cSortie > auj;
    var cleC = c.idEleve ? "i:" + c.idEleve : "n:" + _charlNorm_(c.nom + " " + c.prenom) + "|" + String(c.naissance || "");   // après fusion, un identifiant n'apparaît qu'une fois
    var dejaC = vusCharl[cleC] || (e && e.vuPar);
    if (dejaC) {   // même élève deux fois dans l'export, ou deux élèves Charlemagne sur la même fiche CRM
      if (actifC) stats.doublonsCharl++;
      anomalie("À vérifier", "Doublon Charlemagne", c, e, (c.classe || c.codeClasse) + (c.idEleve ? " · n° " + c.idEleve : ""), e ? e.id : "—",
        vusCharl[cleC] ? "Élève présent deux fois dans l'export Charlemagne sans Identifiant Elève (même nom et date de naissance) : compté une seule fois"
          : "Deux élèves Charlemagne sont rapprochés de la même fiche CRM (déjà rapprochée de « " + e.vuPar + " ») : compté une seule fois",
        "Supprimer le doublon dans Charlemagne, ou corriger l'Identifiant Elève de la fiche CRM");
      if (actifC) ligneComp("Doublon Charlemagne", c, e, null);
      return;
    }
    vusCharl[cleC] = 1;
    if (e) e.vuPar = (c.nom + " " + c.prenom).trim() + (c.idEleve ? " n° " + c.idEleve : "");
    if (actifC) stats.actifsCharl++; else stats.sortisCharl++;
    if (e && !e.inscrit) {
      // Fiche CRM trouvée mais ni inscrite ni préinscrite : hors périmètre de comparaison
      e.vu = true;
      if (actifC) {
        stats.nonInscritsCrm++;
        if (!_charlSansEntreprise_(c.societe)) stats.alt.charlNonInscrit++;
        anomalie("Erreur", "Typologie", c, e, (c.classe || c.codeClasse) + " (actif)", e.typo || "(vide)",
          "Actif dans Charlemagne mais « " + (e.typo || "sans typologie") + " » dans le CRM : il n'est pas compté parmi les inscrits", "Passer la fiche en Inscrit / Préinscrit, ou saisir la sortie dans Charlemagne");
        ligneComp("Non inscrit dans le CRM", c, e, null);
      }
      return;
    }
    if (!e) {
      var prospectC = /prosp|prei/i.test(c.etatPreinscription);
      if (actifC) { stats.absentsCrm++; ligneComp("Absent du CRM", c, null, null); if (!_charlSansEntreprise_(c.societe)) stats.alt.charlAbsentCrm++; }
      if (actifC)
        anomalie(prospectC ? "À vérifier" : "Erreur", "Absent du CRM", c, null, (c.classe || c.codeClasse) + (c.societe && !_charlSansEntreprise_(c.societe) ? " · " + c.societe : ""), "—",
          "Étudiant présent dans Charlemagne (n° " + (c.idEleve || "?") + ") introuvable dans le CRM" + (prospectC ? " (prospect / préinscrit dans Charlemagne)" : ""), "Créer la fiche ou renseigner l'Identifiant Elève (col. X) sur la bonne fiche");
      return;
    }
    e.vu = true; stats.rapproches++; stats[via === "id" ? "parId" : via === "mail" ? "parMail" : "parNom"]++;
    paires.push({ c: c, e: e, via: via });
    if (via !== "id" && c.idEleve) anomalie("Info", "Identifiant Charlemagne", c, e, c.idEleve, e.idEleve || "(vide)",
      "Rapproché par " + (via === "mail" ? "e-mail" : "nom et prénom") + " : l'Identifiant Elève du CRM est " + (e.idEleve ? "différent" : "vide"), "Mettre " + c.idEleve + " en colonne X (Identifiant Elève)");

    // 2. Dates d'entrée / de sortie
    var cEntree = _charlDate_(c.entree);
    if (cSortie && !e.sortie && !e.sortieBrute)
      anomalie(cSortie <= new Date() ? "Erreur" : "À vérifier", "Date de sortie", c, e, _charlFmt_(cSortie) + (c.motif ? " (" + c.motif + ")" : ""), "aucune",
        cSortie <= new Date() ? "Sorti dans Charlemagne mais toujours actif dans le CRM (accès à l'espace étudiant encore ouvert)" : "Sortie programmée dans Charlemagne, absente du CRM", "Reporter la date de sortie dans le CRM");
    else if (!cSortie && (e.sortie || e.sortieBrute))
      anomalie("À vérifier", "Date de sortie", c, e, "aucune", e.sortie ? _charlFmt_(e.sortie) : e.sortieBrute, "Sorti dans le CRM mais toujours présent sans date de sortie dans Charlemagne", "");
    else if (cSortie && e.sortie && Math.abs(cSortie - e.sortie) > 86400000)
      anomalie("À vérifier", "Date de sortie", c, e, _charlFmt_(cSortie), _charlFmt_(e.sortie), "Dates de sortie différentes", "");
    else if (cSortie && !e.sortie && e.sortieBrute)
      anomalie("À vérifier", "Date de sortie", c, e, _charlFmt_(cSortie), e.sortieBrute, "Date de sortie du CRM illisible (saisie en texte non reconnue)", "Corriger la saisie");
    if (cEntree && e.entree && Math.abs(cEntree - e.entree) > 86400000)
      anomalie("À vérifier", "Date d'entrée", c, e, _charlFmt_(cEntree), _charlFmt_(e.entree), "Dates d'entrée différentes", "");
    else if (cEntree && !e.entree)
      anomalie("Info", "Date d'entrée", c, e, _charlFmt_(cEntree), "aucune", "Date d'entrée absente du CRM", "");

    // 3. Entreprise / alternance
    var sansC = _charlSansEntreprise_(c.societe);
    var altCrm = /altern|contrat|apprenti/i.test(e.statut) || !!e.idEnt;
    if (sansC) {
      stats.sansEntCharl++;
      if (altCrm && e.entreprise)
        anomalie("Erreur", "Entreprise", c, e, "SANS ENTREPRISE", e.entreprise + (e.statut ? " (" + e.statut + ")" : ""),
          "Sans alternance dans Charlemagne mais rattaché à une entreprise dans le CRM", "Vérifier le contrat : fin de contrat à reporter dans un des deux logiciels");
      else if (/altern|contrat/i.test(e.statut))
        anomalie("À vérifier", "Entreprise", c, e, "SANS ENTREPRISE", e.statut, "Statut CRM en alternance sans entreprise, et sans entreprise dans Charlemagne", "Corriger le statut CRM (En recherche ?)");
    } else {
      stats.alternantsCharl++;
      if (!e.entreprise) {
        var best = _charlMeilleureEnt_(c.societe, listeEnt);
        anomalie("Erreur", "Entreprise", c, e, c.societe, (e.statut || "sans entreprise"),
          "Alternant dans Charlemagne mais aucune entreprise liée dans le CRM",
          best ? "Lier l'entreprise CRM « " + best.nom + " » (" + best.id + ", ressemblance " + Math.round(best.sim * 100) + " %)" : "Entreprise absente du CRM : la créer puis la lier");
      } else {
        var sim = Math.max(_charlSimEnt_(c.societe, e.entreprise), e.nomEnt ? _charlSimEnt_(c.societe, e.nomEnt) : 0);
        if (sim >= 0.999) stats.entOk++;
        else if (sim >= CHARL_SEUIL_ENT) stats.entApprochant++;
        else {
          var best2 = _charlMeilleureEnt_(c.societe, listeEnt);
          anomalie("Erreur", "Entreprise", c, e, c.societe, e.entreprise, "Entreprise différente entre Charlemagne et le CRM (ressemblance " + Math.round(sim * 100) + " %)",
            best2 && best2.id !== e.idEnt ? "Dans le CRM, « " + c.societe + " » correspond plutôt à « " + best2.nom + " » (" + best2.id + ")" : "Vérifier le contrat en cours");
        }
        if (!/altern|contrat|apprenti/i.test(e.statut))
          anomalie("À vérifier", "Statut", c, e, "Alternant (" + c.societe + ")", e.statut || "vide", "Alternant dans Charlemagne mais statut CRM non « En Alternance »", "Passer le statut à « En Alternance »");
      }
    }
    if (actifC) { if (!sansC && altR(e)) stats.alt.deux++; else if (!sansC) stats.alt.charlSeul++; else if (altR(e)) stats.alt.crmSeul++; }
    else if (altR(e)) stats.alt.crmSortiCharl++;
    if (!actifC) { stats.sortisCharlInscritsCrm = (stats.sortisCharlInscritsCrm || 0) + 1; ligneComp("Sorti dans Charlemagne", c, e, null); }
    else if (sansC) ligneComp(e.entreprise ? "CRM seulement" : "Aucune des deux", c, e, null);
    else if (!e.entreprise) ligneComp("Charlemagne seulement", c, e, null);
    else {
      var simC = Math.max(_charlSimEnt_(c.societe, e.entreprise), e.nomEnt ? _charlSimEnt_(c.societe, e.nomEnt) : 0);
      ligneComp(simC >= 0.999 ? "Identique" : simC >= CHARL_SEUIL_ENT ? "Approchant" : "Différente", c, e, simC);
    }
  });

  // 4. Classes : correspondance dominante apprise sur les paires rapprochées
  var corr = {}, table = [];
  paires.forEach(function (x) {
    var kc = String(x.c.classe || x.c.codeClasse || "").trim(), kr = String(x.e.classe || "").trim();
    if (!kc) return;
    corr[kc] = corr[kc] || { total: 0, crm: {} };
    corr[kc].total++; corr[kc].crm[kr || "(vide)"] = (corr[kc].crm[kr || "(vide)"] || 0) + 1;
  });
  Object.keys(corr).forEach(function (kc) {
    var o = corr[kc], dom = Object.keys(o.crm).sort(function (a, b) { return o.crm[b] - o.crm[a]; })[0];
    o.dominante = dom; o.part = o.crm[dom] / o.total;
    table.push({ charlemagne: kc, crm: dom, effectif: o.total, part: Math.round(o.part * 100) });
  });
  paires.forEach(function (x) {
    var kc = String(x.c.classe || x.c.codeClasse || "").trim(), kr = String(x.e.classe || "").trim() || "(vide)";
    if (!kc || !corr[kc]) return;
    var o = corr[kc];
    if (kr === o.dominante) return;
    var memeNom = _charlNorm_(kc).replace(/ /g, "") === _charlNorm_(kr).replace(/ /g, "");
    if (memeNom) return;
    if (_charlDate_(x.c.sortie) && _charlDate_(x.c.sortie) <= new Date()) return;   // sorti : classe sans importance
    anomalie(kr === "(vide)" ? "Erreur" : "À vérifier", "Classe", x.c, x.e, kc, kr,
      kr === "(vide)" ? "Aucune classe dans le CRM (l'étudiant est traité comme PROSPECT : espace étudiant fermé)" : "Classe différente de celle des autres étudiants de « " + kc + " » (" + o.dominante + " pour " + o.crm[o.dominante] + "/" + o.total + ")",
      o.part >= 0.6 ? "Classe attendue dans le CRM : " + o.dominante : "");
  });

  // 5. Étudiants actifs du CRM absents de l'export
  crm.forEach(function (e) {
    if (e.vu || !e.inscrit) return;
    var g = grp[e.grp] || [e];
    if (g.length > 1 && (g.some(function (x) { return x.vu; }) || g[0] !== e)) {   // fiche en double : la personne est déjà comptée
      var autre = g.filter(function (x) { return x !== e; })[0];
      anomalie("À vérifier", "Doublon CRM", null, e, "—", e.id + " = " + autre.id,
        "Fiche CRM en double pour la même personne (" + (e.idEleve && e.idEleve === autre.idEleve ? "même Identifiant Elève" : "même nom") + ") : comptée une seule fois", "Fusionner ou supprimer la fiche en trop");
      ligneComp("Doublon CRM", null, e, null);
      return;
    }
    stats.absentsCharl++;
    if (altR(e)) stats.alt.crmAbsentCharl++;
    ligneComp("Absent de Charlemagne", null, e, null);
    anomalie("Erreur", "Absent de Charlemagne", null, e, "—", (e.typo || "") + " · " + (e.classe || "sans classe") + (e.entreprise ? " · " + e.entreprise : ""),
      "Inscrit / préinscrit dans le CRM mais absent de l'export Charlemagne", "Vérifier son inscription dans Charlemagne, ou passer la fiche CRM en « Sortie »");
  });

  // 6. Effectif : inscrits + préinscrits du CRM (personnes distinctes) = étudiants actifs de Charlemagne (doublons écartés)
  stats.ecartEffectif = stats.actifsCharl - stats.inscritsCrm;
  stats.altCharl = stats.alt.deux + stats.alt.charlSeul + stats.alt.charlAbsentCrm + stats.alt.charlNonInscrit;
  stats.ecartAlt = stats.altCharl - stats.altCrm;
  if (stats.ecartAlt !== 0)
    anomalie("Erreur", "Effectif", null, { id: "", idEleve: "", nom: "ALTERNANTS", prenom: "", classe: "" }, stats.altCharl + " alternants", stats.altCrm + " alternants",
      "Nombre d'alternants différent (écart " + (stats.ecartAlt > 0 ? "+" : "") + stats.ecartAlt + ") : " + stats.alt.charlSeul + " en alternance dans Charlemagne seulement, " + stats.alt.crmSeul + " dans le CRM seulement, "
        + (stats.alt.charlAbsentCrm + stats.alt.charlNonInscrit) + " alternant(s) Charlemagne absents ou non inscrits au CRM, " + (stats.alt.crmAbsentCharl + stats.alt.crmSortiCharl) + " alternant(s) CRM absents ou sortis de Charlemagne",
      "Voir le comparatif : verdicts « Charlemagne seulement » et « CRM seulement »");
  if (stats.ecartEffectif !== 0)
    anomalie("Erreur", "Effectif", null, { id: "", idEleve: "", nom: "EFFECTIF TOTAL", prenom: "", classe: "" }, stats.actifsCharl + " actifs", stats.inscritsCrm + " inscrits / préinscrits",
      "Les effectifs ne correspondent pas (écart " + (stats.ecartEffectif > 0 ? "+" : "") + stats.ecartEffectif + ", doublons écartés : " + stats.doublonsCharl + " Charlemagne, " + stats.doublonsCrm + " CRM) : " + stats.absentsCrm + " absent(s) du CRM, " + stats.nonInscritsCrm + " non inscrit(s) dans le CRM, " + stats.absentsCharl + " absent(s) de Charlemagne, " + (stats.sortisCharlInscritsCrm || 0) + " sorti(s) dans Charlemagne mais encore inscrit(s) dans le CRM",
      "Traiter les lignes « Absent du CRM », « Typologie », « Absent de Charlemagne » et « Date de sortie »");
  var ordreV = {}; CHARL_VERDICTS.forEach(function (v, i) { ordreV[v] = i; });
  C.sort(function (a, b) { return (ordreV[a.verdict] - ordreV[b.verdict]) || a.nom.localeCompare(b.nom); });
  var parVerdict = {}; C.forEach(function (x) { parVerdict[x.verdict] = (parVerdict[x.verdict] || 0) + 1; });

  var ordreG = { "Erreur": 0, "À vérifier": 1, "Info": 2 };
  A.sort(function (a, b) { return (ordreG[a.gravite] - ordreG[b.gravite]) || a.type.localeCompare(b.type) || a.nom.localeCompare(b.nom); });
  var n = { "Erreur": 0, "À vérifier": 0, "Info": 0 }, parType = {};
  A.forEach(function (a) { n[a.gravite]++; parType[a.type] = (parType[a.type] || 0) + 1; });
  stats.erreurs = n["Erreur"]; stats.aVerifier = n["À vérifier"]; stats.infos = n["Info"];
  table.sort(function (a, b) { return b.effectif - a.effectif; });
  return {
    stats: stats, parType: parType, classes: table, anomalies: A, comparatif: C, parVerdict: parVerdict,
    resume: "Effectif : " + stats.inscritsCrm + " inscrits/préinscrits CRM vs " + stats.actifsCharl + " actifs Charlemagne" + (stats.ecartEffectif ? " (écart " + stats.ecartEffectif + ")" : " ✓") + " — "
      + stats.charlemagne + " étudiants Charlemagne (identifiants distincts), " + stats.rapproches + " rapprochés (" + stats.parId + " par identifiant, " + stats.parMail + " par e-mail, "
      + stats.parNom + " par nom) — " + stats.alternantsCharl + " alternants (" + stats.entOk + " entreprise identique, " + stats.entApprochant + " nom approchant accepté), "
      + stats.sansEntCharl + " sans entreprise — " + n["Erreur"] + " erreur(s), " + n["À vérifier"] + " à vérifier, " + n["Info"] + " info(s)"
  };
}
function _charlMeilleureEnt_(nom, listeEnt) {
  var best = null;
  listeEnt.forEach(function (x) { var s = _charlSimEnt_(nom, x.nom); if (s >= CHARL_SEUIL_ENT && (!best || s > best.sim)) best = { id: x.id, nom: x.nom, sim: s }; });
  return best;
}

function _charlEcrireImport_(ss, lignes, R) {
  var ws = ss.getSheetByName(CHARL_IMPORT) || ss.insertSheet(CHARL_IMPORT);
  ws.clear();
  var rows = [["Import Charlemagne — " + R.date + " — " + R.fichier + " — par " + R.importePar].concat(CHARL_CHAMPS.slice(1).map(function () { return ""; })), CHARL_CHAMPS];
  lignes.forEach(function (l) { rows.push(CHARL_CHAMPS.map(function (c) { return l[c]; })); });
  ws.getRange(1, 1, rows.length, CHARL_CHAMPS.length).setNumberFormat("@").setValues(rows);
  ws.getRange(2, 1, 1, CHARL_CHAMPS.length).setFontWeight("bold");
  ws.setFrozenRows(2);
}
function _charlEcrireComparatif_(ss, R) {
  var ws = ss.getSheetByName(CHARL_COMPARATIF) || ss.insertSheet(CHARL_COMPARATIF);
  ws.clear();
  var rows = [CHARL_COMP_TITRES];
  R.comparatif.forEach(function (x) { rows.push(CHARL_COMP_COLS.map(function (c) { return x[c] == null ? "" : x[c]; })); });
  ws.getRange(1, 1, rows.length, CHARL_COMP_COLS.length).setNumberFormat("@").setValues(rows);
  ws.getRange(1, 1, 1, CHARL_COMP_COLS.length).setFontWeight("bold").setBackground("#1A4E8A").setFontColor("#ffffff");
  ws.setFrozenRows(1);
}
function _charlEcrireRapport_(ss, R) {
  var ws = ss.getSheetByName(CHARL_FEUILLE) || ss.insertSheet(CHARL_FEUILLE);
  ws.clear();
  var rows = [["Contrôle Charlemagne ↔ CRM — " + R.date + " — fichier « " + R.fichier + " » — par " + R.importePar, "", "", "", "", "", "", "", "", ""],
    [R.resume, "", "", "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", "", "", ""],
    ["Gravité", "Type", "ID CRM", "Identifiant Charlemagne", "Étudiant", "Classe", "Charlemagne", "CRM", "Détail", "Action suggérée"]];
  R.anomalies.forEach(function (a) { rows.push([a.gravite, a.type, a.idEtu, a.idEleve, a.nom, a.classe, a.charlemagne, a.crm, a.detail, a.suggestion]); });
  ws.getRange(1, 1, rows.length, 10).setNumberFormat("@").setValues(rows);
  ws.getRange(1, 1).setFontWeight("bold").setFontSize(13);
  ws.getRange(4, 1, 1, 10).setFontWeight("bold").setBackground("#1A4E8A").setFontColor("#ffffff");
  ws.setFrozenRows(4);
  var meta = { date: R.date, fichier: R.fichier, importePar: R.importePar, resume: R.resume, stats: R.stats, parType: R.parType, parVerdict: R.parVerdict, classes: R.classes.slice(0, 80) };
  var json = JSON.stringify(meta);
  while (json.length > 8500 && meta.classes.length) { meta.classes = meta.classes.slice(0, Math.floor(meta.classes.length / 2)); json = JSON.stringify(meta); }   // limite 9 Ko d'une propriété
  PropertiesService.getScriptProperties().setProperty("CHARL_DERNIER", json);
}
