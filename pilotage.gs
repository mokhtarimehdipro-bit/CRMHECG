// ════════════════════════════════════════════════════════════════════════════
//  PILOTAGE GOOGLE WORKSPACE (demande direction 2026-11-25) — page pilotage.html
//  Accès : m.mokhtari@hecg.fr et a.mhoudini@hecg.fr uniquement (vérifié ici, sur l'e-mail de la session).
//
//  1) GROUPES PAR CLASSE : pour chaque classe, un groupe « etu-<classe>@hecg.fr » (étudiants inscrits /
//     préinscrits, adresse @hecg.fr de la fiche) et un groupe « form-<classe>@hecg.fr » (formateurs choisis
//     au cas par cas). « Préparer les groupes » calcule toutes les actions (création, ajouts, retraits,
//     suppression des autres groupes) SANS RIEN MODIFIER ; rien n'est exécuté tant que l'utilisateur
//     n'a pas coché puis « autorisé l'import ». Config : feuille PIL_GROUPES ; plan : PIL_PLAN ; journal : PIL_JOURNAL.
//  2) DRIVE — RAPATRIEMENT DE LA GED : les dossiers « NOM Prénom » créés par le CRM (racine GED étudiants)
//     alors que l'étudiant avait déjà un dossier d'origine (Dossiers étudiants / Candidatures) sont vidés
//     dans un sous-dossier « GED CRM NOM Prénom » du dossier d'origine ; le CRM pointe désormais dessus ;
//     un raccourci le remplace dans la GED ; l'ancien dossier vide part à la corbeille (30 jours).
//     Les nouveaux dossiers suivent la même règle (_gedDossierCrm_).
//  Droits : admin.directory.group (à autoriser une fois : fonction autoriserGroupesWorkspace dans l'éditeur).
// ════════════════════════════════════════════════════════════════════════════
var PIL_ADMINS = ["m.mokhtari@hecg.fr", "a.mhoudini@hecg.fr"];
var PIL_DOMAINE = "hecg.fr";
var PIL_CFG_H = ["Classe", "Groupe étudiants", "Groupe formateurs", "Formateurs retenus (JSON)", "Actif", "Mis à jour le", "Par"];
var PIL_PLAN_H = ["ID action", "Type", "Groupe", "Membre", "Libellé", "Coché par défaut", "Statut", "Résultat", "Préparé le"];
var PIL_JOURNAL_H = ["Date", "Par", "Action", "Groupe", "Membre", "Résultat"];
var PIL_RAP_H = ["ID étudiant", "Étudiant", "Dossier CRM (ID)", "Dossier d'origine (ID)", "Chemin d'origine", "Statut", "Détail", "Candidats (JSON)", "Mis à jour le"];

function _pilAutorise_(p) { return PIL_ADMINS.indexOf(String(p.emailAuteur || "").trim().toLowerCase()) !== -1; }
function _pilFeuille_(ss, nom, H) {
  var ws = ss.getSheetByName(nom);
  if (!ws) { ws = ss.insertSheet(nom); ws.getRange(1, 1, 1, H.length).setValues([H]).setFontWeight("bold").setBackground("#1A4E8A").setFontColor("#ffffff"); ws.setFrozenRows(1); }
  return ws;
}
function _pilNorm_(s) { return String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim(); }
function _pilSlug_(classe) { return _pilNorm_(classe).replace(/ /g, "-").replace(/^-+|-+$/g, "").slice(0, 50) || "classe"; }
function _pilMailOk_(m) { return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(m || "")); }
function _pilJournal_(ss, par, lignes) {
  if (!lignes.length) return;
  var ws = _pilFeuille_(ss, "PIL_JOURNAL", PIL_JOURNAL_H);
  ws.getRange(ws.getLastRow() + 1, 1, lignes.length, PIL_JOURNAL_H.length).setValues(lignes.map(function (l) { return [new Date(), par].concat(l); }));
}

// ── Données de référence ─────────────────────────────────────────────────────
function _pilFormateurs_(ss) {
  var ws; try { ws = getSheetSafe(ss, "Adultes"); } catch (e) { ws = getSheetSafe(ss, "FORMATEURS"); }
  var d = ws.getDataRange().getValues(), out = [];
  for (var i = 1; i < d.length; i++) {
    if (!d[i][0]) continue;
    var pro = String(d[i][4] || "").trim().toLowerCase(), perso = String(d[i][3] || "").trim().toLowerCase();
    out.push({ id: String(d[i][0]), nom: (String(d[i][2] || "") + " " + String(d[i][1] || "")).trim(), email: _pilMailOk_(pro) ? pro : (_pilMailOk_(perso) ? perso : ""),
      sorti: !!String(d[i][10] || "").trim(), classes: String(d[i][14] || "").split(",").map(function (c) { return c.trim(); }).filter(Boolean) });
  }
  return out.sort(function (a, b) { return a.nom.localeCompare(b.nom, "fr"); });
}
function _pilConfig_(ss) {
  var ws = _pilFeuille_(ss, "PIL_GROUPES", PIL_CFG_H), d = ws.getDataRange().getValues(), cfg = {};
  for (var i = 1; i < d.length; i++) {
    var c = String(d[i][0] || "").trim(); if (!c) continue;
    var f = []; try { f = JSON.parse(String(d[i][3] || "[]")); } catch (e) {}
    cfg[c] = { classe: c, etu: String(d[i][1] || "").trim().toLowerCase(), form: String(d[i][2] || "").trim().toLowerCase(), formateurs: f, actif: String(d[i][4]) !== "NON", ligne: i + 1 };
  }
  return cfg;
}
// Classes = classes des étudiants inscrits / préinscrits (+ classes déjà configurées)
function _pilClasses_(ss) {
  var etus = _vsuEtudiants_(ss), cfg = _pilConfig_(ss), parClasse = {};
  etus.forEach(function (e) {
    if (!e.inscrit || !e.classe) return;
    (parClasse[e.classe] = parClasse[e.classe] || []).push({ id: e.id, nom: (e.prenom + " " + e.nom).trim(), email: _pilMailOk_(e.mailHecg) && /@hecg\.fr$/.test(e.mailHecg) ? e.mailHecg : "", typo: e.typo });
  });
  var noms = Object.keys(parClasse); Object.keys(cfg).forEach(function (c) { if (noms.indexOf(c) === -1) noms.push(c); });
  noms.sort(function (a, b) { return a.localeCompare(b, "fr"); });
  return noms.map(function (c) {
    var k = cfg[c] || {}, slug = _pilSlug_(c);
    return { classe: c, etudiants: parClasse[c] || [], etu: k.etu || ("etu-" + slug + "@" + PIL_DOMAINE), form: k.form || ("form-" + slug + "@" + PIL_DOMAINE),
      formateurs: k.formateurs || null, actif: k.actif !== false, configure: !!cfg[c] };
  });
}

// ── API Admin SDK (groupes) ──────────────────────────────────────────────────
function _pilGroupes_() {
  var out = [], page = "";
  do {
    var j = _wsFetch_("get", "/groups?customer=my_customer&maxResults=200" + (page ? "&pageToken=" + page : ""));
    (j.groups || []).forEach(function (g) { out.push({ email: String(g.email).toLowerCase(), nom: g.name || "", membres: +g.directMembersCount || 0, description: g.description || "", alias: (g.aliases || []).map(function (a) { return String(a).toLowerCase(); }) }); });
    page = j.nextPageToken || "";
  } while (page);
  return out;
}
function _pilMembres_(groupe) {
  var out = [], page = "";
  do {
    var j = _wsFetch_("get", "/groups/" + encodeURIComponent(groupe) + "/members?maxResults=200" + (page ? "&pageToken=" + page : ""));
    (j.members || []).forEach(function (m) { out.push({ email: String(m.email || "").toLowerCase(), role: m.role || "MEMBER", type: m.type || "USER" }); });
    page = j.nextPageToken || "";
  } while (page);
  return out;
}

// ── Préparation : plan d'actions, rien n'est modifié ─────────────────────────
function _pilPreparer_(ss) {
  var classes = _pilClasses_(ss).filter(function (c) { return c.actif; }), formateurs = _pilFormateurs_(ss), fParId = {};
  formateurs.forEach(function (f) { fParId[f.id] = f; });
  var existants = _pilGroupes_(), parMail = {}; existants.forEach(function (g) { parMail[g.email] = g; g.alias.forEach(function (a) { parMail[a] = g; }); });
  var actions = [], verif = [], prevus = {}, n = 0;
  var act = function (type, groupe, membre, lib, coche) { actions.push({ id: "A" + (++n), type: type, groupe: groupe, membre: membre || "", lib: lib, coche: coche !== false }); };
  classes.forEach(function (c) {
    var voulu = {
      etu: c.etudiants.filter(function (e) { return e.email; }).map(function (e) { return e.email; }),
      form: (c.formateurs === null ? formateurs.filter(function (f) { return !f.sorti && f.classes.indexOf(c.classe) !== -1; }).map(function (f) { return f.id; }) : c.formateurs)
        .map(function (id) { return fParId[id] ? fParId[id].email : ""; }).filter(Boolean)
    };
    c.etudiants.filter(function (e) { return !e.email; }).forEach(function (e) { verif.push({ classe: c.classe, etudiant: e.nom, id: e.id, probleme: "Inscrit / préinscrit sans adresse @hecg.fr dans sa fiche : ne peut pas être ajouté au groupe" }); });
    [["etu", c.etu, "Étudiants " + c.classe], ["form", c.form, "Formateurs " + c.classe]].forEach(function (x) {
      var cle = x[0], mail = x[1]; prevus[mail] = true;
      if (!_pilMailOk_(mail) || mail.split("@")[1] !== PIL_DOMAINE) { verif.push({ classe: c.classe, etudiant: "", id: "", probleme: "Adresse de groupe invalide : " + mail }); return; }
      var g = parMail[mail], membres = [];
      if (!g) act("creer", mail, "", "Créer le groupe « " + x[2] + " » (" + mail + ")");
      else { try { membres = _pilMembres_(g.email).map(function (m) { return m.email; }); } catch (eM) { verif.push({ classe: c.classe, etudiant: "", id: "", probleme: "Membres illisibles pour " + mail + " : " + eM.message }); } }
      if (g && g.nom !== x[2]) act("renommer", mail, x[2], "Renommer « " + g.nom + " » en « " + x[2] + " »");
      voulu[cle].forEach(function (m) { if (membres.indexOf(m) === -1) act("ajouter", mail, m, "Ajouter " + m + " à " + mail); });
      membres.forEach(function (m) {
        if (voulu[cle].indexOf(m) !== -1) return;
        // retrait coché par défaut seulement pour un étudiant qui n'est plus inscrit dans cette classe
        act("retirer", mail, m, "Retirer " + m + " de " + mail + (cle === "etu" ? " (n'est pas inscrit / préinscrit dans cette classe)" : " (formateur non retenu pour la classe)"), cle === "etu" && /@hecg\.fr$/.test(m));
      });
    });
  });
  // Groupes existants hors plan : suppression proposée, NON cochée (décision au cas par cas)
  existants.forEach(function (g) {
    if (prevus[g.email] || g.alias.some(function (a) { return prevus[a]; })) return;
    act("supprimer", g.email, "", "Supprimer le groupe « " + g.nom + " » (" + g.email + ", " + g.membres + " membre(s))" + (g.description ? " — " + g.description : ""), false);
  });
  var ws = _pilFeuille_(ss, "PIL_PLAN", PIL_PLAN_H), now = new Date();
  if (ws.getLastRow() > 1) ws.getRange(2, 1, ws.getLastRow() - 1, PIL_PLAN_H.length).clearContent();
  if (actions.length) ws.getRange(2, 1, actions.length, PIL_PLAN_H.length).setValues(actions.map(function (a) { return [a.id, a.type, a.groupe, a.membre, a.lib, a.coche ? "OUI" : "NON", "À faire", "", now]; }));
  return { actions: actions, verif: verif, groupes: existants.length, classes: classes.length };
}
function _pilLirePlan_(ss) {
  var ws = _pilFeuille_(ss, "PIL_PLAN", PIL_PLAN_H), d = ws.getDataRange().getValues(), out = [];
  for (var i = 1; i < d.length; i++) if (d[i][0]) out.push({ id: String(d[i][0]), type: String(d[i][1]), groupe: String(d[i][2]), membre: String(d[i][3]), lib: String(d[i][4]), coche: d[i][5] === "OUI", statut: String(d[i][6]), resultat: String(d[i][7]), ligne: i + 1 });
  return out;
}
function _pilAppliquer_(ss, ids, confirmationSuppr, par) {
  var plan = _pilLirePlan_(ss), ws = _pilFeuille_(ss, "PIL_PLAN", PIL_PLAN_H), fin = Date.now() + 240000, res = [], jr = [];
  var aFaire = plan.filter(function (a) { return ids.indexOf(a.id) !== -1 && a.statut !== "Fait"; });
  // ordre sûr : créations, renommages, ajouts, retraits, puis suppressions de groupes
  var ordre = { creer: 0, renommer: 1, ajouter: 2, retirer: 3, supprimer: 4 };
  aFaire.sort(function (a, b) { return ordre[a.type] - ordre[b.type]; });
  if (aFaire.some(function (a) { return a.type === "supprimer"; }) && confirmationSuppr !== "SUPPRIMER") throw new Error("Confirmation « SUPPRIMER » manquante pour la suppression de groupes.");
  var restant = 0;
  aFaire.forEach(function (a) {
    if (Date.now() > fin) { restant++; return; }
    var ok = true, txt = "";
    try {
      if (a.type === "creer") { var nomG = a.lib.match(/« (.+) »/); _wsFetch_("post", "/groups", { email: a.groupe, name: nomG ? nomG[1] : a.groupe, description: "Groupe de classe géré par le CRM HECG" }); txt = "groupe créé"; }
      else if (a.type === "renommer") { _wsFetch_("patch", "/groups/" + encodeURIComponent(a.groupe), { name: a.membre }); txt = "renommé"; }
      else if (a.type === "ajouter") {
        try { _wsFetch_("post", "/groups/" + encodeURIComponent(a.groupe) + "/members", { email: a.membre, role: "MEMBER" }); txt = "ajouté"; }
        catch (eA) { if (/409|already exists|duplicate/i.test(eA.message)) txt = "déjà membre"; else throw eA; }
      }
      else if (a.type === "retirer") { _wsFetch_("delete", "/groups/" + encodeURIComponent(a.groupe) + "/members/" + encodeURIComponent(a.membre)); txt = "retiré"; }
      else if (a.type === "supprimer") { _wsFetch_("delete", "/groups/" + encodeURIComponent(a.groupe)); txt = "groupe supprimé"; }
    } catch (e) { ok = false; txt = String(e.message || e).slice(0, 250); }
    ws.getRange(a.ligne, 7, 1, 2).setValues([[ok ? "Fait" : "Échec", txt]]);
    res.push({ id: a.id, ok: ok, resultat: txt });
    jr.push([a.type, a.groupe, a.membre, (ok ? "OK — " : "ÉCHEC — ") + txt]);
  });
  _pilJournal_(ss, par, jr);
  return { resultats: res, restant: restant };
}

// ── Drive : index des dossiers d'origine + rapatriement ──────────────────────
function _pilDriveDossiers_() {
  var out = [], page = "", tok = ScriptApp.getOAuthToken();
  do {
    var url = "https://www.googleapis.com/drive/v3/files?pageSize=1000&supportsAllDrives=true&includeItemsFromAllDrives=true&corpora=allDrives" +
      "&q=" + encodeURIComponent("mimeType='application/vnd.google-apps.folder' and trashed=false") +
      "&fields=" + encodeURIComponent("nextPageToken,files(id,name,parents,createdTime)") + (page ? "&pageToken=" + encodeURIComponent(page) : "");
    var r = UrlFetchApp.fetch(url, { headers: { Authorization: "Bearer " + tok }, muteHttpExceptions: true });
    if (r.getResponseCode() >= 300) throw new Error("Drive (" + r.getResponseCode() + ") : " + r.getContentText().slice(0, 200));
    var j = JSON.parse(r.getContentText());
    (j.files || []).forEach(function (f) { out.push({ id: f.id, nom: f.name, parents: f.parents || [], cree: f.createdTime || "" }); });
    page = j.nextPageToken || "";
  } while (page);
  return out;
}
// Dossiers d'origine possibles d'un étudiant (nom exact « NOM Prénom » / « Prénom NOM », éventuellement suivi d'un complément)
function _pilCorrespond_(nomDossier, nom, prenom) {
  var fn = _pilNorm_(nomDossier), a = _pilNorm_(nom + " " + prenom), b = _pilNorm_(prenom + " " + nom);
  if (!fn || !_pilNorm_(nom) || !_pilNorm_(prenom) || /^ged crm /.test(fn)) return false;
  return fn === a || fn === b || fn.indexOf(a + " ") === 0 || fn.indexOf(b + " ") === 0;
}
function _pilDriveApercu_(ss) {
  var tous = _pilDriveDossiers_(), parId = {}, enfants = {};
  tous.forEach(function (f) { parId[f.id] = f; f.parents.forEach(function (p) { (enfants[p] = enfants[p] || []).push(f); }); });
  // descendants des racines d'origine (Dossiers étudiants + Candidatures)
  var origines = [], pile = [{ id: GED_SCAN_ROOT_ID, chemin: "Dossiers étudiants" }, { id: GED_CANDIDATURES_ID, chemin: "Candidatures" }], vus = {};
  while (pile.length) {
    var x = pile.pop(); if (vus[x.id]) continue; vus[x.id] = 1;
    (enfants[x.id] || []).forEach(function (f) { var ch = x.chemin + " / " + f.nom; origines.push({ id: f.id, nom: f.nom, cree: f.cree, chemin: ch }); pile.push({ id: f.id, chemin: ch }); });
  }
  // index des origines conservé pour les créations futures (_gedDossierCrm_)
  var wsI = _pilFeuille_(ss, "GED_ORIGINES", ["ID dossier", "Nom", "Nom normalisé", "Créé le", "Chemin"]);
  if (wsI.getLastRow() > 1) wsI.getRange(2, 1, wsI.getLastRow() - 1, 5).clearContent();
  if (origines.length) wsI.getRange(2, 1, origines.length, 5).setValues(origines.map(function (o) { return [o.id, o.nom, _pilNorm_(o.nom), o.cree, o.chemin]; }));

  var etus = {}; getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues().slice(1).forEach(function (r) { if (r[0]) etus[String(r[0]).trim()] = { nom: String(r[2] || "").trim(), prenom: String(r[3] || "").trim() }; });
  var ged = getOrCreateGedEtuSheet(ss).getDataRange().getValues(), lignes = [];
  for (var i = 1; i < ged.length; i++) {
    var id = String(ged[i][0] || "").trim(), fid = String(ged[i][2] || "").trim(), e = etus[id];
    if (!id || !fid || !e) continue;
    var f = parId[fid];
    if (!f || f.parents.indexOf(GED_ETU_ROOT_ID) === -1) continue;   // dossier déjà rangé ailleurs (origine ou sous-dossier GED CRM) : rien à faire
    var cands = origines.filter(function (o) { return o.id !== fid && _pilCorrespond_(o.nom, e.nom, e.prenom) && (!f.cree || !o.cree || o.cree < f.cree); });
    var statut = !cands.length ? "Pas de dossier d'origine" : cands.length === 1 ? "Prêt" : "À choisir";
    lignes.push([id, (e.nom.toUpperCase() + " " + e.prenom).trim(), fid, cands.length === 1 ? cands[0].id : "", cands.length === 1 ? cands[0].chemin : "", statut,
      !cands.length ? "Dossier créé par le CRM, sans dossier antérieur trouvé : il reste dans la GED" : cands.length === 1 ? "Les fichiers iront dans « " + cands[0].chemin + " / GED CRM " + e.nom.toUpperCase() + " " + e.prenom + " »" : cands.length + " dossiers d'origine possibles : choisissez",
      JSON.stringify(cands.slice(0, 8).map(function (o) { return { id: o.id, chemin: o.chemin }; })), new Date()]);
  }
  var ws = _pilFeuille_(ss, "GED_RAPATRIEMENT", PIL_RAP_H), anc = {};
  ws.getDataRange().getValues().slice(1).forEach(function (r) { if (r[0] && r[5] === "Rapatrié") anc[String(r[0])] = r; });
  Object.keys(anc).forEach(function (k) { if (!lignes.some(function (l) { return l[0] === k; })) lignes.push(anc[k]); });   // historique des rapatriements conservé
  if (ws.getLastRow() > 1) ws.getRange(2, 1, ws.getLastRow() - 1, PIL_RAP_H.length).clearContent();
  if (lignes.length) ws.getRange(2, 1, lignes.length, PIL_RAP_H.length).setValues(lignes);
  return { lignes: _pilLireRap_(ss), origines: origines.length, dossiers: tous.length };
}
function _pilLireRap_(ss) {
  var d = _pilFeuille_(ss, "GED_RAPATRIEMENT", PIL_RAP_H).getDataRange().getValues(), out = [];
  for (var i = 1; i < d.length; i++) if (d[i][0]) {
    var c = []; try { c = JSON.parse(String(d[i][7] || "[]")); } catch (e) {}
    out.push({ id: String(d[i][0]), nom: String(d[i][1]), crm: String(d[i][2]), origine: String(d[i][3]), chemin: String(d[i][4]), statut: String(d[i][5]), detail: String(d[i][6]), candidats: c, ligne: i + 1 });
  }
  return out;
}
// Déplace le contenu d'un dossier CRM dans « GED CRM NOM Prénom » du dossier d'origine
function _pilRapatrier_(ss, r, origineId) {
  var crm = DriveApp.getFolderById(r.crm), orig = DriveApp.getFolderById(origineId);
  var nomSous = "GED CRM " + r.nom, it = orig.getFoldersByName(nomSous), sous = it.hasNext() ? it.next() : orig.createFolder(nomSous);
  var nF = 0, nD = 0, erreurs = [];
  var fi = crm.getFiles(); while (fi.hasNext()) { var f = fi.next(); try { f.moveTo(sous); nF++; } catch (e) { erreurs.push(f.getName() + " : " + e.message); } }
  var di = crm.getFolders(); while (di.hasNext()) { var d = di.next(); try { d.moveTo(sous); nD++; } catch (e) { erreurs.push(d.getName() + " : " + e.message); } }
  // le CRM pointe désormais sur le sous-dossier
  var wsG = getOrCreateGedEtuSheet(ss), g = wsG.getDataRange().getValues();
  for (var i = 1; i < g.length; i++) if (String(g[i][0]).trim() === r.id) { wsG.getRange(i + 1, 3, 1, 2).setValues([[sous.getId(), sous.getUrl()]]); break; }
  // raccourci dans la GED + ancien dossier (vide) à la corbeille
  var racc = "";
  try { var s = DriveApp.getFolderById(GED_ETU_ROOT_ID).createShortcut(sous.getId()); s.setName(r.nom); racc = " · raccourci créé dans la GED"; } catch (eS) { racc = " · raccourci non créé (" + eS.message + ")"; }
  var vide = !crm.getFiles().hasNext() && !crm.getFolders().hasNext();
  if (vide) crm.setTrashed(true);
  return { ok: !erreurs.length, sous: sous.getId(), detail: nF + " fichier(s) et " + nD + " dossier(s) déplacés dans « " + nomSous + " »" + racc + (vide ? " · ancien dossier mis à la corbeille" : " · ancien dossier conservé (non vide)") + (erreurs.length ? " · ERREURS : " + erreurs.join(" ; ") : "") };
}
// Dossier GED d'un étudiant pour les nouveaux dépôts : dossier lié, sinon sous-dossier « GED CRM » du dossier d'origine, sinon dossier dans la GED
function _gedDossierCrm_(ss, id, nom, prenom) {
  var ws = getOrCreateGedEtuSheet(ss), d = ws.getDataRange().getValues(), ligne = -1;
  for (var i = 1; i < d.length; i++) if (String(d[i][0]).trim() === String(id)) {
    ligne = i + 1;
    if (String(d[i][2] || "").trim()) { try { return DriveApp.getFolderById(String(d[i][2]).trim()); } catch (e) {} }
    break;
  }
  var lib = (String(nom || "").toUpperCase() + " " + String(prenom || "")).trim() || String(id), f = null;
  try {
    var wo = ss.getSheetByName("GED_ORIGINES");
    if (wo && wo.getLastRow() > 1) {
      var c = wo.getRange(2, 1, wo.getLastRow() - 1, 3).getValues().filter(function (o) { return _pilCorrespond_(o[1], nom, prenom); });
      if (c.length === 1) {
        var orig = DriveApp.getFolderById(String(c[0][0])), it = orig.getFoldersByName("GED CRM " + lib);
        f = it.hasNext() ? it.next() : orig.createFolder("GED CRM " + lib);
        try { DriveApp.getFolderById(GED_ETU_ROOT_ID).createShortcut(f.getId()).setName(lib); } catch (eS) {}
      }
    }
  } catch (eO) { f = null; }
  if (!f) {
    var root = DriveApp.getFolderById(GED_ETU_ROOT_ID), libs = [lib, prenom + " " + nom, nom + " " + prenom];
    for (var k = 0; k < libs.length && !f; k++) { var it2 = root.getFoldersByName(String(libs[k]).trim()); if (it2.hasNext()) f = it2.next(); }
    if (!f) f = root.createFolder(lib);
  }
  if (ligne > 0) ws.getRange(ligne, 3, 1, 2).setValues([[f.getId(), f.getUrl()]]);
  else ws.appendRow([id, (prenom + " " + nom).trim(), f.getId(), f.getUrl()]);
  return f;
}

// ── Point d'entrée ───────────────────────────────────────────────────────────
function handlePilotage(p, ss) {
  if (!_pilAutorise_(p)) return createJsonResponse({ success: false, message: "Accès réservé à m.mokhtari@hecg.fr et a.mhoudini@hecg.fr." });
  var par = String(p.nomAuteur || p.emailAuteur || ""), op = String(p.op || "");
  try {
    if (op === "etat") {
      var droits = true, msgDroits = "";
      try { _wsFetch_("get", "/groups?customer=my_customer&maxResults=1"); } catch (eD) { droits = false; msgDroits = eD.message; }
      return createJsonResponse({ success: true, classes: _pilClasses_(ss), formateurs: _pilFormateurs_(ss), plan: _pilLirePlan_(ss), droitsGroupes: droits, msgDroits: msgDroits, rapatriement: _pilLireRap_(ss) });
    }
    if (op === "sauverClasse") {
      var c = JSON.parse(String(p.classe || "{}")), classe = String(c.classe || "").trim();
      if (!classe) return createJsonResponse({ success: false, message: "Classe manquante." });
      [c.etu, c.form].forEach(function (m) { if (!_pilMailOk_(m) || String(m).split("@")[1] !== PIL_DOMAINE) throw new Error("Adresse de groupe invalide : " + m + " (doit finir par @" + PIL_DOMAINE + ")"); });
      var ws = _pilFeuille_(ss, "PIL_GROUPES", PIL_CFG_H), cfg = _pilConfig_(ss)[classe];
      var ligne = [classe, String(c.etu).toLowerCase(), String(c.form).toLowerCase(), JSON.stringify((c.formateurs || []).map(String)), c.actif === false ? "NON" : "OUI", new Date(), par];
      if (cfg) ws.getRange(cfg.ligne, 1, 1, PIL_CFG_H.length).setValues([ligne]); else ws.appendRow(ligne);
      logAction(par, "", "Modification", "Pilotage Workspace — groupes", classe, "Paramétrage : étudiants → " + ligne[1] + ", formateurs → " + ligne[2] + " (" + (c.formateurs || []).length + " formateur(s) retenu(s))" + (c.actif === false ? ", classe exclue" : ""));
      return createJsonResponse({ success: true });
    }
    if (op === "preparer") {
      var P = _pilPreparer_(ss);
      logAction(par, "", "Action", "Pilotage Workspace — groupes", "Préparation", P.classes + " classe(s), " + P.groupes + " groupe(s) existant(s) : " + P.actions.length + " action(s) proposée(s), " + P.verif.length + " point(s) à vérifier (rien n'a été modifié)");
      return createJsonResponse({ success: true, plan: _pilLirePlan_(ss), verif: P.verif });
    }
    if (op === "appliquer") {
      var ids = (p.ids || []).map(String);
      if (!ids.length) return createJsonResponse({ success: false, message: "Aucune action cochée." });
      var R = _pilAppliquer_(ss, ids, String(p.confirmation || ""), par);
      var nOk = R.resultats.filter(function (x) { return x.ok; }).length;
      logAction(par, "", "Modification", "Pilotage Workspace — groupes", "Import autorisé", nOk + " action(s) réussie(s), " + (R.resultats.length - nOk) + " échec(s)" + (R.restant ? ", " + R.restant + " restante(s) (relancer)" : "") +
        " — " + R.resultats.slice(0, 30).map(function (x) { return x.id + " " + (x.ok ? "OK" : "ÉCHEC") + " (" + x.resultat + ")"; }).join(" ; "));
      return createJsonResponse({ success: true, resultats: R.resultats, restant: R.restant, plan: _pilLirePlan_(ss) });
    }
    if (op === "driveApercu") {
      var A = _pilDriveApercu_(ss);
      logAction(par, "", "Consultation", "Pilotage Workspace — Drive GED", "Rapatriement", A.dossiers + " dossiers lus, " + A.origines + " dossiers d'origine indexés, " + A.lignes.filter(function (l) { return l.statut === "Prêt"; }).length + " dossier(s) CRM prêt(s) à rapatrier");
      return createJsonResponse({ success: true, rapatriement: A.lignes, origines: A.origines });
    }
    if (op === "driveAppliquer") {
      var choix = p.choix || {}, lignes = _pilLireRap_(ss), wsR = _pilFeuille_(ss, "GED_RAPATRIEMENT", PIL_RAP_H), fin = Date.now() + 240000, res = [], restant = 0;
      lignes.forEach(function (r) {
        var orig = choix[r.id]; if (!orig || r.statut === "Rapatrié") return;
        if (r.origine && orig !== r.origine && !r.candidats.some(function (c) { return c.id === orig; })) { res.push({ id: r.id, ok: false, detail: "dossier d'origine non proposé" }); return; }
        if (!r.origine && !r.candidats.some(function (c) { return c.id === orig; })) { res.push({ id: r.id, ok: false, detail: "dossier d'origine non proposé" }); return; }
        if (Date.now() > fin) { restant++; return; }
        try {
          var x = _pilRapatrier_(ss, r, orig), ch = (r.candidats.filter(function (c) { return c.id === orig; })[0] || {}).chemin || r.chemin;
          wsR.getRange(r.ligne, 4, 1, 6).setValues([[orig, ch, x.ok ? "Rapatrié" : "Rapatrié (avec erreurs)", x.detail, JSON.stringify(r.candidats), new Date()]]);
          res.push({ id: r.id, ok: x.ok, detail: x.detail });
          try { _invalidateStudentDetailsCache_(r.id); } catch (eI) {}
        } catch (e) { res.push({ id: r.id, ok: false, detail: String(e.message || e) }); wsR.getRange(r.ligne, 6, 1, 2).setValues([["Échec", String(e.message || e).slice(0, 300)]]); }
      });
      logAction(par, "", "Modification", "Pilotage Workspace — Drive GED", "Rapatriement", res.filter(function (x) { return x.ok; }).length + " dossier(s) rapatrié(s), " + res.filter(function (x) { return !x.ok; }).length + " échec(s)" + (restant ? ", " + restant + " restant(s)" : "") +
        " — " + res.slice(0, 25).map(function (x) { return x.id + " : " + x.detail; }).join(" ; "));
      return createJsonResponse({ success: true, resultats: res, restant: restant, rapatriement: _pilLireRap_(ss) });
    }
    if (op === "contactsAudit" || op === "contactsCorriger") {
      var app = op === "contactsCorriger", RC = _pilContactsAudit_(ss, String(p.compte || ""), app);
      var tot = function (k) { return RC.comptes.reduce(function (a, c) { return a + (c[k] || 0); }, 0); };
      logAction(par, "", app ? "Modification" : "Consultation", "Pilotage Workspace — contacts Google", app ? "Correction" : "Vérification",
        (p.compte ? p.compte : "tous les comptes") + " : " + tot("nbDoublons") + " fiche(s) en double, " + tot("nbTelFix") + " numéro(s) à corriger, " + tot("nbOrphelins") + " n° CRM sur d'autres fiches"
        + (app ? " → " + RC.comptes.map(function (c) { return c.bilan ? c.compte + " : " + c.bilan.supprimes + " supprimé(s), " + c.bilan.telCorriges + " corrigé(s), " + c.bilan.telRetires + " n° retiré(s)" + (c.bilan.erreurs.length ? " (erreurs : " + c.bilan.erreurs.join(" ; ") + ")" : "") : c.compte + " : " + (c.erreur || "?"); }).join(" | ") : "")
        + " — CRM : " + RC.nbFusions + " personne(s) en double, " + RC.nbConflits + " numéro(s) partagé(s)");
      RC.success = true;
      return createJsonResponse(RC);
    }
    if (op === "sauvegardes") return createJsonResponse({ success: true, sauvegardes: sc2ListeSauvegardes_(), suspendus: sc2Suspensions_() });
    if (op === "restaurer") {
      var B = sc2RestaurerSauvegarde_(String(p.id || ""), par);
      logAction(par, "", "Modification", "Pilotage Workspace — contacts Google", "Restauration", "Sauvegarde " + p.id + " (" + (p.libelle || "") + ") restaurée : " + B.restaurees + " fiche(s) remise(s) en état, " + B.recreees + " recréée(s), " + B.retirees + " retirée(s)" + (B.erreurs.length ? " — erreurs : " + B.erreurs.join(" ; ") : "") + ". Synchro de nuit suspendue pour ce compte.");
      return createJsonResponse({ success: true, bilan: B });
    }
    if (op === "reprendreSynchro") {
      sc2Reprendre_(String(p.compte || ""));
      logAction(par, "", "Modification", "Pilotage Workspace — contacts Google", "Synchro réactivée", String(p.compte || ""));
      return createJsonResponse({ success: true });
    }
    if (op === "annuaireTel") {
      var L = _pilAnnuaireTel_(ss);
      logAction(par, "", "Consultation", "Pilotage Workspace — annuaire téléphones", "Contrôle", L.length + " compte(s) étudiant(s) avec un numéro en trop ou incorrect dans l'annuaire (" + L.filter(function (x) { return x.maPresent; }).length + " portant le n° du MA)");
      return createJsonResponse({ success: true, ecarts: L });
    }
    if (op === "annuaireCorriger") {
      var R2 = _pilAnnuaireCorriger_(ss, (p.emails || []).map(String), par);
      logAction(par, "", "Modification", "Pilotage Workspace — annuaire téléphones", "Correction", R2.filter(function (x) { return x.ok; }).length + " profil(s) corrigé(s) (un seul numéro : celui de l'étudiant) — " + R2.slice(0, 25).map(function (x) { return x.email + " : " + x.detail; }).join(" ; "));
      return createJsonResponse({ success: true, resultats: R2 });
    }
    return createJsonResponse({ success: false, message: "Opération inconnue : " + op });
  } catch (e) { return createJsonResponse({ success: false, message: String(e.message || e) }); }
}
// À lancer UNE fois dans l'éditeur Apps Script (compte administrateur) : autorise la gestion des groupes
function autoriserGroupesWorkspace() { ScriptApp.requireScopes(ScriptApp.AuthMode.FULL, ["https://www.googleapis.com/auth/admin.directory.group"]); Logger.log(_pilGroupes_().length + " groupe(s) lu(s) dans Google Workspace."); }

// ── Annuaire Workspace : un seul numéro (celui de l'étudiant) sur le profil @hecg.fr ──
// Le « Profil de l'annuaire » d'un compte étudiant peut porter un n° « Work » (MA / ancien MA) : il s'affiche
// sur les fiches Contacts / WhatsApp et télescope les identités (signalé 2026-11-25).
function _pilChiffres_(t) { var f = sc2FormatPhone(t); return String(f || "").replace(/\D/g, ""); }
function _pilAnnuaireTel_(ss) {
  var d = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues(), parMail = {}, telMa = {};
  try { getSheetSafe(ss, "Partenariat").getDataRange().getValues().slice(1).forEach(function (r) { if (r[0]) telMa[String(r[0]).trim()] = { tel: _pilChiffres_(r[5]) || _pilChiffres_(r[9]), nom: (String(r[4] || "") + " " + String(r[3] || "")).trim() }; }); } catch (e) {}
  for (var i = 1; i < d.length; i++) {
    var m = String(d[i][26] || "").trim().toLowerCase(); if (!m) continue;
    var ma = telMa[String(d[i][12] || "").trim()] || {}, t = _pilChiffres_(d[i][5]);
    parMail[m] = { id: String(d[i][0]), nom: (String(d[i][3] || "") + " " + String(d[i][2] || "")).trim(), tel: t && t !== ma.tel ? sc2FormatPhone(d[i][5]) : "", telMa: ma.tel || "", ma: ma.nom || "" };
  }
  var out = [], page = "";
  do {
    var j = _wsFetch_("get", "/users?customer=my_customer&maxResults=500&projection=full" + (page ? "&pageToken=" + page : ""));
    (j.users || []).forEach(function (u) {
      var em = String(u.primaryEmail || "").toLowerCase(), e = parMail[em]; if (!e) return;
      var ph = (u.phones || []).map(function (p) { return { valeur: String(p.value || ""), type: p.type === "custom" ? (p.customType || "custom") : (p.type || ""), chiffres: _pilChiffres_(p.value) }; });
      var voulu = e.tel ? _pilChiffres_(e.tel) : "";
      var ok = voulu ? (ph.length === 1 && ph[0].chiffres === voulu) : !ph.length;
      if (ok) return;
      out.push({ brut: u.phones || [], email: em, id: e.id, nom: e.nom, actuels: ph.map(function (p) { return p.valeur + " (" + p.type + ")" + (p.chiffres && p.chiffres === e.telMa ? " ← n° du MA " + e.ma : ""); }),
        cible: e.tel || "", maPresent: ph.some(function (p) { return p.chiffres && p.chiffres === e.telMa; }) });
    });
    page = j.nextPageToken || "";
  } while (page);
  return out;
}
function _pilAnnuaireCorriger_(ss, emails, par) {
  var ecarts = _pilAnnuaireTel_(ss), parMail = {}, res = [], jr = [];
  ecarts.forEach(function (x) { parMail[x.email] = x; });
  // Sauvegarde des numéros actuels AVANT correction (restaurable depuis Pilotage Workspace › Contacts Google)
  var aSauver = emails.slice(0, 60).map(function (em) { return parMail[String(em).toLowerCase()]; }).filter(Boolean).map(function (x) { return { email: x.email, phones: x.brut }; });
  if (aSauver.length) sc2SauvegarderAnnuaire_(aSauver);
  emails.slice(0, 60).forEach(function (em) {
    var x = parMail[String(em).toLowerCase()];
    if (!x) { res.push({ email: em, ok: false, detail: "déjà conforme ou compte non relié à une fiche étudiant" }); return; }
    try {
      _wsFetch_("put", "/users/" + encodeURIComponent(x.email), { phones: x.cible ? [{ value: x.cible, type: "mobile", primary: true }] : [] });
      res.push({ email: x.email, ok: true, detail: x.cible ? "seul numéro conservé : " + x.cible : "numéros retirés (aucun n° personnel dans le CRM)" });
      jr.push(["annuaire-telephone", x.email, x.nom, "OK — avant : " + x.actuels.join(", ") + " → après : " + (x.cible || "aucun")]);
    } catch (e) { res.push({ email: x.email, ok: false, detail: String(e.message || e).slice(0, 200) }); jr.push(["annuaire-telephone", x.email, x.nom, "ÉCHEC — " + String(e.message || e).slice(0, 200)]); }
  });
  _pilJournal_(ss, par, jr);
  return res;
}
