// =============================================================================
//  SIRET DES ENTREPRISES (demande direction 2026-11-12)
//  • fiche partenaire : saisie / recherche du SIRET (action saveSiretEntreprise) ;
//  • complétion en lot (action completerSirets, administrateurs) : la PAGE interroge le
//    répertoire SIRENE (refusé aux serveurs Google, voir adresses.gs) et transmet les résultats ;
//    le serveur choisit l'établissement à partir du NOM et de l'ADRESSE de l'entreprise, puis
//    écrit uniquement la colonne « SIRET » de PARTENARIAT (l'adresse n'est jamais modifiée ici).
//  Règle de choix : nom concordant (_adrNomCorrespond_) ET (rue concordante, ou seul établissement
//  actif dans le code postal / la commune). Sinon : « à vérifier », rien n'est écrit.
// =============================================================================

var SIRET_RAPPORT_FEUILLE = "SIRET_RAPPORT";

function _siretNettoyer_(s) { return String(s == null ? "" : s).replace(/\D/g, ""); }
// Contrôle de Luhn (exception La Poste : somme des chiffres multiple de 5)
function _siretValide_(s) {
  s = _siretNettoyer_(s);
  if (!/^\d{14}$/.test(s)) return false;
  var i, t = 0;
  if (s.indexOf("356000000") === 0) { for (i = 0; i < 14; i++) t += +s[i]; return t % 5 === 0; }
  for (i = 0; i < 14; i++) { var d = +s[13 - i]; if (i % 2) { d *= 2; if (d > 9) d -= 9; } t += d; }
  return t % 10 === 0;
}

// Écrit le SIRET sur toutes les lignes (contacts) des entreprises données : map { ID_ENT (majuscules) : siret }
function _siretEcrire_(ss, map) {
  var ws = getSheetSafe(ss, "PARTENARIAT");
  var C = _adrCols_(ws);
  var n = ws.getLastRow();
  if (n < 2) return 0;
  var ids = ws.getRange(2, 2, n - 1, 1).getValues();
  var rg = ws.getRange(2, C.siret + 1, n - 1, 1);
  var col = rg.getValues(), modif = 0;
  for (var i = 0; i < ids.length; i++) {
    var id = String(ids[i][0] || "").trim().toUpperCase();
    if (!id || !Object.prototype.hasOwnProperty.call(map, id)) continue;
    var v = map[id] ? "'" + map[id] : "";
    if (String(col[i][0]).replace(/\D/g, "") !== String(map[id] || "")) modif++;
    col[i][0] = v;
  }
  rg.setValues(col);
  return modif;
}

/** Fiche partenaire : enregistre (ou efface) le SIRET d'une entreprise. */
function handleSaveSiretEntreprise(p, ss) {
  try {
    var id = String(p.id_entreprise || "").trim().toUpperCase();
    if (!id) return createJsonResponse({ success: false, message: "Entreprise manquante." });
    var siret = _siretNettoyer_(p.siret);
    if (siret && !_siretValide_(siret)) return createJsonResponse({ success: false, message: "SIRET invalide : il doit comporter 14 chiffres (clé de contrôle incorrecte)." });
    var m = {}; m[id] = siret;
    var lock = LockService.getScriptLock(); lock.waitLock(20000);
    try { _siretEcrire_(ss, m); } finally { lock.releaseLock(); }
    logAction(p.nomAuteur || p.idAuteur || "", p.roleAuteur || "", "Modification", "Entreprises — SIRET", id,
      siret ? "SIRET " + siret + " enregistré" + (p.source ? " (" + p.source + ")" : "") : "SIRET effacé");
    try { _invalidatePartnersCache_(); _invalidatePartnerDetailsCache_(id); } catch (eC) {}
    return createJsonResponse({ success: true, siret: siret });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

/** Choisit l'établissement SIRENE correspondant à une entreprise du CRM. */
function _siretChoisir_(ent, resultats) {
  var cands = [];
  (resultats || []).forEach(function (r) {
    if (!r || !_adrNomCorrespond_(ent.nom, r)) return;
    _adrEtablissements_(r).forEach(function (e) {
      var rue = !!(ent.adresse && _adrRueCompatible_(e.voie.replace(/^[\d\s\-,]+(bis|ter)?\s*/i, ""), ent.adresse));
      var cp = !!(ent.cp && ent.cp === e.cp);
      var ville = !!(ent.ville && _adrVilleCompatible_(ent.ville, e.ville));
      cands.push({ siret: e.siret, siren: r.siren, nom: r.nom_complet || r.nom_raison_sociale || "", adresse: e.brut, rue: rue, cp: cp, ville: ville, siege: !!(r.siege && r.siege.siret === e.siret) });
    });
  });
  var uniques = function (l) { var v = {}; return l.filter(function (c) { if (v[c.siret]) return false; v[c.siret] = 1; return true; }); };
  var parRue = uniques(cands.filter(function (c) { return c.rue && (c.cp || c.ville || !ent.cp); }));
  if (parRue.length === 1) return { siret: parRue[0].siret, c: parRue[0], raison: "nom + rue" };
  if (parRue.length > 1) {
    var memeCp = parRue.filter(function (c) { return c.cp; });
    if (memeCp.length === 1) return { siret: memeCp[0].siret, c: memeCp[0], raison: "nom + rue + code postal" };
    return { propositions: parRue.slice(0, 5), raison: "plusieurs établissements dans cette rue" };
  }
  var parCp = uniques(cands.filter(function (c) { return c.cp; }));
  if (parCp.length === 1) return { siret: parCp[0].siret, c: parCp[0], raison: "nom + code postal (seul établissement)" };
  var parVille = uniques(cands.filter(function (c) { return c.ville; }));
  if (!parCp.length && parVille.length === 1) return { siret: parVille[0].siret, c: parVille[0], raison: "nom + commune (seul établissement)" };
  var props = uniques(parCp.length ? parCp : parVille.length ? parVille : cands);
  return { propositions: props.slice(0, 5), raison: !cands.length ? "aucune entreprise SIRENE de ce nom" : !ent.cp && !ent.ville ? "pas d'adresse dans le CRM pour départager" : "adresse non concordante ou plusieurs établissements" };
}

/** Complétion en lot : resultats = { ID_ENT : [résultats SIRENE] } (fournis par la page). */
function handleCompleterSirets(p, ss) {
  var isAdmin = String(p.roleAuteur || "") === "Administrateur" || PARAM_MASTER_EMAILS.indexOf(String(p.emailAuteur || "").toLowerCase()) !== -1;
  if (!isAdmin) return createJsonResponse({ success: false, message: "Réservé aux administrateurs." });
  var res = {};
  try { res = typeof p.resultats === "string" ? JSON.parse(p.resultats) : (p.resultats || {}); } catch (e) { res = {}; }
  var ids = Object.keys(res).map(function (k) { return String(k).trim().toUpperCase(); });
  if (!ids.length) return createJsonResponse({ success: false, message: "Aucune entreprise transmise." });
  Object.keys(res).forEach(function (k) { var K = String(k).trim().toUpperCase(); if (K !== k) res[K] = res[k]; });
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return createJsonResponse({ success: false, message: "Une autre mise à jour des entreprises est en cours, réessayez." });
  try {
    var ws = getSheetSafe(ss, "PARTENARIAT");
    var d = ws.getDataRange().getValues();
    var C = _adrColsLecture_(d[0]);
    var ents = {};
    for (var i = 1; i < d.length; i++) {
      var id = String(d[i][1] || "").trim().toUpperCase();
      if (!id || !res[id]) continue;
      ents[id] = ents[id] || { id: id, nom: String(d[i][2] || "").trim(), lignes: [], siretActuel: "" };
      var a = _adrDeLigne_(d[i], C);
      ents[id].lignes.push(a);
      if (a.siret) ents[id].siretActuel = a.siret;
    }
    var ecrire = {}, ecrits = [], incertains = [];
    ids.forEach(function (id) {
      var e = ents[id]; if (!e || e.siretActuel) return;   // jamais d'écrasement d'un SIRET existant
      var ref = _adrReference_(e.lignes);
      var ch = _siretChoisir_({ nom: e.nom, adresse: ref.adresse, cp: String(ref.cp || ""), ville: ref.ville }, res[id]);
      var adr = [ref.adresse, ref.cp, ref.ville].filter(Boolean).join(" ");
      if (ch.siret && _siretValide_(ch.siret)) {
        ecrire[id] = ch.siret;
        ecrits.push({ id: id, nom: e.nom, adresseCrm: adr, siret: ch.siret, nomSirene: ch.c.nom, adresseSirene: ch.c.adresse, raison: ch.raison });
      } else incertains.push({ id: id, nom: e.nom, adresseCrm: adr, raison: ch.raison, propositions: (ch.propositions || []).map(function (c) { return { siret: c.siret, nom: c.nom, adresse: c.adresse }; }) });
    });
    var modifs = Object.keys(ecrire).length ? _siretEcrire_(ss, ecrire) : 0;
    _siretRapport_(ss, ecrits, incertains, p.nomAuteur || p.emailAuteur || "");
    var resume = ids.length + " entreprises sans SIRET recherchées — " + ecrits.length + " SIRET importés (nom + adresse), " + incertains.length + " à vérifier à la main";
    logAction(p.nomAuteur || p.idAuteur || "Administrateur", p.roleAuteur || "", "Modification", "Entreprises — SIRET", "PARTENARIAT",
      resume + (ecrits.length ? " | " + ecrits.slice(0, 40).map(function (x) { return x.nom + " → " + x.siret; }).join(" ; ") : ""));
    try { _invalidatePartnersCache_(); } catch (eC) {}
    return createJsonResponse({ success: true, resume: resume, lignesModifiees: modifs, ecrits: ecrits, incertains: incertains });
  } catch (e) {
    return createJsonResponse({ success: false, message: e.toString() });
  } finally { lock.releaseLock(); }
}

function _siretRapport_(ss, ecrits, incertains, par) {
  var ws = ss.getSheetByName(SIRET_RAPPORT_FEUILLE) || ss.insertSheet(SIRET_RAPPORT_FEUILLE);
  ws.clear();
  var rows = [["Complétion des SIRET — " + Utilities.formatDate(new Date(), "Europe/Paris", "dd/MM/yyyy HH:mm") + " — par " + par, "", "", "", "", "", ""],
    ["Résultat", "ID entreprise", "Entreprise (CRM)", "Adresse CRM", "SIRET", "Établissement SIRENE", "Motif / propositions"]];
  ecrits.forEach(function (x) { rows.push(["Importé", x.id, x.nom, x.adresseCrm, "'" + x.siret, x.nomSirene + " — " + x.adresseSirene, x.raison]); });
  incertains.forEach(function (x) { rows.push(["À vérifier", x.id, x.nom, x.adresseCrm, "", "", x.raison + (x.propositions.length ? " | " + x.propositions.map(function (c) { return c.siret + " " + c.nom + " — " + c.adresse; }).join(" ; ") : "")]); });
  ws.getRange(1, 1, rows.length, 7).setValues(rows);
  ws.getRange(1, 1).setFontWeight("bold");
  ws.getRange(2, 1, 1, 7).setFontWeight("bold").setBackground("#1A4E8A").setFontColor("#ffffff");
  ws.setFrozenRows(2);
}
