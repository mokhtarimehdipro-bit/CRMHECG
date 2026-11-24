// ══════════════════════════════════════════════════════════════════
//  CAMPAGNES — FILTRES AVANCÉS (comme les filtres du tableau de bord)
//  Stockés en JSON dans la colonne L « Filtres » de l'onglet CAMPAGNES
//  (CAMP_COL.FILTRES). Appliqués à l'aperçu (previewComMail) ET à l'envoi
//  (_resolveCampagneRecipients_) : ce que l'on voit = ce qui part.
//
//  Filtres étudiants (s'ajoutent aux cibles cochées, logique ET) :
//    campus[], classes[], statuts[] (sous-chaînes du statut), referent,
//    atelier (CV|CV_NON|LI|LI_NON|BOTH), cvtheque (OUI|NON),
//    creaDu/creaAu (aaaa-mm-jj, date de création de la fiche),
//    echange (7|7-30|30|NEVER, dernier échange du carnet de route),
//    motifSortie (cible « Sortis »).
//  Filtres entreprises :
//    entCreaDu/entCreaAu (date de création de l'entreprise),
//    maSeulement (contacts maîtres d'apprentissage uniquement).
//  Les emails ajoutés à la main ne sont jamais filtrés.
// ══════════════════════════════════════════════════════════════════

// Colonne L : CAMP_COL.FILTRES (défini dans code.gs — pas d'instruction de niveau fichier ici,
// l'ordre de chargement des fichiers Apps Script n'est pas garanti).

function _campNormFiltres_(f) {
  if (!f) return {};
  if (typeof f === "string") {
    var s = f.trim();
    if (s.charAt(0) !== "{") return {};          // anciennes lignes (ancien schéma) : ignorées
    try { f = JSON.parse(s); } catch (e) { return {}; }
  }
  if (typeof f !== "object") return {};
  function arr(v) { return Array.isArray(v) ? v.map(function(x) { return String(x || "").trim(); }).filter(Boolean) : []; }
  function str(v) { return String(v || "").trim(); }
  function dt(v) { v = str(v); return /^\d{4}-\d{2}-\d{2}$/.test(v) ? v : ""; }
  var out = {
    campus: arr(f.campus), classes: arr(f.classes), statuts: arr(f.statuts).map(function(x) { return x.toLowerCase(); }),
    referent: str(f.referent), atelier: str(f.atelier), cvtheque: str(f.cvtheque),
    creaDu: dt(f.creaDu), creaAu: dt(f.creaAu), echange: str(f.echange),
    motifSortie: (str(f.motifSortie) && str(f.motifSortie) !== "Tous") ? str(f.motifSortie) : "",
    entCreaDu: dt(f.entCreaDu), entCreaAu: dt(f.entCreaAu), maSeulement: f.maSeulement === true || f.maSeulement === "true",
    evenements: arr(f.evenements),                                   // prospects : venus à tel salon / événement
    vivier: f.vivier === true || f.vivier === "true",                // + vivier de l'onglet PROSPECTS (page Prospects)
    vivierSources: arr(f.vivierSources)
  };
  // Supprime les clés vides pour un stockage compact
  Object.keys(out).forEach(function(k) {
    var v = out[k];
    if (v === "" || v === false || (Array.isArray(v) && !v.length)) delete out[k];
  });
  return out;
}

function _campFiltresJson_(f) {
  var n = _campNormFiltres_(f);
  return Object.keys(n).length ? JSON.stringify(n) : "";
}

// Résumé lisible (cartes de campagne + Historique)
function _campFiltresResume_(f) {
  var n = _campNormFiltres_(f), parts = [];
  var CAMPUS = { "HECG": "Marseille", "HECG_AIX": "Aix" };
  var ECH = { "7": "échange < 7 j", "7-30": "échange 7-30 j", "30": "échange > 30 j", "NEVER": "jamais d'échange" };
  var AT = { CV: "atelier CV fait", CV_NON: "atelier CV non fait", LI: "LinkedIn fait", LI_NON: "LinkedIn non fait", BOTH: "CV + LinkedIn faits" };
  function fr(d) { return d ? d.split("-").reverse().join("/") : "…"; }
  if (n.campus) parts.push("campus " + n.campus.map(function(c) { return CAMPUS[c] || c; }).join("/"));
  if (n.classes) parts.push("classes " + n.classes.join(", "));
  if (n.statuts) parts.push("statut " + n.statuts.join("/"));
  if (n.referent) parts.push("référent " + n.referent);
  if (n.atelier) parts.push(AT[n.atelier] || n.atelier);
  if (n.cvtheque) parts.push("CVthèque " + (n.cvtheque === "OUI" ? "oui" : "non"));
  if (n.creaDu || n.creaAu) parts.push("fiche créée du " + fr(n.creaDu) + " au " + fr(n.creaAu));
  if (n.echange) parts.push(ECH[n.echange] || n.echange);
  if (n.motifSortie) parts.push("motif « " + n.motifSortie + " »");
  if (n.entCreaDu || n.entCreaAu) parts.push("entreprise créée du " + fr(n.entCreaDu) + " au " + fr(n.entCreaAu));
  if (n.maSeulement) parts.push("maîtres d'apprentissage seulement");
  if (n.evenements) parts.push("venus à : " + n.evenements.join(", "));
  if (n.vivier) parts.push("+ vivier prospects" + (n.vivierSources ? " (" + n.vivierSources.join(", ") + ")" : ""));
  return parts.join(" · ");
}

/**
 * Cœur du ciblage — utilisé par l'aperçu ET par l'envoi.
 * cnf : { target_recherche, b2c_preinscrits, b2c_prospects, b2c_inscrits, b2c_sortis, motif_sortie, b2b_avec, b2b_sans }
 * → [{ email, nom, prenom, source }]
 */
function _campCiblerContacts_(ss, cnf, filtres) {
  cnf = cnf || {};
  var f = _campNormFiltres_(filtres);
  var out = [];
  var motif = f.motifSortie || ((cnf.motif_sortie && cnf.motif_sortie !== "Tous") ? String(cnf.motif_sortie).trim() : "");
  var cibleSortis = cnf.b2c_sortis === true;
  var wantEtu = cnf.target_recherche || cnf.b2c_preinscrits || cnf.b2c_prospects || cnf.b2c_inscrits || cibleSortis;
  var wantEnt = !cibleSortis && (cnf.b2b_avec || cnf.b2b_sans);
  var dE = null;

  if (wantEtu) {
    dE = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues();
    var colCrea = (dE[0] || []).findIndex(function(h) { return String(h || "").trim().toLowerCase() === DATE_CREATION_HEADER.toLowerCase(); });
    var colEvts = _campColEvenements_(dE[0]);
    var dernierEch = null;
    if (f.echange) {
      dernierEch = {};
      var dC = getSheetSafe(ss, "Carnet_route").getDataRange().getValues();
      for (var c = 1; c < dC.length; c++) {
        var idC = String(dC[c][0] || "").trim(); if (!idC) continue;
        var t = new Date(dC[c][2]).getTime(); if (isNaN(t)) continue;
        if (!dernierEch[idC] || t > dernierEch[idC]) dernierEch[idC] = t;
      }
    }
    var now = Date.now();
    for (var i = 1; i < dE.length; i++) {
      var row = dE[i];
      var mail = String(row[4] || "").trim().toLowerCase();
      if (mail.indexOf("@") === -1) continue;
      var statut = String(row[13] || "").toLowerCase();
      var typo = String(row[20] || "").toLowerCase();
      var ok = false;
      if (cibleSortis) {
        if (typo.indexOf("sorti") !== -1 && (!motif || String(row[19] || "").trim() === motif)) ok = true;
      } else if (typo.indexOf("sorti") === -1 && typo.indexOf("refus") === -1) {
        var estPre = typo.indexOf("préin") !== -1 || typo.indexOf("prein") !== -1;
        var estInscritStrict = typo.indexOf("inscrit") !== -1 && !estPre;
        var estProspect = typo.indexOf("prospect") !== -1;
        if (cnf.target_recherche && statut.indexOf("recherche") !== -1 && (estInscritStrict || estPre)) ok = true;
        if (cnf.b2c_preinscrits && estPre) ok = true;
        if (cnf.b2c_prospects && estProspect) ok = true;
        if (cnf.b2c_inscrits && estInscritStrict) ok = true;
      }
      if (!ok) continue;

      // ── Filtres avancés (ET) ──
      if (f.campus && f.campus.indexOf(String(row[22] || "").trim()) === -1) continue;
      if (f.classes && f.classes.indexOf(String(row[14] || "").trim()) === -1) continue;
      if (f.statuts && !f.statuts.some(function(s) { return statut.indexOf(s) !== -1; })) continue;
      if (f.referent && String(row[7] || "").trim() !== f.referent) continue;
      if (f.atelier) {
        var cv = String(row[15] || "").trim().toLowerCase() === "oui", li = String(row[16] || "").trim().toLowerCase() === "oui";
        if (f.atelier === "CV" && !cv) continue;
        if (f.atelier === "CV_NON" && cv) continue;
        if (f.atelier === "LI" && !li) continue;
        if (f.atelier === "LI_NON" && li) continue;
        if (f.atelier === "BOTH" && !(cv && li)) continue;
      }
      if (f.cvtheque) {
        var cvt = String(row[24] || "").trim().toLowerCase() === "oui";
        if (f.cvtheque === "OUI" && !cvt) continue;
        if (f.cvtheque === "NON" && cvt) continue;
      }
      if (f.creaDu || f.creaAu) {
        var k = colCrea >= 0 ? _dateCreationKey_(row[colCrea]) : "";
        if (!k || (f.creaDu && k < f.creaDu) || (f.creaAu && k > f.creaAu)) continue; // fiche sans date : exclue (comme le tableau de bord)
      }
      if (f.echange) {
        var tE = dernierEch[String(row[0] || "").trim()];
        if (f.echange === "NEVER") { if (tE) continue; }
        else {
          if (!tE) continue;
          var j = (now - tE) / 86400000;
          if (f.echange === "7" && !(j <= 7)) continue;
          if (f.echange === "7-30" && !(j > 7 && j < 30)) continue;
          if (f.echange === "30" && !(j >= 30)) continue;
        }
      }
      if (f.evenements) {
        var evts = colEvts >= 0 ? String(row[colEvts] || "").split(" | ").map(function(x) { return x.trim(); }) : [];
        if (!evts.some(function(x) { return f.evenements.indexOf(x) !== -1; })) continue;
      }
      out.push({ email: mail, nom: String((row[3] || "") + " " + (row[2] || "")).trim(), prenom: String(row[3] || "").trim(), source: typo.indexOf("prospect") !== -1 ? "Prospect" : "Étudiant" });
    }
  }

  if (wantEnt) {
    // Onglet Partenariat : A id_contact | B id_entreprise | C nom entreprise | D nom | E prénom | G email | I email 2 | O maître d'apprentissage
    var dP = getSheetSafe(ss, "Partenariat").getDataRange().getValues();
    if (!dE) dE = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues();
    // Entreprise « avec alternants » = au moins un étudiant au statut alternance/apprenti (même règle que la base entreprises)
    var avecAlt = {};
    for (var e = 1; e < dE.length; e++) {
      var idEntEtu = String(dE[e][11] || "").trim().toUpperCase(); if (!idEntEtu) continue;
      var st = String(dE[e][13] || "").toLowerCase();
      if (st.indexOf("altern") !== -1 || st.indexOf("apprenti") !== -1) avecAlt[idEntEtu] = true;
    }
    // Date de création d'une entreprise = plus ancienne date de ses contacts, si tous en ont une
    var colCreaP = (dP[0] || []).findIndex(function(h) { return String(h || "").trim().toLowerCase() === DATE_CREATION_HEADER.toLowerCase(); });
    var creaEnt = {};
    if (f.entCreaDu || f.entCreaAu) {
      for (var q = 1; q < dP.length; q++) {
        var idq = String(dP[q][1] || "").trim().toUpperCase(); if (!idq) continue;
        var kq = colCreaP >= 0 ? _dateCreationKey_(dP[q][colCreaP]) : "";
        if (creaEnt[idq] === null) continue;
        if (!kq) { creaEnt[idq] = null; continue; }
        if (!creaEnt[idq] || kq < creaEnt[idq]) creaEnt[idq] = kq;
      }
    }
    for (var r = 1; r < dP.length; r++) {
      var rp = dP[r];
      var idEnt = String(rp[1] || "").trim().toUpperCase();
      var mailP = String(rp[6] || "").trim().toLowerCase();
      if (mailP.indexOf("@") === -1) mailP = String(rp[8] || "").trim().toLowerCase();
      if (mailP.indexOf("@") === -1) continue;
      var aAlt = !!avecAlt[idEnt];
      if (!((cnf.b2b_avec && aAlt) || (cnf.b2b_sans && !aAlt))) continue;
      if (f.maSeulement && String(rp[14] || "").trim().toLowerCase() !== "oui") continue;
      if (f.entCreaDu || f.entCreaAu) {
        var ke = creaEnt[idEnt];
        if (!ke || (f.entCreaDu && ke < f.entCreaDu) || (f.entCreaAu && ke > f.entCreaAu)) continue;
      }
      var nomC = String((rp[4] || "") + " " + (rp[3] || "")).trim();
      out.push({ email: mailP, nom: (nomC ? nomC + " — " : "") + String(rp[2] || "Entreprise").trim(), prenom: String(rp[4] || "").trim(), source: "Entreprise" });
    }
  }
  // Vivier de l'onglet PROSPECTS (inscriptions ateliers / JPO — page Prospects) : A id | B nom | C prénom | D email | F source
  if (f.vivier && !cibleSortis) {
    var wsV = ss.getSheetByName("PROSPECTS");
    var dV = wsV ? wsV.getDataRange().getValues() : [];
    for (var v = 1; v < dV.length; v++) {
      var mailV = String(dV[v][3] || "").trim().toLowerCase();
      if (mailV.indexOf("@") === -1) continue;
      if (f.vivierSources && f.vivierSources.indexOf(String(dV[v][5] || "").trim()) === -1) continue;
      out.push({ email: mailV, nom: String((dV[v][2] || "") + " " + (dV[v][1] || "")).trim(), prenom: String(dV[v][2] || "").trim(), source: "Vivier prospects" });
    }
  }
  return out;
}

function _campColEvenements_(entete) {
  return (entete || []).findIndex(function(h) {
    return String(h || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().toLowerCase() === "evenements";
  });
}

// Options des listes de filtres (classes, référents, campus présents dans la base)
function handleCampagneFiltresOptions(p, ss) {
  try {
    var dE = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues();
    var classes = {}, refs = {}, campus = {}, evts = {}, colEvts = _campColEvenements_(dE[0]);
    for (var i = 1; i < dE.length; i++) {
      var typo = String(dE[i][20] || "").toLowerCase();
      var c = String(dE[i][14] || "").trim(); if (c) classes[c] = true;
      var r = String(dE[i][7] || "").trim(); if (r && typo.indexOf("refus") === -1) refs[r] = true;
      var ca = String(dE[i][22] || "").trim(); if (ca) campus[ca] = true;
      if (colEvts >= 0) String(dE[i][colEvts] || "").split(" | ").forEach(function(x) { x = x.trim(); if (x) evts[x] = (evts[x] || 0) + 1; });
    }
    var sources = {};
    var wsV = ss.getSheetByName("PROSPECTS");
    if (wsV) wsV.getDataRange().getValues().slice(1).forEach(function(r) { var s = String(r[5] || "").trim(); if (s && String(r[3] || "").indexOf("@") !== -1) sources[s] = (sources[s] || 0) + 1; });
    function compte(o) { return tri(o).map(function(k) { return { v: k, n: o[k] }; }); }
    function tri(o) { return Object.keys(o).sort(function(a, b) { return a.localeCompare(b, "fr"); }); }
    return createJsonResponse({ success: true, classes: tri(classes), referents: tri(refs), campus: tri(campus), evenements: compte(evts), vivierSources: compte(sources) });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// En-tête de la colonne L « Filtres » (posé une fois, si absent)
function _campEnteteFiltres_(sheet) {
  try {
    if (sheet.getMaxColumns() < CAMP_COL.FILTRES + 1) sheet.insertColumnsAfter(sheet.getMaxColumns(), CAMP_COL.FILTRES + 1 - sheet.getMaxColumns());
    var h = sheet.getRange(1, CAMP_COL.FILTRES + 1);
    if (!String(h.getValue() || "").trim()) h.setValue("Filtres");
  } catch (e) {}
}
