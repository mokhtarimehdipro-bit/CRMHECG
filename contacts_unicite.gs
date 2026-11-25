// =========================================================================
// UNICITÉ DES FICHES GOOGLE CONTACTS (demande direction 2026-11-25)
// =========================================================================
// Règle : UNE seule fiche Google Contacts par personne du CRM (étudiant, MA, formateur,
// encadrant) et UN seul numéro par fiche — le numéro de la personne tel qu'il figure
// dans le CRM. Plusieurs fiches au même nom, ou plusieurs numéros sur une fiche, font
// que le téléphone / WhatsApp mélange les identités.
//
// 1. sc2PersonnesCrm      : liste CRM consolidée (une entrée par personne réelle : lignes
//                           ETUDIANTS en double, MA saisi 2 fois, ex-étudiant devenu MA…),
//                           et un numéro ne peut appartenir qu'à UNE personne.
// 2. sc2PlanUnicite       : compare les fiches Google d'un compte à cette liste →
//                           fiche retenue par personne, doublons, numéros à corriger,
//                           numéros CRM qui traînent sur d'autres fiches.
// 3. sc2AppliquerUnicite  : met le plan en œuvre (doublons → corbeille Google Contacts,
//                           récupérables 30 jours ; PATCH limité au champ téléphone).
// La synchro de 2 h (synchroniserTousLesContacts2) applique ce plan avant chaque compte,
// puis réécrit la fiche retenue (nom, prénom, entreprise, tuteur, libellés…).
// Les fiches de diffusion HECG_* / FORM-* (mail @hecg.fr seul, jamais de téléphone) sont
// tenues à part — voir sc2ExtractGroupesDiffusion.

var SC2_SUPPR_MAX_AUTO = 300;   // au-delà, la synchro de nuit n'efface rien (contrôle manuel dans Pilotage Workspace)

function sc2NomF_(s) { return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim(); }
function sc2Tel1_(c) { var p = (c.phones || (c.phone ? [c.phone] : [])).filter(Boolean); return p.length ? sc2PhoneKey(p[0]) : ''; }
function sc2TelVal_(c) { var p = (c.phones || (c.phone ? [c.phone] : [])).filter(Boolean); return p.length ? p[0] : ''; }
function sc2Mails_(c) { return (c.emails || []).map(function (e) { return String(e).trim().toLowerCase(); }).filter(Boolean); }
function sc2NomAff_(c) { return ((c.prenom || '') + ' ' + (c.nom || '')).trim() || c.entreprise || '?'; }
function sc2Cat_(c) { return c._cat === 'etu' ? 'Étudiant' : c._cat === 'part' ? 'Entreprise / MA' : 'Formateur / encadrant'; }

// Plus petit = prioritaire pour porter la fiche (et le numéro en cas de numéro partagé).
function sc2Priorite_(c) {
  var l = (c.labels || [c.label])[0];
  if (c._cat === 'etu') return l === 'Étudiant Inscrit' ? 0 : l === 'Étudiant Préinscrit' ? 1 : l === 'Prospect' ? 6 : 7;
  if (c._cat === 'perso') return (l === 'Formateurs' || l === 'Encadrants') ? 2 : 8;
  return 3;
}

function sc2MemePersonne_(a, b) {
  var ta = sc2Tel1_(a), tb = sc2Tel1_(b);
  if (ta && ta === tb) return true;
  var mb = sc2Mails_(b);
  if (sc2Mails_(a).some(function (x) { return mb.indexOf(x) >= 0; })) return true;
  if (a._cat !== b._cat) return false;                      // ex. étudiant et MA homonymes sans n° ni mail commun
  if (a._cat === 'etu' || a._cat === 'perso') return true;  // même nom + prénom dans la même feuille = même personne
  // Partenariat : homonymes distincts seulement si numéros ET entreprises diffèrent
  return !(ta && tb && ta !== tb && sc2NomF_(a.entreprise) !== sc2NomF_(b.entreprise));
}

function sc2Fusionner_(g) {
  g.sort(function (a, b) { return sc2Priorite_(a) - sc2Priorite_(b) || b._rang - a._rang; });
  var base = g[0];
  if (g.length === 1) return base;
  var m = Object.assign({}, base), labels = new Set((base.labels || [base.label]).filter(Boolean)), vues = {};
  vues[base._cat] = 1;
  g.slice(1).forEach(function (x) {   // libellés d'une AUTRE feuille seulement (jamais « Sorti » + « Inscrit » à la fois)
    if (vues[x._cat]) return; vues[x._cat] = 1;
    (x.labels || [x.label]).forEach(function (l) { if (l) labels.add(l); });
  });
  m.labels = Array.from(labels); delete m.label;
  if (!sc2Tel1_(m)) { var t = g.filter(function (x) { return sc2Tel1_(x); })[0]; m.phones = t ? [sc2TelVal_(t)] : []; }
  if (!sc2Mails_(m).length) { var e = g.filter(function (x) { return sc2Mails_(x).length; })[0]; if (e) m.emails = e.emails.slice(0, 1); }
  if (!m.entreprise) { var en = g.filter(function (x) { return x.entreprise; })[0]; if (en) m.entreprise = en.entreprise; }
  if (!m.contactEnt) { var ce = g.filter(function (x) { return x.contactEnt; })[0]; if (ce) m.contactEnt = ce.contactEnt; }
  var etus = new Set(); g.forEach(function (x) { (x.etudiants || []).forEach(function (n) { etus.add(n); }); });
  if (etus.size) m.etudiants = Array.from(etus);
  m._fusion = g.map(function (x) { return sc2Cat_(x) + (x._cat === 'etu' ? ' (' + ((x.labels || [])[0] || '') + ')' : x.entreprise ? ' (' + x.entreprise + ')' : ''); });
  return m;
}

// Liste CRM consolidée + rapport (fusions, numéros partagés). Remplit SC2_TELS_CRM.
function sc2PersonnesCrm(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var tous = [];
  sc2ExtractEtudiants(ss).forEach(function (c, i) { c._cat = 'etu'; c._rang = i; tous.push(c); });
  sc2ExtractPartenaires(ss).forEach(function (c, i) { c._cat = 'part'; c._rang = i; tous.push(c); });
  sc2ExtractPersonnel(ss).forEach(function (c, i) { c._cat = 'perso'; c._rang = i; tous.push(c); });
  tous = tous.filter(function (c) { return sc2Tel1_(c) || c.entreprise || c.nom || c.prenom; });

  var parent = new Map();
  tous.forEach(function (c) { parent.set(c, c); });
  function racine(c) { while (parent.get(c) !== c) c = parent.get(c); return c; }
  var parNk = {};
  tous.forEach(function (c) { var k = sc2NameKey(c.nom, c.prenom); if (k && sc2NomF_(c.nom) && sc2NomF_(c.prenom)) (parNk[k] = parNk[k] || []).push(c); });
  Object.keys(parNk).forEach(function (k) {
    var g = parNk[k];
    for (var i = 0; i < g.length; i++) for (var j = i + 1; j < g.length; j++) {
      if (sc2MemePersonne_(g[i], g[j])) { var a = racine(g[i]), b = racine(g[j]); if (a !== b) parent.set(b, a); }
    }
  });
  var groupes = new Map();
  tous.forEach(function (c) { var r = racine(c); if (!groupes.has(r)) groupes.set(r, []); groupes.get(r).push(c); });

  var personnes = [], fusions = [];
  groupes.forEach(function (g) { var m = sc2Fusionner_(g); if (g.length > 1) fusions.push({ nom: sc2NomAff_(m), sources: m._fusion }); personnes.push(m); });

  // Un numéro = une seule personne (la prioritaire) ; les autres perdent ce numéro sur leur fiche.
  personnes.sort(function (a, b) { return sc2Priorite_(a) - sc2Priorite_(b); });
  var proprio = {}, conflits = [];
  personnes.forEach(function (c) {
    var k = sc2Tel1_(c); if (!k) return;
    if (proprio[k]) { conflits.push({ tel: sc2TelVal_(c), garde: sc2NomAff_(proprio[k]) + ' — ' + sc2Cat_(proprio[k]), retire: sc2NomAff_(c) + ' — ' + sc2Cat_(c) }); c.phones = []; c._telRetire = true; }
    else proprio[k] = c;
  });
  sc2IndexerTelephones(personnes);
  return { personnes: personnes, fusions: fusions, conflits: conflits, proprio: proprio };
}

// Plan pour UN compte Google (existants = sc2LireContactsExistants).
function sc2PlanUnicite(existants, crm) {
  var personnes = crm.personnes, proprio = crm.proprio;
  var fiches = existants.filter(function (e) { return !sc2EstContactDiffusion(e); });
  var nbDiffusion = existants.length - fiches.length;
  var idxNk = {}, idxMail = {}, idxTel = {};
  fiches.forEach(function (f) {
    var n = (f.names || [])[0] || {}, k = sc2NameKey(n.familyName, n.givenName);
    if (k) (idxNk[k] = idxNk[k] || []).push(f);
    (f.emailAddresses || []).forEach(function (e) { var m = String(e.value || '').trim().toLowerCase(); if (m) (idxMail[m] = idxMail[m] || []).push(f); });
    (f.phoneNumbers || []).forEach(function (p) { var t = sc2PhoneKey(p.value); if (t) (idxTel[t] = idxTel[t] || []).push(f); });
  });
  function telsDe(f) { return (f.phoneNumbers || []).map(function (p) { return sc2PhoneKey(p.value); }).filter(Boolean); }

  // Paires (personne, fiche) candidates avec score ; nkDe[rn] = personnes dont la fiche porte exactement le nom + prénom
  var paires = [], nkDe = {};
  personnes.forEach(function (c, pi) {
    var sc = new Map(), tel = sc2Tel1_(c), fam = sc2NomF_(c.nom);
    function plus(f, s) { sc.set(f, (sc.get(f) || 0) + s); }
    (idxNk[sc2NameKey(c.nom, c.prenom)] || []).forEach(function (f) { plus(f, 4); (nkDe[f.resourceName] = nkDe[f.resourceName] || []).push(pi); });
    sc2Mails_(c).forEach(function (m) { (idxMail[m] || []).forEach(function (f) { plus(f, 2); }); });
    if (tel) (idxTel[tel] || []).forEach(function (f) { if (sc2NomF_(((f.names || [])[0] || {}).familyName) === fam) plus(f, 3); });
    sc.forEach(function (s, f) {
      var t = telsDe(f), bonus = (tel && t.length === 1 && t[0] === tel ? 3 : 0) + (tel && t.indexOf(tel) >= 0 ? 1 : 0)
        + (c.entreprise && sc2NomF_(((f.organizations || [])[0] || {}).name) === sc2NomF_(c.entreprise) ? 1 : 0)
        + Math.min((f.memberships || []).length, 5) * 0.1 - t.length * 0.05;
      paires.push({ pi: pi, f: f, s: s, b: bonus });
    });
  });
  paires.sort(function (a, b) { return b.s - a.s || b.b - a.b; });
  var aff = new Map(), prise = {};
  paires.forEach(function (p) { if (aff.has(p.pi) || prise[p.f.resourceName]) return; aff.set(p.pi, p.f); prise[p.f.resourceName] = p.pi; });

  var plan = { fiches: fiches.length, diffusion: nbDiffusion, personnes: personnes.length, aff: aff,
    doublons: [], telFix: [], orphelins: [], sansFiche: [], horsCrmMultiTel: 0 };

  // Doublons : fiche non retenue portant exactement le nom + prénom d'une personne qui a déjà sa fiche.
  fiches.forEach(function (f) {
    if (prise[f.resourceName] !== undefined) return;
    var pis = (nkDe[f.resourceName] || []).filter(function (pi) { return aff.has(pi); });
    if (!pis.length) return;
    var c = personnes[pis[0]], g = aff.get(pis[0]);
    plan.doublons.push({ rn: f.resourceName, nom: sc2NomAff_(c), cat: sc2Cat_(c),
      tels: (f.phoneNumbers || []).map(function (p) { return p.value; }), garde: g.resourceName,
      telsGarde: (g.phoneNumbers || []).map(function (p) { return p.value; }) });
    prise[f.resourceName] = -1;
  });

  // Fiche retenue : un seul numéro, celui du CRM (à défaut un n° Google qui n'appartient à personne d'autre).
  var utilises = {};
  Object.keys(proprio).forEach(function (k) { utilises[k] = 1; });
  aff.forEach(function (f, pi) {
    var c = personnes[pi], tel = sc2Tel1_(c), avant = (f.phoneNumbers || []).map(function (p) { return p.value; });
    var cible = '';
    if (tel) cible = sc2TelVal_(c);
    else {
      var libre = avant.filter(function (v) { var k = sc2PhoneKey(v); return k && !utilises[k] && !sc2NumeroAutrui(c, v); })[0];
      if (libre) { cible = sc2FormatPhone(libre) || libre; utilises[sc2PhoneKey(libre)] = 1; }
    }
    var ok = cible ? (avant.length === 1 && sc2PhoneKey(avant[0]) === sc2PhoneKey(cible)) : !avant.length;
    if (!ok) plan.telFix.push({ rn: f.resourceName, etag: f.etag, nom: sc2NomAff_(c), cat: sc2Cat_(c), avant: avant, apres: cible,
      motif: avant.length > 1 ? 'plusieurs numéros' : !avant.length ? 'numéro manquant' : !cible ? 'numéro d\'une autre personne' : 'numéro différent du CRM' });
  });

  // Fiches sans lien avec le CRM portant le numéro d'une personne du CRM → on retire ce numéro.
  fiches.forEach(function (f) {
    if (prise[f.resourceName] !== undefined) return;
    var ph = f.phoneNumbers || [];
    if (ph.length > 1) plan.horsCrmMultiTel++;
    var retirer = ph.filter(function (p) { return proprio[sc2PhoneKey(p.value)]; });
    if (!retirer.length) return;
    var n = (f.names || [])[0] || {};
    plan.orphelins.push({ rn: f.resourceName, etag: f.etag, nom: ((n.givenName || '') + ' ' + (n.familyName || '')).trim() || '(sans nom)',
      retirer: retirer.map(function (p) { return p.value + ' → ' + sc2NomAff_(proprio[sc2PhoneKey(p.value)]); }),
      garder: ph.filter(function (p) { return !proprio[sc2PhoneKey(p.value)]; }).map(function (p) { return p.value; }) });
  });

  personnes.forEach(function (c, pi) { if (!aff.has(pi)) plan.sansFiche.push(sc2NomAff_(c) + ' — ' + sc2Cat_(c)); });
  return plan;
}

// Applique le plan. opts.supprimer=false → n'efface aucun doublon. Retourne un bilan.
function sc2AppliquerUnicite(token, plan, opts) {
  opts = opts || {};
  var bilan = { supprimes: 0, telCorriges: 0, telRetires: 0, erreurs: [] };
  if (opts.supprimer !== false && plan.doublons.length) {
    var rns = plan.doublons.map(function (d) { return d.rn; });
    for (var i = 0; i < rns.length; i += 400) {
      var r = sc2ApiCall('https://people.googleapis.com/v1/people:batchDeleteContacts', { method: 'post', contentType: 'application/json',
        headers: sc2AuthHeader(token), payload: JSON.stringify({ resourceNames: rns.slice(i, i + 400) }) });
      if (r.error || r._parseError) bilan.erreurs.push('Suppression doublons : ' + (r.error ? r.error.message : 'réponse non-JSON'));
      else bilan.supprimes += Math.min(400, rns.length - i);
    }
  }
  var maj = plan.telFix.map(function (t) { return { rn: t.rn, etag: t.etag, phones: t.apres ? [t.apres] : [], fix: 1 }; })
    .concat(plan.orphelins.map(function (o) { return { rn: o.rn, etag: o.etag, phones: o.garder, fix: 0 }; }));
  for (var j = 0; j < maj.length; j += 150) {
    var lot = maj.slice(j, j + 150), etags = null;
    for (var essai = 0; essai < 2; essai++) {
      var contacts = {};
      lot.forEach(function (u) { contacts[u.rn] = { etag: (etags && etags[u.rn]) || u.etag, phoneNumbers: u.phones.map(function (v, k) { return { value: v, type: k ? 'other' : 'mobile' }; }) }; });
      var res = sc2ApiCall('https://people.googleapis.com/v1/people:batchUpdateContacts', { method: 'post', contentType: 'application/json',
        headers: sc2AuthHeader(token), payload: JSON.stringify({ contacts: contacts, updateMask: 'phoneNumbers', readMask: 'names' }) });
      if (!res.error && !res._parseError) {
        lot.forEach(function (u) { if (u.fix) bilan.telCorriges++; else bilan.telRetires++; });
        break;
      }
      if (essai === 0 && res.error && /etag|FAILED_PRECONDITION/i.test(String(res.error.message || '') + res.error.status)) { etags = sc2RafraichirEtags(token, lot.map(function (u) { return u.rn; })); continue; }
      bilan.erreurs.push('Numéros (lot ' + (Math.floor(j / 150) + 1) + ') : ' + (res.error ? res.error.message : 'réponse non-JSON'));
      break;
    }
  }
  return bilan;
}

// Résumé compact d'un plan (page Pilotage Workspace / journal)
function sc2ResumePlan_(compte, plan, max) {
  max = max || 400;
  return { compte: compte, fiches: plan.fiches, diffusion: plan.diffusion, personnes: plan.personnes,
    nbDoublons: plan.doublons.length, nbTelFix: plan.telFix.length, nbOrphelins: plan.orphelins.length, nbSansFiche: plan.sansFiche.length,
    horsCrmMultiTel: plan.horsCrmMultiTel,
    doublons: plan.doublons.slice(0, max).map(function (d) { return { nom: d.nom, cat: d.cat, tels: d.tels, telsGarde: d.telsGarde }; }),
    telFix: plan.telFix.slice(0, max).map(function (t) { return { nom: t.nom, cat: t.cat, avant: t.avant, apres: t.apres, motif: t.motif }; }),
    orphelins: plan.orphelins.slice(0, max).map(function (o) { return { nom: o.nom, retirer: o.retirer, garder: o.garder }; }),
    sansFiche: plan.sansFiche.slice(0, 60) };
}

// ── Page Pilotage Workspace : audit / correction d'un compte ──
function _pilContactsAudit_(ss, compte, appliquer) {
  var comptes = compte ? [compte] : SC2_TARGET_USERS.slice();
  comptes.forEach(function (c) { if (SC2_TARGET_USERS.indexOf(c) === -1) throw new Error('Compte non synchronisé : ' + c); });
  var crm = sc2PersonnesCrm(ss), out = [];
  comptes.forEach(function (c) {
    try {
      var token = sc2GetToken(c), existants = sc2LireContactsExistants(token), plan = sc2PlanUnicite(existants, crm);
      var r = sc2ResumePlan_(c, plan, compte ? 400 : 40);
      if (appliquer) r.bilan = sc2AppliquerUnicite(token, plan, { supprimer: true });
      out.push(r);
    } catch (e) { out.push({ compte: c, erreur: String(e.message || e) }); }
  });
  return { comptes: out, fusions: crm.fusions.slice(0, 300), nbFusions: crm.fusions.length, conflits: crm.conflits.slice(0, 300), nbConflits: crm.conflits.length,
    nbPersonnes: crm.personnes.length, sansTel: crm.personnes.filter(function (p) { return !sc2Tel1_(p); }).length };
}

// Éditeur : vérification à blanc de tous les comptes (journal d'exécution)
function verifierUniciteContacts() {
  var R = _pilContactsAudit_(SpreadsheetApp.getActiveSpreadsheet(), '', false);
  Logger.log('CRM : ' + R.nbPersonnes + ' personnes (' + R.nbFusions + ' fusion(s) de lignes en double, ' + R.nbConflits + ' numéro(s) partagé(s), ' + R.sansTel + ' sans numéro).');
  R.comptes.forEach(function (c) {
    if (c.erreur) { Logger.log(c.compte + ' : ERREUR ' + c.erreur); return; }
    Logger.log(c.compte + ' : ' + c.fiches + ' fiches (+' + c.diffusion + ' diffusion) | doublons ' + c.nbDoublons + ' | numéros à corriger ' + c.nbTelFix
      + ' | n° CRM sur fiches hors CRM ' + c.nbOrphelins + ' | personnes sans fiche ' + c.nbSansFiche);
  });
  return R;
}
