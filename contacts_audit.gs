// =========================================================================
// AUDIT D'UN LIBELLÉ GOOGLE CONTACTS (ex. HECG_AIX-BTS CG 2) — réservé direction
// Compare chaque membre du libellé au CRM (ETUDIANTS : classe, campus, statut).
// Lecture seule par défaut ; { appliquer:true } RETIRE seulement du libellé les
// membres en écart (contactGroups/members:modify) — aucune fiche n'est supprimée.
// =========================================================================

function _caNorm_(s) { return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, ' ').trim().toLowerCase(); }

function _caEtudiants_(ss) {
  var d = getSheetSafe(ss, 'ETUDIANTS').getDataRange().getValues(), L = [];
  for (var i = 1; i < d.length; i++) {
    var id = String(d[i][0] || '').trim(); if (!id) continue;
    L.push({ id: id, nom: String(d[i][2] || '').trim(), prenom: String(d[i][3] || '').trim(),
      mail: String(d[i][4] || '').trim().toLowerCase(), mailHecg: String(d[i][26] || '').trim().toLowerCase(),
      typo: String(d[i][20] || '').trim(), classe: sc2NormLabel(d[i][14]), campus: sc2NormLabel(d[i][22]) });
  }
  return L;
}
// Campus CRM : « HECG_AIX » = Aix, « HECG » = Marseille ; classes d'Aix préfixées « AIX_ », celles de Marseille sans préfixe.
function _caCampusOk_(classe, campus) {
  var c = _caNorm_(classe).replace(/^(hecg_|form-|ma-)/, ''), k = _caNorm_(campus);
  if (!k) return true;
  return (k.indexOf('aix') >= 0) === (c.indexOf('aix') === 0);
}
function _caActif_(e) { var t = e.typo.toLowerCase(); return (t.indexOf('inscrit') >= 0 || t.indexOf('alternance') >= 0) && t.indexOf('sorti') < 0; }

function auditerLibelleContacts(libelle, compte, appliquer) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  compte = compte || 'm.mokhtari@hecg.fr';
  var token = sc2GetToken(compte), H = { headers: sc2AuthHeader(token) };
  // 1. groupes
  var groupes = [], pt = null;
  do {
    var g = sc2ApiCall('https://people.googleapis.com/v1/contactGroups?pageSize=1000' + (pt ? '&pageToken=' + encodeURIComponent(pt) : ''), H);
    if (g.error || g._parseError) throw new Error('Lecture des libellés : ' + JSON.stringify(g.error || g._raw));
    (g.contactGroups || []).forEach(function (x) { if (x.groupType === 'USER_CONTACT_GROUP') groupes.push(x); });
    pt = g.nextPageToken;
  } while (pt);
  var cible = _caNorm_(libelle);
  var trouves = groupes.filter(function (x) { return _caNorm_(x.name) === cible; });
  // libellés voisins (même classe sous une autre forme) pour repérer les doublons
  var coeur = cible.replace(/^hecg_|^form-|^ma-/, '').replace(/^(aix|mrs|marseille)[- ]*/, '');
  var voisins = groupes.filter(function (x) { return _caNorm_(x.name).indexOf(coeur) >= 0; })
    .map(function (x) { return { nom: x.name, membres: x.memberCount || 0 }; });
  if (!trouves.length) return { compte: compte, libelle: libelle, introuvable: true, voisins: voisins };

  var etus = _caEtudiants_(ss), parId = {}, parMail = {}, parNom = {};
  etus.forEach(function (e) {
    parId[e.id] = e;
    [e.mail, e.mailHecg].forEach(function (m) { if (m) parMail[m] = e; });
    var k = _caNorm_(e.nom + ' ' + e.prenom); (parNom[k] = parNom[k] || []).push(e);
  });
  var estDiff = /^hecg_/i.test(trouves[0].name), classeLib = trouves[0].name.replace(/^HECG_|^FORM-|^MA-/i, '');

  var res = [];
  trouves.forEach(function (grp) {
    var gd = sc2ApiCall('https://people.googleapis.com/v1/' + grp.resourceName + '?maxMembers=5000', H);
    var rns = gd.memberResourceNames || [];
    for (var i = 0; i < rns.length; i += 20) {
      var q = rns.slice(i, i + 20).map(function (r) { return 'resourceNames=' + encodeURIComponent(r); }).join('&');
      var b = sc2ApiCall('https://people.googleapis.com/v1/people:batchGet?personFields=names,emailAddresses,userDefined,memberships&' + q, H);
      (b.responses || []).forEach(function (r) {
        var p = r.person || {}, n = (p.names || [])[0] || {};
        var mails = (p.emailAddresses || []).map(function (m) { return String(m.value || '').toLowerCase(); });
        var cle = ((p.userDefined || []).filter(function (u) { return u.key === 'HECG_DIFFUSION_ID'; })[0] || {}).value || '';
        var e = null;
        if (/^ETU_/.test(cle)) e = parId[cle.slice(4)] || null;
        if (!e) mails.some(function (m) { return (e = parMail[m] || null); });
        if (!e) { var c = parNom[_caNorm_((n.familyName || '') + ' ' + (n.givenName || ''))]; if (c && c.length === 1) e = c[0]; }
        var autres = (p.memberships || []).map(function (m) { return m.contactGroupMembership && m.contactGroupMembership.contactGroupResourceName; })
          .filter(function (rn) { return rn && rn !== grp.resourceName && rn.indexOf('myContacts') < 0 && rn.indexOf('starred') < 0; })
          .map(function (rn) { var x = groupes.filter(function (y) { return y.resourceName === rn; })[0]; return x ? x.name : rn; });
        var verdict = 'ok', motif = '';
        if (!e) { verdict = 'inconnu'; motif = 'Aucun étudiant correspondant dans le CRM'; }
        else if (!_caActif_(e)) { verdict = 'sorti'; motif = 'Statut CRM : ' + (e.typo || '—'); }
        else if (_caNorm_(e.classe) !== _caNorm_(classeLib)) { verdict = 'classe'; motif = 'Classe CRM : ' + (e.classe || '—'); }
        else if (estDiff && cle !== 'ETU_' + e.id) { verdict = 'doublon'; motif = cle ? 'Fiche de diffusion d\'un autre élève (' + cle + ')' : 'Fiche complète (perso) dans un libellé de diffusion'; }
        var campusOk = !e || _caCampusOk_(classeLib, e.campus);
        res.push({ groupe: grp.name, rn: p.resourceName, nom: ((n.givenName || '') + ' ' + (n.familyName || '')).trim() || n.displayName || '?', mails: mails,
          cle: cle, idCrm: e ? e.id : '', classeCrm: e ? e.classe : '', campusCrm: e ? e.campus : '', statutCrm: e ? e.typo : '',
          verdict: verdict, motif: motif, campusIncoherent: !campusOk, autresLibelles: autres });
      });
    }
  });
  // élèves attendus absents du libellé
  var presents = {}; res.forEach(function (r) { if (r.idCrm) presents[r.idCrm] = 1; });
  var manquants = etus.filter(function (e) { return _caActif_(e) && _caNorm_(e.classe) === _caNorm_(classeLib) && !presents[e.id]; })
    .map(function (e) { return { id: e.id, nom: e.prenom + ' ' + e.nom, campus: e.campus, statut: e.typo, mailHecg: e.mailHecg }; });
  // incohérences dans le CRM lui-même : élèves de la classe dont le campus ne colle pas au préfixe
  var crmIncoh = etus.filter(function (e) { return _caActif_(e) && _caNorm_(e.classe) === _caNorm_(classeLib) && !_caCampusOk_(classeLib, e.campus); })
    .map(function (e) { return { id: e.id, nom: e.prenom + ' ' + e.nom, campus: e.campus, classe: e.classe }; });

  var retires = 0;
  if (appliquer) {
    var parG = {};
    res.filter(function (r) { return r.verdict !== 'ok'; }).forEach(function (r) {
      var gr = trouves.filter(function (x) { return x.name === r.groupe; })[0]; (parG[gr.resourceName] = parG[gr.resourceName] || []).push(r.rn);
    });
    Object.keys(parG).forEach(function (grn) {
      var out = sc2ApiCall('https://people.googleapis.com/v1/' + grn + '/members:modify', { method: 'post', contentType: 'application/json', headers: sc2AuthHeader(token), payload: JSON.stringify({ resourceNamesToRemove: parG[grn] }) });
      if (!out.error && !out._parseError) retires += parG[grn].length;
    });
  }
  return { compte: compte, libelle: libelle, groupes: trouves.map(function (x) { return { nom: x.name, membres: x.memberCount || 0 }; }), voisins: voisins,
    membres: res, manquants: manquants, crmIncoherents: crmIncoh, retires: retires };
}

function handleAuditLibelleContacts(p, ss) {
  if (String(p.emailAuteur || '').trim().toLowerCase() !== 'm.mokhtari@hecg.fr') return createJsonResponse({ success: false, message: 'Audit réservé à la direction.' });
  try {
    var compte = String(p.compte || 'm.mokhtari@hecg.fr').toLowerCase();
    if (SC2_TARGET_USERS.indexOf(compte) < 0) return createJsonResponse({ success: false, message: 'Compte non synchronisé : ' + compte });
    var r = auditerLibelleContacts(String(p.libelle || ''), compte, p.appliquer === true);
    if (p.appliquer === true) logAction(p.emailAuteur, p.roleAuteurValide || '', 'Modification', 'Contacts Google — Audit libellé', p.libelle, r.retires + ' membre(s) retiré(s) du libellé (' + compte + ')');
    return createJsonResponse({ success: true, audit: r });
  } catch (e) { return createJsonResponse({ success: false, message: e.message }); }
}
