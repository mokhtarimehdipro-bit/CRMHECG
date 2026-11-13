// Découverte (lecture seule) des fiches de positionnement historiques dans le Drive — réservé direction
function handleFpDecouverte(p, ss) {
  if (String(p.emailAuteur || '').trim().toLowerCase() !== 'm.mokhtari@hecg.fr') return createJsonResponse({ success: false, message: 'Réservé à la direction.' });
  try {
    var t0 = Date.now(), out = [], vus = {};
    var termes = (p.termes && p.termes.length) ? p.termes : ['positionnement'];
    termes.forEach(function (t) {
      var it = DriveApp.searchFiles("title contains '" + String(t).replace(/'/g, "\\'") + "' and trashed = false");
      while (it.hasNext() && out.length < (p.max || 3000) && Date.now() - t0 < 240000) {
        var f = it.next(), id = f.getId(); if (vus[id]) continue; vus[id] = 1;
        var par = '', parId = '', grand = '';
        try { var pi = f.getParents(); if (pi.hasNext()) { var pf = pi.next(); par = pf.getName(); parId = pf.getId(); var gi = pf.getParents(); if (gi.hasNext()) grand = gi.next().getName(); } } catch (e) {}
        var own = ''; try { own = f.getOwner() ? f.getOwner().getEmail() : ''; } catch (e) {}
        out.push({ id: id, nom: f.getName(), mime: f.getMimeType(), cree: Utilities.formatDate(f.getDateCreated(), 'Europe/Paris', 'yyyy-MM-dd'), modif: Utilities.formatDate(f.getLastUpdated(), 'Europe/Paris', 'yyyy-MM-dd'), taille: f.getSize(), dossier: par, dossierId: parId, parent2: grand, proprio: own });
      }
    });
    return createJsonResponse({ success: true, fichiers: out, ms: Date.now() - t0 });
  } catch (e) { return createJsonResponse({ success: false, message: String(e) }); }
}

// Lecture du contenu d'une fiche (xlsx/Sheet/Doc/PDF→texte limité) pour comprendre sa structure
function handleFpLireFiche(p, ss) {
  if (String(p.emailAuteur || '').trim().toLowerCase() !== 'm.mokhtari@hecg.fr') return createJsonResponse({ success: false, message: 'Réservé à la direction.' });
  try {
    var f = DriveApp.getFileById(p.id), mime = f.getMimeType(), sid = null, tmp = null;
    if (mime === MimeType.GOOGLE_SHEETS) sid = p.id;
    else if (/spreadsheetml|ms-excel/.test(mime)) {
      var r = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + p.id + '/copy?supportsAllDrives=true', { method: 'post', contentType: 'application/json', muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        payload: JSON.stringify({ name: 'TMP lecture fiche', mimeType: 'application/vnd.google-apps.spreadsheet' }) });
      if (r.getResponseCode() !== 200) throw new Error('Conversion : ' + r.getContentText().slice(0, 200));
      sid = tmp = JSON.parse(r.getContentText()).id;
    }
    var res = { mime: mime, onglets: [] };
    if (sid) {
      SpreadsheetApp.openById(sid).getSheets().forEach(function (sh) {
        var v = sh.getDataRange().getDisplayValues(), cells = [];
        for (var i = 0; i < v.length && i < 80; i++) for (var j = 0; j < v[i].length && j < 20; j++) if (String(v[i][j]).trim()) cells.push(String.fromCharCode(65 + j) + (i + 1) + '=' + String(v[i][j]).slice(0, 160));
        res.onglets.push({ nom: sh.getName(), cellules: cells });
      });
    } else if (mime === MimeType.GOOGLE_DOCS) res.texte = DocumentApp.openById(p.id).getBody().getText().slice(0, 8000);
    if (tmp) DriveApp.getFileById(tmp).setTrashed(true);
    return createJsonResponse({ success: true, fiche: res });
  } catch (e) { return createJsonResponse({ success: false, message: String(e) }); }
}

// Inventaire rapide via Drive API v3 (noms contenant un des termes, hors corbeille), avec dossier parent
function handleFpInventaire(p, ss) {
  if (String(p.emailAuteur || '').trim().toLowerCase() !== 'm.mokhtari@hecg.fr') return createJsonResponse({ success: false, message: 'Réservé à la direction.' });
  try {
    var tok = ScriptApp.getOAuthToken(), H = { headers: { Authorization: 'Bearer ' + tok }, muteHttpExceptions: true };
    var termes = (p.termes && p.termes.length) ? p.termes : ['positionnement', 'positionement', 'positionnment'];
    var q = '(' + termes.map(function (t) { return "name contains '" + t + "'"; }).join(' or ') + ') and trashed = false';
    var out = [], pt = '', parents = {};
    do {
      var url = 'https://www.googleapis.com/drive/v3/files?pageSize=1000&corpora=allDrives&includeItemsFromAllDrives=true&supportsAllDrives=true&q=' + encodeURIComponent(q)
        + '&fields=' + encodeURIComponent('nextPageToken,files(id,name,mimeType,createdTime,parents,owners(emailAddress))') + (pt ? '&pageToken=' + pt : '');
      var r = JSON.parse(UrlFetchApp.fetch(url, H).getContentText());
      if (r.error) throw new Error(JSON.stringify(r.error));
      (r.files || []).forEach(function (f) { out.push({ id: f.id, nom: f.name, mime: f.mimeType, cree: String(f.createdTime).slice(0, 10), parent: (f.parents || [])[0] || '', proprio: ((f.owners || [])[0] || {}).emailAddress || '' }); });
      pt = r.nextPageToken || '';
    } while (pt);
    // noms des dossiers parents (+ grand-parent), dédupliqués
    var ids = {}; out.forEach(function (f) { if (f.parent) ids[f.parent] = 1; });
    var info = function (id) {
      if (parents[id]) return parents[id];
      var r = JSON.parse(UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + id + '?supportsAllDrives=true&fields=name,parents', H).getContentText());
      return (parents[id] = { nom: r.name || '', parent: (r.parents || [])[0] || '' });
    };
    Object.keys(ids).forEach(function (id) { var a = info(id); if (a.parent) info(a.parent); });
    out.forEach(function (f) { var a = parents[f.parent] || {}; f.dossier = a.nom || ''; f.parent2 = a.parent && parents[a.parent] ? parents[a.parent].nom : ''; });
    return createJsonResponse({ success: true, fichiers: out });
  } catch (e) { return createJsonResponse({ success: false, message: String(e) }); }
}

function handleFpEtudiants(p, ss) {
  if (String(p.emailAuteur || '').trim().toLowerCase() !== 'm.mokhtari@hecg.fr') return createJsonResponse({ success: false, message: 'Réservé à la direction.' });
  return createJsonResponse({ success: true, etudiants: _vsuEtudiants_(ss).map(function (e) { return { id: e.id, nom: e.nom, prenom: e.prenom, classe: e.classe, typo: e.typo, statut: e.statut, idEleve: e.idEleve }; }) });
}
