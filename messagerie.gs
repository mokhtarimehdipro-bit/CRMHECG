// ==========================================
// MESSAGERIE INSTANTANÉE DU PERSONNEL (nov. 2026) — front : presence.js (pastille « en ligne »)
// Feuille MESSAGERIE : un message par ligne. Statuts : Envoyé (ligne créée) → Remis (le CRM du
// destinataire l'a reçu) → Vu (conversation ouverte par le destinataire).
// Relance par mail si un message n'est pas vu après 10 min (une seule fois par message),
// déclenchée par auditAdressesTick (toutes les 5 min).
// Action « getMessagerie » : préfixe get → pas d'invalidation du cache des données du CRM.
// ==========================================
var MSG_SHEET = "MESSAGERIE";
var MSG_HEADERS = ["ID", "Conversation", "De", "De (nom)", "À", "À (nom)", "Texte", "Envoyé le", "Remis le", "Vu le", "Relance mail le"];
var MSG_RELANCE_MIN = 10;

function _msgSheet_(ss) {
  ss = ss || SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");
  var sh = ss.getSheetByName(MSG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(MSG_SHEET);
    sh.getRange(1, 1, 1, MSG_HEADERS.length).setValues([MSG_HEADERS]).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}
function _msgConv_(a, b) { return [String(a).toLowerCase(), String(b).toLowerCase()].sort().join("|"); }
function _msgLire_(sh) {
  var n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, MSG_HEADERS.length).getValues().map(function (r, i) { r.__row = i + 2; return r; });
}
function _msgIso_(v) { if (!v) return ""; if (v instanceof Date) return v.toISOString(); return String(v); }
function _msgPub_(r, moi) {
  return { id: String(r[0]), de: String(r[2]).toLowerCase(), deNom: String(r[3]), a: String(r[4]).toLowerCase(), aNom: String(r[5]),
           texte: String(r[6]), envoye: _msgIso_(r[7]), remis: _msgIso_(r[8]), vu: _msgIso_(r[9]), moi: String(r[2]).toLowerCase() === moi };
}
// écrit une même valeur dans une colonne pour une liste de lignes (regroupe les lignes contiguës)
function _msgEcrireCol_(sh, rows, col, val) {
  rows.sort(function (a, b) { return a - b; });
  var i = 0;
  while (i < rows.length) {
    var j = i; while (j + 1 < rows.length && rows[j + 1] === rows[j] + 1) j++;
    var vals = []; for (var k = i; k <= j; k++) vals.push([val]);
    sh.getRange(rows[i], col, vals.length, 1).setValues(vals);
    i = j + 1;
  }
}

function handleGetMessagerie(p, ss) {
  try {
    var s = __TRACE_SESSION__ || {};
    if (!s.email || s.role === "Etudiant") return createJsonResponse({ success: false, message: "Accès refusé." });
    var moi = String(s.email).toLowerCase(), op = String(p.op || "boite");
    var sh = _msgSheet_(ss), now = new Date();

    if (op === "envoyer") {
      var a = String(p.a || "").trim().toLowerCase(), texte = String(p.texte || "").trim().slice(0, 2000);
      if (!a || a.indexOf("@") < 1 || a === moi) return createJsonResponse({ success: false, message: "Destinataire invalide." });
      if (!texte) return createJsonResponse({ success: false, message: "Message vide." });
      var lock = LockService.getScriptLock(); lock.waitLock(15000);
      try {
        var row = ["M" + now.getTime() + Math.floor(Math.random() * 1000), _msgConv_(moi, a), moi, s.nom || moi, a, String(p.aNom || a).slice(0, 80), texte, now.toISOString(), "", "", ""];
        sh.appendRow(row);
        try { logAction(s.nom || moi, s.role || "", "Message envoyé", "Messagerie — Envoi", a, "Message instantané à " + (p.aNom || a) + " (contenu non tracé)"); } catch (eL) {}
        return createJsonResponse({ success: true, message: _msgPub_(row, moi) });
      } finally { lock.releaseLock(); }
    }

    var rows = _msgLire_(sh);

    if (op === "fil") {
      var avec = String(p.avec || "").toLowerCase(), conv = _msgConv_(moi, avec);
      var fil = rows.filter(function (r) { return r[1] === conv; });
      // marquer remis + vu ce que l'autre m'a envoyé (conversation ouverte et visible)
      var aVoir = fil.filter(function (r) { return String(r[4]).toLowerCase() === moi && !r[9]; });
      if (aVoir.length && p.vu) {
        var iso = now.toISOString();
        var sansRemis = aVoir.filter(function (r) { return !r[8]; }).map(function (r) { return r.__row; });
        if (sansRemis.length) _msgEcrireCol_(sh, sansRemis, 9, iso);
        _msgEcrireCol_(sh, aVoir.map(function (r) { return r.__row; }), 10, iso);
        aVoir.forEach(function (r) { if (!r[8]) r[8] = iso; r[9] = iso; });
      }
      return createJsonResponse({ success: true, messages: fil.slice(-400).map(function (r) { return _msgPub_(r, moi); }) });
    }

    // op « boite » : conversations + non lus ; marque « remis » les nouveaux messages reçus
    var convs = {}, nouveaux = [], aRemettre = [], isoNow = now.toISOString();
    rows.forEach(function (r) {
      var de = String(r[2]).toLowerCase(), a = String(r[4]).toLowerCase();
      if (de !== moi && a !== moi) return;
      var autre = de === moi ? a : de, nomAutre = de === moi ? String(r[5]) : String(r[3]);
      var c = convs[autre] || (convs[autre] = { email: autre, nom: nomAutre || autre, nonLus: 0, dernier: null });
      if (de !== moi && nomAutre) c.nom = nomAutre;
      c.dernier = { texte: String(r[6]).slice(0, 90), moi: de === moi, quand: _msgIso_(r[7]) };
      if (a === moi && !r[9]) c.nonLus++;
      if (a === moi && !r[8]) { aRemettre.push(r.__row); nouveaux.push(_msgPub_(r, moi)); }
    });
    if (aRemettre.length) _msgEcrireCol_(sh, aRemettre, 9, isoNow);
    var liste = Object.keys(convs).map(function (k) { return convs[k]; })
      .sort(function (x, y) { return String(y.dernier.quand).localeCompare(String(x.dernier.quand)); });
    var total = liste.reduce(function (t, c) { return t + c.nonLus; }, 0);
    return createJsonResponse({ success: true, conversations: liste.slice(0, 40), nonLus: total, nouveaux: nouveaux.slice(-10) });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// Relance mail : messages non vus depuis plus de 10 min (une fois par message, regroupés par destinataire)
function _msgRelances_() {
  var ss = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ"), sh = ss.getSheetByName(MSG_SHEET);
  if (!sh || sh.getLastRow() < 2) return;
  var rows = _msgLire_(sh), lim = Date.now() - MSG_RELANCE_MIN * 60000, parDest = {};
  rows.forEach(function (r) {
    if (r[9] || r[10]) return;
    var t = new Date(r[7]).getTime();
    if (!t || t > lim || Date.now() - t > 3 * 86400000) return; // ignore au-delà de 3 jours
    var a = String(r[4]).toLowerCase();
    (parDest[a] = parDest[a] || []).push(r);
  });
  var esc = function (x) { return String(x).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); };
  var tz = Session.getScriptTimeZone() || "Europe/Paris", fait = [];
  Object.keys(parDest).forEach(function (a) {
    var l = parDest[a], prenom = String(l[0][5] || "").split(/\s+/).pop() || "";
    var lignes = l.map(function (r) {
      return '<tr><td style="padding:10px 14px;border-bottom:1px solid #eef2f7;vertical-align:top;white-space:nowrap;font-weight:700;color:#1A4E8A">' + esc(r[3]) + '</td>'
        + '<td style="padding:10px 14px;border-bottom:1px solid #eef2f7;color:#334155">' + esc(String(r[6]).slice(0, 300)).replace(/\n/g, "<br>") + '</td>'
        + '<td style="padding:10px 14px;border-bottom:1px solid #eef2f7;color:#94a3b8;white-space:nowrap;font-size:12px">' + Utilities.formatDate(new Date(r[7]), tz, "dd/MM HH:mm") + '</td></tr>';
    }).join("");
    var html = '<div style="font-family:Arial,Helvetica,sans-serif;max-width:620px;margin:auto;color:#0f172a">'
      + '<div style="background:linear-gradient(135deg,#1A4E8A,#2563eb);color:#fff;padding:20px 24px;border-radius:14px 14px 0 0;font-size:18px;font-weight:700">💬 ' + l.length + ' message' + (l.length > 1 ? 's' : '') + ' non lu' + (l.length > 1 ? 's' : '') + ' dans le CRM HECG</div>'
      + '<div style="border:1px solid #e6ecf4;border-top:0;border-radius:0 0 14px 14px;padding:18px 10px 22px">'
      + '<p style="margin:0 14px 14px">Bonjour,</p><p style="margin:0 14px 16px">Un ou plusieurs collègues vous ont écrit via la messagerie du CRM et attendent votre réponse :</p>'
      + '<table style="border-collapse:collapse;width:100%;font-size:14px">' + lignes + '</table>'
      + '<p style="margin:18px 14px 0">Pour répondre, ouvrez le CRM puis cliquez sur la pastille <b>👥 en ligne</b> en bas à droite de l\'écran.</p>'
      + '<p style="margin:16px 14px 0;color:#94a3b8;font-size:12px">Message automatique — relance envoyée une seule fois, ' + MSG_RELANCE_MIN + ' minutes après un message non lu.</p></div></div>';
    try {
      MailApp.sendEmail({ to: a, subject: "💬 Message" + (l.length > 1 ? "s" : "") + " non lu" + (l.length > 1 ? "s" : "") + " dans le CRM HECG — " + l.map(function (r) { return r[3]; }).filter(function (v, i, t) { return t.indexOf(v) === i; }).join(", "), htmlBody: html, name: "CRM HECG" });
      l.forEach(function (r) { fait.push(r.__row); });
    } catch (e) { console.error("relance messagerie " + a + " : " + e); }
  });
  if (fait.length) _msgEcrireCol_(sh, fait, 11, new Date().toISOString());
}
