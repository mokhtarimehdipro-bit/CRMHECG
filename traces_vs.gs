// ════════════════════════════════════════════════════════════════════════════
//  TRAÇABILITÉ VIE SCOLAIRE — filet de sécurité du mouchard (HISTORIQUE)
//  Chaque appel venant de vie_scolaire.html (et tout appel « RDV », d'où qu'il vienne)
//  est tracé : écritures toujours (si le module ne l'a pas déjà fait), échecs toujours,
//  consultations dédoublonnées (1 ligne par personne / action / paramètres toutes les 10 min).
// ════════════════════════════════════════════════════════════════════════════
var __LOGS_REQ__ = 0;           // incrémenté par logAction pendant la requête
var __TRACE_SESSION__ = null;   // identité validée (doPostInner_)

var VS_TRACE_CATS = [
  [/rdv|visiteur|drapeau/i, "Suivi des RDV"],
  [/exam(?!.*exb)|examdoc/i, "Examens officiels"],
  [/exb/i, "Examens blancs"],
  [/memoire/i, "Mémoires"],
  [/frise/i, "Frises chronologiques"],
  [/^(get|save|delete|importer)pc|pcinventaire|pcannee/i, "Suivi des PC"],
  [/releve/i, "Relevés de notes"],
  [/note|resultat|referentielmatiere/i, "Notes"],
  [/qualiopi|formulaire|questionnaire/i, "Qualité / Qualiopi"],
  [/salle/i, "Salles"],
  [/personnel/i, "Personnel"],
  [/planning|reftache|migrateplanning/i, "Planning"],
  [/objectif|synthese/i, "Objectifs de la semaine"],
  [/mmok/i, "MMOK"],
  [/charlemagne/i, "Contrôle Charlemagne"],
  [/relance/i, "Relances candidatures"],
  [/tableausuivi|suiviimport|importsuivi/i, "Tableaux de suivi"],
  [/^ged/i, "GED"],
  [/ticket/i, "Tickets maintenance"],
  [/classe/i, "Classes"],
  [/inscription/i, "Inscriptions"],
  [/recruteur/i, "Recruteurs"],
  [/simulation/i, "Simulations"],
  [/stat/i, "Statistiques"],
  [/comparatifdirecteurs/i, "Comparatif directeurs"],
  [/mail/i, "Mails"],
  [/export/i, "Exports"]
];
var VS_TRACE_SECRETS = /^(sessionToken|action|_page|emailAuteur|nomAuteur|roleAuteur|roleAuteurValide|refAuteur|vieScRole|role|idAuteur)$|mdp|pass|token|secret|base64|blob|fichier64|contenu|data64/i;

function _vsTraceCat_(action) {
  for (var i = 0; i < VS_TRACE_CATS.length; i++) if (VS_TRACE_CATS[i][0].test(action)) return "Vie scolaire — " + VS_TRACE_CATS[i][1];
  return "Vie scolaire — Autres";
}
function _vsTraceType_(action, lecture, ok) {
  if (!ok) return "Échec";
  if (lecture) return "Consultation";
  if (/^(delete|archive|raz|masquer)/i.test(action)) return "Suppression";
  if (/^(create|upload|import|aspirer)/i.test(action)) return "Création";
  if (/^(send|install|lancer|sync|log)/i.test(action)) return "Action";
  return "Modification";
}
function _vsTraceVal_(v) {
  if (v === null || v === undefined || v === "") return "";
  if (Array.isArray(v)) return v.length > 5 ? "[" + v.length + " éléments]" : v.map(_vsTraceVal_).join(", ");
  if (typeof v === "object") { var k = Object.keys(v); return k.length > 6 ? "{" + k.length + " champs}" : k.filter(function (x) { return !VS_TRACE_SECRETS.test(x); }).map(function (x) { return x + "=" + _vsTraceVal_(v[x]); }).join(" ; "); }
  var s = String(v); return s.length > 120 ? s.slice(0, 117) + "…" : s;
}
function _vsTraceParams_(params) {
  return Object.keys(params).filter(function (k) { return !VS_TRACE_SECRETS.test(k); })
    .map(function (k) { var v = _vsTraceVal_(params[k]); return v ? k + " : " + v : ""; }).filter(String).join(" | ").slice(0, 900);
}
function _vsTraceCible_(params) {
  var c = params.id || params.etuId || params.idEtu || params.idEtudiant || params.etudiant || params.classe || params.folderId || "";
  if (!c && Array.isArray(params.ids)) c = params.ids.length === 1 ? params.ids[0] : params.ids.length + " éléments";
  return String(c || "-").slice(0, 80);
}

// Appelée par doPost après l'exécution de l'action
function _vsTracer_(params, out, ms) {
  try {
    var action = String(params.action || ""), page = String(params._page || "");
    if (!action || action === "login" || action === "getPresence") return;
    var vs = /^vie_scolaire/i.test(page) || /rdv|visiteur/i.test(action);
    if (!vs) return;
    var S = __TRACE_SESSION__ || {}, auteur = S.nom || S.email || params.nomAuteur || params.emailAuteur || "Inconnu";
    var role = S.vieScRole || S.role || "";
    var lecture = READ_ACTION_RE.test(action);
    var txt = ""; try { txt = String(out.getContent()); } catch (e) {}
    var ok = txt.indexOf('"success":false') === -1;
    if (!lecture && ok && __LOGS_REQ__ > 0) return;   // le module a déjà tracé l'écriture (avec plus de détails)
    var det = _vsTraceParams_(params);
    if (lecture && ok) {
      var cle = "vstr_" + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, auteur + "|" + action + "|" + det));
      var cache = CacheService.getScriptCache();
      if (cache.get(cle)) return;
      cache.put(cle, "1", 600);
    }
    var res = "";
    if (!ok) { var m = txt.match(/"(?:message|error)"\s*:\s*"([^"]{0,200})/); res = " ⇒ ÉCHEC" + (m ? " : " + m[1] : ""); }
    else if (!lecture) res = " ⇒ OK";
    logAction(auteur, role, _vsTraceType_(action, lecture, ok), _vsTraceCat_(action), _vsTraceCible_(params),
      "« " + action + " »" + (det ? " — " + det : "") + res + " (" + ms + " ms)");
  } catch (e) { console.error("_vsTracer_ : " + e); }
}
