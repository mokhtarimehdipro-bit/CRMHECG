// ==========================================
// PRÉSENCE DU PERSONNEL (nov. 2026) — front : presence.js (chargé par hecg-skins.js)
// Chaque page connectée envoie toutes les 60 s la page / le module en cours.
// Stockage en cache (pas de feuille) : une entrée par compte, oubliée après 3 min sans signal.
// Dernière activité de chacun conservée durablement (propriété de script PRESENCE_VU) pour
// afficher « vu le … » des collègues hors ligne ; à défaut, colonne « Dernière connexion » du Personnel.
// Renvoie aussi l'annuaire complet du personnel (onglet Personnel) pour écrire à tout le monde.
// Action en « get… » : lecture pour le versionnage des données (pas d'invalidation de cache).
// ==========================================
var PRESENCE_TTL_S = 180;

function _presVuLire_() { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty("PRESENCE_VU") || "{}") || {}; } catch (e) { return {}; } }

function _presAnnuaire_(ss) {
  var c = CacheService.getScriptCache(), hit = c.get("presence_annuaire_v1");
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  var d = getPersonnelDataCached_(ss), h = (d[0] || []).map(function (x) { return String(x).trim().toLowerCase(); });
  var ix = function (re) { for (var i = 0; i < h.length; i++) if (re.test(h[i])) return i; return -1; };
  var iMail = ix(/^mail/), iNom = ix(/^nomperso$/), iPre = ix(/^prenomperso$/), iFct = ix(/^fonction$/), iRole = ix(/^role$/), iDer = ix(/derni.re connexion/);
  var vus = {}, out = [];
  for (var r = 1; r < d.length; r++) {
    var em = String(d[r][iMail] || "").trim().toLowerCase();
    if (!em || em.indexOf("@") < 1 || vus[em]) continue;
    if (/etudiant/i.test(String(d[r][iRole] || ""))) continue;
    vus[em] = 1;
    var der = iDer > -1 ? d[r][iDer] : "";
    out.push({ email: em, nom: (String(d[r][iPre] || "").trim() + " " + String(d[r][iNom] || "").trim()).trim() || em,
               fonction: String(d[r][iFct] || ""), derniere: der instanceof Date ? der.toISOString() : String(der || "") });
  }
  out.sort(function (a, b) { return a.nom.localeCompare(b.nom); });
  try { c.put("presence_annuaire_v1", JSON.stringify(out), 600); } catch (e) {}
  return out;
}

function handleGetPresence(p, ss) {
  try {
    var s = __TRACE_SESSION__ || {};
    if (!s.email || s.role === "Etudiant") return createJsonResponse({ success: false, message: "Accès refusé." });
    var c = CacheService.getScriptCache(), now = Date.now(), moi = String(s.email).toLowerCase();
    var map = {}; try { map = JSON.parse(c.get("presence_v1") || "{}") || {}; } catch (e) { map = {}; }
    map[moi] = { nom: s.nom || moi, role: s.vieScRole || s.role || "", page: String(p.page || "").slice(0, 60), module: String(p.module || "").slice(0, 60), detail: String(p.detail || "").slice(0, 80), ts: now };
    Object.keys(map).forEach(function (k) { if (now - (map[k].ts || 0) > PRESENCE_TTL_S * 1000) delete map[k]; });
    c.put("presence_v1", JSON.stringify(map), PRESENCE_TTL_S * 4);
    // dernière activité durable
    var vu = _presVuLire_();
    vu[moi] = new Date(now).toISOString();
    try { PropertiesService.getScriptProperties().setProperty("PRESENCE_VU", JSON.stringify(vu)); } catch (eV) {}
    var autres = Object.keys(map).filter(function (k) { return k !== moi; }).map(function (k) {
      var x = map[k]; return { email: k, nom: x.nom, role: x.role, page: x.page, module: x.module, detail: x.detail, il: Math.round((now - x.ts) / 1000) };
    }).sort(function (a, b) { return a.nom.localeCompare(b.nom); });
    var annuaire = [];
    try {
      annuaire = _presAnnuaire_(ss).filter(function (x) { return x.email !== moi; }).map(function (x) {
        return { email: x.email, nom: x.nom, fonction: x.fonction, enLigne: !!map[x.email], vuLe: vu[x.email] || x.derniere || "" };
      });
    } catch (eA) {}
    return createJsonResponse({ success: true, autres: autres, annuaire: annuaire });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}
