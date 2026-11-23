// ==========================================
// PRÉSENCE DU PERSONNEL (nov. 2026) — front : presence.js (chargé par hecg-skins.js)
// Chaque page connectée envoie toutes les 60 s la page / le module en cours.
// Stockage en cache (pas de feuille) : une entrée par compte, oubliée après 3 min sans signal.
// Action en « get… » : lecture pour le versionnage des données (pas d'invalidation de cache).
// ==========================================
var PRESENCE_TTL_S = 180;

function handleGetPresence(p, ss) {
  try {
    var s = __TRACE_SESSION__ || {};
    if (!s.email || s.role === "Etudiant") return createJsonResponse({ success: false, message: "Accès refusé." });
    var c = CacheService.getScriptCache(), now = Date.now(), moi = String(s.email).toLowerCase();
    var map = {}; try { map = JSON.parse(c.get("presence_v1") || "{}") || {}; } catch (e) { map = {}; }
    map[moi] = { nom: s.nom || moi, role: s.vieScRole || s.role || "", page: String(p.page || "").slice(0, 60), module: String(p.module || "").slice(0, 60), detail: String(p.detail || "").slice(0, 80), ts: now };
    Object.keys(map).forEach(function (k) { if (now - (map[k].ts || 0) > PRESENCE_TTL_S * 1000) delete map[k]; });
    c.put("presence_v1", JSON.stringify(map), PRESENCE_TTL_S * 4);
    var autres = Object.keys(map).filter(function (k) { return k !== moi; }).map(function (k) {
      var x = map[k]; return { nom: x.nom, role: x.role, page: x.page, module: x.module, detail: x.detail, il: Math.round((now - x.ts) / 1000) };
    }).sort(function (a, b) { return a.nom.localeCompare(b.nom); });
    return createJsonResponse({ success: true, autres: autres });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}
