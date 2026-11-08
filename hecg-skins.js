/* HECG — Thèmes d'apparence (niveau 14+ et compte maître) — v1, novembre 2026.
   Chargé juste après auth-fetch.js dans chaque page connectée (pas salon.html ni la connexion).
   1. Pose immédiatement <html data-skin="…"> depuis localStorage (pas de flash au chargement).
   2. Vérifie les droits auprès du serveur (action getApparence, mise en cache 6 h) : le thème
      enregistré côté serveur (feuille Personnel, colonne « Apparence ») fait foi → suit l'utilisateur.
   3. Si autorisé : bouton flottant 🎨 en bas à gauche → panneau des thèmes ; le choix est enregistré
      (saveApparence, tracé dans l'Historique « Paramètres — Apparence »).
   Styles : hecg-skins.css. */
(function () {
  if (window.HECG_SKINS) return;
  var API = "https://script.google.com/macros/s/AKfycbwfwxfVa2968AgKLdwwV8VIazVACVPVypKXrEUTc-zTuQ08gda8m5_qPfC2BgMyJMwb/exec";
  var K_SKIN = "hecg_skin", K_ELIG = "hecg_skin_elig", TTL = 6 * 3600 * 1000;
  var THEMES = [
    { id: "classique", nom: "Classique", desc: "L'identité HECG d'origine", c1: "#1A4E8A", c2: "#F26522", bg: "linear-gradient(135deg,#f5f7fb,#eef3fa)", glass: "rgba(255,255,255,.9)", gb: "#e6ebf2", ln: "#dce7f5" },
    { id: "aurore", nom: "Aurore", desc: "Violet & rose électrique", c1: "#8B5CF6", c2: "#EC4899", bg: "linear-gradient(135deg,#F5F3FF,#FDF2F8)", glass: "rgba(255,255,255,.6)", gb: "rgba(255,255,255,.85)", ln: "rgba(139,92,246,.25)" },
    { id: "ocean", nom: "Océan", desc: "Cyan lagon & turquoise", c1: "#0EA5E9", c2: "#14B8A6", bg: "linear-gradient(135deg,#ECFEFF,#F0FDFA)", glass: "rgba(255,255,255,.6)", gb: "rgba(255,255,255,.85)", ln: "rgba(14,165,233,.25)" },
    { id: "emeraude", nom: "Émeraude", desc: "Vert jade & citron vif", c1: "#10B981", c2: "#84CC16", bg: "linear-gradient(135deg,#ECFDF5,#F7FEE7)", glass: "rgba(255,255,255,.6)", gb: "rgba(255,255,255,.85)", ln: "rgba(16,185,129,.25)" },
    { id: "solaire", nom: "Solaire", desc: "Coucher de soleil corail", c1: "#F97316", c2: "#E11D48", bg: "linear-gradient(135deg,#FFF7ED,#FFF1F2)", glass: "rgba(255,255,255,.62)", gb: "rgba(255,255,255,.85)", ln: "rgba(249,115,22,.25)" },
    { id: "saphir", nom: "Saphir & or", desc: "Bleu profond, reflets dorés", c1: "#1D4ED8", c2: "#D4A017", bg: "linear-gradient(135deg,#EEF2FF,#FEFCE8)", glass: "rgba(255,255,255,.66)", gb: "rgba(255,255,255,.9)", ln: "rgba(29,78,216,.2)" },
    { id: "nuit", nom: "Nuit néon", desc: "Mode sombre, cyan & magenta", c1: "#22D3EE", c2: "#E879F9", bg: "linear-gradient(160deg,#060912,#120A24)", glass: "rgba(17,24,45,.6)", gb: "rgba(148,163,184,.2)", ln: "rgba(34,211,238,.3)" }
  ];
  var ids = THEMES.map(function (t) { return t.id; });

  function lsGet(k) { try { return localStorage.getItem(k); } catch (e) { return null; } }
  function lsSet(k, v) { try { localStorage.setItem(k, v); } catch (e) {} }
  function user() { try { return JSON.parse(lsGet("user") || "{}") || {}; } catch (e) { return {}; } }
  function email() { return String(user().email || user().mail || "").trim().toLowerCase(); }

  function apply(id) {
    if (ids.indexOf(id) === -1) id = "classique";
    var h = document.documentElement;
    if (id === "classique") h.removeAttribute("data-skin"); else h.setAttribute("data-skin", id);
    lsSet(K_SKIN, id);
    try { // Graphiques Chart.js : textes et grilles lisibles en mode sombre
      if (window.Chart && Chart.defaults) {
        var dark = id === "nuit";
        Chart.defaults.color = dark ? "#CBD5E1" : "#64748b";
        Chart.defaults.borderColor = dark ? "rgba(148,163,184,.16)" : "rgba(0,0,0,.1)";
        if (Chart.instances) Object.keys(Chart.instances).forEach(function (k) { try { Chart.instances[k].update("none"); } catch (e) {} });
      }
    } catch (e) {}
    var p = document.querySelector(".hsk-panel");
    if (p) p.querySelectorAll(".hsk-card").forEach(function (c) { c.classList.toggle("on", c.getAttribute("data-id") === id); });
  }

  // 1. Application immédiate + feuille de styles
  (function () {
    var s = lsGet(K_SKIN), u = user();
    if (u.role === "Etudiant") s = "classique";
    if (s && s !== "classique" && ids.indexOf(s) !== -1) document.documentElement.setAttribute("data-skin", s);
    var l = document.createElement("link"); l.rel = "stylesheet"; l.href = "hecg-skins.css?v=1";
    document.head.appendChild(l);
  })();

  function post(payload) {
    var u = user();
    payload.sessionToken = u.sessionToken || "";
    payload.idAuteur = u.nom || u.name || "";
    payload.emailAuteur = email();
    return fetch(API, { method: "POST", body: JSON.stringify(payload) }).then(function (r) { return r.json(); });
  }

  // 2. Droits
  function readElig() { try { var e = JSON.parse(lsGet(K_ELIG) || "null"); return e && e.email === email() ? e : null; } catch (er) { return null; } }
  function check(force) {
    var u = user();
    if (!u.sessionToken || u.role === "Etudiant") { apply("classique"); return; }
    var e = readElig();
    if (e) { if (e.ok) mountFab(); else if (lsGet(K_SKIN) !== "classique") apply("classique"); }
    if (!force && e && Date.now() - e.ts < TTL) return;
    post({ action: "getApparence" }).then(function (r) {
      if (!r || r.success === false) return;
      lsSet(K_ELIG, JSON.stringify({ email: email(), ok: !!r.eligible, ts: Date.now(), niveau: r.niveau, master: !!r.master }));
      if (r.eligible) { apply(r.theme || "classique"); mountFab(); }
      else { apply("classique"); unmountFab(); }
    }).catch(function () {});
  }

  // 3. Interface
  var SVG = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 22a10 10 0 1 1 10-10c0 2.2-1.8 3.5-4 3.5h-1.6a1.9 1.9 0 0 0-1.4 3.2 1.9 1.9 0 0 1-1.4 3.3H12Z"/><circle cx="7.5" cy="10.5" r="1.2" fill="#8B5CF6" stroke="none"/><circle cx="11" cy="6.8" r="1.2" fill="#EC4899" stroke="none"/><circle cx="15.5" cy="7.5" r="1.2" fill="#F97316" stroke="none"/><circle cx="17.2" cy="11.5" r="1.2" fill="#0EA5E9" stroke="none"/></svg>';
  function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }

  function mountFab() {
    if (!document.body) { document.addEventListener("DOMContentLoaded", mountFab, { once: true }); return; }
    if (document.querySelector(".hsk-fab")) return;
    var b = document.createElement("button");
    b.type = "button"; b.className = "hsk-fab"; b.title = "Personnaliser l'apparence"; b.setAttribute("aria-label", "Personnaliser l'apparence");
    b.innerHTML = SVG;
    b.addEventListener("click", function (ev) { ev.stopPropagation(); toggle(); });
    document.body.appendChild(b);
  }
  function unmountFab() { var b = document.querySelector(".hsk-fab"), p = document.querySelector(".hsk-panel"); if (b) b.remove(); if (p) p.remove(); }

  function toggle() {
    var p = document.querySelector(".hsk-panel");
    if (p) { p.remove(); return; }
    var cur = lsGet(K_SKIN) || "classique", e = readElig() || {};
    p = document.createElement("div");
    p.className = "hsk-panel"; p.setAttribute("role", "dialog"); p.setAttribute("aria-label", "Apparence");
    p.innerHTML = '<div class="hsk-head"><div class="hsk-title">🎨 Apparence</div><button type="button" class="hsk-x" aria-label="Fermer">×</button></div>'
      + '<div class="hsk-sub">' + (e.master ? "Compte maître" : "Débloqué au niveau 14 · Mythe") + " — le thème vous suit sur tous vos postes.</div>"
      + '<div class="hsk-grid">' + THEMES.map(function (t) {
        return '<button type="button" class="hsk-card' + (t.id === cur ? " on" : "") + '" data-id="' + t.id + '" style="--c1:' + t.c1 + ";--c2:" + t.c2 + ";--bg:" + t.bg + ";--glass:" + t.glass + ";--gb:" + t.gb + ";--ln:" + t.ln + '">'
          + '<div class="hsk-prev"><i></i><b></b><s></s></div><span class="hsk-check">✓</span>'
          + '<div class="hsk-meta"><div class="hsk-name">' + esc(t.nom) + (t.id === "nuit" ? " 🌙" : "") + '</div><div class="hsk-desc">' + esc(t.desc) + "</div></div></button>";
      }).join("") + "</div>"
      + '<div class="hsk-foot">Aperçu immédiat · enregistré automatiquement</div>';
    document.body.appendChild(p);
    p.querySelector(".hsk-x").addEventListener("click", function () { p.remove(); });
    p.addEventListener("click", function (ev) { ev.stopPropagation(); });
    p.querySelectorAll(".hsk-card").forEach(function (c) {
      c.addEventListener("click", function () { choisir(c.getAttribute("data-id")); });
    });
  }
  document.addEventListener("click", function () { var p = document.querySelector(".hsk-panel"); if (p) p.remove(); });
  document.addEventListener("keydown", function (ev) { if (ev.key === "Escape") { var p = document.querySelector(".hsk-panel"); if (p) p.remove(); } });

  var saveTimer = null;
  function choisir(id) {
    apply(id);
    clearTimeout(saveTimer);
    saveTimer = setTimeout(function () {
      post({ action: "saveApparence", theme: id }).then(function (r) {
        if (r && r.success === false) { alert(r.message || "Thème non enregistré."); check(true); }
      }).catch(function () {});
    }, 700);
  }

  window.HECG_SKINS = { apply: apply, recheck: function () { check(true); }, themes: THEMES };
  // Laisse la page charger ses propres données d'abord
  var go = function () { setTimeout(function () { check(false); }, 1200); };
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", go, { once: true }); else go();
  window.addEventListener("load", function () { apply(lsGet(K_SKIN) || "classique"); }); // Chart.js chargé après
})();
