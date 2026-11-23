/* HECG — Présence du personnel (nov. 2026).
   Chargé par hecg-skins.js dans chaque page connectée (personnel uniquement, jamais les étudiants).
   Toutes les 60 s (onglet visible) : signale la page / le module en cours (action getPresence, presence.gs)
   et récupère les collègues actifs depuis moins de 3 min → pastille « 👥 n en ligne » en bas à gauche. */
(function () {
  if (window.HECG_PRESENCE) return; window.HECG_PRESENCE = true;
  var API = "https://script.google.com/macros/s/AKfycbwfwxfVa2968AgKLdwwV8VIazVACVPVypKXrEUTc-zTuQ08gda8m5_qPfC2BgMyJMwb/exec";
  var PAGES = { dashboard: "Tableau de bord", dossier: "Dossier étudiant", partenaires: "Entreprises", dossier_partenaire: "Fiche entreprise", vie_scolaire: "Vie scolaire",
    prospects: "Prospects", offres: "Offres", statistiques: "Statistiques", parametres: "Paramètres", campagnes: "Campagnes", taches: "Tâches", historique: "Historique",
    trophees: "Trophées", salon: "Salon", positionnement: "Fiches de positionnement", cube: "Cube", qualiopi: "Qualiopi" };
  function user() { try { return JSON.parse(localStorage.getItem("user") || "{}") || {}; } catch (e) { return {}; } }
  var u = user();
  if (!u.sessionToken || u.role === "Etudiant") return;
  var esc = function (s) { return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]; }); };
  function txt(sel) { var e = document.querySelector(sel); return e ? String(e.innerText || e.textContent || "").replace(/\s+/g, " ").trim() : ""; }
  function ou() {
    var f = (location.pathname.split("/").pop() || "dashboard.html").replace(/\.html?$/, "") || "dashboard";
    var page = PAGES[f] || (document.title || f).replace(/\s*[-|—].*$/, "");
    var mod = "";
    if (document.getElementById("exOv")) mod = "Simulations d'examens";
    else mod = txt(".nav-btn.active") || txt(".tab-btn.active") || txt("[role=tab][aria-selected=true]");
    mod = mod.replace(/^[^\wÀ-ÿ]+/, "").slice(0, 50);
    var detail = f === "dossier" ? txt("#etuNom") : f === "dossier_partenaire" ? txt("#entNom") || txt("h1") : "";
    return { page: page, module: mod && mod !== page ? mod : "", detail: detail.slice(0, 60) };
  }
  var liste = [], ouvert = false, box = null;
  function rendre() {
    if (!box) {
      box = document.createElement("div"); box.id = "hecgPresence";
      box.style.cssText = "position:fixed;left:76px;bottom:18px;z-index:9980;font-family:'Plus Jakarta Sans',Inter,system-ui,sans-serif";
      document.body.appendChild(box);
    }
    var n = liste.length;
    if (!n) { box.innerHTML = ""; return; }
    var pastille = '<button onclick="window.__presBascule()" style="display:flex;align-items:center;gap:8px;height:40px;padding:0 14px;border-radius:999px;border:1px solid #dfe6f0;background:#fff;box-shadow:0 8px 22px -12px rgba(14,26,43,.45);font-weight:800;font-size:12.5px;color:#1A4E8A;cursor:pointer">'
      + '<span style="width:9px;height:9px;border-radius:50%;background:#10b981;box-shadow:0 0 0 3px #d1fae5"></span>👥 ' + n + ' collègue' + (n > 1 ? 's' : '') + ' en ligne</button>';
    var panneau = ouvert ? '<div style="position:absolute;bottom:48px;left:0;width:330px;max-height:60vh;overflow:auto;background:#fff;border:1px solid #e6ecf4;border-radius:16px;box-shadow:0 18px 40px -18px rgba(14,26,43,.5);padding:12px">'
      + '<div style="font-size:11px;font-weight:900;text-transform:uppercase;letter-spacing:.06em;color:#7d8a9c;margin:0 4px 8px">Actuellement connectés</div>'
      + liste.map(function (x) {
          var ini = String(x.nom || "?").split(/\s+/).map(function (w) { return w.charAt(0); }).join("").slice(0, 2).toUpperCase();
          var memeEndroit = x.page === ou().page && (!x.module || x.module === ou().module);
          return '<div style="display:flex;gap:10px;align-items:flex-start;padding:8px;border-radius:12px;' + (memeEndroit ? 'background:#fff7ed' : '') + '">'
            + '<div style="flex-shrink:0;width:34px;height:34px;border-radius:50%;background:#1A4E8A;color:#fff;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:12px">' + esc(ini) + '</div>'
            + '<div style="min-width:0"><div style="font-weight:800;font-size:13px;color:#0e1a2b">' + esc(x.nom) + (x.role ? ' <span style="font-weight:600;color:#94a3b8;font-size:11px">· ' + esc(x.role) + '</span>' : '') + '</div>'
            + '<div style="font-size:12px;color:#334155">' + esc(x.page) + (x.module ? ' › <b>' + esc(x.module) + '</b>' : '') + '</div>'
            + (x.detail ? '<div style="font-size:11.5px;color:#64748b">' + esc(x.detail) + '</div>' : '')
            + '<div style="font-size:10.5px;color:#94a3b8">' + (x.il < 60 ? "à l'instant" : "il y a " + Math.round(x.il / 60) + " min") + (memeEndroit ? ' · <b style="color:#c2410c">même écran que vous</b>' : '') + '</div></div></div>';
        }).join("") + '</div>' : "";
    box.innerHTML = '<div style="position:relative">' + panneau + pastille + '</div>';
  }
  window.__presBascule = function () { ouvert = !ouvert; rendre(); };
  document.addEventListener("click", function (e) { if (ouvert && box && !box.contains(e.target)) { ouvert = false; rendre(); } });
  var enCours = false;
  function signal() {
    if (enCours || document.visibilityState !== "visible") return;
    enCours = true;
    var o = ou(), us = user();
    fetch(API, { method: "POST", body: JSON.stringify({ action: "getPresence", page: o.page, module: o.module, detail: o.detail, sessionToken: us.sessionToken || "" }) })
      .then(function (r) { return r.json(); })
      .then(function (r) { if (r && r.success) { liste = r.autres || []; rendre(); } })
      .catch(function () {})
      .then(function () { enCours = false; });
  }
  // signal au chargement, puis toutes les 60 s, et à chaque changement de module (clic)
  var dernier = "";
  function siChange() { var o = ou(), k = o.page + "|" + o.module + "|" + o.detail; if (k !== dernier) { dernier = k; signal(); } }
  setTimeout(function () { dernier = ""; siChange(); }, 2500);
  setInterval(signal, 60000);
  document.addEventListener("click", function () { setTimeout(siChange, 1200); }, true);
  document.addEventListener("visibilitychange", function () { if (document.visibilityState === "visible") signal(); });
})();
