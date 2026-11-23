/* HECG — Présence du personnel + messagerie instantanée (nov. 2026).
   Chargé par hecg-skins.js dans chaque page connectée (personnel uniquement, jamais les étudiants).
   • Présence : toutes les 60 s (onglet visible) → action getPresence (presence.gs) ; pastille en bas à droite.
   • Messagerie : action getMessagerie (messagerie.gs), ops boite / fil / envoyer.
       - boîte toutes les 16 s (60 s onglet caché) : non lus, nouveaux messages → notification interne ;
       - fil toutes les 4 s quand une conversation est ouverte : historique + statuts Envoyé / Remis / Vu ;
       - le serveur relance par mail les messages non vus après 10 min. */
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
    var page = PAGES[f] || (document.title || f).replace(/^\(\d+\)\s*/, "").replace(/\s*[-|—].*$/, "");
    var mod = "";
    if (document.getElementById("exOv")) mod = "Simulations d'examens";
    else mod = txt(".nav-btn.active") || txt(".tab-btn.active") || txt("[role=tab][aria-selected=true]");
    mod = mod.replace(/^[^\wÀ-ÿ]+/, "").slice(0, 50);
    var detail = f === "dossier" ? txt("#etuNom") : f === "dossier_partenaire" ? txt("#entNom") || txt("h1") : "";
    return { page: page, module: mod && mod !== page ? mod : "", detail: detail.slice(0, 60) };
  }
  function post(o) { o.sessionToken = user().sessionToken || ""; return fetch(API, { method: "POST", body: JSON.stringify(o) }).then(function (r) { return r.json(); }); }
  function ini(n) { return String(n || "?").split(/\s+/).map(function (w) { return w.charAt(0); }).join("").slice(0, 2).toUpperCase(); }
  var J = ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."];
  function hm(d) { return ("0" + d.getHours()).slice(-2) + ":" + ("0" + d.getMinutes()).slice(-2); }
  function jour(d) {
    var a = new Date(); a.setHours(0, 0, 0, 0); var x = new Date(d); x.setHours(0, 0, 0, 0);
    var diff = Math.round((a - x) / 86400000);
    if (diff === 0) return "Aujourd'hui"; if (diff === 1) return "Hier";
    return J[x.getDay()] + " " + ("0" + x.getDate()).slice(-2) + "/" + ("0" + (x.getMonth() + 1)).slice(-2) + "/" + x.getFullYear();
  }
  function court(iso) { if (!iso) return ""; var d = new Date(iso); return jour(d) === "Aujourd'hui" ? hm(d) : jour(d).replace(/\/\d{4}$/, "") + " " + hm(d); }

  // ---------- état ----------
  var liste = [], convs = [], nonLus = 0, ouvert = false, box = null, chat = null; // chat = { email, nom, messages, charge }
  var CSS = "#hecgPresence{position:fixed;right:18px;bottom:18px;z-index:45;font-family:'Plus Jakarta Sans',Inter,system-ui,sans-serif}"
    + "#hecgPresence *{box-sizing:border-box}"
    + ".hp-pill{display:flex;align-items:center;gap:8px;height:40px;padding:0 14px;border-radius:999px;border:1px solid #dfe6f0;background:#fff;box-shadow:0 8px 22px -12px rgba(14,26,43,.45);font-weight:800;font-size:12.5px;color:#1A4E8A;cursor:pointer;position:relative;margin-left:auto}"
    + ".hp-dot{width:9px;height:9px;border-radius:50%}"
    + ".hp-badge{position:absolute;top:-7px;right:-6px;min-width:20px;height:20px;padding:0 6px;border-radius:999px;background:#ef4444;color:#fff;font-size:11px;font-weight:900;display:flex;align-items:center;justify-content:center;box-shadow:0 0 0 2px #fff;animation:hpPulse 1.6s infinite}"
    + "@keyframes hpPulse{0%,100%{transform:scale(1)}50%{transform:scale(1.12)}}"
    + ".hp-panel{position:absolute;bottom:48px;right:0;width:340px;max-height:66vh;overflow:auto;background:#fff;border:1px solid #e6ecf4;border-radius:16px;box-shadow:0 18px 40px -18px rgba(14,26,43,.5);padding:12px;color:#0f172a}"
    + ".hp-h{font-size:11px;font-weight:900;text-transform:uppercase;letter-spacing:.06em;color:#7d8a9c;margin:4px 4px 8px}"
    + ".hp-row{display:flex;gap:10px;align-items:flex-start;padding:8px;border-radius:12px;cursor:pointer;width:100%;border:0;background:transparent;text-align:left;font:inherit;color:inherit}"
    + ".hp-row:hover{background:#f1f5fb}"
    + ".hp-av{flex-shrink:0;width:34px;height:34px;border-radius:50%;background:#1A4E8A;color:#fff;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:12px;position:relative}"
    + ".hp-av i{position:absolute;right:-1px;bottom:-1px;width:10px;height:10px;border-radius:50%;background:#10b981;border:2px solid #fff}"
    + ".hp-nl{margin-left:auto;align-self:center;min-width:20px;height:20px;padding:0 6px;border-radius:999px;background:#ef4444;color:#fff;font-size:11px;font-weight:900;display:flex;align-items:center;justify-content:center}"
    + ".hp-chat{position:absolute;bottom:48px;right:0;width:370px;max-width:calc(100vw - 36px);height:min(540px,72vh);display:flex;flex-direction:column;background:#fff;border:1px solid #e6ecf4;border-radius:18px;box-shadow:0 24px 60px -20px rgba(14,26,43,.55);overflow:hidden;color:#0f172a}"
    + ".hp-ch{display:flex;align-items:center;gap:10px;padding:12px 14px;background:linear-gradient(135deg,#1A4E8A,#2563eb);color:#fff}"
    + ".hp-ch button{border:0;background:rgba(255,255,255,.16);color:#fff;width:30px;height:30px;border-radius:50%;cursor:pointer;font-size:15px}"
    + ".hp-fil{flex:1;overflow-y:auto;padding:12px 12px 4px;background:#f6f8fc}"
    + ".hp-j{text-align:center;margin:10px 0 8px}.hp-j span{font-size:10.5px;font-weight:800;color:#7d8a9c;background:#e8edf5;padding:3px 10px;border-radius:999px}"
    + ".hp-m{display:flex;margin:3px 0}.hp-m.moi{justify-content:flex-end}"
    + ".hp-b{max-width:78%;padding:8px 11px 5px;border-radius:14px;font-size:13px;line-height:1.4;word-wrap:break-word;white-space:pre-wrap;background:#fff;color:#0f172a;border:1px solid #e6ecf4;border-bottom-left-radius:4px}"
    + ".hp-m.moi .hp-b{background:#1A4E8A;color:#fff;border-color:#1A4E8A;border-bottom-left-radius:14px;border-bottom-right-radius:4px}"
    + ".hp-t{display:block;text-align:right;font-size:10px;opacity:.75;margin-top:3px;white-space:nowrap}"
    + ".hp-st{font-weight:800;margin-left:5px}.hp-st.vu{color:#7dd3fc}"
    + ".hp-sais{display:flex;gap:8px;padding:10px;border-top:1px solid #e6ecf4;background:#fff}"
    + ".hp-sais textarea{flex:1;resize:none;height:42px;border:1px solid #dfe6f0;border-radius:12px;padding:10px 12px;font:inherit;font-size:13px;outline:none;color:#0f172a;background:#fff}"
    + ".hp-sais button{border:0;border-radius:12px;padding:0 14px;background:#F26522;color:#fff;font-weight:800;cursor:pointer}"
    + ".hp-toast{position:fixed;right:18px;bottom:70px;z-index:10005;width:320px;display:flex;gap:10px;padding:12px 14px;border-radius:16px;background:#0f172a;color:#fff;box-shadow:0 20px 40px -14px rgba(0,0,0,.55);cursor:pointer;animation:hpIn .25s ease;font-family:'Plus Jakarta Sans',Inter,system-ui,sans-serif}"
    + "@keyframes hpIn{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:none}}"
    + "html[data-skin=nuit] .hp-pill,html[data-skin=nuit] .hp-panel,html[data-skin=nuit] .hp-chat,html[data-skin=nuit] .hp-sais,html[data-skin=nuit] .hp-sais textarea{background:#0f172a;color:#e2e8f0;border-color:rgba(148,163,184,.2)}"
    + "html[data-skin=nuit] .hp-fil{background:#0a0f1e}html[data-skin=nuit] .hp-b{background:#1e293b;color:#e2e8f0;border-color:#1e293b}html[data-skin=nuit] .hp-row:hover{background:rgba(255,255,255,.05)}";
  var st = document.createElement("style"); st.textContent = CSS; (document.head || document.documentElement).appendChild(st);

  function enLigne(email) { return liste.some(function (x) { return x.email === email; }); }
  function lien(email, nom) { return 'onclick="window.__presChat(\'' + esc(email) + '\',\'' + esc(String(nom).replace(/'/g, "’")) + '\')"'; }
  function rendre() {
    if (!document.body) return;
    if (!box) { box = document.createElement("div"); box.id = "hecgPresence"; box.setAttribute("data-sk-ignore", ""); document.body.appendChild(box); }
    var n = liste.length;
    var pill = '<button class="hp-pill" onclick="window.__presBascule()">'
      + (n ? '<span class="hp-dot" style="background:#10b981;box-shadow:0 0 0 3px #d1fae5"></span>👥 ' + n + ' collègue' + (n > 1 ? 's' : '') + ' en ligne'
           : '<span class="hp-dot" style="background:#94a3b8;box-shadow:0 0 0 3px #e2e8f0"></span><span style="color:#64748b">👤 Personne d’autre en ligne</span>')
      + (nonLus ? '<span class="hp-badge">' + nonLus + '</span>' : '') + '</button>';
    var corps = "";
    if (chat) corps = rendreChat();
    else if (ouvert) {
      var hors = convs.filter(function (c) { return !enLigne(c.email); });
      corps = '<div class="hp-panel"><div class="hp-h">En ligne · cliquez pour discuter</div>'
        + (n ? liste.map(function (x) {
            var c = convs.filter(function (k) { return k.email === x.email; })[0], nl = c ? c.nonLus : 0;
            var meme = x.page === ou().page && (!x.module || x.module === ou().module);
            return '<button class="hp-row" ' + lien(x.email, x.nom) + '>'
              + '<div class="hp-av">' + esc(ini(x.nom)) + '<i></i></div>'
              + '<div style="min-width:0;flex:1"><div style="font-weight:800;font-size:13px">' + esc(x.nom) + (x.role ? ' <span style="font-weight:600;color:#94a3b8;font-size:11px">· ' + esc(x.role) + '</span>' : '') + '</div>'
              + '<div style="font-size:12px;opacity:.8">' + esc(x.page) + (x.module ? ' › <b>' + esc(x.module) + '</b>' : '') + '</div>'
              + (x.detail ? '<div style="font-size:11.5px;color:#64748b">' + esc(x.detail) + '</div>' : '')
              + '<div style="font-size:10.5px;color:#94a3b8">' + (x.il < 60 ? "à l'instant" : "il y a " + Math.round(x.il / 60) + " min") + (meme ? ' · <b style="color:#c2410c">même écran que vous</b>' : '') + '</div></div>'
              + (nl ? '<span class="hp-nl">' + nl + '</span>' : '<span style="margin-left:auto;align-self:center;font-size:16px" title="Écrire">💬</span>') + '</button>';
          }).join("") : '<div style="padding:6px 8px 10px;font-size:12.5px;color:#64748b">Vous êtes seul(e) connecté(e) au CRM pour le moment.</div>')
        + (hors.length ? '<div class="hp-h" style="margin-top:10px">Conversations récentes</div>' + hors.map(function (c) {
            return '<button class="hp-row" ' + lien(c.email, c.nom) + '>'
              + '<div class="hp-av" style="background:#94a3b8">' + esc(ini(c.nom)) + '</div>'
              + '<div style="min-width:0;flex:1"><div style="font-weight:800;font-size:13px">' + esc(c.nom) + ' <span style="font-weight:600;color:#94a3b8;font-size:11px">· hors ligne</span></div>'
              + '<div style="font-size:12px;opacity:.75;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">' + (c.dernier && c.dernier.moi ? 'Vous : ' : '') + esc(c.dernier ? c.dernier.texte : '') + '</div>'
              + '<div style="font-size:10.5px;color:#94a3b8">' + esc(court(c.dernier && c.dernier.quand)) + '</div></div>'
              + (c.nonLus ? '<span class="hp-nl">' + c.nonLus + '</span>' : '') + '</button>';
          }).join("") : '')
        + '</div>';
    }
    var fil0 = box.querySelector(".hp-fil"), basAvant = !fil0 || fil0.scrollHeight - fil0.scrollTop - fil0.clientHeight < 40, posAvant = fil0 ? fil0.scrollTop : 0;
    var ta = box.querySelector(".hp-sais textarea"), brouillon = ta ? ta.value : null, focus = ta && document.activeElement === ta;
    box.innerHTML = '<div style="position:relative;display:flex;flex-direction:column">' + corps + pill + '</div>';
    var f2 = box.querySelector(".hp-fil"); if (f2) f2.scrollTop = basAvant ? f2.scrollHeight : posAvant;
    var t2 = box.querySelector(".hp-sais textarea");
    if (t2) {
      if (brouillon !== null) t2.value = brouillon;
      if (focus || brouillon === null) t2.focus();
      t2.addEventListener("keydown", function (ev) { if (ev.key === "Enter" && !ev.shiftKey) { ev.preventDefault(); envoyer(); } });
    }
    var titre = document.title.replace(/^\(\d+\)\s*/, "");
    document.title = nonLus ? "(" + nonLus + ") " + titre : titre;
  }
  function quand(iso) { var d = new Date(iso); return jour(d) + " à " + hm(d); }
  function statut(m) {
    if (m.pending) return '<span class="hp-st" title="Envoi en cours">🕓</span>';
    if (m.erreur) return '<span class="hp-st" style="color:#fecaca">⚠ non envoyé</span>';
    if (m.vu) return '<span class="hp-st vu" title="Vu le ' + esc(quand(m.vu)) + '">✓✓ Vu ' + hm(new Date(m.vu)) + '</span>';
    if (m.remis) return '<span class="hp-st" title="Remis le ' + esc(quand(m.remis)) + '">✓✓ Remis</span>';
    return '<span class="hp-st" title="Envoyé le ' + esc(quand(m.envoye)) + '">✓ Envoyé</span>';
  }
  function rendreChat() {
    var c = chat, h = "", dj = "";
    (c.messages || []).forEach(function (m) {
      var d = new Date(m.envoye), j = jour(d);
      if (j !== dj) { dj = j; h += '<div class="hp-j"><span>' + esc(j) + '</span></div>'; }
      h += '<div class="hp-m' + (m.moi ? ' moi' : '') + '"><div class="hp-b" title="' + esc(quand(m.envoye)) + '">' + esc(m.texte) + '<span class="hp-t">' + hm(d) + (m.moi ? statut(m) : '') + '</span></div></div>';
    });
    if (!c.charge) h = '<div style="text-align:center;color:#94a3b8;font-size:12px;margin-top:30px">Chargement de l’historique…</div>';
    else if (!(c.messages || []).length) h = '<div style="text-align:center;color:#94a3b8;font-size:12.5px;margin-top:30px">Aucun échange pour l’instant.<br>Écrivez le premier message 👋</div>';
    return '<div class="hp-chat"><div class="hp-ch"><button onclick="window.__presRetour()" title="Retour à la liste">‹</button>'
      + '<div class="hp-av" style="background:rgba(255,255,255,.2)">' + esc(ini(c.nom)) + (enLigne(c.email) ? '<i></i>' : '') + '</div>'
      + '<div style="flex:1;min-width:0"><div style="font-weight:800;font-size:14px">' + esc(c.nom) + '</div><div style="font-size:11px;opacity:.85">' + (enLigne(c.email) ? '● en ligne' : 'hors ligne · relance par mail si non lu après 10 min') + '</div></div>'
      + '<button onclick="window.__presFermer()" title="Fermer">×</button></div>'
      + '<div class="hp-fil">' + h + '</div>'
      + '<div class="hp-sais"><textarea placeholder="Votre message… (Entrée = envoyer, Maj+Entrée = retour à la ligne)" maxlength="2000"></textarea><button onclick="window.__presEnvoyer()">Envoyer</button></div></div>';
  }

  // ---------- actions ----------
  window.__presBascule = function () { if (chat) { chat = null; ouvert = false; } else ouvert = !ouvert; rendre(); if (ouvert) boite(); };
  window.__presRetour = function () { chat = null; ouvert = true; rendre(); };
  window.__presFermer = function () { chat = null; ouvert = false; rendre(); };
  window.__presChat = function (email, nom) { chat = { email: String(email).toLowerCase(), nom: nom, messages: [], charge: false }; rendre(); fil(); };
  window.__presEnvoyer = function () { envoyer(); };
  function envoyer() {
    if (!chat || !box) return;
    var ta = box.querySelector(".hp-sais textarea"), t = ta ? ta.value.trim() : "";
    if (!t) return;
    ta.value = "";
    var c = chat, tmp = { texte: t, envoye: new Date().toISOString(), moi: true, pending: true };
    c.messages.push(tmp); c.charge = true; rendre();
    post({ action: "getMessagerie", op: "envoyer", a: c.email, aNom: c.nom, texte: t }).then(function (r) {
      tmp.pending = false;
      if (r && r.success) { var i = c.messages.indexOf(tmp); if (i > -1) c.messages[i] = r.message; }
      else { tmp.erreur = true; }
      if (chat === c) rendre();
    }).catch(function () { tmp.pending = false; tmp.erreur = true; if (chat === c) rendre(); });
  }
  var filEnCours = false;
  function fil() {
    if (!chat || filEnCours) return;
    var c = chat; filEnCours = true;
    post({ action: "getMessagerie", op: "fil", avec: c.email, vu: document.visibilityState === "visible" }).then(function (r) {
      if (!r || !r.success || chat !== c) return;
      var attente = c.messages.filter(function (m) { return m.pending || m.erreur; });
      var avant = JSON.stringify(c.messages.map(function (m) { return [m.id, m.remis, m.vu]; }));
      c.messages = (r.messages || []).concat(attente);
      var apres = JSON.stringify(c.messages.map(function (m) { return [m.id, m.remis, m.vu]; }));
      convs.forEach(function (k) { if (k.email === c.email) k.nonLus = 0; });
      nonLus = convs.reduce(function (t, k) { return t + k.nonLus; }, 0);
      if (avant !== apres || !c.charge) { c.charge = true; rendre(); }
    }).catch(function () {}).then(function () { filEnCours = false; });
  }
  var vus = {}, premiere = true;
  function boite() {
    return post({ action: "getMessagerie", op: "boite" }).then(function (r) {
      if (!r || !r.success) return;
      convs = r.conversations || [];
      if (chat) convs.forEach(function (k) { if (k.email === chat.email) k.nonLus = 0; });
      nonLus = convs.reduce(function (t, k) { return t + k.nonLus; }, 0);
      (r.nouveaux || []).forEach(function (m) {
        if (vus[m.id]) return; vus[m.id] = 1;
        if (chat && chat.email === m.de) { fil(); return; } // déjà sous les yeux
        notifier(m);
      });
      // à l'ouverture d'une page : rappel des non-lus déjà remis auparavant
      if (premiere && nonLus && !(r.nouveaux || []).length) {
        var c0 = convs.filter(function (k) { return k.nonLus; })[0];
        if (c0) notifier({ de: c0.email, deNom: c0.nom, texte: nonLus + " message" + (nonLus > 1 ? "s" : "") + " non lu" + (nonLus > 1 ? "s" : "") + (c0.dernier ? " — « " + c0.dernier.texte + " »" : ""), envoye: c0.dernier ? c0.dernier.quand : new Date().toISOString() }, true);
      }
      premiere = false;
      rendre();
    }).catch(function () {});
  }
  // notification interne : bulle cliquable + son discret
  function notifier(m, silencieux) {
    if (!document.body) return;
    document.querySelectorAll(".hp-toast").forEach(function (x) { x.remove(); });
    var t = document.createElement("div"); t.className = "hp-toast"; t.setAttribute("data-sk-ignore", "");
    t.innerHTML = '<div class="hp-av" style="background:#F26522">' + esc(ini(m.deNom)) + '</div><div style="min-width:0;flex:1"><div style="font-weight:800;font-size:13px">💬 ' + esc(m.deNom)
      + ' <span style="font-weight:600;opacity:.6;font-size:11px">' + esc(court(m.envoye)) + '</span></div><div style="font-size:12.5px;opacity:.9;max-height:54px;overflow:hidden">' + esc(m.texte) + '</div>'
      + '<div style="font-size:10.5px;opacity:.55;margin-top:3px">Cliquez pour répondre</div></div>';
    t.onclick = function (ev) { ev.stopPropagation(); t.remove(); window.__presChat(m.de, m.deNom); };
    document.body.appendChild(t);
    setTimeout(function () { if (t.isConnected) t.remove(); }, 10000);
    if (silencieux) return;
    try {
      var A = window.AudioContext || window.webkitAudioContext; if (!A) return;
      var ctx = new A(), o = ctx.createOscillator(), g = ctx.createGain();
      o.type = "sine"; o.frequency.setValueAtTime(880, ctx.currentTime); o.frequency.setValueAtTime(1320, ctx.currentTime + .09);
      g.gain.setValueAtTime(.0001, ctx.currentTime); g.gain.exponentialRampToValueAtTime(.12, ctx.currentTime + .02); g.gain.exponentialRampToValueAtTime(.0001, ctx.currentTime + .3);
      o.connect(g); g.connect(ctx.destination); o.start(); o.stop(ctx.currentTime + .32);
    } catch (e) {}
  }
  document.addEventListener("click", function (e) {
    if ((ouvert || chat) && box && e.target.isConnected && !box.contains(e.target) && !(e.target.closest && e.target.closest(".hp-toast"))) { ouvert = false; chat = null; rendre(); }
  });

  // ---------- présence ----------
  var enCours = false;
  function signal() {
    if (enCours || document.visibilityState !== "visible") return;
    enCours = true;
    var o = ou();
    post({ action: "getPresence", page: o.page, module: o.module, detail: o.detail })
      .then(function (r) { if (r && r.success) { liste = r.autres || []; rendre(); } })
      .catch(function () {})
      .then(function () { enCours = false; });
  }
  var dernier = "";
  function siChange() { var o = ou(), k = o.page + "|" + o.module + "|" + o.detail; if (k !== dernier) { dernier = k; signal(); } }
  setTimeout(function () { dernier = ""; siChange(); boite(); }, 2500);
  setInterval(signal, 60000);
  var tick = 0;
  setInterval(function () {
    tick++;
    var visible = document.visibilityState === "visible";
    if (chat && visible) fil();
    if (tick % (visible ? 4 : 15) === 0) boite();
  }, 4000);
  document.addEventListener("click", function () { setTimeout(siChange, 1200); }, true);
  document.addEventListener("visibilitychange", function () { if (document.visibilityState === "visible") { signal(); boite(); if (chat) fil(); } });
})();
