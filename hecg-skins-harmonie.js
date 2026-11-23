/* HECG — Harmonisation des thèmes (nov. 2026). Chargé par hecg-skins.js.
   Les classes Tailwind sont gérées par hecg-skins.css, mais beaucoup de fonds blancs / textes sombres
   sont écrits en dur (styles des modules, HTML généré en JS) et restaient blancs en « Nuit néon ».
   Ce module regarde les couleurs RÉELLES calculées par le navigateur et pose des marqueurs :
     data-sk-surf="g"  fond blanc / gris clair      → verre du thème (sombre en Nuit néon)
     data-sk-surf="c"  fond pastel coloré           → teinte translucide (Nuit néon seulement)
     data-sk-line      bordure claire               → liseré discret (Nuit néon)
     data-sk-ink       texte sombre sur fond sombre → texte clair (Nuit néon)
   Le rendu est défini dans hecg-skins.css (section « Harmonisation »). Traitement par tranches
   (pas de gel de la page) + observation des contenus ajoutés ensuite. */
(function () {
  if (window.HECG_HARMONIE) return; window.HECG_HARMONIE = true;
  var root = document.documentElement;
  var SKIP = { SCRIPT: 1, STYLE: 1, LINK: 1, META: 1, IMG: 1, SVG: 1, PATH: 1, CANVAS: 1, VIDEO: 1, IFRAME: 1, INPUT: 1, SELECT: 1, TEXTAREA: 1, OPTION: 1, BR: 1, HR: 1, NOSCRIPT: 1, TEMPLATE: 1 };
  var file = [], enFile = new WeakSet(), planifie = false, obs = null;

  function skin() { return root.getAttribute("data-skin") || ""; }
  function rgba(s) {
    var m = /rgba?\(([^)]+)\)/.exec(s || ""); if (!m) return null;
    var p = m[1].split(/[\s,\/]+/).filter(Boolean).map(parseFloat);
    return { r: p[0], g: p[1], b: p[2], a: p.length > 3 ? p[3] : 1 };
  }
  function lum(c) { return (0.2126 * c.r + 0.7152 * c.g + 0.0722 * c.b) / 255; }
  function sat(c) { var mx = Math.max(c.r, c.g, c.b), mn = Math.min(c.r, c.g, c.b); return mx ? (mx - mn) / mx : 0; }
  function exclu(el) { return !!(el.closest && el.closest(".hsk-panel,.hsk-fab,.no-theme,[data-sk-ignore]")); }

  // Fond effectif (après thème) : premier ancêtre au fond opaque
  function fondEffectif(el) {
    for (var n = el; n && n.nodeType === 1; n = n.parentElement) {
      var c = rgba(getComputedStyle(n).backgroundColor);
      if (c && c.a > 0.45) return c;
      if (n === document.body) break;
    }
    return skin() === "nuit" ? { r: 10, g: 14, b: 30, a: 1 } : { r: 245, g: 247, b: 251, a: 1 };
  }
  function aTexteDirect(el) {
    for (var n = el.firstChild; n; n = n.nextSibling) if (n.nodeType === 3 && /\S/.test(n.nodeValue)) return true;
    return false;
  }

  function surfaces(el) {
    if (el.hasAttribute("data-sk-surf")) return;
    var cs = getComputedStyle(el), c = rgba(cs.backgroundColor);
    var petit = el.offsetWidth > 0 && el.offsetWidth < 26 && el.offsetHeight < 26; // curseurs d'interrupteurs, puces…
    if (c && c.a > 0.85 && lum(c) > 0.86 && !petit) {
      if (sat(c) < 0.06) el.setAttribute("data-sk-surf", "g");
      else { el.style.setProperty("--sk-sbg", Math.round(c.r) + "," + Math.round(c.g) + "," + Math.round(c.b)); el.setAttribute("data-sk-surf", "c"); }
    }
    if (skin() === "nuit" && !el.hasAttribute("data-sk-line")) {
      var w = parseFloat(cs.borderTopWidth) || parseFloat(cs.borderLeftWidth) || parseFloat(cs.borderBottomWidth) || 0;
      if (w > 0) {
        var b = rgba(cs.borderTopColor && parseFloat(cs.borderTopWidth) ? cs.borderTopColor : (parseFloat(cs.borderLeftWidth) ? cs.borderLeftColor : cs.borderBottomColor));
        if (b && b.a > 0.5 && lum(b) > 0.8 && sat(b) < 0.25) el.setAttribute("data-sk-line", "");
      }
    }
  }
  function encre(el) {
    if (skin() !== "nuit" || el.hasAttribute("data-sk-ink") || !aTexteDirect(el)) return;
    var c = rgba(getComputedStyle(el).color);
    if (!c || lum(c) > 0.5) return;
    if (lum(fondEffectif(el)) > 0.5) return; // vrai fond clair conservé (badge, bouton) : on n'y touche pas
    if (sat(c) > 0.3) el.style.setProperty("--sk-ink", Math.round(c.r) + "," + Math.round(c.g) + "," + Math.round(c.b));
    el.setAttribute("data-sk-ink", sat(c) > 0.3 ? "c" : "g");
  }

  function enfiler(el) {
    if (!el || el.nodeType !== 1 || SKIP[el.tagName.toUpperCase()] || enFile.has(el)) return;
    if (el instanceof SVGElement) return;
    enFile.add(el); file.push(el);
  }
  function enfilerArbre(el) {
    if (!el || el.nodeType !== 1) return;
    enfiler(el);
    if (el.querySelectorAll) { var t = el.querySelectorAll("*"); for (var i = 0; i < t.length; i++) enfiler(t[i]); }
    planifier();
  }
  function planifier() { if (!planifie && file.length) { planifie = true; setTimeout(tranche, 30); } }
  // Deux passes par tranche : fonds d'abord (le texte dépend du fond final), puis textes.
  function tranche() {
    planifie = false;
    if (!skin()) { file = []; enFile = new WeakSet(); return; }
    var t0 = performance.now(), lot = [];
    while (file.length && performance.now() - t0 < 10) {
      var el = file.shift(); enFile.delete(el);
      if (!el.isConnected || exclu(el)) continue;
      try { surfaces(el); lot.push(el); } catch (e) {}
    }
    for (var i = 0; i < lot.length; i++) { try { encre(lot[i]); } catch (e) {} }
    planifier();
  }

  function demarrer() {
    if (!document.body) { document.addEventListener("DOMContentLoaded", demarrer, { once: true }); return; }
    if (!skin()) { if (obs) { obs.disconnect(); obs = null; } return; }
    enfilerArbre(document.body);
    if (!obs) {
      obs = new MutationObserver(function (muts) {
        for (var i = 0; i < muts.length; i++) {
          var a = muts[i].addedNodes;
          for (var j = 0; j < a.length; j++) if (a[j].nodeType === 1) enfilerArbre(a[j]);
          if (a.length && muts[i].target.nodeType === 1) { enfiler(muts[i].target); planifier(); } // nouveau texte dans un élément existant
        }
      });
      obs.observe(document.body, { childList: true, subtree: true });
    }
  }
  // Changement de thème : on relance (l'encre ne s'évalue qu'en Nuit néon)
  document.addEventListener("hecg-skin", function () { setTimeout(demarrer, 50); });
  demarrer();
})();
