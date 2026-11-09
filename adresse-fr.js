/* HECG — Adressage normalisé des entreprises + postes des contacts (novembre 2026).
   Référentiels officiels, gratuits, sans clé :
   • Base Adresse Nationale (BAN) — https://data.geopf.fr/geocodage (autocomplétion)
   • Répertoire SIRENE de l'INSEE — https://recherche-entreprises.api.gouv.fr
   API :
     HecgAdresse.ouvrir({ titre, nomEntreprise, initial, onValider(a) })  → fenêtre d'adresse
       a = { adresse, complement, cp, ville, lat, lon, statut, siret }
       statut : "BAN" (au numéro) | "BAN voie" | "SIRENE" | "Commune" | "Saisie libre" | ""
     HecgAdresse.texte(a) / HecgAdresse.badge(a) / HecgAdresse.estVerifiee(a)
     HecgPostes.attacher(input) → liste déroulante des postes (existants + types), saisie libre possible
   Serveur : adresses.gs (normalisation, une seule adresse par entreprise). */
(function () {
  if (window.HecgAdresse) return;
  var BAN = "https://data.geopf.fr/geocodage/search";
  var SIRENE = "https://recherche-entreprises.api.gouv.fr/search";
  var OK = { "BAN": 1, "BAN voie": 1, "SIRENE": 1 };

  function esc(s) { return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]; }); }
  function casse(s) {
    var petits = { de: 1, du: 1, des: 1, la: 1, le: 1, les: 1, et: 1, "à": 1, au: 1, aux: 1, sur: 1, en: 1 };
    return String(s || "").toLowerCase().split(/(\s+|-)/).map(function (w, i) {
      if (!w.trim() || w === "-") return w;
      if (i > 0 && petits[w]) return w;
      var m = w.match(/^([dl])['’](.+)$/);
      if (m) return (i > 0 ? m[1] : m[1].toUpperCase()) + "'" + m[2].charAt(0).toUpperCase() + m[2].slice(1);
      return w.charAt(0).toUpperCase() + w.slice(1);
    }).join("").replace(/\s+/g, " ").trim();
  }
  function texte(a) {
    if (!a) return "";
    return [a.adresse, a.complement, [a.cp, a.ville].filter(Boolean).join(" ")].filter(function (x) { return String(x || "").trim(); }).join(", ");
  }
  function estVerifiee(a) { return !!(a && OK[a.statut]); }
  function badge(a) {
    if (!a || (!a.adresse && !a.cp && !a.ville)) return '<span class="text-[9px] font-black px-2 py-0.5 rounded-full bg-slate-100 text-slate-400">📍 Adresse manquante</span>';
    if (a.statut === "SIRENE") return '<span class="text-[9px] font-black px-2 py-0.5 rounded-full bg-emerald-100 text-emerald-700" title="Adresse officielle de l\'établissement (répertoire SIRENE de l\'INSEE)">✅ SIRENE</span>';
    if (a.statut === "BAN") return '<span class="text-[9px] font-black px-2 py-0.5 rounded-full bg-emerald-100 text-emerald-700" title="Adresse référencée dans la Base Adresse Nationale">✅ Adresse référencée</span>';
    if (a.statut === "BAN voie") return '<span class="text-[9px] font-black px-2 py-0.5 rounded-full bg-emerald-50 text-emerald-600" title="Voie référencée (numéro absent de la Base Adresse Nationale)">✅ Voie référencée</span>';
    if (a.statut === "Adresses multiples") return '<span class="text-[9px] font-black px-2 py-0.5 rounded-full bg-red-100 text-red-700" title="Les contacts de cette entreprise ont des adresses différentes : choisissez la bonne (✏️)">⛔ Adresses différentes selon les contacts</span>';
    if (a.statut === "Commune") return '<span class="text-[9px] font-black px-2 py-0.5 rounded-full bg-amber-100 text-amber-700" title="Seule la commune est connue : position approximative sur la carte">⚠️ Rue manquante</span>';
    return '<span class="text-[9px] font-black px-2 py-0.5 rounded-full bg-amber-100 text-amber-700" title="Adresse non trouvée dans la Base Adresse Nationale">⚠️ Adresse à vérifier</span>';
  }

  // ---------- Leaflet (aperçu carte) chargé à la demande ----------
  var leafletP = null;
  function leaflet() {
    if (window.L && L.map) return Promise.resolve();
    if (leafletP) return leafletP;
    leafletP = new Promise(function (ok) {
      var l = document.createElement("link"); l.rel = "stylesheet"; l.href = "https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"; document.head.appendChild(l);
      var s = document.createElement("script"); s.src = "https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"; s.onload = function () { ok(); }; s.onerror = function () { ok(); };
      document.head.appendChild(s);
    });
    return leafletP;
  }

  function banVersAdresse(f) {
    var p = f.properties || {}, c = (f.geometry || {}).coordinates || [];
    var voie = casse(p.street || p.name || "");
    var statut = p.type === "housenumber" ? "BAN" : (p.type === "street" || p.type === "locality") ? "BAN voie" : "Commune";
    return {
      adresse: p.type === "municipality" ? "" : (p.housenumber ? p.housenumber + " " : "") + voie,
      cp: p.postcode || "", ville: p.city || "", lat: c[1], lon: c[0], statut: statut
    };
  }

  // ---------- Fenêtre d'adresse ----------
  function ouvrir(opts) {
    opts = opts || {};
    var a = Object.assign({ adresse: "", complement: "", cp: "", ville: "", lat: "", lon: "", statut: "", siret: "" }, opts.initial || {});
    var old = document.getElementById("hadrModal"); if (old) old.remove();
    var m = document.createElement("div");
    m.id = "hadrModal";
    m.className = "fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4";
    m.style.zIndex = "200";
    var inp = "w-full p-3 bg-slate-50 border border-slate-200 rounded-xl text-sm outline-none focus:ring-2 focus:ring-[#1A4E8A]";
    m.innerHTML =
      '<div class="bg-white rounded-[2rem] shadow-2xl w-full max-w-2xl max-h-[92vh] overflow-y-auto p-6 md:p-8 border border-slate-100">'
      + '<div class="flex justify-between items-start mb-5"><div><h2 class="text-xl font-black text-slate-800">📍 ' + esc(opts.titre || "Adresse de l'entreprise") + '</h2>'
      + '<p class="text-[11px] text-slate-400 font-medium mt-1">Une entreprise n\'a qu\'une adresse, commune à tous ses contacts. Choisissez-la dans les référentiels officiels.</p></div>'
      + '<button type="button" data-x class="text-slate-300 hover:text-red-500 text-2xl leading-none">✕</button></div>'
      // SIRENE
      + '<div class="bg-blue-50/60 border border-blue-100 rounded-2xl p-4 mb-5">'
      + '<label class="text-[10px] font-black uppercase text-[#1A4E8A]">🏛️ Option 1 — Retrouver l\'entreprise dans le répertoire SIRENE (INSEE)</label>'
      + '<div class="flex gap-2 mt-2"><input type="text" data-sq class="' + inp + ' bg-white" placeholder="Nom de l\'entreprise, SIREN ou SIRET…">'
      + '<button type="button" data-sgo class="px-4 bg-[#1A4E8A] text-white rounded-xl text-xs font-black whitespace-nowrap">Rechercher</button></div>'
      + '<div data-sres class="mt-2 space-y-1 max-h-56 overflow-y-auto"></div></div>'
      // BAN
      + '<label class="text-[10px] font-black uppercase text-slate-500">🧭 Option 2 — Numéro et voie (Base Adresse Nationale)</label>'
      + '<div class="relative mt-1"><input type="text" data-adr autocomplete="off" class="' + inp + '" placeholder="Ex. : 146 rue Paradis Marseille">'
      + '<div data-sugg class="hidden absolute left-0 right-0 mt-1 bg-white border border-slate-200 rounded-xl shadow-xl overflow-hidden" style="z-index:5"></div></div>'
      + '<p class="text-[10px] text-slate-400 mt-1 ml-1">Tapez le début de l\'adresse puis choisissez une proposition.</p>'
      + '<div class="grid grid-cols-1 md:grid-cols-3 gap-3 mt-4">'
      + '<div class="md:col-span-3"><label class="text-[10px] font-black uppercase text-slate-400 ml-1">Complément (bâtiment, étage, résidence, BP…)</label><input type="text" data-cpl class="' + inp + '" placeholder="Ex. : Immeuble Le Ponant, Bât. B"></div>'
      + '<div><label class="text-[10px] font-black uppercase text-slate-400 ml-1">Code postal</label><input type="text" data-cp maxlength="5" class="' + inp + '" readonly></div>'
      + '<div class="md:col-span-2"><label class="text-[10px] font-black uppercase text-slate-400 ml-1">Ville</label><input type="text" data-ville class="' + inp + '" readonly></div></div>'
      + '<div class="flex items-center justify-between gap-3 mt-3 flex-wrap"><div data-statut></div>'
      + '<button type="button" data-libre class="text-[10px] font-bold text-slate-400 hover:text-[#F26522] underline">Adresse introuvable ? Saisie libre</button></div>'
      + '<div data-map class="mt-4 rounded-2xl overflow-hidden border border-slate-200 hidden" style="height:200px"></div>'
      + '<div class="flex justify-between items-center gap-3 mt-6">'
      + '<button type="button" data-vider class="text-[11px] font-bold text-slate-300 hover:text-red-500">Effacer l\'adresse</button>'
      + '<div class="flex gap-3"><button type="button" data-x class="px-6 py-3 text-slate-400 font-bold hover:bg-slate-100 rounded-2xl">Annuler</button>'
      + '<button type="button" data-ok class="px-8 py-3 bg-[#F26522] text-white rounded-2xl font-black shadow-lg">Valider l\'adresse</button></div></div>'
      + '</div>';
    document.body.appendChild(m);
    function q(sel) { return m.querySelector(sel); }
    var eAdr = q("[data-adr]"), eCpl = q("[data-cpl]"), eCp = q("[data-cp]"), eVille = q("[data-ville]"), eSugg = q("[data-sugg]");
    var carte = null, marqueur = null;

    function remplir() {
      eAdr.value = a.adresse || ""; eCpl.value = a.complement || ""; eCp.value = a.cp || ""; eVille.value = a.ville || "";
      var libre = a.statut === "Saisie libre" || (!a.statut && (a.adresse || a.cp || a.ville));
      eCp.readOnly = eVille.readOnly = !libre;
      eCp.classList.toggle("bg-white", libre); eVille.classList.toggle("bg-white", libre);
      q("[data-statut]").innerHTML = badge(a) + (a.siret ? ' <span class="text-[9px] font-bold text-slate-400 ml-1">SIRET ' + esc(a.siret) + "</span>" : "");
      apercu();
    }
    function apercu() {
      var box = q("[data-map]");
      if (!a.lat || !a.lon) { box.classList.add("hidden"); return; }
      box.classList.remove("hidden");
      leaflet().then(function () {
        if (!window.L) return;
        var ll = [Number(a.lat), Number(a.lon)], z = a.statut === "Commune" ? 12 : 17;
        if (!carte) {
          carte = L.map(box, { zoomControl: true, attributionControl: false }).setView(ll, z);
          L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", { maxZoom: 19 }).addTo(carte);
          marqueur = L.marker(ll).addTo(carte);
        } else { carte.setView(ll, z); marqueur.setLatLng(ll); }
        setTimeout(function () { carte.invalidateSize(); }, 80);
      });
    }
    function fermer() { m.remove(); }

    // --- Autocomplétion BAN
    var tSugg = null, nReq = 0, items = [];
    eAdr.addEventListener("input", function () {
      a.statut = a.statut === "SIRENE" ? "" : a.statut;
      if (OK[a.statut] || a.statut === "Commune") { a.statut = ""; a.lat = ""; a.lon = ""; }
      a.adresse = eAdr.value;
      q("[data-statut]").innerHTML = badge(Object.assign({}, a, { statut: "" }));
      clearTimeout(tSugg);
      var v = eAdr.value.trim();
      if (v.length < 3) { eSugg.classList.add("hidden"); return; }
      tSugg = setTimeout(function () {
        var id = ++nReq;
        fetch(BAN + "?autocomplete=1&limit=6&lat=43.4&lon=5.4&q=" + encodeURIComponent(v.slice(0, 200)))
          .then(function (r) { return r.json(); }).then(function (j) {
            if (id !== nReq) return;
            items = (j.features || []);
            if (!items.length) { eSugg.innerHTML = '<div class="px-4 py-3 text-xs text-slate-400 italic">Aucune adresse trouvée — vérifiez l\'orthographe ou utilisez la saisie libre.</div>'; eSugg.classList.remove("hidden"); return; }
            eSugg.innerHTML = items.map(function (f, i) {
              var p = f.properties || {};
              var t = p.type === "housenumber" ? "" : p.type === "street" ? " · voie" : p.type === "municipality" ? " · commune" : " · lieu-dit";
              return '<button type="button" data-i="' + i + '" class="block w-full text-left px-4 py-2.5 hover:bg-blue-50 border-b border-slate-100 last:border-0">'
                + '<div class="text-sm font-bold text-slate-700">' + esc(casse(p.name)) + '</div>'
                + '<div class="text-[10px] text-slate-400">' + esc((p.postcode || "") + " " + (p.city || "") + " — " + (p.context || "")) + esc(t) + "</div></button>";
            }).join("");
            eSugg.classList.remove("hidden");
          }).catch(function () {});
      }, 220);
    });
    eSugg.addEventListener("mousedown", function (ev) {
      var b = ev.target.closest("[data-i]"); if (!b) return;
      ev.preventDefault();
      var n = banVersAdresse(items[+b.getAttribute("data-i")]);
      a.adresse = n.adresse; a.cp = n.cp; a.ville = n.ville; a.lat = n.lat; a.lon = n.lon; a.statut = n.statut;
      eSugg.classList.add("hidden");
      remplir();
      if (n.statut === "BAN voie") eAdr.focus();
    });
    eAdr.addEventListener("blur", function () { setTimeout(function () { eSugg.classList.add("hidden"); }, 150); });

    // --- SIRENE
    var etabs = [];
    function chercherSirene() {
      var v = q("[data-sq]").value.trim(), box = q("[data-sres]");
      if (v.length < 2) return;
      box.innerHTML = '<p class="text-xs text-slate-400 italic px-1">Recherche dans le répertoire SIRENE…</p>';
      var siret = v.replace(/\s/g, "");
      fetch(SIRENE + "?per_page=8&etat_administratif=A&q=" + encodeURIComponent(siret.match(/^\d{9,14}$/) ? siret.slice(0, 9) : v))
        .then(function (r) { return r.json(); }).then(function (j) {
          etabs = [];
          (j.results || []).forEach(function (r) {
            var liste = (r.matching_etablissements && r.matching_etablissements.length ? r.matching_etablissements : [r.siege]).filter(function (e) { return e && e.etat_administratif !== "F"; });
            if (/^\d{14}$/.test(siret)) {   // SIRET précis : cet établissement seulement
              var exact = liste.concat(r.siege ? [r.siege] : []).filter(function (e) { return e.siret === siret; });
              if (exact.length) liste = exact.slice(0, 1);
            }
            liste.slice(0, 4).forEach(function (e) { etabs.push({ nom: r.nom_complet, naf: r.activite_principale, e: e }); });
          });
          if (!etabs.length) { box.innerHTML = '<p class="text-xs text-slate-400 italic px-1">Aucun établissement actif trouvé.</p>'; return; }
          box.innerHTML = etabs.map(function (x, i) {
            var e = x.e;
            return '<button type="button" data-e="' + i + '" class="block w-full text-left px-3 py-2 bg-white rounded-xl border border-blue-100 hover:border-[#1A4E8A] hover:shadow-sm">'
              + '<div class="flex items-center gap-2 flex-wrap"><span class="text-xs font-black text-slate-700">' + esc(x.nom) + '</span>'
              + (e.est_siege ? '<span class="text-[9px] font-black bg-blue-100 text-blue-700 px-2 rounded-full">Siège</span>' : '<span class="text-[9px] font-black bg-slate-100 text-slate-500 px-2 rounded-full">Établissement</span>')
              + '</div><div class="text-[10px] text-slate-500">' + esc(e.adresse || "") + '</div><div class="text-[9px] text-slate-400">SIRET ' + esc(e.siret || "") + (x.naf ? " · NAF " + esc(x.naf) : "") + "</div></button>";
          }).join("");
        }).catch(function () { box.innerHTML = '<p class="text-xs text-red-400 italic px-1">Répertoire SIRENE indisponible, réessayez ou passez par l\'option 2.</p>'; });
    }
    q("[data-sgo]").addEventListener("click", chercherSirene);
    q("[data-sq]").addEventListener("keydown", function (ev) { if (ev.key === "Enter") { ev.preventDefault(); chercherSirene(); } });
    q("[data-sres]").addEventListener("click", function (ev) {
      var b = ev.target.closest("[data-e]"); if (!b) return;
      var e = etabs[+b.getAttribute("data-e")].e;
      var voie = [e.numero_voie, e.indice_repetition, e.type_voie, e.libelle_voie].filter(Boolean).join(" ");
      if (!voie && e.adresse) {   // établissements secondaires : adresse en un bloc « N VOIE CP VILLE »
        voie = String(e.adresse).replace(/\s*\b\d{5}\b.*$/, "").trim();
        if (e.complement_adresse) voie = voie.replace(e.complement_adresse, "").trim();
      }
      var cand = { adresse: casse(voie), complement: e.complement_adresse ? casse(e.complement_adresse) : "", cp: e.code_postal || "", ville: casse(e.libelle_commune || ""),
        lat: Number(e.latitude) || "", lon: Number(e.longitude) || "", statut: "SIRENE", siret: e.siret || "" };
      // Mise en forme BAN de l'adresse SIRENE (abréviations, casse) quand elle y figure au numéro
      fetch(BAN + "?limit=1&q=" + encodeURIComponent([voie, e.code_postal, e.libelle_commune].join(" ")))
        .then(function (r) { return r.json(); }).then(function (j) {
          var f = (j.features || [])[0];
          if (f && f.properties.type === "housenumber" && f.properties.postcode === cand.cp && f.properties.score >= 0.6) {
            var n = banVersAdresse(f);
            cand.adresse = n.adresse; cand.ville = n.ville; if (!cand.lat) { cand.lat = n.lat; cand.lon = n.lon; }
          }
        }).catch(function () {}).then(function () { a = cand; remplir(); });
    });

    q("[data-libre]").addEventListener("click", function () {
      a.statut = "Saisie libre"; a.lat = ""; a.lon = "";
      remplir(); eAdr.focus();
    });
    q("[data-vider]").addEventListener("click", function () {
      a = { adresse: "", complement: "", cp: "", ville: "", lat: "", lon: "", statut: "", siret: "" }; remplir();
    });
    m.querySelectorAll("[data-x]").forEach(function (b) { b.addEventListener("click", fermer); });
    q("[data-ok]").addEventListener("click", function () {
      a.adresse = eAdr.value.trim(); a.complement = eCpl.value.trim(); a.cp = eCp.value.trim(); a.ville = eVille.value.trim();
      if (a.adresse || a.cp || a.ville) {
        if (!a.ville) { alert("Indiquez au moins la ville."); return; }
        if (a.cp && !/^\d{5}$/.test(a.cp)) { alert("Le code postal doit comporter 5 chiffres."); return; }
        if (!OK[a.statut] && a.statut !== "Commune") {
          if (!confirm("Cette adresse n'a pas été choisie dans la Base Adresse Nationale ni dans SIRENE.\n\nLe serveur va tenter de la retrouver ; à défaut elle sera marquée « à vérifier ».\n\nEnregistrer quand même ?")) return;
          a.statut = "Saisie libre";
        }
      } else a.statut = "";
      fermer();
      if (opts.onValider) opts.onValider(a);
    });

    q("[data-sq]").value = opts.nomEntreprise || "";
    remplir();
    setTimeout(function () { (opts.nomEntreprise && !a.adresse ? q("[data-sq]") : eAdr).focus(); }, 50);
    if (opts.nomEntreprise && !a.adresse && !a.ville) chercherSirene();
  }

  window.HecgAdresse = { ouvrir: ouvrir, texte: texte, badge: badge, estVerifiee: estVerifiee };

  // ---------- Postes des contacts ----------
  var K = "hecg_postes_v1", postesP = null;
  function chargerPostes() {
    if (postesP) return postesP;
    try { var c = JSON.parse(sessionStorage.getItem(K) || "null"); if (c && Date.now() - c.ts < 600000) return (postesP = Promise.resolve(c.data)); } catch (e) {}
    var call = window.apiPost ? window.apiPost({ action: "getPostesContacts" }) : Promise.resolve(null);
    postesP = Promise.resolve(call).then(function (r) {
      var data = r && r.success ? r.data.map(function (x) { return x.poste; }) : ["RH", "Gérant", "Expert comptable", "Collaborateur comptable"];
      try { sessionStorage.setItem(K, JSON.stringify({ ts: Date.now(), data: data })); } catch (e) {}
      return data;
    }).catch(function () { postesP = null; return ["RH", "Gérant", "Expert comptable", "Collaborateur comptable"]; });
    return postesP;
  }
  function attacher(input) {
    if (!input || input.dataset.postes) return;
    input.dataset.postes = "1";
    var id = "dl_" + (input.id || Math.random().toString(36).slice(2));
    var dl = document.createElement("datalist"); dl.id = id;
    input.setAttribute("list", id); input.setAttribute("autocomplete", "off");
    input.parentNode.insertBefore(dl, input.nextSibling);
    var liste = [];
    chargerPostes().then(function (d) { liste = d || []; dl.innerHTML = liste.map(function (p) { return '<option value="' + esc(p) + '"></option>'; }).join(""); });
    input.addEventListener("change", function () {
      var v = input.value.replace(/\s+/g, " ").trim();
      if (!v) return;
      var n = function (s) { return s.normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase(); };
      var exist = liste.filter(function (p) { return n(p) === n(v); })[0];
      input.value = exist || v.charAt(0).toUpperCase() + v.slice(1);
      if (!exist) { liste.push(input.value); dl.insertAdjacentHTML("beforeend", '<option value="' + esc(input.value) + '"></option>'); try { sessionStorage.removeItem(K); } catch (e) {} }
    });
  }
  window.HecgPostes = { attacher: attacher, charger: chargerPostes };
})();
