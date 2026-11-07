/* Plans de salle — éditeur (Vie scolaire → Salles & Simulations)
   Backend : salles_plan.gs (getSallePlan / saveSallePlan). Dépend de vie_scolaire.html : api(), toast(), canWrite(), allSalles.
   Coordonnées en « pixels salle » (non mis à l'échelle) ; la salle est affichée réduite pour tenir à l'écran. */
(function () {
const GRID = 10;
const $ = id => document.getElementById(id);
const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const snap = v => Math.round(v / GRID) * GRID;
const uid = () => 'B' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
const P = { salle: null, plan: null, sel: null, classe: '', ro: true, scale: 1, parClasse: {}, classes: [], dirty: false };

function css() {
  if ($('splanCss')) return;
  const st = document.createElement('style'); st.id = 'splanCss';
  st.textContent = `
  .sp-ov{position:fixed;inset:0;z-index:9998;background:#eef2f8;display:flex;flex-direction:column;font-family:'Plus Jakarta Sans',Inter,system-ui,sans-serif}
  .sp-top{display:flex;flex-wrap:wrap;align-items:center;gap:8px;padding:12px 18px;background:#fff;border-bottom:1px solid #e3eaf3;box-shadow:0 6px 20px -16px rgba(14,26,43,.5)}
  .sp-top h2{margin:0 12px 0 0;font-size:18px;font-weight:800;color:#1A4E8A}.sp-top h2 small{display:block;font-size:11.5px;color:#7d8a9c;font-weight:600}
  .sp-b{height:36px;padding:0 12px;border-radius:10px;border:1px solid #dfe6f0;background:#fff;font:700 12.5px inherit;font-family:inherit;color:#334155;cursor:pointer;white-space:nowrap}
  .sp-b:hover{background:#f3f6fb}.sp-b.pri{background:#1A4E8A;border-color:#1A4E8A;color:#fff}.sp-b.ora{background:#F26522;border-color:#F26522;color:#fff}.sp-b.red{color:#dc2626}
  .sp-b:disabled{opacity:.45;cursor:not-allowed}
  .sp-top select{height:36px;border:1px solid #dfe6f0;border-radius:10px;padding:0 8px;font:700 12.5px inherit;font-family:inherit;width:auto}
  .sp-body{flex:1;display:flex;min-height:0}
  .sp-stage{flex:1;overflow:auto;padding:28px;position:relative}
  .sp-scaler{transform-origin:0 0;position:relative}
  .sp-room{position:relative;background:#fff;border:3px solid #1A4E8A;border-radius:6px;box-shadow:0 30px 60px -40px rgba(14,26,43,.6);
    background-image:linear-gradient(#f1f5fb 1px,transparent 1px),linear-gradient(90deg,#f1f5fb 1px,transparent 1px);background-size:${GRID * 2}px ${GRID * 2}px}
  .sp-room .board{position:absolute;left:50%;top:-3px;transform:translateX(-50%);width:40%;height:8px;background:#334155;border-radius:0 0 4px 4px}
  .sp-room .boardl{position:absolute;left:50%;top:8px;transform:translateX(-50%);font-size:10px;font-weight:800;letter-spacing:.12em;color:#94a3b8;text-transform:uppercase}
  .sp-desk{position:absolute;border-radius:7px;background:linear-gradient(180deg,#fff7ed,#ffedd5);border:2px solid #fdba74;box-shadow:0 6px 12px -10px rgba(14,26,43,.7);display:flex;flex-direction:column;cursor:grab;user-select:none;touch-action:none}
  .sp-desk.prof{background:linear-gradient(180deg,#e8eff8,#dbe6f5);border-color:#1A4E8A}
  .sp-desk.sel{outline:3px solid #F26522;outline-offset:2px;z-index:5}
  .sp-desk .lab{font-size:9.5px;font-weight:800;color:#9a3412;padding:2px 5px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
  .sp-desk.prof .lab{color:#1A4E8A}
  .sp-seats{flex:1;display:flex;gap:3px;padding:0 4px 4px}
  .sp-seat{flex:1;min-width:0;border-radius:5px;background:#fff;border:1.5px dashed #fdba74;display:flex;align-items:center;justify-content:center;text-align:center;font-size:10px;font-weight:700;color:#7c2d12;line-height:1.1;padding:2px;overflow:hidden;word-break:break-word}
  .sp-seat.on{border-style:solid;background:#ecfdf5;border-color:#34d399;color:#065f46}
  .sp-desk.prof .sp-seat{border-color:#93b4dc;color:#1A4E8A}
  .sp-rs{position:absolute;right:-6px;bottom:-6px;width:13px;height:13px;border-radius:4px;background:#F26522;border:2px solid #fff;cursor:nwse-resize;touch-action:none}
  .sp-room>.sp-rs{width:18px;height:18px;right:-10px;bottom:-10px;background:#1A4E8A}
  .sp-side{width:330px;flex-shrink:0;background:#fff;border-left:1px solid #e3eaf3;overflow:auto;padding:18px}
  .sp-side h3{margin:0 0 10px;font-size:13px;font-weight:800;letter-spacing:.06em;text-transform:uppercase;color:#7d8a9c}
  .sp-card{border:1px solid #e3eaf3;border-radius:14px;padding:14px;margin-bottom:14px;background:linear-gradient(180deg,#fff,#f8fafd)}
  .sp-kpi{display:grid;grid-template-columns:1fr 1fr;gap:8px}.sp-kpi div{background:#f5f8fc;border-radius:10px;padding:8px 10px}
  .sp-kpi b{display:block;font-size:20px;color:#0e1a2b}.sp-kpi span{font-size:11px;color:#7d8a9c;font-weight:700}
  .sp-side label{display:block;font-size:10.5px;font-weight:800;letter-spacing:.06em;text-transform:uppercase;color:#94a3b8;margin:10px 0 4px}
  .sp-side input,.sp-side select{width:100%;height:36px;border:1px solid #dfe6f0;border-radius:9px;padding:0 10px;font:600 13px inherit;font-family:inherit}
  .sp-miss{font-size:12px;color:#b91c1c;line-height:1.5;max-height:160px;overflow:auto}
  .sp-hint{font-size:12px;color:#7d8a9c;line-height:1.5}
  @media print{.sp-top,.sp-side{display:none!important}.sp-ov{position:static;background:#fff}.sp-stage{overflow:visible;padding:0}}`;
  document.head.appendChild(st);
}

const placesEleves = () => P.plan.bureaux.filter(b => b.type !== 'prof').reduce((s, b) => s + b.places.length, 0);
const nomsPlaces = () => P.plan.bureaux.filter(b => b.type !== 'prof').flatMap(b => b.places).filter(x => String(x || '').trim());
const etuClasse = () => (P.parClasse[P.classe] || []);

function deskSize(n, type) { return type === 'prof' ? { w: 140, h: 70 } : { w: 70 * n + 10, h: 70 }; }

async function open(salleId) {
  css();
  P.salle = (window.allSalles || []).find(s => s.id === salleId) || { id: salleId, numsalle: salleId.split('||')[1], campus: salleId.split('||')[0], capacite: 0 };
  P.ro = !(typeof canWrite === 'function' && canWrite('salles'));
  const ov = document.createElement('div'); ov.className = 'sp-ov'; ov.id = 'splanOv';
  ov.innerHTML = `<div class="sp-top"><h2>Plan — Salle ${esc(P.salle.numsalle)}<small>${esc(P.salle.campus)} · ${P.salle.capacite || '?'} places configurées</small></h2><span style="flex:1"></span><button class="sp-b" onclick="SPLAN.close()">Fermer</button></div>
    <div class="sp-body"><div class="sp-stage" id="spStage"><p class="sp-hint">Chargement du plan…</p></div><div class="sp-side" id="spSide"></div></div>`;
  document.body.appendChild(ov);
  const r = await api('getSallePlan', { salleId }, { forceRefresh: true });
  if (!r.success) { toast('Erreur : ' + r.message, false); close(true); return; }
  P.parClasse = r.etudiantsParClasse || {}; P.classes = r.classes || [];
  P.plan = r.plan && Array.isArray(r.plan.bureaux) ? r.plan : { largeur: 900, hauteur: 600, bureaux: [] };
  P.plan.bureaux.forEach(b => { if (!b.id) b.id = uid(); if (!Array.isArray(b.places)) b.places = []; });
  P.classe = r.classe || P.salle.classeDebut || '';
  P.sel = null; P.dirty = false; P.majInfo = r.maj ? `Enregistré le ${r.maj}${r.par ? ' par ' + r.par : ''}` : 'Jamais enregistré';
  renderTop(); render();
  if (!P.plan.bureaux.length && !P.ro) generer();
  window.addEventListener('resize', fit);
}
function close(force) {
  if (!force && P.dirty && !confirm('Fermer sans enregistrer les modifications du plan ?')) return;
  const o = $('splanOv'); if (o) o.remove(); window.removeEventListener('resize', fit);
}
function dirty() { P.dirty = true; const b = $('spSave'); if (b) b.textContent = 'Enregistrer •'; }

function renderTop() {
  const top = document.querySelector('#splanOv .sp-top');
  const ro = P.ro;
  top.innerHTML = `<h2>Plan — Salle ${esc(P.salle.numsalle)}<small>${esc(P.salle.campus)} · ${P.salle.capacite || '?'} places configurées · ${esc(P.majInfo)}</small></h2>
    ${ro ? '<span class="sp-hint">Consultation seule</span>' : `<button class="sp-b" onclick="SPLAN.generer()">🧩 Générer une disposition</button>
    <button class="sp-b" onclick="SPLAN.ajouter('eleve')">+ Bureau élève</button><button class="sp-b" onclick="SPLAN.ajouter('prof')">+ Bureau prof</button>`}
    <span style="width:1px;height:26px;background:#e3eaf3;margin:0 4px"></span>
    <select id="spClasse" onchange="SPLAN.setClasse(this.value)"><option value="">— Classe affectée —</option>${P.classes.map(c => `<option value="${esc(c.classe)}" ${c.classe === P.classe ? 'selected' : ''}>${esc(c.classe)} (${c.effectif})</option>`).join('')}</select>
    ${ro ? '' : `<button class="sp-b ora" onclick="SPLAN.remplir()">👥 Remplir avec la classe</button><button class="sp-b" onclick="SPLAN.vider()">Vider les noms</button>`}
    <span style="flex:1"></span>
    <button class="sp-b" onclick="SPLAN.imprimer()">🖨 Imprimer</button>
    ${ro ? '' : `<button class="sp-b pri" id="spSave" onclick="SPLAN.enregistrer()">Enregistrer</button>`}
    <button class="sp-b" onclick="SPLAN.close()">Fermer</button>`;
}

function fit() {
  const st = $('spStage'), sc = document.querySelector('#spStage .sp-scaler'); if (!st || !sc) return;
  const s = Math.min(1, (st.clientWidth - 70) / (P.plan.largeur + 20), Math.max(0.35, (st.clientHeight - 70) / (P.plan.hauteur + 20)));
  P.scale = Math.max(0.3, s);
  sc.style.transform = `scale(${P.scale})`; sc.style.width = (P.plan.largeur + 20) + 'px'; sc.style.height = (P.plan.hauteur + 20) + 'px';
  sc.parentElement.style.minHeight = ((P.plan.hauteur + 30) * P.scale) + 'px';
}

function deskHtml(b) {
  const names = b.places.length ? b.places : [''];
  return `<div class="sp-desk ${b.type === 'prof' ? 'prof' : ''} ${P.sel === b.id ? 'sel' : ''}" data-id="${b.id}" style="left:${b.x}px;top:${b.y}px;width:${b.w}px;height:${b.h}px">
    <div class="lab">${esc(b.label || (b.type === 'prof' ? '👩‍🏫 Bureau prof' : 'Bureau'))}</div>
    <div class="sp-seats">${(b.type === 'prof' ? [b.places[0] || 'Professeur'] : names).map((n, i) => `<div class="sp-seat ${String(n || '').trim() && b.type !== 'prof' ? 'on' : ''}" data-seat="${i}" title="${esc(n || 'Place ' + (i + 1))}">${esc(n || (i + 1))}</div>`).join('')}</div>
    ${P.ro ? '' : '<div class="sp-rs" data-rs="desk"></div>'}</div>`;
}

function render() {
  const st = $('spStage');
  st.innerHTML = `<div><div class="sp-scaler"><div class="sp-room" id="spRoom" style="width:${P.plan.largeur}px;height:${P.plan.hauteur}px">
      <div class="board"></div><div class="boardl">Tableau</div>
      ${P.plan.bureaux.map(deskHtml).join('')}
      ${P.ro ? '' : '<div class="sp-rs" data-rs="room" title="Redimensionner la salle"></div>'}
    </div></div></div>`;
  fit(); bind(); renderSide();
}

function renderSide() {
  const places = placesEleves(), etus = etuClasse(), noms = nomsPlaces();
  const cap = P.salle.capacite || 0;
  const taux = places ? Math.round(etus.length / places * 100) : 0;
  const coul = taux > 100 ? '#dc2626' : (taux >= 85 ? '#f97316' : '#16a34a');
  const sansPlace = etus.filter(e => !noms.includes(e.nom));
  const b = P.plan.bureaux.find(x => x.id === P.sel);
  $('spSide').innerHTML = `<div class="sp-card"><h3>Occupation</h3>
      <div class="sp-kpi"><div><b>${places}</b><span>places élèves</span></div><div><b>${P.classe ? etus.length : '—'}</b><span>${P.classe ? 'étudiants ' + esc(P.classe) : 'aucune classe'}</span></div>
      <div><b style="color:${P.classe ? coul : '#0e1a2b'}">${P.classe && places ? taux + ' %' : '—'}</b><span>taux d'occupation</span></div><div><b>${noms.length}</b><span>places nommées</span></div></div>
      ${P.classe ? `<p class="sp-hint" style="margin:10px 0 0;font-weight:700;color:${coul}">${etus.length > places ? `⚠️ Il manque ${etus.length - places} place(s) pour toute la classe.` : (etus.length === places ? '✓ Salle exactement à la taille de la classe.' : `✓ ${places - etus.length} place(s) libre(s).`)}</p>` : ''}
      ${cap && cap !== places ? `<p class="sp-hint" style="margin:6px 0 0">Capacité configurée de la salle : ${cap} — le plan compte ${places} place(s).</p>` : ''}
      ${P.classe && noms.length && sansPlace.length ? `<label>Étudiants sans place (${sansPlace.length})</label><div class="sp-miss">${sansPlace.map(e => esc(e.nom)).join('<br>')}</div>` : ''}</div>
    ${b ? `<div class="sp-card"><h3>${b.type === 'prof' ? 'Bureau prof' : 'Bureau sélectionné'}</h3>
      <label>Libellé</label><input id="spLab" value="${esc(b.label || '')}" placeholder="${b.type === 'prof' ? 'Bureau prof' : 'Ex : Rang 1 — Bureau 3'}" ${P.ro ? 'disabled' : ''}>
      ${b.type === 'prof' ? `<label>Nom</label><input data-pl="0" value="${esc(b.places[0] || '')}" placeholder="Professeur" ${P.ro ? 'disabled' : ''}>`
        : `<label>Nombre de places</label><select id="spNb" ${P.ro ? 'disabled' : ''}>${[1, 2, 3].map(n => `<option ${b.places.length === n ? 'selected' : ''}>${n}</option>`).join('')}</select>
      ${b.places.map((n, i) => `<label>Place ${i + 1}</label><input data-pl="${i}" value="${esc(n)}" placeholder="Nom de l'étudiant" list="spEtuList" ${P.ro ? 'disabled' : ''}>`).join('')}
      <datalist id="spEtuList">${etus.map(e => `<option value="${esc(e.nom)}">`).join('')}</datalist>`}
      ${P.ro ? '' : `<div style="display:flex;gap:8px;margin-top:14px"><button class="sp-b" onclick="SPLAN.dupliquer()">⧉ Dupliquer</button><button class="sp-b red" onclick="SPLAN.supprimer()">🗑 Supprimer</button></div>`}</div>`
      : `<div class="sp-card"><h3>Mode d'emploi</h3><p class="sp-hint">${P.ro ? 'Cliquez sur un bureau pour voir le détail des places.' : `• Glissez les bureaux pour les placer, la poignée orange les redimensionne.<br>• La poignée bleue (coin de la salle) ajuste la taille de la salle.<br>• Cliquez sur un bureau pour le renommer, changer son nombre de places (1 à 3) ou nommer chaque place.<br>• Choisissez la classe affectée puis « Remplir avec la classe » pour placer les étudiants par ordre alphabétique.`}</p></div>`}`;
  if (b && !P.ro) {
    const lab = $('spLab'); if (lab) lab.oninput = e => { b.label = e.target.value; dirty(); refreshDesk(b); };
    const nb = $('spNb'); if (nb) nb.onchange = e => { const n = +e.target.value; while (b.places.length < n) b.places.push(''); b.places.length = n; const sz = deskSize(n); b.w = Math.max(b.w, sz.w); dirty(); render(); };
    document.querySelectorAll('#spSide [data-pl]').forEach(inp => inp.oninput = e => { b.places[+e.target.dataset.pl] = e.target.value; dirty(); refreshDesk(b); renderSideKpiSoon(); });
  }
}
let _kpiT = null; function renderSideKpiSoon() { clearTimeout(_kpiT); _kpiT = setTimeout(() => { const a = document.activeElement, idx = a && a.dataset ? a.dataset.pl : null; renderSide(); if (idx != null) { const i = document.querySelector(`#spSide [data-pl="${idx}"]`); if (i) { i.focus(); i.setSelectionRange(i.value.length, i.value.length); } } }, 700); }
function refreshDesk(b) { const el = document.querySelector(`.sp-desk[data-id="${b.id}"]`); if (el) el.outerHTML = deskHtml(b); bind(); }

function bind() {
  const room = $('spRoom'); if (!room) return;
  room.querySelectorAll('.sp-desk').forEach(el => {
    el.onpointerdown = ev => {
      const b = P.plan.bureaux.find(x => x.id === el.dataset.id); if (!b) return;
      const seat = ev.target.closest('[data-seat]');
      if (P.sel !== b.id) { P.sel = b.id; room.querySelectorAll('.sp-desk').forEach(d => d.classList.toggle('sel', d.dataset.id === b.id)); renderSide(); }
      if (seat && !ev.target.closest('[data-rs]')) setTimeout(() => { const i = document.querySelector(`#spSide [data-pl="${seat.dataset.seat}"]`); if (i) i.focus(); }, 0);
      if (P.ro) return;
      ev.preventDefault();
      const resize = !!ev.target.closest('[data-rs]');
      const sx = ev.clientX, sy = ev.clientY, ox = b.x, oy = b.y, ow = b.w, oh = b.h;
      el.setPointerCapture(ev.pointerId);
      el.onpointermove = m => {
        const dx = (m.clientX - sx) / P.scale, dy = (m.clientY - sy) / P.scale;
        if (resize) { b.w = Math.max(60, snap(ow + dx)); b.h = Math.max(50, snap(oh + dy)); el.style.width = b.w + 'px'; el.style.height = b.h + 'px'; }
        else { b.x = Math.max(0, Math.min(P.plan.largeur - b.w, snap(ox + dx))); b.y = Math.max(0, Math.min(P.plan.hauteur - b.h, snap(oy + dy))); el.style.left = b.x + 'px'; el.style.top = b.y + 'px'; }
        if (dx || dy) P.moved = true;
      };
      el.onpointerup = () => { el.onpointermove = null; el.onpointerup = null; if (P.moved) dirty(); P.moved = false; };
    };
  });
  room.onpointerdown = ev => {
    if (ev.target === room) { P.sel = null; room.querySelectorAll('.sp-desk').forEach(d => d.classList.remove('sel')); renderSide(); }
  };
  const rs = room.querySelector(':scope > [data-rs="room"]');
  if (rs) rs.onpointerdown = ev => {
    ev.preventDefault(); ev.stopPropagation();
    const sx = ev.clientX, sy = ev.clientY, ow = P.plan.largeur, oh = P.plan.hauteur;
    const minW = Math.max(300, ...P.plan.bureaux.map(b => b.x + b.w)), minH = Math.max(200, ...P.plan.bureaux.map(b => b.y + b.h));
    rs.setPointerCapture(ev.pointerId);
    rs.onpointermove = m => {
      P.plan.largeur = Math.max(minW, snap(ow + (m.clientX - sx) / P.scale)); P.plan.hauteur = Math.max(minH, snap(oh + (m.clientY - sy) / P.scale));
      room.style.width = P.plan.largeur + 'px'; room.style.height = P.plan.hauteur + 'px';
    };
    rs.onpointerup = () => { rs.onpointermove = null; rs.onpointerup = null; dirty(); fit(); };
  };
}

// Place libre la plus proche du centre-haut de la salle
function placeLibre(w, h) {
  for (let y = 60; y + h <= P.plan.hauteur; y += GRID * 2)
    for (let x = 20; x + w <= P.plan.largeur; x += GRID * 2) {
      if (!P.plan.bureaux.some(b => x < b.x + b.w + 10 && x + w + 10 > b.x && y < b.y + b.h + 10 && y + h + 10 > b.y)) return { x, y };
    }
  return { x: 20, y: 60 };
}
function ajouter(type) {
  const n = type === 'prof' ? 1 : 2, sz = deskSize(n, type);
  const pos = type === 'prof' ? { x: snap(P.plan.largeur / 2 - sz.w / 2), y: 30 } : placeLibre(sz.w, sz.h);
  const b = { id: uid(), type, x: pos.x, y: pos.y, w: sz.w, h: sz.h, places: type === 'prof' ? [''] : ['', ''], label: '' };
  P.plan.bureaux.push(b); P.sel = b.id; dirty(); render();
}
function dupliquer() {
  const b = P.plan.bureaux.find(x => x.id === P.sel); if (!b) return;
  const pos = placeLibre(b.w, b.h);
  const c = { ...b, id: uid(), x: pos.x, y: pos.y, places: b.places.map(() => ''), label: '' };
  P.plan.bureaux.push(c); P.sel = c.id; dirty(); render();
}
function supprimer() {
  const b = P.plan.bureaux.find(x => x.id === P.sel); if (!b) return;
  if (b.places.some(x => String(x || '').trim()) && !confirm('Supprimer ce bureau et les noms qu\'il contient ?')) return;
  P.plan.bureaux = P.plan.bureaux.filter(x => x.id !== P.sel); P.sel = null; dirty(); render();
}

function generer() {
  const d = document.createElement('div'); d.className = 'modal-overlay'; d.style.zIndex = 9999;
  const effectif = etuClasse().length || P.salle.capacite || 24;
  d.innerHTML = `<div class="modal-box p-8" style="max-width:480px"><h3 class="text-xl font-black text-[#1A4E8A] mb-1">🧩 Générer une disposition</h3>
    <p class="text-xs text-gray-400 mb-4">Remplace les bureaux actuels par une disposition en rangées.${P.classe ? ` Classe affectée : ${esc(P.classe)} (${etuClasse().length} étudiants).` : ''}</p>
    <div class="grid grid-cols-3 gap-3">
      <div><label>Rangées</label><input type="number" id="gR" min="1" max="12" value="${Math.max(2, Math.ceil(effectif / 12))}"></div>
      <div><label>Bureaux / rangée</label><input type="number" id="gB" min="1" max="12" value="${Math.min(8, Math.max(2, Math.ceil(effectif / (Math.max(2, Math.ceil(effectif / 12)) * 2))))}"></div>
      <div><label>Places / bureau</label><select id="gP"><option>1</option><option selected>2</option><option>3</option></select></div>
    </div>
    <label style="display:flex;align-items:center;gap:8px;margin-top:14px;text-transform:none;font-size:13px;color:#374151"><input type="checkbox" id="gProf" checked style="width:auto"> Ajouter un bureau prof (devant le tableau)</label>
    <label style="display:flex;align-items:center;gap:8px;margin-top:8px;text-transform:none;font-size:13px;color:#374151"><input type="checkbox" id="gAllee" checked style="width:auto"> Allée centrale</label>
    <p class="text-sm font-bold text-[#1A4E8A] mt-4" id="gTot"></p>
    <div class="flex justify-end gap-3 mt-6"><button class="btn-ghost" id="gNo">Annuler</button><button class="btn-orange" id="gGo">Générer</button></div></div>`;
  document.body.appendChild(d);
  const tot = () => { const n = (+$('gR').value || 0) * (+$('gB').value || 0) * (+$('gP').value || 0); $('gTot').textContent = `${n} place(s) élèves${P.classe ? ' pour ' + etuClasse().length + ' étudiant(s)' : ''}`; };
  ['gR', 'gB', 'gP'].forEach(id => $(id).oninput = tot); tot();
  $('gNo').onclick = () => d.remove();
  $('gGo').onclick = () => {
    const R = Math.max(1, Math.min(12, +$('gR').value || 1)), B = Math.max(1, Math.min(12, +$('gB').value || 1)), N = +$('gP').value || 2;
    const allee = $('gAllee').checked && B >= 2;
    const sz = deskSize(N), gapX = 30, gapY = 40, aw = allee ? 70 : 0, top = $('gProf').checked ? 140 : 70;
    const rowW = B * sz.w + (B - 1) * gapX + aw;
    P.plan.largeur = Math.max(P.plan.largeur, snap(rowW + 120)); P.plan.hauteur = Math.max(400, snap(top + R * (sz.h + gapY) + 40));
    const x0 = snap((P.plan.largeur - rowW) / 2);
    const bureaux = [];
    if ($('gProf').checked) { const pz = deskSize(1, 'prof'); bureaux.push({ id: uid(), type: 'prof', x: snap(P.plan.largeur / 2 - pz.w / 2), y: 30, w: pz.w, h: pz.h, places: [''], label: '' }); }
    for (let r = 0; r < R; r++) for (let c = 0; c < B; c++) {
      const x = x0 + c * (sz.w + gapX) + (allee && c >= Math.ceil(B / 2) ? aw : 0);
      bureaux.push({ id: uid(), type: 'eleve', x: snap(x), y: snap(top + r * (sz.h + gapY)), w: sz.w, h: sz.h, places: Array(N).fill(''), label: `Rang ${r + 1} · ${c + 1}` });
    }
    P.plan.bureaux = bureaux; P.sel = null; d.remove(); dirty(); render();
    if (P.classe) remplir(true);
  };
}

function setClasse(c) { P.classe = c; dirty(); renderSide(); }
// Bureaux triés de devant vers le fond, de gauche à droite
const ordreBureaux = () => P.plan.bureaux.filter(b => b.type !== 'prof').sort((a, b) => (Math.round(a.y / 40) - Math.round(b.y / 40)) || (a.x - b.x));
function remplir(silencieux) {
  if (!P.classe) { toast('Choisissez d\'abord la classe affectée', false); return; }
  const etus = etuClasse().map(e => e.nom);
  if (!silencieux && nomsPlaces().length && !confirm(`Remplacer les noms actuels par les ${etus.length} étudiants de ${P.classe} ?`)) return;
  let k = 0;
  ordreBureaux().forEach(b => { b.places = b.places.map(() => etus[k++] || ''); });
  dirty(); render();
  const places = placesEleves();
  toast(etus.length > places ? `⚠️ ${etus.length - places} étudiant(s) sans place : agrandissez la disposition` : `${etus.length} étudiant(s) placé(s) — ${places - etus.length} place(s) libre(s) ✓`, etus.length <= places);
}
function vider() { if (!confirm('Effacer tous les noms des places ?')) return; P.plan.bureaux.forEach(b => { if (b.type !== 'prof') b.places = b.places.map(() => ''); }); dirty(); render(); }

async function enregistrer() {
  const btn = $('spSave'); if (btn) { btn.disabled = true; btn.textContent = 'Enregistrement…'; }
  const r = await api('saveSallePlan', { salleId: P.salle.id, classe: P.classe, plan: P.plan });
  if (btn) btn.disabled = false;
  if (!r.success) { if (btn) btn.textContent = 'Enregistrer •'; toast('Erreur : ' + r.message, false); return; }
  P.dirty = false; if (btn) btn.textContent = 'Enregistré ✓';
  toast('Plan de salle enregistré ✓');
}

function imprimer() {
  const room = $('spRoom'); if (!room) return;
  const w = window.open('', '_blank'); if (!w) { toast('Autorisez les fenêtres pop-up pour imprimer', false); return; }
  const places = placesEleves(), etus = etuClasse();
  const sc = Math.min(1, 1050 / P.plan.largeur, 620 / P.plan.hauteur);
  w.document.write(`<!DOCTYPE html><html><head><meta charset="utf-8"><title>Plan salle ${esc(P.salle.numsalle)}</title><style>${$('splanCss').textContent}
    body{font-family:Arial,sans-serif;margin:18px}h1{font-size:18px;color:#1A4E8A;margin:0 0 4px}p{margin:0 0 12px;color:#64748b;font-size:12px}
    .sp-rs{display:none}.sp-desk.sel{outline:none}@page{size:A4 landscape;margin:10mm}</style></head><body>
    <h1>Plan de salle — Salle ${esc(P.salle.numsalle)} (${esc(P.salle.campus)})</h1>
    <p>${places} place(s) élèves${P.classe ? ` · Classe ${esc(P.classe)} : ${etus.length} étudiant(s) · taux ${places ? Math.round(etus.length / places * 100) : 0} %` : ''} · édité le ${new Date().toLocaleDateString('fr-FR')}</p>
    <div style="transform:scale(${sc});transform-origin:0 0">${room.outerHTML}</div></body></html>`);
  w.document.close(); setTimeout(() => w.print(), 400);
}

window.SPLAN = { open, close: () => close(false), generer, ajouter, dupliquer, supprimer, setClasse, remplir: () => remplir(false), vider, enregistrer, imprimer };
})();
