/* =============================================================================
   🧊 CUBE — Analyse croisée dynamique (Vie scolaire → Tableaux de suivis)
   -----------------------------------------------------------------------------
   - Sources : étudiants, entreprises, contacts, offres, campagnes, activité CRM
     (union possible : colonne « Source ») + enrichissement par la fiche entreprise.
   - Glisser-déposer des champs dans Filtres / Lignes / Colonnes / Valeurs
     (ou clic sur un champ → menu).
   - Agrégations : nombre, nombre distinct, somme, moyenne, min, max.
   - Affichage : valeurs, % du total, % ligne, % colonne ; carte de chaleur ;
     sous-totaux repliables ; tri ; top N ; vue graphique ; détail d'une cellule.
   - Vues enregistrées (navigateur) + modèles rapides.
   - Export CSV et « Envoyer vers Google Sheet » : crée un vrai tableau croisé
     dynamique dans un tableau de suivi (backend cube.gs), resynchronisé chaque heure.
   Dépend de vie_scolaire.html : api(), toast(), escTa(), tableauxSuivi, vieScRole, user.
   ========================================================================== */
(function () {
'use strict';

const CUBE_SRC = [
  { k: 'etudiants',   l: 'Étudiants',    i: '🎓' },
  { k: 'entreprises', l: 'Entreprises',  i: '🏢' },
  { k: 'contacts',    l: 'Contacts',     i: '👤' },
  { k: 'offres',      l: 'Offres',       i: '💼' },
  { k: 'campagnes',   l: 'Campagnes',    i: '📣' },
  { k: 'activite',    l: 'Activité CRM', i: '📈', admin: true }
];
const AGG = { count: 'Nombre', countd: 'Nombre distinct', sum: 'Somme', avg: 'Moyenne', min: 'Min', max: 'Max' };
const VIDE = '(vide)';
const VIEWS_KEY = 'hecg_cube_views_v1';
const PRESETS = [
  { nom: 'Étudiants par classe × statut', sources: ['etudiants'], join: false, rows: ['Classe'], cols: ['Statut'], vals: [{ f: '', agg: 'count' }], filters: {} },
  { nom: 'Alternants par département d\'entreprise', sources: ['etudiants'], join: true, rows: ['Entreprise › Département', 'Entreprise › Ville'], cols: ['Campus'], vals: [{ f: '', agg: 'count' }], filters: {} },
  { nom: 'Ateliers CV / LinkedIn par classe', sources: ['etudiants'], join: false, rows: ['Classe'], cols: ['Atelier CV', 'Atelier LinkedIn'], vals: [{ f: '', agg: 'count' }], filters: {} },
  { nom: 'Offres par mois × qualité', sources: ['offres'], join: false, rows: ['Date · Mois'], cols: ['Qualité'], vals: [{ f: '', agg: 'count' }], filters: {} },
  { nom: 'Entreprises par département (alternants)', sources: ['entreprises'], join: false, rows: ['Département'], cols: [], vals: [{ f: '', agg: 'count' }, { f: 'Nb alternants', agg: 'sum' }], filters: {} },
  { nom: 'Vue croisée entreprises + offres par ville', sources: ['entreprises', 'offres'], join: false, rows: ['Ville'], cols: ['Source'], vals: [{ f: '', agg: 'count' }], filters: {} },
  { nom: 'Activité de l\'équipe par mois', sources: ['activite'], join: false, rows: ['Auteur'], cols: ['Date · Mois'], vals: [{ f: '', agg: 'count' }], filters: {}, admin: true }
];

const S = {
  sources: ['etudiants'], join: false,
  fields: [], rows: [], loadedKey: '',
  cfg: { rows: [], cols: [], vals: [{ f: '', agg: 'count' }], filters: {} },
  opt: { show: 'val', heat: true, totals: true, sort: 'label', top: 0, view: 'table' },
  collapsed: new Set(), fieldSearch: '', canActivite: false, lastPivot: null
};

// ── utilitaires ──────────────────────────────────────────────────────────
const $ = id => document.getElementById(id);
const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const SEP = '\u0001';
const collator = new Intl.Collator('fr', { numeric: true, sensitivity: 'base' });
const fmtNum = (v, dec) => v == null || isNaN(v) ? '' : Number(v).toLocaleString('fr-FR', { maximumFractionDigits: dec == null ? 2 : dec });
const fieldIdx = n => S.fields.findIndex(f => f.n === n);
const fieldType = n => (S.fields.find(f => f.n === n) || {}).t || 'txt';
const typeIcon = t => t === 'num' ? '#' : (t === 'date' ? '📅' : 'Aa');
const cellVal = (row, i) => { const v = row[i]; return v === '' || v == null ? VIDE : String(v); };

// ── styles ───────────────────────────────────────────────────────────────
function injectCss() {
  if ($('cubeCss')) return;
  const st = document.createElement('style');
  st.id = 'cubeCss';
  st.textContent = `
  #cubeOverlay{position:fixed;inset:0;z-index:9000;background:#f5f7fb;display:flex;flex-direction:column;font-family:Inter,ui-sans-serif,system-ui,sans-serif;animation:cubeIn .25s cubic-bezier(.2,.7,.2,1)}
  @keyframes cubeIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:none}}
  #cubeOverlay *{box-sizing:border-box}
  .cube-top{display:flex;align-items:center;gap:.75rem;padding:.7rem 1.25rem;background:rgba(255,255,255,.9);backdrop-filter:blur(12px);border-bottom:1px solid #e6ebf2}
  .cube-top h2{font-size:1.05rem;font-weight:800;color:#1A4E8A;margin:0;letter-spacing:-.01em;display:flex;align-items:center;gap:.5rem}
  .cube-top .sub{font-size:.75rem;color:#94a3b8;font-weight:600}
  .cube-btn{border:1px solid #e2e8f0;background:#fff;color:#334155;border-radius:10px;padding:.45rem .8rem;font-size:.78rem;font-weight:700;cursor:pointer;transition:all .15s;white-space:nowrap}
  .cube-btn:hover{border-color:#1A4E8A;color:#1A4E8A;background:#eef3fa}
  .cube-btn.pri{background:#1A4E8A;color:#fff;border-color:#1A4E8A}.cube-btn.pri:hover{background:#164378;color:#fff}
  .cube-btn.ora{background:#F26522;color:#fff;border-color:#F26522}.cube-btn.ora:hover{background:#dc5714;color:#fff}
  .cube-btn:disabled{opacity:.5;cursor:not-allowed}
  .cube-body{flex:1;display:grid;grid-template-columns:300px 1fr;min-height:0}
  .cube-side{background:#fff;border-right:1px solid #e6ebf2;display:flex;flex-direction:column;min-height:0}
  .cube-side h3{font-size:.66rem;font-weight:800;letter-spacing:.1em;text-transform:uppercase;color:#94a3b8;margin:0 0 .5rem}
  .cube-blk{padding:1rem 1rem .8rem;border-bottom:1px solid #eef2f7}
  .cube-src{display:flex;flex-wrap:wrap;gap:.35rem}
  .cube-chip{border:1.5px solid #e2e8f0;border-radius:999px;padding:.3rem .65rem;font-size:.75rem;font-weight:700;color:#475569;background:#fff;cursor:pointer;user-select:none;transition:all .15s}
  .cube-chip:hover{border-color:#1A4E8A}
  .cube-chip.on{background:#1A4E8A;border-color:#1A4E8A;color:#fff;box-shadow:0 4px 10px -4px rgba(26,78,138,.6)}
  .cube-chk{display:flex;align-items:center;gap:.45rem;font-size:.78rem;font-weight:600;color:#475569;margin-top:.6rem;cursor:pointer}
  .cube-chk input{accent-color:#1A4E8A;width:15px;height:15px}
  .cube-fields{flex:1;overflow:auto;padding:.4rem .6rem 1rem}
  .cube-field{display:flex;align-items:center;gap:.5rem;padding:.42rem .55rem;border-radius:9px;font-size:.8rem;font-weight:600;color:#334155;cursor:grab;user-select:none;border:1px solid transparent;transition:background .12s}
  .cube-field:hover{background:#eef3fa;border-color:#dce7f5}
  .cube-field .ty{width:22px;height:20px;border-radius:6px;background:#f1f5f9;color:#64748b;font-size:.62rem;font-weight:800;display:flex;align-items:center;justify-content:center;flex-shrink:0}
  .cube-field .ty.num{background:#ecfdf5;color:#047857}.cube-field .ty.date{background:#fff7ed;color:#c2410c}
  .cube-field .used{margin-left:auto;width:7px;height:7px;border-radius:50%;background:#F26522}
  .cube-group{font-size:.64rem;font-weight:800;color:#94a3b8;text-transform:uppercase;letter-spacing:.08em;padding:.7rem .55rem .25rem}
  .cube-search{width:100%;border:1.5px solid #e6ebf2;border-radius:10px;padding:.45rem .7rem;font-size:.8rem;outline:none}
  .cube-search:focus{border-color:#1A4E8A;box-shadow:0 0 0 4px rgba(26,78,138,.12)}
  .cube-main{display:flex;flex-direction:column;min-width:0;min-height:0;padding:1rem 1.25rem;gap:.8rem}
  .cube-zones{display:grid;grid-template-columns:1fr 1fr;gap:.6rem}
  .cube-zone{background:#fff;border:1.5px dashed #d5dde8;border-radius:14px;padding:.55rem .6rem;min-height:58px;transition:all .15s}
  .cube-zone.over{border-color:#F26522;background:#fff7f2;box-shadow:0 0 0 4px rgba(242,101,34,.1)}
  .cube-zone .zt{font-size:.62rem;font-weight:800;letter-spacing:.1em;text-transform:uppercase;color:#94a3b8;margin-bottom:.35rem;display:flex;align-items:center;gap:.4rem}
  .cube-zone .zt b{color:#1A4E8A}
  .cube-pills{display:flex;flex-wrap:wrap;gap:.35rem;min-height:26px}
  .cube-pill{display:inline-flex;align-items:center;gap:.35rem;background:#eef3fa;border:1px solid #dce7f5;color:#1A4E8A;border-radius:999px;padding:.22rem .3rem .22rem .65rem;font-size:.75rem;font-weight:700;cursor:grab;max-width:100%}
  .cube-pill.val{background:#fff4ed;border-color:#fed7bf;color:#c2410c}
  .cube-pill.flt{background:#f8fafc;border-color:#e2e8f0;color:#334155}
  .cube-pill .x{width:18px;height:18px;border-radius:50%;border:none;background:rgba(0,0,0,.06);color:inherit;cursor:pointer;font-size:.7rem;line-height:1;display:flex;align-items:center;justify-content:center}
  .cube-pill .x:hover{background:#ef4444;color:#fff}
  html body #cubeOverlay .cube-pill select{border:none;background:transparent;font-weight:800;color:inherit;font-size:.72rem;padding:0 .1rem!important;cursor:pointer;outline:none;-webkit-appearance:auto;appearance:auto;background-image:none!important;box-shadow:none}
  .cube-pill .cnt{background:#1A4E8A;color:#fff;border-radius:999px;font-size:.62rem;padding:.05rem .4rem}
  .cube-hint{font-size:.72rem;color:#b6c0cd;font-weight:600;padding:.2rem .1rem}
  .cube-opts{display:flex;flex-wrap:wrap;align-items:center;gap:.5rem;background:#fff;border:1px solid #e6ebf2;border-radius:14px;padding:.5rem .7rem}
  .cube-seg{display:inline-flex;background:#f1f5f9;border-radius:10px;padding:2px}
  .cube-seg button{border:none;background:transparent;padding:.32rem .6rem;border-radius:8px;font-size:.72rem;font-weight:700;color:#64748b;cursor:pointer}
  .cube-seg button.on{background:#fff;color:#1A4E8A;box-shadow:0 1px 3px rgba(15,27,45,.12)}
  .cube-opts label{font-size:.72rem;font-weight:700;color:#64748b;display:flex;align-items:center;gap:.35rem}
  .cube-opts select{border:1px solid #e2e8f0;border-radius:8px;padding:.25rem 1.6rem .25rem .5rem;font-size:.72rem;font-weight:700;color:#334155}
  .cube-stat{margin-left:auto;font-size:.72rem;color:#94a3b8;font-weight:700}
  .cube-result{flex:1;min-height:0;background:#fff;border:1px solid #e6ebf2;border-radius:16px;overflow:auto;position:relative;box-shadow:0 1px 2px rgba(15,27,45,.04),0 8px 24px -10px rgba(15,27,45,.08)}
  table.cube-t{border-collapse:separate;border-spacing:0;font-size:.8rem;min-width:100%}
  table.cube-t th,table.cube-t td{padding:.45rem .75rem;white-space:nowrap;border-bottom:1px solid #eef2f7}
  table.cube-t thead th{position:sticky;top:0;background:#f8fafc;z-index:2;font-size:.68rem;font-weight:800;color:#475569;text-align:right;border-bottom:1px solid #e2e8f0;letter-spacing:.02em}
  table.cube-t thead tr:nth-child(2) th{top:31px}
  table.cube-t thead th.rh{text-align:left;left:0;z-index:3}
  table.cube-t tbody th{position:sticky;left:0;background:#fff;text-align:left;font-weight:600;color:#1e293b;z-index:1;max-width:360px;overflow:hidden;text-overflow:ellipsis}
  table.cube-t td{text-align:right;font-variant-numeric:tabular-nums;color:#1e293b;cursor:pointer;transition:outline .1s}
  table.cube-t td:hover{outline:2px solid #F26522;outline-offset:-2px}
  table.cube-t tr.sub th,table.cube-t tr.sub td{background:#f8fafc;font-weight:800}
  table.cube-t tr.sub th{cursor:pointer}
  table.cube-t tr.tot th,table.cube-t tr.tot td{background:#eef3fa;font-weight:800;color:#1A4E8A;border-top:2px solid #dce7f5;position:sticky;bottom:0}
  table.cube-t td.ct,table.cube-t th.ct{background:#f8fafc;font-weight:800}
  .cube-tog{display:inline-block;width:16px;color:#94a3b8;font-size:.7rem}
  .cube-empty{display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%;min-height:280px;color:#94a3b8;text-align:center;gap:.6rem;padding:2rem}
  .cube-empty .big{font-size:2.6rem}
  .cube-presets{display:flex;flex-wrap:wrap;gap:.4rem;justify-content:center;max-width:720px}
  .cube-pop{position:fixed;z-index:9100;background:#fff;border:1px solid #e2e8f0;border-radius:14px;box-shadow:0 20px 50px -12px rgba(15,27,45,.35);padding:.5rem;min-width:220px;max-width:340px;animation:cubeIn .15s}
  .cube-pop .it{display:flex;align-items:center;gap:.5rem;padding:.45rem .6rem;border-radius:9px;font-size:.78rem;font-weight:600;color:#334155;cursor:pointer}
  .cube-pop .it:hover{background:#eef3fa;color:#1A4E8A}
  .cube-pop .vals{max-height:260px;overflow:auto;margin:.3rem 0}
  .cube-pop label{display:flex;align-items:center;gap:.45rem;padding:.3rem .5rem;font-size:.76rem;font-weight:600;color:#334155;border-radius:8px;cursor:pointer}
  .cube-pop label:hover{background:#f8fafc}
  .cube-pop label input{accent-color:#1A4E8A}
  .cube-pop .n{margin-left:auto;color:#94a3b8;font-size:.68rem}
  .cube-modal{position:fixed;inset:0;z-index:9200;background:rgba(15,27,45,.45);backdrop-filter:blur(6px);display:flex;align-items:center;justify-content:center;padding:1rem}
  .cube-modal .box{background:#fff;border-radius:18px;box-shadow:0 30px 80px -20px rgba(15,27,45,.45);width:100%;max-width:560px;max-height:90vh;overflow:auto;padding:1.5rem;animation:cubeIn .2s}
  .cube-modal h3{margin:0 0 .25rem;font-size:1.1rem;font-weight:800;color:#1A4E8A}
  .cube-modal label.l{display:block;font-size:.66rem;font-weight:800;letter-spacing:.08em;text-transform:uppercase;color:#94a3b8;margin:1rem 0 .35rem}
  .cube-modal input[type=text],.cube-modal select{width:100%;border:1.5px solid #e6ebf2;border-radius:10px;padding:.55rem .75rem;font-size:.85rem}
  .cube-bars{padding:1.2rem 1.4rem;display:flex;flex-direction:column;gap:.45rem}
  .cube-bar{display:grid;grid-template-columns:minmax(120px,240px) 1fr 80px;align-items:center;gap:.7rem;font-size:.78rem}
  .cube-bar .lb{font-weight:700;color:#334155;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .cube-bar .tr{height:22px;background:#f1f5f9;border-radius:7px;display:flex;overflow:hidden}
  .cube-bar .sg{height:100%;transition:width .4s cubic-bezier(.2,.7,.2,1)}
  .cube-bar .vl{text-align:right;font-weight:800;color:#1A4E8A;font-variant-numeric:tabular-nums}
  .cube-legend{display:flex;flex-wrap:wrap;gap:.8rem;padding:0 1.4rem 1rem;font-size:.72rem;font-weight:700;color:#475569}
  .cube-legend i{display:inline-block;width:10px;height:10px;border-radius:3px;margin-right:.3rem;vertical-align:-1px}
  .cube-loading{position:absolute;inset:0;background:rgba(255,255,255,.75);display:flex;align-items:center;justify-content:center;font-weight:800;color:#1A4E8A;gap:.6rem;z-index:5}
  .cube-spin{width:20px;height:20px;border:3px solid #dce7f5;border-top-color:#1A4E8A;border-radius:50%;animation:cubeSpin .8s linear infinite}
  @keyframes cubeSpin{to{transform:rotate(360deg)}}
  @media(max-width:900px){.cube-body{grid-template-columns:1fr}.cube-side{max-height:40vh}.cube-zones{grid-template-columns:1fr}}
  `;
  document.head.appendChild(st);
}

// ── ouverture / structure ───────────────────────────────────────────────
function open() {
  injectCss();
  S.canActivite = (typeof vieScRole !== 'undefined' && vieScRole === 'Super-admin') || (typeof user !== 'undefined' && user.role === 'Administrateur');
  if ($('cubeOverlay')) { $('cubeOverlay').style.display = 'flex'; return; }
  const ov = document.createElement('div');
  ov.id = 'cubeOverlay';
  ov.innerHTML = `
    <div class="cube-top">
      <h2>🧊 CUBE <span class="sub">Analyse croisée dynamique</span></h2>
      <div style="flex:1"></div>
      <select id="cubeViews" class="cube-btn" title="Vues enregistrées" onchange="CUBE.loadView(this.value)"></select>
      <button class="cube-btn" onclick="CUBE.saveView()" title="Enregistrer la configuration actuelle">💾 Enregistrer la vue</button>
      <button class="cube-btn" onclick="CUBE.reset()" title="Vider lignes, colonnes, valeurs et filtres">↺ Réinitialiser</button>
      <button class="cube-btn" onclick="CUBE.exportCsv()">⬇ CSV</button>
      <button class="cube-btn ora" onclick="CUBE.openExport()">📤 Envoyer vers Google Sheet</button>
      <button class="cube-btn" onclick="CUBE.close()" title="Fermer (Échap)">✕</button>
    </div>
    <div class="cube-body">
      <aside class="cube-side">
        <div class="cube-blk">
          <h3>1 · Sources de données</h3>
          <div class="cube-src" id="cubeSrc"></div>
          <label class="cube-chk"><input type="checkbox" id="cubeJoin" onchange="CUBE.setJoin(this.checked)"> Enrichir avec la fiche entreprise</label>
          <p style="font-size:.68rem;color:#94a3b8;margin:.4rem 0 0">Plusieurs sources = données empilées avec un champ « Source » pour les croiser.</p>
        </div>
        <div class="cube-blk" style="border-bottom:none;padding-bottom:.3rem">
          <h3>2 · Champs <span id="cubeNbF" style="color:#cbd5e1"></span></h3>
          <input class="cube-search" id="cubeFSearch" placeholder="🔍 Rechercher un champ…" oninput="CUBE.searchFields(this.value)">
        </div>
        <div class="cube-fields" id="cubeFields"></div>
      </aside>
      <section class="cube-main">
        <div class="cube-zones">
          <div class="cube-zone" data-zone="filters"><div class="zt">⏷ Filtres</div><div class="cube-pills" id="cubeZ_filters"></div></div>
          <div class="cube-zone" data-zone="cols"><div class="zt">⇥ <b>Colonnes</b></div><div class="cube-pills" id="cubeZ_cols"></div></div>
          <div class="cube-zone" data-zone="rows"><div class="zt">⇣ <b>Lignes</b></div><div class="cube-pills" id="cubeZ_rows"></div></div>
          <div class="cube-zone" data-zone="vals"><div class="zt">Σ <b>Valeurs</b></div><div class="cube-pills" id="cubeZ_vals"></div></div>
        </div>
        <div class="cube-opts" id="cubeOpts"></div>
        <div class="cube-result" id="cubeResult"></div>
      </section>
    </div>`;
  document.body.appendChild(ov);
  document.addEventListener('keydown', onKey);
  bindZones();
  renderSources();
  renderViews();
  loadData();
}
function close() { const o = $('cubeOverlay'); if (o) o.style.display = 'none'; closePop(); }
function onKey(e) {
  if (e.key !== 'Escape' || !$('cubeOverlay') || $('cubeOverlay').style.display === 'none') return;
  if (document.querySelector('.cube-modal')) { document.querySelector('.cube-modal').remove(); return; }
  if (document.querySelector('.cube-pop')) { closePop(); return; }
  close();
}

// ── sources & chargement ────────────────────────────────────────────────
function renderSources() {
  $('cubeSrc').innerHTML = CUBE_SRC.filter(s => !s.admin || S.canActivite).map(s =>
    `<span class="cube-chip ${S.sources.includes(s.k) ? 'on' : ''}" onclick="CUBE.toggleSource('${s.k}')">${s.i} ${s.l}</span>`).join('');
  $('cubeJoin').checked = S.join;
}
function toggleSource(k) {
  const i = S.sources.indexOf(k);
  if (i === -1) S.sources.push(k); else if (S.sources.length > 1) S.sources.splice(i, 1); else { toast('Gardez au moins une source', false); return; }
  renderSources(); loadData();
}
function setJoin(v) { S.join = !!v; loadData(); }

let _loadSeq = 0;
async function loadData(thenCfg) {
  const key = S.sources.slice().sort().join('+') + '|' + S.join;
  if (key === S.loadedKey && S.rows.length) { if (thenCfg) applyCfg(thenCfg); render(); return; }
  const seq = ++_loadSeq;
  const res = $('cubeResult');
  if (res) res.insertAdjacentHTML('beforeend', '<div class="cube-loading" id="cubeLoad"><div class="cube-spin"></div>Chargement des données…</div>');
  try {
    const r = await api('getCubeData', { sources: S.sources, join: S.join });
    if (seq !== _loadSeq) return;
    if (!r.success) { toast('Cube : ' + r.message, false); return; }
    S.fields = r.fields || []; S.rows = r.rows || []; S.loadedKey = key;
    if (typeof r.activiteAutorisee === 'boolean') S.canActivite = r.activiteAutorisee;
    // Retire les champs qui n'existent plus dans les zones
    const ok = n => !n || fieldIdx(n) !== -1;
    if (thenCfg) applyCfg(thenCfg);
    S.cfg.rows = S.cfg.rows.filter(ok); S.cfg.cols = S.cfg.cols.filter(ok);
    S.cfg.vals = S.cfg.vals.filter(v => ok(v.f)); if (!S.cfg.vals.length) S.cfg.vals = [{ f: '', agg: 'count' }];
    Object.keys(S.cfg.filters).forEach(n => { if (!ok(n)) delete S.cfg.filters[n]; });
    renderSources(); render();
  } catch (e) { console.error(e); toast('Erreur de chargement du cube', false); }
  finally { const l = $('cubeLoad'); if (l && seq === _loadSeq) l.remove(); }
}
function applyCfg(c) {
  S.cfg = { rows: (c.rows || []).slice(), cols: (c.cols || []).slice(), vals: (c.vals || [{ f: '', agg: 'count' }]).map(v => ({ ...v })), filters: JSON.parse(JSON.stringify(c.filters || {})) };
  if (c.opt) S.opt = { ...S.opt, ...c.opt };
  S.collapsed = new Set();
}

// ── champs (panneau gauche) ─────────────────────────────────────────────
function usedFields() {
  const u = new Set([...S.cfg.rows, ...S.cfg.cols, ...Object.keys(S.cfg.filters)]);
  S.cfg.vals.forEach(v => v.f && u.add(v.f));
  return u;
}
function renderFields() {
  const q = S.fieldSearch.toLowerCase();
  const used = usedFields();
  const list = S.fields.filter(f => !q || f.n.toLowerCase().includes(q));
  const groups = [
    { t: 'Champs', f: list.filter(f => !f.n.startsWith('Entreprise › ') && !f.n.includes(' · ')) },
    { t: 'Dates (année · trimestre · mois)', f: list.filter(f => f.n.includes(' · ') && !f.n.startsWith('Entreprise › ')) },
    { t: 'Fiche entreprise', f: list.filter(f => f.n.startsWith('Entreprise › ')) }
  ].filter(g => g.f.length);
  $('cubeNbF').textContent = S.fields.length ? '(' + S.fields.length + ')' : '';
  $('cubeFields').innerHTML = groups.map(g => `<div class="cube-group">${g.t}</div>` + g.f.map(f =>
    `<div class="cube-field" draggable="true" data-field="${esc(f.n)}" onclick="CUBE.fieldMenu(event,'${jsq(f.n)}')" title="Glisser vers une zone, ou cliquer">
       <span class="ty ${f.t}">${typeIcon(f.t)}</span><span style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(f.n.replace(/^Entreprise › /, ''))}</span>${used.has(f.n) ? '<span class="used"></span>' : ''}</div>`).join('')).join('')
    || '<p class="cube-hint" style="padding:1rem">Aucun champ.</p>';
  document.querySelectorAll('#cubeFields .cube-field').forEach(el => {
    el.addEventListener('dragstart', e => { e.dataTransfer.setData('text/plain', JSON.stringify({ field: el.dataset.field })); e.dataTransfer.effectAllowed = 'copyMove'; });
  });
}
const jsq = s => String(s).replace(/\\/g, '\\\\').replace(/'/g, "\\'");
function searchFields(v) { S.fieldSearch = v; renderFields(); }

function fieldMenu(ev, n) {
  ev.stopPropagation();
  const isNum = fieldType(n) === 'num';
  openPop(ev, `
    <div style="font-size:.68rem;font-weight:800;color:#94a3b8;padding:.3rem .6rem;text-transform:uppercase;letter-spacing:.08em">${esc(n)}</div>
    <div class="it" onclick="CUBE.addTo('rows','${jsq(n)}')">⇣ Ajouter en lignes</div>
    <div class="it" onclick="CUBE.addTo('cols','${jsq(n)}')">⇥ Ajouter en colonnes</div>
    <div class="it" onclick="CUBE.addTo('vals','${jsq(n)}')">Σ ${isNum ? 'Somme de ce champ' : 'Compter les valeurs distinctes'}</div>
    <div class="it" onclick="CUBE.addTo('filters','${jsq(n)}')">⏷ Filtrer sur ce champ</div>`);
}

// ── zones (glisser-déposer) ─────────────────────────────────────────────
function bindZones() {
  document.querySelectorAll('#cubeOverlay .cube-zone').forEach(z => {
    z.addEventListener('dragover', e => { e.preventDefault(); z.classList.add('over'); });
    z.addEventListener('dragleave', () => z.classList.remove('over'));
    z.addEventListener('drop', e => {
      e.preventDefault(); z.classList.remove('over');
      let d; try { d = JSON.parse(e.dataTransfer.getData('text/plain')); } catch (x) { return; }
      if (!d || !d.field) return;
      if (d.from) removeFrom(d.from, d.field, d.idx, true);
      addTo(z.dataset.zone, d.field, true);
    });
  });
}
function addTo(zone, n, silent) {
  closePop();
  if (zone === 'rows' || zone === 'cols') {
    S.cfg.rows = S.cfg.rows.filter(x => x !== n); S.cfg.cols = S.cfg.cols.filter(x => x !== n);
    S.cfg[zone].push(n);
  } else if (zone === 'vals') {
    const agg = fieldType(n) === 'num' ? 'sum' : 'countd';
    if (S.cfg.vals.length === 1 && !S.cfg.vals[0].f && S.cfg.vals[0].agg === 'count' && !silent) S.cfg.vals.push({ f: n, agg });
    else S.cfg.vals.push({ f: n, agg });
  } else if (zone === 'filters') {
    if (!S.cfg.filters[n]) S.cfg.filters[n] = [];
    render();
    setTimeout(() => { const el = document.querySelector(`#cubeZ_filters [data-f="${CSS.escape(n)}"]`); if (el) el.click(); }, 30);
    return;
  }
  S.collapsed = new Set();
  render();
}
function removeFrom(zone, n, idx, noRender) {
  if (zone === 'rows' || zone === 'cols') S.cfg[zone] = S.cfg[zone].filter(x => x !== n);
  else if (zone === 'vals') { S.cfg.vals.splice(idx, 1); if (!S.cfg.vals.length) S.cfg.vals = [{ f: '', agg: 'count' }]; }
  else if (zone === 'filters') delete S.cfg.filters[n];
  if (!noRender) render();
}
function setAgg(idx, agg) { S.cfg.vals[idx].agg = agg; render(); }
function swapAxes() { const r = S.cfg.rows; S.cfg.rows = S.cfg.cols; S.cfg.cols = r; S.collapsed = new Set(); render(); }

function renderZones() {
  const pill = (zone, n, cls, inner, idx) =>
    `<span class="cube-pill ${cls}" draggable="true" data-zone="${zone}" data-f="${esc(n)}" data-idx="${idx == null ? '' : idx}">${inner}<button class="x" onclick="event.stopPropagation();CUBE.removeFrom('${zone}','${jsq(n)}',${idx == null ? 'null' : idx})" title="Retirer">✕</button></span>`;
  ['rows', 'cols'].forEach(z => {
    $('cubeZ_' + z).innerHTML = S.cfg[z].map(n => pill(z, n, '', esc(n))).join('') || `<span class="cube-hint">Glissez un champ ici</span>`;
  });
  $('cubeZ_vals').innerHTML = S.cfg.vals.map((v, i) => {
    const opts = (v.f ? Object.keys(AGG) : ['count']).filter(a => v.f || a === 'count')
      .filter(a => fieldType(v.f) === 'num' || !v.f || ['count', 'countd'].includes(a))
      .map(a => `<option value="${a}" ${a === v.agg ? 'selected' : ''}>${AGG[a]}</option>`).join('');
    const inner = `<select onchange="CUBE.setAgg(${i},this.value)" onclick="event.stopPropagation()">${opts}</select>${v.f ? ' · ' + esc(v.f) : ' de lignes'}`;
    return pill('vals', v.f || '', 'val', inner, i);
  }).join('');
  $('cubeZ_filters').innerHTML = Object.keys(S.cfg.filters).map(n => {
    const sel = S.cfg.filters[n];
    return pill('filters', n, 'flt', `${esc(n)} ${sel.length ? `<span class="cnt">${sel.length}</span>` : '<span style="color:#94a3b8;font-weight:600">· tout</span>'}`);
  }).join('') || `<span class="cube-hint">Glissez un champ pour filtrer</span>`;
  document.querySelectorAll('#cubeOverlay .cube-pill').forEach(p => {
    p.addEventListener('dragstart', e => e.dataTransfer.setData('text/plain', JSON.stringify({ field: p.dataset.f, from: p.dataset.zone, idx: p.dataset.idx === '' ? null : Number(p.dataset.idx) })));
    if (p.dataset.zone === 'filters') p.addEventListener('click', e => filterPop(e, p.dataset.f));
  });
}

// ── filtres ─────────────────────────────────────────────────────────────
function distinctValues(n) {
  const i = fieldIdx(n); const m = new Map();
  S.rows.forEach(r => { const v = cellVal(r, i); m.set(v, (m.get(v) || 0) + 1); });
  return [...m.entries()].sort((a, b) => collator.compare(a[0], b[0]));
}
function filterPop(ev, n) {
  ev.stopPropagation();
  const vals = distinctValues(n);
  const sel = new Set(S.cfg.filters[n] || []);
  openPop(ev, `
    <div style="display:flex;align-items:center;gap:.4rem;padding:.2rem .3rem .4rem">
      <b style="font-size:.78rem;color:#1A4E8A;flex:1">${esc(n)}</b>
      <button class="cube-btn" style="padding:.2rem .5rem" onclick="CUBE.filterAll('${jsq(n)}',true)">Tout</button>
      <button class="cube-btn" style="padding:.2rem .5rem" onclick="CUBE.filterAll('${jsq(n)}',false)">Aucun</button>
    </div>
    <input class="cube-search" placeholder="Rechercher…" oninput="CUBE.filterSearch(this.value)">
    <div class="vals" id="cubeFVals">${vals.map(([v, c]) => `<label data-v="${esc(v.toLowerCase())}"><input type="checkbox" value="${esc(v)}" ${!sel.size || sel.has(v) ? 'checked' : ''} onchange="CUBE.filterChange('${jsq(n)}')"> <span style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(v)}</span><span class="n">${c}</span></label>`).join('')}</div>
    <p style="font-size:.66rem;color:#94a3b8;margin:.2rem .4rem">Tout coché = aucun filtre.</p>`);
}
function filterSearch(q) { q = q.toLowerCase(); document.querySelectorAll('#cubeFVals label').forEach(l => l.style.display = l.dataset.v.includes(q) ? '' : 'none'); }
function filterChange(n) {
  const all = [...document.querySelectorAll('#cubeFVals input')];
  const on = all.filter(i => i.checked).map(i => i.value);
  S.cfg.filters[n] = on.length === all.length ? [] : on;
  renderZones(); renderResult();
}
function filterAll(n, v) { document.querySelectorAll('#cubeFVals input').forEach(i => { if (i.closest('label').style.display !== 'none') i.checked = v; }); filterChange(n); }

// ── popovers ─────────────────────────────────────────────────────────────
function openPop(ev, html) {
  closePop();
  const p = document.createElement('div');
  p.className = 'cube-pop'; p.innerHTML = html;
  document.body.appendChild(p);
  const r = (ev.currentTarget || ev.target).getBoundingClientRect();
  let x = r.left, y = r.bottom + 6;
  const w = p.offsetWidth, h = p.offsetHeight;
  if (x + w > innerWidth - 8) x = innerWidth - w - 8;
  if (y + h > innerHeight - 8) y = Math.max(8, r.top - h - 6);
  p.style.left = x + 'px'; p.style.top = y + 'px';
  setTimeout(() => document.addEventListener('mousedown', _outside), 0);
}
function _outside(e) { const p = document.querySelector('.cube-pop'); if (p && !p.contains(e.target)) closePop(); }
function closePop() { document.querySelectorAll('.cube-pop').forEach(p => p.remove()); document.removeEventListener('mousedown', _outside); }

// ── options ─────────────────────────────────────────────────────────────
function renderOpts() {
  const seg = (key, items) => `<span class="cube-seg">${items.map(([v, l]) => `<button class="${S.opt[key] === v ? 'on' : ''}" onclick="CUBE.setOpt('${key}','${v}')">${l}</button>`).join('')}</span>`;
  $('cubeOpts').innerHTML = `
    ${seg('view', [['table', '▦ Tableau'], ['chart', '▤ Graphique']])}
    ${seg('show', [['val', 'Valeurs'], ['pctTotal', '% total'], ['pctRow', '% ligne'], ['pctCol', '% colonne']])}
    <label>Trier <select onchange="CUBE.setOpt('sort',this.value)">
      ${[['label', 'Libellé A→Z'], ['desc', 'Valeur ↓'], ['asc', 'Valeur ↑']].map(([v, l]) => `<option value="${v}" ${S.opt.sort === v ? 'selected' : ''}>${l}</option>`).join('')}</select></label>
    <label>Top <select onchange="CUBE.setOpt('top',+this.value)">
      ${[[0, 'Tous'], [5, '5'], [10, '10'], [20, '20'], [50, '50']].map(([v, l]) => `<option value="${v}" ${S.opt.top === v ? 'selected' : ''}>${l}</option>`).join('')}</select></label>
    <label><input type="checkbox" ${S.opt.heat ? 'checked' : ''} onchange="CUBE.setOpt('heat',this.checked)" style="accent-color:#1A4E8A"> Carte de chaleur</label>
    <label><input type="checkbox" ${S.opt.totals ? 'checked' : ''} onchange="CUBE.setOpt('totals',this.checked)" style="accent-color:#1A4E8A"> Totaux</label>
    <button class="cube-btn" style="padding:.3rem .6rem" onclick="CUBE.swapAxes()" title="Inverser lignes et colonnes">⇄ Inverser</button>
    ${S.cfg.rows.length > 1 ? `<button class="cube-btn" style="padding:.3rem .6rem" onclick="CUBE.collapseAll(true)">⊟ Replier</button><button class="cube-btn" style="padding:.3rem .6rem" onclick="CUBE.collapseAll(false)">⊞ Déplier</button>` : ''}
    <span class="cube-stat" id="cubeStat"></span>`;
}
function setOpt(k, v) { S.opt[k] = v; renderOpts(); renderResult(); }

// ── moteur de pivot ─────────────────────────────────────────────────────
function newAcc() { return { n: 0, s: 0, min: Infinity, max: -Infinity, set: null, cnt: 0 }; }
function addAcc(a, raw, agg) {
  a.cnt++;
  if (agg === 'countd') { (a.set || (a.set = new Set())).add(raw); return; }
  if (agg === 'count') return;
  const x = typeof raw === 'number' ? raw : parseFloat(String(raw).replace(',', '.'));
  if (isNaN(x)) return;
  a.n++; a.s += x; if (x < a.min) a.min = x; if (x > a.max) a.max = x;
}
function accVal(a, agg) {
  if (!a) return null;
  switch (agg) {
    case 'count': return a.cnt;
    case 'countd': return a.set ? a.set.size : 0;
    case 'sum': return a.s;
    case 'avg': return a.n ? a.s / a.n : null;
    case 'min': return a.n ? a.min : null;
    case 'max': return a.n ? a.max : null;
  }
  return null;
}

function filteredRows() {
  const flt = Object.entries(S.cfg.filters).filter(([, v]) => v && v.length).map(([n, v]) => [fieldIdx(n), new Set(v)]);
  if (!flt.length) return S.rows;
  return S.rows.filter(r => flt.every(([i, set]) => set.has(cellVal(r, i))));
}

function computePivot() {
  const rows = filteredRows();
  const ri = S.cfg.rows.map(fieldIdx), ci = S.cfg.cols.map(fieldIdx);
  const vals = S.cfg.vals.map(v => ({ ...v, i: v.f ? fieldIdx(v.f) : -1 }));
  const cells = new Map();     // rowPrefixKey + SEP2 + colKey → [acc per val]
  const rowKeys = new Map();   // level → Map(prefixKey → parts)
  const colKeys = new Map();   // colKey → parts
  const SEP2 = '\u0002';
  const touch = (k) => { let a = cells.get(k); if (!a) { a = vals.map(newAcc); cells.set(k, a); } return a; };
  for (const r of rows) {
    const rp = ri.map(i => cellVal(r, i));
    const cp = ci.map(i => cellVal(r, i));
    const ck = cp.join(SEP);
    if (ci.length && !colKeys.has(ck)) colKeys.set(ck, cp);
    const colTargets = ci.length ? [ck, '*'] : ['*'];
    for (let lv = 0; lv <= rp.length; lv++) {
      const pk = lv === 0 ? '' : rp.slice(0, lv).join(SEP);
      if (lv > 0) { let m = rowKeys.get(lv); if (!m) { m = new Map(); rowKeys.set(lv, m); } if (!m.has(pk)) m.set(pk, rp.slice(0, lv)); }
      for (const ct of colTargets) {
        const accs = touch(lv + ':' + pk + SEP2 + ct);
        vals.forEach((v, vi) => addAcc(accs[vi], v.i === -1 ? 1 : r[v.i], v.agg));
      }
    }
  }
  const get = (lv, pk, ck, vi) => { const a = cells.get(lv + ':' + pk + SEP2 + ck); return a ? accVal(a[vi], vals[vi].agg) : null; };
  // colonnes triées
  let cols = [...colKeys.entries()].sort((a, b) => {
    for (let k = 0; k < a[1].length; k++) { const c = collator.compare(a[1][k], b[1][k]); if (c) return c; }
    return 0;
  });
  return { rows, vals, get, cols, rowKeys, nLevels: ri.length };
}

// Arbre des lignes trié + top N
function rowTree(P) {
  const sortVal = (lv, pk) => P.get(lv, pk, '*', 0) || 0;
  const children = (lv, prefix) => {
    const m = P.rowKeys.get(lv); if (!m) return [];
    let list = [...m.entries()].filter(([k]) => lv === 1 || k.startsWith(prefix + SEP));
    if (S.opt.sort === 'label') list.sort((a, b) => collator.compare(a[1][lv - 1], b[1][lv - 1]));
    else list.sort((a, b) => S.opt.sort === 'desc' ? sortVal(lv, b[0]) - sortVal(lv, a[0]) : sortVal(lv, a[0]) - sortVal(lv, b[0]));
    if (S.opt.top && lv === 1) list = list.slice(0, S.opt.top);
    return list.map(([k, parts]) => ({ lv, key: k, label: parts[lv - 1], kids: lv < P.nLevels ? children(lv + 1, k) : [] }));
  };
  return children(1, '');
}

function display(P, raw, lv, pk, ck, vi) {
  if (raw == null) return '';
  const agg = P.vals[vi].agg;
  if (S.opt.show === 'val') return fmtNum(raw, agg === 'avg' ? 2 : (Number.isInteger(raw) ? 0 : 2));
  let den = null;
  if (S.opt.show === 'pctTotal') den = P.get(0, '', '*', vi);
  if (S.opt.show === 'pctRow') den = P.get(lv, pk, '*', vi);
  if (S.opt.show === 'pctCol') den = P.get(0, '', ck, vi);
  if (!den) return '';
  return fmtNum(raw / den * 100, 1) + ' %';
}

// ── rendu ─────────────────────────────────────────────────────────────
function render() { renderFields(); renderZones(); renderOpts(); renderResult(); }

function renderResult() {
  const el = $('cubeResult');
  if (!el) return;
  if (!S.rows.length && !S.fields.length) { el.innerHTML = `<div class="cube-empty"><div class="cube-spin"></div></div>`; return; }
  if (!S.cfg.rows.length && !S.cfg.cols.length) {
    const presets = PRESETS.filter(p => !p.admin || S.canActivite);
    el.innerHTML = `<div class="cube-empty">
      <div class="big">🧊</div>
      <div style="font-weight:800;color:#1A4E8A;font-size:1.05rem">Construisez votre analyse</div>
      <div style="font-size:.82rem;max-width:520px">Glissez des champs de la colonne de gauche vers <b>Lignes</b> et <b>Colonnes</b>, choisissez ce que vous mesurez dans <b>Valeurs</b>. Cliquez sur une cellule pour voir le détail.</div>
      <div style="font-size:.68rem;font-weight:800;text-transform:uppercase;letter-spacing:.1em;margin-top:.8rem">Ou partez d'un modèle</div>
      <div class="cube-presets">${presets.map((p, i) => `<button class="cube-btn" onclick="CUBE.preset(${PRESETS.indexOf(p)})">${esc(p.nom)}</button>`).join('')}</div>
    </div>`;
    const st = $('cubeStat'); if (st) st.textContent = S.rows.length.toLocaleString('fr') + ' lignes chargées';
    return;
  }
  const P = computePivot();
  S.lastPivot = P;
  const st = $('cubeStat');
  if (st) st.textContent = P.rows.length.toLocaleString('fr') + ' / ' + S.rows.length.toLocaleString('fr') + ' lignes';
  if (S.opt.view === 'chart') { el.innerHTML = renderChart(P); return; }
  el.innerHTML = renderTable(P);
}

function renderTable(P) {
  const tree = rowTree(P);
  const nv = P.vals.length;
  const colList = P.cols.length ? P.cols : [];
  const hasCols = colList.length > 0;
  const showColTot = S.opt.totals && hasCols;
  // max pour la carte de chaleur (cellules feuilles uniquement)
  let maxV = [];
  if (S.opt.heat) {
    maxV = P.vals.map((_, vi) => {
      let m = 0;
      const leafLv = Math.max(P.nLevels, 0);
      const keys = leafLv ? [...(P.rowKeys.get(leafLv) || new Map()).keys()] : [''];
      keys.forEach(k => (hasCols ? colList.map(c => c[0]) : ['*']).forEach(ck => { const v = P.get(leafLv, k, ck, vi); if (v > m) m = v; }));
      return m;
    });
  }
  const heat = (v, vi, isLeaf) => {
    if (!S.opt.heat || !isLeaf || v == null || !maxV[vi]) return '';
    const a = Math.max(0, Math.min(1, v / maxV[vi]));
    if (a <= 0) return '';
    return `background:rgba(26,78,138,${(0.06 + a * 0.5).toFixed(3)});${a > 0.62 ? 'color:#fff;' : ''}`;
  };
  const valLabel = v => (AGG[v.agg] || '') + (v.f ? ' · ' + v.f : '');
  const rowHeadTitle = S.cfg.rows.length ? S.cfg.rows.join(' › ') : '';

  // En-têtes
  let thead = '';
  if (hasCols) {
    const colLbl = S.cfg.cols.join(' › ');
    thead += `<tr><th class="rh" rowspan="${nv > 1 ? 2 : 1}">${esc(rowHeadTitle)}<div style="font-weight:600;color:#94a3b8;font-size:.62rem">colonnes : ${esc(colLbl)}</div></th>`;
    colList.forEach(([, parts]) => { thead += `<th colspan="${nv}" style="text-align:center">${esc(parts.join(' › '))}</th>`; });
    if (showColTot) thead += `<th colspan="${nv}" class="ct" style="text-align:center">Total</th>`;
    thead += '</tr>';
    if (nv > 1) {
      thead += '<tr>';
      colList.concat(showColTot ? [[null]] : []).forEach(() => P.vals.forEach(v => { thead += `<th>${esc(valLabel(v))}</th>`; }));
      thead += '</tr>';
    }
  } else {
    thead = `<tr><th class="rh">${esc(rowHeadTitle || 'Total')}</th>${P.vals.map(v => `<th>${esc(valLabel(v))}</th>`).join('')}</tr>`;
  }

  // Corps
  const cellsFor = (lv, pk, isLeaf) => {
    let h = '';
    const targets = hasCols ? colList.map(c => c[0]) : ['*'];
    targets.forEach(ck => P.vals.forEach((v, vi) => {
      const raw = P.get(lv, pk, ck, vi);
      h += `<td style="${heat(raw, vi, isLeaf)}" data-lv="${lv}" data-pk="${esc(pk)}" data-ck="${esc(ck)}">${display(P, raw, lv, pk, ck, vi)}</td>`;
    }));
    if (showColTot) P.vals.forEach((v, vi) => { const raw = P.get(lv, pk, '*', vi); h += `<td class="ct" data-lv="${lv}" data-pk="${esc(pk)}" data-ck="*">${display(P, raw, lv, pk, '*', vi)}</td>`; });
    return h;
  };
  let body = '';
  const walk = (nodes) => nodes.forEach(n => {
    const isLeaf = !n.kids.length;
    const col = S.collapsed.has(n.key);
    const pad = (n.lv - 1) * 18;
    if (isLeaf) body += `<tr><th style="padding-left:${12 + pad}px" title="${esc(n.label)}">${esc(n.label)}</th>${cellsFor(n.lv, n.key, true)}</tr>`;
    else {
      body += `<tr class="sub"><th style="padding-left:${12 + pad}px" onclick="CUBE.toggleRow('${jsq(n.key)}')" title="${esc(n.label)}"><span class="cube-tog">${col ? '▸' : '▾'}</span>${esc(n.label)} <span style="color:#94a3b8;font-weight:600;font-size:.7rem">(${n.kids.length})</span></th>${cellsFor(n.lv, n.key, false)}</tr>`;
      if (!col) walk(n.kids);
    }
  });
  if (P.nLevels) walk(tree);
  if (S.opt.totals || !P.nLevels) body += `<tr class="tot"><th>Total général</th>${cellsFor(0, '', !P.nLevels)}</tr>`;
  return `<table class="cube-t" id="cubeTable"><thead>${thead}</thead><tbody>${body}</tbody></table>`;
}

const PALETTE = ['#1A4E8A', '#F26522', '#10b981', '#8b5cf6', '#06b6d4', '#ec4899', '#f59e0b', '#6366f1', '#14b8a6', '#ef4444', '#84cc16', '#64748b'];
function renderChart(P) {
  const tree = P.nLevels ? rowTree(P) : [{ lv: 0, key: '', label: 'Total', kids: [] }];
  const vi = 0;
  const hasCols = P.cols.length > 0;
  const colsShown = hasCols ? P.cols.slice(0, 12) : [];
  const items = tree.map(n => {
    const lv = n.lv;
    const total = P.get(lv, n.key, '*', vi) || 0;
    const segs = hasCols ? colsShown.map(([ck], i) => ({ v: P.get(lv, n.key, ck, vi) || 0, c: PALETTE[i % PALETTE.length] })) : [{ v: total, c: PALETTE[0] }];
    return { label: n.label, total, segs };
  });
  const max = Math.max(1, ...items.map(i => hasCols ? i.segs.reduce((s, x) => s + x.v, 0) : i.total));
  const v0 = P.vals[0];
  return `<div style="padding:1rem 1.4rem 0;font-size:.75rem;font-weight:800;color:#94a3b8;text-transform:uppercase;letter-spacing:.08em">${esc((AGG[v0.agg] || '') + (v0.f ? ' · ' + v0.f : ' de lignes'))} par ${esc(S.cfg.rows[0] || '—')}${hasCols ? ' × ' + esc(S.cfg.cols.join(' › ')) : ''}</div>
    <div class="cube-bars">${items.map(it => `<div class="cube-bar"><div class="lb" title="${esc(it.label)}">${esc(it.label)}</div>
      <div class="tr">${it.segs.map(s => `<div class="sg" title="${fmtNum(s.v)}" style="width:${(s.v / max * 100).toFixed(2)}%;background:${s.c}"></div>`).join('')}</div>
      <div class="vl">${fmtNum(it.total)}</div></div>`).join('')}</div>
    ${hasCols ? `<div class="cube-legend">${colsShown.map(([, p], i) => `<span><i style="background:${PALETTE[i % PALETTE.length]}"></i>${esc(p.join(' › '))}</span>`).join('')}${P.cols.length > 12 ? '<span>… (12 premières colonnes)</span>' : ''}</div>` : ''}`;
}

function toggleRow(k) { if (S.collapsed.has(k)) S.collapsed.delete(k); else S.collapsed.add(k); renderResult(); }
function collapseAll(v) {
  S.collapsed = new Set();
  if (v && S.lastPivot) for (let lv = 1; lv < S.lastPivot.nLevels; lv++) (S.lastPivot.rowKeys.get(lv) || new Map()).forEach((_, k) => S.collapsed.add(k));
  renderResult();
}

// ── détail d'une cellule ────────────────────────────────────────────────
function onCellClick(e) {
  const td = e.target.closest('#cubeTable td'); if (!td) return;
  const lv = Number(td.dataset.lv), pk = td.dataset.pk, ck = td.dataset.ck;
  const rp = lv ? pk.split(SEP) : [], cp = ck && ck !== '*' ? ck.split(SEP) : [];
  const ri = S.cfg.rows.slice(0, lv).map(fieldIdx), ci = S.cfg.cols.map(fieldIdx);
  const list = filteredRows().filter(r => ri.every((i, k) => cellVal(r, i) === rp[k]) && (!cp.length || ci.every((i, k) => cellVal(r, i) === cp[k])));
  const crit = S.cfg.rows.slice(0, lv).map((n, k) => `${n} = ${rp[k]}`).concat(cp.length ? S.cfg.cols.map((n, k) => `${n} = ${cp[k]}`) : []);
  const showF = S.fields.filter(f => !f.n.includes(' · ')).slice(0, 14);
  const idx = showF.map(f => fieldIdx(f.n));
  const m = document.createElement('div');
  m.className = 'cube-modal';
  m.onclick = ev => { if (ev.target === m) m.remove(); };
  m.innerHTML = `<div class="box" style="max-width:1100px">
    <div style="display:flex;align-items:start;gap:1rem"><div style="flex:1"><h3>🔎 Détail — ${list.length.toLocaleString('fr')} ligne(s)</h3>
    <p style="font-size:.75rem;color:#64748b;margin:0">${esc(crit.join(' · ') || 'Toutes les lignes filtrées')}</p></div>
    <button class="cube-btn" onclick="CUBE.exportDetail()">⬇ CSV</button><button class="cube-btn" onclick="this.closest('.cube-modal').remove()">✕</button></div>
    <div style="overflow:auto;max-height:65vh;margin-top:1rem;border:1px solid #eef2f7;border-radius:12px">
    <table class="cube-t"><thead><tr>${showF.map(f => `<th style="text-align:left">${esc(f.n)}</th>`).join('')}</tr></thead>
    <tbody>${list.slice(0, 500).map(r => `<tr>${idx.map(i => `<td style="text-align:left;cursor:default">${esc(r[i])}</td>`).join('')}</tr>`).join('')}</tbody></table></div>
    ${list.length > 500 ? '<p style="font-size:.72rem;color:#94a3b8;margin:.5rem 0 0">500 premières lignes affichées — le CSV contient tout.</p>' : ''}
  </div>`;
  document.body.appendChild(m);
  S._detail = { list, fields: showF.map(f => f.n), idx };
}

// ── vues, modèles, export ───────────────────────────────────────────────
function getViews() { try { return JSON.parse(localStorage.getItem(VIEWS_KEY) || '[]'); } catch (e) { return []; } }
function renderViews() {
  const v = getViews();
  $('cubeViews').innerHTML = `<option value="">📚 Vues enregistrées (${v.length})</option>` + v.map((x, i) => `<option value="${i}">${esc(x.nom)}</option>`).join('') + (v.length ? '<option value="__del">🗑 Supprimer une vue…</option>' : '');
}
function currentConfig(nom) { return { nom, sources: S.sources.slice(), join: S.join, rows: S.cfg.rows.slice(), cols: S.cfg.cols.slice(), vals: S.cfg.vals.map(v => ({ f: v.f, agg: v.agg })), filters: JSON.parse(JSON.stringify(S.cfg.filters)), opt: { ...S.opt }, show: S.opt.show }; }
function saveView() {
  const nom = prompt('Nom de la vue :', S.cfg.rows.concat(S.cfg.cols).join(' × ') || 'Ma vue');
  if (!nom) return;
  const v = getViews().filter(x => x.nom !== nom); v.push(currentConfig(nom));
  localStorage.setItem(VIEWS_KEY, JSON.stringify(v)); renderViews(); toast('Vue « ' + nom + ' » enregistrée ✓');
}
function loadView(i) {
  const sel = $('cubeViews');
  if (i === '__del') {
    const v = getViews(); const nom = prompt('Nom exact de la vue à supprimer :\n' + v.map(x => '• ' + x.nom).join('\n'));
    if (nom) { localStorage.setItem(VIEWS_KEY, JSON.stringify(v.filter(x => x.nom !== nom))); toast('Vue supprimée'); }
    renderViews(); return;
  }
  if (i === '') return;
  const c = getViews()[Number(i)]; sel.value = '';
  if (c) useConfig(c);
}
function useConfig(c) { S.sources = (c.sources || ['etudiants']).filter(s => s !== 'activite' || S.canActivite); if (!S.sources.length) S.sources = ['etudiants']; S.join = !!c.join; renderSources(); loadData(c); }
function preset(i) { useConfig(PRESETS[i]); }
function reset() { S.cfg = { rows: [], cols: [], vals: [{ f: '', agg: 'count' }], filters: {} }; S.collapsed = new Set(); render(); }

function csvLine(arr) { return arr.map(v => { const s = String(v == null ? '' : v); return /[";\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; }).join(';'); }
function download(name, text) {
  const blob = new Blob(['﻿' + text], { type: 'text/csv;charset=utf-8' });
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 2000);
}
function exportCsv() {
  const t = $('cubeTable');
  if (!t) { toast('Construisez d\'abord un tableau (vue Tableau)', false); return; }
  const lines = [...t.querySelectorAll('tr')].map(tr => csvLine([...tr.children].flatMap(c => { const span = Number(c.getAttribute('colspan') || 1); const txt = c.textContent.replace(/[▾▸]/g, '').trim(); return [txt].concat(Array(span - 1).fill('')); })));
  download('CUBE_HECG_' + new Date().toISOString().slice(0, 10) + '.csv', lines.join('\n'));
}
function exportDetail() {
  const d = S._detail; if (!d) return;
  download('CUBE_detail_' + new Date().toISOString().slice(0, 10) + '.csv', [csvLine(d.fields)].concat(d.list.map(r => csvLine(d.idx.map(i => r[i])))).join('\n'));
}

function openExport() {
  if (!S.cfg.rows.length && !S.cfg.cols.length) { toast('Ajoutez au moins un champ en lignes ou en colonnes', false); return; }
  const mine = (typeof tableauxSuivi !== 'undefined' ? tableauxSuivi : []).filter(t => !t.masque && (t.estAuteur || vieScRole === 'Super-admin'));
  const m = document.createElement('div');
  m.className = 'cube-modal';
  m.onclick = ev => { if (ev.target === m) m.remove(); };
  m.innerHTML = `<div class="box">
    <h3>📤 Envoyer vers Google Sheet</h3>
    <p style="font-size:.78rem;color:#64748b;margin:0">Crée un <b>vrai tableau croisé dynamique Google Sheets</b> avec cette configuration, alimenté par un onglet de données masqué. Il est <b>reconstruit automatiquement</b> à chaque synchronisation (toutes les heures, ou bouton 🔄).</p>
    <label class="l">Tableau de suivi</label>
    <select id="cubeExpTab">${mine.map(t => `<option value="${esc(t.id)}">${esc(t.nom)}</option>`).join('')}<option value="__new">➕ Nouveau tableau de suivi…</option></select>
    <div id="cubeExpNewWrap" style="display:${mine.length ? 'none' : 'block'}"><label class="l">Nom du nouveau tableau</label><input type="text" id="cubeExpNew" placeholder="Ex : Analyses CUBE 2026-2027"></div>
    <label class="l">Nom de l'onglet</label>
    <input type="text" id="cubeExpNom" value="${esc((S.cfg.rows.concat(S.cfg.cols)).join(' × ').slice(0, 60))}">
    <div style="background:#f8fafc;border-radius:12px;padding:.7rem .85rem;margin-top:1rem;font-size:.74rem;color:#475569;line-height:1.5">
      <b>Sources :</b> ${esc(S.sources.map(s => (CUBE_SRC.find(x => x.k === s) || {}).l).join(' + '))}${S.join ? ' (+ fiche entreprise)' : ''}<br>
      <b>Lignes :</b> ${esc(S.cfg.rows.join(', ') || '—')} · <b>Colonnes :</b> ${esc(S.cfg.cols.join(', ') || '—')}<br>
      <b>Valeurs :</b> ${esc(S.cfg.vals.map(v => AGG[v.agg] + (v.f ? ' · ' + v.f : '')).join(', '))}<br>
      <b>Filtres :</b> ${esc(Object.entries(S.cfg.filters).filter(([, v]) => v.length).map(([n, v]) => n + ' (' + v.length + ')').join(', ') || 'aucun')}
    </div>
    <div style="display:flex;justify-content:flex-end;gap:.5rem;margin-top:1.2rem">
      <button class="cube-btn" onclick="this.closest('.cube-modal').remove()">Annuler</button>
      <button class="cube-btn ora" id="cubeExpBtn" onclick="CUBE.doExport()">Créer le tableau croisé</button>
    </div></div>`;
  document.body.appendChild(m);
  const sel = $('cubeExpTab');
  if (!mine.length) sel.value = '__new';
  sel.onchange = () => { $('cubeExpNewWrap').style.display = sel.value === '__new' ? 'block' : 'none'; };
}
async function doExport() {
  const btn = $('cubeExpBtn');
  btn.disabled = true; btn.textContent = 'Création en cours…';
  try {
    let id = $('cubeExpTab').value;
    if (id === '__new') {
      const nom = $('cubeExpNew').value.trim();
      if (!nom) { toast('Donnez un nom au nouveau tableau', false); return; }
      const r = await api('createTableauSuivi', { nom });
      if (!r.success) { toast('Erreur : ' + r.message, false); return; }
      id = r.id;
    }
    const cfg = currentConfig($('cubeExpNom').value.trim() || 'Cube');
    cfg.filters = Object.fromEntries(Object.entries(cfg.filters).filter(([, v]) => v.length));
    const r = await api('exportCubeToSheet', { id, config: cfg });
    if (!r.success) { toast('Erreur : ' + r.message, false); return; }
    document.querySelector('.cube-modal')?.remove();
    toast(`Tableau croisé « ${r.onglet} » créé (${r.lignes} lignes) ✓`);
    if (r.url) window.open(r.url, '_blank');
    if (typeof loadTableauxSuivi === 'function') loadTableauxSuivi();
  } finally { btn.disabled = false; btn.textContent = 'Créer le tableau croisé'; }
}

document.addEventListener('click', e => { if (e.target.closest && e.target.closest('#cubeTable td')) onCellClick(e); });

window.CUBE = { open, close, toggleSource, setJoin, searchFields, fieldMenu, addTo, removeFrom, setAgg, swapAxes, filterSearch, filterChange, filterAll,
  setOpt, toggleRow, collapseAll, saveView, loadView, preset, reset, exportCsv, exportDetail, openExport, doExport };
})();
