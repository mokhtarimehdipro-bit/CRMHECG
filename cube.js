/* =============================================================================
   🧊 CUBE v2 — Analyse croisée dynamique (Vie scolaire → Tableaux de suivis)
   -----------------------------------------------------------------------------
   Interface : Plus Jakarta Sans, icônes vectorielles, tuiles de sources,
   ajout rapide d'un champ en un clic (Lignes / Colonnes / Valeur), glisser-déposer,
   bandeau d'indicateurs, tableau à sous-totaux repliables, carte de chaleur,
   graphique, détail d'une cellule, vues enregistrées, modèles, export CSV et
   « Envoyer vers Google Sheets » (vrai tableau croisé dynamique, backend cube.gs).
   Dépend de vie_scolaire.html : api(), toast(), tableauxSuivi, vieScRole, user.
   ========================================================================== */
(function () {
'use strict';

// ── icônes (traits 24×24) ───────────────────────────────────────────────
const IC = {
  cube: 'M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z M3.3 7 12 12l8.7-5 M12 22V12',
  rows: 'M3 6h18 M3 12h18 M3 18h18',
  cols: 'M6 3v18 M12 3v18 M18 3v18',
  sigma: 'M18 7V4H6l6 8-6 8h12v-3',
  filter: 'M22 3H2l8 9.46V19l4 2v-8.54L22 3z',
  search: 'M11 4a7 7 0 1 0 0 14 7 7 0 1 0 0-14z M21 21l-4.35-4.35',
  save: 'M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z M17 21v-8H7v8 M7 3v5h8',
  reset: 'M3 12a9 9 0 1 0 3-6.7L3 8 M3 3v5h5',
  download: 'M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4 M7 10l5 5 5-5 M12 15V3',
  send: 'M22 2 11 13 M22 2l-7 20-4-9-9-4 20-7z',
  x: 'M18 6 6 18 M6 6l12 12',
  table: 'M3 3h18v18H3z M3 9h18 M3 15h18 M9 3v18',
  chart: 'M3 3v18h18 M7 16V10 M12 16V6 M17 16v-3',
  swap: 'M7 16V4 M3 8l4-4 4 4 M17 8v12 M21 16l-4 4-4-4',
  plus: 'M12 5v14 M5 12h14',
  grad: 'M22 10 12 5 2 10l10 5 10-5z M6 12v5c3 2 9 2 12 0v-5',
  build: 'M4 21V5l8-3v19 M12 21V9l8 3v9 M2 21h20',
  user: 'M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2 M12 3a4 4 0 1 0 0 8 4 4 0 1 0 0-8z',
  brief: 'M3 7h18v13H3z M8 7V5a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2',
  mega: 'M3 11v2a1 1 0 0 0 1 1h3l5 4V6L7 10H4a1 1 0 0 0-1 1z M16 9a4 4 0 0 1 0 6',
  pulse: 'M22 12h-4l-3 9L9 3l-3 9H2',
  check: 'M20 6 9 17l-5-5',
  chevD: 'm6 9 6 6 6-6', chevR: 'm9 6 6 6-6 6',
  book: 'M4 19.5A2.5 2.5 0 0 1 6.5 17H20V3H6.5A2.5 2.5 0 0 0 4 5.5z',
  spark: 'M12 3l1.9 5.8L20 10l-6.1 1.2L12 17l-1.9-5.8L4 10l6.1-1.2z',
  alert: 'M12 9v4 M12 17h.01 M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z'
};
const svg = (k, s = 16, w = 2) => `<svg width="${s}" height="${s}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="${w}" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="${IC[k]}"/></svg>`;

const CUBE_SRC = [
  { k: 'etudiants',   l: 'Étudiants',    ic: 'grad',  c: '#1A4E8A' },
  { k: 'entreprises', l: 'Entreprises',  ic: 'build', c: '#6366f1' },
  { k: 'contacts',    l: 'Contacts',     ic: 'user',  c: '#0891b2' },
  { k: 'offres',      l: 'Offres',       ic: 'brief', c: '#059669' },
  { k: 'campagnes',   l: 'Campagnes',    ic: 'mega',  c: '#F26522' },
  { k: 'activite',    l: 'Activité CRM', ic: 'pulse', c: '#db2777', admin: true }
];
const AGG = { count: 'Nombre', countd: 'Nombre distinct', sum: 'Somme', avg: 'Moyenne', min: 'Min', max: 'Max' };
const VIDE = '(vide)';
const VIEWS_KEY = 'hecg_cube_views_v1';
const PRESETS = [
  { nom: 'Étudiants par classe et statut', d: 'Répartition des statuts dans chaque classe', ic: 'grad', sources: ['etudiants'], join: false, rows: ['Classe'], cols: ['Statut'], vals: [{ f: '', agg: 'count' }], filters: {} },
  { nom: 'Alternants par territoire', d: 'Département et ville des entreprises d\'accueil', ic: 'build', sources: ['etudiants'], join: true, rows: ['Entreprise › Département', 'Entreprise › Ville'], cols: ['Campus'], vals: [{ f: '', agg: 'count' }], filters: {} },
  { nom: 'Ateliers CV & LinkedIn', d: 'Taux de participation par classe', ic: 'check', sources: ['etudiants'], join: false, rows: ['Classe'], cols: ['Atelier CV'], vals: [{ f: '', agg: 'count' }], filters: {}, opt: { show: 'pctRow' } },
  { nom: 'Offres par mois et qualité', d: 'Flux d\'offres dans le temps', ic: 'brief', sources: ['offres'], join: false, rows: ['Date · Mois'], cols: ['Qualité'], vals: [{ f: '', agg: 'count' }], filters: {}, opt: { view: 'chart' } },
  { nom: 'Entreprises par département', d: 'Nombre d\'entreprises et d\'alternants', ic: 'build', sources: ['entreprises'], join: false, rows: ['Département'], cols: [], vals: [{ f: '', agg: 'count' }, { f: 'Nb alternants', agg: 'sum' }], filters: {}, opt: { sort: 'desc' } },
  { nom: 'Entreprises + offres par ville', d: 'Croisement de deux sources', ic: 'spark', sources: ['entreprises', 'offres'], join: false, rows: ['Ville'], cols: ['Source'], vals: [{ f: '', agg: 'count' }], filters: {}, opt: { sort: 'desc', top: 20 } },
  { nom: 'Activité de l\'équipe', d: 'Actions par personne et par mois', ic: 'pulse', sources: ['activite'], join: false, rows: ['Auteur'], cols: ['Date · Mois'], vals: [{ f: '', agg: 'count' }], filters: {}, admin: true }
];
const OPT_DEFAULT = { show: 'val', heat: true, totals: true, sort: 'label', top: 0, view: 'table' };

const S = {
  sources: ['etudiants'], join: false,
  fields: [], rows: [], loadedKey: '', error: '', loading: false,
  cfg: { rows: [], cols: [], vals: [{ f: '', agg: 'count' }], filters: {} },
  opt: { ...OPT_DEFAULT },
  collapsed: new Set(), fieldSearch: '', canActivite: false, lastPivot: null
};

// ── utilitaires ──────────────────────────────────────────────────────────
const $ = id => document.getElementById(id);
const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const jsq = s => String(s).replace(/\\/g, '\\\\').replace(/'/g, "\\'");
const SEP = '\u0001';
const collator = new Intl.Collator('fr', { numeric: true, sensitivity: 'base' });
const fmtNum = (v, dec) => v == null || isNaN(v) ? '' : Number(v).toLocaleString('fr-FR', { maximumFractionDigits: dec == null ? 2 : dec });
const fieldIdx = n => S.fields.findIndex(f => f.n === n);
const fieldType = n => (S.fields.find(f => f.n === n) || {}).t || 'txt';
const typeIcon = t => t === 'num' ? '123' : (t === 'date' ? 'DATE' : 'ABC');
const cellVal = (row, i) => { const v = row[i]; return v === '' || v == null ? VIDE : String(v); };
const shortName = n => n.replace(/^Entreprise › /, '');
const valLabel = v => (AGG[v.agg] || '') + (v.f ? ' · ' + v.f : '');

// ── styles ───────────────────────────────────────────────────────────────
function injectCss() {
  if ($('cubeCss')) return;
  const ln = document.createElement('link');
  ln.rel = 'stylesheet';
  ln.href = 'https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap';
  document.head.appendChild(ln);
  const st = document.createElement('style');
  st.id = 'cubeCss';
  st.textContent = `
  #cubeOverlay{--b:#1A4E8A;--b6:#164378;--b5:#2a62a6;--b1:#e8eff8;--b0:#f3f7fc;--o:#F26522;--o1:#fff1e8;
    --ink:#0e1a2b;--tx:#3b4859;--mu:#7d8a9c;--li:#e7ecf3;--li2:#f0f3f8;--bg:#f4f6fa;
    position:fixed;inset:0;z-index:9000;display:flex;flex-direction:column;color:var(--tx);
    background:radial-gradient(1100px 520px at 100% -10%,rgba(26,78,138,.07),transparent 60%),radial-gradient(800px 420px at -5% 105%,rgba(242,101,34,.05),transparent 60%),var(--bg);
    font-family:'Plus Jakarta Sans',Inter,ui-sans-serif,system-ui,sans-serif;font-size:13.5px;line-height:1.45;
    -webkit-font-smoothing:antialiased;font-feature-settings:'tnum' 0;animation:cbIn .3s cubic-bezier(.2,.7,.2,1)}
  @keyframes cbIn{from{opacity:0;transform:translateY(10px) scale(.995)}to{opacity:1;transform:none}}
  #cubeOverlay *,#cubeOverlay *::before,#cubeOverlay *::after{box-sizing:border-box}
  #cubeOverlay button{font-family:inherit}
  #cubeOverlay svg,.cb-pop svg,.cb-modal svg{flex-shrink:0;display:inline-block;vertical-align:-2px}

  /* barre supérieure */
  .cb-top{height:64px;display:flex;align-items:center;gap:14px;padding:0 22px;background:rgba(255,255,255,.78);backdrop-filter:saturate(180%) blur(16px);border-bottom:1px solid var(--li)}
  .cb-logo{width:38px;height:38px;border-radius:12px;display:flex;align-items:center;justify-content:center;color:#fff;background:linear-gradient(135deg,#2a62a6,#1A4E8A 55%,#123865);box-shadow:0 8px 18px -8px rgba(26,78,138,.8),inset 0 1px 0 rgba(255,255,255,.25)}
  .cb-title{display:flex;flex-direction:column;line-height:1.15}
  .cb-title b{font-size:17px;font-weight:800;color:var(--ink);letter-spacing:-.02em}
  .cb-title span{font-size:12px;color:var(--mu);font-weight:500}
  .cb-sp{flex:1}
  .cb-btn{display:inline-flex;align-items:center;gap:7px;height:36px;padding:0 13px;border-radius:11px;border:1px solid var(--li);background:#fff;color:var(--tx);font-size:13px;font-weight:600;cursor:pointer;transition:all .18s;white-space:nowrap}
  .cb-btn:hover{border-color:#cbd6e4;background:var(--b0);color:var(--b)}
  .cb-btn:active{transform:scale(.98)}
  .cb-btn.ico{width:36px;padding:0;justify-content:center}
  .cb-btn.pri{background:var(--b);border-color:var(--b);color:#fff;box-shadow:0 6px 16px -8px rgba(26,78,138,.8)}
  .cb-btn.pri:hover{background:var(--b6);color:#fff}
  .cb-btn.ora{background:linear-gradient(180deg,#f47637,#F26522);border-color:#e85d1c;color:#fff;box-shadow:0 8px 18px -8px rgba(242,101,34,.8),inset 0 1px 0 rgba(255,255,255,.2)}
  .cb-btn.ora:hover{background:#dc5714;color:#fff}
  .cb-btn:disabled{opacity:.5;cursor:not-allowed}
  .cb-sel{height:36px;border-radius:11px;border:1px solid var(--li);background:#fff;padding:0 30px 0 12px!important;font:600 13px 'Plus Jakarta Sans',Inter,sans-serif;color:var(--tx);cursor:pointer}
  .cb-div{width:1px;height:24px;background:var(--li)}

  /* structure */
  .cb-body{flex:1;display:grid;grid-template-columns:316px 1fr;min-height:0}
  .cb-side{display:flex;flex-direction:column;min-height:0;background:rgba(255,255,255,.72);border-right:1px solid var(--li)}
  .cb-sec{padding:18px 18px 14px}
  .cb-lbl{display:flex;align-items:center;gap:8px;font-size:11px;font-weight:700;letter-spacing:.09em;text-transform:uppercase;color:var(--mu);margin-bottom:11px}
  .cb-lbl .n{margin-left:auto;font-size:11px;letter-spacing:0;background:var(--li2);color:var(--mu);border-radius:999px;padding:1px 8px;text-transform:none}

  /* sources */
  .cb-srcs{display:grid;grid-template-columns:1fr 1fr;gap:8px}
  .cb-src{position:relative;display:flex;align-items:center;gap:9px;padding:10px 10px;border-radius:13px;border:1px solid var(--li);background:#fff;cursor:pointer;transition:all .18s;font-size:12.5px;font-weight:600;color:var(--tx);user-select:none}
  .cb-src:hover{border-color:#cbd6e4;transform:translateY(-1px);box-shadow:0 6px 14px -10px rgba(14,26,43,.35)}
  .cb-src .i{width:28px;height:28px;border-radius:9px;display:flex;align-items:center;justify-content:center;background:color-mix(in srgb,var(--c) 12%,#fff);color:var(--c)}
  .cb-src.on{border-color:var(--c);background:color-mix(in srgb,var(--c) 6%,#fff);box-shadow:0 0 0 3px color-mix(in srgb,var(--c) 14%,transparent)}
  .cb-src.on .i{background:var(--c);color:#fff}
  .cb-src .ck{position:absolute;top:6px;right:6px;width:15px;height:15px;border-radius:50%;background:var(--c);color:#fff;display:none;align-items:center;justify-content:center}
  .cb-src.on .ck{display:flex}
  .cb-switch{display:flex;align-items:center;gap:10px;margin-top:12px;padding:10px 12px;border-radius:12px;background:var(--b0);cursor:pointer;font-size:12.5px;font-weight:600;color:var(--tx);user-select:none}
  .cb-switch input{display:none}
  .cb-switch .tg{width:34px;height:20px;border-radius:999px;background:#cfd8e4;position:relative;transition:background .2s;flex-shrink:0}
  .cb-switch .tg::after{content:'';position:absolute;top:2px;left:2px;width:16px;height:16px;border-radius:50%;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,.25);transition:transform .2s}
  .cb-switch input:checked+.tg{background:var(--b)}
  .cb-switch input:checked+.tg::after{transform:translateX(14px)}
  .cb-note{font-size:11.5px;color:var(--mu);margin:9px 2px 0;line-height:1.45}

  /* champs */
  .cb-search{position:relative}
  .cb-search svg{position:absolute;left:11px;top:50%;transform:translateY(-50%);color:var(--mu)}
  #cubeOverlay .cb-search input{width:100%;height:38px;border:1px solid var(--li)!important;border-radius:11px;padding:0 12px 0 34px!important;font:500 13px 'Plus Jakarta Sans',Inter,sans-serif;background:#fff;outline:none;transition:all .18s}
  #cubeOverlay .cb-search input:focus{border-color:var(--b)!important;box-shadow:0 0 0 4px rgba(26,78,138,.12)}
  .cb-fields{flex:1;overflow:auto;padding:2px 10px 18px}
  .cb-grp{font-size:10.5px;font-weight:700;letter-spacing:.09em;text-transform:uppercase;color:#a3aebd;padding:12px 8px 6px}
  .cb-f{position:relative;display:flex;align-items:center;gap:9px;height:36px;padding:0 8px;border-radius:10px;cursor:grab;user-select:none;transition:background .15s}
  .cb-f:hover{background:#fff;box-shadow:0 1px 0 var(--li),0 4px 12px -8px rgba(14,26,43,.3)}
  .cb-f .t{font-size:8.5px;font-weight:800;letter-spacing:.04em;width:34px;height:20px;border-radius:6px;display:flex;align-items:center;justify-content:center;background:var(--li2);color:#8592a5;flex-shrink:0}
  .cb-f .t.num{background:#e9f8f1;color:#0f8a5f}.cb-f .t.date{background:#fff1e8;color:#d4561a}
  .cb-f .nm{flex:1;min-width:0;font-size:13px;font-weight:500;color:var(--ink);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .cb-f.used .nm{font-weight:700;color:var(--b)}
  .cb-f .dot{width:6px;height:6px;border-radius:50%;background:var(--o);flex-shrink:0}
  .cb-f .qa{display:none;gap:3px}
  .cb-f:hover .qa{display:flex}
  .cb-f:hover .dot{display:none}
  .cb-qa{height:24px;min-width:24px;padding:0 6px;border-radius:7px;border:1px solid var(--li);background:#fff;color:var(--mu);font-size:10.5px;font-weight:700;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:3px;transition:all .15s}
  .cb-qa:hover{background:var(--b);border-color:var(--b);color:#fff}
  .cb-qa.v:hover{background:var(--o);border-color:var(--o)}

  /* zone principale */
  .cb-main{display:flex;flex-direction:column;gap:14px;min-width:0;min-height:0;padding:18px 22px 20px}
  .cb-zones{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px}
  .cb-zone{background:#fff;border:1px solid var(--li);border-radius:16px;padding:12px 12px 10px;min-height:92px;transition:all .18s;box-shadow:0 1px 2px rgba(14,26,43,.03)}
  .cb-zone .zh{display:flex;align-items:center;gap:8px;font-size:12px;font-weight:700;color:var(--ink);margin-bottom:9px}
  .cb-zone .zh .zi{width:24px;height:24px;border-radius:8px;display:flex;align-items:center;justify-content:center;background:var(--zc1);color:var(--zc)}
  .cb-zone .zh small{margin-left:auto;font-size:11px;color:#a3aebd;font-weight:600}
  .cb-zone.over{border-color:var(--zc);background:color-mix(in srgb,var(--zc) 5%,#fff);box-shadow:0 0 0 4px color-mix(in srgb,var(--zc) 14%,transparent)}
  .cb-pills{display:flex;flex-wrap:wrap;gap:6px}
  .cb-ph{width:100%;border:1.5px dashed #dbe2ec;border-radius:10px;padding:8px 10px;font-size:12px;color:#a3aebd;font-weight:500;text-align:center}
  .cb-pill{display:inline-flex;align-items:center;gap:6px;max-width:100%;height:30px;padding:0 5px 0 11px;border-radius:9px;background:var(--zc1);color:var(--zc);font-size:12.5px;font-weight:600;cursor:grab;border:1px solid transparent;transition:all .15s}
  .cb-pill:hover{border-color:color-mix(in srgb,var(--zc) 35%,transparent)}
  .cb-pill .lb{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .cb-pill .x{width:20px;height:20px;border-radius:6px;border:none;background:transparent;color:inherit;opacity:.55;cursor:pointer;display:flex;align-items:center;justify-content:center}
  .cb-pill .x:hover{opacity:1;background:rgba(0,0,0,.07)}
  .cb-pill .cnt{background:var(--zc);color:#fff;border-radius:999px;font-size:10.5px;font-weight:700;padding:1px 7px}
  #cubeOverlay .cb-pill select{border:none!important;background:rgba(255,255,255,.7)!important;background-image:none!important;font:700 11.5px 'Plus Jakarta Sans',Inter,sans-serif;color:inherit;padding:2px 4px!important;border-radius:6px;cursor:pointer;outline:none;box-shadow:none!important;-webkit-appearance:auto;appearance:auto}

  /* barre d'outils */
  .cb-bar{display:flex;flex-wrap:wrap;align-items:center;gap:10px}
  .cb-seg{display:inline-flex;background:#e9edf3;border-radius:11px;padding:3px;gap:2px}
  .cb-seg button{display:inline-flex;align-items:center;gap:6px;height:30px;padding:0 11px;border:none;border-radius:8px;background:transparent;color:var(--mu);font-size:12.5px;font-weight:600;cursor:pointer;transition:all .15s}
  .cb-seg button:hover{color:var(--ink)}
  .cb-seg button.on{background:#fff;color:var(--b);box-shadow:0 1px 3px rgba(14,26,43,.12)}
  .cb-ctl{display:inline-flex;align-items:center;gap:7px;font-size:12.5px;font-weight:600;color:var(--mu)}
  #cubeOverlay .cb-ctl select{height:32px;border:1px solid var(--li)!important;border-radius:9px;background-color:#fff;padding:0 28px 0 10px!important;font:600 12.5px 'Plus Jakarta Sans',Inter,sans-serif;color:var(--ink)}
  .cb-chk{display:inline-flex;align-items:center;gap:8px;font-size:12.5px;font-weight:600;color:var(--tx);cursor:pointer;user-select:none}
  .cb-chk input{display:none}
  .cb-chk .bx{width:17px;height:17px;border-radius:5px;border:1.5px solid #c5cfdc;display:flex;align-items:center;justify-content:center;color:#fff;transition:all .15s}
  .cb-chk input:checked+.bx{background:var(--b);border-color:var(--b)}

  /* indicateurs */
  .cb-kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:12px}
  .cb-kpi{background:#fff;border:1px solid var(--li);border-radius:16px;padding:12px 16px;display:flex;flex-direction:column;gap:2px;box-shadow:0 1px 2px rgba(14,26,43,.03)}
  .cb-kpi span{font-size:11px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:var(--mu)}
  .cb-kpi b{font-size:22px;font-weight:800;color:var(--ink);letter-spacing:-.02em;font-feature-settings:'tnum'}
  .cb-kpi.acc b{color:var(--b)}

  /* résultat */
  .cb-res{flex:1;min-height:0;position:relative;background:#fff;border:1px solid var(--li);border-radius:18px;overflow:auto;box-shadow:0 1px 2px rgba(14,26,43,.03),0 18px 40px -26px rgba(14,26,43,.28)}
  table.cb-t{border-collapse:separate;border-spacing:0;min-width:100%;font-size:13px;font-feature-settings:'tnum'}
  table.cb-t th,table.cb-t td{padding:10px 16px;white-space:nowrap;border-bottom:1px solid var(--li2)}
  table.cb-t thead th{position:sticky;top:0;z-index:2;background:#fafbfd;font-size:11.5px;font-weight:700;color:var(--mu);text-align:right;border-bottom:1px solid var(--li)}
  table.cb-t thead tr:nth-child(2) th{top:39px;font-weight:600;font-size:11px}
  table.cb-t thead th.rh{left:0;z-index:3;text-align:left;color:var(--ink);font-size:12px}
  table.cb-t thead th.grp{text-align:center;color:var(--ink);font-size:12px;border-left:1px solid var(--li2)}
  table.cb-t tbody th{position:sticky;left:0;z-index:1;background:#fff;text-align:left;font-weight:500;color:var(--ink);max-width:380px;overflow:hidden;text-overflow:ellipsis}
  table.cb-t tbody tr:hover th,table.cb-t tbody tr:hover td{background-color:#f7f9fc}
  table.cb-t td{text-align:right;color:var(--ink);cursor:pointer;position:relative}
  table.cb-t td.h{color:#fff}
  table.cb-t td:hover{box-shadow:inset 0 0 0 2px var(--o)}
  table.cb-t td.z{color:#c3ccd8}
  table.cb-t tr.sub th,table.cb-t tr.sub td{background:#f6f8fb;font-weight:700}
  table.cb-t tr.sub th{cursor:pointer}
  table.cb-t tr.sub th .tw{display:inline-flex;vertical-align:-3px;margin-right:6px;color:var(--mu);transition:transform .2s}
  table.cb-t tr.sub th .k{display:inline-block;margin-left:8px;font-size:11px;font-weight:600;color:#a3aebd}
  table.cb-t tr.tot th,table.cb-t tr.tot td{position:sticky;bottom:0;background:#eef3fa;font-weight:800;color:var(--b);border-top:1px solid #d5e1f0;border-bottom:none}
  table.cb-t .ct{background:#fafbfd;font-weight:700;border-left:1px solid var(--li)}
  table.cb-t tr.tot .ct{background:#e3ebf6}

  /* graphique */
  .cb-chart{padding:22px 26px}
  .cb-chart h4{margin:0 0 18px;font-size:13px;font-weight:700;color:var(--ink)}
  .cb-chart h4 span{color:var(--mu);font-weight:500}
  .cb-row{display:grid;grid-template-columns:minmax(120px,230px) 1fr 90px;align-items:center;gap:14px;padding:6px 0}
  .cb-row .lb{font-size:12.5px;font-weight:600;color:var(--tx);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .cb-row .tr{height:26px;border-radius:8px;background:var(--li2);display:flex;overflow:hidden}
  .cb-row .sg{height:100%;transition:width .5s cubic-bezier(.2,.7,.2,1)}
  .cb-row .sg:first-child{border-radius:8px 0 0 8px}
  .cb-row .vl{text-align:right;font-weight:800;color:var(--ink);font-feature-settings:'tnum';font-size:13px}
  .cb-legend{display:flex;flex-wrap:wrap;gap:14px;margin-top:18px;padding-top:14px;border-top:1px solid var(--li2);font-size:12px;font-weight:600;color:var(--tx)}
  .cb-legend i{display:inline-block;width:10px;height:10px;border-radius:3px;margin-right:6px}

  /* états */
  .cb-empty{min-height:100%;display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;padding:36px 28px;gap:6px}
  .cb-empty .ill{width:64px;height:64px;border-radius:20px;display:flex;align-items:center;justify-content:center;color:var(--b);background:linear-gradient(135deg,#eef3fa,#dfe9f6);margin-bottom:8px;box-shadow:inset 0 1px 0 #fff}
  .cb-empty h3{margin:0;font-size:20px;font-weight:800;color:var(--ink);letter-spacing:-.02em}
  .cb-empty p{margin:0;font-size:13.5px;color:var(--mu);max-width:560px}
  .cb-presets{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:12px;width:100%;max-width:960px;margin-top:22px}
  .cb-preset{text-align:left;display:flex;gap:12px;align-items:flex-start;padding:14px;border-radius:14px;border:1px solid var(--li);background:#fff;cursor:pointer;transition:all .18s}
  .cb-preset:hover{border-color:#c6d4e6;transform:translateY(-2px);box-shadow:0 12px 24px -16px rgba(26,78,138,.5)}
  .cb-preset .pi{width:34px;height:34px;border-radius:10px;flex-shrink:0;display:flex;align-items:center;justify-content:center;background:var(--b1);color:var(--b)}
  .cb-preset b{display:block;font-size:13px;font-weight:700;color:var(--ink);margin-bottom:2px}
  .cb-preset span{font-size:12px;color:var(--mu);line-height:1.35}
  .cb-err .ill{background:linear-gradient(135deg,#fff1e8,#ffe2cf);color:var(--o)}
  .cb-loading{position:absolute;inset:0;z-index:6;display:flex;align-items:center;justify-content:center;gap:12px;background:rgba(255,255,255,.72);backdrop-filter:blur(2px);font-weight:700;color:var(--b);font-size:13.5px}
  .cb-spin{width:22px;height:22px;border-radius:50%;border:3px solid #dbe6f4;border-top-color:var(--b);animation:cbSpin .8s linear infinite}
  @keyframes cbSpin{to{transform:rotate(360deg)}}
  .cb-skel{height:14px;border-radius:7px;background:linear-gradient(90deg,#eef2f7,#f7f9fc,#eef2f7);background-size:200% 100%;animation:cbSk 1.2s infinite}
  @keyframes cbSk{to{background-position:-200% 0}}

  /* popover, modales */
  .cb-pop{position:fixed;z-index:9100;min-width:240px;max-width:340px;padding:8px;border-radius:14px;background:#fff;border:1px solid var(--li,#e7ecf3);box-shadow:0 24px 60px -16px rgba(14,26,43,.35);font-family:'Plus Jakarta Sans',Inter,sans-serif;font-size:13px;color:#3b4859;animation:cbIn .16s}
  .cb-pop .ph{font-size:11px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:#7d8a9c;padding:6px 8px 8px;display:flex;align-items:center;gap:6px}
  .cb-pop .it{display:flex;align-items:center;gap:10px;padding:9px 10px;border-radius:9px;font-weight:600;cursor:pointer}
  .cb-pop .it:hover{background:#f3f7fc;color:#1A4E8A}
  .cb-pop .vals{max-height:280px;overflow:auto;margin:6px 0}
  .cb-pop label{display:flex;align-items:center;gap:9px;padding:6px 8px;border-radius:8px;cursor:pointer;font-weight:500}
  .cb-pop label:hover{background:#f6f8fb}
  .cb-pop label input{accent-color:#1A4E8A;width:15px;height:15px}
  .cb-pop .n{margin-left:auto;color:#a3aebd;font-size:11.5px;font-feature-settings:'tnum'}
  .cb-pop .mini{height:28px;padding:0 10px;border-radius:8px;border:1px solid #e7ecf3;background:#fff;font:600 12px 'Plus Jakarta Sans',Inter,sans-serif;color:#3b4859;cursor:pointer}
  .cb-pop .mini:hover{border-color:#1A4E8A;color:#1A4E8A}
  .cb-pop input.s{width:100%;height:34px;border:1px solid #e7ecf3!important;border-radius:9px;padding:0 10px!important;font:500 12.5px 'Plus Jakarta Sans',Inter,sans-serif;outline:none}
  .cb-modal{position:fixed;inset:0;z-index:9200;background:rgba(14,26,43,.42);backdrop-filter:blur(8px);display:flex;align-items:center;justify-content:center;padding:20px;font-family:'Plus Jakarta Sans',Inter,sans-serif;color:#3b4859}
  .cb-modal .box{width:100%;max-width:580px;max-height:90vh;overflow:auto;background:#fff;border-radius:22px;padding:26px;box-shadow:0 40px 90px -24px rgba(14,26,43,.5);animation:cbIn .22s}
  .cb-modal h3{margin:0;font-size:19px;font-weight:800;color:#0e1a2b;letter-spacing:-.02em}
  .cb-modal .sub{font-size:13px;color:#7d8a9c;margin:6px 0 0;line-height:1.5}
  .cb-modal .fl{display:block;font-size:11px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:#7d8a9c;margin:18px 0 7px}
  .cb-modal input[type=text],.cb-modal select{width:100%;height:42px;border:1px solid #e7ecf3!important;border-radius:11px;padding:0 12px;font:500 13.5px 'Plus Jakarta Sans',Inter,sans-serif;color:#0e1a2b}
  .cb-recap{margin-top:18px;border-radius:14px;background:#f6f8fb;padding:14px 16px;display:grid;grid-template-columns:auto 1fr;gap:6px 14px;font-size:12.5px}
  .cb-recap dt{font-weight:700;color:#7d8a9c}.cb-recap dd{margin:0;color:#0e1a2b;font-weight:500}
  .cb-modal .ft{display:flex;justify-content:flex-end;gap:10px;margin-top:22px}
  @media(max-width:1200px){.cb-zones{grid-template-columns:repeat(2,minmax(0,1fr))}}
  @media(max-width:900px){.cb-body{grid-template-columns:1fr}.cb-side{max-height:42vh}.cb-top .hide-sm{display:none}}
  `;
  document.head.appendChild(st);
}

// ── ouverture / structure ───────────────────────────────────────────────
const ZONES = [
  { k: 'filters', l: 'Filtres',  ic: 'filter', c: '#475569', c1: '#eef1f5', ph: 'Glissez un champ pour filtrer' },
  { k: 'rows',    l: 'Lignes',   ic: 'rows',   c: '#1A4E8A', c1: '#e8eff8', ph: 'Glissez un champ ici' },
  { k: 'cols',    l: 'Colonnes', ic: 'cols',   c: '#6d28d9', c1: '#f1ebfd', ph: 'Glissez un champ ici' },
  { k: 'vals',    l: 'Valeurs',  ic: 'sigma',  c: '#d4561a', c1: '#fff1e8', ph: '' }
];
function open() {
  injectCss();
  S.canActivite = (typeof vieScRole !== 'undefined' && vieScRole === 'Super-admin') || (typeof user !== 'undefined' && user.role === 'Administrateur');
  if ($('cubeOverlay')) { $('cubeOverlay').style.display = 'flex'; return; }
  const ov = document.createElement('div');
  ov.id = 'cubeOverlay';
  ov.innerHTML = `
    <header class="cb-top">
      <div class="cb-logo">${svg('cube', 20, 1.8)}</div>
      <div class="cb-title"><b>Cube</b><span id="cubeSub">Analyse croisée dynamique</span></div>
      <div class="cb-sp"></div>
      <select id="cubeViews" class="cb-sel hide-sm" title="Vues enregistrées" onchange="CUBE.loadView(this.value)"></select>
      <button class="cb-btn hide-sm" onclick="CUBE.saveView()">${svg('save', 15)}Enregistrer</button>
      <button class="cb-btn hide-sm" onclick="CUBE.reset()">${svg('reset', 15)}Réinitialiser</button>
      <button class="cb-btn hide-sm" onclick="CUBE.exportCsv()">${svg('download', 15)}CSV</button>
      <div class="cb-div hide-sm"></div>
      <button class="cb-btn ora" onclick="CUBE.openExport()">${svg('send', 15)}Envoyer vers Google Sheets</button>
      <button class="cb-btn ico" onclick="CUBE.close()" title="Fermer (Échap)">${svg('x', 17)}</button>
    </header>
    <div class="cb-body">
      <aside class="cb-side">
        <div class="cb-sec">
          <div class="cb-lbl">Sources de données</div>
          <div class="cb-srcs" id="cubeSrc"></div>
          <label class="cb-switch"><input type="checkbox" id="cubeJoin" onchange="CUBE.setJoin(this.checked)"><span class="tg"></span>Ajouter la fiche entreprise</label>
          <p class="cb-note">Sélectionnez plusieurs sources pour les croiser : un champ « Source » est ajouté.</p>
        </div>
        <div class="cb-sec" style="padding-top:4px;padding-bottom:8px">
          <div class="cb-lbl">Champs <span class="n" id="cubeNbF">0</span></div>
          <div class="cb-search">${svg('search', 15)}<input id="cubeFSearch" placeholder="Rechercher un champ" oninput="CUBE.searchFields(this.value)"></div>
        </div>
        <div class="cb-fields" id="cubeFields"></div>
      </aside>
      <section class="cb-main">
        <div class="cb-zones">${ZONES.map(z => `
          <div class="cb-zone" data-zone="${z.k}" style="--zc:${z.c};--zc1:${z.c1}">
            <div class="zh"><span class="zi">${svg(z.ic, 14)}</span>${z.l}<small id="cubeZN_${z.k}"></small></div>
            <div class="cb-pills" id="cubeZ_${z.k}"></div>
          </div>`).join('')}
        </div>
        <div class="cb-bar" id="cubeOpts"></div>
        <div class="cb-kpis" id="cubeKpis"></div>
        <div class="cb-res" id="cubeResult"></div>
      </section>
    </div>`;
  document.body.appendChild(ov);
  document.addEventListener('keydown', onKey);
  bindZones();
  renderSources();
  renderViews();
  render();
  loadData();
}
function close() { const o = $('cubeOverlay'); if (o) o.style.display = 'none'; closePop(); }
function onKey(e) {
  if (e.key !== 'Escape' || !$('cubeOverlay') || $('cubeOverlay').style.display === 'none') return;
  if (document.querySelector('.cb-modal')) { document.querySelector('.cb-modal').remove(); return; }
  if (document.querySelector('.cb-pop')) { closePop(); return; }
  close();
}

// ── sources & chargement ────────────────────────────────────────────────
function renderSources() {
  $('cubeSrc').innerHTML = CUBE_SRC.filter(s => !s.admin || S.canActivite).map(s =>
    `<div class="cb-src ${S.sources.includes(s.k) ? 'on' : ''}" style="--c:${s.c}" onclick="CUBE.toggleSource('${s.k}')">
       <span class="i">${svg(s.ic, 15)}</span>${s.l}<span class="ck">${svg('check', 10, 3)}</span></div>`).join('');
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
  S.loading = true; S.error = '';
  showLoading(true);
  try {
    const r = await api('getCubeData', { sources: S.sources, join: S.join });
    if (seq !== _loadSeq) return;
    if (!r || !r.success) {
      const msg = (r && r.message) || 'Réponse vide du serveur';
      S.error = /action inconnue/i.test(msg) ? 'serveur' : msg;
      S.fields = []; S.rows = []; S.loadedKey = '';
      render(); return;
    }
    S.fields = r.fields || []; S.rows = r.rows || []; S.loadedKey = key;
    if (typeof r.activiteAutorisee === 'boolean') S.canActivite = r.activiteAutorisee;
    if (thenCfg) applyCfg(thenCfg);
    const ok = n => !n || fieldIdx(n) !== -1;
    S.cfg.rows = S.cfg.rows.filter(ok); S.cfg.cols = S.cfg.cols.filter(ok);
    S.cfg.vals = S.cfg.vals.filter(v => ok(v.f)); if (!S.cfg.vals.length) S.cfg.vals = [{ f: '', agg: 'count' }];
    Object.keys(S.cfg.filters).forEach(n => { if (!ok(n)) delete S.cfg.filters[n]; });
    renderSources(); render();
  } catch (e) {
    console.error(e);
    if (seq === _loadSeq) { S.error = 'reseau'; render(); }
  } finally {
    if (seq === _loadSeq) { S.loading = false; showLoading(false); }
  }
}
function showLoading(on) {
  const res = $('cubeResult'); if (!res) return;
  const ex = $('cubeLoad');
  if (on && !ex) res.insertAdjacentHTML('beforeend', `<div class="cb-loading" id="cubeLoad"><div class="cb-spin"></div>Chargement des données…</div>`);
  if (!on && ex) ex.remove();
}
function applyCfg(c) {
  S.cfg = { rows: (c.rows || []).slice(), cols: (c.cols || []).slice(), vals: (c.vals || [{ f: '', agg: 'count' }]).map(v => ({ ...v })), filters: JSON.parse(JSON.stringify(c.filters || {})) };
  S.opt = { ...OPT_DEFAULT, ...(c.opt || {}) };
  S.collapsed = new Set();
}

// ── champs ──────────────────────────────────────────────────────────────
function usedFields() {
  const u = new Set([...S.cfg.rows, ...S.cfg.cols, ...Object.keys(S.cfg.filters)]);
  S.cfg.vals.forEach(v => v.f && u.add(v.f));
  return u;
}
function renderFields() {
  const q = S.fieldSearch.toLowerCase();
  const used = usedFields();
  $('cubeNbF').textContent = S.fields.length;
  if (!S.fields.length) {
    $('cubeFields').innerHTML = S.loading || !S.error
      ? Array.from({ length: 9 }, (_, i) => `<div style="padding:11px 8px"><div class="cb-skel" style="width:${55 + (i * 17) % 40}%"></div></div>`).join('')
      : '<p class="cb-note" style="padding:8px">Aucun champ disponible.</p>';
    return;
  }
  const list = S.fields.filter(f => !q || f.n.toLowerCase().includes(q));
  const groups = [
    { t: 'Informations', f: list.filter(f => !f.n.startsWith('Entreprise › ') && !f.n.includes(' · ')) },
    { t: 'Périodes', f: list.filter(f => f.n.includes(' · ') && !f.n.startsWith('Entreprise › ')) },
    { t: 'Fiche entreprise', f: list.filter(f => f.n.startsWith('Entreprise › ')) }
  ].filter(g => g.f.length);
  $('cubeFields').innerHTML = groups.map(g => `<div class="cb-grp">${g.t}</div>` + g.f.map(f => {
    const n = jsq(f.n);
    return `<div class="cb-f ${used.has(f.n) ? 'used' : ''}" draggable="true" data-field="${esc(f.n)}" ondblclick="CUBE.addTo('rows','${n}')" title="${esc(f.n)} — glisser vers une zone, double-clic = Lignes">
       <span class="t ${f.t}">${typeIcon(f.t)}</span><span class="nm">${esc(shortName(f.n))}</span>
       ${used.has(f.n) ? '<span class="dot"></span>' : ''}
       <span class="qa">
         <button class="cb-qa" title="Ajouter en lignes" onclick="event.stopPropagation();CUBE.addTo('rows','${n}')">${svg('rows', 12, 2.4)}</button>
         <button class="cb-qa" title="Ajouter en colonnes" onclick="event.stopPropagation();CUBE.addTo('cols','${n}')">${svg('cols', 12, 2.4)}</button>
         <button class="cb-qa v" title="${f.t === 'num' ? 'Somme' : 'Nombre distinct'}" onclick="event.stopPropagation();CUBE.addTo('vals','${n}')">${svg('sigma', 12, 2.4)}</button>
         <button class="cb-qa" title="Filtrer" onclick="event.stopPropagation();CUBE.addTo('filters','${n}')">${svg('filter', 11, 2.4)}</button>
       </span></div>`;
  }).join('')).join('') || '<p class="cb-note" style="padding:8px">Aucun champ ne correspond.</p>';
  document.querySelectorAll('#cubeFields .cb-f').forEach(el => {
    el.addEventListener('dragstart', e => { e.dataTransfer.setData('text/plain', JSON.stringify({ field: el.dataset.field })); e.dataTransfer.effectAllowed = 'copyMove'; });
  });
}
function searchFields(v) { S.fieldSearch = v; renderFields(); }

// ── zones ───────────────────────────────────────────────────────────────
function bindZones() {
  document.querySelectorAll('#cubeOverlay .cb-zone').forEach(z => {
    z.addEventListener('dragover', e => { e.preventDefault(); z.classList.add('over'); });
    z.addEventListener('dragleave', e => { if (!z.contains(e.relatedTarget)) z.classList.remove('over'); });
    z.addEventListener('drop', e => {
      e.preventDefault(); z.classList.remove('over');
      let d; try { d = JSON.parse(e.dataTransfer.getData('text/plain')); } catch (x) { return; }
      if (!d || (!d.field && d.from !== 'vals')) return;
      if (d.from === z.dataset.zone) return;
      if (d.from) removeFrom(d.from, d.field, d.idx, true);
      if (d.field) addTo(z.dataset.zone, d.field, true); else render();
    });
  });
}
function addTo(zone, n, silent) {
  closePop();
  if (zone === 'rows' || zone === 'cols') {
    S.cfg.rows = S.cfg.rows.filter(x => x !== n); S.cfg.cols = S.cfg.cols.filter(x => x !== n);
    S.cfg[zone].push(n);
  } else if (zone === 'vals') {
    S.cfg.vals.push({ f: n, agg: fieldType(n) === 'num' ? 'sum' : 'countd' });
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
  const pill = (zone, n, inner, idx) =>
    `<span class="cb-pill" draggable="true" data-zone="${zone}" data-f="${esc(n)}" data-idx="${idx == null ? '' : idx}">${inner}<button class="x" onclick="event.stopPropagation();CUBE.removeFrom('${zone}','${jsq(n)}',${idx == null ? 'null' : idx})" title="Retirer">${svg('x', 12, 2.6)}</button></span>`;
  ZONES.forEach(z => {
    let html = '';
    if (z.k === 'rows' || z.k === 'cols') html = S.cfg[z.k].map(n => pill(z.k, n, `<span class="lb">${esc(shortName(n))}</span>`)).join('');
    if (z.k === 'vals') html = S.cfg.vals.map((v, i) => {
      const allowed = !v.f ? ['count'] : (fieldType(v.f) === 'num' ? Object.keys(AGG) : ['count', 'countd']);
      const opts = allowed.map(a => `<option value="${a}" ${a === v.agg ? 'selected' : ''}>${AGG[a]}</option>`).join('');
      return pill('vals', v.f || '', `<select onchange="CUBE.setAgg(${i},this.value)" onclick="event.stopPropagation()" onmousedown="event.stopPropagation()">${opts}</select><span class="lb">${v.f ? esc(shortName(v.f)) : 'de lignes'}</span>`, i);
    }).join('');
    if (z.k === 'filters') html = Object.keys(S.cfg.filters).map(n => {
      const sel = S.cfg.filters[n];
      return pill('filters', n, `<span class="lb">${esc(shortName(n))}</span>${sel.length ? `<span class="cnt">${sel.length}</span>` : ''}`);
    }).join('');
    $('cubeZ_' + z.k).innerHTML = html || (z.ph ? `<div class="cb-ph">${z.ph}</div>` : '');
    const nb = z.k === 'filters' ? Object.keys(S.cfg.filters).length : (z.k === 'vals' ? S.cfg.vals.length : S.cfg[z.k].length);
    $('cubeZN_' + z.k).textContent = nb ? nb : '';
  });
  document.querySelectorAll('#cubeOverlay .cb-pill').forEach(p => {
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
    <div class="ph">${svg('filter', 12)} ${esc(n)}</div>
    <div style="display:flex;gap:6px;padding:0 4px 8px">
      <input class="s" placeholder="Rechercher une valeur" oninput="CUBE.filterSearch(this.value)">
      <button class="mini" onclick="CUBE.filterAll('${jsq(n)}',true)">Tout</button>
      <button class="mini" onclick="CUBE.filterAll('${jsq(n)}',false)">Aucun</button>
    </div>
    <div class="vals" id="cubeFVals">${vals.map(([v, c]) => `<label data-v="${esc(v.toLowerCase())}"><input type="checkbox" value="${esc(v)}" ${!sel.size || sel.has(v) ? 'checked' : ''} onchange="CUBE.filterChange('${jsq(n)}')"><span style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(v)}</span><span class="n">${c.toLocaleString('fr')}</span></label>`).join('')}</div>
    <div style="font-size:11.5px;color:#a3aebd;padding:2px 8px 4px">Tout coché = pas de filtre</div>`);
}
function filterSearch(q) { q = q.toLowerCase(); document.querySelectorAll('#cubeFVals label').forEach(l => l.style.display = l.dataset.v.includes(q) ? '' : 'none'); }
function filterChange(n) {
  const all = [...document.querySelectorAll('#cubeFVals input')];
  const on = all.filter(i => i.checked).map(i => i.value);
  S.cfg.filters[n] = on.length === all.length ? [] : on;
  renderZones(); renderKpis(); renderResult();
}
function filterAll(n, v) { document.querySelectorAll('#cubeFVals input').forEach(i => { if (i.closest('label').style.display !== 'none') i.checked = v; }); filterChange(n); }

// ── popovers ─────────────────────────────────────────────────────────────
function openPop(ev, html) {
  closePop();
  const p = document.createElement('div');
  p.className = 'cb-pop'; p.innerHTML = html;
  document.body.appendChild(p);
  const r = (ev.currentTarget || ev.target).getBoundingClientRect();
  let x = r.left, y = r.bottom + 8;
  const w = p.offsetWidth, h = p.offsetHeight;
  if (x + w > innerWidth - 10) x = innerWidth - w - 10;
  if (y + h > innerHeight - 10) y = Math.max(10, r.top - h - 8);
  p.style.left = x + 'px'; p.style.top = y + 'px';
  setTimeout(() => document.addEventListener('mousedown', _outside), 0);
}
function _outside(e) { const p = document.querySelector('.cb-pop'); if (p && !p.contains(e.target)) closePop(); }
function closePop() { document.querySelectorAll('.cb-pop').forEach(p => p.remove()); document.removeEventListener('mousedown', _outside); }

// ── barre d'outils ──────────────────────────────────────────────────────
function renderOpts() {
  const seg = (key, items) => `<span class="cb-seg">${items.map(([v, l, ic]) => `<button class="${S.opt[key] === v ? 'on' : ''}" onclick="CUBE.setOpt('${key}','${v}')">${ic ? svg(ic, 14) : ''}${l}</button>`).join('')}</span>`;
  const chk = (key, l) => `<label class="cb-chk"><input type="checkbox" ${S.opt[key] ? 'checked' : ''} onchange="CUBE.setOpt('${key}',this.checked)"><span class="bx">${svg('check', 11, 3)}</span>${l}</label>`;
  $('cubeOpts').innerHTML = `
    ${seg('view', [['table', 'Tableau', 'table'], ['chart', 'Graphique', 'chart']])}
    ${seg('show', [['val', 'Valeurs'], ['pctTotal', '% du total'], ['pctRow', '% ligne'], ['pctCol', '% colonne']])}
    <span class="cb-ctl">Tri <select onchange="CUBE.setOpt('sort',this.value)">${[['label', 'Alphabétique'], ['desc', 'Plus grand d\'abord'], ['asc', 'Plus petit d\'abord']].map(([v, l]) => `<option value="${v}" ${S.opt.sort === v ? 'selected' : ''}>${l}</option>`).join('')}</select></span>
    <span class="cb-ctl">Afficher <select onchange="CUBE.setOpt('top',+this.value)">${[[0, 'Tout'], [5, 'Top 5'], [10, 'Top 10'], [20, 'Top 20'], [50, 'Top 50']].map(([v, l]) => `<option value="${v}" ${S.opt.top === v ? 'selected' : ''}>${l}</option>`).join('')}</select></span>
    ${chk('heat', 'Carte de chaleur')}
    ${chk('totals', 'Totaux')}
    <span style="flex:1"></span>
    <button class="cb-btn" onclick="CUBE.swapAxes()" title="Inverser lignes et colonnes">${svg('swap', 14)}Inverser</button>
    ${S.cfg.rows.length > 1 ? `<button class="cb-btn" onclick="CUBE.collapseAll(true)">Tout replier</button><button class="cb-btn" onclick="CUBE.collapseAll(false)">Tout déplier</button>` : ''}`;
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
  const cells = new Map(), rowKeys = new Map(), colKeys = new Map(), children = new Map();
  const SEP2 = '\u0002';
  const touch = k => { let a = cells.get(k); if (!a) { a = vals.map(newAcc); cells.set(k, a); } return a; };
  for (const r of rows) {
    const rp = ri.map(i => cellVal(r, i));
    const cp = ci.map(i => cellVal(r, i));
    const ck = cp.join(SEP);
    if (ci.length && !colKeys.has(ck)) colKeys.set(ck, cp);
    const colTargets = ci.length ? [ck, '*'] : ['*'];
    let parent = '';
    for (let lv = 0; lv <= rp.length; lv++) {
      const pk = lv === 0 ? '' : rp.slice(0, lv).join(SEP);
      if (lv > 0) {
        let m = rowKeys.get(lv); if (!m) { m = new Map(); rowKeys.set(lv, m); }
        if (!m.has(pk)) { m.set(pk, rp.slice(0, lv)); const ch = children.get(parent) || []; ch.push(pk); children.set(parent, ch); }
      }
      for (const ct of colTargets) {
        const accs = touch(lv + ':' + pk + SEP2 + ct);
        vals.forEach((v, vi) => addAcc(accs[vi], v.i === -1 ? 1 : r[v.i], v.agg));
      }
      parent = pk;
    }
  }
  const get = (lv, pk, ck, vi) => { const a = cells.get(lv + ':' + pk + SEP2 + ck); return a ? accVal(a[vi], vals[vi].agg) : null; };
  const cols = [...colKeys.entries()].sort((a, b) => {
    for (let k = 0; k < a[1].length; k++) { const c = collator.compare(a[1][k], b[1][k]); if (c) return c; }
    return 0;
  });
  return { rows, vals, get, cols, rowKeys, children, nLevels: ri.length };
}
function rowTree(P) {
  const sortVal = (lv, pk) => P.get(lv, pk, '*', 0) || 0;
  const build = (lv, parent) => {
    const keys = P.children.get(parent) || [];
    const m = P.rowKeys.get(lv);
    let list = keys.map(k => [k, m.get(k)]);
    if (S.opt.sort === 'label') list.sort((a, b) => collator.compare(a[1][lv - 1], b[1][lv - 1]));
    else list.sort((a, b) => S.opt.sort === 'desc' ? sortVal(lv, b[0]) - sortVal(lv, a[0]) : sortVal(lv, a[0]) - sortVal(lv, b[0]));
    if (S.opt.top && lv === 1) list = list.slice(0, S.opt.top);
    return list.map(([k, parts]) => ({ lv, key: k, label: parts[lv - 1], kids: lv < P.nLevels ? build(lv + 1, k) : [] }));
  };
  return P.nLevels ? build(1, '') : [];
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
function render() { renderFields(); renderZones(); renderOpts(); renderKpis(); renderResult(); }

function renderKpis() {
  const el = $('cubeKpis'); if (!el) return;
  const sub = $('cubeSub');
  if (sub) sub.textContent = S.rows.length ? `${S.rows.length.toLocaleString('fr')} lignes · ${S.sources.map(s => (CUBE_SRC.find(x => x.k === s) || {}).l).join(' + ')}` : 'Analyse croisée dynamique';
  if (!S.rows.length || (!S.cfg.rows.length && !S.cfg.cols.length)) { el.style.display = 'none'; return; }
  const P = computePivot();
  S.lastPivot = P;
  el.style.display = '';
  const groups = P.nLevels ? (P.rowKeys.get(1) || new Map()).size : 1;
  el.innerHTML = `
    <div class="cb-kpi"><span>Lignes analysées</span><b>${P.rows.length.toLocaleString('fr')}</b></div>
    <div class="cb-kpi"><span>${esc(S.cfg.rows[0] ? shortName(S.cfg.rows[0]) : 'Groupes')}</span><b>${groups.toLocaleString('fr')}</b></div>
    ${P.cols.length ? `<div class="cb-kpi"><span>${esc(shortName(S.cfg.cols.join(' › ')))}</span><b>${P.cols.length.toLocaleString('fr')}</b></div>` : ''}
    ${P.vals.slice(0, 2).map((v, vi) => `<div class="cb-kpi acc"><span>${esc(valLabel(v))}</span><b>${fmtNum(P.get(0, '', '*', vi), v.agg === 'avg' ? 2 : 0)}</b></div>`).join('')}`;
}

function renderResult() {
  const el = $('cubeResult');
  if (!el) return;
  const loadingHtml = S.loading ? `<div class="cb-loading" id="cubeLoad"><div class="cb-spin"></div>Chargement des données…</div>` : '';
  if (S.error) {
    const isSrv = S.error === 'serveur', isNet = S.error === 'reseau';
    el.innerHTML = `<div class="cb-empty cb-err"><div class="ill">${svg('alert', 28, 1.8)}</div>
      <h3>${isSrv ? 'Le serveur n\'est pas encore à jour' : 'Impossible de charger les données'}</h3>
      <p>${isSrv ? 'Le CUBE a besoin de la dernière version du script Google Apps Script. Une fois la nouvelle version mise en production, rouvrez le CUBE.' : (isNet ? 'Vérifiez votre connexion puis réessayez.' : esc(S.error))}</p>
      <button class="cb-btn pri" style="margin-top:16px" onclick="CUBE.retry()">${svg('reset', 14)}Réessayer</button></div>`;
    return;
  }
  if (!S.fields.length) { el.innerHTML = `<div class="cb-empty"><div class="ill">${svg('cube', 28, 1.6)}</div><h3>Préparation du cube…</h3><p>Lecture des données du CRM.</p></div>` + loadingHtml; return; }
  if (!S.cfg.rows.length && !S.cfg.cols.length) {
    const presets = PRESETS.filter(p => !p.admin || S.canActivite);
    el.innerHTML = `<div class="cb-empty">
      <div class="ill">${svg('cube', 30, 1.6)}</div>
      <h3>Que voulez-vous analyser ?</h3>
      <p>Survolez un champ à gauche et cliquez sur ${svg('rows', 12, 2.4)} <b>Lignes</b>, ${svg('cols', 12, 2.4)} <b>Colonnes</b> ou ${svg('sigma', 12, 2.4)} <b>Valeur</b> — ou glissez-le dans une zone. Cliquez sur une cellule pour voir le détail.</p>
      <div class="cb-presets">${presets.map(p => `<button class="cb-preset" onclick="CUBE.preset(${PRESETS.indexOf(p)})"><span class="pi">${svg(p.ic, 17)}</span><span><b>${esc(p.nom)}</b><span>${esc(p.d)}</span></span></button>`).join('')}</div>
    </div>` + loadingHtml;
    return;
  }
  const P = S.lastPivot && S.lastPivot.rows === filteredRows() ? S.lastPivot : computePivot();
  S.lastPivot = P;
  el.innerHTML = (S.opt.view === 'chart' ? renderChart(P) : renderTable(P)) + loadingHtml;
}

function renderTable(P) {
  const tree = rowTree(P);
  const nv = P.vals.length;
  const colList = P.cols;
  const hasCols = colList.length > 0;
  const showColTot = S.opt.totals && hasCols;
  const leafLv = P.nLevels;
  let maxV = [];
  if (S.opt.heat) {
    const keys = leafLv ? [...(P.rowKeys.get(leafLv) || new Map()).keys()] : [''];
    const targets = hasCols ? colList.map(c => c[0]) : ['*'];
    maxV = P.vals.map((_, vi) => { let m = 0; keys.forEach(k => targets.forEach(ck => { const v = P.get(leafLv, k, ck, vi); if (v > m) m = v; })); return m; });
  }
  const cellAttr = (v, vi, isLeaf) => {
    if (v == null || v === 0) return ' class="z"';
    if (!S.opt.heat || !isLeaf || !maxV[vi]) return '';
    const a = Math.max(0, Math.min(1, v / maxV[vi]));
    const al = (0.05 + a * 0.62).toFixed(3);
    return ` style="background:rgba(26,78,138,${al})"${a > 0.55 ? ' class="h"' : ''}`;
  };
  const rowHeadTitle = S.cfg.rows.map(shortName).join(' › ');
  let thead = '';
  if (hasCols) {
    thead += `<tr><th class="rh" rowspan="${nv > 1 ? 2 : 1}">${esc(rowHeadTitle || '—')}</th>`;
    colList.forEach(([, parts]) => { thead += `<th class="grp" colspan="${nv}">${esc(parts.join(' › '))}</th>`; });
    if (showColTot) thead += `<th class="grp ct" colspan="${nv}">Total</th>`;
    thead += '</tr>';
    if (nv > 1) { thead += '<tr>'; colList.concat(showColTot ? [[null]] : []).forEach(() => P.vals.forEach(v => { thead += `<th>${esc(valLabel(v))}</th>`; })); thead += '</tr>'; }
  } else {
    thead = `<tr><th class="rh">${esc(rowHeadTitle || 'Total')}</th>${P.vals.map(v => `<th>${esc(valLabel(v))}</th>`).join('')}</tr>`;
  }
  const cellsFor = (lv, pk, isLeaf) => {
    let h = '';
    const targets = hasCols ? colList.map(c => c[0]) : ['*'];
    targets.forEach(ck => P.vals.forEach((v, vi) => {
      const raw = P.get(lv, pk, ck, vi);
      h += `<td${cellAttr(raw, vi, isLeaf)} data-lv="${lv}" data-pk="${esc(pk)}" data-ck="${esc(ck)}">${display(P, raw, lv, pk, ck, vi) || '·'}</td>`;
    }));
    if (showColTot) P.vals.forEach((v, vi) => { const raw = P.get(lv, pk, '*', vi); h += `<td class="ct" data-lv="${lv}" data-pk="${esc(pk)}" data-ck="*">${display(P, raw, lv, pk, '*', vi)}</td>`; });
    return h;
  };
  let body = '';
  const walk = nodes => nodes.forEach(n => {
    const isLeaf = !n.kids.length;
    const col = S.collapsed.has(n.key);
    const pad = 16 + (n.lv - 1) * 22;
    if (isLeaf) body += `<tr><th style="padding-left:${pad + (P.nLevels > 1 ? 20 : 0)}px" title="${esc(n.label)}">${esc(n.label)}</th>${cellsFor(n.lv, n.key, true)}</tr>`;
    else {
      body += `<tr class="sub"><th style="padding-left:${pad}px" onclick="CUBE.toggleRow('${jsq(n.key)}')" title="${esc(n.label)}"><span class="tw">${svg(col ? 'chevR' : 'chevD', 14, 2.4)}</span>${esc(n.label)}<span class="k">${n.kids.length}</span></th>${cellsFor(n.lv, n.key, false)}</tr>`;
      if (!col) walk(n.kids);
    }
  });
  walk(tree);
  if (S.opt.totals || !P.nLevels) body += `<tr class="tot"><th>Total général</th>${cellsFor(0, '', !P.nLevels)}</tr>`;
  return `<table class="cb-t" id="cubeTable"><thead>${thead}</thead><tbody>${body}</tbody></table>`;
}

const PALETTE = ['#1A4E8A', '#F26522', '#10b981', '#8b5cf6', '#06b6d4', '#ec4899', '#f59e0b', '#6366f1', '#14b8a6', '#ef4444', '#84cc16', '#64748b'];
function renderChart(P) {
  const tree = P.nLevels ? rowTree(P) : [{ lv: 0, key: '', label: 'Total', kids: [] }];
  const vi = 0;
  const hasCols = P.cols.length > 0;
  const colsShown = hasCols ? P.cols.slice(0, 12) : [];
  const items = tree.map(n => {
    const total = P.get(n.lv, n.key, '*', vi) || 0;
    const segs = hasCols ? colsShown.map(([ck], i) => ({ v: P.get(n.lv, n.key, ck, vi) || 0, c: PALETTE[i % PALETTE.length], l: P.cols[i][1].join(' › ') })) : [{ v: total, c: 'linear-gradient(90deg,#2a62a6,#1A4E8A)', l: '' }];
    return { label: n.label, total, segs };
  });
  const max = Math.max(1, ...items.map(i => hasCols ? i.segs.reduce((s, x) => s + x.v, 0) : i.total));
  const v0 = P.vals[0];
  return `<div class="cb-chart"><h4>${esc(valLabel(v0).replace(/^Nombre$/, 'Nombre de lignes'))} <span>par ${esc(shortName(S.cfg.rows[0] || '—'))}${hasCols ? ' et ' + esc(S.cfg.cols.map(shortName).join(' › ')) : ''}</span></h4>
    ${items.map(it => `<div class="cb-row"><div class="lb" title="${esc(it.label)}">${esc(it.label)}</div>
      <div class="tr">${it.segs.map(s => `<div class="sg" title="${esc(s.l ? s.l + ' : ' : '')}${fmtNum(s.v)}" style="width:${(s.v / max * 100).toFixed(2)}%;background:${s.c}"></div>`).join('')}</div>
      <div class="vl">${fmtNum(it.total)}</div></div>`).join('')}
    ${hasCols ? `<div class="cb-legend">${colsShown.map(([, p], i) => `<span><i style="background:${PALETTE[i % PALETTE.length]}"></i>${esc(p.join(' › '))}</span>`).join('')}${P.cols.length > 12 ? '<span>… 12 premières colonnes</span>' : ''}</div>` : ''}</div>`;
}

function toggleRow(k) { if (S.collapsed.has(k)) S.collapsed.delete(k); else S.collapsed.add(k); renderResult(); }
function collapseAll(v) {
  S.collapsed = new Set();
  if (v && S.lastPivot) for (let lv = 1; lv < S.lastPivot.nLevels; lv++) (S.lastPivot.rowKeys.get(lv) || new Map()).forEach((_, k) => S.collapsed.add(k));
  renderResult();
}
function retry() { S.error = ''; S.loadedKey = ''; render(); loadData(); }

// ── détail d'une cellule ────────────────────────────────────────────────
function onCellClick(e) {
  const td = e.target.closest('#cubeTable td'); if (!td) return;
  const lv = Number(td.dataset.lv), pk = td.dataset.pk, ck = td.dataset.ck;
  const rp = lv ? pk.split(SEP) : [], cp = ck && ck !== '*' ? ck.split(SEP) : [];
  const ri = S.cfg.rows.slice(0, lv).map(fieldIdx), ci = S.cfg.cols.map(fieldIdx);
  const list = filteredRows().filter(r => ri.every((i, k) => cellVal(r, i) === rp[k]) && (!cp.length || ci.every((i, k) => cellVal(r, i) === cp[k])));
  const crit = S.cfg.rows.slice(0, lv).map((n, k) => `<b>${esc(shortName(n))}</b> ${esc(rp[k])}`).concat(cp.length ? S.cfg.cols.map((n, k) => `<b>${esc(shortName(n))}</b> ${esc(cp[k])}`) : []);
  const showF = S.fields.filter(f => !f.n.includes(' · ')).slice(0, 14);
  const idx = showF.map(f => fieldIdx(f.n));
  const m = document.createElement('div');
  m.className = 'cb-modal';
  m.onclick = ev => { if (ev.target === m) m.remove(); };
  m.innerHTML = `<div class="box" style="max-width:1120px">
    <div style="display:flex;align-items:flex-start;gap:12px"><div style="flex:1"><h3>${list.length.toLocaleString('fr')} ligne${list.length > 1 ? 's' : ''}</h3>
    <p class="sub">${crit.join(' &nbsp;·&nbsp; ') || 'Toutes les lignes filtrées'}</p></div>
    <button class="cb-btn" onclick="CUBE.exportDetail()">${svg('download', 14)}CSV</button><button class="cb-btn ico" onclick="this.closest('.cb-modal').remove()">${svg('x', 16)}</button></div>
    <div style="overflow:auto;max-height:64vh;margin-top:18px;border:1px solid #e7ecf3;border-radius:14px">
    <table class="cb-t"><thead><tr>${showF.map(f => `<th style="text-align:left">${esc(shortName(f.n))}</th>`).join('')}</tr></thead>
    <tbody>${list.slice(0, 500).map(r => `<tr>${idx.map(i => `<td style="text-align:left;cursor:default">${esc(r[i])}</td>`).join('')}</tr>`).join('')}</tbody></table></div>
    ${list.length > 500 ? '<p class="sub">500 premières lignes affichées — le CSV contient tout.</p>' : ''}
  </div>`;
  document.body.appendChild(m);
  S._detail = { list, fields: showF.map(f => f.n), idx };
}

// ── vues, modèles, export ───────────────────────────────────────────────
function getViews() { try { return JSON.parse(localStorage.getItem(VIEWS_KEY) || '[]'); } catch (e) { return []; } }
function renderViews() {
  const v = getViews();
  $('cubeViews').innerHTML = `<option value="">Mes vues (${v.length})</option>` + v.map((x, i) => `<option value="${i}">${esc(x.nom)}</option>`).join('') + (v.length ? '<option value="__del">Supprimer une vue…</option>' : '');
}
function currentConfig(nom) { return { nom, sources: S.sources.slice(), join: S.join, rows: S.cfg.rows.slice(), cols: S.cfg.cols.slice(), vals: S.cfg.vals.map(v => ({ f: v.f, agg: v.agg })), filters: JSON.parse(JSON.stringify(S.cfg.filters)), opt: { ...S.opt }, show: S.opt.show }; }
function saveView() {
  const nom = prompt('Nom de la vue :', S.cfg.rows.concat(S.cfg.cols).map(shortName).join(' × ') || 'Ma vue');
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
function reset() { S.cfg = { rows: [], cols: [], vals: [{ f: '', agg: 'count' }], filters: {} }; S.opt = { ...OPT_DEFAULT }; S.collapsed = new Set(); render(); }

function csvLine(arr) { return arr.map(v => { const s = String(v == null ? '' : v); return /[";\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; }).join(';'); }
function download(name, text) {
  const blob = new Blob(['﻿' + text], { type: 'text/csv;charset=utf-8' });
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 2000);
}
function exportCsv() {
  const t = $('cubeTable');
  if (!t) { toast('Affichez d\'abord un tableau (vue Tableau)', false); return; }
  const lines = [...t.querySelectorAll('tr')].map(tr => csvLine([...tr.children].flatMap(c => { const span = Number(c.getAttribute('colspan') || 1); const txt = c.textContent.replace(/·/g, '').trim(); return [txt].concat(Array(span - 1).fill('')); })));
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
  m.className = 'cb-modal';
  m.onclick = ev => { if (ev.target === m) m.remove(); };
  const flt = Object.entries(S.cfg.filters).filter(([, v]) => v.length).map(([n, v]) => shortName(n) + ' (' + v.length + ')').join(', ');
  m.innerHTML = `<div class="box">
    <div style="display:flex;gap:14px;align-items:flex-start">
      <div class="cb-logo" style="background:linear-gradient(135deg,#f47637,#F26522)">${svg('send', 18, 1.9)}</div>
      <div><h3>Envoyer vers Google Sheets</h3>
      <p class="sub">Un vrai tableau croisé dynamique Google Sheets est créé avec cette configuration. Il se met à jour tout seul à chaque synchronisation (toutes les heures).</p></div>
    </div>
    <label class="fl">Tableau de suivi</label>
    <select id="cubeExpTab">${mine.map(t => `<option value="${esc(t.id)}">${esc(t.nom)}</option>`).join('')}<option value="__new">+ Nouveau tableau de suivi…</option></select>
    <div id="cubeExpNewWrap" style="display:${mine.length ? 'none' : 'block'}"><label class="fl">Nom du nouveau tableau</label><input type="text" id="cubeExpNew" placeholder="Ex : Analyses 2026-2027"></div>
    <label class="fl">Nom de l'onglet</label>
    <input type="text" id="cubeExpNom" value="${esc(S.cfg.rows.concat(S.cfg.cols).map(shortName).join(' × ').slice(0, 60))}">
    <dl class="cb-recap">
      <dt>Sources</dt><dd>${esc(S.sources.map(s => (CUBE_SRC.find(x => x.k === s) || {}).l).join(' + '))}${S.join ? ' + fiche entreprise' : ''}</dd>
      <dt>Lignes</dt><dd>${esc(S.cfg.rows.map(shortName).join(', ') || '—')}</dd>
      <dt>Colonnes</dt><dd>${esc(S.cfg.cols.map(shortName).join(', ') || '—')}</dd>
      <dt>Valeurs</dt><dd>${esc(S.cfg.vals.map(valLabel).join(', '))}</dd>
      <dt>Filtres</dt><dd>${esc(flt || 'aucun')}</dd>
    </dl>
    <div class="ft"><button class="cb-btn" onclick="this.closest('.cb-modal').remove()">Annuler</button>
      <button class="cb-btn ora" id="cubeExpBtn" onclick="CUBE.doExport()">${svg('send', 14)}Créer le tableau croisé</button></div></div>`;
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
    document.querySelector('.cb-modal')?.remove();
    toast(`Tableau croisé « ${r.onglet} » créé (${r.lignes} lignes) ✓`);
    if (r.url) window.open(r.url, '_blank');
    if (typeof loadTableauxSuivi === 'function') loadTableauxSuivi();
  } finally { if (btn.isConnected) { btn.disabled = false; btn.innerHTML = svg('send', 14) + 'Créer le tableau croisé'; } }
}

document.addEventListener('click', e => { if (e.target.closest && e.target.closest('#cubeTable td')) onCellClick(e); });

window.CUBE = { open, close, toggleSource, setJoin, searchFields, addTo, removeFrom, setAgg, swapAxes, filterSearch, filterChange, filterAll,
  setOpt, toggleRow, collapseAll, saveView, loadView, preset, reset, retry, exportCsv, exportDetail, openExport, doExport, useConfig };
})();
