/* =============================================================================
   Vie scolaire — modules « Suivi des mémoires » et « Frises chronologiques »
   Backend : vs_extras.gs. Dépend de vie_scolaire.html : api(), toast(), user, vieScRole.
   ========================================================================== */
(function () {
'use strict';

// ── socle commun ─────────────────────────────────────────────────────────
const $ = id => document.getElementById(id);
const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const jsq = s => String(s == null ? '' : s).replace(/\\/g, '\\\\').replace(/'/g, "\\'");
const norm = s => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
const todayISO = () => { const d = new Date(); return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); };
const fmtD = iso => iso ? iso.split('-').reverse().join('/') : '';
const MOIS = ['janv.', 'févr.', 'mars', 'avr.', 'mai', 'juin', 'juil.', 'août', 'sept.', 'oct.', 'nov.', 'déc.'];
const fmtLong = iso => { if (!iso) return ''; const [y, m, d] = iso.split('-'); return `${+d} ${MOIS[+m - 1]} ${y}`; };
const daysTo = iso => Math.round((new Date(iso + 'T12:00:00') - new Date(todayISO() + 'T12:00:00')) / 86400000);
const IC = {
  plus: '<path d="M12 5v14M5 12h14"/>', down: '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="m7 10 5 5 5-5"/><path d="M12 15V3"/>',
  search: '<circle cx="11" cy="11" r="7"/><path d="m21 21-4.35-4.35"/>', x: '<path d="M18 6 6 18M6 6l12 12"/>', check: '<path d="M20 6 9 17l-5-5"/>',
  edit: '<path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4Z"/>', trash: '<path d="M3 6h18"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6"/><path d="M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/>',
  grid: '<rect x="3" y="3" width="7" height="7" rx="1.5"/><rect x="14" y="3" width="7" height="7" rx="1.5"/><rect x="3" y="14" width="7" height="7" rx="1.5"/><rect x="14" y="14" width="7" height="7" rx="1.5"/>',
  list: '<path d="M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01"/>', book: '<path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20V3H6.5A2.5 2.5 0 0 0 4 5.5z"/><path d="M8 7h8M8 11h6"/>',
  user: '<circle cx="12" cy="8" r="4"/><path d="M4 21a8 8 0 0 1 16 0"/>', pin: '<path d="M12 22s7-6.2 7-12a7 7 0 1 0-14 0c0 5.8 7 12 7 12z"/><circle cx="12" cy="10" r="2.5"/>',
  bell: '<path d="M6 8a6 6 0 0 1 12 0c0 7 3 9 3 9H3s3-2 3-9"/><path d="M10.3 21a1.94 1.94 0 0 0 3.4 0"/>', clock: '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
  flag: '<path d="M4 22V4"/><path d="M4 4h13l-2.5 4.5L17 13H4"/>',
  tag: '<path d="M20.6 13.4 13.4 20.6a2 2 0 0 1-2.8 0L3 13V3h10l7.6 7.6a2 2 0 0 1 0 2.8z"/><circle cx="7.5" cy="7.5" r="1.5"/>',
  file: '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/>', upload: '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="m17 8-5-5-5 5"/><path d="M12 3v12"/>', ext: '<path d="M15 3h6v6"/><path d="M10 14 21 3"/><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/>', building: '<rect x="4" y="3" width="16" height="18" rx="1.5"/><path d="M9 7h.01M15 7h.01M9 11h.01M15 11h.01M9 15h.01M15 15h.01M10 21v-3h4v3"/>', chat: '<path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/><path d="M8 9h8M8 13h5"/>', mail: '<rect x="2" y="4" width="20" height="16" rx="2"/><path d="m22 7-10 6L2 7"/>'
};
const svg = (k, s = 16, w = 2) => `<svg width="${s}" height="${s}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="${w}" stroke-linecap="round" stroke-linejoin="round">${IC[k]}</svg>`;
const PALETTE = ['#1A4E8A', '#F26522', '#16a34a', '#6d28d9', '#0f766e', '#db2777', '#0891b2', '#b45309', '#e5484d', '#475569'];
function hue(s) { let h = 0; for (const c of String(s)) h = (h * 31 + c.charCodeAt(0)) % 360; return h; }
function avatar(nom, sz = 34) { const ini = String(nom || '?').trim().split(/\s+/).map(w => w[0]).join('').slice(0, 2).toUpperCase(); const h = hue(nom); return `<span class="vx-av" style="width:${sz}px;height:${sz}px;font-size:${Math.round(sz * .36)}px;background:linear-gradient(135deg,hsl(${h},55%,55%),hsl(${(h + 35) % 360},55%,40%))">${esc(ini)}</span>`; }
function kwColor(k) { const h = hue(norm(k)); return `background:hsl(${h},70%,95%);color:hsl(${h},55%,32%);border-color:hsl(${h},60%,86%)`; }
function csvDownload(name, rows) {
  const L = a => a.map(v => { const s = String(v == null ? '' : v); return /[";\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; }).join(';');
  const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob(['﻿' + rows.map(L).join('\n')], { type: 'text/csv;charset=utf-8' })); a.download = name; a.click();
}
const errTxt = m => /inconnue/i.test(String(m || '')) ? "le serveur n'est pas encore à jour (mise en production en attente)" : String(m || 'erreur inconnue');
function modal(html, w = 640) {
  const m = document.createElement('div'); m.className = 'vx-modal vx';
  m.innerHTML = `<div class="box" style="max-width:${w}px">${html}</div>`;
  m.addEventListener('mousedown', e => { if (e.target === m) m.remove(); });
  document.body.appendChild(m); return m;
}

function injectCss() {
  if ($('vxCss')) return;
  const ln = document.createElement('link'); ln.rel = 'stylesheet'; ln.href = 'https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap'; document.head.appendChild(ln);
  const st = document.createElement('style'); st.id = 'vxCss';
  st.textContent = `
  .vx{--b:#1A4E8A;--b1:#e8eff8;--b0:#f3f7fc;--o:#F26522;--ink:#0e1a2b;--tx:#3b4859;--mu:#7d8a9c;--li:#e7ecf3;--li2:#f1f4f8;--ok:#16a34a;--red:#e5484d;font-family:'Plus Jakarta Sans',Inter,ui-sans-serif,system-ui,sans-serif;color:var(--tx);font-size:13.5px}
  .vx svg{display:inline-block;vertical-align:-2px;flex-shrink:0}
  .vx-hd{display:flex;flex-wrap:wrap;align-items:flex-end;justify-content:space-between;gap:14px;margin-bottom:20px}
  .vx-hd h1{margin:0;font-size:26px;font-weight:800;color:var(--ink);letter-spacing:-.025em}
  .vx-hd p{margin:4px 0 0;color:var(--mu);font-size:13.5px}
  .vx-btn{display:inline-flex;align-items:center;gap:7px;height:38px;padding:0 15px;border-radius:11px;border:1px solid var(--li);background:#fff;color:var(--tx);font:600 13px 'Plus Jakarta Sans',Inter,sans-serif;cursor:pointer;transition:all .18s;white-space:nowrap;box-shadow:0 1px 2px rgba(14,26,43,.04)}
  .vx-btn:hover{border-color:#cdd8e6;color:var(--b);transform:translateY(-1px);box-shadow:0 8px 18px -12px rgba(14,26,43,.45)}
  .vx-btn:active{transform:none}
  .vx-btn.pri{background:linear-gradient(180deg,#24589a,#1A4E8A);border-color:#164378;color:#fff;box-shadow:0 8px 18px -10px rgba(26,78,138,.9),inset 0 1px 0 rgba(255,255,255,.15)}
  .vx-btn.ora{background:linear-gradient(180deg,#f47637,#F26522);border-color:#e85d1c;color:#fff;box-shadow:0 8px 18px -10px rgba(242,101,34,.9),inset 0 1px 0 rgba(255,255,255,.2)}
  .vx-btn.sm{height:32px;padding:0 11px;font-size:12.5px;border-radius:9px}
  .vx-btn.ico{width:34px;padding:0;justify-content:center;height:34px}
  .vx-btn.danger:hover{color:var(--red);border-color:#f3c4c6}
  .vx-kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:14px;margin-bottom:18px}
  .vx-kpi{background:#fff;border:1px solid var(--li);border-radius:16px;padding:16px 18px;box-shadow:0 1px 2px rgba(14,26,43,.03),0 12px 30px -24px rgba(14,26,43,.3)}
  .vx-kpi span{font-size:11px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:var(--mu)}
  .vx-kpi b{display:block;font-size:26px;font-weight:800;color:var(--ink);letter-spacing:-.02em;margin-top:2px}
  .vx-kpi small{font-size:12px;color:var(--mu);font-weight:500}
  .vx-bar{display:flex;flex-wrap:wrap;align-items:center;gap:8px;background:#fff;border:1px solid var(--li);border-radius:16px;padding:10px 12px;margin-bottom:16px}
  .vx-search{position:relative;flex:1;min-width:200px}
  .vx-search svg{position:absolute;left:12px;top:50%;transform:translateY(-50%);color:var(--mu)}
  .vx .vx-search input{width:100%;height:36px;border:1px solid var(--li);border-radius:10px;padding:0 12px 0 35px;font:500 13px 'Plus Jakarta Sans',Inter,sans-serif;outline:none;background:#fff}
  .vx select.vx-sel{height:36px;border:1px solid var(--li);border-radius:10px;padding:0 30px 0 11px;font:600 12.5px 'Plus Jakarta Sans',Inter,sans-serif;color:var(--ink);background-color:#fff;width:auto;max-width:200px}
  .vx-seg{display:inline-flex;background:#eef1f5;border-radius:10px;padding:3px}
  .vx-seg button{display:inline-flex;align-items:center;gap:6px;height:30px;padding:0 10px;border:none;border-radius:8px;background:none;color:var(--mu);font:600 12.5px 'Plus Jakarta Sans',Inter,sans-serif;cursor:pointer}
  .vx-seg button.on{background:#fff;color:var(--b);box-shadow:0 1px 3px rgba(14,26,43,.12)}
  .vx-av{border-radius:50%;display:inline-flex;align-items:center;justify-content:center;color:#fff;font-weight:800;flex-shrink:0;box-shadow:inset 0 0 0 2px rgba(255,255,255,.45)}
  .vx-chip{display:inline-flex;align-items:center;gap:5px;padding:3px 9px;border-radius:999px;font-size:11.5px;font-weight:600;background:var(--li2);color:var(--tx);white-space:nowrap}
  .vx-chip.b{background:var(--b1);color:var(--b)}
  .vx-kw{display:inline-flex;align-items:center;gap:4px;padding:3px 10px;border-radius:999px;font-size:11.5px;font-weight:700;border:1px solid transparent;cursor:pointer;transition:transform .12s}
  .vx-kw:hover{transform:translateY(-1px)}
  .vx-kw.on{box-shadow:0 0 0 2px currentColor}
  .vx-cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(330px,1fr));gap:16px}
  .vx-card{position:relative;background:#fff;border:1px solid var(--li);border-radius:18px;padding:18px 18px 16px;box-shadow:0 1px 2px rgba(14,26,43,.03),0 14px 30px -26px rgba(14,26,43,.4);transition:all .2s;display:flex;flex-direction:column;gap:12px}
  .vx-card:hover{transform:translateY(-3px);box-shadow:0 22px 40px -24px rgba(26,78,138,.45);border-color:#d5e0ee}
  .vx-card .top{display:flex;align-items:center;gap:12px}
  .vx-card .nm{font-size:15px;font-weight:700;color:var(--ink);line-height:1.2}
  .vx-card .meta{display:flex;flex-wrap:wrap;gap:5px;margin-top:4px}
  .vx-card .sujet{font-size:14px;line-height:1.5;color:var(--ink);font-weight:500;display:-webkit-box;-webkit-line-clamp:3;-webkit-box-orient:vertical;overflow:hidden;min-height:42px}
  .vx-card .sujet.empty{color:#b3bdca;font-style:italic}
  .vx-card .kws{display:flex;flex-wrap:wrap;gap:6px}
  .vx-card .ft{display:flex;align-items:center;gap:8px;padding-top:12px;border-top:1px solid var(--li2);font-size:12.5px;color:var(--mu)}
  .vx-card .ft b{color:var(--ink);font-weight:600}
  .vx-card .acts{position:absolute;top:12px;right:12px;display:flex;gap:4px;opacity:0;transition:opacity .15s}
  .vx-card:hover .acts{opacity:1}
  .vx-st{font-size:11px;font-weight:700;padding:3px 9px;border-radius:999px}
  .vx-grp{display:flex;align-items:center;gap:10px;margin:22px 0 12px;font-size:13px;font-weight:700;color:var(--ink)}
  .vx-grp:first-child{margin-top:4px}
  .vx-grp span{font-size:12px;color:var(--mu);font-weight:600;background:var(--li2);padding:1px 8px;border-radius:999px}
  .vx-grp::after{content:'';flex:1;height:1px;background:var(--li)}
  .vx-table{width:100%;background:#fff;border:1px solid var(--li);border-radius:16px;overflow:hidden;border-collapse:separate;border-spacing:0}
  .vx-table th{text-align:left;font-size:11.5px;font-weight:700;color:var(--mu);padding:12px 14px;background:#fafbfd;border-bottom:1px solid var(--li);cursor:pointer;user-select:none;white-space:nowrap}
  .vx-table td{padding:12px 14px;border-bottom:1px solid var(--li2);font-size:13px;vertical-align:middle}
  .vx-table tr:last-child td{border-bottom:none}
  .vx-table tbody tr{cursor:pointer;transition:background .12s}
  .vx-table tbody tr:hover{background:#f8fafd}
  .vx-empty{background:#fff;border:1.5px dashed #d8e0ea;border-radius:18px;padding:48px 24px;text-align:center;color:var(--mu)}
  .vx-empty .ill{width:60px;height:60px;border-radius:18px;margin:0 auto 12px;display:flex;align-items:center;justify-content:center;background:linear-gradient(135deg,#eef3fa,#dfe9f6);color:var(--b)}
  .vx-empty h3{margin:0 0 4px;font-size:17px;font-weight:800;color:var(--ink)}
  .vx-modal{position:fixed;inset:0;z-index:10000;background:rgba(14,26,43,.42);backdrop-filter:blur(7px);display:flex;align-items:center;justify-content:center;padding:20px}
  .vx-modal .box{width:100%;max-height:92vh;overflow:auto;background:#fff;border-radius:22px;padding:28px;box-shadow:0 40px 90px -24px rgba(14,26,43,.5);animation:vxIn .2s}
  @keyframes vxIn{from{opacity:0;transform:translateY(8px) scale(.99)}to{opacity:1;transform:none}}
  .vx-modal h2{margin:0;font-size:20px;font-weight:800;color:var(--ink);letter-spacing:-.02em}
  .vx-modal .sub{margin:4px 0 0;font-size:13px;color:var(--mu)}
  .vx .fl{display:block;font-size:11px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:var(--mu);margin:16px 0 7px}
  .vx .in{width:100%;min-height:42px;border:1px solid var(--li);border-radius:11px;padding:10px 12px;font:500 14px 'Plus Jakarta Sans',Inter,sans-serif;color:var(--ink);outline:none;background:#fff;transition:all .15s;text-transform:none;letter-spacing:0}
  .vx .in:focus{border-color:var(--b);box-shadow:0 0 0 4px rgba(26,78,138,.1)}
  .vx textarea.in{min-height:90px;resize:vertical;line-height:1.5}
  .vx .g2{display:grid;grid-template-columns:1fr 1fr;gap:0 14px}
  .vx .mft{display:flex;justify-content:flex-end;gap:10px;margin-top:24px;flex-wrap:wrap}
  .vx-ac{position:relative}
  .vx-acl{position:absolute;left:0;right:0;top:calc(100% + 6px);z-index:5;background:#fff;border:1px solid var(--li);border-radius:12px;box-shadow:0 20px 50px -16px rgba(14,26,43,.35);max-height:260px;overflow:auto;padding:5px}
  .vx-acl .it{display:flex;align-items:center;gap:10px;padding:8px 10px;border-radius:9px;cursor:pointer;font-size:13px;font-weight:500;color:var(--ink)}
  .vx-acl .it:hover,.vx-acl .it.act{background:var(--b0);color:var(--b)}
  .vx-acl .it small{margin-left:auto;color:var(--mu);font-size:11.5px}
  .vx-tags{display:flex;flex-wrap:wrap;gap:6px;align-items:center;min-height:42px;border:1px solid var(--li);border-radius:11px;padding:6px 8px;cursor:text;background:#fff}
  .vx-tags:focus-within{border-color:var(--b);box-shadow:0 0 0 4px rgba(26,78,138,.1)}
  .vx-tags input{flex:1;min-width:120px;border:none;outline:none;font:500 14px 'Plus Jakarta Sans',Inter,sans-serif;padding:4px}
  .vx-tags .vx-kw button{border:none;background:none;color:inherit;cursor:pointer;padding:0;margin-left:2px;opacity:.6}
  .vx-sugg{display:flex;flex-wrap:wrap;gap:6px;margin-top:8px}
  .vx-sugg .vx-kw{opacity:.85}
  .vx-fixed{display:flex;gap:8px;margin-top:10px}

  /* frises */
  .fr-list{display:grid;grid-template-columns:repeat(auto-fill,minmax(270px,1fr));gap:14px;margin-bottom:22px}
  .fr-item{position:relative;background:#fff;border:1px solid var(--li);border-radius:16px;padding:16px 16px 14px;cursor:pointer;transition:all .2s;overflow:hidden}
  .fr-item::before{content:'';position:absolute;left:0;top:0;bottom:0;width:4px;background:var(--c)}
  .fr-item:hover{transform:translateY(-2px);box-shadow:0 18px 34px -24px rgba(14,26,43,.45)}
  .fr-item.on{border-color:var(--c);box-shadow:0 0 0 3px color-mix(in srgb,var(--c) 18%,transparent),0 18px 34px -24px rgba(14,26,43,.45)}
  .fr-item h4{margin:0;font-size:14.5px;font-weight:700;color:var(--ink)}
  .fr-item .nx{margin-top:8px;font-size:12.5px;color:var(--mu);display:flex;align-items:center;gap:6px}
  .fr-prog{height:6px;border-radius:999px;background:var(--li2);margin-top:12px;overflow:hidden}
  .fr-prog i{display:block;height:100%;border-radius:999px;background:var(--c)}
  .fr-board{background:#fff;border:1px solid var(--li);border-radius:20px;padding:24px 26px;box-shadow:0 1px 2px rgba(14,26,43,.03),0 18px 40px -30px rgba(14,26,43,.35)}
  .fr-head{display:flex;flex-wrap:wrap;align-items:flex-start;gap:14px;justify-content:space-between}
  .fr-head h2{margin:0;font-size:21px;font-weight:800;color:var(--ink);letter-spacing:-.02em;display:flex;align-items:center;gap:10px}
  .fr-head h2 i{width:12px;height:12px;border-radius:50%;background:var(--c)}
  .fr-next{display:flex;align-items:center;gap:14px;margin-top:18px;padding:14px 16px;border-radius:14px;background:linear-gradient(90deg,color-mix(in srgb,var(--c) 9%,#fff),#fff)}
  .fr-cd{min-width:74px;text-align:center;padding:8px 10px;border-radius:12px;background:var(--c);color:#fff;box-shadow:0 10px 20px -12px var(--c)}
  .fr-cd b{display:block;font-size:22px;font-weight:800;line-height:1}
  .fr-cd span{font-size:10.5px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;opacity:.9}
  .fr-scroll{overflow-x:auto;padding:12px 4px 6px;margin-top:10px}
  .fr-track{position:relative;min-width:100%;height:340px}
  .fr-line{position:absolute;left:0;right:0;top:170px;height:4px;border-radius:4px;background:var(--li)}
  .fr-line i{position:absolute;left:0;top:0;bottom:0;border-radius:4px;background:var(--c)}
  .fr-today{position:absolute;top:120px;height:104px;width:0;border-left:2px dashed var(--o)}
  .fr-today span{position:absolute;top:-22px;left:50%;transform:translateX(-50%);font-size:10.5px;font-weight:800;color:var(--o);background:#fff4ed;border-radius:999px;padding:2px 8px;white-space:nowrap}
  .fr-pt{position:absolute;top:160px;width:24px;height:24px;margin-left:-12px;border-radius:50%;background:#fff;border:3px solid var(--c);box-shadow:0 4px 10px -4px rgba(14,26,43,.4);cursor:pointer;display:flex;align-items:center;justify-content:center;color:#fff;transition:transform .15s;z-index:2}
  .fr-pt:hover{transform:scale(1.15)}
  .fr-pt.done{background:var(--ok);border-color:var(--ok)}
  .fr-pt.late{border-color:var(--red);background:#fff1f1}
  .fr-pt.next{box-shadow:0 0 0 6px color-mix(in srgb,var(--c) 20%,transparent)}
  .fr-card{position:absolute;width:200px;margin-left:-100px;background:#fff;border:1px solid var(--li);border-radius:14px;padding:11px 13px;box-shadow:0 14px 28px -20px rgba(14,26,43,.45);font-size:12.5px}
  .fr-card.up{bottom:200px}.fr-card.dn{top:200px}
  .fr-card::after{content:'';position:absolute;left:50%;width:2px;height:16px;background:var(--li);transform:translateX(-50%)}
  .fr-card.up::after{top:100%}.fr-card.dn::after{bottom:100%}
  .fr-card .d{font-size:11px;font-weight:800;letter-spacing:.04em;text-transform:uppercase;color:var(--c)}
  .fr-card .t{font-size:13.5px;font-weight:700;color:var(--ink);margin-top:2px;line-height:1.3}
  .fr-card .r{color:var(--mu);margin-top:4px;line-height:1.35}
  .fr-card.done{opacity:.6}.fr-card.done .t{text-decoration:line-through}
  .fr-card .j{display:inline-block;margin-top:7px;font-size:10.5px;font-weight:800;padding:2px 8px;border-radius:999px;background:var(--li2);color:var(--mu)}
  .fr-card .j.soon{background:#fff4ed;color:#c2410c}.fr-card .j.late{background:#fde8e8;color:#c62828}.fr-card .j.ok{background:#e8f6ee;color:#15803d}
  .fr-steps{display:flex;flex-direction:column;gap:8px}
  .fr-step{display:grid;grid-template-columns:150px 1fr 1fr 34px;gap:8px;align-items:center}
  .fr-rem{display:flex;flex-wrap:wrap;gap:6px}
  .fr-rem label{display:inline-flex;align-items:center;gap:6px;padding:6px 11px;border-radius:999px;border:1px solid var(--li);font-size:12.5px;font-weight:600;cursor:pointer;user-select:none;text-transform:none;letter-spacing:0;color:var(--tx);margin:0}
  .fr-rem input{display:none}
  .fr-rem label:has(input:checked){background:var(--b);border-color:var(--b);color:#fff}
  .fr-pp{border:1px solid var(--li);border-radius:13px;padding:12px;background:#fff}
  .fr-tabs{display:flex;flex-wrap:wrap;gap:6px;align-items:center}
  .fr-tabs button{padding:6px 12px;border-radius:999px;border:1px solid var(--li);background:#fff;font:600 12.5px 'Plus Jakarta Sans',Inter,sans-serif;cursor:pointer;color:var(--tx)}
  .fr-tabs button.on{background:var(--b);border-color:var(--b);color:#fff}
  .fr-sel{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:10px}
  .fr-sel:empty{display:none}
  .fr-sel .vx-chip{padding:4px 6px 4px 10px;font-size:12px}
  .fr-sel .vx-chip.etu{background:#e8f0fb;color:#1A4E8A}.fr-sel .vx-chip.equipe{background:#fff1e8;color:#c2410c}.fr-sel .vx-chip.form{background:#e8f6ee;color:#15803d}
  .fr-sel .vx-chip button{border:0;background:rgba(0,0,0,.06);border-radius:50%;width:18px;height:18px;line-height:16px;cursor:pointer;color:inherit;font-size:13px;padding:0}
  .fr-res{max-height:220px;overflow:auto;margin-top:8px;display:flex;flex-direction:column;gap:1px}
  .fr-res .o{display:flex;align-items:center;gap:9px;padding:6px 8px;border-radius:9px;cursor:pointer;font-size:13px;color:var(--ink)}
  .fr-res .o:hover{background:var(--li2)}
  .fr-res .o small{color:var(--mu);margin-left:auto;text-align:right;max-width:45%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .fr-res .o .plus{color:var(--b);font-weight:800}
  .vx-dot{width:22px;height:22px;border-radius:50%;cursor:pointer;border:2px solid #fff;box-shadow:0 0 0 1px var(--li)}
  .vx-dot.on{box-shadow:0 0 0 2px var(--ink)}
  @media(max-width:760px){.vx .g2{grid-template-columns:1fr}.fr-step{grid-template-columns:1fr 1fr}}

  /* ═════ v2 : plus aéré, dégradés & ombres (nov. 2026) ═════ */
  .vx{font-size:14px}
  .vx-hd{margin-bottom:28px;padding:28px 32px;border-radius:26px;background:linear-gradient(135deg,rgba(255,255,255,.95),rgba(242,246,252,.9));border:1px solid #e3eaf3;box-shadow:0 1px 0 #fff inset,0 30px 60px -42px rgba(26,78,138,.5);position:relative;overflow:hidden}
  .vx-hd::before{content:'';position:absolute;left:0;right:0;top:0;height:4px;background:linear-gradient(90deg,#1A4E8A,#3f7cc4 45%,#F26522)}
  .vx-hd h1{font-size:30px}
  .vx-btn{height:42px;padding:0 18px;border-radius:13px;background:linear-gradient(180deg,#fff,#f2f5fa);border-color:#dfe6f0;box-shadow:0 1px 0 #fff inset,0 8px 16px -12px rgba(14,26,43,.4)}
  .vx-btn:hover{background:linear-gradient(180deg,#fff,#e9eff8)}
  .vx-btn.pri{background:linear-gradient(180deg,#2d64a8,#1A4E8A 60%,#163f73);box-shadow:0 1px 0 rgba(255,255,255,.2) inset,0 14px 26px -12px rgba(26,78,138,.85)}
  .vx-btn.ora{background:linear-gradient(180deg,#f7844a,#F26522 60%,#e2571a);box-shadow:0 1px 0 rgba(255,255,255,.25) inset,0 14px 26px -12px rgba(242,101,34,.9)}
  .vx-btn.sm{height:36px}.vx-btn.ico{width:38px;height:38px}
  .vx-kpis{gap:18px;margin-bottom:24px}
  .vx-kpi{padding:20px 22px;border-radius:20px;background:linear-gradient(160deg,#ffffff,#f3f6fb);border-color:#e3eaf3;box-shadow:0 1px 0 #fff inset,0 20px 40px -32px rgba(14,26,43,.5)}
  .vx-kpi b{font-size:30px;background:linear-gradient(90deg,#1A4E8A,#3f7cc4);-webkit-background-clip:text;background-clip:text;color:transparent}
  .vx-bar{padding:14px 16px;border-radius:20px;gap:10px;background:linear-gradient(180deg,rgba(255,255,255,.95),rgba(245,248,252,.95));border-color:#e3eaf3;box-shadow:0 18px 36px -30px rgba(14,26,43,.45);margin-bottom:20px}
  .vx .vx-search input,.vx select.vx-sel{height:40px;border-radius:12px;background:linear-gradient(180deg,#fff,#f6f8fc);border-color:#dfe6f0}
  .vx-cards{gap:22px}
  .vx-card{padding:24px 24px 20px;border-radius:22px;gap:14px;background:linear-gradient(165deg,#ffffff 0%,#f6f9fd 100%);border-color:#e3eaf3;box-shadow:0 1px 0 #fff inset,0 22px 44px -34px rgba(14,26,43,.55)}
  .vx-card::before{content:'';position:absolute;left:24px;right:24px;top:0;height:3px;border-radius:0 0 4px 4px;background:linear-gradient(90deg,#1A4E8A,#F26522);opacity:.0;transition:opacity .2s}
  .vx-card:hover::before{opacity:1}
  .vx-card .nm{font-size:16px}
  .vx-card .sujet{font-size:14.5px}
  .vx-card .ft{padding-top:14px}
  .vx-docchip{display:flex;align-items:center;gap:7px;font-size:12.5px;color:#475569;padding:8px 12px;border-radius:12px;background:linear-gradient(90deg,#eef3fa,#f7f9fc)}
  .vx-docchip b{color:var(--b)}
  .vx-grp{margin:30px 0 14px;font-size:14px}
  .vx-table{border-radius:20px;background:linear-gradient(180deg,#fff,#f8fafd);box-shadow:0 20px 40px -32px rgba(14,26,43,.5)}
  .vx-table td{padding:15px 16px}
  .vx-empty{border-radius:24px;padding:60px 28px;background:linear-gradient(180deg,rgba(255,255,255,.9),rgba(244,247,251,.9))}
  .vx-modal .box{border-radius:28px;padding:36px 40px;background:linear-gradient(180deg,#ffffff,#f5f8fc);border:1px solid #e6ecf4}
  .vx-modal h2{font-size:24px}
  .vx .in{min-height:48px;border-radius:13px;background:linear-gradient(180deg,#fff,#f8fafd);border-color:#dfe6f0;font-size:14.5px}
  .vx .g2{gap:0 18px}
  .vx .g3{display:grid;grid-template-columns:110px 170px 1fr;gap:0 12px}
  .vx-tags{min-height:50px;border-radius:13px;background:linear-gradient(180deg,#fff,#f8fafd);border-color:#dfe6f0}
  .vx-note{display:flex;align-items:center;gap:8px;font-size:13px;color:#64748b;padding:12px 14px;border-radius:12px;background:linear-gradient(90deg,#f1f5fa,#f8fafc);border:1px dashed #d8e1ec}
  .vx-docs{display:flex;flex-direction:column;gap:8px;margin-bottom:12px;position:relative}
  .vx-doc{display:flex;align-items:center;gap:14px;padding:12px 14px;border-radius:14px;background:linear-gradient(180deg,#fff,#f6f9fc);border:1px solid #e3eaf3;box-shadow:0 10px 20px -18px rgba(14,26,43,.5)}
  .vx-doc .ver{min-width:46px;height:34px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:13px;color:#fff;background:linear-gradient(135deg,#3f7cc4,#1A4E8A);box-shadow:0 8px 14px -8px rgba(26,78,138,.8)}
  .vx-doc:first-child .ver{background:linear-gradient(135deg,#f7844a,#F26522);box-shadow:0 8px 14px -8px rgba(242,101,34,.9)}
  .vx-doc .in2{flex:1;min-width:0}
  .vx-doc a{display:block;font-weight:700;color:var(--ink);text-decoration:none;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .vx-doc a:hover{color:var(--b)}
  .vx-doc small{font-size:12px;color:var(--mu)}
  .vx-drop .dz{display:flex;align-items:center;gap:14px;padding:18px 20px;border-radius:16px;border:2px dashed #cdd9e8;background:linear-gradient(135deg,#f5f8fd,#fbfcfe);color:var(--b);cursor:pointer;transition:all .18s}
  .vx-drop .dz:hover,.vx-drop .dz.over{border-color:var(--o);background:linear-gradient(135deg,#fff4ed,#fffaf6);color:var(--o)}
  .vx-drop .dz b{display:block;font-size:14px;color:var(--ink)}
  .vx-drop .dz span{font-size:12.5px;color:var(--mu)}
  .vx-upl{padding:16px 18px;border-radius:16px;background:linear-gradient(180deg,#f7f9fc,#eff3f9);border:1px solid #e3eaf3}
  .vx-upl .fn{display:flex;align-items:center;gap:8px;font-size:13.5px;color:var(--ink)}
  .vx-upl .fn span{color:var(--mu);font-size:12px}
  .vx-acl{border-radius:14px;background:linear-gradient(180deg,#fff,#f8fafd);max-height:300px}
  .vx-tabs{display:inline-flex;gap:6px;padding:5px;margin:-6px 0 22px;border-radius:16px;background:linear-gradient(180deg,#e9eef5,#eef2f8);box-shadow:inset 0 1px 2px rgba(14,26,43,.06)}
  .vx-tabs button{display:inline-flex;align-items:center;gap:8px;height:40px;padding:0 18px;border:none;border-radius:12px;background:none;font:700 13.5px 'Plus Jakarta Sans',Inter,sans-serif;color:var(--mu);cursor:pointer;transition:all .18s}
  .vx-tabs button.on{background:linear-gradient(180deg,#fff,#f6f9fd);color:var(--b);box-shadow:0 6px 16px -10px rgba(26,78,138,.6)}
  .vx-lastdoc{display:inline-flex;align-items:center;gap:7px;padding:4px 10px 4px 4px;border-radius:999px;background:linear-gradient(90deg,#eef3fa,#f7f9fc);border:1px solid #dce7f5;color:var(--b);font-size:12px;font-weight:700;text-decoration:none;white-space:nowrap;transition:all .15s}
  .vx-lastdoc:hover{background:linear-gradient(90deg,#1A4E8A,#2d64a8);color:#fff}
  .vx-lastdoc .ver{padding:2px 7px;border-radius:999px;background:linear-gradient(135deg,#f7844a,#F26522);color:#fff;font-size:11px}
  .ck-top{display:flex;flex-wrap:wrap;align-items:center;gap:14px;margin-bottom:18px}
  .ck-year{display:flex;align-items:center;gap:10px}
  .ck-year label{font-size:11px;font-weight:800;letter-spacing:.08em;text-transform:uppercase;color:var(--mu)}
  .ck-arch,.ck-live{display:inline-flex;align-items:center;gap:7px;font-size:12.5px;font-weight:700;padding:7px 13px;border-radius:999px}
  .ck-arch{background:linear-gradient(90deg,#f1f5fa,#e9eef5);color:#475569}
  .ck-live{background:linear-gradient(90deg,#e8f6ee,#f2fbf5);color:#15803d}
  .ck-dips{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:16px;margin-bottom:20px}
  .ck-dip{text-align:left;display:flex;flex-direction:column;gap:6px;padding:18px 20px;border-radius:20px;border:1px solid #e3eaf3;background:linear-gradient(160deg,#fff,#f4f7fb);box-shadow:0 1px 0 #fff inset,0 18px 34px -28px rgba(14,26,43,.5);cursor:pointer;font-family:inherit;transition:all .2s}
  .ck-dip:hover{transform:translateY(-2px)}
  .ck-dip.on{background:linear-gradient(150deg,#24589a,#1A4E8A 60%,#163f73);border-color:#1A4E8A;box-shadow:0 22px 38px -20px rgba(26,78,138,.85)}
  .ck-dip .t{font-size:17px;font-weight:800;color:var(--ink)}
  .ck-dip .v{font-size:13px;color:var(--mu)}.ck-dip .v b{font-size:20px;color:var(--ink)}
  .ck-dip .bar{height:7px;border-radius:99px;background:#e6ebf2;overflow:hidden}
  .ck-dip .bar i{display:block;height:100%;border-radius:99px;background:linear-gradient(90deg,#F26522,#16a34a)}
  .ck-dip .p{font-size:11.5px;color:var(--mu)}
  .ck-dip.on .t,.ck-dip.on .v b{color:#fff}.ck-dip.on .v,.ck-dip.on .p{color:rgba(255,255,255,.75)}.ck-dip.on .bar{background:rgba(255,255,255,.2)}
  .ck-card{padding:22px 24px;border-radius:22px;background:linear-gradient(175deg,#fff,#f6f9fd);border:1px solid #e3eaf3;box-shadow:0 1px 0 #fff inset,0 24px 46px -36px rgba(14,26,43,.55);margin-bottom:18px}
  .ck-h{display:flex;flex-wrap:wrap;align-items:center;justify-content:space-between;gap:10px;margin-bottom:12px}
  .ck-h h3{margin:0;font-size:16px;font-weight:800;color:var(--ink);display:flex;align-items:center;gap:10px}
  .ck-h h3 span{font-size:12px;padding:2px 9px;border-radius:99px;background:#e9eef6;color:#516178}
  .ck-legend{display:flex;gap:6px}.ck-legend span{font-size:11px;font-weight:700;color:var(--mu);padding:4px 10px;border-radius:99px;background:#eef2f7}
  .ck-t{width:100%;border-collapse:separate;border-spacing:0 6px}
  .ck-t th{text-align:left;font-size:11px;font-weight:800;letter-spacing:.06em;text-transform:uppercase;color:#8a96a8;padding:4px 12px}
  .ck-t td{padding:12px;background:linear-gradient(180deg,#fff,#fafcfe);border-top:1px solid #e9eef5;border-bottom:1px solid #e9eef5;vertical-align:middle}
  .ck-t td:first-child{border-left:1px solid #e9eef5;border-radius:14px 0 0 14px}.ck-t td:last-child{border-right:1px solid #e9eef5;border-radius:0 14px 14px 0}
  .ck-t tr.full td{background:linear-gradient(180deg,#f3fbf6,#ecf8f1);border-color:#d3eedc}
  .ck-etu{display:flex;align-items:center;gap:10px;cursor:pointer}
  .ck-etu b{display:block;font-size:13.5px;color:var(--ink)}.ck-etu small{font-size:11.5px;color:var(--mu)}
  .ck-etu:hover b{color:var(--b)}
  .ck-ok{display:inline-flex;align-items:center;gap:5px;padding:4px 8px;border-radius:99px;font-size:11.5px;font-weight:800;vertical-align:middle}
  .ck-ok.y{background:linear-gradient(180deg,#22c55e,#16a34a);color:#fff;box-shadow:0 6px 12px -8px rgba(22,163,74,.9)}
  .ck-ok.n{background:#fde8e8;color:#c62828}
  .ck-sub{display:block;margin-top:5px;font-size:12px;color:var(--tx);max-width:260px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .ck-t .vx-lastdoc{margin-left:6px}
  .ck-box input,.ck-pick input,.ck-cyc input{display:none}
  .ck-box span,.ck-pick .b{width:22px;height:22px;border-radius:7px;border:1.8px solid #c3ccd8;display:inline-flex;align-items:center;justify-content:center;color:transparent;cursor:pointer;background:#fff;flex-shrink:0}
  .ck-box input:checked+span{background:linear-gradient(180deg,#2d64a8,#1A4E8A);border-color:#1A4E8A;color:#fff}
  .ck-cyc{display:inline-flex;align-items:center;gap:9px;font-size:12.5px;font-weight:700;color:var(--mu);cursor:pointer;user-select:none;text-transform:none;letter-spacing:0;margin:0}
  .ck-cyc .sw{width:36px;height:21px;border-radius:99px;background:#cfd8e4;position:relative;transition:background .2s;flex-shrink:0}
  .ck-cyc .sw::after{content:'';position:absolute;top:2px;left:2px;width:17px;height:17px;border-radius:50%;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,.25);transition:transform .2s}
  .ck-cyc.on{color:#15803d}.ck-cyc.on .sw{background:#16a34a}.ck-cyc.on .sw::after{transform:translateX(15px)}
  .ck-prog{display:inline-block;width:70px;height:7px;border-radius:99px;background:#e6ebf2;overflow:hidden;vertical-align:middle;margin-right:8px}
  .ck-prog i{display:block;height:100%;background:linear-gradient(90deg,#F26522,#f59e0b)}
  .ck-prog.p4 i{background:linear-gradient(90deg,#22c55e,#16a34a)}
  .ck-n{font-size:12.5px;color:var(--ink)}.ck-n.ok{color:#15803d}
  .ck-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(250px,1fr));gap:10px}
  .ck-pick{display:flex;align-items:center;gap:10px;padding:10px 12px;border-radius:14px;border:1px solid #e3eaf3;background:linear-gradient(180deg,#fff,#f8fafd);cursor:pointer;transition:all .15s;text-transform:none;letter-spacing:0;margin:0}
  .ck-pick:hover{border-color:#bcd0e8;box-shadow:0 10px 20px -16px rgba(26,78,138,.6)}
  .ck-pick .n b{display:block;font-size:13px;color:var(--ink)}.ck-pick .n small{font-size:11.5px;color:var(--mu)}
  .ck-pick input:checked+.b{background:linear-gradient(180deg,#2d64a8,#1A4E8A);border-color:#1A4E8A;color:#fff}
  .ck-pick:has(input:checked){border-color:#1A4E8A;background:linear-gradient(180deg,#eef4fc,#e5eefa)}
  .ck-pick .m{font-size:11px;font-weight:700;margin-top:2px}.ck-pick .m.y{color:#15803d}.ck-pick .m.n{color:#a3aebd}
  .ck-filt{display:flex;flex-wrap:wrap;align-items:center;gap:10px;margin-bottom:12px}
  .ck-mail{display:block;font-size:11px;color:#1d4ed8;text-decoration:none;margin-top:2px;max-width:240px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .ck-mail:hover{text-decoration:underline}.ck-mail.perso{color:#94a3b8;font-style:italic}
  .ck-rap{display:inline-block;margin-left:6px;font-size:10px;font-weight:800;padding:1px 7px;border-radius:99px;background:#fff4ed;color:#c2410c;vertical-align:middle}
  @media(max-width:1100px){.ck-dips{grid-template-columns:repeat(2,minmax(0,1fr))}}
  .vx-kw-new{background:linear-gradient(135deg,#fff4ed,#ffe9dc);color:#c2410c;border-color:#fcd2b6;font-weight:800}
  .vx-link{display:inline-flex;align-items:center;gap:6px;margin-top:8px;border:none;background:none;padding:2px 0;font:600 12.5px 'Plus Jakarta Sans',Inter,sans-serif;color:var(--o);cursor:pointer}
  .vx-link:hover{text-decoration:underline}
  .vx-autre{grid-template-columns:1fr 1fr;gap:10px;margin-top:10px;padding:12px;border-radius:14px;background:linear-gradient(135deg,#fff4ed,#fffaf6);border:1px solid #fcd9c2}
  /* frises */
  .fr-list{gap:18px;margin-bottom:28px}
  .fr-item{padding:20px 22px 18px;border-radius:20px;background:linear-gradient(160deg,#ffffff,color-mix(in srgb,var(--c) 5%,#f6f8fc));border-color:#e3eaf3;box-shadow:0 1px 0 #fff inset,0 18px 36px -30px rgba(14,26,43,.5)}
  .fr-item::before{width:5px}
  .fr-item h4{font-size:15.5px}
  .fr-prog{height:8px;margin-top:14px;background:#e9eef5}
  .fr-prog i{background:linear-gradient(90deg,color-mix(in srgb,var(--c) 70%,#fff),var(--c))}
  .fr-board{padding:34px 38px;border-radius:28px;background:linear-gradient(170deg,#ffffff 0%,#f6f9fd 60%,color-mix(in srgb,var(--c) 5%,#f4f7fb) 100%);border-color:#e3eaf3;box-shadow:0 1px 0 #fff inset,0 40px 80px -56px rgba(14,26,43,.6)}
  .fr-head h2{font-size:25px}
  .fr-head h2 i{width:14px;height:14px;box-shadow:0 0 0 5px color-mix(in srgb,var(--c) 18%,transparent)}
  .fr-next{margin-top:24px;padding:18px 22px;border-radius:18px;background:linear-gradient(100deg,color-mix(in srgb,var(--c) 12%,#fff),rgba(255,255,255,.6) 70%);border:1px solid color-mix(in srgb,var(--c) 18%,#e3eaf3);box-shadow:0 18px 34px -28px var(--c)}
  .fr-cd{min-width:86px;padding:12px;border-radius:16px;background:linear-gradient(160deg,color-mix(in srgb,var(--c) 75%,#fff),var(--c));box-shadow:0 14px 26px -12px var(--c)}
  .fr-cd b{font-size:26px}
  .fr-scroll{padding:22px 8px 12px;margin-top:18px}
  .fr-track{height:420px}
  .fr-line{top:206px;height:6px;background:linear-gradient(90deg,#e6ebf2,#dfe6ef)}
  .fr-line i{background:linear-gradient(90deg,color-mix(in srgb,var(--c) 50%,#fff),var(--c));opacity:.55 !important}
  .fr-today{top:150px;height:120px;border-left:2px dashed var(--o)}
  .fr-today span{top:auto;bottom:-26px;background:linear-gradient(180deg,#fff4ed,#ffe9dc);box-shadow:0 6px 12px -8px rgba(242,101,34,.8)}
  .fr-pt{top:194px;width:30px;height:30px;margin-left:-15px;border-width:4px;box-shadow:0 8px 16px -6px rgba(14,26,43,.45)}
  .fr-pt.done{background:linear-gradient(160deg,#22c55e,#16a34a)}
  .fr-pt.next{box-shadow:0 0 0 8px color-mix(in srgb,var(--c) 20%,transparent),0 8px 16px -6px rgba(14,26,43,.45);animation:frPulse 2.2s ease-in-out infinite}
  @keyframes frPulse{50%{box-shadow:0 0 0 14px color-mix(in srgb,var(--c) 8%,transparent),0 8px 16px -6px rgba(14,26,43,.45)}}
  .fr-card{width:230px;margin-left:-115px;padding:15px 17px;border-radius:18px;background:linear-gradient(170deg,#ffffff,#f6f9fd);border-color:#e3eaf3;box-shadow:0 1px 0 #fff inset,0 20px 36px -24px rgba(14,26,43,.55);font-size:13px}
  .fr-card.up{bottom:242px}.fr-card.dn{top:248px}
  .fr-card::after{height:22px;width:2px;background:linear-gradient(180deg,#dfe6ef,#cbd6e4)}
  .fr-card .t{font-size:14.5px}
  .fr-card .j{margin-top:9px;padding:3px 10px}
  .fr-steps{gap:10px}
  .fr-step{grid-template-columns:170px 1fr 1fr 40px;gap:10px;padding:10px;border-radius:14px;background:linear-gradient(180deg,#fff,#f6f9fc);border:1px solid #e3eaf3}
  .fr-rem{gap:8px}
  .fr-rem label{padding:8px 14px;background:linear-gradient(180deg,#fff,#f5f8fc);border-color:#dfe6f0}
  .fr-rem label:has(input:checked){background:linear-gradient(180deg,#2d64a8,#1A4E8A);box-shadow:0 8px 16px -10px rgba(26,78,138,.8)}
  .vx-dot{width:26px;height:26px}
  #toast{z-index:20000 !important}
  .vx-vgrp{border:1px solid #e3eaf3;border-radius:16px;padding:10px;background:#fbfcfe;display:flex;flex-direction:column;gap:6px}
  .vx-vgrp .vh{display:flex;align-items:center;gap:10px;padding:2px 4px 4px}
  .vx-vgrp .vh .ver{min-width:46px;height:30px;border-radius:9px;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:13px;color:#fff;background:linear-gradient(135deg,#3f7cc4,#1A4E8A)}
  .vx-vgrp .vh b{color:var(--ink);font-size:13px}.vx-vgrp .vh small{color:var(--mu);font-size:12px}
  .vx-vgrp .vx-doc{padding:9px 12px;box-shadow:none}
  .vx-vgrp .vx-doc > svg{color:#1A4E8A;flex:none}
  .vx-fsel{display:flex;align-items:center;gap:8px;padding:6px 0;border-bottom:1px dashed #dfe6ef}
  .vx-fsel b{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .vx-verpick{display:flex;flex-wrap:wrap;align-items:center;gap:6px;margin-top:10px;font-size:12px;color:var(--mu);font-weight:600}
  .vx-verpick button{font:700 12px 'Plus Jakarta Sans',Inter,sans-serif;padding:4px 11px;border-radius:999px;border:1px solid #dfe6f0;background:#fff;color:#1A4E8A;cursor:pointer}
  .vx-verpick button.on{background:#1A4E8A;color:#fff;border-color:#1A4E8A}
  .vx-ent{display:inline-flex;align-items:center;gap:6px;font-size:12.5px;font-weight:700;color:#0f766e;background:#e6f6f3;border:1px solid #c7ebe4;padding:3px 10px;border-radius:999px;max-width:100%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;align-self:flex-start}
  /* regroupement par diplôme + carnet d'échanges */
  .vx-dip{display:flex;align-items:center;flex-wrap:wrap;gap:10px;margin:30px 0 4px;padding:12px 18px;border-radius:16px;background:linear-gradient(90deg,#1A4E8A,#2d64a8);color:#fff;box-shadow:0 14px 26px -20px rgba(26,78,138,.9)}
  .vx-dip:first-child{margin-top:4px}
  .vx-dip .t{font-size:16px;font-weight:800;letter-spacing:.2px}
  .vx-dip .n{font-size:12px;font-weight:700;background:rgba(255,255,255,.18);padding:2px 10px;border-radius:999px}
  .vx-dip .w{font-size:11.5px;font-weight:700;background:#fff4ed;color:#c2410c;padding:2px 10px;border-radius:999px}
  .vx-dip + .vx-table{margin-top:12px}
  .vx-echchip{background:linear-gradient(90deg,#f3effd,#f8f6fe);color:#4c1d95;flex-wrap:wrap}
  .vx-echchip .tx{flex-basis:100%;color:#64748b;font-size:12px;font-style:italic}
  .vx-echn{display:inline-flex;align-items:center;gap:5px;font-size:11.5px;font-weight:700;color:#6d28d9;background:#f3effd;padding:2px 9px;border-radius:999px;white-space:nowrap}
  .vx-ech{border:1px solid var(--li);border-radius:16px;padding:12px;background:#fbfcfe;margin-bottom:6px}
  .vx-echf .r{display:grid;grid-template-columns:170px 1fr;gap:8px;margin-bottom:8px}
  .vx-echf textarea{min-height:64px}
  .vx-echf .a{display:flex;align-items:center;justify-content:space-between;gap:8px;margin-top:8px}
  .vx-echf .a .sub{font-size:11.5px;color:var(--mu);margin:0}
  .vx-echl{display:flex;flex-direction:column;gap:8px;margin-top:12px;max-height:340px;overflow:auto;padding-right:2px}
  .vx-echi{background:#fff;border:1px solid var(--li);border-left:4px solid var(--c);border-radius:12px;padding:9px 12px}
  .vx-echi .h{display:flex;align-items:center;gap:8px;font-size:12px}
  .vx-echi .h .ty{font-weight:800;color:var(--c);background:color-mix(in srgb,var(--c) 10%,#fff);padding:1px 8px;border-radius:999px}
  .vx-echi .h b{color:var(--ink)}
  .vx-echi .h .by{color:var(--mu);flex:1}
  .vx-echi .tx{margin-top:5px;font-size:13px;color:var(--tx);line-height:1.5;white-space:normal}
  .ck-ech{display:inline-flex;align-items:center;gap:5px;margin-top:4px;font:600 11.5px 'Plus Jakarta Sans',Inter,sans-serif;color:#7d8a9c;background:#f1f4f8;border:1px dashed #d5dde8;border-radius:999px;padding:2px 9px;cursor:pointer}
  .ck-ech.on{color:#6d28d9;background:#f3effd;border:1px solid #e4dafb}
  .ck-ech:hover{filter:brightness(.97)}
  .ck-t tr.ck-gr td{background:#f3f7fc;padding:8px 14px;font-size:12.5px;border-top:1px solid #e3eaf3}
  .ck-t tr.ck-gr td b{color:#1A4E8A;font-weight:800;margin-right:10px}
  .ck-t tr.ck-gr td span{color:#7d8a9c;font-weight:600;margin-right:10px}
  .ck-t tr.ck-gr td span.ok{color:#15803d}
  .ck-dips{grid-template-columns:repeat(5,minmax(0,1fr)) !important}
  .ck-dip.all .t::before{content:'★ ';color:#F26522}
  @media(max-width:1100px){.ck-dips{grid-template-columns:repeat(2,minmax(0,1fr)) !important}}
  @media(max-width:760px){.vx .g3{grid-template-columns:1fr}}
  `;
  document.head.appendChild(st);
}

// ═══════════════════════════ MÉMOIRES ═══════════════════════════
const MEM_KW = ['Fiscalité', 'CDG', 'IA', 'Dématérialisation', 'Comptabilité', 'Audit', 'Finance', 'Paie', 'Droit social', 'RH', 'Marketing', 'RSE', 'Gestion de trésorerie', 'Consolidation', 'Digitalisation'];
const MEM_ST = { 'En cours': ['#e8eff8', '#1A4E8A'], 'Sujet validé': ['#fff4ed', '#c2410c'], 'Rendu': ['#ede9fe', '#6d28d9'], 'Soutenu': ['#e8f6ee', '#15803d'] };
const M = { data: [], etudiants: [], personnel: [], formateurs: [], motsClesRef: [], view: 'cards', sort: 'diplome', kw: '', sortCol: 'classe', sortDir: 1 };
const DIP_ORDRE = ['BTS CG', 'BTS GPME', 'DCG', 'DSCG'];
// Même règle que _ckDiplome_ (vs_extras.gs) : BTSCG1A, AIX_BTS1GPME, DCG2A, DSCG1A…
function dipOf(classe) {
  const c = String(classe || '').toUpperCase().normalize('NFD').replace(/[^A-Z]/g, '');
  if (c.includes('DSCG')) return 'DSCG';
  if (c.includes('DCG')) return 'DCG';
  if (c.includes('GPME')) return 'BTS GPME';
  if (c.includes('BTSCG') || (c.includes('BTS') && c.includes('CG'))) return 'BTS CG';
  return 'Autre';
}
const dipRank = d => { const i = DIP_ORDRE.indexOf(d); return i === -1 ? 99 : i; };
const ECH_TYPES = { 'Échange': '#1A4E8A', 'Rendez-vous': '#6d28d9', 'Mail': '#0891b2', 'Appel': '#0f766e', 'Relance': '#c2410c', 'Note': '#475569' };

async function memLoad(force) {
  const r = await api('getMemoires', {}, force ? { forceRefresh: true } : {});
  if (!r.success) { $('memBody').innerHTML = `<div class="vx-empty"><div class="ill">${svg('book', 26)}</div><h3>Chargement impossible</h3><p>${esc(errTxt(r.message))}</p></div>`; return; }
  M.data = r.data || []; M.etudiants = r.etudiants || []; M.personnel = r.personnel || []; M.formateurs = r.formateurs || []; M.motsClesRef = r.motsClesRef || MEM_KW.slice(); M.entreprises = r.entreprises || [];
  M.data.forEach(m => { m.docs = m.docs || []; m.echanges = m.echanges || []; m.diplome = dipOf(m.classe); });
  memRender();
}
function memInit() {
  injectCss();
  const sec = $('section-memoires');
  if (!sec.dataset.ready) {
    sec.dataset.ready = '1';
    sec.classList.add('vx');
    sec.innerHTML = `
      <div class="vx-hd"><div><h1>Suivi des mémoires</h1><p>Sujets, mots-clés et formateurs référents de chaque étudiant</p></div>
        <div style="display:flex;gap:8px;flex-wrap:wrap"><button class="vx-btn" onclick="VX.memKwManage()">${svg('tag', 15)} Mots-clés</button><button class="vx-btn" onclick="VX.memExport()">${svg('down', 15)} Exporter Excel</button><button class="vx-btn ora" onclick="VX.memEdit()">${svg('plus', 15, 2.4)} Ajouter un mémoire</button></div></div>
      <div class="vx-tabs"><button id="memTabM" class="on" onclick="VX.memTab('m')">${svg('book', 15)} Mémoires</button><button id="memTabC" onclick="VX.memTab('c')">${svg('check', 15, 2.4)} Checklist par diplôme</button></div>
      <div id="memCk" style="display:none"></div>
      <div id="memMain">
      <div class="vx-kpis" id="memKpis"></div>
      <div class="vx-bar">
        <div class="vx-search">${svg('search', 15)}<input id="memQ" placeholder="Étudiant, sujet, mot-clé, formateur…" oninput="VX.memRender()"></div>
        <select class="vx-sel" id="memFDip" onchange="VX.memRender()"></select>
        <select class="vx-sel" id="memFClasse" onchange="VX.memRender()"></select>
        <select class="vx-sel" id="memFCampus" onchange="VX.memRender()"></select>
        <select class="vx-sel" id="memFForm" onchange="VX.memRender()"></select>
        <select class="vx-sel" id="memSort" onchange="VX.memSort(this.value)">
          <option value="diplome">Grouper par diplôme</option><option value="classe">Grouper par classe</option><option value="campus">Grouper par campus</option><option value="formateur">Grouper par formateur</option><option value="statut">Grouper par statut</option><option value="etudiant">Trier par étudiant</option><option value="recent">Plus récents</option></select>
        <span class="vx-seg"><button id="memVC" class="on" onclick="VX.memView('cards')">${svg('grid', 14)}Cartes</button><button id="memVT" onclick="VX.memView('table')">${svg('list', 14)}Tableau</button></span>
      </div>
      <div id="memKwBar" style="display:flex;flex-wrap:wrap;gap:6px;margin:-4px 0 16px"></div>
      <div id="memBody"><div class="vx-empty"><div class="ill">${svg('book', 26)}</div><h3>Chargement…</h3></div></div>
      </div>`;
  }
  memLoad();
}
function memFiltered() {
  const q = norm($('memQ').value), fd = $('memFDip').value, fc = $('memFClasse').value, fca = $('memFCampus').value, ff = $('memFForm').value;
  return M.data.filter(m => (!fd || m.diplome === fd) && (!fc || m.classe === fc) && (!fca || m.campus === fca) && (!ff || m.formateur === ff)
    && (!M.kw || m.motsCles.some(k => norm(k) === norm(M.kw)))
    && (!q || norm([m.etudiant, m.sujet, m.formateur, m.classe, m.campus, m.entreprise, m.motsCles.join(' ')].join(' ')).includes(q)));
}
function fillSel(id, label, values) {
  const el = $(id), v = el.value;
  el.innerHTML = `<option value="">${label}</option>` + [...new Set(values.filter(Boolean))].sort((a, b) => a.localeCompare(b, 'fr')).map(x => `<option>${esc(x)}</option>`).join('');
  el.value = v;
}
function memRender() {
  { const el = $('memFDip'), v = el.value, dips = [...new Set(M.data.map(m => m.diplome))].sort((a, b) => dipRank(a) - dipRank(b));
    el.innerHTML = '<option value="">Tous les diplômes</option>' + dips.map(d => `<option>${esc(d)}</option>`).join(''); el.value = v; }
  fillSel('memFClasse', 'Toutes les classes', M.data.map(m => m.classe));
  fillSel('memFCampus', 'Tous les campus', M.data.map(m => m.campus));
  fillSel('memFForm', 'Tous les formateurs', M.data.map(m => m.formateur));
  const kwCount = {}; M.data.forEach(m => m.motsCles.forEach(k => { kwCount[k] = (kwCount[k] || 0) + 1; }));
  const topKw = Object.entries(kwCount).sort((a, b) => b[1] - a[1]);
  $('memKwBar').innerHTML = topKw.slice(0, 18).map(([k, n]) => `<span class="vx-kw ${norm(M.kw) === norm(k) ? 'on' : ''}" style="${kwColor(k)}" onclick="VX.memKw('${jsq(k)}')">${esc(k)} <small style="opacity:.7">${n}</small></span>`).join('');
  const list = memFiltered();
  $('memKpis').innerHTML = `
    <div class="vx-kpi"><span>Mémoires suivis</span><b>${M.data.length}</b><small>${list.length !== M.data.length ? list.length + ' affichés' : 'tous affichés'}</small></div>
    <div class="vx-kpi"><span>Classes</span><b>${new Set(M.data.map(m => m.classe).filter(Boolean)).size}</b><small>${new Set(M.data.map(m => m.campus).filter(Boolean)).size} campus</small></div>
    <div class="vx-kpi"><span>Formateurs référents</span><b>${new Set(M.data.map(m => m.formateur).filter(Boolean)).size}</b><small>${M.data.filter(m => !m.formateur).length} sans référent</small></div>
    <div class="vx-kpi"><span>Thème n°1</span><b style="font-size:20px">${topKw[0] ? esc(topKw[0][0]) : '—'}</b><small>${topKw[0] ? topKw[0][1] + ' mémoire(s)' : ''}</small></div>`;
  const body = $('memBody');
  if (!M.data.length) { body.innerHTML = `<div class="vx-empty"><div class="ill">${svg('book', 26)}</div><h3>Aucun mémoire suivi</h3><p>Ajoutez un premier étudiant : sujet, mots-clés et formateur référent.</p><button class="vx-btn ora" style="margin-top:14px" onclick="VX.memEdit()">${svg('plus', 15, 2.4)} Ajouter un mémoire</button></div>`; return; }
  if (!list.length) { body.innerHTML = `<div class="vx-empty"><h3>Aucun résultat</h3><p>Modifiez la recherche ou les filtres.</p></div>`; return; }
  if (M.sort === 'diplome') {
    // Diplôme → classe : un bloc par diplôme, sous-groupes par classe
    const parDip = {};
    list.forEach(m => { (parDip[m.diplome] = parDip[m.diplome] || []).push(m); });
    body.innerHTML = Object.keys(parDip).sort((a, b) => dipRank(a) - dipRank(b)).map(d => {
      const l = parDip[d], sansRef = l.filter(m => !m.formateur).length, sansSujet = l.filter(m => !m.sujet).length;
      const head = `<div class="vx-dip"><span class="t">${esc(d === 'Autre' ? 'Autres formations' : d)}</span><span class="n">${l.length} mémoire${l.length > 1 ? 's' : ''}</span>
        ${sansSujet ? `<span class="w">${sansSujet} sans sujet</span>` : ''}${sansRef ? `<span class="w">${sansRef} sans référent</span>` : ''}</div>`;
      if (M.view === 'table') return head + memTable(l);
      const parCl = {};
      l.slice().sort((a, b) => (a.etudiant || '').localeCompare(b.etudiant || '', 'fr')).forEach(m => { const c = m.classe || 'Classe non renseignée'; (parCl[c] = parCl[c] || []).push(m); });
      return head + Object.keys(parCl).sort((a, b) => a.localeCompare(b, 'fr', { numeric: true })).map(c =>
        `<div class="vx-grp">${esc(c)} <span>${parCl[c].length}</span></div><div class="vx-cards">${parCl[c].map(memCard).join('')}</div>`).join('');
    }).join('');
    return;
  }
  if (M.view === 'table') { body.innerHTML = memTable(list); return; }
  const key ={ classe: 'classe', campus: 'campus', formateur: 'formateur', statut: 'statut' }[M.sort];
  const sorted = list.slice().sort((a, b) => M.sort === 'recent' ? (b.majLe || '').localeCompare(a.majLe || '') : (a.etudiant || '').localeCompare(b.etudiant || '', 'fr'));
  if (!key) { body.innerHTML = `<div class="vx-cards">${sorted.map(memCard).join('')}</div>`; return; }
  const groups = {};
  sorted.forEach(m => { const g = m[key] || 'Non renseigné'; (groups[g] = groups[g] || []).push(m); });
  body.innerHTML = Object.keys(groups).sort((a, b) => a.localeCompare(b, 'fr', { numeric: true })).map(g =>
    `<div class="vx-grp">${esc(g)} <span>${groups[g].length}</span></div><div class="vx-cards">${groups[g].map(memCard).join('')}</div>`).join('');
}
function memCard(m) {
  const st = MEM_ST[m.statut] || MEM_ST['En cours'];
  return `<div class="vx-card" onclick="VX.memEdit('${jsq(m.id)}')">
    <div class="acts"><button class="vx-btn ico sm" title="Modifier">${svg('edit', 14)}</button></div>
    <div class="top">${avatar(m.etudiant, 42)}<div style="min-width:0"><div class="nm">${esc(m.etudiant)}</div>
      <div class="meta"><span class="vx-chip b">${esc(m.classe || '—')}</span><span class="vx-chip">${svg('pin', 11)} ${esc(m.campus || '—')}</span><span class="vx-st" style="background:${st[0]};color:${st[1]}">${esc(m.statut)}</span></div></div></div>
    ${m.entreprise ? `<div class="vx-ent">${svg('building', 13)} ${esc(m.entreprise)}</div>` : ''}
    <div class="sujet ${m.sujet ? '' : 'empty'}">${esc(m.sujet || 'Sujet à définir')}</div>
    ${m.motsCles.length ? `<div class="kws">${m.motsCles.map(k => `<span class="vx-kw" style="${kwColor(k)}" onclick="event.stopPropagation();VX.memKw('${jsq(k)}')">${esc(k)}</span>`).join('')}</div>` : ''}
    ${m.docs.length ? `<div class="vx-docchip">${svg('file', 13)} ${m.docs.length} document${m.docs.length > 1 ? 's' : ''} · dernière version <b>${esc(m.docs[m.docs.length - 1].version)}</b>${(n => n > 1 ? ' (' + n + ' fichiers)' : '')(m.docs.filter(d => String(d.version).toUpperCase() === String(m.docs[m.docs.length - 1].version).toUpperCase()).length)} du ${fmtD(m.docs[m.docs.length - 1].date)}</div>` : ''}
    ${m.echanges.length ? `<div class="vx-docchip vx-echchip">${svg('chat', 13)} ${m.echanges.length} échange${m.echanges.length > 1 ? 's' : ''} · dernier le <b>${fmtD(m.echanges[0].date)}</b><span class="tx">${esc(m.echanges[0].texte.slice(0, 70))}${m.echanges[0].texte.length > 70 ? '…' : ''}</span></div>` : ''}
    <div class="ft">${m.formateur ? avatar(m.formateur, 24) + `<span>Référent <b>${esc(m.formateur)}</b></span>` : `<span style="color:#c2410c;font-weight:600">${svg('user', 13)} Sans formateur référent</span>`}</div>
  </div>`;
}
function memTable(list) {
  const cols = [['etudiant', 'Étudiant'], ['classe', 'Classe'], ['campus', 'Campus'], ['sujet', 'Sujet'], ['motsCles', 'Mots-clés'], ['entreprise', 'Entreprise'], ['formateur', 'Formateur référent'], ['statut', 'Statut'], ['dernierDoc', 'Dernier document'], ['dernierEch', 'Carnet']];
  list.forEach(m => { const d = m.docs[m.docs.length - 1]; m.dernierDoc = d ? d.date + ' ' + d.version : ''; m.dernierEch = m.echanges.length ? m.echanges[0].date : ''; });
  const sorted = list.slice().sort((a, b) => M.sortDir * String(Array.isArray(a[M.sortCol]) ? a[M.sortCol].join(',') : a[M.sortCol] || '').localeCompare(String(Array.isArray(b[M.sortCol]) ? b[M.sortCol].join(',') : b[M.sortCol] || ''), 'fr', { numeric: true }));
  return `<table class="vx-table"><thead><tr>${cols.map(([k, l]) => `<th onclick="VX.memSortCol('${k}')">${l}${M.sortCol === k ? (M.sortDir > 0 ? ' ↑' : ' ↓') : ''}</th>`).join('')}</tr></thead><tbody>
    ${sorted.map(m => { const st = MEM_ST[m.statut] || MEM_ST['En cours']; return `<tr onclick="VX.memEdit('${jsq(m.id)}')">
      <td><div style="display:flex;align-items:center;gap:10px">${avatar(m.etudiant, 30)}<div><b style="color:var(--ink);font-weight:600">${esc(m.etudiant)}</b>${m.etuMail ? `<div style="font-size:11.5px;color:var(--mu)">${esc(m.etuMail)}</div>` : ''}</div></div></td>
      <td><span class="vx-chip b">${esc(m.classe)}</span></td><td>${esc(m.campus)}</td>
      <td style="max-width:340px">${esc(m.sujet)}</td>
      <td><div style="display:flex;flex-wrap:wrap;gap:4px">${m.motsCles.map(k => `<span class="vx-kw" style="${kwColor(k)}">${esc(k)}</span>`).join('')}</div></td>
      <td>${m.entreprise ? `<span class="vx-ent">${svg('building', 12)} ${esc(m.entreprise)}</span>` : '<span style="color:#c3ccd8">—</span>'}</td>
      <td>${esc(m.formateur)}${m.formateurMail ? `<div style="font-size:11.5px;color:var(--mu)">${esc(m.formateurMail)}</div>` : ''}</td><td><span class="vx-st" style="background:${st[0]};color:${st[1]}">${esc(m.statut)}</span></td>
      <td>${(d => d ? `<a class="vx-lastdoc" href="${esc(d.url)}" target="_blank" rel="noopener" onclick="event.stopPropagation()" title="${esc(d.nom)}"><span class="ver">${esc(d.version)}</span><span>${fmtD(d.date)}</span>${svg('ext', 12)}</a>` : '<span style="color:#c3ccd8">—</span>')(m.docs[m.docs.length - 1])}</td>
      <td>${m.echanges.length ? `<span class="vx-echn" title="${esc(m.echanges[0].texte)}">${svg('chat', 12)} ${m.echanges.length} · ${fmtD(m.echanges[0].date)}</span>` : '<span style="color:#c3ccd8">—</span>'}</td></tr>`; }).join('')}
  </tbody></table>`;
}
function memEdit(id, etuPre, focusEch) {
  const m = id ? M.data.find(x => x.id === id) : null;
  const pre = !m && etuPre ? M.etudiants.find(x => x.id === etuPre) : null;
  const E = { id: m ? m.id : '', etuId: m ? m.etuId : (pre ? pre.id : ''), motsCles: m ? m.motsCles.slice() : [], formateurRef: m ? m.formateurRef : '', formateurNom: m ? m.formateur : '', entId: m ? m.entrepriseId : '', entNom: m ? m.entreprise : '' };
  const etuLabel = e => e ? `${e.prenom} ${e.nom}` : '';
  const dlg = modal(`
    <h2>${m ? 'Modifier le mémoire' : 'Nouveau mémoire'}</h2><p class="sub">La classe et le campus sont repris automatiquement de la fiche étudiant.</p>
    <label class="fl">Étudiant</label>
    <div class="vx-ac"><input class="in" id="meEtu" placeholder="Tapez un nom ou un prénom…" autocomplete="off" value="${esc(m ? m.etudiant : (pre ? pre.prenom + ' ' + pre.nom : ''))}"><div class="vx-acl" id="meEtuL" style="display:none"></div></div>
    <div class="vx-fixed" id="meFix"></div>
    <label class="fl">Sujet du mémoire</label><textarea class="in" id="meSujet" placeholder="Ex : L'impact de l'IA sur la révision comptable dans les cabinets">${esc(m ? m.sujet : '')}</textarea>
    <label class="fl">Mots-clés</label>
    <div class="vx-tags" id="meTags" onclick="document.getElementById('meKwIn').focus()"></div>
    <div class="vx-sugg" id="meSugg"></div>
    <div class="g2">
      <div><label class="fl">Formateur référent</label><div class="vx-ac"><input class="in" id="meForm" autocomplete="off" placeholder="Rechercher un formateur…" value="${esc(m ? m.formateur : '')}"><div class="vx-acl" id="meFormL" style="display:none"></div></div>
        <button type="button" class="vx-link" id="meAutreBtn">${svg('plus', 13, 2.4)} Autre référent : préciser nom et prénom</button>
        <div class="vx-autre" id="meAutre" style="display:none"><input class="in" id="meAPrenom" placeholder="Prénom"><input class="in" id="meANom" placeholder="Nom"></div></div>
      <div><label class="fl">Statut</label><select class="in" id="meSt">${Object.keys(MEM_ST).map(s => `<option ${m && m.statut === s ? 'selected' : ''}>${s}</option>`).join('')}</select></div>
    </div>
    <label class="fl">Entreprise <span style="font-weight:500;color:var(--mu)">(facultatif — reprise de la fiche étudiant si elle est renseignée)</span></label>
    <div class="vx-ac"><input class="in" id="meEnt" autocomplete="off" placeholder="Rechercher une entreprise du CRM ou saisir un nom…" value="${esc(m ? m.entreprise : '')}"><div class="vx-acl" id="meEntL" style="display:none"></div></div>
    <label class="fl">Notes</label><textarea class="in" id="meNotes" style="min-height:60px" placeholder="Facultatif">${esc(m ? m.notes : '')}</textarea>
    <label class="fl" id="meEchLbl">${svg('chat', 13)} Carnet d'échanges${m && m.echanges.length ? ` <span class="vx-echn">${m.echanges.length}</span>` : ''}</label>
    ${m ? `<div class="vx-ech">
        <div class="vx-echf">
          <div class="r"><input type="date" class="in" id="meEchDate" value="${todayISO()}"><select class="in" id="meEchType">${Object.keys(ECH_TYPES).map(t => `<option>${t}</option>`).join('')}</select></div>
          <textarea class="in" id="meEchTxt" placeholder="Ce qui a été dit, décidé ou demandé (avec l'étudiant, le référent, le jury…)"></textarea>
          <div class="a"><span class="sub">Ctrl + Entrée pour ajouter</span><button class="vx-btn pri sm" id="meEchAdd">${svg('plus', 13, 2.4)} Ajouter au carnet</button></div>
        </div>
        <div class="vx-echl" id="meEchL"></div></div>`
      : `<div class="vx-note">${svg('chat', 14)} Enregistrez d'abord le mémoire : vous pourrez ensuite tenir son carnet d'échanges.</div>`}
    <label class="fl">Documents & versions</label>
    ${m ? `<div class="vx-docs" id="meDocs"></div>
      <div class="vx-drop" id="meDrop"><input type="file" id="meFile" multiple hidden>
        <div class="dz" id="meDz">${svg('upload', 22)}<div><b>Déposer un ou plusieurs documents</b><span>Glissez les fichiers ici ou cliquez — PDF, Word, etc. (20 Mo max chacun)</span></div></div>
        <div class="vx-upl" id="meUpl" style="display:none">
          <div class="fn" id="meFn"></div>
          <div class="vx-verpick" id="meVerPick"></div>
          <div class="g3"><div><label class="fl" style="margin-top:8px">Version</label><input class="in" id="meVer"></div><div><label class="fl" style="margin-top:8px">Date</label><input type="date" class="in" id="meDate"></div><div><label class="fl" style="margin-top:8px">Commentaire</label><input class="in" id="meCom" placeholder="Facultatif"></div></div>
          <div style="display:flex;gap:8px;justify-content:flex-end;margin-top:12px"><button class="vx-btn sm" id="meUplCancel">Annuler</button><button class="vx-btn pri sm" id="meUplGo">${svg('upload', 14)} Déposer</button></div>
        </div></div>` : `<div class="vx-note">${svg('file', 14)} Enregistrez d'abord le mémoire : vous pourrez ensuite y déposer les versions (V1, V2…) datées.</div>`}
    <div class="mft">${m ? `<button class="vx-btn danger" style="margin-right:auto" id="meDel">${svg('trash', 14)} Retirer</button>` : ''}
      <button class="vx-btn" onclick="this.closest('.vx-modal').remove()">Annuler</button>
      ${m ? '' : `<button class="vx-btn" id="meSaveNext">Enregistrer et ajouter un autre</button>`}
      <button class="vx-btn pri" id="meSave">Enregistrer</button></div>`, 860);
  const fix = () => {
    const e = M.etudiants.find(x => x.id === E.etuId);
    $('meFix').innerHTML = e ? `<span class="vx-chip b">${esc(e.classe || 'Classe ?')}</span><span class="vx-chip">${svg('pin', 11)} ${esc(e.campus || 'Campus ?')}</span>` : '';
  };
  fix();
  // auto-complétion étudiant
  const inp = $('meEtu'), L = $('meEtuL'); let items = [], act = 0;
  const draw = () => {
    const q = norm(inp.value).trim();
    if (q.length < 2) { L.style.display = 'none'; return; }
    const w = q.split(/\s+/);
    items = M.etudiants.filter(e => w.every(x => norm(e.prenom + ' ' + e.nom + ' ' + e.classe).includes(x))).slice(0, 12);
    act = 0;
    L.innerHTML = items.length ? items.map((e, i) => `<div class="it ${i === act ? 'act' : ''}" data-i="${i}">${avatar(etuLabel(e), 24)}${esc(etuLabel(e))}<small>${esc(e.classe)} · ${esc(e.campus)}</small></div>`).join('') : `<div class="it" style="color:var(--mu)">${M.etudiants.length ? 'Aucun étudiant trouvé' : 'Liste des étudiants indisponible (serveur pas à jour)'}</div>`;
    L.style.display = 'block';
  };
  const pick = e => { if (!e) return; E.etuId = e.id; inp.value = etuLabel(e); L.style.display = 'none'; fix(); entDeEtu(e); $('meSujet').focus(); };
  inp.addEventListener('input', () => { E.etuId = ''; fix(); draw(); });
  inp.addEventListener('keydown', ev => {
    if (L.style.display === 'none') return;
    if (ev.key === 'ArrowDown' || ev.key === 'ArrowUp') { ev.preventDefault(); act = Math.max(0, Math.min(items.length - 1, act + (ev.key === 'ArrowDown' ? 1 : -1))); L.querySelectorAll('[data-i]').forEach((el, i) => el.classList.toggle('act', i === act)); }
    if (ev.key === 'Enter') { ev.preventDefault(); pick(items[act]); }
  });
  L.addEventListener('mousedown', ev => { const it = ev.target.closest('[data-i]'); if (it) { ev.preventDefault(); pick(items[+it.dataset.i]); } });
  inp.addEventListener('blur', () => setTimeout(() => L.style.display = 'none', 150));
  // mots-clés
  const kwPool = () => { const seen = new Set(); return M.motsClesRef.concat(...M.data.map(x => x.motsCles)).filter(k => { const n = norm(k); if (!n || seen.has(n)) return false; seen.add(n); return true; }); };
  const drawTags = () => {
    $('meTags').innerHTML = E.motsCles.map((k, i) => `<span class="vx-kw" style="${kwColor(k)}">${esc(k)}<button data-rm="${i}">${svg('x', 11, 2.6)}</button></span>`).join('') + `<input id="meKwIn" placeholder="${E.motsCles.length ? 'Ajouter…' : 'Tapez un mot-clé (existant ou nouveau) puis Entrée'}">`;
    drawSugg('');
    const ki = $('meKwIn');
    ki.addEventListener('input', () => drawSugg(ki.value));
    ki.addEventListener('keydown', ev => {
      if ((ev.key === 'Enter' || ev.key === ',') && ki.value.trim()) { ev.preventDefault(); addKw(ki.value.trim()); }
      if (ev.key === 'Backspace' && !ki.value && E.motsCles.length) { E.motsCles.pop(); drawTags(); $('meKwIn').focus(); }
    });
  };
  // suggestions : mots-clés existants qui correspondent à la saisie + « Créer » si nouveau
  const drawSugg = q => {
    const nq = norm(q).trim();
    const dispo = kwPool().filter(k => !E.motsCles.some(x => norm(x) === norm(k)) && (!nq || norm(k).includes(nq)));
    const exact = nq && kwPool().some(k => norm(k) === nq);
    $('meSugg').innerHTML = (nq && !exact ? `<span class="vx-kw vx-kw-new" data-create="${esc(q.trim())}">${svg('plus', 11, 2.6)} Créer le mot-clé « ${esc(q.trim())} »</span>` : '')
      + dispo.slice(0, nq ? 12 : 18).map(k => `<span class="vx-kw" style="${kwColor(k)}" data-add="${esc(k)}">+ ${esc(k)}</span>`).join('');
  };
  const addKw = k => {
    const hit = kwPool().find(x => norm(x) === norm(k));
    if (!hit) return createKw(k);
    if (!E.motsCles.some(x => norm(x) === norm(hit))) E.motsCles.push(hit);
    drawTags(); $('meKwIn').focus();
  };
  // nouveau mot-clé : ajouté au mémoire ET enregistré dans la liste partagée
  const createKw = k => {
    k = k.trim().replace(/\s+/g, ' ').slice(0, 40); if (!k) return;
    k = k.charAt(0).toUpperCase() + k.slice(1);
    M.motsClesRef.push(k);
    if (!E.motsCles.some(x => norm(x) === norm(k))) E.motsCles.push(k);
    drawTags(); $('meKwIn').focus();
    api('saveMotsClesMemoire', { ajouter: [k] }).then(r => { if (r.success) { M.motsClesRef = r.motsCles; toast('Mot-clé « ' + k + ' » ajouté à la liste ✓'); } else toast('Erreur : ' + errTxt(r.message), false); });
  };
  dlg.addEventListener('click', ev => {
    const rm = ev.target.closest('[data-rm]'); if (rm) { E.motsCles.splice(+rm.dataset.rm, 1); drawTags(); return; }
    const ad = ev.target.closest('[data-add]'); if (ad) addKw(ad.dataset.add);
    const cr = ev.target.closest('[data-create]'); if (cr) createKw(cr.dataset.create);
  });
  drawTags();
  // formateur : combobox (formateurs actifs + équipe) avec saisie libre
  const fIn = $('meForm'), fL = $('meFormL');
  const fPool = () => M.formateurs.map(x => ({ ...x, t: 'Formateur', s: x.matieres })).concat(M.personnel.map(x => ({ ...x, t: 'Équipe', s: x.role })));
  let fItems = [];
  const fDraw = () => {
    const q = norm(fIn.value).trim();
    fItems = fPool().filter(x => !q || norm(x.nom).includes(q)).slice(0, 40);
    fL.innerHTML = (fItems.length ? fItems.map((x, i) => `<div class="it" data-fi="${i}">${avatar(x.nom, 24)}${esc(x.nom)}<small>${esc(x.t)}${x.s ? ' · ' + esc(String(x.s).slice(0, 40)) : ''}</small></div>`).join('') : '')
      + (fIn.value.trim() && !fItems.some(x => norm(x.nom) === q) ? `<div class="it" data-free="1" style="color:var(--b)">${svg('plus', 13, 2.4)} Utiliser « ${esc(fIn.value.trim())} »</div>` : '')
      + `<div class="it" data-autre="1" style="color:var(--o)">${svg('user', 13)} Autre… préciser nom et prénom</div>`
      + (!fPool().length ? `<div class="it" style="color:var(--mu)">Liste des formateurs indisponible — saisissez le nom</div>` : '');
    fL.style.display = 'block';
  };
  fIn.addEventListener('focus', fDraw);
  fIn.addEventListener('input', () => { E.formateurRef = ''; E.formateurNom = fIn.value.trim(); fDraw(); });
  fL.addEventListener('mousedown', ev => { ev.preventDefault();
    const it = ev.target.closest('.it'); if (!it) return;
    if (it.dataset.autre) { fL.style.display = 'none'; openAutre(); return; }
    if (it.dataset.fi !== undefined) { const x = fItems[+it.dataset.fi]; E.formateurRef = x.ref; E.formateurNom = x.nom; fIn.value = x.nom; }
    else { E.formateurRef = ''; E.formateurNom = fIn.value.trim(); }
    fL.style.display = 'none'; });
  fIn.addEventListener('blur', () => setTimeout(() => fL.style.display = 'none', 150));
  // « Autre » : deux champs prénom / nom → « Prénom NOM »
  const openAutre = () => {
    const box = $('meAutre'); box.style.display = 'grid'; $('meAutreBtn').style.display = 'none';
    E.formateurRef = '';
    const parts = (E.formateurNom || '').trim().split(/\s+/);
    if (parts.length > 1 && !$('meAPrenom').value && !$('meANom').value) { $('meAPrenom').value = parts[0]; $('meANom').value = parts.slice(1).join(' '); }
    $('meAPrenom').focus();
  };
  const syncAutre = () => {
    const p = $('meAPrenom').value.trim(), n = $('meANom').value.trim().toUpperCase();
    fIn.value = (p + ' ' + n).trim(); E.formateurRef = ''; E.formateurNom = fIn.value;
  };
  $('meAutreBtn').onclick = openAutre;
  $('meAPrenom').addEventListener('input', syncAutre);
  $('meANom').addEventListener('input', syncAutre);
  // documents
  if (m) {
    // Documents regroupés par version (une version peut contenir plusieurs fichiers)
    const versions = () => { const g = {}; m.docs.forEach(d => { const v = String(d.version || '—').toUpperCase(); (g[v] = g[v] || []).push(d); });
      return Object.keys(g).sort((a, b) => (parseInt(b.replace(/\D/g, ''), 10) || 0) - (parseInt(a.replace(/\D/g, ''), 10) || 0)).map(v => [v, g[v]]); };
    const drawDocs = () => {
      $('meDocs').innerHTML = m.docs.length ? versions().map(([v, l]) => `<div class="vx-vgrp"><div class="vh"><span class="ver">${esc(v)}</span><b>${l.length} fichier${l.length > 1 ? 's' : ''}</b><small>${fmtLong(l[l.length - 1].date)}</small></div>
        ${l.map(d => `<div class="vx-doc">${svg('file', 16)}
        <div class="in2"><a href="${esc(d.url)}" target="_blank" rel="noopener">${esc(d.nom)} ${svg('ext', 12)}</a><small>${fmtLong(d.date)}${d.par ? ' · déposé par ' + esc(d.par) : ''}${d.commentaire ? ' · ' + esc(d.commentaire) : ''}</small></div>
        <button class="vx-btn ico sm danger" data-docdel="${esc(d.id)}" title="Retirer de la liste (le fichier reste dans Drive)">${svg('trash', 13)}</button></div>`).join('')}</div>`).join('')
        : `<div class="vx-note">${svg('file', 14)} Aucun document pour l'instant.</div>`;
    };
    drawDocs();
    let files = [];
    const nextVer = () => 'V' + (m.docs.reduce((mx, d) => Math.max(mx, parseInt(String(d.version).replace(/\D/g, ''), 10) || 0), 0) + 1);
    const drawFiles = () => {
      $('meFn').innerHTML = files.map((f, i) => `<div class="vx-fsel">${svg('file', 15)} <b>${esc(f.name)}</b> <span>${(f.size / 1024 / 1024).toFixed(2)} Mo</span><button type="button" class="vx-btn ico sm" data-frm="${i}" title="Retirer">${svg('x', 12, 2.6)}</button></div>`).join('')
        + `<button type="button" class="vx-link" id="meMore">${svg('plus', 13, 2.4)} Ajouter d'autres fichiers à cet envoi</button>`;
      $('meMore').onclick = () => $('meFile').click();
      $('meUplGo').innerHTML = svg('upload', 14) + (files.length > 1 ? ' Déposer les ' + files.length + ' fichiers' : ' Déposer');
    };
    const drawVerPick = () => {
      const vs = versions().map(([v]) => v), nv = nextVer(), cur = $('meVer').value.trim().toUpperCase();
      $('meVerPick').innerHTML = `<span>Ranger dans :</span><button type="button" class="${cur === nv ? 'on' : ''}" data-vp="${nv}">Nouvelle version ${nv}</button>`
        + vs.map(v => `<button type="button" class="${cur === v ? 'on' : ''}" data-vp="${esc(v)}">Ajouter à ${esc(v)}</button>`).join('');
    };
    const choose = list => {
      const ok = [...(list || [])].filter(f => { if (f.size > 20 * 1024 * 1024) { toast(f.name + ' : trop volumineux (20 Mo max)', false); return false; } return true; });
      if (!ok.length) return;
      const premier = !files.length;
      files = files.concat(ok);
      if (premier) { $('meVer').value = nextVer(); $('meDate').value = todayISO(); $('meCom').value = ''; }
      drawFiles(); drawVerPick();
      $('meDz').style.display = 'none'; $('meUpl').style.display = 'block';
    };
    $('meVerPick').addEventListener('click', ev => { const b = ev.target.closest('[data-vp]'); if (!b) return; $('meVer').value = b.dataset.vp; drawVerPick(); });
    $('meVer').addEventListener('input', drawVerPick);
    $('meFn').addEventListener('click', ev => { const b = ev.target.closest('[data-frm]'); if (!b) return; files.splice(+b.dataset.frm, 1); if (!files.length) $('meUplCancel').click(); else drawFiles(); });
    const dz = $('meDz');
    dz.onclick = () => $('meFile').click();
    $('meFile').onchange = e => { choose(e.target.files); e.target.value = ''; };
    dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('over'); });
    dz.addEventListener('dragleave', () => dz.classList.remove('over'));
    dz.addEventListener('drop', e => { e.preventDefault(); dz.classList.remove('over'); choose(e.dataTransfer.files); });
    $('meUplCancel').onclick = () => { files = []; $('meFile').value = ''; $('meUpl').style.display = 'none'; $('meDz').style.display = 'flex'; };
    $('meUplGo').onclick = async () => {
      if (!files.length) return;
      const btn = $('meUplGo'); btn.disabled = true;
      const version = $('meVer').value.trim().toUpperCase() || nextVer(), date = $('meDate').value, com = $('meCom').value.trim();
      const restants = [];
      let ok = 0;
      for (let i = 0; i < files.length; i++) {
        const f = files[i];
        btn.innerHTML = `Envoi ${i + 1}/${files.length}…`;
        try {
          const b64 = await new Promise((res, rej) => { const fr = new FileReader(); fr.onload = () => res(String(fr.result).split(',')[1] || ''); fr.onerror = rej; fr.readAsDataURL(f); });
          const r = await api('uploadMemoireDoc', { memoireId: m.id, data: b64, fileName: f.name, mimeType: f.type || 'application/octet-stream', version, date, commentaire: com });
          if (!r.success) { restants.push(f); toast(f.name + ' : ' + errTxt(r.message), false); continue; }
          m.docs.push({ ...r.doc, par: user.nom || '' }); ok++;
        } catch (e) { restants.push(f); toast(f.name + ' : envoi impossible (' + e + ')', false); }
      }
      btn.disabled = false;
      drawDocs(); memRender();
      if (restants.length) { files = restants; drawFiles(); drawVerPick(); }
      else $('meUplCancel').click();
      if (ok) toast(ok + ' fichier' + (ok > 1 ? 's' : '') + ' déposé' + (ok > 1 ? 's' : '') + ' dans ' + version + ' ✓');
    };
    dlg.addEventListener('click', async ev => {
      const del = ev.target.closest('[data-docdel]'); if (!del) return;
      if (!confirm('Retirer ce document de la liste ? Le fichier reste dans le dossier Drive de l\'étudiant.')) return;
      const r = await api('archiveMemoireDoc', { id: del.dataset.docdel });
      if (!r.success) { toast('Erreur : ' + errTxt(r.message), false); return; }
      m.docs = m.docs.filter(d => d.id !== del.dataset.docdel); drawDocs(); memRender();
    });
  }
  // entreprise : auto-complétion sur les entreprises du CRM, saisie libre acceptée
  const entIn = $('meEnt'), entL = $('meEntL'); let eItems = [], eAct = 0;
  const entDraw = () => {
    const q = norm(entIn.value).trim();
    if (q.length < 2) { entL.style.display = 'none'; return; }
    const w = q.split(/\s+/);
    eItems = (M.entreprises || []).filter(x => w.every(t => norm(x.nom).includes(t))).slice(0, 12); eAct = 0;
    entL.innerHTML = (eItems.length ? eItems.map((x, i) => `<div class="it ${i === eAct ? 'act' : ''}" data-ei="${i}">${svg('building', 14)} ${esc(x.nom)}<small>${esc(x.id)}</small></div>`).join('') : '')
      + `<div class="it" data-elibre="1" style="color:var(--mu)">${svg('plus', 13, 2.4)} Utiliser « ${esc(entIn.value.trim())} » (hors base entreprises)</div>`;
    entL.style.display = 'block';
  };
  const entPick = x => { if (x) { E.entId = x.id; E.entNom = x.nom; entIn.value = x.nom; } else { E.entId = ''; E.entNom = entIn.value.trim(); } entL.style.display = 'none'; };
  const entDeEtu = e => {
    if (!e || entIn.value.trim() || !e.entId) return;
    const x = (M.entreprises || []).find(k => k.id.toUpperCase() === String(e.entId).toUpperCase());
    if (x) { entPick(x); toast('Entreprise reprise de la fiche étudiant : ' + x.nom); }
  };
  if (!m && pre) entDeEtu(pre);
  entIn.addEventListener('input', () => { E.entId = ''; E.entNom = entIn.value.trim(); entDraw(); });
  entIn.addEventListener('keydown', ev => {
    if (entL.style.display === 'none') return;
    if (ev.key === 'ArrowDown' || ev.key === 'ArrowUp') { ev.preventDefault(); eAct = Math.max(0, Math.min(eItems.length - 1, eAct + (ev.key === 'ArrowDown' ? 1 : -1))); entL.querySelectorAll('[data-ei]').forEach((el, i) => el.classList.toggle('act', i === eAct)); }
    if (ev.key === 'Enter') { ev.preventDefault(); entPick(eItems[eAct]); }
  });
  entL.addEventListener('mousedown', ev => { const it = ev.target.closest('[data-ei],[data-elibre]'); if (!it) return; ev.preventDefault(); entPick(it.dataset.ei !== undefined ? eItems[+it.dataset.ei] : null); });
  entIn.addEventListener('blur', () => setTimeout(() => entL.style.display = 'none', 150));
  // carnet d'échanges
  if (m) {
    const drawEch = () => {
      $('meEchL').innerHTML = m.echanges.length ? m.echanges.map(x => `<div class="vx-echi" style="--c:${ECH_TYPES[x.type] || '#1A4E8A'}">
        <div class="h"><span class="ty">${esc(x.type)}</span><b>${fmtLong(x.date)}</b><span class="by">${esc(x.auteur)}</span>
          <button class="vx-btn ico sm danger" data-echdel="${esc(x.id)}" title="Retirer cette entrée">${svg('trash', 12)}</button></div>
        <div class="tx">${esc(x.texte).replace(/\n/g, '<br>')}</div></div>`).join('')
        : `<div class="vx-note">${svg('chat', 14)} Aucun échange noté pour l'instant.</div>`;
      const lb = $('meEchLbl'); if (lb) lb.innerHTML = `${svg('chat', 13)} Carnet d'échanges${m.echanges.length ? ` <span class="vx-echn">${m.echanges.length}</span>` : ''}`;
    };
    drawEch();
    const syncCk = () => {
      memRender();
      if (C.data) Object.values(C.data.diplomes).flat().forEach(e => { if (e.memoireId === m.id) { e.nbEchanges = m.echanges.length; e.dernierEchange = m.echanges[0] ? m.echanges[0].date : ''; } });
      if (C.open && C.data) ckRender();
    };
    const addEch = async () => {
      const txt = $('meEchTxt').value.trim(); if (!txt) { $('meEchTxt').focus(); return; }
      const btn = $('meEchAdd'); btn.disabled = true;
      try {
        const r = await api('saveEchangeMemoire', { memoireId: m.id, echange: { date: $('meEchDate').value, type: $('meEchType').value, texte: txt } });
        if (!r.success) { toast('Erreur : ' + errTxt(r.message), false); return; }
        m.echanges.unshift(r.echange);
        m.echanges.sort((a, b) => (b.date + b.creeLe).localeCompare(a.date + a.creeLe));
        $('meEchTxt').value = ''; drawEch(); syncCk(); toast('Ajouté au carnet ✓');
      } finally { btn.disabled = false; }
    };
    $('meEchAdd').onclick = addEch;
    $('meEchTxt').addEventListener('keydown', ev => { if (ev.key === 'Enter' && (ev.ctrlKey || ev.metaKey)) { ev.preventDefault(); addEch(); } });
    dlg.addEventListener('click', async ev => {
      const del = ev.target.closest('[data-echdel]'); if (!del) return;
      if (!confirm('Retirer cette entrée du carnet ? (elle reste archivée dans la feuille MEMOIRES_ECHANGES)')) return;
      const r = await api('saveEchangeMemoire', { retirer: del.dataset.echdel });
      if (!r.success) { toast('Erreur : ' + errTxt(r.message), false); return; }
      m.echanges = m.echanges.filter(x => x.id !== del.dataset.echdel); drawEch(); syncCk();
    });
    if (focusEch) setTimeout(() => { $('meEchLbl').scrollIntoView({ behavior: 'smooth', block: 'start' }); $('meEchTxt').focus({ preventScroll: true }); }, 60);
  }
  const save = async next => {
    const k = $('meKwIn').value.trim(); if (k) addKw(k);
    if (!E.etuId) { toast(M.etudiants.length ? 'Choisissez un étudiant dans la liste' : "Liste des étudiants indisponible : " + errTxt('Action inconnue'), false); inp.focus(); return; }
    const fNom = $('meForm').value.trim();
    const payload = { id: E.id, etuId: E.etuId, sujet: $('meSujet').value.trim(), motsCles: E.motsCles, formateurRef: fNom && fNom === E.formateurNom ? E.formateurRef : '', formateur: fNom, statut: $('meSt').value, notes: $('meNotes').value.trim(), entreprise: $('meEnt').value.trim(), entrepriseId: $('meEnt').value.trim() && $('meEnt').value.trim() === E.entNom ? E.entId : '' };
    const r = await api('saveMemoire', { memoire: payload });
    if (!r.success) { toast('Erreur : ' + errTxt(r.message), false); return; }
    toast(m ? 'Mémoire mis à jour ✓' : 'Mémoire ajouté ✓');
    dlg.remove(); await memLoad(true); if (C.open) ckLoad();
    if (next) memEdit(); else if (!m && r.id) { memEdit(r.id); toast('Vous pouvez maintenant déposer les documents du mémoire'); }
  };
  $('meSave').onclick = () => save(false);
  if ($('meSaveNext')) $('meSaveNext').onclick = () => save(true);
  if ($('meDel')) $('meDel').onclick = async () => {
    if (!confirm('Retirer ce mémoire du suivi ? (il reste archivé dans la feuille MEMOIRES)')) return;
    const r = await api('saveMemoire', { memoire: { id: m.id, etuId: m.etuId, sujet: m.sujet, motsCles: m.motsCles, formateurRef: m.formateurRef, statut: m.statut, notes: m.notes, archiver: true } });
    if (!r.success) { toast('Erreur : ' + errTxt(r.message), false); return; }
    dlg.remove(); toast('Mémoire retiré'); memLoad(true);
  };
  if (!m) inp.focus();
}
// ═══════════════════════════ CHECKLIST PAR DIPLÔME ═══════════════════════════
const CK_DIP = ['BTS CG', 'BTS GPME', 'DCG', 'DSCG'];
const C = { open: false, annee: '', data: null, dip: 'ALL', loading: false, sel: new Set(), q: '', cls: '' };
// Mail affiché : mail HECG en priorité, sinon mail perso (grisé, signalé)
const ckMailEtu = e => e.mailHecg ? `<a class="ck-mail" href="mailto:${esc(e.mailHecg)}" onclick="event.stopPropagation()">${esc(e.mailHecg)}</a>`
  : (e.mailPerso ? `<a class="ck-mail perso" href="mailto:${esc(e.mailPerso)}" title="Pas de mail HECG — mail personnel" onclick="event.stopPropagation()">${esc(e.mailPerso)} (perso)</a>` : '');
function memTab(t) {
  C.open = t === 'c';
  $('memTabM').classList.toggle('on', !C.open); $('memTabC').classList.toggle('on', C.open);
  $('memMain').style.display = C.open ? 'none' : ''; $('memCk').style.display = C.open ? '' : 'none';
  if (C.open && !C.data) ckLoad();
}
async function ckLoad(annee) {
  C.loading = true; if (annee) C.annee = annee;
  if (!C.data) $('memCk').innerHTML = `<div class="vx-empty"><div class="ill">${svg('check', 26, 2.4)}</div><h3>Chargement de la checklist…</h3></div>`;
  const r = await api('getMemoiresChecklist', C.annee ? { annee: C.annee } : {}, { forceRefresh: true });
  C.loading = false;
  if (!r.success) { $('memCk').innerHTML = `<div class="vx-empty"><h3>Chargement impossible</h3><p>${esc(errTxt(r.message))}</p></div>`; return; }
  C.data = r; C.annee = r.annee;
  ckRender();
}
const ckScore = e => [!!e.sujet.trim(), e.nbDocs > 0, !!e.referent.trim(), e.cyclade];
// Regroupe une liste d'étudiants par classe (ordre naturel : BTSCG1A < BTSCG2A…)
const ckParClasse = l => { const g = {}; l.forEach(e => { const c = e.classe || 'Classe non renseignée'; (g[c] = g[c] || []).push(e); }); return Object.keys(g).sort((a, b) => a.localeCompare(b, 'fr', { numeric: true })).map(c => [c, g[c].sort((a, b) => a.etudiant.localeCompare(b.etudiant, 'fr'))]); };
function ckRender() {
  const r = C.data, ro = r.archivee;
  const stats = d => { const l = (r.diplomes[d] || []).filter(e => e.concerne); return { n: l.length, ok: l.filter(e => ckScore(e).every(Boolean)).length, pts: l.reduce((s, e) => s + ckScore(e).filter(Boolean).length, 0) }; };
  const all = C.dip === 'ALL';
  const list = all ? [] : (r.diplomes[C.dip] || []);
  const concernes = list.filter(e => e.concerne), autres = list.filter(e => !e.concerne && !e.sorti);
  const yes = (ok, txt, title) => `<span class="ck-ok ${ok ? 'y' : 'n'}" title="${esc(title || '')}">${ok ? svg('check', 13, 3) : svg('x', 12, 3)}${txt ? `<span>${txt}</span>` : ''}</span>`;
  const row = e => {
    const sc = ckScore(e), n = sc.filter(Boolean).length;
    const d = e.dernierDoc;
    return `<tr class="${n === 4 ? 'full' : ''}">
      <td style="width:44px">${ro ? '' : `<label class="ck-box"><input type="checkbox" checked onchange="VX.ckConcerne(['${jsq(e.etuId)}'],false)"><span>${svg('check', 12, 3)}</span></label>`}</td>
      <td><div class="ck-etu" ${e.memoireId || !ro ? `onclick="VX.memEdit(${e.memoireId ? `'${jsq(e.memoireId)}'` : `null,'${jsq(e.etuId)}'`})"` : ''}>${avatar(e.etudiant, 32)}<div><b>${esc(e.etudiant)}${e.rapprochement === 'nom' ? '<span class="ck-rap" title="Mémoire retrouvé par le nom de l\'étudiant (ID différent) — vérifiez la fiche">rapproché par le nom</span>' : ''}</b><small>${esc(e.classe)} · ${esc(e.campus)}${e.sorti ? ' · <span style="color:#c62828">sorti</span>' : ''}</small>${ckMailEtu(e)}${e.memoireId && !ro ? `<button class="ck-ech ${e.nbEchanges ? 'on' : ''}" onclick="event.stopPropagation();VX.memEdit('${jsq(e.memoireId)}',null,true)" title="Ouvrir le carnet d'échanges">${svg('chat', 12)} ${e.nbEchanges ? e.nbEchanges + ' échange' + (e.nbEchanges > 1 ? 's' : '') + ' · ' + fmtD(e.dernierEchange) : 'Carnet vide'}</button>` : (e.nbEchanges ? `<span class="ck-ech on">${svg('chat', 12)} ${e.nbEchanges} échange${e.nbEchanges > 1 ? 's' : ''}</span>` : '')}</div></div></td>
      <td>${yes(sc[0], '', e.sujet || 'Sujet non renseigné')}${e.sujet ? `<span class="ck-sub" title="${esc(e.sujet)}">${esc(e.sujet.slice(0, 60))}${e.sujet.length > 60 ? '…' : ''}</span>` : `<span class="ck-sub" style="color:#c62828">${e.memoireId ? 'Sujet à saisir' : 'Aucun mémoire saisi'}</span>`}</td>
      <td>${yes(sc[1], '')}${d ? `<a class="vx-lastdoc" href="${esc(d.url)}" target="_blank" rel="noopener"><span class="ver">${esc(d.version)}</span><span>${fmtD(d.date)}</span>${svg('ext', 12)}</a>` : ''}</td>
      <td>${yes(sc[2], '')}${e.referent ? `<span class="ck-sub">${esc(e.referent)}</span>${e.referentMail ? `<a class="ck-mail" href="mailto:${esc(e.referentMail)}">${esc(e.referentMail)}</a>` : ''}` : ''}</td>
      <td>${ro ? yes(sc[3], e.cycladeDate ? fmtD(e.cycladeDate) : '') : `<label class="ck-cyc ${e.cyclade ? 'on' : ''}"><input type="checkbox" ${e.cyclade ? 'checked' : ''} onchange="VX.ckCyclade('${jsq(e.etuId)}',this.checked)"><span class="sw"></span>${e.cyclade ? 'Déposé' + (e.cycladeDate ? ' le ' + fmtD(e.cycladeDate) : '') : 'À déposer'}</label>`}</td>
      <td><span class="ck-prog p${n}"><i style="width:${n * 25}%"></i></span><b class="ck-n ${n === 4 ? 'ok' : ''}">${n === 4 ? 'Complet' : n + '/4'}</b></td></tr>`;
  };
  $('memCk').innerHTML = `
    <div class="ck-top">
      <div class="ck-year"><label>Année scolaire</label><select class="vx-sel" onchange="VX.ckYear(this.value)">${r.annees.map(a => `<option ${a === r.annee ? 'selected' : ''}>${a}</option>`).join('')}<option value="__new">Autre année…</option></select></div>
      ${ro ? `<span class="ck-arch">${svg('book', 14)} Année archivée — consultation seule</span>` : `<span class="ck-live">${svg('clock', 14)} Mise à jour automatique avec les étudiants inscrits</span>`}
      <span style="flex:1"></span>
      <button class="vx-btn" onclick="VX.ckExport()">${svg('down', 15)} Exporter Excel</button>
    </div>
    <div class="ck-dips">${[{ d: 'ALL', l: 'Tous les diplômes' }].concat(CK_DIP.map(d => ({ d, l: d }))).map(({ d, l }) => { const s2 = d === 'ALL' ? CK_DIP.map(stats).reduce((a, b) => ({ n: a.n + b.n, ok: a.ok + b.ok, pts: a.pts + b.pts }), { n: 0, ok: 0, pts: 0 }) : stats(d); const pct = s2.n ? Math.round(s2.pts / (s2.n * 4) * 100) : 0;
      const nAutres = (d === 'ALL' ? CK_DIP.flatMap(x => r.diplomes[x] || []) : (r.diplomes[d] || [])).filter(e => !e.concerne && !e.sorti).length;
      return `<button class="ck-dip ${d === 'ALL' ? 'all ' : ''}${C.dip === d ? 'on' : ''}" onclick="VX.ckDip('${d}')">
      <span class="t">${l}</span><span class="v"><b>${s2.ok}</b>/${s2.n} complet${s2.ok > 1 ? 's' : ''}</span>
      <span class="bar"><i style="width:${pct}%"></i></span><span class="p">${pct} % des contrôles validés${d === 'ALL' ? '' : ' · ' + nAutres + ' autre(s) inscrit(s)'}</span></button>`; }).join('')}</div>
    ${(all ? CK_DIP : [C.dip]).map(dip => { const cc = (r.diplomes[dip] || []).filter(e => e.concerne); if (all && !cc.length) return '';
      return `<div class="ck-card">
      <div class="ck-h"><h3>${dip} · étudiants qui passent le mémoire en ${esc(r.annee)} <span>${cc.length}</span></h3>
        ${all ? `<button class="vx-btn sm" onclick="VX.ckDip('${dip}')">Ouvrir ${dip}</button>` : '<div class="ck-legend"><span>Sujet</span><span>Mémoire inséré</span><span>Prof référent</span><span>Dépôt Cyclade</span></div>'}</div>
      ${cc.length ? `<table class="ck-t"><thead><tr><th></th><th>Étudiant</th><th>Sujet</th><th>Mémoire inséré</th><th>Prof référent</th><th>Dépôt Cyclade</th><th>Avancement</th></tr></thead><tbody>${ckParClasse(cc).map(([cl, l]) => { const ok = l.filter(e => ckScore(e).every(Boolean)).length;
          return `<tr class="ck-gr"><td colspan="7"><b>${esc(cl)}</b><span>${l.length} étudiant${l.length > 1 ? 's' : ''}</span><span class="${ok === l.length ? 'ok' : ''}">${ok}/${l.length} complet${ok > 1 ? 's' : ''}</span></td></tr>` + l.map(row).join(''); }).join('')}</tbody></table>`
        : `<div class="vx-note" style="margin:6px 0">${svg('check', 14, 2.4)} ${ro ? 'Aucun étudiant archivé pour ce diplôme.' : 'Cochez ci-dessous les étudiants qui passent le mémoire cette année.'}</div>`}
    </div>`; }).join('') || `<div class="ck-card"><div class="vx-note" style="margin:6px 0">${svg('check', 14, 2.4)} ${ro ? 'Aucun étudiant archivé pour cette année.' : 'Aucun étudiant dans la checklist : ouvrez un diplôme pour cocher ceux qui passent le mémoire.'}</div></div>`}
    ${ro ? '' : (all ? `<div class="vx-note" style="margin-top:4px">${svg('plus', 14, 2.4)} Pour ajouter des étudiants à la checklist, ouvrez un diplôme ci-dessus.</div>` : ckOthersHtml(autres))}`;
}
// Sélection des étudiants qui passent le mémoire : on coche (recherche + filtre classe), puis
// « Ajouter la sélection » — le rapprochement avec les mémoires déjà saisis est fait par le serveur.
function ckOthersHtml(autres) {
  const classes = [...new Set(autres.map(e => e.classe).filter(Boolean))].sort();
  const q = norm(C.q);
  const vis = autres.filter(e => (!C.cls || e.classe === C.cls) && (!q || norm(e.etudiant + ' ' + e.classe).includes(q)));
  const nSel = autres.filter(e => C.sel.has(e.etuId)).length;
  return `<div class="ck-card ck-others">
      <div class="ck-h"><h3>Autres inscrits en ${C.dip} <span>${autres.length}</span></h3>
        ${autres.length ? `<div style="display:flex;gap:8px;flex-wrap:wrap">
          <button class="vx-btn sm" onclick="VX.ckSelAll(true)">${svg('check', 13, 2.4)} Tout sélectionner${vis.length !== autres.length ? ' (filtrés)' : ''}</button>
          ${nSel ? `<button class="vx-btn sm" onclick="VX.ckSelAll(false)">Désélectionner</button>` : ''}
          <button class="vx-btn sm ora" ${nSel ? '' : 'disabled style="opacity:.5;cursor:not-allowed"'} onclick="VX.ckAddSel()">${svg('plus', 13, 2.4)} Ajouter la sélection${nSel ? ' (' + nSel + ')' : ''}</button></div>` : ''}</div>
      ${autres.length ? `<div class="ck-filt">
          <div class="vx-search">${svg('search', 15)}<input id="ckQ" placeholder="Rechercher un étudiant…" value="${esc(C.q)}" oninput="VX.ckFilter(this.value,null)"></div>
          <select class="vx-sel" onchange="VX.ckFilter(null,this.value)"><option value="">Toutes les classes</option>${classes.map(c => `<option ${c === C.cls ? 'selected' : ''}>${esc(c)}</option>`).join('')}</select>
        </div>
        <div class="ck-grid">${vis.map(e => `<label class="ck-pick"><input type="checkbox" ${C.sel.has(e.etuId) ? 'checked' : ''} onchange="VX.ckSel('${jsq(e.etuId)}',this.checked)"><span class="b">${svg('check', 12, 3)}</span>${avatar(e.etudiant, 28)}<span class="n"><b>${esc(e.etudiant)}</b><small>${esc(e.classe)} · ${esc(e.campus)}</small>
          <span class="m ${e.memoireId ? 'y' : 'n'}">${e.memoireId ? '✓ Mémoire déjà saisi' + (e.sujet ? '' : ' (sans sujet)') : 'Pas encore de mémoire'}</span></span></label>`).join('') || '<p class="sub">Aucun étudiant ne correspond à la recherche.</p>'}</div>`
        : `<p class="sub" style="margin:4px 0 0">Tous les inscrits de ce diplôme sont dans la checklist.</p>`}
    </div>`;
}
function ckSel(id, on) { if (on) C.sel.add(id); else C.sel.delete(id); ckRender(); }
function ckSelAll(on) {
  const autres = (C.data.diplomes[C.dip] || []).filter(e => !e.concerne && !e.sorti);
  const q = norm(C.q);
  autres.filter(e => (!C.cls || e.classe === C.cls) && (!q || norm(e.etudiant + ' ' + e.classe).includes(q)))
    .forEach(e => on ? C.sel.add(e.etuId) : C.sel.delete(e.etuId));
  if (!on) C.sel.clear();
  ckRender();
}
function ckFilter(q, cls) {
  if (q !== null) C.q = q; if (cls !== null) C.cls = cls;
  ckRender();
  if (q !== null) { const i = $('ckQ'); if (i) { i.focus(); i.setSelectionRange(i.value.length, i.value.length); } }
}
async function ckAddSel() {
  const autres = (C.data.diplomes[C.dip] || []).filter(e => !e.concerne && !e.sorti && C.sel.has(e.etuId));
  if (!autres.length) return;
  const ids = autres.map(e => e.etuId);
  C.sel.clear();
  await ckConcerne(ids, true);
  // Bilan du rapprochement avec les mémoires déjà en place
  const avec = autres.filter(e => e.memoireId).length, sujet = autres.filter(e => e.sujet.trim()).length,
        ref = autres.filter(e => e.referent.trim()).length, doc = autres.filter(e => e.nbDocs > 0).length;
  modal(`<h2>${ids.length} étudiant${ids.length > 1 ? 's' : ''} ajouté${ids.length > 1 ? 's' : ''} — ${esc(C.dip)}</h2>
    <p class="sub">Rapprochement avec les mémoires déjà saisis dans le CRM :</p>
    <div style="display:grid;grid-template-columns:repeat(2,1fr);gap:10px;margin:14px 0">
      ${[['Mémoire déjà saisi', avec], ['Sujet fourni', sujet], ['Formateur référent', ref], ['Mémoire joint', doc]].map(([l, n]) => `<div class="ck-card" style="margin:0;padding:14px 16px"><b style="font-size:22px;color:${n === ids.length ? '#15803d' : (n ? '#c2410c' : '#c62828')}">${n}</b><span style="color:#7d8a9c"> / ${ids.length}</span><div style="font-size:12.5px;font-weight:700;margin-top:2px">${l}</div></div>`).join('')}
    </div>
    ${avec < ids.length ? `<p class="sub"><b>Sans mémoire saisi :</b> ${autres.filter(e => !e.memoireId).map(e => esc(e.etudiant)).join(', ')}</p>` : ''}
    <div class="mft"><button class="vx-btn pri" onclick="this.closest('.vx-modal').remove()">OK</button></div>`, 560);
}
function ckDip(d) { C.dip = d; C.sel.clear(); C.q = ''; C.cls = ''; ckRender(); }
function ckYear(v) {
  if (v === '__new') { const a = prompt('Année scolaire (ex : 2027-2028) :', ''); if (!a || !/^\d{4}-\d{4}$/.test(a.trim())) { ckRender(); return; } v = a.trim(); }
  C.data = null; ckLoad(v);
}
async function ckConcerne(ids, on) {
  const all = Object.values(C.data.diplomes).flat();
  all.forEach(e => { if (ids.includes(e.etuId)) e.concerne = on; });
  ckRender();
  const r = await api('saveMemoiresChecklist', { annee: C.annee, etuIds: ids, concerne: on });
  if (!r.success) { all.forEach(e => { if (ids.includes(e.etuId)) e.concerne = !on; }); ckRender(); toast('Erreur : ' + errTxt(r.message), false); return; }
  toast(on ? (ids.length > 1 ? ids.length + ' étudiants ajoutés' : 'Étudiant ajouté') + ' à la checklist ✓' : 'Retiré de la checklist');
}
async function ckCyclade(id, on) {
  const e = Object.values(C.data.diplomes).flat().find(x => x.etuId === id); if (!e) return;
  e.cyclade = on; e.cycladeDate = on ? todayISO() : ''; ckRender();
  const r = await api('saveMemoiresChecklist', { annee: C.annee, cycladeEtuId: id, cyclade: on, cycladeDate: todayISO() });
  if (!r.success) { e.cyclade = !on; ckRender(); toast('Erreur : ' + errTxt(r.message), false); return; }
  toast(on ? 'Dépôt Cyclade enregistré ✓' : 'Dépôt Cyclade annulé');
}
async function ckExport() {
  const r = C.data; if (!r) return;
  toast('Préparation du fichier Excel…');
  let X; try { X = await loadExcelJS(); } catch (e) { toast(e.message, false); return; }
  const wb = new X.Workbook(); wb.creator = 'CRM HECG'; wb.created = new Date();
  const fill = argb => ({ type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + argb } });
  const OK = { bg: 'E8F6EE', fg: '15803D' }, KO = { bg: 'FDE8E8', fg: 'C62828' }, MID = { bg: 'FFF4ED', fg: 'C2410C' };
  // Une ligne par étudiant : identité + mails HECG (étudiant et formateur référent) + 4 contrôles colorés
  const H = ['Étudiant', 'Classe', 'Campus', 'Mail HECG étudiant', 'Sujet du mémoire', 'Mémoire joint', 'Formateur référent', 'Mail formateur référent', 'Dépôt Cyclade', 'Avancement'];
  const W = [26, 13, 15, 32, 46, 20, 24, 32, 17, 13];
  const mailCell = (c, mail, perso) => {
    if (!mail) { c.value = '✗ Non renseigné'; c.font = { italic: true, color: { argb: 'FF' + KO.fg } }; return; }
    c.value = { text: mail + (perso ? ' (perso)' : ''), hyperlink: 'mailto:' + mail };
    c.font = { underline: true, color: { argb: perso ? 'FF94A3B8' : 'FF1D4ED8' }, italic: !!perso };
  };
  const synth = [];
  CK_DIP.forEach(d => {
    const list = (r.diplomes[d] || []).filter(e => e.concerne);
    const ws = wb.addWorksheet(d, { views: [{ state: 'frozen', ySplit: 5, xSplit: 1 }], properties: { tabColor: { argb: 'FF' + XB.blue } }, pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
    ws.columns = W.map(w => ({ width: w }));
    const ok = list.filter(e => ckScore(e).every(Boolean)).length;
    const cnt = j => list.filter(e => ckScore(e)[j]).length;
    synth.push({ d, n: list.length, ok, sujet: cnt(0), doc: cnt(1), ref: cnt(2), cyc: cnt(3) });
    xlTitle(ws, `Suivi des mémoires ${d} — ${r.annee}`, `${list.length} étudiant(s) · ${ok} dossier(s) complet(s) · sujet ${cnt(0)}/${list.length} · mémoire joint ${cnt(1)}/${list.length} · référent ${cnt(2)}/${list.length} · Cyclade ${cnt(3)}/${list.length} · édition du ${fmtLong(todayISO())}${r.archivee ? ' · année archivée' : ''}`, H.length);
    // Légende des couleurs
    const lg = ws.getRow(4); lg.height = 18;
    [[1, '■ Fait', OK], [2, '■ À faire', KO], [3, '■ Partiel', MID]].forEach(([c, t, k]) => { const cell = lg.getCell(c); cell.value = t; cell.font = { bold: true, size: 9, color: { argb: 'FF' + k.fg } }; cell.fill = fill(k.bg); cell.alignment = { horizontal: 'center' }; });
    xlHeader(ws, 5, H);
    list.forEach((e, i) => {
      const sc = ckScore(e), n = sc.filter(Boolean).length, rw = ws.getRow(6 + i);
      rw.values = [e.etudiant, e.classe, e.campus, '', e.sujet || '✗ À définir', e.dernierDoc ? '' : '✗ Aucun', e.referent || '✗ À désigner', '', e.cyclade ? '✓ ' + (e.cycladeDate ? fmtD(e.cycladeDate) : 'Déposé') : '✗ À déposer', n === 4 ? '✓ Complet' : n + ' / 4'];
      rw.height = Math.max(26, Math.min(70, 15 * Math.ceil((e.sujet || '').length / 50)));
      rw.getCell(1).font = { bold: true, color: { argb: 'FF' + XB.ink } };
      mailCell(rw.getCell(4), e.mailHecg || e.mailPerso, !e.mailHecg && !!e.mailPerso);
      if (e.referent) mailCell(rw.getCell(8), e.referentMail, false); else { rw.getCell(8).value = '—'; rw.getCell(8).font = { color: { argb: 'FF' + XB.mu } }; }
      sc.forEach((v, j) => { const c = rw.getCell([5, 6, 7, 9][j]); const k = v ? OK : KO; c.fill = fill(k.bg); if (!(j === 1 && v)) c.font = { color: { argb: 'FF' + k.fg }, bold: !v }; });
      if (e.dernierDoc) { const c = rw.getCell(6); c.value = { text: '✓ ' + e.dernierDoc.version + ' du ' + fmtD(e.dernierDoc.date) + (e.nbDocs > 1 ? ' (' + e.nbDocs + ' versions)' : ''), hyperlink: e.dernierDoc.url }; c.font = { color: { argb: 'FF' + OK.fg }, underline: true, bold: true }; }
      const pc = rw.getCell(10); pc.alignment = { horizontal: 'center', vertical: 'middle' };
      pc.font = { bold: true, color: { argb: n === 4 ? 'FFFFFFFF' : 'FF' + (n >= 2 ? MID.fg : KO.fg) } };
      pc.fill = fill(n === 4 ? '16A34A' : (n >= 2 ? MID.bg : KO.bg));
      [2, 3, 9].forEach(c => rw.getCell(c).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true });
    });
    xlBand(ws, 6, 5 + list.length, H.length);
    if (list.length) ws.autoFilter = { from: { row: 5, column: 1 }, to: { row: 5 + list.length, column: H.length } };
    else { ws.mergeCells(6, 1, 6, H.length); const c = ws.getCell(6, 1); c.value = 'Aucun étudiant coché pour ce diplôme.'; c.font = { italic: true, color: { argb: 'FF' + XB.mu } }; }
  });
  // Synthèse par diplôme (première feuille)
  const sy = wb.addWorksheet('Synthèse', { properties: { tabColor: { argb: 'FF' + XB.orange } } });
  sy.orderNo = -1; // ExcelJS ordonne les feuilles par orderNo : la synthèse passe en premier
  const SH = ['Diplôme', 'Étudiants', 'Sujet fourni', 'Mémoire joint', 'Formateur référent', 'Dépôt Cyclade', 'Dossiers complets'];
  sy.columns = [16, 12, 15, 15, 19, 15, 18].map(w => ({ width: w }));
  xlTitle(sy, `Suivi des mémoires — synthèse ${r.annee}`, `Édition du ${fmtLong(todayISO())}`, SH.length);
  xlHeader(sy, 4, SH);
  synth.forEach((s, i) => {
    const rw = sy.getRow(5 + i); rw.values = [s.d, s.n, s.sujet, s.doc, s.ref, s.cyc, s.ok]; rw.height = 26;
    rw.getCell(1).font = { bold: true };
    [3, 4, 5, 6, 7].forEach(c => { const v = rw.getCell(c).value, k = !s.n ? null : (v === s.n ? OK : (v ? MID : KO)); const cell = rw.getCell(c);
      cell.value = s.n ? `${v} / ${s.n}  (${Math.round(v / s.n * 100)} %)` : '—'; if (k) { cell.fill = fill(k.bg); cell.font = { bold: true, color: { argb: 'FF' + k.fg } }; } });
    for (let c = 1; c <= SH.length; c++) rw.getCell(c).alignment = { horizontal: c === 1 ? 'left' : 'center', vertical: 'middle', indent: c === 1 ? 1 : 0 };
  });
  xlBand(sy, 5, 4 + synth.length, SH.length);
  await xlDownload(wb, `Checklist_memoires_${r.annee}.xlsx`);
  toast('Export Excel prêt ✓');
}

function memKwManage() {
  const usage = k => M.data.filter(m => m.motsCles.some(x => norm(x) === norm(k))).length;
  const dlg = modal(`<h2>Mots-clés des mémoires</h2><p class="sub">Créez vos propres mots-clés : ils sont proposés à toute l'équipe lors de la saisie d'un mémoire.</p>
    <label class="fl">Nouveau mot-clé</label>
    <div style="display:flex;gap:10px"><input class="in" id="kwNew" placeholder="Ex : Contrôle interne, ESG, Blockchain…" maxlength="40"><button class="vx-btn pri" id="kwAdd">${svg('plus', 15, 2.4)} Ajouter</button></div>
    <label class="fl">Liste actuelle <span id="kwN" style="text-transform:none;letter-spacing:0;color:#a3aebd"></span></label>
    <div id="kwList" style="display:flex;flex-wrap:wrap;gap:8px"></div>
    <p class="sub" style="margin-top:14px">Retirer un mot-clé le masque des suggestions ; les mémoires qui l'utilisent le conservent.</p>
    <div class="mft"><button class="vx-btn" onclick="this.closest('.vx-modal').remove()">Fermer</button></div>`, 720);
  const draw = () => {
    const list = M.motsClesRef.slice().sort((a, b) => a.localeCompare(b, 'fr'));
    $('kwN').textContent = '(' + list.length + ')';
    $('kwList').innerHTML = list.map(k => `<span class="vx-kw" style="${kwColor(k)};padding:6px 8px 6px 12px;font-size:12.5px">${esc(k)} <small style="opacity:.6">${usage(k)}</small><button data-kwdel="${esc(k)}" title="Retirer de la liste" style="border:none;background:none;color:inherit;cursor:pointer;opacity:.6;padding:0 0 0 4px">${svg('x', 12, 2.6)}</button></span>`).join('') || '<p class="sub">Aucun mot-clé.</p>';
  };
  const add = async () => {
    const k0 = $('kwNew').value.trim().replace(/\s+/g, ' '); if (!k0) return;
    const k = k0.charAt(0).toUpperCase() + k0.slice(1);
    if (M.motsClesRef.some(x => norm(x) === norm(k))) { toast('« ' + k + ' » existe déjà', false); return; }
    M.motsClesRef.push(k); $('kwNew').value = ''; draw(); $('kwNew').focus();
    const r = await api('saveMotsClesMemoire', { ajouter: [k] });
    if (!r.success) { M.motsClesRef = M.motsClesRef.filter(x => x !== k); draw(); toast('Erreur : ' + errTxt(r.message), false); return; }
    M.motsClesRef = r.motsCles; draw(); toast('Mot-clé « ' + k + ' » ajouté ✓');
  };
  $('kwAdd').onclick = add;
  $('kwNew').addEventListener('keydown', e => { if (e.key === 'Enter') add(); });
  dlg.addEventListener('click', async e => {
    const b = e.target.closest('[data-kwdel]'); if (!b) return;
    const k = b.dataset.kwdel;
    if (!confirm('Retirer « ' + k + ' » des suggestions ?')) return;
    const avant = M.motsClesRef.slice(); M.motsClesRef = M.motsClesRef.filter(x => x !== k); draw();
    const r = await api('saveMotsClesMemoire', { retirer: [k] });
    if (!r.success) { M.motsClesRef = avant; draw(); toast('Erreur : ' + errTxt(r.message), false); return; }
    M.motsClesRef = r.motsCles; draw();
  });
  draw(); $('kwNew').focus();
}

let _exceljs = null;
function loadExcelJS() {
  if (window.ExcelJS) return Promise.resolve(window.ExcelJS);
  if (_exceljs) return _exceljs;
  _exceljs = new Promise((res, rej) => { const sc = document.createElement('script'); sc.src = 'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js'; sc.onload = () => res(window.ExcelJS); sc.onerror = () => { _exceljs = null; rej(new Error('Bibliothèque Excel indisponible (réseau)')); }; document.head.appendChild(sc); });
  return _exceljs;
}
function hslHex(h, s, l) { s /= 100; l /= 100; const k = n => (n + h / 30) % 12, a = s * Math.min(l, 1 - l), f = n => l - a * Math.max(-1, Math.min(k(n) - 3, Math.min(9 - k(n), 1))); return [f(0), f(8), f(4)].map(x => Math.round(x * 255).toString(16).padStart(2, '0')).join('').toUpperCase(); }
const kwHex = k => { const h = hue(norm(k)); return { fg: hslHex(h, 55, 32), bg: hslHex(h, 70, 94) }; };
const XB = { blue: '1A4E8A', orange: 'F26522', ink: '0E1A2B', mu: '7D8A9C', line: 'E3EAF3', band: 'F5F8FC' };
function xlTitle(ws, titre, sousTitre, nbCols) {
  ws.mergeCells(1, 1, 1, nbCols); ws.mergeCells(2, 1, 2, nbCols); ws.mergeCells(3, 1, 3, nbCols);
  const t = ws.getCell(1, 1); t.value = titre; t.font = { name: 'Calibri', size: 18, bold: true, color: { argb: 'FFFFFFFF' } };
  t.fill = { type: 'gradient', gradient: 'angle', degree: 0, stops: [{ position: 0, color: { argb: 'FF1A4E8A' } }, { position: 1, color: { argb: 'FF2D64A8' } }] };
  t.alignment = { vertical: 'middle', indent: 1 }; ws.getRow(1).height = 36;
  const st = ws.getCell(2, 1); st.value = sousTitre; st.font = { size: 10.5, italic: true, color: { argb: 'FF' + XB.mu } }; st.alignment = { indent: 1, vertical: 'middle' }; ws.getRow(2).height = 20;
  ws.getCell(3, 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + XB.orange } }; ws.getRow(3).height = 4;
}
function xlHeader(ws, rowIdx, headers) {
  const r = ws.getRow(rowIdx); r.values = headers; r.height = 28;
  r.eachCell(c => { c.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + XB.blue } };
    c.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }; c.border = { bottom: { style: 'medium', color: { argb: 'FF' + XB.orange } } }; });
}
function xlBand(ws, from, to, nbCols) {
  for (let r = from; r <= to; r++) for (let c = 1; c <= nbCols; c++) {
    const cell = ws.getCell(r, c);
    if (!cell.fill || cell.fill.type !== 'pattern' || !cell.fill.fgColor || cell.fill.fgColor.argb === undefined) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: (r - from) % 2 ? 'FF' + XB.band : 'FFFFFFFF' } };
    cell.border = { top: { style: 'thin', color: { argb: 'FF' + XB.line } }, bottom: { style: 'thin', color: { argb: 'FF' + XB.line } }, left: { style: 'thin', color: { argb: 'FF' + XB.line } }, right: { style: 'thin', color: { argb: 'FF' + XB.line } } };
    cell.alignment = Object.assign({ vertical: 'middle', wrapText: true }, cell.alignment || {});
  }
}
async function xlDownload(wb, name) {
  const buf = await wb.xlsx.writeBuffer();
  const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })); a.download = name; a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 3000);
}
const STX = { 'En cours': ['E8EFF8', '1A4E8A'], 'Sujet validé': ['FFF4ED', 'C2410C'], 'Rendu': ['EDE9FE', '6D28D9'], 'Soutenu': ['E8F6EE', '15803D'] };

async function memExport() {
  const list = memFiltered().slice().sort((a, b) => (a.classe + a.etudiant).localeCompare(b.classe + b.etudiant, 'fr'));
  toast('Préparation du fichier Excel…');
  let X; try { X = await loadExcelJS(); } catch (e) { toast(e.message, false); return; }
  const wb = new X.Workbook(); wb.creator = 'CRM HECG'; wb.created = new Date();
  const ws = wb.addWorksheet('Mémoires', { views: [{ state: 'frozen', ySplit: 4 }], properties: { tabColor: { argb: 'FF' + XB.blue } } });
  const H = ['Étudiant', 'Mail étudiant', 'Classe', 'Campus', 'Sujet du mémoire', 'Mots-clés', 'Formateur référent', 'Mail formateur', 'Statut', 'Dernier document', 'Date dernier doc.', 'Nb doc.', 'Notes', 'Entreprise'];
  ws.columns = [26, 30, 13, 16, 46, 34, 24, 30, 14, 22, 15, 9, 34, 28].map(w => ({ width: w }));
  xlTitle(ws, 'Suivi des mémoires — HECG', `Export du ${fmtLong(todayISO())} · ${list.length} mémoire(s)` + (list.length !== M.data.length ? ' (filtrés)' : ''), H.length);
  xlHeader(ws, 4, H);
  list.forEach((m, i) => {
    const d = m.docs[m.docs.length - 1];
    const r = ws.getRow(5 + i);
    r.values = [m.etudiant, m.etuMail || '', m.classe, m.campus, m.sujet, '', m.formateur, m.formateurMail || '', m.statut, d ? d.version + ' — ' + d.nom : '', d ? new Date(d.date + 'T12:00:00') : '', m.docs.length || '', m.notes || '', m.entreprise || ''];
    r.height = Math.max(24, Math.min(80, 15 * Math.ceil((m.sujet || '').length / 55)));
    r.getCell(1).font = { bold: true, color: { argb: 'FF' + XB.ink } };
    if (m.etuMail) { r.getCell(2).value = { text: m.etuMail, hyperlink: 'mailto:' + m.etuMail }; r.getCell(2).font = { color: { argb: 'FF1D4ED8' }, underline: true }; }
    if (m.formateurMail) { r.getCell(8).value = { text: m.formateurMail, hyperlink: 'mailto:' + m.formateurMail }; r.getCell(8).font = { color: { argb: 'FF1D4ED8' }, underline: true }; }
    if (m.motsCles.length) r.getCell(6).value = { richText: m.motsCles.flatMap((k, j) => { const c = kwHex(k); return (j ? [{ text: '   ', font: { size: 10 } }] : []).concat([{ text: '● ' + k, font: { bold: true, size: 10.5, color: { argb: 'FF' + c.fg } } }]); }) };
    const sx = STX[m.statut] || STX['En cours'];
    r.getCell(9).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + sx[0] } }; r.getCell(9).font = { bold: true, color: { argb: 'FF' + sx[1] } }; r.getCell(9).alignment = { horizontal: 'center', vertical: 'middle' };
    if (d) { r.getCell(10).value = { text: d.version + ' — ' + d.nom, hyperlink: d.url }; r.getCell(10).font = { color: { argb: 'FF1D4ED8' }, underline: true }; r.getCell(11).numFmt = 'dd/mm/yyyy'; }
    [3, 4, 11, 12].forEach(c => r.getCell(c).alignment = { horizontal: 'center', vertical: 'middle' });
    if (!m.formateur) { r.getCell(7).value = 'À désigner'; r.getCell(7).font = { italic: true, color: { argb: 'FFC2410C' } }; }
  });
  xlBand(ws, 5, 4 + list.length, H.length);
  ws.autoFilter = { from: { row: 4, column: 1 }, to: { row: 4 + list.length, column: H.length } };
  // Synthèse : mots-clés colorés + statuts
  const sy = wb.addWorksheet('Synthèse', { properties: { tabColor: { argb: 'FF' + XB.orange } } });
  sy.columns = [{ width: 32 }, { width: 14 }, { width: 6 }, { width: 22 }, { width: 14 }];
  xlTitle(sy, 'Synthèse des mémoires', `${list.length} mémoire(s) · ${new Set(list.map(m => m.classe)).size} classe(s)`, 5);
  sy.getRow(4).values = ['Mot-clé', 'Mémoires', '', 'Statut', 'Mémoires'];
  [1, 2, 4, 5].forEach(c => { const cell = sy.getCell(4, c); cell.font = { bold: true, color: { argb: 'FFFFFFFF' } }; cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + XB.blue } }; cell.alignment = { horizontal: 'center', vertical: 'middle' }; });
  sy.getRow(4).height = 26;
  const kwc = {}; list.forEach(m => m.motsCles.forEach(k => kwc[k] = (kwc[k] || 0) + 1));
  Object.entries(kwc).sort((a, b) => b[1] - a[1]).forEach(([k, n], i) => { const c = kwHex(k); const a1 = sy.getCell(5 + i, 1), b1 = sy.getCell(5 + i, 2);
    a1.value = k; a1.font = { bold: true, color: { argb: 'FF' + c.fg } }; a1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + c.bg } };
    b1.value = n; b1.alignment = { horizontal: 'center' }; b1.font = { bold: true }; b1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + c.bg } }; sy.getRow(5 + i).height = 22; });
  Object.keys(STX).forEach((st, i) => { const sx = STX[st]; const a1 = sy.getCell(5 + i, 4), b1 = sy.getCell(5 + i, 5);
    a1.value = st; b1.value = list.filter(m => m.statut === st).length;
    [a1, b1].forEach(c => { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + sx[0] } }; c.font = { bold: true, color: { argb: 'FF' + sx[1] } }; }); b1.alignment = { horizontal: 'center' }; });
  await xlDownload(wb, 'Memoires_HECG_' + todayISO() + '.xlsx');
  toast('Export Excel prêt ✓');
}


// ═══════════════════════════ FRISES ═══════════════════════════
const F = { data: [], cur: null };
const REMS = [[30, 'J-30'], [14, 'J-14'], [7, 'J-7'], [3, 'J-3'], [2, 'J-2'], [1, 'Veille'], [0, 'Jour J']];

async function frLoad(force) {
  const r = await api('getFrises', {}, force ? { forceRefresh: true } : {});
  if (!r.success) { $('frBody').innerHTML = `<div class="vx-empty"><div class="ill">${svg('clock', 26)}</div><h3>Chargement impossible</h3><p>${esc(errTxt(r.message))}</p></div>`; return; }
  F.data = r.data || []; F.personnes = r.personnes || [];
  if (!F.data.find(f => f.id === F.cur)) F.cur = (F.data[0] || {}).id || null;
  frRender();
}
function frInit() {
  injectCss();
  const sec = $('section-frises');
  if (!sec.dataset.ready) {
    sec.dataset.ready = '1'; sec.classList.add('vx');
    sec.innerHTML = `<div class="vx-hd"><div><h1>Frises chronologiques</h1><p>Échéances et livrables, avec rappels automatiques par mail</p></div>
      <button class="vx-btn ora" onclick="VX.frEdit()">${svg('plus', 15, 2.4)} Nouvelle frise</button></div>
      <div id="frBody"><div class="vx-empty"><div class="ill">${svg('clock', 26)}</div><h3>Chargement…</h3></div></div>`;
  }
  frLoad();
}
function nextStep(f) { return f.etapes.filter(e => e.statut !== 'Rendu').sort((a, b) => a.date.localeCompare(b.date))[0]; }
function jLabel(n) { return n === 0 ? "Aujourd'hui" : (n === 1 ? 'Demain' : (n < 0 ? `En retard de ${-n} j` : `Dans ${n} jours`)); }
function frRender() {
  const body = $('frBody');
  if (!F.data.length) { body.innerHTML = `<div class="vx-empty"><div class="ill">${svg('clock', 26)}</div><h3>Aucune frise</h3><p>Créez une frise, ajoutez vos dates et ce qu'il faut rendre : vous recevrez des rappels par mail avant chaque échéance.</p><button class="vx-btn ora" style="margin-top:14px" onclick="VX.frEdit()">${svg('plus', 15, 2.4)} Créer ma première frise</button></div>`; return; }
  const cards = F.data.map(f => {
    const n = nextStep(f), done = f.etapes.filter(e => e.statut === 'Rendu').length, tot = f.etapes.length || 1;
    return `<div class="fr-item ${f.id === F.cur ? 'on' : ''}" style="--c:${esc(f.couleur)}" onclick="VX.frSel('${jsq(f.id)}')">
      <h4>${esc(f.titre)}</h4>
      <div class="nx">${n ? `${svg('flag', 13)} ${esc(n.titre)} · <b style="color:${daysTo(n.date) <= 2 ? '#e5484d' : 'inherit'}">${jLabel(daysTo(n.date))}</b>` : `${svg('check', 13)} Toutes les étapes sont rendues`}</div>
      <div class="fr-prog"><i style="width:${Math.round(done / tot * 100)}%"></i></div>
      <div class="nx" style="margin-top:6px">${done}/${f.etapes.length} étape(s) rendue(s)</div></div>`;
  }).join('');
  const f = F.data.find(x => x.id === F.cur);
  body.innerHTML = `<div class="fr-list">${cards}</div>` + (f ? frBoard(f) : '');
  const sc = document.querySelector('.fr-scroll'), td = document.querySelector('.fr-today');
  if (sc && td) sc.scrollLeft = Math.max(0, td.offsetLeft - sc.clientWidth / 3);
}
function frBoard(f) {
  const steps = f.etapes.slice().sort((a, b) => a.date.localeCompare(b.date));
  const n = nextStep(f);
  if (!steps.length) return `<div class="fr-board" style="--c:${esc(f.couleur)}"><div class="fr-head"><h2><i></i>${esc(f.titre)}</h2><button class="vx-btn" onclick="VX.frEdit('${jsq(f.id)}')">${svg('edit', 14)} Modifier</button></div><div class="vx-empty" style="margin-top:18px"><h3>Aucune échéance</h3><p>Ajoutez des dates et ce qu'il faut rendre.</p></div></div>`;
  const today = todayISO();
  const dates = steps.map(s => s.date).concat([today]).sort();
  const t0 = new Date(dates[0] + 'T12:00:00'), t1 = new Date(dates[dates.length - 1] + 'T12:00:00');
  const span = Math.max(1, (t1 - t0) / 86400000);
  const W = Math.max(steps.length * 260, 980), PAD = 140;
  const x = iso => PAD + ((new Date(iso + 'T12:00:00') - t0) / 86400000) / span * (W - 2 * PAD);
  // décale les points trop proches pour éviter les chevauchements
  const xs = []; steps.forEach((s, i) => { let v = x(s.date); if (i && v - xs[i - 1] < 130) v = xs[i - 1] + 130; xs.push(v); });
  const width = Math.max(W, xs[xs.length - 1] + PAD);
  const doneCount = steps.filter(s => s.statut === 'Rendu').length;
  const progX = Math.min(x(today), width);
  return `<div class="fr-board" style="--c:${esc(f.couleur)}">
    <div class="fr-head"><div><h2><i></i>${esc(f.titre)}</h2>${f.description ? `<p style="margin:6px 0 0;color:var(--mu)">${esc(f.description)}</p>` : ''}
      <div style="display:flex;flex-wrap:wrap;gap:6px;margin-top:10px"><span class="vx-chip">${svg('bell', 12)} Rappels : ${String(f.rappels).split(',').map(r => +r === 0 ? 'jour J' : (+r === 1 ? 'veille' : 'J-' + r)).join(', ')}</span>${frDestChip(f)}<span class="vx-chip b">${doneCount}/${steps.length} rendue(s)</span></div></div>
      <div style="display:flex;gap:8px"><button class="vx-btn" onclick="VX.frPdf('${jsq(f.id)}')">${svg('down', 14)} Exporter en PDF</button><button class="vx-btn" onclick="VX.frEdit('${jsq(f.id)}')">${svg('edit', 14)} Modifier</button></div></div>
    ${n ? `<div class="fr-next"><div class="fr-cd"><b>${daysTo(n.date) < 0 ? '!' : (daysTo(n.date) === 0 ? 'J' : 'J-' + daysTo(n.date))}</b><span>${daysTo(n.date) < 0 ? 'retard' : (daysTo(n.date) === 0 ? "aujourd'hui" : 'jours')}</span></div>
      <div><div style="font-size:11px;font-weight:800;letter-spacing:.07em;text-transform:uppercase;color:var(--mu)">Prochaine échéance · ${fmtLong(n.date)}</div><div style="font-size:16px;font-weight:700;color:var(--ink);margin-top:2px">${esc(n.titre)}</div>${n.aRendre ? `<div style="color:var(--tx);margin-top:2px">À rendre : ${esc(n.aRendre)}</div>` : ''}</div>
      <button class="vx-btn sm" style="margin-left:auto" onclick="VX.frToggle('${jsq(n.id)}',true)">${svg('check', 14, 2.4)} Marquer rendu</button></div>` : ''}
    <div class="fr-scroll"><div class="fr-track" style="width:${width}px">
      <div class="fr-line"><i style="width:${progX}px;opacity:.35"></i></div>
      <div class="fr-today" style="left:${x(today)}px"><span>Aujourd'hui</span></div>
      ${steps.map((s, i) => {
        const j = daysTo(s.date), done = s.statut === 'Rendu', late = !done && j < 0;
        const cls = done ? 'ok' : (late ? 'late' : (j <= 3 ? 'soon' : ''));
        return `<div class="fr-pt ${done ? 'done' : ''} ${late ? 'late' : ''} ${n && n.id === s.id ? 'next' : ''}" style="left:${xs[i]}px" title="${done ? 'Rendu — cliquer pour annuler' : 'Cliquer pour marquer rendu'}" onclick="VX.frToggle('${jsq(s.id)}',${!done})">${done ? svg('check', 12, 3) : ''}</div>
        <div class="fr-card ${i % 2 ? 'dn' : 'up'} ${done ? 'done' : ''}" style="left:${xs[i]}px"><div class="d">${fmtLong(s.date)}</div><div class="t">${esc(s.titre)}</div>${s.aRendre ? `<div class="r">${esc(s.aRendre)}</div>` : ''}<span class="j ${cls}">${done ? 'Rendu' + (s.renduLe ? ' le ' + fmtD(s.renduLe) : '') : jLabel(j)}</span></div>`;
      }).join('')}
    </div></div></div>`;
}
function frSel(id) { F.cur = id; frRender(); }
const FR_TYPES = { etu: ['Étudiants', 'étudiant'], equipe: ['Équipe', 'membre de l\'équipe'], form: ['Formateurs', 'formateur'] };
function frDestChip(f) {
  const tous = String(f.destinataires || f.emailCreateur || '').split(/[,;\s]+/).filter(x => x.includes('@'));
  const parts = f.participants || [], n = t => parts.filter(p => p.t === t).length;
  const det = ['etu', 'equipe', 'form'].filter(t => n(t)).map(t => n(t) + ' ' + FR_TYPES[t][1] + (n(t) > 1 ? 's' : '')).join(' · ');
  const titre = parts.map(p => p.nom + ' <' + p.email + '>').concat(tous.filter(e => !parts.some(p => p.email === e))).join('\n');
  return `<span class="vx-chip" title="${esc(titre)}">${svg('mail', 12)} ${tous.length} destinataire(s) des rappels${det ? ' — ' + esc(det) : ''}</span>`;
}

// Export PDF : document A4 paysage mis en page (frise + tableau détaillé), puis impression
// « Enregistrer au format PDF » du navigateur.
function frPdf(id) {
  const f = F.data.find(x => x.id === id); if (!f) return;
  const steps = f.etapes.slice().sort((a, b) => a.date.localeCompare(b.date));
  const c = f.couleur || '#1A4E8A';
  const today = todayISO();
  const dates = steps.map(s => s.date).concat([today]).sort();
  const t0 = new Date(dates[0] + 'T12:00:00'), t1 = new Date(dates[dates.length - 1] + 'T12:00:00');
  const span = Math.max(1, (t1 - t0) / 86400000);
  const pos = iso => 4 + ((new Date(iso + 'T12:00:00') - t0) / 86400000) / span * 92;
  const xs = []; steps.forEach((s, i) => { let v = pos(s.date); if (i && v - xs[i - 1] < 9) v = xs[i - 1] + 9; xs.push(Math.min(v, 98)); });
  const done = steps.filter(s => s.statut === 'Rendu').length;
  const n = nextStep(f);
  const etat = s => { const j = daysTo(s.date); return s.statut === 'Rendu' ? ['Rendu' + (s.renduLe ? ' le ' + fmtD(s.renduLe) : ''), '#15803d', '#e8f6ee'] : (j < 0 ? ['En retard de ' + (-j) + ' j', '#c62828', '#fde8e8'] : (j === 0 ? ["Aujourd'hui", '#c2410c', '#fff1e8'] : ['Dans ' + j + ' j', '#475569', '#eef2f7'])); };
  const html = `<!DOCTYPE html><html lang="fr"><head><meta charset="utf-8"><title>${esc(f.titre)} — frise chronologique</title>
  <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap">
  <style>
    @page{size:A4 landscape;margin:14mm}
    *{box-sizing:border-box;-webkit-print-color-adjust:exact;print-color-adjust:exact}
    body{margin:0;font-family:'Plus Jakarta Sans',Inter,Arial,sans-serif;color:#0e1a2b;font-size:12px}
    .hd{display:flex;align-items:flex-end;justify-content:space-between;border-bottom:3px solid ${c};padding-bottom:10px}
    .hd h1{margin:0;font-size:24px;letter-spacing:-.02em}
    .hd p{margin:4px 0 0;color:#64748b;font-size:12.5px}
    .hd .r{text-align:right;color:#64748b;font-size:11px}
    .hd img{height:30px;margin-bottom:6px}
    .kp{display:flex;gap:10px;margin:14px 0 6px}
    .kp div{flex:1;padding:9px 12px;border-radius:10px;background:#f4f7fb;border:1px solid #e3eaf3}
    .kp span{display:block;font-size:9.5px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:#7d8a9c}
    .kp b{font-size:15px}
    .tl{position:relative;height:230px;margin:10px 0 4px}
    .ln{position:absolute;left:2%;right:1%;top:112px;height:5px;border-radius:5px;background:#e3e9f1}
    .td{position:absolute;top:84px;height:60px;border-left:2px dashed #F26522}
    .td{z-index:3}
    .td span{position:absolute;bottom:-17px;left:-34px;width:68px;text-align:center;font-size:9px;font-weight:800;color:#F26522;background:#fff4ed;border-radius:99px;padding:1px 0;z-index:3}
    .pt{position:absolute;top:104px;width:21px;height:21px;margin-left:-10.5px;border-radius:50%;border:4px solid ${c};background:#fff}
    .pt.ok{background:#16a34a;border-color:#16a34a}.pt.late{border-color:#e5484d;background:#fff1f1}
    .cd{position:absolute;width:138px;margin-left:-69px;padding:7px 9px;border-radius:9px;background:#fff;border:1px solid #dfe6ef;box-shadow:0 4px 10px -6px rgba(14,26,43,.35)}
    .cd.up{bottom:132px}.cd.dn{top:136px}
    .cd .d{font-size:9px;font-weight:800;letter-spacing:.04em;text-transform:uppercase;color:${c}}
    .cd .t{font-size:11px;font-weight:700;margin-top:1px;line-height:1.25}
    .cd .s{font-size:9px;font-weight:700;display:inline-block;margin-top:4px;padding:1px 6px;border-radius:99px}
    table{width:100%;border-collapse:collapse;margin-top:14px;font-size:11.5px}
    th{text-align:left;font-size:9.5px;letter-spacing:.07em;text-transform:uppercase;color:#7d8a9c;padding:7px 9px;border-bottom:2px solid #e3eaf3;background:#f7f9fc}
    td{padding:8px 9px;border-bottom:1px solid #edf1f6;vertical-align:top}
    tr{page-break-inside:avoid}
    .badge{font-size:10px;font-weight:700;padding:2px 8px;border-radius:99px;white-space:nowrap}
    .ft{margin-top:12px;font-size:9.5px;color:#94a3b8;display:flex;justify-content:space-between}
  </style></head><body>
  <div class="hd"><div><img src="https://hecg.fr/wp-content/uploads/2025/03/LOGO-FORMAT-PNG-1-1.png" onerror="this.remove()"><h1>${esc(f.titre)}</h1>${f.description ? `<p>${esc(f.description)}</p>` : ''}</div>
    <div class="r">Frise chronologique<br>Édition du ${fmtLong(today)}</div></div>
  <div class="kp"><div><span>Échéances</span><b>${steps.length}</b></div><div><span>Rendues</span><b>${done} / ${steps.length}</b></div>
    <div><span>Prochaine échéance</span><b>${n ? esc(n.titre) + ' · ' + fmtLong(n.date) : 'Aucune'}</b></div>
    <div><span>Période</span><b>${steps.length ? fmtLong(steps[0].date) + ' → ' + fmtLong(steps[steps.length - 1].date) : '—'}</b></div></div>
  <div class="tl"><div class="ln"></div>
    ${today >= dates[0] && today <= dates[dates.length - 1] ? `<div class="td" style="left:${pos(today)}%"><span>Aujourd'hui</span></div>` : ''}
    ${steps.map((s, i) => { const e = etat(s); return `<div class="pt ${s.statut === 'Rendu' ? 'ok' : (daysTo(s.date) < 0 ? 'late' : '')}" style="left:${xs[i]}%"></div>
      <div class="cd ${i % 2 ? 'dn' : 'up'}" style="left:${Math.max(7, Math.min(93, xs[i]))}%"><div class="d">${fmtLong(s.date)}</div><div class="t">${esc(s.titre)}</div><span class="s" style="color:${e[1]};background:${e[2]}">${e[0]}</span></div>`; }).join('')}
  </div>
  <table><thead><tr><th style="width:110px">Date</th><th>Échéance</th><th>À rendre</th><th style="width:150px">Statut</th></tr></thead><tbody>
    ${steps.map(s => { const e = etat(s); return `<tr><td><b>${fmtLong(s.date)}</b></td><td>${esc(s.titre)}</td><td>${esc(s.aRendre || '—')}</td><td><span class="badge" style="color:${e[1]};background:${e[2]}">${e[0]}</span></td></tr>`; }).join('')}
  </tbody></table>
  <div class="ft"><span>Rappels automatiques : ${String(f.rappels).split(',').map(r => +r === 0 ? 'jour J' : (+r === 1 ? 'veille' : 'J-' + r)).join(', ')}</span><span>CRM HECG</span></div>
  <script>window.onload=function(){setTimeout(function(){window.print();},450);};<\/script>
  </body></html>`;
  const w = window.open('', '_blank');
  if (!w) { toast('Autorisez les fenêtres pop-up pour exporter le PDF', false); return; }
  w.document.open(); w.document.write(html); w.document.close();
}
async function frToggle(id, rendu) {
  const f = F.data.find(x => x.etapes.some(e => e.id === id)); const e = f && f.etapes.find(x => x.id === id);
  if (!e) return;
  e.statut = rendu ? 'Rendu' : 'À faire'; e.renduLe = rendu ? todayISO() : ''; frRender();
  const r = await api('toggleEtapeFrise', { id, rendu });
  if (!r.success) { e.statut = rendu ? 'À faire' : 'Rendu'; frRender(); toast('Erreur : ' + errTxt(r.message), false); return; }
  toast(rendu ? 'Étape marquée rendue ✓' : 'Étape remise à faire');
}
function frEdit(id) {
  const f = id ? F.data.find(x => x.id === id) : null;
  const E = { couleur: f ? f.couleur : PALETTE[0], rappels: (f ? f.rappels : '7,2,0').split(',').map(Number), etapes: f ? f.etapes.map(e => ({ ...e })) : [{ date: '', titre: '', aRendre: '' }],
    parts: f ? (f.participants || []).map(p => ({ ...p })) : [], tab: 'etu', q: '' };
  const autres = f ? String(f.destinataires || '').split(/[,;\s]+/).filter(e => e.includes('@') && !E.parts.some(p => p.email === e.toLowerCase())).join(', ') : (user.email || '');
  const dlg = modal(`<h2>${f ? 'Modifier la frise' : 'Nouvelle frise'}</h2><p class="sub">Un mail de rappel est envoyé avant chaque échéance non rendue à toutes les personnes ajoutées ci-dessous (étudiants, équipe, formateurs).</p>
    <label class="fl">Titre</label><input class="in" id="frT" value="${esc(f ? f.titre : '')}" placeholder="Ex : Calendrier des mémoires BTS CG 2026">
    <label class="fl">Description</label><input class="in" id="frD" value="${esc(f ? f.description : '')}" placeholder="Facultatif">
    <label class="fl">Couleur</label><div style="display:flex;gap:8px" id="frC"></div>
    <label class="fl">Échéances</label><div class="fr-steps" id="frSteps"></div>
    <button class="vx-btn sm" style="margin-top:10px" id="frAdd">${svg('plus', 14, 2.4)} Ajouter une échéance</button>
    <label class="fl">Rappels par mail</label><div class="fr-rem" id="frR">${REMS.map(([n, l]) => `<label><input type="checkbox" value="${n}" ${E.rappels.includes(n) ? 'checked' : ''}>${l}</label>`).join('')}</div>
    <label class="fl">Personnes qui reçoivent les rappels</label>
    <div class="fr-pp"><div class="fr-sel" id="frSel"></div>
      <div class="fr-tabs" id="frTabs"></div>
      <div style="display:flex;gap:8px;margin-top:8px;flex-wrap:wrap"><input class="in" id="frQ" style="flex:1;min-width:220px" placeholder="Rechercher un nom, une classe, un mail…"><select class="in" id="frCls" style="width:auto;flex:0 0 auto"></select></div>
      <div class="fr-res" id="frRes"></div></div>
    <label class="fl">Autres adresses mail (facultatif)</label><input class="in" id="frDest" value="${esc(autres)}" placeholder="email1@hecg.fr, email2@hecg.fr">
    <div class="mft">${f ? `<button class="vx-btn danger" style="margin-right:auto" id="frDel">${svg('trash', 14)} Supprimer</button>` : ''}<button class="vx-btn" onclick="this.closest('.vx-modal').remove()">Annuler</button><button class="vx-btn pri" id="frSave">Enregistrer</button></div>`, 900);
  const drawC = () => { $('frC').innerHTML = PALETTE.map(c => `<span class="vx-dot ${c === E.couleur ? 'on' : ''}" style="background:${c}" data-c="${c}"></span>`).join(''); };
  const readSteps = () => { dlg.querySelectorAll('.fr-step').forEach((row, i) => { E.etapes[i].date = row.querySelector('[data-k=date]').value; E.etapes[i].titre = row.querySelector('[data-k=titre]').value; E.etapes[i].aRendre = row.querySelector('[data-k=ar]').value; }); };
  const drawS = () => {
    $('frSteps').innerHTML = E.etapes.map((e, i) => `<div class="fr-step"><input type="date" class="in" data-k="date" value="${esc(e.date)}"><input class="in" data-k="titre" value="${esc(e.titre)}" placeholder="Échéance (ex : Remise du plan)"><input class="in" data-k="ar" value="${esc(e.aRendre)}" placeholder="À rendre (ex : plan détaillé + biblio)"><button class="vx-btn ico danger" data-del="${i}" title="Retirer">${svg('trash', 14)}</button></div>`).join('');
  };
  const P = F.personnes || [];
  const drawP = () => {
    $('frSel').innerHTML = E.parts.map((p, i) => `<span class="vx-chip ${p.t}" title="${esc(FR_TYPES[p.t][1] + ' — ' + p.email)}">${esc(p.nom || p.email)}<button type="button" data-rm="${i}" title="Retirer">×</button></span>`).join('');
    $('frTabs').innerHTML = Object.keys(FR_TYPES).map(t => `<button type="button" data-tab="${t}" class="${E.tab === t ? 'on' : ''}">${FR_TYPES[t][0]} (${P.filter(p => p.t === t).length})</button>`).join('')
      + `<span style="margin-left:auto;font-size:12px;color:var(--mu)">${E.parts.length} sélectionné(s)</span>`;
    const pris = new Set(E.parts.map(p => p.email)), q = norm(E.q);
    const liste = P.filter(p => p.t === E.tab && !pris.has(p.email) && (!q || norm(p.nom + ' ' + p.info + ' ' + p.email).includes(q)));
    $('frRes').innerHTML = liste.slice(0, 60).map(p => `<div class="o" data-add="${esc(p.email)}"><span class="plus">+</span>${esc(p.nom)}<small>${esc(p.info || p.email)}</small></div>`).join('')
      || `<div style="padding:8px;color:var(--mu);font-size:13px">${P.length ? 'Aucune personne à ajouter' : 'Liste des personnes indisponible (rechargez la page)'}</div>`;
    const cls = [...new Set(P.filter(p => p.t === 'etu').map(p => p.info).filter(Boolean))].sort();
    $('frCls').style.display = E.tab === 'etu' ? '' : 'none';
    $('frCls').innerHTML = '<option value="">+ Ajouter toute une classe</option>' + cls.map(c => `<option value="${esc(c)}">${esc(c)} (${P.filter(p => p.t === 'etu' && p.info === c).length})</option>`).join('');
  };
  const addP = p => { if (p && !E.parts.some(x => x.email === p.email)) E.parts.push({ t: p.t, id: p.id, nom: p.nom, email: p.email }); };
  $('frQ').oninput = () => { E.q = $('frQ').value; drawP(); };
  $('frCls').onchange = () => { const c = $('frCls').value; if (!c) return; P.filter(p => p.t === 'etu' && p.info === c).forEach(addP); toast(`Classe ${c} ajoutée aux rappels`); drawP(); };
  drawC(); drawS(); drawP();
  dlg.addEventListener('click', ev => {
    const tb = ev.target.closest('[data-tab]'); if (tb) { E.tab = tb.dataset.tab; drawP(); return; }
    const ad = ev.target.closest('[data-add]'); if (ad) { addP(P.find(p => p.email === ad.dataset.add && p.t === E.tab)); drawP(); return; }
    const rm = ev.target.closest('[data-rm]'); if (rm) { E.parts.splice(+rm.dataset.rm, 1); drawP(); return; }
    const c = ev.target.closest('[data-c]'); if (c) { E.couleur = c.dataset.c; drawC(); }
    const d = ev.target.closest('[data-del]'); if (d) { readSteps(); E.etapes.splice(+d.dataset.del, 1); drawS(); }
  });
  $('frAdd').onclick = () => { readSteps(); E.etapes.push({ date: '', titre: '', aRendre: '' }); drawS(); const ins = dlg.querySelectorAll('.fr-step [data-k=date]'); ins[ins.length - 1].focus(); };
  $('frSave').onclick = async () => {
    readSteps();
    const titre = $('frT').value.trim(); if (!titre) { toast('Donnez un titre à la frise', false); return; }
    const incompletes = E.etapes.filter(e => (e.date || e.titre) && !(e.date && e.titre.trim()));
    if (incompletes.length) { toast('Chaque échéance doit avoir une date et un titre', false); return; }
    const rappels = [...dlg.querySelectorAll('#frR input:checked')].map(i => i.value).join(',');
    const r = await api('saveFrise', { frise: { id: f ? f.id : '', titre, description: $('frD').value.trim(), couleur: E.couleur, rappels: rappels || '0', destinataires: $('frDest').value, participants: E.parts, etapes: E.etapes.filter(e => e.date && e.titre.trim()) } });
    if (!r.success) { toast('Erreur : ' + errTxt(r.message), false); return; }
    if (!E.parts.length && !$('frDest').value.includes('@')) toast('Aucun destinataire : les rappels iront au créateur de la frise', false);
    dlg.remove(); F.cur = r.id; toast(f ? 'Frise mise à jour ✓' : 'Frise créée ✓ — rappels activés'); frLoad(true);
  };
  if ($('frDel')) $('frDel').onclick = async () => {
    if (!confirm('Supprimer cette frise ? Les rappels s\'arrêteront (elle reste archivée dans la feuille FRISES).')) return;
    const r = await api('saveFrise', { frise: { id: f.id, titre: f.titre, archiver: true } });
    if (!r.success) { toast('Erreur : ' + errTxt(r.message), false); return; }
    dlg.remove(); toast('Frise supprimée'); frLoad(true);
  };
  if (!f) $('frT').focus();
}

// ═══════════════════════════ CSS des modules RDV / Directeurs / Examens ═══════════════════════════
function injectCss2() {
  if ($('vxCss2')) return;
  const st = document.createElement('style'); st.id = 'vxCss2';
  st.textContent = `
  .vs2-t{width:100%;border-collapse:separate;border-spacing:0;background:#fff;border:1px solid #e3eaf3;border-radius:16px;overflow:hidden;font-size:13px}
  .vs2-t th{text-align:left;font-size:11px;font-weight:800;color:#7d8a9c;text-transform:uppercase;padding:10px 12px;background:#f7f9fc;border-bottom:1px solid #e3eaf3;white-space:nowrap}
  .vs2-t td{padding:10px 12px;border-top:1px solid #eef2f7;vertical-align:top}
  .vs2-t tr.gr td{background:#f3f7fc;font-size:12.5px;padding:7px 12px}
  .vs2-t tr.gr td b{color:#1A4E8A;font-weight:800;margin-right:10px}
  .vs2-t tr.gr td span{color:#7d8a9c;font-weight:600;margin-right:10px}
  .vs2-n b{display:block;color:#0e1a2b;font-weight:700}.vs2-n small{color:#7d8a9c;font-size:11.5px}
  .vs2-n a{color:#1A4E8A;font-weight:700;text-decoration:none}.vs2-n a:hover{text-decoration:underline}
  .rdv-c{min-width:190px;border:1px solid #e3eaf3;border-radius:12px;padding:8px 10px;background:#fbfcfe;cursor:pointer}
  .rdv-c:hover{border-color:#1A4E8A}
  .rdv-c .d{font-weight:800;color:#0e1a2b;font-size:13px}.rdv-c .s{font-size:11.5px;color:#64748b}
  .rdv-c .rp{font-size:11.5px;color:#c2410c;font-weight:700}
  .rdv-b{display:flex;gap:4px;margin-top:6px;flex-wrap:wrap}
  .rdv-b button{font:700 11px 'Plus Jakarta Sans',Inter,sans-serif;border:1px solid #dfe6f0;background:#fff;border-radius:8px;padding:3px 7px;cursor:pointer}
  .rdv-b button.h{color:#15803d}.rdv-b button.r{color:#c2410c}.rdv-b button.a{color:#c62828}.rdv-b button.o{color:#475569}
  .rdv-b button:hover{filter:brightness(.95);border-color:currentColor}
  .st-b{display:inline-block;font-size:11px;font-weight:800;padding:2px 8px;border-radius:999px}
  .st-Planifié{background:#e8eff8;color:#1A4E8A}.st-Honoré{background:#e8f6ee;color:#15803d}.st-Reporté{background:#fff4ed;color:#c2410c}
  .st-Annulé{background:#fde8e8;color:#c62828}.st-Autre{background:#eef2f7;color:#475569}.st-Àplanifier{background:#f1f4f8;color:#7d8a9c}
  .st-retard{box-shadow:0 0 0 2px #fca5a5}
  .vs2-chips{display:flex;flex-wrap:wrap;gap:6px;margin:0 0 12px}
  .vs2-chips button{font:800 12px 'Plus Jakarta Sans',Inter,sans-serif;border-radius:10px;padding:6px 11px;border:1px solid transparent;cursor:pointer}
  .vs2-chips button.on{box-shadow:0 0 0 2px #1A4E8A}
  .c-ok{background:#e8f6ee;color:#15803d}.c-ko{background:#fde8e8;color:#c62828}.c-wa{background:#fff4ed;color:#c2410c}.c-ne{background:#eef2f7;color:#475569}
  .vs2-band{display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:14px;margin-bottom:16px}
  .vs2-cmp{border-radius:20px;padding:16px 20px;border:2px solid;display:flex;align-items:center;gap:18px;flex-wrap:wrap}
  .vs2-cmp.ok{background:#f0fbf4;border-color:#bbf7d0}.vs2-cmp.ko{background:#fef2f2;border-color:#fecaca}
  .vs2-cmp .k{font-size:11px;font-weight:800;text-transform:uppercase;color:#64748b}.vs2-cmp .v{font-size:30px;font-weight:800;color:#0e1a2b;line-height:1.1}
  .vs2-cmp .eq{font-size:28px;font-weight:900}.vs2-cmp.ok .eq{color:#16a34a}.vs2-cmp.ko .eq{color:#dc2626}
  .vs2-cmp .t{flex-basis:100%;font-size:12.5px;font-weight:700}.vs2-cmp.ok .t{color:#166534}.vs2-cmp.ko .t{color:#991b1b}
  .sw{display:inline-block;width:12px;height:12px;border-radius:3px;border:1px solid #cbd5e1;vertical-align:-2px;margin-right:4px}
  .ex-tg{font:800 12px 'Plus Jakarta Sans',Inter,sans-serif;border-radius:10px;padding:6px 10px;border:1px dashed #cbd5e1;background:#fff;color:#7d8a9c;cursor:pointer;white-space:nowrap}
  .ex-tg.on{border:1px solid #16a34a;background:#e8f6ee;color:#15803d}
  .ex-ok{display:inline-flex;align-items:center;gap:4px;font-size:12px;font-weight:800;padding:5px 10px;border-radius:10px;white-space:nowrap}
  .ex-ok.y{background:#e8f6ee;color:#15803d}.ex-ok.n{background:#f1f4f8;color:#94a3b8}
  .ex-ep{display:flex;flex-wrap:wrap;gap:4px}.ex-ep span{font-size:11px;font-weight:700;background:#eef3fa;color:#1A4E8A;padding:2px 7px;border-radius:999px}
  .ex-doc{display:flex;align-items:center;gap:6px;font-size:12px;margin-bottom:3px}.ex-doc a{color:#1A4E8A;font-weight:700;text-decoration:none;max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .ex-doc button{border:0;background:none;color:#cbd5e1;cursor:pointer}.ex-doc button:hover{color:#c62828}
  .rdv-c.attente{border:2px solid #fb923c;background:#fff7ed}.rdv-c.fait{background:#f0fdf4;border-color:#bbf7d0}
  .rdv-c .pa{font-size:11px;font-weight:800;color:#c2410c;text-transform:uppercase;letter-spacing:.02em;margin-bottom:2px}
  .vs2-n .pr,.rdv-c .pr{font-size:11.5px;font-style:italic;color:#7c3aed}
  .vs2-t tr.prio td{background:#fef2f2}.rdv-prio{display:inline-block;font-size:10.5px;font-weight:800;color:#fff;background:#dc2626;border-radius:6px;padding:1px 6px;vertical-align:1px}
  .rdv-cp{display:inline-block;font-weight:800;color:#1A4E8A;background:#eef3fa;border-radius:6px;padding:0 5px}
  .rdv-qui{display:flex;flex-wrap:wrap;gap:8px;align-items:center;background:#fff;border:1px solid #e3eaf3;border-radius:16px;padding:10px 12px;margin-bottom:14px}
  .rdv-qui .t{font-size:11px;text-transform:uppercase;color:#7d8a9c;margin-right:4px}
  .rdv-qui button{display:flex;align-items:center;gap:8px;border:1px solid #e3eaf3;background:#fbfcfe;border-radius:10px;padding:5px 9px;cursor:pointer;font:600 12.5px 'Plus Jakarta Sans',Inter,sans-serif}
  .rdv-qui button:hover{border-color:#1A4E8A}
  .rdv-qui .bar{display:flex;width:70px;height:7px;border-radius:9px;background:#eef2f7;overflow:hidden}.rdv-qui .bar i{display:block;height:100%}
  .rdv-qui i.h{background:#16a34a}.rdv-qui i.a{background:#fb923c}
  .rdv-qui .c b{color:#15803d}.rdv-qui .c em{color:#c2410c;font-style:normal;font-weight:800}
  .rdv-qui .lg{font-size:11px;color:#7d8a9c;margin-left:auto;display:flex;gap:5px;align-items:center}.rdv-qui .lg i{display:inline-block;width:10px;height:7px;border-radius:3px}
  .rdv-dt{display:inline-flex;align-items:center;gap:6px;font-size:12px;font-weight:700;color:#64748b;background:#f7f9fc;border:1px solid #e3eaf3;border-radius:10px;padding:3px 8px}
  .rdv-dt input{border:1px solid #e3eaf3;border-radius:7px;padding:3px 5px;font:600 12px 'Plus Jakarta Sans',Inter,sans-serif}
  .rdv-dt button{border:0;background:none;color:#94a3b8;cursor:pointer}
  .rdv-selbar{position:sticky;top:0;z-index:20;display:flex;flex-wrap:wrap;gap:8px;align-items:center;background:#1A4E8A;color:#fff;border-radius:14px;padding:10px 14px;margin-bottom:12px;box-shadow:0 12px 26px -16px rgba(26,78,138,.9)}
  .rdv-selbar b{margin-right:auto}
  .vs2-t tr.sel td{box-shadow:inset 0 0 0 9999px rgba(26,78,138,.07)}
  .rdv-flag{flex:none;font:800 11px 'Plus Jakarta Sans',Inter,sans-serif;border:1px dashed #cbd5e1;background:#fff;color:#94a3b8;border-radius:8px;padding:2px 6px;cursor:pointer;white-space:nowrap}
  .rdv-flag.on{border-style:solid}
  .rdv-niv{display:grid;grid-template-columns:repeat(2,1fr);gap:8px;margin:10px 0}
  .rdv-niv label{display:flex;flex-direction:column;gap:2px;border:2px solid var(--b);background:var(--b);color:var(--c);font-weight:800;border-radius:12px;padding:9px 12px;cursor:pointer}
  .rdv-niv label:has(input:checked){border-color:var(--c)}.rdv-niv input{display:none}.rdv-niv small{font-weight:600;color:#64748b}
  .rdv-mois{position:relative;height:16px;background:#f1f4f8;border-radius:6px;overflow:hidden}.rdv-mois i{position:absolute;left:0;height:8px;border-radius:4px}
  .rdv-mois i.p{top:0;background:#fb923c}.rdv-mois i.h{bottom:0;background:#16a34a}
  .rdv-lg{display:inline-block;width:10px;height:7px;border-radius:3px;margin:0 3px 0 8px}.rdv-lg.p{background:#fb923c}.rdv-lg.h{background:#16a34a}
  .rdv-afaire{border-color:#fdba74;background:#fffbf5}.rdv-afaire .t{color:#c2410c}
  .rdv-cls{position:absolute;z-index:50;top:42px;left:0;min-width:220px;max-height:340px;overflow:auto;background:#fff;border:1px solid #e3eaf3;border-radius:14px;padding:10px;box-shadow:0 20px 40px -20px rgba(14,26,43,.45)}
  .rdv-cls label{display:flex;gap:8px;align-items:center;font-size:13px;padding:4px 6px;border-radius:8px;cursor:pointer}.rdv-cls label:hover{background:#f3f7fc}
  .rdv-ml{border:1px solid #e3eaf3;border-radius:12px;background:#fbfcfe}
  .rdv-ml .h{display:flex;gap:10px;align-items:center;padding:8px 10px;font-size:13px;cursor:pointer}.rdv-ml .h span{flex:1;min-width:0}.rdv-ml small{color:#64748b;font-size:11.5px}
  .rdv-ml .b{border-top:1px dashed #e3eaf3;padding:10px 12px;font-size:13px;line-height:1.5;background:#fff;border-radius:0 0 12px 12px}
  .vs2-mv{background:#fff;border:1px solid #e3eaf3;border-radius:14px;padding:10px 14px;margin-bottom:8px}.vs2-mv summary{cursor:pointer;font-size:13.5px}
  .ex-num{width:120px;border:1px solid #e3eaf3;border-radius:8px;padding:5px 8px;font:600 12px 'Plus Jakarta Sans',Inter,sans-serif}
  `;
  document.head.appendChild(st);
}
const dl = (name, rows) => csvDownload(name, rows);
const dipOrdreTri = d => { const i = DIP_ORDRE.indexOf(d); return i === -1 ? 99 : i; };
// Regroupe une liste par diplôme puis par classe → [[diplome, [[classe, items]]]]
function parDipClasse(l, dipF = x => x.diplome, clF = x => x.classe) {
  const g = {};
  l.forEach(x => { const d = dipF(x) || 'Autre', c = clF(x) || 'Classe non renseignée'; ((g[d] = g[d] || {})[c] = g[d][c] || []).push(x); });
  return Object.keys(g).sort((a, b) => dipOrdreTri(a) - dipOrdreTri(b) || a.localeCompare(b)).map(d => [d, Object.keys(g[d]).sort((a, b) => a.localeCompare(b, 'fr', { numeric: true })).map(c => [c, g[d][c].sort((a, b) => String(a.nom).localeCompare(String(b.nom), 'fr'))])]);
}

// ═══════════════════════════ SUIVI DES RDV EN ENTREPRISE ═══════════════════════════
const R = { data: null, annee: '', onglet: 'suivi', classes: new Set(), analyse: null, sel: new Set(), visibles: [] };
async function rdvInit() {
  injectCss(); injectCss2();
  const sec = $('section-rdv');
  if (!sec.dataset.ready) {
    sec.dataset.ready = '1'; sec.classList.add('vx');
    sec.innerHTML = `<div class="vx-hd"><div><h1>Suivi des RDV en entreprise</h1><p>Chaque étudiant en alternance reçoit la visite d'un membre de l'équipe sur son lieu de travail — historique complet par étudiant et par entreprise</p></div>
      <div style="display:flex;gap:8px;flex-wrap:wrap">
        <button class="vx-btn" id="rdvParamBtn" style="display:none" onclick="VX.rdvParams()">${svg('edit', 15)} Paramètres</button>
        <button class="vx-btn ora" id="rdvMailBtn" style="display:none" onclick="VX.rdvMails()">${svg('mail', 15)} Envoyer le lien de RDV</button>
        <button class="vx-btn" id="rdvCharlBtn" style="display:none" onclick="VX.rdvCharl()">${svg('down', 15)} Reprendre depuis Charlemagne</button>
        <button class="vx-btn" id="rdvHarmBtn" style="display:none" onclick="VX.rdvHarmoniser()">${svg('user', 15)} Harmoniser les noms</button>
        <button class="vx-btn" onclick="VX.rdvExport()">${svg('down', 15)} Exporter Excel</button>
        <button class="vx-btn" id="rdvImpBtn" style="display:none" onclick="VX.rdvImport()" title="Réimporter un export retravaillé dans Excel (aperçu avant enregistrement)">${svg('upload', 15)} Importer Excel</button></div></div>
      <div class="vx-tabs" style="margin-bottom:14px"><button id="rdvTabS" class="on" onclick="VX.rdvOnglet('suivi')">${svg('list', 15)} Suivi des RDV</button><button id="rdvTabA" onclick="VX.rdvOnglet('analyse')">${svg('grid', 15)} Analyse des visites</button></div>
      <div class="vx-kpis" id="rdvKpis"></div>
      <div class="vx-bar">
        <select class="vx-sel" id="rdvAnnee" onchange="VX.rdvLoad(this.value)"></select>
        <div class="vx-search">${svg('search', 15)}<input id="rdvQ" placeholder="Étudiant, entreprise, visiteur…" oninput="VX.rdvRender()"></div>
        <select class="vx-sel" id="rdvDip" onchange="VX.rdvRender()"></select>
        <div style="position:relative"><button class="vx-btn sm" id="rdvClsBtn" onclick="VX.rdvClsPanel()">${svg('user', 13)} Classes : toutes ▾</button><div id="rdvClsPanel" class="rdv-cls" style="display:none"></div></div>
        <select class="vx-sel" id="rdvResp" onchange="VX.rdvRender()"></select>
        <span class="rdv-dt">${svg('clock', 13)} RDV pris du <input type="date" id="rdvDu" onchange="VX.rdvRender()"> au <input type="date" id="rdvAu" onchange="VX.rdvRender()"><button type="button" title="Effacer les dates" onclick="VX.rdvDates()">✕</button></span>
        <select class="vx-sel" id="rdvVue" onchange="VX.rdvRender()"><option value="alt">Alternants</option><option value="tous">Tous les inscrits</option><option value="sansresp">⚠ Alternants sans responsable visite</option><option value="prio">🚩 Drapeaux rouges (prioritaires)</option><option value="drap">⚑ Tous les drapeaux</option><option value="attente">⏳ RDV pris, non honorés</option><option value="venir">RDV à venir</option><option value="suivre">À suivre (date passée)</option><option value="sans">Sans aucun RDV</option><option value="nonenvoye">Lien de RDV non envoyé</option></select>
        <select class="vx-sel" id="rdvTri" onchange="VX.rdvRender()"><option value="classe">Grouper par diplôme et classe</option><option value="resp">Grouper par responsable visite</option><option value="cp">Trier par code postal du lieu</option><option value="nom">Trier par étudiant</option></select>
      </div>
      <div id="rdvSelBar"></div>
      <div id="rdvQui"></div>
      <div id="rdvBody"><div class="vx-empty"><h3>Chargement…</h3></div></div>`;
    document.addEventListener('mousedown', ev => { const p = $('rdvClsPanel'); if (p && p.style.display !== 'none' && !ev.target.closest('#rdvClsPanel,#rdvClsBtn')) p.style.display = 'none'; });
  }
  rdvLoad();
}
async function rdvLoad(annee) {
  if (annee) { R.annee = annee; R.analyse = null; }
  const r = await api('getRdv', R.annee ? { annee: R.annee } : {}, { forceRefresh: true });
  if (!r.success) { $('rdvBody').innerHTML = `<div class="vx-empty"><h3>Chargement impossible</h3><p>${esc(errTxt(r.message))}</p></div>`; return; }
  R.data = r; R.annee = r.annee; r.envois = r.envois || {};
  ['rdvCharlBtn', 'rdvParamBtn', 'rdvMailBtn', 'rdvHarmBtn', 'rdvImpBtn'].forEach(id => $(id).style.display = r.peutSuivre ? '' : 'none');
  $('rdvAnnee').innerHTML = r.annees.map(a => `<option ${a === r.annee ? 'selected' : ''}>${a}</option>`).join('');
  const dips = [...new Set(r.etudiants.map(e => e.diplome))].sort((a, b) => dipOrdreTri(a) - dipOrdreTri(b));
  let v = $('rdvDip').value;
  $('rdvDip').innerHTML = '<option value="">Tous les diplômes</option>' + dips.map(d => `<option>${esc(d)}</option>`).join(''); $('rdvDip').value = v;
  const resps = [...new Set(r.etudiants.map(e => e.referent).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'fr'));
  v = $('rdvResp').value;
  $('rdvResp').innerHTML = '<option value="">Tous les responsables</option>' + resps.map(x => `<option>${esc(x)}</option>`).join('') + '<option value="__aucun">Sans responsable</option>'; $('rdvResp').value = v;
  rdvRender();
}
const rdvAuj = () => todayISO();
const rdvAsuivre = v => ['Planifié', 'Reporté'].includes(v.statut) && v.date && v.date < rdvAuj();
const rdvAttente = v => ['Planifié', 'Reporté'].includes(v.statut) && !!(v.datePris || v.date);   // RDV pris, pas encore honoré
// Drapeaux persistants (RDV_DRAPEAUX) : rouge posé automatiquement quand un responsable est attribué sans
// aucun historique ; il ne disparaît plus tout seul, on le lève à la main.
const DRAP = { Rouge: ['#dc2626', '#fef2f2'], Orange: ['#ea580c', '#fff7ed'], Jaune: ['#ca8a04', '#fefce8'] };
const rdvPrio = e => e.drapeau === 'Rouge';
const joursDepuis = iso => iso ? Math.round((new Date(todayISO() + 'T12:00:00') - new Date(iso + 'T12:00:00')) / 86400000) : null;
function rdvClsPanel() {
  const p = $('rdvClsPanel');
  if (p.style.display !== 'none') { p.style.display = 'none'; return; }
  const cls = [...new Set(R.data.etudiants.filter(e => !$('rdvDip').value || e.diplome === $('rdvDip').value).map(e => e.classe || 'Classe non renseignée'))].sort((a, b) => a.localeCompare(b, 'fr', { numeric: true }));
  p.innerHTML = `<div style="margin-bottom:8px"><button class="vx-btn sm" onclick="VX.rdvCls('__tout')">Toutes les classes</button></div>`
    + cls.map(c => `<label><input type="checkbox" ${R.classes.has(c) ? 'checked' : ''} onchange="VX.rdvCls('${jsq(c)}',this.checked)"> ${esc(c)}</label>`).join('');
  p.style.display = 'block';
}
function rdvCls(c, on) {
  if (c === '__tout') { R.classes.clear(); $('rdvClsPanel').style.display = 'none'; }
  else if (on) R.classes.add(c); else R.classes.delete(c);
  $('rdvClsBtn').innerHTML = svg('user', 13) + ' Classes : ' + (R.classes.size ? R.classes.size + ' sélectionnée' + (R.classes.size > 1 ? 's' : '') : 'toutes') + ' ▾';
  rdvRender();
}
function rdvOnglet(o) {
  R.onglet = o;
  $('rdvTabS').classList.toggle('on', o === 'suivi'); $('rdvTabA').classList.toggle('on', o === 'analyse');
  $('rdvVue').style.display = $('rdvTri').style.display = o === 'suivi' ? '' : 'none';
  rdvRender();
}
// Étudiants correspondant aux filtres (diplôme, classes, responsable, dates de RDV pris, recherche) — commun au suivi, à l'analyse
function rdvFiltre(avecVue = true) {
  const D = R.data;
  const q = norm($('rdvQ').value), dip = $('rdvDip').value, resp = $('rdvResp').value, vue = $('rdvVue').value, du = $('rdvDu').value, au = $('rdvAu').value;
  const vAnnee = D.visites.filter(v => v.annee === D.annee);
  const parEtu = {}; vAnnee.forEach(v => (parEtu[v.etuId] = parEtu[v.etuId] || []).push(v));
  Object.values(parEtu).forEach(l => l.sort((a, b) => (a.datePris || a.creeLe || '').localeCompare(b.datePris || b.creeLe || '')));
  let L = D.etudiants.filter(e => (!dip || e.diplome === dip) && (!R.classes.size || R.classes.has(e.classe || 'Classe non renseignée'))
    && (!resp || (resp === '__aucun' ? !e.referent : e.referent === resp))
    && (!q || norm([e.nom, e.classe, e.entreprise, e.referent, e.drapeauCom, ...(parEtu[e.id] || []).map(v => v.visiteur + ' ' + v.entreprise)].join(' ')).includes(q)));
  if (du || au) L = L.filter(e => (parEtu[e.id] || []).some(v => v.datePris && (!du || v.datePris >= du) && (!au || v.datePris <= au)));
  if (avecVue) {
    if (vue === 'alt') L = L.filter(e => e.alternant || parEtu[e.id]);
    if (vue === 'sansresp') L = L.filter(e => e.alternant && !e.referent);
    if (vue === 'prio') L = L.filter(rdvPrio);
    if (vue === 'drap') L = L.filter(e => e.drapeau);
    if (vue === 'sans') L = L.filter(e => e.alternant && !(parEtu[e.id] || []).some(v => v.statut !== 'Annulé'));
    if (vue === 'suivre') L = L.filter(e => (parEtu[e.id] || []).some(rdvAsuivre));
    if (vue === 'venir') L = L.filter(e => (parEtu[e.id] || []).some(v => ['Planifié', 'Reporté'].includes(v.statut) && v.date >= rdvAuj()));
    if (vue === 'nonenvoye') L = L.filter(e => e.alternant && !D.envois[e.id]);
    if (vue === 'attente') L = L.filter(e => (parEtu[e.id] || []).some(rdvAttente));
  }
  return { L, parEtu, vAnnee };
}
function rdvRender() {
  const D = R.data; if (!D) return;
  if (R.onglet === 'analyse') { $('rdvSelBar').innerHTML = ''; rdvAnalyse(); return; }
  const { L, parEtu, vAnnee } = rdvFiltre(true);
  R.visibles = L.map(e => e.id);
  const histo = {}; D.visites.forEach(v => { histo[v.etuId] = (histo[v.etuId] || 0) + 1; });
  const alts = D.etudiants.filter(e => e.alternant);
  const honores = vAnnee.filter(v => v.statut === 'Honoré').length, venir = vAnnee.filter(v => ['Planifié', 'Reporté'].includes(v.statut) && v.date >= rdvAuj()).length;
  const retard = vAnnee.filter(rdvAsuivre).length;
  const attente = vAnnee.filter(rdvAttente).length, prio = D.etudiants.filter(rdvPrio).length, sansResp = alts.filter(e => !e.referent).length;
  $('rdvQui').innerHTML = rdvAFaire(L, parEtu) + rdvParVisiteur(L, parEtu);
  $('rdvKpis').innerHTML = `
    <div class="vx-kpi"><span>Alternants inscrits</span><b>${alts.length}</b><small>${D.annee}</small></div>
    <div class="vx-kpi" style="cursor:pointer;border-color:${sansResp ? '#fca5a5' : '#e3eaf3'}" onclick="VX.rdvVue('sansresp')"><span>⚠ Sans responsable visite</span><b style="color:${sansResp ? '#c62828' : '#15803d'}">${sansResp}</b><small>alternant(s) sans visiteur désigné</small></div>
    <div class="vx-kpi" style="cursor:pointer;border-color:#fca5a5" onclick="VX.rdvVue('prio')"><span>🚩 Drapeaux rouges</span><b style="color:${prio ? '#c62828' : 'inherit'}">${prio}</b><small>à lever à la main</small></div>
    <div class="vx-kpi" style="cursor:pointer;border-color:#fdba74" onclick="VX.rdvVue('attente')"><span>⏳ RDV pris, non honorés</span><b style="color:#c2410c">${attente}</b><small>${venir} à venir · ${retard} date passée</small></div>
    <div class="vx-kpi"><span>RDV honorés</span><b style="color:#15803d">${honores}</b><small>${vAnnee.length} RDV dans l'année</small></div>`;
  rdvSelBar();
  if (!L.length) { $('rdvBody').innerHTML = `<div class="vx-empty"><h3>Aucun étudiant pour ce filtre</h3></div>`; return; }
  const maxR = Math.max(1, ...L.map(e => (parEtu[e.id] || []).length));
  const cell = v => {
    const st = v.statut, ret = rdvAsuivre(v);
    const att = rdvAttente(v);
    return `<div class="rdv-c ${att ? 'attente' : ''} ${st === 'Honoré' ? 'fait' : ''}" onclick="VX.rdvEdit('${jsq(v.id)}')">
      ${att ? `<div class="pa">⏳ RDV pris — ${ret ? 'date passée, à confirmer' : 'à venir'}</div>` : ''}
      <div class="d">${v.datePris ? 'Pris le ' + fmtD(v.datePris) : 'Date à fixer'}</div>
      ${v.date && v.date !== v.datePris ? `<div class="rp">→ reporté au ${fmtD(v.date)}</div>` : ''}
      <div class="s">${v.visiteur ? esc(v.visiteur) : '<i>visiteur ?</i>'}${v.entreprise ? ' · ' + esc(v.entreprise) : ''}</div>
      ${v.precision ? `<div class="pr">${esc(v.precision)}</div>` : ''}
      <span class="st-b st-${st.replace(/\s/g, '')} ${ret ? 'st-retard' : ''}" title="${ret ? 'Date passée : indiquez si le RDV a été honoré' : ''}">${esc(st)}${ret ? ' · à suivre' : ''}</span>
      ${D.peutSuivre && !['Honoré', 'Annulé'].includes(st) && v.date ? `<div class="rdv-b" onclick="event.stopPropagation()">
        <button class="h" onclick="VX.rdvStatut('${jsq(v.id)}','Honoré')" title="RDV honoré">✓ Honoré</button>
        <button class="r" onclick="VX.rdvStatut('${jsq(v.id)}','Reporté')" title="Reporté : nouvelle date">↻ Reporté</button>
        <button class="a" onclick="VX.rdvStatut('${jsq(v.id)}','Annulé')" title="Annulé">✕ Annulé</button>
        <button class="o" onclick="VX.rdvStatut('${jsq(v.id)}','Autre')" title="Autre (préciser)">… Autre</button></div>` : ''}
    </div>`;
  };
  const drap = e => {
    const c = DRAP[e.drapeau];
    const tip = e.drapeau ? `Drapeau ${e.drapeau.toLowerCase()} posé le ${fmtD(e.drapeauLe)} par ${e.drapeauPar}${e.drapeauCom ? ' — ' + e.drapeauCom : ''} (cliquer pour modifier ou lever)` : (e.drapeauLeve ? 'Drapeau levé le ' + e.drapeauLeve + ' — cliquer pour en poser un' : 'Poser un drapeau de priorité / un commentaire');
    return `<button type="button" class="rdv-flag ${e.drapeau ? 'on' : ''}" style="${c ? `color:${c[0]};background:${c[1]};border-color:${c[0]}` : ''}" title="${esc(tip)}" onclick="VX.rdvDrapeau(['${jsq(e.id)}'])">${e.drapeau ? '🚩 ' + esc(e.drapeau) : '⚑'}</button>`;
  };
  const ligne = e => {
    const vs = parEtu[e.id] || [], env = D.envois[e.id], c = DRAP[e.drapeau], j = joursDepuis(e.referentDepuis);
    return `<tr class="${R.sel.has(e.id) ? 'sel' : ''}" style="${c ? `background:${c[1]}` : ''}"><td style="width:30px"><input type="checkbox" ${R.sel.has(e.id) ? 'checked' : ''} onchange="VX.rdvSel('${jsq(e.id)}',this.checked)"></td>
      <td class="vs2-n"><div style="display:flex;gap:6px;align-items:flex-start">${drap(e)}<div><b>${esc(e.nom)}</b><small>${esc(e.classe)} · ${esc(e.typo)}${e.alternant ? '' : ' · pas en alternance'}</small></div></div>${e.drapeauCom ? `<div class="pr" style="color:#991b1b">💬 ${esc(e.drapeauCom)}</div>` : ''}${histo[e.id] ? `<div><button class="vx-link" onclick="VX.rdvHisto('${jsq(e.id)}')">${svg('clock', 12)} Historique (${histo[e.id]})</button></div>` : ''}</td>
      <td class="vs2-n">${e.entreprise ? `<b>${esc(e.entreprise)}</b><small>${e.cp ? `<span class="rdv-cp">${esc(e.cp)}</span> ` : ''}${esc(e.lieu || 'adresse non renseignée')}</small>` : '<small style="color:#c2410c;font-weight:700">Aucune entreprise dans le CRM</small>'}</td>
      <td class="vs2-n"><button class="vx-link" style="margin-top:0" onclick="VX.rdvAffecter(['${jsq(e.id)}'])">${e.referent ? esc(e.referent) : '<i style="color:#c62828">⚠ à désigner</i>'}</button>${e.referentSource === 'Charlemagne' ? '<small> (Charlemagne)</small>' : ''}${e.referentPrecision ? `<div class="pr">${esc(e.referentPrecision)}</div>` : ''}
        ${e.referentDepuis ? `<small style="display:block" title="Date d'attribution du responsable">attribué le ${fmtD(e.referentDepuis)}${j !== null ? ` · <b style="color:${j > 30 ? '#c62828' : j > 14 ? '#c2410c' : '#64748b'}">il y a ${j} j</b>` : ''}</small>` : ''}
        ${env ? `<small style="display:block;color:#0f766e;font-weight:700" title="${esc(env.map(x => fmtD(x.le) + ' → ' + x.dest).join('\n'))}">${svg('mail', 11)} lien envoyé le ${fmtD(env[env.length - 1].le)}${env.length > 1 ? ' (' + env.length + '×)' : ''}</small>` : ''}</td>
      ${Array.from({ length: maxR }, (_, i) => `<td>${vs[i] ? cell(vs[i]) : ''}</td>`).join('')}
      <td><button class="vx-btn sm" onclick="VX.rdvNew('${jsq(e.id)}')">${svg('plus', 13, 2.4)} RDV</button></td></tr>`;
  };
  const tousSel = L.length && L.every(e => R.sel.has(e.id));
  const head = `<tr><th><input type="checkbox" title="Tout sélectionner (liste affichée)" ${tousSel ? 'checked' : ''} onchange="VX.rdvSelTout(this.checked)"></th><th>Étudiant</th><th>Entreprise (lieu du RDV)</th><th>Responsable visite</th>${Array.from({ length: maxR }, (_, i) => `<th>Date RDV pris ${i + 1}</th>`).join('')}<th></th></tr>`;
  const tri = $('rdvTri').value;
  if (tri === 'resp') {
    const g = {}; L.forEach(e => (g[e.referent || '— Sans responsable —'] = g[e.referent || '— Sans responsable —'] || []).push(e));
    $('rdvBody').innerHTML = Object.keys(g).sort((a, b) => a.startsWith('—') - b.startsWith('—') || a.localeCompare(b, 'fr')).map(r => {
      const l = g[r].sort((a, b) => (a.classe + a.nom).localeCompare(b.classe + b.nom, 'fr', { numeric: true }));
      const hon = l.filter(e => (parEtu[e.id] || []).some(v => v.statut === 'Honoré')).length;
      return `<div class="vx-dip"><span class="t">${esc(r)}</span><span class="n">${l.length} étudiant(s)</span><span class="n">${hon} visité(s)</span></div>
        <div style="overflow-x:auto;margin-top:12px"><table class="vs2-t"><thead>${head}</thead><tbody>${l.map(ligne).join('')}</tbody></table></div>`;
    }).join('');
  } else if (tri === 'cp') {
    const g = {}; L.forEach(e => { const k = e.cp ? e.cp.slice(0, 2) : '—'; (g[k] = g[k] || []).push(e); });
    $('rdvBody').innerHTML = Object.keys(g).sort((a, b) => (a === '—') - (b === '—') || a.localeCompare(b)).map(dep => {
      const l = g[dep].sort((a, b) => (a.cp || '99999').localeCompare(b.cp || '99999') || (a.ville || '').localeCompare(b.ville || '', 'fr') || a.nom.localeCompare(b.nom, 'fr'));
      return `<div class="vx-dip"><span class="t">${dep === '—' ? 'Lieu sans code postal' : 'Département ' + esc(dep)}</span><span class="n">${l.length} étudiant(s)</span></div>
        <div style="overflow-x:auto;margin-top:12px"><table class="vs2-t"><thead>${head}</thead><tbody>${l.map(ligne).join('')}</tbody></table></div>`;
    }).join('');
  } else if (tri === 'nom') {
    $('rdvBody').innerHTML = `<div style="overflow-x:auto"><table class="vs2-t"><thead>${head}</thead><tbody>${L.slice().sort((a, b) => a.nom.localeCompare(b.nom, 'fr')).map(ligne).join('')}</tbody></table></div>`;
  } else {
    $('rdvBody').innerHTML = parDipClasse(L).map(([d, classes]) => `<div class="vx-dip"><span class="t">${esc(d === 'Autre' ? 'Autres formations' : d)}</span><span class="n">${classes.reduce((s, c) => s + c[1].length, 0)} étudiant(s)</span></div>
      <div style="overflow-x:auto;margin-top:12px"><table class="vs2-t"><thead>${head}</thead><tbody>${classes.map(([c, l]) => `<tr class="gr"><td colspan="${maxR + 5}"><b>${esc(c)}</b><span>${l.length} étudiant(s)</span><span>${l.filter(e => !e.referent && e.alternant).length ? `<b style="color:#c62828">${l.filter(e => !e.referent && e.alternant).length} sans responsable</b>` : ''}</span></td></tr>` + l.map(ligne).join('')).join('')}</tbody></table></div>`).join('');
  }
}
// ── Sélection multiple : affectation en masse d'un responsable, drapeaux ──
function rdvSelBar() {
  const n = R.sel.size;
  $('rdvSelBar').innerHTML = n ? `<div class="rdv-selbar"><b>${n} étudiant(s) sélectionné(s)</b>
    <button class="vx-btn sm pri" onclick="VX.rdvAffecter()">${svg('user', 13)} Affecter un responsable visite</button>
    <button class="vx-btn sm" onclick="VX.rdvDrapeau()">⚑ Drapeau / commentaire</button>
    <button class="vx-btn sm" onclick="VX.rdvSelTout(false, true)">Vider la sélection</button></div>` : '';
}
function rdvSel(id, on) { if (on) R.sel.add(id); else R.sel.delete(id); rdvSelBar(); const tr = event && event.target && event.target.closest('tr'); if (tr) tr.classList.toggle('sel', on); }
function rdvSelTout(on, tout) { if (tout) R.sel.clear(); else (R.visibles || []).forEach(id => on ? R.sel.add(id) : R.sel.delete(id)); rdvRender(); }
function rdvAffecter(ids) {
  ids = ids || [...R.sel];
  const L = R.data.etudiants.filter(e => ids.includes(e.id)); if (!L.length) return;
  const e = L.length === 1 ? L[0] : null;
  const dejaR = L.filter(x => x.referent).length;
  const dlg = modal(`<h2>Responsable des visites${e ? ' — ' + esc(e.nom) : ` — ${L.length} étudiants`}</h2>
    <p class="sub">Membre de l'équipe chargé des RDV en ${esc(R.annee)}. La date d'attribution est enregistrée.${!e && dejaR ? ` <b style="color:#c2410c">${dejaR} étudiant(s) ont déjà un responsable : il sera remplacé.</b>` : ''}</p>
    ${!e ? `<div class="vx-note" style="display:block;max-height:120px;overflow:auto;margin-bottom:8px">${L.map(x => esc(x.nom) + (x.referent ? ` <i>(${esc(x.referent)})</i>` : '')).join(', ')}</div>` : ''}
    <select class="in" id="rfVis">${rdvEquipeOptions(e ? (e.referentEmail || e.referent) : '')}</select>
    <input class="in" id="rfLibre" placeholder="Prénom NOM" style="display:none;margin-top:6px">
    <label class="fl">Précision (zone libre)</label><input class="in" id="rfPrec" placeholder="Ex : à sa demande" value="${esc(e ? e.referentPrecision || '' : '')}">
    <div class="mft"><button class="vx-btn" onclick="this.closest('.vx-modal').remove()">Annuler</button><button class="vx-btn pri" id="rfOk">Affecter${e ? '' : ' aux ' + L.length + ' étudiants'}</button></div>`, 520);
  $('rfVis').onchange = () => { $('rfLibre').style.display = $('rfVis').value === '__autre' ? '' : 'none'; };
  $('rfOk').onclick = async () => {
    const who = rdvVisiteurChoisi('rfVis', 'rfLibre');
    if (!who.visiteur && !confirm('Aucun responsable choisi : retirer le responsable ?')) return;
    const r = await api('saveReferentRdv', { annee: R.annee, etuIds: L.map(x => x.id), nom: who.visiteur, email: who.visiteurEmail, precision: $('rfPrec').value.trim() });
    if (!r.success) { toast('Erreur : ' + errTxt(r.message), false); return; }
    dlg.remove(); R.sel.clear(); toast('Responsable affecté à ' + L.length + ' étudiant(s) ✓'); rdvLoad();
  };
}
function rdvDrapeau(ids) {
  ids = ids || [...R.sel];
  const L = R.data.etudiants.filter(e => ids.includes(e.id)); if (!L.length) return;
  const e = L.length === 1 ? L[0] : null;
  const dlg = modal(`<h2>Drapeau de priorité${e ? ' — ' + esc(e.nom) : ` — ${L.length} étudiants`}</h2>
    <p class="sub">Les drapeaux restent affichés jusqu'à ce que vous les leviez. ${e && e.drapeau ? `Actuel : <b>${esc(e.drapeau)}</b>, posé le ${fmtD(e.drapeauLe)} par ${esc(e.drapeauPar)}.` : e && e.drapeauLeve ? `Dernier drapeau levé le ${esc(e.drapeauLeve)}.` : ''}</p>
    <div class="rdv-niv">${['Rouge', 'Orange', 'Jaune'].map(n => `<label style="--c:${DRAP[n][0]};--b:${DRAP[n][1]}"><input type="radio" name="dgN" value="${n}" ${(e && e.drapeau === n) || (!e && n === 'Rouge') || (e && !e.drapeau && n === 'Rouge') ? 'checked' : ''}> 🚩 ${n}<small>${n === 'Rouge' ? 'urgent' : n === 'Orange' ? 'important' : 'à surveiller'}</small></label>`).join('')}
      <label style="--c:#15803d;--b:#f0fdf4"><input type="radio" name="dgN" value="lever"> ✓ Lever le drapeau<small>plus besoin</small></label></div>
    <label class="fl">Commentaire</label><textarea class="in" id="dgCom" style="min-height:70px" placeholder="Pourquoi cette priorité, ce qui reste à faire…">${esc(e ? e.drapeauCom || '' : '')}</textarea>
    <div class="mft"><button class="vx-btn" onclick="this.closest('.vx-modal').remove()">Annuler</button><button class="vx-btn pri" id="dgOk">Enregistrer</button></div>`, 560);
  $('dgOk').onclick = async () => {
    const n = dlg.querySelector('[name="dgN"]:checked').value;
    const r = await api('saveDrapeauRdv', { annee: R.annee, etuIds: L.map(x => x.id), niveau: n === 'lever' ? '' : n, commentaire: $('dgCom').value.trim() });
    if (!r.success) { toast('Erreur : ' + errTxt(r.message), false); return; }
    dlg.remove(); toast(n === 'lever' ? 'Drapeau levé ✓' : 'Drapeau ' + n.toLowerCase() + ' enregistré ✓'); rdvLoad();
  };
}

// ── « Qui a combien de visites à faire » : alternants attribués à chaque responsable et pas encore visités cette année ──
function rdvAFaire(L, parEtu) {
  const g = {};
  L.filter(e => e.alternant).forEach(e => {
    const vs = parEtu[e.id] || [];
    if (vs.some(v => v.statut === 'Honoré')) return;   // déjà visité cette année
    const k = e.referent || '— Sans responsable —';
    const o = g[k] = g[k] || { total: 0, pris: 0, sans: 0 };
    o.total++; if (vs.some(rdvAttente)) o.pris++; else o.sans++;
  });
  const l = Object.entries(g).sort((a, b) => a[0].startsWith('—') - b[0].startsWith('—') || b[1].total - a[1].total);
  if (!l.length) return `<div class="rdv-qui rdv-afaire"><b class="t">Visites à faire</b><span style="color:#15803d;font-weight:800">✅ Tous les alternants affichés ont été visités cette année</span></div>`;
  const tot = l.reduce((s, [, o]) => s + o.total, 0);
  return `<div class="rdv-qui rdv-afaire"><b class="t">Visites à faire · ${tot}</b>${l.map(([k, o]) => `<button type="button" onclick="${k.startsWith('—') ? "VX.rdvVue('sansresp')" : `VX.rdvResp('${jsq(k)}')`}" title="${o.total} alternant(s) pas encore visité(s) cette année : ${o.pris} avec un RDV pris, ${o.sans} sans RDV">
    <span class="n" style="${k.startsWith('—') ? 'color:#c62828' : ''}">${esc(k)}</span><span class="c"><b style="color:#0e1a2b">${o.total}</b> à faire${o.pris ? ` · <em>${o.pris} RDV pris</em>` : ''}${o.sans ? ` · <span style="color:#c62828;font-weight:800">${o.sans} sans RDV</span>` : ''}</span></button>`).join('')}</div>`;
}
// ── Résumé « qui a combien de visites » (périmètre = filtres actuels) ──
function rdvParVisiteur(L, parEtu) {
  const g = {};
  L.forEach(e => (parEtu[e.id] || []).forEach(v => {
    const k = v.visiteur || '— visiteur non renseigné —';
    const o = g[k] = g[k] || { hon: 0, att: 0, autres: 0 };
    if (v.statut === 'Honoré') o.hon++; else if (['Planifié', 'Reporté'].includes(v.statut)) o.att++; else o.autres++;
  }));
  const l = Object.entries(g).sort((a, b) => (b[1].hon + b[1].att) - (a[1].hon + a[1].att));
  if (!l.length) return '';
  const max = Math.max(...l.map(([, o]) => o.hon + o.att), 1);
  return `<div class="rdv-qui"><b class="t">Qui a combien de visites</b>${l.map(([k, o]) => `<button type="button" onclick="VX.rdvQ('${jsq(k.startsWith('—') ? '' : k)}')" title="${o.hon} honorée(s), ${o.att} prise(s) non honorée(s)${o.autres ? ', ' + o.autres + ' annulée(s)/autre' : ''}">
    <span class="n">${esc(k)}</span><span class="bar"><i class="h" style="width:${o.hon / max * 100}%"></i><i class="a" style="width:${o.att / max * 100}%"></i></span><span class="c"><b>${o.hon}</b>${o.att ? ` + <em>${o.att}</em>` : ''}</span></button>`).join('')}
    <span class="lg"><i class="h"></i> honorées <i class="a"></i> prises, non honorées</span></div>`;
}
// ── Harmonisation des noms de visiteurs ──
async function rdvHarmoniser() {
  const r = await api('harmoniserVisiteurs', {});
  if (!r.success) { toast('Erreur : ' + errTxt(r.message), false); return; }
  const G = r.groupes, diff = G.filter(g => g.brut !== g.nom || g.precision);
  const dlg = modal(`<h2>Harmoniser les noms des visiteurs</h2>
    <p class="sub">${G.length} écriture(s) différente(s) trouvée(s) dans les RDV et les responsables. Le texte libre (« à sa demande », parenthèses…) est rangé dans la zone <b>Précision</b>. Pour un nom non reconnu, choisissez le membre de l'équipe correspondant : il sera retenu pour les prochains imports.</p>
    <div style="max-height:56vh;overflow:auto"><table class="vs2-t"><thead><tr><th>Écrit</th><th>Devient</th><th>Précision</th><th>Fois</th><th>Membre de l'équipe</th></tr></thead><tbody>
    ${G.map((g, i) => `<tr><td>${esc(g.brut)}</td><td><b>${esc(g.nom)}</b>${g.reconnu ? ' <span class="st-b c-ok">équipe</span>' : g.ambigu ? ' <span class="st-b c-wa">plusieurs possibles</span>' : ' <span class="st-b c-ne">non reconnu</span>'}</td><td><i>${esc(g.precision || '')}</i></td><td>${g.n}</td>
      <td>${g.reconnu ? '' : `<select class="in" style="margin:0;min-width:180px" data-alias="${i}"><option value="">— garder « ${esc(g.nom)} » —</option>${r.equipe.map(q => `<option value="${esc(q.email)}">${esc(q.nom)}</option>`).join('')}</select>`}</td></tr>`).join('')}
    </tbody></table></div>
    <div class="mft"><button class="vx-btn" onclick="this.closest('.vx-modal').remove()">Annuler</button><button class="vx-btn pri" id="hvOk">Harmoniser (${diff.length} écriture(s) à corriger)</button></div>`, 960);
  $('hvOk').onclick = async () => {
    const alias = {}; dlg.querySelectorAll('[data-alias]').forEach(s => { if (s.value) alias[_rdvNomSeul(G[+s.dataset.alias].brut)] = s.value; });
    const o = await api('harmoniserVisiteurs', { appliquer: true, alias });
    if (!o.success) { toast('Erreur : ' + errTxt(o.message), false); return; }
    dlg.remove(); toast(`Noms harmonisés : ${o.modifies.rdv} RDV, ${o.modifies.responsables} responsable(s) ✓`); rdvLoad();
  };
}
const _rdvNomSeul = s => String(s).replace(/\(([^)]*)\)/g, ' ').replace(/\s*(?:[-–:,;]\s*)?(?:^|\s)(?:à|a)\s+(?:sa|la|leur|ma|notre|votre)\s+demande\b.*$/i, '').replace(/\s+[-–:]\s+.*$/, '').replace(/\s+/g, ' ').trim();

// ── Import d'un fichier Excel retravaillé (export « Suivi des RDV ») ──
const RDV_COLS_IMPORT = { idRdv: ['id rdv'], etuId: ['id etudiant'], responsable: ['responsable visite'], precisionResp: ['precision responsable'],
  datePris: ['date rdv pris'], date: ['date rdv apres report', 'date rdv'], visiteur: ['visiteur'], precision: ['precision'], statut: ['statut'], commentaire: ['commentaire'] };
function rdvImport() {
  let inp = $('rdvImpFile');
  if (!inp) { inp = document.createElement('input'); inp.type = 'file'; inp.id = 'rdvImpFile'; inp.accept = '.xlsx,.xls,.csv'; inp.hidden = true; document.body.appendChild(inp); inp.onchange = ev => { const f = ev.target.files[0]; ev.target.value = ''; if (f) rdvImportLire(f); }; }
  inp.click();
}
async function rdvImportLire(f) {
  if (typeof XLSX === 'undefined') { toast('Lecteur Excel indisponible, rechargez la page', false); return; }
  try {
    const wb = XLSX.read(await f.arrayBuffer(), { type: 'array', cellDates: true });
    const iso = v => { if (v instanceof Date && !isNaN(v)) return v.getFullYear() + '-' + String(v.getMonth() + 1).padStart(2, '0') + '-' + String(v.getDate()).padStart(2, '0'); return String(v ?? '').trim(); };
    let lignes = null;
    for (const nf of wb.SheetNames) {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[nf], { header: 1, raw: true, defval: '' });
      const ih = rows.findIndex(r => r.map(x => norm(x).replace(/[^a-z ]/g, '').trim()).includes('id rdv'));
      if (ih < 0) continue;
      const h = rows[ih].map(x => norm(x).replace(/[^a-z ]/g, ' ').replace(/\s+/g, ' ').trim()), idx = {};
      Object.entries(RDV_COLS_IMPORT).forEach(([k, cands]) => { for (const c of cands) { const i = h.indexOf(c); if (i >= 0 && !Object.values(idx).includes(i)) { idx[k] = i; break; } } });
      lignes = rows.slice(ih + 1).filter(r => r.some(x => String(x).trim())).map(r => { const o = {}; Object.entries(idx).forEach(([k, i]) => { o[k] = ['datePris', 'date'].includes(k) ? iso(r[i]) : String(r[i] ?? '').trim(); }); return o; });
      break;
    }
    if (!lignes) throw new Error("colonne « ID RDV » introuvable : partez d'un export « Exporter Excel » de cette page");
    const r = await api('importerRdvExcel', { lignes: JSON.stringify(lignes), annee: R.annee, fichier: f.name });
    if (!r.success) throw new Error(errTxt(r.message));
    rdvImportApercu(r, lignes, f.name);
  } catch (e) { toast('Import impossible : ' + e.message, false); }
}
function rdvImportApercu(r, lignes, fichier) {
  const n = r.modifs.length + r.creations.length + r.refModifs.length;
  const dlg = modal(`<h2>Import « ${esc(fichier)} » — aperçu</h2>
    <p class="sub"><b>${r.modifs.length}</b> RDV modifié(s) · <b>${r.creations.length}</b> nouveau(x) RDV · <b>${r.refModifs.length}</b> responsable(s) mis à jour${r.erreurs.length ? ` · <b style="color:#c62828">${r.erreurs.length}</b> ligne(s) ignorée(s)` : ''}. Rien n'est encore enregistré.</p>
    <div style="max-height:56vh;overflow:auto">
    ${r.erreurs.length ? `<details class="vx-note" style="display:block;margin-bottom:10px" open><summary><b>Lignes ignorées</b></summary>${r.erreurs.map(x => `<div>${esc(x)}</div>`).join('')}</details>` : ''}
    ${r.modifs.length ? `<table class="vs2-t" style="margin-bottom:12px"><thead><tr><th>Étudiant</th><th>Modification</th></tr></thead><tbody>${r.modifs.map(m => `<tr><td><b>${esc(m.etudiant)}</b></td><td>${m.changements.map(c => `<div><b>${esc(c.champ)}</b> : <span style="color:#94a3b8;text-decoration:line-through">${esc(c.avant || '—')}</span> → <b style="color:#15803d">${esc(c.apres || '—')}</b></div>`).join('')}</td></tr>`).join('')}</tbody></table>` : ''}
    ${r.creations.length ? `<table class="vs2-t" style="margin-bottom:12px"><thead><tr><th>Nouveau RDV</th><th>Pris le</th><th>Date</th><th>Visiteur</th><th>Statut</th></tr></thead><tbody>${r.creations.map(c => `<tr><td><b>${esc(c.etudiant)}</b></td><td>${fmtD(c.datePris)}</td><td>${fmtD(c.date)}</td><td>${esc(c.visiteur || '—')}${c.precision ? ` <i>(${esc(c.precision)})</i>` : ''}</td><td>${esc(c.statut)}</td></tr>`).join('')}</tbody></table>` : ''}
    ${r.refModifs.length ? `<table class="vs2-t"><thead><tr><th>Responsable visite</th><th>Avant</th><th>Après</th></tr></thead><tbody>${r.refModifs.map(c => `<tr><td><b>${esc(c.etudiant)}</b></td><td style="color:#94a3b8">${esc(c.avant)}</td><td><b style="color:#15803d">${esc(c.apres)}</b></td></tr>`).join('')}</tbody></table>` : ''}
    ${!n ? '<div class="vx-note">Aucun changement par rapport au CRM.</div>' : ''}</div>
    <div class="mft"><button class="vx-btn" onclick="this.closest('.vx-modal').remove()">Annuler</button><button class="vx-btn pri" id="riOk" ${n ? '' : 'disabled'}>Appliquer ${n} changement(s)</button></div>`, 900);
  $('riOk').onclick = async () => {
    $('riOk').disabled = true; $('riOk').innerHTML = 'Enregistrement…';
    const o = await api('importerRdvExcel', { lignes: JSON.stringify(lignes), annee: R.annee, fichier, appliquer: true });
    if (!o.success) { toast('Erreur : ' + errTxt(o.message), false); $('riOk').disabled = false; return; }
    dlg.remove(); toast(`Import appliqué : ${o.modifies} modifié(s), ${o.crees} créé(s), ${o.responsables} responsable(s) ✓`); rdvLoad();
  };
}

// ── Analyse des visites (respecte les filtres diplôme / classes / responsable / recherche) ──
async function rdvAnalyse() {
  const D = R.data;
  if (!R.analyse || R.analyse.annee !== D.annee) {
    $('rdvBody').innerHTML = `<div class="vx-empty"><h3>Calcul de l'analyse…</h3></div>`;
    const r = await api('getRdvAnalyse', { annee: D.annee }, { forceRefresh: true });
    if (!r.success) { $('rdvBody').innerHTML = `<div class="vx-empty"><h3>Analyse impossible</h3><p>${esc(errTxt(r.message))}</p></div>`; return; }
    R.analyse = r;
    if (R.onglet !== 'analyse') return;
  }
  const { L, parEtu, vAnnee } = rdvFiltre(false);
  const ids = new Set(L.map(e => e.id));
  $('rdvQui').innerHTML = rdvAFaire(L, parEtu) + rdvParVisiteur(L, parEtu);
  const hon = v => v.statut === 'Honoré';
  const visites = vAnnee.filter(v => ids.has(v.etuId) && hon(v)).length;
  const alts = L.filter(e => e.alternant), visites1 = alts.filter(e => (parEtu[e.id] || []).some(hon));
  const pct = (a, b) => b ? Math.round(a / b * 100) : 0;
  const mv = R.analyse.mouvements.filter(m => ids.has(m.etuId) || (!R.classes.size && !$('rdvDip').value && !$('rdvResp').value && !$('rdvQ').value));
  const nb = t => new Set(mv.filter(m => m.type === t).map(m => m.etuId)).size;
  $('rdvKpis').innerHTML = `
    <div class="vx-kpi"><span>Visites effectuées</span><b style="color:#15803d">${visites}</b><small>RDV honorés en ${esc(D.annee)}</small></div>
    <div class="vx-kpi"><span>Alternants visités</span><b>${visites1.length} / ${alts.length}</b><small>${pct(visites1.length, alts.length)} % des alternants</small></div>
    <div class="vx-kpi"><span>Alternants non visités</span><b style="color:${alts.length - visites1.length ? '#c2410c' : 'inherit'}">${alts.length - visites1.length}</b><small>${pct(alts.length - visites1.length, alts.length)} % des alternants</small></div>
    <div class="vx-kpi"><span>Changement d'alternance</span><b>${nb('Changement')}</b><small>dans l'année</small></div>
    <div class="vx-kpi"><span>Alternance perdue</span><b style="color:${nb('Perte') ? '#c62828' : 'inherit'}">${nb('Perte')}</b><small>dans l'année</small></div>
    <div class="vx-kpi"><span>Alternance gagnée</span><b style="color:#15803d">${nb('Gain')}</b><small>dans l'année</small></div>`;
  const parCl = {};
  L.forEach(e => { const c = e.classe || 'Classe non renseignée'; const o = parCl[c] = parCl[c] || { diplome: e.diplome, ins: 0, alt: 0, vis: 0, visites: 0, venir: 0 };
    o.ins++; if (e.alternant) { o.alt++; if ((parEtu[e.id] || []).some(hon)) o.vis++; }
    (parEtu[e.id] || []).forEach(v => { if (hon(v)) o.visites++; if (['Planifié', 'Reporté'].includes(v.statut) && v.date >= rdvAuj()) o.venir++; }); });
  // Par responsable : étudiants attribués + TOUTES leurs visites (historique toutes années) honorées et prises non honorées
  const parResp = {}, refDe = {};
  L.forEach(e => { const r = e.referent || '— Sans responsable —'; refDe[e.id] = r; const o = parResp[r] = parResp[r] || { suivis: 0, alt: 0, vis: 0, honAn: 0, honHist: 0, attente: 0 };
    o.suivis++; if (e.alternant) { o.alt++; if ((parEtu[e.id] || []).some(hon)) o.vis++; } });
  D.visites.filter(v => ids.has(v.etuId)).forEach(v => { const o = parResp[refDe[v.etuId]]; if (!o) return;
    if (hon(v)) { o.honHist++; if (v.annee === D.annee) o.honAn++; } else if (rdvAttente(v)) o.attente++; });
  // RDV par mois (septembre → août) : pris (date de prise), honorés, pris non honorés
  const y0 = +String(D.annee).split('-')[0], mois = Array.from({ length: 12 }, (_, i) => { const m = (8 + i) % 12, y = y0 + (i >= 4 ? 1 : 0); return { cle: y + '-' + String(m + 1).padStart(2, '0'), lib: MOIS[m] + ' ' + y, pris: 0, hon: 0, att: 0 }; });
  const parMois = {}; mois.forEach(m => parMois[m.cle] = m);
  vAnnee.filter(v => ids.has(v.etuId)).forEach(v => {
    const kp = (v.datePris || '').slice(0, 7), kd = (v.date || v.datePris || '').slice(0, 7);
    if (parMois[kp]) { parMois[kp].pris++; if (rdvAttente(v)) parMois[kp].att++; }
    if (hon(v) && parMois[kd]) parMois[kd].hon++;
  });
  const maxM = Math.max(1, ...mois.map(m => Math.max(m.pris, m.hon)));
  const barre = p => `<div style="display:flex;align-items:center;gap:8px"><div style="flex:1;min-width:70px;height:8px;border-radius:9px;background:#eef2f7;overflow:hidden"><i style="display:block;height:100%;width:${p}%;background:${p >= 80 ? '#16a34a' : p >= 50 ? '#F26522' : '#e5484d'}"></i></div><b>${p} %</b></div>`;
  const lstMv = t => { const l = mv.filter(m => m.type === t).sort((a, b) => b.date.localeCompare(a.date));
    return `<details class="vs2-mv" ${l.length ? '' : ''}><summary><b>${t === 'Gain' ? 'Alternances gagnées' : t === 'Perte' ? 'Alternances perdues' : "Changements d'alternance"}</b> · ${l.length}</summary>
      ${l.length ? `<table class="vs2-t" style="margin-top:8px"><thead><tr><th>Date</th><th>Étudiant</th><th>Classe</th><th>Avant</th><th>Après</th></tr></thead><tbody>${l.map(m => `<tr><td>${fmtD(m.date)}</td><td><b>${esc(m.etudiant)}</b></td><td>${esc(m.classe)}</td><td>${esc(m.avant || '—')}</td><td>${esc(m.apres || '—')}</td></tr>`).join('')}</tbody></table>` : '<p class="sub" style="margin:8px 0">Aucun.</p>'}</details>`; };
  $('rdvBody').innerHTML = `
    <div class="vx-dip"><span class="t">Par classe</span></div>
    <div style="overflow-x:auto;margin-top:12px"><table class="vs2-t"><thead><tr><th>Classe</th><th>Inscrits</th><th>En alternance</th><th>Visites effectuées</th><th>Alternants visités</th><th>Non visités</th><th>Taux de visite</th><th>RDV à venir</th></tr></thead><tbody>
    ${Object.keys(parCl).sort((a, b) => dipOrdreTri(parCl[a].diplome) - dipOrdreTri(parCl[b].diplome) || a.localeCompare(b, 'fr', { numeric: true })).map(c => { const o = parCl[c]; return `<tr><td><b>${esc(c)}</b> <small style="color:#94a3b8">${esc(o.diplome)}</small></td><td>${o.ins}</td><td>${o.alt}</td><td>${o.visites}</td><td>${o.vis}</td><td style="color:${o.alt - o.vis ? '#c2410c' : 'inherit'};font-weight:800">${o.alt - o.vis}</td><td style="min-width:150px">${barre(pct(o.vis, o.alt))}</td><td>${o.venir}</td></tr>`; }).join('')}
    </tbody></table></div>
    <div class="vx-dip"><span class="t">RDV par mois</span></div>
    <div style="overflow-x:auto;margin-top:12px"><table class="vs2-t"><thead><tr><th>Mois</th><th>RDV pris</th><th>RDV honorés</th><th>Pris, non honorés</th><th style="width:45%"></th></tr></thead><tbody>
    ${mois.map(m => `<tr><td><b>${esc(m.lib)}</b></td><td>${m.pris}</td><td style="color:#15803d;font-weight:800">${m.hon}</td><td style="color:${m.att ? '#c2410c' : 'inherit'};font-weight:800">${m.att}</td>
      <td><div class="rdv-mois"><i class="p" style="width:${m.pris / maxM * 100}%"></i><i class="h" style="width:${m.hon / maxM * 100}%"></i></div></td></tr>`).join('')}
    <tr class="gr"><td><b>Total</b></td><td><b>${mois.reduce((s, m) => s + m.pris, 0)}</b></td><td><b>${mois.reduce((s, m) => s + m.hon, 0)}</b></td><td><b>${mois.reduce((s, m) => s + m.att, 0)}</b></td><td><span style="font-size:11px;color:#7d8a9c"><i class="rdv-lg p"></i> pris <i class="rdv-lg h"></i> honorés</span></td></tr>
    </tbody></table></div>
    <div class="vx-dip"><span class="t">Par responsable visite</span><span class="n">historique toutes années inclus</span></div>
    <div style="overflow-x:auto;margin-top:12px"><table class="vs2-t"><thead><tr><th>Responsable</th><th>Étudiants suivis</th><th>Alternants</th><th>Honorés ${esc(D.annee)}</th><th>Honorés (historique)</th><th>Pris, non honorés</th><th>Alternants visités</th><th>Non visités</th><th>Taux</th></tr></thead><tbody>
    ${Object.keys(parResp).sort((a, b) => a.startsWith('—') - b.startsWith('—') || a.localeCompare(b, 'fr')).map(r => { const o = parResp[r]; return `<tr><td><b>${esc(r)}</b></td><td>${o.suivis}</td><td>${o.alt}</td><td style="color:#15803d;font-weight:800">${o.honAn}</td><td>${o.honHist}</td><td style="color:${o.attente ? '#c2410c' : 'inherit'};font-weight:800">${o.attente}</td><td>${o.vis}</td><td style="color:${o.alt - o.vis ? '#c2410c' : 'inherit'};font-weight:800">${o.alt - o.vis}</td><td style="min-width:150px">${barre(pct(o.vis, o.alt))}</td></tr>`; }).join('')}
    </tbody></table></div>
    <div class="vx-dip"><span class="t">Mouvements d'alternance en ${esc(D.annee)}</span></div>
    <div class="vx-note" style="margin:12px 0">${svg('clock', 14)} Relevé quotidien des entreprises des étudiants inscrits${R.analyse.suiviDepuis ? ' depuis le ' + fmtD(R.analyse.suiviDepuis) : ''}. Avant cette date, seuls les départs enregistrés dans le CRM (changement d'entreprise, entreprise retirée) sont connus : les alternances gagnées ne sont comptées qu'à partir du premier relevé.</div>
    ${lstMv('Changement')}${lstMv('Perte')}${lstMv('Gain')}`;
}

// ── Paramètres : lien de calendrier de chaque responsable + modèle du mail ──
async function rdvParams() {
  const r = await api('getRdvParams', {}, { forceRefresh: true });
  if (!r.success) { toast('Erreur : ' + errTxt(r.message), false); return; }
  const ex = { contact: 'Mme Dupont', etudiant: 'Léa Martin', prenom_etudiant: 'Léa', classe: 'BTSCG2A', entreprise: 'Cabinet Exemple', responsable: user.nom || 'Responsable', lien_rdv: 'https://calendly.com/exemple', annee: R.annee };
  const rep = s => String(s).replace(/\{\{\s*([a-z_]+)\s*\}\}/gi, (m, k) => ex[k.toLowerCase()] !== undefined ? ex[k.toLowerCase()] : m);
  const dlg = modal(`<h2>Paramètres du suivi des RDV</h2><p class="sub">Lien de prise de RDV de chaque responsable de visite (Google Agenda, Calendly, Microsoft Bookings…) et modèle du mail envoyé aux entreprises.</p>
    <label class="fl">Liens de calendrier</label>
    <div style="max-height:260px;overflow:auto;border:1px solid #e3eaf3;border-radius:14px"><table class="vs2-t" style="border:0"><thead><tr><th>Membre de l'équipe</th><th>Lien de prise de RDV</th></tr></thead><tbody>
      ${r.responsables.map((x, i) => `<tr><td class="vs2-n"><b>${esc(x.nom)}</b><small>${esc(x.role || '')} · ${esc(x.email)}</small></td><td><div style="display:flex;gap:6px"><input class="in" data-lien="${i}" value="${esc(x.lien)}" placeholder="https://…" style="margin:0">${x.lien ? `<a class="vx-btn ico sm" href="${esc(x.lien)}" target="_blank" rel="noopener" title="Tester le lien">${svg('ext', 13)}</a>` : ''}</div></td></tr>`).join('')}
    </tbody></table></div>
    <label class="fl">Objet du mail</label><input class="in" id="rpObj" value="${esc(r.modele.objet)}">
    <label class="fl">Corps du mail</label><textarea class="in" id="rpCorps" style="min-height:220px;font-family:inherit">${esc(r.modele.corps)}</textarea>
    <div style="display:flex;flex-wrap:wrap;gap:6px;margin-top:8px;align-items:center"><span class="sub" style="margin:0">Insérer :</span>${r.variables.map(v => `<button type="button" class="vx-chip" style="cursor:pointer;border:1px solid #dfe6f0" data-var="${v}">{{${v}}}</button>`).join('')}
      <button type="button" class="vx-link" style="margin:0 0 0 auto" id="rpDefaut">Rétablir le modèle par défaut</button></div>
    <label class="fl">Aperçu (exemple)</label><div id="rpApercu" class="vx-note" style="display:block;white-space:pre-wrap;line-height:1.5"></div>
    <div class="mft"><button class="vx-btn" onclick="this.closest('.vx-modal').remove()">Annuler</button><button class="vx-btn pri" id="rpOk">Enregistrer</button></div>`, 900);
  const apercu = () => { $('rpApercu').innerHTML = '<b>' + esc(rep($('rpObj').value)) + '</b>\n\n' + esc(rep($('rpCorps').value)); };
  apercu();
  $('rpObj').oninput = apercu; $('rpCorps').oninput = apercu;
  let dernier = $('rpCorps');
  $('rpObj').onfocus = () => dernier = $('rpObj'); $('rpCorps').onfocus = () => dernier = $('rpCorps');
  dlg.querySelectorAll('[data-var]').forEach(b => b.onclick = () => {
    const el = dernier, t = '{{' + b.dataset.var + '}}', s = el.selectionStart ?? el.value.length;
    el.value = el.value.slice(0, s) + t + el.value.slice(el.selectionEnd ?? s); el.focus(); el.selectionStart = el.selectionEnd = s + t.length; apercu();
  });
  $('rpDefaut').onclick = () => { $('rpObj').value = r.modeleDefaut.objet; $('rpCorps').value = r.modeleDefaut.corps; apercu(); };
  $('rpOk').onclick = async () => {
    const responsables = r.responsables.map((x, i) => ({ email: x.email, nom: x.nom, lien: dlg.querySelector(`[data-lien="${i}"]`).value.trim() }));
    const o = await api('saveRdvParams', { responsables, modele: { objet: $('rpObj').value.trim(), corps: $('rpCorps').value } });
    if (!o.success) { toast('Erreur : ' + errTxt(o.message), false); return; }
    dlg.remove(); toast('Paramètres enregistrés ✓');
  };
}

// ── Envoi du lien de prise de RDV : sélection → aperçu → envoi ──
function rdvMails() {
  const D = R.data, vAn = D.visites.filter(v => v.annee === D.annee);
  const resps = [...new Set(D.etudiants.filter(e => e.referent).map(e => e.referent))].sort((a, b) => a.localeCompare(b, 'fr'));
  if (!resps.length) { toast("Aucun responsable visite n'est encore attribué", false); return; }
  const pre = $('rdvResp').value && $('rdvResp').value !== '__aucun' ? $('rdvResp').value : resps[0];
  const dlg = modal(`<h2>Envoyer le lien de prise de RDV</h2>
    <p class="sub">On choisit UN responsable visite : son lien de calendrier part à l'entreprise de chacun des étudiants qu'il doit visiter (un mail par étudiant, le responsable en copie et en adresse de réponse).</p>
    <label class="fl">Responsable visite</label><select class="in" id="rmResp">${resps.map(r => `<option ${r === pre ? 'selected' : ''}>${esc(r)}</option>`).join('')}</select>
    <div id="rmCompte" class="vx-note" style="margin-top:8px"></div>
    <label class="fl">Destinataires</label>
    <label style="display:flex;gap:8px;align-items:center;font-size:13px;margin:4px 0"><input type="radio" name="rmC" value="ma" checked> Le maître d'apprentissage (sinon un contact de l'entreprise)</label>
    <label style="display:flex;gap:8px;align-items:center;font-size:13px;margin:4px 0"><input type="radio" name="rmC" value="tous"> Tous les contacts de l'entreprise</label>
    <label class="fl">Exclure</label>
    <label style="display:flex;gap:8px;align-items:center;font-size:13px;margin:4px 0"><input type="checkbox" id="rmEx1" checked> les étudiants déjà visités ou ayant un RDV à venir</label>
    <label style="display:flex;gap:8px;align-items:center;font-size:13px;margin:4px 0"><input type="checkbox" id="rmEx2" checked> ceux à qui le lien a déjà été envoyé cette année</label>
    <div class="mft"><button class="vx-btn" onclick="this.closest('.vx-modal').remove()">Annuler</button><button class="vx-btn pri" id="rmGo">${svg('search', 14)} Voir l'aperçu des mails</button></div>`, 620);
  const cibles = () => {
    const resp = $('rmResp').value;
    return D.etudiants.filter(e => e.referent === resp && e.alternant)
      .filter(e => !($('rmEx1').checked && vAn.some(v => v.etuId === e.id && (v.statut === 'Honoré' || (['Planifié', 'Reporté'].includes(v.statut) && v.date >= rdvAuj())))))
      .filter(e => !($('rmEx2').checked && D.envois[e.id]));
  };
  const compte = () => { const t = D.etudiants.filter(e => e.referent === $('rmResp').value && e.alternant).length, c = cibles().length;
    $('rmCompte').innerHTML = `${svg('user', 14)} <span><b>${t}</b> alternant(s) attribué(s) à ${esc($('rmResp').value)} · <b>${c}</b> à contacter avec ces critères</span>`; };
  ['rmResp', 'rmEx1', 'rmEx2'].forEach(id => $(id).onchange = compte); compte();
  $('rmGo').onclick = async () => {
    const cible = dlg.querySelector('[name="rmC"]:checked').value, ids = cibles().map(e => e.id);
    if (!ids.length) { toast('Aucun étudiant à contacter avec ces critères', false); return; }
    $('rmGo').disabled = true; $('rmGo').innerHTML = 'Préparation…';
    const r = await api('previewMailsRdv', { annee: R.annee, etuIds: ids, cible });
    if (!r.success) { toast('Erreur : ' + errTxt(r.message), false); $('rmGo').disabled = false; $('rmGo').innerHTML = "Voir l'aperçu des mails"; return; }
    const rp = $('rmResp').value; dlg.remove(); rdvApercu(r, cible, rp);
  };
}
function rdvApercu(r, cible, resp) {
  const ok = r.items.filter(x => !x.bloquant), ko = r.items.filter(x => x.bloquant);
  const raisons = {}; ko.forEach(x => (raisons[x.bloquant] = raisons[x.bloquant] || []).push(x.etudiant));
  const sel = new Set(ok.map(x => x.etuId));
  const dlg = modal(`<h2>Aperçu avant envoi${resp ? ' — ' + esc(resp) : ''}</h2>
    <p class="sub"><b>${ok.length}</b> mail(s) prêt(s)${ko.length ? ` · <b style="color:#c2410c">${ko.length}</b> impossible(s)` : ''}${r.quota != null ? ` · quota d'envoi restant aujourd'hui : ${r.quota}` : ''}. Cliquez sur un mail pour le lire tel qu'il partira.</p>
    ${ko.length ? `<details class="vx-note" style="display:block;margin-bottom:10px"><summary><b>Non envoyables (${ko.length})</b> — à corriger avant</summary>${Object.keys(raisons).map(k => `<div style="margin-top:6px"><b>${esc(k)}</b> : ${raisons[k].map(esc).join(', ')}</div>`).join('')}</details>` : ''}
    <div style="max-height:52vh;overflow:auto;display:flex;flex-direction:column;gap:6px">${ok.map((x, i) => `<div class="rdv-ml">
      <label class="h"><input type="checkbox" data-ml="${esc(x.etuId)}" checked><span><b>${esc(x.etudiant)}</b> <small>${esc(x.classe)} · ${esc(x.entreprise)}</small><small style="display:block">À : ${x.destinataires.map(c => esc(c.nom || c.email) + ' &lt;' + esc(c.email) + '&gt;' + (c.ma ? ' (MA)' : '')).join(', ')} · Cc : ${esc(x.responsable)}${x.responsableEmail ? ' &lt;' + esc(x.responsableEmail) + '&gt;' : ''}</small></span><button type="button" class="vx-btn sm" data-voir="${i}">Lire</button></label>
      <div class="b" data-corps="${i}" style="display:none"><div><b>Objet :</b> ${esc(x.objet)}</div><div style="white-space:pre-wrap;margin-top:8px">${esc(x.corps).replace(/(https?:\/\/[^\s<]+)/g, '<a href="$1" target="_blank" rel="noopener">$1</a>')}</div></div></div>`).join('') || '<div class="vx-note">Aucun mail envoyable.</div>'}</div>
    <div class="mft"><button class="vx-btn" style="margin-right:auto" id="rmTest" ${ok.length ? '' : 'disabled'} title="Envoie le premier mail coché, uniquement à polepedagogique@hecg.fr">${svg('mail', 14)} Mail de test à polepedagogique@hecg.fr</button><button class="vx-btn" onclick="this.closest('.vx-modal').remove()">Annuler</button><button class="vx-btn ora" id="rmSend" ${ok.length ? '' : 'disabled'}>${svg('mail', 14)} Envoyer ${ok.length} mail(s)</button></div>`, 900);
  $('rmTest').onclick = async () => {
    const id = ok.map(x => x.etuId).find(x => sel.has(x)); if (!id) { toast('Cochez au moins un mail', false); return; }
    $('rmTest').disabled = true;
    const o = await api('sendMailsRdv', { annee: R.annee, etuIds: [id], cible, test: true });
    $('rmTest').disabled = false;
    if (!o.success) { toast('Erreur : ' + errTxt(o.message), false); return; }
    toast('Mail de test envoyé à ' + o.a + ' (exemple : ' + o.exemple + ') ✓');
  };
  const maj = () => { $('rmSend').innerHTML = svg('mail', 14) + ' Envoyer ' + sel.size + ' mail(s)'; $('rmSend').disabled = !sel.size; };
  dlg.querySelectorAll('[data-ml]').forEach(c => c.onchange = () => { if (c.checked) sel.add(c.dataset.ml); else sel.delete(c.dataset.ml); maj(); });
  dlg.querySelectorAll('[data-voir]').forEach(b => b.onclick = ev => { ev.preventDefault(); const d = dlg.querySelector(`[data-corps="${b.dataset.voir}"]`); d.style.display = d.style.display === 'none' ? 'block' : 'none'; });
  $('rmSend').onclick = async () => {
    if (!sel.size || !confirm('Envoyer ' + sel.size + ' mail(s) maintenant ?')) return;
    $('rmSend').disabled = true; $('rmSend').innerHTML = 'Envoi en cours…';
    const o = await api('sendMailsRdv', { annee: R.annee, etuIds: [...sel], cible });
    if (!o.success) { toast('Erreur : ' + errTxt(o.message), false); maj(); return; }
    dlg.remove(); toast(o.envoyes + ' mail(s) envoyé(s) ✓' + (o.echecs.length ? ' — ' + o.echecs.length + ' échec(s)' : ''), !o.echecs.length);
    if (o.echecs.length) alert('Échecs :\n' + o.echecs.join('\n'));
    rdvLoad();
  };
}
function rdvEquipeOptions(sel) {
  return `<option value="">— choisir —</option>` + (R.data.equipe || []).map(x => `<option value="${esc(x.email)}" ${sel && (sel === x.email || norm(sel) === norm(x.nom)) ? 'selected' : ''}>${esc(x.nom)}${x.role ? ' · ' + esc(x.role) : ''}</option>`).join('') + `<option value="__autre">Autre (saisie libre)…</option>`;
}
function rdvVisiteurChoisi(selId, libreId) {
  const v = $(selId).value;
  if (v === '__autre') return { visiteur: $(libreId).value.trim(), visiteurEmail: '' };
  const x = (R.data.equipe || []).find(q => q.email === v);
  return x ? { visiteur: x.nom, visiteurEmail: x.email } : { visiteur: '', visiteurEmail: '' };
}
function rdvForm(e, v) {
  const D = R.data, ref = v ? v.visiteurEmail || v.visiteur : (e.referentEmail || e.referent);
  const dlg = modal(`<h2>${v ? 'RDV de ' + esc(v.etudiant) : 'Nouveau RDV — ' + esc(e.nom)}</h2>
    <p class="sub">${v ? `${esc(v.entreprise || '')}${v.lieu ? ' — ' + esc(v.lieu) : ''} · statut <b>${esc(v.statut)}</b>` : `${esc(e.entreprise || 'Entreprise non renseignée')}${e.lieu ? ' — ' + esc(e.lieu) : ''} (lieu figé dans l'historique)`}</p>
    <label class="fl">Membre de l'équipe qui se déplace</label>
    <select class="in" id="rvVis">${rdvEquipeOptions(ref)}</select>
    <input class="in" id="rvLibre" placeholder="Prénom NOM" style="display:none;margin-top:6px" value="${esc(v && !v.visiteurEmail ? v.visiteur : '')}">
    <label class="fl">Précision (zone libre)</label><input class="in" id="rvPrec" placeholder="Ex : à sa demande, remplace un collègue…" value="${esc(v ? v.precision || '' : '')}">
    <label class="fl">${v ? 'Date du RDV' : 'Date RDV pris'}</label><input type="date" class="in" id="rvDate" value="${v ? (v.date || '') : ''}" ${v && v.statut === 'Honoré' ? 'disabled' : ''}>
    <label class="fl">Commentaire</label><textarea class="in" id="rvCom" style="min-height:60px">${esc(v ? v.commentaire : '')}</textarea>
    ${v && v.historique && v.historique.length ? `<label class="fl">Historique du RDV</label><div class="vx-note" style="display:block">${v.historique.map(h => `<div><b>${esc(h.le)}</b> — ${esc(h.statut)}${h.dateReport ? ' au ' + fmtD(h.dateReport) : h.date ? ' (' + fmtD(h.date) + ')' : ''}${h.commentaire ? ' — ' + esc(h.commentaire) : ''} <span style="color:#94a3b8">${esc(h.par || '')}</span></div>`).join('')}</div>` : ''}
    <div class="mft">${v ? `<button class="vx-btn danger" style="margin-right:auto" id="rvDel">${svg('trash', 14)} Retirer</button>` : ''}<button class="vx-btn" onclick="this.closest('.vx-modal').remove()">Annuler</button><button class="vx-btn pri" id="rvOk">Enregistrer</button></div>`, 560);
  if (v && !v.visiteurEmail && v.visiteur) { $('rvVis').value = '__autre'; $('rvLibre').style.display = ''; }
  $('rvVis').onchange = () => { $('rvLibre').style.display = $('rvVis').value === '__autre' ? '' : 'none'; };
  $('rvOk').onclick = async () => {
    const who = rdvVisiteurChoisi('rvVis', 'rvLibre');
    const payload = Object.assign({ date: $('rvDate').value, commentaire: $('rvCom').value.trim(), precision: $('rvPrec').value.trim() }, who, v ? { id: v.id } : { etuId: e.id, annee: R.annee });
    const r = await api('saveRdv', { rdv: payload });
    if (!r.success) { toast('Erreur : ' + errTxt(r.message), false); return; }
    dlg.remove(); toast(v ? 'RDV mis à jour ✓' : 'RDV enregistré ✓'); rdvLoad();
  };
  if ($('rvDel')) $('rvDel').onclick = async () => {
    if (!confirm('Retirer ce RDV du suivi ? (il reste archivé dans la feuille RDV_VISITES)')) return;
    const r = await api('saveRdv', { rdv: { id: v.id, archiver: true } });
    if (!r.success) { toast('Erreur : ' + errTxt(r.message), false); return; }
    dlg.remove(); toast('RDV retiré'); rdvLoad();
  };
}
function rdvNew(etuId) { rdvForm(R.data.etudiants.find(e => e.id === etuId), null); }
function rdvEdit(id) { const v = R.data.visites.find(x => x.id === id); if (v) rdvForm(R.data.etudiants.find(e => e.id === v.etuId) || {}, v); }
async function rdvStatut(id, statut) {
  const v = R.data.visites.find(x => x.id === id); if (!v) return;
  const go = async (extra) => {
    const r = await api('statutRdv', Object.assign({ id, statut }, extra || {}));
    if (!r.success) { toast('Erreur : ' + errTxt(r.message), false); return false; }
    toast('RDV ' + statut.toLowerCase() + ' ✓'); rdvLoad(); return true;
  };
  if (statut === 'Honoré') { go(); return; }
  const dlg = modal(`<h2>RDV ${statut === 'Autre' ? '— autre issue' : statut.toLowerCase()}</h2><p class="sub">${esc(v.etudiant)} · ${esc(v.entreprise)} · prévu le ${fmtD(v.date)}</p>
    ${statut === 'Reporté' ? `<label class="fl">Nouvelle date du RDV</label><input type="date" class="in" id="rsDate">` : ''}
    <label class="fl">Commentaire${statut === 'Autre' ? ' (obligatoire)' : ''}</label><textarea class="in" id="rsCom" style="min-height:60px" placeholder="${statut === 'Annulé' ? 'Motif de l\'annulation' : statut === 'Reporté' ? 'Raison du report' : 'Que s\'est-il passé ?'}"></textarea>
    <div class="mft"><button class="vx-btn" onclick="this.closest('.vx-modal').remove()">Annuler</button><button class="vx-btn pri" id="rsOk">Valider</button></div>`, 480);
  $('rsOk').onclick = async () => {
    const extra = { commentaire: $('rsCom').value.trim() };
    if (statut === 'Reporté') { extra.dateReport = $('rsDate').value; if (!extra.dateReport) { toast('Indiquez la nouvelle date', false); return; } }
    if (statut === 'Autre' && !extra.commentaire) { toast('Précisez ce qui s\'est passé', false); return; }
    if (await go(extra)) dlg.remove();
  };
}
function rdvReferent(etuId) { rdvAffecter([etuId]); }
function rdvHisto(etuId) {
  const e = R.data.etudiants.find(x => x.id === etuId) || {};
  const l = R.data.visites.filter(v => v.etuId === etuId).sort((a, b) => (b.date || '').localeCompare(a.date || ''));
  modal(`<h2>Historique des visites — ${esc(e.nom || (l[0] || {}).etudiant || '')}</h2><p class="sub">Toutes les années, avec l'entreprise et le lieu au moment du RDV.</p>
    <table class="vs2-t"><thead><tr><th>Date</th><th>Année</th><th>Visiteur</th><th>Structure (lieu)</th><th>Statut</th></tr></thead><tbody>
    ${l.map(v => `<tr><td><b>${fmtD(v.date || v.datePris)}</b>${v.datePris && v.date !== v.datePris ? `<div style="font-size:11px;color:#94a3b8">pris le ${fmtD(v.datePris)}</div>` : ''}</td><td>${esc(v.annee)}</td><td>${esc(v.visiteur || '—')}</td>
      <td class="vs2-n"><b>${esc(v.entreprise || '—')}</b><small>${esc(v.lieu || '')}</small></td><td><span class="st-b st-${v.statut.replace(/\s/g, '')}">${esc(v.statut)}</span>${v.source === 'Charlemagne' ? '<div style="font-size:11px;color:#94a3b8">repris de Charlemagne</div>' : ''}</td></tr>`).join('')}
    </tbody></table><div class="mft"><button class="vx-btn pri" onclick="this.closest('.vx-modal').remove()">Fermer</button></div>`, 820);
}
async function rdvCharl() {
  if (!confirm("Reprendre les responsables et les dates de visite du dernier import Charlemagne ?\n\n• le responsable de visite devient le responsable de l'année (une saisie manuelle n'est jamais écrasée) ;\n• chaque date de visite trouvée crée un RDV (passée = Honoré), sans doublon.")) return;
  const r = await api('importerRdvCharlemagne', {});
  if (!r.success) { toast('Erreur : ' + errTxt(r.message), false); return; }
  const S = r.stats;
  toast(`Charlemagne : ${S.responsables} responsable(s), ${S.visites} visite(s) créée(s)${S.nonRapproches ? ', ' + S.nonRapproches + ' étudiant(s) non retrouvé(s)' : ''}`);
  rdvLoad();
}
async function rdvExport() {
  const D = R.data; if (!D) return;
  let X; try { X = await loadExcelJS(); } catch (e) { toast(e.message, false); return; }
  const wb = new X.Workbook(), ws = wb.addWorksheet('RDV ' + D.annee, { views: [{ state: 'frozen', ySplit: 4 }] });
  const H = ['ID RDV', 'ID étudiant', 'Diplôme', 'Classe', 'Étudiant', 'Responsable visite', 'Précision responsable', 'N° RDV', 'Date RDV pris', 'Date RDV (après report)', 'Visiteur', 'Précision', 'Entreprise', 'Code postal', 'Lieu', 'Statut', 'Commentaire'];
  ws.columns = [16, 12, 12, 14, 26, 22, 20, 8, 14, 16, 22, 20, 28, 11, 36, 12, 40].map(w => ({ width: w }));
  xlTitle(ws, 'Suivi des RDV en entreprise — ' + D.annee, 'Export du ' + fmtLong(todayISO()) + ' — modifiable puis réimportable (« Importer Excel ») : ne pas toucher aux colonnes ID ; ajouter une ligne avec l\'ID étudiant et une date pour créer un RDV', H.length); xlHeader(ws, 4, H);
  const SC = { 'Honoré': ['E8F6EE', '15803D'], 'Planifié': ['E8EFF8', '1A4E8A'], 'Reporté': ['FFF4ED', 'C2410C'], 'Annulé': ['FDE8E8', 'C62828'], 'Autre': ['EEF2F7', '475569'], 'À planifier': ['F1F4F8', '7D8A9C'] };
  let r = 5;
  parDipClasse(D.etudiants.filter(e => e.alternant || D.visites.some(v => v.etuId === e.id && v.annee === D.annee))).forEach(([d, cls]) => cls.forEach(([c, l]) => l.forEach(e => {
    const vs = D.visites.filter(v => v.etuId === e.id && v.annee === D.annee).sort((a, b) => (a.datePris || '').localeCompare(b.datePris || ''));
    (vs.length ? vs : [null]).forEach((v, i) => {
      ws.getRow(r).values = [v ? v.id : '', e.id, d, c, e.nom, e.referent || '', e.referentPrecision || '', v ? i + 1 : '', v && v.datePris ? new Date(v.datePris + 'T12:00:00') : '', v && v.date ? new Date(v.date + 'T12:00:00') : '', v ? v.visiteur : '', v ? v.precision || '' : '', v ? v.entreprise : e.entreprise, e.cp || '', v ? v.lieu : e.lieu, v ? v.statut : 'Aucun RDV', v ? v.commentaire : ''];
      [9, 10].forEach(k => ws.getCell(r, k).numFmt = 'dd/mm/yyyy');
      [1, 2].forEach(k => ws.getCell(r, k).font = { color: { argb: 'FF94A3B8' }, size: 9 });
      const sc = v ? SC[v.statut] : ['FDE8E8', 'C62828'];
      if (sc) { const cl = ws.getCell(r, 16); cl.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + sc[0] } }; cl.font = { bold: true, color: { argb: 'FF' + sc[1] } }; }
      r++;
    });
  })));
  xlBand(ws, 5, r - 1, H.length);
  ws.autoFilter = { from: { row: 4, column: 1 }, to: { row: Math.max(4, r - 1), column: H.length } };
  await xlDownload(wb, 'Suivi_RDV_' + D.annee + '.xlsx');
}

// ═══════════════════════════ COMPARATIF AVEC LE SUIVI DES DIRECTEURS ═══════════════════════════
const DR = { data: null, lecture: 1, verdict: '', q: '' };
const DIR_COUL = {
  'Absent du CRM': 'c-ko', 'Non inscrit dans le CRM': 'c-ko', 'Absent du fichier directeur': 'c-ko', 'Rouge chez le directeur, inscrit dans le CRM': 'c-ko',
  'Doublon dans le fichier directeur': 'c-wa', 'Présent, diplôme différent': 'c-wa', 'Présent des deux côtés': 'c-ok', 'Sorti des deux côtés': 'c-ne',
  'Vert chez le directeur, pas alternant dans le CRM': 'c-ko', 'Alternant dans le CRM, pas vert chez le directeur': 'c-ko', 'Alternance des deux côtés': 'c-ok', 'Sans alternance des deux côtés': 'c-ne' };
function dirInit() {
  injectCss(); injectCss2();
  const sec = $('section-directeurs');
  if (!sec.dataset.ready) {
    sec.dataset.ready = '1'; sec.classList.add('vx');
    sec.innerHTML = `<div class="vx-hd"><div><h1>Comparatif avec le suivi des directeurs</h1><p>Lecture seule du classeur « Suivi inscription étudiants » (fond rouge = plus dans l'établissement, fond vert = en alternance). Rien n'est écrit, ni dans le classeur ni dans le CRM.</p></div>
      <div style="display:flex;gap:8px;flex-wrap:wrap"><a class="vx-btn" id="dirLien" target="_blank" rel="noopener" style="display:none">${svg('ext', 14)} Ouvrir le classeur</a><button class="vx-btn pri" onclick="VX.dirLoad()">${svg('clock', 15)} Actualiser</button></div></div>
      <div id="dirOnglets" style="margin-bottom:14px"></div><div id="dirBody"><div class="vx-empty"><h3>Lecture du classeur des directeurs…</h3></div></div>`;
  }
  if (!DR.data) dirLoad();
}
async function dirLoad() {
  const choisis = [...document.querySelectorAll('[data-dironglet]:checked')].map(x => x.value);
  $('dirBody').innerHTML = `<div class="vx-empty"><h3>Lecture du classeur des directeurs…</h3></div>`;
  const r = await api('getComparatifDirecteurs', choisis.length ? { onglets: choisis } : {}, { forceRefresh: true });
  if (!r.success) { $('dirBody').innerHTML = `<div class="vx-empty"><h3>Lecture impossible</h3><p>${esc(errTxt(r.message))}</p></div>`; return; }
  DR.data = r;
  if (r.url) { $('dirLien').href = r.url; $('dirLien').style.display = ''; }
  $('dirOnglets').innerHTML = `<div class="vx-bar" style="margin:0"><b style="font-size:12px;color:#7d8a9c;text-transform:uppercase">Onglets comparés</b>${r.onglets.map(o => `<label class="vx-chip" style="cursor:pointer"><input type="checkbox" data-dironglet value="${esc(o)}" ${r.choisis.includes(o) ? 'checked' : ''}> ${esc(o.trim())}</label>`).join('')}<button class="vx-btn sm" onclick="VX.dirLoad()">Comparer</button></div>`;
  dirRender();
}
function dirRender() {
  const D = DR.data; if (!D || !D.stats) { $('dirBody').innerHTML = `<div class="vx-empty"><h3>${esc((D && D.message) || 'Choisissez les onglets à comparer')}</h3></div>`; return; }
  const S = D.stats;
  const carte = (titre, a, la, b, lb, detail) => `<div class="vs2-cmp ${a === b ? 'ok' : 'ko'}"><div><div class="k">${la}</div><div class="v">${a}</div></div><div class="eq">${a === b ? '=' : '≠'}</div><div><div class="k">${lb}</div><div class="v">${b}</div></div><div class="t">${titre} — ${a === b ? '✅ identique' : '❌ écart de ' + Math.abs(a - b)}${detail ? ' · ' + detail : ''}</div></div>`;
  const L = DR.lecture === 1 ? D.lecture1 : DR.lecture === 2 ? D.lecture2 : null;
  const autres = Object.entries(D.couleursAutres || {});
  let corps = '';
  if (L) {
    const cpt = {}; L.forEach(x => cpt[x.verdict] = (cpt[x.verdict] || 0) + 1);
    const ordre = Object.keys(DIR_COUL).filter(v => cpt[v]);
    const q = norm(DR.q);
    const F = L.filter(x => (!DR.verdict || x.verdict === DR.verdict) && (!q || norm([x.dir && x.dir.texte, x.dir && x.dir.classe, x.crm && x.crm.nom, x.crm && x.crm.classe].join(' ')).includes(q)))
      .sort((a, b) => ordre.indexOf(a.verdict) - ordre.indexOf(b.verdict) || String((a.crm || a.dir).nom).localeCompare(String((b.crm || b.dir).nom), 'fr'));
    corps = `<div class="vs2-chips"><button class="c-ne ${DR.verdict ? '' : 'on'}" onclick="VX.dirVerdict('')">Tous · ${L.length}</button>${ordre.map(v => `<button class="${DIR_COUL[v]} ${DR.verdict === v ? 'on' : ''}" onclick="VX.dirVerdict('${jsq(v)}')">${esc(v)} · ${cpt[v]}</button>`).join('')}</div>
      <div class="vx-bar"><div class="vx-search">${svg('search', 15)}<input id="dirQ" placeholder="Étudiant, classe…" value="${esc(DR.q)}" oninput="VX.dirQ(this.value)"></div><button class="vx-btn sm" onclick="VX.dirCsv()">${svg('down', 13)} Export CSV</button></div>
      <div style="overflow-x:auto"><table class="vs2-t"><thead><tr><th>Fichier des directeurs</th><th>CRM</th><th>${DR.lecture === 1 ? 'Inscription' : 'Alternance'}</th></tr></thead><tbody>
      ${F.slice(0, 1500).map(x => `<tr><td class="vs2-n">${x.dir ? `<b><span class="sw" style="background:${esc(x.dir.couleur)}"></span>${esc(x.dir.nom)}</b><small>${esc(x.dir.classe)} · ${esc(x.dir.campus || x.dir.onglet.trim())} · case ${esc(x.dir.ref)}</small>${x.dir.texte !== x.dir.nom ? `<small style="display:block;color:#94a3b8">${esc(x.dir.texte.slice(0, 140))}</small>` : ''}` : '<small style="color:#c62828;font-weight:700">absent du fichier</small>'}</td>
        <td class="vs2-n">${x.crm ? `<a href="dossier.html?id=${encodeURIComponent(x.crm.id)}" target="_blank">${esc(x.crm.nom)}</a><small>${esc(x.crm.classe || 'sans classe')} · ${esc(x.crm.typo || '?')} · ${esc(x.crm.statut || '')}</small>${x.crm.entreprise ? `<small style="display:block">${esc(x.crm.entreprise)}</small>` : ''}` : '<small style="color:#c62828;font-weight:700">introuvable dans le CRM</small>'}</td>
        <td><span class="st-b ${DIR_COUL[x.verdict]}">${esc(x.verdict)}</span></td></tr>`).join('') || '<tr><td colspan="3" style="text-align:center;color:#94a3b8;padding:30px">Aucune ligne</td></tr>'}</tbody></table></div>`;
  } else {
    corps = `<table class="vs2-t"><thead><tr><th>Classe (fichier directeur)</th><th>Inscrits directeur</th><th>…retrouvés inscrits au CRM</th><th>Verts (alternance) directeur</th><th>…alternants au CRM</th></tr></thead><tbody>
      ${D.parClasse.map(c => `<tr><td><b>${esc(c.classe)}</b></td><td>${c.dir}</td><td style="color:${c.crm === c.dir ? '#15803d' : '#c62828'};font-weight:800">${c.crm}</td><td>${c.dirVert}</td><td style="color:${c.crmAlt === c.dirVert ? '#15803d' : '#c62828'};font-weight:800">${c.crmAlt}</td></tr>`).join('')}</tbody></table>`;
  }
  $('dirBody').innerHTML = `<div class="vs2-band">
      ${carte('1re lecture : nombre d\'inscrits', S.dirInscrits, 'Directeurs (hors rouge)', S.crmInscrits, 'CRM inscrits + préinscrits', S.dirRouges + ' case(s) rouge(s) ignorée(s)')}
      ${carte('2e lecture : alternance', S.dirAlternants, 'Directeurs (cases vertes)', S.crmAlternants, 'CRM en alternance', S.altDirSeul + ' vert(s) sans alternance CRM · ' + S.altCrmSeul + ' alternant(s) CRM pas en vert')}
    </div>
    ${autres.length ? `<div class="vx-note" style="margin-bottom:12px">Autres couleurs du classeur, comptées comme « inscrit sans alternance » : ${autres.map(([c, n]) => `<span><span class="sw" style="background:${esc(c)}"></span>${n}</span>`).join(' · ')}</div>` : ''}
    <div class="vx-tabs" style="margin-bottom:14px"><button class="${DR.lecture === 1 ? 'on' : ''}" onclick="VX.dirLecture(1)">1re lecture — inscrits</button><button class="${DR.lecture === 2 ? 'on' : ''}" onclick="VX.dirLecture(2)">2e lecture — alternance</button><button class="${DR.lecture === 3 ? 'on' : ''}" onclick="VX.dirLecture(3)">Par classe</button></div>
    ${corps}`;
}
function dirCsv() {
  const L = DR.lecture === 1 ? DR.data.lecture1 : DR.data.lecture2;
  dl('comparatif_directeurs_' + (DR.lecture === 1 ? 'inscrits' : 'alternance') + '.csv', [['Constat', 'Nom (directeur)', 'Classe (directeur)', 'Campus', 'Case', 'Texte de la case', 'Couleur', 'Étudiant CRM', 'ID CRM', 'Classe CRM', 'Typologie', 'Statut', 'Entreprise']]
    .concat(L.map(x => [x.verdict, x.dir ? x.dir.nom : '', x.dir ? x.dir.classe : '', x.dir ? x.dir.campus : '', x.dir ? x.dir.ref : '', x.dir ? x.dir.texte : '', x.dir ? (x.dir.absent ? 'rouge' : x.dir.vert ? 'vert' : x.dir.couleur) : '',
      x.crm ? x.crm.nom : '', x.crm ? x.crm.id : '', x.crm ? x.crm.classe : '', x.crm ? x.crm.typo : '', x.crm ? x.crm.statut : '', x.crm ? x.crm.entreprise : ''])));
}

// ═══════════════════════════ INSCRIPTIONS AUX EXAMENS OFFICIELS ═══════════════════════════
const EX = { data: null, session: '', q: '', dip: '', filtre: '' };
function exInit() {
  injectCss(); injectCss2();
  const sec = $('section-examens');
  if (!sec.dataset.ready) {
    sec.dataset.ready = '1'; sec.classList.add('vx');
    sec.innerHTML = `<div class="vx-hd"><div><h1>Inscriptions aux examens officiels</h1><p>Par diplôme et par classe : épreuves, inscription (à cocher) et justificatif (« Justificatif OK » dès qu'un document est déposé)</p></div>
      <div style="display:flex;gap:8px;flex-wrap:wrap"><button class="vx-btn" onclick="VX.exExport()">${svg('down', 15)} Exporter Excel</button></div></div>
      <div class="vx-kpis" id="exKpis"></div>
      <div class="vx-bar">
        <select class="vx-sel" id="exSession" onchange="VX.exLoad(this.value)"></select>
        <div class="vx-search">${svg('search', 15)}<input id="exQ" placeholder="Étudiant, classe, n° candidat…" oninput="VX.exRender()"></div>
        <select class="vx-sel" id="exDip" onchange="VX.exRender()"></select>
        <select class="vx-sel" id="exFiltre" onchange="VX.exRender()"><option value="">Tous</option><option value="noninscrit">Non inscrits</option><option value="nojustif">Sans justificatif</option><option value="complet">Inscrits + justificatif</option></select>
      </div>
      <input type="file" id="exFile" multiple hidden>
      <div id="exBody"><div class="vx-empty"><h3>Chargement…</h3></div></div>`;
    $('exFile').onchange = ev => { exUpload([...ev.target.files]); ev.target.value = ''; };
  }
  exLoad();
}
async function exLoad(session) {
  if (session) EX.session = session;
  const r = await api('getExamens', EX.session ? { session: EX.session } : {}, { forceRefresh: true });
  if (!r.success) { $('exBody').innerHTML = `<div class="vx-empty"><h3>Chargement impossible</h3><p>${esc(errTxt(r.message))}</p></div>`; return; }
  EX.data = r; EX.session = r.session;
  $('exSession').innerHTML = r.sessions.map(s => `<option ${s === r.session ? 'selected' : ''} value="${s}">Session ${s}</option>`).join('');
  const dips = [...new Set(r.etudiants.map(e => e.diplome))].sort((a, b) => dipOrdreTri(a) - dipOrdreTri(b));
  const v = $('exDip').value; $('exDip').innerHTML = '<option value="">Tous les diplômes</option>' + dips.map(d => `<option>${esc(d)}</option>`).join(''); $('exDip').value = v;
  exRender();
}
function exFiltrees() {
  const q = norm($('exQ').value), dip = $('exDip').value, f = $('exFiltre').value;
  return EX.data.etudiants.filter(e => (!dip || e.diplome === dip) && (!q || norm([e.nom, e.classe, e.numCandidat, e.epreuves.join(' ')].join(' ')).includes(q))
    && (!f || (f === 'noninscrit' && !e.inscrit) || (f === 'nojustif' && !e.justifOk) || (f === 'complet' && e.inscrit && e.justifOk)));
}
function exRender() {
  const D = EX.data; if (!D) return;
  const all = D.etudiants, L = exFiltrees();
  const ins = all.filter(e => e.inscrit).length, jus = all.filter(e => e.justifOk).length, comp = all.filter(e => e.inscrit && e.justifOk).length;
  $('exKpis').innerHTML = `<div class="vx-kpi"><span>Étudiants</span><b>${all.length}</b><small>session ${esc(D.session)}</small></div>
    <div class="vx-kpi"><span>Inscrits</span><b style="color:#15803d">${ins}</b><small>${all.length - ins} non inscrit(s)</small></div>
    <div class="vx-kpi"><span>Justificatif OK</span><b style="color:#15803d">${jus}</b><small>${all.length - jus} sans justificatif</small></div>
    <div class="vx-kpi"><span>Dossiers complets</span><b>${comp}</b><small>inscrit + justificatif</small></div>`;
  if (!L.length) { $('exBody').innerHTML = `<div class="vx-empty"><h3>Aucun étudiant pour ce filtre</h3></div>`; return; }
  $('exBody').innerHTML = parDipClasse(L).map(([d, classes]) => {
    const n = classes.reduce((s, c) => s + c[1].length, 0), ni = classes.reduce((s, c) => s + c[1].filter(e => e.inscrit).length, 0);
    return `<div class="vx-dip"><span class="t">${esc(d === 'Autre' ? 'Autres formations' : d)}</span><span class="n">${n} étudiant(s)</span><span class="n">${ni} inscrit(s)</span></div>
    <div style="overflow-x:auto;margin-top:12px"><table class="vs2-t"><thead><tr><th>Étudiant</th><th>Épreuves</th><th>N° candidat</th><th>État 1 · Inscrit</th><th>Justificatifs</th><th>État 2 · Justificatif</th></tr></thead><tbody>
    ${classes.map(([c, l]) => `<tr class="gr"><td colspan="6"><b>${esc(c)}</b><span>${l.length} étudiant(s)</span><span>${l.filter(e => e.inscrit).length} inscrit(s) · ${l.filter(e => e.justifOk).length} justificatif(s)</span>
      ${l.some(e => !e.inscrit) ? `<button class="vx-btn sm" style="float:right" onclick="VX.exInscrireClasse('${jsq(d)}','${jsq(c)}')">${svg('check', 12, 2.6)} Marquer toute la classe inscrite</button>` : ''}</td></tr>` + l.map(e => `<tr>
      <td class="vs2-n"><b>${esc(e.nom)}</b><small>${esc(e.typo || '')}${e.sorti ? ' · <span style="color:#c62828">plus inscrit</span>' : ''}</small>${e.mail ? `<small style="display:block">${esc(e.mail)}</small>` : ''}</td>
      <td><div class="ex-ep">${e.epreuves.map(x => `<span>${esc(x)}</span>`).join('')}</div><button class="vx-link" onclick="VX.exEpreuves('${jsq(e.id)}')">${svg('edit', 12)} ${e.epreuves.length ? 'Modifier' : 'Choisir les épreuves'}</button></td>
      <td><input class="ex-num" value="${esc(e.numCandidat)}" placeholder="—" onchange="VX.exNum('${jsq(e.id)}',this.value)"></td>
      <td><button class="ex-tg ${e.inscrit ? 'on' : ''}" onclick="VX.exInscrit('${jsq(e.id)}',${!e.inscrit})" title="${e.inscrit ? 'Inscrit le ' + fmtD(e.inscritLe) + ' par ' + esc(e.inscritPar) + ' — cliquer pour annuler' : 'Cliquer pour marquer inscrit'}">${e.inscrit ? '✓ Inscrit' + (e.inscritLe ? ' le ' + fmtD(e.inscritLe) : '') : '○ Marquer inscrit'}</button></td>
      <td>${e.docs.map(x => `<div class="ex-doc">${svg('file', 12)}<a href="${esc(x.url)}" target="_blank" rel="noopener" title="${esc(x.nom)}">${esc(x.nom.replace(/^.*— /, ''))}</a><button title="Retirer" onclick="VX.exDocDel('${jsq(e.id)}','${jsq(x.id)}')">✕</button></div>`).join('')}
        <button class="vx-link" onclick="VX.exPick('${jsq(e.id)}')">${svg('upload', 12)} Déposer</button></td>
      <td><span class="ex-ok ${e.justifOk ? 'y' : 'n'}">${e.justifOk ? '✓ Justificatif OK' : 'En attente'}</span></td></tr>`).join('')).join('')}
    </tbody></table></div>`; }).join('');
}
async function exSave(ids, champs, msg) {
  const r = await api('saveExamen', { session: EX.session, etuIds: ids, champs });
  if (!r.success) { toast('Erreur : ' + errTxt(r.message), false); return false; }
  if (msg) toast(msg); return true;
}
async function exInscrit(id, on) {
  const e = EX.data.etudiants.find(x => x.id === id); if (!e) return;
  if (!on && !confirm('Annuler l\'inscription de ' + e.nom + ' ?')) return;
  if (await exSave([id], { inscrit: on }, on ? 'Inscription enregistrée ✓' : 'Inscription annulée')) { e.inscrit = on; e.inscritLe = on ? todayISO() : ''; e.inscritPar = on ? (user.nom || '') : ''; exRender(); }
}
async function exInscrireClasse(d, c) {
  const l = EX.data.etudiants.filter(e => e.diplome === d && (e.classe || 'Classe non renseignée') === c && !e.inscrit);
  if (!l.length || !confirm('Marquer inscrits les ' + l.length + ' étudiant(s) non inscrits de ' + c + ' ?')) return;
  if (await exSave(l.map(e => e.id), { inscrit: true }, l.length + ' inscription(s) enregistrée(s) ✓')) { l.forEach(e => { e.inscrit = true; e.inscritLe = todayISO(); e.inscritPar = user.nom || ''; }); exRender(); }
}
async function exNum(id, v) { const e = EX.data.etudiants.find(x => x.id === id); if (e && await exSave([id], { numCandidat: v.trim() }, 'N° candidat enregistré ✓')) e.numCandidat = v.trim(); }
function exEpreuves(id) {
  const e = EX.data.etudiants.find(x => x.id === id); if (!e) return;
  const ref = (EX.data.epreuvesRef[e.diplome] || []).slice();
  e.epreuves.forEach(x => { if (!ref.includes(x)) ref.push(x); });
  const dlg = modal(`<h2>Épreuves — ${esc(e.nom)}</h2><p class="sub">${esc(e.diplome)} · ${esc(e.classe)} — cochez les épreuves auxquelles l'étudiant est inscrit.</p>
    <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:6px;margin:12px 0">${ref.map((x, i) => `<label class="vx-chip" style="cursor:pointer;justify-content:flex-start;white-space:normal"><input type="checkbox" data-ep value="${esc(x)}" ${e.epreuves.includes(x) ? 'checked' : ''}> ${esc(x)}</label>`).join('') || '<p class="sub">Aucune épreuve dans le référentiel pour ce diplôme : ajoutez-les ci-dessous.</p>'}</div>
    <label class="fl">Autre épreuve</label><input class="in" id="exEpLibre" placeholder="Ex : UE 4 — Droit fiscal">
    <div class="mft"><button class="vx-btn" onclick="this.closest('.vx-modal').remove()">Annuler</button><button class="vx-btn pri" id="exEpOk">Enregistrer</button></div>`, 720);
  $('exEpOk').onclick = async () => {
    const l = [...dlg.querySelectorAll('[data-ep]:checked')].map(x => x.value); const libre = $('exEpLibre').value.trim(); if (libre) l.push(libre);
    if (await exSave([id], { epreuves: l }, 'Épreuves enregistrées ✓')) { e.epreuves = l; dlg.remove(); exRender(); }
  };
}
let exCible = null;
function exPick(id) { exCible = id; $('exFile').click(); }
async function exUpload(files) {
  const e = EX.data.etudiants.find(x => x.id === exCible); if (!e || !files.length) return;
  for (const f of files) {
    if (f.size > 20 * 1024 * 1024) { toast(f.name + ' : trop volumineux (20 Mo max)', false); continue; }
    toast('Envoi de ' + f.name + '…');
    const b64 = await new Promise((res, rej) => { const fr = new FileReader(); fr.onload = () => res(String(fr.result).split(',')[1] || ''); fr.onerror = rej; fr.readAsDataURL(f); });
    const r = await api('uploadExamDoc', { session: EX.session, etuId: e.id, data: b64, fileName: f.name, mimeType: f.type || 'application/octet-stream' });
    if (!r.success) { toast(f.name + ' : ' + errTxt(r.message), false); continue; }
    e.docs.push(r.doc); e.justifOk = true;
  }
  exRender(); toast('Justificatif déposé — état « Justificatif OK » ✓');
}
async function exDocDel(etuId, docId) {
  if (!confirm('Retirer ce justificatif ? (le fichier reste dans Drive)')) return;
  const r = await api('archiveExamDoc', { id: docId });
  if (!r.success) { toast('Erreur : ' + errTxt(r.message), false); return; }
  const e = EX.data.etudiants.find(x => x.id === etuId); if (e) { e.docs = e.docs.filter(x => x.id !== docId); e.justifOk = e.docs.length > 0; }
  exRender();
}
async function exExport() {
  const D = EX.data; if (!D) return;
  let X; try { X = await loadExcelJS(); } catch (e) { toast(e.message, false); return; }
  const wb = new X.Workbook(), ws = wb.addWorksheet('Inscriptions ' + D.session, { views: [{ state: 'frozen', ySplit: 4 }] });
  const H = ['Diplôme', 'Classe', 'Étudiant', 'Mail', 'Épreuves', 'N° candidat', 'Inscrit', 'Inscrit le', 'Justificatif OK', 'Justificatifs', 'Commentaire'];
  ws.columns = [12, 14, 26, 30, 44, 14, 10, 12, 14, 40, 30].map(w => ({ width: w }));
  xlTitle(ws, 'Inscriptions aux examens officiels — session ' + D.session, 'Export du ' + fmtLong(todayISO()), H.length); xlHeader(ws, 4, H);
  const OK = ['E8F6EE', '15803D'], KO = ['FDE8E8', 'C62828'];
  let r = 5;
  parDipClasse(exFiltrees()).forEach(([d, cls]) => cls.forEach(([c, l]) => l.forEach(e => {
    ws.getRow(r).values = [d, c, e.nom, e.mail, e.epreuves.join(', '), e.numCandidat, e.inscrit ? 'Oui' : 'Non', e.inscritLe ? new Date(e.inscritLe + 'T12:00:00') : '', e.justifOk ? 'Oui' : 'Non', e.docs.map(x => x.url).join('\n'), e.commentaire];
    ws.getCell(r, 8).numFmt = 'dd/mm/yyyy';
    [[7, e.inscrit], [9, e.justifOk]].forEach(([k, ok]) => { const cl = ws.getCell(r, k), sc = ok ? OK : KO; cl.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + sc[0] } }; cl.font = { bold: true, color: { argb: 'FF' + sc[1] } }; cl.alignment = { horizontal: 'center' }; });
    r++;
  })));
  xlBand(ws, 5, r - 1, H.length);
  ws.autoFilter = { from: { row: 4, column: 1 }, to: { row: Math.max(4, r - 1), column: H.length } };
  await xlDownload(wb, 'Inscriptions_examens_' + D.session + '.xlsx');
}

window.VX = {
  memInit, memRender, memEdit, memExport, memKwManage, memTab, ckDip, ckYear, ckConcerne, ckCyclade, ckExport,
  ckSel, ckSelAll, ckFilter, ckAddSel,
  memSort: v => { M.sort = v; memRender(); }, memView: v => { M.view = v; $('memVC').classList.toggle('on', v === 'cards'); $('memVT').classList.toggle('on', v === 'table'); memRender(); },
  memKw: k => { M.kw = norm(M.kw) === norm(k) ? '' : k; memRender(); }, memSortCol: k => { M.sortDir = M.sortCol === k ? -M.sortDir : 1; M.sortCol = k; memRender(); },
  frInit, frSel, frToggle, frEdit, frPdf,
  rdvInit, rdvLoad: a => rdvLoad(a), rdvRender, rdvOnglet, rdvClsPanel, rdvCls, rdvParams, rdvMails, rdvHarmoniser, rdvImport,
  rdvSel, rdvSelTout, rdvAffecter, rdvDrapeau, rdvResp: r => { $('rdvResp').value = r; rdvRender(); }, rdvDates: () => { $('rdvDu').value = ''; $('rdvAu').value = ''; rdvRender(); },
  rdvVue: v => { $('rdvVue').value = v; if (R.onglet !== 'suivi') rdvOnglet('suivi'); else rdvRender(); },
  rdvQ: q => { $('rdvQ').value = q; rdvRender(); }, rdvNew, rdvEdit, rdvStatut, rdvReferent, rdvHisto, rdvCharl, rdvExport,
  dirInit, dirLoad, dirCsv, dirVerdict: v => { DR.verdict = DR.verdict === v ? '' : v; dirRender(); }, dirLecture: n => { DR.lecture = n; DR.verdict = ''; dirRender(); },
  dirQ: v => { DR.q = v; dirRender(); const i = $('dirQ'); if (i) { i.focus(); i.setSelectionRange(v.length, v.length); } },
  exInit, exLoad: s => exLoad(s), exRender, exInscrit, exInscrireClasse, exNum, exEpreuves, exPick, exDocDel, exExport
};
})();
