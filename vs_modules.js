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
  flag: '<path d="M4 22V4"/><path d="M4 4h13l-2.5 4.5L17 13H4"/>', mail: '<rect x="2" y="4" width="20" height="16" rx="2"/><path d="m22 7-10 6L2 7"/>'
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
  .vx-dot{width:22px;height:22px;border-radius:50%;cursor:pointer;border:2px solid #fff;box-shadow:0 0 0 1px var(--li)}
  .vx-dot.on{box-shadow:0 0 0 2px var(--ink)}
  @media(max-width:760px){.vx .g2{grid-template-columns:1fr}.fr-step{grid-template-columns:1fr 1fr}}
  `;
  document.head.appendChild(st);
}

// ═══════════════════════════ MÉMOIRES ═══════════════════════════
const MEM_KW = ['Fiscalité', 'CDG', 'IA', 'Dématérialisation', 'Comptabilité', 'Audit', 'Finance', 'Paie', 'Droit social', 'RH', 'Marketing', 'RSE', 'Gestion de trésorerie', 'Consolidation', 'Digitalisation'];
const MEM_ST = { 'En cours': ['#e8eff8', '#1A4E8A'], 'Sujet validé': ['#fff4ed', '#c2410c'], 'Rendu': ['#ede9fe', '#6d28d9'], 'Soutenu': ['#e8f6ee', '#15803d'] };
const M = { data: [], etudiants: [], personnel: [], view: 'cards', sort: 'classe', kw: '', sortCol: 'classe', sortDir: 1 };

async function memLoad(force) {
  const r = await api('getMemoires', {}, force ? { forceRefresh: true } : {});
  if (!r.success) { $('memBody').innerHTML = `<div class="vx-empty"><h3>Chargement impossible</h3><p>${esc(r.message)}</p></div>`; return; }
  M.data = r.data || []; M.etudiants = r.etudiants || []; M.personnel = r.personnel || [];
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
        <div style="display:flex;gap:8px;flex-wrap:wrap"><button class="vx-btn" onclick="VX.memExport()">${svg('down', 15)} Exporter</button><button class="vx-btn ora" onclick="VX.memEdit()">${svg('plus', 15, 2.4)} Ajouter un mémoire</button></div></div>
      <div class="vx-kpis" id="memKpis"></div>
      <div class="vx-bar">
        <div class="vx-search">${svg('search', 15)}<input id="memQ" placeholder="Étudiant, sujet, mot-clé, formateur…" oninput="VX.memRender()"></div>
        <select class="vx-sel" id="memFClasse" onchange="VX.memRender()"></select>
        <select class="vx-sel" id="memFCampus" onchange="VX.memRender()"></select>
        <select class="vx-sel" id="memFForm" onchange="VX.memRender()"></select>
        <select class="vx-sel" id="memSort" onchange="VX.memSort(this.value)">
          <option value="classe">Grouper par classe</option><option value="campus">Grouper par campus</option><option value="formateur">Grouper par formateur</option><option value="statut">Grouper par statut</option><option value="etudiant">Trier par étudiant</option><option value="recent">Plus récents</option></select>
        <span class="vx-seg"><button id="memVC" class="on" onclick="VX.memView('cards')">${svg('grid', 14)}Cartes</button><button id="memVT" onclick="VX.memView('table')">${svg('list', 14)}Tableau</button></span>
      </div>
      <div id="memKwBar" style="display:flex;flex-wrap:wrap;gap:6px;margin:-4px 0 16px"></div>
      <div id="memBody"><div class="vx-empty"><div class="ill">${svg('book', 26)}</div><h3>Chargement…</h3></div></div>`;
  }
  memLoad();
}
function memFiltered() {
  const q = norm($('memQ').value), fc = $('memFClasse').value, fca = $('memFCampus').value, ff = $('memFForm').value;
  return M.data.filter(m => (!fc || m.classe === fc) && (!fca || m.campus === fca) && (!ff || m.formateur === ff)
    && (!M.kw || m.motsCles.some(k => norm(k) === norm(M.kw)))
    && (!q || norm([m.etudiant, m.sujet, m.formateur, m.classe, m.campus, m.motsCles.join(' ')].join(' ')).includes(q)));
}
function fillSel(id, label, values) {
  const el = $(id), v = el.value;
  el.innerHTML = `<option value="">${label}</option>` + [...new Set(values.filter(Boolean))].sort((a, b) => a.localeCompare(b, 'fr')).map(x => `<option>${esc(x)}</option>`).join('');
  el.value = v;
}
function memRender() {
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
  if (M.view === 'table') { body.innerHTML = memTable(list); return; }
  const key = { classe: 'classe', campus: 'campus', formateur: 'formateur', statut: 'statut' }[M.sort];
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
    <div class="sujet ${m.sujet ? '' : 'empty'}">${esc(m.sujet || 'Sujet à définir')}</div>
    ${m.motsCles.length ? `<div class="kws">${m.motsCles.map(k => `<span class="vx-kw" style="${kwColor(k)}" onclick="event.stopPropagation();VX.memKw('${jsq(k)}')">${esc(k)}</span>`).join('')}</div>` : ''}
    <div class="ft">${m.formateur ? avatar(m.formateur, 24) + `<span>Référent <b>${esc(m.formateur)}</b></span>` : `<span style="color:#c2410c;font-weight:600">${svg('user', 13)} Sans formateur référent</span>`}</div>
  </div>`;
}
function memTable(list) {
  const cols = [['etudiant', 'Étudiant'], ['classe', 'Classe'], ['campus', 'Campus'], ['sujet', 'Sujet'], ['motsCles', 'Mots-clés'], ['formateur', 'Formateur référent'], ['statut', 'Statut']];
  const sorted = list.slice().sort((a, b) => M.sortDir * String(Array.isArray(a[M.sortCol]) ? a[M.sortCol].join(',') : a[M.sortCol] || '').localeCompare(String(Array.isArray(b[M.sortCol]) ? b[M.sortCol].join(',') : b[M.sortCol] || ''), 'fr', { numeric: true }));
  return `<table class="vx-table"><thead><tr>${cols.map(([k, l]) => `<th onclick="VX.memSortCol('${k}')">${l}${M.sortCol === k ? (M.sortDir > 0 ? ' ↑' : ' ↓') : ''}</th>`).join('')}</tr></thead><tbody>
    ${sorted.map(m => { const st = MEM_ST[m.statut] || MEM_ST['En cours']; return `<tr onclick="VX.memEdit('${jsq(m.id)}')">
      <td><div style="display:flex;align-items:center;gap:10px">${avatar(m.etudiant, 30)}<b style="color:var(--ink);font-weight:600">${esc(m.etudiant)}</b></div></td>
      <td><span class="vx-chip b">${esc(m.classe)}</span></td><td>${esc(m.campus)}</td>
      <td style="max-width:340px">${esc(m.sujet)}</td>
      <td><div style="display:flex;flex-wrap:wrap;gap:4px">${m.motsCles.map(k => `<span class="vx-kw" style="${kwColor(k)}">${esc(k)}</span>`).join('')}</div></td>
      <td>${esc(m.formateur)}</td><td><span class="vx-st" style="background:${st[0]};color:${st[1]}">${esc(m.statut)}</span></td></tr>`; }).join('')}
  </tbody></table>`;
}
function memEdit(id) {
  const m = id ? M.data.find(x => x.id === id) : null;
  const E = { id: m ? m.id : '', etuId: m ? m.etuId : '', motsCles: m ? m.motsCles.slice() : [] };
  const etuLabel = e => e ? `${e.prenom} ${e.nom}` : '';
  const dlg = modal(`
    <h2>${m ? 'Modifier le mémoire' : 'Nouveau mémoire'}</h2><p class="sub">La classe et le campus sont repris automatiquement de la fiche étudiant.</p>
    <label class="fl">Étudiant</label>
    <div class="vx-ac"><input class="in" id="meEtu" placeholder="Tapez un nom ou un prénom…" autocomplete="off" value="${esc(m ? m.etudiant : '')}"><div class="vx-acl" id="meEtuL" style="display:none"></div></div>
    <div class="vx-fixed" id="meFix"></div>
    <label class="fl">Sujet du mémoire</label><textarea class="in" id="meSujet" placeholder="Ex : L'impact de l'IA sur la révision comptable dans les cabinets">${esc(m ? m.sujet : '')}</textarea>
    <label class="fl">Mots-clés</label>
    <div class="vx-tags" id="meTags" onclick="document.getElementById('meKwIn').focus()"></div>
    <div class="vx-sugg" id="meSugg"></div>
    <div class="g2">
      <div><label class="fl">Formateur référent</label><select class="in" id="meForm"><option value="">— Choisir —</option>${M.personnel.slice().sort((a, b) => a.nom.localeCompare(b.nom, 'fr')).map(p => `<option value="${esc(p.ref)}" ${m && m.formateurRef === p.ref ? 'selected' : ''}>${esc(p.nom)}${p.role ? ' · ' + esc(p.role) : ''}</option>`).join('')}</select></div>
      <div><label class="fl">Statut</label><select class="in" id="meSt">${Object.keys(MEM_ST).map(s => `<option ${m && m.statut === s ? 'selected' : ''}>${s}</option>`).join('')}</select></div>
    </div>
    <label class="fl">Notes</label><textarea class="in" id="meNotes" style="min-height:60px" placeholder="Facultatif">${esc(m ? m.notes : '')}</textarea>
    <div class="mft">${m ? `<button class="vx-btn danger" style="margin-right:auto" id="meDel">${svg('trash', 14)} Retirer</button>` : ''}
      <button class="vx-btn" onclick="this.closest('.vx-modal').remove()">Annuler</button>
      ${m ? '' : `<button class="vx-btn" id="meSaveNext">Enregistrer et ajouter un autre</button>`}
      <button class="vx-btn pri" id="meSave">Enregistrer</button></div>`, 680);
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
    L.innerHTML = items.length ? items.map((e, i) => `<div class="it ${i === act ? 'act' : ''}" data-i="${i}">${avatar(etuLabel(e), 24)}${esc(etuLabel(e))}<small>${esc(e.classe)} · ${esc(e.campus)}</small></div>`).join('') : '<div class="it" style="color:var(--mu)">Aucun étudiant trouvé</div>';
    L.style.display = 'block';
  };
  const pick = e => { if (!e) return; E.etuId = e.id; inp.value = etuLabel(e); L.style.display = 'none'; fix(); $('meSujet').focus(); };
  inp.addEventListener('input', () => { E.etuId = ''; fix(); draw(); });
  inp.addEventListener('keydown', ev => {
    if (L.style.display === 'none') return;
    if (ev.key === 'ArrowDown' || ev.key === 'ArrowUp') { ev.preventDefault(); act = Math.max(0, Math.min(items.length - 1, act + (ev.key === 'ArrowDown' ? 1 : -1))); L.querySelectorAll('[data-i]').forEach((el, i) => el.classList.toggle('act', i === act)); }
    if (ev.key === 'Enter') { ev.preventDefault(); pick(items[act]); }
  });
  L.addEventListener('mousedown', ev => { const it = ev.target.closest('[data-i]'); if (it) { ev.preventDefault(); pick(items[+it.dataset.i]); } });
  inp.addEventListener('blur', () => setTimeout(() => L.style.display = 'none', 150));
  // mots-clés
  const allKw = [...new Set(MEM_KW.concat(...M.data.map(x => x.motsCles)))];
  const drawTags = () => {
    $('meTags').innerHTML = E.motsCles.map((k, i) => `<span class="vx-kw" style="${kwColor(k)}">${esc(k)}<button data-rm="${i}">${svg('x', 11, 2.6)}</button></span>`).join('') + `<input id="meKwIn" placeholder="${E.motsCles.length ? '' : 'Tapez puis Entrée (fisca, CDG, IA…)'}">`;
    $('meSugg').innerHTML = allKw.filter(k => !E.motsCles.some(x => norm(x) === norm(k))).slice(0, 16).map(k => `<span class="vx-kw" style="${kwColor(k)}" data-add="${esc(k)}">+ ${esc(k)}</span>`).join('');
    const ki = $('meKwIn');
    ki.addEventListener('keydown', ev => {
      if ((ev.key === 'Enter' || ev.key === ',') && ki.value.trim()) { ev.preventDefault(); addKw(ki.value.trim()); }
      if (ev.key === 'Backspace' && !ki.value && E.motsCles.length) { E.motsCles.pop(); drawTags(); $('meKwIn').focus(); }
    });
  };
  const addKw = k => {
    const hit = allKw.find(x => norm(x) === norm(k)) || allKw.find(x => norm(x).startsWith(norm(k)) && norm(k).length >= 3) || k;
    if (!E.motsCles.some(x => norm(x) === norm(hit))) E.motsCles.push(hit);
    drawTags(); $('meKwIn').focus();
  };
  dlg.addEventListener('click', ev => {
    const rm = ev.target.closest('[data-rm]'); if (rm) { E.motsCles.splice(+rm.dataset.rm, 1); drawTags(); return; }
    const ad = ev.target.closest('[data-add]'); if (ad) addKw(ad.dataset.add);
  });
  drawTags();
  const save = async next => {
    const k = $('meKwIn').value.trim(); if (k) addKw(k);
    if (!E.etuId) { toast('Choisissez un étudiant dans la liste', false); inp.focus(); return; }
    const payload = { id: E.id, etuId: E.etuId, sujet: $('meSujet').value.trim(), motsCles: E.motsCles, formateurRef: $('meForm').value, statut: $('meSt').value, notes: $('meNotes').value.trim() };
    const r = await api('saveMemoire', { memoire: payload });
    if (!r.success) { toast('Erreur : ' + r.message, false); return; }
    toast(m ? 'Mémoire mis à jour ✓' : 'Mémoire ajouté ✓');
    dlg.remove(); await memLoad(true);
    if (next) memEdit();
  };
  $('meSave').onclick = () => save(false);
  if ($('meSaveNext')) $('meSaveNext').onclick = () => save(true);
  if ($('meDel')) $('meDel').onclick = async () => {
    if (!confirm('Retirer ce mémoire du suivi ? (il reste archivé dans la feuille MEMOIRES)')) return;
    const r = await api('saveMemoire', { memoire: { id: m.id, etuId: m.etuId, sujet: m.sujet, motsCles: m.motsCles, formateurRef: m.formateurRef, statut: m.statut, notes: m.notes, archiver: true } });
    if (!r.success) { toast('Erreur : ' + r.message, false); return; }
    dlg.remove(); toast('Mémoire retiré'); memLoad(true);
  };
  if (!m) inp.focus();
}
function memExport() {
  const list = memFiltered();
  csvDownload('Memoires_HECG_' + todayISO() + '.csv', [['Étudiant', 'Classe', 'Campus', 'Sujet', 'Mots-clés', 'Formateur référent', 'Statut', 'Notes', 'Mis à jour le']]
    .concat(list.map(m => [m.etudiant, m.classe, m.campus, m.sujet, m.motsCles.join(', '), m.formateur, m.statut, m.notes, fmtD(m.majLe)])));
}

// ═══════════════════════════ FRISES ═══════════════════════════
const F = { data: [], cur: null };
const REMS = [[30, 'J-30'], [14, 'J-14'], [7, 'J-7'], [3, 'J-3'], [2, 'J-2'], [1, 'Veille'], [0, 'Jour J']];

async function frLoad(force) {
  const r = await api('getFrises', {}, force ? { forceRefresh: true } : {});
  if (!r.success) { $('frBody').innerHTML = `<div class="vx-empty"><h3>Chargement impossible</h3><p>${esc(r.message)}</p></div>`; return; }
  F.data = r.data || [];
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
  const W = Math.max(steps.length * 230, 900), PAD = 120;
  const x = iso => PAD + ((new Date(iso + 'T12:00:00') - t0) / 86400000) / span * (W - 2 * PAD);
  // décale les points trop proches pour éviter les chevauchements
  const xs = []; steps.forEach((s, i) => { let v = x(s.date); if (i && v - xs[i - 1] < 110) v = xs[i - 1] + 110; xs.push(v); });
  const width = Math.max(W, xs[xs.length - 1] + PAD);
  const doneCount = steps.filter(s => s.statut === 'Rendu').length;
  const progX = Math.min(x(today), width);
  return `<div class="fr-board" style="--c:${esc(f.couleur)}">
    <div class="fr-head"><div><h2><i></i>${esc(f.titre)}</h2>${f.description ? `<p style="margin:6px 0 0;color:var(--mu)">${esc(f.description)}</p>` : ''}
      <div style="display:flex;flex-wrap:wrap;gap:6px;margin-top:10px"><span class="vx-chip">${svg('bell', 12)} Rappels : ${String(f.rappels).split(',').map(r => +r === 0 ? 'jour J' : (+r === 1 ? 'veille' : 'J-' + r)).join(', ')}</span><span class="vx-chip">${svg('mail', 12)} ${esc(f.destinataires || f.emailCreateur)}</span><span class="vx-chip b">${doneCount}/${steps.length} rendue(s)</span></div></div>
      <div style="display:flex;gap:8px"><button class="vx-btn" onclick="VX.frEdit('${jsq(f.id)}')">${svg('edit', 14)} Modifier</button></div></div>
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
async function frToggle(id, rendu) {
  const f = F.data.find(x => x.etapes.some(e => e.id === id)); const e = f && f.etapes.find(x => x.id === id);
  if (!e) return;
  e.statut = rendu ? 'Rendu' : 'À faire'; e.renduLe = rendu ? todayISO() : ''; frRender();
  const r = await api('toggleEtapeFrise', { id, rendu });
  if (!r.success) { e.statut = rendu ? 'À faire' : 'Rendu'; frRender(); toast('Erreur : ' + r.message, false); return; }
  toast(rendu ? 'Étape marquée rendue ✓' : 'Étape remise à faire');
}
function frEdit(id) {
  const f = id ? F.data.find(x => x.id === id) : null;
  const E = { couleur: f ? f.couleur : PALETTE[0], rappels: (f ? f.rappels : '7,2,0').split(',').map(Number), etapes: f ? f.etapes.map(e => ({ ...e })) : [{ date: '', titre: '', aRendre: '' }] };
  const dlg = modal(`<h2>${f ? 'Modifier la frise' : 'Nouvelle frise'}</h2><p class="sub">Vous recevrez un mail de rappel avant chaque échéance non rendue.</p>
    <label class="fl">Titre</label><input class="in" id="frT" value="${esc(f ? f.titre : '')}" placeholder="Ex : Calendrier des mémoires BTS CG 2026">
    <label class="fl">Description</label><input class="in" id="frD" value="${esc(f ? f.description : '')}" placeholder="Facultatif">
    <label class="fl">Couleur</label><div style="display:flex;gap:8px" id="frC"></div>
    <label class="fl">Échéances</label><div class="fr-steps" id="frSteps"></div>
    <button class="vx-btn sm" style="margin-top:10px" id="frAdd">${svg('plus', 14, 2.4)} Ajouter une échéance</button>
    <label class="fl">Rappels par mail</label><div class="fr-rem" id="frR">${REMS.map(([n, l]) => `<label><input type="checkbox" value="${n}" ${E.rappels.includes(n) ? 'checked' : ''}>${l}</label>`).join('')}</div>
    <label class="fl">Destinataires des rappels</label><input class="in" id="frDest" value="${esc(f ? f.destinataires : (user.email || ''))}" placeholder="email1@hecg.fr, email2@hecg.fr">
    <div class="mft">${f ? `<button class="vx-btn danger" style="margin-right:auto" id="frDel">${svg('trash', 14)} Supprimer</button>` : ''}<button class="vx-btn" onclick="this.closest('.vx-modal').remove()">Annuler</button><button class="vx-btn pri" id="frSave">Enregistrer</button></div>`, 760);
  const drawC = () => { $('frC').innerHTML = PALETTE.map(c => `<span class="vx-dot ${c === E.couleur ? 'on' : ''}" style="background:${c}" data-c="${c}"></span>`).join(''); };
  const readSteps = () => { dlg.querySelectorAll('.fr-step').forEach((row, i) => { E.etapes[i].date = row.querySelector('[data-k=date]').value; E.etapes[i].titre = row.querySelector('[data-k=titre]').value; E.etapes[i].aRendre = row.querySelector('[data-k=ar]').value; }); };
  const drawS = () => {
    $('frSteps').innerHTML = E.etapes.map((e, i) => `<div class="fr-step"><input type="date" class="in" data-k="date" value="${esc(e.date)}"><input class="in" data-k="titre" value="${esc(e.titre)}" placeholder="Échéance (ex : Remise du plan)"><input class="in" data-k="ar" value="${esc(e.aRendre)}" placeholder="À rendre (ex : plan détaillé + biblio)"><button class="vx-btn ico danger" data-del="${i}" title="Retirer">${svg('trash', 14)}</button></div>`).join('');
  };
  drawC(); drawS();
  dlg.addEventListener('click', ev => {
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
    const r = await api('saveFrise', { frise: { id: f ? f.id : '', titre, description: $('frD').value.trim(), couleur: E.couleur, rappels: rappels || '0', destinataires: $('frDest').value, etapes: E.etapes.filter(e => e.date && e.titre.trim()) } });
    if (!r.success) { toast('Erreur : ' + r.message, false); return; }
    dlg.remove(); F.cur = r.id; toast(f ? 'Frise mise à jour ✓' : 'Frise créée ✓ — rappels activés'); frLoad(true);
  };
  if ($('frDel')) $('frDel').onclick = async () => {
    if (!confirm('Supprimer cette frise ? Les rappels s\'arrêteront (elle reste archivée dans la feuille FRISES).')) return;
    const r = await api('saveFrise', { frise: { id: f.id, titre: f.titre, archiver: true } });
    if (!r.success) { toast('Erreur : ' + r.message, false); return; }
    dlg.remove(); toast('Frise supprimée'); frLoad(true);
  };
  if (!f) $('frT').focus();
}

window.VX = {
  memInit, memRender, memEdit, memExport,
  memSort: v => { M.sort = v; memRender(); }, memView: v => { M.view = v; $('memVC').classList.toggle('on', v === 'cards'); $('memVT').classList.toggle('on', v === 'table'); memRender(); },
  memKw: k => { M.kw = norm(M.kw) === norm(k) ? '' : k; memRender(); }, memSortCol: k => { M.sortDir = M.sortCol === k ? -M.sortDir : 1; M.sortCol = k; memRender(); },
  frInit, frSel, frToggle, frEdit
};
})();
