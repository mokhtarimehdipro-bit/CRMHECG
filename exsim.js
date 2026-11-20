/* Simulations d'examens blancs — Vie scolaire › Salles & Simulations › Simulations (refonte nov. 2026)
   Backend : exsim.gs (action 'exsim', op ref|liste|get|enregistrer|supprimer).
   Dépend de vie_scolaire.html : api(), _apiFresh(), toast(), refMatieres.
   Règles dures : une salle = une seule classe à la fois ; un surveillant = une seule salle à la fois ;
   un surveillant n'est proposé que sur ses disponibilités (par défaut : toute la semaine).
   L'optimisation (regroupement + mélange nominatif + plans de classe) n'est ouverte qu'une fois la semaine « terminée ». */
(function () {
const $ = id => document.getElementById(id);
const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
const JOURS = ['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi'];
const COUL = ['#2563eb', '#ea580c', '#059669', '#9333ea', '#dc2626', '#0891b2', '#a16207', '#db2777', '#4f46e5', '#4d7c0f', '#0f766e', '#b45309'];
const h2m = h => { const m = String(h || '').match(/^(\d{1,2}):(\d{2})$/); return m ? (+m[1]) * 60 + (+m[2]) : NaN; };
const m2h = m => String(Math.floor(m / 60)).padStart(2, '0') + ':' + String(m % 60).padStart(2, '0');
const ov = (a0, a1, b0, b1) => a0 < b1 && b0 < a1;
const uid = p => (p || 'E') + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
const norm = s => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase();
const PPM = 1.1; // pixels par minute dans l'EDT

const X = { ref: null, id: null, version: '', S: null, ecriture: false, dirty: false, onglet: 'params', classeSel: null, draft: null,
  dispoEdit: null, jourVue: 'tous', etatVue: 'salle', etatFiltre: '', etatVersion: 'optim', prop: null, planSel: null, derniereDuree: 120 };

/* ---------- appels serveur ---------- */
const lire = extra => _apiFresh('exsim', extra, null);
async function chargerRef(force) {
  if (X.ref && !force) return X.ref;
  const r = await lire({ op: 'ref' });
  if (!r.success) throw new Error(r.message || 'Référentiel indisponible');
  X.ref = r; return r;
}

/* ---------- accès aux données ---------- */
const S = () => X.S;
const salle = id => (S().salles || []).find(s => s.id === id) || { id, num: String(id).split('||')[1] || id, campus: String(id).split('||')[0] || '', capacite: 0 };
const surv = id => (S().surveillants || []).find(s => s.id === id);
const plan = id => (X.ref && X.ref.plans && X.ref.plans[id]) || null;
function capSalle(id) {
  const p = plan(id);
  if (p) { const n = p.bureaux.filter(b => b.type !== 'prof').reduce((t, b) => t + (b.places || []).length, 0); if (n) return n; }
  return salle(id).capacite || 0;
}
const etudiants = c => (X.ref && X.ref.etudiants && X.ref.etudiants[c]) || [];
function effectif(c) { const n = etudiants(c).length; if (n) return n; const k = (S().classes || []).find(x => x.classe === c); return k ? k.effectif || 0 : 0; }
const couleur = c => { const i = (S().classes || []).findIndex(x => x.classe === c); return COUL[(i < 0 ? 0 : i) % COUL.length]; };
const dateJour = j => { if (!S().lundi) return ''; const d = new Date(S().lundi + 'T12:00:00'); d.setDate(d.getDate() + j); return String(d.getDate()).padStart(2, '0') + '/' + String(d.getMonth() + 1).padStart(2, '0'); };
const libJour = j => JOURS[j] + (S().lundi ? ' ' + dateJour(j) : '');
const h0 = () => h2m(S().hDebut || '08:00'), h1 = () => h2m(S().hFin || '19:00');
const lecture = () => !X.ecriture || S().statut === 'terminee';
function dispoOk(s, jour, a0, a1) {
  const pl = s.dispo && s.dispo[String(jour)];
  if (!pl) return true;                       // pas de restriction = disponible toute la journée
  return pl.some(p => h2m(p[0]) <= a0 && h2m(p[1]) >= a1);
}
function dispoTexte(s, jour) {
  const pl = s.dispo && s.dispo[String(jour)];
  if (!pl) return 'Toute la journée';
  if (!pl.length) return 'Absent';
  return pl.map(p => p[0] + '–' + p[1]).join(', ');
}
function partage(ep) { // répartition de l'effectif entre les salles de l'épreuve
  let rest = effectif(ep.classe); const out = {};
  const ss = ep.salles.slice().sort((a, b) => capSalle(b.id) - capSalle(a.id));
  ss.forEach((sa, i) => { const n = i === ss.length - 1 ? rest : Math.min(rest, capSalle(sa.id)); out[sa.id] = n; rest -= n; });
  return out;
}
function matieresDe(c) {
  const out = new Set((X.ref.matieresExb || {})[c] || []);
  const ref = window.refMatieres || {}, nc = norm(c);
  Object.keys(ref).forEach(k => {
    const ok = (k === 'BTS CG' && nc.includes('BTS')) || (k === 'DSCG' && nc.includes('DSCG')) || (k === 'DCG' && nc.includes('DCG') && !nc.includes('DSCG')) || nc.includes(norm(k));
    if (ok) ref[k].forEach(m => out.add(m.matiere));
  });
  return [...out];
}

/* ---------- contrôles ---------- */
function etatSalle(id, D) {
  const a0 = h2m(D.debut), a1 = h2m(D.fin);
  for (const e of S().epreuves) {
    if (e.id === D.id || e.jour !== D.jour || !ov(a0, a1, h2m(e.debut), h2m(e.fin))) continue;
    if (e.salles.some(sa => sa.id === id)) return { ok: false, r: e.classe === D.classe ? 'déjà prise par cette classe sur ce créneau' : `occupée par ${e.classe}${e.matiere ? ' (' + e.matiere + ')' : ''} ${e.debut}–${e.fin}` };
  }
  return { ok: true };
}
function etatSurv(sid, D, salleId) {
  const s = surv(sid); if (!s) return { ok: false, r: 'surveillant retiré de la simulation' };
  const a0 = h2m(D.debut), a1 = h2m(D.fin);
  if (!dispoOk(s, D.jour, a0, a1)) return { ok: false, r: 'non disponible (' + dispoTexte(s, D.jour) + ')' };
  if (D.salles.some(sa => sa.id !== salleId && sa.surv.includes(sid))) return { ok: false, r: 'déjà dans une autre salle de cette épreuve' };
  for (const e of S().epreuves) {
    if (e.id === D.id || e.jour !== D.jour || !ov(a0, a1, h2m(e.debut), h2m(e.fin))) continue;
    const sa = e.salles.find(x => x.surv.includes(sid));
    if (sa && sa.id !== salleId) return { ok: false, r: `déjà en salle ${salle(sa.id).num} (${e.classe} ${e.debut}–${e.fin})` };
  }
  return { ok: true };
}
function erreursEpreuve(D) {
  const err = [], a0 = h2m(D.debut), a1 = h2m(D.fin);
  if (!(a1 > a0)) err.push('L\'heure de fin doit être après l\'heure de début.');
  if (a0 < h0() || a1 > h1()) err.push(`L'épreuve sort de la plage horaire de la simulation (${S().hDebut}–${S().hFin}).`);
  if (D.jour >= S().nbJours) err.push('Jour hors de la semaine.');
  for (const e of S().epreuves) if (e.id !== D.id && e.classe === D.classe && e.jour === D.jour && ov(a0, a1, h2m(e.debut), h2m(e.fin))) err.push(`La classe a déjà « ${e.matiere || 'une épreuve'} » de ${e.debut} à ${e.fin}.`);
  D.salles.forEach(sa => {
    if (!S().salles.some(s => s.id === sa.id)) err.push(`Salle ${salle(sa.id).num} : retirée de la simulation.`);
    const st = etatSalle(sa.id, D); if (!st.ok) err.push(`Salle ${salle(sa.id).num} : ${st.r}.`);
    sa.surv.forEach(sid => { const t = etatSurv(sid, D, sa.id); if (!t.ok) err.push(`${surv(sid) ? surv(sid).nom : sid} : ${t.r}.`); });
  });
  return err;
}
function analyse() {
  const C = [], A = [], parEp = {};
  const add = (L, txt, ids, type) => { L.push({ txt, ids: ids || [], type }); (ids || []).forEach(i => { (parEp[i] = parEp[i] || { c: 0, a: 0 })[L === C ? 'c' : 'a']++; }); };
  const E = S().epreuves, ratio = +S().ratio || 25;
  E.forEach(ep => {
    const lab = `${ep.classe} · ${ep.matiere || 'épreuve'} (${libJour(ep.jour)} ${ep.debut}–${ep.fin})`;
    erreursEpreuve(ep).forEach(t => add(C, lab + ' — ' + t, [ep.id], 'conflit'));
    if (!ep.salles.length) add(A, lab + ' — aucune salle', [ep.id], 'sansSalle');
    ep.salles.forEach(sa => { if (!sa.surv.length) add(A, `${lab} — salle ${salle(sa.id).num} sans surveillant`, [ep.id], 'sansSurv'); });
    const cap = ep.salles.reduce((t, sa) => t + capSalle(sa.id), 0), n = effectif(ep.classe);
    if (ep.salles.length && cap < n) add(A, `${lab} — ${n} étudiants pour ${cap} places`, [ep.id], 'capacite');
    const ns = ep.salles.reduce((t, sa) => t + sa.surv.length, 0), besoin = Math.max(ep.salles.length, Math.ceil(n / ratio));
    if (ep.salles.length && ns && ns < besoin) add(A, `${lab} — ${ns} surveillant(s) pour ${n} étudiants (repère : 1 pour ${ratio})`, [ep.id], 'encadrement');
  });
  // dédoublonnage des conflits vus des deux côtés
  const vus = new Set(), CC = C.filter(c => { if (vus.has(c.txt)) return false; vus.add(c.txt); return true; });
  (S().classes || []).forEach(c => { if (!E.some(e => e.classe === c.classe)) A.push({ txt: `${c.classe} — aucune épreuve planifiée`, ids: [], type: 'classeVide' }); });
  return { conflits: CC, alertes: A, parEp };
}

/* ---------- styles ---------- */
function css() {
  if ($('exsimCss')) return;
  const st = document.createElement('style'); st.id = 'exsimCss';
  st.textContent = `
  .ex-ov{position:fixed;inset:0;z-index:9990;background:#eef2f8;display:flex;flex-direction:column;font-family:'Plus Jakarta Sans',Inter,system-ui,sans-serif;color:#0e1a2b}
  .ex-top{display:flex;flex-wrap:wrap;align-items:center;gap:8px;padding:10px 18px;background:#fff;border-bottom:1px solid #e3eaf3}
  .ex-top h2{margin:0;font-size:17px;font-weight:800;color:#1A4E8A}.ex-top h2 small{display:block;font-size:11.5px;color:#7d8a9c;font-weight:600}
  .ex-b{height:34px;padding:0 12px;border-radius:10px;border:1px solid #dfe6f0;background:#fff;font-weight:700;font-size:12.5px;font-family:inherit;color:#334155;cursor:pointer;white-space:nowrap;display:inline-flex;align-items:center;gap:6px}
  .ex-b:hover{background:#f3f6fb}.ex-b.pri{background:#1A4E8A;border-color:#1A4E8A;color:#fff}.ex-b.ora{background:#F26522;border-color:#F26522;color:#fff}.ex-b.vio{background:#6d28d9;border-color:#6d28d9;color:#fff}
  .ex-b.red{color:#dc2626}.ex-b.sm{height:28px;padding:0 9px;font-size:11.5px}.ex-b:disabled{opacity:.45;cursor:not-allowed}
  .ex-tabs{display:flex;gap:4px;padding:8px 18px 0;background:#fff;border-bottom:1px solid #e3eaf3;overflow-x:auto}
  .ex-tab{padding:9px 14px;border:0;background:none;font:700 12.5px inherit;font-family:inherit;color:#64748b;border-bottom:3px solid transparent;cursor:pointer;white-space:nowrap}
  .ex-tab.on{color:#1A4E8A;border-bottom-color:#F26522}.ex-tab .n{display:inline-block;min-width:18px;padding:0 5px;border-radius:9px;background:#eef2f8;font-size:10.5px;margin-left:4px}
  .ex-body{flex:1;overflow:auto;padding:18px}
  .ex-card{background:#fff;border:1px solid #e3eaf3;border-radius:16px;padding:16px;margin-bottom:14px}
  .ex-card h3{margin:0 0 10px;font-size:12px;font-weight:800;letter-spacing:.06em;text-transform:uppercase;color:#7d8a9c}
  .ex-body label.l{display:block;font-size:10.5px;font-weight:800;letter-spacing:.06em;text-transform:uppercase;color:#94a3b8;margin:10px 0 4px}
  .ex-body input[type=text],.ex-body input[type=date],.ex-body input[type=number],.ex-body select,.ex-body textarea{width:100%;height:36px;border:1px solid #dfe6f0;border-radius:9px;padding:0 10px;font:600 13px inherit;font-family:inherit;background:#fff;color:#0e1a2b}
  .ex-body textarea{height:60px;padding:8px 10px}
  .ex-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:10px}
  .ex-chip{display:inline-flex;align-items:center;gap:5px;padding:3px 9px;border-radius:999px;font-size:11.5px;font-weight:700;background:#eef2f8;color:#334155;margin:2px}
  .ex-chip.bad{background:#fee2e2;color:#b91c1c}.ex-chip button{border:0;background:none;cursor:pointer;color:inherit;font-weight:800;padding:0}
  .ex-tbl{width:100%;border-collapse:collapse;font-size:12.5px}.ex-tbl th{text-align:left;font-size:10.5px;text-transform:uppercase;letter-spacing:.05em;color:#7d8a9c;padding:6px 8px;border-bottom:2px solid #e3eaf3;background:#f8fafd}
  .ex-tbl td{padding:6px 8px;border-bottom:1px solid #eef2f8;vertical-align:top}
  .ex-dc{cursor:pointer;border-radius:8px;padding:4px 6px;font-size:11.5px;font-weight:700;text-align:center}
  .ex-dc.full{background:#ecfdf5;color:#047857}.ex-dc.part{background:#fffbeb;color:#b45309}.ex-dc.abs{background:#f1f5f9;color:#94a3b8;text-decoration:line-through}
  .ex-dc.sel{outline:2px solid #F26522}
  .ex-lock{background:#fff7ed;border:1px solid #fed7aa;color:#9a3412;border-radius:12px;padding:10px 14px;font-size:12.5px;font-weight:600;margin-bottom:14px}
  .ex-edt{display:flex;gap:14px;align-items:flex-start}
  .ex-cls{width:220px;flex-shrink:0}.ex-cls .it{display:flex;align-items:center;gap:8px;padding:8px 10px;border-radius:10px;cursor:pointer;font-size:12.5px;font-weight:700;border:1px solid transparent}
  .ex-cls .it:hover{background:#fff}.ex-cls .it.on{background:#fff;border-color:#dfe6f0;box-shadow:0 4px 12px -8px rgba(14,26,43,.4)}
  .ex-dot{width:10px;height:10px;border-radius:3px;flex-shrink:0}
  .ex-cal{flex:1;min-width:0;background:#fff;border:1px solid #e3eaf3;border-radius:16px;padding:10px;overflow-x:auto}
  .ex-calg{display:grid;min-width:640px}
  .ex-calh{font-size:11.5px;font-weight:800;color:#1A4E8A;text-align:center;padding:4px 0 8px}
  .ex-col{position:relative;border-left:1px solid #e3eaf3;cursor:copy}
  .ex-col.ro{cursor:default}
  .ex-tcol{position:relative}.ex-tcol span{position:absolute;right:6px;font-size:10px;color:#94a3b8;font-weight:700;transform:translateY(-50%)}
  .ex-blk{position:absolute;left:3px;right:3px;border-radius:8px;padding:4px 6px;font-size:11px;line-height:1.25;overflow:hidden;cursor:pointer;border-left:4px solid;box-shadow:0 4px 10px -8px rgba(14,26,43,.6)}
  .ex-blk b{display:block;font-size:11.5px}.ex-blk.err{outline:2px solid #dc2626}.ex-blk.warn{outline:2px dashed #f59e0b}.ex-blk.sel{outline:3px solid #F26522;z-index:3}
  .ex-blk.ghost{opacity:.35;pointer-events:none;border-left-style:dashed}
  .ex-form{width:390px;flex-shrink:0;position:sticky;top:0}
  .ex-srow{border:1px solid #e3eaf3;border-radius:10px;padding:7px 9px;margin-bottom:6px;font-size:12px}
  .ex-srow.off{opacity:.55;background:#f8fafc}.ex-srow.bad{border-color:#fca5a5;background:#fef2f2}
  .ex-err{background:#fef2f2;border:1px solid #fecaca;color:#b91c1c;border-radius:10px;padding:8px 10px;font-size:12px;font-weight:600;margin:8px 0}
  .ex-ok{background:#ecfdf5;border:1px solid #a7f3d0;color:#047857;border-radius:10px;padding:8px 10px;font-size:12px;font-weight:600;margin:8px 0}
  .ex-tl{display:grid;grid-template-columns:170px 1fr;align-items:center;border-bottom:1px solid #f1f5f9}
  .ex-tl .lb{font-size:11.5px;font-weight:700;padding:5px 8px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
  .ex-tl .tr{position:relative;height:30px;background:repeating-linear-gradient(90deg,transparent 0,transparent calc(var(--h) - 1px),#eef2f8 calc(var(--h) - 1px),#eef2f8 var(--h))}
  .ex-tl .bk{position:absolute;top:3px;bottom:3px;border-radius:6px;font-size:10.5px;font-weight:700;padding:2px 5px;overflow:hidden;white-space:nowrap;text-overflow:ellipsis;border-left:3px solid;cursor:pointer}
  .ex-tl .na{position:absolute;top:0;bottom:0;background:repeating-linear-gradient(45deg,#e2e8f0 0,#e2e8f0 3px,transparent 3px,transparent 7px)}
  .ex-kpi{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin-bottom:14px}
  .ex-kpi div{background:#fff;border:1px solid #e3eaf3;border-radius:14px;padding:10px 12px}.ex-kpi b{display:block;font-size:22px}.ex-kpi span{font-size:11px;color:#7d8a9c;font-weight:700}
  .ex-plan{position:relative;background:#fff;border:3px solid #1A4E8A;border-radius:6px;margin:10px 0}
  .ex-seat{position:absolute;border-radius:6px;border:1.5px solid #cbd5e1;background:#fff;font-size:9.5px;line-height:1.15;padding:2px 3px;overflow:hidden;border-left-width:5px}
  .ex-seat i{font-style:normal;color:#94a3b8;font-weight:800}.ex-seat.vide{background:#f8fafc;border-style:dashed}
  .ex-msg{position:fixed;left:50%;bottom:22px;transform:translateX(-50%);z-index:10001;padding:10px 18px;border-radius:14px;color:#fff;font-weight:700;font-size:13px;box-shadow:0 10px 30px -10px rgba(0,0,0,.4)}
  `;
  document.head.appendChild(st);
}
function msg(t, ok = true) {
  const old = $('exMsg'); if (old) old.remove();
  const d = document.createElement('div'); d.id = 'exMsg'; d.className = 'ex-msg'; d.style.background = ok ? '#059669' : '#dc2626'; d.textContent = t;
  document.body.appendChild(d); setTimeout(() => d.remove(), 4200);
}

/* ---------- liste des simulations (onglet Simulations) ---------- */
async function liste() {
  const box = $('exsimListe'); if (!box) return;
  box.innerHTML = '<p class="text-center text-gray-400 py-10">Chargement des simulations…</p>';
  const r = await lire({ op: 'liste' });
  if (!r.success) { box.innerHTML = `<p class="text-red-600 font-bold">${esc(r.message)}</p>`; return; }
  X.ecriture = !!r.ecriture;
  const bt = $('btnExsimNew'); if (bt) bt.classList.toggle('hidden', !X.ecriture);
  if (!r.data.length) { box.innerHTML = `<div class="text-center text-gray-400 py-14 bg-white rounded-2xl border border-gray-100"><div class="text-4xl mb-2">🧪</div><p class="font-bold text-gray-500">Aucune simulation d'examens</p><p class="text-sm mt-1">Créez une simulation : surveillants et disponibilités, classes, salles, puis l'emploi du temps de chaque classe.</p></div>`; return; }
  box.innerHTML = '<div class="space-y-3">' + r.data.map(s => {
    const st = s.optim ? ['✨ Optimisée', '#6d28d9', '#f5f3ff'] : s.statut === 'terminee' ? ['✅ Semaine terminée', '#047857', '#ecfdf5'] : ['🛠 En construction', '#b45309', '#fffbeb'];
    let sem = '—'; if (s.lundi) { const d = new Date(s.lundi + 'T12:00:00'); sem = 'Semaine du ' + d.toLocaleDateString('fr-FR', { day: '2-digit', month: 'long', year: 'numeric' }); }
    return `<div class="bg-white rounded-2xl border border-gray-100 shadow-sm p-5 flex flex-wrap gap-4 justify-between items-center">
      <div><div class="flex items-center gap-2 flex-wrap"><h3 class="text-lg font-black text-[#1A4E8A]">${esc(s.titre || 'Sans titre')}</h3>
        <span style="background:${st[2]};color:${st[1]}" class="text-xs font-black px-2 py-0.5 rounded-full">${st[0]}</span></div>
        <div class="flex gap-4 mt-1 text-sm text-gray-500 flex-wrap"><span>📅 ${esc(sem)}</span><span>🎓 ${s.nbClasses} classe(s)</span><span>🏫 ${s.nbSalles} salle(s)</span><span>👁 ${s.nbSurv} surveillant(s)</span><span>📝 ${s.nbEpreuves} épreuve(s)</span></div>
        <div class="text-xs text-gray-400 mt-1">Mis à jour le ${esc(s.maj)}${s.par ? ' par ' + esc(s.par) : ''}</div></div>
      <div class="flex gap-2 flex-wrap"><button class="btn-primary text-xs py-1" onclick="EXSIM.ouvrir('${esc(s.id)}')">Ouvrir</button>
        ${X.ecriture ? `<button class="btn-ghost text-xs py-1" onclick="EXSIM.dupliquer('${esc(s.id)}')">⧉ Dupliquer</button><button class="btn-danger text-xs py-1" onclick="EXSIM.supprimer('${esc(s.id)}','${esc(s.titre).replace(/&#39;/g, '')}')">🗑</button>` : ''}</div></div>`;
  }).join('') + '</div>';
}
function lundiProchain() { const d = new Date(); d.setHours(12, 0, 0, 0); d.setDate(d.getDate() + ((8 - d.getDay()) % 7 || 7)); return d.toISOString().slice(0, 10); }
function normaliser(o) {
  const s = Object.assign({ titre: '', lundi: lundiProchain(), nbJours: 5, hDebut: '08:00', hFin: '18:00', pas: 15, ratio: 25, statut: 'brouillon', surveillants: [], classes: [], salles: [], epreuves: [], optim: null }, o || {});
  s.epreuves.forEach(e => { e.salles = (e.salles || []).map(sa => ({ id: sa.id, surv: sa.surv || [] })); e.jour = +e.jour || 0; });
  s.surveillants.forEach(v => { v.dispo = v.dispo || {}; });
  return s;
}
async function nouveau() {
  try { await chargerRef(true); } catch (e) { toast(e.message, false); return; }
  X.id = null; X.version = ''; X.dirty = false; X.prmOptim = null; X.S = normaliser({ titre: 'Examens blancs — ' });
  ouvrirUI('params');
}
async function ouvrir(id) {
  try {
    const [r] = await Promise.all([lire({ op: 'get', id }), chargerRef(true)]);
    if (!r.success) { toast(r.message, false); return; }
    X.id = r.id; X.version = r.version; X.ecriture = !!r.ecriture; X.dirty = false; X.prmOptim = null; X.S = normaliser(r.sim);
    ouvrirUI(X.S.epreuves.length ? 'globale' : 'params');
  } catch (e) { toast(e.message, false); }
}
async function dupliquer(id) {
  const r = await lire({ op: 'get', id }); if (!r.success) { toast(r.message, false); return; }
  try { await chargerRef(true); } catch (e) { toast(e.message, false); return; }
  const s = normaliser(r.sim); s.titre += ' (copie)'; s.statut = 'brouillon'; s.optim = null;
  X.id = null; X.version = ''; X.prmOptim = null; X.S = s; X.dirty = true; ouvrirUI('params');
}
async function supprimer(id, titre) {
  if (!confirm(`Supprimer définitivement la simulation « ${titre} » ?`)) return;
  const r = await api('exsim', { op: 'supprimer', id }); if (!r.success) { toast(r.message, false); return; }
  toast('Simulation supprimée ✓'); liste();
}

/* ---------- fenêtre principale ---------- */
const ONGLETS = [['params', '⚙ Paramètres'], ['surv', '👁 Surveillants'], ['classes', '🎓 Classes'], ['salles', '🏫 Salles'], ['edt', '🗓 EDT par classe'], ['globale', '🌐 Vue globale'], ['etats', '📊 États'], ['optim', '✨ Optimisation & plans']];
function ouvrirUI(onglet) {
  css(); fermerUI(true);
  const o = document.createElement('div'); o.className = 'ex-ov'; o.id = 'exOv';
  o.innerHTML = '<div class="ex-top" id="exTop"></div><div class="ex-tabs" id="exTabs"></div><div class="ex-body" id="exBody"></div>';
  document.body.appendChild(o); document.body.style.overflow = 'hidden';
  X.onglet = onglet; X.draft = null; X.prop = null; X.planSel = null; X.dispoEdit = null;
  X.classeSel = (S().classes[0] || {}).classe || null;
  document.addEventListener('keydown', raccourcis); window.addEventListener('beforeunload', avantQuitter);
  rendre();
}
function avantQuitter(e) { if (X.dirty) { e.preventDefault(); e.returnValue = ''; } }
function raccourcis(e) { if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's' && $('exOv')) { e.preventDefault(); enregistrer(); } }
function fermerUI(force) {
  if (!force && X.dirty && !confirm('Des modifications ne sont pas enregistrées. Fermer quand même ?')) return;
  const o = $('exOv'); if (o) o.remove(); document.body.style.overflow = '';
  document.removeEventListener('keydown', raccourcis); window.removeEventListener('beforeunload', avantQuitter);
  if (!force) { X.dirty = false; liste(); }
}
function modifie() { X.dirty = true; rendreTop(); }
function rendre() { rendreTop(); rendreTabs(); rendreCorps(); }
function rendreTop() {
  const s = S(), an = analyse(), t = $('exTop'); if (!t) return;
  const st = s.optim ? ['✨ Optimisée', '#6d28d9', '#f5f3ff'] : s.statut === 'terminee' ? ['✅ Semaine terminée', '#047857', '#ecfdf5'] : ['🛠 En construction', '#b45309', '#fffbeb'];
  t.innerHTML = `<h2>🧪 ${esc(s.titre || 'Nouvelle simulation')}<small>${s.lundi ? 'Semaine du ' + esc(dateJour(0)) + ' au ' + esc(dateJour(s.nbJours - 1)) : 'Semaine à définir'} · ${s.classes.length} classe(s) · ${s.salles.length} salle(s) · ${s.surveillants.length} surveillant(s) · ${s.epreuves.length} épreuve(s)</small></h2>
    <span style="background:${st[2]};color:${st[1]}" class="text-xs font-black px-2 py-1 rounded-full">${st[0]}</span>
    ${an.conflits.length ? `<button class="ex-b sm red" onclick="EXSIM.onglet('globale')">⛔ ${an.conflits.length} conflit(s)</button>` : '<span class="text-xs font-bold text-emerald-700">✓ Aucun conflit</span>'}
    ${an.alertes.length ? `<button class="ex-b sm" style="color:#b45309" onclick="EXSIM.onglet('globale')">⚠ ${an.alertes.length} point(s) d'attention</button>` : ''}
    <span style="flex:1"></span>
    ${X.dirty ? '<span class="text-xs font-bold text-orange-600">● Non enregistré</span>' : ''}
    ${X.ecriture ? (s.statut === 'terminee' ? '<button class="ex-b" onclick="EXSIM.rouvrir()">🔓 Rouvrir la semaine</button>' : '<button class="ex-b" style="color:#047857" onclick="EXSIM.terminer()">✅ Terminer la semaine</button>') : ''}
    ${X.ecriture ? '<button class="ex-b pri" id="exSave" onclick="EXSIM.enregistrer()">💾 Enregistrer</button>' : ''}
    <button class="ex-b" onclick="EXSIM.fermer()">Fermer</button>`;
}
function rendreTabs() {
  const s = S(), n = { surv: s.surveillants.length, classes: s.classes.length, salles: s.salles.length, edt: s.epreuves.length };
  $('exTabs').innerHTML = ONGLETS.map(([k, l]) => `<button class="ex-tab ${X.onglet === k ? 'on' : ''}" onclick="EXSIM.onglet('${k}')">${l}${n[k] != null ? `<span class="n">${n[k]}</span>` : ''}${k === 'optim' && s.statut !== 'terminee' ? ' 🔒' : ''}</button>`).join('');
}
function onglet(k) { X.onglet = k; if (k !== 'edt') X.draft = null; rendreTabs(); rendreCorps(); $('exBody').scrollTop = 0; }
function rendreCorps() {
  const b = $('exBody'); if (!b) return;
  const f = { params: vParams, surv: vSurv, classes: vClasses, salles: vSalles, edt: vEdt, globale: vGlobale, etats: vEtats, optim: vOptim }[X.onglet];
  b.innerHTML = (lecture() && ['params', 'surv', 'classes', 'salles', 'edt'].includes(X.onglet)
    ? `<div class="ex-lock">${X.ecriture ? '🔒 La semaine est terminée : lecture seule. Cliquez sur « Rouvrir la semaine » pour la modifier' + (S().optim ? ' (l\'optimisation et les plans de classe seront alors effacés).' : '.') : '👁 Consultation seule (modification réservée aux profils Super-admin et Pédagogie).'}</div>` : '') + f();
}

/* ---------- 1. Paramètres ---------- */
function vParams() {
  const s = S(), ro = lecture() ? 'disabled' : '';
  const heures = []; for (let m = 6 * 60; m <= 22 * 60; m += 30) heures.push(m2h(m));
  return `<div class="ex-card" style="max-width:860px"><h3>Simulation</h3>
    <label class="l">Titre</label><input type="text" ${ro} value="${esc(s.titre)}" oninput="EXSIM.set('titre',this.value)" placeholder="ex : Examens blancs — session 2">
    <div class="ex-grid">
      <div><label class="l">Semaine (lundi)</label><input type="date" ${ro} value="${esc(s.lundi)}" onchange="EXSIM.setLundi(this.value)"></div>
      <div><label class="l">Jours</label><select ${ro} onchange="EXSIM.setJours(+this.value)"><option value="5" ${s.nbJours === 5 ? 'selected' : ''}>Lundi → vendredi</option><option value="6" ${s.nbJours === 6 ? 'selected' : ''}>Lundi → samedi</option></select></div>
      <div><label class="l">Début de journée</label><select ${ro} onchange="EXSIM.set('hDebut',this.value,true)">${heures.map(h => `<option ${h === s.hDebut ? 'selected' : ''}>${h}</option>`).join('')}</select></div>
      <div><label class="l">Fin de journée</label><select ${ro} onchange="EXSIM.set('hFin',this.value,true)">${heures.map(h => `<option ${h === s.hFin ? 'selected' : ''}>${h}</option>`).join('')}</select></div>
      <div><label class="l">Pas des horaires</label><select ${ro} onchange="EXSIM.set('pas',+this.value)"><option value="15" ${+s.pas === 15 ? 'selected' : ''}>15 min</option><option value="30" ${+s.pas === 30 ? 'selected' : ''}>30 min</option></select></div>
      <div><label class="l">Repère d'encadrement</label><select ${ro} onchange="EXSIM.set('ratio',+this.value)">${[15, 20, 25, 30, 35, 40, 50].map(n => `<option value="${n}" ${+s.ratio === n ? 'selected' : ''}>1 surveillant pour ${n} étudiants</option>`).join('')}</select></div>
    </div></div>
    <div class="ex-card" style="max-width:860px"><h3>Comment ça marche</h3><ol style="margin:0;padding-left:18px;font-size:13px;line-height:1.8;color:#334155">
      <li><b>Surveillants</b> : ajoutez les surveillants de la semaine et réglez leurs disponibilités (par défaut : toute la semaine).</li>
      <li><b>Classes</b> et <b>Salles</b> : cochez les classes concernées et les locaux disponibles.</li>
      <li><b>EDT par classe</b> : cliquez dans le calendrier pour placer chaque épreuve. Seules les salles libres et les surveillants disponibles et libres sont proposés.</li>
      <li><b>Vue globale</b> et <b>États</b> : contrôlez la semaine par salle, par classe et par surveillant, puis imprimez.</li>
      <li><b>Terminer la semaine</b> (aucun conflit) débloque l'<b>optimisation</b> : regroupement des classes dans moins de salles, mélange nominatif des étudiants et plans de classe.</li></ol></div>`;
}
function set(k, v, rer) { if (lecture()) return; S()[k] = v; modifie(); if (k === 'titre') rendreTop(); if (rer) rendreCorps(); }
function setLundi(v) { if (!v) return; const d = new Date(v + 'T12:00:00'), w = d.getDay(); d.setDate(d.getDate() - (w === 0 ? 6 : w - 1)); S().lundi = d.toISOString().slice(0, 10); modifie(); rendreCorps(); }
function setJours(n) {
  const hors = S().epreuves.filter(e => e.jour >= n);
  if (hors.length && !confirm(`${hors.length} épreuve(s) sont planifiées le samedi et seront supprimées. Continuer ?`)) { rendreCorps(); return; }
  S().epreuves = S().epreuves.filter(e => e.jour < n); S().nbJours = n; modifie(); rendreTabs(); rendreCorps();
}

/* ---------- 2. Surveillants & disponibilités ---------- */
function vSurv() {
  const s = S(), ro = lecture(), pris = new Set(s.surveillants.map(v => v.id));
  const dispoListe = (X.ref.personnes || []).filter(p => !pris.has(p.id));
  const heuresAff = sid => s.epreuves.reduce((t, e) => t + (e.salles.some(sa => sa.surv.includes(sid)) ? h2m(e.fin) - h2m(e.debut) : 0), 0);
  let ed = '';
  if (X.dispoEdit && !ro) {
    const v = surv(X.dispoEdit.sid), j = X.dispoEdit.jour;
    if (v) {
      const pl = v.dispo[String(j)];
      ed = `<div class="ex-card" style="border-color:#F26522"><h3>Disponibilité — ${esc(v.nom)} · ${esc(libJour(j))}</h3>
        <div class="flex flex-wrap gap-2 items-center">
          <button class="ex-b sm ${!pl ? 'pri' : ''}" onclick="EXSIM.dispo('plein')">Toute la journée</button>
          <button class="ex-b sm ${pl && !pl.length ? 'pri' : ''}" onclick="EXSIM.dispo('absent')">Absent</button>
          <span class="text-xs font-bold text-gray-500 ml-2">Plages :</span>
          <input type="text" id="exPlages" style="width:260px" placeholder="08:00-12:00, 13:30-17:00" value="${pl && pl.length ? esc(pl.map(p => p[0] + '-' + p[1]).join(', ')) : ''}">
          <button class="ex-b sm ora" onclick="EXSIM.dispo('plages')">Appliquer</button>
          <span style="flex:1"></span>
          <button class="ex-b sm" onclick="EXSIM.dispo('semaine')" title="Copie la disponibilité de ce jour sur tous les jours">↔ Appliquer à toute la semaine</button>
          <button class="ex-b sm" onclick="EXSIM.dispoEdit(null)">Fermer</button></div>
        <p class="text-xs text-gray-400 mt-2">Si une épreuve déjà planifiée devient incompatible, elle apparaît en conflit dans la vue globale.</p></div>`;
    }
  }
  return `${ro ? '' : `<div class="ex-card"><h3>Ajouter des surveillants</h3><div class="flex flex-wrap gap-2 items-center">
      <input type="text" list="exPersList" id="exAddSurv" style="max-width:360px" placeholder="Rechercher dans le personnel et les formateurs…" onkeydown="if(event.key==='Enter')EXSIM.ajoutSurv()">
      <datalist id="exPersList">${dispoListe.map(p => `<option value="${esc(p.nom)}">${esc(p.source)}</option>`).join('')}</datalist>
      <button class="ex-b pri" onclick="EXSIM.ajoutSurv()">+ Ajouter</button>
      <span class="text-xs text-gray-400">Un nom absent des listes est ajouté comme surveillant externe. Un surveillant ajouté est disponible toute la semaine.</span></div></div>`}
    ${ed}
    <div class="ex-card"><h3>Disponibilités de la semaine ${ro ? '' : '<span class="normal-case font-semibold text-gray-400">— cliquez sur une case pour la modifier</span>'}</h3>
    ${s.surveillants.length ? `<div style="overflow-x:auto"><table class="ex-tbl"><thead><tr><th>Surveillant</th>${Array.from({ length: s.nbJours }, (_, j) => `<th style="text-align:center">${esc(libJour(j))}</th>`).join('')}<th>Affecté</th><th></th></tr></thead><tbody>
      ${s.surveillants.map(v => `<tr><td><b>${esc(v.nom)}</b><div class="text-xs text-gray-400">${esc(v.source || '')}</div></td>
        ${Array.from({ length: s.nbJours }, (_, j) => { const pl = v.dispo[String(j)], cl = !pl ? 'full' : pl.length ? 'part' : 'abs', sel = X.dispoEdit && X.dispoEdit.sid === v.id && X.dispoEdit.jour === j ? 'sel' : '';
          return `<td><div class="ex-dc ${cl} ${sel}" ${ro ? '' : `onclick="EXSIM.dispoEdit('${esc(v.id)}',${j})"`}>${esc(dispoTexte(v, j))}</div></td>`; }).join('')}
        <td class="text-xs font-bold">${(heuresAff(v.id) / 60).toFixed(1).replace('.', ',')} h</td>
        <td>${ro ? '' : `<button class="ex-b sm red" onclick="EXSIM.retirerSurv('${esc(v.id)}')">Retirer</button>`}</td></tr>`).join('')}
      </tbody></table></div>` : '<p class="text-sm text-gray-400">Aucun surveillant pour l\'instant.</p>'}</div>`;
}
function ajoutSurv() {
  const inp = $('exAddSurv'), nom = String(inp.value || '').trim(); if (!nom) return;
  const p = (X.ref.personnes || []).find(x => norm(x.nom) === norm(nom));
  const id = p ? p.id : 'X-' + norm(nom).replace(/[^A-Z0-9]+/g, '-');
  if (surv(id)) { msg(nom + ' est déjà dans la liste.', false); return; }
  S().surveillants.push({ id, nom: p ? p.nom : nom, source: p ? p.source : 'Externe', dispo: {} });
  S().surveillants.sort((a, b) => a.nom.localeCompare(b.nom, 'fr'));
  modifie(); rendreTabs(); rendreCorps(); setTimeout(() => $('exAddSurv') && $('exAddSurv').focus(), 0);
}
function retirerSurv(id) {
  const n = S().epreuves.filter(e => e.salles.some(sa => sa.surv.includes(id))).length;
  if (n && !confirm(`Ce surveillant est affecté à ${n} épreuve(s). Le retirer de la simulation et de ces épreuves ?`)) return;
  S().epreuves.forEach(e => e.salles.forEach(sa => { sa.surv = sa.surv.filter(x => x !== id); }));
  S().surveillants = S().surveillants.filter(v => v.id !== id); modifie(); rendreTabs(); rendreCorps();
}
function dispoEdit(sid, jour) { X.dispoEdit = sid ? { sid, jour } : null; rendreCorps(); }
function dispo(mode) {
  const v = surv(X.dispoEdit.sid), j = String(X.dispoEdit.jour); if (!v) return;
  if (mode === 'plein') delete v.dispo[j];
  else if (mode === 'absent') v.dispo[j] = [];
  else if (mode === 'semaine') { const cur = v.dispo[j]; for (let k = 0; k < S().nbJours; k++) { if (cur) v.dispo[String(k)] = JSON.parse(JSON.stringify(cur)); else delete v.dispo[String(k)]; } }
  else {
    const txt = $('exPlages').value, pl = [];
    for (const morceau of txt.split(/[,;]+/)) {
      const m = morceau.trim().replace(/h/gi, ':').match(/^(\d{1,2}):?(\d{2})?\s*[-–à]\s*(\d{1,2}):?(\d{2})?$/i);
      if (!morceau.trim()) continue;
      if (!m) { msg('Plage illisible : « ' + morceau.trim() + ' » (format 08:00-12:00)', false); return; }
      const a = m2h(+m[1] * 60 + (+m[2] || 0)), b = m2h(+m[3] * 60 + (+m[4] || 0));
      if (h2m(b) <= h2m(a)) { msg('Plage incohérente : ' + a + '–' + b, false); return; }
      pl.push([a, b]);
    }
    if (!pl.length) { msg('Indiquez au moins une plage, ou choisissez « Absent ».', false); return; }
    pl.sort((x, y) => h2m(x[0]) - h2m(y[0])); v.dispo[j] = pl;
  }
  modifie(); rendreCorps();
}

/* ---------- 3. Classes ---------- */
function vClasses() {
  const s = S(), ro = lecture() ? 'disabled' : '', sel = new Set(s.classes.map(c => c.classe));
  const q = norm(X.filtreClasse || '');
  const L = (X.ref.classes || []).filter(c => !q || norm(c.classe).includes(q));
  return `<div class="ex-card"><h3>Classes concernées par la simulation</h3>
    <div class="flex gap-2 flex-wrap items-center mb-3"><input type="text" style="max-width:280px" placeholder="Filtrer…" value="${esc(X.filtreClasse || '')}" oninput="EXSIM.filtreClasse(this.value)" id="exFC">
      ${ro ? '' : `<button class="ex-b sm" onclick="EXSIM.toutesClasses(true)">Tout cocher (filtre)</button><button class="ex-b sm" onclick="EXSIM.toutesClasses(false)">Tout décocher (filtre)</button>`}
      <span class="text-xs text-gray-400">Effectifs : étudiants actifs du CRM (hors sortis et prospects).</span></div>
    <div class="ex-grid">${L.map(c => `<label class="ex-srow ${sel.has(c.classe) ? '' : 'off'}" style="display:flex;gap:8px;align-items:center;cursor:pointer;opacity:1">
      <input type="checkbox" ${ro} ${sel.has(c.classe) ? 'checked' : ''} onchange="EXSIM.classe('${esc(c.classe)}',this.checked)" style="width:auto">
      ${sel.has(c.classe) ? `<span class="ex-dot" style="background:${couleur(c.classe)}"></span>` : ''}
      <span><b>${esc(c.classe)}</b><br><span class="text-xs text-gray-500">${c.effectif} étudiant(s)${(X.ref.matieresExb[c.classe] || []).length ? ' · ' + X.ref.matieresExb[c.classe].length + ' matière(s) EXB' : ''}</span></span></label>`).join('')}</div></div>`;
}
function filtreClasse(v) { X.filtreClasse = v; rendreCorps(); const i = $('exFC'); if (i) { i.focus(); i.setSelectionRange(v.length, v.length); } }
function classe(c, on) {
  const s = S();
  if (on) { if (!s.classes.some(x => x.classe === c)) s.classes.push({ classe: c, effectif: effectifRef(c) }); }
  else {
    const n = s.epreuves.filter(e => e.classe === c).length;
    if (n && !confirm(`${c} a ${n} épreuve(s) planifiée(s). Les supprimer ?`)) { rendreCorps(); return; }
    s.epreuves = s.epreuves.filter(e => e.classe !== c); s.classes = s.classes.filter(x => x.classe !== c);
  }
  if (!X.classeSel || !s.classes.some(x => x.classe === X.classeSel)) X.classeSel = (s.classes[0] || {}).classe || null;
  modifie(); rendreTabs(); rendreCorps();
}
const effectifRef = c => ((X.ref.classes || []).find(x => x.classe === c) || {}).effectif || 0;
function toutesClasses(on) {
  const q = norm(X.filtreClasse || '');
  const L = (X.ref.classes || []).filter(c => !q || norm(c.classe).includes(q));
  if (!on && S().epreuves.some(e => L.some(c => c.classe === e.classe)) && !confirm('Des épreuves planifiées seront supprimées. Continuer ?')) return;
  L.forEach(c => { if (on) { if (!S().classes.some(x => x.classe === c.classe)) S().classes.push({ classe: c.classe, effectif: c.effectif }); } else { S().epreuves = S().epreuves.filter(e => e.classe !== c.classe); S().classes = S().classes.filter(x => x.classe !== c.classe); } });
  X.classeSel = (S().classes[0] || {}).classe || null; modifie(); rendreTabs(); rendreCorps();
}

/* ---------- 4. Salles ---------- */
function vSalles() {
  const s = S(), ro = lecture() ? 'disabled' : '', sel = new Set(s.salles.map(x => x.id));
  const parCampus = {}; (X.ref.salles || []).forEach(x => (parCampus[x.campus || '—'] = parCampus[x.campus || '—'] || []).push(x));
  return Object.keys(parCampus).sort().map(cp => `<div class="ex-card"><h3>${esc(cp)} ${ro ? '' : `<button class="ex-b sm" style="margin-left:8px" onclick="EXSIM.campus('${esc(cp)}',true)">Tout cocher</button> <button class="ex-b sm" onclick="EXSIM.campus('${esc(cp)}',false)">Tout décocher</button>`}</h3>
    <div class="ex-grid">${parCampus[cp].map(x => { const p = plan(x.id), np = p ? p.bureaux.filter(b => b.type !== 'prof').reduce((t, b) => t + (b.places || []).length, 0) : 0;
      return `<label class="ex-srow ${sel.has(x.id) ? '' : 'off'}" style="display:flex;gap:8px;align-items:center;cursor:pointer;opacity:1">
      <input type="checkbox" ${ro} ${sel.has(x.id) ? 'checked' : ''} onchange="EXSIM.salle('${esc(x.id)}',this.checked)" style="width:auto">
      <span><b>Salle ${esc(x.num)}</b><br><span class="text-xs text-gray-500">${x.capacite || '?'} places${p ? ` · 🪑 plan : ${np} places` : ' · pas de plan (grille auto)'}</span></span></label>`; }).join('')}</div></div>`).join('')
    + '<p class="text-xs text-gray-400">Le plan de salle (bouton 🪑 de l\'onglet Salles) sert à placer les étudiants nominativement après l\'optimisation ; sans plan, une grille est générée selon la capacité.</p>';
}
function salleCoche(id, on) {
  const s = S();
  if (on) { const x = (X.ref.salles || []).find(y => y.id === id); if (x && !s.salles.some(y => y.id === id)) s.salles.push(Object.assign({}, x)); }
  else {
    const n = s.epreuves.filter(e => e.salles.some(sa => sa.id === id)).length;
    if (n && !confirm(`Cette salle est utilisée par ${n} épreuve(s). La retirer de ces épreuves ?`)) { rendreCorps(); return; }
    s.epreuves.forEach(e => { e.salles = e.salles.filter(sa => sa.id !== id); }); s.salles = s.salles.filter(y => y.id !== id);
  }
  s.salles.sort((a, b) => (a.campus + a.num).localeCompare(b.campus + b.num, 'fr', { numeric: true }));
  modifie(); rendreTabs(); rendreCorps();
}
function campus(cp, on) {
  const L = (X.ref.salles || []).filter(x => (x.campus || '—') === cp);
  if (!on && S().epreuves.some(e => e.salles.some(sa => L.some(x => x.id === sa.id))) && !confirm('Des épreuves utilisent ces salles ; elles en seront retirées. Continuer ?')) return;
  L.forEach(x => { if (on) { if (!S().salles.some(y => y.id === x.id)) S().salles.push(Object.assign({}, x)); } else { S().epreuves.forEach(e => { e.salles = e.salles.filter(sa => sa.id !== x.id); }); S().salles = S().salles.filter(y => y.id !== x.id); } });
  S().salles.sort((a, b) => (a.campus + a.num).localeCompare(b.campus + b.num, 'fr', { numeric: true }));
  modifie(); rendreTabs(); rendreCorps();
}

/* ---------- 5. EDT par classe ---------- */
function vEdt() {
  const s = S();
  if (!s.classes.length) return '<div class="ex-card">Cochez d\'abord les classes (onglet 🎓 Classes).</div>';
  if (!s.salles.length || !s.surveillants.length) return `<div class="ex-card">Il manque ${!s.salles.length ? 'les salles (onglet 🏫 Salles)' : ''}${!s.salles.length && !s.surveillants.length ? ' et ' : ''}${!s.surveillants.length ? 'les surveillants (onglet 👁 Surveillants)' : ''} avant de construire les emplois du temps.</div>`;
  if (!s.classes.some(c => c.classe === X.classeSel)) X.classeSel = s.classes[0].classe;
  const an = analyse(), c = X.classeSel, H = (h1() - h0()) * PPM, ro = lecture();
  const cls = s.classes.map(k => {
    const E = s.epreuves.filter(e => e.classe === k.classe), hh = E.reduce((t, e) => t + h2m(e.fin) - h2m(e.debut), 0);
    const pb = E.reduce((t, e) => t + ((an.parEp[e.id] || {}).c || 0), 0), at = E.reduce((t, e) => t + ((an.parEp[e.id] || {}).a || 0), 0);
    return `<div class="it ${k.classe === c ? 'on' : ''}" onclick="EXSIM.choixClasse('${esc(k.classe)}')"><span class="ex-dot" style="background:${couleur(k.classe)}"></span>
      <span style="flex:1;min-width:0"><span style="display:block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(k.classe)}</span><span class="text-xs text-gray-400 font-semibold">${E.length} épreuve(s) · ${(hh / 60).toFixed(1).replace('.', ',')} h · ${effectif(k.classe)} étu.</span></span>
      ${pb ? '<span title="conflit">⛔</span>' : at ? '<span title="point d\'attention">⚠</span>' : E.length ? '<span>✓</span>' : ''}</div>`;
  }).join('');
  const heures = []; for (let m = Math.ceil(h0() / 60) * 60; m <= h1(); m += 60) heures.push(m);
  const cols = Array.from({ length: s.nbJours }, (_, j) => {
    const blocs = s.epreuves.filter(e => e.classe === c && e.jour === j).map(e => {
      const p = an.parEp[e.id] || {}, col = couleur(c), top = (h2m(e.debut) - h0()) * PPM, h = (h2m(e.fin) - h2m(e.debut)) * PPM;
      return `<div class="ex-blk ${p.c ? 'err' : p.a ? 'warn' : ''} ${X.draft && X.draft.id === e.id ? 'sel' : ''}" style="top:${top}px;height:${Math.max(h, 18)}px;background:${col}1a;border-left-color:${col}" onclick="event.stopPropagation();EXSIM.editer('${e.id}')" title="${esc(e.matiere)} ${e.debut}–${e.fin}">
        <b>${esc(e.matiere || 'Épreuve')}</b>${e.debut}–${e.fin}<br>${e.salles.map(sa => '🏫 ' + esc(salle(sa.id).num) + (sa.surv.length ? ' · ' + sa.surv.map(x => esc((surv(x) || { nom: '?' }).nom)).join(', ') : ' · <span style="color:#b45309">sans surveillant</span>')).join('<br>') || '<span style="color:#b45309">aucune salle</span>'}</div>`;
    }).join('');
    const ghost = X.draft && X.draft.classe === c && X.draft.jour === j && !s.epreuves.some(e => e.id === X.draft.id)
      ? `<div class="ex-blk ghost sel" style="top:${(h2m(X.draft.debut) - h0()) * PPM}px;height:${Math.max((h2m(X.draft.fin) - h2m(X.draft.debut)) * PPM, 18)}px;background:${couleur(c)}22;border-left-color:${couleur(c)}"><b>Nouvelle épreuve</b>${X.draft.debut}–${X.draft.fin}</div>` : '';
    return `<div class="ex-col ${ro ? 'ro' : ''}" style="height:${H}px;background:repeating-linear-gradient(180deg,#f1f5f9 0,#f1f5f9 1px,transparent 1px,transparent ${30 * PPM}px)" ${ro ? '' : `onclick="EXSIM.clicCal(event,${j})"`}>${blocs}${ghost}</div>`;
  }).join('');
  return `<div class="ex-edt"><div class="ex-cls">${cls}</div>
    <div class="ex-cal"><div class="flex items-center gap-2 mb-2 flex-wrap"><span class="ex-dot" style="background:${couleur(c)}"></span><b style="font-size:15px;color:#1A4E8A">${esc(c)}</b>
      <span class="text-xs text-gray-500 font-bold">${effectif(c)} étudiant(s)</span><span style="flex:1"></span>
      ${ro ? '' : `<span class="text-xs text-gray-400">Cliquez dans une journée pour créer une épreuve</span>`}</div>
      <div class="ex-calg" style="grid-template-columns:48px repeat(${s.nbJours},1fr)"><div></div>${Array.from({ length: s.nbJours }, (_, j) => `<div class="ex-calh">${esc(libJour(j))}</div>`).join('')}
      <div class="ex-tcol" style="height:${H}px">${heures.map(m => `<span style="top:${(m - h0()) * PPM}px">${m2h(m)}</span>`).join('')}</div>${cols}</div></div>
    <div class="ex-form" id="exForm">${vForm()}</div></div>`;
}
function choixClasse(c) { X.classeSel = c; X.draft = null; rendreCorps(); }
function clicCal(ev, j) {
  if (lecture()) return;
  const r = ev.currentTarget.getBoundingClientRect(), pas = +S().pas || 15;
  let a0 = h0() + Math.floor((ev.clientY - r.top) / PPM / pas) * pas;
  let a1 = Math.min(h1(), a0 + X.derniereDuree); if (a1 - a0 < pas) { a0 = Math.max(h0(), h1() - X.derniereDuree); a1 = h1(); }
  X.draft = { id: uid('E'), classe: X.classeSel, jour: j, debut: m2h(a0), fin: m2h(a1), matiere: '', note: '', salles: [], _nouveau: true };
  rendreCorps();
}
function editer(id) { const e = S().epreuves.find(x => x.id === id); if (!e) return; X.draft = JSON.parse(JSON.stringify(e)); X.onglet = 'edt'; X.classeSel = e.classe; rendreTabs(); rendreCorps(); }
function vForm() {
  const D = X.draft, ro = lecture();
  if (!D) return `<div class="ex-card"><h3>Épreuve</h3><p class="text-sm text-gray-500">${ro ? 'Cliquez sur une épreuve pour voir son détail.' : 'Cliquez dans le calendrier pour créer une épreuve, ou sur une épreuve existante pour la modifier.'}</p>
    <p class="text-xs text-gray-400 mt-2">Seules les salles libres sur le créneau et les surveillants disponibles et non affectés ailleurs peuvent être choisis.</p></div>`;
  const pas = +S().pas || 15, opts = []; for (let m = h0(); m <= h1(); m += pas) opts.push(m2h(m));
  const dis = ro ? 'disabled' : '', n = effectif(D.classe), err = erreursEpreuve(D), part = partage(D);
  const capTot = D.salles.reduce((t, sa) => t + capSalle(sa.id), 0), ratio = +S().ratio || 25;
  const mats = matieresDe(D.classe);
  const sallesHtml = S().salles.map(x => {
    const on = D.salles.find(sa => sa.id === x.id), st = etatSalle(x.id, D);
    if (!on) return `<div class="ex-srow ${st.ok ? '' : 'off'}"><label style="display:flex;gap:8px;align-items:center;${st.ok && !ro ? 'cursor:pointer' : ''}"><input type="checkbox" ${st.ok && !ro ? '' : 'disabled'} onchange="EXSIM.dSalle('${esc(x.id)}',true)" style="width:auto">
      <span><b>Salle ${esc(x.num)}</b> <span class="text-gray-400">· ${capSalle(x.id)} pl. · ${esc(x.campus)}</span>${st.ok ? '' : `<br><span class="text-xs" style="color:#b91c1c">⛔ ${esc(st.r)}</span>`}</span></label></div>`;
    const besoin = Math.max(1, Math.ceil((part[x.id] || 0) / ratio));
    const options = S().surveillants.filter(v => !on.surv.includes(v.id)).map(v => { const t = etatSurv(v.id, D, x.id); return `<option value="${esc(v.id)}" ${t.ok ? '' : 'disabled'}>${esc(v.nom)}${t.ok ? '' : ' — ' + esc(t.r)}</option>`; });
    const nbOk = S().surveillants.filter(v => !on.surv.includes(v.id) && etatSurv(v.id, D, x.id).ok).length;
    return `<div class="ex-srow ${st.ok ? '' : 'bad'}" style="border-color:${st.ok ? '#93c5fd' : ''};background:${st.ok ? '#f5f9ff' : ''}">
      <label style="display:flex;gap:8px;align-items:center"><input type="checkbox" checked ${dis} onchange="EXSIM.dSalle('${esc(x.id)}',false)" style="width:auto">
      <span style="flex:1"><b>Salle ${esc(x.num)}</b> <span class="text-gray-500">· ${part[x.id] || 0} étu. / ${capSalle(x.id)} pl.</span></span></label>
      ${st.ok ? '' : `<div class="text-xs" style="color:#b91c1c">⛔ ${esc(st.r)}</div>`}
      <div style="margin-top:4px">${on.surv.map(sid => { const t = etatSurv(sid, D, x.id); return `<span class="ex-chip ${t.ok ? '' : 'bad'}" title="${esc(t.r || '')}">👁 ${esc((surv(sid) || { nom: sid }).nom)}${t.ok ? '' : ' ⛔'}${ro ? '' : `<button onclick="EXSIM.dSurv('${esc(x.id)}','${esc(sid)}',false)">×</button>`}</span>`; }).join('')}
      ${on.surv.length < besoin ? `<span class="text-xs" style="color:#b45309">repère : ${besoin} surveillant(s)</span>` : ''}</div>
      ${ro ? '' : `<select style="margin-top:4px;height:32px;font-size:12px" onchange="if(this.value)EXSIM.dSurv('${esc(x.id)}',this.value,true)"><option value="">+ Ajouter un surveillant (${nbOk} disponible(s))</option>${options.join('')}</select>`}</div>`;
  }).join('');
  const ex = S().epreuves.some(e => e.id === D.id);
  return `<div class="ex-card" style="border-color:${couleur(D.classe)}"><h3>${ex ? 'Épreuve' : 'Nouvelle épreuve'} — ${esc(D.classe)}</h3>
    <label class="l">Matière / épreuve</label><input type="text" ${dis} list="exMatList" value="${esc(D.matiere)}" oninput="EXSIM.dSet('matiere',this.value)" placeholder="ex : UE 4 — Droit fiscal">
    <datalist id="exMatList">${mats.map(m => `<option value="${esc(m)}">`).join('')}</datalist>
    <div style="display:grid;grid-template-columns:1.3fr 1fr 1fr;gap:8px">
      <div><label class="l">Jour</label><select ${dis} onchange="EXSIM.dSet('jour',+this.value,true)">${Array.from({ length: S().nbJours }, (_, j) => `<option value="${j}" ${j === D.jour ? 'selected' : ''}>${esc(libJour(j))}</option>`).join('')}</select></div>
      <div><label class="l">Début</label><select ${dis} onchange="EXSIM.dSet('debut',this.value,true)">${opts.map(h => `<option ${h === D.debut ? 'selected' : ''}>${h}</option>`).join('')}</select></div>
      <div><label class="l">Fin</label><select ${dis} onchange="EXSIM.dSet('fin',this.value,true)">${opts.map(h => `<option ${h === D.fin ? 'selected' : ''}>${h}</option>`).join('')}</select></div></div>
    <p class="text-xs text-gray-500 mt-1">Durée : ${h2m(D.fin) > h2m(D.debut) ? m2h(h2m(D.fin) - h2m(D.debut)).replace(':', ' h ') : '—'} · ${n} étudiant(s) · ${capTot} place(s) choisie(s) ${capTot && capTot < n ? '<b style="color:#b91c1c">— places insuffisantes</b>' : ''}</p>
    <label class="l">Salles (libres sur ce créneau)</label>${sallesHtml || '<p class="text-xs text-gray-400">Aucune salle cochée dans la simulation.</p>'}
    <label class="l">Remarque</label><input type="text" ${dis} value="${esc(D.note || '')}" oninput="EXSIM.dSet('note',this.value)" placeholder="facultatif (calculatrice autorisée, tiers-temps…)">
    ${err.length ? `<div class="ex-err">${err.map(esc).join('<br>')}</div>` : ''}
    ${ro ? '' : `<div class="flex gap-2 flex-wrap mt-3"><button class="ex-b ora" ${err.length ? 'disabled' : ''} onclick="EXSIM.dValider()">✔ ${ex ? 'Mettre à jour' : 'Ajouter l\'épreuve'}</button>
      ${ex ? `<button class="ex-b" onclick="EXSIM.dDupliquer()">⧉ Dupliquer</button><button class="ex-b red" onclick="EXSIM.dSupprimer()">Supprimer</button>` : ''}
      <button class="ex-b" onclick="EXSIM.dAnnuler()">Annuler</button></div>`}</div>`;
}
const majForm = () => { const f = $('exForm'); if (f) f.innerHTML = vForm(); };
function dSet(k, v, rer) {
  X.draft[k] = v;
  if (k === 'debut' && h2m(X.draft.fin) <= h2m(v)) X.draft.fin = m2h(Math.min(h1(), h2m(v) + X.derniereDuree));
  if (rer) rendreCorps();
}
function dSalle(id, on) { const D = X.draft; if (on) D.salles.push({ id, surv: [] }); else D.salles = D.salles.filter(sa => sa.id !== id); majForm(); }
function dSurv(sid, v, on) { const sa = X.draft.salles.find(x => x.id === sid); if (!sa) return; if (on) { if (!sa.surv.includes(v)) sa.surv.push(v); } else sa.surv = sa.surv.filter(x => x !== v); majForm(); }
function dValider() {
  const D = X.draft, err = erreursEpreuve(D); if (err.length) { majForm(); return; }
  delete D._nouveau; D.matiere = String(D.matiere || '').trim();
  const i = S().epreuves.findIndex(e => e.id === D.id);
  if (i >= 0) S().epreuves[i] = D; else S().epreuves.push(D);
  S().epreuves.sort((a, b) => a.jour - b.jour || h2m(a.debut) - h2m(b.debut));
  X.derniereDuree = h2m(D.fin) - h2m(D.debut); X.draft = null; modifie(); rendreTabs(); rendreCorps();
}
function dSupprimer() { if (!confirm('Supprimer cette épreuve ?')) return; S().epreuves = S().epreuves.filter(e => e.id !== X.draft.id); X.draft = null; modifie(); rendreTabs(); rendreCorps(); }
function dAnnuler() { X.draft = null; rendreCorps(); }
function dDupliquer() {
  const D = JSON.parse(JSON.stringify(X.draft)); D.id = uid('E'); D._nouveau = true;
  D.jour = (D.jour + 1) % S().nbJours;
  D.salles = D.salles.filter(sa => etatSalle(sa.id, D).ok).map(sa => ({ id: sa.id, surv: sa.surv.filter(v => etatSurv(v, D, sa.id).ok) }));
  X.draft = D; rendreCorps(); msg('Copie préparée sur ' + JOURS[D.jour] + ' : ajustez puis validez.');
}

/* ---------- 6. Vue globale ---------- */
function affectations(version) { // lignes {jour,a0,a1,salleId,parts:[{classe,matiere,n,epId}],surv[]}
  const s = S(), R = [], opt = version === 'optim' && s.optim && s.optim.groupes;
  const dansOptim = new Set(opt ? s.optim.epIds : []);
  s.epreuves.forEach(e => {
    if (dansOptim.has(e.id)) return;
    const pt = partage(e);
    e.salles.forEach(sa => R.push({ jour: e.jour, a0: h2m(e.debut), a1: h2m(e.fin), salleId: sa.id, parts: [{ classe: e.classe, matiere: e.matiere, n: pt[sa.id] || 0, epId: e.id, a0: h2m(e.debut), a1: h2m(e.fin) }], surv: sa.surv, note: e.note }));
  });
  if (opt) s.optim.groupes.forEach(g => R.push({ jour: g.jour, a0: g.a0, a1: g.a1, salleId: g.salleId, parts: g.parts, surv: g.surv, groupe: g.id }));
  return R.sort((a, b) => a.jour - b.jour || a.a0 - b.a0);
}
function vGlobale() {
  const s = S(), an = analyse(), D0 = h0(), D1 = h1(), span = D1 - D0, pct = m => ((m - D0) / span * 100).toFixed(3) + '%';
  const version = s.optim ? X.etatVersion : 'init', rows = affectations(version);
  const jours = X.jourVue === 'tous' ? Array.from({ length: s.nbJours }, (_, j) => j) : [+X.jourVue];
  const blocTitre = r => `${libJour(r.jour)} ${m2h(r.a0)}–${m2h(r.a1)} · Salle ${salle(r.salleId).num} · ${r.parts.map(p => p.classe + ' (' + p.n + ') ' + (p.matiere || '')).join(' + ')} · Surveillance : ${r.surv.map(x => (surv(x) || { nom: '?' }).nom).join(', ') || 'aucune'}`;
  const bloc = (r, lab) => { const c = couleur(r.parts[0].classe), multi = r.parts.length > 1;
    return `<div class="bk" style="left:${pct(r.a0)};width:calc(${((r.a1 - r.a0) / span * 100).toFixed(3)}% - 2px);background:${multi ? 'linear-gradient(90deg,' + r.parts.map(p => couleur(p.classe) + '2e').join(',') + ')' : c + '24'};border-left-color:${c}" title="${esc(blocTitre(r))}" onclick="EXSIM.editer('${r.parts[0].epId}')">${esc(lab)}</div>`; };
  const graduations = `<div class="ex-tl" style="border:0"><div></div><div style="position:relative;height:16px">${Array.from({ length: Math.floor(span / 60) + 1 }, (_, k) => D0 + k * 60).filter(m => m <= D1).map(m => `<span style="position:absolute;left:${pct(m)};font-size:10px;color:#94a3b8;font-weight:700;transform:translateX(-50%)">${m2h(m).replace(':00', 'h')}</span>`).join('')}</div></div>`;
  const tracks = `--h:${(60 / span * 100).toFixed(4)}%`;
  const jourHtml = jours.map(j => {
    const R = rows.filter(r => r.jour === j);
    const lignesSalles = s.salles.map(x => { const L = R.filter(r => r.salleId === x.id);
      return `<div class="ex-tl"><div class="lb">🏫 ${esc(x.num)} <span class="text-gray-400">· ${capSalle(x.id)}</span></div><div class="tr" style="${tracks}">${L.map(r => bloc(r, r.parts.map(p => p.classe).join(' + ') + ' · ' + r.parts.map(p => p.matiere || '').filter(Boolean).join(' / '))).join('')}</div></div>`; }).join('');
    const lignesSurv = s.surveillants.map(v => {
      const pl = v.dispo[String(j)], na = [];
      if (pl) { let t = D0; pl.slice().sort((a, b) => h2m(a[0]) - h2m(b[0])).forEach(p => { if (h2m(p[0]) > t) na.push([t, Math.min(D1, h2m(p[0]))]); t = Math.max(t, h2m(p[1])); }); if (t < D1) na.push([t, D1]); }
      const L = R.filter(r => r.surv.includes(v.id));
      return `<div class="ex-tl"><div class="lb">👁 ${esc(v.nom)}</div><div class="tr" style="${tracks}">${na.filter(z => z[1] > z[0]).map(z => `<div class="na" style="left:${pct(Math.max(D0, z[0]))};width:${((Math.min(D1, z[1]) - Math.max(D0, z[0])) / span * 100).toFixed(3)}%" title="Non disponible"></div>`).join('')}${L.map(r => bloc(r, 'Salle ' + salle(r.salleId).num + ' · ' + r.parts.map(p => p.classe).join(' + '))).join('')}</div></div>`; }).join('');
    return `<div class="ex-card"><h3 style="color:#1A4E8A;font-size:13px">${esc(libJour(j))}</h3>${graduations}${lignesSalles}
      <div class="text-xs font-black uppercase text-gray-400 mt-3 mb-1">Surveillants <span class="normal-case font-semibold">(hachuré = non disponible)</span></div>${lignesSurv}</div>`;
  }).join('');
  const lst = (L, coul, titre) => L.length ? `<div class="ex-card" style="border-color:${coul}"><h3 style="color:${coul}">${titre} (${L.length})</h3><div style="max-height:220px;overflow:auto;font-size:12.5px;line-height:1.7">${L.map(c => `<div>${c.ids.length ? `<a href="#" onclick="event.preventDefault();EXSIM.editer('${c.ids[0]}')" style="color:inherit">${esc(c.txt)}</a>` : esc(c.txt)}</div>`).join('')}</div></div>` : '';
  return `${lst(an.conflits, '#dc2626', '⛔ Conflits à corriger')}${lst(an.alertes, '#b45309', '⚠ Points d\'attention')}
    ${!an.conflits.length && !an.alertes.length && s.epreuves.length ? '<div class="ex-ok">✓ Aucun conflit ni point d\'attention.</div>' : ''}
    <div class="flex gap-2 flex-wrap items-center mb-3">${['tous'].concat(Array.from({ length: s.nbJours }, (_, j) => String(j))).map(k => `<button class="ex-b sm ${String(X.jourVue) === k ? 'pri' : ''}" onclick="EXSIM.jourVue('${k}')">${k === 'tous' ? 'Toute la semaine' : esc(libJour(+k))}</button>`).join('')}
      ${s.optim ? `<span style="flex:1"></span><button class="ex-b sm ${version === 'init' ? 'pri' : ''}" onclick="EXSIM.version('init')">Avant optimisation</button><button class="ex-b sm ${version === 'optim' ? 'vio' : ''}" onclick="EXSIM.version('optim')">Après optimisation</button>` : ''}</div>
    <div class="flex gap-3 flex-wrap mb-3">${s.classes.map(c => `<span class="ex-chip"><span class="ex-dot" style="background:${couleur(c.classe)}"></span>${esc(c.classe)}</span>`).join('')}</div>
    ${s.salles.length ? jourHtml : '<div class="ex-card">Aucune salle sélectionnée.</div>'}`;
}

/* ---------- 7. États ---------- */
function etatsHtml(vue, filtre, version) {
  const s = S(), R = affectations(version), H = m => (m / 60).toFixed(1).replace('.', ',') + ' h';
  const survN = L => L.map(x => esc((surv(x) || { nom: '?' }).nom)).join(', ') || '<span style="color:#b45309">—</span>';
  const ligneJ = r => `<td>${esc(libJour(r.jour))}</td><td>${m2h(r.a0)}–${m2h(r.a1)}</td>`;
  const sec = (titre, sous, corps) => `<section class="ex-card ex-etat"><h3 style="color:#1A4E8A;font-size:14px;text-transform:none;letter-spacing:0">${titre}</h3><p class="text-xs text-gray-500" style="margin:-6px 0 8px">${sous}</p>${corps}</section>`;
  if (vue === 'salle') return s.salles.filter(x => !filtre || x.id === filtre).map(x => {
    const L = R.filter(r => r.salleId === x.id), occ = L.reduce((t, r) => t + r.a1 - r.a0, 0), dispo = (h1() - h0()) * s.nbJours;
    return sec(`🏫 Salle ${esc(x.num)} — ${esc(x.campus)}`, `${capSalle(x.id)} places · occupée ${H(occ)} sur ${H(dispo)} (${dispo ? Math.round(occ / dispo * 100) : 0} %)`,
      L.length ? `<table class="ex-tbl"><thead><tr><th>Jour</th><th>Horaires</th><th>Classe(s) · matière</th><th>Effectif</th><th>Surveillants</th></tr></thead><tbody>${L.map(r => `<tr>${ligneJ(r)}<td>${r.parts.map(p => `<span class="ex-dot" style="display:inline-block;background:${couleur(p.classe)}"></span> <b>${esc(p.classe)}</b> · ${esc(p.matiere || '')} <span class="text-gray-400">${m2h(p.a0)}–${m2h(p.a1)}</span>`).join('<br>')}</td><td>${r.parts.reduce((t, p) => t + p.n, 0)} / ${capSalle(x.id)}</td><td>${survN(r.surv)}</td></tr>`).join('')}</tbody></table>` : '<p class="text-sm text-gray-400">Salle libre toute la semaine.</p>');
  }).join('');
  if (vue === 'classe') return s.classes.filter(c => !filtre || c.classe === filtre).map(c => {
    const E = s.epreuves.filter(e => e.classe === c.classe), tot = E.reduce((t, e) => t + h2m(e.fin) - h2m(e.debut), 0);
    return sec(`<span class="ex-dot" style="display:inline-block;background:${couleur(c.classe)}"></span> ${esc(c.classe)}`, `${effectif(c.classe)} étudiant(s) · ${E.length} épreuve(s) · ${H(tot)}`,
      E.length ? `<table class="ex-tbl"><thead><tr><th>Jour</th><th>Horaires</th><th>Matière</th><th>Salle(s) · effectif</th><th>Surveillants</th></tr></thead><tbody>${E.map(e => {
        const L = R.filter(r => r.parts.some(p => p.epId === e.id));
        return `<tr><td>${esc(libJour(e.jour))}</td><td>${e.debut}–${e.fin}</td><td><b>${esc(e.matiere || 'Épreuve')}</b>${e.note ? `<div class="text-xs text-gray-500">${esc(e.note)}</div>` : ''}</td>
          <td>${L.map(r => `Salle ${esc(salle(r.salleId).num)} · ${r.parts.filter(p => p.epId === e.id).reduce((t, p) => t + p.n, 0)} étu.${r.parts.length > 1 ? ' <span class="text-gray-400">(avec ' + esc(r.parts.filter(p => p.epId !== e.id).map(p => p.classe).join(', ')) + ')</span>' : ''}`).join('<br>') || '—'}</td>
          <td>${L.map(r => survN(r.surv)).join('<br>')}</td></tr>`; }).join('')}</tbody></table>` : '<p class="text-sm text-gray-400">Aucune épreuve.</p>');
  }).join('');
  return s.surveillants.filter(v => !filtre || v.id === filtre).map(v => {
    const L = R.filter(r => r.surv.includes(v.id)), tot = L.reduce((t, r) => t + r.a1 - r.a0, 0);
    return sec(`👁 ${esc(v.nom)}`, `${H(tot)} de surveillance · disponibilités : ${Array.from({ length: s.nbJours }, (_, j) => JOURS[j].slice(0, 3) + '. ' + dispoTexte(v, j)).map(esc).join(' · ')}`,
      L.length ? `<table class="ex-tbl"><thead><tr><th>Jour</th><th>Horaires</th><th>Salle</th><th>Classe(s) · matière</th><th>Avec</th></tr></thead><tbody>${L.map(r => `<tr>${ligneJ(r)}<td><b>Salle ${esc(salle(r.salleId).num)}</b></td><td>${r.parts.map(p => esc(p.classe) + ' · ' + esc(p.matiere || '')).join('<br>')}</td><td>${survN(r.surv.filter(x => x !== v.id))}</td></tr>`).join('')}</tbody></table>` : '<p class="text-sm text-gray-400">Aucune surveillance.</p>');
  }).join('');
}
function vEtats() {
  const s = S(), version = s.optim ? X.etatVersion : 'init';
  const opts = X.etatVue === 'salle' ? s.salles.map(x => [x.id, 'Salle ' + x.num + ' — ' + x.campus]) : X.etatVue === 'classe' ? s.classes.map(c => [c.classe, c.classe]) : s.surveillants.map(v => [v.id, v.nom]);
  if (X.etatFiltre && !opts.some(o => o[0] === X.etatFiltre)) X.etatFiltre = '';
  return `<div class="flex gap-2 flex-wrap items-center mb-3">
    ${[['salle', '🏫 Par salle'], ['classe', '🎓 Par classe'], ['surv', '👁 Par surveillant']].map(([k, l]) => `<button class="ex-b ${X.etatVue === k ? 'pri' : ''}" onclick="EXSIM.etatVue('${k}')">${l}</button>`).join('')}
    <select style="max-width:260px" onchange="EXSIM.etatFiltre(this.value)"><option value="">Tous</option>${opts.map(o => `<option value="${esc(o[0])}" ${o[0] === X.etatFiltre ? 'selected' : ''}>${esc(o[1])}</option>`).join('')}</select>
    ${s.optim ? `<button class="ex-b sm ${version === 'init' ? 'pri' : ''}" onclick="EXSIM.version('init')">Avant optimisation</button><button class="ex-b sm ${version === 'optim' ? 'vio' : ''}" onclick="EXSIM.version('optim')">Après optimisation</button>` : ''}
    <span style="flex:1"></span><button class="ex-b ora" onclick="EXSIM.imprimerEtat()">🖨 Imprimer / PDF</button></div>
    ${etatsHtml(X.etatVue, X.etatFiltre, version) || '<div class="ex-card">Rien à afficher.</div>'}`;
}
function fenetre(titre, corps) {
  const w = window.open('', '_blank'); if (!w) { msg('Autorisez les fenêtres pop-up pour imprimer.', false); return; }
  w.document.write(`<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8"><title>${esc(titre)}</title><style>
    body{font-family:Arial,Helvetica,sans-serif;color:#0e1a2b;margin:18px;font-size:12px}h1{font-size:18px;color:#1A4E8A;margin:0 0 4px}.sub{color:#64748b;margin-bottom:14px}
    .ex-card{border:1px solid #cbd5e1;border-radius:10px;padding:12px;margin-bottom:12px;page-break-inside:avoid}.ex-card h3{margin:0 0 8px}
    .ex-tbl{width:100%;border-collapse:collapse}.ex-tbl th,.ex-tbl td{border:1px solid #cbd5e1;padding:4px 6px;text-align:left;vertical-align:top}.ex-tbl th{background:#f1f5f9;font-size:10px;text-transform:uppercase}
    .ex-dot{width:9px;height:9px;border-radius:2px;display:inline-block}.text-xs{font-size:10px}.text-gray-400,.text-gray-500{color:#64748b}.text-sm{font-size:12px}
    .page{page-break-after:always}.page:last-child{page-break-after:auto}
    .ex-plan{position:relative;border:3px solid #1A4E8A;border-radius:6px;margin:8px 0}
    .ex-seat{position:absolute;border-radius:5px;border:1.5px solid #94a3b8;font-size:8.5px;line-height:1.1;padding:1px 2px;overflow:hidden;border-left-width:4px;background:#fff}
    .ex-seat i{font-style:normal;color:#64748b;font-weight:bold}.ex-seat.vide{background:#f8fafc;border-style:dashed}
    @media print{body{margin:8mm}button{display:none}}@page{size:A4 landscape;margin:8mm}</style></head><body>
    <button onclick="print()" style="float:right;padding:8px 14px;font-weight:bold">🖨 Imprimer</button>${corps}</body></html>`);
  w.document.close();
}
function imprimerEtat() {
  const s = S(), version = s.optim ? X.etatVersion : 'init', lib = { salle: 'par salle', classe: 'par classe', surv: 'par surveillant' }[X.etatVue];
  fenetre(`${s.titre} — état ${lib}`, `<h1>${esc(s.titre)} — état ${lib}${version === 'optim' ? ' (après optimisation)' : ''}</h1><div class="sub">Semaine du ${esc(dateJour(0))} au ${esc(dateJour(s.nbJours - 1))} · imprimé le ${new Date().toLocaleString('fr-FR')}</div>${etatsHtml(X.etatVue, X.etatFiltre, version)}`);
}

/* ---------- terminer / rouvrir / enregistrer ---------- */
function bloquants() { const an = analyse(); return { an, L: an.conflits.concat(an.alertes.filter(a => a.type === 'sansSalle' || a.type === 'sansSurv')) }; }
async function terminer() {
  const s = S(), { an, L } = bloquants();
  if (!s.epreuves.length) { msg('Aucune épreuve planifiée.', false); return; }
  if (L.length) { msg(`Impossible de terminer : ${L.length} point(s) bloquant(s) (conflits, épreuves sans salle ou sans surveillant).`, false); onglet('globale'); return; }
  const vides = an.alertes.filter(a => a.type === 'classeVide');
  if (!confirm((vides.length ? `Attention : ${vides.length} classe(s) sans épreuve (${vides.map(v => v.txt.split(' — ')[0]).join(', ')}).\n\n` : '') + 'Terminer la semaine ? Elle passera en lecture seule et l\'optimisation sera débloquée.')) return;
  s.statut = 'terminee'; X.dirty = true;
  if (!(await enregistrer())) { s.statut = 'brouillon'; rendre(); return; }
  onglet('optim'); rendreTop();
}
async function rouvrir() {
  const s = S();
  if (!confirm('Rouvrir la semaine pour la modifier ?' + (s.optim ? '\n\nL\'optimisation et les plans de classe seront effacés (vous pourrez les recalculer après avoir terminé à nouveau).' : ''))) return;
  const sv = { statut: s.statut, optim: s.optim };
  s.statut = 'brouillon'; s.optim = null; X.prop = null; X.dirty = true;
  if (!(await enregistrer())) { Object.assign(s, sv); }
  rendre();
}
async function enregistrer() {
  if (!X.ecriture) return false;
  const s = S(); if (!String(s.titre || '').trim()) { msg('Donnez un titre à la simulation (onglet Paramètres).', false); onglet('params'); return false; }
  const b = $('exSave'); if (b) { b.disabled = true; b.textContent = 'Enregistrement…'; }
  try {
    const r = await api('exsim', { op: 'enregistrer', id: X.id || '', version: X.version, sim: JSON.stringify(s) });
    if (!r.success) { msg(r.message, false); return false; }
    X.id = r.id; X.version = r.version; X.dirty = false; msg('Simulation enregistrée ✓'); return true;
  } catch (e) { msg('Erreur réseau : ' + e.message, false); return false; }
  finally { rendreTop(); }
}

/* ---------- 8. Optimisation & plans de classe ---------- */
function sieges(salleId) { // places dans l'ordre de lecture, coordonnées en px
  const p = plan(salleId), out = [];
  if (p) {
    p.bureaux.filter(b => b.type !== 'prof').slice().sort((a, b) => a.y - b.y || a.x - b.x).forEach(b => {
      const n = (b.places || []).length || 1, w = (b.w - 8) / n;
      for (let i = 0; i < n; i++) out.push({ x: b.x + 4 + i * w, y: b.y + 16, w: w - 3, h: Math.max(26, b.h - 20) });
    });
    return { L: out, W: p.largeur || 900, H: p.hauteur || 600, plan: true };
  }
  const cap = Math.max(1, capSalle(salleId)), cols = Math.max(4, Math.round(Math.sqrt(cap * 1.6))), w = 112, h = 46;
  for (let i = 0; i < cap; i++) { const r = Math.floor(i / cols), c = i % cols; out.push({ x: 20 + c * (w + 10) + Math.floor(c / 2) * 18, y: 50 + r * (h + 22), w, h }); }
  return { L: out, W: 40 + cols * (w + 10) + Math.floor(cols / 2) * 18, H: 70 + Math.ceil(cap / cols) * (h + 22), plan: false };
}
function voisins(L) {
  return L.map((a, i) => L.map((b, j) => j).filter(j => j !== i && Math.abs((L[j].x + L[j].w / 2) - (a.x + a.w / 2)) <= a.w * 1.45 && Math.abs((L[j].y + L[j].h / 2) - (a.y + a.h / 2)) <= a.h * 1.7));
}
function placer(g, prm) {
  const G = sieges(g.salleId), L = G.L, V = voisins(L), occ = L.map(() => []), places = [];
  const parts = g.parts.map(p => ({ p, file: p.etus.slice() }));
  const total = g.parts.reduce((t, p) => t + p.n, 0);
  let ordre = L.map((_, i) => i);
  if (prm.unSurDeux && pic(g.parts, g.a0, g.a1) * 2 <= L.length) { // damier : rangées décalées, une place sur deux
    const rangs = {}; L.forEach((s, i) => { const k = Math.round(s.y / 10); (rangs[k] = rangs[k] || []).push(i); });
    const prim = []; Object.keys(rangs).sort((a, b) => a - b).forEach((k, r) => rangs[k].sort((a, b) => L[a].x - L[b].x).forEach((i, c) => { if ((c + r) % 2 === 0) prim.push(i); }));
    ordre = prim.concat(ordre.filter(i => !prim.includes(i)));
  }
  let progres = true;
  while (progres && parts.some(x => x.file.length)) {
    progres = false;
    for (const i of ordre) {
      const cands = parts.filter(x => x.file.length && !occ[i].some(o => ov(o.a0, o.a1, x.p.a0, x.p.a1)));
      if (!cands.length) continue;
      const score = x => { let mm = 0; V[i].forEach(j => occ[j].forEach(o => { if (o.classe === x.p.classe && ov(o.a0, o.a1, x.p.a0, x.p.a1)) mm++; })); return mm * 1000 - x.file.length; };
      cands.sort((a, b) => score(a) - score(b));
      const x = cands[0], e = x.file.shift();
      occ[i].push({ classe: x.p.classe, a0: x.p.a0, a1: x.p.a1 });
      places.push({ s: i, id: e.id, nom: e.nom, classe: x.p.classe, epId: x.p.epId });
      progres = true;
      if (!parts.some(y => y.file.length)) break;
    }
  }
  g.places = places; g.nonPlaces = total - places.length;
}
function pic(parts, a0, a1) { const pts = [a0].concat(parts.map(p => p.a0).filter(t => t > a0 && t < a1)); return Math.max(0, ...pts.map(t => parts.filter(p => p.a0 <= t && t < p.a1).reduce((s, p) => s + p.n, 0))); }
function optimiser(prm) {
  const s = S(), selS = new Set(prm.salles), selC = new Set(prm.classes), ratio = +prm.ratio || 25;
  const items = [], fixes = [];
  s.epreuves.forEach(e => {
    const a0 = h2m(e.debut), a1 = h2m(e.fin);
    if (selC.has(e.classe) && e.salles.length && e.salles.every(sa => selS.has(sa.id))) {
      const etus = etudiants(e.classe).map(x => ({ id: x.id, nom: x.nom }));
      const n = effectif(e.classe); while (etus.length < n) etus.push({ id: '', nom: 'Étudiant ' + (etus.length + 1) });
      items.push({ e, a0, a1, jour: e.jour, etus, n: etus.length });
    } else e.salles.forEach(sa => fixes.push({ jour: e.jour, a0, a1, salleId: sa.id, surv: sa.surv }));
  });
  const occupe = []; fixes.forEach(f => f.surv.forEach(sid => occupe.push({ sid, jour: f.jour, a0: f.a0, a1: f.a1 })));
  const groupes = [], sansPlace = [];
  for (let j = 0; j < s.nbJours; j++) {
    const its = items.filter(i => i.jour === j).sort((a, b) => a.a0 - b.a0);
    const clusters = []; let cur = null, fin = -1;
    its.forEach(i => { if (cur && i.a0 < fin) { cur.push(i); fin = Math.max(fin, i.a1); } else { cur = [i]; clusters.push(cur); fin = i.a1; } });
    clusters.forEach(C => {
      const libre = (id, a0, a1) => !fixes.some(f => f.jour === j && f.salleId === id && ov(f.a0, f.a1, a0, a1));
      const libreCap = (r, it) => {
        if (!libre(r.id, it.a0, it.a1)) return 0;
        if (!prm.melanger && r.parts.some(p => p.classe !== it.e.classe && ov(p.a0, p.a1, it.a0, it.a1))) return 0;
        return r.cap - pic(r.parts.filter(p => ov(p.a0, p.a1, it.a0, it.a1)), it.a0, it.a1);
      };
      // remplissage pour un ordre d'épreuves donné (premier qui convient, ouverture de la plus grande salle libre, puis réduction)
      const remplir = ordre => {
        const rooms = [...selS].map(id => ({ id, cap: capSalle(id), parts: [] })).filter(r => r.cap > 0).sort((a, b) => b.cap - a.cap), manques = [];
        ordre.forEach(it => {
          let rest = it.n, off = 0;
          while (rest > 0) {
            const used = rooms.filter(r => r.parts.length), unused = rooms.filter(r => !r.parts.length);
            let r = used.filter(x => libreCap(x, it) >= rest).sort((a, b) => libreCap(a, it) - libreCap(b, it))[0];
            if (!r) r = unused.filter(x => libreCap(x, it) > 0).sort((a, b) => libreCap(b, it) - libreCap(a, it))[0];
            if (!r) r = used.filter(x => libreCap(x, it) > 0).sort((a, b) => libreCap(b, it) - libreCap(a, it))[0];
            if (!r) { manques.push(`${it.e.classe} · ${it.e.matiere || 'épreuve'} (${libJour(j)}) : ${rest} étudiant(s) sans place`); break; }
            const k = Math.min(rest, libreCap(r, it));
            r.parts.push({ epId: it.e.id, classe: it.e.classe, matiere: it.e.matiere, a0: it.a0, a1: it.a1, n: k, etus: it.etus.slice(off, off + k) });
            off += k; rest -= k;
          }
        });
        // réduction : chaque salle utilisée est remplacée par une salle plus petite libre si elle suffit
        rooms.filter(r => r.parts.length).sort((a, b) => b.cap - a.cap).forEach(r => {
          const besoin = pic(r.parts, Math.min(...r.parts.map(p => p.a0)), Math.max(...r.parts.map(p => p.a1)));
          const alt = rooms.filter(x => !x.parts.length && x.cap >= besoin && x.cap < r.cap && r.parts.every(p => libre(x.id, p.a0, p.a1))).sort((a, b) => a.cap - b.cap)[0];
          if (alt) { alt.parts = r.parts; r.parts = []; }
        });
        const U = rooms.filter(r => r.parts.length);
        // score : élèves sans place, puis nombre de salles, puis places mobilisées, puis classes coupées en plusieurs salles
        return { rooms, manques, score: [manques.length, U.length, U.reduce((t, r) => t + r.cap, 0), U.reduce((t, r) => t + r.parts.length, 0)] };
      };
      // plusieurs ordres essayés (tri par effectif + mélanges reproductibles) : on garde le meilleur
      let graine = 7 + j * 101 + C.length;
      const alea = () => { graine = (graine * 16807) % 2147483647; return graine / 2147483647; };
      const ordres = [C.slice().sort((a, b) => b.n - a.n), C.slice().sort((a, b) => a.n - b.n)];
      for (let t = 0; t < (C.length > 1 ? 300 : 0); t++) { const o = C.slice(); for (let i = o.length - 1; i > 0; i--) { const k = Math.floor(alea() * (i + 1)); [o[i], o[k]] = [o[k], o[i]]; } ordres.push(o); }
      let best = null;
      const meilleur = (x, y) => { const i = x.score.findIndex((v, k) => v !== y.score[k]); return i >= 0 && x.score[i] < y.score[i]; };
      ordres.forEach(o => { const r = remplir(o); if (!best || meilleur(r, best)) best = r; });
      const rooms = best.rooms; best.manques.forEach(m => sansPlace.push(m));
      // une salle peut accueillir des épreuves successives : un groupe = parties qui se chevauchent
      rooms.filter(r => r.parts.length).forEach(r => {
        const P = r.parts.slice().sort((a, b) => a.a0 - b.a0); let gcur = null, gf = -1;
        P.forEach(p => { if (gcur && p.a0 < gf) { gcur.parts.push(p); gf = Math.max(gf, p.a1); gcur.a1 = gf; } else { gcur = { id: uid('G'), jour: j, salleId: r.id, a0: p.a0, a1: p.a1, parts: [p], surv: [] }; groupes.push(gcur); gf = p.a1; } });
      });
    });
  }
  // surveillants : on garde en priorité ceux qui surveillaient déjà ces épreuves dans cette salle
  const initial = new Set(); items.forEach(it => it.e.salles.forEach(sa => sa.surv.forEach(x => initial.add(x))));
  const libreSurv = (sid, g) => { const v = surv(sid); return v && dispoOk(v, g.jour, g.a0, g.a1) && !occupe.some(o => o.sid === sid && o.jour === g.jour && ov(o.a0, o.a1, g.a0, g.a1)); };
  groupes.slice().sort((a, b) => pic(b.parts, b.a0, b.a1) - pic(a.parts, a.a0, a.a1)).forEach(g => {
    g.pic = pic(g.parts, g.a0, g.a1);
    const besoin = Math.max(1, Math.ceil(g.pic / ratio));
    const eps = g.parts.map(p => s.epreuves.find(e => e.id === p.epId));
    const ordre = [].concat(...eps.map(e => (e.salles.find(sa => sa.id === g.salleId) || { surv: [] }).surv), ...eps.map(e => [].concat(...e.salles.map(sa => sa.surv))), [...initial], s.surveillants.map(v => v.id));
    for (const sid of ordre) { if (g.surv.length >= besoin) break; if (!g.surv.includes(sid) && libreSurv(sid, g)) { g.surv.push(sid); occupe.push({ sid, jour: g.jour, a0: g.a0, a1: g.a1 }); } }
    g.manque = besoin - g.surv.length;
  });
  groupes.forEach(g => placer(g, prm));
  groupes.sort((a, b) => a.jour - b.jour || a.a0 - b.a0 || salle(a.salleId).num.localeCompare(salle(b.salleId).num, 'fr', { numeric: true }));
  // bilan avant / après
  const avant = { salles: 0, hSalle: 0, hSurv: 0 }, apres = { salles: groupes.length, hSalle: 0, hSurv: 0 };
  items.forEach(it => it.e.salles.forEach(sa => { avant.salles++; avant.hSalle += it.a1 - it.a0; avant.hSurv += (it.a1 - it.a0) * sa.surv.length; }));
  groupes.forEach(g => { apres.hSalle += g.a1 - g.a0; apres.hSurv += (g.a1 - g.a0) * g.surv.length; });
  const gardes = new Set([].concat(...groupes.map(g => g.surv)));
  return { params: prm, date: new Date().toISOString(), epIds: items.map(i => i.e.id), groupes: groupes.map(g => ({ id: g.id, jour: g.jour, salleId: g.salleId, a0: g.a0, a1: g.a1, pic: g.pic, surv: g.surv, manque: g.manque, nonPlaces: g.nonPlaces,
    parts: g.parts.map(p => ({ epId: p.epId, classe: p.classe, matiere: p.matiere, a0: p.a0, a1: p.a1, n: p.n })), places: g.places })),
    liberes: [...initial].filter(x => !gardes.has(x)), avant, apres, sansPlace };
}
function vOptim() {
  const s = S();
  if (s.statut !== 'terminee') {
    const { L } = bloquants();
    return `<div class="ex-card" style="max-width:820px"><h3>🔒 Optimisation verrouillée</h3><p style="font-size:13px;line-height:1.7">L'optimisation se débloque lorsque la semaine est <b>terminée</b> : toutes les épreuves placées, chacune avec sa ou ses salles et au moins un surveillant par salle, sans aucun conflit.</p>
      <ul style="font-size:13px;line-height:1.8;margin:8px 0;padding-left:18px"><li>${s.epreuves.length ? '✅' : '⬜'} ${s.epreuves.length} épreuve(s) planifiée(s)</li><li>${L.length ? '⬜' : '✅'} ${L.length ? L.length + ' point(s) bloquant(s) à corriger (voir la vue globale)' : 'Aucun point bloquant'}</li></ul>
      ${X.ecriture ? `<button class="ex-b" style="color:#047857" onclick="EXSIM.terminer()">✅ Terminer la semaine</button>` : ''}</div>`;
  }
  const P = X.prmOptim = X.prmOptim || { salles: s.salles.map(x => x.id), classes: s.classes.map(c => c.classe), ratio: +s.ratio || 25, melanger: true, unSurDeux: true };
  const ro = !X.ecriture;
  const choix = s.optim ? '' : `<div class="ex-card"><h3>Périmètre de l'optimisation (toute la semaine)</h3>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px">
      <div><label class="l">Salles à optimiser</label>${s.salles.map(x => `<label style="display:flex;gap:6px;align-items:center;font-size:12.5px;margin:3px 0"><input type="checkbox" style="width:auto" ${P.salles.includes(x.id) ? 'checked' : ''} onchange="EXSIM.prm('salles','${esc(x.id)}',this.checked)"> Salle ${esc(x.num)} <span class="text-gray-400">· ${capSalle(x.id)} pl.</span></label>`).join('')}</div>
      <div><label class="l">Classes à optimiser</label>${s.classes.map(c => `<label style="display:flex;gap:6px;align-items:center;font-size:12.5px;margin:3px 0"><input type="checkbox" style="width:auto" ${P.classes.includes(c.classe) ? 'checked' : ''} onchange="EXSIM.prm('classes','${esc(c.classe)}',this.checked)"><span class="ex-dot" style="background:${couleur(c.classe)}"></span> ${esc(c.classe)} <span class="text-gray-400">· ${effectif(c.classe)}</span></label>`).join('')}</div></div>
    <div class="flex flex-wrap gap-4 items-center mt-3" style="font-size:12.5px;font-weight:600">
      <label style="display:flex;gap:6px;align-items:center"><input type="checkbox" style="width:auto" ${P.melanger ? 'checked' : ''} onchange="EXSIM.prm('melanger',null,this.checked)"> Regrouper plusieurs classes dans une même salle (mélange nominatif)</label>
      <label style="display:flex;gap:6px;align-items:center"><input type="checkbox" style="width:auto" ${P.unSurDeux ? 'checked' : ''} onchange="EXSIM.prm('unSurDeux',null,this.checked)"> Une place sur deux quand la salle le permet</label>
      <label style="display:flex;gap:6px;align-items:center">Encadrement : 1 surveillant pour <input type="number" min="5" max="80" value="${P.ratio}" style="width:70px" onchange="EXSIM.prm('ratio',null,+this.value)"> étudiants</label></div>
    <p class="text-xs text-gray-400 mt-2">Seules les épreuves dont la classe et toutes les salles sont cochées sont optimisées ; les autres restent telles quelles et leurs salles restent bloquées sur leurs créneaux. Les étudiants d'une classe ne sont jamais mis côte à côte quand une autre classe partage la salle.</p>
    ${ro ? '' : `<button class="ex-b vio mt-3" onclick="EXSIM.calculer()">✨ Calculer l'optimisation</button>`}</div>`;
  const R = s.optim || X.prop;
  if (!R) return choix;
  const H = m => (m / 60).toFixed(1).replace('.', ',') + ' h';
  const kpi = `<div class="ex-kpi"><div><b>${R.avant.salles} → ${R.apres.salles}</b><span>salles·créneaux</span></div><div><b>${H(R.avant.hSalle)} → ${H(R.apres.hSalle)}</b><span>heures de salle occupées</span></div>
    <div><b>${H(R.avant.hSurv)} → ${H(R.apres.hSurv)}</b><span>heures de surveillance</span></div><div><b>${R.liberes.length}</b><span>surveillant(s) entièrement libéré(s)</span></div></div>`;
  const alertes = [].concat(R.sansPlace, R.groupes.filter(g => g.manque > 0).map(g => `${libJour(g.jour)} ${m2h(g.a0)}–${m2h(g.a1)} salle ${salle(g.salleId).num} : il manque ${g.manque} surveillant(s) disponible(s)`), R.groupes.filter(g => g.nonPlaces > 0).map(g => `Salle ${salle(g.salleId).num} ${libJour(g.jour)} : ${g.nonPlaces} étudiant(s) sans place sur le plan`));
  const tableau = `<table class="ex-tbl"><thead><tr><th>Jour</th><th>Horaires</th><th>Salle</th><th>Classes regroupées</th><th>Occupation</th><th>Surveillants</th><th></th></tr></thead><tbody>${R.groupes.map((g, k) => `<tr>
    <td>${esc(libJour(g.jour))}</td><td>${m2h(g.a0)}–${m2h(g.a1)}</td><td><b>Salle ${esc(salle(g.salleId).num)}</b></td>
    <td>${g.parts.map(p => `<span class="ex-dot" style="display:inline-block;background:${couleur(p.classe)}"></span> ${esc(p.classe)} (${p.n}) · ${esc(p.matiere || '')} <span class="text-gray-400">${m2h(p.a0)}–${m2h(p.a1)}</span>`).join('<br>')}</td>
    <td>${g.pic} / ${capSalle(g.salleId)}</td><td>${g.surv.map(x => esc((surv(x) || { nom: '?' }).nom)).join(', ')}${g.manque > 0 ? ` <b style="color:#dc2626">(manque ${g.manque})</b>` : ''}</td>
    <td>${s.optim ? `<button class="ex-b sm" onclick="EXSIM.voirPlan(${k})">🪑 Plan</button>` : ''}</td></tr>`).join('')}</tbody></table>`;
  if (!s.optim) return choix + `<div class="ex-card" style="border-color:#6d28d9"><h3 style="color:#6d28d9">Proposition d'optimisation</h3>${kpi}
    ${alertes.length ? `<div class="ex-err">${alertes.map(esc).join('<br>')}</div>` : ''}
    ${R.liberes.length ? `<p class="text-sm mb-2"><b>Surveillants libérés :</b> ${R.liberes.map(x => esc((surv(x) || { nom: x }).nom)).join(', ')}</p>` : ''}${tableau}
    ${ro ? '' : `<div class="flex gap-2 mt-3"><button class="ex-b vio" onclick="EXSIM.appliquer()">✔ Appliquer et générer les plans de classe</button><button class="ex-b" onclick="EXSIM.abandonProp()">Abandonner</button></div>`}</div>`;
  const pl = X.planSel != null && s.optim.groupes[X.planSel] ? vPlan(s.optim.groupes[X.planSel]) : '';
  return `<div class="ex-card" style="border-color:#6d28d9"><h3 style="color:#6d28d9">✨ Optimisation appliquée <span class="normal-case font-semibold text-gray-400">— le ${esc(new Date(s.optim.date).toLocaleString('fr-FR'))}</span></h3>${kpi}
    ${alertes.length ? `<div class="ex-err">${alertes.map(esc).join('<br>')}</div>` : ''}
    <div class="flex gap-2 flex-wrap mb-3"><button class="ex-b ora" onclick="EXSIM.imprimerPlans()">🖨 Tous les plans de classe</button><button class="ex-b" onclick="EXSIM.imprimerRepartition()">🖨 Répartition nominative par classe</button><button class="ex-b" onclick="EXSIM.imprimerEmargement()">🖨 Listes d'émargement par salle</button>
      ${ro ? '' : '<span style="flex:1"></span><button class="ex-b red" onclick="EXSIM.retirerOptim()">Retirer l\'optimisation</button>'}</div>${tableau}</div>${pl}`;
}
function prm(k, v, on) {
  const P = X.prmOptim;
  if (k === 'salles' || k === 'classes') { P[k] = P[k].filter(x => x !== v); if (on) P[k].push(v); } else P[k] = on;
  X.prop = null; rendreCorps();
}
function calculer() {
  const P = X.prmOptim; if (!P.salles.length || !P.classes.length) { msg('Cochez au moins une salle et une classe.', false); return; }
  X.prop = optimiser(JSON.parse(JSON.stringify(P)));
  if (!X.prop.groupes.length) msg('Aucune épreuve dans ce périmètre.', false);
  rendreCorps();
}
function abandonProp() { X.prop = null; rendreCorps(); }
async function appliquer() {
  if (!X.prop) return; S().optim = X.prop; X.prop = null; X.dirty = true;
  if (!(await enregistrer())) { S().optim = null; rendreCorps(); return; }
  X.etatVersion = 'optim'; rendre();
}
async function retirerOptim() {
  if (!confirm('Retirer l\'optimisation et les plans de classe ? Les emplois du temps des classes ne changent pas.')) return;
  const o = S().optim; S().optim = null; X.planSel = null; X.dirty = true;
  if (!(await enregistrer())) S().optim = o;
  rendre();
}
function planHtml(g, echelle) {
  const G = sieges(g.salleId), occ = {}; (g.places || []).forEach(p => (occ[p.s] = occ[p.s] || []).push(p));
  const k = echelle || Math.min(1, 1000 / G.W);
  const court = n => { const t = String(n || '').split(' '); return esc(t.slice(1).join(' ') + ' ' + (t[0] || '').charAt(0) + '.'); };
  return `<div class="ex-plan" style="width:${G.W * k}px;height:${G.H * k}px"><div style="position:absolute;left:50%;top:4px;transform:translateX(-50%);font-size:10px;font-weight:800;letter-spacing:.12em;color:#94a3b8">TABLEAU</div>
    ${G.L.map((st, i) => { const o = occ[i] || [], c = o.length ? couleur(o[0].classe) : '#cbd5e1';
      return `<div class="ex-seat ${o.length ? '' : 'vide'}" style="left:${st.x * k}px;top:${st.y * k}px;width:${st.w * k}px;height:${st.h * k}px;border-left-color:${c}"><i>${i + 1}</i> ${o.map(p => `<b>${court(p.nom)}</b><br><span style="color:#64748b">${esc(p.classe)}</span>`).join('<br>')}</div>`; }).join('')}</div>`;
}
function legendeGroupe(g) { return g.parts.map(p => `<span class="ex-chip"><span class="ex-dot" style="background:${couleur(p.classe)}"></span>${esc(p.classe)} · ${esc(p.matiere || '')} · ${m2h(p.a0)}–${m2h(p.a1)} · ${p.n} étu.</span>`).join(' '); }
function vPlan(g) {
  const L = (g.places || []).slice().sort((a, b) => a.nom.localeCompare(b.nom, 'fr'));
  return `<div class="ex-card"><div class="flex items-center gap-2 flex-wrap"><h3 style="margin:0;color:#1A4E8A;font-size:14px;text-transform:none;letter-spacing:0">🪑 Plan de classe — Salle ${esc(salle(g.salleId).num)} · ${esc(libJour(g.jour))} ${m2h(g.a0)}–${m2h(g.a1)}</h3><span style="flex:1"></span>
    <button class="ex-b sm" onclick="EXSIM.imprimerPlans(${S().optim.groupes.indexOf(g)})">🖨 Imprimer ce plan</button><button class="ex-b sm" onclick="EXSIM.voirPlan(null)">Fermer</button></div>
    <div class="mt-2">${legendeGroupe(g)}</div>${sieges(g.salleId).plan ? '' : '<p class="text-xs text-gray-400 mt-1">Grille automatique (aucun plan enregistré pour cette salle).</p>'}
    <div style="overflow:auto">${planHtml(g)}</div>
    <details><summary class="text-sm font-bold cursor-pointer">Liste alphabétique (${L.length})</summary><table class="ex-tbl mt-2"><thead><tr><th>Étudiant</th><th>Classe</th><th>Place</th></tr></thead><tbody>${L.map(p => `<tr><td>${esc(p.nom)}</td><td>${esc(p.classe)}</td><td><b>${p.s + 1}</b></td></tr>`).join('')}</tbody></table></details></div>`;
}
function voirPlan(k) { X.planSel = k; rendreCorps(); if (k != null) setTimeout(() => { const b = $('exBody'); b.scrollTop = b.scrollHeight; }, 30); }
function imprimerPlans(k) {
  const s = S(), G = k != null ? [s.optim.groupes[k]] : s.optim.groupes;
  fenetre(`${s.titre} — plans de classe`, G.map(g => {
    const L = (g.places || []).slice().sort((a, b) => a.nom.localeCompare(b.nom, 'fr'));
    return `<div class="page"><h1>Salle ${esc(salle(g.salleId).num)} — ${esc(libJour(g.jour))} · ${m2h(g.a0)}–${m2h(g.a1)}</h1>
      <div class="sub">${esc(s.titre)} · ${g.parts.map(p => esc(p.classe) + ' — ' + esc(p.matiere || '') + ' (' + m2h(p.a0) + '–' + m2h(p.a1) + ', ' + p.n + ' étu.)').join(' · ')} · Surveillance : ${g.surv.map(x => esc((surv(x) || { nom: '?' }).nom)).join(', ')}</div>
      ${planHtml(g, Math.min(1, 1000 / sieges(g.salleId).W))}
      <table class="ex-tbl" style="margin-top:8px;font-size:10px;columns:2"><thead><tr><th>Étudiant</th><th>Classe</th><th>Place</th></tr></thead><tbody>${L.map(p => `<tr><td>${esc(p.nom)}</td><td>${esc(p.classe)}</td><td><b>${p.s + 1}</b></td></tr>`).join('')}</tbody></table></div>`;
  }).join(''));
}
function imprimerRepartition() {
  const s = S(), O = s.optim;
  const corps = s.classes.map(c => {
    const E = s.epreuves.filter(e => e.classe === c.classe); if (!E.length) return '';
    return E.map(e => {
      const G = O.groupes.filter(g => g.parts.some(p => p.epId === e.id));
      const L = G.length ? [].concat(...G.map(g => (g.places || []).filter(p => p.epId === e.id).map(p => ({ nom: p.nom, salle: salle(g.salleId).num, place: p.s + 1 })))) : [];
      L.sort((a, b) => a.nom.localeCompare(b.nom, 'fr'));
      const sallesInit = !G.length ? e.salles.map(sa => 'Salle ' + salle(sa.id).num).join(', ') : '';
      return `<div class="page"><h1>${esc(c.classe)} — ${esc(e.matiere || 'Épreuve')}</h1><div class="sub">${esc(libJour(e.jour))} · ${e.debut}–${e.fin} · ${esc(s.titre)}</div>
        ${L.length ? `<table class="ex-tbl"><thead><tr><th>Étudiant</th><th>Salle</th><th>Place</th></tr></thead><tbody>${L.map(p => `<tr><td>${esc(p.nom)}</td><td>Salle ${esc(p.salle)}</td><td><b>${p.place}</b></td></tr>`).join('')}</tbody></table>` : `<p>Épreuve non optimisée : ${esc(sallesInit || '—')}</p>`}</div>`;
    }).join('');
  }).join('');
  fenetre(`${s.titre} — répartition nominative`, corps);
}
function imprimerEmargement() {
  const s = S();
  fenetre(`${s.titre} — émargement`, s.optim.groupes.map(g => {
    const L = (g.places || []).slice().sort((a, b) => a.classe.localeCompare(b.classe, 'fr') || a.nom.localeCompare(b.nom, 'fr'));
    return `<div class="page"><h1>Émargement — Salle ${esc(salle(g.salleId).num)} · ${esc(libJour(g.jour))} ${m2h(g.a0)}–${m2h(g.a1)}</h1><div class="sub">${g.parts.map(p => esc(p.classe) + ' — ' + esc(p.matiere || '')).join(' · ')} · Surveillance : ${g.surv.map(x => esc((surv(x) || { nom: '?' }).nom)).join(', ')}</div>
      <table class="ex-tbl"><thead><tr><th>Place</th><th>Étudiant</th><th>Classe</th><th style="width:28%">Signature arrivée</th><th style="width:16%">Heure de sortie</th></tr></thead><tbody>${L.map(p => `<tr style="height:26px"><td><b>${p.s + 1}</b></td><td>${esc(p.nom)}</td><td>${esc(p.classe)}</td><td></td><td></td></tr>`).join('')}</tbody></table></div>`;
  }).join(''));
}

window.EXSIM = { _t: { optimiser: p => optimiser(p), analyse: () => analyse(), X }, liste, nouveau, ouvrir, dupliquer, supprimer, fermer: () => fermerUI(false), onglet, enregistrer, terminer, rouvrir,
  set, setLundi, setJours, ajoutSurv, retirerSurv, dispoEdit, dispo, filtreClasse, classe, toutesClasses, salle: salleCoche, campus,
  choixClasse, clicCal, editer, dSet, dSalle, dSurv, dValider, dSupprimer, dAnnuler, dDupliquer,
  jourVue: k => { X.jourVue = k; rendreCorps(); }, version: v => { X.etatVersion = v; rendreCorps(); },
  etatVue: k => { X.etatVue = k; X.etatFiltre = ''; rendreCorps(); }, etatFiltre: v => { X.etatFiltre = v; rendreCorps(); }, imprimerEtat,
  prm, calculer, abandonProp, appliquer, retirerOptim, voirPlan, imprimerPlans, imprimerRepartition, imprimerEmargement };
})();
