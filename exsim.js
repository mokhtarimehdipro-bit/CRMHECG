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
// « Prénom NOM » → « NOM Prénom » (les mots tout en majuscules forment le nom)
const nomTri = n => { const w = String(n || '').trim().split(/\s+/), up = w.filter(x => /[A-ZÀ-Þ]/.test(x) && x === x.toUpperCase()), lo = w.filter(x => !up.includes(x)); return (up.length ? up.join(' ') + ' ' + lo.join(' ') : w.join(' ')).trim(); };
const norm = s => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase();
const PPM = 1.35; // pixels par minute dans l'EDT

const X = { ref: null, id: null, version: '', S: null, ecriture: false, dirty: false, onglet: 'params', classeSel: null, draft: null,
  dispoEdit: null, jourVue: 'tous', gVue: 'chrono', etatVue: 'salle', etatFiltre: '', etatVersion: 'optim', prop: null, planSel: null, derniereDuree: 120 };

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
// Matières propres aux simulations (absentes du référentiel des notes)
const MATIERES_SIM = {
  GPME: ['CEJM', 'Anglais', 'CGE', 'GRH', 'GDR', 'SFDT', 'GRCF'],
  BACHELOR: ['Bloc 1', 'Bloc 2', 'Bloc 3', 'Bloc 4']
};
function matieresDe(c) {
  const out = new Set((X.ref.matieresExb || {})[c] || []);
  const ref = (typeof refMatieres !== 'undefined' && refMatieres) || window.refMatieres || {}, nc = norm(c);
  const gpme = nc.includes('GPME'), bach = nc.includes('BACH');
  if (gpme) MATIERES_SIM.GPME.forEach(m => out.add(m));
  if (bach) MATIERES_SIM.BACHELOR.forEach(m => out.add(m));
  Object.keys(ref).forEach(k => {
    const ok = (k === 'BTS CG' && nc.includes('BTS') && !gpme) || (k === 'DSCG' && nc.includes('DSCG')) || (k === 'DCG' && nc.includes('DCG') && !nc.includes('DSCG')) || nc.includes(norm(k));
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
    if (sa) return { ok: false, r: `déjà en salle ${salle(sa.id).num} (${e.classe} ${e.debut}–${e.fin})` };
  }
  return { ok: true };
}
function erreursEpreuve(D) {
  const err = [], a0 = h2m(D.debut), a1 = h2m(D.fin);
  if (!(a1 > a0)) err.push('L\'heure de fin doit être après l\'heure de début.');
  if (a0 < h0() || a1 > h1()) err.push(`L'épreuve sort de la plage horaire de la simulation (${S().hDebut}–${S().hFin}).`);
  if (D.jour >= S().nbJours) err.push('Jour hors de la semaine.');
  D.salles.forEach(sa => {
    if (!S().salles.some(s => s.id === sa.id)) err.push(`Salle ${salle(sa.id).num} : retirée de la simulation.`);
    const st = etatSalle(sa.id, D); if (!st.ok) err.push(`Salle ${salle(sa.id).num} : ${st.r}.`);
    sa.surv.forEach(sid => { const t = etatSurv(sid, D, sa.id); if (!t.ok) err.push(`${surv(sid) ? surv(sid).nom : sid} : ${t.r}.`); });
  });
  return err;
}
// Même classe sur deux épreuves en même temps : autorisé (oraux sur plusieurs créneaux) mais signalé
function simultanees(D) { const a0 = h2m(D.debut), a1 = h2m(D.fin); return S().epreuves.filter(e => e.id !== D.id && e.classe === D.classe && e.jour === D.jour && ov(a0, a1, h2m(e.debut), h2m(e.fin))); }
function alertesClasse(D) { return simultanees(D).map(e => `La classe a déjà « ${e.matiere || 'une épreuve'} » de ${e.debut} à ${e.fin} : deux épreuves en même temps (normal pour des oraux en plusieurs créneaux, sinon à corriger).`); }
function analyse() {
  const C = [], A = [], parEp = {};
  const add = (L, txt, ids, type) => { L.push({ txt, ids: ids || [], type }); (ids || []).forEach(i => { (parEp[i] = parEp[i] || { c: 0, a: 0 })[L === C ? 'c' : 'a']++; }); };
  const E = S().epreuves;
  E.forEach(ep => {
    const lab = `${ep.classe} · ${ep.matiere || 'épreuve'} (${libJour(ep.jour)} ${ep.debut}–${ep.fin})`;
    erreursEpreuve(ep).forEach(t => add(C, lab + ' — ' + t, [ep.id], 'conflit'));
    if (!ep.salles.length) add(A, lab + ' — aucune salle', [ep.id], 'sansSalle');
    ep.salles.forEach(sa => { if (!sa.surv.length) add(A, `${lab} — salle ${salle(sa.id).num} sans surveillant`, [ep.id], 'sansSurv'); });
    const cap = ep.salles.reduce((t, sa) => t + capSalle(sa.id), 0), n = effectif(ep.classe);
    if (ep.salles.length && cap < n) add(A, `${lab} — ${n} étudiants pour ${cap} places`, [ep.id], 'capacite');
    simultanees(ep).filter(e => e.id > ep.id).forEach(e => add(A, `${ep.classe} — deux épreuves en même temps le ${libJour(ep.jour)} : « ${ep.matiere || 'épreuve'} » ${ep.debut}–${ep.fin} et « ${e.matiere || 'épreuve'} » ${e.debut}–${e.fin}`, [ep.id, e.id], 'classeSimult'));
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
  .ex-ov{position:fixed;inset:0;z-index:9990;background:#f3f6fb;display:flex;flex-direction:column;font-family:'Plus Jakarta Sans',Inter,system-ui,sans-serif;color:#0e1a2b;font-size:14px}
  .ex-top{display:flex;flex-wrap:wrap;align-items:center;gap:14px;padding:20px 40px;background:#fff;border-bottom:1px solid #e6ecf4}
  .ex-top h2{margin:0;font-size:21px;font-weight:800;color:#1A4E8A;letter-spacing:-.01em}.ex-top h2 small{display:block;margin-top:5px;font-size:12.5px;color:#7d8a9c;font-weight:600}
  .ex-badge{font-size:12px;font-weight:800;padding:7px 13px;border-radius:999px;white-space:nowrap}
  .ex-b{height:40px;padding:0 16px;border-radius:12px;border:1px solid #dfe6f0;background:#fff;font-weight:700;font-size:13px;font-family:inherit;color:#334155;cursor:pointer;white-space:nowrap;display:inline-flex;align-items:center;gap:8px;transition:background .15s,box-shadow .15s}
  .ex-b:hover{background:#f5f8fc;box-shadow:0 6px 16px -10px rgba(14,26,43,.35)}
  .ex-b.pri{background:#1A4E8A;border-color:#1A4E8A;color:#fff}.ex-b.ora{background:#F26522;border-color:#F26522;color:#fff}.ex-b.vio{background:#6d28d9;border-color:#6d28d9;color:#fff}
  .ex-b.pri:hover,.ex-b.ora:hover,.ex-b.vio:hover{filter:brightness(1.07)}
  .ex-b.red{color:#dc2626}.ex-b.sm{height:32px;padding:0 12px;font-size:12px;border-radius:10px}.ex-b:disabled{opacity:.45;cursor:not-allowed;box-shadow:none}
  .ex-tabs{display:flex;gap:6px;padding:12px 40px;background:#fff;border-bottom:1px solid #e6ecf4;overflow-x:auto}
  .ex-tab{display:inline-flex;align-items:center;gap:9px;padding:8px 16px 8px 8px;border:0;border-radius:999px;background:none;font-weight:700;font-size:13px;font-family:inherit;color:#64748b;cursor:pointer;white-space:nowrap;transition:background .15s}
  .ex-tab:hover{background:#f3f6fb}
  .ex-tab .k{width:26px;height:26px;border-radius:50%;background:#eef2f8;color:#64748b;display:inline-flex;align-items:center;justify-content:center;font-size:12px;font-weight:800}
  .ex-tab.on{background:#1A4E8A;color:#fff}.ex-tab.on .k{background:#F26522;color:#fff}
  .ex-tab .n{font-size:11px;font-weight:800;opacity:.65}
  .ex-body{flex:1;overflow:auto;padding:36px 40px 70px}
  .ex-body>*{max-width:1440px;margin-left:auto;margin-right:auto}
  .ex-card{background:#fff;border:1px solid #e6ecf4;border-radius:20px;padding:26px 28px;margin-bottom:24px;box-shadow:0 1px 2px rgba(14,26,43,.04),0 14px 34px -26px rgba(14,26,43,.35)}
  .ex-card h3{margin:0 0 18px;font-size:12px;font-weight:800;letter-spacing:.08em;text-transform:uppercase;color:#8a97a8}
  .ex-body label.l{display:block;font-size:11px;font-weight:800;letter-spacing:.07em;text-transform:uppercase;color:#94a3b8;margin:20px 0 8px}
  .ex-body input[type=text],.ex-body input[type=date],.ex-body input[type=number],.ex-body select,.ex-body textarea{width:100%;height:44px;border:1px solid #dfe6f0;border-radius:12px;padding:0 14px;font-weight:600;font-size:14px;font-family:inherit;background:#fff;color:#0e1a2b;transition:border-color .15s,box-shadow .15s}
  .ex-body input:focus,.ex-body select:focus{outline:none;border-color:#1A4E8A;box-shadow:0 0 0 4px rgba(26,78,138,.12)}
  .ex-aide{font-size:13px;color:#64748b;line-height:1.65;margin:0 0 16px}
  .ex-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(240px,1fr));gap:16px 18px}
  .ex-pas{display:flex;flex-wrap:wrap;gap:10px}
  .ex-pill{height:40px;padding:0 18px;border-radius:999px;border:1px solid #dfe6f0;background:#fff;font-weight:700;font-size:13px;font-family:inherit;color:#334155;cursor:pointer}
  .ex-pill.on{background:#1A4E8A;border-color:#1A4E8A;color:#fff}.ex-pill:disabled{opacity:.5;cursor:not-allowed}
  .ex-add{display:grid;grid-template-columns:repeat(auto-fit,minmax(380px,1fr));gap:20px}
  .ex-addc{background:#f8fafd;border:1px solid #e6ecf4;border-radius:16px;padding:22px}
  .ex-addt{font-weight:800;color:#1A4E8A;font-size:15px;margin-bottom:6px}
  .ex-row{display:flex;gap:10px;align-items:center}.ex-row input{flex:1;min-width:0}
  .ex-sect{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:24px}
  .ex-chip{display:inline-flex;align-items:center;gap:6px;padding:6px 12px;border-radius:999px;font-size:12px;font-weight:700;background:#eef2f8;color:#334155;margin:3px 4px 3px 0}
  .ex-chip.bad{background:#fee2e2;color:#b91c1c}.ex-chip button{border:0;background:none;cursor:pointer;color:inherit;font-weight:800;padding:0 0 0 2px;font-size:14px}
  .ex-tbl{width:100%;border-collapse:separate;border-spacing:0;font-size:13.5px}
  .ex-tbl th{text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:#7d8a9c;padding:12px 14px;border-bottom:1px solid #e6ecf4;background:#f8fafd}
  .ex-tbl th:first-child{border-top-left-radius:12px}.ex-tbl th:last-child{border-top-right-radius:12px}
  .ex-tbl td{padding:14px;border-bottom:1px solid #eff3f8;vertical-align:top;line-height:1.55}
  .ex-tbl tbody tr:hover td{background:#fafcff}
  .ex-dc{cursor:pointer;border-radius:10px;padding:9px 10px;font-size:12px;font-weight:700;text-align:center;transition:transform .1s}
  .ex-dc:hover{transform:translateY(-1px)}
  .ex-dc.full{background:#ecfdf5;color:#047857}.ex-dc.part{background:#fffbeb;color:#b45309}.ex-dc.abs{background:#f1f5f9;color:#94a3b8;text-decoration:line-through}
  .ex-dc.sel{outline:2px solid #F26522;outline-offset:1px}
  .ex-lock{background:#fff7ed;border:1px solid #fed7aa;color:#9a3412;border-radius:14px;padding:14px 18px;font-size:13.5px;font-weight:600;margin-bottom:24px}
  .ex-edt{display:flex;gap:24px;align-items:flex-start}
  .ex-cls{width:190px;flex-shrink:0}.ex-cls .it{display:flex;align-items:center;gap:10px;padding:12px 14px;border-radius:14px;cursor:pointer;font-size:13px;font-weight:700;border:1px solid transparent;margin-bottom:6px;transition:background .15s}
  .ex-cls .it:hover{background:#fff}.ex-cls .it.on{background:#fff;border-color:#dfe6f0;box-shadow:0 8px 20px -14px rgba(14,26,43,.45)}
  .ex-dot{width:11px;height:11px;border-radius:4px;flex-shrink:0;display:inline-block}
  .ex-cal{flex:1;min-width:0;background:#fff;border:1px solid #e6ecf4;border-radius:20px;padding:22px 20px;overflow-x:auto;box-shadow:0 14px 34px -26px rgba(14,26,43,.35)}
  .ex-calg{display:grid;min-width:680px}
  .ex-calh{font-size:12.5px;font-weight:800;color:#1A4E8A;text-align:center;padding:4px 0 14px}
  .ex-col{position:relative;border-left:1px solid #e6ecf4;cursor:copy;transition:background-color .12s}
  .ex-col.ro{cursor:default}.ex-col.drop{background-color:#fff4ec!important}
  .ex-tcol{position:relative}.ex-tcol span{position:absolute;right:8px;font-size:10.5px;color:#94a3b8;font-weight:700;transform:translateY(-50%)}
  .ex-blk{position:absolute;left:4px;right:4px;border-radius:10px;padding:6px 8px;font-size:11.5px;line-height:1.35;overflow:hidden;cursor:grab;border-left:4px solid;box-shadow:0 6px 14px -10px rgba(14,26,43,.6);touch-action:none;user-select:none}
  .ex-blk b{display:block;font-size:12px}.ex-blk.err{outline:2px solid #dc2626}.ex-blk.warn{outline:2px dashed #f59e0b}.ex-blk.sel{outline:3px solid #F26522;z-index:3}
  .ex-blk.ghost{opacity:.45;pointer-events:none;border-left-style:dashed}.ex-blk.dragging{opacity:.3}
  .ex-ghost{position:fixed!important;z-index:10002;pointer-events:none;box-shadow:0 22px 44px -14px rgba(14,26,43,.55);opacity:.95;cursor:grabbing}
  .ex-gt{font-size:10.5px;font-weight:800;color:#F26522;margin-bottom:2px}
  .ex-form{position:fixed;top:0;right:0;bottom:0;width:460px;max-width:94vw;overflow-y:auto;z-index:9995;background:#f3f6fb;padding:16px;box-shadow:-24px 0 48px -24px rgba(14,26,43,.45);border-left:1px solid #e6ecf4}.ex-form:empty{display:none}
  .ex-dupb{position:absolute;top:4px;right:4px;width:22px;height:22px;border-radius:7px;border:1px solid #dfe6f0;background:#fff;font-size:12px;line-height:1;cursor:pointer;opacity:0;transition:opacity .15s;padding:0}.ex-blk:hover .ex-dupb{opacity:1}
  .ex-srow{border:1px solid #e6ecf4;border-radius:14px;padding:12px 14px;margin-bottom:10px;font-size:13px}
  .ex-srow.off{opacity:.55;background:#f8fafc}.ex-srow.bad{border-color:#fca5a5;background:#fef2f2}
  .ex-err{background:#fef2f2;border:1px solid #fecaca;color:#b91c1c;border-radius:12px;padding:12px 14px;font-size:12.5px;font-weight:600;margin:14px 0;line-height:1.6}
  .ex-warn{background:#fffbeb;border:1px solid #fde68a;color:#92400e;border-radius:12px;padding:12px 14px;font-size:12.5px;font-weight:600;margin:14px 0;line-height:1.6}
  .ex-ok{background:#ecfdf5;border:1px solid #a7f3d0;color:#047857;border-radius:12px;padding:12px 14px;font-size:13px;font-weight:600;margin:0 0 24px}
  .ex-tl{display:grid;grid-template-columns:200px 1fr;align-items:center;border-bottom:1px solid #f1f5f9}
  .ex-tl .lb{font-size:12.5px;font-weight:700;padding:8px 12px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
  .ex-tl .tr{position:relative;height:40px;background:repeating-linear-gradient(90deg,transparent 0,transparent calc(var(--h) - 1px),#eef2f8 calc(var(--h) - 1px),#eef2f8 var(--h))}
  .ex-tl .bk{position:absolute;top:6px;bottom:6px;border-radius:8px;font-size:11.5px;font-weight:700;padding:4px 8px;overflow:hidden;white-space:nowrap;text-overflow:ellipsis;border-left:3px solid;cursor:pointer}
  .ex-tl .na{position:absolute;top:0;bottom:0;background:repeating-linear-gradient(45deg,#e2e8f0 0,#e2e8f0 3px,transparent 3px,transparent 7px)}
  .ex-kpi{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:16px;margin-bottom:24px}
  .ex-kpi div{background:#f8fafd;border:1px solid #e6ecf4;border-radius:16px;padding:18px 20px}.ex-kpi b{display:block;font-size:24px;color:#0e1a2b}.ex-kpi span{font-size:12px;color:#7d8a9c;font-weight:700}
  .ex-plan{position:relative;background:#fff;border:3px solid #1A4E8A;border-radius:8px;margin:18px 0}
  .ex-seat{position:absolute;border-radius:6px;border:1.5px solid #cbd5e1;background:#fff;font-size:9.5px;line-height:1.2;padding:2px 4px;overflow:hidden;border-left-width:5px}
  .ex-seat i{font-style:normal;color:#94a3b8;font-weight:800}.ex-seat.vide{background:#f8fafc;border-style:dashed}
  .ex-msg{position:fixed;left:50%;bottom:26px;transform:translateX(-50%);z-index:10003;padding:12px 20px;border-radius:14px;color:#fff;font-weight:700;font-size:13.5px;box-shadow:0 14px 34px -12px rgba(0,0,0,.45);max-width:90vw}
  .ex-ov .gap-2{gap:12px}.ex-ov .mb-3{margin-bottom:20px}.ex-ov .mt-3{margin-top:20px}.ex-ov .mt-2{margin-top:12px}.ex-ov .text-xs{font-size:12px}
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
  const s = Object.assign({ titre: '', lundi: lundiProchain(), nbJours: 5, hDebut: '08:00', hFin: '18:00', pas: 15, dureeDefaut: 120, statut: 'brouillon', surveillants: [], classes: [], salles: [], epreuves: [], optim: null }, o || {});
  s.epreuves.forEach(e => { e.salles = (e.salles || []).map(sa => ({ id: sa.id, surv: sa.surv || [] })); e.jour = +e.jour || 0; });
  s.surveillants.forEach(v => { v.dispo = v.dispo || {}; });
  return s;
}
async function nouveau() {
  try { await chargerRef(true); const l = await lire({ op: 'liste' }); X.ecriture = !!(l && l.ecriture); } catch (e) { toast(e.message, false); return; }
  if (!X.ecriture) { toast('Création réservée au personnel.', false); return; }
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
const ONGLETS = [['params', 'Paramètres'], ['surv', 'Surveillants'], ['classes', 'Classes'], ['salles', 'Salles'], ['edt', 'Emplois du temps'], ['globale', 'Vue globale'], ['etats', 'États & exports'], ['optim', 'Optimisation & plans']];
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
    <span class="ex-badge" style="background:${st[2]};color:${st[1]}">${st[0]}</span>
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
  $('exTabs').innerHTML = ONGLETS.map(([k, l], i) => `<button class="ex-tab ${X.onglet === k ? 'on' : ''}" onclick="EXSIM.onglet('${k}')"><span class="k">${i + 1}</span>${l}${n[k] != null ? `<span class="n">${n[k]}</span>` : ''}${k === 'optim' && s.statut !== 'terminee' ? ' 🔒' : ''}</button>`).join('');
}
function onglet(k) { X.onglet = k; if (k !== 'edt') X.draft = null; rendreTabs(); rendreCorps(); $('exBody').scrollTop = 0; }
function rendreCorps() {
  const b = $('exBody'); if (!b) return;
  const f = { params: vParams, surv: vSurv, classes: vClasses, salles: vSalles, edt: vEdt, globale: vGlobale, etats: vEtats, optim: vOptim }[X.onglet];
  b.innerHTML = (lecture() && ['params', 'surv', 'classes', 'salles', 'edt'].includes(X.onglet)
    ? `<div class="ex-lock">${X.ecriture ? '🔒 La semaine est terminée : lecture seule. Cliquez sur « Rouvrir la semaine » pour la modifier' + (S().optim ? ' (l\'optimisation et les plans de classe seront alors effacés).' : '.') : '👁 Consultation seule.'}</div>` : '') + f();
}

/* ---------- 1. Paramètres ---------- */
function vParams() {
  const s = S(), ro = lecture() ? 'disabled' : '';
  const heures = []; for (let m = 6 * 60; m <= 22 * 60; m += 15) heures.push(m2h(m));
  const PAS = [[5, '5 min'], [10, '10 min'], [15, '15 min'], [20, '20 min'], [30, '30 min'], [45, '45 min'], [60, '1 heure']];
  const DUREES = [60, 90, 120, 150, 180, 210, 240, 270, 300, 330, 360];
  return `<div class="ex-card" style="max-width:1000px"><h3>Simulation</h3>
    <label class="l" style="margin-top:0">Titre</label><input type="text" ${ro} value="${esc(s.titre)}" oninput="EXSIM.set('titre',this.value)" placeholder="ex : Examens blancs — session 2">
    <div class="ex-grid">
      <div><label class="l">Semaine (lundi)</label><input type="date" ${ro} value="${esc(s.lundi)}" onchange="EXSIM.setLundi(this.value)"></div>
      <div><label class="l">Jours</label><select ${ro} onchange="EXSIM.setJours(+this.value)"><option value="5" ${s.nbJours === 5 ? 'selected' : ''}>Lundi → vendredi</option><option value="6" ${s.nbJours === 6 ? 'selected' : ''}>Lundi → samedi</option></select></div>
      <div><label class="l">Début de journée</label><select ${ro} onchange="EXSIM.set('hDebut',this.value,true)">${heures.map(h => `<option ${h === s.hDebut ? 'selected' : ''}>${h}</option>`).join('')}</select></div>
      <div><label class="l">Fin de journée</label><select ${ro} onchange="EXSIM.set('hFin',this.value,true)">${heures.map(h => `<option ${h === s.hFin ? 'selected' : ''}>${h}</option>`).join('')}</select></div>
    </div></div>
  <div class="ex-card" style="max-width:1000px"><h3>Horaires des épreuves</h3>
    <p class="ex-aide">Le <b>pas</b> fixe la précision des heures de début et de fin proposées, ainsi que l'aimantation lors d'un clic ou d'un glisser-déposer dans le calendrier.</p>
    <div class="ex-pas">${PAS.map(([v, l]) => `<button class="ex-pill ${+s.pas === v ? 'on' : ''}" ${ro} onclick="EXSIM.set('pas',${v},true)">${l}</button>`).join('')}</div>
    <div class="ex-grid" style="margin-top:6px"><div><label class="l">Durée proposée pour une nouvelle épreuve</label><select ${ro} onchange="EXSIM.set('dureeDefaut',+this.value)">${DUREES.map(d => `<option value="${d}" ${+s.dureeDefaut === d ? 'selected' : ''}>${Math.floor(d / 60)} h${d % 60 ? ' ' + String(d % 60).padStart(2, '0') : ''}</option>`).join('')}</select></div></div></div>
  <div class="ex-card" style="max-width:1000px"><h3>Encadrement</h3>
    <p class="ex-aide" style="margin:0">Règle appliquée : <b>un surveillant par salle</b>. Une épreuve ne peut être validée que si elle a au moins une salle, et chaque salle au moins un surveillant (vous pouvez en ajouter plusieurs).</p></div>
  <div class="ex-card" style="max-width:1000px"><h3>Comment ça marche</h3><ol style="margin:0;padding-left:20px;font-size:13.5px;line-height:2;color:#334155">
      <li><b>Surveillants</b> : ajoutez-les (prénom et nom libres, ou depuis le personnel) et réglez leurs disponibilités — par défaut, toute la semaine.</li>
      <li><b>Classes</b> et <b>Salles</b> : cochez les classes concernées et les locaux disponibles.</li>
      <li><b>Emplois du temps</b> : cliquez dans le calendrier pour créer une épreuve ; faites glisser une épreuve pour la déplacer sur un autre jour ou un autre horaire. Seules les salles libres et les surveillants disponibles sont proposés.</li>
      <li><b>Vue globale</b> et <b>États</b> : contrôlez la semaine par salle, par classe et par surveillant, puis exportez en Excel ou en PDF.</li>
      <li><b>Terminer la semaine</b> (aucun conflit) débloque l'<b>optimisation</b> : regroupement des classes, mélange nominatif des étudiants et plans de classe.</li></ol></div>`;
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
  return `${ro ? '' : `<div class="ex-card"><h3>Ajouter des surveillants</h3>
    <div class="ex-add">
      <div class="ex-addc"><div class="ex-addt">Nouveau surveillant</div><p class="ex-aide">Saisissez librement le prénom et le nom (vacataire, intervenant, parent…).</p>
        <div class="ex-row"><input type="text" id="exSPrenom" placeholder="Prénom" onkeydown="if(event.key==='Enter')EXSIM.ajoutLibre()"><input type="text" id="exSNom" placeholder="Nom" onkeydown="if(event.key==='Enter')EXSIM.ajoutLibre()"><button class="ex-b pri" onclick="EXSIM.ajoutLibre()">+ Ajouter</button></div></div>
      <div class="ex-addc"><div class="ex-addt">Depuis le personnel ou les formateurs</div><p class="ex-aide">Recherchez une personne déjà enregistrée dans le CRM.</p>
        <div class="ex-row"><input type="text" list="exPersList" id="exAddSurv" placeholder="Rechercher…" onkeydown="if(event.key==='Enter')EXSIM.ajoutSurv()"><button class="ex-b" onclick="EXSIM.ajoutSurv()">+ Ajouter</button></div>
        <datalist id="exPersList">${dispoListe.map(p => `<option value="${esc(p.nom)}">${esc(p.source)}</option>`).join('')}</datalist></div>
    </div>
    <p class="ex-aide" style="margin:18px 0 0">Tout surveillant ajouté est disponible toute la semaine ; ajustez ensuite ses disponibilités dans le tableau ci-dessous.</p></div>`}
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
  S().surveillants.sort((a, b) => nomTri(a.nom).localeCompare(nomTri(b.nom), 'fr'));
  modifie(); rendreTabs(); rendreCorps(); setTimeout(() => $('exAddSurv') && $('exAddSurv').focus(), 0);
}
function ajoutLibre() {
  const pr = String($('exSPrenom').value || '').trim(), no = String($('exSNom').value || '').trim();
  if (!pr || !no) { msg('Indiquez le prénom et le nom du surveillant.', false); return; }
  const prenom = pr.toLowerCase().replace(/(^|[\s-])\S/g, c => c.toUpperCase()), nom = prenom + ' ' + no.toUpperCase();
  const id = 'X-' + norm(nom).replace(/[^A-Z0-9]+/g, '-');
  if (surv(id) || S().surveillants.some(v => norm(v.nom) === norm(nom))) { msg(nom + ' est déjà dans la liste.', false); return; }
  S().surveillants.push({ id, nom, source: 'Ajout libre', dispo: {} });
  S().surveillants.sort((a, b) => nomTri(a.nom).localeCompare(nomTri(b.nom), 'fr'));
  modifie(); rendreTabs(); rendreCorps(); msg(nom + ' ajouté ✓'); setTimeout(() => $('exSPrenom') && $('exSPrenom').focus(), 0);
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
    const EJ = s.epreuves.filter(e => e.classe === c && e.jour === j).sort((a, b) => h2m(a.debut) - h2m(b.debut) || h2m(b.fin) - h2m(a.fin));
    const voie = {}, nbV = {}; // répartition en colonnes des épreuves qui se chevauchent
    { let grp = [], finG = -1, fins = []; const clore = () => { const n = fins.length || 1; grp.forEach(id => nbV[id] = n); grp = []; fins = []; };
      EJ.forEach(e => { const a0 = h2m(e.debut), a1 = h2m(e.fin); if (grp.length && a0 >= finG) clore();
        let v = fins.findIndex(f => f <= a0); if (v < 0) { v = fins.length; fins.push(a1); } else fins[v] = a1;
        voie[e.id] = v; grp.push(e.id); finG = Math.max(grp.length > 1 ? finG : -1, a1); });
      clore(); }
    const blocs = EJ.map(e => {
      const p = an.parEp[e.id] || {}, col = couleur(c), top = (h2m(e.debut) - h0()) * PPM, h = (h2m(e.fin) - h2m(e.debut)) * PPM;
      const nv = nbV[e.id] || 1, lr = nv > 1 ? `left:calc(${voie[e.id] * 100 / nv}% + 2px);right:auto;width:calc(${100 / nv}% - 4px);` : '';
      return `<div class="ex-blk ${p.c ? 'err' : p.a ? 'warn' : ''} ${X.draft && X.draft.id === e.id ? 'sel' : ''}" style="${lr}top:${top}px;height:${Math.max(h, 18)}px;background:${col}1a;border-left-color:${col}" ${ro ? `onclick="event.stopPropagation();EXSIM.editer('${e.id}')"` : `onpointerdown="EXSIM.glisser(event,'${e.id}')" onclick="event.stopPropagation()"`} title="${esc(e.matiere)} ${e.debut}–${e.fin}">
        ${ro ? '' : `<button class="ex-dupb" title="Dupliquer cette épreuve" onpointerdown="event.stopPropagation()" onclick="event.stopPropagation();EXSIM.dupRapide('${e.id}')">⧉</button>`}<b>${esc(e.matiere || 'Épreuve')}</b>${e.debut}–${e.fin}<br>${e.salles.map(sa => '🏫 ' + esc(salle(sa.id).num) + (sa.surv.length ? ' · ' + sa.surv.map(x => esc((surv(x) || { nom: '?' }).nom)).join(', ') : ' · <span style="color:#b45309">sans surveillant</span>')).join('<br>') || '<span style="color:#b45309">aucune salle</span>'}</div>`;
    }).join('');
    const sv = X.draft && s.epreuves.find(e => e.id === X.draft.id);
    const ghost = X.draft && X.draft.classe === c && X.draft.jour === j && (!sv || sv.jour !== X.draft.jour || sv.debut !== X.draft.debut || sv.fin !== X.draft.fin)
      ? `<div class="ex-blk ghost sel" style="top:${(h2m(X.draft.debut) - h0()) * PPM}px;height:${Math.max((h2m(X.draft.fin) - h2m(X.draft.debut)) * PPM, 18)}px;background:${couleur(c)}22;border-left-color:${couleur(c)}"><b>${sv ? 'Nouvel horaire' : 'Nouvelle épreuve'}</b>${X.draft.debut}–${X.draft.fin}</div>` : '';
    return `<div class="ex-col ${ro ? 'ro' : ''}" data-j="${j}" style="height:${H}px;background:repeating-linear-gradient(180deg,#f1f5f9 0,#f1f5f9 1px,transparent 1px,transparent ${30 * PPM}px)" ${ro ? '' : `onclick="EXSIM.clicCal(event,${j})"`}>${blocs}${ghost}</div>`;
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
  if (lecture() || Date.now() - (X.ignoreClic || 0) < 400) return;
  const r = ev.currentTarget.getBoundingClientRect(), pas = +S().pas || 15;
  let a0 = h0() + Math.floor((ev.clientY - r.top) / PPM / pas) * pas;
  const duree = +S().dureeDefaut || X.derniereDuree;
  let a1 = Math.min(h1(), a0 + duree); if (a1 - a0 < pas) { a0 = Math.max(h0(), h1() - duree); a1 = h1(); }
  X.draft = { id: uid('E'), classe: X.classeSel, jour: j, debut: m2h(a0), fin: m2h(a1), matiere: '', note: '', salles: [], _nouveau: true };
  rendreCorps();
}
function editer(id) { const e = S().epreuves.find(x => x.id === id); if (!e) return; X.draft = JSON.parse(JSON.stringify(e)); X.onglet = 'edt'; X.classeSel = e.classe; rendreTabs(); rendreCorps(); }
function vForm() {
  const D = X.draft, ro = lecture();
  if (!D) return '';
  const pas = +S().pas || 15, opts = []; for (let m = h0(); m <= h1(); m += pas) opts.push(m2h(m));
  [D.debut, D.fin].forEach(h => { if (!opts.includes(h)) opts.push(h); }); opts.sort((a, b) => h2m(a) - h2m(b));
  const dis = ro ? 'disabled' : '', n = effectif(D.classe), err = erreursEpreuve(D), part = partage(D);
  const capTot = D.salles.reduce((t, sa) => t + capSalle(sa.id), 0), manq = manquants(D);
  const mats = matieresDe(D.classe);
  const sallesHtml = S().salles.map(x => {
    const on = D.salles.find(sa => sa.id === x.id), st = etatSalle(x.id, D);
    if (!on) return `<div class="ex-srow ${st.ok ? '' : 'off'}"><label style="display:flex;gap:8px;align-items:center;${st.ok && !ro ? 'cursor:pointer' : ''}"><input type="checkbox" ${st.ok && !ro ? '' : 'disabled'} onchange="EXSIM.dSalle('${esc(x.id)}',true)" style="width:auto">
      <span><b>Salle ${esc(x.num)}</b> <span class="text-gray-400">· ${capSalle(x.id)} pl. · ${esc(x.campus)}</span>${st.ok ? '' : `<br><span class="text-xs" style="color:#b91c1c">⛔ ${esc(st.r)}</span>`}</span></label></div>`;
    const options = S().surveillants.filter(v => !on.surv.includes(v.id)).map(v => { const t = etatSurv(v.id, D, x.id); return `<option value="${esc(v.id)}" ${t.ok ? '' : 'disabled'}>${esc(v.nom)}${t.ok ? '' : ' — ' + esc(t.r)}</option>`; });
    const nbOk = S().surveillants.filter(v => !on.surv.includes(v.id) && etatSurv(v.id, D, x.id).ok).length;
    return `<div class="ex-srow ${st.ok ? '' : 'bad'}" style="border-color:${st.ok ? '#93c5fd' : ''};background:${st.ok ? '#f5f9ff' : ''}">
      <label style="display:flex;gap:8px;align-items:center"><input type="checkbox" checked ${dis} onchange="EXSIM.dSalle('${esc(x.id)}',false)" style="width:auto">
      <span style="flex:1"><b>Salle ${esc(x.num)}</b> <span class="text-gray-500">· ${part[x.id] || 0} étu. / ${capSalle(x.id)} pl.</span></span></label>
      ${st.ok ? '' : `<div class="text-xs" style="color:#b91c1c">⛔ ${esc(st.r)}</div>`}
      <div style="margin-top:4px">${on.surv.map(sid => { const t = etatSurv(sid, D, x.id); return `<span class="ex-chip ${t.ok ? '' : 'bad'}" title="${esc(t.r || '')}">👁 ${esc((surv(sid) || { nom: sid }).nom)}${t.ok ? '' : ' ⛔'}${ro ? '' : `<button onclick="EXSIM.dSurv('${esc(x.id)}','${esc(sid)}',false)">×</button>`}</span>`; }).join('')}
      ${on.surv.length ? '' : `<span class="text-xs" style="color:#b45309;font-weight:700">⚠ Un surveillant est requis pour cette salle</span>`}</div>
      ${ro ? '' : `<select style="margin-top:4px;height:32px;font-size:12px" onchange="if(this.value)EXSIM.dSurv('${esc(x.id)}',this.value,true)"><option value="">+ Ajouter un surveillant (${nbOk} disponible(s))</option>${options.join('')}</select>`}</div>`;
  }).join('');
  const ex = S().epreuves.some(e => e.id === D.id);
  return `<div class="ex-card" style="border-color:${couleur(D.classe)}"><h3 style="display:flex;align-items:center;gap:8px"><span style="flex:1">${ex ? 'Épreuve' : 'Nouvelle épreuve'} — ${esc(D.classe)}</span>${ex && !ro ? `<button class="ex-b sm" title="Dupliquer" onclick="document.getElementById('exDupBloc').scrollIntoView({behavior:'smooth'})">⧉ Dupliquer</button>` : ''}<button class="ex-b sm" title="Fermer" onclick="EXSIM.dAnnuler()">✕</button></h3>
    <label class="l" style="margin-top:0">Classe</label><select ${dis} onchange="EXSIM.dSet('classe',this.value,true)">${S().classes.map(c => `<option value="${esc(c.classe)}" ${c.classe === D.classe ? 'selected' : ''}>${esc(c.classe)} · ${effectif(c.classe)} étu.</option>`).join('')}</select>
    <label class="l">Matière / épreuve</label><input type="text" ${dis} list="exMatList" value="${esc(D.matiere)}" oninput="EXSIM.dSet('matiere',this.value)" placeholder="ex : UE 4 — Droit fiscal">
    <datalist id="exMatList">${mats.map(m => `<option value="${esc(m)}">`).join('')}</datalist>
    <div style="display:grid;grid-template-columns:1.3fr 1fr 1fr;gap:8px">
      <div><label class="l">Jour</label><select ${dis} onchange="EXSIM.dSet('jour',+this.value,true)">${Array.from({ length: S().nbJours }, (_, j) => `<option value="${j}" ${j === D.jour ? 'selected' : ''}>${esc(libJour(j))}</option>`).join('')}</select></div>
      <div><label class="l">Début</label><select ${dis} onchange="EXSIM.dSet('debut',this.value,true)">${opts.map(h => `<option ${h === D.debut ? 'selected' : ''}>${h}</option>`).join('')}</select></div>
      <div><label class="l">Fin</label><select ${dis} onchange="EXSIM.dSet('fin',this.value,true)">${opts.map(h => `<option ${h === D.fin ? 'selected' : ''}>${h}</option>`).join('')}</select></div></div>
    <p class="text-xs text-gray-500 mt-1">Durée : ${h2m(D.fin) > h2m(D.debut) ? (x => Math.floor(x / 60) + ' h' + (x % 60 ? ' ' + String(x % 60).padStart(2, '0') : ''))(h2m(D.fin) - h2m(D.debut)) : '—'} · ${n} étudiant(s) · ${capTot} place(s) choisie(s) ${capTot && capTot < n ? '<b style="color:#b91c1c">— places insuffisantes</b>' : ''}</p>
    <label class="l">Salles (libres sur ce créneau)</label>${sallesHtml || '<p class="text-xs text-gray-400">Aucune salle cochée dans la simulation.</p>'}
    <label class="l">Remarque</label><input type="text" ${dis} value="${esc(D.note || '')}" oninput="EXSIM.dSet('note',this.value)" placeholder="facultatif (calculatrice autorisée, tiers-temps…)">
    ${err.length ? `<div class="ex-err">${err.map(esc).join('<br>')}</div>` : ''}
    ${alertesClasse(D).length ? `<div class="ex-warn">⚠ ${alertesClasse(D).map(esc).join('<br>⚠ ')}</div>` : ''}
    ${!ro && manq.length ? `<div class="ex-warn">À compléter avant validation :<br>${manq.map(esc).join('<br>')}</div>` : ''}
    ${ro ? '' : `<div class="flex gap-2 flex-wrap mt-3"><button class="ex-b ora" ${err.length || manq.length ? 'disabled' : ''} onclick="EXSIM.dValider()">✔ ${ex ? 'Mettre à jour' : 'Ajouter l\'épreuve'}</button>
      ${ex ? `<button class="ex-b red" onclick="EXSIM.dSupprimer()">Supprimer</button>` : ''}
      <button class="ex-b" onclick="EXSIM.dAnnuler()">Annuler</button></div>
    ${ex ? `<div id="exDupBloc" style="margin-top:22px;padding:18px;border-radius:16px;background:#f8fafd;border:1px solid #e6ecf4">
      <div class="ex-addt" style="font-size:14px">⧉ Dupliquer cette épreuve</div>
      <p class="ex-aide" style="margin-bottom:10px">Même matière et même durée : choisissez la classe, le jour et l'heure de début, puis affectez la salle et le surveillant.</p>
      <div style="display:grid;grid-template-columns:1.4fr 1fr 0.8fr;gap:10px">
        <select id="exDupC">${S().classes.map(c => `<option value="${esc(c.classe)}" ${c.classe === dupClasseDefaut(D) ? 'selected' : ''}>${esc(c.classe)}</option>`).join('')}</select>
        <select id="exDupJ">${Array.from({ length: S().nbJours }, (_, j) => `<option value="${j}" ${j === D.jour ? 'selected' : ''}>${esc(libJour(j))}</option>`).join('')}</select>
        <select id="exDupH">${opts.filter(h => h2m(h) + h2m(D.fin) - h2m(D.debut) <= h1()).map(h => `<option ${h === D.debut ? 'selected' : ''}>${h}</option>`).join('')}</select></div>
      <button class="ex-b pri" style="margin-top:10px;width:100%;justify-content:center" onclick="EXSIM.dDupliquer()">⧉ Créer la copie</button></div>` : ''}`}</div>`;
}
function glisser(ev, id) {
  if (lecture() || ev.button !== 0) return;
  const el = ev.currentTarget, e = S().epreuves.find(x => x.id === id); if (!e) return;
  ev.preventDefault();
  const r0 = el.getBoundingClientRect(), dy = ev.clientY - r0.top, x0 = ev.clientX, y0 = ev.clientY;
  const duree = h2m(e.fin) - h2m(e.debut), pas = +S().pas || 15;
  let gh = null, lab = null, cible = null;
  const calc = (cx, cy) => {
    const col = document.elementsFromPoint(cx, cy).find(n => n.classList && n.classList.contains('ex-col'));
    if (!col) return null;
    const rc = col.getBoundingClientRect();
    let a0 = h0() + Math.round((cy - dy - rc.top) / PPM / pas) * pas;
    a0 = Math.max(h0(), Math.min(h1() - duree, a0));
    return { col, rc, jour: +col.dataset.j, a0 };
  };
  const nettoie = () => document.querySelectorAll('.ex-col.drop').forEach(n => n.classList.remove('drop'));
  const mv = m => {
    if (!gh && Math.hypot(m.clientX - x0, m.clientY - y0) < 6) return;
    if (!gh) {
      gh = el.cloneNode(true); gh.classList.add('ex-ghost'); gh.classList.remove('sel', 'err', 'warn');
      gh.removeAttribute('onpointerdown'); gh.removeAttribute('onclick');
      Object.assign(gh.style, { width: r0.width + 'px', height: r0.height + 'px', background: '#fff', right: 'auto' });
      lab = document.createElement('div'); lab.className = 'ex-gt'; gh.prepend(lab);
      document.body.appendChild(gh); el.classList.add('dragging');
    }
    const c = calc(m.clientX, m.clientY); nettoie();
    if (c) {
      c.col.classList.add('drop');
      Object.assign(gh.style, { left: (c.rc.left + 4) + 'px', width: (c.rc.width - 8) + 'px', top: (c.rc.top + (c.a0 - h0()) * PPM) + 'px' });
      lab.textContent = `→ ${libJour(c.jour)} · ${m2h(c.a0)}–${m2h(c.a0 + duree)}`;
    } else { Object.assign(gh.style, { left: (m.clientX - 20) + 'px', top: (m.clientY - dy) + 'px' }); lab.textContent = 'Déposez sur une journée'; }
    cible = c;
  };
  const up = () => {
    document.removeEventListener('pointermove', mv); document.removeEventListener('pointerup', up); nettoie();
    if (!gh) { editer(id); return; }
    gh.remove(); el.classList.remove('dragging'); X.ignoreClic = Date.now();
    if (!cible || (cible.jour === e.jour && m2h(cible.a0) === e.debut)) return;
    const D = JSON.parse(JSON.stringify(e)); D.jour = cible.jour; D.debut = m2h(cible.a0); D.fin = m2h(cible.a0 + duree);
    if (!erreursEpreuve(D).length && !manquants(D).length) {
      S().epreuves[S().epreuves.findIndex(x => x.id === id)] = D;
      S().epreuves.sort((a, b) => a.jour - b.jour || h2m(a.debut) - h2m(b.debut));
      X.draft = null; modifie(); rendreTabs(); rendreCorps(); msg(`Épreuve déplacée : ${libJour(D.jour)} · ${D.debut}–${D.fin}`);
    } else {
      X.draft = D; rendreCorps();
      msg('Déplacement à confirmer : la salle ou un surveillant n\'est pas disponible sur ce créneau. Ajustez à droite puis validez.', false);
    }
  };
  document.addEventListener('pointermove', mv); document.addEventListener('pointerup', up);
}
const majForm = () => { const f = $('exForm'); if (f) f.innerHTML = vForm(); };
function dSet(k, v, rer) {
  X.draft[k] = v;
  if (k === 'classe') X.classeSel = v;
  if (k === 'debut' && h2m(X.draft.fin) <= h2m(v)) X.draft.fin = m2h(Math.min(h1(), h2m(v) + X.derniereDuree));
  if (rer) rendreCorps();
}
function dSalle(id, on) { const D = X.draft; if (on) D.salles.push({ id, surv: [] }); else D.salles = D.salles.filter(sa => sa.id !== id); majForm(); }
function dSurv(sid, v, on) { const sa = X.draft.salles.find(x => x.id === sid); if (!sa) return; if (on) { if (!sa.surv.includes(v)) sa.surv.push(v); } else sa.surv = sa.surv.filter(x => x !== v); majForm(); }
function manquants(D) {
  const m = [];
  if (!D.salles.length) m.push('• Affectez au moins une salle.');
  D.salles.forEach(sa => { if (!sa.surv.length) m.push(`• Salle ${salle(sa.id).num} : affectez au moins un surveillant.`); });
  return m;
}
function dValider() {
  const D = X.draft, err = erreursEpreuve(D).concat(manquants(D)); if (err.length) { majForm(); return; }
  delete D._nouveau; D.matiere = String(D.matiere || '').trim();
  const i = S().epreuves.findIndex(e => e.id === D.id);
  if (i >= 0) S().epreuves[i] = D; else S().epreuves.push(D);
  S().epreuves.sort((a, b) => a.jour - b.jour || h2m(a.debut) - h2m(b.debut));
  X.derniereDuree = h2m(D.fin) - h2m(D.debut); X.draft = null; modifie(); rendreTabs(); rendreCorps();
}
function dSupprimer() { if (!confirm('Supprimer cette épreuve ?')) return; S().epreuves = S().epreuves.filter(e => e.id !== X.draft.id); X.draft = null; modifie(); rendreTabs(); rendreCorps(); }
function dAnnuler() { X.draft = null; rendreCorps(); }
// classe proposée pour la copie : la première classe qui n'a pas encore d'épreuve sur ce créneau
function dupClasseDefaut(D) {
  const a0 = h2m(D.debut), a1 = h2m(D.fin);
  const libre = S().classes.find(c => c.classe !== D.classe && !S().epreuves.some(e => e.classe === c.classe && e.jour === D.jour && ov(a0, a1, h2m(e.debut), h2m(e.fin))));
  return (libre || S().classes.find(c => c.classe !== D.classe) || { classe: D.classe }).classe;
}
function dDupliquer() {
  const src = S().epreuves.find(e => e.id === X.draft.id) || X.draft;
  const D = JSON.parse(JSON.stringify(src)); D.id = uid('E'); D._nouveau = true;
  D.classe = ($('exDupC') || {}).value || D.classe; D.jour = +(($('exDupJ') || {}).value ?? D.jour);
  const hd = ($('exDupH') || {}).value; if (hd) { const du = h2m(D.fin) - h2m(D.debut); D.debut = hd; D.fin = m2h(h2m(hd) + du); }
  ouvrirCopie(D);
}
// Copie rapide depuis le calendrier : même classe, même jour, juste après l'original (ou même horaire si la journée est finie)
function dupRapide(id) {
  if (lecture()) return;
  const src = S().epreuves.find(e => e.id === id); if (!src) return;
  const D = JSON.parse(JSON.stringify(src)); D.id = uid('E'); D._nouveau = true;
  const du = h2m(src.fin) - h2m(src.debut);
  if (h2m(src.fin) + du <= h1()) { D.debut = src.fin; D.fin = m2h(h2m(src.fin) + du); }
  ouvrirCopie(D);
}
function ouvrirCopie(D) {
  // on garde la salle et les surveillants seulement s'ils sont libres sur le nouveau créneau
  D.salles = D.salles.filter(sa => etatSalle(sa.id, D).ok).map(sa => ({ id: sa.id, surv: sa.surv.filter(v => etatSurv(v, D, sa.id).ok) }));
  X.draft = D; X.classeSel = D.classe; X.onglet = 'edt'; rendreTabs(); rendreCorps();
  msg(`Copie pour ${D.classe} (${libJour(D.jour)} ${D.debut}–${D.fin}) : choisissez la salle et le surveillant, puis validez.`);
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
/* ---------- grilles demi-journées (classes × jours / jours × classes) ---------- */
const MIDI = 12 * 60; // une épreuve qui commence avant 12 h est « matin »
const DEMI = ['Matin', 'Après-midi'];
function celluleEp(classe, jour, demi, R) {
  return S().epreuves.filter(e => e.classe === classe && e.jour === jour && (h2m(e.debut) < MIDI ? 0 : 1) === demi)
    .sort((a, b) => h2m(a.debut) - h2m(b.debut))
    .map(e => { const L = R.filter(r => r.parts.some(p => p.epId === e.id));
      const lieux = L.map(r => 'Salle ' + salle(r.salleId).num + (r.surv.length ? ' (' + nomsSurv(r.surv) + ')' : '')).join(', ');
      return { e, txt: `${e.matiere || 'Épreuve'}\n${e.debut} – ${e.fin}${lieux ? '\n' + lieux : ''}` }; });
}
// mode 'cj' : lignes = classes (bloc matin puis bloc après-midi), colonnes = jours
// mode 'jc' : lignes = jours (matin / après-midi), colonnes = classes
function grille(mode) {
  const s = S(), R = affectations(s.optim ? X.etatVersion : 'init'), J = Array.from({ length: s.nbJours }, (_, j) => j);
  if (mode === 'cj') return { cols: J.map(libJour), groupes: DEMI.map((d, k) => ({ titre: d,
    lignes: s.classes.map(c => ({ lab: c.classe, coul: couleur(c.classe), cells: J.map(j => celluleEp(c.classe, j, k, R)) }))
      .filter(l => l.cells.some(L => L.length)) })).filter(g => g.lignes.length) }; // classes sans épreuve sur la demi-journée masquées
  return { cols: s.classes.map(c => c.classe), colCoul: s.classes.map(c => couleur(c.classe)), groupes: J.map(j => ({ titre: libJour(j),
    lignes: DEMI.map((d, k) => ({ lab: d, cells: s.classes.map(c => celluleEp(c.classe, j, k, R)) })) })) };
}
function vGrille(mode) {
  const G = grille(mode), n = G.cols.length;
  if (!n || !G.groupes.length) return '<div class="ex-card">Rien à afficher : ajoutez des classes et des épreuves.</div>';
  const th = G.cols.map((c, i) => `<th style="min-width:130px">${G.colCoul ? `<span class="ex-dot" style="display:inline-block;background:${G.colCoul[i]}"></span> ` : ''}${esc(c)}</th>`).join('');
  const corps = G.groupes.map(g => `<tr><td colspan="${n + 1}" style="background:#1A4E8A;color:#fff;font-weight:800;font-size:12px;letter-spacing:.05em;text-transform:uppercase;padding:7px 12px">${esc(g.titre)}</td></tr>` +
    g.lignes.map(l => `<tr><td style="font-weight:800;white-space:nowrap;border-left:4px solid ${l.coul || '#F26522'}">${esc(l.lab)}</td>${l.cells.map(L => `<td style="vertical-align:top">${L.map(x => {
      const c = couleur(x.e.classe), t = x.txt.split('\n');
      return `<div onclick="EXSIM.editer('${x.e.id}')" style="cursor:pointer;margin-bottom:4px;padding:5px 7px;border-radius:8px;background:${c}1a;border-left:3px solid ${c};font-size:12px;line-height:1.35"><b>${esc(t[0])}</b><br>${t.slice(1).map(esc).join('<br>')}</div>`; }).join('')}</td>`).join('')}</tr>`).join('')).join('');
  return `<div class="ex-card" style="overflow-x:auto"><table class="ex-tbl" style="min-width:${180 + n * 140}px"><thead><tr><th>${mode === 'cj' ? 'Classe' : 'Jour'}</th>${th}</tr></thead><tbody>${corps}</tbody></table>
    <p class="text-xs text-gray-400 mt-2">Matin = épreuve qui commence avant 12 h. Cliquez sur une épreuve pour l'ouvrir.</p></div>`;
}
const LIBG = { cj: 'Classes par jour (matin - après-midi)', jc: 'Jours par classe (matin - après-midi)' };
async function exportGrille(mode, format) {
  const G = grille(mode), n = G.cols.length, titre = LIBG[mode];
  if (!n) { msg('Rien à exporter.', false); return; }
  const tete = [mode === 'cj' ? 'Demi-journée' : 'Jour', mode === 'cj' ? 'Classe' : 'Demi-journée'].concat(G.cols);
  const lignes = []; G.groupes.forEach(g => g.lignes.forEach((l, i) => lignes.push({ g: g.titre, prem: i === 0, nb: g.lignes.length, l })));
  const val = L => L.map(x => x.txt).join('\n\n');
  try {
    if (format === 'xlsx') {
      msg('Préparation du fichier Excel…'); await libExcel();
      const aoa = [[`${S().titre || 'Simulation'} — ${titre}`], [`HECG · ${semaineTxt()} · exporté le ${new Date().toLocaleString('fr-FR')}`], [], tete];
      lignes.forEach(x => aoa.push([x.prem ? x.g : '', x.l.lab].concat(x.l.cells.map(val))));
      const ws = XS.utils.aoa_to_sheet(aoa), nc = n + 2, rows = [{ hpt: 30 }, { hpt: 18 }, { hpt: 6 }, { hpt: 24 }];
      const merges = [{ s: { r: 0, c: 0 }, e: { r: 0, c: nc - 1 } }, { s: { r: 1, c: 0 }, e: { r: 1, c: nc - 1 } }];
      let r0 = 4; G.groupes.forEach(g => { if (g.lignes.length > 1) merges.push({ s: { r: r0, c: 0 }, e: { r: r0 + g.lignes.length - 1, c: 0 } }); r0 += g.lignes.length; });
      for (let R = 0; R < aoa.length; R++) for (let c = 0; c < nc; c++) {
        const a = XS.utils.encode_cell({ r: R, c }); if (!ws[a]) ws[a] = { t: 's', v: '' };
        let st;
        if (R === 0) st = XST.T; else if (R === 1) st = XST.ST; else if (R === 2) st = {}; else if (R === 3) st = XST.H;
        else {
          const i = R - 4, x = lignes[i];
          st = JSON.parse(JSON.stringify(c === 0 ? XST.H : (i % 2 ? XST.D2 : XST.D1)));
          if (c === 0) st.alignment = { vertical: 'center', horizontal: 'center', wrapText: true };
          if (c === 1) { st.font = Object.assign({}, st.font, { bold: true }); if (x.l.coul) st.border = Object.assign({}, st.border, { left: { style: 'thick', color: { rgb: x.l.coul.slice(1).toUpperCase() } } }); }
        }
        ws[a].s = JSON.parse(JSON.stringify(st));
      }
      lignes.forEach(x => { const nl = Math.max(1, ...x.l.cells.map(L => val(L).split('\n').length)); rows.push({ hpt: 15 * nl + 6 }); });
      ws['!merges'] = merges; ws['!rows'] = rows; ws['!cols'] = [{ wch: 14 }, { wch: 18 }].concat(G.cols.map(() => ({ wch: 26 })));
      ws['!ref'] = XS.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: aoa.length - 1, c: nc - 1 } });
      const wb = XS.utils.book_new(); XS.utils.book_append_sheet(wb, ws, mode === 'cj' ? 'Classes x jours' : 'Jours x classes');
      XS.writeFile(wb, nomFichier(titre, 'xlsx')); msg('Fichier Excel téléchargé ✓');
    } else {
      const { doc, lg } = await pdfNouveau(), fs = n > 8 ? 6.6 : n > 5 ? 7.6 : 8.6;
      const body = lignes.map(x => (x.prem ? [{ content: pdfT(x.g), rowSpan: x.nb, styles: { fillColor: [26, 78, 138], textColor: 255, fontStyle: 'bold', valign: 'middle', halign: 'center' } }] : [])
        .concat([pdfT(x.l.lab)], x.l.cells.map(L => pdfT(val(L)))));
      doc.autoTable({ startY: TOP + 2, head: [tete.map(pdfT)], body, theme: 'grid', margin: { left: MX, right: MX, top: TOP, bottom: 17 },
        styles: { font: 'helvetica', fontSize: fs, cellPadding: 1.8, lineColor: [221, 228, 238], lineWidth: 0.2, textColor: [14, 26, 43], valign: 'top', overflow: 'linebreak' },
        headStyles: { fillColor: [26, 78, 138], textColor: 255, fontStyle: 'bold' },
        columnStyles: { 0: { cellWidth: 20 }, 1: { cellWidth: mode === 'cj' ? 26 : 20, fontStyle: 'bold' } },
        didDrawCell: d => { if (d.section === 'body' && d.column.index === 1) { const x = lignes[d.row.index]; if (x && x.l.coul) { doc.setFillColor(...RGB(x.l.coul)); doc.rect(d.cell.x, d.cell.y, 1.4, d.cell.height, 'F'); } } } });
      pdfHabiller(doc, lg, titre); doc.save(nomFichier(titre, 'pdf')); msg('PDF téléchargé ✓');
    }
  } catch (e) { msg('Export impossible : ' + e.message, false); }
}
function vGlobale() {
  const s = S();
  const choix = `<div class="flex gap-2 flex-wrap items-center mb-3"><span class="text-xs font-black uppercase text-gray-400">Vue</span>${[['chrono', 'Chronologie salles / surveillants'], ['cj', 'Classes × jours'], ['jc', 'Jours × classes']].map(([k, l]) => `<button class="ex-b sm ${X.gVue === k ? 'pri' : ''}" onclick="EXSIM.gVue('${k}')">${l}</button>`).join('')}
    ${X.gVue !== 'chrono' ? `<span style="flex:1"></span>${s.optim ? `<button class="ex-b sm ${X.etatVersion === 'init' ? 'pri' : ''}" onclick="EXSIM.version('init')">Avant optimisation</button><button class="ex-b sm ${X.etatVersion === 'optim' ? 'vio' : ''}" onclick="EXSIM.version('optim')">Après optimisation</button>` : ''}
      <button class="ex-b sm" onclick="EXSIM.exportGrille('${X.gVue}','xlsx')">⬇ Excel</button><button class="ex-b sm ora" onclick="EXSIM.exportGrille('${X.gVue}','pdf')">⬇ PDF</button>` : ''}</div>`;
  return choix + (X.gVue === 'chrono' ? vChrono() : vGrille(X.gVue));
}
function vChrono() {
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
  return `<div class="ex-sect">
    ${[['salle', '🏫 Par salle'], ['classe', '🎓 Par classe'], ['surv', '👁 Par surveillant']].map(([k, l]) => `<button class="ex-b ${X.etatVue === k ? 'pri' : ''}" onclick="EXSIM.etatVue('${k}')">${l}</button>`).join('')}
    <select style="max-width:260px" onchange="EXSIM.etatFiltre(this.value)"><option value="">Tous</option>${opts.map(o => `<option value="${esc(o[0])}" ${o[0] === X.etatFiltre ? 'selected' : ''}>${esc(o[1])}</option>`).join('')}</select>
    ${s.optim ? `<button class="ex-b sm ${version === 'init' ? 'pri' : ''}" onclick="EXSIM.version('init')">Avant optimisation</button><button class="ex-b sm ${version === 'optim' ? 'vio' : ''}" onclick="EXSIM.version('optim')">Après optimisation</button>` : ''}
    <span style="flex:1"></span><button class="ex-b" onclick="EXSIM.exportExcel()">⬇ Excel (simulation complète)</button><button class="ex-b ora" onclick="EXSIM.pdfEtat()">⬇ PDF de cette vue</button></div>
    ${etatsHtml(X.etatVue, X.etatFiltre, version) || '<div class="ex-card">Rien à afficher.</div>'}`;
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
  const s = S(), selS = new Set(prm.salles), selC = new Set(prm.classes);
  const items = [], fixes = [];
  s.epreuves.forEach(e => {
    const a0 = h2m(e.debut), a1 = h2m(e.fin);
    if (selC.has(e.classe) && e.salles.length && e.salles.every(sa => selS.has(sa.id))) {
      const etus = etudiants(e.classe).map(x => ({ id: x.id, nom: x.nom })).sort((a, b) => nomTri(a.nom).localeCompare(nomTri(b.nom), 'fr'));
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
        // jamais deux épreuves de la même classe dans la même salle en même temps (oraux simultanés)
        if (r.parts.some(p => p.epId !== it.e.id && p.classe === it.e.classe && ov(p.a0, p.a1, it.a0, it.a1))) return 0;
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
    const besoin = 1;   // règle HECG : un surveillant par salle
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
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:28px">
      <div><label class="l">Salles à optimiser</label>${s.salles.map(x => `<label style="display:flex;gap:6px;align-items:center;font-size:12.5px;margin:3px 0"><input type="checkbox" style="width:auto" ${P.salles.includes(x.id) ? 'checked' : ''} onchange="EXSIM.prm('salles','${esc(x.id)}',this.checked)"> Salle ${esc(x.num)} <span class="text-gray-400">· ${capSalle(x.id)} pl.</span></label>`).join('')}</div>
      <div><label class="l">Classes à optimiser</label>${s.classes.map(c => `<label style="display:flex;gap:6px;align-items:center;font-size:12.5px;margin:3px 0"><input type="checkbox" style="width:auto" ${P.classes.includes(c.classe) ? 'checked' : ''} onchange="EXSIM.prm('classes','${esc(c.classe)}',this.checked)"><span class="ex-dot" style="background:${couleur(c.classe)}"></span> ${esc(c.classe)} <span class="text-gray-400">· ${effectif(c.classe)}</span></label>`).join('')}</div></div>
    <div class="flex flex-wrap gap-4 items-center mt-3" style="font-size:12.5px;font-weight:600">
      <label style="display:flex;gap:6px;align-items:center"><input type="checkbox" style="width:auto" ${P.melanger ? 'checked' : ''} onchange="EXSIM.prm('melanger',null,this.checked)"> Regrouper plusieurs classes dans une même salle (mélange nominatif)</label>
      <label style="display:flex;gap:6px;align-items:center"><input type="checkbox" style="width:auto" ${P.unSurDeux ? 'checked' : ''} onchange="EXSIM.prm('unSurDeux',null,this.checked)"> Une place sur deux quand la salle le permet</label>
      <span class="text-xs text-gray-500">Encadrement : 1 surveillant par salle</span></div>
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
    <div class="ex-sect"><button class="ex-b ora" onclick="EXSIM.pdfPlans()">⬇ Plans de classe (PDF)</button><button class="ex-b" onclick="EXSIM.pdfRepartition()">⬇ Répartition nominative (PDF)</button><button class="ex-b" onclick="EXSIM.pdfEmargement()">⬇ Émargement par salle (PDF)</button><button class="ex-b" onclick="EXSIM.exportExcel()">⬇ Tout en Excel</button>
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
  const court = n => esc(nomCourt(n));
  return `<div class="ex-plan" style="width:${G.W * k}px;height:${G.H * k}px"><div style="position:absolute;left:50%;top:4px;transform:translateX(-50%);font-size:10px;font-weight:800;letter-spacing:.12em;color:#94a3b8">TABLEAU</div>
    ${G.L.map((st, i) => { const o = occ[i] || [], c = o.length ? couleur(o[0].classe) : '#cbd5e1';
      return `<div class="ex-seat ${o.length ? '' : 'vide'}" style="left:${st.x * k}px;top:${st.y * k}px;width:${st.w * k}px;height:${st.h * k}px;border-left-color:${c}"><i>${i + 1}</i> ${o.map(p => `<b>${court(p.nom)}</b><br><span style="color:#64748b">${esc(p.classe)}</span>`).join('<br>')}</div>`; }).join('')}</div>`;
}
function nomCourt(n) { const t = nomTri(n).split(' '), w = String(n || '').trim().split(/\s+/), up = w.filter(x => /[A-ZÀ-Þ]/.test(x) && x === x.toUpperCase()); const pre = w.filter(x => !up.includes(x)); return (up.join(' ') || t[0]) + (pre.length ? ' ' + pre[0].charAt(0) + '.' : ''); }
function legendeGroupe(g) { return g.parts.map(p => `<span class="ex-chip"><span class="ex-dot" style="background:${couleur(p.classe)}"></span>${esc(p.classe)} · ${esc(p.matiere || '')} · ${m2h(p.a0)}–${m2h(p.a1)} · ${p.n} étu.</span>`).join(' '); }
function vPlan(g) {
  const L = (g.places || []).slice().sort((a, b) => nomTri(a.nom).localeCompare(nomTri(b.nom), 'fr'));
  return `<div class="ex-card"><div class="flex items-center gap-2 flex-wrap"><h3 style="margin:0;color:#1A4E8A;font-size:14px;text-transform:none;letter-spacing:0">🪑 Plan de classe — Salle ${esc(salle(g.salleId).num)} · ${esc(libJour(g.jour))} ${m2h(g.a0)}–${m2h(g.a1)}</h3><span style="flex:1"></span>
    <button class="ex-b sm ora" onclick="EXSIM.pdfPlans(${S().optim.groupes.indexOf(g)})">⬇ PDF de ce plan</button><button class="ex-b sm" onclick="EXSIM.voirPlan(null)">Fermer</button></div>
    <div class="mt-2">${legendeGroupe(g)}</div>${sieges(g.salleId).plan ? '' : '<p class="text-xs text-gray-400 mt-1">Grille automatique (aucun plan enregistré pour cette salle).</p>'}
    <div style="overflow:auto">${planHtml(g)}</div>
    <details><summary class="text-sm font-bold cursor-pointer">Liste alphabétique (${L.length})</summary><table class="ex-tbl mt-2"><thead><tr><th>Étudiant</th><th>Classe</th><th>Place</th></tr></thead><tbody>${L.map(p => `<tr><td>${esc(nomTri(p.nom))}</td><td>${esc(p.classe)}</td><td><b>${p.s + 1}</b></td></tr>`).join('')}</tbody></table></details></div>`;
}
function voirPlan(k) { X.planSel = k; rendreCorps(); if (k != null) setTimeout(() => { const b = $('exBody'); b.scrollTop = b.scrollHeight; }, 30); }
/* ---------- exports Excel / PDF aux couleurs HECG (bleu #1A4E8A, orange #F26522) ---------- */
const RGB = h => [parseInt(h.slice(1, 3), 16), parseInt(h.slice(3, 5), 16), parseInt(h.slice(5, 7), 16)];
const PALE = (h, k) => RGB(h).map(v => Math.round(v + (255 - v) * (k == null ? 0.84 : k)));
const charger = src => new Promise((ok, ko) => { const t = document.createElement('script'); t.src = src; t.onload = ok; t.onerror = () => ko(new Error('bibliothèque indisponible')); document.head.appendChild(t); });
let XS = null, LOGO;
async function libExcel() {
  if (XS) return XS;
  const ancien = window.XLSX;                 // la page utilise déjà SheetJS : on le restaure après chargement
  await charger('https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js');
  XS = window.XLSX; if (ancien) window.XLSX = ancien;
  return XS;
}
async function libPdf() {
  if (!(window.jspdf && window.jspdf.jsPDF)) await charger('https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js');
  if (!window.jspdf.jsPDF.API.autoTable) await charger('https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.2/jspdf.plugin.autotable.min.js');
  return window.jspdf.jsPDF;
}
async function logo() {
  if (LOGO !== undefined) return LOGO;
  try {
    const b = await (await fetch('logo-hecg.png')).blob();
    const url = await new Promise(ok => { const f = new FileReader(); f.onload = () => ok(f.result); f.readAsDataURL(b); });
    const img = await new Promise((ok, ko) => { const i = new Image(); i.onload = () => ok(i); i.onerror = ko; i.src = url; });
    const c = document.createElement('canvas'), k = 260 / img.width; c.width = 260; c.height = Math.round(img.height * k);
    c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
    LOGO = { url: c.toDataURL('image/png'), r: c.height / c.width };
  } catch (e) { LOGO = null; }
  return LOGO;
}
const nomFichier = (t, ext) => ((S().titre || 'Simulation') + ' - ' + t).replace(/[\\/:*?"<>|]+/g, ' ').replace(/\s+/g, ' ').trim() + '.' + ext;
const semaineTxt = () => S().lundi ? `Semaine du ${dateJour(0)} au ${dateJour(S().nbJours - 1)}` : '';
const H1 = m => (m / 60).toFixed(1).replace('.', ',') + ' h';
const nomsSurv = L => L.map(x => (surv(x) || { nom: '?' }).nom).join(', ') || '—';
const LIBV = { salle: 'par salle', classe: 'par classe', surv: 'par surveillant' };

// Données communes aux exports : une « section » = titre, sous-titre, colonnes, lignes
function sections(vue, filtre, version) {
  const s = S(), R = affectations(version);
  if (vue === 'salle') return s.salles.filter(x => !filtre || x.id === filtre).map(x => {
    const L = R.filter(r => r.salleId === x.id), occ = L.reduce((t, r) => t + r.a1 - r.a0, 0), dispo = (h1() - h0()) * s.nbJours;
    return { titre: `Salle ${x.num} — ${x.campus}`, sous: `${capSalle(x.id)} places · occupée ${H1(occ)} sur ${H1(dispo)} (${dispo ? Math.round(occ / dispo * 100) : 0} %)`,
      cols: ['Jour', 'Horaires', 'Classe(s) et épreuve', 'Effectif', 'Surveillant(s)'], larg: [18, 15, 50, 12, 34],
      couleurs: L.map(r => couleur(r.parts[0].classe)), vide: 'Salle libre toute la semaine.',
      rows: L.map(r => [libJour(r.jour), m2h(r.a0) + ' – ' + m2h(r.a1), r.parts.map(p => p.classe + ' · ' + (p.matiere || 'Épreuve') + (r.parts.length > 1 ? ' (' + m2h(p.a0) + '–' + m2h(p.a1) + ')' : '')).join('\n'),
        r.parts.reduce((t, p) => t + p.n, 0) + ' / ' + capSalle(x.id), nomsSurv(r.surv)]) };
  });
  if (vue === 'classe') return s.classes.filter(c => !filtre || c.classe === filtre).map(c => {
    const E = s.epreuves.filter(e => e.classe === c.classe), tot = E.reduce((t, e) => t + h2m(e.fin) - h2m(e.debut), 0);
    return { titre: c.classe, coul: couleur(c.classe), sous: `${effectif(c.classe)} étudiant(s) · ${E.length} épreuve(s) · ${H1(tot)} d'épreuves`,
      cols: ['Jour', 'Horaires', 'Épreuve', 'Salle(s)', 'Surveillant(s)'], larg: [18, 15, 38, 34, 34], couleurs: E.map(() => couleur(c.classe)), vide: 'Aucune épreuve.',
      rows: E.map(e => { const L = R.filter(r => r.parts.some(p => p.epId === e.id));
        return [libJour(e.jour), e.debut + ' – ' + e.fin, (e.matiere || 'Épreuve') + (e.note ? '\n' + e.note : ''),
          L.map(r => 'Salle ' + salle(r.salleId).num + ' · ' + r.parts.filter(p => p.epId === e.id).reduce((t, p) => t + p.n, 0) + ' étu.' + (r.parts.length > 1 ? ' (avec ' + r.parts.filter(p => p.epId !== e.id).map(p => p.classe).join(', ') + ')' : '')).join('\n') || '—',
          L.map(r => nomsSurv(r.surv)).join('\n')]; }) };
  });
  return s.surveillants.filter(v => !filtre || v.id === filtre).map(v => {
    const L = R.filter(r => r.surv.includes(v.id)), tot = L.reduce((t, r) => t + r.a1 - r.a0, 0);
    return { titre: v.nom, sous: `${H1(tot)} de surveillance · disponibilités : ${Array.from({ length: s.nbJours }, (_, j) => JOURS[j].slice(0, 3) + '. ' + dispoTexte(v, j)).join(' · ')}`,
      cols: ['Jour', 'Horaires', 'Salle', 'Classe(s) et épreuve', 'Avec'], larg: [18, 15, 14, 50, 34], couleurs: L.map(r => couleur(r.parts[0].classe)), vide: 'Aucune surveillance.',
      rows: L.map(r => [libJour(r.jour), m2h(r.a0) + ' – ' + m2h(r.a1), 'Salle ' + salle(r.salleId).num, r.parts.map(p => p.classe + ' · ' + (p.matiere || 'Épreuve')).join('\n'), nomsSurv(r.surv.filter(x => x !== v.id))]) };
  });
}
const placesTriees = g => (g.places || []).slice().sort((a, b) => nomTri(a.nom).localeCompare(nomTri(b.nom), 'fr'));
const titreGroupe = g => `Salle ${salle(g.salleId).num} — ${libJour(g.jour)} · ${m2h(g.a0)}–${m2h(g.a1)}`;
const sousGroupe = g => g.parts.map(p => `${p.classe} · ${p.matiere || 'Épreuve'} (${m2h(p.a0)}–${m2h(p.a1)}, ${p.n} étu.)`).join(' + ') + ' · Surveillance : ' + nomsSurv(g.surv);
function blocsPlans() {
  return S().optim.groupes.map(g => { const L = (g.places || []).slice().sort((a, b) => a.s - b.s);
    return { titre: titreGroupe(g), sous: sousGroupe(g), cols: ['Place', 'Étudiant', 'Classe'], rows: L.map(p => [String(p.s + 1), nomTri(p.nom), p.classe]), couleurs: L.map(p => couleur(p.classe)), vide: '—' }; });
}
function blocsRepartition() {
  const s = S(), O = s.optim, out = [];
  s.classes.forEach(c => s.epreuves.filter(e => e.classe === c.classe).forEach(e => {
    const G = O.groupes.filter(g => g.parts.some(p => p.epId === e.id));
    const L = [].concat(...G.map(g => (g.places || []).filter(p => p.epId === e.id).map(p => ({ nom: p.nom, salle: salle(g.salleId).num, place: p.s + 1 })))).sort((a, b) => nomTri(a.nom).localeCompare(nomTri(b.nom), 'fr'));
    out.push({ titre: `${c.classe} — ${e.matiere || 'Épreuve'}`, coul: couleur(c.classe), sous: `${libJour(e.jour)} · ${e.debut}–${e.fin}` + (G.length ? '' : ' · épreuve non optimisée : ' + e.salles.map(sa => 'salle ' + salle(sa.id).num).join(', ')),
      cols: ['Étudiant', 'Salle', 'Place'], rows: L.map(p => [nomTri(p.nom), 'Salle ' + p.salle, String(p.place)]), couleurs: L.map(() => couleur(c.classe)), vide: 'Pas de placement nominatif pour cette épreuve.' });
  }));
  return out;
}

/* ----- Excel ----- */
const bord = c => { const l = { style: 'thin', color: { rgb: c } }; return { top: l, bottom: l, left: l, right: l }; };
const XST = {
  T: { font: { name: 'Calibri', bold: true, sz: 16, color: { rgb: 'FFFFFF' } }, fill: { patternType: 'solid', fgColor: { rgb: '1A4E8A' } }, alignment: { vertical: 'center', indent: 1 } },
  ST: { font: { name: 'Calibri', italic: true, sz: 10, color: { rgb: 'DCE6F4' } }, fill: { patternType: 'solid', fgColor: { rgb: '1A4E8A' } }, alignment: { vertical: 'center', indent: 1 } },
  O: { fill: { patternType: 'solid', fgColor: { rgb: 'F26522' } } },
  B: { font: { name: 'Calibri', bold: true, sz: 13, color: { rgb: '1A4E8A' } }, alignment: { vertical: 'bottom' }, border: { bottom: { style: 'medium', color: { rgb: 'F26522' } } } },
  BS: { font: { name: 'Calibri', italic: true, sz: 10, color: { rgb: '64748B' } }, alignment: { vertical: 'top', wrapText: true } },
  H: { font: { name: 'Calibri', bold: true, sz: 11, color: { rgb: 'FFFFFF' } }, fill: { patternType: 'solid', fgColor: { rgb: '1A4E8A' } }, alignment: { vertical: 'center', wrapText: true, indent: 1 }, border: bord('1A4E8A') },
  D1: { font: { name: 'Calibri', sz: 10.5, color: { rgb: '0E1A2B' } }, alignment: { vertical: 'top', wrapText: true, indent: 1 }, border: bord('DCE3ED') },
  D2: { font: { name: 'Calibri', sz: 10.5, color: { rgb: '0E1A2B' } }, fill: { patternType: 'solid', fgColor: { rgb: 'F4F7FB' } }, alignment: { vertical: 'top', wrapText: true, indent: 1 }, border: bord('DCE3ED') },
  V: { font: { name: 'Calibri', italic: true, sz: 10, color: { rgb: '94A3B8' } } }
};
function feuille(titre, blocs, largeurs) {
  const n = largeurs.length, aoa = [], st = [], hts = [], merges = [];
  const ligne = (vals, style, h, fus, coul) => { const r = aoa.length; aoa.push(vals); st.push({ style, coul }); hts.push(h); if (fus) merges.push({ s: { r, c: 0 }, e: { r, c: n - 1 } }); };
  ligne([`${S().titre || 'Simulation'} — ${titre}`], 'T', 32, true);
  ligne([`HECG · ${semaineTxt()} · exporté le ${new Date().toLocaleString('fr-FR')}`], 'ST', 18, true);
  ligne([''], 'O', 5, true);
  ligne([], null, 12);
  blocs.forEach(b => {
    ligne([b.titre], 'B', 24, true);
    if (b.sous) ligne([b.sous], 'BS', Math.max(16, 14 * Math.ceil(b.sous.length / (largeurs.reduce((t, w) => t + w, 0) * 1.1))), true);
    ligne(b.cols, 'H', 22);
    if (!b.rows.length) ligne([b.vide || '—'], 'V', 18, true);
    b.rows.forEach((row, i) => {
      const nl = Math.max(1, ...row.map((v, c) => String(v == null ? '' : v).split('\n').reduce((t, l) => t + Math.max(1, Math.ceil(l.length / Math.max(6, (largeurs[c] || 12) - 3))), 0)));
      ligne(row, i % 2 ? 'D2' : 'D1', 15 * nl + 5, false, b.couleurs && b.couleurs[i]);
    });
    ligne([], null, 14);
  });
  const ws = XS.utils.aoa_to_sheet(aoa);
  st.forEach((x, r) => {
    if (!x.style) return;
    for (let c = 0; c < n; c++) {
      const a = XS.utils.encode_cell({ r, c }); if (!ws[a]) ws[a] = { t: 's', v: '' };
      const sty = JSON.parse(JSON.stringify(XST[x.style]));
      if (c === 0 && x.coul) sty.border = Object.assign({}, sty.border, { left: { style: 'thick', color: { rgb: x.coul.slice(1).toUpperCase() } } });
      ws[a].s = sty;
    }
  });
  ws['!merges'] = merges; ws['!cols'] = largeurs.map(w => ({ wch: w })); ws['!rows'] = hts.map(h => (h ? { hpt: h } : {}));
  ws['!ref'] = XS.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: aoa.length - 1, c: n - 1 } });
  return ws;
}
async function exportExcel() {
  try {
    msg('Préparation du fichier Excel…');
    await libExcel();
    const s = S(), wb = XS.utils.book_new(), version = s.optim ? 'optim' : 'init';
    const stat = s.optim ? 'Optimisée' : s.statut === 'terminee' ? 'Semaine terminée' : 'En construction';
    XS.utils.book_append_sheet(wb, feuille('Synthèse', [
      { titre: 'Simulation', cols: ['Élément', 'Valeur'], rows: [['Titre', s.titre], ['Semaine', semaineTxt()], ['Plage horaire', s.hDebut + ' – ' + s.hFin], ['Statut', stat], ['Épreuves', String(s.epreuves.length)], ['Encadrement', '1 surveillant par salle']] },
      { titre: 'Classes', cols: ['Classe', 'Étudiants'], rows: s.classes.map(c => [c.classe, String(effectif(c.classe))]), couleurs: s.classes.map(c => couleur(c.classe)), vide: 'Aucune classe.' },
      { titre: 'Salles', cols: ['Salle', 'Places'], rows: s.salles.map(x => ['Salle ' + x.num + ' — ' + x.campus, String(capSalle(x.id))]), vide: 'Aucune salle.' }
    ], [36, 60]), 'Synthèse');
    const lj = Array.from({ length: s.nbJours }, (_, j) => libJour(j));
    XS.utils.book_append_sheet(wb, feuille('Disponibilités des surveillants', [{ titre: 'Disponibilités de la semaine', cols: ['Surveillant'].concat(lj),
      rows: s.surveillants.map(v => [v.nom].concat(lj.map((_, j) => dispoTexte(v, j)))), vide: 'Aucun surveillant.' }], [30].concat(lj.map(() => 20))), 'Disponibilités');
    [['salle', 'Par salle'], ['classe', 'Par classe'], ['surv', 'Par surveillant']].forEach(([v, l]) => {
      const B = sections(v, '', version);
      XS.utils.book_append_sheet(wb, feuille('État ' + LIBV[v] + (version === 'optim' ? ' (après optimisation)' : ''), B, B.length ? B[0].larg : [18, 15, 40, 12, 30]), l);
    });
    if (s.optim) {
      XS.utils.book_append_sheet(wb, feuille('Plans de classe', blocsPlans(), [10, 38, 24]), 'Plans de classe');
      XS.utils.book_append_sheet(wb, feuille('Répartition nominative', blocsRepartition(), [38, 16, 10]), 'Répartition');
    }
    XS.writeFile(wb, nomFichier('export', 'xlsx'));
    msg('Fichier Excel téléchargé ✓');
  } catch (e) { msg('Export Excel impossible : ' + e.message, false); }
}

/* ----- PDF ----- */
const PW = 297, PH = 210, MX = 14, TOP = 38;
const pdfT = t => String(t == null ? '' : t).replace(/→/g, '->').replace(/[–—]/g, '-').replace(/[‘’]/g, "'").replace(/…/g, '...').replace(/œ/g, 'oe').replace(/Œ/g, 'OE').replace(/[^\n\x20-\x7E\xA0-\xFF]/g, '');
async function pdfNouveau() {
  msg('Préparation du PDF…');
  const J = await libPdf(), lg = await logo();
  return { doc: new J({ orientation: 'landscape', unit: 'mm', format: 'a4' }), lg };
}
function pdfHabiller(doc, lg, titre) {
  const n = doc.getNumberOfPages(), edite = new Date().toLocaleDateString('fr-FR');
  for (let i = 1; i <= n; i++) {
    doc.setPage(i);
    let x = MX;
    if (lg) { const h = 17, w = h / lg.r; doc.addImage(lg.url, 'PNG', MX, 6, w, h, 'logoHECG'); x = MX + w + 6; }
    else { doc.setFont('helvetica', 'bold'); doc.setFontSize(22); doc.setTextColor(26, 78, 138); doc.text('HECG', MX, 18); doc.setFillColor(242, 101, 34); doc.circle(MX + 27.5, 16.6, 1.7, 'F'); x = MX + 35; }
    doc.setFont('helvetica', 'bold'); doc.setFontSize(16); doc.setTextColor(26, 78, 138); doc.text(pdfT(titre), x, 14);
    doc.setFont('helvetica', 'normal'); doc.setFontSize(9.5); doc.setTextColor(100, 116, 139); doc.text(pdfT((S().titre || '') + '  ·  ' + semaineTxt()), x, 20.5);
    doc.setFillColor(26, 78, 138); doc.rect(0, 27, PW, 1.3, 'F'); doc.setFillColor(242, 101, 34); doc.rect(0, 28.3, PW * 0.32, 1.3, 'F');
    doc.setDrawColor(227, 234, 243); doc.setLineWidth(0.3); doc.line(MX, PH - 12, PW - MX, PH - 12);
    doc.setFontSize(8); doc.setTextColor(125, 138, 156); doc.text(pdfT('HECG  ·  hecg.fr  ·  édité le ' + edite), MX, PH - 6.5);
    doc.text(`Page ${i} / ${n}`, PW - MX, PH - 6.5, { align: 'right' });
  }
}
function pdfTitreSection(doc, y, b) {
  if (y > PH - 48) { doc.addPage(); y = TOP + 2; }
  doc.setFillColor(...(b.coul ? RGB(b.coul) : [242, 101, 34])); doc.roundedRect(MX, y - 4.6, 1.8, 6.4, 0.8, 0.8, 'F');
  doc.setFont('helvetica', 'bold'); doc.setFontSize(13); doc.setTextColor(26, 78, 138); doc.text(pdfT(b.titre), MX + 5, y);
  y += 5.5;
  if (b.sous) { doc.setFont('helvetica', 'normal'); doc.setFontSize(8.8); doc.setTextColor(100, 116, 139); const L = doc.splitTextToSize(pdfT(b.sous), PW - 2 * MX - 5); doc.text(L, MX + 5, y); y += L.length * 4; }
  return y + 2;
}
function pdfSection(doc, y, b, opts) {
  y = pdfTitreSection(doc, y, b);
  if (!b.rows.length) { doc.setFont('helvetica', 'italic'); doc.setFontSize(9.5); doc.setTextColor(148, 163, 184); doc.text(pdfT(b.vide || '—'), MX + 5, y + 3); return y + 14; }
  doc.autoTable(Object.assign({
    startY: y, head: [b.cols.map(pdfT)], body: b.rows.map(r => r.map(pdfT)), theme: 'grid',
    margin: { left: MX, right: MX, top: TOP, bottom: 17 },
    styles: { font: 'helvetica', fontSize: 9.2, cellPadding: { top: 3, bottom: 3, left: 3.5, right: 3 }, lineColor: [221, 228, 238], lineWidth: 0.2, textColor: [14, 26, 43], valign: 'middle', overflow: 'linebreak' },
    headStyles: { fillColor: [26, 78, 138], textColor: 255, fontStyle: 'bold', fontSize: 9.2 },
    alternateRowStyles: { fillColor: [244, 247, 251] },
    didDrawCell: d => { if (d.section === 'body' && d.column.index === 0 && b.couleurs && b.couleurs[d.row.index]) { doc.setFillColor(...RGB(b.couleurs[d.row.index])); doc.rect(d.cell.x, d.cell.y, 1.5, d.cell.height, 'F'); } }
  }, opts || {}));
  return doc.lastAutoTable.finalY + 12;
}
async function pdfEtat() {
  try {
    const { doc, lg } = await pdfNouveau(), s = S(), version = s.optim ? X.etatVersion : 'init';
    const B = sections(X.etatVue, X.etatFiltre, version); let y = TOP + 4;
    if (!B.length) { msg('Rien à exporter dans cette vue.', false); return; }
    B.forEach(b => { y = pdfSection(doc, y, b); });
    pdfHabiller(doc, lg, 'État ' + LIBV[X.etatVue] + (version === 'optim' ? ' (après optimisation)' : ''));
    doc.save(nomFichier('état ' + LIBV[X.etatVue], 'pdf')); msg('PDF téléchargé ✓');
  } catch (e) { msg('Export PDF impossible : ' + e.message, false); }
}
function pdfPlanSalle(doc, g, y) {
  const P = sieges(g.salleId), zw = PW - 2 * MX, zh = 108, k = Math.min(zw / P.W, zh / P.H), w = P.W * k, h = P.H * k, ox = MX + (zw - w) / 2, oy = y;
  const occ = {}; (g.places || []).forEach(p => (occ[p.s] = occ[p.s] || []).push(p));
  doc.setDrawColor(26, 78, 138); doc.setLineWidth(0.8); doc.roundedRect(ox, oy, w, h, 1.5, 1.5, 'S');
  doc.setFillColor(51, 65, 85); doc.rect(ox + w * 0.3, oy, w * 0.4, 1.4, 'F');
  doc.setFont('helvetica', 'bold'); doc.setFontSize(6.5); doc.setTextColor(148, 163, 184); doc.text('TABLEAU', ox + w / 2, oy + 4.6, { align: 'center' });
  P.L.forEach((st, i) => {
    const x = ox + st.x * k, yy = oy + st.y * k, sw = st.w * k, sh = st.h * k, o = occ[i] || [];
    if (o.length) {
      const c = couleur(o[0].classe);
      doc.setFillColor(...PALE(c, 0.88)); doc.setDrawColor(...PALE(c, 0.45)); doc.setLineWidth(0.25); doc.roundedRect(x, yy, sw, sh, 0.9, 0.9, 'FD');
      doc.setFillColor(...RGB(c)); doc.rect(x, yy, 1.2, sh, 'F');
    } else {
      doc.setFillColor(248, 250, 252); doc.setDrawColor(203, 213, 225); doc.setLineWidth(0.2); doc.setLineDashPattern([0.8, 0.8], 0); doc.roundedRect(x, yy, sw, sh, 0.9, 0.9, 'FD'); doc.setLineDashPattern([], 0);
    }
    const fs = Math.max(4.2, Math.min(7.5, sh * 0.8)), lh = fs * 0.3528 * 1.18, tx = x + 2.2, tw = sw - 3;
    doc.setFont('helvetica', 'bold'); doc.setFontSize(fs * 0.9); doc.setTextColor(148, 163, 184); doc.text(String(i + 1), tx, yy + lh * 0.95);
    if (o.length) {
      const p = o[0];
      doc.setFont('helvetica', 'bold'); doc.setFontSize(fs); doc.setTextColor(14, 26, 43);
      doc.text(doc.splitTextToSize(pdfT(nomCourt(p.nom)), tw)[0] || '', tx, yy + lh * 2);
      doc.setFont('helvetica', 'normal'); doc.setFontSize(fs * 0.85); doc.setTextColor(100, 116, 139);
      if (sh > lh * 2.6) doc.text(doc.splitTextToSize(pdfT(p.classe + (o.length > 1 ? ' (+' + (o.length - 1) + ')' : '')), tw)[0] || '', tx, yy + lh * 2.95);
    }
  });
  return oy + h + 8;
}
async function pdfPlans(k) {
  try {
    const { doc, lg } = await pdfNouveau(), s = S(), G = k != null ? [s.optim.groupes[k]] : s.optim.groupes;
    G.forEach((g, gi) => {
      if (gi) doc.addPage();
      let y = pdfTitreSection(doc, TOP + 4, { titre: titreGroupe(g), sous: sousGroupe(g) });
      // légende des classes
      let lx = MX + 5; doc.setFontSize(8.5);
      g.parts.forEach(p => { const t = pdfT(`${p.classe} (${p.n})`), tw = doc.getTextWidth(t) + 9; doc.setFillColor(...PALE(couleur(p.classe), 0.85)); doc.roundedRect(lx, y - 3.4, tw, 5.4, 2.7, 2.7, 'F'); doc.setFillColor(...RGB(couleur(p.classe))); doc.circle(lx + 3, y - 0.7, 1.1, 'F'); doc.setFont('helvetica', 'bold'); doc.setTextColor(14, 26, 43); doc.text(t, lx + 5.5, y + 0.4); lx += tw + 3; });
      y = pdfPlanSalle(doc, g, y + 5);
      const L = placesTriees(g), moitie = Math.ceil(L.length / 2), body = [];
      for (let i = 0; i < moitie; i++) { const a = L[i], b = L[i + moitie]; body.push([nomTri(a.nom), a.classe, String(a.s + 1)].concat(b ? [nomTri(b.nom), b.classe, String(b.s + 1)] : ['', '', ''])); }
      if (y > PH - 40) { doc.addPage(); y = TOP + 4; }
      pdfSection(doc, y, { titre: 'Liste alphabétique', cols: ['Étudiant', 'Classe', 'Place', 'Étudiant', 'Classe', 'Place'], rows: body, vide: 'Aucun étudiant placé.' },
        { styles: { fontSize: 8.4, cellPadding: { top: 1.8, bottom: 1.8, left: 3, right: 3 }, lineColor: [221, 228, 238], lineWidth: 0.2, textColor: [14, 26, 43] }, columnStyles: { 2: { halign: 'center', fontStyle: 'bold', cellWidth: 16 }, 5: { halign: 'center', fontStyle: 'bold', cellWidth: 16 } } });
    });
    pdfHabiller(doc, lg, k != null ? 'Plan de classe' : 'Plans de classe');
    doc.save(nomFichier(k != null ? 'plan salle ' + salle(G[0].salleId).num + ' ' + JOURS[G[0].jour] : 'plans de classe', 'pdf')); msg('PDF téléchargé ✓');
  } catch (e) { msg('Export PDF impossible : ' + e.message, false); }
}
async function pdfRepartition() {
  try {
    const { doc, lg } = await pdfNouveau();
    blocsRepartition().forEach((b, i) => { if (i) doc.addPage(); pdfSection(doc, TOP + 4, b, { columnStyles: { 2: { halign: 'center', fontStyle: 'bold', cellWidth: 24 } } }); });
    pdfHabiller(doc, lg, 'Répartition nominative des étudiants');
    doc.save(nomFichier('répartition nominative', 'pdf')); msg('PDF téléchargé ✓');
  } catch (e) { msg('Export PDF impossible : ' + e.message, false); }
}
async function pdfEmargement() {
  try {
    const { doc, lg } = await pdfNouveau();
    S().optim.groupes.forEach((g, i) => {
      if (i) doc.addPage();
      const L = (g.places || []).slice().sort((a, b) => a.classe.localeCompare(b.classe, 'fr') || nomTri(a.nom).localeCompare(nomTri(b.nom), 'fr'));
      pdfSection(doc, TOP + 4, { titre: 'Émargement — ' + titreGroupe(g), sous: sousGroupe(g), cols: ['Place', 'Étudiant', 'Classe', 'Signature à l\'arrivée', 'Heure de sortie'],
        rows: L.map(p => [String(p.s + 1), nomTri(p.nom), p.classe, '', '']), couleurs: L.map(p => couleur(p.classe)), vide: 'Aucun étudiant placé.' },
        { styles: { fontSize: 9, minCellHeight: 9, cellPadding: { top: 2, bottom: 2, left: 3.5, right: 3 }, lineColor: [221, 228, 238], lineWidth: 0.2, textColor: [14, 26, 43], valign: 'middle' },
          columnStyles: { 0: { halign: 'center', fontStyle: 'bold', cellWidth: 16 }, 3: { cellWidth: 72 }, 4: { cellWidth: 34 } } });
    });
    pdfHabiller(doc, lg, 'Listes d\'émargement');
    doc.save(nomFichier('émargement', 'pdf')); msg('PDF téléchargé ✓');
  } catch (e) { msg('Export PDF impossible : ' + e.message, false); }
}

window.EXSIM = { _t: { optimiser: p => optimiser(p), analyse: () => analyse(), X }, liste, nouveau, ouvrir, dupliquer, supprimer, fermer: () => fermerUI(false), onglet, enregistrer, terminer, rouvrir,
  set, setLundi, setJours, ajoutSurv, retirerSurv, dispoEdit, dispo, filtreClasse, classe, toutesClasses, salle: salleCoche, campus,
  choixClasse, clicCal, editer, glisser, ajoutLibre, dSet, dSalle, dSurv, dValider, dSupprimer, dAnnuler, dDupliquer, dupRapide,
  jourVue: k => { X.jourVue = k; rendreCorps(); }, gVue: k => { X.gVue = k; rendreCorps(); }, exportGrille, version: v => { X.etatVersion = v; rendreCorps(); },
  etatVue: k => { X.etatVue = k; X.etatFiltre = ''; rendreCorps(); }, etatFiltre: v => { X.etatFiltre = v; rendreCorps(); },
  prm, calculer, abandonProp, appliquer, retirerOptim, voirPlan, exportExcel, pdfEtat, pdfPlans, pdfRepartition, pdfEmargement };
})();
