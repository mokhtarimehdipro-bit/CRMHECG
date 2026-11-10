// =============================================================================
//  Vie scolaire — Qualiopi (33 indicateurs, preuves datées), Maintenance, Formulaires
//  Chargé par vie_scolaire.html après le script principal (utilise api(), toast(), canWrite()).
// =============================================================================
const vxEsc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const vxDate = iso => { const m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})/); return m ? `${m[3]}/${m[2]}/${m[1]}` : (iso || ''); };

// ══ QUALIOPI — référentiel national qualité (décret n° 2026-728, en vigueur au 01/11/2026) ══
const QUALIOPI_DATA = [
  { critere: 1, titre: "Conditions d'information du public", indicateurs: [
    { num: "1", modif: true, texte: "Information accessible, détaillée et vérifiable sur les prestations : prérequis, objectifs, reconnaissance, durée, modalités pédagogiques et de financement, délais, tarifs, contacts, méthodes, évaluation, accessibilité handicap — sans aucune mention de nature à induire le public en erreur.", attentes: "Fiches formations datées (site, plaquettes), CGV, captures d'écran ; vérifier l'absence de promesses non prouvées (placement, gratuité, poursuites d'études)." },
    { num: "2", modif: true, texte: "Diffusion d'indicateurs de résultats adaptés, avec leurs modalités de calcul (ou dispositifs existants).", attentes: "Taux de réussite, satisfaction, abandon, rupture, insertion publiés et datés + méthode de calcul ; lien InserJeunes." },
    { num: "3", modif: true, texte: "Taux d'obtention des certifications, validation par blocs, équivalences, passerelles, suites de parcours (dont poursuites d'études) et débouchés.", attentes: "Taux par certification, blocs de compétences, passerelles, débouchés, lien InserJeunes / France compétences." }
  ]},
  { critere: 2, titre: "Identification des objectifs et adaptation au public", indicateurs: [
    { num: "4", texte: "Analyse du besoin du bénéficiaire en lien avec l'entreprise et/ou le financeur.", attentes: "Grille d'entretien amont, fiche de poste, dossier de candidature, compte rendu d'analyse du besoin." },
    { num: "5", texte: "Objectifs opérationnels et évaluables de la prestation.", attentes: "Objectifs formulés en verbes d'action, programmes et plaquettes." },
    { num: "6", texte: "Contenus et modalités de mise en œuvre adaptés aux objectifs et aux publics.", attentes: "Progressions pédagogiques de l'année, emplois du temps, syllabus." },
    { num: "7", modif: true, texte: "Adéquation des contenus aux exigences de la certification visée, et capacité à assurer cette certification (y compris comme organisme habilité).", attentes: "Tableaux croisés référentiel / programme, habilitation ou convention avec le certificateur, inscriptions aux examens." },
    { num: "8", texte: "Procédures de positionnement et d'évaluation des acquis à l'entrée.", attentes: "Fiches de positionnement renseignées, tests d'entrée, décisions d'admission." }
  ]},
  { critere: 3, titre: "Adaptation des prestations, accueil, suivi et évaluation", indicateurs: [
    { num: "9", texte: "Information des bénéficiaires sur les conditions de déroulement.", attentes: "Livret d'accueil daté, règlement intérieur, calendrier, preuve de remise." },
    { num: "10", texte: "Mise en œuvre et adaptation de la prestation, de l'accompagnement et du suivi.", attentes: "Parcours individualisés, aménagements, suivi personnalisé tracé." },
    { num: "11", texte: "Évaluation de l'atteinte des objectifs par les bénéficiaires.", attentes: "Relevés de notes, DST, examens blancs, livrets scolaires, attestations de fin de formation." },
    { num: "12", modif: true, texte: "Engagement des bénéficiaires, prévention des ruptures, et prévention / traitement de toute violence (dont VSS), harcèlement ou discrimination.", attentes: "Procédure ruptures, suivi des absences, procédure et référent violences/harcèlement/discriminations, affichages." },
    { num: "13", cfa: true, texte: "Alternance : anticipation des missions avec l'entreprise, coordination et progressivité centre / entreprise.", attentes: "Livret de suivi renseigné et signé, comptes rendus de visites, outil de liaison (NC mineure en 2024)." },
    { num: "14", modif: true, cfa: true, texte: "Accompagnement socio-professionnel, éducatif et citoyen ; procédure de traitement sans délai des ruptures liées à des difficultés, violences ou discriminations.", attentes: "Aides (permis, équipement, logement), actions mixité / égalité (NC majeure en 2024), procédure de traitement immédiat." },
    { num: "15", modif: true, cfa: true, texte: "Information des apprentis sur leurs droits et devoirs, santé et sécurité ; information renforcée (mineurs, VSS, médiateur de l'apprentissage, inspection du travail).", attentes: "Livret / page droits et devoirs, SST, coordonnées du médiateur, preuve d'information." },
    { num: "16", texte: "Conditions de présentation à la certification conformes aux exigences du certificateur.", attentes: "Inscriptions et convocations aux examens de l'année." }
  ]},
  { critere: 4, titre: "Adéquation des moyens pédagogiques, techniques et d'encadrement", indicateurs: [
    { num: "17", texte: "Moyens humains et techniques adaptés, environnement approprié (locaux, équipements).", attentes: "Baux, DUERP de l'année, liste des équipements, registre d'accessibilité." },
    { num: "18", texte: "Mobilisation et coordination des intervenants internes et externes.", attentes: "Organigramme de l'année, liste des formateurs, CR de réunions pédagogiques." },
    { num: "19", modif: true, texte: "Ressources pédagogiques mises à disposition et appropriation ; suivi effectif des modules à distance ; référent pédagogique par formation.", attentes: "Supports, plateformes, traces de connexion, désignation des référents pédagogiques." },
    { num: "20", modif: true, cfa: true, texte: "Personnel dédié à la mobilité nationale et internationale, référent handicap, conseil de perfectionnement ; gouvernance participative (apprentis, formateurs, entreprises).", attentes: "Nominations des référents, CR des conseils de perfectionnement, composition." }
  ]},
  { critere: 5, titre: "Qualification et développement des connaissances des personnels", indicateurs: [
    { num: "21", texte: "Détermination, mobilisation et évaluation des compétences des intervenants.", attentes: "CV datés, diplômes, matrice de compétences, évaluations des formateurs." },
    { num: "22", texte: "Entretien et développement des compétences des salariés.", attentes: "Plan de développement des compétences, attestations de formation, entretiens professionnels." }
  ]},
  { critere: 6, titre: "Inscription dans l'environnement professionnel", indicateurs: [
    { num: "23", texte: "Veille légale et réglementaire sur la formation professionnelle, et exploitation.", attentes: "Abonnements, notes de veille récentes, CR montrant l'exploitation." },
    { num: "24", texte: "Veille sur les compétences, métiers et emplois, et exploitation.", attentes: "Congrès OEC, revues professionnelles, intégration dans les cours." },
    { num: "25", texte: "Veille sur les innovations pédagogiques et technologiques, et exploitation.", attentes: "Outils testés, formations, retours d'expérience." },
    { num: "26", texte: "Mobilisation des expertises, outils et réseaux pour accueillir et accompagner les publics en situation de handicap.", attentes: "Réseau (AGEFIPH, RHF), conventions, aménagements réalisés, badge handicap." },
    { num: "27", modif: true, texte: "Sous-traitance / portage salarial : conformité au référentiel et traçabilité dans les contrats.", attentes: "Contrats avec clause Qualiopi — ou justification de non-application." },
    { num: "28", texte: "Formation en situation de travail : réseau de partenaires pour coconstruire l'ingénierie et favoriser l'accueil en entreprise.", attentes: "Conventions et partenariats, CVthèque, réseau cabinets et entreprises." },
    { num: "29", cfa: true, texte: "Actions concourant à l'insertion professionnelle ou à la poursuite d'études.", attentes: "Job dating, offres diffusées, partenariats France Travail, ateliers CV." }
  ]},
  { critere: 7, titre: "Recueil et prise en compte des appréciations et réclamations", indicateurs: [
    { num: "30", modif: true, texte: "Recueil des appréciations des parties prenantes : bénéficiaires, financeurs, équipes pédagogiques, entreprises.", attentes: "Enquêtes de l'année (apprentis, entreprises, formateurs, OPCO) et taux de retour." },
    { num: "31", texte: "Traitement des difficultés, réclamations et aléas.", attentes: "Formulaire de réclamation, registre, tickets de maintenance traités, réponses." },
    { num: "32", modif: true, texte: "Mesures d'amélioration à partir des appréciations et réclamations, et analyse des risques sur la qualité des formations.", attentes: "Plan d'amélioration continue daté, analyse des risques, CR de revue qualité." },
    { num: "33", nouveau: true, cfa: true, texte: "Évaluation des contenus et des enseignements par les apprenants (distincte de la satisfaction), résultats partagés avec les équipes, démarche d'amélioration et mesure de son efficacité.", attentes: "Questionnaire d'évaluation des enseignements (module Formulaires), analyse, CR de partage, plan d'actions." }
  ]}
];
const CRITERE_COLORS = { 1: '#2563eb', 2: '#4338ca', 3: '#7c3aed', 4: '#9333ea', 5: '#db2777', 6: '#e11d48', 7: '#ea580c' };
let qualiopiActions = [], qFiltre = { texte: '', etat: '' };

// Statut d'une preuve : à jour / à revoir (< 60 jours) / obsolète
function qStatut(p) {
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const m = String(p.nomPreuve || '').match(/(20\d\d)\s*[-\/_ ]\s*(20\d\d)/);
  const finAnnee = today.getMonth() >= 7 ? today.getFullYear() + 1 : today.getFullYear();
  if (m && Number(m[2]) > Number(m[1]) && Number(m[2]) < finAnnee) return { etat: 'obsolete', raison: `année ${m[1]}-${m[2]} terminée` };
  const val = p.validite === '' || p.validite === undefined || p.validite === null ? (p.type === 'document' ? 24 : 12) : Number(p.validite);
  if (val === 0) return { etat: 'ok', raison: 'validité permanente' };
  const d = String(p.dateDocument || p.dateDepot || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!d) return { etat: 'inconnu', raison: 'date du document inconnue' };
  const exp = new Date(+d[1], +d[2] - 1 + val, +d[3]);
  const j = Math.round((exp - today) / 86400000);
  const expTxt = exp.toLocaleDateString('fr-FR');
  if (j < 0) return { etat: 'obsolete', raison: `périmé depuis le ${expTxt}` };
  if (j <= 60) return { etat: 'bientot', raison: `à revoir avant le ${expTxt}` };
  return { etat: 'ok', raison: `valable jusqu'au ${expTxt}` };
}
const Q_PASTILLE = { ok: ['#dcfce7', '#15803d', '✓ À jour'], bientot: ['#fef3c7', '#b45309', '⏳ À revoir'], obsolete: ['#fee2e2', '#b91c1c', '⚠ Obsolète'], inconnu: ['#f3f4f6', '#6b7280', '? Non daté'] };

async function loadQualiopi() {
  const [r, ra] = await Promise.all([api('getQualiopiPreuves'), api('getQualiopiActions').catch(() => ({}))]);
  qualiopiPreuves = r.success ? r.data : [];
  qualiopiActions = ra && ra.success ? ra.data : [];
  renderQualiopi();
  if (canWrite('qualiopi') && typeof loadQuestionnaireSurvey === 'function') loadQuestionnaireSurvey();
}
function qLignePreuve(p) {
  const st = qStatut(p), [bg, fg, lib] = Q_PASTILLE[st.etat];
  const ic = p.captureUrl ? '🖼️' : p.type === 'lien' ? '🔗' : '📄';
  return `<div class="qx-preuve">
    <div style="min-width:0;flex:1">
      <div style="display:flex;align-items:center;gap:.4rem;flex-wrap:wrap">
        <span>${ic}</span>
        <a href="${vxEsc(p.urlFichier)}" target="_blank" rel="noopener" class="qx-titre" title="${vxEsc(p.nomPreuve)}">${vxEsc(p.nomPreuve)}</a>
        <span class="qx-pill" style="background:${bg};color:${fg}" title="${vxEsc(st.raison)}">${lib}</span>
      </div>
      ${p.commentaire ? `<p class="qx-com">${vxEsc(p.commentaire)}</p>` : ''}
      <p class="qx-meta">Document du ${vxDate(p.dateDocument) || '—'} · ${vxEsc(st.raison)}${p.source ? ' · ' + vxEsc(p.source) : ''}</p>
    </div>
    <div class="qx-actions">
      ${p.captureUrl ? `<a href="${vxEsc(p.captureUrl)}" target="_blank" rel="noopener" title="Capture d'écran datée">🖼 Capture</a>` : ''}
      ${canWrite('qualiopi') ? `<button onclick="editQualiopiPreuve('${vxEsc(p.id)}')" title="Dater / modifier">✏️</button><button onclick="deleteQualiopiPreuve('${vxEsc(p.id)}')" title="Retirer" class="qx-del">🗑</button>` : ''}
    </div></div>`;
}
function renderQualiopi() {
  const par = {};
  qualiopiPreuves.forEach(p => { (par[p.indicateur] = par[p.indicateur] || []).push(p); });
  let total = 0, couverts = 0, nbObs = 0, nbBientot = 0;
  qualiopiPreuves.forEach(p => { const e = qStatut(p).etat; if (e === 'obsolete') nbObs++; if (e === 'bientot') nbBientot++; });
  const t = (qFiltre.texte || '').toLowerCase();
  const garde = p => (!t || (p.nomPreuve + ' ' + p.commentaire).toLowerCase().includes(t)) && (!qFiltre.etat || qFiltre.etat === 'sans' || qStatut(p).etat === qFiltre.etat);
  const actionsOuvertes = qualiopiActions.filter(a => a.statut !== 'Fait');
  let html = `<style>
    .qx-bar{display:flex;flex-wrap:wrap;gap:.6rem;align-items:center;margin-bottom:1.25rem}
    .qx-bar input,.qx-bar select{border:1px solid #e5e7eb;border-radius:.75rem;padding:.55rem .8rem;font-size:.85rem}
    .qx-kpi{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:.75rem;margin-bottom:1.25rem}
    .qx-kpi div{background:#fff;border:1px solid #f1f5f9;border-radius:1rem;padding:.9rem 1rem}
    .qx-kpi b{display:block;font-size:1.5rem;font-weight:900}.qx-kpi span{font-size:.7rem;font-weight:800;text-transform:uppercase;color:#9ca3af}
    .qx-preuve{display:flex;justify-content:space-between;gap:.75rem;background:#fff;border:1px solid #f1f5f9;border-radius:.85rem;padding:.6rem .8rem}
    .qx-titre{font-size:.85rem;font-weight:700;color:#1f2937;text-decoration:none;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:420px}
    .qx-titre:hover{color:#1A4E8A;text-decoration:underline}
    .qx-pill{font-size:.65rem;font-weight:800;padding:.15rem .5rem;border-radius:999px;white-space:nowrap}
    .qx-com{font-size:.75rem;color:#4b5563;margin-top:.2rem}.qx-meta{font-size:.68rem;color:#9ca3af;margin-top:.15rem}
    .qx-actions{display:flex;gap:.3rem;align-items:flex-start;flex-shrink:0}
    .qx-actions a,.qx-actions button{background:#f3f4f6;color:#4b5563;border:0;border-radius:.5rem;padding:.3rem .55rem;font-size:.7rem;font-weight:700;cursor:pointer;text-decoration:none}
    .qx-actions .qx-del{background:#fef2f2;color:#dc2626}
    .qx-badge{font-size:.62rem;font-weight:900;padding:.15rem .5rem;border-radius:999px;margin-left:.3rem}
    details.qx-liste>summary{cursor:pointer;font-size:.75rem;font-weight:800;color:#1A4E8A;margin-top:.6rem;list-style:none}
    details.qx-liste>summary::-webkit-details-marker{display:none}
    .qx-act{display:flex;gap:.6rem;align-items:flex-start;padding:.55rem .2rem;border-bottom:1px dashed #f1f5f9;font-size:.82rem}
    .qx-act input{margin-top:.2rem;width:16px;height:16px;flex-shrink:0}
  </style>`;
  // Indicateurs couverts : au moins une preuve non obsolète
  QUALIOPI_DATA.forEach(cr => cr.indicateurs.forEach(ind => { total++; if ((par[ind.num] || []).some(p => qStatut(p).etat !== 'obsolete')) couverts++; }));
  html += `<div class="qx-kpi">
    <div><b style="color:#1A4E8A">${couverts}/${total}</b><span>indicateurs avec preuve à jour</span></div>
    <div><b>${qualiopiPreuves.length}</b><span>preuves enregistrées</span></div>
    <div><b style="color:#b91c1c">${nbObs}</b><span>preuves obsolètes</span></div>
    <div><b style="color:#b45309">${nbBientot}</b><span>à revoir sous 60 j</span></div>
    <div><b style="color:#ea580c">${actionsOuvertes.length}</b><span>actions à mener</span></div></div>
  <div class="qx-bar">
    <input type="search" placeholder="🔎 Rechercher une preuve…" value="${vxEsc(qFiltre.texte)}" oninput="qFiltre.texte=this.value;renderQualiopi()" style="flex:1;min-width:200px">
    <select onchange="qFiltre.etat=this.value;renderQualiopi()">
      ${[['', 'Toutes les preuves'], ['obsolete', '⚠ Obsolètes'], ['bientot', '⏳ À revoir'], ['ok', '✓ À jour'], ['sans', '⬜ Indicateurs sans preuve à jour']].map(o => `<option value="${o[0]}" ${qFiltre.etat === o[0] ? 'selected' : ''}>${o[1]}</option>`).join('')}
    </select>
    <span style="font-size:.72rem;color:#6b7280">Référentiel : décret n° 2026-728 — 33 indicateurs (en vigueur depuis le 01/11/2026)</span>
  </div>`;
  // Plan d'actions
  if (qualiopiActions.length) {
    const ordre = { Haute: 0, Moyenne: 1, Basse: 2 };
    const liste = qualiopiActions.slice().sort((a, b) => (a.statut === 'Fait') - (b.statut === 'Fait') || (ordre[a.priorite] ?? 3) - (ordre[b.priorite] ?? 3) || Number(a.indicateur) - Number(b.indicateur));
    const coul = { Haute: '#dc2626', Moyenne: '#d97706', Basse: '#6b7280' };
    html += `<details class="qualiopi-card" style="border-left:4px solid #ea580c;margin-bottom:1.5rem" ${actionsOuvertes.length ? 'open' : ''}>
      <summary style="cursor:pointer;font-weight:900;color:#1f2937">🧭 Plan d'actions avant audit — ${actionsOuvertes.length} à faire / ${qualiopiActions.length}</summary>
      <div style="margin-top:.75rem">${liste.map(a => `<div class="qx-act" style="${a.statut === 'Fait' ? 'opacity:.5;text-decoration:line-through' : ''}">
        <input type="checkbox" ${a.statut === 'Fait' ? 'checked' : ''} ${canWrite('qualiopi') ? '' : 'disabled'} onchange="qualiopiActionStatut('${vxEsc(a.id)}', this.checked)">
        <span style="font-weight:900;color:${coul[a.priorite] || '#374151'};min-width:62px">Ind. ${vxEsc(a.indicateur)}</span>
        <span style="flex:1">${vxEsc(a.action)}</span>
        <span class="qx-pill" style="background:#f8fafc;color:${coul[a.priorite] || '#374151'}">${vxEsc(a.priorite)}</span></div>`).join('')}</div></details>`;
  }
  // Zone libre
  const libres = (par['LIBRE'] || []).filter(garde);
  html += `<div class="qualiopi-card" style="border:2px dashed #93c5fd;background:linear-gradient(180deg,#f8fbff,#fff);margin-bottom:1.75rem">
    <div style="display:flex;justify-content:space-between;gap:1rem;align-items:flex-start">
      <div><p style="font-size:.95rem;font-weight:800">🗂 Zone libre <span style="font-size:.7rem;color:#9ca3af">${(par['LIBRE'] || []).length} document(s)</span></p>
      <p style="font-size:.75rem;color:#6b7280">Rapports d'audit, fiches de non-conformité, référentiel, tableaux de pilotage…</p></div>
      ${canWrite('qualiopi') ? `<button onclick="openQualiopiModal('LIBRE','0')" class="btn-orange" style="font-size:.75rem;padding:.45rem .8rem">+ Ajouter</button>` : ''}
    </div>${libres.length ? `<div style="display:flex;flex-direction:column;gap:.4rem;margin-top:.75rem">${libres.map(qLignePreuve).join('')}</div>` : ''}</div>`;
  QUALIOPI_DATA.forEach(cr => {
    const cc = CRITERE_COLORS[cr.critere] || '#374151';
    let bloc = '';
    cr.indicateurs.forEach(ind => {
      const toutes = par[ind.num] || [], aJour = toutes.filter(p => qStatut(p).etat !== 'obsolete');
      if (qFiltre.etat === 'sans' && aJour.length) return;
      const visibles = toutes.filter(garde);
      if ((qFiltre.texte || (qFiltre.etat && qFiltre.etat !== 'sans')) && !visibles.length) return;
      const obs = toutes.length - aJour.length;
      const bg = aJour.length ? '#f0fdf4' : '#fff', brd = aJour.length ? '#bbf7d0' : (toutes.length ? '#fecaca' : '#f3f4f6');
      const ouvrir = qFiltre.texte || qFiltre.etat || visibles.length <= 3;
      bloc += `<div class="qualiopi-card" style="background:${bg};border-color:${brd}">
        <div style="display:flex;justify-content:space-between;gap:1rem;align-items:flex-start">
          <div style="flex:1">
            <div style="display:flex;align-items:center;gap:.4rem;flex-wrap:wrap;margin-bottom:.45rem">
              <span style="font-size:.75rem;font-weight:900;color:${cc}">Indicateur ${ind.num}</span>
              ${ind.nouveau ? '<span class="qx-badge" style="background:#fee2e2;color:#b91c1c">NOUVEAU 2026</span>' : ''}
              ${ind.modif ? '<span class="qx-badge" style="background:#fef3c7;color:#92400e">MODIFIÉ 2026</span>' : ''}
              ${ind.cfa ? '<span class="qx-badge" style="background:#e0e7ff;color:#3730a3">CFA</span>' : ''}
              <span class="qx-pill" style="background:${aJour.length ? '#dcfce7' : '#f3f4f6'};color:${aJour.length ? '#15803d' : '#9ca3af'}">${aJour.length ? '✓ ' + aJour.length + ' preuve(s) à jour' : '⬜ Aucune preuve à jour'}</span>
              ${obs ? `<span class="qx-pill" style="background:#fee2e2;color:#b91c1c">⚠ ${obs} obsolète(s)</span>` : ''}
            </div>
            <p style="font-size:.85rem;font-weight:600;color:#1f2937">${ind.texte}</p>
            <p style="font-size:.75rem;color:#6b7280;font-style:italic;margin-top:.3rem">📋 <strong>Preuves attendues :</strong> ${ind.attentes}</p>
          </div>
          ${canWrite('qualiopi') ? `<div style="display:flex;flex-direction:column;gap:.4rem;flex-shrink:0">
            <button onclick="openQualiopiModal('${ind.num}','${cr.critere}')" style="background:#1A4E8A;color:#fff;padding:.4rem .75rem;border-radius:.6rem;font-weight:700;font-size:.75rem;border:none;cursor:pointer">+ Ajouter</button>
            <button onclick="ouvrirDossierQualiopi('${ind.num}','${cr.critere}')" style="background:#f3f4f6;color:#4b5563;padding:.4rem .75rem;border-radius:.6rem;font-weight:700;font-size:.75rem;border:none;cursor:pointer">📂 Dossier</button></div>` : ''}
        </div>
        ${visibles.length ? `<details class="qx-liste" ${ouvrir ? 'open' : ''}><summary>▸ ${visibles.length} preuve(s)</summary><div style="display:flex;flex-direction:column;gap:.4rem;margin-top:.5rem">${visibles.sort((a, b) => String(b.dateDocument).localeCompare(String(a.dateDocument))).map(qLignePreuve).join('')}</div></details>` : ''}
      </div>`;
    });
    if (bloc) html += `<div style="margin-bottom:1.5rem"><div style="display:flex;align-items:center;gap:.75rem;margin-bottom:.75rem">
      <span style="background:${cc};color:#fff;padding:.2rem .75rem;border-radius:999px;font-size:.75rem;font-weight:900">Critère ${cr.critere}</span>
      <h2 style="font-size:1.1rem;font-weight:900;color:#1f2937">${cr.titre}</h2></div>${bloc}</div>`;
  });
  document.getElementById('qualiopiContent').innerHTML = html;
  document.getElementById('qualiopiPct').textContent = Math.round(couverts / total * 100) + '%';
}
async function qualiopiActionStatut(id, fait) {
  const r = await api('saveQualiopiAction', { id, statut: fait ? 'Fait' : 'À faire' });
  if (!r.success) { toast('Erreur : ' + r.message, false); return; }
  const a = qualiopiActions.find(x => x.id === id); if (a) a.statut = fait ? 'Fait' : 'À faire';
  renderQualiopi();
}

// ── Ajout / modification d'une preuve : document OU lien web ──
let qEdition = null;
function qModal() {
  let m = document.getElementById('qxModal');
  if (m) return m;
  m = document.createElement('div');
  m.id = 'qxModal'; m.className = 'modal-overlay hidden';
  m.innerHTML = `<div class="modal-box p-6" style="max-width:560px">
    <h3 class="text-xl font-black text-[#1A4E8A] mb-1" id="qxTitre">Ajouter une preuve</h3>
    <p class="text-sm text-blue-700 font-bold mb-4" id="qxInd"></p>
    <div id="qxChoix" style="display:grid;grid-template-columns:1fr 1fr;gap:.6rem;margin-bottom:1rem">
      <button type="button" data-t="document" onclick="qxType('document')" class="qx-choix">📎<b>Document</b><span>PDF, Word, Excel, image…</span></button>
      <button type="button" data-t="lien" onclick="qxType('lien')" class="qx-choix">🔗<b>Lien web</b><span>Page du site, Drive, plateforme…</span></button>
    </div>
    <style>.qx-choix{display:flex;flex-direction:column;align-items:center;gap:.15rem;padding:.9rem;border:2px solid #e5e7eb;border-radius:1rem;background:#fff;cursor:pointer;font-size:1.4rem}
      .qx-choix b{font-size:.85rem;color:#1f2937}.qx-choix span{font-size:.68rem;color:#9ca3af}.qx-choix.on{border-color:#1A4E8A;background:#eff6ff}</style>
    <div class="space-y-3">
      <div><label>Libellé de la preuve</label><input type="text" id="qxNom" placeholder="ex. : Livret d'accueil 2026-2027"></div>
      <div id="qxBlocFichier"><label>Fichier</label><input type="file" id="qxFichier" accept=".pdf,.jpg,.jpeg,.png,.docx,.xlsx,.pptx,.doc,.xls"></div>
      <div id="qxBlocLien" class="hidden"><label>Adresse (URL)</label><input type="url" id="qxUrl" placeholder="https://hecg.fr/…"></div>
      <div><label>Commentaire (ce que la preuve démontre)</label><textarea id="qxCom" rows="2" placeholder="ex. : prérequis, durée et tarifs publiés sur la fiche BTS CG"></textarea></div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:.6rem">
        <div><label>Date du document *</label><input type="date" id="qxDate" required></div>
        <div><label>Valable pendant</label><select id="qxVal"><option value="6">6 mois</option><option value="12">1 an</option><option value="24">2 ans</option><option value="36">3 ans</option><option value="0">Permanent</option></select></div>
      </div>
      <p class="text-xs text-gray-400">La preuve passe automatiquement en « À revoir » 60 jours avant son échéance, puis en « Obsolète ». Un libellé contenant une année scolaire terminée (ex. 2024-2025) est aussi signalé obsolète.</p>
    </div>
    <div class="flex justify-end gap-3 mt-5"><button onclick="document.getElementById('qxModal').classList.add('hidden')" class="btn-ghost">Annuler</button>
      <button onclick="saveQualiopiPreuve()" id="qxSave" class="btn-orange">Enregistrer</button></div></div>`;
  document.body.appendChild(m);
  return m;
}
function qxType(t) {
  qEdition.type = t;
  document.querySelectorAll('#qxChoix .qx-choix').forEach(b => b.classList.toggle('on', b.dataset.t === t));
  document.getElementById('qxBlocFichier').classList.toggle('hidden', t !== 'document');
  document.getElementById('qxBlocLien').classList.toggle('hidden', t !== 'lien');
  const v = document.getElementById('qxVal'); if (!qEdition.id) v.value = t === 'lien' ? '12' : '24';
}
function openQualiopiModal(num, critere) {
  const m = qModal(); qEdition = { num, critere, type: 'document' };
  document.getElementById('qxTitre').textContent = 'Ajouter une preuve';
  document.getElementById('qxInd').textContent = num === 'LIBRE' ? 'Zone libre — document hors indicateur' : 'Indicateur ' + num + ' (critère ' + critere + ')';
  document.getElementById('qxChoix').style.display = '';
  ['qxNom', 'qxUrl', 'qxCom'].forEach(i => document.getElementById(i).value = '');
  document.getElementById('qxFichier').value = '';
  document.getElementById('qxDate').value = new Date().toISOString().slice(0, 10);
  qxType('document');
  m.classList.remove('hidden');
}
function closeQualiopiModal() { const m = document.getElementById('qxModal'); if (m) m.classList.add('hidden'); }
function editQualiopiPreuve(id) {
  const p = qualiopiPreuves.find(x => x.id === id); if (!p) return;
  const m = qModal(); qEdition = { id, num: p.indicateur, type: p.type };
  document.getElementById('qxTitre').textContent = 'Dater / modifier la preuve';
  document.getElementById('qxInd').textContent = p.indicateur === 'LIBRE' ? 'Zone libre' : 'Indicateur ' + p.indicateur;
  document.getElementById('qxChoix').style.display = 'none';
  document.getElementById('qxBlocFichier').classList.add('hidden');
  document.getElementById('qxBlocLien').classList.add('hidden');
  document.getElementById('qxNom').value = p.nomPreuve;
  document.getElementById('qxCom').value = p.commentaire || '';
  document.getElementById('qxDate').value = p.dateDocument || '';
  const v = p.validite === '' ? (p.type === 'document' ? 24 : 12) : p.validite;
  document.getElementById('qxVal').value = String([0, 6, 12, 24, 36].includes(Number(v)) ? v : 24);
  m.classList.remove('hidden');
}
async function saveQualiopiPreuve() {
  const nom = document.getElementById('qxNom').value.trim(), com = document.getElementById('qxCom').value.trim();
  const dateDocument = document.getElementById('qxDate').value, validite = document.getElementById('qxVal').value;
  if (!nom) { toast('Saisissez un libellé.', false); return; }
  if (!dateDocument) { toast('Indiquez la date du document : tous les éléments Qualiopi doivent être datés.', false); return; }
  const btn = document.getElementById('qxSave'); btn.disabled = true; btn.textContent = 'Enregistrement…';
  const fin = (r, msg) => { btn.disabled = false; btn.textContent = 'Enregistrer'; if (!r.success) { toast('Erreur : ' + r.message, false); return; } toast(msg); closeQualiopiModal(); loadQualiopi(); };
  if (qEdition.id) { fin(await api('updateQualiopiPreuve', { id: qEdition.id, nomPreuve: nom, commentaire: com, dateDocument, validite }), 'Preuve mise à jour ✓'); return; }
  if (qEdition.type === 'lien') {
    const url = document.getElementById('qxUrl').value.trim();
    if (!/^https?:\/\//i.test(url)) { toast('Saisissez une adresse commençant par https://', false); btn.disabled = false; btn.textContent = 'Enregistrer'; return; }
    fin(await api('saveQualiopiPreuve', { indicateur: qEdition.num, nomPreuve: nom, urlFichier: url, commentaire: com, dateDocument, validite }), 'Lien enregistré ✓');
    return;
  }
  const file = document.getElementById('qxFichier').files[0];
  if (!file) { toast('Sélectionnez un fichier.', false); btn.disabled = false; btn.textContent = 'Enregistrer'; return; }
  const reader = new FileReader();
  reader.onload = async e => fin(await api('uploadQualiopiFile', { base64Data: e.target.result.split(',')[1], mimeType: file.type || 'application/octet-stream', fileName: file.name,
    indicateur: qEdition.num, critere: qEdition.critere, nomPreuve: nom, commentaire: com, dateDocument, validite }), 'Document déposé ✓');
  reader.readAsDataURL(file);
}
async function deleteQualiopiPreuve(id) {
  const p = qualiopiPreuves.find(x => x.id === id);
  if (!confirm('Retirer la preuve « ' + (p ? p.nomPreuve : '') + ' » ?\n(le fichier reste dans le Drive)')) return;
  const r = await api('deleteQualiopiPreuve', { id });
  if (!r.success) { toast('Erreur : ' + r.message, false); return; }
  toast('Preuve retirée'); loadQualiopi();
}

// ══ TICKETS DE MAINTENANCE ══
let maintData = null, maintFiltre = { statut: 'ouverts', texte: '' }, maintSel = new Set();
let maintVue = (() => { try { return localStorage.getItem('maintVue') || 'liste'; } catch (e) { return 'liste'; } })();
const MAINT_STATUTS = ['Nouveau', 'Pris en charge', 'En attente', 'Résolu', 'Clôturé', 'Annulé', 'Sans suite'];
const MAINT_FERMES = ['Résolu', 'Clôturé', 'Annulé', 'Sans suite'];
const MAINT_COUL = { 'Nouveau': ['#dbeafe', '#1d4ed8'], 'Pris en charge': ['#fef3c7', '#b45309'], 'En attente': ['#ede9fe', '#6d28d9'], 'Résolu': ['#dcfce7', '#15803d'],
  'Clôturé': ['#d1fae5', '#065f46'], 'Annulé': ['#fee2e2', '#991b1b'], 'Sans suite': ['#f3f4f6', '#6b7280'] };
const maintOuvert = t => !MAINT_FERMES.includes(t.statut);
const maintUrgent = t => /urgent|haute|critique|bloquant/i.test(t.urgence + ' ' + t.priorite);
async function loadMaintenance() {
  const s = document.getElementById('section-maintenance');
  if (!maintData) s.innerHTML = `<div class="text-center text-gray-400 py-16">Chargement des tickets…</div>`;
  const r = await api('getTicketsMaintenance');
  if (!r.success) {
    s.innerHTML = `<h1 class="text-2xl font-black text-[#1A4E8A] mb-4">🛠️ Tickets de maintenance</h1><div class="qualiopi-card" style="border-left:4px solid #dc2626">
      <p class="font-bold text-red-700">${vxEsc(r.message)}</p>${r.lien ? `<a href="${vxEsc(r.lien)}" target="_blank" class="text-sm text-blue-600 underline">Ouvrir le classeur des réponses</a>` : ''}</div>`;
    return;
  }
  // Nouveaux tickets depuis la dernière visite (contrôle d'arrivée)
  let vus = []; try { vus = JSON.parse(localStorage.getItem('maintVus') || '[]'); } catch (e) {}
  r.nouveaux = vus.length ? r.tickets.filter(t => !vus.includes(t.cle)).map(t => t.cle) : [];
  try { localStorage.setItem('maintVus', JSON.stringify(r.tickets.map(t => t.cle))); } catch (e) {}
  maintData = r; maintSel = new Set([...maintSel].filter(c => r.tickets.some(t => t.cle === c)));
  renderMaintenance();
}
function maintListe() {
  const T = maintData.tickets, q = maintFiltre.texte.toLowerCase(), f = maintFiltre.statut;
  return T.filter(t => t._enCours || (f === 'archives' ? t.archive : !t.archive)
    && (t._enCours || f === 'tous' || f === 'archives' || (f === 'ouverts' ? maintOuvert(t) : t.statut === f))
    && (!q || [t.numero, t.description, t.lieu, t.categorie, t.nom, t.email, t.campus, t.commentaire, t.responsable, t.statut].join(' ').toLowerCase().includes(q)));
}
function renderMaintenance() {
  const s = document.getElementById('section-maintenance'), r = maintData, T = r.tickets, A = T.filter(t => !t.archive);
  const ouverts = A.filter(maintOuvert), urg = ouverts.filter(maintUrgent), liste = maintListe(), ctl = r.controle || {};
  const cats = {}; ouverts.forEach(t => { const c = t.categorie || 'Autre'; cats[c] = (cats[c] || 0) + 1; });
  const synchroOk = ctl.lignesSheet == null || ctl.lignesSheet === ctl.affiches;
  const visibles = new Set(liste.map(t => t.cle)), nbSel = [...maintSel].filter(c => visibles.has(c)).length;
  s.innerHTML = `<style>
    .mt-kpi{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:.75rem;margin:1rem 0 1.25rem}
    .mt-kpi div{border-radius:1.1rem;padding:1rem 1.1rem;color:#fff;box-shadow:0 12px 24px -16px rgba(15,23,42,.5);cursor:pointer}
    .mt-kpi b{display:block;font-size:1.8rem;font-weight:900;line-height:1}.mt-kpi span{font-size:.72rem;font-weight:700;opacity:.9}
    .mt-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:.9rem}
    .mt-card{background:#fff;border:1px solid #eef2f7;border-radius:1.1rem;padding:1rem 1.1rem;display:flex;flex-direction:column;gap:.5rem;transition:transform .15s,box-shadow .15s;cursor:pointer;position:relative;overflow:hidden}
    .mt-card:hover{transform:translateY(-2px);box-shadow:0 14px 30px -18px rgba(26,78,138,.5)}
    .mt-card::before{content:"";position:absolute;left:0;top:0;bottom:0;width:4px;background:var(--bord)}
    .mt-pill{font-size:.65rem;font-weight:800;padding:.18rem .55rem;border-radius:999px;white-space:nowrap}
    .mt-chip{background:#f1f5f9;color:#475569;font-size:.68rem;font-weight:700;padding:.15rem .5rem;border-radius:.5rem}
    .mt-tab{width:100%;border-collapse:separate;border-spacing:0;font-size:.82rem}
    .mt-tab th{position:sticky;top:0;background:#f8fafc;text-align:left;font-size:.66rem;text-transform:uppercase;color:#64748b;padding:.55rem .6rem;border-bottom:1px solid #e2e8f0;white-space:nowrap}
    .mt-tab td{padding:.55rem .6rem;border-bottom:1px solid #f1f5f9;vertical-align:top}
    .mt-tab tr:hover td{background:#f8fbff}.mt-tab tr.sel td{background:#eff6ff}
    .mt-act{display:inline-flex;align-items:center;justify-content:center;width:1.8rem;height:1.8rem;border-radius:.55rem;border:1px solid #e2e8f0;background:#fff;font-size:.8rem;cursor:pointer}
    .mt-act:hover{background:#f1f5f9}.mt-act[disabled]{opacity:.35;cursor:default}
    .mt-bulk{position:sticky;top:0;z-index:5;display:flex;flex-wrap:wrap;gap:.5rem;align-items:center;background:#1A4E8A;color:#fff;border-radius:.9rem;padding:.55rem .8rem;margin-bottom:.75rem}
    .mt-bulk button{background:rgba(255,255,255,.15);border-radius:.6rem;padding:.3rem .7rem;font-size:.75rem;font-weight:800}
    .mt-bulk button:hover{background:rgba(255,255,255,.28)}
    .mt-new{background:#f97316;color:#fff;font-size:.6rem;font-weight:900;padding:.1rem .4rem;border-radius:999px;margin-left:.3rem}
  </style>
  <div class="flex flex-wrap justify-between items-start gap-3">
    <div><h1 class="text-2xl font-black text-[#1A4E8A]">🛠️ Tickets de maintenance</h1>
      <p class="text-sm text-gray-500 mt-1">Signalements reçus par le formulaire « ${vxEsc(r.titre)} » — le suivi est enregistré dans le CRM, les réponses d'origine ne sont pas modifiées.</p></div>
    <div class="flex gap-2 flex-wrap">${r.formUrl ? `<button onclick="navigator.clipboard.writeText('${vxEsc(r.formUrl)}');toast('Lien du formulaire copié ✓')" class="btn-ghost text-sm">📋 Copier le lien du formulaire</button><a href="${vxEsc(r.formUrl)}" target="_blank" class="btn-ghost text-sm">📝 Ouvrir le formulaire</a>` : ''}
      <a href="${vxEsc(r.sheetUrl)}" target="_blank" class="btn-ghost text-sm">📊 Réponses (Sheet)</a><button onclick="loadMaintenance()" class="btn-orange text-sm">↻ Actualiser</button></div>
  </div>
  <div class="mt-2 text-xs rounded-xl px-3 py-2 inline-flex flex-wrap gap-x-3 gap-y-1 items-center" style="background:${synchroOk ? '#ecfdf5' : '#fef2f2'};color:${synchroOk ? '#065f46' : '#991b1b'}">
    <b>${synchroOk ? '✅ Synchronisé' : '⚠️ Écart de synchronisation'}</b>
    <span>${ctl.lignesSheet != null ? `${ctl.lignesSheet} réponse(s) dans le Sheet (onglet « ${vxEsc(ctl.onglet || '')} ») · ${ctl.affiches} ticket(s) dans l'application` : ''}</span>
    ${ctl.dernier ? `<span>· dernier reçu : ${vxEsc(ctl.dernier)}</span>` : ''}${ctl.lu ? `<span>· lu à ${vxEsc(String(ctl.lu).slice(11))}</span>` : ''}
    ${r.nouveaux && r.nouveaux.length ? `<span class="mt-new">${r.nouveaux.length} nouveau(x) depuis votre dernière visite</span>` : ''}
    ${ctl.colonnesReperees && (ctl.colonnesReperees.date < 0 || ctl.colonnesReperees.description < 0) ? `<span>· ⚠️ colonne ${ctl.colonnesReperees.date < 0 ? '« horodatage »' : '« description »'} non reconnue dans le Sheet</span>` : ''}
  </div>
  <div class="mt-kpi">
    <div onclick="maintFiltre.statut='ouverts';renderMaintenance()" style="background:linear-gradient(135deg,#1A4E8A,#3b82f6)"><b>${ouverts.length}</b><span>tickets ouverts</span></div>
    <div onclick="maintFiltre.statut='ouverts';maintFiltre.texte='';renderMaintenance()" style="background:linear-gradient(135deg,#dc2626,#f97316)"><b>${urg.length}</b><span>urgents ouverts</span></div>
    <div onclick="maintFiltre.statut='Nouveau';renderMaintenance()" style="background:linear-gradient(135deg,#d97706,#fbbf24)"><b>${A.filter(t => t.statut === 'Nouveau').length}</b><span>non pris en charge</span></div>
    <div onclick="maintFiltre.statut='Clôturé';renderMaintenance()" style="background:linear-gradient(135deg,#059669,#34d399)"><b>${A.filter(t => ['Résolu', 'Clôturé'].includes(t.statut)).length}</b><span>résolus / clôturés</span></div>
    <div onclick="maintFiltre.statut='archives';renderMaintenance()" style="background:linear-gradient(135deg,#475569,#94a3b8)"><b>${T.filter(t => t.archive).length}</b><span>archivés · ${T.length} reçus</span></div>
  </div>
  ${Object.keys(cats).length ? `<div class="flex flex-wrap gap-2 mb-3">${Object.entries(cats).sort((a, b) => b[1] - a[1]).map(([c, n]) => `<span class="mt-chip">${vxEsc(c)} · ${n}</span>`).join('')}</div>` : ''}
  <div class="flex flex-wrap gap-2 mb-4 items-center">
    <select onchange="maintFiltre.statut=this.value;maintSel.clear();renderMaintenance()" class="border border-gray-200 rounded-xl px-3 py-2 text-sm">
      ${[['ouverts', 'Tickets ouverts'], ['tous', 'Tous (hors archives)']].concat(MAINT_STATUTS.map(x => [x, x]), [['archives', '🗄 Archivés']]).map(o => `<option value="${o[0]}" ${maintFiltre.statut === o[0] ? 'selected' : ''}>${o[1]}</option>`).join('')}</select>
    <input id="mtRecherche" type="search" placeholder="🔎 N° de ticket, lieu, catégorie, description, responsable…" value="${vxEsc(maintFiltre.texte)}" oninput="maintFiltre.texte=this.value;renderMaintenance();const e=document.getElementById('mtRecherche');e.focus();e.setSelectionRange(e.value.length,e.value.length)" class="border border-gray-200 rounded-xl px-3 py-2 text-sm flex-1 min-w-[220px]">
    <div class="inline-flex rounded-xl border border-gray-200 overflow-hidden text-sm font-bold">
      ${[['liste', '☰ Liste'], ['cartes', '▦ Cartes']].map(([v, l]) => `<button onclick="maintVue='${v}';try{localStorage.setItem('maintVue','${v}')}catch(e){};renderMaintenance()" class="px-3 py-2 ${maintVue === v ? 'bg-[#1A4E8A] text-white' : 'bg-white text-gray-600'}">${l}</button>`).join('')}</div>
  </div>
  ${nbSel ? `<div class="mt-bulk"><b class="text-sm mr-2">${nbSel} sélectionné(s)</b>
    <button onclick="maintActionGroupee('cloturer')">✓ Clôturer</button><button onclick="maintActionGroupee('relancer')">🔔 Relancer</button>
    <button onclick="maintActionGroupee('annuler')">✕ Annuler</button>
    ${maintFiltre.statut === 'archives' ? `<button onclick="maintActionGroupee('desarchiver')">↩ Désarchiver</button>` : `<button onclick="maintActionGroupee('archiver')">🗄 Archiver</button>`}
    <button onclick="maintSel.clear();renderMaintenance()" style="margin-left:auto">Désélectionner</button></div>` : ''}
  ${liste.length ? (maintVue === 'liste' ? maintTableau(liste, r.nouveaux || []) : maintCartes(liste, r.nouveaux || [])) : '<p class="text-gray-400 text-sm">Aucun ticket pour ce filtre ✅</p>'}`;
}
function maintBoutons(t) {
  const c = vxEsc(t.cle).replace(/'/g, "\\'"), o = maintOuvert(t);
  if (t._enCours) return `<span class="text-xs font-bold text-[#1A4E8A] whitespace-nowrap">⏳ ${vxEsc(t._enCours)}…</span>`;
  return `<button class="mt-act" title="Modifier" onclick="event.stopPropagation();ouvrirTicket('${c}')">✏️</button>
    ${o ? `<button class="mt-act" title="Clôturer" onclick="event.stopPropagation();maintAction('${c}','cloturer')">✓</button>` : `<button class="mt-act" title="Rouvrir" onclick="event.stopPropagation();maintAction('${c}','rouvrir')">↺</button>`}
    <button class="mt-act" title="Relancer le responsable par mail" onclick="event.stopPropagation();maintAction('${c}','relancer')">🔔</button>
    <button class="mt-act" title="Annuler" ${t.statut === 'Annulé' ? 'disabled' : ''} onclick="event.stopPropagation();maintAction('${c}','annuler')">✕</button>
    ${t.archive ? `<button class="mt-act" title="Désarchiver" onclick="event.stopPropagation();maintAction('${c}','desarchiver')">↩</button>` : `<button class="mt-act" title="Archiver" onclick="event.stopPropagation();maintAction('${c}','archiver')">🗄</button>`}`;
}
function maintNum(t) { return t.numero ? `<span class="font-black text-[#1A4E8A] whitespace-nowrap" style="font-variant-numeric:tabular-nums">${vxEsc(t.numero)}</span>` : '<span class="text-gray-300">—</span>'; }
function maintPill(t) { const [bg, fg] = MAINT_COUL[t.statut] || MAINT_COUL['Nouveau']; return `<span class="mt-pill" style="background:${bg};color:${fg}">${vxEsc(t.statut)}</span>`; }
function maintTableau(liste, nouveaux) {
  const tous = liste.every(t => maintSel.has(t.cle));
  return `<div class="bg-white rounded-2xl border border-gray-100 overflow-auto" style="max-height:70vh"><table class="mt-tab"><thead><tr>
    <th><input type="checkbox" ${tous ? 'checked' : ''} onchange="maintToutSel(this.checked)"></th><th>N°</th><th>Reçu le</th><th>Statut</th><th>Priorité</th><th>Signalement</th><th>Lieu</th><th>Demandeur</th><th>Responsable</th><th>Suivi</th><th class="text-right">Actions</th></tr></thead><tbody>
    ${liste.map(t => { const c = vxEsc(t.cle).replace(/'/g, "\\'"), u = maintUrgent(t); return `<tr class="${maintSel.has(t.cle) ? 'sel' : ''}" style="cursor:pointer;${t._enCours ? 'opacity:.55' : ''}" onclick="ouvrirTicket('${c}')">
      <td onclick="event.stopPropagation()"><input type="checkbox" ${maintSel.has(t.cle) ? 'checked' : ''} onchange="maintBasculer('${c}',this.checked)"></td>
      <td>${maintNum(t)}</td>
      <td class="whitespace-nowrap text-gray-500 text-xs">${vxEsc(t.date)}${nouveaux.includes(t.cle) ? '<span class="mt-new">nouveau</span>' : ''}</td>
      <td>${maintPill(t)}${t.archive ? ' <span class="mt-chip">🗄</span>' : ''}</td>
      <td class="text-xs ${u ? 'text-red-700 font-black' : 'text-gray-500'}">${u ? '🔥 ' : ''}${vxEsc(t.priorite || t.urgence || '—')}</td>
      <td style="min-width:240px;max-width:380px"><p class="font-bold text-gray-800" style="display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden">${vxEsc(t.description || t.categorie || 'Signalement')}</p>${t.categorie ? `<span class="mt-chip">🏷 ${vxEsc(t.categorie)}</span>` : ''}${t.corrige ? ' <span class="mt-chip" title="Contenu corrigé dans le CRM">✎ corrigé</span>' : ''}</td>
      <td class="text-xs text-gray-600">${vxEsc([t.campus, t.lieu].filter(Boolean).join(' · ') || '—')}</td>
      <td class="text-xs text-gray-600">${vxEsc(t.nom || t.email || '—')}</td>
      <td class="text-xs">${t.responsable ? `👤 ${vxEsc(t.responsable)}` : '<span class="text-gray-300">—</span>'}${t.relances ? `<div class="text-[10px] text-orange-600">🔔 ${t.relances} relance(s)${t.derniereRelance ? ' · ' + vxEsc(t.derniereRelance.slice(0, 10)) : ''}</div>` : ''}</td>
      <td class="text-xs text-gray-500" style="max-width:220px"><span style="display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden">${vxEsc(t.commentaire || '')}</span>${t.maj ? `<div class="text-[10px] text-gray-400">${vxEsc(t.maj)}</div>` : ''}</td>
      <td class="whitespace-nowrap text-right">${maintBoutons(t)}</td></tr>`; }).join('')}
  </tbody></table></div>`;
}
function maintCartes(liste, nouveaux) {
  return `<div class="mt-grid">${liste.map(t => {
    const fg = (MAINT_COUL[t.statut] || MAINT_COUL['Nouveau'])[1], u = maintUrgent(t), c = vxEsc(t.cle).replace(/'/g, "\\'");
    return `<div class="mt-card" style="--bord:${u ? '#dc2626' : fg}" onclick="ouvrirTicket('${c}')">
      <div class="flex justify-between gap-2"><span>${maintNum(t)} ${maintPill(t)}${nouveaux.includes(t.cle) ? '<span class="mt-new">nouveau</span>' : ''}</span>
        ${t.urgence || t.priorite ? `<span class="mt-pill" style="background:${u ? '#fee2e2' : '#f8fafc'};color:${u ? '#b91c1c' : '#64748b'}">${u ? '🔥 ' : ''}${vxEsc(t.priorite || t.urgence)}</span>` : ''}</div>
      <p class="font-bold text-gray-800 text-sm" style="display:-webkit-box;-webkit-line-clamp:3;-webkit-box-orient:vertical;overflow:hidden">${vxEsc(t.description || t.categorie || 'Signalement')}</p>
      <div class="flex flex-wrap gap-1">${t.categorie ? `<span class="mt-chip">🏷 ${vxEsc(t.categorie)}</span>` : ''}${t.campus ? `<span class="mt-chip">🏫 ${vxEsc(t.campus)}</span>` : ''}${t.lieu ? `<span class="mt-chip">📍 ${vxEsc(t.lieu)}</span>` : ''}</div>
      <p class="text-[11px] text-gray-400 mt-auto">${vxEsc(t.date)}${t.nom ? ' · ' + vxEsc(t.nom) : t.email ? ' · ' + vxEsc(t.email) : ''}${t.responsable ? ' · 👤 ' + vxEsc(t.responsable) : ''}${t.relances ? ' · 🔔 ' + t.relances : ''}</p>
      <div class="flex gap-1 justify-end">${maintBoutons(t)}</div></div>`;
  }).join('')}</div>`;
}
function maintBasculer(cle, on) { on ? maintSel.add(cle) : maintSel.delete(cle); renderMaintenance(); }
function maintToutSel(on) { maintListe().forEach(t => on ? maintSel.add(t.cle) : maintSel.delete(t.cle)); renderMaintenance(); }

const MAINT_LIB = { cloturer: 'clôturer', annuler: 'annuler', archiver: 'archiver', desarchiver: 'désarchiver', rouvrir: 'rouvrir', relancer: 'relancer' };
function maintResume(t) { return (t.description || t.categorie || 'Signalement').slice(0, 80); }
const maintNom = t => (t.numero ? t.numero + ' — ' : '') + maintResume(t).slice(0, 50);
function maintAppliquer(res) {
  (res || []).forEach(x => { const t = maintData.tickets.find(y => y.cle === x.cle); if (!t) return;
    Object.assign(t, { statut: x.statut, archive: x.archive, journal: x.journal, relances: x.relances, derniereRelance: x.derniereRelance, resolu: x.resolu, maj: x.maj, par: x.par, responsable: x.responsable, responsableMail: x.responsableMail, ancienneCle: '' });
    if (x.corr) { ['campus', 'lieu', 'categorie', 'description'].forEach(k => { t[k] = x.corr[k] != null ? x.corr[k] : t.original[k]; }); t.corrige = Object.keys(x.corr).length > 0; } });
}
async function maintEnvoyer(cles, op, extra = {}) {
  const T = maintData.tickets.filter(t => cles.includes(t.cle)), resumes = {}, anciennesCles = {};
  T.forEach(t => { resumes[t.cle] = maintResume(t); if (t.ancienneCle) anciennesCles[t.cle] = t.ancienneCle; });
  let r;
  try { r = await api('updateTicketMaintenance', Object.assign({ op, cles, resumes, anciennesCles }, extra)); }
  catch (e) { r = { success: false, message: 'Pas de réponse du serveur (' + (e && e.message || e) + '). Réessayez.' }; }
  if (!r || typeof r !== 'object') r = { success: false, message: 'Réponse du serveur illisible' };
  if (r.resultats) maintAppliquer(r.resultats);
  return r;
}
const MAINT_EN_COURS = { cloturer: 'Clôture', annuler: 'Annulation', archiver: 'Archivage', desarchiver: 'Désarchivage', rouvrir: 'Réouverture', relancer: 'Relance', maj: 'Enregistrement' };
function maintOptimiste(t, op) {
  if (op === 'cloturer') t.statut = 'Clôturé'; else if (op === 'annuler') t.statut = 'Annulé'; else if (op === 'rouvrir') { t.statut = 'Pris en charge'; t.archive = false; }
  else if (op === 'archiver') t.archive = true; else if (op === 'desarchiver') t.archive = false;
}
function maintRafraichirModal(cle) { const m = document.getElementById('mtModal'); if (m && !m.classList.contains('hidden')) ouvrirTicket(cle); }
async function maintAction(cle, op) {
  const t = maintData.tickets.find(x => x.cle === cle); if (!t || t._enCours) return;
  let extra = {};
  if (op === 'relancer') {
    if (!t.responsable) { toast("Désignez d'abord un responsable (✏️ Modifier)", false); ouvrirTicket(cle); return; }
    const msg = prompt(`Relancer ${t.responsable} par mail (copie direction) pour le ticket ${t.numero || ''} ?\nMessage complémentaire (facultatif) :`, '');
    if (msg === null) return; extra.message = msg;
  } else if (op === 'annuler' && !confirm(`Annuler le ticket ${maintNom(t)} ?`)) return;
  const avant = { statut: t.statut, archive: t.archive };
  maintOptimiste(t, op); t._enCours = MAINT_EN_COURS[op] || 'En cours';
  renderMaintenance(); maintRafraichirModal(cle);
  const r = await maintEnvoyer([cle], op, extra);
  delete t._enCours;
  if (!r.success || (r.erreurs && r.erreurs.length) || !(r.resultats || []).length) {
    Object.assign(t, avant);
    toast(`${t.numero || 'Ticket'} : ${r.message || (r.erreurs || []).join(' ; ') || 'action non enregistrée'}`, false);
  } else toast(op === 'relancer' ? `${t.numero} : relance envoyée à ${t.responsable} ✓` : `${t.numero || 'Ticket'} ${op === 'rouvrir' ? 'rouvert' : MAINT_LIB[op].replace(/er$/, 'é')} ✓`);
  renderMaintenance(); maintRafraichirModal(cle);
}
async function maintActionGroupee(op) {
  const visibles = new Set(maintListe().map(t => t.cle)), cles = [...maintSel].filter(c => visibles.has(c));
  if (!cles.length) return;
  let extra = {};
  if (op === 'relancer') {
    const sans = maintData.tickets.filter(t => cles.includes(t.cle) && !t.responsable).length;
    const msg = prompt(`Relancer par mail les responsables de ${cles.length} ticket(s)${sans ? ` (${sans} sans responsable seront ignorés)` : ''} ?\nMessage complémentaire (facultatif) :`, '');
    if (msg === null) return; extra.message = msg;
  } else if (!confirm(`${MAINT_LIB[op][0].toUpperCase() + MAINT_LIB[op].slice(1)} ${cles.length} ticket(s) ?`)) return;
  const T = maintData.tickets.filter(t => cles.includes(t.cle)), avant = T.map(t => ({ t, statut: t.statut, archive: t.archive }));
  T.forEach(t => { if (op !== 'relancer') maintOptimiste(t, op); t._enCours = MAINT_EN_COURS[op]; });
  maintSel.clear(); renderMaintenance();
  const r = await maintEnvoyer(cles, op, extra);
  const faits = new Set((r.resultats || []).map(x => x.cle));
  avant.forEach(a => { delete a.t._enCours; if (!faits.has(a.t.cle)) Object.assign(a.t, { statut: a.statut, archive: a.archive }); });
  const ok = faits.size;
  toast(r.success === false && !ok ? 'Erreur : ' + r.message : `${ok} ticket(s) traité(s)${op === 'relancer' ? ` · ${r.relances || 0} relance(s) envoyée(s)` : ''}${r.erreurs && r.erreurs.length ? ` · ${r.erreurs.length} non traité(s) : ${r.erreurs.slice(0, 2).join(' ; ')}` : ''}`, !(r.erreurs && r.erreurs.length) && r.success !== false);
  renderMaintenance();
}

// Responsable : membre du personnel ou contact externe (création à la volée)
function maintOptionsResp(t) {
  const eq = (maintData.equipe || []).slice().sort((a, b) => a.nom.localeCompare(b.nom)), ct = (maintData.contacts || []).slice().sort((a, b) => a.nom.localeCompare(b.nom));
  const connu = !t.responsable || eq.some(x => x.nom === t.responsable) || ct.some(x => x.nom === t.responsable);
  const opt = (x, g) => `<option value="${vxEsc(g + '|' + x.nom)}" data-mail="${vxEsc(x.mail || '')}" ${x.nom === t.responsable ? 'selected' : ''}>${vxEsc(x.nom)}${x.societe ? ' — ' + vxEsc(x.societe) : ''}${x.mail ? '' : ' (sans mail)'}</option>`;
  return `<option value="">— Personne —</option>
    ${!connu ? `<option value="${vxEsc('libre|' + t.responsable)}" data-mail="${vxEsc(t.responsableMail || '')}" selected>${vxEsc(t.responsable)} (saisie libre)</option>` : ''}
    <optgroup label="Équipe HECG">${eq.map(x => opt(x, 'eq')).join('')}</optgroup>
    <optgroup label="Contacts externes">${ct.map(x => opt(x, 'ct')).join('')}</optgroup>
    <option value="__nouveau">➕ Créer un autre contact…</option>`;
}
function maintRespChange(sel) {
  const box = document.getElementById('mtNouveauContact');
  box.classList.toggle('hidden', sel.value !== '__nouveau');
  if (sel.value === '__nouveau') document.getElementById('mtNcNom').focus();
}
async function maintCreerContact() {
  const d = { op: 'contact', nom: document.getElementById('mtNcNom').value.trim(), mail: document.getElementById('mtNcMail').value.trim(), tel: document.getElementById('mtNcTel').value.trim(), societe: document.getElementById('mtNcSoc').value.trim() };
  if (!d.nom) { toast('Nom obligatoire', false); return; }
  if (!d.mail && !confirm("Sans adresse mail, ce contact ne pourra pas être relancé par mail. Continuer ?")) return;
  let r; try { r = await api('updateTicketMaintenance', d); } catch (e) { r = { success: false, message: 'Pas de réponse du serveur' }; }
  if (!r.success) { toast(r.message, false); return; }
  maintData.contacts = (maintData.contacts || []).concat([r.contact]);
  const sel = document.getElementById('mtResp'), o = document.createElement('option');
  o.value = 'ct|' + r.contact.nom; o.dataset.mail = r.contact.mail; o.textContent = r.contact.nom + (r.contact.societe ? ' — ' + r.contact.societe : '');
  sel.querySelector('optgroup[label="Contacts externes"]').appendChild(o); sel.value = o.value;
  document.getElementById('mtNouveauContact').classList.add('hidden');
  toast('Contact créé ✓');
}
function ouvrirTicket(cle) {
  const t = maintData.tickets.find(x => x.cle === cle); if (!t) return;
  const c = vxEsc(t.cle).replace(/'/g, "\\'");
  let m = document.getElementById('mtModal');
  if (!m) { m = document.createElement('div'); m.id = 'mtModal'; m.className = 'modal-overlay hidden'; m.onclick = e => { if (e.target === m) m.classList.add('hidden'); }; document.body.appendChild(m); }
  const champ = (id, lab, k, rows) => `<div class="${rows ? 'col-span-2' : ''}"><label>${lab}${t.original && t.original[k] !== t[k] ? ` <span class="text-[10px] text-orange-600" title="${vxEsc(t.original[k])}">(corrigé — original au survol)</span>` : ''}</label>
    ${rows ? `<textarea id="${id}" rows="${rows}">${vxEsc(t[k])}</textarea>` : `<input type="text" id="${id}" value="${vxEsc(t[k])}">`}</div>`;
  m.innerHTML = `<div class="modal-box p-6" style="max-width:760px">
    <div class="flex justify-between items-start gap-3 mb-3"><div><h3 class="text-xl font-black text-[#1A4E8A]">🛠️ Ticket ${vxEsc(t.numero || '')} <span class="text-sm font-bold text-gray-400">reçu le ${vxEsc(t.date)}</span></h3>
      <p class="text-xs text-gray-500">${vxEsc(t.nom || '')}${t.email ? ' · ' + vxEsc(t.email) : ''} ${maintPill(t)}${t.archive ? ' <span class="mt-chip">🗄 archivé</span>' : ''}</p></div>
      <div class="flex gap-1 flex-wrap justify-end">${maintBoutons(t)}</div></div>
    <details class="bg-gray-50 rounded-2xl p-3 mb-4 text-sm"><summary class="font-bold text-gray-600 cursor-pointer">Réponse d'origine du formulaire</summary>
      <div class="space-y-1 mt-2" style="max-height:30vh;overflow:auto">${t.champs.map(([k, v]) => `<p><span class="font-bold text-gray-500">${vxEsc(k)} :</span> ${/^https?:\/\//.test(v) ? `<a href="${vxEsc(v)}" target="_blank" class="text-blue-600 underline">ouvrir</a>` : vxEsc(v)}</p>`).join('')}</div></details>
    <div class="grid grid-cols-2 gap-3">
      ${champ('mtDesc', 'Description', 'description', 3)}
      ${champ('mtCat', 'Catégorie', 'categorie')}${champ('mtCampus', 'Campus', 'campus')}
      ${champ('mtLieu', 'Lieu / salle', 'lieu')}
      <div><label>Statut</label><select id="mtStatut">${MAINT_STATUTS.map(s => `<option ${s === t.statut ? 'selected' : ''}>${s}</option>`).join('')}</select></div>
      <div><label>Priorité</label><select id="mtPrio">${['', 'Basse', 'Normale', 'Haute', 'Urgente'].map(s => `<option ${s === t.priorite ? 'selected' : ''}>${s}</option>`).join('')}</select></div>
      <div><label>Qui s'en occupe ?</label><select id="mtResp" onchange="maintRespChange(this)">${maintOptionsResp(t)}</select></div>
      <div id="mtNouveauContact" class="col-span-2 hidden rounded-2xl border border-dashed border-[#1A4E8A] p-3 bg-blue-50/40">
        <p class="text-xs font-black text-[#1A4E8A] mb-2">Nouveau contact (prestataire, propriétaire, technicien…)</p>
        <div class="grid grid-cols-2 gap-2"><input id="mtNcNom" type="text" placeholder="Nom *"><input id="mtNcSoc" type="text" placeholder="Société / rôle">
          <input id="mtNcMail" type="email" placeholder="Mail (pour les relances)"><input id="mtNcTel" type="tel" placeholder="Téléphone"></div>
        <div class="flex justify-end gap-2 mt-2"><button onclick="document.getElementById('mtResp').value='';maintRespChange(document.getElementById('mtResp'))" class="btn-ghost text-xs">Annuler</button><button onclick="maintCreerContact()" class="btn-orange text-xs">Créer le contact</button></div></div>
      <div class="col-span-2"><label>Suivi / commentaire</label><textarea id="mtCom" rows="3" placeholder="Intervention prévue, pièce commandée, résolu le…">${vxEsc(t.commentaire)}</textarea></div>
    </div>
    ${t.journal && t.journal.length ? `<div class="mt-4"><p class="text-xs font-black text-gray-500 uppercase mb-1">Historique</p><div class="text-xs text-gray-600 space-y-1" style="max-height:20vh;overflow:auto">${t.journal.slice().reverse().map(j => `<p><span class="text-gray-400">${vxEsc(j.d)}</span> · <b>${vxEsc(j.par)}</b> — ${vxEsc(j.evt)}</p>`).join('')}</div></div>` : ''}
    ${t.maj ? `<p class="text-xs text-gray-400 mt-2">Dernière mise à jour : ${vxEsc(t.maj)} par ${vxEsc(t.par)}${t.resolu ? ' · clôturé le ' + vxEsc(t.resolu) : ''}${t.relances ? ' · ' + t.relances + ' relance(s)' : ''}</p>` : ''}
    <div class="flex justify-end gap-3 mt-5"><button onclick="document.getElementById('mtModal').classList.add('hidden')" class="btn-ghost">Fermer</button>
      <button id="mtSave" onclick="saveTicket('${c}')" class="btn-orange">Enregistrer</button></div></div>`;
  m.classList.remove('hidden');
}
async function saveTicket(cle) {
  const t = maintData.tickets.find(x => x.cle === cle), sel = document.getElementById('mtResp');
  if (sel.value === '__nouveau') { toast("Créez le contact ou choisissez une personne", false); return; }
  const opt = sel.selectedOptions[0], respNom = sel.value ? sel.value.split('|').slice(1).join('|') : '';
  const corrections = { description: document.getElementById('mtDesc').value.trim(), categorie: document.getElementById('mtCat').value.trim(), campus: document.getElementById('mtCampus').value.trim(), lieu: document.getElementById('mtLieu').value.trim() };
  const d = { statut: document.getElementById('mtStatut').value, priorite: document.getElementById('mtPrio').value, responsable: respNom, responsableMail: sel.value && opt ? (opt.dataset.mail || '') : '',
    commentaire: document.getElementById('mtCom').value.trim(), corrections, originaux: t.original };
  const btn = document.getElementById('mtSave');
  if (btn) { if (btn.disabled) return; btn.disabled = true; btn.textContent = '⏳ Enregistrement…'; }
  const r = await maintEnvoyer([cle], 'maj', d);
  if (!r.success || !(r.resultats || []).length) { if (btn) { btn.disabled = false; btn.textContent = 'Enregistrer'; } toast('Erreur : ' + (r.message || 'non enregistré'), false); return; }
  Object.assign(t, { priorite: d.priorite, commentaire: d.commentaire });
  document.getElementById('mtModal').classList.add('hidden'); toast(`${t.numero || 'Ticket'} mis à jour ✓`); renderMaintenance();
}

// ══ CENTRE DES FORMULAIRES ══
let formsData = [], formsPageUrl = '', fEd = null;
const F_TYPES = [
  ['texte', 'Réponse courte', '✏️'], ['paragraphe', 'Paragraphe', '📝'], ['choix', 'Choix unique', '🔘'], ['cases', 'Cases à cocher', '☑️'], ['liste', 'Liste déroulante', '🔽'],
  ['etoiles', 'Note en étoiles', '⭐'], ['echelle', 'Échelle', '📏'], ['nps', 'Recommandation 0-10', '📣'], ['oui_non', 'Oui / Non', '👍'], ['date', 'Date', '📅'],
  ['email', 'E-mail', '📧'], ['tel', 'Téléphone', '📞'], ['nombre', 'Nombre', '🔢'], ['section', 'Titre de section', '🔖']];
const F_LIB = Object.fromEntries(F_TYPES.map(t => [t[0], t]));
async function loadFormulaires() {
  const s = document.getElementById('section-formulaires');
  s.innerHTML = `<div class="text-center text-gray-400 py-16">Chargement des formulaires…</div>`;
  const r = await api('getFormulaires');
  if (!r.success) { s.innerHTML = `<p class="text-red-600 font-bold">${vxEsc(r.message)}</p>`; return; }
  formsData = r.data; formsPageUrl = r.pageUrl; renderFormulaires();
}
function renderFormulaires() {
  const s = document.getElementById('section-formulaires');
  const stat = { publié: ['#dcfce7', '#15803d', '● En ligne'], brouillon: ['#f3f4f6', '#6b7280', '○ Brouillon'], fermé: ['#fee2e2', '#b91c1c', '■ Clôturé'] };
  s.innerHTML = `<style>
    .fm-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(290px,1fr));gap:1rem;margin-top:1.25rem}
    .fm-card{background:#fff;border:1px solid #eef2f7;border-radius:1.25rem;overflow:hidden;display:flex;flex-direction:column;transition:transform .15s,box-shadow .15s}
    .fm-card:hover{transform:translateY(-3px);box-shadow:0 20px 40px -24px rgba(26,78,138,.55)}
    .fm-band{height:70px;position:relative}.fm-band span{position:absolute;right:12px;top:10px;font-size:.65rem;font-weight:800;padding:.2rem .55rem;border-radius:999px;background:rgba(255,255,255,.9)}
    .fm-body{padding:1rem 1.1rem;display:flex;flex-direction:column;gap:.45rem;flex:1}
    .fm-btns{display:flex;flex-wrap:wrap;gap:.35rem;padding:0 1.1rem 1rem}
    .fm-btns button,.fm-btns a{font-size:.72rem;font-weight:700;padding:.4rem .65rem;border-radius:.6rem;border:0;cursor:pointer;background:#f3f4f6;color:#374151;text-decoration:none}
    .fm-new{border:2px dashed #cbd5e1;border-radius:1.25rem;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:.4rem;min-height:210px;cursor:pointer;color:#64748b;background:linear-gradient(180deg,#f8fafc,#fff);font-weight:700}
    .fm-new:hover{border-color:#1A4E8A;color:#1A4E8A}
  </style>
  <div class="flex flex-wrap justify-between items-start gap-3">
    <div><h1 class="text-2xl font-black text-[#1A4E8A]">🧾 Centre des formulaires</h1>
    <p class="text-sm text-gray-500 mt-1">Tous vos formulaires au même endroit : créés ici (mise en page soignée, réponses dans le CRM) ou formulaires Google existants référencés.</p></div>
    <div class="flex gap-2"><button onclick="fmExterne()" class="btn-ghost text-sm">🔗 Référencer un formulaire Google</button><button onclick="fmEditer()" class="btn-orange text-sm">+ Nouveau formulaire</button></div>
  </div>
  <div class="fm-grid">
    <div class="fm-new" onclick="fmEditer()"><span style="font-size:2rem">✨</span>Créer un formulaire<span class="text-xs font-medium text-gray-400">questionnaire, enquête, inscription…</span></div>
    ${formsData.map(f => {
      const th = f.theme || {}, st = stat[f.statut] || stat.brouillon;
      return `<div class="fm-card"><div class="fm-band" style="background:linear-gradient(120deg,${th.couleur || '#1A4E8A'},${th.accent || '#F26522'})"><span style="color:${st[1]}">${st[2]}</span></div>
        <div class="fm-body"><p class="font-black text-gray-800">${vxEsc(f.titre)}</p>
          <p class="text-xs text-gray-500" style="display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden">${vxEsc(f.description)}</p>
          <div class="flex flex-wrap gap-1 mt-auto">
            <span class="mt-chip" style="background:#f1f5f9;font-size:.66rem;font-weight:700;padding:.15rem .5rem;border-radius:.5rem">${f.type === 'externe' ? 'Google Forms' : 'Formulaire HECG'}</span>
            ${f.indicateur ? `<span style="background:#eef2ff;color:#3730a3;font-size:.66rem;font-weight:800;padding:.15rem .5rem;border-radius:.5rem">Qualiopi ${vxEsc(f.indicateur)}</span>` : ''}
            <span style="background:#ecfdf5;color:#047857;font-size:.66rem;font-weight:800;padding:.15rem .5rem;border-radius:.5rem">${f.nbReponses == null ? (f.accesRefuse ? '🔒 accès à partager' : '—') : f.nbReponses + ' réponse(s)'}</span>
          </div></div>
        <div class="fm-btns">
          <button onclick="fmReponses('${f.id}')">📊 Réponses</button>
          ${f.type === 'interne' ? `<button onclick="fmEditer('${f.id}')">✏️ Modifier</button>` : `<button onclick="fmExterne('${f.id}')">✏️ Modifier</button>`}
          ${f.lienPublic ? `<a href="${vxEsc(f.lienPublic)}" target="_blank">👁 Ouvrir</a><button onclick="navigator.clipboard.writeText('${vxEsc(f.lienPublic)}');toast('Lien copié ✓')">📋 Lien</button><button onclick="fmQr('${vxEsc(f.lienPublic)}','${vxEsc(f.titre).replace(/'/g, '')}')">▦ QR</button>` : ''}
          <button onclick="fmArchiver('${f.id}')" style="background:#fef2f2;color:#dc2626">🗑</button>
        </div></div>`;
    }).join('')}
  </div>`;
}
function fmQr(url, titre) {
  const src = 'https://api.qrserver.com/v1/create-qr-code/?size=420x420&margin=16&data=' + encodeURIComponent(url);
  const w = window.open('', '_blank');
  w.document.write(`<title>QR — ${titre}</title><body style="font-family:Arial;text-align:center;padding:40px"><h2>${titre}</h2><img src="${src}" style="width:340px"><p style="color:#666;word-break:break-all">${url}</p><button onclick="print()">Imprimer</button></body>`);
}
async function fmArchiver(id) {
  const f = formsData.find(x => x.id === id);
  if (!confirm('Archiver le formulaire « ' + f.titre + ' » ? Il disparaît de la liste ; ses réponses sont conservées.')) return;
  const r = await api('deleteFormulaire', { id }); if (!r.success) { toast(r.message, false); return; }
  toast('Formulaire archivé'); loadFormulaires();
}
function fmExterne(id) {
  const f = id ? formsData.find(x => x.id === id) : null;
  const titre = prompt('Titre du formulaire Google :', f ? f.titre : ''); if (!titre) return;
  const url = prompt('Lien de réponse du formulaire (https://docs.google.com/forms/…/viewform) :', f ? f.urlExterne : ''); if (url === null) return;
  const sheet = prompt('Lien du Google Sheet des réponses (pour afficher les réponses ici) :', f && f.sheetExterne ? 'https://docs.google.com/spreadsheets/d/' + f.sheetExterne : ''); if (sheet === null) return;
  const ind = prompt('Indicateur Qualiopi concerné (facultatif, ex. 30) :', f ? f.indicateur : '') || '';
  api('saveFormulaire', { formulaire: JSON.stringify({ id: f ? f.id : '', titre, description: f ? f.description : '', type: 'externe', statut: 'publié', urlExterne: url.trim(), sheetExterne: sheet.trim(), indicateur: ind, theme: f ? f.theme : {} }) })
    .then(r => { if (!r.success) toast(r.message, false); else { toast('Formulaire référencé ✓'); loadFormulaires(); } });
}

// ── Éditeur de formulaire ──
function fmEditer(id) {
  const f = id ? JSON.parse(JSON.stringify(formsData.find(x => x.id === id))) : null;
  fEd = f || { id: '', titre: '', description: '', statut: 'brouillon', type: 'interne', questions: [{ id: 'q1', type: 'texte', label: 'Votre nom', obligatoire: true }], theme: { couleur: '#1A4E8A', accent: '#F26522', fond: 'degrade', police: 'moderne' }, indicateur: '', messageFin: '' };
  fEd.theme = Object.assign({ couleur: '#1A4E8A', accent: '#F26522', fond: 'degrade', police: 'moderne' }, fEd.theme || {});
  let m = document.getElementById('fmModal');
  if (!m) { m = document.createElement('div'); m.id = 'fmModal'; m.className = 'modal-overlay hidden'; document.body.appendChild(m); }
  m.classList.remove('hidden'); fmRendreEditeur();
}
const FM_PALETTES = [['#1A4E8A', '#F26522'], ['#0F766E', '#F59E0B'], ['#5B3F86', '#EC4899'], ['#111827', '#22D3EE'], ['#9D174D', '#F97316'], ['#065F46', '#84CC16'], ['#1E3A8A', '#A855F7']];
function fmRendreEditeur() {
  const m = document.getElementById('fmModal'), f = fEd, th = f.theme;
  m.innerHTML = `<div class="modal-box" style="max-width:1100px;width:96vw;height:92vh;display:flex;flex-direction:column;padding:0;overflow:hidden">
    <div style="padding:1rem 1.4rem;border-bottom:1px solid #eef2f7;display:flex;justify-content:space-between;align-items:center;gap:1rem;background:linear-gradient(120deg,${th.couleur},${th.accent});color:#fff">
      <div><p style="font-size:.7rem;font-weight:800;opacity:.85;text-transform:uppercase">Éditeur de formulaire</p><p style="font-weight:900;font-size:1.1rem">${vxEsc(f.titre || 'Nouveau formulaire')}</p></div>
      <div style="display:flex;gap:.5rem;align-items:center">
        <select id="fmStatut" style="border-radius:.6rem;padding:.4rem .6rem;color:#111;font-size:.8rem">${[['brouillon', '○ Brouillon'], ['publié', '● Publié (en ligne)'], ['fermé', '■ Clôturé']].map(o => `<option value="${o[0]}" ${f.statut === o[0] ? 'selected' : ''}>${o[1]}</option>`).join('')}</select>
        <button onclick="fmSauver()" style="background:#fff;color:${th.couleur};font-weight:900;padding:.5rem 1rem;border-radius:.7rem;border:0;cursor:pointer">💾 Enregistrer</button>
        <button onclick="document.getElementById('fmModal').classList.add('hidden')" style="background:rgba(255,255,255,.2);color:#fff;font-weight:900;padding:.5rem .8rem;border-radius:.7rem;border:0;cursor:pointer">✕</button></div></div>
    <div style="display:grid;grid-template-columns:300px 1fr;flex:1;min-height:0">
      <div style="border-right:1px solid #eef2f7;padding:1rem;overflow:auto;background:#fafbfd">
        <p class="text-xs font-black text-gray-400 uppercase mb-2">Paramètres</p>
        <label>Titre</label><input id="fmTitre" value="${vxEsc(f.titre)}" oninput="fEd.titre=this.value">
        <label class="mt-2 block">Description / consignes</label><textarea id="fmDesc" rows="3" oninput="fEd.description=this.value">${vxEsc(f.description)}</textarea>
        <label class="mt-2 block">Message de remerciement</label><textarea rows="2" oninput="fEd.messageFin=this.value" placeholder="Merci pour votre réponse !">${vxEsc(f.messageFin)}</textarea>
        <label class="mt-2 block">Indicateur Qualiopi (facultatif)</label><input value="${vxEsc(f.indicateur)}" oninput="fEd.indicateur=this.value" placeholder="ex. 30, 33">
        <p class="text-xs font-black text-gray-400 uppercase mt-4 mb-2">Apparence</p>
        <div style="display:flex;flex-wrap:wrap;gap:.4rem">${FM_PALETTES.map(p => `<button title="Palette" onclick="fEd.theme.couleur='${p[0]}';fEd.theme.accent='${p[1]}';fmRendreEditeur()" style="width:34px;height:34px;border-radius:50%;border:${th.couleur === p[0] ? '3px solid #111' : '2px solid #fff'};box-shadow:0 2px 6px rgba(0,0,0,.2);background:linear-gradient(135deg,${p[0]} 50%,${p[1]} 50%);cursor:pointer"></button>`).join('')}</div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:.4rem;margin-top:.5rem">
          <div><label>Couleur</label><input type="color" value="${th.couleur}" onchange="fEd.theme.couleur=this.value;fmRendreEditeur()" style="height:36px;padding:2px"></div>
          <div><label>Accent</label><input type="color" value="${th.accent}" onchange="fEd.theme.accent=this.value;fmRendreEditeur()" style="height:36px;padding:2px"></div></div>
        <label class="mt-2 block">Fond</label><select onchange="fEd.theme.fond=this.value">${[['degrade', 'Dégradé lumineux'], ['aurore', 'Aurore'], ['clair', 'Clair'], ['sombre', 'Sombre élégant']].map(o => `<option value="${o[0]}" ${th.fond === o[0] ? 'selected' : ''}>${o[1]}</option>`).join('')}</select>
        <label class="mt-2 block">Typographie</label><select onchange="fEd.theme.police=this.value">${[['moderne', 'Moderne (Fraunces + Inter)'], ['ronde', 'Ronde (Nunito)']].map(o => `<option value="${o[0]}" ${th.police === o[0] ? 'selected' : ''}>${o[1]}</option>`).join('')}</select>
        <label class="mt-2 block">Image de bandeau (URL, facultatif)</label><input value="${vxEsc(th.image || '')}" oninput="fEd.theme.image=this.value" placeholder="https://…">
        ${f.id && f.statut === 'publié' ? `<a href="${vxEsc(formsPageUrl + '?f=' + f.id)}" target="_blank" class="btn-ghost text-sm mt-4 inline-block">👁 Voir le formulaire en ligne</a>` : '<p class="text-xs text-gray-400 mt-4">Publiez le formulaire pour obtenir son lien public et son QR code.</p>'}
      </div>
      <div style="padding:1.2rem 1.4rem;overflow:auto;background:linear-gradient(135deg,#eef3fb,#fdf1ea)">
        <div id="fmQuestions"></div>
        <div style="background:#fff;border:2px dashed #cbd5e1;border-radius:1rem;padding:.9rem;margin-top:.9rem">
          <p class="text-xs font-black text-gray-400 uppercase mb-2">Ajouter</p>
          <div style="display:flex;flex-wrap:wrap;gap:.35rem">${F_TYPES.map(t => `<button onclick="fmAjouter('${t[0]}')" style="background:#f8fafc;border:1px solid #e5e7eb;border-radius:.6rem;padding:.4rem .6rem;font-size:.75rem;font-weight:700;cursor:pointer">${t[2]} ${t[1]}</button>`).join('')}</div>
        </div>
      </div>
    </div></div>`;
  fmRendreQuestions();
}
function fmRendreQuestions() {
  const box = document.getElementById('fmQuestions'), qs = fEd.questions, th = fEd.theme;
  box.innerHTML = qs.map((q, i) => {
    const t = F_LIB[q.type] || F_LIB.texte;
    if (q.type === 'section') return `<div style="display:flex;gap:.5rem;align-items:center;margin:1rem 0 .4rem">
      <span style="width:24px;height:3px;background:${th.accent};border-radius:3px"></span>
      <input value="${vxEsc(q.label)}" oninput="fEd.questions[${i}].label=this.value" placeholder="Titre de la section" style="font-weight:900;font-size:1.05rem;background:transparent;border:0;border-bottom:1px dashed #94a3b8;border-radius:0;flex:1">
      ${fmOutils(i)}</div>`;
    const opts = ['choix', 'cases', 'liste'].includes(q.type) ? `<label class="mt-2 block">Options (une par ligne)</label><textarea rows="${Math.max(3, (q.options || []).length + 1)}" oninput="fEd.questions[${i}].options=this.value.split('\\n').map(s=>s.trim()).filter(Boolean)">${vxEsc((q.options || []).join('\n'))}</textarea>` : '';
    const ech = q.type === 'echelle' ? `<div style="display:grid;grid-template-columns:70px 70px 1fr 1fr;gap:.4rem;margin-top:.5rem">
      <div><label>De</label><select onchange="fEd.questions[${i}].min=+this.value">${[0, 1].map(v => `<option ${q.min == v ? 'selected' : ''}>${v}</option>`).join('')}</select></div>
      <div><label>À</label><select onchange="fEd.questions[${i}].max=+this.value">${[3, 4, 5, 6, 7, 10].map(v => `<option ${(q.max || 5) == v ? 'selected' : ''}>${v}</option>`).join('')}</select></div>
      <div><label>Libellé min</label><input value="${vxEsc(q.minLabel || '')}" oninput="fEd.questions[${i}].minLabel=this.value"></div>
      <div><label>Libellé max</label><input value="${vxEsc(q.maxLabel || '')}" oninput="fEd.questions[${i}].maxLabel=this.value"></div></div>` : '';
    return `<div style="background:#fff;border-radius:1rem;padding:.9rem 1rem;margin-top:.6rem;box-shadow:0 8px 20px -16px rgba(15,23,42,.4);border-left:4px solid ${th.couleur}">
      <div style="display:flex;justify-content:space-between;align-items:center;gap:.5rem;margin-bottom:.4rem">
        <select onchange="fEd.questions[${i}].type=this.value;fmRendreQuestions()" style="width:auto;font-size:.75rem;font-weight:800;padding:.3rem .5rem">${F_TYPES.filter(x => x[0] !== 'section').map(x => `<option value="${x[0]}" ${x[0] === q.type ? 'selected' : ''}>${x[2]} ${x[1]}</option>`).join('')}</select>
        <div style="display:flex;align-items:center;gap:.6rem"><label style="display:flex;align-items:center;gap:.3rem;font-size:.75rem;font-weight:700;margin:0;text-transform:none"><input type="checkbox" ${q.obligatoire ? 'checked' : ''} onchange="fEd.questions[${i}].obligatoire=this.checked" style="width:auto"> Obligatoire</label>${fmOutils(i)}</div></div>
      <input value="${vxEsc(q.label)}" oninput="fEd.questions[${i}].label=this.value" placeholder="Intitulé de la question" style="font-weight:700">
      <input value="${vxEsc(q.aide || '')}" oninput="fEd.questions[${i}].aide=this.value" placeholder="Aide / précision (facultatif)" style="margin-top:.4rem;font-size:.8rem">
      ${opts}${ech}</div>`;
  }).join('') || '<p class="text-gray-400 text-sm">Ajoutez une première question ci-dessous.</p>';
}
function fmOutils(i) {
  return `<span style="display:flex;gap:.2rem">${['↑', '↓', '⧉', '🗑'].map((s, k) => `<button title="${['Monter', 'Descendre', 'Dupliquer', 'Supprimer'][k]}" onclick="fmOutil(${i},${k})" style="background:${k === 3 ? '#fef2f2' : '#f3f4f6'};color:${k === 3 ? '#dc2626' : '#374151'};border:0;border-radius:.45rem;width:28px;height:28px;cursor:pointer;font-weight:900">${s}</button>`).join('')}</span>`;
}
function fmOutil(i, k) {
  const q = fEd.questions;
  if (k === 0 && i > 0) [q[i - 1], q[i]] = [q[i], q[i - 1]];
  if (k === 1 && i < q.length - 1) [q[i + 1], q[i]] = [q[i], q[i + 1]];
  if (k === 2) q.splice(i + 1, 0, Object.assign(JSON.parse(JSON.stringify(q[i])), { id: 'q' + Date.now().toString(36) }));
  if (k === 3) q.splice(i, 1);
  fmRendreQuestions();
}
function fmAjouter(type) {
  const q = { id: 'q' + Date.now().toString(36), type, label: type === 'section' ? 'Nouvelle section' : '', obligatoire: false };
  if (['choix', 'cases', 'liste'].includes(type)) q.options = ['Option 1', 'Option 2'];
  if (type === 'echelle') Object.assign(q, { min: 1, max: 5, minLabel: 'Pas du tout', maxLabel: 'Tout à fait' });
  fEd.questions.push(q); fmRendreQuestions();
  const box = document.getElementById('fmQuestions'); box.lastElementChild && box.lastElementChild.scrollIntoView({ behavior: 'smooth', block: 'center' });
}
async function fmSauver() {
  fEd.statut = document.getElementById('fmStatut').value;
  if (!fEd.titre.trim()) { toast('Donnez un titre au formulaire.', false); return; }
  const vides = fEd.questions.filter(q => q.type !== 'section' && !String(q.label || '').trim());
  if (vides.length) { toast(vides.length + ' question(s) sans intitulé.', false); return; }
  const r = await api('saveFormulaire', { formulaire: JSON.stringify(fEd) });
  if (!r.success) { toast(r.message, false); return; }
  fEd.id = r.id; toast(fEd.statut === 'publié' ? 'Formulaire publié ✓ — lien copié' : 'Formulaire enregistré ✓');
  if (fEd.statut === 'publié' && r.lienPublic) { try { navigator.clipboard.writeText(r.lienPublic); } catch (e) {} }
  document.getElementById('fmModal').classList.add('hidden'); loadFormulaires();
}

// ── Réponses & statistiques ──
async function fmReponses(id) {
  const f = formsData.find(x => x.id === id);
  let m = document.getElementById('frModal');
  if (!m) { m = document.createElement('div'); m.id = 'frModal'; m.className = 'modal-overlay hidden'; m.onclick = e => { if (e.target === m) m.classList.add('hidden'); }; document.body.appendChild(m); }
  m.innerHTML = `<div class="modal-box p-6" style="max-width:1000px;width:95vw;max-height:90vh;overflow:auto"><p class="text-gray-400">Chargement des réponses…</p></div>`;
  m.classList.remove('hidden');
  const r = await api('getFormulaireReponses', { id });
  const box = m.firstElementChild;
  if (!r.success) { box.innerHTML = `<p class="text-red-600 font-bold">${vxEsc(r.message)}</p>`; return; }
  const th = f.theme || {}, coul = th.couleur || '#1A4E8A', acc = th.accent || '#F26522';
  let cols, lignes, questions;
  if (r.externe) { cols = r.colonnes; lignes = r.lignes; questions = cols.map((c, i) => ({ id: String(i), label: c, type: 'auto' })); }
  else { questions = r.questions.filter(q => q.type !== 'section'); cols = ['Date'].concat(questions.map(q => q.label)); lignes = r.reponses.map(x => [x.date].concat(questions.map(q => Array.isArray(x.reponses[q.id]) ? x.reponses[q.id].join(', ') : (x.reponses[q.id] ?? '')))); }
  const n = lignes.length;
  const stats = questions.map((q, qi) => {
    const col = r.externe ? qi : qi + 1;
    const vals = lignes.map(l => String(l[col] ?? '').trim()).filter(Boolean);
    if (!vals.length || (r.externe && qi === 0)) return '';
    const nums = vals.map(Number).filter(v => !isNaN(v));
    const estNum = nums.length === vals.length && (q.type !== 'texte');
    if (q.type === 'nps' || (estNum && Math.max(...nums) === 10 && Math.min(...nums) >= 0 && q.type === 'nps')) {
      const p = nums.filter(v => v >= 9).length, d = nums.filter(v => v <= 6).length, score = Math.round((p - d) / nums.length * 100);
      return fmCarteStat(q.label, `<b style="font-size:1.8rem;color:${score >= 0 ? '#059669' : '#dc2626'}">${score > 0 ? '+' : ''}${score}</b><span class="text-xs text-gray-400 ml-2">score de recommandation (${nums.length} rép.)</span>`);
    }
    if (estNum && ['etoiles', 'echelle', 'nombre', 'auto'].includes(q.type)) {
      const moy = nums.reduce((a, b) => a + b, 0) / nums.length, max = q.type === 'etoiles' ? 5 : (q.max || Math.max(...nums, 5));
      return fmCarteStat(q.label, `<b style="font-size:1.8rem;color:${coul}">${moy.toFixed(1).replace('.', ',')}</b><span class="text-gray-400"> / ${max}</span>
        <div style="height:8px;background:#f1f5f9;border-radius:9px;margin-top:.4rem"><div style="height:8px;width:${Math.min(100, moy / max * 100)}%;border-radius:9px;background:linear-gradient(90deg,${coul},${acc})"></div></div><p class="text-xs text-gray-400 mt-1">${nums.length} réponse(s)</p>`);
    }
    const freq = {}; vals.forEach(v => v.split(/,\s*/).forEach(x => { freq[x] = (freq[x] || 0) + 1; }));
    const ent = Object.entries(freq).sort((a, b) => b[1] - a[1]);
    if (ent.length <= 12 && vals.every(v => v.length < 80)) return fmCarteStat(q.label, ent.map(([k, c]) => `<div style="margin-top:.35rem"><div class="flex justify-between text-xs"><span class="font-semibold">${vxEsc(k)}</span><span class="text-gray-400">${c} (${Math.round(c / vals.length * 100)} %)</span></div>
      <div style="height:7px;background:#f1f5f9;border-radius:9px"><div style="height:7px;width:${c / ent[0][1] * 100}%;border-radius:9px;background:${coul}"></div></div></div>`).join(''));
    return fmCarteStat(q.label, `<div style="max-height:160px;overflow:auto">${vals.slice(-12).reverse().map(v => `<p class="text-xs bg-gray-50 rounded-lg px-2 py-1 mt-1">« ${vxEsc(v)} »</p>`).join('')}</div>`, true);
  }).join('');
  box.innerHTML = `<div class="flex justify-between items-start gap-3 mb-4"><div><h3 class="text-xl font-black" style="color:${coul}">📊 ${vxEsc(f.titre)}</h3><p class="text-sm text-gray-500">${n} réponse(s)${r.externe ? ' — formulaire Google' : ''}</p></div>
    <div class="flex gap-2">${n ? `<button onclick="fmCsv()" class="btn-ghost text-sm">⬇️ CSV</button>` : ''}<button onclick="document.getElementById('frModal').classList.add('hidden')" class="btn-ghost text-sm">Fermer</button></div></div>
    ${n ? `<div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:.8rem">${stats}</div>
    <details class="mt-5"><summary class="cursor-pointer font-bold text-sm text-gray-600">Voir le tableau des réponses</summary><div style="overflow:auto;margin-top:.6rem"><table class="text-xs w-full"><thead><tr>${cols.map(c => `<th class="text-left p-2 bg-gray-50">${vxEsc(c)}</th>`).join('')}</tr></thead>
      <tbody>${lignes.slice().reverse().map(l => `<tr class="border-t">${cols.map((c, i) => `<td class="p-2 align-top">${vxEsc(l[i])}</td>`).join('')}</tr>`).join('')}</tbody></table></div></details>` : '<p class="text-gray-400">Aucune réponse pour le moment.</p>'}`;
  window._fmExport = { cols, lignes, titre: f.titre };
}
function fmCarteStat(titre, corps, large) { return `<div style="background:#fff;border:1px solid #eef2f7;border-radius:1rem;padding:.9rem 1rem;${large ? 'grid-column:span 1' : ''}"><p class="text-xs font-black text-gray-500 mb-2">${vxEsc(titre)}</p>${corps}</div>`; }
function fmCsv() {
  const e = window._fmExport; if (!e) return;
  const csv = '﻿' + [e.cols].concat(e.lignes).map(l => l.map(v => '"' + String(v ?? '').replace(/"/g, '""') + '"').join(';')).join('\r\n');
  const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
  a.download = (e.titre || 'reponses').replace(/[^\w-]+/g, '_') + '.csv'; a.click();
}

// ══ AUDIT ESPACE RECRUTEURS (hecg.fr) ↔ CRM ══
let rcData = null;
async function loadRecruteurs() {
  const s = document.getElementById('section-recruteurs');
  s.innerHTML = '<div class="text-center text-gray-400 py-16">Chargement du dernier audit…</div>';
  const r = await api('getAuditRecruteurs');
  if (!r.success) { s.innerHTML = `<p class="text-red-600 font-bold">${vxEsc(r.message)}</p>`; return; }
  rcData = r; renderRecruteurs();
}
function renderRecruteurs() {
  const s = document.getElementById('section-recruteurs'), r = rcData, d = r.dernier;
  const coul = { 'Absent du site': ['#fee2e2', '#b91c1c'], 'À retirer du site': ['#ffedd5', '#c2410c'], 'À passer indisponible': ['#ffedd5', '#c2410c'], 'À repasser disponible': ['#fef3c7', '#92400e'], 'Inconnu du CRM': ['#e0e7ff', '#3730a3'], 'Homonymes': ['#f3f4f6', '#374151'] };
  const types = {}; r.ecarts.forEach(e => types[e.type] = (types[e.type] || 0) + 1);
  const kpis = d ? [['Profils sur le site', d.profilsSite, '#1A4E8A'], ['En recherche (CRM)', d.enRecherche, '#0f766e'], ['Rapprochés', d.presents, '#059669'], ['Affichés indisponibles', d.indisponibles == null ? '—' : d.indisponibles, '#475569'], ['Écarts', d.ecarts, '#dc2626'], ['CV en attente de lecture', d.cvEnAttente, '#d97706']] : [];
  s.innerHTML = `<div class="flex flex-wrap justify-between items-start gap-3 mb-4">
      <div><h1 class="text-2xl font-black text-[#1A4E8A]">🎯 Espace recruteurs ↔ CRM & relances</h1>
      <p class="text-sm text-gray-500 mt-1">Vérifie que chaque étudiant « En recherche » figure sur <a href="https://hecg.fr/espace-recruteurs" target="_blank" class="underline">hecg.fr/espace-recruteurs</a> et que les étudiants placés ou sortis y sont affichés « Indisponible » (profils déjà choisis laissés en vitrine : conforme). Chaque CV est copié dans le Drive et sa date d'édition est reportée dans le dossier étudiant. Audit automatique chaque nuit.</p></div>
      <div class="flex gap-2"><button onclick="rcMdp()" class="btn-ghost text-sm">🔑 ${r.mdpConfigure ? 'Changer le mot de passe' : 'Enregistrer le mot de passe'}</button>
      <button id="btnRcAudit" onclick="rcLancer()" class="btn-orange text-sm" ${r.mdpConfigure ? '' : 'disabled'}>▶ Lancer l'audit maintenant</button></div></div>
    ${!r.mdpConfigure ? '<div class="qualiopi-card" style="border-left:4px solid #f59e0b"><p class="font-bold text-amber-700">Enregistrez une fois le mot de passe de la page « Espace recruteurs » : il est conservé côté serveur, jamais dans le code.</p></div>' : ''}
    ${d ? `<div class="grid grid-cols-2 md:grid-cols-6 gap-3 mb-4">${kpis.map(k => `<div class="bg-white rounded-2xl border border-gray-100 p-4"><p class="text-2xl font-black" style="color:${k[2]}">${k[1]}</p><p class="text-[11px] font-bold text-gray-400 uppercase">${k[0]}</p></div>`).join('')}</div>
      <p class="text-xs text-gray-400 mb-3">Dernier audit : ${vxEsc(d.date)} — ${vxEsc(d.auteur)}${d.notes ? ' · ' + d.notes + ' note(s) ajoutée(s) aux dossiers' : ''}${d.cvEnAttente ? ' · les CV restants sont lus automatiquement toutes les 5 minutes' : ''}</p>` : '<p class="text-gray-400 text-sm mb-3">Aucun audit encore réalisé.</p>'}
    ${Object.keys(types).length ? `<div class="flex flex-wrap gap-2 mb-3">${Object.entries(types).map(([t, n]) => { const c = coul[t] || coul.Homonymes; return `<span class="text-xs font-bold px-3 py-1 rounded-full" style="background:${c[0]};color:${c[1]}">${vxEsc(t)} · ${n}</span>`; }).join('')}</div>` : ''}
    <div class="bg-white rounded-2xl border border-gray-100 overflow-x-auto"><table class="w-full text-sm"><thead class="bg-gray-50 text-[11px] uppercase text-gray-500"><tr><th class="text-left p-3">Écart</th><th class="text-left p-3">Étudiant</th><th class="text-left p-3">Note</th><th class="text-left p-3">CV site</th></tr></thead>
    <tbody>${r.ecarts.length ? r.ecarts.map(e => { const c = coul[e.type] || coul.Homonymes; return `<tr class="border-t align-top">
      <td class="p-3"><span class="text-[11px] font-black px-2 py-1 rounded-lg" style="background:${c[0]};color:${c[1]}">${vxEsc(e.type)}</span></td>
      <td class="p-3">${e.id ? `<a href="dossier.html?id=${encodeURIComponent(e.id)}" target="_blank" class="font-bold text-[#1A4E8A] hover:underline">${vxEsc(e.nom)}</a><div class="text-[10px] text-gray-400">${vxEsc(e.id)}</div>` : `<b>${vxEsc(e.nom)}</b>`}</td>
      <td class="p-3 text-gray-600">${vxEsc(e.detail)}</td><td class="p-3">${e.url ? `<a href="${vxEsc(e.url)}" target="_blank" class="text-blue-600 underline text-xs">PDF</a>` : '—'}</td></tr>`; }).join('')
      : '<tr><td colspan="4" class="text-center text-gray-400 py-8">Aucun écart ✅</td></tr>'}</tbody></table></div>
    <div id="rcRelances" class="mt-6"></div>`;
  loadRelances();
}
async function rcMdp() {
  const mdp = prompt("Mot de passe de la page « Espace recruteurs » de hecg.fr :");
  if (!mdp) return;
  const r = await api('saveMdpRecruteurs', { mdp });
  if (!r.success) { toast(r.message, false); return; }
  toast('Mot de passe enregistré ✓'); loadRecruteurs();
}
async function rcLancer() {
  const b = document.getElementById('btnRcAudit'); b.disabled = true; b.textContent = "⏳ Audit en cours (jusqu'à 3 min)…";
  try {
    const r = await api('lancerAuditRecruteurs');
    if (!r.success) { toast(r.message, false); if (r.mdp) rcMdp(); return; }
    toast(`Audit terminé : ${r.rapport.ecarts.length} écart(s), ${r.rapport.cvLus} CV lus`);
    loadRecruteurs();
  } finally { b.disabled = false; b.textContent = "▶ Lancer l'audit maintenant"; }
}

// ══ RELANCE QUOTIDIENNE DES ÉTUDIANTS EN RECHERCHE ══
let relData = null;
async function loadRelances() {
  const box = document.getElementById('rcRelances'); if (!box) return;
  box.innerHTML = '<p class="text-gray-400 text-sm">Chargement des relances…</p>';
  const r = await api('getRelanceCandidatures');
  if (!r.success) { box.innerHTML = `<p class="text-red-600 text-sm">${vxEsc(r.message)}</p>`; return; }
  relData = r; renderRelances();
}
function renderRelances() {
  const box = document.getElementById('rcRelances'), r = relData, d = r.dernier;
  box.innerHTML = `<div class="bg-white rounded-2xl border border-gray-100 p-5">
    <div class="flex flex-wrap justify-between items-start gap-3">
      <div><h2 class="text-lg font-black text-[#1A4E8A]">📬 Relance quotidienne des étudiants en recherche</h2>
        <p class="text-xs text-gray-500 mt-1">Du lundi au vendredi à partir de ${r.heure} h : un mail rappelle de mettre à jour l'application aux étudiants « En recherche » sans nouvelle candidature depuis plus de 24 h. Copie systématique : <b>${vxEsc(r.copie)}</b>.</p></div>
      <div class="flex items-center gap-3">
        <label class="flex items-center gap-2 text-sm font-bold" style="text-transform:none;margin:0"><input type="checkbox" ${r.active ? 'checked' : ''} onchange="relActiver(this.checked)" style="width:18px;height:18px"> Relance automatique active</label>
        <button id="btnRelEnvoyer" onclick="relEnvoyer()" class="btn-orange text-sm" ${r.cibles.length ? '' : 'disabled'}>✉️ Envoyer maintenant (${r.cibles.length})</button></div></div>
    <p class="text-xs text-gray-400 mt-2">${d ? `Dernier envoi : ${vxEsc(d.date)} — ${d.envoyes} mail(s)${d.erreurs ? ', ' + d.erreurs + ' erreur(s)' : ''} (${vxEsc(d.auteur)})` : 'Aucun envoi pour le moment.'}${r.envoyeAujourdhui ? ' · déjà envoyée aujourd\'hui' : ''}</p>
    <details class="mt-3"><summary class="cursor-pointer text-sm font-bold text-gray-600">👀 ${r.cibles.length} étudiant(s) recevraient la relance aujourd'hui</summary>
      <div class="overflow-x-auto mt-2"><table class="w-full text-xs"><thead class="bg-gray-50 text-gray-500 uppercase"><tr><th class="text-left p-2">Étudiant</th><th class="text-left p-2">Classe</th><th class="text-left p-2">Candidatures</th><th class="text-left p-2">Dernière</th></tr></thead>
      <tbody>${r.cibles.map(c => `<tr class="border-t"><td class="p-2"><a href="dossier.html?id=${encodeURIComponent(c.id)}" target="_blank" class="font-bold text-[#1A4E8A] hover:underline">${vxEsc(c.prenom + ' ' + c.nom)}</a></td><td class="p-2">${vxEsc(c.classe)}</td><td class="p-2">${c.nb}</td><td class="p-2 ${c.derniere ? '' : 'text-red-600 font-bold'}">${vxEsc(c.derniere || 'aucune')}</td></tr>`).join('')}</tbody></table></div></details></div>`;
}
async function relActiver(on) {
  const r = await api('saveRelanceCandidatures', { active: on });
  if (!r.success) { toast(r.message, false); return; }
  toast(on ? 'Relance quotidienne activée ✓' : 'Relance quotidienne désactivée');
}
async function relEnvoyer() {
  if (!confirm(`Envoyer maintenant la relance à ${relData.cibles.length} étudiant(s) (copie ${relData.copie}) ?\nLa relance automatique de ce jour sera alors considérée comme faite.`)) return;
  const b = document.getElementById('btnRelEnvoyer'); b.disabled = true; b.textContent = '⏳ Envoi…';
  const r = await api('saveRelanceCandidatures', { envoyerMaintenant: true });
  if (!r.success) { toast(r.message, false); b.disabled = false; return; }
  toast(`${r.envoyes} relance(s) envoyée(s)` + (r.erreurs && r.erreurs.length ? ` — ${r.erreurs.length} erreur(s)` : ' ✓'));
  loadRelances();
}
