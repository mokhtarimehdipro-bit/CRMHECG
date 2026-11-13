// =========================================================================
// MOTEUR DE SYNCHRONISATION GOOGLE CONTACTS V2 — HECG
// Corrections vs V1 :
//   • Protection JSON.parse sur TOUS les appels API
//   • try/catch par utilisateur (un échec n'arrête pas les autres)
//   • Vérification directe de myContacts.memberCount avant toute création
//   • Purge forcée si compteur fantôme détecté (memberCount >> contacts visibles)
//   • Retry 429 dans mettreAJourContactsEnBatch
//   • Màj par lot (batchUpdateContacts, 200/requête) au lieu d'1 patch/contact
//     → ~200x plus rapide, les 8 utilisateurs tiennent dans un seul run
//   • Repli automatique en PATCH individuel si un lot échoue en bloc (1 seul
//     etag périmé rejette sinon les 200 contacts du lot chez Google)
//   • diagnostiquerTous() — état complet des 8 utilisateurs en un seul run
//   • resetterCompte(email) — purge manuelle d'un compte bloqué
// =========================================================================

const SC2_SERVICE_ACCOUNT_EMAIL = 'robot-signature-hecg@signature-hecg-officiel.iam.gserviceaccount.com';

const SC2_PRIVATE_KEY = '-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQCsdka615zrvROl\neHyFa/2aawECkcjwU6vQUhIXAq9JI/IsSNKhGU6cZRKe/wxX9KTyRL6ztnIvW0Pt\n1lAkDy4zKLO1dEpBpHNW+PyQrDIoG/ARwD0T14jV+4SmPYKevMWfCIs4X+vYHnoX\nFIZo4ei2tmbDoAk54eNDKf+oAXVKkFDNoHsQdBXLkynk6igC2dnEnVcGKattezOV\nTnfhPLK+cj/tTaJHRegU1aqdwoDfbA6FVFby7tlXsMc29OOewOLwj52UeW/XmDv8\n7AKNloBgc0QrEYkVS09HlO0hMrTOX7/Ao2JxovPg/R6BPjHxb1oX4Je7rs1onn7z\n5CwNEkslAgMBAAECggEAEmhRZlHrItI8jZXNnKQJHnk7U13iF5ymowaPfbtAoErg\n508ihCViWZkEIspQM/cdv+oMfLwFdf6Ewpb0WNTx9m3quHxgDJ+T2/2ZX4uxksxg\nlFRzcHG53jUJVIEONwkpAq9zxKGgV6HxIBOFwR4Tq6TOVST4tx/gFOQfsHvvW/Tc\nfrdBFx+WbSBV+kcWqB3wq0xHr9LgW54SZ8YM6H0FGald34S3IF/vwyjtLIYgRQ3/\nZNJIjRQRpOlgjUMXazkzluX9RGK3ydL7zP+xQli5BH/aImV+SuzzRGDLxue6zQCB\njYkU7y79uqr6AK8Gq6cf+/ypnviiG+ax2jLJv1Y/oQKBgQDVpw81mXz7xf+KAAdS\n1yGbixhoVDafaTaE2e8CQ2rFlh85DyRy3f7zQ+DyNOXukx3LjZXkbcMz060FaLKy\nX7Qzep2b2C1s9g+25ZIHZvUxncQvPq9K6vsexgMLJZo3dzfUTVj1H5VNzkCklHrB\nX1bYmGDbQVPCsapzCZxsTAt4PQKBgQDOpSrqP+8tLDt+a+RfSFExL+68jXpEsDAd\nJwa405Wg77o2YCumdbfRdZ6q0MBKAuGWRTuY7zUirhnhsgSu++gq+U82+KWYgjR4\nECkKaSDo+XmppZTCRGX+nJOY3y1qd0cB4hW+9cHI/++UGsP5yWbynfYIr/+U4zT6\nypmVS9dlCQKBgQCMbCqg7eqpiC82QmKN3fumwbsfBwqHp50/oAVpFWpdxxdqZztr\ni+D/fkOgrYfaUDMrEDnOUx4TODLl9TRN7H0BwLtKLMFedjNJ4IUj/FV3cNv6uVZ5\nBQxb44UolGRRxDebf+LR6Ro2czMleLld0w2/ehdexAcLVb5TsaNvwmNfeQKBgQC8\nt3RQx6CbJXkTxF6kcbvMatThF2dhEXJvPTPTWU+d0TDC9eMHOxxrSrpjjw78yFLS\nVFnQGizxhgQW7OeAEof9rv8b2coJVGesej2wxz+J5EOqnZAUNjjbZI0aoD6uq02K\nt7laUr/t22YlYKg3FypQSdfmKS0FANZibuIByWhlWQKBgHwQGyNHkEddB5IBNQ4G\nSRC2IzeOb+SvDmj+vgox3tM9mntOTa5nAJueD0tCi+Bor3WA3YxFdcOIf2ly39Ts\niTCYNtR7IBmXhrehndsqKfTz7K0mYm3GWqyQWyVPaRyHKb3Kc3eeMlAnweUYHbfV\nt2epNa0zscpTqcKo+JOs042c\n-----END PRIVATE KEY-----\n';

const SC2_TARGET_USERS = [
  'm.mokhtari@hecg.fr', 'a.mhoudini@hecg.fr', 'n.mas@hecg.fr', 'd.benyounes@hecg.fr',
  'v.velcker@hecg.fr', 'contact@hecg.fr', 'k.bouhassane@hecg.fr',
  'communication@hecg.fr'
];

// ANCIEN type custom People API utilisé un temps pour le téléphone du responsable
// entreprise sur la fiche étudiant. Abandonné : un même numéro dialable présent sur
// DEUX fiches Google Contacts distinctes (celle du MA lui-même — sc2ExtractPartenaires —
// et celle de l'étudiant) fait que iOS/WhatsApp résolvent l'identité de l'appelant de
// façon incohérente pour ce numéro (confusion de nom dans les deux sens constatée en
// usage réel). sc2BuildPerson n'émet plus ce type ; la constante ne sert plus qu'à
// filtrer/nettoyer les résidus laissés par d'anciens runs (voir sc2ResoudrePhonesExistant,
// l'index de matching dans sc2SyncerContacts, et l'audit sansTel).
const SC2_TYPE_TEL_RESPONSABLE = 'Responsable entreprise';

// ==========================================
// HELPERS — appel API sécurisé (protection JSON.parse)
// ==========================================

function sc2ApiCall(url, options) {
  options = options || {};
  options.muteHttpExceptions = true;
  const resp = UrlFetchApp.fetch(url, options);
  const raw = resp.getContentText();
  const code = resp.getResponseCode();
  let data;
  try { data = JSON.parse(raw); }
  catch(e) {
    return { _parseError: true, _httpCode: code, _raw: raw.substring(0, 400) };
  }
  return data;
}

function sc2B64url(input) {
  const bytes = (typeof input === 'string') ? Utilities.newBlob(input).getBytes() : input;
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, '');
}

function sc2GetToken(userEmail) {
  const now = Math.floor(Date.now() / 1000);
  const header = sc2B64url(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
  const claim  = sc2B64url(JSON.stringify({
    iss: SC2_SERVICE_ACCOUNT_EMAIL, sub: userEmail,
    scope: 'https://www.googleapis.com/auth/contacts',
    aud: 'https://oauth2.googleapis.com/token',
    iat: now, exp: now + 3600
  }));
  const sig = sc2B64url(Utilities.computeRsaSha256Signature(header + '.' + claim, SC2_PRIVATE_KEY));
  const jwt = header + '.' + claim + '.' + sig;
  const res = sc2ApiCall('https://oauth2.googleapis.com/token', {
    method: 'post',
    payload: { grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: jwt }
  });
  if (res._parseError || !res.access_token)
    throw new Error('Auth échouée : ' + (res._raw || JSON.stringify(res.error || res)));
  return res.access_token;
}

function sc2AuthHeader(token) { return { Authorization: 'Bearer ' + token }; }

// ==========================================
// DIAGNOSTIC — ÉTAT DE TOUS LES UTILISATEURS
// ==========================================
// Lancer en priorité pour identifier le blocage.
// Examine : connexions visibles, myContacts.memberCount, groupes utilisateur,
// test de création de contact.

function diagnostiquerTous() {
  Logger.log('╔══════════════════════════════════════════════╗');
  Logger.log('║  DIAGNOSTIC SYNCHRO CONTACTS — TOUS USERS  ║');
  Logger.log('╚══════════════════════════════════════════════╝');

  for (const email of SC2_TARGET_USERS) {
    Logger.log('\n── ' + email + ' ──');
    let token;
    try { token = sc2GetToken(email); }
    catch(e) { Logger.log('  ❌ Auth : ' + e.message); continue; }

    // 1. connections.list — combien Google voit-il de contacts ?
    const connData = sc2ApiCall(
      'https://people.googleapis.com/v1/people/me/connections?personFields=names&pageSize=1',
      { headers: sc2AuthHeader(token) }
    );
    if (connData._parseError) {
      Logger.log('  ❌ connections.list : réponse non-JSON (HTTP ' + connData._httpCode + ') : ' + connData._raw);
    } else if (connData.error) {
      Logger.log('  ❌ connections.list : [' + connData.error.code + '] ' + connData.error.message);
    } else {
      Logger.log('  connections.list → totalItems = ' + (connData.totalItems || connData.totalPeople || 0));
    }

    // 2. myContacts.memberCount — le VRAI compteur de quota
    const mcData = sc2ApiCall(
      'https://people.googleapis.com/v1/contactGroups/myContacts?maxMembers=0',
      { headers: sc2AuthHeader(token) }
    );
    if (mcData._parseError) {
      Logger.log('  ❌ myContacts.memberCount : réponse non-JSON');
    } else if (mcData.error) {
      Logger.log('  ❌ myContacts : [' + mcData.error.code + '] ' + mcData.error.message);
    } else {
      const mc = mcData.memberCount !== undefined ? mcData.memberCount : 'N/A';
      Logger.log('  myContacts.memberCount = ' + mc + (mc > 0 ? (mc > 100 ? ' ⚠️ POTENTIELLEMENT FANTÔME' : '') : ' ✅'));
    }

    // 3. Groupes utilisateur
    const grpData = sc2ApiCall(
      'https://people.googleapis.com/v1/contactGroups',
      { headers: sc2AuthHeader(token) }
    );
    if (grpData._parseError || grpData.error) {
      Logger.log('  ❌ contactGroups : erreur');
    } else {
      const userGrps = (grpData.contactGroups || []).filter(g => g.groupType === 'USER_CONTACT_GROUP');
      Logger.log('  Groupes utilisateur : ' + userGrps.length + (userGrps.length ? ' → ' + userGrps.map(g=>g.name).join(', ') : ''));
    }

    // 4. Test de création (sans membership → "Other Contacts" — ne touche pas myContacts)
    const testCreate = sc2ApiCall('https://people.googleapis.com/v1/people:createContact', {
      method: 'post', contentType: 'application/json',
      headers: sc2AuthHeader(token),
      payload: JSON.stringify({ names: [{ givenName: 'DIAG_TEST', familyName: 'HECG_V2' }], phoneNumbers: [{ value: '+33600000099' }] })
    });
    if (testCreate.resourceName) {
      Logger.log('  Test createContact (sans membership) : ✅ OK → ' + testCreate.resourceName);
      // Nettoyage immédiat du contact test
      sc2ApiCall('https://people.googleapis.com/v1/people:batchDeleteContacts', {
        method: 'post', contentType: 'application/json',
        headers: sc2AuthHeader(token),
        payload: JSON.stringify({ resourceNames: [testCreate.resourceName] })
      });
    } else {
      Logger.log('  Test createContact (sans membership) : ❌ ' + JSON.stringify(testCreate.error || testCreate));
    }
  }

  Logger.log('\n══ FIN DIAGNOSTIC ══');
  Logger.log('⬆ Si myContacts.memberCount >> totalItems → exécuter resetterCompte(email) pour chaque compte bloqué.');
}

// ==========================================
// AUDIT + IMPORT AUTOMATIQUE DES TÉLÉPHONES MANQUANTS
// ==========================================
// La synchro complète (synchroniserTousLesContacts2) réimporte déjà le tél manquant à
// chaque passage, mais elle tourne par rotation sur les 8 comptes avec un budget par
// utilisateur : un contact peut donc attendre plusieurs jours avant d'être corrigé.
// Cet outil audite IMMÉDIATEMENT les fiches complètes (jamais les fiches de diffusion
// HECG_*/FORM-*, qui n'ont volontairement aucun tél) sans numéro sur Google, vérifie si
// le CRM en a un, et l'importe tout de suite par un PATCH minimal (ne touche QUE le
// champ téléphone, rien d'autre sur la fiche).
//
// UTILISATION :
//   1. auditerTelephonesManquantsTous()             → tous les comptes, avec import auto
//   2. auditerTelephonesManquants('user@hecg.fr')    → un seul compte
// Les fiches pour lesquelles le CRM n'a PAS de téléphone non plus sont listées à part
// dans les logs ("à compléter à la source") — impossible de les importer, il faut
// d'abord renseigner le numéro dans le Sheet.

function auditerTelephonesManquantsTous() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tousLesContacts = [...sc2ExtractEtudiants(ss), ...sc2ExtractPartenaires(ss), ...sc2ExtractPersonnel(ss)];
  Logger.log('🔍 Audit téléphones manquants — ' + tousLesContacts.length + ' contacts source, 8 comptes.');

  let totalSansTel = 0, totalImportes = 0, totalSansCRM = 0;
  for (const email of SC2_TARGET_USERS) {
    const res = auditerTelephonesManquants(email, tousLesContacts);
    totalSansTel  += res.sansTel;
    totalImportes += res.importes;
    totalSansCRM  += res.sansCRM;
    Utilities.sleep(1000);
  }
  Logger.log('\n══ BILAN AUDIT ══');
  Logger.log('📇 ' + totalSansTel + ' fiche(s) sans téléphone détectée(s) au total (tous comptes confondus, doublons inter-comptes possibles).');
  Logger.log('✅ ' + totalImportes + ' importée(s) automatiquement depuis le CRM.');
  Logger.log('⚠️ ' + totalSansCRM + ' sans aucun numéro dans le CRM non plus — à compléter à la source (voir détail ci-dessus).');
}

function auditerTelephonesManquants(emailUtilisateur, tousLesContacts) {
  const email = emailUtilisateur || SC2_TARGET_USERS[0];
  if (!tousLesContacts) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    tousLesContacts = [...sc2ExtractEtudiants(ss), ...sc2ExtractPartenaires(ss), ...sc2ExtractPersonnel(ss)];
  }
  Logger.log('\n── Audit téléphones : ' + email + ' ──');

  let token;
  try { token = sc2GetToken(email); }
  catch(e) { Logger.log('  ❌ Auth : ' + e.message); return { sansTel: 0, importes: 0, sansCRM: 0 }; }

  const existants = sc2LireContactsExistants(token);

  // Index CRM par email et par nom normalisé (même logique que sc2SyncerContacts).
  const emailIdx = {}, nameIdx = {};
  tousLesContacts.forEach(c => {
    (c.emails || (c.email && c.email.includes('@') ? [c.email] : [])).forEach(em => {
      const k = String(em).trim().toLowerCase();
      if (k && !emailIdx[k]) emailIdx[k] = c;
    });
    const nk = sc2NameKey(c.nom, c.prenom);
    if (nk && !nameIdx[nk]) nameIdx[nk] = c;
  });

  // Le tél du responsable entreprise (type SC2_TYPE_TEL_RESPONSABLE) ne compte pas comme
  // "numéro propre" du contact : une fiche qui n'a QUE ce numéro doit quand même être
  // détectée comme manquant son propre téléphone.
  const sansTel = existants.filter(e => !sc2EstContactDiffusion(e)
    && !(e.phoneNumbers || []).some(p => p.type !== SC2_TYPE_TEL_RESPONSABLE));
  Logger.log('  📇 ' + sansTel.length + ' fiche(s) complète(s) sans téléphone sur ' + existants.length + ' au total.');

  let importes = 0, sansCRM = 0;
  sansTel.forEach(e => {
    const n = (e.names || [])[0] || {};
    const nomComplet = ((n.givenName || '') + ' ' + (n.familyName || '')).trim() || '(sans nom)';

    let crm = null;
    (e.emailAddresses || []).some(em => {
      const k = String(em.value || '').trim().toLowerCase();
      if (k && emailIdx[k]) { crm = emailIdx[k]; return true; }
      return false;
    });
    if (!crm) {
      const nk = sc2NameKey(n.familyName, n.givenName);
      if (nk && nameIdx[nk]) crm = nameIdx[nk];
    }

    const crmPhone = crm && (crm.phones || (crm.phone ? [crm.phone] : []))[0];
    if (!crmPhone) {
      sansCRM++;
      Logger.log('  ℹ️ ' + nomComplet + ' — aucun numéro dans le CRM non plus (à compléter à la source).');
      return;
    }

    if (sc2ImporterTelephoneUnique(token, e.resourceName, e.etag, crmPhone)) {
      importes++;
      Logger.log('  ✅ ' + nomComplet + ' — ' + crmPhone + ' importé.');
    } else {
      Logger.log('  ⚠️ ' + nomComplet + ' — échec import (voir logs ci-dessus).');
    }
    Utilities.sleep(400);
  });

  Logger.log('  📊 ' + email + ' : ' + importes + ' importé(s), ' + sansCRM + ' sans numéro dans le CRM.');
  return { sansTel: sansTel.length, importes, sansCRM };
}

// PATCH minimal : ne touche QUE le champ téléphone d'un contact (updatePersonFields
// scopé), sans passer par sc2BuildPerson (qui exige une appartenance à un groupe pour
// ne pas renvoyer null — inutile ici, on ne modifie rien d'autre sur la fiche).
// Retente une fois avec un etag frais en cas de conflit (fiche modifiée entre-temps).
function sc2ImporterTelephoneUnique(token, resourceName, etag, phone) {
  const url = 'https://people.googleapis.com/v1/' + resourceName + ':updateContact?updatePersonFields=phoneNumbers';
  const payload = { etag: etag, phoneNumbers: [{ value: phone, type: 'mobile' }] };
  let result = sc2ApiCall(url, { method: 'patch', contentType: 'application/json', headers: sc2AuthHeader(token), payload: JSON.stringify(payload) });

  if (result.error && /etag/i.test(result.error.message || '')) {
    const fresh = sc2ApiCall('https://people.googleapis.com/v1/' + resourceName + '?personFields=names', { headers: sc2AuthHeader(token) });
    if (!fresh._parseError && !fresh.error && fresh.etag) {
      payload.etag = fresh.etag;
      result = sc2ApiCall(url, { method: 'patch', contentType: 'application/json', headers: sc2AuthHeader(token), payload: JSON.stringify(payload) });
    }
  }

  if (result._parseError || result.error) {
    Logger.log('    ⚠️ Import tél ' + resourceName + ' : ' + (result.error ? '[' + result.error.code + '] ' + result.error.message : 'réponse non-JSON'));
    return false;
  }
  return true;
}

// ==========================================
// NETTOYAGE CIBLÉ — RÉSIDUS TÉL RESPONSABLE ENTREPRISE SUR FICHES ÉTUDIANT
// ==========================================
// sc2BuildPerson n'émet plus le tél du responsable entreprise sur la fiche étudiant (un
// même numéro sur 2 fiches faisait que WhatsApp/iOS affichaient parfois le nom de
// l'étudiant en appelant le MA — voir sc2BuildPerson). La synchro normale est censée
// nettoyer automatiquement les fiches déjà synchronisées avec l'ancien comportement
// (sc2ContactAChange détecte l'écart puis sc2MettreAJourEnBatch remplace tout le tableau
// phoneNumbers), mais si le nettoyage tarde (gros volume de contacts, rotation par
// budget sur 8 comptes) ou pour un contrôle immédiat et certain, cet outil retire
// directement et UNIQUEMENT les entrées phoneNumbers de type SC2_TYPE_TEL_RESPONSABLE,
// sans toucher au reste de la fiche (PATCH scopé updatePersonFields=phoneNumbers, comme
// sc2ImporterTelephoneUnique).
//
// UTILISATION :
//   1. nettoyerTelResponsableResiduelTous()             → tous les comptes
//   2. nettoyerTelResponsableResiduel('user@hecg.fr')    → un seul compte

function nettoyerTelResponsableResiduelTous() {
  Logger.log('🔍 Nettoyage résidus tél responsable entreprise — 8 comptes.');
  let total = 0;
  for (const email of SC2_TARGET_USERS) {
    total += nettoyerTelResponsableResiduel(email);
    Utilities.sleep(1000);
  }
  Logger.log('\n══ BILAN ══');
  Logger.log(total > 0
    ? '✅ ' + total + ' fiche(s) nettoyée(s) au total (tous comptes confondus).'
    : '✅ Aucun résidu trouvé — rien à nettoyer.');
}

function nettoyerTelResponsableResiduel(emailUtilisateur) {
  const email = emailUtilisateur || SC2_TARGET_USERS[0];
  Logger.log('\n── Nettoyage résidus : ' + email + ' ──');

  let token;
  try { token = sc2GetToken(email); }
  catch(e) { Logger.log('  ❌ Auth : ' + e.message); return 0; }

  const existants = sc2LireContactsExistants(token);
  const aNettoyer = existants.filter(e =>
    !sc2EstContactDiffusion(e) && (e.phoneNumbers || []).some(p => p.type === SC2_TYPE_TEL_RESPONSABLE)
  );
  Logger.log('  📇 ' + aNettoyer.length + ' fiche(s) avec résidu sur ' + existants.length + ' au total.');

  let nettoyes = 0;
  aNettoyer.forEach(e => {
    const n = (e.names || [])[0] || {};
    const nomComplet = ((n.givenName || '') + ' ' + (n.familyName || '')).trim() || '(sans nom)';
    const phonesRestants = (e.phoneNumbers || []).filter(p => p.type !== SC2_TYPE_TEL_RESPONSABLE);

    const url = 'https://people.googleapis.com/v1/' + e.resourceName + ':updateContact?updatePersonFields=phoneNumbers';
    let payload = { etag: e.etag, phoneNumbers: phonesRestants };
    let result = sc2ApiCall(url, { method: 'patch', contentType: 'application/json', headers: sc2AuthHeader(token), payload: JSON.stringify(payload) });

    if (result.error && /etag/i.test(result.error.message || '')) {
      const fresh = sc2ApiCall('https://people.googleapis.com/v1/' + e.resourceName + '?personFields=names', { headers: sc2AuthHeader(token) });
      if (!fresh._parseError && !fresh.error && fresh.etag) {
        payload.etag = fresh.etag;
        result = sc2ApiCall(url, { method: 'patch', contentType: 'application/json', headers: sc2AuthHeader(token), payload: JSON.stringify(payload) });
      }
    }

    if (result._parseError || result.error) {
      Logger.log('  ⚠️ ' + nomComplet + ' — échec : ' + (result.error ? '[' + result.error.code + '] ' + result.error.message : 'réponse non-JSON'));
    } else {
      nettoyes++;
      Logger.log('  ✅ ' + nomComplet + ' — résidu retiré.');
    }
    Utilities.sleep(400);
  });

  Logger.log('  📊 ' + email + ' : ' + nettoyes + '/' + aNettoyer.length + ' fiche(s) nettoyée(s).');
  return nettoyes;
}

// ==========================================
// AUDIT DES LIBELLÉS MANQUANTS (lecture seule)
// ==========================================
// Compare, pour chaque contact CRM (étudiants/partenaires/personnel), les libellés
// ATTENDUS (calculés depuis les feuilles, même logique que la synchro) aux groupes
// RÉELLEMENT présents sur sa fiche Google. Ne modifie RIEN (aucun appel d'écriture) —
// sert uniquement à localiser précisément où et pourquoi des libellés manquent :
//   • "groupe Google jamais créé" → le groupe lui-même n'existe pas encore côté Google
//     (cause la plus probable d'un manque systématique — voir le bilan en fin de log,
//     qui liste tous les libellés concernés)
//   • fiche trouvée mais libellé absent de ses memberships → la fiche attend juste sa
//     prochaine synchro (synchroniserTousLesContacts2 la corrigera au run suivant)
//   • "aucune fiche Google trouvée" → le contact CRM n'a jamais été matché à une fiche
//     existante (à vérifier : tél/mail/nom cohérents entre le Sheet et Google Contacts)
//
// UTILISATION :
//   1. auditerLibellesManquantsTous()             → tous les comptes
//   2. auditerLibellesManquants('user@hecg.fr')    → un seul compte

function auditerLibellesManquantsTous() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tousLesContacts = [...sc2ExtractEtudiants(ss), ...sc2ExtractPartenaires(ss), ...sc2ExtractPersonnel(ss)];
  Logger.log('🔍 Audit libellés manquants — ' + tousLesContacts.length + ' contacts source, 8 comptes.');

  const groupesJamaisCrees = new Set();
  let totalIncomplets = 0;
  for (const email of SC2_TARGET_USERS) {
    const res = auditerLibellesManquants(email, tousLesContacts);
    totalIncomplets += res.contactsIncomplets;
    res.groupesManquants.forEach(g => groupesJamaisCrees.add(g));
    Utilities.sleep(1000);
  }

  Logger.log('\n══ BILAN AUDIT LIBELLÉS ══');
  Logger.log('📇 ' + totalIncomplets + ' fiche(s) avec au moins un libellé manquant (tous comptes confondus, doublons inter-comptes possibles).');
  if (groupesJamaisCrees.size > 0) {
    Logger.log('⚠️ ' + groupesJamaisCrees.size + ' libellé(s) SANS GROUPE Google créé sur au moins un compte — cause la plus probable des manques : '
      + [...groupesJamaisCrees].join(', '));
  } else if (totalIncomplets > 0) {
    Logger.log('✅ Tous les libellés attendus ont bien un groupe Google créé — les manques viennent donc de fiches pas encore resynchronisées. Relancer synchroniserTousLesContacts2() ou attendre le prochain run nocturne (2h).');
  } else {
    Logger.log('✅ Aucun libellé manquant détecté.');
  }
}

function auditerLibellesManquants(emailUtilisateur, tousLesContacts) {
  const email = emailUtilisateur || SC2_TARGET_USERS[0];
  if (!tousLesContacts) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    tousLesContacts = [...sc2ExtractEtudiants(ss), ...sc2ExtractPartenaires(ss), ...sc2ExtractPersonnel(ss)];
  }
  Logger.log('\n── Audit libellés : ' + email + ' ──');

  let token;
  try { token = sc2GetToken(email); }
  catch(e) { Logger.log('  ❌ Auth : ' + e.message); return { contactsIncomplets: 0, groupesManquants: [] }; }

  // Liste UNIQUEMENT les groupes existants (aucune création) — permet de distinguer
  // "le groupe existe mais le contact n'y est pas encore" de "le groupe n'existe même pas".
  const nomVersRn = {}, rnVersNom = {};
  let nextPageToken = null;
  do {
    const pageUrl = 'https://people.googleapis.com/v1/contactGroups'
      + (nextPageToken ? '?pageToken=' + encodeURIComponent(nextPageToken) : '');
    const data = sc2ApiCall(pageUrl, { headers: sc2AuthHeader(token) });
    if (data._parseError || data.error) { Logger.log('  ⚠️ Lecture des groupes échouée.'); break; }
    (data.contactGroups || []).forEach(g => { nomVersRn[g.name] = g.resourceName; rnVersNom[g.resourceName] = g.name; });
    nextPageToken = data.nextPageToken || null;
  } while (nextPageToken);

  const existants = sc2LireContactsExistants(token);
  const phoneIdx = {}, nameIdx = {}, emailIdx = {};
  existants.forEach(e => {
    if (sc2EstContactDiffusion(e)) return;
    (e.phoneNumbers || []).forEach(p => {
      if (p.type === SC2_TYPE_TEL_RESPONSABLE) return;
      const k = sc2PhoneKey(p.value);
      if (k && !phoneIdx[k]) phoneIdx[k] = e;
    });
    (e.emailAddresses || []).forEach(em => {
      const k = String(em.value || '').trim().toLowerCase();
      if (k && !emailIdx[k]) emailIdx[k] = e;
    });
    const n = (e.names || [])[0];
    if (n) {
      const nk = sc2NameKey(n.familyName, n.givenName);
      if (nk && !nameIdx[nk]) nameIdx[nk] = e;
    }
  });

  let contactsIncomplets = 0, contactsIntrouvables = 0;
  const groupesManquants = new Set();

  tousLesContacts.forEach(c => {
    const labelsAttendus = (c.labels || [c.label]).filter(Boolean);
    if (!labelsAttendus.length) return;

    const phones = c.phones || (c.phone ? [c.phone] : []);
    let existing = null;
    for (const p of phones) { const k = sc2PhoneKey(p); if (k && phoneIdx[k]) { existing = phoneIdx[k]; break; } }
    if (!existing) {
      const emails = c.emails || (c.email && c.email.includes('@') ? [c.email] : []);
      for (const em of emails) { const k = String(em).trim().toLowerCase(); if (k && emailIdx[k]) { existing = emailIdx[k]; break; } }
    }
    if (!existing) {
      const nk = sc2NameKey(c.nom, c.prenom);
      if (nk && nameIdx[nk]) existing = nameIdx[nk];
    }

    const nomComplet = ((c.prenom || '') + ' ' + (c.nom || '')).trim() || c.entreprise || '(sans nom)';

    if (!existing) {
      contactsIntrouvables++;
      Logger.log('  ❓ ' + nomComplet + ' — aucune fiche Google trouvée (pas encore synchronisé ?).');
      return;
    }

    const actuels = new Set(
      (existing.memberships || [])
        .map(m => m.contactGroupMembership && m.contactGroupMembership.contactGroupResourceName)
        .filter(rn => rn && !rn.includes('contactGroups/myContacts') && !rn.includes('contactGroups/starred'))
        .map(rn => rnVersNom[rn])
        .filter(Boolean)
    );

    const manquants = labelsAttendus.filter(l => !actuels.has(l));
    if (manquants.length > 0) {
      contactsIncomplets++;
      const sansGroupe = manquants.filter(l => !nomVersRn[l]);
      sansGroupe.forEach(l => groupesManquants.add(l));
      const suffixe = sansGroupe.length > 0 ? '  [groupe Google jamais créé : ' + sansGroupe.join(', ') + ']' : '';
      Logger.log('  ⚠️ ' + nomComplet + ' — libellé(s) manquant(s) : ' + manquants.join(', ') + suffixe);
    }
  });

  Logger.log('  📊 ' + email + ' : ' + contactsIncomplets + ' fiche(s) incomplète(s), '
    + contactsIntrouvables + ' introuvable(s) sur ' + tousLesContacts.length + ' contacts CRM.');
  return { contactsIncomplets, groupesManquants: [...groupesManquants] };
}

// ==========================================
// DIAGNOSTIC PERMISSIONS WRITE (nouveau)
// ==========================================
// Teste séparément les opérations READ, UPDATE et CREATE pour distinguer :
//   • Problème de permissions (scope contacts manquant dans DWD ou Admin Console)
//   • Vrai problème de quota myContacts
// Lancer sur UN seul utilisateur — utilise le premier dans TARGET_USERS.

function diagnostiquerPermissionsWrite() {
  const email = SC2_TARGET_USERS[0];
  Logger.log('╔═══════════════════════════════════════════════╗');
  Logger.log('║  DIAGNOSTIC WRITE PERMISSIONS : ' + email);
  Logger.log('╚═══════════════════════════════════════════════╝');

  let token;
  try { token = sc2GetToken(email); }
  catch(e) { Logger.log('❌ Auth : ' + e.message); return; }
  Logger.log('✅ Auth JWT OK');

  // ── TEST 1 : READ (connections.list) ──────────────────────────────────
  const rConn = sc2ApiCall(
    'https://people.googleapis.com/v1/people/me/connections?personFields=names,phoneNumbers&pageSize=5',
    { headers: sc2AuthHeader(token) }
  );
  const firstContact = (!rConn._parseError && !rConn.error && (rConn.connections || []).length > 0)
    ? rConn.connections[0] : null;
  Logger.log('TEST 1 — READ connections.list : ' + (rConn.error ? '❌ ' + rConn.error.message : (rConn._parseError ? '❌ non-JSON' : '✅ OK (' + (rConn.connections||[]).length + ' contacts retournés)')));

  // ── TEST 2 : UPDATE d'un contact existant ────────────────────────────
  if (firstContact) {
    const rn = firstContact.resourceName;
    const etag = firstContact.etag || '';
    const nom = ((firstContact.names||[])[0]?.familyName) || 'TEST';
    const patchPayload = { etag: etag, names: [{ givenName: (firstContact.names||[])[0]?.givenName||'', familyName: nom }] };
    const rUpdate = sc2ApiCall(
      'https://people.googleapis.com/v1/' + rn + ':updateContact?updatePersonFields=names',
      { method: 'patch', contentType: 'application/json', headers: sc2AuthHeader(token), payload: JSON.stringify(patchPayload) }
    );
    if (rUpdate.error) {
      Logger.log('TEST 2 — UPDATE contact existant (' + rn + ') : ❌ [' + rUpdate.error.code + '] ' + rUpdate.error.message);
      Logger.log('         → Si UPDATE échoue aussi, c\'est un problème de PERMISSIONS (scope write manquant).');
    } else if (rUpdate._parseError) {
      Logger.log('TEST 2 — UPDATE contact existant : ❌ réponse non-JSON (HTTP ' + rUpdate._httpCode + ')');
    } else {
      Logger.log('TEST 2 — UPDATE contact existant : ✅ OK → les écritures SIMPLES fonctionnent');
    }
  } else {
    Logger.log('TEST 2 — UPDATE : ⚠️ aucun contact existant à mettre à jour (skipped)');
  }

  // ── TEST 3 : CREATE sans membership (→ tentative myContacts) ─────────
  const rCreate = sc2ApiCall('https://people.googleapis.com/v1/people:createContact', {
    method: 'post', contentType: 'application/json',
    headers: sc2AuthHeader(token),
    payload: JSON.stringify({ names: [{ givenName: 'DIAG_WRITE_TEST', familyName: 'HECG_V2' }], phoneNumbers: [{ value: '+33600000099' }] })
  });
  if (rCreate.resourceName) {
    Logger.log('TEST 3 — CREATE contact (sans groupe) : ✅ OK → création fonctionne ! Quota non bloqué.');
    sc2ApiCall('https://people.googleapis.com/v1/people:batchDeleteContacts', {
      method: 'post', contentType: 'application/json', headers: sc2AuthHeader(token),
      payload: JSON.stringify({ resourceNames: [rCreate.resourceName] })
    });
    Logger.log('         → Contact test supprimé.');
  } else if (rCreate.error) {
    Logger.log('TEST 3 — CREATE contact (sans groupe) : ❌ [' + rCreate.error.code + '] ' + rCreate.error.message);
  } else {
    Logger.log('TEST 3 — CREATE : ❌ réponse inattendue : ' + JSON.stringify(rCreate).substring(0,200));
  }

  // ── TEST 4 : Vérifier les scopes OAuth réellement accordés ───────────
  // Appel tokeninfo sur le token JWT pour voir les scopes effectifs
  const rToken = sc2ApiCall('https://www.googleapis.com/oauth2/v1/tokeninfo?access_token=' + token, {});
  if (!rToken._parseError && !rToken.error) {
    Logger.log('TEST 4 — Scopes OAuth du token : ' + (rToken.scope || '(non renvoyé)'));
    const hasWrite = (rToken.scope || '').includes('https://www.googleapis.com/auth/contacts') &&
                    !(rToken.scope || '').includes('readonly');
    Logger.log('         → Scope write contacts : ' + (hasWrite ? '✅ Présent' : '⚠️ ABSENT ou readonly seulement'));
  } else {
    Logger.log('TEST 4 — tokeninfo : ' + (rToken._parseError ? 'non-JSON' : rToken.error?.message));
  }

  Logger.log('');
  Logger.log('══ INTERPRÉTATION ══════════════════════════════════════════');
  Logger.log('• TEST 2 ✅ + TEST 3 ❌ MY_CONTACTS_OVERFLOW : quota myContacts — attendre 24h');
  Logger.log('• TEST 2 ❌ + TEST 3 ❌ même erreur           : PERMISSIONS — vérifier Google Admin Console');
  Logger.log('• TEST 4 scope absent/readonly                : scope DWD incomplet — corriger Admin Console');
  Logger.log('• TEST 3 ✅                                   : tout fonctionne, problème résolu');
  Logger.log('════════════════════════════════════════════════════════════');
}

// ==========================================
// VIDAGE CORBEILLE VIA contactGroups/all
// ==========================================
// L'API GData étant fermée, on utilise le groupe système "all" (People API) qui peut
// exposer les resourceNames de contacts en corbeille non retournés par connections.list.
// Fonctionne car Google compte active+corbeille dans la limite MY_CONTACTS_OVERFLOW_COUNT.
//
// UTILISATION :
//   1. viderCorbeilleViaGroupesTous()          → tous les utilisateurs
//   2. viderCorbeilleViaGroupes('user@hecg.fr') → un seul utilisateur

function viderCorbeilleViaGroupesTous() {
  Logger.log('╔═══════════════════════════════════════════════╗');
  Logger.log('║  VIDAGE CORBEILLE — TOUS UTILISATEURS        ║');
  Logger.log('╚═══════════════════════════════════════════════╝');
  for (let i = 0; i < SC2_TARGET_USERS.length; i++) {
    viderCorbeilleViaGroupes(SC2_TARGET_USERS[i]);
    if (i < SC2_TARGET_USERS.length - 1) Utilities.sleep(3000);
  }
  Logger.log('\n✅ Terminé. Relancez diagnostiquerPermissionsWrite() pour vérifier.');
}

function viderCorbeilleViaGroupes(emailUtilisateur) {
  const email = emailUtilisateur || SC2_TARGET_USERS[0];
  Logger.log('\n🗑️ Vidage corbeille : ' + email);

  let token;
  try { token = sc2GetToken(email); }
  catch(e) { Logger.log('  ❌ Auth : ' + e.message); return; }

  // ── Phase 1 : Récupérer TOUS les memberResourceNames du groupe "all" ────
  // "all" est le groupe système Google Contacts qui inclut tout (actifs + potentiellement corbeille)
  const allGroupData = sc2ApiCall(
    'https://people.googleapis.com/v1/contactGroups/all?maxMembers=30000',
    { headers: sc2AuthHeader(token) }
  );

  if (allGroupData._parseError || allGroupData.error) {
    Logger.log('  ❌ contactGroups/all : ' + JSON.stringify(allGroupData.error || allGroupData._raw));
    return;
  }

  const allRns = allGroupData.memberResourceNames || [];
  Logger.log('  contactGroups/all : memberCount=' + (allGroupData.memberCount || 0)
    + ' | memberResourceNames retournés : ' + allRns.length);

  // ── Phase 2 : Récupérer les contacts ACTIFS (connections.list) ──────────
  const activeContacts = sc2LireContactsExistants(token);
  const activeRns = new Set(activeContacts.map(c => c.resourceName).filter(Boolean));
  Logger.log('  Contacts actifs (connections.list) : ' + activeRns.size);

  // ── Phase 3 : Différence = contacts en corbeille ─────────────────────────
  const corbeille = allRns.filter(rn => !activeRns.has(rn));
  Logger.log('  Contacts en corbeille détectés : ' + corbeille.length);

  if (corbeille.length === 0) {
    Logger.log('  ✅ Aucun contact en corbeille détecté via cette méthode.');
    // Informer si le memberCount total correspond bien aux actifs seulement
    if ((allGroupData.memberCount || 0) > activeRns.size) {
      Logger.log('  ⚠️ MAIS memberCount(' + allGroupData.memberCount + ') > actifs('
        + activeRns.size + ') → des contacts de corbeille sont comptés mais non accessibles.');
      Logger.log('  → Videz manuellement : contacts.google.com → Corbeille → Vider la corbeille.');
    }
    return;
  }

  // ── Phase 4 : Suppression définitive par lots de 500 ─────────────────────
  let totalDeleted = 0;
  for (let i = 0; i < corbeille.length; i += 500) {
    const batch = corbeille.slice(i, i + 500);
    const result = sc2ApiCall('https://people.googleapis.com/v1/people:batchDeleteContacts', {
      method: 'post', contentType: 'application/json',
      headers: sc2AuthHeader(token),
      payload: JSON.stringify({ resourceNames: batch })
    });
    if (result._parseError || result.error) {
      Logger.log('  ⚠️ Lot ' + (Math.floor(i / 500) + 1) + ' : erreur → '
        + JSON.stringify(result.error || result._raw).substring(0, 200));
    } else {
      totalDeleted += batch.length;
      Logger.log('  ✅ Lot ' + (Math.floor(i / 500) + 1) + ' : '
        + batch.length + ' supprimés (total : ' + totalDeleted + '/' + corbeille.length + ')');
    }
    Utilities.sleep(800);
  }

  Logger.log('  📊 ' + email + ' : ' + totalDeleted + '/' + corbeille.length
    + ' contacts supprimés définitivement.');
}

// ==========================================
// RESET D'UN COMPTE BLOQUÉ (fantôme)
// ==========================================
// À lancer sur chaque utilisateur dont myContacts.memberCount est anormalement élevé.
// 1. Supprime tous les contacts visibles (connections.list)
// 2. Supprime tous les groupes utilisateur (deleteContacts=false → reset compteur interne)
// 3. Attend 3s pour propagation
// La synchro normale recrée tout proprement.

function resetterCompte(emailUtilisateur) {
  const email = emailUtilisateur || SC2_TARGET_USERS[0];
  Logger.log('🔄 Reset du compte : ' + email);

  let token;
  try { token = sc2GetToken(email); }
  catch(e) { Logger.log('❌ Auth : ' + e.message); return; }

  // 1. Lire tous les contacts existants et les supprimer par lots de 500
  const contacts = sc2LireContactsExistants(token);
  Logger.log('  → ' + contacts.length + ' contacts à supprimer...');
  if (contacts.length > 0) {
    const rns = contacts.map(c => c.resourceName).filter(Boolean);
    for (let i = 0; i < rns.length; i += 500) {
      sc2ApiCall('https://people.googleapis.com/v1/people:batchDeleteContacts', {
        method: 'post', contentType: 'application/json',
        headers: sc2AuthHeader(token),
        payload: JSON.stringify({ resourceNames: rns.slice(i, i + 500) })
      });
      Utilities.sleep(1000);
    }
    Logger.log('  ✅ ' + rns.length + ' contacts supprimés.');
  }

  // 2. Supprimer tous les groupes (reset compteur fantôme)
  sc2PurgerGroupes(token);

  Utilities.sleep(3000);
  Logger.log('✅ Reset terminé pour ' + email + '. Relancez synchroniserTousLesContacts().');
}

// Raccourci pour resetter TOUS les comptes bloqués (à utiliser avec prudence)
function resetterTousLesComptes() {
  Logger.log('⚠️ Reset de TOUS les comptes — cela supprimera TOUS les contacts pour les 8 utilisateurs.');
  for (const email of SC2_TARGET_USERS) {
    resetterCompte(email);
    Utilities.sleep(2000);
  }
}

// ==========================================
// POINT D'ENTRÉE PRINCIPAL — SMART SYNC V2
// ==========================================

// Budget global par exécution : Apps Script coupe à 30 min (comptes Workspace).
// On se limite à 25 min pour laisser une marge de sécurité (cleanup, logs).
const SC2_RUN_BUDGET_MS  = 25 * 60 * 1000;
// Budget max par utilisateur : évite qu'un seul compte avec un gros arriéré
// (ex. rattrapage massif) monopolise tout le budget du run et bloque les 7 autres.
const SC2_USER_BUDGET_MS = 5  * 60 * 1000;
const SC2_PROP_NEXT_USER_INDEX = 'SC2_NEXT_USER_INDEX';

function synchroniserTousLesContacts2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const dataEtu       = sc2ExtractEtudiants(ss);
  const dataPart      = sc2ExtractPartenaires(ss);
  const dataPerso     = sc2ExtractPersonnel(ss);
  const dataDiffusion = [...sc2ExtractGroupesDiffusion(ss), ...sc2ExtractGroupesFormateursParClasse(ss)];
  const tousLesContacts = [...dataEtu, ...dataPart, ...dataPerso];

  Logger.log('🔍 ' + tousLesContacts.length + ' contacts source + ' + dataDiffusion.length
    + ' membre(s) de diffusion HECG_*/FORM-*. Début de la synchronisation V2...');

  // Contrôle : contacts source (étudiants + pro) sans numéro de téléphone exploitable.
  // Ils sont quand même synchronisés (nom/mail/labels) — le téléphone sera importé dès
  // qu'il sera renseigné dans le CRM. Ce log sert juste à repérer les fiches à compléter.
  const sansTelephone = tousLesContacts.filter(c => !((c.phones && c.phones.length > 0) || c.phone));
  if (sansTelephone.length > 0) {
    Logger.log('⚠️ ' + sansTelephone.length + ' contact(s) source sans téléphone exploitable (à compléter dans le CRM) : '
      + sansTelephone.slice(0, 30).map(c => (String(c.prenom || '') + ' ' + String(c.nom || '')).trim() || c.entreprise || '?').join(', ')
      + (sansTelephone.length > 30 ? '…' : ''));
  }

  const tousLibelles = new Set();
  tousLesContacts.forEach(c => {
    (c.labels || [c.label]).forEach(l => { if (l) tousLibelles.add(l); });
  });
  dataDiffusion.forEach(m => {
    (m.labels || []).forEach(l => { if (l) tousLibelles.add(l); });
  });

  const contactsValides = tousLesContacts.filter(c =>
    (c.phones && c.phones.length > 0) || c.phone || c.entreprise || c.nom || c.prenom
  );

  // Rotation : on reprend là où le run précédent s'est arrêté, pour que
  // les 8 comptes avancent tour à tour au lieu de rester bloqués sur le premier.
  const props = PropertiesService.getScriptProperties();
  const startIndex = (parseInt(props.getProperty(SC2_PROP_NEXT_USER_INDEX), 10) || 0) % SC2_TARGET_USERS.length;
  const runStart = Date.now();

  let traites = 0;

  for (let offset = 0; offset < SC2_TARGET_USERS.length; offset++) {
    const idx = (startIndex + offset) % SC2_TARGET_USERS.length;
    const userEmail = SC2_TARGET_USERS[idx];

    if (Date.now() - runStart > SC2_RUN_BUDGET_MS) {
      props.setProperty(SC2_PROP_NEXT_USER_INDEX, String(idx));
      Logger.log('\n⏰ Budget global (' + Math.round(SC2_RUN_BUDGET_MS / 60000) + ' min) atteint. '
        + traites + '/' + SC2_TARGET_USERS.length + ' utilisateur(s) traité(s) cette fois. '
        + 'Reprise à ' + userEmail + ' au prochain run.');
      return;
    }

    Logger.log('\n➡️ Traitement : ' + userEmail);
    try {
      // Chaque utilisateur dispose d'un budget propre, plafonné par ce qu'il
      // reste du budget global du run.
      const userDeadline = Math.min(runStart + SC2_RUN_BUDGET_MS, Date.now() + SC2_USER_BUDGET_MS);
      sc2SyncUser(userEmail, contactsValides, tousLibelles, userDeadline, dataDiffusion);
    } catch(e) {
      Logger.log('❌ Erreur non récupérée pour ' + userEmail + ' : ' + e.message + '\n' + (e.stack || ''));
    }
    traites++;
  }

  // Les 8 utilisateurs ont été traités dans ce run → prochain run repart du début.
  props.setProperty(SC2_PROP_NEXT_USER_INDEX, '0');
  Logger.log('\n✅ Synchronisation V2 terminée (cycle complet des ' + SC2_TARGET_USERS.length + ' utilisateurs).');
}

function sc2SyncUser(email, contactsValides, tousLibelles, deadline, dataDiffusion) {
  let token;
  try { token = sc2GetToken(email); }
  catch(e) { Logger.log('  ❌ Auth : ' + e.message); return; }

  // 1. Vérifier l'état de myContacts AVANT tout
  const mcState = sc2CheckMyContacts(token);
  Logger.log('  myContacts.memberCount = ' + mcState.memberCount + ' | connections = ?');

  // 2. Préparer les groupes
  const groupesMap = sc2PreparerGroupes(token, tousLibelles);

  // 3. Lire les contacts existants
  const existants = sc2LireContactsExistants(token);
  Logger.log('  📖 ' + existants.length + ' contacts existants lus.');

  // 4. Détecter les compteurs fantômes et purger si nécessaire
  //    Critère : memberCount > contacts visibles + marge de 50
  const needsPurge = mcState.memberCount > existants.length + 50;
  if (needsPurge) {
    Logger.log('  ⚠️ Compteur fantôme détecté (memberCount=' + mcState.memberCount + ' > existants=' + existants.length + '). Purge...');
    sc2PurgerGroupes(token);
    Utilities.sleep(3000);
    // Reconstruire les groupes après purge
    Object.keys(groupesMap).forEach(k => delete groupesMap[k]);
    const freshMap = sc2PreparerGroupes(token, tousLibelles);
    Object.assign(groupesMap, freshMap);
    // Relire les contacts (maintenant vide après purge propre)
    existants.splice(0, existants.length, ...sc2LireContactsExistants(token));
    Logger.log('  📖 Après purge : ' + existants.length + ' contacts.');
  }

  // 5. Smart sync
  const stats = sc2SyncerContacts(token, contactsValides, existants, groupesMap, deadline);
  Logger.log('  ✅ ' + email + ' : ' + stats.created + ' créés, ' + stats.updated + ' màj, ' + stats.deleted + ' supprimés.');

  // 5bis. Groupes de diffusion HECG_* (mailing-list par classe / formateurs — fiches
  // séparées, volontairement isolées de la fiche complète, voir sc2ExtractGroupesDiffusion)
  if (dataDiffusion && dataDiffusion.length > 0 && (!deadline || Date.now() < deadline)) {
    sc2SyncerDiffusion(token, dataDiffusion, existants, groupesMap, deadline);
  }

  // 6. Nettoyer les groupes obsolètes (rapide, sauté si le budget est déjà consommé)
  if (!deadline || Date.now() < deadline) {
    sc2NettoyerGroupesObsoletes(token, tousLibelles);
  }
}

// ==========================================
// LECTURE DE L'ÉTAT myContacts
// ==========================================

function sc2CheckMyContacts(token) {
  const data = sc2ApiCall(
    'https://people.googleapis.com/v1/contactGroups/myContacts?maxMembers=0',
    { headers: sc2AuthHeader(token) }
  );
  if (data._parseError || data.error) return { memberCount: 0, error: data.error || data._raw };
  return { memberCount: data.memberCount || 0, error: null };
}

// ==========================================
// LECTURE DES CONTACTS EXISTANTS (sécurisée)
// ==========================================

function sc2LireContactsExistants(token) {
  const baseUrl = 'https://people.googleapis.com/v1/people/me/connections'
    + '?personFields=names,phoneNumbers,emailAddresses,organizations,addresses,memberships,userDefined,relations'
    + '&pageSize=1000';
  const contacts = [];
  let pageToken = null;
  let page = 0;
  do {
    const url = baseUrl + (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');
    const data = sc2ApiCall(url, { headers: sc2AuthHeader(token) });
    if (data._parseError) {
      Logger.log('  ⚠️ lireContactsExistants : réponse non-JSON page ' + page + ' (HTTP ' + data._httpCode + ')');
      break;
    }
    if (data.error) {
      Logger.log('  ⚠️ lireContactsExistants : erreur API [' + data.error.code + '] : ' + data.error.message);
      break;
    }
    (data.connections || []).forEach(c => contacts.push(c));
    pageToken = data.nextPageToken || null;
  } while (pageToken && ++page < 30);
  return contacts;
}

// ==========================================
// PRÉPARATION DES GROUPES (sécurisée)
// ==========================================

function sc2PreparerGroupes(token, libellesNecessaires) {
  const urlBase = 'https://people.googleapis.com/v1/contactGroups';
  const map = {};
  let nextPageToken = null;
  do {
    const pageUrl = urlBase + (nextPageToken ? '?pageToken=' + encodeURIComponent(nextPageToken) : '');
    const data = sc2ApiCall(pageUrl, { headers: sc2AuthHeader(token) });
    if (data._parseError) {
      Logger.log('  ⚠️ sc2PreparerGroupes : réponse non-JSON (HTTP ' + data._httpCode + ')');
      return map;
    }
    if (data.error) {
      Logger.log('  ⚠️ sc2PreparerGroupes : [' + data.error.code + '] ' + data.error.message);
      return map;
    }
    (data.contactGroups || []).forEach(g => { map[g.name] = g.resourceName; });
    nextPageToken = data.nextPageToken || null;
  } while (nextPageToken);

  for (const nom of libellesNecessaires) {
    if (!nom || map[nom]) continue;
    let attempts = 0;
    while (attempts < 3) {
      const created = sc2ApiCall(urlBase, {
        method: 'post', contentType: 'application/json',
        headers: sc2AuthHeader(token),
        payload: JSON.stringify({ contactGroup: { name: nom } })
      });
      if (created._parseError) { Logger.log('  ⚠️ Groupe "' + nom + '" : réponse non-JSON'); break; }
      if (created.resourceName) { map[nom] = created.resourceName; break; }
      if (created.error && created.error.code === 429) {
        Logger.log('  ⏳ Rate limit groupe "' + nom + '" — pause 65s...');
        Utilities.sleep(65000); attempts++;
      } else if (created.error && created.error.code === 409) {
        // Conflit : groupe déjà existant (course condition) → relire
        const refetch = sc2ApiCall(urlBase, { headers: sc2AuthHeader(token) });
        if (!refetch._parseError && !refetch.error) {
          (refetch.contactGroups || []).forEach(g => { map[g.name] = g.resourceName; });
        }
        break;
      } else {
        Logger.log('  ⚠️ Groupe "' + nom + '" non créé : ' + JSON.stringify(created.error || created));
        break;
      }
    }
    Utilities.sleep(700);
  }

  // Résumé explicite des libellés dont le groupe n'a PAS pu être créé — sans ce résumé,
  // un échec de création reste noyé dans le log et le libellé manque silencieusement sur
  // TOUS les contacts concernés à chaque run (voir auditerLibellesManquants pour un
  // diagnostic détaillé, contact par contact).
  const echecs = [...libellesNecessaires].filter(l => l && !map[l]);
  if (echecs.length > 0) {
    Logger.log('  ⚠️⚠️ ' + echecs.length + ' groupe(s) N\'ONT PAS PU être créé(s) — libellé(s) manquant(s) sur les contacts concernés tant que ce n\'est pas résolu : ' + echecs.join(', '));
  }

  return map;
}

// ==========================================
// PURGE DES GROUPES (reset compteur fantôme)
// ==========================================

function sc2PurgerGroupes(token) {
  const urlBase = 'https://people.googleapis.com/v1/contactGroups';
  let groupRns = [];
  let nextPageToken = null;
  do {
    const pageUrl = urlBase + (nextPageToken ? '?pageToken=' + encodeURIComponent(nextPageToken) : '');
    const data = sc2ApiCall(pageUrl, { headers: sc2AuthHeader(token) });
    if (data._parseError || data.error) break;
    (data.contactGroups || []).forEach(g => {
      if (g.groupType === 'USER_CONTACT_GROUP') groupRns.push(g.resourceName);
    });
    nextPageToken = data.nextPageToken || null;
  } while (nextPageToken);

  if (!groupRns.length) { Logger.log('  ℹ️ Aucun groupe à purger.'); return; }
  Logger.log('  🔄 Purge de ' + groupRns.length + ' groupes (reset compteur myContacts)...');
  for (const rn of groupRns) {
    sc2ApiCall('https://people.googleapis.com/v1/' + rn + '?deleteContacts=false', {
      method: 'delete', headers: sc2AuthHeader(token)
    });
    Utilities.sleep(300);
  }
  Logger.log('  ✅ ' + groupRns.length + ' groupes purgés.');
}

// ==========================================
// NETTOYAGE DES GROUPES OBSOLÈTES (sécurisé)
// ==========================================

function sc2NettoyerGroupesObsoletes(token, libellesActifs) {
  const urlBase = 'https://people.googleapis.com/v1/contactGroups';
  let userGroups = [];
  let nextPageToken = null;
  do {
    const pageUrl = urlBase + (nextPageToken ? '?pageToken=' + encodeURIComponent(nextPageToken) : '');
    const data = sc2ApiCall(pageUrl, { headers: sc2AuthHeader(token) });
    if (data._parseError || data.error) return;
    (data.contactGroups || []).forEach(g => {
      if (g.groupType === 'USER_CONTACT_GROUP') userGroups.push(g);
    });
    nextPageToken = data.nextPageToken || null;
  } while (nextPageToken);

  const obsoletes = userGroups.filter(g => !libellesActifs.has(g.name));
  if (!obsoletes.length) return;
  Logger.log('  🗑️ Suppression de ' + obsoletes.length + ' groupe(s) obsolète(s)...');
  for (const g of obsoletes) {
    sc2ApiCall('https://people.googleapis.com/v1/' + g.resourceName + '?deleteContacts=false', {
      method: 'delete', headers: sc2AuthHeader(token)
    });
    Utilities.sleep(300);
  }
}

// ==========================================
// SYNCHRONISATION INTELLIGENTE
// ==========================================

// Logique UPSERT-only : jamais de suppression automatique.
// - Si un contact CRM correspond à un contact Google (par tél ou nom) → mise à jour
// - Si aucune correspondance → création
// - Les contacts Google sans équivalent CRM sont CONSERVÉS (jamais supprimés)
// Pour supprimer manuellement un contact, utilisez directement contacts.google.com.
function sc2SyncerContacts(token, newContacts, existants, groupesMap, deadline) {
  existants   = existants   || [];
  newContacts = newContacts || [];

  // ── Index des contacts existants ────────────────────────────────────────
  // Un contact sans téléphone (ou avec un nom légèrement différent — accents, ordre,
  // espaces) ne matchait auparavant ni par tél ni par nom exact : un DOUBLON était créé
  // à la place du contact existant, qui restait donc éternellement sans téléphone.
  // L'email — souvent plus stable que le nom — comble ce trou, et la clé de nom est
  // désormais normalisée (accents/casse/espaces) pour tolérer les petites variations
  // d'orthographe.
  //
  // IMPORTANT : les fiches de diffusion HECG_* (marqueur userDefined
  // "HECG_DIFFUSION_ID", email HECG uniquement, jamais de téléphone ni mail perso —
  // volontairement séparées de la fiche complète, voir "GROUPES DE DIFFUSION HECG_*"
  // plus bas) sont exclues de ces index. Sans ça, un étudiant sans fiche complète
  // pouvait matcher par NOM sa propre fiche de diffusion et s'y retrouver mélangé —
  // c'est ce qui causait, sur iPhone, la fiche "pro" présente mais la fiche
  // "perso + téléphone" absente.
  const phoneIdx = {}, nameIdx = {}, emailIdx = {};
  existants.forEach(e => {
    if (sc2EstContactDiffusion(e)) return;
    (e.phoneNumbers || []).forEach(p => {
      // Le tél du responsable entreprise (voir SC2_TYPE_TEL_RESPONSABLE) est dupliqué
      // sur la fiche étudiant à titre informatif : il ne doit JAMAIS servir à identifier
      // ce contact, sinon un autre étudiant du même tuteur — ou la fiche "Partenaire" du
      // tuteur lui-même — matcherait par erreur sur la fiche de ce premier étudiant.
      if (p.type === SC2_TYPE_TEL_RESPONSABLE) return;
      const k = sc2PhoneKey(p.value);
      if (k && !phoneIdx[k]) phoneIdx[k] = e;
    });
    (e.emailAddresses || []).forEach(em => {
      const k = String(em.value || '').trim().toLowerCase();
      if (k && !emailIdx[k]) emailIdx[k] = e;
    });
    const n = (e.names || [])[0];
    if (n) {
      const nk = sc2NameKey(n.familyName, n.givenName);
      if (nk && !nameIdx[nk]) nameIdx[nk] = e;
    }
  });

  // ── Tri UPSERT : créer ou mettre à jour, jamais supprimer ───────────────
  const toCreate = [], toUpdate = [];

  for (const c of newContacts) {
    const phones = c.phones || (c.phone ? [c.phone] : []);
    let existing = null;

    // Recherche par téléphone (priorité)
    for (const p of phones) {
      const k = sc2PhoneKey(p);
      if (k && phoneIdx[k]) { existing = phoneIdx[k]; break; }
    }
    // Fallback : recherche par email
    if (!existing) {
      const emails = c.emails || (c.email && c.email.includes('@') ? [c.email] : []);
      for (const em of emails) {
        const k = String(em).trim().toLowerCase();
        if (k && emailIdx[k]) { existing = emailIdx[k]; break; }
      }
    }
    // Fallback : recherche par nom+prénom (normalisé)
    if (!existing) {
      const nk = sc2NameKey(c.nom, c.prenom);
      if (nk && nameIdx[nk]) existing = nameIdx[nk];
    }

    if (existing) {
      // Politique téléphone : un numéro déjà présent côté Google n'est JAMAIS écrasé
      // par celui du CRM (on le réharmonise juste au même format) ; s'il est absent,
      // on importe celui du CRM.
      const c2 = sc2ResoudrePhonesExistant(c, existing);
      // Vérifier si une mise à jour est nécessaire (évite les appels API inutiles)
      if (sc2ContactAChange(c2, existing, groupesMap)) {
        toUpdate.push({
          contactData: c2, resourceName: existing.resourceName, etag: existing.etag,
          // Groupes que ce contact doit quitter (ex. ancienne classe après changement)
          // — voir sc2GroupesARetirer : le champ memberships de update/batchUpdate
          // n'AJOUTE pas fiablement les RETRAITS, il faut les faire explicitement.
          groupsToRemove: sc2GroupesARetirer(c2, existing, groupesMap)
        });
      }
      // Sinon : aucun changement détecté, on skip silencieusement
    } else {
      // Pas de doublon trouvé → créer
      toCreate.push(c);
    }
  }

  Logger.log('  📊 À créer : ' + toCreate.length
    + ' | À mettre à jour : ' + toUpdate.length
    + ' | Conservés sans modif : ' + (existants.length - toUpdate.length)
    + ' | (aucune suppression automatique)');

  let created = 0, updated = 0;
  if (toCreate.length > 0) created = sc2ImporterEnBatch(token, toCreate, groupesMap, deadline);
  if (toUpdate.length > 0) {
    updated = sc2MettreAJourEnBatch(token, toUpdate, groupesMap, deadline);
    // Retraits de groupe explicites — voir sc2GroupesARetirer/sc2RetirerAppartenancesGroupes.
    // Fait même si le budget a coupé sc2MettreAJourEnBatch en cours de route : les
    // contacts non traités cette fois seront de toute façon retentés (et donc retirés)
    // au prochain run puisqu'ils resteront détectés en toUpdate tant que le décalage
    // persiste ; ceux déjà traités doivent voir leur retrait appliqué maintenant.
    sc2RetirerAppartenancesGroupes(token, toUpdate, groupesMap);
  }

  return { created, updated, deleted: 0 };
}

// Calcule les groupes dont ce contact doit être RETIRÉ (groupes actuels côté Google qui
// ne correspondent plus à aucun libellé CRM attendu). Nécessaire car ni updateContact ni
// batchUpdateContacts ne retirent fiablement un contact d'un groupe via le champ
// memberships — seul l'AJOUT fonctionne de façon fiable par ce biais (constaté : un
// étudiant qui change de classe/statut gardait son ancien libellé de classe). Le retrait
// doit passer par contactGroups.members.modify (voir sc2RetirerAppartenancesGroupes).
function sc2GroupesARetirer(c, existing, groupesMap) {
  const labels = c.labels || [c.label];
  const expectedGrps = new Set(labels.filter(l => l && groupesMap[l]).map(l => groupesMap[l]));
  return (existing.memberships || [])
    .map(m => m.contactGroupMembership && m.contactGroupMembership.contactGroupResourceName)
    .filter(rn => rn && !rn.includes('contactGroups/myContacts') && !rn.includes('contactGroups/starred') && !expectedGrps.has(rn));
}

// Retire explicitement chaque contact des groupes calculés par sc2GroupesARetirer, via
// contactGroups.members.modify (regroupé par groupe, par lots de 500 pour rester dans
// des limites raisonnables — cohérent avec les autres opérations par lot du fichier).
// Log NOMINATIF (contact + nom du groupe, pas juste un total agrégé) : indispensable
// pour pouvoir auditer après coup qu'un retrait donné était légitime, en particulier
// lors du premier run suivant le déploiement de cette fonction (rattrapage potentiel de
// mois d'incohérences accumulées → volumes élevés attendus une seule fois).
function sc2RetirerAppartenancesGroupes(token, toUpdate, groupesMap) {
  const parGroupe = {};
  const nomParResourceName = {};
  toUpdate.forEach(u => {
    const cd = u.contactData || {};
    nomParResourceName[u.resourceName] = ((cd.prenom || '') + ' ' + (cd.nom || '')).trim() || cd.entreprise || u.resourceName;
    (u.groupsToRemove || []).forEach(rn => {
      if (!parGroupe[rn]) parGroupe[rn] = [];
      parGroupe[rn].push(u.resourceName);
    });
  });
  const groupes = Object.keys(parGroupe);
  if (!groupes.length) return;

  const nomParGroupeRn = {};
  Object.keys(groupesMap || {}).forEach(nom => { nomParGroupeRn[groupesMap[nom]] = nom; });

  let totalRetraits = 0;
  groupes.forEach(groupRn => {
    const membres = parGroupe[groupRn];
    const nomGroupe = nomParGroupeRn[groupRn] || groupRn;
    for (let i = 0; i < membres.length; i += 500) {
      const chunk = membres.slice(i, i + 500);
      const result = sc2ApiCall('https://people.googleapis.com/v1/' + groupRn + '/members:modify', {
        method: 'post', contentType: 'application/json',
        headers: sc2AuthHeader(token),
        payload: JSON.stringify({ resourceNamesToRemove: chunk })
      });
      if (result._parseError || result.error) {
        Logger.log('  ⚠️ Retrait groupe "' + nomGroupe + '" : '
          + (result.error ? '[' + result.error.code + '] ' + result.error.message : 'réponse non-JSON'));
      } else {
        totalRetraits += chunk.length;
        const noms = chunk.map(rn => nomParResourceName[rn] || rn);
        Logger.log('  🔻 Retiré(s) de "' + nomGroupe + '" : ' + noms.join(', '));
      }
      Utilities.sleep(300);
    }
  });
  if (totalRetraits > 0) {
    Logger.log('  🔻 ' + totalRetraits + ' retrait(s) d\'appartenance à un groupe au total (détail ci-dessus).');
  }
}

// Politique téléphone pour un contact déjà existant côté Google :
// - S'il a déjà au moins un numéro PROPRE → on le CONSERVE tel quel (jamais remplacé par
//   celui du CRM, pour ne pas écraser une correction/note manuelle), en le réharmonisant
//   juste au même format que les numéros importés (harmonisation universelle demandée).
// - S'il n'en a aucun → on importe celui du CRM (déjà harmonisé par sc2FormatPhone).
// Filtre défensif : exclut toute entrée SC2_TYPE_TEL_RESPONSABLE encore présente d'un
// ancien run (ce type n'est plus émis par sc2BuildPerson — voir son commentaire, un même
// numéro sur 2 fiches distinctes causait des confusions de nom WhatsApp/iOS). Sans ce
// filtre, un résidu serait à tort traité comme "numéro propre" du contact et ne serait
// jamais nettoyé (le remplacement complet de phoneNumbers au prochain update s'en charge).
function sc2ResoudrePhonesExistant(c, existing) {
  const existants = (existing.phoneNumbers || [])
    .filter(p => p.type !== SC2_TYPE_TEL_RESPONSABLE)
    .map(p => p.value)
    .filter(Boolean)
    .map(p => sc2FormatPhone(p) || p);
  const uniques = [...new Set(existants)];
  const phones = uniques.length > 0 ? uniques : (c.phones || (c.phone ? [c.phone] : []));
  return Object.assign({}, c, { phones });
}

// Détecte si le contact CRM diffère du contact Google existant
// Champs comparés : nom, prénom, téléphones, emails, entreprise, libellés (classes/groupes)
function sc2ContactAChange(c, existing, groupesMap) {
  // Nom + prénom
  const existName = (existing.names || [])[0] || {};
  if (String(c.nom    || '') !== String(existName.familyName || '')) return true;
  if (String(c.prenom || '') !== String(existName.givenName  || '')) return true;

  // Téléphones (ensemble non ordonné, comparaison sur chiffres uniquement). Si une
  // fiche a encore un résidu SC2_TYPE_TEL_RESPONSABLE d'un ancien run, la taille des
  // ensembles diffère ici à raison — c'est justement ce qui déclenche le nettoyage.
  const cPhones = new Set((c.phones || (c.phone ? [c.phone] : [])).map(sc2PhoneKey).filter(Boolean));
  const ePhones = new Set((existing.phoneNumbers || []).map(p => sc2PhoneKey(p.value)).filter(Boolean));
  if (cPhones.size !== ePhones.size) return true;
  for (const p of cPhones) { if (!ePhones.has(p)) return true; }

  // Emails (ensemble non ordonné, insensible à la casse)
  const cEmails = new Set((c.emails || (c.email && c.email.includes('@') ? [c.email] : []))
    .map(e => String(e).toLowerCase().trim()).filter(Boolean));
  const eEmails = new Set((existing.emailAddresses || [])
    .map(e => String(e.value || '').toLowerCase().trim()).filter(Boolean));
  if (cEmails.size !== eEmails.size) return true;
  for (const e of cEmails) { if (!eEmails.has(e)) return true; }

  // Entreprise (organisation principale)
  const cOrg = String(c.entreprise || '').trim();
  const eOrg = String(((existing.organizations || [])[0] || {}).name || '').trim();
  if (cOrg !== eOrg) return true;

  // Adresse entreprise (fiche MA uniquement — voir c.adresseEnt dans sc2BuildPerson)
  const cAddr = c.adresseEnt || {};
  const eAddr = (existing.addresses || [])[0] || {};
  if (String(cAddr.adresse || '') !== String(eAddr.streetAddress || '')) return true;
  if (String(cAddr.cp      || '') !== String(eAddr.postalCode   || '')) return true;
  if (String(cAddr.ville   || '') !== String(eAddr.city         || '')) return true;

  // Le champ relations (RELATED, ex-tuteur) n'est plus utilisé par sc2BuildPerson
  // (remplacé par un champ personnalisé — candidat plausible au bug iPhone "fiche
  // visible mais téléphone manquant"). Une fiche qui en a encore un depuis un ancien
  // sync doit être mise à jour pour que "relations" (dans l'updateMask) soit nettoyé.
  if (existing.relations && existing.relations.length > 0) return true;

  // Libellés / groupes (classes, statuts…)
  const labels = c.labels || [c.label];
  const expectedGrps = new Set(labels.filter(l => l && groupesMap[l]).map(l => groupesMap[l]));
  const actualGrps   = new Set(
    (existing.memberships || [])
      .map(m => m.contactGroupMembership && m.contactGroupMembership.contactGroupResourceName)
      .filter(rn => rn && !rn.includes('contactGroups/myContacts') && !rn.includes('contactGroups/starred'))
  );
  if (expectedGrps.size !== actualGrps.size) return true;
  for (const g of expectedGrps) { if (!actualGrps.has(g)) return true; }

  return false;
}

// ==========================================
// CRÉATION EN BATCH
// ==========================================

// Utilise batchCreateContacts (200 contacts/requête) au lieu de createContact (1/requête)
// → ~200x plus rapide : 1686 contacts ≈ 9 requêtes ≈ 20 secondes (au lieu de 40 min)
function sc2ImporterEnBatch(token, contacts, groupesMap, deadline) {
  const BATCH_SIZE = 200;
  let total = 0;

  for (let i = 0; i < contacts.length; i += BATCH_SIZE) {
    if (deadline && Date.now() > deadline) {
      Logger.log('  ⏰ Budget utilisateur atteint — arrêt création après ' + total + ' contacts ('
        + (contacts.length - i) + ' restants, repris au prochain run).');
      break;
    }

    const lot = contacts.slice(i, i + BATCH_SIZE);

    // Construire les personnes du lot (filtrer celles sans groupe valide)
    const contactsPayload = [];
    for (const c of lot) {
      const person = sc2BuildPerson(c, groupesMap);
      if (person) contactsPayload.push({ contactPerson: person });
    }
    if (contactsPayload.length === 0) continue;

    const numLot = Math.floor(i / BATCH_SIZE) + 1;
    const result = sc2ApiCall('https://people.googleapis.com/v1/people:batchCreateContacts', {
      method: 'post', contentType: 'application/json',
      headers: sc2AuthHeader(token),
      payload: JSON.stringify({ contacts: contactsPayload, readMask: 'names' })
    });

    if (result._parseError) {
      Logger.log('  ⚠️ Lot ' + numLot + ' : réponse non-JSON — ignoré.');
    } else if (result.error) {
      const code = result.error.code, msg = result.error.message || '';
      Logger.log('  ⚠️ Lot ' + numLot + ' [' + code + '] : ' + msg);

      if (code === 429 && (msg.includes('OVERFLOW') || msg.includes('MY_CONTACTS'))) {
        Logger.log('  🚫 MY_CONTACTS_OVERFLOW — corbeille non vidée pour cet utilisateur. Arrêt.');
        break;
      }
      if (code === 429) {
        Logger.log('  ⏳ Rate limit — pause 65s puis réessai...');
        Utilities.sleep(65000);
        i -= BATCH_SIZE; // réessayer ce lot
        continue;
      }
    } else {
      const created = (result.createdPeople || []).length;
      total += created;
      Logger.log('  ✅ Lot ' + numLot + '/' + Math.ceil(contacts.length / BATCH_SIZE)
        + ' : ' + created + ' créés (total : ' + total + ')');
    }

    Utilities.sleep(1000); // 1 s entre les lots pour éviter le rate limiting
  }
  return total;
}

// ==========================================
// MISE À JOUR EN BATCH (avec retry 429)
// ==========================================

// Utilise batchUpdateContacts (200 contacts/requête) au lieu d'un PATCH updateContact
// par contact → même gain que sc2ImporterEnBatch (~200x plus rapide). C'est ce qui
// permet aux 8 utilisateurs de tenir dans un seul run au lieu d'étaler les màj sur
// plusieurs jours (5 min de budget/utilisateur ne suffisaient qu'à ~85 patchs individuels).
//
// L'etag lu au moment de sc2LireContactsExistants est renvoyé sur chaque contact
// (Google l'exige désormais : "Request must set person.etag ... for each source that
// is being updated", sinon TOUT le batch — et même le repli individuel — échoue en
// FAILED_PRECONDITION). Un etag périmé isolé (fréquent — Google le fait tourner en
// interne même sans édition manuelle) ne doit faire échouer que CE contact : le repli
// en PATCH individuel (sc2MettreAJourIndividuel) relit alors un etag frais et retente.

// Récupère l'etag ACTUEL d'un lot de resourceNames via people:batchGet (GET avec
// resourceNames répétés en query string) au lieu d'un GET par contact. people:batchGet
// accepte jusqu'à 200 resourceNames côté API, mais UrlFetchApp rejette une URL aussi
// longue ("Limit Exceeded: URLFetch URL Length") — on découpe donc en petits paquets
// de resourceNames par requête, ce qui reste bien plus rapide que 200 GET individuels.
const SC2_BATCHGET_CHUNK = 40;

function sc2RafraichirEtags(token, resourceNames) {
  const map = {};
  if (!resourceNames || !resourceNames.length) return map;
  for (let i = 0; i < resourceNames.length; i += SC2_BATCHGET_CHUNK) {
    const chunk = resourceNames.slice(i, i + SC2_BATCHGET_CHUNK);
    const qs = chunk.map(rn => 'resourceNames=' + encodeURIComponent(rn)).join('&');
    const data = sc2ApiCall('https://people.googleapis.com/v1/people:batchGet?' + qs + '&personFields=names',
      { headers: sc2AuthHeader(token) });
    if (data._parseError || data.error || !data.responses) continue; // ce paquet échoue, on tente les suivants
    data.responses.forEach(r => {
      if (r.person && r.person.resourceName && r.person.etag) map[r.person.resourceName] = r.person.etag;
    });
  }
  return map;
}

function sc2MettreAJourEnBatch(token, updates, groupesMap, deadline) {
  const FIELDS = 'names,phoneNumbers,emailAddresses,organizations,addresses,memberships,userDefined,relations';
  const BATCH_SIZE = 200;
  let count = 0;

  for (let i = 0; i < updates.length; i += BATCH_SIZE) {
    if (deadline && Date.now() > deadline) {
      Logger.log('  ⏰ Budget utilisateur atteint — arrêt mise à jour après ' + count + ' contacts ('
        + (updates.length - i) + ' restants, repris au prochain run).');
      break;
    }

    const lot = updates.slice(i, i + BATCH_SIZE);
    const contactsMap = {};
    for (const u of lot) {
      const person = sc2BuildPerson(u.contactData, groupesMap);
      if (!person) continue;
      if (u.etag) person.etag = u.etag;
      contactsMap[u.resourceName] = person;
    }
    if (Object.keys(contactsMap).length === 0) continue;

    const numLot = Math.floor(i / BATCH_SIZE) + 1;
    const result = sc2ApiCall('https://people.googleapis.com/v1/people:batchUpdateContacts', {
      method: 'post', contentType: 'application/json',
      headers: sc2AuthHeader(token),
      payload: JSON.stringify({ contacts: contactsMap, updateMask: FIELDS, readMask: 'names' })
    });

    if (result._parseError) {
      Logger.log('  ⚠️ Lot màj ' + numLot + ' : réponse non-JSON — repli en individuel.');
      count += sc2MettreAJourIndividuel(token, contactsMap, FIELDS, deadline);
    } else if (result.error) {
      const code = result.error.code, msg = result.error.message || '';
      if (code === 429) {
        Logger.log('  ⚠️ Lot màj ' + numLot + ' [' + code + '] : ' + msg);
        Logger.log('  ⏳ Rate limit — pause 65s puis réessai...');
        Utilities.sleep(65000);
        i -= BATCH_SIZE; // réessayer ce lot
        continue;
      }
      // Étag périmé (fréquent — Google le fait tourner en interne même sans édition
      // manuelle, et c'est désormais obligatoire sur chaque écriture) : ça rejette TOUT
      // le lot de 200 d'un coup. Avant de basculer sur le repli individuel (1 requête +
      // 700ms/contact — peut à lui seul consommer le budget de 5 min/utilisateur), on
      // rafraîchit les etags de tout le lot en UN SEUL appel (batchGet) et on retente le
      // batch une fois. Ne laisse au repli individuel que les vrais échecs résiduels.
      if (/etag/i.test(msg)) {
        const freshEtags = sc2RafraichirEtags(token, Object.keys(contactsMap));
        let refreshed = 0;
        Object.keys(contactsMap).forEach(rn => { if (freshEtags[rn]) { contactsMap[rn].etag = freshEtags[rn]; refreshed++; } });

        if (refreshed > 0) {
          const retry = sc2ApiCall('https://people.googleapis.com/v1/people:batchUpdateContacts', {
            method: 'post', contentType: 'application/json',
            headers: sc2AuthHeader(token),
            payload: JSON.stringify({ contacts: contactsMap, updateMask: FIELDS, readMask: 'names' })
          });
          if (!retry._parseError && !retry.error) {
            const results = retry.updateResult || {};
            let lotOk = 0; const echecs = {};
            Object.keys(results).forEach(rn => {
              const pr = results[rn];
              if (pr && pr.person) lotOk++; else echecs[rn] = contactsMap[rn];
            });
            count += lotOk;
            Logger.log('  ✅ Lot màj ' + numLot + ' (après rafraîchissement etag, ' + refreshed + '/' + Object.keys(contactsMap).length + ') : '
              + lotOk + ' mis à jour' + (Object.keys(echecs).length ? ', ' + Object.keys(echecs).length + ' échec(s) résiduel(s) — repli individuel' : ''));
            if (Object.keys(echecs).length > 0) count += sc2MettreAJourIndividuel(token, echecs, FIELDS, deadline);
            Utilities.sleep(1000);
            continue;
          }
        }
        Logger.log('  ⚠️ Lot màj ' + numLot + ' : rafraîchissement etag sans effet — repli en individuel.');
      }
      // Échec transactionnel du lot entier (pas de détail par contact fourni par
      // Google dans ce cas). On bascule en PATCH individuel pour ce lot afin de ne
      // pas perdre les ~199 autres mises à jour valides.
      Logger.log('  ⚠️ Lot màj ' + numLot + ' [' + code + '] : ' + msg + ' — repli en individuel.');
      count += sc2MettreAJourIndividuel(token, contactsMap, FIELDS, deadline);
    } else {
      const results = result.updateResult || {};
      let lotOk = 0, lotErr = 0;
      Object.keys(results).forEach(rn => {
        const pr = results[rn];
        if (pr && pr.person) {
          lotOk++;
        } else {
          lotErr++;
          const st = pr && pr.status;
          Logger.log('  ⚠️ Màj ' + rn + ' : ' + (st ? '[' + st.code + '] ' + st.message : 'échec'));
        }
      });
      count += lotOk;
      Logger.log('  ✅ Lot màj ' + numLot + '/' + Math.ceil(updates.length / BATCH_SIZE)
        + ' : ' + lotOk + ' mis à jour' + (lotErr ? ', ' + lotErr + ' erreur(s)' : '')
        + ' (total : ' + count + ')');
    }

    Utilities.sleep(1000); // 1 s entre les lots pour éviter le rate limiting
  }
  return count;
}

// Repli PATCH individuel (:updateContact) pour un lot dont le batchUpdateContacts
// a échoué en bloc. Plus lent mais isole le(s) contact(s) fautif(s) sans sacrifier
// les ~199 autres mises à jour valides du lot.
function sc2MettreAJourIndividuel(token, contactsMap, fields, deadline) {
  let count = 0;
  const resourceNames = Object.keys(contactsMap);
  for (let j = 0; j < resourceNames.length; j++) {
    if (deadline && Date.now() > deadline) {
      Logger.log('  ⏰ Budget utilisateur atteint pendant le repli individuel — arrêt après ' + count
        + ' contacts (' + (resourceNames.length - j) + ' restants de ce lot, repris au prochain run).');
      break;
    }
    const resourceName = resourceNames[j];
    let person = contactsMap[resourceName];
    const url = 'https://people.googleapis.com/v1/' + resourceName + ':updateContact'
      + '?updatePersonFields=' + encodeURIComponent(fields);

    let retriedRate = false, retriedEtag = false;
    while (true) {
      const result = sc2ApiCall(url, {
        method: 'patch', contentType: 'application/json',
        headers: sc2AuthHeader(token),
        payload: JSON.stringify(person)
      });
      if (result._parseError) {
        Logger.log('  ⚠️ Màj individuelle ' + resourceName + ' : réponse non-JSON'); break;
      }
      if (result.error) {
        const code = result.error.code, msg = result.error.message || '';
        if (code === 429 && !retriedRate) {
          Logger.log('  ⏳ Rate limit màj individuelle — pause 65s...');
          Utilities.sleep(65000); retriedRate = true; continue;
        }
        // Etag périmé (ou manquant) : on relit le contact pour récupérer l'etag à jour
        // et on retente une seule fois avec ce nouvel etag.
        if (!retriedEtag && (code === 400 || code === 409) && /etag/i.test(msg)) {
          const fresh = sc2ApiCall('https://people.googleapis.com/v1/' + resourceName + '?personFields=names',
            { headers: sc2AuthHeader(token) });
          retriedEtag = true;
          if (!fresh._parseError && !fresh.error && fresh.etag) {
            person = Object.assign({}, person, { etag: fresh.etag });
            continue;
          }
        }
        Logger.log('  ⚠️ Erreur màj individuelle ' + resourceName + ' [' + code + '] : ' + msg);
      } else { count++; }
      break;
    }
    Utilities.sleep(700);
  }
  return count;
}

// ==========================================
// CONSTRUCTION D'UN CONTACT PEOPLE API
// ==========================================

function sc2BuildPerson(c, groupesMap) {
  const phoneList = c.phones || (c.phone ? [c.phone] : []);
  const emailList = c.emails || (c.email && c.email.includes('@') ? [c.email] : []);
  const labels = c.labels || [c.label];
  const memberships = labels
    .filter(l => l && groupesMap[l])
    .map(l => ({ contactGroupMembership: { contactGroupResourceName: groupesMap[l] } }));

  if (!memberships.length) return null;

  const person = {
    names:       [{ givenName: c.prenom || '', familyName: c.nom || '' }],
    memberships: memberships
  };
  const phoneEntries = phoneList.map((p, i) => ({ value: p, type: i === 0 ? 'mobile' : 'work' }));
  // Le numéro du responsable entreprise (tuteur) n'est PAS dupliqué ici (voir extra
  // ci-dessous pour son nom/entreprise). Essayé un temps comme 2e phoneNumbers typé,
  // mais un même numéro dialable présent sur DEUX fiches Google Contacts distinctes
  // (celle du MA lui-même — sc2ExtractPartenaires — et celle de l'étudiant) fait que
  // iOS/WhatsApp résolvent l'identité de l'appelant de façon incohérente pour ce numéro
  // (tantôt le nom de l'étudiant, tantôt celui du MA) : ambiguïté structurelle, pas un
  // problème de libellé. Le numéro du MA ne vit donc que sur SA fiche dédiée.
  if (phoneEntries.length > 0) person.phoneNumbers = phoneEntries;
  if (emailList.length > 0)
    person.emailAddresses = emailList.map(e => ({ value: e, type: 'work' }));
  if (c.entreprise)
    person.organizations = [{ name: c.entreprise, type: 'work' }];
  // Adresse de l'entreprise, sur la fiche du MA uniquement (sc2ExtractPartenaires).
  if (c.adresseEnt) {
    const a = c.adresseEnt;
    person.addresses = [{ streetAddress: a.adresse || '', postalCode: a.cp || '', city: a.ville || '', type: 'work' }];
  }

  const extra = [];
  if (c.contactEnt) {
    // Nom du tuteur en champ personnalisé, informatif seulement — jamais son téléphone
    // (voir commentaire au-dessus de phoneEntries). Pour l'appeler, on cherche sa fiche
    // dédiée par son nom (sc2ExtractPartenaires). Pas d'entreprise ici : déjà sur
    // person.organizations (c.entreprise) juste au-dessus — inutile de la dupliquer.
    const ce = c.contactEnt;
    if (ce.nomComplet) extra.push({ key: 'Tuteur entreprise', value: ce.nomComplet });
  }
  if (c.etudiants && c.etudiants.length > 0)
    extra.push({ key: 'Étudiants HECG', value: c.etudiants.join(', ') });
  if (c.diffusionKey)
    extra.push({ key: 'HECG_DIFFUSION_ID', value: c.diffusionKey });
  if (extra.length > 0) person.userDefined = extra;

  return person;
}

// ==========================================
// EXTRACTION DES FEUILLES (identique V1)
// ==========================================

function sc2FormatPhone(phoneRaw) {
  if (!phoneRaw) return '';
  let cleaned = String(phoneRaw).trim().replace(/[^\d+]/g, '');
  let digits9 = null;
  if (cleaned.startsWith('+33') || cleaned.startsWith('0033')) {
    // Tolère l'erreur de saisie fréquente "+33 06 12 34 56 78" (0 conservé après l'indicatif)
    let rest = cleaned.startsWith('+33') ? cleaned.substring(3) : cleaned.substring(4);
    if (rest.length === 10 && rest.startsWith('0')) rest = rest.substring(1);
    if (rest.length === 9) digits9 = rest;
  }
  else if (cleaned.startsWith('0') && cleaned.length === 10)                digits9 = cleaned.substring(1);
  else if (/^\d{9}$/.test(cleaned))                                         digits9 = cleaned;
  if (digits9) return '+33 ' + digits9[0] + ' ' + digits9.slice(1,3) + ' ' + digits9.slice(3,5) + ' ' + digits9.slice(5,7) + ' ' + digits9.slice(7,9);
  if (!cleaned.startsWith('+')) cleaned = '+' + cleaned;
  return cleaned.length > 8 ? cleaned : '';
}

function sc2PhoneKey(phone) { return String(phone || '').replace(/\D/g, ''); }

// Normalise une valeur de cellule destinée à devenir un libellé/groupe (classe, statut,
// campus...) : espace insécable → espace normal, espaces multiples/tabulations →
// simple espace, trim. Sans ça, une classe saisie deux fois avec un espace "en trop"
// (fréquent en copier-coller depuis un autre outil) crée un DEUXIÈME groupe Google
// quasi-identique, et une partie des étudiants de cette classe se retrouve dans le
// mauvais groupe (ou sans le groupe qu'ont leurs camarades) — symptôme observé de
// libellés manquants sur certains contacts alors qu'ils sont présents sur d'autres.
function sc2NormLabel(s) {
  return String(s || '').replace(/ /g, ' ').replace(/\s+/g, ' ').trim();
}

// Un contact Google est une fiche de diffusion HECG_* (email HECG seul, jamais de
// téléphone ni mail perso — voir "GROUPES DE DIFFUSION HECG_*" plus bas) s'il porte le
// marqueur userDefined "HECG_DIFFUSION_ID". Ces fiches ne doivent jamais être prises
// pour la fiche complète (perso + téléphone) d'un étudiant/formateur, même si le nom
// correspond — c'est pour ça que sc2SyncerContacts les exclut de ses index de
// correspondance (phone/email/nom).
function sc2EstContactDiffusion(e) {
  return (e.userDefined || []).some(u => u.key === 'HECG_DIFFUSION_ID');
}

// Clé de correspondance nom+prénom tolérante aux accents/casse/espaces multiples
// (ex. "Béatrice" / "beatrice", "Jean-Paul" / "jean paul") pour retrouver un contact
// déjà existant même si son nom n'est pas orthographié à l'identique côté Google.
function sc2NameKey(familyName, givenName) {
  function norm(s) {
    return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '')
      .toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
  }
  const f = norm(familyName), g = norm(givenName);
  if (!f && !g) return '';
  return f + '|' + g;
}

function sc2BuildPartMap(ss) {
  const data = getSheetSafe(ss, 'Partenariat').getDataRange().getValues();
  const map  = {};
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0] || '').trim();
    if (!id) continue;
    map[id] = {
      nomComplet: [String(data[i][4]||'').trim(), String(data[i][3]||'').trim()].filter(Boolean).join(' '),
      tel: sc2FormatPhone(data[i][5]) || sc2FormatPhone(data[i][9]),
      nomEnt: String(data[i][2] || '').trim()
    };
  }
  return map;
}

// Classification centralisée du statut "typo" (colonne U de ETUDIANTS, ex. "Inscrit",
// "Préinscrit", "Sorti", "Alternance"…) en un des 4 libellés principaux. Centralisée
// ici pour que sc2ExtractEtudiants ET sc2BuildMAClassesParContact (libellés MA-<classe>
// des maîtres d'apprentissage) restent TOUJOURS d'accord sur qui compte comme
// Inscrit/Préinscrit — sinon un MA pourrait se voir attribuer (ou perdre) un libellé
// MA-<classe> de façon incohérente avec le libellé de l'étudiant lui-même.
function sc2ClassifierTypoEtudiant(typoLower) {
  if (typoLower.includes('prein') || typoLower.includes('préin')) return 'Étudiant Préinscrit';
  if (typoLower.includes('alumni'))                                return 'Alumni';   // ancien étudiant formé (fin d'études)
  if (typoLower.includes('sorti'))                                 return 'Étudiant Sorti';
  if (typoLower.includes('inscrit') || typoLower.includes('alternance')) return 'Étudiant Inscrit';
  return 'Prospect';
}

function sc2BuildEtudiantsParContact(ss) {
  const data = getSheetSafe(ss, 'ETUDIANTS').getDataRange().getValues();
  const map  = {};
  for (let i = 1; i < data.length; i++) {
    const idCont = String(data[i][12] || '').trim();
    if (!idCont) continue;
    const nom = [String(data[i][3]||'').trim(), String(data[i][2]||'').trim()].filter(Boolean).join(' ');
    if (!nom) continue;
    if (!map[idCont]) map[idCont] = [];
    map[idCont].push(nom);
  }
  return map;
}

// Pour chaque MA (contact de la feuille Partenariat, identifié par idCont), l'ensemble
// des classes de ses étudiants ACTUELLEMENT Inscrits ou Préinscrits (jamais Sortis ni
// Prospects) — sert à générer les libellés "MA-<classe>" sur la fiche du MA (voir
// sc2ExtractPartenaires). Recalculé à chaque synchro depuis ETUDIANTS (jamais mémorisé),
// donc un changement de classe/statut d'un étudiant met bien à jour les libellés MA-*
// de son maître d'apprentissage au run suivant.
function sc2BuildMAClassesParContact(ss) {
  const data = getSheetSafe(ss, 'ETUDIANTS').getDataRange().getValues();
  const map = {};
  for (let i = 1; i < data.length; i++) {
    const idCont = String(data[i][12] || '').trim();
    const classe = sc2NormLabel(data[i][14]);
    if (!idCont || !classe) continue;
    const typo = String(data[i][20] || '').toLowerCase();
    const mainLabel = sc2ClassifierTypoEtudiant(typo);
    if (mainLabel !== 'Étudiant Inscrit' && mainLabel !== 'Étudiant Préinscrit') continue;
    if (!map[idCont]) map[idCont] = new Set();
    map[idCont].add(classe);
  }
  return map;
}

function sc2ExtractEtudiants(ss) {
  const data    = getSheetSafe(ss, 'ETUDIANTS').getDataRange().getValues();
  const partMap = sc2BuildPartMap(ss);
  const list    = [];
  for (let i = 1; i < data.length; i++) {
    const nom    = String(data[i][2]  || '').trim();
    const prenom = String(data[i][3]  || '').trim();
    const mail   = String(data[i][4]  || '').trim();
    const tel    = sc2FormatPhone(data[i][5]);
    const typo   = String(data[i][20] || '').toLowerCase();
    const statut = sc2NormLabel(data[i][13]);
    const classe = sc2NormLabel(data[i][14]);
    const campus = sc2NormLabel(data[i][22]);
    const idCont = String(data[i][12] || '').trim();
    const nomEnt = String(data[i][21] || '').trim();
    if (!prenom && !nom) continue;
    const mainLabel = sc2ClassifierTypoEtudiant(typo);
    const labelsSet = new Set([mainLabel]);
    const isSorti = mainLabel === 'Étudiant Sorti' || mainLabel === 'Alumni';
    if (mainLabel !== 'Prospect') {
      if (isSorti) { if (statut.toLowerCase().includes('contrat') && statut) labelsSet.add(statut); }
      else { if (statut) labelsSet.add(statut); if (classe) labelsSet.add(classe); if (campus) labelsSet.add(campus); }
    }
    // Fiche complète : UNIQUEMENT l'adresse perso. L'adresse HECG (@hecg.fr) vit
    // exclusivement sur la fiche de diffusion HECG_<classe> séparée (voir
    // sc2ExtractGroupesDiffusion) — jamais les deux sur la même fiche, sinon Gmail
    // affiche les deux adresses dès qu'on tape N'IMPORTE LEQUEL des deux libellés
    // (classe ou HECG_<classe>), quelle que soit celle marquée "primaire".
    list.push({
      nom, prenom, phones: tel ? [tel] : [], emails: (mail && mail.includes('@')) ? [mail] : [],
      labels: [...labelsSet], entreprise: nomEnt, contactEnt: partMap[idCont] || null
    });
  }
  return list;
}

function sc2ExtractPartenaires(ss) {
  const data           = getSheetSafe(ss, 'Partenariat').getDataRange().getValues();
  const etudiantsMap   = sc2BuildEtudiantsParContact(ss);
  const maClassesMap   = sc2BuildMAClassesParContact(ss);
  const list           = [];
  for (let i = 1; i < data.length; i++) {
    const idContact  = String(data[i][0] || '').trim();
    const entreprise = String(data[i][2] || '').trim();
    const nom        = String(data[i][3] || '').trim();
    const prenom     = String(data[i][4] || '').trim();
    const tel1       = sc2FormatPhone(data[i][5]);
    const email1     = String(data[i][6] || '').trim();
    const email2     = String(data[i][8] || '').trim();
    const tel2       = sc2FormatPhone(data[i][9]);
    const adresse    = String(data[i][10] || '').trim();
    const cp         = String(data[i][11] || '').trim();
    const ville      = String(data[i][12] || '').trim();
    if (!entreprise && !nom && !prenom) continue;
    const phones = []; if (tel1) phones.push(tel1); if (tel2 && tel2 !== tel1) phones.push(tel2);
    const emails = []; if (email1.includes('@')) emails.push(email1); if (email2.includes('@') && email2 !== email1) emails.push(email2);
    // MA-<classe> : un libellé par classe où ce contact a au moins un étudiant
    // actuellement Inscrit ou Préinscrit (voir sc2BuildMAClassesParContact). Recalculé
    // à chaque synchro, donc se met à jour automatiquement (ajout/retrait de classe)
    // au fil des changements de statut/classe dans ETUDIANTS.
    const maLabels = [...(maClassesMap[idContact] || [])].map(cl => 'MA-' + cl);
    list.push({
      nom, prenom, phones, emails, labels: ['Entreprise', ...maLabels], entreprise,
      adresseEnt: (adresse || cp || ville) ? { adresse, cp, ville } : null,
      etudiants: etudiantsMap[idContact] || []
    });
  }
  return list;
}

function sc2ExtractPersonnel(ss) {
  var ws;
  try { ws = getSheetSafe(ss, 'Adultes'); } catch(e) { ws = getSheetSafe(ss, 'FORMATEURS'); }
  var data = ws.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var nom    = String(data[i][1]  || '').trim();
    var prenom = String(data[i][2]  || '').trim();
    if (!nom && !prenom) continue;
    var mailPerso = String(data[i][3]  || '').trim();
    var mailPro   = String(data[i][4]  || '').trim();
    var dateSortie = data[i][10];
    var tel        = sc2FormatPhone(data[i][11]);
    var etat       = String(data[i][12] || 'NON').trim().toUpperCase();
    var estSorti   = dateSortie instanceof Date ? !isNaN(dateSortie.getTime()) : String(dateSortie || '').trim() !== '';
    var label = (etat === 'OUI')
      ? (estSorti ? 'Formateurs sortis' : 'Formateurs')
      : (estSorti ? 'Encadrants sortis' : 'Encadrants');
    var emails = [];
    if (mailPro.includes('@'))                          emails.push(mailPro);
    if (mailPerso.includes('@') && mailPerso !== mailPro) emails.push(mailPerso);
    list.push({ nom, prenom, phones: tel ? [tel] : [], emails, label });
  }
  return list;
}

// ==========================================
// GROUPES DE DIFFUSION HECG_* (listes de diffusion par classe / formateurs)
// ==========================================
// Séparation stricte et volontaire, confirmée après test réel : Gmail expose TOUTES
// les adresses mail d'un contact dès qu'on tape N'IMPORTE QUEL libellé auquel il
// appartient, y compris quand une adresse est marquée "primaire" — il n'existe aucun
// moyen de scoper une adresse à un libellé précis sur une fiche unique. Pour garantir
// que :
//   • un libellé "HECG_<classe>" (ou "HECG_Formateurs") ne renvoie QUE l'adresse
//     @hecg.fr,
//   • un libellé "classe" (ou "Formateurs"/"Encadrants") ne renvoie QUE l'adresse
//     perso (côté étudiants — pour le personnel la fiche complète a toujours les 2),
// on maintient deux fiches Google Contacts distinctes par étudiant :
//   • la fiche COMPLÈTE (sc2ExtractEtudiants) — nom, tél, mail perso, entreprise,
//     tuteur... TOUTES les infos SAUF le mail HECG — rattachée aux libellés "classe".
//   • la fiche de DIFFUSION (ci-dessous) — nom + UNIQUEMENT l'adresse mail HECG,
//     jamais le téléphone ni le mail perso — rattachée au libellé "HECG_<classe>".
//
// Ces fiches de diffusion sont identifiées par un marqueur stable (userDefined
// "HECG_DIFFUSION_ID") plutôt que par téléphone/nom, pour ne jamais fusionner avec la
// fiche complète du même étudiant/formateur — voir aussi sc2EstContactDiffusion(),
// utilisée dans sc2SyncerContacts pour exclure ces fiches de ses index de
// correspondance (sinon la fiche complète peut s'y raccrocher par nom et s'y faire
// fusionner puis vider par cette synchro-ci, qui la ramènerait à son état minimal).

function sc2ExtractGroupesDiffusion(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const list = [];

  // Étudiants actifs : un membre par classe, mail = uniquement colonne AA "Mail HECG"
  const dataEtu = getSheetSafe(ss, 'ETUDIANTS').getDataRange().getValues();
  for (let i = 1; i < dataEtu.length; i++) {
    const id       = String(dataEtu[i][0]  || '').trim();
    const nom      = String(dataEtu[i][2]  || '').trim();
    const prenom   = String(dataEtu[i][3]  || '').trim();
    const typo     = String(dataEtu[i][20] || '').toLowerCase();
    const classe   = sc2NormLabel(dataEtu[i][14]);
    const mailHecg = String(dataEtu[i][26] || '').trim();
    if (!id || !classe || !mailHecg.includes('@')) continue;
    // Inclut Inscrit + Préinscrit + Alternance ; exclut uniquement les étudiants Sortis.
    const estActif = typo.includes('inscrit') || typo.includes('alternance');
    const estSorti = typo.includes('sorti');
    if (!estActif || estSorti) continue;
    list.push({ nom, prenom, emails: [mailHecg], labels: ['HECG_' + classe], diffusionKey: 'ETU_' + id });
  }

  // Formateurs actifs : un seul groupe HECG_Formateurs, mail = uniquement mailPro (colonne E, domaine HECG)
  let wsAdultes;
  try { wsAdultes = getSheetSafe(ss, 'Adultes'); } catch (e) { wsAdultes = getSheetSafe(ss, 'FORMATEURS'); }
  const dataAdultes = wsAdultes.getDataRange().getValues();
  for (let i = 1; i < dataAdultes.length; i++) {
    const id = String(dataAdultes[i][0] || '').trim();
    if (!id) continue;
    const nom        = String(dataAdultes[i][1]  || '').trim();
    const prenom     = String(dataAdultes[i][2]  || '').trim();
    const mailPro    = String(dataAdultes[i][4]  || '').trim();
    const dateSortie = dataAdultes[i][10];
    const etat       = String(dataAdultes[i][12] || 'NON').trim().toUpperCase();
    const estSorti   = dateSortie instanceof Date ? !isNaN(dateSortie.getTime()) : String(dateSortie || '').trim() !== '';
    if (etat !== 'OUI' || estSorti || !mailPro.includes('@')) continue;
    list.push({ nom, prenom, emails: [mailPro], labels: ['HECG_Formateurs'], diffusionKey: 'FORM_' + id });
  }

  return list;
}

// Groupes FORM-<classe> : un membre par (formateur, classe où il/elle intervient),
// mail = UNIQUEMENT l'adresse pro (@hecg.fr) — jamais le perso, même logique que les
// groupes HECG_* ci-dessus (Gmail expose toutes les adresses d'un contact dès qu'on
// tape un de ses libellés, donc la pureté du groupe exige une fiche dédiée par email).
// Classes = colonne "Classes" (O) de la feuille Adultes/FORMATEURS, renseignée via les
// cases à cocher de vie_scolaire.html (Personnel encadrant) — toujours les mêmes noms
// que la colonne classe de ETUDIANTS, jamais de nomenclature séparée.
// Ne couvre que les FORMATEURS (etatFormateur = OUI), pas les encadrants.
function sc2ExtractGroupesFormateursParClasse(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const list = [];
  let ws;
  try { ws = getSheetSafe(ss, 'Adultes'); } catch (e) { ws = getSheetSafe(ss, 'FORMATEURS'); }
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0] || '').trim();
    if (!id) continue;
    const nom        = String(data[i][1]  || '').trim();
    const prenom     = String(data[i][2]  || '').trim();
    const mailPro    = String(data[i][4]  || '').trim();
    const dateSortie = data[i][10];
    const etat       = String(data[i][12] || 'NON').trim().toUpperCase();
    const estSorti   = dateSortie instanceof Date ? !isNaN(dateSortie.getTime()) : String(dateSortie || '').trim() !== '';
    const classes    = String(data[i][14] || '').split(',').map(c => sc2NormLabel(c)).filter(Boolean);
    if (etat !== 'OUI' || estSorti || !mailPro.includes('@') || !classes.length) continue;
    classes.forEach(classe => {
      list.push({ nom, prenom, emails: [mailPro], labels: ['FORM-' + classe], diffusionKey: 'FORMCLASSE_' + id + '_' + classe });
    });
  }
  return list;
}

// Synchronise les membres des groupes de diffusion HECG_* pour un utilisateur donné.
// Appariement strict par marqueur userDefined "HECG_DIFFUSION_ID" (jamais par nom/téléphone)
// pour ne jamais toucher aux fiches complètes existantes.
function sc2SyncerDiffusion(token, members, existants, groupesMap, deadline) {
  existants = existants || [];
  members   = members   || [];

  const keyIdx = {};
  existants.forEach(e => {
    (e.userDefined || []).forEach(u => {
      if (u.key === 'HECG_DIFFUSION_ID' && u.value && !keyIdx[u.value]) keyIdx[u.value] = e;
    });
  });

  const toCreate = [], toUpdate = [];
  for (const m of members) {
    const existing = keyIdx[m.diffusionKey];
    if (existing) {
      if (sc2ContactAChange(m, existing, groupesMap)) {
        // Un étudiant qui change de classe garde le MÊME diffusionKey ('ETU_'+id) — sa
        // fiche HECG_* existante doit donc basculer de HECG_<ancienneClasse> vers
        // HECG_<nouvelleClasse>, pas être recréée. Même limitation API que la fiche
        // complète (voir sc2GroupesARetirer) : le retrait de l'ancien groupe doit être
        // fait explicitement.
        toUpdate.push({
          contactData: m, resourceName: existing.resourceName, etag: existing.etag,
          groupsToRemove: sc2GroupesARetirer(m, existing, groupesMap)
        });
      }
    } else {
      toCreate.push(m);
    }
  }

  Logger.log('  📋 Diffusion HECG_* : ' + toCreate.length + ' à créer, ' + toUpdate.length
    + ' à mettre à jour, ' + (members.length - toCreate.length - toUpdate.length) + ' inchangé(s).');

  let created = 0, updated = 0;
  if (toCreate.length > 0) created = sc2ImporterEnBatch(token, toCreate, groupesMap, deadline);
  if (toUpdate.length > 0) {
    updated = sc2MettreAJourEnBatch(token, toUpdate, groupesMap, deadline);
    sc2RetirerAppartenancesGroupes(token, toUpdate, groupesMap);
  }
  if (created || updated) Logger.log('  ✅ Diffusion HECG_* : ' + created + ' créés, ' + updated + ' mis à jour.');
  return { created, updated };
}

// Prévisualisation SANS AUCUN appel à l'API Google Contacts — lit uniquement le
// classeur et affiche, par libellé HECG_*/FORM-*, le nombre et le détail des membres
// qui seraient synchronisés. À lancer directement depuis l'éditeur Apps Script pour
// vérifier la colonne AA "Mail HECG" et la colonne "Classes" (Adultes) avant de
// laisser tourner la synchro réelle.
function previsualiserGroupesDiffusion() {
  const membres = [...sc2ExtractGroupesDiffusion(), ...sc2ExtractGroupesFormateursParClasse()];
  const parGroupe = {};
  membres.forEach(m => {
    (m.labels || []).forEach(l => {
      if (!parGroupe[l]) parGroupe[l] = [];
      parGroupe[l].push(m.prenom + ' ' + m.nom + ' <' + m.emails[0] + '>');
    });
  });

  const groupes = Object.keys(parGroupe).sort();
  if (groupes.length === 0) {
    Logger.log('ℹ️ Aucun membre détecté — la colonne AA "Mail HECG" (ETUDIANTS) et/ou mailPro (Adultes) sont probablement encore vides.');
    return;
  }
  Logger.log('📋 ' + groupes.length + ' groupe(s) de diffusion HECG_* détecté(s) — ' + membres.length + ' membre(s) au total :\n');
  groupes.forEach(g => {
    Logger.log('── ' + g + ' (' + parGroupe[g].length + ') ──');
    parGroupe[g].forEach(ligne => Logger.log('   • ' + ligne));
  });
}

// Supprime les fiches de diffusion HECG_* orphelines (marqueur userDefined
// HECG_DIFFUSION_ID) dont la clé ne correspond plus à aucun étudiant/formateur actif
// dans le CRM (ex. étudiant sorti depuis) — la synchro normale ne les crée/màj jamais
// mais ne les supprime pas non plus automatiquement (politique UPSERT-only).
function nettoyerFichesDiffusionOrphelinesTous() {
  const membresValides = new Set(sc2ExtractGroupesDiffusion().map(m => m.diffusionKey));
  Logger.log('🧹 Nettoyage des fiches de diffusion HECG_* orphelines — tous comptes (' + membresValides.size + ' clé(s) valide(s)).');
  for (const email of SC2_TARGET_USERS) {
    nettoyerFichesDiffusionOrphelines(email, membresValides);
    Utilities.sleep(1000);
  }
  Logger.log('✅ Terminé.');
}

function nettoyerFichesDiffusionOrphelines(emailUtilisateur, membresValides) {
  const email = emailUtilisateur || SC2_TARGET_USERS[0];
  let token;
  try { token = sc2GetToken(email); }
  catch(e) { Logger.log('❌ Auth (' + email + ') : ' + e.message); return; }

  const existants = sc2LireContactsExistants(token);
  const rns = existants.filter(e => {
    const key = ((e.userDefined || []).find(u => u.key === 'HECG_DIFFUSION_ID') || {}).value;
    return key && !membresValides.has(key);
  }).map(e => e.resourceName).filter(Boolean);

  if (!rns.length) { Logger.log('  ℹ️ ' + email + ' : aucune fiche de diffusion orpheline.'); return; }
  Logger.log('  🗑️ ' + email + ' : ' + rns.length + ' fiche(s) de diffusion orpheline(s) à supprimer...');
  for (let i = 0; i < rns.length; i += 500) {
    sc2ApiCall('https://people.googleapis.com/v1/people:batchDeleteContacts', {
      method: 'post', contentType: 'application/json', headers: sc2AuthHeader(token),
      payload: JSON.stringify({ resourceNames: rns.slice(i, i + 500) })
    });
    Utilities.sleep(500);
  }
  Logger.log('  ✅ ' + email + ' : ' + rns.length + ' fiche(s) supprimée(s).');
}

// ==========================================
// DÉCLENCHEUR QUOTIDIEN
// ==========================================

// Supprime l'ancien trigger V1 (synchroniserTousLesContacts) et installe V2 à 2h Paris.
// À lancer UNE SEULE FOIS manuellement après avoir validé que synchroniserTousLesContacts2 fonctionne.
function creerDeclencheurQuotidien2() {
  const OLD_FN = 'synchroniserTousLesContacts';  // ancien (synchrocontacts.gs)
  const NEW_FN = 'synchroniserTousLesContacts2'; // nouveau (synchro2.gs)

  let supprimesV1 = 0, supprimesV2 = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === OLD_FN) { ScriptApp.deleteTrigger(t); supprimesV1++; }
    if (fn === NEW_FN) { ScriptApp.deleteTrigger(t); supprimesV2++; }
  });

  if (supprimesV1 > 0) Logger.log('🗑️  Ancien trigger V1 supprimé (' + supprimesV1 + ').');
  if (supprimesV2 > 0) Logger.log('🔄 Trigger V2 existant supprimé pour recréation propre (' + supprimesV2 + ').');

  ScriptApp.newTrigger(NEW_FN)
    .timeBased().atHour(2).everyDays(1).inTimezone('Europe/Paris').create();

  // Vérification : lister tous les triggers actifs du projet
  Logger.log('\n📋 Triggers actifs après configuration :');
  ScriptApp.getProjectTriggers().forEach(t => {
    Logger.log('  • ' + t.getHandlerFunction()
      + ' — ' + t.getEventType()
      + (t.getTriggerSource() === ScriptApp.TriggerSource.CLOCK
        ? ' — déclenchement horaire'
        : ''));
  });
  Logger.log('\n✅ synchroniserTousLesContacts2 programmé chaque jour à 2h (Europe/Paris).');
}

// Affiche tous les triggers actifs sans rien modifier (diagnostic)
function listerTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  if (triggers.length === 0) { Logger.log('Aucun trigger actif.'); return; }
  Logger.log('📋 ' + triggers.length + ' trigger(s) actif(s) :');
  triggers.forEach(t => {
    Logger.log('  • ' + t.getHandlerFunction()
      + ' | source : ' + t.getTriggerSource()
      + ' | type : ' + t.getEventType());
  });
}
