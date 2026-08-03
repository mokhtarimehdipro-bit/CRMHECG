// =========================================================================
// MOTEUR DE SYNCHRONISATION GOOGLE CONTACTS V2 — HECG
// Corrections vs V1 :
//   • Protection JSON.parse sur TOUS les appels API
//   • try/catch par utilisateur (un échec n'arrête pas les autres)
//   • Vérification directe de myContacts.memberCount avant toute création
//   • Purge forcée si compteur fantôme détecté (memberCount >> contacts visibles)
//   • Retry 429 dans mettreAJourContactsEnBatch
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

function synchroniserTousLesContacts2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const dataEtu   = sc2ExtractEtudiants(ss);
  const dataPart  = sc2ExtractPartenaires(ss);
  const dataPerso = sc2ExtractPersonnel(ss);
  const tousLesContacts = [...dataEtu, ...dataPart, ...dataPerso];

  Logger.log('🔍 ' + tousLesContacts.length + ' contacts source. Début de la synchronisation V2...');

  const tousLibelles = new Set();
  tousLesContacts.forEach(c => {
    (c.labels || [c.label]).forEach(l => { if (l) tousLibelles.add(l); });
  });

  const contactsValides = tousLesContacts.filter(c =>
    (c.phones && c.phones.length > 0) || c.phone || c.entreprise || c.nom || c.prenom
  );

  for (const userEmail of SC2_TARGET_USERS) {
    Logger.log('\n➡️ Traitement : ' + userEmail);
    try {
      sc2SyncUser(userEmail, contactsValides, tousLibelles);
    } catch(e) {
      Logger.log('❌ Erreur non récupérée pour ' + userEmail + ' : ' + e.message + '\n' + (e.stack || ''));
    }
  }

  Logger.log('\n✅ Synchronisation V2 terminée.');
}

function sc2SyncUser(email, contactsValides, tousLibelles) {
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
  const stats = sc2SyncerContacts(token, contactsValides, existants, groupesMap);
  Logger.log('  ✅ ' + email + ' : ' + stats.created + ' créés, ' + stats.updated + ' màj, ' + stats.deleted + ' supprimés.');

  // 6. Nettoyer les groupes obsolètes
  sc2NettoyerGroupesObsoletes(token, tousLibelles);
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
    + '?personFields=names,phoneNumbers,emailAddresses,organizations,memberships,userDefined,relations'
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
function sc2SyncerContacts(token, newContacts, existants, groupesMap) {
  existants   = existants   || [];
  newContacts = newContacts || [];

  // ── Index des contacts existants ────────────────────────────────────────
  const phoneIdx = {}, nameIdx = {};
  existants.forEach(e => {
    (e.phoneNumbers || []).forEach(p => {
      const k = sc2PhoneKey(p.value);
      if (k && !phoneIdx[k]) phoneIdx[k] = e;
    });
    const n = (e.names || [])[0];
    if (n) {
      const nk = (String(n.familyName || '') + '|' + String(n.givenName || '')).toLowerCase();
      if (nk !== '|' && !nameIdx[nk]) nameIdx[nk] = e;
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
    // Fallback : recherche par nom+prénom
    if (!existing) {
      const nk = (String(c.nom || '') + '|' + String(c.prenom || '')).toLowerCase();
      if (nk !== '|' && nameIdx[nk]) existing = nameIdx[nk];
    }

    if (existing) {
      // Vérifier si une mise à jour est nécessaire (évite les appels API inutiles)
      if (sc2ContactAChange(c, existing, groupesMap)) {
        toUpdate.push({ contactData: c, resourceName: existing.resourceName, etag: existing.etag });
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
  if (toCreate.length > 0) created = sc2ImporterEnBatch(token, toCreate, groupesMap);
  if (toUpdate.length > 0) updated = sc2MettreAJourEnBatch(token, toUpdate, groupesMap);

  return { created, updated, deleted: 0 };
}

// Détecte si le contact CRM diffère du contact Google existant
// Champs comparés : nom, prénom, téléphones, emails, entreprise, libellés (classes/groupes)
function sc2ContactAChange(c, existing, groupesMap) {
  // Nom + prénom
  const existName = (existing.names || [])[0] || {};
  if (String(c.nom    || '') !== String(existName.familyName || '')) return true;
  if (String(c.prenom || '') !== String(existName.givenName  || '')) return true;

  // Téléphones (ensemble non ordonné, comparaison sur chiffres uniquement)
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
function sc2ImporterEnBatch(token, contacts, groupesMap) {
  const BATCH_SIZE = 200;
  let total = 0;

  for (let i = 0; i < contacts.length; i += BATCH_SIZE) {
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

function sc2MettreAJourEnBatch(token, updates, groupesMap) {
  const FIELDS = 'names,phoneNumbers,emailAddresses,organizations,memberships,userDefined,relations';
  let count = 0;

  for (const { contactData, resourceName, etag } of updates) {
    const person = sc2BuildPerson(contactData, groupesMap);
    if (!person) continue;
    if (etag) person.etag = etag;

    const url = 'https://people.googleapis.com/v1/' + resourceName + ':updateContact'
      + '?updatePersonFields=' + encodeURIComponent(FIELDS);

    let retried = false;
    while (true) {
      const result = sc2ApiCall(url, {
        method: 'patch', contentType: 'application/json',
        headers: sc2AuthHeader(token),
        payload: JSON.stringify(person)
      });
      if (result._parseError) {
        Logger.log('  ⚠️ Màj ' + resourceName + ' : réponse non-JSON'); break;
      }
      if (result.error) {
        const code = result.error.code;
        if (code === 429 && !retried) {
          Logger.log('  ⏳ Rate limit màj — pause 65s...');
          Utilities.sleep(65000); retried = true; continue;
        }
        Logger.log('  ⚠️ Erreur màj ' + resourceName + ' [' + code + '] : ' + result.error.message);
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
  if (phoneList.length > 0)
    person.phoneNumbers = phoneList.map((p, i) => ({ value: p, type: i === 0 ? 'mobile' : 'work' }));
  if (emailList.length > 0)
    person.emailAddresses = emailList.map(e => ({ value: e, type: 'work' }));
  if (c.entreprise)
    person.organizations = [{ name: c.entreprise, type: 'work' }];

  const extra = [];
  if (c.contactEnt) {
    const ce = c.contactEnt;
    if (ce.nomComplet) person.relations = [{ person: ce.nomComplet, type: 'manager' }];
    if (ce.nomEnt) extra.push({ key: 'Entreprise contact',     value: ce.nomEnt });
    if (ce.tel)    extra.push({ key: 'Tél contact entreprise', value: ce.tel });
  }
  if (c.etudiants && c.etudiants.length > 0)
    extra.push({ key: 'Étudiants HECG', value: c.etudiants.join(', ') });
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
  if (cleaned.startsWith('+33') && cleaned.substring(3).length === 9)      digits9 = cleaned.substring(3);
  else if (cleaned.startsWith('0033') && cleaned.substring(4).length === 9) digits9 = cleaned.substring(4);
  else if (cleaned.startsWith('0') && cleaned.length === 10)                digits9 = cleaned.substring(1);
  else if (/^\d{9}$/.test(cleaned))                                         digits9 = cleaned;
  if (digits9) return '+33 ' + digits9[0] + ' ' + digits9.slice(1,3) + ' ' + digits9.slice(3,5) + ' ' + digits9.slice(5,7) + ' ' + digits9.slice(7,9);
  if (!cleaned.startsWith('+')) cleaned = '+' + cleaned;
  return cleaned.length > 8 ? cleaned : '';
}

function sc2PhoneKey(phone) { return String(phone || '').replace(/\D/g, ''); }

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

function sc2ExtractEtudiants(ss) {
  const data    = getSheetSafe(ss, 'ETUDIANTS').getDataRange().getValues();
  const partMap = sc2BuildPartMap(ss);
  const list    = [];
  for (let i = 1; i < data.length; i++) {
    const nom    = String(data[i][2]  || '').trim();
    const prenom = String(data[i][3]  || '').trim();
    const tel    = sc2FormatPhone(data[i][5]);
    const typo   = String(data[i][20] || '').toLowerCase();
    const statut = String(data[i][13] || '').trim();
    const classe = String(data[i][14] || '').trim();
    const campus = String(data[i][22] || '').trim();
    const idCont = String(data[i][12] || '').trim();
    const nomEnt = String(data[i][21] || '').trim();
    if (!tel || (!prenom && !nom)) continue;
    let mainLabel;
    if (typo.includes('prein') || typo.includes('préin')) mainLabel = 'Étudiant Préinscrit';
    else if (typo.includes('sorti'))                       mainLabel = 'Étudiant Sorti';
    else if (typo.includes('inscrit') || typo.includes('alternance')) mainLabel = 'Étudiant Inscrit';
    else                                                   mainLabel = 'Prospect';
    const labelsSet = new Set([mainLabel]);
    const isSorti = mainLabel === 'Étudiant Sorti';
    if (mainLabel !== 'Prospect') {
      if (isSorti) { if (statut.toLowerCase().includes('contrat') && statut) labelsSet.add(statut); }
      else { if (statut) labelsSet.add(statut); if (classe) labelsSet.add(classe); if (campus) labelsSet.add(campus); }
    }
    list.push({ nom, prenom, phones: [tel], emails: [], labels: [...labelsSet], entreprise: nomEnt, contactEnt: partMap[idCont] || null });
  }
  return list;
}

function sc2ExtractPartenaires(ss) {
  const data           = getSheetSafe(ss, 'Partenariat').getDataRange().getValues();
  const etudiantsMap   = sc2BuildEtudiantsParContact(ss);
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
    if (!entreprise && !nom && !prenom) continue;
    const phones = []; if (tel1) phones.push(tel1); if (tel2 && tel2 !== tel1) phones.push(tel2);
    const emails = []; if (email1.includes('@')) emails.push(email1); if (email2.includes('@') && email2 !== email1) emails.push(email2);
    list.push({ nom, prenom, phones, emails, label: 'Entreprise', entreprise, etudiants: etudiantsMap[idContact] || [] });
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
