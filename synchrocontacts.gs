// =========================================================================
// MOTEUR DE SYNCHRONISATION GOOGLE CONTACTS - HECG
// ⚠️ Aucune bibliothèque externe requise — auth JWT native via Utilities
// Stratégie : smart sync (update/create/delete diff) — jamais de reset total
// =========================================================================

const SERVICE_ACCOUNT_EMAIL = 'robot-signature-hecg@signature-hecg-officiel.iam.gserviceaccount.com';

const PRIVATE_KEY = '-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQCsdka615zrvROl\neHyFa/2aawECkcjwU6vQUhIXAq9JI/IsSNKhGU6cZRKe/wxX9KTyRL6ztnIvW0Pt\n1lAkDy4zKLO1dEpBpHNW+PyQrDIoG/ARwD0T14jV+4SmPYKevMWfCIs4X+vYHnoX\nFIZo4ei2tmbDoAk54eNDKf+oAXVKkFDNoHsQdBXLkynk6igC2dnEnVcGKattezOV\nTnfhPLK+cj/tTaJHRegU1aqdwoDfbA6FVFby7tlXsMc29OOewOLwj52UeW/XmDv8\n7AKNloBgc0QrEYkVS09HlO0hMrTOX7/Ao2JxovPg/R6BPjHxb1oX4Je7rs1onn7z\n5CwNEkslAgMBAAECggEAEmhRZlHrItI8jZXNnKQJHnk7U13iF5ymowaPfbtAoErg\n508ihCViWZkEIspQM/cdv+oMfLwFdf6Ewpb0WNTx9m3quHxgDJ+T2/2ZX4uxksxg\nlFRzcHG53jUJVIEONwkpAq9zxKGgV6HxIBOFwR4Tq6TOVST4tx/gFOQfsHvvW/Tc\nfrdBFx+WbSBV+kcWqB3wq0xHr9LgW54SZ8YM6H0FGald34S3IF/vwyjtLIYgRQ3/\nZNJIjRQRpOlgjUMXazkzluX9RGK3ydL7zP+xQli5BH/aImV+SuzzRGDLxue6zQCB\njYkU7y79uqr6AK8Gq6cf+/ypnviiG+ax2jLJv1Y/oQKBgQDVpw81mXz7xf+KAAdS\n1yGbixhoVDafaTaE2e8CQ2rFlh85DyRy3f7zQ+DyNOXukx3LjZXkbcMz060FaLKy\nX7Qzep2b2C1s9g+25ZIHZvUxncQvPq9K6vsexgMLJZo3dzfUTVj1H5VNzkCklHrB\nX1bYmGDbQVPCsapzCZxsTAt4PQKBgQDOpSrqP+8tLDt+a+RfSFExL+68jXpEsDAd\nJwa405Wg77o2YCumdbfRdZ6q0MBKAuGWRTuY7zUirhnhsgSu++gq+U82+KWYgjR4\nECkKaSDo+XmppZTCRGX+nJOY3y1qd0cB4hW+9cHI/++UGsP5yWbynfYIr/+U4zT6\nypmVS9dlCQKBgQCMbCqg7eqpiC82QmKN3fumwbsfBwqHp50/oAVpFWpdxxdqZztr\ni+D/fkOgrYfaUDMrEDnOUx4TODLl9TRN7H0BwLtKLMFedjNJ4IUj/FV3cNv6uVZ5\nBQxb44UolGRRxDebf+LR6Ro2czMleLld0w2/ehdexAcLVb5TsaNvwmNfeQKBgQC8\nt3RQx6CbJXkTxF6kcbvMatThF2dhEXJvPTPTWU+d0TDC9eMHOxxrSrpjjw78yFLS\nVFnQGizxhgQW7OeAEof9rv8b2coJVGesej2wxz+J5EOqnZAUNjjbZI0aoD6uq02K\nt7laUr/t22YlYKg3FypQSdfmKS0FANZibuIByWhlWQKBgHwQGyNHkEddB5IBNQ4G\nSRC2IzeOb+SvDmj+vgox3tM9mntOTa5nAJueD0tCi+Bor3WA3YxFdcOIf2ly39Ts\niTCYNtR7IBmXhrehndsqKfTz7K0mYm3GWqyQWyVPaRyHKb3Kc3eeMlAnweUYHbfV\nt2epNa0zscpTqcKo+JOs042c\n-----END PRIVATE KEY-----\n';

const TARGET_USERS = [
  'm.mokhtari@hecg.fr', 'a.mhoudini@hecg.fr', 'n.mas@hecg.fr', 'd.benyounes@hecg.fr',
  'v.velcker@hecg.fr', 'contact@hecg.fr', 'k.bouhassane@hecg.fr',
  'communication@hecg.fr'
];

// ==========================================
// POINT D'ENTRÉE PRINCIPAL — SMART SYNC
// ==========================================
// Principe : on ne supprime JAMAIS tout d'un coup.
// On lit les contacts existants, on calcule le diff, puis on :
//   • crée les contacts nouveaux
//   • met à jour les contacts modifiés
//   • supprime les contacts qui ne sont plus dans le tableur
// Cela évite le MY_CONTACTS_OVERFLOW_COUNT causé par le cycle
// "supprimer 1652 contacts → quota bloqué → impossibilité de recréer".

function synchroniserTousLesContacts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const dataEtu   = extractEtudiants(ss);
  const dataPart  = extractPartenaires(ss);
  const dataPerso = extractPersonnelVieScolaire(ss);
  const tousLesContacts = [...dataEtu, ...dataPart, ...dataPerso];

  Logger.log(`🔍 ${tousLesContacts.length} contacts trouvés. Début de la synchronisation...`);

  for (const userEmail of TARGET_USERS) {
    const email = userEmail.toLowerCase().trim();
    Logger.log(`\n➡️ Traitement pour : ${email}`);

    let token;
    try {
      token = getServiceAccountToken(email);
    } catch(e) {
      Logger.log(`❌ Échec d'authentification pour ${email} : ${e.message}`);
      continue;
    }

    // 1. Libellés nécessaires → créer les groupes manquants
    const tousLesLibelles = new Set();
    tousLesContacts.forEach(c => {
      const lbls = c.labels || [c.label];
      lbls.forEach(l => { if (l) tousLesLibelles.add(l); });
    });
    const groupesMap = preparerGroupes(token, tousLesLibelles);

    // 2. Contacts valides à synchroniser
    const contactsValides = tousLesContacts.filter(c =>
      (c.phones && c.phones.length > 0) || !!c.phone || !!c.entreprise || !!c.nom || !!c.prenom
    );

    // 3. Lire les contacts déjà présents dans Google Contacts
    const existants = lireContactsExistants(token);
    Logger.log(`📖 ${existants.length} contacts existants lus.`);

    // 4. Si 0 contacts mais des groupes existent → leurs compteurs internes sont corrompus
    //    (ghost member counts causant MY_CONTACTS_OVERFLOW_COUNT sur batchCreate).
    //    On les purge sans deleteContacts (aucun contact à supprimer) pour réinitialiser.
    if (existants.length === 0) {
      purgerGroupesFantomes(token);
      // Reconstruire la map des groupes après purge
      Object.keys(groupesMap).forEach(k => delete groupesMap[k]);
      const freshMap = preparerGroupes(token, tousLesLibelles);
      Object.assign(groupesMap, freshMap);
    }

    // 5. Smart sync : create / update / delete
    const stats = syncerContacts(token, contactsValides, existants, groupesMap);
    Logger.log(`✅ Terminé pour ${email} : ${stats.created} créés, ${stats.updated} mis à jour, ${stats.deleted} supprimés.`);

    // 6. Supprimer les groupes qui ne correspondent plus à aucun libellé actif
    nettoyerGroupesObsoletes(token, tousLesLibelles);
  }
}

// ==========================================
// LECTURE DES FEUILLES
// ==========================================

function extractEtudiants(ss) {
  const sheet   = getSheetSafe(ss, "ETUDIANTS");
  const data    = sheet.getDataRange().getValues();
  const partMap = buildPartenaireMap(ss);
  let contacts  = [];

  for (let i = 1; i < data.length; i++) {
    const nom    = String(data[i][2]  || "").trim();
    const prenom = String(data[i][3]  || "").trim();
    const tel    = formatPhone(data[i][5]);
    const typo   = String(data[i][20] || "").toLowerCase();
    const statut = String(data[i][13] || "").trim();
    const classe = String(data[i][14] || "").trim();
    const campus = String(data[i][22] || "").trim();
    const idCont = String(data[i][12] || "").trim();
    const nomEnt = String(data[i][21] || "").trim();

    if (!tel || (!prenom && !nom)) continue;

    // Label principal — ordre important : préinscrit AVANT inscrit
    let mainLabel;
    if (typo.includes("prein") || typo.includes("préin")) {
      mainLabel = "Étudiant Préinscrit";
    } else if (typo.includes("sorti")) {
      mainLabel = "Étudiant Sorti";
    } else if (typo.includes("inscrit") || typo.includes("alternance")) {
      mainLabel = "Étudiant Inscrit";
    } else {
      mainLabel = "Prospect";
    }

    const labelsSet = new Set([mainLabel]);
    const isProspect  = mainLabel === "Prospect";
    const isSorti     = mainLabel === "Étudiant Sorti";
    const isEnContrat = statut.toLowerCase().includes("contrat");

    if (isProspect) {
      // Aucun libellé supplémentaire
    } else if (isSorti) {
      if (isEnContrat && statut) labelsSet.add(statut);
    } else {
      if (statut) labelsSet.add(statut);
      if (classe)  labelsSet.add(classe);
      if (campus)  labelsSet.add(campus);
    }

    contacts.push({
      nom, prenom,
      phones:     [tel],
      emails:     [],
      labels:     [...labelsSet],
      entreprise: nomEnt,
      contactEnt: partMap[idCont] || null
    });
  }
  return contacts;
}

function buildPartenaireMap(ss) {
  const sheet = getSheetSafe(ss, "Partenariat");
  const data  = sheet.getDataRange().getValues();
  const map   = {};
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0] || "").trim();
    if (!id) continue;
    const prenom = String(data[i][4] || "").trim();
    const nom    = String(data[i][3] || "").trim();
    const tel    = formatPhone(data[i][5]) || formatPhone(data[i][9]);
    const nomEnt = String(data[i][2] || "").trim();
    map[id] = {
      nomComplet: [prenom, nom].filter(Boolean).join(" "),
      tel,
      nomEnt
    };
  }
  return map;
}

function extractPartenaires(ss) {
  const sheet            = getSheetSafe(ss, "Partenariat");
  const data             = sheet.getDataRange().getValues();
  const etudiantsParCont = buildEtudiantsParContact(ss);
  let contacts = [];

  for (let i = 1; i < data.length; i++) {
    const idContact  = String(data[i][0] || "").trim();
    const entreprise = String(data[i][2] || "").trim();
    const nom        = String(data[i][3] || "").trim();
    const prenom     = String(data[i][4] || "").trim();
    const tel1       = formatPhone(data[i][5]);
    const email1     = String(data[i][6] || "").trim();
    const email2     = String(data[i][8] || "").trim();
    const tel2       = formatPhone(data[i][9]);

    if (!entreprise && !nom && !prenom) continue;

    const phones = [];
    if (tel1) phones.push(tel1);
    if (tel2 && tel2 !== tel1) phones.push(tel2);

    const emails = [];
    if (email1 && email1.includes('@')) emails.push(email1);
    if (email2 && email2.includes('@') && email2 !== email1) emails.push(email2);

    contacts.push({ nom, prenom, phones, emails, label: "Entreprise", entreprise, etudiants: etudiantsParCont[idContact] || [] });
  }
  return contacts;
}

function buildEtudiantsParContact(ss) {
  const sheet = getSheetSafe(ss, "ETUDIANTS");
  const data  = sheet.getDataRange().getValues();
  const map   = {};
  for (let i = 1; i < data.length; i++) {
    const idCont = String(data[i][12] || "").trim();
    if (!idCont) continue;
    const prenom = String(data[i][3] || "").trim();
    const nom    = String(data[i][2] || "").trim();
    const nomComplet = [prenom, nom].filter(Boolean).join(" ");
    if (!nomComplet) continue;
    if (!map[idCont]) map[idCont] = [];
    map[idCont].push(nomComplet);
  }
  return map;
}

// Feuille "Adultes" (ou "FORMATEURS") — colonnes :
// A=ID B=Nom C=Prenom D=Mail_Perso E=Mail_Pro L=Tel K=Date_sortie M=Etat_formateur (OUI/NON)
function extractPersonnelVieScolaire(ss) {
  var ws;
  try { ws = getSheetSafe(ss, "Adultes"); } catch(e) { ws = getSheetSafe(ss, "FORMATEURS"); }
  var data = ws.getDataRange().getValues();
  var contacts = [];

  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;

    var nom       = String(data[i][1]  || "").trim();
    var prenom    = String(data[i][2]  || "").trim();
    var mailPerso = String(data[i][3]  || "").trim();
    var mailPro   = String(data[i][4]  || "").trim();
    var dateSortie = data[i][10];
    var tel       = formatPhone(data[i][11]);
    var etatFormateur = String(data[i][12] || "NON").trim().toUpperCase();

    if (!nom && !prenom) continue;

    var estSorti = dateSortie instanceof Date
      ? !isNaN(dateSortie.getTime())
      : String(dateSortie || "").trim() !== "";

    var isFormateur = (etatFormateur === "OUI");
    var label = isFormateur
      ? (estSorti ? "Formateurs sortis" : "Formateurs")
      : (estSorti ? "Encadrants sortis" : "Encadrants");

    var emails = [];
    if (mailPro   && mailPro.includes('@'))   emails.push(mailPro);
    if (mailPerso && mailPerso.includes('@') && mailPerso !== mailPro) emails.push(mailPerso);

    contacts.push({ nom: nom, prenom: prenom, phones: tel ? [tel] : [], emails: emails, label: label });
  }

  return contacts;
}

// ==========================================
// FORMATAGE ET NORMALISATION DES TÉLÉPHONES
// ==========================================

function formatPhone(phoneRaw) {
  if (!phoneRaw) return "";
  let cleaned = String(phoneRaw).trim().replace(/[^\d+]/g, '');

  let digits9 = null;

  if (cleaned.startsWith('+33') && cleaned.substring(3).length === 9) {
    digits9 = cleaned.substring(3);
  } else if (cleaned.startsWith('0033') && cleaned.substring(4).length === 9) {
    digits9 = cleaned.substring(4);
  } else if (cleaned.startsWith('0') && cleaned.length === 10) {
    digits9 = cleaned.substring(1);
  } else if (/^\d{9}$/.test(cleaned)) {
    digits9 = cleaned;
  }

  if (digits9) {
    return `+33 ${digits9[0]} ${digits9.slice(1,3)} ${digits9.slice(3,5)} ${digits9.slice(5,7)} ${digits9.slice(7,9)}`;
  }

  if (!cleaned.startsWith('+')) cleaned = '+' + cleaned;
  return cleaned.length > 8 ? cleaned : "";
}

function phoneKey(phone) {
  return String(phone || "").replace(/\D/g, '');
}

// ==========================================
// AUTHENTIFICATION JWT NATIVE (sans bibliothèque)
// ==========================================

function getServiceAccountToken(userEmail) {
  const now = Math.floor(Date.now() / 1000);

  const header = _b64url(JSON.stringify({ alg: "RS256", typ: "JWT" }));
  const claim  = _b64url(JSON.stringify({
    iss:   SERVICE_ACCOUNT_EMAIL,
    sub:   userEmail,
    scope: 'https://www.googleapis.com/auth/contacts',
    aud:   'https://oauth2.googleapis.com/token',
    iat:   now,
    exp:   now + 3600
  }));

  const toSign    = header + '.' + claim;
  const signature = _b64url(Utilities.computeRsaSha256Signature(toSign, PRIVATE_KEY));
  const jwt       = toSign + '.' + signature;

  const resp = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    payload: {
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      assertion: jwt
    },
    muteHttpExceptions: true
  });

  const result = JSON.parse(resp.getContentText());
  if (!result.access_token) {
    throw new Error(resp.getContentText());
  }
  return result.access_token;
}

function _b64url(input) {
  const bytes = (typeof input === 'string')
    ? Utilities.newBlob(input).getBytes()
    : input;
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, '');
}

// ==========================================
// LECTURE DES CONTACTS EXISTANTS
// ==========================================

function lireContactsExistants(token) {
  const baseUrl = 'https://people.googleapis.com/v1/people/me/connections'
    + '?personFields=names,phoneNumbers,emailAddresses,organizations,memberships,userDefined,relations'
    + '&pageSize=1000';
  const contacts = [];
  let pageToken = null;
  let page = 0;
  do {
    const url = baseUrl + (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');
    const data = JSON.parse(UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    }).getContentText());
    (data.connections || []).forEach(c => contacts.push(c));
    pageToken = data.nextPageToken || null;
  } while (pageToken && ++page < 30);
  return contacts;
}

// ==========================================
// SYNCHRONISATION INTELLIGENTE (DIFF)
// ==========================================

function syncerContacts(token, newContacts, existants, groupesMap) {
  // Index des contacts existants par téléphone (clé principale) et par nom (fallback)
  const phoneIdx = {}, nameIdx = {};
  existants.forEach(e => {
    (e.phoneNumbers || []).forEach(p => {
      const k = phoneKey(p.value);
      if (k && !phoneIdx[k]) phoneIdx[k] = e;
    });
    const n = (e.names || [])[0];
    if (n) {
      const nk = (String(n.familyName || '') + '|' + String(n.givenName || '')).toLowerCase();
      if (nk !== '|' && !nameIdx[nk]) nameIdx[nk] = e;
    }
  });

  const toCreate = [], toUpdate = [];
  const seenRns = new Set();

  for (const c of newContacts) {
    const phones = c.phones || (c.phone ? [c.phone] : []);
    let existing = null;

    // Recherche par téléphone d'abord
    for (const p of phones) {
      const k = phoneKey(p);
      if (k && phoneIdx[k]) { existing = phoneIdx[k]; break; }
    }
    // Fallback par nom+prénom
    if (!existing) {
      const nk = (String(c.nom || '') + '|' + String(c.prenom || '')).toLowerCase();
      if (nk !== '|' && nameIdx[nk]) existing = nameIdx[nk];
    }

    if (existing) {
      seenRns.add(existing.resourceName);
      toUpdate.push({ contactData: c, resourceName: existing.resourceName, etag: existing.etag });
    } else {
      toCreate.push(c);
    }
  }

  // Contacts dans Google qui ne correspondent plus à aucun contact du tableur → supprimer
  const toDelete = existants
    .filter(e => !seenRns.has(e.resourceName))
    .map(e => e.resourceName);

  Logger.log(`📊 Créer: ${toCreate.length}, Màj: ${toUpdate.length}, Supprimer: ${toDelete.length}`);

  let created = 0, updated = 0, deleted = 0;

  if (toCreate.length > 0) {
    created = importerContactsEnBatch(token, toCreate, groupesMap);
  }

  if (toUpdate.length > 0) {
    mettreAJourContactsEnBatch(token, toUpdate, groupesMap);
    updated = toUpdate.length;
  }

  if (toDelete.length > 0) {
    Logger.log(`🗑️ Suppression de ${toDelete.length} contacts obsolètes...`);
    for (let i = 0; i < toDelete.length; i += 500) {
      UrlFetchApp.fetch('https://people.googleapis.com/v1/people:batchDeleteContacts', {
        method: 'post',
        contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + token },
        payload: JSON.stringify({ resourceNames: toDelete.slice(i, i + 500) }),
        muteHttpExceptions: true
      });
      if (i + 500 < toDelete.length) Utilities.sleep(1000);
    }
    deleted = toDelete.length;
  }

  return { created, updated, deleted };
}

// ==========================================
// CRÉATION EN BATCH (nouveaux contacts)
// ==========================================

// Crée les contacts un par un via createContact (batchCreateContacts déclenche
// MY_CONTACTS_OVERFLOW_COUNT même après quelques dizaines de créations).
// Cadence : 1500ms entre chaque appel → ~40 créations/min, conservateur.
// Budget : 5 min par utilisateur (inclut les pauses retry).
// Chaque contact bénéficie d'un seul retry en cas de 429 (évite la boucle infinie).
// Le smart sync reprend exactement où il en était à chaque relance.
function importerContactsEnBatch(token, contacts, groupesMap) {
  const PAUSE_MS  = 1500;
  const BUDGET_MS = 5 * 60 * 1000; // 5 min max par utilisateur (pauses incluses)
  const runStart  = Date.now();
  let total   = 0;
  let retried = false;

  for (let i = 0; i < contacts.length; i++) {
    if (Date.now() - runStart > BUDGET_MS) {
      Logger.log(`⏰ Budget 5 min atteint. Arrêt après ${total} contacts créés. ${contacts.length - i} restants — relancez pour continuer.`);
      break;
    }

    const person = buildContactPerson(contacts[i], groupesMap);
    if (!person) { retried = false; continue; }

    const resp = UrlFetchApp.fetch('https://people.googleapis.com/v1/people:createContact', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify(person),
      muteHttpExceptions: true
    });

    const result = JSON.parse(resp.getContentText());
    if (result.error) {
      const code    = result.error.code;
      const status  = result.error.status  || '';
      const msg     = result.error.message || '';
      const details = result.error.details || [];
      const reason  = details.length > 0 ? (details[0].reason || '') : '';
      Logger.log(`⚠️ Contact ${i} [${code}] status=${status} reason=${reason} : ${msg}`);

      if (code === 429) {
        const isOverflow = reason.includes('OVERFLOW') || msg.includes('OVERFLOW');
        if (isOverflow) {
          if (i === 0 || retried) {
            // Quota myContacts saturé pour la journée (fenêtre ~24h).
            // Inutile de réessayer — passer à l'utilisateur suivant.
            Logger.log(`🚫 Quota myContacts saturé — relancez demain pour cet utilisateur.`);
            break;
          }
          // Premier échec overflow non en contact 0 → un seul retry après 2 min
          Logger.log(`⏳ OVERFLOW — pause 2 min avant réessai unique du contact ${i}...`);
          Utilities.sleep(120000);
          retried = true;
          i--;
        } else if (!retried) {
          Logger.log(`⏳ Rate limit — pause 65s avant réessai unique...`);
          Utilities.sleep(65000);
          retried = true;
          i--;
        } else {
          Logger.log(`⏭️ Contact ${i} ignoré après double échec — sera repris au prochain run.`);
          retried = false;
        }
      } else {
        retried = false;
      }
    } else {
      total++;
      retried = false;
    }

    Utilities.sleep(PAUSE_MS);
  }
  return total;
}

// ==========================================
// MISE À JOUR EN BATCH (contacts existants modifiés)
// ==========================================

function mettreAJourContactsEnBatch(token, updates, groupesMap) {
  const UPDATE_FIELDS = 'names,phoneNumbers,emailAddresses,organizations,memberships,userDefined,relations';

  for (const { contactData, resourceName, etag } of updates) {
    const person = buildContactPerson(contactData, groupesMap);
    if (!person) continue;
    if (etag) person.etag = etag;

    const url = `https://people.googleapis.com/v1/${resourceName}:updateContact`
      + `?updatePersonFields=${encodeURIComponent(UPDATE_FIELDS)}`;

    const resp = UrlFetchApp.fetch(url, {
      method: 'patch',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify(person),
      muteHttpExceptions: true
    });

    const result = JSON.parse(resp.getContentText());
    if (result.error) {
      Logger.log(`⚠️ Erreur màj ${resourceName} [${result.error.code}] : ${result.error.message}`);
    }

    Utilities.sleep(700);
  }
}

// ==========================================
// PURGE DES GROUPES AVEC COMPTEURS FANTÔMES
// ==========================================
// Appelée quand connections.list = 0 mais des groupes subsistent.
// Leurs memberCount internes peuvent rester > 0 côté Google même après suppression
// de contacts, ce qui provoque MY_CONTACTS_OVERFLOW_COUNT sur batchCreateContacts.
// On supprime uniquement les groupes (deleteContacts=false) pour forcer un reset.

function purgerGroupesFantomes(token) {
  const urlBase = 'https://people.googleapis.com/v1/contactGroups';
  let groupRns = [];
  let nextPageToken = null;

  do {
    const pageUrl = urlBase + (nextPageToken ? '?pageToken=' + encodeURIComponent(nextPageToken) : '');
    const data = JSON.parse(UrlFetchApp.fetch(pageUrl, {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    }).getContentText());
    (data.contactGroups || []).forEach(g => {
      if (g.groupType === 'USER_CONTACT_GROUP') groupRns.push(g.resourceName);
    });
    nextPageToken = data.nextPageToken || null;
  } while (nextPageToken);

  if (groupRns.length === 0) return;

  Logger.log(`🔄 Purge de ${groupRns.length} groupes fantômes (reset compteurs)...`);
  for (const rn of groupRns) {
    UrlFetchApp.fetch('https://people.googleapis.com/v1/' + rn + '?deleteContacts=false', {
      method: 'delete',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });
    Utilities.sleep(300);
  }
  Logger.log(`✅ ${groupRns.length} groupes supprimés — compteurs réinitialisés.`);
}

// ==========================================
// NETTOYAGE DES GROUPES OBSOLÈTES
// ==========================================

function nettoyerGroupesObsoletes(token, libellesActifs) {
  const urlBase = 'https://people.googleapis.com/v1/contactGroups';
  let userGroups = [];
  let nextPageToken = null;
  do {
    const pageUrl = urlBase + (nextPageToken ? '?pageToken=' + encodeURIComponent(nextPageToken) : '');
    const data = JSON.parse(UrlFetchApp.fetch(pageUrl, {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    }).getContentText());
    (data.contactGroups || []).forEach(g => {
      if (g.groupType === 'USER_CONTACT_GROUP') userGroups.push(g);
    });
    nextPageToken = data.nextPageToken || null;
  } while (nextPageToken);

  const obsoletes = userGroups.filter(g => !libellesActifs.has(g.name));
  if (obsoletes.length === 0) return;

  Logger.log(`🗑️ Suppression de ${obsoletes.length} groupe(s) obsolète(s)...`);
  for (const g of obsoletes) {
    UrlFetchApp.fetch('https://people.googleapis.com/v1/' + g.resourceName + '?deleteContacts=false', {
      method: 'delete',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });
    Utilities.sleep(300);
  }
  Logger.log(`✅ ${obsoletes.length} groupe(s) obsolète(s) supprimé(s).`);
}

// ==========================================
// API GOOGLE CONTACTS — UTILITAIRES
// ==========================================

function preparerGroupes(token, libellesNecessaires) {
  const urlBase = 'https://people.googleapis.com/v1/contactGroups';

  const map = {};
  let nextPageToken = null;
  do {
    const pageUrl = urlBase + (nextPageToken ? '?pageToken=' + encodeURIComponent(nextPageToken) : '');
    const resp = UrlFetchApp.fetch(pageUrl, {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });
    const raw = resp.getContentText();
    let data;
    try {
      data = JSON.parse(raw);
    } catch(e) {
      Logger.log(`⚠️ preparerGroupes — réponse non-JSON (HTTP ${resp.getResponseCode()}) : ${raw.substring(0, 300)}`);
      return map;
    }
    if (data.error) {
      Logger.log(`⚠️ preparerGroupes — erreur API [${data.error.code}] : ${data.error.message}`);
      return map;
    }
    (data.contactGroups || []).forEach(g => { map[g.name] = g.resourceName; });
    nextPageToken = data.nextPageToken || null;
  } while (nextPageToken);

  for (const nom of libellesNecessaires) {
    if (!nom || map[nom]) continue;

    let attempts = 0;
    while (attempts < 3) {
      const createRes = UrlFetchApp.fetch(urlBase, {
        method: 'post',
        contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + token },
        payload: JSON.stringify({ contactGroup: { name: nom } }),
        muteHttpExceptions: true
      });
      const created = JSON.parse(createRes.getContentText());

      if (created.resourceName) {
        map[nom] = created.resourceName;
        break;
      } else if (created.error && created.error.code === 429) {
        Logger.log(`⏳ Rate limit — pause 65s avant réessai du groupe "${nom}"...`);
        Utilities.sleep(65000);
        attempts++;
      } else if (created.error && created.error.code === 409) {
        const refetch = JSON.parse(UrlFetchApp.fetch(urlBase, {
          headers: { Authorization: 'Bearer ' + token },
          muteHttpExceptions: true
        }).getContentText());
        (refetch.contactGroups || []).forEach(g => { map[g.name] = g.resourceName; });
        break;
      } else {
        Logger.log(`⚠️ Groupe "${nom}" non créé : ${JSON.stringify(created.error || created)}`);
        break;
      }
    }
    Utilities.sleep(700);
  }
  return map;
}

function buildContactPerson(contactData, groupesMap) {
  const phoneList = contactData.phones || (contactData.phone ? [contactData.phone] : []);
  const emailList = contactData.emails
    || (contactData.email && contactData.email.includes('@') ? [contactData.email] : []);

  const labels = contactData.labels || [contactData.label];
  const memberships = labels
    .filter(l => l && groupesMap[l])
    .map(l => ({ contactGroupMembership: { contactGroupResourceName: groupesMap[l] } }));

  if (memberships.length === 0) return null;

  const person = {
    names:       [{ givenName: contactData.prenom || '', familyName: contactData.nom || '' }],
    memberships: memberships
  };

  if (phoneList.length > 0) {
    person.phoneNumbers = phoneList.map((p, i) => ({ value: p, type: i === 0 ? 'mobile' : 'work' }));
  }
  if (emailList.length > 0) {
    person.emailAddresses = emailList.map(e => ({ value: e, type: 'work' }));
  }
  if (contactData.entreprise) {
    person.organizations = [{ name: contactData.entreprise, type: 'work' }];
  }

  const champsSup = [];
  if (contactData.contactEnt) {
    const ce = contactData.contactEnt;
    if (ce.nomComplet) person.relations = [{ person: ce.nomComplet, type: 'manager' }];
    if (ce.nomEnt) champsSup.push({ key: 'Entreprise contact',     value: ce.nomEnt });
    if (ce.tel)    champsSup.push({ key: 'Tél contact entreprise', value: ce.tel });
  }
  if (contactData.etudiants && contactData.etudiants.length > 0) {
    champsSup.push({ key: 'Étudiants HECG', value: contactData.etudiants.join(', ') });
  }
  if (champsSup.length > 0) person.userDefined = champsSup;

  return person;
}

// ==========================================
// DIAGNOSTIC — à exécuter pour identifier le blocage
// ==========================================

function diagnostiquerContacts() {
  const email = TARGET_USERS[0].toLowerCase().trim();
  Logger.log('=== Diagnostic Google Contacts pour : ' + email + ' ===');

  let token;
  try { token = getServiceAccountToken(email); }
  catch(e) { Logger.log('❌ Auth échouée : ' + e.message); return; }

  // 1. Nombre de contacts vus par connections.list (totalItems)
  const connResp = UrlFetchApp.fetch(
    'https://people.googleapis.com/v1/people/me/connections?personFields=names&pageSize=1',
    { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
  );
  const connData = JSON.parse(connResp.getContentText());
  Logger.log('connections.list → totalItems : ' + (connData.totalItems || connData.totalPeople || 0));

  // 2. Nombre réel dans le groupe système "myContacts" (la source du quota OVERFLOW)
  const mcResp = UrlFetchApp.fetch(
    'https://people.googleapis.com/v1/contactGroups/myContacts?maxMembers=0',
    { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
  );
  const mcData = JSON.parse(mcResp.getContentText());
  Logger.log('myContacts système → memberCount : ' + (mcData.memberCount !== undefined ? mcData.memberCount : 'N/A'));

  // 3. Nombre de groupes utilisateur existants
  const grpResp = UrlFetchApp.fetch(
    'https://people.googleapis.com/v1/contactGroups',
    { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
  );
  const grpData = JSON.parse(grpResp.getContentText());
  const userGroups = (grpData.contactGroups || []).filter(g => g.groupType === 'USER_CONTACT_GROUP');
  Logger.log('Groupes utilisateur : ' + userGroups.length + ' → ' + userGroups.map(g => g.name).join(', '));

  // Réponse brute myContacts pour debug
  Logger.log('myContacts raw : ' + mcResp.getContentText().substring(0, 300));

  const rnsANettoyer = [];

  // 4a. createContact SANS membership (→ Other Contacts)
  const t1 = JSON.parse(UrlFetchApp.fetch('https://people.googleapis.com/v1/people:createContact', {
    method: 'post', contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ names: [{ givenName: 'TEST1', familyName: 'DIAG' }], phoneNumbers: [{ value: '+33600000001' }] }),
    muteHttpExceptions: true
  }).getContentText());
  Logger.log('createContact SANS membership : ' + (t1.resourceName ? '✅ OK → ' + t1.resourceName : '❌ ' + JSON.stringify(t1.error)));
  if (t1.resourceName) rnsANettoyer.push(t1.resourceName);

  // 4b. createContact AVEC membership (→ My Contacts)
  const premierGroupe = userGroups.length > 0 ? userGroups[0].resourceName : null;
  if (premierGroupe) {
    const t2 = JSON.parse(UrlFetchApp.fetch('https://people.googleapis.com/v1/people:createContact', {
      method: 'post', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({
        names: [{ givenName: 'TEST2', familyName: 'DIAG' }],
        phoneNumbers: [{ value: '+33600000002' }],
        memberships: [{ contactGroupMembership: { contactGroupResourceName: premierGroupe } }]
      }),
      muteHttpExceptions: true
    }).getContentText());
    Logger.log('createContact AVEC membership (' + userGroups[0].name + ') : ' + (t2.resourceName ? '✅ OK → ' + t2.resourceName : '❌ ' + JSON.stringify(t2.error)));
    if (t2.resourceName) rnsANettoyer.push(t2.resourceName);
  } else {
    Logger.log('createContact AVEC membership : ⚠️ aucun groupe existant pour tester');
  }

  // 4c. batchCreateContacts avec 1 seul contact SANS membership
  const t3 = JSON.parse(UrlFetchApp.fetch('https://people.googleapis.com/v1/people:batchCreateContacts', {
    method: 'post', contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({
      contacts: [{ contactPerson: { names: [{ givenName: 'TEST3', familyName: 'DIAG' }], phoneNumbers: [{ value: '+33600000003' }] } }],
      readMask: 'names'
    }),
    muteHttpExceptions: true
  }).getContentText());
  Logger.log('batchCreateContacts (1 contact SANS membership) : ' + (t3.error ? '❌ ' + JSON.stringify(t3.error) : '✅ OK → ' + JSON.stringify((t3.createdPeople || []).map(p => p.person && p.person.resourceName))));
  if (t3.createdPeople) t3.createdPeople.forEach(p => p.person && p.person.resourceName && rnsANettoyer.push(p.person.resourceName));

  // 4d. batchCreateContacts avec 1 seul contact AVEC membership
  if (premierGroupe) {
    const t4 = JSON.parse(UrlFetchApp.fetch('https://people.googleapis.com/v1/people:batchCreateContacts', {
      method: 'post', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({
        contacts: [{ contactPerson: {
          names: [{ givenName: 'TEST4', familyName: 'DIAG' }],
          phoneNumbers: [{ value: '+33600000004' }],
          memberships: [{ contactGroupMembership: { contactGroupResourceName: premierGroupe } }]
        } }],
        readMask: 'names'
      }),
      muteHttpExceptions: true
    }).getContentText());
    Logger.log('batchCreateContacts (1 contact AVEC membership) : ' + (t4.error ? '❌ ' + JSON.stringify(t4.error) : '✅ OK → ' + JSON.stringify((t4.createdPeople || []).map(p => p.person && p.person.resourceName))));
    if (t4.createdPeople) t4.createdPeople.forEach(p => p.person && p.person.resourceName && rnsANettoyer.push(p.person.resourceName));
  }

  // Nettoyage de tous les contacts tests créés
  if (rnsANettoyer.length > 0) {
    UrlFetchApp.fetch('https://people.googleapis.com/v1/people:batchDeleteContacts', {
      method: 'post', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({ resourceNames: rnsANettoyer }),
      muteHttpExceptions: true
    });
    Logger.log('🗑️ ' + rnsANettoyer.length + ' contact(s) test supprimé(s).');
  }

  Logger.log('=== Fin du diagnostic ===');
}

// ==========================================
// DÉCLENCHEUR AUTOMATIQUE (à exécuter UNE SEULE FOIS)
// ==========================================

function creerDeclencheurQuotidien() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'synchroniserTousLesContacts') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('synchroniserTousLesContacts')
    .timeBased()
    .atHour(2)
    .everyDays(1)
    .inTimezone('Europe/Paris')
    .create();
  Logger.log('✅ Déclencheur activé : synchronisation automatique chaque jour à 2h (heure de Paris).');
}
