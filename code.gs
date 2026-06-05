
/**
 * CRM HECG - MOTEUR BACKEND V12.5 "PRO-COM"
 * Base de données centralisée, Intelligence Commerciale & Marketing Automation
 */

const FOLDER_ID_CV = "1vX3WhpqqLkcmZkMWw5EwzqFpQQxtl-_x";
const FOLDER_ID_COM = "1yQWMv2z_FtP8wSrc3WmidNR6uNikjuJ8";
const SENDER_EMAIL = "m.mokhtari@hecg.fr";

function getSignatureHTML() { return `<div style="margin-top:30px;"><a href="https://www.linkedin.com/in/mehdi-mokhtari-757225218" target="_blank"><img src="https://drive.google.com/uc?export=view&id=11vwuqgHIhK4uvPUIWBzuGZhs1ledNNT0" alt="Signature" style="display:block;max-width:500px;height:auto;border:none;"></a></div>`; }

function getSheetSafe(ss, name) {
  var s = ss.getSheetByName(name) || ss.getSheetByName(name.toLowerCase()) || ss.getSheetByName(name.charAt(0).toUpperCase() + name.slice(1).toLowerCase());
  if(!s) throw new Error("Onglet introuvable : " + name); return s;
}

function createJsonResponse(data) { return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); }

// ==========================================
// GESTION DES SESSIONS (sécurité)
// ==========================================
function _getSessionSheet(ss) {
  var ws = ss.getSheetByName("Sessions");
  if (!ws) {
    ws = ss.insertSheet("Sessions");
    ws.appendRow(["Token","UserRef","Email","Nom","Role","VieScRole","CreatedAt"]);
    ws.setFrozenRows(1);
    ws.hideSheet();
  }
  return ws;
}

function _createSession(ss, userRef, email, nom, role, vieScRole) {
  var token = Utilities.getUuid();
  var ws = _getSessionSheet(ss);
  ws.appendRow([token, String(userRef||""), String(email||""), String(nom||""), String(role||""), String(vieScRole||""), new Date()]);
  // Nettoyage des sessions de +7 jours (max 100 lignes)
  try {
    var data = ws.getDataRange().getValues();
    var now = new Date();
    var toDelete = [];
    for (var i = data.length - 1; i >= 1; i--) {
      var created = data[i][6];
      if (created instanceof Date && (now - created) > 7 * 24 * 60 * 60 * 1000) toDelete.push(i + 1);
    }
    toDelete.forEach(function(row) { ws.deleteRow(row); });
  } catch(e) {}
  return token;
}

function _validateSession(ss, token) {
  if (!token || String(token).length < 10) return null;
  try {
    var ws = _getSessionSheet(ss);
    var data = ws.getDataRange().getValues();
    var now = new Date();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(token)) {
        var created = data[i][6];
        if (created instanceof Date && (now - created) < 7 * 24 * 60 * 60 * 1000) {
          return { ref: String(data[i][1]), email: String(data[i][2]), nom: String(data[i][3]), role: String(data[i][4]), vieScRole: String(data[i][5]) };
        }
        return null; // Session expirée
      }
    }
  } catch(e) {}
  return null;
}

/**
 * À exécuter UNE SEULE FOIS depuis l'éditeur Apps Script (▶ Exécuter).
 * Met à jour la validation de données de la colonne N (Statut) de l'onglet ETUDIANTS
 * pour autoriser les 5 statuts d'employabilité.
 */
function fixStatutDataValidation() {
  const ss = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");
  const ws = ss.getSheetByName("ETUDIANTS");
  if (!ws) { Logger.log("Onglet ETUDIANTS introuvable."); return; }

  const statuts = ["En recherche", "En Alternance", "En Initial", "En contrat", "Non su"];
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(statuts, true)
    .setAllowInvalid(true)   // "Afficher avertissement" plutôt que "Rejeter" — le backend peut toujours écrire
    .build();

  const lastRow = Math.max(ws.getLastRow(), 2);
  ws.getRange(2, 14, lastRow - 1, 1).setDataValidation(rule);
  Logger.log("Validation mise à jour sur " + (lastRow - 1) + " lignes (colonne N).");
}

// ==========================================
// DISPATCHER PRINCIPAL
// ==========================================
function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const action = params.action;
    const ss = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");
    var p = JSON.parse(e.postData.contents);

    // ── Validation de session ────────────────────────────────────
    // Actions publiques (lecture / login) qui ne nécessitent pas de token
    var PUBLIC_ACTIONS = new Set([
      'login','publicRegister','getDashboardData','getStats','getPartners','getPartnerHistory',
      'getPartnerDetails','getAllAlerts','getStudentDetails','getComEvents','getCampagnes',
      'getEvents','getParticipants','getProspects','getOffres','getHistory','getBilanJournalier',
      'getRapportHebdomadaire','getFormateurs','getPersonnel','getPersonnelList','getPersonnelStaff',
      'getReferentielMatieres','getNotes','getStatsNotes','getClassesInfo','getPlanningTaches',
      'getObjectifsSemaine','getObjectifsStats','getObjectifsStatsSemaines','getSyntheseSemaine',
      'getStatsMatieres','getRefTaches','getQualiopiPreuves','getQualiopiDossierUrl',
      'gedListFolder','gedGetDroits','gedGetStudentDocs','gedGetSubfolder',
      'getMissions','getMailsLog','getMailDetail','getSettings','getStudentsMinList',
      'getHistoriqueStatutNotes','getTropheeData','getActivityStats','getHistoriqueStatutNotes',
      'getSalles','getSimulations','getSuiviEXB','getMMOK','getAllStudentAlerts',
      'getOffreNotes','checkPartnerDuplicate','getClassesList','getPartnerAlerts','getStudentAlerts'
    ]);

    var sessionToken = params.sessionToken || p.sessionToken || "";
    var session = null;
    if (!PUBLIC_ACTIONS.has(action)) {
      session = _validateSession(ss, sessionToken);
      if (sessionToken && !session) {
        // Token fourni mais invalide/expiré → rejeter l'écriture
        return createJsonResponse({ success: false, message: "Session expirée. Veuillez vous reconnecter.", sessionExpired: true });
      }
      if (session) {
        // Enrichir les paramètres avec l'identité validée (remplace l'identité client)
        params.emailAuteur = session.email || params.emailAuteur;
        params.nomAuteur   = session.nom   || params.nomAuteur;
        p.emailAuteur      = session.email || p.emailAuteur;
        p.nomAuteur        = session.nom   || p.nomAuteur;
        // vieScRole : garder celui envoyé par le client si cohérent avec la session
        if (!params.vieScRole && session.vieScRole) params.vieScRole = session.vieScRole;
        if (!p.vieScRole   && session.vieScRole) p.vieScRole = session.vieScRole;
      }
      // Sans token : accepter pour compatibilité mais emailAuteur sera "inconnu" dans les logs
    }
    // ────────────────────────────────────────────────────────────

    switch (action) {
      case "login": return handleLogin(params, ss);
      case "getDashboardData": return handleDashboardData(params, ss);
      case "getStats": return handleGetStats(params, ss);
      case "addStudent": return handleAddStudent(params, ss);
      case "updateStudent": return handleUpdateStudent(params, ss);
      case "changeStudentMdp": return handleChangeStudentMdp(params, ss);
      case "deleteStudent": return handleDeleteStudent(params, ss);
      case "getStudentDetails": return handleStudentDetails(params, ss);
      case "addNote": return handleAddNote(params, ss);
      case "uploadCV": return handleUploadCV(params, ss);
      case "getPartners": return handleGetPartners(params, ss);
      case "getPartnerHistory": return handleGetPartnerHistory(params, ss);
      case "addPartner": return handleAddPartner(params, ss); 
      case "addPartnerNote": return handleAddPartnerNote(params, ss);
      case "setPartnerAlert": return handleSetPartnerAlert(params, ss);
      case "updateContact": return handleUpdateContact(params, ss);
      case "deleteContact": return handleDeleteContact(params, ss);
      case "updateEnterprise": return handleUpdateEnterprise(params, ss);
      case "deleteEnterprise": return handleDeleteEnterprise(params, ss);
      case "linkMAToStudents": return handleLinkMAToStudents(params, ss);
      case "confirmFormerAlternant": return handleConfirmFormerAlternant(params, ss);
      case "removeFormerAlternant": return handleRemoveFormerAlternant(params, ss);
      case "signalFormerAlternant": return handleSignalFormerAlternant(params, ss);
      case "detachStudentFromEnterprise": return handleDetachStudentFromEnterprise(params, ss);
      case "mergePartners": return handleMergePartners(params, ss);
      case "checkPartnerDuplicate": return handleCheckDuplicate(params, ss);
      case "getPartnerDetails": return handleGetPartnerDetails(params, ss);
      
      case "getAllAlerts": return handleGetAllAlerts(params, ss);
      case "completeAlert": return handleCompleteAlert(params, ss);
      case "updateAlert": return handleUpdateAlert(params, ss);
      case "cancelAlert": return handleCancelAlert(params, ss);
      case "testAlert": return handleTestAlert(params, ss);
      case "setStudentAlert": return handleSetStudentAlert(params, ss);
      case "getAllStudentAlerts": return handleGetAllStudentAlerts(params, ss);
      case "completeStudentAlert": return handleCompleteStudentAlert(params, ss);
      case "updateStudentAlert": return handleUpdateStudentAlert(params, ss);
      case "cancelStudentAlert": return handleCancelStudentAlert(params, ss);
      case "getOffres": return handleGetOffres(params, ss);
      case "addOffreNote": return handleAddOffreNote(params, ss);
      case "getOffreNotes": return handleGetOffreNotes(params, ss);
      case "addOffre": return handleAddOffre(params, ss);
      case "updateOffre": return handleUpdateOffre(params, ss);
      case "deleteOffre": return handleDeleteOffre(params, ss);
      case "partagerOffre": return handlePartagerOffre(params, ss);
      case "postulerOffre": return handlePostulerOffre(params, ss);
      case 'aspirerOffres': return handleAspirerOffresGmail(p, ss);
      // --- ROUTES POUR LE MODULE COMMUNICATION ---
      case "getComEvents": return handleGetComEvents(params, ss);
      case "saveComEvent": return handleSaveComEvent(params, ss);
      case "deleteComEvent": return handleDeleteComEvent(params, ss);
      case "previewComMail": return handlePreviewComMail(params, ss);
      case "sendComMail": return handleSendComMail(params, ss);
      case "getClassesList": return handleGetClassesList(ss);
      case "getMailsLog":   return handleGetMailsLog(params, ss);
      case "getMailDetail": return handleGetMailDetail(params, ss);
      
      // --- ROUTES POUR LES ALERTES ---
      case "getPartnerDetails": return handleGetPartnerDetails(params, ss);
      case "getAllAlerts": return handleGetAllAlerts(params, ss);
      case "completeAlert": return handleCompleteAlert(params, ss);
     
      
      
     
      
      
      // --- ROUTES POUR LE MODULE CAMPAGNES (VOTRE ERREUR ÉTAIT ICI) ---
      case "getCampagnes": return handleGetCampagnes(ss);
      case "addCampagne": return handleCreateCampaign(params, ss);
      case "updateCampaign": return handleUpdateCampaign(params, ss);
      case "deleteCampagne": return handleDeleteCampaign(params, ss);
      case "duplicateCampaign": return handleDuplicateCampaign(params, ss);
      case "runCampagne": return handleRunCampagneManual(params, ss);
      case "testCampaign": return handleTestCampaign(params, ss);

      case "getSettings": return handleGetSettings(ss);
      case "saveSettings": return handleSaveSettings(params, ss);
      case "getStudentsMinList": return handleGetStudentsMinList(ss);
      case "getEvents": return handleGetEvents(params, ss);
      case "addEvent": return handleAddEvent(params, ss);
      case "getParticipants": return handleGetParticipants(params, ss);
      case "markPresence": return handleMarkPresence(params, ss);
      case "publicRegister": return handlePublicRegister(params, ss);
      case "getProspects": return handleGetProspects(params, ss);
      case "deleteProspect": return handleDeleteProspect(params, ss);
      case "getTemplateRelance": return handleGetTemplateRelance(params, ss);
      case "saveTemplateRelance": return handleSaveTemplateRelance(params, ss);
      case "testTemplateRelance": return handleTestTemplateRelance(params, ss);

      case 'getHistory': return handleGetHistory(p, ss);
      case 'getBilanJournalier': return handleGetBilanJournalier(p, ss);
      case 'getRapportHebdomadaire': return handleGetRapportHebdomadaire(p, ss);

      // --- ROUTES MISSIONS ---
      case 'getMissions': return handleGetMissions(params, ss);
      case 'updateMissionStatus': return handleUpdateMissionStatus(params, ss);
      case 'addMissionNote': return handleAddMissionNote(params, ss);
      case 'aspirerMissionsNow': return handleAspirerMissionsNow(params, ss);
      case 'updateMissionField': return handleUpdateMissionField(params, ss);
      case 'addMission': return handleAddMission(params, ss);

      // --- ROUTES VIE SCOLAIRE ---
      case 'getReferentielMatieres':     return handleGetReferentielMatieres(p, ss);
      case 'getNotes':                   return handleGetNotes(p, ss);
      case 'saveNote':                   return handleSaveNote(p, ss);
      case 'deleteNote':                 return handleDeleteNote(p, ss);
      case 'updateResultatNote':         return handleUpdateResultatNote(p, ss);
      case 'getStatsNotes':              return handleGetStatsNotes(p, ss);
      case 'getHistoriqueStatutNotes':   return handleGetHistoriqueStatutNotes(p, ss);
      case 'getClassesInfo':             return handleGetClassesInfo(p, ss);
      case 'saveClasseInfo':             return handleSaveClasseInfo(p, ss);
      case 'getFormateurs':              return handleGetFormateurs(p, ss);
      case 'saveFormateur':              return handleSaveFormateur(p, ss);
      case 'deleteFormateur':            return handleDeleteFormateur(p, ss);
      case 'getPlanningTaches':          return handleGetPlanningTaches(p, ss);
      case 'savePlanningTache':          return handleSavePlanningTache(p, ss);
      case 'deletePlanningTache':        return handleDeletePlanningTache(p, ss);
      case 'previewInscriptions':        return handlePreviewInscriptions(p, ss);
      case 'aspirerInscriptions':        return handleAspirerInscriptions(p, ss);
      case 'gedListFolder':              return handleGedListFolder(p, ss);
      case 'gedCreateFolder':            return handleGedCreateFolder(p, ss);
      case 'gedGetDroits':               return handleGedGetDroits(p, ss);
      case 'gedSetDroits':               return handleGedSetDroits(p, ss);
      case 'gedUploadFile':              return handleGedUploadFile(p, ss);
      case 'gedGetStudentDocs':          return handleGedGetStudentDocs(p, ss);
      case 'gedLinkStudentFolder':       return handleGedLinkStudentFolder(p, ss);
      case 'gedUploadStudentDoc':        return handleGedUploadStudentDoc(p, ss);
      case 'gedAutoScanStudents':        return handleGedAutoScanStudents(p, ss);
      case 'gedScanForStudent':          return handleGedScanForStudent(p, ss);
      case 'gedGetSubfolder':            return handleGedGetSubfolder(p, ss);
      case 'getQualiopiPreuves':         return handleGetQualiopiPreuves(p, ss);
      case 'saveQualiopiPreuve':         return handleSaveQualiopiPreuve(p, ss);
      case 'deleteQualiopiPreuve':       return handleDeleteQualiopiPreuve(p, ss);
      case 'saveNotesBulk':              return handleSaveNotesBulk(p, ss);
      case 'uploadReleveNote':           return handleUploadReleveNote(p, ss);
      case 'getPersonnelList':           return handleGetPersonnelList(p, ss);
      case 'uploadQualiopiFile':         return handleUploadQualiopiFile(p, ss);
      case 'getQualiopiDossierUrl':      return handleGetQualiopiDossierUrl(p, ss);
      case 'getRefTaches':               return handleGetRefTaches(p, ss);
      case 'saveRefTache':               return handleSaveRefTache(p, ss);
      case 'getPersonnel':               return handleGetPersonnel(p, ss);
      case 'savePersonnel':              return handleSavePersonnel(p, ss);
      case 'deletePersonnel':            return handleDeletePersonnel(p, ss);
      case 'getObjectifsSemaine':        return handleGetObjectifsSemaine(p, ss);
      case 'saveObjectifSemaine':        return handleSaveObjectifSemaine(p, ss);
      case 'deleteObjectifSemaine':      return handleDeleteObjectifSemaine(p, ss);
      case 'updateObjectifEtat':         return handleUpdateObjectifEtat(p, ss);
      case 'getObjectifsStats':          return handleGetObjectifsStats(p, ss);
      case 'getObjectifsStatsSemaines':  return handleGetObjectifsStatsSemaines(p, ss);
      case 'getSyntheseSemaine':         return handleGetSyntheseSemaine(p, ss);
      case 'saveSyntheseSemaine':        return handleSaveSyntheseSemaine(p, ss);
      case 'getStatsMatieres':           return handleGetStatsMatieres(p, ss);

      // --- ROUTES TROPHEES ---
      case 'migratePlanningWeek': return handleMigratePlanningWeek(p, ss);
      case 'getTropheeData':    return handleGetTropheeData(p, ss);
      case 'saveTropheePrefs':  return handleSaveTropheePrefs(p, ss);
      case 'markTropheesVus':   return handleMarkTropheesVus(p, ss);
      case 'getActivityStats':  return handleGetActivityStats(p, ss);

      // --- ROUTES SALLES & SIMULATIONS ---
      case 'getSalles':         return handleGetSalles(p, ss);
      case 'saveSalle':         return handleSaveSalle(p, ss);
      case 'deleteSalle':       return handleDeleteSalle(p, ss);
      case 'getSimulations':    return handleGetSimulations(p, ss);
      case 'saveSimulation':    return handleSaveSimulation(p, ss);
      case 'deleteSimulation':  return handleDeleteSimulation(p, ss);

      // --- ROUTES SUIVI EXB ---
      case 'getSuiviEXB':       return handleGetSuiviEXB(p, ss);
      case 'saveSuiviEXB':      return handleSaveSuiviEXB(p, ss);
      case 'deleteSuiviEXB':    return handleDeleteSuiviEXB(p, ss);
      case 'createEXBBatch':       return handleCreateEXBBatch(p, ss);
      case 'createEXBBatchByTypo': return handleCreateEXBBatchByTypo(p, ss);
      case 'uploadEXBFile':     return handleUploadEXBFile(p, ss);
      case 'deleteEXBSession':  return handleDeleteEXBSession(p, ss);

      // --- ROUTES MMOK ---
      case 'getMMOK':           return handleGetMMOK(p, ss);
      case 'saveMMOK':          return handleSaveMMOK(p, ss);
      case 'deleteMMOK':        return handleDeleteMMOK(p, ss);

      default:
        return createJsonResponse({ success: false, message: "Action inconnue : " + action });
    } // <--- CETTE ACCOLADE FERME LE SWITCH
  } catch(e) {
    // LE MOUCHARD : Si ça crashe, on force l'écriture de l'erreur dans la case Z1 du 1er onglet
    try {
      var sheetSecours = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
      sheetSecours.getRange("Z1").setValue("🚨 CRASH SERVEUR : " + e.toString());
    } catch(e2) {} 
    
    return createJsonResponse({ success: false, message: "Erreur doPost : " + e.toString() });
  }
}

function doGet(e) {
  if (e.parameter.page === 'inscription') {
    const temp = HtmlService.createTemplateFromFile('inscription');
    temp.eventId = e.parameter.id;
    return temp.evaluate().setTitle("Inscription HECG").addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  return HtmlService.createHtmlOutput("Portail HECG Prive");
}

// ==========================================
// AUTHENTIFICATION & ETUDIANTS
// ==========================================
function handleLogin(params, ss) {
  const id = String(params.identifiant).trim(); const mdp = String(params.password).trim();
  const sP = getSheetSafe(ss, "Personnel"); const dP = sP.getDataRange().getValues();
  for(let i=1; i<dP.length; i++) {
    if((id == String(dP[i][0]) || id == String(dP[i][6])) && mdp == String(dP[i][4])) {
      const email = String(dP[i][6] || "").trim().toLowerCase();
      const nom = String(dP[i][2]||"") + " " + String(dP[i][1]||"");
      const role = String(dP[i][5]||"");
      const vieScRole = String(dP[i][7] || "").trim();
      const sessionToken = _createSession(ss, dP[i][0], email, nom.trim(), role, vieScRole);
      return createJsonResponse({ success: true, role, nom: nom.trim(), ref: dP[i][0], email, vieScRole, sessionToken });
    }
  }
  const sE = getSheetSafe(ss, "ETUDIANTS"); const dE = sE.getDataRange().getValues();
  for(let j=1; j<dE.length; j++) if(id == String(dE[j][0]) && mdp == String(dE[j][1])) return createJsonResponse({ success: true, role: "Etudiant", nom: dE[j][3]+" "+dE[j][2], ref: dE[j][0] });
  return createJsonResponse({ success: false, message: "Identifiants incorrects." });
}

function calculerTypoEtu(entree, sortie, classe) {
  if (sortie && String(sortie).trim() !== "") return "SORTIE";
  if (!classe || String(classe).trim() === "") return "PROSPECT";
  var dEntree = new Date(); var strEntree = String(entree).toLowerCase().trim();
  if (strEntree.includes("rentree")) { var year = strEntree.replace(/[^0-9]/g, ''); if (year.length === 4) dEntree = new Date(year + "-09-01"); }
  else if (entree) { dEntree = new Date(entree); } else { return "PROSPECT"; }
  if (!isNaN(dEntree.getTime())) { var now = new Date(); if (dEntree > now) return "PRÉINSCRIT"; else return "INSCRIT"; }
  return "INSCRIT";
}

function handleDashboardData(params, ss) {
  const role = params.role; const userId = params.userId; 
  const dE = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues();
  const dC = getSheetSafe(ss, "Carnet_route").getDataRange().getValues();
  const dP = getSheetSafe(ss, "Personnel").getDataRange().getValues();
  const dernierEchangeMap = {};
  for (let i = 1; i < dC.length; i++) {
    const etuId = String(dC[i][0]); const dateEch = new Date(dC[i][2]);
    if (!dernierEchangeMap[etuId] || dateEch > dernierEchangeMap[etuId]) dernierEchangeMap[etuId] = dateEch;
  }
  const list = [];
  for (let k = 1; k < dE.length; k++) {
    if (role === "Administrateur" || role === "Communication" || role === "Superviseur") {
      list.push({ 
        id: dE[k][0], nom: dE[k][2], prenom: dE[k][3], mail: dE[k][4], tel: dE[k][5], cv: dE[k][6], id_ent: dE[k][11], id_cont: dE[k][12], 
        statut: dE[k][13], classe: dE[k][14], faitCV: dE[k][15], faitLI: dE[k][16], entree: dE[k][17], sortie: dE[k][18], 
        motif: dE[k][19], typo: dE[k][20], entreprise: dE[k][21], 
        campus: dE[k][22], // <--- NOUVEAU : Lecture Colonne W
        refPerso: dE[k][7], dernierEchange: dernierEchangeMap[dE[k][0]] || null
      });
    }
  }
  const referents = [];
  for(let r=1; r<dP.length; r++) if(dP[r][0]) referents.push({ ref: dP[r][0], nom: dP[r][2] + " " + dP[r][1] });
  return createJsonResponse({ success: true, data: list, referents: referents });
}

function handleAddStudent(p, ss) {
  const ws = getSheetSafe(ss, "ETUDIANTS");
  const id = "ETU" + new Date().getTime().toString().slice(-5);
  
  const typoAAppliquer = p.forceTypo || calculerTypoEtu(p.entree, p.sortie, p.classe);

  const newRow = [
    id, 
    "P@ss123", 
    p.nom, 
    p.prenom, 
    p.mail, 
    p.tel, 
    p.cv, 
    p.refPerso, 
    "", "", "", 
    p.id_ent || "",    // Correction : ID de l'entreprise
    p.id_cont || "",   // Correction : ID du tuteur
    p.statut, 
    p.classe, 
    p.faitCV || "Non", 
    p.faitLI || "Non", 
    p.entree, 
    p.sortie, 
    p.motifSortie, 
    typoAAppliquer,    // Typologie propre
    p.entreprise || "", 
    p.campus || ""
  ];
  ws.appendRow(newRow); 

  logAction(p.idAuteur, p.roleAuteur, "Création", "Dossier étudiant", id, "A créé la fiche de " + p.prenom);

  return createJsonResponse({ success: true });
}





function handleDeleteStudent(p, ss) {
  const ws = getSheetSafe(ss, "ETUDIANTS");
  const d = ws.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (String(d[i][0]) === p.id_etu) {
      const nomEtu = (d[i][3] || "") + " " + (d[i][2] || "");
      const classeEtu = d[i][14] || "?";
      const statutEtu = d[i][13] || "?";
      ws.deleteRow(i + 1);
      const logDetail = "A supprimé la fiche de " + nomEtu.trim() + " (Classe: " + classeEtu + ", Statut: " + statutEtu + ")";
      logAction(p.idAuteur, p.roleAuteur, "Suppression", "Dossier étudiant", p.id_etu, logDetail);
      notifierSuppression(p.idAuteur, p.roleAuteur, "Étudiant supprimé", nomEtu.trim() + " — Classe : " + classeEtu + " — Statut : " + statutEtu + " — ID : " + p.id_etu);
      return createJsonResponse({ success: true });
    }
  }
  return createJsonResponse({ success: false });
}

function handleStudentDetails(params, ss) { 
  const id = params.targetId; 
  const dE = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues(); 
  let prof = null;
  for (let i = 1; i < dE.length; i++) {
    if (String(dE[i][0]).trim() === id) { 
      prof = { id: id, mdp: dE[i][1], nom: dE[i][2], prenom: dE[i][3], mail: dE[i][4], tel: dE[i][5], cv: dE[i][6], refPerso: String(dE[i][7] || "").trim(), statut: dE[i][13], classe: dE[i][14], faitCV: dE[i][15], faitLI: dE[i][16], entree: dE[i][17], sortie: dE[i][18], motif: dE[i][19], typo: dE[i][20], id_ent: dE[i][11], id_cont: dE[i][12], entreprise: dE[i][21], campus: dE[i][22] };
      break;
    }
  }
  if (prof && prof.refPerso) {
    try {
      const dPerso = getSheetSafe(ss, "Personnel").getDataRange().getValues();
      for (let j = 1; j < dPerso.length; j++) {
        if (String(dPerso[j][0] || "").trim() === prof.refPerso) {
          prof.referent = { nom: dPerso[j][1] || "", prenom: dPerso[j][2] || "", mail: dPerso[j][6] || "" };
          break;
        }
      }
    } catch(eRef) {}
  }
  const idUpper = String(id).trim().toUpperCase();
  const isEtudiant = params.requesterRole === "Etudiant";
  const allNotes = getSheetSafe(ss, "Carnet_route").getDataRange().getValues().filter(r => String(r[0]).trim().toUpperCase() === idUpper);
  const notes = allNotes
    .filter(r => !isEtudiant || r[4] !== false)
    .map(r => ({ auteur: r[1], date: r[2], texte: r[3], visible: r[4] !== false }));

  // --- NOUVEAU : RÉCUPÉRATION DES OFFRES PARTAGÉES ---
  const dP = getSheetSafe(ss, "PARTAGES").getDataRange().getValues();
  const dO = getSheetSafe(ss, "OFFRES").getDataRange().getValues();
  const offersMap = {};
  for(let j=1; j<dO.length; j++) {
    offersMap[String(dO[j][0])] = { poste: dO[j][14], ent: dO[j][4], url: dO[j][1], qualite: dO[j][13] };
  }
  const studentOffers = dP.filter(r => String(r[0]) === id).map(r => {
    const info = offersMap[String(r[1])] || {};
    return { id: r[1], poste: info.poste || "Inconnu", entreprise: info.ent || "-", url: info.url, etat: r[3], qualite: info.qualite };
  }).filter(o => o.qualite !== "Obsolète"); // Cache les périmées

  // Historique entreprises depuis ANCIENS_ALTERNANTS
  const alternanceHistory = [];
  try {
    const wsA = ss.getSheetByName("ANCIENS_ALTERNANTS");
    if (wsA && wsA.getLastRow() > 1) {
      const dA = wsA.getDataRange().getValues();
      const dPart = getSheetSafe(ss, "PARTENARIAT").getDataRange().getValues();
      // Map entreprise : idEnt → {nom, contact}
      const entMap = {}, contactMap = {};
      for (let i = 1; i < dPart.length; i++) {
        const eid = String(dPart[i][1]||"").trim().toUpperCase();
        if (eid && !entMap[eid])  entMap[eid] = String(dPart[i][2]||"").trim();
        if (eid && !contactMap[eid]) contactMap[eid] = {
          nom:   String(dPart[i][4]||"")+" "+String(dPart[i][3]||""),
          tel:   String(dPart[i][5]||"").trim(),
          mail:  String(dPart[i][6]||"").trim(),
          poste: String(dPart[i][13]||"").trim()
        };
      }
      for (let k = 1; k < dA.length; k++) {
        if (String(dA[k][2]||"").trim() !== id) continue;
        const status = String(dA[k][6]||"").trim();
        if (status === "Erreur") continue;
        const idEnt = String(dA[k][1]||"").trim().toUpperCase();
        // Nom entreprise : priorité colonne 7 (stockée), sinon lookup PARTENARIAT
        const nomEnt = String(dA[k][7]||"").trim() || entMap[idEnt] || idEnt;
        const contact = contactMap[idEnt] || null;
        let dateStr = "";
        if (dA[k][5]) {
          try { dateStr = Utilities.formatDate(new Date(dA[k][5]), Session.getScriptTimeZone(), "dd/MM/yyyy"); } catch(e) { dateStr = String(dA[k][5]); }
        }
        alternanceHistory.push({
          id:          String(dA[k][0]),
          entreprise:  nomEnt,
          idEntreprise: idEnt,
          dateDepart:  dateStr,
          status:      status,
          classe:      String(dA[k][8]||"").trim(),
          campus:      String(dA[k][9]||"").trim(),
          motif:       String(dA[k][10]||"").trim(),
          contact:     contact ? contact.nom.trim() : "",
          contactTel:  contact ? contact.tel : "",
          contactMail: contact ? contact.mail : "",
          contactPoste:contact ? contact.poste : ""
        });
      }
    }
  } catch(eH) {}

  return createJsonResponse({ success: true, profil: prof, notes: notes.reverse(), offres: studentOffers, alternanceHistory: alternanceHistory });
}

function handleAddNote(p, ss) {
  const visibleEtudiant = p.visibleEtudiant !== false;
  getSheetSafe(ss, "Carnet_route").appendRow([p.targetId, p.idAuteur || "Inconnu", new Date(), p.contenu, visibleEtudiant]);
  const apercu = p.contenu ? String(p.contenu).substring(0, 80) + (p.contenu.length > 80 ? "…" : "") : "";
  logAction(p.idAuteur, p.roleAuteur, "Création", "Carnet de route", p.targetId, "Note ajoutée : « " + apercu + " »");

  // Si c'est l'étudiant lui-même qui écrit, on notifie son référent par mail
  if (p.isStudentNote) {
    try {
      const dE = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues();
      const dP = getSheetSafe(ss, "Personnel").getDataRange().getValues();
      let etuRow = null;
      for (let i = 1; i < dE.length; i++) {
        if (String(dE[i][0]).trim() === String(p.targetId).trim()) { etuRow = dE[i]; break; }
      }
      if (etuRow) {
        const prenomEtu = etuRow[3], nomEtu = etuRow[2];
        const refPerso = String(etuRow[7] || "").trim();
        let mailRef = null;
        for (let j = 1; j < dP.length; j++) {
          const staffId = String(dP[j][0] || "").trim();
          const staffNom = (String(dP[j][2] || "") + " " + String(dP[j][1] || "")).trim();
          const staffEmail = String(dP[j][6] || "").trim();
          if ((staffId === refPerso || staffNom.toLowerCase() === refPerso.toLowerCase()) && staffEmail.includes("@")) {
            mailRef = staffEmail; break;
          }
        }
        if (mailRef) {
          const subject = "[CRM HECG] Nouvelle mise à jour de " + prenomEtu + " " + nomEtu;
          const body = `<div style="font-family:Arial,sans-serif;padding:20px;border:1px solid #eee;border-radius:10px;">
            <h2 style="color:#1A4E8A;">📝 Mise à jour étudiant</h2>
            <p><b>${prenomEtu} ${nomEtu}</b> a ajouté une note dans son carnet de route :</p>
            <div style="background:#f9f9f9;padding:15px;border-left:4px solid #F26522;border-radius:5px;margin:15px 0;">
              ${String(p.contenu).replace(/\n/g,'<br>')}
            </div>
            <p style="font-size:11px;color:#999;">Consultez le dossier étudiant dans le CRM HECG pour plus de détails.</p>
          </div>`;
          GmailApp.sendEmail(mailRef, subject, "", { htmlBody: body });
          try { logMailDetail(SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ"), { auteur: p.idAuteur || "Étudiant", categorie: "Échange étudiant", sousCat: "Carnet de route — notif. référent", objet: subject, destinataires: [mailRef], corps: body, sourceId: p.targetId || "" }); } catch(eLog) {}
        }
      }
    } catch(eNotif) { console.error("Erreur notif référent : " + eNotif.toString()); }
  }

  // Si c'est un responsable qui écrit, on notifie l'étudiant par mail
  if (!p.isStudentNote) {
    try {
      const dE = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues();
      const dP = getSheetSafe(ss, "Personnel").getDataRange().getValues();

      // 1. Trouver l'email de l'étudiant cible
      let mailEtu = null;
      for (let i = 1; i < dE.length; i++) {
        if (String(dE[i][0]).trim() === String(p.targetId).trim()) {
          mailEtu = String(dE[i][4] || "").trim();
          break;
        }
      }

      // 2. Trouver le nom complet + email du responsable auteur
      let nomRefComplet = p.idAuteur || "Votre référent";
      let mailRefComplet = null;
      for (let j = 1; j < dP.length; j++) {
        const staffNom = (String(dP[j][2] || "") + " " + String(dP[j][1] || "")).trim();
        const staffEmail = String(dP[j][6] || "").trim();
        if (staffNom.toLowerCase() === (p.idAuteur || "").toLowerCase() && staffEmail.includes("@")) {
          nomRefComplet = staffNom;
          mailRefComplet = staffEmail;
          break;
        }
      }

      if (mailEtu && mailEtu.includes("@")) {
        const referentInfo = mailRefComplet
          ? `${nomRefComplet} (<a href="mailto:${mailRefComplet}">${mailRefComplet}</a>)`
          : nomRefComplet;
        const subject = "[HECG] Nouvelle note dans votre carnet d'échange";
        const body = `<div style="font-family:Arial,sans-serif;padding:20px;border:1px solid #eee;border-radius:10px;max-width:560px;">
          <h2 style="color:#1A4E8A;margin-top:0;">📝 Carnet d'échange HECG</h2>
          <p>Votre référent <strong>${referentInfo}</strong> a ajouté une note à votre carnet d'échange.</p>
          <p style="margin-top:20px;">Si vous n'avez pas accès à votre espace étudiant, demandez-lui directement vos identifiants.</p>
          <hr style="border:none;border-top:1px solid #eee;margin:20px 0;">
          <p style="font-size:11px;color:#aaa;">HECG — Espace étudiant. Ne pas répondre directement à cet e-mail.</p>
        </div>`;
        GmailApp.sendEmail(mailEtu, subject, "", { htmlBody: body, from: SENDER_EMAIL });
        try { logMailDetail(SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ"), { auteur: p.idAuteur || "Référent", categorie: "Échange étudiant", sousCat: "Carnet de route — notif. étudiant", objet: subject, destinataires: [mailEtu], corps: body, sourceId: p.targetId || "" }); } catch(eLog) {}
      }
    } catch(eNotif) { console.error("Erreur notif etudiant : " + eNotif.toString()); }
  }

  return createJsonResponse({ success: true });
}

function handleUploadCV(p, ss) { 
  const folderId = (p.userId === 'FLYER' || p.userId === 'PJ') ? FOLDER_ID_COM : FOLDER_ID_CV;
  const file = DriveApp.getFolderById(folderId).createFile(Utilities.newBlob(Utilities.base64Decode(p.base64Data), p.mimeType, p.fileName)); 
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); 
  if (p.userId !== 'FLYER' && p.userId !== 'PJ') {
    const ws = getSheetSafe(ss, "ETUDIANTS"); const d = ws.getDataRange().getValues(); 
    for (let i = 1; i < d.length; i++) if (String(d[i][0]) === p.userId) { ws.getRange(i+1, 7).setValue(file.getUrl()); break; }
  }
  const typeDoc = (p.userId === 'FLYER' || p.userId === 'PJ') ? "Document com" : "CV";
  logAction(p.idAuteur, p.roleAuteur, "Création", "Documents", p.userId || p.targetId, "Fichier uploadé : " + p.fileName + " (" + typeDoc + ")");
  return createJsonResponse({ success: true, url: file.getUrl() });
}

// ==========================================
// STATISTIQUES & B2B
// ==========================================
function handleGetStats(params, ss) {
  const dE = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues(); 
  const dP = getSheetSafe(ss, "Partenariat").getDataRange().getValues();
  const dO = getSheetSafe(ss, "OFFRES").getDataRange().getValues(); 
  const stats = { inscrits: [], preinscrits: [], prospects: [], sorties: [], partenaires_avec: [], partenaires_sans: [], offres_actives: [], offres_non_partagees: [], offres_non_postulees: [] };
  const entAccueil = new Set(); const today = new Date(); today.setHours(0,0,0,0);
  for (let i = 1; i < dE.length; i++) {
    if (!dE[i][0]) continue;
    const typo = String(dE[i][20] || "PROSPECT").trim();
    const typoLow = typo.toLowerCase();
    const aAlt = String(dE[i][13]).toLowerCase().includes("alternance");
    if (aAlt && dE[i][21]) entAccueil.add(String(dE[i][21]).toLowerCase().trim());

    const etuObj = { id: dE[i][0], nom: dE[i][2] + " " + dE[i][3], classe: dE[i][14], aAlternance: aAlt, motif: dE[i][19], email: dE[i][4], campus: dE[i][22] };

    if (typoLow.includes("sorti") || typoLow.includes("refus")) stats.sorties.push(etuObj);
    else if (typoLow === "prospect") stats.prospects.push(etuObj);
    else if (typoLow.includes("préin")) stats.preinscrits.push(etuObj);
    else stats.inscrits.push(etuObj);
  }
  const ents = {};
  for (let j = 1; j < dP.length; j++) if (dP[j][2]) {
    const n = String(dP[j][2]).toLowerCase().trim();
    if (!ents[n]) ents[n] = { nom: dP[j][2], nbContacts: 1 }; else ents[n].nbContacts++;
  }
  for (const k in ents) { if (entAccueil.has(k)) stats.partenaires_avec.push(ents[k]); else stats.partenaires_sans.push(ents[k]); }
  return createJsonResponse({ success: true, data: stats });
}

function handleGetPartners(params, ss) {
  const wsP = getSheetSafe(ss, "Partenariat"); const wsE = getSheetSafe(ss, "ETUDIANTS");
  const dP = wsP.getDataRange().getValues(); const dE = wsE.getDataRange().getValues();
  const entreprises = {};
  for (let i = 1; i < dP.length; i++) {
    const idEnt = String(dP[i][1] || "").trim(); if (!idEnt) continue;
    if (!entreprises[idEnt]) entreprises[idEnt] = { id: idEnt, nom: dP[i][2], adresse: String(dP[i][10] || "").trim(), cp: String(dP[i][11] || "").trim(), ville: String(dP[i][12] || "").trim(), contacts: [], alternants: [], anciensAlternants: [] };
    entreprises[idEnt].contacts.push({ id_contact: dP[i][0], nom: dP[i][3], prenom: dP[i][4], tel: dP[i][5], email: dP[i][6], tel2: dP[i][9], mail2: dP[i][8], poste: String(dP[i][13] || "").trim(), is_ma: String(dP[i][14] || "Non").trim() });
  }
  for (let j = 1; j < dE.length; j++) {
    const idEntEtu = String(dE[j][11]).trim();
    if (entreprises[idEntEtu] && String(dE[j][13]).toLowerCase().includes("alternance")) entreprises[idEntEtu].alternants.push({ nom: dE[j][2], prenom: dE[j][3] });
  }
  // Anciens alternants (changements d'entreprise détectés)
  const wsA = ss.getSheetByName("ANCIENS_ALTERNANTS");
  if (wsA && wsA.getLastRow() > 1) {
    const dA = wsA.getDataRange().getValues();
    for (let k = 1; k < dA.length; k++) {
      const idEnt = String(dA[k][1] || "").trim();
      const status = String(dA[k][6] || "").trim();
      if (entreprises[idEnt] && status !== "Erreur") {
        entreprises[idEnt].anciensAlternants.push({ id: String(dA[k][0]), nom: dA[k][3], prenom: dA[k][4], date: dA[k][5], status: status });
      }
    }
  }
  // Dernier échange par entreprise (SUIVI_PARTENARIAT) — comparaison insensible à la casse
  try {
    const allSheets = ss.getSheets();
    const wsSuivi = allSheets.find(s => s.getName().toUpperCase() === "SUIVI_PARTENARIAT") || null;
    if (wsSuivi && wsSuivi.getLastRow() > 1) {
      const dSuivi = wsSuivi.getDataRange().getValues();
      // Index entreprises en majuscules pour matching souple
      const entUpperMap = {};
      Object.keys(entreprises).forEach(k => { entUpperMap[k.toUpperCase()] = k; });
      const lastMap = {};
      for (let s = 1; s < dSuivi.length; s++) {
        const rawId = String(dSuivi[s][0] || "").trim(); if (!rawId) continue;
        const keyUp = rawId.toUpperCase();
        const noteDate = dSuivi[s][2];
        const t = noteDate instanceof Date ? noteDate.getTime() : new Date(noteDate).getTime();
        if (!lastMap[keyUp] || t > lastMap[keyUp].t) {
          lastMap[keyUp] = { t, date: noteDate, auteur: String(dSuivi[s][1] || ""), texte: String(dSuivi[s][3] || "") };
        }
      }
      Object.keys(lastMap).forEach(keyUp => {
        const realKey = entUpperMap[keyUp]; if (!realKey) return;
        const e = lastMap[keyUp];
        const df = e.date instanceof Date ? Utilities.formatDate(e.date, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : String(e.date);
        entreprises[realKey].dateEchange = df;
        entreprises[realKey].contenuEchange = (e.auteur ? e.auteur + " : " : "") + e.texte;
      });
    }
  } catch(eSuivi) {}
  return createJsonResponse({ success: true, data: Object.values(entreprises) });
}

function _getOrCreateAnciensSheet(ss) {
  let s = ss.getSheetByName("ANCIENS_ALTERNANTS");
  if (!s) {
    s = ss.insertSheet("ANCIENS_ALTERNANTS");
    s.getRange(1, 1, 1, 11).setValues([["ID","ID_Entreprise","ID_Etudiant","Nom","Prenom","Date_Depart","Status","Nom_Entreprise","Classe","Campus","Motif"]]);
    s.setFrozenRows(1);
  }
  return s;
}

// Enregistre un lien historique étudiant ↔ entreprise (évite les doublons actifs)
function _saveAncienLien(ss, data) {
  var ws = _getOrCreateAnciensSheet(ss);
  var idEtu = String(data.idEtu||"").trim();
  var idEnt = String(data.idEnt||"").trim().toUpperCase();
  if (!idEtu || !idEnt) return false;
  var existing = ws.getDataRange().getValues();
  var doublon = existing.slice(1).some(function(r){
    return String(r[1]).trim().toUpperCase() === idEnt
        && String(r[2]).trim() === idEtu
        && String(r[6]) !== "Erreur";
  });
  if (doublon) return false;
  ws.appendRow([
    "ANC-" + Date.now(),
    idEnt,
    idEtu,
    String(data.nom||""),
    String(data.prenom||""),
    new Date(),
    "Confirmé",
    String(data.nomEnt||""),
    String(data.classe||""),
    String(data.campus||""),
    String(data.motif||"")
  ]);
  return true;
}

function handleSignalFormerAlternant(p, ss) {
  try {
    const s = _getOrCreateAnciensSheet(ss);
    const idEnt = String(p.idEntreprise || "").trim().toUpperCase();
    const idEtu = String(p.idEtudiant || "").trim().toUpperCase();

    // Vérifier doublon
    const existing = s.getDataRange().getValues().slice(1);
    const doublon = existing.some(r => String(r[1]).trim().toUpperCase() === idEnt && String(r[2]).trim() === idEtu && String(r[6]) !== "Erreur");
    if (doublon) return createJsonResponse({ success: false, message: "Un mouvement est déjà en cours pour cet alternant." });

    // Récupérer classe, campus, nomEnt depuis ETUDIANTS
    const sheetEtu = getSheetSafe(ss, "ETUDIANTS");
    const dataEtu = sheetEtu.getDataRange().getValues();
    let rowEtu = -1, classe = "", campus = "";
    for (let i = 1; i < dataEtu.length; i++) {
      if (String(dataEtu[i][0]).trim() === idEtu) {
        rowEtu = i + 1;
        classe = String(dataEtu[i][14] || "").trim();
        campus = String(dataEtu[i][22] || "").trim();
        break;
      }
    }

    // Nom de l'entreprise depuis PARTENARIAT
    let nomEnt = "";
    try {
      const dPart = getSheetSafe(ss, "PARTENARIAT").getDataRange().getValues();
      for (let i = 1; i < dPart.length; i++) {
        if (String(dPart[i][1]).trim().toUpperCase() === idEnt) { nomEnt = String(dPart[i][2] || "").trim(); break; }
      }
    } catch(e2) {}

    // Créer l'entrée ANCIENS_ALTERNANTS (status Confirmé : action intentionnelle depuis dossier entreprise)
    s.appendRow(["ANC-" + Date.now(), idEnt, idEtu, p.nom || "", p.prenom || "", new Date(), "Confirmé", nomEnt, classe, campus, "Départ signalé depuis dossier entreprise"]);

    // Mise à jour bilatérale : retirer l'entreprise du dossier étudiant
    if (rowEtu !== -1) {
      sheetEtu.getRange(rowEtu, 12).setValue("");
      sheetEtu.getRange(rowEtu, 13).setValue("");
      sheetEtu.getRange(rowEtu, 22).setValue("");
      logAction(p.idAuteur, p.roleAuteur, "Modification", "Dossier étudiant", idEtu, "Entreprise « " + nomEnt + " » retirée automatiquement (départ signalé depuis dossier entreprise)");
    }

    logAction(p.idAuteur, p.roleAuteur, "Création", "AncienAlternant", idEtu, (p.prenom||"") + " " + (p.nom||"") + " — départ de « " + nomEnt + " » enregistré");
    return createJsonResponse({ success: true });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleConfirmFormerAlternant(p, ss) {
  const s = _getOrCreateAnciensSheet(ss);
  const data = s.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(p.id)) { s.getRange(i + 1, 7).setValue("Confirmé"); return createJsonResponse({ success: true }); }
  }
  return createJsonResponse({ success: false, message: "Entrée introuvable" });
}

function handleRemoveFormerAlternant(p, ss) {
  const s = _getOrCreateAnciensSheet(ss);
  const data = s.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(p.id)) { s.getRange(i + 1, 7).setValue("Erreur"); return createJsonResponse({ success: true }); }
  }
  return createJsonResponse({ success: false, message: "Entrée introuvable" });
}

// ==========================================
// OFFRES
// ==========================================
function handleGetOffres(p, ss) {
  try {
    const dO = getSheetSafe(ss, "OFFRES").getDataRange().getValues();
    const dP = getSheetSafe(ss, "PARTAGES").getDataRange().getValues();
    const dE = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues();

    // Derniere note tracee par offre (SUIVI_OFFRES)
    const lastNoteMap = {};
    try {
      const dSO = _getSuiviOffresSheet(ss).getDataRange().getValues();
      for (let s = 1; s < dSO.length; s++) {
        const idO = String(dSO[s][1] || "").trim().toUpperCase();
        if (!idO) continue;
        const d = dSO[s][3];
        const t = d instanceof Date ? d.getTime() : new Date(d).getTime();
        if (!lastNoteMap[idO] || t > lastNoteMap[idO].t) {
          lastNoteMap[idO] = { t, auteur: String(dSO[s][2] || ""), date: d instanceof Date ? Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : String(d || ""), texte: String(dSO[s][4] || "") };
        }
      }
    } catch(eSO) {}

    // Dictionnaire ultra-precis : on indexe par ID ET par Email
    const etuMap = {};
    for(let j=1; j<dE.length; j++) {
      let id = String(dE[j][0] || "").trim().toUpperCase();
      let mail = String(dE[j][4] || "").trim().toLowerCase();
      let nomComplet = (dE[j][3] || "") + " " + (dE[j][2] || "");
      if(id) etuMap[id] = nomComplet;
      if(mail) etuMap[mail] = nomComplet;
    }

    const res = [];
    for (let i = 1; i < dO.length; i++) {
      if (dO[i][0]) {
        const idOffre = String(dO[i][0]).trim();
        const sharedWith = dP
          .filter(r => String(r[1]).trim() === idOffre)
          .map(r => {
             let key = String(r[0]).trim();
             return etuMap[key.toUpperCase()] || etuMap[key.toLowerCase()] || ("Inconnu (" + key + ")");
          });

        const lastNote = lastNoteMap[idOffre.toUpperCase()] || null;
        res.push({
          id: idOffre, url: dO[i][1], date: dO[i][2], plateforme: dO[i][3],
          entreprise: dO[i][4], ville: String(dO[i][5] || "").trim(), missions: dO[i][6],
          qualite: dO[i][9],
          etat: dO[i][13],
          poste: dO[i][14],
          destinataires: sharedWith,
          dernierNote: lastNote
        });
      }
    }
    return createJsonResponse({ success: true, data: res });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}
function handleAddOffre(p, ss) { 
  getSheetSafe(ss, "OFFRES").appendRow([
    "OFR" + new Date().getTime().toString().slice(-6), // A : ID
    p.url,              // B : Lien
    new Date(),         // C : Date
    p.plateforme,       // D : Plateforme
    p.entreprise,       // E : Entreprise
    p.ville || "",      // F : Ville
    p.missions,         // G : Missions
    "",                 // H
    "",                 // I
    p.qualite || "À qualifier", // J : Qualité_offre (Index 9)
    "",                 // K
    "",                 // L
    "",                 // M
    "Non partagée",              // N : Etat_offre (Index 13) — géré automatiquement par handlePartagerOffre
    p.poste             // O : Poste
  ]); 

  logAction(p.idAuteur, p.roleAuteur, "Création", "Offres", "Nouvelle Offre", "A ajouté l'offre : " + p.poste + " chez " + p.entreprise + " (" + (p.plateforme || "?") + ") — " + (p.ville || "Ville ?") + " — Statut: À qualifier");
  return createJsonResponse({ success: true });
}



// --- MISE À JOUR D'UNE OFFRE EXISTANTE ---
function handleUpdateOffre(p, ss) {
  try {
    const sheet = getSheetSafe(ss, "OFFRES");
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(p.id)) {
        const oldRow = data[i];

        // Construire le détail du log avec comparaison avant/après
        const changes = [];
        if ((oldRow[9] || "À qualifier") !== (p.qualite || "")) changes.push("Statut: " + (oldRow[9] || "À qualifier") + " → " + p.qualite);
        if ((oldRow[13] || "Non partagée") !== (p.etat || "") && p.etat) changes.push("État: " + (oldRow[13] || "Non partagée") + " → " + p.etat);
        if ((oldRow[14] || "") !== (p.poste || "")) changes.push("Poste: " + (oldRow[14] || "-") + " → " + p.poste);
        if ((oldRow[4] || "") !== (p.entreprise || "")) changes.push("Entreprise: " + (oldRow[4] || "-") + " → " + p.entreprise);
        if ((oldRow[5] || "") !== (p.ville || "")) changes.push("Ville: " + (oldRow[5] || "-") + " → " + p.ville);

        // Tracer la modification des notes internes dans SUIVI_OFFRES
        const oldMissions = String(oldRow[6] || "").trim();
        const newMissions = String(p.missions || "").trim();
        if (newMissions && newMissions !== oldMissions) {
          try {
            const wsNote = _getSuiviOffresSheet(ss);
            wsNote.appendRow(["ON-" + Date.now(), String(p.id).trim(), p.idAuteur || "Inconnu", new Date(), "📝 Notes internes : " + newMissions]);
          } catch(eN) {}
          changes.push("Notes internes modifiées");
        }

        // 1. Mise à jour des informations classiques
        sheet.getRange(i + 1, 2).setValue(p.url);
        sheet.getRange(i + 1, 4).setValue(p.plateforme);
        sheet.getRange(i + 1, 5).setValue(p.entreprise);
        sheet.getRange(i + 1, 6).setValue(p.ville || "");
        sheet.getRange(i + 1, 7).setValue(p.missions);
        sheet.getRange(i + 1, 15).setValue(p.poste);
        sheet.getRange(i + 1, 10).setValue(p.qualite);
        if (p.etat) sheet.getRange(i + 1, 14).setValue(p.etat);

        const logDetail = changes.length > 0
          ? "Offre " + (oldRow[14] || p.poste) + " | " + changes.join(" | ")
          : "Mise à jour offre " + (p.poste || p.id);
        logAction(p.idAuteur, p.roleAuteur, "Modification", "Offres", p.id, logDetail);
        return createJsonResponse({ success: true, message: "Offre mise à jour !" });
      }
    }
    return createJsonResponse({ success: false, message: "Offre introuvable dans le tableur." });
  } catch (e) {
    return createJsonResponse({ success: false, message: e.toString() });
  }
}
// ==========================================
// COMMUNICATION & CAMPAGNES
// ==========================================
function handleGetCampagnes(ss) {
  try {
    const sheet = getSheetSafe(ss, "CAMPAGNES");
    const data = sheet.getDataRange().getValues();
    const campaigns = [];
    for (let i = 1; i < data.length; i++) {
      campaigns.push({
        id: data[i][0], nom: data[i][1], frequence: data[i][2],
        dateCreation: data[i][3], dernierEnvoi: data[i][4],
        objet: data[i][5], message: data[i][6], statut: data[i][7], cibles: data[i][8]
      });
    }
    return createJsonResponse({ success: true, data: campaigns });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleGetSettings(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = getSheetSafe(ss, "CONFIG").getDataRange().getValues();
  let emailTest = "mokhtarimehdi.pro@gmail.com";
  for(let i=1; i<data.length; i++) if(String(data[i][0]).trim() === "email_test") emailTest = data[i][1]; 
  return createJsonResponse({ success: true, email_test: emailTest });
}

function handleSaveSettings(p, ss) {
  const ws = getSheetSafe(ss, "CONFIG"); const data = ws.getDataRange().getValues();
  for(let i=1; i<data.length; i++) if(String(data[i][0]).trim() === "email_test") { ws.getRange(i+1, 2).setValue(p.email_test); return createJsonResponse({ success: true }); }
  ws.appendRow(["email_test", p.email_test]); return createJsonResponse({ success: true });
}

function handleSaveCampaign(p, ss) {
  const ws = getSheetSafe(ss, "CAMPAGNES");
  const isNew = !p.id;
  const id = isNew ? "CMP" + new Date().getTime() : p.id;
  
  // On s'assure que la cible est bien une chaîne de caractères
  const cibleStr = typeof p.cible === 'object' ? JSON.stringify(p.cible) : p.cible;

  // Structure des colonnes (14 au total) :
  // A:ID, B:Nom, C:Objet, D:Corps/Accroche, E:Cible, F:Type, G:DernierEnvoi, 
  // H:Statut, I:PJ_URL, J:Flyer_URL, K:LienExt, L:Campus, M:DateEvt, N:Horaires
  
  const rowData = [
    id, p.nom, p.objet, "", cibleStr, p.type, 
    p.dernierEnvoi || "", p.statut || "Brouillon", p.pjUrl || "", p.flyerUrl || "", 
    p.lienExterne || "", p.campus, p.dateEvt, p.horaires, p.emailsLibres || "", 
    p.accroche || "" // <-- ACCROCHE BIEN PLACÉE EN 16ÈME POSITION (COLONNE P)
  ];

  if (isNew) {
    ws.appendRow(rowData);
  } else {
    const data = ws.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === id) {
        // On met à jour de la colonne B (index 2) à N (index 14) -> soit 13 colonnes
        // On ne touche pas à l'ID (colonne A)
        const updateData = rowData.slice(1); 
        ws.getRange(i + 1, 2, 1, updateData.length).setValues([updateData]);
        break;
      }
    }
  }
  return createJsonResponse({ success: true, message: "Campagne sauvegardée avec succès." });
}

function handleDeleteCampaign(p, ss) {
  try {
    const sheet = getSheetSafe(ss, "CAMPAGNES");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === p.id) {
        sheet.getRange(i + 1, 8).setValue("Archivé"); // On change le statut en H
        break;
      }
    }
    return createJsonResponse({ success: true, message: "Campagne rangée dans l'historique." });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

/**
 * MOTEUR D'ENVOI HECG - ZERO EMOJI & LIEN INSCRIPTION PROSPECTS
 */
/**
 * ÉTAPE 2 : MOTEUR D'ENVOI INTELLIGENT HECG
 */
function handleSendComMail(p, ss) {
  const isTest = p.isTest;
  const cmpId = p.id;
  const type = p.type;
  const subject = p.objet;
  
  // ON PREND DIRECTEMENT LA LISTE VALIDÉE PAR L'UTILISATEUR
  let emailsArray = [];
  if (isTest) {
    emailsArray.push(p.testEmail);
  } else {
    emailsArray = p.finalEmails || [];
  }

  if (emailsArray.length === 0) return createJsonResponse({ success: false, message: "Aucun destinataire validé pour l'envoi." });



  // 2. PRÉPARATION DU CONTENU HTML (SANS EMOJI)
  const addrMarseille = "39 Rue Sainte Cecile, 13005 Marseille";
  const addrAix = "Le Mansard B, 2 Pl. Romee de Villeneuve, 13090 Aix-en-Provence";
  const campusAddr = (p.campus === "Aix") ? addrAix : addrMarseille;
  const regURL = p.lienExterne || (ScriptApp.getService().getUrl() + "?page=inscription&id=" + cmpId);
  
  let htmlBody = "";
  if (type === "Atelier CV") {
    htmlBody = getTemplateCV().replace(/{{INSERER_ICI_ADRESSE_CAMPUS}}/g, campusAddr).replace(/{{VOTRE_LIEN_D_INSCRIPTION_ICI}}/g, regURL);
  } else if (type === "Atelier LinkedIn") {
    htmlBody = getTemplateLinkedIn().replace(/{{INSERER_ICI_ADRESSE_CAMPUS}}/g, campusAddr).replace(/{{VOTRE_LIEN_D_INSCRIPTION_ICI}}/g, regURL);
  } else if (type === "Atelier Mixte") {
    htmlBody = getTemplateMixte().replace(/{{INSERER_ICI_ADRESSE_CAMPUS}}/g, campusAddr).replace(/{{VOTRE_LIEN_D_INSCRIPTION_ICI}}/g, regURL);
  } else if (type === "JPO") {
    const infoJPO = (p.dateEvt ? new Date(p.dateEvt).toLocaleDateString('fr-FR') : "") + " | " + (p.horaires || "") + " | Campus " + (p.campus || "Marseille");
    htmlBody = getTemplateJPO().replace(/{{INSERER_DATE_ICI}}/g, infoJPO).replace(/{{INSERER_HORAIRES_ICI}}/g, p.horaires || "").replace(/{{INSERER_CAMPUS_ET_ADRESSE_ICI}}/g, campusAddr).replace(/{{LIEN_INSCRIPTION_JPO_ICI}}/g, regURL);
  } else {
    // --- LOGIQUE POUR LE TEMPLATE SALON ---
    htmlBody = getTemplateSalon();

    // 1. Texte de secours (si le champ accroche est vide dans l'interface)
    const texteSecours = "Choisir sa formation, c'est avant tout une rencontre humaine. Nos équipes se déplacent pour échanger avec vous, comprendre vos ambitions et dessiner ensemble votre futur parcours en alternance.";

    // 2. On choisit le texte (Accroche ou Secours)
    let texteFinal = (p.accroche && p.accroche.trim() !== "") 
                     ? p.accroche.replace(/\n/g, '<br>') 
                     : texteSecours;

    // 3. On fait les remplacements
    htmlBody = htmlBody
      .replace(/{{TEXTE_ACCROCHE}}/g, texteFinal)
      .replace(/{{NOM_DU_SALON_OU_EVENEMENT}}/g, p.nom || "")
      .replace(/{{DATE_ET_HORAIRES}}/g, p.dateEvt ? new Date(p.dateEvt).toLocaleDateString('fr-FR') : "")
      .replace(/{{LIEU_ET_ADRESSE}}/g, p.lieuLibre || campusAddr)
      .replace(/{{LIEN_INSCRIPTION_ICI}}/g, regURL);
  }

  // Insertion du Flyer si présent
  if (p.flyerUrl) {
    htmlBody = `<div style="text-align:center; margin-bottom:20px;"><img src="${p.flyerUrl}" style="max-width:100%; border-radius:10px;"></div>` + htmlBody;
  }

  // 3. ENVOI PAR PAQUETS (ANTI-SPAM)
  const options = { bcc: "", htmlBody: htmlBody + getSignatureHTML(), name: "HECG - Ecole de Gestion", body: "Merci de consulter ce mail en mode HTML pour voir l'invitation HECG." };
  if (p.pjUrl) { try { let fid = p.pjUrl.match(/id=([^&]+)/) || p.pjUrl.match(/d\/([^/]+)/); if(fid) options.attachments = [DriveApp.getFileById(fid[1]).getBlob()]; } catch(e){} }

  const chunkSize = 50;
  let sentCount = 0;
  let sentEmails = [];
  const errorsList = [];

  for (let i = 0; i < emailsArray.length; i += chunkSize) {
    const chunk = emailsArray.slice(i, i + chunkSize);
    try {
      options.bcc = chunk.join(',');
      GmailApp.sendEmail(Session.getActiveUser().getEmail(), subject, "", options);
      sentCount += chunk.length;
      sentEmails = sentEmails.concat(chunk);
    } catch(eChunk) {
      errorsList.push("Lot " + (Math.floor(i/chunkSize)+1) + " (" + chunk.length + " dest.) : " + eChunk.toString());
    }
  }

  if (sentCount === 0 && errorsList.length > 0) {
    return createJsonResponse({ success: false, message: "Envoi totalement échoué.\n" + errorsList.join("\n") });
  }

  // Mise à jour statut dès qu'au moins un mail est parti (ou si test réussi)
  if (!isTest) {
    try {
      const ws = getSheetSafe(ss, "CAMPAGNES_COM");
      const d = ws.getDataRange().getValues();
      for(let i = 1; i < d.length; i++) {
        if(String(d[i][0]) === cmpId) {
          ws.getRange(i+1, 7).setValue(new Date());
          ws.getRange(i+1, 8).setValue("Diffusé");
          break;
        }
      }
    } catch(eStatus) {}
  }

  const modeDiffusion = isTest ? "Test" : "Envoi réel";
  const logDetail = modeDiffusion + " — " + (p.nom || p.type || "Campagne") + " | Objet: « " + (p.objet || subject || "?") + " » | " + sentCount + "/" + emailsArray.length + " envoyé(s)" + (errorsList.length > 0 ? " | " + errorsList.length + " lot(s) en erreur" : "");
  logAction(p.idAuteur, p.roleAuteur, "Action", "Campagne", cmpId, logDetail);
  logMailDetail(ss, {
    auteur: p.idAuteur || "Système",
    categorie: "Com & Événement",
    sousCat: (isTest ? "[TEST] " : "") + (p.type || "Campagne"),
    objet: subject,
    destinataires: sentEmails,
    corps: htmlBody,
    sourceId: cmpId
  });

  const partialSuccess = errorsList.length > 0 && sentCount > 0;
  const msgBase = isTest ? "Mail de test envoyé !" : (partialSuccess ? "Envoi partiel : " + sentCount + "/" + emailsArray.length + " contacts atteints." : "Campagne diffusée avec succès à " + sentCount + " contacts.");
  return createJsonResponse({
    success: true,
    partialSuccess: partialSuccess,
    sent: sentCount,
    failed: emailsArray.length - sentCount,
    errors: errorsList,
    message: msgBase
  });
}

// --- TRAMES NETTOYÉES SANS EMOJI ---
function getTemplateCV() {
  return `<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Atelier CV HECG</title>
</head>
<body style="margin: 0; padding: 0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f7f9; color: #333;">

    <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width: 600px; background-color: #ffffff; margin-top: 20px; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 10px rgba(0,0,0,0.05);">
        
        <!-- HEADER AVEC LOGO (LIEN VERS SITE WEB) -->
        <tr>
            <td align="center" style="padding: 30px 20px;">
                <a href="https://hecg.fr/" target="_blank">
                    <img src="https://hecg.fr/wp-content/uploads/2025/03/LOGO-FORMAT-PNG-1-1.png" alt="Logo HECG" width="180" style="display: block; border: 0;">
                </a>
            </td>
        </tr>

        <!-- HERO SECTION AVEC DÉGRADÉ -->
        <tr>
            <td style="background: linear-gradient(135deg, #004a99 0%, #002a5c 100%); padding: 40px 30px; text-align: center;">
                <h1 style="color: #ffffff; margin: 0; font-size: 26px; font-weight: 700; letter-spacing: -0.5px;">Transformez votre CV en aimant à recruteurs 🚀</h1>
                <p style="color: #d1e3ff; font-size: 16px; margin-top: 15px;">Participez à nos ateliers exclusifs pour décrocher votre prochaine alternance.</p>
            </td>
        </tr>

        <!-- CORPS DU MAIL -->
        <tr>
            <td style="padding: 20px 30px 40px 30px;">
                
                <!-- LISTE À PUCES STYLISÉE -->
                <table width="100%" cellpadding="0" cellspacing="0" style="margin: 10px 0;">
                    <tr>
                        <td style="padding: 10px 0;">
                            <span style="background-color: #ff8c00; color: white; padding: 2px 8px; border-radius: 4px; font-weight: bold; margin-right: 10px;">✔</span> 
                            <span style="font-size: 15px;">Mise en page pro et moderne aux couleurs de l'établissement</span>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding: 10px 0;">
                            <span style="background-color: #ff8c00; color: white; padding: 2px 8px; border-radius: 4px; font-weight: bold; margin-right: 10px;">✔</span> 
                            <span style="font-size: 15px;">Mots-clés stratégiques pour les recruteurs</span>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding: 10px 0;">
                            <span style="background-color: #ff8c00; color: white; padding: 2px 8px; border-radius: 4px; font-weight: bold; margin-right: 10px;">✔</span> 
                            <span style="font-size: 15px;">Feedback personnalisé en direct</span>
                        </td>
                    </tr>
                </table>

                <!-- INFO LIEU DYNAMIQUE -->
                <div style="background-color: #f8f9fa; border-left: 4px solid #004a99; padding: 20px; margin-top: 30px;">
                    <p style="margin: 0; font-weight: bold; color: #004a99;">📍 Lieu de l'atelier :</p>
                    <p style="margin: 5px 0 0 0; font-size: 15px; color: #555;">
                        <!-- REMPLACER ICI PAR L'ADRESSE : 
                             Marseille : 39 Rue Sainte Cécile, 13005 Marseille 
                             Aix : Le Mansard B, 2 Pl. Romée de Villeneuve, 13090 Aix-en-Provence 
                        -->
                        <strong>{{INSERER_ICI_ADRESSE_CAMPUS}}</strong>
                    </p>
                </div>

                <!-- CALL TO ACTION -->
                <table width="100%" cellpadding="0" cellspacing="0" style="margin-top: 40px;">
                    <tr>
                        <td align="center">
                            <!-- === ATTENTION : INSERER VOTRE LIEN D'INSCRIPTION CI-DESSOUS === -->
                            <a href="{{VOTRE_LIEN_D_INSCRIPTION_ICI}}" style="background-color: #ff8c00; color: #ffffff; padding: 18px 35px; text-decoration: none; font-size: 18px; font-weight: bold; border-radius: 50px; display: inline-block;">
                                M'inscrire à l'atelier
                            </a>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>

        <!-- FOOTER & RÉSEAUX SOCIAUX -->
        <tr>
            <td style="background-color: #001a38; color: #ffffff; padding: 40px 30px; text-align: center;">
                <p style="font-size: 14px; margin-bottom: 20px; opacity: 0.8;">Suivez les actualités de l'HECG :</p>
                
                <table align="center" cellpadding="0" cellspacing="0">
                    <tr>
                        <!-- LINKEDIN -->
                        <td style="padding: 0 15px;">
                            <a href="https://linkedin.com/in/hecg-aix-marseille-a21232239" target="_blank">
                                <img src="https://cdn-icons-png.flaticon.com/512/174/174857.png" alt="LinkedIn" width="30" style="display: block; filter: brightness(0) invert(1);">
                            </a>
                        </td>
                        <!-- INSTAGRAM -->
                        <td style="padding: 0 15px;">
                            <a href="https://www.instagram.com/hecgecole" target="_blank">
                                <img src="https://cdn-icons-png.flaticon.com/512/174/174855.png" alt="Instagram" width="30" style="display: block; filter: brightness(0) invert(1);">
                            </a>
                        </td>
                        <!-- TIKTOK -->
                        <td style="padding: 0 15px;">
                            <a href="https://www.tiktok.com/@hecgmarseille" target="_blank">
                                <img src="https://cdn-icons-png.flaticon.com/512/3046/3046121.png" alt="TikTok" width="30" style="display: block; filter: brightness(0) invert(1);">
                            </a>
                        </td>
                    </tr>
                </table>

                <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid rgba(255,255,255,0.1); font-size: 12px; opacity: 0.6;">
                    <p>© 2026 HECG - Hautes Etudes de Comptabilité et de Gestion<br>
                    <a href="https://hecg.fr/" style="color: #ffffff; text-decoration: underline;">www.hecg.fr</a></p>
                </div>
            </td>
        </tr>
    </table>

</body>
</html>`;
}
function getTemplateLinkedIn() {
  return `<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Atelier LinkedIn HECG</title>
</head>
<body style="margin: 0; padding: 0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f7f9; color: #333;">

    <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width: 600px; background-color: #ffffff; margin-top: 20px; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 10px rgba(0,0,0,0.05);">
        
        <!-- HEADER AVEC LOGO (LIEN VERS SITE WEB) -->
        <tr>
            <td align="center" style="padding: 30px 20px;">
                <a href="https://hecg.fr/" target="_blank">
                    <img src="https://hecg.fr/wp-content/uploads/2025/03/LOGO-FORMAT-PNG-1-1.png" alt="Logo HECG" width="180" style="display: block; border: 0;">
                </a>
            </td>
        </tr>

        <!-- HERO SECTION AVEC DÉGRADÉ -->
        <tr>
            <td style="background: linear-gradient(135deg, #004a99 0%, #002a5c 100%); padding: 40px 30px; text-align: center;">
                <h1 style="color: #ffffff; margin: 0; font-size: 26px; font-weight: 700; letter-spacing: -0.5px;">Devenez incontournable sur LinkedIn 🚀</h1>
                <p style="color: #d1e3ff; font-size: 16px; margin-top: 15px;">Participez à nos ateliers exclusifs pour décrocher votre prochaine alternance.</p>
            </td>
        </tr>

        <!-- CORPS DU MAIL -->
        <tr>
            <td style="padding: 20px 30px 40px 30px;">
                
                <!-- LISTE À PUCES STYLISÉE -->
                <table width="100%" cellpadding="0" cellspacing="0" style="margin: 10px 0;">
                    <tr>
                        <td style="padding: 10px 0;">
                            <span style="background-color: #ff8c00; color: white; padding: 2px 8px; border-radius: 4px; font-weight: bold; margin-right: 10px;">✔</span> 
                            <span style="font-size: 15px;">Profil pro et moderne aux couleurs de l'établissement</span>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding: 10px 0;">
                            <span style="background-color: #ff8c00; color: white; padding: 2px 8px; border-radius: 4px; font-weight: bold; margin-right: 10px;">✔</span> 
                            <span style="font-size: 15px;">Maîtrise des outils de réseautage stratégique</span>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding: 10px 0;">
                            <span style="background-color: #ff8c00; color: white; padding: 2px 8px; border-radius: 4px; font-weight: bold; margin-right: 10px;">✔</span> 
                            <span style="font-size: 15px;">Feedback personnalisé sur votre profil en direct</span>
                        </td>
                    </tr>
                </table>

                <!-- INFO LIEU DYNAMIQUE -->
                <div style="background-color: #f8f9fa; border-left: 4px solid #004a99; padding: 20px; margin-top: 30px;">
                    <p style="margin: 0; font-weight: bold; color: #004a99;">📍 Lieu de l'atelier :</p>
                    <p style="margin: 5px 0 0 0; font-size: 15px; color: #555;">
                        <!-- REMPLACER ICI PAR L'ADRESSE : 
                             Marseille : 39 Rue Sainte Cécile, 13005 Marseille 
                             Aix : Le Mansard B, 2 Pl. Romée de Villeneuve, 13090 Aix-en-Provence 
                        -->
                        <strong>{{INSERER_ICI_ADRESSE_CAMPUS}}</strong>
                    </p>
                </div>

                <!-- CALL TO ACTION -->
                <table width="100%" cellpadding="0" cellspacing="0" style="margin-top: 40px;">
                    <tr>
                        <td align="center">
                            <!-- === ATTENTION : INSERER VOTRE LIEN D'INSCRIPTION CI-DESSOUS === -->
                            <a href="{{VOTRE_LIEN_D_INSCRIPTION_ICI}}" style="background-color: #ff8c00; color: #ffffff; padding: 18px 35px; text-decoration: none; font-size: 18px; font-weight: bold; border-radius: 50px; display: inline-block;">
                                M'inscrire à l'atelier
                            </a>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>

        <!-- FOOTER & RÉSEAUX SOCIAUX -->
        <tr>
            <td style="background-color: #001a38; color: #ffffff; padding: 40px 30px; text-align: center;">
                <p style="font-size: 14px; margin-bottom: 20px; opacity: 0.8;">Suivez les actualités de l'HECG :</p>
                
                <table align="center" cellpadding="0" cellspacing="0">
                    <tr>
                        <!-- LINKEDIN -->
                        <td style="padding: 0 15px;">
                            <a href="https://linkedin.com/in/hecg-aix-marseille-a21232239" target="_blank">
                                <img src="https://cdn-icons-png.flaticon.com/512/174/174857.png" alt="LinkedIn" width="30" style="display: block; filter: brightness(0) invert(1);">
                            </a>
                        </td>
                        <!-- INSTAGRAM -->
                        <td style="padding: 0 15px;">
                            <a href="https://www.instagram.com/hecgecole" target="_blank">
                                <img src="https://cdn-icons-png.flaticon.com/512/174/174855.png" alt="Instagram" width="30" style="display: block; filter: brightness(0) invert(1);">
                            </a>
                        </td>
                        <!-- TIKTOK -->
                        <td style="padding: 0 15px;">
                            <a href="https://www.tiktok.com/@hecgmarseille" target="_blank">
                                <img src="https://cdn-icons-png.flaticon.com/512/3046/3046121.png" alt="TikTok" width="30" style="display: block; filter: brightness(0) invert(1);">
                            </a>
                        </td>
                    </tr>
                </table>

                <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid rgba(255,255,255,0.1); font-size: 12px; opacity: 0.6;">
                    <p>© 2026 HECG - Hautes Etudes de Comptabilité et de Gestion<br>
                    <a href="https://hecg.fr/" style="color: #ffffff; text-decoration: underline;">www.hecg.fr</a></p>
                </div>
            </td>
        </tr>
    </table>

</body>
</html>`;
}
function getTemplateMixte() {
  return `<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Atelier CV & LinkedIn HECG</title>
</head>
<body style="margin: 0; padding: 0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f7f9; color: #333;">

    <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width: 600px; background-color: #ffffff; margin-top: 20px; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 10px rgba(0,0,0,0.05);">
        
        <!-- HEADER AVEC LOGO (LIEN VERS SITE WEB) -->
        <tr>
            <td align="center" style="padding: 30px 20px;">
                <a href="https://hecg.fr/" target="_blank">
                    <img src="https://hecg.fr/wp-content/uploads/2025/03/LOGO-FORMAT-PNG-1-1.png" alt="Logo HECG" width="180" style="display: block; border: 0;">
                </a>
            </td>
        </tr>

        <!-- HERO SECTION AVEC DÉGRADÉ -->
        <tr>
            <td style="background: linear-gradient(135deg, #004a99 0%, #002a5c 100%); padding: 40px 30px; text-align: center;">
                <h1 style="color: #ffffff; margin: 0; font-size: 26px; font-weight: 700; letter-spacing: -0.5px;">Boostez votre visibilité : CV & LinkedIn 🚀</h1>
                <p style="color: #d1e3ff; font-size: 16px; margin-top: 15px;">Un atelier combiné exclusif pour décrocher votre prochaine alternance.</p>
            </td>
        </tr>

        <!-- CORPS DU MAIL -->
        <tr>
            <td style="padding: 20px 30px 40px 30px;">
                
                <!-- LISTE À PUCES STYLISÉE -->
                <table width="100%" cellpadding="0" cellspacing="0" style="margin: 10px 0;">
                    <tr>
                        <td style="padding: 10px 0;">
                            <span style="background-color: #ff8c00; color: white; padding: 2px 8px; border-radius: 4px; font-weight: bold; margin-right: 10px;">✔</span> 
                            <span style="font-size: 15px;">CV pro aux couleurs de l'établissement</span>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding: 10px 0;">
                            <span style="background-color: #ff8c00; color: white; padding: 2px 8px; border-radius: 4px; font-weight: bold; margin-right: 10px;">✔</span> 
                            <span style="font-size: 15px;">Optimisation complète du profil LinkedIn</span>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding: 10px 0;">
                            <span style="background-color: #ff8c00; color: white; padding: 2px 8px; border-radius: 4px; font-weight: bold; margin-right: 10px;">✔</span> 
                            <span style="font-size: 15px;">Stratégies de réseautage et feedback en direct</span>
                        </td>
                    </tr>
                </table>

                <!-- INFO LIEU DYNAMIQUE -->
                <div style="background-color: #f8f9fa; border-left: 4px solid #004a99; padding: 20px; margin-top: 30px;">
                    <p style="margin: 0; font-weight: bold; color: #004a99;">📍 Lieu de l'atelier :</p>
                    <p style="margin: 5px 0 0 0; font-size: 15px; color: #555;">
                        <!-- REMPLACER ICI PAR L'ADRESSE : 
                             Marseille : 39 Rue Sainte Cécile, 13005 Marseille 
                             Aix : Le Mansard B, 2 Pl. Romée de Villeneuve, 13090 Aix-en-Provence 
                        -->
                        <strong>{{INSERER_ICI_ADRESSE_CAMPUS}}</strong>
                    </p>
                </div>

                <!-- CALL TO ACTION -->
                <table width="100%" cellpadding="0" cellspacing="0" style="margin-top: 40px;">
                    <tr>
                        <td align="center">
                            <!-- === ATTENTION : INSERER VOTRE LIEN D'INSCRIPTION CI-DESSOUS === -->
                            <a href="{{VOTRE_LIEN_D_INSCRIPTION_ICI}}" style="background-color: #ff8c00; color: #ffffff; padding: 18px 35px; text-decoration: none; font-size: 18px; font-weight: bold; border-radius: 50px; display: inline-block;">
                                M'inscrire à l'atelier combiné
                            </a>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>

        <!-- FOOTER & RÉSEAUX SOCIAUX -->
        <tr>
            <td style="background-color: #001a38; color: #ffffff; padding: 40px 30px; text-align: center;">
                <p style="font-size: 14px; margin-bottom: 20px; opacity: 0.8;">Suivez les actualités de l'HECG :</p>
                
                <table align="center" cellpadding="0" cellspacing="0">
                    <tr>
                        <!-- LINKEDIN -->
                        <td style="padding: 0 15px;">
                            <a href="https://linkedin.com/in/hecg-aix-marseille-a21232239" target="_blank">
                                <img src="https://cdn-icons-png.flaticon.com/512/174/174857.png" alt="LinkedIn" width="30" style="display: block; filter: brightness(0) invert(1);">
                            </a>
                        </td>
                        <!-- INSTAGRAM -->
                        <td style="padding: 0 15px;">
                            <a href="https://www.instagram.com/hecgecole" target="_blank">
                                <img src="https://cdn-icons-png.flaticon.com/512/174/174855.png" alt="Instagram" width="30" style="display: block; filter: brightness(0) invert(1);">
                            </a>
                        </td>
                        <!-- TIKTOK -->
                        <td style="padding: 0 15px;">
                            <a href="https://www.tiktok.com/@hecgmarseille" target="_blank">
                                <img src="https://cdn-icons-png.flaticon.com/512/3046/3046121.png" alt="TikTok" width="30" style="display: block; filter: brightness(0) invert(1);">
                            </a>
                        </td>
                    </tr>
                </table>

                <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid rgba(255,255,255,0.1); font-size: 12px; opacity: 0.6;">
                    <p>© 2026 HECG - Hautes Etudes de Comptabilité et de Gestion<br>
                    <a href="https://hecg.fr/" style="color: #ffffff; text-decoration: underline;">www.hecg.fr</a></p>
                </div>
            </td>
        </tr>
    </table>

</body>
</html>`;
}
function getTemplateJPO() {
  return `<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>JPO HECG - Journée Portes Ouvertes</title>
</head>
<body style="margin: 0; padding: 0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f7f9; color: #333;">

    <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width: 600px; background-color: #ffffff; margin-top: 20px; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 10px rgba(0,0,0,0.05);">
        
        <!-- HEADER AVEC LOGO (LIEN VERS SITE WEB) -->
        <tr>
            <td align="center" style="padding: 30px 20px;">
                <a href="https://hecg.fr/" target="_blank">
                    <img src="https://hecg.fr/wp-content/uploads/2025/03/LOGO-FORMAT-PNG-1-1.png" alt="Logo HECG" width="180" style="display: block; border: 0;">
                </a>
            </td>
        </tr>

        <!-- HERO SECTION AVEC DÉGRADÉ -->
        <tr>
            <td style="background: linear-gradient(135deg, #004a99 0%, #002a5c 100%); padding: 45px 30px; text-align: center;">
                <h1 style="color: #ffffff; margin: 0; font-size: 28px; font-weight: 700; letter-spacing: -0.5px;">HECG a le plaisir de vous accueillir pour ses JPO</h1>
                <p style="color: #d1e3ff; font-size: 16px; margin-top: 15px;">Venez découvrir nos programmes et rencontrer nos équipes pédagogiques.</p>
            </td>
        </tr>

        <!-- CORPS DU MAIL -->
        <tr>
            <td style="padding: 20px 30px 40px 30px;">
                
                <p style="font-size: 16px; line-height: 1.6; color: #444; text-align: center;">
                    Vous hésitez encore sur votre orientation ? Profitez de cette journée pour poser toutes vos questions et visiter nos infrastructures.
                </p>

                <!-- LISTE DES AVANTAGES -->
                <table width="100%" cellpadding="0" cellspacing="0" style="margin: 10px 0;">
                    <tr>
                        <td style="padding: 10px 0;">
                            <span style="background-color: #ff8c00; color: white; padding: 2px 8px; border-radius: 4px; font-weight: bold; margin-right: 10px;">✔</span> 
                            <span style="font-size: 15px;">Présentation des cursus (Comptabilité, Gestion, Paye...)</span>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding: 10px 0;">
                            <span style="background-color: #ff8c00; color: white; padding: 2px 8px; border-radius: 4px; font-weight: bold; margin-right: 10px;">✔</span> 
                            <span style="font-size: 15px;">Échanges avec nos étudiants et diplômés</span>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding: 10px 0;">
                            <span style="background-color: #ff8c00; color: white; padding: 2px 8px; border-radius: 4px; font-weight: bold; margin-right: 10px;">✔</span> 
                            <span style="font-size: 15px;">Aide à la recherche d'alternance</span>
                        </td>
                    </tr>
                </table>

                <!-- INFO PRATIQUES (DATE, HEURE, LIEU) -->
                <div style="background-color: #f8f9fa; border-left: 4px solid #004a99; padding: 25px; margin-top: 30px; border-radius: 0 4px 4px 0;">
                    <p style="margin: 0; font-weight: bold; color: #004a99; font-size: 18px;">📅 Informations pratiques :</p>
                    <p style="margin: 10px 0 5px 0; font-size: 15px; color: #333;">
                        <strong>Date :</strong> {{INSERER_DATE_ICI}}
                    </p>
                    <p style="margin: 5px 0 5px 0; font-size: 15px; color: #333;">
                        <strong>Horaires :</strong> {{INSERER_HORAIRES_ICI}}
                    </p>
                    <p style="margin: 5px 0 0 0; font-size: 15px; color: #333;">
                        <strong>Campus :</strong> {{INSERER_CAMPUS_ET_ADRESSE_ICI}}
                    </p>
                </div>

                <!-- CALL TO ACTION -->
                <table width="100%" cellpadding="0" cellspacing="0" style="margin-top: 40px;">
                    <tr>
                        <td align="center">
                            <!-- === ATTENTION : INSERER LE LIEN D'INSCRIPTION JPO CI-DESSOUS === -->
                            <a href="{{LIEN_INSCRIPTION_JPO_ICI}}" style="background-color: #ff8c00; color: #ffffff; padding: 18px 35px; text-decoration: none; font-size: 18px; font-weight: bold; border-radius: 50px; display: inline-block; box-shadow: 0 4px 6px rgba(255,140,0,0.3);">
                                Je m'inscris à la JPO
                            </a>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>

        <!-- FOOTER & RÉSEAUX SOCIAUX -->
        <tr>
            <td style="background-color: #001a38; color: #ffffff; padding: 40px 30px; text-align: center;">
                <p style="font-size: 14px; margin-bottom: 20px; opacity: 0.8;">Restez connecté avec l'HECG :</p>
                
                <table align="center" cellpadding="0" cellspacing="0">
                    <tr>
                        <!-- LINKEDIN -->
                        <td style="padding: 0 15px;">
                            <a href="https://linkedin.com/in/hecg-aix-marseille-a21232239" target="_blank">
                                <img src="https://cdn-icons-png.flaticon.com/512/174/174857.png" alt="LinkedIn" width="30" style="display: block; filter: brightness(0) invert(1);">
                            </a>
                        </td>
                        <!-- INSTAGRAM -->
                        <td style="padding: 0 15px;">
                            <a href="https://www.instagram.com/hecgecole" target="_blank">
                                <img src="https://cdn-icons-png.flaticon.com/512/174/174855.png" alt="Instagram" width="30" style="display: block; filter: brightness(0) invert(1);">
                            </a>
                        </td>
                        <!-- TIKTOK -->
                        <td style="padding: 0 15px;">
                            <a href="https://www.tiktok.com/@hecgmarseille" target="_blank">
                                <img src="https://cdn-icons-png.flaticon.com/512/3046/3046121.png" alt="TikTok" width="30" style="display: block; filter: brightness(0) invert(1);">
                            </a>
                        </td>
                    </tr>
                </table>

                <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid rgba(255,255,255,0.1); font-size: 12px; opacity: 0.6;">
                    <p>© 2026 HECG - Hautes Etudes de Comptabilité et de Gestion<br>
                    <a href="https://hecg.fr/" style="color: #ffffff; text-decoration: underline;">www.hecg.fr</a></p>
                </div>
            </td>
        </tr>
    </table>

</body>
</html>`;
}
function getTemplateSalon() {
  return `<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>HECG au Salon</title>
</head>
<body style="margin: 0; padding: 0; background-color: #001a38; font-family: 'Segoe UI', Arial, sans-serif;">

    <!-- CONTENEUR PRINCIPAL AVEC DÉGRADÉ DE BLEU -->
    <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="background: linear-gradient(180deg, #004a99 0%, #001a38 100%); min-height: 100vh;">
        <tr>
            <td align="center" style="padding: 40px 10px;">
                
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width: 600px;">
                    
                    <!-- LOGO -->
                    <tr>
                        <td align="center" style="padding-bottom: 50px;">
                            <a href="https://hecg.fr/" target="_blank">
                                <img src="https://hecg.fr/wp-content/uploads/2025/03/LOGO-FORMAT-PNG-1-1.png" alt="Logo HECG" width="160" style="display: block; border: 0; filter: brightness(0) invert(1);">
                            </a>
                        </td>
                    </tr>

                    <!-- CONTENU VISUEL / PRÉSENTATION -->
                    <tr>
                        <td align="center" style="padding: 0 20px;">
                            <h1 style="color: #ffffff; font-size: 32px; font-weight: 800; line-height: 1.2; margin: 0; letter-spacing: -1px;">
                                Venez nous retrouver.
                            </h1>
                            
                            <div style="width: 50px; height: 4px; background-color: #FF7A00; margin: 25px 0;"></div>
                            
                            <p style="color: #d1e3ff; font-size: 18px; line-height: 1.6; margin: 0; font-weight: 400;">
                                {{TEXTE_ACCROCHE}}
                            </p>
                        </td>
                    </tr>

                    <!-- CARD INFO STYLISÉE -->
                    <tr>
                        <td style="padding: 50px 20px 0 20px;">
                            <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: rgba(255, 255, 255, 0.05); border: 1px solid rgba(255, 255, 255, 0.1); border-radius: 20px; backdrop-filter: blur(10px);">
                                <tr>
                                    <td style="padding: 35px; text-align: center;">
                                        <p style="color: #FF7A00; font-weight: bold; text-transform: uppercase; letter-spacing: 2px; margin: 0 0 15px 0; font-size: 13px;">
                                            Prochain Événement
                                        </p>
                                        <h2 style="color: #ffffff; font-size: 24px; margin: 0; font-weight: 600;">
                                            {{NOM_DU_SALON_OU_EVENEMENT}}
                                        </h2>
                                        
                                        <table border="0" cellpadding="0" cellspacing="0" width="100%" style="margin-top: 25px;">
                                            <tr>
                                                <td align="center" style="color: #ffffff; font-size: 16px; opacity: 0.9;">
                                                    <span style="display: inline-block; padding: 5px 10px; border-radius: 5px; background: rgba(255,255,255,0.1); margin: 5px;"> {{DATE_ET_HORAIRES}}</span>
                                                    <span style="display: inline-block; padding: 5px 10px; border-radius: 5px; background: rgba(255,255,255,0.1); margin: 5px;"> {{LIEU_ET_ADRESSE}}</span>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>

                    <!-- BOUTON ACTION -->
                    <tr>
                        <td align="center" style="padding: 50px 20px;">
                            <a href="{{LIEN_INSCRIPTION_ICI}}" style="background-color: #FF7A00; color: #ffffff; padding: 20px 45px; text-decoration: none; font-size: 18px; font-weight: bold; border-radius: 12px; display: inline-block; box-shadow: 0 10px 20px rgba(255, 122, 0, 0.3);">
                                Lien de l'organisateur
                            </a>
                        </td>
                    </tr>

                    <!-- FOOTER & RÉSEAUX -->
                    <tr>
                        <td align="center" style="padding: 40px 20px; border-top: 1px solid rgba(255, 255, 255, 0.1);">
                            <p style="color: #ffffff; font-size: 14px; opacity: 0.6; margin-bottom: 25px;">Retrouvez-nous sur nos réseaux sociaux</p>
                            
                            <table border="0" cellpadding="0" cellspacing="0">
                                <tr>
                                    <td style="padding: 0 15px;">
                                        <a href="https://linkedin.com/in/hecg-aix-marseille-a21232239" target="_blank">
                                            <img src="https://cdn-icons-png.flaticon.com/512/174/174857.png" alt="LinkedIn" width="28" style="filter: brightness(0) invert(1);">
                                        </a>
                                    </td>
                                    <td style="padding: 0 15px;">
                                        <a href="https://www.instagram.com/hecgecole" target="_blank">
                                            <img src="https://cdn-icons-png.flaticon.com/512/174/174855.png" alt="Instagram" width="28" style="filter: brightness(0) invert(1);">
                                        </a>
                                    </td>
                                    <td style="padding: 0 15px;">
                                        <a href="https://www.tiktok.com/@hecgmarseille" target="_blank">
                                            <img src="https://cdn-icons-png.flaticon.com/512/3046/3046121.png" alt="TikTok" width="28" style="filter: brightness(0) invert(1);">
                                        </a>
                                    </td>
                                </tr>
                            </table>

                            <div style="margin-top: 40px; color: #ffffff; font-size: 11px; opacity: 0.4; line-height: 1.5;">
                                © 2026 HECG - Hautes Etudes de Comptabilité et de Gestion<br>
                                <a href="https://hecg.fr/" style="color: #ffffff;">www.hecg.fr</a>
                            </div>
                        </td>
                    </tr>

                </table>
                
            </td>
        </tr>
    </table>

</body>
</html>`;
}

// ==========================================
// PROSPECTS & INSCRIPTIONS
// ==========================================
function handlePublicRegister(p, ss) {
  // 1. On cherche l'événement dans le NOUVEL onglet
  const sheetCom = ss.getSheetByName("CAMPAGNES_COM");
  const dataCom = sheetCom.getDataRange().getValues();
  let eventName = "Événement inconnu";
  
  for(let i=1; i<dataCom.length; i++) {
    if(String(dataCom[i][0]) === String(p.eventId)) {
      eventName = dataCom[i][1]; // On récupère le Nom de l'événement (colonne B)
      break;
    }
  }

  // 2. On ajoute le prospect dans l'onglet PROSPECTS
  const sheetPros = getSheetSafe(ss, "PROSPECTS");
  const newId = "PR" + new Date().getTime();
  
  sheetPros.appendRow([
    newId,
    p.nom,
    p.prenom,
    p.email,
    p.tel,
    eventName,             // Source : Le nom de l'atelier/salon
    new Date(),            // Date d'inscription
    p.eventId              // ID de l'événement pour le suivi
  ]);

  return createJsonResponse({ success: true, message: "Inscription validée" });
}

function handleGetProspects(params, ss) {
  const data = getSheetSafe(ss, "PROSPECTS").getDataRange().getValues();
  const list = [];
  for(let i=1; i<data.length; i++) if(data[i][0]) list.push({ id: data[i][0], nom: data[i][1], prenom: data[i][2], mail: data[i][3], tel: data[i][4], source: data[i][5], comm: data[i][6] });
  return createJsonResponse({ success: true, data: list.reverse() });
}

function handleDeleteProspect(params, ss) { 
  const ws = getSheetSafe(ss, "PROSPECTS"); const d = ws.getDataRange().getValues();
  for(let i=1; i<d.length; i++) if(String(d[i][0]) === params.id) { ws.deleteRow(i+1); return createJsonResponse({success: true}); }
  return createJsonResponse({success: false}); 
}

// ==========================================
// TRAMES HTML (NETTOYEES SANS EMOJI)
// ==========================================
function getTemplateCV() {
  return `<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Atelier CV HECG</title>
</head>
<body style="margin: 0; padding: 0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f7f9; color: #333;">

    <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width: 600px; background-color: #ffffff; margin-top: 20px; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 10px rgba(0,0,0,0.05);">
        
        <!-- HEADER AVEC LOGO (LIEN VERS SITE WEB) -->
        <tr>
            <td align="center" style="padding: 30px 20px;">
                <a href="https://hecg.fr/" target="_blank">
                    <img src="https://hecg.fr/wp-content/uploads/2025/03/LOGO-FORMAT-PNG-1-1.png" alt="Logo HECG" width="180" style="display: block; border: 0;">
                </a>
            </td>
        </tr>

        <!-- HERO SECTION AVEC DÉGRADÉ -->
        <tr>
            <td style="background: linear-gradient(135deg, #004a99 0%, #002a5c 100%); padding: 40px 30px; text-align: center;">
                <h1 style="color: #ffffff; margin: 0; font-size: 26px; font-weight: 700; letter-spacing: -0.5px;">Transformez votre CV en aimant à recruteurs </h1>
                <p style="color: #d1e3ff; font-size: 16px; margin-top: 15px;">Participez à nos ateliers exclusifs pour décrocher votre prochaine alternance.</p>
            </td>
        </tr>

        <!-- CORPS DU MAIL -->
        <tr>
            <td style="padding: 20px 30px 40px 30px;">
                
                <!-- LISTE À PUCES STYLISÉE -->
                <table width="100%" cellpadding="0" cellspacing="0" style="margin: 10px 0;">
                    <tr>
                        <td style="padding: 10px 0;">
                            <span style="background-color: #ff8c00; color: white; padding: 2px 8px; border-radius: 4px; font-weight: bold; margin-right: 10px;">✔</span> 
                            <span style="font-size: 15px;">Mise en page pro et moderne aux couleurs de l'établissement</span>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding: 10px 0;">
                            <span style="background-color: #ff8c00; color: white; padding: 2px 8px; border-radius: 4px; font-weight: bold; margin-right: 10px;">✔</span> 
                            <span style="font-size: 15px;">Mots-clés stratégiques pour les recruteurs</span>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding: 10px 0;">
                            <span style="background-color: #ff8c00; color: white; padding: 2px 8px; border-radius: 4px; font-weight: bold; margin-right: 10px;">✔</span> 
                            <span style="font-size: 15px;">Feedback personnalisé en direct</span>
                        </td>
                    </tr>
                </table>

                <!-- INFO LIEU DYNAMIQUE -->
                <div style="background-color: #f8f9fa; border-left: 4px solid #004a99; padding: 20px; margin-top: 30px;">
                    <p style="margin: 0; font-weight: bold; color: #004a99;"> Lieu de l'atelier :</p>
                    <p style="margin: 5px 0 0 0; font-size: 15px; color: #555;">
                        <!-- REMPLACER ICI PAR L'ADRESSE : 
                             Marseille : 39 Rue Sainte Cécile, 13005 Marseille 
                             Aix : Le Mansard B, 2 Pl. Romée de Villeneuve, 13090 Aix-en-Provence 
                        -->
                        <strong>{{INSERER_ICI_ADRESSE_CAMPUS}}</strong>
                    </p>
                </div>

                <!-- CALL TO ACTION -->
                <table width="100%" cellpadding="0" cellspacing="0" style="margin-top: 40px;">
                    <tr>
                        <td align="center">
                            <!-- === ATTENTION : INSERER VOTRE LIEN D'INSCRIPTION CI-DESSOUS === -->
                            <a href="{{VOTRE_LIEN_D_INSCRIPTION_ICI}}" style="background-color: #ff8c00; color: #ffffff; padding: 18px 35px; text-decoration: none; font-size: 18px; font-weight: bold; border-radius: 50px; display: inline-block;">
                                M'inscrire à l'atelier
                            </a>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>

        <!-- FOOTER & RÉSEAUX SOCIAUX -->
        <tr>
            <td style="background-color: #001a38; color: #ffffff; padding: 40px 30px; text-align: center;">
                <p style="font-size: 14px; margin-bottom: 20px; opacity: 0.8;">Suivez les actualités de l'HECG :</p>
                
                <table align="center" cellpadding="0" cellspacing="0">
                    <tr>
                        <!-- LINKEDIN -->
                        <td style="padding: 0 15px;">
                            <a href="https://linkedin.com/in/hecg-aix-marseille-a21232239" target="_blank">
                                <img src="https://cdn-icons-png.flaticon.com/512/174/174857.png" alt="LinkedIn" width="30" style="display: block; filter: brightness(0) invert(1);">
                            </a>
                        </td>
                        <!-- INSTAGRAM -->
                        <td style="padding: 0 15px;">
                            <a href="https://www.instagram.com/hecgecole" target="_blank">
                                <img src="https://cdn-icons-png.flaticon.com/512/174/174855.png" alt="Instagram" width="30" style="display: block; filter: brightness(0) invert(1);">
                            </a>
                        </td>
                        <!-- TIKTOK -->
                        <td style="padding: 0 15px;">
                            <a href="https://www.tiktok.com/@hecgmarseille" target="_blank">
                                <img src="https://cdn-icons-png.flaticon.com/512/3046/3046121.png" alt="TikTok" width="30" style="display: block; filter: brightness(0) invert(1);">
                            </a>
                        </td>
                    </tr>
                </table>

                <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid rgba(255,255,255,0.1); font-size: 12px; opacity: 0.6;">
                    <p>© 2026 HECG - Hautes Etudes de Comptabilité et de Gestion<br>
                    <a href="https://hecg.fr/" style="color: #ffffff; text-decoration: underline;">www.hecg.fr</a></p>
                </div>
            </td>
        </tr>
    </table>

</body>
</html>`;
}
function getTemplateLinkedIn() {
  return `<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Atelier LinkedIn HECG</title>
</head>
<body style="margin: 0; padding: 0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f7f9; color: #333;">

    <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width: 600px; background-color: #ffffff; margin-top: 20px; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 10px rgba(0,0,0,0.05);">
        
        <!-- HEADER AVEC LOGO (LIEN VERS SITE WEB) -->
        <tr>
            <td align="center" style="padding: 30px 20px;">
                <a href="https://hecg.fr/" target="_blank">
                    <img src="https://hecg.fr/wp-content/uploads/2025/03/LOGO-FORMAT-PNG-1-1.png" alt="Logo HECG" width="180" style="display: block; border: 0;">
                </a>
            </td>
        </tr>

        <!-- HERO SECTION AVEC DÉGRADÉ -->
        <tr>
            <td style="background: linear-gradient(135deg, #004a99 0%, #002a5c 100%); padding: 40px 30px; text-align: center;">
                <h1 style="color: #ffffff; margin: 0; font-size: 26px; font-weight: 700; letter-spacing: -0.5px;">Devenez incontournable sur LinkedIn </h1>
                <p style="color: #d1e3ff; font-size: 16px; margin-top: 15px;">Participez à nos ateliers exclusifs pour décrocher votre prochaine alternance.</p>
            </td>
        </tr>

        <!-- CORPS DU MAIL -->
        <tr>
            <td style="padding: 20px 30px 40px 30px;">
                
                <!-- LISTE À PUCES STYLISÉE -->
                <table width="100%" cellpadding="0" cellspacing="0" style="margin: 10px 0;">
                    <tr>
                        <td style="padding: 10px 0;">
                            <span style="background-color: #ff8c00; color: white; padding: 2px 8px; border-radius: 4px; font-weight: bold; margin-right: 10px;">✔</span> 
                            <span style="font-size: 15px;">Profil pro et moderne aux couleurs de l'établissement</span>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding: 10px 0;">
                            <span style="background-color: #ff8c00; color: white; padding: 2px 8px; border-radius: 4px; font-weight: bold; margin-right: 10px;">✔</span> 
                            <span style="font-size: 15px;">Maîtrise des outils de réseautage stratégique</span>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding: 10px 0;">
                            <span style="background-color: #ff8c00; color: white; padding: 2px 8px; border-radius: 4px; font-weight: bold; margin-right: 10px;">✔</span> 
                            <span style="font-size: 15px;">Feedback personnalisé sur votre profil en direct</span>
                        </td>
                    </tr>
                </table>

                <!-- INFO LIEU DYNAMIQUE -->
                <div style="background-color: #f8f9fa; border-left: 4px solid #004a99; padding: 20px; margin-top: 30px;">
                    <p style="margin: 0; font-weight: bold; color: #004a99;"> Lieu de l'atelier :</p>
                    <p style="margin: 5px 0 0 0; font-size: 15px; color: #555;">
                        <!-- REMPLACER ICI PAR L'ADRESSE : 
                             Marseille : 39 Rue Sainte Cécile, 13005 Marseille 
                             Aix : Le Mansard B, 2 Pl. Romée de Villeneuve, 13090 Aix-en-Provence 
                        -->
                        <strong>{{INSERER_ICI_ADRESSE_CAMPUS}}</strong>
                    </p>
                </div>

                <!-- CALL TO ACTION -->
                <table width="100%" cellpadding="0" cellspacing="0" style="margin-top: 40px;">
                    <tr>
                        <td align="center">
                            <!-- === ATTENTION : INSERER VOTRE LIEN D'INSCRIPTION CI-DESSOUS === -->
                            <a href="{{VOTRE_LIEN_D_INSCRIPTION_ICI}}" style="background-color: #ff8c00; color: #ffffff; padding: 18px 35px; text-decoration: none; font-size: 18px; font-weight: bold; border-radius: 50px; display: inline-block;">
                                M'inscrire à l'atelier
                            </a>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>

        <!-- FOOTER & RÉSEAUX SOCIAUX -->
        <tr>
            <td style="background-color: #001a38; color: #ffffff; padding: 40px 30px; text-align: center;">
                <p style="font-size: 14px; margin-bottom: 20px; opacity: 0.8;">Suivez les actualités de l'HECG :</p>
                
                <table align="center" cellpadding="0" cellspacing="0">
                    <tr>
                        <!-- LINKEDIN -->
                        <td style="padding: 0 15px;">
                            <a href="https://linkedin.com/in/hecg-aix-marseille-a21232239" target="_blank">
                                <img src="https://cdn-icons-png.flaticon.com/512/174/174857.png" alt="LinkedIn" width="30" style="display: block; filter: brightness(0) invert(1);">
                            </a>
                        </td>
                        <!-- INSTAGRAM -->
                        <td style="padding: 0 15px;">
                            <a href="https://www.instagram.com/hecgecole" target="_blank">
                                <img src="https://cdn-icons-png.flaticon.com/512/174/174855.png" alt="Instagram" width="30" style="display: block; filter: brightness(0) invert(1);">
                            </a>
                        </td>
                        <!-- TIKTOK -->
                        <td style="padding: 0 15px;">
                            <a href="https://www.tiktok.com/@hecgmarseille" target="_blank">
                                <img src="https://cdn-icons-png.flaticon.com/512/3046/3046121.png" alt="TikTok" width="30" style="display: block; filter: brightness(0) invert(1);">
                            </a>
                        </td>
                    </tr>
                </table>

                <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid rgba(255,255,255,0.1); font-size: 12px; opacity: 0.6;">
                    <p>© 2026 HECG - Hautes Etudes de Comptabilité et de Gestion<br>
                    <a href="https://hecg.fr/" style="color: #ffffff; text-decoration: underline;">www.hecg.fr</a></p>
                </div>
            </td>
        </tr>
    </table>

</body>
</html>`;
}
function getTemplateMixte() {
  return `<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Atelier CV & LinkedIn HECG</title>
</head>
<body style="margin: 0; padding: 0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f7f9; color: #333;">

    <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width: 600px; background-color: #ffffff; margin-top: 20px; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 10px rgba(0,0,0,0.05);">
        
        <!-- HEADER AVEC LOGO (LIEN VERS SITE WEB) -->
        <tr>
            <td align="center" style="padding: 30px 20px;">
                <a href="https://hecg.fr/" target="_blank">
                    <img src="https://hecg.fr/wp-content/uploads/2025/03/LOGO-FORMAT-PNG-1-1.png" alt="Logo HECG" width="180" style="display: block; border: 0;">
                </a>
            </td>
        </tr>

        <!-- HERO SECTION AVEC DÉGRADÉ -->
        <tr>
            <td style="background: linear-gradient(135deg, #004a99 0%, #002a5c 100%); padding: 40px 30px; text-align: center;">
                <h1 style="color: #ffffff; margin: 0; font-size: 26px; font-weight: 700; letter-spacing: -0.5px;">Boostez votre visibilité : CV & LinkedIn </h1>
                <p style="color: #d1e3ff; font-size: 16px; margin-top: 15px;">Un atelier combiné exclusif pour décrocher votre prochaine alternance.</p>
            </td>
        </tr>

        <!-- CORPS DU MAIL -->
        <tr>
            <td style="padding: 20px 30px 40px 30px;">
                
                <!-- LISTE À PUCES STYLISÉE -->
                <table width="100%" cellpadding="0" cellspacing="0" style="margin: 10px 0;">
                    <tr>
                        <td style="padding: 10px 0;">
                            <span style="background-color: #ff8c00; color: white; padding: 2px 8px; border-radius: 4px; font-weight: bold; margin-right: 10px;">✔</span> 
                            <span style="font-size: 15px;">CV pro aux couleurs de l'établissement</span>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding: 10px 0;">
                            <span style="background-color: #ff8c00; color: white; padding: 2px 8px; border-radius: 4px; font-weight: bold; margin-right: 10px;">✔</span> 
                            <span style="font-size: 15px;">Optimisation complète du profil LinkedIn</span>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding: 10px 0;">
                            <span style="background-color: #ff8c00; color: white; padding: 2px 8px; border-radius: 4px; font-weight: bold; margin-right: 10px;">✔</span> 
                            <span style="font-size: 15px;">Stratégies de réseautage et feedback en direct</span>
                        </td>
                    </tr>
                </table>

                <!-- INFO LIEU DYNAMIQUE -->
                <div style="background-color: #f8f9fa; border-left: 4px solid #004a99; padding: 20px; margin-top: 30px;">
                    <p style="margin: 0; font-weight: bold; color: #004a99;"> Lieu de l'atelier :</p>
                    <p style="margin: 5px 0 0 0; font-size: 15px; color: #555;">
                        <!-- REMPLACER ICI PAR L'ADRESSE : 
                             Marseille : 39 Rue Sainte Cécile, 13005 Marseille 
                             Aix : Le Mansard B, 2 Pl. Romée de Villeneuve, 13090 Aix-en-Provence 
                        -->
                        <strong>{{INSERER_ICI_ADRESSE_CAMPUS}}</strong>
                    </p>
                </div>

                <!-- CALL TO ACTION -->
                <table width="100%" cellpadding="0" cellspacing="0" style="margin-top: 40px;">
                    <tr>
                        <td align="center">
                            <!-- === ATTENTION : INSERER VOTRE LIEN D'INSCRIPTION CI-DESSOUS === -->
                            <a href="{{VOTRE_LIEN_D_INSCRIPTION_ICI}}" style="background-color: #ff8c00; color: #ffffff; padding: 18px 35px; text-decoration: none; font-size: 18px; font-weight: bold; border-radius: 50px; display: inline-block;">
                                M'inscrire à l'atelier combiné
                            </a>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>

        <!-- FOOTER & RÉSEAUX SOCIAUX -->
        <tr>
            <td style="background-color: #001a38; color: #ffffff; padding: 40px 30px; text-align: center;">
                <p style="font-size: 14px; margin-bottom: 20px; opacity: 0.8;">Suivez les actualités de l'HECG :</p>
                
                <table align="center" cellpadding="0" cellspacing="0">
                    <tr>
                        <!-- LINKEDIN -->
                        <td style="padding: 0 15px;">
                            <a href="https://linkedin.com/in/hecg-aix-marseille-a21232239" target="_blank">
                                <img src="https://cdn-icons-png.flaticon.com/512/174/174857.png" alt="LinkedIn" width="30" style="display: block; filter: brightness(0) invert(1);">
                            </a>
                        </td>
                        <!-- INSTAGRAM -->
                        <td style="padding: 0 15px;">
                            <a href="https://www.instagram.com/hecgecole" target="_blank">
                                <img src="https://cdn-icons-png.flaticon.com/512/174/174855.png" alt="Instagram" width="30" style="display: block; filter: brightness(0) invert(1);">
                            </a>
                        </td>
                        <!-- TIKTOK -->
                        <td style="padding: 0 15px;">
                            <a href="https://www.tiktok.com/@hecgmarseille" target="_blank">
                                <img src="https://cdn-icons-png.flaticon.com/512/3046/3046121.png" alt="TikTok" width="30" style="display: block; filter: brightness(0) invert(1);">
                            </a>
                        </td>
                    </tr>
                </table>

                <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid rgba(255,255,255,0.1); font-size: 12px; opacity: 0.6;">
                    <p>© 2026 HECG - Hautes Etudes de Comptabilité et de Gestion<br>
                    <a href="https://hecg.fr/" style="color: #ffffff; text-decoration: underline;">www.hecg.fr</a></p>
                </div>
            </td>
        </tr>
    </table>

</body>
</html>`;
}
function getTemplateJPO() {
  return `<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>JPO HECG - Journée Portes Ouvertes</title>
</head>
<body style="margin: 0; padding: 0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f7f9; color: #333;">

    <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width: 600px; background-color: #ffffff; margin-top: 20px; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 10px rgba(0,0,0,0.05);">
        
        <!-- HEADER AVEC LOGO (LIEN VERS SITE WEB) -->
        <tr>
            <td align="center" style="padding: 30px 20px;">
                <a href="https://hecg.fr/" target="_blank">
                    <img src="https://hecg.fr/wp-content/uploads/2025/03/LOGO-FORMAT-PNG-1-1.png" alt="Logo HECG" width="180" style="display: block; border: 0;">
                </a>
            </td>
        </tr>

        <!-- HERO SECTION AVEC DÉGRADÉ -->
        <tr>
            <td style="background: linear-gradient(135deg, #004a99 0%, #002a5c 100%); padding: 45px 30px; text-align: center;">
                <h1 style="color: #ffffff; margin: 0; font-size: 28px; font-weight: 700; letter-spacing: -0.5px;">HECG a le plaisir de vous accueillir pour ses JPO</h1>
                <p style="color: #d1e3ff; font-size: 16px; margin-top: 15px;">Venez découvrir nos programmes et rencontrer nos équipes pédagogiques.</p>
            </td>
        </tr>

        <!-- CORPS DU MAIL -->
        <tr>
            <td style="padding: 20px 30px 40px 30px;">
                
                <p style="font-size: 16px; line-height: 1.6; color: #444; text-align: center;">
                    Vous hésitez encore sur votre orientation ? Profitez de cette journée pour poser toutes vos questions et visiter nos infrastructures.
                </p>

                <!-- LISTE DES AVANTAGES -->
                <table width="100%" cellpadding="0" cellspacing="0" style="margin: 10px 0;">
                    <tr>
                        <td style="padding: 10px 0;">
                            <span style="background-color: #ff8c00; color: white; padding: 2px 8px; border-radius: 4px; font-weight: bold; margin-right: 10px;">✔</span> 
                            <span style="font-size: 15px;">Présentation des cursus (Comptabilité, Gestion, Paye...)</span>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding: 10px 0;">
                            <span style="background-color: #ff8c00; color: white; padding: 2px 8px; border-radius: 4px; font-weight: bold; margin-right: 10px;">✔</span> 
                            <span style="font-size: 15px;">Échanges avec nos étudiants et diplômés</span>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding: 10px 0;">
                            <span style="background-color: #ff8c00; color: white; padding: 2px 8px; border-radius: 4px; font-weight: bold; margin-right: 10px;">✔</span> 
                            <span style="font-size: 15px;">Aide à la recherche d'alternance</span>
                        </td>
                    </tr>
                </table>

                <!-- INFO PRATIQUES (DATE, HEURE, LIEU) -->
                <div style="background-color: #f8f9fa; border-left: 4px solid #004a99; padding: 25px; margin-top: 30px; border-radius: 0 4px 4px 0;">
                    <p style="margin: 0; font-weight: bold; color: #004a99; font-size: 18px;"> Informations pratiques :</p>
                    <p style="margin: 10px 0 5px 0; font-size: 15px; color: #333;">
                        <strong>Date :</strong> {{INSERER_DATE_ICI}}
                    </p>
                    <p style="margin: 5px 0 5px 0; font-size: 15px; color: #333;">
                        <strong>Horaires :</strong> {{INSERER_HORAIRES_ICI}}
                    </p>
                    <p style="margin: 5px 0 0 0; font-size: 15px; color: #333;">
                        <strong>Campus :</strong> {{INSERER_CAMPUS_ET_ADRESSE_ICI}}
                    </p>
                </div>

                <!-- CALL TO ACTION -->
                <table width="100%" cellpadding="0" cellspacing="0" style="margin-top: 40px;">
                    <tr>
                        <td align="center">
                            <!-- === ATTENTION : INSERER LE LIEN D'INSCRIPTION JPO CI-DESSOUS === -->
                            <a href="{{LIEN_INSCRIPTION_JPO_ICI}}" style="background-color: #ff8c00; color: #ffffff; padding: 18px 35px; text-decoration: none; font-size: 18px; font-weight: bold; border-radius: 50px; display: inline-block; box-shadow: 0 4px 6px rgba(255,140,0,0.3);">
                                Je m'inscris à la JPO
                            </a>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>

        <!-- FOOTER & RÉSEAUX SOCIAUX -->
        <tr>
            <td style="background-color: #001a38; color: #ffffff; padding: 40px 30px; text-align: center;">
                <p style="font-size: 14px; margin-bottom: 20px; opacity: 0.8;">Restez connecté avec l'HECG :</p>
                
                <table align="center" cellpadding="0" cellspacing="0">
                    <tr>
                        <!-- LINKEDIN -->
                        <td style="padding: 0 15px;">
                            <a href="https://linkedin.com/in/hecg-aix-marseille-a21232239" target="_blank">
                                <img src="https://cdn-icons-png.flaticon.com/512/174/174857.png" alt="LinkedIn" width="30" style="display: block; filter: brightness(0) invert(1);">
                            </a>
                        </td>
                        <!-- INSTAGRAM -->
                        <td style="padding: 0 15px;">
                            <a href="https://www.instagram.com/hecgecole" target="_blank">
                                <img src="https://cdn-icons-png.flaticon.com/512/174/174855.png" alt="Instagram" width="30" style="display: block; filter: brightness(0) invert(1);">
                            </a>
                        </td>
                        <!-- TIKTOK -->
                        <td style="padding: 0 15px;">
                            <a href="https://www.tiktok.com/@hecgmarseille" target="_blank">
                                <img src="https://cdn-icons-png.flaticon.com/512/3046/3046121.png" alt="TikTok" width="30" style="display: block; filter: brightness(0) invert(1);">
                            </a>
                        </td>
                    </tr>
                </table>

                <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid rgba(255,255,255,0.1); font-size: 12px; opacity: 0.6;">
                    <p>© 2026 HECG - Hautes Etudes de Comptabilité et de Gestion<br>
                    <a href="https://hecg.fr/" style="color: #ffffff; text-decoration: underline;">www.hecg.fr</a></p>
                </div>
            </td>
        </tr>
    </table>

</body>
</html>`;
}
function getTemplateSalon() {
  return `<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>HECG au Salon</title>
</head>
<body style="margin: 0; padding: 0; background-color: #001a38; font-family: 'Segoe UI', Arial, sans-serif;">

    <!-- CONTENEUR PRINCIPAL AVEC DÉGRADÉ DE BLEU -->
    <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="background: linear-gradient(180deg, #004a99 0%, #001a38 100%); min-height: 100vh;">
        <tr>
            <td align="center" style="padding: 40px 10px;">
                
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width: 600px;">
                    
                    <!-- LOGO -->
                    <tr>
                        <td align="center" style="padding-bottom: 50px;">
                            <a href="https://hecg.fr/" target="_blank">
                                <img src="https://hecg.fr/wp-content/uploads/2025/03/LOGO-FORMAT-PNG-1-1.png" alt="Logo HECG" width="160" style="display: block; border: 0; filter: brightness(0) invert(1);">
                            </a>
                        </td>
                    </tr>

                    <!-- CONTENU VISUEL / PRÉSENTATION -->
                    <tr>
                        <td align="center" style="padding: 0 20px;">
                            <h1 style="color: #ffffff; font-size: 32px; font-weight: 800; line-height: 1.2; margin: 0; letter-spacing: -1px;">
                                Venez nous retrouver.
                            </h1>
                            
                            <div style="width: 50px; height: 4px; background-color: #FF7A00; margin: 25px 0;"></div>
                            
                            <p style="color: #d1e3ff; font-size: 18px; line-height: 1.6; margin: 0; font-weight: 400;">
                                {{TEXTE_ACCROCHE}}
                            </p>
                        </td>
                    </tr>

                    <!-- CARD INFO STYLISÉE -->
                    <tr>
                        <td style="padding: 50px 20px 0 20px;">
                            <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: rgba(255, 255, 255, 0.05); border: 1px solid rgba(255, 255, 255, 0.1); border-radius: 20px; backdrop-filter: blur(10px);">
                                <tr>
                                    <td style="padding: 35px; text-align: center;">
                                        <p style="color: #FF7A00; font-weight: bold; text-transform: uppercase; letter-spacing: 2px; margin: 0 0 15px 0; font-size: 13px;">
                                            Prochain Événement
                                        </p>
                                        <h2 style="color: #ffffff; font-size: 24px; margin: 0; font-weight: 600;">
                                            {{NOM_DU_SALON_OU_EVENEMENT}}
                                        </h2>
                                        
                                        <table border="0" cellpadding="0" cellspacing="0" width="100%" style="margin-top: 25px;">
                                            <tr>
                                                <td align="center" style="color: #ffffff; font-size: 16px; opacity: 0.9;">
                                                    <span style="display: inline-block; padding: 5px 10px; border-radius: 5px; background: rgba(255,255,255,0.1); margin: 5px;"> {{DATE_ET_HORAIRES}}</span>
                                                    <span style="display: inline-block; padding: 5px 10px; border-radius: 5px; background: rgba(255,255,255,0.1); margin: 5px;"> {{LIEU_ET_ADRESSE}}</span>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>

                    <!-- BOUTON ACTION -->
                    <tr>
                        <td align="center" style="padding: 50px 20px;">
                            <a href="{{LIEN_INSCRIPTION_ICI}}" style="background-color: #FF7A00; color: #ffffff; padding: 20px 45px; text-decoration: none; font-size: 18px; font-weight: bold; border-radius: 12px; display: inline-block; box-shadow: 0 10px 20px rgba(255, 122, 0, 0.3);">
                                Lien de l'organisateur
                            </a>
                        </td>
                    </tr>

                    <!-- FOOTER & RÉSEAUX -->
                    <tr>
                        <td align="center" style="padding: 40px 20px; border-top: 1px solid rgba(255, 255, 255, 0.1);">
                            <p style="color: #ffffff; font-size: 14px; opacity: 0.6; margin-bottom: 25px;">Retrouvez-nous sur nos réseaux sociaux</p>
                            
                            <table border="0" cellpadding="0" cellspacing="0">
                                <tr>
                                    <td style="padding: 0 15px;">
                                        <a href="https://linkedin.com/in/hecg-aix-marseille-a21232239" target="_blank">
                                            <img src="https://cdn-icons-png.flaticon.com/512/174/174857.png" alt="LinkedIn" width="28" style="filter: brightness(0) invert(1);">
                                        </a>
                                    </td>
                                    <td style="padding: 0 15px;">
                                        <a href="https://www.instagram.com/hecgecole" target="_blank">
                                            <img src="https://cdn-icons-png.flaticon.com/512/174/174855.png" alt="Instagram" width="28" style="filter: brightness(0) invert(1);">
                                        </a>
                                    </td>
                                    <td style="padding: 0 15px;">
                                        <a href="https://www.tiktok.com/@hecgmarseille" target="_blank">
                                            <img src="https://cdn-icons-png.flaticon.com/512/3046/3046121.png" alt="TikTok" width="28" style="filter: brightness(0) invert(1);">
                                        </a>
                                    </td>
                                </tr>
                            </table>

                            <div style="margin-top: 40px; color: #ffffff; font-size: 11px; opacity: 0.4; line-height: 1.5;">
                                © 2026 HECG - Hautes Etudes de Comptabilité et de Gestion<br>
                                <a href="https://hecg.fr/" style="color: #ffffff;">www.hecg.fr</a>
                            </div>
                        </td>
                    </tr>

                </table>
                
            </td>
        </tr>
    </table>

</body>
</html>`;
}
// ==========================================
// VALIDATION DES PRÉSENCES ET MISE À JOUR AUTO
// ==========================================
function handleMarkPresence(p, ss) {
  // 1. Mise à jour de l'onglet INSCRIPTIONS (Le Pointage)
  const sI = getSheetSafe(ss, "INSCRIPTIONS");
  const dI = sI.getDataRange().getValues();
  let participantTrouve = false;

  for(let i=1; i<dI.length; i++) {
    // On valide la présence en utilisant l'email d'origine (celui de l'inscription)
    if(dI[i][0] === p.eventId && String(dI[i][1]).toLowerCase().trim() === String(p.emailOriginal).toLowerCase().trim()) {
      sI.getRange(i+1, 6).setValue("Présent"); // Colonne F
      participantTrouve = true;
      break;
    }
  }

  // 2. Rapprochement et mise à jour de l'onglet ETUDIANTS
  const sE = getSheetSafe(ss, "ETUDIANTS");
  const dE = sE.getDataRange().getValues();
  let etudiantMisAJour = false;
  
  // On utilise l'email de rapprochement (soit l'original, soit celui forcé par l'équipe Com)
  const targetEmail = String(p.emailEtudiant).toLowerCase().trim();

  for(let j=1; j<dE.length; j++) {
    const mailEtu = String(dE[j][4] || "").toLowerCase().trim(); // Colonne E

    if (mailEtu === targetEmail && mailEtu !== "") {
      // ÉTUDIANT TROUVÉ ! On met à jour selon le type d'événement
      const typeEvt = p.typeEvt || "";
      
      if (typeEvt.includes("CV") || typeEvt === "Atelier Mixte") {
        sE.getRange(j+1, 16).setValue("Oui"); // Colonne P (Fait CV)
      }
      if (typeEvt.includes("LinkedIn") || typeEvt === "Atelier Mixte") {
        sE.getRange(j+1, 17).setValue("Oui"); // Colonne Q (Fait LinkedIn)
      }
      
      // On ajoute une petite note dans son carnet de route pour garder l'historique
      getSheetSafe(ss, "Carnet_route").appendRow([dE[j][0], "Système Auto", new Date(), `📍 A participé à l'événement : ${typeEvt}. Profil mis à jour.`]);
      
      etudiantMisAJour = true;
      break;
    }
  }

  // 3. Retour d'information à l'interface
  if (participantTrouve) {
    if (etudiantMisAJour) {
      return createJsonResponse({ success: true, message: "Présence validée et profil étudiant mis à jour avec succès !" });
    } else {
      return createJsonResponse({ success: true, message: "Présence validée (Prospect externe ou étudiant non trouvé dans la base)." });
    }
  } else {
    return createJsonResponse({ success: false, message: "Inscrit introuvable pour cet événement." });
  }
}
function handleGetComEvents(p, ss) {
  const sheet = ss.getSheetByName("CAMPAGNES_COM");
  if (!sheet) return createJsonResponse({ success: true, data: [] });
  
  const data = sheet.getDataRange().getValues();
  const list = [];
  for(let i = 1; i < data.length; i++) {
    if(!data[i][0]) continue;
    const row = data[i]; // <-- LIGNE CRUCIALE AJOUTÉE ICI
    list.push({
                id: row[0],
                nom: row[1],
                objet: row[2],
                cible: row[4],
                type: row[5],
                dateEvt: row[12],
                campus: row[11],
                statut: row[7],
                pjUrl: row[8],
                flyerUrl: row[9],
                lienExterne: row[10],
                horaires: row[13],
                emailsLibres: row[14],
                accroche: row[15] // <-- AJOUTÉ ICI (Colonne P)
            });
  }
  return createJsonResponse({ success: true, data: list.reverse() });
}

function handleSaveComEvent(p, ss) {
  let sheet = ss.getSheetByName("CAMPAGNES_COM");
  if (!sheet) {
    sheet = ss.insertSheet("CAMPAGNES_COM");
    sheet.appendRow(["ID","Nom","Objet","Corps_Accroche","Cible","Type_Action","Dernier_Envoi","Statut","PJ_URL","Flyer_URL","Lien_Externe","Campus","Date_Evt","Horaires","Emails_Libres","Accroche"]);
  }
  
  const isNew = !p.id;
  const id = isNew ? "COM" + new Date().getTime() : p.id;
  const cibleStr = typeof p.cible === 'object' ? JSON.stringify(p.cible) : p.cible;
  
  const rowData = [
    id, p.nom, p.objet, "", cibleStr, p.type, 
    p.dernierEnvoi || "", p.statut || "Brouillon", p.pjUrl || "", p.flyerUrl || "", 
    p.lienExterne || "", p.campus, p.dateEvt, p.horaires, p.emailsLibres || "", 
    p.accroche || "" // <-- ACCROCHE BIEN PLACÉE EN 16ÈME POSITION (COLONNE P)
  ];

  if (isNew) {
    sheet.appendRow(rowData);
  } else {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === id) {
        sheet.getRange(i + 1, 1, 1, rowData.length).setValues([rowData]);
        break;
      }
    }
  }
  const evtDateStr = p.dateEvt ? new Date(p.dateEvt).toLocaleDateString('fr-FR') : "sans date";
  const logDetail = (p.nom || "Événement") + " | Type: " + (p.type || "?") + " | Date: " + evtDateStr + " | Statut: " + (p.statut || "Brouillon");
  logAction(p.idAuteur, p.roleAuteur, isNew ? "Création" : "Modification", "Com & Événements", id, logDetail);
  return createJsonResponse({ success: true, message: "Événement enregistré avec succès." });
}

// CETTE FONCTION EST LE POINT D'ENTRÉE UNIQUE POUR LE FORMULAIRE
function doPostManual(json) {
  try {
    var p = JSON.parse(json);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // On récupère le nom de l'événement dans CAMPAGNES_COM
    var sheetCom = ss.getSheetByName("CAMPAGNES_COM");
    var eventName = "Événement direct";
    if (sheetCom) {
      var dataCom = sheetCom.getDataRange().getValues();
      for (var i = 1; i < dataCom.length; i++) {
        if (String(dataCom[i][0]) === String(p.eventId)) {
          eventName = dataCom[i][1]; // Nom en colonne B
          break;
        }
      }
    }

    // On écrit dans PROSPECTS (A à H)
    var sheetPros = ss.getSheetByName("PROSPECTS");
    if (!sheetPros) {
      sheetPros = ss.insertSheet("PROSPECTS");
      sheetPros.appendRow(["ID_Prospect", "Nom", "Prenom", "Email", "Telephone", "Source", "Commentaire", "ID_Evenement"]);
    }

    sheetPros.appendRow([
      "PR" + new Date().getTime(), // A: ID
      p.nom,                        // B: Nom
      p.prenom,                     // C: Prenom
      p.email,                      // D: Email
      p.tel,                        // E: Tel
      eventName,                    // F: Source (Nom de l'event)
      "Inscrit via formulaire",     // G: Commentaire
      p.eventId                     // H: ID de l'event (Lien pour le pointage)
    ]);

    return "SUCCESS";
  } catch (e) {
    throw new Error("Erreur lors de l'inscription : " + e.toString());
  }
}
function handleGetParticipants(p, ss) {
  try {
    const sheetPros = ss.getSheetByName("PROSPECTS");
    if (!sheetPros) return createJsonResponse({ success: true, data: [] });

    const dataPros = sheetPros.getDataRange().getValues();
    const sheetEtu = ss.getSheetByName("ETUDIANTS");
    const dataEtu = sheetEtu ? sheetEtu.getDataRange().getValues() : [];
    
    const participants = [];
    const eventIdCible = String(p.eventId).trim();

    // On parcourt l'onglet PROSPECTS (à partir de la ligne 2)
    for (let i = 1; i < dataPros.length; i++) {
      // COLONNE H (index 7) : C'est là qu'est stocké l'ID COM...
      const idEvtInscrit = String(dataPros[i][7] || "").trim();

      if (idEvtInscrit === eventIdCible) {
        const emailInscrit = String(dataPros[i][3] || "").toLowerCase().trim();
        
        // Vérification si c'est un étudiant connu (Rapprochement auto)
        let isEtudiant = false;
        for (let j = 1; j < dataEtu.length; j++) {
          if (String(dataEtu[j][4]).toLowerCase().trim() === emailInscrit) {
            isEtudiant = true;
            break;
          }
        }

        participants.push({
          id: dataPros[i][0],
          nom: dataPros[i][1],
          prenom: dataPros[i][2],
          email: emailInscrit,
          tel: dataPros[i][4],
          presence: dataPros[i][6] === "Présent" ? "Présent" : "Inscrit",
          isEtudiant: isEtudiant
        });
      }
    }

    return createJsonResponse({ success: true, data: participants });
  } catch (e) {
    return createJsonResponse({ success: false, message: e.toString() });
  }
}
function handleGetStudentsMinList(ss) {
  const sheet = ss.getSheetByName("ETUDIANTS");
  const data = sheet.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    const mail = String(data[i][4] || "").trim();
    if (mail && mail.includes('@')) {
      list.push({
        id:     String(data[i][0]  || ""),
        nom:    String(data[i][2]  || ""),   // col C
        prenom: String(data[i][3]  || ""),   // col D
        email:  mail,                          // col E
        classe: String(data[i][14] || "").trim(), // col O
        statut: String(data[i][13] || "").toLowerCase(), // col N
        typo:   String(data[i][20] || "").toLowerCase()  // col U
      });
    }
  }
  list.sort((a, b) => a.nom.localeCompare(b.nom));
  return createJsonResponse({ success: true, data: list });
}



function handlePreviewComMail(p, ss) {
  try {
    const wsEtu = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues();
    const wsPart = getSheetSafe(ss, "Partenariat").getDataRange().getValues();
    let contactsList = [];

    let cnf = p.cible || {};
    if (typeof cnf === 'string') { try { cnf = JSON.parse(cnf); } catch(e) { cnf = {}; } }

    // --- CHEMIN SPÉCIAL ATELIERS (auto-ciblage) ---
    const typeEvt = String(p.type || "").trim();
    const isAtelier = typeEvt === "Atelier CV" || typeEvt === "Atelier LinkedIn" || typeEvt === "Atelier Mixte";
    if (isAtelier) {
      const isAtelierCVOnly  = typeEvt === "Atelier CV";
      const isAtelierLIOnly  = typeEvt === "Atelier LinkedIn";
      const isAtelierMixte   = typeEvt === "Atelier Mixte";
      const campusCible = String(p.campus || "").toLowerCase().trim();
      // Mapping noms du formulaire → valeurs réelles de la colonne W
      var CAMPUS_MAP = { "marseille": "hecg", "aix": "hecg_aix" };
      var campusCibleFeuille = CAMPUS_MAP[campusCible] || campusCible;
      const classesCiblesAtelier = Array.isArray(cnf.classesCibles) ? cnf.classesCibles : [];
      for (let i = 1; i < wsEtu.length; i++) {
        const row = wsEtu[i];
        const mail = String(row[4] || "").trim().toLowerCase();
        if (!mail.includes('@')) continue;
        const typo   = String(row[20] || "").toLowerCase();
        const statut = String(row[13] || "").toLowerCase();
        const faitCV = String(row[15] || "").toLowerCase().includes("oui");
        const faitLI = String(row[16] || "").toLowerCase().includes("oui");
        const campus = String(row[22] || "").toLowerCase();
        const classe = String(row[14] || "").trim();
        // Inscrit ou Préinscrit uniquement (pas sorti, pas prospect)
        if (!typo.includes("inscrit") || typo.includes("sorti")) continue;
        // En recherche (exclut alternance, initial, etc.)
        if (!statut.includes("recherche")) continue;
        // Filtre campus : correspondance exacte avec la valeur de la feuille
        if (campusCible && campus && campus !== campusCibleFeuille) continue;
        // Filtre par classe si une sélection partielle est précisée
        if (classesCiblesAtelier.length > 0 && !classesCiblesAtelier.includes(classe)) continue;
        // Déjà passé l'atelier ?
        if (isAtelierCVOnly && faitCV) continue;
        if (isAtelierLIOnly && faitLI) continue;
        if (isAtelierMixte && faitCV && faitLI) continue;
        contactsList.push({ email: mail, nom: (row[3] || "") + " " + (row[2] || ""), source: "Atelier", classe: classe });
      }
      // Emails ajoutés manuellement
      if (p.emailsLibres) {
        String(p.emailsLibres).split(',').forEach(function(raw) {
          const mail = raw.trim().toLowerCase();
          if (mail.includes('@')) contactsList.push({ email: mail, nom: mail, source: "Ajouté" });
        });
      }
      const unique = contactsList.filter(function(v, i, a) { return a.findIndex(function(t) { return t.email === v.email; }) === i; });
      return createJsonResponse({ success: true, data: unique });
    }

    const cibleSortis = cnf.b2c_sortis === true;
    const motifSelectionne = String(cnf.motif_sortie || "Tous").trim();

    // 1. FILTRAGE ETUDIANTS
    for(let i = 1; i < wsEtu.length; i++) {
      const row = wsEtu[i];
      let mail = String(row[4] || "").trim().toLowerCase();
      if(!mail.includes('@')) continue;

      let status = String(row[13] || "").toLowerCase(); // Colonne N (Statut)
      let motifEtu = String(row[19] || "").trim();      // Colonne T (Motif sortie)
      let typo = String(row[20] || "").toLowerCase();   // Colonne U (Typologie)
      
      let ok = false;

      if (cibleSortis) {
        // --- MODE RÉCUPÉRATION (SORTIS) ---
        if (typo.includes("sorti")) {
          if (motifSelectionne === "Tous" || motifEtu === motifSelectionne) ok = true;
        }
      } else {
        // --- MODE CLASSIQUE (On exclut d'office les sortis et refusés) ---
        if (!typo.includes("sorti") && !typo.includes("refus")) {
          
          // 1. Etudiants en recherche : (Statut Recherche) ET (Inscrit OU Préinscrit)
          if (cnf.target_recherche && status.includes("recherche") && typo.includes("inscrit")) {
            ok = true;
          }

          // 2. Etudiants préinscrits uniquement (Peu importe le statut)
          if (cnf.b2c_preinscrits && typo.includes("préin")) {
            ok = true;
          }

          // 3. Etudiants prospects uniquement
          if (cnf.b2c_prospects && typo.includes("prospect")) {
            ok = true;
          }

          // 4. Inscrits définitivement (Tous les inscrits, peu importe le statut)
          if (cnf.b2c_inscrits && typo.includes("inscrit")) ok = true;
          
        }
      }

      if (ok) contactsList.push({ email: mail, nom: (row[3] || "") + " " + (row[2] || ""), source: "Étudiant" });
    }

    // 2. FILTRAGE ENTREPRISES
    // --- FILTRAGE ENTREPRISES ---
    if (!cibleSortis) {
  for(let j = 1; j < wsPart.length; j++) {
    let rowP = wsPart[j];
    let mailP = String(rowP[3] || "").trim(); // Colonne D
    if(!mailP.includes('@')) continue;

    // On vérifie si la colonne J (index 9) n'est pas vide
    let cellAlt = String(rowP[9] || "").trim();
    let nbAlt = (cellAlt === "" || cellAlt === "0") ? 0 : 1; 

    let okP = false;
    if (cnf.b2b_avec && nbAlt > 0) okP = true;
    if (cnf.b2b_sans && nbAlt === 0) okP = true;

    if (okP) contactsList.push({ email: mailP.toLowerCase(), nom: rowP[1] || "Entreprise", source: "Entreprise" });
  }
}
    
    // Suppression des doublons
    const unique = contactsList.filter((v, i, a) => a.findIndex(t => t.email === v.email) === i);
    return createJsonResponse({ success: true, data: unique });

  } catch(e) { 
    return createJsonResponse({ success: false, message: "Erreur : " + e.toString() }); 
  }
}



function handleDeleteComEvent(p, ss) {
  const ws = getSheetSafe(ss, "CAMPAGNES_COM");
  const data = ws.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === p.id) {
      const nomEvt = data[i][1] || "?";
      const typeEvt = data[i][5] || "?";
      ws.getRange(i + 1, 8).setValue("Archivé");
      logAction(p.idAuteur, p.roleAuteur, "Suppression", "Com & Événements", p.id, "Événement archivé : " + nomEvt + " (" + typeEvt + ")");
      break;
    }
  }
  return createJsonResponse({ success: true, message: "Événement rangé dans l'historique." });
}


function sendWeeklyReminders() {
  const ss = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");
  const wsCom = ss.getSheetByName("CAMPAGNES_COM");
  const dataCom = wsCom.getDataRange().getValues();
  
  // 1. On cherche la configuration "Relance Hebdo"
  let config = dataCom.find(row => row[0] === "RELANCE_HEBDO");
  if (!config || config[7] !== "Actif") return; // On ne fait rien si elle n'est pas "Actif"

  const subject = config[3]; // L'objet que vous avez modifié dans l'appli
  const hook = config[4];    // L'accroche / texte (si besoin)

  // 2. Récupération des offres de l'onglet OFFRES
  const wsOffres = ss.getSheetByName("OFFRES");
  const offres = wsOffres.getDataRange().getValues();
  let offresHtml = "<div style='font-family: Arial; line-height: 1.6;'>";
  
  for(let j=1; j < offres.length; j++) {
    // On ne prend que les offres "Reçue" (Col J = index 9)
    if(String(offres[j][9]).includes("Reçue")) {
       offresHtml += `
         <div style="margin-bottom: 15px; padding: 10px; border-left: 4px solid #F26522; background: #f9f9f9;">
           <b style="color: #1A4E8A;">${offres[j][14]}</b> chez <i>${offres[j][4]}</i><br>
           <a href="${offres[j][1]}" style="color: #F26522; font-weight: bold; text-decoration: none;">Voir l'offre →</a>
         </div>`;
    }
  }
  offresHtml += "</div>";

  // 3. Envoi aux étudiants ciblés (En recherche + Inscrits/Prein)
  const wsEtu = ss.getSheetByName("ETUDIANTS");
  const etuData = wsEtu.getDataRange().getValues();

  etuData.forEach((row, i) => {
    if (i === 0) return;
    let mail = String(row[4]).trim();
    let prenom = row[3];
    let statut = String(row[13]).toLowerCase();
    let typo = String(row[20]).toLowerCase();

    if (statut.includes("recherche") && typo.includes("inscrit")) {
      let body = `<h3>Bonjour ${prenom},</h3>
                  <p>La recherche d'alternance demande de la persévérance. Voici les dernières opportunités partagées cette semaine :</p>
                  ${offresHtml}
                  <p>N'oublie pas de relancer les entreprises après 5 jours !<br><b>L'équipe HECG</b></p>`;
      
      GmailApp.sendEmail(mail, subject, "", { htmlBody: body });
    }
  });

  // Mise à jour de la date de "Dernier Envoi" (Col G)
  for(let k=1; k<dataCom.length; k++) {
    if(dataCom[k][0] === "RELANCE_HEBDO") {
      wsCom.getRange(k+1, 7).setValue(new Date());
      break;
    }
  }
}

function processWeeklyCampagnes() {
  const ss = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");
  const sheetCamp = ss.getSheetByName("CAMPAGNES"); // Onglet lié à campagnes.html
  const dataCamp = sheetCamp.getDataRange().getValues();
  
  // 1. Trouver la campagne active hebdomadaire
  const camp = dataCamp.find(r => r[2] === "Hebdomadaire" && r[8] === "Active");
  if (!camp) return;

  const subjectTemplate = camp[6]; // Colonne Objet
  const bodyTemplate = camp[7];    // Colonne Message/Corps

  // 2. Préparer la liste des offres (depuis l'onglet OFFRES)
  const offres = ss.getSheetByName("OFFRES").getDataRange().getValues();
  let htmlOffres = "<ul style='color: #1A4E8A; font-weight: bold;'>";
  offres.slice(1).forEach(o => {
    if(o[9] === "Reçue") { // Filtrage des offres actives
      htmlOffres += `<li>${o[14]} chez ${o[4]} - <a href='${o[1]}'>Postuler</a></li>`;
    }
  });
  htmlOffres += "</ul>";

  // 3. Envoyer à chaque étudiant ciblé
  const etudiants = ss.getSheetByName("ETUDIANTS").getDataRange().getValues();
  etudiants.slice(1).forEach(etu => {
    const statut = String(etu[13]).toLowerCase();
    const typo = String(etu[20]).toLowerCase();
    
    // Filtre : Recherche + (Inscrit ou Préinscrit)
    if (statut.includes("recherche") && typo.includes("inscrit")) {
      const email = etu[4];
      const prenom = etu[3];

      // Remplacement des variables dynamiques
      const finalSubject = subjectTemplate.replace("{{PRENOM}}", prenom);
      const finalBody = bodyTemplate
        .replace("{{PRENOM}}", prenom)
        .replace("{{LISTE_OFFRES}}", htmlOffres);

      GmailApp.sendEmail(email, finalSubject, "", { htmlBody: finalBody });
    }
  });
}

// Créer une campagne
function handleCreateCampaign(p, ss) {
  try {
    const sheet = getSheetSafe(ss, "CAMPAGNES");
    const newId = "CMP-" + Math.floor(10000 + Math.random() * 90000);
    const newRow = [
      newId, p.nom, p.frequence, new Date(), "", p.objet, p.message, "Active", p.cibles
    ];
    sheet.appendRow(newRow);
    return createJsonResponse({ success: true, message: "Campagne créée !" });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// Mettre à jour (Modifier)
function handleUpdateCampaign(p, ss) {
  try {
    const sheet = getSheetSafe(ss, "CAMPAGNES");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === p.id) {
        sheet.getRange(i + 1, 2, 1, 8).setValues([[p.nom, p.frequence, data[i][3], data[i][4], p.objet, p.message, p.statut, p.cibles]]);
        break;
      }
    }
    return createJsonResponse({ success: true, message: "Campagne mise à jour !" });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// Fonction pour lancer manuellement la campagne depuis le bouton 🚀
function handleRunCampagneManual(p, ss) {
  try {
    // On appelle la logique globale de traitement hebdomadaire
    processWeeklyCampagnes(); 
    return createJsonResponse({ success: true, message: "Lancement de la campagne effectué !" });
  } catch (e) {
    return createJsonResponse({ success: false, message: e.toString() });
  }
}

// Fonction pour envoyer un test à l'admin
function handleTestCampaign(p, ss) {
  try {
    const sheet = getSheetSafe(ss, "CAMPAGNES");
    const data = sheet.getDataRange().getValues();
    const camp = data.find(r => String(r[0]) === String(p.campaignId));
    
    if (!camp) throw new Error("Campagne introuvable.");

    // Récupération de l'email de test dans les réglages
    const settings = handleGetSettings(ss);
    const emailDest = JSON.parse(settings.getContent()).email_test;

    GmailApp.sendEmail(emailDest, "[TEST] " + camp[5], "", { htmlBody: camp[6] });
    return createJsonResponse({ success: true, message: "Email de test envoyé à " + emailDest });
  } catch (e) {
    return createJsonResponse({ success: false, message: e.toString() });
  }
}

// Fonction pour dupliquer une campagne existante
function handleDuplicateCampaign(p, ss) {
  try {
    const sheet = getSheetSafe(ss, "CAMPAGNES");
    const data = sheet.getDataRange().getValues();
    const original = data.find(r => String(r[0]) === String(p.campaignId));
    
    if (!original) throw new Error("Original introuvable.");

    const newId = "CMP-" + Math.floor(10000 + Math.random() * 90000);
    const newRow = [...original];
    newRow[0] = newId;
    newRow[1] = original[1] + " (Copie)";
    newRow[3] = new Date();
    newRow[4] = "";
    newRow[7] = "À lancer"; // Statut par défaut
    
    sheet.appendRow(newRow);
    return createJsonResponse({ success: true, message: "Campagne dupliquée !" });
  } catch (e) {
    return createJsonResponse({ success: false, message: e.toString() });
  }
}

function runScheduledCampaigns() {
  const ss = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");
  const sheetCamp = getSheetSafe(ss, "CAMPAGNES");
  const dataCamp = sheetCamp.getDataRange().getValues();

  // Normalisation à minuit pour éviter les décalages d'heure dus au trigger
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  let nbLancees = 0;
  let nbIgnorees = 0;
  let nbErreurs = 0;
  const erreurDetails = [];

  for (let i = 1; i < dataCamp.length; i++) {
    const row = dataCamp[i];
    const idCmp = row[0];
    const freq = row[2]; // Colonne C
    const statut = row[7]; // Colonne H

    if (statut !== "Active") { nbIgnorees++; continue; }

    let dernierEnvoi = null;
    if (row[4]) {
      dernierEnvoi = new Date(row[4]);
      dernierEnvoi.setHours(0, 0, 0, 0); // Normalisation identique
    }

    let doitPartir = false;

    if (!dernierEnvoi) {
      doitPartir = true;
    } else {
      const joursEcoules = Math.floor((today - dernierEnvoi) / (1000 * 60 * 60 * 24));

      if (freq === "Hebdomadaire" && joursEcoules >= 7) {
        doitPartir = true;
      } else if (freq === "Mensuelle" && joursEcoules >= 28) {
        doitPartir = true;
      } else if (freq === "Compte à rebours" && joursEcoules >= 1) {
        doitPartir = true;
      }
      // "Une fois" avec dernierEnvoi existant → ne repart pas
    }

    if (doitPartir) {
      try {
        Logger.log("Lancement de la campagne : " + row[1]);
        executeSendingLogic(row, ss);
        sheetCamp.getRange(i + 1, 5).setValue(new Date());
        if (freq === "Une fois") sheetCamp.getRange(i + 1, 8).setValue("Terminée");
        nbLancees++;
      } catch (e) {
        nbErreurs++;
        erreurDetails.push("Campagne [" + idCmp + "] : " + e.toString());
      }
    } else {
      nbIgnorees++;
    }
  }

  logExecutionAuto(ss, "runScheduledCampaigns", nbLancees, nbIgnorees, nbErreurs, erreurDetails.join(" | "));
}
// ==========================================
// MOTEUR DES ALERTES PARTENAIRES
// ==========================================

// 1. Lire et afficher les alertes actives
function handleGetAllAlerts(p, ss) {
  try {
    const sA = getSheetSafe(ss, "ALERTES_PARTENAIRES");
    const sP = getSheetSafe(ss, "PARTENARIAT");
    const dA = sA.getDataRange().getValues();
    const dP = sP.getDataRange().getValues();

    // Dictionnaire pour trouver le nom de l'entreprise via son ID (Sécurisé majuscules/espaces)
    const partMap = {};
    for(let j = 1; j < dP.length; j++) {
      // On associe l'ID (Col B) au NOM (Col C) en ignorant les espaces/majuscules
      let cleanId = String(dP[j][1] || "").trim().toUpperCase();
      if(cleanId) partMap[cleanId] = String(dP[j][2] || "Entreprise sans nom").trim(); 
    }

    const list = [];
    for(let i = 1; i < dA.length; i++) {
      // On ignore les lignes vides, terminées ou annulées
      if(!dA[i][0] || dA[i][5] === "Terminé" || dA[i][5] === "Annulé") continue; 
      
      let dateObj = new Date(dA[i][2]);
      let idAlerte = String(dA[i][1]).trim().toUpperCase();
      let nomTrouve = partMap[idAlerte] ? partMap[idAlerte] : ("ID : " + dA[i][1]);
      
      list.push({
        id: dA[i][0],
        partId: dA[i][1],
        nomPartenaire: nomTrouve,
        dateRaw: dateObj.toISOString(),
        date: dateObj.toLocaleDateString('fr-FR'),
        msg: dA[i][3],
        freq: dA[i][4],
        destinataire: dA[i][6] || "" // <-- ON LIT BIEN L'EMAIL ICI
      });
    }
    
    // On trie pour avoir les plus urgentes en premier
    list.sort((a, b) => new Date(a.dateRaw) - new Date(b.dateRaw));
    
    return createJsonResponse({ success: true, data: list });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// 2. Marquer comme Terminé (ou relancer si récurrent)
function handleCompleteAlert(p, ss) {
  try {
    const sA = getSheetSafe(ss, "ALERTES_PARTENAIRES");
    const dA = sA.getDataRange().getValues();

    for(let i = 1; i < dA.length; i++) {
      if(String(dA[i][0]) === p.id) {
        const idPart = String(dA[i][1]);
        const msgAlerte = String(dA[i][3] || "").substring(0, 60);
        if(dA[i][4] === "Unique") {
          sA.getRange(i+1, 6).setValue("Terminé");
          logAction(p.idAuteur, p.roleAuteur, "Action", "Alertes Partenaires", idPart, "Alerte marquée Terminée : « " + msgAlerte + " »");
        } else {
          let oldDate = new Date(dA[i][2]);
          if(dA[i][4] === "Hebdomadaire") oldDate.setDate(oldDate.getDate() + 7);
          if(dA[i][4] === "Mensuelle") oldDate.setMonth(oldDate.getMonth() + 1);
          if(dA[i][4] === "Trimestrielle") oldDate.setMonth(oldDate.getMonth() + 3);
          sA.getRange(i+1, 3).setValue(oldDate);
          logAction(p.idAuteur, p.roleAuteur, "Action", "Alertes Partenaires", idPart, "Alerte récurrente complétée, prochaine: " + oldDate.toLocaleDateString('fr-FR') + " : « " + msgAlerte + " »");
        }
        notifierMMOKEtCreerMission(ss, p, "Partenaire", idPart, "cloture", msgAlerte);
        return createJsonResponse({ success: true });
      }
    }
    return createJsonResponse({ success: false, message: "Alerte non trouvée" });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// 3. Modifier l'alerte
function handleUpdateAlert(p, ss) {
  try {
    const sA = getSheetSafe(ss, "ALERTES_PARTENAIRES");
    const dA = sA.getDataRange().getValues();

    for(let i = 1; i < dA.length; i++) {
      if(String(dA[i][0]) === p.id) {
        const idPart = String(dA[i][1]);
        sA.getRange(i+1, 3).setValue(p.date);
        sA.getRange(i+1, 4).setValue(p.msg);
        sA.getRange(i+1, 5).setValue(p.freq);
        sA.getRange(i+1, 7).setValue(p.destinataire);
        const newDate = p.date ? new Date(p.date).toLocaleDateString('fr-FR') : "?";
        logAction(p.idAuteur, p.roleAuteur, "Modification", "Alertes Partenaires", idPart, "Alerte modifiée → Date: " + newDate + " | « " + String(p.msg || "").substring(0, 60) + " »");
        return createJsonResponse({ success: true, message: "Alerte mise à jour avec succès !" });
      }
    }
    return createJsonResponse({ success: false });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// 4. Annuler l'alerte
function handleCancelAlert(p, ss) {
  try {
    const sA = getSheetSafe(ss, "ALERTES_PARTENAIRES");
    const dA = sA.getDataRange().getValues();

    for(let i = 1; i < dA.length; i++) {
      if(String(dA[i][0]) === p.id) {
        const idPart = String(dA[i][1]);
        const msgAlerte = String(dA[i][3] || "").substring(0, 60);
        sA.getRange(i+1, 6).setValue("Annulé");
        logAction(p.idAuteur, p.roleAuteur, "Action", "Alertes Partenaires", idPart, "Alerte annulée : « " + msgAlerte + " »");
        return createJsonResponse({ success: true, message: "Alerte annulée" });
      }
    }
    return createJsonResponse({ success: false });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// 5. Tester l'alerte (Envoi du Mail)
function handleTestAlert(p, ss) {
  try {
    const sA = getSheetSafe(ss, "ALERTES_PARTENAIRES");
    const sP = getSheetSafe(ss, "PARTENARIAT");
    const dA = sA.getDataRange().getValues();
    
    let alertRow = dA.find(r => String(r[0]) === String(p.id));
    if(!alertRow) throw new Error("Alerte introuvable");
    
    const dP = sP.getDataRange().getValues();
    const pRow = dP.find(r => String(r[1]).trim().toUpperCase() === String(alertRow[1]).trim().toUpperCase());
    let partName = pRow ? String(pRow[2]).trim() : ("ID : " + alertRow[1]);

    const mailDest = alertRow[6] ? String(alertRow[6]).trim() : "m.mokhtari@hecg.fr";
    const subject = "[TEST ALERTE CRM] À contacter : " + partName;
    const body = `<div style="font-family: Arial, sans-serif; padding: 20px; border: 1px solid #eee; border-radius: 10px;">
                    <h2 style="color: #1A4E8A;"> Alerte HECG</h2>
                    <p><b> Entreprise concernée :</b> ${partName}</p>
                    <p><b> Date prévue :</b> ${new Date(alertRow[2]).toLocaleDateString('fr-FR')}</p>
                    <p><b> Fréquence :</b> ${alertRow[4]}</p>
                    <hr style="border: none; border-top: 1px solid #ddd; margin: 20px 0;">
                    <p><b> Votre note / message :</b></p>
                    <div style="background: #f9f9f9; padding: 15px; border-left: 4px solid #F26522; border-radius: 5px;">
                      ${String(alertRow[3]).replace(/\n/g, '<br>')}
                    </div>
                  </div>`;
    
    GmailApp.sendEmail(mailDest, subject, "", { htmlBody: body });
    return createJsonResponse({ success: true, message: "Un email de test a été envoyé à " + mailDest });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// ==========================================
// ENVOI AUTOMATIQUE DES ALERTES PARTENAIRES (trigger quotidien)
// ==========================================
function envoyerAlertesPartenaires() {
  try {
    const ss = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");
    const sA = getSheetSafe(ss, "ALERTES_PARTENAIRES");
    const sP = getSheetSafe(ss, "PARTENARIAT");
    const dA = sA.getDataRange().getValues();
    const dP = sP.getDataRange().getValues();

    const today = new Date(); today.setHours(0, 0, 0, 0);

    const partMap = {};
    for (let j = 1; j < dP.length; j++) {
      const cleanId = String(dP[j][1] || "").trim().toUpperCase();
      if (cleanId) partMap[cleanId] = String(dP[j][2] || "Entreprise").trim();
    }

    let nbEnvoyes = 0, nbIgnores = 0, erreurs = "";

    for (let i = 1; i < dA.length; i++) {
      if (!dA[i][0]) continue;
      if (dA[i][5] === "Terminé" || dA[i][5] === "Annulé") { nbIgnores++; continue; }

      const dateAlerte = new Date(dA[i][2]); dateAlerte.setHours(0, 0, 0, 0);
      if (dateAlerte > today) { nbIgnores++; continue; }

      const mailDest = dA[i][6] ? String(dA[i][6]).trim() : "m.mokhtari@hecg.fr";
      if (!mailDest.includes("@")) { nbIgnores++; continue; }

      const partName = partMap[String(dA[i][1]).trim().toUpperCase()] || ("ID : " + dA[i][1]);
      const isRetard = dateAlerte < today;
      const subject = (isRetard ? "[⚠️ RETARD] " : "[🔔 ALERTE] ") + "À contacter : " + partName;
      const body = `<div style="font-family:Arial,sans-serif;padding:20px;border:1px solid #eee;border-radius:10px;">
        <h2 style="color:#1A4E8A;">🔔 Alerte CRM HECG</h2>
        ${isRetard ? '<p style="color:#dc2626;font-weight:bold;">⚠️ Cette alerte était prévue le ' + new Date(dA[i][2]).toLocaleDateString('fr-FR') + ' et n\'a pas encore été traitée.</p>' : ''}
        <p><b>Entreprise :</b> ${partName}</p>
        <p><b>Date prévue :</b> ${new Date(dA[i][2]).toLocaleDateString('fr-FR')} &nbsp;|&nbsp; <b>Fréquence :</b> ${dA[i][4]}</p>
        <hr style="border:none;border-top:1px solid #ddd;margin:15px 0;">
        <p><b>Note / message :</b></p>
        <div style="background:#f9f9f9;padding:15px;border-left:4px solid #F26522;border-radius:5px;">${String(dA[i][3]).replace(/\n/g,'<br>')}</div>
      </div>`;

      try {
        GmailApp.sendEmail(mailDest, subject, "", { htmlBody: body });
        nbEnvoyes++;
        try { logMailDetail(ss, { auteur: "Système (CRON)", categorie: "Alerte", sousCat: "Alerte partenaire — " + (dA[i][4] || "Unique"), objet: subject, destinataires: [mailDest], corps: body, sourceId: String(dA[i][0] || "") }); } catch(eLog) {}
        if (dA[i][4] === "Unique") {
          sA.getRange(i + 1, 6).setValue("Terminé");
        } else {
          const newDate = new Date(dA[i][2]);
          if (dA[i][4] === "Hebdomadaire") newDate.setDate(newDate.getDate() + 7);
          else if (dA[i][4] === "Mensuelle") newDate.setMonth(newDate.getMonth() + 1);
          else if (dA[i][4] === "Trimestrielle") newDate.setMonth(newDate.getMonth() + 3);
          sA.getRange(i + 1, 3).setValue(newDate);
        }
      } catch (eEnvoi) {
        erreurs += mailDest + ": " + eEnvoi.toString() + "; ";
      }
    }
    logExecutionAuto(ss, "envoyerAlertesPartenaires", nbEnvoyes, nbIgnores, erreurs ? 1 : 0, erreurs);
  } catch (e) {
    try {
      const ss2 = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");
      logExecutionAuto(ss2, "envoyerAlertesPartenaires", 0, 0, 1, "CRASH : " + e.toString());
    } catch(e2) {}
  }
}

/**
 * À exécuter UNE SEULE FOIS depuis l'éditeur Apps Script pour installer le trigger quotidien à 8h.
 */
function createAlertsTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "envoyerAlertesPartenaires") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("envoyerAlertesPartenaires").timeBased().everyDays(1).atHour(8).create();
}

// ==========================================
// CRÉATION D'UNE ALERTE DEPUIS LE DOSSIER PARTENAIRE
// ==========================================
/**
 * Notifie MMOK par email et crée une mission pour l'auteur
 * quand un collaborateur (≠ m.mokhtari@hecg.fr) crée ou clôture une alerte.
 * @param {string} action  "creation" | "cloture"
 */
function notifierMMOKEtCreerMission(ss, p, typeAlerte, idCible, action, msgAlerte) {
  let emailAuteur = String(p.emailAuteur || "").trim().toLowerCase();

  // Fallback : résolution de l'email via p.ref (Personnel col A → col G)
  if (!emailAuteur && p.ref) {
    try {
      const dPerso = getSheetSafe(ss, "Personnel").getDataRange().getValues();
      for (let k = 1; k < dPerso.length; k++) {
        if (String(dPerso[k][0] || "").trim() === String(p.ref).trim()) {
          emailAuteur = String(dPerso[k][6] || "").trim().toLowerCase();
          break;
        }
      }
    } catch(eR) {}
  }

  if (!emailAuteur || emailAuteur === "m.mokhtari@hecg.fr") return;

  const nomAuteur  = String(p.idAuteur || emailAuteur).trim();
  const actionLib  = action === "creation" ? "créé" : "clôturé";
  const now        = new Date();
  const dateStr    = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy");
  const heureStr   = Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm");

  // --- 1. Email de notification à MMOK ---
  const sujet = "[CRM HECG] " + nomAuteur + " a " + actionLib + " une alerte " + typeAlerte + " (" + idCible + ")";
  const html  = '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">'
    + '<div style="background:#1A4E8A;padding:20px;border-radius:12px 12px 0 0;">'
    + '<h2 style="color:#fff;margin:0;font-size:18px;">🔔 Notification CRM HECG</h2></div>'
    + '<div style="background:#f8fafc;padding:20px;border-radius:0 0 12px 12px;border:1px solid #e2e8f0;">'
    + '<p style="font-size:15px;color:#1e293b;"><strong>' + nomAuteur + '</strong> a <strong>' + actionLib
    + '</strong> une alerte <strong>' + typeAlerte + ' ' + idCible + '</strong>.</p>'
    + '<div style="background:#fff;border-left:4px solid #F26522;padding:12px 16px;margin:16px 0;border-radius:4px;">'
    + '<p style="margin:0;color:#475569;font-size:13px;white-space:pre-wrap;">' + msgAlerte + '</p></div>'
    + '<p style="font-size:12px;color:#94a3b8;">Le ' + dateStr + ' à ' + heureStr + '</p></div></div>';
  try {
    GmailApp.sendEmail("m.mokhtari@hecg.fr", sujet, msgAlerte, { htmlBody: html, name: "CRM HECG" });
  } catch(eM) { console.error("notifierMMOK mail error: " + eM.message); }

  // --- 2. Mission pour l'auteur ---
  let sheet;
  try { sheet = getSheetSafe(ss, "MISSIONS"); }
  catch(e) {
    sheet = ss.insertSheet("MISSIONS");
    sheet.appendRow(["ID","Date","Heure","Email_Dest","Nom_Collab","CC","Objet","Corps","Statut","Notes","Date_Maj"]);
    sheet.setFrozenRows(1);
    sheet.getRange("1:1").setFontWeight("bold");
  }
  let nomCollab = nomAuteur;
  try {
    const dP2 = getSheetSafe(ss, "Personnel").getDataRange().getValues();
    for (let k = 1; k < dP2.length; k++) {
      if (String(dP2[k][6] || "").trim().toLowerCase() === emailAuteur) {
        nomCollab = (String(dP2[k][2] || "") + " " + String(dP2[k][1] || "")).trim();
        break;
      }
    }
  } catch(eP) {}
  const idMsn  = "MSN-" + now.getTime() + "-" + Math.floor(Math.random() * 1000);
  const objet  = "[" + (action === "creation" ? "Alerte créée" : "Alerte clôturée") + "] " + typeAlerte + " " + idCible + " – " + msgAlerte.substring(0, 50);
  sheet.appendRow([idMsn, dateStr, heureStr, emailAuteur, nomCollab, "", objet, msgAlerte, "Attribué", "", now, "", "", String(p.emailAuteur || emailAuteur)]);
  logAction(p.idAuteur || "Système", p.roleAuteur || "Auto", "Création", "Missions", idCible,
    "Mission auto créée (" + action + " alerte " + typeAlerte + ") pour " + nomAuteur);
}

function handleSetPartnerAlert(p, ss) {
  try {
    const sheet = getSheetSafe(ss, "ALERTES_PARTENAIRES");
    const newId = "AL" + Math.floor(1000 + Math.random() * 9000);
    
    // LA CORRECTION EST ICI : on récupère bien le "targetId" envoyé par le dossier
    const idPartenaire = String(p.targetId || p.partnerId || p.id_partenaire || p.id || "").trim().toUpperCase() || "ID_MANQUANT";
    
    const newRow = [
      newId,
      idPartenaire,                           // Col B : ID de l'entreprise
      p.date || p.dateAlerte || new Date(),   // Col C : Date
      p.message || p.msg || "Relance à effectuer", // Col D : Message
      p.frequence || p.freq || "Unique",      // Col E : Fréquence
      "À faire",                              // Col F : Statut
      p.destinataire || "m.mokhtari@hecg.fr"  // Col G : Email
    ];
    
    sheet.appendRow(newRow);
    const dateAlerte = (p.date || p.dateAlerte) ? new Date(p.date || p.dateAlerte).toLocaleDateString('fr-FR') : "?";
    const msgApercu = (p.message || p.msg || "Relance").substring(0, 60);
    logAction(p.idAuteur, p.roleAuteur, "Création", "Alertes Partenaires", idPartenaire, "Alerte programmée le " + dateAlerte + " : « " + msgApercu + " » (" + (p.frequence || p.freq || "Unique") + ")");
    const msnId = creerMissionDepuisAlerte(ss, p, "Partenaire", idPartenaire);
    notifierMMOKEtCreerMission(ss, p, "Partenaire", idPartenaire, "creation", p.message || p.msg || "Relance");
    const msgRetour = msnId ? "Alerte créée + Mission " + msnId + " déclenchée !" : "Alerte créée avec succès dans la base !";
    return createJsonResponse({ success: true, message: msgRetour, missionId: msnId || null });
  } catch(e) {
    return createJsonResponse({ success: false, message: "Erreur Serveur : " + e.toString() });
  }
}
// FONCTION POUR RÉCUPÉRER TOUS LES DÉTAILS D'UN PARTENAIRE (HISTORIQUE, ALTERNANTS, ETC.)
function handleGetPartnerDetails(params, ss) {
  try {
    const id = String(params.targetId || "").trim().toUpperCase();
    const sheetPart = getSheetSafe(ss, "PARTENARIAT");
    const dP = sheetPart.getDataRange().getValues();
    let info = null;
    const contacts = [];

    for (let i = 1; i < dP.length; i++) {
      if (String(dP[i][1]).trim().toUpperCase() === id) {
        if (!info) info = { entreprise: String(dP[i][2] || "").trim(), adresse: String(dP[i][10] || "").trim(), cp: String(dP[i][11] || "").trim(), ville: String(dP[i][12] || "").trim() };
        contacts.push({
          id_contact: String(dP[i][0] || "").trim(),
          nom: String(dP[i][3] || "").trim(),
          prenom: String(dP[i][4] || "").trim(),
          tel: String(dP[i][5] || "").trim(),
          mail: String(dP[i][6] || "").trim(),
          tel2: String(dP[i][9] || "").trim(),
          mail2: String(dP[i][8] || "").trim(),
          poste: String(dP[i][13] || "").trim(),
          is_ma: String(dP[i][14] || "Non").trim()
        });
      }
    }

    if (!info) return createJsonResponse({ success: false, message: "Entreprise introuvable" });
    
    const dataEtu = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues();
    const alternants = dataEtu
      .filter(r => String(r[11]).trim().toUpperCase() === id && String(r[13]).toLowerCase().includes("alternance"))
      .map(r => ({ id: String(r[0]).trim(), nom: r[3] + " " + r[2], prenom: String(r[3]).trim(), nomFamille: String(r[2]).trim(), classe: r[14] }));
    
    // --- LA CORRECTION EST ICI : ON UTILISE SUIVI_PARTENARIAT ---
    const suivi = getSheetSafe(ss, "SUIVI_PARTENARIAT").getDataRange().getValues()
      .filter(r => String(r[0]).trim().toUpperCase() === id)
      .map(r => ({ 
        auteur: r[1], // Colonne B
        date: r[2],   // Colonne C
        texte: r[3]   // Colonne D (Compte_rendu)
      }))
      .reverse(); // Pour avoir les plus récents en haut
    
    let alertes = [];
    try { alertes = getSheetSafe(ss, "ALERTES_PARTENAIRES").getDataRange().getValues().filter(r => String(r[1]).trim().toUpperCase() === id).map(r => ({ date: r[2], msg: r[3], freq: r[4], statut: r[5] })); } catch(eA) {}
    let offres = [];
    try { offres = getSheetSafe(ss, "OFFRES").getDataRange().getValues().filter(r => String(r[4]).trim().toLowerCase() === info.entreprise.toLowerCase()).map(r => ({ poste: r[14], url: r[1], etat: r[9] })); } catch(eO) {}

    let anciensAlternants = [];
    try {
      const wsA = ss.getSheetByName("ANCIENS_ALTERNANTS");
      if (wsA && wsA.getLastRow() > 1) {
        // Construire un map idEtu → infos actuelles (pour enrichir avec classe/campus si manquants dans l'historique)
        const etuMap = {};
        try {
          const dEtu = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues();
          for (let ei = 1; ei < dEtu.length; ei++) {
            const eid = String(dEtu[ei][0]).trim();
            if (eid) etuMap[eid] = { classe: String(dEtu[ei][14]||""), campus: String(dEtu[ei][22]||""), mail: String(dEtu[ei][4]||""), tel: String(dEtu[ei][5]||"") };
          }
        } catch(eMap) {}
        anciensAlternants = wsA.getDataRange().getValues()
          .slice(1)
          .filter(r => String(r[1]).trim().toUpperCase() === id && String(r[6]) !== "Erreur")
          .map(r => {
            const idEtu = String(r[2]).trim();
            const info = etuMap[idEtu] || {};
            let dateStr = "";
            if (r[5]) { try { dateStr = Utilities.formatDate(new Date(r[5]), Session.getScriptTimeZone(), "dd/MM/yyyy"); } catch(e) { dateStr = String(r[5]); } }
            return {
              id:     String(r[0]),
              idEtu:  idEtu,
              nom:    r[3],
              prenom: r[4],
              date:   dateStr,
              status: String(r[6]),
              classe: String(r[8]||"") || info.classe || "",
              campus: String(r[9]||"") || info.campus || "",
              motif:  String(r[10]||""),
              mail:   info.mail || "",
              tel:    info.tel || ""
            };
          });
      }
    } catch(eAnc) {}

    return createJsonResponse({ success: true, data: { info, contacts, alternants, suivi, alertes, offres, anciensAlternants } });
  } catch (e) { 
    return createJsonResponse({ success: false, message: e.toString() }); 
  }
}
function handleUpdateEnterprise(p, ss) {
  try {
    const sheet = getSheetSafe(ss, "PARTENARIAT");
    const data = sheet.getDataRange().getValues();
    const idCible = String(p.id_entreprise).trim().toUpperCase();
    let found = false;
    const changes = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim().toUpperCase() === idCible) {
        if (!found) {
          // Capture des anciennes valeurs sur la première ligne trouvée
          if (p.nouveauNom !== undefined && p.nouveauNom && data[i][2] !== p.nouveauNom) changes.push("Nom: " + data[i][2] + " → " + p.nouveauNom);
          if (p.adresse !== undefined && data[i][10] !== p.adresse) changes.push("Adresse: " + (data[i][10] || "-") + " → " + (p.adresse || "-"));
          if (p.cp !== undefined && String(data[i][11]) !== String(p.cp)) changes.push("CP: " + (data[i][11] || "-") + " → " + (p.cp || "-"));
          if (p.ville !== undefined && data[i][12] !== p.ville) changes.push("Ville: " + (data[i][12] || "-") + " → " + (p.ville || "-"));
        }
        if (p.nouveauNom !== undefined && p.nouveauNom) sheet.getRange(i + 1, 3).setValue(p.nouveauNom);
        if (p.adresse !== undefined) sheet.getRange(i + 1, 11).setValue(p.adresse);
        if (p.cp !== undefined) sheet.getRange(i + 1, 12).setValue(p.cp);
        if (p.ville !== undefined) sheet.getRange(i + 1, 13).setValue(p.ville);
        found = true;
      }
    }
    if (found) {
      const logDetail = changes.length > 0 ? changes.join(" | ") : "Mise à jour fiche entreprise";
      logAction(p.idAuteur, p.roleAuteur, "Modification", "Entreprise", idCible, logDetail);
      return createJsonResponse({ success: true });
    }
    return createJsonResponse({ success: false, message: "ID non trouvé" });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleUpdateContact(p, ss) {
  try {
    const sheet = getSheetSafe(ss, "PARTENARIAT");
    const data = sheet.getDataRange().getValues();
    const idCible = String(p.id_contact).trim();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === idCible) {
        const oldRow = data[i];
        const changes = [];
        const nomComplet = (oldRow[4] || "") + " " + (oldRow[3] || "");
        if (p.nom !== undefined && oldRow[3] !== p.nom) changes.push("Nom: " + (oldRow[3] || "-") + " → " + p.nom);
        if (p.prenom !== undefined && oldRow[4] !== p.prenom) changes.push("Prénom: " + (oldRow[4] || "-") + " → " + p.prenom);
        if (p.tel !== undefined && String(oldRow[5]) !== String(p.tel)) changes.push("Tél: " + (oldRow[5] || "-") + " → " + p.tel);
        if (p.mail !== undefined && oldRow[6] !== p.mail) changes.push("Mail: " + (oldRow[6] || "-") + " → " + p.mail);
        if (p.mail2 !== undefined && oldRow[8] !== p.mail2) changes.push("Mail2: " + (oldRow[8] || "-") + " → " + p.mail2);
        if (p.tel2 !== undefined && String(oldRow[9]) !== String(p.tel2)) changes.push("Tél2: " + (oldRow[9] || "-") + " → " + p.tel2);

        if (p.nom !== undefined) sheet.getRange(i + 1, 4).setValue(p.nom);
        if (p.prenom !== undefined) sheet.getRange(i + 1, 5).setValue(p.prenom);
        if (p.tel !== undefined) sheet.getRange(i + 1, 6).setValue(p.tel);
        if (p.mail !== undefined) sheet.getRange(i + 1, 7).setValue(p.mail);
        if (p.mail2 !== undefined) sheet.getRange(i + 1, 9).setValue(p.mail2);
        if (p.tel2 !== undefined) sheet.getRange(i + 1, 10).setValue(p.tel2);

        const logDetail = "Contact " + nomComplet.trim() + (changes.length > 0 ? " | " + changes.join(" | ") : " — coordonnées mises à jour");
        logAction(p.idAuteur, p.roleAuteur, "Modification", "Contact", idCible, logDetail);
        return createJsonResponse({ success: true });
      }
    }
    return createJsonResponse({ success: false, message: "Contact introuvable" });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleDeleteContact(p, ss) {
  try {
    const sheet = getSheetSafe(ss, "PARTENARIAT");
    const data = sheet.getDataRange().getValues();
    const idCible = String(p.id_contact).trim();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]).trim() === idCible) {
        const nomContact = (data[i][4] || "") + " " + (data[i][3] || "");
        const nomEnt = data[i][2] || "?";
        sheet.deleteRow(i + 1);
        logAction(p.idAuteur, p.roleAuteur, "Suppression", "Contact", idCible, "A supprimé le contact " + nomContact.trim() + " (Entreprise : " + nomEnt + ")");
        notifierSuppression(p.idAuteur, p.roleAuteur, "Contact entreprise supprimé", nomContact.trim() + " — Entreprise : " + nomEnt + " — ID : " + idCible);
        return createJsonResponse({ success: true });
      }
    }
    return createJsonResponse({ success: false, message: "Contact introuvable" });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// 1. FONCTION POUR PARTAGER UNE OFFRE À UN ÉTUDIANT
function handlePartagerOffre(p, ss) {
  try {
    const sheetPartages = getSheetSafe(ss, "PARTAGES");
    const dataPartages = sheetPartages.getDataRange().getValues();
    const sheetOffres = getSheetSafe(ss, "OFFRES");
    const dataOffres = sheetOffres.getDataRange().getValues();
    const dataEtu = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues();

    const offreRow = dataOffres.find(r => String(r[0]).trim() === String(p.id_offre).trim());
    if (!offreRow) return createJsonResponse({ success: false, message: "Offre introuvable." });

    const infoMail = { poste: offreRow[14] || "Offre", ent: offreRow[4] || "Entreprise", url: offreRow[1] };
    let cibles = [];
    const now = new Date();

    if (p.emailsCibles && Array.isArray(p.emailsCibles) && p.emailsCibles.length > 0) {
      // Mode liste explicite : l'utilisateur a édité les destinataires manuellement
      p.emailsCibles.forEach(rawMail => {
        const mail = String(rawMail || "").trim().toLowerCase();
        if (!mail.includes('@')) return;
        let prenom = mail.split('@')[0];
        let id = 'EXT-' + mail;
        for (let i = 1; i < dataEtu.length; i++) {
          if (String(dataEtu[i][4]).trim().toLowerCase() === mail) {
            prenom = String(dataEtu[i][3] || prenom);
            id = dataEtu[i][0];
            break;
          }
        }
        if (!cibles.some(c => c.mail === mail)) cibles.push({ id, mail, prenom });
      });
    } else {
      // Mode filtrage par classe (comportement existant)
      for (let i = 1; i < dataEtu.length; i++) {
        let id = dataEtu[i][0];
        let prenom = dataEtu[i][3];
        let mail = String(dataEtu[i][4]).trim();
        let statut = String(dataEtu[i][13]).toLowerCase();
        let classe = String(dataEtu[i][14]).trim();
        let typo = String(dataEtu[i][20]).toLowerCase();
        let estEnRecherche = statut.includes("recherche");
        let estLegitime = typo.includes("inscrit");
        let estExclu = typo.includes("sorti") || typo.includes("refus") || statut.includes("alternance") || statut.includes("initial") || statut.includes("placé") || statut.includes("contrat") || statut.includes("non su");
        if (estEnRecherche && estLegitime && !estExclu && mail.includes('@')) {
          let matchClasse = true;
          if (p.classesCibles && p.classesCibles.length > 0 && !p.classesCibles.includes("TOUS")) {
            matchClasse = p.classesCibles.includes(classe);
          }
          if (matchClasse) cibles.push({ id, mail, prenom });
        }
      }
    }

    if (cibles.length === 0) return createJsonResponse({ success: false, message: "Aucun étudiant ne correspond aux critères (Inscrit/Préinscrit en recherche de cette classe)." });

    let nbAjoutes = 0;
    var mailsEnvoyes = [];
    var sujetOffre = "Nouvelle offre ciblée : " + infoMail.poste;
    var corpsOffre = "";
    cibles.forEach(etu => {
      const dejaFait = dataPartages.some(r => String(r[0]).trim() === String(etu.id).trim() && String(r[1]).trim() === String(p.id_offre).trim());
      if (!dejaFait) {
        sheetPartages.appendRow([etu.id, p.id_offre, now, "Partagée"]);
        try {
          const corps = `<div style="font-family:Arial; padding:20px; border:1px solid #eee; border-radius:15px;"><h2 style="color:#1A4E8A;">Bonjour ${etu.prenom},</h2><p>Une nouvelle offre a été sélectionnée pour ton profil :</p><div style="background:#f9f9f9; padding:15px; border-left:5px solid #F26522; margin:20px 0;"><b style="font-size:16px;">${infoMail.poste}</b><br><span>Entreprise : ${infoMail.ent}</span></div><p>Tu peux postuler directement depuis ton dossier étudiant HECG.</p><a href="${infoMail.url}" style="display:inline-block; padding:12px 25px; background:#F26522; color:white; text-decoration:none; border-radius:8px; font-weight:bold;">Voir l'annonce</a></div>`;
          corpsOffre = corps;
          GmailApp.sendEmail(etu.mail, sujetOffre, "", { htmlBody: corps + getSignatureHTML() });
          mailsEnvoyes.push(etu.mail);
        } catch(e) {}
        nbAjoutes++;
      }
    });

    // Mise à jour du statut : toujours écrire "Partagée", quel que soit nbAjoutes
    for (let i = 1; i < dataOffres.length; i++) {
      if (String(dataOffres[i][0]).trim() === String(p.id_offre).trim()) {
        sheetOffres.getRange(i + 1, 14).setValue("Partagée"); // Colonne N
        break;
      }
    }
    SpreadsheetApp.flush(); // Forcer la validation de l'écriture avant le retour
    logAction(p.idAuteur, p.roleAuteur, "Action", "Offres", p.id_offre, "A partagé l'offre avec " + nbAjoutes + " étudiant(s)");
    if (mailsEnvoyes.length > 0) {
      logMailDetail(ss, {
        auteur: p.idAuteur || "Système",
        categorie: "Offre partagée",
        sousCat: infoMail.poste + " — " + infoMail.ent,
        objet: sujetOffre,
        destinataires: mailsEnvoyes,
        corps: corpsOffre,
        sourceId: p.id_offre
      });
    }
    return createJsonResponse({ success: true, message: nbAjoutes + " partage(s) effectué(s). L'offre est passée en 'Partagée'." });
  } catch(e) { return createJsonResponse({ success: false, message: "Erreur serveur : " + e.toString() }); }
}

// 2. FONCTION POUR ENREGISTRER QUAND UN ÉTUDIANT CLIQUE SUR "POSTULER"
function handlePostulerOffre(p, ss) {
  try {
    const sheet = getSheetSafe(ss, "PARTAGES");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(p.id_etu) && String(data[i][1]) === String(p.id_offre)) {
        sheet.getRange(i + 1, 4).setValue("Postulée");
        logAction(p.idAuteur, p.roleAuteur, "Action", "Candidature", p.id_etu, "L'étudiant a cliqué sur POSTULER pour l'offre " + p.id_offre);
        return createJsonResponse({ success: true });
      }
    }
    return createJsonResponse({ success: false });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}
function nettoyerOffresPerimees() {
  const ss = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");
  const sheet = getSheetSafe(ss, "OFFRES");
  const data = sheet.getDataRange().getValues();
  const today = new Date();
  
  for (let i = 1; i < data.length; i++) {
    const dateOffre = new Date(data[i][2]);
    const diffJours = (today - dateOffre) / (1000 * 60 * 60 * 24);
    
    // Si l'offre a plus de 14 jours et n'est pas déjà obsolète ou pourvue
    if (diffJours > 14 && data[i][13] !== "Obsolète") {
      sheet.getRange(i + 1, 14).setValue("Obsolète"); // Colonne N
    }
  }
}
function handleDeleteOffre(p, ss) {
  try {
    const sheet = getSheetSafe(ss, "OFFRES");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(p.id)) {
        const oldRow = data[i];
        const infoLog = (oldRow[14] || "?") + " chez " + (oldRow[4] || "?") + " (" + (oldRow[9] || "?") + ")";
        sheet.deleteRow(i + 1);
        const sheetP = getSheetSafe(ss, "PARTAGES");
        const dataP = sheetP.getDataRange().getValues();
        for (let j = dataP.length - 1; j >= 1; j--) {
          if (String(dataP[j][1]) === String(p.id)) sheetP.deleteRow(j + 1);
        }
        logAction(p.idAuteur, p.roleAuteur, "Suppression", "Offres", p.id, "A supprimé l'offre : " + infoLog);
        return createJsonResponse({ success: true });
      }
    }
    return createJsonResponse({ success: false });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// --- JOURNAL TRACÉ DES OFFRES ---
function _getSuiviOffresSheet(ss) {
  let ws = ss.getSheetByName("SUIVI_OFFRES");
  if (!ws) {
    ws = ss.insertSheet("SUIVI_OFFRES");
    ws.appendRow(["ID", "ID_Offre", "Auteur", "Date", "Note"]);
    ws.setFrozenRows(1);
  }
  return ws;
}

function handleAddOffreNote(p, ss) {
  try {
    if (!p.id_offre || !p.texte) return createJsonResponse({ success: false, message: "Paramètres manquants." });
    const ws = _getSuiviOffresSheet(ss);
    const id = "ON-" + Date.now();
    ws.appendRow([id, String(p.id_offre).trim(), p.idAuteur || "Inconnu", new Date(), String(p.texte).trim()]);
    logAction(p.idAuteur, p.roleAuteur, "Note", "Offres", p.id_offre, "Note ajoutée : " + String(p.texte).substring(0, 80));
    return createJsonResponse({ success: true, id });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleGetOffreNotes(p, ss) {
  try {
    const ws = _getSuiviOffresSheet(ss);
    const data = ws.getDataRange().getValues();
    const idCible = String(p.id_offre || "").trim().toUpperCase();
    const notes = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim().toUpperCase() === idCible) {
        const d = data[i][3];
        notes.push({
          id: String(data[i][0]),
          auteur: String(data[i][2] || ""),
          date: d instanceof Date ? Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : String(d || ""),
          texte: String(data[i][4] || "")
        });
      }
    }
    notes.reverse();
    return createJsonResponse({ success: true, data: notes });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// --- CRÉATION D'UNE ENTREPRISE ET D'UN CONTACT ---
function handleAddPartner(p, ss) {
  try {
    // 1. On cible l'unique onglet qui gère les entreprises et les contacts
    const sheetPart = getSheetSafe(ss, "PARTENARIAT"); 

    let idEnt = p.id_entreprise;

    // 2. Si c'est une NOUVELLE entreprise, on génère un ID Entreprise
    if (!idEnt || String(idEnt).trim() === "") {
      idEnt = "ENT" + Utilities.formatDate(new Date(), "GMT", "yyMMddHHmmss");
    }

    // 3. On crée le nouveau contact lié à cette entreprise
    const idCont = "CONT" + Utilities.formatDate(new Date(), "GMT", "yyMMddHHmmss") + Math.floor(Math.random() * 10);
    
    // 4. On insère tout sur une seule ligne dans PARTENARIAT
    // L'ordre respecte la lecture de votre CRM : ID_Cont | ID_Ent | Nom_Ent | Nom | Prénom | Tel1 | Mail1 | Vide | Mail2 | Tel2
    const nouvelleLigne = [
      idCont,
      idEnt,
      p.nom_entreprise || "",
      p.nom || "",
      p.prenom || "",
      p.tel1 || "",
      p.mail1 || "",
      "",               // Colonne H
      p.mail2 || "",    // Colonne I
      p.tel2 || "",     // Colonne J
      p.adresse || "",  // Colonne K
      p.cp || "",       // Colonne L
      p.ville || "",    // Colonne M
      p.poste || "",    // Colonne N — Poste du contact
      p.is_ma || "Non"  // Colonne O — Maître d'apprentissage
    ];
    
    sheetPart.appendRow(nouvelleLigne);

    // 5. On laisse une trace dans l'historique
    // 5. On laisse une trace dans l'historique (avec sécurité sur le nom)
    const auteurReconnu = p.idAuteur || p.auteur || "Inconnu";
    const roleReconnu = p.roleAuteur || p.role || "N/A";
    const isNewEnt = !p.id_entreprise || String(p.id_entreprise).trim() === "";
    const logDetail = (isNewEnt ? "Nouvelle entreprise : " : "Nouveau contact ajouté à : ") + (p.nom_entreprise || "?") + " | Contact : " + (p.prenom || "") + " " + (p.nom || "") + (p.mail1 ? " <" + p.mail1 + ">" : "") + (p.ville ? " — " + p.ville : "");
    logAction(auteurReconnu, roleReconnu, "Création", isNewEnt ? "Entreprise" : "Contact", idEnt, logDetail);
    
    return createJsonResponse({ success: true, message: "Interlocuteur (et entreprise) ajouté avec succès !", id_contact: idCont, id_entreprise: idEnt });

  } catch (e) {
    return createJsonResponse({ success: false, message: "Erreur d'ajout : " + e.toString() });
  }
}

function handleDeleteEnterprise(p, ss) {
  try {
    const sheet = getSheetSafe(ss, "PARTENARIAT");
    const data = sheet.getDataRange().getValues();
    const idCible = String(p.id_entreprise).trim().toUpperCase();
    let nomEnt = "";
    let nbRows = 0;
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][1]).trim().toUpperCase() === idCible) {
        if (!nomEnt) nomEnt = data[i][2];
        sheet.deleteRow(i + 1);
        nbRows++;
      }
    }
    if (nbRows === 0) return createJsonResponse({ success: false, message: "Entreprise introuvable" });
    logAction(p.idAuteur, p.roleAuteur, "Suppression", "Entreprise", idCible, "A supprimé l'entreprise «" + nomEnt + "» et ses " + nbRows + " contact(s)");
    notifierSuppression(p.idAuteur, p.roleAuteur, "Entreprise supprimée", nomEnt + " (ID : " + idCible + ") — " + nbRows + " contact(s) supprimé(s)");
    return createJsonResponse({ success: true });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleLinkMAToStudents(p, ss) {
  try {
    const sheet = getSheetSafe(ss, "ETUDIANTS");
    const data = sheet.getDataRange().getValues();
    const ids = Array.isArray(p.ids_etudiants) ? p.ids_etudiants : [];
    const idEnt = String(p.id_entreprise || "").trim();
    const idCont = String(p.id_contact || "").trim();
    const nomEnt = String(p.nom_entreprise || "").trim();
    let updated = 0;
    for (let i = 1; i < data.length; i++) {
      if (ids.includes(String(data[i][0]).trim())) {
        sheet.getRange(i + 1, 12).setValue(idEnt);   // col L — id_ent
        sheet.getRange(i + 1, 13).setValue(idCont);  // col M — id_cont
        sheet.getRange(i + 1, 22).setValue(nomEnt);  // col V — nom entreprise
        updated++;
      }
    }
    logAction(p.idAuteur, p.roleAuteur, "Liaison MA", "Contact", idCont, "MA rattaché à " + updated + " étudiant(s) de " + nomEnt);
    return createJsonResponse({ success: true, updated });
  } catch (e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// --- PETIT BLOC A : CALCUL DATE EXPIRATION ---
function calculerDateExpiration(jours) {
  var date = new Date();
  date.setDate(date.getDate() + jours);
  return date;
}

function aspirerLinkedIn() {
  const ss = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");
  const sheet = getSheetSafe(ss, "OFFRES");
  const label = GmailApp.getUserLabelByName("Alertes/LINKEDIN_ALERTES");
  if (!label) return;
  const threads = label.getThreads();

  // --- MÉMOIRE GLOBALE ---
  const data = sheet.getDataRange().getValues();
  const urlsExistantes = data.map(r => nettoyerURL(String(r[1])));
  const postesEntreprises = data.map(r => (String(r[14]) + String(r[4])).toLowerCase().trim());

  threads.forEach(thread => {
    const htmlBody = thread.getMessages()[0].getBody();
    const regex = /<a href="(https:\/\/www\.linkedin\.com\/comm\/jobs\/view\/[0-9?&_=-]+)"[^>]*>(.*?)<\/a>.*?<span>([^<]+)<\/span>/g;
    let match;

    while ((match = regex.exec(htmlBody)) !== null) {
      let urlOffre = nettoyerURL(match[1]);
      let poste = match[2].replace(/<[^>]*>/g, '').trim();
      let entreprise = match[3].trim();
      let identifiantUnique = (poste + entreprise).toLowerCase().trim();

      // VÉRIFICATION CROISÉE
      if (!urlsExistantes.includes(urlOffre) && !postesEntreprises.includes(identifiantUnique)) {
        sheet.appendRow(["LK-"+Date.now(), urlOffre, new Date(), "LinkedIn", entreprise, "", "", "", "", "À qualifier", "", "", "", "", poste, "Alternance", "Détail sur lien", calculerDateExpiration(15)]);
        urlsExistantes.push(urlOffre);
        postesEntreprises.push(identifiantUnique);
      }
    }
    thread.removeLabel(label).markRead();
  });
}

function aspirerGoogleAlerts() {
  const ss = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");
  const sheet = getSheetSafe(ss, "OFFRES");
  
  // 1. RECHERCHE
  const threads = GmailApp.search('from:(notify-noreply@google.com OR googlealerts-noreply@google.com)', 0, 20);
  console.log("Conversations examinées : " + threads.length);

  const data = sheet.getDataRange().getValues();
  // L'Anti-doublon : Vu qu'on fabrique l'URL, on vérifie si (Poste + Entreprise) existe déjà !
  const offresExistantes = data.map(r => (String(r[14]) + String(r[4])).toLowerCase().trim().replace(/\s+/g, ''));

  let nbNouveautes = 0;

  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(msg => {
      // On récupère le texte pur, exactement comme dans notre diagnostic !
      const body = msg.getPlainBody();
      const lignes = body.split('\n').map(l => l.trim()).filter(l => l.length > 0);
      
      // On parcourt les lignes une par une
      for (let i = 0; i < lignes.length; i++) {
        
        // NOTRE POINT DE REPÈRE : Le mot "via "
        if (lignes[i].startsWith("via ") && i >= 3) {
          const plateforme = lignes[i].replace("via ", "").trim();
          const lieu = lignes[i-1];
          const entreprise = lignes[i-2];
          const poste = lignes[i-3];
          
          // Sécurité : On s'assure que ce sont bien des titres et pas des phrases longues
          if (entreprise.length > 60 || poste.length > 120) continue;

          // Création de la clé anti-doublon (Ex: "alternant(e)gestionnairepaiehexanet")
          const cleDoublon = (poste + entreprise).toLowerCase().replace(/\s+/g, '');

          if (!offresExistantes.includes(cleDoublon)) {
            
            // LA MAGIE : On fabrique un lien Google Jobs intelligent qui pointe sur cette offre
            const motsCles = encodeURIComponent(poste + " " + entreprise + " " + lieu);
            const lienIntelligent = "https://www.google.com/search?q=" + motsCles + "&ibp=htl;jobs";

            const idOffre = "GO-" + Date.now() + "-" + Math.floor(Math.random() * 100);
            
            // On range tout proprement dans les bonnes colonnes
            sheet.appendRow([
              idOffre, 
              lienIntelligent, // La colonne URL
              new Date(), 
              plateforme, 
              entreprise, 
              lieu, "", "", "", "À qualifier", "", "", "", "Non partagée", 
              poste, // Colonne O
              "Alternance", 
              "Alerte Google"
            ]);
            
            offresExistantes.push(cleDoublon);
            nbNouveautes++;
          }
        }
      }
    });
  });

  console.log("--- TERMINÉ : " + nbNouveautes + " nouvelles offres ajoutées ---");
}

const FT_CLIENT_ID = "PAR_crmhecg_9ac657d456ba54feb599d45ab15c33a98ddc114b41990d3c87bd04814fb55623";
const FT_CLIENT_SECRET = "fa072af71be48f94e3040ff6abca9724dbe6e5363652c4f35560029ba7ade169";

function getFranceTravailToken() {
  const url = "https://entreprise.francetravail.fr/connexion/oauth2/access_token?realm=%2Fpartenaire";
  const payload = { 
    "grant_type": "client_credentials", 
    "client_id": FT_CLIENT_ID, 
    "client_secret": FT_CLIENT_SECRET, 
    "scope": "api_offresdemploiv2 o2dsoffre" // Scope corrigé
  };
  
  const options = { 
    "method": "post", 
    "contentType": "application/x-www-form-urlencoded", 
    "payload": Object.keys(payload).map(k => k + '=' + encodeURIComponent(payload[k])).join('&'),
    "muteHttpExceptions": true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());
  
  if (json.error) {
    throw new Error("Erreur France Travail : " + json.error_description);
  }
  
  return json.access_token;
}

// --- 3. LA FONCTION PRINCIPALE D'ASPIRATION ---
function aspirerFranceTravail() {
  const ss = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");
  const sheet = getSheetSafe(ss, "OFFRES");
  const token = getFranceTravailToken();

  // Liste des recherches (Marseille 13055 et Aix 13001 + 20km)
  const recherches = [
    "https://api.francetravail.io/partenaire/offresdemploi/v2/offres/search?motsCles=comptabilite+alternance&commune=13055&distance=20",
    "https://api.francetravail.io/partenaire/offresdemploi/v2/offres/search?motsCles=paie+alternance&commune=13055&distance=20",
    "https://api.francetravail.io/partenaire/offresdemploi/v2/offres/search?motsCles=gestion+administrative+alternance&commune=13055&distance=20",
    "https://api.francetravail.io/partenaire/offresdemploi/v2/offres/search?motsCles=comptabilite+alternance&commune=13001&distance=20",
    "https://api.francetravail.io/partenaire/offresdemploi/v2/offres/search?motsCles=paie+alternance&commune=13001&distance=20",
    "https://api.francetravail.io/partenaire/offresdemploi/v2/offres/search?motsCles=gestion+administrative+alternance&commune=13001&distance=20"
  ];

  // On récupère les URLs déjà présentes pour éviter les doublons
  const dataExistante = sheet.getDataRange().getValues();
  const urlsExistantes = dataExistante.map(r => String(r[1])); 

  let totalAjoutees = 0;
  const now = new Date();
  const dateExp = new Date();
  dateExp.setDate(now.getDate() + 15);

  recherches.forEach(url => {
    const res = UrlFetchApp.fetch(url, { headers: { "Authorization": "Bearer " + token }, "muteHttpExceptions": true });
    
    if (res.getResponseCode() === 200) {
      const json = JSON.parse(res.getContentText());
      const resultats = json.resultats;

      if (resultats && resultats.length > 0) {
        Logger.log("Trouvé " + resultats.length + " offres pour : " + url);
        
        resultats.forEach(job => {
          // Vérification doublon par URL
          if (!urlsExistantes.includes(job.origineOffre.urlOrigine)) {
            
            sheet.appendRow([
              "FT-" + job.id,           // A : ID
              job.origineOffre.urlOrigine, // B : URL
              now,                      // C : Date Parution
              "France Travail",         // D : Plateforme
              job.entreprise ? job.entreprise.nom : "Anonyme", // E : Entreprise
              "", "", "", "",           // F à I : Vides
              "À qualifier",            // J : Statut
              "", "", "", "",           // K à N : Vides
              job.intitule,             // O : Poste
              job.typeContratLibelle,   // P : Type (CDD, Apprentissage...)
              job.description.slice(0, 500) + "...", // Q : Missions
              dateExp                   // R : Date Expiration J+15
            ]);
            
            urlsExistantes.push(job.origineOffre.urlOrigine);
            totalAjoutees++;
          }
        });
      } else {
        Logger.log("Aucune nouvelle offre sur cette URL.");
      }
    } else {
      Logger.log("Erreur sur l'URL : " + url + " (Code : " + res.getResponseCode() + ")");
    }
  });

  Logger.log("=== FIN : " + totalAjoutees + " offres ajoutées au total ===");
}

function aspirerLaBonneAlternance() {
  const ss = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");
  const sheet = getSheetSafe(ss, "OFFRES");

  // Endpoint v1 public (pas de clé API requise) — sources=matcha,peJob pour exclure les formations
  const apiUrl = "https://labonnealternance.apprentissage.beta.gouv.fr/api/v1/jobsEtFormations?romes=M1203,M1206,M1604,M1605&insee=13055&latitude=43.296482&longitude=5.36978&radius=30&sources=matcha,peJob&caller=crm_hecg";

  const response = UrlFetchApp.fetch(apiUrl, { muteHttpExceptions: true });

  if (response.getResponseCode() !== 200) {
    Logger.log("Erreur LBA (" + response.getResponseCode() + ") : " + response.getContentText().slice(0, 300));
    return;
  }

  const json = JSON.parse(response.getContentText());
  const jobsBlock = json.jobs || {};

  // Anti-doublon : on charge URL (col B) et clé Poste+Entreprise (col O+E) existants
  const dataExistante = sheet.getDataRange().getValues();
  const urlsExistantes = dataExistante.map(r => String(r[1]).toLowerCase().trim());
  const clesExistantes = dataExistante.map(r => (String(r[14]) + String(r[4])).toLowerCase().replace(/\s+/g, ''));

  let totalAjoutees = 0;

  // Les deux sources : matchas (offres LBA directes) + peJobs (France Travail via LBA)
  const tableaux = [
    (jobsBlock.matchas && jobsBlock.matchas.results) ? jobsBlock.matchas.results : [],
    (jobsBlock.peJobs  && jobsBlock.peJobs.results)  ? jobsBlock.peJobs.results  : []
  ];

  tableaux.forEach(tableau => {
    tableau.forEach(job => {
      const id         = job.id || "";
      const url        = (job.contact && job.contact.url) ? job.contact.url : "";
      const entreprise = (job.company && job.company.name) ? job.company.name : "";
      const lieu       = (job.place   && job.place.city)   ? job.place.city   : "";
      const poste      = job.title || "";

      if (!poste || !url) return;

      const urlNorm    = url.toLowerCase().trim();
      const cleDoublon = (poste + entreprise).toLowerCase().replace(/\s+/g, '');

      if (urlsExistantes.includes(urlNorm) || clesExistantes.includes(cleDoublon)) return;

      sheet.appendRow([
        "LBA-" + id,               // A  : ID
        url,                       // B  : URL
        new Date(),                // C  : Date d'aspiration
        "La Bonne Alternance",     // D  : Plateforme
        entreprise,                // E  : Entreprise
        lieu,                      // F  : Lieu
        "", "", "",                // G, H, I
        "À qualifier",             // J  : Statut
        "", "", "",                // K, L, M
        "Non partagée",            // N  : Partage
        poste,                     // O  : Poste
        "Alternance",              // P  : Type de contrat
        "",                        // Q  : Description (vide)
        calculerDateExpiration(15) // R  : Date expiration J+15
      ]);

      urlsExistantes.push(urlNorm);
      clesExistantes.push(cleDoublon);
      totalAjoutees++;
    });
  });

  Logger.log("=== LBA : " + totalAjoutees + " offres ajoutées au total ===");
}

function toutAspirer() {
  console.log("🚀 Lancement de l'aspiration globale...");

  try {
    aspirerFranceTravail();
  } catch(e) { console.error("Erreur France Travail : " + e.message); }

  try {
    aspirerLaBonneAlternance();
  } catch(e) { console.error("Erreur La Bonne Alternance : " + e.message); }
  
  try {
    aspirerLinkedIn();
  } catch(e) { console.error("Erreur LinkedIn : " + e.message); }

  try {
    AspirerOffresGmail();
  } catch(e) { console.error("Erreur MMOK GOOGLE ALERTES : " + e.message); }
  
  try {
    aspirerGoogleAlerts();
  } catch(e) { console.error("Erreur Google Alerts : " + e.message); }

  try {
    aspirerHelloWork();
  } catch(e) { console.error("Erreur HelloWork : " + e.message); }

  try {
    aspirerMissionsGmail();
  } catch(e) { console.error("Erreur Missions Gmail : " + e.message); }

  console.log("🏁 Fin de l'aspiration globale.");
}

// --- OUTIL : NETTOYEUR D'URL ---
function nettoyerURL(url) {
  if (!url) return "";
  return url.split('?')[0].split('#')[0].trim();
}

function envoyerBilanOffres() {
  const ss = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");
  const sheet = ss.getSheetByName("OFFRES") || getSheetSafe(ss, "OFFRES");
  const data = sheet.getDataRange().getValues();
  
  let nbAQualifier = 0;
  let nbNonPartages = 0;
  
  for (let i = 1; i < data.length; i++) {
    let qualite = String(data[i][9]).toLowerCase().trim();  // Colonne J : Qualité (Vérifiée, À qualifier, Obsolète)
    let etat    = String(data[i][13]).toLowerCase().trim(); // Colonne N : Etat (Partagée, Non partagée)

    // On ignore les offres obsolètes
    if (!qualite.includes("obsolete")) {

      // Offres à qualifier
      if (qualite.includes("qualifier") || qualite === "") {
        nbAQualifier++;
      }

      // Offres non partagées
      if (etat.includes("non") || etat === "") {
        nbNonPartages++;
      }
    }
  }
  
  const message = "Bonjour l'équipe,\n\nVoici le point actuel sur votre Bourse aux Offres HECG :\n\n" +
                  "" + nbAQualifier + " offres sont en attente d'être qualifiées.\n" +
                  " " + nbNonPartages + " offres actives n'ont pas encore été partagées aux étudiants.\n\n" +
                  "Connectez-vous sur votre interface pour les traiter.\n\n" +
                  "Passez une excellente journée !";
                  
  GmailApp.sendEmail("polepedagogique@hecg.fr", " Point sur vos offres HECG", message);
}

// --- NETTOYEUR AUTOMATIQUE : OFFRES EXPIRÉES ---
function verifierExpirationOffres() {
  const ss = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");
  const sheet = ss.getSheetByName("OFFRES") || getSheetSafe(ss, "OFFRES");
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  
  let nbNettoyees = 0;

  // On boucle sur toutes les lignes (en ignorant la ligne 1 des titres)
  for (let i = 1; i < data.length; i++) {
    let qualiteActuelle = String(data[i][9]).toLowerCase().trim(); // Colonne J (Index 9)
    let dateExpiration = data[i][17]; // Colonne R (Index 17)

    // Si l'offre n'est pas déjà obsolète ou pourvue...
    if (!qualiteActuelle.includes("obsolète") && !qualiteActuelle.includes("obsolete") && !qualiteActuelle.includes("pourvue")) {
      
      // On vérifie que la case R contient bien une date valide
      if (dateExpiration instanceof Date && !isNaN(dateExpiration)) {
        
        // Si la date d'expiration est plus vieille que la date de maintenant
        if (dateExpiration < now) {
          // On écrit "Obsolète" dans la Colonne J (10ème colonne pour le setValue)
          sheet.getRange(i + 1, 10).setValue("Obsolète");
          nbNettoyees++;
        }
      }
    }
  }
  
  Logger.log("Nettoyage terminé : " + nbNettoyees + " offre(s) expirée(s).");
}

// ==========================================
// GESTION DE LA CAMPAGNE RELANCE HEBDO
// ==========================================

// 1. Lire le modèle (ou charger celui par défaut)
function handleGetTemplateRelance(p, ss) {
  let template = PropertiesService.getScriptProperties().getProperty("TEMPLATE_RELANCE");
  
  // Si c'est la première fois, on met un template par défaut (le vôtre)
  if (!template) {
    template = `<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8"></head><body style="background:#001a38; color:white; padding:20px; font-family:Arial;">
      <h2 style="color:#FF7A00;">Bonjour {{PRENOM}},</h2>
      <p>Tu as {{NB_OFFRES}} offre(s) en attente :</p>
      {{INSERER_VOS_OFFRES_ICI}}
      <p>Contacte ton référent : polepedagogique@hecg.fr</p>
    </body></html>`;
  }
  return createJsonResponse({ success: true, template: template });
}

// 2. Sauvegarder le modèle
function handleSaveTemplateRelance(p, ss) {
  PropertiesService.getScriptProperties().setProperty("TEMPLATE_RELANCE", p.template);
  return createJsonResponse({ success: true, message: "Modèle HTML sauvegardé avec succès !" });
}

// 3. Tester le modèle (Envoi d'un faux mail)
function handleTestTemplateRelance(p, ss) {
  let template = PropertiesService.getScriptProperties().getProperty("TEMPLATE_RELANCE") || "Erreur : Aucun modèle.";
  
  // On remplace par des fausses données pour le test
  let htmlTest = template
    .replace("{{PRENOM}}", "Étudiant Test")
    .replace("{{NB_OFFRES}}", "2")
    .replace("{{INSERER_VOS_OFFRES_ICI}}", `
      <div style="border-left: 4px solid #FF7A00; padding:10px; margin-bottom:10px;">
        <strong>Assistant Comptable (F/H)</strong><br>🏢 Boîte Test SARL
      </div>
    `);
    
  try {
    GmailApp.sendEmail(p.email, "TEST - Relance Alternance", "", { htmlBody: htmlTest });
    return createJsonResponse({ success: true, message: "Email de test envoyé à " + p.email });
  } catch(e) {
    return createJsonResponse({ success: false, message: "Erreur d'envoi : " + e.toString() });
  }
}

// --- RELANCE HEBDOMADAIRE DES ÉTUDIANTS ---
function envoyerRelanceHebdomadaire() {
  const ss = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");
  const sheetEtu = getSheetSafe(ss, "ETUDIANTS");
  const sheetOffres = getSheetSafe(ss, "OFFRES");
  const sheetPartages = getSheetSafe(ss, "PARTAGES");

  const dataEtu = sheetEtu.getDataRange().getValues();
  const dataOffres = sheetOffres.getDataRange().getValues();
  const dataPartages = sheetPartages.getDataRange().getValues();

  // 1. Mémoriser toutes les offres pour les retrouver vite
  const offresMap = {};
  for (let i = 1; i < dataOffres.length; i++) {
    offresMap[String(dataOffres[i][0]).trim()] = {
      url: dataOffres[i][1],
      entreprise: dataOffres[i][4],
      qualite: String(dataOffres[i][9]).toLowerCase().trim(), // Colonne J
      poste: dataOffres[i][14],
      notesInternes: dataOffres[i][15] // <-- AJOUT : Récupère la Colonne P (Notes)
    };
  }

  let nbMailsEnvoyes = 0;
  let nbIgnores = 0;
  let nbErreurs = 0;
  const erreurDetails = [];
  var mailsRelances = [];

  // 2. Boucler sur les étudiants
  for (let i = 1; i < dataEtu.length; i++) {
    let idEtu = String(dataEtu[i][0]).trim();
    let prenom = dataEtu[i][3];
    let mailEtu = dataEtu[i][4];
    let statut = String(dataEtu[i][13] || "").toLowerCase();
    let typo = String(dataEtu[i][20] || "").toLowerCase();

    // Vérifier que l'étudiant est bien "En recherche" et non "Sorti/Refus"
    if (!statut.includes("recherche") || typo.includes("sorti") || typo.includes("refus") || !mailEtu.includes("@")) {
      nbIgnores++;
      continue;
    }

    // 3. Trouver les offres partagées mais non postulées
    let offresEnAttente = [];
    for (let p = 1; p < dataPartages.length; p++) {
      let partIdEtu = String(dataPartages[p][0]).trim();
      let partIdOffre = String(dataPartages[p][1]).trim();
      let partEtat = String(dataPartages[p][3] || "").toLowerCase();

      if (partIdEtu === idEtu && !partEtat.includes("postul") && !partEtat.includes("refus")) {
        let offreDetails = offresMap[partIdOffre];
        if (offreDetails && !offreDetails.qualite.includes("obsolète") && !offreDetails.qualite.includes("obsolete") && !offreDetails.qualite.includes("pourvue")) {
          offresEnAttente.push(offreDetails);
        }
      }
    }

    if (offresEnAttente.length === 0) { nbIgnores++; continue; }

    // 4. Construction du mail
    let offresHtml = "";
    offresEnAttente.forEach(o => {
      offresHtml += `
        <div style="background: rgba(255,255,255,0.05); padding: 15px; margin-bottom: 12px; border-radius: 8px; border-left: 4px solid #FF7A00; text-align: left;">
            <strong style="color: #ffffff; font-size: 16px;">${o.poste} <span style="color: #FF7A00; font-size: 12px;">(missions internes)</span></strong><br>
            <span style="color: #d1e3ff; font-size: 14px;">🏢 ${o.entreprise}</span><br>
            <a href="${o.url}" target="_blank" style="display: inline-block; margin-top: 10px; color: #FF7A00; text-decoration: none; font-weight: bold; font-size: 14px;">Voir l'annonce ↗</a>
        </div>
      `;
    });

    let htmlTemplate = PropertiesService.getScriptProperties().getProperty("TEMPLATE_RELANCE");
    if (!htmlTemplate) htmlTemplate = "Erreur : Modèle non configuré.";

    let htmlBody = htmlTemplate
      .replace("{{PRENOM}}", prenom)
      .replace("{{NB_OFFRES}}", offresEnAttente.length)
      .replace("{{INSERER_VOS_OFFRES_ICI}}", offresHtml);

    // 5. Envoi avec capture d'erreur précise
    try {
      GmailApp.sendEmail(mailEtu, "🎯 " + offresEnAttente.length + " offre(s) d'alternance en attente !", "", { htmlBody: htmlBody });
      nbMailsEnvoyes++;
      mailsRelances.push(mailEtu);
    } catch (e) {
      nbErreurs++;
      erreurDetails.push(mailEtu + " → " + e.toString());
      Logger.log("Erreur d'envoi pour " + mailEtu + " : " + e.toString());
    }
  }

  Logger.log("Relances terminées : " + nbMailsEnvoyes + " mail(s) envoyé(s), " + nbIgnores + " ignoré(s), " + nbErreurs + " erreur(s).");
  if (mailsRelances.length > 0) {
    try { logMailDetail(ss, { auteur: "Système (CRON)", categorie: "Relance auto", sousCat: "Relance étudiant hebdo", objet: "🎯 Offre(s) d'alternance en attente", destinataires: mailsRelances, corps: "Relance hebdomadaire des offres d'alternance en attente. " + nbMailsEnvoyes + " étudiant(s) relancé(s).", sourceId: "RELANCE-HEBDO-" + new Date().toLocaleDateString('fr-FR') }); } catch(eLog) {}
  }
  logExecutionAuto(ss, "envoyerRelanceHebdomadaire", nbMailsEnvoyes, nbIgnores, nbErreurs, erreurDetails.join(" | "));
}
function forcerAutorisations() {
  SpreadsheetApp.getActiveSpreadsheet();
  try { DriveApp.getRootFolder(); } catch(e) {}
  try { GmailApp.getAliases(); } catch(e) {}
}

/**
 * Alerte automatique hebdomadaire ciblée (Préinscrits & Inscrits en recherche)
 * Envoyé chaque lundi à 10h à polepedagogique@hecg.fr
 */
function envoyerAlertesSuiviPeda() {
  try {
    const SPREADSHEET_ID = "1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ";
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetEtu = getSheetSafe(ss, "ETUDIANTS");
    const sheetNotes = getSheetSafe(ss, "Carnet_route");
    
    const dataEtu = sheetEtu.getDataRange().getValues();
    const dataNotes = sheetNotes.getDataRange().getValues();
    const today = new Date();
    today.setHours(0, 0, 0, 0); // Normalisation à minuit pour éviter le drift du trigger
    
    // 1. Map des derniers échanges (Carnet de route)
    const mapDernierEchange = {};
    for (let i = 1; i < dataNotes.length; i++) {
      let id = String(dataNotes[i][0]).trim();
      let dateNote = new Date(dataNotes[i][2]);
      if (!mapDernierEchange[id] || dateNote > mapDernierEchange[id]) {
        mapDernierEchange[id] = dateNote;
      }
    }

    let listeAlerte7 = "";
    let listeAlerte30 = "";
    let count7 = 0;
    let count30 = 0;

    // 2. Scan des étudiants avec filtres stricts
    for (let j = 1; j < dataEtu.length; j++) {
      let id = String(dataEtu[j][0]).trim();
      let nom = dataEtu[j][2];
      let prenom = dataEtu[j][3];
      let statut = String(dataEtu[j][13]).toLowerCase(); // Colonne N
      let typo = String(dataEtu[j][20]).toLowerCase();   // Colonne U
      
      // FILTRE : Uniquement Préinscrits OU (Inscrits ET en recherche)
      // On exclut d'office Alternance, Initial, Prospects, Sorties.
      let estCible = (typo === "préinscrit") || (typo === "inscrit" && statut.includes("recherche"));
      
      if (!estCible) continue;

      let dernierContact = mapDernierEchange[id];
      let diffJours = 0;
      let texteJours = "";

      if (!dernierContact) {
        diffJours = 999; // Cas "Jamais contacté"
        texteJours = "Jamais contacté";
      } else {
        let diffMs = today - dernierContact;
        diffJours = Math.floor(diffMs / (1000 * 60 * 60 * 24));
        texteJours = diffJours + " jours";
      }

      let ligne = `<li><b>${nom.toUpperCase()} ${prenom}</b> : ${texteJours} sans échange</li>`;

      if (diffJours > 30) {
        listeAlerte30 += ligne;
        count30++;
      } else if (diffJours > 7) {
        listeAlerte7 += ligne;
        count7++;
      }
    }

    // 3. Construction du mail
    if (count7 === 0 && count30 === 0) return; // On n'envoie rien si tout est à jour

    let htmlBody = `
      <div style="font-family: 'Segoe UI', Arial, sans-serif; color: #333; line-height: 1.6;">
        <div style="background-color: #1A4E8A; color: white; padding: 20px; border-radius: 10px 10px 0 0;">
          <h1 style="margin: 0; font-size: 20px;"> Rapport de Suivi Pédagogique</h1>
          <p style="margin: 5px 0 0 0; opacity: 0.8;">Cible : Préinscrits & Inscrits en recherche d'entreprise</p>
        </div>
        
        <div style="padding: 20px; border: 1px solid #1A4E8A; border-top: none; border-radius: 0 0 10px 10px;">
          <p>Bonjour l'équipe,</p>
          <p>Voici les profils prioritaires qui nécessitent une relance immédiate au <b>${today.toLocaleDateString('fr-FR')}</b> :</p>
          
          <h3 style="color: #d9534f; border-bottom: 2px solid #d9534f; padding-bottom: 5px;"> URGENCE : +30 JOURS OU JAMAIS (${count30})</h3>
          <ul style="list-style-type: none; padding-left: 0;">${listeAlerte30 || "<li>Aucun profil.</li>"}</ul>

          <h3 style="color: #f0ad4e; border-bottom: 2px solid #f0ad4e; padding-bottom: 5px; margin-top: 30px;"> ALERTE : +7 JOURS SANS CONTACT (${count7})</h3>
          <ul style="list-style-type: none; padding-left: 0;">${listeAlerte7 || "<li>Aucun profil.</li>"}</ul>
          
          <p style="margin-top: 30px; font-size: 13px; color: #666;">
            <i>Note : Ce rapport exclut les étudiants déjà placés (Alternance), les Initiaux et les simples Prospects.</i>
          </p>
        </div>
      </div>
    `;

    let nbEnvoyes = 0;
    let erreurEnvoi = "";
    try {
      MailApp.sendEmail({
        to: "polepedagogique@hecg.fr",
        subject: "Alertes Suivi HECG : " + (count7 + count30) + " profils prioritaires",
        htmlBody: htmlBody
      });
      nbEnvoyes = 1;
      try { logMailDetail(ss, { auteur: "Système (CRON)", categorie: "Alerte", sousCat: "Relance équipe hebdo — Suivi pédago", objet: "Alertes Suivi HECG : " + (count7 + count30) + " profils prioritaires", destinataires: ["polepedagogique@hecg.fr"], corps: "Rapport hebdo : " + count30 + " profils URGENCE (+30j), " + count7 + " profils ALERTE (+7j).", sourceId: "SUIVI-PEDA-" + new Date().toLocaleDateString('fr-FR') }); } catch(eLog) {}
    } catch (mailErr) {
      erreurEnvoi = mailErr.toString();
      console.error("Erreur envoi alerte suivi : " + erreurEnvoi);
    }

    logExecutionAuto(ss, "envoyerAlertesSuiviPeda", nbEnvoyes, (count7 + count30), nbEnvoyes === 0 ? 1 : 0, erreurEnvoi);

  } catch (e) {
    console.error("Erreur alerte : " + e.toString());
    try {
      const ss2 = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");
      logExecutionAuto(ss2, "envoyerAlertesSuiviPeda", 0, 0, 1, "CRASH global : " + e.toString());
    } catch(e2) {}
  }
}

// ==========================================
// NOTIFICATION EMAIL — ACTIONS CRITIQUES
// ==========================================
function notifierSuppression(auteur, role, type, details) {
  try {
    const date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    const sujet = "⚠️ [HECG CRM] Suppression : " + type;
    const corps = `<div style="font-family:Arial,sans-serif;padding:20px;max-width:600px;border:2px solid #e74c3c;border-radius:12px;">
      <h2 style="color:#e74c3c;margin-top:0;">⚠️ Action de suppression détectée</h2>
      <table style="width:100%;border-collapse:collapse;">
        <tr><td style="padding:6px 0;color:#555;font-weight:bold;width:130px;">Type :</td><td style="padding:6px 0;color:#333;">${type}</td></tr>
        <tr><td style="padding:6px 0;color:#555;font-weight:bold;">Réalisée par :</td><td style="padding:6px 0;color:#333;">${auteur} (${role})</td></tr>
        <tr><td style="padding:6px 0;color:#555;font-weight:bold;">Date :</td><td style="padding:6px 0;color:#333;">${date}</td></tr>
        <tr><td style="padding:6px 0;color:#555;font-weight:bold;">Détail :</td><td style="padding:6px 0;color:#333;">${details}</td></tr>
      </table>
      <p style="margin-top:16px;font-size:11px;color:#999;">Ce message est automatique — CRM HECG.</p>
    </div>`;
    GmailApp.sendEmail("m.mokhtari@hecg.fr", sujet, "", { htmlBody: corps });
  } catch(eNotif) {
    console.error("Erreur envoi notification suppression : " + eNotif.toString());
  }
}

function logAction(idAuteur, roleAuteur, typeAction, categorie, idCible, details) {
  try {
    // On récupère le classeur actif
    const ss = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");
    let sheetLog;
    
    try {
      sheetLog = getSheetSafe(ss, "HISTORIQUE");
    } catch (e) {
      // Si la feuille n'existe pas, on l'arrête là pour éviter un crash
      console.error("La feuille HISTORIQUE est introuvable");
      return;
    }

    const idAction = "LOG-" + Date.now() + "-" + Math.floor(Math.random() * 1000);
    
    // On ajoute la ligne
    sheetLog.appendRow([
      idAction, 
      new Date(), 
      idAuteur || "Inconnu", 
      roleAuteur || "N/A", 
      typeAction, 
      categorie, 
      idCible || "-", 
      details || ""
    ]);
    
  } catch (e) {
    console.error("Erreur critique dans logAction : " + e.toString());
  }
}

/**
 * Enregistre un e-mail sortant dans LOGS_MAILS_DETAIL.
 * @param {Object} ss - Spreadsheet
 * @param {Object} data - { auteur, categorie, sousCat, objet, destinataires[], corps, sourceId }
 */
function logMailDetail(ss, data) {
  try {
    var sheet = ss.getSheetByName("LOGS_MAILS_DETAIL");
    if (!sheet) {
      sheet = ss.insertSheet("LOGS_MAILS_DETAIL");
      sheet.appendRow(["ID","Horodatage","Auteur","Categorie","Sous_Categorie","Objet","Nb_Destinataires","Destinataires","Apercu_Corps","Source_ID"]);
      sheet.setFrozenRows(1);
    }
    var emails = Array.isArray(data.destinataires) ? data.destinataires.filter(function(e){ return e && String(e).includes("@"); }) : [];
    var apercu = String(data.corps || "").replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim().substring(0, 400);
    var id = "MAIL-" + Date.now() + "-" + Math.floor(Math.random() * 1000);
    sheet.appendRow([id, new Date(), data.auteur || "Système", data.categorie || "Autre", data.sousCat || "", data.objet || "", emails.length, emails.join(";"), apercu, data.sourceId || ""]);
    return id;
  } catch(e) { console.error("Erreur logMailDetail : " + e.toString()); return null; }
}

/**
 * Journal complet des e-mails sortants — réservé Super-admin Vie Scolaire.
 * Source principale : LOGS_MAILS_DETAIL. Fallback : LOGS_EMAILS (historique auto CRON).
 */
function handleGetMailsLog(p, ss) {
  if (p.vieScRole !== "Super-admin") return createJsonResponse({ success: false, message: "Accès réservé au Super-admin." });
  try {
    var entries = [];

    // ── 1. LOGS_MAILS_DETAIL (source principale) ─────────────────────────
    try {
      var wsD = ss.getSheetByName("LOGS_MAILS_DETAIL");
      if (wsD && wsD.getLastRow() > 1) {
        wsD.getDataRange().getValues().slice(1).forEach(function(r) {
          entries.push({
            id: String(r[0]), date: r[1], auteur: String(r[2] || ""),
            categorie: String(r[3] || ""), sousCat: String(r[4] || ""),
            objet: String(r[5] || ""), nbDest: Number(r[6] || 0),
            apercu: String(r[8] || "").substring(0, 120), sourceId: String(r[9] || ""),
            hasDetail: true
          });
        });
      }
    } catch(eD) {}

    // ── 2. HISTORIQUE (envois manuels historiques : campagnes + offres) ────────
    try {
      var wsH = ss.getSheetByName("HISTORIQUE");
      if (wsH && wsH.getLastRow() > 1) {
        // Éviter les doublons avec LOGS_MAILS_DETAIL (sourceId déjà présent)
        var detailSourceIds = new Set(entries.filter(function(e){ return e.hasDetail; }).map(function(e){ return e.sourceId; }).filter(Boolean));
        wsH.getDataRange().getValues().slice(1).forEach(function(r) {
          var typeAction = String(r[4] || "");
          var categorie  = String(r[5] || "");
          var details    = String(r[7] || "");
          var sourceId   = String(r[6] || "");
          var cat = null;
          if (typeAction === "Action" && categorie === "Campagne") cat = "Com & Événement";
          else if (typeAction === "Action" && categorie === "Offres" && details.includes("partagé")) cat = "Offre partagée";
          if (!cat) return;
          if (detailSourceIds.has(sourceId)) return; // Déjà dans LOGS_MAILS_DETAIL
          var nbMatch = details.match(/(\d+)\s+destinataire/);
          var objetParse = "";
          var objetParts = details.split("Objet:");
          if (objetParts[1]) {
            objetParse = objetParts[1].replace(/«\s*/,"").replace(/\s*».*/,"").trim();
          } else {
            objetParse = details.split("|")[0].replace(/^(Test|Envoi réel)\s*—\s*/,"").trim().substring(0,80);
          }
          entries.push({
            id: "HIST-" + String(r[0]), date: r[1], auteur: String(r[2] || ""),
            categorie: cat, sousCat: details.split("—")[0].trim().substring(0,40),
            objet: objetParse, nbDest: nbMatch ? parseInt(nbMatch[1]) : 0,
            apercu: details, sourceId: sourceId, hasDetail: false
          });
        });
      }
    } catch(eH) {}

    // ── 3. LOGS_EMAILS (relances auto CRON sans LOGS_MAILS_DETAIL) ───────────
    try {
      var wsL = ss.getSheetByName("LOGS_EMAILS");
      if (wsL && wsL.getLastRow() > 1) {
        wsL.getDataRange().getValues().slice(1).forEach(function(r) {
          var nbEnv = Number(r[3] || 0);
          entries.push({
            id: "LEMAIL-" + String(r[0]), date: r[1], auteur: "Système (Auto)",
            categorie: "Relance auto", sousCat: String(r[2] || ""),
            objet: String(r[2] || ""), nbDest: nbEnv,
            apercu: "Envoyés : " + nbEnv + " | Ignorés : " + (r[4]||0) + (Number(r[5]) > 0 ? " | Erreurs : " + r[5] : ""),
            sourceId: "", hasDetail: false
          });
        });
      }
    } catch(eL) {}

    entries.sort(function(a, b){ return new Date(b.date) - new Date(a.date); });

    // ── Filtres ──────────────────────────────────────────────────────────
    var dateDebut  = p.dateDebut ? new Date(p.dateDebut)              : null;
    var dateFin    = p.dateFin   ? new Date(p.dateFin + "T23:59:59") : null;
    var catFiltree = String(p.categorie || "");
    var recherche  = String(p.recherche  || "").toLowerCase();

    var filtered = entries.filter(function(e) {
      var d = new Date(e.date);
      if (dateDebut && d < dateDebut) return false;
      if (dateFin   && d > dateFin)   return false;
      if (catFiltree && !e.categorie.toLowerCase().includes(catFiltree.toLowerCase()) && !e.sousCat.toLowerCase().includes(catFiltree.toLowerCase())) return false;
      if (recherche && !e.objet.toLowerCase().includes(recherche)
                    && !e.auteur.toLowerCase().includes(recherche)
                    && !e.apercu.toLowerCase().includes(recherche)
                    && !e.sousCat.toLowerCase().includes(recherche)
                    && !e.sourceId.toLowerCase().includes(recherche)) return false;
      return true;
    });

    var stats = { total: filtered.length, campagne: 0, offre: 0, relance: 0, alerte: 0, echange: 0 };
    filtered.forEach(function(e) {
      var c = (e.categorie + " " + e.sousCat).toLowerCase();
      if (c.includes("événement") || c.includes("evenement") || c.includes("com &")) stats.campagne++;
      else if (c.includes("offre")) stats.offre++;
      else if (c.includes("relance")) stats.relance++;
      else if (c.includes("alerte")) stats.alerte++;
      else if (c.includes("échange") || c.includes("echange") || c.includes("carnet")) stats.echange++;
    });

    return createJsonResponse({ success: true, data: filtered.slice(0, 500), stats: stats });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

/**
 * Retourne le détail complet d'un mail loggué (destinataires + aperçu corps).
 */
function handleGetMailDetail(p, ss) {
  if (p.vieScRole !== "Super-admin") return createJsonResponse({ success: false, message: "Accès réservé." });
  try {
    var id = String(p.id || "");
    var wsD = ss.getSheetByName("LOGS_MAILS_DETAIL");
    if (wsD && wsD.getLastRow() > 1) {
      var data = wsD.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]) === id) {
          var dest = String(data[i][7] || "");
          var emails = dest ? dest.split(";").map(function(e){ return e.trim(); }).filter(function(e){ return e; }) : [];
          return createJsonResponse({ success: true, data: {
            id: String(data[i][0]), date: data[i][1],
            auteur: String(data[i][2] || ""), categorie: String(data[i][3] || ""),
            sousCat: String(data[i][4] || ""), objet: String(data[i][5] || ""),
            nbDest: Number(data[i][6] || 0), destinataires: emails,
            apercuCorps: String(data[i][8] || ""), sourceId: String(data[i][9] || "")
          }});
        }
      }
    }
    return createJsonResponse({ success: false, message: "Entrée introuvable." });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

/**
 * Mouchard dédié aux exécutions automatiques (triggers CRON).
 * Écrit dans l'onglet LOGS_EMAILS (créé si absent).
 */
function logExecutionAuto(ss, nomFonction, nbEnvoyes, nbIgnores, nbErreurs, detailErreurs) {
  try {
    let sheet;
    try {
      sheet = getSheetSafe(ss, "LOGS_EMAILS");
    } catch (e) {
      sheet = ss.insertSheet("LOGS_EMAILS");
      sheet.appendRow(["ID", "Horodatage", "Fonction", "Nb Envoyés", "Nb Ignorés", "Nb Erreurs", "Détails Erreurs"]);
      sheet.setFrozenRows(1);
    }
    const id = "AUTO-" + Date.now();
    sheet.appendRow([id, new Date(), nomFonction, nbEnvoyes || 0, nbIgnores || 0, nbErreurs || 0, detailErreurs || ""]);
  } catch (e) {
    console.error("Impossible d'écrire dans LOGS_EMAILS : " + e.toString());
  }
}

/**
 * RÉCUPÉRATION DE L'HISTORIQUE — Réservé aux Administrateurs uniquement.
 */
function handleGetHistory(p, ss) {
  if (p.roleAuteur !== "Administrateur") {
    return createJsonResponse({ success: false, message: "Accès refusé" });
  }

  try {
    const sheetLog = getSheetSafe(ss, "HISTORIQUE");
    const data = sheetLog.getDataRange().getValues();

    if (data.length <= 1) return createJsonResponse({ success: true, logs: [], studentsMap: {}, personnelNames: [] });

    const headers = data[0];
    const rows = data.slice(1).reverse();

    let logs = rows.map(r => {
      let obj = {};
      headers.forEach((h, i) => obj[h] = r[i]);
      return obj;
    });


    // Map id étudiant → "Prénom Nom"
    const studentsMap = {};
    try {
      const dE = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues();
      for (let i = 1; i < dE.length; i++) {
        if (!dE[i][0]) continue;
        const id = String(dE[i][0]).trim();
        const fullName = (String(dE[i][3] || "").trim() + " " + String(dE[i][2] || "").trim()).trim();
        if (id && fullName) studentsMap[id] = fullName;
      }
    } catch(eE) {}

    // Liste des noms du Personnel (format "Prénom Nom" identique aux logs)
    const personnelNames = [];
    try {
      const dP = getSheetSafe(ss, "Personnel").getDataRange().getValues();
      for (let i = 1; i < dP.length; i++) {
        if (!dP[i][0]) continue;
        const fullName = (String(dP[i][2] || "").trim() + " " + String(dP[i][1] || "").trim()).trim();
        if (fullName) personnelNames.push(fullName);
      }
    } catch(eP) {}

    return createJsonResponse({ success: true, logs, studentsMap, personnelNames });
  } catch (e) {
    return createJsonResponse({ success: false, message: e.toString() });
  }
}


// ==========================================
// INDICATEURS CRM — BILAN JOURNALIER & HEBDO
// ==========================================

/**
 * Calcule tous les KPIs pour une période donnée.
 * dateDebut / dateFin au format "yyyy-MM-dd" (string) ou null pour tout.
 */
function computeIndicateurs(ss, dateDebut, dateFin) {
  var debut = dateDebut ? new Date(dateDebut + 'T00:00:00') : null;
  var fin   = dateFin   ? new Date(dateFin   + 'T23:59:59') : null;

  // ── HISTORIQUE ────────────────────────────────────────────────────
  var shLog   = getSheetSafe(ss, "HISTORIQUE");
  var dataLog = shLog.getDataRange().getValues();
  var hdrs    = dataLog[0];

  function col(keywords) {
    for (var ki = 0; ki < keywords.length; ki++) {
      var kw = keywords[ki];
      for (var hi = 0; hi < hdrs.length; hi++) {
        if (String(hdrs[hi]).toLowerCase().includes(kw)) return hi;
      }
    }
    return -1;
  }
  var iDate   = col(['date']);          if (iDate   < 0) iDate   = 1;
  var iAuteur = col(['auteur']);        if (iAuteur < 0) iAuteur = 2;
  var iRole   = col(['role','rôle']);   if (iRole   < 0) iRole   = 3;
  var iType   = col(['type','action']); if (iType   < 0) iType   = 4;
  var iCat    = col(['cat']);           if (iCat    < 0) iCat    = 5;
  var iCible  = col(['cible']);         if (iCible  < 0) iCible  = 6;
  var iDetail = col(['détail','detail']);if (iDetail < 0) iDetail = 7;

  // Personnel set
  var personnelSet = new Set();
  try {
    var dP = getSheetSafe(ss, "Personnel").getDataRange().getValues();
    for (var pi = 1; pi < dP.length; pi++) {
      if (!dP[pi][0]) continue;
      var pn = ((dP[pi][2] || "") + " " + (dP[pi][1] || "")).trim();
      if (pn) personnelSet.add(pn.toLowerCase());
    }
  } catch(e) {}

  // Étudiants map id → nom
  var dE = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues();
  var studentsMap = {};
  for (var ei = 1; ei < dE.length; ei++) {
    if (!dE[ei][0]) continue;
    studentsMap[String(dE[ei][0]).trim()] = ((dE[ei][3] || "") + " " + (dE[ei][2] || "")).trim();
  }

  // Compteurs
  var totalActions = 0;
  var offresVerifiees = 0, notesCarnetEtu = 0, notesCarnetEnt = 0;
  var contactsEntCreated = 0, contactsEntDeleted = 0;
  var etuCreated = 0, etuDeleted = 0, infosModifiees = 0;
  var encadrantsCount = {}, etudiantsCount = {};

  // Détails par KPI
  var dTotal = [], dOffresVerif = [], dNotesEtu = [], dNotesEnt = [];
  var dContCreated = [], dContDeleted = [], dEtuCreated = [], dEtuDeleted = [], dInfosModif = [];

  function makeEntry(row) {
    var dv = row[iDate];
    var ds = dv ? new Date(dv).toLocaleString('fr-FR', {day:'2-digit',month:'2-digit',year:'numeric',hour:'2-digit',minute:'2-digit'}) : '-';
    return {
      date:      ds,
      auteur:    String(row[iAuteur] || "Inconnu"),
      role:      String(row[iRole]   || ""),
      type:      String(row[iType]   || ""),
      categorie: String(row[iCat]    || ""),
      cible:     String(row[iCible]  || ""),
      detail:    String(row[iDetail] || "")
    };
  }

  var STU_CATS = ['dossier étudiant','carnet de route','note étudiant','note etudiant'];

  dataLog.slice(1).forEach(function(row) {
    var dv = row[iDate];
    if (!dv) return;
    var ld = new Date(dv);
    if (isNaN(ld)) return;
    if (debut && ld < debut) return;
    if (fin   && ld > fin)   return;

    var type   = String(row[iType]   || "").toLowerCase();
    var cat    = String(row[iCat]    || "").toLowerCase();
    var auteur = String(row[iAuteur] || "Inconnu");
    var detail = String(row[iDetail] || "").toLowerCase();
    var cible  = String(row[iCible]  || "");
    var entry  = makeEntry(row);

    totalActions++;
    dTotal.push(entry);

    // Encadrants
    if (personnelSet.size === 0 || personnelSet.has(auteur.toLowerCase())) {
      encadrantsCount[auteur] = (encadrantsCount[auteur] || 0) + 1;
    }

    // Étudiants les plus actifs (via logs carnet/dossier)
    if (STU_CATS.some(function(c){ return cat.includes(c); }) && cible) {
      var nom = studentsMap[cible] || cible;
      etudiantsCount[nom] = (etudiantsCount[nom] || 0) + 1;
    }

    // Offres vérifiées (log)
    if (cat === "offres" && detail.includes("→ vérifiée")) {
      offresVerifiees++;
      dOffresVerif.push(entry);
    }

    // Notes carnet étudiants
    if (cat.includes("carnet de route") && type.includes("créat")) {
      notesCarnetEtu++;
      dNotesEtu.push(entry);
    }

    // Notes carnet entreprise
    if (cat.includes("note partenaire") && type.includes("créat")) {
      notesCarnetEnt++;
      dNotesEnt.push(entry);
    }

    // Contacts entreprises créés
    if ((cat === "entreprise" || cat === "contact") && type.includes("créat")) {
      contactsEntCreated++;
      dContCreated.push(entry);
    }

    // Contacts entreprises supprimés
    if ((cat === "entreprise" || cat === "contact") && type.includes("suppr")) {
      contactsEntDeleted++;
      dContDeleted.push(entry);
    }

    // Étudiants créés
    if (cat.includes("dossier étudiant") && type.includes("créat")) {
      etuCreated++;
      dEtuCreated.push(entry);
    }

    // Étudiants supprimés
    if (cat.includes("dossier étudiant") && type.includes("suppr")) {
      etuDeleted++;
      dEtuDeleted.push(entry);
    }

    // Infos modifiées (coordonnées)
    if (type.includes("modif") && (cat.includes("dossier étudiant") || cat === "contact" || cat === "entreprise")) {
      infosModifiees++;
      dInfosModif.push(entry);
    }
  });

  // Classements
  var encadrantsRanking = Object.keys(encadrantsCount)
    .map(function(n){ return { name: n, count: encadrantsCount[n] }; })
    .sort(function(a,b){ return b.count - a.count; });
  var etudiantsRanking = Object.keys(etudiantsCount)
    .map(function(n){ return { name: n, count: etudiantsCount[n] }; })
    .sort(function(a,b){ return b.count - a.count; });

  // ── OFFRES (état actuel, non filtré par date) ─────────────────────
  var offresAQualifier = 0;
  var dOffresAQualif = [];
  try {
    var dO = getSheetSafe(ss, "OFFRES").getDataRange().getValues();
    for (var oi = 1; oi < dO.length; oi++) {
      if (!dO[oi][0]) continue;
      var qual = String(dO[oi][9] || "").trim();
      if (qual === "À qualifier" || qual === "A qualifier") {
        offresAQualifier++;
        dOffresAQualif.push({ id: dO[oi][0], poste: dO[oi][14] || '-', entreprise: dO[oi][4] || '-', ville: dO[oi][5] || '-' });
      }
    }
  } catch(e) {}

  // ── ÉTUDIANTS (état actuel, indicateurs statiques) ────────────────
  var dernierEchangeMap = {};
  try {
    var dC = getSheetSafe(ss, "Carnet_route").getDataRange().getValues();
    for (var ci = 1; ci < dC.length; ci++) {
      var eId = String(dC[ci][0]);
      var de  = new Date(dC[ci][2]);
      if (!isNaN(de) && (!dernierEchangeMap[eId] || de > dernierEchangeMap[eId])) {
        dernierEchangeMap[eId] = de;
      }
    }
  } catch(e) {}

  var unMoisAvant = new Date(); unMoisAvant.setDate(unMoisAvant.getDate() - 30);

  var etuSansAlt = [], etuAvecAlt = [], etuCVNon = [], etuNonContact = [];

  for (var i = 1; i < dE.length; i++) {
    if (!dE[i][0]) continue;
    var id2    = String(dE[i][0]).trim();
    var nom2   = ((dE[i][3] || "") + " " + (dE[i][2] || "")).trim();
    var stat   = String(dE[i][13] || "");
    var classe = String(dE[i][14] || "");
    var fCV    = String(dE[i][15] || "Non");
    var typo   = String(dE[i][20] || "");
    var idEnt  = String(dE[i][11] || "").trim();
    var campus = String(dE[i][22] || "");

    // Normalisation : PRÉINSCRIT (avec accent) est la valeur réelle dans le Sheet
    var typoNorm = typo.toLowerCase().replace(/[éèêë]/g, 'e').trim();
    var actif      = (typoNorm === "preinscrit" || typoNorm === "inscrit");
    var enRecherche = stat.toLowerCase().trim() === "en recherche";

    if (actif && enRecherche && !idEnt)
      etuSansAlt.push({ id: id2, nom: nom2, classe: classe, campus: campus, statut: stat });

    if (actif && enRecherche && idEnt)
      etuAvecAlt.push({ id: id2, nom: nom2, classe: classe, campus: campus });

    if (actif && enRecherche && fCV !== "Oui")
      etuCVNon.push({ id: id2, nom: nom2, classe: classe, campus: campus });

    if (typoNorm === "preinscrit" && enRecherche) {
      var dc = dernierEchangeMap[id2];
      if (!dc || dc < unMoisAvant) {
        etuNonContact.push({ id: id2, nom: nom2, classe: classe, campus: campus,
          dernierContact: dc ? dc.toLocaleDateString('fr-FR') : 'Jamais' });
      }
    }
  }

  return {
    periode: { debut: dateDebut || 'tout', fin: dateFin || 'tout' },
    totalActions:       totalActions,
    offresVerifiees:    offresVerifiees,
    offresAQualifier:   offresAQualifier,
    notesCarnetEtu:     notesCarnetEtu,
    notesCarnetEnt:     notesCarnetEnt,
    contactsEntCreated: contactsEntCreated,
    contactsEntDeleted: contactsEntDeleted,
    etuCreated:         etuCreated,
    etuDeleted:         etuDeleted,
    infosModifiees:     infosModifiees,
    encadrantsRanking:  encadrantsRanking,
    etudiantsRanking:   etudiantsRanking,
    etuSansAlternance:        { count: etuSansAlt.length,    detail: etuSansAlt },
    etuAvecAlternance:        { count: etuAvecAlt.length,    detail: etuAvecAlt },
    etuCVNonFait:             { count: etuCVNon.length,      detail: etuCVNon },
    etuPreinscritNonContacte: { count: etuNonContact.length, detail: etuNonContact },
    details: {
      actions:          dTotal,
      offresVerif:      dOffresVerif,
      offresAQualifier: dOffresAQualif,
      notesEtu:         dNotesEtu,
      notesEnt:         dNotesEnt,
      contactsCreated:  dContCreated,
      contactsDeleted:  dContDeleted,
      etuCreated:       dEtuCreated,
      etuDeleted:       dEtuDeleted,
      infosModif:       dInfosModif
    }
  };
}

/**
 * KPIs additionnels propres au rapport hebdomadaire.
 */
function computeHebdoKPIs(ss, lundi, dimanche) {
  // Offres : taux de transformation
  var dO = getSheetSafe(ss, "OFFRES").getDataRange().getValues();
  var totalRecu = 0, totalVerif = 0, semRecu = 0, semVerif = 0;
  var lundiD = new Date(lundi); lundiD.setHours(0,0,0,0);
  var dimD   = new Date(dimanche); dimD.setHours(23,59,59,999);

  for (var oi = 1; oi < dO.length; oi++) {
    if (!dO[oi][0]) continue;
    var dOff  = new Date(dO[oi][2]);
    var qual  = String(dO[oi][9] || "").trim().toLowerCase().replace(/[éèêë]/g, 'e');
    // Seules les offres explicitement marquées "Vérifiée" en colonne J comptent
    var isVerif = (qual === "verifiee");
    totalRecu++;
    if (isVerif) totalVerif++;
    if (!isNaN(dOff) && dOff >= lundiD && dOff <= dimD) {
      semRecu++;
      if (isVerif) semVerif++;
    }
  }

  // Étudiants actifs en recherche par campus
  var dE = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues();
  var actifsByCampus = {}, totalActifs = 0;
  for (var ei = 1; ei < dE.length; ei++) {
    if (!dE[ei][0]) continue;
    var stat    = String(dE[ei][13] || "").toLowerCase().trim();
    var typoH   = String(dE[ei][20] || "").toLowerCase().replace(/[éèêë]/g, 'e').trim();
    var campus  = String(dE[ei][22] || "Principal");
    if ((typoH === "preinscrit" || typoH === "inscrit") && stat === "en recherche") {
      totalActifs++;
      actifsByCampus[campus] = (actifsByCampus[campus] || 0) + 1;
    }
  }

  // Objectifs de la semaine
  var objTotal = 0, objAtteints = 0;
  try {
    var mondayStr = Utilities.formatDate(lundi, Session.getScriptTimeZone(), "yyyy-MM-dd");
    var wsObj = getOrCreateObjectifsSheet(ss);
    var dObj  = wsObj.getDataRange().getValues();
    for (var oo = 1; oo < dObj.length; oo++) {
      var semO = fmtSheetDate(dObj[oo][1]);
      if (semO === mondayStr) {
        objTotal++;
        if (normValide(dObj[oo][8]) === "TRUE") objAtteints++;
      }
    }
  } catch(e) {}

  return {
    totalOffresRecu:          totalRecu,
    totalOffresVerif:         totalVerif,
    tauxTransfoTotal:         totalRecu > 0 ? Math.round(totalVerif / totalRecu * 100) : 0,
    semOffresRecu:            semRecu,
    semOffresVerif:           semVerif,
    tauxTransfoSemaine:       semRecu > 0 ? Math.round(semVerif / semRecu * 100) : 0,
    totalActifsEnRecherche:   totalActifs,
    actifsByCampus:           actifsByCampus,
    ratioOffresEtu:           totalActifs > 0 ? (totalVerif / totalActifs).toFixed(2) : "N/A",
    objTotal:                 objTotal,
    objAtteints:              objAtteints,
    tauxObjectifs:            objTotal > 0 ? Math.round(objAtteints / objTotal * 100) : 0
  };
}

function handleGetBilanJournalier(p, ss) {
  if (p.roleAuteur !== "Administrateur") return createJsonResponse({ success: false, message: "Accès refusé" });
  try {
    var tz    = Session.getScriptTimeZone();
    var today = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
    var debut = p.dateDebut || today;
    var fin   = p.dateFin   || today;
    var data  = computeIndicateurs(ss, debut, fin);
    return createJsonResponse({ success: true, data: data });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleGetRapportHebdomadaire(p, ss) {
  if (p.roleAuteur !== "Administrateur") return createJsonResponse({ success: false, message: "Accès refusé" });
  try {
    var tz  = Session.getScriptTimeZone();
    var now = new Date();
    var dow = now.getDay();
    var diffL = (dow === 0) ? -6 : 1 - dow;
    var lundi = new Date(now); lundi.setDate(now.getDate() + diffL); lundi.setHours(0,0,0,0);
    var dim   = new Date(lundi); dim.setDate(lundi.getDate() + 6); dim.setHours(23,59,59,999);

    var debut = p.dateDebut || Utilities.formatDate(lundi, tz, "yyyy-MM-dd");
    var fin   = p.dateFin   || Utilities.formatDate(dim,   tz, "yyyy-MM-dd");

    var data  = computeIndicateurs(ss, debut, fin);
    data.hebdo = computeHebdoKPIs(ss, lundi, dim);
    return createJsonResponse({ success: true, data: data });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// ── EMAILS AUTOMATIQUES ────────────────────────────────────────────

function sendBilanJournalierAuto() {
  try {
    var ss     = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");
    var tz     = Session.getScriptTimeZone();
    var today  = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
    var todayFr= Utilities.formatDate(new Date(), tz, "dd/MM/yyyy");
    var data   = computeIndicateurs(ss, today, today);
    var html   = buildBilanJournalierHTML(data, todayFr);
    GmailApp.sendEmail("m.mokhtari@hecg.fr", "📊 Bilan du jour HECG — " + todayFr, "",
      { htmlBody: html, name: "CRM HECG — Automatique", from: SENDER_EMAIL });
    logAction("Système","Auto","Action","Bilan journalier","Mail","Bilan du " + todayFr + " envoyé");
  } catch(e) { Logger.log("Erreur sendBilanJournalierAuto : " + e.toString()); }
}

function sendRapportHebdomadaireAuto() {
  try {
    var ss  = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");
    var tz  = Session.getScriptTimeZone();
    var now = new Date();
    var dow = now.getDay();
    var diffL = (dow === 0) ? -6 : 1 - dow;
    var lundi = new Date(now); lundi.setDate(now.getDate() + diffL); lundi.setHours(0,0,0,0);
    var dim   = new Date(lundi); dim.setDate(lundi.getDate() + 6); dim.setHours(23,59,59,999);
    var lundiStr = Utilities.formatDate(lundi, tz, "dd/MM/yyyy");
    var dimStr   = Utilities.formatDate(dim,   tz, "dd/MM/yyyy");
    var debut    = Utilities.formatDate(lundi, tz, "yyyy-MM-dd");
    var fin      = Utilities.formatDate(dim,   tz, "yyyy-MM-dd");

    var data = computeIndicateurs(ss, debut, fin);
    data.hebdo = computeHebdoKPIs(ss, lundi, dim);
    var html = buildRapportHebdomadaireHTML(data, lundiStr, dimStr);
    GmailApp.sendEmail("m.mokhtari@hecg.fr", "📈 Rapport Hebdomadaire HECG — Semaine du " + lundiStr, "",
      { htmlBody: html, name: "CRM HECG — Automatique", from: SENDER_EMAIL });
    logAction("Système","Auto","Action","Rapport hebdomadaire","Mail","Rapport semaine " + lundiStr + " → " + dimStr + " envoyé");
  } catch(e) { Logger.log("Erreur sendRapportHebdomadaireAuto : " + e.toString()); }
}

// ── BUILDERS EMAIL HTML ────────────────────────────────────────────

function _kpiRow(label, value, color, bg) {
  return '<tr style="background:' + (bg||'#fff') + '"><td style="padding:9px 14px;font-size:13px">' + label + '</td>'
       + '<td style="padding:9px 14px;font-weight:800;font-size:17px;color:' + (color||'#1A4E8A') + '">' + value + '</td></tr>';
}

function _rankTable(ranking, color) {
  if (!ranking || !ranking.length) return '<p style="color:#9ca3af;font-size:12px;font-style:italic">Aucune donnée</p>';
  return '<table style="width:100%;border-collapse:collapse">'
    + ranking.slice(0,5).map(function(e,i){
        return '<tr><td style="padding:4px 8px;font-weight:700;width:28px">' + (i+1) + '</td>'
             + '<td style="padding:4px 8px">' + e.name + '</td>'
             + '<td style="padding:4px 8px;font-weight:700;color:' + color + '">' + e.count + ' action' + (e.count>1?'s':'') + '</td></tr>';
      }).join('') + '</table>';
}

function buildBilanJournalierHTML(data, dateStr) {
  var h = '<table style="width:100%;border-collapse:collapse">';
  return '<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;max-width:680px;margin:auto;background:#f1f5f9;padding:24px">'
  + '<div style="background:#1A4E8A;color:#fff;padding:26px 30px;border-radius:12px 12px 0 0">'
  + '<h1 style="margin:0;font-size:21px">📊 Bilan du jour — ' + dateStr + '</h1>'
  + '<p style="margin:6px 0 0;opacity:.8;font-size:12px">CRM HECG — Rapport automatique journalier (16h30)</p></div>'
  + '<div style="background:#fff;padding:24px 30px;border-radius:0 0 12px 12px;box-shadow:0 2px 10px rgba(0,0,0,.08)">'

  + '<h3 style="color:#1A4E8A;font-size:13px;text-transform:uppercase;border-bottom:2px solid #e2e8f0;padding-bottom:6px">📈 Activité du jour</h3>'
  + h
  + _kpiRow('Total actions', data.totalActions, '#1A4E8A', '#f8fafc')
  + _kpiRow('✅ Offres passées en vérifiées', data.offresVerifiees, '#7c3aed')
  + _kpiRow('⏳ Offres encore à qualifier (total actuel)', data.offresAQualifier, '#d97706', '#fffbeb')
  + _kpiRow('📓 Notes carnet étudiants', data.notesCarnetEtu, '#f59e0b')
  + _kpiRow('🏢 Notes carnet entreprise', data.notesCarnetEnt, '#6366f1', '#f8fafc')
  + '</table>'

  + '<h3 style="color:#1A4E8A;font-size:13px;text-transform:uppercase;border-bottom:2px solid #e2e8f0;padding-bottom:6px;margin-top:18px">🏢 Contacts Entreprises</h3>'
  + h
  + _kpiRow('➕ Nouveaux contacts créés', data.contactsEntCreated, '#16a34a', '#f0fdf4')
  + _kpiRow('🗑️ Contacts supprimés', data.contactsEntDeleted, '#dc2626', '#fef2f2')
  + '</table>'

  + '<h3 style="color:#1A4E8A;font-size:13px;text-transform:uppercase;border-bottom:2px solid #e2e8f0;padding-bottom:6px;margin-top:18px">🎓 Étudiants</h3>'
  + h
  + _kpiRow('➕ Nouveaux étudiants créés', data.etuCreated, '#16a34a', '#f0fdf4')
  + _kpiRow('🗑️ Étudiants supprimés', data.etuDeleted, '#dc2626', '#fef2f2')
  + _kpiRow('✏️ Informations modifiées (coordonnées)', data.infosModifiees, '#2563eb', '#eff6ff')
  + '</table>'

  + '<div style="display:flex;gap:20px;margin-top:18px">'
  + '<div style="flex:1"><h3 style="color:#1A4E8A;font-size:12px;text-transform:uppercase;border-bottom:2px solid #e2e8f0;padding-bottom:5px">🏆 Encadrants les plus actifs</h3>'
  + _rankTable(data.encadrantsRanking, '#1A4E8A') + '</div>'
  + '<div style="flex:1"><h3 style="color:#f59e0b;font-size:12px;text-transform:uppercase;border-bottom:2px solid #fde68a;padding-bottom:5px">🎓 Étudiants les plus actifs</h3>'
  + _rankTable(data.etudiantsRanking, '#f59e0b') + '</div>'
  + '</div>'

  + '<h3 style="color:#1A4E8A;font-size:13px;text-transform:uppercase;border-bottom:2px solid #e2e8f0;padding-bottom:6px;margin-top:18px">🔍 Situation actuelle — Recherche d\'alternance</h3>'
  + h
  + _kpiRow('🔴 En recherche SANS alternance (préinscrits + inscrits)', data.etuSansAlternance.count,   '#dc2626', '#fef2f2')
  + _kpiRow('🟢 En recherche AVEC alternance (préinscrits + inscrits)', data.etuAvecAlternance.count,   '#16a34a', '#f0fdf4')
  + _kpiRow('📄 En recherche — CV non fait',                            data.etuCVNonFait.count,         '#d97706', '#fffbeb')
  + _kpiRow('⚠️ Préinscrits en recherche non contactés &gt; 1 mois',   data.etuPreinscritNonContacte.count, '#dc2626', '#fef2f2')
  + '</table>'

  + '<p style="margin-top:22px;color:#9ca3af;font-size:10px;text-align:center">Rapport généré automatiquement — CRM HECG — Consultez les détails sur l\'onglet Historique → Indicateurs du quotidien</p>'
  + '</div></body></html>';
}

function buildRapportHebdomadaireHTML(data, lundiStr, dimStr) {
  var hb = data.hebdo || {};
  var h  = '<table style="width:100%;border-collapse:collapse">';

  var campusRows = Object.keys(hb.actifsByCampus || {}).map(function(c){
    return '<tr><td style="padding:5px 12px">' + c + '</td><td style="padding:5px 12px;font-weight:700;color:#1A4E8A">' + (hb.actifsByCampus[c]) + ' étu.</td>'
         + '<td style="padding:5px 12px;color:#555">' + hb.totalOffresVerif + ' offres vérif. / ' + (hb.actifsByCampus[c]) + ' étu. = <b>' + (hb.actifsByCampus[c] > 0 ? (hb.totalOffresVerif / hb.actifsByCampus[c]).toFixed(1) : 'N/A') + '</b></td></tr>';
  }).join('');

  var tauxColor = (hb.tauxObjectifs || 0) >= 70 ? '#16a34a' : (hb.tauxObjectifs >= 50 ? '#d97706' : '#dc2626');

  return '<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;max-width:680px;margin:auto;background:#f1f5f9;padding:24px">'
  + '<div style="background:linear-gradient(135deg,#1A4E8A,#1d4ed8);color:#fff;padding:28px 30px;border-radius:12px 12px 0 0">'
  + '<h1 style="margin:0;font-size:22px">📈 Rapport Hebdomadaire HECG</h1>'
  + '<p style="margin:7px 0 0;opacity:.85;font-size:13px">Semaine du ' + lundiStr + ' au ' + dimStr + ' — Envoi automatique vendredi 17h30</p></div>'
  + '<div style="background:#fff;padding:24px 30px;border-radius:0 0 12px 12px;box-shadow:0 2px 10px rgba(0,0,0,.08)">'

  + '<h3 style="color:#1A4E8A;font-size:13px;text-transform:uppercase;border-bottom:2px solid #e2e8f0;padding-bottom:6px">🎯 Indicateurs Majeurs</h3>'
  + h
  + _kpiRow('📊 Taux de transformation total (offres vérifiées / reçues)',
      (hb.tauxTransfoTotal||0) + '% (' + (hb.totalOffresVerif||0) + ' / ' + (hb.totalOffresRecu||0) + ')', '#7c3aed', '#f5f3ff')
  + _kpiRow('📅 Taux de transformation cette semaine',
      (hb.tauxTransfoSemaine||0) + '% (' + (hb.semOffresVerif||0) + ' / ' + (hb.semOffresRecu||0) + ')', '#6366f1', '#eef2ff')
  + _kpiRow('🎓 Étudiants actifs en recherche (total)',
      (hb.totalActifsEnRecherche||0) + ' étudiants', '#1A4E8A', '#eff6ff')
  + _kpiRow('🏆 Objectifs atteints cette semaine',
      (hb.tauxObjectifs||0) + '% (' + (hb.objAtteints||0) + '/' + (hb.objTotal||0) + ')', tauxColor, '#f8fafc')
  + '</table>'

  + (campusRows ? '<h4 style="color:#555;font-size:11px;margin-top:12px;text-transform:uppercase">Offres vérifiées / Étudiants en recherche — par campus</h4>'
    + h + campusRows + '</table>' : '')

  + '<h3 style="color:#1A4E8A;font-size:13px;text-transform:uppercase;border-bottom:2px solid #e2e8f0;padding-bottom:6px;margin-top:20px">📈 Activité de la semaine</h3>'
  + h
  + _kpiRow('Total actions', data.totalActions, '#1A4E8A', '#f8fafc')
  + _kpiRow('✅ Offres passées en vérifiées', data.offresVerifiees, '#7c3aed')
  + _kpiRow('⏳ Offres encore à qualifier (total actuel)', data.offresAQualifier, '#d97706', '#fffbeb')
  + _kpiRow('📓 Notes carnet étudiants', data.notesCarnetEtu, '#f59e0b')
  + _kpiRow('🏢 Notes carnet entreprise', data.notesCarnetEnt, '#6366f1', '#f8fafc')
  + _kpiRow('➕ Nouveaux contacts entreprises', data.contactsEntCreated, '#16a34a', '#f0fdf4')
  + _kpiRow('🗑️ Contacts entreprises supprimés', data.contactsEntDeleted, '#dc2626', '#fef2f2')
  + _kpiRow('➕ Nouveaux étudiants créés', data.etuCreated, '#16a34a', '#f0fdf4')
  + _kpiRow('🗑️ Étudiants supprimés', data.etuDeleted, '#dc2626', '#fef2f2')
  + _kpiRow('✏️ Informations modifiées', data.infosModifiees, '#2563eb', '#eff6ff')
  + '</table>'

  + '<div style="display:flex;gap:20px;margin-top:18px">'
  + '<div style="flex:1"><h3 style="color:#1A4E8A;font-size:12px;text-transform:uppercase;border-bottom:2px solid #e2e8f0;padding-bottom:5px">🏆 Encadrants les plus actifs</h3>'
  + _rankTable(data.encadrantsRanking, '#1A4E8A') + '</div>'
  + '<div style="flex:1"><h3 style="color:#f59e0b;font-size:12px;text-transform:uppercase;border-bottom:2px solid #fde68a;padding-bottom:5px">🎓 Étudiants les plus actifs</h3>'
  + _rankTable(data.etudiantsRanking, '#f59e0b') + '</div>'
  + '</div>'

  + '<h3 style="color:#1A4E8A;font-size:13px;text-transform:uppercase;border-bottom:2px solid #e2e8f0;padding-bottom:6px;margin-top:18px">🔍 Situation actuelle — Recherche d\'alternance</h3>'
  + h
  + _kpiRow('🔴 En recherche SANS alternance', data.etuSansAlternance.count,   '#dc2626', '#fef2f2')
  + _kpiRow('🟢 En recherche AVEC alternance', data.etuAvecAlternance.count,   '#16a34a', '#f0fdf4')
  + _kpiRow('📄 En recherche — CV non fait',   data.etuCVNonFait.count,         '#d97706', '#fffbeb')
  + _kpiRow('⚠️ Préinscrits non contactés &gt; 1 mois', data.etuPreinscritNonContacte.count, '#dc2626', '#fef2f2')
  + '</table>'

  + '<p style="margin-top:22px;color:#9ca3af;font-size:10px;text-align:center">Rapport hebdomadaire automatique — CRM HECG — Consultez les détails sur l\'onglet Historique → Indicateurs majeurs</p>'
  + '</div></body></html>';
}

/**
 * À exécuter UNE SEULE FOIS depuis l'éditeur Apps Script.
 * Crée les déclencheurs temporels automatiques.
 */
function setupIndicateursTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    var fn = t.getHandlerFunction();
    if (fn === 'sendBilanJournalierAuto' || fn === 'sendRapportHebdomadaireAuto') ScriptApp.deleteTrigger(t);
  });
  // Bilan journalier : tous les jours à ~16h (Apps Script ne garantit pas les minutes)
  ScriptApp.newTrigger('sendBilanJournalierAuto').timeBased().everyDays(1).atHour(16).create();
  // Rapport hebdo : chaque vendredi à ~17h
  ScriptApp.newTrigger('sendRapportHebdomadaireAuto').timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(17).create();
  Logger.log("✅ Triggers créés : bilan journalier 16h + rapport hebdo vendredi 17h");
}

function extraireVille(texte) {
  if (!texte) return "";
  var mArr = texte.match(/marseille[\s\-]*([0-9]{1,2})(?:\s*e(?:r|me|eme)?)?(?:\s*arrondissement)?/i);
  if (mArr) { var n = parseInt(mArr[1]); return "Marseille " + n + (n === 1 ? "er" : "e"); }
  if (/marseille/i.test(texte)) return "Marseille";
  var villes = ["Aix-en-Provence","Paris","Lyon","Bordeaux","Toulon","Nice","Montpellier","Lille","Nantes","Strasbourg","Grenoble","Toulouse","Rennes","Rouen"];
  for (var i = 0; i < villes.length; i++) { if (new RegExp(villes[i],'i').test(texte)) return villes[i]; }
  return "";
}
function handleAspirerOffresGmail(p, ss) {
  try {
    // Cherche les mails non lus de l'alerte Google
    const threads = GmailApp.search('from:notify-noreply@google.com subject:"Alertes d\'emploi de Google" is:unread', 0, 5);
    const wsOffres = getSheetSafe(ss, "OFFRES");
    let nbOffresAjoutees = 0;

    threads.forEach(thread => {
      const messages = thread.getMessages();
      messages.forEach(msg => {
        if (msg.isUnread()) {
          const body = msg.getPlainBody(); // Le texte brut du mail
          
          // --- LOGIQUE D'EXTRACTION BASIQUE ---
          // Google Alertes met souvent l'intitulé du poste et l'entreprise sur des lignes séparées.
          // Ici on capture l'alerte de façon générique (il faudra sûrement affiner avec tes vrais mails)
          const lignes = body.split('\n').map(l => l.trim()).filter(l => l.length > 0);
          
          // Exemple d'ajout brut dans le Sheet (Génère un ID, met la date, etc.)
          const idOffre = "OFF-" + Date.now() + "-" + Math.floor(Math.random()*100);
          const dateJour = new Date().toLocaleDateString('fr-FR');
          
          // On ajoute une ligne (Colonne A: ID, B: Date, C: Titre, etc... À ADAPTER SELON TES COLONNES)
          var villeDetectee = extraireVille(body);
          wsOffres.appendRow([idOffre, dateJour, "Alerte Google Job", "Import Auto", "À traiter", villeDetectee, "", "", "", "", "", "", "", "Nouvelle"]);
          
          nbOffresAjoutees++;
          msg.markRead(); // On marque comme lu pour ne pas le ré-aspirer demain
        }
      });
    });

    logAction(p.idAuteur, p.roleAuteur, "Création", "Offres", "Import", `Aspiré ${nbOffresAjoutees} offres depuis Gmail`);
    return createJsonResponse({ success: true, message: `${nbOffresAjoutees} offres importées depuis Gmail !` });
  } catch (e) {
    return createJsonResponse({ success: false, message: "Erreur aspiration : " + e.toString() });
  }
}

function handleGetClassesList(ss) {
  const d = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues();
  const classes = [];
  for (let i = 1; i < d.length; i++) {
    let c = String(d[i][14] || "").trim(); // Colonne O (Index 14)
    if (c && !classes.includes(c)) classes.push(c);
  }
  return createJsonResponse({ success: true, data: classes.sort() });
}

function diagnosticLiens() {
  // On prend juste LE DERNIER mail d'alerte
  const threads = GmailApp.search('from:(notify-noreply@google.com OR googlealerts-noreply@google.com)', 0, 1);
  
  if (threads.length === 0) {
    console.log("Aucun mail trouvé pour le test.");
    return;
  }
  
  const msg = threads[0].getMessages()[0];
  const htmlBody = msg.getBody();
  
  // On cherche ABSOLUMENT TOUS les liens du mail
  const tousLesLiens = htmlBody.match(/href="([^"]+)"/g);
  
  if (tousLesLiens) {
    console.log("--- VOICI LES LIENS CACHÉS DANS VOTRE MAIL ---");
    // On affiche les 15 premiers liens trouvés
    for(let i = 0; i < Math.min(15, tousLesLiens.length); i++) {
      console.log(tousLesLiens[i]);
    }
  } else {
    console.log("Aucun lien trouvé. Le mail est peut-être du texte brut.");
  }
}

function diagnosticUltime() {
  const threads = GmailApp.search('from:(notify-noreply@google.com OR googlealerts-noreply@google.com)', 0, 1);
  
  if (threads.length === 0) {
    console.log("Aucun mail trouvé.");
    return;
  }
  
  const msg = threads[0].getMessages()[0];
  
  // On récupère le texte brut du mail (sans aucun code caché)
  const texteBrut = msg.getPlainBody();
  
  // On cherche tout ce qui commence par "http" et s'arrête à un espace ou retour à la ligne
  const tousLesLiensBruts = texteBrut.match(/https?:\/\/[^\s]+/g);
  
  if (tousLesLiensBruts) {
    console.log("--- BINGO ! VOICI LES LIENS TROUVÉS DANS LE TEXTE ---");
    // On affiche les 15 premiers liens trouvés
    for(let i = 0; i < Math.min(15, tousLesLiensBruts.length); i++) {
      console.log("Lien " + (i+1) + " : " + tousLesLiensBruts[i]);
    }
  } else {
    console.log("Décidément, il n'y a aucun lien commençant par http dans ce mail !");
    // On affiche les 500 premiers caractères du mail pour voir de quoi il est fait
    console.log("Aperçu du contenu : " + texteBrut.substring(0, 500));
  }
}

/**
 * Enregistre un nouvel échange dans l'onglet SUIVI_PARTENARIAT
 * Colonnes attendues : ID_Partenaire | Auteur | Date | Compte_rendu
 */
function handleAddPartnerNote(params, ss) {
  try {
    const sheet = getSheetSafe(ss, "SUIVI_PARTENARIAT");

    // Sécurisation de l'auteur
    const auteurFinal = params.idAuteur || params.auteur || "Système";
    const roleFinal = params.roleAuteur || "N/A";

    const idPartenaire = String(params.id_entreprise || params.targetId || "").trim();
    const nouvelleLigne = [
      idPartenaire,
      auteurFinal,
      new Date(),
      params.texte
    ];

    sheet.appendRow(nouvelleLigne);

    const apercuNote = params.texte ? String(params.texte).substring(0, 80) + (params.texte.length > 80 ? "…" : "") : "";
    logAction(auteurFinal, roleFinal, "Création", "Note Partenaire", idPartenaire, "Échange ajouté : « " + apercuNote + " »");

    return { 
      success: true, 
      message: "Note enregistrée avec succès dans le carnet de bord." 
    };
  } catch (e) {
    return { 
      success: false, 
      message: "Erreur lors de l'enregistrement de la note : " + e.toString() 
    };
  }
}

function handleGetPartnerHistory(params, ss) {
  try {
    const id = String(params.id_entreprise || params.targetId || "").trim().toUpperCase();
    const sheet = getSheetSafe(ss, "SUIVI_PARTENARIAT");
    const rows = sheet.getDataRange().getValues();
    const data = [];
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0] || "").trim().toUpperCase() !== id) continue;
      data.push({ auteur: rows[i][1], date: rows[i][2], texte: rows[i][3] });
    }
    data.reverse();
    return createJsonResponse({ success: true, data });
  } catch(e) {
    return createJsonResponse({ success: false, data: [], message: e.toString() });
  }
}

function handleGetActivityStats(p, ss) {
  try {
    const period = parseInt(p.period) || 0; // 0=tout, sinon nb de jours
    const now = new Date();
    const tz = Session.getScriptTimeZone();

    // Email → Prénom Nom (PERSONNEL col G=email, C=prénom, B=nom)
    const persMap = {};
    try {
      const dPers = getSheetSafe(ss, "PERSONNEL").getDataRange().getValues();
      for (let i = 1; i < dPers.length; i++) {
        const em = String(dPers[i][6] || "").trim().toLowerCase();
        if (em) persMap[em] = (String(dPers[i][2]||"")+" "+String(dPers[i][1]||"")).trim();
      }
    } catch(e) {}

    // ID étudiant → Prénom Nom (ETUDIANTS col A=id, C=nom, D=prénom)
    const etuMap = {};
    try {
      const dEtu = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues();
      for (let i = 1; i < dEtu.length; i++) {
        const id = String(dEtu[i][0] || "").trim();
        if (id) etuMap[id] = (String(dEtu[i][3]||"")+" "+String(dEtu[i][2]||"")).trim();
      }
    } catch(e) {}

    const histData = getSheetSafe(ss, "HISTORIQUE").getDataRange().getValues();
    const STU_CATS = ['dossier étudiant','carnet de route','note étudiant','note etudiant'];
    const cutoff = period > 0 ? new Date(now.getTime() - period * 86400000) : null;

    const byPerson = {}, byStudent = {};
    // Pour "aujourd'hui" (dashboard) : cutoff = minuit aujourd'hui
    const todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const byPersonToday = {}, byStudentToday = {};

    for (let i = 1; i < histData.length; i++) {
      const row = histData[i];
      if (!row[1]) continue;
      const d = new Date(row[1]);
      const em = String(row[2] || "").trim().toLowerCase();
      const cat = String(row[5] || "").trim().toLowerCase();
      const cible = String(row[6] || "").trim();
      if (!em) continue;

      const name = persMap[em] || em;
      const isStu = STU_CATS.some(c => cat.includes(c));

      // Stats période (filtres trophées)
      if (!cutoff || d >= cutoff) {
        byPerson[name] = (byPerson[name] || 0) + 1;
        if (isStu && cible) {
          const sName = etuMap[cible] || cible;
          byStudent[sName] = (byStudent[sName] || 0) + 1;
        }
      }

      // Stats "aujourd'hui" (dashboard)
      if (d >= todayStart) {
        byPersonToday[name] = (byPersonToday[name] || 0) + 1;
        if (isStu && cible) {
          const sName = etuMap[cible] || cible;
          byStudentToday[sName] = (byStudentToday[sName] || 0) + 1;
        }
      }
    }

    const topPerson  = Object.keys(byPersonToday).sort((a,b) => byPersonToday[b]-byPersonToday[a])[0] || null;
    const topStudent = Object.keys(byStudentToday).sort((a,b) => byStudentToday[b]-byStudentToday[a])[0] || null;

    return createJsonResponse({
      success: true, byPerson, byStudent,
      topPerson,  topPersonCount:  topPerson  ? byPersonToday[topPerson]  : 0,
      topStudent, topStudentCount: topStudent ? byStudentToday[topStudent] : 0
    });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

/**
 * CRM HECG - Système de Sauvegarde Automatique
 * Crée une copie de la base dans le dossier spécifié 2 fois par jour.
 */

function backupDatabase() {
  // ID du dossier de destination (extrait de votre lien)
  const FOLDER_ID_BACKUPS = "1xTUfVgbMVT_puAyvUvO77HrGkIKwuKuC"; 
  const MAX_BACKUPS_TO_KEEP = 14; // Garde environ 1 semaine de copies (2x par jour)
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const folder = DriveApp.getFolderById(FOLDER_ID_BACKUPS);
    
    // Génération du nom avec date et heure
    const now = new Date();
    const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd_HH'h'mm");
    const backupName = "Sauvegarde_CRM_" + timestamp;
    
    // Création de la copie
    const currentFile = DriveApp.getFileById(ss.getId());
    currentFile.makeCopy(backupName, folder);
    
    console.log("✅ Sauvegarde réussie : " + backupName);
    
    // Nettoyage des anciennes versions
    cleanupOldBackups(folder, MAX_BACKUPS_TO_KEEP);
    
  } catch (e) {
    console.error("❌ Erreur lors de la sauvegarde : " + e.toString());
  }
}

/**
 * Supprime les fichiers les plus anciens si le nombre limite est dépassé.
 */
function cleanupOldBackups(folder, maxFiles) {
  const files = folder.getFiles();
  const fileList = [];
  
  while (files.hasNext()) {
    fileList.push(files.next());
  }
  
  // Trier par date de création (du plus récent au plus ancien)
  fileList.sort((a, b) => b.getDateCreated() - a.getDateCreated());
  
  // Supprimer les fichiers qui dépassent la limite
  if (fileList.length > maxFiles) {
    for (let i = maxFiles; i < fileList.length; i++) {
      fileList[i].setTrashed(true);
      console.log("🗑️ Ancienne sauvegarde supprimée : " + fileList[i].getName());
    }
  }
}
 function handleUpdateStudent(p, ss) {
    const idCherche = String(p.id_etu || "").trim().toUpperCase();
    let actualIdInSheet = idCherche;

    try {
      const sheet = getSheetSafe(ss, "ETUDIANTS");
      const data = sheet.getDataRange().getValues();
      let rowIndex = -1;

      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim().toUpperCase() === idCherche) {
          rowIndex = i + 1;
          break;
        }
      }

      if (rowIndex === -1) return createJsonResponse({ success: false, message: "Élève introuvable" });

      const oldRow = data[rowIndex - 1];
      actualIdInSheet = String(oldRow[0]);

      const oldIdEnt = String(oldRow[11] || "").trim();
      const newIdEnt = String(p.id_ent || "").trim();
      // Enregistrer l'historique dès qu'une entreprise est remplacée ou supprimée (toutes situations)
      if (oldIdEnt && newIdEnt !== oldIdEnt) {
        try {
          _saveAncienLien(ss, {
            idEtu: actualIdInSheet,
            nom: String(oldRow[2]||""),
            prenom: String(oldRow[3]||""),
            classe: String(oldRow[14]||""),
            campus: String(oldRow[22]||""),
            idEnt: oldIdEnt,
            nomEnt: String(oldRow[21]||""),
            motif: newIdEnt ? "Changement d'entreprise" : "Entreprise retirée"
          });
        } catch(e) {}
      }

      const typoCalcule = calculerTypoEtu(p.entree, p.sortie, p.classe);
      const typoFinale = typoCalcule === "SORTIE" ? "SORTIE" : (p.forceTypo || typoCalcule);

      const updatedRow = [
        oldRow[0],           // A : id_etu
        p.mdp || oldRow[1],  // B : MDP
        p.nom,               // C : Nom_etu
        p.prenom,            // D : Prenom_etu
        p.mail,              // E : Mail_etu
        p.tel,               // F : Tel_etu
        p.cv,                // G : CV
        p.refPerso,          // H : Ref_Perso
        oldRow[8]  || "",    // I : (préservé)
        oldRow[9]  || "",    // J : (préservé)
        oldRow[10] || "",    // K : (préservé)
        p.id_ent || "",      // L : id_entreprise
        p.id_cont || "",     // M : id_contact
        p.statut,            // N : Statut
        p.classe,            // O : Classe
        p.faitCV || "Non",   // P : Atelier_CV
        p.faitLI || "Non",   // Q : Atelier_LinkedIn
        "'" + p.entree,      // R : Entree_etu
        p.sortie,            // S : sortie_etu
        p.motifSortie !== undefined ? p.motifSortie : (oldRow[19] || ""),  // T : motifsortie
        typoFinale,          // U : Typo
        p.entreprise || "",  // V : Nom_entreprise
        p.campus || "",      // W : Campus
        oldRow[23] || ""     // X : Identifiant Elève
      ];

      sheet.getRange(rowIndex, 1, 1, updatedRow.length).setValues([updatedRow]);
    } catch (erreur) {
      return createJsonResponse({ success: false, message: erreur.toString() });
    }

    try {
      var diffParts = [];
      var fieldChecks = [
        [2, "Nom", p.nom], [3, "Prénom", p.prenom], [4, "Email", p.mail], [5, "Tél", p.tel],
        [13, "Statut", p.statut], [21, "Entreprise", p.entreprise], [14, "Classe", p.classe],
        [6, "CV", p.cv], [20, "Typo", undefined]
      ];
      fieldChecks.forEach(function(fc) {
        if (fc[2] === undefined) return;
        var oldVal = String(oldRow[fc[0]] || "").trim();
        var newVal = String(fc[2] || "").trim();
        if (oldVal !== newVal) diffParts.push(fc[1] + ": « " + oldVal + " » → « " + newVal + " »");
      });
      var texteLog = diffParts.length > 0 ? "Modifié — " + diffParts.join(" | ") : (p.detailsLog && p.detailsLog !== "" ? p.detailsLog : "Mise à jour fiche (aucun changement détecté)");
      logAction(p.idAuteur, p.roleAuteur, "Modification", "Dossier étudiant", actualIdInSheet, texteLog);
    } catch (logErreur) {
      console.error("Erreur de journalisation (non bloquante) : " + logErreur.toString());
    }

    return createJsonResponse({ success: true, message: "Mise à jour réussie" });
  }



  function handleChangeStudentMdp(p, ss) {
    try {
      const idCherche = String(p.id_etu || "").trim().toUpperCase();
      const sheet = getSheetSafe(ss, "ETUDIANTS");
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim().toUpperCase() === idCherche) {
          sheet.getRange(i + 1, 2).setValue(p.mdp);
          return createJsonResponse({ success: true });
        }
      }
      return createJsonResponse({ success: false, message: "Élève introuvable" });
    } catch (e) {
      return createJsonResponse({ success: false, message: e.toString() });
    }
  }

function handleChangeStudentMdp(p, ss) {
  try {
    const idCherche = String(p.id_etu || "").trim().toUpperCase();
    const sheet = getSheetSafe(ss, "ETUDIANTS");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toUpperCase() === idCherche) {
        sheet.getRange(i + 1, 2).setValue(p.mdp); // Colonne B uniquement
        return createJsonResponse({ success: true });
      }
    }
    return createJsonResponse({ success: false, message: "Élève introuvable" });
  } catch (e) {
    return createJsonResponse({ success: false, message: e.toString() });
  }
}

// ==========================================
// ALERTES ÉTUDIANTS (parallèle aux alertes partenaires)
// Sheet : ALERTES_ETUDIANTS — Colonnes : A:ID | B:ID_Etu | C:Date | D:Message | E:Fréquence | F:Statut | G:Email_Destinataire
// ==========================================

function handleSetStudentAlert(p, ss) {
  try {
    let sheet;
    try { sheet = getSheetSafe(ss, "ALERTES_ETUDIANTS"); }
    catch(e) {
      sheet = ss.insertSheet("ALERTES_ETUDIANTS");
      sheet.appendRow(["ID", "ID_Etudiant", "Date", "Message", "Frequence", "Statut", "Email_Destinataire"]);
      sheet.setFrozenRows(1);
    }
    const newId = "ALE" + Math.floor(1000 + Math.random() * 9000);
    const idEtu = String(p.targetId || p.id_etudiant || "").trim();
    sheet.appendRow([
      newId,
      idEtu,
      p.date || new Date(),
      p.message || p.msg || "Relance à effectuer",
      p.frequence || p.freq || "Unique",
      "À faire",
      p.destinataire || "m.mokhtari@hecg.fr"
    ]);
    const dateAlerte = p.date ? new Date(p.date).toLocaleDateString('fr-FR') : "?";
    const msgApercu = (p.message || p.msg || "Relance").substring(0, 60);
    logAction(p.idAuteur, p.roleAuteur, "Création", "Alertes Étudiants", idEtu, "Alerte programmée le " + dateAlerte + " : « " + msgApercu + " » (" + (p.frequence || p.freq || "Unique") + ")");
    const msnId = creerMissionDepuisAlerte(ss, p, "Étudiant", idEtu);
    notifierMMOKEtCreerMission(ss, p, "Étudiant", idEtu, "creation", p.message || p.msg || "Relance");
    const msgRetour = msnId ? "Alerte étudiant créée + Mission " + msnId + " déclenchée !" : "Alerte étudiant créée !";
    return createJsonResponse({ success: true, message: msgRetour, missionId: msnId || null });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleGetAllStudentAlerts(p, ss) {
  try {
    let sheet;
    try { sheet = getSheetSafe(ss, "ALERTES_ETUDIANTS"); }
    catch(e) { return createJsonResponse({ success: true, data: [] }); }
    const dA = sheet.getDataRange().getValues();
    const dE = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues();

    const etuMap = {};
    for (let j = 1; j < dE.length; j++) {
      const eid = String(dE[j][0] || "").trim();
      if (eid) etuMap[eid] = (String(dE[j][3] || "").trim() + " " + String(dE[j][2] || "").trim()).trim();
    }

    const list = [];
    for (let i = 1; i < dA.length; i++) {
      if (!dA[i][0] || dA[i][5] === "Terminé" || dA[i][5] === "Annulé") continue;
      const dateObj = new Date(dA[i][2]);
      list.push({
        id: dA[i][0],
        etuId: dA[i][1],
        nomEtudiant: etuMap[String(dA[i][1]).trim()] || ("ID : " + dA[i][1]),
        dateRaw: dateObj.toISOString(),
        date: dateObj.toLocaleDateString('fr-FR'),
        msg: dA[i][3],
        freq: dA[i][4],
        destinataire: dA[i][6] || ""
      });
    }
    list.sort((a, b) => new Date(a.dateRaw) - new Date(b.dateRaw));
    return createJsonResponse({ success: true, data: list });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleCompleteStudentAlert(p, ss) {
  try {
    const sheet = getSheetSafe(ss, "ALERTES_ETUDIANTS");
    const dA = sheet.getDataRange().getValues();
    for (let i = 1; i < dA.length; i++) {
      if (String(dA[i][0]) === p.id) {
        const idEtu = String(dA[i][1]);
        const msgAlerte = String(dA[i][3] || "").substring(0, 60);
        if (dA[i][4] === "Unique") {
          sheet.getRange(i + 1, 6).setValue("Terminé");
          logAction(p.idAuteur, p.roleAuteur, "Action", "Alertes Étudiants", idEtu, "Alerte marquée Terminée : « " + msgAlerte + " »");
        } else {
          const newDate = new Date(dA[i][2]);
          if (dA[i][4] === "Hebdomadaire") newDate.setDate(newDate.getDate() + 7);
          else if (dA[i][4] === "Mensuelle") newDate.setMonth(newDate.getMonth() + 1);
          else if (dA[i][4] === "Trimestrielle") newDate.setMonth(newDate.getMonth() + 3);
          sheet.getRange(i + 1, 3).setValue(newDate);
          logAction(p.idAuteur, p.roleAuteur, "Action", "Alertes Étudiants", idEtu, "Alerte récurrente complétée, prochaine: " + newDate.toLocaleDateString('fr-FR') + " : « " + msgAlerte + " »");
        }
        notifierMMOKEtCreerMission(ss, p, "Étudiant", idEtu, "cloture", msgAlerte);
        return createJsonResponse({ success: true });
      }
    }
    return createJsonResponse({ success: false, message: "Alerte non trouvée" });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleUpdateStudentAlert(p, ss) {
  try {
    const sheet = getSheetSafe(ss, "ALERTES_ETUDIANTS");
    const dA = sheet.getDataRange().getValues();
    for (let i = 1; i < dA.length; i++) {
      if (String(dA[i][0]) === p.id) {
        const idEtu = String(dA[i][1]);
        sheet.getRange(i + 1, 3).setValue(p.date);
        sheet.getRange(i + 1, 4).setValue(p.msg);
        sheet.getRange(i + 1, 5).setValue(p.freq);
        sheet.getRange(i + 1, 7).setValue(p.destinataire);
        const newDate = p.date ? new Date(p.date).toLocaleDateString('fr-FR') : "?";
        logAction(p.idAuteur, p.roleAuteur, "Modification", "Alertes Étudiants", idEtu, "Alerte modifiée → Date: " + newDate + " | « " + String(p.msg || "").substring(0, 60) + " »");
        return createJsonResponse({ success: true });
      }
    }
    return createJsonResponse({ success: false });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleCancelStudentAlert(p, ss) {
  try {
    const sheet = getSheetSafe(ss, "ALERTES_ETUDIANTS");
    const dA = sheet.getDataRange().getValues();
    for (let i = 1; i < dA.length; i++) {
      if (String(dA[i][0]) === p.id) {
        const idEtu = String(dA[i][1]);
        const msgAlerte = String(dA[i][3] || "").substring(0, 60);
        sheet.getRange(i + 1, 6).setValue("Annulé");
        logAction(p.idAuteur, p.roleAuteur, "Action", "Alertes Étudiants", idEtu, "Alerte annulée : « " + msgAlerte + " »");
        return createJsonResponse({ success: true });
      }
    }
    return createJsonResponse({ success: false });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// ==========================================
// ASPIRATION HELLOWORK
// ==========================================

function aspirerHelloWork() {
  const ss = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");
  const sheet = getSheetSafe(ss, "OFFRES");

  // 1. RECHERCHE : emails non lus de HelloWork
  const threads = GmailApp.search('from:(alerte@emails.hellowork.com OR notification@emails.hellowork.com) is:unread', 0, 20);
  console.log("HelloWork - Conversations examinées : " + threads.length);

  const data = sheet.getDataRange().getValues();

  // 2. ANTI-DOUBLON : clé (Poste + Entreprise) + URL
  const offresExistantes = data.map(r =>
    (String(r[14]) + String(r[4])).toLowerCase().replace(/\s+/g, '')
  );
  const urlsExistantes = data.map(r => String(r[1]).trim());

  let nbNouveautes = 0;

  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      if (!msg.isUnread()) return;

      const body = msg.getBody(); // HTML brut

      // 3. REGEX : un bloc complet d'offre HelloWork
      // Groupe 1 = URL (href dans <a class="dark-style-color...">)
      // Groupe 2 = Poste (texte du lien)
      // Groupe 3 = Entreprise (texte du <td class="dark-style-color"> suivant)
      // Groupe 4 = Lieu (texte du <td style="background-color:#F6F6F6"> suivant)
      const regexBloc = /<a\b[^>]*class="[^"]*dark-style-color[^"]*"[^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a>[\s\S]*?<td\b[^>]*class=(?:"[^"]*dark-style-color[^"]*"|dark-style-color)[^>]*>([\s\S]*?)<\/td>[\s\S]*?<td\b[^>]*background-color:#F6F6F6[^>]*>([\s\S]*?)<\/td>/gi;

      let match;
      let nbMatchsDansCeMessage = 0;
      while ((match = regexBloc.exec(body)) !== null) {
        const url        = match[1].replace(/&amp;/g, '&').trim();
        const poste      = match[2].replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim();
        const entreprise = match[3].replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim();
        const lieu       = match[4].replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim();

        if (!poste || !entreprise) continue;

        // 4. VÉRIFICATION DOUBLON
        const cleDoublon = (poste + entreprise).toLowerCase().replace(/\s+/g, '');
        if (offresExistantes.includes(cleDoublon) || urlsExistantes.includes(url)) continue;

        // 5. INSERTION
        const idOffre = "HW-" + Date.now() + "-" + Math.floor(Math.random() * 100);
        sheet.appendRow([
          idOffre,           // A : ID
          url,               // B : URL
          new Date(),        // C : Date d'aspiration
          "HelloWork",       // D : Plateforme
          entreprise,        // E : Entreprise
          lieu,              // F : Lieu
          "", "", "",        // G, H, I
          "À qualifier",     // J (index 9)
          "", "", "",        // K, L, M
          "Non partagée",    // N (index 13)
          poste,             // O (index 14)
          "Alternance",      // P
          "Alerte HelloWork" // Q
        ]);

        offresExistantes.push(cleDoublon);
        urlsExistantes.push(url);
        nbNouveautes++;
        nbMatchsDansCeMessage++;
      }

      console.log("HelloWork - Email analysé, regex matches : " + nbMatchsDansCeMessage + " | Sujet : " + msg.getSubject());

      // 6. MARQUER COMME LU seulement si le regex a trouvé des données (sinon conserver pour re-traitement après fix)
      if (nbMatchsDansCeMessage > 0) msg.markRead();
    });
  });

  console.log("HelloWork - TERMINÉ : " + nbNouveautes + " nouvelles offres ajoutées.");
}

// ==========================================
// TRIGGER : ASPIRATION QUOTIDIENNE 06h + 17h
// ==========================================

function creerTriggersAspiration() {
  // Supprimer les triggers toutAspirer existants pour éviter les doublons
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "toutAspirer") ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger("toutAspirer").timeBased().everyDays(1).atHour(6).create();
  ScriptApp.newTrigger("toutAspirer").timeBased().everyDays(1).atHour(17).create();

  console.log("Triggers créés : toutAspirer à 06h00 et 17h00.");
}

// ==========================================
// DIAGNOSTIC HELLOWORK (À supprimer après)
// ==========================================

function diagnosticHelloWork() {
  const threads = GmailApp.search('from:(alerte@emails.hellowork.com OR notification@emails.hellowork.com)', 0, 1);
  if (threads.length === 0) { console.log("Aucun email HelloWork trouvé."); return; }

  const msg = threads[0].getMessages()[0];
  const body = msg.getBody();

  console.log("=== SUJET ===");
  console.log(msg.getSubject());
  console.log("=== TAILLE HTML total : " + body.length + " caractères ===");

  // Trouver où commence le contenu des offres (après le gros bloc CSS)
  const bodyTagIdx = body.indexOf('<body');
  console.log("=== <body> commence à l'index : " + bodyTagIdx + " ===");

  // Chercher les premiers liens <a href= dans le body
  const firstHref = body.indexOf('href=', bodyTagIdx);
  console.log("=== Premier href à l'index : " + firstHref + " ===");

  // Montrer 3000 caractères autour du premier lien pour voir la structure réelle
  const start = Math.max(0, firstHref - 200);
  console.log("=== STRUCTURE AUTOUR DU PREMIER LIEN ===");
  console.log(body.substring(start, start + 3000));

  // Montrer le HTML autour du 1er <a class="dark-style-color"> pour voir la structure entreprise/lieu
  const idxDarkLink = body.indexOf('class="dark-style-color');
  if (idxDarkLink > -1) {
    console.log("=== HTML autour du 1er lien dark-style-color (structure offre) ===");
    console.log(body.substring(idxDarkLink - 100, idxDarkLink + 4000));
  }
}

// ==========================================
// MODULE MISSIONS — ASPIRATION, API
// ==========================================

/**
 * Scanne les emails ENVOYÉS par m.mokhtari@hecg.fr contenant #mission ou #MMOK,
 * applique un libellé "Mission_Traitee" pour éviter les doublons,
 * et insère les données dans l'onglet MISSIONS.
 * Structure MISSIONS : A:ID | B:Date | C:Heure | D:Email_Dest | E:Nom_Collab | F:CC | G:Objet | H:Corps | I:Statut | J:Notes | K:Date_Maj
 */
function aspirerMissionsGmail() {
  const ss = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");

  // --- Initialisation de la feuille MISSIONS ---
  let sheet;
  try {
    sheet = getSheetSafe(ss, "MISSIONS");
  } catch(e) {
    sheet = ss.insertSheet("MISSIONS");
    sheet.appendRow(["ID", "Date", "Heure", "Email_Dest", "Nom_Collab", "CC", "Objet", "Corps", "Statut", "Notes", "Date_Maj"]);
    sheet.setFrozenRows(1);
    sheet.getRange("1:1").setFontWeight("bold");
  }

  // --- Libellé anti-doublon ---
  const LABEL_NAME = "Mission_Traitee";
  let label = GmailApp.getUserLabelByName(LABEL_NAME);
  if (!label) label = GmailApp.createLabel(LABEL_NAME);

  // --- Récupération du référentiel Personnel pour les noms ---
  const dPerso = getSheetSafe(ss, "Personnel").getDataRange().getValues();
  const emailToNom = {};
  for (let i = 1; i < dPerso.length; i++) {
    const email = String(dPerso[i][6] || "").trim().toLowerCase();
    if (email) emailToNom[email] = (dPerso[i][2] + " " + dPerso[i][1]).trim(); // Nom Prénom
  }

  // --- Recherche dans les mails ENVOYÉS ---
  // On cherche dans "sent" les mails contenant les hashtags cibles
  const query = 'in:sent (#mission OR #Mission OR #MISSION OR #MMOK)';
  const threads = GmailApp.search(query, 0, 50);

  let nbInseres = 0;

  threads.forEach(function(thread) {
    // Si le thread a déjà le libellé → déjà traité, on skip
    const labels = thread.getLabels().map(l => l.getName());
    if (labels.includes(LABEL_NAME)) return;

    const messages = thread.getMessages();
    messages.forEach(function(msg) {
      const fromRaw = msg.getFrom() || "";
      // On ne traite que les mails envoyés par MMOK
      if (!fromRaw.toLowerCase().includes("m.mokhtari@hecg.fr")) return;

      const dateObj    = msg.getDate();
      const dateStr    = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "dd/MM/yyyy");
      const heureStr   = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "HH:mm");
      const toRaw      = msg.getTo() || "";
      const ccRaw      = msg.getCc() || "";
      const subject    = msg.getSubject() || "(Sans objet)";

      // Nettoyage du corps : plain text, sans les lignes de citation >
      let corps = msg.getPlainBody() || "";
      corps = corps.split("\n").filter(function(l) {
        return !l.trim().startsWith(">") && !l.trim().startsWith("On ") && l.trim() !== "";
      }).slice(0, 80).join("\n").trim();
      if (corps.length > 3000) corps = corps.substring(0, 3000) + "\n[...]";

      // Extraction du premier destinataire
      const emailDest = toRaw.split(",")[0].replace(/<|>/g, "").trim().toLowerCase();
      const nomCollab = emailToNom[emailDest] || emailDest;

      // Génération d'un ID unique
      const idMission = "MSN-" + dateObj.getTime() + "-" + Math.floor(Math.random() * 1000);

      sheet.appendRow([
        idMission,   // A: ID
        dateStr,     // B: Date
        heureStr,    // C: Heure
        emailDest,   // D: Email Destinataire
        nomCollab,   // E: Nom Collaborateur
        ccRaw,       // F: CC
        subject,     // G: Objet
        corps,       // H: Corps
        "Attribué",  // I: Statut
        "",          // J: Notes
        new Date()   // K: Date_Maj
      ]);

      nbInseres++;
    });

    // On applique le libellé pour ne plus retraiter ce thread
    thread.addLabel(label);
  });

  console.log("Missions Gmail : " + nbInseres + " mission(s) insérée(s).");
  if (nbInseres > 0) {
    logAction("Système", "Auto", "Création", "Missions", "Import Gmail", nbInseres + " mission(s) aspirée(s)");
  }
}

// ==========================================
// MISSIONS — HANDLERS API
// ==========================================

/**
 * Renvoie la liste des missions.
 * Admin/MMOK → toutes les missions.
 * Collaborateur → uniquement ses missions (filtrage sur Email_Dest).
 */
function handleGetMissions(p, ss) {
  try {
    let sheet;
    try { sheet = getSheetSafe(ss, "MISSIONS"); }
    catch(e) { return createJsonResponse({ success: true, data: [] }); }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return createJsonResponse({ success: true, data: [] });

    const isAdmin = p.roleAuteur === "Administrateur" || String(p.emailAuteur || "").toLowerCase() === "m.mokhtari@hecg.fr";

    const missions = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;

      // Filtrage par collaborateur si non-admin
      if (!isAdmin) {
        const emailDest = String(row[3] || "").toLowerCase();
        const emailUser = String(p.emailAuteur || "").toLowerCase();
        if (emailDest !== emailUser) continue;
      }

      missions.push({
        id:             String(row[0]),
        date:           String(row[1] || ""),
        heure:          String(row[2] || ""),
        email:          String(row[3] || ""),
        collab:         String(row[4] || ""),
        cc:             String(row[5] || ""),
        objet:          String(row[6] || ""),
        corps:          String(row[7] || ""),
        statut:         String(row[8] || "Attribué"),
        notes:          String(row[9] || ""),
        dateMaj:        row[10] ? new Date(row[10]).toLocaleDateString('fr-FR') : "",
        priorite:       row[11] ? Number(row[11]) : 3,
        resume:         String(row[12] || ""),
        emailCreateur:  String(row[13] || "")
      });
    }

    // Plus récentes en premier
    missions.sort(function(a, b) {
      return new Date(b.date.split('/').reverse().join('-')) - new Date(a.date.split('/').reverse().join('-'));
    });

    return createJsonResponse({ success: true, data: missions });
  } catch(e) {
    return createJsonResponse({ success: false, message: e.toString() });
  }
}

/**
 * Met à jour le statut d'une mission.
 * Réservé Administrateur ou m.mokhtari@hecg.fr.
 * Statuts valides : Attribué, Clôturé, Annulé, Problématique
 */
function handleUpdateMissionStatus(p, ss) {
  try {
    const isAdmin    = p.roleAuteur === "Administrateur" || String(p.emailAuteur || "").toLowerCase() === "m.mokhtari@hecg.fr";
    const emailUser  = String(p.emailAuteur || "").toLowerCase();
    const statutsValides = ["Attribué", "Clôturé", "Annulé", "Problématique"];

    if (!statutsValides.includes(p.statut)) {
      return createJsonResponse({ success: false, message: "Statut invalide : " + p.statut });
    }

    const sheet = getSheetSafe(ss, "MISSIONS");
    const data  = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(p.id)) {
        const emailCreateur = String(data[i][13] || "").toLowerCase();
        const isCreator     = emailCreateur !== "" && emailCreateur === emailUser;
        if (!isAdmin && !isCreator) {
          return createJsonResponse({ success: false, message: "Seul le créateur de cette mission ou un administrateur peut modifier son statut." });
        }
        const ancienStatut = String(data[i][8] || "");
        sheet.getRange(i + 1, 9).setValue(p.statut);   // Col I : Statut
        sheet.getRange(i + 1, 11).setValue(new Date()); // Col K : Date_Maj
        logAction(p.idAuteur, p.roleAuteur, "Modification", "Missions", p.id,
          "Statut : " + ancienStatut + " → " + p.statut + " | " + (data[i][6] || ""));
        // Sync : si mission Clôturée/Annulée → mettre la tâche planning liée à Terminé/Annulé
        if (p.statut === "Clôturé" || p.statut === "Annulé") {
          try {
            var planWs = ss.getSheetByName("PLANNING");
            if (planWs) {
              var planData = planWs.getDataRange().getValues();
              var newPlanStatut = p.statut === "Clôturé" ? "Terminé" : "Annulé";
              for (var pi = 1; pi < planData.length; pi++) {
                if (String(planData[pi][12]||"") === String(p.id)) {
                  planWs.getRange(pi+1, 6).setValue(newPlanStatut); break;
                }
              }
            }
          } catch(eSync) {}
        }
        return createJsonResponse({ success: true });
      }
    }
    return createJsonResponse({ success: false, message: "Mission introuvable." });
  } catch(e) {
    return createJsonResponse({ success: false, message: e.toString() });
  }
}

/**
 * Ajoute une note à une mission.
 * La note est horodatée et concaténée au contenu existant.
 */
function handleAddMissionNote(p, ss) {
  try {
    const sheet = getSheetSafe(ss, "MISSIONS");
    const data  = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(p.id)) {
        const now         = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
        const auteur      = p.idAuteur || "?";
        const nouvelleNote = "[" + now + " — " + auteur + "] " + String(p.note || "").trim();
        const notesExist  = String(data[i][9] || "").trim();
        const notesFinal  = notesExist ? notesExist + "\n" + nouvelleNote : nouvelleNote;

        sheet.getRange(i + 1, 10).setValue(notesFinal); // Col J : Notes
        sheet.getRange(i + 1, 11).setValue(new Date()); // Col K : Date_Maj
        logAction(p.idAuteur, p.roleAuteur, "Création", "Missions", p.id,
          "Note ajoutée sur mission : « " + String(p.note || "").substring(0, 60) + " »");
        return createJsonResponse({ success: true });
      }
    }
    return createJsonResponse({ success: false, message: "Mission introuvable." });
  } catch(e) {
    return createJsonResponse({ success: false, message: e.toString() });
  }
}

/**
 * Met à jour un champ libre d'une mission (priorite, resume).
 * Autorisé : propriétaire de la mission ou admin.
 */
function handleUpdateMissionField(p, ss) {
  try {
    const isAdmin = p.roleAuteur === "Administrateur" || String(p.emailAuteur || "").toLowerCase() === "m.mokhtari@hecg.fr";
    const allowed = ['priorite', 'resume'];
    if (!allowed.includes(p.field)) return createJsonResponse({ success: false, message: "Champ non autorisé." });
    const colMap = { priorite: 12, resume: 13 }; // col L, M (1-indexed)
    const col = colMap[p.field];

    const sheet = getSheetSafe(ss, "MISSIONS");
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) !== String(p.id)) continue;
      const emailDest = String(data[i][3] || "").toLowerCase();
      const emailUser = String(p.emailAuteur || "").toLowerCase();
      if (!isAdmin && emailDest !== emailUser) return createJsonResponse({ success: false, message: "Accès refusé." });
      sheet.getRange(i+1, col).setValue(p.field === 'priorite' ? Number(p.val) : String(p.val));
      sheet.getRange(i+1, 11).setValue(new Date());
      return createJsonResponse({ success: true });
    }
    return createJsonResponse({ success: false, message: "Mission introuvable." });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

/**
 * Création manuelle d'une mission.
 * Tout utilisateur peut créer pour lui-même.
 * m.mokhtari@hecg.fr / Administrateur peut créer pour n'importe qui.
 */
function handleAddMission(p, ss) {
  try {
    const isAdmin = p.roleAuteur === "Administrateur" || String(p.emailAuteur || "").toLowerCase() === "m.mokhtari@hecg.fr";
    let emailDest = String(p.emailDest || p.emailAuteur || "").toLowerCase();
    let nomCollab = String(p.nomCollab || p.idAuteur || "");
    if (!isAdmin) { emailDest = String(p.emailAuteur || "").toLowerCase(); nomCollab = p.idAuteur || ""; }

    const sheet = getSheetSafe(ss, "MISSIONS");
    const now   = new Date();
    const id    = "M" + now.getTime();
    sheet.appendRow([
      id,
      Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy"),
      Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm"),
      emailDest,
      nomCollab,
      String(p.cc || ""),
      String(p.objet || ""),
      String(p.corps || ""),
      "Attribué",
      "",
      now,
      Number(p.priorite || 3),
      String(p.resume || ""),
      String(p.emailAuteur || "")
    ]);
    logAction(p.idAuteur, p.roleAuteur, "Création", "Missions", id,
      "Mission manuelle : " + String(p.objet || "").substring(0,80) + " | Pour : " + nomCollab);
    return createJsonResponse({ success: true, id: id });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

/**
 * Déclenche l'aspiration Gmail à la demande depuis l'UI.
 * Réservé Administrateur / m.mokhtari@hecg.fr.
 * Retourne le nombre de missions insérées et la liste des objets trouvés.
 */
function handleAspirerMissionsNow(p, ss) {
  try {
    const isAdmin = p.roleAuteur === "Administrateur" ||
                    String(p.emailAuteur || "").toLowerCase() === "m.mokhtari@hecg.fr";
    if (!isAdmin) return createJsonResponse({ success: false, message: "Accès refusé." });

    // On relance l'aspiration et on capture le résultat
    const result = aspirerMissionsGmailAvecRetour(ss);
    return createJsonResponse({ success: true, nb: result.nb, objets: result.objets });
  } catch(e) {
    return createJsonResponse({ success: false, message: e.toString() });
  }
}

// Normalise un objet mail pour la déduplication : retire Re:/Fwd:/Tr: et met en minuscules
function normaliserObjetMission(s) {
  return String(s || "").replace(/^(re|fw|fwd|tr|réf|ref)\s*:\s*/gi, "").trim().toLowerCase();
}

/**
 * Version de l'aspiration qui retourne un objet { nb, objets[] } au lieu de juste logger.
 * Partage toute la logique avec aspirerMissionsGmail().
 */
function aspirerMissionsGmailAvecRetour(ss) {
  if (!ss) ss = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");

  let sheet;
  try { sheet = getSheetSafe(ss, "MISSIONS"); }
  catch(e) {
    sheet = ss.insertSheet("MISSIONS");
    sheet.appendRow(["ID","Date","Heure","Email_Dest","Nom_Collab","CC","Objet","Corps","Statut","Notes","Date_Maj","Gmail_ID"]);
    sheet.setFrozenRows(1);
    sheet.getRange("1:1").setFontWeight("bold");
  }

  // Ajoute l'en-tête Gmail_ID (col L) si la feuille existait sans cette colonne
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (!headerRow.includes("Gmail_ID")) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue("Gmail_ID");
  }

  const LABEL_NAME = "Mission_Traitee";
  let label = GmailApp.getUserLabelByName(LABEL_NAME);
  if (!label) label = GmailApp.createLabel(LABEL_NAME);

  // Référentiel Personnel email → Nom
  const dPerso = getSheetSafe(ss, "Personnel").getDataRange().getValues();
  const emailToNom = {};
  for (let i = 1; i < dPerso.length; i++) {
    const email = String(dPerso[i][6] || "").trim().toLowerCase();
    if (email) emailToNom[email] = (dPerso[i][2] + " " + dPerso[i][1]).trim();
  }

  // Boîte envoyée uniquement, lus et non lus, tous hashtags mission
  const query = 'in:sent (#mission OR #missions OR #Mission OR #MMOK)';
  const threads = GmailApp.search(query, 0, 500);

  // Index des doublons : ID Gmail (col L, index 11) pour les nouveaux enregistrements
  // + fallback emailDest|objet_normalisé (col D+G) pour les anciens sans Gmail_ID
  const existingData = sheet.getDataRange().getValues();
  const existingKeys = new Set();
  for (let i = 1; i < existingData.length; i++) {
    const gmailId = String(existingData[i][11] || "").trim();
    if (gmailId) {
      existingKeys.add(gmailId);
    } else {
      existingKeys.add(String(existingData[i][3]).toLowerCase() + '|' + normaliserObjetMission(String(existingData[i][6] || "")));
    }
  }

  let nbInseres = 0;
  const objetsInseres = [];

  threads.forEach(function(thread) {
    const messages = thread.getMessages();
    let threadHasNew = false;

    messages.forEach(function(msg) {
      const rawBody = msg.getPlainBody() || "";

      // Vérification stricte : le hashtag doit être dans le corps brut (hors citations)
      const bodyLines = rawBody.split("\n").filter(function(l) {
        return l.trim() !== "" && !l.trim().startsWith(">") && !l.trim().startsWith("On ");
      }).join("\n");
      if (!/#missions?/i.test(bodyLines) && !/#MMOK/.test(bodyLines)) return;

      const dateObj  = msg.getDate();
      const dateStr  = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "dd/MM/yyyy");
      const heureStr = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "HH:mm");
      const toRaw    = msg.getTo() || "";
      const emailDest = toRaw.split(",")[0].replace(/[<>]/g, "").replace(/.*\s/, "").trim().toLowerCase();
      const subject   = msg.getSubject() || "(Sans objet)";

      // Déduplication par message ID Gmail (évite les doublons sur relances multiples)
      const dupKey = msg.getId();
      if (existingKeys.has(dupKey)) return;
      existingKeys.add(dupKey);

      const ccRaw  = msg.getCc() || "";

      // Nettoyage corps : supprime les citations et lignes vides
      let corps = bodyLines.split("\n").slice(0, 80).join("\n").trim();
      if (corps.length > 3000) corps = corps.substring(0, 3000) + "\n[...]";

      const nomCollab = emailToNom[emailDest] || emailDest;
      const idMission = "MSN-" + dateObj.getTime() + "-" + Math.floor(Math.random() * 1000);

      sheet.appendRow([idMission, dateStr, heureStr, emailDest, nomCollab, ccRaw, subject, corps, "Attribué", "", new Date(), msg.getId(), "", SENDER_EMAIL]);
      nbInseres++;
      threadHasNew = true;
      objetsInseres.push(subject);
    });

    if (threadHasNew) thread.addLabel(label);
  });

  if (nbInseres > 0) {
    logAction("Système", "Auto", "Création", "Missions", "Import Gmail", nbInseres + " mission(s) aspirée(s) : " + objetsInseres.join(" | "));
  }
  return { nb: nbInseres, objets: objetsInseres };
}

// Réécrit aspirerMissionsGmail pour déléguer à la version avec retour
function aspirerMissionsGmail() {
  aspirerMissionsGmailAvecRetour(null);
}

/**
 * À exécuter UNE SEULE FOIS depuis l'éditeur Apps Script.
 * Installe un trigger automatique toutes les 30 minutes pour les missions.
 */
function creerTriggerMissions() {
  // Supprime l'ancien trigger missions s'il existe
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "aspirerMissionsGmail") ScriptApp.deleteTrigger(t);
  });
  // Crée le nouveau : toutes les 30 minutes
  ScriptApp.newTrigger("aspirerMissionsGmail")
    .timeBased()
    .everyMinutes(30)
    .create();
  console.log("Trigger missions créé : toutes les 30 minutes.");
}

/**
 * Crée une mission dans l'onglet MISSIONS depuis une alerte (partenaire ou étudiant)
 * quand le message contient #MMOK, #mission ou #Mission.
 * Retourne l'ID de la mission créée, ou null si aucun hashtag détecté.
 */
function creerMissionDepuisAlerte(ss, p, typeAlerte, idCible) {
  const msg = String(p.message || p.msg || "");
  if (!/#MMOK|#[Mm]ission/.test(msg)) return null;

  let sheet;
  try { sheet = getSheetSafe(ss, "MISSIONS"); }
  catch(e) {
    sheet = ss.insertSheet("MISSIONS");
    sheet.appendRow(["ID","Date","Heure","Email_Dest","Nom_Collab","CC","Objet","Corps","Statut","Notes","Date_Maj"]);
    sheet.setFrozenRows(1);
    sheet.getRange("1:1").setFontWeight("bold");
  }

  const emailDest = String(p.destinataire || "m.mokhtari@hecg.fr").trim().toLowerCase();
  let nomCollab = emailDest;
  try {
    const dPerso = getSheetSafe(ss, "Personnel").getDataRange().getValues();
    for (let i = 1; i < dPerso.length; i++) {
      if (String(dPerso[i][6] || "").trim().toLowerCase() === emailDest) {
        nomCollab = (String(dPerso[i][2] || "") + " " + String(dPerso[i][1] || "")).trim();
        break;
      }
    }
  } catch(eP) {}

  const now      = new Date();
  const dateStr  = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy");
  const heureStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm");
  const idMission = "MSN-" + now.getTime() + "-" + Math.floor(Math.random() * 1000);
  const texteObjet = msg.replace(/#MMOK|#[Mm]issions?/g, "").trim().substring(0, 60);
  const objet = typeAlerte + " [" + idCible + "]" + (texteObjet ? " – " + texteObjet : "");

  sheet.appendRow([idMission, dateStr, heureStr, emailDest, nomCollab, "", objet, msg, "Attribué", "", now, "", "", String(p.emailAuteur || "")]);
  logAction(p.idAuteur || "Système", p.roleAuteur || "Auto", "Création", "Missions", idCible,
    "Mission créée via alerte " + typeAlerte + " : « " + objet.substring(0, 80) + " »");
  return idMission;
}

// ==========================================
// IMPORT ONE-SHOT — HISTORIQUE NOTES
// ==========================================
/**
 * Migre les notes depuis IMPORTS_NOTES et IMPORTS_NOTESBTS vers NOTES_EXAMENS.
 * À exécuter UNE SEULE FOIS manuellement depuis l'éditeur Apps Script.
 * Ne pas ajouter à un trigger.
 */
function importerHistoriqueNotes() {
  const ss = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");

  // --- 1. Référentiel ETUDIANTS : Prénom (col C, idx 2) + Nom (col D, idx 3) → id_etu (col A, idx 0) ---
  const dataEtu = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues();
  const etuMap  = {};
  for (let i = 1; i < dataEtu.length; i++) {
    const id     = String(dataEtu[i][0] || "").trim();
    const prenom = String(dataEtu[i][2] || "").trim().toLowerCase().replace(/\s+/g, " ");
    const nom    = String(dataEtu[i][3] || "").trim().toLowerCase().replace(/\s+/g, " ");
    if (!id || (!prenom && !nom)) continue;
    const k1 = (prenom + " " + nom).trim();
    const k2 = (nom    + " " + prenom).trim();
    if (k1) etuMap[k1] = id;
    if (k2 && k2 !== k1) etuMap[k2] = id;
  }

  // --- 2. Référentiel_matieres : Diplome (col A, idx 0) + Matiere (col B, idx 1) → Coefficient (col D, idx 3) ---
  const dataRef = getSheetSafe(ss, "Referentiel_matieres").getDataRange().getValues();
  const refMap  = {};
  for (let i = 1; i < dataRef.length; i++) {
    const dip = String(dataRef[i][0] || "").trim().toLowerCase();
    const mat = String(dataRef[i][1] || "").trim().toLowerCase();
    if (!dip || !mat) continue;
    const coeffRaw = dataRef[i][3];
    refMap[dip + "|" + mat] = (coeffRaw !== "" && coeffRaw !== null && coeffRaw !== undefined)
      ? Number(coeffRaw)
      : null;
  }

  // --- 3. Feuille cible NOTES_EXAMENS ---
  let wsNotes;
  try {
    wsNotes = getSheetSafe(ss, "NOTES_EXAMENS");
  } catch(e) {
    wsNotes = ss.insertSheet("NOTES_EXAMENS");
    wsNotes.appendRow(["ID_Note","ID_Etudiant","Nom_Prenom","Classe","Annee_Obtention",
                       "Matiere","Note","Coefficient","Points","Resultat_Global",
                       "Sujet_Memoire","Referent_Memoire"]);
    wsNotes.setFrozenRows(1);
    wsNotes.getRange("1:1").setFontWeight("bold");
  }

  const existingRows = wsNotes.getLastRow() - 1; // hors en-tête
  if (existingRows > 0) {
    console.log("⚠️  NOTES_EXAMENS contient déjà " + existingRows + " ligne(s). L'import va AJOUTER à la suite.");
  }

  // --- 4. Dépivotage des onglets sources ---
  const SOURCES       = ["IMPORTS_NOTES", "IMPORTS_NOTESBTS"];
  const notFoundEtu   = new Set();
  const notFoundRef   = new Set();
  const rowsToInsert  = [];
  let   cpt           = 0;

  SOURCES.forEach(function(nomOnglet) {
    let wsImport;
    try {
      wsImport = getSheetSafe(ss, nomOnglet);
    } catch(e) {
      console.log("⚠️  Onglet introuvable : " + nomOnglet + " → ignoré.");
      return;
    }

    const data = wsImport.getDataRange().getValues();
    if (data.length < 2) {
      console.log("ℹ️  Onglet " + nomOnglet + " : aucune donnée.");
      return;
    }

    const headers    = data[0]; // Ligne 1 : index 0-2 = métadonnées, index 3+ = matières
    let   nbEtuTraites = 0;

    for (let i = 1; i < data.length; i++) {
      const row          = data[i];
      const nomPrenomRaw = String(row[0] || "").trim();
      const diplome      = String(row[1] || "").trim();
      const annee        = String(row[2] || "").trim();
      if (!nomPrenomRaw) continue;
      nbEtuTraites++;

      // Recherche de l'ID étudiant (essai clé "prénom nom" et "nom prénom")
      const searchKey = nomPrenomRaw.toLowerCase().replace(/\s+/g, " ");
      const idEtu     = etuMap[searchKey] || "";
      if (!idEtu) notFoundEtu.add(nomPrenomRaw);

      const diplomeKey = diplome.toLowerCase().trim();

      // Dépivotage : colonnes de notes à partir de l'index 3
      for (let j = 3; j < headers.length; j++) {
        const matiere = String(headers[j] || "").trim();
        if (!matiere) continue;

        const cellVal = row[j];
        if (cellVal === "" || cellVal === null || cellVal === undefined) continue;
        const note = parseFloat(cellVal);
        if (isNaN(note)) continue;

        // Recherche du coefficient
        const refKey  = diplomeKey + "|" + matiere.toLowerCase().trim();
        const hasCoef = (refKey in refMap);
        const coeff   = hasCoef ? refMap[refKey] : null;
        if (!hasCoef) notFoundRef.add(diplome + " | " + matiere);

        const points = (coeff !== null && coeff !== undefined)
          ? Math.round(note * coeff * 100) / 100
          : "";

        rowsToInsert.push([
          "NOT-" + Date.now() + "-" + (++cpt), // ID_Note unique
          idEtu,
          nomPrenomRaw,
          diplome,
          annee,
          matiere,
          note,
          (coeff !== null && coeff !== undefined) ? coeff : "",
          points,
          "", // Resultat_Global
          "", // Sujet_Memoire
          ""  // Referent_Memoire
        ]);
      }
    }
    console.log("✅ " + nomOnglet + " : " + nbEtuTraites + " étudiant(s) traité(s).");
  });

  // --- 5. Insertion en batch (1 seul appel API) ---
  if (rowsToInsert.length > 0) {
    wsNotes.getRange(wsNotes.getLastRow() + 1, 1, rowsToInsert.length, 12).setValues(rowsToInsert);
  }

  // --- 6. Rapport final dans la console ---
  console.log("==================================================");
  console.log("IMPORT TERMINÉ — " + rowsToInsert.length + " note(s) insérée(s) dans NOTES_EXAMENS");
  console.log("==================================================");

  if (notFoundEtu.size > 0) {
    console.log("❌ ÉTUDIANTS NON TROUVÉS dans ETUDIANTS (" + notFoundEtu.size + ") :");
    notFoundEtu.forEach(function(n) { console.log("   • " + n); });
    console.log("   → Vérifiez la casse / orthographe dans les onglets IMPORT (col A)");
    console.log("   → Ces lignes ont quand même été insérées avec ID_Etudiant vide.");
  } else {
    console.log("✅ Tous les étudiants ont été trouvés.");
  }

  if (notFoundRef.size > 0) {
    console.log("❌ MATIÈRES NON TROUVÉES dans Referentiel_matieres (" + notFoundRef.size + ") :");
    notFoundRef.forEach(function(m) { console.log("   • " + m); });
    console.log("   → Vérifiez le couple Diplome/Matiere dans Referentiel_matieres.");
    console.log("   → Ces notes ont été insérées avec Coefficient et Points vides.");
  } else {
    console.log("✅ Toutes les matières ont été trouvées.");
  }
  console.log("==================================================");
}

// ==========================================
// VIE SCOLAIRE — RÉFÉRENTIEL MATIÈRES
// ==========================================

function handleGetReferentielMatieres(p, ss) {
  try {
    const data = getSheetSafe(ss, "Referentiel_matieres").getDataRange().getValues();
    const result = {};
    for (let i = 1; i < data.length; i++) {
      const dip  = String(data[i][0] || "").trim();
      const mat  = String(data[i][1] || "").trim();
      const type = String(data[i][2] || "principale").trim();
      const coef = parseFloat(data[i][3]) || 1;
      if (!dip || !mat) continue;
      if (!result[dip]) result[dip] = [];
      result[dip].push({ matiere: mat, type: type, coeff: coef });
    }
    return createJsonResponse({ success: true, data: result });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// ==========================================
// VIE SCOLAIRE — NOTES EXAMENS
// ==========================================

function handleGetNotes(p, ss) {
  try {
    const data = getSheetSafe(ss, "NOTES_EXAMENS").getDataRange().getValues();
    const notes = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      if (p.idEtu    && String(row[1]).trim() !== String(p.idEtu).trim())   continue;
      if (p.classe   && String(row[3]).trim() !== p.classe)    continue;
      if (p.annee    && String(row[4]).trim() !== p.annee)     continue;
      // maxAnnee : pour DCG/DSCG cumulatif — inclut toutes les années <= maxAnnee
      if (p.maxAnnee && p.maxAnnee !== '' && String(row[4]).trim() > p.maxAnnee) continue;
      if (p.diplome  && String(row[3]) !== "" && !String(row[3]).toLowerCase().includes(p.diplome.toLowerCase())) continue;
      const rawNote = String(row[6]||"").trim();
      const noteVal = rawNote === "" ? null : (rawNote.toUpperCase() === "NE" ? "NE" : parseFloat(rawNote));
      const classeVal = String(row[3]);
      const effectiveClasse = (!classeVal.trim() && p.diplome) ? p.diplome : classeVal;
      notes.push({ id: String(row[0]), idEtu: String(row[1]), nomPrenom: String(row[2]),
        classe: effectiveClasse, annee: String(row[4]), matiere: String(row[5]),
        note: noteVal,
        coeff: row[7] !== "" ? parseFloat(row[7]) : null,
        points: (noteVal !== null && noteVal !== "NE" && row[8] !== "") ? parseFloat(row[8]) : null,
        resultat: String(row[9] || ""), sujetMemoire: String(row[10] || ""),
        referentMemoire: String(row[11] || "") });
    }
    return createJsonResponse({ success: true, data: notes });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleSaveNote(p, ss) {
  try {
    if (p.vieScRole !== "Super-admin" && p.vieScRole !== "Pédagogie")
      return createJsonResponse({ success: false, message: "Accès refusé." });
    const ws   = getSheetSafe(ss, "NOTES_EXAMENS");
    const data = ws.getDataRange().getValues();
    const rawN = (p.note !== "" && p.note !== null && p.note !== undefined) ? String(p.note).trim() : "";
    const isNE = rawN.toUpperCase() === "NE";
    const note  = rawN === "" ? null : (isNE ? "NE" : parseFloat(rawN));
    const coeff = (p.coeff !== "" && p.coeff !== null && p.coeff !== undefined) ? parseFloat(p.coeff) : null;
    const points = (note !== null && !isNE && coeff !== null) ? Math.round(parseFloat(note) * coeff * 100) / 100 : null;
    if (p.id) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) !== String(p.id)) continue;
        if (note  !== null) ws.getRange(i+1,7).setValue(note);
        if (coeff !== null) ws.getRange(i+1,8).setValue(coeff);
        if (points !== null) ws.getRange(i+1,9).setValue(points);
        else if (isNE) ws.getRange(i+1,9).setValue(0);
        if (p.sujetMemoire    !== undefined) ws.getRange(i+1,11).setValue(p.sujetMemoire || "");
        if (p.referentMemoire !== undefined) ws.getRange(i+1,12).setValue(p.referentMemoire || "");
        logAction(p.emailAuteur||"?", p.vieScRole||"?", "Modification", "Vie scolaire", p.id, "Note: "+p.matiere+" = "+note);
        return createJsonResponse({ success: true });
      }
      return createJsonResponse({ success: false, message: "Note introuvable." });
    }
    const id = "NOT-" + new Date().getTime() + "-" + Math.floor(Math.random()*1000);
    ws.appendRow([id, p.idEtu||"", p.nomPrenom||"", p.classe||"", p.annee||"", p.matiere||"",
      note ?? "", coeff ?? "", points ?? "", "", p.sujetMemoire||"", p.referentMemoire||""]);
    logAction(p.emailAuteur||"?", p.vieScRole||"?", "Création", "Vie scolaire", id,
      (p.nomPrenom||"?") + " — " + (p.matiere||"?") + " = " + note + "/20");
    return createJsonResponse({ success: true, id: id });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleDeleteNote(p, ss) {
  try {
    if (p.vieScRole !== "Super-admin") return createJsonResponse({ success: false, message: "Accès refusé." });
    const ws = getSheetSafe(ss, "NOTES_EXAMENS"); const data = ws.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(p.id)) {
        ws.deleteRow(i + 1);
        logAction(p.emailAuteur||"?", "Super-admin", "Suppression", "Vie scolaire", p.id, "Note supprimée");
        return createJsonResponse({ success: true });
      }
    }
    return createJsonResponse({ success: false, message: "Introuvable." });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleUpdateResultatNote(p, ss) {
  try {
    if (p.vieScRole !== "Super-admin" && p.vieScRole !== "Pédagogie")
      return createJsonResponse({ success: false, message: "Accès refusé." });
    const ws   = getSheetSafe(ss, "NOTES_EXAMENS");
    const data = ws.getDataRange().getValues();
    let ancienResultat = ""; let found = false;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) !== String(p.idEtu) || String(data[i][4]) !== String(p.annee)) continue;
      if (p.diplome && !String(data[i][3]).toUpperCase().includes(String(p.diplome).toUpperCase())) continue;
      if (!ancienResultat) ancienResultat = String(data[i][9] || "");
      ws.getRange(i+1, 10).setValue(p.resultat);
      found = true;
    }
    if (!found) return createJsonResponse({ success: false, message: "Aucune note trouvée." });
    // Historique
    let wsH;
    try { wsH = getSheetSafe(ss, "NOTES_HISTORIQUE_STATUT"); }
    catch(e) {
      wsH = ss.insertSheet("NOTES_HISTORIQUE_STATUT");
      wsH.appendRow(["ID","Date","Heure","ID_Etudiant","Nom_Prenom","Diplome","Annee","Ancien_Statut","Nouveau_Statut","Auteur"]);
      wsH.setFrozenRows(1); wsH.getRange("1:1").setFontWeight("bold");
    }
    const now = new Date();
    wsH.appendRow(["HST-"+now.getTime(),
      Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy"),
      Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm"),
      p.idEtu, p.nomPrenom||"", p.diplome||"", p.annee, ancienResultat, p.resultat, p.emailAuteur||"?"]);
    logAction(p.emailAuteur||"?", p.vieScRole||"?", "Modification", "Vie scolaire", p.idEtu,
      "Résultat: "+ancienResultat+" → "+p.resultat+" ("+p.diplome+" "+p.annee+")");
    return createJsonResponse({ success: true });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleGetHistoriqueStatutNotes(p, ss) {
  try {
    let wsH; try { wsH = getSheetSafe(ss, "NOTES_HISTORIQUE_STATUT"); } catch(e) { return createJsonResponse({ success: true, data: [] }); }
    const data = wsH.getDataRange().getValues(); const hist = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      if (p.idEtu && String(data[i][3]) !== String(p.idEtu)) continue;
      hist.push({ id: data[i][0], date: data[i][1], heure: data[i][2], idEtu: data[i][3],
        nomPrenom: data[i][4], diplome: data[i][5], annee: data[i][6],
        ancienStatut: data[i][7], nouveauStatut: data[i][8], auteur: data[i][9] });
    }
    return createJsonResponse({ success: true, data: hist });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleGetStatsNotes(p, ss) {
  try {
    const dataN = getSheetSafe(ss, "NOTES_EXAMENS").getDataRange().getValues();
    const dataR = getSheetSafe(ss, "Referentiel_matieres").getDataRange().getValues();
    const refMap = {};
    for (var i = 1; i < dataR.length; i++) {
      var rk   = String(dataR[i][0]||"").trim().toLowerCase() + "|" + String(dataR[i][1]||"").trim().toLowerCase();
      var type = String(dataR[i][2]||"principale").trim().toLowerCase();
      var coef = parseFloat(dataR[i][3]) || 1;
      refMap[rk] = { coeff: coef, type: type };
    }
    var diplome = (p.diplome||"").toLowerCase();
    var isDCG   = diplome.startsWith("dcg") && !diplome.startsWith("dscg");
    var isDSCG  = diplome.startsWith("dscg");
    var isBTS   = diplome.startsWith("bts");
    var annee   = String(p.annee||"").trim();

    // Collecte des notes
    // Pour DCG/DSCG : cumulatif (toutes années <= annee) + meilleure note par UE
    var byEtu = {};
    for (var i = 1; i < dataN.length; i++) {
      var row = dataN[i]; if (!row[0]) continue;
      var rowAnnee = String(row[4]).trim();
      if (p.diplome && !String(row[3]).toLowerCase().includes(diplome)) continue;
      if (p.classe  && String(row[3]).trim() !== p.classe) continue;
      // Filtre année : exact pour BTS, cumulatif pour DCG
      if (isBTS && annee && rowAnnee !== annee) continue;
      if ((isDCG||isDSCG) && annee && rowAnnee > annee) continue;
      var idEtu = String(row[1]); var mat = String(row[5]).trim();
      var rawNote = String(row[6]||"").trim();
      var note = rawNote === "" ? null : (rawNote.toUpperCase() === "NE" ? "NE" : parseFloat(rawNote));
      if (!byEtu[idEtu]) byEtu[idEtu] = { nomPrenom: row[2], classe: row[3], annee: rowAnnee, notesByUE: {}, resultat: "", lastAnnee: "" };
      if (rowAnnee >= byEtu[idEtu].lastAnnee) { byEtu[idEtu].classe = row[3]; byEtu[idEtu].annee = rowAnnee; byEtu[idEtu].lastAnnee = rowAnnee; }
      if (row[9]) byEtu[idEtu].resultat = String(row[9]);
      var storedCoeff = parseFloat(row[7]) || 1;
      // Pour DCG/DSCG : garder la meilleure note par UE (NE = dispensé, compte si pas de note numérique)
      // Clé normalisée : "UE 2 - Droit..." et "UE 2" → même clé pour éviter les doublons inter-années
      if (isDCG||isDSCG) {
        var kue = normMatiere(mat);
        var existing = byEtu[idEtu].notesByUE[kue];
        if (!existing) {
          byEtu[idEtu].notesByUE[kue] = { mat: mat, note: note, coeff: storedCoeff };
        } else if (note !== null && note !== "NE" && (existing.note === null || existing.note === "NE" || note > existing.note)) {
          byEtu[idEtu].notesByUE[kue] = { mat: mat, note: note, coeff: storedCoeff };
        }
        // Si déjà une note numérique on ne remplace pas par NE
      } else {
        byEtu[idEtu].notesByUE[mat.toLowerCase()] = { mat: mat, note: note, coeff: storedCoeff };
      }
    }

    var students = [];
    Object.keys(byEtu).forEach(function(idEtu) {
      var etu = byEtu[idEtu];
      var tp=0, tc=0, pp=0, cp=0, tpGen=0, tcGen=0;
      var tpOk=0, ueOk=0, ueTotal=0;
      var ueElim=0, ueRecu=0;
      Object.keys(etu.notesByUE).forEach(function(kue) {
        var entry = etu.notesByUE[kue]; var note = entry.note;
        if (note === null) return;
        // NE = dispensé par équivalence : compte comme UE validée, ne compte pas dans les points
        if (note === "NE") {
          ueTotal++; ueOk++;
          return;
        }
        if (isNaN(note)) return;
        var rk = diplome + "|" + kue; var ref = refMap[rk];
        if (!ref) {
          if (isDCG||isDSCG) ref = { coeff:1, type:"principale" };
          else if (isBTS) {
            // Matière BTS absente du référentiel : inclure dans la moyenne générale avec le coeff sauvegardé
            // coeff=0 → note informative (ex: Rattrapage), ne compte pas dans la moyenne
            var sc = entry.coeff;
            if (!sc) return;
            tp += note * sc; tc += sc; ueTotal++;
            return;
          } else return;
        }
        if (ref.type === "autre") return;
        tp += note * ref.coeff; tc += ref.coeff; ueTotal++;
        if (isBTS && /^E[5-8]\b/i.test(entry.mat.trim())) { pp += note * ref.coeff; cp += ref.coeff; }
        if (ref.type === "generale") { tpGen += note * ref.coeff; tcGen += ref.coeff; }
        if (isDCG||isDSCG) {
          if (note >= 10) ueOk++;
          else if (note >= 6) ueRecu++;
          else ueElim++;
        }
      });
      var moyGen = tc > 0 ? Math.round(tp/tc*100)/100 : null;
      var moyPro = cp > 0 ? Math.round(pp/cp*100)/100 : null;
      var moyGeneral = tcGen > 0 ? Math.round(tpGen/tcGen*100)/100 : null;
      // Différentiel DCG : total points obtenus - nb_UE * 10 (toutes UE, pas seulement les validées)
      var diff = null;
      if (isBTS && moyGen !== null) diff = Math.round((moyGen-10)*100)/100;
      else if ((isDCG||isDSCG) && ueTotal > 0) diff = Math.round((tp - ueTotal*10)*100)/100;
      var nbUEObl = isDCG ? 13 : (isDSCG ? 6 : 0);
      // Résultat automatique
      var resultatAuto = "";
      if (isBTS && moyGen !== null) {
        if (moyGen >= 10) resultatAuto = "Admis";
        else if (moyGen >= 8 && moyPro !== null && moyPro >= 10) resultatAuto = "Rattrapage";
        else resultatAuto = "Recalé";
      } else if (isDCG||isDSCG) {
        // DCG/DSCG : Admis si moy ≥ 10 ET aucune UE éliminatoire (note < 6)
        // Sinon Recalé dans tous les autres cas (moy < 10 ou au moins une UE éliminatoire)
        if (ueTotal >= nbUEObl) {
          resultatAuto = (moyGen !== null && moyGen >= 10 && ueElim === 0) ? "Admis" : "Recalé";
        } else {
          resultatAuto = "En cours";
        }
      }
      // Pour DCG stats : n'inclure que les étudiants ayant toutes leurs UE dans l'année sélectionnée
      if ((isDCG||isDSCG) && ueTotal < nbUEObl) return; // pas encore 13 UE
      if ((isDCG||isDSCG) && annee && etu.lastAnnee !== annee) return; // dernière saisie pas sur cette année
      students.push({ idEtu: idEtu, nomPrenom: etu.nomPrenom, classe: etu.classe, annee: etu.annee,
        moyGen: moyGen, moyPro: moyPro, moyGeneral: moyGeneral,
        totalPoints: Math.round(tp*100)/100, totalCoeff: tc,
        differentiels: diff, ueTotal: ueTotal, ueOk: ueOk, ueElim: ueElim, ueRecu: ueRecu,
        resultatAuto: resultatAuto, resultatManuel: etu.resultat, rang: 0 });
    });
    students.sort(function(a,b){ return (b.moyGen||0) - (a.moyGen||0); });
    students.forEach(function(s,i){ s.rang = i+1; });
    var wm = students.filter(function(s){ return s.moyGen !== null; });
    var nbT = wm.length;
    var getRes = function(s){ return s.resultatManuel || s.resultatAuto; };
    var nbA   = wm.filter(function(s){ return getRes(s)==="Admis"; }).length;
    var nbR   = wm.filter(function(s){ return getRes(s)==="Rattrapage"; }).length;
    var nbRec = wm.filter(function(s){ return ["Recalé","Non admis"].includes(getRes(s)); }).length;
    var moys  = wm.map(function(s){ return s.moyGen; });
    var moy   = nbT > 0 ? moys.reduce(function(a,b){ return a+b; },0)/nbT : 0;
    var ec    = nbT > 0 ? Math.sqrt(moys.reduce(function(a,b){ return a+Math.pow(b-moy,2); },0)/nbT) : 0;
    var ns10  = moys.filter(function(n){ return n>=10; }).length;
    return createJsonResponse({ success: true, students: students, promo: {
      nbTotal: nbT, nbAdmis: nbA, nbRattrapage: nbR, nbRecale: nbRec,
      moyenne: Math.round(moy*100)/100, noteMax: moys.length?Math.max.apply(null,moys):0,
      noteMin: moys.length?Math.min.apply(null,moys):0, ecartType: Math.round(ec*100)/100,
      nbSup10: ns10, nbInf10: nbT-ns10,
      pcSup10: nbT>0?Math.round(ns10/nbT*100):0, tauxReussite: nbT>0?Math.round(nbA/nbT*100):0 }});
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// ==========================================
// VIE SCOLAIRE — CLASSES INFO
// ==========================================

function handleGetClassesInfo(p, ss) {
  try {
    const data = getSheetSafe(ss, "CLASSES_INFO").getDataRange().getValues(); const rows = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      rows.push({ classe: data[i][0], pp: data[i][1], del1: data[i][2], del2: data[i][3],
        sup1: data[i][4], sup2: data[i][5], arretS1: data[i][6], conseilS1: data[i][7],
        arretS2: data[i][8], conseilS2: data[i][9],
        heureConseilS1: String(data[i][10]||""), heureConseilS2: String(data[i][11]||"") });
    }
    return createJsonResponse({ success: true, data: rows });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleSaveClasseInfo(p, ss) {
  try {
    const ws   = getSheetSafe(ss, "CLASSES_INFO");
    const data = ws.getDataRange().getValues();
    // Validation : délégués et suppléants doivent être dans ETUDIANTS
    const dEtu = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues();
    function nomConnu(nom) {
      if (!nom || !nom.trim()) return true;
      const nl = nom.toLowerCase().trim();
      for (let i = 1; i < dEtu.length; i++) {
        const k1 = (String(dEtu[i][2]||"")+" "+String(dEtu[i][3]||"")).toLowerCase().trim();
        const k2 = (String(dEtu[i][3]||"")+" "+String(dEtu[i][2]||"")).toLowerCase().trim();
        if (nl===k1 || nl===k2) return true;
      }
      return false;
    }
    for (const n of [p.del1, p.del2, p.sup1, p.sup2]) {
      if (n && n.trim() && !nomConnu(n))
        return createJsonResponse({ success: false, message: "Nom non reconnu dans ETUDIANTS : « "+n+" »" });
    }
    const rowVal = [p.classe||"", p.pp||"", p.del1||"", p.del2||"", p.sup1||"", p.sup2||"",
                    p.arretS1||"", p.conseilS1||"", p.arretS2||"", p.conseilS2||"",
                    p.heureConseilS1||"", p.heureConseilS2||""];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === (p.classe||"").trim().toLowerCase()) {
        ws.getRange(i+1, 1, 1, 12).setValues([rowVal]);
        logAction(p.emailAuteur||"?", p.vieScRole||"?", "Modification", "Vie scolaire", p.classe, "Infos classe: "+p.classe);
        return createJsonResponse({ success: true });
      }
    }
    ws.appendRow(rowVal);
    logAction(p.emailAuteur||"?", p.vieScRole||"?", "Création", "Vie scolaire", p.classe, "Nouvelle classe: "+p.classe);
    return createJsonResponse({ success: true });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// VIE SCOLAIRE — PERSONNEL ENCADRANT (onglet Adultes)
// Colonnes : A=ID B=Nom C=Prenom D=Mail_Perso E=Mail_Pro F=Adresse G=CP H=Ville I=Matieres J=Date_entree K=Date_sortie L=Tel M=Etat_formateur
// ==========================================

function handleGetPersonnel(p, ss) {
  try {
    var ws; try { ws = getSheetSafe(ss, "Adultes"); } catch(e) { ws = getSheetSafe(ss, "FORMATEURS"); }
    var data = ws.getDataRange().getValues(); var rows = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var dateE = data[i][9]; var dateS = data[i][10];
      rows.push({
        id:              String(data[i][0]),
        nom:             String(data[i][1]||""),
        prenom:          String(data[i][2]||""),
        mailPerso:       String(data[i][3]||""),
        mailPro:         String(data[i][4]||""),
        adresse:         String(data[i][5]||""),
        cp:              String(data[i][6]||""),
        ville:           String(data[i][7]||""),
        matieres:        String(data[i][8]||""),
        dateEntree:      dateE instanceof Date ? Utilities.formatDate(dateE,Session.getScriptTimeZone(),"dd/MM/yyyy") : String(dateE||""),
        dateSortie:      dateS instanceof Date ? Utilities.formatDate(dateS,Session.getScriptTimeZone(),"dd/MM/yyyy") : String(dateS||""),
        tel:             String(data[i][11]||""),
        etatFormateur:   String(data[i][12]||"NON").trim().toUpperCase(),
      });
    }
    return createJsonResponse({ success: true, data: rows });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleSavePersonnel(p, ss) {
  try {
    if (p.vieScRole !== "Super-admin") return createJsonResponse({ success:false, message:"Accès refusé." });
    var ws; try { ws = getSheetSafe(ss, "Adultes"); } catch(e) { ws = getSheetSafe(ss, "FORMATEURS"); }
    var data = ws.getDataRange().getValues();
    var row = [p.id||"", p.nom||"", p.prenom||"", p.mailPerso||"", p.mailPro||"",
               p.adresse||"", p.cp||"", p.ville||"", p.matieres||"",
               p.dateEntree||"", p.dateSortie||"", p.tel||"", (p.etatFormateur||"NON").toUpperCase()];
    if (p.id) {
      for (var i=1;i<data.length;i++) {
        if (String(data[i][0])===String(p.id)) {
          ws.getRange(i+1,1,1,13).setValues([row]);
          logAction(p.emailAuteur||"?",p.vieScRole||"?","Modification","Vie scolaire",p.id,"Personnel: "+p.nom+" "+p.prenom);
          return createJsonResponse({ success:true });
        }
      }
    }
    var id = "ADU-"+new Date().getTime().toString().slice(-7);
    row[0]=id;
    ws.appendRow(row);
    logAction(p.emailAuteur||"?",p.vieScRole||"?","Création","Vie scolaire",id,"Personnel: "+p.nom+" "+p.prenom);
    return createJsonResponse({ success:true, id:id });
  } catch(e) { return createJsonResponse({ success:false, message:e.toString() }); }
}

function handleDeletePersonnel(p, ss) {
  try {
    if (p.vieScRole !== "Super-admin") return createJsonResponse({ success:false, message:"Accès refusé." });
    var ws; try { ws = getSheetSafe(ss, "Adultes"); } catch(e) { ws = getSheetSafe(ss, "FORMATEURS"); }
    var data = ws.getDataRange().getValues();
    for (var i=1;i<data.length;i++) {
      if (String(data[i][0])===String(p.id)) {
        ws.deleteRow(i+1);
        logAction(p.emailAuteur||"?",p.vieScRole||"?","Suppression","Vie scolaire",p.id,"Personnel supprimé");
        return createJsonResponse({ success:true });
      }
    }
    return createJsonResponse({ success:false, message:"Introuvable." });
  } catch(e) { return createJsonResponse({ success:false, message:e.toString() }); }
}

// Kept for backward compatibility
function handleGetFormateurs(p, ss) { return handleGetPersonnel(p, ss); }
function handleSaveFormateur(p, ss) { return handleSavePersonnel(p, ss); }
function handleDeleteFormateur(p, ss) { return handleDeletePersonnel(p, ss); }

// ==========================================
// VIE SCOLAIRE — PLANNING
// ==========================================

function handleGetPlanningTaches(p, ss) {
  try {
    const data = getSheetSafe(ss, "PLANNING").getDataRange().getValues();
    const isAdmin = p.vieScRole === "Super-admin";
    const userRef   = String(p.userRef||"").trim();
    const userEmail = String(p.emailAuteur||"").toLowerCase();
    // Live priority lookup from REFTACHE
    let refMap = {};
    try {
      const rd = getSheetSafe(ss, "REFTACHE").getDataRange().getValues();
      for (let j = 1; j < rd.length; j++) {
        if (rd[j][0]) refMap[String(rd[j][0])] = { nom: String(rd[j][1]||""), priorite: Number(rd[j][2]||1) };
      }
    } catch(eR) {}
    const tasks = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      if (!isAdmin) {
        const tRef   = String(data[i][8]||"").trim();
        const tEmail = String(data[i][4]||"").trim().toLowerCase();
        const match  = (tRef && tRef === userRef) || (!tRef && tEmail === userEmail);
        if (!match) continue;
      }
      const tacheRef = String(data[i][10]||"");
      const refEntry = refMap[tacheRef] || {};
      tasks.push({ id: String(data[i][0]), titre: data[i][1],
        dateDebut: fmtSheetDate(data[i][2]), dateFin: fmtSheetDate(data[i][3]),
        referent: data[i][4], statut: data[i][5],
        description: data[i][6], demiJournee: String(data[i][7]||"Matin"),
        referentRef: String(data[i][8]||""), referentNom: String(data[i][9]||""),
        tacheRef: tacheRef, priorite: refEntry.priorite || Number(data[i][11]||1),
        missionId: String(data[i][12]||"")
      });
    }
    return createJsonResponse({ success: true, data: tasks });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleSavePlanningTache(p, ss) {
  try {
    const ws = getSheetSafe(ss, "PLANNING"); const data = ws.getDataRange().getValues();
    const isAdmin = p.vieScRole === "Super-admin";
    const row = [
      p.id||"", p.titre||"", p.dateDebut||"", p.dateFin||"",
      p.referentEmail||p.referent||"", p.statut||"À faire",
      p.description||"", p.demiJournee||"Matin",
      p.referentRef||"", p.referentNom||"",
      p.tacheRef||"", p.priorite||1
    ];
    if (p.id) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) !== String(p.id)) continue;
        const existingMissionId = String(data[i][12]||"");
        const rowWithMission = [
          p.id||"", p.titre||"", p.dateDebut||"", p.dateFin||"",
          p.referentEmail||p.referent||"", p.statut||"À faire",
          p.description||"", p.demiJournee||"Matin",
          p.referentRef||"", p.referentNom||"",
          p.tacheRef||"", p.priorite||1, existingMissionId
        ];
        ws.getRange(i+1,1,1,13).setValues([rowWithMission]);
        logAction(p.emailAuteur||"?", p.vieScRole||"?", "Modification", "Vie scolaire", p.id, "Tâche: "+p.titre);
        // Sync vers mission : si tâche Terminée/Annulée → clôturer la mission liée
        if (existingMissionId && (p.statut === "Terminé" || p.statut === "Annulé")) {
          try {
            var mSyncWs = getSheetSafe(ss, "MISSIONS"); var mSyncData = mSyncWs.getDataRange().getValues();
            for (var ms = 1; ms < mSyncData.length; ms++) {
              if (String(mSyncData[ms][0]) === existingMissionId) {
                mSyncWs.getRange(ms+1, 9).setValue(p.statut === "Terminé" ? "Clôturé" : "Annulé");
                mSyncWs.getRange(ms+1, 11).setValue(new Date()); break;
              }
            }
          } catch(eS) {}
        }
        if (p.referentEmail && shouldSendEmailToReferent(p.referentEmail)) {
          try { GmailApp.sendEmail(p.referentEmail, "[HECG] Tâche mise à jour : "+p.titre,
            "Bonjour "+(p.referentNom||"")+",\n\nVotre tâche «"+p.titre+"» a été mise à jour.\n\nStatut : "+p.statut+"\nDate : "+p.dateDebut+" ("+p.demiJournee+")\n\nCordialement,\nL'équipe HECG",
            { name: "CRM HECG" }); } catch(eM) {}
        }
        return createJsonResponse({ success: true });
      }
    }
    if (!isAdmin) return createJsonResponse({ success: false, message: "Seul le super-admin peut créer des tâches." });
    const id = "TSK-" + new Date().getTime().toString().slice(-7); row[0] = id;
    ws.appendRow([...row, ""]); // col M (missionId) vide au départ
    logAction(p.emailAuteur||"?", p.vieScRole||"?", "Création", "Vie scolaire", id, "Tâche: "+p.titre);

    // Créer automatiquement une mission pour la personne assignée
    var missionId = null;
    var emailDest = String(p.referentEmail || p.referent || "").trim().toLowerCase();
    if (emailDest) {
      try {
        var mSheet = getSheetSafe(ss, "MISSIONS");
        var mNow   = new Date();
        missionId  = "M" + mNow.getTime();
        mSheet.appendRow([
          missionId,
          Utilities.formatDate(mNow, Session.getScriptTimeZone(), "dd/MM/yyyy"),
          Utilities.formatDate(mNow, Session.getScriptTimeZone(), "HH:mm"),
          emailDest,
          String(p.referentNom || ""),
          "",
          String(p.titre || ""),
          String(p.description || ""),
          "Attribué",
          "",
          mNow,
          Number(p.priorite || 3),
          "[Planning] " + String(p.titre || ""),
          String(p.emailAuteur || "")
        ]);
        // Stocker missionId dans col M de la tâche planning
        ws.getRange(ws.getLastRow(), 13).setValue(missionId);
        logAction(p.emailAuteur||"?", p.vieScRole||"?", "Création", "Missions", missionId,
          "Mission auto (planning) : " + String(p.titre||"").substring(0,80) + " | Pour : " + String(p.referentNom||emailDest));
      } catch(eMission) {}
    }

    if (p.referentEmail && shouldSendEmailToReferent(p.referentEmail)) {
      try { GmailApp.sendEmail(p.referentEmail, "[HECG] Nouvelle tâche assignée : "+p.titre,
        "Bonjour "+(p.referentNom||"")+",\n\nUne nouvelle tâche vous a été assignée.\n\nTâche : "+p.titre+"\nDate : "+p.dateDebut+" ("+p.demiJournee+")\nStatut : "+(p.statut||"À faire")+"\n\nCordialement,\nL'équipe HECG",
        { name: "CRM HECG" }); } catch(eM) {}
    }
    return createJsonResponse({ success: true, id: id, missionId: missionId });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleDeletePlanningTache(p, ss) {
  try {
    if (p.vieScRole !== "Super-admin") return createJsonResponse({ success: false, message: "Accès refusé." });
    const ws = getSheetSafe(ss, "PLANNING"); const data = ws.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(p.id)) {
        const titre = String(data[i][1]||"");
        ws.deleteRow(i+1);
        logAction(p.emailAuteur||"?", p.vieScRole||"?", "Suppression", "Vie scolaire", p.id, "Tâche supprimée: "+titre);
        return createJsonResponse({ success: true });
      }
    }
    return createJsonResponse({ success: false, message: "Introuvable." });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// ==========================================
// VIE SCOLAIRE — REFTACHE
// ==========================================

function handleGetRefTaches(p, ss) {
  try {
    const data = getSheetSafe(ss, "REFTACHE").getDataRange().getValues(); const list = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      list.push({ id: String(data[i][0]), nom: String(data[i][1]||""), priorite: Number(data[i][2]||1) });
    }
    return createJsonResponse({ success: true, data: list });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleSaveRefTache(p, ss) {
  try {
    if (p.vieScRole !== "Super-admin") return createJsonResponse({ success: false, message: "Accès refusé." });
    const ws = getSheetSafe(ss, "REFTACHE"); const data = ws.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(p.id)) {
        ws.getRange(i+1,3).setValue(Number(p.priorite));
        logAction(p.emailAuteur||"?", p.vieScRole||"?", "Modification", "Vie scolaire", p.id, "Priorité tâche → "+p.priorite);
        return createJsonResponse({ success: true });
      }
    }
    return createJsonResponse({ success: false, message: "Tâche introuvable." });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// ══════════════════════════════════════════════════════════════════
// OBJECTIFS HEBDOMADAIRES
// Onglet OBJECTIFS_SEM : A=ID B=SemaineDebut C=Nom D=Quantite E=Priorite F=Descriptif G=EtatMiSemaine H=EtatFinSemaine I=Valide J=CreePar K=DateCre L=RetourEquipes
// ══════════════════════════════════════════════════════════════════

function getOrCreateObjectifsSheet(ss) {
  var ws = ss.getSheetByName("OBJECTIFS_SEM");
  if (!ws) {
    ws = ss.insertSheet("OBJECTIFS_SEM");
    ws.getRange(1,1,1,12).setValues([["ID","SemaineDebut","Nom","Quantite","Priorite","Descriptif","EtatMiSemaine","EtatFinSemaine","Valide","CreePar","DateCre","RetourEquipes"]]);
    ws.setFrozenRows(1);
  }
  return ws;
}

function normValide(v) {
  return (v === true || String(v).toUpperCase() === 'TRUE') ? 'TRUE' : 'FALSE';
}

function fmtSheetDate(val) {
  if (!val) return "";
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return String(val);
}

function handleGetObjectifsSemaine(p, ss) {
  try {
    var ws = getOrCreateObjectifsSheet(ss);
    var data = ws.getDataRange().getValues();
    var sem = p.semaineDebut || "";
    var dateFin = p.dateFin || "";
    var result = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var cellSem = fmtSheetDate(data[i][1]);
      if (sem) {
        if (dateFin) { if (cellSem < sem || cellSem > dateFin) continue; }
        else { if (cellSem !== sem) continue; }
      }
      result.push({
        id:            String(data[i][0]),
        semaineDebut:  cellSem,
        nom:           String(data[i][2]),
        quantite:      String(data[i][3]),
        priorite:      Number(data[i][4]) || 1,
        descriptif:    String(data[i][5]),
        etatMiSemaine: String(data[i][6]||""),
        etatFinSemaine:String(data[i][7]||""),
        valide:        normValide(data[i][8]),
        creePar:       String(data[i][9]||""),
        dateCre:       String(data[i][10]||""),
        retourEquipes: String(data[i][11]||""),
      });
    }
    return createJsonResponse({ success: true, data: result });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleSaveObjectifSemaine(p, ss) {
  try {
    if (p.vieScRole !== "Super-admin") return createJsonResponse({ success: false, message: "Accès refusé." });
    const ws = getOrCreateObjectifsSheet(ss);
    const data = ws.getDataRange().getValues();
    if (p.id) {
      // Mise à jour
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]) !== String(p.id)) continue;
        ws.getRange(i+1,3,1,7).setValues([[p.nom, p.quantite||"", Number(p.priorite)||1, p.descriptif||"", data[i][6], data[i][7], data[i][8]]]);
        logAction(p.emailAuteur||"?", p.vieScRole||p.roleAuteur||"?", "Modification", "Vie scolaire", p.id, "Objectif: "+p.nom);
        return createJsonResponse({ success: true });
      }
      return createJsonResponse({ success: false, message: "Objectif introuvable." });
    } else {
      // Création — stocker semaineDebut comme texte (apostrophe) pour éviter la conversion en Date
      var id = "OBJ_" + new Date().getTime();
      var newRow = [id, p.semaineDebut||"", p.nom, p.quantite||"", Number(p.priorite)||1, p.descriptif||"", "", "", "FALSE", p.emailAuteur||"?", new Date().toISOString(), ""];
      ws.appendRow(newRow);
      // Forcer colonne B en format texte pour éviter que Sheets convertisse en Date
      var lastRow = ws.getLastRow();
      ws.getRange(lastRow, 2).setNumberFormat("@");
      logAction(p.emailAuteur||"?", p.vieScRole||p.roleAuteur||"?", "Création", "Vie scolaire", id, "Objectif: "+p.nom);
      return createJsonResponse({ success: true, id: id });
    }
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleDeleteObjectifSemaine(p, ss) {
  try {
    if (p.vieScRole !== "Super-admin") return createJsonResponse({ success: false, message: "Accès refusé." });
    const ws = getOrCreateObjectifsSheet(ss);
    const data = ws.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(p.id)) {
        const nomObj = String(data[i][2]||"");
        ws.deleteRow(i+1);
        logAction(p.emailAuteur||"?", p.vieScRole||p.roleAuteur||"?", "Suppression", "Vie scolaire", p.id, "Objectif supprimé: "+nomObj);
        return createJsonResponse({ success: true });
      }
    }
    return createJsonResponse({ success: false, message: "Objectif introuvable." });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleUpdateObjectifEtat(p, ss) {
  try {
    const isSA = p.roleAuteur === 'Super-admin';
    const ws = getOrCreateObjectifsSheet(ss);
    const data = ws.getDataRange().getValues();

    // Append-only mode for retourEquipes (non-SA users)
    if (p.field === 'appendRetourEquipes') {
      const newEntry = String(p.val || "").trim();
      if (!newEntry) return createJsonResponse({ success: false, message: "Entrée vide." });
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(p.id)) {
          const existing = String(data[i][11] || "").trim();
          const newVal = existing ? existing + "\n" + newEntry : newEntry;
          ws.getRange(i+1, 12).setValue(newVal);
          logAction(p.emailAuteur||"?", p.roleAuteur||"?", "Modification", "Vie scolaire", p.id, "Retour équipe objectif: "+String(data[i][2]||""));
          return createJsonResponse({ success: true });
        }
      }
      return createJsonResponse({ success: false, message: "Objectif introuvable." });
    }

    // Block non-SA from modifying states and validation
    if (!isSA && (p.field === 'etatMiSemaine' || p.field === 'etatFinSemaine' || p.field === 'valide')) {
      return createJsonResponse({ success: false, message: "Accès refusé." });
    }

    const colMap = { etatMiSemaine: 7, etatFinSemaine: 8, valide: 9, retourEquipes: 12 };
    const col = colMap[p.field];
    if (!col) return createJsonResponse({ success: false, message: "Champ inconnu : " + p.field });
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(p.id)) {
        ws.getRange(i+1, col).setValue(p.val);
        logAction(p.emailAuteur||"?", p.roleAuteur||"?", "Modification", "Vie scolaire", p.id, "Objectif "+p.field+": "+p.val+" — "+String(data[i][2]||""));
        return createJsonResponse({ success: true });
      }
    }
    return createJsonResponse({ success: false, message: "Objectif introuvable." });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleGetObjectifsStats(p, ss) {
  try {
    var ws = getOrCreateObjectifsSheet(ss);
    var data = ws.getDataRange().getValues();
    var byMonth = {}; // clé = "YYYY-MM"
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var sem = fmtSheetDate(data[i][1]); // "2026-05-18"
      var month = sem.slice(0,7); // "2026-05"
      if (!month) continue;
      if (!byMonth[month]) byMonth[month] = { total:0, ok:0, nok:0, enCours:0, detail:[] };
      var valide = normValide(data[i][8]);
      var etatFin = String(data[i][7]||"");
      byMonth[month].total++;
      if (valide==="TRUE") byMonth[month].ok++;
      else if (etatFin) byMonth[month].nok++;
      else byMonth[month].enCours++;
      byMonth[month].detail.push({
        id:            String(data[i][0]),
        semaineDebut:  sem,
        nom:           String(data[i][2]),
        quantite:      String(data[i][3]||""),
        priorite:      Number(data[i][4])||1,
        etatMiSemaine: String(data[i][6]||""),
        etatFinSemaine:etatFin,
        valide:        valide,
        retourEquipes: String(data[i][11]||""),
      });
    }
    var months = Object.keys(byMonth).sort();
    var totalAll=0, totalOk=0, totalNok=0, totalEnCours=0;
    months.forEach(function(m){ totalAll+=byMonth[m].total; totalOk+=byMonth[m].ok; totalNok+=byMonth[m].nok; totalEnCours+=byMonth[m].enCours; });
    return createJsonResponse({ success:true, byMonth:byMonth, months:months, totaux:{total:totalAll,ok:totalOk,nok:totalNok,enCours:totalEnCours} });
  } catch(e) { return createJsonResponse({ success:false, message:e.toString() }); }
}

function handleGetObjectifsStatsSemaines(p, ss) {
  try {
    var ws = getOrCreateObjectifsSheet(ss);
    var data = ws.getDataRange().getValues();
    var bySem = {};
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var sem = fmtSheetDate(data[i][1]);
      if (!sem) continue;
      if (!bySem[sem]) bySem[sem] = { total:0, ok:0, nok:0, enCours:0, detail:[] };
      var valide = normValide(data[i][8]);
      var etatFin = String(data[i][7]||"");
      bySem[sem].total++;
      if (valide==="TRUE") bySem[sem].ok++;
      else if (etatFin) bySem[sem].nok++;
      else bySem[sem].enCours++;
      bySem[sem].detail.push({
        nom:           String(data[i][2]),
        quantite:      String(data[i][3]||""),
        priorite:      Number(data[i][4])||1,
        etatFinSemaine:etatFin,
        retourEquipes: String(data[i][11]||""),
        valide:        valide,
      });
    }
    var semaines = Object.keys(bySem).sort();
    return createJsonResponse({ success:true, bySem:bySem, semaines:semaines });
  } catch(e) { return createJsonResponse({ success:false, message:e.toString() }); }
}

// ══════════════════════════════════════════════════════════════════
// SYNTHÈSE HEBDOMADAIRE
// Onglet SYNTH_SEM : A=SemaineDebut B=PointsForts C=AxesAmelioration D=RapportSemaine E=BulletinSFLIFLA F=DateModif
// ══════════════════════════════════════════════════════════════════

function getOrCreateSynthSheet(ss) {
  var ws = ss.getSheetByName("SYNTH_SEM");
  if (!ws) {
    ws = ss.insertSheet("SYNTH_SEM");
    ws.getRange(1,1,1,7).setValues([["SemaineDebut","PointsForts","AxesAmelioration","RapportSemaine","BulletinSFLIFLA","DateModif","Commentaires"]]);
    ws.setFrozenRows(1);
  }
  return ws;
}

function handleGetSyntheseSemaine(p, ss) {
  try {
    var ws = getOrCreateSynthSheet(ss);
    var data = ws.getDataRange().getValues();
    var sem = p.semaineDebut || "";
    var dateDebut = p.dateDebut || sem;
    var dateFin   = p.dateFin   || sem;
    var result = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var cellSem = fmtSheetDate(data[i][0]);
      if (dateDebut && dateFin) {
        if (cellSem < dateDebut || cellSem > dateFin) continue;
      } else if (sem) {
        if (cellSem !== sem) continue;
      }
      result.push({
        semaineDebut:    cellSem,
        pointsForts:     String(data[i][1]||""),
        axesAmelioration:String(data[i][2]||""),
        rapportSemaine:  String(data[i][3]||""),
        bulletinSFLIFLA: String(data[i][4]||""),
        commentaires:    String(data[i][6]||""),
      });
    }
    return createJsonResponse({ success:true, data:result });
  } catch(e) { return createJsonResponse({ success:false, message:e.toString() }); }
}

function handleSaveSyntheseSemaine(p, ss) {
  try {
    var isSA = p.roleAuteur === 'Super-admin';
    var ws = getOrCreateSynthSheet(ss);
    var data = ws.getDataRange().getValues();
    var sem = p.semaineDebut || "";
    if (!sem) return createJsonResponse({ success:false, message:"semaineDebut requis." });

    // Append-only comment (accessible à tous)
    if (p.field === 'appendComment') {
      var newEntry = String(p.val||"").trim();
      if (!newEntry) return createJsonResponse({ success:false, message:"Commentaire vide." });
      for (var i = 1; i < data.length; i++) {
        if (fmtSheetDate(data[i][0]) === sem) {
          var existing = String(data[i][6]||"").trim();
          ws.getRange(i+1, 7).setValue(existing ? existing+"\n"+newEntry : newEntry);
          logAction(p.emailAuteur||"?", p.roleAuteur||"?", "Modification", "Vie scolaire", sem, "Commentaire synthèse semaine");
          return createJsonResponse({ success:true });
        }
      }
      // Ligne inexistante : créer avec seulement la semaine + commentaire
      ws.appendRow([sem, "", "", "", "", new Date().toISOString(), newEntry]);
      ws.getRange(ws.getLastRow(), 1).setNumberFormat("@");
      logAction(p.emailAuteur||"?", p.roleAuteur||"?", "Création", "Vie scolaire", sem, "Commentaire synthèse semaine");
      return createJsonResponse({ success:true });
    }

    // Modification principale — réservée Super-admin
    if (!isSA) return createJsonResponse({ success:false, message:"Accès refusé." });

    var fieldCol = { pointsForts:2, axesAmelioration:3, rapportSemaine:4, bulletinSFLIFLA:5 };
    for (var i = 1; i < data.length; i++) {
      if (fmtSheetDate(data[i][0]) === sem) {
        if (p.field) {
          var col = fieldCol[p.field];
          if (!col) return createJsonResponse({ success:false, message:"Champ inconnu." });
          ws.getRange(i+1, col).setValue(p.val||"");
          ws.getRange(i+1, 6).setValue(new Date().toISOString());
        } else {
          ws.getRange(i+1,2,1,5).setValues([[p.pointsForts||"",p.axesAmelioration||"",p.rapportSemaine||"",p.bulletinSFLIFLA||"",new Date().toISOString()]]);
        }
        logAction(p.emailAuteur||"?", p.roleAuteur||"?", "Modification", "Vie scolaire", sem, "Synthèse semaine mise à jour");
        return createJsonResponse({ success:true });
      }
    }
    ws.appendRow([sem, p.pointsForts||"", p.axesAmelioration||"", p.rapportSemaine||"", p.bulletinSFLIFLA||"", new Date().toISOString(), ""]);
    ws.getRange(ws.getLastRow(), 1).setNumberFormat("@");
    logAction(p.emailAuteur||"?", p.roleAuteur||"?", "Création", "Vie scolaire", sem, "Synthèse semaine créée");
    return createJsonResponse({ success:true });
  } catch(e) { return createJsonResponse({ success:false, message:e.toString() }); }
}

function handleGetStatsMatieres(p, ss) {
  try {
    var dataN = getSheetSafe(ss, "NOTES_EXAMENS").getDataRange().getValues();
    var diplome = (p.diplome||"").toLowerCase();
    var byMat = {}; // matiere -> { annees: { annee -> {total,>=10,6-9,<6,notes:[]} } }
    for (var i=1;i<dataN.length;i++) {
      var row = dataN[i]; if(!row[0]) continue;
      if(p.diplome && !String(row[3]).toLowerCase().includes(diplome)) continue;
      var mat   = String(row[5]).trim(); if(!mat) continue;
      var annee = String(row[4]).trim();
      var note  = row[6]!==""?parseFloat(row[6]):null;
      if(note===null||isNaN(note)) continue;
      if(!byMat[mat]) byMat[mat]={};
      if(!byMat[mat][annee]) byMat[mat][annee]={total:0,ok:0,recu:0,elim:0,somme:0,notes:[]};
      byMat[mat][annee].total++; byMat[mat][annee].somme+=note; byMat[mat][annee].notes.push(note);
      if(note>=10) byMat[mat][annee].ok++;
      else if(note>=6) byMat[mat][annee].recu++;
      else byMat[mat][annee].elim++;
    }
    // Résumé par matière toutes années
    var result = Object.keys(byMat).sort().map(function(mat){
      var annees = Object.keys(byMat[mat]).sort();
      var total=0,ok=0,recu=0,elim=0,somme=0;
      annees.forEach(function(a){ var d=byMat[mat][a]; total+=d.total; ok+=d.ok; recu+=d.recu; elim+=d.elim; somme+=d.somme; });
      return { mat:mat, total:total, ok:ok, recu:recu, elim:elim,
        moy: total>0?Math.round(somme/total*100)/100:null,
        tauxOk: total>0?Math.round(ok/total*100):0,
        tauxRecu: total>0?Math.round(recu/total*100):0,
        tauxElim: total>0?Math.round(elim/total*100):0,
        parAnnee: annees.map(function(a){ var d=byMat[mat][a];
          return { annee:a, total:d.total, ok:d.ok, recu:d.recu, elim:d.elim,
            moy:d.total>0?Math.round(d.somme/d.total*100)/100:null,
            tauxOk:d.total>0?Math.round(d.ok/d.total*100):0 }; })
      };
    });
    return createJsonResponse({ success:true, data:result });
  } catch(e) { return createJsonResponse({ success:false, message:e.toString() }); }
}

function shouldSendEmailToReferent(email) {
  if (!email || email.indexOf('@') < 0) return false;
  try {
    const store = PropertiesService.getScriptProperties();
    const key   = "PLN_MAIL_" + email.replace(/[^a-zA-Z0-9]/g,"_");
    const last  = store.getProperty(key);
    const now   = new Date().getTime();
    if (!last || (now - parseInt(last)) > 600000) {
      store.setProperty(key, now.toString());
      return true;
    }
    return false;
  } catch(e) { return false; }
}

// ==========================================
// VIE SCOLAIRE — ASPIRATEUR INSCRIPTIONS
// ==========================================

function extractInscriptionFields(body) {
  var b = body.replace(/&nbsp;/g,' ');
  function ext(field) {
    var rx = new RegExp(field.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')+'[:\\s]*([^<\n\r]+)','i');
    var m = b.match(rx); return m ? m[1].replace(/<[^>]+>/g,'').trim() : '';
  }
  return { nom: ext('Nom'), prenom: ext('Prénom'), email: ext('Email'), tel: ext('N° de Téléphone'),
           formation: ext('Formation souhaitée'), campus: ext('Campus souhaité') };
}

function emailExistsInEtudiants(ws, email) {
  if (!email) return false;
  var data = ws.getDataRange().getValues();
  var el = email.toLowerCase().trim();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][4]||"").toLowerCase().trim() === el) return true;
  }
  return false;
}

function handlePreviewInscriptions(p, ss) {
  try {
    if (p.vieScRole !== "Super-admin") return createJsonResponse({ success: false, message: "Accès refusé." });
    var LABEL = "Inscription_Traitee";
    var label = GmailApp.getUserLabelByName(LABEL);
    if (!label) label = GmailApp.createLabel(LABEL);
    var query = '(from:email@hecg.fr OR subject:"demande de renseignements") subject:inscription -label:'+LABEL;
    var threads = GmailApp.search(query, 0, 50);
    var ws = getSheetSafe(ss, "ETUDIANTS");
    var previews = [];
    threads.forEach(function(thread) {
      if (thread.getLabels().some(function(l){ return l.getName()===LABEL; })) return;
      thread.getMessages().forEach(function(msg) {
        var f = extractInscriptionFields(msg.getBody());
        if (!f.nom || !f.prenom) return;
        var doublon = emailExistsInEtudiants(ws, f.email);
        previews.push({ nom:f.nom, prenom:f.prenom, email:f.email, tel:f.tel,
                        formation:f.formation, campus:f.campus, doublon:doublon,
                        msgId:msg.getId(), threadId:thread.getId() });
      });
    });
    return createJsonResponse({ success: true, previews: previews });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleAspirerInscriptions(p, ss) {
  try {
    if (p.vieScRole !== "Super-admin") return createJsonResponse({ success: false, message: "Accès refusé." });
    var LABEL = "Inscription_Traitee";
    var label = GmailApp.getUserLabelByName(LABEL);
    if (!label) label = GmailApp.createLabel(LABEL);
    var query = '(from:email@hecg.fr OR subject:"demande de renseignements") subject:inscription -label:'+LABEL;
    var threads = GmailApp.search(query, 0, 50);
    var ws = getSheetSafe(ss, "ETUDIANTS"); var nb = 0; var inseres = [];
    // Filtre par sélection utilisateur (tableau de msgId)
    var selectedMsgIds = p.selectedMsgIds || null;
    threads.forEach(function(thread) {
      if (thread.getLabels().some(function(l){ return l.getName()===LABEL; })) return;
      var hasInsert = false;
      thread.getMessages().forEach(function(msg) {
        var f = extractInscriptionFields(msg.getBody());
        if (!f.nom || !f.prenom) return;
        if (emailExistsInEtudiants(ws, f.email)) return;
        // Si une sélection est fournie, ignorer les messages non sélectionnés
        if (selectedMsgIds && selectedMsgIds.indexOf(msg.getId()) === -1) return;
        var id = "ETU"+new Date().getTime().toString().slice(-5)+Math.floor(Math.random()*100);
        ws.appendRow([id,"",f.prenom,f.nom,f.email,f.tel,"","","","","","","","Prospect","","Non","Non","","","","PROSPECT",f.formation,f.campus]);
        logAction("Système","Auto","Création","ETUDIANTS",id,"Inscription web: "+f.prenom+" "+f.nom+" — "+f.formation+" — "+f.campus);
        nb++; inseres.push(f.prenom+" "+f.nom+" ("+f.formation+" / "+f.campus+")");
        hasInsert = true;
      });
      // Archiver le fil seulement si au moins un message a été importé depuis ce fil
      if (hasInsert) {
        thread.addLabel(label);
        thread.moveToArchive();
      }
    });
    return createJsonResponse({ success: true, nb: nb, inseres: inseres });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// ==========================================
// VIE SCOLAIRE — GED (GOOGLE DRIVE)
// ==========================================

const GED_ROOT_ID = "1zTxBYhZO3E8PeTGEksKKk3IBWk6btriR";

function handleGedListFolder(p, ss) {
  try {
    const fid    = p.folderId || GED_ROOT_ID;
    const folder = DriveApp.getFolderById(fid);
    const items  = [];
    const subs   = folder.getFolders();
    while (subs.hasNext()) { const f=subs.next(); items.push({ type:"folder", id:f.getId(), name:f.getName(), url:f.getUrl() }); }
    const fils   = folder.getFiles();
    while (fils.hasNext()) {
      const f=fils.next();
      items.push({ type:"file", id:f.getId(), name:f.getName(), url:f.getUrl(),
        mimeType:f.getMimeType(), size:f.getSize(),
        date:Utilities.formatDate(f.getLastUpdated(),Session.getScriptTimeZone(),"dd/MM/yyyy") });
    }
    // Vérification droits GED pour utilisateur non admin
    let droitsOk = true;
    if (p.vieScRole !== "Super-admin") {
      try {
        const wsD = getSheetSafe(ss, "GED_DROITS"); const dD = wsD.getDataRange().getValues();
        droitsOk = false;
        for (let i=1;i<dD.length;i++) {
          if (String(dD[i][0])!==fid) continue;
          const membres = JSON.parse(String(dD[i][2]||"[]"));
          if (membres.includes(p.emailAuteur) || membres.includes(p.vieScRole)) { droitsOk=true; break; }
        }
        if (!droitsOk && fid === GED_ROOT_ID) droitsOk = true; // Racine toujours accessible
      } catch(e) { droitsOk = true; }
    }
    if (!droitsOk) return createJsonResponse({ success: false, message: "Accès non autorisé à ce dossier." });
    return createJsonResponse({ success: true, items: items, folderName: folder.getName(), folderId: fid });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleGedCreateFolder(p, ss) {
  try {
    if (p.vieScRole !== "Super-admin") return createJsonResponse({ success: false, message: "Accès refusé." });
    if (!p.name || !p.name.trim()) return createJsonResponse({ success: false, message: "Nom requis." });
    const parent = DriveApp.getFolderById(p.parentId || GED_ROOT_ID);
    const nf     = parent.createFolder(p.name.trim());
    logAction(p.emailAuteur||"?","Super-admin","Création","GED",nf.getId(),"Dossier: "+p.name);
    return createJsonResponse({ success: true, id: nf.getId(), name: nf.getName(), url: nf.getUrl() });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleGedGetDroits(p, ss) {
  try {
    let wsD; try { wsD = getSheetSafe(ss,"GED_DROITS"); } catch(e) { return createJsonResponse({ success:true, membres:[] }); }
    const data = wsD.getDataRange().getValues();
    for (let i=1;i<data.length;i++) {
      if (String(data[i][0])===p.folderId) {
        try { return createJsonResponse({ success:true, membres:JSON.parse(String(data[i][2]||"[]")) }); }
        catch(e2) { return createJsonResponse({ success:true, membres:[] }); }
      }
    }
    return createJsonResponse({ success:true, membres:[] });
  } catch(e) { return createJsonResponse({ success:false, message:e.toString() }); }
}

function handleGedSetDroits(p, ss) {
  try {
    if (p.vieScRole !== "Super-admin") return createJsonResponse({ success:false, message:"Accès refusé." });
    let wsD;
    try { wsD = getSheetSafe(ss,"GED_DROITS"); }
    catch(e) { wsD=ss.insertSheet("GED_DROITS"); wsD.appendRow(["Drive_Folder_ID","Nom_Dossier","Membres_Autorises"]); wsD.setFrozenRows(1); }
    const data=wsD.getDataRange().getValues(); const val=JSON.stringify(p.membres||[]);
    for (let i=1;i<data.length;i++) { if (String(data[i][0])===p.folderId) { wsD.getRange(i+1,3).setValue(val); return createJsonResponse({ success:true }); } }
    wsD.appendRow([p.folderId, p.folderName||"", val]);
    return createJsonResponse({ success:true });
  } catch(e) { return createJsonResponse({ success:false, message:e.toString() }); }
}

function handleGedUploadFile(p, ss) {
  try {
    if (!p.folderId||!p.base64Data||!p.fileName||!p.mimeType)
      return createJsonResponse({ success:false, message:"Paramètres manquants." });
    const folder = DriveApp.getFolderById(p.folderId);
    const blob   = Utilities.newBlob(Utilities.base64Decode(p.base64Data), p.mimeType, p.fileName);
    const file   = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    logAction(p.emailAuteur||"?", p.vieScRole||"?", "Création","GED",file.getId(),"Fichier: "+p.fileName);
    return createJsonResponse({ success:true, id:file.getId(), url:file.getUrl(), name:file.getName() });
  } catch(e) { return createJsonResponse({ success:false, message:e.toString() }); }
}

// ==========================================
// GED — DOSSIERS ÉTUDIANTS
// ==========================================
const GED_ETU_ROOT_ID   = "1VbWpZ-m2puTE1cjdPKaIrqoQScDLGHSb";
const GED_SCAN_ROOT_ID  = "1oX-eay610TeLfRh8ng25ck2I9pCRp5pS"; // Arborescence Drive étudiants à scanner

function getOrCreateGedEtuSheet(ss) {
  var ws;
  try { ws = getSheetSafe(ss, "GED_ETU"); }
  catch(e) {
    ws = ss.insertSheet("GED_ETU");
    ws.getRange(1,1,1,4).setValues([["ID_Etudiant","NomPrenom","FolderID","FolderUrl"]]);
    ws.setFrozenRows(1); ws.getRange("1:1").setFontWeight("bold");
  }
  return ws;
}

function handleGedGetStudentDocs(p, ss) {
  try {
    var ws = getOrCreateGedEtuSheet(ss);
    var data = ws.getDataRange().getValues();
    var folderId = null, folderUrl = null;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(p.idEtu)) {
        folderId = String(data[i][2]||"").trim();
        folderUrl = String(data[i][3]||"").trim();
        break;
      }
    }
    if (!folderId) return createJsonResponse({ success: true, linked: false });
    var folder = DriveApp.getFolderById(folderId);
    var files = [];
    var fils = folder.getFiles();
    while (fils.hasNext()) {
      var f = fils.next();
      files.push({ id: f.getId(), name: f.getName(), url: f.getUrl(),
        mimeType: f.getMimeType(), size: f.getSize(),
        date: Utilities.formatDate(f.getLastUpdated(), Session.getScriptTimeZone(), "dd/MM/yyyy") });
    }
    files.sort(function(a,b){ return b.date.localeCompare(a.date); });
    var subfolders = [];
    var subs = folder.getFolders();
    while (subs.hasNext()) {
      var sf = subs.next();
      subfolders.push({ id: sf.getId(), name: sf.getName(), url: sf.getUrl() });
    }
    return createJsonResponse({ success: true, linked: true, folderId: folderId, folderUrl: folderUrl, files: files, subfolders: subfolders });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleGedLinkStudentFolder(p, ss) {
  try {
    if (p.vieScRole !== "Super-admin") return createJsonResponse({ success: false, message: "Accès refusé." });
    var fid = (p.folderUrl||"").replace(/.*\/folders\//,"").replace(/[?&].*/,"") || p.folderId||"";
    if (!fid) return createJsonResponse({ success: false, message: "URL ou ID invalide." });
    var folder = DriveApp.getFolderById(fid);
    var url = folder.getUrl();
    var ws = getOrCreateGedEtuSheet(ss);
    var data = ws.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(p.idEtu)) {
        ws.getRange(i+1,3).setValue(fid); ws.getRange(i+1,4).setValue(url);
        logAction(p.emailAuteur||p.idAuteur||"?", p.vieScRole||"?", "Liaison dossier GED", "GED_ETU", p.idEtu, (p.emailAuteur||"?") + " a lié le dossier Drive « " + folder.getName() + " » aux documents de " + (p.nomPrenom||p.idEtu||"?"));
        return createJsonResponse({ success: true, folderId: fid, folderUrl: url });
      }
    }
    ws.appendRow([p.idEtu||"", p.nomPrenom||"", fid, url]);
    logAction(p.emailAuteur||p.idAuteur||"?", p.vieScRole||"?", "Liaison dossier GED", "GED_ETU", p.idEtu, (p.emailAuteur||"?") + " a lié le dossier Drive « " + folder.getName() + " » aux documents de " + (p.nomPrenom||p.idEtu||"?"));
    return createJsonResponse({ success: true, folderId: fid, folderUrl: url });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleGedUploadStudentDoc(p, ss) {
  try {
    if (!p.base64Data || !p.fileName || !p.mimeType) return createJsonResponse({ success: false, message: "Paramètres manquants." });
    var ws = getOrCreateGedEtuSheet(ss);
    var data = ws.getDataRange().getValues();
    var folderId = null;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(p.idEtu)) { folderId = String(data[i][2]||"").trim(); break; }
    }
    if (!folderId) {
      // Auto-create dans le dossier racine si pas encore lié
      var root = DriveApp.getFolderById(GED_ETU_ROOT_ID);
      var newFolder = null;
      try { var it=root.getFoldersByName(p.nomPrenom||p.idEtu); newFolder=it.hasNext()?it.next():root.createFolder(p.nomPrenom||p.idEtu); }
      catch(ec){ newFolder=root.createFolder(p.nomPrenom||p.idEtu); }
      folderId = newFolder.getId();
      // Link automatiquement
      for (var j = 1; j < data.length; j++) {
        if (String(data[j][0]) === String(p.idEtu)) { ws.getRange(j+1,3).setValue(folderId); ws.getRange(j+1,4).setValue(newFolder.getUrl()); break; }
      }
      if (folderId === newFolder.getId()) ws.appendRow([p.idEtu||"", p.nomPrenom||"", folderId, newFolder.getUrl()]);
    }
    var folder = DriveApp.getFolderById(folderId);
    var blob = Utilities.newBlob(Utilities.base64Decode(p.base64Data), p.mimeType, p.fileName);
    var file = folder.createFile(blob);
    var nomEtu = p.nomPrenom || p.idEtu || "?";
    var auteurLog = p.emailAuteur || p.idAuteur || "?";
    logAction(auteurLog, p.vieScRole||p.roleAuteur||"?", "Dépôt document", "GED_ETU", p.idEtu, auteurLog + " a déposé « " + p.fileName + " » dans les documents de " + nomEtu);
    return createJsonResponse({ success: true, id: file.getId(), url: file.getUrl(), name: file.getName() });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleGedAutoScanStudents(p, ss) {
  try {
    if (p.vieScRole !== "Super-admin") return createJsonResponse({ success: false, message: "Accès refusé." });
    var root = DriveApp.getFolderById(GED_ETU_ROOT_ID);
    var items = [];
    var subs = root.getFolders();
    while (subs.hasNext()) { var f=subs.next(); items.push({ name:f.getName(), id:f.getId(), url:f.getUrl() }); }
    // Tentative de lister aussi les fichiers (raccourcis)
    var fils = root.getFiles();
    while (fils.hasNext()) { var ff=fils.next(); if(ff.getMimeType()==='application/vnd.google-apps.shortcut') items.push({ name:ff.getName(), id:ff.getId(), url:ff.getUrl(), isShortcut:true }); }

    var dE = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues();
    var students = [];
    for (var i = 1; i < dE.length; i++) { if(dE[i][0]) students.push({ id:String(dE[i][0]), nom:String(dE[i][2]||""), prenom:String(dE[i][3]||"") }); }

    var ws = getOrCreateGedEtuSheet(ss);
    var gedData = ws.getDataRange().getValues();
    var linkedMap = {};
    for (var k = 1; k < gedData.length; k++) if (gedData[k][0]) linkedMap[String(gedData[k][0])] = String(gedData[k][2]||"");

    var matched = 0;
    items.forEach(function(item){
      var nameNorm = item.name.toLowerCase().trim();
      var stu = students.find(function(s){
        var n1=(s.nom+' '+s.prenom).toLowerCase().trim();
        var n2=(s.prenom+' '+s.nom).toLowerCase().trim();
        return n1===nameNorm||n2===nameNorm;
      });
      if (stu && !linkedMap[stu.id]) {
        var existRow = -1;
        for (var gi=1;gi<gedData.length;gi++) if(String(gedData[gi][0])===stu.id){existRow=gi;break;}
        if (existRow>0) { ws.getRange(existRow+1,3).setValue(item.id); ws.getRange(existRow+1,4).setValue(item.url||""); }
        else { ws.appendRow([stu.id, stu.nom+' '+stu.prenom, item.id, item.url||""]); }
        gedData.push([stu.id,'',item.id,'']); // update local cache
        linkedMap[stu.id] = item.id; matched++;
      }
    });
    return createJsonResponse({ success:true, total:items.length, matched:matched });
  } catch(e) { return createJsonResponse({ success:false, message:e.toString() }); }
}

/**
 * Recherche les dossiers Drive correspondant à un étudiant dans GED_SCAN_ROOT_ID (tous niveaux, BFS).
 * Retourne les candidats sans lier — la liaison est confirmée manuellement via gedLinkStudentFolder.
 */
function handleGedScanForStudent(p, ss) {
  try {
    var idEtu  = String(p.idEtu  || "").trim();
    var nom    = String(p.nom    || "").trim();
    var prenom = String(p.prenom || "").trim();

    // Récupère nom/prénom depuis ETUDIANTS si non fournis
    if (idEtu && (!nom || !prenom)) {
      var dE = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues();
      for (var i = 1; i < dE.length; i++) {
        if (String(dE[i][0]).trim() === idEtu) {
          nom    = String(dE[i][2] || "").trim();
          prenom = String(dE[i][3] || "").trim();
          break;
        }
      }
    }
    if (!nom && !prenom) return createJsonResponse({ success: false, message: "Nom étudiant introuvable." });

    // Normalisation : minuscules, suppression accents (Unicode range correct), espaces normalisés
    function norm(s) {
      return String(s || "").toLowerCase()
        .normalize("NFD").replace(/[̀-ͯ]/g, "")
        .replace(/[^a-z0-9]/g, " ").replace(/\s+/g, " ").trim();
    }

    // Le drive utilise le format "NOM Prénom" (NOM en caps, Prénom 1ère lettre maj)
    // On compare après normalisation pour être insensible à la casse et aux accents
    var nomNorm    = norm(nom);
    var prenomNorm = norm(prenom);

    function match(folderName) {
      var fn = norm(folderName);
      if (!fn || !nomNorm) return false;

      // Combinaisons exactes "NOM Prénom" / "Prénom NOM"
      var exact1 = nomNorm + (prenomNorm ? " " + prenomNorm : "");
      var exact2 = prenomNorm ? prenomNorm + " " + nomNorm : "";
      if (fn === exact1 || (exact2 && fn === exact2)) return true;
      if (prenomNorm && (fn.startsWith(exact1 + " ") || fn.startsWith(exact2 + " "))) return true;
      if (!prenomNorm && fn === nomNorm) return true;

      // Le dossier contient la combinaison nom+prénom quelque part
      if (fn.indexOf(exact1) !== -1 || (exact2 && fn.indexOf(exact2) !== -1)) return true;

      // Correspondance par tokens : tous les mots significatifs (>1 lettre) du nom ET prénom
      // apparaissent dans le nom du dossier
      var nomTokens    = nomNorm.split(" ").filter(function(t){ return t.length > 1; });
      var prenomTokens = prenomNorm ? prenomNorm.split(" ").filter(function(t){ return t.length > 1; }) : [];
      var allTokens    = nomTokens.concat(prenomTokens);
      if (allTokens.length >= 2) {
        if (allTokens.every(function(t){ return fn.indexOf(t) !== -1; })) return true;
      }
      // Nom seul + prénom partiel (au moins 3 lettres du prénom)
      if (nomTokens.length >= 1 && prenomTokens.length >= 1) {
        var nomOk  = nomTokens.every(function(t){ return fn.indexOf(t) !== -1; });
        var prenomPartial = prenomTokens.some(function(t){ return fn.indexOf(t.slice(0,3)) !== -1; });
        if (nomOk && prenomPartial) return true;
      }
      // Tolérance 1 caractère sur le nom (noms courts peuvent être abrégés)
      if (nomNorm.length >= 5) {
        var nomPrefix4 = nomNorm.slice(0, 4);
        if (fn.indexOf(nomPrefix4) !== -1 && prenomNorm.length >= 2 && fn.indexOf(prenomNorm.slice(0, 2)) !== -1) return true;
      }
      return false;
    }

    var candidates = [];
    var root = DriveApp.getFolderById(GED_SCAN_ROOT_ID);

    // BFS — parcours en largeur sur TOUTE l'arborescence (profondeur illimitée)
    // Garde-fous : max 2000 dossiers inspectés ET max 25 secondes (limite Apps Script = 30s)
    var MAX_INSPECT  = 2000;
    var MAX_TIME_MS  = 25000;
    var startTime    = Date.now();
    var inspected    = 0;
    var truncated    = false;
    var queue = [{ folder: root, path: "" }];

    outer: while (queue.length > 0) {
      // Vérification temps restant
      if (Date.now() - startTime > MAX_TIME_MS) { truncated = true; break outer; }
      if (inspected >= MAX_INSPECT)              { truncated = true; break outer; }

      var current = queue.shift();
      inspected++;

      try {
        var children = current.folder.getFolders();
        while (children.hasNext()) {
          var child     = children.next();
          var childPath = current.path ? current.path + " / " + child.getName() : child.getName();
          if (match(child.getName())) {
            // Correspondance → candidat proposé, on n'inspecte pas ses sous-dossiers
            candidates.push({ id: child.getId(), name: child.getName(), url: child.getUrl(), path: childPath });
          } else {
            // Pas de correspondance → on descend pour chercher plus profond
            queue.push({ folder: child, path: childPath });
          }
        }
      } catch(eChild) {}
    }

    return createJsonResponse({
      success: true, candidates: candidates,
      nom: nom, prenom: prenom,
      inspected: inspected, truncated: truncated
    });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// ==========================================
// VIE SCOLAIRE — QUALIOPI
// ==========================================

function handleGetQualiopiPreuves(p, ss) {
  try {
    const data = getSheetSafe(ss,"QUALIOPI_PREUVES").getDataRange().getValues(); const rows = [];
    for (let i=1;i<data.length;i++) {
      if (!data[i][0]) continue;
      rows.push({ indicateur:String(data[i][0]), nomPreuve:String(data[i][1]),
        urlFichier:String(data[i][2]||""), dateDepot:String(data[i][3]||"") });
    }
    return createJsonResponse({ success:true, data:rows });
  } catch(e) { return createJsonResponse({ success:false, message:e.toString() }); }
}

function handleSaveQualiopiPreuve(p, ss) {
  try {
    if (p.vieScRole !== "Super-admin") return createJsonResponse({ success:false, message:"Accès refusé." });
    const ws   = getSheetSafe(ss,"QUALIOPI_PREUVES");
    const date = Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"dd/MM/yyyy");
    ws.appendRow([p.indicateur||"", p.nomPreuve||"", p.urlFichier||"", p.dateDepot||date]);
    logAction(p.emailAuteur||"?","Super-admin","Création","Vie scolaire",p.indicateur,"Qualiopi preuve: "+p.nomPreuve);
    return createJsonResponse({ success:true });
  } catch(e) { return createJsonResponse({ success:false, message:e.toString() }); }
}

function handleDeleteQualiopiPreuve(p, ss) {
  try {
    if (p.vieScRole !== "Super-admin") return createJsonResponse({ success:false, message:"Accès refusé." });
    const ws = getSheetSafe(ss,"QUALIOPI_PREUVES"); const data = ws.getDataRange().getValues();
    for (let i=1;i<data.length;i++) {
      if (String(data[i][0])===String(p.indicateur) && String(data[i][1])===String(p.nomPreuve)
          && String(data[i][2])===String(p.urlFichier)) {
        ws.deleteRow(i+1); return createJsonResponse({ success:true });
      }
    }
    return createJsonResponse({ success:false, message:"Preuve introuvable." });
  } catch(e) { return createJsonResponse({ success:false, message:e.toString() }); }
}

// ==========================================
// VIE SCOLAIRE — NOTES EN MASSE & RELEVÉ
// ==========================================

function normMatiere(s) {
  var m = String(s||'').trim().match(/^(UE\s*\d+[a-zA-Z]*)/i);
  return m ? m[1].toLowerCase().replace(/\s+/g,'') : String(s||'').trim().toLowerCase();
}

function handleSaveNotesBulk(p, ss) {
  try {
    if (p.vieScRole !== "Super-admin" && p.vieScRole !== "Pédagogie")
      return createJsonResponse({ success: false, message: "Accès refusé." });
    const ws   = getSheetSafe(ss, "NOTES_EXAMENS");
    const data = ws.getDataRange().getValues();
    const notes = p.notes || [];  // [{matiere, note, coeff}]
    let nbInsert = 0, nbUpdate = 0;
    for (const n of notes) {
      if (n.note === "" || n.note === null || n.note === undefined) continue;
      const rawN = String(n.note).trim();
      const isNE = rawN.toUpperCase() === "NE";
      const note   = isNE ? "NE" : parseFloat(rawN);
      const coeff  = parseFloat(n.coeff) || 0;
      const points = isNE ? 0 : Math.round(parseFloat(note) * coeff * 100) / 100;
      // Cherche note existante : même idEtu + annee + matiere (matching normalisé sur nom UE)
      let found = false;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][1]) !== String(p.idEtu)) continue;
        if (String(data[i][4]).trim() !== String(p.annee).trim()) continue;
        if (normMatiere(data[i][5]) !== normMatiere(n.matiere)) continue;
        ws.getRange(i+1,7).setValue(note);
        ws.getRange(i+1,8).setValue(coeff);
        ws.getRange(i+1,9).setValue(isNE ? 0 : points);
        if (!String(data[i][3]).trim() && (p.classe || p.diplome)) ws.getRange(i+1,4).setValue(p.classe||p.diplome||"");
        if (n.sujetMemoire)    ws.getRange(i+1,11).setValue(n.sujetMemoire);
        if (n.referentMemoire) ws.getRange(i+1,12).setValue(n.referentMemoire);
        data[i][6]=note; data[i][7]=coeff; data[i][8]=points; // maj mémoire
        found = true; nbUpdate++;
        break;
      }
      if (!found) {
        const id = "NOT-" + new Date().getTime() + "-" + Math.floor(Math.random()*999);
        ws.appendRow([id, p.idEtu||"", p.nomPrenom||"", p.classe||p.diplome||"", p.annee||"",
          n.matiere||"", note, coeff, points, "", n.sujetMemoire||"", n.referentMemoire||""]);
        nbInsert++;
      }
    }
    logAction(p.emailAuteur||"?", p.vieScRole||"?", "Modification", "Vie scolaire", p.idEtu,
      "Saisie notes: "+p.nomPrenom+" — "+p.annee+" — "+nbInsert+" créées, "+nbUpdate+" mises à jour");
    return createJsonResponse({ success: true, nbInsert: nbInsert, nbUpdate: nbUpdate });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

const RELEVE_ROOT_ID = "1UGf647KDLmzbn8uZPn8xLnzxmchFAgl-";

function handleUploadReleveNote(p, ss) {
  try {
    if (p.vieScRole !== "Super-admin" && p.vieScRole !== "Pédagogie")
      return createJsonResponse({ success: false, message: "Accès refusé." });
    if (!p.base64Data || !p.fileName || !p.mimeType)
      return createJsonResponse({ success: false, message: "Paramètres manquants." });
    // Trouver/créer le sous-dossier [Diplôme]
    const root    = DriveApp.getFolderById(RELEVE_ROOT_ID);
    const diplome = (p.diplome || "Autres").replace(/[\\/:*?"<>|]/g, '_');
    const annee   = (p.annee   || "").replace(/[\\/:*?"<>|]/g, '_');
    function getOrCreate(parent, name) {
      const it = parent.getFoldersByName(name);
      return it.hasNext() ? it.next() : parent.createFolder(name);
    }
    const dipFolder  = getOrCreate(root, diplome);
    const annFolder  = getOrCreate(dipFolder, annee);
    // Renommer le fichier : ANNEE_NOM_PRENOM_DIPLOME_CLASSE.ext
    const ext        = p.fileName.includes('.') ? p.fileName.split('.').pop() : '';
    const nomClean   = (p.nom||p.nomPrenom||"").replace(/[\\/:*?"<>|]/g,'_').replace(/\s+/g,'_');
    const prenomClean= (p.prenom||"").replace(/[\\/:*?"<>|]/g,'_').replace(/\s+/g,'_');
    const dipClean   = (p.diplome||"").replace(/[\\/:*?"<>|\\s]/g,'_');
    const classClean = (p.classe||"").replace(/[\\/:*?"<>|]/g,'_').replace(/\s+/g,'_');
    const nameParts  = [annee, nomClean, prenomClean, dipClean, classClean].filter(Boolean);
    const newName    = (nameParts.join("_") + (ext ? '.'+ext : '')).replace(/__+/g,'_');
    const blob  = Utilities.newBlob(Utilities.base64Decode(p.base64Data), p.mimeType, newName);
    const file  = annFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    logAction(p.emailAuteur||"?", p.vieScRole||"?", "Création","Vie scolaire",file.getId(),"Relevé de notes: "+newName);
    return createJsonResponse({ success: true, id: file.getId(), url: file.getUrl(), name: newName });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// ==========================================
// VIE SCOLAIRE — PERSONNEL LIST (GED DROITS)
// ==========================================

function handleGetPersonnelList(p, ss) {
  try {
    const data = getSheetSafe(ss, "Personnel").getDataRange().getValues();
    const list = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      const email = String(data[i][6]||"").trim().toLowerCase();
      const nom   = String(data[i][2]||"")+" "+String(data[i][1]||"");
      list.push({ ref: String(data[i][0]), nom: nom.trim(), email: email, role: String(data[i][5]||"") });
    }
    return createJsonResponse({ success: true, data: list });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// ==========================================
// VIE SCOLAIRE — QUALIOPI UPLOAD FICHIER
// ==========================================

const QUALIOPI_ROOT_ID = "1Byn9V9FX1N0enndO1CnI8DnIlnu-YoKl";

function handleUploadQualiopiFile(p, ss) {
  try {
    if (p.vieScRole !== "Super-admin") return createJsonResponse({ success: false, message: "Accès refusé." });
    if (!p.base64Data || !p.fileName || !p.mimeType)
      return createJsonResponse({ success: false, message: "Paramètres manquants." });
    // Sous-dossier : CRITERE_INDICATEUR (ex: "1_3")
    const folderName = String(p.critere||"0") + "_" + String(p.indicateur||"0");
    const root = DriveApp.getFolderById(QUALIOPI_ROOT_ID);
    const it   = root.getFoldersByName(folderName);
    const folder = it.hasNext() ? it.next() : root.createFolder(folderName);
    const blob = Utilities.newBlob(Utilities.base64Decode(p.base64Data), p.mimeType, p.fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    // Sauvegarder la preuve en base
    const ws   = getSheetSafe(ss,"QUALIOPI_PREUVES");
    const date = Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"dd/MM/yyyy");
    ws.appendRow([p.indicateur||"", p.nomPreuve||p.fileName, file.getUrl(), p.dateDepot||date]);
    logAction(p.emailAuteur||"?","Super-admin","Création","Vie scolaire",p.indicateur,"Qualiopi preuve: "+p.fileName);
    return createJsonResponse({ success: true, id: file.getId(), url: file.getUrl(), folderUrl: folder.getUrl() });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleGetQualiopiDossierUrl(p, ss) {
  try {
    if (p.vieScRole !== "Super-admin") return createJsonResponse({ success: false, message: "Accès refusé." });
    const folderName = String(p.critere||"0") + "_" + String(p.indicateur||"0");
    const root = DriveApp.getFolderById(QUALIOPI_ROOT_ID);
    const it   = root.getFoldersByName(folderName);
    if (it.hasNext()) return createJsonResponse({ success: true, url: it.next().getUrl() });
    return createJsonResponse({ success: true, url: root.getUrl() + "?usp=sharing" }); // parent si pas encore créé
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// ==========================================
// SEEDER ONE-SHOT — RÉFÉRENTIEL MATIÈRES
// ==========================================
/**
 * À exécuter UNE SEULE FOIS depuis l'éditeur Apps Script.
 * Initialise Referentiel_matieres avec BTS CG, DCG et DSCG.
 */
function seederReferentielMatieres() {
  const ss = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");
  let ws;
  try { ws = getSheetSafe(ss, "Referentiel_matieres"); }
  catch(e) {
    ws = ss.insertSheet("Referentiel_matieres");
    ws.appendRow(["Diplome","Matiere","Type","Coefficient"]);
    ws.setFrozenRows(1); ws.getRange("1:1").setFontWeight("bold");
  }
  ws.getRange(1,3).setValue("Type"); // Assure l'en-tête col C
  const lastRow = ws.getLastRow();
  if (lastRow > 1) ws.deleteRows(2, lastRow - 1);
  const rows = [
    // BTS CG — coeff officiels (arrêté BTS CG)
    ["BTS CG","CEJM",      "generale",   3],
    ["BTS CG","Math",       "generale",   2],
    ["BTS CG","CGE",        "generale",   5],
    ["BTS CG","Anglais",    "generale",   2],
    ["BTS CG","E5",         "pro",        4],
    ["BTS CG","E6",         "pro",        3],
    ["BTS CG","E7",         "pro",        3],
    ["BTS CG","E8",         "pro",        4],
    ["BTS CG","EF1",        "facultative",2],
    ["BTS CG","EF2",        "facultative",2],
    ["BTS CG","EF3",        "facultative",2],
    ["BTS CG","Autre",      "autre",      1],
    // DCG — 14 UE, coeff = 1 pour toutes, UE14 facultative
    ["DCG","UE 1 - Introduction au droit",                      "principale",1],
    ["DCG","UE 2 - Droit des sociétés et des groupements",      "principale",1],
    ["DCG","UE 3 - Droit social",                               "principale",1],
    ["DCG","UE 4 - Droit fiscal",                               "principale",1],
    ["DCG","UE 5 - Économie contemporaine",                     "principale",1],
    ["DCG","UE 6 - Finance d'entreprise",                       "principale",1],
    ["DCG","UE 7 - Management et organisation",                 "principale",1],
    ["DCG","UE 8 - Systèmes d'information de gestion",          "principale",1],
    ["DCG","UE 9 - Introduction à la comptabilité",             "principale",1],
    ["DCG","UE 10 - Comptabilité approfondie",                  "principale",1],
    ["DCG","UE 11 - Contrôle de gestion",                       "principale",1],
    ["DCG","UE 12 - Anglais des affaires",                      "principale",1],
    ["DCG","UE 13 - Relations professionnelles",                "principale",1],
    ["DCG","UE 14 - Espagnol des affaires",                     "facultative",1],
    // DSCG — 6 UE obligatoires + UE7 facultative
    ["DSCG","UE 1 - Gestion juridique, fiscale et sociale",     "principale",1],
    ["DSCG","UE 2 - Finance",                                   "principale",1],
    ["DSCG","UE 3 - Management et contrôle de gestion",         "principale",1],
    ["DSCG","UE 4 - Comptabilité et audit",                     "principale",1],
    ["DSCG","UE 5 - Management des systèmes d'information",     "principale",1],
    ["DSCG","UE 6 - Relations professionnelles",                "principale",1],
    ["DSCG","UE 7 - Langue vivante étrangère",                  "facultative",1],
  ];
  ws.getRange(2, 1, rows.length, 4).setValues(rows);
  console.log("✅ seederReferentielMatieres : " + rows.length + " matières initialisées.");
}

// ==========================================
// PLANNING — MIGRATION TACHES → MISSIONS (semaine 25-30 mai 2026)
// ==========================================

function handleMigratePlanningWeek(p, ss) {
  try {
    if (p.vieScRole !== "Super-admin") return createJsonResponse({ success: false, message: "Accès refusé." });
    var planWs   = getSheetSafe(ss, "PLANNING");
    var planData = planWs.getDataRange().getValues();
    var mSheet   = getSheetSafe(ss, "MISSIONS");
    var nb = 0;
    for (var i = 1; i < planData.length; i++) {
      var dateDebut = String(planData[i][2]||"").trim();
      if (dateDebut < "2026-05-25" || dateDebut > "2026-05-30") continue;
      if (String(planData[i][12]||"").trim()) continue; // déjà migré
      var emailDest = String(planData[i][4]||"").trim().toLowerCase();
      if (!emailDest) continue;
      var mNow = new Date();
      var mId  = "M" + (mNow.getTime() + i); // unicité garantie
      mSheet.appendRow([
        mId,
        Utilities.formatDate(mNow, Session.getScriptTimeZone(), "dd/MM/yyyy"),
        Utilities.formatDate(mNow, Session.getScriptTimeZone(), "HH:mm"),
        emailDest,
        String(planData[i][9]||""),
        "",
        String(planData[i][1]||""),
        String(planData[i][6]||""),
        "Attribué",
        "",
        mNow,
        Number(planData[i][11]||3),
        "[Planning] " + String(planData[i][1]||"")
      ]);
      planWs.getRange(i+1, 13).setValue(mId);
      logAction(p.emailAuteur||"?", p.vieScRole||"?", "Création", "Missions", mId,
        "Migration planning 25-30 mai : " + String(planData[i][1]||"").substring(0,60));
      nb++;
    }
    return createJsonResponse({ success: true, nb: nb });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// ==========================================
// TROPHEES — GET DATA
// ==========================================

function handleGetTropheeData(p, ss) {
  try {
    const emailUser  = String(p.emailUser  || "").trim().toLowerCase();
    const nomUser    = String(p.nomUser    || "").trim();
    if (!emailUser) return createJsonResponse({ success: false, message: "emailUser manquant." });

    // --- 1. Lire HISTORIQUE et compter par catégorie ---
    var histWs, histData;
    try { histWs = getSheetSafe(ss, "HISTORIQUE"); histData = histWs.getDataRange().getValues(); }
    catch(e) { histData = []; }

    const XP_MAP = {
      Actions:1, Offres:10, CarnetRoute:5, CarnetEntreprise:15,
      Missions:15, Objectifs:50, Alertes:10, Com:35, Entreprise:25, Etudiant:20, Special:100
    };

    var count = { Actions:0, Offres:0, CarnetRoute:0, CarnetEntreprise:0, Missions:0, Objectifs:0, Alertes:0, Com:0, Entreprise:0, Etudiant:0, Special:0 };
    var actionRows = [];
    var missionsSelfCreated = 0;

    for (var i = 1; i < histData.length; i++) {
      var row = histData[i];
      var auteur  = String(row[2]||"").trim().toLowerCase();
      var role    = String(row[3]||"").trim();
      var typeAct = String(row[4]||"").trim();
      var cat     = String(row[5]||"").trim();
      var cible   = String(row[6]||"").trim();
      var details = String(row[7]||"").trim();
      var dateVal = row[1];

      // Matcher l'utilisateur (col C = nom ou email)
      var match = (auteur === emailUser) || (auteur === nomUser.toLowerCase()) || (auteur.indexOf(emailUser) !== -1);
      if (!match) continue;

      // Compter Actions (toutes)
      count.Actions++;

      // Catégories spécifiques
      if (cat === 'Offres' && (typeAct === 'Modification' || typeAct === 'Création')) count.Offres++;
      if (cat === 'Carnet de route') count.CarnetRoute++;
      if (cat === 'Note Partenaire') count.CarnetEntreprise++;
      if (cat === 'Missions' && typeAct === 'Modification' && details.toLowerCase().indexOf('clôturé') !== -1) count.Missions++;
      if (cat === 'Missions' && typeAct === 'Création') missionsSelfCreated++;
      if ((cat === 'Alertes Partenaires' || cat === 'Alertes Étudiants') && typeAct === 'Création') count.Alertes++;
      if (cat === 'Com & Événements' && typeAct === 'Création') count.Com++;
      if (cat === 'Entreprise' && typeAct === 'Création') count.Entreprise++;
      if (cat === 'Dossier étudiant' && typeAct === 'Création') count.Etudiant++;

      actionRows.push({
        date: dateVal ? Utilities.formatDate(new Date(dateVal), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : "",
        type: typeAct, categorie: cat, cible: cible, details: details
      });
    }

    // --- 2. Compter objectifs validés (équipe) ---
    try {
      var objWs = ss.getSheetByName("OBJECTIFS_SEM");
      if (objWs) {
        var objData = objWs.getDataRange().getValues();
        var objValides = 0;
        for (var j = 1; j < objData.length; j++) {
          if (String(objData[j][8]||"").toLowerCase() === 'true' || objData[j][8] === true) objValides++;
        }
        count.Objectifs = objValides;
      }
    } catch(e) {}

    // --- 2b. Missions assignées et clôturées (crédit plein pour l'assigné) ---
    try {
      var missWs = ss.getSheetByName("MISSIONS");
      if (missWs) {
        var missData = missWs.getDataRange().getValues();
        for (var mi = 1; mi < missData.length; mi++) {
          var mDest = String(missData[mi][3]||"").trim().toLowerCase();
          var mStatut = String(missData[mi][8]||"").trim();
          if (mDest === emailUser && mStatut === 'Clôturé') count.Missions++;
        }
      }
    } catch(e) {}
    // 1/3 crédit pour les missions créées de manière autonome
    count.Missions = Math.round(count.Missions + missionsSelfCreated / 3);

    // --- 3. Calculer XP total ---
    var xpTotal = 0;
    for (var cat in XP_MAP) { xpTotal += (count[cat] || 0) * XP_MAP[cat]; }

    // --- 4. Prefs + ref utilisateur (Personnel) — lu tôt pour le matching TROPHESPEC ---
    var prefs = { pseudo: "", personnageChoisi: "", emblemeChoisi: "", tropheesVus: "" };
    var personnelRow = -1;
    var userRef = "";
    try {
      var persWs = getSheetSafe(ss, "Personnel");
      var persData = persWs.getDataRange().getValues();
      for (var pr = 1; pr < persData.length; pr++) {
        if (String(persData[pr][6]||"").trim().toLowerCase() === emailUser) {
          personnelRow = pr + 1;
          userRef                = String(persData[pr][0]||"").trim().toLowerCase();
          prefs.pseudo           = String(persData[pr][8]||"");
          prefs.personnageChoisi = String(persData[pr][9]||"");
          prefs.emblemeChoisi    = String(persData[pr][10]||"");
          prefs.tropheesVus      = String(persData[pr][11]||"");
          break;
        }
      }
    } catch(e) {}

    // --- 5. Lire feuille RECOMPENSE ---
    var trophees = [];
    try {
      var recWs = getSheetSafe(ss, "RECOMPENSE");
      var recData = recWs.getDataRange().getValues();
      const CAT_NORM = {
        'Actions':'Actions','Offres':'Offres','Carnet de route':'CarnetRoute',
        'Carnet entreprise':'CarnetEntreprise','Missions':'Missions','Objectifs':'Objectifs',
        'Alerte':'Alertes','com':'Com','Entreprise':'Entreprise','Etudiant':'Etudiant','Spécial':'Special'
      };
      for (var r = 1; r < recData.length; r++) {
        var catRec = String(recData[r][3]||"").trim();
        if (!catRec) continue;
        var qte = Number(recData[r][4]||0);
        if (!qte) continue;
        var idRec  = recData[r][0] ? String(recData[r][0]).trim() : ('t_'+r);
        var nom    = String(recData[r][1]||"").trim();
        var urlRec = String(recData[r][2]||"").trim();
        var catKey = CAT_NORM[catRec] || catRec;
        var tier   = String(recData[r][5]||"").trim().toLowerCase();
        var unlocked = (count[catKey] || 0) >= qte;
        trophees.push({ id: idRec, nom: nom, url: urlRec, tier: tier, categorie: catRec, catKey: catKey, quantite: qte, unlocked: unlocked, count: count[catKey]||0 });
      }
    } catch(e) {}

    // --- 6. Lire TROPHESPEC (trophées spéciaux attribués manuellement depuis le Sheet) ---
    try {
      var specWs = ss.getSheetByName("TROPHESPEC");
      if (specWs) {
        var specData = specWs.getDataRange().getValues();
        for (var sp = 1; sp < specData.length; sp++) {
          if (!specData[sp][0]) continue;
          var spId   = String(specData[sp][0]).trim();
          var spNom  = String(specData[sp][1]||"").trim();
          var spUsrs = String(specData[sp][2]||"").trim().toLowerCase();
          var spType = String(specData[sp][3]||"").trim().toLowerCase();
          var spDesc = String(specData[sp][4]||"").trim();
          var spUrl  = String(specData[sp][5]||"").trim();
          if (!spId || !spNom) continue;
          var destList = spUsrs.split(",").map(function(s){ return s.trim(); });
          var isForUser = spUsrs === "tous" || destList.indexOf(emailUser) !== -1 || (userRef && destList.indexOf(userRef) !== -1);
          if (!isForUser) continue;
          var spTier = spType.indexOf("arc") !== -1 ? "arc-en-ciel" : "or";
          trophees.push({ id: spId, nom: spNom, url: spUrl, tier: spTier,
            categorie: "Spécial", catKey: "Special", quantite: 1, count: 1,
            unlocked: true, description: spDesc, isSpecial: true });
        }
      }
    } catch(e) {}

    // --- 7. Lire feuille PERSONNAGE ---
    var personnages = [];
    try {
      var persoWs = getSheetSafe(ss, "PERSONNAGE");
      var persoData = persoWs.getDataRange().getValues();
      for (var p2 = 1; p2 < persoData.length; p2++) {
        if (!persoData[p2][0] && !persoData[p2][1]) continue;
        var pRef = persoData[p2][0] ? String(persoData[p2][0]).trim() : ('p_'+p2);
        var pNom = String(persoData[p2][1]||"").trim();
        var pUrl = String(persoData[p2][2]||"").trim();
        var pXp  = Number(persoData[p2][3]||0);
        personnages.push({
          ref: pRef, nom: pNom, url: pUrl, xpRequired: pXp,
          unlocked: xpTotal >= pXp
        });
      }
    } catch(e) {}

    var niveau = personnages.filter(function(c){ return c.unlocked; }).length;

    // --- 8. Lire SUIVITROPHEE pour dates d'obtention (réguliers + spéciaux) ---
    var suivi = {};
    try {
      var suiviWs = ss.getSheetByName("SUIVITROPHEE");
      if (!suiviWs) { suiviWs = ss.insertSheet("SUIVITROPHEE"); suiviWs.appendRow(["refpersonnel","idrecompense","datetrophee"]); }
      var suiviData = suiviWs.getDataRange().getValues();
      for (var s = 1; s < suiviData.length; s++) {
        if (String(suiviData[s][0]||"").trim().toLowerCase() === emailUser) {
          suivi[String(suiviData[s][1])] = suiviData[s][2];
        }
      }
    } catch(e) {}

    // Enregistrer nouvelles obtentions (réguliers + spéciaux)
    var newUnlocks = [];
    trophees.forEach(function(t) {
      if (t.unlocked && !suivi[t.id]) {
        newUnlocks.push(t.id);
        try { suiviWs.appendRow([emailUser, t.id, new Date()]); } catch(e) {}
        suivi[t.id] = new Date();
      }
      t.dateObtention = suivi[t.id] ? Utilities.formatDate(new Date(suivi[t.id]), Session.getScriptTimeZone(), "dd/MM/yyyy") : null;
    });

    return createJsonResponse({
      success: true,
      xpTotal: xpTotal, niveau: niveau, count: count,
      trophees: trophees, personnages: personnages,
      actionRows: actionRows, prefs: prefs, newUnlocks: newUnlocks
    });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// ==========================================
// TROPHEES — SAVE PREFS
// ==========================================

function handleSaveTropheePrefs(p, ss) {
  try {
    const emailUser = String(p.emailUser || "").trim().toLowerCase();
    if (!emailUser) return createJsonResponse({ success: false, message: "emailUser manquant." });
    var persWs = getSheetSafe(ss, "Personnel");
    var persData = persWs.getDataRange().getValues();
    for (var i = 1; i < persData.length; i++) {
      if (String(persData[i][6]||"").trim().toLowerCase() === emailUser) {
        var row = i + 1;
        if (p.field === 'pseudo')               persWs.getRange(row, 9).setValue(String(p.val||""));   // col I
        else if (p.field === 'personnageChoisi') persWs.getRange(row, 10).setValue(String(p.val||"")); // col J
        else if (p.field === 'emblemeChoisi')    persWs.getRange(row, 11).setValue(String(p.val||"")); // col K
        return createJsonResponse({ success: true });
      }
    }
    return createJsonResponse({ success: false, message: "Utilisateur introuvable." });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// ==========================================
// TROPHEES — MARK VUS
// ==========================================

function handleMarkTropheesVus(p, ss) {
  try {
    const emailUser = String(p.emailUser || "").trim().toLowerCase();
    if (!emailUser) return createJsonResponse({ success: false, message: "emailUser manquant." });
    var persWs = getSheetSafe(ss, "Personnel");
    var persData = persWs.getDataRange().getValues();
    for (var i = 1; i < persData.length; i++) {
      if (String(persData[i][6]||"").trim().toLowerCase() === emailUser) {
        persWs.getRange(i + 1, 12).setValue(new Date().toISOString()); // col L
        return createJsonResponse({ success: true });
      }
    }
    return createJsonResponse({ success: false, message: "Utilisateur introuvable." });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// ==========================================
// ALTERNANTS — DÉTACHEMENT BIDIRECTIONNEL
// ==========================================

function handleDetachStudentFromEnterprise(p, ss) {
  try {
    var idEtu = String(p.idEtu || "").trim().toUpperCase();
    var sheet = getSheetSafe(ss, "ETUDIANTS");
    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toUpperCase() === idEtu) { rowIndex = i + 1; break; }
    }
    if (rowIndex === -1) return createJsonResponse({ success: false, message: "Étudiant introuvable." });

    var oldRow = data[rowIndex - 1];
    var oldIdEnt = String(oldRow[11] || "").trim();
    var oldNomEnt = String(oldRow[21] || "").trim();
    if (!oldIdEnt) return createJsonResponse({ success: false, message: "Aucune entreprise liée à cet étudiant." });

    _saveAncienLien(ss, {
      idEtu: idEtu,
      nom: String(oldRow[2]||""),
      prenom: String(oldRow[3]||""),
      classe: String(oldRow[14]||""),
      campus: String(oldRow[22]||""),
      idEnt: oldIdEnt,
      nomEnt: oldNomEnt,
      motif: "Retrait manuel"
    });

    sheet.getRange(rowIndex, 12).setValue("");
    sheet.getRange(rowIndex, 13).setValue("");
    sheet.getRange(rowIndex, 22).setValue("");

    var auteur = p.emailAuteur || p.idAuteur || "?";
    var nomEtu = String(oldRow[3]||"") + " " + String(oldRow[2]||"");
    logAction(auteur, p.roleAuteur||"?", "Modification", "Dossier étudiant", idEtu,
      auteur + " a retiré l'entreprise « " + oldNomEnt + " » du dossier de " + nomEtu + " — Mouvement enregistré dans Anciens alternants");
    logAction(auteur, p.roleAuteur||"?", "Création", "AncienAlternant", idEtu,
      nomEtu + " quitte « " + oldNomEnt + " » (retrait depuis dossier étudiant)");

    return createJsonResponse({ success: true, oldEntreprise: oldNomEnt, oldIdEnt: oldIdEnt, nomEtu: nomEtu });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// ==========================================
// GED — NAVIGATION SOUS-DOSSIERS
// ==========================================

function handleGedGetSubfolder(p, ss) {
  try {
    if (!p.folderId) return createJsonResponse({ success: false, message: "folderId manquant." });
    var folder = DriveApp.getFolderById(p.folderId);
    var files = [];
    var fils = folder.getFiles();
    while (fils.hasNext()) {
      var f = fils.next();
      files.push({ id: f.getId(), name: f.getName(), url: f.getUrl(),
        mimeType: f.getMimeType(), size: f.getSize(),
        date: Utilities.formatDate(f.getLastUpdated(), Session.getScriptTimeZone(), "dd/MM/yyyy") });
    }
    files.sort(function(a,b){ return b.date.localeCompare(a.date); });
    var subfolders = [];
    var subs = folder.getFolders();
    while (subs.hasNext()) {
      var sf = subs.next();
      subfolders.push({ id: sf.getId(), name: sf.getName(), url: sf.getUrl() });
    }
    return createJsonResponse({ success: true, folderName: folder.getName(), folderId: p.folderId, files: files, subfolders: subfolders });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// ==========================================
// SALLES & SIMULATIONS
// ==========================================

function _getSallesSheet(ss) {
  var ws = ss.getSheetByName("Classe");
  if (!ws) throw new Error("Onglet 'Classe' introuvable.");
  return ws;
}

function _getSimulationsSheet(ss) {
  var ws = ss.getSheetByName("Simulations");
  if (!ws) {
    ws = ss.insertSheet("Simulations");
    ws.appendRow(["ID","Titre","DateDebut","DateFin","Data","SavedAt"]);
    ws.setFrozenRows(1);
    ws.hideSheet();
  }
  return ws;
}

function _countStudentsByClasse(ss) {
  var map = {};
  try {
    var ws = getSheetSafe(ss, "ETUDIANTS");
    var data = ws.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var classe = String(data[i][14]||"").trim();
      var statut = String(data[i][20]||"").trim();
      if (!classe || statut === "SORTIE" || statut === "PROSPECT") continue;
      map[classe] = (map[classe]||0) + 1;
    }
  } catch(e) {}
  return map;
}

function handleGetSalles(p, ss) {
  try {
    var ws = _getSallesSheet(ss);
    var data = ws.getDataRange().getValues();
    var classeCount = _countStudentsByClasse(ss);
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      var campus = String(data[i][0]||"").trim();
      var numsalle = String(data[i][1]||"").trim();
      if (!campus && !numsalle) continue;
      var capacite = parseInt(data[i][2])||0;
      var classeDebut = String(data[i][3]||"").trim();
      var effectifDebut = classeDebut ? (classeCount[classeDebut]||0) : (parseInt(data[i][4])||0);
      var classeFin = String(data[i][6]||"").trim();
      var effectifFin = classeFin ? (classeCount[classeFin]||0) : (parseInt(data[i][7])||0);
      var tauxDebut = capacite > 0 && effectifDebut > 0 ? Math.round(effectifDebut/capacite*100) : 0;
      var tauxFin   = capacite > 0 && effectifFin > 0   ? Math.round(effectifFin/capacite*100)   : 0;
      rows.push({
        id: campus + "||" + numsalle,
        campus: campus, numsalle: numsalle, capacite: capacite,
        classeDebut: classeDebut, effectifDebut: effectifDebut, tauxDebut: tauxDebut,
        classeFin: classeFin, effectifFin: effectifFin, tauxFin: tauxFin,
        commentaire: String(data[i][9]||"").trim(),
        _row: i + 1
      });
    }
    return createJsonResponse({ success: true, data: rows });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleSaveSalle(p, ss) {
  try {
    if (p.vieScRole !== "Super-admin" && p.vieScRole !== "Pédagogie")
      return createJsonResponse({ success: false, message: "Accès refusé." });
    var ws = _getSallesSheet(ss);
    var data = ws.getDataRange().getValues();
    var classeCount = _countStudentsByClasse(ss);
    var campus = String(p.campus||"").trim();
    var numsalle = String(p.numsalle||"").trim();
    var capacite = parseInt(p.capacite)||0;
    var classeDebut = String(p.classeDebut||"").trim();
    var classeFin = String(p.classeFin||"").trim();
    var effectifDebut = classeDebut ? (classeCount[classeDebut]||0) : 0;
    var effectifFin   = classeFin   ? (classeCount[classeFin]||0)   : 0;
    var tauxDebut = capacite > 0 && effectifDebut > 0 ? Math.round(effectifDebut/capacite*100) : 0;
    var tauxFin   = capacite > 0 && effectifFin   > 0 ? Math.round(effectifFin/capacite*100)   : 0;
    var commentaire = String(p.commentaire||"").trim();
    // Find existing row to update
    var oldId = String(p.id||"");
    var parts = oldId.split("||");
    var oldCampus = parts[0]||"", oldNum = parts[1]||"";
    var found = false;
    for (var i = 1; i < data.length; i++) {
      var rc = String(data[i][0]||"").trim(), rn = String(data[i][1]||"").trim();
      if (rc === oldCampus && rn === oldNum) {
        ws.getRange(i+1, 1, 1, 10).setValues([[campus, numsalle, capacite, classeDebut, effectifDebut, tauxDebut+'%', classeFin, effectifFin, tauxFin+'%', commentaire]]);
        found = true; break;
      }
    }
    if (!found) {
      ws.appendRow([campus, numsalle, capacite, classeDebut, effectifDebut, tauxDebut+'%', classeFin, effectifFin, tauxFin+'%', commentaire]);
    }
    return createJsonResponse({ success: true });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleDeleteSalle(p, ss) {
  try {
    if (p.vieScRole !== "Super-admin" && p.vieScRole !== "Pédagogie")
      return createJsonResponse({ success: false, message: "Accès refusé." });
    var ws = _getSallesSheet(ss);
    var data = ws.getDataRange().getValues();
    var parts = String(p.id||"").split("||");
    var campus = parts[0]||"", numsalle = parts[1]||"";
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]||"").trim() === campus && String(data[i][1]||"").trim() === numsalle) {
        ws.deleteRow(i+1); return createJsonResponse({ success: true });
      }
    }
    return createJsonResponse({ success: false, message: "Salle introuvable." });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function _fmtDateStr(val) {
  if (!val) return "";
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  var s = String(val).trim();
  // Si format dd/MM/yyyy, convertir
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return m[3]+"-"+m[2].padStart(2,"0")+"-"+m[1].padStart(2,"0");
  return s;
}

function handleGetSimulations(p, ss) {
  try {
    var ws = _getSimulationsSheet(ss);
    var data = ws.getDataRange().getValues();
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      rows.push({
        id: String(data[i][0]), titre: String(data[i][1]||""),
        dateDebut: _fmtDateStr(data[i][2]),
        dateFin: _fmtDateStr(data[i][3]),
        data: String(data[i][4]||"{}"), savedAt: String(data[i][5]||"")
      });
    }
    rows.reverse();
    return createJsonResponse({ success: true, data: rows });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleSaveSimulation(p, ss) {
  try {
    if (p.vieScRole !== "Super-admin" && p.vieScRole !== "Pédagogie")
      return createJsonResponse({ success: false, message: "Accès refusé." });
    var ws = _getSimulationsSheet(ss);
    var data = ws.getDataRange().getValues();
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    var id = String(p.id||"").trim();
    if (id) {
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]) === id) {
          ws.getRange(i+1, 1, 1, 6).setValues([[id, p.titre||"", p.dateDebut||"", p.dateFin||"", p.data||"{}", now]]);
          return createJsonResponse({ success: true, id: id });
        }
      }
    }
    id = "SIM-" + Date.now();
    ws.appendRow([id, p.titre||"", p.dateDebut||"", p.dateFin||"", p.data||"{}", now]);
    return createJsonResponse({ success: true, id: id });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleDeleteSimulation(p, ss) {
  try {
    if (p.vieScRole !== "Super-admin" && p.vieScRole !== "Pédagogie")
      return createJsonResponse({ success: false, message: "Accès refusé." });
    var ws = _getSimulationsSheet(ss);
    var data = ws.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === String(p.id||"")) {
        ws.deleteRow(i+1); return createJsonResponse({ success: true });
      }
    }
    return createJsonResponse({ success: false, message: "Simulation introuvable." });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// ==========================================
// SUIVI EXB
// ==========================================

function _getSuiviEXBSheet(ss) {
  var ws = ss.getSheetByName("SuiviEXB");
  if (!ws) throw new Error("Onglet 'SuiviEXB' introuvable.");
  return ws;
}

function handleGetSuiviEXB(p, ss) {
  try {
    var ws = _getSuiviEXBSheet(ss);
    var data = ws.getDataRange().getValues();
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0] && !data[i][1]) continue;
      rows.push({
        id: "EXB-" + i,
        refexb:      String(data[i][0]||"").trim(),
        classe:      String(data[i][1]||"").trim(),
        matiere:     String(data[i][2]||"").trim(),
        formateur:   String(data[i][3]||"").trim(),
        sujet:       String(data[i][4]||"").trim(),
        corrige:     String(data[i][5]||"").trim(),
        impression:  String(data[i][6]||"").trim(),
        commentaire: String(data[i][7]||"").trim(),
        sujetFileId: String(data[i][8]||"").trim(),
        sujetUrl:    String(data[i][9]||"").trim(),
        corrigeFileId:String(data[i][10]||"").trim(),
        corrigeUrl:  String(data[i][11]||"").trim(),
        _row: i + 1
      });
    }
    return createJsonResponse({ success: true, data: rows });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleSaveSuiviEXB(p, ss) {
  try {
    if (p.vieScRole !== "Super-admin" && p.vieScRole !== "Pédagogie")
      return createJsonResponse({ success: false, message: "Accès refusé." });
    var ws = _getSuiviEXBSheet(ss);
    var data = ws.getDataRange().getValues();
    var row = [
      String(p.refexb||""), String(p.classe||""), String(p.matiere||""), String(p.formateur||""),
      String(p.sujet||""), String(p.corrige||""), String(p.impression||""), String(p.commentaire||""),
      String(p.sujetFileId||""), String(p.sujetUrl||""), String(p.corrigeFileId||""), String(p.corrigeUrl||"")
    ];
    var idStr = String(p.id||"");
    if (idStr.startsWith("EXB-")) {
      var rowIdx = parseInt(idStr.slice(4));
      if (rowIdx >= 1 && rowIdx < data.length) {
        ws.getRange(rowIdx+1, 1, 1, 12).setValues([row]);
        return createJsonResponse({ success: true });
      }
    }
    ws.appendRow(row);
    return createJsonResponse({ success: true });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// Génération batch depuis le backend (lit ETUDIANTS directement → fiable même si front pas chargé)
function handleCreateEXBBatchByTypo(p, ss) {
  try {
    if (p.vieScRole !== "Super-admin" && p.vieScRole !== "Pédagogie")
      return createJsonResponse({ success: false, message: "Accès refusé." });
    var typo = String(p.typo || "").trim();
    var refexb = String(p.refexb || "").trim();
    if (!typo || !refexb) return createJsonResponse({ success: false, message: "Paramètres manquants (typo, refexb)." });

    // Lire les entrées existantes pour éviter les doublons
    var exbWs = _getSuiviEXBSheet(ss);
    var exbData = exbWs.getDataRange().getValues();
    var existing = new Set();
    for (var i = 1; i < exbData.length; i++) {
      existing.add(String(exbData[i][0]) + '||' + String(exbData[i][1]) + '||' + String(exbData[i][2]));
    }

    var toCreate = [];
    var emptyRow = function(classe, matiere) {
      return [refexb, classe, matiere, '', '', '', '', '', '', '', '', ''];
    };

    if (typo === '1') {
      // ── BTS CG : lire les classes depuis ETUDIANTS ──────────────
      var etuWs = getSheetSafe(ss, "ETUDIANTS");
      var etuData = etuWs.getDataRange().getValues();
      var btsClasses = [];
      var btsSet = new Set();
      for (var j = 1; j < etuData.length; j++) {
        var classe = String(etuData[j][14] || "").trim(); // colonne O = classe
        var sortie  = String(etuData[j][18] || "").trim(); // colonne S = sortie
        if (sortie) continue; // ignorer étudiants sortis
        if (/bts\s*cg/i.test(classe) && !btsSet.has(classe)) {
          btsSet.add(classe); btsClasses.push(classe);
        }
      }
      btsClasses.sort();
      var btsMatieres = ['Math', 'Anglais', 'CGE', 'CEJM', 'E5', 'Oraux PROS'];
      btsClasses.forEach(function(cl) {
        btsMatieres.forEach(function(mat) {
          var k = refexb + '||' + cl + '||' + mat;
          if (!existing.has(k)) toCreate.push(emptyRow(cl, mat));
        });
      });

      // ── DCG : UE 1 à UE 12, classe = "DCG" ─────────────────────
      var dcgMatieres = ['UE 1','UE 2','UE 3','UE 4','UE 5','UE 6','UE 7','UE 8','UE 9','UE 10','UE 11','UE 12'];
      dcgMatieres.forEach(function(mat) {
        var k = refexb + '||DCG||' + mat;
        if (!existing.has(k)) toCreate.push(emptyRow('DCG', mat));
      });

    } else if (typo === '2') {
      // DSCG : classes depuis ETUDIANTS + matières refMatieres
      var etuWs2 = getSheetSafe(ss, "ETUDIANTS");
      var etuData2 = etuWs2.getDataRange().getValues();
      var dscgSet = new Set();
      for (var k2 = 1; k2 < etuData2.length; k2++) {
        var cl2 = String(etuData2[k2][14] || "").trim();
        if (/dscg/i.test(cl2)) dscgSet.add(cl2);
      }
      // Matières DSCG hardcodées (UE 1-7)
      var dscgMatieres = ['UE 1','UE 2','UE 3','UE 4','UE 5','UE 6','UE 7'];
      [...dscgSet].sort().forEach(function(cl) {
        dscgMatieres.forEach(function(mat) {
          var k = refexb + '||' + cl + '||' + mat;
          if (!existing.has(k)) toCreate.push(emptyRow(cl, mat));
        });
      });

    } else if (typo === '3') {
      // Autres filières
      var etuWs3 = getSheetSafe(ss, "ETUDIANTS");
      var etuData3 = etuWs3.getDataRange().getValues();
      var otherSet = new Set();
      for (var k3 = 1; k3 < etuData3.length; k3++) {
        var cl3 = String(etuData3[k3][14] || "").trim();
        var sortie3 = String(etuData3[k3][18] || "").trim();
        if (sortie3) continue;
        if (cl3 && !/bts\s*cg/i.test(cl3) && !/dcg/i.test(cl3) && !/dscg/i.test(cl3)) otherSet.add(cl3);
      }
      [...otherSet].sort().forEach(function(cl) {
        var k = refexb + '||' + cl + '||(à définir)';
        if (!existing.has(k)) toCreate.push(emptyRow(cl, '(à définir)'));
      });

    } else {
      return createJsonResponse({ success: false, message: "Typo inconnue : " + typo });
    }

    if (!toCreate.length) return createJsonResponse({ success: true, created: 0, refexb: refexb });

    toCreate.forEach(function(row) { exbWs.appendRow(row); });
    return createJsonResponse({ success: true, created: toCreate.length, refexb: refexb });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleCreateEXBBatch(p, ss) {
  try {
    if (p.vieScRole !== "Super-admin" && p.vieScRole !== "Pédagogie")
      return createJsonResponse({ success: false, message: "Accès refusé." });
    var ws = _getSuiviEXBSheet(ss);
    var items = p.items || [];
    var created = 0;
    items.forEach(function(item) {
      ws.appendRow([String(item.refexb||""), String(item.classe||""), String(item.matiere||""), String(item.formateur||""),
        String(item.sujet||""), String(item.corrige||""), String(item.impression||""), String(item.commentaire||""),
        "","","",""]);
      created++;
    });
    return createJsonResponse({ success: true, created: created });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleUploadEXBFile(p, ss) {
  try {
    if (!p.base64Data || !p.fileName || !p.mimeType)
      return createJsonResponse({ success: false, message: "Paramètres manquants." });
    var blob = Utilities.newBlob(Utilities.base64Decode(p.base64Data), p.mimeType, p.fileName);
    var folder;
    if (p.folderId) {
      try { folder = DriveApp.getFolderById(p.folderId); } catch(e) {}
    }
    if (!folder) {
      // Créer / utiliser le dossier "EXB_Documents" dans le dossier GED COM
      try {
        var root = DriveApp.getFolderById(FOLDER_ID_COM);
        var it = root.getFoldersByName("EXB_Documents");
        folder = it.hasNext() ? it.next() : root.createFolder("EXB_Documents");
      } catch(e) {
        folder = DriveApp.getRootFolder();
      }
    }
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return createJsonResponse({ success: true, id: file.getId(), url: file.getUrl(), name: file.getName() });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleDeleteSuiviEXB(p, ss) {
  try {
    if (p.vieScRole !== "Super-admin" && p.vieScRole !== "Pédagogie")
      return createJsonResponse({ success: false, message: "Accès refusé." });
    var ws = _getSuiviEXBSheet(ss);
    var data = ws.getDataRange().getValues();
    var idStr = String(p.id||"");
    if (idStr.startsWith("EXB-")) {
      var rowIdx = parseInt(idStr.slice(4));
      if (rowIdx >= 1 && rowIdx < data.length) {
        ws.deleteRow(rowIdx+1); return createJsonResponse({ success: true });
      }
    }
    return createJsonResponse({ success: false, message: "Entrée introuvable." });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleDeleteEXBSession(p, ss) {
  try {
    if (p.vieScRole !== "Super-admin" && p.vieScRole !== "Pédagogie")
      return createJsonResponse({ success: false, message: "Accès refusé." });
    var ws = _getSuiviEXBSheet(ss);
    var data = ws.getDataRange().getValues();
    var ids = Array.isArray(p.ids) ? p.ids.map(String) : [];
    // Extraire les index de lignes (tri décroissant pour supprimer sans décalage)
    var rowsToDelete = [];
    ids.forEach(function(idStr) {
      if (idStr.startsWith("EXB-")) {
        var rowIdx = parseInt(idStr.slice(4));
        if (rowIdx >= 1 && rowIdx < data.length) rowsToDelete.push(rowIdx + 1);
      }
    });
    rowsToDelete.sort(function(a,b){return b-a;});
    rowsToDelete.forEach(function(row){ ws.deleteRow(row); });
    return createJsonResponse({ success: true, deleted: rowsToDelete.length });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// ==========================================
// MMOK — RÉSEAUX SOCIAUX
// ==========================================

function _getMMOKSheet(ss) {
  var ws = ss.getSheetByName("MMOK");
  if (!ws) throw new Error("Onglet 'MMOK' introuvable.");
  return ws;
}

function handleGetMMOK(p, ss) {
  try {
    if (p.vieScRole !== "Super-admin")
      return createJsonResponse({ success: false, message: "Accès Super-admin requis." });
    var ws = _getMMOKSheet(ss);
    var data = ws.getDataRange().getValues();
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0] && !data[i][1] && !data[i][4]) continue;
      var dp = data[i][3], dr = data[i][6];
      rows.push({
        id: "MMOK-" + i,
        refpost: String(data[i][0]||"").trim(),
        plateforme: String(data[i][1]||"").trim(),
        theme: String(data[i][2]||"").trim(),
        datePrevue: dp instanceof Date ? Utilities.formatDate(dp, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(dp||"").trim(),
        ideeContenu: String(data[i][4]||"").trim(),
        contenu: String(data[i][5]||"").trim(),
        dateReelle: dr instanceof Date ? Utilities.formatDate(dr, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(dr||"").trim(),
        etat: String(data[i][7]||"").trim(),
        _row: i + 1
      });
    }
    return createJsonResponse({ success: true, data: rows });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleSaveMMOK(p, ss) {
  try {
    if (p.vieScRole !== "Super-admin")
      return createJsonResponse({ success: false, message: "Accès Super-admin requis." });
    var ws = _getMMOKSheet(ss);
    var data = ws.getDataRange().getValues();
    var refpost = String(p.refpost||"").trim() || ("MMOK-" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmmss"));
    var row = [refpost, String(p.plateforme||""), String(p.theme||""), String(p.datePrevue||""), String(p.ideeContenu||""), String(p.contenu||""), String(p.dateReelle||""), String(p.etat||"Planifié")];
    var idStr = String(p.id||"");
    if (idStr.startsWith("MMOK-")) {
      var rowIdx = parseInt(idStr.slice(5));
      if (rowIdx >= 1 && rowIdx < data.length) {
        ws.getRange(rowIdx+1, 1, 1, 8).setValues([row]);
        return createJsonResponse({ success: true });
      }
    }
    ws.appendRow(row);
    return createJsonResponse({ success: true });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

function handleDeleteMMOK(p, ss) {
  try {
    if (p.vieScRole !== "Super-admin")
      return createJsonResponse({ success: false, message: "Accès Super-admin requis." });
    var ws = _getMMOKSheet(ss);
    var data = ws.getDataRange().getValues();
    var idStr = String(p.id||"");
    if (idStr.startsWith("MMOK-")) {
      var rowIdx = parseInt(idStr.slice(5));
      if (rowIdx >= 1 && rowIdx < data.length) {
        ws.deleteRow(rowIdx+1); return createJsonResponse({ success: true });
      }
    }
    return createJsonResponse({ success: false, message: "Post introuvable." });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}
