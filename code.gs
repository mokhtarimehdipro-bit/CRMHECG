
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
// DISPATCHER PRINCIPAL
// ==========================================
function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const action = params.action;
    const ss = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");
    var p = JSON.parse(e.postData.contents);

    switch (action) {
      case "login": return handleLogin(params, ss);
      case "getDashboardData": return handleDashboardData(params, ss);
      case "getStats": return handleGetStats(params, ss);
      case "addStudent": return handleAddStudent(params, ss);
      case "updateStudent": return handleUpdateStudent(params, ss);
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
      case "mergePartners": return handleMergePartners(params, ss);
      case "checkPartnerDuplicate": return handleCheckDuplicate(params, ss);
      case "getPartnerDetails": return handleGetPartnerDetails(params, ss);
      
      case "getAllAlerts": return handleGetAllAlerts(params, ss);
      case "completeAlert": return handleCompleteAlert(params, ss);
      case "updateAlert": return handleUpdateAlert(params, ss);
      case "cancelAlert": return handleCancelAlert(params, ss);
      case "testAlert": return handleTestAlert(params, ss);
      case "getOffres": return handleGetOffres(params, ss);
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
  for(let i=1; i<dP.length; i++) if((id == String(dP[i][0]) || id == String(dP[i][6])) && mdp == String(dP[i][4])) return createJsonResponse({ success: true, role: dP[i][5], nom: dP[i][2]+" "+dP[i][1], ref: dP[i][0] }); 
  const sE = getSheetSafe(ss, "ETUDIANTS"); const dE = sE.getDataRange().getValues(); 
  for(let j=1; j<dE.length; j++) if(id == String(dE[j][0]) && mdp == String(dE[j][1])) return createJsonResponse({ success: true, role: "Etudiant", nom: dE[j][3]+" "+dE[j][2], ref: dE[j][0] }); 
  return createJsonResponse({ success: false, message: "Identifiants incorrects." }); 
}

function calculerTypoEtu(entree, sortie, classe, forceTypo) {
  if (forceTypo === "Refuse" || forceTypo === "Attente") return forceTypo;
  if (sortie && String(sortie).trim() !== "") return "Sortie";
  if (!classe || String(classe).trim() === "") return "Prospect";
  var dEntree = new Date(); var strEntree = String(entree).toLowerCase().trim();
  if (strEntree.includes("rentree")) { var year = strEntree.replace(/[^0-9]/g, ''); if (year.length === 4) dEntree = new Date(year + "-09-01"); } 
  else if (entree) { dEntree = new Date(entree); } else { return "Prospect"; }
  if (!isNaN(dEntree.getTime())) { var now = new Date(); if (dEntree > now) return "Preinscrit"; else return "Inscrit"; }
  return "Inscrit";
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
  
  // Correction Typo "Auto" à la création
  let typoAAppliquer = p.forceTypo;
  if (!typoAAppliquer || typoAAppliquer === "Auto") {
      typoAAppliquer = calculerTypoEtu(p.entree, p.sortie, p.classe, null);
  }

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





function handleDeleteStudent(p, ss) { const ws = getSheetSafe(ss, "ETUDIANTS"); const d = ws.getDataRange().getValues(); for (let i = 1; i < d.length; i++) if (String(d[i][0]) === p.id_etu) { ws.deleteRow(i + 1); logAction(p.idAuteur, p.roleAuteur, "Suppression", "Dossier étudiant", p.id_etu, "A supprimé une fiche"); return createJsonResponse({ success: true }); } return createJsonResponse({ success: false }); }

function handleStudentDetails(params, ss) { 
  const id = params.targetId; 
  const dE = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues(); 
  let prof = null;
  for (let i = 1; i < dE.length; i++) {
    if (String(dE[i][0]).trim() === id) { 
      prof = { id: id, mdp: dE[i][1], nom: dE[i][2], prenom: dE[i][3], mail: dE[i][4], tel: dE[i][5], cv: dE[i][6], statut: dE[i][13], classe: dE[i][14], faitCV: dE[i][15], faitLI: dE[i][16], entree: dE[i][17], sortie: dE[i][18], motif: dE[i][19], typo: dE[i][20], id_ent: dE[i][11], id_cont: dE[i][12], entreprise: dE[i][21], campus: dE[i][22] };
      break; 
    }
  }
  const idUpper = String(id).trim().toUpperCase();
  const notes = getSheetSafe(ss, "Carnet_route").getDataRange().getValues().filter(r => String(r[0]).trim().toUpperCase() === idUpper).map(r => ({ auteur: r[1], date: r[2], texte: r[3] }));

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

  return createJsonResponse({ success: true, profil: prof, notes: notes.reverse(), offres: studentOffers }); 
}

function handleAddNote(p, ss) { 
  getSheetSafe(ss, "Carnet_route").appendRow([p.targetId, p.auteur, new Date(), p.contenu]); 
  logAction(p.idAuteur, p.roleAuteur, "Création", "Carnet de route", p.targetId, "A ajouté un compte-rendu/note");
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
  logAction(p.idAuteur, p.roleAuteur, "Création", "Carnet de route", p.targetId, "A ajouté une note");
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
    let typo = String(dE[i][20] || "Prospect").trim();
    if ((typo.toLowerCase().includes("prein") || typo.toLowerCase().includes("preinscrit")) && dE[i][17]) {
        if (new Date(dE[i][17]) <= today) typo = "Inscrit";
    }
    const aAlt = String(dE[i][13]).toLowerCase().includes("alternance");
    if (aAlt && dE[i][21]) entAccueil.add(String(dE[i][21]).toLowerCase().trim());
    
    // NOUVEAU : On injecte le campus dans l'objet Etudiant pour les stats
    const etuObj = { id: dE[i][0], nom: dE[i][2] + " " + dE[i][3], classe: dE[i][14], aAlternance: aAlt, motif: dE[i][19], email: dE[i][4], campus: dE[i][22] };
    
    if (typo.toLowerCase().includes("sorti") || typo.toLowerCase().includes("refus")) stats.sorties.push(etuObj); 
    else if (typo.toLowerCase() === "prospect") stats.prospects.push(etuObj); 
    else if (typo.toLowerCase().includes("prein") || typo.toLowerCase().includes("préin")) stats.preinscrits.push(etuObj);
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
    if (!entreprises[idEnt]) entreprises[idEnt] = { id: idEnt, nom: dP[i][2], contacts: [], alternants: [] };
    entreprises[idEnt].contacts.push({ id_contact: dP[i][0], nom: dP[i][3], prenom: dP[i][4], tel: dP[i][5], email: dP[i][6], tel2: dP[i][9], mail2: dP[i][8] });
  }
  for (let j = 1; j < dE.length; j++) {
    const idEntEtu = String(dE[j][11]).trim();
    if (entreprises[idEntEtu] && String(dE[j][13]).toLowerCase().includes("alternance")) entreprises[idEntEtu].alternants.push({ nom: dE[j][2], prenom: dE[j][3] });
  }
  return createJsonResponse({ success: true, data: Object.values(entreprises) });
}

// ==========================================
// OFFRES
// ==========================================
function handleGetOffres(p, ss) {
  try {
    const dO = getSheetSafe(ss, "OFFRES").getDataRange().getValues();
    const dP = getSheetSafe(ss, "PARTAGES").getDataRange().getValues();
    const dE = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues();
    
    // Dictionnaire ultra-précis : on indexe par ID ET par Email
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
             // On cherche par ID (majuscule) ou par Mail (minuscule)
             return etuMap[key.toUpperCase()] || etuMap[key.toLowerCase()] || ("Inconnu (" + key + ")");
          });

        res.push({
          id: idOffre, url: dO[i][1], date: dO[i][2], plateforme: dO[i][3],
          entreprise: dO[i][4], missions: dO[i][6], 
          qualite: dO[i][9], // Colonne J (Index 9)
          etat: dO[i][13],   // Colonne N (Index 13)
          poste: dO[i][14],
          destinataires: sharedWith
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
    p.type,             // F : Type
    p.missions,         // G : Missions
    "",                 // H
    "",                 // I
    p.qualite || "À qualifier", // J : Qualité_offre (Index 9)
    "",                 // K
    "",                 // L
    "",                 // M
    p.etat || "Non partagée",   // N : Etat_offre (Index 13)
    p.poste             // O : Poste
  ]); 

  logAction(p.idAuteur, p.roleAuteur, "Création", "Offres", "Nouvelle Offre", "A ajouté l'offre : " + p.poste);
  return createJsonResponse({ success: true }); 
}



// --- MISE À JOUR D'UNE OFFRE EXISTANTE ---
function handleUpdateOffre(p, ss) {
  try {
    const sheet = getSheetSafe(ss, "OFFRES");
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(p.id)) {
        
        // 1. Mise à jour des informations classiques
        sheet.getRange(i + 1, 2).setValue(p.url);          // Col B (Index 2)
        sheet.getRange(i + 1, 4).setValue(p.plateforme);   // Col D (Index 4)
        sheet.getRange(i + 1, 5).setValue(p.entreprise);   // Col E (Index 5)
        sheet.getRange(i + 1, 6).setValue(p.type);         // Col F (Index 6)
        sheet.getRange(i + 1, 7).setValue(p.missions);     // Col G (Index 7)
        sheet.getRange(i + 1, 15).setValue(p.poste);       // Col O (Index 15)

        // 2. LA LOGIQUE INVERSÉE (COMME VOUS L'AVEZ DEMANDÉ)
        // La Qualité (Validé, Obsolète...) remonte bien dans la Colonne J (Index 10)
        sheet.getRange(i + 1, 10).setValue(p.qualite);
        
        // L'État (Partagé, Non partagé...) remonte bien dans la Colonne N (Index 14)
        sheet.getRange(i + 1, 14).setValue(p.etat);
        
        logAction(p.idAuteur, p.roleAuteur, "Modification", "Offres", p.id, "A modifié l'offre");
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

  try {
    const chunkSize = 50; 
    for (let i = 0; i < emailsArray.length; i += chunkSize) {
      options.bcc = emailsArray.slice(i, i + chunkSize).join(',');
      GmailApp.sendEmail(Session.getActiveUser().getEmail(), subject, "", options);
    }
    // Mise à jour du statut UNIQUEMENT si ce n'est pas un envoi de test
    if (!isTest) {
      const ws = getSheetSafe(ss, "CAMPAGNES_COM"); 
      const d = ws.getDataRange().getValues();
      for(let i = 1; i < d.length; i++) { 
        if(String(d[i][0]) === cmpId) { 
          ws.getRange(i+1, 7).setValue(new Date()); // Col G : Date d'envoi
          ws.getRange(i+1, 8).setValue("Diffusé");  // Col H : Changement du statut
          break; 
        } 
      }
    }
    logAction(p.idAuteur, p.roleAuteur, "Action", "Campagne", cmpId, "A diffusé la campagne à " + emailsArray.length + " contacts");
    return createJsonResponse({ 
      success: true, 
      message: isTest ? "Mail de test envoyé !" : `Campagne diffusée avec succès à ${emailsArray.length} contacts.` 
    });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
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
  logAction(p.idAuteur, p.roleAuteur, isNew ? "Création" : "Modification", "Com", id, "A sauvegardé l'événement");
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
    // On ne prend que ceux qui ont un mail et on formate pour la liste
    if (data[i][4]) { 
      list.push({
        nom: data[i][1],
        prenom: data[i][2],
        email: data[i][4],
        classe: data[i][5] || "Sans classe"
      });
    }
  }
  // Tri alphabétique par nom
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
          if (motifSelectionne === "Tous" || motifEtudiant === motifSelectionne) ok = true;
        }
      } else {
        // --- MODE CLASSIQUE (On exclut d'office les sortis et refusés) ---
        if (!typo.includes("sorti") && !typo.includes("refus")) {
          
          // 1. Etudiants en recherche : (Statut Recherche) ET (Inscrit OU Prein)
          if (cnf.target_recherche && status.includes("recherche") && (typo.includes("inscrit") || typo.includes("prein"))) {
            ok = true;
          }
          
          // 2. Etudiants préinscrits uniquement (Peu importe le statut)
          if (cnf.b2c_preinscrits && typo.includes("prein")) {
            ok = true;
          }
          
          // 3. Etudiants prospects uniquement
          if (cnf.b2c_prospects && typo.includes("prospect")) {
            ok = true;
          }
          
          // 4. Inscrits définitivement (Tous les inscrits, peu importe le statut)
          if (cnf.b2c_inscrits && (typo.includes("inscrit") || typo.includes("prein"))) ok = true;
          
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
      // Au lieu de supprimer la ligne (ws.deleteRow), on change le statut en colonne H (index 8)
      ws.getRange(i + 1, 8).setValue("Archivé");
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

    if (statut.includes("recherche") && (typo.includes("inscrit") || typo.includes("prein"))) {
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
    
    // Filtre : Recherche + (Inscrit ou Preinscrit)
    if (statut.includes("recherche") && (typo.includes("inscrit") || typo.includes("prein"))) {
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
        if(dA[i][4] === "Unique") {
          sA.getRange(i+1, 6).setValue("Terminé");
        } else {
          // Si c'est récurrent, on repousse la date
          let oldDate = new Date(dA[i][2]);
          if(dA[i][4] === "Hebdomadaire") oldDate.setDate(oldDate.getDate() + 7);
          if(dA[i][4] === "Mensuelle") oldDate.setMonth(oldDate.getMonth() + 1);
          if(dA[i][4] === "Trimestrielle") oldDate.setMonth(oldDate.getMonth() + 3);
          sA.getRange(i+1, 3).setValue(oldDate);
        }
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
        sA.getRange(i+1, 3).setValue(p.date);
        sA.getRange(i+1, 4).setValue(p.msg);
        sA.getRange(i+1, 5).setValue(p.freq);
        sA.getRange(i+1, 7).setValue(p.destinataire); // <-- ECRITURE DE L'EMAIL ICI
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
        sA.getRange(i+1, 6).setValue("Annulé");
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
// CRÉATION D'UNE ALERTE DEPUIS LE DOSSIER PARTENAIRE
// ==========================================
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
    logAction(p.idAuteur, p.roleAuteur, "Création", "Alertes", idPartenaire, "A programmé une alerte");
    return createJsonResponse({ success: true, message: "Alerte créée avec succès dans la base !" });
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
    
    for (let i = 1; i < dP.length; i++) {
      if (String(dP[i][1]).trim().toUpperCase() === id) {
        info = { entreprise: dP[i][2], contact: dP[i][3] + " " + dP[i][4], mail: dP[i][6], tel: dP[i][5] };
        break;
      }
    }
    
    if (!info) return createJsonResponse({ success: false, message: "Entreprise introuvable" });
    
    const dataEtu = getSheetSafe(ss, "ETUDIANTS").getDataRange().getValues();
    const alternants = dataEtu.filter(r => String(r[11]).trim().toUpperCase() === id).map(r => ({ nom: r[3] + " " + r[2], classe: r[14] }));
    
    // --- LA CORRECTION EST ICI : ON UTILISE SUIVI_PARTENARIAT ---
    const suivi = getSheetSafe(ss, "SUIVI_PARTENARIAT").getDataRange().getValues()
      .filter(r => String(r[0]).trim().toUpperCase() === id)
      .map(r => ({ 
        auteur: r[1], // Colonne B
        date: r[2],   // Colonne C
        texte: r[3]   // Colonne D (Compte_rendu)
      }))
      .reverse(); // Pour avoir les plus récents en haut
    
    const alertes = getSheetSafe(ss, "ALERTES_PARTENAIRES").getDataRange().getValues().filter(r => String(r[1]).trim().toUpperCase() === id).map(r => ({ date: r[2], msg: r[3], freq: r[4], statut: r[5] }));
    const offres = getSheetSafe(ss, "OFFRES").getDataRange().getValues().filter(r => String(r[4]).trim().toLowerCase() === info.entreprise.toLowerCase()).map(r => ({ poste: r[14], url: r[1], etat: r[9] }));
    
    return createJsonResponse({ success: true, data: { info, alternants, suivi, alertes, offres } });
  } catch (e) { 
    return createJsonResponse({ success: false, message: e.toString() }); 
  }
}
function handleUpdateEnterprise(p, ss) {
  try {
    const sheet = getSheetSafe(ss, "PARTENARIAT");
    const data = sheet.getDataRange().getValues();
    const idCible = String(p.id_entreprise).trim().toUpperCase();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim().toUpperCase() === idCible) {
        sheet.getRange(i + 1, 3).setValue(p.nouveauNom); // Mise à jour Colonne C
        return createJsonResponse({ success: true });
      }
    }
    logAction(p.idAuteur, p.roleAuteur, "Modification", "Entreprise", idCible, "A renommé l'entreprise en " + p.nouveauNom);
    return createJsonResponse({ success: false, message: "ID non trouvé" });
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

    for (let i = 1; i < dataEtu.length; i++) {
      let id = dataEtu[i][0];
      let prenom = dataEtu[i][3];
      let mail = String(dataEtu[i][4]).trim();
      let statut = String(dataEtu[i][13]).toLowerCase(); // Colonne N
      let classe = String(dataEtu[i][14]).trim(); // Colonne O
      let typo = String(dataEtu[i][20]).toLowerCase(); // Colonne U

      // RÈGLE STRICTE (Le videur)
      let estEnRecherche = statut.includes("recherche");
      let estLegitime = typo.includes("inscrit") || typo.includes("préin") || typo.includes("prein");
      let estExclu = typo.includes("sorti") || typo.includes("refus") || statut.includes("alternance") || statut.includes("initial") || statut.includes("placé");

      if (estEnRecherche && estLegitime && !estExclu && mail.includes('@')) {
        // FILTRE PAR CLASSE
        let matchClasse = true;
        if (p.classesCibles && p.classesCibles.length > 0 && !p.classesCibles.includes("TOUS")) {
          matchClasse = p.classesCibles.includes(classe);
        }

        if (matchClasse) cibles.push({ id: id, mail: mail, prenom: prenom });
      }
    }

    if (cibles.length === 0) return createJsonResponse({ success: false, message: "Aucun étudiant ne correspond aux critères (Inscrit/Préinscrit en recherche de cette classe)." });

    let nbAjoutes = 0;
    cibles.forEach(etu => {
      const dejaFait = dataPartages.some(r => String(r[0]).trim() === String(etu.id).trim() && String(r[1]).trim() === String(p.id_offre).trim());
      if (!dejaFait) {
        sheetPartages.appendRow([etu.id, p.id_offre, now, "Partagée"]);
        try {
          const sujet = "Nouvelle offre ciblée : " + infoMail.poste;
          const corps = `<div style="font-family:Arial; padding:20px; border:1px solid #eee; border-radius:15px;"><h2 style="color:#1A4E8A;">Bonjour ${etu.prenom},</h2><p>Une nouvelle offre a été sélectionnée pour ton profil :</p><div style="background:#f9f9f9; padding:15px; border-left:5px solid #F26522; margin:20px 0;"><b style="font-size:16px;">${infoMail.poste}</b><br><span>Entreprise : ${infoMail.ent}</span></div><p>Tu peux postuler directement depuis ton dossier étudiant HECG.</p><a href="${infoMail.url}" style="display:inline-block; padding:12px 25px; background:#F26522; color:white; text-decoration:none; border-radius:8px; font-weight:bold;">Voir l'annonce</a></div>`;
          GmailApp.sendEmail(etu.mail, sujet, "", { htmlBody: corps + getSignatureHTML() });
        } catch(e) {}
        nbAjoutes++;
      }
    });

    if (nbAjoutes > 0) {
      for (let i = 1; i < dataOffres.length; i++) {
        if (String(dataOffres[i][0]).trim() === String(p.id_offre).trim()) {
          sheetOffres.getRange(i + 1, 14).setValue("Partagée"); // Colonne N
          break;
        }
      }
    }
    logAction(p.idAuteur, p.roleAuteur, "Action", "Offres", p.id_offre, "A partagé l'offre avec " + nbAjoutes + " étudiant(s)");
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
        sheet.deleteRow(i + 1);
        // On nettoie aussi les partages liés à cette offre
        const sheetP = getSheetSafe(ss, "PARTAGES");
        const dataP = sheetP.getDataRange().getValues();
        for (let j = dataP.length - 1; j >= 1; j--) {
          if (String(dataP[j][1]) === String(p.id)) sheetP.deleteRow(j + 1);
        }
        logAction(p.idAuteur, p.roleAuteur, "Suppression", "Offres", p.id, "A supprimé l'offre");
        return createJsonResponse({ success: true });
      }
    }
    return createJsonResponse({ success: false });
  } catch(e) { return createJsonResponse({ success: false, message: e.toString() }); }
}

// --- CRÉATION D'UNE ENTREPRISE ET D'UN CONTACT ---
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
      "",               // Colonne H (généralement vide ou statut)
      p.mail2 || "",    // Colonne I
      p.tel2 || ""      // Colonne J
    ];
    
    sheetPart.appendRow(nouvelleLigne);

    // 5. On laisse une trace dans l'historique
    // 5. On laisse une trace dans l'historique (avec sécurité sur le nom)
    const auteurReconnu = p.idAuteur || p.auteur || "Inconnu";
    const roleReconnu = p.roleAuteur || p.role || "N/A";
    logAction(auteurReconnu, roleReconnu, "Création", "Entreprise", idEnt, "A ajouté un contact / entreprise");
    
    return createJsonResponse({ success: true, message: "Interlocuteur (et entreprise) ajouté avec succès !" });

  } catch (e) {
    return createJsonResponse({ success: false, message: "Erreur d'ajout : " + e.toString() });
  }
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

function toutAspirer() {
  console.log("🚀 Lancement de l'aspiration globale...");
  
  try {
    aspirerFranceTravail();
  } catch(e) { console.error("Erreur France Travail : " + e.message); }
  
  try {
    aspirerLinkedIn();
  } catch(e) { console.error("Erreur LinkedIn : " + e.message); }

  try {
    AspirerOffresGmail();
  } catch(e) { console.error("Erreur MMOK GOOGLE ALERTES : " + e.message); }
  
  try {
    aspirerGoogleAlerts();
  } catch(e) { console.error("Erreur Google Alerts : " + e.message); }
  
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
    let etat = String(data[i][9]).toLowerCase().trim();    // Colonne J : Etat_offre
    let qualite = String(data[i][13]).toLowerCase().trim();// Colonne N : Qualité_offre
    
    // On ignore les offres obsolètes pour le compte à traiter
    if (!etat.includes("obsolète") && !etat.includes("obsolete")) {
      
      // On compte les "A qualifier"
      if (etat.includes("qualifier") || etat === "") {
        nbAQualifier++;
      }
      
      // On compte les "Non partagée"
      if (qualite.includes("non partagée") || qualite.includes("non partage") || qualite === "") {
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
    } catch (e) {
      nbErreurs++;
      erreurDetails.push(mailEtu + " → " + e.toString());
      Logger.log("Erreur d'envoi pour " + mailEtu + " : " + e.toString());
    }
  }

  Logger.log("Relances terminées : " + nbMailsEnvoyes + " mail(s) envoyé(s), " + nbIgnores + " ignoré(s), " + nbErreurs + " erreur(s).");
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

function logAction(idAuteur, roleAuteur, typeAction, categorie, idCible, details) {
  try {
    const ss = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");
    let sheetLog;
    try {
      sheetLog = getSheetSafe(ss, "HISTORIQUE");
    } catch (e) {
      // Crée l'onglet HISTORIQUE s'il est absent
      sheetLog = ss.insertSheet("HISTORIQUE");
      sheetLog.appendRow(["ID Action", "Date", "Auteur", "Rôle", "Type", "Catégorie", "Cible", "Détails"]);
      sheetLog.setFrozenRows(1);
    }
    const idAction = "LOG-" + Date.now() + "-" + Math.floor(Math.random() * 1000);
    sheetLog.appendRow([idAction, new Date(), idAuteur || "Inconnu", roleAuteur || "N/A", typeAction, categorie, idCible || "-", details || ""]);
  } catch (e) { console.error("Erreur Mouchard : " + e.toString()); }
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
 * RÉCUPÉRATION DE L'HISTORIQUE (Réservé aux Administrateurs)
 */
function handleGetHistory(p, ss) {
  // Sécurité double : on vérifie le rôle envoyé
  if (p.roleAuteur !== "Administrateur") {
    return createJsonResponse({ success: false, message: "Accès refusé" });
  }

  try {
    const sheetLog = getSheetSafe(ss, "HISTORIQUE");
    const data = sheetLog.getDataRange().getValues();
    
    if (data.length <= 1) return createJsonResponse({ success: true, logs: [] });

    // On récupère les titres (ligne 1) et les données (lignes suivantes)
    const headers = data[0];
    const rows = data.slice(1).reverse(); // On inverse pour avoir les plus récents en haut

    const logs = rows.map(r => {
      let obj = {};
      headers.forEach((h, i) => obj[h] = r[i]);
      return obj;
    });

    return createJsonResponse({ success: true, logs: logs });
  } catch (e) {
    return createJsonResponse({ success: false, message: e.toString() });
  }
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
          wsOffres.appendRow([idOffre, dateJour, "Alerte Google Job", "Import Auto", "À traiter", "", "", "", "", "", "", "", "", "Nouvelle"]);
          
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

    const nouvelleLigne = [
      params.targetId,       
      auteurFinal,         
      new Date(),            
      params.texte           
    ];

    sheet.appendRow(nouvelleLigne);

    // On passe les variables sécurisées au mouchard
    logAction(auteurFinal, roleFinal, "Création", "Note Partenaire", params.targetId, "A ajouté un échange");

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
  let actualIdInSheet = idCherche; // Sera écrasé par la valeur réelle du sheet dans Bloc 1

  // Bloc 1 (critique) : écriture des données étudiant. Un échec ici bloque la réponse.
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
    actualIdInSheet = String(oldRow[0]); // ID RÉEL tel que stocké dans la feuille (casse d'origine)

    let typoFinale = p.forceTypo;
    if (!typoFinale || typoFinale === "Auto") {
      typoFinale = calculerTypoEtu(p.entree, p.sortie, p.classe, null);
    }

    const updatedRow = [
      oldRow[0],           // A : id_etu
      p.mdp || oldRow[1],  // B : MDP
      p.nom,               // C : Nom_etu
      p.prenom,            // D : Prenom_etu
      p.mail,              // E : Mail_etu
      p.tel,               // F : Tel_etu
      p.cv,                // G : CV
      p.refPerso,          // H : Ref_Perso
      oldRow[8],           // I : (préservé)
      oldRow[9],           // J : (préservé)
      oldRow[10],          // K : (préservé)
      p.id_ent || "",      // L : id_entreprise
      p.id_cont || "",     // M : id_contact
      p.statut,            // N : Statut
      p.classe,            // O : Classe
      p.faitCV || "Non",   // P : Atelier_CV
      p.faitLI || "Non",   // Q : Atelier_LinkedIn
      "'" + p.entree,      // R : Entree_etu
      p.sortie,            // S : sortie_etu
      p.motifSortie,       // T : motifsortie
      typoFinale,          // U : Typo
      p.entreprise || "",  // V : Nom_entreprise
      p.campus || "",      // W : Campus
      oldRow[23]           // X : Identifiant Elève
    ];

    sheet.getRange(rowIndex, 1, 1, updatedRow.length).setValues([updatedRow]);
  } catch (erreur) {
    return createJsonResponse({ success: false, message: erreur.toString() });
  }

  // Bloc 2 (non-critique) : journalisation dans HISTORIQUE uniquement.
  // Le Carnet de route est réservé aux échanges manuels (handleAddNote).
  try {
    const msgLog = p.detailsLog || "Mise à jour manuelle de la fiche";
    logAction(p.idAuteur, p.roleAuteur, "Modification", "Dossier étudiant", actualIdInSheet, msgLog);
  } catch (logErreur) {
    console.error("Erreur de journalisation (non bloquante) : " + logErreur.toString());
  }

  return createJsonResponse({ success: true, message: "Mise à jour réussie" });
}
