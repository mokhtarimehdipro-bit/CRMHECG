// ═══════════ IMPORT PONCTUEL — leads « Salon de l'Étudiant de Marseille » du 21/11/2026 ═══════════
// Source : 222HECG_leads_export-2.xlsx (export des badges scannés). Contacts corrigés d'après les
// commentaires des commerciaux (coordonnées de l'élève plutôt que celles du parent / du badge).
// Passe par handleAddProspectsSalon (même circuit que salon.html : dédoublonnage mail/tél, colonne
// « Événements », carnet de route, mouchard), puis recale les dates au 21/11/2026 (heure du scan).
// Lancé une seule fois par auditAdressesTick (garde : propriété IMPORT_LEADS_SALON_2111).
var LEADS_SALON_2111 = [
{
"cid": "LEADS2111-2",
"prenom": "Léana",
"nom": "MEGHNI",
"mail": "meghnileana@gmail.com",
"tel": "0749492466",
"campus": "HECG",
"heure": "16:35",
"note": "Profil : Lycéen — niveau : Terminale\nDomaines : Je ne sais pas quoi faire\nFormations envisagées : Classes préparatoires, Ecoles d'ingénieurs\nRencontré par : Myriam NETTAH à 16h35\nContact du badge : 0680727832 / sarah-sev@hotmail.fr\nCode postal : 13270"
},
{
"cid": "LEADS2111-3",
"prenom": "Axel",
"nom": "BRUN",
"mail": "axel.brun091209@gmail.com",
"tel": "0619821319",
"campus": "HECG",
"heure": "16:07",
"note": "Profil : Lycéen\nDomaines : Banque assurance finance immobilier, Commerce vente marketing, Comptabilité gestion ressources humaines\nRencontré par : Myriam NETTAH à 16h07\nCode postal : 84120"
},
{
"cid": "LEADS2111-4",
"prenom": "Maëva",
"nom": "CARVIN",
"mail": "maevacarvin066@gmail.com",
"tel": "0766964956",
"campus": "HECG",
"heure": "15:54",
"note": "Profil : Lycéen\nDomaines : Commerce vente marketing\nRencontré par : Myriam NETTAH à 15h54\nContact du badge : 0648371821 / sophieferretti@orange.fr\nCode postal : 13127\nBadge scanné au nom de Sophie FERRETTI"
},
{
"cid": "LEADS2111-6",
"prenom": "Lilly",
"nom": "MEYNIER",
"mail": "lilly.meynier@gmail.com",
"tel": "0665143835",
"campus": "HECG",
"heure": "15:53",
"note": "Profil : Parent — niveau : 3ème;Seconde Générale et Technologique\nDomaines : Banque assurance finance immobilier, Comptabilité gestion ressources humaines\nFormations envisagées : BTS, BUT, Ecoles de commerce - IEP, Université - licence, Université - master\nRencontré par : Rayan BEN JEDDOU à 15h53\nContact du badge : 0684114593 / vandre13@wanadoo.fr\nCode postal : 13590\nPrénom déduit de l'adresse mail — 2 badges scannés (élève + parent). Parent : Virginie MEYNIER"
},
{
"cid": "LEADS2111-7",
"prenom": "Louna",
"nom": "BEDERIAN",
"mail": "lounabede@outlook.fr",
"tel": "0788389891",
"campus": "HECG",
"heure": "15:50",
"note": "Profil : Lycéen — niveau : Terminale technologique\nDomaines : Je ne sais pas quoi faire\nFormations envisagées : BTS\nRencontré par : Myriam NETTAH à 15h50\nCode postal : 05300"
},
{
"cid": "LEADS2111-8",
"prenom": "Hugo",
"nom": "SORAZU",
"mail": "hugosorazu@gmail.com",
"tel": "0761752020",
"campus": "HECG",
"heure": "15:41",
"note": "Profil : Lycéen — niveau : Seconde Générale et Technologique\nDomaines : Je ne sais pas quoi faire\nFormations envisagées : BTS, BUT\nRencontré par : Myriam NETTAH à 15h41\nCode postal : 13190\nTél. noté sur le stand : 0761752030"
},
{
"cid": "LEADS2111-9",
"prenom": "Romane",
"nom": "BARD",
"mail": "r.bard2a@gmail.com",
"tel": "0760631544",
"campus": "HECG",
"heure": "15:32",
"note": "Profil : Lycéen — niveau : Terminale Pro\nDomaines : Armée défense sécurité, Commerce vente marketing, Communication information journalisme, Comptabilité gestion ressources humaines\nFormations envisagées : BTS\nRencontré par : Myriam NETTAH à 15h32\nContact du badge : crt.bard@gmail.com\nCode postal : 13840"
},
{
"cid": "LEADS2111-10",
"prenom": "Candice",
"nom": "FRANCESCHI",
"mail": "23candice10@gmail.com",
"tel": "0636096097",
"campus": "HECG",
"heure": "15:20",
"note": "Profil : Parent — niveau : Terminale\nDomaines : Je ne sais pas quoi faire\nFormations envisagées : BTS\nRencontré par : Rayan BEN JEDDOU à 15h20\nContact du badge : 0622334947 / carole.franceschi@yahoo.fr\nCode postal : 13960\nPrénom déduit du mail, nom du parent (à confirmer). Parent : Carole FRANCESCHI"
},
{
"cid": "LEADS2111-11",
"prenom": "Lorena",
"nom": "BLEIN",
"mail": "lorenablein13@gmail.com",
"tel": "0670789901",
"campus": "HECG",
"heure": "15:16",
"note": "Profil : Lycéen — niveau : Seconde Générale et Technologique\nDomaines : Action sociale animation sciences sociales, Banque assurance finance immobilier, Comptabilité gestion ressources humaines, Santé paramédical, Tourisme hôtellerie restauration\nFormations envisagées : Alternance, Ecoles Spécialisées, Université - licence, Université - master\nRencontré par : Myriam NETTAH à 15h16\nCode postal : 84240"
},
{
"cid": "LEADS2111-12",
"prenom": "Eliott",
"nom": "GALLET",
"mail": "eliott.gallet007@gmail.com",
"tel": "0786644689",
"campus": "HECG",
"heure": "15:14",
"note": "Profil : Lycéen\nDomaines : Droit science politique administration publique\nRencontré par : Myriam NETTAH à 15h14\nCode postal : 13090"
},
{
"cid": "LEADS2111-12b",
"prenom": "Nikolas",
"nom": "BISSON",
"mail": "bisson.nikolas@gmail.com",
"tel": "0616613419",
"campus": "HECG",
"heure": "15:14",
"note": "Venu avec Eliott GALLET (badge de ce dernier)\nRencontré par : Myriam NETTAH"
},
{
"cid": "LEADS2111-13",
"prenom": "Sarah",
"nom": "MOREL",
"mail": "sarahmrl0605@gmail.com",
"tel": "0651756513",
"campus": "HECG",
"heure": "15:08",
"note": "Profil : Lycéen — niveau : Terminale\nDomaines : Je ne sais pas quoi faire\nFormations envisagées : Ecoles d'ingénieurs, Ecoles de commerce - IEP, Ecoles Spécialisées\nRencontré par : Rayan BEN JEDDOU à 15h08\nCode postal : 83470"
},
{
"cid": "LEADS2111-14",
"prenom": "Cassandre",
"nom": "BOIZIS",
"mail": "cassandre.boizis13320@gmail.com",
"tel": "0674412726",
"campus": "HECG",
"heure": "14:40",
"note": "Profil : Parent — niveau : Première technologique\nDomaines : Banque assurance finance immobilier, Commerce vente marketing\nFormations envisagées : Alternance, BUT\nRencontré par : Myriam NETTAH à 14h40\nContact du badge : 0674089410 / lolo1976.uguen@gmail.com\nCode postal : 13320\nNom déduit du mail. Parent : Laurence UGUEN"
},
{
"cid": "LEADS2111-15",
"prenom": "Lili",
"nom": "ECK",
"mail": "lili.eck@icloud.com",
"tel": "0636159497",
"campus": "HECG",
"heure": "14:24",
"note": "Profil : Lycéen\nDomaines : Action sociale animation sciences sociales, Banque assurance finance immobilier, Commerce vente marketing, Comptabilité gestion ressources humaines, Economie\nFormations envisagées : Classes préparatoires\nRencontré par : Rayan BEN JEDDOU à 14h24\nContact du badge : 0661772338 / lili.eck@icloud.fr\nCode postal : 99999"
},
{
"cid": "LEADS2111-16",
"prenom": "Eva",
"nom": "ROBLES",
"mail": "eva09robles@gmail.com",
"tel": "0680142116",
"campus": "HECG",
"heure": "14:02",
"note": "Profil : Lycéen — niveau : Première\nDomaines : Je ne sais pas quoi faire\nFormations envisagées : Alternance, Bachelor, BTS, BUT, Ecoles d'ingénieurs, Université - master\nRencontré par : Rayan BEN JEDDOU à 14h02\nCode postal : 83640"
},
{
"cid": "LEADS2111-17",
"prenom": "Alicia",
"nom": "KALDIRIMDJIAN",
"mail": "ali6kld@gmail.com",
"tel": "0646365173",
"campus": "HECG",
"heure": "14:02",
"note": "Profil : Lycéen\nDomaines : Je ne sais pas quoi faire\nFormations envisagées : Je ne sais pas quoi faire\nRencontré par : Myriam NETTAH à 14h02\nContact du badge : 0674497560 / ali6kld@gmai.com\nCode postal : 13770"
},
{
"cid": "LEADS2111-18",
"prenom": "Raphaël",
"nom": "ISOARD",
"mail": "raphael.isoard2010@icloud.com",
"tel": "0631882076",
"campus": "HECG",
"heure": "13:40",
"note": "Profil : Lycéen — niveau : Première\nDomaines : Armée défense sécurité, Banque assurance finance immobilier, Commerce vente marketing, Economie, Mécanique aéronautique automobile\nFormations envisagées : Alternance, BTS, BUT, Ecoles d'ingénieurs, Ecoles de commerce - IEP\nRencontré par : Rayan BEN JEDDOU à 13h40\nContact du badge : 0675000401 / castaing.marjorie85@gmail.com\nCode postal : 13120"
},
{
"cid": "LEADS2111-19",
"prenom": "Chloé Léa",
"nom": "SAGGIORO",
"mail": "saggiorochloe@gmail.com",
"tel": "0611527119",
"campus": "HECG",
"heure": "12:42",
"note": "Profil : Lycéen\nDomaines : Je ne sais pas quoi faire\nFormations envisagées : Je ne sais pas quoi faire\nRencontré par : Rayan BEN JEDDOU à 12h42\nCode postal : 13008"
},
{
"cid": "LEADS2111-20",
"prenom": "Lucas",
"nom": "TRAVERSO",
"mail": "lucastraverso290909@gmail.com",
"tel": "0781245863",
"campus": "HECG",
"heure": "12:36",
"note": "Profil : Lycéen\nDomaines : Je ne sais pas quoi faire\nFormations envisagées : Je ne sais pas quoi faire\nRencontré par : Rayan BEN JEDDOU à 12h36\nContact du badge : 0668584787\nCode postal : 13008"
},
{
"cid": "LEADS2111-21",
"prenom": "Sania",
"nom": "BENALLOU",
"mail": "sania.benallou31@gmail.com",
"tel": "0659048995",
"campus": "HECG",
"heure": "12:18",
"note": "Profil : Lycéen — niveau : Terminale technologique\nDomaines : Banque assurance finance immobilier, Commerce vente marketing, Communication information journalisme, Economie, International études à l'étranger, Luxe, Tourisme hôtellerie restauration\nFormations envisagées : BTS, BUT, Ecoles Spécialisées\nRencontré par : Myriam NETTAH à 12h18\nContact du badge : 0665685798\nCode postal : 13710"
},
{
"cid": "LEADS2111-22",
"prenom": "Mateo",
"nom": "PINO BRUGUIER",
"mail": "pinobruguier.mateo@yahoo.com",
"tel": "0613511898",
"campus": "HECG",
"heure": "12:18",
"note": "Profil : Lycéen — niveau : Terminale\nDomaines : Arts audiovisuel culture, Banque assurance finance immobilier, Luxe\nFormations envisagées : BUT, Ecoles de commerce - IEP\nRencontré par : Rayan BEN JEDDOU à 12h18\nCode postal : 13540"
},
{
"cid": "LEADS2111-23",
"prenom": "Gabrielle",
"nom": "THERY",
"mail": "gabytld2009@outlook.fr",
"tel": "0749459465",
"campus": "HECG",
"heure": "12:05",
"note": "Profil : Lycéen — niveau : Terminale\nDomaines : Je ne sais pas quoi faire\nFormations envisagées : Je ne sais pas quoi faire\nMotivation : 5/5\nRencontré par : Myriam NETTAH à 12h05\nCode postal : 04100"
},
{
"cid": "LEADS2111-24",
"prenom": "Andy",
"nom": "CALMON",
"mail": "andycalmon3@gmail.com",
"tel": "0629923797",
"campus": "HECG",
"heure": "12:04",
"note": "Profil : Parent — niveau : Terminale\nDomaines : Banque assurance finance immobilier\nFormations envisagées : Alternance, BTS, Université - master\nRencontré par : Rayan BEN JEDDOU à 12h04\nContact du badge : 0610323894 / ivatzarevska@hotmail.com\nCode postal : 13700\nIdentité déduite du mail. Parent : Iva TZAREVSKA"
},
{
"cid": "LEADS2111-25",
"prenom": "Noah",
"nom": "DELMER",
"mail": "mauricedelmer0@gmail.com",
"tel": "0788476863",
"campus": "HECG",
"heure": "11:55",
"note": "Profil : Lycéen — niveau : Première\nDomaines : Je ne sais pas quoi faire\nFormations envisagées : Je ne sais pas quoi faire\nRencontré par : Myriam NETTAH à 11h55\nContact du badge : 0699034686 / benoit.delmer@clarke-energy.com\nCode postal : 13300"
},
{
"cid": "LEADS2111-26",
"prenom": "Maxime",
"nom": "GRIPPI",
"mail": "m.grippi13@icloud.com",
"tel": "0646246229",
"campus": "HECG",
"heure": "11:53",
"note": "Profil : Lycéen\nDomaines : Armée défense sécurité, Luxe, Santé paramédical\nFormations envisagées : Université - licence, Université - master\nRencontré par : Myriam NETTAH à 11h53\nCode postal : 13580"
},
{
"cid": "LEADS2111-27",
"prenom": "Léane",
"nom": "MARON",
"mail": "leane.ll101023@gmail.com",
"tel": "",
"campus": "HECG",
"heure": "11:31",
"note": "Profil : Parent — niveau : Première\nFormations envisagées : BUT, Classes préparatoires, Ecoles d'ingénieurs\nRencontré par : Rayan BEN JEDDOU à 11h31\nContact du badge : 0609249660 / anneisabellemaron@gmail.com\nCode postal : 13340\nPrénom déduit du mail, nom du parent (à confirmer). Parent : Anne-Isabelle MARON"
},
{
"cid": "LEADS2111-28",
"prenom": "Océane",
"nom": "FONTANINI",
"mail": "oceanefontanini@hotmail.com",
"tel": "0673582578",
"campus": "HECG",
"heure": "11:22",
"note": "Profil : Lycéen — niveau : Terminale\nDomaines : Droit science politique administration publique\nFormations envisagées : Alternance, BTS, BUT, Université - licence, Université - master\nRencontré par : Rayan BEN JEDDOU à 11h22\nContact du badge : 0623839587\nCode postal : 83910"
},
{
"cid": "LEADS2111-29",
"prenom": "Tessa",
"nom": "BERNON",
"mail": "bernon.tessa@gmail.com",
"tel": "0651102344",
"campus": "HECG",
"heure": "11:20",
"note": "Profil : Lycéen — niveau : Terminale technologique\nDomaines : Action sociale animation sciences sociales\nFormations envisagées : BUT\nRencontré par : Rayan BEN JEDDOU à 11h20\nCode postal : 13300"
},
{
"cid": "LEADS2111-30",
"prenom": "Martin",
"nom": "CASCALES",
"mail": "cascalesmartin93@gmail.com",
"tel": "0787166672",
"campus": "HECG",
"heure": "11:19",
"note": "Profil : Lycéen — niveau : Terminale\nDomaines : Communication information journalisme, Comptabilité gestion ressources humaines\nFormations envisagées : Alternance, BUT\nRencontré par : Rayan BEN JEDDOU à 11h19\nCode postal : 13116"
},
{
"cid": "LEADS2111-31",
"prenom": "Jun",
"nom": "VALENSI",
"mail": "valensijun@gmail.com",
"tel": "0651306108",
"campus": "HECG",
"heure": "11:18",
"note": "Profil : Lycéen\nDomaines : Commerce vente marketing\nRencontré par : Rayan BEN JEDDOU à 11h18\nCode postal : 13330"
},
{
"cid": "LEADS2111-32",
"prenom": "Maée",
"nom": "MELLAN",
"mail": "mellanmaee3@gmail.com",
"tel": "0762472391",
"campus": "HECG",
"heure": "11:18",
"note": "Profil : Lycéen — niveau : Seconde Générale et Technologique\nDomaines : Communication information journalisme\nFormations envisagées : Je ne sais pas quoi faire\nRencontré par : Rayan BEN JEDDOU à 11h18\nCode postal : 13430"
},
{
"cid": "LEADS2111-33",
"prenom": "Mathieu",
"nom": "MEARELLI",
"mail": "mathmath13500@gmail.com",
"tel": "0772204098",
"campus": "HECG",
"heure": "11:16",
"note": "Profil : Parent — niveau : Seconde Générale et Technologique\nDomaines : Commerce vente marketing\nFormations envisagées : Je ne sais pas quoi faire\nRencontré par : Myriam NETTAH à 11h16\nContact du badge : 0628714956 / mearelli.isabelle@orange.fr\nCode postal : 13500"
},
{
"cid": "LEADS2111-34",
"prenom": "Héloïse",
"nom": "AUDA",
"mail": "heloiseauda04@gmail.com",
"tel": "0767338360",
"campus": "HECG",
"heure": "10:59",
"note": "Profil : Lycéen — niveau : Première\nDomaines : Je ne sais pas quoi faire\nFormations envisagées : Classes préparatoires, Ecoles d'ingénieurs, Ecoles de commerce - IEP, Université - licence\nRencontré par : Rayan BEN JEDDOU à 10h59\nCode postal : 83470"
},
{
"cid": "LEADS2111-35",
"prenom": "Yliann",
"nom": "MOGIS",
"mail": "yliann04000@gmail.com",
"tel": "0633473961",
"campus": "HECG",
"heure": "10:38",
"note": "Profil : Lycéen\nDomaines : Action sociale animation sciences sociales, Architecture géographie aménagement, Digital Informatique multimédia, Electronique automatique, Génie civil bâtiment, Secretariat assistant de gestion\nFormations envisagées : Alternance, BTS\nRencontré par : Rayan BEN JEDDOU à 10h38\nCode postal : 04000"
},
{
"cid": "LEADS2111-36",
"prenom": "Maëlys",
"nom": "PEJCL",
"mail": "lseghi13120@aol.com",
"tel": "0615650197",
"campus": "HECG",
"heure": "10:38",
"note": "Profil : Lycéen — niveau : Première technologique\nDomaines : Je ne sais pas quoi faire\nFormations envisagées : Je ne sais pas quoi faire\nRencontré par : Myriam NETTAH à 10h38\nCode postal : 13680"
},
{
"cid": "LEADS2111-37",
"prenom": "Antoine",
"nom": "MAIXANT",
"mail": "maixantantoine@icloud.com",
"tel": "0623951352",
"campus": "HECG",
"heure": "10:34",
"note": "Profil : Parent — niveau : Terminale technologique\nDomaines : Banque assurance finance immobilier, Commerce vente marketing, Comptabilité gestion ressources humaines\nFormations envisagées : Alternance, BTS, BUT\nRencontré par : Myriam NETTAH à 10h34\nContact du badge : 0620557287 / maixantkristel@sfr.fr\nCode postal : 13770\nParent : Kristel VERDU"
},
{
"cid": "LEADS2111-38",
"prenom": "Leandre",
"nom": "VIGOUROUX",
"mail": "leandrevig@icloud.com",
"tel": "0783016666",
"campus": "HECG",
"heure": "10:33",
"note": "Profil : Lycéen — niveau : Terminale technologique\nDomaines : Commerce vente marketing, Economie\nFormations envisagées : Ecoles de commerce - IEP\nRencontré par : Myriam NETTAH à 10h33\nCode postal : 13410"
},
{
"cid": "LEADS2111-39",
"prenom": "Léna",
"nom": "VERHAEGHE",
"mail": "verhaeghelena7@gmail.com",
"tel": "0782427436",
"campus": "HECG",
"heure": "10:28",
"note": "Profil : Lycéen\nDomaines : Banque assurance finance immobilier, Comptabilité gestion ressources humaines\nFormations envisagées : Je ne sais pas quoi faire\nRencontré par : Myriam NETTAH à 10h28\nCode postal : 13119"
},
{
"cid": "LEADS2111-40",
"prenom": "Valentin",
"nom": "ASSO",
"mail": "assovalentin@icloud.com",
"tel": "0652192847",
"campus": "HECG",
"heure": "10:24",
"note": "Profil : Lycéen\nDomaines : Commerce vente marketing\nFormations envisagées : Je ne sais pas quoi faire\nRencontré par : Rayan BEN JEDDOU à 10h24\nContact du badge : bazuso56@gmail.com\nCode postal : 13590"
},
{
"cid": "LEADS2111-41",
"prenom": "Laetitia",
"nom": "MONTARELLO",
"mail": "laetitiamontarello@hotmail.fr",
"tel": "0616251125",
"campus": "HECG",
"heure": "10:22",
"note": "Profil : Parent — niveau : Terminale Pro\nDomaines : Banque assurance finance immobilier\nFormations envisagées : Alternance, Bachelor, BTS\nRencontré par : Myriam NETTAH à 10h22\nCode postal : 13220\nParent — coordonnées de l'enfant non renseignées"
}
];

function _importLeadsSalon2111Auto_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty("IMPORT_LEADS_SALON_2111")) return;
  props.setProperty("IMPORT_LEADS_SALON_2111", "en cours " + new Date().toISOString());
  var ss = SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ");
  var debut = new Date(Date.now() - 60000);
  var heureParCid = {};
  var prospects = LEADS_SALON_2111.map(function(x) {
    heureParCid[x.cid] = x.heure;
    return { cid: x.cid, nom: x.nom, prenom: x.prenom, mail: x.mail, tel: x.tel, classe: "", campus: x.campus, note: x.note,
             salon: "Salon de l'Étudiant de Marseille", salonDate: "2026-11-21", salonLieu: "Marseille" };
  });
  var out = JSON.parse(handleAddProspectsSalon({ prospects: prospects, nomAuteur: "Mehdi MOKHTARI", roleAuteur: "Administrateur" }, ss).getContent());
  if (!out.success) { props.setProperty("IMPORT_LEADS_SALON_2111", "ERREUR " + out.message); return; }
  var dateDe = function(cid) { var h = String(heureParCid[cid] || "12:00").split(":"); return new Date(2026, 10, 21, Number(h[0]), Number(h[1])); };
  var parId = {}, bilan = { ajoute: 0, existant: 0, deja: 0, erreur: 0, existants: [] };
  out.resultats.forEach(function(r) {
    bilan[r.statut] = (bilan[r.statut] || 0) + 1;
    if (r.statut === "existant") bilan.existants.push(r.id);
    if (r.id && !parId[r.id]) parId[r.id] = { date: dateDe(r.cid), nouveau: r.statut === "ajoute" };
  });
  // Date de création des nouvelles fiches = heure du scan sur le salon
  var ws = getSheetSafe(ss, "ETUDIANTS"), colDC = _getDateCreationCol_(ws);
  var ids = ws.getRange(1, 1, ws.getLastRow(), 1).getValues();
  for (var i = 1; i < ids.length; i++) {
    var p = parId[String(ids[i][0])];
    if (p && p.nouveau) ws.getRange(i + 1, colDC).setValue(p.date);
    if (p) { try { _invalidateStudentDetailsCache_(String(ids[i][0])); } catch (eC) {} }
  }
  // Notes du carnet de route datées du 21/11 (seulement celles écrites par cet import)
  var wc = getSheetSafe(ss, "Carnet_route"), n = wc.getLastRow(), from = Math.max(2, n - 120);
  var cr = wc.getRange(from, 1, n - from + 1, 4).getValues();
  for (var j = 0; j < cr.length; j++) {
    var q = parId[String(cr[j][0])];
    if (q && cr[j][2] instanceof Date && cr[j][2] >= debut && String(cr[j][3]).indexOf("📍") === 0) wc.getRange(from + j, 3).setValue(q.date);
  }
  try { CacheService.getScriptCache().removeAll(["students_min_v1", "salon_ref_v2"]); } catch (eK) {}
  try { logAction("Mehdi MOKHTARI", "Administrateur", "Création", "Salon — Import leads", "EVT 21/11/2026",
    "Import du fichier 222HECG_leads_export-2.xlsx (Salon de l'Étudiant de Marseille, 21/11/2026) : " + bilan.ajoute + " prospect(s) créé(s), " + bilan.existant + " déjà connu(s) (événement ajouté à leur fiche)"); } catch (eL) {}
  props.setProperty("IMPORT_LEADS_SALON_2111", "OK " + new Date().toISOString() + " " + JSON.stringify(bilan));
}
