/* =============================================================================
   OFFRES — Détection « alternance ou pas » (novembre 2026)
   La colonne P (Type de contrat) était remplie d'office avec « Alternance » par les
   aspirateurs (LinkedIn, HelloWork, alertes Google) : elle n'est pas fiable.
   _classerContratOffre_ analyse l'intitulé (O), la plateforme (D), le type (P) et les
   notes (G) et renvoie { code, label, raison } :
     alternance — marqueur explicite (alternance, apprenti, contrat pro, BTS/DCG…) ou source sûre
     stage      — stage / stagiaire : ce n'est PAS une alternance
     non        — CDI, CDD, intérim, poste confirmé/senior… : ce n'est PAS une alternance
     inconnu    — rien ne permet de l'affirmer : alternance non confirmée, à vérifier
   Utilisé par getOffres (cartes), getStudentDetails (dossier) et les aspirateurs
   (les offres « non » ne sont plus importées).
   ========================================================================== */

function _offreNormTxt_(s) {
  return String(s || "")
    .replace(/&#x([0-9a-f]+);/gi, function (m, h) { return String.fromCharCode(parseInt(h, 16)); })
    .replace(/&#(\d+);/g, function (m, d) { return String.fromCharCode(+d); })
    .replace(/&amp;/g, "&").replace(/&quot;/g, '"')
    .normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
}

var OFFRE_RX_ALT   = /alternan|apprenti|apprentissage|contrat pro\b|contrat de pro|professionnalisation|work[ -]?study|\bbts\b|\bdcg\b|\bdscg\b|but gea|\bbac ?\+ ?[2-5]\b.*(alternance|ecole)|rythme ecole|en cours de formation/;
var OFFRE_RX_STAGE = /\bstages?\b|\bstagiaire|\binternship\b|\bintern\b|memorialiste/;
var OFFRE_RX_NON   = [
  [/\bcdi\b/, "CDI"], [/\bcdd\b/, "CDD"], [/interim/, "intérim"], [/\bfreelance|independant|auto-?entrepreneur/, "indépendant"],
  [/job d.?ete|saisonnier|saisonniere|extra\b/, "job saisonnier"], [/vacataire|benevol/, "vacation / bénévolat"], [/\bmi ?-? ?temps|temps partiel/, "temps partiel"],
  [/confirme|experimente|\bsenior\b/, "poste confirmé"], [/chef de missions?\b|responsable comptable|responsable paie|directeur|directrice|head of|manager confirme/, "poste d'encadrement"],
  [/expert[- ]comptable(?! stagiaire)/, "expert-comptable"], [/\d{2} ?-? ?\d{0,2} ?k ?(€|eur)|\d{2} ?k€/, "salaire de poste confirmé"]
];
var OFFRE_SOURCES_SURES = /la bonne alternance|engagement jeunes|alternance\.emploi|walt|hellowork alternance/;

function _classerContratOffre_(id, plateforme, poste, typeP, notes) {
  var titre = _offreNormTxt_(poste), type = _offreNormTxt_(typeP).trim(), pf = _offreNormTxt_(plateforme);
  var idU = String(id || "").toUpperCase();
  // La valeur « alternance » seule en colonne P est celle posée d'office → ignorée
  var typeFiable = type && type !== "alternance" ? type : "";
  // France Travail : les apprentissages y sont libellés « CDD - 24 mois » → le type ne prouve pas un CDD classique
  var typeNon = idU.indexOf("FT-") === 0 ? "" : typeFiable;
  var txt = titre + " | " + typeFiable + " | " + _offreNormTxt_(notes).slice(0, 400);

  if (OFFRE_RX_ALT.test(titre) || OFFRE_RX_ALT.test(typeFiable))
    return { code: "alternance", label: "Alternance", raison: "" };
  if (OFFRE_RX_STAGE.test(titre) || /^stage/.test(type))
    return { code: "stage", label: "Stage — pas une alternance", raison: "stage" };
  for (var i = 0; i < OFFRE_RX_NON.length; i++) {
    if (OFFRE_RX_NON[i][0].test(titre) || (typeNon && OFFRE_RX_NON[i][0].test(typeNon)))
      return { code: "non", label: "Pas une alternance", raison: OFFRE_RX_NON[i][1] };
  }
  if (OFFRE_RX_ALT.test(txt)) return { code: "alternance", label: "Alternance", raison: "" };
  // Sources sûres : plateformes d'alternance, France Travail (filtré à l'aspiration), saisies manuelles
  if (OFFRE_SOURCES_SURES.test(pf) || idU.indexOf("LBA-") === 0 || idU.indexOf("FT-") === 0 ||
      idU.indexOf("OFR") === 0 || /direct partenaire/.test(pf))
    return { code: "alternance", label: "Alternance", raison: "" };
  return { code: "inconnu", label: "Alternance non confirmée", raison: "rien dans l'intitulé n'indique une alternance" };
}

// Valeur à écrire en colonne P par les aspirateurs
function _typeContratAspire_(cls) {
  return cls.code === "alternance" ? "Alternance" : cls.code === "stage" ? "Stage" : cls.code === "non" ? "Pas une alternance (" + cls.raison + ")" : "À vérifier";
}

/** Contrôle (éditeur Apps Script) : répartition des offres actuelles par catégorie. */
function controlerContratsOffres() {
  var d = getSheetSafe(SpreadsheetApp.openById("1TfghkrbVnei_vQTdO3_jXW6-gqVQIKKYGtQN4_W_iWQ"), "OFFRES").getDataRange().getValues();
  var tot = {}, ex = {};
  for (var i = 1; i < d.length; i++) {
    if (!d[i][0]) continue;
    var c = _classerContratOffre_(d[i][0], d[i][3], d[i][14], d[i][15], d[i][6]);
    tot[c.code] = (tot[c.code] || 0) + 1;
    (ex[c.code] = ex[c.code] || []).length < 15 && ex[c.code].push(String(d[i][14]) + (c.raison ? " [" + c.raison + "]" : ""));
  }
  Logger.log(JSON.stringify(tot)); Logger.log(JSON.stringify(ex, null, 1));
  return tot;
}
