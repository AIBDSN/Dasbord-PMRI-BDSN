/* ============================================================
   V3 - Avis PM loading + parsing (loaded before dashboard.js)
   ============================================================ */

const MAINT0910_RULES = [
  { keys:['fuite externe','fuite sur ci','fuite clapet','fuite robinet','fuite tuyaut','fuite detend','fuite au niveau'],
    nature:'P', delai:'IMMÉDIAT', jours:0, traitement:'Référentiel classification fuites — ISG immédiat' },
  { keys:['ocg non accessible','regard visible et robinet non accessible'],
    nature:'P', delai:'P-1 mois', jours:30, traitement:'Réalisation lors de la GM + info hiérarchique' },
  { keys:['pénétration ci non étanche'],
    nature:'P', delai:'P-1 mois', jours:30, traitement:'Traitement lors de la gamme' },
  { keys:['dysfonctionnement détendeur'],
    nature:'P', delai:'P-1 mois', jours:30, traitement:'Processus régulateur / ISG Dépannage' },
  { keys:['indications erronées sur plaque repère oci'],
    nature:'P', delai:'P-1 mois', jours:30, traitement:'Mesures conservatoires immédiates' },
  { keys:['by pass frauduleux'],
    nature:'P', delai:'P-1 mois', jours:30, traitement:'Info hiérarchique + fermeture robinet + Contentieux' },
  { keys:['branchement improductif non sécurisé'],
    nature:'P', delai:'P-1 mois', jours:30, traitement:'Procédure de sécurisation' },
  { keys:['défaillance comptage','bruit compteur'],
    nature:'P', delai:'P-1 mois', jours:30, traitement:'ISG Dépannage' },
  { keys:['lot de rappel','rappel de lot'],
    nature:'P', delai:'P-3 mois', jours:90, traitement:'Remplacement détendeur + traçabilité GMAO' },
  { keys:['accès impossible à la ci','accès impossible ci'],
    nature:'P', delai:'P-18 mois', jours:548, traitement:'Procédure Accessibilité + info syndic' },
  { keys:['accès immeuble impossible'],
    nature:'P', delai:'P-18 mois', jours:548, traitement:'Procédure Accessibilité + info syndic' },
  { keys:['accès impossible aux ouvrages','accès impossible conduite'],
    nature:'P', delai:'P-18 mois', jours:548, traitement:'Procédure Accessibilité + info syndic' },
  { keys:['absence ou non visibilité robinet','non visibilité robinet'],
    nature:'P', delai:'P-1 an', jours:365, traitement:'Info hiérarchique + info gestionnaire + démarche autorisation' },
  { keys:['non déclenchement ddmp','absence de ddmp'],
    nature:'P', delai:'P-1 an', jours:365, traitement:'Info hiérarchique + programmation travaux' },
  { keys:['réparation provisoire en place'],
    nature:'P', delai:'P-1 an', jours:365, traitement:'Vérification étanchéité + MAJ O²' },
  { keys:['défaut de fixation sur ouvrage générant un danger'],
    nature:'P', delai:'P-1 an', jours:365, traitement:'Info CE immédiate + mesures conservatoires' },
  { keys:['plaque d\'identif','plaque de repérage absente','absence plaque repère'],
    nature:'R', delai:'R-1 mois', jours:30, traitement:'Réalisation lors de la GM' },
  { keys:['indications erronées sur plaque repère'],
    nature:'R', delai:'R-1 mois', jours:30, traitement:'Traitement lors de la GM + info hiérarchique' },
  { keys:['sens de fermeture non conventionnelle'],
    nature:'R', delai:'R-1 mois', jours:30, traitement:'Pose plaque informative T122' },
  // ── Règles complémentaires terrain (mars 2026) ──────────────────

  // Accès BRC
  { keys:['acces impossible conduite montante','acces impossible cm',
          "pb d'acces aux equipements",'pb acces aux equipements',
          'acces impossible equipements du poste','acces robinet rdd impossible'],
    nature:'P', delai:'P-18 mois', jours:548, traitement:'Procédure Accessibilité + info syndic' },

  // Manœuvrabilité
  { keys:['difficilement manoeuvrable'],
    nature:'R', delai:'R-1 an', jours:365, traitement:'Info hiérarchique + programmation' },

  // Branchement improductif (= non sécurisé SVR10 p.59)
  { keys:['branchement improductif non obture','branchement improductif non obtue'],
    nature:'P', delai:'P-1 mois', jours:30, traitement:'Procédure sécurisation + macaron + bouchon (SVR10 p.59)' },

  // Plaques / repérage
  { keys:['abs.plaque consigne','abs plaque consigne','absence plaque consigne'],
    nature:'P', delai:'P-1 an', jours:365, traitement:'Remise au propriétaire (CM > 400 mbar ou > 10 logements)' },
  { keys:['plaque de reperage erronee','plaque reperage erronee'],
    nature:'R', delai:'R-1 mois', jours:30, traitement:'Traitement lors de la GM' },
  { keys:['reperage manquant pour certains br.part','reperage manquant branchements particuliers'],
    nature:'R', delai:'R-3', jours:null, traitement:'Info hiérarchique + traçabilité' },
  { keys:['absence de reperage (si plusieurs cm)','absence reperage si plusieurs cm'],
    nature:'R', delai:'R-1 an', jours:365, traitement:'Traitement lors de la GM' },

  // Clé de manœuvre
  { keys:['absence de cle de manoeuvre f','absence cle de manoeuvre'],
    nature:'P', delai:'P-1 an', jours:365, traitement:'Remise au propriétaire' },

  // Tampon / trappe
  { keys:['absence de tampon ou trappe','absence tampon ou trappe','tampon absent','trappe absente'],
    nature:'R', delai:'R-1 mois', jours:30, traitement:'Info hiérarchique immédiate — balisage + sécurisation journée + remplacement programmé' },
  { keys:['tampon ou trappe deteriore','tampon deteriore','trappe deterioree'],
    nature:'R', delai:'P-1 mois', jours:30, traitement:'Info hiérarchique immédiate — balisage + sécurisation journée + remplacement programmé' },

  // Objets / gaine
  { keys:['objets deposes dans la gaine','objets dans la gaine'],
    nature:'INFO', delai:'N/A', jours:null, traitement:'Signalement — hors périmètre MAINT0910' },
  { keys:['objets accroches a la tuyauterie'],
    nature:'R', delai:'R-3', jours:null, traitement:'Info hiérarchique + traçabilité' },
  { keys:['probleme fermeture/ouverture de la gaine','probleme fermeture ouverture de la gaine',
          'probleme fermeture','verre dormant'],
    nature:'R', delai:'R-3', jours:null, traitement:'Info hiérarchique + programmation' },

  // Protection mécanique
  { keys:['abs prot.meca','abs prot meca','abs. prot.meca',
          'protec meca insuf','protection mecanique insuffisante','protection mecanique absente'],
    nature:'R', delai:'R-3', jours:null, traitement:'Info syndic pour traitement' },

  // Chambre de vanne / armoire
  { keys:['chambre de vanne deteriore','chambre de vanne fissure'],
    nature:'R', delai:'R-3', jours:null, traitement:'Info hiérarchique + programmation' },
  { keys:["deterioration de l'armoire",'deterioration armoire'],
    nature:'R', delai:'R-3', jours:null, traitement:'Info hiérarchique + programmation' },
  { keys:['probleme fermeture/ouverture de coffret','probleme fermeture ouverture de coffret'],
    nature:'R', delai:'R-1 an', jours:365, traitement:'Info hiérarchique + programmation' },

  // Mesure compensatoire
  { keys:['mesure compensatoire endommagee','mesure compensatoire deterioree'],
    nature:'R', delai:'R-3', jours:null, traitement:'Info hiérarchique + traçabilité' },

  // Détendeur / régulateur spécifiques
  { keys:['cdg detendeur non-normalise','cdg detendeur non normalise'],
    nature:'P', delai:'P-3 mois', jours:90, traitement:'Remplacement détendeur CDG non-normalisé + traçabilité GMAO' },
  { keys:['pression regulee incorrecte','pression regulee non conforme'],
    nature:'P', delai:'IMMÉDIAT', jours:0, traitement:'ISG immédiat — processus régulateur' },
  { keys:['non etancheite interne regulateur'],
    nature:'P', delai:'P-1 mois', jours:30, traitement:'Info hiérarchique + processus régulateur + ISG Dépannage' },
  { keys:['fermeture intempestive'],
    nature:'R', delai:'P-3 mois', jours:90, traitement:'Info hiérarchique + programmation travaux' },

  // Butée
  { keys:['butee cassee ou introuvable','butee cassee','butee introuvable'],
    nature:'P', delai:'P-1 an', jours:365, traitement:'Info CE immédiate + traçabilité Carpathe + procédure déblocage robinets (Annexe)' },

  // Non étanchéité
  { keys:['non etancheite interne (incidt ou autre)','non etancheite interne incident'],
    nature:'R', delai:'R-1 an', jours:365, traitement:'Info CE + traçabilité Carpathe + analyse BEX/GROP/BERG + intervention ou déclassement' },
  { keys:['non etancheite a la fermeture'],
    nature:'R', delai:'R-1 an', jours:365, traitement:'Info CE + traçabilité Carpathe + analyse BEX/GROP/BERG + intervention ou déclassement' },

  // Défaut transmission état
  { keys:["defaut transmission d'etat",'defaut transmission etat'],
    nature:'R', delai:'R-1 an', jours:365, traitement:'Info hiérarchique + programmation' },

  // Ventilation / zone transition
  { keys:['absence de ventilation galerie technique','absence ventilation galerie'],
    nature:'INFO', delai:'N/A', jours:null, traitement:'Signalement — hors périmètre correctif standard' },
  { keys:['zone de transition en mauvais etat','zone transition mauvais etat'],
    nature:'INFO', delai:'N/A', jours:null, traitement:'Signalement — hors périmètre correctif standard' },

  // Corrosion / vétusté générique
  { keys:['enrouillement bride/manchon','enrouillement bride','enrouillement manchon'],
    nature:'R', delai:'R-1 an', jours:365, traitement:'Débriefing + photos + analyse technique + programmation éventuelle' },
  { keys:['vetuste (hors dgi)','vetuste hors dgi',
          'fixations vetustes','vetuste tuyauterie','vetuste assemblages','vetuste accessoire'],
    nature:'R', delai:'R-1 an', jours:365, traitement:'Débriefing + photos + analyse technique + programmation éventuelle' },
  { keys:['tuy. vetuste/corrodee/chocs','tuyauterie vetuste corrodee chocs'],
    nature:'R', delai:'R-1 an', jours:365, traitement:'Débriefing + photos + analyse technique + programmation éventuelle' },
  { keys:['corrosion superficielle canalisation'],
    nature:'R', delai:'R-1 an', jours:365, traitement:'Débriefing + photos + analyse technique + programmation éventuelle' },

  // TDR et hors périmètre
  { keys:['tdr - adresse','tdr-annee','tdr-','saisie non conforme','en local ferme','photo absente'],
    nature:'INFO', delai:'N/A', jours:null, traitement:'Mise à jour données SAP — hors périmètre MAINT0910' },
];

function normalizeStr(s) {
  return String(s || '').toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/\u0153/g, 'oe').replace(/\u00e6/g, 'ae')
    .replace(/['']/g, "'")
    .replace(/\s+/g, ' ').trim();
}

function isRobReseau(ouvrage) {
  const up = String(ouvrage || '').toUpperCase();
  return up.includes('ROB') && !up.includes('BRC');
}

const MAINT0410_RULES = [
  { keys:['fuite externe','fuite robinet','fuite sur robinet'],
    nature:'P', delai:'IMMÉDIAT', jours:0, traitement:'ISG immédiat — référentiel classification fuites' },
  { keys:['non manoeuvrabilite','non manoeuvrable','manoeuvrabilite impossible'],
    nature:'P', delai:'P-1 mois', jours:30, traitement:'Processus déblocage robinet (cf. Annexe MAINT0410) + info hiérarchique' },
  { keys:['acces impossible','non accessible','regard non accessible','acces robinet'],
    nature:'P', delai:'P-1 mois', jours:30, traitement:'Démarche autorisation + info hiérarchique + procédure accessibilité' },
  { keys:['absence ou non visibilite robinet','non visibilite robinet'],
    nature:'P', delai:'P-1 an', jours:365, traitement:'Info hiérarchique + info gestionnaire + démarche autorisation' },
  { keys:['plaque absente','plaque erronee','plaque de reperage'],
    nature:'R', delai:'R-1 mois', jours:30, traitement:'Traitement lors de la GM' },
  { keys:['sens de fermeture non conventionnelle'],
    nature:'R', delai:'R-1 mois', jours:30, traitement:'Pose plaque informative' },
  { keys:['vetuste','corrosion','deterioration','enrouillement'],
    nature:'R', delai:'R-1 an', jours:365, traitement:'Débriefing + photos + analyse technique + programmation éventuelle' },
  { keys:['joint isolant'],
    nature:'R', delai:'R-1 an', jours:365, traitement:'Remplacement joint isolant lors de la maintenance' },
  { keys:['butee cassee','butee introuvable'],
    nature:'P', delai:'P-1 an', jours:365, traitement:'Info CE immédiate + traçabilité Carpathe + procédure déblocage robinets' },
  { keys:['etat rob','non conforme'],
    nature:'P', delai:'IMMÉDIAT', jours:0, traitement:'ISG immédiat — état robinet non conforme' },
];

function findMaintRule(defail, ouvrage) {
  if (!defail) return { rule: null, ref: null };
  const df   = normalizeStr(defail);
  const rob  = isRobReseau(ouvrage);
  const rules = rob ? MAINT0410_RULES : MAINT0910_RULES;
  const ref   = rob ? 'MAINT0410'     : 'MAINT0910';
  for (const rule of rules) {
    if (rule.keys.some(k => df.includes(normalizeStr(k)))) return { rule, ref };
  }
  return { rule: null, ref: null };
}

function loadAvisFile(input) {
  const file = input.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const wb  = XLSX.read(e.target.result, { type:'array', cellDates:true });
      const ws  = wb.Sheets[wb.SheetNames[0]];
      const raw = XLSX.utils.sheet_to_json(ws, { defval:'' });
      AVIS_DATA = raw.map(parseAvisRow);

      document.getElementById('drop-avis-label').textContent = '✓ ' + file.name.replace(/\.xlsx?$/i,'');
      document.getElementById('file-avis-drop-zone').classList.add('loaded');
      document.getElementById('tab-avis-count').textContent = AVIS_DATA.length.toLocaleString('de-DE');

      linkAvisToOT();
      if (typeof applyFilters === 'function') applyFilters();
      else if (typeof FILTERED_AVIS !== 'undefined') FILTERED_AVIS = [...AVIS_DATA];

      document.getElementById('empty-state').style.display = 'none';
      document.getElementById('dashboard').style.display   = 'block';

      setMode('avis');
    } catch(err) { alert('Erreur lecture avis PM : ' + err.message); console.error(err); }
  };
  reader.readAsArrayBuffer(file);
}

function parseAvisRow(row) {
  const defail  = String(row['Code défaillance (Description)'] || row['Défaillance'] || '');
  const ouvrage = String(row['Poste technique'] || row["Désignation de l'objet technique"] || row['Ouvrage détecté'] || row['Ouvrage'] || '');
  const prio    = String(row['Priorité'] || '');
  const debut   = parseDate(row['Début souhaité (Europe/Paris)'] || row['Date début']);
  const now     = new Date();
  const ageDays = debut ? Math.floor((now - debut) / 86400000) : 0;
  const otRaw   = String(row['Ordre de travail'] || row['OT correctif'] || '').trim();
  const otNum   = /^[-–\s]*$/.test(otRaw) ? '' : otRaw;
  const hasOT   = otNum.length > 0;

  const { rule, ref } = findMaintRule(defail, ouvrage);
  let conformite, confCode, depasse = 0;
  if (rule) {
    if (rule.nature === 'INFO') {
      conformite = 'ℹ️ Hors périmètre'; confCode = 'INFO';
    } else if (rule.jours === null) {
      conformite = '⚠️ Sans délai fixé'; confCode = 'NF';
    } else {
      depasse = ageDays - rule.jours;
      if (depasse > 0) { conformite = '⛔ +' + depasse + 'j'; confCode = 'KO'; }
      else             { conformite = '✅ Dans les délais';    confCode = 'OK'; }
    }
  } else {
    conformite = '❓ Non qualifié'; confCode = '?';
  }

  const isF0    = prio.includes('F0');
  const prioGrp = isF0                  ? 'F0'
                : prio.startsWith('P1') ? 'P1'
                : prio.startsWith('P2') ? 'P2'
                : prio.startsWith('R1') ? 'R1'
                : prio.startsWith('R2') ? 'R2'
                : prio.startsWith('R3') ? 'R3' : 'Autre';

  return {
    avis:       String(row['Avis'] || '').trim(),
    prio, prioGrp, isF0,
    defail, ouvrage,
    cause:      String(row['Code de la cause (Description)'] || row['Cause'] || ''),
    ville:      String(row['Ville normalisée'] || row['Ville'] || ''),
    rue:        (String(row['Nº de rue'] || '') + ' ' + String(row['Rue'] || '') + ' ' + String(row['Adresse'] || '')).trim(),
    objet:      ouvrage,
    decl:       String(row["Déclaré par nom d'utilisateur"] || ''),
    dateRef:    debut,
    dateStr:    debut ? debut.toLocaleDateString('fr-FR') : '—',
    ageDays,    hasOT, otNum,
    conformite, confCode, depasse,
    delaiLabel: rule ? rule.delai      : '—',
    traitement: rule ? rule.traitement : 'Vérifier référentiel MAINT',
    nature:     rule ? rule.nature     : '?',
    refLabel:   ref  || '—',
    entiteAvis: shortEntite(String(row['Entité en charge'] || row['Entité'] || '')),
    ouvrageAvis: isRobReseau(ouvrage) ? 'Robinet Réseau' : 'CICM / BRC',
    otLie:      null,
  };
}

function linkAvisToOT() {
  if (!ALL_DATA.length) return;
  AVIS_DATA.forEach(a => {
    if (!a.hasOT) return;
    const found = ALL_DATA.find(d => d.numSAP === a.otNum || d.ordre === a.otNum);
    a.otLie = found ? found.numSAP : null;
  });
}

window.GRDFDashV3 = window.GRDFDashV3 || {};
window.GRDFDashV3.avis = { normalizeStr, findMaintRule, parseAvisRow, linkAvisToOT, loadAvisFile };
window.loadAvisFile = loadAvisFile;
