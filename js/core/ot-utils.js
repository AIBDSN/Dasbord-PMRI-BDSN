/* ============================================================
   V3 - OT parsing utilities (loaded before dashboard.js)
   ============================================================ */

function parseDate(val) {
  if (!val || val === '') return null;
  if (val instanceof Date) return val;
  const d = new Date(val);
  return isNaN(d) ? null : d;
}

function isEnRetard(statut, dateRef, statutOrdre) {
  if (!['ACTI','EPR','PRG','OGDI'].includes(statut)) return false;
  if (statutOrdre && !statutOrdre.includes('Lancé')) return false;
  const d = parseDate(dateRef);
  if (!d) return false;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return d < today;
}

function statutLabel(s) {
  return {TERM:'Terminé','TERM ANO':'Terminé avec anomalie',EPR:'En préparation',
          ACTI:'Actif',PRG:'Programmé',OGDI:'OGDI','TERM CTOK':'TERM CTOK'}[s] || s;
}

function shortEntite(e) {
  const m = e.match(/\((AI101\w+)\)/);
  if (m) return m[1];
  return e.replace(/\s*\(.*?\)\s*/g,'').trim() || e;
}

function classifyOuvrage(gamme, ordre, travail) {
  if (gamme.includes('Remplacement régulateur') || /^Rempla/i.test(ordre) ||
      /remplacement.+r.gulateur/i.test(ordre))   return 'TDR - Remplacement';
  if (gamme.includes('visite préalable') ||
      /^Visite/i.test(ordre))                    return 'TDR - Visite préalable';
  if (gamme.includes('TDR'))                     return 'TDR - Remplacement';
  if (/DDMP/i.test(gamme) || /^DDMP/i.test(ordre)) return 'DDMP';
  if (gamme.includes('CICM') || gamme.includes('OCG') ||
      /^(F1|F3)-/.test(gamme))                   return 'CICM / BRC';
  if (gamme.includes('Robinet') || gamme.includes('décompression'))
                                                 return 'Robinet Réseau';
  if (gamme.includes('PDR'))                     return 'PDR';
  if (gamme.includes('PDL'))                     return 'PDL';
  if (gamme.includes('SCI'))                     return 'SCI';
  if (/Inventaire|Enquête/i.test(gamme))         return 'Enquête / Inventaire';
  if (gamme.includes('DPBE'))                    return 'DPBE';
  if (/^SINIS|^SINRV/.test(ordre))              return 'SIN (Point singulier)';
  if (/^VSIC/.test(ordre))                      return 'VSIC';
  if (/^BRCIS|^BRCRV|^BRCRF/.test(ordre))      return 'CICM / BRC';
  if (/^ROBIS|^ROB2AG|^RDDIS/.test(ordre))     return 'Robinet Réseau';
  if (/^PIS[\s-]/.test(ordre))                 return 'Robinet Réseau';
  if (/^PDRIS|^PDRRV/.test(ordre))             return 'PDR';
  if (/^PDLIS|^PDLRV/.test(ordre))             return 'PDL';
  if (travail === 'Exploitation (EXP)')         return 'Exploitation';
  return 'Autre';
}

function parseRow(row) {
  const gamme      = String(row['Désignation de la gamme'] || '');
  const ordre      = String(row['Ordre'] || '');
  const travail    = String(row["Type de travail"] || '');
  const statut     = String(row['Statut utilisateur'] || '');
  const statutOrdre= String(row["Description du statut de l'ordre"] || '');
  const typeOrdre  = String(row["Type d'ordre"] || '');
  const termine    = ['TERM','TERM ANO','TERM CTOK'].includes(statut);
  const attenteReferent = termine && statutOrdre === 'Lancé';
  const incoherence = !termine && statutOrdre === 'Clôturé commercialement';

  const numMatch = ordre.match(/\((\d+)\)/);
  const numSAP   = numMatch ? numMatch[1] : ordre;

  const nature = typeOrdre.includes('PREV') ? 'Préventif'
               : typeOrdre.includes('CORR') ? 'Correctif'
               : typeOrdre.includes('EXPL') ? 'Exploitation'
               : 'Autre';

  return {
    ordre, numSAP, nature, typeOrdre, statut,
    statutSimple:   statutLabel(statut),
    statutOrdre,
    typeTravail:    travail,
    ouvrage:        classifyOuvrage(gamme, ordre, travail),
    ville:          String(row['Ville normalisée'] || ''),
    entite:         shortEntite(String(row['Entité en charge'] || '')),
    dateRef:        parseDate(row['Date de référence']),
    dateDebut:      parseDate(row['Date de début']),
    rue:            String(row['Nº de rue'] || '') + ' ' + String(row['Rue'] || ''),
    gamme, termine,
    enRetard:       isEnRetard(statut, row['Date de référence'], statutOrdre),
    attenteReferent,
    incoherence,
  };
}

window.GRDFDashV3 = window.GRDFDashV3 || {};
window.GRDFDashV3.otUtils = { parseDate, isEnRetard, statutLabel, shortEntite, classifyOuvrage, parseRow };
