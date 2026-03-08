/* ============================================================
   V3 - KPI + plan d'activite (loaded before dashboard.js)
   ============================================================ */

(function () {
// ── Modèles d'activité théorique par secteur ─────────────────
// Source : objectifs_cumul_par_ouvrage_avec_TDR_3onglets.xlsx
// BDSN = total agence, SAR = secteur Argenteuil, VLG = secteur Villeneuve

const ACTIVITY_MODEL_BDSN = {
   1: { BRC:   0, ROB:   0, PDL:0, PDR:0, SIN:0, TDR:  0 },
   2: { BRC: 103, ROB:  84, PDL:0, PDR:0, SIN:0, TDR:  0 },
   3: { BRC: 111, ROB:  84, PDL:0, PDR:0, SIN:0, TDR:  3 },
   4: { BRC: 103, ROB:  84, PDL:0, PDR:0, SIN:0, TDR:  0 },
   5: { BRC:  33, ROB:  84, PDL:0, PDR:0, SIN:0, TDR:  0 },
   6: { BRC: 103, ROB:  56, PDL:0, PDR:0, SIN:0, TDR:  0 },
   7: { BRC:  53, ROB:  56, PDL:0, PDR:0, SIN:0, TDR:  3 },
   8: { BRC:  22, ROB:  28, PDL:0, PDR:0, SIN:0, TDR:  0 },
   9: { BRC:  22, ROB:  28, PDL:0, PDR:0, SIN:0, TDR:  0 },
  10: { BRC:  45, ROB:  56, PDL:0, PDR:0, SIN:1, TDR:  0 },
  11: { BRC:  45, ROB:  56, PDL:0, PDR:0, SIN:1, TDR:  0 },
  12: { BRC:  52, ROB:  56, PDL:0, PDR:0, SIN:0, TDR:  3 },
  13: { BRC:  45, ROB:  56, PDL:0, PDR:0, SIN:0, TDR:  0 },
  14: { BRC:  16, ROB:  56, PDL:0, PDR:0, SIN:1, TDR: 24 },
  15: { BRC:  45, ROB:  56, PDL:0, PDR:0, SIN:1, TDR: 24 },
  16: { BRC:  52, ROB:  14, PDL:0, PDR:0, SIN:1, TDR: 27 },
  17: { BRC:  22, ROB:   5, PDL:0, PDR:0, SIN:1, TDR:  4 },
  18: { BRC:  22, ROB:   5, PDL:0, PDR:0, SIN:1, TDR:  4 },
  19: { BRC:   7, ROB:  12, PDL:0, PDR:0, SIN:1, TDR: 23 },
  20: { BRC:  43, ROB:  12, PDL:0, PDR:0, SIN:1, TDR: 26 },
  21: { BRC:  24, ROB:  12, PDL:0, PDR:0, SIN:1, TDR: 22 },
  22: { BRC:  24, ROB:  12, PDL:0, PDR:0, SIN:1, TDR: 22 },
  23: { BRC:   7, ROB:  12, PDL:0, PDR:0, SIN:1, TDR: 22 },
  24: { BRC:  24, ROB:  12, PDL:0, PDR:0, SIN:1, TDR: 21 },
  25: { BRC:  31, ROB:  12, PDL:0, PDR:0, SIN:1, TDR: 24 },
  26: { BRC:  24, ROB:  12, PDL:0, PDR:0, SIN:1, TDR: 21 },
  27: { BRC:   7, ROB:  12, PDL:0, PDR:0, SIN:1, TDR: 21 },
  28: { BRC:  24, ROB:  12, PDL:0, PDR:0, SIN:1, TDR: 22 },
  29: { BRC:  31, ROB:  12, PDL:0, PDR:0, SIN:0, TDR: 23 },
  30: { BRC:  24, ROB:  12, PDL:0, PDR:0, SIN:0, TDR: 21 },
  31: { BRC:  24, ROB:   0, PDL:0, PDR:0, SIN:0, TDR:  4 },
  32: { BRC:  24, ROB:   0, PDL:0, PDR:0, SIN:0, TDR:  4 },
  33: { BRC:  24, ROB:   0, PDL:0, PDR:0, SIN:0, TDR:  4 },
  34: { BRC:  24, ROB:   0, PDL:0, PDR:0, SIN:0, TDR:  4 },
  35: { BRC:  14, ROB:  12, PDL:0, PDR:0, SIN:0, TDR: 23 },
  36: { BRC:  24, ROB:  12, PDL:0, PDR:0, SIN:0, TDR: 21 },
  37: { BRC:  24, ROB:  12, PDL:0, PDR:0, SIN:1, TDR: 21 },
  38: { BRC:  24, ROB:  12, PDL:0, PDR:0, SIN:1, TDR: 20 },
  39: { BRC:  13, ROB:  12, PDL:0, PDR:0, SIN:1, TDR: 20 },
  40: { BRC:  24, ROB:  12, PDL:0, PDR:0, SIN:1, TDR: 19 },
  41: { BRC:  24, ROB:  12, PDL:0, PDR:0, SIN:1, TDR: 18 },
  42: { BRC:  24, ROB:  12, PDL:0, PDR:0, SIN:0, TDR: 18 },
  43: { BRC:  12, ROB:   5, PDL:0, PDR:0, SIN:0, TDR:  4 },
  44: { BRC:  12, ROB:   5, PDL:0, PDR:0, SIN:0, TDR:  1 },
  45: { BRC:  30, ROB:   0, PDL:0, PDR:0, SIN:0, TDR: 11 },
  46: { BRC:  24, ROB:   0, PDL:0, PDR:0, SIN:0, TDR:  0 },
  47: { BRC:  24, ROB:   0, PDL:0, PDR:0, SIN:0, TDR:  0 },
  48: { BRC:  24, ROB:   0, PDL:0, PDR:0, SIN:0, TDR:  0 },
  49: { BRC:   0, ROB:   0, PDL:0, PDR:0, SIN:0, TDR:  0 },
  50: { BRC:   6, ROB:   0, PDL:0, PDR:0, SIN:0, TDR:  1 },
  51: { BRC:   0, ROB:   0, PDL:0, PDR:0, SIN:0, TDR:  0 },
  52: { BRC:   0, ROB:   0, PDL:0, PDR:0, SIN:0, TDR:  0 },
  53: { BRC:   0, ROB:   0, PDL:0, PDR:0, SIN:0, TDR:  0 },
};

const ACTIVITY_MODEL_SAR = {
   1: { BRC:  0, ROB:  0, PDL:0, PDR:0, SIN:0, TDR: 0 },
   2: { BRC: 33, ROB: 24, PDL:0, PDR:0, SIN:0, TDR: 0 },
   3: { BRC: 35, ROB: 24, PDL:0, PDR:0, SIN:0, TDR: 1 },
   4: { BRC: 33, ROB: 24, PDL:0, PDR:0, SIN:0, TDR: 0 },
   5: { BRC: 33, ROB: 24, PDL:0, PDR:0, SIN:0, TDR: 0 },
   6: { BRC: 33, ROB: 16, PDL:0, PDR:0, SIN:0, TDR: 0 },
   7: { BRC: 18, ROB: 16, PDL:0, PDR:0, SIN:0, TDR: 1 },
   8: { BRC:  8, ROB:  8, PDL:0, PDR:0, SIN:0, TDR: 0 },
   9: { BRC:  8, ROB:  8, PDL:0, PDR:0, SIN:0, TDR: 0 },
  10: { BRC: 16, ROB: 16, PDL:0, PDR:0, SIN:1, TDR: 0 },
  11: { BRC: 16, ROB: 16, PDL:0, PDR:0, SIN:1, TDR: 0 },
  12: { BRC: 18, ROB: 16, PDL:0, PDR:0, SIN:0, TDR: 1 },
  13: { BRC: 16, ROB: 16, PDL:0, PDR:0, SIN:0, TDR: 0 },
  14: { BRC: 16, ROB: 16, PDL:0, PDR:0, SIN:0, TDR: 5 },
  15: { BRC: 16, ROB: 16, PDL:0, PDR:0, SIN:0, TDR: 5 },
  16: { BRC: 18, ROB:  4, PDL:0, PDR:0, SIN:0, TDR: 6 },
  17: { BRC:  8, ROB:  1, PDL:0, PDR:0, SIN:0, TDR: 1 },
  18: { BRC:  8, ROB:  1, PDL:0, PDR:0, SIN:0, TDR: 1 },
  19: { BRC:  7, ROB:  4, PDL:0, PDR:0, SIN:0, TDR: 4 },
  20: { BRC:  9, ROB:  4, PDL:0, PDR:0, SIN:1, TDR: 5 },
  21: { BRC:  7, ROB:  4, PDL:0, PDR:0, SIN:1, TDR: 4 },
  22: { BRC:  7, ROB:  4, PDL:0, PDR:0, SIN:0, TDR: 4 },
  23: { BRC:  7, ROB:  4, PDL:0, PDR:0, SIN:0, TDR: 4 },
  24: { BRC:  7, ROB:  4, PDL:0, PDR:0, SIN:0, TDR: 4 },
  25: { BRC:  9, ROB:  4, PDL:0, PDR:0, SIN:0, TDR: 5 },
  26: { BRC:  7, ROB:  4, PDL:0, PDR:0, SIN:1, TDR: 4 },
  27: { BRC:  7, ROB:  4, PDL:0, PDR:0, SIN:1, TDR: 4 },
  28: { BRC:  7, ROB:  4, PDL:0, PDR:0, SIN:0, TDR: 4 },
  29: { BRC:  9, ROB:  4, PDL:0, PDR:0, SIN:0, TDR: 4 },
  30: { BRC:  7, ROB:  4, PDL:0, PDR:0, SIN:0, TDR: 4 },
  31: { BRC:  7, ROB:  0, PDL:0, PDR:0, SIN:0, TDR: 1 },
  32: { BRC:  7, ROB:  0, PDL:0, PDR:0, SIN:0, TDR: 1 },
  33: { BRC:  7, ROB:  0, PDL:0, PDR:0, SIN:0, TDR: 1 },
  34: { BRC:  7, ROB:  0, PDL:0, PDR:0, SIN:0, TDR: 1 },
  35: { BRC:  9, ROB:  4, PDL:0, PDR:0, SIN:0, TDR: 4 },
  36: { BRC:  7, ROB:  4, PDL:0, PDR:0, SIN:0, TDR: 4 },
  37: { BRC:  7, ROB:  4, PDL:0, PDR:0, SIN:0, TDR: 4 },
  38: { BRC:  7, ROB:  4, PDL:0, PDR:0, SIN:0, TDR: 3 },
  39: { BRC:  8, ROB:  4, PDL:0, PDR:0, SIN:0, TDR: 3 },
  40: { BRC:  7, ROB:  4, PDL:0, PDR:0, SIN:0, TDR: 3 },
  41: { BRC:  7, ROB:  4, PDL:0, PDR:0, SIN:0, TDR: 3 },
  42: { BRC:  7, ROB:  4, PDL:0, PDR:0, SIN:0, TDR: 2 },
  43: { BRC:  4, ROB:  1, PDL:0, PDR:0, SIN:0, TDR: 1 },
  44: { BRC:  4, ROB:  1, PDL:0, PDR:0, SIN:0, TDR: 0 },
  45: { BRC:  8, ROB:  0, PDL:0, PDR:0, SIN:0, TDR: 1 },
  46: { BRC:  7, ROB:  0, PDL:0, PDR:0, SIN:0, TDR: 0 },
  47: { BRC:  7, ROB:  0, PDL:0, PDR:0, SIN:0, TDR: 0 },
  48: { BRC:  7, ROB:  0, PDL:0, PDR:0, SIN:0, TDR: 0 },
  49: { BRC:  0, ROB:  0, PDL:0, PDR:0, SIN:0, TDR: 0 },
  50: { BRC:  1, ROB:  0, PDL:0, PDR:0, SIN:0, TDR: 0 },
  51: { BRC:  0, ROB:  0, PDL:0, PDR:0, SIN:0, TDR: 0 },
  52: { BRC:  0, ROB:  0, PDL:0, PDR:0, SIN:0, TDR: 0 },
  53: { BRC:  0, ROB:  0, PDL:0, PDR:0, SIN:0, TDR: 0 },
};

const ACTIVITY_MODEL_VLG = {
   1: { BRC:  0, ROB:  0, PDL:0, PDR:0, SIN:0, TDR:  0 },
   2: { BRC: 71, ROB: 60, PDL:0, PDR:0, SIN:0, TDR:  0 },
   3: { BRC: 77, ROB: 60, PDL:0, PDR:0, SIN:0, TDR:  2 },
   4: { BRC: 71, ROB: 60, PDL:0, PDR:0, SIN:0, TDR:  0 },
   5: { BRC:  0, ROB: 60, PDL:0, PDR:0, SIN:0, TDR:  0 },
   6: { BRC: 71, ROB: 40, PDL:0, PDR:0, SIN:0, TDR:  0 },
   7: { BRC: 34, ROB: 40, PDL:0, PDR:0, SIN:0, TDR:  2 },
   8: { BRC: 14, ROB: 20, PDL:0, PDR:0, SIN:0, TDR:  0 },
   9: { BRC: 14, ROB: 20, PDL:0, PDR:0, SIN:0, TDR:  0 },
  10: { BRC: 28, ROB: 40, PDL:0, PDR:0, SIN:0, TDR:  0 },
  11: { BRC: 28, ROB: 40, PDL:0, PDR:0, SIN:0, TDR:  0 },
  12: { BRC: 33, ROB: 40, PDL:0, PDR:0, SIN:0, TDR:  2 },
  13: { BRC: 28, ROB: 40, PDL:0, PDR:0, SIN:0, TDR:  0 },
  14: { BRC:  0, ROB: 40, PDL:0, PDR:0, SIN:1, TDR: 19 },
  15: { BRC: 28, ROB: 40, PDL:0, PDR:0, SIN:1, TDR: 19 },
  16: { BRC: 33, ROB: 10, PDL:0, PDR:0, SIN:1, TDR: 21 },
  17: { BRC: 14, ROB:  4, PDL:0, PDR:0, SIN:1, TDR:  3 },
  18: { BRC: 14, ROB:  4, PDL:0, PDR:0, SIN:1, TDR:  3 },
  19: { BRC:  0, ROB:  9, PDL:0, PDR:0, SIN:1, TDR: 19 },
  20: { BRC: 33, ROB:  9, PDL:0, PDR:0, SIN:1, TDR: 21 },
  21: { BRC: 16, ROB:  9, PDL:0, PDR:0, SIN:1, TDR: 18 },
  22: { BRC: 16, ROB:  9, PDL:0, PDR:0, SIN:1, TDR: 18 },
  23: { BRC:  0, ROB:  9, PDL:0, PDR:0, SIN:1, TDR: 18 },
  24: { BRC: 16, ROB:  9, PDL:0, PDR:0, SIN:1, TDR: 17 },
  25: { BRC: 21, ROB:  9, PDL:0, PDR:0, SIN:1, TDR: 19 },
  26: { BRC: 16, ROB:  9, PDL:0, PDR:0, SIN:1, TDR: 17 },
  27: { BRC:  0, ROB:  9, PDL:0, PDR:0, SIN:1, TDR: 17 },
  28: { BRC: 16, ROB:  9, PDL:0, PDR:0, SIN:1, TDR: 18 },
  29: { BRC: 21, ROB:  9, PDL:0, PDR:0, SIN:0, TDR: 19 },
  30: { BRC: 16, ROB:  9, PDL:0, PDR:0, SIN:0, TDR: 17 },
  31: { BRC: 16, ROB:  0, PDL:0, PDR:0, SIN:0, TDR:  3 },
  32: { BRC: 16, ROB:  0, PDL:0, PDR:0, SIN:0, TDR:  3 },
  33: { BRC: 16, ROB:  0, PDL:0, PDR:0, SIN:0, TDR:  3 },
  34: { BRC: 16, ROB:  0, PDL:0, PDR:0, SIN:0, TDR:  3 },
  35: { BRC:  5, ROB:  9, PDL:0, PDR:0, SIN:0, TDR: 19 },
  36: { BRC: 16, ROB:  9, PDL:0, PDR:0, SIN:0, TDR: 17 },
  37: { BRC: 16, ROB:  9, PDL:0, PDR:0, SIN:1, TDR: 17 },
  38: { BRC: 16, ROB:  9, PDL:0, PDR:0, SIN:1, TDR: 17 },
  39: { BRC:  5, ROB:  9, PDL:0, PDR:0, SIN:1, TDR: 17 },
  40: { BRC: 16, ROB:  9, PDL:0, PDR:0, SIN:1, TDR: 16 },
  41: { BRC: 16, ROB:  9, PDL:0, PDR:0, SIN:1, TDR: 15 },
  42: { BRC: 16, ROB:  9, PDL:0, PDR:0, SIN:0, TDR: 16 },
  43: { BRC:  8, ROB:  4, PDL:0, PDR:0, SIN:0, TDR:  3 },
  44: { BRC:  8, ROB:  4, PDL:0, PDR:0, SIN:0, TDR:  1 },
  45: { BRC: 21, ROB:  0, PDL:0, PDR:0, SIN:0, TDR: 10 },
  46: { BRC: 16, ROB:  0, PDL:0, PDR:0, SIN:0, TDR:  0 },
  47: { BRC: 16, ROB:  0, PDL:0, PDR:0, SIN:0, TDR:  0 },
  48: { BRC: 16, ROB:  0, PDL:0, PDR:0, SIN:0, TDR:  0 },
  49: { BRC:  0, ROB:  0, PDL:0, PDR:0, SIN:0, TDR:  0 },
  50: { BRC:  5, ROB:  0, PDL:0, PDR:0, SIN:0, TDR:  1 },
  51: { BRC:  0, ROB:  0, PDL:0, PDR:0, SIN:0, TDR:  0 },
  52: { BRC:  0, ROB:  0, PDL:0, PDR:0, SIN:0, TDR:  0 },
  53: { BRC:  0, ROB:  0, PDL:0, PDR:0, SIN:0, TDR:  0 },
};

// Alias pour compatibilité — retourne le bon modèle selon le filtre entité actif
function getActivityModel() {
  const entite = document.getElementById('filter-entite')?.value || '';
  if (entite.includes('SAR')) return ACTIVITY_MODEL_SAR;
  if (entite.includes('VLG')) return ACTIVITY_MODEL_VLG;
  return ACTIVITY_MODEL_BDSN;
}

// Retourne les semaines triées du modèle actif
function getModelWeeks() {
  return Object.keys(getActivityModel()).map(Number).sort((a,b)=>a-b);
}

// Mapping type ouvrage dashboard → clé dans ACTIVITY_MODEL
const OUVRAGE_TO_MODEL = {
  'CICM / BRC':             'BRC',
  'Robinet Réseau':         'ROB',
  'PDL':                    'PDL',
  'PDR':                    'PDR',
  'SIN (Point singulier)':  'SIN',
  'TDR - Remplacement':     'TDR',
  'TDR - Visite préalable': 'TDR',
};

// Retourne le numéro de semaine ISO d'une date
function getISOWeek(d) {
  const date = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  const dayNum = date.getUTCDay() || 7;
  date.setUTCDate(date.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
  return Math.ceil((((date - yearStart) / 86400000) + 1) / 7);
}


function renderKPIs() {
  const total          = FILTERED.length;
  const term           = FILTERED.filter(d=>d.statut==='TERM').length;
  const termAno        = FILTERED.filter(d=>d.statut==='TERM ANO').length;
  const encours        = FILTERED.filter(d=>['ACTI','EPR','PRG','OGDI'].includes(d.statut)).length;
  const retard         = FILTERED.filter(d=>d.enRetard).length;
  const attenteRef     = FILTERED.filter(d=>d.attenteReferent).length;
  const taux           = total>0 ? ((term+termAno)/total*100).toFixed(1) : '—';

  document.getElementById('kpi-total').textContent       = total.toLocaleString('de-DE');
  document.getElementById('kpi-taux').textContent        = total>0 ? taux+' %' : '—';
  document.getElementById('kpi-retard').textContent      = retard.toLocaleString('fr-FR');
  document.getElementById('kpi-anomalie').textContent    = termAno.toLocaleString('fr-FR');
  document.getElementById('kpi-encours').textContent     = encours.toLocaleString('fr-FR');
  document.getElementById('kpi-attente').textContent     = attenteRef.toLocaleString('fr-FR');

  const trends = window.GRDFDashV3.services.weeklyTrends.computeWeeklyTrends();
  const realises = trends.preventif.realises;
  const tone = realises.trend === 'up' ? 'positive' : realises.trend === 'down' ? 'negative' : 'neutral';
  window.GRDFDashV3.services.weeklyTrends.upsertKpiExtra(
    'kpi-total',
    'prev-realises-week',
    `Réalisés semaine: ${realises.current} · ${window.GRDFDashV3.services.weeklyTrends.formatVsS1(realises)}`,
    tone
  );
}

function computePlanActivite() {
  const now = new Date();
  const currentWeek = getISOWeek(now);
  const currentYear = now.getFullYear();

  const allWeeks = getModelWeeks();
  const mappedOuvrages = Object.keys(OUVRAGE_TO_MODEL);

  // ── Déterminer quelles colonnes du modèle utiliser selon le filtre ouvrage actif ──
  // Si un filtre ouvrage est actif ET qu'il est dans le périmètre plan, on restreint
  // Sinon on prend toutes les colonnes du périmètre
  const filtreOuvrage = document.getElementById('filter-ouvrage')?.value || '';
  let colonnesPlan; // tableau de clés : ['BRC'], ['ROB'], ou ['BRC','ROB','PDL','PDR','SIN']

  if (filtreOuvrage && OUVRAGE_TO_MODEL[filtreOuvrage]) {
    // Filtre actif sur un ouvrage du périmètre → plan = uniquement cette colonne
    colonnesPlan = [OUVRAGE_TO_MODEL[filtreOuvrage]];
  } else if (filtreOuvrage && !OUVRAGE_TO_MODEL[filtreOuvrage]) {
    // Filtre actif sur un ouvrage hors périmètre (ex: TDR) → pas de plan applicable
    colonnesPlan = [];
  } else {
    // Pas de filtre → toutes les colonnes du périmètre
    colonnesPlan = ['BRC','ROB','PDL','PDR','SIN'];
  }

  // Fonction helper : somme des colonnes actives pour une semaine
  const planSemaine = (r) => colonnesPlan.reduce((acc, col) => acc + (r[col]||0), 0);

  // ── Plan cumulé jusqu'à la semaine courante ──
  let theoriqueTotal = 0;
  allWeeks.filter(w => w <= currentWeek).forEach(w => {
    theoriqueTotal += planSemaine(getActivityModel()[w]);
  });

  // ── Plan annuel total ──
  const annuelTheo = allWeeks.reduce((acc, w) => acc + planSemaine(getActivityModel()[w]), 0);

  // ── Réel : OT terminés du périmètre (respectant le filtre ouvrage), ventilés par dateDebut ──
  const reelParSemaine = {};
  FILTERED.forEach(d => {
    if (!d.termine) return;
    // Si filtre ouvrage actif hors périmètre → pas de réel plan non plus
    if (colonnesPlan.length === 0) return;
    // Si filtre ouvrage actif dans périmètre → uniquement cet ouvrage
    // Si pas de filtre → tous les ouvrages mappés
    if (filtreOuvrage && OUVRAGE_TO_MODEL[filtreOuvrage]) {
      if (d.ouvrage !== filtreOuvrage) return;
    } else {
      if (!mappedOuvrages.includes(d.ouvrage)) return;
    }
    const date = d.dateDebut || d.dateRef;
    if (!date) return;
    if (date.getFullYear() !== currentYear) return;
    const w = getISOWeek(date);
    reelParSemaine[w] = (reelParSemaine[w] || 0) + 1;
  });

  // ── Total réel cumulé jusqu'à la semaine courante ──
  const realTermine = allWeeks
    .filter(w => w <= currentWeek)
    .reduce((acc, w) => acc + (reelParSemaine[w] || 0), 0);

  // ── % réalisation et écart ──
  const pctRealisation = theoriqueTotal > 0 ? Math.round(realTermine / theoriqueTotal * 100) : 0;
  const ecart    = realTermine - Math.round(theoriqueTotal);
  const enAvance = ecart >= 0;

  // ── Données graphique : semaine par semaine ──
  let cumTheo = 0;
  let cumReel = 0;
  const chartLabels = [];
  const chartTheo   = [];
  const chartReel   = [];

  allWeeks.forEach(w => {
    cumTheo += planSemaine(getActivityModel()[w]);
    chartLabels.push('S' + w);
    chartTheo.push(Math.round(cumTheo));

    if (w <= currentWeek) {
      cumReel += (reelParSemaine[w] || 0);
      chartReel.push(cumReel);
    } else {
      chartReel.push(null);
    }
  });

  return {
    currentWeek,
    theoriqueTotal: Math.round(theoriqueTotal),
    annuelTheo:     Math.round(annuelTheo),
    realTermine,
    pctRealisation,
    ecart,
    enAvance,
    chartLabels,
    chartTheo,
    chartReel,
    colonnesPlan,
    horsPérimètre: colonnesPlan.length === 0,
  };
}

function renderPlanActivite() {
  const plan = computePlanActivite();

  const alertEl    = document.getElementById('plan-alert');
  const kpiTheoEl  = document.getElementById('plan-kpi-theorique');
  const kpiReelEl  = document.getElementById('plan-kpi-realise');
  const kpiPctEl   = document.getElementById('plan-kpi-pct');
  const kpiEcartEl = document.getElementById('plan-kpi-ecart');

  // ── Cas : ouvrage filtré hors périmètre plan ──
  if (plan.horsPérimètre) {
    kpiTheoEl.textContent  = '—'; kpiReelEl.textContent = '—';
    kpiPctEl.textContent   = '—'; kpiPctEl.style.color  = C.gray;
    kpiEcartEl.textContent = '—'; kpiEcartEl.style.color = C.gray;
    if (kpiTheoEl.nextElementSibling)
      kpiTheoEl.nextElementSibling.textContent = 'Hors périmètre plan d\'activité';
    alertEl.className = 'plan-alert plan-alert-warn';
    alertEl.innerHTML = `ℹ️ <strong>Hors périmètre</strong> — Ce type d'ouvrage n'a pas d'objectif dans le plan d'activité`;
    destroyChart('planActivite');
    return;
  }

  // ── KPIs ──
  const kpiColor = plan.enAvance ? C.green : C.red;
  const ecartTxt = (plan.enAvance ? '+' : '') + plan.ecart + ' OT';
  const pctColor = plan.pctRealisation >= 90 ? C.green : plan.pctRealisation >= 70 ? C.orange : C.red;

  kpiTheoEl.textContent  = plan.theoriqueTotal.toLocaleString('fr-FR');
  kpiReelEl.textContent  = plan.realTermine.toLocaleString('fr-FR');
  kpiPctEl.textContent   = plan.pctRealisation + ' %';
  kpiPctEl.style.color   = pctColor;
  kpiEcartEl.textContent = ecartTxt;
  kpiEcartEl.style.color = kpiColor;

  if (kpiTheoEl.nextElementSibling)
    kpiTheoEl.nextElementSibling.textContent =
      `Plan cumulé S${plan.currentWeek} · objectif annuel : ${plan.annuelTheo.toLocaleString('fr-FR')} OT`;

  // ── Badge alerte ──
  if (plan.pctRealisation >= 90) {
    alertEl.className = 'plan-alert plan-alert-ok';
    alertEl.innerHTML = `✅ <strong>Dans le plan</strong> — ${plan.pctRealisation}% de l'objectif cumulé atteint à S${plan.currentWeek}`;
  } else if (plan.pctRealisation >= 70) {
    alertEl.className = 'plan-alert plan-alert-warn';
    alertEl.innerHTML = `⚠️ <strong>Légèrement en retard</strong> — ${plan.pctRealisation}% de l'objectif cumulé S${plan.currentWeek} (écart : ${Math.abs(plan.ecart)} OT)`;
  } else {
    alertEl.className = 'plan-alert plan-alert-danger';
    alertEl.innerHTML = `🔴 <strong>En retard sur le plan</strong> — ${plan.pctRealisation}% de l'objectif cumulé S${plan.currentWeek} (écart : ${Math.abs(plan.ecart)} OT)`;
  }

  // ── Graphique ──
  destroyChart('planActivite');
  const ctx = document.getElementById('chartPlanActivite').getContext('2d');

  // ── Zoom : fenêtre de semaines à afficher ──
  const zoom = window._planZoom || 0; // 0 = année complète, N = currentWeek±N
  let idxMin = 0, idxMax = plan.chartLabels.length - 1;
  if (zoom > 0) {
    const ci = plan.chartLabels.indexOf('S' + plan.currentWeek);
    idxMin = Math.max(0, ci - zoom);
    idxMax = Math.min(plan.chartLabels.length - 1, ci + zoom);
  }
  const labels  = plan.chartLabels.slice(idxMin, idxMax + 1);
  const dTheo   = plan.chartTheo.slice(idxMin, idxMax + 1);
  const dReel   = plan.chartReel.slice(idxMin, idxMax + 1);

  // ── Labels mois : afficher le mois quand il change ──
  // Semaine ISO → date approximative (lundi de la semaine)
  const MOIS_COURT = ['Jan','Fév','Mar','Avr','Mai','Jun','Jul','Aoû','Sep','Oct','Nov','Déc'];
  function weekToDate(w) {
    const d = new Date(2026, 0, 4); // 4 jan 2026 = S1
    d.setDate(d.getDate() + (w - 1) * 7);
    return d;
  }
  // Construire les labels avec mois intercalés
  let lastMonth = -1;
  const tickLabels = labels.map(lbl => {
    const w = parseInt(lbl.replace('S',''));
    const d = weekToDate(w);
    const m = d.getMonth();
    if (m !== lastMonth) { lastMonth = m; return `${lbl}\n${MOIS_COURT[m]}`; }
    return lbl;
  });

  CHARTS['planActivite'] = new Chart(ctx, {
    type: 'line',
    data: {
      labels: tickLabels,
      datasets: [
        {
          label: 'Plan théorique (cumulé)',
          data: dTheo,
          borderColor: C.blue,
          backgroundColor: 'rgba(0,49,137,0.06)',
          borderWidth: 2,
          borderDash: [6, 3],
          pointRadius: 0,
          pointHoverRadius: 5,
          fill: false,
          tension: 0.2,
          spanGaps: true,
        },
        {
          label: 'Réalisé (cumulé)',
          data: dReel,
          borderColor: C.green,
          backgroundColor: 'rgba(27,122,62,0.10)',
          borderWidth: 2.5,
          pointRadius: ctx2 => {
            const ci = labels.indexOf('S' + plan.currentWeek);
            return ctx2.dataIndex === ci ? 6 : 0;
          },
          pointHoverRadius: 6,
          pointBackgroundColor: C.green,
          pointBorderColor: '#fff',
          pointBorderWidth: 2,
          fill: true,
          tension: 0.2,
          spanGaps: false,
        },
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode:'index', intersect:false },
      plugins: {
        legend: { position:'bottom', labels:{ font:{family:'DM Sans',size:11}, boxWidth:14, padding:14 } },
        tooltip: {
          filter: item => item.raw !== null,
          callbacks: {
            title: items => `Semaine ${labels[items[0].dataIndex]}`,
            label: ctx2 => {
              if (ctx2.raw === null) return null;
              if (ctx2.datasetIndex === 0) {
                const reel = dReel[ctx2.dataIndex];
                if (reel !== null) {
                  const diff = reel - ctx2.raw;
                  const sign = diff >= 0 ? '+' : '';
                  return ` Plan : ${ctx2.raw} OT  (écart réel : ${sign}${diff})`;
                }
                return ` Plan : ${ctx2.raw} OT prévus (cumulé)`;
              }
              const theo = dTheo[ctx2.dataIndex];
              const diff = ctx2.raw - theo;
              const sign = diff >= 0 ? '+' : '';
              return ` Réalisé : ${ctx2.raw} OT  (${sign}${diff} vs plan)`;
            }
          }
        }
      },
      scales: {
        x: {
          grid: { color:'#f1f5f9' },
          ticks: {
            font: ctx2 => {
              // Gras pour les labels avec le mois (contiennent \n)
              const lbl = tickLabels[ctx2.index] || '';
              return { family:'DM Sans', size: lbl.includes('\n') ? 10 : 9,
                       weight: lbl.includes('\n') ? '600' : '400' };
            },
            maxRotation: 0,
            maxTicksLimit: zoom > 0 ? idxMax - idxMin + 1 : 22,
            color: ctx2 => {
              const lbl = tickLabels[ctx2.index] || '';
              return lbl.includes('\n') ? '#1e293b' : '#94a3b8';
            },
          },
          afterDraw(chart) {
            const ci = labels.indexOf('S' + plan.currentWeek);
            if (ci < 0) return;
            const { ctx: c, chartArea:{ top, bottom } } = chart;
            const x = chart.scales.x.getPixelForIndex(ci);
            c.save();
            c.beginPath(); c.setLineDash([4,3]);
            c.strokeStyle = '#94a3b8'; c.lineWidth = 1.2;
            c.moveTo(x, top); c.lineTo(x, bottom); c.stroke();
            c.setLineDash([]);
            c.font = 'bold 10px DM Sans,sans-serif';
            c.fillStyle = '#64748b'; c.textAlign = 'center';
            c.fillText("Aujourd'hui", x, top - 4);
            c.restore();
          }
        },
        y: {
          grid: { color:'#f1f5f9' },
          ticks: { font:{family:'DM Sans',size:10} },
          title: { display:true, text:'OT cumulés', font:{family:'DM Sans',size:10}, color:'#94a3b8' }
        }
      }
    }
  });
}

// ── Zoom graphique plan d'activité ────────────────────────────
function setPlanZoom(n) {
  window._planZoom = n;
  // Mettre à jour les boutons actifs
  document.querySelectorAll('.plan-zoom-btn').forEach(b => {
    const v = b.getAttribute('onclick').match(/\d+/);
    const val = v ? +v[0] : 0;
    b.classList.toggle('active', val === n);
  });
  renderPlanActivite();
}
  window.GRDFDashV3 = window.GRDFDashV3 || {};
  window.GRDFDashV3.features = window.GRDFDashV3.features || {};
  window.GRDFDashV3.features.kpiPlan = {
    ACTIVITY_MODEL_BDSN, ACTIVITY_MODEL_SAR, ACTIVITY_MODEL_VLG, OUVRAGE_TO_MODEL,
    getActivityModel, getModelWeeks, getISOWeek,
    renderKPIs, computePlanActivite, renderPlanActivite, setPlanZoom,
  };
})();

