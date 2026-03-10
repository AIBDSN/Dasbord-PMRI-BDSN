/* ============================================================
   GRDF MAINTENANCE DASHBOARD — dashboard.js
   ============================================================ */

let ALL_DATA = [];
let FILTERED  = [];
let FILTERED_PREV = [];   // préventif filtré
let FILTERED_COR  = [];   // correctif filtré
let CHARTS    = {};
let TABLE_MODE = 'retard';
let COR_TABLE_MODE = 'cor-ouverts';
let CURRENT_MODE = 'preventif'; // 'preventif' | 'correctif'
// ── Variables globales Avis PM ────────────────────────────────
let AVIS_DATA = [];
let AVIS_MODE = 'f0';
const PRIO_ORD_AVIS  = { F0:0, P1:1, P2:2, R1:3, R2:4, R3:5, Autre:6 };
const PRIO_ORD_SORT  = { F0:0, P1:1, P2:2, R1:3, R2:4, R3:5, Autre:6 };
const SORT_STATE     = { prev:{col:null,dir:1}, cor:{col:null,dir:1}, avis:{col:null,dir:1} };
let ROWS_PREV = [], ROWS_COR = [], ROWS_AVIS = [];


const C = {
  blue:'#003189', blueMid:'#1a4faf', blueLight:'#93b4e8',
  green:'#1b7a3e', greenLight:'#6dbf8a',
  orange:'#c95f00', orangeLight:'#f4a35a',
  red:'#b91c1c', redLight:'#e87777',
  gray:'#94a3b8', purple:'#6d28d9', teal:'#0f766e',
};
const PALETTE = [C.blue,C.green,C.orange,C.red,C.purple,C.teal,C.blueLight,C.greenLight,C.orangeLight,C.redLight,C.gray];

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

// ── Chargement fichier ────────────────────────────────────────
function loadFile(input) {
  const file = input.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const wb  = XLSX.read(e.target.result, { type:'array', cellDates:true });
      const ws  = wb.Sheets[wb.SheetNames[0]];
      const raw = XLSX.utils.sheet_to_json(ws, { defval:'' });
      ALL_DATA  = raw.map(parseRow);
      document.getElementById('file-meta').textContent = `${file.name} · ${ALL_DATA.length} OT`;
      document.getElementById('drop-label').textContent = '📂 Changer le fichier';

      // ── Date d'extraction GOTAM = lastModified du fichier ──
      const extractDate = new Date(file.lastModified);
      const extractStr = extractDate.toLocaleDateString('fr-FR', {
        day:'2-digit', month:'2-digit', year:'numeric'
      }) + ' à ' + extractDate.toLocaleTimeString('fr-FR', {
        hour:'2-digit', minute:'2-digit'
      });
      document.getElementById('extraction-date').textContent = `Extraction GOTAM du : ${extractStr}`;
      document.getElementById('extraction-badge').style.display = 'flex';

      initFilters();
      FILTERED_PREV = ALL_DATA.filter(d => d.nature === 'Préventif' || d.nature === 'Exploitation');
      FILTERED_COR  = ALL_DATA.filter(d => d.nature === 'Correctif');
      FILTERED = FILTERED_PREV;
      CURRENT_MODE = 'preventif';
      document.getElementById('tab-prev').classList.add('active');
      document.getElementById('tab-cor').classList.remove('active');
      document.getElementById('view-preventif').style.display = '';
      document.getElementById('view-correctif').style.display = 'none';
      // Compteurs sur les onglets
      document.getElementById('tab-prev-count').textContent = FILTERED_PREV.length.toLocaleString('fr-FR');
      document.getElementById('tab-cor-count').textContent  = FILTERED_COR.length.toLocaleString('fr-FR');
      render();
      document.getElementById('empty-state').style.display = 'none';
      document.getElementById('dashboard').style.display   = 'block';
    } catch(err) { alert('Erreur lecture fichier : ' + err.message); }
  };
  reader.readAsArrayBuffer(file);
}

// ── Parsing ───────────────────────────────────────────────────
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

  // Numéro SAP court : extrait entre parenthèses ex: "BRCIS 95219 (30000253247)" → "30000253247"
  const numMatch = ordre.match(/\((\d+)\)/);
  const numSAP   = numMatch ? numMatch[1] : ordre;

  // Nature simplifiée : Préventif / Correctif / Exploitation
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

// ── Classification ouvrage ────────────────────────────────────
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

// ── Filtres ───────────────────────────────────────────────────
function initFilters() {
  populateSelect('filter-entite',  [...new Set(ALL_DATA.map(d=>d.entite))].sort());
  populateSelect('filter-ouvrage', [...new Set(ALL_DATA.map(d=>d.ouvrage))].sort());
}

function populateSelect(id, values) {
  const sel = document.getElementById(id);
  sel.innerHTML = '<option value="">Tous</option>';
  values.filter(Boolean).forEach(v => {
    const o = document.createElement('option');
    o.value = v; o.textContent = v; sel.appendChild(o);
  });
}

function applyFilters() {
  const entite  = document.getElementById('filter-entite').value;
  const ouvrage = document.getElementById('filter-ouvrage').value;

  const base = ALL_DATA.filter(d =>
    (!entite  || d.entite  === entite)  &&
    (!ouvrage || d.ouvrage === ouvrage)
  );

  FILTERED_PREV = base.filter(d => d.nature === 'Préventif' || d.nature === 'Exploitation');
  FILTERED_COR  = base.filter(d => d.nature === 'Correctif');
  FILTERED = CURRENT_MODE === 'correctif' ? FILTERED_COR : FILTERED_PREV;

  render();
}

function resetFilters() {
  ['filter-entite','filter-ouvrage'].forEach(id => document.getElementById(id).value = '');
  FILTERED_PREV = ALL_DATA.filter(d => d.nature === 'Préventif' || d.nature === 'Exploitation');
  FILTERED_COR  = ALL_DATA.filter(d => d.nature === 'Correctif');
  FILTERED = CURRENT_MODE === 'correctif' ? FILTERED_COR : FILTERED_PREV;
  render();
}

// ── Navigation entre modes ────────────────────────────────────
function setMode(mode) {
  CURRENT_MODE = mode;
  document.getElementById('tab-prev').classList.toggle('active',  mode === 'preventif');
  document.getElementById('tab-cor').classList.toggle('active',   mode === 'correctif');
  document.getElementById('tab-avis').classList.toggle('active',  mode === 'avis');
  document.getElementById('view-preventif').style.display = mode === 'preventif' ? '' : 'none';
  document.getElementById('view-correctif').style.display = mode === 'correctif' ? '' : 'none';
  document.getElementById('view-avis').style.display      = mode === 'avis'      ? '' : 'none';
  if (mode === 'avis') { renderAvis(); return; }
  FILTERED = mode === 'correctif' ? FILTERED_COR : FILTERED_PREV;
  render();
}

// ── Rendu global ──────────────────────────────────────────────
function render() {
  if (CURRENT_MODE === 'correctif') {
    renderCorrectif();
  } else {
    renderKPIs();
    renderGauges();
    renderChartStatut();
    renderChartOuvrageBar();
    renderPlanActivite();
    renderTable();
  }
}

// ── KPIs ──────────────────────────────────────────────────────
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
}

// ── Jauges speedometer ────────────────────────────────────────

// Ordre d'affichage fixe des jauges
const GAUGE_ORDER = [
  'CICM / BRC', 'DDMP', 'Robinet Réseau',
  'TDR - Remplacement', 'TDR - Visite préalable',
  'SIN (Point singulier)', 'PDL', 'PDR', 'DPBE', 'VSIC',
  'Enquête / Inventaire', 'Exploitation', 'Autre',
];

function renderGauges() {
  const container = document.getElementById('gauges-container');
  container.innerHTML = '';

  const present = [...new Set(FILTERED.map(d=>d.ouvrage))];
  const ouvrages = [
    ...GAUGE_ORDER.filter(o => present.includes(o)),
    ...present.filter(o => !GAUGE_ORDER.includes(o)).sort(),
  ];

  const now = new Date();
  const currentWeek = getISOWeek(now);
  const currentYear = now.getFullYear();
  const allWeeks = getModelWeeks().filter(w => w <= currentWeek);

  ouvrages.forEach(ouvrage => {
    const sub    = FILTERED.filter(d=>d.ouvrage===ouvrage);
    const total  = sub.length;
    const done   = sub.filter(d=>d.termine).length;
    const retard = sub.filter(d=>d.enRetard).length;

    // % clôture global → affiché sur la jauge
    const pctCloture = total > 0 ? Math.round(done / total * 100) : 0;

    // Cible plan → marqueur sur l'arc (en % du total d'OT)
    const modelKey = OUVRAGE_TO_MODEL[ouvrage];
    let pctCible = -1; // -1 = pas de marqueur
    let theoPlan = 0, reelPlan = 0;

    if (modelKey) {
      theoPlan = allWeeks.reduce((acc, w) => acc + (getActivityModel()[w][modelKey]||0), 0);
      reelPlan = sub.filter(d => {
        if (!d.termine) return false;
        const date = d.dateDebut || d.dateRef;
        return date && date.getFullYear() === currentYear;
      }).length;
      // Cible exprimée en % du total d'OT (plafonné à 100%)
      pctCible = total > 0 ? Math.min(Math.round(theoPlan / total * 100), 100) : -1;
    }

    // Couleur : comparaison % clôture vs cible plan (tolérance ±5%)
    let color;
    if (pctCible < 0) {
      color = pctCloture >= 90 ? C.green : pctCloture >= 60 ? C.orange : C.red;
    } else {
      const ecart = pctCloture - pctCible;
      color = ecart >= -5 ? C.green : ecart >= -20 ? C.orange : C.red;
    }

    // Badge plan
    let badgePlan = '';
    if (modelKey) {
      const ecart = reelPlan - Math.round(theoPlan);
      const sign  = ecart >= 0 ? '+' : '';
      const ecartColor = ecart >= 0 ? C.green : C.red;
      badgePlan = `<span class="gauge-plan-badge" style="color:${ecartColor}">Plan S${currentWeek} : ${reelPlan}/${Math.round(theoPlan)} (${sign}${ecart})</span>`;
    } else {
      badgePlan = `<span class="gauge-plan-hors">Hors périmètre plan</span>`;
    }

    const card = document.createElement('div');
    card.className = 'gauge-card';
    card.innerHTML = `
      <div class="gauge-title">${esc(ouvrage)}</div>
      <div class="gauge-wrap">
        <canvas class="gauge-canvas" width="160" height="92"
                data-pct="${pctCloture}"
                data-pct-cible="${pctCible}"
                data-color="${color}"></canvas>
      </div>
      <div class="gauge-pct" style="color:${color}">${pctCloture}%</div>
      <div class="gauge-meta">
        <span>${done} / ${total} clôturés</span>
        ${badgePlan}
        ${retard>0 ? `<span class="gauge-retard">⚠ ${retard} en retard</span>` : '<span class="gauge-ok">✓ À jour</span>'}
      </div>`;
    container.appendChild(card);
  });

  container.querySelectorAll('.gauge-canvas').forEach(canvas => {
    drawSpeedometer(canvas, +canvas.dataset.pct, canvas.dataset.color, +canvas.dataset.pctCible);
  });
}

function drawSpeedometer(canvas, pct, color, pctCible) {
  const ctx = canvas.getContext('2d');
  const w=canvas.width, h=canvas.height;
  const cx=w/2, cy=h-12, r=Math.min(w,h*2)/2-10;
  ctx.clearRect(0,0,w,h);

  // Track de fond
  ctx.beginPath(); ctx.arc(cx,cy,r,Math.PI,0,false);
  ctx.lineWidth=13; ctx.strokeStyle='#e2e8f0'; ctx.lineCap='round'; ctx.stroke();

  // Zones couleur de fond
  [{from:0,to:0.6,c:'#fca5a5'},{from:0.6,to:0.85,c:'#fed7aa'},{from:0.85,to:1,c:'#86efac'}].forEach(z => {
    ctx.beginPath(); ctx.arc(cx,cy,r,Math.PI+z.from*Math.PI,Math.PI+z.to*Math.PI,false);
    ctx.lineWidth=13; ctx.strokeStyle=z.c; ctx.lineCap='butt'; ctx.stroke();
  });

  // Arc de progression (% clôture)
  ctx.beginPath(); ctx.arc(cx,cy,r,Math.PI,Math.PI+(pct/100)*Math.PI,false);
  ctx.lineWidth=13; ctx.strokeStyle=color; ctx.lineCap='round'; ctx.stroke();

  // ── Marqueur cible plan (trait vertical sur l'arc) ──
  if (pctCible >= 0) {
    const ca = Math.PI + (pctCible/100) * Math.PI;
    const r1 = r - 8, r2 = r + 8;
    ctx.beginPath();
    ctx.moveTo(cx + r1*Math.cos(ca), cy + r1*Math.sin(ca));
    ctx.lineTo(cx + r2*Math.cos(ca), cy + r2*Math.sin(ca));
    ctx.lineWidth = 2.5; ctx.strokeStyle = '#1e293b'; ctx.lineCap = 'round'; ctx.stroke();
    // Petit losange sur le marqueur
    ctx.beginPath(); ctx.arc(cx + r*Math.cos(ca), cy + r*Math.sin(ca), 3, 0, Math.PI*2);
    ctx.fillStyle = '#1e293b'; ctx.fill();
  }

  // Aiguille (% clôture réel)
  const na = Math.PI + (pct/100)*Math.PI;
  ctx.beginPath(); ctx.moveTo(cx,cy); ctx.lineTo(cx+(r-4)*Math.cos(na),cy+(r-4)*Math.sin(na));
  ctx.lineWidth=2.5; ctx.strokeStyle='#1e293b'; ctx.lineCap='round'; ctx.stroke();

  // Centre
  ctx.beginPath(); ctx.arc(cx,cy,4,0,Math.PI*2); ctx.fillStyle='#1e293b'; ctx.fill();

  // Labels 0% / 100%
  ctx.font='8px DM Sans,sans-serif'; ctx.fillStyle='#94a3b8';
  ctx.textAlign='left';  ctx.fillText('0%',3,cy+2);
  ctx.textAlign='right'; ctx.fillText('100%',w-3,cy+2);
}

// ── Chart: Statut donut ───────────────────────────────────────
function renderChartStatut() {
  destroyChart('statut');
  const map=countBy(FILTERED,'statutSimple');
  const ctx=document.getElementById('chartStatut').getContext('2d');
  CHARTS['statut']=new Chart(ctx,{
    type:'doughnut',
    data:{labels:Object.keys(map),datasets:[{data:Object.values(map),backgroundColor:PALETTE,borderWidth:2,borderColor:'#fff'}]},
    options:{responsive:false, maintainAspectRatio:true, cutout:'62%',
      plugins:{legend:{position:'bottom',labels:{font:{family:'DM Sans',size:11},boxWidth:12,padding:10}}}}
  });
}

// ── Chart: Fait vs Reste à faire par ouvrage ──────────────────
function renderChartOuvrageBar() {
  destroyChart('ouvrageBar');
  const ouvrages = [...new Set(FILTERED.map(d=>d.ouvrage))].sort();
  const fait     = ouvrages.map(o => FILTERED.filter(d=>d.ouvrage===o && d.termine).length);
  const restant  = ouvrages.map(o => FILTERED.filter(d=>d.ouvrage===o && !d.termine).length);

  const ctx = document.getElementById('chartOuvrageBar').getContext('2d');
  CHARTS['ouvrageBar'] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: ouvrages,
      datasets: [
        { label:'Fait', data:fait,    backgroundColor:C.green,  borderRadius:3, stack:'s' },
        { label:'Reste à faire', data:restant, backgroundColor:C.redLight, borderRadius:3, stack:'s' },
      ]
    },
    options: {
      indexAxis: 'y',
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { position:'bottom', labels:{ font:{family:'DM Sans',size:11}, boxWidth:12, padding:12 } },
        tooltip: {
          callbacks: {
            label: ctx => {
              const total = fait[ctx.dataIndex] + restant[ctx.dataIndex];
              const pct = total>0 ? Math.round(ctx.raw/total*100) : 0;
              return ` ${ctx.dataset.label} : ${ctx.raw} (${pct}%)`;
            }
          }
        }
      },
      scales: {
        x: { stacked:true, grid:{color:'#f1f5f9'}, ticks:{font:{family:'DM Sans',size:10}} },
        y: { stacked:true, grid:{display:false},   ticks:{font:{family:'DM Sans',size:11}} }
      }
    }
  });
}

// ── Section : Plan d'activité ─────────────────────────────────
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
function setTableMode(mode) {
  TABLE_MODE = mode;
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.toggle('active', b.dataset.mode===mode));
  renderTable();
}

function renderTable() {
  let rows;
  const labels = {
    retard:   '🔴 OT en retard',
    afaire:   '🔵 Reste à faire — OT actifs (terrain non terminé, ordre Lancé)',
    termine:  '✅ OT terminés',
    attente:  '⏳ Attente traitement référent',
    incohere: '⚠️ Incohérences',
    all:      '📋 Tous les OT'
  };
  switch(TABLE_MODE) {
    case 'retard':   rows=FILTERED.filter(d=>d.enRetard); break;
    case 'afaire':   rows=FILTERED.filter(d=>!d.termine && d.statutOrdre.includes('Lancé')); break;
    case 'termine':  rows=FILTERED.filter(d=>d.termine); break;
    case 'attente':  rows=FILTERED.filter(d=>d.attenteReferent); break;
    case 'incohere': rows=FILTERED.filter(d=>d.incoherence); break;
    default:         rows=[...FILTERED];
  }
  rows.sort((a,b)=>(a.dateRef||0)-(b.dateRef||0));
  document.getElementById('table-section-title').textContent = `${labels[TABLE_MODE]} — ${rows.length} OT`;
  ROWS_PREV = rows;
  // Réinitialiser état tri
  SORT_STATE.prev = { col: null, dir: 1 };
  document.querySelectorAll('#table-retard thead th').forEach(h => h.classList.remove('sort-asc','sort-desc'));
  const q = document.getElementById('search-prev')?.value || '';
  _renderPrevRows(ROWS_PREV, q);
}

function renderTableRows(rows) {
  const tbody=document.getElementById('table-retard-body');
  tbody.innerHTML='';
  rows.forEach(d => {
    const tr=document.createElement('tr');
    const dateStr=d.dateRef ? d.dateRef.toLocaleDateString('fr-FR') : '—';
    const bc=d.statut==='TERM'?'badge-term':d.statut==='TERM ANO'?'badge-ano':
             d.statut==='ACTI'?'badge-acti':d.statut==='EPR'?'badge-epr':'badge-prg';
    const soLabel = d.statutOrdre.includes('commercialement') ? 'Clôturé commercialement'
                  : d.statutOrdre.includes('techniquement')   ? 'Clôturé techniquement'
                  : 'Lancé';
    const soClass = d.statutOrdre.includes('commercialement') ? 'so-clocom'
                  : d.statutOrdre.includes('techniquement')   ? 'so-clotec' : 'so-lance';
    const natureCls = d.nature === 'Préventif'   ? 'badge-nature-prev'
                    : d.nature === 'Correctif'    ? 'badge-nature-cor'
                    : d.nature === 'Exploitation' ? 'badge-nature-exp'
                    : 'badge-nature-autre';
    tr.innerHTML=`
      <td class="mono num-sap">${esc(d.numSAP)}</td>
      <td class="mono ordre-libelle">${esc(d.ordre)}</td>
      <td><strong>${esc(d.ouvrage)}</strong></td>
      <td>${esc(d.typeTravail)}</td>
      <td><span class="badge-nature ${natureCls}">${esc(d.nature)}</span></td>
      <td><span class="badge-statut ${bc}">${esc(d.statut)}</span></td>
      <td><span class="badge-so ${soClass}">${soLabel}</span></td>
      <td>${esc(d.entite)}</td>
      <td>${esc(d.ville)}</td>
      <td class="${d.enRetard?'date-retard':''}">${dateStr}</td>`;
    tbody.appendChild(tr);
  });
}

// ══════════════════════════════════════════════════════════════
//  VUE CORRECTIF
// ══════════════════════════════════════════════════════════════
function renderCorrectif() {
  const cor = FILTERED_COR;
  const now = new Date();

  const total    = cor.length;
  const clotures = cor.filter(d => d.termine || !d.statutOrdre.includes('Lancé')).length;
  const ouverts  = total - clotures;
  const taux     = total > 0 ? Math.round(clotures / total * 100) : 0;
  const anomalies= cor.filter(d => d.statut === 'TERM ANO').length;

  // Ancienneté : OT ouverts dont dateRef est dans le passé
  function moisDepuis(d) {
    if (!d.dateRef) return 0;
    return (now - d.dateRef) / (1000 * 60 * 60 * 24 * 30.44);
  }
  const ouvertsList = cor.filter(d => !d.termine && d.statutOrdre.includes('Lancé') && d.dateRef && d.dateRef < now);
  const depasses    = ouvertsList.length;
  const t1 = ouvertsList.filter(d => { const m=moisDepuis(d); return m>=1 && m<3; }).length;
  const t2 = ouvertsList.filter(d => { const m=moisDepuis(d); return m>=6 && m<12; }).length;
  const t3 = ouvertsList.filter(d => moisDepuis(d) >= 12).length;

  // KPIs
  const set = (id, v) => { const el=document.getElementById(id); if(el) el.textContent=v; };
  set('cor-kpi-total',     total.toLocaleString('fr-FR'));
  set('cor-kpi-ouverts',   ouverts.toLocaleString('fr-FR'));
  set('cor-kpi-clotures',  clotures.toLocaleString('fr-FR'));
  set('cor-kpi-taux',      taux + ' %');
  set('cor-kpi-depasses',  depasses.toLocaleString('fr-FR'));
  set('cor-kpi-anomalies', anomalies.toLocaleString('fr-FR'));

  // Tranches ancienneté
  const pct = n => ouvertsList.length > 0 ? Math.round(n/ouvertsList.length*100)+'%' : '—';
  set('cor-t1-val', t1); set('cor-t1-pct', pct(t1));
  set('cor-t2-val', t2); set('cor-t2-pct', pct(t2));
  set('cor-t3-val', t3); set('cor-t3-pct', pct(t3));

  // Graphique ancienneté (barres horizontales)
  destroyChart('corAnciennete');
  const ctxA = document.getElementById('chartCorAnciennete')?.getContext('2d');
  if (ctxA) {
    CHARTS['corAnciennete'] = new Chart(ctxA, {
      type: 'bar',
      data: {
        labels: ['1–3 mois', '6–12 mois', '>12 mois'],
        datasets: [{ data:[t1,t2,t3], backgroundColor:['#fde047','#f4a35a',C.red], borderRadius:4 }]
      },
      options: {
        indexAxis:'y', responsive:true, maintainAspectRatio:false,
        plugins:{ legend:{display:false} },
        scales:{
          x:{ grid:{color:'#f1f5f9'}, ticks:{font:{family:'DM Sans',size:10}} },
          y:{ grid:{display:false},   ticks:{font:{family:'DM Sans',size:11}} }
        }
      }
    });
  }

  // Graphique par ouvrage (barres : ouverts vs clôturés)
  destroyChart('corOuvrage');
  const ouvrages = [...new Set(cor.map(d=>d.ouvrage))].sort();
  const ctxO = document.getElementById('chartCorOuvrage')?.getContext('2d');
  if (ctxO) {
    CHARTS['corOuvrage'] = new Chart(ctxO, {
      type:'bar',
      data:{
        labels: ouvrages,
        datasets:[
          { label:'Ouverts',  data: ouvrages.map(o=>cor.filter(d=>d.ouvrage===o&&!d.termine&&d.statutOrdre.includes('Lancé')).length), backgroundColor:C.red,   borderRadius:3, stack:'s' },
          { label:'Clôturés', data: ouvrages.map(o=>cor.filter(d=>d.ouvrage===o&&(d.termine||!d.statutOrdre.includes('Lancé'))).length), backgroundColor:C.green, borderRadius:3, stack:'s' },
        ]
      },
      options:{
        indexAxis:'y', responsive:true, maintainAspectRatio:false,
        plugins:{ legend:{position:'bottom',labels:{font:{family:'DM Sans',size:11},boxWidth:12,padding:10}} },
        scales:{
          x:{ stacked:true, grid:{color:'#f1f5f9'}, ticks:{font:{family:'DM Sans',size:10}} },
          y:{ stacked:true, grid:{display:false},   ticks:{font:{family:'DM Sans',size:10}} }
        }
      }
    });
  }

  // Graphique par entité (donut)
  destroyChart('corEntite');
  const entMap = countBy(cor, 'entite');
  const ctxE = document.getElementById('chartCorEntite')?.getContext('2d');
  if (ctxE) {
    CHARTS['corEntite'] = new Chart(ctxE, {
      type:'doughnut',
      data:{ labels:Object.keys(entMap), datasets:[{ data:Object.values(entMap), backgroundColor:PALETTE, borderWidth:2, borderColor:'#fff' }] },
      options:{ responsive:false, maintainAspectRatio:true, cutout:'62%',
        plugins:{ legend:{ position:'bottom', labels:{ font:{family:'DM Sans',size:11}, boxWidth:12, padding:10 } } }
      }
    });
  }

  // Courbe cumulée : entrées vs clôtures par mois 2026
  destroyChart('corCourbe');
  const ctxC = document.getElementById('chartCorCourbe')?.getContext('2d');
  if (ctxC) {
    const moisLabels = ['Jan','Fév','Mar','Avr','Mai','Jun','Jul','Aoû','Sep','Oct','Nov','Déc'];
    let cumEntree=0, cumCloture=0;
    const datEntree=[], datCloture=[];
    for (let m=0; m<12; m++) {
      const entrees  = cor.filter(d => d.dateRef && d.dateRef.getFullYear()===2026 && d.dateRef.getMonth()===m).length;
      const clot     = cor.filter(d => d.termine && d.dateRef && d.dateRef.getFullYear()===2026 && d.dateRef.getMonth()===m).length;
      cumEntree  += entrees;
      cumCloture += clot;
      const isFuture = m > now.getMonth();
      datEntree.push(isFuture ? null : cumEntree);
      datCloture.push(isFuture ? null : cumCloture);
    }
    CHARTS['corCourbe'] = new Chart(ctxC, {
      type:'line',
      data:{ labels:moisLabels, datasets:[
        { label:'OT entrants (cumulé)', data:datEntree,  borderColor:C.red,   backgroundColor:'rgba(185,28,28,.07)', borderWidth:2, tension:.3, fill:true, pointRadius:4, spanGaps:false },
        { label:'OT clôturés (cumulé)', data:datCloture, borderColor:C.green, backgroundColor:'rgba(27,122,62,.07)', borderWidth:2, tension:.3, fill:true, pointRadius:4, spanGaps:false },
      ]},
      options:{
        responsive:true, maintainAspectRatio:false,
        interaction:{ mode:'index', intersect:false },
        plugins:{ legend:{ position:'bottom', labels:{ font:{family:'DM Sans',size:11}, boxWidth:14, padding:14 } } },
        scales:{
          x:{ grid:{color:'#f1f5f9'}, ticks:{font:{family:'DM Sans',size:11}} },
          y:{ grid:{color:'#f1f5f9'}, ticks:{font:{family:'DM Sans',size:10}},
              title:{display:true, text:'OT cumulés', font:{family:'DM Sans',size:10}, color:'#94a3b8'} }
        }
      }
    });
  }

  // Table
  renderCorTable();
}

// ── Mode table correctif ──────────────────────────────────────
function setCorTableMode(mode) {
  COR_TABLE_MODE = mode;
  document.querySelectorAll('#view-correctif .tab-btn').forEach(b =>
    b.classList.toggle('active', b.dataset.mode === mode)
  );
  renderCorTable();
}

function renderCorTable() {
  const now = new Date();
  function moisDepuis(d) {
    if (!d.dateRef) return 0;
    return (now - d.dateRef) / (1000 * 60 * 60 * 24 * 30.44);
  }
  function ancienneteLabel(d) {
    if (d.termine || !d.statutOrdre.includes('Lancé')) return null; // clôturé
    if (!d.dateRef || d.dateRef >= now) return null; // pas encore dépassé
    const m = moisDepuis(d);
    if (m >= 12) return { txt:'>12 mois', cls:'badge-anc-danger' };
    if (m >= 6)  return { txt:'6–12 mois', cls:'badge-anc-warn' };
    if (m >= 1)  return { txt:'1–3 mois', cls:'badge-anc-low' };
    return { txt:'< 1 mois', cls:'badge-anc-ok' };
  }

  const labels = {
    'cor-ouverts':  '🔴 OT correctifs ouverts',
    'cor-depasses': '⏰ OT correctifs dépassés',
    'cor-clotures': '✅ OT correctifs clôturés',
    'cor-all':      '📋 Tous les OT correctifs',
  };

  let rows;
  switch(COR_TABLE_MODE) {
    case 'cor-ouverts':  rows = FILTERED_COR.filter(d => !d.termine && d.statutOrdre.includes('Lancé')); break;
    case 'cor-depasses': rows = FILTERED_COR.filter(d => !d.termine && d.statutOrdre.includes('Lancé') && d.dateRef && d.dateRef < now); break;
    case 'cor-clotures': rows = FILTERED_COR.filter(d => d.termine || !d.statutOrdre.includes('Lancé')); break;
    default:             rows = [...FILTERED_COR];
  }
  rows.sort((a,b) => (a.dateRef||0) - (b.dateRef||0));
  // Enrichir avec _ancDays pour le tri
  const now2 = new Date();
  rows.forEach(d => { d._ancDays = d.dateRef ? Math.floor((now2-d.dateRef)/86400000) : 0; });

  const titleEl = document.getElementById('cor-table-title');
  if (titleEl) titleEl.textContent = `${labels[COR_TABLE_MODE]} — ${rows.length} OT`;

  ROWS_COR = rows;
  SORT_STATE.cor = { col: null, dir: 1 };
  document.querySelectorAll('#table-cor thead th').forEach(h => h.classList.remove('sort-asc','sort-desc'));
  const q = document.getElementById('search-cor')?.value || '';
  _renderCorRows(ROWS_COR, q);
}

function exportExcelCor() {
  const now = new Date();
  function moisDepuis(d) { return d.dateRef ? (now-d.dateRef)/(1000*60*60*24*30.44) : 0; }
  function ancTxt(d) {
    if (d.termine || !d.statutOrdre.includes('Lancé')) return 'Clôturé';
    if (!d.dateRef || d.dateRef >= now) return 'Non dépassé';
    const m = moisDepuis(d);
    if (m>=12) return '>12 mois'; if (m>=6) return '6-12 mois'; if (m>=1) return '1-3 mois';
    return '<1 mois';
  }
  let rows;
  switch(COR_TABLE_MODE) {
    case 'cor-ouverts':  rows=FILTERED_COR.filter(d=>!d.termine&&d.statutOrdre.includes('Lancé')); break;
    case 'cor-depasses': rows=FILTERED_COR.filter(d=>!d.termine&&d.statutOrdre.includes('Lancé')&&d.dateRef&&d.dateRef<now); break;
    case 'cor-clotures': rows=FILTERED_COR.filter(d=>d.termine||!d.statutOrdre.includes('Lancé')); break;
    default:             rows=[...FILTERED_COR];
  }
  const data = rows.map(d => ({
    'N° SAP':        d.numSAP,
    'Libellé ordre': d.ordre,
    'Type d\'ouvrage': d.ouvrage,
    'Type de travail': d.typeTravail,
    'Statut terrain':  d.statut,
    'Statut ordre':    d.statutOrdre,
    'Entité':          d.entite,
    'Ville':           d.ville,
    'Date référence':  d.dateRef ? d.dateRef.toLocaleDateString('fr-FR') : '',
    'Ancienneté':      ancTxt(d),
  }));
  const ws = XLSX.utils.json_to_sheet(data);
  ws['!autofilter'] = { ref: ws['!ref'] };
  ws['!cols'] = [{wch:14},{wch:45},{wch:22},{wch:20},{wch:14},{wch:24},{wch:14},{wch:22},{wch:14},{wch:12}];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Correctif');
  XLSX.writeFile(wb, `GRDF_Correctif_${new Date().toISOString().slice(0,10)}.xlsx`);
}

// ── Helpers ───────────────────────────────────────────────────
function destroyChart(id) { if(CHARTS[id]){CHARTS[id].destroy();delete CHARTS[id];} }
function countBy(data,key) { const m={}; data.forEach(d=>{const v=d[key]||'N/A';m[v]=(m[v]||0)+1;}); return m; }
function esc(str) { return String(str||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
function getFilterValue(id, fallback='') {
  const el = document.getElementById(id);
  return el ? (el.value || fallback) : fallback;
}

const _fd = document.getElementById('footer-date');
if (_fd) _fd.textContent = 'Tableau de bord généré le ' + new Date().toLocaleDateString('fr-FR',{day:'2-digit',month:'long',year:'numeric'});

// ── Export Excel ──────────────────────────────────────────────
function exportExcel() {
  const labels = {
    retard:'OT_en_retard', afaire:'Clotures_SAP_non_remontes',
    attente:'Attente_referent', incohere:'Incoherences',
    termine:'OT_termines', all:'Tous_OT'
  };

  let rows;
  switch(TABLE_MODE) {
    case 'retard':   rows=FILTERED.filter(d=>d.enRetard); break;
    case 'afaire':   rows=FILTERED.filter(d=>!d.termine && d.statutOrdre.includes('Lancé')); break;
    case 'attente':  rows=FILTERED.filter(d=>d.attenteReferent); break;
    case 'incohere': rows=FILTERED.filter(d=>d.incoherence); break;
    case 'termine':  rows=FILTERED.filter(d=>d.termine); break;
    default:         rows=[...FILTERED];
  }

  const data = rows.map(d => ({
    'N° SAP':               d.numSAP,
    'Libellé ordre':        d.ordre,
    'Nature':               d.nature,
    'Type d\'ouvrage':      d.ouvrage,
    'Type de travail':      d.typeTravail,
    'Statut terrain':       d.statut,
    'Statut ordre':         d.statutOrdre.replace(' (2)','').replace(' (3)','').replace(' (6)',''),
    'Entité':               d.entite,
    'Ville':                d.ville,
    'Date référence':       d.dateRef ? d.dateRef.toLocaleDateString('fr-FR') : '',
    'Gamme':                d.gamme,
    'En retard':            d.enRetard ? 'Oui' : 'Non',
  }));

  const ws = XLSX.utils.json_to_sheet(data);
  const range = XLSX.utils.decode_range(ws['!ref']);
  for (let C2 = range.s.c; C2 <= range.e.c; C2++) {
    const cell = ws[XLSX.utils.encode_cell({r:0, c:C2})];
    if (cell) cell.s = { font:{bold:true}, fill:{fgColor:{rgb:'003189'}}, alignment:{horizontal:'center'} };
  }
  ws['!autofilter'] = { ref: ws['!ref'] };
  ws['!cols'] = [
    {wch:14},{wch:45},{wch:14},{wch:22},{wch:20},{wch:14},{wch:24},
    {wch:14},{wch:22},{wch:14},{wch:30},{wch:10}
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'OT');

  const meta = XLSX.utils.aoa_to_sheet([
    ['Export GRDF — Agence AI Boucles de Seine Nord'],
    ['Date d\'export', new Date().toLocaleDateString('fr-FR')],
    ['Filtre entité', getFilterValue('filter-entite', 'Toutes')],
    ['Filtre ouvrage', getFilterValue('filter-ouvrage', 'Tous')],
    ['Filtre travail', getFilterValue('filter-travail', 'Tous')],
    ['Nombre d\'OT exportés', rows.length],
  ]);
  XLSX.utils.book_append_sheet(wb, meta, 'Informations');

  const date = new Date().toISOString().slice(0,10);
  XLSX.writeFile(wb, `GRDF_${labels[TABLE_MODE]}_${date}.xlsx`);
}

// ── Export PDF ────────────────────────────────────────────────
async function exportPDF() {
  if (!ALL_DATA.length) { alert('Aucune donnée chargée.'); return; }
  const btn = document.querySelector('.btn-pdf');
  btn.innerHTML = '⏳ Génération...'; btn.disabled = true;

  try {
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({ orientation:'landscape', unit:'mm', format:'a3' });
    const W=420, H=297, M=14;

    const hex2rgb = h => [parseInt(h.slice(1,3),16),parseInt(h.slice(3,5),16),parseInt(h.slice(5,7),16)];
    const setTxt = (color,size,style='normal') => {
      pdf.setTextColor(...hex2rgb(color)); pdf.setFontSize(size); pdf.setFont('helvetica',style);
    };

    pdf.setFillColor(0,49,137); pdf.rect(0,0,W,20,'F');

    try {
      const logoEl = document.querySelector('img.logo-img');
      const c = document.createElement('canvas'); c.width=221; c.height=221;
      c.getContext('2d').drawImage(logoEl,0,0,221,221);
      pdf.addImage(c.toDataURL('image/png'),'PNG',M,2.5,15,15);
    } catch(e){}

    setTxt('#ffffff',12,'bold');
    pdf.text('Agence AI Boucles de Seine Nord — Tableau de bord maintenance', M+18, 11);
    const ds = new Date().toLocaleDateString('fr-FR',{day:'2-digit',month:'long',year:'numeric'});
    setTxt('#b4c8ff',7.5,'normal');
    pdf.text('Généré le '+ds, W-M, 7, {align:'right'});
    const fe=getFilterValue('filter-entite', 'Toutes entités');
    const fo=getFilterValue('filter-ouvrage', 'Tous ouvrages');
    const ft=getFilterValue('filter-travail', 'Tous types');
    pdf.text(`Filtres : ${fe} · ${fo} · ${ft} · ${FILTERED.length} OT`, W-M, 15, {align:'right'});

    const total   = FILTERED.length;
    const term    = FILTERED.filter(d=>d.statut==='TERM').length;
    const termAno = FILTERED.filter(d=>d.statut==='TERM ANO').length;
    const retard  = FILTERED.filter(d=>d.enRetard).length;
    const encours = FILTERED.filter(d=>['ACTI','EPR','PRG','OGDI'].includes(d.statut)).length;
    const attente = FILTERED.filter(d=>d.attenteReferent).length;
    const taux    = total>0 ? ((term+termAno)/total*100).toFixed(1)+' %' : '—';

    const kpis = [
      {label:'Total OT',                    value:total.toLocaleString('de-DE'), accent:'#003189'},
      {label:'Taux de clôture',             value:taux,                          accent:'#1b7a3e'},
      {label:'OT en retard',                value:String(retard),                accent:'#b91c1c'},
      {label:'Terminés avec anomalie',      value:String(termAno),               accent:'#c95f00'},
      {label:'En cours',                    value:String(encours),               accent:'#64748b'},
      {label:'Attente référent',            value:String(attente),               accent:'#6d28d9'},
    ];
    const kpiY=23, kpiH=16, kpiW=(W-M*2-10)/6;
    kpis.forEach((k,i) => {
      const x=M+i*(kpiW+2);
      pdf.setFillColor(255,255,255); pdf.setDrawColor(220,228,238); pdf.setLineWidth(0.2);
      pdf.roundedRect(x,kpiY,kpiW,kpiH,1.5,1.5,'FD');
      pdf.setFillColor(...hex2rgb(k.accent));
      pdf.roundedRect(x,kpiY,1.8,kpiH,1,1,'F');
      pdf.rect(x+0.9,kpiY,0.9,kpiH,'F');
      setTxt(k.accent,13,'bold');
      pdf.text(k.value, x+kpiW/2, kpiY+8.5, {align:'center'});
      setTxt('#64748b',5.5,'normal');
      pdf.text(k.label.toUpperCase(), x+kpiW/2, kpiY+14, {align:'center'});
    });

    const gY = kpiY+kpiH+4;
    const gaugesEl = document.getElementById('section-gauges');
    const gCanvas = await html2canvas(gaugesEl, {scale:2.5, backgroundColor:'#ffffff', logging:false});
    const gImg = gCanvas.toDataURL('image/png');
    const gH = 78;
    pdf.setFillColor(255,255,255); pdf.setDrawColor(220,228,238); pdf.setLineWidth(0.25);
    pdf.roundedRect(M,gY,W-M*2,gH,2,2,'FD');
    setTxt('#1e293b',7,'bold');
    pdf.text('AVANCEMENT PAR TYPE D\'OUVRAGE', M+4, gY+5.5);
    [[`≥90%`,'#1b7a3e'],[`60–90%`,'#c95f00'],[`<60%`,'#b91c1c']].forEach((l,i)=>{
      pdf.setFillColor(...hex2rgb(l[1])); pdf.circle(W-M-58+i*22,gY+5,1.2,'F');
      setTxt('#64748b',6,'normal'); pdf.text(l[0],W-M-55+i*22,gY+5.8);
    });
    pdf.addImage(gImg,'PNG',M+2,gY+7,W-M*2-4,gH-10);

    const cY = gY+gH+4;
    const cH = H-cY-10;
    const chartsEl = document.getElementById('section-charts');
    const prevStyle = chartsEl.getAttribute('style')||'';
    chartsEl.style.cssText = `height:${Math.round(cH/0.264583)}px; overflow:hidden;`;
    await new Promise(r=>setTimeout(r,100));
    const cCanvas = await html2canvas(chartsEl, {scale:2, backgroundColor:'#f6f8fd', logging:false});
    chartsEl.setAttribute('style', prevStyle);
    const cImg = cCanvas.toDataURL('image/png');
    pdf.addImage(cImg,'PNG',M,cY,W-M*2,cH);

    pdf.setFillColor(0,49,137); pdf.rect(0,H-8,W,8,'F');
    setTxt('#ffffff',7,'normal');
    pdf.text('GRDF — Agence AI Boucles de Seine Nord — Confidentiel interne',W/2,H-3.5,{align:'center'});
    setTxt('#b4c8ff',7,'normal');
    pdf.text('1 / 1',W-M,H-3.5,{align:'right'});

    const dateFile=new Date().toISOString().slice(0,10);
    pdf.save(`GRDF_Rapport_Maintenance_${dateFile}.pdf`);

  } catch(err) {
    console.error('PDF error:', err);
    alert('Erreur lors de la génération PDF : '+err.message);
  } finally {
    btn.innerHTML='<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="12" y1="12" x2="12" y2="18"/><line x1="9" y1="15" x2="15" y2="15"/></svg> Exporter PDF';
    btn.disabled=false;
  }
}

// ── Export PPTX ───────────────────────────────────────────────
async function exportPPTX() {
  if (!ALL_DATA.length) { alert('Aucune donnée chargée.'); return; }
  const btn = document.querySelector('.btn-pptx');
  btn.innerHTML = '⏳ Génération...'; btn.disabled = true;

  try {
    const pres = new PptxGenJS();
    pres.layout = 'LAYOUT_WIDE';
    pres.title = 'GRDF - Tableau de bord maintenance';

    const W=13.3, H=7.5;
    const BLUE='003189', GREEN='1b7a3e', ORANGE='c95f00', RED='b91c1c',
          GRAY='64748b', PURPLE='6d28d9', WHITE='ffffff', DARK='1e293b', LIGHT='f6f8fd';
    const mk = () => ({ type:'outer', blur:8, offset:2, angle:135, color:'000000', opacity:0.08 });

    let logoData = null;
    try {
      const logoEl = document.querySelector('img.logo-img');
      const c = document.createElement('canvas'); c.width=221; c.height=221;
      c.getContext('2d').drawImage(logoEl,0,0,221,221);
      logoData = c.toDataURL('image/png');
    } catch(e){}

    function footer(slide) {
      slide.addShape(pres.ShapeType.rect, {x:0,y:7.1,w:W,h:0.4,fill:{color:BLUE},line:{color:BLUE}});
      slide.addText('GRDF — Agence AI Boucles de Seine Nord — Confidentiel interne',
        {x:0,y:7.1,w:W,h:0.4,fontSize:8,color:'93b4e8',align:'center',valign:'middle',fontFace:'Calibri'});
    }

    const total   = FILTERED.length;
    const term    = FILTERED.filter(d=>d.statut==='TERM').length;
    const termAno = FILTERED.filter(d=>d.statut==='TERM ANO').length;
    const retard  = FILTERED.filter(d=>d.enRetard).length;
    const encours = FILTERED.filter(d=>['ACTI','EPR','PRG','OGDI'].includes(d.statut)).length;
    const attente = FILTERED.filter(d=>d.attenteReferent).length;
    const taux    = total>0 ? ((term+termAno)/total*100).toFixed(1) : '0';
    const mois    = ['janvier','février','mars','avril','mai','juin','juillet','août','septembre','octobre','novembre','décembre'];
    const now     = new Date();
    const dateStr = `${String(now.getDate()).padStart(2,'0')} ${mois[now.getMonth()]} ${now.getFullYear()}`;

    // ── SLIDE 1 : Titre ──────────────────────────────────────
    const s1 = pres.addSlide();
    s1.background = { color: BLUE };
    if (logoData) s1.addImage({data:logoData, x:0.5, y:1.8, w:1.6, h:1.6});
    s1.addText('Tableau de bord maintenance',
      {x:2.5,y:1.9,w:10,h:1.1,fontSize:40,bold:true,color:WHITE,fontFace:'Calibri'});
    s1.addText('Agence AI Boucles de Seine Nord',
      {x:2.5,y:3.05,w:10,h:0.65,fontSize:22,color:'93b4e8',fontFace:'Calibri'});
    s1.addText(`Données au ${dateStr}  ·  ${total.toLocaleString('de-DE')} OT`,
      {x:2.5,y:3.75,w:10,h:0.45,fontSize:14,color:'6b8fd4',fontFace:'Calibri'});
    s1.addShape(pres.ShapeType.rect,{x:0,y:6.9,w:W,h:0.6,fill:{color:'002070'},line:{color:'002070'}});
    s1.addText('Confidentiel interne — GRDF',
      {x:0,y:6.9,w:W,h:0.6,fontSize:10,color:'6b8fd4',align:'center',valign:'middle',fontFace:'Calibri'});

    // ── SLIDE 2 : KPIs ───────────────────────────────────────
    const s2 = pres.addSlide();
    s2.background = { color: LIGHT };
    s2.addShape(pres.ShapeType.rect,{x:0,y:0,w:W,h:0.65,fill:{color:BLUE},line:{color:BLUE}});
    s2.addText('INDICATEURS CLÉS',{x:0.4,y:0,w:8,h:0.65,fontSize:14,bold:true,color:WHITE,valign:'middle',fontFace:'Calibri'});
    s2.addText(`${dateStr}  ·  ${total} OT`,{x:8,y:0,w:5,h:0.65,fontSize:11,color:'93b4e8',align:'right',valign:'middle',fontFace:'Calibri',margin:10});

    const kpis = [
      {label:'Total OT',               value:total.toLocaleString('de-DE'), accent:BLUE},
      {label:'Taux de clôture',        value:taux+'%',                      accent:GREEN},
      {label:'OT en retard',           value:String(retard),                accent:RED},
      {label:'Terminés avec anomalie', value:String(termAno),               accent:ORANGE},
      {label:'En cours',               value:String(encours),               accent:GRAY},
      {label:'Attente référent',       value:String(attente),               accent:PURPLE},
    ];
    const kW=1.92, kH=1.55, kY=0.85, kGap=0.16;
    kpis.forEach((k,i) => {
      const x=0.35+i*(kW+kGap);
      s2.addShape(pres.ShapeType.rect,{x,y:kY,w:kW,h:kH,fill:{color:WHITE},line:{color:'e2e8f0',pt:0.5},shadow:mk()});
      s2.addShape(pres.ShapeType.rect,{x,y:kY,w:0.07,h:kH,fill:{color:k.accent},line:{color:k.accent}});
      s2.addText(k.value,{x:x+0.15,y:kY+0.18,w:kW-0.2,h:0.75,fontSize:30,bold:true,color:k.accent,fontFace:'Calibri'});
      s2.addText(k.label.toUpperCase(),{x:x+0.15,y:kY+1.0,w:kW-0.2,h:0.42,fontSize:8.5,color:'94a3b8',fontFace:'Calibri'});
    });

    const statMap = {};
    FILTERED.forEach(d => { statMap[d.statutSimple] = (statMap[d.statutSimple]||0)+1; });
    const statEntries = Object.entries(statMap).sort((a,b)=>b[1]-a[1]);
    const tY=2.6;
    s2.addShape(pres.ShapeType.rect,{x:0.35,y:tY,w:12.6,h:0.42,fill:{color:BLUE},line:{color:BLUE}});
    ['Statut terrain','Nb OT','% du total'].forEach((h,i)=>{
      const xs=[0.45,5.5,9.0]; const ws=[5,3,3];
      s2.addText(h,{x:xs[i],y:tY,w:ws[i],h:0.42,fontSize:10,bold:true,color:WHITE,valign:'middle',fontFace:'Calibri'});
    });
    statEntries.forEach(([label,nb],i)=>{
      const ry=tY+0.42+i*0.4;
      s2.addShape(pres.ShapeType.rect,{x:0.35,y:ry,w:12.6,h:0.4,fill:{color:i%2===0?WHITE:'f8fafc'},line:{color:'e2e8f0',pt:0.3}});
      s2.addText(label,   {x:0.45,y:ry,w:5,  h:0.4,fontSize:10,color:DARK,valign:'middle',fontFace:'Calibri'});
      s2.addText(String(nb),{x:5.5, y:ry,w:3,  h:0.4,fontSize:10,bold:true,color:DARK,valign:'middle',fontFace:'Calibri'});
      s2.addText(Math.round(nb/total*100)+'%',{x:9.0,y:ry,w:3,h:0.4,fontSize:10,color:GRAY,valign:'middle',fontFace:'Calibri'});
    });
    footer(s2);

    // ── SLIDE 3 : Avancement par ouvrage ─────────────────────
    const s3 = pres.addSlide();
    s3.background = { color: LIGHT };
    s3.addShape(pres.ShapeType.rect,{x:0,y:0,w:W,h:0.65,fill:{color:BLUE},line:{color:BLUE}});
    s3.addText("AVANCEMENT PAR TYPE D'OUVRAGE",{x:0.4,y:0,w:9,h:0.65,fontSize:14,bold:true,color:WHITE,valign:'middle',fontFace:'Calibri'});
    s3.addText(dateStr,{x:9,y:0,w:4,h:0.65,fontSize:11,color:'93b4e8',align:'right',valign:'middle',fontFace:'Calibri',margin:10});

    [[`≥90%`,GREEN],[`60–90%`,ORANGE],[`<60%`,RED]].forEach(([l,c],i)=>{
      s3.addShape(pres.ShapeType.ellipse,{x:7.4+i*1.9,y:0.77,w:0.15,h:0.15,fill:{color:c},line:{color:c}});
      s3.addText(l,{x:7.58+i*1.9,y:0.73,w:1.7,h:0.22,fontSize:8,color:GRAY,fontFace:'Calibri'});
    });

    const ouvrages = [...new Set(FILTERED.map(d=>d.ouvrage))].sort();
    const nCols = Math.min(ouvrages.length, 6);
    const gW=(W-0.8)/nCols-0.12, gH=ouvrages.length>6?2.75:3.0, gStartY=1.0;

    ouvrages.forEach((ouvrage,idx)=>{
      const col=idx%nCols, row=Math.floor(idx/nCols);
      const gx=0.4+col*(gW+0.12), gy=gStartY+row*(gH+0.15);
      const sub=FILTERED.filter(d=>d.ouvrage===ouvrage);
      const pct=sub.length>0?Math.round(sub.filter(d=>d.termine).length/sub.length*100):0;
      const fait=sub.filter(d=>d.termine).length;
      const retardN=sub.filter(d=>d.enRetard).length;
      const color=pct>=90?GREEN:pct>=60?ORANGE:RED;

      s3.addShape(pres.ShapeType.rect,{x:gx,y:gy,w:gW,h:gH,fill:{color:WHITE},line:{color:'e2e8f0',pt:0.5},shadow:mk()});
      s3.addText(ouvrage,{x:gx+0.05,y:gy+0.08,w:gW-0.1,h:0.32,fontSize:8.5,bold:true,color:DARK,align:'center',fontFace:'Calibri'});
      const bx=gx+0.12, bw=gW-0.24, by=gy+0.48, bh=0.16;
      s3.addShape(pres.ShapeType.rect,{x:bx,y:by,w:bw,h:bh,fill:{color:'e2e8f0'},line:{color:'e2e8f0'}});
      s3.addShape(pres.ShapeType.rect,{x:bx,y:by,w:Math.max(bw*pct/100,0.04),h:bh,fill:{color:color},line:{color:color}});
      s3.addText(pct+'%',{x:gx,y:gy+0.7,w:gW,h:0.52,fontSize:26,bold:true,color:color,align:'center',fontFace:'Calibri'});
      s3.addText(`${fait} / ${sub.length}`,{x:gx,y:gy+1.26,w:gW,h:0.28,fontSize:8.5,color:GRAY,align:'center',fontFace:'Calibri'});
      const badgeY=gy+gH-0.36;
      if (retardN>0) {
        s3.addShape(pres.ShapeType.rect,{x:gx+0.08,y:badgeY,w:gW-0.16,h:0.28,fill:{color:'fde8e8'},line:{color:'fde8e8'}});
        s3.addText(`⚠ ${retardN} en retard`,{x:gx+0.08,y:badgeY,w:gW-0.16,h:0.28,fontSize:7.5,bold:true,color:RED,align:'center',valign:'middle',fontFace:'Calibri'});
      } else {
        s3.addShape(pres.ShapeType.rect,{x:gx+0.08,y:badgeY,w:gW-0.16,h:0.28,fill:{color:'d6f0e0'},line:{color:'d6f0e0'}});
        s3.addText('✓ À jour',{x:gx+0.08,y:badgeY,w:gW-0.16,h:0.28,fontSize:7.5,color:GREEN,align:'center',valign:'middle',fontFace:'Calibri'});
      }
    });
    footer(s3);

    // ── SLIDE 4 : Fait vs Reste à faire ──────────────────────
    const s4 = pres.addSlide();
    s4.background = { color: LIGHT };
    s4.addShape(pres.ShapeType.rect,{x:0,y:0,w:W,h:0.65,fill:{color:BLUE},line:{color:BLUE}});
    s4.addText("FAIT VS RESTE À FAIRE — PAR TYPE D'OUVRAGE",{x:0.4,y:0,w:11,h:0.65,fontSize:14,bold:true,color:WHITE,valign:'middle',fontFace:'Calibri'});

    const sorted=[...ouvrages].sort((a,b)=>{
      return FILTERED.filter(d=>d.ouvrage===b).length - FILTERED.filter(d=>d.ouvrage===a).length;
    });
    const chartData = [
      { name:'Terminés',      labels:sorted, values:sorted.map(o=>FILTERED.filter(d=>d.ouvrage===o&&d.termine).length) },
      { name:'Reste à faire', labels:sorted, values:sorted.map(o=>FILTERED.filter(d=>d.ouvrage===o&&!d.termine).length) },
    ];
    s4.addChart(pres.charts.BAR, chartData, {
      x:0.4,y:0.8,w:12.5,h:5.4,
      barDir:'bar', barGrouping:'stacked',
      chartColors:[GREEN,'e87777'],
      chartArea:{fill:{color:WHITE},roundedCorners:true},
      catAxisLabelColor:DARK, catAxisLabelFontSize:10,
      valAxisLabelColor:GRAY, valAxisLabelFontSize:9,
      valGridLine:{color:'e2e8f0',size:0.5},
      catGridLine:{style:'none'},
      showLegend:true, legendPos:'b', legendFontSize:10,
    });
    footer(s4);

    // ── SLIDE 5 : Plan d'activité ─────────────────────────────
    const s5 = pres.addSlide();
    s5.background = { color: LIGHT };
    s5.addShape(pres.ShapeType.rect,{x:0,y:0,w:W,h:0.65,fill:{color:BLUE},line:{color:BLUE}});
    s5.addText("SUIVI DU PLAN D'ACTIVITÉ",{x:0.4,y:0,w:9,h:0.65,fontSize:14,bold:true,color:WHITE,valign:'middle',fontFace:'Calibri'});
    s5.addText(dateStr,{x:9,y:0,w:4,h:0.65,fontSize:11,color:'93b4e8',align:'right',valign:'middle',fontFace:'Calibri',margin:10});

    const plan = computePlanActivite();
    const planKpis = [
      {label:'Plan cumulé S'+plan.currentWeek, value:String(plan.theoriqueTotal), accent:BLUE},
      {label:'Réalisé (BRC+ROB+PDL+PDR+SIN)',  value:String(plan.realTermine),    accent:GREEN},
      {label:'% Réalisation',                  value:plan.pctRealisation+'%',     accent:plan.pctRealisation>=90?GREEN:plan.pctRealisation>=70?ORANGE:RED},
      {label:'Écart (réel - plan)',             value:(plan.enAvance?'+':'')+plan.ecart+' OT', accent:plan.enAvance?GREEN:RED},
    ];
    const pkW=2.9, pkH=1.3, pkY=0.8, pkGap=0.2;
    planKpis.forEach((k,i)=>{
      const x=0.35+i*(pkW+pkGap);
      s5.addShape(pres.ShapeType.rect,{x,y:pkY,w:pkW,h:pkH,fill:{color:WHITE},line:{color:'e2e8f0',pt:0.5},shadow:mk()});
      s5.addShape(pres.ShapeType.rect,{x,y:pkY,w:0.07,h:pkH,fill:{color:k.accent},line:{color:k.accent}});
      s5.addText(k.value,{x:x+0.15,y:pkY+0.1,w:pkW-0.2,h:0.7,fontSize:28,bold:true,color:k.accent,fontFace:'Calibri'});
      s5.addText(k.label.toUpperCase(),{x:x+0.15,y:pkY+0.88,w:pkW-0.2,h:0.34,fontSize:7.5,color:'94a3b8',fontFace:'Calibri'});
    });

    // Chart courbe plan vs réel
    const lineLabels = plan.chartLabels;
    const lineData = [
      { name:'Plan théorique (cumulé)', labels:lineLabels, values:plan.chartTheo },
      { name:'Réalisé (cumulé)',        labels:lineLabels, values:plan.chartReel },
    ];
    s5.addChart(pres.charts.LINE, lineData, {
      x:0.4, y:2.25, w:12.5, h:4.5,
      chartColors:[BLUE, GREEN],
      lineDataSymbol:'none',
      chartArea:{fill:{color:WHITE},roundedCorners:true},
      catAxisLabelColor:DARK, catAxisLabelFontSize:9,
      valAxisLabelColor:GRAY, valAxisLabelFontSize:9,
      valGridLine:{color:'e2e8f0',size:0.5},
      catGridLine:{style:'none'},
      showLegend:true, legendPos:'b', legendFontSize:10,
    });
    footer(s5);

    const dateFile = now.toISOString().slice(0,10);
    await pres.writeFile({ fileName:`GRDF_Rapport_Maintenance_${dateFile}.pptx` });

  } catch(err) {
    console.error('PPTX error:', err);
    alert('Erreur : ' + err.message);
  } finally {
    btn.innerHTML = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="2" y="3" width="20" height="14" rx="2"/><path d="M8 21h8M12 17v4"/></svg> PowerPoint';
    btn.disabled = false;
  }
}



// ════════════════════════════════════════════════════════════════
//  MODULE AVIS PM — Chargement, parsing, rendu, export
// ════════════════════════════════════════════════════════════════

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

// ── Helpers normalisation ─────────────────────────────────────
function normalizeStr(s) {
  return String(s || '').toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/\u0153/g, 'oe').replace(/\u00e6/g, 'ae')
    .replace(/['']/g, "'")
    .replace(/\s+/g, ' ').trim();
}

function isRobReseau(ouvrage) {
  // Poste technique format : "D95491-ROB00056" (ROB réseau → MAINT0410)
  // VS "D95582-BRC00025-BR01-ROB1" (robinet BRC → MAINT0910)
  // Règle : contient ROB mais PAS BRC
  const up = String(ouvrage || '').toUpperCase();
  return up.includes('ROB') && !up.includes('BRC');
}

// ── MAINT0410 : règles ROB réseau (schéma de vannage uniquement) ──
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

      // Afficher le dashboard si pas encore visible
      document.getElementById('empty-state').style.display = 'none';
      document.getElementById('dashboard').style.display   = 'block';

      // Switcher automatiquement sur l'onglet avis
      setMode('avis');
    } catch(err) { alert('Erreur lecture avis PM : ' + err.message); console.error(err); }
  };
  reader.readAsArrayBuffer(file);
}

function parseAvisRow(row) {
  const defail  = String(row['Code défaillance (Description)'] || row['Défaillance'] || '');
  // Poste technique = "D95491-ROB00056" ou "D92025-BRC00268" → détermine MAINT0410 vs MAINT0910
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


function renderAvis() {
  if (!AVIS_DATA.length) {
    document.getElementById('avis-empty-banner').style.display = '';
    document.getElementById('avis-content').style.display      = 'none';
    return;
  }
  document.getElementById('avis-empty-banner').style.display = 'none';
  document.getElementById('avis-content').style.display      = '';
  renderAvisKPIs();
  renderAvisCharts();
  renderAvisTable();
}

function renderAvisKPIs() {
  const total      = AVIS_DATA.length;
  const f0         = AVIS_DATA.filter(a => a.isF0).length;
  const f0SansOT   = AVIS_DATA.filter(a => a.isF0 && !a.hasOT).length;
  const sansOT     = AVIS_DATA.filter(a => !a.hasOT).length;
  const ko         = AVIS_DATA.filter(a => a.confCode === 'KO').length;
  const lies       = AVIS_DATA.filter(a => a.otLie).length;
  const nonQualif  = AVIS_DATA.filter(a => a.confCode === '?').length;

  document.getElementById('avis-kpi-total').textContent      = total.toLocaleString('de-DE');
  document.getElementById('avis-kpi-f0').textContent         = f0;
  document.getElementById('avis-kpi-f0-ot').textContent      = f0SansOT;
  document.getElementById('avis-kpi-sans-ot').textContent    = sansOT;
  document.getElementById('avis-kpi-ko').textContent         = ko;
  document.getElementById('avis-kpi-lies').textContent       = lies;
  const nqEl = document.getElementById('avis-kpi-nonqualif');
  if (nqEl) nqEl.textContent = nonQualif;

  // ── Focus BRC (MAINT0910) ──
  const brcAll  = AVIS_DATA.filter(a => a.refLabel === 'MAINT0910');
  const brcP18  = brcAll.filter(a => a.delaiLabel === 'P-18 mois');
  const brcP    = brcAll.filter(a => a.nature === 'P' && a.delaiLabel !== 'P-18 mois' && a.delaiLabel !== 'IMMÉDIAT');
  const brcR    = brcAll.filter(a => a.nature === 'R');

  const set = (id, val, detail) => {
    const el = document.getElementById(id);
    const elD = document.getElementById(id + '-detail');
    if (el)  el.textContent = val;
    if (elD) elD.textContent = detail;
  };

  const pct = (sub, all) => all.length ? Math.round(sub.length/all.length*100) + ' % du total BRC' : '';

  set('kpi-brc-p18',  brcP18.length, pct(brcP18, brcAll));
  set('kpi-brc-p',    brcP.length,   pct(brcP,   brcAll));
  set('kpi-brc-r',    brcR.length,   pct(brcR,   brcAll));

  // ── Focus ROB (MAINT0410) ──
  const robAll  = AVIS_DATA.filter(a => a.refLabel === 'MAINT0410');
  const robNacc = robAll.filter(a => normalizeStr(a.defail||'').includes('acces'));
  const robNman = robAll.filter(a => normalizeStr(a.defail||'').includes('manoeuv'));

  set('kpi-rob-nacc',  robNacc.length, pct(robNacc, robAll));
  set('kpi-rob-nman',  robNman.length, pct(robNman, robAll));
  set('kpi-rob-all',   robAll.length,  robAll.length + ' robinets réseau');
}

function renderAvisCharts() {
  const prioMap  = {};
  AVIS_DATA.forEach(a => prioMap[a.prioGrp] = (prioMap[a.prioGrp]||0)+1);
  const pOrder   = ['F0','P1','P2','R1','R2','R3','Autre'];
  const pLabels  = { F0:'F0 Fuite', P1:'P1 Prescrit', P2:'P2 Prescrit', R1:'R1 Recom.', R2:'R2 Recom.', R3:'R3 Recom.', Autre:'Autre' };
  const pColors  = { F0:'#7f1d1d', P1:'#b91c1c', P2:'#c95f00', R1:'#1a4faf', R2:'#6d28d9', R3:'#0f766e', Autre:'#94a3b8' };
  const keys     = pOrder.filter(k => prioMap[k]);

  if (CHARTS.avisPrio) CHARTS.avisPrio.destroy();
  CHARTS.avisPrio = new Chart(
    document.getElementById('chartAvisPrio').getContext('2d'), {
      type:'doughnut',
      data:{ labels: keys.map(k=>pLabels[k]),
             datasets:[{ data: keys.map(k=>prioMap[k]),
                         backgroundColor: keys.map(k=>pColors[k]), borderWidth:2 }] },
      options:{ responsive:true, plugins:{ legend:{ position:'right', labels:{ font:{size:11} } } } }
    }
  );

  const buckets  = ['0–3 mois','3–6 mois','6–12 mois','1–2 ans','> 2 ans'];
  const bucketFn = d => d<=90?'0–3 mois':d<=180?'3–6 mois':d<=365?'6–12 mois':d<=730?'1–2 ans':'> 2 ans';
  const bMap     = {};
  AVIS_DATA.filter(a=>!a.hasOT).forEach(a => { const b=bucketFn(a.ageDays); bMap[b]=(bMap[b]||0)+1; });

  if (CHARTS.avisAge) CHARTS.avisAge.destroy();
  CHARTS.avisAge = new Chart(
    document.getElementById('chartAvisAge').getContext('2d'), {
      type:'bar',
      data:{ labels: buckets,
             datasets:[{ label:'Avis sans OT', data: buckets.map(b=>bMap[b]||0),
               backgroundColor:['#1b7a3e','#c95f00','#c95f00','#b91c1c','#7f1d1d'], borderRadius:4 }] },
      options:{ responsive:true, maintainAspectRatio:false,
        plugins:{ legend:{ display:false } },
        scales:{ y:{ beginAtZero:true, grid:{color:'#e2e8f0'} }, x:{ grid:{display:false} } } }
    }
  );
}

function setAvisMode(mode) {
  AVIS_MODE = mode;
  document.querySelectorAll('[data-avis-mode]').forEach(b =>
    b.classList.toggle('active', b.dataset.avisMode === mode)
  );
  renderAvisTable();
}

function renderAvisTable() {
  const titles = {
    'f0':        '🚨 Fuites F0',
    'sans-ot':   '⚠️ Avis sans OT correctif',
    'ko':        '⛔ Hors délai MAINT0910',
    'nonqualif': '? NON QUALIFIÉES',
    'brc-p18':   '🗂 BRC — P-18 MOIS (Accès impossibles)',
    'brc-p':     '🗂 BRC — PRESCRIT (P)',
    'brc-r':     '🗂 BRC — RECOMMANDÉ (R)',
    'rob-nacc':  '🔧 ROB — NON ACCESSIBLE (NACC)',
    'rob-nman':  '🔧 ROB — NON MANŒUVRABLE (NMAN)',
    'lies':      '🔗 Avis liés à un OT chargé',
    'all':       '📋 Tous les avis',
  };
  let rows;
  switch(AVIS_MODE) {
    case 'f0':        rows = AVIS_DATA.filter(a => a.isF0);           break;
    case 'sans-ot':   rows = AVIS_DATA.filter(a => !a.hasOT);         break;
    case 'ko':        rows = AVIS_DATA.filter(a => a.confCode==='KO'); break;
    case 'nonqualif': rows = AVIS_DATA.filter(a => a.confCode==='?');  break;
    case 'brc-p18':   rows = AVIS_DATA.filter(a => a.refLabel==='MAINT0910' && a.delaiLabel==='P-18 mois'); break;
    case 'brc-p':     rows = AVIS_DATA.filter(a => a.refLabel==='MAINT0910' && a.nature==='P' && a.delaiLabel!=='P-18 mois' && a.delaiLabel!=='IMMÉDIAT'); break;
    case 'brc-r':     rows = AVIS_DATA.filter(a => a.refLabel==='MAINT0910' && a.nature==='R'); break;
    case 'rob-nacc':  rows = AVIS_DATA.filter(a => a.refLabel==='MAINT0410' && normalizeStr(a.defail||'').includes('acces')); break;
    case 'rob-nman':  rows = AVIS_DATA.filter(a => a.refLabel==='MAINT0410' && normalizeStr(a.defail||'').includes('manoeuv')); break;
    case 'lies':      rows = AVIS_DATA.filter(a => a.otLie);           break;
    default:          rows = [...AVIS_DATA];
  }
  rows = [...rows].sort((a,b) =>
    (PRIO_ORD_AVIS[a.prioGrp]??9) - (PRIO_ORD_AVIS[b.prioGrp]??9) || b.ageDays - a.ageDays
  );
  document.getElementById('avis-table-title').textContent =
    (titles[AVIS_MODE] || 'Avis') + ' — ' + rows.length + ' avis';

  ROWS_AVIS = rows;
  SORT_STATE.avis = { col: null, dir: 1 };
  document.querySelectorAll('#table-avis thead th').forEach(h => h.classList.remove('sort-asc','sort-desc'));
  const q = document.getElementById('search-avis')?.value || '';
  _renderAvisRows(ROWS_AVIS, q);
}

function exportExcelAvis() {
  if (!AVIS_DATA.length) return;

  // Reproduit le même filtre que renderAvisTable
  let filtered;
  switch(AVIS_MODE) {
    case 'f0':        filtered = AVIS_DATA.filter(a => a.isF0);           break;
    case 'sans-ot':   filtered = AVIS_DATA.filter(a => !a.hasOT);         break;
    case 'ko':        filtered = AVIS_DATA.filter(a => a.confCode==='KO'); break;
    case 'nonqualif': filtered = AVIS_DATA.filter(a => a.confCode==='?');  break;
    case 'brc-p18':   filtered = AVIS_DATA.filter(a => a.refLabel==='MAINT0910' && a.delaiLabel==='P-18 mois'); break;
    case 'brc-p':     filtered = AVIS_DATA.filter(a => a.refLabel==='MAINT0910' && a.nature==='P' && a.delaiLabel!=='P-18 mois' && a.delaiLabel!=='IMMÉDIAT'); break;
    case 'brc-r':     filtered = AVIS_DATA.filter(a => a.refLabel==='MAINT0910' && a.nature==='R'); break;
    case 'rob-nacc':  filtered = AVIS_DATA.filter(a => a.refLabel==='MAINT0410' && normalizeStr(a.defail||'').includes('acces')); break;
    case 'rob-nman':  filtered = AVIS_DATA.filter(a => a.refLabel==='MAINT0410' && normalizeStr(a.defail||'').includes('manoeuv')); break;
    case 'lies':      filtered = AVIS_DATA.filter(a => a.otLie);           break;
    default:          filtered = [...AVIS_DATA];
  }
  filtered = [...filtered].sort((a,b) =>
    (PRIO_ORD_AVIS[a.prioGrp]??9) - (PRIO_ORD_AVIS[b.prioGrp]??9) || b.ageDays - a.ageDays
  );

  const tabLabels = {
    'f0':'Fuites_F0', 'sans-ot':'Sans_OT', 'ko':'Hors_Delai',
    'nonqualif':'Non_Qualifiees',
    'brc-p18':'BRC_P18', 'brc-p':'BRC_Prescrit', 'brc-r':'BRC_Recommande',
    'rob-nacc':'ROB_NACC', 'rob-nman':'ROB_NMAN',
    'lies':'Lies_a_OT', 'all':'Tous'
  };

  const rows = filtered.map(a => ({
    'N° Avis':            a.avis,
    'Priorité':           a.prio,
    'Défaillance':        a.defail,
    'Cause':              a.cause,
    'Nature MAINT0910':   a.nature,
    'Délai MAINT0910':    a.delaiLabel,
    'Conformité':         a.conformite,
    'Dépassement (j)':    a.depasse > 0 ? a.depasse : '',
    'Traitement prescrit':a.traitement,
    'Ville':              a.ville,
    'Adresse':            a.rue,
    'Ancienneté (j)':     a.ageDays,
    'Date début':         a.dateStr,
    'OT correctif':       a.otNum || '',
    'OT lié au fichier':  a.otLie ? 'Oui (' + a.otLie + ')' : (a.hasOT ? 'Externe' : 'Non'),
  }));

  const ws  = XLSX.utils.json_to_sheet(rows);
  ws['!autofilter'] = { ref: ws['!ref'] };
  ws['!cols'] = [{wch:12},{wch:14},{wch:40},{wch:30},{wch:8},{wch:14},{wch:22},{wch:10},
                 {wch:40},{wch:20},{wch:28},{wch:10},{wch:12},{wch:16},{wch:12}];
  const wbX = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wbX, ws, 'Avis PM');
  const label = tabLabels[AVIS_MODE] || 'Avis';
  XLSX.writeFile(wbX, `GRDF_Avis_PM_${label}_${new Date().toISOString().slice(0,10)}.xlsx`);
}

// ════════════════════════════════════════════════════════════════
//  TRI ET RECHERCHE DANS LES TABLEAUX
// ════════════════════════════════════════════════════════════════

// État de tri par table


function sortTable(table, th, col, type) {
  const state = SORT_STATE[table];
  // Inverser si même colonne
  state.dir = (state.col === col) ? -state.dir : 1;
  state.col = col;

  // Mise à jour visuelle des en-têtes
  const thead = th.closest('thead');
  thead.querySelectorAll('th').forEach(h => h.classList.remove('sort-asc','sort-desc'));
  th.classList.add(state.dir === 1 ? 'sort-asc' : 'sort-desc');

  // Comparateur selon type
  const cmp = (a, b) => {
    let va = a[col], vb = b[col];
    if (type === 'date') {
      va = va ? va.getTime() : 0;
      vb = vb ? vb.getTime() : 0;
      return (va - vb) * state.dir;
    }
    if (type === 'num') {
      return ((+va||0) - (+vb||0)) * state.dir;
    }
    if (type === 'prio') {
      return ((PRIO_ORD_SORT[va]??9) - (PRIO_ORD_SORT[vb]??9)) * state.dir;
    }
    return String(va||'').localeCompare(String(vb||''), 'fr', {sensitivity:'base'}) * state.dir;
  };

  if (table === 'prev') {
    ROWS_PREV = [...ROWS_PREV].sort(cmp);
    _renderPrevRows(ROWS_PREV, document.getElementById('search-prev').value);
  } else if (table === 'cor') {
    ROWS_COR = [...ROWS_COR].sort(cmp);
    _renderCorRows(ROWS_COR, document.getElementById('search-cor').value);
  } else {
    ROWS_AVIS = [...ROWS_AVIS].sort(cmp);
    _renderAvisRows(ROWS_AVIS, document.getElementById('search-avis').value);
  }
}

// ── Recherche texte ──────────────────────────────────────────
function filterTableSearch(table) {
  const query = document.getElementById('search-' + table).value;
  if (table === 'prev')  _renderPrevRows(ROWS_PREV, query);
  if (table === 'cor')   _renderCorRows(ROWS_COR, query);
  if (table === 'avis')  _renderAvisRows(ROWS_AVIS, query);
}

function clearTableSearch(table) {
  document.getElementById('search-' + table).value = '';
  filterTableSearch(table);
}

function matchSearch(row, query) {
  if (!query) return true;
  const q = query.toLowerCase();
  return Object.values(row).some(v =>
    v && String(v).toLowerCase().includes(q)
  );
}

// ── Renderers avec recherche ─────────────────────────────────

function _renderPrevRows(rows, query) {
  const tbody = document.getElementById('table-retard-body');
  tbody.innerHTML = '';
  const visible = rows.filter(d => matchSearch(d, query));
  document.getElementById('search-prev-count').textContent =
    query ? `${visible.length} / ${rows.length} résultats` : '';

  visible.forEach(d => {
    const tr = document.createElement('tr');
    const dateStr = d.dateRef ? d.dateRef.toLocaleDateString('fr-FR') : '—';
    const bc = d.statut==='TERM'?'badge-term':d.statut==='TERM ANO'?'badge-ano':
               d.statut==='ACTI'?'badge-acti':d.statut==='EPR'?'badge-epr':'badge-prg';
    const soLabel = d.statutOrdre.includes('commercialement') ? 'Clôturé commercialement'
                  : d.statutOrdre.includes('techniquement')   ? 'Clôturé techniquement' : 'Lancé';
    const soClass = d.statutOrdre.includes('commercialement') ? 'so-clocom'
                  : d.statutOrdre.includes('techniquement')   ? 'so-clotec' : 'so-lance';
    const natureCls = d.nature==='Préventif'?'badge-nature-prev':d.nature==='Correctif'?'badge-nature-cor':
                      d.nature==='Exploitation'?'badge-nature-exp':'badge-nature-autre';
    tr.innerHTML =
      `<td class="mono num-sap">${esc(d.numSAP)}</td>` +
      `<td class="mono ordre-libelle">${esc(d.ordre)}</td>` +
      `<td><strong>${esc(d.ouvrage)}</strong></td>` +
      `<td>${esc(d.typeTravail)}</td>` +
      `<td><span class="badge-nature ${natureCls}">${esc(d.nature)}</span></td>` +
      `<td><span class="badge-statut ${bc}">${esc(d.statut)}</span></td>` +
      `<td><span class="badge-so ${soClass}">${soLabel}</span></td>` +
      `<td>${esc(d.entite)}</td>` +
      `<td>${esc(d.ville)}</td>` +
      `<td class="${d.enRetard?'date-retard':''}">${dateStr}</td>`;
    tbody.appendChild(tr);
  });
}

function _renderCorRows(rows, query) {
  const tbody = document.getElementById('table-cor-body');
  tbody.innerHTML = '';
  const now = new Date();
  const visible = rows.filter(d => matchSearch(d, query));
  document.getElementById('search-cor-count').textContent =
    query ? `${visible.length} / ${rows.length} résultats` : '';

  visible.forEach(d => {
    const tr = document.createElement('tr');
    const dateStr = d.dateRef ? d.dateRef.toLocaleDateString('fr-FR') : '—';
    const bc = d.statut==='TERM'?'badge-term':d.statut==='TERM ANO'?'badge-ano':
               d.statut==='ACTI'?'badge-acti':d.statut==='EPR'?'badge-epr':'badge-prg';
    const soLabel = d.statutOrdre.includes('commercialement')?'Clôturé commercialement':
                    d.statutOrdre.includes('techniquement')?'Clôturé techniquement':'Lancé';
    const soClass = d.statutOrdre.includes('commercialement')?'so-clocom':
                    d.statutOrdre.includes('techniquement')?'so-clotec':'so-lance';
    const ancDays = d.dateRef ? Math.floor((now-d.dateRef)/86400000) : 0;
    const ancCls = ancDays>=365?'badge-anc-danger':ancDays>=180?'badge-anc-warn':ancDays>=30?'badge-anc-low':'badge-anc-ok';
    const ancTxt = d.termine||!d.statutOrdre.includes('Lancé') ? 'Clôturé'
                 : ancDays>=365 ? '>12 mois' : ancDays>=180 ? '6–12 mois'
                 : ancDays>=30  ? '1–6 mois' : '<1 mois';
    tr.innerHTML =
      `<td class="mono num-sap">${esc(d.numSAP)}</td>` +
      `<td class="mono ordre-libelle">${esc(d.ordre)}</td>` +
      `<td><strong>${esc(d.ouvrage)}</strong></td>` +
      `<td>${esc(d.typeTravail)}</td>` +
      `<td><span class="badge-statut ${bc}">${esc(d.statut)}</span></td>` +
      `<td><span class="badge-so ${soClass}">${soLabel}</span></td>` +
      `<td>${esc(d.entite)}</td>` +
      `<td>${esc(d.ville)}</td>` +
      `<td>${dateStr}</td>` +
      `<td><span class="badge-anciennete ${ancCls}">${ancTxt}</span></td>`;
    tbody.appendChild(tr);
  });
}

function _renderAvisRows(rows, query) {
  const tbody = document.getElementById('table-avis-body');
  tbody.innerHTML = '';
  const visible = rows.filter(a => matchSearch(a, query));
  document.getElementById('search-avis-count').textContent =
    query ? `${visible.length} / ${rows.length} résultats` : '';

  const prioLbl = { F0:'F0 Fuite', P1:'P1 Prescrit', P2:'P2 Prescrit',
                    R1:'R1 Recom.', R2:'R2 Recom.', R3:'R3 Recom.', Autre:'Autre' };
  const prioCls = { F0:'prio-f0', P1:'prio-p1', P2:'prio-p2',
                    R1:'prio-r1', R2:'prio-r2', R3:'prio-r3', Autre:'prio-autre' };
  const confCls = { OK:'conf-ok', KO:'conf-ko', INFO:'conf-info', NF:'conf-nf', '?':'conf-unk' };

  visible.forEach(a => {
    const tr  = document.createElement('tr');
    const pcl = prioCls[a.prioGrp] || 'prio-autre';
    const ccl = confCls[a.confCode] || 'conf-unk';
    const ageCl = a.ageDays>730?'age-crit':a.ageDays>365?'age-warn':'';
    let otCell;
    if (a.otLie)      otCell = `<span class="badge-ot-lie">${esc(a.otLie)}</span>`;
    else if (a.hasOT) otCell = `<span class="badge-ot-ext">${esc(a.otNum.substring(0,22))}</span>`;
    else              otCell = `<span class="no-ot">Aucun OT</span>`;

    tr.innerHTML =
      `<td class="mono" style="font-size:11px">${esc(a.avis)}</td>` +
      `<td><span class="badge-prio ${pcl}">${prioLbl[a.prioGrp]||a.prioGrp}</span></td>` +
      `<td style="max-width:230px;white-space:normal;line-height:1.3"><strong style="font-size:11.5px">${esc(a.defail)}</strong></td>` +
      `<td style="font-size:10.5px;color:#64748b">${esc(a.delaiLabel)}</td>` +
      `<td><span class="badge-conf ${ccl}">${esc(a.conformite)}</span></td>` +
      `<td>${esc(a.ville)}</td>` +
      `<td style="font-size:11px;color:#475569;max-width:190px;white-space:normal">${esc(a.rue)}</td>` +
      `<td class="${ageCl}" style="font-family:'DM Mono',monospace;font-size:11.5px">${a.ageDays>0?a.ageDays+'j':a.dateStr}</td>` +
      `<td>${otCell}</td>`;
    tbody.appendChild(tr);
  });
}
