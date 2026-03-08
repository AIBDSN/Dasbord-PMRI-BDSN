/* ============================================================
   V3 - Filters service (loaded before dashboard.js)
   ============================================================ */

(function () {
  function filterAvisData(entite, ouvrage) {
    if (!Array.isArray(AVIS_DATA) || !AVIS_DATA.length) return [];

    const otByNumSAP = {};
    ALL_DATA.forEach(d => { otByNumSAP[d.numSAP] = d; });

    return AVIS_DATA.filter(a => {
      const linkedOT = a.otLie ? otByNumSAP[a.otLie] : null;

      const entiteMatch = !entite
        || (a.entiteAvis && a.entiteAvis === entite)
        || (linkedOT && linkedOT.entite === entite);

      const ouvrageMatch = !ouvrage
        || (a.ouvrageAvis && a.ouvrageAvis === ouvrage)
        || (linkedOT && linkedOT.ouvrage === ouvrage)
        || (ouvrage === 'Robinet Réseau' && a.refLabel === 'MAINT0410')
        || (ouvrage === 'CICM / BRC' && a.refLabel === 'MAINT0910');

      return entiteMatch && ouvrageMatch;
    });
  }

  function populateSelect(id, values) {
    const sel = document.getElementById(id);
    sel.innerHTML = '<option value="">Tous</option>';
    values.filter(Boolean).forEach(v => {
      const o = document.createElement('option');
      o.value = v; o.textContent = v; sel.appendChild(o);
    });
  }

  function initFilters() {
    populateSelect('filter-entite',  [...new Set(ALL_DATA.map(d=>d.entite))].sort());
    populateSelect('filter-ouvrage', [...new Set(ALL_DATA.map(d=>d.ouvrage))].sort());
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
    FILTERED_AVIS = filterAvisData(entite, ouvrage);

    if (CURRENT_MODE === 'avis') renderAvis();
    else render();
  }

  function resetFilters() {
    ['filter-entite','filter-ouvrage'].forEach(id => document.getElementById(id).value = '');
    FILTERED_PREV = ALL_DATA.filter(d => d.nature === 'Préventif' || d.nature === 'Exploitation');
    FILTERED_COR  = ALL_DATA.filter(d => d.nature === 'Correctif');
    FILTERED = CURRENT_MODE === 'correctif' ? FILTERED_COR : FILTERED_PREV;
    FILTERED_AVIS = [...AVIS_DATA];
    if (CURRENT_MODE === 'avis') renderAvis();
    else render();
  }

  window.GRDFDashV3 = window.GRDFDashV3 || {};
  window.GRDFDashV3.services = window.GRDFDashV3.services || {};
  window.GRDFDashV3.services.filters = { populateSelect, initFilters, applyFilters, resetFilters };
})();
