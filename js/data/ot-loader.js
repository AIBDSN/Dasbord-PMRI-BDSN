/* ============================================================
   V3 - OT file loading (loaded before dashboard.js)
   ============================================================ */

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
      document.getElementById('tab-prev-count').textContent = FILTERED_PREV.length.toLocaleString('fr-FR');
      document.getElementById('tab-cor-count').textContent  = FILTERED_COR.length.toLocaleString('fr-FR');
      render();
      document.getElementById('empty-state').style.display = 'none';
      document.getElementById('dashboard').style.display   = 'block';
    } catch(err) { alert('Erreur lecture fichier : ' + err.message); }
  };
  reader.readAsArrayBuffer(file);
}

window.GRDFDashV3 = window.GRDFDashV3 || {};
window.GRDFDashV3.loaders = window.GRDFDashV3.loaders || {};
window.GRDFDashV3.loaders.loadFile = loadFile;
window.loadFile = loadFile;
