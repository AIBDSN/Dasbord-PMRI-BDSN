/* ============================================================
   V3 - Dashboard helpers (loaded before dashboard.js)
   ============================================================ */

(function () {
  function destroyChart(id) { if(CHARTS[id]){CHARTS[id].destroy();delete CHARTS[id];} }
  function countBy(data,key) { const m={}; data.forEach(d=>{const v=d[key]||'N/A';m[v]=(m[v]||0)+1;}); return m; }
  function esc(str) { return String(str||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
  function getFilterValue(id, fallback='') {
    const el = document.getElementById(id);
    return el ? (el.value || fallback) : fallback;
  }

  window.GRDFDashV3 = window.GRDFDashV3 || {};
  window.GRDFDashV3.coreHelpers = { destroyChart, countBy, esc, getFilterValue };
})();
