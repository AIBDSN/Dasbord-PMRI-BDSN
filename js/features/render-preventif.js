/* ============================================================
   V3 - Render preventif (loaded before dashboard.js)
   ============================================================ */

(function () {
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
  window.GRDFDashV3 = window.GRDFDashV3 || {};
  window.GRDFDashV3.features = window.GRDFDashV3.features || {};
  window.GRDFDashV3.features.renderPreventif = {
    renderChartStatut, renderChartOuvrageBar, setTableMode,
    renderTable, renderTableRows, _renderPrevRows,
  };
})();

