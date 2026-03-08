/* ============================================================
   V3 - Weekly trends indicators (loaded before dashboard.js)
   ============================================================ */

(function () {
  function getISOWeekInfo(d) {
    const date = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    const dayNum = date.getUTCDay() || 7;
    date.setUTCDate(date.getUTCDate() + 4 - dayNum);
    const isoYear = date.getUTCFullYear();
    const yearStart = new Date(Date.UTC(isoYear, 0, 1));
    const isoWeek = Math.ceil((((date - yearStart) / 86400000) + 1) / 7);
    return { year: isoYear, week: isoWeek };
  }

  function weeksInISOYear(year) {
    return getISOWeekInfo(new Date(Date.UTC(year, 11, 31))).week;
  }

  function previousWeekOf(info) {
    if (info.week > 1) return { year: info.year, week: info.week - 1 };
    const prevYear = info.year - 1;
    return { year: prevYear, week: weeksInISOYear(prevYear) };
  }

  function sameISOWeek(date, info) {
    if (!date) return false;
    const d = (date instanceof Date) ? date : new Date(date);
    if (isNaN(d)) return false;
    const wk = getISOWeekInfo(d);
    return wk.year === info.year && wk.week === info.week;
  }

  function summarize(current, prev) {
    const delta = current - prev;
    const pct = prev !== 0 ? ((delta / prev) * 100) : null;
    const trend = delta > 0 ? 'up' : delta < 0 ? 'down' : 'flat';
    return { current, prev, delta, pct, trend };
  }

  function getAvisSource() {
    if (typeof getAvisDataSource === 'function') return getAvisDataSource();
    if (Array.isArray(FILTERED_AVIS)) return FILTERED_AVIS;
    if (Array.isArray(AVIS_DATA)) return AVIS_DATA;
    return [];
  }

  function computeWeeklyTrends() {
    const now = new Date();
    const currentWeek = getISOWeekInfo(now);
    const prevWeek = previousWeekOf(currentWeek);

    const prevSource = Array.isArray(FILTERED_PREV) ? FILTERED_PREV : [];
    const corSource = Array.isArray(FILTERED_COR) ? FILTERED_COR : [];
    const avisSource = getAvisSource();

    const prevDoneDate = d => d.dateDebut || d.dateRef;
    const corCreationDate = d => d.dateRef;
    const corClosureDate = d => d.dateRef;
    const avisCreationDate = a => a.dateRef;

    const prevRealisesCurrent = prevSource.filter(d => d.termine && sameISOWeek(prevDoneDate(d), currentWeek)).length;
    const prevRealisesPrev = prevSource.filter(d => d.termine && sameISOWeek(prevDoneDate(d), prevWeek)).length;

    const corCreationsCurrent = corSource.filter(d => sameISOWeek(corCreationDate(d), currentWeek)).length;
    const corCreationsPrev = corSource.filter(d => sameISOWeek(corCreationDate(d), prevWeek)).length;

    const corCloturesCurrent = corSource.filter(d =>
      (d.termine || !d.statutOrdre.includes('Lancé')) && sameISOWeek(corClosureDate(d), currentWeek)
    ).length;
    const corCloturesPrev = corSource.filter(d =>
      (d.termine || !d.statutOrdre.includes('Lancé')) && sameISOWeek(corClosureDate(d), prevWeek)
    ).length;

    const avisNouveauxCurrent = avisSource.filter(a => sameISOWeek(avisCreationDate(a), currentWeek)).length;

    return {
      week: currentWeek,
      prevWeek: prevWeek,
      preventif: {
        realises: summarize(prevRealisesCurrent, prevRealisesPrev),
      },
      avis: {
        stock: { current: avisSource.length },
        nouveaux: { current: avisNouveauxCurrent },
      },
      correctif: {
        creations: summarize(corCreationsCurrent, corCreationsPrev),
        clotures: summarize(corCloturesCurrent, corCloturesPrev),
        soldeNet: {
          current: corCreationsCurrent - corCloturesCurrent,
          prev: corCreationsPrev - corCloturesPrev,
          delta: (corCreationsCurrent - corCloturesCurrent) - (corCreationsPrev - corCloturesPrev),
          trend: (corCreationsCurrent - corCloturesCurrent) > (corCreationsPrev - corCloturesPrev) ? 'up'
               : (corCreationsCurrent - corCloturesCurrent) < (corCreationsPrev - corCloturesPrev) ? 'down'
               : 'flat',
        },
      },
    };
  }

  function formatVsS1(metric) {
    const arrow = metric.trend === 'up' ? '↗' : metric.trend === 'down' ? '↘' : '→';
    const sign = metric.delta > 0 ? '+' : '';
    return `${arrow} ${sign}${metric.delta} vs S-1`;
  }

  function toneColor(tone) {
    if (tone === 'positive') return '#15803d';
    if (tone === 'negative') return '#b91c1c';
    return '#475569';
  }

  function upsertKpiExtra(valueId, key, text, tone) {
    const valueEl = document.getElementById(valueId);
    const card = valueEl ? valueEl.closest('.kpi') : null;
    if (!card) return;

    let el = card.querySelector(`[data-kpi-extra="${key}"]`);
    if (!el) {
      el = document.createElement('div');
      el.setAttribute('data-kpi-extra', key);
      el.style.fontSize = '11px';
      el.style.fontWeight = '600';
      el.style.marginTop = '4px';
      el.style.lineHeight = '1.2';
      card.appendChild(el);
    }
    el.style.color = toneColor(tone);
    el.textContent = text;
  }

  window.GRDFDashV3 = window.GRDFDashV3 || {};
  window.GRDFDashV3.services = window.GRDFDashV3.services || {};
  window.GRDFDashV3.services.weeklyTrends = {
    computeWeeklyTrends,
    formatVsS1,
    upsertKpiExtra,
  };
})();
