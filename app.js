(() => {
  'use strict';

  // ---- Configuration -------------------------------------------------
  // Google Sheet id, taken from the share link. Each entry in SHEETS is one
  // tab in the workbook = one strategic goal ("الهدف الاستراتيجي"). gid
  // identifies each tab uniquely (see the sheet's URL after "#gid=" when its
  // tab is open) and is immune to whitespace/renaming issues that would break
  // exact-name matching. If a goal is renamed / added / removed, update this list.
  const SHEET_ID = '1sdtmMervz3DEMSLlFGUrwX-IU1uX_DghGiTiTIm_RBs';
  const SHEETS = [
    { gid: '464265919', name: 'بناء الهوية القيمية' },
    { gid: '696696481', name: 'تعزيز السلوك الايجابي' },
    { gid: '1383058834', name: 'رفع وعي الوالدين' },
    { gid: '1501680650', name: 'تأهيل الممارسين مع الطفل' },
    { gid: '933553380', name: 'بناء موارد مالية مستدامة' },
    { gid: '1849689338', name: 'ابتكار برامج وخدمات نوعية' },
    { gid: '1775610743', name: 'بناء الشراكات الفاعلة' },
    { gid: '1392541017', name: 'تحسين الحوكمة الأساسية' },
    { gid: '319288641', name: 'تأهيل الكوادر الوظيفية' },
  ];

  const TATWEEL = /ـ/g;

  function cleanLabel(v) {
    if (v === null || v === undefined) return '';
    return String(v).replace(TATWEEL, '').trim();
  }

  function normText(v) {
    if (v === null || v === undefined) return '';
    return String(v).trim();
  }

  function toNumber(v) {
    if (v === null || v === undefined) return null;
    if (typeof v === 'number') return v;
    const s = String(v).trim().replace(/[,،٬\s]/g, '');
    if (s === '') return null;
    const n = parseFloat(s);
    return Number.isNaN(n) ? null : n;
  }

  function findLabelCell(row, prefix) {
    for (let i = 0; i < row.length; i++) {
      const v = row[i];
      if (v === null || v === undefined) continue;
      if (cleanLabel(v).startsWith(prefix)) return i;
    }
    return null;
  }

  function valueAfter(row, idx) {
    for (let j = idx + 1; j < row.length; j++) {
      if (row[j] !== null && row[j] !== undefined && row[j] !== '') return row[j];
    }
    return null;
  }

  // ---- Fetch + parse one sheet ----------------------------------------

  // The gviz JSON endpoint infers one data type per column and silently nulls
  // out any cell that doesn't fit it (e.g. a text header sitting in a column
  // whose other rows are numeric). This sheet mixes label rows and numeric
  // task rows in the same columns, so we use the CSV export instead: it
  // returns raw displayed text for every cell with no type coercion.
  function parseCSV(text) {
    const rows = [];
    let row = [];
    let field = '';
    let inQuotes = false;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (inQuotes) {
        if (c === '"') {
          if (text[i + 1] === '"') { field += '"'; i++; }
          else inQuotes = false;
        } else {
          field += c;
        }
      } else if (c === '"') {
        inQuotes = true;
      } else if (c === ',') {
        row.push(field); field = '';
      } else if (c === '\r') {
        // skip, \n handles the line break
      } else if (c === '\n') {
        row.push(field); field = '';
        rows.push(row); row = [];
      } else {
        field += c;
      }
    }
    if (field !== '' || row.length) { row.push(field); rows.push(row); }
    return rows.map(r => r.map(v => (v === '' ? null : v)));
  }

  async function fetchSheetGrid(gid, sheetName) {
    const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv&gid=${gid}`;
    const res = await fetch(url, { cache: 'no-store' });
    if (!res.ok) throw new Error(`تعذر جلب الورقة: ${sheetName} (${res.status})`);
    const text = await res.text();
    return parseCSV(text);
  }

  function parseSheet(goalName, grid) {
    const initiatives = [];
    let i = 0;
    const n = grid.length;

    while (i < n) {
      const row = grid[i];
      const projIdx = findLabelCell(row, 'المشروع');
      if (projIdx === null) { i++; continue; }

      const record = {
        goal: goalName,
        name: normText(valueAfter(row, projIdx)),
        seq: null,
        description: '',
        indicator: '',
        target: null,
        achieved: null,
        responsible: '',
        cost: 0,
      };

      let j = i + 1;
      let costCol = null;
      let inSubtable = false;

      while (j < n) {
        const r2 = grid[j];

        if (j !== i && findLabelCell(r2, 'المشروع') !== null) break;

        let ci = findLabelCell(r2, 'تسلسل');
        if (ci !== null) record.seq = valueAfter(r2, ci);

        ci = findLabelCell(r2, 'شرح');
        if (ci !== null) record.description = normText(valueAfter(r2, ci));

        ci = findLabelCell(r2, 'مؤشر أداء');
        if (ci !== null) {
          record.indicator = normText(valueAfter(r2, ci));
          const ti = findLabelCell(r2, 'المستهدف');
          if (ti !== null) record.target = toNumber(valueAfter(r2, ti));
          const ai = findLabelCell(r2, 'المتحقق');
          if (ai !== null) record.achieved = toNumber(valueAfter(r2, ai));
        }

        ci = findLabelCell(r2, 'القسم المنفذ');
        if (ci !== null) record.responsible = normText(valueAfter(r2, ci));

        if (findLabelCell(r2, 'خطوات التنفيذ') !== null) {
          for (let k = 0; k < r2.length; k++) {
            if (r2[k] !== null && cleanLabel(r2[k]) === 'التكلفة') costCol = k;
          }
          inSubtable = true;
          j++;
          continue;
        }

        if (inSubtable) {
          if (findLabelCell(r2, 'نسبة إنجاز') !== null) {
            inSubtable = false;
          } else if (costCol !== null && costCol < r2.length) {
            const num = toNumber(r2[costCol]);
            if (num !== null) record.cost += num;
          }
        }

        j++;
      }

      initiatives.push(record);
      i = j;
    }

    return initiatives;
  }

  // ---- Load all sheets --------------------------------------------------

  async function loadAllData() {
    const settled = await Promise.allSettled(
      SHEETS.map(s => fetchSheetGrid(s.gid, s.name).then(grid => parseSheet(s.name, grid)))
    );

    const initiatives = [];
    const errors = [];
    settled.forEach((res, idx) => {
      if (res.status === 'fulfilled') {
        initiatives.push(...res.value);
      } else {
        errors.push(`${SHEETS[idx].name}: ${res.reason.message || res.reason}`);
      }
    });

    return { initiatives, errors };
  }

  // ---- Rendering ----------------------------------------------------

  const fmtInt = n => Math.round(n).toLocaleString('en-US');

  function renderKPIs(initiatives) {
    const goals = new Set(initiatives.map(x => x.goal));
    const totalCost = initiatives.reduce((s, x) => s + (x.cost || 0), 0);

    let targetSum = 0, achievedSum = 0, withTarget = 0;
    initiatives.forEach(x => {
      if (x.target !== null && x.target > 0) {
        targetSum += x.target;
        achievedSum += Math.min(x.achieved || 0, x.target);
        withTarget++;
      }
    });
    const completion = targetSum > 0 ? (achievedSum / targetSum) * 100 : 0;

    document.getElementById('kpiGoals').textContent = fmtInt(goals.size);
    document.getElementById('kpiInitiatives').textContent = fmtInt(initiatives.length);
    document.getElementById('kpiCompletion').textContent = withTarget ? `${completion.toFixed(1)}%` : '—';
    document.getElementById('kpiCost').textContent = fmtInt(totalCost);
  }

  function renderChart(initiatives) {
    const byGoal = new Map();
    initiatives.forEach(x => {
      byGoal.set(x.goal, (byGoal.get(x.goal) || 0) + (x.cost || 0));
    });
    const rows = Array.from(byGoal.entries()).sort((a, b) => b[1] - a[1]);
    const container = document.getElementById('costChart');

    if (!rows.length || rows.every(r => r[1] === 0)) {
      container.innerHTML = '<div class="chart-empty">لا توجد بيانات تكلفة متاحة بعد</div>';
      return;
    }

    const max = Math.max(...rows.map(r => r[1]), 1);
    const rowH = 34;
    const gap = 10;
    const topPad = 8;
    const height = rows.length * (rowH + gap) + topPad;
    const width = 900;
    const labelW = 230;
    const valueW = 90;
    const trackX = labelW;
    const trackW = width - labelW - valueW;

    let bars = '';
    rows.forEach(([goal, cost], idx) => {
      const y = topPad + idx * (rowH + gap);
      const barW = Math.max((cost / max) * trackW, cost > 0 ? 3 : 0);
      const label = goal.length > 28 ? goal.slice(0, 27) + '…' : goal;
      bars += `
        <g>
          <text class="bar-label" x="${width - 6}" y="${y + rowH / 2 + 4}" text-anchor="end">${escapeXML(label)}</text>
          <rect class="bar-track" x="${trackX}" y="${y}" width="${trackW}" height="${rowH * 0.55}" rx="6" transform="translate(0, ${rowH * 0.225})" />
          <rect class="bar-fill" x="${trackX + trackW - barW}" y="${y}" width="${barW}" height="${rowH * 0.55}" rx="6" transform="translate(0, ${rowH * 0.225})">
            <title>${escapeXML(goal)}: ${fmtInt(cost)} ر.س</title>
          </rect>
          <text class="bar-value" x="${trackX + trackW - barW - 8}" y="${y + rowH / 2 + 4}" text-anchor="end">${fmtInt(cost)}</text>
        </g>`;
    });

    container.innerHTML = `<svg viewBox="0 0 ${width} ${height}" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="التكلفة حسب الهدف الاستراتيجي">${bars}</svg>`;
  }

  function escapeXML(s) {
    return String(s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function populateGoalFilter(initiatives) {
    const select = document.getElementById('goalFilter');
    const existing = new Set(Array.from(select.options).map(o => o.value));
    const goals = Array.from(new Set(initiatives.map(x => x.goal)));
    goals.forEach(g => {
      if (!existing.has(g)) {
        const opt = document.createElement('option');
        opt.value = g;
        opt.textContent = g;
        select.appendChild(opt);
      }
    });
  }

  function renderTable(initiatives) {
    const search = document.getElementById('searchInput').value.trim().toLowerCase();
    const goalFilter = document.getElementById('goalFilter').value;
    const body = document.getElementById('initiativesBody');

    const filtered = initiatives.filter(x => {
      if (goalFilter && x.goal !== goalFilter) return false;
      if (!search) return true;
      const hay = `${x.name} ${x.indicator} ${x.responsible} ${x.description}`.toLowerCase();
      return hay.includes(search);
    });

    if (!filtered.length) {
      body.innerHTML = '<tr><td colspan="9" class="empty-row">لا توجد نتائج مطابقة</td></tr>';
      return;
    }

    body.innerHTML = filtered.map(x => {
      const target = x.target;
      const achieved = x.achieved;
      const pct = (target && target > 0) ? Math.min((achieved || 0) / target * 100, 100) : null;
      const pctLabel = pct === null ? '—' : `${pct.toFixed(0)}%`;
      const progressBar = pct === null ? '—' : `
        <div class="progress-cell">
          <div class="progress-track"><div class="progress-fill" style="width:${pct}%"></div></div>
          <span class="progress-pct">${pctLabel}</span>
        </div>`;

      return `
        <tr>
          <td><span class="goal-tag">${escapeXML(x.goal)}</span></td>
          <td class="name-cell">
            <strong>${escapeXML(x.name || '—')}</strong>
            ${x.seq !== null ? `<span class="seq-badge">مبادرة رقم ${escapeXML(String(x.seq))}</span>` : ''}
          </td>
          <td><div class="desc-cell" title="${escapeXML(x.description)}">${escapeXML(x.description || '—')}</div></td>
          <td>${escapeXML(x.responsible || '—')}</td>
          <td>${escapeXML(x.indicator || '—')}</td>
          <td class="num">${target !== null ? fmtInt(target) : '—'}</td>
          <td class="num">${achieved !== null ? fmtInt(achieved) : '—'}</td>
          <td>${progressBar}</td>
          <td class="num">${x.cost ? fmtInt(x.cost) : '—'}</td>
        </tr>`;
    }).join('');
  }

  function renderStatus(errors) {
    const banner = document.getElementById('statusBanner');
    if (!errors.length) {
      banner.hidden = true;
      return;
    }
    banner.hidden = false;
    banner.innerHTML = `تعذر تحميل بعض البيانات: ${errors.map(escapeXML).join(' — ')}`;
  }

  // ---- Main -----------------------------------------------------------

  let allInitiatives = [];

  async function refresh() {
    const btn = document.getElementById('refreshBtn');
    btn.disabled = true;
    btn.classList.add('spinning');
    try {
      const { initiatives, errors } = await loadAllData();
      allInitiatives = initiatives;
      renderStatus(errors);
      renderKPIs(initiatives);
      renderChart(initiatives);
      populateGoalFilter(initiatives);
      renderTable(initiatives);
      document.getElementById('lastUpdated').textContent =
        `آخر تحديث: ${new Date().toLocaleString('ar-SA')}`;
    } catch (err) {
      renderStatus([err.message || String(err)]);
      document.getElementById('initiativesBody').innerHTML =
        '<tr><td colspan="9" class="empty-row">تعذر تحميل البيانات</td></tr>';
    } finally {
      btn.disabled = false;
      btn.classList.remove('spinning');
    }
  }

  document.getElementById('refreshBtn').addEventListener('click', refresh);
  document.getElementById('searchInput').addEventListener('input', () => renderTable(allInitiatives));
  document.getElementById('goalFilter').addEventListener('change', () => renderTable(allInitiatives));

  refresh();
})();
