// Walmart Sales & Google Sheets Studio Logic
document.addEventListener('DOMContentLoaded', () => {
  // Elements
  const tabGoogleSheetsBtn = document.getElementById('tab-google-sheets-btn');
  const tabLocalViewerBtn = document.getElementById('tab-local-viewer-btn');
  const tabProcedureBtn = document.getElementById('tab-procedure-btn');
  const tabGoogleSheetsContent = document.getElementById('tab-google-sheets-content');
  const tabLocalViewerContent = document.getElementById('tab-local-viewer-content');
  const tabProcedureContent = document.getElementById('tab-procedure-content');

  // Embed elements
  const sheetsUrlInput = document.getElementById('sheets-url-input');
  const btnApplySheetUrl = document.getElementById('btn-apply-sheet-url');
  const googleSheetsFrame = document.getElementById('google-sheets-frame');
  const btnToggleInstructions = document.getElementById('btn-toggle-instructions');
  const instructionsPanel = document.getElementById('instructions-panel');
  const btnFullscreenEmbed = document.getElementById('btn-fullscreen-embed');
  const iframeWrapper = document.getElementById('iframe-wrapper');
  const iframeLoading = document.getElementById('iframe-loading');

  // KPI elements
  const totalSalesValue = document.getElementById('total-sales-value');
  const totalRecordsValue = document.getElementById('total-records-value');
  const totalStoresValue = document.getElementById('total-stores-value');
  const topDeptValue = document.getElementById('top-dept-value');
  const topDeptSub = document.getElementById('top-dept-sub');
  const deptsBarsContainer = document.getElementById('depts-bars-container');

  // Local table elements
  const sheetsTabsList = document.getElementById('sheets-tabs-list');
  const sheetSearchInput = document.getElementById('sheet-search-input');
  const sheetMetaInfo = document.getElementById('sheet-meta-info');
  const btnPrevPage = document.getElementById('btn-prev-page');
  const btnNextPage = document.getElementById('btn-next-page');
  const pageIndicator = document.getElementById('page-indicator');
  const tableHead = document.getElementById('table-head');
  const tableBody = document.getElementById('table-body');

  // State
  let currentSheet = 'raw_departamento';
  let currentPage = 1;
  let totalPages = 1;
  let currentSearch = '';
  let searchTimeout = null;

  // 1. Tab Switching
  function switchTab(target) {
    tabGoogleSheetsBtn.classList.remove('active');
    tabLocalViewerBtn.classList.remove('active');
    if (tabProcedureBtn) tabProcedureBtn.classList.remove('active');

    tabGoogleSheetsContent.style.display = 'none';
    tabLocalViewerContent.style.display = 'none';
    if (tabProcedureContent) tabProcedureContent.style.display = 'none';

    if (target === 'sheets') {
      tabGoogleSheetsBtn.classList.add('active');
      tabGoogleSheetsContent.style.display = 'block';
    } else if (target === 'local') {
      tabLocalViewerBtn.classList.add('active');
      tabLocalViewerContent.style.display = 'block';
      if (!sheetsTabsList.hasChildNodes()) {
        loadSummary();
      }
    } else if (target === 'procedure') {
      if (tabProcedureBtn) tabProcedureBtn.classList.add('active');
      if (tabProcedureContent) tabProcedureContent.style.display = 'block';
    }
  }

  tabGoogleSheetsBtn.addEventListener('click', () => switchTab('sheets'));
  tabLocalViewerBtn.addEventListener('click', () => switchTab('local'));
  if (tabProcedureBtn) {
    tabProcedureBtn.addEventListener('click', () => switchTab('procedure'));
  }

  // 2. Google Sheets URL Converter & Loader
  function formatSheetsEmbedUrl(rawUrl, mode = 'preview') {
    const trimmed = rawUrl.trim();
    if (!trimmed) return '';

    // Extract sheet ID: /d/([a-zA-Z0-9_-]+)
    const idMatch = trimmed.match(/\/d\/([a-zA-Z0-9_-]+)/);
    // Extract gid if present
    const gidMatch = trimmed.match(/[?&#]gid=([0-9]+)/);
    const gid = gidMatch ? gidMatch[1] : null;

    if (idMatch && idMatch[1]) {
      const sheetId = idMatch[1];
      if (mode === 'pubhtml' || trimmed.includes('/pubhtml')) {
        const gidParam = gid ? `&gid=${gid}` : '';
        return `https://docs.google.com/spreadsheets/d/${sheetId}/pubhtml?widget=true&headers=false${gidParam}`;
      }
      // Preview mode: Works without needing "Publicar en la web"
      const gidParam = gid ? `?gid=${gid}` : '';
      return `https://docs.google.com/spreadsheets/d/${sheetId}/preview${gidParam}`;
    }

    // If it's already an embed URL without standard /d/ pattern
    return trimmed;
  }

  const sheetsModeSelect = document.getElementById('sheets-mode-select');

  function loadGoogleSheet(url) {
    const mode = sheetsModeSelect ? sheetsModeSelect.value : 'preview';
    const embedUrl = formatSheetsEmbedUrl(url, mode);
    if (!embedUrl) return;

    iframeLoading.classList.add('visible');
    googleSheetsFrame.src = embedUrl;

    // Remove loading overlay once loaded
    googleSheetsFrame.onload = () => {
      iframeLoading.classList.remove('visible');
    };
    setTimeout(() => {
      iframeLoading.classList.remove('visible');
    }, 4000);
  }

  if (sheetsModeSelect) {
    sheetsModeSelect.addEventListener('change', () => {
      loadGoogleSheet(sheetsUrlInput.value);
    });
  }

  btnApplySheetUrl.addEventListener('click', () => {
    loadGoogleSheet(sheetsUrlInput.value);
  });

  sheetsUrlInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      loadGoogleSheet(sheetsUrlInput.value);
    }
  });

  // Toggle instructions
  btnToggleInstructions.addEventListener('click', () => {
    instructionsPanel.classList.toggle('open');
  });

  // Toggle fullscreen embed
  btnFullscreenEmbed.addEventListener('click', () => {
    iframeWrapper.classList.toggle('fullscreen');
    if (iframeWrapper.classList.contains('fullscreen')) {
      btnFullscreenEmbed.textContent = '✕ Reducir';
    } else {
      btnFullscreenEmbed.textContent = '⛶ Expandir';
    }
  });

  // 3. Load Excel Summary & KPIs
  async function loadSummary() {
    try {
      const res = await fetch('/api/excel/summary');
      const data = await res.json();

      if (data.error) return;

      // Fill KPIs
      totalSalesValue.textContent = data.formattedSales || '$1,912,142,474';
      totalRecordsValue.textContent = (data.totalTransactions || 95880).toLocaleString();
      totalStoresValue.textContent = `${data.totalStores || 45} Sucursales`;

      if (data.topDepartments && data.topDepartments.length > 0) {
        const top1 = data.topDepartments[0];
        topDeptValue.textContent = top1.name;
        const pct = ((top1.sales / data.totalSales) * 100).toFixed(1);
        topDeptSub.textContent = `$${(top1.sales / 1000000).toFixed(1)}M (${pct}% del total)`;

        // Render Top Departments bars
        renderDepartmentBars(data.topDepartments, data.totalSales);
      }

      // Render Sheet Tabs
      if (data.sheetNames && data.sheetNames.length > 0) {
        renderSheetTabs(data.sheetNames);
        currentSheet = data.sheetNames.includes('raw_departamento') ? 'raw_departamento' : data.sheetNames[0];
        loadSheetData();
      }
    } catch (err) {
      console.error('Error fetching summary:', err);
    }
  }

  function renderDepartmentBars(topDepts, totalSales) {
    deptsBarsContainer.innerHTML = '';
    topDepts.forEach((dept) => {
      const pct = Math.min(100, Math.max(5, (dept.sales / totalSales) * 100)).toFixed(1);
      const formattedM = (dept.sales / 1000000).toFixed(2);

      const row = document.createElement('div');
      row.className = 'dept-bar-row';
      row.innerHTML = `
        <div class="dept-bar-labels">
          <span class="dept-bar-name">${dept.name}</span>
          <span class="dept-bar-val">$${formattedM}M (${pct}%)</span>
        </div>
        <div class="dept-progress-track">
          <div class="dept-progress-fill" style="width: ${pct}%"></div>
        </div>
      `;
      deptsBarsContainer.appendChild(row);
    });
  }

  function renderSheetTabs(sheetNames) {
    sheetsTabsList.innerHTML = '';
    sheetNames.forEach((name) => {
      const chip = document.createElement('button');
      chip.className = `sheet-tab-chip ${name === currentSheet ? 'active' : ''}`;
      chip.textContent = name;
      chip.addEventListener('click', () => {
        document.querySelectorAll('.sheet-tab-chip').forEach(c => c.classList.remove('active'));
        chip.classList.add('active');
        currentSheet = name;
        currentPage = 1;
        currentSearch = '';
        sheetSearchInput.value = '';
        loadSheetData();
      });
      sheetsTabsList.appendChild(chip);
    });
  }

  // 4. Load Paginated Sheet Data
  async function loadSheetData() {
    sheetMetaInfo.textContent = `Cargando '${currentSheet}'...`;
    try {
      const url = `/api/excel/sheet?name=${encodeURIComponent(currentSheet)}&page=${currentPage}&limit=50&search=${encodeURIComponent(currentSearch)}`;
      const res = await fetch(url);
      const data = await res.json();

      if (data.error) {
        sheetMetaInfo.textContent = `Error: ${data.error}`;
        return;
      }

      totalPages = data.totalPages || 1;
      pageIndicator.textContent = `Página ${data.page} de ${totalPages}`;
      sheetMetaInfo.textContent = `Hoja: ${data.sheetName} · ${data.totalRows.toLocaleString()} registros encontrados`;

      btnPrevPage.disabled = data.page <= 1;
      btnNextPage.disabled = data.page >= totalPages;

      renderTable(data.headers, data.rows);
    } catch (err) {
      console.error('Error fetching sheet:', err);
      sheetMetaInfo.textContent = 'Error al cargar datos de la hoja';
    }
  }

  function renderTable(headers, rows) {
    // Header
    tableHead.innerHTML = '';
    const trHead = document.createElement('tr');
    headers.forEach((h, idx) => {
      const th = document.createElement('th');
      th.textContent = h !== undefined && h !== null && h !== '' ? h : `Columna ${idx + 1}`;
      trHead.appendChild(th);
    });
    tableHead.appendChild(trHead);

    // Body
    tableBody.innerHTML = '';
    if (!rows || rows.length === 0) {
      const emptyTr = document.createElement('tr');
      const emptyTd = document.createElement('td');
      emptyTd.colSpan = Math.max(headers.length, 1);
      emptyTd.textContent = 'No hay registros en esta página';
      emptyTd.style.textAlign = 'center';
      emptyTd.style.padding = '30px';
      emptyTd.style.color = 'var(--text-dim)';
      emptyTr.appendChild(emptyTd);
      tableBody.appendChild(emptyTr);
      return;
    }

    rows.forEach((row) => {
      const tr = document.createElement('tr');
      headers.forEach((_, colIdx) => {
        const td = document.createElement('td');
        const val = row[colIdx];
        if (val === undefined || val === null || val === '') {
          td.textContent = '-';
          td.style.color = 'var(--text-dim)';
        } else if (typeof val === 'number') {
          td.textContent = Number.isInteger(val) ? val.toLocaleString() : val.toFixed(2);
          td.style.fontFamily = 'var(--font-mono)';
        } else {
          td.textContent = String(val);
        }
        tr.appendChild(td);
      });
      tableBody.appendChild(tr);
    });
  }

  // Pagination events
  btnPrevPage.addEventListener('click', () => {
    if (currentPage > 1) {
      currentPage--;
      loadSheetData();
    }
  });

  btnNextPage.addEventListener('click', () => {
    if (currentPage < totalPages) {
      currentPage++;
      loadSheetData();
    }
  });

  // Search input with debounce
  sheetSearchInput.addEventListener('input', (e) => {
    clearTimeout(searchTimeout);
    searchTimeout = setTimeout(() => {
      currentSearch = e.target.value.trim();
      currentPage = 1;
      loadSheetData();
    }, 350);
  });

  // Initialize summary on startup
  loadSummary();
});
