document.addEventListener('DOMContentLoaded', () => {
  // DOM Elements
  const dropZone = document.getElementById('dropZone');
  const fileInput = document.getElementById('fileInput');
  const fileInfo = document.getElementById('fileInfo');
  const fileNameDisplay = document.getElementById('fileNameDisplay');
  const fileMetaDisplay = document.getElementById('fileMetaDisplay');
  const clearFileBtn = document.getElementById('clearFileBtn');
  const fileBadge = document.getElementById('fileBadge');
  
  const promptInput = document.getElementById('promptInput');
  const analyzeBtn = document.getElementById('analyzeBtn');
  const loadSampleBtn = document.getElementById('loadSampleBtn');
  const presetChips = document.querySelectorAll('.preset-chip');
  
  const summaryCard = document.getElementById('summaryCard');
  const summaryText = document.getElementById('summaryText');

  const tabOriginalBtn = document.getElementById('tabOriginalBtn');
  const tabProcessedBtn = document.getElementById('tabProcessedBtn');
  const originalTableView = document.getElementById('originalTableView');
  const processedTableView = document.getElementById('processedTableView');
  
  const originalBadgeCount = document.getElementById('originalBadgeCount');
  const processedBadgeCount = document.getElementById('processedBadgeCount');
  const originalTableRowsText = document.getElementById('originalTableRowsText');
  const processedTableRowsText = document.getElementById('processedTableRowsText');
  
  const originalTableHead = document.getElementById('originalTableHead');
  const originalTableBody = document.getElementById('originalTableBody');
  const processedTableHead = document.getElementById('processedTableHead');
  const processedTableBody = document.getElementById('processedTableBody');
  
  const downloadExcelBtn = document.getElementById('downloadExcelBtn');
  const loadingOverlay = document.getElementById('loadingOverlay');

  // Application State
  let currentFile = null;
  let activeFileName = '';
  let originalData = [];
  let processedData = [];

  // Max DOM rendering limit for fast table performance
  const RENDER_LIMIT = 250;

  // Drag & Drop File Handlers
  dropZone.addEventListener('click', () => fileInput.click());
  
  dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('border-excel-500', 'bg-emerald-50/50');
  });

  dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('border-excel-500', 'bg-emerald-50/50');
  });

  dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('border-excel-500', 'bg-emerald-50/50');
    if (e.dataTransfer.files.length) {
      handleFileUpload(e.dataTransfer.files[0]);
    }
  });

  fileInput.addEventListener('change', (e) => {
    if (e.target.files.length) {
      handleFileUpload(e.target.files[0]);
    }
  });

  clearFileBtn.addEventListener('click', resetState);

  // Load Sample Data Handler
  loadSampleBtn.addEventListener('click', async () => {
    try {
      showLoading(true);
      const sampleResponse = await fetch('/sales_data.xlsx');
      if (!sampleResponse.ok) {
        throw new Error('Sample Excel file not found on server.');
      }
      const blob = await sampleResponse.blob();
      const sampleFile = new File([blob], 'sales_data.xlsx', {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
      await handleFileUpload(sampleFile);
      showToast('Sample sales dataset loaded successfully!', 'success');
    } catch (err) {
      console.warn(err);
      showToast(err.message || 'Could not fetch sample file.', 'error');
    } finally {
      showLoading(false);
    }
  });

  // Preset chip clicks
  presetChips.forEach(chip => {
    chip.addEventListener('click', () => {
      const promptText = chip.getAttribute('data-prompt');
      promptInput.value = promptText;
      updateAnalyzeButtonState();
    });
  });

  promptInput.addEventListener('input', updateAnalyzeButtonState);

  // Tab Navigation Handlers
  tabOriginalBtn.addEventListener('click', () => switchTab('original'));
  tabProcessedBtn.addEventListener('click', () => switchTab('processed'));

  /**
   * Helper to safely parse API responses and throw human-readable errors
   */
  async function parseApiResponse(res) {
    const contentType = res.headers.get('content-type') || '';
    if (contentType.includes('application/json')) {
      const json = await res.json();
      if (!res.ok) {
        throw new Error(json.error || `Server error (${res.status})`);
      }
      return json;
    } else {
      const text = await res.text();
      const cleanText = text.replace(/<[^>]*>/g, '').trim();
      throw new Error(`Server returned error (${res.status}): ${cleanText.slice(0, 150) || res.statusText}`);
    }
  }

  /**
   * Parse Excel File client-side via SheetJS (bypasses payload limits)
   * with server fallback if client parsing fails.
   */
  async function handleFileUpload(file) {
    if (!file) return;

    const validExtensions = ['.xlsx', '.xls', '.csv'];
    const hasValidExt = validExtensions.some(ext => file.name.toLowerCase().endsWith(ext));
    if (!hasValidExt) {
      showToast('Please upload a valid .xlsx, .xls, or .csv file', 'error');
      return;
    }

    currentFile = file;
    activeFileName = file.name;

    showLoading(true);

    // Attempt Client-Side SheetJS parsing (Fast, 0 network payload)
    if (window.XLSX) {
      try {
        const arrayBuffer = await file.arrayBuffer();
        const workbook = window.XLSX.read(arrayBuffer, { type: 'array' });

        if (!workbook.SheetNames || !workbook.SheetNames.length) {
          throw new Error('The uploaded file contains no sheets.');
        }

        const activeSheet = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[activeSheet];
        originalData = window.XLSX.utils.sheet_to_json(worksheet, { defval: '' });
        processedData = [];

        onDataLoaded(file.name, activeSheet);
        showToast(`Loaded ${originalData.length.toLocaleString()} rows from ${file.name}`, 'success');
        showLoading(false);
        return;
      } catch (clientErr) {
        console.warn('Client-side SheetJS parsing failed, falling back to server route:', clientErr);
      }
    }

    // Server Fallback Route
    try {
      const formData = new FormData();
      formData.append('excelFile', file);

      const res = await fetch('/api/upload', {
        method: 'POST',
        body: formData
      });

      const result = await parseApiResponse(res);
      originalData = result.data || [];
      processedData = [];

      onDataLoaded(file.name, result.activeSheet || 'Sheet1');
      showToast(`Loaded ${originalData.length.toLocaleString()} rows from ${file.name}`, 'success');
    } catch (err) {
      console.error('File Upload Error:', err);
      showToast(err.message, 'error');
      resetState();
    } finally {
      showLoading(false);
    }
  }

  function onDataLoaded(fileName, sheetName) {
    fileNameDisplay.textContent = fileName;
    fileMetaDisplay.textContent = `${originalData.length.toLocaleString()} rows • Sheet: ${sheetName}`;
    fileInfo.classList.remove('hidden');
    fileBadge.classList.remove('hidden');

    renderExcelGrid(originalTableHead, originalTableBody, originalData);
    originalBadgeCount.textContent = originalData.length.toLocaleString();

    if (originalData.length > RENDER_LIMIT) {
      originalTableRowsText.textContent = `Showing top ${RENDER_LIMIT} of ${originalData.length.toLocaleString()} rows`;
    } else {
      originalTableRowsText.textContent = `${originalData.length.toLocaleString()} spreadsheet rows`;
    }

    processedTableHead.innerHTML = `<tr><th class="p-3 text-center text-slate-400 font-normal">Execute an AI instruction to display transformed spreadsheet output</th></tr>`;
    processedTableBody.innerHTML = '';
    processedBadgeCount.textContent = '0';
    processedTableRowsText.textContent = 'Awaiting formula execution';
    summaryCard.classList.add('hidden');
    downloadExcelBtn.disabled = true;

    switchTab('original');
    updateAnalyzeButtonState();
  }

  // Analyze Button Click Handler (Uses Lightweight ~2 KB Metadata payload)
  analyzeBtn.addEventListener('click', async () => {
    const prompt = promptInput.value.trim();
    if (!originalData.length) {
      showToast('Please upload an Excel file first.', 'error');
      return;
    }
    if (!prompt) {
      showToast('Please enter an analysis prompt or select a preset action.', 'error');
      return;
    }

    try {
      showLoading(true);

      // Lightweight metadata payload (2 KB) avoids Vercel 4.5 MB body limit
      const payload = {
        columns: Object.keys(originalData[0] || {}),
        sampleRows: originalData.slice(0, Math.min(10, originalData.length)),
        totalRows: originalData.length,
        prompt
      };

      const response = await fetch('/api/analyze', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });

      const result = await parseApiResponse(response);

      // Execute AI Code on full dataset in client browser (Instant execution, 0 payload)
      if (result.code && typeof result.code === 'string') {
        let cleanCode = result.code.trim();
        if (cleanCode.startsWith('```')) {
          cleanCode = cleanCode.replace(/```(?:javascript|js)?/gi, '').replace(/```/g, '').trim();
        }

        try {
          const transformFn = new Function('data', `return (${cleanCode})(data);`);
          processedData = transformFn(originalData);
        } catch (evalErr) {
          console.warn('Browser evaluation of Gemini code failed:', evalErr);
          processedData = result.processedData || originalData;
        }
      } else if (result.processedData) {
        processedData = result.processedData;
      } else {
        processedData = originalData;
      }

      // Update Summary Card
      summaryText.textContent = result.summary || 'Processing complete.';
      summaryCard.classList.remove('hidden');

      // Render Processed Excel Grid
      renderExcelGrid(processedTableHead, processedTableBody, processedData);
      processedBadgeCount.textContent = processedData.length.toLocaleString();
      
      if (processedData.length > RENDER_LIMIT) {
        processedTableRowsText.textContent = `Showing top ${RENDER_LIMIT} of ${processedData.length.toLocaleString()} output rows`;
      } else {
        processedTableRowsText.textContent = `${processedData.length.toLocaleString()} output rows`;
      }

      // Enable Download button
      downloadExcelBtn.disabled = processedData.length === 0;

      // Switch to Processed Tab
      switchTab('processed');
      showToast(`Transformed ${processedData.length.toLocaleString()} rows successfully!`, 'success');
    } catch (err) {
      console.error('AI Analysis Error:', err);
      showToast(err.message, 'error');
    } finally {
      showLoading(false);
    }
  });

  // Download Processed Excel Handler (Client-side SheetJS Export)
  downloadExcelBtn.addEventListener('click', async () => {
    if (!processedData.length) {
      showToast('No processed data available to export.', 'error');
      return;
    }

    try {
      showLoading(true);

      // Client-Side SheetJS Export (Instant, 0 network payload)
      if (window.XLSX) {
        const worksheet = window.XLSX.utils.json_to_sheet(processedData);
        const workbook = window.XLSX.utils.book_new();
        window.XLSX.utils.book_append_sheet(workbook, worksheet, 'Analyzed Data');

        const baseName = activeFileName ? activeFileName.replace(/\.[^/.]+$/, '') : 'Data';
        const downloadFileName = `Analyzed_${baseName}.xlsx`;

        window.XLSX.writeFile(workbook, downloadFileName);
        showToast(`Exported ${processedData.length.toLocaleString()} rows to Excel!`, 'success');
        showLoading(false);
        return;
      }

      // Server Export Fallback Route
      const res = await fetch('/api/export', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          processedData,
          fileName: activeFileName
        })
      });

      if (!res.ok) {
        const errorJson = await res.json().catch(() => ({ error: 'Export failed' }));
        throw new Error(errorJson.error || 'Failed to generate Excel file');
      }

      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Analyzed_${activeFileName || 'Data.xlsx'}`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);

      showToast(`Exported ${processedData.length.toLocaleString()} rows to Excel!`, 'success');
    } catch (err) {
      console.error('Export Error:', err);
      showToast(err.message, 'error');
    } finally {
      showLoading(false);
    }
  });

  // Helper to convert 0-index column to Excel letter (0 -> A, 1 -> B, 25 -> Z, 26 -> AA)
  function getExcelColumnLetter(index) {
    let letter = '';
    while (index >= 0) {
      letter = String.fromCharCode((index % 26) + 65) + letter;
      index = Math.floor(index / 26) - 1;
    }
    return letter;
  }

  // Excel Grid Table Renderer (with performance virtualization limit)
  function renderExcelGrid(headElem, bodyElem, dataArray) {
    headElem.innerHTML = '';
    bodyElem.innerHTML = '';

    if (!dataArray || dataArray.length === 0) {
      headElem.innerHTML = `<tr><th class="p-3 text-center text-slate-400 font-normal">No spreadsheet data available</th></tr>`;
      return;
    }

    const headers = Object.keys(dataArray[0]);

    // Build Excel Column Header Row
    const trHead = document.createElement('tr');
    
    // Top-left row index corner cell (#)
    const thCorner = document.createElement('th');
    thCorner.className = 'w-12 p-2 px-3 text-center font-mono font-bold text-slate-400 bg-slate-100 border-r border-b border-slate-300 text-[11px] select-none';
    thCorner.textContent = '#';
    trHead.appendChild(thCorner);

    headers.forEach((header, idx) => {
      const th = document.createElement('th');
      th.className = 'p-2.5 px-4 font-bold text-slate-700 uppercase tracking-wide whitespace-nowrap excel-grid-th text-xs select-none';
      const letter = getExcelColumnLetter(idx);
      th.innerHTML = `<span class="text-excel-600 font-mono text-[11px] mr-1.5">${letter}</span><span>${header}</span>`;
      trHead.appendChild(th);
    });
    headElem.appendChild(trHead);

    // Limit DOM rows rendered to prevent browser tab lag on massive files
    const renderData = dataArray.slice(0, RENDER_LIMIT);

    // Build Table Body Rows with Row Numbers
    const fragment = document.createDocumentFragment();
    renderData.forEach((row, rowIndex) => {
      const tr = document.createElement('tr');
      tr.className = 'hover:bg-emerald-50/40 transition duration-150 odd:bg-white even:bg-slate-50/70';

      // Left Row Number Cell (1, 2, 3...)
      const tdRowNum = document.createElement('td');
      tdRowNum.className = 'w-12 p-2 px-3 text-center font-mono text-slate-500 font-semibold excel-row-header text-[11px] select-none';
      tdRowNum.textContent = rowIndex + 1;
      tr.appendChild(tdRowNum);

      headers.forEach(header => {
        const td = document.createElement('td');
        td.className = 'p-2.5 px-4 text-slate-800 font-normal whitespace-nowrap excel-grid-td text-xs';
        let value = row[header];
        if (value === null || value === undefined) value = '';
        if (typeof value === 'number') {
          if (header.toLowerCase().includes('revenue') || header.toLowerCase().includes('price') || header.toLowerCase().includes('cost') || header.toLowerCase().includes('profit') || header.toLowerCase().includes('tax')) {
            value = '$' + Number(value).toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 2 });
          } else {
            value = Number(value).toLocaleString();
          }
        }
        td.textContent = value;
        tr.appendChild(td);
      });
      fragment.appendChild(tr);
    });
    bodyElem.appendChild(fragment);

    // Add informative footer row if dataset exceeds render limit
    if (dataArray.length > RENDER_LIMIT) {
      const trInfo = document.createElement('tr');
      trInfo.className = 'bg-amber-50/70 border-t border-amber-200';
      const tdInfo = document.createElement('td');
      tdInfo.colSpan = headers.length + 1;
      tdInfo.className = 'p-2.5 px-4 text-amber-800 text-center text-xs font-semibold';
      tdInfo.textContent = `⚡ Displaying top ${RENDER_LIMIT} rows out of ${dataArray.length.toLocaleString()} total rows. All ${dataArray.length.toLocaleString()} rows will be exported when downloading the Excel file.`;
      trInfo.appendChild(tdInfo);
      bodyElem.appendChild(trInfo);
    }
  }

  // Switch Tab View (Excel Theme)
  function switchTab(tab) {
    if (tab === 'original') {
      originalTableView.classList.remove('hidden');
      processedTableView.classList.add('hidden');
      
      tabOriginalBtn.className = 'px-4 py-2 rounded-t-lg text-xs font-bold bg-white text-excel-700 border-t-2 border-x border-excel-500 shadow-sm transition flex items-center space-x-2';
      tabProcessedBtn.className = 'px-4 py-2 rounded-t-lg text-xs font-semibold text-slate-600 hover:text-slate-900 border border-transparent transition flex items-center space-x-2';
    } else {
      originalTableView.classList.add('hidden');
      processedTableView.classList.remove('hidden');

      tabOriginalBtn.className = 'px-4 py-2 rounded-t-lg text-xs font-semibold text-slate-600 hover:text-slate-900 border border-transparent transition flex items-center space-x-2';
      tabProcessedBtn.className = 'px-4 py-2 rounded-t-lg text-xs font-bold bg-white text-excel-700 border-t-2 border-x border-excel-500 shadow-sm transition flex items-center space-x-2';
    }
  }

  // State Resets
  function resetState() {
    currentFile = null;
    activeFileName = '';
    originalData = [];
    processedData = [];
    fileInput.value = '';

    fileInfo.classList.add('hidden');
    fileBadge.classList.add('hidden');
    summaryCard.classList.add('hidden');

    originalTableHead.innerHTML = `<tr><th class="p-3 text-center text-slate-400 font-normal">Upload an Excel file to display original rows</th></tr>`;
    originalTableBody.innerHTML = '';
    processedTableHead.innerHTML = `<tr><th class="p-3 text-center text-slate-400 font-normal">Execute an AI instruction to display transformed spreadsheet output</th></tr>`;
    processedTableBody.innerHTML = '';

    originalBadgeCount.textContent = '0';
    processedBadgeCount.textContent = '0';
    originalTableRowsText.textContent = 'No workbook loaded';
    processedTableRowsText.textContent = 'Awaiting formula execution';

    downloadExcelBtn.disabled = true;
    updateAnalyzeButtonState();
    switchTab('original');
  }

  function updateAnalyzeButtonState() {
    const hasData = originalData.length > 0;
    const hasPrompt = promptInput.value.trim().length > 0;
    analyzeBtn.disabled = !(hasData && hasPrompt);
  }

  function showLoading(show) {
    if (show) {
      loadingOverlay.classList.remove('hidden');
    } else {
      loadingOverlay.classList.add('hidden');
    }
  }

  function showToast(message, type = 'info') {
    const toastContainer = document.getElementById('toastContainer');
    const toast = document.createElement('div');
    
    let bgColors = 'bg-white text-slate-800 border-slate-300';
    let icon = 'info';
    
    if (type === 'success') {
      bgColors = 'bg-white text-excel-800 border-emerald-400 shadow-emerald-500/10';
      icon = 'check-circle-2';
    } else if (type === 'error') {
      bgColors = 'bg-white text-rose-800 border-rose-300 shadow-rose-500/10';
      icon = 'alert-circle';
    }

    toast.className = `pointer-events-auto flex items-center space-x-2.5 px-4 py-3 rounded-lg border ${bgColors} shadow-lg text-xs font-semibold transition duration-300 transform translate-y-2 opacity-0 max-w-md`;
    toast.innerHTML = `
      <i data-lucide="${icon}" class="w-4 h-4 text-excel-600 flex-shrink-0"></i>
      <span class="break-words">${message}</span>
    `;

    toastContainer.appendChild(toast);
    if (window.lucide) window.lucide.createIcons();

    setTimeout(() => {
      toast.classList.remove('translate-y-2', 'opacity-0');
    }, 10);

    setTimeout(() => {
      toast.classList.add('opacity-0', 'translate-y-2');
      setTimeout(() => toast.remove(), 300);
    }, 4500);
  }
});
