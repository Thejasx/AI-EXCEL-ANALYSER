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
      showToast('Could not fetch sample file directly.', 'info');
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

  // Upload File API Call
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

    const formData = new FormData();
    formData.append('excelFile', file);

    try {
      showLoading(true);
      const res = await fetch('/api/upload', {
        method: 'POST',
        body: formData
      });

      const result = await res.json();
      if (!res.ok || !result.success) {
        throw new Error(result.error || 'Failed to parse Excel file');
      }

      originalData = result.data || [];
      processedData = []; // Clear previous AI output

      // Update UI File Cards
      fileNameDisplay.textContent = file.name;
      fileMetaDisplay.textContent = `${originalData.length} rows • Sheet: ${result.activeSheet || 'Sheet1'}`;
      fileInfo.classList.remove('hidden');
      fileBadge.classList.remove('hidden');

      // Update Tables
      renderExcelGrid(originalTableHead, originalTableBody, originalData);
      originalBadgeCount.textContent = originalData.length;
      originalTableRowsText.textContent = `${originalData.length} spreadsheet rows`;

      // Reset Processed Table
      processedTableHead.innerHTML = `<tr><th class="p-3 text-center text-slate-400 font-normal">Execute an AI instruction to display transformed spreadsheet output</th></tr>`;
      processedTableBody.innerHTML = '';
      processedBadgeCount.textContent = '0';
      processedTableRowsText.textContent = 'Awaiting formula execution';
      summaryCard.classList.add('hidden');
      downloadExcelBtn.disabled = true;

      switchTab('original');
      updateAnalyzeButtonState();
      showToast(`Loaded ${originalData.length} rows from ${file.name}`, 'success');
    } catch (err) {
      console.error(err);
      showToast(err.message, 'error');
      resetState();
    } finally {
      showLoading(false);
    }
  }

  // Analyze Button Click Handler
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

      const response = await fetch('/api/analyze', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          data: originalData,
          prompt
        })
      });

      const result = await response.json();
      if (!response.ok || !result.success) {
        throw new Error(result.error || 'Failed to process request with Gemini API');
      }

      processedData = result.processedData || [];

      // Update Summary Card
      summaryText.textContent = result.summary || 'Processing complete.';
      summaryCard.classList.remove('hidden');

      // Render Processed Excel Grid
      renderExcelGrid(processedTableHead, processedTableBody, processedData);
      processedBadgeCount.textContent = processedData.length;
      processedTableRowsText.textContent = `${processedData.length} output rows`;

      // Enable Download button
      downloadExcelBtn.disabled = processedData.length === 0;

      // Switch to Processed Tab
      switchTab('processed');
      showToast('Spreadsheet transformation completed!', 'success');
    } catch (err) {
      console.error(err);
      showToast(err.message, 'error');
    } finally {
      showLoading(false);
    }
  });

  // Download Processed Excel Handler
  downloadExcelBtn.addEventListener('click', async () => {
    if (!processedData.length) {
      showToast('No processed data available to export.', 'error');
      return;
    }

    try {
      showLoading(true);
      const res = await fetch('/api/export', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          processedData,
          fileName: activeFileName
        })
      });

      if (!res.ok) {
        const errorJson = await res.json();
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

      showToast('Excel file downloaded successfully!', 'success');
    } catch (err) {
      console.error(err);
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

  // Excel Grid Table Renderer
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

    // Build Table Body Rows with Row Numbers
    dataArray.forEach((row, rowIndex) => {
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
          if (header.toLowerCase().includes('revenue') || header.toLowerCase().includes('price') || header.toLowerCase().includes('cost')) {
            value = '$' + Number(value).toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 2 });
          } else {
            value = Number(value).toLocaleString();
          }
        }
        td.textContent = value;
        tr.appendChild(td);
      });
      bodyElem.appendChild(tr);
    });
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

    toast.className = `pointer-events-auto flex items-center space-x-2.5 px-4 py-3 rounded-lg border ${bgColors} shadow-lg text-xs font-semibold transition duration-300 transform translate-y-2 opacity-0`;
    toast.innerHTML = `
      <i data-lucide="${icon}" class="w-4 h-4 text-excel-600 flex-shrink-0"></i>
      <span>${message}</span>
    `;

    toastContainer.appendChild(toast);
    if (window.lucide) window.lucide.createIcons();

    setTimeout(() => {
      toast.classList.remove('translate-y-2', 'opacity-0');
    }, 10);

    setTimeout(() => {
      toast.classList.add('opacity-0', 'translate-y-2');
      setTimeout(() => toast.remove(), 300);
    }, 3500);
  }
});
