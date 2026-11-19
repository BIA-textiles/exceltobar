/**
 * BarcodeStudio - Professional Excel to Barcode Application
 * Supports .xlsx, .xls, .csv parsing, live barcode rendering, customizable templates, and high-DPI print/export.
 */

(function () {
  'use strict';

  // State Management
  const state = {
    workbook: null,
    fileName: '',
    sheetNames: [],
    currentSheet: '',
    columns: [],
    selectedColumnIndex: 0,
    columnName: 'Column 1',
    rawItems: [], // [{ index: 1, value: "ABC-123", raw: "ABC-123" }]
    validCount: 0,
    errorCount: 0,
    settings: {
      format: 'CODE128',
      height: 85,       // ↑ taller bars = easier to scan
      width: 2.5,       // ↑ wider bars = more reliable scanning
      showColName: true,
      showText: true,
      showCounter: true,
      trim: true,
      layout: 'grid-3'
    }
  };

  // DOM Elements
  const elements = {
    dropZone: document.getElementById('dropZone'),
    fileInput: document.getElementById('fileInput'),
    fileMetaBar: document.getElementById('fileMetaBar'),
    metaFileName: document.getElementById('metaFileName'),
    metaDetails: document.getElementById('metaDetails'),
    sheetSelect: document.getElementById('sheetSelect'),
    columnSelect: document.getElementById('columnSelect'),
    btnClearFile: document.getElementById('btnClearFile'),

    barcodeFormat: document.getElementById('barcodeFormat'),
    printLayout: document.getElementById('printLayout'),
    barcodeHeight: document.getElementById('barcodeHeight'),
    valBarcodeHeight: document.getElementById('valBarcodeHeight'),
    barcodeWidth: document.getElementById('barcodeWidth'),
    valBarcodeWidth: document.getElementById('valBarcodeWidth'),

    chkShowColName: document.getElementById('chkShowColName'),
    chkShowText: document.getElementById('chkShowText'),
    chkShowCounter: document.getElementById('chkShowCounter'),
    chkTrim: document.getElementById('chkTrim'),

    resultsToolbar: document.getElementById('resultsToolbar'),
    statTotalCount: document.getElementById('statTotalCount'),
    statColumnName: document.getElementById('statColumnName'),
    statValidCount: document.getElementById('statValidCount'),
    statErrorBadge: document.getElementById('statErrorBadge'),
    statErrorCount: document.getElementById('statErrorCount'),
    searchInput: document.getElementById('searchInput'),
    noSearchResults: document.getElementById('noSearchResults'),

    barcodeContainer: document.getElementById('barcodeContainer'),
    emptyState: document.getElementById('emptyState'),

    btnSampleData: document.getElementById('btnSampleData'),
    btnDownloadTemplate: document.getElementById('btnDownloadTemplate'),
    btnPrintAll: document.getElementById('btnPrintAll'),
    btnToolbarPrint: document.getElementById('btnToolbarPrint'),
    btnDownloadZip: document.getElementById('btnDownloadZip'),

    printDocumentHeader: document.getElementById('printDocumentHeader'),
    printDocTitle: document.getElementById('printDocTitle'),
    printDocSub: document.getElementById('printDocSub'),
    printDate: document.getElementById('printDate'),
    toastContainer: document.getElementById('toastContainer')
  };

  // Initialize
  function init() {
    bindEvents();
    syncLayoutClass();
  }

  // Event Listeners
  function bindEvents() {
    // File Drag and Drop
    elements.dropZone.addEventListener('dragover', (e) => {
      e.preventDefault();
      e.stopPropagation();
      elements.dropZone.classList.add('dragover');
    });

    elements.dropZone.addEventListener('dragleave', (e) => {
      e.preventDefault();
      e.stopPropagation();
      elements.dropZone.classList.remove('dragover');
    });

    elements.dropZone.addEventListener('drop', (e) => {
      e.preventDefault();
      e.stopPropagation();
      elements.dropZone.classList.remove('dragover');
      if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
        handleFile(e.dataTransfer.files[0]);
      }
    });

    // File Input change
    elements.fileInput.addEventListener('change', (e) => {
      if (e.target.files && e.target.files.length > 0) {
        handleFile(e.target.files[0]);
      }
    });

    // Clear file
    elements.btnClearFile.addEventListener('click', resetData);

    // Sheet selection change
    elements.sheetSelect.addEventListener('change', (e) => {
      state.currentSheet = e.target.value;
      extractColumnsAndData();
    });

    // Column selection change
    elements.columnSelect.addEventListener('change', (e) => {
      state.selectedColumnIndex = parseInt(e.target.value, 10);
      extractColumnData();
    });

    // Settings adjustments
    elements.barcodeFormat.addEventListener('change', (e) => {
      state.settings.format = e.target.value;
      renderBarcodes();
    });

    elements.printLayout.addEventListener('change', (e) => {
      state.settings.layout = e.target.value;
      syncLayoutClass();
    });

    elements.barcodeHeight.addEventListener('input', (e) => {
      state.settings.height = parseInt(e.target.value, 10);
      elements.valBarcodeHeight.textContent = state.settings.height + 'px';
      renderBarcodes();
    });

    elements.barcodeWidth.addEventListener('input', (e) => {
      state.settings.width = parseFloat(e.target.value);
      elements.valBarcodeWidth.textContent = state.settings.width + 'px';
      renderBarcodes();
    });

    elements.chkShowColName.addEventListener('change', (e) => {
      state.settings.showColName = e.target.checked;
      toggleCardHeaders();
    });

    elements.chkShowText.addEventListener('change', (e) => {
      state.settings.showText = e.target.checked;
      renderBarcodes();
    });

    elements.chkShowCounter.addEventListener('change', (e) => {
      state.settings.showCounter = e.target.checked;
      toggleCardCounters();
    });

    elements.chkTrim.addEventListener('change', (e) => {
      state.settings.trim = e.target.checked;
      extractColumnData();
    });

    // Search filter
    elements.searchInput.addEventListener('input', (e) => {
      filterBarcodes(e.target.value.trim().toLowerCase());
    });

    // Header & Toolbar actions
    elements.btnSampleData.addEventListener('click', loadSampleData);
    elements.btnDownloadTemplate.addEventListener('click', downloadExcelTemplate);
    elements.btnPrintAll.addEventListener('click', printLabels);
    elements.btnToolbarPrint.addEventListener('click', printLabels);
    elements.btnDownloadZip.addEventListener('click', downloadAllAsZip);
  }

  // Handle incoming file
  function handleFile(file) {
    if (!file) return;

    const extension = file.name.split('.').pop().toLowerCase();
    if (!['xlsx', 'xls', 'csv'].includes(extension)) {
      showToast('Please upload a valid Excel (.xlsx, .xls) or CSV (.csv) file.', 'error');
      return;
    }

    state.fileName = file.name;
    const reader = new FileReader();

    reader.onload = function (e) {
      try {
        const data = e.target.result;
        // Check if SheetJS XLSX is available
        if (typeof XLSX !== 'undefined') {
          const workbook = XLSX.read(data, { type: 'binary', cellDates: true });
          state.workbook = workbook;
          state.sheetNames = workbook.SheetNames;
          state.currentSheet = workbook.SheetNames[0];

          // Populate sheet dropdown
          elements.sheetSelect.innerHTML = '';
          state.sheetNames.forEach((sheetName) => {
            const opt = document.createElement('option');
            opt.value = sheetName;
            opt.textContent = sheetName;
            elements.sheetSelect.appendChild(opt);
          });

          extractColumnsAndData();
          showToast(`Loaded "${file.name}" successfully!`, 'success');
        } else {
          // Fallback parser for CSV/Text if offline without CDN
          parseDelimitedText(data, file.name);
        }
      } catch (err) {
        console.error('File parsing error:', err);
        showToast('Error reading spreadsheet file: ' + err.message, 'error');
      }
    };

    if (extension === 'csv' && typeof XLSX === 'undefined') {
      reader.readAsText(file);
    } else {
      reader.readAsBinaryString(file);
    }
  }

  // Extract columns from current SheetJS worksheet
  function extractColumnsAndData() {
    if (!state.workbook || !state.currentSheet) return;

    const worksheet = state.workbook.Sheets[state.currentSheet];
    // Convert sheet to array of arrays (header in row 0)
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, blankrows: false, defval: '' });

    if (!jsonData || jsonData.length === 0) {
      showToast('The selected sheet is empty.', 'error');
      resetData();
      return;
    }

    // Determine headers from row 0
    const headerRow = jsonData[0] || [];
    state.columns = [];
    const maxCols = Math.max(...jsonData.map(row => (Array.isArray(row) ? row.length : 0)));

    for (let c = 0; c < maxCols; c++) {
      let colName = (headerRow[c] !== undefined && headerRow[c] !== null && String(headerRow[c]).trim() !== '') 
        ? String(headerRow[c]).trim() 
        : `Column ${c + 1}`;
      state.columns.push(colName);
    }

    // Populate Column Dropdown (Column 1 is index 0 by default)
    elements.columnSelect.innerHTML = '';
    state.columns.forEach((name, index) => {
      const opt = document.createElement('option');
      opt.value = index;
      opt.textContent = `${index === 0 ? '★ ' : ''}Col ${index + 1}: ${name}`;
      elements.columnSelect.appendChild(opt);
    });

    state.selectedColumnIndex = 0;
    elements.columnSelect.value = 0;

    // Store raw rows for fast column switching
    state.sheetRows = jsonData;

    // Update metadata badge
    elements.fileMetaBar.style.display = 'flex';
    elements.metaFileName.textContent = state.fileName;
    elements.metaDetails.textContent = `${state.currentSheet} • ${jsonData.length - 1} data rows`;

    extractColumnData();
  }

  // Extract data from the chosen column
  function extractColumnData() {
    if (!state.sheetRows || state.sheetRows.length <= 1) return;

    const rows = state.sheetRows;
    const colIdx = state.selectedColumnIndex;
    state.columnName = state.columns[colIdx] || `Column ${colIdx + 1}`;

    const items = [];
    // Start from row 1 (row 0 is header)
    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      let val = row[colIdx];
      if (val === undefined || val === null) continue;
      
      let strVal = String(val);
      if (state.settings.trim) {
        strVal = strVal.trim();
      }

      if (strVal !== '') {
        items.push({
          index: items.length + 1,
          value: strVal,
          colName: state.columnName
        });
      }
    }

    state.rawItems = items;
    renderBarcodes();
  }

  // Fallback CSV Parser if offline and SheetJS is unavailable
  function parseDelimitedText(text, filename) {
    const lines = text.split(/\r\n|\n/).filter(line => line.trim() !== '');
    if (lines.length === 0) {
      showToast('The CSV file is empty.', 'error');
      return;
    }

    const rows = lines.map(line => {
      // Simple split by comma or tab
      const delimiter = line.includes('\t') ? '\t' : ',';
      return line.split(delimiter).map(cell => cell.replace(/^"(.*)"$/, '$1').trim());
    });

    state.sheetRows = rows;
    state.sheetNames = ['Sheet 1'];
    state.currentSheet = 'Sheet 1';
    state.columns = rows[0].map((col, idx) => col || `Column ${idx + 1}`);

    elements.columnSelect.innerHTML = '';
    state.columns.forEach((name, idx) => {
      const opt = document.createElement('option');
      opt.value = idx;
      opt.textContent = `${idx === 0 ? '★ ' : ''}Col ${idx + 1}: ${name}`;
      elements.columnSelect.appendChild(opt);
    });

    elements.sheetSelect.innerHTML = '<option value="Sheet 1">Sheet 1</option>';
    elements.fileMetaBar.style.display = 'flex';
    elements.metaFileName.textContent = filename;
    elements.metaDetails.textContent = `${rows.length - 1} rows parsed`;

    state.selectedColumnIndex = 0;
    extractColumnData();
  }

  // Render Barcode Cards into DOM
  function renderBarcodes() {
    if (!state.rawItems || state.rawItems.length === 0) {
      elements.emptyState.style.display = 'block';
      elements.resultsToolbar.style.display = 'none';
      elements.barcodeContainer.innerHTML = '';
      elements.btnPrintAll.disabled = true;
      return;
    }

    elements.emptyState.style.display = 'none';
    elements.resultsToolbar.style.display = 'flex';
    elements.btnPrintAll.disabled = false;

    // Reset stats
    state.validCount = 0;
    state.errorCount = 0;

    elements.statTotalCount.textContent = state.rawItems.length;
    elements.statColumnName.textContent = state.columnName;

    const fragment = document.createDocumentFragment();

    state.rawItems.forEach((item, i) => {
      const card = document.createElement('div');
      card.className = 'barcode-card';
      card.dataset.value = item.value.toLowerCase();
      card.dataset.index = item.index;

      // Card Header: Column Name + Index
      const headerBar = document.createElement('div');
      headerBar.className = 'card-header-bar';
      headerBar.style.display = state.settings.showColName || state.settings.showCounter ? 'flex' : 'none';

      const colTag = document.createElement('span');
      colTag.className = 'col-header-tag';
      colTag.textContent = item.colName;
      colTag.style.display = state.settings.showColName ? 'inline-flex' : 'none';

      const indexTag = document.createElement('span');
      indexTag.className = 'card-index';
      indexTag.textContent = `#${item.index}`;
      indexTag.style.display = state.settings.showCounter ? 'inline-block' : 'none';

      headerBar.appendChild(colTag);
      headerBar.appendChild(indexTag);
      card.appendChild(headerBar);

      // Barcode Render Box — CANVAS based (3× internal resolution for print/scanner quality)
      const renderBox = document.createElement('div');
      renderBox.className = 'barcode-render-box';
      card.appendChild(renderBox);

      // Render Barcode using JsBarcode (canvas mode) or offline SVG fallback
      let renderSuccess = false;
      let errorMsg = '';

      // 3× scale = effective ~288 DPI when printed from 96 DPI screen
      const RENDER_SCALE = 3;
      let barcodeElem = null; // canvas OR svg, used by download buttons

      if (typeof JsBarcode !== 'undefined') {
        // ── Canvas renderer: crisp raster at 3× for high-DPI print/scan output
        const canvas = document.createElement('canvas');
        canvas.id = `barcode-${item.index}`;
        canvas.setAttribute('role', 'img');
        canvas.setAttribute('aria-label', `Barcode for ${item.value}`);
        renderBox.appendChild(canvas);
        barcodeElem = canvas;

        try {
          JsBarcode(canvas, item.value, {
            format:       state.settings.format,
            width:        Math.round(state.settings.width * RENDER_SCALE),
            height:       Math.round(state.settings.height * RENDER_SCALE),
            displayValue: state.settings.showText,
            fontSize:     Math.round(13 * RENDER_SCALE),
            font:         'JetBrains Mono, monospace',
            fontOptions:  'bold',
            textMargin:   Math.round(4 * RENDER_SCALE),
            margin:       Math.round(10 * RENDER_SCALE),
            marginLeft:   Math.round(16 * RENDER_SCALE),
            marginRight:  Math.round(16 * RENDER_SCALE),
            background:   '#ffffff',
            lineColor:    '#000000',
            valid: function (v) { renderSuccess = v; }
          });
          renderSuccess = true;

          // CSS-scale the canvas DOWN so it displays at 1× but stores 3× data for print
          if (canvas.width > 0) {
            const dispW = Math.round(canvas.width / RENDER_SCALE);
            const dispH = Math.round(canvas.height / RENDER_SCALE);
            canvas.style.cssText = `
              width: 100%;
              max-width: ${dispW}px;
              height: auto;
              display: block;
              margin: 0 auto;
              image-rendering: crisp-edges;
              image-rendering: -webkit-optimize-contrast;
            `;
          }
        } catch (e) {
          renderSuccess = false;
          errorMsg = e.message || 'Invalid format or characters';
        }

      } else {
        // ── Offline SVG fallback (no CDN)
        const svgElem = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
        svgElem.id = `barcode-${item.index}`;
        svgElem.setAttribute('shape-rendering', 'crispEdges'); // critical for scanner quality
        renderBox.appendChild(svgElem);
        barcodeElem = svgElem;

        renderSuccess = generateCode128Svg(svgElem, item.value, {
          height:   state.settings.height,
          width:    state.settings.width,
          showText: state.settings.showText
        });
        if (!renderSuccess) errorMsg = 'Encoding error (working offline)';
      }

      if (!renderSuccess) {
        state.errorCount++;
        card.classList.add('card-error');
        renderBox.innerHTML = `
          <div class="card-error-msg">
            ⚠️ Cannot encode in <strong>${state.settings.format}</strong>.<br>
            <small style="opacity:0.8;">${errorMsg || 'Check character constraints for this format'}</small><br>
            <small style="opacity:0.65;">Try switching to <strong>CODE 128</strong> (supports all alphanumeric + symbols)</small>
          </div>
        `;
      } else {
        state.validCount++;
      }

      // Value Text Row below (if text toggle is on and JsBarcode didn't embed it, or as clean text)
      if (!state.settings.showText) {
        const valRow = document.createElement('div');
        valRow.className = 'barcode-value-row';
        valRow.innerHTML = `<span class="barcode-value-text">${escapeHtml(item.value)}</span>`;
        card.appendChild(valRow);
      }

      // Card Action Buttons (Hover bar on screen)
      const actionsBar = document.createElement('div');
      actionsBar.className = 'card-actions-bar no-print';

      // Copy Value Button
      const btnCopy = document.createElement('button');
      btnCopy.className = 'card-act-btn';
      btnCopy.title = 'Copy value';
      btnCopy.innerHTML = `
        <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect>
          <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path>
        </svg> Copy
      `;
      btnCopy.addEventListener('click', () => copyToClipboard(item.value, btnCopy));

      // Download Single PNG Button
      const btnDl = document.createElement('button');
      btnDl.className = 'card-act-btn';
      btnDl.title = 'Download Barcode PNG';
      btnDl.innerHTML = `
        <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
          <polyline points="7 10 12 15 17 10"></polyline>
          <line x1="12" y1="15" x2="12" y2="3"></line>
        </svg> PNG
      `;
      btnDl.addEventListener('click', () => downloadSingleBarcode(barcodeElem, item.value));

      actionsBar.appendChild(btnCopy);
      actionsBar.appendChild(btnDl);
      card.appendChild(actionsBar);

      fragment.appendChild(card);
    });

    elements.barcodeContainer.innerHTML = '';
    elements.barcodeContainer.appendChild(fragment);

    // Update Counter Badges
    elements.statValidCount.textContent = state.validCount;
    elements.statErrorCount.textContent = state.errorCount;
    elements.statErrorBadge.style.display = state.errorCount > 0 ? 'inline-flex' : 'none';

    // Apply layout class
    syncLayoutClass();
  }

  // Toggle Column Name Tag on cards
  function toggleCardHeaders() {
    const colTags = elements.barcodeContainer.querySelectorAll('.col-header-tag');
    colTags.forEach(el => {
      el.style.display = state.settings.showColName ? 'inline-flex' : 'none';
    });
    const headerBars = elements.barcodeContainer.querySelectorAll('.card-header-bar');
    headerBars.forEach(el => {
      el.style.display = state.settings.showColName || state.settings.showCounter ? 'flex' : 'none';
    });
  }

  // Toggle Index Tag on cards
  function toggleCardCounters() {
    const counterTags = elements.barcodeContainer.querySelectorAll('.card-index');
    counterTags.forEach(el => {
      el.style.display = state.settings.showCounter ? 'inline-block' : 'none';
    });
    const headerBars = elements.barcodeContainer.querySelectorAll('.card-header-bar');
    headerBars.forEach(el => {
      el.style.display = state.settings.showColName || state.settings.showCounter ? 'flex' : 'none';
    });
  }

  // Sync Layout Class on container
  function syncLayoutClass() {
    elements.barcodeContainer.className = `barcode-container layout-${state.settings.layout}`;
  }

  // Filter Barcodes by Search Keyword
  function filterBarcodes(query) {
    const cards = elements.barcodeContainer.querySelectorAll('.barcode-card');
    let visibleCount = 0;

    cards.forEach(card => {
      const val = card.dataset.value;
      if (!query || val.includes(query)) {
        card.style.display = '';
        visibleCount++;
      } else {
        card.style.display = 'none';
      }
    });

    elements.noSearchResults.style.display = (visibleCount === 0 && cards.length > 0) ? 'block' : 'none';
  }

  // Copy text to clipboard
  function copyToClipboard(text, btn) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(() => {
        showCopyFeedback(btn);
      }).catch(() => {
        fallbackCopy(text, btn);
      });
    } else {
      fallbackCopy(text, btn);
    }
  }

  function fallbackCopy(text, btn) {
    const textarea = document.createElement('textarea');
    textarea.value = text;
    document.body.appendChild(textarea);
    textarea.select();
    try {
      document.execCommand('copy');
      showCopyFeedback(btn);
    } catch (e) {
      showToast('Could not copy to clipboard', 'error');
    }
    document.body.removeChild(textarea);
  }

  function showCopyFeedback(btn) {
    const origHtml = btn.innerHTML;
    btn.innerHTML = '✓ Copied';
    btn.classList.add('btn-copied');
    setTimeout(() => {
      btn.innerHTML = origHtml;
      btn.classList.remove('btn-copied');
    }, 1500);
  }

  // Convert SVG to PNG Data URL
  function svgToPngDataUrl(svgElement) {
    return new Promise((resolve, reject) => {
      try {
        const svgString = new XMLSerializer().serializeToString(svgElement);
        const svgBlob = new Blob([svgString], { type: 'image/svg+xml;charset=utf-8' });
        const URL = window.URL || window.webkitURL || window;
        const blobURL = URL.createObjectURL(svgBlob);
        
        const img = new Image();
        img.onload = function () {
          const canvas = document.createElement('canvas');
          // High resolution scale factor (2x for crisp print quality)
          const scale = 2;
          const bbox = svgElement.getBoundingClientRect();
          const width = (bbox.width || 200) * scale;
          const height = (bbox.height || 80) * scale;

          canvas.width = width;
          canvas.height = height;

          const ctx = canvas.getContext('2d');
          ctx.fillStyle = '#ffffff';
          ctx.fillRect(0, 0, width, height);
          ctx.drawImage(img, 0, 0, width, height);

          URL.revokeObjectURL(blobURL);
          resolve(canvas.toDataURL('image/png'));
        };
        img.onerror = function (e) {
          URL.revokeObjectURL(blobURL);
          reject(e);
        };
        img.src = blobURL;
      } catch (err) {
        reject(err);
      }
    });
  }

  // Download single barcode as PNG
  function downloadSingleBarcode(elem, filename) {
    let pngPromise;

    if (elem && elem.tagName === 'CANVAS') {
      // Canvas already stores 3× high-res pixel data — use directly
      pngPromise = Promise.resolve(elem.toDataURL('image/png'));
    } else if (elem && elem.tagName === 'svg') {
      // SVG fallback: convert to rasterized PNG
      pngPromise = svgToPngDataUrl(elem);
    } else {
      showToast('No barcode element found to download.', 'error');
      return;
    }

    pngPromise.then(pngDataUrl => {
      const a = document.createElement('a');
      a.download = `barcode_${sanitizeFilename(filename)}.png`;
      a.href = pngDataUrl;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      showToast(`Downloaded: ${filename}.png`, 'success');
    }).catch(err => {
      console.error('Download error:', err);
      showToast('Failed to export image: ' + err.message, 'error');
    });
  }

  // Download All as ZIP archive
  async function downloadAllAsZip() {
    if (typeof JSZip === 'undefined') {
      showToast('JSZip library is loading or blocked. Printing or individual download is available.', 'warning');
      return;
    }

    const cards = elements.barcodeContainer.querySelectorAll('.barcode-card:not(.card-error)');
    if (cards.length === 0) {
      showToast('No valid barcodes to export.', 'warning');
      return;
    }

    showToast(`Preparing ZIP with ${cards.length} barcode images...`, 'info');

    const zip = new JSZip();
    const folder = zip.folder('barcodes');
    let processed = 0;

    for (let card of cards) {
      const elem = card.querySelector('canvas') || card.querySelector('svg');
      const val = card.dataset.value;
      const index = card.dataset.index;
      if (elem) {
        try {
          let dataUrl;
          if (elem.tagName === 'CANVAS') {
            // Canvas: use toDataURL directly — already high-resolution
            dataUrl = elem.toDataURL('image/png');
          } else {
            // SVG fallback: rasterize
            dataUrl = await svgToPngDataUrl(elem);
          }
          const base64Data = dataUrl.replace(/^data:image\/png;base64,/, '');
          const filename = `${String(index).padStart(3, '0')}_${sanitizeFilename(val)}.png`;
          folder.file(filename, base64Data, { base64: true });
          processed++;
        } catch (e) {
          console.warn('Skipping card in zip:', val, e);
        }
      }
    }

    zip.generateAsync({ type: 'blob' }).then((content) => {
      const a = document.createElement('a');
      a.href = URL.createObjectURL(content);
      a.download = `barcodes_${state.columnName.replace(/\s+/g, '_')}.zip`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      showToast(`ZIP archive with ${processed} barcodes generated!`, 'success');
    }).catch(err => {
      showToast('Failed to generate ZIP: ' + err.message, 'error');
    });
  }

  // Print Labels Sheet
  function printLabels() {
    if (state.rawItems.length === 0) {
      showToast('No barcodes available to print.', 'warning');
      return;
    }

    // Configure print document header details
    elements.printDocTitle.textContent = `BIA Textiles — Barcode Sheet | Column: ${state.columnName}`;
    elements.printDocSub.textContent = `Total Labels: ${state.validCount} | File: ${state.fileName || 'Sample Data'}`;
    const now = new Date();
    elements.printDate.textContent = now.toLocaleDateString() + ' ' + now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });

    window.print();
  }

  // Load Demo / Sample Data
  function loadSampleData() {
    const demoItems = [
      'PROD-88219-X',
      'SKU-99401-ALPHA',
      'INV-55104-BETA',
      'PART-10294-ENG',
      'MOD-77312-REV3',
      'TAG-44210-BLUE',
      'LOG-66890-EAST',
      'BOX-11204-PALLET',
      'SER-33821-QR7',
      'COMP-99120-MAX',
      'TOOL-22419-HEX',
      'KIT-88710-PRO'
    ];

    state.fileName = 'Sample_Products.xlsx';
    state.sheetNames = ['Inventory'];
    state.currentSheet = 'Inventory';
    state.columns = ['Product_SKU', 'Category', 'Warehouse_Location'];
    state.selectedColumnIndex = 0;
    state.columnName = 'Product_SKU';

    // Populate Sheet & Column selectors
    elements.sheetSelect.innerHTML = '<option value="Inventory">Inventory</option>';
    elements.columnSelect.innerHTML = '';
    state.columns.forEach((name, idx) => {
      const opt = document.createElement('option');
      opt.value = idx;
      opt.textContent = `${idx === 0 ? '★ ' : ''}Col ${idx + 1}: ${name}`;
      elements.columnSelect.appendChild(opt);
    });

    elements.fileMetaBar.style.display = 'flex';
    elements.metaFileName.textContent = state.fileName;
    elements.metaDetails.textContent = `Sample Sheet • ${demoItems.length} items`;

    state.rawItems = demoItems.map((val, i) => ({
      index: i + 1,
      value: val,
      colName: state.columnName
    }));

    renderBarcodes();
    showToast('Demo data loaded successfully!', 'success');
  }

  // Download Sample Excel Template
  function downloadExcelTemplate() {
    const sampleHeaders = ['Product_SKU', 'Item_Description', 'Quantity'];
    const sampleRows = [
      ['WIDGET-001-A', 'Heavy Duty Steel Bracket', '25'],
      ['WIDGET-002-B', 'Aluminum Mounting Plate', '40'],
      ['CABLE-USB-C', 'Braided Type-C Fast Cable', '100'],
      ['SENSOR-TEMP-V2', 'Digital Thermal Sensor', '15'],
      ['BATTERY-LI-2000', 'Rechargeable 2000mAh Battery', '60']
    ];

    if (typeof XLSX !== 'undefined') {
      const wb = XLSX.utils.book_new();
      const wsData = [sampleHeaders, ...sampleRows];
      const ws = XLSX.utils.aoa_to_sheet(wsData);

      // Auto column widths
      ws['!cols'] = [{ wch: 20 }, { wch: 30 }, { wch: 12 }];

      XLSX.utils.book_append_sheet(wb, ws, 'Inventory');
      XLSX.writeFile(wb, 'Barcode_Import_Template.xlsx');
      showToast('Downloaded "Barcode_Import_Template.xlsx"', 'success');
    } else {
      // Fallback CSV template
      let csvContent = sampleHeaders.join(',') + '\n';
      sampleRows.forEach(row => {
        csvContent += row.join(',') + '\n';
      });

      const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = 'Barcode_Import_Template.csv';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      showToast('Downloaded "Barcode_Import_Template.csv"', 'success');
    }
  }

  // Reset Application Data
  function resetData() {
    state.workbook = null;
    state.fileName = '';
    state.sheetRows = null;
    state.rawItems = [];
    state.columns = [];
    state.columnName = 'Column 1';

    elements.fileInput.value = '';
    elements.fileMetaBar.style.display = 'none';
    elements.barcodeContainer.innerHTML = '';
    elements.resultsToolbar.style.display = 'none';
    elements.emptyState.style.display = 'block';
    elements.btnPrintAll.disabled = true;
    elements.searchInput.value = '';
    elements.noSearchResults.style.display = 'none';

    showToast('Cleared all data.', 'info');
  }

  // Toast Notification System
  function showToast(message, type = 'info') {
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;

    let icon = 'ℹ️';
    if (type === 'success') icon = '✅';
    if (type === 'error') icon = '❌';
    if (type === 'warning') icon = '⚠️';

    toast.innerHTML = `<span>${icon}</span><span>${escapeHtml(message)}</span>`;
    elements.toastContainer.appendChild(toast);

    setTimeout(() => {
      toast.style.transition = 'opacity 0.3s, transform 0.3s';
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(10px)';
      setTimeout(() => {
        if (toast.parentNode) toast.parentNode.removeChild(toast);
      }, 300);
    }, 3500);
  }

  // Sanitize file names
  function sanitizeFilename(name) {
    return String(name).replace(/[^a-zA-Z0-9_\-\.]/g, '_').substring(0, 32);
  }

  // Helper escape HTML
  function escapeHtml(str) {
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  /**
   * Built-in Offline Code128 SVG Generator Fallback
   * Ensures the app generates valid barcodes even if CDNs are blocked or offline!
   */
  function generateCode128Svg(svg, text, opts) {
    // Basic Code 128B pattern table for ASCII 32 to 126
    const patterns = [
      '212222', '222122', '222221', '121223', '121322', '131222', '122213', '122312', '132212', '221213',
      '221312', '231212', '112232', '122132', '122231', '113222', '123122', '123221', '223211', '221132',
      '221231', '213212', '223112', '312131', '311222', '321122', '321221', '312212', '322112', '322211',
      '212123', '212321', '232121', '111323', '131123', '131321', '112313', '132113', '132311', '211313',
      '231113', '231311', '112133', '112331', '132131', '113123', '113321', '133121', '313121', '211331',
      '231131', '213113', '213311', '213131', '311123', '311321', '331121', '312113', '312311', '332111',
      '314111', '221411', '431111', '111224', '111422', '121124', '121421', '141122', '141221', '112214',
      '112412', '122114', '122411', '142112', '142211', '241211', '221114', '413111', '241112', '134111',
      '111242', '121142', '121241', '114212', '124112', '124211', '411212', '421112', '421211', '212141',
      '214121', '412121', '111143', '111341', '131141', '114113', '114311', '411113', '411311', '113141',
      '114131', '311141', '411131', '211412', '211214', '211232', '2331112'
    ];

    const START_B = 104;
    const STOP = 106;

    const codes = [START_B];
    let checkSum = START_B;

    for (let i = 0; i < text.length; i++) {
      const code = text.charCodeAt(i) - 32;
      if (code < 0 || code > 94) return false;
      codes.push(code);
      checkSum += code * (i + 1);
    }

    codes.push(checkSum % 103);
    codes.push(STOP);

    let patternString = '';
    codes.forEach(c => {
      patternString += patterns[c] || '';
    });

    const barWidth = opts.width || 2.5;
    const barHeight = opts.height || 85;
    const quietZone = Math.max(barWidth * 14, 20); // ISO standard: at least 10x narrowest bar

    let totalWidth = quietZone * 2;
    for (let i = 0; i < patternString.length; i++) {
      totalWidth += parseInt(patternString[i], 10) * barWidth;
    }

    const textPad = opts.showText ? 22 : 8;
    const totalHeight = barHeight + textPad + 8;

    svg.setAttribute('width', totalWidth);
    svg.setAttribute('height', totalHeight);
    svg.setAttribute('viewBox', `0 0 ${totalWidth} ${totalHeight}`);
    svg.style.display = 'block';
    svg.style.maxWidth = '100%';

    let x = quietZone;
    // White background rect — essential for scanner contrast
    let paths = `<rect x="0" y="0" width="${totalWidth}" height="${totalHeight}" fill="#ffffff"/>`;

    for (let i = 0; i < patternString.length; i++) {
      const w = parseInt(patternString[i], 10) * barWidth;
      if (i % 2 === 0) {
        // Black bars with 2px top padding, to give visual breathing room
        paths += `<rect x="${x}" y="6" width="${w}" height="${barHeight}" fill="#000000"/>`;
      }
      x += w;
    }

    let textElement = '';
    if (opts.showText) {
      textElement = `<text x="${totalWidth / 2}" y="${barHeight + 20}" text-anchor="middle" font-family="JetBrains Mono, monospace" font-size="12" font-weight="700" fill="#000000">${escapeHtml(text)}</text>`;
    }

    svg.innerHTML = paths + textElement;
    return true;
  }

  // Boot
  document.addEventListener('DOMContentLoaded', init);
})();
