/**
 * 20x20 Spreadsheet Grid & Excel-Compatible Formula Engine
 * - Light Mode
 * - Immediate In-Cell Editing & Mobile Virtual Keyboard Popup on Tap
 * - Multi-Cell Block Selection (Shift + Click, Shift + Arrows, Mouse Drag)
 * - Real-Time Dynamic Statistics (Selected Range or Entire Sheet)
 * - Completely Decoupled from Calculator
 */

class Spreadsheet {
  constructor(rows = 20, cols = 20) {
    this.rowsCount = rows; // 1 to 20
    this.colsCount = cols; // A to T (20 cols)
    this.colHeaders = [];
    for (let i = 0; i < this.colsCount; i++) {
      this.colHeaders.push(String.fromCharCode(65 + i)); // A, B, C... T
    }

    this.data = {}; // key: "A1", value: { raw: string, val: any }
    
    // Selection & Edit State
    this.selectedCell = 'A1';
    this.rangeStart = { col: 0, row: 1 };
    this.rangeEnd = { col: 0, row: 1 };
    this.isSelecting = false;
    this.isEditing = false;
    this.editingCellId = null;

    // DOM Elements
    this.gridWrapper = document.getElementById('sheetGridWrapper');
    this.activeCellLabel = document.getElementById('activeCellBox');
    this.formulaInput = document.getElementById('formulaInput');
    this.statSum = document.getElementById('statSum');
    this.statAvg = document.getElementById('statAvg');
    this.statCount = document.getElementById('statCount');
    this.statModeTag = document.getElementById('statModeTag');

    this.initGrid();
    this.initEvents();
    this.updateRealtimeStats();
  }

  initGrid() {
    if (!this.gridWrapper) return;

    let html = '<table class="sheet-table" id="sheetTable">';
    
    // Header Row (Corner + A..T)
    html += '<thead><tr><th class="corner-header"></th>';
    for (let c = 0; c < this.colsCount; c++) {
      html += `<th id="colHeader_${this.colHeaders[c]}">${this.colHeaders[c]}</th>`;
    }
    html += '</tr></thead><tbody>';

    // Data Rows 1..20
    for (let r = 1; r <= this.rowsCount; r++) {
      html += `<tr><th class="row-header" id="rowHeader_${r}">${r}</th>`;
      for (let c = 0; c < this.colsCount; c++) {
        const cellId = `${this.colHeaders[c]}${r}`;
        html += `<td class="sheet-cell" id="cell_${cellId}" data-cell="${cellId}" data-col="${c}" data-row="${r}"></td>`;
      }
      html += '</tr>';
    }

    html += '</tbody></table>';
    this.gridWrapper.innerHTML = html;
    this.selectSingleCell('A1', false);
  }

  initEvents() {
    const table = document.getElementById('sheetTable');
    if (!table) return;

    // Cell Click / Touch Handler (Immediate focus for mobile virtual keyboard)
    table.addEventListener('click', (e) => {
      if (window.setActivePanel) window.setActivePanel('sheet');

      const td = e.target.closest('.sheet-cell');
      if (!td) return;

      const col = parseInt(td.dataset.col, 10);
      const row = parseInt(td.dataset.row, 10);
      const cellId = td.dataset.cell;

      if (e.shiftKey) {
        // Shift + Click: Block Range Selection
        if (this.isEditing) this.endEdit();
        this.rangeEnd = { col, row };
        this.applyRangeSelection();
      } else {
        // Single Click / Tap: Select and open in-cell editor immediately
        this.rangeStart = { col, row };
        this.rangeEnd = { col, row };
        this.selectSingleCell(cellId, true);
      }
    });

    // Mouse Drag Selection (Desktop)
    table.addEventListener('mousedown', (e) => {
      if (e.shiftKey) return;
      const td = e.target.closest('.sheet-cell');
      if (!td || e.target.classList.contains('cell-editor')) return;

      const col = parseInt(td.dataset.col, 10);
      const row = parseInt(td.dataset.row, 10);
      this.isSelecting = true;
      this.rangeStart = { col, row };
      this.rangeEnd = { col, row };
    });

    table.addEventListener('mouseover', (e) => {
      if (!this.isSelecting) return;
      const td = e.target.closest('.sheet-cell');
      if (td) {
        const col = parseInt(td.dataset.col, 10);
        const row = parseInt(td.dataset.row, 10);
        this.rangeEnd = { col, row };
        if (this.isEditing) this.endEdit();
        this.applyRangeSelection();
      }
    });

    window.addEventListener('mouseup', () => {
      if (this.isSelecting) {
        this.isSelecting = false;
      }
    });

    // Formula Input Synchronization
    if (this.formulaInput) {
      this.formulaInput.addEventListener('focus', () => {
        if (window.setActivePanel) window.setActivePanel('sheet');
      });

      this.formulaInput.addEventListener('input', (e) => {
        if (this.selectedCell) {
          this.setCellValue(this.selectedCell, e.target.value, false);
        }
      });

      this.formulaInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
          e.preventDefault();
          this.recalculateAll();
          this.moveSelection(1, 0, false);
        } else if (e.key === 'Escape') {
          this.formulaInput.value = this.getCellRaw(this.selectedCell);
          this.recalculateAll();
        }
      });
    }

    // Activate spreadsheet panel on click
    const sheetPanel = document.getElementById('sheetPanel');
    if (sheetPanel) {
      sheetPanel.addEventListener('mousedown', () => {
        if (window.setActivePanel) window.setActivePanel('sheet');
      });
    }

    // Keyboard Navigation & Range Selection (Shift + Arrows) & Direct Typing
    window.addEventListener('keydown', (e) => {
      if (window.activePanel !== 'sheet') return;

      // If formula input is focused, let it handle
      if (document.activeElement === this.formulaInput) return;

      // If currently editing inside a cell, let the editor handle it
      if (this.isEditing) return;

      const key = e.key;
      const shift = e.shiftKey;

      if (key === 'ArrowUp') {
        e.preventDefault();
        this.moveSelection(-1, 0, shift);
      } else if (key === 'ArrowDown') {
        e.preventDefault();
        this.moveSelection(1, 0, shift);
      } else if (key === 'ArrowLeft') {
        e.preventDefault();
        this.moveSelection(0, -1, shift);
      } else if (key === 'ArrowRight') {
        e.preventDefault();
        this.moveSelection(0, 1, shift);
      } else if (key === 'Tab') {
        e.preventDefault();
        this.moveSelection(0, shift ? -1 : 1, false);
      } else if (key === 'Enter') {
        e.preventDefault();
        this.startEdit(this.selectedCell);
      } else if (key === 'Delete' || key === 'Backspace') {
        e.preventDefault();
        this.clearSelectedRange();
      } else if (key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
        // Direct typing: Start editing immediately
        e.preventDefault();
        this.startEdit(this.selectedCell, key);
      }
    });

    // Toolbar Buttons
    const btnSampleData = document.getElementById('btnSampleData');
    if (btnSampleData) {
      btnSampleData.addEventListener('click', () => this.loadSampleData());
    }

    const btnClearSheet = document.getElementById('btnClearSheet');
    if (btnClearSheet) {
      btnClearSheet.addEventListener('click', () => this.clearSheet());
    }

    const btnExportCSV = document.getElementById('btnExportCSV');
    if (btnExportCSV) {
      btnExportCSV.addEventListener('click', () => this.exportCSV());
    }
  }

  selectSingleCell(cellId, autoStartEdit = false) {
    if (this.isEditing && this.editingCellId !== cellId) {
      this.endEdit();
    }

    this.clearSelectionStyles();

    this.selectedCell = cellId;
    const colChar = cellId.charAt(0);
    const rowNum = parseInt(cellId.slice(1), 10);
    const colIdx = this.colHeaders.indexOf(colChar);

    this.rangeStart = { col: colIdx, row: rowNum };
    this.rangeEnd = { col: colIdx, row: rowNum };

    const td = document.getElementById(`cell_${cellId}`);
    if (td) {
      td.classList.add('is-selected');
      if (typeof td.scrollIntoView === 'function') {
        td.scrollIntoView({ block: 'nearest', inline: 'nearest' });
      }
    }

    // Highlight headers
    this.highlightHeaders(colIdx, colIdx, rowNum, rowNum);

    // Update formula bar
    if (this.activeCellLabel) this.activeCellLabel.textContent = cellId;
    if (this.formulaInput) this.formulaInput.value = this.getCellRaw(cellId);

    this.updateRealtimeStats();

    // On mobile / user tap, immediately start in-cell edit so on-screen keyboard pops up
    if (autoStartEdit) {
      this.startEdit(cellId);
    }
  }

  applyRangeSelection() {
    this.clearSelectionStyles();

    const minCol = Math.min(this.rangeStart.col, this.rangeEnd.col);
    const maxCol = Math.max(this.rangeStart.col, this.rangeEnd.col);
    const minRow = Math.min(this.rangeStart.row, this.rangeEnd.row);
    const maxRow = Math.max(this.rangeStart.row, this.rangeEnd.row);

    const anchorId = `${this.colHeaders[this.rangeStart.col]}${this.rangeStart.row}`;
    this.selectedCell = anchorId;

    const isMulti = (minCol !== maxCol || minRow !== maxRow);

    for (let r = minRow; r <= maxRow; r++) {
      for (let c = minCol; c <= maxCol; c++) {
        const id = `${this.colHeaders[c]}${r}`;
        const td = document.getElementById(`cell_${id}`);
        if (td) {
          if (id === anchorId) {
            td.classList.add('is-selected');
          }
          if (isMulti) {
            td.classList.add('is-range-selected');
            if (r === minRow) td.classList.add('is-range-top');
            if (r === maxRow) td.classList.add('is-range-bottom');
            if (c === minCol) td.classList.add('is-range-left');
            if (c === maxCol) td.classList.add('is-range-right');
          }
        }
      }
    }

    this.highlightHeaders(minCol, maxCol, minRow, maxRow);

    // Update active label & formula input
    if (this.activeCellLabel) {
      if (isMulti) {
        const startId = `${this.colHeaders[minCol]}${minRow}`;
        const endId = `${this.colHeaders[maxCol]}${maxRow}`;
        this.activeCellLabel.textContent = `${startId}:${endId}`;
      } else {
        this.activeCellLabel.textContent = anchorId;
      }
    }
    if (this.formulaInput) {
      this.formulaInput.value = this.getCellRaw(anchorId);
    }

    this.updateRealtimeStats();
  }

  clearSelectionStyles() {
    document.querySelectorAll('.sheet-cell.is-selected, .sheet-cell.is-range-selected').forEach(el => {
      el.classList.remove('is-selected', 'is-range-selected', 'is-range-top', 'is-range-bottom', 'is-range-left', 'is-range-right');
    });
    document.querySelectorAll('.sheet-table th.header-active').forEach(el => el.classList.remove('header-active'));
  }

  highlightHeaders(minCol, maxCol, minRow, maxRow) {
    for (let c = minCol; c <= maxCol; c++) {
      const el = document.getElementById(`colHeader_${this.colHeaders[c]}`);
      if (el) el.classList.add('header-active');
    }
    for (let r = minRow; r <= maxRow; r++) {
      const el = document.getElementById(`rowHeader_${r}`);
      if (el) el.classList.add('header-active');
    }
  }

  moveSelection(rowDelta, colDelta, isShift = false) {
    if (isShift) {
      let nextCol = this.rangeEnd.col + colDelta;
      let nextRow = this.rangeEnd.row + rowDelta;
      nextCol = Math.max(0, Math.min(this.colsCount - 1, nextCol));
      nextRow = Math.max(1, Math.min(this.rowsCount, nextRow));
      this.rangeEnd = { col: nextCol, row: nextRow };
      this.applyRangeSelection();
    } else {
      let colIdx = this.rangeStart.col + colDelta;
      let rowNum = this.rangeStart.row + rowDelta;
      colIdx = Math.max(0, Math.min(this.colsCount - 1, colIdx));
      rowNum = Math.max(1, Math.min(this.rowsCount, rowNum));
      const nextCell = `${this.colHeaders[colIdx]}${rowNum}`;
      this.selectSingleCell(nextCell, false);
    }
  }

  clearSelectedRange() {
    const minCol = Math.min(this.rangeStart.col, this.rangeEnd.col);
    const maxCol = Math.max(this.rangeStart.col, this.rangeEnd.col);
    const minRow = Math.min(this.rangeStart.row, this.rangeEnd.row);
    const maxRow = Math.max(this.rangeStart.row, this.rangeEnd.row);

    for (let r = minRow; r <= maxRow; r++) {
      for (let c = minCol; c <= maxCol; c++) {
        const id = `${this.colHeaders[c]}${r}`;
        this.setCellValue(id, '', false);
      }
    }
    this.recalculateAll();
  }

  startEdit(cellId, initialChar = null) {
    if (this.isEditing && this.editingCellId === cellId) return;
    if (this.isEditing) {
      this.endEdit();
    }

    const td = document.getElementById(`cell_${cellId}`);
    if (!td) return;

    this.isEditing = true;
    this.editingCellId = cellId;
    td.classList.add('is-editing');

    const curRaw = initialChar !== null ? initialChar : this.getCellRaw(cellId);
    const editor = document.createElement('input');
    editor.type = 'text';
    editor.inputMode = 'text'; // Brings up mobile software keyboard
    editor.autocapitalize = 'off';
    editor.autocomplete = 'off';
    editor.autocorrect = 'off';
    editor.spellcheck = false;
    editor.className = 'cell-editor';
    editor.value = curRaw;

    td.innerHTML = '';
    td.appendChild(editor);

    // Focus editor immediately
    editor.focus();
    if (initialChar === null) {
      editor.select();
    }

    // Keep formula bar in sync while typing
    editor.addEventListener('input', (e) => {
      this.setCellValue(cellId, e.target.value, false);
      if (this.formulaInput) {
        this.formulaInput.value = e.target.value;
      }
    });

    const commit = (shouldMove = false) => {
      const val = editor.value;
      this.isEditing = false;
      this.editingCellId = null;
      td.classList.remove('is-editing');
      this.setCellValue(cellId, val, true);
      if (shouldMove) {
        this.moveSelection(1, 0, false);
      }
    };

    editor.addEventListener('blur', () => {
      if (this.isEditing && this.editingCellId === cellId) {
        commit(false);
      }
    });

    editor.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        commit(true);
      } else if (e.key === 'Escape') {
        e.preventDefault();
        this.isEditing = false;
        this.editingCellId = null;
        td.classList.remove('is-editing');
        this.renderCell(cellId);
      } else if (e.key === 'Tab') {
        e.preventDefault();
        commit(false);
        this.moveSelection(0, e.shiftKey ? -1 : 1, false);
      }
    });
  }

  endEdit() {
    if (!this.isEditing) return;
    this.isEditing = false;
    const editingId = this.editingCellId;
    this.editingCellId = null;

    document.querySelectorAll('.cell-editor').forEach(ed => {
      const td = ed.closest('.sheet-cell');
      if (td) {
        const cellId = td.dataset.cell;
        const val = ed.value;
        td.classList.remove('is-editing');
        this.setCellValue(cellId, val, true);
      }
    });

    if (editingId) {
      this.renderCell(editingId);
    }
  }

  getCellRaw(cellId) {
    return this.data[cellId]?.raw || '';
  }

  getCellValue(cellId) {
    return this.data[cellId]?.val;
  }

  setCellValue(cellId, rawVal, triggerRecalc = true) {
    if (!this.data[cellId]) {
      this.data[cellId] = { raw: '', val: null };
    }
    this.data[cellId].raw = String(rawVal).trim();
    if (triggerRecalc) {
      this.recalculateAll();
    }
    if (this.formulaInput && this.selectedCell === cellId && document.activeElement !== this.formulaInput) {
      this.formulaInput.value = this.data[cellId].raw;
    }
  }

  recalculateAll() {
    const visited = new Set();
    for (let r = 1; r <= this.rowsCount; r++) {
      for (let c = 0; c < this.colsCount; c++) {
        const cellId = `${this.colHeaders[c]}${r}`;
        this.evalCell(cellId, visited);
        this.renderCell(cellId);
      }
    }
    this.updateRealtimeStats();
  }

  evalCell(cellId, visited = new Set()) {
    const item = this.data[cellId];
    if (!item || !item.raw) {
      if (item) item.val = null;
      return null;
    }

    const raw = item.raw;

    if (raw.startsWith('=')) {
      try {
        const formula = raw.slice(1);
        item.val = this.evaluateFormula(formula, cellId, visited);
      } catch (err) {
        item.val = '#ERROR!';
      }
    } else {
      const num = Number(raw);
      if (!isNaN(num) && raw !== '') {
        item.val = num;
      } else {
        item.val = raw;
      }
    }

    return item.val;
  }

  renderCell(cellId) {
    const td = document.getElementById(`cell_${cellId}`);
    if (!td || td.classList.contains('is-editing')) return;

    const val = this.getCellValue(cellId);
    if (val === null || val === undefined || val === '') {
      td.textContent = '';
      td.classList.remove('is-error', 'text-left');
    } else if (typeof val === 'string' && val.startsWith('#')) {
      td.textContent = val;
      td.classList.add('is-error');
    } else if (typeof val === 'number') {
      const formatted = Number.isInteger(val) ? val.toLocaleString() : (Math.abs(val) < 0.0001 && val !== 0 ? val.toExponential(4) : parseFloat(val.toFixed(4)).toLocaleString());
      td.textContent = formatted;
      td.classList.remove('is-error', 'text-left');
    } else {
      td.textContent = String(val);
      td.classList.remove('is-error');
      td.classList.add('text-left');
    }
  }

  evaluateFormula(expr, currentCell, callStack = new Set()) {
    if (callStack.has(currentCell)) {
      return '#REF!';
    }
    const nextStack = new Set(callStack);
    nextStack.add(currentCell);

    let parsed = expr.toUpperCase().trim();

    // 1. Expand Range Functions: SUM(A1:A5), AVERAGE(B1:B10), etc.
    parsed = parsed.replace(/\b([A-Z_]+)\s*\(\s*([A-T]\d+)\s*:\s*([A-T]\d+)\s*\)/g, (match, fn, startCell, endCell) => {
      const values = this.getRangeValues(startCell, endCell, nextStack);
      return this.computeFormulaFunction(fn, values);
    });

    // 2. Expand List Functions: SUM(A1, B2, 10), etc.
    parsed = parsed.replace(/\b([A-Z_]+)\s*\(([^()]+)\)/g, (match, fn, argsStr) => {
      const argTokens = argsStr.split(',').map(s => s.trim());
      const values = [];
      for (const token of argTokens) {
        if (/^[A-T]\d+$/.test(token)) {
          const v = this.resolveCellRef(token, nextStack);
          if (v !== null && !isNaN(Number(v))) values.push(Number(v));
        } else {
          const num = Number(token);
          if (!isNaN(num)) values.push(num);
        }
      }
      return this.computeFormulaFunction(fn, values);
    });

    // 3. Replace single cell references (e.g. A1, B2)
    parsed = parsed.replace(/\b([A-T]\d+)\b/g, (match, refCell) => {
      const v = this.resolveCellRef(refCell, nextStack);
      if (typeof v === 'string' && v.startsWith('#')) throw new Error(v);
      return Number(v) || 0;
    });

    // 4. Safe evaluate arithmetic expression
    try {
      parsed = parsed.replace(/\^/g, '**');
      // eslint-disable-next-line no-new-func
      const result = Function(`"use strict"; return (${parsed});`)();
      return typeof result === 'number' && !isNaN(result) ? result : '#VALUE!';
    } catch (e) {
      return '#VALUE!';
    }
  }

  resolveCellRef(cellId, callStack) {
    if (!this.data[cellId]) return 0;
    if (this.data[cellId].val !== undefined && this.data[cellId].val !== null) {
      return this.data[cellId].val;
    }
    return this.evalCell(cellId, callStack);
  }

  getRangeValues(startCell, endCell, callStack) {
    const startCol = startCell.charAt(0);
    const startRow = parseInt(startCell.slice(1), 10);
    const endCol = endCell.charAt(0);
    const endRow = parseInt(endCell.slice(1), 10);

    const minColIdx = Math.min(this.colHeaders.indexOf(startCol), this.colHeaders.indexOf(endCol));
    const maxColIdx = Math.max(this.colHeaders.indexOf(startCol), this.colHeaders.indexOf(endCol));
    const minRow = Math.min(startRow, endRow);
    const maxRow = Math.max(startRow, endRow);

    const values = [];
    for (let c = minColIdx; c <= maxColIdx; c++) {
      for (let r = minRow; r <= maxRow; r++) {
        const id = `${this.colHeaders[c]}${r}`;
        const val = this.resolveCellRef(id, callStack);
        if (val !== null && val !== undefined && val !== '' && !isNaN(Number(val))) {
          values.push(Number(val));
        }
      }
    }
    return values;
  }

  computeFormulaFunction(fn, values) {
    if (!values || values.length === 0) {
      if (fn === 'COUNT' || fn === 'SUM') return 0;
      return 0;
    }

    switch (fn) {
      case 'SUM':
        return values.reduce((a, b) => a + b, 0);
      case 'AVERAGE':
      case 'AVG':
        return values.reduce((a, b) => a + b, 0) / values.length;
      case 'COUNT':
      case 'COUNTA':
        return values.length;
      case 'MAX':
        return Math.max(...values);
      case 'MIN':
        return Math.min(...values);
      case 'PRODUCT':
        return values.reduce((a, b) => a * b, 1);
      case 'MEDIAN': {
        const sorted = [...values].sort((a, b) => a - b);
        const mid = Math.floor(sorted.length / 2);
        return sorted.length % 2 !== 0 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
      }
      case 'STDEV': {
        const mean = values.reduce((a, b) => a + b, 0) / values.length;
        const variance = values.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / values.length;
        return Math.sqrt(variance);
      }
      case 'SQRT':
        return Math.sqrt(values[0] || 0);
      case 'ABS':
        return Math.abs(values[0] || 0);
      default:
        return values[0] || 0;
    }
  }

  /**
   * Real-time Statistics Calculation:
   * - If a multi-cell block is selected: calculates for that selection range!
   * - If 1 cell is selected: calculates for all populated numbers in the entire 20x20 sheet!
   */
  updateRealtimeStats() {
    let sum = 0;
    let count = 0;

    const minCol = Math.min(this.rangeStart.col, this.rangeEnd.col);
    const maxCol = Math.max(this.rangeStart.col, this.rangeEnd.col);
    const minRow = Math.min(this.rangeStart.row, this.rangeEnd.row);
    const maxRow = Math.max(this.rangeStart.row, this.rangeEnd.row);

    const isMultiSelection = (minCol !== maxCol || minRow !== maxRow);

    if (isMultiSelection) {
      for (let r = minRow; r <= maxRow; r++) {
        for (let c = minCol; c <= maxCol; c++) {
          const cellId = `${this.colHeaders[c]}${r}`;
          const item = this.data[cellId];
          if (item && typeof item.val === 'number' && !isNaN(item.val) && isFinite(item.val)) {
            sum += item.val;
            count++;
          }
        }
      }
      if (this.statModeTag) {
        const startId = `${this.colHeaders[minCol]}${minRow}`;
        const endId = `${this.colHeaders[maxCol]}${maxRow}`;
        this.statModeTag.textContent = `🎯 선택 영역 [${startId}:${endId}] 통계`;
      }
    } else {
      for (let r = 1; r <= this.rowsCount; r++) {
        for (let c = 0; c < this.colsCount; c++) {
          const cellId = `${this.colHeaders[c]}${r}`;
          const item = this.data[cellId];
          if (item && typeof item.val === 'number' && !isNaN(item.val) && isFinite(item.val)) {
            sum += item.val;
            count++;
          }
        }
      }
      if (this.statModeTag) {
        this.statModeTag.textContent = `⚡ 전체 20×20 바둑판 통계`;
      }
    }

    const avg = count > 0 ? (sum / count) : 0;

    if (this.statSum) {
      this.statSum.textContent = Number.isInteger(sum) ? sum.toLocaleString() : parseFloat(sum.toFixed(4)).toLocaleString();
    }
    if (this.statAvg) {
      this.statAvg.textContent = count > 0 ? (Number.isInteger(avg) ? avg.toLocaleString() : parseFloat(avg.toFixed(4)).toLocaleString()) : '0';
    }
    if (this.statCount) {
      this.statCount.textContent = count.toLocaleString();
    }
  }

  loadSampleData() {
    this.clearSheet();

    this.setCellValue('A1', '항목', false);
    this.setCellValue('B1', '수량', false);
    this.setCellValue('C1', '단가', false);
    this.setCellValue('D1', '금액(=B*C)', false);

    const items = [
      ['센서 모듈 A', 15, 12500],
      ['MCU 보드', 8, 34000],
      ['배터리 팩', 20, 8900],
      ['디스플레이 LCD', 12, 18500],
      ['통신 안테나', 30, 4500],
      ['알루미늄 케이스', 10, 22000]
    ];

    items.forEach((item, idx) => {
      const row = idx + 2;
      this.setCellValue(`A${row}`, item[0], false);
      this.setCellValue(`B${row}`, item[1], false);
      this.setCellValue(`C${row}`, item[2], false);
      this.setCellValue(`D${row}`, `=B${row}*C${row}`, false);
    });

    this.setCellValue('A9', '총합계', false);
    this.setCellValue('D9', '=SUM(D2:D7)', false);
    this.setCellValue('A10', '평균 단가', false);
    this.setCellValue('C10', '=AVERAGE(C2:C7)', false);

    this.recalculateAll();
    this.selectSingleCell('A1', false);
    this.showToast('예제 데이터가 로드되었습니다.');
  }

  clearSheet() {
    this.data = {};
    for (let r = 1; r <= this.rowsCount; r++) {
      for (let c = 0; c < this.colsCount; c++) {
        const cellId = `${this.colHeaders[c]}${r}`;
        const td = document.getElementById(`cell_${cellId}`);
        if (td) {
          td.textContent = '';
          td.className = 'sheet-cell';
        }
      }
    }
    if (this.formulaInput) this.formulaInput.value = '';
    this.updateRealtimeStats();
    this.showToast('스프레드시트가 초기화되었습니다.');
  }

  exportCSV() {
    let csv = '';
    for (let r = 1; r <= this.rowsCount; r++) {
      const rowVals = [];
      for (let c = 0; c < this.colsCount; c++) {
        const id = `${this.colHeaders[c]}${r}`;
        const val = this.getCellValue(id);
        const cellText = val !== null && val !== undefined ? String(val).replace(/"/g, '""') : '';
        rowVals.push(`"${cellText}"`);
      }
      csv += rowVals.join(',') + '\r\n';
    }

    const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', `spreadsheet_20x20_${new Date().toISOString().slice(0,10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    this.showToast('CSV 파일이 다운로드되었습니다.');
  }

  showToast(msg) {
    const container = document.getElementById('toastContainer');
    if (!container) return;
    const toast = document.createElement('div');
    toast.className = 'toast';
    toast.innerHTML = `<span>📊</span> <span>${msg}</span>`;
    container.appendChild(toast);
    setTimeout(() => toast.remove(), 2200);
  }
}

window.Spreadsheet = Spreadsheet;
