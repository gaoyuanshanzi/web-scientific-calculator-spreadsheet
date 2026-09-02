/**
 * 20×50 Spreadsheet Grid & Excel-Compatible Formula Engine
 * - Light Mode
 * - Immediate In-Cell Editing & Mobile Virtual Keyboard Popup on Tap
 * - Multi-Cell Block Selection (Shift + Click, Shift + Arrows, Mouse Drag)
 * - Real-Time Dynamic Statistics (Selected Range or Entire Sheet)
 * - Completely Decoupled from Calculator
 */

class Spreadsheet {
  constructor(rows = 50, cols = 20) {
    this.rowsCount = rows; // 1 to 50
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

    // Formula Cell Pointing & Range Picking State (Excel-like formula cell selection)
    this.isPickingFormula = false;
    this.formulaRefStart = null;
    this.formulaRefEnd = null;
    this.pickingState = null;
    this.pointCell = null;

    // DOM Elements
    this.gridWrapper = document.getElementById('sheetGridWrapper');
    this.activeCellLabel = document.getElementById('activeCellBox');
    this.formulaInput = document.getElementById('formulaInput');
    this.statSum = document.getElementById('statSum');
    this.statAvg = document.getElementById('statAvg');
    this.statCount = document.getElementById('statCount');
    this.statVar = document.getElementById('statVar');
    this.statStdev = document.getElementById('statStdev');
    this.statSS = document.getElementById('statSS');
    this.statSumSq = document.getElementById('statSumSq');
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

    // Cell Click / Touch Handler
    table.addEventListener('click', (e) => {
      if (window.setActivePanel) window.setActivePanel('sheet');

      const td = e.target.closest('.sheet-cell');
      if (!td) return;

      const col = parseInt(td.dataset.col, 10);
      const row = parseInt(td.dataset.row, 10);
      const cellId = td.dataset.cell;

      const activeEd = document.querySelector('.cell-editor');
      const isFormulaMode = (this.isEditing && activeEd && activeEd.value.startsWith('=')) ||
                            (document.activeElement === this.formulaInput && this.formulaInput.value.startsWith('='));

      // If in formula mode and clicked another cell, picking was handled by mousedown
      if (isFormulaMode && this.editingCellId !== cellId) {
        return;
      }

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

    // Mouse Drag Selection / Formula Cell Pointing
    table.addEventListener('mousedown', (e) => {
      if (e.shiftKey) return;
      const td = e.target.closest('.sheet-cell');
      if (!td) return;

      const col = parseInt(td.dataset.col, 10);
      const row = parseInt(td.dataset.row, 10);
      const cellId = td.dataset.cell;

      const activeEd = document.querySelector('.cell-editor');
      const isFormulaMode = (this.isEditing && activeEd && activeEd.value.startsWith('=')) ||
                            (document.activeElement === this.formulaInput && this.formulaInput.value.startsWith('='));

      if (isFormulaMode) {
        // If clicking inside the current cell being edited, let user position text cursor
        if (e.target.classList.contains('cell-editor') && this.editingCellId === cellId) {
          return;
        }

        e.preventDefault(); // Keep focus in editor
        this.isPickingFormula = true;
        this.formulaRefStart = { col, row };
        this.formulaRefEnd = { col, row };
        this.insertOrUpdateFormulaRef(cellId);
        this.highlightFormulaRange(col, col, row, row);
        return;
      }

      if (e.target.classList.contains('cell-editor')) return;

      this.isSelecting = true;
      this.rangeStart = { col, row };
      this.rangeEnd = { col, row };
    });

    table.addEventListener('mouseover', (e) => {
      const td = e.target.closest('.sheet-cell');
      if (!td) return;

      const col = parseInt(td.dataset.col, 10);
      const row = parseInt(td.dataset.row, 10);

      if (this.isPickingFormula) {
        this.formulaRefEnd = { col, row };
        const minC = Math.min(this.formulaRefStart.col, this.formulaRefEnd.col);
        const maxC = Math.max(this.formulaRefStart.col, this.formulaRefEnd.col);
        const minR = Math.min(this.formulaRefStart.row, this.formulaRefEnd.row);
        const maxR = Math.max(this.formulaRefStart.row, this.formulaRefEnd.row);
        const startId = `${this.colHeaders[minC]}${minR}`;
        const endId = `${this.colHeaders[maxC]}${maxR}`;
        const refStr = (startId === endId) ? startId : `${startId}:${endId}`;
        this.insertOrUpdateFormulaRef(refStr);
        this.highlightFormulaRange(minC, maxC, minR, maxR);
        return;
      }

      if (this.isSelecting) {
        this.rangeEnd = { col, row };
        if (this.isEditing) this.endEdit();
        this.applyRangeSelection();
      }
    });

    window.addEventListener('mouseup', () => {
      if (this.isPickingFormula) {
        this.isPickingFormula = false;
        const ed = document.querySelector('.cell-editor') || this.formulaInput;
        if (ed) ed.focus();
      }
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
        if (['+', '-', '*', '/', '^', '(', ')', ',', ':'].includes(e.key)) {
          this.pickingState = null;
          this.pointCell = null;
          this.clearFormulaRefHighlights();
        }

        if (e.key === 'Enter') {
          e.preventDefault();
          this.pickingState = null;
          this.pointCell = null;
          this.clearFormulaRefHighlights();
          this.formulaInput.blur();
          this.recalculateAll();
          this.moveSelection(1, 0, false);
        } else if (e.key === 'Escape') {
          e.preventDefault();
          this.pickingState = null;
          this.pointCell = null;
          this.clearFormulaRefHighlights();
          this.formulaInput.value = this.getCellRaw(this.selectedCell);
          this.recalculateAll();
          this.formulaInput.blur();
          // Move focus & cursor to the active cell
          if (this.selectedCell) {
            this.selectSingleCell(this.selectedCell, true);
          }
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

      // Excel Ctrl + D (Fill Down / 아래로 채우기)
      if ((e.ctrlKey || e.metaKey) && (e.key === 'd' || e.key === 'D')) {
        e.preventDefault();
        this.fillDown();
        return;
      }

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
    const btnFillDown = document.getElementById('btnFillDown');
    if (btnFillDown) {
      btnFillDown.addEventListener('click', () => {
        if (this.isEditing) this.endEdit();
        this.fillDown();
      });
    }

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
      if (!this.isPickingFormula) {
        this.pointCell = null;
        this.clearFormulaRefHighlights();
      }
    });

    const commit = (shouldMove = false) => {
      const val = editor.value;
      this.isEditing = false;
      this.editingCellId = null;
      this.pickingState = null;
      this.pointCell = null;
      this.clearFormulaRefHighlights();
      td.classList.remove('is-editing');
      this.setCellValue(cellId, val, true);
      if (shouldMove) {
        this.moveSelection(1, 0, false);
      }
    };

    editor.addEventListener('blur', () => {
      if (this.isPickingFormula) return;
      if (this.isEditing && this.editingCellId === cellId) {
        commit(false);
      }
    });

    editor.addEventListener('keydown', (e) => {
      // If typing an operator while in formula, reset picking state so next click inserts new reference
      if (['+', '-', '*', '/', '^', '(', ')', ',', ':'].includes(e.key)) {
        this.pickingState = null;
        this.pointCell = null;
        this.clearFormulaRefHighlights();
      }

      // Excel Ctrl + D (Fill Down) inside active editor
      if ((e.ctrlKey || e.metaKey) && (e.key === 'd' || e.key === 'D')) {
        e.preventDefault();
        commit(false);
        this.fillDown();
        return;
      }

      if (e.key === 'Enter') {
        e.preventDefault();
        commit(true);
        return;
      } else if (e.key === 'Escape') {
        e.preventDefault();
        this.isEditing = false;
        this.editingCellId = null;
        this.pickingState = null;
        this.pointCell = null;
        this.clearFormulaRefHighlights();
        td.classList.remove('is-editing');
        this.renderCell(cellId);
        return;
      } else if (e.key === 'Tab') {
        e.preventDefault();
        commit(false);
        this.moveSelection(0, e.shiftKey ? -1 : 1, false);
        return;
      }

      // Check for Excel Point Mode with Arrow Keys
      const val = editor.value;
      const isFormula = val.startsWith('=');
      const cursorPos = editor.selectionStart;
      const charBeforeCursor = cursorPos > 0 ? val.charAt(cursorPos - 1) : '';
      const isAfterOperator = /[=+\-*/^,:(]/.test(charBeforeCursor);

      if (isFormula && (this.pointCell || isAfterOperator) && ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)) {
        e.preventDefault();
        
        if (!this.pointCell) {
          const editColChar = cellId.charAt(0);
          const editRowNum = parseInt(cellId.slice(1), 10);
          const editColIdx = this.colHeaders.indexOf(editColChar);
          this.pointCell = { col: editColIdx, row: editRowNum };
        }

        if (e.key === 'ArrowUp') this.pointCell.row = Math.max(1, this.pointCell.row - 1);
        if (e.key === 'ArrowDown') this.pointCell.row = Math.min(this.rowsCount, this.pointCell.row + 1);
        if (e.key === 'ArrowLeft') this.pointCell.col = Math.max(0, this.pointCell.col - 1);
        if (e.key === 'ArrowRight') this.pointCell.col = Math.min(this.colsCount - 1, this.pointCell.col + 1);

        const targetCellId = `${this.colHeaders[this.pointCell.col]}${this.pointCell.row}`;
        this.insertOrUpdateFormulaRef(targetCellId);
        this.highlightFormulaRange(this.pointCell.col, this.pointCell.col, this.pointCell.row, this.pointCell.row);
        return;
      }

      if (e.key === 'ArrowUp') {
        e.preventDefault();
        commit(false);
        this.moveSelection(-1, 0, false);
      } else if (e.key === 'ArrowDown') {
        e.preventDefault();
        commit(false);
        this.moveSelection(1, 0, false);
      } else if (e.key === 'ArrowLeft') {
        // Only navigate if cursor is at the start of the text
        if (editor.selectionStart === 0 && editor.selectionEnd === 0) {
          e.preventDefault();
          commit(false);
          this.moveSelection(0, -1, false);
        }
      } else if (e.key === 'ArrowRight') {
        // Only navigate if cursor is at the end of the text
        if (editor.selectionStart === editor.value.length && editor.selectionEnd === editor.value.length) {
          e.preventDefault();
          commit(false);
          this.moveSelection(0, 1, false);
        }
      }
    });
  }

  endEdit() {
    if (!this.isEditing) return;
    this.isEditing = false;
    const editingId = this.editingCellId;
    this.editingCellId = null;
    this.pickingState = null;
    this.pointCell = null;
    this.clearFormulaRefHighlights();

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
   * - If 1 cell is selected: calculates for all populated numbers in the entire sheet!
   */
  updateRealtimeStats() {
    const values = [];

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
            values.push(item.val);
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
            values.push(item.val);
          }
        }
      }
      if (this.statModeTag) {
        this.statModeTag.textContent = `⚡ 전체 20×50 통계`;
      }
    }

    const count = values.length;
    const sum = count > 0 ? values.reduce((a, b) => a + b, 0) : 0;
    const avg = count > 0 ? sum / count : 0;
    const sumSq = count > 0 ? values.reduce((a, b) => a + b * b, 0) : 0;
    const ss = count > 0 ? values.reduce((a, b) => a + Math.pow(b - avg, 2), 0) : 0; // 편차제곱합
    const variance = count > 0 ? ss / count : 0;
    const stdev = Math.sqrt(variance);

    const fmt = (v) => {
      if (count === 0) return '0';
      return Number.isInteger(v) ? v.toLocaleString() : parseFloat(v.toFixed(4)).toLocaleString();
    };

    if (this.statSum)   this.statSum.textContent   = fmt(sum);
    if (this.statAvg)   this.statAvg.textContent   = fmt(avg);
    if (this.statCount) this.statCount.textContent  = count.toLocaleString();
    if (this.statVar)   this.statVar.textContent    = fmt(variance);
    if (this.statStdev) this.statStdev.textContent  = fmt(stdev);
    if (this.statSS)    this.statSS.textContent     = fmt(ss);
    if (this.statSumSq) this.statSumSq.textContent  = fmt(sumSq);
  }

  /**
   * Excel Ctrl + D (Fill Down / 아래로 채우기)
   * 1. If multiple rows are selected (minRow < maxRow):
   *    Copies the top row's values/formulas down to all cells below in each selected column.
   *    Formulas automatically adjust relative row references (e.g. =A1+B1 -> =A2+B2).
   * 2. If a single row is selected (minRow === maxRow):
   *    Copies the cell immediately above (row - 1) into the current cell.
   */
  fillDown() {
    const minCol = Math.min(this.rangeStart.col, this.rangeEnd.col);
    const maxCol = Math.max(this.rangeStart.col, this.rangeEnd.col);
    const minRow = Math.min(this.rangeStart.row, this.rangeEnd.row);
    const maxRow = Math.max(this.rangeStart.row, this.rangeEnd.row);

    if (minRow < maxRow) {
      // Case 1: Multi-row selection
      for (let c = minCol; c <= maxCol; c++) {
        const sourceCellId = `${this.colHeaders[c]}${minRow}`;
        const sourceRaw = this.getCellRaw(sourceCellId);

        for (let r = minRow + 1; r <= maxRow; r++) {
          const targetCellId = `${this.colHeaders[c]}${r}`;
          const rowOffset = r - minRow;
          const adjustedVal = this.adjustFormulaRowOffset(sourceRaw, rowOffset);
          this.setCellValue(targetCellId, adjustedVal, false);
        }
      }
      this.recalculateAll();
      const startId = `${this.colHeaders[minCol]}${minRow}`;
      const endId = `${this.colHeaders[maxCol]}${maxRow}`;
      this.showToast(`아래로 채우기(Ctrl+D) 완료: [${startId}:${endId}]`);
    } else {
      // Case 2: Single-row selection (copy from row above)
      if (minRow > 1) {
        for (let c = minCol; c <= maxCol; c++) {
          const sourceCellId = `${this.colHeaders[c]}${minRow - 1}`;
          const sourceRaw = this.getCellRaw(sourceCellId);
          const targetCellId = `${this.colHeaders[c]}${minRow}`;
          const adjustedVal = this.adjustFormulaRowOffset(sourceRaw, 1);
          this.setCellValue(targetCellId, adjustedVal, false);
        }
        this.recalculateAll();
        if (this.formulaInput && this.selectedCell) {
          this.formulaInput.value = this.getCellRaw(this.selectedCell);
        }
        this.showToast(`위 셀 복사(Ctrl+D) 완료: ${this.selectedCell}`);
      } else {
        this.showToast('1행 위에는 복사할 셀이 없습니다.');
      }
    }
  }

  /**
   * Adjust relative row references in formulas when filling down
   * e.g. =A1*B1 with offset +1 -> =A2*B2
   * $A$1 or A$1 (absolute row reference) remains unchanged
   */
  adjustFormulaRowOffset(raw, rowOffset) {
    if (!raw || typeof raw !== 'string' || !raw.startsWith('=')) {
      return raw;
    }
    return raw.replace(/(?<![A-Za-z0-9_])(\$?)([A-T])(\$?)(\d+)(?![A-Za-z0-9_])/gi, (match, colLock, colLetter, rowLock, rowStr) => {
      if (rowLock === '$') {
        return match;
      }
      const origRow = parseInt(rowStr, 10);
      const newRow = origRow + rowOffset;
      if (newRow >= 1 && newRow <= this.rowsCount) {
        return `${colLock}${colLetter.toUpperCase()}${rowLock}${newRow}`;
      }
      return match;
    });
  }

  /**
   * Excel Formula Reference Insertion & Updating:
   * Inserts or replaces a referenced cell/range at the current formula insertion point
   */
  insertOrUpdateFormulaRef(refStr) {
    const editor = document.querySelector('.cell-editor') || this.formulaInput;
    if (!editor) return;

    const text = editor.value;
    if (!this.pickingState) {
      const start = editor.selectionStart !== null ? Math.min(editor.selectionStart, text.length) : text.length;
      const end = editor.selectionEnd !== null ? Math.min(editor.selectionEnd, text.length) : text.length;
      this.pickingState = {
        startPos: start,
        refLength: refStr.length
      };
      const before = text.substring(0, start);
      const after = text.substring(end);
      editor.value = before + refStr + after;
      editor.selectionStart = editor.selectionEnd = start + refStr.length;
    } else {
      const start = this.pickingState.startPos;
      const oldLen = this.pickingState.refLength;
      const before = text.substring(0, start);
      const after = text.substring(start + oldLen);
      editor.value = before + refStr + after;
      this.pickingState.refLength = refStr.length;
      editor.selectionStart = editor.selectionEnd = start + refStr.length;
    }

    if (this.editingCellId) {
      this.setCellValue(this.editingCellId, editor.value, false);
    }
    if (this.formulaInput && editor !== this.formulaInput) {
      this.formulaInput.value = editor.value;
    }
  }

  clearFormulaRefHighlights() {
    document.querySelectorAll('.sheet-cell.is-formula-ref').forEach(el => {
      el.classList.remove('is-formula-ref');
    });
  }

  highlightFormulaRange(minCol, maxCol, minRow, maxRow) {
    this.clearFormulaRefHighlights();
    for (let r = minRow; r <= maxRow; r++) {
      for (let c = minCol; c <= maxCol; c++) {
        const id = `${this.colHeaders[c]}${r}`;
        const td = document.getElementById(`cell_${id}`);
        if (td) {
          td.classList.add('is-formula-ref');
        }
      }
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
