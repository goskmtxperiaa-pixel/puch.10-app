// 【完全統合版 script.js】
// CSV・Excel・相対参照・スマホ表示修正をすべて含んでいます
const NUM_COLS = 50;
const NUM_ROWS = 200;

let state = {
  cells: {}, selected: [], isEditing: false, focusCell: null, inputBuffer: '', 
  referenceMode: false, keyboardMode: 'hidden', handedness: 'right', keyboardMinimized: false,
  clipboard: null
};

const undoStack = [];
function saveHistory() {
  undoStack.push(JSON.stringify(state.cells));
  if (undoStack.length > 20) undoStack.shift();
}
function performUndo() {
  if (undoStack.length > 0) {
    const prevState = undoStack.pop();
    state.cells = JSON.parse(prevState);
    if (state.focusCell && state.cells[state.focusCell]) {
      state.inputBuffer = state.cells[state.focusCell].input || '';
    } else { state.inputBuffer = ''; }
    checkReferenceMode(); renderCells();
  }
}

// UI Elements
const gridEl = document.getElementById('grid');
const formulaBar = document.getElementById('formula-bar'); 
const toggleInputBtn = document.getElementById('btn-toggle-input');
const customKeyboard = document.getElementById('custom-keyboard');
const kbNormal = document.getElementById('kb-normal');
const kbMinimized = document.getElementById('kb-minimized');
const formulaHintEl = document.getElementById('formula-hint');
const buttonTooltipEl = document.getElementById('button-tooltip');
const kbResizeHandle = document.getElementById('kb-resize-handle');
const vKeys = [document.getElementById('v-key-1'), document.getElementById('v-key-2'), document.getElementById('v-key-3')];
const btnOpen = document.getElementById('btn-open');
const btnCopy = document.getElementById('btn-copy');
const btnPaste = document.getElementById('btn-paste');
const fileInput = document.getElementById('file-input');
const btnExportCsv = document.getElementById('btn-export-csv');
const btnExportExcel = document.getElementById('btn-export-excel');

// Keyboard Resizing
let isResizing = false;
kbResizeHandle.addEventListener('pointerdown', (e) => {
    isResizing = true;
    document.body.style.cursor = 'ns-resize';
    kbResizeHandle.setPointerCapture(e.pointerId);
});
window.addEventListener('pointermove', (e) => {
    if (!isResizing) return;
    const availableSpace = window.innerHeight - e.clientY - 40; 
    const newHeight = Math.max(18, Math.min(100, availableSpace / 4));
    document.documentElement.style.setProperty('--kb-row-height', `${newHeight}px`);
    const newFontSize = Math.max(9, Math.min(18, 9 + (newHeight - 18) * 0.3));
    document.documentElement.style.setProperty('--kb-font-size', `${newFontSize}px`);
});
window.addEventListener('pointerup', () => {
    if (isResizing) { isResizing = false; document.body.style.cursor = ''; }
});

// Grid logic
function getColLabel(ci) {
  let l = ''; let t = ci;
  while (t >= 0) { l = String.fromCharCode(65 + (t % 26)) + l; t = Math.floor(t / 26) - 1; }
  return l;
}
function getCellId(c, r) { return getColLabel(c) + (r + 1); }

function initGrid() {
  let h = '';
  for (let r = -1; r < NUM_ROWS; r++) {
    for (let c = -1; c < NUM_COLS; c++) {
      if (r === -1 && c === -1) h += `<div class="cell header-corner"></div>`;
      else if (r === -1) h += `<div class="cell header-row">${getColLabel(c)}</div>`;
      else if (c === -1) h += `<div class="cell header-col">${r + 1}</div>`;
      else { const id = getCellId(c, r); h += `<div class="cell" id="cell-${id}" data-col="${c}" data-row="${r}" data-id="${id}"></div>`; }
    }
  }
  gridEl.innerHTML = h;
}

gridEl.addEventListener('click', (e) => {
    const cell = e.target.closest('.cell');
    if (!cell || cell.classList.contains('header-row') || cell.classList.contains('header-col')) return;
    handleShortPress(cell);
});

function renderCells() {
  document.querySelectorAll('.cell.selected').forEach(el => el.classList.remove('selected', 'editing'));
  Object.keys(state.cells).forEach(id => {
      const el = document.getElementById(`cell-${id}`);
      if (el) el.textContent = state.cells[id].display;
  });
  if (state.focusCell) {
      const el = document.getElementById(`cell-${state.focusCell}`);
      if (el) el.classList.add('selected');
      formulaBar.value = state.isEditing ? state.inputBuffer : (state.cells[state.focusCell]?.input || '');
  }
}

function handleShortPress(cellEl) {
  const id = cellEl.dataset.id;
  if (state.isEditing && state.focusCell && state.focusCell !== id && state.inputBuffer.startsWith('=')) {
      state.inputBuffer += id; renderCells(); return;
  }
  state.selected = [id]; state.focusCell = id; 
  state.inputBuffer = state.cells[id]?.input || '';
  state.isEditing = true;
  renderCells();
}

function commitInput() {
  if (state.focusCell) {
    if (!state.cells[state.focusCell]) state.cells[state.focusCell] = {};
    state.cells[state.focusCell].input = state.inputBuffer;
    state.cells[state.focusCell].display = calculateDisplay(state.inputBuffer);
    state.isEditing = false;
  }
  recomputeAllFormulas(); renderCells();
}

formulaBar.addEventListener('keydown', (e) => { if (e.key === 'Enter') commitInput(); });

function calculateDisplay(inputStr) {
  if (!inputStr.startsWith('=')) return inputStr;
  try {
    let formula = inputStr.substring(1).toUpperCase();
    formula = formula.replace(/π/g, 'Math.PI').replace(/√/g, 'Math.sqrt');
    formula = formula.replace(/[A-Z]+[0-9]+/g, (match) => {
      const cellData = state.cells[match];
      return cellData && !isNaN(parseFloat(cellData.display)) ? parseFloat(cellData.display) : 0;
    });
    return eval(formula);
  } catch(e) { return 'ERR'; }
}

function recomputeAllFormulas() {
  for (let pass = 0; pass < 2; pass++) {
    for (const id in state.cells) {
      if (state.cells[id].input?.startsWith('=')) {
        state.cells[id].display = calculateDisplay(state.cells[id].input);
      }
    }
  }
}

function checkReferenceMode() { state.referenceMode = state.inputBuffer.startsWith('='); }

// --- FILE OPERATIONS ---
btnOpen.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
        const workbook = XLSX.read(evt.target.result, { type: 'binary' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        state.cells = {};
        json.forEach((row, ri) => {
            row.forEach((val, ci) => {
                if (val != null) {
                    const id = getCellId(ci, ri);
                    state.cells[id] = { input: val.toString(), display: val.toString() };
                }
            });
        });
        recomputeAllFormulas(); renderCells();
    };
    reader.readAsBinaryString(file);
});

function exportFile(format) {
    const data = [];
    for (let r = 0; r < NUM_ROWS; r++) {
        const row = [];
        for (let c = 0; c < NUM_COLS; c++) {
            const id = getCellId(c, r);
            row.push(state.cells[id] ? state.cells[id].display : '');
        }
        data.push(row);
    }
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    XLSX.writeFile(wb, `puchitto_export.${format === 'xlsx' ? 'xlsx' : 'csv'}`, { bookType: format });
}
btnExportExcel.addEventListener('click', () => exportFile('xlsx'));
btnExportCsv.addEventListener('click', () => exportFile('csv'));

// --- COPY & PASTE (Relative Reference) ---
btnCopy.addEventListener('click', () => {
    if (state.focusCell) {
        state.clipboard = { data: JSON.parse(JSON.stringify(state.cells[state.focusCell] || { input: '', display: '' })), sourceId: state.focusCell };
    }
});
btnPaste.addEventListener('click', () => {
    if (state.focusCell && state.clipboard) {
        saveHistory();
        const input = state.clipboard.data.input;
        state.cells[state.focusCell] = {
            input: input.startsWith('=') ? shiftFormula(input, state.clipboard.sourceId, state.focusCell) : input,
            display: ''
        };
        recomputeAllFormulas(); renderCells();
    }
});

function shiftFormula(formula, srcId, dstId) {
    const srcC = getCoords(
