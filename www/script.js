// 【最終版 script.js】
// すべてコピーして www/script.js に貼り付けてください
const NUM_COLS = 50;
const NUM_ROWS = 200;

// Application State
let state = {
  cells: {}, 
  selected: [], 
  isEditing: false, 
  focusCell: null, 
  inputBuffer: '', 
  referenceMode: false, 
  keyboardMode: 'hidden', 
  handedness: 'right', 
  keyboardMinimized: false,
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
    if (state.keyboardMode === 'text') formulaBar.value = state.inputBuffer;
  }
}

// Elements
const gridEl = document.getElementById('grid');
const formulaBar = document.getElementById('formula-bar'); 
const toggleInputBtn = document.getElementById('btn-toggle-input');
const customKeyboard = document.getElementById('custom-keyboard');
const kbNormal = document.getElementById('kb-normal');
const kbMinimized = document.getElementById('kb-minimized');
const kbHandle = document.getElementById('kb-handle');
const formulaHintEl = document.getElementById('formula-hint');
const buttonTooltipEl = document.getElementById('button-tooltip');
const kbResizeHandle = document.getElementById('kb-resize-handle');
const vKeys = [
    document.getElementById('v-key-1'),
    document.getElementById('v-key-2'),
    document.getElementById('v-key-3')
];
const btnOpen = document.getElementById('btn-open');
const btnCopy = document.getElementById('btn-copy');
const btnPaste = document.getElementById('btn-paste');
const fileInput = document.getElementById('file-input');
const btnExportCsv = document.getElementById('btn-export-csv');
const btnExportExcel = document.getElementById('btn-export-excel');

// --- Keyboard Height Resize Logic ---
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
    if (isResizing) {
        isResizing = false;
        document.body.style.cursor = '';
        localStorage.setItem('kb-row-height', getComputedStyle(document.documentElement).getPropertyValue('--kb-row-height'));
    }
});

const savedHeight = localStorage.getItem('kb-row-height');
if (savedHeight) {
    document.documentElement.style.setProperty('--kb-row-height', savedHeight);
}

// Vibrate & Sound
let audioCtx = null;
function feedback() {
  if (navigator.vibrate) navigator.vibrate(15);
  try {
    if (!audioCtx) audioCtx = new (window.AudioContext || window.webkitAudioContext)();
    if (audioCtx.state === 'suspended') audioCtx.resume();
    const osc = audioCtx.createOscillator();
    const gain = audioCtx.createGain();
    osc.type = 'triangle';
    osc.frequency.setValueAtTime(600, audioCtx.currentTime);
    osc.frequency.exponentialRampToValueAtTime(200, audioCtx.currentTime + 0.05);
    gain.gain.setValueAtTime(0.3, audioCtx.currentTime);
    gain.gain.exponentialRampToValueAtTime(0.01, audioCtx.currentTime + 0.05);
    osc.connect(gain); gain.connect(audioCtx.destination);
    osc.start(); osc.stop(audioCtx.currentTime + 0.05);
  } catch(e) {}
}

function getColLabel(colIndex) {
  let label = ''; let temp = colIndex;
  while (temp >= 0) {
    label = String.fromCharCode(65 + (temp % 26)) + label;
    temp = Math.floor(temp / 26) - 1;
  }
  return label;
}
function getCellId(col, row) { return getColLabel(col) + (row + 1); }

function initGrid() {
  let html = '';
  for (let r = -1; r < NUM_ROWS; r++) {
    for (let c = -1; c < NUM_COLS; c++) {
      if (r === -1 && c === -1) html += `<div class="cell header-corner"></div>`;
      else if (r === -1) html += `<div class="cell header-row">${getColLabel(c)}</div>`;
      else if (c === -1) html += `<div class="cell header-col">${r + 1}</div>`;
      else {
        const id = getCellId(c, r);
        html += `<div class="cell" id="cell-${id}" data-col="${c}" data-row="${r}" data-id="${id}"></div>`;
      }
    }
  }
  gridEl.innerHTML = html;
  
  let longPressTimer; let isDragging = false; let startDragCell = null; let startX = 0, startY = 0;
  gridEl.addEventListener('pointerdown', (e) => {
    const cell = e.target.closest('.cell');
    if (!cell || cell.classList.contains('header-row') || cell.classList.contains('header-col') || cell.classList.contains('header-corner')) return;
    startDragCell = cell; isDragging = false; startX = e.clientX; startY = e.clientY;
    longPressTimer = setTimeout(() => {
      longPressTimer = null; state.pendingLongPress = cell; cell.classList.add('editing');
      if (navigator.vibrate) navigator.vibrate(20);
    }, 500);
  });
  gridEl.addEventListener('pointermove', (e) => {
    if (startDragCell && !isDragging) {
      if (Math.abs(e.clientX - startX) > 5 || Math.abs(e.clientY - startY) > 5) {
        clearTimeout(longPressTimer); longPressTimer = null;
        if (state.pendingLongPress) { state.pendingLongPress.classList.remove('editing'); state.pendingLongPress = null; }
        isDragging = true;
      }
    }
    if (isDragging) e.preventDefault();
  });
  gridEl.addEventListener('pointerup', (e) => {
    const cell = e.target.closest('.cell');
    if (longPressTimer) {
      clearTimeout(longPressTimer); longPressTimer = null;
      if (!isDragging) handleShortPress(cell);
    } else if (state.pendingLongPress) {
      handleLongPress(state.pendingLongPress); state.pendingLongPress = null;
    } 
    startDragCell = null; isDragging = false;
  });
}

function renderCells() {
  document.querySelectorAll('.cell.selected').forEach(el => el.classList.remove('selected', 'editing'));
  document.querySelectorAll('.cell.multi-selected').forEach(el => el.classList.remove('multi-selected'));
  const allIds = new Set([...state.selected, ...Object.keys(state.cells)]);
  if (state.focusCell) allIds.add(state.focusCell);
  allIds.forEach(id => {
      const el = document.getElementById(`cell-${id}`);
      if (!el) return;
      if (state.selected.includes(id)) el.classList.add(state.selected.length === 1 ? 'selected
