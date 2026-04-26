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
      if (state.selected.includes(id)) el.classList.add(state.selected.length === 1 ? 'selected' : 'multi-selected');
      if (id === state.focusCell && state.isEditing) el.classList.add('editing');
      const cellData = state.cells[id] || { display: '' };
      if (id === state.focusCell && state.isEditing) el.innerHTML = state.inputBuffer + '<span class="cursor"></span>';
      else el.textContent = cellData.display;
  });

  let activeHint = null;
  if (state.isEditing && state.inputBuffer.startsWith('=')) {
      const match = state.inputBuffer.match(/(\b[A-Z]+|√)\([^)]*$/);
      if (match) {
          const funcHints = {
              'SUM': '合計: =SUM(A1:B2)', 'AVERAGE': '平均: =AVERAGE(A1:B2)',
              'MAX': '最大: =MAX(A1:B2)', 'MIN': '最小: =MIN(A1:B2)',
              'COUNT': '個数: =COUNT(A1:B2)', 'ROUND': '四捨五入: =ROUND(3.14, 0)',
              'SIN': 'サイン: =SIN(角度*π/180)', 'COS': 'コサイン: =COS(角度*π/180)', 
              'TAN': 'タンジェント: =TAN(角度*π/180)',
              'LOG': 'ログ: =LOG(100)', 'ABS': '絶対値: =ABS(-5)', 
              'EXP': '指数: =EXP(2)', '√': 'ルート: =√(9)'
          };
          activeHint = funcHints[match[1]];
      } else if (state.inputBuffer === '=') activeHint = '数式を入力してください (例: =A1+B1)';
  }
  formulaHintEl.textContent = activeHint || (state.isEditing ? '数式の入力が可能です（=A1*2 など）' : 'セルを選択して入力を開始してください');

  if (state.focusCell) {
      const selectedCellData = state.cells[state.focusCell] || { input: '' };
      formulaBar.value = state.isEditing ? state.inputBuffer : selectedCellData.input;
  }
  updateVariableKey();
}

function updateVariableKey() {
    vKeys.forEach(k => { if(k){k.textContent = ''; k.dataset.key = ''; k.classList.remove('active');} });
    if (!state.isEditing) return;
    const buf = state.inputBuffer;
    let suggestions = [];
    const funcMatch = buf.match(/([A-z]+|√)\([^)]*$/);
    const lastFunc = funcMatch ? funcMatch[1].toUpperCase() : null;

    if (lastFunc) {
        if (['SIN', 'COS', 'TAN'].includes(lastFunc)) {
            if (buf.endsWith('(')) suggestions = ['30', 'π', ')'];
            else if (buf.match(/[0-9]$/)) suggestions = ['*π/180)', '*', '/'];
            else if (buf.match(/[π]$/)) suggestions = ['*', '/', ')'];
            else if (buf.match(/[*\/]$/)) suggestions = ['π', '180', ')'];
            else suggestions = ['*', '/', ')'];
        } else if (['SUM', 'AVERAGE', 'MAX', 'MIN', 'COUNT'].includes(lastFunc)) {
            if (buf.endsWith('(')) suggestions = ['A1', 'B1', ')'];
            else if (buf.match(/[0-9]$/)) suggestions = [':', ',', ')'];
            else if (buf.endsWith(':')) suggestions = ['A2', 'C1', 'B2'];
            else suggestions = [':', ')', ','];
        } else suggestions = [')', '+', '*'];
    } else {
        if (buf.match(/[0-9π]$/)) suggestions = ['+', '-', '*'];
        else if (buf.endsWith('=')) suggestions = ['SUM(', 'SIN(', 'COS('];
        else suggestions = [')', '=', '('];
    }

    suggestions.slice(0, 3).forEach((s, i) => {
        if (vKeys[i]) {
            vKeys[i].textContent = s; vKeys[i].dataset.key = s;
            if (s) vKeys[i].classList.add('active');
        }
    });
}

function handleShortPress(cellEl) {
  if (!cellEl) return;
  const id = cellEl.dataset.id;
  if (state.isEditing && state.focusCell && state.focusCell !== id && state.inputBuffer.startsWith('=')) {
      feedback(); state.inputBuffer += id; renderCells(); return;
  }
  if (state.isEditing && state.focusCell && state.focusCell !== id) commitInput();
  state.selected = [id]; state.focusCell = id; 
  state.inputBuffer = (state.cells[id] && state.cells[id].input) || '';
  toggleInputBtn.style.display = 'flex';
  if (state.keyboardMode === 'text') { state.isEditing = true; formulaBar.focus(); }
  else if (state.keyboardMode === 'numeric') { state.isEditing = true; }
  else state.isEditing = false;
  renderCells();
}

function handleLongPress(cellEl) {
  feedback(); const id = cellEl.dataset.id;
  state.selected = [id]; state.focusCell = id; state.isEditing = true;
  state.inputBuffer = (state.cells[id] && state.cells[id].input) || '';
  state.keyboardMode = 'text'; toggleInputBtn.style.display = 'flex'; toggleInputBtn.textContent = '数値入力';
  formulaBar.focus(); customKeyboard.classList.add('hidden'); renderCells();
}

formulaBar.addEventListener('input', (e) => {
  state.inputBuffer = e.target.value; state.isEditing = true;
  updateVariableKey(); renderCells();
});
formulaBar.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') { commitInput(); formulaBar.blur(); state.keyboardMode = 'hidden'; toggleInputBtn.style.display = 'none'; }
});

function calculateDisplay(inputStr) {
  if (!inputStr.startsWith('=')) return inputStr;
  try {
    let formula = inputStr.substring(1);
    formula = formula.replace(/π|Π|pi/gi, 'PI');
    formula = formula.replace(/÷/g, '/').replace(/×/g, '*').replace(/√/g, 'SQRT');
    formula = formula.replace(/([A-Za-z]+[0-9]+):([A-Za-z]+[0-9]+)/g, (match, p1, p2) => resolveRange(p1, p2));
    formula = formula.replace(/[A-Za-z]+[0-9]+/g, (match) => {
      const cellData = state.cells[match.toUpperCase()];
      return cellData && !isNaN(parseFloat(cellData.display)) ? parseFloat(cellData.display) : 0;
    });
    formula = formula.toUpperCase();
    formula = formula.replace(/\^/g, '**');

    const mathContext = `
      const SUM = (...args) => args.reduce((a,b)=>parseFloat(a||0)+parseFloat(b||0), 0);
      const AVERAGE = (...args) => args.length ? SUM(...args)/args.length : 0;
      const MAX = (...args) => Math.max(...args); const MIN = (...args) => Math.min(...args);
      const COUNT = (...args) => args.filter(x => x !== undefined && x !== '').length;
      const ROUND = (v, d=0) => Number(Math.round(v+'e'+d)+'e-'+d);
      const SIN = v => Math.sin(v); const COS = v => Math.cos(v); const TAN = v => Math.tan(v);
      const LOG = Math.log10; const ABS = Math.abs; const EXP = Math.exp;
      const SQRT = Math.sqrt; const PI = Math.PI;
    `;
    const result = new Function(`${mathContext} return ${formula}`)();
    return isNaN(result) ? 'DIV/0' : Math.round(result * 1000) / 1000;
  } catch(e) { return 'ERR'; }
}

function resolveRange(s, e) {
    const colToInt = (l) => {
        let n = 0; for(let i=0; i<l.length; i++) n = n*26+(l.toUpperCase().charCodeAt(i)-64);
        return n-1;
    };
    let s_upper = s.toUpperCase(), e_upper = e.toUpperCase();
    let c1 = colToInt(s_upper.replace(/[0-9]/g, '')), r1 = parseInt(s_upper.replace(/[A-Z]/g, ''))-1;
    let c2 = colToInt(e_upper.replace(/[0-9]/g, '')), r2 = parseInt(e_upper.replace(/[A-Z]/g, ''))-1;
    let vals = [];
    for(let r=Math.min(r1,r2); r<=Math.max(r1,r2); r++) {
        for(let c=Math.min(c1,c2); c<=Math.max(c1,c2); c++) {
            const cellData = state.cells[getCellId(c, r)];
            vals.push(cellData && !isNaN(parseFloat(cellData.display)) ? parseFloat(cellData.display) : 0);
        }
    }
    return vals.join(',');
}

function recomputeAllFormulas() {
  for (let pass = 0; pass < 3; pass++) {
    let changed = false;
    for (const id in state.cells) {
      if (state.cells[id].input && state.cells[id].input.startsWith('=')) {
        const newD = calculateDisplay(state.cells[id].input);
        if (newD !== state.cells[id].display) { state.cells[id].display = newD; changed = true; }
      }
    }
    if (!changed) break;
  }
}
function checkReferenceMode() { state.referenceMode = state.inputBuffer.startsWith('='); }
function commitInput() {
  if (state.focusCell) {
    if ((state.cells[state.focusCell]?.input || '') !== state.inputBuffer) saveHistory();
    if (!state.cells[state.focusCell]) state.cells[state.focusCell] = {};
    state.cells[state.focusCell].input = state.inputBuffer;
    state.cells[state.focusCell].display = calculateDisplay(state.inputBuffer);
    state.isEditing = false;
  }
  recomputeAllFormulas(); renderCells();
}

toggleInputBtn.addEventListener('click', (e) => {
    feedback();
    state.keyboardMinimized = false; 
    if (state.keyboardMode === 'text') {
        state.keyboardMode = 'numeric'; customKeyboard.classList.remove('hidden');
        kbNormal.classList.remove('hidden'); kbMinimized.classList.add('hidden');
        toggleInputBtn.textContent = '文字入力';
    } else {
        state.keyboardMode = 'text'; formulaBar.focus(); customKeyboard.classList.add('hidden');
        toggleInputBtn.textContent = '数値入力';
    }
});

const keyInfo = {
    'delete-all': 'すべて消去', 'backspace': '1文字消去', 'toggle-fn': 'Fn切替', 'minimize': '一時的に隠す',
    'restore': 'キーボードを再表示', 'switch-hand': '左右入替', '(': '(', ')': ')', 'sum': '合計', 'avr': '平均',
    '+': '+', '-': '-', '/': '/', '*': '*', 'sqrt': '√', 'pi': 'π',
    'undo': '一つ戻す', ':': ':', '=': '=', 'enter': '確定', 'max': '最大', 'min': '最小', 'count': '個数',
    '^': '累乗', '%': '%', 'round': '四捨五入', 'sin': 'sin', 'cos': 'cos', 'tan': 'tan', 'log': 'log', 'abs': 'abs', 'exp': 'exp', ',': ','
};

customKeyboard.addEventListener('pointerenter', (e) => {
    const keyEl = e.target.closest('.kb-key');
    if (keyEl) buttonTooltipEl.textContent = keyInfo[keyEl.dataset.key] || keyEl.textContent;
}, true);

customKeyboard.addEventListener('click', (e) => {
  const keyEl = e.target.closest('.kb-key'); if (!keyEl) return;
  feedback(); const key = keyEl.dataset.key; state.isEditing = true;
  if (!key) return; 

  if (key === 'delete-all') {
      saveHistory(); state.selected.forEach(id => { if (state.cells[id]) delete state.cells[id]; });
      state.inputBuffer = ''; state.isEditing = false; recomputeAllFormulas(); renderCells();
  } else if (key === 'backspace') {
      state.inputBuffer = state.inputBuffer.slice(0, -1); updateVariableKey(); renderCells();
  } else if (key === 'enter') commitInput();
  else if (key === 'minimize') { 
      state.keyboardMinimized = true; kbNormal.classList.add('hidden'); kbMinimized.classList.remove('hidden'); 
  } else if (key === 'restore') {
      state.keyboardMinimized = false; kbNormal.classList.remove('hidden'); kbMinimized.classList.add('hidden');
  } else if (key === 'switch-hand') { state.handedness = state.handedness === 'right' ? 'left' : 'right'; applyHandedness(); }
  else if (key === 'toggle-fn') {
      state.fnActive = !state.fnActive;
      document.querySelector('.fn-btn').classList.toggle('active', state.fnActive);
      document.getElementById('kb-primary-contents').classList.toggle('hidden', state.fnActive);
      document.getElementById('kb-secondary-contents').classList.toggle('hidden', !state.fnActive);
  } else if (['avr','sum','max','min','count','round','sin','cos','tan','log','abs','exp','sqrt'].includes(key)) {
       const f = (key==='avr'?'AVERAGE':(key==='sqrt'?'√':key.toUpperCase())) + '(';
       state.inputBuffer = (state.inputBuffer.startsWith('=')?state.inputBuffer:'=') + f;
       renderCells();
  } else if (key === 'undo') performUndo();
  else if (key.startsWith('arrow-')) moveSelection(key);
  else {
      state.inputBuffer += (key==='pi'?'π':key);
      updateVariableKey(); renderCells();
  }
});

function applyHandedness() {
    const isL = state.handedness === 'left';
    kbNormal.classList.toggle('left-hand', isL);
    toggleInputBtn.classList.toggle('left-hand', isL);
    document.querySelector('.top-bar-main').classList.toggle('left-hand', isL);
}

function moveSelection(a) {
    if (!state.focusCell) return;
    const el = document.getElementById(`cell-${state.focusCell}`);
    let c = parseInt(el.dataset.col), r = parseInt(el.dataset.row);
    if (a === 'arrow-left') c = Math.max(0, c - 1);
    if (a === 'arrow-right') c = Math.min(NUM_COLS - 1, c + 1);
    if (a === 'arrow-up') r = Math.max(0, r - 1);
    if (a === 'arrow-down') r = Math.min(NUM_ROWS - 1, r + 1);
    commitInput();
    const nid = getCellId(c, r); state.focusCell = nid; state.selected = [nid];
    state.inputBuffer = (state.cells[nid] && state.cells[nid].input) || '';
    renderCells();
    document.getElementById(`cell-${nid}`).scrollIntoView({ block: 'nearest', inline: 'nearest' });
}

// --- File & Copy/Paste Functions ---
btnOpen.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
        const data = evt.target.result;
        const workbook = XLSX.read(data, { type: 'binary' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        
        state.cells = {};
        json.forEach((row, ri) => {
            row.forEach((val, ci) => {
                if (val !== undefined && val !== null) {
                    const id = getCellId(ci, ri);
                    state.cells[id] = { input: val.toString(), display: val.toString() };
                }
            });
        });
        recomputeAllFormulas();
        renderCells();
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
    const worksheet = XLSX.utils.aoa_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
    if (format === 'xlsx') {
        XLSX.writeFile(workbook, "puchitto_export.xlsx");
    } else {
        XLSX.writeFile(workbook, "puchitto_export.csv", { bookType: 'csv' });
    }
}
btnExportExcel.addEventListener('click', () => exportFile('xlsx'));
btnExportCsv.addEventListener('click', () => exportFile('csv'));

btnCopy.addEventListener('click', () => {
    if (state.focusCell) {
        state.clipboard = {
            data: JSON.parse(JSON.stringify(state.cells[state.focusCell] || { input: '', display: '' })),
            sourceId: state.focusCell
        };
        feedback();
    }
});

btnPaste.addEventListener('click', () => {
    if (state.focusCell && state.clipboard) {
        saveHistory();
        const targetId = state.focusCell;
        const sourceId = state.clipboard.sourceId;
        const input = state.clipboard.data.input;

        if (input.startsWith('=')) {
            state.cells[targetId] = {
                input: shiftFormula(input, sourceId, targetId),
                display: '' 
            };
        } else {
            state.cells[targetId] = JSON.parse(JSON.stringify(state.clipboard.data));
        }
        recomputeAllFormulas();
        renderCells();
        feedback();
    }
});

function shiftFormula(formula, srcId, dstId) {
    const srcCoord = getCoordsFromId(srcId);
    const dstCoord = getCoordsFromId(dstId);
    const dCol = dstCoord.col - srcCoord.col;
    const dRow = dstCoord.row - srcCoord.row;
    return formula.replace(/([A-Z]+)([0-9]+)/g, (match, colPart, rowPart) => {
        const col = colLabelToInt(colPart);
        const row = parseInt(rowPart) - 1;
        const newCol = col + dCol;
        const newRow = row + dRow;
        if (newCol < 0 || newRow < 0 || newCol >= NUM_COLS || newRow >= NUM_ROWS) return match;
        return getCellId(newCol, newRow);
    });
}

function getCoordsFromId(id) {
    const colPart = id.replace(/[0-9]/g, '');
    const rowPart = id.replace(/[A-Z]/g, '');
    return { col: colLabelToInt(colPart), row: parseInt(rowPart) - 1 };
}

function colLabelToInt(label) {
    let n = 0;
    for (let i = 0; i < label.length; i++) {
        n = n * 26 + (label.charCodeAt(i) - 64);
    }
    return n - 1;
}

initGrid(); renderCells();
document.documentElement.style.setProperty('--col-width', `calc((100vw - 32px) / 4)`);
document.documentElement.style.setProperty('--row-height', `calc(((100vw - 32px) / 4) * 0.4)`);
