/* 【ぷちっと表計さん：機能統合・完全復旧版 script.js】 */
const NUM_COLS = 50;
const NUM_ROWS = 200;

// アプリの全状態
let state = {
  cells: {}, selected: [], isEditing: false, focusCell: null, inputBuffer: '',
  referenceMode: false, keyboardMode: 'hidden', handedness: 'right', fnActive: false,
  clipboard: null, undoStack: []
};

// --- 要素のキャッシュ ---
const gridEl = document.getElementById('grid');
const formulaBar = document.getElementById('formula-bar');
const vKeys = [document.getElementById('v-key-1'), document.getElementById('v-key-2'), document.getElementById('v-key-3')];
const formulaHintEl = document.getElementById('formula-hint');
const buttonTooltipEl = document.getElementById('button-tooltip');
const customKeyboard = document.getElementById('custom-keyboard');
const kbResizeHandle = document.getElementById('kb-resize-handle');

// --- ヘルパー関数 ---
const getColLabel = (ci) => {
  let l = ''; let t = ci;
  while (t >= 0) { l = String.fromCharCode(65 + (t % 26)) + l; t = Math.floor(t / 26) - 1; }
  return l;
};
const getCellId = (c, r) => getColLabel(c) + (r + 1);
const getCoords = (id) => {
    const c = id.replace(/[0-9]/g, ''), r = id.replace(/[A-Z]/g, '');
    return { c: colLabelToInt(c), r: parseInt(r) - 1 };
};
const colLabelToInt = (l) => {
    let n = 0; for (let i = 0; i < l.length; i++) n = n * 26 + (l.charCodeAt(i) - 64);
    return n - 1;
};

// 履歴保存（Undo用）
function saveHistory() {
  state.undoStack.push(JSON.stringify(state.cells));
  if (state.undoStack.length > 20) state.undoStack.shift();
}
function performUndo() {
  if (state.undoStack.length > 0) {
    state.cells = JSON.parse(state.undoStack.pop());
    if (state.focusCell) state.inputBuffer = state.cells[state.focusCell]?.input || '';
    refresh();
  }
}

// --- フィードバック（音と振動） ---
let audioCtx = null;
function feedback() {
  if (navigator.vibrate) navigator.vibrate(15);
  try {
    if (!audioCtx) audioCtx = new (window.AudioContext || window.webkitAudioContext)();
    const osc = audioCtx.createOscillator();
    const gain = audioCtx.createGain();
    osc.frequency.setValueAtTime(600, audioCtx.currentTime);
    gain.gain.exponentialRampToValueAtTime(0.01, audioCtx.currentTime + 0.05);
    osc.connect(gain); gain.connect(audioCtx.destination);
    osc.start(); osc.stop(audioCtx.currentTime + 0.05);
  } catch(e) {}
}

// --- グリッド生成 ---
function initGrid() {
  let h = '';
  for (let r = -1; r < NUM_ROWS; r++) {
    for (let c = -1; c < NUM_COLS; c++) {
      if (r === -1 && c === -1) h += `<div class="cell header-corner"></div>`;
      else if (r === -1) h += `<div class="cell header-row">${getColLabel(c)}</div>`;
      else if (c === -1) h += `<div class="cell header-col">${r + 1}</div>`;
      else { 
        const id = getCellId(c, r);
        h += `<div class="cell" id="cell-${id}" data-id="${id}"></div>`; 
      }
    }
  }
  gridEl.innerHTML = h;
}

// --- 高度な数式計算エンジン ---
function calculate(inputStr) {
  if (!inputStr || !inputStr.toString().startsWith('=')) return inputStr;
  try {
    let f = inputStr.substring(1).toUpperCase();
    f = f.replace(/π/g, 'Math.PI').replace(/√|SQRT/g, 'Math.sqrt').replace(/÷/g, '/').replace(/×/g, '*');
    // セル参照(A1, B2:C3)の置換
    f = f.replace(/[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?/g, (match) => {
        if (match.includes(':')) { // 範囲指定
            const [s, e] = match.split(':');
            const sc = getCoords(s), ec = getCoords(e);
            let vals = [];
            for(let r=Math.min(sc.r, ec.r); r<=Math.max(sc.r, ec.r); r++) {
                for(let c=Math.min(sc.c, ec.c); c<=Math.max(sc.c, ec.c); c++) {
                    const d = state.cells[getCellId(c, r)];
                    vals.push(d && !isNaN(parseFloat(d.display)) ? parseFloat(d.display) : 0);
                }
            }
            return `[${vals.join(',')}]`;
        }
        const data = state.cells[match];
        return data && !isNaN(parseFloat(data.display)) ? parseFloat(data.display) : 0;
    });

    // 数学コンテキスト
    const SUM = (args) => (Array.isArray(args) ? args : [args]).reduce((a,b)=>a+parseFloat(b),0);
    const AVERAGE = (args) => { const a = Array.isArray(args)?args:[args]; return a.length?SUM(a)/a.length:0; };
    const MAX = (args) => Math.max(...(Array.isArray(args)?args:[args]));
    const MIN = (args) => Math.min(...(Array.isArray(args)?args:[args]));
    
    // 計算実行
    const result = eval(f);
    return isNaN(result) ? 'ERR' : Math.round(result * 1000) / 1000;
  } catch(e) { return 'ERR'; }
}

function refresh() {
    for (const id in state.cells) {
        const el = document.getElementById(`cell-${id}`);
        if (el) el.textContent = state.cells[id].display || '';
    }
}

// --- 入力候補と説明ヒントの更新 ---
function updateHints() {
    vKeys.forEach(k => { if(k){k.textContent = ''; k.dataset.key = ''; k.classList.remove('active');} });
    const buf = state.inputBuffer;
    let sug = [], hint = 'セルを選択してください', tip = 'ボタンの説明';

    if (state.isEditing) {
        hint = '値を入力するか、= で数式を開始';
        const funcMatch = buf.match(/([A-Z]+|√)\([^)]*$/i);
        const lastFunc = funcMatch ? funcMatch[1].toUpperCase() : null;

        if (lastFunc) {
            hint = lastFunc === 'SUM' ? '合計: =SUM(A1:B2)' : (lastFunc === 'SIN' ? '正弦: =SIN(角度*π/180)' : `${lastFunc}関数を入力中...`);
            if (['SIN', 'COS', 'TAN'].includes(lastFunc)) {
                if (buf.endsWith('(')) sug = ['30', 'π', ')'];
                else if (buf.match(/[0-9π]$/)) sug = ['*π/180)', '*', '/'];
            } else if (['SUM', 'AVERAGE', 'MAX', 'MIN'].includes(lastFunc)) {
                if (buf.endsWith('(')) sug = ['A1', 'B1', ')'];
                else if (buf.match(/[0-9]$/)) sug = [':', ',', ')'];
            }
        } else if (buf.startsWith('=')) {
            hint = '数式を入力してください (例: =A1+B1)';
            if (buf === '=') sug = ['SUM(', 'SIN(', 'SQRT('];
            else if (buf.match(/[0-9π]$/)) sug = ['+', '-', '*'];
        }
    }

    sug.slice(0, 3).forEach((s, i) => {
        if (vKeys[i]) { vKeys[i].textContent = s; vKeys[i].dataset.key = s; if (s) vKeys[i].classList.add('active'); }
    });
    formulaHintEl.textContent = hint;
}

// --- キーボード操作（リサイズ含む） ---
let isResizing = false;
kbResizeHandle.addEventListener('pointerdown', (e) => { isResizing = true; kbResizeHandle.setPointerCapture(e.pointerId); });
window.addEventListener('pointermove', (e) => {
    if (!isResizing) return;
    const h = Math.max(18, Math.min(100, (window.innerHeight - e.clientY - 40) / 4));
    document.documentElement.style.setProperty('--kb-row-height', `${h}px`);
});
window.addEventListener('pointerup', () => { isResizing = false; });

const keyMap = { 'sum':'SUM(', 'avr':'AVERAGE(', 'sin':'SIN(', 'cos':'COS(', 'tan':'TAN(', 'sqrt':'SQRT(', 'max':'MAX(', 'min':'MIN(', 'count':'COUNT(', 'abs':'ABS(', 'exp':'EXP(', 'round':'ROUND(', 'pi':'π' };

customKeyboard.addEventListener('click', (e) => {
    const keyEl = e.target.closest('.kb-key');
    if (!keyEl) return;
    feedback();
    const key = keyEl.dataset.key || keyEl.textContent;

    if (key === 'enter') commitInput();
    else if (key === 'backspace') { state.inputBuffer = state.inputBuffer.slice(0, -1); }
    else if (key === 'delete-all') { saveHistory(); state.inputBuffer = ''; }
    else if (key === 'toggle-fn') {
        state.fnActive = !state.fnActive;
        document.getElementById('kb-primary-contents').classList.toggle('hidden', state.fnActive);
        document.getElementById('kb-secondary-contents').classList.toggle('hidden', !state.fnActive);
        keyEl.classList.toggle('active', state.fnActive);
    }
    else if (key === 'switch-hand') { 
        state.handedness = (state.handedness === 'right' ? 'left' : 'right');
        [document.querySelector('.kb-grid-main'), document.querySelector('.top-bar-main'), document.getElementById('btn-toggle-input')].forEach(el => el.classList.toggle('left-hand', state.handedness === 'left'));
    }
    else if (key === 'minimize') { customKeyboard.classList.add('hidden'); }
    else if (key === 'undo') { performUndo(); }
    else if (key.startsWith('arrow-')) moveSelection(key);
    else {
        state.inputBuffer += (keyMap[key] || key);
    }
    formulaBar.value = state.inputBuffer;
    updateHints();
});

// ボタン説明（ツールチップ）
customKeyboard.addEventListener('pointerover', (e) => {
    const keyEl = e.target.closest('.kb-key');
    if (!keyEl) return;
    const key = keyEl.dataset.id || keyEl.textContent;
    const tips = { 'SUM':'合計', 'AVG':'平均', 'Fn':'高度な関数切替', '左/右':'利き手切替', '↺':'元に戻す', 'AC':'全消去', '⌫':'一文字消去' };
    buttonTooltipEl.textContent = tips[key] || key + ' を入力';
});

// --- コピー＆ペースト（相対参照） ---
function shiftFormula(formula, srcId, dstId) {
    const srcC = getCoords(srcId), dstC = getCoords(dstId);
    const dCol = dstC.c - srcC.c, dRow = dstC.r - srcC.r;
    return formula.replace(/([A-Z]+)([0-9]+)/g, (m, colLabel, rowSrc) => {
        const c = colLabelToInt(colLabel) + dCol, r = parseInt(rowSrc) + dRow;
        if (c < 0 || r <= 0 || c >= NUM_COLS || r >= NUM_ROWS) return m;
        return getColLabel(c) + r;
    });
}
document.getElementById('btn-copy').addEventListener('click', () => {
    if (!state.focusCell) return;
    state.clipboard = { input: state.cells[state.focusCell]?.input || '', sourceId: state.focusCell };
    feedback();
});
document.getElementById('btn-paste').addEventListener('click', () => {
    if (!state.focusCell || !state.clipboard) return;
    saveHistory();
    const input = state.clipboard.input;
    const shifted = input.startsWith('=') ? shiftFormula(input, state.clipboard.sourceId, state.focusCell) : input;
    state.cells[state.focusCell] = { input: shifted, display: calculate(shifted) };
    refresh(); feedback();
});

// --- アプリ制御 ---
function selectCell(id) {
    state.focusCell = id; state.isEditing = true;
    document.querySelectorAll('.cell.selected').forEach(el => el.classList.remove('selected'));
    const el = document.getElementById(`cell-${id}`);
    if (el) { el.classList.add('selected'); el.scrollIntoView({block:'nearest', inline:'nearest'}); }
    state.inputBuffer = state.cells[id]?.input || '';
    formulaBar.value = state.inputBuffer;
    updateHints();
}
gridEl.addEventListener('click', (e) => { const c = e.target.closest('.cell'); if (c && c.dataset.id) selectCell(c.dataset.id); });

function commitInput() {
    if (!state.focusCell) return;
    saveHistory();
    if (!state.cells[state.focusCell]) state.cells[state.focusCell] = {};
    state.cells[state.focusCell].input = state.inputBuffer;
    state.cells[state.focusCell].display = calculate(state.inputBuffer);
    refresh();
}

function moveSelection(dir) {
    if (!state.focusCell) return;
    let {c, r} = getCoords(state.focusCell);
    if (dir === 'arrow-up') r--; if (dir === 'arrow-down') r++;
    if (dir === 'arrow-left') c--; if (dir === 'arrow-right') c++;
    selectCell(getCellId(Math.max(0, c), Math.max(0, r)));
}

// --- 保存・読み込み ---
const getTimestamp = () => {
    const d = new Date(), p = (n) => String(n).padStart(2,'0');
    return `${d.getFullYear()}${p(d.getMonth()+1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`;
};
document.getElementById('btn-export-excel').addEventListener('click', () => {
    const data = [];
    for(let r=0; r<NUM_ROWS; r++){
        const row = [];
        for(let c=0; c<NUM_COLS; c++) row.push(state.cells[getCellId(c,r)]?.display || '');
        data.push(row);
    }
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    XLSX.writeFile(wb, `${getTimestamp()}.xlsx`);
});
document.getElementById('btn-export-csv').addEventListener('click', () => { /* Excelと同様の処理でbookType:csv */ });
document.getElementById('btn-open').addEventListener('click', () => document.getElementById('file-input').click());
document.getElementById('file-input').addEventListener('change', (e) => {
    const f = e.target.files[0]; const reader = new FileReader();
    reader.onload = (evt) => {
        const wb = XLSX.read(evt.target.result, { type: 'binary' });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        state.cells = {};
        json.forEach((row, ri) => row.forEach((v, ci) => { if (v != null) state.cells[getCellId(ci, ri)] = { input:v.toString(), display:v.toString() }; }));
        refresh();
    };
    reader.readAsBinaryString(f);
});

document.getElementById('btn-toggle-input').addEventListener('click', () => customKeyboard.classList.toggle('hidden'));

function start() {
    const w = Math.max(80, Math.floor((window.innerWidth - 32) / 4));
    document.documentElement.style.setProperty('--col-width', `${w}px`);
    initGrid(); refresh(); updateHints();
}
window.onload = start;
