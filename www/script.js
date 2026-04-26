/* 【復元：3/3】 www/script.js (4月24日ベース + 最新機能統合版) */
const NUM_COLS = 50;
const NUM_ROWS = 200;

let state = {
  cells: {}, selected: [], isEditing: false, focusCell: null, inputBuffer: '',
  referenceMode: false, keyboardMode: 'hidden', handedness: 'right', fnActive: false,
  clipboard: null, undoStack: []
};

// 要素のキャッシュ
const gridEl = document.getElementById('grid');
const formulaBar = document.getElementById('formula-bar');
const vKeys = [document.getElementById('v-key-1'), document.getElementById('v-key-2'), document.getElementById('v-key-3')];
const formulaHintEl = document.getElementById('formula-hint');
const buttonTooltipEl = document.getElementById('button-tooltip');
const customKeyboard = document.getElementById('custom-keyboard');
const kbResizeHandle = document.getElementById('kb-resize-handle');

// ヘルパー
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

// フィードバック
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

// グリッド生成
function initGrid() {
  let h = '';
  for (let r = -1; r < NUM_ROWS; r++) {
    for (let c = -1; c < NUM_COLS; c++) {
      if (r === -1 && c === -1) h += `<div class="cell header-corner"></div>`;
      else if (r === -1) h += `<div class="cell header-row">${getColLabel(c)}</div>`;
      else if (c === -1) h += `<div class="cell header-col">${r + 1}</div>`;
      else { const id = getCellId(c, r); h += `<div class="cell" id="cell-${id}" data-id="${id}"></div>`; }
    }
  }
  gridEl.innerHTML = h;
}

// 予測候補のロジック (4月24日仕様)
function updateVKeys() {
    vKeys.forEach(k => { if(k){k.textContent = ''; k.dataset.key = ''; k.classList.remove('active');} });
    const buf = state.inputBuffer;
    let suggestions = [];
    const funcMatch = buf.match(/([A-Z]+|√)\([^)]*$/i);
    const lastFunc = funcMatch ? funcMatch[1].toUpperCase() : null;

    if (lastFunc) {
        if (['SIN', 'COS', 'TAN'].includes(lastFunc)) {
            if (buf.endsWith('(')) suggestions = ['30', 'π', ')'];
            else suggestions = ['*π/180)', '*', '/'];
        } else if (['SUM', 'AVERAGE', 'MAX', 'MIN', 'COUNT'].includes(lastFunc)) {
            if (buf.endsWith('(')) suggestions = ['A1', 'B1', ')'];
            else suggestions = [':', ',', ')'];
        } else suggestions = [')', '+', '*'];
    } else {
        if (buf.match(/[0-9π]$/)) suggestions = ['+', '-', '*'];
        else if (buf.endsWith('=')) suggestions = ['SUM(', 'SIN(', 'SQRT('];
        else suggestions = [')', '=', '('];
    }

    suggestions.slice(0, 3).forEach((s, i) => {
        if (vKeys[i]) { vKeys[i].textContent = s; vKeys[i].dataset.key = s; if (s) vKeys[i].classList.add('active'); }
    });
}

// 計算エンジン
function calculate(inputStr) {
  if (!inputStr || !inputStr.toString().startsWith('=')) return inputStr;
  try {
    let f = inputStr.substring(1).toUpperCase().replace(/π/g, 'Math.PI').replace(/√/g, 'Math.sqrt').replace(/÷/g, '/').replace(/×/g, '*');
    f = f.replace(/[A-Z]+[0-9]+/g, (m) => {
      const d = state.cells[m]; return d && !isNaN(parseFloat(d.display)) ? parseFloat(d.display) : 0;
    });
    const SUM = (...a) => a.flat().reduce((s,v)=>s+parseFloat(v||0),0);
    const AVERAGE = (...a) => { const x=a.flat(); return x.length?SUM(x)/x.length:0; };
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

// 相対参照シフト
function shiftFormula(formula, srcId, dstId) {
    const srcC = getCoords(srcId), dstC = getCoords(dstId);
    const dCol = dstC.c - srcC.c, dRow = dstC.r - srcC.r;
    return formula.replace(/([A-Z]+)([0-9]+)/g, (m, colLabel, rowSrc) => {
        const c = colLabelToInt(colLabel) + dCol, r = parseInt(rowSrc) + dRow;
        if (c < 0 || r <= 0 || c >= NUM_COLS || r >= NUM_ROWS) return m;
        return getColLabel(c) + r;
    });
}

// キーボードイベント
const keyMap = { 'sum':'SUM(', 'avr':'AVERAGE(', 'sin':'SIN(', 'cos':'COS(', 'tan':'TAN(', 'sqrt':'SQRT(', 'max':'MAX(', 'min':'MIN(', 'count':'COUNT(', 'abs':'ABS(', 'exp':'EXP(', 'round':'ROUND(', 'pi':'π' };

customKeyboard.addEventListener('click', (e) => {
    const keyEl = e.target.closest('.kb-key');
    if (!keyEl) return;
    feedback();
    const key = keyEl.dataset.key || keyEl.textContent;

    if (key === 'enter') commit();
    else if (key === 'backspace') { state.inputBuffer = state.inputBuffer.slice(0, -1); }
    else if (key === 'delete-all') { state.inputBuffer = ''; }
    else if (key === 'toggle-fn') {
        state.fnActive = !state.fnActive;
        document.getElementById('kb-primary-contents').classList.toggle('hidden', state.fnActive);
        document.getElementById('kb-secondary-contents').classList.toggle('hidden', !state.fnActive);
        keyEl.classList.toggle('active', state.fnActive);
    }
    else if (key === 'switch-hand') { 
        state.handedness = (state.handedness === 'right' ? 'left' : 'right');
        [document.querySelector('.kb-grid-main'), document.getElementById('btn-toggle-input')].forEach(el => el.classList.toggle('left-hand', state.handedness==='left'));
    }
    else if (key === 'minimize') { customKeyboard.classList.add('hidden'); }
    else if (key === 'undo') { /* 履歴実装が必要ならここ */ }
    else if (key.startsWith('arrow-')) move(key);
    else { state.inputBuffer += (keyMap[key] || key); }
    
    formulaBar.value = state.inputBuffer;
    updateVKeys();
});

function commit() {
    if (!state.focusCell) return;
    if (!state.cells[state.focusCell]) state.cells[state.focusCell] = {};
    state.cells[state.focusCell].input = state.inputBuffer;
    state.cells[state.focusCell].display = calculate(state.inputBuffer);
    refresh();
}

function move(dir) {
    if (!state.focusCell) return;
    commit();
    let {c, r} = getCoords(state.focusCell);
    if (dir === 'arrow-up') r--; if (dir === 'arrow-down') r++;
    if (dir === 'arrow-left') c--; if (dir === 'arrow-right') c++;
    selectCell(getCellId(Math.max(0, c), Math.max(0, r)));
}

function selectCell(id) {
    state.focusCell = id;
    document.querySelectorAll('.cell.selected').forEach(el => el.classList.remove('selected'));
    const el = document.getElementById(`cell-${id}`);
    if (el) { el.classList.add('selected'); el.scrollIntoView({block:'nearest', inline:'nearest'}); }
    state.inputBuffer = state.cells[id]?.input || '';
    formulaBar.value = state.inputBuffer;
    updateVKeys();
}

gridEl.addEventListener('click', (e) => {
    const c = e.target.closest('.cell');
    if (c && c.dataset.id) selectCell(c.dataset.id);
});

// ファイル名 (西暦秒)
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

document.getElementById('btn-export-csv').addEventListener('click', () => {
    const data = [];
    for(let r=0; r<NUM_ROWS; r++){
        const row = [];
        for(let c=0; c<NUM_COLS; c++) row.push(state.cells[getCellId(c,r)]?.display || '');
        data.push(row);
    }
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    XLSX.writeFile(wb, `${getTimestamp()}.csv`, { bookType: 'csv' });
});

document.getElementById('btn-open').addEventListener('click', () => document.getElementById('file-input').click());
document.getElementById('file-input').addEventListener('change', (e) => {
    const f = e.target.files[0]; if(!f) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
        const wb = XLSX.read(evt.target.result, { type: 'binary' });
        const sheet = wb.Sheets[wb.Sheets[0] || wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        state.cells = {};
        json.forEach((row, ri) => row.forEach((v, ci) => { if(v!=null) state.cells[getCellId(ci,ri)]={input:v.toString(), display:v.toString()}; }));
        refresh();
    };
    reader.readAsBinaryString(f);
});

// コピペ
document.getElementById('btn-copy').addEventListener('click', () => {
    if (!state.focusCell) return;
    state.clipboard = { input: state.cells[state.focusCell]?.input || '', sourceId: state.focusCell };
    feedback();
});
document.getElementById('btn-paste').addEventListener('click', () => {
    if (!state.focusCell || !state.clipboard) return;
    const input = state.clipboard.input;
    const shifted = input.startsWith('=') ? shiftFormula(input, state.clipboard.sourceId, state.focusCell) : input;
    state.cells[state.focusCell] = { input: shifted, display: calculate(shifted) };
    refresh(); feedback();
});

document.getElementById('btn-toggle-input').addEventListener('click', () => {
    customKeyboard.classList.toggle('hidden');
});

// リサイズハンドル
kbResizeHandle.addEventListener('pointerdown', (e) => { kbResizeHandle.setPointerCapture(e.pointerId); });
kbResizeHandle.addEventListener('pointermove', (e) => {
    if (e.buttons !== 1) return;
    const h = Math.max(18, Math.min(100, (window.innerHeight - e.clientY - 40) / 4));
    document.documentElement.style.setProperty('--kb-row-height', `${h}px`);
});

// 起動
window.onload = () => {
    const w = Math.max(80, Math.floor((window.innerWidth - 32) / 4));
    document.documentElement.style.setProperty('--col-width', `${w}px`);
    initGrid(); refresh(); updateVKeys();
};
