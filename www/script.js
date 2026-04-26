/* 【www/script.js】 */
const NUM_COLS = 50;
const NUM_ROWS = 200;

// アプリの状態管理
let state = {
    cells: {},          // セルデータ {A1: {input: "=1+1", display: "2"}}
    selected: null,     // 現在選択中のセルID (例: "A1")
    isEditing: false, 
    clipboard: null     // コピーしたセルの情報 {input: "...", sourceId: "A1"}
};

// --- 初期化処理 ---
const gridEl = document.getElementById('grid');
const formulaBar = document.getElementById('formula-bar');

function initGrid() {
    let html = '';
    for (let r = -1; r < NUM_ROWS; r++) {
        for (let c = -1; c < NUM_COLS; c++) {
            if (r === -1 && c === -1) html += `<div class="cell header-corner"></div>`;
            else if (r === -1) html += `<div class="cell header-row">${getColLabel(c)}</div>`;
            else if (c === -1) html += `<div class="cell header-col">${r + 1}</div>`;
            else {
                const id = getCellId(c, r);
                html += `<div class="cell" id="cell-${id}" data-id="${id}"></div>`;
            }
        }
    }
    gridEl.innerHTML = html;
}

// 列記号(A, B...Z, AA...)を作る
function getColLabel(ci) {
    let l = ''; let t = ci;
    while (t >= 0) { l = String.fromCharCode(65 + (t % 26)) + l; t = Math.floor(t / 26) - 1; }
    return l;
}
function getCellId(c, r) { return getColLabel(c) + (r + 1); }

// --- セル選択・表示 ---
gridEl.addEventListener('click', (e) => {
    const cell = e.target.closest('.cell');
    if (!cell || !cell.dataset.id) return;
    
    // 選択状態の更新
    document.querySelectorAll('.cell.selected').forEach(el => el.classList.remove('selected'));
    cell.classList.add('selected');
    
    state.selected = cell.dataset.id;
    formulaBar.value = state.cells[state.selected]?.input || '';
});

// 入力が確定した時
formulaBar.addEventListener('change', commitInput);

function commitInput() {
    if (!state.selected) return;
    const val = formulaBar.value;
    if (!state.cells[state.selected]) state.cells[state.selected] = {};
    
    state.cells[state.selected].input = val;
    state.cells[state.selected].display = calculate(val);
    
    refreshGrid();
}

function calculate(val) {
    if (!val.startsWith('=')) return val;
    try {
        let formula = val.substring(1).toUpperCase();
        // セル参照(A1など)を数値に置き換え
        formula = formula.replace(/[A-Z]+[0-9]+/g, (m) => {
            const data = state.cells[m];
            return data && !isNaN(data.display) ? data.display : 0;
        });
        // 簡易的な数式計算
        return eval(formula);
    } catch(e) { return 'ERR'; }
}

function refreshGrid() {
    for (const id in state.cells) {
        const el = document.getElementById(`cell-${id}`);
        if (el) el.textContent = state.cells[id].display;
    }
}

// --- コピー＆ペースト（相対参照対応） ---
document.getElementById('btn-copy').addEventListener('click', () => {
    if (!state.selected) return;
    state.clipboard = {
        input: state.cells[state.selected]?.input || '',
        sourceId: state.selected
    };
});

document.getElementById('btn-paste').addEventListener('click', () => {
    if (!state.selected || !state.clipboard) return;
    const input = state.clipboard.input;
    const shiftedInput = input.startsWith('=') ? shiftFormula(input, state.clipboard.sourceId, state.selected) : input;
    
    if (!state.cells[state.selected]) state.cells[state.selected] = {};
    state.cells[state.selected].input = shiftedInput;
    state.cells[state.selected].display = calculate(shiftedInput);
    
    refreshGrid();
    formulaBar.value = shiftedInput;
});

function shiftFormula(formula, srcId, dstId) {
    const srcC = decodeId(srcId), dstC = decodeId(dstId);
    const dC = dstC.c - srcC.c, dR = dstC.r - srcC.r;
    return formula.replace(/([A-Z]+)([0-9]+)/g, (m, colLabel, rowSrc) => {
        const c = colLabelToInt(colLabel) + dC;
        const r = parseInt(rowSrc) + dR;
        return (c >= 0 && r > 0) ? getColLabel(c) + r : m;
    });
}
function decodeId(id) {
    const c = id.replace(/[0-9]/g, ''), r = id.replace(/[A-Z]/g, '');
    return { c: colLabelToInt(c), r: parseInt(r) - 1 };
}
function colLabelToInt(l) {
    let n = 0; for (let i = 0; i < l.length; i++) n = n * 26 + (l.charCodeAt(i) - 64);
    return n - 1;
}

// --- 保存・読み込み（西暦秒ファイル名） ---
function getTimestamp() {
    const d = new Date();
    const pad = (n) => String(n).padStart(2, '0');
    return `${d.getFullYear()}${pad(d.getMonth()+1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
}

document.getElementById('btn-export-excel').addEventListener('click', () => saveFile('xlsx'));
document.getElementById('btn-export-csv').addEventListener('click', () => saveFile('csv'));

function saveFile(format) {
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
    XLSX.writeFile(wb, `${getTimestamp()}.${format}`, { bookType: format });
}

document.getElementById('btn-open').addEventListener('click', () => document.getElementById('file-input').click());
document.getElementById('file-input').addEventListener('change', (e) => {
    const file = e.target.files[0];
    const reader = new FileReader();
    reader.onload = (evt) => {
        const workbook = XLSX.read(evt.target.result, { type: 'binary' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        state.cells = {};
        json.forEach((row, ri) => {
            row.forEach((v, ci) => {
                if (v != null) {
                    const id = getCellId(ci, ri);
                    state.cells[id] = { input: v.toString(), display: v.toString() };
                }
            });
        });
        refreshGrid();
    };
    reader.readAsBinaryString(file);
});

// --- キーボードと表示の調整 ---
document.getElementById('btn-toggle-input').addEventListener('click', () => {
    document.getElementById('custom-keyboard').classList.toggle('hidden');
});

function closeKeyboard() { document.getElementById('custom-keyboard').classList.add('hidden'); }
function commitInputFromKb() { commitInput(); closeKeyboard(); }

// アプリ開始
function start() {
    // スマホの幅に合わせてセル幅を調整
    const w = Math.max(80, Math.floor((window.innerWidth - 32) / 4));
    document.documentElement.style.setProperty('--col-width', `${w}px`);
    initGrid();
    refreshGrid();
}
window.onload = start;
