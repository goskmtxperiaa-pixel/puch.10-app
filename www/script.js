// 【script.js 更新版】
// これをコピーして GitHub の www/script.js に上書きしてください
const NUM_COLS = 30; 
const NUM_ROWS = 100;

let state = {
  cells: {}, selected: [], focusCell: null, inputBuffer: '', isEditing: false, clipboard: null
};

function getColLabel(ci) {
  let l = ''; let t = ci;
  while (t >= 0) { l = String.fromCharCode(65 + (t % 26)) + l; t = Math.floor(t / 26) - 1; }
  return l;
}
function getCellId(c, r) { return getColLabel(c) + (r + 1); }

const gridEl = document.getElementById('grid');
const formulaBar = document.getElementById('formula-bar');

function initGrid() {
  if(!gridEl) return;
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

gridEl.addEventListener('click', (e) => {
  const cell = e.target.closest('.cell');
  if (!cell || cell.dataset.id === undefined) return;
  document.querySelectorAll('.cell.selected').forEach(el => el.classList.remove('selected'));
  cell.classList.add('selected');
  state.focusCell = cell.dataset.id;
  state.inputBuffer = state.cells[state.focusCell]?.input || '';
  if(formulaBar) formulaBar.value = state.inputBuffer;
});

function commitInput() {
  if (!state.focusCell) return;
  const val = formulaBar.value;
  if (!state.cells[state.focusCell]) state.cells[state.focusCell] = {};
  state.cells[state.focusCell].input = val;
  state.cells[state.focusCell].display = val;
  const el = document.getElementById(`cell-${state.focusCell}`);
  if (el) el.textContent = val;
}

// ファイル名生成関数
function getTimestamp() {
    const now = new Date();
    const pad = (n) => String(n).padStart(2, '0');
    return `${now.getFullYear()}${pad(now.getMonth()+1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
}

// 保存・読込設定
document.getElementById('btn-export-excel')?.addEventListener('click', () => exportFile('xlsx'));
document.getElementById('btn-export-csv')?.addEventListener('click', () => exportFile('csv'));

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
    XLSX.writeFile(wb, `${getTimestamp()}.${format === 'xlsx' ? 'xlsx' : 'csv'}`, { bookType: format });
}

function startApp() {
  const cellW = Math.max(80, Math.floor((window.innerWidth - 32) / 4));
  document.documentElement.style.setProperty('--col-width', `${cellW}px`);
  document.documentElement.style.setProperty('--row-height', `40px`);
  initGrid();
}
window.onload = startApp;
window.commitInput = commitInput; // ボタンからも呼べるように
