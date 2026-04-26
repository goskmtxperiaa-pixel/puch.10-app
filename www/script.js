// 【www/script.js】
const NUM_COLS = 30; // 動作を軽くするため一時的に減らしています
const NUM_ROWS = 50;

let state = {
  cells: {}, selected: [], focusCell: null, inputBuffer: '', isEditing: false, clipboard: null
};

function getColLabel(ci) {
  let l = ''; let t = ci;
  while (t >= 0) { l = String.fromCharCode(65 + (t % 26)) + l; t = Math.floor(t / 26) - 1; }
  return l;
}

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
        const id = getColLabel(c) + (r + 1);
        h += `<div class="cell" id="cell-${id}" data-id="${id}">${state.cells[id]?.display || ''}</div>`; 
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
  state.cells[state.focusCell].display = val; // 簡易表示
  const el = document.getElementById(`cell-${state.focusCell}`);
  if (el) el.textContent = val;
}

// ボタン設定
const safeAddListener = (id, evt, fn) => {
  const el = document.getElementById(id);
  if (el) el.addEventListener(evt, fn);
};

safeAddListener('btn-toggle-input', 'click', () => {
    document.getElementById('custom-keyboard').classList.toggle('hidden');
});

// 初期起動
function startApp() {
  const cellW = Math.max(80, Math.floor((window.innerWidth - 32) / 4));
  document.documentElement.style.setProperty('--col-width', `${cellW}px`);
  document.documentElement.style.setProperty('--row-height', `40px`);
  initGrid();
}
window.onload = startApp;
