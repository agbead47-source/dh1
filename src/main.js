// 화면 연결. 업로드 → 표준화·합산 → 비교 → 검토 → 엑셀 다운로드.
import { parseExcel } from './parse.js';
import { prepare } from './normalize.js';
import { diffOrders } from './diff.js';
import { defaultReason, isFilled, needsReason } from './reasons.js';
import { buildExportRows, buildWorkbook, exportFileName } from './export.js';
import * as ui from './ui/render.js';

const $ = (sel) => document.querySelector(sel);
const el = {
  slots: { prev: $('.drop[data-slot="prev"]'), curr: $('.drop[data-slot="curr"]') },
  swap: $('#swap'),
  sample: $('#sample'),
  summary: $('#summary-line'),
  review: $('#review'),
  empty: $('#empty'),
  mainTitle: $('#main-title'),
  mainSub: $('#main-sub'),
  scopes: document.querySelectorAll('[data-scope="selected"]'),
  shipTo: $('#ship-to'),
  cards: $('#cards'),
  typeChart: $('#chart-type'),
  delayChart: $('#chart-delay'),
  shipChart: $('#chart-ship'),
  chips: $('#chips'),
  list: $('#list'),
  optDelivery: $('#opt-delivery'),
  warn: $('#export-warn'),
  download: $('#download'),
};

const state = {
  files: { prev: null, curr: null }, // { name, orders, stats }
  errors: { prev: '', curr: '' },
  changes: [],
  reasons: new Map(), // 변경 키 → { code, text }
  shipTo: '',
  filter: 'all',
};

// ---------- 계산 ----------

const isChanged = (c) => c.type !== 'same';
const forShipTo = () => state.changes.filter((c) => c.shipTo === state.shipTo);
const isMissing = (c) => needsReason(c.type) && !isFilled(state.reasons.get(c.key));

function shipToOptions() {
  const map = new Map();
  for (const c of state.changes) {
    const o = map.get(c.shipTo) ?? { name: c.shipTo, changed: 0 };
    if (isChanged(c)) o.changed++;
    map.set(c.shipTo, o);
  }
  return [...map.values()].sort((a, b) => a.name.localeCompare(b.name));
}

function recompute() {
  const { prev, curr } = state.files;
  state.changes = prev && curr ? diffOrders(prev.orders, curr.orders) : [];
  state.reasons = new Map(
    state.changes.filter((c) => needsReason(c.type)).map((c) => [c.key, defaultReason(c.type)]),
  );
  state.filter = 'all';
  const options = shipToOptions();
  if (!options.some((o) => o.name === state.shipTo)) {
    // 변경이 가장 많은 납품처를 먼저 보여 준다
    state.shipTo = [...options].sort((a, b) => b.changed - a.changed)[0]?.name ?? '';
  }
  render();
}

// ---------- 그리기 ----------

function render() {
  for (const slot of ['prev', 'curr']) ui.renderSlot(el.slots[slot], state.files[slot], state.errors[slot]);
  ui.renderSummaryLine(el.summary, state.files);
  el.swap.disabled = !(state.files.prev && state.files.curr);

  const ready = Boolean(state.files.prev && state.files.curr);
  el.review.hidden = !ready;
  el.download.disabled = !ready;
  ui.renderEmpty(el.empty, state.files, ready);

  if (!ready) {
    // 빈 상태: 대시보드 틀은 그대로 두고 값만 비운다
    ui.renderShipTo(el.shipTo, [], '');
    ui.renderMainHead(el.mainTitle, el.mainSub, el.scopes, null, state.files);
    ui.renderCards(el.cards, null);
    ui.renderTypeChart(el.typeChart, null);
    ui.renderDelayChart(el.delayChart, null);
    ui.renderShipChart(el.shipChart, null);
    ui.renderExportWarn(el.warn, 0);
    return;
  }
  renderSelection();
}

/** 고객사가 바뀌면 다시 그리는 것 전부. */
function renderSelection() {
  const options = shipToOptions();
  ui.renderShipTo(el.shipTo, options, state.shipTo);
  ui.renderMainHead(el.mainTitle, el.mainSub, el.scopes, state.shipTo, state.files);
  ui.renderShipChart(el.shipChart, options, state.shipTo);
  const delayed = forShipTo()
    .filter((c) => isChanged(c) && c.delayDays > 0)
    .sort((a, b) => b.delayDays - a.delayDays);
  ui.renderDelayChart(el.delayChart, delayed.slice(0, ui.DELAY_TOP_N), delayed.length);
  renderCounts();
  renderList();
}

/** 사유를 고칠 때마다 다시 계산하는 숫자들. 목록은 다시 그리지 않는다(입력 포커스 유지). */
function renderCounts() {
  const all = forShipTo();
  const changed = all.filter(isChanged);
  const missing = changed.filter(isMissing).length;
  ui.renderCards(el.cards, {
    lines: all.reduce((s, c) => s + (c.curr ? c.curr.lines.length : 0), 0),
    changed: changed.length,
    delayed: changed.filter((c) => c.delayDays >= ui.DELAY_ALERT_DAYS).length,
    missing,
  });
  const counts = {};
  for (const c of changed) counts[c.type] = (counts[c.type] ?? 0) + 1;
  ui.renderChips(el.chips, counts, state.filter);
  ui.renderTypeChart(el.typeChart, counts, state.filter);
  ui.renderExportWarn(el.warn, missing);
}

function renderList() {
  const items = state.changes
    .map((change, id) => ({ id, change }))
    .filter(({ change: c }) => c.shipTo === state.shipTo && isChanged(c))
    .filter(({ change: c }) => state.filter === 'all' || c.type === state.filter);
  ui.renderList(el.list, items, state.reasons);
}

// ---------- 파일 읽기 ----------

async function loadFile(slot, file) {
  state.errors[slot] = '';
  try {
    const rawRows = await parseExcel(await file.arrayBuffer());
    state.files[slot] = { name: file.name, ...prepare(rawRows) };
  } catch (err) {
    console.error(err);
    state.files[slot] = null;
    state.errors[slot] = err instanceof Error ? err.message : String(err);
  }
  recompute();
}

for (const [slot, drop] of Object.entries(el.slots)) {
  const input = drop.querySelector('input');
  input.addEventListener('change', () => {
    if (input.files[0]) loadFile(slot, input.files[0]);
    input.value = ''; // 같은 파일을 다시 골라도 change가 나도록
  });
  drop.addEventListener('dragover', (e) => {
    e.preventDefault();
    drop.classList.add('is-over');
  });
  drop.addEventListener('dragleave', () => drop.classList.remove('is-over'));
  drop.addEventListener('drop', (e) => {
    e.preventDefault();
    drop.classList.remove('is-over');
    const file = e.dataTransfer.files[0];
    if (file) loadFile(slot, file);
  });
}

// 드롭 영역 밖에 떨어뜨렸을 때 브라우저가 파일을 열어 버리지 않게 한다
for (const type of ['dragover', 'drop']) window.addEventListener(type, (e) => e.preventDefault());

el.swap.addEventListener('click', () => {
  [state.files.prev, state.files.curr] = [state.files.curr, state.files.prev];
  recompute();
});

el.sample.addEventListener('click', async () => {
  const [prev, curr] = await Promise.all([
    import('../tests/fixtures/prev.json'),
    import('../tests/fixtures/curr.json'),
  ]);
  state.errors = { prev: '', curr: '' };
  state.files.prev = { name: '샘플_이전.xlsx (가짜 데이터)', ...prepare(prev.default) };
  state.files.curr = { name: '샘플_최신.xlsx (가짜 데이터)', ...prepare(curr.default) };
  recompute();
});

// ---------- 검토 화면 ----------

function selectShipTo(name) {
  state.shipTo = name;
  state.filter = 'all';
  renderSelection();
}

el.shipTo.addEventListener('change', () => selectShipTo(el.shipTo.value));

el.shipChart.addEventListener('click', (e) => {
  const bar = e.target.closest('[data-ship]');
  const option = bar && shipToOptions()[Number(bar.dataset.ship)];
  if (option && option.name !== state.shipTo) selectShipTo(option.name);
});

function setFilter(code) {
  state.filter = code;
  renderCounts();
  renderList();
}

el.chips.addEventListener('click', (e) => {
  const chip = e.target.closest('[data-filter]');
  if (chip) setFilter(chip.dataset.filter);
});

// 유형 차트의 막대는 칩과 같은 필터다. 켜진 막대를 다시 누르면 전체로 돌아간다
el.typeChart.addEventListener('click', (e) => {
  const bar = e.target.closest('[data-filter]');
  if (bar) setFilter(bar.dataset.filter === state.filter ? 'all' : bar.dataset.filter);
});

function onReasonInput(e) {
  const row = e.target.closest('.row');
  if (!row) return;
  const change = state.changes[Number(row.dataset.id)];
  const select = row.querySelector('.reason-select');
  const input = row.querySelector('.reason-text');
  if (!change || !select) return;

  const sel = { code: select.value, text: input.value };
  state.reasons.set(change.key, sel);
  input.hidden = sel.code !== 'custom';
  select.classList.toggle('is-missing', !isFilled(sel));
  if (e.target === select && sel.code === 'custom') input.focus();
  renderCounts();
}
el.list.addEventListener('change', onReasonInput);
el.list.addEventListener('input', onReasonInput);

// ---------- 다운로드 ----------

el.download.addEventListener('click', async () => {
  el.download.disabled = true;
  try {
    const rows = buildExportRows({ shipTo: state.shipTo, changes: state.changes, reasons: state.reasons });
    const wb = buildWorkbook(rows, { includeDeliveryDate: el.optDelivery.checked });
    const blob = new Blob([await wb.xlsx.writeBuffer()], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = exportFileName(state.shipTo);
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  } catch (err) {
    console.error(err);
    el.warn.textContent = `엑셀을 만들지 못했습니다: ${err instanceof Error ? err.message : err}`;
  } finally {
    el.download.disabled = !(state.files.prev && state.files.curr);
  }
});

render();
