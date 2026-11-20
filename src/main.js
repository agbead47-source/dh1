// 화면 연결. 업로드 → 표준화·합산 → 비교 → Pending Order 표에서 검토 → 엑셀 다운로드.
import { parseWorkbook } from './parse.js';
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
  steps: {
    upload: $('[data-step="upload"]'),
    review: $('[data-step="review"]'),
    export: $('[data-step="export"]'),
  },
  sideProgress: $('#side-progress'),
  review: $('#review'),
  empty: $('#empty'),
  mainTitle: $('#main-title'),
  mainSub: $('#main-sub'),
  scopes: document.querySelectorAll('[data-scope="selected"]'),
  shipTo: $('#ship-to'),
  cards: $('#cards'),
  chartsFold: $('#charts-fold'),
  typeChart: $('#chart-type'),
  delayChart: $('#chart-delay'),
  shipChart: $('#chart-ship'),
  views: $('#views'),
  chips: $('#chips'),
  reasonTools: $('#reason-tools'),
  table: $('#po-table'),
  rows: $('#po-rows'),
  cancelled: $('#po-cancelled'),
  shown: $('#shown'),
  gone: $('#gone'),
  goneTitle: $('#gone-title'),
  goneList: $('#gone-list'),
  optDelivery: $('#opt-delivery'),
  warn: $('#export-warn'),
  download: $('#download'),
  exportFile: $('#export-file'),
};

const state = {
  files: { prev: null, curr: null }, // { name, source, orders, stats }
  errors: { prev: '', curr: '' },
  changes: [],
  idByKey: new Map(), // 변경 키 → changes 순번
  reasons: new Map(), // 변경 키 → { code, text }
  shipTo: '',
  view: 'changed', // 'all' 엑셀에 나가는 전체 줄 | 'changed' 변경된 주문만 | 'missing' 사유 미입력만
  filter: 'all', // 유형 필터. 보기와 함께 걸린다
  downloaded: new Set(), // 지금 내용 그대로 엑셀을 내려받은 납품처
};

// ---------- 계산 ----------

const isReady = () => Boolean(state.files.prev && state.files.curr);
const isChanged = (c) => c.type !== 'same';
const forShipTo = () => state.changes.filter((c) => c.shipTo === state.shipTo);
const isMissing = (c) => needsReason(c.type) && !isFilled(state.reasons.get(c.key));

/** 화면의 표와 내려받는 엑셀이 같은 행을 쓴다. */
const exportRows = () => buildExportRows({ shipTo: state.shipTo, changes: state.changes, reasons: state.reasons });

function shipToOptions() {
  const map = new Map();
  for (const c of state.changes) {
    const o = map.get(c.shipTo) ?? { name: c.shipTo, changed: 0, missing: 0 };
    if (isChanged(c)) o.changed++;
    if (isMissing(c)) o.missing++;
    map.set(c.shipTo, o);
  }
  return [...map.values()].sort((a, b) => a.name.localeCompare(b.name));
}

function recompute() {
  const { prev, curr } = state.files;
  state.changes = prev && curr ? diffOrders(prev.orders, curr.orders) : [];
  state.idByKey = new Map(state.changes.map((c, id) => [c.key, id]));
  state.reasons = new Map(
    state.changes.filter((c) => needsReason(c.type)).map((c) => [c.key, defaultReason(c.type)]),
  );
  state.view = 'changed';
  state.filter = 'all';
  state.downloaded.clear();
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
  const ready = isReady();
  el.swap.disabled = !ready;
  el.review.hidden = !ready;
  el.download.disabled = !ready;
  ui.renderEmpty(el.empty, state.files, state.errors, ready);

  if (!ready) {
    // 빈 상태: 대시보드 틀은 그대로 두고 값만 비운다
    ui.renderShipTo(el.shipTo, [], '');
    ui.renderMainHead(el.mainTitle, el.mainSub, el.scopes, null, state.files);
    ui.renderCards(el.cards, null, state.view);
    ui.renderTypeChart(el.typeChart, null);
    ui.renderDelayChart(el.delayChart, null);
    ui.renderShipChart(el.shipChart, null);
    ui.renderSideProgress(el.sideProgress, null);
    ui.renderExportWarn(el.warn, 0);
    renderSteps(null);
    el.exportFile.textContent = '파일 2개를 올리면 내려받을 수 있습니다.';
    return;
  }
  renderSelection();
}

/** 고객사가 바뀌면 다시 그리는 것 전부. */
function renderSelection() {
  ui.renderMainHead(el.mainTitle, el.mainSub, el.scopes, state.shipTo, state.files);
  ui.renderShipChart(el.shipChart, shipToOptions(), state.shipTo);
  const delayed = forShipTo()
    .filter((c) => isChanged(c) && c.delayDays > 0)
    .sort((a, b) => b.delayDays - a.delayDays);
  ui.renderDelayChart(el.delayChart, delayed.slice(0, ui.DELAY_TOP_N), delayed.length);
  el.exportFile.textContent = `파일명: ${exportFileName(state.shipTo)}`;
  renderTable();
  renderGone();
}

/** 작업 순서 3단계(올리기 → 검토 → 내려받기)가 어디까지 왔는지. progress가 null이면 아직 비교 전이다. */
function renderSteps(progress) {
  const loaded = ['prev', 'curr'].filter((slot) => state.files[slot]).length;
  const failed = Boolean(state.errors.prev || state.errors.curr);
  const upload = failed
    ? { state: 'error', status: '다시 올려 주세요' }
    : loaded === 2
      ? { state: 'done', status: '완료' }
      : { state: 'current', status: `${loaded} / 2개` };
  let review = { state: 'todo', status: '' };
  let exportStep = { state: 'todo', status: '' };
  if (progress) {
    review = progress.missing
      ? { state: 'current', status: `미입력 ${progress.missing}건` }
      : { state: 'done', status: '완료' };
    if (state.downloaded.has(state.shipTo)) exportStep = { state: 'done', status: '내려받음' };
    else if (!progress.missing) exportStep = { state: 'current', status: '할 차례' };
  }
  ui.renderSteps(el.steps, { upload, review, export: exportStep });
}

/** 사유를 고칠 때마다 다시 계산하는 숫자들. 표는 다시 그리지 않는다(입력 포커스 유지). */
function renderCounts(rows = exportRows()) {
  const all = forShipTo();
  const changed = all.filter(isChanged);
  const progress = {
    need: changed.filter((c) => needsReason(c.type)).length,
    missing: changed.filter(isMissing).length,
  };
  ui.renderCards(
    el.cards,
    {
      lines: rows.length,
      cancelledLines: rows.filter((r) => r.status === 'cancelled').length,
      orders: all.length,
      changed: changed.length,
      delayed: changed.filter((c) => c.delayDays >= ui.DELAY_ALERT_DAYS).length,
      maxDelay: Math.max(0, ...changed.map((c) => c.delayDays)),
      ...progress,
    },
    state.view,
  );
  const counts = {};
  for (const c of changed) counts[c.type] = (counts[c.type] ?? 0) + 1;
  ui.renderViews(el.views, { lines: rows.length, changed: changed.length, missing: progress.missing }, state.view);
  ui.renderChips(el.chips, counts, state.filter);
  ui.renderReasonTools(el.reasonTools, progress);
  ui.renderTypeChart(el.typeChart, counts, state.filter);
  // 드롭다운의 고객사별 미입력 건수도 함께 갱신한다. 고르는 중에는 건드리지 않는다
  if (document.activeElement !== el.shipTo) ui.renderShipTo(el.shipTo, shipToOptions(), state.shipTo);
  ui.renderSideProgress(el.sideProgress, progress);
  renderSteps(progress);
  ui.renderExportWarn(el.warn, progress.missing);
  ui.renderShown(el.shown, el.table.querySelectorAll('tbody tr[data-id]').length, rows.length);

  const gone = all.filter((c) => c.type === 'gone');
  ui.renderGoneTitle(el.goneTitle, gone.length, gone.filter(isMissing).length);
}

/** 표에 그릴 항목. 사유 입력은 주문마다 한 줄에만 둔다: 변경된 첫 줄, 없으면 첫 줄. */
function toItems(rows) {
  const editorRow = new Map();
  for (const r of rows) {
    const first = editorRow.get(r.key);
    if (!first || (first.status !== 'changed' && r.status === 'changed')) editorRow.set(r.key, r);
  }
  const seen = new Set();
  return rows.map((row) => {
    const id = state.idByKey.get(row.key);
    const change = state.changes[id];
    const lead = !seen.has(row.key);
    seen.add(row.key);
    return {
      row,
      id,
      change,
      lead,
      editor: row.status !== 'cancelled' && needsReason(row.type) && editorRow.get(row.key) === row,
      sel: state.reasons.get(row.key),
    };
  });
}

const inView = ({ row, change }) => {
  if (state.filter !== 'all' && row.type !== state.filter) return false;
  if (state.view === 'changed') return isChanged(change);
  if (state.view === 'missing') return row.status !== 'cancelled' && isMissing(change);
  return true;
};

const EMPTY_TEXT = {
  all: '이 조건에 해당하는 줄이 없습니다.',
  changed: '이 조건으로 바뀐 줄이 없습니다. "전체"를 누르면 엑셀에 나가는 모든 줄을 볼 수 있습니다.',
  missing: '사유를 입력하지 않은 줄이 없습니다.',
};

function renderTable() {
  const rows = exportRows();
  const items = toItems(rows).filter(inView);
  ui.renderTable(el.rows, items.filter((i) => i.row.status !== 'cancelled'), EMPTY_TEXT[state.view]);
  ui.renderCancelled(el.cancelled, items.filter((i) => i.row.status === 'cancelled'));
  el.table.classList.toggle('no-delivery', !el.optDelivery.checked);
  renderCounts(rows);
}

/** 이전 파일에만 있는 주문. Pending Order 줄이 아니므로 표 아래에 따로 둔다. */
function renderGone() {
  const items = state.changes
    .map((change, id) => ({ id, change, sel: state.reasons.get(change.key) }))
    .filter(({ change: c }) => c.shipTo === state.shipTo && c.type === 'gone');
  el.gone.hidden = !items.length;
  ui.renderGone(el.goneList, items);
}

// ---------- 파일 읽기 ----------

async function loadFile(slot, file) {
  state.errors[slot] = '';
  try {
    const { source, rows } = await parseWorkbook(await file.arrayBuffer());
    state.files[slot] = { name: file.name, source, ...prepare(rows) };
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
  state.files.prev = { name: '샘플_이전.xlsx (가짜 데이터)', source: 'erp', ...prepare(prev.default) };
  state.files.curr = { name: '샘플_최신.xlsx (가짜 데이터)', source: 'erp', ...prepare(curr.default) };
  recompute();
});

// 빈 상태 안내의 샘플 버튼은 사이드바의 샘플 버튼과 같은 일을 한다
el.empty.addEventListener('click', (e) => {
  if (e.target.closest('[data-action="sample"]')) el.sample.click();
});

// ---------- 검토 화면 ----------

function selectShipTo(name) {
  state.shipTo = name;
  state.view = 'changed';
  state.filter = 'all';
  renderSelection();
}

el.shipTo.addEventListener('change', () => selectShipTo(el.shipTo.value));

el.shipChart.addEventListener('click', (e) => {
  const bar = e.target.closest('[data-ship]');
  const option = bar && shipToOptions()[Number(bar.dataset.ship)];
  if (option && option.name !== state.shipTo) selectShipTo(option.name);
});

function scrollTo(target) {
  const calm = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  target.scrollIntoView({ block: 'start', behavior: calm ? 'auto' : 'smooth' });
}

function setView(view) {
  state.view = view;
  renderTable();
}

function setFilter(code) {
  // 사라진 주문은 표의 줄이 아니다. 표 아래 구역으로 데려간다
  if (code === 'gone') {
    if (!el.gone.hidden) scrollTo(el.gone);
    return;
  }
  state.filter = code;
  renderTable();
}

el.views.addEventListener('click', (e) => {
  const chip = e.target.closest('[data-view]');
  if (chip) setView(chip.dataset.view);
});

el.chips.addEventListener('click', (e) => {
  const chip = e.target.closest('[data-filter]');
  if (chip) setFilter(chip.dataset.filter);
});

// 유형 차트의 막대는 칩과 같은 필터다. 켜진 막대를 다시 누르면 전체로 돌아간다
el.typeChart.addEventListener('click', (e) => {
  const bar = e.target.closest('[data-filter]');
  if (bar) setFilter(bar.dataset.filter === state.filter ? 'all' : bar.dataset.filter);
});

// 요약 카드의 "사유 미입력"을 누르면 유형 필터를 풀고 미입력만 모아 표로 데려간다. 다시 누르면 "변경만"으로
el.cards.addEventListener('click', (e) => {
  if (!e.target.closest('[data-action="missing"]')) return;
  state.filter = 'all';
  setView(state.view === 'missing' ? 'changed' : 'missing');
  if (state.view === 'missing') scrollTo(el.review);
});

/**
 * 사유를 고르거나 직접 입력할 때. 표 전체를 다시 그리지 않고 그 주문의 줄과 숫자만 고친다.
 * 그래서 "사유 미입력" 보기에서 사유를 채워도 그 줄은 보기를 바꿀 때까지 남는다.
 */
function onReasonInput(e) {
  const holder = e.target.closest('[data-id]');
  if (!holder) return;
  const id = holder.dataset.id;
  const change = state.changes[Number(id)];
  const select = holder.querySelector('.reason-select');
  const input = holder.querySelector('.reason-text');
  if (!change || !select) return;

  const sel = { code: select.value, text: input.value };
  state.reasons.set(change.key, sel);
  input.hidden = sel.code !== 'custom';
  select.classList.toggle('is-missing', !isFilled(sel));
  if (e.target === select && sel.code === 'custom') input.focus();
  state.downloaded.delete(state.shipTo); // 내려받은 뒤에 사유를 고치면 다시 받아야 한다

  const rows = exportRows();
  if (change.type === 'gone') {
    // 주문 취소로 고르면 표 맨 아래에 줄이 생기고, 바꾸면 사라진다
    holder.classList.toggle('is-missing', !isFilled(sel));
    holder.querySelector('.gone-note').textContent = ui.goneNote(isFilled(sel) ? sel.code : '');
    const show = state.filter === 'all' && state.view !== 'missing';
    ui.renderCancelled(el.cancelled, show ? toItems(rows).filter((i) => i.row.status === 'cancelled') : []);
  } else {
    const trs = [...el.rows.querySelectorAll(`tr[data-id="${id}"]`)];
    ui.updateOrderRows(trs, rows.filter((r) => r.key === change.key), !isFilled(sel));
  }
  renderCounts(rows);
}

for (const box of [el.rows, el.goneList]) {
  box.addEventListener('change', onReasonInput);
  box.addEventListener('input', onReasonInput);
}

// ---------- 다운로드 ----------

el.optDelivery.addEventListener('change', () => {
  el.table.classList.toggle('no-delivery', !el.optDelivery.checked);
  state.downloaded.delete(state.shipTo);
  if (isReady()) renderCounts();
});

el.download.addEventListener('click', async () => {
  el.download.disabled = true;
  try {
    const wb = buildWorkbook(exportRows(), { includeDeliveryDate: el.optDelivery.checked });
    const blob = new Blob([await wb.xlsx.writeBuffer()], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = exportFileName(state.shipTo);
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
    state.downloaded.add(state.shipTo);
    renderCounts();
  } catch (err) {
    console.error(err);
    el.warn.textContent = `엑셀을 만들지 못했습니다: ${err instanceof Error ? err.message : err}`;
  } finally {
    el.download.disabled = !isReady();
  }
});

render();
