// 화면 연결. 업로드 → 표준화·합산 → 비교 → Pending Order 표에서 검토 → 엑셀 다운로드.
import { parseWorkbook } from './parse.js';
import { prepare, normalizeShipped } from './normalize.js';
import { diffOrders, compareChange } from './diff.js';
import { defaultReason, isFilled, needsReason } from './reasons.js';
import { buildExportRows, buildShippedRows, buildWorkbook, exportFileName, fileLabelFor } from './export.js';
import { buildMailDraft, DEFAULT_DELAY_DAYS } from './mail.js';
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
  tabs: $('#tabs'),
  panes: { pending: $('#pane-pending'), shipped: $('#pane-shipped'), gone: $('#pane-gone'), mail: $('#pane-mail') },
  mail: {
    recipient: $('#mail-recipient'),
    sender: $('#mail-sender'),
    days: $('#mail-days'),
    subject: $('#mail-subject'),
    body: $('#mail-body'),
    summary: $('#mail-summary'),
    warn: $('#mail-warn'),
    status: $('#mail-status'),
    copySubject: $('#copy-subject'),
    copyBody: $('#copy-body'),
  },
  goneEmpty: $('#gone-empty'),
  compare: $('#compare'),
  compareBody: $('#compare-body'),
  shTable: $('#sh-table'),
  shRows: $('#sh-rows'),
  shShown: $('#sh-shown'),
  gone: $('#gone'),
  goneTitle: $('#gone-title'),
  goneList: $('#gone-list'),
  optDelivery: $('#opt-delivery'),
  optStyle: $('#opt-style'),
  optShipped: $('#opt-shipped'),
  optLabel: $('#opt-label'),
  warn: $('#export-warn'),
  download: $('#download'),
  exportFile: $('#export-file'),
};

// 고객사 선택의 "전체". 실제 일정표처럼 모든 납품처를 한 파일로 낸다.
const ALL = '__ALL__';

const state = {
  files: { prev: null, curr: null }, // { name, source, shipped, orders, stats }
  errors: { prev: '', curr: '' },
  changes: [],
  idByKey: new Map(), // 변경 키 → changes 순번
  reasons: new Map(), // 변경 키 → { code, text }
  shipTo: ALL,
  tab: 'pending', // 보고 있는 탭: 'pending' | 'shipped' | 'gone' | 'mail'
  view: 'changed', // 'all' 엑셀에 나가는 전체 줄 | 'changed' 변경된 주문만 | 'missing' 사유 미입력만
  filter: 'all', // 유형 필터. 보기와 함께 걸린다
  downloaded: new Set(), // 지금 내용 그대로 엑셀을 내려받은 납품처
};

// ---------- 계산 ----------

const isReady = () => Boolean(state.files.prev && state.files.curr);
const isChanged = (c) => c.type !== 'same';
const inScope = (c) => state.shipTo === ALL || c.shipTo === state.shipTo;
const forShipTo = () => state.changes.filter(inScope);
/** 출력 함수에 넘기는 납품처. 전체면 null. */
const scope = () => (state.shipTo === ALL ? null : state.shipTo);
const scopeName = () => (state.shipTo === ALL ? '전체 고객사' : state.shipTo);

/** 파일명에 넣을 이름. 직접 적은 값이 있으면 그것, 없으면 납품처나 매출처에서 만든다. */
const autoLabel = () =>
  fileLabelFor({ shipTo: scope(), soldTo: state.files.curr?.stats.soldTo || state.files.prev?.stats.soldTo });
const label = () => el.optLabel.value.trim() || autoLabel();
const fileName = () => exportFileName(label());
const isMissing = (c) => needsReason(c.type) && !isFilled(state.reasons.get(c.key));

/** 화면의 표와 내려받는 엑셀이 같은 행을 쓴다. */
const exportRows = () => buildExportRows({ shipTo: scope(), changes: state.changes, reasons: state.reasons });

/** Shipped Orders 줄. 이전 파일이 일정표면 그 출하 기록을 넘겨받는다(최신 파일의 것은 쓰지 않는다). */
const shippedRows = () =>
  buildShippedRows({
    shipTo: scope(),
    changes: state.changes,
    reasons: state.reasons,
    shipped: state.files.prev?.shipped ?? [],
  });

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

/** 드롭다운 목록: 맨 위에 전체, 그 아래 납품처. */
function selectOptions() {
  const options = shipToOptions();
  const sum = (key) => options.reduce((s, o) => s + o[key], 0);
  return [{ name: ALL, label: '전체 (한 파일로)', changed: sum('changed'), missing: sum('missing') }, ...options];
}

function recompute() {
  const { prev, curr } = state.files;
  state.changes = prev && curr ? diffOrders(prev.orders, curr.orders) : [];
  state.idByKey = new Map(state.changes.map((c, id) => [c.key, id]));
  state.reasons = new Map(
    state.changes.filter((c) => needsReason(c.type)).map((c) => [c.key, defaultReason(c.type)]),
  );
  state.tab = 'pending';
  state.view = 'changed';
  state.filter = 'all';
  state.downloaded.clear();
  // 실제 일정표가 전체 납품처 한 파일이라 전체를 기본으로 둔다
  if (state.shipTo !== ALL && !shipToOptions().some((o) => o.name === state.shipTo)) state.shipTo = ALL;
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
  ui.renderMainHead(el.mainTitle, el.mainSub, el.scopes, scopeName(), state.files);
  ui.renderShipChart(el.shipChart, shipToOptions(), state.shipTo);
  const delayed = forShipTo()
    .filter((c) => isChanged(c) && c.delayDays > 0)
    .sort((a, b) => b.delayDays - a.delayDays);
  ui.renderDelayChart(el.delayChart, delayed.slice(0, ui.DELAY_TOP_N), delayed.length, state.shipTo === ALL);
  // 좁은 화면의 카드형 표에서 SHIP TO는 전체 고객사를 볼 때만 보인다
  el.table.classList.toggle('one-ship', state.shipTo !== ALL);
  el.optLabel.placeholder = autoLabel();
  el.exportFile.textContent = `파일명: ${fileName()}`;
  renderTable();
  renderGone();
  renderShipped();
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
  const gone = all.filter((c) => c.type === 'gone');
  const goneMissing = gone.filter(isMissing).length;
  ui.renderCards(
    el.cards,
    {
      lines: rows.length,
      cancelledLines: rows.filter((r) => r.status === 'cancelled').length,
      orders: all.length,
      changed: changed.length,
      delayed: changed.filter((c) => c.delayDays >= ui.DELAY_ALERT_DAYS).length,
      maxDelay: Math.max(0, ...changed.map((c) => c.delayDays)),
      goneMissing,
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
  if (document.activeElement !== el.shipTo) ui.renderShipTo(el.shipTo, selectOptions(), state.shipTo);
  ui.renderSideProgress(el.sideProgress, progress);
  renderSteps(progress);
  ui.renderExportWarn(el.warn, progress.missing);
  ui.renderShown(el.shown, el.table.querySelectorAll('tbody tr[data-id]').length, rows.length);
  ui.renderTabs(
    el.tabs,
    { pending: rows.length, shipped: shippedRows().length, gone: gone.length, goneMissing },
    state.tab,
  );
  ui.renderGoneTitle(el.goneTitle, gone.length, goneMissing);
  renderMail();
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

/** Shipped Orders 탭. 보기 전용이라 그 탭을 볼 때만 그린다. */
function renderShipped() {
  for (const [name, pane] of Object.entries(el.panes)) pane.hidden = name !== state.tab;
  if (state.tab !== 'shipped') return;
  const rows = shippedRows();
  ui.renderShipped(el.shRows, rows);
  ui.renderShippedShown(el.shShown, rows);
  el.shTable.classList.toggle('no-delivery', !el.optDelivery.checked);
}

function setTab(tab) {
  state.tab = tab;
  renderShipped();
  renderCounts();
}

/** 최신 파일에 없는 주문. 엑셀의 줄이 아니므로 확인용 탭에 따로 둔다. */
function renderGone() {
  const items = state.changes
    .map((change, id) => ({ id, change, sel: state.reasons.get(change.key) }))
    .filter(({ change: c }) => inScope(c) && c.type === 'gone');
  el.gone.hidden = !items.length;
  el.goneEmpty.hidden = items.length > 0;
  ui.renderGone(el.goneList, items, state.shipTo === ALL);
}

// ---------- 파일 읽기 ----------

async function loadFile(slot, file) {
  state.errors[slot] = '';
  try {
    const { source, rows, shipped } = await parseWorkbook(await file.arrayBuffer());
    state.files[slot] = { name: file.name, source, shipped: normalizeShipped(shipped ?? []), ...prepare(rows) };
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
  state.files.prev = { name: '샘플_이전.xlsx (가짜 데이터)', source: 'erp', shipped: [], ...prepare(prev.default) };
  state.files.curr = { name: '샘플_최신.xlsx (가짜 데이터)', source: 'erp', shipped: [], ...prepare(curr.default) };
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
  // 사라진 주문은 표의 줄이 아니다. 확인용 탭으로 데려간다
  if (code === 'gone') {
    setTab('gone');
    scrollTo(el.tabs);
    return;
  }
  state.filter = code;
  showPending();
  renderTable();
}

/** 보기와 필터는 Pending Order의 것이다. 다른 탭을 보고 있었다면 돌아온다. */
function showPending() {
  if (state.tab === 'pending') return;
  state.tab = 'pending';
  renderShipped();
}

el.tabs.addEventListener('click', (e) => {
  const tab = e.target.closest('[data-tab]');
  if (tab) setTab(tab.dataset.tab);
});

// 탭은 좌우 화살표와 Home, End로 옮긴다
el.tabs.addEventListener('keydown', (e) => {
  const tabs = [...el.tabs.querySelectorAll('[role="tab"]')];
  const at = tabs.indexOf(document.activeElement);
  const to = { ArrowRight: at + 1, ArrowLeft: at - 1, Home: 0, End: tabs.length - 1 }[e.key];
  if (at < 0 || to === undefined) return;
  e.preventDefault();
  const next = tabs[(to + tabs.length) % tabs.length];
  setTab(next.dataset.tab);
  next.focus();
});

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
  showPending();
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
  state.downloaded.clear(); // 내려받은 뒤에 사유를 고치면 다시 받아야 한다

  const rows = exportRows();
  if (change.type === 'gone') {
    // 주문 취소로 고르면 표 맨 아래에 줄이 생기고, 바꾸면 사라진다
    holder.classList.toggle('is-missing', !isFilled(sel));
    holder.querySelector('.gone-note').textContent = ui.goneNote(isFilled(sel) ? sel.code : '');
    const show = state.filter === 'all' && state.view !== 'missing';
    ui.renderCancelled(el.cancelled, show ? toItems(rows).filter((i) => i.row.status === 'cancelled') : []);
    renderShipped(); // 출하 완료로 고르면 Shipped Orders에 줄이 생긴다
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

// ---------- 이전·최신 나란히 비교 ----------

function openCompare(id) {
  const change = state.changes[id];
  if (!change) return;
  ui.renderCompare(el.compareBody, change, compareChange(change), {
    prev: state.files.prev?.name ?? '',
    curr: state.files.curr?.name ?? '',
  });
  if (!el.compare.open) el.compare.showModal();
  el.compare.scrollTop = 0;
}

for (const box of [el.rows, el.cancelled, el.goneList]) {
  box.addEventListener('click', (e) => {
    const btn = e.target.closest('[data-action="compare"]');
    const row = btn?.closest('[data-id]');
    if (row) openCompare(Number(row.dataset.id));
  });
}

// 창 바깥(어두운 배경)을 누르면 닫는다
el.compare.addEventListener('click', (e) => {
  if (e.target === el.compare) el.compare.close();
});

// ---------- 안내 메일 초안 ----------

const MAIL_PREFS_KEY = 'dh-delivery-mail-prefs';

/** 지연 기준일. 비었거나 1보다 작으면 기본값을 쓴다. */
function mailDelayDays() {
  const n = Math.floor(Number(el.mail.days.value));
  return Number.isFinite(n) && n >= 1 ? n : DEFAULT_DELAY_DAYS;
}

/** 메일 탭을 볼 때만 그린다. 사유, 고객사, 사라진 주문 선택, 파일명 이름이 바뀌면 renderCounts를 거쳐 여기로 온다. */
function renderMail() {
  if (state.tab !== 'mail' || !isReady()) return;
  const delayDays = mailDelayDays();
  const draft = buildMailDraft({
    changes: state.changes,
    reasons: state.reasons,
    shipTo: scope(),
    label: label(),
    delayDays,
    recipient: el.mail.recipient.value,
    sender: el.mail.sender.value,
  });
  ui.renderMail(el.mail, draft, {
    delayDays,
    noReason: draft.delayed.filter((d) => !d.reason).length,
    goneMissing: forShipTo().filter((c) => c.type === 'gone' && isMissing(c)).length,
  });
}

// 호칭, 이름, 기준일만 이 PC에 남긴다(주문 데이터는 저장하지 않는다). 저장이 막혀 있어도 그대로 동작한다
try {
  const saved = JSON.parse(localStorage.getItem(MAIL_PREFS_KEY) ?? '{}');
  if (typeof saved.recipient === 'string') el.mail.recipient.value = saved.recipient;
  if (typeof saved.sender === 'string') el.mail.sender.value = saved.sender;
  if (Number(saved.delayDays) >= 1) el.mail.days.value = String(Math.floor(Number(saved.delayDays)));
} catch {
  // 저장된 값을 못 읽으면 기본값으로 시작한다
}

for (const input of [el.mail.recipient, el.mail.sender, el.mail.days]) {
  input.addEventListener('input', () => {
    try {
      localStorage.setItem(
        MAIL_PREFS_KEY,
        JSON.stringify({ recipient: el.mail.recipient.value, sender: el.mail.sender.value, delayDays: mailDelayDays() }),
      );
    } catch {
      // 저장하지 못해도 초안은 만든다
    }
    renderMail();
  });
}

// 기준일 칸을 비우거나 0을 넣고 나가면 실제로 쓰인 값으로 되돌려 보여 준다
el.mail.days.addEventListener('change', () => {
  el.mail.days.value = String(mailDelayDays());
});

let mailStatusTimer;
function showMailStatus(text, failed = false) {
  clearTimeout(mailStatusTimer);
  el.mail.status.textContent = text;
  el.mail.status.classList.toggle('is-error', failed);
  mailStatusTimer = setTimeout(() => (el.mail.status.textContent = ''), failed ? 8000 : 2500);
}

/** 클립보드에 넣는다. 막혀 있으면 그 부분을 선택해 두고 Ctrl+C를 안내한다. */
async function copyFrom(node, what) {
  try {
    await navigator.clipboard.writeText(node.textContent);
    showMailStatus(`${what}을 복사했습니다`);
  } catch (err) {
    console.error(err);
    const range = document.createRange();
    range.selectNodeContents(node);
    const selection = window.getSelection();
    selection.removeAllRanges();
    selection.addRange(range);
    showMailStatus(`복사하지 못했습니다. ${what}을 선택해 두었으니 Ctrl+C로 복사하세요.`, true);
  }
}

el.mail.copySubject.addEventListener('click', () => copyFrom(el.mail.subject, '제목'));
el.mail.copyBody.addEventListener('click', () => copyFrom(el.mail.body, '본문'));

// ---------- 다운로드 ----------

el.optDelivery.addEventListener('change', () => {
  el.table.classList.toggle('no-delivery', !el.optDelivery.checked);
  el.shTable.classList.toggle('no-delivery', !el.optDelivery.checked);
  state.downloaded.delete(state.shipTo);
  if (isReady()) renderCounts();
});

/** 줄 배경색 범례의 문구는 용도에 따라 다르다(고객 발송용 엑셀에는 색이 없다). */
const syncStyle = () => {
  el.review.dataset.style = el.optStyle.value;
};
syncStyle();

// 발송용 설정을 바꾸면 이미 내려받은 파일과 달라지므로 다시 받아야 한다
for (const input of [el.optStyle, el.optShipped, el.optLabel]) {
  input.addEventListener('input', () => {
    syncStyle();
    state.downloaded.delete(state.shipTo);
    if (!isReady()) return;
    el.exportFile.textContent = `파일명: ${fileName()}`;
    renderCounts();
  });
}

el.download.addEventListener('click', async () => {
  el.download.disabled = true;
  try {
    const wb = buildWorkbook(exportRows(), {
      style: el.optStyle.value,
      includeDeliveryDate: el.optDelivery.checked,
      includeShipped: el.optShipped.checked,
      shippedRows: shippedRows(),
    });
    const blob = new Blob([await wb.xlsx.writeBuffer()], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = fileName();
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
