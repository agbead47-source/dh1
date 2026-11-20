// 검토 화면 렌더링. 상태는 main.js가 갖고, 여기서는 받은 값을 그리기만 한다.
import { reasonOptions, needsReason, isFilled } from '../reasons.js';

export const TYPE_LABELS = {
  date: '납기 변경',
  qty: '수량 증감',
  split: '분할',
  merge: '통합',
  new: '신규',
  gone: '사라짐',
};

export const DELAY_ALERT_DAYS = 10;
export const DELAY_TOP_N = 5;

const ESC = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };
const esc = (v) => String(v ?? '').replace(/[&<>"']/g, (ch) => ESC[ch]);
const num = (n) => Number(n).toLocaleString('ko-KR');
const signed = (n) => (n > 0 ? `+${num(n)}` : `−${num(-n)}`);
const shipName = (name) => name || '(납품처 없음)';
const pct = (value, max) => (max > 0 ? Math.round((value / max) * 1000) / 10 : 0);

/** 날짜는 월-일이 먼저 읽히게 연도를 흐리게 한다. 글자는 그대로라 복사하면 YYYY-MM-DD다. */
const dateHtml = (v) => {
  const m = /^(\d{4}-)(\d{2}-\d{2})$/.exec(String(v ?? ''));
  return m ? `<span class="yr">${m[1]}</span>${m[2]}` : esc(v);
};

/** 버튼 묶음을 다시 그려도 키보드 포커스가 같은 버튼에 남게 한다. */
function setHtmlKeepFocus(el, html) {
  const active = document.activeElement;
  const key = el.contains(active) ? active.dataset.key : undefined;
  el.innerHTML = html;
  if (key === undefined) return;
  [...el.querySelectorAll('[data-key]')].find((node) => node.dataset.key === key)?.focus();
}

// ---------- 사이드바 ----------

/** 업로드 칸. 읽은 파일의 정리 결과(원본 → 고객 기준)도 여기에 보여 준다. */
export function renderSlot(el, file, error) {
  const body = el.querySelector('[data-role="body"]');
  el.classList.toggle('is-loaded', Boolean(file) && !error);
  el.classList.toggle('is-error', Boolean(error));
  if (error) {
    body.innerHTML = `<strong>읽지 못했습니다</strong><span>${esc(error)}</span><span class="drop-again">눌러서 다른 파일을 선택하세요.</span>`;
  } else if (file) {
    const s = file.stats;
    const extra = [`잔여 재등록(점) ${num(s.dottedRows)}건`, `합산 ${num(s.mergedLines)}건`];
    if (s.skipped) extra.push(`빈 행 제외 ${num(s.skipped)}건`);
    if (s.noDate) extra.push(`출하일 없음 ${num(s.noDate)}건`);
    const kind = file.source === 'schedule' ? '보냈던 일정표로 읽음' : '전산자료로 읽음';
    body.innerHTML = `<strong>${esc(file.name)}</strong><span class="drop-kind">${kind}</span><span class="drop-stat">원본 ${num(s.rawCount)}행 → 고객 기준 ${num(s.lineCount)}행</span><span class="drop-extra">${extra.join(', ')}</span>`;
  } else {
    body.textContent = body.dataset.prompt;
  }
}

/** 작업 순서 3단계의 상태. states는 { upload, review, export } 각각 { state, status }. */
export function renderSteps(els, states) {
  for (const [name, el] of Object.entries(els)) {
    el.dataset.state = states[name].state;
    el.querySelector('[data-role="status"]').textContent = states[name].status;
  }
}

export function renderShipTo(el, options, selected) {
  el.disabled = !options.length;
  el.innerHTML = options.length
    ? options
        .map((o) => {
          const missing = o.missing ? `, 미입력 ${num(o.missing)}건` : '';
          return `<option value="${esc(o.name)}"${o.name === selected ? ' selected' : ''}>${esc(o.label ?? shipName(o.name))} (변경 ${num(o.changed)}건${missing})</option>`;
        })
        .join('')
    : '<option>파일 2개를 올리면 고를 수 있습니다</option>';
}

/** 사유 입력 진행 표시. need는 사유가 필요한 변경 건 수, missing은 그중 미입력. */
function progressHtml(need, missing) {
  if (!need) return '<span class="progress-text">사유가 필요한 변경 건이 없습니다</span>';
  const filled = need - missing;
  return `
    <span class="progress-text">사유 <b>${num(filled)}</b> / ${num(need)}건 입력${missing ? '' : ' 완료'}</span>
    <span class="bar-track" aria-hidden="true"><span class="bar-fill" style="width:${pct(filled, need)}%"></span></span>`;
}

/** 사이드바 2단계의 진행 표시. data가 null이면 비운다. */
export function renderSideProgress(el, data) {
  el.innerHTML = data ? progressHtml(data.need, data.missing) : '';
  el.classList.toggle('is-done', Boolean(data && data.need && !data.missing));
}

export function renderExportWarn(el, missing) {
  el.textContent = missing ? `사유 미입력 ${num(missing)}건. 그대로 받을 수도 있습니다.` : '';
}

// ---------- 본문 머리, 빈 상태 ----------

/** selected가 null이면 아직 비교 전이다. */
export function renderMainHead(titleEl, subEl, scopeEls, selected, files) {
  const ready = selected !== null;
  titleEl.textContent = ready ? `${shipName(selected)} 변경 현황` : '변경 현황';
  subEl.textContent = ready ? `${files.prev.name} → ${files.curr.name}` : '';
  for (const el of scopeEls) el.textContent = ready ? `${shipName(selected)} 기준` : '선택한 고객사 기준';
}

export function renderEmpty(el, files, errors, ready) {
  el.hidden = ready;
  if (ready) return;
  const failed = Boolean(errors.prev || errors.curr);
  let lead;
  if (failed) lead = '파일을 읽지 못했습니다. 빨갛게 표시된 칸의 안내를 보고 전산 출하계획 엑셀(.xlsx)을 다시 올려 주세요.';
  else if (!files.prev && !files.curr) lead = '이전 파일과 최신 파일을 올리면 바뀐 주문이 여기에 정리됩니다.';
  else if (files.prev) lead = '이전 파일을 읽었습니다. 최신 파일을 올려 주세요.';
  else lead = '최신 파일을 읽었습니다. 이전 파일을 올려 주세요.';
  el.classList.toggle('is-error', failed);
  el.innerHTML = `
    <p class="empty-lead">${lead}</p>
    <p class="empty-note">이전 파일은 전산자료나 지난번에 보냈던 Delivery Schedule, 최신 파일은 전산자료를 올립니다. 실제 파일 없이 화면을 먼저 보려면
      <button type="button" class="link-btn" data-action="sample">가짜 샘플 데이터로 보기</button></p>`;
}

// ---------- 요약 카드 ----------

/**
 * data가 null이면 값 대신 줄표를 보여 준다(빈 상태).
 * "사유 미입력" 카드는 버튼이다. 누르면 목록을 미입력 건만 모아 본다.
 */
export function renderCards(el, data, view) {
  const onlyMissing = view === 'missing';
  const d = data ?? {};
  const card = (label, key, sub, alert) => {
    const value = data ? data[key] : null;
    const shown = value === null ? '–' : num(value);
    const cls = `card${alert && value ? ' is-alert' : ''}${value === null ? ' is-blank' : ''}`;
    const inner = `<span class="card-label">${label}</span><span class="card-value">${shown}</span><span class="card-sub">${data ? sub : ''}</span>`;
    if (key !== 'missing' || !data) return `<div class="${cls}">${inner}</div>`;
    // 할 일이 남았는지(빨강), 끝났는지(청록)를 카드 전체로 알린다. 누르면 무엇이 되는지는 밑줄 친 한 줄로 적는다
    const todo = d.missing ? ' is-todo' : d.need ? ' is-done' : '';
    const act = onlyMissing ? '변경만 보기로 돌아가기' : d.missing ? '미입력만 모아 보기' : '';
    return `<button type="button" class="${cls}${todo} card-btn" data-action="missing" data-key="missing" aria-pressed="${onlyMissing}">${inner}${act ? `<span class="card-act">${act}</span>` : ''}</button>`;
  };
  // 사유 미입력에는 표에 줄이 없는 "최신 파일에 없는 주문"도 들어 있다. 어디서 고르는지 적어 둔다
  const inGone = d.goneMissing ? `. ${num(d.goneMissing)}건은 '최신 파일에 없는 주문' 탭` : '';
  const missingSub = !d.need
    ? '사유가 필요한 건 없음'
    : d.missing
      ? `${num(d.need)}건 중 ${num(d.need - d.missing)}건 입력${inGone}`
      : `${num(d.need)}건 모두 입력. 내려받을 수 있습니다`;
  setHtmlKeepFocus(
    el,
    card('엑셀에 나가는 줄', 'lines', d.cancelledLines ? `변경 없는 줄 포함, 주문 취소 ${num(d.cancelledLines)}줄 포함` : '변경 없는 줄 포함') +
      card('변경된 주문', 'changed', `주문 ${num(d.orders ?? 0)}건 중`) +
      card(`${DELAY_ALERT_DAYS}일 이상 지연`, 'delayed', d.maxDelay > 0 ? `가장 큰 지연 ${num(d.maxDelay)}일` : '늦어진 주문 없음', true) +
      card('사유 미입력', 'missing', missingSub, true),
  );
}

// ---------- 차트 ----------

const chartEmpty = (text) => `<p class="chart-empty">${text}</p>`;
const NO_DATA = '파일 2개를 올리면 표시됩니다.';

/** 변경 유형별 건수(선택한 고객사). 막대를 누르면 그 유형으로 거른다. counts가 null이면 빈 상태. */
export function renderTypeChart(el, counts, active) {
  if (!counts) {
    el.innerHTML = chartEmpty(NO_DATA);
    return;
  }
  const max = Math.max(0, ...Object.values(counts));
  setHtmlKeepFocus(
    el,
    Object.entries(TYPE_LABELS)
      .map(([code, label]) => {
        const n = counts[code] ?? 0;
        return `
        <button type="button" class="bar-row${code === active ? ' is-active' : ''}${n ? '' : ' is-zero'}" data-filter="${code}" data-key="${code}" aria-pressed="${code === active}">
          <span class="bar-label">${label}</span>
          <span class="bar-track"><span class="bar-fill" style="width:${pct(n, max)}%"></span></span>
          <span class="bar-value">${num(n)}<span class="visually-hidden">건</span></span>
        </button>`;
      })
      .join(''),
  );
}

/** 지연이 큰 주문 상위(선택한 고객사). items는 delayDays 내림차순, total은 지연된 주문 전체 수. */
export function renderDelayChart(el, items, total, showShip = false) {
  if (!items) {
    el.innerHTML = chartEmpty(NO_DATA);
    return;
  }
  if (!items.length) {
    el.innerHTML = chartEmpty('출하일이 늦어진 주문이 없습니다.');
    return;
  }
  const max = items[0].delayDays;
  const rows = items
    .map((c) => {
      const alert = c.delayDays >= DELAY_ALERT_DAYS;
      return `
      <li class="delay-row${alert ? ' is-alert' : ''}">
        <span class="delay-name">${showShip ? `<span class="ship-of">${esc(shipName(c.shipTo))}</span> ` : ''}<b>${esc(c.po)}</b> ${esc(c.partNo)}</span>
        <span class="bar-value">${num(c.delayDays)}일${alert ? '<span class="delay-flag">기준 초과</span>' : ''}</span>
        <span class="bar-track"><span class="bar-fill" style="width:${pct(c.delayDays, max)}%"></span></span>
      </li>`;
    })
    .join('');
  const more = total > items.length ? `지연 ${num(total)}건 중 상위 ${num(items.length)}건. ` : `지연 ${num(total)}건. `;
  el.innerHTML = `<ol class="delay-list">${rows}</ol><p class="chart-note">${more}빨간 막대는 ${DELAY_ALERT_DAYS}일 이상 지연입니다.</p>`;
}

/** 고객사별 변경 건수(전체). 막대를 누르면 그 고객사로 바꾼다. data-ship은 options의 순번이다. */
export function renderShipChart(el, options, selected) {
  if (!options) {
    el.innerHTML = chartEmpty(NO_DATA);
    return;
  }
  const max = Math.max(0, ...options.map((o) => o.changed));
  const total = options.reduce((s, o) => s + o.changed, 0);
  setHtmlKeepFocus(
    el,
    options
      .map((o, i) => {
        const on = o.name === selected;
        return `
        <button type="button" class="bar-row${on ? ' is-active' : ''}${o.changed ? '' : ' is-zero'}" data-ship="${i}" data-key="${i}" aria-pressed="${on}" title="${esc(shipName(o.name))}">
          <span class="bar-label">${esc(shipName(o.name))}</span>
          <span class="bar-track"><span class="bar-fill" style="width:${pct(o.changed, max)}%"></span></span>
          <span class="bar-value">${num(o.changed)}<span class="visually-hidden">건</span></span>
        </button>`;
      })
      .join('') + `<p class="chart-note">고객사 ${num(options.length)}곳, 변경 합계 ${num(total)}건. 굵게 표시한 줄이 지금 보는 고객사입니다.</p>`,
  );
}

// ---------- 보기 전환, 유형 칩, 사유 진행 ----------

const chipHtml = (attr, code, label, count, active, extra = '') =>
  `<button type="button" class="chip${extra}${code === active ? ' is-active' : ''}" ${attr}="${code}" data-key="${code}" aria-pressed="${code === active}">${label}${count ? ` <b>${count}</b>` : ''}</button>`;

/** 보기 전환: 전체 줄 / 변경된 주문만 / 사유 미입력만. */
export function renderViews(el, { lines, changed, missing }, active) {
  setHtmlKeepFocus(
    el,
    chipHtml('data-view', 'all', '전체', `${num(lines)}줄`, active) +
      chipHtml('data-view', 'changed', '변경만', `${num(changed)}건`, active) +
      chipHtml('data-view', 'missing', '사유 미입력', `${num(missing)}건`, active, ` chip-missing${missing ? ' has-missing' : ''}`),
  );
}

/** 표에 줄이 있는 유형만 칩으로 둔다. 사라진 주문은 표 아래에 따로 있다. */
export const TABLE_TYPES = ['date', 'qty', 'split', 'merge', 'new'];

export function renderChips(el, counts, active) {
  setHtmlKeepFocus(
    el,
    '<span class="chips-label">유형</span>' +
      chipHtml('data-filter', 'all', '모두', '', active) +
      TABLE_TYPES.map((code) => chipHtml('data-filter', code, TYPE_LABELS[code], num(counts[code] ?? 0), active)).join(''),
  );
}

export function renderReasonTools(el, { need, missing }) {
  el.innerHTML = `<span class="progress${need && !missing ? ' is-done' : ''}">${progressHtml(need, missing)}</span>`;
}

// ---------- Pending Order 표 ----------

/**
 * PO 끝의 점으로 구별한 잔여 재등록 표시. 전산 등록 PO(점 포함)와 뜻은 눌러서 펼친다
 * (마우스 없이 터치와 키보드로도 읽을 수 있게).
 */
function reissueHtml(c) {
  const r = c.reissue;
  if (!r?.dotted) return '';
  const erp = (c.curr ?? c.prev).erpPos.map(esc).join(', ');
  const label = r.newlyDotted ? '이번에 재등록' : '잔여 재등록';
  const help = r.newlyDotted
    ? '이전 파일에는 점이 없었고, 최신 파일에서 PO 끝에 점을 찍어 다시 등록된 주문입니다.'
    : '일부 출하 뒤 PO 끝에 점을 찍어 다시 등록한 잔여 오더입니다.';
  const mixed = r.mixed
    ? '<p class="is-alert">점 있는 행과 없는 행의 수량이 한 줄로 더해졌습니다. 원래 PO의 잔여 행이 남아 두 번 더해진 것은 아닌지 전산에서 확인하세요.</p>'
    : '';
  return `
    <details class="reissue">
      <summary>
        <span class="tag is-reissue${r.newlyDotted ? ' is-new' : ''}">${label}</span>${r.mixed ? '<span class="tag is-alert">중복 확인</span>' : ''}
        <span class="visually-hidden">설명 펼치기</span>
      </summary>
      <div class="reissue-help"><p>전산 등록 PO <b>${erp}</b></p><p>${help} 고객용 엑셀에는 점 없는 PO로 나갑니다.</p>${mixed}</div>
    </details>`;
}

/** 변경 없는 주문은 재등록이 새로 생겼거나 중복이 의심될 때만 표시한다. */
const showReissue = (c) => c.type !== 'same' || c.reissue?.mixed || c.reissue?.newlyDotted;

function deltaHtml(c) {
  // "얼마나 바뀌었나"는 배지 바로 아래에 굵은 글자로. 기준일 이상 지연만 빨간 바탕이다
  const tags = [];
  if (c.delayDays > 0) {
    tags.push(`<span class="delta${c.delayDays >= DELAY_ALERT_DAYS ? ' is-alert' : ''}">${num(c.delayDays)}일 지연</span>`);
  } else if (c.delayDays < 0) {
    tags.push(`<span class="delta is-early">${num(-c.delayDays)}일 앞당김</span>`);
  }
  if (c.qtyDiff && c.type !== 'new') tags.push(`<span class="delta">수량 ${signed(c.qtyDiff)}</span>`);
  return tags.join('');
}

function reasonHtml(c, sel) {
  const code = sel?.code ?? '';
  const options = reasonOptions(c.type)
    .map((r) => `<option value="${r.code}"${r.code === code ? ' selected' : ''}>${esc(r.ko)}</option>`)
    .join('');
  const placeholder = c.type === 'gone' ? '출하 완료 / 주문 취소 선택' : '사유 선택';
  return `
    <select class="reason-select${isFilled(sel) ? '' : ' is-missing'}" aria-label="${esc(c.po)} ${esc(c.partNo)} 변경 사유">
      <option value=""${code ? '' : ' selected'}>${placeholder}</option>${options}
    </select>
    <input class="reason-text" type="text" aria-label="직접 입력한 사유" placeholder="REMARKS에 그대로 들어갈 문장" value="${esc(sel?.text)}"${code === 'custom' ? '' : ' hidden'} />`;
}

/**
 * 표의 한 줄. 열 순서는 엑셀과 같고, 맨 앞의 "변경" 칸만 화면용이다.
 * item: { row(buildExportRows의 행), id(changes 순번), change, lead(그 주문의 첫 줄), editor(사유 입력 줄), sel }
 */
const COMPARE_BUTTON =
  '<button type="button" class="link-btn compare-btn" data-action="compare" aria-haspopup="dialog">이전·최신 비교</button>';

function rowHtml({ row: r, id, change: c, lead, editor, sel }) {
  const cancelled = r.status === 'cancelled';
  const badge = cancelled
    ? '<span class="badge badge-gone">주문 취소</span>'
    : r.type === 'same'
      ? ''
      : `<span class="badge badge-${r.type}">${TYPE_LABELS[r.type]}</span>`;
  // 변경 칸에는 "무엇이, 얼마나"만 둔다. 재등록 표시와 비교 버튼은 PO 아래로 내려 한 단계 가볍게
  const marks = badge + (lead && !cancelled ? deltaHtml(c) : '');
  const sub = (lead && showReissue(c) ? reissueHtml(c) : '') + (lead && c ? COMPARE_BUTTON : '');
  const missing = !cancelled && needsReason(r.type) && !isFilled(sel);
  return `
    <tr class="st-${r.status}${missing ? ' is-missing' : ''}" data-id="${id}">
      <td class="c-mark">${marks ? `<div class="marks">${marks}</div>` : ''}</td>
      <td class="c-order">${dateHtml(r.orderDate)}</td>
      <td class="c-delivery">${dateHtml(r.deliveryDate)}</td>
      <td class="c-prev${r.exwChanged ? ' is-changed' : ''}">${dateHtml(r.prevExw)}</td>
      <td class="c-exw${r.exwChanged ? ' is-changed' : ''}">${dateHtml(r.exw)}</td>
      <td class="c-ship">${esc(r.shipTo)}</td>
      <td class="c-po"><span class="po-no">${esc(r.po)}</span>${sub ? `<div class="po-sub">${sub}</div>` : ''}</td>
      <td class="c-partno">${esc(r.partNo)}</td>
      <td class="c-name">${esc(r.partName)}</td>
      <td class="c-qty">${num(r.qty)}</td>
      <td class="c-unit">${esc(r.unit)}</td>
      <td class="c-remarks">${editor ? `<div class="reason">${reasonHtml(c, sel)}</div>` : ''}<div class="remark-out">${esc(r.remarks)}</div></td>
    </tr>`;
}

export function renderTable(el, items, emptyText) {
  el.innerHTML = items.length
    ? items.map(rowHtml).join('')
    : `<tr class="empty-row"><td colspan="12" class="empty">${emptyText}</td></tr>`;
}

/** 주문 취소로 고른 건. 엑셀에서도 맨 아래에 들어간다. */
export function renderCancelled(el, items) {
  el.innerHTML = items.map(rowHtml).join('');
}

/** 사유를 고친 주문의 줄만 고친다. trs와 rows는 같은 주문의 줄이 같은 순서로 온다. */
export function updateOrderRows(trs, rows, missing) {
  trs.forEach((tr, i) => {
    tr.classList.toggle('is-missing', missing);
    const out = tr.querySelector('.remark-out');
    if (out && rows[i]) out.textContent = rows[i].remarks;
  });
}

export function renderShown(el, shown, total) {
  el.textContent =
    shown === total
      ? `엑셀에 나가는 ${num(total)}줄을 모두 보고 있습니다.`
      : `엑셀에 나가는 ${num(total)}줄 중 ${num(shown)}줄을 보고 있습니다.`;
}

// ---------- 시트 탭, Shipped Orders 표 ----------

/** 탭 이름 옆의 줄 수와 선택 상태. counts는 { pending, shipped }. */
export function renderTabs(el, counts, active) {
  for (const tab of el.querySelectorAll('[role="tab"]')) {
    const on = tab.dataset.tab === active;
    tab.setAttribute('aria-selected', String(on));
    tab.tabIndex = on ? 0 : -1;
    const name = tab.dataset.tab;
    // 확인용 탭은 줄이 아니라 주문 건수를 세고, 아직 고르지 않은 건이 있으면 빨갛게 알린다
    const count = tab.querySelector('b'); // 안내 메일 탭에는 건수가 없다
    if (count) count.textContent = name === 'gone' ? `${num(counts.gone)}건` : `${num(counts[name])}줄`;
    if (name === 'gone') tab.classList.toggle('is-alert', counts.goneMissing > 0);
  }
}

/** 이번에 옮겨 적은 줄은 선적 정보가 아직 없다. 빈칸인 이유를 칸에 적어 둔다. */
const shipCell = (value, added) => (value ? esc(value) : added ? '<span class="cell-todo">엑셀에서 입력</span>' : '');

/** 읽기 전용. rows는 buildShippedRows 결과(넘겨받은 줄 carried, 이번에 추가한 줄 added). */
export function renderShipped(el, rows) {
  el.closest('table').classList.toggle('is-empty', !rows.length);
  if (!rows.length) {
    el.innerHTML =
      '<tr class="empty-row"><td colspan="15" class="empty">출하 기록이 없습니다. 이전 파일로 지난번에 보냈던 일정표를 올리면 그 Shipped Orders를 넘겨받고, "최신 파일에 없는 주문" 탭에서 출하 완료를 고르면 여기에 줄이 생깁니다.</td></tr>';
    return;
  }
  el.innerHTML = rows
    .map((r) => {
      const added = r.status === 'added';
      return `
      <tr class="${added ? 'st-new' : 'st-same'}">
        <td class="c-mark">${added ? '<span class="badge badge-new">이번에 추가</span>' : ''}</td>
        <td class="c-date">${dateHtml(r.orderDate)}</td>
        <td class="c-date c-delivery">${dateHtml(r.deliveryDate)}</td>
        <td class="c-date">${dateHtml(r.exw)}</td>
        <td class="c-date">${shipCell(r.etd, added)}</td>
        <td class="c-date">${shipCell(r.eta, added)}</td>
        <td>${esc(r.shipTo)}</td>
        <td>${shipCell(r.invoiceNo, added)}</td>
        <td>${shipCell(r.shipmentMode, added)}</td>
        <td class="c-po">${esc(r.po)}</td>
        <td class="c-partno">${esc(r.partNo)}</td>
        <td class="c-name">${esc(r.partName)}</td>
        <td class="c-qty">${num(r.qty)}</td>
        <td class="c-unit">${esc(r.unit)}</td>
        <td class="c-note">${esc(r.remarks)}</td>
      </tr>`;
    })
    .join('');
}

export function renderShippedShown(el, rows) {
  const added = rows.filter((r) => r.status === 'added').length;
  el.textContent = rows.length
    ? `넘겨받은 ${num(rows.length - added)}줄, 이번에 추가한 ${num(added)}줄. 내려받는 엑셀의 Shipped Orders 시트에 들어갑니다.`
    : '줄이 없으면 엑셀에 Shipped Orders 시트를 만들지 않습니다.';
}

// ---------- 이전 파일에만 있는 주문 ----------

const GONE_NOTES = {
  '': '골라야 합니다',
  shipped: 'Shipped Orders로 옮겨집니다',
  cancelled: 'Pending Order 맨 아래에 Order cancelled 줄로 들어갑니다',
};
export const goneNote = (code) => GONE_NOTES[code] ?? '';

export function renderGoneTitle(el, total, missing) {
  el.innerHTML = `확인할 주문 ${num(total)}건${missing ? ` <span class="tag is-alert">미선택 ${num(missing)}건</span>` : ''}`;
}

/** items: { id, change, sel }. showShip이면(전체 고객사 보기) 주문마다 납품처를 적는다. */
export function renderGone(el, items, showShip = false) {
  el.innerHTML = items
    .map(({ id, change: c, sel }) => {
      const lines = c.prev.lines
        .map(
          (l) =>
            `<div><span class="d">${l.date ? dateHtml(l.date) : '출하일 미정'}</span><span class="q">${num(l.qty)}<span class="u">${esc(c.prev.unit)}</span></span></div>`,
        )
        .join('');
      return `
      <li class="row gone-row${isFilled(sel) ? '' : ' is-missing'}" data-id="${id}">
        <span class="badge badge-gone">${TYPE_LABELS.gone}</span>
        <div class="ident">
          <div class="po">${showShip ? `<span class="ship-of">${esc(shipName(c.shipTo))}</span> ` : ''}${esc(c.po)}</div>
          <div class="part"><span>${esc(c.partNo)}</span>${esc(c.partName)}</div>
          <div class="po-sub">${reissueHtml(c)}${COMPARE_BUTTON}</div>
        </div>
        <div class="lines"><span class="visually-hidden">이전 출하 계획</span>${lines}</div>
        <div class="reason">${reasonHtml(c, sel)}<span class="gone-note">${goneNote(sel?.code ?? '')}</span></div>
      </li>`;
    })
    .join('');
}

// ---------- 안내 메일 초안 ----------

/**
 * els: { subject, body, summary, warn }. draft는 buildMailDraft 결과.
 * 제목과 본문은 사용자 입력(호칭, 이름)이 섞인 일반 텍스트라 textContent로만 넣는다.
 * info: { detail(상세 내용 넣기), delayDays, noReason(본문에 적힌 지연 주문 중 사유 미입력), goneMissing(출하/취소를 고르지 않은 주문) }
 */
export function renderMail(els, draft, { detail, delayDays, noReason, goneMissing }) {
  els.subject.textContent = draft.subject;
  els.body.textContent = draft.body;
  if (!detail) {
    // 기본은 짧은 안내다. 주문 번호나 건수가 본문에 없으니 요약과 경고도 필요 없다
    els.summary.textContent = '짧은 안내만 적습니다. 변경 내용은 첨부한 일정표에 있습니다.';
    els.warn.textContent = '';
    return;
  }
  els.summary.textContent = draft.delayed.length
    ? `지연 기준 ${num(delayDays)}일 이상 ${num(draft.delayed.length)}건을 본문에 한 건씩 적었습니다.`
    : `${num(delayDays)}일 이상 늦어진 주문이 없어 지연 안내 문단은 넣지 않았습니다.`;
  const warns = [];
  if (noReason) warns.push(`사유 미입력 ${num(noReason)}건은 사유 없이 적힙니다.`);
  if (goneMissing) warns.push(`출하 완료/주문 취소를 고르지 않은 ${num(goneMissing)}건은 요약에 들어가지 않습니다.`);
  els.warn.textContent = warns.join(' ');
}

// ---------- 이전·최신 나란히 비교 창 ----------

const DATE_STATE = { same: '그대로', changed: '수량 변경', added: '새로 생김', removed: '없어짐' };
const qtyCell = (v) => (v == null ? '<span class="none">없음</span>' : num(v));
const diffCell = (v) => (v ? signed(v) : '');

function sourcesHtml(title, fileName, order, sources) {
  if (!order) return `<section class="cmp-side"><h4>${title}</h4><p class="cmp-file">${esc(fileName)}</p><p class="cmp-none">이 파일에는 이 주문이 없습니다.</p></section>`;
  const body = order.fromSchedule
    ? '<p class="cmp-none">보냈던 일정표에서 읽은 주문이라 전산 등록 행(품목코드, 제조사, 수주번호)이 없습니다.</p>'
    : `<div class="cmp-scroll"><table class="cmp-table">
        <thead><tr><th>출하일</th><th>전산 PO</th><th>품목코드</th><th>제조사</th><th>수주번호</th><th class="n">수량</th></tr></thead>
        <tbody>${sources
          .map(
            (s) =>
              `<tr><td>${s.date ? dateHtml(s.date) : '미정'}</td><td>${esc(s.erpPo)}</td><td>${esc(s.itemCode)}</td><td>${esc(s.maker)}</td><td>${esc(s.orderNo)}</td><td class="n">${num(s.qty)}</td></tr>`,
          )
          .join('')}</tbody>
      </table></div>`;
  return `<section class="cmp-side"><h4>${title} <span class="cmp-count">${num(sources.length)}행</span></h4><p class="cmp-file">${esc(fileName)}</p>${body}</section>`;
}

/**
 * 한 주문의 이전·최신을 나란히 보여 준다.
 * @param {object} c 변경 건
 * @param {object} d compareChange(c) 결과
 * @param {{prev: string, curr: string}} names 파일 이름
 */
export function renderCompare(el, c, d, names) {
  const unit = (c.curr ?? c.prev).unit;
  const notes = [];
  if (c.type === 'same' && d.sourcesChanged) notes.push('고객 기준으로는 변경이 없지만 전산 등록 행(공급처 배분 등)은 달라졌습니다. 고객 일정표에는 변화가 없습니다.');
  if (c.type === 'new') notes.push('이전 파일에 없던 주문입니다.');
  if (c.type === 'gone') notes.push('최신 파일에 없는 주문입니다. 출하 완료인지 주문 취소인지 확인이 필요합니다.');
  if (c.reissue?.newlyDotted) notes.push('이전 파일에는 점이 없었고 최신 파일에서 PO 끝에 점을 찍어 다시 등록되었습니다(일부 출하 뒤의 잔여 재등록).');
  if (c.reissue?.mixed) notes.push('점 있는 행과 없는 행이 한 주문으로 합산되었습니다. 두 번 더해진 것은 아닌지 아래 전산 등록 행에서 확인하세요.');

  const dateRows = d.dates
    .map(
      (r) => `<tr class="ds-${r.state}">
        <td>${r.date ? dateHtml(r.date) : '출하일 미정'}</td>
        <td class="n">${qtyCell(r.prevQty)}</td>
        <td class="n">${qtyCell(r.currQty)}</td>
        <td class="n diff">${diffCell(r.diff)}</td>
        <td><span class="cmp-state">${DATE_STATE[r.state]}</span></td>
      </tr>`,
    )
    .join('');

  el.innerHTML = `
    <header class="cmp-head">
      <div class="marks"><span class="badge badge-${c.type === 'same' ? 'same' : c.type}">${c.type === 'same' ? '변경 없음' : TYPE_LABELS[c.type]}</span>${deltaHtml(c)}</div>
      <h3 id="compare-title">${esc(c.po)} <span class="cmp-part">${esc(c.partNo)}</span></h3>
      <p class="cmp-sub"><span class="ship-of">${esc(shipName(c.shipTo))}</span> ${esc(c.partName)}</p>
    </header>
    ${notes.map((n) => `<p class="cmp-note">${n}</p>`).join('')}

    <section>
      <h4>고객에게 보이는 출하 계획 <span class="cmp-count">단위 ${esc(unit)}</span></h4>
      <div class="cmp-scroll"><table class="cmp-table cmp-dates">
        <thead><tr><th>출하일</th><th class="n">이전</th><th class="n">최신</th><th class="n">차이</th><th>상태</th></tr></thead>
        <tbody>${dateRows}</tbody>
        <tfoot><tr><th>합계</th><td class="n">${qtyCell(d.prevTotal)}</td><td class="n">${qtyCell(d.currTotal)}</td><td class="n diff">${diffCell(d.qtyDiff)}</td><td></td></tr></tfoot>
      </table></div>
    </section>

    <div class="cmp-sides">
      ${sourcesHtml('이전 파일의 전산 등록 행', names.prev, c.prev, d.prevSources)}
      ${sourcesHtml('최신 파일의 전산 등록 행', names.curr, c.curr, d.currSources)}
    </div>`;
}

// ---------- 열 정렬 ----------

/** 머리 칸의 글자를 버튼으로 감싼다(키보드로도 누를 수 있게). 처음 한 번만 부른다. */
export function initSortHeaders(table) {
  for (const th of table.querySelectorAll('th[data-sort]')) {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'th-sort';
    btn.append(...th.childNodes);
    const arrow = document.createElement('span');
    arrow.className = 'sort-arrow';
    arrow.setAttribute('aria-hidden', 'true');
    btn.append(arrow);
    th.append(btn);
  }
}

/** 정렬 중인 열에 방향을 표시하고, 표 위에 안내와 해제 버튼을 보여 준다. sort는 { key, dir } 또는 null. */
export function renderSort(table, noteEl, sort, label) {
  for (const th of table.querySelectorAll('th[data-sort]')) {
    const on = sort && th.dataset.sort === sort.key;
    th.setAttribute('aria-sort', on ? (sort.dir === 'asc' ? 'ascending' : 'descending') : 'none');
    th.querySelector('.sort-arrow').textContent = on ? (sort.dir === 'asc' ? '▲' : '▼') : '';
  }
  noteEl.hidden = !sort;
  if (!sort) return;
  noteEl.innerHTML = `<span><b>${esc(label)}</b> ${sort.dir === 'asc' ? '오름차순' : '내림차순'}으로 보는 중입니다. 화면에서 보는 순서만 바뀌고, 엑셀은 출하일 순서로 내려받습니다.</span> <button type="button" class="link-btn" data-action="sort-reset">엑셀 순서로 되돌리기</button>`;
}
