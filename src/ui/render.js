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

/** 버튼 묶음을 다시 그려도 키보드 포커스가 같은 버튼에 남게 한다. */
function setHtmlKeepFocus(el, html) {
  const active = document.activeElement;
  const key = el.contains(active) ? active.dataset.key : undefined;
  el.innerHTML = html;
  if (key === undefined) return;
  [...el.querySelectorAll('[data-key]')].find((node) => node.dataset.key === key)?.focus();
}

// ---------- 사이드바 ----------

export function renderSlot(el, file, error) {
  const body = el.querySelector('[data-role="body"]');
  el.classList.toggle('is-loaded', Boolean(file) && !error);
  el.classList.toggle('is-error', Boolean(error));
  if (error) {
    body.innerHTML = `<strong>읽지 못했습니다</strong><span>${esc(error)}</span>`;
  } else if (file) {
    body.innerHTML = `<strong>${esc(file.name)}</strong><span>원본 ${num(file.stats.rawCount)}행</span>`;
  }
}

function statLine(label, s) {
  const extra = [`잔여 재등록(점) ${num(s.dottedRows)}건`, `합산 ${num(s.mergedLines)}건`];
  if (s.skipped) extra.push(`빈 행 제외 ${num(s.skipped)}건`);
  if (s.noDate) extra.push(`출하일 없음 ${num(s.noDate)}건`);
  return `<p><b>${label}</b> 원본 ${num(s.rawCount)}행 → 고객 기준 ${num(s.lineCount)}행 (${extra.join(', ')})</p>`;
}

export function renderSummaryLine(el, files) {
  el.innerHTML = [files.prev && statLine('이전', files.prev.stats), files.curr && statLine('최신', files.curr.stats)]
    .filter(Boolean)
    .join('');
}

export function renderShipTo(el, options, selected) {
  el.disabled = !options.length;
  el.innerHTML = options.length
    ? options
        .map(
          (o) =>
            `<option value="${esc(o.name)}"${o.name === selected ? ' selected' : ''}>${esc(shipName(o.name))} · 변경 ${num(o.changed)}건</option>`,
        )
        .join('')
    : '<option>파일 2개를 올리면 고를 수 있습니다</option>';
}

export function renderExportWarn(el, missing) {
  el.textContent = missing ? `사유 미입력 ${num(missing)}건이 남아 있습니다` : '';
}

// ---------- 본문 머리, 빈 상태 ----------

/** selected가 null이면 아직 비교 전이다. */
export function renderMainHead(titleEl, subEl, scopeEls, selected, files) {
  const ready = selected !== null;
  titleEl.textContent = ready ? `${shipName(selected)} 변경 현황` : '변경 현황';
  subEl.textContent = ready ? `${files.prev.name} → ${files.curr.name}` : '';
  for (const el of scopeEls) el.textContent = ready ? `${shipName(selected)} 기준` : '선택한 고객사 기준';
}

export function renderEmpty(el, files, ready) {
  el.hidden = ready;
  if (ready) return;
  const step = (done, text) =>
    `<li class="${done ? 'is-done' : ''}">${text}<span class="step-state">${done ? '완료' : '할 일'}</span></li>`;
  const next = !files.prev && !files.curr
    ? '왼쪽에 이전 파일과 최신 파일을 올리면 변경 건이 여기에 정리됩니다.'
    : files.prev
      ? '이전 파일을 읽었습니다. 최신 파일을 올려 주세요.'
      : '최신 파일을 읽었습니다. 이전 파일을 올려 주세요.';
  el.innerHTML = `
    <p class="empty-lead">${next}</p>
    <ol class="steps">
      ${step(Boolean(files.prev), '이전 전산자료 올리기')}
      ${step(Boolean(files.curr), '최신 전산자료 올리기')}
      ${step(false, '변경 건 검토하고 사유 입력')}
      ${step(false, '고객 양식 엑셀 내려받기')}
    </ol>
    <p class="empty-note">먼저 둘러보려면 왼쪽의 "가짜 샘플 데이터로 보기"를 누르세요.</p>`;
}

// ---------- 요약 카드 ----------

/** data가 null이면 값 대신 줄표를 보여 준다(빈 상태). */
export function renderCards(el, data) {
  const card = (label, key, alert) => {
    const value = data ? data[key] : null;
    const shown = value === null ? '–' : num(value);
    return `<div class="card${alert && value ? ' is-alert' : ''}${value === null ? ' is-blank' : ''}"><span class="card-label">${label}</span><span class="card-value">${shown}</span></div>`;
  };
  el.innerHTML =
    card('전체 주문 행', 'lines') +
    card('변경 건', 'changed') +
    card(`${DELAY_ALERT_DAYS}일 이상 지연`, 'delayed', true) +
    card('사유 미입력', 'missing', true);
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
export function renderDelayChart(el, items, total) {
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
        <span class="delay-name"><b>${esc(c.po)}</b> ${esc(c.partNo)}</span>
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

// ---------- 필터 칩 ----------

export function renderChips(el, counts, active) {
  const total = Object.values(counts).reduce((s, n) => s + n, 0);
  const chip = (code, label, n) =>
    `<button type="button" class="chip${code === active ? ' is-active' : ''}" data-filter="${code}" data-key="${code}" aria-pressed="${code === active}">${label} <b>${num(n)}</b></button>`;
  setHtmlKeepFocus(
    el,
    chip('all', '전체', total) +
      Object.entries(TYPE_LABELS)
        .map(([code, label]) => chip(code, label, counts[code] ?? 0))
        .join(''),
  );
}

// ---------- 변경 목록 ----------

function linesHtml(order, unit) {
  if (!order) return '<div class="lines is-none">없음</div>';
  return `<div class="lines">${order.lines
    .map((l) => `<div><span class="d">${esc(l.date || '출하일 미정')}</span><span class="q">${num(l.qty)} ${esc(unit)}</span></div>`)
    .join('')}</div>`;
}

/** PO 끝의 점으로 구별한 잔여 재등록 표시. 전산 등록 PO(점 포함)를 함께 보여 준다. */
function reissueHtml(c) {
  const r = c.reissue;
  if (!r?.dotted) return '';
  const erp = (c.curr ?? c.prev).erpPos.map(esc).join(', ');
  const tags = [
    r.newlyDotted
      ? `<span class="tag is-reissue is-new" title="이전 파일에는 점이 없었고 최신 파일에서 점을 찍어 다시 등록되었습니다">이번에 재등록</span>`
      : `<span class="tag is-reissue" title="일부 출하 뒤 PO 끝에 점을 찍어 다시 등록한 잔여 오더입니다">잔여 재등록</span>`,
  ];
  if (r.mixed) {
    tags.push(
      `<span class="tag is-alert" title="점 있는 행과 없는 행의 수량이 한 줄로 더해졌습니다. 원래 PO의 잔여 행이 남아 두 번 더해진 것은 아닌지 확인하세요">중복 확인</span>`,
    );
  }
  return `<div class="reissue">${tags.join('')}<span class="erp-po">전산 등록 ${erp}</span></div>`;
}

function tagsHtml(c) {
  const tags = [];
  if (c.delayDays > 0) {
    tags.push(`<span class="tag${c.delayDays >= DELAY_ALERT_DAYS ? ' is-alert' : ''}">${c.delayDays}일 지연</span>`);
  } else if (c.delayDays < 0) {
    tags.push(`<span class="tag">${-c.delayDays}일 앞당김</span>`);
  }
  if (c.qtyDiff) tags.push(`<span class="tag">수량 ${signed(c.qtyDiff)}</span>`);
  return tags.join('');
}

function reasonHtml(c, sel) {
  if (!needsReason(c.type)) return '<span class="no-reason">사유 불필요</span>';
  const code = sel?.code ?? '';
  const options = reasonOptions(c.type)
    .map((r) => `<option value="${r.code}"${r.code === code ? ' selected' : ''}>${esc(r.ko)}</option>`)
    .join('');
  const placeholder = c.type === 'gone' ? '출하 완료 / 주문 취소 선택' : '사유 선택';
  return `
    <select class="reason-select${isFilled(sel) ? '' : ' is-missing'}" aria-label="변경 사유">
      <option value=""${code ? '' : ' selected'}>${placeholder}</option>${options}
    </select>
    <input class="reason-text" type="text" aria-label="직접 입력한 사유" placeholder="REMARKS에 그대로 들어갈 문장" value="${esc(sel?.text)}"${code === 'custom' ? '' : ' hidden'} />`;
}

/** 변경 목록. 각 행의 data-id는 changes 배열의 순번이다. */
export function renderList(el, items, reasons) {
  if (!items.length) {
    el.innerHTML = '<li class="empty">이 조건에 해당하는 변경 건이 없습니다.</li>';
    return;
  }
  el.innerHTML = items
    .map(({ id, change: c }) => {
      const unit = (c.curr ?? c.prev).unit;
      return `
      <li class="row" data-id="${id}">
        <span class="badge badge-${c.type}">${TYPE_LABELS[c.type]}</span>
        <div class="ident">
          <div class="po">${esc(c.po)}</div>
          <div class="part"><span>${esc(c.partNo)}</span>${esc(c.partName)}</div>
          ${reissueHtml(c)}
        </div>
        <div class="change">
          ${linesHtml(c.prev, unit)}
          <span class="arrow" aria-label="에서">→</span>
          ${linesHtml(c.curr, unit)}
          <div class="tags">${tagsHtml(c)}</div>
        </div>
        <div class="reason">${reasonHtml(c, reasons.get(c.key))}</div>
      </li>`;
    })
    .join('');
}
