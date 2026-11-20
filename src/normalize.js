// 표준화와 고객 기준 합산. DOM에 의존하지 않는 순수 함수.

const DAY_MS = 86400000;
const KEY_SEP = '\u001f';
const pad = (n) => String(n).padStart(2, '0');

function ymd(y, m, d) {
  if (!(m >= 1 && m <= 12 && d >= 1 && d <= 31)) return '';
  return `${y}-${pad(m)}-${pad(d)}`;
}

/** 날짜 값을 `YYYY-MM-DD` 문자열로 바꾼다. Date는 UTC 기준으로 읽어 하루 밀림을 막는다. */
export function toDateStr(v) {
  if (v == null || v === '') return '';
  if (v instanceof Date) {
    if (Number.isNaN(v.getTime())) return '';
    return ymd(v.getUTCFullYear(), v.getUTCMonth() + 1, v.getUTCDate());
  }
  if (typeof v === 'number') {
    // 엑셀 날짜 일련번호(서식이 풀린 날짜 칸)
    if (v > 0 && v < 2958466) return toDateStr(new Date(Date.UTC(1899, 11, 30) + Math.floor(v) * DAY_MS));
    v = String(v); // 20261118 같은 숫자
  }
  const m = String(v).trim().match(/^(\d{4})\s*[-./]?\s*(\d{1,2})\s*[-./]?\s*(\d{1,2})/);
  return m ? ymd(Number(m[1]), Number(m[2]), Number(m[3])) : '';
}

const text = (v) => (v == null ? '' : String(v).trim());

/** PO번호 끝의 점을 모두 뗀다. 중간의 점은 그대로 둔다. */
export function cleanPo(v) {
  return text(v).replace(/\.+$/, '').trim();
}

function toQty(v) {
  if (typeof v === 'number') return Number.isFinite(v) ? v : 0;
  const n = Number(text(v).replace(/,/g, ''));
  return Number.isFinite(n) ? n : 0;
}

export function makeKey(shipTo, po, partNo) {
  return [shipTo, po, partNo].join(KEY_SEP);
}

/**
 * 원본 행을 표준화한다. PO번호나 품번이 없는 행은 건너뛴다.
 * @returns {{rows: object[], stats: {rawCount:number, skipped:number, dotRemoved:number, noDate:number}}}
 */
export function normalizeRows(rawRows) {
  const rows = [];
  const stats = { rawCount: rawRows.length, skipped: 0, dotRemoved: 0, noDate: 0 };
  for (const raw of rawRows) {
    const rawPo = text(raw['고객PO번호']);
    const po = cleanPo(rawPo);
    const partNo = text(raw['고객품번']);
    if (!po || !partNo) {
      stats.skipped++;
      continue;
    }
    if (po !== rawPo) stats.dotRemoved++;
    const date = toDateStr(raw['공장출하일']);
    if (!date) stats.noDate++;
    const shipTo = text(raw['납품처']);
    rows.push({
      key: makeKey(shipTo, po, partNo),
      shipTo,
      po,
      partNo,
      partName: text(raw['고객품명']),
      date,
      qty: toQty(raw['계획수량']),
      unit: text(raw['수주단위']),
      orderDate: toDateStr(raw['수주일자']),
      itemCode: text(raw['품목코드']),
      maker: text(raw['제조사']),
      orderNo: text(raw['수주번호']),
    });
  }
  return { rows, stats };
}

/**
 * 표준화한 행을 고객 기준으로 합산한다.
 * 같은 키, 같은 출하일은 수량을 더해 한 줄로, 출하일이 다르면 줄을 나눈다.
 * @returns {{orders: Map<string, object>, stats: {lineCount:number, mergedLines:number}}}
 */
export function aggregate(rows) {
  const orders = new Map();
  for (const r of rows) {
    let o = orders.get(r.key);
    if (!o) {
      o = {
        key: r.key,
        shipTo: r.shipTo,
        po: r.po,
        partNo: r.partNo,
        partName: '',
        unit: '',
        orderDate: '',
        byDate: new Map(),
      };
      orders.set(r.key, o);
    }
    if (!o.partName) o.partName = r.partName;
    if (!o.unit) o.unit = r.unit;
    if (r.orderDate && (!o.orderDate || r.orderDate < o.orderDate)) o.orderDate = r.orderDate;
    const line = o.byDate.get(r.date) ?? { date: r.date, qty: 0, sources: 0 };
    line.qty += r.qty;
    line.sources++;
    o.byDate.set(r.date, line);
  }

  const stats = { lineCount: 0, mergedLines: 0 };
  for (const o of orders.values()) {
    const all = [...o.byDate.values()].sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));
    stats.lineCount += all.length;
    stats.mergedLines += all.filter((l) => l.sources > 1).length;
    o.lines = all.map(({ date, qty }) => ({ date, qty }));
    o.total = o.lines.reduce((s, l) => s + l.qty, 0);
    delete o.byDate;
  }
  return { orders, stats };
}

/** 원본 행 → 합산된 주문 Map과 정리 통계. */
export function prepare(rawRows) {
  const { rows, stats } = normalizeRows(rawRows);
  const { orders, stats: aggStats } = aggregate(rows);
  return { orders, stats: { ...stats, ...aggStats } };
}
