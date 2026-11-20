// 이전·최신 비교와 변경 분류. DOM에 의존하지 않는 순수 함수.

const DAY_MS = 86400000;

export const totalQty = (lines) => lines.reduce((s, l) => s + l.qty, 0);

export function linesEqual(a, b) {
  return a.length === b.length && a.every((l, i) => l.date === b[i].date && l.qty === b[i].qty);
}

/** `YYYY-MM-DD` 두 날짜의 차이(to - from, 일). 날짜가 비어 있으면 null. */
export function daysBetween(from, to) {
  if (!from || !to) return null;
  return Math.round((Date.parse(`${to}T00:00:00Z`) - Date.parse(`${from}T00:00:00Z`)) / DAY_MS);
}

const lastDate = (lines) => (lines.length ? lines[lines.length - 1].date : '');

/** 한 키의 이전(prev)·최신(curr) 주문을 보고 변경 유형을 정한다. 판정 순서는 CLAUDE.md 5.3. */
export function classify(prev, curr) {
  if (!prev) return 'new';
  if (!curr) return 'gone';
  if (linesEqual(prev.lines, curr.lines)) return 'same';
  if (curr.lines.length > prev.lines.length) return 'split';
  if (totalQty(curr.lines) !== totalQty(prev.lines)) return 'qty';
  if (curr.lines.length < prev.lines.length) return 'merge';
  // 줄 수와 총수량이 같다. 날짜가 바뀌었거나, 날짜별 수량 배분이 바뀐 경우.
  return 'date';
}

/**
 * 합산한 이전·최신 주문 Map을 비교한다. 변경 없는 건(`same`)도 결과에 넣는다.
 * @returns {Array<{key, type, shipTo, po, partNo, partName, prev, curr, delayDays, qtyDiff}>}
 */
export function diffOrders(prevOrders, currOrders) {
  const keys = new Set([...prevOrders.keys(), ...currOrders.keys()]);
  const changes = [];
  for (const key of keys) {
    const prev = prevOrders.get(key) ?? null;
    const curr = currOrders.get(key) ?? null;
    const ref = curr ?? prev;
    const both = prev && curr;
    changes.push({
      key,
      type: classify(prev, curr),
      shipTo: ref.shipTo,
      po: ref.po,
      partNo: ref.partNo,
      partName: ref.partName,
      prev,
      curr,
      delayDays: both ? daysBetween(lastDate(prev.lines), lastDate(curr.lines)) : null,
      qtyDiff: both ? totalQty(curr.lines) - totalQty(prev.lines) : null,
    });
  }
  const firstDate = (c) => (c.curr ?? c.prev).lines[0]?.date ?? '';
  const cmp = (a, b) => (a < b ? -1 : a > b ? 1 : 0);
  changes.sort(
    (a, b) =>
      cmp(a.shipTo, b.shipTo) || cmp(firstDate(a), firstDate(b)) || cmp(a.po, b.po) || cmp(a.partNo, b.partNo),
  );
  return changes;
}

/**
 * curr 각 줄에 대응하는 Previous EXW를 정한다.
 * 같은 날짜가 prev에 있으면 그 날짜, 없으면 같은 순번의 prev 날짜, 순번이 넘치면 prev의 마지막 날짜.
 */
export function pairPrevDates(prevLines, currLines) {
  if (!prevLines.length) return currLines.map(() => '');
  const prevDates = new Set(prevLines.map((l) => l.date));
  return currLines.map((line, i) => {
    if (prevDates.has(line.date)) return line.date;
    return prevLines[Math.min(i, prevLines.length - 1)].date;
  });
}
