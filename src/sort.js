// 화면 표의 열 정렬. DOM에 의존하지 않는 순수 함수.
// 화면에서 보는 순서만 바꾼다. 내려받는 엑셀의 줄 순서(출하일 순)는 export.js가 정하고 여기서 건드리지 않는다.

const isEmpty = (v) => v == null || v === '';

function compare(a, b) {
  if (typeof a === 'number' && typeof b === 'number') return a - b;
  const x = String(a);
  const y = String(b);
  return x < y ? -1 : x > y ? 1 : 0;
}

/**
 * 목록을 한 열 기준으로 정렬한 새 배열을 돌려준다.
 * - 값이 빈 줄은 오름차순이든 내림차순이든 맨 뒤에 둔다(엑셀과 같다).
 * - 값이 같은 줄끼리는 원래 순서를 지킨다.
 * @param {object[]} items
 * @param {(item: object) => string|number|null|undefined} getValue 정렬에 쓸 값
 * @param {'asc'|'desc'} dir
 */
export function sortBy(items, getValue, dir = 'asc') {
  const sign = dir === 'desc' ? -1 : 1;
  return items
    .map((item, index) => ({ item, index, value: getValue(item) }))
    .sort((a, b) => {
      const ae = isEmpty(a.value);
      const be = isEmpty(b.value);
      if (ae !== be) return ae ? 1 : -1;
      if (ae) return a.index - b.index;
      return sign * compare(a.value, b.value) || a.index - b.index;
    })
    .map((x) => x.item);
}

/** 머리 칸을 누를 때마다 오름차순 → 내림차순 → 정렬 없음(엑셀 순서)으로 돈다. */
export function nextSort(current, key) {
  if (!current || current.key !== key) return { key, dir: 'asc' };
  if (current.dir === 'asc') return { key, dir: 'desc' };
  return null;
}
