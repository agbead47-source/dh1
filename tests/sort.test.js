import { describe, it, expect } from 'vitest';
import { sortBy, nextSort } from '../src/sort.js';

const rows = [
  { po: 'PO-3', exw: '2026-12-01', qty: 500 },
  { po: 'PO-1', exw: '', qty: 12000 },
  { po: 'PO-2', exw: '2026-11-18', qty: 900 },
  { po: 'PO-4', exw: '2026-11-18', qty: 900 },
];
const pos = (list) => list.map((r) => r.po);

describe('sortBy', () => {
  it('글자 열을 오름차순, 내림차순으로 정렬한다', () => {
    expect(pos(sortBy(rows, (r) => r.po, 'asc'))).toEqual(['PO-1', 'PO-2', 'PO-3', 'PO-4']);
    expect(pos(sortBy(rows, (r) => r.po, 'desc'))).toEqual(['PO-4', 'PO-3', 'PO-2', 'PO-1']);
  });

  it('숫자 열은 글자 순서가 아니라 크기로 정렬한다', () => {
    expect(sortBy(rows, (r) => r.qty, 'asc').map((r) => r.qty)).toEqual([500, 900, 900, 12000]);
    expect(sortBy(rows, (r) => r.qty, 'desc').map((r) => r.qty)).toEqual([12000, 900, 900, 500]);
  });

  it('빈 값은 어느 방향이든 맨 뒤에 둔다', () => {
    expect(pos(sortBy(rows, (r) => r.exw, 'asc'))).toEqual(['PO-2', 'PO-4', 'PO-3', 'PO-1']);
    expect(pos(sortBy(rows, (r) => r.exw, 'desc'))).toEqual(['PO-3', 'PO-2', 'PO-4', 'PO-1']);
  });

  it('값이 같은 줄은 원래 순서를 지키고, 원본 배열은 바꾸지 않는다', () => {
    const before = pos(rows);
    expect(pos(sortBy(rows, (r) => r.qty, 'asc')).slice(1, 3)).toEqual(['PO-2', 'PO-4']);
    expect(pos(sortBy(rows, (r) => r.qty, 'desc')).slice(1, 3)).toEqual(['PO-2', 'PO-4']);
    expect(pos(rows)).toEqual(before);
  });
});

describe('nextSort', () => {
  it('오름차순 → 내림차순 → 정렬 없음 순서로 돈다', () => {
    const a = nextSort(null, 'qty');
    expect(a).toEqual({ key: 'qty', dir: 'asc' });
    const b = nextSort(a, 'qty');
    expect(b).toEqual({ key: 'qty', dir: 'desc' });
    expect(nextSort(b, 'qty')).toBeNull();
  });

  it('다른 열을 누르면 그 열의 오름차순부터 시작한다', () => {
    expect(nextSort({ key: 'qty', dir: 'desc' }, 'po')).toEqual({ key: 'po', dir: 'asc' });
  });
});
