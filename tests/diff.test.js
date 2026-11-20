import { describe, it, expect } from 'vitest';
import { prepare, makeKey } from '../src/normalize.js';
import { diffOrders, classify, daysBetween, pairPrevDates, compareChange } from '../src/diff.js';
import prevRaw from './fixtures/prev.json';
import currRaw from './fixtures/curr.json';

const changes = diffOrders(prepare(prevRaw).orders, prepare(currRaw).orders);
const find = (shipTo, po, partNo) => changes.find((c) => c.key === makeKey(shipTo, po, partNo));
const order = (...lines) => ({ lines: lines.map(([date, qty]) => ({ date, qty })) });

describe('diffOrders (fixtures)', () => {
  it('점 있는 PO와 없는 PO는 같은 주문이라 변경 없음', () => {
    expect(find('ALPHA', 'PO-1001', 'A100').type).toBe('same');
  });

  it('품목코드 배분만 바뀌면 same', () => {
    const c = find('ALPHA', 'PO-1002', 'A200');
    expect(c.type).toBe('same');
    expect(c.curr.lines).toEqual([{ date: '2026-11-20', qty: 120000 }]);
  });

  it('날짜만 변경되면 date, delayDays 계산', () => {
    const c = find('ALPHA', 'PO-1003', 'A300');
    expect(c).toMatchObject({ type: 'date', delayDays: 13, qtyDiff: 0 });
  });

  it('날짜가 앞당겨지면 date, delayDays 음수', () => {
    expect(find('ALPHA', 'PO-1004', 'A300')).toMatchObject({ type: 'date', delayDays: -5 });
  });

  it('총수량이 바뀌면 qty, qtyDiff 계산', () => {
    expect(find('ALPHA', 'PO-1005', 'A400')).toMatchObject({ type: 'qty', qtyDiff: -2000, delayDays: 0 });
  });

  it('한 줄이 두 날짜로 나뉘면 split', () => {
    const c = find('ALPHA', 'PO-1006', 'A500');
    expect(c).toMatchObject({ type: 'split', qtyDiff: 0, delayDays: 14 });
    expect(c.curr.lines).toHaveLength(2);
  });

  it('두 줄이 한 날짜로 합쳐지면 merge', () => {
    expect(find('ALPHA', 'PO-1007', 'A600')).toMatchObject({ type: 'merge', qtyDiff: 0, delayDays: 0 });
  });

  it('curr에만 있으면 new', () => {
    const c = find('ALPHA', 'PO-1008', 'A800');
    expect(c).toMatchObject({ type: 'new', prev: null, delayDays: null, qtyDiff: null });
  });

  it('prev에만 있으면 gone', () => {
    const c = find('ALPHA', 'PO-1009', 'A700');
    expect(c).toMatchObject({ type: 'gone', curr: null, delayDays: null, qtyDiff: null });
  });

  it('0으로 시작하는 품번과 숫자 품번이 문자열 키로 맞춰진다', () => {
    expect(find('ALPHA', 'PO-1010', '000123400000').type).toBe('same');
    expect(find('ALPHA', 'PO-1011', '55501').type).toBe('same');
  });

  it('납품처 끝 공백이 달라도 같은 납품처로 비교한다', () => {
    expect(find('BETA', 'PO-2001', 'B100').type).toBe('same');
    expect(find('BETA', 'PO-2002', 'B200')).toMatchObject({ type: 'date', delayDays: 14 });
    expect(new Set(changes.map((c) => c.shipTo))).toEqual(new Set(['ALPHA', 'BETA']));
  });

  it('점 있는 행과 없는 행이 함께 있던 주문은 중복 확인 대상으로 표시한다', () => {
    // prev: PO-1001. 과 PO-1001 이 함께 있음 / curr: PO-1001. 만 있음
    expect(find('ALPHA', 'PO-1001', 'A100').reissue).toEqual({ dotted: true, newlyDotted: false, mixed: false });
    const swapped = diffOrders(prepare(currRaw).orders, prepare(prevRaw).orders);
    const c = swapped.find((x) => x.key === makeKey('ALPHA', 'PO-1001', 'A100'));
    expect(c.reissue).toEqual({ dotted: true, newlyDotted: false, mixed: true });
    expect(c.curr.erpPos).toEqual(['PO-1001', 'PO-1001.']);
  });

  it('이전에 점이 없다가 최신에 점이 생기면 신규·사라짐이 아니라 같은 주문의 변경이다', () => {
    const c = find('ALPHA', 'PO-1005', 'A400');
    expect(c).toMatchObject({ type: 'qty', qtyDiff: -2000, po: 'PO-1005' });
    expect(c.reissue).toEqual({ dotted: true, newlyDotted: true, mixed: false });
    expect(c.curr.erpPos).toEqual(['PO-1005.']);
    expect(changes.some((x) => x.po === 'PO-1005.')).toBe(false);
  });

  it('점이 없는 주문은 재등록 표시가 없다', () => {
    expect(find('ALPHA', 'PO-1003', 'A300').reissue).toEqual({ dotted: false, newlyDotted: false, mixed: false });
  });

  it('유형별 건수', () => {
    const count = {};
    for (const c of changes) count[c.type] = (count[c.type] ?? 0) + 1;
    expect(count).toEqual({ same: 5, date: 3, qty: 1, split: 1, merge: 1, new: 2, gone: 1 });
  });
});

describe('classify', () => {
  it('분할이면서 총수량도 바뀌면 split', () => {
    expect(classify(order(['2026-11-18', 1000]), order(['2026-11-18', 500], ['2026-11-25', 300]))).toBe('split');
  });

  it('줄 수가 줄고 총수량도 다르면 qty', () => {
    expect(classify(order(['2026-11-18', 500], ['2026-11-25', 500]), order(['2026-11-25', 700]))).toBe('qty');
  });

  it('줄 수가 같고 총수량이 다르면 날짜가 바뀌어도 qty', () => {
    expect(classify(order(['2026-11-18', 500]), order(['2026-11-25', 700]))).toBe('qty');
  });

  it('줄 수와 총수량이 같고 날짜별 배분만 바뀌면 date', () => {
    const prev = order(['2026-11-18', 100], ['2026-11-25', 200]);
    const curr = order(['2026-11-18', 200], ['2026-11-25', 100]);
    expect(classify(prev, curr)).toBe('date');
  });
});

describe('daysBetween', () => {
  it('월을 넘어가는 날짜 차이를 센다', () => {
    expect(daysBetween('2026-11-18', '2026-12-01')).toBe(13);
    expect(daysBetween('2026-12-01', '2026-11-18')).toBe(-13);
    expect(daysBetween('2026-12-31', '2027-01-01')).toBe(1);
  });

  it('날짜가 비면 null', () => {
    expect(daysBetween('', '2026-12-01')).toBeNull();
  });
});

describe('pairPrevDates', () => {
  const lines = (...dates) => dates.map((date) => ({ date, qty: 1 }));

  it('같은 날짜가 prev에 있으면 그 날짜', () => {
    expect(pairPrevDates(lines('2026-11-18', '2026-11-25'), lines('2026-11-25'))).toEqual(['2026-11-25']);
  });

  it('없으면 같은 순번의 prev 날짜', () => {
    expect(pairPrevDates(lines('2026-11-18', '2026-11-25'), lines('2026-11-20', '2026-12-02'))).toEqual([
      '2026-11-18',
      '2026-11-25',
    ]);
  });

  it('순번이 넘치면 prev의 마지막 날짜', () => {
    expect(pairPrevDates(lines('2026-11-18'), lines('2026-11-18', '2026-12-02', '2026-12-09'))).toEqual([
      '2026-11-18',
      '2026-11-18',
      '2026-11-18',
    ]);
  });

  it('prev가 없으면 빈칸', () => {
    expect(pairPrevDates([], lines('2026-11-18'))).toEqual(['']);
  });
});

describe('compareChange (이전·최신 나란히 비교)', () => {
  const byPo = (po) => changes.find((c) => c.po === po);

  it('출하일별로 이전 수량과 최신 수량을 맞춰 놓는다', () => {
    // 분할: 11-18 10,000 → 11-18 6,000 + 12-02 4,000
    expect(compareChange(byPo('PO-1006'))).toMatchObject({
      dates: [
        { date: '2026-11-18', prevQty: 10000, currQty: 6000, diff: -4000, state: 'changed' },
        { date: '2026-12-02', prevQty: null, currQty: 4000, diff: 4000, state: 'added' },
      ],
      prevTotal: 10000,
      currTotal: 10000,
      qtyDiff: 0,
      delayDays: 14,
    });
  });

  it('날짜가 옮겨지면 빠진 날짜와 생긴 날짜가 따로 보인다', () => {
    expect(compareChange(byPo('PO-1003')).dates).toEqual([
      { date: '2026-11-18', prevQty: 2000, currQty: null, diff: -2000, state: 'removed' },
      { date: '2026-12-01', prevQty: null, currQty: 2000, diff: 2000, state: 'added' },
    ]);
  });

  it('변경 없는 날짜는 same', () => {
    expect(compareChange(byPo('PO-1002')).dates).toEqual([
      { date: '2026-11-20', prevQty: 120000, currQty: 120000, diff: 0, state: 'same' },
    ]);
  });

  it('신규와 사라짐은 한쪽이 비어 있다', () => {
    expect(compareChange(byPo('PO-1008'))).toMatchObject({ prevTotal: null, currTotal: 2400, qtyDiff: null, prevSources: [] });
    expect(compareChange(byPo('PO-1009'))).toMatchObject({ prevTotal: 800, currTotal: null, currSources: [] });
  });

  it('전산에 등록된 원본 행(품목코드, 제조사, 수주번호, 전산 PO)을 이전·최신 각각 보여 준다', () => {
    // 품목코드 배분만 바뀐 주문: 고객 기준으로는 same이지만 전산 행은 다르다
    const d = compareChange(byPo('PO-1002'));
    expect(d.prevSources.map((s) => [s.maker, s.itemCode, s.qty])).toEqual([
      ['106', 'FX0000000003', 60000],
      ['110', 'FX0000000004', 60000],
    ]);
    expect(d.currSources.map((s) => [s.maker, s.itemCode, s.qty])).toEqual([
      ['106', 'FX0000000003', 40000],
      ['110', 'FX0000000004', 80000],
    ]);
    expect(d.sourcesChanged).toBe(true);
    // 점 찍어 재등록된 주문: 전산 PO 표기가 보인다
    expect(compareChange(byPo('PO-1005')).currSources[0]).toMatchObject({ erpPo: 'PO-1005.', date: '2026-11-20', qty: 3000, orderNo: 'T000000007' });
    expect(compareChange(byPo('PO-1003')).sourcesChanged).toBe(true);
    expect(compareChange(byPo('PO-1010')).sourcesChanged).toBe(false);
  });
});
