import { describe, it, expect } from 'vitest';
import { toDateStr, cleanPo, normalizeRows, aggregate, prepare, makeKey } from '../src/normalize.js';
import prevRaw from './fixtures/prev.json';

const row = (over = {}) => ({
  공장출하일: '2026-11-18',
  납품처: 'ALPHA',
  고객PO번호: 'PO-1',
  고객품번: 'A1',
  계획수량: 100,
  ...over,
});

describe('toDateStr', () => {
  it('Date 객체를 UTC 기준 YYYY-MM-DD로 바꾼다', () => {
    expect(toDateStr(new Date(Date.UTC(2026, 10, 18)))).toBe('2026-11-18');
    // 한국 시간 자정으로 만들어진 값도 UTC 날짜로 읽는다(하루 밀림 확인용 경계값)
    expect(toDateStr(new Date('2026-11-18T00:00:00Z'))).toBe('2026-11-18');
    expect(toDateStr(new Date('2026-11-18T23:59:59Z'))).toBe('2026-11-18');
  });

  it('여러 문자열 형식을 통일한다', () => {
    expect(toDateStr('2026-11-18')).toBe('2026-11-18');
    expect(toDateStr('2026.11.18')).toBe('2026-11-18');
    expect(toDateStr('2026/1/5')).toBe('2026-01-05');
    expect(toDateStr('20261118')).toBe('2026-11-18');
    expect(toDateStr(' 2026-11-18 00:00:00 ')).toBe('2026-11-18');
  });

  it('엑셀 일련번호와 8자리 숫자를 처리한다', () => {
    expect(toDateStr(46344)).toBe('2026-11-18');
    expect(toDateStr(20261118)).toBe('2026-11-18');
  });

  it('빈 값과 알 수 없는 값은 빈 문자열', () => {
    expect(toDateStr(null)).toBe('');
    expect(toDateStr('')).toBe('');
    expect(toDateStr('미정')).toBe('');
    expect(toDateStr(new Date('x'))).toBe('');
  });
});

describe('cleanPo', () => {
  it('끝의 점을 모두 뗀다', () => {
    expect(cleanPo('PO-1001.')).toBe('PO-1001');
    expect(cleanPo('PO-1001...')).toBe('PO-1001');
    expect(cleanPo(' PO-1001. ')).toBe('PO-1001');
  });

  it('중간의 점은 건드리지 않는다', () => {
    expect(cleanPo('PO.10.01')).toBe('PO.10.01');
    expect(cleanPo('PO.10.01.')).toBe('PO.10.01');
  });
});

describe('normalizeRows', () => {
  it('PO 끝에 점이 있는 행과 없는 행이 같은 키가 된다', () => {
    const { rows, stats } = normalizeRows([row({ 고객PO번호: 'PO-1.' }), row({ 고객PO번호: 'PO-1' })]);
    expect(rows[0].key).toBe(rows[1].key);
    expect(rows[0].po).toBe('PO-1');
    expect(stats.dotRemoved).toBe(1);
  });

  it('납품처 끝 공백을 없애 같은 납품처로 묶는다', () => {
    const { rows } = normalizeRows([row({ 납품처: 'BETA ' }), row({ 납품처: 'BETA' })]);
    expect(rows[0].shipTo).toBe('BETA');
    expect(rows[0].key).toBe(rows[1].key);
  });

  it('품번은 항상 문자열이고 앞의 0을 유지한다', () => {
    const { rows } = normalizeRows([row({ 고객품번: '002537100000' }), row({ 고객품번: 55501 })]);
    expect(rows[0].partNo).toBe('002537100000');
    expect(rows[1].partNo).toBe('55501');
  });

  it('빈 행, PO번호나 품번이 없는 행은 건너뛴다', () => {
    const { rows, stats } = normalizeRows([
      {},
      { 납품처: 'ALPHA', 계획수량: 10 },
      row({ 고객PO번호: '.' }),
      row({ 고객품번: '  ' }),
      row(),
    ]);
    expect(rows).toHaveLength(1);
    expect(stats).toMatchObject({ rawCount: 5, skipped: 4 });
  });

  it('쉼표가 든 수량 문자열을 숫자로 읽는다', () => {
    const { rows } = normalizeRows([row({ 계획수량: '59,200' })]);
    expect(rows[0].qty).toBe(59200);
  });
});

describe('aggregate', () => {
  it('같은 키, 같은 날짜, 품목코드만 다른 2행을 한 줄로 합산한다', () => {
    const { rows } = normalizeRows([
      row({ 계획수량: 59200, 품목코드: 'FX1', 제조사: '106' }),
      row({ 계획수량: 60800, 품목코드: 'FX2', 제조사: '110' }),
    ]);
    const { orders, stats } = aggregate(rows);
    expect(orders.size).toBe(1);
    const order = orders.get(makeKey('ALPHA', 'PO-1', 'A1'));
    expect(order.lines).toEqual([{ date: '2026-11-18', qty: 120000 }]);
    expect(order.total).toBe(120000);
    expect(stats).toEqual({ lineCount: 1, mergedLines: 1 });
  });

  it('같은 키인데 출하일이 다르면 날짜 오름차순으로 따로 둔다', () => {
    const { rows } = normalizeRows([
      row({ 공장출하일: '2026-11-25', 계획수량: 200 }),
      row({ 공장출하일: '2026-11-18', 계획수량: 300 }),
    ]);
    const order = [...aggregate(rows).orders.values()][0];
    expect(order.lines).toEqual([
      { date: '2026-11-18', qty: 300 },
      { date: '2026-11-25', qty: 200 },
    ]);
  });

  it('수주일자는 가장 이른 날짜를 쓴다', () => {
    const { rows } = normalizeRows([row({ 수주일자: '2026-07-10' }), row({ 수주일자: '2026-07-01' })]);
    expect([...aggregate(rows).orders.values()][0].orderDate).toBe('2026-07-01');
  });
});

describe('prepare (fixtures)', () => {
  const { orders, stats } = prepare(prevRaw);

  it('정리 통계를 낸다', () => {
    expect(stats).toMatchObject({ rawCount: 18, skipped: 2, dotRemoved: 1, lineCount: 13, mergedLines: 3 });
  });

  it('점 있는 PO와 없는 PO를 묶어 합산한다', () => {
    expect(orders.get(makeKey('ALPHA', 'PO-1001', 'A100')).lines).toEqual([{ date: '2026-11-18', qty: 1500 }]);
  });

  it('끝 공백이 다른 납품처 행을 묶는다', () => {
    expect(orders.get(makeKey('BETA', 'PO-2001', 'B100')).total).toBe(5000);
    expect([...orders.values()].some((o) => o.shipTo !== o.shipTo.trim())).toBe(false);
  });
});
