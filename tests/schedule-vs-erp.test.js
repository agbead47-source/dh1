// 이 도구의 기본 흐름: 지난번에 고객에게 보낸 배송일정표(이전)와 최신 전산자료(최신)를 비교한다.
import { describe, it, expect } from 'vitest';
import { prepare, normalizeShipped } from '../src/normalize.js';
import { diffOrders } from '../src/diff.js';
import { buildExportRows, buildShippedRows } from '../src/export.js';
import { slotNotice, EXPECTED_SOURCE } from '../src/parse.js';
import scheduleRaw from './fixtures/schedule.json';
import shippedRaw from './fixtures/shipped.json';
import erpRaw from './fixtures/curr.json';

const schedule = prepare(scheduleRaw);
const erp = prepare(erpRaw);
const changes = diffOrders(schedule.orders, erp.orders);
const byPo = (po) => changes.find((c) => c.po === po);

describe('배송일정표(이전) × 전산자료(최신)', () => {
  it('일정표 한 줄과 전산의 여러 행(공급처별, 점 찍힌 PO)이 같은 주문으로 맞춰진다', () => {
    // 전산: PO-1002가 106, 110 두 행 / 일정표: 한 줄 120,000
    expect(byPo('PO-1002').type).toBe('same');
    // 전산: PO-1001. (점) / 일정표: PO-1001
    expect(byPo('PO-1001')).toMatchObject({ type: 'same', reissue: { dotted: true, newlyDotted: false } });
  });

  it('변경 유형을 가려낸다', () => {
    const count = {};
    for (const c of changes) count[c.type] = (count[c.type] ?? 0) + 1;
    expect(count).toEqual({ same: 5, date: 3, qty: 1, split: 1, merge: 1, new: 2, gone: 1 });
    expect(byPo('PO-1003')).toMatchObject({ type: 'date', delayDays: 13 });
    expect(byPo('PO-1005')).toMatchObject({ type: 'qty', qtyDiff: -2000 });
    expect(byPo('PO-1009').type).toBe('gone');
  });

  it('일정표에는 점 정보가 없으므로 이번에 재등록으로 단정하지 않는다', () => {
    expect(changes.some((c) => c.reissue.newlyDotted)).toBe(false);
    expect(byPo('PO-1005').reissue).toMatchObject({ dotted: true, newlyDotted: false });
  });

  it('새 일정표의 Previous EXW는 보냈던 일정표의 날짜, Updated EXW는 전산의 날짜다', () => {
    const rows = buildExportRows({ shipTo: null, changes, reasons: new Map() });
    expect(rows.find((r) => r.po === 'PO-1003')).toMatchObject({ prevExw: '2026-11-18', exw: '2026-12-01', exwChanged: true });
  });

  it('전산에 없는 고객 희망 납기와 지난 REMARKS를 일정표에서 넘겨받는다', () => {
    const rows = buildExportRows({ shipTo: null, changes, reasons: new Map() });
    const row = (po) => rows.find((r) => r.po === po);
    expect(row('PO-1002')).toMatchObject({ status: 'same', deliveryDate: '2026-11-15', remarks: 'INV-2611-01 (ETD 11/25, ETA 12/28)' });
    expect(row('PO-2001')).toMatchObject({ status: 'same', remarks: 'AIR AGREED' });
    // 지난번의 NEW ORDER는 넘기지 않는다
    expect(row('PO-1010')).toMatchObject({ status: 'same', remarks: '' });
    // 신규 주문은 일정표에 없었으므로 희망 납기를 모른다
    expect(row('PO-1008')).toMatchObject({ status: 'new', deliveryDate: '', remarks: 'NEW ORDER' });
  });

  it('일정표의 Shipped Orders를 넘겨받고, 출하 완료로 고른 건을 붙인다', () => {
    const shipped = normalizeShipped(shippedRaw);
    const gone = byPo('PO-1009');
    const rows = buildShippedRows({ shipTo: null, changes, reasons: new Map([[gone.key, { code: 'shipped', text: '' }]]), shipped });
    expect(rows.map((r) => `${r.status}:${r.shipTo}:${r.po}`)).toEqual([
      'carried:ALPHA:PO-0901',
      'carried:ALPHA:PO-0901',
      'carried:BETA:PO-0902',
      'added:ALPHA:PO-1009',
    ]);
    expect(rows.at(-1)).toMatchObject({ deliveryDate: '2026-11-12', exw: '2026-11-19', etd: '', invoiceNo: '' });
  });
});

describe('올린 칸과 파일 종류가 맞는지 알려 준다', () => {
  it('이전 칸은 보냈던 일정표, 최신 칸은 전산자료를 기대한다', () => {
    expect(EXPECTED_SOURCE).toEqual({ prev: 'schedule', curr: 'erp' });
    expect(slotNotice('prev', 'schedule')).toBe('');
    expect(slotNotice('curr', 'erp')).toBe('');
  });

  it('칸과 종류가 다르면 막지 않고 안내만 한다', () => {
    expect(slotNotice('prev', 'erp')).toContain('전산자료');
    expect(slotNotice('prev', 'erp')).toContain('보냈던 일정표');
    expect(slotNotice('curr', 'schedule')).toContain('일정표');
    expect(slotNotice('curr', 'schedule')).toContain('전산자료');
  });
});
