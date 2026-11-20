import { describe, it, expect } from 'vitest';
import ExcelJS from 'exceljs';
import { parseExcel, parseWorkbook, matrixToRows } from '../src/parse.js';
import { prepare } from '../src/normalize.js';
import { diffOrders } from '../src/diff.js';
import { buildExportRows, buildWorkbook, exportFileName } from '../src/export.js';
import prevRaw from './fixtures/prev.json';
import currRaw from './fixtures/curr.json';

async function makeXlsx() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('출하계획');
  // 쓰지 않는 열을 섞고 열 순서를 바꿔 헤더 이름으로 찾는지 확인한다
  ws.addRow(['담당자', '고객품번', '납품처', '공장출하일', '고객PO번호', '재고', '계획수량', '수주단위', '고객품명']);
  ws.addRow(['kim', 55501, 'BETA ', new Date(Date.UTC(2026, 10, 18)), 'PO-9001.', 3, 1200, 'PCS', 'H7 12V 55W']);
  ws.addRow([]);
  ws.addRow(['lee', '000123400000', 'BETA', new Date(Date.UTC(2026, 10, 25)), 'PO-9002', 0, 800, 'PCS', 'H4 12V']);
  wb.addWorksheet('Sheet2').addRow(['고객PO번호', '무시되는 목록']);
  return wb.xlsx.writeBuffer();
}

describe('parseExcel', () => {
  it('출하계획 시트만 읽고 헤더 이름으로 열을 찾는다', async () => {
    const rows = await parseExcel(await makeXlsx());
    expect(rows).toHaveLength(2);
    expect(rows[0]).toMatchObject({ 고객PO번호: 'PO-9001.', 고객품번: 55501, 납품처: 'BETA ', 계획수량: 1200 });
    expect(rows[0]['공장출하일']).toBeInstanceOf(Date);
    expect(rows[0]).not.toHaveProperty('담당자');
  });

  it('읽은 결과가 표준화까지 이어진다(날짜 하루 밀림 없음)', async () => {
    const { orders } = prepare(await parseExcel(await makeXlsx()));
    const list = [...orders.values()];
    expect(list.map((o) => [o.shipTo, o.po, o.partNo, o.lines[0].date])).toEqual([
      ['BETA', 'PO-9001', '55501', '2026-11-18'],
      ['BETA', 'PO-9002', '000123400000', '2026-11-25'],
    ]);
  });

  it('출하계획 시트가 없으면 알려준다', async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('Sheet1').addRow(['a']);
    await expect(parseExcel(await wb.xlsx.writeBuffer())).rejects.toThrow('출하계획');
  });

  it('필수 열이 없으면 어떤 열인지 알려준다', () => {
    expect(() => matrixToRows([['납품처', '고객품번'], ['A', 'B']])).toThrow(/공장출하일.*고객PO번호.*계획수량/);
  });
});

describe('buildExportRows', () => {
  const changes = diffOrders(prepare(prevRaw).orders, prepare(currRaw).orders);
  const reasons = new Map();
  const keyOf = (po) => changes.find((c) => c.po === po).key;
  reasons.set(keyOf('PO-1003'), { code: 'production', text: '' });
  reasons.set(keyOf('PO-1005'), { code: 'partial', text: '' });
  reasons.set(keyOf('PO-1006'), { code: 'custom', text: '고객과 협의' });
  reasons.set(keyOf('PO-1009'), { code: 'cancelled', text: '' });
  const rows = buildExportRows({ shipTo: 'ALPHA', changes, reasons });
  const of = (po) => rows.filter((r) => r.po === po);

  it('선택한 납품처의 curr 전체 줄을 출하일, PO 순으로 낸다', () => {
    expect(rows.every((r) => r.shipTo === 'ALPHA')).toBe(true);
    const body = rows.filter((r) => r.status !== 'cancelled');
    expect(body).toHaveLength(11);
    expect(body.map((r) => `${r.exw} ${r.po}`)).toEqual([...body.map((r) => `${r.exw} ${r.po}`)].sort());
  });

  it('변경 없는 줄은 Previous EXW가 같고 REMARKS가 비어 있다', () => {
    expect(of('PO-1002')[0]).toMatchObject({ status: 'same', prevExw: '2026-11-20', exw: '2026-11-20', qty: 120000, remarks: '' });
  });

  it('납기 변경은 이전 날짜와 영문 사유를 넣는다', () => {
    expect(of('PO-1003')[0]).toMatchObject({
      status: 'changed',
      exwChanged: true,
      prevExw: '2026-11-18',
      exw: '2026-12-01',
      remarks: 'Delayed due to production schedule',
    });
  });

  it('수량이 바뀌면 Previous QTY를 덧붙인다', () => {
    expect(of('PO-1005')[0]).toMatchObject({
      status: 'changed',
      exwChanged: false,
      remarks: 'Partially shipped, balance pending / Previous QTY 5,000',
    });
  });

  it('분할은 넘치는 줄에 prev의 마지막 날짜를 짝짓고, 직접 입력 사유를 그대로 쓴다', () => {
    expect(of('PO-1006')).toMatchObject([
      { exw: '2026-11-18', prevExw: '2026-11-18', exwChanged: false, status: 'changed', remarks: '고객과 협의' },
      { exw: '2026-12-02', prevExw: '2026-11-18', exwChanged: true, status: 'changed', remarks: '고객과 협의' },
    ]);
  });

  it('점 찍어 재등록된 주문도 고객용 PO에는 점이 없다', () => {
    expect(rows.some((r) => r.po.endsWith('.'))).toBe(false);
    expect(of('PO-1001')).toHaveLength(1);
    expect(of('PO-1005')).toHaveLength(1);
  });

  it('신규는 Previous EXW가 빈칸', () => {
    expect(of('PO-1008')[0]).toMatchObject({ status: 'new', prevExw: '', remarks: '' });
  });

  it('주문 취소한 gone은 맨 아래, 출하 완료나 미선택 gone은 출력하지 않는다', () => {
    expect(rows.at(-1)).toMatchObject({ status: 'cancelled', po: 'PO-1009', prevExw: '2026-11-19', exw: '', remarks: 'Order cancelled' });
    const shipped = new Map(reasons).set(keyOf('PO-1009'), { code: 'shipped', text: '' });
    expect(buildExportRows({ shipTo: 'ALPHA', changes, reasons: shipped }).some((r) => r.po === 'PO-1009')).toBe(false);
    expect(buildExportRows({ shipTo: 'ALPHA', changes, reasons: new Map() }).some((r) => r.po === 'PO-1009')).toBe(false);
  });

  it('워크북을 만들고 다시 읽을 수 있다', async () => {
    const buffer = await buildWorkbook(rows, { includeDeliveryDate: false }).xlsx.writeBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    const ws = wb.getWorksheet('Pending Order');
    expect(ws.getRow(1).values.slice(1)).toEqual([
      'ORDER DATE', 'Previous EXW', 'Updated EXW', 'SHIP TO', 'PO NO.', 'PART NO.', 'PART NAME', 'QTY', 'UNIT', 'REMARKS',
    ]);
    expect(ws.rowCount).toBe(rows.length + 1);
    expect(ws.getRow(1).getCell(1).fill.fgColor.argb).toBe('FF007082');
    const first = ws.getRow(2);
    expect(first.getCell(3).value).toEqual(new Date(Date.UTC(2026, 10, 18)));
    expect(first.getCell(3).numFmt).toBe('yyyy-mm-dd');
    expect(first.getCell(8).numFmt).toBe('#,##0');
  });

  it('DELIVERY DATE 열을 넣을 수 있다', () => {
    const ws = buildWorkbook(rows).getWorksheet('Pending Order');
    expect(ws.getRow(1).getCell(2).value).toBe('DELIVERY DATE');
  });
});

describe('exportFileName', () => {
  it('납품처와 생성일로 파일명을 만든다', () => {
    expect(exportFileName('Herth+Buss', new Date(2026, 10, 20))).toBe('DELIVERY SCHEDULE_Herth+Buss_2026.11.20.xlsx');
    expect(exportFileName('A/B', new Date(2026, 0, 5))).toBe('DELIVERY SCHEDULE_A_B_2026.01.05.xlsx');
  });
});

// 직전에 고객에게 보낸 Delivery Schedule(Pending Order 시트)을 이전 파일로 올리는 경우
async function makeScheduleXlsx() {
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet('Shipped Orders').addRow(['ORDER DATE', 'EXW', 'PO NO.']);
  const ws = wb.addWorksheet('Pending Order');
  ws.addRow(['ORDER DATE', 'DELIVERY DATE', 'Previous EXW', 'Updated EXW', 'SHIP TO', 'PO NO.', 'PART NO.', 'PART NAME', 'QTY', 'UNIT', 'REMARKS']);
  const d = (m, day) => new Date(Date.UTC(2026, m - 1, day));
  ws.addRow([d(7, 1), d(10, 30), d(11, 10), d(11, 18), 'ALPHA ', 'PO-1001', 'A100', 'H4 12V 60/55W STD', 1500, 'PCS', 'INV-01 (ETD 11/20)']);
  ws.addRow([d(7, 5), d(11, 1), d(11, 18), d(11, 18), 'ALPHA ', 'PO-1003', 'A300', 'H1 12V 55W STD', 2000, 'PCS', null]);
  ws.addRow([d(7, 20), d(11, 5), d(11, 20), null, 'ALPHA ', 'PO-1012', 123400, 'H3 12V 55W STD', 900, 'PCS', 'Schedule under review']);
  return wb.xlsx.writeBuffer();
}

describe('Delivery Schedule(Pending Order)을 이전 파일로 읽기', () => {
  it('출하계획 시트가 없고 Pending Order 시트가 있으면 일정표로 읽는다', async () => {
    const result = await parseWorkbook(await makeScheduleXlsx());
    expect(result.source).toBe('schedule');
    expect(result.rows).toHaveLength(3);
    expect(result.rows[0]).toMatchObject({
      납품처: 'ALPHA ',
      고객PO번호: 'PO-1001',
      고객품번: 'A100',
      계획수량: 1500,
      이전REMARKS: 'INV-01 (ETD 11/20)',
      출처: 'schedule',
    });
    expect(result.rows[0]['공장출하일']).toEqual(new Date(Date.UTC(2026, 10, 18))); // Updated EXW
    expect(result.rows[0]['희망납기']).toEqual(new Date(Date.UTC(2026, 9, 30))); // DELIVERY DATE
  });

  it('전산자료는 그대로 erp로 읽는다', async () => {
    expect((await parseWorkbook(await makeXlsx())).source).toBe('erp');
  });

  it('두 시트가 모두 없으면 두 이름을 다 알려준다', async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('Sheet1').addRow(['a']);
    await expect(parseWorkbook(await wb.xlsx.writeBuffer())).rejects.toThrow(/출하계획.*Pending Order/);
  });

  it('일정표의 희망 납기와 REMARKS를 주문에 보관하고 EXW가 빈 줄도 읽는다', async () => {
    const { orders } = prepare(await parseExcel(await makeScheduleXlsx()));
    const list = [...orders.values()];
    expect(list.map((o) => [o.shipTo, o.po, o.partNo, o.lines[0].date])).toEqual([
      ['ALPHA', 'PO-1001', 'A100', '2026-11-18'],
      ['ALPHA', 'PO-1003', 'A300', '2026-11-18'],
      ['ALPHA', 'PO-1012', '123400', ''],
    ]);
    expect(list[0]).toMatchObject({ deliveryDate: '2026-10-30', remarksByDate: { '2026-11-18': 'INV-01 (ETD 11/20)' }, fromSchedule: true });
    expect(list[1].remarksByDate).toEqual({});
  });

  it('이전=일정표, 최신=전산자료로 비교하면 희망 납기와 변경 없는 줄의 REMARKS를 넘겨받는다', async () => {
    const prev = prepare(await parseExcel(await makeScheduleXlsx())).orders;
    const curr = prepare(currRaw).orders;
    const ch = diffOrders(prev, curr);
    const out = buildExportRows({ shipTo: 'ALPHA', changes: ch, reasons: new Map([[ch.find((c) => c.po === 'PO-1003').key, { code: 'production', text: '' }]]) });
    const row = (po) => out.find((r) => r.po === po);
    // 변경 없는 줄: 지난 표의 REMARKS 유지
    expect(row('PO-1001')).toMatchObject({ status: 'same', deliveryDate: '2026-10-30', remarks: 'INV-01 (ETD 11/20)' });
    // 변경된 줄: 이번 사유로 바뀜, 희망 납기는 유지
    expect(row('PO-1003')).toMatchObject({ status: 'changed', deliveryDate: '2026-11-01', prevExw: '2026-11-18', exw: '2026-12-01', remarks: 'Delayed due to production schedule' });
    // 지난 표에 없던 주문: 희망 납기 모름
    expect(row('PO-1008')).toMatchObject({ status: 'new', deliveryDate: '' });
  });

  it('일정표에는 점 정보가 없으므로 이번에 재등록으로 단정하지 않는다', async () => {
    const prev = prepare(await parseExcel(await makeScheduleXlsx())).orders;
    const ch = diffOrders(prev, prepare(currRaw).orders);
    expect(ch.find((c) => c.po === 'PO-1001').reissue).toEqual({ dotted: true, newlyDotted: false, mixed: false });
  });
});
