import { describe, it, expect } from 'vitest';
import ExcelJS from 'exceljs';
import { parseExcel, parseWorkbook, matrixToRows } from '../src/parse.js';
import { prepare, normalizeShipped } from '../src/normalize.js';
import { diffOrders } from '../src/diff.js';
import { buildExportRows, buildShippedRows, buildWorkbook, exportFileName, fileLabelFor, changedOnly, addedOnly } from '../src/export.js';
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
    expect(of('PO-1008')[0]).toMatchObject({ status: 'new', prevExw: '', remarks: 'NEW ORDER' });
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
    expect(ws.getRow(1).getCell(1).fill.fgColor.argb).toBe('FFD9E1F2');
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

// Shipped Orders: 이전 일정표의 시트를 넘겨받고, 출하 완료로 고른 건을 아래에 붙인다
async function makeScheduleWithShipped() {
  const wb = new ExcelJS.Workbook();
  const d = (m, day) => new Date(Date.UTC(2026, m - 1, day));
  const sh = wb.addWorksheet('Shipped Orders');
  sh.addRow(['ORDER DATE', 'DELIVERY DATE', 'EXW', 'ETD', 'ETA', 'SHIP TO', 'Invoice No.', 'SHIPMENT MODE', 'PO NO.', 'PART NO.', 'PART NAME', 'QTY', 'UNIT', 'REMARKS']);
  sh.addRow([d(3, 2), d(7, 1), d(7, 3), d(7, 10), d(8, 20), 'ALPHA ', 'INV-0001', 'FCL', 'PO-0901', 'A900', 'H1 12V 55W STD', 3000, 'PCS', null]);
  sh.addRow([d(3, 9), d(7, 1), d(7, 3), d(7, 10), d(8, 20), 'BETA', 'INV-0002', 'AIR', 'PO-0902', 123400, 'H3 12V 55W STD', 500, 'PCS', 'Partial']);
  const ws = wb.addWorksheet('Pending Order');
  ws.addRow(['ORDER DATE', 'DELIVERY DATE', 'Previous EXW', 'Updated EXW', 'SHIP TO', 'PO NO.', 'PART NO.', 'PART NAME', 'QTY', 'UNIT', 'REMARKS']);
  ws.addRow([d(7, 18), d(11, 10), d(11, 19), d(11, 19), 'ALPHA', 'PO-1009', 'A700', 'HB4 12V 51W STD', 800, 'PCS', null]);
  ws.addRow([d(7, 5), d(11, 1), d(11, 18), d(11, 18), 'ALPHA', 'PO-1003', 'A300', 'H1 12V 55W STD', 2000, 'PCS', null]);
  return wb.xlsx.writeBuffer();
}

describe('Shipped Orders', () => {
  it('일정표의 Shipped Orders 시트를 함께 읽는다', async () => {
    const { shipped } = await parseWorkbook(await makeScheduleWithShipped());
    expect(normalizeShipped(shipped)).toEqual([
      { orderDate: '2026-03-02', deliveryDate: '2026-07-01', exw: '2026-07-03', etd: '2026-07-10', eta: '2026-08-20', shipTo: 'ALPHA', invoiceNo: 'INV-0001', shipmentMode: 'FCL', po: 'PO-0901', partNo: 'A900', partName: 'H1 12V 55W STD', qty: 3000, unit: 'PCS', remarks: '' },
      { orderDate: '2026-03-09', deliveryDate: '2026-07-01', exw: '2026-07-03', etd: '2026-07-10', eta: '2026-08-20', shipTo: 'BETA', invoiceNo: 'INV-0002', shipmentMode: 'AIR', po: 'PO-0902', partNo: '123400', partName: 'H3 12V 55W STD', qty: 500, unit: 'PCS', remarks: 'Partial' },
    ]);
  });

  it('전산자료에는 Shipped Orders가 없다', async () => {
    expect((await parseWorkbook(await makeXlsx())).shipped).toEqual([]);
  });

  it('선택한 납품처의 기존 출하 건을 넘겨받고, 출하 완료로 고른 건을 아래에 붙인다', async () => {
    const wbIn = await parseWorkbook(await makeScheduleWithShipped());
    const ch = diffOrders(prepare(wbIn.rows).orders, prepare(currRaw).orders);
    const gone = ch.find((c) => c.po === 'PO-1009');
    expect(gone.type).toBe('gone');
    const shipped = normalizeShipped(wbIn.shipped);

    const out = buildShippedRows({ shipTo: 'ALPHA', changes: ch, reasons: new Map([[gone.key, { code: 'shipped', text: '' }]]), shipped });
    expect(out).toMatchObject([
      { status: 'carried', po: 'PO-0901', invoiceNo: 'INV-0001', etd: '2026-07-10' },
      { status: 'added', key: gone.key, po: 'PO-1009', partNo: 'A700', orderDate: '2026-07-18', deliveryDate: '2026-11-10', exw: '2026-11-19', qty: 800, unit: 'PCS', etd: '', eta: '', invoiceNo: '', shipmentMode: '', remarks: '' },
    ]);

    // 주문 취소나 미선택은 Shipped Orders에 넣지 않는다
    for (const code of ['cancelled', '']) {
      const rows = buildShippedRows({ shipTo: 'ALPHA', changes: ch, reasons: new Map([[gone.key, { code, text: '' }]]), shipped });
      expect(rows.map((r) => r.po)).toEqual(['PO-0901']);
    }
  });

  it('출하 건이 있으면 워크북에 Shipped Orders 시트를 Pending Order 앞에 넣는다', async () => {
    const shippedRows = [
      { status: 'carried', orderDate: '2026-03-02', deliveryDate: '2026-07-01', exw: '2026-07-03', etd: '2026-07-10', eta: '2026-08-20', shipTo: 'ALPHA', invoiceNo: 'INV-0001', shipmentMode: 'FCL', po: 'PO-0901', partNo: 'A900', partName: 'H1', qty: 3000, unit: 'PCS', remarks: '' },
      { status: 'added', orderDate: '2026-07-18', deliveryDate: '', exw: '2026-11-19', etd: '', eta: '', shipTo: 'ALPHA', invoiceNo: '', shipmentMode: '', po: 'PO-1009', partNo: 'A700', partName: 'HB4', qty: 800, unit: 'PCS', remarks: '' },
    ];
    const buffer = await buildWorkbook([], { shippedRows }).xlsx.writeBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    expect(wb.worksheets.map((s) => s.name)).toEqual(['Shipped Orders', 'Pending Order']);
    const ws = wb.getWorksheet('Shipped Orders');
    expect(ws.getRow(1).values.slice(1)).toEqual(['ORDER DATE', 'DELIVERY DATE', 'EXW', 'ETD', 'ETA', 'SHIP TO', 'Invoice No.', 'SHIPMENT MODE', 'PO NO.', 'PART NO.', 'PART NAME', 'QTY', 'UNIT', 'REMARKS']);
    expect(ws.rowCount).toBe(3);
    expect(ws.getRow(2).getCell(4).value).toEqual(new Date(Date.UTC(2026, 6, 10)));
    expect(ws.getRow(3).getCell(4).value).toBeNull();
    expect(ws.getRow(3).getCell(9).value).toBe('PO-1009');
  });

  it('출하 건이 없으면 Pending Order 시트만 만든다', () => {
    expect(buildWorkbook([]).worksheets.map((s) => s.name)).toEqual(['Pending Order']);
    expect(buildWorkbook([], { shippedRows: [] }).worksheets.map((s) => s.name)).toEqual(['Pending Order']);
  });
});

describe('출하일이 빈 줄의 위치', () => {
  it('실제 일정표처럼 출하일 미정 줄은 날짜 있는 줄 뒤, 취소 줄 앞에 둔다', () => {
    const base = { 납품처: 'ALPHA', 고객품번: 'A1', 계획수량: 10 };
    const prev = prepare([
      { ...base, 고객PO번호: 'PO-3', 공장출하일: '2026-11-18' },
      { ...base, 고객PO번호: 'PO-4', 공장출하일: '2026-11-18' },
    ]).orders;
    const curr = prepare([
      { ...base, 고객PO번호: 'PO-1' },
      { ...base, 고객PO번호: 'PO-2', 공장출하일: '2026-12-01' },
      { ...base, 고객PO번호: 'PO-3', 공장출하일: '2026-11-18' },
    ]).orders;
    const ch = diffOrders(prev, curr);
    const cancelled = new Map([[ch.find((c) => c.po === 'PO-4').key, { code: 'cancelled', text: '' }]]);
    const out = buildExportRows({ shipTo: 'ALPHA', changes: ch, reasons: cancelled });
    expect(out.map((r) => `${r.po}:${r.exw || '미정'}:${r.status}`)).toEqual([
      'PO-3:2026-11-18:same',
      'PO-2:2026-12-01:new',
      'PO-1:미정:new',
      'PO-4:미정:cancelled',
    ]);
  });
});

// 고객 발송용 설정: 전체 납품처 한 파일, 실제 양식 서식, 파일명
describe('고객 발송용 출력', () => {
  const changes = diffOrders(prepare(prevRaw).orders, prepare(currRaw).orders);
  const reasons = new Map();

  it('납품처를 지정하지 않으면 전체 납품처를 한 표로 낸다(출하일, 납품처, PO 순)', () => {
    const all = buildExportRows({ shipTo: null, changes, reasons });
    expect(new Set(all.map((r) => r.shipTo))).toEqual(new Set(['ALPHA', 'BETA']));
    const alpha = buildExportRows({ shipTo: 'ALPHA', changes, reasons });
    const beta = buildExportRows({ shipTo: 'BETA', changes, reasons });
    expect(all).toHaveLength(alpha.length + beta.length);
    const order = all.map((r) => `${r.exw}|${r.shipTo}|${r.po}`);
    expect(order).toEqual([...order].sort());
  });

  it('전체 납품처의 Shipped Orders도 한 번에 낸다', () => {
    const shipped = [
      { shipTo: 'ALPHA', po: 'PO-0901', partNo: 'A900' },
      { shipTo: 'BETA', po: 'PO-0902', partNo: 'B900' },
    ];
    expect(buildShippedRows({ shipTo: null, changes, reasons, shipped }).map((r) => r.po)).toEqual(['PO-0901', 'PO-0902']);
    expect(buildShippedRows({ shipTo: 'BETA', changes, reasons, shipped }).map((r) => r.po)).toEqual(['PO-0902']);
  });

  it('신규 주문은 REMARKS에 NEW ORDER를 적는다', () => {
    const rows = buildExportRows({ shipTo: 'ALPHA', changes, reasons });
    expect(rows.find((r) => r.po === 'PO-1008')).toMatchObject({ status: 'new', prevExw: '', remarks: 'NEW ORDER' });
  });

  const sample = [
    { status: 'changed', exwChanged: true, orderDate: '2026-07-01', deliveryDate: '2026-10-30', prevExw: '2026-11-18', exw: '2026-12-01', shipTo: 'ALPHA', po: 'PO-1', partNo: 'A1', partName: 'x', qty: 1000, unit: 'PCS', remarks: 'r' },
    { status: 'new', exwChanged: false, orderDate: '2026-11-10', deliveryDate: '', prevExw: '', exw: '2026-12-04', shipTo: 'ALPHA', po: 'PO-2', partNo: 'A2', partName: 'y', qty: 2400, unit: 'PCS', remarks: 'NEW ORDER' },
  ];
  const fillOf = (cell) => (cell.fill?.pattern === 'solid' ? cell.fill.fgColor.argb : null);

  it('고객 발송용(기본): 배경색 없이 바뀐 Updated EXW만 빨간 굵은 글자, 헤더는 실제 양식의 연한 파란색', () => {
    const ws = buildWorkbook(sample).getWorksheet('Pending Order');
    expect(fillOf(ws.getRow(1).getCell(1))).toBe('FFD9E1F2');
    expect(ws.getRow(1).getCell(1).font).toMatchObject({ bold: true, name: '맑은 고딕', size: 11 });
    expect(fillOf(ws.getRow(2).getCell(1))).toBeNull();
    expect(fillOf(ws.getRow(3).getCell(1))).toBeNull();
    expect(ws.getRow(2).getCell(4).font).toMatchObject({ bold: true, color: { argb: 'FFFF0000' } });
    expect(ws.getRow(2).getCell(3).font.color).toBeUndefined();
    expect(ws.getRow(3).getCell(4).font.color).toBeUndefined();
    expect(ws.columns.map((c) => c.width)).toEqual([18, 21, 13, 13, 14, 26, 19, 51, 10, 6, 100]);
  });

  it('내부 검토용: 변경 줄은 노란색, 신규 줄은 초록색 배경을 더한다', () => {
    const ws = buildWorkbook(sample, { style: 'review' }).getWorksheet('Pending Order');
    expect(fillOf(ws.getRow(2).getCell(1))).toBe('FFFFF2CC');
    expect(fillOf(ws.getRow(3).getCell(1))).toBe('FFE2F0D9');
    expect(ws.getRow(2).getCell(4).font).toMatchObject({ bold: true, color: { argb: 'FFFF0000' } });
  });

  it('Shipped Orders 시트를 빼는 설정', () => {
    const shippedRows = [{ status: 'carried', orderDate: '2026-03-02', exw: '2026-07-03', shipTo: 'ALPHA', po: 'PO-0901', partNo: 'A900', qty: 1 }];
    expect(buildWorkbook(sample, { shippedRows }).worksheets.map((s) => s.name)).toEqual(['Shipped Orders', 'Pending Order']);
    expect(buildWorkbook(sample, { shippedRows, includeShipped: false }).worksheets.map((s) => s.name)).toEqual(['Pending Order']);
  });

  it('파일명에 넣을 이름: 납품처를 골랐으면 납품처, 전체면 매출처에서 기호를 뺀 이름', () => {
    expect(fileLabelFor({ shipTo: 'BETA', soldTo: 'DH-E' })).toBe('BETA');
    expect(fileLabelFor({ shipTo: null, soldTo: 'DH-E' })).toBe('DHE');
    expect(fileLabelFor({ shipTo: null, soldTo: '' })).toBe('ALL');
    expect(exportFileName(fileLabelFor({ shipTo: null, soldTo: 'DH-E' }), new Date(2026, 10, 20))).toBe('DELIVERY SCHEDULE_DHE_2026.11.20.xlsx');
  });

  it('전산자료의 매출처를 읽어 가장 많은 값을 통계에 둔다', () => {
    const rows = [
      { 매출처: 'DH-E', 납품처: 'ALPHA', 고객PO번호: 'PO-1', 고객품번: 'A1', 계획수량: 1, 공장출하일: '2026-11-18' },
      { 매출처: 'DH-E ', 납품처: 'BETA', 고객PO번호: 'PO-2', 고객품번: 'A1', 계획수량: 1, 공장출하일: '2026-11-18' },
      { 매출처: 'OTHER', 납품처: 'BETA', 고객PO번호: 'PO-3', 고객품번: 'A1', 계획수량: 1, 공장출하일: '2026-11-18' },
    ];
    expect(prepare(rows).stats.soldTo).toBe('DH-E');
    expect(prepare(currRaw).stats.soldTo).toBe('');
    expect(matrixToRows([['매출처', '공장출하일', '납품처', '고객PO번호', '고객품번', '계획수량'], ['DH-E', '2026-11-18', 'A', 'P', 'X', 1]])[0]['매출처']).toBe('DH-E');
  });
});

describe('사유가 엑셀에서 빠지지 않는다', () => {
  it('줄이 줄었는데 남은 줄이 이전과 같아도 사유와 Previous QTY를 첫 줄에 적는다', () => {
    const base = { 납품처: 'ALPHA', 고객PO번호: 'PO-7', 고객품번: 'A1' };
    const prev = prepare([
      { ...base, 공장출하일: '2026-11-18', 계획수량: 500 },
      { ...base, 공장출하일: '2026-11-25', 계획수량: 500 },
    ]).orders;
    const curr = prepare([{ ...base, 공장출하일: '2026-11-25', 계획수량: 500 }]).orders;
    const ch = diffOrders(prev, curr);
    expect(ch[0].type).toBe('qty');
    const out = buildExportRows({ shipTo: 'ALPHA', changes: ch, reasons: new Map([[ch[0].key, { code: 'partial', text: '' }]]) });
    expect(out).toMatchObject([
      { status: 'changed', exwChanged: false, exw: '2026-11-25', remarks: 'Partially shipped, balance pending / Previous total QTY 1,000' },
    ]);
  });
});

describe('넘겨받는 REMARKS', () => {
  it('지난 표의 NEW ORDER는 변경 없는 줄에 넘기지 않고, 다른 문구는 넘긴다', () => {
    const base = { 납품처: 'ALPHA', 고객품번: 'A1', 계획수량: 10, 공장출하일: '2026-11-18' };
    const prev = prepare([
      { ...base, 고객PO번호: 'PO-1', 이전REMARKS: 'NEW ORDER', 출처: 'schedule' },
      { ...base, 고객PO번호: 'PO-2', 이전REMARKS: 'INV-01 (ETD 11/20)', 출처: 'schedule' },
    ]).orders;
    const curr = prepare([
      { ...base, 고객PO번호: 'PO-1' },
      { ...base, 고객PO번호: 'PO-2' },
    ]).orders;
    const out = buildExportRows({ shipTo: null, changes: diffOrders(prev, curr), reasons: new Map() });
    expect(out.map((r) => [r.po, r.status, r.remarks])).toEqual([
      ['PO-1', 'same', ''],
      ['PO-2', 'same', 'INV-01 (ETD 11/20)'],
    ]);
  });
});

describe('수정된 부분만 내려받기', () => {
  const changes = diffOrders(prepare(prevRaw).orders, prepare(currRaw).orders);
  const gone = changes.find((c) => c.type === 'gone');

  it('변경 없는 줄을 빼고 변경, 신규, 취소 줄만 남긴다(순서는 그대로)', () => {
    const reasons = new Map([[gone.key, { code: 'cancelled', text: '' }]]);
    const all = buildExportRows({ shipTo: null, changes, reasons });
    const only = changedOnly(all);
    expect(only.length).toBeGreaterThan(0);
    expect(only.every((r) => r.status !== 'same')).toBe(true);
    expect(only).toEqual(all.filter((r) => r.status !== 'same'));
    expect(new Set(only.map((r) => r.status))).toEqual(new Set(['changed', 'new', 'cancelled']));
    // 변경 없는 주문은 한 줄도 들어가지 않는다
    expect(only.some((r) => r.po === 'PO-1002')).toBe(false);
  });

  it('Shipped Orders는 이번에 옮겨 적은 줄만 남긴다', () => {
    const reasons = new Map([[gone.key, { code: 'shipped', text: '' }]]);
    const shipped = [{ shipTo: 'ALPHA', po: 'PO-0901', partNo: 'A900' }];
    const rows = buildShippedRows({ shipTo: null, changes, reasons, shipped });
    expect(rows.map((r) => r.status)).toEqual(['carried', 'added']);
    expect(addedOnly(rows).map((r) => r.po)).toEqual([gone.po]);
  });

  it('파일명 끝에 CHANGES를 붙여 전체 파일과 구별한다', () => {
    expect(exportFileName('DHE', new Date(2026, 10, 20), 'CHANGES')).toBe('DELIVERY SCHEDULE_DHE_2026.11.20_CHANGES.xlsx');
    expect(exportFileName('DHE', new Date(2026, 10, 20))).toBe('DELIVERY SCHEDULE_DHE_2026.11.20.xlsx');
  });
});
