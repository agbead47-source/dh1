// 고객 양식 Delivery Schedule 엑셀 생성.
import ExcelJS from 'exceljs';
import { pairPrevDates } from './diff.js';
import { remarkFor } from './reasons.js';

export const SHEET_NAME = 'Pending Order';

const COLUMNS = [
  { key: 'orderDate', header: 'ORDER DATE', width: 13, kind: 'date' },
  { key: 'deliveryDate', header: 'DELIVERY DATE', width: 15, kind: 'date' },
  { key: 'prevExw', header: 'Previous EXW', width: 14, kind: 'date' },
  { key: 'exw', header: 'Updated EXW', width: 14, kind: 'date' },
  { key: 'shipTo', header: 'SHIP TO', width: 16 },
  { key: 'po', header: 'PO NO.', width: 22 },
  { key: 'partNo', header: 'PART NO.', width: 18 },
  { key: 'partName', header: 'PART NAME', width: 36 },
  { key: 'qty', header: 'QTY', width: 12, kind: 'number' },
  { key: 'unit', header: 'UNIT', width: 8 },
  { key: 'remarks', header: 'REMARKS', width: 52 },
];

const COLOR = {
  header: 'FF007082',
  white: 'FFFFFFFF',
  changed: 'FFFFF2CC',
  new: 'FFE2F0D9',
  cancelled: 'FFEDEDED',
  red: 'FFCD0A2D',
  border: 'FFC8C8C8',
};

const fmtQty = (n) => n.toLocaleString('en-US');
const cmp = (a, b) => (a < b ? -1 : a > b ? 1 : 0);

function joinRemarks(...parts) {
  return parts.filter(Boolean).join(' / ');
}

/**
 * 선택한 납품처의 출력 행을 만든다(엑셀과 무관한 순수 데이터).
 * status: 'same' | 'changed' | 'new' | 'cancelled'
 * @param {{shipTo: string, changes: object[], reasons: Map<string, {code:string, text:string}>}} args
 */
export function buildExportRows({ shipTo, changes, reasons }) {
  const rows = [];
  const cancelled = [];

  for (const c of changes) {
    if (c.shipTo !== shipTo) continue;
    const sel = reasons.get(c.key);

    if (c.type === 'gone') {
      if (sel?.code !== 'cancelled') continue; // 출하 완료 또는 미선택은 출력하지 않는다
      for (const line of c.prev.lines) {
        cancelled.push({
          status: 'cancelled',
          exwChanged: false,
          orderDate: c.prev.orderDate,
          deliveryDate: '',
          prevExw: line.date,
          exw: '',
          shipTo: c.shipTo,
          po: c.po,
          partNo: c.partNo,
          partName: c.partName,
          qty: line.qty,
          unit: c.prev.unit,
          remarks: remarkFor(sel),
        });
      }
      continue;
    }

    const isNew = c.type === 'new';
    const isSame = c.type === 'same';
    const prevLines = c.prev?.lines ?? [];
    const prevDates = pairPrevDates(prevLines, c.curr.lines);
    const single = prevLines.length === 1 && c.curr.lines.length === 1;
    const qtyNote =
      !isNew && c.qtyDiff ? `Previous ${single ? '' : 'total '}QTY ${fmtQty(c.prev.total)}` : '';

    c.curr.lines.forEach((line, i) => {
      const lineChanged =
        !isNew && !isSame && !prevLines.some((p) => p.date === line.date && p.qty === line.qty);
      rows.push({
        status: isNew ? 'new' : lineChanged ? 'changed' : 'same',
        exwChanged: lineChanged && prevDates[i] !== line.date,
        orderDate: c.curr.orderDate,
        deliveryDate: '',
        prevExw: isNew ? '' : prevDates[i],
        exw: line.date,
        shipTo: c.shipTo,
        po: c.po,
        partNo: c.partNo,
        partName: c.partName,
        qty: line.qty,
        unit: c.curr.unit,
        remarks: lineChanged ? joinRemarks(remarkFor(sel), qtyNote) : '',
      });
    });
  }

  rows.sort((a, b) => cmp(a.exw, b.exw) || cmp(a.po, b.po) || cmp(a.partNo, b.partNo));
  cancelled.sort((a, b) => cmp(a.prevExw, b.prevExw) || cmp(a.po, b.po) || cmp(a.partNo, b.partNo));
  return [...rows, ...cancelled];
}

// 엑셀은 날짜를 UTC 자정 기준으로 저장한다.
function toExcelDate(str) {
  if (!str) return null;
  const [y, m, d] = str.split('-').map(Number);
  return new Date(Date.UTC(y, m - 1, d));
}

const solid = (argb) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });

/**
 * 출력 행으로 서식 있는 워크북을 만든다.
 * @param {object[]} rows buildExportRows 결과
 * @param {{includeDeliveryDate?: boolean}} options
 */
export function buildWorkbook(rows, { includeDeliveryDate = true } = {}) {
  const cols = COLUMNS.filter((c) => includeDeliveryDate || c.key !== 'deliveryDate');
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(SHEET_NAME, { views: [{ state: 'frozen', ySplit: 1 }] });
  ws.columns = cols.map((c) => ({ key: c.key, width: c.width }));

  const thin = { style: 'thin', color: { argb: COLOR.border } };
  const border = { top: thin, left: thin, bottom: thin, right: thin };

  const head = ws.addRow(cols.map((c) => c.header));
  head.height = 22;
  head.eachCell((cell) => {
    cell.fill = solid(COLOR.header);
    cell.font = { bold: true, color: { argb: COLOR.white } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = border;
  });

  const fills = { changed: COLOR.changed, new: COLOR.new, cancelled: COLOR.cancelled };
  for (const r of rows) {
    const row = ws.addRow(cols.map((c) => (c.kind === 'date' ? toExcelDate(r[c.key]) : r[c.key])));
    cols.forEach((c, i) => {
      const cell = row.getCell(i + 1);
      cell.border = border;
      cell.alignment = { vertical: 'middle', horizontal: c.kind === 'date' || c.key === 'unit' ? 'center' : undefined };
      if (c.kind === 'date') cell.numFmt = 'yyyy-mm-dd';
      if (c.kind === 'number') cell.numFmt = '#,##0';
      if (fills[r.status]) cell.fill = solid(fills[r.status]);
      if (c.key === 'exw' && r.exwChanged) cell.font = { bold: true, color: { argb: COLOR.red } };
    });
  }

  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: cols.length } };
  return wb;
}

/** `DELIVERY SCHEDULE_{납품처}_{YYYY.MM.DD}.xlsx` (날짜는 생성일, 사용자 PC 기준) */
export function exportFileName(shipTo, now = new Date()) {
  const pad = (n) => String(n).padStart(2, '0');
  const date = `${now.getFullYear()}.${pad(now.getMonth() + 1)}.${pad(now.getDate())}`;
  const safe = shipTo.replace(/[\\/:*?"<>|]/g, '_');
  return `DELIVERY SCHEDULE_${safe}_${date}.xlsx`;
}
