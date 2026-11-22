// 고객 양식 Delivery Schedule 엑셀 생성.
import ExcelJS from 'exceljs';
import { pairPrevDates } from './diff.js';
import { remarkFor } from './reasons.js';

export const SHEET_NAME = 'Pending Order';

const COLUMNS = [
  { key: 'orderDate', header: 'ORDER DATE', width: 18, kind: 'date' },
  { key: 'deliveryDate', header: 'DELIVERY DATE', width: 21, kind: 'date' },
  { key: 'prevExw', header: 'Previous EXW', width: 13, kind: 'date' },
  { key: 'exw', header: 'Updated EXW', width: 13, kind: 'date' },
  { key: 'shipTo', header: 'SHIP TO', width: 14 },
  { key: 'po', header: 'PO NO.', width: 26 },
  { key: 'partNo', header: 'PART NO.', width: 19 },
  { key: 'partName', header: 'PART NAME', width: 51 },
  { key: 'qty', header: 'QTY', width: 10, kind: 'number' },
  { key: 'unit', header: 'UNIT', width: 6 },
  { key: 'remarks', header: 'REMARKS', width: 100 },
];

export const SHIPPED_SHEET_NAME = 'Shipped Orders';

const SHIPPED_COLUMNS = [
  { key: 'orderDate', header: 'ORDER DATE', width: 14, kind: 'date' },
  { key: 'deliveryDate', header: 'DELIVERY DATE', width: 22, kind: 'date' },
  { key: 'exw', header: 'EXW', width: 13, kind: 'date' },
  { key: 'etd', header: 'ETD', width: 13, kind: 'date' },
  { key: 'eta', header: 'ETA', width: 13, kind: 'date' },
  { key: 'shipTo', header: 'SHIP TO', width: 14 },
  { key: 'invoiceNo', header: 'Invoice No.', width: 13 },
  { key: 'shipmentMode', header: 'SHIPMENT MODE', width: 14 },
  { key: 'po', header: 'PO NO.', width: 26 },
  { key: 'partNo', header: 'PART NO.', width: 17 },
  { key: 'partName', header: 'PART NAME', width: 51 },
  { key: 'qty', header: 'QTY', width: 11, kind: 'number' },
  { key: 'unit', header: 'UNIT', width: 6 },
  { key: 'remarks', header: 'REMARKS', width: 40 },
];

const COLOR = {
  header: 'FFD9E1F2', // 실제 일정표의 헤더(엑셀 기본 파랑의 연한 색)
  changed: 'FFFFF2CC',
  new: 'FFE2F0D9',
  cancelled: 'FFEDEDED',
  red: 'FFFF0000', // 실제 일정표가 쓰는 빨강
  border: 'FFC8C8C8',
};

export const NEW_ORDER_REMARK = 'NEW ORDER';

const fmtQty = (n) => n.toLocaleString('en-US');
const cmp = (a, b) => (a < b ? -1 : a > b ? 1 : 0);

/** 지난 일정표에서 넘겨받는 REMARKS. 지난번의 NEW ORDER는 이제 신규가 아니므로 넘기지 않는다. */
function carriedRemark(prevOrder, date) {
  const remark = prevOrder?.remarksByDate?.[date] ?? '';
  return remark.trim().toUpperCase() === NEW_ORDER_REMARK ? '' : remark;
}

function joinRemarks(...parts) {
  return parts.filter(Boolean).join(' / ');
}

/**
 * 출력 행을 만든다(엑셀과 무관한 순수 데이터). shipTo가 null이면 전체 납품처를 한 표로 낸다(실제 일정표 방식).
 * status: 'same' | 'changed' | 'new' | 'cancelled'. key와 type은 화면에서 줄과 변경 건을 잇는 데 쓴다(엑셀에는 나가지 않는다).
 * @param {{shipTo: string|null, changes: object[], reasons: Map<string, {code:string, text:string}>}} args
 */
export function buildExportRows({ shipTo, changes, reasons }) {
  const rows = [];
  const cancelled = [];

  for (const c of changes) {
    if (shipTo != null && c.shipTo !== shipTo) continue;
    const sel = reasons.get(c.key);

    if (c.type === 'gone') {
      if (sel?.code !== 'cancelled') continue; // 출하 완료 또는 미선택은 출력하지 않는다
      for (const line of c.prev.lines) {
        cancelled.push({
          key: c.key,
          type: c.type,
          status: 'cancelled',
          exwChanged: false,
          orderDate: c.prev.orderDate,
          deliveryDate: c.prev.deliveryDate || '',
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

    // 줄 단위로 바뀌었는지 본다. 주문은 바뀌었는데(예: 두 줄 중 한 줄이 빠짐) 남은 줄이 모두
    // 이전과 같으면 사유를 적을 곳이 없어지므로, 그때는 첫 줄을 변경 줄로 삼는다.
    const changedFlags = c.curr.lines.map(
      (line) => !isNew && !isSame && !prevLines.some((p) => p.date === line.date && p.qty === line.qty),
    );
    if (!isNew && !isSame && !changedFlags.includes(true)) changedFlags[0] = true;

    c.curr.lines.forEach((line, i) => {
      const lineChanged = changedFlags[i];
      rows.push({
        key: c.key,
        type: c.type,
        status: isNew ? 'new' : lineChanged ? 'changed' : 'same',
        exwChanged: lineChanged && prevDates[i] !== line.date,
        orderDate: c.curr.orderDate,
        // 고객 희망 납기는 전산자료에 없어 지난 일정표에서 넘겨받는다
        deliveryDate: c.curr.deliveryDate || c.prev?.deliveryDate || '',
        prevExw: isNew ? '' : prevDates[i],
        exw: line.date,
        shipTo: c.shipTo,
        po: c.po,
        partNo: c.partNo,
        partName: c.partName,
        qty: line.qty,
        unit: c.curr.unit,
        // 신규는 NEW ORDER, 변경된 줄은 이번 사유, 변경 없는 줄은 지난 일정표의 REMARKS를 그대로 둔다
        remarks: isNew
          ? NEW_ORDER_REMARK
          : lineChanged
            ? joinRemarks(remarkFor(sel), qtyNote)
            : carriedRemark(c.prev, line.date),
      });
    });
  }

  // 출하일 오름차순. 출하일이 정해지지 않은 줄은 실제 일정표처럼 맨 뒤에 둔다.
  rows.sort(
    (a, b) =>
      cmp(!a.exw, !b.exw) || cmp(a.exw, b.exw) || cmp(a.shipTo, b.shipTo) || cmp(a.po, b.po) || cmp(a.partNo, b.partNo),
  );
  cancelled.sort(
    (a, b) => cmp(a.prevExw, b.prevExw) || cmp(a.shipTo, b.shipTo) || cmp(a.po, b.po) || cmp(a.partNo, b.partNo),
  );
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
 * Shipped Orders 시트의 행을 만든다.
 * - 이전 일정표에 있던 선택 납품처의 출하 건은 그대로 넘겨받는다(status: 'carried').
 * - 사라진 주문 중 담당자가 "출하 완료"로 고른 건을 아래에 붙인다(status: 'added').
 *   ETD, ETA, Invoice No., SHIPMENT MODE는 전산 출하계획에 없어 빈칸으로 둔다.
 * @param {{shipTo: string, changes: object[], reasons: Map, shipped?: object[]}} args shipped는 normalizeShipped 결과
 */
export function buildShippedRows({ shipTo, changes, reasons, shipped = [] }) {
  const wanted = (name) => shipTo == null || name === shipTo;
  const carried = shipped.filter((r) => wanted(r.shipTo)).map((r) => ({ ...r, status: 'carried' }));
  const added = [];
  for (const c of changes) {
    if (!wanted(c.shipTo) || c.type !== 'gone' || reasons.get(c.key)?.code !== 'shipped') continue;
    for (const line of c.prev.lines) {
      added.push({
        key: c.key,
        status: 'added',
        orderDate: c.prev.orderDate,
        deliveryDate: c.prev.deliveryDate || '',
        exw: line.date,
        etd: '',
        eta: '',
        shipTo: c.shipTo,
        invoiceNo: '',
        shipmentMode: '',
        po: c.po,
        partNo: c.partNo,
        partName: c.partName,
        qty: line.qty,
        unit: c.prev.unit,
        remarks: '',
      });
    }
  }
  added.sort((a, b) => cmp(a.exw, b.exw) || cmp(a.po, b.po) || cmp(a.partNo, b.partNo));
  return [...carried, ...added];
}

// 실제 일정표는 모든 칸이 맑은 고딕 11 굵은 글자다.
const BASE_FONT = { name: '맑은 고딕', size: 11, bold: true };

function addSheet(wb, name, cols, rows, style = () => ({})) {
  const ws = wb.addWorksheet(name, { views: [{ state: 'frozen', ySplit: 1 }] });
  ws.columns = cols.map((c) => ({ key: c.key, width: c.width }));

  const thin = { style: 'thin', color: { argb: COLOR.border } };
  const border = { top: thin, left: thin, bottom: thin, right: thin };

  const head = ws.addRow(cols.map((c) => c.header));
  head.height = 22;
  head.eachCell((cell) => {
    cell.fill = solid(COLOR.header);
    cell.font = { ...BASE_FONT };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = border;
  });

  for (const r of rows) {
    const row = ws.addRow(cols.map((c) => (c.kind === 'date' ? toExcelDate(r[c.key]) : r[c.key] === '' ? null : r[c.key])));
    cols.forEach((c, i) => {
      const cell = row.getCell(i + 1);
      const { fill, font } = style(r, c);
      cell.border = border;
      cell.alignment = {
        vertical: 'middle',
        horizontal: c.kind === 'date' || c.key === 'unit' ? 'center' : undefined,
        shrinkToFit: c.key === 'partName' || undefined,
        wrapText: c.key === 'remarks' || undefined,
      };
      if (c.kind === 'date') cell.numFmt = 'yyyy-mm-dd';
      if (c.kind === 'number') cell.numFmt = '#,##0';
      if (fill) cell.fill = solid(fill);
      cell.font = { ...BASE_FONT, ...font };
    });
  }

  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: cols.length } };
  return ws;
}

/**
 * 출력 행으로 서식 있는 워크북을 만든다.
 *
 * 발송용 설정(options)
 * - style: 'customer'(기본)는 고객에게 보내는 실제 양식. 배경색 없이 바뀐 Updated EXW만 빨간 굵은 글자.
 *          'review'는 내부 검토용. 변경 줄 노랑, 신규 줄 초록, 취소 줄 회색, 옮겨 적은 출하 줄 초록 배경을 더한다.
 * - includeDeliveryDate: DELIVERY DATE 열을 넣을지(원하는 고객에게만 보낸다)
 * - includeShipped: Shipped Orders 시트를 넣을지. 넣을 줄(shippedRows)이 없으면 시트를 만들지 않는다.
 * @param {object[]} rows buildExportRows 결과
 * @param {{style?: 'customer'|'review', includeDeliveryDate?: boolean, includeShipped?: boolean, shippedRows?: object[]}} options
 */
export function buildWorkbook(rows, { style = 'customer', includeDeliveryDate = true, includeShipped = true, shippedRows = [] } = {}) {
  const keep = (c) => includeDeliveryDate || c.key !== 'deliveryDate';
  const review = style === 'review';
  const wb = new ExcelJS.Workbook();

  if (includeShipped && shippedRows.length) {
    addSheet(wb, SHIPPED_SHEET_NAME, SHIPPED_COLUMNS.filter(keep), shippedRows, (r) => ({
      fill: review && r.status === 'added' ? COLOR.new : undefined,
    }));
  }

  const fills = { changed: COLOR.changed, new: COLOR.new, cancelled: COLOR.cancelled };
  addSheet(wb, SHEET_NAME, COLUMNS.filter(keep), rows, (r, c) => ({
    fill: review ? fills[r.status] : undefined,
    font: c.key === 'exw' && r.exwChanged ? { color: { argb: COLOR.red } } : undefined,
  }));
  return wb;
}

/** 파일명에 넣을 고객 이름. 납품처를 골랐으면 그 납품처, 전체면 매출처에서 기호를 뺀 이름(DH-E → DHE). */
export function fileLabelFor({ shipTo, soldTo }) {
  if (shipTo != null) return shipTo;
  return String(soldTo ?? '').replace(/[^0-9A-Za-z가-힣]/g, '') || 'ALL';
}

/** `DELIVERY SCHEDULE_{이름}_{YYYY.MM.DD}.xlsx` (이름은 fileLabelFor, 날짜는 생성일, 사용자 PC 기준) */
export function exportFileName(shipTo, now = new Date(), suffix = '') {
  const pad = (n) => String(n).padStart(2, '0');
  const date = `${now.getFullYear()}.${pad(now.getMonth() + 1)}.${pad(now.getDate())}`;
  const safe = shipTo.replace(/[\\/:*?"<>|]/g, '_');
  return `DELIVERY SCHEDULE_${safe}_${date}${suffix ? `_${suffix}` : ''}.xlsx`;
}

// ---------- 수정된 부분만 내려받기 ----------

/** 수정된 부분만 담은 파일의 이름 끝에 붙여 전체 일정표와 구별한다. */
export const CHANGES_SUFFIX = 'CHANGES';

/** Pending Order 줄 중 변경, 신규, 취소 줄만 남긴다(변경 없는 줄을 뺀다). 순서는 그대로다. */
export function changedOnly(rows) {
  return rows.filter((r) => r.status !== 'same');
}

/** Shipped Orders 줄 중 이번에 옮겨 적은 줄만 남긴다(넘겨받은 지난 출하 기록을 뺀다). */
export function addedOnly(shippedRows) {
  return shippedRows.filter((r) => r.status === 'added');
}
