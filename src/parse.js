// 엑셀 → 원본 행 배열. DOM에 의존하지 않는다.
// 원본 행은 헤더 이름(한글)을 키로 갖는 객체다. 값 정리는 normalize.js가 한다.
import ExcelJS from 'exceljs';

export const SHEET_NAME = '출하계획';

export const COLUMNS = [
  '공장출하일',
  '납품처',
  '고객PO번호',
  '고객품번',
  '고객품명',
  '계획수량',
  '수주단위',
  '수주일자',
  '품목코드',
  '제조사',
  '수주번호',
];

// 이 열이 없으면 비교 자체가 불가능하다.
export const REQUIRED_COLUMNS = ['공장출하일', '납품처', '고객PO번호', '고객품번', '계획수량'];

// 직전에 고객에게 보낸 Delivery Schedule을 이전 파일로 올릴 수도 있다.
// 그 파일의 Pending Order 시트 열을 전산자료와 같은 이름으로 바꿔 읽는다.
export const SCHEDULE_SHEET_NAME = 'Pending Order';

export const SCHEDULE_COLUMNS = {
  'Updated EXW': '공장출하일',
  'SHIP TO': '납품처',
  'PO NO.': '고객PO번호',
  'PART NO.': '고객품번',
  'PART NAME': '고객품명',
  QTY: '계획수량',
  UNIT: '수주단위',
  'ORDER DATE': '수주일자',
  'DELIVERY DATE': '희망납기',
  REMARKS: '이전REMARKS',
};

const headerName = (v) => String(v ?? '').replace(/\s+/g, '').toLowerCase();

/** ExcelJS 셀 값(수식, 서식 있는 글자, 링크 등)을 단순 값으로 바꾼다. */
export function cellToValue(v) {
  if (v == null) return null;
  if (v instanceof Date) return v;
  if (typeof v === 'object') {
    if (Array.isArray(v.richText)) return v.richText.map((t) => t.text ?? '').join('');
    if ('result' in v) return cellToValue(v.result);
    if ('formula' in v || 'sharedFormula' in v) return null; // 계산 결과가 없는 수식
    if ('text' in v) return cellToValue(v.text);
    if ('error' in v) return null;
    return null;
  }
  return v;
}

const isBlank = (v) => v == null || (typeof v === 'string' && v.trim() === '');

/**
 * 2차원 배열(첫 행이 헤더)을 원본 행 객체 배열로 바꾼다.
 * 열 순서가 아니라 헤더 이름으로 열을 찾는다. 사용하는 열이 모두 빈 행은 버린다.
 * @param {Record<string,string>} [rename] 파일의 헤더 이름 → 원본 행의 키. 없으면 전산자료 열 그대로.
 */
export function matrixToRows(matrix, rename) {
  const pairs = rename ? Object.entries(rename) : COLUMNS.map((name) => [name, name]);
  const header = (matrix[0] ?? []).map((v) => headerName(cellToValue(v)));
  const index = {};
  for (const [label, name] of pairs) {
    const i = header.indexOf(headerName(label));
    if (i >= 0) index[name] = i;
  }
  const missing = pairs.filter(([, name]) => REQUIRED_COLUMNS.includes(name) && !(name in index)).map(([label]) => label);
  if (missing.length) {
    throw new Error(`필수 열을 찾지 못했습니다: ${missing.join(', ')}`);
  }

  const rows = [];
  for (let r = 1; r < matrix.length; r++) {
    const line = matrix[r] ?? [];
    const row = {};
    let hasValue = false;
    for (const [, name] of pairs) {
      if (!(name in index)) continue;
      const v = cellToValue(line[index[name]]);
      if (isBlank(v)) continue;
      row[name] = v;
      hasValue = true;
    }
    if (hasValue) rows.push(row);
  }
  return rows;
}

function sheetToMatrix(ws) {
  const matrix = [];
  const colCount = ws.columnCount;
  for (let r = 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const line = new Array(colCount);
    for (let c = 1; c <= colCount; c++) line[c - 1] = row.getCell(c).value;
    matrix.push(line);
  }
  return matrix;
}

/**
 * xlsx 파일 내용(ArrayBuffer)을 읽는다.
 * - `출하계획` 시트가 있으면 전산자료(source: 'erp')
 * - 없고 `Pending Order` 시트가 있으면 고객에게 보냈던 일정표(source: 'schedule')
 * @returns {Promise<{source: 'erp'|'schedule', sheet: string, rows: object[]}>}
 */
export async function parseWorkbook(buffer) {
  const wb = new ExcelJS.Workbook();
  try {
    await wb.xlsx.load(buffer);
  } catch (err) {
    throw new Error('엑셀 파일을 읽지 못했습니다. xlsx 형식인지 확인해 주세요.', { cause: err });
  }

  const erp = wb.getWorksheet(SHEET_NAME);
  if (erp) return { source: 'erp', sheet: SHEET_NAME, rows: matrixToRows(sheetToMatrix(erp)) };

  const schedule = wb.getWorksheet(SCHEDULE_SHEET_NAME);
  if (schedule) {
    const rows = matrixToRows(sheetToMatrix(schedule), SCHEDULE_COLUMNS);
    for (const row of rows) row['출처'] = 'schedule';
    return { source: 'schedule', sheet: SCHEDULE_SHEET_NAME, rows };
  }

  const names = wb.worksheets.map((s) => s.name).join(', ');
  throw new Error(
    `'${SHEET_NAME}'(전산자료) 또는 '${SCHEDULE_SHEET_NAME}'(보냈던 일정표) 시트가 없습니다. (파일의 시트: ${names || '없음'})`,
  );
}

/** xlsx 파일 내용을 읽어 원본 행 배열만 돌려준다. */
export async function parseExcel(buffer) {
  return (await parseWorkbook(buffer)).rows;
}
