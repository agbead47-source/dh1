// 변경 사유 목록(한글: 화면, 영문: REMARKS). 기업 확인 전까지 쓰는 임시 목록이다.

const ALL = ['date', 'qty', 'split', 'merge'];

export const REASONS = [
  { code: 'production', ko: '생산 일정 지연', en: 'Delayed due to production schedule', types: ['date', 'split'] },
  { code: 'material', ko: '원자재 입고 지연', en: 'Delayed due to raw material supply', types: ['date', 'split'] },
  { code: 'shipping', ko: '선적 일정 조정', en: 'Adjusted to shipping schedule', types: ['date', 'split', 'merge'] },
  { code: 'container', ko: '컨테이너 물량 조정', en: 'Adjusted for container loading', types: ['split', 'merge'] },
  { code: 'customer', ko: '고객 요청', en: "Changed at customer's request", types: ALL },
  { code: 'partial', ko: '부분 출하', en: 'Partially shipped, balance pending', types: ['qty'] },
  { code: 'air', ko: '에어 분리 출하', en: 'Partial quantity shipped separately by air', types: ['qty', 'split'] },
  { code: 'custom', ko: '직접 입력', en: '', types: ALL },
];

// 사라진 건은 출하 완료인지 주문 취소인지 담당자가 고른다.
export const GONE_OPTIONS = [
  { code: 'shipped', ko: '출하 완료', en: '' },
  { code: 'cancelled', ko: '주문 취소', en: 'Order cancelled' },
];

const BY_CODE = new Map([...REASONS, ...GONE_OPTIONS].map((r) => [r.code, r]));

export const needsReason = (type) => type !== 'new' && type !== 'same';

/** 해당 변경 유형에서 고를 수 있는 사유 목록. */
export function reasonOptions(type) {
  if (type === 'gone') return GONE_OPTIONS;
  return REASONS.filter((r) => r.types.includes(type));
}

/** 유형별 기본 사유. `date`만 기본값이 있다. */
export function defaultReason(type) {
  return { code: type === 'date' ? 'production' : '', text: '' };
}

/** 사유가 입력되었는가. 직접 입력은 문장이 있어야 입력된 것으로 본다. */
export function isFilled(sel) {
  if (!sel || !sel.code) return false;
  if (sel.code === 'custom') return Boolean(sel.text && sel.text.trim());
  return BY_CODE.has(sel.code);
}

/** 선택한 사유의 REMARKS 문구. 직접 입력은 입력한 문장 그대로. */
export function remarkFor(sel) {
  if (!isFilled(sel)) return '';
  if (sel.code === 'custom') return sel.text.trim();
  return BY_CODE.get(sel.code).en;
}
