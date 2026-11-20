import { describe, it, expect } from 'vitest';
import { prepare } from '../src/normalize.js';
import { diffOrders } from '../src/diff.js';
import { buildMailDraft } from '../src/mail.js';
import prevRaw from './fixtures/prev.json';
import currRaw from './fixtures/curr.json';

const changes = diffOrders(prepare(prevRaw).orders, prepare(currRaw).orders);
const keyOf = (po) => changes.find((c) => c.po === po).key;
const reasons = new Map([
  [keyOf('PO-1003'), { code: 'production', text: '' }],
  [keyOf('PO-1006'), { code: 'custom', text: 'Split per container plan' }],
  [keyOf('PO-1009'), { code: 'cancelled', text: '' }],
]);
const date = new Date(2026, 10, 20);
const simple = (over = {}) => buildMailDraft({ changes, reasons, shipTo: null, label: 'DHE', date, ...over });
// 아래 buildMailDraft 묶음은 상세 내용을 넣는 선택 사항(detail: true)을 본다
const draft = (over = {}) => simple({ detail: true, ...over });

describe('buildMailDraft 기본(간단한 안내)', () => {
  it('일정표를 첨부했다는 짧은 안내만 적고, 변경 요약과 주문 목록은 넣지 않는다', () => {
    const d = simple({ sender: 'Sales Admin Team' });
    expect(d.subject).toBe('Updated Delivery Schedule - DHE (2026.11.20)');
    expect(d.body).toBe(
      [
        'Dear Customer,',
        'Please find attached the updated delivery schedule as of November 20, 2026.',
        'Changed EXW dates are marked in red in the Pending Order sheet.',
        'Please let us know if you have any questions.',
        'Best regards,\nSales Admin Team',
      ].join('\n\n') + '\n',
    );
    expect(d.body).not.toContain('PO-');
  });

  it('변경이 없으면 변경 없음으로 적는다', () => {
    const same = diffOrders(prepare(currRaw).orders, prepare(currRaw).orders);
    const d = buildMailDraft({ changes: same, reasons: new Map(), shipTo: null, label: 'DHE', date });
    expect(d.body).toContain('There are no changes since the previous schedule.');
    expect(d.body).not.toContain('marked in red');
  });

  it('요약 숫자는 화면에서 쓸 수 있게 계속 돌려준다', () => {
    expect(simple().summary).toEqual({ schedule: 5, quantity: 1, added: 2, cancelled: 1, shipped: 0, delayed: 3 });
  });
});

describe('buildMailDraft', () => {
  it('제목에 고객 이름과 날짜가 들어간다', () => {
    expect(draft().subject).toBe('Updated Delivery Schedule - DHE (2026.11.20)');
  });

  it('첨부 안내와 기준일, 변경 요약(주문 수)을 적는다', () => {
    const { body, summary } = draft();
    expect(body).toContain('Please find attached the updated delivery schedule as of November 20, 2026.');
    expect(summary).toEqual({ schedule: 5, quantity: 1, added: 2, cancelled: 1, shipped: 0, delayed: 3 });
    expect(body).toContain('- Delivery date changed: 5 orders');
    expect(body).toContain('- Quantity changed: 1 order');
    expect(body).toContain('- New orders: 2 orders');
    expect(body).toContain('- Cancelled: 1 order');
    expect(body).not.toContain('Shipped');
  });

  it('기준일 이상 지연된 주문만, 많이 늦어진 순으로 사유와 함께 따로 적는다', () => {
    const { body, delayed } = draft();
    expect(delayed.map((d) => [d.po, d.days])).toEqual([
      ['PO-1006', 14],
      ['PO-2002', 14],
      ['PO-1003', 13],
    ]);
    expect(body).toContain('The following orders are delayed by 10 days or more:');
    expect(body).toContain('- PO PO-1003 / A300 (H1 12V 55W STD): EXW 2026-11-18 -> 2026-12-01, 13 days. Delayed due to production schedule');
    expect(body).toContain('- PO PO-1006 / A500 (H11 12V 55W STD): EXW 2026-11-18 -> 2026-12-02, 14 days. Split per container plan');
    // 사유를 안 고른 건은 사유 없이 적는다
    expect(body).toContain('- PO PO-2002 / B200 (H7 12V 55W LONGLIFE): EXW 2026-11-24 -> 2026-12-08, 14 days.\n');
    // 앞당겨진 건과 기준 미만은 넣지 않는다
    expect(body).not.toContain('PO-1004');
  });

  it('지연 기준일을 바꿀 수 있고, 해당 주문이 없으면 그 문단을 뺀다', () => {
    expect(draft({ delayDays: 14 }).delayed.map((d) => d.po)).toEqual(['PO-1006', 'PO-2002']);
    const none = draft({ delayDays: 30 });
    expect(none.delayed).toEqual([]);
    expect(none.body).not.toContain('delayed by');
  });

  it('납품처를 고르면 그 납품처의 주문만 센다', () => {
    const beta = draft({ shipTo: 'BETA', label: 'BETA' });
    expect(beta.subject).toBe('Updated Delivery Schedule - BETA (2026.11.20)');
    expect(beta.summary).toEqual({ schedule: 1, quantity: 0, added: 1, cancelled: 0, shipped: 0, delayed: 1 });
    expect(beta.body).not.toContain('PO-1003');
  });

  it('받는 사람과 보내는 사람을 넣을 수 있고, 비워 두면 기본 문구를 쓴다', () => {
    const named = draft({ recipient: 'Mr. Kim', sender: 'Sales Admin Team' });
    expect(named.body.startsWith('Dear Mr. Kim,\n')).toBe(true);
    expect(named.body.trimEnd().endsWith('Best regards,\nSales Admin Team')).toBe(true);
    expect(draft().body.startsWith('Dear Customer,\n')).toBe(true);
    expect(draft().body.trimEnd().endsWith('Best regards,')).toBe(true);
  });

  it('출하 완료로 고른 건은 Shipped Orders로 옮겼다고 적는다', () => {
    const shipped = new Map(reasons).set(keyOf('PO-1009'), { code: 'shipped', text: '' });
    const d = draft({ reasons: shipped });
    expect(d.summary).toMatchObject({ cancelled: 0, shipped: 1 });
    expect(d.body).toContain('- Shipped (moved to the Shipped Orders sheet): 1 order');
  });

  it('변경이 하나도 없으면 변경 없음으로 적는다', () => {
    const same = diffOrders(prepare(currRaw).orders, prepare(currRaw).orders);
    const d = buildMailDraft({ changes: same, reasons: new Map(), shipTo: null, label: 'DHE', date, detail: true });
    expect(d.body).toContain('There are no changes since the previous schedule.');
    expect(d.body).not.toContain('Summary of changes');
  });
});
