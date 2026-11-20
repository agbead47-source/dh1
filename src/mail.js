// 고객 안내 메일 초안(영문). DOM에 의존하지 않는 순수 함수.
// 정해진 문장에 숫자와 주문 정보를 채워 넣을 뿐이고, 외부로 아무것도 보내지 않는다.
// 메일을 실제로 보내지도 않는다. 담당자가 복사해서 사내 메일에 붙여 넣는다.
import { remarkFor } from './reasons.js';

export const DEFAULT_DELAY_DAYS = 10;

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const pad = (n) => String(n).padStart(2, '0');
const orders = (n) => `${n} order${n === 1 ? '' : 's'}`;
const last = (lines) => (lines.length ? lines[lines.length - 1].date : '');

/**
 * 변경 내용으로 안내 메일 초안을 만든다.
 * @param {object} args
 * @param {object[]} args.changes diffOrders 결과
 * @param {Map<string, {code:string, text:string}>} args.reasons 변경 키 → 사유
 * @param {string|null} args.shipTo 납품처. null이면 전체
 * @param {string} args.label 제목에 넣을 고객 이름(파일명에 쓰는 이름과 같다)
 * @param {Date} [args.date] 기준일(메일을 쓰는 날)
 * @param {number} [args.delayDays] 이 일수 이상 늦어진 주문을 본문에 따로 적는다
 * @param {string} [args.recipient] 받는 사람 호칭. 비우면 Customer
 * @param {string} [args.sender] 보내는 사람 이름. 비우면 서명 없이 끝낸다
 * @returns {{subject: string, body: string, summary: object, delayed: object[]}}
 */
export function buildMailDraft({
  changes,
  reasons,
  shipTo = null,
  label = '',
  date = new Date(),
  delayDays = DEFAULT_DELAY_DAYS,
  recipient = '',
  sender = '',
}) {
  const scoped = changes.filter((c) => shipTo == null || c.shipTo === shipTo);
  const count = (test) => scoped.filter(test).length;
  const chosen = (c) => reasons.get(c.key)?.code;

  const delayed = scoped
    .filter((c) => c.curr && c.prev && c.type !== 'same' && c.delayDays >= delayDays)
    .sort((a, b) => b.delayDays - a.delayDays || (a.po < b.po ? -1 : a.po > b.po ? 1 : 0))
    .map((c) => ({
      key: c.key,
      po: c.po,
      partNo: c.partNo,
      partName: c.partName,
      from: last(c.prev.lines),
      to: last(c.curr.lines),
      days: c.delayDays,
      reason: remarkFor(reasons.get(c.key)),
    }));

  const summary = {
    schedule: count((c) => ['date', 'split', 'merge'].includes(c.type)),
    quantity: count((c) => c.type === 'qty'),
    added: count((c) => c.type === 'new'),
    cancelled: count((c) => c.type === 'gone' && chosen(c) === 'cancelled'),
    shipped: count((c) => c.type === 'gone' && chosen(c) === 'shipped'),
    delayed: delayed.length,
  };

  const ymd = `${date.getFullYear()}.${pad(date.getMonth() + 1)}.${pad(date.getDate())}`;
  const longDate = `${MONTHS[date.getMonth()]} ${date.getDate()}, ${date.getFullYear()}`;
  const subject = `Updated Delivery Schedule${label ? ` - ${label}` : ''} (${ymd})`;

  const parts = [`Dear ${recipient.trim() || 'Customer'},`, `Please find attached the updated delivery schedule as of ${longDate}.`];

  const lines = [
    summary.schedule && `- Delivery date changed: ${orders(summary.schedule)}`,
    summary.quantity && `- Quantity changed: ${orders(summary.quantity)}`,
    summary.added && `- New orders: ${orders(summary.added)}`,
    summary.shipped && `- Shipped (moved to the Shipped Orders sheet): ${orders(summary.shipped)}`,
    summary.cancelled && `- Cancelled: ${orders(summary.cancelled)}`,
  ].filter(Boolean);

  if (lines.length) {
    parts.push(['Summary of changes since the previous schedule:', ...lines].join('\n'));
  } else {
    parts.push('There are no changes since the previous schedule.');
  }

  if (delayed.length) {
    parts.push(
      [
        `The following orders are delayed by ${delayDays} days or more:`,
        ...delayed.map((d) => {
          const name = d.partName ? ` (${d.partName})` : '';
          const reason = d.reason ? ` ${d.reason}` : '';
          return `- PO ${d.po} / ${d.partNo}${name}: EXW ${d.from} -> ${d.to}, ${d.days} days.${reason}`;
        }),
      ].join('\n'),
      'We apologize for the inconvenience. Please let us know if any of these orders need a different arrangement.',
    );
  }

  if (lines.length) parts.push('Changed EXW dates are marked in red in the Pending Order sheet.');
  parts.push(['Best regards,', sender.trim()].filter(Boolean).join('\n'));

  return { subject, body: `${parts.join('\n\n')}\n`, summary, delayed };
}
