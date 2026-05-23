const amountKeywords = ['金額', '残高', '売上高', '営業利益', '固定費', '変動費', '単価', '差異額', '小計', '合計', '入力値'];
const scoreKeywords = ['行別得点', '問題得点', '得点', '満点'];
const japaneseTextKeywords = ['借方科目', '貸方科目', '勘定科目', '項目名', '摘要', '処理区分', '部門名', '工程', '表区分', '差異区分'];

export function toHalfWidthNumber(value: string): string {
  return value
    .replace(/[０-９]/g, (char) => String.fromCharCode(char.charCodeAt(0) - 0xfee0))
    .replace(/[，、]/g, ',')
    .replace(/[－ー―]/g, '-')
    .replace(/[￥円]/g, '');
}

export function normalizeNumericInput(value: string): string {
  const half = toHalfWidthNumber(value);
  const cleaned = half.replace(/[,$\s]/g, '');
  if (cleaned === '' || /^-?\d*$/.test(cleaned)) return cleaned;
  return half;
}

function addCommas(integerText: string): string {
  const negative = integerText.startsWith('-');
  const body = negative ? integerText.slice(1) : integerText;
  if (!body) return integerText;
  return `${negative ? '-' : ''}${body.replace(/\B(?=(\d{3})+(?!\d))/g, ',')}`;
}

export function formatAmount(value: string): string {
  const normalized = normalizeNumericInput(value);
  if (normalized === '' || normalized === '-') return normalized;
  if (/^-?\d+$/.test(normalized)) return addCommas(normalized);
  return toHalfWidthNumber(value);
}

export function isAmountColumn(columnName: string): boolean {
  const normalized = columnName.trim();
  if (scoreKeywords.some((keyword) => normalized.includes(keyword))) return false;
  return amountKeywords.some((keyword) => normalized.includes(keyword));
}

export function isJapaneseTextColumn(columnName: string): boolean {
  const normalized = columnName.trim();
  return japaneseTextKeywords.some((keyword) => normalized.includes(keyword));
}

export function isInvalidAmount(value: string): boolean {
  if (!value) return false;
  const normalized = normalizeNumericInput(value);
  return normalized !== '' && !/^-?\d*$/.test(normalized);
}
