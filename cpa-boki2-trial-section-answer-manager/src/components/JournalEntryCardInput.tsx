import type { AnswerRow, AnswerState, TemplateDefinition } from '../types';
import { formatAmount, normalizeNumericInput, toHalfWidthNumber } from '../utils/numberFormat';

interface Props {
  answer: AnswerState;
  template: TemplateDefinition;
  onChange: (answer: AnswerState) => void;
}

function makeRow(columnCount: number, index: number): AnswerRow {
  const cells = Array.from({ length: columnCount }, (_, col) => (col === 0 ? String(index + 1) : ''));
  return { id: crypto.randomUUID(), cells, grade: '未採点', points: 0 };
}

function findColumnIndex(columns: string[], patterns: string[], fallback: number): number {
  const index = columns.findIndex((column) => patterns.some((pattern) => column.includes(pattern)));
  return index >= 0 ? index : fallback;
}

export function JournalEntryCardInput({ answer, template, onChange }: Props) {
  const columns = answer.columns;
  const debitAccountIndex = findColumnIndex(columns, ['借方科目'], 2);
  const debitAmountIndex = findColumnIndex(columns, ['借方金額'], 3);
  const creditAccountIndex = findColumnIndex(columns, ['貸方科目'], 4);
  const creditAmountIndex = findColumnIndex(columns, ['貸方金額'], 5);

  const updateCell = (rowIndex: number, colIndex: number, value: string, amount = false) => {
    const rows = answer.rows.map((row, index) => {
      if (index !== rowIndex) return row;
      const cells = [...row.cells];
      cells[colIndex] = amount ? normalizeNumericInput(value) : value;
      return { ...row, cells };
    });
    onChange({ ...answer, rows });
  };

  const addEntry = () => {
    const rows = [...answer.rows, makeRow(answer.columns.length, answer.rows.length)];
    onChange({ ...answer, rows });
  };

  const deleteEntry = (rowIndex: number) => {
    const rows = answer.rows.filter((_, index) => index !== rowIndex);
    onChange({ ...answer, rows: rows.length ? rows : [makeRow(answer.columns.length, 0)] });
  };

  const renderAccountInput = (row: AnswerRow, rowIndex: number, colIndex: number, label: string) => {
    const value = row.cells[colIndex] ?? '';
    const options = template.optionColumns?.[columns[colIndex] ?? ''];
    return (
      <label>{label}
        {options ? (
          <select value={value} onChange={(event) => updateCell(rowIndex, colIndex, event.target.value)}>
            <option value="">選択</option>
            {options.map((option) => <option key={option} value={option}>{option}</option>)}
          </select>
        ) : (
          <input value={value} onChange={(event) => updateCell(rowIndex, colIndex, event.target.value)} placeholder="例：仕掛品" lang="ja" autoComplete="off" />
        )}
      </label>
    );
  };

  const renderAmountInput = (row: AnswerRow, rowIndex: number, colIndex: number, label: string) => {
    const value = row.cells[colIndex] ?? '';
    return (
      <label>{label}
        <input
          value={formatAmount(value)}
          onChange={(event) => updateCell(rowIndex, colIndex, toHalfWidthNumber(event.target.value), true)}
          inputMode="numeric"
          pattern="[0-9,\\-]*"
          placeholder="例：10,000"
          autoComplete="off"
        />
      </label>
    );
  };

  return (
    <section className="mobile-journal-input" aria-label="スマホ用仕訳入力">
      <div className="mobile-journal-heading">
        <p className="eyebrow">答案入力</p>
        <h2>仕訳カード入力</h2>
      </div>
      {answer.rows.map((row, rowIndex) => (
        <article className="journal-entry-card" key={row.id}>
          <div className="journal-entry-card-header">
            <strong>仕訳{rowIndex + 1}</strong>
            <button className="secondary" onClick={() => deleteEntry(rowIndex)} disabled={answer.rows.length <= 1}>削除</button>
          </div>
          <div className="journal-entry-side">
            <h3>借方</h3>
            {renderAccountInput(row, rowIndex, debitAccountIndex, '科目')}
            {renderAmountInput(row, rowIndex, debitAmountIndex, '金額')}
          </div>
          <div className="journal-entry-side">
            <h3>貸方</h3>
            {renderAccountInput(row, rowIndex, creditAccountIndex, '科目')}
            {renderAmountInput(row, rowIndex, creditAmountIndex, '金額')}
          </div>
        </article>
      ))}
      <button className="accent add-journal-entry" onClick={addEntry}>＋ 仕訳を追加</button>
    </section>
  );
}
