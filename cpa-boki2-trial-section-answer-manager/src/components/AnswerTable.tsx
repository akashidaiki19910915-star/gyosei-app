import type { AnswerRow, AnswerState, GradeMark, TemplateDefinition } from '../types';
import { formatAmount, isAmountColumn, isInvalidAmount, isJapaneseTextColumn, normalizeNumericInput, toHalfWidthNumber } from '../utils/numberFormat';
import { handleTableCellKeyDown } from '../utils/tableNavigation';

const gradeOptions: GradeMark[] = ['未採点', '○', '△', '×'];

interface Props {
  answer: AnswerState;
  template: TemplateDefinition;
  onChange: (answer: AnswerState) => void;
  onApplyRowPoints: () => void;
}

function makeRow(columnCount: number, index: number): AnswerRow {
  const cells = Array.from({ length: columnCount }, (_, col) => (col === 0 ? String(index + 1) : ''));
  return { id: crypto.randomUUID(), cells, grade: '未採点', points: 0 };
}

export function createRows(columnCount: number, count: number): AnswerRow[] {
  return Array.from({ length: count }, (_, index) => makeRow(columnCount, index));
}

function columnClass(column: string): string {
  if (column === '行番号') return 'col-row-number';
  if (column.includes('設問番号')) return 'col-question-number';
  if (isAmountColumn(column)) return 'col-amount';
  if (column.includes('メモ')) return 'col-memo';
  if (isJapaneseTextColumn(column)) return 'col-account';
  return 'col-generic';
}

function cellInputMode(column: string): React.HTMLAttributes<HTMLInputElement>['inputMode'] {
  if (isAmountColumn(column)) return 'numeric';
  return 'text';
}

function placeholderFor(column: string): string {
  if (isAmountColumn(column)) return '例：200,000';
  if (isJapaneseTextColumn(column)) return column.includes('科目') ? '例：仕入' : '日本語入力';
  return '';
}

export function AnswerTable({ answer, template, onChange, onApplyRowPoints }: Props) {
  const updateHeader = (index: number, value: string) => {
    const columns = [...answer.columns];
    columns[index] = value;
    onChange({ ...answer, columns, rows: answer.rows.map((row) => ({ ...row, cells: row.cells.slice(0, columns.length) })) });
  };

  const updateCell = (rowIndex: number, colIndex: number, value: string) => {
    const column = answer.columns[colIndex] ?? '';
    const nextValue = isAmountColumn(column) ? normalizeNumericInput(value) : value;
    const rows = answer.rows.map((row, index) => {
      if (index !== rowIndex) return row;
      const cells = [...row.cells];
      cells[colIndex] = nextValue;
      return { ...row, cells };
    });
    onChange({ ...answer, rows });
  };

  const updateCellOnBlur = (rowIndex: number, colIndex: number, value: string) => {
    const column = answer.columns[colIndex] ?? '';
    if (!isAmountColumn(column)) return;
    const normalized = normalizeNumericInput(value);
    const rows = answer.rows.map((row, index) => {
      if (index !== rowIndex) return row;
      const cells = [...row.cells];
      cells[colIndex] = normalized;
      return { ...row, cells };
    });
    onChange({ ...answer, rows });
  };

  const updateGrade = (rowIndex: number, grade: GradeMark) => {
    const rows = answer.rows.map((row, index) => index === rowIndex ? { ...row, grade } : row);
    onChange({ ...answer, rows });
  };

  const updatePoints = (rowIndex: number, points: number) => {
    const rows = answer.rows.map((row, index) => index === rowIndex ? { ...row, points } : row);
    onChange({ ...answer, rows });
  };

  const addRows = () => {
    const start = answer.rows.length;
    const rows = [...answer.rows, ...Array.from({ length: 10 }, (_, index) => makeRow(answer.columns.length, start + index))];
    onChange({ ...answer, rows });
  };

  const addColumn = () => {
    const columns = [...answer.columns, `追加列${answer.columns.length + 1}`];
    const rows = answer.rows.map((row) => ({ ...row, cells: [...row.cells, ''] }));
    onChange({ ...answer, columns, rows });
  };

  const compactRows = () => {
    const rows = answer.rows.filter((row) => row.cells.some((cell, index) => index > 0 && cell.trim() !== '') || row.grade !== '未採点' || Number(row.points) !== 0);
    onChange({ ...answer, rows: rows.length ? rows : createRows(answer.columns.length, 1) });
  };

  const renderCellInput = (row: AnswerRow, rowIndex: number, column: string, colIndex: number) => {
    const value = row.cells[colIndex] ?? '';
    const amount = isAmountColumn(column);
    const invalidAmount = amount && isInvalidAmount(value);
    const commonProps = {
      'data-answer-cell': 'true',
      'data-row-index': rowIndex,
      'data-col-index': colIndex,
      onKeyDown: (event: React.KeyboardEvent<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>) => handleTableCellKeyDown(event, rowIndex, colIndex),
    };

    if (template.optionColumns?.[column]) {
      return (
        <select
          {...commonProps}
          value={value}
          onChange={(event) => updateCell(rowIndex, colIndex, event.target.value)}
          aria-label={`${column} ${rowIndex + 1}行目`}
        >
          <option value="">選択</option>
          {template.optionColumns[column].map((option) => <option key={option} value={option}>{option}</option>)}
        </select>
      );
    }

    if (column.includes('メモ')) {
      return (
        <textarea
          {...commonProps}
          className="memo-cell-input"
          value={value}
          title={value}
          onChange={(event) => updateCell(rowIndex, colIndex, event.target.value)}
          lang="ja"
          inputMode="text"
          autoComplete="off"
          spellCheck={false}
          aria-label={`${column} ${rowIndex + 1}行目`}
        />
      );
    }

    return (
      <input
        {...commonProps}
        className={invalidAmount ? 'invalid-amount' : ''}
        type="text"
        value={amount ? formatAmount(value) : value}
        title={value}
        onChange={(event) => updateCell(rowIndex, colIndex, amount ? toHalfWidthNumber(event.target.value) : event.target.value)}
        onBlur={(event) => updateCellOnBlur(rowIndex, colIndex, event.target.value)}
        lang={isJapaneseTextColumn(column) ? 'ja' : undefined}
        inputMode={cellInputMode(column)}
        pattern={amount ? '[0-9,\\-]*' : undefined}
        autoComplete="off"
        spellCheck={false}
        placeholder={placeholderFor(column)}
        aria-label={`${column} ${rowIndex + 1}行目`}
      />
    );
  };

  return (
    <section className="panel answer-panel">
      <div className="section-heading answer-heading">
        <h2>答案入力テーブル</h2>
        <div className="button-row answer-actions">
          <button onClick={addRows}>10行追加</button>
          <button onClick={addColumn}>列追加</button>
          <button className="secondary" onClick={compactRows}>空行整理</button>
          <button className="accent" onClick={onApplyRowPoints}>行別得点合計を問題得点へ反映</button>
        </div>
      </div>
      <div className={`table-wrap answer-table-wrap ${answer.templateId === 'journal' ? 'journal-table-wrap' : ''}`}>
        <table className={`answer-table ${answer.templateId === 'journal' ? 'journal-table' : ''}`}>
          <colgroup>
            {answer.columns.map((column, index) => <col key={`${column}-${index}`} className={columnClass(column)} />)}
            <col className="col-grade" />
            <col className="col-points" />
          </colgroup>
          <thead>
            <tr>
              {answer.columns.map((column, index) => (
                <th key={`${column}-${index}`} className={columnClass(column)}>
                  <input value={column} onChange={(event) => updateHeader(index, event.target.value)} aria-label={`列名${index + 1}`} />
                </th>
              ))}
              <th className="col-grade">採点</th>
              <th className="col-points">行別得点</th>
            </tr>
          </thead>
          <tbody>
            {answer.rows.map((row, rowIndex) => (
              <tr key={row.id}>
                {answer.columns.map((column, colIndex) => (
                  <td key={`${row.id}-${colIndex}`} className={columnClass(column)}>
                    {renderCellInput(row, rowIndex, column, colIndex)}
                  </td>
                ))}
                <td className="col-grade">
                  <select
                    data-answer-cell="true"
                    data-row-index={rowIndex}
                    data-col-index={answer.columns.length}
                    value={row.grade}
                    onChange={(event) => updateGrade(rowIndex, event.target.value as GradeMark)}
                    onKeyDown={(event) => handleTableCellKeyDown(event, rowIndex, answer.columns.length)}
                    aria-label={`採点 ${rowIndex + 1}行目`}
                  >
                    {gradeOptions.map((grade) => <option key={grade} value={grade}>{grade}</option>)}
                  </select>
                </td>
                <td className="col-points">
                  <input
                    data-answer-cell="true"
                    data-row-index={rowIndex}
                    data-col-index={answer.columns.length + 1}
                    type="text"
                    inputMode="numeric"
                    pattern="[0-9,\\-]*"
                    value={String(row.points)}
                    onChange={(event) => updatePoints(rowIndex, Number(normalizeNumericInput(event.target.value) || 0))}
                    onKeyDown={(event) => handleTableCellKeyDown(event, rowIndex, answer.columns.length + 1)}
                    autoComplete="off"
                    spellCheck={false}
                    aria-label={`行別得点 ${rowIndex + 1}行目`}
                  />
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  );
}
