import type { AnswerRow, AnswerState, GradeMark, TemplateDefinition } from '../types';

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

export function AnswerTable({ answer, template, onChange, onApplyRowPoints }: Props) {
  const updateHeader = (index: number, value: string) => {
    const columns = [...answer.columns];
    columns[index] = value;
    onChange({ ...answer, columns, rows: answer.rows.map((row) => ({ ...row, cells: row.cells.slice(0, columns.length) })) });
  };

  const updateCell = (rowIndex: number, colIndex: number, value: string) => {
    const rows = answer.rows.map((row, index) => {
      if (index !== rowIndex) return row;
      const cells = [...row.cells];
      cells[colIndex] = value;
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

  return (
    <section className="panel answer-panel">
      <div className="section-heading">
        <h2>答案入力テーブル</h2>
        <div className="button-row">
          <button onClick={addRows}>10行追加</button>
          <button onClick={addColumn}>列追加</button>
          <button className="secondary" onClick={compactRows}>空行整理</button>
          <button className="accent" onClick={onApplyRowPoints}>行別得点合計を問題得点へ反映</button>
        </div>
      </div>
      <div className="table-wrap">
        <table className="answer-table">
          <thead>
            <tr>
              {answer.columns.map((column, index) => (
                <th key={`${column}-${index}`}>
                  <input value={column} onChange={(event) => updateHeader(index, event.target.value)} aria-label={`列名${index + 1}`} />
                </th>
              ))}
              <th>採点</th>
              <th>行別得点</th>
            </tr>
          </thead>
          <tbody>
            {answer.rows.map((row, rowIndex) => (
              <tr key={row.id}>
                {answer.columns.map((column, colIndex) => (
                  <td key={`${row.id}-${colIndex}`}>
                    {template.optionColumns?.[column] ? (
                      <select value={row.cells[colIndex] ?? ''} onChange={(event) => updateCell(rowIndex, colIndex, event.target.value)}>
                        <option value="">選択</option>
                        {template.optionColumns[column].map((option) => <option key={option} value={option}>{option}</option>)}
                      </select>
                    ) : (
                      <input value={row.cells[colIndex] ?? ''} onChange={(event) => updateCell(rowIndex, colIndex, event.target.value)} />
                    )}
                  </td>
                ))}
                <td>
                  <select value={row.grade} onChange={(event) => updateGrade(rowIndex, event.target.value as GradeMark)}>
                    {gradeOptions.map((grade) => <option key={grade} value={grade}>{grade}</option>)}
                  </select>
                </td>
                <td><input type="number" value={row.points} onChange={(event) => updatePoints(rowIndex, Number(event.target.value || 0))} /></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  );
}
