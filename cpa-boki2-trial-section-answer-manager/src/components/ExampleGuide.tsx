import type { TemplateDefinition } from '../types';

interface Props {
  template: TemplateDefinition;
}

export function ExampleGuide({ template }: Props) {
  const guide = template.exampleGuide;

  return (
    <details className="example-guide">
      <summary>記入例・使い方</summary>
      <div className="example-guide-body">
        <p className="example-warning">この記入例は架空の入力例です。教材の解答ではありません。実際の問題では、紙教材・PDF・CPA画面の解答解説を見ながら自己採点してください。</p>

        <section>
          <h3>このテンプレートで入力する内容</h3>
          <p>{guide.description}</p>
        </section>

        <section>
          <h3>使い方</h3>
          <ul>
            {guide.howToUse.map((item) => <li key={item}>{item}</li>)}
          </ul>
        </section>

        <section>
          <h3>架空の記入例</h3>
          <div className="example-table-wrap">
            <table className="example-table">
              <thead>
                <tr>
                  {Array.from(new Set(guide.exampleRows.flatMap((row) => Object.keys(row)))).map((key) => <th key={key}>{key}</th>)}
                </tr>
              </thead>
              <tbody>
                {guide.exampleRows.map((row, rowIndex) => {
                  const keys = Array.from(new Set(guide.exampleRows.flatMap((item) => Object.keys(item))));
                  return (
                    <tr key={`${template.id}-example-${rowIndex}`}>
                      {keys.map((key) => <td key={key}>{row[key] || ''}</td>)}
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </section>

        <section>
          <h3>採点時の入力例</h3>
          <p>{guide.scoringExample}</p>
        </section>

        <section>
          <h3>注意点</h3>
          <ul>
            {guide.notes.map((note) => <li key={note}>{note}</li>)}
          </ul>
        </section>
      </div>
    </details>
  );
}
