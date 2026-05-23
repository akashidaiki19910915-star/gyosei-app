import { useEffect, useMemo, useState } from 'react';
import type { AnswerState, MissReason, MistakeCard, ProblemDefinition } from '../types';
import { formatDateTime, nowIso } from '../utils/dates';

interface Props {
  problem: ProblemDefinition;
  answer: AnswerState;
  cards: MistakeCard[];
  onSave: (card: MistakeCard) => Promise<void>;
  onDelete: (id: string) => Promise<void>;
  onExportCsv: () => void;
}

const missReasonOptions: MissReason[] = ['論点理解不足', '仕訳ミス', '借方貸方逆', '金額ミス', '集計ミス', '転記ミス', '表の入力位置ミス', '下書き不足', '時間不足', '解答欄形式の誤認', '問題文読み落とし', 'その他'];

function emptyCard(problemId: string): MistakeCard {
  const now = nowIso();
  return {
    id: crypto.randomUUID(),
    problemId,
    correctFlow: '',
    mistakeReason: '',
    firstReviewPoint: '',
    preSolveChecklist: '',
    preventionPhrase: '',
    missReasons: [],
    importance: '中',
    createdAt: now,
    updatedAt: now,
  };
}

export function MistakeCardPanel({ problem, answer, cards, onSave, onDelete, onExportCsv }: Props) {
  const problemCards = useMemo(() => cards.filter((card) => card.problemId === problem.id).sort((a, b) => b.updatedAt.localeCompare(a.updatedAt)), [cards, problem.id]);
  const latest = problemCards[0];
  const needsCard = answer.rank === 'C' || (answer.maxScore > 0 && answer.score / answer.maxScore < 0.7);
  const [draft, setDraft] = useState<MistakeCard>(() => latest ?? emptyCard(problem.id));

  useEffect(() => {
    setDraft(latest ?? emptyCard(problem.id));
  }, [latest, problem.id]);

  const update = (patch: Partial<MistakeCard>) => setDraft((current) => ({ ...current, ...patch, updatedAt: nowIso() }));
  const toggleReason = (reason: MissReason) => {
    const exists = draft.missReasons.includes(reason);
    update({ missReasons: exists ? draft.missReasons.filter((item) => item !== reason) : [...draft.missReasons, reason] });
  };

  const save = async () => {
    await onSave({ ...draft, problemId: problem.id, updatedAt: nowIso() });
  };

  const startNew = () => setDraft(emptyCard(problem.id));
  const editCard = (card: MistakeCard) => setDraft(card);

  return (
    <details className="management-panel recovery-card-panel" open>
      <summary>白紙再現・解き直しカード</summary>
      <div className="panel-body">
        <p className="notice-small">問題文・解答・解説をそのまま保存しないでください。ここには、自分の言葉で「処理手順」「ミス原因」「次回の注意点」だけを書いてください。</p>
        {needsCard && problemCards.length === 0 && <p className="warning-text">C判定または70%未満です。次回同じミスを防ぐため、白紙再現・解き直しカードを作成してください。</p>}
        {latest && (
          <div className="priority-card">
            <span className="status-badge">最新カード</span>
            <h3>{latest.preventionPhrase || '同じミスを防ぐ一言は未入力'}</h3>
            <p><strong>次回最初に見る注意点：</strong>{latest.firstReviewPoint || '未入力'}</p>
            <p><strong>更新：</strong>{formatDateTime(latest.updatedAt)}</p>
          </div>
        )}
        <div className="form-grid two-columns">
          <label>重要度
            <select value={draft.importance} onChange={(event) => update({ importance: event.target.value as MistakeCard['importance'] })}>
              <option value="高">高</option><option value="中">中</option><option value="低">低</option>
            </select>
          </label>
          <label>同じミスを防ぐ一言
            <input value={draft.preventionPhrase} onChange={(event) => update({ preventionPhrase: event.target.value })} placeholder="例：金額を入れる前に、処理時点を確認する。" />
          </label>
        </div>
        <label>正しい処理の流れ
          <textarea value={draft.correctFlow} onChange={(event) => update({ correctFlow: event.target.value })} placeholder="例：売上原価を計算してから、期末商品の処理、貸倒引当金、減価償却の順に確認する。" />
        </label>
        <label>なぜ間違えたか
          <textarea value={draft.mistakeReason} onChange={(event) => update({ mistakeReason: event.target.value })} placeholder="例：問題文の「期末に一括処理」という条件を読み落とした。" />
        </label>
        <label>次回最初に見る注意点
          <textarea value={draft.firstReviewPoint} onChange={(event) => update({ firstReviewPoint: event.target.value })} placeholder="例：処理条件、日付、期末整理か期中処理かを先に見る。" />
        </label>
        <label>次回は何を確認してから解くか
          <textarea value={draft.preSolveChecklist} onChange={(event) => update({ preSolveChecklist: event.target.value })} placeholder="例：設問の時点、処理対象、資料の使い分けを最初に確認する。" />
        </label>
        <div>
          <strong>関連するミス原因</strong>
          <div className="check-list wide">
            {missReasonOptions.map((reason) => <label key={reason}><input type="checkbox" checked={draft.missReasons.includes(reason)} onChange={() => toggleReason(reason)} />{reason}</label>)}
          </div>
        </div>
        <div className="button-row">
          <button className="accent" onClick={save}>カード保存</button>
          <button className="secondary" onClick={startNew}>新規カード</button>
          <button onClick={onExportCsv}>解き直しカードCSV</button>
        </div>
        <div className="table-wrap short-wrap">
          <table className="compact-table">
            <thead><tr><th>重要度</th><th>更新日</th><th>一言</th><th>注意点</th><th>操作</th></tr></thead>
            <tbody>
              {problemCards.map((card) => (
                <tr key={card.id}>
                  <td>{card.importance}</td><td>{formatDateTime(card.updatedAt)}</td><td>{card.preventionPhrase}</td><td>{card.firstReviewPoint}</td>
                  <td><button onClick={() => editCard(card)}>編集</button> <button className="danger" onClick={() => onDelete(card.id)}>削除</button></td>
                </tr>
              ))}
              {problemCards.length === 0 && <tr><td colSpan={5}>この問題IDの解き直しカードは未作成です。</td></tr>}
            </tbody>
          </table>
        </div>
      </div>
    </details>
  );
}
