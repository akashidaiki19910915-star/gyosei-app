import { useEffect, useMemo, useState } from 'react';
import { originalQuestions } from './data/originalQuestions';
import type { AnswerAttempt, AnswerLine, AnswerTemplateId, MissReason, QuestionItem, ReviewRank, ReviewSchedule, SourceType, StudyBackupPayload } from './originalStudyTypes';
import { exportStudyData, getAttempts, getReviewSchedules, getStoredQuestions, importStudyData, saveAttempt, saveQuestion, saveReviewSchedule, storageSummary } from './originalStudyStorage';
import './styles.css';

type AppMode = 'solve' | 'grading' | 'review' | 'management';

type QuestionForm = Omit<QuestionItem, 'conditions' | 'modelAnswer'> & {
  conditionsText: string;
  modelAnswerJson: string;
};

const missReasons: MissReason[] = ['論点理解不足', '仕訳ミス', '借方貸方逆', '金額ミス', '集計ミス', '転記ミス', '表の入力位置ミス', '下書き不足', '時間不足', '解答形式の誤認', '問題文読み落とし', 'その他'];
const templateLabels: Record<AnswerTemplateId, string> = { journal: '仕訳テンプレート', numeric: '数値入力テンプレート', statementTable: '表形式テンプレート', free: '自由テンプレート' };
const sourceTypes: SourceType[] = ['original', 'public_domain', 'licensed', 'user_created'];

function todayString(): string {
  return new Date().toISOString().slice(0, 10);
}

function addDays(days: number): string {
  const date = new Date();
  date.setDate(date.getDate() + days);
  return date.toISOString().slice(0, 10);
}

function newLine(itemName = ''): AnswerLine {
  return { id: crypto.randomUUID(), itemName, debitAccount: '', debitAmount: '', creditAccount: '', creditAmount: '', value: '', value1: '', value2: '', value3: '', value4: '', memo: '', grade: '未採点', points: 0 };
}

function emptyRows(question: QuestionItem): AnswerLine[] {
  const modelRows = question.modelAnswer?.length ? question.modelAnswer : [newLine('')];
  return modelRows.map((row) => ({ ...newLine(row.itemName ?? ''), id: crypto.randomUUID() }));
}

function normalizeNumberText(value: string): string {
  return value.replace(/[０-９]/g, (char) => String.fromCharCode(char.charCodeAt(0) - 0xFEE0)).replace(/,/g, '').replace(/[^0-9.-]/g, '');
}

function formatAmount(value = ''): string {
  const normalized = normalizeNumberText(value);
  if (!normalized || normalized === '-' || normalized === '.') return normalized;
  const numeric = Number(normalized);
  if (!Number.isFinite(numeric)) return normalized;
  return numeric.toLocaleString('ja-JP');
}

function rowPointsTotal(rows: AnswerLine[]): number {
  return rows.reduce((sum, row) => sum + (Number(row.points) || 0), 0);
}

function questionToForm(question?: QuestionItem): QuestionForm {
  const base: QuestionItem = question ?? {
    id: `ユーザー-${Date.now()}`,
    examType: 'boki2',
    subject: '商業簿記',
    section: '第1問対策',
    topic: '',
    title: '',
    questionText: '',
    conditions: [],
    answerTemplateId: 'journal',
    maxScore: 4,
    estimatedMinutes: 3,
    difficulty: '標準',
    explanation: '',
    sourceType: 'user_created',
    isVerified: false,
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    modelAnswer: [newLine('')],
    modelAnswerText: '',
  };
  return { ...base, conditionsText: base.conditions.join('\n'), modelAnswerJson: JSON.stringify(base.modelAnswer ?? [], null, 2) };
}

function formToQuestion(form: QuestionForm): QuestionItem {
  let parsed: AnswerLine[] = [];
  try {
    parsed = JSON.parse(form.modelAnswerJson || '[]') as AnswerLine[];
  } catch {
    parsed = [];
  }
  const sourceType = form.sourceType === 'original' ? 'user_created' : form.sourceType;
  return {
    id: form.id.trim(),
    examType: 'boki2',
    subject: form.subject,
    section: form.section,
    topic: form.topic.trim(),
    title: form.title.trim(),
    questionText: form.questionText.trim(),
    conditions: form.conditionsText.split('\n').map((line) => line.trim()).filter(Boolean),
    answerTemplateId: form.answerTemplateId,
    maxScore: Number(form.maxScore) || 0,
    estimatedMinutes: Number(form.estimatedMinutes) || undefined,
    difficulty: form.difficulty,
    explanation: form.explanation,
    sourceType,
    isVerified: form.isVerified,
    createdAt: form.createdAt || new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    modelAnswer: parsed.map((row) => ({ ...newLine(row.itemName ?? ''), ...row, id: row.id || crypto.randomUUID() })),
    modelAnswerText: form.modelAnswerText,
  };
}

function buildReviewSchedule(question: QuestionItem, rank: ReviewRank, score: number, previous?: ReviewSchedule): ReviewSchedule {
  const aStreak = rank === 'A' ? (previous?.aStreak ?? 0) + 1 : 0;
  const cStreak = rank === 'C' ? (previous?.cStreak ?? 0) + 1 : 0;
  const days = rank === 'A' ? (aStreak >= 3 ? 30 : aStreak >= 2 ? 14 : 7) : rank === 'B' ? 3 : 1;
  return {
    id: question.id,
    questionId: question.id,
    nextReviewDate: addDays(days),
    latestRank: rank,
    latestScore: score,
    maxScore: question.maxScore,
    aStreak,
    cStreak,
    attempts: (previous?.attempts ?? 0) + 1,
    isWeak: cStreak >= 2,
    updatedAt: new Date().toISOString(),
  };
}

function problemLabel(question: QuestionItem): string {
  return `${question.id}　${question.title || question.topic}`;
}

export default function App() {
  const [mode, setMode] = useState<AppMode>('solve');
  const [questions, setQuestions] = useState<QuestionItem[]>(originalQuestions);
  const [attempts, setAttempts] = useState<AnswerAttempt[]>([]);
  const [reviews, setReviews] = useState<ReviewSchedule[]>([]);
  const [selectedId, setSelectedId] = useState(originalQuestions[0].id);
  const [rows, setRows] = useState<AnswerLine[]>(() => emptyRows(originalQuestions[0]));
  const [score, setScore] = useState(0);
  const [rank, setRank] = useState<ReviewRank>('B');
  const [selectedReasons, setSelectedReasons] = useState<MissReason[]>([]);
  const [reviewMemo, setReviewMemo] = useState('');
  const [message, setMessage] = useState('待機中');
  const [form, setForm] = useState<QuestionForm>(() => questionToForm());
  const [summary, setSummary] = useState<{ questionCount: number; attemptCount: number; reviewCount: number; usage: number | null; quota: number | null } | null>(null);

  const selected = useMemo(() => questions.find((question) => question.id === selectedId) ?? questions[0], [questions, selectedId]);
  const latestReview = useMemo(() => reviews.find((review) => review.questionId === selected.id), [reviews, selected.id]);
  const dueCards = useMemo(() => {
    const today = todayString();
    return reviews
      .map((review) => ({ review, question: questions.find((question) => question.id === review.questionId) }))
      .filter((item): item is { review: ReviewSchedule; question: QuestionItem } => Boolean(item.question))
      .sort((a, b) => a.review.nextReviewDate.localeCompare(b.review.nextReviewDate))
      .filter((item) => item.review.nextReviewDate <= today || item.review.latestRank === 'C' || item.review.latestRank === 'B' || item.review.isWeak);
  }, [questions, reviews]);

  const untouchedQuestions = useMemo(() => questions.filter((question) => !attempts.some((attempt) => attempt.questionId === question.id)), [questions, attempts]);

  async function refreshAll() {
    const [stored, attemptRows, reviewRows, storage] = await Promise.all([getStoredQuestions(), getAttempts(), getReviewSchedules(), storageSummary()]);
    const merged = new Map<string, QuestionItem>();
    originalQuestions.forEach((question) => merged.set(question.id, question));
    stored.forEach((question) => merged.set(question.id, question));
    const nextQuestions = Array.from(merged.values()).sort((a, b) => a.id.localeCompare(b.id, 'ja'));
    setQuestions(nextQuestions);
    setAttempts(attemptRows);
    setReviews(reviewRows);
    setSummary(storage);
  }

  useEffect(() => { void refreshAll(); }, []);

  useEffect(() => {
    setRows(emptyRows(selected));
    setScore(0);
    setRank('B');
    setSelectedReasons([]);
    setReviewMemo('');
  }, [selected.id]);

  function updateRow(index: number, patch: Partial<AnswerLine>) {
    setRows((current) => current.map((row, rowIndex) => rowIndex === index ? { ...row, ...patch } : row));
  }

  function addRow() {
    setRows((current) => [...current, newLine('')]);
  }

  function applyRowPoints() {
    setScore(rowPointsTotal(rows));
  }

  function toggleReason(reason: MissReason) {
    setSelectedReasons((current) => current.includes(reason) ? current.filter((item) => item !== reason) : [...current, reason]);
  }

  async function saveCurrentAttempt() {
    const finalScore = score || rowPointsTotal(rows);
    const review = buildReviewSchedule(selected, rank, finalScore, latestReview);
    const attempt: AnswerAttempt = {
      id: crypto.randomUUID(),
      questionId: selected.id,
      questionTitle: selected.title,
      subject: selected.subject,
      section: selected.section,
      topic: selected.topic,
      answeredAt: new Date().toISOString(),
      answerTemplateId: selected.answerTemplateId,
      rows,
      score: finalScore,
      maxScore: selected.maxScore,
      rank,
      missReasons: selectedReasons,
      reviewMemo,
      nextReviewDate: review.nextReviewDate,
    };
    await saveAttempt(attempt);
    await saveReviewSchedule(review);
    await refreshAll();
    setMessage(`保存しました。次回復習日：${review.nextReviewDate}${review.isWeak ? '（連続Cの苦手問題）' : ''}`);
    const currentIndex = questions.findIndex((question) => question.id === selected.id);
    const next = questions[currentIndex + 1] ?? questions[0];
    setSelectedId(next.id);
    setMode('solve');
  }

  async function saveQuestionForm() {
    const question = formToQuestion(form);
    if (!question.id || !question.title) {
      setMessage('問題IDとタイトルは必須です');
      return;
    }
    await saveQuestion(question);
    await refreshAll();
    setSelectedId(question.id);
    setMessage('ユーザー作成問題を保存しました');
  }

  async function handleExport() {
    const payload = await exportStudyData(originalQuestions);
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'boki2-original-study-backup.json';
    a.click();
    URL.revokeObjectURL(url);
    setMessage('JSONエクスポートを作成しました');
  }

  async function handleImport(file?: File) {
    if (!file) return;
    try {
      const payload = JSON.parse(await file.text()) as StudyBackupPayload;
      await importStudyData(payload);
      await refreshAll();
      setMessage('JSONインポートが完了しました');
    } catch {
      setMessage('JSONインポートに失敗しました');
    }
  }

  function renderProblemCard() {
    return <section className="panel question-reference-card original-question-card">
      <p className="eyebrow">問題文</p>
      <h2>{problemLabel(selected)}</h2>
      <p className="readable-question-text">{selected.questionText || '問題文は未登録です。必要に応じて管理画面から問題文を追加してください。'}</p>
      {selected.conditions.length > 0 && <div className="condition-list"><strong>条件</strong>{selected.conditions.map((condition) => <span key={condition}>{condition}</span>)}</div>}
      <div className="answer-summary-grid compact-meta-grid">
        <div><span>科目</span><strong>{selected.subject}</strong></div><div><span>大問対策</span><strong>{selected.section}</strong></div><div><span>論点</span><strong>{selected.topic}</strong></div><div><span>満点</span><strong>{selected.maxScore}点</strong></div><div><span>形式</span><strong>{templateLabels[selected.answerTemplateId]}</strong></div><div><span>出典</span><strong>{selected.sourceType}</strong></div>
      </div>
    </section>;
  }

  function renderAnswerInput(readOnly = false) {
    const commonReview = (row: AnswerLine, index: number) => <>
      <label>メモ<textarea value={row.memo ?? ''} disabled={readOnly} onChange={(event) => updateRow(index, { memo: event.target.value })} /></label>
      <label>採点<select value={row.grade} disabled={readOnly} onChange={(event) => updateRow(index, { grade: event.target.value as AnswerLine['grade'] })}><option>未採点</option><option>○</option><option>△</option><option>×</option></select></label>
      <label>得点<input value={String(row.points || '')} disabled={readOnly} inputMode="numeric" onChange={(event) => updateRow(index, { points: Number(normalizeNumberText(event.target.value)) || 0 })} /></label>
    </>;

    if (selected.answerTemplateId === 'journal') {
      return <section className="answer-input-main original-answer-input"><div className="pc-answer-only"><table className="answer-table"><thead><tr><th>行</th><th>借方科目</th><th>借方金額</th><th>貸方科目</th><th>貸方金額</th><th>メモ</th><th>採点</th><th>得点</th></tr></thead><tbody>{rows.map((row, index) => <tr key={row.id}><td>{index + 1}</td><td><input value={row.debitAccount ?? ''} disabled={readOnly} onChange={(event) => updateRow(index, { debitAccount: event.target.value })} /></td><td><input value={formatAmount(row.debitAmount)} disabled={readOnly} inputMode="numeric" onChange={(event) => updateRow(index, { debitAmount: normalizeNumberText(event.target.value) })} /></td><td><input value={row.creditAccount ?? ''} disabled={readOnly} onChange={(event) => updateRow(index, { creditAccount: event.target.value })} /></td><td><input value={formatAmount(row.creditAmount)} disabled={readOnly} inputMode="numeric" onChange={(event) => updateRow(index, { creditAmount: normalizeNumberText(event.target.value) })} /></td><td><input value={row.memo ?? ''} disabled={readOnly} onChange={(event) => updateRow(index, { memo: event.target.value })} /></td><td><select value={row.grade} disabled={readOnly} onChange={(event) => updateRow(index, { grade: event.target.value as AnswerLine['grade'] })}><option>未採点</option><option>○</option><option>△</option><option>×</option></select></td><td><input value={String(row.points || '')} disabled={readOnly} onChange={(event) => updateRow(index, { points: Number(normalizeNumberText(event.target.value)) || 0 })} /></td></tr>)}</tbody></table></div><div className="mobile-journal-input">{rows.map((row, index) => <article className="journal-entry-card" key={row.id}><h3>仕訳{index + 1}</h3><div className="journal-entry-side"><h4>借方</h4><label>借方科目<input value={row.debitAccount ?? ''} disabled={readOnly} onChange={(event) => updateRow(index, { debitAccount: event.target.value })} /></label><label>借方金額<input value={formatAmount(row.debitAmount)} disabled={readOnly} inputMode="numeric" onChange={(event) => updateRow(index, { debitAmount: normalizeNumberText(event.target.value) })} /></label></div><div className="journal-entry-side"><h4>貸方</h4><label>貸方科目<input value={row.creditAccount ?? ''} disabled={readOnly} onChange={(event) => updateRow(index, { creditAccount: event.target.value })} /></label><label>貸方金額<input value={formatAmount(row.creditAmount)} disabled={readOnly} inputMode="numeric" onChange={(event) => updateRow(index, { creditAmount: normalizeNumberText(event.target.value) })} /></label></div>{commonReview(row, index)}</article>)}</div>{!readOnly && <button className="secondary" onClick={addRow}>＋ 行を追加</button>}</section>;
    }

    return <section className="answer-input-main original-answer-input"><div className="pc-answer-only"><table className="answer-table"><thead><tr><th>行</th><th>項目名</th><th>{selected.answerTemplateId === 'numeric' ? '入力欄' : '入力値1'}</th>{selected.answerTemplateId !== 'numeric' && <><th>入力値2</th><th>入力値3</th><th>入力値4</th></>}<th>メモ</th><th>採点</th><th>得点</th></tr></thead><tbody>{rows.map((row, index) => <tr key={row.id}><td>{index + 1}</td><td><input value={row.itemName ?? ''} disabled={readOnly} onChange={(event) => updateRow(index, { itemName: event.target.value })} /></td><td><input value={formatAmount(row.value ?? row.value1)} disabled={readOnly} inputMode="numeric" onChange={(event) => updateRow(index, selected.answerTemplateId === 'numeric' ? { value: normalizeNumberText(event.target.value) } : { value1: normalizeNumberText(event.target.value) })} /></td>{selected.answerTemplateId !== 'numeric' && <><td><input value={formatAmount(row.value2)} disabled={readOnly} onChange={(event) => updateRow(index, { value2: normalizeNumberText(event.target.value) })} /></td><td><input value={formatAmount(row.value3)} disabled={readOnly} onChange={(event) => updateRow(index, { value3: normalizeNumberText(event.target.value) })} /></td><td><input value={formatAmount(row.value4)} disabled={readOnly} onChange={(event) => updateRow(index, { value4: normalizeNumberText(event.target.value) })} /></td></>}<td><input value={row.memo ?? ''} disabled={readOnly} onChange={(event) => updateRow(index, { memo: event.target.value })} /></td><td><select value={row.grade} disabled={readOnly} onChange={(event) => updateRow(index, { grade: event.target.value as AnswerLine['grade'] })}><option>未採点</option><option>○</option><option>△</option><option>×</option></select></td><td><input value={String(row.points || '')} disabled={readOnly} onChange={(event) => updateRow(index, { points: Number(normalizeNumberText(event.target.value)) || 0 })} /></td></tr>)}</tbody></table></div><div className="mobile-journal-input">{rows.map((row, index) => <article className="journal-entry-card" key={row.id}><h3>解答{index + 1}</h3><label>項目名<input value={row.itemName ?? ''} disabled={readOnly} onChange={(event) => updateRow(index, { itemName: event.target.value })} /></label><label>入力値<input value={formatAmount(row.value ?? row.value1)} disabled={readOnly} inputMode="numeric" onChange={(event) => updateRow(index, selected.answerTemplateId === 'numeric' ? { value: normalizeNumberText(event.target.value) } : { value1: normalizeNumberText(event.target.value) })} /></label>{commonReview(row, index)}</article>)}</div>{!readOnly && <button className="secondary" onClick={addRow}>＋ 行を追加</button>}</section>;
  }

  function renderModelAnswer() {
    return <section className="panel model-answer-card"><p className="eyebrow">模範答案・解説</p><h2>自己採点用</h2>{selected.modelAnswer?.length ? <div className="model-answer-list">{selected.modelAnswer.map((row, index) => <div key={`${selected.id}-${index}`}><strong>{index + 1}</strong><span>{selected.answerTemplateId === 'journal' ? `${row.debitAccount || ''} ${formatAmount(row.debitAmount)} / ${row.creditAccount || ''} ${formatAmount(row.creditAmount)}` : `${row.itemName || ''}：${formatAmount(row.value ?? row.value1)}`}</span></div>)}</div> : <p>{selected.modelAnswerText || '模範答案は未登録です。'}</p>}<p className="readable-question-text">{selected.explanation || '解説は未登録です。'}</p></section>;
  }

  return <main className="app-shell simplified-shell original-study-app">
    <header className="app-header compact-header"><div><h1>日商簿記2級 オリジナル問題演習・復習管理アプリ</h1><p>PDF抽出に依存せず、オリジナル問題データから問題文・答案欄・自己採点・復習予定を同期します。</p></div><div className="practice-progress-box"><strong>{questions.length}</strong><span>内蔵・追加問題</span></div></header>
    <nav className="mode-nav"><button className={mode === 'solve' ? 'active' : ''} onClick={() => setMode('solve')}>今すぐ解く</button><button className={mode === 'grading' ? 'active' : ''} onClick={() => setMode('grading')}>採点する</button><button className={mode === 'review' ? 'active' : ''} onClick={() => setMode('review')}>復習する</button><button className={mode === 'management' ? 'active' : ''} onClick={() => setMode('management')}>管理</button></nav>
    <div className="status-bar">{message}</div>

    {mode === 'solve' && <section className="mode-section solve-mode-section"><section className="panel answer-summary-card"><label>問題ID<select value={selected.id} onChange={(event) => setSelectedId(event.target.value)}>{questions.map((question) => <option key={question.id} value={question.id}>{problemLabel(question)}</option>)}</select></label><div className="button-row"><button onClick={() => setRows(emptyRows(selected))}>一時保存代わりに入力を初期化</button><button className="accent" onClick={() => { applyRowPoints(); setMode('grading'); }}>採点へ進む</button></div></section>{renderProblemCard()}<section className="focused-answer-area answer-only-main"><div className="focused-answer-title"><div><p className="eyebrow">答案入力</p><h2>{templateLabels[selected.answerTemplateId]}</h2></div><button className="secondary" onClick={applyRowPoints}>行別得点合計を問題得点へ反映</button></div>{renderAnswerInput(false)}</section></section>}

    {mode === 'grading' && <section className="mode-section grading-mode-section">{renderProblemCard()}<section className="focused-answer-area"><div className="focused-answer-title"><div><p className="eyebrow">自分の答案</p><h2>{selected.title}</h2></div><button className="secondary" onClick={applyRowPoints}>行別得点合計を反映</button></div>{renderAnswerInput(false)}</section>{renderModelAnswer()}<section className="panel self-score-box"><h2>自己採点</h2><div className="score-grid"><label>問題得点<input type="number" value={score} onChange={(event) => setScore(Number(event.target.value || 0))} /></label><label>満点<input type="number" value={selected.maxScore} disabled /></label><label>A/B/C判定<select value={rank} onChange={(event) => setRank(event.target.value as ReviewRank)}><option value="A">A：自力で解けた・説明できる</option><option value="B">B：正解または惜しいが手順に不安</option><option value="C">C：不正解・解説なしでは再現できない</option></select></label><label>次回復習日<input value={buildReviewSchedule(selected, rank, score || rowPointsTotal(rows), latestReview).nextReviewDate} disabled /></label></div><h3>ミス原因</h3><div className="check-list compact-check-list">{missReasons.map((reason) => <label key={reason}><input type="checkbox" checked={selectedReasons.includes(reason)} onChange={() => toggleReason(reason)} />{reason}</label>)}</div><label>復習メモ<textarea value={reviewMemo} onChange={(event) => setReviewMemo(event.target.value)} placeholder="次回解く前に見る注意点" /></label><button className="accent" onClick={() => void saveCurrentAttempt()}>保存して次へ</button></section></section>}

    {mode === 'review' && <section className="mode-section review-mode-section"><section className="panel progress-overview-panel"><div><p className="eyebrow">復習する</p><h2>今日やる問題をカードで確認</h2></div><div className="progress-summary-grid"><div><strong>{dueCards.filter((item) => item.review.nextReviewDate <= todayString()).length}</strong><span>今日やる問題</span></div><div><strong>{dueCards.filter((item) => item.review.nextReviewDate < todayString()).length}</strong><span>期限超過</span></div><div><strong>{reviews.filter((item) => item.latestRank === 'C').length}</strong><span>C判定</span></div><div><strong>{reviews.filter((item) => item.latestRank === 'B').length}</strong><span>B判定</span></div><div><strong>{untouchedQuestions.length}</strong><span>未着手</span></div><div><strong>{reviews.filter((item) => item.isWeak).length}</strong><span>連続ミス</span></div></div></section>{dueCards.length === 0 && <section className="panel"><h2>今日の復習はありません</h2><button className="accent" onClick={() => { setSelectedId(untouchedQuestions[0]?.id ?? questions[0].id); setMode('solve'); }}>未着手問題を解く</button></section>}{dueCards.map(({ question, review }) => <article className="panel priority-card" key={question.id}><p className="eyebrow">{review.isWeak ? '連続ミス' : review.nextReviewDate <= todayString() ? '今日の最優先' : '復習候補'}</p><h3>{question.id}　{question.title}</h3><p>前回：{review.latestRank || '未判定'}判定 / {review.latestScore}点 / 復習日：{review.nextReviewDate}</p><button onClick={() => { setSelectedId(question.id); setMode('solve'); }}>解く</button></article>)}</section>}

    {mode === 'management' && <section className="mode-section progress-mode-section"><section className="panel"><p className="eyebrow">管理</p><h2>問題追加・JSON・保存状態</h2><p>通常学習では開かなくてよい画面です。PDF補助取込は今回の主機能ではありません。</p></section><details className="panel management-panel" open><summary>問題一覧</summary><div className="management-list">{questions.map((question) => <button key={question.id} onClick={() => { setSelectedId(question.id); setForm(questionToForm(question)); }}>{question.id}｜{question.title}｜{question.sourceType}</button>)}</div></details><details className="panel management-panel" open><summary>問題追加・編集</summary><div className="question-form-grid"><label>問題ID<input value={form.id} onChange={(event) => setForm({ ...form, id: event.target.value })} /></label><label>科目<select value={form.subject} onChange={(event) => setForm({ ...form, subject: event.target.value as QuestionItem['subject'] })}><option>商業簿記</option><option>工業簿記</option></select></label><label>大問対策<select value={form.section} onChange={(event) => setForm({ ...form, section: event.target.value as QuestionItem['section'] })}><option>第1問対策</option><option>第2問対策</option><option>第3問対策</option><option>第4問対策</option><option>第5問対策</option></select></label><label>論点<input value={form.topic} onChange={(event) => setForm({ ...form, topic: event.target.value })} /></label><label>タイトル<input value={form.title} onChange={(event) => setForm({ ...form, title: event.target.value })} /></label><label>答案テンプレート<select value={form.answerTemplateId} onChange={(event) => setForm({ ...form, answerTemplateId: event.target.value as AnswerTemplateId })}><option value="journal">仕訳</option><option value="numeric">数値入力</option><option value="statementTable">表形式</option><option value="free">自由</option></select></label><label>満点<input type="number" value={form.maxScore} onChange={(event) => setForm({ ...form, maxScore: Number(event.target.value) })} /></label><label>想定時間<input type="number" value={form.estimatedMinutes ?? ''} onChange={(event) => setForm({ ...form, estimatedMinutes: Number(event.target.value) })} /></label><label>難易度<select value={form.difficulty} onChange={(event) => setForm({ ...form, difficulty: event.target.value as QuestionItem['difficulty'] })}><option>易しい</option><option>標準</option><option>やや難</option><option>難しい</option></select></label><label>sourceType<select value={form.sourceType} onChange={(event) => setForm({ ...form, sourceType: event.target.value as SourceType })}>{sourceTypes.map((type) => <option key={type}>{type}</option>)}</select></label><label>確認済み<select value={String(form.isVerified)} onChange={(event) => setForm({ ...form, isVerified: event.target.value === 'true' })}><option value="true">true</option><option value="false">false</option></select></label><label className="wide-field">問題文<textarea value={form.questionText} onChange={(event) => setForm({ ...form, questionText: event.target.value })} /></label><label className="wide-field">条件<textarea value={form.conditionsText} onChange={(event) => setForm({ ...form, conditionsText: event.target.value })} /></label><label className="wide-field">解説<textarea value={form.explanation ?? ''} onChange={(event) => setForm({ ...form, explanation: event.target.value })} /></label><label className="wide-field">模範答案JSON<textarea value={form.modelAnswerJson} onChange={(event) => setForm({ ...form, modelAnswerJson: event.target.value })} /></label><label className="wide-field">模範答案メモ<textarea value={form.modelAnswerText ?? ''} onChange={(event) => setForm({ ...form, modelAnswerText: event.target.value })} /></label></div><button className="accent" onClick={() => void saveQuestionForm()}>ユーザー作成問題として保存</button><button className="secondary" onClick={() => setForm(questionToForm())}>新規問題フォーム</button></details><details className="panel management-panel"><summary>JSONインポート・エクスポート / CSV出力</summary><div className="button-row"><button onClick={() => void handleExport()}>JSONエクスポート</button><label className="import-label">JSONインポート<input type="file" accept="application/json" onChange={(event) => void handleImport(event.target.files?.[0])} /></label><button onClick={() => { const csv = attempts.map((a) => [a.answeredAt, a.questionId, a.questionTitle, a.score, a.maxScore, a.rank, a.nextReviewDate].join(',')).join('\n'); const blob = new Blob([`日時,問題ID,タイトル,得点,満点,判定,次回復習日\n${csv}`], { type: 'text/csv;charset=utf-8' }); const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = 'boki2-attempts.csv'; a.click(); URL.revokeObjectURL(url); }}>CSV出力</button></div></details><details className="panel management-panel"><summary>履歴一覧</summary><div className="history-list">{attempts.map((attempt) => <div key={attempt.id}><strong>{attempt.questionId}</strong><span>{attempt.score}/{attempt.maxScore}点・{attempt.rank}・{attempt.nextReviewDate}</span></div>)}</div></details><details className="panel management-panel"><summary>PDF補助取込</summary><p className="warning-text">PDFの種類によっては正確に抽出できません。学習画面では確認済みの問題データのみ使用してください。今回の主機能ではないため、PDF抽出精度改善・OCR・完全自動問題生成は実装していません。</p></details><details className="panel management-panel"><summary>保存状態確認</summary><p>IndexedDBに questions / answerAttempts / reviewSchedules / settings / userProblems / histories を保存します。</p><p>問題数：{summary?.questionCount ?? '-'} / 履歴：{summary?.attemptCount ?? '-'} / 復習予定：{summary?.reviewCount ?? '-'}</p><button onClick={() => void refreshAll()}>保存状態を再確認</button></details></section>}
  </main>;
}
