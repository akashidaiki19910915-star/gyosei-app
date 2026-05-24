import { useEffect, useMemo, useState } from 'react';
import { externalProblemRefs } from './data/cpaTrialSections';
import { approvedOriginalQuestions, draftOriginalQuestions, originalQuestions } from './data/originalQuestions';
import type { AnswerAttempt, AnswerLine, AnswerTemplateId, ExternalProblemRef, MissReason, PracticeItem, QuestionItem, ReviewRank, ReviewSchedule, StudyBackupPayload, StudyMode } from './originalStudyTypes';
import { exportStudyData, getAttempts, getReviewSchedules, getStoredExternalProblemRefs, getStoredQuestions, importStudyData, saveAttempt, saveExternalProblemRef, saveQuestion, saveReviewSchedule, storageSummary } from './originalStudyStorage';
import './styles.css';

type AppMode = 'solve' | 'grading' | 'review' | 'management';
type ActiveStudyChoice = 'external_material' | 'built_in_question';

const missReasons: MissReason[] = ['論点理解不足', '仕訳ミス', '借方貸方逆', '金額ミス', '集計ミス', '転記ミス', '表の入力位置ミス', '下書き不足', '時間不足', '解答形式の誤認', '問題文読み落とし', 'その他'];
const templateLabels: Record<AnswerTemplateId, string> = { journal: '仕訳テンプレート', numeric: '数値入力テンプレート', statementTable: '表形式テンプレート', accountLedger: '勘定記入テンプレート', free: '自由テンプレート' };
const COPYRIGHT_NOTE = 'CPA問題集の試験対策編の区分のみ参考。問題文・解答・解説・数値・表構成は未使用。';

function todayString(): string { return new Date().toISOString().slice(0, 10); }
function addDays(days: number): string { const date = new Date(); date.setDate(date.getDate() + days); return date.toISOString().slice(0, 10); }
function normalizeNumberText(value: string): string { return value.replace(/[０-９]/g, (char) => String.fromCharCode(char.charCodeAt(0) - 0xFEE0)).replace(/,/g, '').replace(/[^0-9.-]/g, ''); }
function formatAmount(value = ''): string { const normalized = normalizeNumberText(value); if (!normalized || normalized === '-' || normalized === '.') return normalized; const numeric = Number(normalized); return Number.isFinite(numeric) ? numeric.toLocaleString('ja-JP') : normalized; }
function isBuiltIn(item: PracticeItem): item is QuestionItem { return item.studyMode === 'built_in_question'; }
function newLine(itemName = ''): AnswerLine { return { id: crypto.randomUUID(), itemName, debitDate: '', debitSummary: '', debitAccount: '', debitAmount: '', creditDate: '', creditSummary: '', creditAccount: '', creditAmount: '', value: '', value1: '', value2: '', value3: '', value4: '', memo: '', grade: '未採点', points: 0 }; }
function lineCountForTemplate(template: AnswerTemplateId): number { return template === 'journal' ? 4 : template === 'accountLedger' ? 6 : template === 'statementTable' ? 8 : template === 'numeric' ? 4 : 6; }
function emptyRows(item: PracticeItem): AnswerLine[] { const sourceRows = isBuiltIn(item) && item.modelAnswer.length ? item.modelAnswer : Array.from({ length: lineCountForTemplate(item.answerTemplateId) }, () => newLine('')); return sourceRows.map((row) => ({ ...newLine(row.itemName ?? ''), id: crypto.randomUUID() })); }
function rowPointsTotal(rows: AnswerLine[]): number { return rows.reduce((sum, row) => sum + (Number(row.points) || 0), 0); }
function labelOf(item: PracticeItem): string { return `${item.id}　${item.topic || item.title}`; }
function itemKey(item: PracticeItem): string { return `${item.studyMode}:${item.id}`; }
function parseItemKey(key: string): { mode: StudyMode; id: string } { const [mode, ...rest] = key.split(':'); return { mode: mode as StudyMode, id: rest.join(':') }; }
function nextReviewDateFor(rank: ReviewRank, previous?: ReviewSchedule): string { const aStreak = rank === 'A' ? (previous?.aStreak ?? 0) + 1 : 0; const days = rank === 'A' ? (aStreak >= 3 ? 30 : aStreak >= 2 ? 14 : 7) : rank === 'B' ? 3 : 1; return addDays(days); }
function buildReviewSchedule(item: PracticeItem, rank: ReviewRank, score: number, previous?: ReviewSchedule, manualDate?: string): ReviewSchedule { const aStreak = rank === 'A' ? (previous?.aStreak ?? 0) + 1 : 0; const cStreak = rank === 'C' ? (previous?.cStreak ?? 0) + 1 : 0; return { id: itemKey(item), examType: 'boki2', studyMode: item.studyMode, questionId: item.id, cpaTrialSectionRef: item.cpaTrialSectionRef, nextReviewDate: manualDate || nextReviewDateFor(rank, previous), latestRank: rank, latestScore: score, maxScore: item.maxScore, aStreak, cStreak, attempts: (previous?.attempts ?? 0) + 1, isWeak: cStreak >= 2 || previous?.isWeak === true, updatedAt: new Date().toISOString() }; }

function modelAnswerText(item: PracticeItem): string {
  if (!isBuiltIn(item)) return '';
  if (!item.modelAnswer.length) return item.modelAnswerText ?? '';
  return item.modelAnswer.map((row, index) => {
    if (item.answerTemplateId === 'journal') return `${index + 1}. 借方 ${row.debitAccount || '-'} ${formatAmount(row.debitAmount)} / 貸方 ${row.creditAccount || '-'} ${formatAmount(row.creditAmount)}`;
    if (item.answerTemplateId === 'accountLedger') return `${index + 1}. ${row.itemName || row.debitSummary || row.creditSummary || '行'}　借方${row.debitDate || ''} ${row.debitSummary || ''} ${formatAmount(row.debitAmount)} / 貸方${row.creditDate || ''} ${row.creditSummary || ''} ${formatAmount(row.creditAmount)}`;
    return `${index + 1}. ${row.itemName || '項目'}：${formatAmount(row.value ?? row.value1)} ${row.value2 ? `/ ${formatAmount(row.value2)}` : ''} ${row.value3 ? `/ ${formatAmount(row.value3)}` : ''} ${row.value4 ? `/ ${formatAmount(row.value4)}` : ''}`;
  }).join('\n');
}

export default function App() {
  const [mode, setMode] = useState<AppMode>('solve');
  const [choice, setChoice] = useState<ActiveStudyChoice>('external_material');
  const [externalRefs, setExternalRefs] = useState<ExternalProblemRef[]>(externalProblemRefs);
  const [questions, setQuestions] = useState<QuestionItem[]>(originalQuestions);
  const [attempts, setAttempts] = useState<AnswerAttempt[]>([]);
  const [reviews, setReviews] = useState<ReviewSchedule[]>([]);
  const [selectedExternalId, setSelectedExternalId] = useState(externalProblemRefs[0].id);
  const [selectedQuestionId, setSelectedQuestionId] = useState(approvedOriginalQuestions[0]?.id ?? originalQuestions[0].id);
  const [rows, setRows] = useState<AnswerLine[]>(() => emptyRows(externalProblemRefs[0]));
  const [score, setScore] = useState(0);
  const [rank, setRank] = useState<ReviewRank>('B');
  const [manualReviewDate, setManualReviewDate] = useState('');
  const [selectedReasons, setSelectedReasons] = useState<MissReason[]>([]);
  const [reviewMemo, setReviewMemo] = useState('');
  const [message, setMessage] = useState('待機中');
  const [questionJson, setQuestionJson] = useState('');
  const [externalJson, setExternalJson] = useState('');
  const [summary, setSummary] = useState<{ questionCount: number; externalRefCount: number; attemptCount: number; reviewCount: number; usage: number | null; quota: number | null } | null>(null);

  const approvedQuestions = useMemo(() => questions.filter((question) => question.verificationStatus === 'approved'), [questions]);
  const selectedItem: PracticeItem = useMemo(() => {
    if (choice === 'external_material') return externalRefs.find((item) => item.id === selectedExternalId) ?? externalRefs[0];
    return approvedQuestions.find((item) => item.id === selectedQuestionId) ?? approvedQuestions[0] ?? questions[0];
  }, [approvedQuestions, choice, externalRefs, questions, selectedExternalId, selectedQuestionId]);
  const latestReview = useMemo(() => reviews.find((review) => review.id === itemKey(selectedItem)), [reviews, selectedItem]);
  const dueCards = useMemo(() => {
    const today = todayString();
    const allItems: PracticeItem[] = [...externalRefs, ...approvedQuestions];
    return reviews.map((review) => ({ review, item: allItems.find((candidate) => itemKey(candidate) === review.id) })).filter((entry): entry is { review: ReviewSchedule; item: PracticeItem } => Boolean(entry.item)).sort((a, b) => {
      const aOver = a.review.nextReviewDate < today ? 0 : 1;
      const bOver = b.review.nextReviewDate < today ? 0 : 1;
      if (aOver !== bOver) return aOver - bOver;
      if (a.review.isWeak !== b.review.isWeak) return a.review.isWeak ? -1 : 1;
      if (a.review.latestRank !== b.review.latestRank) return a.review.latestRank === 'C' ? -1 : b.review.latestRank === 'C' ? 1 : 0;
      return a.review.nextReviewDate.localeCompare(b.review.nextReviewDate);
    });
  }, [approvedQuestions, externalRefs, reviews]);
  const untouchedItems = useMemo(() => [...externalRefs, ...approvedQuestions].filter((item) => !attempts.some((attempt) => attempt.studyMode === item.studyMode && attempt.questionId === item.id)), [approvedQuestions, attempts, externalRefs]);

  async function refreshAll() {
    const [storedQuestions, storedRefs, attemptRows, reviewRows, storage] = await Promise.all([getStoredQuestions(), getStoredExternalProblemRefs(), getAttempts(), getReviewSchedules(), storageSummary()]);
    const mergedQuestions = new Map<string, QuestionItem>();
    originalQuestions.forEach((question) => mergedQuestions.set(question.id, question));
    storedQuestions.forEach((question) => mergedQuestions.set(question.id, question));
    const mergedRefs = new Map<string, ExternalProblemRef>();
    externalProblemRefs.forEach((ref) => mergedRefs.set(ref.id, ref));
    storedRefs.forEach((ref) => mergedRefs.set(ref.id, ref));
    setQuestions(Array.from(mergedQuestions.values()).sort((a, b) => a.id.localeCompare(b.id, 'ja')));
    setExternalRefs(Array.from(mergedRefs.values()).sort((a, b) => a.id.localeCompare(b.id, 'ja')));
    setAttempts(attemptRows);
    setReviews(reviewRows);
    setSummary(storage);
  }

  useEffect(() => { void refreshAll(); }, []);
  useEffect(() => { setRows(emptyRows(selectedItem)); setScore(0); setRank('B'); setManualReviewDate(''); setSelectedReasons([]); setReviewMemo(''); }, [selectedItem.id, selectedItem.studyMode]);

  function updateRow(index: number, patch: Partial<AnswerLine>) { setRows((current) => current.map((row, rowIndex) => rowIndex === index ? { ...row, ...patch } : row)); }
  function addRow() { setRows((current) => [...current, newLine('')]); }
  function applyRowPoints() { setScore(rowPointsTotal(rows)); }
  function toggleReason(reason: MissReason) { setSelectedReasons((current) => current.includes(reason) ? current.filter((item) => item !== reason) : [...current, reason]); }
  function selectPracticeItem(item: PracticeItem) { setChoice(item.studyMode); if (item.studyMode === 'external_material') setSelectedExternalId(item.id); else setSelectedQuestionId(item.id); setMode('solve'); }

  async function saveCurrentAttempt() {
    const finalScore = score || rowPointsTotal(rows);
    const review = buildReviewSchedule(selectedItem, rank, finalScore, latestReview, manualReviewDate);
    const attempt: AnswerAttempt = { id: crypto.randomUUID(), examType: 'boki2', studyMode: selectedItem.studyMode, questionId: selectedItem.id, questionTitle: selectedItem.title, subject: selectedItem.subject, section: selectedItem.section, topic: selectedItem.topic, cpaTrialSectionRef: selectedItem.cpaTrialSectionRef, answeredAt: new Date().toISOString(), answerTemplateId: selectedItem.answerTemplateId, rows, score: finalScore, maxScore: selectedItem.maxScore, rank, missReasons: selectedReasons, reviewMemo, nextReviewDate: review.nextReviewDate };
    await saveAttempt(attempt);
    await saveReviewSchedule(review);
    await refreshAll();
    setMessage(`保存しました。次回復習日：${review.nextReviewDate}${review.isWeak ? '（苦手固定）' : ''}`);
    const list = selectedItem.studyMode === 'external_material' ? externalRefs : approvedQuestions;
    const currentIndex = list.findIndex((item) => item.id === selectedItem.id);
    const next = list[currentIndex + 1] ?? list[0];
    selectPracticeItem(next);
  }

  async function saveQuestionJson() {
    try {
      const parsed = JSON.parse(questionJson) as QuestionItem;
      if (parsed.studyMode !== 'built_in_question') throw new Error('studyMode');
      const question: QuestionItem = { ...parsed, sourceType: parsed.sourceType === 'initial_original' ? 'user_created' : parsed.sourceType, updatedAt: new Date().toISOString() };
      await saveQuestion(question);
      await refreshAll();
      setMessage('オリジナル問題を保存しました');
    } catch {
      setMessage('オリジナル問題JSONの保存に失敗しました');
    }
  }

  async function saveExternalJson() {
    try {
      const parsed = JSON.parse(externalJson) as ExternalProblemRef;
      if (parsed.studyMode !== 'external_material') throw new Error('studyMode');
      await saveExternalProblemRef({ ...parsed, updatedAt: new Date().toISOString() });
      await refreshAll();
      setMessage('外部教材モード用問題IDを保存しました');
    } catch {
      setMessage('外部教材モード用JSONの保存に失敗しました');
    }
  }

  async function handleExport() {
    const payload = await exportStudyData(originalQuestions, externalProblemRefs);
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'cross-qualification-study-backup.json';
    a.click();
    URL.revokeObjectURL(url);
    setMessage('JSONバックアップを作成しました');
  }

  async function handleImport(file?: File) {
    if (!file) return;
    try { await importStudyData(JSON.parse(await file.text()) as StudyBackupPayload); await refreshAll(); setMessage('JSON復元が完了しました'); } catch { setMessage('JSON復元に失敗しました'); }
  }

  function renderStudyChoice() {
    return <section className="study-choice-grid"><button className={choice === 'external_material' ? 'study-choice active' : 'study-choice'} onClick={() => setChoice('external_material')}><strong>CPA教材を見ながら解く</strong><span>手元の問題集・PDFを見ながら答案を入力し、自己採点と復習管理を行います。</span></button><button className={choice === 'built_in_question' ? 'study-choice active' : 'study-choice'} onClick={() => setChoice('built_in_question')}><strong>教材なしで解く</strong><span>検証済みのオリジナル短問を解き、解答解説を見て復習できます。</span></button></section>;
  }

  function renderProblemSelector() {
    const options = choice === 'external_material' ? externalRefs : approvedQuestions;
    const value = choice === 'external_material' ? selectedExternalId : selectedQuestionId;
    return <section className="panel answer-summary-card"><label>{choice === 'external_material' ? '外部教材の問題ID' : 'approved オリジナル短問'}<select value={value} onChange={(event) => choice === 'external_material' ? setSelectedExternalId(event.target.value) : setSelectedQuestionId(event.target.value)}>{options.map((item) => <option key={item.id} value={item.id}>{labelOf(item)}</option>)}</select></label><div className="button-row"><button onClick={() => setRows(emptyRows(selectedItem))}>一時保存</button><button className="accent" onClick={() => { applyRowPoints(); setMode('grading'); }}>採点へ進む</button></div></section>;
  }

  function renderQuestionCard() {
    return <section className="panel question-reference-card original-question-card"><p className="eyebrow">{choice === 'external_material' ? '外部教材モード' : '教材なしモード'}</p><h2>{labelOf(selectedItem)}</h2>{choice === 'external_material' ? <p className="readable-question-text">問題文・解答・解説はお手元のCPAラーニング問題集・PDF・紙教材で確認してください。この画面では、問題ID、答案入力、自己採点、復習管理だけを行います。</p> : <p className="readable-question-text">{isBuiltIn(selectedItem) ? selectedItem.questionText : ''}</p>}{isBuiltIn(selectedItem) && selectedItem.conditions.length > 0 && <div className="condition-list"><strong>条件</strong>{selectedItem.conditions.map((condition) => <span key={condition}>{condition}</span>)}</div>}<div className="answer-summary-grid compact-meta-grid"><div><span>資格</span><strong>日商簿記2級</strong></div><div><span>科目</span><strong>{selectedItem.subject}</strong></div><div><span>大問対策</span><strong>{selectedItem.section}</strong></div><div><span>論点</span><strong>{selectedItem.topic}</strong></div><div><span>区分Ref</span><strong>{selectedItem.cpaTrialSectionRef}</strong></div><div><span>形式</span><strong>{templateLabels[selectedItem.answerTemplateId]}</strong></div>{choice === 'external_material' && <div><span>ページメモ</span><strong>{(selectedItem as ExternalProblemRef).pageMemo || '未設定'}</strong></div>}{isBuiltIn(selectedItem) && <div><span>検証</span><strong>{selectedItem.verificationStatus}</strong></div>}</div></section>;
  }

  function renderAnswerInput(readOnly = false) {
    const reviewCells = (row: AnswerLine, index: number) => <><label>メモ<input value={row.memo ?? ''} disabled={readOnly} onChange={(e) => updateRow(index, { memo: e.target.value })} /></label><label>採点<select value={row.grade} disabled={readOnly} onChange={(e) => updateRow(index, { grade: e.target.value as AnswerLine['grade'] })}><option>未採点</option><option>○</option><option>△</option><option>×</option></select></label><label>得点<input value={String(row.points || '')} disabled={readOnly} inputMode="numeric" onChange={(e) => updateRow(index, { points: Number(normalizeNumberText(e.target.value)) || 0 })} /></label></>;
    if (selectedItem.answerTemplateId === 'journal') return <section className="answer-input-main original-answer-input"><div className="pc-answer-only"><table className="answer-table"><thead><tr><th>行</th><th>借方科目</th><th>借方金額</th><th>貸方科目</th><th>貸方金額</th><th>メモ</th><th>採点</th><th>得点</th></tr></thead><tbody>{rows.map((row, index) => <tr key={row.id}><td>{index + 1}</td><td><input value={row.debitAccount ?? ''} disabled={readOnly} onChange={(e) => updateRow(index, { debitAccount: e.target.value })} /></td><td><input value={formatAmount(row.debitAmount)} disabled={readOnly} inputMode="numeric" onChange={(e) => updateRow(index, { debitAmount: normalizeNumberText(e.target.value) })} /></td><td><input value={row.creditAccount ?? ''} disabled={readOnly} onChange={(e) => updateRow(index, { creditAccount: e.target.value })} /></td><td><input value={formatAmount(row.creditAmount)} disabled={readOnly} inputMode="numeric" onChange={(e) => updateRow(index, { creditAmount: normalizeNumberText(e.target.value) })} /></td><td><input value={row.memo ?? ''} disabled={readOnly} onChange={(e) => updateRow(index, { memo: e.target.value })} /></td><td><select value={row.grade} disabled={readOnly} onChange={(e) => updateRow(index, { grade: e.target.value as AnswerLine['grade'] })}><option>未採点</option><option>○</option><option>△</option><option>×</option></select></td><td><input value={String(row.points || '')} disabled={readOnly} onChange={(e) => updateRow(index, { points: Number(normalizeNumberText(e.target.value)) || 0 })} /></td></tr>)}</tbody></table></div><div className="mobile-journal-input">{rows.map((row, index) => <article className="journal-entry-card" key={row.id}><h3>仕訳{index + 1}</h3><div className="journal-entry-side"><h4>借方</h4><label>借方科目<input value={row.debitAccount ?? ''} disabled={readOnly} onChange={(e) => updateRow(index, { debitAccount: e.target.value })} /></label><label>借方金額<input value={formatAmount(row.debitAmount)} disabled={readOnly} inputMode="numeric" onChange={(e) => updateRow(index, { debitAmount: normalizeNumberText(e.target.value) })} /></label></div><div className="journal-entry-side"><h4>貸方</h4><label>貸方科目<input value={row.creditAccount ?? ''} disabled={readOnly} onChange={(e) => updateRow(index, { creditAccount: e.target.value })} /></label><label>貸方金額<input value={formatAmount(row.creditAmount)} disabled={readOnly} inputMode="numeric" onChange={(e) => updateRow(index, { creditAmount: normalizeNumberText(e.target.value) })} /></label></div>{reviewCells(row, index)}</article>)}</div>{!readOnly && <button className="secondary" onClick={addRow}>＋ 行を追加</button>}</section>;
    if (selectedItem.answerTemplateId === 'accountLedger') return <section className="answer-input-main original-answer-input"><table className="answer-table"><thead><tr><th>行</th><th>勘定科目名</th><th>借方日付</th><th>借方摘要</th><th>借方金額</th><th>貸方日付</th><th>貸方摘要</th><th>貸方金額</th><th>メモ</th><th>採点</th><th>得点</th></tr></thead><tbody>{rows.map((row, index) => <tr key={row.id}><td>{index + 1}</td><td><input value={row.itemName ?? ''} disabled={readOnly} onChange={(e) => updateRow(index, { itemName: e.target.value })} /></td><td><input value={row.debitDate ?? ''} disabled={readOnly} onChange={(e) => updateRow(index, { debitDate: e.target.value })} /></td><td><input value={row.debitSummary ?? ''} disabled={readOnly} onChange={(e) => updateRow(index, { debitSummary: e.target.value })} /></td><td><input value={formatAmount(row.debitAmount)} disabled={readOnly} onChange={(e) => updateRow(index, { debitAmount: normalizeNumberText(e.target.value) })} /></td><td><input value={row.creditDate ?? ''} disabled={readOnly} onChange={(e) => updateRow(index, { creditDate: e.target.value })} /></td><td><input value={row.creditSummary ?? ''} disabled={readOnly} onChange={(e) => updateRow(index, { creditSummary: e.target.value })} /></td><td><input value={formatAmount(row.creditAmount)} disabled={readOnly} onChange={(e) => updateRow(index, { creditAmount: normalizeNumberText(e.target.value) })} /></td><td><input value={row.memo ?? ''} disabled={readOnly} onChange={(e) => updateRow(index, { memo: e.target.value })} /></td><td><select value={row.grade} disabled={readOnly} onChange={(e) => updateRow(index, { grade: e.target.value as AnswerLine['grade'] })}><option>未採点</option><option>○</option><option>△</option><option>×</option></select></td><td><input value={String(row.points || '')} disabled={readOnly} onChange={(e) => updateRow(index, { points: Number(normalizeNumberText(e.target.value)) || 0 })} /></td></tr>)}</tbody></table>{!readOnly && <button className="secondary" onClick={addRow}>＋ 行を追加</button>}</section>;
    return <section className="answer-input-main original-answer-input"><table className="answer-table"><thead><tr><th>行</th><th>項目名</th><th>{selectedItem.answerTemplateId === 'numeric' ? '入力値' : '入力値1'}</th>{selectedItem.answerTemplateId !== 'numeric' && <><th>入力値2</th><th>入力値3</th><th>入力値4</th></>}<th>メモ</th><th>採点</th><th>得点</th></tr></thead><tbody>{rows.map((row, index) => <tr key={row.id}><td>{index + 1}</td><td><input value={row.itemName ?? ''} disabled={readOnly} onChange={(e) => updateRow(index, { itemName: e.target.value })} /></td><td><input value={formatAmount(row.value ?? row.value1)} disabled={readOnly} onChange={(e) => updateRow(index, selectedItem.answerTemplateId === 'numeric' ? { value: normalizeNumberText(e.target.value) } : { value1: normalizeNumberText(e.target.value) })} /></td>{selectedItem.answerTemplateId !== 'numeric' && <><td><input value={formatAmount(row.value2)} disabled={readOnly} onChange={(e) => updateRow(index, { value2: normalizeNumberText(e.target.value) })} /></td><td><input value={formatAmount(row.value3)} disabled={readOnly} onChange={(e) => updateRow(index, { value3: normalizeNumberText(e.target.value) })} /></td><td><input value={formatAmount(row.value4)} disabled={readOnly} onChange={(e) => updateRow(index, { value4: normalizeNumberText(e.target.value) })} /></td></>}<td><input value={row.memo ?? ''} disabled={readOnly} onChange={(e) => updateRow(index, { memo: e.target.value })} /></td><td><select value={row.grade} disabled={readOnly} onChange={(e) => updateRow(index, { grade: e.target.value as AnswerLine['grade'] })}><option>未採点</option><option>○</option><option>△</option><option>×</option></select></td><td><input value={String(row.points || '')} disabled={readOnly} onChange={(e) => updateRow(index, { points: Number(normalizeNumberText(e.target.value)) || 0 })} /></td></tr>)}</tbody></table>{!readOnly && <button className="secondary" onClick={addRow}>＋ 行を追加</button>}</section>;
  }

  function renderModelAnswer() {
    if (!isBuiltIn(selectedItem)) return null;
    return <section className="panel model-answer-card"><p className="eyebrow">模範解答・解説</p><pre>{modelAnswerText(selectedItem)}</pre><p className="readable-question-text">{selectedItem.explanation}</p></section>;
  }

  return <main className="app-shell simplified-shell original-study-app">
    <header className="app-header compact-header"><div><h1>資格横断型 問題演習・採点・復習管理アプリ</h1><p>簿記2級では「CPA教材を見ながら解く」と「教材なしで解く」の2モードを使います。PDF抽出・教材本文取り込みは行いません。</p></div><div className="practice-progress-box"><strong>{externalRefs.length}</strong><span>外部教材区分</span></div></header>
    <nav className="mode-nav"><button className={mode === 'solve' ? 'active' : ''} onClick={() => setMode('solve')}>今すぐ解く</button><button className={mode === 'grading' ? 'active' : ''} onClick={() => setMode('grading')}>採点する</button><button className={mode === 'review' ? 'active' : ''} onClick={() => setMode('review')}>復習する</button><button className={mode === 'management' ? 'active' : ''} onClick={() => setMode('management')}>管理</button></nav>
    <div className="status-bar">{message}</div>

    {mode === 'solve' && <section className="mode-section solve-mode-section">{renderStudyChoice()}{renderProblemSelector()}{renderQuestionCard()}<section className="focused-answer-area answer-only-main"><div className="focused-answer-title"><div><p className="eyebrow">答案入力</p><h2>{templateLabels[selectedItem.answerTemplateId]}</h2></div><button className="secondary" onClick={applyRowPoints}>行別得点合計を問題得点へ反映</button></div>{renderAnswerInput(false)}</section></section>}

    {mode === 'grading' && <section className="mode-section grading-mode-section">{renderQuestionCard()}<section className="focused-answer-area"><div className="focused-answer-title"><div><p className="eyebrow">自分の答案</p><h2>{selectedItem.title}</h2></div><button className="secondary" onClick={applyRowPoints}>行別得点合計を反映</button></div>{renderAnswerInput(false)}</section>{renderModelAnswer()}<section className="panel self-score-box"><h2>自己採点</h2><div className="score-grid"><label>問題得点<input type="number" value={score} onChange={(e) => setScore(Number(e.target.value || 0))} /></label><label>満点<input type="number" value={selectedItem.maxScore} disabled /></label><label>A/B/C判定<select value={rank} onChange={(e) => { const next = e.target.value as ReviewRank; setRank(next); setManualReviewDate(nextReviewDateFor(next, latestReview)); }}><option value="A">A：自力で解けた・説明できる</option><option value="B">B：正解または惜しいが手順に不安</option><option value="C">C：不正解・解説なしでは再現できない</option></select></label><label>次回復習日<input type="date" value={manualReviewDate || nextReviewDateFor(rank, latestReview)} onChange={(e) => setManualReviewDate(e.target.value)} /></label></div><h3>ミス原因</h3><div className="check-list compact-check-list">{missReasons.map((reason) => <label key={reason}><input type="checkbox" checked={selectedReasons.includes(reason)} onChange={() => toggleReason(reason)} />{reason}</label>)}</div><label>復習メモ<textarea value={reviewMemo} onChange={(e) => setReviewMemo(e.target.value)} placeholder="次回解く前に見る注意点" /></label><button className="accent" onClick={() => void saveCurrentAttempt()}>保存して次へ</button></section></section>}

    {mode === 'review' && <section className="mode-section review-mode-section"><section className="panel progress-overview-panel"><div><p className="eyebrow">復習する</p><h2>今日やる問題をカードで確認</h2></div><div className="progress-summary-grid"><div><strong>{dueCards.filter(({ review }) => review.nextReviewDate <= todayString()).length}</strong><span>今日の復習</span></div><div><strong>{dueCards.filter(({ review }) => review.nextReviewDate < todayString()).length}</strong><span>期限超過</span></div><div><strong>{reviews.filter((item) => item.isWeak).length}</strong><span>苦手固定</span></div><div><strong>{reviews.filter((item) => item.latestRank === 'C').length}</strong><span>C判定</span></div><div><strong>{reviews.filter((item) => item.latestRank === 'B').length}</strong><span>B判定</span></div><div><strong>{untouchedItems.length}</strong><span>未着手</span></div></div></section>{dueCards.length === 0 && <section className="panel"><h2>今日の復習はありません</h2><button className="accent" onClick={() => selectPracticeItem(untouchedItems[0] ?? externalRefs[0])}>未着手問題を解く</button></section>}{dueCards.map(({ item, review }) => <article className="panel priority-card" key={review.id}><p className="eyebrow">{item.studyMode === 'external_material' ? '外部教材' : '教材なし'} / {review.isWeak ? '苦手固定' : review.nextReviewDate <= todayString() ? '今日の復習' : '復習候補'}</p><h3>{item.id}　{item.topic}</h3><p>Ref：{item.cpaTrialSectionRef} / 前回：{review.latestRank || '未判定'} / {review.latestScore}点 / 復習日：{review.nextReviewDate}</p><button onClick={() => selectPracticeItem(item)}>解く</button></article>)}</section>}

    {mode === 'management' && <section className="mode-section progress-mode-section"><section className="panel"><p className="eyebrow">管理</p><h2>低頻度機能</h2><p>通常学習では開かなくてよい画面です。PDF抽出・PDF範囲指定は今回実装していません。</p></section><details className="panel management-panel" open><summary>問題ID一覧・品質状態</summary><div className="progress-summary-grid"><div><strong>{externalRefs.length}</strong><span>外部教材区分</span></div><div><strong>{questions.filter((q) => q.verificationStatus === 'approved').length}</strong><span>approved</span></div><div><strong>{questions.filter((q) => q.verificationStatus === 'draft').length}</strong><span>draft</span></div><div><strong>{questions.filter((q) => q.verificationStatus === 'rejected').length}</strong><span>rejected</span></div></div><div className="management-list">{externalRefs.map((item) => <button key={item.id} onClick={() => { setExternalJson(JSON.stringify(item, null, 2)); selectPracticeItem(item); }}>{item.id}｜{item.topic}｜{item.cpaTrialSectionRef}</button>)}{questions.map((item) => <button key={item.id} onClick={() => { setQuestionJson(JSON.stringify(item, null, 2)); if (item.verificationStatus === 'approved') selectPracticeItem(item); }}>{item.id}｜{item.topic}｜{item.cpaTrialSectionRef}｜{item.verificationStatus}</button>)}</div></details><details className="panel management-panel"><summary>外部教材モード用問題ID追加</summary><textarea className="json-editor" value={externalJson} onChange={(e) => setExternalJson(e.target.value)} placeholder="ExternalProblemRef JSON" /><button className="accent" onClick={() => void saveExternalJson()}>外部教材問題IDを保存</button></details><details className="panel management-panel"><summary>オリジナル問題追加・編集・verificationStatus変更</summary><p>通常演習に出るのは verificationStatus が approved の問題だけです。draft は管理画面でのみ確認できます。</p><textarea className="json-editor" value={questionJson} onChange={(e) => setQuestionJson(e.target.value)} placeholder="QuestionItem JSON" /><button className="accent" onClick={() => void saveQuestionJson()}>オリジナル問題を保存</button></details><details className="panel management-panel"><summary>JSONバックアップ・JSON復元 / CSV出力</summary><div className="button-row"><button onClick={() => void handleExport()}>JSONバックアップ</button><label className="import-label">JSON復元<input type="file" accept="application/json" onChange={(e) => void handleImport(e.target.files?.[0])} /></label><button onClick={() => { const csv = attempts.map((a) => [a.answeredAt, a.studyMode, a.questionId, a.cpaTrialSectionRef, a.score, a.maxScore, a.rank, a.nextReviewDate].join(',')).join('\n'); const blob = new Blob([`日時,モード,問題ID,cpaTrialSectionRef,得点,満点,判定,次回復習日\n${csv}`], { type: 'text/csv;charset=utf-8' }); const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = 'study-attempts.csv'; a.click(); URL.revokeObjectURL(url); }}>CSV出力</button></div></details><details className="panel management-panel"><summary>履歴一覧</summary><div className="history-list">{attempts.map((attempt) => <div key={attempt.id}><strong>{attempt.studyMode === 'external_material' ? '外部教材' : '教材なし'}｜{attempt.questionId}</strong><span>{attempt.cpaTrialSectionRef}｜{attempt.score}/{attempt.maxScore}点｜{attempt.rank}｜{attempt.nextReviewDate}</span></div>)}</div></details><details className="panel management-panel"><summary>PDF補助取込</summary><p className="warning-text">PDF抽出、PDF範囲指定、PDF画像表示改善は今回実装していません。CPA教材本文・解答・解説・数値・表構成はアプリに収録しません。</p></details><details className="panel management-panel"><summary>保存状態確認</summary><p>IndexedDBに questions / externalProblemRefs / answerAttempts / reviewSchedules / settings / userProblems / histories を保存します。</p><p>問題数：{summary?.questionCount ?? '-'} / 外部教材区分：{summary?.externalRefCount ?? '-'} / 履歴：{summary?.attemptCount ?? '-'} / 復習予定：{summary?.reviewCount ?? '-'}</p><button onClick={() => void refreshAll()}>保存状態を再確認</button></details></section>}
  </main>;
}
