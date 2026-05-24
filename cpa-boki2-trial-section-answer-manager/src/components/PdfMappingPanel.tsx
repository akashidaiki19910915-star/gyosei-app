import { useEffect, useMemo, useState } from 'react';
import { problemCatalog } from '../data/problemCatalog';
import type { MaterialPdf, MaterialPdfMapping, TimeMode } from '../types';
import { nowIso } from '../utils/dates';
import { timeModes } from '../utils/disposableStudy';

interface Props {
  pdfs: MaterialPdf[];
  mappings: MaterialPdfMapping[];
  selectedProblemId: string;
  onSave: (mapping: MaterialPdfMapping) => Promise<void>;
  onDelete: (id: string) => Promise<void>;
  onOpenProblem: (problemId: string) => void;
}

function blankMapping(problemId: string, pdfId = ''): MaterialPdfMapping {
  const problem = problemCatalog.find((item) => item.id === problemId) ?? problemCatalog[0];
  const now = nowIso();
  return {
    id: crypto.randomUUID(),
    problemId: problem.id,
    displayId: problem.displayId,
    materialPdfId: pdfId,
    problemPageStart: 1,
    problemPageEnd: 1,
    answerPageStart: undefined,
    answerPageEnd: undefined,
    explanationPageStart: undefined,
    explanationPageEnd: undefined,
    estimatedMinutes: '10分',
    difficulty: '標準',
    memo: '',
    createdAt: now,
    updatedAt: now,
  };
}

function numberOrUndefined(value: string): number | undefined {
  const parsed = Number(value);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : undefined;
}

export function PdfMappingPanel({ pdfs, mappings, selectedProblemId, onSave, onDelete, onOpenProblem }: Props) {
  const [draft, setDraft] = useState<MaterialPdfMapping>(() => blankMapping(selectedProblemId, pdfs[0]?.id ?? ''));
  const mappedProblemIds = useMemo(() => new Set(mappings.map((mapping) => mapping.problemId)), [mappings]);
  const unmapped = problemCatalog.filter((problem) => !mappedProblemIds.has(problem.id));

  useEffect(() => {
    const existing = mappings.find((mapping) => mapping.problemId === selectedProblemId) ?? blankMapping(selectedProblemId, pdfs[0]?.id ?? '');
    setDraft(existing);
  }, [selectedProblemId, mappings, pdfs]);

  const selectedProblem = problemCatalog.find((problem) => problem.id === draft.problemId) ?? problemCatalog[0];
  const update = (patch: Partial<MaterialPdfMapping>) => setDraft((current) => ({ ...current, ...patch, displayId: (problemCatalog.find((problem) => problem.id === (patch.problemId ?? current.problemId)) ?? selectedProblem).displayId, updatedAt: nowIso() }));
  const save = async () => {
    if (!draft.materialPdfId) return;
    await onSave({ ...draft, displayId: selectedProblem.displayId, updatedAt: nowIso() });
  };

  return (
    <details className="panel management-panel">
      <summary>PDFページ紐付け：{mappings.length}件 / 未紐付け {unmapped.length}件</summary>
      <div className="panel-body">
        <div className="form-grid three-columns">
          <label>問題ID
            <select value={draft.problemId} onChange={(event) => update({ problemId: event.target.value })}>
              {problemCatalog.map((problem) => <option key={problem.id} value={problem.id}>{problem.sectionLabel} {problem.displayId} {problem.topic}</option>)}
            </select>
          </label>
          <label>対象PDF
            <select value={draft.materialPdfId} onChange={(event) => update({ materialPdfId: event.target.value })}>
              <option value="">PDFを選択</option>
              {pdfs.map((pdf) => <option key={pdf.id} value={pdf.id}>{pdf.title}</option>)}
            </select>
          </label>
          <label>想定時間
            <select value={draft.estimatedMinutes} onChange={(event) => update({ estimatedMinutes: event.target.value as TimeMode })}>
              {timeModes.map((mode) => <option key={mode} value={mode}>{mode}</option>)}
            </select>
          </label>
        </div>
        <div className="form-grid six-columns">
          <label>問題開始P<input type="number" min="1" value={draft.problemPageStart} onChange={(event) => update({ problemPageStart: Number(event.target.value) || 1 })} /></label>
          <label>問題終了P<input type="number" min="1" value={draft.problemPageEnd} onChange={(event) => update({ problemPageEnd: Number(event.target.value) || draft.problemPageStart })} /></label>
          <label>解答開始P<input type="number" min="1" value={draft.answerPageStart ?? ''} onChange={(event) => update({ answerPageStart: numberOrUndefined(event.target.value) })} /></label>
          <label>解答終了P<input type="number" min="1" value={draft.answerPageEnd ?? ''} onChange={(event) => update({ answerPageEnd: numberOrUndefined(event.target.value) })} /></label>
          <label>解説開始P<input type="number" min="1" value={draft.explanationPageStart ?? ''} onChange={(event) => update({ explanationPageStart: numberOrUndefined(event.target.value) })} /></label>
          <label>解説終了P<input type="number" min="1" value={draft.explanationPageEnd ?? ''} onChange={(event) => update({ explanationPageEnd: numberOrUndefined(event.target.value) })} /></label>
        </div>
        <div className="form-grid two-columns">
          <label>難易度
            <select value={draft.difficulty} onChange={(event) => update({ difficulty: event.target.value as MaterialPdfMapping['difficulty'] })}>
              <option value="易">易</option><option value="標準">標準</option><option value="難">難</option>
            </select>
          </label>
          <label>メモ<input value={draft.memo} onChange={(event) => update({ memo: event.target.value })} placeholder="教材本文は入れず、ページ確認用の短いメモだけ" /></label>
        </div>
        <div className="button-row">
          <button className="accent" disabled={!draft.materialPdfId} onClick={() => { void save(); }}>紐付け保存</button>
          <button className="secondary" onClick={() => setDraft(blankMapping(selectedProblemId, pdfs[0]?.id ?? ''))}>新規作成</button>
        </div>
        <div className="table-wrap short-wrap">
          <table className="compact-table">
            <thead><tr><th>問題ID</th><th>論点</th><th>PDF</th><th>問題P</th><th>解答P</th><th>解説P</th><th>時間</th><th>難易度</th><th>操作</th></tr></thead>
            <tbody>
              {mappings.map((mapping) => {
                const problem = problemCatalog.find((item) => item.id === mapping.problemId) ?? problemCatalog[0];
                const pdf = pdfs.find((item) => item.id === mapping.materialPdfId);
                return (
                  <tr key={mapping.id}>
                    <td>{problem.displayId}</td><td>{problem.topic}</td><td>{pdf?.title ?? 'PDF未登録'}</td>
                    <td>{mapping.problemPageStart}-{mapping.problemPageEnd}</td>
                    <td>{mapping.answerPageStart ? `${mapping.answerPageStart}-${mapping.answerPageEnd ?? mapping.answerPageStart}` : '-'}</td>
                    <td>{mapping.explanationPageStart ? `${mapping.explanationPageStart}-${mapping.explanationPageEnd ?? mapping.explanationPageStart}` : '-'}</td>
                    <td>{mapping.estimatedMinutes}</td><td>{mapping.difficulty}</td>
                    <td><button onClick={() => { setDraft(mapping); onOpenProblem(mapping.problemId); }}>編集</button> <button className="danger" onClick={() => { void onDelete(mapping.id); }}>削除</button></td>
                  </tr>
                );
              })}
              {mappings.length === 0 && <tr><td colSpan={9}>PDFページ紐付けは未登録です。</td></tr>}
            </tbody>
          </table>
        </div>
        {unmapped.length > 0 && <p className="notice-small">未紐付け例：{unmapped.slice(0, 10).map((problem) => `${problem.displayId} ${problem.topic}`).join(' / ')}{unmapped.length > 10 ? ' ...' : ''}</p>}
      </div>
    </details>
  );
}
