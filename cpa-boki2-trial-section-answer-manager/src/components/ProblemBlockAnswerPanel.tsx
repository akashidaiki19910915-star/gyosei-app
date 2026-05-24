import { useEffect, useMemo, useState } from 'react';
import { getTemplateById } from '../data/templateCatalog';
import type { AnswerBlockState, AnswerState, MaterialPdf, MaterialPdfMapping, MissReason, ProblemBlock, ReviewRank } from '../types';
import { sumRowPoints } from '../utils/scoring';
import { AnswerTable, createRows } from './AnswerTable';
import { PdfQuestionPreview } from './PdfQuestionPreview';

const missReasons: MissReason[] = [
  '論点理解不足',
  '仕訳ミス',
  '借方貸方逆',
  '金額ミス',
  '集計ミス',
  '転記ミス',
  '表の入力位置ミス',
  '下書き不足',
  '時間不足',
  '解答欄形式の誤認',
  '問題文読み落とし',
  'その他',
];

interface Props {
  answer: AnswerState;
  mapping?: MaterialPdfMapping;
  pdf?: MaterialPdf;
  mode?: 'practice' | 'grading';
  onChange: (answer: AnswerState) => void;
}

function questionTitle(index: number, title?: string): string {
  const trimmed = title?.trim();
  if (!trimmed || trimmed.startsWith('問題文')) return `設問${index + 1}`;
  return trimmed;
}

function defaultBlockFromAnswer(answer: AnswerState): AnswerBlockState {
  return {
    id: 'legacy-main-block',
    title: '設問1',
    description: 'PDFの該当部分を見て、直下の答案欄へ入力します。',
    templateId: answer.templateId,
    templateName: answer.templateName,
    columns: answer.columns,
    rows: answer.rows,
    score: answer.score,
    maxScore: answer.maxScore,
    rowPointsTotal: answer.rowPointsTotal,
    rank: answer.rank,
    missReasons: answer.missReasons,
    nextReviewPoint: '',
    memo: answer.reviewMemo,
    collapsed: false,
  };
}

function blockFromMapping(block: ProblemBlock, index: number): AnswerBlockState {
  const template = getTemplateById(block.templateId);
  return {
    id: block.id || crypto.randomUUID(),
    title: questionTitle(index, block.title),
    description: block.description ?? '',
    problemPageStart: block.problemPageStart,
    problemPageEnd: block.problemPageEnd,
    cropTopPercent: block.cropTopPercent,
    cropBottomPercent: block.cropBottomPercent,
    answerPageStart: block.answerPageStart,
    answerPageEnd: block.answerPageEnd,
    explanationPageStart: block.explanationPageStart,
    explanationPageEnd: block.explanationPageEnd,
    estimatedMinutes: block.estimatedMinutes,
    difficulty: block.difficulty,
    templateId: block.templateId,
    templateName: template.name,
    columns: template.columns,
    rows: block.rows && block.rows.length > 0 ? block.rows : createRows(template.columns.length, template.initialRows),
    score: 0,
    maxScore: 0,
    rowPointsTotal: 0,
    rank: '',
    missReasons: [],
    nextReviewPoint: '',
    memo: block.memo ?? '',
    collapsed: false,
  };
}

function normalizeBlocks(answer: AnswerState, mapping?: MaterialPdfMapping): AnswerBlockState[] {
  if (answer.problemBlocks && answer.problemBlocks.length > 0) {
    return answer.problemBlocks.map((block, index) => ({ ...block, title: questionTitle(index, block.title) }));
  }
  if (mapping?.problemBlocks && mapping.problemBlocks.length > 0) return mapping.problemBlocks.map(blockFromMapping);
  return [defaultBlockFromAnswer(answer)];
}

function summarize(answer: AnswerState, blocks: AnswerBlockState[]): AnswerState {
  const score = blocks.reduce((total, block) => total + (Number(block.score) || 0), 0);
  const maxScore = blocks.reduce((total, block) => total + (Number(block.maxScore) || 0), 0) || answer.maxScore;
  const rowPointsTotal = blocks.reduce((total, block) => total + (Number(block.rowPointsTotal) || 0), 0);
  const missReasons = Array.from(new Set(blocks.flatMap((block) => block.missReasons)));
  return {
    ...answer,
    problemBlocks: blocks,
    score,
    maxScore,
    rowPointsTotal,
    missReasons,
    rows: blocks[0]?.rows ?? answer.rows,
    columns: blocks[0]?.columns ?? answer.columns,
    templateId: blocks[0]?.templateId ?? answer.templateId,
    templateName: blocks[0]?.templateName ?? answer.templateName,
  };
}

function blockToAnswer(base: AnswerState, block: AnswerBlockState): AnswerState {
  return {
    ...base,
    templateId: block.templateId,
    templateName: block.templateName,
    columns: block.columns,
    rows: block.rows,
    score: block.score,
    maxScore: block.maxScore,
    rowPointsTotal: block.rowPointsTotal,
    rank: block.rank,
    missReasons: block.missReasons,
    reviewMemo: block.memo,
  };
}

export function ProblemBlockAnswerPanel({ answer, mapping, pdf, mode = 'practice', onChange }: Props) {
  const blocks = useMemo(() => normalizeBlocks(answer, mapping), [answer.problemBlocks, answer.problemId, mapping?.id, mapping?.problemBlocks]);
  const [activeIndex, setActiveIndex] = useState(0);

  useEffect(() => {
    if (!answer.problemBlocks || answer.problemBlocks.length === 0) onChange(summarize(answer, blocks));
    // 初期化だけを目的にしているため、answer全体は依存に入れない。
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [answer.problemId, mapping?.id]);

  const updateBlocks = (nextBlocks: AnswerBlockState[]) => onChange(summarize(answer, nextBlocks));

  const updateBlock = (index: number, patch: Partial<AnswerBlockState>) => {
    updateBlocks(blocks.map((block, blockIndex) => blockIndex === index ? { ...block, ...patch } : block));
  };

  const updateBlockAnswer = (index: number, blockAnswer: AnswerState) => {
    updateBlock(index, {
      templateId: blockAnswer.templateId,
      templateName: blockAnswer.templateName,
      columns: blockAnswer.columns,
      rows: blockAnswer.rows,
      score: blockAnswer.score,
      maxScore: blockAnswer.maxScore,
      rowPointsTotal: blockAnswer.rowPointsTotal,
      rank: blockAnswer.rank,
      missReasons: blockAnswer.missReasons,
      memo: blockAnswer.reviewMemo,
    });
  };

  const applyRowPoints = (index: number) => {
    const block = blocks[index];
    const total = sumRowPoints(block.rows);
    updateBlock(index, { rowPointsTotal: total, score: total });
  };

  const toggleReason = (index: number, reason: MissReason) => {
    const block = blocks[index];
    const exists = block.missReasons.includes(reason);
    updateBlock(index, { missReasons: exists ? block.missReasons.filter((item) => item !== reason) : [...block.missReasons, reason] });
  };

  const moveActive = (delta: number) => setActiveIndex((current) => Math.max(0, Math.min(blocks.length - 1, current + delta)));

  return (
    <section className="question-practice-panel">
      <div className="panel question-practice-summary">
        <div className="section-heading">
          <div>
            <p className="eyebrow">設問別演習</p>
            <h2>PDFを読む → 直下に入力する</h2>
            <p>設問ごとに、問題PDFの該当部分を見ながら答案を入力します。採点・メモは必要な時だけ開きます。</p>
          </div>
          <div className="button-row">
            <button onClick={() => moveActive(-1)}>前の設問へ</button>
            <button onClick={() => moveActive(1)}>次の設問へ</button>
          </div>
        </div>
      </div>

      {blocks.map((block, index) => {
        const blockAnswer = blockToAnswer(answer, block);
        const page = block.problemPageStart ?? mapping?.problemPageStart ?? 1;
        return (
          <section key={block.id} className={`panel question-card ${index === activeIndex ? 'active-question-card' : ''}`}>
            <div className="question-card-header">
              <div>
                <p className="eyebrow">設問{index + 1}</p>
                <h2>{questionTitle(index, block.title)}</h2>
                {block.description && <p>{block.description}</p>}
              </div>
              <div className="question-card-meta">
                {block.estimatedMinutes && <span>目安 {block.estimatedMinutes}分</span>}
                {block.difficulty && <span>{block.difficulty}</span>}
              </div>
            </div>

            <PdfQuestionPreview
              pdf={pdf}
              page={page}
              cropTopPercent={block.cropTopPercent}
              cropBottomPercent={block.cropBottomPercent}
              title="PDFの該当部分"
              compact={mode === 'practice'}
            />

            <div className="answer-input-label">答案入力</div>
            <AnswerTable answer={blockAnswer} template={getTemplateById(block.templateId)} onChange={(next) => updateBlockAnswer(index, next)} onApplyRowPoints={() => applyRowPoints(index)} />

            <details className="question-review-details" open={mode === 'grading'}>
              <summary>採点・メモを開く</summary>
              <div className="block-meta-grid question-score-grid">
                <label>得点<input type="number" value={block.score} onChange={(event) => updateBlock(index, { score: Number(event.target.value) || 0 })} /></label>
                <label>満点<input type="number" value={block.maxScore} onChange={(event) => updateBlock(index, { maxScore: Number(event.target.value) || 0 })} /></label>
                <label>判定
                  <select value={block.rank} onChange={(event) => updateBlock(index, { rank: event.target.value as ReviewRank })}>
                    <option value="">未選択</option><option value="A">A</option><option value="B">B</option><option value="C">C</option>
                  </select>
                </label>
              </div>
              <div className="block-review-grid">
                <label>次回注意点<textarea value={block.nextReviewPoint} onChange={(event) => updateBlock(index, { nextReviewPoint: event.target.value })} placeholder="次回この設問を解く前に見る注意点" /></label>
                <label>メモ<textarea value={block.memo} onChange={(event) => updateBlock(index, { memo: event.target.value })} placeholder="教材本文・解答本文は貼らず、自分のミスと処理順だけを書く" /></label>
              </div>
              <div className="check-list block-miss-list">
                {missReasons.map((reason) => <label key={reason}><input type="checkbox" checked={block.missReasons.includes(reason)} onChange={() => toggleReason(index, reason)} />{reason}</label>)}
              </div>
            </details>
          </section>
        );
      })}
    </section>
  );
}
