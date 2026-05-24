import { useEffect, useMemo, useState } from 'react';
import { getTemplateById } from '../data/templateCatalog';
import type { AnswerBlockState, AnswerState, MaterialPdfMapping, MissReason, ProblemBlock, ReviewRank } from '../types';
import { sumRowPoints } from '../utils/scoring';
import { AnswerTable, createRows } from './AnswerTable';

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
  onChange: (answer: AnswerState) => void;
}

function defaultBlockFromAnswer(answer: AnswerState): AnswerBlockState {
  return {
    id: 'legacy-main-block',
    title: '問題文①',
    description: '既存の答案入力テーブルを1ブロックとして表示しています。',
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
    title: block.title || `問題文${index + 1}`,
    description: block.description ?? '',
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
    collapsed: index !== 0,
  };
}

function normalizeBlocks(answer: AnswerState, mapping?: MaterialPdfMapping): AnswerBlockState[] {
  if (answer.problemBlocks && answer.problemBlocks.length > 0) return answer.problemBlocks;
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

export function ProblemBlockAnswerPanel({ answer, mapping, onChange }: Props) {
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

  const setCollapsed = (index: number, collapsed: boolean) => updateBlock(index, { collapsed });
  const expandAll = () => updateBlocks(blocks.map((block) => ({ ...block, collapsed: false })));
  const collapseAll = () => updateBlocks(blocks.map((block, index) => ({ ...block, collapsed: index !== activeIndex })));
  const moveActive = (delta: number) => {
    const next = Math.max(0, Math.min(blocks.length - 1, activeIndex + delta));
    setActiveIndex(next);
    updateBlocks(blocks.map((block, index) => ({ ...block, collapsed: index !== next })));
  };

  return (
    <section className="problem-block-answer-panel">
      <div className="panel block-summary-panel">
        <div className="section-heading">
          <div>
            <p className="eyebrow">問題文ブロック別答案</p>
            <h2>合計 {answer.score} / {answer.maxScore}</h2>
            <p>問題文ごとに答案入力テーブルを分けます。既存データは1ブロックとして扱います。</p>
          </div>
          <div className="button-row">
            <button onClick={expandAll}>すべて展開</button>
            <button onClick={collapseAll}>すべて折りたたみ</button>
            <button onClick={() => moveActive(-1)}>前の問題文へ</button>
            <button onClick={() => moveActive(1)}>次の問題文へ</button>
          </div>
        </div>
      </div>

      {blocks.map((block, index) => {
        const blockAnswer = blockToAnswer(answer, block);
        return (
          <section key={block.id} className={`panel problem-block-card ${index === activeIndex ? 'active-block' : ''}`}>
            <div className="section-heading block-card-heading">
              <div>
                <p className="eyebrow">問題文{index + 1}</p>
                <h2>{block.title}</h2>
                {block.description && <p>{block.description}</p>}
              </div>
              <div className="block-score-box">
                <strong>{block.score} / {block.maxScore || '-'}</strong>
                <span>{block.rank || '未判定'}</span>
              </div>
            </div>
            <div className="block-meta-grid">
              <label>ブロック得点<input type="number" value={block.score} onChange={(event) => updateBlock(index, { score: Number(event.target.value) || 0 })} /></label>
              <label>ブロック満点<input type="number" value={block.maxScore} onChange={(event) => updateBlock(index, { maxScore: Number(event.target.value) || 0 })} /></label>
              <label>ブロック判定
                <select value={block.rank} onChange={(event) => updateBlock(index, { rank: event.target.value as ReviewRank })}>
                  <option value="">未選択</option><option value="A">A</option><option value="B">B</option><option value="C">C</option>
                </select>
              </label>
              <button className="secondary" onClick={() => setCollapsed(index, !block.collapsed)}>{block.collapsed ? '開く' : '折りたたむ'}</button>
            </div>
            {!block.collapsed && (
              <>
                <AnswerTable answer={blockAnswer} template={getTemplateById(block.templateId)} onChange={(next) => updateBlockAnswer(index, next)} onApplyRowPoints={() => applyRowPoints(index)} />
                <div className="block-review-grid">
                  <label>次回注意点<textarea value={block.nextReviewPoint} onChange={(event) => updateBlock(index, { nextReviewPoint: event.target.value })} placeholder="次回このブロックを解く前に見る注意点" /></label>
                  <label>ブロックメモ<textarea value={block.memo} onChange={(event) => updateBlock(index, { memo: event.target.value })} placeholder="教材本文・解答本文は貼らず、自分のミスと処理順だけを書く" /></label>
                </div>
                <div className="check-list block-miss-list">
                  {missReasons.map((reason) => <label key={reason}><input type="checkbox" checked={block.missReasons.includes(reason)} onChange={() => toggleReason(index, reason)} />{reason}</label>)}
                </div>
              </>
            )}
          </section>
        );
      })}
    </section>
  );
}
