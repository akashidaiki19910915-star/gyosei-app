import { problemCatalog, subjectLabels } from '../data/problemCatalog';
import type { AnswerInputType, AnswerTemplateType, MaterialPdfMapping, PdfProblemExtractionCandidate, ProblemDefinition, QuestionCard, TemplateId } from '../types';
import { nowIso } from './dates';

const NO_QUESTION_TEXT = 'この問題は問題文未登録です。原本を見るか、教材設定で問題文を登録してください。';
const HEADER_PATTERNS = [
  /試験対策編/g,
  /いちばんわかる日商簿記2級/g,
  /CPAラーニング/g,
  /第\s*[1-5]\s*問対策(?:[-ー－]?\d+)?/g,
  /商業簿記|工業簿記/g,
  /解答・解説|解答解説/g,
];

function templateToInputType(templateId?: TemplateId): AnswerInputType {
  if (templateId === 'journal') return 'journalEntry';
  if (templateId === 'genericNumber') return 'numeric';
  return 'memoOnly';
}

function templateToTemplateType(templateId?: TemplateId): AnswerTemplateType {
  if (templateId === 'journal') return '仕訳テーブル';
  if (templateId === 'genericNumber') return '金額計算';
  return 'メモのみ';
}

function difficultyToMinutes(value?: MaterialPdfMapping['estimatedMinutes']): number {
  const parsed = Number(String(value ?? '').replace(/[^0-9]/g, ''));
  return Number.isFinite(parsed) && parsed > 0 ? parsed : 10;
}

function stripNoise(raw: string, displayId?: string): string {
  let text = raw.replace(/\r\n?/g, '\n').replace(/[\t ]+/g, ' ');
  HEADER_PATTERNS.forEach((pattern) => { text = text.replace(pattern, ' '); });
  text = text.replace(/(^|\s)\d{1,4}(?=\s|$)/g, ' ');
  if (displayId) {
    const escaped = displayId.replace('-', '\\s*-\\s*');
    text = text.replace(new RegExp(`問題?\s*${escaped}`, 'g'), ' ');
    text = text.replace(new RegExp(`(^|\\s)${escaped}(?=\\s|$)`, 'g'), ' ');
  }
  text = text.replace(/\s*\(第\s*[1-5]\s*問対策[-ー－]?\d*\)\s*/g, ' ');
  text = text.replace(/\s+/g, ' ').trim();
  return text;
}

function splitChoices(text: string): { body: string; choices: string[] } {
  const choicePattern = /([アイウエオカキクケコ])\s*[　 ]*([^アイウエオカキクケコ\n]{1,30})(?=\s+[アイウエオカキクケコ]\s|$)/g;
  const choices: string[] = [];
  let match: RegExpExecArray | null;
  while ((match = choicePattern.exec(text)) !== null) {
    const label = match[1];
    const value = match[2].trim();
    if (value && !choices.includes(`${label}　${value}`)) choices.push(`${label}　${value}`);
  }
  if (choices.length === 0) return { body: text, choices };
  const firstChoiceIndex = text.search(/アイウエオ|ア\s|ア　/);
  return {
    body: firstChoiceIndex > 0 ? text.slice(0, firstChoiceIndex).trim() : text.replace(choicePattern, '').trim(),
    choices,
  };
}

function paragraphize(text: string): string {
  const normalized = text.replace(/。\s*/g, '。\n\n').replace(/(次の[^。]+。)/, '$1\n\n');
  return normalized.replace(/\n{3,}/g, '\n\n').trim();
}

export function cleanupQuestionText(raw: string, displayId?: string): { questionText: string; choices: string[]; needsReview: boolean } {
  const stripped = stripNoise(raw, displayId);
  if (!stripped || stripped.length < 20 || /�{2,}/.test(stripped)) return { questionText: NO_QUESTION_TEXT, choices: [], needsReview: true };
  const answerStart = stripped.search(/解答|解説|正解|答案/);
  const problemOnly = answerStart > 40 ? stripped.slice(0, answerStart).trim() : stripped;
  const { body, choices } = splitChoices(problemOnly);
  const questionText = paragraphize(body);
  return { questionText: questionText || NO_QUESTION_TEXT, choices, needsReview: questionText.length < 20 };
}

export function questionCardFromMapping(mapping: MaterialPdfMapping): QuestionCard {
  const problem = problemCatalog.find((item) => item.id === mapping.problemId) ?? problemCatalog[0];
  const rawText = mapping.problemBlocks?.map((block) => block.questionText).filter(Boolean).join('\n\n') || mapping.problemText || '';
  const cleaned = cleanupQuestionText(rawText, problem.displayId);
  const now = mapping.updatedAt || nowIso();
  return {
    id: mapping.problemId,
    examType: '日商簿記2級',
    subject: subjectLabels[problem.subject],
    section: problem.sectionLabel,
    topic: problem.topic,
    title: `${problem.displayId}：${problem.topic}`,
    questionNumber: problem.displayId,
    questionText: cleaned.questionText,
    choices: cleaned.choices,
    answerInputType: templateToInputType(problem.defaultTemplateId),
    answerTemplateType: templateToTemplateType(problem.defaultTemplateId),
    estimatedMinutes: difficultyToMinutes(mapping.estimatedMinutes),
    difficulty: mapping.difficulty,
    sourceMaterialId: mapping.materialPdfId,
    sourcePageStart: mapping.problemPageStart,
    sourcePageEnd: mapping.problemPageEnd,
    answerPageStart: mapping.answerPageStart,
    answerPageEnd: mapping.answerPageEnd,
    explanationPageStart: mapping.explanationPageStart,
    explanationPageEnd: mapping.explanationPageEnd,
    extractionConfidence: mapping.extractionConfidence ?? 0,
    needsReview: cleaned.needsReview || mapping.extractionStatus !== 'autoLinked',
    sourceProblemId: mapping.problemId,
    sourceMappingId: mapping.id,
    createdAt: mapping.createdAt || now,
    updatedAt: now,
  };
}

export function fallbackQuestionCard(problem: ProblemDefinition): QuestionCard {
  const now = nowIso();
  return {
    id: problem.id,
    examType: '日商簿記2級',
    subject: subjectLabels[problem.subject],
    section: problem.sectionLabel,
    topic: problem.topic,
    title: `${problem.displayId}：${problem.topic}`,
    questionNumber: problem.displayId,
    questionText: NO_QUESTION_TEXT,
    choices: [],
    answerInputType: templateToInputType(problem.defaultTemplateId),
    answerTemplateType: templateToTemplateType(problem.defaultTemplateId),
    estimatedMinutes: 10,
    difficulty: '標準',
    sourceMaterialId: '',
    sourcePageStart: 1,
    sourcePageEnd: 1,
    extractionConfidence: 0,
    needsReview: true,
    sourceProblemId: problem.id,
    createdAt: now,
    updatedAt: now,
  };
}

export function questionCardFromCandidate(candidate: PdfProblemExtractionCandidate): QuestionCard {
  const problem = problemCatalog.find((item) => item.id === candidate.suggestedProblemId) ?? problemCatalog[0];
  const cleaned = cleanupQuestionText(candidate.problemText, problem.displayId);
  const now = nowIso();
  const pages = candidate.sourcePages.length ? [...candidate.sourcePages].sort((a, b) => a - b) : [1];
  const answerPages = candidate.answerPages.length ? [...candidate.answerPages].sort((a, b) => a - b) : [];
  const explanationPages = candidate.explanationPages.length ? [...candidate.explanationPages].sort((a, b) => a - b) : [];
  return {
    id: `${candidate.materialPdfId}-${problem.id}`,
    examType: '日商簿記2級',
    subject: subjectLabels[problem.subject],
    section: problem.sectionLabel,
    topic: problem.topic,
    title: `${problem.displayId}：${problem.topic}`,
    questionNumber: problem.displayId,
    questionText: cleaned.questionText,
    choices: cleaned.choices,
    answerInputType: templateToInputType(candidate.templateId),
    answerTemplateType: templateToTemplateType(candidate.templateId),
    estimatedMinutes: 10,
    difficulty: candidate.confidence >= 75 ? '標準' : '難',
    sourceMaterialId: candidate.materialPdfId,
    sourcePageStart: pages[0],
    sourcePageEnd: pages[pages.length - 1],
    answerPageStart: answerPages[0],
    answerPageEnd: answerPages[answerPages.length - 1],
    explanationPageStart: explanationPages[0],
    explanationPageEnd: explanationPages[explanationPages.length - 1],
    extractionConfidence: candidate.confidence,
    needsReview: cleaned.needsReview || candidate.status !== 'autoLinked',
    sourceProblemId: problem.id,
    createdAt: now,
    updatedAt: now,
  };
}
