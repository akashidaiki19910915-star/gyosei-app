import { problemCatalog } from '../data/problemCatalog';
import type { ExtractedPdfPage, MaterialPdf, PdfProblemExtractionCandidate, ProblemDefinition, SectionId, Subject, TemplateId } from '../types';
import { extractAllPdfPageTexts, extractPdfPageText } from './pdfRenderer';

export interface PdfAutoLinkCandidate {
  page: number;
  labels: string[];
  preview: string;
}

const KEYWORDS = ['第1問対策', '第2問対策', '第3問対策', '第4問対策', '第5問対策', '問題', '解答', '解説', '仕訳', '連結', 'CVP', '標準原価', '直接原価計算', '工業簿記', '商業簿記'];

const SECTION_PATTERNS: Record<SectionId, RegExp[]> = {
  q1: [/第\s*1\s*問/, /1\s*-\s*\d+/],
  q2: [/第\s*2\s*問/, /2\s*-\s*\d+/],
  q3: [/第\s*3\s*問/, /3\s*-\s*\d+/],
  q4: [/第\s*4\s*問/, /4\s*-\s*\d+/],
  q5: [/第\s*5\s*問/, /5\s*-\s*\d+/],
};

function compact(text: string): string {
  return text.replace(/\s+/g, ' ').trim();
}

function escapeRegExp(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function containsProblemId(text: string, displayId: string): boolean {
  const escaped = escapeRegExp(displayId).replace('\\-', '\\s*-\\s*');
  return new RegExp(`(^|[^0-9])${escaped}([^0-9]|$)`).test(text);
}

function subjectFromText(text: string): Subject | '' {
  if (/工業簿記|原価計算|標準原価|CVP|直接原価|部門別|工程別/.test(text)) return 'industrial';
  if (/商業簿記|連結|株主資本|有価証券|本支店|決算/.test(text)) return 'commercial';
  return '';
}

function scorePageForProblem(page: ExtractedPdfPage, problem: ProblemDefinition): number {
  const text = page.text;
  let score = 0;
  if (containsProblemId(text, problem.displayId)) score += 55;
  if (text.includes(problem.sectionLabel)) score += 18;
  if (problem.topic.split(/[、・\s]+/).filter(Boolean).some((word) => text.includes(word))) score += 12;
  if (SECTION_PATTERNS[problem.sectionId].some((pattern) => pattern.test(text))) score += 10;
  const subject = subjectFromText(text);
  if (subject && subject === problem.subject) score += 5;
  if (/解答|解説|解答・解説/.test(text) && !/問題|問\d|資料|次の/.test(text)) score -= 25;
  return score;
}

function buildAnswerText(pages: ExtractedPdfPage[], startPage: number): { answerText: string; answerPages: number[]; explanationText: string; explanationPages: number[] } {
  const nearby = pages.filter((page) => page.page > startPage && page.page <= startPage + 8);
  const answerPages = nearby.filter((page) => /解答|答案|正解/.test(page.text)).slice(0, 3);
  const explanationPages = nearby.filter((page) => /解説|ポイント|考え方|仕訳/.test(page.text)).slice(0, 3);
  return {
    answerText: compact(answerPages.map((page) => `P${page.page} ${page.text}`).join('\n\n')).slice(0, 12000),
    answerPages: answerPages.map((page) => page.page),
    explanationText: compact(explanationPages.map((page) => `P${page.page} ${page.text}`).join('\n\n')).slice(0, 12000),
    explanationPages: explanationPages.map((page) => page.page),
  };
}

function confidenceToStatus(confidence: number): PdfProblemExtractionCandidate['status'] {
  if (confidence >= 75) return 'autoLinked';
  if (confidence >= 35) return 'needsReview';
  return 'unmatched';
}

function templateKind(templateId: TemplateId): PdfProblemExtractionCandidate['problemKind'] {
  if (templateId === 'journal') return 'journal';
  if (templateId === 'genericNumber') return 'numericInput';
  return 'calculationTable';
}

export async function extractProblemCandidatesFromPdf(pdf: MaterialPdf, onProgress?: (page: number, pageCount: number) => void): Promise<PdfProblemExtractionCandidate[]> {
  const pages = await extractAllPdfPageTexts(pdf.pdfBlob, onProgress);
  const usedPages = new Set<number>();
  return problemCatalog.map((problem) => {
    const ranked = pages
      .filter((page) => page.text.trim().length > 20)
      .map((page) => ({ page, score: scorePageForProblem(page, problem) }))
      .sort((a, b) => b.score - a.score);
    const best = ranked[0];
    const fallback = pages.find((page) => !usedPages.has(page.page) && page.text.trim().length > 60) ?? pages[0];
    const page = best && best.score > 0 ? best.page : fallback;
    const confidence = Math.max(0, Math.min(100, best?.score ?? 0));
    if (page) usedPages.add(page.page);
    const answer = page ? buildAnswerText(pages, page.page) : { answerText: '', answerPages: [], explanationText: '', explanationPages: [] };
    const problemText = page ? compact(page.text).slice(0, 16000) : '';
    const status = confidenceToStatus(confidence);
    return {
      id: crypto.randomUUID(),
      materialPdfId: pdf.id,
      suggestedProblemId: problem.id,
      displayId: problem.displayId,
      subject: problem.subject,
      sectionId: problem.sectionId,
      sectionLabel: problem.sectionLabel,
      topic: problem.topic,
      templateId: problem.defaultTemplateId,
      problemKind: templateKind(problem.defaultTemplateId),
      problemText,
      answerText: answer.answerText,
      explanationText: answer.explanationText,
      sourcePages: page ? [page.page] : [],
      answerPages: answer.answerPages,
      explanationPages: answer.explanationPages,
      confidence,
      status,
      reason: status === 'autoLinked' ? '問題ID・見出し・論点語の一致度が高い候補です。' : status === 'needsReview' ? '候補はありますが、手動確認を推奨します。' : '自動紐づけできませんでした。手動確認してください。',
    };
  });
}

export async function suggestPdfLinkCandidates(pdf: MaterialPdf, maxPages = 80): Promise<PdfAutoLinkCandidate[]> {
  const limit = Math.min(pdf.pageCount, maxPages);
  const candidates: PdfAutoLinkCandidate[] = [];
  for (let page = 1; page <= limit; page += 1) {
    try {
      const text = await extractPdfPageText(pdf.pdfBlob, page);
      if (!text.trim()) continue;
      const labels = KEYWORDS.filter((keyword) => text.includes(keyword));
      if (labels.length === 0) continue;
      candidates.push({ page, labels, preview: text.replace(/\s+/g, ' ').slice(0, 120) });
    } catch {
      // テキスト抽出できないページは候補なしとして扱う。OCRは行わない。
    }
  }
  return candidates;
}
