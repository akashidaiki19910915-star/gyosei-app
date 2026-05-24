import type { MaterialPdf } from '../types';
import { extractPdfPageText } from './pdfRenderer';

export interface PdfAutoLinkCandidate {
  page: number;
  labels: string[];
  preview: string;
}

const KEYWORDS = ['第1問対策', '第2問対策', '第3問対策', '第4問対策', '第5問対策', '問題', '解答', '解説', '仕訳', '連結', 'CVP', '標準原価', '直接原価計算', '工業簿記', '商業簿記'];

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
