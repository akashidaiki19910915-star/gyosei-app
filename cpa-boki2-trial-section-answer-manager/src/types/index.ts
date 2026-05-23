export type Subject = 'commercial' | 'industrial';
export type SectionId = 'q1' | 'q2' | 'q3' | 'q4' | 'q5';
export type TemplateId = 'journal' | 'genericNumber' | 'equityStatement' | 'consolidation' | 'ledger' | 'financialStatements' | 'processCosting' | 'departmentCosting' | 'standardCosting' | 'cvp' | 'variableFullCosting' | 'freeTable';
export type GradeMark = '未採点' | '○' | '△' | '×';
export type ReviewRank = '' | 'A' | 'B' | 'C';
export type MissReason = '論点理解不足' | '仕訳ミス' | '借方貸方逆' | '金額ミス' | '集計ミス' | '転記ミス' | '表の入力位置ミス' | '下書き不足' | '時間不足' | '解答欄形式の誤認' | '問題文読み落とし' | 'その他';

export interface ProblemDefinition {
  id: string;
  subject: Subject;
  sectionId: SectionId;
  sectionLabel: string;
  displayId: string;
  topic: string;
  defaultTemplateId: TemplateId;
}

export interface TemplateDefinition {
  id: TemplateId;
  name: string;
  initialRows: number;
  columns: string[];
  optionColumns?: Record<string, string[]>;
}

export interface AnswerRow {
  id: string;
  cells: string[];
  grade: GradeMark;
  points: number;
}

export interface AnswerState {
  id: string;
  problemId: string;
  templateId: TemplateId;
  templateName: string;
  columns: string[];
  rows: AnswerRow[];
  draftMemo: string;
  reviewMemo: string;
  score: number;
  maxScore: number;
  rowPointsTotal: number;
  rank: ReviewRank;
  nextReviewDate: string;
  missReasons: MissReason[];
  scoringStarted: boolean;
  scored: boolean;
  scoredAt: string;
  updatedAt: string;
  createdAt: string;
}

export interface HistoryEntry {
  id: string;
  savedAt: string;
  subject: Subject;
  sectionId: SectionId;
  sectionLabel: string;
  problemId: string;
  displayId: string;
  topic: string;
  templateId: TemplateId;
  templateName: string;
  score: number;
  maxScore: number;
  rowPointsTotal: number;
  scored: boolean;
  scoredAt: string;
  rank: ReviewRank;
  nextReviewDate: string;
  missReasons: MissReason[];
  reviewMemo: string;
  draftMemoSummary: string;
  snapshot: AnswerState;
}

export interface BackupPayload {
  exportedAt: string;
  appName: string;
  version: string;
  answers: AnswerState[];
  histories: HistoryEntry[];
  settings: Record<string, unknown>[];
  templates: Record<string, unknown>[];
}
