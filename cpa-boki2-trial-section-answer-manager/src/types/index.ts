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

export interface TemplateExampleGuide {
  description: string;
  howToUse: string[];
  exampleRows: Record<string, string>[];
  scoringExample: string;
  notes: string[];
}

export interface TemplateDefinition {
  id: TemplateId;
  name: string;
  initialRows: number;
  columns: string[];
  optionColumns?: Record<string, string[]>;
  exampleGuide: TemplateExampleGuide;
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

export type QueueStatusLabel = '期限超過' | '今日復習' | 'C判定' | 'B判定' | 'A判定' | '未着手' | '低得点' | 'ミス多い' | '日付未設定C';

export interface StudyQueueItem {
  priority: number;
  statusLabels: QueueStatusLabel[];
  problem: ProblemDefinition;
  latestHistory?: HistoryEntry;
  previousScore?: number;
  maxScore?: number;
  rank: ReviewRank;
  nextReviewDate: string;
  missReasons: MissReason[];
  memoSummary: string;
  reason: string;
}

export interface DashboardProblemStats {
  problem: ProblemDefinition;
  attempts: number;
  latestScore: number | null;
  highestScore: number | null;
  lowestScore: number | null;
  latestRank: ReviewRank;
  latestMissReasons: MissReason[];
  latestReviewDate: string;
  trend: '改善' | '横ばい' | '悪化' | 'データ不足';
}

export interface BackupMetadata {
  id: 'backupMeta';
  lastBackupAt: string;
  lastRestoreAt: string;
  backupCount: number;
  lastBackupHistoryCount: number;
  lastBackupAnswerCount: number;
}

export interface StorageSafetyInfo {
  saveMethod: 'IndexedDB';
  historyCount: number;
  answerCount: number;
  lastBackupAt: string;
  lastRestoreAt: string;
  backupCount: number;
  lastBackupHistoryCount: number;
  lastBackupAnswerCount: number;
  estimatedUsage: number | null;
  estimatedQuota: number | null;
  persistentStatus: '許可済み' | '未許可' | '非対応' | '確認不可';
  storageManagerSupported: boolean;
  warnings: string[];
}

export interface ExamSetProblemState {
  problemId: string;
  status: '未開始' | '開始' | '完了';
  score: number;
  rank: ReviewRank;
}

export interface ExamSetRecord {
  id: string;
  name: string;
  createdAt: string;
  startedAt: string;
  finishedAt: string;
  durationSeconds: number;
  selectedProblemIds: string[];
  problemStates: ExamSetProblemState[];
  problemScores: Record<string, number>;
  totalScore: number;
  passLineReached: boolean;
  memo: string;
  relatedHistoryIds: string[];
}

export interface BackupPayload {
  exportedAt: string;
  appName: string;
  version: string;
  answers: AnswerState[];
  histories: HistoryEntry[];
  settings: Record<string, unknown>[];
  templates: Record<string, unknown>[];
  backups?: Record<string, unknown>[];
  examSets?: ExamSetRecord[];
  backupMetadata?: BackupMetadata | null;
}
