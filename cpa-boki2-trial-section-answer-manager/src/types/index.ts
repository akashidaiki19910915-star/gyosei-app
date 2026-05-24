export type Subject = 'commercial' | 'industrial';
export type SectionId = 'q1' | 'q2' | 'q3' | 'q4' | 'q5';
export type TemplateId = 'journal' | 'genericNumber' | 'equityStatement' | 'consolidation' | 'ledger' | 'financialStatements' | 'processCosting' | 'departmentCosting' | 'standardCosting' | 'cvp' | 'variableFullCosting' | 'freeTable';
export type GradeMark = '未採点' | '○' | '△' | '×';
export type ReviewRank = '' | 'A' | 'B' | 'C';
export type MissReason = '論点理解不足' | '仕訳ミス' | '借方貸方逆' | '金額ミス' | '集計ミス' | '転記ミス' | '表の入力位置ミス' | '下書き不足' | '時間不足' | '解答欄形式の誤認' | '問題文読み落とし' | 'その他';
export type TimeMode = '1分' | '3分' | '5分' | '10分' | '30分' | '90分';
export type GradingStatus = '正解' | '部分正解' | '不正解' | '迷いあり';
export type BlockDifficulty = '易しい' | '標準' | 'やや難' | '難しい';

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

export interface ProblemBlock {
  id: string;
  title: string;
  description?: string;
  problemPageStart?: number;
  problemPageEnd?: number;
  answerPageStart?: number;
  answerPageEnd?: number;
  explanationPageStart?: number;
  explanationPageEnd?: number;
  estimatedMinutes?: number;
  difficulty?: BlockDifficulty;
  templateId: TemplateId;
  rows?: AnswerRow[];
  memo?: string;
}

export interface AnswerBlockState {
  id: string;
  title: string;
  description?: string;
  templateId: TemplateId;
  templateName: string;
  columns: string[];
  rows: AnswerRow[];
  score: number;
  maxScore: number;
  rowPointsTotal: number;
  rank: ReviewRank;
  missReasons: MissReason[];
  nextReviewPoint: string;
  memo: string;
  collapsed?: boolean;
}

export interface AnswerState {
  id: string;
  problemId: string;
  templateId: TemplateId;
  templateName: string;
  columns: string[];
  rows: AnswerRow[];
  problemBlocks?: AnswerBlockState[];
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

export type MasteryStatus = '未着手' | '1回演習済み' | '2回演習済み' | '3回以上演習済み' | '合格水準' | '再復習対象' | '危険問題' | '期限超過';
export type MasteryFilter = '全体' | '未着手' | '危険問題' | '期限超過' | 'C判定' | 'B判定' | '合格水準' | '商業簿記' | '工業簿記' | '第1問対策' | '第2問対策' | '第3問対策' | '第4問対策' | '第5問対策';

export interface MasteryMapItem {
  problem: ProblemDefinition;
  latestScore: number | null;
  maxScore: number | null;
  scoreRate: number | null;
  latestRank: ReviewRank;
  attempts: number;
  aCount: number;
  bCount: number;
  cCount: number;
  lastPracticedAt: string;
  nextReviewDate: string;
  status: MasteryStatus;
  action: string;
  sortPriority: number;
}

export interface MistakeCard {
  id: string;
  problemId: string;
  correctFlow: string;
  mistakeReason: string;
  firstReviewPoint: string;
  preSolveChecklist: string;
  preventionPhrase: string;
  missReasons: MissReason[];
  importance: '高' | '中' | '低';
  createdAt: string;
  updatedAt: string;
}

export interface RecoveryPlanRow {
  sectionId: SectionId;
  sectionLabel: string;
  score: number;
  maxScore: number;
  lostPoints: number;
}

export interface RecoveryPlan {
  source: '直近90分セット' | '問題別履歴';
  totalScore: number;
  maxScore: number;
  shortage70: number;
  shortage72: number;
  shortage80: number;
  sectionRows: RecoveryPlanRow[];
  biggestLossSection: string;
  mostFrequentMissReason: string;
  nextFocus: string;
  recoveryCandidates: string[];
  latestExamSet?: ExamSetRecord;
}

export interface MaterialPdf {
  id: string;
  title: string;
  sourceName: string;
  examType: '日商簿記2級';
  bookType: '商業簿記' | '工業簿記' | 'その他';
  fileName: string;
  mimeType: string;
  size: number;
  pageCount: number;
  pdfBlob: Blob;
  createdAt: string;
  updatedAt: string;
}

export type MaterialPdfMetadata = Omit<MaterialPdf, 'pdfBlob'>;

export interface MaterialPdfMapping {
  id: string;
  problemId: string;
  displayId: string;
  materialPdfId: string;
  problemPageStart: number;
  problemPageEnd: number;
  answerPageStart?: number;
  answerPageEnd?: number;
  explanationPageStart?: number;
  explanationPageEnd?: number;
  estimatedMinutes: TimeMode;
  difficulty: '易' | '標準' | '難';
  memo: string;
  problemBlocks?: ProblemBlock[];
  createdAt: string;
  updatedAt: string;
}

export interface ReviewState {
  id: string;
  problemId: string;
  correctStreak: number;
  wrongStreak: number;
  lastResult: GradingStatus | '';
  lastReviewedAt: string;
  nextReviewDate: string;
  intervalDays: number;
  easeLevel: number;
  updatedAt: string;
}

export interface GradingResult {
  status: GradingStatus | '';
  score: number;
  maxScore: number;
  scoreRate: number;
  autoGraded: false;
  detailRows: string[];
}

export interface PracticeSession {
  id: string;
  examType: '日商簿記2級';
  problemId: string;
  startedAt: string;
  submittedAt: string;
  durationSeconds: number;
  selectedTimeMode: TimeMode;
  answerSnapshot?: AnswerState;
  gradingMode: 'self';
  gradingResult: GradingResult;
  reviewState?: ReviewState;
  openedAnswer: boolean;
  openedExplanation: boolean;
  memo: string;
  completed: boolean;
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
  mistakeCards?: MistakeCard[];
  materialPdfMappings?: MaterialPdfMapping[];
  materialPdfMetadata?: MaterialPdfMetadata[];
  practiceSessions?: PracticeSession[];
  reviewStates?: ReviewState[];
  backupMetadata?: BackupMetadata | null;
}