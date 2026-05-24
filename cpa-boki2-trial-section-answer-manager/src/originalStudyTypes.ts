export type ExamType = 'boki2' | 'kensetsu_keiri_2' | 'gyouseishoshi' | 'benrishi' | string;
export type SubjectName = '商業簿記' | '工業簿記';
export type SectionName = '第1問対策' | '第2問対策' | '第3問対策' | '第4問対策' | '第5問対策';
export type Difficulty = '易しい' | '標準' | 'やや難' | '難しい';
export type SourceType = 'original' | 'public_domain' | 'licensed' | 'user_created';
export type AnswerTemplateId = 'journal' | 'numeric' | 'statementTable' | 'free';
export type GradeMark = '未採点' | '○' | '△' | '×';
export type ReviewRank = '' | 'A' | 'B' | 'C';
export type MissReason = '論点理解不足' | '仕訳ミス' | '借方貸方逆' | '金額ミス' | '集計ミス' | '転記ミス' | '表の入力位置ミス' | '下書き不足' | '時間不足' | '解答形式の誤認' | '問題文読み落とし' | 'その他';

export interface AnswerLine {
  id: string;
  itemName?: string;
  debitAccount?: string;
  debitAmount?: string;
  creditAccount?: string;
  creditAmount?: string;
  value?: string;
  value1?: string;
  value2?: string;
  value3?: string;
  value4?: string;
  memo?: string;
  grade: GradeMark;
  points: number;
}

export interface QuestionItem {
  id: string;
  examType: 'boki2';
  subject: SubjectName;
  section: SectionName;
  topic: string;
  title: string;
  questionText: string;
  conditions: string[];
  answerTemplateId: AnswerTemplateId;
  maxScore: number;
  estimatedMinutes?: number;
  difficulty?: Difficulty;
  explanation?: string;
  sourceType: SourceType;
  isVerified: boolean;
  createdAt: string;
  updatedAt: string;
  modelAnswer?: AnswerLine[];
  modelAnswerText?: string;
}

export interface AnswerAttempt {
  id: string;
  questionId: string;
  questionTitle: string;
  subject: SubjectName;
  section: SectionName;
  topic: string;
  answeredAt: string;
  answerTemplateId: AnswerTemplateId;
  rows: AnswerLine[];
  score: number;
  maxScore: number;
  rank: ReviewRank;
  missReasons: MissReason[];
  reviewMemo: string;
  nextReviewDate: string;
}

export interface ReviewSchedule {
  id: string;
  questionId: string;
  nextReviewDate: string;
  latestRank: ReviewRank;
  latestScore: number;
  maxScore: number;
  aStreak: number;
  cStreak: number;
  attempts: number;
  isWeak: boolean;
  updatedAt: string;
}

export interface StudyBackupPayload {
  exportedAt: string;
  appName: string;
  version: string;
  questions: QuestionItem[];
  answerAttempts: AnswerAttempt[];
  reviewSchedules: ReviewSchedule[];
}
