export type ExamType = 'boki2' | 'kensetsu_keiri2' | 'benrishi' | 'gyoseishoshi';
export type StudyMode = 'external_material' | 'built_in_question';
export type SubjectName = '商業簿記' | '工業簿記';
export type SectionName = '第1問対策' | '第2問対策' | '第3問対策' | '第4問対策' | '第5問対策';
export type Difficulty = '易しい' | '標準' | 'やや難' | '難しい';
export type SourceType = 'external_material_reference' | 'initial_original' | 'user_created' | 'law_text' | 'official_past_exam' | 'licensed';
export type VerificationStatus = 'draft' | 'self_checked' | 'source_checked' | 'approved' | 'rejected';
export type CpaTrialSectionRef =
  | 'commercial_1_1'
  | 'commercial_1_2'
  | 'commercial_1_3'
  | 'commercial_1_4'
  | 'commercial_1_5'
  | 'commercial_2_1'
  | 'commercial_2_2'
  | 'commercial_2_3'
  | 'commercial_2_4'
  | 'commercial_2_5'
  | 'commercial_2_6'
  | 'commercial_3_1'
  | 'commercial_3_2'
  | 'commercial_3_3'
  | 'industrial_4_1'
  | 'industrial_4_2'
  | 'industrial_4_3'
  | 'industrial_4_4'
  | 'industrial_5_1'
  | 'industrial_5_2'
  | 'industrial_5_3';
export type ReferenceScope = 'cpa_trial_section_structure_only';
export type AnswerTemplateId = 'journal' | 'numeric' | 'statementTable' | 'accountLedger' | 'free';
export type GradeMark = '未採点' | '正解' | '不正解' | '要確認' | '○' | '△' | '×';
export type ReviewRank = '' | 'A' | 'B' | 'C';
export type MissReason = '論点理解不足' | '仕訳ミス' | '借方貸方逆' | '金額ミス' | '集計ミス' | '転記ミス' | '表の入力位置ミス' | '下書き不足' | '時間不足' | '解答形式の誤認' | '問題文読み落とし' | 'その他';
export type VisualAidType = 'industrial_account_flow' | 'wip_box';

export interface QualityCheck {
  scopeChecked: boolean;
  cpaTrialSectionChecked: boolean;
  accountTitleChecked: boolean;
  journalChecked: boolean;
  amountChecked: boolean;
  explanationChecked: boolean;
  difficultyChecked: boolean;
  copySimilarityChecked: boolean;
  reviewer: string;
  checkedAt: string;
  evidenceMemo: string;
}

export interface IndustrialAccountFlowData {
  title: string;
  description?: string;
  nodes: string[];
  edges: [string, string][];
}

export interface WipBoxRow {
  leftLabel: string;
  rightLabel: string;
}

export interface WipBoxData {
  title: string;
  description?: string;
  rows: WipBoxRow[];
  notes: string[];
}

export interface VisualAidDataMap {
  industrial_account_flow?: IndustrialAccountFlowData;
  wip_box?: WipBoxData;
}

export interface AnswerLine {
  id: string;
  itemName?: string;
  debitDate?: string;
  debitSummary?: string;
  debitAccount?: string;
  debitAmount?: string;
  creditDate?: string;
  creditSummary?: string;
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

export interface ExternalProblemRef {
  id: string;
  examType: 'boki2';
  studyMode: 'external_material';
  qualificationName: '日商簿記2級';
  materialName: string;
  subject: SubjectName;
  section: SectionName;
  topic: string;
  title: string;
  pageMemo: string;
  studyMemo: string;
  answerTemplateId: AnswerTemplateId;
  cpaTrialSectionRef: CpaTrialSectionRef;
  maxScore: number;
  estimatedMinutes?: number;
  difficulty?: Difficulty;
  sourceType: 'external_material_reference';
  referenceScope: ReferenceScope;
  copyrightNote: string;
  visualAidTypes?: VisualAidType[];
  visualAidData?: VisualAidDataMap;
  createdAt: string;
  updatedAt: string;
}

export interface QuestionItem {
  id: string;
  examType: 'boki2';
  studyMode: 'built_in_question';
  subject: SubjectName;
  section: SectionName;
  topic: string;
  title: string;
  questionText: string;
  conditions: string[];
  answerTemplateId: AnswerTemplateId;
  modelAnswer: AnswerLine[];
  explanation: string;
  maxScore: number;
  estimatedMinutes?: number;
  difficulty?: Difficulty;
  sourceType: SourceType;
  verificationStatus: VerificationStatus;
  qualityCheck: QualityCheck;
  cpaTrialSectionRef: CpaTrialSectionRef;
  referenceScope: ReferenceScope;
  copyrightNote: string;
  isVerified: boolean;
  createdAt: string;
  updatedAt: string;
  modelAnswerText?: string;
  visualAidTypes?: VisualAidType[];
  visualAidData?: VisualAidDataMap;
}

export type PracticeItem = ExternalProblemRef | QuestionItem;

export interface AnswerAttempt {
  id: string;
  examType: ExamType;
  studyMode: StudyMode;
  questionId: string;
  questionTitle: string;
  subject: SubjectName;
  section: SectionName;
  topic: string;
  cpaTrialSectionRef: CpaTrialSectionRef;
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
  examType: ExamType;
  studyMode: StudyMode;
  questionId: string;
  cpaTrialSectionRef: CpaTrialSectionRef;
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
}
