import { cpaTrialSections, findTrialSection } from './cpaTrialSections';
import type { AnswerLine, CpaTrialSectionRef, QuestionItem, VerificationStatus } from '../originalStudyTypes';

const NOW = '2026-05-24T00:00:00.000Z';
const COPYRIGHT_NOTE = 'CPA問題集の試験対策編の区分のみ参考。問題文・解答・解説・数値・表構成は未使用。';
const INDUSTRIAL_ACCOUNT_FLOW = {
  title: '勘定連絡図',
  description: '材料・賃金・経費が製造活動を通じて仕掛品、製品、売上原価へ流れる関係を確認する図です。',
  nodes: ['材料', '賃金', '経費', '製造間接費', '仕掛品', '製品', '売上原価'],
  edges: [
    ['材料', '仕掛品'],
    ['賃金', '仕掛品'],
    ['経費', '製造間接費'],
    ['製造間接費', '仕掛品'],
    ['仕掛品', '製品'],
    ['製品', '売上原価'],
  ] as [string, string][],
};

function line(input: Partial<AnswerLine>): AnswerLine {
  return { id: crypto.randomUUID(), grade: '未採点', points: 0, ...input };
}

function quality(approved: boolean, memo: string) {
  return {
    scopeChecked: approved,
    cpaTrialSectionChecked: true,
    accountTitleChecked: approved,
    journalChecked: approved,
    amountChecked: approved,
    explanationChecked: approved,
    difficultyChecked: approved,
    copySimilarityChecked: true,
    reviewer: 'ChatGPT implementation draft',
    checkedAt: approved ? '2026-05-24' : '',
    evidenceMemo: memo,
  };
}

function q(input: { ref: CpaTrialSectionRef; title: string; questionText: string; conditions?: string[]; modelAnswer?: AnswerLine[]; explanation?: string; maxScore?: number; status?: VerificationStatus }): QuestionItem {
  const section = findTrialSection(input.ref);
  const approved = (input.status ?? 'approved') === 'approved';
  const base: QuestionItem = {
    id: `短問-${section.id}`,
    examType: 'boki2',
    studyMode: 'built_in_question',
    subject: section.subject,
    section: section.section,
    topic: section.topic,
    title: input.title,
    questionText: input.questionText,
    conditions: input.conditions ?? [],
    answerTemplateId: section.template,
    modelAnswer: input.modelAnswer ?? [line({ itemName: '解答欄', value: '', points: 0 })],
    explanation: input.explanation ?? 'この問題は管理画面で検証後に通常演習へ表示してください。',
    maxScore: input.maxScore ?? Math.min(section.maxScore, 10),
    estimatedMinutes: Math.min(section.minutes, 8),
    difficulty: approved ? '標準' : 'やや難',
    sourceType: 'initial_original',
    verificationStatus: input.status ?? 'approved',
    qualityCheck: quality(approved, approved ? '一般的な簿記2級知識に基づく短問として、仕訳・金額・解説を自己検証済み。CPA教材本文・数値・表構成は未使用。' : '区分対応のみ設定。通常演習には出さない。'),
    cpaTrialSectionRef: input.ref,
    referenceScope: 'cpa_trial_section_structure_only',
    copyrightNote: COPYRIGHT_NOTE,
    isVerified: approved,
    createdAt: NOW,
    updatedAt: NOW,
  };
  if (input.ref === 'industrial_4_1') {
    return {
      ...base,
      visualAidTypes: ['industrial_account_flow'],
      visualAidData: { industrial_account_flow: INDUSTRIAL_ACCOUNT_FLOW },
    };
  }
  return base;
}

const approvedQuestions: QuestionItem[] = [
  q({ ref: 'commercial_1_1', title: '掛仕入と仕入諸掛', questionText: '商品200,000円を掛けで仕入れ、引取運賃6,000円を現金で支払った。引取運賃は仕入原価に含める。必要な仕訳を答えなさい。', conditions: ['税金は考慮しない。'], modelAnswer: [line({ debitAccount: '仕入', debitAmount: '206000', creditAccount: '買掛金', creditAmount: '200000', points: 2 }), line({ debitAccount: '', debitAmount: '', creditAccount: '現金', creditAmount: '6000', points: 2 })], explanation: '仕入諸掛は仕入原価に含めるため、仕入は206,000円になる。貸方は買掛金と現金支払に分ける。', maxScore: 4 }),
  q({ ref: 'commercial_1_2', title: '電子記録債権への振替', questionText: '売掛金120,000円について、取引先の承諾を得て電子記録債権へ振り替えた。必要な仕訳を答えなさい。', modelAnswer: [line({ debitAccount: '電子記録債権', debitAmount: '120000', creditAccount: '売掛金', creditAmount: '120000', points: 4 })], explanation: '売掛金という通常債権を電子記録債権へ振り替える。', maxScore: 4 }),
  q({ ref: 'commercial_1_3', title: '備品売却損', questionText: '取得原価300,000円、減価償却累計額180,000円の備品を100,000円で売却し、代金は翌月受け取ることにした。必要な仕訳を答えなさい。', conditions: ['未収入金勘定を用いる。'], modelAnswer: [line({ debitAccount: '減価償却累計額', debitAmount: '180000', creditAccount: '備品', creditAmount: '300000', points: 2 }), line({ debitAccount: '未収入金', debitAmount: '100000', creditAccount: '', creditAmount: '', points: 2 }), line({ debitAccount: '固定資産売却損', debitAmount: '20000', creditAccount: '', creditAmount: '', points: 2 })], explanation: '帳簿価額は120,000円。売却価額100,000円との差額20,000円は売却損。', maxScore: 6 }),
  q({ ref: 'commercial_1_4', title: 'ファイナンス・リース開始', questionText: '所有権移転外ファイナンス・リース取引として機械を取得した。リース料総額500,000円、利息相当額は考慮しない。開始時の仕訳を答えなさい。', conditions: ['リース資産、リース債務を用いる。'], modelAnswer: [line({ debitAccount: 'リース資産', debitAmount: '500000', creditAccount: 'リース債務', creditAmount: '500000', points: 4 })], explanation: 'ファイナンス・リースは売買取引に準じて資産と債務を認識する。', maxScore: 4 }),
  q({ ref: 'commercial_1_5', title: '外貨建売掛金の決済', questionText: '外貨建売掛金1,000ドルを決済し、普通預金に入金された。売掛金計上時は1ドル140円、決済時は1ドル145円である。必要な仕訳を答えなさい。', modelAnswer: [line({ debitAccount: '普通預金', debitAmount: '145000', creditAccount: '売掛金', creditAmount: '140000', points: 3 }), line({ debitAccount: '', debitAmount: '', creditAccount: '為替差益', creditAmount: '5000', points: 3 })], explanation: '入金145,000円と売掛金帳簿価額140,000円との差額5,000円は為替差益。', maxScore: 6 }),
  q({ ref: 'industrial_4_1', title: '材料消費額の計上', questionText: '直接材料80,000円、間接材料12,000円を当月製造のために消費した。必要な仕訳を答えなさい。', conditions: ['直接材料は仕掛品、間接材料は製造間接費へ振り替える。'], modelAnswer: [line({ debitAccount: '仕掛品', debitAmount: '80000', creditAccount: '材料', creditAmount: '92000', points: 2 }), line({ debitAccount: '製造間接費', debitAmount: '12000', creditAccount: '', creditAmount: '', points: 2 })], explanation: '直接材料は仕掛品へ、間接材料は製造間接費へ集計する。', maxScore: 4 }),
  q({ ref: 'industrial_5_1', title: '直接材料費差異', questionText: '標準価格は1kgあたり500円、標準消費量は完成品1個あたり2kgである。完成品100個、実際消費量230kg、実際価格480円/kgの場合、価格差異と数量差異を求めなさい。', conditions: ['有利差異はプラス、不利差異はマイナスで入力する。'], modelAnswer: [line({ itemName: '価格差異', value: '4600', points: 4 }), line({ itemName: '数量差異', value: '-15000', points: 4 })], explanation: '価格差異=(500-480)×230=4,600円有利。数量差異=(200kg-230kg)×500=-15,000円不利。', maxScore: 8 }),
  q({ ref: 'industrial_5_2', title: '損益分岐点売上高', questionText: '販売価格1,000円、変動費率60%、固定費200,000円である。損益分岐点売上高を求めなさい。', conditions: ['金額は円単位で入力する。'], modelAnswer: [line({ itemName: '損益分岐点売上高', value: '500000', points: 6 })], explanation: '貢献利益率は40%。200,000円÷40%=500,000円。', maxScore: 6 }),
  q({ ref: 'industrial_5_3', title: '全部原価計算と直接原価計算の利益差', questionText: '固定製造間接費は月額120,000円。当月は生産1,000個、販売900個、月初製品なしである。全部原価計算の営業利益が直接原価計算よりいくら大きいか求めなさい。', conditions: ['差額は円単位で入力する。'], modelAnswer: [line({ itemName: '営業利益差額', value: '12000', points: 6 })], explanation: '固定製造間接費の単価は120円。期末在庫100個に含まれる固定費12,000円だけ、全部原価計算の利益が大きい。', maxScore: 6 }),
];

const approvedRefs = new Set(approvedQuestions.map((question) => question.cpaTrialSectionRef));
const draftQuestions: QuestionItem[] = cpaTrialSections.filter((section) => !approvedRefs.has(section.ref)).map((section) => q({
  ref: section.ref,
  title: `${section.topic}（ドラフト）`,
  questionText: 'この区分は試験対策編の対応カテゴリとして登録済みです。問題本文・模範解答・解説は未検証のため、通常演習には表示しません。',
  conditions: ['管理画面で検証後、verificationStatusをapprovedにしてください。'],
  status: 'draft',
  maxScore: Math.min(section.maxScore, 10),
}));

export const originalQuestions: QuestionItem[] = [...approvedQuestions, ...draftQuestions];
export const approvedOriginalQuestions: QuestionItem[] = originalQuestions.filter((question) => question.verificationStatus === 'approved');
export const draftOriginalQuestions: QuestionItem[] = originalQuestions.filter((question) => question.verificationStatus === 'draft');
