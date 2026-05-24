import type { AnswerTemplateId, CpaTrialSectionRef, ExternalProblemRef, SectionName, SubjectName } from '../originalStudyTypes';

const NOW = '2026-05-24T00:00:00.000Z';
const COPYRIGHT_NOTE = 'CPA問題集の試験対策編の区分のみ参考。問題文・解答・解説・数値・表構成は未使用。';

type SectionDef = {
  ref: CpaTrialSectionRef;
  id: string;
  subject: SubjectName;
  section: SectionName;
  topic: string;
  template: AnswerTemplateId;
  maxScore: number;
  minutes: number;
};

export const cpaTrialSections: SectionDef[] = [
  { ref: 'commercial_1_1', id: '1-1', subject: '商業簿記', section: '第1問対策', topic: '商品、収益認識、現金預金', template: 'journal', maxScore: 20, minutes: 12 },
  { ref: 'commercial_1_2', id: '1-2', subject: '商業簿記', section: '第1問対策', topic: '債権債務、有価証券', template: 'journal', maxScore: 20, minutes: 12 },
  { ref: 'commercial_1_3', id: '1-3', subject: '商業簿記', section: '第1問対策', topic: '有形固定資産、無形資産', template: 'journal', maxScore: 20, minutes: 12 },
  { ref: 'commercial_1_4', id: '1-4', subject: '商業簿記', section: '第1問対策', topic: 'リース取引、引当金、株式会社会計', template: 'journal', maxScore: 20, minutes: 12 },
  { ref: 'commercial_1_5', id: '1-5', subject: '商業簿記', section: '第1問対策', topic: '外貨建取引、税効果会計、本支店会計、連結会計', template: 'journal', maxScore: 20, minutes: 12 },
  { ref: 'commercial_2_1', id: '2-1', subject: '商業簿記', section: '第2問対策', topic: '株主資本等変動計算書', template: 'statementTable', maxScore: 20, minutes: 18 },
  { ref: 'commercial_2_2', id: '2-2', subject: '商業簿記', section: '第2問対策', topic: '連結会計①', template: 'statementTable', maxScore: 20, minutes: 18 },
  { ref: 'commercial_2_3', id: '2-3', subject: '商業簿記', section: '第2問対策', topic: '連結会計②', template: 'statementTable', maxScore: 20, minutes: 18 },
  { ref: 'commercial_2_4', id: '2-4', subject: '商業簿記', section: '第2問対策', topic: '勘定記入①（商品）', template: 'accountLedger', maxScore: 20, minutes: 18 },
  { ref: 'commercial_2_5', id: '2-5', subject: '商業簿記', section: '第2問対策', topic: '勘定記入②（有形固定資産）', template: 'accountLedger', maxScore: 20, minutes: 18 },
  { ref: 'commercial_2_6', id: '2-6', subject: '商業簿記', section: '第2問対策', topic: '勘定記入③（有価証券）', template: 'accountLedger', maxScore: 20, minutes: 18 },
  { ref: 'commercial_3_1', id: '3-1', subject: '商業簿記', section: '第3問対策', topic: '決算の総合問題①', template: 'free', maxScore: 20, minutes: 22 },
  { ref: 'commercial_3_2', id: '3-2', subject: '商業簿記', section: '第3問対策', topic: '決算の総合問題②', template: 'free', maxScore: 20, minutes: 22 },
  { ref: 'commercial_3_3', id: '3-3', subject: '商業簿記', section: '第3問対策', topic: '本支店会計の総合問題', template: 'free', maxScore: 20, minutes: 22 },
  { ref: 'industrial_4_1', id: '4-1', subject: '工業簿記', section: '第4問対策', topic: '工業簿記 仕訳問題', template: 'journal', maxScore: 28, minutes: 15 },
  { ref: 'industrial_4_2', id: '4-2', subject: '工業簿記', section: '第4問対策', topic: '工程別総合原価計算', template: 'statementTable', maxScore: 28, minutes: 18 },
  { ref: 'industrial_4_3', id: '4-3', subject: '工業簿記', section: '第4問対策', topic: '部門別原価計算', template: 'statementTable', maxScore: 28, minutes: 18 },
  { ref: 'industrial_4_4', id: '4-4', subject: '工業簿記', section: '第4問対策', topic: '財務諸表作成', template: 'statementTable', maxScore: 28, minutes: 18 },
  { ref: 'industrial_5_1', id: '5-1', subject: '工業簿記', section: '第5問対策', topic: '標準原価計算', template: 'numeric', maxScore: 12, minutes: 12 },
  { ref: 'industrial_5_2', id: '5-2', subject: '工業簿記', section: '第5問対策', topic: 'CVP分析', template: 'numeric', maxScore: 12, minutes: 12 },
  { ref: 'industrial_5_3', id: '5-3', subject: '工業簿記', section: '第5問対策', topic: '全部原価計算と直接原価計算', template: 'numeric', maxScore: 12, minutes: 12 },
];

export const externalProblemRefs: ExternalProblemRef[] = cpaTrialSections.map((item) => ({
  id: item.id,
  examType: 'boki2',
  studyMode: 'external_material',
  qualificationName: '日商簿記2級',
  materialName: 'CPAラーニング 日商簿記2級 試験対策編（外部教材参照）',
  subject: item.subject,
  section: item.section,
  topic: item.topic,
  title: item.topic,
  pageMemo: '',
  studyMemo: '',
  answerTemplateId: item.template,
  cpaTrialSectionRef: item.ref,
  maxScore: item.maxScore,
  estimatedMinutes: item.minutes,
  difficulty: '標準',
  sourceType: 'external_material_reference',
  referenceScope: 'cpa_trial_section_structure_only',
  copyrightNote: COPYRIGHT_NOTE,
  createdAt: NOW,
  updatedAt: NOW,
}));

export function findTrialSection(ref: CpaTrialSectionRef): SectionDef {
  return cpaTrialSections.find((item) => item.ref === ref) ?? cpaTrialSections[0];
}
