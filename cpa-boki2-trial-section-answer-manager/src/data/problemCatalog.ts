import type { ProblemDefinition, ProblemKind, SectionId, Subject, TemplateId } from '../types';

type Row = [Subject, SectionId, string, string, string, TemplateId];

const rows: Row[] = [
  ['commercial', 'q1', '第1問対策', '1-1', '商品、収益認識、現金預金', 'journal'],
  ['commercial', 'q1', '第1問対策', '1-2', '債権債務、有価証券', 'journal'],
  ['commercial', 'q1', '第1問対策', '1-3', '有形固定資産、無形資産', 'journal'],
  ['commercial', 'q1', '第1問対策', '1-4', 'リース取引、引当金、株式会社会計', 'journal'],
  ['commercial', 'q1', '第1問対策', '1-5', '外貨建取引、税効果会計、本支店会計、連結会計', 'journal'],
  ['commercial', 'q2', '第2問対策', '2-1', '株主資本等変動計算書', 'equityStatement'],
  ['commercial', 'q2', '第2問対策', '2-2', '連結会計①', 'consolidation'],
  ['commercial', 'q2', '第2問対策', '2-3', '連結会計②', 'consolidation'],
  ['commercial', 'q2', '第2問対策', '2-4', '勘定記入① 商品', 'ledger'],
  ['commercial', 'q2', '第2問対策', '2-5', '勘定記入② 有形固定資産', 'ledger'],
  ['commercial', 'q2', '第2問対策', '2-6', '勘定記入③ 有価証券', 'ledger'],
  ['commercial', 'q3', '第3問対策', '3-1', '決算の総合問題①', 'financialStatements'],
  ['commercial', 'q3', '第3問対策', '3-2', '決算の総合問題②', 'financialStatements'],
  ['commercial', 'q3', '第3問対策', '3-3', '本支店会計の総合問題', 'financialStatements'],
  ['industrial', 'q4', '第4問対策', '4-1', '工業簿記 仕訳問題', 'journal'],
  ['industrial', 'q4', '第4問対策', '4-2', '工程別総合原価計算', 'processCosting'],
  ['industrial', 'q4', '第4問対策', '4-3', '部門別原価計算', 'departmentCosting'],
  ['industrial', 'q4', '第4問対策', '4-4', '財務諸表作成', 'financialStatements'],
  ['industrial', 'q5', '第5問対策', '5-1', '標準原価計算', 'standardCosting'],
  ['industrial', 'q5', '第5問対策', '5-2', 'CVP分析', 'cvp'],
  ['industrial', 'q5', '第5問対策', '5-3', '全部原価計算と直接原価計算', 'variableFullCosting'],
];

function problemKindFromTemplate(templateId: TemplateId): ProblemKind {
  if (templateId === 'journal') return 'journal';
  if (templateId === 'genericNumber') return 'numericInput';
  return 'calculationTable';
}

export const problemCatalog: ProblemDefinition[] = rows.map(([subject, sectionId, sectionLabel, displayId, topic, defaultTemplateId]) => ({
  id: `${subject}-${displayId}`,
  qualificationId: 'nissho_boki2',
  subject,
  subjectId: subject,
  sectionId,
  sectionLabel,
  displayId,
  topic,
  topicId: displayId,
  problemKind: problemKindFromTemplate(defaultTemplateId),
  defaultTemplateId,
}));

export const subjectLabels: Record<Subject, string> = {
  commercial: '商業簿記',
  industrial: '工業簿記',
};

export const sectionLabels: Record<SectionId, string> = {
  q1: '第1問対策',
  q2: '第2問対策',
  q3: '第3問対策',
  q4: '第4問対策',
  q5: '第5問対策',
};

export function getProblemById(problemId: string): ProblemDefinition {
  const problem = problemCatalog.find((item) => item.id === problemId);
  if (!problem) return problemCatalog[0];
  return problem;
}
