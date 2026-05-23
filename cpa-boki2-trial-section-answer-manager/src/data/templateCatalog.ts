import type { TemplateDefinition } from '../types';

export const templateCatalog: TemplateDefinition[] = [
  { id: 'journal', name: '仕訳テーブル', initialRows: 30, columns: ['行番号', '設問番号', '会社名・立場', '借方科目', '借方金額', '貸方科目', '貸方金額', 'メモ'] },
  { id: 'genericNumber', name: '汎用数値入力', initialRows: 30, columns: ['行番号', '項目名', '入力値1', '入力値2', '入力値3', 'メモ'] },
  { id: 'equityStatement', name: '株主資本等変動計算書', initialRows: 30, columns: ['行番号', '項目名', '当期首残高', '当期変動額', '当期末残高', 'メモ'] },
  { id: 'consolidation', name: '連結会計', initialRows: 40, columns: ['行番号', '処理区分', '借方科目', '借方金額', '貸方科目', '貸方金額', '連結上の意味・メモ'], optionColumns: { 処理区分: ['資本連結', 'のれん', '非支配株主持分', '当期純利益', '配当', '未実現利益', '債権債務相殺', 'その他'] } },
  { id: 'ledger', name: '勘定記入', initialRows: 30, columns: ['行番号', '勘定科目', '借方摘要', '借方金額', '貸方摘要', '貸方金額', '残高', 'メモ'] },
  { id: 'financialStatements', name: '財務諸表・精算表', initialRows: 40, columns: ['行番号', '表区分', '項目名', '借方・資産・費用', '貸方・負債純資産・収益', '入力値', '小計・合計', 'メモ'], optionColumns: { 表区分: ['決算整理仕訳', '損益計算書', '貸借対照表', '精算表', '本店側処理', '支店側処理', '未達事項', '内部利益', '合併財務諸表', 'その他'] } },
  { id: 'processCosting', name: '工程別総合原価計算', initialRows: 30, columns: ['行番号', '工程', '月初仕掛品', '当月投入', '完成品', '月末仕掛品', '換算量', '単価', '金額', 'メモ'] },
  { id: 'departmentCosting', name: '部門別原価計算', initialRows: 30, columns: ['行番号', '部門名', '配賦基準', '第一次集計', '第二次集計', '補助部門費配賦額', '製造部門費', 'メモ'] },
  { id: 'standardCosting', name: '標準原価計算・差異分析', initialRows: 30, columns: ['行番号', '差異区分', '標準数量・時間', '標準単価・賃率', '実際数量・時間', '実際単価・賃率', '差異額', '有利・不利', 'メモ'], optionColumns: { '有利・不利': ['有利', '不利', '判定保留'] } },
  { id: 'cvp', name: 'CVP分析', initialRows: 30, columns: ['行番号', '項目名', '販売単価', '変動費単価', '貢献利益単価', '固定費', '売上高', '営業利益', '損益分岐点', '安全余裕率', 'メモ'] },
  { id: 'variableFullCosting', name: '全部原価計算・直接原価計算', initialRows: 30, columns: ['行番号', '項目名', '全部原価計算', '直接原価計算', '固定製造原価調整', '差額', 'メモ'] },
  { id: 'freeTable', name: '完全自由テーブル', initialRows: 30, columns: ['列1', '列2', '列3', '列4', '列5', '列6', '列7', '列8', 'メモ', '補助欄'] },
];

export function getTemplateById(id: string): TemplateDefinition {
  const template = templateCatalog.find((item) => item.id === id);
  if (!template) return templateCatalog[0];
  return template;
}
