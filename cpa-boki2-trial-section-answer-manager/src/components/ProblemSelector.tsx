import { problemCatalog, sectionLabels, subjectLabels } from '../data/problemCatalog';
import { templateCatalog } from '../data/templateCatalog';
import type { ProblemDefinition, SectionId, Subject, TemplateId } from '../types';

interface Props {
  selectedProblem: ProblemDefinition;
  selectedTemplateId: TemplateId;
  onSelectProblem: (problemId: string) => void;
  onSelectTemplate: (templateId: TemplateId) => void;
}

export function ProblemSelector({ selectedProblem, selectedTemplateId, onSelectProblem, onSelectTemplate }: Props) {
  const subjects = Array.from(new Set(problemCatalog.map((problem) => problem.subject))) as Subject[];
  const sections = Array.from(new Set(problemCatalog.filter((problem) => problem.subject === selectedProblem.subject).map((problem) => problem.sectionId))) as SectionId[];
  const problems = problemCatalog.filter((problem) => problem.subject === selectedProblem.subject && problem.sectionId === selectedProblem.sectionId);

  return (
    <aside className="panel sidebar">
      <h2>問題選択</h2>
      <label>科目</label>
      <select value={selectedProblem.subject} onChange={(event) => {
        const next = problemCatalog.find((problem) => problem.subject === event.target.value) ?? problemCatalog[0];
        onSelectProblem(next.id);
      }}>
        {subjects.map((subject) => <option key={subject} value={subject}>{subjectLabels[subject]}</option>)}
      </select>

      <label>大問対策</label>
      <select value={selectedProblem.sectionId} onChange={(event) => {
        const next = problemCatalog.find((problem) => problem.subject === selectedProblem.subject && problem.sectionId === event.target.value) ?? selectedProblem;
        onSelectProblem(next.id);
      }}>
        {sections.map((section) => <option key={section} value={section}>{sectionLabels[section]}</option>)}
      </select>

      <label>問題ID</label>
      <select value={selectedProblem.id} onChange={(event) => onSelectProblem(event.target.value)}>
        {problems.map((problem) => <option key={problem.id} value={problem.id}>{problem.displayId} {problem.topic}</option>)}
      </select>

      <label>使用テンプレート</label>
      <select value={selectedTemplateId} onChange={(event) => onSelectTemplate(event.target.value as TemplateId)}>
        {templateCatalog.map((template) => <option key={template.id} value={template.id}>{template.name}</option>)}
      </select>

      <div className="notice-small">
        問題文・解答・解説は収録していません。ここでは問題IDと論点名だけを管理します。
      </div>
    </aside>
  );
}
