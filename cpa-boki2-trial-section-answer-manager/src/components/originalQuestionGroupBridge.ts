import type { CpaTrialSectionRef, QuestionItem, ReviewSchedule, AnswerAttempt } from '../originalStudyTypes';

const DB_NAME = 'nissho-boki2-original-study-app';
const DB_VERSION = 2;
const BRIDGE_ID = 'boki2-original-question-group-selector';
const STYLE_ID = 'boki2-original-question-group-selector-style';
const OPEN_GROUP_KEY = 'boki2-original-question-open-group-ref';
const SESSION_KEY = 'boki2-original-question-continuous-session';

type BridgeWindow = Window & { __boki2OriginalQuestionGroups?: QuestionItem[]; __boki2OriginalQuestionGroupBridgeInstalled?: boolean };
type StoredAttempt = Pick<AnswerAttempt, 'studyMode' | 'questionId'>;
type StoredReview = Pick<ReviewSchedule, 'studyMode' | 'questionId' | 'nextReviewDate' | 'latestRank' | 'isWeak'>;
type QuestionStatus = '未学習' | '学習済み' | '復習対象';
type GroupStats = Record<string, { studied: number; due: number; statuses: Record<string, QuestionStatus> }>;
type ContinuousSession = { groupRef: string; questionIds: string[]; currentIndex: number };

type QuestionGroup = {
  ref: CpaTrialSectionRef;
  label: string;
  subject: string;
  section: string;
  topic: string;
  questions: QuestionItem[];
};

function sectionNumber(ref: string): string {
  const match = ref.match(/_(\d+)_(\d+)$/);
  return match ? `${match[1]}-${match[2]}` : ref;
}

function groupLabel(question: QuestionItem): string {
  return `${sectionNumber(question.cpaTrialSectionRef)} ${question.topic}`;
}

function groupQuestions(questions: QuestionItem[]): QuestionGroup[] {
  const map = new Map<CpaTrialSectionRef, QuestionGroup>();
  questions
    .filter((question) => question.verificationStatus === 'approved')
    .forEach((question) => {
      const ref = question.cpaTrialSectionRef;
      const current = map.get(ref);
      if (current) {
        current.questions.push(question);
        return;
      }
      map.set(ref, {
        ref,
        label: groupLabel(question),
        subject: question.subject,
        section: question.section,
        topic: question.topic,
        questions: [question],
      });
    });
  return Array.from(map.values()).sort((a, b) => sectionNumber(a.ref).localeCompare(sectionNumber(b.ref), 'ja', { numeric: true }));
}

function todayString(): string {
  return new Date().toISOString().slice(0, 10);
}

function readSession(): ContinuousSession | null {
  try {
    const raw = sessionStorage.getItem(SESSION_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw) as ContinuousSession;
    if (!parsed || !Array.isArray(parsed.questionIds) || !parsed.questionIds.length) return null;
    return parsed;
  } catch {
    return null;
  }
}

function writeSession(session: ContinuousSession | null): void {
  if (!session) {
    sessionStorage.removeItem(SESSION_KEY);
    return;
  }
  sessionStorage.setItem(SESSION_KEY, JSON.stringify(session));
}

function openDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function getAllFromStore<T>(storeName: string): Promise<T[]> {
  if (!('indexedDB' in window)) return [];
  try {
    const db = await openDb();
    if (!db.objectStoreNames.contains(storeName)) {
      db.close();
      return [];
    }
    return await new Promise<T[]>((resolve, reject) => {
      const tx = db.transaction(storeName, 'readonly');
      const request = tx.objectStore(storeName).getAll();
      request.onsuccess = () => resolve((request.result ?? []) as T[]);
      request.onerror = () => reject(request.error);
      tx.oncomplete = () => db.close();
      tx.onerror = () => { db.close(); reject(tx.error); };
    });
  } catch (error) {
    console.warn('original question group bridge storage read skipped', error);
    return [];
  }
}

async function loadStats(groups: QuestionGroup[]): Promise<GroupStats> {
  const [attempts, reviews] = await Promise.all([
    getAllFromStore<StoredAttempt>('answerAttempts'),
    getAllFromStore<StoredReview>('reviewSchedules'),
  ]);
  const today = todayString();
  const attemptsByQuestion = new Set(attempts.filter((attempt) => attempt.studyMode === 'built_in_question').map((attempt) => attempt.questionId));
  const reviewsByQuestion = new Map(reviews.filter((review) => review.studyMode === 'built_in_question').map((review) => [review.questionId, review]));
  const stats: GroupStats = {};
  groups.forEach((group) => {
    const statuses: Record<string, QuestionStatus> = {};
    let studied = 0;
    let due = 0;
    group.questions.forEach((question) => {
      const review = reviewsByQuestion.get(question.id);
      const isDue = Boolean(review && (review.isWeak || review.nextReviewDate <= today || review.latestRank === 'B' || review.latestRank === 'C'));
      if (isDue) {
        statuses[question.id] = '復習対象';
        due += 1;
        studied += 1;
      } else if (attemptsByQuestion.has(question.id)) {
        statuses[question.id] = '学習済み';
        studied += 1;
      } else {
        statuses[question.id] = '未学習';
      }
    });
    stats[group.ref] = { studied, due, statuses };
  });
  return stats;
}

function findBuiltInSelect(): HTMLSelectElement | null {
  const selects = Array.from(document.querySelectorAll<HTMLSelectElement>('section.answer-summary-card select'));
  return selects.find((select) => select.closest('label')?.textContent?.includes('approved オリジナル短問')) ?? null;
}

function setNativeSelectValue(select: HTMLSelectElement, value: string): void {
  const setter = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value')?.set;
  if (setter) setter.call(select, value);
  else select.value = value;
  select.dispatchEvent(new Event('input', { bubbles: true }));
  select.dispatchEvent(new Event('change', { bubbles: true }));
}

function selectQuestionById(id: string): void {
  const select = findBuiltInSelect();
  if (!select) return;
  setNativeSelectValue(select, id);
  const solveButton = Array.from(document.querySelectorAll<HTMLButtonElement>('.mode-nav button')).find((button) => button.textContent?.trim() === '今すぐ解く');
  solveButton?.click();
}

function questionRank(status: QuestionStatus): number {
  if (status === '未学習') return 0;
  if (status === '復習対象') return 1;
  return 2;
}

function startContinuous(group: QuestionGroup, stats: GroupStats): void {
  const statuses = stats[group.ref]?.statuses ?? {};
  const ordered = [...group.questions].sort((a, b) => questionRank(statuses[a.id] ?? '未学習') - questionRank(statuses[b.id] ?? '未学習'));
  const questionIds = ordered.map((question) => question.id);
  writeSession({ groupRef: group.ref, questionIds, currentIndex: 0 });
  sessionStorage.setItem(OPEN_GROUP_KEY, group.ref);
  selectQuestionById(questionIds[0]);
}

function advanceContinuousAfterSave(): void {
  const session = readSession();
  if (!session) return;
  const nextIndex = session.currentIndex + 1;
  if (nextIndex >= session.questionIds.length) {
    writeSession(null);
    return;
  }
  const nextSession = { ...session, currentIndex: nextIndex };
  writeSession(nextSession);
  setTimeout(() => selectQuestionById(nextSession.questionIds[nextIndex]), 150);
}

function injectStyle(): void {
  if (document.getElementById(STYLE_ID)) return;
  const style = document.createElement('style');
  style.id = STYLE_ID;
  style.textContent = `
    .original-question-group-selector { display: grid; gap: 12px; }
    .original-question-group-card { border: 1px solid #d8dee9; border-radius: 14px; padding: 14px; background: #fff; box-shadow: 0 1px 4px rgba(0,0,0,.04); }
    .original-question-group-card.is-current { border-color: #2563eb; box-shadow: 0 0 0 2px rgba(37,99,235,.12); }
    .original-question-group-header { display: flex; justify-content: space-between; gap: 12px; align-items: flex-start; flex-wrap: wrap; }
    .original-question-group-title { display: grid; gap: 4px; }
    .original-question-group-title button { border: 0; background: transparent; padding: 0; text-align: left; font: inherit; cursor: pointer; color: #111827; }
    .original-question-group-title strong { font-size: 1.05rem; }
    .original-question-group-title span, .original-question-card-meta { color: #64748b; font-size: .9rem; }
    .original-question-group-actions { display: flex; flex-wrap: wrap; gap: 8px; }
    .original-question-list { display: grid; gap: 8px; margin-top: 12px; padding-top: 12px; border-top: 1px solid #e5e7eb; }
    .original-question-row { display: grid; grid-template-columns: minmax(0, 1fr) auto; gap: 10px; align-items: center; border: 1px solid #edf2f7; border-radius: 10px; padding: 10px; background: #f8fafc; }
    .original-question-row-title { display: grid; gap: 3px; }
    .original-question-status { display: inline-flex; align-items: center; width: fit-content; border-radius: 999px; background: #e0f2fe; color: #0369a1; padding: 2px 8px; font-size: .8rem; }
    .original-question-continuous-banner { display: flex; justify-content: space-between; align-items: center; gap: 10px; border-radius: 12px; padding: 10px 12px; background: #fff7ed; border: 1px solid #fed7aa; }
    .original-question-group-selector button { cursor: pointer; }
    @media (max-width: 720px) { .original-question-row { grid-template-columns: 1fr; } .original-question-group-header { display: grid; } }
  `;
  document.head.appendChild(style);
}

function renderGroupSelector(root: HTMLElement, select: HTMLSelectElement, groups: QuestionGroup[], stats: GroupStats): void {
  const currentId = select.value;
  const currentGroup = groups.find((group) => group.questions.some((question) => question.id === currentId));
  const openRef = sessionStorage.getItem(OPEN_GROUP_KEY) as CpaTrialSectionRef | null;
  const session = readSession();
  const sessionGroup = session ? groups.find((group) => group.ref === session.groupRef) : null;
  const sessionIndex = session && session.questionIds.includes(currentId) ? session.questionIds.indexOf(currentId) : session?.currentIndex ?? -1;

  root.innerHTML = `
    <section class="panel original-question-group-selector" aria-label="教材なしモードの論点グループ選択">
      <div>
        <p class="eyebrow">教材なしモード</p>
        <h2>論点グループから選ぶ</h2>
        <p class="readable-question-text">問題数が増えても選びやすいよう、1問ずつではなく論点グループ単位で表示しています。</p>
      </div>
      ${session && sessionGroup ? `<div class="original-question-continuous-banner"><strong>${sessionGroup.label}　${Math.max(1, sessionIndex + 1)} / ${session.questionIds.length}問目</strong><button type="button" data-end-continuous="true">連続演習を終了</button></div>` : ''}
      ${groups.map((group) => {
        const groupStats = stats[group.ref] ?? { studied: 0, due: 0, statuses: {} };
        const isOpen = openRef === group.ref;
        const isCurrent = currentGroup?.ref === group.ref;
        const unstudied = Math.max(0, group.questions.length - groupStats.studied);
        return `<article class="original-question-group-card ${isCurrent ? 'is-current' : ''}" data-group-ref="${group.ref}">
          <div class="original-question-group-header">
            <div class="original-question-group-title">
              <button type="button" data-toggle-group="${group.ref}"><strong>${group.label}</strong></button>
              <span>${group.subject} / ${group.section}</span>
              <span>問題数：${group.questions.length}問 / 学習済み：${groupStats.studied}問 / 未学習：${unstudied}問 / 復習対象：${groupStats.due}問</span>
            </div>
            <div class="original-question-group-actions">
              <button type="button" data-open-list="${group.ref}">1問ずつ選んで解く</button>
              <button type="button" data-start-continuous="${group.ref}">連続で解く</button>
            </div>
          </div>
          ${isOpen ? `<div class="original-question-list">
            ${group.questions.map((question) => {
              const status = groupStats.statuses[question.id] ?? '未学習';
              return `<div class="original-question-row">
                <div class="original-question-row-title">
                  <strong>${question.title}</strong>
                  <span class="original-question-card-meta">${question.difficulty ?? '標準'} / ${question.maxScore}点 / ${question.estimatedMinutes ?? '-'}分</span>
                  <span class="original-question-status">${status}</span>
                </div>
                <button type="button" data-select-question="${question.id}">解く</button>
              </div>`;
            }).join('')}
          </div>` : ''}
        </article>`;
      }).join('')}
    </section>
  `;

  root.querySelectorAll<HTMLButtonElement>('[data-toggle-group], [data-open-list]').forEach((button) => {
    button.addEventListener('click', () => {
      const ref = button.dataset.toggleGroup || button.dataset.openList;
      if (!ref) return;
      sessionStorage.setItem(OPEN_GROUP_KEY, ref);
      scheduleRender();
    });
  });
  root.querySelectorAll<HTMLButtonElement>('[data-select-question]').forEach((button) => {
    button.addEventListener('click', () => {
      const id = button.dataset.selectQuestion;
      if (!id) return;
      writeSession(null);
      selectQuestionById(id);
      scheduleRender();
    });
  });
  root.querySelectorAll<HTMLButtonElement>('[data-start-continuous]').forEach((button) => {
    button.addEventListener('click', () => {
      const ref = button.dataset.startContinuous as CpaTrialSectionRef | undefined;
      const group = groups.find((candidate) => candidate.ref === ref);
      if (!group) return;
      startContinuous(group, stats);
      scheduleRender();
    });
  });
  root.querySelector<HTMLButtonElement>('[data-end-continuous]')?.addEventListener('click', () => {
    writeSession(null);
    scheduleRender();
  });
}

let renderScheduled = false;
function scheduleRender(): void {
  if (renderScheduled) return;
  renderScheduled = true;
  requestAnimationFrame(() => {
    renderScheduled = false;
    void renderBridge();
  });
}

async function renderBridge(): Promise<void> {
  const win = window as BridgeWindow;
  const questions = win.__boki2OriginalQuestionGroups ?? [];
  if (!questions.length) return;
  const select = findBuiltInSelect();
  const sourcePanel = select?.closest<HTMLElement>('section.answer-summary-card');
  if (!select || !sourcePanel) {
    document.getElementById(BRIDGE_ID)?.remove();
    return;
  }
  sourcePanel.style.display = 'none';
  let root = document.getElementById(BRIDGE_ID);
  if (!root) {
    root = document.createElement('div');
    root.id = BRIDGE_ID;
    sourcePanel.insertAdjacentElement('afterend', root);
  }
  const groups = groupQuestions(questions);
  const stats = await loadStats(groups);
  renderGroupSelector(root, select, groups, stats);
}

function handleSaveClick(event: MouseEvent): void {
  const target = event.target;
  if (!(target instanceof HTMLElement)) return;
  const button = target.closest('button');
  if (!button || button.textContent?.trim() !== '保存して次へ') return;
  const session = readSession();
  if (!session) return;
  setTimeout(() => advanceContinuousAfterSave(), 300);
}

export function installOriginalQuestionGroupSelectorBridge(questions: QuestionItem[]): void {
  if (typeof window === 'undefined' || typeof document === 'undefined') return;
  const win = window as BridgeWindow;
  win.__boki2OriginalQuestionGroups = questions;
  injectStyle();
  if (win.__boki2OriginalQuestionGroupBridgeInstalled) {
    scheduleRender();
    return;
  }
  win.__boki2OriginalQuestionGroupBridgeInstalled = true;
  const observer = new MutationObserver((mutations) => {
    const root = document.getElementById(BRIDGE_ID);
    if (root && mutations.every((mutation) => root.contains(mutation.target as Node))) return;
    scheduleRender();
  });
  observer.observe(document.body, { childList: true, subtree: true });
  document.addEventListener('click', handleSaveClick, true);
  scheduleRender();
}
