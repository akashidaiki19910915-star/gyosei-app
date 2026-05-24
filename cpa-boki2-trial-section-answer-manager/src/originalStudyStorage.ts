import type { AnswerAttempt, ExternalProblemRef, QuestionItem, ReviewSchedule, StudyBackupPayload } from './originalStudyTypes';

const DB_NAME = 'nissho-boki2-original-study-app';
const DB_VERSION = 2;
const STORES = ['questions', 'externalProblemRefs', 'answerAttempts', 'reviewSchedules', 'settings', 'userProblems', 'histories'] as const;
type StoreName = (typeof STORES)[number];

function openDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => {
      const db = request.result;
      STORES.forEach((name) => {
        if (!db.objectStoreNames.contains(name)) db.createObjectStore(name, { keyPath: 'id' });
      });
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function tx<T>(storeName: StoreName, mode: IDBTransactionMode, run: (store: IDBObjectStore) => IDBRequest<T>): Promise<T> {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(storeName, mode);
    const request = run(transaction.objectStore(storeName));
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
    transaction.oncomplete = () => db.close();
    transaction.onerror = () => { db.close(); reject(transaction.error); };
  });
}

export async function saveQuestion(question: QuestionItem): Promise<void> {
  const target = question.sourceType === 'user_created' ? 'userProblems' : 'questions';
  await tx(target, 'readwrite', (store) => store.put(question));
}

export async function saveQuestions(questions: QuestionItem[]): Promise<void> {
  const db = await openDb();
  await new Promise<void>((resolve, reject) => {
    const transaction = db.transaction(['questions', 'userProblems'], 'readwrite');
    questions.forEach((question) => transaction.objectStore(question.sourceType === 'user_created' ? 'userProblems' : 'questions').put(question));
    transaction.oncomplete = () => { db.close(); resolve(); };
    transaction.onerror = () => { db.close(); reject(transaction.error); };
  });
}

export async function saveExternalProblemRef(ref: ExternalProblemRef): Promise<void> {
  await tx('externalProblemRefs', 'readwrite', (store) => store.put(ref));
}

export async function saveExternalProblemRefs(refs: ExternalProblemRef[]): Promise<void> {
  const db = await openDb();
  await new Promise<void>((resolve, reject) => {
    const transaction = db.transaction('externalProblemRefs', 'readwrite');
    const store = transaction.objectStore('externalProblemRefs');
    refs.forEach((ref) => store.put(ref));
    transaction.oncomplete = () => { db.close(); resolve(); };
    transaction.onerror = () => { db.close(); reject(transaction.error); };
  });
}

export async function getStoredExternalProblemRefs(): Promise<ExternalProblemRef[]> {
  const rows = await tx<ExternalProblemRef[]>('externalProblemRefs', 'readonly', (store) => store.getAll());
  return rows.sort((a, b) => a.id.localeCompare(b.id, 'ja'));
}

export async function getStoredQuestions(): Promise<QuestionItem[]> {
  const [questions, userProblems] = await Promise.all([
    tx<QuestionItem[]>('questions', 'readonly', (store) => store.getAll()),
    tx<QuestionItem[]>('userProblems', 'readonly', (store) => store.getAll()),
  ]);
  return [...questions, ...userProblems].sort((a, b) => a.id.localeCompare(b.id, 'ja'));
}

export async function saveAttempt(attempt: AnswerAttempt): Promise<void> {
  await tx('answerAttempts', 'readwrite', (store) => store.put(attempt));
  await tx('histories', 'readwrite', (store) => store.put(attempt));
}

export async function getAttempts(): Promise<AnswerAttempt[]> {
  const rows = await tx<AnswerAttempt[]>('answerAttempts', 'readonly', (store) => store.getAll());
  return rows.sort((a, b) => b.answeredAt.localeCompare(a.answeredAt));
}

export async function saveReviewSchedule(schedule: ReviewSchedule): Promise<void> {
  await tx('reviewSchedules', 'readwrite', (store) => store.put(schedule));
}

export async function getReviewSchedules(): Promise<ReviewSchedule[]> {
  const rows = await tx<ReviewSchedule[]>('reviewSchedules', 'readonly', (store) => store.getAll());
  return rows.sort((a, b) => a.nextReviewDate.localeCompare(b.nextReviewDate));
}

export async function exportStudyData(appQuestions: QuestionItem[], appExternalRefs: ExternalProblemRef[]): Promise<StudyBackupPayload> {
  const [storedQuestions, storedExternalRefs, answerAttempts, reviewSchedules] = await Promise.all([getStoredQuestions(), getStoredExternalProblemRefs(), getAttempts(), getReviewSchedules()]);
  const mergedQuestions = new Map<string, QuestionItem>();
  appQuestions.forEach((question) => mergedQuestions.set(question.id, question));
  storedQuestions.forEach((question) => mergedQuestions.set(question.id, question));
  const mergedRefs = new Map<string, ExternalProblemRef>();
  appExternalRefs.forEach((ref) => mergedRefs.set(ref.id, ref));
  storedExternalRefs.forEach((ref) => mergedRefs.set(ref.id, ref));
  return {
    exportedAt: new Date().toISOString(),
    appName: '資格横断型 問題演習・採点・復習管理アプリ',
    version: '2.0.0',
    questions: Array.from(mergedQuestions.values()).sort((a, b) => a.id.localeCompare(b.id, 'ja')),
    externalProblemRefs: Array.from(mergedRefs.values()).sort((a, b) => a.id.localeCompare(b.id, 'ja')),
    answerAttempts,
    reviewSchedules,
  };
}

export async function importStudyData(payload: StudyBackupPayload): Promise<void> {
  if (!payload || !Array.isArray(payload.questions)) throw new Error('invalid payload');
  const db = await openDb();
  await new Promise<void>((resolve, reject) => {
    const transaction = db.transaction(['questions', 'externalProblemRefs', 'userProblems', 'answerAttempts', 'reviewSchedules', 'histories'], 'readwrite');
    payload.questions.forEach((question) => transaction.objectStore(question.sourceType === 'user_created' ? 'userProblems' : 'questions').put(question));
    (payload.externalProblemRefs ?? []).forEach((ref) => transaction.objectStore('externalProblemRefs').put(ref));
    (payload.answerAttempts ?? []).forEach((attempt) => {
      transaction.objectStore('answerAttempts').put(attempt);
      transaction.objectStore('histories').put(attempt);
    });
    (payload.reviewSchedules ?? []).forEach((schedule) => transaction.objectStore('reviewSchedules').put(schedule));
    transaction.oncomplete = () => { db.close(); resolve(); };
    transaction.onerror = () => { db.close(); reject(transaction.error); };
  });
}

export async function storageSummary(): Promise<{ questionCount: number; externalRefCount: number; attemptCount: number; reviewCount: number; usage: number | null; quota: number | null }> {
  const [questions, externalRefs, attempts, reviews] = await Promise.all([getStoredQuestions(), getStoredExternalProblemRefs(), getAttempts(), getReviewSchedules()]);
  const estimate = navigator.storage?.estimate ? await navigator.storage.estimate() : null;
  return { questionCount: questions.length, externalRefCount: externalRefs.length, attemptCount: attempts.length, reviewCount: reviews.length, usage: estimate?.usage ?? null, quota: estimate?.quota ?? null };
}
