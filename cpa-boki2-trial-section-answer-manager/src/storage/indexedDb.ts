import type { AnswerState, BackupMetadata, BackupPayload, ExamSetRecord, HistoryEntry } from '../types';

const DB_NAME = 'cpa-boki2-trial-section-answer-manager';
const DB_VERSION = 2;
const STORES = ['answers', 'histories', 'settings', 'templates', 'backups', 'examSets'] as const;

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
    const store = transaction.objectStore(storeName);
    const request = run(store);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
    transaction.oncomplete = () => db.close();
    transaction.onerror = () => {
      db.close();
      reject(transaction.error);
    };
  });
}

export async function saveAnswer(answer: AnswerState): Promise<void> {
  await tx('answers', 'readwrite', (store) => store.put(answer));
}

export async function getAnswer(id: string): Promise<AnswerState | undefined> {
  return tx('answers', 'readonly', (store) => store.get(id));
}

export async function getAllAnswers(): Promise<AnswerState[]> {
  return tx('answers', 'readonly', (store) => store.getAll());
}

export async function addHistory(history: HistoryEntry): Promise<void> {
  await tx('histories', 'readwrite', (store) => store.put(history));
}

export async function getHistories(): Promise<HistoryEntry[]> {
  const rows = await tx<HistoryEntry[]>('histories', 'readonly', (store) => store.getAll());
  return rows.sort((a, b) => b.savedAt.localeCompare(a.savedAt));
}

export async function deleteHistory(id: string): Promise<void> {
  await tx('histories', 'readwrite', (store) => store.delete(id));
}

export async function clearHistories(): Promise<void> {
  await tx('histories', 'readwrite', (store) => store.clear());
}

export async function saveExamSet(record: ExamSetRecord): Promise<void> {
  await tx('examSets', 'readwrite', (store) => store.put(record));
}

export async function getExamSets(): Promise<ExamSetRecord[]> {
  const rows = await tx<ExamSetRecord[]>('examSets', 'readonly', (store) => store.getAll());
  return rows.sort((a, b) => b.createdAt.localeCompare(a.createdAt));
}

export async function deleteExamSet(id: string): Promise<void> {
  await tx('examSets', 'readwrite', (store) => store.delete(id));
}

export async function getBackupMetadata(): Promise<BackupMetadata | null> {
  const row = await tx<BackupMetadata | undefined>('backups', 'readonly', (store) => store.get('backupMeta'));
  return row ?? null;
}

export async function saveBackupMetadata(meta: BackupMetadata): Promise<void> {
  await tx('backups', 'readwrite', (store) => store.put(meta));
}

export async function recordBackupMade(historyCount: number, answerCount: number): Promise<BackupMetadata> {
  const current = await getBackupMetadata();
  const next: BackupMetadata = {
    id: 'backupMeta',
    lastBackupAt: new Date().toISOString(),
    lastRestoreAt: current?.lastRestoreAt ?? '',
    backupCount: (current?.backupCount ?? 0) + 1,
    lastBackupHistoryCount: historyCount,
    lastBackupAnswerCount: answerCount,
  };
  await saveBackupMetadata(next);
  return next;
}

export async function recordRestoreMade(): Promise<BackupMetadata> {
  const current = await getBackupMetadata();
  const next: BackupMetadata = {
    id: 'backupMeta',
    lastBackupAt: current?.lastBackupAt ?? '',
    lastRestoreAt: new Date().toISOString(),
    backupCount: current?.backupCount ?? 0,
    lastBackupHistoryCount: current?.lastBackupHistoryCount ?? 0,
    lastBackupAnswerCount: current?.lastBackupAnswerCount ?? 0,
  };
  await saveBackupMetadata(next);
  return next;
}

export async function exportAllData(): Promise<BackupPayload> {
  return {
    exportedAt: new Date().toISOString(),
    appName: 'CPA日商簿記2級 試験対策編 解答・復習管理アプリ',
    version: '1.1.0',
    answers: await getAllAnswers(),
    histories: await getHistories(),
    settings: await tx<Record<string, unknown>[]>('settings', 'readonly', (store) => store.getAll()),
    templates: await tx<Record<string, unknown>[]>('templates', 'readonly', (store) => store.getAll()),
    backups: await tx<Record<string, unknown>[]>('backups', 'readonly', (store) => store.getAll()),
    examSets: await getExamSets(),
    backupMetadata: await getBackupMetadata(),
  };
}

export async function importAllData(payload: BackupPayload): Promise<void> {
  const db = await openDb();
  await new Promise<void>((resolve, reject) => {
    const transaction = db.transaction([...STORES], 'readwrite');
    STORES.forEach((name) => transaction.objectStore(name).clear());
    payload.answers?.forEach((item) => transaction.objectStore('answers').put(item));
    payload.histories?.forEach((item) => transaction.objectStore('histories').put(item));
    payload.settings?.forEach((item) => transaction.objectStore('settings').put(item));
    payload.templates?.forEach((item) => transaction.objectStore('templates').put(item));
    payload.backups?.forEach((item) => transaction.objectStore('backups').put(item));
    payload.examSets?.forEach((item) => transaction.objectStore('examSets').put(item));
    if (payload.backupMetadata) transaction.objectStore('backups').put(payload.backupMetadata);
    transaction.oncomplete = () => {
      db.close();
      resolve();
    };
    transaction.onerror = () => {
      db.close();
      reject(transaction.error);
    };
  });
  await recordRestoreMade();
}
