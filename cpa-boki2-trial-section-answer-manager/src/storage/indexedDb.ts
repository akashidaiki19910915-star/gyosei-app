import type { AnswerState, BackupPayload, HistoryEntry } from '../types';

const DB_NAME = 'cpa-boki2-trial-section-answer-manager';
const DB_VERSION = 1;
const STORES = ['answers', 'histories', 'settings', 'templates', 'backups'] as const;

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

export async function exportAllData(): Promise<BackupPayload> {
  return {
    exportedAt: new Date().toISOString(),
    appName: 'CPA日商簿記2級 試験対策編 解答・復習管理アプリ',
    version: '1.0.0',
    answers: await getAllAnswers(),
    histories: await getHistories(),
    settings: await tx<Record<string, unknown>[]>('settings', 'readonly', (store) => store.getAll()),
    templates: await tx<Record<string, unknown>[]>('templates', 'readonly', (store) => store.getAll()),
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
    transaction.oncomplete = () => {
      db.close();
      resolve();
    };
    transaction.onerror = () => {
      db.close();
      reject(transaction.error);
    };
  });
}
