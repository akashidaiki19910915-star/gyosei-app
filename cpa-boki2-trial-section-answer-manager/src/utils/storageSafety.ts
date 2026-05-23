import type { BackupMetadata, StorageSafetyInfo } from '../types';

function daysSince(value: string): number | null {
  if (!value) return null;
  const time = new Date(value).getTime();
  if (Number.isNaN(time)) return null;
  return Math.floor((Date.now() - time) / 86400000);
}

export async function getStorageSafetyInfo(historyCount: number, answerCount: number, metadata: BackupMetadata | null): Promise<StorageSafetyInfo> {
  const storageSupported = typeof navigator !== 'undefined' && 'storage' in navigator;
  let estimatedUsage: number | null = null;
  let estimatedQuota: number | null = null;
  let persistentStatus: StorageSafetyInfo['persistentStatus'] = storageSupported ? '確認不可' : '非対応';

  try {
    if (storageSupported && navigator.storage.estimate) {
      const estimate = await navigator.storage.estimate();
      estimatedUsage = estimate.usage ?? null;
      estimatedQuota = estimate.quota ?? null;
    }
    if (storageSupported && navigator.storage.persisted) {
      persistentStatus = (await navigator.storage.persisted()) ? '許可済み' : '未許可';
    }
  } catch {
    persistentStatus = '確認不可';
  }

  const lastBackupAt = metadata?.lastBackupAt ?? '';
  const backupAge = daysSince(lastBackupAt);
  const warnings: string[] = [];
  if (!lastBackupAt) warnings.push('JSONバックアップ未取得です。ブラウザデータ削除・端末変更に備えてバックアップしてください。');
  if (backupAge !== null && backupAge >= 7) warnings.push('最終JSONバックアップから7日以上経過しています。');
  if (historyCount >= 50 && !lastBackupAt) warnings.push('履歴件数が50件以上ありますが、JSONバックアップが未取得です。');
  if (!storageSupported) warnings.push('このブラウザはStorageManager APIに非対応です。保存状態の詳細確認ができません。');
  if (persistentStatus === '未許可') warnings.push('永続ストレージが許可されていません。ブラウザ判断で保存データが削除される可能性があります。');

  return {
    saveMethod: 'IndexedDB',
    historyCount,
    answerCount,
    lastBackupAt,
    lastRestoreAt: metadata?.lastRestoreAt ?? '',
    backupCount: metadata?.backupCount ?? 0,
    lastBackupHistoryCount: metadata?.lastBackupHistoryCount ?? 0,
    lastBackupAnswerCount: metadata?.lastBackupAnswerCount ?? 0,
    estimatedUsage,
    estimatedQuota,
    persistentStatus,
    storageManagerSupported: storageSupported,
    warnings,
  };
}

export async function requestPersistentStorage(): Promise<'許可済み' | '未許可' | '非対応'> {
  if (typeof navigator === 'undefined' || !('storage' in navigator) || !navigator.storage.persist) return '非対応';
  try {
    return (await navigator.storage.persist()) ? '許可済み' : '未許可';
  } catch {
    return '未許可';
  }
}

export function formatBytes(value: number | null): string {
  if (value === null) return '-';
  if (value < 1024) return `${value} B`;
  if (value < 1024 * 1024) return `${(value / 1024).toFixed(1)} KB`;
  if (value < 1024 * 1024 * 1024) return `${(value / 1024 / 1024).toFixed(1)} MB`;
  return `${(value / 1024 / 1024 / 1024).toFixed(2)} GB`;
}
