import type { StorageSafetyInfo } from '../types';
import { formatBytes, requestPersistentStorage } from '../utils/storageSafety';

interface Props {
  info: StorageSafetyInfo | null;
  onBackup: () => void;
  onRestoreFile: (file: File | undefined) => void;
  onRefresh: () => void;
  onMessage: (message: string) => void;
}

export function BackupSafetyPanel({ info, onBackup, onRestoreFile, onRefresh, onMessage }: Props) {
  const enablePersistent = async () => {
    const result = await requestPersistentStorage();
    onMessage(`永続ストレージ要求結果：${result}`);
    onRefresh();
  };

  return (
    <details className="panel management-panel" open>
      <summary>保存安全性チェック</summary>
      <div className="panel-body">
        <p className="warning-text">このアプリの学習履歴はブラウザ内に保存されています。ブラウザデータ削除・端末変更・シークレットモード利用により消える可能性があります。定期的にJSONバックアップを取得してください。</p>
        {info ? (
          <>
            <div className="metric-grid">
              <div><span>保存方式</span><strong>{info.saveMethod}</strong></div>
              <div><span>履歴件数</span><strong>{info.historyCount}</strong></div>
              <div><span>保存済み答案</span><strong>{info.answerCount}</strong></div>
              <div><span>最終バックアップ</span><strong>{info.lastBackupAt || '未取得'}</strong></div>
              <div><span>最終復元</span><strong>{info.lastRestoreAt || '-'}</strong></div>
              <div><span>バックアップ回数</span><strong>{info.backupCount}</strong></div>
              <div><span>推定使用量</span><strong>{formatBytes(info.estimatedUsage)}</strong></div>
              <div><span>推定上限</span><strong>{formatBytes(info.estimatedQuota)}</strong></div>
              <div><span>永続ストレージ</span><strong>{info.persistentStatus}</strong></div>
            </div>
            {info.warnings.length ? (
              <ul className="warning-list">{info.warnings.map((warning) => <li key={warning}>{warning}</li>)}</ul>
            ) : <p className="ok-text">現時点で重大なバックアップ警告はありません。</p>}
          </>
        ) : <p>保存状態を確認中です。</p>}
        <div className="button-row">
          <button onClick={onBackup}>JSONバックアップ</button>
          <label className="import-label">JSON復元<input type="file" accept="application/json" onChange={(event) => onRestoreFile(event.target.files?.[0])} /></label>
          <button className="secondary" onClick={enablePersistent}>永続ストレージを有効化</button>
          <button className="secondary" onClick={onRefresh}>保存状態を再確認</button>
        </div>
      </div>
    </details>
  );
}
