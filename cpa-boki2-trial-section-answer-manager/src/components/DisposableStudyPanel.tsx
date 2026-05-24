import type { PracticeSession, TimeMode } from '../types';
import { timeModeDescription, timeModes } from '../utils/disposableStudy';
import { formatDateTime } from '../utils/dates';

interface Props {
  activeSession: PracticeSession | null;
  quickHint: string;
  onQuickResume: () => void | Promise<void>;
  onSelectMode: (mode: TimeMode) => void | Promise<void>;
}

export function DisposableStudyPanel({ activeSession, quickHint, onQuickResume, onSelectMode }: Props) {
  return (
    <section className="panel disposable-panel">
      <div className="disposable-hero">
        <div>
          <p className="eyebrow">可処分学習時間最大化MVP</p>
          <h2>30秒以内に演習開始</h2>
          <p>{activeSession ? `未完了セッション：${activeSession.selectedTimeMode} / ${formatDateTime(activeSession.startedAt)}` : quickHint}</p>
        </div>
        <button className="resume-button" onClick={() => { void onQuickResume(); }}>今すぐ再開</button>
      </div>
      <div className="time-mode-panel">
        <h3>空き時間から選ぶ</h3>
        <div className="time-mode-grid">
          {timeModes.map((mode) => (
            <button key={mode} className="time-mode-button" onClick={() => { void onSelectMode(mode); }}>
              <strong>{mode}</strong>
              <span>{timeModeDescription(mode)}</span>
            </button>
          ))}
        </div>
      </div>
      <p className="legal-notice">
        この機能は、正規に取得・購入した教材PDFを、利用者本人の私的学習のために端末内で整理・参照する目的のものです。教材PDF・問題文・解答・解説を第三者と共有したり、GitHub・SNS・クラウド公開領域へアップロードしたりしないでください。DRM等の技術的保護手段を回避して取り込むことはしないでください。
      </p>
    </section>
  );
}
