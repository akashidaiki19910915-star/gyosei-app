import { useEffect, useRef, useState } from 'react';

const EXAM_SECONDS = 90 * 60;

export function Timer() {
  const [seconds, setSeconds] = useState(EXAM_SECONDS);
  const [status, setStatus] = useState('開始前');
  const intervalRef = useRef<number | null>(null);

  useEffect(() => {
    if (seconds <= 0 && intervalRef.current !== null) {
      window.clearInterval(intervalRef.current);
      intervalRef.current = null;
      setStatus('試験終了');
    } else if (seconds <= 600 && intervalRef.current !== null) {
      setStatus('残り10分');
    }
  }, [seconds]);

  const start = () => {
    if (intervalRef.current !== null) return;
    if (seconds <= 0) setSeconds(EXAM_SECONDS);
    setStatus('試験中');
    intervalRef.current = window.setInterval(() => setSeconds((value) => Math.max(value - 1, 0)), 1000);
  };

  const pause = () => {
    if (intervalRef.current !== null) window.clearInterval(intervalRef.current);
    intervalRef.current = null;
    setStatus('一時停止中');
  };

  const reset = () => {
    if (intervalRef.current !== null) window.clearInterval(intervalRef.current);
    intervalRef.current = null;
    setSeconds(EXAM_SECONDS);
    setStatus('開始前');
  };

  const minutes = Math.floor(seconds / 60).toString().padStart(2, '0');
  const rest = (seconds % 60).toString().padStart(2, '0');

  return (
    <section className="timer-card">
      <div className="timer-display">{minutes}:{rest}</div>
      <div className={`timer-status ${status === '残り10分' ? 'warning' : ''}`}>{status}</div>
      <div className="button-row">
        <button onClick={start}>開始</button>
        <button onClick={pause}>一時停止</button>
        <button onClick={start}>再開</button>
        <button className="secondary" onClick={reset}>リセット</button>
      </div>
    </section>
  );
}
