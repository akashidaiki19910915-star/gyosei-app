const VISUAL_AID_ROOT_ID = 'industrial-account-flow-visual-aid-root';
const VISUAL_AID_STYLE_ID = 'industrial-account-flow-visual-aid-style';

const VISUAL_AID_STYLE = `
  .visual-aid-panel {
    border: 1px solid #cfe5d6;
    background: #f6fbf7;
    border-radius: 14px;
    margin: 16px 0;
    padding: 14px;
  }
  .visual-aid-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 12px;
  }
  .visual-aid-header h2 { margin: 2px 0 0; }
  .visual-aid-body { margin-top: 12px; }
  .visual-aid-description { margin: 0 0 12px; color: #31543d; line-height: 1.7; }
  .visual-aid-diagram {
    background: #fff;
    border: 1px solid #d9eadf;
    border-radius: 12px;
    padding: 14px;
    overflow-x: auto;
  }
  .flow-row {
    display: flex;
    align-items: center;
    gap: 8px;
    margin: 8px 0;
    min-width: 620px;
  }
  .flow-node {
    min-width: 96px;
    padding: 10px 12px;
    border: 1px solid #8fc39e;
    border-radius: 10px;
    background: #eef8f0;
    text-align: center;
    font-weight: 700;
    color: #173a22;
  }
  .flow-main-node { background: #dff2e4; border-color: #4c9b61; }
  .flow-arrow { color: #2f7d45; font-weight: 800; }
  .flow-edge-list {
    display: flex;
    flex-wrap: wrap;
    gap: 8px;
    margin-top: 12px;
  }
  .flow-edge-list span {
    background: #f1f6f3;
    border: 1px solid #d8e8dd;
    border-radius: 999px;
    padding: 5px 9px;
    color: #31543d;
    font-size: 0.9rem;
  }
  @media (max-width: 760px) {
    .visual-aid-header { align-items: flex-start; flex-direction: column; }
    .flow-row {
      align-items: stretch;
      flex-direction: column;
      min-width: 0;
    }
    .flow-arrow { transform: rotate(90deg); align-self: center; }
    .flow-node { width: 100%; box-sizing: border-box; }
  }
`;

const INDUSTRIAL_ACCOUNT_FLOW_HTML = `
  <section class="panel visual-aid-panel" aria-label="図表で確認">
    <div class="visual-aid-header">
      <div>
        <p class="eyebrow">図表で確認</p>
        <h2>勘定連絡図</h2>
      </div>
      <button type="button" class="secondary" data-visual-aid-toggle>図表で確認</button>
    </div>
    <div class="visual-aid-body" data-visual-aid-body hidden>
      <p class="visual-aid-description">材料・賃金・経費が製造活動を通じて仕掛品、製品、売上原価へ流れる関係を確認する図です。</p>
      <div class="visual-aid-diagram industrial-account-flow" aria-label="勘定連絡図">
        <div class="flow-row flow-input-row">
          <div class="flow-node">材料</div><div class="flow-arrow" aria-hidden="true">→</div><div class="flow-node flow-main-node">仕掛品</div><div class="flow-arrow" aria-hidden="true">→</div><div class="flow-node">製品</div><div class="flow-arrow" aria-hidden="true">→</div><div class="flow-node">売上原価</div>
        </div>
        <div class="flow-row flow-support-row">
          <div class="flow-node">賃金</div><div class="flow-arrow" aria-hidden="true">→</div><div class="flow-node flow-main-node">仕掛品</div>
        </div>
        <div class="flow-row flow-support-row">
          <div class="flow-node">経費</div><div class="flow-arrow" aria-hidden="true">→</div><div class="flow-node">製造間接費</div><div class="flow-arrow" aria-hidden="true">→</div><div class="flow-node flow-main-node">仕掛品</div>
        </div>
        <div class="flow-edge-list" aria-label="表示している流れ">
          <span>材料 → 仕掛品</span><span>賃金 → 仕掛品</span><span>経費 → 製造間接費</span><span>製造間接費 → 仕掛品</span><span>仕掛品 → 製品</span><span>製品 → 売上原価</span>
        </div>
      </div>
    </div>
  </section>
`;

function ensureVisualAidStyle() {
  if (document.getElementById(VISUAL_AID_STYLE_ID)) return;
  const style = document.createElement('style');
  style.id = VISUAL_AID_STYLE_ID;
  style.textContent = VISUAL_AID_STYLE;
  document.head.appendChild(style);
}

function currentQuestionHasIndustrial41(): boolean {
  const questionCard = document.querySelector('.question-reference-card');
  const text = questionCard?.textContent ?? '';
  return text.includes('教材なしモード') && text.includes('industrial_4_1');
}

function removeVisualAidIfNeeded() {
  const root = document.getElementById(VISUAL_AID_ROOT_ID);
  if (root && !currentQuestionHasIndustrial41()) root.remove();
}

function ensureVisualAid() {
  if (typeof document === 'undefined') return;
  ensureVisualAidStyle();
  const answerArea = document.querySelector('.focused-answer-area.answer-only-main');
  if (!answerArea) return;
  if (!currentQuestionHasIndustrial41()) {
    removeVisualAidIfNeeded();
    return;
  }
  if (document.getElementById(VISUAL_AID_ROOT_ID)) return;

  const root = document.createElement('div');
  root.id = VISUAL_AID_ROOT_ID;
  root.innerHTML = INDUSTRIAL_ACCOUNT_FLOW_HTML;
  answerArea.parentElement?.insertBefore(root, answerArea);

  const button = root.querySelector<HTMLButtonElement>('[data-visual-aid-toggle]');
  const body = root.querySelector<HTMLElement>('[data-visual-aid-body]');
  button?.addEventListener('click', () => {
    if (!body || !button) return;
    const nextOpen = body.hasAttribute('hidden');
    if (nextOpen) {
      body.removeAttribute('hidden');
      button.textContent = '図表を閉じる';
    } else {
      body.setAttribute('hidden', '');
      button.textContent = '図表で確認';
    }
  });
}

export function installVisualAidBridge() {
  if (typeof window === 'undefined' || typeof document === 'undefined') return;
  if ((window as Window & { __boki2VisualAidBridgeInstalled?: boolean }).__boki2VisualAidBridgeInstalled) return;
  (window as Window & { __boki2VisualAidBridgeInstalled?: boolean }).__boki2VisualAidBridgeInstalled = true;

  const observer = new MutationObserver(() => ensureVisualAid());
  observer.observe(document.body, { childList: true, subtree: true });
  window.setTimeout(() => ensureVisualAid(), 0);
}
