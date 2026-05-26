import './visualAids.css';

const VISUAL_AID_ROOT_ID = 'industrial-account-flow-visual-aid-root';

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

function currentQuestionHasIndustrial41(): boolean {
  const questionCard = document.querySelector('.question-reference-card');
  return Boolean(questionCard?.textContent?.includes('industrial_4_1'));
}

function removeVisualAidIfNeeded() {
  const root = document.getElementById(VISUAL_AID_ROOT_ID);
  if (root && !currentQuestionHasIndustrial41()) root.remove();
}

function ensureVisualAid() {
  if (typeof document === 'undefined') return;
  const questionCard = document.querySelector('.question-reference-card');
  const answerArea = document.querySelector('.focused-answer-area.answer-only-main');
  if (!questionCard || !answerArea) return;
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
