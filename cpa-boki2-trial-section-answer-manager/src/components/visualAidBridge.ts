import { createElement } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { AccountFlowEditor } from './AccountFlowEditor';

const VISUAL_AID_ROOT_ID = 'industrial-account-flow-visual-aid-root';
const VISUAL_AID_STYLE_ID = 'industrial-account-flow-visual-aid-style';
const EDITOR_MOUNT_ID = 'industrial-account-flow-editor-root';

type BridgeWindow = Window & {
  __boki2VisualAidBridgeInstalled?: boolean;
  __boki2AccountFlowEditorRoot?: Root | null;
};

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
  .wip-box-title {
    text-align: center;
    font-weight: 800;
    color: #173a22;
    margin-bottom: 10px;
  }
  .wip-box-frame {
    display: grid;
    grid-template-columns: minmax(150px, 1fr) minmax(120px, 0.8fr) minmax(150px, 1fr);
    border: 2px solid #4c9b61;
    border-radius: 12px;
    overflow: hidden;
    min-width: 520px;
  }
  .wip-box-side {
    display: grid;
    gap: 8px;
    padding: 12px;
    background: #fbfefd;
  }
  .wip-box-left { border-right: 1px solid #d9eadf; }
  .wip-box-right { border-left: 1px solid #d9eadf; }
  .wip-box-center {
    display: flex;
    align-items: center;
    justify-content: center;
    background: #dff2e4;
    color: #173a22;
    font-weight: 800;
    letter-spacing: 0.08em;
  }
  .wip-box-entry {
    border: 1px solid #8fc39e;
    border-radius: 10px;
    background: #eef8f0;
    padding: 10px;
    text-align: center;
    font-weight: 700;
    color: #173a22;
  }
  .wip-box-notes {
    display: flex;
    flex-wrap: wrap;
    gap: 8px;
    margin-top: 12px;
  }
  .wip-box-notes span {
    background: #f1f6f3;
    border: 1px solid #d8e8dd;
    border-radius: 999px;
    color: #31543d;
    padding: 5px 10px;
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
    .wip-box-frame {
      grid-template-columns: 1fr;
      min-width: 0;
    }
    .wip-box-left,
    .wip-box-right { border: 0; }
    .wip-box-center { padding: 12px; border-top: 1px solid #d9eadf; border-bottom: 1px solid #d9eadf; }
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
      <div class="visual-aid-mode-tabs" role="tablist" aria-label="勘定連絡図の表示切替">
        <button type="button" class="active" data-visual-mode="basic">基本図を見る</button>
        <button type="button" data-visual-mode="editor">自分で作る</button>
      </div>
      <div class="account-flow-static-area" data-visual-basic-area>
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
      <div class="account-flow-editor-area" data-visual-editor-area hidden>
        <p class="visual-aid-description">サイドバーからボックスをドラッグして追加し、接続点から矢印を伸ばして勘定の流れを作れます。図表データはこのブラウザに一時保存されます。</p>
        <div id="${EDITOR_MOUNT_ID}"></div>
      </div>
    </div>
  </section>
`;

const WIP_BOX_HTML = `
  <section class="panel visual-aid-panel" aria-label="図表で確認">
    <div class="visual-aid-header">
      <div>
        <p class="eyebrow">図表で確認</p>
        <h2>仕掛品BOX</h2>
      </div>
      <button type="button" class="secondary" data-visual-aid-toggle>図表で確認</button>
    </div>
    <div class="visual-aid-body" data-visual-aid-body hidden>
      <p class="visual-aid-description">仕掛品BOXは、月初仕掛品と当月投入が、完成品と月末仕掛品へ分かれる関係を整理するための図です。</p>
      <div class="visual-aid-diagram wip-box-diagram" aria-label="仕掛品BOX">
        <div class="wip-box-title">仕掛品BOX</div>
        <div class="wip-box-frame">
          <div class="wip-box-side wip-box-left" aria-label="投入側">
            <div class="wip-box-entry">月初仕掛品</div>
            <div class="wip-box-entry">当月投入</div>
          </div>
          <div class="wip-box-center" aria-hidden="true">仕掛品</div>
          <div class="wip-box-side wip-box-right" aria-label="完成・月末側">
            <div class="wip-box-entry">完成品</div>
            <div class="wip-box-entry">月末仕掛品</div>
          </div>
        </div>
        <div class="wip-box-notes" aria-label="確認ポイント">
          <span>材料費</span><span>加工費</span><span>換算量</span><span>完成品換算量</span>
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

function getCurrentVisualAidKind(): 'industrial_account_flow' | 'wip_box' | null {
  const questionCard = document.querySelector('.question-reference-card');
  const text = questionCard?.textContent ?? '';
  if (!text.includes('教材なしモード')) return null;
  if (text.includes('industrial_4_1')) return 'industrial_account_flow';
  if (text.includes('industrial_4_2') || text.includes('industrial_4_4')) return 'wip_box';
  return null;
}

function unmountEditor() {
  const appWindow = window as BridgeWindow;
  try {
    appWindow.__boki2AccountFlowEditorRoot?.unmount();
  } catch (error) {
    console.warn('勘定連絡図エディタのunmountをスキップしました', error);
  }
  appWindow.__boki2AccountFlowEditorRoot = null;
}

function removeVisualAidIfNeeded() {
  const root = document.getElementById(VISUAL_AID_ROOT_ID);
  if (root && !getCurrentVisualAidKind()) {
    unmountEditor();
    root.remove();
  }
}

function mountEditorIfNeeded() {
  const appWindow = window as BridgeWindow;
  const target = document.getElementById(EDITOR_MOUNT_ID);
  if (!target || appWindow.__boki2AccountFlowEditorRoot) return;
  const editorRoot = createRoot(target);
  editorRoot.render(createElement(AccountFlowEditor));
  appWindow.__boki2AccountFlowEditorRoot = editorRoot;
}

function wireVisualAidControls(root: HTMLElement) {
  const button = root.querySelector<HTMLButtonElement>('[data-visual-aid-toggle]');
  const body = root.querySelector<HTMLElement>('[data-visual-aid-body]');
  const basicArea = root.querySelector<HTMLElement>('[data-visual-basic-area]');
  const editorArea = root.querySelector<HTMLElement>('[data-visual-editor-area]');
  const modeButtons = Array.from(root.querySelectorAll<HTMLButtonElement>('[data-visual-mode]'));

  function setMode(mode: 'basic' | 'editor') {
    if (!basicArea || !editorArea) return;
    basicArea.hidden = mode !== 'basic';
    editorArea.hidden = mode !== 'editor';
    modeButtons.forEach((modeButton) => modeButton.classList.toggle('active', modeButton.dataset.visualMode === mode));
    if (mode === 'editor') mountEditorIfNeeded();
  }

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

  modeButtons.forEach((modeButton) => {
    modeButton.addEventListener('click', () => {
      const mode = modeButton.dataset.visualMode === 'editor' ? 'editor' : 'basic';
      setMode(mode);
    });
  });
}

function ensureVisualAid() {
  if (typeof document === 'undefined') return;
  ensureVisualAidStyle();
  const answerArea = document.querySelector('.focused-answer-area.answer-only-main');
  if (!answerArea) return;
  const visualAidKind = getCurrentVisualAidKind();
  if (!visualAidKind) {
    removeVisualAidIfNeeded();
    return;
  }
  const existingRoot = document.getElementById(VISUAL_AID_ROOT_ID);
  if (existingRoot?.dataset.visualAidKind === visualAidKind) return;
  if (existingRoot) {
    unmountEditor();
    existingRoot.remove();
  }

  const root = document.createElement('div');
  root.id = VISUAL_AID_ROOT_ID;
  root.dataset.visualAidKind = visualAidKind;
  root.innerHTML = visualAidKind === 'industrial_account_flow' ? INDUSTRIAL_ACCOUNT_FLOW_HTML : WIP_BOX_HTML;
  answerArea.parentElement?.insertBefore(root, answerArea);
  wireVisualAidControls(root);
}

export function installVisualAidBridge() {
  if (typeof window === 'undefined' || typeof document === 'undefined') return;
  const appWindow = window as BridgeWindow;
  if (appWindow.__boki2VisualAidBridgeInstalled) return;
  appWindow.__boki2VisualAidBridgeInstalled = true;

  const observer = new MutationObserver(() => ensureVisualAid());
  observer.observe(document.body, { childList: true, subtree: true });
  window.setTimeout(() => ensureVisualAid(), 0);
}
