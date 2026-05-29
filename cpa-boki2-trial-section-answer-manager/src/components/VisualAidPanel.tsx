import { useState } from 'react';
import type { IndustrialAccountFlowData, PracticeItem, VisualAidType, WipBoxData } from '../originalStudyTypes';
import './visualAids.css';

function renderIndustrialAccountFlow(data: IndustrialAccountFlowData) {
  const nodeSet = new Set(data.nodes);
  const safeEdges = data.edges.filter(([from, to]) => nodeSet.has(from) && nodeSet.has(to));

  return (
    <div className="visual-aid-diagram industrial-account-flow" aria-label="勘定連絡図">
      <div className="flow-row flow-input-row">
        <div className="flow-node">材料</div>
        <div className="flow-arrow" aria-hidden="true">→</div>
        <div className="flow-node flow-main-node">仕掛品</div>
        <div className="flow-arrow" aria-hidden="true">→</div>
        <div className="flow-node">製品</div>
        <div className="flow-arrow" aria-hidden="true">→</div>
        <div className="flow-node">売上原価</div>
      </div>
      <div className="flow-row flow-support-row">
        <div className="flow-node">賃金</div>
        <div className="flow-arrow" aria-hidden="true">→</div>
        <div className="flow-node flow-main-node">仕掛品</div>
      </div>
      <div className="flow-row flow-support-row">
        <div className="flow-node">経費</div>
        <div className="flow-arrow" aria-hidden="true">→</div>
        <div className="flow-node">製造間接費</div>
        <div className="flow-arrow" aria-hidden="true">→</div>
        <div className="flow-node flow-main-node">仕掛品</div>
      </div>
      <div className="flow-edge-list" aria-label="表示している流れ">
        {safeEdges.map(([from, to]) => <span key={`${from}-${to}`}>{from} → {to}</span>)}
      </div>
    </div>
  );
}

function renderWipBox(data: WipBoxData) {
  return (
    <div className="visual-aid-diagram wip-box-diagram" aria-label="仕掛品BOX">
      <div className="wip-box-title">{data.title}</div>
      <div className="wip-box-frame">
        <div className="wip-box-side wip-box-left" aria-label="投入側">
          {data.rows.map((row) => <div className="wip-box-entry" key={`left-${row.leftLabel}`}>{row.leftLabel}</div>)}
        </div>
        <div className="wip-box-center" aria-hidden="true">仕掛品</div>
        <div className="wip-box-side wip-box-right" aria-label="完成・月末側">
          {data.rows.map((row) => <div className="wip-box-entry" key={`right-${row.rightLabel}`}>{row.rightLabel}</div>)}
        </div>
      </div>
      <div className="wip-box-notes" aria-label="確認ポイント">
        {data.notes.map((note) => <span key={note}>{note}</span>)}
      </div>
    </div>
  );
}

function hasVisualAid(item: PracticeItem): boolean {
  return Array.isArray(item.visualAidTypes) && item.visualAidTypes.length > 0;
}

export function VisualAidPanel({ item }: { item: PracticeItem }) {
  const [open, setOpen] = useState(false);
  const firstType = item.visualAidTypes?.[0] as VisualAidType | undefined;
  const accountFlow = item.visualAidData?.industrial_account_flow;
  const wipBox = item.visualAidData?.wip_box;
  const title = firstType === 'wip_box' ? wipBox?.title : accountFlow?.title;

  if (!hasVisualAid(item) || !firstType) return null;

  return (
    <section className="panel visual-aid-panel" aria-label="図表で確認">
      <div className="visual-aid-header">
        <div>
          <p className="eyebrow">図表で確認</p>
          <h2>{title ?? '図表'}</h2>
        </div>
        <button type="button" className="secondary" onClick={() => setOpen((value) => !value)}>
          {open ? '図表を閉じる' : '図表で確認'}
        </button>
      </div>
      {open && firstType === 'industrial_account_flow' && accountFlow && (
        <div className="visual-aid-body">
          <p className="visual-aid-description">
            {accountFlow.description ?? '材料・賃金・経費が製造活動を通じて仕掛品、製品、売上原価へ流れる関係を確認する図です。'}
          </p>
          {renderIndustrialAccountFlow(accountFlow)}
        </div>
      )}
      {open && firstType === 'wip_box' && wipBox && (
        <div className="visual-aid-body">
          <p className="visual-aid-description">
            {wipBox.description ?? '仕掛品BOXは、月初仕掛品と当月投入が、完成品と月末仕掛品へ分かれる関係を整理するための図です。'}
          </p>
          {renderWipBox(wipBox)}
        </div>
      )}
    </section>
  );
}

export default VisualAidPanel;
