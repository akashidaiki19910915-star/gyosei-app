import { useState } from 'react';
import type { IndustrialAccountFlowData, PracticeItem, VisualAidType } from '../originalStudyTypes';
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

function hasVisualAid(item: PracticeItem): boolean {
  return Array.isArray(item.visualAidTypes) && item.visualAidTypes.length > 0;
}

export function VisualAidPanel({ item }: { item: PracticeItem }) {
  const [open, setOpen] = useState(false);
  const firstType = item.visualAidTypes?.[0] as VisualAidType | undefined;
  const accountFlow = item.visualAidData?.industrial_account_flow;

  if (!hasVisualAid(item) || !firstType) return null;

  return (
    <section className="panel visual-aid-panel" aria-label="図表で確認">
      <div className="visual-aid-header">
        <div>
          <p className="eyebrow">図表で確認</p>
          <h2>{accountFlow?.title ?? '図表'}</h2>
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
    </section>
  );
}

export default VisualAidPanel;
