import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { ChangeEvent, DragEvent } from 'react';
import {
  addEdge,
  applyEdgeChanges,
  applyNodeChanges,
  Background,
  Controls,
  Handle,
  MarkerType,
  MiniMap,
  Position,
  ReactFlow,
  ReactFlowProvider,
  useReactFlow,
} from '@xyflow/react';
import type { Connection, Edge, EdgeChange, Node, NodeChange, NodeProps, NodeTypes } from '@xyflow/react';
import '@xyflow/react/dist/style.css';
import './accountFlowEditor.css';

type AccountBoxData = {
  accountName: string;
  amount: string;
  memo: string;
  onChange?: (nodeId: string, patch: Partial<AccountBoxData>) => void;
};

type AccountBoxNode = Node<AccountBoxData, 'accountBox'>;
type AccountFlowDraft = {
  nodes: AccountBoxNode[];
  edges: Edge[];
};

const STORAGE_KEY = 'boki2-industrial-4-1-account-flow-draft-v1';
const paletteItems = ['材料', '賃金', '経費', '製造間接費', '仕掛品', '製品', '売上原価', '自由入力ボックス'];

function stripRuntimeNodeData(node: AccountBoxNode): AccountBoxNode {
  const { onChange: _onChange, ...data } = node.data;
  return { ...node, data };
}

function sampleDraft(): AccountFlowDraft {
  return {
    nodes: [
      { id: 'sample-material', type: 'accountBox', position: { x: 40, y: 90 }, data: { accountName: '材料', amount: '', memo: '' } },
      { id: 'sample-wip', type: 'accountBox', position: { x: 260, y: 90 }, data: { accountName: '仕掛品', amount: '', memo: '' } },
      { id: 'sample-product', type: 'accountBox', position: { x: 480, y: 90 }, data: { accountName: '製品', amount: '', memo: '' } },
      { id: 'sample-cogs', type: 'accountBox', position: { x: 700, y: 90 }, data: { accountName: '売上原価', amount: '', memo: '' } },
    ],
    edges: [
      { id: 'edge-sample-material-wip', source: 'sample-material', target: 'sample-wip', markerEnd: { type: MarkerType.ArrowClosed } },
      { id: 'edge-sample-wip-product', source: 'sample-wip', target: 'sample-product', markerEnd: { type: MarkerType.ArrowClosed } },
      { id: 'edge-sample-product-cogs', source: 'sample-product', target: 'sample-cogs', markerEnd: { type: MarkerType.ArrowClosed } },
    ],
  };
}

function loadDraft(): AccountFlowDraft {
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    if (!raw) return sampleDraft();
    const parsed = JSON.parse(raw) as Partial<AccountFlowDraft>;
    if (!Array.isArray(parsed.nodes) || !Array.isArray(parsed.edges)) return sampleDraft();
    return {
      nodes: parsed.nodes.map((node) => stripRuntimeNodeData({
        ...node,
        type: 'accountBox',
        data: {
          accountName: String(node.data?.accountName ?? ''),
          amount: String(node.data?.amount ?? ''),
          memo: String(node.data?.memo ?? ''),
        },
      } as AccountBoxNode)),
      edges: parsed.edges,
    };
  } catch (error) {
    console.warn('勘定連絡図エディタの一時保存データを読み込めませんでした', error);
    return sampleDraft();
  }
}

function saveDraft(nodes: AccountBoxNode[], edges: Edge[]) {
  try {
    const payload: AccountFlowDraft = {
      nodes: nodes.map(stripRuntimeNodeData),
      edges,
    };
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
    window.dispatchEvent(new CustomEvent('boki2-account-flow-draft-saved', { detail: payload }));
  } catch (error) {
    console.warn('勘定連絡図エディタの一時保存に失敗しました', error);
  }
}

function AccountBoxNodeView({ id, data, selected }: NodeProps) {
  const box = data as AccountBoxData;
  const update = box.onChange;

  const handleChange = (field: keyof Omit<AccountBoxData, 'onChange'>) => (event: ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) => {
    update?.(id, { [field]: event.currentTarget.value });
  };

  return (
    <div className={selected ? 'account-flow-node selected' : 'account-flow-node'}>
      <Handle type="target" position={Position.Left} className="account-flow-handle" />
      <Handle type="source" position={Position.Right} className="account-flow-handle" />
      <label>
        <span>勘定科目</span>
        <input value={box.accountName} onChange={handleChange('accountName')} lang="ja" autoComplete="off" spellCheck={false} />
      </label>
      <label>
        <span>金額</span>
        <input value={box.amount} onChange={handleChange('amount')} inputMode="numeric" autoComplete="off" spellCheck={false} placeholder="例：2000" />
      </label>
      <label>
        <span>メモ</span>
        <textarea value={box.memo} onChange={handleChange('memo')} rows={2} lang="ja" spellCheck={false} />
      </label>
    </div>
  );
}

const nodeTypes: NodeTypes = { accountBox: AccountBoxNodeView };

function AccountFlowEditorCanvas() {
  const initial = useMemo(() => loadDraft(), []);
  const [nodes, setNodes] = useState<AccountBoxNode[]>(initial.nodes);
  const [edges, setEdges] = useState<Edge[]>(initial.edges);
  const [selectedNodeId, setSelectedNodeId] = useState<string | null>(null);
  const [selectedEdgeId, setSelectedEdgeId] = useState<string | null>(null);
  const [status, setStatus] = useState('図表データは自動で一時保存されます。');
  const flowWrapperRef = useRef<HTMLDivElement | null>(null);
  const { screenToFlowPosition } = useReactFlow();

  const updateNodeData = useCallback((nodeId: string, patch: Partial<AccountBoxData>) => {
    setNodes((current) => current.map((node) => node.id === nodeId ? { ...node, data: { ...node.data, ...patch } } : node));
  }, []);

  const nodesWithHandlers = useMemo<AccountBoxNode[]>(() => nodes.map((node) => ({
    ...node,
    data: { ...node.data, onChange: updateNodeData },
  })), [nodes, updateNodeData]);

  useEffect(() => {
    saveDraft(nodes, edges);
  }, [nodes, edges]);

  const onNodesChange = useCallback((changes: NodeChange[]) => setNodes((current) => applyNodeChanges(changes, current) as AccountBoxNode[]), []);
  const onEdgesChange = useCallback((changes: EdgeChange[]) => setEdges((current) => applyEdgeChanges(changes, current)), []);
  const onConnect = useCallback((connection: Connection) => setEdges((current) => addEdge({ ...connection, markerEnd: { type: MarkerType.ArrowClosed } }, current)), []);

  const addNode = useCallback((accountName: string, x = 120, y = 120) => {
    const label = accountName === '自由入力ボックス' ? '' : accountName;
    const id = `node-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
    setNodes((current) => [...current, { id, type: 'accountBox', position: { x, y }, data: { accountName: label, amount: '', memo: '' } }]);
    setStatus(`${accountName}を追加しました。`);
  }, []);

  const onDragStart = (event: DragEvent<HTMLButtonElement>, accountName: string) => {
    event.dataTransfer.setData('application/x-account-flow-node', accountName);
    event.dataTransfer.effectAllowed = 'move';
  };

  const onDragOver = (event: DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    event.dataTransfer.dropEffect = 'move';
  };

  const onDrop = (event: DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    const accountName = event.dataTransfer.getData('application/x-account-flow-node');
    if (!accountName) return;
    const position = screenToFlowPosition({ x: event.clientX, y: event.clientY });
    addNode(accountName, position.x, position.y);
  };

  const deleteSelected = () => {
    if (selectedNodeId) {
      setNodes((current) => current.filter((node) => node.id !== selectedNodeId));
      setEdges((current) => current.filter((edge) => edge.source !== selectedNodeId && edge.target !== selectedNodeId));
      setSelectedNodeId(null);
      setStatus('選択中のボックスを削除しました。');
      return;
    }
    if (selectedEdgeId) {
      setEdges((current) => current.filter((edge) => edge.id !== selectedEdgeId));
      setSelectedEdgeId(null);
      setStatus('選択中の矢印を削除しました。');
    }
  };

  const resetSample = () => {
    if (!window.confirm('現在の勘定連絡図をサンプルに戻しますか？')) return;
    const next = sampleDraft();
    setNodes(next.nodes);
    setEdges(next.edges);
    setSelectedNodeId(null);
    setSelectedEdgeId(null);
    setStatus('材料 → 仕掛品 → 製品 → 売上原価のサンプルに戻しました。');
  };

  return (
    <div className="account-flow-editor-shell">
      <aside className="account-flow-palette" aria-label="ボックス追加サイドバー">
        <h3>ボックスを追加</h3>
        <p>ドラッグして右のキャンバスへ追加します。クリックでも追加できます。</p>
        <div className="account-flow-palette-list">
          {paletteItems.map((name) => (
            <button key={name} type="button" draggable onDragStart={(event) => onDragStart(event, name)} onClick={() => addNode(name)}>
              {name}
            </button>
          ))}
        </div>
        <div className="account-flow-editor-actions">
          <button type="button" className="secondary" onClick={deleteSelected} disabled={!selectedNodeId && !selectedEdgeId}>選択中を削除</button>
          <button type="button" className="secondary" onClick={resetSample}>サンプルに戻す</button>
        </div>
        <p className="account-flow-editor-note">金額欄は文字列のまま保存します。自動採点とは連動しません。</p>
      </aside>
      <div className="account-flow-canvas-panel" ref={flowWrapperRef} onDragOver={onDragOver} onDrop={onDrop}>
        <ReactFlow
          nodes={nodesWithHandlers}
          edges={edges}
          nodeTypes={nodeTypes}
          onNodesChange={onNodesChange}
          onEdgesChange={onEdgesChange}
          onConnect={onConnect}
          onNodeClick={(_, node) => { setSelectedNodeId(node.id); setSelectedEdgeId(null); }}
          onEdgeClick={(_, edge) => { setSelectedEdgeId(edge.id); setSelectedNodeId(null); }}
          onPaneClick={() => { setSelectedNodeId(null); setSelectedEdgeId(null); }}
          fitView
          defaultEdgeOptions={{ markerEnd: { type: MarkerType.ArrowClosed } }}
        >
          <MiniMap pannable zoomable />
          <Controls />
          <Background />
        </ReactFlow>
      </div>
      <p className="account-flow-editor-status">{status}</p>
    </div>
  );
}

export function AccountFlowEditor() {
  return (
    <ReactFlowProvider>
      <AccountFlowEditorCanvas />
    </ReactFlowProvider>
  );
}

export default AccountFlowEditor;
