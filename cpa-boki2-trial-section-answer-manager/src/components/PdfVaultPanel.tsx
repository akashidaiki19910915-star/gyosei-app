import { useState } from 'react';
import type { MaterialPdf } from '../types';
import { formatDateTime, nowIso } from '../utils/dates';
import { formatFileSize, getPdfPageCount } from '../utils/pdfRenderer';

interface Props {
  pdfs: MaterialPdf[];
  onSavePdf: (pdf: MaterialPdf) => Promise<void>;
  onDeletePdf: (id: string) => Promise<void>;
  onMessage: (message: string) => void;
}

export function PdfVaultPanel({ pdfs, onSavePdf, onDeletePdf, onMessage }: Props) {
  const [title, setTitle] = useState('');
  const [sourceName, setSourceName] = useState('CPA問題集');
  const [bookType, setBookType] = useState<MaterialPdf['bookType']>('商業簿記');
  const [loading, setLoading] = useState(false);

  async function importPdf(file: File | undefined) {
    if (!file) return;
    if (file.type !== 'application/pdf' && !file.name.toLowerCase().endsWith('.pdf')) {
      onMessage('PDFファイルのみ取り込めます');
      return;
    }
    try {
      setLoading(true);
      const now = nowIso();
      const pageCount = await getPdfPageCount(file);
      await onSavePdf({
        id: crypto.randomUUID(),
        title: title.trim() || file.name.replace(/\.pdf$/i, ''),
        sourceName: sourceName.trim() || '教材PDF',
        examType: '日商簿記2級',
        bookType,
        fileName: file.name,
        mimeType: file.type || 'application/pdf',
        size: file.size,
        pageCount,
        pdfBlob: file,
        createdAt: now,
        updatedAt: now,
      });
      setTitle('');
      onMessage('PDF教材を端末内IndexedDBへ保存しました。GitHubや外部サーバーには送信していません。');
    } catch {
      onMessage('PDFの読み込みに失敗しました。破損ファイル、保護ファイル、ブラウザ非対応の可能性があります。');
    } finally {
      setLoading(false);
    }
  }

  async function removePdf(pdf: MaterialPdf) {
    if (!window.confirm('このPDF教材データと問題IDへの紐付けを削除します。答案履歴・白紙再現カードは削除されません。本当に削除しますか？')) return;
    await onDeletePdf(pdf.id);
  }

  return (
    <details className="panel management-panel" open>
      <summary>PDF教材Vault</summary>
      <div className="panel-body">
        <p className="legal-notice compact">
          この機能は、正規に取得・購入した教材PDFを、利用者本人の私的学習のために端末内で整理・参照する目的のものです。教材PDF・問題文・解答・解説を第三者と共有したり、GitHub・SNS・クラウド公開領域へアップロードしたりしないでください。DRM等の技術的保護手段を回避して取り込むことはしないでください。
        </p>
        <div className="form-grid three-columns">
          <label>表示名
            <input value={title} onChange={(event) => setTitle(event.target.value)} placeholder="例：商業簿記 試験対策編" />
          </label>
          <label>教材区分
            <select value={bookType} onChange={(event) => setBookType(event.target.value as MaterialPdf['bookType'])}>
              <option value="商業簿記">商業簿記</option>
              <option value="工業簿記">工業簿記</option>
              <option value="その他">その他</option>
            </select>
          </label>
          <label>出所メモ
            <input value={sourceName} onChange={(event) => setSourceName(event.target.value)} placeholder="例：正規取得PDF" />
          </label>
        </div>
        <label className="import-label pdf-import-label">PDFを選択して端末内に保存
          <input type="file" accept="application/pdf" disabled={loading} onChange={(event) => { void importPdf(event.target.files?.[0]); }} />
        </label>
        <div className="table-wrap short-wrap">
          <table className="compact-table">
            <thead><tr><th>教材名</th><th>区分</th><th>ファイル名</th><th>ページ数</th><th>サイズ</th><th>保存日時</th><th>操作</th></tr></thead>
            <tbody>
              {pdfs.map((pdf) => (
                <tr key={pdf.id}>
                  <td>{pdf.title}</td>
                  <td>{pdf.bookType}</td>
                  <td>{pdf.fileName}</td>
                  <td>{pdf.pageCount || '-'}</td>
                  <td>{formatFileSize(pdf.size)}</td>
                  <td>{formatDateTime(pdf.createdAt)}</td>
                  <td><button className="danger" onClick={() => { void removePdf(pdf); }}>削除</button></td>
                </tr>
              ))}
              {pdfs.length === 0 && <tr><td colSpan={7}>PDF教材は未登録です。端末内のPDFを選択してください。</td></tr>}
            </tbody>
          </table>
        </div>
      </div>
    </details>
  );
}
