import * as pdfjsLib from 'pdfjs-dist';
import pdfWorkerUrl from 'pdfjs-dist/build/pdf.worker.mjs?url';

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;

export async function getPdfPageCount(blob: Blob): Promise<number> {
  const data = new Uint8Array(await blob.arrayBuffer());
  const task = pdfjsLib.getDocument({ data });
  const pdf = await task.promise;
  return pdf.numPages;
}

export async function renderPdfPageToCanvas(blob: Blob, pageNumber: number, scale: number, canvas: HTMLCanvasElement): Promise<void> {
  const data = new Uint8Array(await blob.arrayBuffer());
  const task = pdfjsLib.getDocument({ data });
  const pdf = await task.promise;
  const safePage = Math.max(1, Math.min(pageNumber, pdf.numPages));
  const page = await pdf.getPage(safePage);
  const viewport = page.getViewport({ scale });
  const context = canvas.getContext('2d');
  if (!context) return;
  canvas.width = Math.floor(viewport.width);
  canvas.height = Math.floor(viewport.height);
  context.clearRect(0, 0, canvas.width, canvas.height);
  const renderTask = page.render({ canvasContext: context, viewport } as any);
  await renderTask.promise;
}

export async function extractPdfPageText(blob: Blob, pageNumber: number): Promise<string> {
  const data = new Uint8Array(await blob.arrayBuffer());
  const task = pdfjsLib.getDocument({ data });
  const pdf = await task.promise;
  const safePage = Math.max(1, Math.min(pageNumber, pdf.numPages));
  const page = await pdf.getPage(safePage);
  const textContent = await page.getTextContent();
  return textContent.items.map((item) => ('str' in item ? item.str : '')).join(' ');
}

export function formatFileSize(size: number): string {
  if (size >= 1024 * 1024) return `${(size / 1024 / 1024).toFixed(1)}MB`;
  if (size >= 1024) return `${Math.round(size / 1024)}KB`;
  return `${size}B`;
}
