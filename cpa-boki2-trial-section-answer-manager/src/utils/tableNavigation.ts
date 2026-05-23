import type { KeyboardEvent as ReactKeyboardEvent } from 'react';

type NavigableElement = HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement;

function findCell(rowIndex: number, colIndex: number): NavigableElement | null {
  return document.querySelector<NavigableElement>(`[data-answer-cell="true"][data-row-index="${rowIndex}"][data-col-index="${colIndex}"]`);
}

function isTextArea(target: EventTarget | null): target is HTMLTextAreaElement {
  return target instanceof HTMLTextAreaElement;
}

function selectContent(element: NavigableElement): void {
  if (element instanceof HTMLInputElement) element.select();
}

export function focusCell(rowIndex: number, colIndex: number): void {
  const element = findCell(rowIndex, colIndex);
  if (!element) return;
  element.focus();
  selectContent(element);
}

export function handleTableCellKeyDown(
  event: ReactKeyboardEvent<NavigableElement>,
  rowIndex: number,
  colIndex: number,
): void {
  const native = event.nativeEvent as KeyboardEvent;
  if (native.isComposing) return;

  const key = event.key;
  const textarea = isTextArea(event.target);
  if (textarea && (key === 'ArrowLeft' || key === 'ArrowRight' || key === 'ArrowUp' || key === 'ArrowDown' || key === 'Enter')) return;

  let nextRow = rowIndex;
  let nextCol = colIndex;
  if (key === 'ArrowRight') nextCol += 1;
  else if (key === 'ArrowLeft') nextCol -= 1;
  else if (key === 'ArrowDown' || (key === 'Enter' && !event.shiftKey)) nextRow += 1;
  else if (key === 'ArrowUp' || (key === 'Enter' && event.shiftKey)) nextRow -= 1;
  else return;

  if (nextRow < 0 || nextCol < 0) return;
  const target = findCell(nextRow, nextCol);
  if (!target) return;
  event.preventDefault();
  target.focus();
  selectContent(target);
}
