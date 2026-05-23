import type { KeyboardEvent as ReactKeyboardEvent } from 'react';

type NavigableElement = HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement;

interface TableNavigationOptions {
  isEditing: boolean;
  enterEditMode: () => void;
  exitEditMode: () => void;
}

function findCell(rowIndex: number, colIndex: number): NavigableElement | null {
  return document.querySelector<NavigableElement>(`[data-answer-cell="true"][data-row-index="${rowIndex}"][data-col-index="${colIndex}"]`);
}

function isTextArea(target: EventTarget | null): target is HTMLTextAreaElement {
  return target instanceof HTMLTextAreaElement;
}

function selectContent(element: NavigableElement): void {
  if (element instanceof HTMLInputElement) element.select();
}

export function focusCell(rowIndex: number, colIndex: number, select = true): void {
  const element = findCell(rowIndex, colIndex);
  if (!element) return;
  element.focus();
  if (select) selectContent(element);
}

function moveToCell(rowIndex: number, colIndex: number, options: TableNavigationOptions): void {
  if (rowIndex < 0 || colIndex < 0) return;
  const target = findCell(rowIndex, colIndex);
  if (!target) return;
  options.exitEditMode();
  target.focus();
  selectContent(target);
}

export function handleTableCellKeyDown(
  event: ReactKeyboardEvent<NavigableElement>,
  rowIndex: number,
  colIndex: number,
  options: TableNavigationOptions,
): void {
  const native = event.nativeEvent as KeyboardEvent;
  if (native.isComposing) return;

  const key = event.key;
  const textarea = isTextArea(event.target);

  if (key === 'F2') {
    event.preventDefault();
    options.enterEditMode();
    return;
  }

  if (key === 'Escape') {
    if (options.isEditing) event.preventDefault();
    options.exitEditMode();
    return;
  }

  if (options.isEditing) {
    if (textarea && key === 'Enter' && (event.altKey || event.ctrlKey)) return;
    if (key.startsWith('Arrow') || key === 'Home' || key === 'End') return;
    if (key === 'Enter') {
      event.preventDefault();
      moveToCell(event.shiftKey ? rowIndex - 1 : rowIndex + 1, colIndex, options);
    }
    return;
  }

  let nextRow = rowIndex;
  let nextCol = colIndex;
  if (key === 'ArrowRight') nextCol += 1;
  else if (key === 'ArrowLeft') nextCol -= 1;
  else if (key === 'ArrowDown' || (key === 'Enter' && !event.shiftKey)) nextRow += 1;
  else if (key === 'ArrowUp' || (key === 'Enter' && event.shiftKey)) nextRow -= 1;
  else return;

  event.preventDefault();
  moveToCell(nextRow, nextCol, options);
}
