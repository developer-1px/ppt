import type { KeyboardEvent as ReactKeyboardEvent } from 'react'
import { isIMESafe, isPrintable } from '@interactive-os/keyboard'
import {
  insertTextAtSelection,
  isSelectionAtTextEnd,
  normalizeEditableText,
} from './editableTextDom'
import { isHistoryShortcut } from './editorKeyboard'

export const PLAIN_TEXT_BLOCK_EDITOR_SELECTOR =
  '[data-editing="true"][contenteditable]' as const
export const PLAIN_TEXT_BLOCK_EDITOR_CONTENT_EDITABLE = 'plaintext-only' as const

export type PlainTextBlockEditorBeforeInput =
  | { kind: 'line-break' }
  | { kind: 'text'; text: string }

export function plainTextBlockEditorAttributes({
  blockId,
  editing,
}: {
  blockId: string
  editing: boolean
}) {
  return {
    contentEditable: editing
      ? PLAIN_TEXT_BLOCK_EDITOR_CONTENT_EDITABLE
      : undefined,
    'data-editing': editing ? 'true' : undefined,
    'data-editing-block': editing ? blockId : undefined,
    spellCheck: editing ? false : undefined,
    suppressContentEditableWarning: editing ? true : undefined,
  } as const
}

export function isPlainTextBlockEditorKeySafe(
  event: ReactKeyboardEvent<HTMLElement>,
) {
  return isIMESafe(event.nativeEvent)
}

export function isPlainTextBlockEditorUndoShortcut(
  event: ReactKeyboardEvent<HTMLElement>,
) {
  return (
    isHistoryShortcut(event.nativeEvent) &&
    event.key.toLowerCase() === 'z' &&
    !event.shiftKey
  )
}

export function isPlainTextBlockEditorCancelShortcut(
  event: ReactKeyboardEvent<HTMLElement>,
) {
  return event.key === 'Escape'
}

export function isPlainTextBlockEditorCommitShortcut(
  event: ReactKeyboardEvent<HTMLElement>,
) {
  return event.key === 'Enter' && (event.metaKey || event.ctrlKey)
}

export function isPlainTextBlockEditorLineBreakKey(
  event: ReactKeyboardEvent<HTMLElement>,
) {
  return event.key === 'Enter'
}

export function isPlainTextBlockEditorPrintableKey(
  event: ReactKeyboardEvent<HTMLElement>,
) {
  return isPrintable(event.nativeEvent)
}

export function readPlainTextBlockEditorBeforeInput(
  event: InputEvent,
): PlainTextBlockEditorBeforeInput | null {
  if (event.isComposing) {
    return null
  }

  if (event.inputType === 'insertText' && event.data !== null) {
    return { kind: 'text', text: event.data }
  }

  if (
    event.inputType === 'insertLineBreak' ||
    event.inputType === 'insertParagraph'
  ) {
    return { kind: 'line-break' }
  }

  return null
}

export function insertPlainTextBlockEditorText(
  element: HTMLElement,
  text: string,
) {
  insertTextAtSelection(element, text)
}

export function insertPlainTextBlockEditorLineBreak(element: HTMLElement) {
  const beforeText = element.textContent ?? ''
  const lineBreakAtEnd = isSelectionAtTextEnd(element)

  insertPlainTextBlockEditorText(element, '\n')

  return {
    beforeText,
    lineBreakAtEnd,
  }
}

export function normalizePlainTextBlockEditorText(text: string) {
  return normalizeEditableText(text)
}
