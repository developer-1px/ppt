import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useRef,
  type ClipboardEvent as ReactClipboardEvent,
  type FormEvent as ReactFormEvent,
  type KeyboardEvent as ReactKeyboardEvent,
  type MouseEvent as ReactMouseEvent,
  type PointerEvent as ReactPointerEvent,
} from 'react'
import {
  EMPTY_TEXT_BOX_HEIGHT,
  rectToAutoHeightStyle,
  type Rect,
  type SlideBlock,
} from './retouchModel'
import {
  applyAutoHeightStyle,
  autoHeightRect,
  insertTextAtSelection,
  isSelectionAtTextEnd,
  normalizeEditableText,
  placeCaretAtEnd,
  placeCaretFromPoint,
  rememberStageScroll,
  restoreStageScroll,
} from './editableTextDom'
import { isHistoryShortcut } from './editorKeyboard'
import type { Point } from './layoutInteraction'
import { usePendingTrailingLineBreakInput } from './usePendingTrailingLineBreakInput'

type SlideBlockElementProps = {
  block: SlideBlock
  className: string
  editing: boolean
  initialClientPoint?: Point
  onClick: (event: ReactMouseEvent<HTMLElement>) => void
  onPointerDown: (event: ReactPointerEvent<HTMLElement>) => void
  onCancel: () => void
  onCommit: (text: string, rect: Rect) => void
  minimumHeight: number
  rect: Rect
  selected: boolean
  text: string
}

export function SlideBlockElement({
  block,
  className,
  editing,
  initialClientPoint,
  onClick,
  onPointerDown,
  onCancel,
  onCommit,
  minimumHeight,
  rect,
  selected,
  text,
}: SlideBlockElementProps) {
  const blockRef = useRef<HTMLElement | null>(null)
  const editorRef = useRef<HTMLElement | null>(null)
  const editingSessionRef = useRef<{ blockId: string; text: string } | null>(null)
  const rectRef = useRef(rect)
  const committedRef = useRef(false)
  const {
    clearPendingTrailingLineBreak,
    insertAfterPendingTrailingLineBreak,
    rememberTrailingLineBreak,
    repairPendingTrailingLineBreak,
  } = usePendingTrailingLineBreakInput()

  useEffect(() => {
    rectRef.current = rect
  }, [rect])

  useLayoutEffect(() => {
    if (!editing) {
      committedRef.current = false
      editingSessionRef.current = null
      return
    }

    const element = editorRef.current

    if (!element) {
      return
    }

    const editingSession = editingSessionRef.current

    if (
      !editingSession ||
      editingSession.blockId !== block.id ||
      editingSession.text !== text
    ) {
      element.textContent = text
      editingSessionRef.current = { blockId: block.id, text }
    }

    committedRef.current = false
    const scrollPosition = rememberStageScroll(element)

    element.focus({ preventScroll: true })
    restoreStageScroll(scrollPosition)
    if (
      !initialClientPoint ||
      !placeCaretFromPoint(element, initialClientPoint.x, initialClientPoint.y)
    ) {
      placeCaretAtEnd(element)
    }
    restoreStageScroll(scrollPosition)
  }, [block.id, editing, initialClientPoint, text])

  const syncAutoHeight = useCallback(
    (element: HTMLElement) => {
      const blockElement = blockRef.current

      if (!blockElement) {
        return rectRef.current
      }

      const effectiveMinimumHeight =
        element.textContent?.length === 0 ? EMPTY_TEXT_BOX_HEIGHT : minimumHeight
      const nextRect = autoHeightRect(
        blockElement,
        rectRef.current,
        effectiveMinimumHeight,
      )

      rectRef.current = nextRect
      applyAutoHeightStyle(blockElement, nextRect, effectiveMinimumHeight)

      return nextRect
    },
    [minimumHeight],
  )

  const commit = useCallback(() => {
    if (!editing || committedRef.current) {
      return
    }

    const element = editorRef.current
    const nextRect = element ? syncAutoHeight(element) : rectRef.current

    committedRef.current = true
    element?.blur()
    onCommit(normalizeEditableText(element?.textContent ?? text), nextRect)
  }, [editing, onCommit, syncAutoHeight, text])

  function resetDraft() {
    const element = editorRef.current
    const blockElement = blockRef.current

    if (!element || !blockElement) {
      return
    }

    const effectiveMinimumHeight =
      text.length === 0 ? EMPTY_TEXT_BOX_HEIGHT : minimumHeight
    clearPendingTrailingLineBreak()
    element.textContent = text
    rectRef.current = rect
    applyAutoHeightStyle(blockElement, rect, effectiveMinimumHeight)
  }

  function undoDraft() {
    resetDraft()

    if (editorRef.current) {
      placeCaretAtEnd(editorRef.current)
    }
  }

  function handleKeyDown(event: ReactKeyboardEvent<HTMLElement>) {
    if (event.nativeEvent.isComposing) {
      return
    }

    if (
      isHistoryShortcut(event.nativeEvent) &&
      event.key.toLowerCase() === 'z' &&
      !event.shiftKey
    ) {
      event.preventDefault()
      event.stopPropagation()
      undoDraft()
      return
    }

    if (event.key === 'Escape') {
      event.preventDefault()
      event.stopPropagation()
      resetDraft()
      onCancel()
      return
    }

    if (event.key === 'Enter' && (event.metaKey || event.ctrlKey)) {
      event.preventDefault()
      event.stopPropagation()
      commit()
      return
    }

    if (event.key === 'Enter') {
      event.preventDefault()
      event.stopPropagation()
      const beforeText = event.currentTarget.textContent ?? ''
      const lineBreakAtEnd = isSelectionAtTextEnd(event.currentTarget)

      insertTextAtSelection(event.currentTarget, '\n')
      rememberTrailingLineBreak(lineBreakAtEnd, beforeText)
      syncAutoHeight(event.currentTarget)
      return
    }

    if (
      event.key.length === 1 &&
      !event.altKey &&
      !event.ctrlKey &&
      !event.metaKey
    ) {
      event.preventDefault()
      event.stopPropagation()
      if (!insertAfterPendingTrailingLineBreak(event.currentTarget, event.key)) {
        insertTextAtSelection(event.currentTarget, event.key)
      }
      syncAutoHeight(event.currentTarget)
    }
  }

  function handleBeforeInput(event: ReactFormEvent<HTMLElement>) {
    const nativeEvent = event.nativeEvent as InputEvent

    if (nativeEvent.isComposing) {
      return
    }

    if (nativeEvent.inputType === 'insertText' && nativeEvent.data !== null) {
      event.preventDefault()
      event.stopPropagation()
      if (
        !insertAfterPendingTrailingLineBreak(
          event.currentTarget,
          nativeEvent.data,
        )
      ) {
        insertTextAtSelection(event.currentTarget, nativeEvent.data)
      }
      syncAutoHeight(event.currentTarget)
      return
    }

    if (
      nativeEvent.inputType === 'insertLineBreak' ||
      nativeEvent.inputType === 'insertParagraph'
    ) {
      event.preventDefault()
      event.stopPropagation()
      const beforeText = event.currentTarget.textContent ?? ''
      const lineBreakAtEnd = isSelectionAtTextEnd(event.currentTarget)

      insertTextAtSelection(event.currentTarget, '\n')
      rememberTrailingLineBreak(lineBreakAtEnd, beforeText)
      syncAutoHeight(event.currentTarget)
    }
  }

  function handlePaste(event: ReactClipboardEvent<HTMLElement>) {
    event.preventDefault()
    event.stopPropagation()
    const text = event.clipboardData.getData('text/plain')

    if (!insertAfterPendingTrailingLineBreak(event.currentTarget, text)) {
      insertTextAtSelection(event.currentTarget, text)
    }
    syncAutoHeight(event.currentTarget)
  }

  function handleInput(event: ReactFormEvent<HTMLElement>) {
    repairPendingTrailingLineBreak(event.currentTarget)
    syncAutoHeight(event.currentTarget)
  }

  const textContent = (
    <span
      className="slide-block-text"
      contentEditable={editing ? ('plaintext-only' as const) : undefined}
      data-editing={editing ? 'true' : undefined}
      data-editing-block={editing ? block.id : undefined}
      onBeforeInput={editing ? handleBeforeInput : undefined}
      onBlur={editing ? commit : undefined}
      onClick={(event: ReactMouseEvent<HTMLElement>) => {
        if (editing) {
          event.stopPropagation()
        }
      }}
      onInput={editing ? handleInput : undefined}
      onKeyDown={editing ? handleKeyDown : undefined}
      onPaste={editing ? handlePaste : undefined}
      onPointerDown={(event: ReactPointerEvent<HTMLElement>) => {
        if (editing) {
          event.stopPropagation()
        }
      }}
      ref={(element: HTMLElement | null) => {
        editorRef.current = element
      }}
      spellCheck={editing ? false : undefined}
      suppressContentEditableWarning={editing ? true : undefined}
    >
      {editing ? null : text}
    </span>
  )

  const sharedProps = {
    'data-block': block.id,
    'data-empty': text.length === 0 ? 'true' : undefined,
    'data-role': block.role,
    'data-selected': selected ? 'true' : 'false',
    className,
    onClick: (event: ReactMouseEvent<HTMLElement>) => {
      event.stopPropagation()
      if (!editing) {
        onClick(event)
      }
    },
    onPointerDown: (event: ReactPointerEvent<HTMLElement>) => {
      if (editing) {
        event.stopPropagation()
        return
      }

      onPointerDown(event)
    },
    ref: (element: HTMLElement | null) => {
      blockRef.current = element
    },
    style: rectToAutoHeightStyle(rect, minimumHeight),
    tabIndex: 0,
  }

  if (block.tag === 'h1') {
    return <h1 {...sharedProps}>{textContent}</h1>
  }

  if (block.tag === 'p') {
    return <p {...sharedProps}>{textContent}</p>
  }

  return <div {...sharedProps}>{textContent}</div>
}
