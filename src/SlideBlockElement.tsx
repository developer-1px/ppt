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
import type { Point } from 'canvas/core'
import {
  EMPTY_TEXT_BOX_HEIGHT,
  rectToAutoHeightStyle,
  type Rect,
  type SlideBlock,
} from './retouchModel'
import {
  applyAutoHeightStyle,
  autoHeightRect,
  placeCaretAtEnd,
  placeCaretFromPoint,
  rememberStageScroll,
  restoreStageScroll,
} from './editableTextDom'
import {
  HTML_SLIDE_CLASSES,
  htmlSlideBlockAttributes,
} from './htmlSlideContract'
import {
  insertPlainTextBlockEditorLineBreak,
  insertPlainTextBlockEditorText,
  isPlainTextBlockEditorCancelShortcut,
  isPlainTextBlockEditorCommitShortcut,
  isPlainTextBlockEditorKeySafe,
  isPlainTextBlockEditorLineBreakKey,
  isPlainTextBlockEditorPrintableKey,
  isPlainTextBlockEditorUndoShortcut,
  normalizePlainTextBlockEditorText,
  plainTextBlockEditorAttributes,
  readPlainTextBlockEditorBeforeInput,
} from './plainTextBlockEditor'
import { usePendingTrailingLineBreakInput } from './usePendingTrailingLineBreakInput'

type SlideBlockElementProps = {
  block: SlideBlock
  blockIndex: number
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
  blockIndex,
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
    onCommit(
      normalizePlainTextBlockEditorText(element?.textContent ?? text),
      nextRect,
    )
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
    if (!isPlainTextBlockEditorKeySafe(event)) {
      return
    }

    if (isPlainTextBlockEditorUndoShortcut(event)) {
      consumeEditorEvent(event)
      undoDraft()
      return
    }

    if (isPlainTextBlockEditorCancelShortcut(event)) {
      consumeEditorEvent(event)
      resetDraft()
      onCancel()
      return
    }

    if (isPlainTextBlockEditorCommitShortcut(event)) {
      consumeEditorEvent(event)
      commit()
      return
    }

    if (isPlainTextBlockEditorLineBreakKey(event)) {
      consumeEditorEvent(event)
      const lineBreak = insertPlainTextBlockEditorLineBreak(event.currentTarget)

      rememberTrailingLineBreak(
        lineBreak.lineBreakAtEnd,
        lineBreak.beforeText,
      )
      syncAutoHeight(event.currentTarget)
      return
    }

    if (isPlainTextBlockEditorPrintableKey(event)) {
      consumeEditorEvent(event)
      if (!insertAfterPendingTrailingLineBreak(event.currentTarget, event.key)) {
        insertPlainTextBlockEditorText(event.currentTarget, event.key)
      }
      syncAutoHeight(event.currentTarget)
    }
  }

  function handleBeforeInput(event: ReactFormEvent<HTMLElement>) {
    const input = readPlainTextBlockEditorBeforeInput(event.nativeEvent)

    if (!input) {
      return
    }

    if (input.kind === 'text') {
      consumeEditorEvent(event)
      if (
        !insertAfterPendingTrailingLineBreak(
          event.currentTarget,
          input.text,
        )
      ) {
        insertPlainTextBlockEditorText(event.currentTarget, input.text)
      }
      syncAutoHeight(event.currentTarget)
      return
    }

    consumeEditorEvent(event)
    const lineBreak = insertPlainTextBlockEditorLineBreak(event.currentTarget)

    rememberTrailingLineBreak(lineBreak.lineBreakAtEnd, lineBreak.beforeText)
    syncAutoHeight(event.currentTarget)
  }

  function handlePaste(event: ReactClipboardEvent<HTMLElement>) {
    consumeEditorEvent(event)
    const text = event.clipboardData.getData('text/plain')

    if (!insertAfterPendingTrailingLineBreak(event.currentTarget, text)) {
      insertPlainTextBlockEditorText(event.currentTarget, text)
    }
    syncAutoHeight(event.currentTarget)
  }

  function handleInput(event: ReactFormEvent<HTMLElement>) {
    repairPendingTrailingLineBreak(event.currentTarget)
    syncAutoHeight(event.currentTarget)
  }

  function handleTextEventPropagation(event: {
    stopPropagation: () => void
  }) {
    if (editing) {
      event.stopPropagation()
    }
  }

  function handleBlockClick(event: ReactMouseEvent<HTMLElement>) {
    event.stopPropagation()
    if (!editing) {
      onClick(event)
    }
  }

  function handleBlockPointerDown(event: ReactPointerEvent<HTMLElement>) {
    if (editing) {
      event.stopPropagation()
      return
    }

    onPointerDown(event)
  }

  function setEditorElement(element: HTMLElement | null) {
    editorRef.current = element
  }

  function setBlockElement(element: HTMLElement | null) {
    blockRef.current = element
  }

  const textContent = (
    <span
      className={HTML_SLIDE_CLASSES.blockText}
      {...plainTextBlockEditorAttributes({
        blockId: block.id,
        editing,
      })}
      onBeforeInput={editing ? handleBeforeInput : undefined}
      onBlur={editing ? commit : undefined}
      onClick={handleTextEventPropagation}
      onInput={editing ? handleInput : undefined}
      onKeyDown={editing ? handleKeyDown : undefined}
      onPaste={editing ? handlePaste : undefined}
      onPointerDown={handleTextEventPropagation}
      ref={setEditorElement}
    >
      {editing ? null : text}
    </span>
  )

  const sharedProps = {
    ...htmlSlideBlockAttributes(block, blockIndex),
    'data-empty': text.length === 0 ? 'true' : undefined,
    'data-selected': selected ? 'true' : 'false',
    className,
    onClick: handleBlockClick,
    onPointerDown: handleBlockPointerDown,
    ref: setBlockElement,
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

function consumeEditorEvent(event: {
  preventDefault: () => void
  stopPropagation: () => void
}) {
  event.preventDefault()
  event.stopPropagation()
}
