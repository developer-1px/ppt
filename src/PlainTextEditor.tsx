import {
  useCallback,
  useEffect,
  useRef,
  type CSSProperties,
  type ClipboardEvent,
  type KeyboardEvent,
} from 'react'
import {
  EMPTY_TEXT_BOX_HEIGHT,
  SLIDE_HEIGHT,
  SLIDE_WIDTH,
  type Rect,
  type SlideBlock,
  rectToAutoHeightStyle,
} from './retouchModel'

type PlainTextEditorProps = {
  block: SlideBlock
  initialClientPoint?: {
    x: number
    y: number
  }
  minimumHeight: number
  onCancel: () => void
  onCommit: (text: string, rect: Rect) => void
  rect: Rect
}

export function PlainTextEditor({
  block,
  initialClientPoint,
  minimumHeight,
  onCancel,
  onCommit,
  rect,
}: PlainTextEditorProps) {
  const editorRef = useRef<HTMLDivElement | null>(null)
  const rectRef = useRef(rect)
  const committedRef = useRef(false)

  useEffect(() => {
    rectRef.current = rect
  }, [rect])

  useEffect(() => {
    const editor = editorRef.current

    if (!editor) {
      return
    }

    editor.focus()
    if (
      !initialClientPoint ||
      !placeCaretFromPoint(editor, initialClientPoint.x, initialClientPoint.y)
    ) {
      placeCaretAtEnd(editor)
    }
  }, [initialClientPoint])

  const syncAutoHeight = useCallback((editor: HTMLElement) => {
    const effectiveMinimumHeight =
      editor.textContent?.length === 0 ? EMPTY_TEXT_BOX_HEIGHT : minimumHeight
    const nextRect = autoHeightRect(
      editor,
      rectRef.current,
      effectiveMinimumHeight,
    )
    rectRef.current = nextRect
    applyEditorStyle(editor, nextRect, effectiveMinimumHeight)

    return nextRect
  }, [minimumHeight])

  const commit = useCallback(() => {
    if (committedRef.current) {
      return
    }

    const editor = editorRef.current
    const nextRect = editor ? syncAutoHeight(editor) : rectRef.current
    committedRef.current = true
    onCommit(editor?.textContent ?? block.text, nextRect)
  }, [block.text, onCommit, syncAutoHeight])

  function handleKeyDown(event: KeyboardEvent<HTMLDivElement>) {
    if (event.nativeEvent.isComposing) {
      return
    }

    if (event.key === 'Escape') {
      event.preventDefault()
      event.stopPropagation()
      onCancel()
      return
    }

    if (
      event.key === 'Enter' &&
      (event.metaKey || event.ctrlKey || !event.shiftKey)
    ) {
      event.preventDefault()
      event.stopPropagation()
      commit()
    }
  }

  function handlePaste(event: ClipboardEvent<HTMLDivElement>) {
    event.preventDefault()
    document.execCommand('insertText', false, event.clipboardData.getData('text/plain'))
  }

  return (
    <div
      className={`plain-text-editor ${block.className}`}
      contentEditable="plaintext-only"
      onBlur={commit}
      onInput={(event) => syncAutoHeight(event.currentTarget)}
      onKeyDown={handleKeyDown}
      onPaste={handlePaste}
      onPointerDown={(event) => event.stopPropagation()}
      ref={editorRef}
      spellCheck={false}
      style={editorStyle(rect, minimumHeight)}
      suppressContentEditableWarning
    >
      {block.text}
    </div>
  )
}

function editorStyle(rect: Rect, minimumHeight: number): CSSProperties {
  return rectToAutoHeightStyle(rect, minimumHeight)
}

function applyEditorStyle(
  editor: HTMLElement,
  rect: Rect,
  minimumHeight: number,
) {
  editor.style.left = `${(rect.x / SLIDE_WIDTH) * 100}%`
  editor.style.top = `${(rect.y / SLIDE_HEIGHT) * 100}%`
  editor.style.width = `${(rect.width / SLIDE_WIDTH) * 100}%`
  editor.style.height = 'auto'
  editor.style.minHeight =
    minimumHeight > 0 ? `${(minimumHeight / SLIDE_HEIGHT) * 100}%` : ''
}

function autoHeightRect(mount: HTMLElement, rect: Rect, minimumHeight: number): Rect {
  const mountBox = mount.getBoundingClientRect()

  if (mountBox.width === 0) {
    return rect
  }

  const slideUnitsPerCssPixel = rect.width / mountBox.width
  const contentHeight = mountBox.height * slideUnitsPerCssPixel
  const height = Math.max(minimumHeight, Math.ceil(contentHeight))
  const y = Math.min(rect.y, Math.max(0, SLIDE_HEIGHT - height))

  return rect.height === height && rect.y === y ? rect : { ...rect, y, height }
}

function placeCaretAtEnd(root: HTMLElement) {
  const selection = window.getSelection()
  const range = document.createRange()
  range.selectNodeContents(root)
  range.collapse(false)
  selection?.removeAllRanges()
  selection?.addRange(range)
}

function placeCaretFromPoint(root: HTMLElement, x: number, y: number) {
  const range = readCaretRangeFromPoint(x, y)

  if (!range || !root.contains(range.startContainer)) {
    return false
  }

  const selection = window.getSelection()
  selection?.removeAllRanges()
  selection?.addRange(range)

  return true
}

function readCaretRangeFromPoint(x: number, y: number) {
  if ('caretPositionFromPoint' in document) {
    const position = document.caretPositionFromPoint(x, y)

    if (position) {
      const range = document.createRange()
      range.setStart(position.offsetNode, position.offset)
      range.collapse(true)

      return range
    }
  }

  if ('caretRangeFromPoint' in document) {
    return document.caretRangeFromPoint(x, y)
  }

  return null
}
