import {
  useCallback,
  useEffect,
  useRef,
  type CSSProperties,
  type ClipboardEvent,
  type KeyboardEvent,
} from 'react'
import {
  MIN_BLOCK_SIZE,
  SLIDE_HEIGHT,
  rectToStyle,
  type Rect,
  type SlideBlock,
} from './retouchModel'

type PlainTextEditorProps = {
  block: SlideBlock
  minimumHeight: number
  onCancel: () => void
  onCommit: (text: string, rect: Rect) => void
  rect: Rect
}

export function PlainTextEditor({
  block,
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
    placeCaretAtEnd(editor)
  }, [])

  const commit = useCallback(() => {
    if (committedRef.current) {
      return
    }

    const editor = editorRef.current
    const nextRect = editor
      ? autoHeightRect(editor, rectRef.current, minimumHeight)
      : rectRef.current
    committedRef.current = true
    onCommit(editor?.textContent ?? block.text, nextRect)
  }, [block.text, minimumHeight, onCommit])

  function handleKeyDown(event: KeyboardEvent<HTMLDivElement>) {
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
  return {
    ...rectToStyle(rect),
    height: 'auto',
    minHeight: `${(Math.max(minimumHeight, MIN_BLOCK_SIZE) / SLIDE_HEIGHT) * 100}%`,
  }
}

function autoHeightRect(mount: HTMLElement, rect: Rect, minimumHeight: number): Rect {
  const mountBox = mount.getBoundingClientRect()

  if (mountBox.width === 0) {
    return rect
  }

  const slideUnitsPerCssPixel = rect.width / mountBox.width
  const contentHeight = mountBox.height * slideUnitsPerCssPixel
  const height = Math.max(minimumHeight, MIN_BLOCK_SIZE, Math.ceil(contentHeight))
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
