import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  type CSSProperties,
} from 'react'
import {
  createNanoDocument,
  createNanoView,
  blockTextPointer,
  point,
  type NanoBlock,
  type NanoDocument,
  type NanoDocumentEngine,
} from 'nano-edit'
import {
  MIN_BLOCK_SIZE,
  SLIDE_HEIGHT,
  rectToStyle,
  type Rect,
  type SlideBlock,
} from './retouchModel'

type NanoTextEditorProps = {
  block: SlideBlock
  minimumHeight: number
  onCancel: () => void
  onCommit: (text: string, rect: Rect) => void
  onRectChange: (rect: Rect) => void
  rect: Rect
}

export function NanoTextEditor({
  block,
  minimumHeight,
  onCancel,
  onCommit,
  onRectChange,
  rect,
}: NanoTextEditorProps) {
  const mountRef = useRef<HTMLDivElement | null>(null)
  const engineRef = useRef<NanoDocumentEngine | null>(null)
  const rectRef = useRef(rect)
  const onCancelRef = useRef(onCancel)
  const onCommitRef = useRef(onCommit)
  const onRectChangeRef = useRef(onRectChange)
  const initialDocument = useMemo(
    () => nanoDocumentFromText(block.id, block.text),
    [block.id, block.text],
  )

  useEffect(() => {
    rectRef.current = rect
  }, [rect])

  useEffect(() => {
    onCancelRef.current = onCancel
  }, [onCancel])

  useEffect(() => {
    onCommitRef.current = onCommit
  }, [onCommit])

  useEffect(() => {
    onRectChangeRef.current = onRectChange
  }, [onRectChange])

  const commit = useCallback(() => {
    const engine = engineRef.current

    onCommitRef.current(
      engine ? textFromNanoDocument(engine.value) : block.text,
      rectRef.current,
    )
  }, [block.text])

  const handleEditorKeyDown = useCallback(
    (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        event.preventDefault()
        event.stopPropagation()
        onCancelRef.current()
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
    },
    [commit],
  )

  useEffect(() => {
    const mount = mountRef.current

    if (!mount) {
      return
    }

    mount.addEventListener('keydown', handleEditorKeyDown, true)

    return () => {
      mount.removeEventListener('keydown', handleEditorKeyDown, true)
    }
  }, [handleEditorKeyDown])

  useEffect(() => {
    const mount = mountRef.current

    if (!mount) {
      return
    }

    const engine = createNanoDocument(initialDocument)
    const lastBlockIndex = Math.max(0, initialDocument.blocks.length - 1)
    const lastBlock = initialDocument.blocks[lastBlockIndex]
    engine.selection?.collapse(point(
      blockTextPointer(lastBlockIndex),
      lastBlock ? blockText(lastBlock).length : 0,
    ))
    const handle = createNanoView({ mount, engine })
    engineRef.current = engine

    const measure = () => {
      const nextRect = autoHeightRect(mount, rectRef.current, minimumHeight)
      if (Math.abs(nextRect.height - rectRef.current.height) >= 1) {
        rectRef.current = nextRect
        onRectChangeRef.current(nextRect)
      }
    }
    const observer = new ResizeObserver(() => measure())
    const editor = mount.querySelector<HTMLElement>('.editor')
    if (editor) {
      observer.observe(editor)
    }

    let caretFrame = 0
    let measureFrame = 0
    const focusFrame = requestAnimationFrame(() => {
      const editor = mount.querySelector<HTMLElement>('.ProseMirror')
      editor?.focus()
      if (editor) {
        caretFrame = requestAnimationFrame(() => placeCaretAtEnd(editor))
      }
      measureFrame = requestAnimationFrame(measure)
    })

    return () => {
      observer.disconnect()
      cancelAnimationFrame(focusFrame)
      cancelAnimationFrame(caretFrame)
      cancelAnimationFrame(measureFrame)
      handle.destroy()
      engineRef.current = null
    }
  }, [block.text.length, initialDocument, minimumHeight])

  return (
    <div
      className={`nano-text-editor ${block.className}`}
      onBlur={(event) => {
        if (!event.currentTarget.contains(event.relatedTarget as Node | null)) {
          commit()
        }
      }}
      onPointerDown={(event) => event.stopPropagation()}
      ref={mountRef}
      style={rectToStyle(rect) as CSSProperties}
    />
  )
}

function nanoDocumentFromText(blockId: string, text: string): NanoDocument {
  const lines = text.split('\n')
  const blocks = lines.length > 0 ? lines : ['']

  return {
    blocks: blocks.map((line, index) => ({
      id: `${blockId}-line-${index}`,
      marks: [],
      text: line,
      type: 'paragraph',
    })),
  }
}

function textFromNanoDocument(document: NanoDocument) {
  return document.blocks.map(blockText).join('\n')
}

function blockText(block: NanoBlock) {
  return 'text' in block && typeof block.text === 'string' ? block.text : ''
}

function autoHeightRect(mount: HTMLElement, rect: Rect, minimumHeight: number): Rect {
  const editor = mount.querySelector<HTMLElement>('.editor')
  const mountBox = mount.getBoundingClientRect()
  const editorBox = editor?.getBoundingClientRect()

  if (!editor || mountBox.width === 0 || !editorBox) {
    return rect
  }

  const style = getComputedStyle(mount)
  const verticalChrome =
    readPixels(style.paddingTop) +
    readPixels(style.paddingBottom) +
    readPixels(style.borderTopWidth) +
    readPixels(style.borderBottomWidth)
  const slideUnitsPerCssPixel = rect.width / mountBox.width
  const contentHeight = (editorBox.height + verticalChrome) * slideUnitsPerCssPixel
  const height = Math.max(minimumHeight, MIN_BLOCK_SIZE, Math.ceil(contentHeight))
  const y = Math.min(rect.y, Math.max(0, SLIDE_HEIGHT - height))

  return rect.height === height && rect.y === y ? rect : { ...rect, y, height }
}

function readPixels(value: string) {
  const number = Number.parseFloat(value)
  return Number.isFinite(number) ? number : 0
}

function placeCaretAtEnd(root: HTMLElement) {
  const selection = window.getSelection()
  const range = document.createRange()
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
  let lastTextNode: Node | null = null

  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    lastTextNode = node
  }

  if (lastTextNode?.textContent) {
    range.setStart(lastTextNode, lastTextNode.textContent.length)
  } else {
    range.selectNodeContents(root)
    range.collapse(false)
  }

  selection?.removeAllRanges()
  selection?.addRange(range)
}
