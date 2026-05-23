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
  rect: Rect
}

export function NanoTextEditor({
  block,
  minimumHeight,
  onCancel,
  onCommit,
  rect,
}: NanoTextEditorProps) {
  const mountRef = useRef<HTMLDivElement | null>(null)
  const engineRef = useRef<NanoDocumentEngine | null>(null)
  const rectRef = useRef(rect)
  const onCancelRef = useRef(onCancel)
  const onCommitRef = useRef(onCommit)
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

  const commit = useCallback(() => {
    const engine = engineRef.current
    const rect = mountRef.current
      ? autoHeightRect(mountRef.current, rectRef.current, minimumHeight)
      : rectRef.current

    onCommitRef.current(
      engine ? textFromNanoDocument(engine.value) : block.text,
      rect,
    )
  }, [block.text, minimumHeight])

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

    let caretFrame = 0
    const focusFrame = requestAnimationFrame(() => {
      const editor = mount.querySelector<HTMLElement>('.ProseMirror')
      editor?.focus()
      if (editor) {
        caretFrame = requestAnimationFrame(() => placeCaretAtEnd(editor))
      }
    })

    return () => {
      cancelAnimationFrame(focusFrame)
      cancelAnimationFrame(caretFrame)
      handle.destroy()
      engineRef.current = null
    }
  }, [block.text.length, initialDocument])

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
      style={editorStyle(rect, minimumHeight)}
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
