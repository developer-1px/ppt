import {
  useEffect,
  useMemo,
  useRef,
  type CSSProperties,
  type KeyboardEvent as ReactKeyboardEvent,
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
  rectToStyle,
  type Rect,
  type SlideBlock,
} from './retouchModel'

type NanoTextEditorProps = {
  block: SlideBlock
  onCancel: () => void
  onCommit: (text: string) => void
  rect: Rect
}

export function NanoTextEditor({
  block,
  onCancel,
  onCommit,
  rect,
}: NanoTextEditorProps) {
  const mountRef = useRef<HTMLDivElement | null>(null)
  const engineRef = useRef<NanoDocumentEngine | null>(null)
  const initialDocument = useMemo(
    () => nanoDocumentFromText(block.id, block.text),
    [block.id, block.text],
  )

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

  function commit() {
    const engine = engineRef.current

    if (!engine) {
      onCommit(block.text)
      return
    }

    onCommit(textFromNanoDocument(engine.value))
  }

  function handleKeyDown(event: ReactKeyboardEvent<HTMLDivElement>) {
    if (event.key === 'Escape') {
      event.preventDefault()
      event.stopPropagation()
      onCancel()
      return
    }

    if (event.key === 'Enter' && (event.metaKey || event.ctrlKey)) {
      event.preventDefault()
      event.stopPropagation()
      commit()
    }
  }

  return (
    <div
      className={`nano-text-editor ${block.className}`}
      onBlur={(event) => {
        if (!event.currentTarget.contains(event.relatedTarget as Node | null)) {
          commit()
        }
      }}
      onKeyDownCapture={handleKeyDown}
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
