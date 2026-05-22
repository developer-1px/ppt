import {
  createElement,
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type CSSProperties,
  type KeyboardEvent as ReactKeyboardEvent,
  type PointerEvent as ReactPointerEvent,
  type RefObject,
} from 'react'
import {
  RESIZE_HANDLES,
  SAMPLE_SLIDES,
  SLIDE_HEIGHT,
  SLIDE_WIDTH,
  clamp,
  createInitialDeckEdits,
  exportRetouchDeck,
  getRect,
  getText,
  moveRect,
  rectToStyle,
  resetLayoutPatch,
  resizeRect,
  setLayoutPatch,
  setTextPatch,
  type DeckEdits,
  type Rect,
  type ResizeHandle,
  type RetouchSlide,
  type SlideBlock,
} from './retouchModel'
import './App.css'

type Mode = 'text' | 'layout'

type HistoryState = {
  past: DeckEdits[]
  present: DeckEdits
  future: DeckEdits[]
}

type EditingState = {
  blockId: string
  draft: string
}

type Interaction =
  | {
      kind: 'move'
      blockId: string
      startClientPoint: Point
      startPoint: Point
      startRect: Rect
    }
  | {
      kind: 'resize'
      blockId: string
      handle: ResizeHandle
      startClientPoint: Point
      startPoint: Point
      startRect: Rect
    }

type Point = {
  x: number
  y: number
}

type DraftLayout = {
  blockId: string
  rect: Rect
}

type SnapGuides = {
  x: number | null
  y: number | null
}

const DRAG_THRESHOLD = 4

const initialHistory: HistoryState = {
  past: [],
  present: createInitialDeckEdits(SAMPLE_SLIDES),
  future: [],
}

function App() {
  const slideRef = useRef<HTMLDivElement | null>(null)
  const editorRef = useRef<HTMLTextAreaElement | null>(null)

  const [mode, setMode] = useState<Mode>('text')
  const [activeSlideId, setActiveSlideId] = useState(SAMPLE_SLIDES[0].id)
  const [history, setHistory] = useState<HistoryState>(initialHistory)
  const [selectedBlockId, setSelectedBlockId] = useState<string | null>(null)
  const [editing, setEditing] = useState<EditingState | null>(null)
  const [interaction, setInteraction] = useState<Interaction | null>(null)
  const [draftLayout, setDraftLayout] = useState<DraftLayout | null>(null)
  const [snapGuides, setSnapGuides] = useState<SnapGuides>({
    x: null,
    y: null,
  })
  const [exportOpen, setExportOpen] = useState(false)

  const activeSlide = useMemo(
    () => SAMPLE_SLIDES.find((slide) => slide.id === activeSlideId) ?? SAMPLE_SLIDES[0],
    [activeSlideId],
  )
  const slideEdits = history.present[activeSlide.id]
  const selectedBlock = useMemo(
    () =>
      selectedBlockId
        ? activeSlide.blocks.find((block) => block.id === selectedBlockId) ?? null
        : null,
    [activeSlide.blocks, selectedBlockId],
  )
  const selectedRect = selectedBlock
    ? getCurrentRect(selectedBlock, slideEdits, draftLayout)
    : null
  const exportCode = useMemo(
    () => exportRetouchDeck(SAMPLE_SLIDES, history.present),
    [history.present],
  )
  const canResetSelected =
    selectedBlockId !== null &&
    slideEdits.layoutPatches[selectedBlockId] !== undefined

  const commitEdits = useCallback((updater: (edits: DeckEdits) => DeckEdits) => {
    setHistory((current) => {
      const next = updater(current.present)

      if (next === current.present) {
        return current
      }

      return {
        past: [...current.past, current.present],
        present: next,
        future: [],
      }
    })
  }, [])

  const undo = useCallback(() => {
    setHistory((current) => {
      const previous = current.past.at(-1)

      if (!previous) {
        return current
      }

      return {
        past: current.past.slice(0, -1),
        present: previous,
        future: [current.present, ...current.future],
      }
    })
  }, [])

  const redo = useCallback(() => {
    setHistory((current) => {
      const next = current.future[0]

      if (!next) {
        return current
      }

      return {
        past: [...current.past, current.present],
        present: next,
        future: current.future.slice(1),
      }
    })
  }, [])

  const readSlidePoint = useCallback((event: Pick<PointerEvent, 'clientX' | 'clientY'>) => {
    const rect = slideRef.current?.getBoundingClientRect()

    if (!rect) {
      return null
    }

    return {
      x: clamp(((event.clientX - rect.left) / rect.width) * SLIDE_WIDTH, 0, SLIDE_WIDTH),
      y: clamp(((event.clientY - rect.top) / rect.height) * SLIDE_HEIGHT, 0, SLIDE_HEIGHT),
    }
  }, [])

  const calculateInteractionRect = useCallback(
    (nextPoint: Point, currentInteraction: Interaction) => {
      const dx = nextPoint.x - currentInteraction.startPoint.x
      const dy = nextPoint.y - currentInteraction.startPoint.y

      if (currentInteraction.kind === 'move') {
        return moveRect(currentInteraction.startRect, dx, dy)
      }

      return resizeRect(
        currentInteraction.startRect,
        currentInteraction.handle,
        dx,
        dy,
      )
    },
    [],
  )

  const hasMeaningfulPointerDelta = useCallback((start: Point, next: Point) => {
    return (
      Math.abs(next.x - start.x) >= DRAG_THRESHOLD ||
      Math.abs(next.y - start.y) >= DRAG_THRESHOLD
    )
  }, [])

  useEffect(() => {
    if (!editing) {
      return
    }

    editorRef.current?.focus()
    editorRef.current?.select()
  }, [editing])

  useEffect(() => {
    if (!interaction) {
      return
    }

    const currentInteraction = interaction

    function handlePointerMove(event: PointerEvent) {
      const point = readSlidePoint(event)

      if (!point) {
        return
      }

      if (
        !hasMeaningfulPointerDelta(currentInteraction.startClientPoint, {
          x: event.clientX,
          y: event.clientY,
        })
      ) {
        return
      }

      const rect = calculateInteractionRect(point, currentInteraction)
      setDraftLayout({
        blockId: currentInteraction.blockId,
        rect,
      })
      setSnapGuides({
        x: rect.x,
        y: rect.y,
      })
    }

    function handlePointerUp(event: PointerEvent) {
      const point = readSlidePoint(event)

      if (
        !point ||
        !hasMeaningfulPointerDelta(currentInteraction.startClientPoint, {
          x: event.clientX,
          y: event.clientY,
        })
      ) {
        setInteraction(null)
        setDraftLayout(null)
        setSnapGuides({ x: null, y: null })
        return
      }

      const rect = calculateInteractionRect(point, currentInteraction)

      if (rectEquals(rect, currentInteraction.startRect)) {
        setInteraction(null)
        setDraftLayout(null)
        setSnapGuides({ x: null, y: null })
        return
      }

      commitEdits((edits) =>
        setLayoutPatch(edits, activeSlide.id, currentInteraction.blockId, rect),
      )
      setSelectedBlockId(currentInteraction.blockId)
      setInteraction(null)
      setDraftLayout(null)
      setSnapGuides({ x: null, y: null })
    }

    window.addEventListener('pointermove', handlePointerMove)
    window.addEventListener('pointerup', handlePointerUp)
    window.addEventListener('pointercancel', handlePointerUp)

    return () => {
      window.removeEventListener('pointermove', handlePointerMove)
      window.removeEventListener('pointerup', handlePointerUp)
      window.removeEventListener('pointercancel', handlePointerUp)
    }
  }, [
    activeSlide.id,
    calculateInteractionRect,
    commitEdits,
    hasMeaningfulPointerDelta,
    interaction,
    readSlidePoint,
  ])

  function clearTransientState() {
    setEditing(null)
    setInteraction(null)
    setDraftLayout(null)
    setSnapGuides({ x: null, y: null })
  }

  function changeMode(nextMode: Mode) {
    setMode(nextMode)
    clearTransientState()
  }

  function selectSlide(slideId: string) {
    setActiveSlideId(slideId)
    setSelectedBlockId(null)
    clearTransientState()
  }

  function startTextEdit(block: SlideBlock) {
    if (mode !== 'text') {
      return
    }

    setSelectedBlockId(block.id)
    setEditing({
      blockId: block.id,
      draft: getText(block, slideEdits),
    })
  }

  function commitTextEdit() {
    if (!editing) {
      return
    }

    const block = activeSlide.blocks.find((item) => item.id === editing.blockId)

    if (!block) {
      setEditing(null)
      return
    }

    const currentText = getText(block, slideEdits)
    const nextText = editing.draft
    setEditing(null)

    if (nextText === currentText) {
      return
    }

    commitEdits((edits) =>
      setTextPatch(edits, activeSlide.id, editing.blockId, nextText),
    )
  }

  function handleTextKeyDown(event: ReactKeyboardEvent<HTMLTextAreaElement>) {
    if (event.key === 'Escape') {
      event.preventDefault()
      setEditing(null)
      return
    }

    if (event.key === 'Enter' && (event.metaKey || event.ctrlKey)) {
      event.preventDefault()
      commitTextEdit()
    }
  }

  function handleBlockPointerDown(
    event: ReactPointerEvent<HTMLElement>,
    block: SlideBlock,
  ) {
    if (mode !== 'layout') {
      return
    }

    const point = readSlidePoint(event)

    if (!point) {
      return
    }

    event.preventDefault()
    event.stopPropagation()
    setSelectedBlockId(block.id)
    setInteraction({
      kind: 'move',
      blockId: block.id,
      startClientPoint: {
        x: event.clientX,
        y: event.clientY,
      },
      startPoint: point,
      startRect: getCurrentRect(block, slideEdits, draftLayout),
    })
  }

  function handleResizePointerDown(
    event: ReactPointerEvent<HTMLButtonElement>,
    handle: ResizeHandle,
  ) {
    if (!selectedBlock || !selectedRect || mode !== 'layout') {
      return
    }

    const point = readSlidePoint(event)

    if (!point) {
      return
    }

    event.preventDefault()
    event.stopPropagation()
    setInteraction({
      kind: 'resize',
      blockId: selectedBlock.id,
      handle,
      startClientPoint: {
        x: event.clientX,
        y: event.clientY,
      },
      startPoint: point,
      startRect: selectedRect,
    })
  }

  function resetSelectedLayout() {
    if (!selectedBlockId || !canResetSelected) {
      return
    }

    commitEdits((edits) =>
      resetLayoutPatch(edits, activeSlide.id, selectedBlockId),
    )
  }

  return (
    <main className="retouch-app" data-mode={mode}>
      <aside className="slide-rail" aria-label="Slides">
        {SAMPLE_SLIDES.map((slide, index) => (
          <button
            aria-current={slide.id === activeSlide.id ? 'page' : undefined}
            className="slide-thumb"
            key={slide.id}
            onClick={() => selectSlide(slide.id)}
            type="button"
          >
            <span className="thumb-number">{index + 1}</span>
            <MiniSlide slide={slide} />
            <span className="thumb-name">{slide.name}</span>
          </button>
        ))}
      </aside>

      <section className="retouch-workspace">
        <header className="topbar">
          <div className="mode-toggle" role="tablist" aria-label="Mode">
            <button
              aria-selected={mode === 'text'}
              className="mode-button"
              onClick={() => changeMode('text')}
              role="tab"
              type="button"
            >
              Text
            </button>
            <button
              aria-selected={mode === 'layout'}
              className="mode-button"
              onClick={() => changeMode('layout')}
              role="tab"
              type="button"
            >
              Layout
            </button>
          </div>

          <div className="toolbar" role="toolbar" aria-label="Actions">
            <button
              disabled={history.past.length === 0}
              onClick={undo}
              type="button"
            >
              Undo
            </button>
            <button
              disabled={history.future.length === 0}
              onClick={redo}
              type="button"
            >
              Redo
            </button>
            <button
              disabled={!canResetSelected}
              onClick={resetSelectedLayout}
              type="button"
            >
              Reset
            </button>
            <button
              aria-pressed={exportOpen}
              onClick={() => setExportOpen((open) => !open)}
              type="button"
            >
              Export
            </button>
          </div>
        </header>

        <div
          className="stage-shell"
          onClick={() => {
            if (mode === 'layout') {
              setSelectedBlockId(null)
            }
          }}
        >
          <div className="slide-frame">
            <div
              className="slide-canvas"
              data-slide={activeSlide.id}
              ref={slideRef}
              style={{ '--accent': activeSlide.accent } as CSSProperties}
            >
              {activeSlide.blocks.map((block) => {
                const rect = getCurrentRect(block, slideEdits, draftLayout)
                const selected = block.id === selectedBlockId
                const className = [
                  'slide-block',
                  block.className,
                  selected ? 'is-selected' : '',
                  mode === 'text' ? 'is-text-mode' : 'is-layout-mode',
                ]
                  .filter(Boolean)
                  .join(' ')

                return (
                  <SlideBlockElement
                    block={block}
                    className={className}
                    key={block.id}
                    onClick={() => startTextEdit(block)}
                    onPointerDown={(event) =>
                      handleBlockPointerDown(event, block)
                    }
                    rect={rect}
                    selected={selected}
                    text={getText(block, slideEdits)}
                  />
                )
              })}

              {mode === 'layout' && selectedRect ? (
                <SelectionOverlay
                  onResizePointerDown={handleResizePointerDown}
                  rect={selectedRect}
                />
              ) : null}

              {mode === 'text' && editing ? (
                <TextEditor
                  block={selectedBlock}
                  draft={editing.draft}
                  editorRef={editorRef}
                  onBlur={commitTextEdit}
                  onChange={(draft) =>
                    setEditing((current) =>
                      current ? { ...current, draft } : current,
                    )
                  }
                  onKeyDown={handleTextKeyDown}
                  rect={selectedRect}
                />
              ) : null}

              {mode === 'layout' && snapGuides.x !== null ? (
                <div
                  className="snap-guide snap-guide-x"
                  style={{ left: `${(snapGuides.x / SLIDE_WIDTH) * 100}%` }}
                />
              ) : null}
              {mode === 'layout' && snapGuides.y !== null ? (
                <div
                  className="snap-guide snap-guide-y"
                  style={{ top: `${(snapGuides.y / SLIDE_HEIGHT) * 100}%` }}
                />
              ) : null}
            </div>
          </div>
        </div>

        {exportOpen ? (
          <section className="export-panel" aria-label="Export">
            <textarea readOnly spellCheck={false} value={exportCode} />
          </section>
        ) : null}
      </section>
    </main>
  )
}

function SlideBlockElement({
  block,
  className,
  onClick,
  onPointerDown,
  rect,
  selected,
  text,
}: {
  block: SlideBlock
  className: string
  onClick: () => void
  onPointerDown: (event: ReactPointerEvent<HTMLElement>) => void
  rect: Rect
  selected: boolean
  text: string
}) {
  return createElement(
    block.tag,
    {
      'data-block': block.id,
      'data-role': block.role,
      'data-selected': selected ? 'true' : 'false',
      className,
      onClick: (event: MouseEvent) => {
        event.stopPropagation()
        onClick()
      },
      onPointerDown,
      style: rectToStyle(rect),
      tabIndex: 0,
    },
    text,
  )
}

function MiniSlide({ slide }: { slide: RetouchSlide }) {
  return (
    <span className="mini-slide" style={{ '--accent': slide.accent } as CSSProperties}>
      {slide.blocks.map((block) => (
        <span
          className={`mini-block ${block.role}`}
          key={block.id}
          style={rectToStyle({
            x: block.x,
            y: block.y,
            width: block.width,
            height: block.height,
          })}
        />
      ))}
    </span>
  )
}

function SelectionOverlay({
  onResizePointerDown,
  rect,
}: {
  onResizePointerDown: (
    event: ReactPointerEvent<HTMLButtonElement>,
    handle: ResizeHandle,
  ) => void
  rect: Rect
}) {
  return (
    <div className="selection-overlay" style={rectToStyle(rect)}>
      {RESIZE_HANDLES.map((handle) => (
        <button
          aria-label={`Resize ${handle}`}
          className="resize-handle"
          data-handle={handle}
          key={handle}
          onClick={(event) => event.stopPropagation()}
          onPointerDown={(event) => onResizePointerDown(event, handle)}
          type="button"
        />
      ))}
    </div>
  )
}

function TextEditor({
  block,
  draft,
  editorRef,
  onBlur,
  onChange,
  onKeyDown,
  rect,
}: {
  block: SlideBlock | null
  draft: string
  editorRef: RefObject<HTMLTextAreaElement | null>
  onBlur: () => void
  onChange: (draft: string) => void
  onKeyDown: (event: ReactKeyboardEvent<HTMLTextAreaElement>) => void
  rect: Rect | null
}) {
  if (!block || !rect) {
    return null
  }

  return (
    <textarea
      aria-label="Text"
      className={`text-editor ${block.className}`}
      onBlur={onBlur}
      onChange={(event) => onChange(event.target.value)}
      onKeyDown={onKeyDown}
      onPointerDown={(event) => event.stopPropagation()}
      ref={editorRef}
      spellCheck={false}
      style={rectToStyle(rect)}
      value={draft}
    />
  )
}

function getCurrentRect(
  block: SlideBlock,
  slideEdits: ReturnType<typeof createInitialDeckEdits>[string],
  draftLayout: DraftLayout | null,
) {
  return draftLayout?.blockId === block.id
    ? draftLayout.rect
    : getRect(block, slideEdits)
}

function rectEquals(a: Rect, b: Rect) {
  return (
    a.x === b.x &&
    a.y === b.y &&
    a.width === b.width &&
    a.height === b.height
  )
}

export default App
