import {
  createElement,
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type CSSProperties,
  type PointerEvent as ReactPointerEvent,
} from 'react'
import type { JSONPatchOperation, Pointer } from 'zod-crud'
import { useJSONDocument } from 'zod-crud/react'
import { PlainTextEditor } from './PlainTextEditor'
import {
  RESIZE_HANDLES,
  MIN_BLOCK_SIZE,
  SAMPLE_DECK,
  SAMPLE_SLIDES,
  SLIDE_HEIGHT,
  SLIDE_WIDTH,
  RetouchDeckSchema,
  blockLocationFromPointer,
  blockPointer,
  blockTextPointer,
  exportRetouchDeck,
  findBlockLocation,
  findSlideIndex,
  getRect,
  moveRect,
  rectEquals,
  rectToStyle,
  resizeRect,
  setLayoutPatch,
  setTextPatch,
  type Rect,
  type ResizeHandle,
  type RetouchSlide,
  type SlideBlock,
} from './retouchModel'
import './App.css'

type Mode = 'text' | 'layout'

type EditingState = {
  pointer: Pointer
}

type Interaction =
  | {
      kind: 'move'
      pointer: Pointer
      startClientPoint: Point
      startPoint: Point
      startRect: Rect
    }
  | {
      kind: 'resize'
      pointer: Pointer
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
  pointer: Pointer
  rect: Rect
}

type SnapGuides = {
  x: number | null
  y: number | null
}

const DRAG_THRESHOLD = 4

function App() {
  const doc = useJSONDocument(RetouchDeckSchema, SAMPLE_DECK, {
    history: 200,
    selection: { mode: 'extended' },
  })
  const slideRef = useRef<HTMLDivElement | null>(null)
  const stageRef = useRef<HTMLDivElement | null>(null)
  const exportTextareaRef = useRef<HTMLTextAreaElement | null>(null)

  const [mode, setMode] = useState<Mode>('text')
  const [activeSlideId, setActiveSlideId] = useState(SAMPLE_SLIDES[0].id)
  const [editing, setEditing] = useState<EditingState | null>(null)
  const [interaction, setInteraction] = useState<Interaction | null>(null)
  const [draftLayout, setDraftLayout] = useState<DraftLayout | null>(null)
  const [snapGuides, setSnapGuides] = useState<SnapGuides>({
    x: null,
    y: null,
  })
  const [exportOpen, setExportOpen] = useState(false)
  const [copiedExportCode, setCopiedExportCode] = useState<string | null>(null)

  const activeSlideIndex = Math.max(0, findSlideIndex(doc.value, activeSlideId))
  const activeSlide = doc.value.slides[activeSlideIndex] ?? doc.value.slides[0]
  const focusPointer = doc.selection?.focusPointer ?? null
  const selectedLocation = useMemo(
    () => blockLocationFromPointer(doc.value, focusPointer),
    [doc.value, focusPointer],
  )
  const selectedPointer =
    selectedLocation?.slide.id === activeSlide.id ? selectedLocation.pointer : null
  const selectedBlock =
    selectedLocation?.slide.id === activeSlide.id ? selectedLocation.block : null
  const selectedRect =
    selectedPointer && selectedBlock
      ? getCurrentRect(selectedPointer, selectedBlock, draftLayout)
      : null
  const exportCode = useMemo(() => exportRetouchDeck(doc.value), [doc.value])
  const exportCopied = exportOpen && copiedExportCode === exportCode
  const baseSelectedLocation =
    selectedBlock === null
      ? null
      : findBlockLocation(SAMPLE_DECK, activeSlide.id, selectedBlock.id)
  const canResetSelected = Boolean(
    selectedPointer &&
      selectedBlock &&
      baseSelectedLocation &&
      !rectEquals(getRect(selectedBlock), getRect(baseSelectedLocation.block)),
  )

  const commitPatch = useCallback(
    (
      patch: JSONPatchOperation[],
      pointer: Pointer,
      label: string,
      mergeKey?: string,
    ) => {
      if (patch.length === 0) {
        return
      }

      doc.commit(patch, {
        label,
        mergeKey,
        origin: 'ppt-retouch',
        selection: { type: 'collapse', pointer },
      })
    },
    [doc],
  )

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
        pointer: currentInteraction.pointer,
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

      commitPatch(
        setLayoutPatch(currentInteraction.pointer, rect),
        currentInteraction.pointer,
        `${currentInteraction.kind} layout`,
      )
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
    calculateInteractionRect,
    commitPatch,
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
    doc.selection?.empty()
    stageRef.current?.scrollTo({ left: 0, top: 0 })
    clearTransientState()
  }

  function selectBlock(pointer: Pointer) {
    doc.selection?.selectRanges([pointer])
  }

  function startTextEdit(pointer: Pointer) {
    if (mode !== 'text') {
      return
    }

    selectBlock(pointer)
    setEditing({ pointer })
  }

  function cancelTextEdit() {
    setEditing(null)
    setDraftLayout(null)
  }

  function commitTextEdit(pointer: Pointer, text: string, rect: Rect) {
    const location = blockLocationFromPointer(doc.value, pointer)
    setEditing(null)
    setDraftLayout(null)

    if (!location) {
      return
    }

    const currentRect = getRect(location.block)
    const patch = [
      ...(text === location.block.text ? [] : setTextPatch(pointer, text)),
      ...(rectEquals(rect, currentRect) ? [] : setLayoutPatch(pointer, rect)),
    ]

    commitPatch(
      patch,
      pointer,
      text === location.block.text ? 'resize text box' : 'edit text',
      text === location.block.text ? undefined : `text:${blockTextPointer(pointer)}`,
    )
  }

  function handleBlockPointerDown(
    event: ReactPointerEvent<HTMLElement>,
    pointer: Pointer,
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
    selectBlock(pointer)
    setInteraction({
      kind: 'move',
      pointer,
      startClientPoint: {
        x: event.clientX,
        y: event.clientY,
      },
      startPoint: point,
      startRect: getCurrentRect(pointer, block, draftLayout),
    })
  }

  function handleResizePointerDown(
    event: ReactPointerEvent<HTMLButtonElement>,
    handle: ResizeHandle,
  ) {
    if (!selectedBlock || !selectedPointer || !selectedRect || mode !== 'layout') {
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
      pointer: selectedPointer,
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
    if (
      mode !== 'layout' ||
      !selectedPointer ||
      !selectedBlock ||
      !baseSelectedLocation ||
      !canResetSelected
    ) {
      return
    }

    commitPatch(
      setLayoutPatch(selectedPointer, getRect(baseSelectedLocation.block)),
      selectedPointer,
      'reset layout',
    )
  }

  async function copyExportCode() {
    try {
      await navigator.clipboard.writeText(exportCode)
    } catch {
      exportTextareaRef.current?.select()
      document.execCommand('copy')
    }
    setCopiedExportCode(exportCode)
  }

  return (
    <main className="retouch-app" data-mode={mode}>
      <aside className="slide-rail" aria-label="Slides">
        {doc.value.slides.map((slide, index) => (
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
              disabled={!doc.history.canUndo}
              onClick={() => doc.history.undo()}
              type="button"
            >
              Undo
            </button>
            <button
              disabled={!doc.history.canRedo}
              onClick={() => doc.history.redo()}
              type="button"
            >
              Redo
            </button>
            <button
              disabled={mode !== 'layout' || !canResetSelected}
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
              doc.selection?.empty()
            }
          }}
          ref={stageRef}
        >
          <div className="slide-frame">
            <div
              className="slide-canvas"
              data-slide={activeSlide.id}
              ref={slideRef}
              style={{ '--accent': activeSlide.accent } as CSSProperties}
            >
              {activeSlide.blocks.map((block, blockIndex) => {
                const pointer = blockPointer(activeSlideIndex, blockIndex)
                const rect = getCurrentRect(pointer, block, draftLayout)
                const selected = pointer === selectedPointer
                const editingThisBlock =
                  mode === 'text' && editing?.pointer === pointer
                const className = [
                  'slide-block',
                  block.className,
                  selected ? 'is-selected' : '',
                  mode === 'text' ? 'is-text-mode' : 'is-layout-mode',
                ]
                  .filter(Boolean)
                  .join(' ')

                if (editingThisBlock) {
                  return (
                    <PlainTextEditor
                      block={block}
                      key={`${block.id}:editor`}
                      minimumHeight={
                        findBlockLocation(SAMPLE_DECK, activeSlide.id, block.id)
                          ?.block.height ?? MIN_BLOCK_SIZE
                      }
                      onCancel={cancelTextEdit}
                      onCommit={(text, rect) => commitTextEdit(pointer, text, rect)}
                      rect={rect}
                    />
                  )
                }

                return (
                  <SlideBlockElement
                    block={block}
                    className={className}
                    key={block.id}
                    onClick={() => {
                      if (mode === 'layout') {
                        selectBlock(pointer)
                      } else {
                        startTextEdit(pointer)
                      }
                    }}
                    onPointerDown={(event) =>
                      handleBlockPointerDown(event, pointer, block)
                    }
                    rect={rect}
                    selected={selected}
                    text={block.text}
                  />
                )
              })}

              {mode === 'layout' && selectedRect ? (
                <SelectionOverlay
                  onResizePointerDown={handleResizePointerDown}
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
            <div className="export-actions">
              <button onClick={copyExportCode} type="button">
                {exportCopied ? 'Copied' : 'Copy'}
              </button>
            </div>
            <textarea
              readOnly
              ref={exportTextareaRef}
              spellCheck={false}
              value={exportCode}
            />
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
          style={rectToStyle(getRect(block))}
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

function getCurrentRect(
  pointer: Pointer,
  block: SlideBlock,
  draftLayout: DraftLayout | null,
) {
  return draftLayout?.pointer === pointer ? draftLayout.rect : getRect(block)
}

function clamp(value: number, min: number, max: number) {
  return Math.min(max, Math.max(min, value))
}

export default App
