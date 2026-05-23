import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  type CSSProperties,
  type ClipboardEvent as ReactClipboardEvent,
  type FormEvent as ReactFormEvent,
  type KeyboardEvent as ReactKeyboardEvent,
  type MouseEvent as ReactMouseEvent,
  type PointerEvent as ReactPointerEvent,
} from 'react'
import { Check, Code2, Download, Redo2, RotateCcw, Undo2 } from 'lucide-react'
import type { JSONPatchOperation, Pointer } from 'zod-crud'
import { useJSONDocument } from 'zod-crud/react'
import {
  RESIZE_HANDLES,
  SAMPLE_DECK,
  SAMPLE_SLIDES,
  SLIDE_HEIGHT,
  SLIDE_WIDTH,
  EMPTY_TEXT_BOX_HEIGHT,
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
  rectToAutoHeightStyle,
  rectToStyle,
  resizeRect,
  setArrangePatch,
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
  clientPoint?: Point
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

const DRAG_THRESHOLD = 8
const STORAGE_KEY = 'ppt-retouch:v1:deck'
const STORAGE_VERSION = 1

function App() {
  const initialDeck = useMemo(() => readInitialDeck(), [])
  const doc = useJSONDocument(RetouchDeckSchema, initialDeck, {
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
  const [copiedExportCode, setCopiedExportCode] = useState<string | null>(null)
  const [downloadedExportCode, setDownloadedExportCode] = useState<string | null>(null)
  const [visualSelectionRect, setVisualSelectionRect] = useState<Rect | null>(null)

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
  const exportCopied = copiedExportCode === exportCode
  const exportDownloaded = downloadedExportCode === exportCode
  const hasDeckChanges = !deckEquals(doc.value, SAMPLE_DECK)
  const baseSelectedLocation =
    selectedBlock === null
      ? null
      : findBlockLocation(SAMPLE_DECK, activeSlide.id, selectedBlock.id)
  const canResetSelected = Boolean(
      selectedPointer &&
      selectedBlock &&
      baseSelectedLocation &&
      !arrangeRectEquals(getRect(selectedBlock), getRect(baseSelectedLocation.block)),
  )
  const canReset = mode === 'layout' ? canResetSelected : hasDeckChanges
  const resetTitle = mode === 'layout' ? 'Reset layout' : 'Reset deck'

  useEffect(() => {
    persistDeck(doc.value)
  }, [doc.value])

  useLayoutEffect(() => {
    if (mode !== 'layout' || !selectedBlock || !slideRef.current) {
      setVisualSelectionRect(null)
      return
    }

    const slideBox = slideRef.current.getBoundingClientRect()
    const block = slideRef.current.querySelector<HTMLElement>(
      `[data-block="${selectedBlock.id}"]`,
    )

    if (!block || slideBox.width === 0 || slideBox.height === 0) {
      setVisualSelectionRect(null)
      return
    }

    const blockBox = block.getBoundingClientRect()
    const nextRect = {
      x: ((blockBox.left - slideBox.left) / slideBox.width) * SLIDE_WIDTH,
      y: ((blockBox.top - slideBox.top) / slideBox.height) * SLIDE_HEIGHT,
      width: (blockBox.width / slideBox.width) * SLIDE_WIDTH,
      height: (blockBox.height / slideBox.height) * SLIDE_HEIGHT,
    }

    setVisualSelectionRect((currentRect) =>
      currentRect && rectClose(currentRect, nextRect) ? currentRect : nextRect,
    )
  }, [activeSlide.id, doc.value, draftLayout, mode, selectedBlock])

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
      setSnapGuides(guidesForInteraction(rect, currentInteraction))
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
        setArrangePatch(currentInteraction.pointer, rect),
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

  useEffect(() => {
    function handleKeyDown(event: KeyboardEvent) {
      if (
        event.defaultPrevented ||
        !isHistoryShortcut(event) ||
        isEditableTarget(event.target)
      ) {
        return
      }

      const key = event.key.toLowerCase()

      if (key === 'z' && event.shiftKey && doc.history.canRedo) {
        event.preventDefault()
        doc.history.redo()
        return
      }

      if (key === 'z' && doc.history.canUndo) {
        event.preventDefault()
        doc.history.undo()
        return
      }

      if (key === 'y' && doc.history.canRedo) {
        event.preventDefault()
        doc.history.redo()
      }
    }

    window.addEventListener('keydown', handleKeyDown)

    return () => {
      window.removeEventListener('keydown', handleKeyDown)
    }
  }, [doc.history])

  useEffect(() => {
    function handleFocusedBlockEditKey(event: KeyboardEvent) {
      if (
        mode !== 'text' ||
        event.defaultPrevented ||
        isEditableTarget(event.target) ||
        (event.key !== 'Enter' && event.key !== 'F2')
      ) {
        return
      }

      const activeElement = document.activeElement

      if (!(activeElement instanceof HTMLElement)) {
        return
      }

      const blockElement = activeElement.closest<HTMLElement>('[data-block]')
      const blockId = blockElement?.dataset.block

      if (!blockId) {
        return
      }

      const location = findBlockLocation(doc.value, activeSlide.id, blockId)

      if (!location) {
        return
      }

      event.preventDefault()
      event.stopPropagation()
      doc.selection?.selectRanges([location.pointer])
      setEditing({ pointer: location.pointer })
    }

    window.addEventListener('keydown', handleFocusedBlockEditKey)

    return () => {
      window.removeEventListener('keydown', handleFocusedBlockEditKey)
    }
  }, [activeSlide.id, doc.selection, doc.value, mode])

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

  function startTextEdit(pointer: Pointer, clientPoint?: Point) {
    if (mode !== 'text') {
      return
    }

    selectBlock(pointer)
    setEditing({ clientPoint, pointer })
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
    const textChanged = text !== location.block.text
    const layoutChanged =
      rect.x !== currentRect.x ||
      rect.y !== currentRect.y ||
      rect.width !== currentRect.width ||
      (textChanged && rect.height !== currentRect.height)
    const patch = [
      ...(textChanged ? setTextPatch(pointer, text) : []),
      ...(layoutChanged ? setLayoutPatch(pointer, rect) : []),
    ]

    commitPatch(
      patch,
      pointer,
      textChanged ? 'edit text' : 'resize text box',
      textChanged ? `text:${blockTextPointer(pointer)}` : undefined,
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
      setArrangePatch(selectedPointer, getRect(baseSelectedLocation.block)),
      selectedPointer,
      'reset layout',
    )
  }

  function resetDeck() {
    if (mode === 'layout' || !hasDeckChanges) {
      return
    }

    clearTransientState()
    doc.selection?.empty()
    setActiveSlideId(SAMPLE_SLIDES[0].id)
    setCopiedExportCode(null)
    setDownloadedExportCode(null)
    stageRef.current?.scrollTo({ left: 0, top: 0 })
    doc.reset(SAMPLE_DECK)
  }

  function resetCurrentTarget() {
    if (mode === 'layout') {
      resetSelectedLayout()
      return
    }

    resetDeck()
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

  function downloadExportCode() {
    const url = URL.createObjectURL(
      new Blob([exportCode], { type: 'text/html;charset=utf-8' }),
    )
    const anchor = document.createElement('a')

    anchor.href = url
    anchor.download = 'retouched-slides.html'
    anchor.rel = 'noopener'
    document.body.append(anchor)
    anchor.click()
    anchor.remove()
    URL.revokeObjectURL(url)
    setDownloadedExportCode(exportCode)
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
              Arrange
            </button>
          </div>

          <div className="toolbar" role="toolbar" aria-label="Actions">
            <button
              aria-label="Undo"
              disabled={!doc.history.canUndo}
              onClick={() => doc.history.undo()}
              title="Undo"
              type="button"
            >
              <Undo2 aria-hidden="true" size={16} strokeWidth={2.2} />
            </button>
            <button
              aria-label="Redo"
              disabled={!doc.history.canRedo}
              onClick={() => doc.history.redo()}
              title="Redo"
              type="button"
            >
              <Redo2 aria-hidden="true" size={16} strokeWidth={2.2} />
            </button>
            <button
              aria-label="Reset"
              disabled={!canReset}
              onClick={resetCurrentTarget}
              title={resetTitle}
              type="button"
            >
              <RotateCcw aria-hidden="true" size={16} strokeWidth={2.2} />
            </button>
            <button
              aria-label="Copy HTML"
              aria-pressed={exportCopied}
              data-copy-state={exportCopied ? 'copied' : 'idle'}
              onClick={copyExportCode}
              title={exportCopied ? 'Copied' : 'Copy HTML'}
              type="button"
            >
              {exportCopied ? (
                <Check aria-hidden="true" size={16} strokeWidth={2.4} />
              ) : (
                <Code2 aria-hidden="true" size={16} strokeWidth={2.2} />
              )}
            </button>
            <button
              aria-label="Download HTML"
              aria-pressed={exportDownloaded}
              data-download-state={exportDownloaded ? 'downloaded' : 'idle'}
              onClick={downloadExportCode}
              title={exportDownloaded ? 'Downloaded' : 'Download HTML'}
              type="button"
            >
              {exportDownloaded ? (
                <Check aria-hidden="true" size={16} strokeWidth={2.4} />
              ) : (
                <Download aria-hidden="true" size={16} strokeWidth={2.2} />
              )}
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
                const minimumHeight =
                  block.text.length === 0 ? EMPTY_TEXT_BOX_HEIGHT : 0
                const editingThisBlock =
                  mode === 'text' && editing?.pointer === pointer
                const className = [
                  'slide-block',
                  block.className,
                  selected ? 'is-selected' : '',
                  editingThisBlock ? 'is-editing' : '',
                  mode === 'text' ? 'is-text-mode' : 'is-layout-mode',
                ]
                  .filter(Boolean)
                  .join(' ')

                return (
                  <SlideBlockElement
                    block={block}
                    className={className}
                    editing={editingThisBlock}
                    initialClientPoint={
                      editingThisBlock ? editing.clientPoint : undefined
                    }
                    key={block.id}
                    minimumHeight={minimumHeight}
                    onCancel={cancelTextEdit}
                    onClick={(event) => {
                      if (mode === 'layout') {
                        selectBlock(pointer)
                      } else {
                        startTextEdit(pointer, {
                          x: event.clientX,
                          y: event.clientY,
                        })
                      }
                    }}
                    onPointerDown={(event) =>
                      handleBlockPointerDown(event, pointer, block)
                    }
                    onCommit={(text, nextRect) =>
                      commitTextEdit(pointer, text, nextRect)
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
                  rect={visualSelectionRect ?? selectedRect}
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

        <textarea
          aria-hidden="true"
          className="export-buffer"
          readOnly
          ref={exportTextareaRef}
          tabIndex={-1}
          value={exportCode}
        />
      </section>
    </main>
  )
}

function SlideBlockElement({
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
}: {
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
}) {
  const rectRef = useRef(rect)
  const committedRef = useRef(false)

  useEffect(() => {
    rectRef.current = rect
  }, [rect])

  useEffect(() => {
    if (!editing) {
      committedRef.current = false
      return
    }

    const element = readBlockElement(block.id)

    if (!element) {
      return
    }

    committedRef.current = false
    element.focus()
    if (
      !initialClientPoint ||
      !placeCaretFromPoint(element, initialClientPoint.x, initialClientPoint.y)
    ) {
      placeCaretAtEnd(element)
    }
  }, [block.id, editing, initialClientPoint])

  const syncAutoHeight = useCallback((element: HTMLElement) => {
    const effectiveMinimumHeight =
      element.textContent?.length === 0 ? EMPTY_TEXT_BOX_HEIGHT : minimumHeight
    const nextRect = autoHeightRect(element, rectRef.current, effectiveMinimumHeight)

    rectRef.current = nextRect
    applyAutoHeightStyle(element, nextRect, effectiveMinimumHeight)

    return nextRect
  }, [minimumHeight])

  const commit = useCallback(() => {
    if (!editing || committedRef.current) {
      return
    }

    const element = readBlockElement(block.id)
    const nextRect = element ? syncAutoHeight(element) : rectRef.current

    committedRef.current = true
    element?.blur()
    onCommit(element?.textContent ?? text, nextRect)
  }, [block.id, editing, onCommit, syncAutoHeight, text])

  function resetDraft() {
    const element = readBlockElement(block.id)

    if (!element) {
      return
    }

    const effectiveMinimumHeight =
      text.length === 0 ? EMPTY_TEXT_BOX_HEIGHT : minimumHeight
    element.textContent = text
    rectRef.current = rect
    applyAutoHeightStyle(element, rect, effectiveMinimumHeight)
  }

  function handleKeyDown(event: ReactKeyboardEvent<HTMLElement>) {
    if (event.nativeEvent.isComposing) {
      return
    }

    if (event.key === 'Escape') {
      event.preventDefault()
      event.stopPropagation()
      resetDraft()
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

  function handlePaste(event: ReactClipboardEvent<HTMLElement>) {
    event.preventDefault()
    document.execCommand('insertText', false, event.clipboardData.getData('text/plain'))
  }

  const sharedProps = {
    'data-block': block.id,
    'data-editing': editing ? 'true' : undefined,
    'data-empty': text.length === 0 ? 'true' : undefined,
    'data-role': block.role,
    'data-selected': selected ? 'true' : 'false',
    className,
    contentEditable: editing ? ('plaintext-only' as const) : undefined,
    onBlur: editing ? commit : undefined,
    onClick: (event: ReactMouseEvent<HTMLElement>) => {
      event.stopPropagation()
      if (!editing) {
        onClick(event)
      }
    },
    onInput: editing
      ? (event: ReactFormEvent<HTMLElement>) => syncAutoHeight(event.currentTarget)
      : undefined,
    onKeyDown: editing ? handleKeyDown : undefined,
    onPaste: editing ? handlePaste : undefined,
    onPointerDown: (event: ReactPointerEvent<HTMLElement>) => {
      if (editing) {
        event.stopPropagation()
        return
      }

      onPointerDown(event)
    },
    spellCheck: editing ? false : undefined,
    style: rectToAutoHeightStyle(rect, minimumHeight),
    suppressContentEditableWarning: editing ? true : undefined,
    tabIndex: 0,
  }

  if (block.tag === 'h1') {
    return <h1 {...sharedProps}>{text}</h1>
  }

  if (block.tag === 'p') {
    return <p {...sharedProps}>{text}</p>
  }

  return <div {...sharedProps}>{text}</div>
}

function readBlockElement(blockId: string) {
  return document.querySelector<HTMLElement>(`[data-block="${blockId}"]`)
}

function applyAutoHeightStyle(
  element: HTMLElement,
  rect: Rect,
  minimumHeight: number,
) {
  element.style.left = `${(rect.x / SLIDE_WIDTH) * 100}%`
  element.style.top = `${(rect.y / SLIDE_HEIGHT) * 100}%`
  element.style.width = `${(rect.width / SLIDE_WIDTH) * 100}%`
  element.style.height = 'auto'
  element.style.minHeight =
    minimumHeight > 0 ? `${(minimumHeight / SLIDE_HEIGHT) * 100}%` : ''
}

function autoHeightRect(element: HTMLElement, rect: Rect, minimumHeight: number): Rect {
  const elementBox = element.getBoundingClientRect()

  if (elementBox.width === 0) {
    return rect
  }

  const slideUnitsPerCssPixel = rect.width / elementBox.width
  const contentHeight = elementBox.height * slideUnitsPerCssPixel
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

function guidesForInteraction(rect: Rect, interaction: Interaction): SnapGuides {
  if (interaction.kind === 'move') {
    return { x: rect.x, y: rect.y }
  }

  return {
    x: interaction.handle === 'e' ? rect.x + rect.width : rect.x,
    y: null,
  }
}

function clamp(value: number, min: number, max: number) {
  return Math.min(max, Math.max(min, value))
}

function rectClose(a: Rect, b: Rect) {
  return (
    Math.abs(a.x - b.x) < 0.5 &&
    Math.abs(a.y - b.y) < 0.5 &&
    Math.abs(a.width - b.width) < 0.5 &&
    Math.abs(a.height - b.height) < 0.5
  )
}

function arrangeRectEquals(a: Rect, b: Rect) {
  return a.x === b.x && a.y === b.y && a.width === b.width
}

function deckEquals(a: unknown, b: unknown) {
  const parsedA = RetouchDeckSchema.safeParse(a)
  const parsedB = RetouchDeckSchema.safeParse(b)

  if (!parsedA.success || !parsedB.success) {
    return false
  }

  return JSON.stringify(parsedA.data) === JSON.stringify(parsedB.data)
}

function isHistoryShortcut(event: KeyboardEvent) {
  return (event.metaKey || event.ctrlKey) && !event.altKey
}

function isEditableTarget(target: EventTarget | null) {
  if (!(target instanceof HTMLElement)) {
    return false
  }

  return Boolean(target.closest('[contenteditable], textarea, input, select'))
}

function readInitialDeck() {
  if (typeof window === 'undefined') {
    return SAMPLE_DECK
  }

  try {
    const raw = window.localStorage.getItem(STORAGE_KEY)

    if (!raw) {
      return SAMPLE_DECK
    }

    const payload: unknown = JSON.parse(raw)

    if (
      !payload ||
      typeof payload !== 'object' ||
      !('version' in payload) ||
      payload.version !== STORAGE_VERSION ||
      !('deck' in payload)
    ) {
      window.localStorage.removeItem(STORAGE_KEY)
      return SAMPLE_DECK
    }

    const parsed = RetouchDeckSchema.safeParse(payload.deck)

    if (!parsed.success) {
      window.localStorage.removeItem(STORAGE_KEY)
      return SAMPLE_DECK
    }

    return parsed.data
  } catch {
    window.localStorage.removeItem(STORAGE_KEY)
    return SAMPLE_DECK
  }
}

function persistDeck(deck: unknown) {
  if (typeof window === 'undefined') {
    return
  }

  try {
    window.localStorage.setItem(
      STORAGE_KEY,
      JSON.stringify({
        version: STORAGE_VERSION,
        deck,
      }),
    )
  } catch {
    // Autosave is best-effort; editing must keep working if storage is unavailable.
  }
}

export default App
