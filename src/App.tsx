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
import type { JSONPatchOperation, Pointer, SelectionAction } from 'zod-crud'
import { useJSONDocument } from 'zod-crud/react'
import {
  RESIZE_HANDLES,
  GRID_SIZE,
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
      pointers: Pointer[]
      startClientPoint: Point
      startPoint: Point
      startRect: Rect
      startRects: DraftLayoutRect[]
    }
  | {
      kind: 'resize'
      pointer: Pointer
      handle: ResizeHandle
      startClientPoint: Point
      startPoint: Point
      startRect: Rect
      startRects: DraftLayoutRect[]
    }

type Point = {
  x: number
  y: number
}

type DraftLayout = {
  rects: DraftLayoutRect[]
}

type DraftLayoutRect = {
  pointer: Pointer
  rect: Rect
}

type SnapGuides = {
  x: number | null
  y: number | null
}

const DRAG_THRESHOLD = 8
const STORAGE_KEY = 'ppt-retouch:v3:deck'
const STORAGE_VERSION = 1
const CARET_PLACEHOLDER = '\u200B'

function App() {
  const initialDeck = useMemo(() => readInitialDeck(), [])
  const doc = useJSONDocument(RetouchDeckSchema, initialDeck, {
    history: 200,
    selection: { mode: 'extended' },
  })
  const slideRef = useRef<HTMLDivElement | null>(null)
  const stageRef = useRef<HTMLDivElement | null>(null)
  const exportTextareaRef = useRef<HTMLTextAreaElement | null>(null)
  const previousExportCodeRef = useRef<string | null>(null)
  const suppressBlockClickRef = useRef(false)
  const suppressStageClickRef = useRef(false)

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
  const [failedCopyExportCode, setFailedCopyExportCode] = useState<string | null>(null)
  const [visualSelectionRect, setVisualSelectionRect] = useState<Rect | null>(null)

  const activeSlideIndex = Math.max(0, findSlideIndex(doc.value, activeSlideId))
  const activeSlide = doc.value.slides[activeSlideIndex] ?? doc.value.slides[0]
  const selectedPointers = (doc.selection?.selectedPointers ?? []).filter((pointer) => {
    const location = blockLocationFromPointer(doc.value, pointer)

    return location?.slide.id === activeSlide.id
  })
  const selectedPointerSet = new Set(selectedPointers)
  const focusPointer = doc.selection?.focusPointer ?? null
  const primaryPointer =
    focusPointer && selectedPointerSet.has(focusPointer)
      ? focusPointer
      : (selectedPointers.at(-1) ?? null)
  const selectedLocation = useMemo(
    () => blockLocationFromPointer(doc.value, primaryPointer),
    [doc.value, primaryPointer],
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
  const exportStatusMatchesVisibleSlide = !editing && !interaction && !draftLayout
  const exportCopied = exportStatusMatchesVisibleSlide && copiedExportCode === exportCode
  const exportDownloaded =
    exportStatusMatchesVisibleSlide && downloadedExportCode === exportCode
  const exportCopyFailed =
    exportStatusMatchesVisibleSlide && failedCopyExportCode === exportCode
  const copyState = exportCopied
    ? 'copied'
    : exportCopyFailed
      ? 'failed'
      : 'idle'
  const copyTitle =
    copyState === 'copied'
      ? 'Copied'
      : copyState === 'failed'
        ? 'Copy failed'
        : 'Copy HTML'
  const hasDeckChanges = !deckEquals(doc.value, SAMPLE_DECK)
  const changedSlideIds = useMemo(() => changedSlides(doc.value), [doc.value])
  const baseSelectedLocation =
    selectedBlock === null
      ? null
      : findBlockLocation(SAMPLE_DECK, activeSlide.id, selectedBlock.id)
  const canResetSelectedLayout = Boolean(
    selectedPointer &&
      selectedBlock &&
      baseSelectedLocation &&
      !arrangeResetEquals(selectedBlock, baseSelectedLocation.block),
  )
  const canResetSelectedText = Boolean(
    selectedPointer &&
      selectedBlock &&
      baseSelectedLocation &&
      (!textResetEquals(selectedBlock, baseSelectedLocation.block) ||
        editing?.pointer === selectedPointer),
  )
  const canResetDeck = !selectedPointer && hasDeckChanges
  const canReset =
    mode === 'layout'
      ? canResetSelectedLayout
      : canResetSelectedText || canResetDeck
  const resetScope =
    mode === 'layout' ? 'layout' : selectedPointer ? 'text' : 'deck'
  const resetTitle =
    resetScope === 'layout'
      ? 'Reset layout'
      : resetScope === 'text'
        ? 'Reset text'
        : 'Reset deck'

  useEffect(() => {
    persistDeck(doc.value)
  }, [doc.value])

  useEffect(() => {
    if (previousExportCodeRef.current === null) {
      previousExportCodeRef.current = exportCode
      return
    }

    if (previousExportCodeRef.current === exportCode) {
      return
    }

    previousExportCodeRef.current = exportCode
    setFailedCopyExportCode(null)
  }, [exportCode])

  useLayoutEffect(() => {
    if (
      mode !== 'layout' ||
      selectedPointers.length !== 1 ||
      !selectedBlock ||
      !slideRef.current
    ) {
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
  }, [activeSlide.id, doc.value, draftLayout, mode, selectedBlock, selectedPointers.length])

  const commitPatch = useCallback(
    (
      patch: JSONPatchOperation[],
      pointer: Pointer,
      label: string,
      mergeKey?: string,
      selection?: SelectionAction,
    ) => {
      if (patch.length === 0) {
        return
      }

      setFailedCopyExportCode(null)
      doc.commit(patch, {
        label,
        mergeKey,
        origin: 'ppt-retouch',
        selection: selection ?? { type: 'collapse', pointer },
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

  const calculateInteractionState = useCallback(
    (nextPoint: Point, currentInteraction: Interaction) => {
      const dx = nextPoint.x - currentInteraction.startPoint.x
      const dy = nextPoint.y - currentInteraction.startPoint.y

      if (currentInteraction.kind === 'move') {
        return calculateMoveInteractionState(
          dx,
          dy,
          currentInteraction,
          activeSlide.blocks,
          activeSlideIndex,
        )
      }

      const rect = resizeRect(
        currentInteraction.startRect,
        currentInteraction.handle,
        dx,
        dy,
      )

      return {
        guides: guidesForInteraction(rect, currentInteraction),
        rects: [{ pointer: currentInteraction.pointer, rect }],
      }
    },
    [activeSlide.blocks, activeSlideIndex],
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

      const { guides, rects } = calculateInteractionState(point, currentInteraction)
      setDraftLayout({
        rects,
      })
      setSnapGuides(guides)
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
        suppressStageClickRef.current = true
        setInteraction(null)
        setDraftLayout(null)
        setSnapGuides({ x: null, y: null })
        return
      }

      const { rects } = calculateInteractionState(point, currentInteraction)

      if (draftRectsEqual(rects, currentInteraction.startRects)) {
        suppressStageClickRef.current = true
        setInteraction(null)
        setDraftLayout(null)
        setSnapGuides({ x: null, y: null })
        return
      }

      suppressBlockClickRef.current = true
      window.setTimeout(() => {
        suppressBlockClickRef.current = false
      }, 0)
      const selection =
        currentInteraction.kind === 'move'
          ? selectionActionForPointers(
              currentInteraction.pointers,
              currentInteraction.pointer,
            )
          : undefined

      suppressStageClickRef.current = true
      commitPatch(
        rects.flatMap(({ pointer, rect }) =>
          setArrangePatch(pointer, rect, {
            includeHeight:
              currentInteraction.kind === 'resize' &&
              resizeHandleAffectsHeight(currentInteraction.handle),
          }),
        ),
        currentInteraction.pointer,
        `${currentInteraction.kind} layout`,
        undefined,
        selection,
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
    calculateInteractionState,
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

  useEffect(() => {
    function handleTextSelectionKey(event: KeyboardEvent) {
      if (
        mode !== 'text' ||
        event.defaultPrevented ||
        editing ||
        !selectedPointer ||
        isEditableTarget(event.target) ||
        isControlTarget(event.target) ||
        event.key !== 'Escape'
      ) {
        return
      }

      event.preventDefault()
      event.stopPropagation()
      doc.selection?.empty()
    }

    window.addEventListener('keydown', handleTextSelectionKey)

    return () => {
      window.removeEventListener('keydown', handleTextSelectionKey)
    }
  }, [doc.selection, editing, mode, selectedPointer])

  useEffect(() => {
    function handleLayoutKey(event: KeyboardEvent) {
      if (
        mode !== 'layout' ||
        event.defaultPrevented ||
        editing ||
        interaction ||
        selectedPointers.length === 0 ||
        isEditableTarget(event.target) ||
        isControlTarget(event.target)
      ) {
        return
      }

      if (event.key === 'Escape') {
        event.preventDefault()
        event.stopPropagation()
        doc.selection?.empty()
        return
      }

      const delta = arrowKeyDelta(event.key, event.shiftKey)

      if (!delta) {
        return
      }

      const targets = selectedPointers
        .map((pointer) => {
          const location = blockLocationFromPointer(doc.value, pointer)

          if (!location || location.slide.id !== activeSlide.id) {
            return null
          }

          return {
            pointer,
            rect: moveRect(getRect(location.block), delta.x, delta.y),
            startRect: getRect(location.block),
          }
        })
        .filter(
          (
            target,
          ): target is { pointer: Pointer; rect: Rect; startRect: Rect } =>
            target !== null,
        )

      if (
        targets.length === 0 ||
        targets.every((target) => rectEquals(target.rect, target.startRect))
      ) {
        return
      }

      event.preventDefault()
      event.stopPropagation()
      commitPatch(
        targets.flatMap((target) => setArrangePatch(target.pointer, target.rect)),
        targets.at(-1)?.pointer ?? selectedPointer ?? targets[0].pointer,
        'nudge layout',
        `layout:nudge:${targets.map((target) => target.pointer).join('|')}`,
        selectionActionForPointers(
          targets.map((target) => target.pointer),
          targets.at(-1)?.pointer,
        ),
      )
    }

    window.addEventListener('keydown', handleLayoutKey)

    return () => {
      window.removeEventListener('keydown', handleLayoutKey)
    }
  }, [
    commitPatch,
    doc.selection,
    editing,
    interaction,
    mode,
    activeSlide.id,
    doc.value,
    selectedPointers,
    selectedPointer,
  ])

  function clearTransientState() {
    setEditing(null)
    setInteraction(null)
    setDraftLayout(null)
    setSnapGuides({ x: null, y: null })
  }

  function changeMode(nextMode: Mode) {
    if (mode === 'text') {
      commitActiveTextEdit()
    }

    if (mode !== 'text' && nextMode === 'text') {
      doc.selection?.empty()
    }

    setMode(nextMode)
    clearTransientState()
  }

  function selectSlide(slideId: string) {
    commitActiveTextEdit()
    setActiveSlideId(slideId)
    doc.selection?.empty()
    stageRef.current?.scrollTo({ left: 0, top: 0 })
    clearTransientState()
  }

  function selectBlock(pointer: Pointer, additive = false) {
    if (additive) {
      doc.selection?.togglePointer(pointer)
      return
    }

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
    setEditing(null)
    setDraftLayout(null)
    commitTextPatch(pointer, text, rect)
  }

  function commitActiveTextEdit() {
    if (!editing) {
      return doc.value
    }

    const activeEditing = editing
    const location = blockLocationFromPointer(doc.value, activeEditing.pointer)
    const element = slideRef.current?.querySelector<HTMLElement>(
      '[data-editing="true"][contenteditable]',
    )
    const blockElement = element?.closest<HTMLElement>('[data-block]')

    setEditing(null)
    setDraftLayout(null)

    if (!location || !element || !blockElement) {
      return doc.value
    }

    const text = normalizeEditableText(element.textContent ?? '')
    const baseBlock = findBlockLocation(
      SAMPLE_DECK,
      location.slide.id,
      location.block.id,
    )?.block
    const minimumHeight = minimumHeightForBlock(
      location.block,
      baseBlock,
      getRect(location.block),
    )
    const rect = autoHeightRect(blockElement, getRect(location.block), minimumHeight)

    return commitTextPatch(activeEditing.pointer, text, rect)
  }

  function commitTextPatch(pointer: Pointer, text: string, rect: Rect) {
    const location = blockLocationFromPointer(doc.value, pointer)

    if (!location) {
      return doc.value
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

    return doc.value
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

    if (hasSelectionModifier(event)) {
      return
    }

    event.preventDefault()
    event.stopPropagation()
    const activePointers = selectedPointerSet.has(pointer)
      ? selectedPointers
      : [pointer]
    const startRects = activePointers
      .map((activePointer) => {
        const location = blockLocationFromPointer(doc.value, activePointer)

        if (!location || location.slide.id !== activeSlide.id) {
          return null
        }

        return {
          pointer: activePointer,
          rect: getCurrentRect(activePointer, location.block, draftLayout),
        }
      })
      .filter((draftRect): draftRect is DraftLayoutRect => draftRect !== null)

    if (!selectedPointerSet.has(pointer)) {
      selectBlock(pointer)
    }

    setInteraction({
      kind: 'move',
      pointer,
      pointers: startRects.map((draftRect) => draftRect.pointer),
      startClientPoint: {
        x: event.clientX,
        y: event.clientY,
      },
      startPoint: point,
      startRect: getCurrentRect(pointer, block, draftLayout),
      startRects,
    })
  }

  function handleResizePointerDown(
    event: ReactPointerEvent<HTMLButtonElement>,
    handle: ResizeHandle,
  ) {
    if (
      selectedPointers.length !== 1 ||
      !selectedBlock ||
      !selectedPointer ||
      !selectedRect ||
      mode !== 'layout'
    ) {
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
      startRects: [{ pointer: selectedPointer, rect: selectedRect }],
    })
  }

  function resetSelectedLayout() {
    if (
      mode !== 'layout' ||
      !selectedPointer ||
      !selectedBlock ||
      !baseSelectedLocation ||
      !canResetSelectedLayout
    ) {
      return
    }

    commitPatch(
      setArrangePatch(selectedPointer, getRect(baseSelectedLocation.block), {
        includeHeight:
          selectedBlock.text === baseSelectedLocation.block.text &&
          selectedBlock.height !== baseSelectedLocation.block.height,
      }),
      selectedPointer,
      'reset layout',
    )
  }

  function resetSelectedText() {
    if (mode !== 'text') {
      return
    }

    const pointer = editing?.pointer ?? selectedPointer
    const location = pointer ? blockLocationFromPointer(doc.value, pointer) : null
    const baseLocation = location
      ? findBlockLocation(SAMPLE_DECK, location.slide.id, location.block.id)
      : null

    if (!pointer || !location || !baseLocation) {
      return
    }

    const element =
      editing?.pointer === pointer
        ? slideRef.current?.querySelector<HTMLElement>(
            '[data-editing="true"][contenteditable]',
          )
        : null
    const blockElement = element?.closest<HTMLElement>('[data-block]')
    const liveText = normalizeEditableText(
      element?.textContent ?? location.block.text,
    )
    const liveMinimumHeight =
      liveText.length === 0
        ? EMPTY_TEXT_BOX_HEIGHT
        : minimumHeightForBlock(
            location.block,
            baseLocation.block,
            getRect(location.block),
          )
    const liveRect = element && blockElement
      ? autoHeightRect(blockElement, getRect(location.block), liveMinimumHeight)
      : getRect(location.block)
    const resetRect = {
      ...liveRect,
      height: baseLocation.block.height,
    }

    if (element && liveText !== location.block.text) {
      commitTextPatch(pointer, liveText, liveRect)
    }

    setEditing(null)
    setDraftLayout(null)

    const patch = [
      ...(liveText !== baseLocation.block.text ||
      location.block.text !== baseLocation.block.text
        ? setTextPatch(pointer, baseLocation.block.text)
        : []),
      ...(liveRect.height !== resetRect.height
        ? setLayoutPatch(pointer, resetRect)
        : []),
    ]

    commitPatch(patch, pointer, 'reset text')
  }

  function resetDeck() {
    if (mode === 'layout' || !hasDeckChanges) {
      return
    }

    commitActiveTextEdit()
    clearTransientState()
    doc.selection?.empty()
    setActiveSlideId(SAMPLE_SLIDES[0].id)
    setCopiedExportCode(null)
    setDownloadedExportCode(null)
    setFailedCopyExportCode(null)
    stageRef.current?.scrollTo({ left: 0, top: 0 })
    doc.commit([{ op: 'replace', path: '', value: SAMPLE_DECK }], {
      label: 'reset deck',
      origin: 'ppt-retouch',
    })
  }

  function resetCurrentTarget() {
    if (mode === 'layout') {
      resetSelectedLayout()
      return
    }

    if (selectedPointer) {
      resetSelectedText()
      return
    }

    resetDeck()
  }

  function undoDocumentChange() {
    commitActiveTextEdit()
    doc.history.undo()
  }

  function redoDocumentChange() {
    commitActiveTextEdit()
    doc.history.redo()
  }

  async function copyExportCode() {
    const nextExportCode = exportRetouchDeck(commitActiveTextEdit())
    const copied = await writeExportToClipboard(nextExportCode, exportTextareaRef.current)

    if (!copied) {
      setCopiedExportCode(null)
      setFailedCopyExportCode(nextExportCode)
      return
    }

    setFailedCopyExportCode(null)
    setCopiedExportCode(nextExportCode)
  }

  function downloadExportCode() {
    const nextExportCode = exportRetouchDeck(commitActiveTextEdit())
    const url = URL.createObjectURL(
      new Blob([nextExportCode], { type: 'text/html;charset=utf-8' }),
    )
    const anchor = document.createElement('a')

    anchor.href = url
    anchor.download = 'retouched-slides.html'
    anchor.rel = 'noopener'
    document.body.append(anchor)
    anchor.click()
    anchor.remove()
    URL.revokeObjectURL(url)
    setFailedCopyExportCode(null)
    setDownloadedExportCode(nextExportCode)
  }

  return (
    <main className="retouch-app" data-mode={mode}>
      <aside className="slide-rail" aria-label="Slides">
        {doc.value.slides.map((slide, index) => {
          const changed = changedSlideIds.has(slide.id)

          return (
            <button
              aria-current={slide.id === activeSlide.id ? 'page' : undefined}
              aria-label={changed ? `${slide.name}, modified` : slide.name}
              className="slide-thumb"
              data-changed={changed ? 'true' : 'false'}
              key={slide.id}
              onClick={() => selectSlide(slide.id)}
              type="button"
            >
              <span className="thumb-number">{index + 1}</span>
              <MiniSlide slide={slide} />
              <span className="thumb-name">{slide.name}</span>
              {changed ? <span aria-hidden="true" className="thumb-change" /> : null}
            </button>
          )
        })}
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
              data-action="undo"
              disabled={!doc.history.canUndo}
              onClick={undoDocumentChange}
              title="Undo"
              type="button"
            >
              <Undo2 aria-hidden="true" size={16} strokeWidth={2.2} />
            </button>
            <button
              aria-label="Redo"
              data-action="redo"
              disabled={!doc.history.canRedo}
              onClick={redoDocumentChange}
              title="Redo"
              type="button"
            >
              <Redo2 aria-hidden="true" size={16} strokeWidth={2.2} />
            </button>
            <button
              aria-label={resetTitle}
              data-action="reset"
              data-reset-scope={resetScope}
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
              data-action="copy-html"
              data-copy-state={copyState}
              onClick={copyExportCode}
              title={copyTitle}
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
              data-action="download-html"
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
          onClick={(event) => {
            if (suppressStageClickRef.current) {
              suppressStageClickRef.current = false
              return
            }

            const target =
              event.target instanceof HTMLElement ? event.target : null

            if (target?.closest('[data-block], .selection-overlay, .resize-handle')) {
              return
            }

            if (mode === 'text') {
              commitActiveTextEdit()
            }
            doc.selection?.empty()
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
                const selected = selectedPointerSet.has(pointer)
                const baseBlock = findBlockLocation(
                  SAMPLE_DECK,
                  activeSlide.id,
                  block.id,
                )?.block
                const minimumHeight = minimumHeightForBlock(
                  block,
                  baseBlock,
                  rect,
                  interaction?.kind === 'resize' &&
                    interaction.pointer === pointer &&
                    resizeHandleAffectsHeight(interaction.handle),
                )
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
                      if (suppressBlockClickRef.current) {
                        suppressBlockClickRef.current = false
                        return
                      }

                      if (mode === 'layout') {
                        selectBlock(pointer, hasSelectionModifier(event))
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

              {mode === 'layout' && selectedPointers.length === 1 && selectedRect ? (
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
  const blockRef = useRef<HTMLElement | null>(null)
  const editorRef = useRef<HTMLElement | null>(null)
  const editingSessionRef = useRef<{ blockId: string; text: string } | null>(null)
  const pendingTrailingLineBreakRef = useRef<{ beforeText: string } | null>(null)
  const rectRef = useRef(rect)
  const committedRef = useRef(false)

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

  const syncAutoHeight = useCallback((element: HTMLElement) => {
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
  }, [minimumHeight])

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
    pendingTrailingLineBreakRef.current = null
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
      pendingTrailingLineBreakRef.current = lineBreakAtEnd ? { beforeText } : null
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
      if (!insertTextAfterPendingTrailingLineBreak(event.currentTarget, event.key)) {
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
        !insertTextAfterPendingTrailingLineBreak(
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
      pendingTrailingLineBreakRef.current = lineBreakAtEnd ? { beforeText } : null
      syncAutoHeight(event.currentTarget)
    }
  }

  function handlePaste(event: ReactClipboardEvent<HTMLElement>) {
    event.preventDefault()
    event.stopPropagation()
    const text = event.clipboardData.getData('text/plain')

    if (!insertTextAfterPendingTrailingLineBreak(event.currentTarget, text)) {
      insertTextAtSelection(event.currentTarget, text)
    }
    syncAutoHeight(event.currentTarget)
  }

  function insertTextAfterPendingTrailingLineBreak(
    element: HTMLElement,
    insertedText: string,
  ) {
    const pendingLineBreak = pendingTrailingLineBreakRef.current

    pendingTrailingLineBreakRef.current = null

    if (
      !pendingLineBreak ||
      element.textContent !== `${pendingLineBreak.beforeText}\n`
    ) {
      return false
    }

    element.textContent = `${pendingLineBreak.beforeText}\n${insertedText}`
    placeCaretAtTextOffset(
      element,
      pendingLineBreak.beforeText.length + 1 + insertedText.length,
    )

    return true
  }

  function handleInput(event: ReactFormEvent<HTMLElement>) {
    const pendingLineBreak = pendingTrailingLineBreakRef.current

    if (pendingLineBreak) {
      pendingTrailingLineBreakRef.current = null
      repairTrailingLineBreakInput(event.currentTarget, pendingLineBreak.beforeText)
    }

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

function isSelectionAtTextEnd(root: HTMLElement) {
  const selection = window.getSelection()

  if (!selection || selection.rangeCount === 0) {
    return true
  }

  const range = selection.getRangeAt(0)

  if (!root.contains(range.commonAncestorContainer)) {
    return true
  }

  const afterCaret = range.cloneRange()
  afterCaret.selectNodeContents(root)
  afterCaret.setStart(range.endContainer, range.endOffset)

  return afterCaret.toString().length === 0
}

function insertTextAtSelection(root: HTMLElement, text: string) {
  const selection = window.getSelection()
  const insertedText = text

  if (!selection || selection.rangeCount === 0) {
    root.append(document.createTextNode(insertedText))
    placeCaretAtEnd(root)
    return
  }

  const range = selection.getRangeAt(0)

  if (!root.contains(range.commonAncestorContainer)) {
    root.append(document.createTextNode(insertedText))
    placeCaretAtEnd(root)
    return
  }

  const textNode = document.createTextNode(insertedText)

  range.deleteContents()
  range.insertNode(textNode)
  range.setStart(textNode, insertedText.length)
  range.collapse(true)
  selection.removeAllRanges()
  selection.addRange(range)
}

function repairTrailingLineBreakInput(root: HTMLElement, beforeText: string) {
  const text = root.textContent ?? ''

  if (text === `${beforeText}\n`) {
    return
  }

  if (!text.startsWith(beforeText) || !text.endsWith('\n')) {
    return
  }

  const insertedText = text.slice(beforeText.length, -1)

  if (insertedText.length === 0) {
    return
  }

  root.textContent = `${beforeText}\n${insertedText}`
  placeCaretAtTextOffset(root, beforeText.length + 1 + insertedText.length)
}

function placeCaretAtTextOffset(root: HTMLElement, offset: number) {
  const selection = window.getSelection()
  const range = document.createRange()
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
  let remainingOffset = offset
  let node = walker.nextNode()

  while (node) {
    const textLength = node.textContent?.length ?? 0

    if (remainingOffset <= textLength) {
      range.setStart(node, remainingOffset)
      range.collapse(true)
      selection?.removeAllRanges()
      selection?.addRange(range)
      return
    }

    remainingOffset -= textLength
    node = walker.nextNode()
  }

  placeCaretAtEnd(root)
}

function normalizeEditableText(text: string) {
  return text.replaceAll(CARET_PLACEHOLDER, '')
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

function rememberStageScroll(element: HTMLElement) {
  const stage = element.closest<HTMLElement>('.stage-shell')

  return stage
    ? {
        element: stage,
        left: stage.scrollLeft,
        top: stage.scrollTop,
      }
    : null
}

function restoreStageScroll(
  position: { element: HTMLElement; left: number; top: number } | null,
) {
  if (!position) {
    return
  }

  position.element.scrollLeft = position.left
  position.element.scrollTop = position.top
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
  return (
    draftLayout?.rects.find((draftRect) => draftRect.pointer === pointer)?.rect ??
    getRect(block)
  )
}

function calculateMoveInteractionState(
  dx: number,
  dy: number,
  interaction: Extract<Interaction, { kind: 'move' }>,
  blocks: SlideBlock[],
  slideIndex: number,
) {
  const primaryStartRect = interaction.startRect
  const movedPrimaryRect = moveRect(primaryStartRect, dx, dy)
  const primarySnap = snapMoveRectToSlideBlocks(
    movedPrimaryRect,
    interaction.pointer,
    blocks,
    slideIndex,
    interaction.pointers,
  )
  const groupDelta = clampGroupDelta(
    interaction.startRects,
    primarySnap.rect.x - primaryStartRect.x,
    primarySnap.rect.y - primaryStartRect.y,
  )

  return {
    guides: primarySnap.guides,
    rects: interaction.startRects.map(({ pointer, rect }) => ({
      pointer,
      rect: {
        ...rect,
        x: rect.x + groupDelta.x,
        y: rect.y + groupDelta.y,
      },
    })),
  }
}

function clampGroupDelta(rects: DraftLayoutRect[], dx: number, dy: number) {
  const minX = Math.min(...rects.map(({ rect }) => rect.x))
  const minY = Math.min(...rects.map(({ rect }) => rect.y))
  const maxX = Math.max(...rects.map(({ rect }) => rect.x + rect.width))
  const maxY = Math.max(...rects.map(({ rect }) => rect.y + rect.height))

  return {
    x: clamp(dx, -minX, SLIDE_WIDTH - maxX),
    y: clamp(dy, -minY, SLIDE_HEIGHT - maxY),
  }
}

function draftRectsEqual(a: DraftLayoutRect[], b: DraftLayoutRect[]) {
  return (
    a.length === b.length &&
    a.every((draftRect) => {
      const other = b.find((candidate) => candidate.pointer === draftRect.pointer)

      return other ? rectEquals(draftRect.rect, other.rect) : false
    })
  )
}

function selectionActionForPointers(
  pointers: Pointer[],
  primaryPointer = pointers.at(-1),
): SelectionAction {
  if (pointers.length <= 1) {
    return { type: 'collapse', pointer: pointers[0] ?? primaryPointer ?? '' }
  }

  return {
    type: 'selectRanges',
    ranges: pointers,
    primaryIndex: Math.max(0, pointers.indexOf(primaryPointer ?? pointers.at(-1)!)),
  }
}

function hasSelectionModifier(
  event:
    | ReactMouseEvent<HTMLElement>
    | ReactPointerEvent<HTMLElement>
    | ReactPointerEvent<HTMLButtonElement>,
) {
  return event.shiftKey || event.metaKey || event.ctrlKey
}

function snapMoveRectToSlideBlocks(
  rect: Rect,
  pointer: Pointer,
  blocks: SlideBlock[],
  slideIndex: number,
  excludedPointers: Pointer[] = [pointer],
) {
  const excludedPointerSet = new Set(excludedPointers)
  const peerRects = blocks
    .filter(
      (_, blockIndex) =>
        !excludedPointerSet.has(blockPointer(slideIndex, blockIndex)),
    )
    .map(getRect)
  const xSnap = closestAxisSnap(
    [rect.x, rect.x + rect.width / 2, rect.x + rect.width],
    peerRects.flatMap((peerRect) => [
      peerRect.x,
      peerRect.x + peerRect.width / 2,
      peerRect.x + peerRect.width,
    ]),
  )
  const ySnap = closestAxisSnap(
    [rect.y, rect.y + rect.height / 2, rect.y + rect.height],
    peerRects.flatMap((peerRect) => [
      peerRect.y,
      peerRect.y + peerRect.height / 2,
      peerRect.y + peerRect.height,
    ]),
  )
  const snappedRect = {
    ...rect,
    x: clamp(rect.x + (xSnap?.offset ?? 0), 0, SLIDE_WIDTH - rect.width),
    y: clamp(rect.y + (ySnap?.offset ?? 0), 0, SLIDE_HEIGHT - rect.height),
  }
  const slideGuideX = alignedGuideFor(
    snappedRect.x,
    snappedRect.x + snappedRect.width / 2,
    snappedRect.x + snappedRect.width,
    SLIDE_WIDTH,
  )
  const slideGuideY = alignedGuideFor(
    snappedRect.y,
    snappedRect.y + snappedRect.height / 2,
    snappedRect.y + snappedRect.height,
    SLIDE_HEIGHT,
  )

  return {
    guides: {
      x: slideGuideX ?? xSnap?.guide ?? null,
      y: slideGuideY ?? ySnap?.guide ?? null,
    },
    rect: snappedRect,
  }
}

function closestAxisSnap(sources: number[], targets: number[]) {
  let closest: { distance: number; guide: number; offset: number } | null = null

  for (const source of sources) {
    for (const target of targets) {
      const distance = Math.abs(source - target)

      if (distance > GRID_SIZE) {
        continue
      }

      if (!closest || distance < closest.distance) {
        closest = {
          distance,
          guide: target,
          offset: target - source,
        }
      }
    }
  }

  return closest
}

function guidesForInteraction(rect: Rect, interaction: Interaction): SnapGuides {
  const left = rect.x
  const centerX = rect.x + rect.width / 2
  const right = rect.x + rect.width
  const top = rect.y
  const centerY = rect.y + rect.height / 2
  const bottom = rect.y + rect.height

  if (interaction.kind === 'move') {
    return {
      x: alignedGuideFor(left, centerX, right, SLIDE_WIDTH),
      y: alignedGuideFor(top, centerY, bottom, SLIDE_HEIGHT),
    }
  }

  const activeEdgeX = interaction.handle.includes('e')
    ? right
    : interaction.handle.includes('w')
      ? left
      : null
  const activeEdgeY = interaction.handle.includes('s')
    ? bottom
    : interaction.handle.includes('n')
      ? top
      : null

  return {
    x:
      activeEdgeX === null
        ? null
        : alignedGuideFor(activeEdgeX, centerX, activeEdgeX, SLIDE_WIDTH),
    y:
      activeEdgeY === null
        ? null
        : alignedGuideFor(activeEdgeY, centerY, activeEdgeY, SLIDE_HEIGHT),
  }
}

function alignedGuideFor(start: number, center: number, end: number, containerSize: number) {
  const target = containerSize / 2

  if (Math.abs(center - target) < 0.5) {
    return target
  }

  if (Math.abs(start) < 0.5) {
    return 0
  }

  if (Math.abs(end - containerSize) < 0.5) {
    return containerSize
  }

  return null
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

function resizeHandleAffectsHeight(handle: ResizeHandle) {
  return handle.includes('n') || handle.includes('s')
}

function minimumHeightForBlock(
  block: SlideBlock,
  baseBlock: SlideBlock | undefined,
  rect: Rect,
  draftHeightResize = false,
) {
  if (block.text.length === 0) {
    return EMPTY_TEXT_BOX_HEIGHT
  }

  if (
    draftHeightResize ||
    (baseBlock && block.text === baseBlock.text && block.height !== baseBlock.height)
  ) {
    return rect.height
  }

  return 0
}

function arrangeResetEquals(a: SlideBlock, b: SlideBlock) {
  const heightChangedByArrange = a.text === b.text && a.height !== b.height

  return (
    a.x === b.x &&
    a.y === b.y &&
    a.width === b.width &&
    !heightChangedByArrange
  )
}

function textResetEquals(a: SlideBlock, b: SlideBlock) {
  return a.text === b.text && a.height === b.height
}

async function writeExportToClipboard(
  exportCode: string,
  fallbackTextarea: HTMLTextAreaElement | null,
) {
  try {
    await navigator.clipboard.writeText(exportCode)
    return true
  } catch {
    if (!fallbackTextarea) {
      return false
    }

    fallbackTextarea.value = exportCode
    fallbackTextarea.select()

    return document.execCommand('copy')
  }
}

function deckEquals(a: unknown, b: unknown) {
  const parsedA = RetouchDeckSchema.safeParse(a)
  const parsedB = RetouchDeckSchema.safeParse(b)

  if (!parsedA.success || !parsedB.success) {
    return false
  }

  return JSON.stringify(parsedA.data) === JSON.stringify(parsedB.data)
}

function changedSlides(deck: unknown) {
  const parsed = RetouchDeckSchema.safeParse(deck)
  const parsedBase = RetouchDeckSchema.safeParse(SAMPLE_DECK)

  if (!parsed.success || !parsedBase.success) {
    return new Set<string>()
  }

  return new Set(
    parsed.data.slides
      .filter((slide) => {
        const baseSlide = parsedBase.data.slides.find(
          (candidate) => candidate.id === slide.id,
        )

        return !baseSlide || JSON.stringify(slide) !== JSON.stringify(baseSlide)
      })
      .map((slide) => slide.id),
  )
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

function isControlTarget(target: EventTarget | null) {
  if (!(target instanceof HTMLElement)) {
    return false
  }

  return Boolean(target.closest('button, [role="tab"], [role="toolbar"]'))
}

function arrowKeyDelta(key: string, largeStep: boolean): Point | null {
  const step = largeStep ? GRID_SIZE * 5 : GRID_SIZE

  if (key === 'ArrowLeft') {
    return { x: -step, y: 0 }
  }

  if (key === 'ArrowRight') {
    return { x: step, y: 0 }
  }

  if (key === 'ArrowUp') {
    return { x: 0, y: -step }
  }

  if (key === 'ArrowDown') {
    return { x: 0, y: step }
  }

  return null
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
    if (deckEquals(deck, SAMPLE_DECK)) {
      window.localStorage.removeItem(STORAGE_KEY)
      return
    }

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
