import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
} from 'react'
import type { JSONPatchOperation, Pointer, SelectionAction } from 'zod-crud'
import { useJSONDocument } from 'zod-crud/react'
import {
  MIN_BLOCK_SIZE,
  SAMPLE_DECK,
  SAMPLE_SLIDES,
  SLIDE_HEIGHT,
  SLIDE_WIDTH,
  RetouchDeckSchema,
  blockLocationFromPointer,
  blockPointer,
  clamp,
  findSlideIndex,
  getRect,
  rectEquals,
  slideBlocksPointer,
  slideAccentPointer,
  slideNamePointer,
  slidePointer,
  setLayoutPatch,
  snap,
  type Rect,
} from './retouchModel'
import { exportRetouchDeck } from './retouchExport'
import { PresentationOverlay } from './PresentationOverlay'
import { RetouchWorkspace } from './RetouchWorkspace'
import { SlideRail } from './SlideRail'
import { createTextBlock, duplicateBlock } from './slideBlockOperations'
import { createBlankSlide, duplicateSlide } from './slideDeckOperations'
import {
  alignRectToBounds,
  alignmentBounds,
  distributeRects,
  type AlignSelectionAction,
  type DistributeSelectionAction,
} from './selectionAlignment'
import {
  blockOrderChanged,
  reorderBlocksByLayer,
  type LayerOrderAction,
} from './selectionLayerOrder'
import { useExportControls } from './useExportControls'
import { useRetouchLayoutInteraction } from './useRetouchLayoutInteraction'
import { useRetouchMarqueeSelection } from './useRetouchMarqueeSelection'
import { useRetouchKeyboardShortcuts } from './useRetouchKeyboardShortcuts'
import { useRetouchResetActions } from './useRetouchResetActions'
import { useRetouchSelectionState } from './useRetouchSelectionState'
import { useRetouchTextEditing } from './useRetouchTextEditing'
import { useVisualSelectionRect } from './useVisualSelectionRect'
import {
  changedSlides,
  deckEquals,
  persistDeck,
  readInitialDeck,
} from './retouchPersistence'
import {
  getCurrentRect,
  hasSelectionModifier,
  selectionActionForPointers,
  type Point,
} from './layoutInteraction'
import './App.css'

type Mode = 'text' | 'layout'
type CanvasView = 'slide' | 'grid'
type RectField = keyof Rect
type CanvasZoom = 'fit' | number

const CANVAS_ZOOM_STEPS = [0.5, 0.75, 1, 1.25, 1.5, 2]

type EditingState = {
  clientPoint?: Point
  pointer: Pointer
}

function finiteNumber(value: number, fallback: number) {
  return Number.isFinite(value) ? value : fallback
}

function normalizeInspectorRect(
  rect: Rect,
  currentRect: Rect,
  changedField?: RectField,
): Rect {
  const nextRect = { ...currentRect }

  if (changedField === 'x') {
    nextRect.x = clamp(
      snap(finiteNumber(rect.x, currentRect.x)),
      0,
      SLIDE_WIDTH - currentRect.width,
    )
  }

  if (changedField === 'y') {
    nextRect.y = clamp(
      snap(finiteNumber(rect.y, currentRect.y)),
      0,
      SLIDE_HEIGHT - currentRect.height,
    )
  }

  if (changedField === 'width') {
    nextRect.width = clamp(
      snap(finiteNumber(rect.width, currentRect.width)),
      MIN_BLOCK_SIZE,
      SLIDE_WIDTH - currentRect.x,
    )
  }

  if (changedField === 'height') {
    nextRect.height = clamp(
      snap(finiteNumber(rect.height, currentRect.height)),
      MIN_BLOCK_SIZE,
      SLIDE_HEIGHT - currentRect.y,
    )
  }

  return nextRect
}

function App() {
  const initialDeck = useMemo(() => readInitialDeck(), [])
  const doc = useJSONDocument(RetouchDeckSchema, initialDeck, {
    history: 200,
    selection: { mode: 'extended' },
  })
  const slideRef = useRef<HTMLDivElement | null>(null)
  const stageRef = useRef<HTMLDivElement | null>(null)
  const suppressBlockClickRef = useRef(false)
  const suppressStageClickRef = useRef(false)

  const [mode, setMode] = useState<Mode>('text')
  const [canvasView, setCanvasView] = useState<CanvasView>('slide')
  const [canvasZoom, setCanvasZoom] = useState<CanvasZoom>('fit')
  const [activeSlideId, setActiveSlideId] = useState(SAMPLE_SLIDES[0].id)
  const [editing, setEditing] = useState<EditingState | null>(null)
  const [notesBySlideId, setNotesBySlideId] = useState<Record<string, string>>({})
  const [presenting, setPresenting] = useState(false)

  const activeSlideIndex = Math.max(0, findSlideIndex(doc.value, activeSlideId))
  const activeSlide = doc.value.slides[activeSlideIndex] ?? doc.value.slides[0]
  const hasDeckChanges = !deckEquals(doc.value, SAMPLE_DECK)
  const changedSlideIds = useMemo(() => changedSlides(doc.value), [doc.value])
  const {
    baseSelectedLocation,
    canReset,
    canResetSelectedLayout,
    resetScope,
    resetTitle,
    selectedBlock,
    selectedPointer,
    selectedPointerSet,
    selectedPointers,
  } = useRetouchSelectionState({
    activeSlide,
    deckValue: doc.value,
    editing,
    focusPointer: doc.selection?.focusPointer ?? null,
    hasDeckChanges,
    mode,
    selectedPointersFromDocument: doc.selection?.selectedPointers ?? [],
  })

  useEffect(() => {
    persistDeck(doc.value)
  }, [doc.value])

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

      doc.commit(patch, {
        label,
        mergeKey,
        origin: 'ppt-retouch',
        selection: selection ?? { type: 'collapse', pointer },
      })
    },
    [doc],
  )

  const {
    clearLayoutInteraction,
    draftLayout,
    interaction,
    snapGuides,
    startMoveInteraction,
    startResizeInteraction,
  } = useRetouchLayoutInteraction({
    activeSlideBlocks: activeSlide.blocks,
    activeSlideId: activeSlide.id,
    activeSlideIndex,
    commitPatch,
    deckValue: doc.value,
    mode,
    selectedBlock,
    selectedPointer,
    selectedPointerSet,
    selectedPointers,
    selectBlock,
    slideRef,
    suppressBlockClickRef,
    suppressStageClickRef,
  })
  const {
    cancelTextEdit,
    commitActiveTextEdit,
    commitTextEdit,
    commitTextPatch,
    startTextEdit,
  } = useRetouchTextEditing({
    clearLayoutInteraction,
    commitPatch,
    deckValue: doc.value,
    editing,
    mode,
    selectBlock,
    setEditing,
    slideRef,
  })
  const {
    clearMarqueeSelection,
    marqueeRect,
    startMarqueeSelection,
  } = useRetouchMarqueeSelection({
    activeSlide,
    activeSlideIndex,
    mode,
    selectedPointers,
    selection: doc.selection,
    slideRef,
    suppressStageClickRef,
  })
  const selectedRect =
    selectedPointer && selectedBlock
      ? getCurrentRect(selectedPointer, selectedBlock, draftLayout)
      : null
  const exportCode = useMemo(() => exportRetouchDeck(doc.value), [doc.value])
  const exportStatusMatchesVisibleSlide = !editing && !interaction && !draftLayout
  const {
    copyExportCode,
    copyState,
    copyTitle,
    downloadExportCode,
    exportCopied,
    exportDownloaded,
    exportTextareaRef,
    resetExportFeedback,
  } = useExportControls({
    exportCode,
    readCommittedExportCode: () => exportRetouchDeck(commitActiveTextEdit()),
    statusMatchesVisibleSlide: exportStatusMatchesVisibleSlide,
  })
  const {
    redoDocumentChange,
    resetCurrentTarget,
    undoDocumentChange,
  } = useRetouchResetActions({
    baseSelectedLocation,
    canResetSelectedLayout,
    clearTransientState,
    commitActiveTextEdit,
    commitDeckReset: () => {
      doc.commit([{ op: 'replace', path: '', value: SAMPLE_DECK }], {
        label: 'reset deck',
        origin: 'ppt-retouch',
      })
    },
    commitPatch,
    commitTextPatch,
    deckValue: doc.value,
    editing,
    hasDeckChanges,
    history: doc.history,
    mode,
    resetExportFeedback,
    selectedBlock,
    selectedPointer,
    selection: doc.selection,
    setActiveSlideId,
    setEditing,
    slideRef,
    stageRef,
  })
  const visualSelectionRect = useVisualSelectionRect({
    activeSlideId: activeSlide.id,
    deckValue: doc.value,
    draftLayout,
    mode,
    selectedBlock,
    selectedPointerCount: selectedPointers.length,
    slideRef,
  })

  useRetouchKeyboardShortcuts({
    activeSlideId: activeSlide.id,
    commitPatch,
    deckValue: doc.value,
    editing,
    history: doc.history,
    interaction,
    mode,
    selectedPointer,
    selectedPointers,
    selection: doc.selection,
    setEditing,
  })

  function clearTransientState() {
    setEditing(null)
    clearLayoutInteraction()
    clearMarqueeSelection()
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

  function changeCanvasView(nextView: CanvasView) {
    if (nextView === 'grid') {
      commitActiveTextEdit()
      doc.selection?.empty()
      clearTransientState()
    }

    setCanvasView(nextView)
  }

  function zoomStep(direction: -1 | 1) {
    const currentZoom = canvasZoom === 'fit' ? 1 : canvasZoom
    const currentIndex = CANVAS_ZOOM_STEPS.findIndex((step) => step >= currentZoom)
    const fallbackIndex = CANVAS_ZOOM_STEPS.indexOf(1)
    const nextIndex =
      direction > 0
        ? currentZoom >= CANVAS_ZOOM_STEPS.at(-1)!
          ? CANVAS_ZOOM_STEPS.length - 1
          : Math.max(0, currentIndex) + 1
        : currentZoom <= CANVAS_ZOOM_STEPS[0]
          ? 0
          : currentIndex > 0
            ? currentIndex - 1
            : fallbackIndex

    setCanvasZoom(CANVAS_ZOOM_STEPS[nextIndex] ?? 1)
  }

  function fitCanvasZoom() {
    setCanvasZoom('fit')
  }

  function activateSlide(slideId: string) {
    setActiveSlideId(slideId)
    setCanvasView('slide')
    doc.selection?.empty()
    stageRef.current?.scrollTo({ left: 0, top: 0 })
    clearTransientState()
  }

  function selectSlide(slideId: string) {
    commitActiveTextEdit()
    activateSlide(slideId)
  }

  function addSlide() {
    commitActiveTextEdit()
    const nextSlide = createBlankSlide(doc.value.slides)
    const insertIndex = activeSlideIndex + 1

    doc.commit([{ op: 'add', path: slidePointer(insertIndex), value: nextSlide }], {
      label: 'add slide',
      origin: 'ppt-retouch',
    })
    activateSlide(nextSlide.id)
  }

  function copySlide() {
    commitActiveTextEdit()
    const nextSlide = duplicateSlide(activeSlide, doc.value.slides)
    const insertIndex = activeSlideIndex + 1

    doc.commit([{ op: 'add', path: slidePointer(insertIndex), value: nextSlide }], {
      label: 'duplicate slide',
      origin: 'ppt-retouch',
    })
    setNotesBySlideId((current) => ({
      ...current,
      [nextSlide.id]: current[activeSlide.id] ?? '',
    }))
    activateSlide(nextSlide.id)
  }

  function deleteSlide() {
    if (doc.value.slides.length <= 1) {
      return
    }

    commitActiveTextEdit()
    const nextSlide =
      doc.value.slides[activeSlideIndex + 1] ??
      doc.value.slides[activeSlideIndex - 1] ??
      doc.value.slides[0]

    doc.commit([{ op: 'remove', path: slidePointer(activeSlideIndex) }], {
      label: 'delete slide',
      origin: 'ppt-retouch',
    })
    setNotesBySlideId((current) => {
      const next = { ...current }
      delete next[activeSlide.id]
      return next
    })
    activateSlide(nextSlide.id)
  }

  function moveSlide(direction: -1 | 1) {
    const nextIndex = activeSlideIndex + direction

    if (nextIndex < 0 || nextIndex >= doc.value.slides.length) {
      return
    }

    commitActiveTextEdit()
    doc.commit(
      [
        { op: 'remove', path: slidePointer(activeSlideIndex) },
        { op: 'add', path: slidePointer(nextIndex), value: activeSlide },
      ],
      {
        label: direction < 0 ? 'move slide up' : 'move slide down',
        origin: 'ppt-retouch',
      },
    )
    setActiveSlideId(activeSlide.id)
    setCanvasView('slide')
    doc.selection?.empty()
    clearTransientState()
  }

  function changeSlideName(name: string) {
    const nextName = name.trim() || 'Untitled'

    if (nextName === activeSlide.name) {
      return
    }

    commitActiveTextEdit()
    doc.commit(
      [
        {
          op: 'replace',
          path: slideNamePointer(activeSlideIndex),
          value: nextName,
        },
      ],
      {
        label: 'rename slide',
        origin: 'ppt-retouch',
      },
    )
    clearTransientState()
  }

  function changeSlideAccent(accent: string) {
    if (accent === activeSlide.accent) {
      return
    }

    commitActiveTextEdit()
    doc.commit(
      [
        {
          op: 'replace',
          path: slideAccentPointer(activeSlideIndex),
          value: accent,
        },
      ],
      {
        label: 'change slide accent',
        origin: 'ppt-retouch',
      },
    )
    clearTransientState()
  }

  function insertTextBlock() {
    commitActiveTextEdit()
    const nextBlock = createTextBlock(activeSlide)
    const blockIndex = activeSlide.blocks.length
    const pointer = blockPointer(activeSlideIndex, blockIndex)

    doc.commit([{ op: 'add', path: pointer, value: nextBlock }], {
      label: 'add text block',
      origin: 'ppt-retouch',
      selection: { type: 'collapse', pointer },
    })
    setCanvasView('slide')
    setMode('text')
    clearLayoutInteraction()
    setEditing({ pointer })
    stageRef.current?.scrollTo({ left: 0, top: 0 })
  }

  function selectedActiveBlockLocations() {
    return selectedPointers
      .map((pointer) => blockLocationFromPointer(doc.value, pointer))
      .filter(
        (location): location is NonNullable<typeof location> =>
          location !== null && location.slide.id === activeSlide.id,
      )
      .sort((a, b) => a.blockIndex - b.blockIndex)
  }

  function duplicateSelectedBlock() {
    const locations = selectedActiveBlockLocations()

    if (locations.length === 0) {
      return
    }

    commitActiveTextEdit()
    const duplicatedBlocks: typeof activeSlide.blocks = []

    for (const location of locations) {
      duplicatedBlocks.push(
        duplicateBlock(location.block, {
          ...activeSlide,
          blocks: [...activeSlide.blocks, ...duplicatedBlocks],
        }),
      )
    }

    const insertIndex = locations.at(-1)!.blockIndex + 1
    const duplicatePointers = duplicatedBlocks.map((_, offset) =>
      blockPointer(activeSlideIndex, insertIndex + offset),
    )

    doc.commit(duplicatedBlocks.map((block, offset) => ({
      op: 'add',
      path: blockPointer(activeSlideIndex, insertIndex + offset),
      value: block,
    })), {
      label: locations.length > 1 ? 'duplicate blocks' : 'duplicate block',
      origin: 'ppt-retouch',
      selection: selectionActionForPointers(duplicatePointers),
    })
    setCanvasView('slide')
    setMode('layout')
    setEditing(null)
    clearLayoutInteraction()
  }

  function deleteSelectedBlock() {
    const locations = selectedActiveBlockLocations()

    if (locations.length === 0) {
      return
    }

    commitActiveTextEdit()
    const nextSelectionIndex = Math.min(
      locations[0].blockIndex,
      activeSlide.blocks.length - locations.length - 1,
    )

    doc.commit(
      [...locations]
        .sort((a, b) => b.blockIndex - a.blockIndex)
        .map((location) => ({
          op: 'remove',
          path: blockPointer(activeSlideIndex, location.blockIndex),
        })),
      {
        label: locations.length > 1 ? 'delete blocks' : 'delete block',
        origin: 'ppt-retouch',
      },
    )
    setCanvasView('slide')
    setMode('layout')
    setEditing(null)
    clearLayoutInteraction()

    if (nextSelectionIndex >= 0) {
      doc.selection?.selectRanges?.([blockPointer(activeSlideIndex, nextSelectionIndex)])
    } else {
      doc.selection?.empty()
    }
  }

  function alignSelectedBlocks(action: AlignSelectionAction) {
    const locations = selectedActiveBlockLocations()

    if (locations.length === 0) {
      return
    }

    const bounds = alignmentBounds(locations.map((location) => getRect(location.block)))

    if (!bounds) {
      return
    }

    const targets = locations.map((location) => ({
      pointer: location.pointer,
      rect: alignRectToBounds(getRect(location.block), bounds, action),
      startRect: getRect(location.block),
    }))

    if (targets.every((target) => rectEquals(target.rect, target.startRect))) {
      return
    }

    commitActiveTextEdit()
    setCanvasView('slide')
    setMode('layout')
    clearTransientState()
    commitPatch(
      targets.flatMap((target) => setLayoutPatch(target.pointer, target.rect)),
      targets.at(-1)?.pointer ?? selectedPointer ?? targets[0].pointer,
      'align selection',
      undefined,
      selectionActionForPointers(
        targets.map((target) => target.pointer),
        selectedPointer ?? targets.at(-1)?.pointer,
      ),
    )
  }

  function distributeSelectedBlocks(action: DistributeSelectionAction) {
    const locations = selectedActiveBlockLocations()

    if (locations.length < 3) {
      return
    }

    const targets = distributeRects(
      locations.map((location) => ({
        item: location,
        rect: getRect(location.block),
      })),
      action,
    ).map(({ item: location, rect }) => ({
      pointer: location.pointer,
      rect,
      startRect: getRect(location.block),
    }))

    if (targets.every((target) => rectEquals(target.rect, target.startRect))) {
      return
    }

    commitActiveTextEdit()
    setCanvasView('slide')
    setMode('layout')
    clearTransientState()
    commitPatch(
      targets.flatMap((target) => setLayoutPatch(target.pointer, target.rect)),
      targets.at(-1)?.pointer ?? selectedPointer ?? targets[0].pointer,
      'distribute selection',
      undefined,
      selectionActionForPointers(
        targets.map((target) => target.pointer),
        selectedPointer ?? targets.at(-1)?.pointer,
      ),
    )
  }

  function changeSelectedLayerOrder(action: LayerOrderAction) {
    const locations = selectedActiveBlockLocations()

    if (locations.length === 0) {
      return
    }

    const selectedIds = locations.map((location) => location.block.id)
    const nextBlocks = reorderBlocksByLayer(activeSlide.blocks, selectedIds, action)

    if (!blockOrderChanged(activeSlide.blocks, nextBlocks)) {
      return
    }

    const selectedIdSet = new Set(selectedIds)
    const nextSelectedPointers = nextBlocks
      .map((block, blockIndex) =>
        selectedIdSet.has(block.id) ? blockPointer(activeSlideIndex, blockIndex) : null,
      )
      .filter((pointer): pointer is Pointer => pointer !== null)

    commitActiveTextEdit()
    doc.commit(
      [
        {
          op: 'replace',
          path: slideBlocksPointer(activeSlideIndex),
          value: nextBlocks,
        },
      ],
      {
        label: 'reorder layers',
        origin: 'ppt-retouch',
        selection: selectionActionForPointers(
          nextSelectedPointers,
          nextSelectedPointers.at(-1),
        ),
      },
    )
    setCanvasView('slide')
    setMode('layout')
    clearTransientState()
  }

  function changeSelectedBlockRect(rect: Rect, changedField?: RectField) {
    if (!selectedPointer || !selectedBlock || !selectedRect) {
      return
    }

    const nextRect = normalizeInspectorRect(rect, selectedRect, changedField)

    if (rectEquals(nextRect, selectedRect)) {
      return
    }

    commitActiveTextEdit()
    setCanvasView('slide')
    setMode('layout')
    clearTransientState()
    commitPatch(
      setLayoutPatch(selectedPointer, nextRect),
      selectedPointer,
      'edit block geometry',
      `layout:geometry:${selectedPointer}`,
    )
  }

  function startPresentation() {
    commitActiveTextEdit()
    setCanvasView('slide')
    doc.selection?.empty()
    clearTransientState()
    setPresenting(true)
  }

  function closePresentation() {
    setPresenting(false)
  }

  function navigatePresentation(direction: -1 | 1) {
    const nextIndex = Math.min(
      doc.value.slides.length - 1,
      Math.max(0, activeSlideIndex + direction),
    )
    const nextSlide = doc.value.slides[nextIndex]

    if (!nextSlide || nextSlide.id === activeSlide.id) {
      return
    }

    setActiveSlideId(nextSlide.id)
    doc.selection?.empty()
    clearTransientState()
  }

  function selectBlock(pointer: Pointer, additive = false) {
    if (additive) {
      doc.selection?.togglePointer(pointer)
      return
    }

    doc.selection?.selectRanges([pointer])
  }

  return (
    <main className="retouch-app" data-mode={mode}>
      <SlideRail
        activeSlideId={activeSlide.id}
        canMoveSlideDown={activeSlideIndex < doc.value.slides.length - 1}
        canMoveSlideUp={activeSlideIndex > 0}
        canvasView={canvasView}
        changedSlideIds={changedSlideIds}
        canDeleteSlide={doc.value.slides.length > 1}
        onChangeCanvasView={changeCanvasView}
        onAddSlide={addSlide}
        onDeleteSlide={deleteSlide}
        onDuplicateSlide={copySlide}
        onMoveSlideDown={() => moveSlide(1)}
        onMoveSlideUp={() => moveSlide(-1)}
        onSelectSlide={selectSlide}
        slides={doc.value.slides}
      />

      <RetouchWorkspace
        activeSlide={activeSlide}
        activeSlideIndex={activeSlideIndex}
        activeSlideAccent={activeSlide.accent}
        canRedo={doc.history.canRedo}
        canReset={canReset}
        canUndo={doc.history.canUndo}
        canZoomIn={canvasZoom === 'fit' || canvasZoom < CANVAS_ZOOM_STEPS.at(-1)!}
        canZoomOut={canvasZoom === 'fit' || canvasZoom > CANVAS_ZOOM_STEPS[0]}
        canvasView={canvasView}
        canvasZoom={canvasZoom}
        changedSlideIds={changedSlideIds}
        copyState={copyState}
        copyTitle={copyTitle}
        draftLayout={draftLayout}
        editing={editing}
        exportCode={exportCode}
        exportCopied={exportCopied}
        exportDownloaded={exportDownloaded}
        exportTextareaRef={exportTextareaRef}
        interaction={interaction}
        mode={mode}
        marqueeRect={marqueeRect}
        notes={notesBySlideId[activeSlide.id] ?? ''}
        onBlockClick={(event, pointer) => {
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
        onBlockPointerDown={startMoveInteraction}
        onCancelTextEdit={cancelTextEdit}
        onChangeMode={changeMode}
        onCommitTextEdit={commitTextEdit}
        onCopyExport={copyExportCode}
        onCanvasPointerDown={startMarqueeSelection}
        onDeleteBlock={deleteSelectedBlock}
        onDownloadExport={downloadExportCode}
        onDuplicateBlock={duplicateSelectedBlock}
        onAlignSelection={alignSelectedBlocks}
        onDistributeSelection={distributeSelectedBlocks}
        onLayerOrderChange={changeSelectedLayerOrder}
        onInsertTextBlock={insertTextBlock}
        onZoomFit={fitCanvasZoom}
        onZoomIn={() => zoomStep(1)}
        onZoomOut={() => zoomStep(-1)}
        onSlideAccentChange={changeSlideAccent}
        onSlideNameChange={changeSlideName}
        onNotesChange={(notes) =>
          setNotesBySlideId((current) => ({
            ...current,
            [activeSlide.id]: notes,
          }))
        }
        onOpenSlide={selectSlide}
        onPresent={startPresentation}
        onRedo={redoDocumentChange}
        onReset={resetCurrentTarget}
        onResizePointerDown={startResizeInteraction}
        onSelectedRectChange={changeSelectedBlockRect}
        onStageBackgroundClick={() => {
          if (mode === 'text') {
            commitActiveTextEdit()
          }
          doc.selection?.empty()
        }}
        onUndo={undoDocumentChange}
        resetScope={resetScope}
        resetTitle={resetTitle}
        selectedBlock={selectedBlock}
        selectedPointerSet={selectedPointerSet}
        selectedPointers={selectedPointers}
        selectedRect={selectedRect}
        slides={doc.value.slides}
        slideRef={slideRef}
        snapGuides={snapGuides}
        stageRef={stageRef}
        suppressStageClickRef={suppressStageClickRef}
        visualSelectionRect={visualSelectionRect}
        zoomLabel={
          canvasZoom === 'fit' ? 'Fit' : `${Math.round(canvasZoom * 100)}%`
        }
      />

      {presenting ? (
        <PresentationOverlay
          activeSlide={activeSlide}
          activeSlideIndex={activeSlideIndex}
          notes={notesBySlideId[activeSlide.id] ?? ''}
          onClose={closePresentation}
          onNext={() => navigatePresentation(1)}
          onPrevious={() => navigatePresentation(-1)}
          slideCount={doc.value.slides.length}
        />
      ) : null}
    </main>
  )
}

export default App
