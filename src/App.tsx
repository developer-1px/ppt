import { useCallback, useMemo, useRef, useState } from 'react'
import type { JSONPatchOperation, Pointer, SelectionSnap } from 'zod-crud'
import { useJSONDocument } from 'zod-crud/react'
import {
  SAMPLE_DECK,
  SAMPLE_SLIDES,
  RetouchDeckSchema,
  blockLocationFromPointer,
  blockPointer,
  getRect,
  rectEquals,
  setLayoutPatch,
  type Rect,
  type SlideBlock,
} from './retouchModel'
import { exportRetouchDeck } from './retouchExport'
import { PresentationOverlay } from './PresentationOverlay'
import { RetouchWorkspace } from './RetouchWorkspace'
import { SlideRail } from './SlideRail'
import {
  addBlocksPatch,
  blockInsertPointers,
  createTextBlock,
  duplicateBlocks,
} from './slideBlockOperations'
import {
  addSlidePatch,
  createBlankSlide,
  duplicateSlide,
  setSlideAccentPatch,
  setSlideNamePatch,
} from './slideDeckOperations'
import {
  alignRectToBounds,
  alignmentBounds,
  distributeRects,
  type AlignSelectionAction,
  type DistributeSelectionAction,
} from './selectionAlignment'
import {
  createLayerOrderPatch,
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
import { useRetouchDraftPersistence } from './retouchPersistence'
import { changedSlideIdsFromBaseline } from './retouchDirtyState'
import { createRetouchCollection } from './retouchCollection'
import { createRetouchIdResolver } from './retouchIdResolver'
import {
  getCurrentRect,
  hasSelectionModifier,
  selectionSnapForPointers,
} from './layoutInteraction'
import { useCanvasViewTabs } from './useCanvasViewTabs'
import {
  normalizeInspectorRect,
  type RectField,
} from './inspectorGeometry'
import {
  canDecreaseCanvasZoom,
  canIncreaseCanvasZoom,
  canvasZoomLabel,
  nextCanvasZoom,
  type CanvasZoom,
} from './canvasZoom'
import type {
  CanvasView,
  EditingState,
  RetouchMode,
} from './retouchViewState'
import './App.css'

function App() {
  const doc = useJSONDocument(RetouchDeckSchema, SAMPLE_DECK, {
    history: 200,
    selection: { mode: 'extended' },
  })
  const slideRef = useRef<HTMLDivElement | null>(null)
  const stageRef = useRef<HTMLDivElement | null>(null)
  const suppressBlockClickRef = useRef(false)
  const suppressStageClickRef = useRef(false)

  const [mode, setMode] = useState<RetouchMode>('text')
  const [canvasView, setCanvasView] = useState<CanvasView>('slide')
  const [canvasZoom, setCanvasZoom] = useState<CanvasZoom>('fit')
  const [activeSlideId, setActiveSlideId] = useState(SAMPLE_SLIDES[0].id)
  const [editing, setEditing] = useState<EditingState | null>(null)
  const [blockClipboard, setBlockClipboard] = useState<SlideBlock[]>([])
  const [notesBySlideId, setNotesBySlideId] = useState<Record<string, string>>({})
  const [presenting, setPresenting] = useState(false)

  const retouchIds = useMemo(() => createRetouchIdResolver(doc), [doc])
  const retouchCollection = useMemo(() => createRetouchCollection(doc), [doc])
  const activeSlideIndex = Math.max(
    0,
    retouchIds.resolveSlideIndex(activeSlideId) ?? 0,
  )
  const activeSlide = doc.value.slides[activeSlideIndex] ?? doc.value.slides[0]
  const { hasDeckChanges } = useRetouchDraftPersistence(doc)
  const changedSlideIds = useMemo(
    () => changedSlideIdsFromBaseline(doc.value),
    [doc.value],
  )
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

  const commitPatch = useCallback(
    (
      patch: JSONPatchOperation[],
      pointer: Pointer,
      label: string,
      mergeKey?: string,
      selection?: SelectionSnap,
    ) => {
      if (patch.length === 0) {
        return
      }

      doc.commit(patch, {
        label,
        mergeKey,
        origin: 'ppt-retouch',
        selection: selection ?? selectionSnapForPointers([pointer]),
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
    canPasteSelection: blockClipboard.length > 0,
    onCopySelection: copySelectedBlocks,
    onDeleteSelection: deleteSelectedBlock,
    onDuplicateSelection: duplicateSelectedBlock,
    onPasteSelection: pasteCopiedBlocks,
    onSelectAllBlocks: selectAllBlocks,
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

  function changeMode(nextMode: RetouchMode) {
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

  const {
    canvasViewPanelProps,
    canvasViewTabProps,
    canvasViewTablistProps,
  } = useCanvasViewTabs({
    canvasView,
    onChange: changeCanvasView,
  })

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

    doc.commit(addSlidePatch(nextSlide, insertIndex), {
      label: 'add slide',
      origin: 'ppt-retouch',
    })
    activateSlide(nextSlide.id)
  }

  function copySlide() {
    commitActiveTextEdit()
    const nextSlide = duplicateSlide(activeSlide, doc.value.slides)
    const insertIndex = activeSlideIndex + 1

    doc.commit(addSlidePatch(nextSlide, insertIndex), {
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

    const deleted = retouchCollection.deleteSlide(activeSlideIndex)
    if (!deleted.ok) {
      return
    }

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
    const moved = retouchCollection.moveSlide(activeSlideIndex, direction)
    if (!moved.ok) {
      return
    }

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
    doc.commit(setSlideNamePatch(activeSlideIndex, nextName), {
      label: 'rename slide',
      origin: 'ppt-retouch',
    })
    clearTransientState()
  }

  function changeSlideAccent(accent: string) {
    if (accent === activeSlide.accent) {
      return
    }

    commitActiveTextEdit()
    doc.commit(setSlideAccentPatch(activeSlideIndex, accent), {
      label: 'change slide accent',
      origin: 'ppt-retouch',
    })
    clearTransientState()
  }

  function insertTextBlock() {
    commitActiveTextEdit()
    const nextBlock = createTextBlock(activeSlide)
    const blockIndex = activeSlide.blocks.length
    const pointer = blockPointer(activeSlideIndex, blockIndex)

    doc.commit(addBlocksPatch({
      blocks: [nextBlock],
      insertIndex: blockIndex,
      slideIndex: activeSlideIndex,
    }), {
      label: 'add text block',
      origin: 'ppt-retouch',
      selection: selectionSnapForPointers([pointer]),
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

  function copySelectedBlocks() {
    const locations = selectedActiveBlockLocations()

    if (locations.length === 0) {
      return
    }

    setBlockClipboard(locations.map((location) => ({ ...location.block })))
  }

  function pasteCopiedBlocks() {
    if (blockClipboard.length === 0) {
      return
    }

    commitActiveTextEdit()
    const pastedBlocks = duplicateBlocks(blockClipboard, activeSlide)
    const insertIndex = activeSlide.blocks.length
    const pastedPointers = blockInsertPointers({
      blocks: pastedBlocks,
      insertIndex,
      slideIndex: activeSlideIndex,
    })

    doc.commit(addBlocksPatch({
      blocks: pastedBlocks,
      insertIndex,
      slideIndex: activeSlideIndex,
    }), {
      label: 'paste blocks',
      origin: 'ppt-retouch',
      selection: selectionSnapForPointers(pastedPointers),
    })
    setBlockClipboard(pastedBlocks)
    setCanvasView('slide')
    setMode('layout')
    setEditing(null)
    clearLayoutInteraction()
  }

  function duplicateSelectedBlock() {
    const locations = selectedActiveBlockLocations()

    if (locations.length === 0) {
      return
    }

    commitActiveTextEdit()
    const duplicatedBlocks = duplicateBlocks(
      locations.map((location) => location.block),
      activeSlide,
    )
    const insertIndex = locations.at(-1)!.blockIndex + 1
    const duplicatePointers = blockInsertPointers({
      blocks: duplicatedBlocks,
      insertIndex,
      slideIndex: activeSlideIndex,
    })

    doc.commit(addBlocksPatch({
      blocks: duplicatedBlocks,
      insertIndex,
      slideIndex: activeSlideIndex,
    }), {
      label: locations.length > 1 ? 'duplicate blocks' : 'duplicate block',
      origin: 'ppt-retouch',
      selection: selectionSnapForPointers(duplicatePointers),
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

    const deleted = retouchCollection.deleteBlocks(
      locations.map((location) => location.pointer),
    )
    if (!deleted.ok) {
      return
    }

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
      selectionSnapForPointers(
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
      selectionSnapForPointers(
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
    const layerOrderPatch = createLayerOrderPatch({
      action,
      activeSlideIndex,
      doc,
      selectedIds,
      selectedPointers: locations.map((location) => location.pointer),
    })

    if (!layerOrderPatch) {
      return
    }

    commitActiveTextEdit()
    doc.commit(
      layerOrderPatch.operations,
      {
        label: 'reorder layers',
        origin: 'ppt-retouch',
        selection: selectionSnapForPointers(
          layerOrderPatch.nextSelectedPointers,
          layerOrderPatch.nextSelectedPointers.at(-1),
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

  function selectAllBlocks() {
    if (activeSlide.blocks.length === 0) {
      doc.selection?.empty()
      return
    }

    setCanvasView('slide')
    setMode('layout')
    clearTransientState()
    doc.selection?.selectRanges(
      activeSlide.blocks.map((_, blockIndex) =>
        blockPointer(activeSlideIndex, blockIndex),
      ),
    )
  }

  return (
    <main className="retouch-app" data-mode={mode}>
      <SlideRail
        activeSlideId={activeSlide.id}
        canMoveSlideDown={activeSlideIndex < doc.value.slides.length - 1}
        canMoveSlideUp={activeSlideIndex > 0}
        canvasViewTablistProps={canvasViewTablistProps}
        canvasViewTabProps={canvasViewTabProps}
        changedSlideIds={changedSlideIds}
        canDeleteSlide={doc.value.slides.length > 1}
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
        canZoomIn={canIncreaseCanvasZoom(canvasZoom)}
        canZoomOut={canDecreaseCanvasZoom(canvasZoom)}
        canvasView={canvasView}
        canvasViewPanelProps={canvasViewPanelProps}
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
        onZoomFit={() => setCanvasZoom('fit')}
        onZoomIn={() => setCanvasZoom((zoom) => nextCanvasZoom(zoom, 1))}
        onZoomOut={() => setCanvasZoom((zoom) => nextCanvasZoom(zoom, -1))}
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
        zoomLabel={canvasZoomLabel(canvasZoom)}
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
