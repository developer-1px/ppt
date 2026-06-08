import {
  useCallback,
  useMemo,
  useRef,
  useState,
  type MouseEvent as ReactMouseEvent,
  type PointerEvent as ReactPointerEvent,
} from 'react'
import { isAdditivePointerInput } from 'canvas/foundation'
import type { JSONPatchOperation, Pointer, SelectionSnap } from 'zod-crud'
import { useJSONDocument } from 'zod-crud/react'
import {
  RetouchDeckSchema,
} from './retouchModel'
import { SAMPLE_DECK, SAMPLE_SLIDES } from './sampleDeck'
import { exportRetouchDeck } from './retouchExport'
import { PresentationOverlay } from './PresentationOverlay'
import { RetouchWorkspace } from './RetouchWorkspace'
import { SlideRail } from './SlideRail'
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
import { useRetouchSlideCommands } from './useRetouchSlideCommands'
import { useRetouchBlockCommands } from './useRetouchBlockCommands'
import { useRetouchPresentationCommands } from './useRetouchPresentationCommands'
import {
  getCurrentRect,
} from './layoutInteraction'
import { selectionSnapForPointers } from './retouchSelectionSnap'
import { useRetouchViewCommands } from './useRetouchViewCommands'
import {
  canDecreaseCanvasZoom,
  canIncreaseCanvasZoom,
  canvasZoomLabel,
  nextCanvasZoom,
  type CanvasZoom,
} from './canvasZoom'
import type { RetouchPatchCommit } from './retouchSurfaceContract'
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

  const commitRetouchPatch = useCallback<RetouchPatchCommit>(
    (patch, options) => {
      if (patch.length === 0) {
        return
      }

      const commitOptions =
        typeof options === 'string' ? { label: options } : options

      doc.commit(patch, {
        ...commitOptions,
        origin: 'ppt-retouch',
      })
    },
    [doc],
  )

  const commitPatch = useCallback(
    (
      patch: JSONPatchOperation[],
      pointer: Pointer,
      label: string,
      mergeKey?: string,
      selection?: SelectionSnap,
    ) => {
      commitRetouchPatch(patch, {
        label,
        mergeKey,
        selection: selection ?? selectionSnapForPointers([pointer]),
      })
    },
    [commitRetouchPatch],
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
      commitRetouchPatch([{ op: 'replace', path: '', value: SAMPLE_DECK }], 'reset deck')
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

  function clearTransientState() {
    setEditing(null)
    clearLayoutInteraction()
    clearMarqueeSelection()
  }

  const {
    canvasViewPanelProps,
    canvasViewTabProps,
    canvasViewTablistProps,
    changeMode,
    enterLayoutMode,
  } = useRetouchViewCommands({
    canvasView,
    clearSelection: () => doc.selection?.empty(),
    clearTransientState,
    commitActiveTextEdit,
    mode,
    setCanvasView,
    setMode,
  })

  const {
    activeSlideNotes,
    addSlide,
    changeActiveSlideNotes,
    changeSlideAccent,
    changeSlideName,
    copySlide,
    deleteSlide,
    moveSlide,
    selectSlide,
  } = useRetouchSlideCommands({
    activeSlide,
    activeSlideIndex,
    clearTransientState,
    commitActiveTextEdit,
    commitRetouchPatch,
    doc,
    retouchCollection,
    setActiveSlideId,
    setCanvasView,
    stageRef,
  })

  const {
    alignSelectedBlocks,
    canPasteSelection,
    changeSelectedBlockRect,
    changeSelectedLayerOrder,
    copySelectedBlocks,
    deleteSelectedBlock,
    distributeSelectedBlocks,
    duplicateSelectedBlock,
    insertTextBlock,
    pasteCopiedBlocks,
    selectAllBlocks,
  } = useRetouchBlockCommands({
    activeSlide,
    activeSlideIndex,
    clearLayoutInteraction,
    commitActiveTextEdit,
    commitPatch,
    commitRetouchPatch,
    doc,
    enterLayoutMode,
    retouchCollection,
    selectedBlock,
    selectedPointer,
    selectedPointers,
    selectedRect,
    setCanvasView,
    setEditing,
    setMode,
    stageRef,
  })

  useRetouchKeyboardShortcuts({
    activeSlideId: activeSlide.id,
    commitPatch,
    deckValue: doc.value,
    editing,
    history: doc.history,
    interaction,
    mode,
    canPasteSelection,
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

  const {
    closePresentation,
    navigatePresentation,
    presenting,
    startPresentation,
  } = useRetouchPresentationCommands({
    activeSlide,
    activeSlideIndex,
    clearTransientState,
    commitActiveTextEdit,
    doc,
    setActiveSlideId,
    setCanvasView,
  })

  function selectBlock(pointer: Pointer, additive = false) {
    if (additive) {
      doc.selection?.togglePointer(pointer)
      return
    }

    doc.selection?.selectRanges([pointer])
  }

  function handleBlockClick(
    event: ReactPointerEvent<HTMLElement> | ReactMouseEvent<HTMLElement>,
    pointer: Pointer,
  ) {
    if (suppressBlockClickRef.current) {
      suppressBlockClickRef.current = false
      return
    }

    if (mode === 'layout') {
      selectBlock(pointer, isAdditivePointerInput(event))
      return
    }

    startTextEdit(pointer, {
      x: event.clientX,
      y: event.clientY,
    })
  }

  function handleStageBackgroundClick() {
    if (mode === 'text') {
      commitActiveTextEdit()
    }
    doc.selection?.empty()
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
        notes={activeSlideNotes}
        onBlockClick={handleBlockClick}
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
        onNotesChange={changeActiveSlideNotes}
        onOpenSlide={selectSlide}
        onPresent={startPresentation}
        onRedo={redoDocumentChange}
        onReset={resetCurrentTarget}
        onResizePointerDown={startResizeInteraction}
        onSelectedRectChange={changeSelectedBlockRect}
        onStageBackgroundClick={handleStageBackgroundClick}
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
          notes={activeSlideNotes}
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
