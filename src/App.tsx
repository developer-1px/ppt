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
  SAMPLE_DECK,
  SAMPLE_SLIDES,
  RetouchDeckSchema,
  blockPointer,
  findSlideIndex,
  slideAccentPointer,
  slideNamePointer,
  slidePointer,
} from './retouchModel'
import { exportRetouchDeck } from './retouchExport'
import { PresentationOverlay } from './PresentationOverlay'
import { RetouchWorkspace } from './RetouchWorkspace'
import { SlideRail } from './SlideRail'
import { createTextBlock } from './slideBlockOperations'
import { createBlankSlide, duplicateSlide } from './slideDeckOperations'
import { useExportControls } from './useExportControls'
import { useRetouchLayoutInteraction } from './useRetouchLayoutInteraction'
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
  type Point,
} from './layoutInteraction'
import './App.css'

type Mode = 'text' | 'layout'
type CanvasView = 'slide' | 'grid'

type EditingState = {
  clientPoint?: Point
  pointer: Pointer
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
        canvasView={canvasView}
        changedSlideIds={changedSlideIds}
        canDeleteSlide={doc.value.slides.length > 1}
        onChangeCanvasView={changeCanvasView}
        onAddSlide={addSlide}
        onDeleteSlide={deleteSlide}
        onDuplicateSlide={copySlide}
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
        canvasView={canvasView}
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
        onDownloadExport={downloadExportCode}
        onInsertTextBlock={insertTextBlock}
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
