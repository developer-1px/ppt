import type {
  MouseEvent as ReactMouseEvent,
  MutableRefObject,
  PointerEvent as ReactPointerEvent,
  RefObject,
} from 'react'
import type { Pointer } from 'zod-crud'
import { DeckGrid } from './DeckGrid'
import { InspectorPanel } from './InspectorPanel'
import { StageCanvas } from './StageCanvas'
import { Topbar } from './Topbar'
import type {
  Rect,
  ResizeHandle,
  RetouchSlide,
  SlideBlock,
} from './retouchModel'
import type {
  DraftLayout,
  Interaction,
  Point,
  SnapGuides,
} from './layoutInteraction'

type Mode = 'text' | 'layout'

type EditingState = {
  clientPoint?: Point
  pointer: Pointer
}

type RetouchWorkspaceProps = {
  activeSlide: RetouchSlide
  activeSlideAccent: string
  activeSlideIndex: number
  canRedo: boolean
  canReset: boolean
  canUndo: boolean
  canvasView: 'slide' | 'grid'
  copyState: 'copied' | 'failed' | 'idle'
  copyTitle: string
  draftLayout: DraftLayout | null
  editing: EditingState | null
  exportCode: string
  exportCopied: boolean
  exportDownloaded: boolean
  exportTextareaRef: RefObject<HTMLTextAreaElement | null>
  interaction: Interaction | null
  mode: Mode
  notes: string
  onBlockClick: (
    event: ReactPointerEvent<HTMLElement> | ReactMouseEvent<HTMLElement>,
    pointer: Pointer,
  ) => void
  onBlockPointerDown: (
    event: ReactPointerEvent<HTMLElement>,
    pointer: Pointer,
    block: SlideBlock,
  ) => void
  onCancelTextEdit: () => void
  onChangeMode: (mode: Mode) => void
  onCommitTextEdit: (pointer: Pointer, text: string, rect: Rect) => void
  onCopyExport: () => void
  onDeleteBlock: () => void
  onDownloadExport: () => void
  onDuplicateBlock: () => void
  onInsertTextBlock: () => void
  onNotesChange: (notes: string) => void
  onOpenSlide: (slideId: string) => void
  onPresent: () => void
  onRedo: () => void
  onReset: () => void
  onResizePointerDown: (
    event: ReactPointerEvent<HTMLButtonElement>,
    handle: ResizeHandle,
  ) => void
  onSelectedRectChange: (rect: Rect, changedField?: keyof Rect) => void
  onStageBackgroundClick: () => void
  onSlideAccentChange: (accent: string) => void
  onSlideNameChange: (name: string) => void
  onUndo: () => void
  resetScope: 'deck' | 'layout' | 'text'
  resetTitle: string
  selectedPointerSet: Set<Pointer>
  selectedBlock: SlideBlock | null
  selectedPointers: Pointer[]
  selectedRect: Rect | null
  changedSlideIds: Set<string>
  slides: RetouchSlide[]
  slideRef: RefObject<HTMLDivElement | null>
  snapGuides: SnapGuides
  stageRef: RefObject<HTMLDivElement | null>
  suppressStageClickRef: MutableRefObject<boolean>
  visualSelectionRect: Rect | null
}

export function RetouchWorkspace({
  activeSlide,
  activeSlideAccent,
  activeSlideIndex,
  canRedo,
  canReset,
  canUndo,
  canvasView,
  copyState,
  copyTitle,
  draftLayout,
  editing,
  exportCode,
  exportCopied,
  exportDownloaded,
  exportTextareaRef,
  interaction,
  mode,
  notes,
  onBlockClick,
  onBlockPointerDown,
  onCancelTextEdit,
  onChangeMode,
  onCommitTextEdit,
  onCopyExport,
  onDeleteBlock,
  onDownloadExport,
  onDuplicateBlock,
  onInsertTextBlock,
  onNotesChange,
  onOpenSlide,
  onPresent,
  onRedo,
  onReset,
  onResizePointerDown,
  onSelectedRectChange,
  onStageBackgroundClick,
  onSlideAccentChange,
  onSlideNameChange,
  onUndo,
  resetScope,
  resetTitle,
  changedSlideIds,
  selectedBlock,
  selectedPointerSet,
  selectedPointers,
  selectedRect,
  slides,
  slideRef,
  snapGuides,
  stageRef,
  suppressStageClickRef,
  visualSelectionRect,
}: RetouchWorkspaceProps) {
  return (
    <section className="retouch-workspace">
      <Topbar
        canRedo={canRedo}
        canReset={canReset}
        canUndo={canUndo}
        copyState={copyState}
        copyTitle={copyTitle}
        exportCopied={exportCopied}
        exportDownloaded={exportDownloaded}
        mode={mode}
        onChangeMode={onChangeMode}
        onCopyExport={onCopyExport}
        onDownloadExport={onDownloadExport}
        onInsertTextBlock={onInsertTextBlock}
        onPresent={onPresent}
        onRedo={onRedo}
        onReset={onReset}
        onUndo={onUndo}
        resetScope={resetScope}
        resetTitle={resetTitle}
      />

      {canvasView === 'grid' ? (
        <div className="stage-shell stage-shell-grid" ref={stageRef}>
          <DeckGrid
            activeSlideId={activeSlide.id}
            changedSlideIds={changedSlideIds}
            onOpenSlide={onOpenSlide}
            slides={slides}
          />
        </div>
      ) : (
        <div
          className="stage-shell"
          onClick={(event) => {
            if (suppressStageClickRef.current) {
              suppressStageClickRef.current = false
              return
            }

            const target = event.target instanceof HTMLElement ? event.target : null

            if (target?.closest('[data-block], .selection-overlay, .resize-handle')) {
              return
            }

            onStageBackgroundClick()
          }}
          ref={stageRef}
        >
          <StageCanvas
            activeSlide={activeSlide}
            activeSlideIndex={activeSlideIndex}
            draftLayout={draftLayout}
            editing={editing}
            interaction={interaction}
            mode={mode}
            onBlockClick={onBlockClick}
            onBlockPointerDown={onBlockPointerDown}
            onCancelTextEdit={onCancelTextEdit}
            onCommitTextEdit={onCommitTextEdit}
            onResizePointerDown={onResizePointerDown}
            selectedPointerSet={selectedPointerSet}
            selectedPointers={selectedPointers}
            selectedRect={selectedRect}
            slideRef={slideRef}
            snapGuides={snapGuides}
            visualSelectionRect={visualSelectionRect}
          />
        </div>
      )}

      <textarea
        aria-hidden="true"
        className="export-buffer"
        readOnly
        ref={exportTextareaRef}
        tabIndex={-1}
        value={exportCode}
      />

      <InspectorPanel
        activeSlideAccent={activeSlideAccent}
        activeSlideName={activeSlide.name}
        canvasView={canvasView}
        mode={mode}
        notes={notes}
        onDeleteBlock={onDeleteBlock}
        onDuplicateBlock={onDuplicateBlock}
        onNotesChange={onNotesChange}
        onPresent={onPresent}
        onSelectedRectChange={onSelectedRectChange}
        onSlideAccentChange={onSlideAccentChange}
        onSlideNameChange={onSlideNameChange}
        selectedBlock={selectedBlock}
        selectedCount={selectedPointers.length}
        selectedRect={selectedRect}
      />
    </section>
  )
}
