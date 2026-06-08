import type {
  HTMLAttributes,
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
  AlignSelectionAction,
  DistributeSelectionAction,
} from './selectionAlignment'
import type { LayerOrderAction } from './selectionLayerOrder'
import type {
  DraftLayout,
  Interaction,
  Point,
  SnapGuides,
} from './layoutInteraction'
import { cssVariables } from './cssVariables'

type Mode = 'text' | 'layout'
type CanvasZoom = 'fit' | number

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
  canZoomIn: boolean
  canZoomOut: boolean
  canvasView: 'slide' | 'grid'
  canvasViewPanelProps: HTMLAttributes<HTMLDivElement>
  canvasZoom: CanvasZoom
  copyState: 'copied' | 'failed' | 'idle'
  copyTitle: string
  draftLayout: DraftLayout | null
  editing: EditingState | null
  exportCode: string
  exportCopied: boolean
  exportDownloaded: boolean
  exportTextareaRef: RefObject<HTMLTextAreaElement | null>
  interaction: Interaction | null
  marqueeRect: Rect | null
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
  onAlignSelection: (action: AlignSelectionAction) => void
  onCancelTextEdit: () => void
  onCanvasPointerDown: (event: ReactPointerEvent<HTMLDivElement>) => void
  onChangeMode: (mode: Mode) => void
  onCommitTextEdit: (pointer: Pointer, text: string, rect: Rect) => void
  onCopyExport: () => void
  onDeleteBlock: () => void
  onDownloadExport: () => void
  onDistributeSelection: (action: DistributeSelectionAction) => void
  onDuplicateBlock: () => void
  onInsertTextBlock: () => void
  onLayerOrderChange: (action: LayerOrderAction) => void
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
  onZoomFit: () => void
  onZoomIn: () => void
  onZoomOut: () => void
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
  zoomLabel: string
}

export function RetouchWorkspace({
  activeSlide,
  activeSlideAccent,
  activeSlideIndex,
  canRedo,
  canReset,
  canUndo,
  canZoomIn,
  canZoomOut,
  canvasView,
  canvasViewPanelProps,
  canvasZoom,
  copyState,
  copyTitle,
  draftLayout,
  editing,
  exportCode,
  exportCopied,
  exportDownloaded,
  exportTextareaRef,
  interaction,
  marqueeRect,
  mode,
  notes,
  onBlockClick,
  onBlockPointerDown,
  onAlignSelection,
  onCancelTextEdit,
  onCanvasPointerDown,
  onChangeMode,
  onCommitTextEdit,
  onCopyExport,
  onDeleteBlock,
  onDownloadExport,
  onDistributeSelection,
  onDuplicateBlock,
  onInsertTextBlock,
  onLayerOrderChange,
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
  onZoomFit,
  onZoomIn,
  onZoomOut,
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
  zoomLabel,
}: RetouchWorkspaceProps) {
  return (
    <section className="retouch-workspace">
      <Topbar
        canRedo={canRedo}
        canReset={canReset}
        canUndo={canUndo}
        canZoomIn={canZoomIn}
        canZoomOut={canZoomOut}
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
        onZoomFit={onZoomFit}
        onZoomIn={onZoomIn}
        onZoomOut={onZoomOut}
        resetScope={resetScope}
        resetTitle={resetTitle}
        zoomLabel={zoomLabel}
      />

      {canvasView === 'grid' ? (
        <div
          {...canvasViewPanelProps}
          className="stage-shell stage-shell-grid"
          ref={stageRef}
        >
          <DeckGrid
            activeSlideId={activeSlide.id}
            changedSlideIds={changedSlideIds}
            onOpenSlide={onOpenSlide}
            slides={slides}
          />
        </div>
      ) : (
        <div
          {...canvasViewPanelProps}
          className="stage-shell"
          data-zoom-mode={canvasZoom === 'fit' ? 'fit' : 'manual'}
          onClick={(event) => {
            if (suppressStageClickRef.current) {
              suppressStageClickRef.current = false
              return
            }

            const target = event.target instanceof HTMLElement ? event.target : null

            if (
              target?.closest(
                '[data-block], .selection-overlay, .resize-handle, .marquee-selection',
              )
            ) {
              return
            }

            onStageBackgroundClick()
          }}
          ref={stageRef}
          style={
            canvasZoom === 'fit'
              ? undefined
              : cssVariables({ '--canvas-zoom': canvasZoom })
          }
        >
          <StageCanvas
            activeSlide={activeSlide}
            activeSlideIndex={activeSlideIndex}
            draftLayout={draftLayout}
            editing={editing}
            interaction={interaction}
            marqueeRect={marqueeRect}
            mode={mode}
            onBlockClick={onBlockClick}
            onBlockPointerDown={onBlockPointerDown}
            onCanvasPointerDown={onCanvasPointerDown}
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
        onAlignSelection={onAlignSelection}
        onDeleteBlock={onDeleteBlock}
        onDistributeSelection={onDistributeSelection}
        onDuplicateBlock={onDuplicateBlock}
        onLayerOrderChange={onLayerOrderChange}
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
