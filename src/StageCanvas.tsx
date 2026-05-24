import type {
  CSSProperties,
  MouseEvent as ReactMouseEvent,
  PointerEvent as ReactPointerEvent,
  RefObject,
} from 'react'
import type { Pointer } from 'zod-crud'
import {
  RESIZE_HANDLES,
  SAMPLE_DECK,
  SLIDE_HEIGHT,
  SLIDE_WIDTH,
  blockPointer,
  findBlockLocation,
  rectToStyle,
  type Rect,
  type ResizeHandle,
  type RetouchSlide,
  type SlideBlock,
} from './retouchModel'
import { SlideBlockElement } from './SlideBlockElement'
import {
  getCurrentRect,
  minimumHeightForBlock,
  resizeHandleAffectsHeight,
  type DraftLayout,
  type Interaction,
  type Point,
  type SnapGuides,
} from './layoutInteraction'

type EditingState = {
  clientPoint?: Point
  pointer: Pointer
}

type StageCanvasProps = {
  activeSlide: RetouchSlide
  activeSlideIndex: number
  draftLayout: DraftLayout | null
  editing: EditingState | null
  interaction: Interaction | null
  marqueeRect: Rect | null
  mode: 'text' | 'layout'
  onBlockClick: (
    event: ReactPointerEvent<HTMLElement> | ReactMouseEvent<HTMLElement>,
    pointer: Pointer,
  ) => void
  onBlockPointerDown: (
    event: ReactPointerEvent<HTMLElement>,
    pointer: Pointer,
    block: SlideBlock,
  ) => void
  onCanvasPointerDown: (event: ReactPointerEvent<HTMLDivElement>) => void
  onCancelTextEdit: () => void
  onCommitTextEdit: (pointer: Pointer, text: string, rect: Rect) => void
  onResizePointerDown: (
    event: ReactPointerEvent<HTMLButtonElement>,
    handle: ResizeHandle,
  ) => void
  selectedPointerSet: Set<Pointer>
  selectedPointers: Pointer[]
  selectedRect: Rect | null
  slideRef: RefObject<HTMLDivElement | null>
  snapGuides: SnapGuides
  visualSelectionRect: Rect | null
}

export function StageCanvas({
  activeSlide,
  activeSlideIndex,
  draftLayout,
  editing,
  interaction,
  marqueeRect,
  mode,
  onBlockClick,
  onBlockPointerDown,
  onCanvasPointerDown,
  onCancelTextEdit,
  onCommitTextEdit,
  onResizePointerDown,
  selectedPointerSet,
  selectedPointers,
  selectedRect,
  slideRef,
  snapGuides,
  visualSelectionRect,
}: StageCanvasProps) {
  const blockEntries = activeSlide.blocks.map((block, blockIndex) => {
    const pointer = blockPointer(activeSlideIndex, blockIndex)

    return {
      block,
      pointer,
      rect: getCurrentRect(pointer, block, draftLayout),
      selected: selectedPointerSet.has(pointer),
    }
  })
  const selectedBounds = rectBounds(
    blockEntries
      .filter(({ pointer }) => selectedPointerSet.has(pointer))
      .map(({ rect }) => rect),
  )
  const overlayRect =
    selectedPointers.length === 1 && selectedRect
      ? (visualSelectionRect ?? selectedRect)
      : selectedBounds

  return (
    <div className="slide-frame">
      <div
        className="slide-canvas"
        data-slide={activeSlide.id}
        onPointerDown={onCanvasPointerDown}
        ref={slideRef}
        style={{ '--accent': activeSlide.accent } as CSSProperties}
      >
        {blockEntries.map(({ block, pointer, rect, selected }) => {
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
          const editingThisBlock = mode === 'text' && editing?.pointer === pointer
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
              initialClientPoint={editingThisBlock ? editing.clientPoint : undefined}
              key={block.id}
              minimumHeight={minimumHeight}
              onCancel={onCancelTextEdit}
              onClick={(event) => onBlockClick(event, pointer)}
              onPointerDown={(event) => onBlockPointerDown(event, pointer, block)}
              onCommit={(text, nextRect) => onCommitTextEdit(pointer, text, nextRect)}
              rect={rect}
              selected={selected}
              text={block.text}
            />
          )
        })}

        {mode === 'layout' && overlayRect ? (
          <SelectionOverlay
            onResizePointerDown={onResizePointerDown}
            rect={overlayRect}
            resizable={selectedPointers.length === 1}
          />
        ) : null}

        {mode === 'layout' && marqueeRect ? (
          <div className="marquee-selection" style={rectToStyle(marqueeRect)} />
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
  )
}

function SelectionOverlay({
  onResizePointerDown,
  rect,
  resizable,
}: {
  onResizePointerDown: (
    event: ReactPointerEvent<HTMLButtonElement>,
    handle: ResizeHandle,
  ) => void
  rect: Rect
  resizable: boolean
}) {
  return (
    <div
      className="selection-overlay"
      data-resizable={resizable ? 'true' : 'false'}
      style={rectToStyle(rect)}
    >
      {resizable
        ? RESIZE_HANDLES.map((handle) => (
            <button
              aria-label={`Resize ${handle}`}
              className="resize-handle"
              data-handle={handle}
              key={handle}
              onClick={(event) => event.stopPropagation()}
              onPointerDown={(event) => onResizePointerDown(event, handle)}
              type="button"
            />
          ))
        : null}
    </div>
  )
}

function rectBounds(rects: Rect[]) {
  if (rects.length === 0) {
    return null
  }

  const left = Math.min(...rects.map((rect) => rect.x))
  const top = Math.min(...rects.map((rect) => rect.y))
  const right = Math.max(...rects.map((rect) => rect.x + rect.width))
  const bottom = Math.max(...rects.map((rect) => rect.y + rect.height))

  return {
    x: left,
    y: top,
    width: right - left,
    height: bottom - top,
  }
}
