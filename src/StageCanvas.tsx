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
  mode,
  onBlockClick,
  onBlockPointerDown,
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
  return (
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

        {mode === 'layout' && selectedPointers.length === 1 && selectedRect ? (
          <SelectionOverlay
            onResizePointerDown={onResizePointerDown}
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
