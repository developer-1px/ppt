import type {
  MouseEvent as ReactMouseEvent,
  PointerEvent as ReactPointerEvent,
  RefObject,
} from 'react'
import type { ResizeHandle } from 'canvas/core'
import type { Pointer } from 'zod-crud'
import {
  blockPointer,
  findBlockLocation,
  type Rect,
  type RetouchSlide,
  type SlideBlock,
} from './retouchModel'
import { SAMPLE_DECK } from './sampleDeck'
import { SlideBlockElement } from './SlideBlockElement'
import { StageLayoutChrome } from './StageLayoutChrome'
import { htmlSlideRootAttributes } from './htmlSlideContract'
import {
  getCurrentRect,
  minimumHeightForBlock,
  resizeHandleAffectsHeight,
  type DraftLayout,
  type Interaction,
  type SnapGuides,
} from './layoutInteraction'
import { cssVariables } from './cssVariables'
import type { EditingState, RetouchMode } from './retouchViewState'
import './StageCanvas.css'

type StageCanvasProps = {
  activeSlide: RetouchSlide
  activeSlideIndex: number
  draftLayout: DraftLayout | null
  editing: EditingState | null
  interaction: Interaction | null
  marqueeRect: Rect | null
  mode: RetouchMode
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
      blockIndex,
      pointer,
      rect: getCurrentRect(pointer, block, draftLayout),
      selected: selectedPointerSet.has(pointer),
    }
  })
  const surfaceItems = blockEntries.map(({ pointer, rect }) => ({
    pointer,
    rect,
  }))

  return (
    <div className="slide-frame">
      <div
        className="slide-canvas"
        {...htmlSlideRootAttributes(activeSlide.id)}
        onPointerDown={onCanvasPointerDown}
        ref={slideRef}
        style={cssVariables({ '--accent': activeSlide.accent })}
      >
        {blockEntries.map(({ block, blockIndex, pointer, rect, selected }) => {
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
              blockIndex={blockIndex}
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

        {mode === 'layout' ? (
          <StageLayoutChrome
            marqueeRect={marqueeRect}
            onResizePointerDown={onResizePointerDown}
            selectedPointers={selectedPointers}
            selectedRect={selectedRect}
            snapGuides={snapGuides}
            surfaceItems={surfaceItems}
            visualSelectionRect={visualSelectionRect}
          />
        ) : null}
      </div>
    </div>
  )
}
