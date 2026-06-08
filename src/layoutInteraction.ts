import { clamp, type Point } from 'canvas/core'
import type { Pointer, SelectionSnap } from 'zod-crud'
import {
  EMPTY_TEXT_BOX_HEIGHT,
  GRID_SIZE,
  SLIDE_HEIGHT,
  SLIDE_WIDTH,
  blockPointer,
  getRect,
  moveRect,
  rectEquals,
  type Rect,
  type ResizeHandle,
  type SlideBlock,
} from './retouchModel'

export type Interaction =
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

export type { Point }

export type DraftLayout = {
  rects: DraftLayoutRect[]
}

export type DraftLayoutRect = {
  pointer: Pointer
  rect: Rect
}

export type SnapGuides = {
  x: number | null
  y: number | null
}

type SelectionModifierEvent = {
  ctrlKey: boolean
  metaKey: boolean
  shiftKey: boolean
}

export function getCurrentRect(
  pointer: Pointer,
  block: SlideBlock,
  draftLayout: DraftLayout | null,
) {
  return (
    draftLayout?.rects.find((draftRect) => draftRect.pointer === pointer)?.rect ??
    getRect(block)
  )
}

export function calculateMoveInteractionState(
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

export function draftRectsEqual(a: DraftLayoutRect[], b: DraftLayoutRect[]) {
  return (
    a.length === b.length &&
    a.every((draftRect) => {
      const other = b.find((candidate) => candidate.pointer === draftRect.pointer)

      return other ? rectEquals(draftRect.rect, other.rect) : false
    })
  )
}

export function selectionSnapForPointers(
  pointers: Pointer[],
  primaryPointer = pointers.at(-1),
): SelectionSnap {
  const selectedPointers = [...new Set(pointers)]
  const fallbackPrimary = selectedPointers.at(-1) ?? null
  const primary = primaryPointer && selectedPointers.includes(primaryPointer)
    ? primaryPointer
    : fallbackPrimary
  const primaryIndex = primary ? selectedPointers.indexOf(primary) : -1

  return {
    anchor: primary,
    focus: primary,
    primaryIndex,
    selectedPointers,
    selectionRanges: selectedPointers.map((pointer) => ({
      anchor: pointer,
      focus: pointer,
    })),
  }
}

export function hasSelectionModifier(event: SelectionModifierEvent) {
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

export function guidesForInteraction(rect: Rect, interaction: Interaction): SnapGuides {
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

export function rectClose(a: Rect, b: Rect) {
  return (
    Math.abs(a.x - b.x) < 0.5 &&
    Math.abs(a.y - b.y) < 0.5 &&
    Math.abs(a.width - b.width) < 0.5 &&
    Math.abs(a.height - b.height) < 0.5
  )
}

export function resizeHandleAffectsHeight(handle: ResizeHandle) {
  return handle.includes('n') || handle.includes('s')
}

export function minimumHeightForBlock(
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
    (baseBlock &&
      block.text === baseBlock.text &&
      (block.height !== baseBlock.height || block.width !== baseBlock.width))
  ) {
    return rect.height
  }

  return 0
}

export function arrangeResetEquals(a: SlideBlock, b: SlideBlock) {
  const heightChangedByArrange = a.text === b.text && a.height !== b.height

  return (
    a.x === b.x &&
    a.y === b.y &&
    a.width === b.width &&
    !heightChangedByArrange
  )
}

export function textResetEquals(a: SlideBlock, b: SlideBlock) {
  return a.text === b.text && a.height === b.height
}
