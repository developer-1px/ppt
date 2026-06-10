import type { PointerEvent as ReactPointerEvent } from 'react'
import {
  RESIZE_HANDLES,
  type ResizeHandle,
} from 'canvas/core'
import type { Pointer } from 'zod-crud'
import {
  SLIDE_HEIGHT,
  SLIDE_WIDTH,
  rectToStyle,
  type Rect,
} from './retouchModel'
import type { SnapGuides } from './layoutInteraction'
import {
  retouchCanvasSceneEntries,
  retouchCanvasSelectionBounds,
} from './retouchCanvasScene'

type StageLayoutChromeProps = {
  marqueeRect: Rect | null
  onResizePointerDown: (
    event: ReactPointerEvent<HTMLButtonElement>,
    handle: ResizeHandle,
  ) => void
  selectedPointers: readonly Pointer[]
  selectedRect: Rect | null
  snapGuides: SnapGuides
  surfaceItems: readonly { pointer: Pointer; rect: Rect }[]
  visualSelectionRect: Rect | null
}

export function StageLayoutChrome({
  marqueeRect,
  onResizePointerDown,
  selectedPointers,
  selectedRect,
  snapGuides,
  surfaceItems,
  visualSelectionRect,
}: StageLayoutChromeProps) {
  const singleSelectionRect =
    selectedPointers.length === 1 && selectedRect !== null
      ? (visualSelectionRect ?? selectedRect)
      : null
  const selectedBounds =
    singleSelectionRect !== null || selectedPointers.length === 0
      ? null
      : retouchCanvasSelectionBounds(
          retouchCanvasSceneEntries(surfaceItems),
          selectedPointers,
        )
  const overlayRect = singleSelectionRect ?? selectedBounds

  return (
    <>
      {overlayRect !== null ? (
        <SelectionOverlay
          onResizePointerDown={onResizePointerDown}
          rect={overlayRect}
          resizable={selectedPointers.length === 1}
        />
      ) : null}

      {marqueeRect !== null ? (
        <div className="marquee-selection" style={rectToStyle(marqueeRect)} />
      ) : null}

      {snapGuides.x !== null ? (
        <div
          className="snap-guide snap-guide-x"
          style={{ left: `${(snapGuides.x / SLIDE_WIDTH) * 100}%` }}
        />
      ) : null}
      {snapGuides.y !== null ? (
        <div
          className="snap-guide snap-guide-y"
          style={{ top: `${(snapGuides.y / SLIDE_HEIGHT) * 100}%` }}
        />
      ) : null}
    </>
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
