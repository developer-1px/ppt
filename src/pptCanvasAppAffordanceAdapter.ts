import {
  CANVAS_COMMAND_PALETTE_ITEMS_MODEL,
  filterCanvasCommandPaletteItems,
  type CanvasCommandPaletteItem,
} from 'canvas/app/command-palette-items'
import {
  CANVAS_MINIMAP_READ_MODEL,
  getCanvasMinimapPointFromViewportOffset,
  getCanvasMinimapReadModel,
  getCanvasMinimapWorldPoint,
  type CanvasMinimapItemBounds,
  type CanvasMinimapReadModel,
  type CanvasMinimapSize,
} from 'canvas/app/minimap-model'
import {
  CANVAS_PASTE_POSITION_MODEL,
} from 'canvas/app/paste-position'
import {
  CANVAS_KEYBOARD_TEMPORARY_PAN_MODEL,
  CANVAS_KEYBOARD_TEMPORARY_PAN_SHORTCUT_LABEL,
} from 'canvas/app/keyboard-system-shortcuts'
import {
  CANVAS_WHEEL_VIEWPORT_HORIZONTAL_PAN_MODIFIER,
  CANVAS_WHEEL_VIEWPORT_MODEL,
  CANVAS_WHEEL_VIEWPORT_PAN_MODE,
  CANVAS_WHEEL_VIEWPORT_ZOOM_MODIFIER,
} from 'canvas/app/viewport-controls'
import {
  CANVAS_INLINE_EDIT_DOM_MODEL,
  inlineEditHistoryDirectionFromInputType,
  insertInlineEditText,
  isInlineEditLineBreakInput,
} from 'canvas/app/inline-edit-dom'
import {
  CANVAS_POINTER_CLICK_MEMORY_MODEL,
} from 'canvas/app/pointer-click-memory'

export const PPT_COMMAND_PALETTE_ITEMS_MODEL =
  CANVAS_COMMAND_PALETTE_ITEMS_MODEL
export const PPT_MINIMAP_READ_MODEL = CANVAS_MINIMAP_READ_MODEL
export const PPT_PASTE_POSITION_MODEL = CANVAS_PASTE_POSITION_MODEL
export const PPT_KEYBOARD_TEMPORARY_PAN_MODEL =
  CANVAS_KEYBOARD_TEMPORARY_PAN_MODEL
export const PPT_KEYBOARD_TEMPORARY_PAN_SHORTCUT_LABEL =
  CANVAS_KEYBOARD_TEMPORARY_PAN_SHORTCUT_LABEL
export const PPT_WHEEL_VIEWPORT_HORIZONTAL_PAN_MODIFIER =
  CANVAS_WHEEL_VIEWPORT_HORIZONTAL_PAN_MODIFIER
export const PPT_WHEEL_VIEWPORT_MODEL = CANVAS_WHEEL_VIEWPORT_MODEL
export const PPT_WHEEL_VIEWPORT_PAN_MODE = CANVAS_WHEEL_VIEWPORT_PAN_MODE
export const PPT_WHEEL_VIEWPORT_ZOOM_MODIFIER =
  CANVAS_WHEEL_VIEWPORT_ZOOM_MODIFIER
export const PPT_INLINE_EDIT_DOM_MODEL = CANVAS_INLINE_EDIT_DOM_MODEL
export const PPT_POINTER_CLICK_MEMORY_MODEL =
  CANVAS_POINTER_CLICK_MEMORY_MODEL

export const filterPPTCommandPaletteItems =
  filterCanvasCommandPaletteItems
export const getPPTMinimapPointFromViewportOffset =
  getCanvasMinimapPointFromViewportOffset
export const getPPTMinimapReadModel = getCanvasMinimapReadModel
export const getPPTMinimapWorldPoint = getCanvasMinimapWorldPoint
export const getPPTInlineEditHistoryDirectionFromInputType =
  inlineEditHistoryDirectionFromInputType
export const insertPPTInlineEditText = insertInlineEditText
export const isPPTInlineEditLineBreakInput = isInlineEditLineBreakInput

export type PPTCommandPaletteItemBase = CanvasCommandPaletteItem
export type PPTMinimapItemBounds = CanvasMinimapItemBounds
export type PPTMinimapReadModel = CanvasMinimapReadModel
export type PPTMinimapSize = CanvasMinimapSize
