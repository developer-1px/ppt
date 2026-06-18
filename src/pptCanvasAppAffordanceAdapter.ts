import {
  CANVAS_COMMENT_THREAD_MODEL,
  CANVAS_KEYBOARD_COMMAND_DISPATCH_MODEL,
  CANVAS_KEYBOARD_COMMAND_INTENT_MODEL,
  CANVAS_KEYBOARD_NUDGE_INTENT_MODEL,
  CANVAS_KEYBOARD_NUDGE_KEYS,
  CANVAS_KEYBOARD_NUDGE_LARGE_STEP,
  CANVAS_KEYBOARD_NUDGE_MODEL,
  CANVAS_KEYBOARD_NUDGE_STEP,
  CANVAS_KEYBOARD_TOOL_DISPATCH_MODEL,
  CANVAS_KEYBOARD_VIEWPORT_INTENT_MODEL,
  CANVAS_KEYBOARD_VIEWPORT_MODEL,
  CANVAS_MENU_FOCUS_MODEL,
  CANVAS_MENU_ITEM_PROPS,
  CANVAS_MENU_KEYBOARD_KEYS,
  CANVAS_MENU_ROVING_FOCUS_MODEL,
  CANVAS_MODAL_FOCUS_LIFECYCLE_MODEL,
  CANVAS_RADIO_GROUP_FOCUS_MODEL,
  CANVAS_RADIO_GROUP_KEYBOARD_MODEL,
  CANVAS_RADIO_GROUP_MODEL,
  CANVAS_RESIZE_POINTER_MODIFIERS_MODEL,
  CANVAS_SELECTION_TOOLBAR_DROPDOWN_MENU_MODEL,
  CANVAS_TABS_ROVING_FOCUS_MODEL,
  CANVAS_TOOLBAR_FOCUS_MODEL,
  CANVAS_TOOLBAR_ITEM_PROPS,
  CANVAS_TOOLBAR_KEYBOARD_MODEL,
  CANVAS_TOOLBAR_ROVING_FOCUS_MODEL,
  bindCanvasEventListener,
  bindCanvasEventListeners,
  cancelCanvasAnimationFrameTask,
  cancelCanvasDeferredFocus,
  captureCanvasPointerFromEvent,
  centerCanvasViewportAtWorldPoint,
  createCanvasDataTransferImportActionPlan,
  createCanvasPastePositionKey,
  createCanvasRichClipboardHTML,
  createCanvasTabsDescriptor,
  downloadCanvasTextFile,
  fitCanvasViewportToBounds,
  focusCanvasElement,
  focusCanvasElementBySelectorOnNextFrame,
  focusCanvasElementOnNextFrame,
  getCanvasClientViewportSize,
  getCanvasCommandPaletteKeyboardIntent,
  getCanvasContextMenuKeyboardIntent,
  getCanvasContextMenuPosition,
  getCanvasDataTransferText,
  getCanvasEditableFieldKeyboardIntent,
  getCanvasFindInputKeyboardIntent,
  getCanvasFloatingAnchorForBounds,
  getCanvasImportedImageSize,
  getCanvasInlineEditKeyboardIntent,
  getCanvasKeyboardBuiltinCommandShortcutIntent,
  getCanvasKeyboardNudgeShortcutIntent,
  getCanvasKeyboardSystemShortcutIntent,
  getCanvasKeyboardToolShortcutIntent,
  getCanvasKeyboardViewportShortcutIntent,
  getCanvasMenuTriggerKeyboardIntent,
  getCanvasModalBackdropPointerIntent,
  getCanvasModalKeyboardIntent,
  getCanvasPasteOffsetForBounds,
  getCanvasPastePositionSession,
  getCanvasPointerLocalGeometry,
  getCanvasPointerTransformModifierState,
  getCanvasPresentationKeyboardIntent,
  getCanvasRadioTabIndex,
  getCanvasResizeHandleDoubleClickIntent,
  getCanvasSelectionListModifierState,
  getCanvasTabsKeyboardIntent,
  getCanvasWorldClientPoint,
  handleCanvasRadioGroupKeyDown,
  isCanvasControlTarget,
  isCanvasKeyboardCommandIntent,
  isCanvasKeyboardToolIntent,
  isCanvasKeyboardTypingTarget,
  isCanvasKeyboardViewportIntent,
  isCanvasTargetWithinSelector,
  isCanvasWheelPassthroughTarget,
  measureCanvasElementOverflow,
  measureCanvasTextBlocks,
  readCanvasRichClipboardFromDataTransfer,
  resetCanvasViewport,
  runCanvasKeyboardCommandIntent,
  runCanvasKeyboardToolIntent,
  runCanvasKeyboardViewportIntent,
  scheduleCanvasAnimationFrameTask,
  scheduleCanvasTimeoutTask,
  setCanvasDataTransferDropEffect,
  setCanvasDataTransferText,
  shouldReleaseCanvasKeyboardTemporaryPan,
  stringifyCanvasRichClipboardPayload,
  trapCanvasModalTabFocus,
  useCanvasMenuRovingFocus,
  useCanvasModalFocusLifecycle,
  useCanvasToolbarRovingFocus,
  writeCanvasClipboardText,
  writeCanvasRichClipboardPayload,
  zoomCanvasViewport,
  type CanvasFloatingAnchor,
  type CanvasDeferredSelectorFocusInput,
  type CanvasFocusableElement,
  type CanvasKeyboardToolIntent,
  type CanvasPastePositionMemory,
  type CanvasPointerClickMemory,
  type CanvasRichClipboardReadFormat,
  type CanvasRichClipboardWriteMode,
  type CanvasTabsDescriptor,
} from 'canvas/app'
import {
  CANVAS_COMMAND_PALETTE_ITEMS_MODEL,
  filterCanvasCommandPaletteItems,
  type CanvasCommandPaletteItem,
} from 'canvas/app/command-palette-items'
import {
  readCanvasClipboardImageSource,
} from 'canvas/app/image-clipboard'
import {
  CANVAS_IMAGE_IMPORT_MODEL,
  getCanvasDataImageSourceFromDataTransfer,
  getCanvasImageFileFromDataTransfer,
  getCanvasImageFileFromList,
  getCanvasSVGImageSourceFromDataTransfer,
  isCanvasImageBlob,
  readCanvasImageFileSource,
  resolveCanvasImageSourceNaturalSize,
  routeCanvasImagePasteReplace,
  type CanvasImageImportFormat,
  type CanvasImageImportSource,
  type CanvasImagePasteReplaceRoute,
  type CanvasImagePasteReplaceTarget,
} from 'canvas/app/image-import'
import {
  CANVAS_MEDIA_IMPORT_MODEL,
  getCanvasMediaSourceFromDataTransfer,
  getCanvasMediaSourceFromText,
  routeCanvasMediaSourceObjectHyperlink,
  type CanvasMediaObjectHyperlinkRoute,
  type CanvasMediaObjectHyperlinkTarget,
} from 'canvas/app/media-import'
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
  CANVAS_TABLE_IMPORT_MODEL,
  getCanvasTableColumnCount,
  getCanvasTableComponentSize,
  getCanvasTableFileFromDataTransfer,
  getCanvasTableFileFromList,
  getCanvasTableSourceFromDataTransfer,
  getCanvasTableSourceFromHTML,
  getCanvasTableSourceFromText,
  normalizeCanvasTableRows,
  readCanvasTableFileSource,
  type CanvasTableImportFormat,
  type CanvasTableImportSource,
} from 'canvas/app/table-import'
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
import {
  CANVAS_TEXT_PASTE_IMPORT_MODEL,
  createCanvasTextPasteItems,
  getCanvasRichTextPasteSourceFromDataTransfer,
  getCanvasTextPasteSourcesFromDataTransfer,
  type CanvasRichTextPasteSource,
} from 'canvas/app/text-paste-import'
import { useCanvasAppStageElement } from 'canvas/app/stage-element'
import { getNextCanvasDrawingPoints } from 'canvas/app/pointer-drawing'
import {
  CANVAS_LASER_TRAIL_OVERLAY_MODEL,
  previewCanvasPointerLaserInteraction,
  startCanvasPointerLaserInteraction,
  type CanvasPointerLaserInteraction,
} from 'canvas/app/pointer-laser'
import {
  previewCanvasPointerPanInteraction,
  startCanvasPointerPanInteraction,
  type CanvasPointerPanInteraction,
} from 'canvas/app/pointer-pan-interaction'
import { getCanvasPointerStartProjection } from 'canvas/app/pointer-start-session'
import {
  getCanvasEraserHitItemIds,
  getCanvasMergedEraserHitIds,
} from 'canvas/app/eraser-hit-testing'

export const PPT_COMMAND_PALETTE_ITEMS_MODEL =
  CANVAS_COMMAND_PALETTE_ITEMS_MODEL
export const PPT_CANVAS_IMAGE_IMPORT_MODEL = CANVAS_IMAGE_IMPORT_MODEL
export const PPT_CANVAS_MEDIA_IMPORT_MODEL = CANVAS_MEDIA_IMPORT_MODEL
export const PPT_MINIMAP_READ_MODEL = CANVAS_MINIMAP_READ_MODEL
export const PPT_PASTE_POSITION_MODEL = CANVAS_PASTE_POSITION_MODEL
export const PPT_CANVAS_TABLE_IMPORT_MODEL = CANVAS_TABLE_IMPORT_MODEL
export const PPT_CANVAS_TEXT_PASTE_IMPORT_MODEL =
  CANVAS_TEXT_PASTE_IMPORT_MODEL
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
export const PPT_LASER_TRAIL_OVERLAY_MODEL =
  CANVAS_LASER_TRAIL_OVERLAY_MODEL
export const PPT_COMMENT_THREAD_MODEL = CANVAS_COMMENT_THREAD_MODEL
export const PPT_KEYBOARD_COMMAND_DISPATCH_MODEL =
  CANVAS_KEYBOARD_COMMAND_DISPATCH_MODEL
export const PPT_KEYBOARD_COMMAND_INTENT_MODEL =
  CANVAS_KEYBOARD_COMMAND_INTENT_MODEL
export const PPT_KEYBOARD_NUDGE_INTENT_MODEL =
  CANVAS_KEYBOARD_NUDGE_INTENT_MODEL
export const PPT_KEYBOARD_NUDGE_KEYS = CANVAS_KEYBOARD_NUDGE_KEYS
export const PPT_KEYBOARD_NUDGE_LARGE_STEP =
  CANVAS_KEYBOARD_NUDGE_LARGE_STEP
export const PPT_KEYBOARD_NUDGE_MODEL = CANVAS_KEYBOARD_NUDGE_MODEL
export const PPT_KEYBOARD_NUDGE_STEP = CANVAS_KEYBOARD_NUDGE_STEP
export const PPT_KEYBOARD_TOOL_DISPATCH_MODEL =
  CANVAS_KEYBOARD_TOOL_DISPATCH_MODEL
export const PPT_KEYBOARD_VIEWPORT_INTENT_MODEL =
  CANVAS_KEYBOARD_VIEWPORT_INTENT_MODEL
export const PPT_KEYBOARD_VIEWPORT_MODEL =
  CANVAS_KEYBOARD_VIEWPORT_MODEL
export const PPT_MENU_FOCUS_MODEL = CANVAS_MENU_FOCUS_MODEL
export const PPT_MENU_ITEM_PROPS = CANVAS_MENU_ITEM_PROPS
export const PPT_MENU_KEYBOARD_KEYS = CANVAS_MENU_KEYBOARD_KEYS
export const PPT_MENU_ROVING_FOCUS_MODEL =
  CANVAS_MENU_ROVING_FOCUS_MODEL
export const PPT_MODAL_FOCUS_LIFECYCLE_MODEL =
  CANVAS_MODAL_FOCUS_LIFECYCLE_MODEL
export const PPT_RADIO_GROUP_FOCUS_MODEL =
  CANVAS_RADIO_GROUP_FOCUS_MODEL
export const PPT_RADIO_GROUP_KEYBOARD_MODEL =
  CANVAS_RADIO_GROUP_KEYBOARD_MODEL
export const PPT_RADIO_GROUP_MODEL = CANVAS_RADIO_GROUP_MODEL
export const PPT_RESIZE_POINTER_MODIFIERS_MODEL =
  CANVAS_RESIZE_POINTER_MODIFIERS_MODEL
export const PPT_SELECTION_TOOLBAR_DROPDOWN_MENU_MODEL =
  CANVAS_SELECTION_TOOLBAR_DROPDOWN_MENU_MODEL
export const PPT_TABS_ROVING_FOCUS_MODEL =
  CANVAS_TABS_ROVING_FOCUS_MODEL
export const PPT_TOOLBAR_FOCUS_MODEL = CANVAS_TOOLBAR_FOCUS_MODEL
export const PPT_TOOLBAR_ITEM_PROPS = CANVAS_TOOLBAR_ITEM_PROPS
export const PPT_TOOLBAR_KEYBOARD_MODEL =
  CANVAS_TOOLBAR_KEYBOARD_MODEL
export const PPT_TOOLBAR_ROVING_FOCUS_MODEL =
  CANVAS_TOOLBAR_ROVING_FOCUS_MODEL

export const bindPPTCanvasEventListener = bindCanvasEventListener
export const bindPPTCanvasEventListeners = bindCanvasEventListeners
export const cancelPPTCanvasAnimationFrameTask =
  cancelCanvasAnimationFrameTask
export const cancelPPTCanvasDeferredFocus = cancelCanvasDeferredFocus
export const capturePPTCanvasPointerFromEvent =
  captureCanvasPointerFromEvent
export const centerPPTCanvasViewportAtWorldPoint =
  centerCanvasViewportAtWorldPoint
export const createPPTCanvasDataTransferImportActionPlan =
  createCanvasDataTransferImportActionPlan
export const createPPTCanvasPastePositionKey =
  createCanvasPastePositionKey
export const createPPTCanvasRichClipboardHTML =
  createCanvasRichClipboardHTML
export const createPPTCanvasTabsDescriptor = createCanvasTabsDescriptor
export const downloadPPTCanvasTextFile = downloadCanvasTextFile
export const fitPPTCanvasViewportToBounds = fitCanvasViewportToBounds
export const focusPPTCanvasElement = focusCanvasElement
export function focusPPTCanvasElementBySelectorOnNextFrame<
  TElement extends Element & CanvasFocusableElement = HTMLElement,
>(input: CanvasDeferredSelectorFocusInput<TElement>) {
  return focusCanvasElementBySelectorOnNextFrame<TElement>(input)
}
export const focusPPTCanvasElementOnNextFrame =
  focusCanvasElementOnNextFrame
export const getPPTCanvasClientViewportSize =
  getCanvasClientViewportSize
export const getPPTCanvasCommandPaletteKeyboardIntent =
  getCanvasCommandPaletteKeyboardIntent
export const getPPTCanvasContextMenuKeyboardIntent =
  getCanvasContextMenuKeyboardIntent
export const getPPTCanvasContextMenuPosition =
  getCanvasContextMenuPosition
export const getPPTCanvasDataTransferText = getCanvasDataTransferText
export const getPPTCanvasEditableFieldKeyboardIntent =
  getCanvasEditableFieldKeyboardIntent
export const getPPTCanvasFindInputKeyboardIntent =
  getCanvasFindInputKeyboardIntent
export const getPPTCanvasFloatingAnchorForBounds =
  getCanvasFloatingAnchorForBounds
export const getPPTCanvasImportedImageSize =
  getCanvasImportedImageSize
export const getPPTCanvasInlineEditKeyboardIntent =
  getCanvasInlineEditKeyboardIntent
export const getPPTCanvasKeyboardBuiltinCommandShortcutIntent =
  getCanvasKeyboardBuiltinCommandShortcutIntent
export const getPPTCanvasKeyboardNudgeShortcutIntent =
  getCanvasKeyboardNudgeShortcutIntent
export const getPPTCanvasKeyboardSystemShortcutIntent =
  getCanvasKeyboardSystemShortcutIntent
export const getPPTCanvasKeyboardToolShortcutIntent =
  getCanvasKeyboardToolShortcutIntent
export const getPPTCanvasKeyboardViewportShortcutIntent =
  getCanvasKeyboardViewportShortcutIntent
export const getPPTCanvasMenuTriggerKeyboardIntent =
  getCanvasMenuTriggerKeyboardIntent
export const getPPTCanvasModalBackdropPointerIntent =
  getCanvasModalBackdropPointerIntent
export const getPPTCanvasModalKeyboardIntent =
  getCanvasModalKeyboardIntent
export const getPPTCanvasPasteOffsetForBounds =
  getCanvasPasteOffsetForBounds
export const getPPTCanvasPastePositionSession =
  getCanvasPastePositionSession
export const getPPTCanvasPointerLocalGeometry =
  getCanvasPointerLocalGeometry
export const getPPTCanvasPointerTransformModifierState =
  getCanvasPointerTransformModifierState
export const getPPTCanvasPresentationKeyboardIntent =
  getCanvasPresentationKeyboardIntent
export const getPPTCanvasRadioTabIndex = getCanvasRadioTabIndex
export const getPPTCanvasResizeHandleDoubleClickIntent =
  getCanvasResizeHandleDoubleClickIntent
export const getPPTCanvasSelectionListModifierState =
  getCanvasSelectionListModifierState
export const getPPTCanvasTabsKeyboardIntent =
  getCanvasTabsKeyboardIntent
export const getPPTCanvasWorldClientPoint = getCanvasWorldClientPoint
export const handlePPTCanvasRadioGroupKeyDown =
  handleCanvasRadioGroupKeyDown
export const isPPTCanvasControlTarget = isCanvasControlTarget
export const isPPTCanvasKeyboardCommandIntent =
  isCanvasKeyboardCommandIntent
export const isPPTCanvasKeyboardToolIntent =
  isCanvasKeyboardToolIntent
export const isPPTCanvasKeyboardTypingTarget =
  isCanvasKeyboardTypingTarget
export const isPPTCanvasKeyboardViewportIntent =
  isCanvasKeyboardViewportIntent
export const isPPTCanvasTargetWithinSelector =
  isCanvasTargetWithinSelector
export const isPPTCanvasWheelPassthroughTarget =
  isCanvasWheelPassthroughTarget
export const measurePPTCanvasElementOverflow =
  measureCanvasElementOverflow
export const measurePPTCanvasTextBlocks = measureCanvasTextBlocks
export const readPPTCanvasRichClipboardFromDataTransfer =
  readCanvasRichClipboardFromDataTransfer
export const resetPPTCanvasViewport = resetCanvasViewport
export const runPPTCanvasKeyboardCommandIntent =
  runCanvasKeyboardCommandIntent
export const runPPTCanvasKeyboardToolIntent =
  runCanvasKeyboardToolIntent
export const runPPTCanvasKeyboardViewportIntent =
  runCanvasKeyboardViewportIntent
export const schedulePPTCanvasAnimationFrameTask =
  scheduleCanvasAnimationFrameTask
export const schedulePPTCanvasTimeoutTask = scheduleCanvasTimeoutTask
export const setPPTCanvasDataTransferDropEffect =
  setCanvasDataTransferDropEffect
export const setPPTCanvasDataTransferText = setCanvasDataTransferText
export const shouldReleasePPTCanvasKeyboardTemporaryPan =
  shouldReleaseCanvasKeyboardTemporaryPan
export const stringifyPPTCanvasRichClipboardPayload =
  stringifyCanvasRichClipboardPayload
export const trapPPTCanvasModalTabFocus = trapCanvasModalTabFocus
export const usePPTCanvasMenuRovingFocus = useCanvasMenuRovingFocus
export const usePPTCanvasModalFocusLifecycle =
  useCanvasModalFocusLifecycle
export const usePPTCanvasToolbarRovingFocus =
  useCanvasToolbarRovingFocus
export const writePPTCanvasClipboardText = writeCanvasClipboardText
export const writePPTCanvasRichClipboardPayload =
  writeCanvasRichClipboardPayload
export const zoomPPTCanvasViewport = zoomCanvasViewport

export const filterPPTCommandPaletteItems =
  filterCanvasCommandPaletteItems
export const createPPTCanvasTextPasteItems = createCanvasTextPasteItems
export const getPPTCanvasDataImageSourceFromDataTransfer =
  getCanvasDataImageSourceFromDataTransfer
export const readPPTCanvasClipboardImageSource =
  readCanvasClipboardImageSource
export const getPPTCanvasImageFileFromDataTransfer =
  getCanvasImageFileFromDataTransfer
export const getPPTCanvasImageFileFromList = getCanvasImageFileFromList
export function getPPTCanvasDataTransferFiles(
  dataTransfer: DataTransfer | null,
) {
  const fileList = Array.from(dataTransfer?.files ?? [])
  const types = Array.from(dataTransfer?.types ?? [])
  const shouldReadItemFiles = fileList.length === 0 &&
    (types.length === 0 || types.includes('Files'))
  const files = [
    ...fileList,
    ...(shouldReadItemFiles
      ? Array.from(dataTransfer?.items ?? [])
        .map((item) => item.kind === 'file' ? item.getAsFile() : null)
        .filter((file): file is File => file !== null)
      : []),
  ]
  const seen = new Set<string>()

  return files.filter((file) => {
    const key = [
      file.name,
      file.type,
      file.size,
      file.lastModified,
    ].join(':')

    if (seen.has(key)) {
      return false
    }

    seen.add(key)
    return true
  })
}
export function getPPTCanvasImageFilesFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  return getPPTCanvasDataTransferFiles(dataTransfer).filter(isCanvasImageBlob)
}
export const routePPTCanvasImagePasteReplace =
  routeCanvasImagePasteReplace
export const getPPTCanvasMediaSourceFromDataTransfer =
  getCanvasMediaSourceFromDataTransfer
export const getPPTCanvasMediaSourceFromText =
  getCanvasMediaSourceFromText
export const routePPTCanvasMediaSourceObjectHyperlink =
  routeCanvasMediaSourceObjectHyperlink
export const getPPTMinimapPointFromViewportOffset =
  getCanvasMinimapPointFromViewportOffset
export const getPPTMinimapReadModel = getCanvasMinimapReadModel
export const getPPTMinimapWorldPoint = getCanvasMinimapWorldPoint
export const getNextPPTCanvasDrawingPoints = getNextCanvasDrawingPoints
export const getPPTCanvasEraserHitItemIds = getCanvasEraserHitItemIds
export const getPPTCanvasMergedEraserHitIds =
  getCanvasMergedEraserHitIds
export const getPPTCanvasPointerStartProjection =
  getCanvasPointerStartProjection
export const getPPTCanvasRichTextPasteSourceFromDataTransfer =
  getCanvasRichTextPasteSourceFromDataTransfer
export const getPPTCanvasSVGImageSourceFromDataTransfer =
  getCanvasSVGImageSourceFromDataTransfer
export const getPPTCanvasTableColumnCount = getCanvasTableColumnCount
export const getPPTCanvasTableComponentSize =
  getCanvasTableComponentSize
export const getPPTCanvasTableFileFromDataTransfer =
  getCanvasTableFileFromDataTransfer
export const getPPTCanvasTableFileFromList = getCanvasTableFileFromList
export const getPPTCanvasTableSourceFromDataTransfer =
  getCanvasTableSourceFromDataTransfer
export const getPPTCanvasTableSourceFromHTML =
  getCanvasTableSourceFromHTML
export const getPPTCanvasTableSourceFromText =
  getCanvasTableSourceFromText
export const getPPTCanvasTextPasteSourcesFromDataTransfer =
  getCanvasTextPasteSourcesFromDataTransfer
export const getPPTInlineEditHistoryDirectionFromInputType =
  inlineEditHistoryDirectionFromInputType
export const insertPPTInlineEditText = insertInlineEditText
export const isPPTInlineEditLineBreakInput = isInlineEditLineBreakInput
export const normalizePPTCanvasTableRows = normalizeCanvasTableRows
export const previewPPTCanvasPointerLaserInteraction =
  previewCanvasPointerLaserInteraction
export const previewPPTCanvasPointerPanInteraction =
  previewCanvasPointerPanInteraction
export const startPPTCanvasPointerLaserInteraction =
  startCanvasPointerLaserInteraction
export const startPPTCanvasPointerPanInteraction =
  startCanvasPointerPanInteraction
export const readPPTCanvasImageFileSource = readCanvasImageFileSource
export const readPPTCanvasTableFileSource = readCanvasTableFileSource
export const resolvePPTCanvasImageSourceNaturalSize =
  resolveCanvasImageSourceNaturalSize
export const usePPTCanvasAppStageElement = useCanvasAppStageElement

export type PPTCommandPaletteItemBase = CanvasCommandPaletteItem
export type PPTCanvasImageImportFormat = CanvasImageImportFormat
export type PPTCanvasImageImportSource = CanvasImageImportSource
export type PPTCanvasImagePasteReplaceRoute =
  CanvasImagePasteReplaceRoute
export type PPTCanvasImagePasteReplaceTarget =
  CanvasImagePasteReplaceTarget
export type PPTCanvasMediaObjectHyperlinkRoute =
  CanvasMediaObjectHyperlinkRoute
export type PPTCanvasMediaObjectHyperlinkTarget =
  CanvasMediaObjectHyperlinkTarget
export type PPTCanvasRichTextPasteSource = CanvasRichTextPasteSource
export type PPTCanvasTableImportFormat = CanvasTableImportFormat
export type PPTCanvasTableImportSource = CanvasTableImportSource
export type PPTMinimapItemBounds = CanvasMinimapItemBounds
export type PPTMinimapReadModel = CanvasMinimapReadModel
export type PPTMinimapSize = CanvasMinimapSize
export type PPTCanvasEraserHitInput =
  Parameters<typeof getCanvasEraserHitItemIds>[0]
export type PPTCanvasPointerLaserInteractionBase =
  CanvasPointerLaserInteraction
export type PPTCanvasPointerPanInteractionBase =
  CanvasPointerPanInteraction
export type PPTCanvasFloatingAnchor = CanvasFloatingAnchor
export type PPTCanvasKeyboardToolIntent = CanvasKeyboardToolIntent
export type PPTCanvasPastePositionMemory = CanvasPastePositionMemory
export type PPTCanvasPointerClickMemory = CanvasPointerClickMemory
export type PPTCanvasRichClipboardReadFormat =
  CanvasRichClipboardReadFormat
export type PPTCanvasRichClipboardWriteMode =
  CanvasRichClipboardWriteMode
export type PPTCanvasTabsDescriptor<TId extends string> =
  CanvasTabsDescriptor<TId>
