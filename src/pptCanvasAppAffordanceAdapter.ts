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
  CANVAS_LASER_TRAIL_OVERLAY_MODEL,
  CANVAS_MENU_FOCUS_MODEL,
  CANVAS_MENU_ITEM_PROPS,
  CANVAS_MENU_KEYBOARD_KEYS,
  CANVAS_MENU_ROVING_FOCUS_MODEL,
  CANVAS_MODAL_FOCUS_LIFECYCLE_MODEL,
  CANVAS_RADIO_GROUP_FOCUS_MODEL,
  CANVAS_RADIO_GROUP_KEYBOARD_MODEL,
  CANVAS_RADIO_GROUP_MODEL,
  CANVAS_RESIZE_POINTER_MODIFIERS_MODEL,
  CANVAS_RICH_TEXT_PASTE_SUPPORTED_FORMATS,
  CANVAS_SELECTION_TOOLBAR_DROPDOWN_MENU_MODEL,
  CANVAS_TABS_ROVING_FOCUS_MODEL,
  CANVAS_TABLE_FILE_IMPORT_SUPPORTED_FORMATS,
  CANVAS_TABLE_IMPORT_MODEL,
  CANVAS_TABLE_IMPORT_SUPPORTED_FORMATS,
  CANVAS_TEXT_PASTE_SUPPORTED_FORMATS,
  CANVAS_TEXT_PASTE_IMPORT_MODEL,
  CANVAS_TOOLBAR_FOCUS_MODEL,
  CANVAS_TOOLBAR_ITEM_PROPS,
  CANVAS_TOOLBAR_KEYBOARD_MODEL,
  CANVAS_TOOLBAR_ROVING_FOCUS_MODEL,
  CANVAS_WHEEL_VIEWPORT_HORIZONTAL_PAN_MODIFIER,
  CANVAS_WHEEL_VIEWPORT_MODEL,
  CANVAS_WHEEL_VIEWPORT_PAN_MODE,
  CANVAS_WHEEL_VIEWPORT_ZOOM_MODIFIER,
  bindCanvasEventListener,
  bindCanvasEventListeners,
  cancelCanvasAnimationFrameTask,
  cancelCanvasDeferredFocus,
  captureCanvasPointerFromEvent,
  centerCanvasViewportAtWorldPoint,
  createCanvasDataTransferImportActionPlan,
  createCanvasDataTransferImportActionPlanFromRegistry,
  createCanvasDataTransferImportRegistry,
  CANVAS_DATA_TRANSFER_TEXT_MIME_TYPE,
  createCanvasExternalClipboardImagePasteActionResolver,
  createCanvasExternalClipboardPasteActionPlan,
  createCanvasClipboardCommandEffectPlan,
  createCanvasPastePositionKey,
  createCanvasRichClipboardHTML,
  createCanvasTabsDescriptor,
  copyCanvasClipboardSelection,
  cutCanvasClipboardSelection,
  duplicateCanvasClipboardSelection,
  createCanvasTextPasteItems,
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
  getCanvasDataTransferImportRegistryMetadata,
  getCanvasEditableFieldKeyboardIntent,
  getCanvasEraserHitItemIds,
  getCanvasExternalClipboardPasteCommandRoute,
  getCanvasFindInputKeyboardIntent,
  getCanvasFloatingAnchorForBounds,
  getCanvasImportedImageSize,
  getCanvasImageInsertCenter,
  getCanvasInlineEditKeyboardIntent,
  getCanvasKeyboardBuiltinCommandShortcutIntent,
  getCanvasKeyboardNudgeShortcutIntent,
  getCanvasKeyboardSystemShortcutIntent,
  getCanvasKeyboardToolShortcutIntent,
  getCanvasKeyboardViewportShortcutIntent,
  getCanvasMergedEraserHitIds,
  getCanvasMediaInsertPosition,
  getCanvasMenuTriggerKeyboardIntent,
  getCanvasModalBackdropPointerIntent,
  getCanvasModalKeyboardIntent,
  getNextCanvasDrawingPoints,
  getCanvasPasteOffsetForBounds,
  getCanvasPastePositionSession,
  getCanvasPointerLocalGeometry,
  getCanvasPointerTransformModifierState,
  getCanvasPresentationKeyboardIntent,
  getCanvasRadioTabIndex,
  getCanvasRichClipboardJSONFromHTML,
  getCanvasResizeHandleDoubleClickIntent,
  getCanvasRichTextPasteSourceFromDataTransfer,
  getCanvasSelectionListModifierState,
  getCanvasTableColumnCount,
  getCanvasTableComponentSize,
  getCanvasTableFileFromDataTransfer,
  getCanvasTableFileFromList,
  getCanvasTableFilesFromDataTransfer,
  getCanvasTableFilesFromList,
  getCanvasTableInsertCenter,
  getCanvasTableSourceFromDataTransfer,
  getCanvasTableSourceFromHTML,
  getCanvasTableSourceFromText,
  getCanvasTabsKeyboardIntent,
  getCanvasTextPasteInsertPosition,
  getCanvasTextPasteSourceCandidatesFromDataTransfer,
  getCanvasTextPasteSourceText,
  getCanvasTextPasteSourcesFromDataTransfer,
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
  readCanvasDataTransferJSONCandidate,
  readCanvasRichClipboardFromDataTransfer,
  readCanvasTableFileSource,
  readCanvasTableFileSources,
  resetCanvasViewport,
  routeCanvasTableImportTargetReplace,
  routeCanvasTextPasteReplace,
  runCanvasDataTransferImportActionPlan,
  runCanvasKeyboardCommandIntent,
  runCanvasKeyboardToolIntent,
  runCanvasKeyboardViewportIntent,
  runCanvasWheelViewport,
  scheduleCanvasAnimationFrameTask,
  scheduleCanvasTimeoutTask,
  setCanvasDataTransferDropEffect,
  setCanvasDataTransferText,
  shouldReleaseCanvasKeyboardTemporaryPan,
  stringifyCanvasRichClipboardPayload,
  normalizeCanvasTableRows,
  applyCanvasClipboardCommandEffect,
  applyCanvasStandardDocumentEffect,
  commitCanvasAppHostItemsChange,
  executeCanvasClipboardCommand,
  executeCanvasStandardCommand,
  pasteCanvasClipboardSelection,
  trapCanvasModalTabFocus,
  transformCanvasAppItemsChange,
  useCanvasMenuRovingFocus,
  useCanvasModalFocusLifecycle,
  useCanvasAppStageElement,
  useCanvasToolbarRovingFocus,
  writeCanvasClipboardText,
  writeCanvasRichClipboardPayload,
  zoomCanvasViewport,
  type CanvasFloatingAnchor,
  type CanvasClipboardCommand,
  type CanvasClipboardCommandEffect,
  type CanvasClipboardCommandEffectContext,
  type CanvasClipboardCommandEffectPlanContext,
  type CanvasClipboardCommandExecutionContext,
  type CanvasClipboardCommandExecutionResult,
  type CanvasAppItemsChange,
  type CanvasAppItemsChangeTransformer,
  type CanvasAppHostItemsChangeCommitResult,
  type CommitCanvasAppHostItemsChangeArgs,
  type CanvasStandardCommand,
  type CanvasStandardCommandDocumentEffect,
  type CanvasStandardCommandDocumentEffectContext,
  type CanvasStandardCommandEffectPlanContext,
  type CanvasStandardCommandExecutionContext,
  type CanvasStandardCommandItemsChange,
  type CanvasDeferredSelectorFocusInput,
  type CanvasDataTransferImportRegistry,
  type CanvasDataTransferImportRegistryResolver,
  type CanvasDataTransferImportRegistryResolverMetadata,
  type CanvasFocusableElement,
  type CanvasKeyboardToolIntent,
  type CanvasPastePositionMemory,
  type CanvasPointerClickMemory,
  previewCanvasPointerLaserInteraction,
  previewCanvasPointerPanInteraction,
  startCanvasPointerLaserInteraction,
  type CanvasPointerLaserInteraction,
  startCanvasPointerPanInteraction,
  type CanvasPointerPanInteraction,
  type CanvasRichClipboardReadFormat,
  type CanvasRichClipboardWriteMode,
  type CanvasRichTextPasteSource,
  type RunCanvasClipboardCommand,
  type CanvasTabsDescriptor,
  type CanvasTableImportFormat,
  type CanvasTableImportSource,
  type CanvasTableImportTargetReplaceRoute,
  type CanvasTableImportTargetReplaceTarget,
  type CanvasTextPasteReplaceRoute,
  type CanvasTextPasteReplaceTarget,
  type CanvasTextPasteSource,
} from 'canvas/app'
import {
  CANVAS_COMPONENT_DEFINITION_REGISTRY,
} from 'canvas/host'
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
  getCanvasHTMLDataImageSourcesFromDataTransfer,
  getCanvasImageFileFromDataTransfer,
  getCanvasImageFileFromList,
  getCanvasImageFilesFromDataTransfer,
  getCanvasImageSourceFromDataTransfer,
  getCanvasSVGImageSourceFromDataTransfer,
  readCanvasImageFileSource,
  readCanvasImageFileSources,
  resolveCanvasImageSourceNaturalSize,
  routeCanvasImagePasteReplace,
  type CanvasImageImportFormat,
  type CanvasImageImportSource,
  type CanvasImagePasteReplaceRoute,
  type CanvasImagePasteReplaceTarget,
} from 'canvas/app/image-import'
import {
  CANVAS_MEDIA_IMPORT_MODEL,
  CANVAS_MEDIA_SOURCE_JSON_MIME_TYPE,
  CANVAS_MEDIA_SOURCE_JSON_TYPES,
  getCanvasMediaSourceFromDataTransfer,
  getCanvasMediaSourceFromJSONDataTransfer,
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
  CANVAS_KEYBOARD_TEMPORARY_PAN_MODEL,
  CANVAS_KEYBOARD_TEMPORARY_PAN_SHORTCUT_LABEL,
} from 'canvas/app/keyboard-system-shortcuts'
import {
  CANVAS_INLINE_EDIT_DOM_MODEL,
  inlineEditHistoryDirectionFromInputType,
  insertInlineEditText,
  isInlineEditLineBreakInput,
} from 'canvas/app/inline-edit-dom'
import {
  CANVAS_POINTER_CLICK_MEMORY_MODEL,
} from 'canvas/app/pointer-click-memory'
import { getCanvasPointerStartProjection } from 'canvas/app/pointer-start-session'
export const PPT_COMMAND_PALETTE_ITEMS_MODEL =
  CANVAS_COMMAND_PALETTE_ITEMS_MODEL
export const PPT_CANVAS_IMAGE_IMPORT_MODEL = CANVAS_IMAGE_IMPORT_MODEL
export const PPT_CANVAS_MEDIA_IMPORT_MODEL = CANVAS_MEDIA_IMPORT_MODEL
export const PPT_MINIMAP_READ_MODEL = CANVAS_MINIMAP_READ_MODEL
export const PPT_PASTE_POSITION_MODEL = CANVAS_PASTE_POSITION_MODEL
export const PPT_CANVAS_TABLE_IMPORT_MODEL = CANVAS_TABLE_IMPORT_MODEL
export const PPT_CANVAS_TEXT_PASTE_IMPORT_MODEL =
  CANVAS_TEXT_PASTE_IMPORT_MODEL
export const PPT_CANVAS_TABLE_FILE_IMPORT_SUPPORTED_FORMATS =
  CANVAS_TABLE_FILE_IMPORT_SUPPORTED_FORMATS
export const PPT_CANVAS_TABLE_IMPORT_SUPPORTED_FORMATS =
  CANVAS_TABLE_IMPORT_SUPPORTED_FORMATS
export const PPT_CANVAS_RICH_TEXT_PASTE_SUPPORTED_FORMATS =
  CANVAS_RICH_TEXT_PASTE_SUPPORTED_FORMATS
export const PPT_CANVAS_TEXT_PASTE_SUPPORTED_FORMATS =
  CANVAS_TEXT_PASTE_SUPPORTED_FORMATS
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
export const applyPPTCanvasClipboardCommandEffect =
  applyCanvasClipboardCommandEffect
export const applyPPTCanvasStandardDocumentEffect =
  applyCanvasStandardDocumentEffect
export const commitPPTCanvasAppHostItemsChange =
  commitCanvasAppHostItemsChange
export const createPPTCanvasClipboardCommandEffectPlan =
  createCanvasClipboardCommandEffectPlan
export const createPPTCanvasDataTransferImportActionPlan =
  createCanvasDataTransferImportActionPlan
export const createPPTCanvasDataTransferImportActionPlanFromRegistry =
  createCanvasDataTransferImportActionPlanFromRegistry
export const createPPTCanvasDataTransferImportRegistry =
  createCanvasDataTransferImportRegistry
export const createPPTCanvasExternalClipboardImagePasteActionResolver =
  createCanvasExternalClipboardImagePasteActionResolver
export const createPPTCanvasExternalClipboardPasteActionPlan =
  createCanvasExternalClipboardPasteActionPlan
export const createPPTCanvasPastePositionKey =
  createCanvasPastePositionKey
export const createPPTCanvasRichClipboardHTML =
  createCanvasRichClipboardHTML
export const getPPTCanvasRichClipboardJSONFromHTML =
  getCanvasRichClipboardJSONFromHTML
export const createPPTCanvasTabsDescriptor = createCanvasTabsDescriptor
export const copyPPTCanvasClipboardSelection =
  copyCanvasClipboardSelection
export const cutPPTCanvasClipboardSelection =
  cutCanvasClipboardSelection
export const duplicatePPTCanvasClipboardSelection =
  duplicateCanvasClipboardSelection
export const downloadPPTCanvasTextFile = downloadCanvasTextFile
export const executePPTCanvasClipboardCommand =
  executeCanvasClipboardCommand
export const executePPTCanvasStandardCommand =
  executeCanvasStandardCommand
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
export const PPT_CANVAS_DATA_TRANSFER_TEXT_MIME_TYPE =
  CANVAS_DATA_TRANSFER_TEXT_MIME_TYPE
export const getPPTCanvasDataTransferImportRegistryMetadata =
  getCanvasDataTransferImportRegistryMetadata
export const getPPTCanvasEditableFieldKeyboardIntent =
  getCanvasEditableFieldKeyboardIntent
export const getPPTCanvasExternalClipboardPasteCommandRoute =
  getCanvasExternalClipboardPasteCommandRoute
export const getPPTCanvasFindInputKeyboardIntent =
  getCanvasFindInputKeyboardIntent
export const getPPTCanvasFloatingAnchorForBounds =
  getCanvasFloatingAnchorForBounds
export const getPPTCanvasImportedImageSize =
  getCanvasImportedImageSize
export const getPPTCanvasImageInsertCenter =
  getCanvasImageInsertCenter
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
export const readPPTCanvasDataTransferJSONCandidate =
  readCanvasDataTransferJSONCandidate
export const readPPTCanvasRichClipboardFromDataTransfer =
  readCanvasRichClipboardFromDataTransfer
export const resetPPTCanvasViewport = resetCanvasViewport
export const runPPTCanvasDataTransferImportActionPlan =
  runCanvasDataTransferImportActionPlan
export const runPPTCanvasKeyboardCommandIntent =
  runCanvasKeyboardCommandIntent
export const runPPTCanvasKeyboardToolIntent =
  runCanvasKeyboardToolIntent
export const runPPTCanvasKeyboardViewportIntent =
  runCanvasKeyboardViewportIntent
export const runPPTCanvasWheelViewport = runCanvasWheelViewport
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
export const transformPPTCanvasAppItemsChange =
  transformCanvasAppItemsChange
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
export const getPPTCanvasHTMLDataImageSourcesFromDataTransfer =
  getCanvasHTMLDataImageSourcesFromDataTransfer
export const readPPTCanvasClipboardImageSource =
  readCanvasClipboardImageSource
export const getPPTCanvasImageFileFromDataTransfer =
  getCanvasImageFileFromDataTransfer
export const getPPTCanvasImageFileFromList = getCanvasImageFileFromList
export const getPPTCanvasImageFilesFromDataTransfer =
  getCanvasImageFilesFromDataTransfer
export const getPPTCanvasImageSourceFromDataTransfer =
  getCanvasImageSourceFromDataTransfer
export const routePPTCanvasImagePasteReplace =
  routeCanvasImagePasteReplace
export const getPPTCanvasMediaSourceFromDataTransfer =
  getCanvasMediaSourceFromDataTransfer
export const getPPTCanvasMediaSourceFromJSONDataTransfer =
  getCanvasMediaSourceFromJSONDataTransfer
export const getPPTCanvasMediaSourceFromText =
  getCanvasMediaSourceFromText
export const PPT_CANVAS_MEDIA_SOURCE_JSON_MIME_TYPE =
  CANVAS_MEDIA_SOURCE_JSON_MIME_TYPE
export const PPT_CANVAS_MEDIA_SOURCE_JSON_TYPES =
  CANVAS_MEDIA_SOURCE_JSON_TYPES
export const getPPTCanvasMediaInsertPosition =
  getCanvasMediaInsertPosition
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
export const getPPTCanvasTableFilesFromDataTransfer =
  getCanvasTableFilesFromDataTransfer
export const getPPTCanvasTableFilesFromList =
  getCanvasTableFilesFromList
export const getPPTCanvasTableInsertCenter =
  getCanvasTableInsertCenter
export const getPPTCanvasTableSourceFromDataTransfer =
  getCanvasTableSourceFromDataTransfer
export const getPPTCanvasTableSourceFromHTML =
  getCanvasTableSourceFromHTML
export const getPPTCanvasTableSourceFromText =
  getCanvasTableSourceFromText
export const routePPTCanvasTableImportTargetReplace =
  routeCanvasTableImportTargetReplace
export const getPPTCanvasTextPasteSourcesFromDataTransfer =
  getCanvasTextPasteSourcesFromDataTransfer
export const getPPTCanvasTextPasteSourceCandidatesFromDataTransfer =
  getCanvasTextPasteSourceCandidatesFromDataTransfer
export const getPPTCanvasTextPasteSourceText = getCanvasTextPasteSourceText
export const getPPTCanvasTextPasteInsertPosition =
  getCanvasTextPasteInsertPosition
export const routePPTCanvasTextPasteReplace =
  routeCanvasTextPasteReplace
export const getPPTInlineEditHistoryDirectionFromInputType =
  inlineEditHistoryDirectionFromInputType
export const insertPPTInlineEditText = insertInlineEditText
export const isPPTInlineEditLineBreakInput = isInlineEditLineBreakInput
export const normalizePPTCanvasTableRows = normalizeCanvasTableRows
export const pastePPTCanvasClipboardSelection =
  pasteCanvasClipboardSelection
export const previewPPTCanvasPointerLaserInteraction =
  previewCanvasPointerLaserInteraction
export const previewPPTCanvasPointerPanInteraction =
  previewCanvasPointerPanInteraction
export const startPPTCanvasPointerLaserInteraction =
  startCanvasPointerLaserInteraction
export const startPPTCanvasPointerPanInteraction =
  startCanvasPointerPanInteraction
export const readPPTCanvasImageFileSource = readCanvasImageFileSource
export const readPPTCanvasImageFileSources = readCanvasImageFileSources
export const readPPTCanvasTableFileSource = readCanvasTableFileSource
export const readPPTCanvasTableFileSources = readCanvasTableFileSources
export const resolvePPTCanvasImageSourceNaturalSize =
  resolveCanvasImageSourceNaturalSize
export const usePPTCanvasAppStageElement = useCanvasAppStageElement

export type PPTCommandPaletteItemBase = CanvasCommandPaletteItem
export type PPTCanvasClipboardCommand = CanvasClipboardCommand
export type PPTCanvasClipboardCommandEffect<TItem extends { id: string }> =
  CanvasClipboardCommandEffect<TItem>
export type PPTCanvasClipboardCommandEffectContext<
  TItem extends { id: string },
> = CanvasClipboardCommandEffectContext<TItem>
export type PPTCanvasClipboardCommandEffectPlanContext<
  TItem extends { id: string },
> = CanvasClipboardCommandEffectPlanContext<TItem>
export type PPTCanvasClipboardCommandExecutionContext<
  TItem extends { id: string },
> = CanvasClipboardCommandExecutionContext<TItem>
export type PPTCanvasClipboardCommandExecutionResult<
  TItem extends { id: string },
> = CanvasClipboardCommandExecutionResult<TItem>
export type PPTCanvasAppItemsChange<TItem extends { id: string }> =
  CanvasAppItemsChange<TItem>
export type PPTCanvasAppItemsChangeTransformer<
  TItem extends { id: string },
> = CanvasAppItemsChangeTransformer<TItem>
export type PPTCanvasAppHostItemsChangeCommitResult<
  TItem extends { id: string },
> = CanvasAppHostItemsChangeCommitResult<TItem>
export type CommitPPTCanvasAppHostItemsChangeArgs<
  TItem extends { id: string },
> = CommitCanvasAppHostItemsChangeArgs<TItem>
export type PPTCanvasStandardCommand = CanvasStandardCommand
export type PPTCanvasStandardCommandDocumentEffect<
  TItem extends { id: string },
> = CanvasStandardCommandDocumentEffect<TItem>
export type PPTCanvasStandardCommandDocumentEffectContext<
  TItem extends { id: string },
> = CanvasStandardCommandDocumentEffectContext<TItem>
export type PPTCanvasStandardCommandEffectPlanContext<
  TItem extends { id: string },
> = CanvasStandardCommandEffectPlanContext<TItem>
export type PPTCanvasStandardCommandExecutionContext<
  TItem extends { id: string },
> = CanvasStandardCommandExecutionContext<TItem>
export type PPTCanvasStandardCommandItemsChange<
  TItem extends { id: string },
> = CanvasStandardCommandItemsChange<TItem>
export type PPTCanvasDataTransferImportRegistry<
  TAction,
  TScope extends string,
> = CanvasDataTransferImportRegistry<TAction, TScope>
export type PPTCanvasDataTransferImportRegistryResolver<
  TAction,
  TScope extends string,
> = CanvasDataTransferImportRegistryResolver<TAction, TScope>
export type PPTCanvasDataTransferImportRegistryResolverMetadata<
  TScope extends string,
> = CanvasDataTransferImportRegistryResolverMetadata<TScope>
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
export type PPTCanvasTextPasteReplaceRoute =
  CanvasTextPasteReplaceRoute
export type PPTCanvasTextPasteReplaceTarget =
  CanvasTextPasteReplaceTarget
export type PPTCanvasTextPasteSource = CanvasTextPasteSource
export type PPTCanvasTableImportFormat = CanvasTableImportFormat
export type PPTCanvasTableImportTargetReplaceRoute =
  CanvasTableImportTargetReplaceRoute
export type PPTCanvasTableImportTargetReplaceTarget =
  CanvasTableImportTargetReplaceTarget
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
export type RunPPTCanvasClipboardCommand = RunCanvasClipboardCommand
export type PPTCanvasTabsDescriptor<TId extends string> =
  CanvasTabsDescriptor<TId>
export const PPT_CANVAS_COMPONENT_DEFINITION_REGISTRY =
  CANVAS_COMPONENT_DEFINITION_REGISTRY
