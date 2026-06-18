import {
  ALargeSmall,
  AlignCenter,
  AlignCenterHorizontal,
  AlignCenterVertical,
  AlignEndHorizontal,
  AlignEndVertical,
  AlignHorizontalDistributeCenter,
  AlignLeft,
  AlignRight,
  AlignStartHorizontal,
  AlignStartVertical,
  AlignVerticalDistributeCenter,
  ArrowRight,
  Baseline,
  Bold,
  BringToFront,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  ChevronUp,
  Circle,
  ClipboardPaste,
  Command,
  Copy,
  CopyPlus,
  Diamond,
  Download,
  Eraser,
  Eye,
  EyeOff,
  FilePlus2,
  FlipHorizontal2,
  FlipVertical2,
  Frame,
  Grid2X2,
  Group,
  Hand,
  Highlighter,
  ImagePlus,
  Italic,
  Keyboard,
  Layers,
  List,
  ListOrdered,
  Lock,
  Map as MapIcon,
  Maximize2,
  MessageSquare,
  Minus,
  Moon,
  MousePointer2,
  MoveDown,
  MoveUp,
  Paintbrush,
  PencilLine,
  PenLine,
  Play,
  Plus,
  Redo2,
  RotateCw,
  Ruler,
  Search,
  SendToBack,
  Square,
  StickyNote,
  Sun,
  Table2,
  Trash2,
  Type,
  Undo2,
  Ungroup,
  Unlock,
  Underline,
  X,
  ZoomIn,
  ZoomOut,
} from 'lucide-react'
import {
  Fragment,
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  type ChangeEvent as ReactChangeEvent,
  type ClipboardEvent as ReactClipboardEvent,
  type CSSProperties,
  type DragEvent as ReactDragEvent,
  type FormEvent as ReactFormEvent,
  type KeyboardEvent as ReactKeyboardEvent,
  type MouseEvent as ReactMouseEvent,
  type PointerEvent as ReactPointerEvent,
  type RefObject,
} from 'react'
import {
  CANVAS_DOM_ALIGNMENT_POPOVER_MODEL,
  CANVAS_DOM_ALIGNMENT_PREVIEW_GUIDE_MODEL,
} from './pptDomEditAffordanceAdapter'
import {
  createSlideEditColorSwatchPaletteDescriptor,
  createSlideEditClipboardPasteCommandEffect,
  createSlideEditClipboardPayload,
  createSlideEditObjectAccessibilityDescriptor,
  createSlideEditLayoutPlaceholderDescriptor,
  createSlideEditLayerPaneDescriptor,
  createSlideEditObjectCornerRadiusDescriptor,
  createSlideEditObjectFillOpacityDescriptor,
  createSlideEditObjectHyperlinkDescriptor,
  createSlideEditObjectImageCropDescriptor,
  createSlideEditObjectImageReplaceDescriptor,
  createSlideEditObjectAnimationDescriptor,
  createSlideEditObjectOpacityDescriptor,
  createSlideEditObjectShadowDescriptor,
  createSlideEditObjectStrokeLineStyleDescriptor,
  createSlideEditStyleClipboardDescriptor,
  createSlideEditStyleClipboardPasteCommandEffect,
  createSlideEditThemeDescriptor,
  createSlideEditTextFontFamilyDescriptor,
  createSlideEditTextFrameInsetDescriptor,
  createSlideEditTextParagraphSpacingDescriptor,
  createSlideEditTextVerticalAlignmentDescriptor,
  createSlideEditTransitionDescriptor,
  getSlideEditFrameGuideGeometry,
  getSlideEditColorSwatchCommandEffect,
  getSlideEditColorSwatchId,
  getSlideEditColorWithAlphaCSS,
  getSlideEditDeckNavigationKeyboardIntent,
  getSlideEditLayoutApplyCommandEffect,
  getSlideEditLayerPaneCommandEffect,
  getSlideEditLayerPaneDropIndicator,
  getSlideEditLayerPaneKeyboardIntent,
  getSlideEditLayerPaneResolvedFocusObjectId,
  getSlideEditObjectVisibilityCommandAvailability,
  getSlideEditObjectVisibilityCommandEffect,
  getSlideEditObjectVisibilityState,
  getSlideEditObjectAccessibilityCommandEffect,
  getSlideEditObjectAnimationBuildOrder,
  getSlideEditObjectAnimationCSSStyle,
  getSlideEditObjectCornerRadiusCommandEffect,
  getSlideEditObjectCornerRadiusCSS,
  getSlideEditObjectCornerRadiusPreviewCSS,
  getSlideEditObjectFillOpacityCommandEffect,
  getSlideEditObjectHyperlinkCommandEffect,
  getSlideEditObjectImageCropCommandEffect,
  getSlideEditObjectImageCropPositionCSS,
  getSlideEditObjectImageReplaceCommandEffect,
  getSlideEditObjectAnimationUpdateCommandEffect,
  getSlideEditObjectOpacityCommandEffect,
  getSlideEditObjectShadowCommandEffect,
  getSlideEditObjectShadowFilter,
  normalizeSlideEditObjectAltTextStorageValue,
  normalizeSlideEditObjectHyperlinkStorageUrl,
  getSlideEditLayoutPlaceholderVisibilityDescriptor,
  getSlideEditObjectStrokeLineStyleBorderStyle,
  getSlideEditObjectStrokeLineStyleCommandEffect,
  getSlideEditObjectStrokeLineStyleDashArray,
  createSlideEditRailDescriptor,
  getSlideEditRailKeyboardCommandEffect,
  getSlideEditRailListboxKeyboardIntent,
  getSlideEditRailPointerCommandEffect,
  getSlideEditResolvedLayoutPlaceholder,
  getSlideEditStyleClipboardCategoryDescriptors,
  getSlideEditStyleClipboardCopyCommandEffect,
  getSlideEditStyleClipboardKeyboardIntent,
  getSlideEditStyleClipboardPasteAvailability,
  getSlideEditTextAutoFitGestureCommandEffect,
  getSlideEditTextOverflowIndicatorState,
  getSlideEditTextFontFamilyCSS,
  getSlideEditTextFontFamilyCommandEffect,
  getSlideEditTextFormattingKeyboardIntent,
  getSlideEditTextFrameInsetCommandEffect,
  getSlideEditTextFrameInsetPaddingCSS,
  getSlideEditTextParagraphSpacingCommandEffect,
  getSlideEditTextParagraphSpacingCSSStyle,
  getSlideEditTextVerticalAlignmentCommandEffect,
  getSlideEditTextVerticalAlignmentFlexAlignItems,
  getSlideEditTransitionCSSStyle,
  getSlideEditTransitionUpdateCommandEffect,
  mapSlideEditClipboardPasteObjects,
  normalizeSlideEditClipboardSelectedObjectIds,
  normalizeSlideEditObjectCornerRadius,
  normalizeSlideEditObjectAnimationDelayMs,
  normalizeSlideEditObjectAnimationDurationMs,
  normalizeSlideEditObjectAnimationOrder,
  normalizeSlideEditColorSwatchValue,
  normalizeSlideEditObjectFillOpacity,
  normalizeSlideEditObjectImageCropFit,
  normalizeSlideEditObjectImageCropValue,
  normalizeSlideEditObjectOpacity,
  isSlideEditObjectStrokeLineStyleValue,
  normalizeSlideEditObjectStrokeLineStyle,
  normalizeSlideEditColorHex,
  normalizeSlideEditTextFontFamily,
  normalizeSlideEditTextFrameInsetValue,
  normalizeSlideEditTextLineHeightRatio,
  normalizeSlideEditTextParagraphSpacingAmount,
  normalizeSlideEditTextVerticalAlignment,
  SLIDE_EDIT_DEFAULT_TRANSITION,
  SLIDE_EDIT_OBJECT_ANIMATION_LIMITS,
  SLIDE_EDIT_COLOR_SWATCH_CHANNELS,
  SLIDE_EDIT_OBJECT_ANIMATION_TRIGGERS,
  SLIDE_EDIT_OBJECT_ANIMATION_TYPES,
  SLIDE_EDIT_OBJECT_STROKE_LINE_STYLE_OPTIONS,
  SLIDE_EDIT_STYLE_CLIPBOARD_COPY_FORMATTING_SHORTCUT,
  SLIDE_EDIT_STYLE_CLIPBOARD_PASTE_FORMATTING_SHORTCUT,
  SLIDE_EDIT_TRANSITION_TIMING_LIMITS,
  SLIDE_EDIT_TRANSITION_TYPES,
  SLIDE_EDIT_TEXT_BOX_SIZE_MODES,
  SLIDE_EDIT_TEXT_VERTICAL_ALIGNMENT_OPTIONS,
  SLIDE_EDIT_LAYER_PANE_COMMANDS,
  SLIDE_EDIT_LAYER_PANE_DROP_INDICATOR_MODEL,
  SLIDE_EDIT_LAYER_PANE_KEYBOARD_INTENT_MODEL,
  SLIDE_EDIT_LAYER_PANE_KEYBOARD_KEYS,
  SLIDE_EDIT_RAIL_KEYBOARD_KEYS,
  toSlideEditObjectCornerRadiusAttributeValue,
  toSlideEditObjectFillOpacityAttributeValue,
  toSlideEditObjectOpacityAttributeValue,
  toSlideEditRailHostCommandEffect,
  type SlideEditFrameGuideConfig,
  type SlideEditFrameGuideGeometry,
  type SlideEditLayerPaneCommandDescriptor,
  type SlideEditLayerPaneDescriptor,
  type SlideEditLayerPaneDropPlacement,
  type SlideEditLayerPaneHostCommandEffect,
  type SlideEditLayerPaneIntent,
  type SlideEditLayerPaneKeyboardIntent,
  type SlideEditLayerPaneRowDescriptor,
  type SlideEditLayoutApplyHostCommandEffect,
  type SlideEditLayoutDescriptor,
  type SlideEditMasterDescriptor,
  type SlideEditObjectAccessibilityDescriptor,
  type SlideEditObjectAccessibilityHostCommandEffect,
  type SlideEditObjectCornerRadiusDescriptor,
  type SlideEditObjectCornerRadiusHostCommandEffect,
  type SlideEditObjectFillOpacityDescriptor,
  type SlideEditObjectFillOpacityHostCommandEffect,
  type SlideEditObjectHyperlinkDescriptor,
  type SlideEditObjectHyperlinkHostCommandEffect,
  type SlideEditObjectImageCropDescriptor,
  type SlideEditObjectImageCropHostCommandEffect,
  type SlideEditObjectImageReplaceDescriptor,
  type SlideEditObjectImageReplaceHostCommandEffect,
  type SlideEditBuiltInAnimationTrigger,
  type SlideEditColorSwatchBuiltInChannelId,
  type SlideEditColorSwatchHostCommandEffect,
  type SlideEditColorSwatchPaletteDescriptor,
  type SlideEditColorSwatchSelection,
  type SlideEditBuiltInAnimationType,
  type SlideEditClipboardObjectMetadata,
  type SlideEditClipboardOperation,
  type SlideEditClipboardPasteHostCommandEffect,
  type SlideEditClipboardPasteObjectMapping,
  type SlideEditClipboardPasteTarget,
  type SlideEditClipboardPayload,
  type SlideEditClipboardRemapPolicy,
  type SlideEditObjectAnimationDescriptor,
  type SlideEditObjectAnimationHostCommandEffect,
  type SlideEditObjectAnimationUpdateCommand,
  type SlideEditObjectOpacityDescriptor,
  type SlideEditObjectOpacityHostCommandEffect,
  type SlideEditObjectShadowDescriptor,
  type SlideEditObjectShadowHostCommandEffect,
  type SlideEditObjectStrokeLineStyleDescriptor,
  type SlideEditObjectStrokeLineStyleHostCommandEffect,
  type SlideEditObjectVisibilityCommandAvailability,
  type SlideEditObjectVisibilityCommandId,
  type SlideEditObjectVisibilityDescriptor,
  type SlideEditObjectVisibilityHostCommandEffect,
  type SlideEditObjectSelectionPolicy,
  type SlideEditObjectVisibilityState,
  type SlideEditPlaceholderDescriptor,
  type SlideEditRailListboxOptionDescriptor,
  type SlideEditRailHostCommandEffect,
  type SlideEditRailThumbnailDescriptor,
  type SlideEditResolvedLayoutPlaceholder,
  type SlideEditStyleClipboardBuiltInCategoryId,
  type SlideEditStyleClipboardCopyFormattingCommand,
  type SlideEditStyleClipboardDescriptor,
  type SlideEditStyleClipboardHostCommandEffect,
  type SlideEditStyleClipboardPasteFormattingCommand,
  type SlideEditStyleClipboardTargetInput,
  type SlideEditThemeColorToken,
  type SlideEditTextFontFamilyDescriptor,
  type SlideEditTextFontFamilyHostCommandEffect,
  type SlideEditTextFrameInsetDescriptor,
  type SlideEditTextFrameInsetHostCommandEffect,
  type SlideEditTextAutoFitHostCommandEffect,
  type SlideEditTextBoxMeasurement,
  type SlideEditTextBoxSizeMode,
  type SlideEditTextOverflowIndicatorState,
  type SlideEditTextResizeHandle,
  type SlideEditTextParagraphSpacingDescriptor,
  type SlideEditTextParagraphSpacingFieldId,
  type SlideEditTextParagraphSpacingHostCommandEffect,
  type SlideEditTextParagraphSpacingUpdateCommand,
  type SlideEditTextVerticalAlignmentDescriptor,
  type SlideEditTextVerticalAlignmentHostCommandEffect,
  type SlideEditSlideTransitionDescriptor,
  type SlideEditTransitionHostCommandEffect,
  type SlideEditTransitionUpdateCommand,
} from './pptSlideEditAffordanceAdapter'
import {
  PPT_RESIZE_HANDLES,
  clampPPTCanvasBoundsToFrame,
  clampPPTCanvasPointToBounds,
  clampPPTCanvasValue,
  createPPTCanvasSequentialIdFactory,
  getPPTCanvasBoundsAnchorPoints,
  getPPTCanvasBoundsCenter,
  getPPTCanvasHandlePoint,
  getPPTCanvasPointDistance,
  normalizePPTCanvasBounds,
  normalizePPTCanvasPointsToLocalBounds,
  uniquePPTCanvasValues,
  type Bounds,
  type Point,
  type ResizeHandle,
  type Tool,
  type Viewport,
} from './pptCanvasCoreAdapter'
import {
  EMPTY_PPT_CANVAS_SNAP_GUIDES,
  PPT_MARQUEE_SELECTION_MODEL,
  deletePPTCanvasSelectionItems,
  getPPTCanvasFullySelectedItemGroupIds,
  getPPTCanvasGroupExpandedSelectionIds,
  getPPTCanvasGroupedItemPointerSelection,
  getPPTCanvasGroupedItemSelection,
  getPPTCanvasMarqueeSelection,
  getPPTCanvasMoveSnap,
  getPPTCanvasSelectedItems,
  getPPTCanvasSingleItemSelection,
  isAdditivePPTPointerInput,
  mapPPTCanvasSelectionItems,
  movePPTCanvasSelection,
  normalizePPTCanvasRotationDegrees,
  removePPTCanvasSelectionIds,
  resizePPTCanvasSelection,
  type PPTCanvasSnapGuides,
} from './pptCanvasFoundationAdapter'
import {
  PPT_DEFAULT_THEME_ID,
  PPT_SPLIT_LAYOUT_ID,
  PPT_SLIDE_HEIGHT,
  PPT_SLIDE_WIDTH,
  PPT_TITLE_BODY_LAYOUT_ID,
  PPTDeckSchema,
  PPTElementSchema,
  PPTSlideSchema,
  PPTTextBodySchema,
  createPPTElementId,
  createPPTTextBody,
  findPPTElement,
  findPPTSlide,
  isPPTTextElement,
  readPPTText,
  replacePPTElementText,
  updatePPTDeckElement,
  updatePPTDeckSlide,
  type PPTDeck,
  type PPTComment,
  type PPTCommentThreadMessage,
  type PPTElement,
  type PPTElementAccessibility,
  type PPTElementAnimation,
  type PPTElementHyperlink,
  type PPTElementShadow,
  type PPTFill,
  type PPTFreeform,
  type PPTImage,
  type PPTImageCrop,
  type PPTImageFit,
  type PPTLine,
  type PPTLineConnection,
  type PPTLineMarker,
  type PPTLineRoute,
  type PPTParagraph,
  type PPTRun,
  type PPTShape,
  type PPTShapeKind,
  type PPTSlide,
  type PPTSlideTransition,
  type PPTStroke,
  type PPTStrokeDash,
  type PPTTable,
  type PPTTextBody,
  type PPTTextAutoFit,
  type PPTTextElement,
  type PPTTextStyle,
} from './pptModel'
import { SAMPLE_PPT_DECK } from './pptSampleDeck'
import {
  createPPTCanvasScene,
  canFlipPPTElements,
  getPPTElementGroupIndexRange,
  getPPTElementGroupMemberIds,
  getPPTElementPointerSelection,
  insertPPTSlideAtTargetPlacement,
  movePPTElementsToIndex,
  movePPTSlideToTargetPlacement,
  canSelectSameTypePPTElements,
  pptGeometryToBounds,
  pptCanvasTransformAdapter,
  selectSameTypePPTElements,
  canTidyPPTElements,
  flipPPTElements,
  tidyPPTElements,
} from './pptCanvasAdapter'
import {
  alignPPTCanvasCommand,
  createPPTCanvasAffordanceConfig,
  createPPTCanvasShape,
  createPPTCanvasText,
  deletePPTCanvasCommand,
  distributePPTCanvasCommand,
  duplicatePPTCanvasCommand,
  getPPTCanvasWheelViewport,
  groupPPTCanvasCommand,
  lockPPTCanvasCommand,
  nudgePPTCanvasCommand,
  PPT_COMMAND_AFFORDANCES,
  PPT_TOOL_AFFORDANCES,
  reorderPPTCanvasCommand,
  selectAllPPTCanvasCommand,
  ungroupPPTCanvasCommand,
  unlockAllPPTCanvasCommand,
  type PPTCanvasAlignMode,
  type PPTCanvasCommandItemsResult,
  type PPTCanvasCreatedShapeKind,
  type PPTCanvasCreationAdapter,
  type PPTCanvasDistributeMode,
  type PPTCanvasReorderMode,
} from './pptCanvasEngineAdapter'
import {
  createPPTCanvasCssBoundsTransform,
  createPPTCanvasSvgFreehandPathData,
  createPPTCanvasSvgPathData,
  escapePPTCanvasXmlAttribute,
} from './pptCanvasRendererAdapter'
import {
  PPT_COMMENT_THREAD_MODEL,
  PPT_KEYBOARD_COMMAND_DISPATCH_MODEL,
  PPT_KEYBOARD_COMMAND_INTENT_MODEL,
  PPT_KEYBOARD_NUDGE_INTENT_MODEL,
  PPT_KEYBOARD_NUDGE_KEYS,
  PPT_KEYBOARD_NUDGE_LARGE_STEP,
  PPT_KEYBOARD_NUDGE_MODEL,
  PPT_KEYBOARD_NUDGE_STEP,
  PPT_KEYBOARD_TOOL_DISPATCH_MODEL,
  PPT_KEYBOARD_VIEWPORT_INTENT_MODEL,
  PPT_KEYBOARD_VIEWPORT_MODEL,
  PPT_MENU_FOCUS_MODEL,
  PPT_MENU_ITEM_PROPS,
  PPT_MENU_KEYBOARD_KEYS,
  PPT_MENU_ROVING_FOCUS_MODEL,
  PPT_MODAL_FOCUS_LIFECYCLE_MODEL,
  PPT_RADIO_GROUP_FOCUS_MODEL,
  PPT_RADIO_GROUP_KEYBOARD_MODEL,
  PPT_RADIO_GROUP_MODEL,
  PPT_RESIZE_POINTER_MODIFIERS_MODEL,
  PPT_SELECTION_TOOLBAR_DROPDOWN_MENU_MODEL,
  PPT_TABS_ROVING_FOCUS_MODEL,
  PPT_TOOLBAR_FOCUS_MODEL,
  PPT_TOOLBAR_ITEM_PROPS,
  PPT_TOOLBAR_KEYBOARD_MODEL,
  PPT_TOOLBAR_ROVING_FOCUS_MODEL,
  bindPPTCanvasEventListener,
  bindPPTCanvasEventListeners,
  cancelPPTCanvasAnimationFrameTask,
  cancelPPTCanvasDeferredFocus,
  capturePPTCanvasPointerFromEvent,
  centerPPTCanvasViewportAtWorldPoint,
  createPPTCanvasPastePositionKey,
  createPPTCanvasRichClipboardHTML,
  createPPTCanvasTabsDescriptor,
  downloadPPTCanvasTextFile,
  filterPPTCommandPaletteItems,
  fitPPTCanvasViewportToBounds,
  focusPPTCanvasElement,
  focusPPTCanvasElementBySelectorOnNextFrame,
  focusPPTCanvasElementOnNextFrame,
  getPPTInlineEditHistoryDirectionFromInputType,
  getPPTCanvasClientViewportSize,
  getPPTCanvasCommandPaletteKeyboardIntent,
  getPPTCanvasContextMenuKeyboardIntent,
  getPPTCanvasContextMenuPosition,
  getPPTCanvasDataTransferText,
  getPPTCanvasEditableFieldKeyboardIntent,
  getPPTCanvasFindInputKeyboardIntent,
  getPPTCanvasFloatingAnchorForBounds,
  getPPTCanvasInlineEditKeyboardIntent,
  getPPTCanvasKeyboardBuiltinCommandShortcutIntent,
  getPPTCanvasKeyboardNudgeShortcutIntent,
  getPPTCanvasKeyboardSystemShortcutIntent,
  getPPTCanvasKeyboardToolShortcutIntent,
  getPPTCanvasKeyboardViewportShortcutIntent,
  getPPTCanvasMenuTriggerKeyboardIntent,
  getPPTMinimapPointFromViewportOffset,
  getPPTMinimapReadModel,
  getPPTMinimapWorldPoint,
  getPPTCanvasModalBackdropPointerIntent,
  getPPTCanvasModalKeyboardIntent,
  getPPTCanvasPasteOffsetForBounds,
  getPPTCanvasPastePositionSession,
  getPPTCanvasPointerLocalGeometry,
  getPPTCanvasPointerTransformModifierState,
  getPPTCanvasPresentationKeyboardIntent,
  getPPTCanvasRadioTabIndex,
  getPPTCanvasResizeHandleDoubleClickIntent,
  getPPTCanvasSelectionListModifierState,
  getPPTCanvasTabsKeyboardIntent,
  getPPTCanvasWorldClientPoint,
  handlePPTCanvasRadioGroupKeyDown,
  insertPPTInlineEditText,
  isPPTCanvasControlTarget,
  isPPTCanvasKeyboardCommandIntent,
  isPPTCanvasKeyboardToolIntent,
  isPPTCanvasKeyboardTypingTarget,
  isPPTCanvasKeyboardViewportIntent,
  isPPTCanvasTargetWithinSelector,
  isPPTCanvasWheelPassthroughTarget,
  isPPTInlineEditLineBreakInput,
  measurePPTCanvasElementOverflow,
  measurePPTCanvasTextBlocks,
  PPT_COMMAND_PALETTE_ITEMS_MODEL,
  PPT_INLINE_EDIT_DOM_MODEL,
  PPT_KEYBOARD_TEMPORARY_PAN_MODEL,
  PPT_KEYBOARD_TEMPORARY_PAN_SHORTCUT_LABEL,
  PPT_MINIMAP_READ_MODEL,
  PPT_PASTE_POSITION_MODEL,
  PPT_POINTER_CLICK_MEMORY_MODEL,
  PPT_WHEEL_VIEWPORT_HORIZONTAL_PAN_MODIFIER,
  PPT_WHEEL_VIEWPORT_MODEL,
  PPT_WHEEL_VIEWPORT_PAN_MODE,
  PPT_WHEEL_VIEWPORT_ZOOM_MODIFIER,
  readPPTCanvasRichClipboardFromDataTransfer,
  resetPPTCanvasViewport,
  runPPTCanvasKeyboardCommandIntent,
  runPPTCanvasKeyboardToolIntent,
  runPPTCanvasKeyboardViewportIntent,
  schedulePPTCanvasAnimationFrameTask,
  schedulePPTCanvasTimeoutTask,
  setPPTCanvasDataTransferDropEffect,
  setPPTCanvasDataTransferText,
  shouldReleasePPTCanvasKeyboardTemporaryPan,
  stringifyPPTCanvasRichClipboardPayload,
  trapPPTCanvasModalTabFocus,
  usePPTCanvasMenuRovingFocus,
  usePPTCanvasModalFocusLifecycle,
  usePPTCanvasToolbarRovingFocus,
  writePPTCanvasRichClipboardPayload,
  zoomPPTCanvasViewport,
  type PPTCanvasFloatingAnchor,
  type PPTCanvasKeyboardToolIntent,
  type PPTCanvasPastePositionMemory,
  type PPTCanvasPointerClickMemory,
  type PPTCanvasRichClipboardReadFormat,
  type PPTCanvasRichClipboardWriteMode,
  type PPTCanvasTabsDescriptor,
  type PPTCommandPaletteItemBase,
  type PPTMinimapItemBounds as PPTMinimapItemBoundsBase,
  type PPTMinimapReadModel as PPTMinimapReadModelBase,
  type PPTMinimapSize as PPTMinimapSizeBase,
} from './pptCanvasAppAffordanceAdapter'
import {
  getNextPPTDrawingPoints,
  getPPTPointerStartProjection,
  PPT_LASER_TRAIL_OVERLAY_MODEL,
  previewPPTPointerLaserInteraction,
  previewPPTPointerPanInteraction,
  startPPTPointerLaserInteraction,
  startPPTPointerPanInteraction,
  usePPTCanvasStageElement,
  type PPTPointerLaserInteraction,
  type PPTPointerPanInteraction,
} from './pptCanvasInteractionAdapter'
import {
  getPPTEraserHitElementIds,
  mergePPTEraserHitElementIds,
} from './pptCanvasEraserAdapter'
import {
  createPPTCanvasCommandAdapter,
  createPPTElementIdFactory,
  getPPTElementIdPrefix,
  getPPTElementsBounds,
  getPPTCanvasCommandAvailability,
  updatePPTElementBounds,
} from './pptCommandAdapter'
import {
  exportPPTDeckHTML,
  exportPPTSelectionSVG,
  exportPPTSlideSVG,
} from './pptExport'
import {
  PPT_DEFAULT_TABLE_ROWS,
  PPT_DECK_MARKDOWN_OUTLINE_IMPORT_FORMAT,
  PPT_DECK_MARKDOWN_OUTLINE_IMPORT_MODEL,
  PPT_FALLBACK_HTML_IMPORT_MODEL,
  PPT_IMPORT_EXTENSION,
  canHandlePPTStageDropImport,
  createPPTDeckMarkdownOutlineSlides,
  createPPTFallbackHTMLImportEffect,
  createPPTFallbackHTMLImageElement,
  createPPTFallbackHTMLSelectionElements,
  createPPTFallbackHTMLSelectionImportEffect,
  createPPTFallbackHTMLShapeElement,
  createPPTFallbackHTMLTableElement,
  createPPTFallbackHTMLTextElement,
  createPPTImageImportEffect,
  createPPTImportedImageElements,
  getPPTImageFileFromList,
  createPPTMediaElement,
  createPPTTableElements,
  createPPTTableImportEffect,
  createPPTRichTextPasteElement,
  createPPTTextPasteElement,
  getPPTClipboardImportActions,
  getPPTDeckMarkdownOutlineSourceFromDataTransfer,
  getPPTStageDropImportAction,
  getPPTTableColumnCount,
  getPPTTextPasteSourcesFromDataTransfer,
  getPPTMediaSourceFromText,
  normalizePPTTableRows,
  PPT_IMAGE_IMPORT_MODEL,
  readPPTClipboardImageSource,
  readPPTImageFileSource,
  readPPTTableFileSource,
  resolvePPTImageSourceNaturalSize,
  type PPTImageImportSource,
  stringifyPPTTableRows,
  PPT_MEDIA_IMPORT_MODEL,
  PPT_TABLE_IMPORT_MODEL,
  PPT_TEXT_PASTE_IMPORT_MODEL,
  type PPTClipboardImportAction,
  type PPTDeckMarkdownOutlineSource,
  type PPTFallbackHTMLImportEffect,
  type PPTFallbackHTMLImageSource,
  type PPTFallbackHTMLSelectionSource,
  type PPTFallbackHTMLShapeSource,
  type PPTFallbackHTMLTableSource,
  type PPTFallbackHTMLTextSource,
  type PPTImageImportEffect,
  type PPTMediaImportResult,
  type PPTMediaImportSource,
  type PPTRichTextPasteSource,
  type PPTStageDropImportAction,
  type PPTTableImportEffect,
  type PPTTableImportSource,
  type PPTTextPasteImportResult,
} from './pptImportExtension'
import './App.css'

const PPT_CANVAS_COMMAND_CONFIG = createPPTCanvasAffordanceConfig({
  commands: {
    group: true,
    lockSelection: true,
    ungroup: true,
    unlockAll: true,
  },
})
const PPT_CANVAS_STANDARD_COMMAND_INTENT_KINDS = new Set([
  'copy-selection',
  'cut-selection',
  'delete-selection',
  'duplicate-selection',
  'group-selection',
  'lock-selection',
  'paste-selection',
  'redo-history',
  'reorder-selection',
  'select-all',
  'ungroup-selection',
  'undo-history',
  'unlock-all',
])
const PPT_RECENT_COLOR_LIMIT = 8
const PPT_PARAGRAPH_ALIGN_OPTIONS = ['left', 'center', 'right'] as const
const PPT_SLIDE_RAIL_HIT_TARGET_PADDING = 6
const PPT_SLIDE_RAIL_THUMB_GAP = 8
const PPT_SLIDE_RAIL_THUMB_HEIGHT = 86
const PPT_SLIDE_RAIL_THUMB_PADDING = 10
const PPT_SLIDE_RAIL_THUMB_WIDTH = 112

const PPT_FRAME_GUIDE_CONFIG = Object.freeze({
  columns: {
    count: 4,
    gutter: 18,
    margin: 84,
  },
  margin: {
    bottom: 64,
    left: 84,
    right: 84,
    top: 64,
  },
  rulerGuides: [
    { axis: 'x', id: 'ruler-title-left', offset: 128 },
    { axis: 'y', id: 'ruler-title-baseline', offset: 122 },
  ],
  safeArea: {
    bottom: 76,
    left: 92,
    right: 92,
    top: 76,
  },
} as const satisfies SlideEditFrameGuideConfig)

const PPT_MASTER_ID = 'ppt-master-default'

const PPT_THEME_DESCRIPTOR = createSlideEditThemeDescriptor({
  colorTokens: [
    {
      label: 'Background',
      role: 'background',
      tokenId: 'ppt-color-background',
      value: '#f8fafc',
    },
    {
      label: 'Surface',
      role: 'surface',
      tokenId: 'ppt-color-surface',
      value: '#ffffff',
    },
    {
      label: 'Text',
      role: 'text',
      tokenId: 'ppt-color-text',
      value: '#111827',
    },
    {
      label: 'Accent',
      role: 'accent',
      tokenId: 'ppt-color-accent',
      value: '#2563eb',
    },
  ],
  defaultStyle: {
    colorTokenIds: {
      background: 'ppt-color-background',
      text: 'ppt-color-text',
    },
    fontTokenIds: {
      body: 'ppt-font-body',
      heading: 'ppt-font-heading',
    },
    spacingTokenIds: {
      gap: 'ppt-space-gap',
      margin: 'ppt-space-margin',
    },
  },
  fontTokens: [
    {
      family: 'Inter',
      label: 'Body',
      role: 'body',
      size: 24,
      tokenId: 'ppt-font-body',
      weight: 500,
    },
    {
      family: 'Inter Display',
      label: 'Heading',
      role: 'heading',
      size: 56,
      tokenId: 'ppt-font-heading',
      weight: 700,
    },
  ],
  name: 'PPT default',
  spacingTokens: [
    {
      label: 'Margin',
      role: 'slide-margin',
      tokenId: 'ppt-space-margin',
      value: 84,
    },
    {
      label: 'Gap',
      role: 'object-gap',
      tokenId: 'ppt-space-gap',
      value: 28,
    },
  ],
  themeId: PPT_DEFAULT_THEME_ID,
})

const PPT_MASTER_DESCRIPTOR: SlideEditMasterDescriptor = {
  defaultStyle: {
    colorTokenIds: {
      background: 'ppt-color-background',
      text: 'ppt-color-text',
    },
    fontTokenIds: {
      body: 'ppt-font-body',
      heading: 'ppt-font-heading',
    },
  },
  layoutIds: [
    PPT_TITLE_BODY_LAYOUT_ID,
    PPT_SPLIT_LAYOUT_ID,
  ],
  masterId: PPT_MASTER_ID,
  name: 'PPT master',
  themeId: PPT_DEFAULT_THEME_ID,
}

const PPT_LAYOUT_DESCRIPTORS: SlideEditLayoutDescriptor[] = [
  {
    defaultStyle: {
      spacingTokenIds: {
        gap: 'ppt-space-gap',
        margin: 'ppt-space-margin',
      },
    },
    layoutId: PPT_TITLE_BODY_LAYOUT_ID,
    masterId: PPT_MASTER_ID,
    name: 'Title and body',
    placeholders: [
      createSlideEditLayoutPlaceholderDescriptor({
        defaultBounds: { h: 96, w: 900, x: 84, y: 76 },
        defaultStyle: {
          fontTokenIds: {
            heading: 'ppt-font-heading',
          },
        },
        placeholderId: 'title',
        role: 'title',
        title: 'Title',
      }),
      createSlideEditLayoutPlaceholderDescriptor({
        defaultBounds: { h: 420, w: 900, x: 84, y: 214 },
        placeholderId: 'body',
        role: 'body',
        title: 'Body',
      }),
    ],
  },
  {
    defaultStyle: {
      spacingTokenIds: {
        gap: 'ppt-space-gap',
        margin: 'ppt-space-margin',
      },
    },
    layoutId: PPT_SPLIT_LAYOUT_ID,
    masterId: PPT_MASTER_ID,
    name: 'Title with side panel',
    placeholders: [
      createSlideEditLayoutPlaceholderDescriptor({
        defaultBounds: { h: 86, w: 700, x: 84, y: 76 },
        defaultStyle: {
          fontTokenIds: {
            heading: 'ppt-font-heading',
          },
        },
        placeholderId: 'title',
        role: 'title',
        title: 'Title',
      }),
      createSlideEditLayoutPlaceholderDescriptor({
        defaultBounds: { h: 404, w: 680, x: 84, y: 214 },
        placeholderId: 'body',
        role: 'body',
        title: 'Body',
      }),
      createSlideEditLayoutPlaceholderDescriptor({
        defaultBounds: { h: 506, w: 300, x: 896, y: 112 },
        defaultStyle: {
          colorTokenIds: {
            fill: 'ppt-color-surface',
          },
        },
        placeholderId: 'media',
        role: 'media',
        title: 'Side panel',
      }),
    ],
  },
]

const PPT_LAYOUT_BY_ID = new Map(PPT_LAYOUT_DESCRIPTORS.map((layout) => [
  layout.layoutId,
  layout,
]))

function getPPTLayoutDescriptor(layoutId: string | null | undefined) {
  return PPT_LAYOUT_BY_ID.get(layoutId ?? '') ?? PPT_LAYOUT_DESCRIPTORS[0]
}

function getPPTLayoutPlaceholders(
  layout: SlideEditLayoutDescriptor,
  slide: PPTSlide,
): SlideEditResolvedLayoutPlaceholder[] {
  const hiddenPlaceholderIds = getPPTHiddenPlaceholderIdSet(slide)

  return layout.placeholders.map((placeholder) =>
    withPPTSlidePlaceholderVisibilityOverride(getSlideEditResolvedLayoutPlaceholder({
      layout,
      master: PPT_MASTER_DESCRIPTOR,
      placeholder,
      theme: PPT_THEME_DESCRIPTOR,
    }), hiddenPlaceholderIds))
}

function getPPTLayoutPlaceholderVisibilityDescriptors(
  layout: SlideEditLayoutDescriptor,
  slide: PPTSlide,
): SlideEditPlaceholderDescriptor<string, string>[] {
  const hiddenPlaceholderIds = getPPTHiddenPlaceholderIdSet(slide)

  return layout.placeholders.map((placeholder) =>
    withPPTSlidePlaceholderVisibilityOverride(
      getSlideEditLayoutPlaceholderVisibilityDescriptor({
        placeholder,
        slideId: slide.id,
      }),
      hiddenPlaceholderIds,
    ))
}

function withPPTSlidePlaceholderVisibilityOverride<
  TPlaceholder extends { isVisible: boolean; placeholderId: string },
>(
  placeholder: TPlaceholder,
  hiddenPlaceholderIds: ReadonlySet<string>,
): TPlaceholder {
  return hiddenPlaceholderIds.has(placeholder.placeholderId)
    ? {
        ...placeholder,
        isVisible: false,
      }
    : placeholder
}

function getPPTHiddenPlaceholderIdSet(slide: PPTSlide) {
  return new Set(slide.hiddenPlaceholderIds ?? [])
}

const canvasAlignModeAvailabilityKey = {
  alignBottom: 'alignBottom',
  alignCenter: 'alignCenter',
  alignLeft: 'alignLeft',
  alignMiddle: 'alignMiddle',
  alignRight: 'alignRight',
  alignTop: 'alignTop',
} as const satisfies Record<
  PPTCanvasAlignMode,
  keyof ReturnType<typeof getPPTCanvasCommandAvailability>
>

const canvasDistributeModeAvailabilityKey = {
  distributeHorizontal: 'distributeHorizontal',
  distributeVertical: 'distributeVertical',
} as const satisfies Record<
  PPTCanvasDistributeMode,
  keyof ReturnType<typeof getPPTCanvasCommandAvailability>
>

const canvasReorderModeAvailabilityKey = {
  bringForward: 'bringForward',
  bringToFront: 'bringToFront',
  sendBackward: 'sendBackward',
  sendToBack: 'sendToBack',
} as const satisfies Record<
  PPTCanvasReorderMode,
  keyof ReturnType<typeof getPPTCanvasCommandAvailability>
>

type PPTCommandSurface = 'context-menu' | 'selection-floating-bar'
type PPTClipboardOperation = SlideEditClipboardOperation
type PPTClipboardObjectMetadata = SlideEditClipboardObjectMetadata<string, string, string>
type PPTClipboardPayload = SlideEditClipboardPayload<string, string, PPTElement, string, string>
type PPTClipboard = PPTClipboardPayload
type PPTClipboardPasteTarget = SlideEditClipboardPasteTarget<string>
type PPTClipboardPasteObjectMapping =
  SlideEditClipboardPasteObjectMapping<string, string, string>
type PPTClipboardPasteHostCommandEffect =
  SlideEditClipboardPasteHostCommandEffect<string, string, PPTElement, string, string>
type PPTClipboardPastePositionEffect = {
  anchor: Point
  clipboardBounds: Bounds | null
  clipboardObjectCount: number
  model: typeof PPT_PASTE_POSITION_MODEL
  pasteIndex: number
  viewportCenter: Point | null
}
type PPTClipboardPastePositionMemory = PPTCanvasPastePositionMemory
const PPT_RICH_CLIPBOARD_MODEL = 'canvas-board-io-ppt-rich-clipboard' as const
const PPT_RICH_CLIPBOARD_KIND = 'interactive-os.ppt.selection' as const
const PPT_RICH_CLIPBOARD_VERSION = 1
const PPT_RICH_CLIPBOARD_JSON_MIME_TYPE =
  'application/vnd.interactive-os.ppt.selection+json'
const PPT_RICH_CLIPBOARD_HTML_ROOT_ATTRIBUTE = 'data-ppt-rich-clipboard'
const PPT_RICH_CLIPBOARD_HTML_JSON_SCRIPT_ATTRIBUTE =
  'data-ppt-rich-clipboard-json'
const PPT_SLIDE_CLIPBOARD_MODEL = 'canvas-board-io-ppt-slide-clipboard' as const
const PPT_SLIDE_CLIPBOARD_KIND = 'interactive-os.ppt.slide' as const
const PPT_SLIDE_CLIPBOARD_VERSION = 1
const PPT_SLIDE_CLIPBOARD_JSON_MIME_TYPE =
  'application/vnd.interactive-os.ppt.slide+json'
const PPT_SLIDE_CLIPBOARD_HTML_ROOT_ATTRIBUTE = 'data-ppt-slide-clipboard'
const PPT_SLIDE_CLIPBOARD_HTML_JSON_SCRIPT_ATTRIBUTE =
  'data-ppt-slide-clipboard-json'
const PPT_SLIDE_JSON_IMPORT_MODEL = 'ppt-slide-json-import' as const
const PPT_SLIDE_JSON_IMPORT_FORMAT = 'application-json-ppt-slide' as const
const PPT_SLIDE_JSON_TEXT_IMPORT_FORMAT = 'text-json-ppt-slide' as const
const PPT_SLIDE_JSON_MIME_TYPE =
  'application/vnd.interactive-os.ppt.slide.raw+json'
const PPT_SLIDE_NOTES_IMPORT_MODEL = 'ppt-slide-notes-import' as const
const PPT_SLIDE_NOTES_JSON_IMPORT_FORMAT = 'application-json-ppt-notes' as const
const PPT_SLIDE_NOTES_MARKDOWN_IMPORT_FORMAT =
  'text-markdown-ppt-notes' as const
const PPT_SLIDE_NOTES_TEXT_IMPORT_FORMAT = 'text-plain-ppt-notes' as const
const PPT_SLIDE_NOTES_JSON_MIME_TYPE =
  'application/vnd.interactive-os.ppt.notes+json'
const PPT_SLIDE_METADATA_IMPORT_MODEL = 'ppt-slide-metadata-import' as const
const PPT_SLIDE_METADATA_JSON_IMPORT_FORMAT =
  'application-json-ppt-slide-metadata' as const
const PPT_SLIDE_METADATA_JSON_MIME_TYPE =
  'application/vnd.interactive-os.ppt.slide-metadata+json'
const PPT_SLIDE_LAYOUT_IMPORT_MODEL = 'ppt-slide-layout-import' as const
const PPT_SLIDE_LAYOUT_JSON_IMPORT_FORMAT =
  'application-json-ppt-slide-layout' as const
const PPT_SLIDE_LAYOUT_JSON_MIME_TYPE =
  'application/vnd.interactive-os.ppt.slide-layout+json'
const PPT_SLIDE_TRANSITION_IMPORT_MODEL = 'ppt-slide-transition-import' as const
const PPT_SLIDE_TRANSITION_JSON_IMPORT_FORMAT =
  'application-json-ppt-slide-transition' as const
const PPT_SLIDE_TRANSITION_JSON_MIME_TYPE =
  'application/vnd.interactive-os.ppt.slide-transition+json'
const PPT_OBJECT_ANIMATION_IMPORT_MODEL = 'ppt-object-animation-import' as const
const PPT_OBJECT_ANIMATION_JSON_IMPORT_FORMAT =
  'application-json-ppt-object-animation' as const
const PPT_OBJECT_ANIMATION_JSON_MIME_TYPE =
  'application/vnd.interactive-os.ppt.object-animation+json'
const PPT_OBJECT_STYLE_IMPORT_MODEL = 'ppt-object-style-import' as const
const PPT_OBJECT_STYLE_JSON_IMPORT_FORMAT =
  'application-json-ppt-object-style' as const
const PPT_OBJECT_STYLE_JSON_MIME_TYPE =
  'application/vnd.interactive-os.ppt.object-style+json'
const PPT_OBJECT_METADATA_IMPORT_MODEL = 'ppt-object-metadata-import' as const
const PPT_OBJECT_METADATA_JSON_IMPORT_FORMAT =
  'application-json-ppt-object-metadata' as const
const PPT_OBJECT_METADATA_JSON_MIME_TYPE =
  'application/vnd.interactive-os.ppt.object-metadata+json'
const PPT_OBJECT_STATE_IMPORT_MODEL = 'ppt-object-state-import' as const
const PPT_OBJECT_STATE_JSON_IMPORT_FORMAT =
  'application-json-ppt-object-state' as const
const PPT_OBJECT_STATE_JSON_MIME_TYPE =
  'application/vnd.interactive-os.ppt.object-state+json'
const PPT_OBJECT_LAYER_IMPORT_MODEL = 'ppt-object-layer-import' as const
const PPT_OBJECT_LAYER_JSON_IMPORT_FORMAT =
  'application-json-ppt-object-layer' as const
const PPT_OBJECT_LAYER_JSON_MIME_TYPE =
  'application/vnd.interactive-os.ppt.object-layer+json'
const PPT_IMAGE_CROP_IMPORT_MODEL = 'ppt-image-crop-import' as const
const PPT_IMAGE_CROP_JSON_IMPORT_FORMAT =
  'application-json-ppt-image-crop' as const
const PPT_IMAGE_CROP_JSON_MIME_TYPE =
  'application/vnd.interactive-os.ppt.image-crop+json'
const PPT_IMAGE_REPLACE_IMPORT_MODEL = 'ppt-image-replace-import' as const
const PPT_IMAGE_REPLACE_JSON_IMPORT_FORMAT =
  'application-json-ppt-image-replace' as const
const PPT_IMAGE_REPLACE_JSON_MIME_TYPE =
  'application/vnd.interactive-os.ppt.image-replace+json'
const PPT_SHAPE_STYLE_IMPORT_MODEL = 'ppt-shape-style-import' as const
const PPT_SHAPE_STYLE_JSON_IMPORT_FORMAT =
  'application-json-ppt-shape-style' as const
const PPT_SHAPE_STYLE_JSON_MIME_TYPE =
  'application/vnd.interactive-os.ppt.shape-style+json'
const PPT_TEXT_STYLE_IMPORT_MODEL = 'ppt-text-style-import' as const
const PPT_TEXT_STYLE_JSON_IMPORT_FORMAT =
  'application-json-ppt-text-style' as const
const PPT_TEXT_STYLE_JSON_MIME_TYPE =
  'application/vnd.interactive-os.ppt.text-style+json'
const PPT_TEXT_BODY_IMPORT_MODEL = 'ppt-text-body-import' as const
const PPT_TEXT_BODY_JSON_IMPORT_FORMAT =
  'application-json-ppt-text-body' as const
const PPT_TEXT_BODY_JSON_MIME_TYPE =
  'application/vnd.interactive-os.ppt.text-body+json'
const PPT_TEXT_AUTOFIT_IMPORT_MODEL = 'ppt-text-autofit-import' as const
const PPT_TEXT_AUTOFIT_JSON_IMPORT_FORMAT =
  'application-json-ppt-text-autofit' as const
const PPT_TEXT_AUTOFIT_JSON_MIME_TYPE =
  'application/vnd.interactive-os.ppt.text-autofit+json'
const PPT_LINE_STYLE_IMPORT_MODEL = 'ppt-line-style-import' as const
const PPT_LINE_STYLE_JSON_IMPORT_FORMAT =
  'application-json-ppt-line-style' as const
const PPT_LINE_STYLE_JSON_MIME_TYPE =
  'application/vnd.interactive-os.ppt.line-style+json'
const PPT_OBJECT_TRANSFORM_IMPORT_MODEL = 'ppt-object-transform-import' as const
const PPT_OBJECT_TRANSFORM_JSON_IMPORT_FORMAT =
  'application-json-ppt-object-transform' as const
const PPT_OBJECT_TRANSFORM_JSON_MIME_TYPE =
  'application/vnd.interactive-os.ppt.object-transform+json'
const PPT_MEDIA_JSON_IMPORT_MODEL = 'ppt-media-json-import' as const
const PPT_MEDIA_JSON_IMPORT_FORMAT = 'application-json-ppt-media' as const
const PPT_MEDIA_JSON_MIME_TYPE =
  'application/vnd.interactive-os.ppt.media+json'
const PPT_COMMENT_IMPORT_MODEL = 'ppt-comment-import' as const
const PPT_COMMENT_JSON_IMPORT_FORMAT =
  'application-json-ppt-comment' as const
const PPT_COMMENT_JSON_MIME_TYPE =
  'application/vnd.interactive-os.ppt.comment+json'
const PPT_HTML_CLIPBOARD_MODEL = 'canvas-rich-html-clipboard' as const
const PPT_HTML_CLIPBOARD_KIND = 'interactive-os.ppt.html-export' as const
const PPT_HTML_CLIPBOARD_VERSION = 1
const PPT_HTML_CLIPBOARD_JSON_MIME_TYPE =
  'application/vnd.interactive-os.ppt.html-export+json'
const PPT_DECK_HTML_IMPORT_MODEL = 'ppt-deck-html-import' as const
const PPT_DECK_HTML_IMPORT_FORMAT = 'text-html-ppt-deck' as const
const PPT_DECK_HTML_FALLBACK_IMPORT_FORMAT =
  'text-html-ppt-deck-fallback' as const
const PPT_DECK_JSON_IMPORT_MODEL = 'ppt-deck-json-import' as const
const PPT_DECK_JSON_IMPORT_FORMAT = 'application-json-ppt-deck' as const
const PPT_DECK_JSON_TEXT_IMPORT_FORMAT = 'text-json-ppt-deck' as const
const PPT_DECK_JSON_MIME_TYPE =
  'application/vnd.interactive-os.ppt.deck+json'
const PPT_ELEMENTS_JSON_IMPORT_MODEL = 'ppt-elements-json-import' as const
const PPT_ELEMENTS_JSON_IMPORT_FORMAT = 'application-json-ppt-elements' as const
const PPT_ELEMENTS_JSON_TEXT_IMPORT_FORMAT = 'text-json-ppt-elements' as const
const PPT_ELEMENTS_JSON_MIME_TYPE =
  'application/vnd.interactive-os.ppt.elements+json'
const PPT_SLIDE_SVG_CLIPBOARD_MODEL = 'canvas-rich-slide-svg-clipboard' as const
const PPT_SLIDE_SVG_CLIPBOARD_KIND = 'interactive-os.ppt.slide-svg-export' as const
const PPT_SLIDE_SVG_CLIPBOARD_VERSION = 1
const PPT_SLIDE_SVG_CLIPBOARD_JSON_MIME_TYPE =
  'application/vnd.interactive-os.ppt.slide-svg-export+json'
const PPT_SELECTION_SVG_CLIPBOARD_MODEL = 'canvas-rich-selection-svg-clipboard' as const
const PPT_SELECTION_SVG_CLIPBOARD_KIND =
  'interactive-os.ppt.selection-svg-export' as const
const PPT_SELECTION_SVG_CLIPBOARD_VERSION = 1
const PPT_SELECTION_SVG_CLIPBOARD_JSON_MIME_TYPE =
  'application/vnd.interactive-os.ppt.selection-svg-export+json'
const PPT_TABLE_CLIPBOARD_MODEL = 'canvas-rich-table-clipboard' as const
const PPT_TABLE_CLIPBOARD_KIND = 'interactive-os.ppt.table-export' as const
const PPT_TABLE_CLIPBOARD_VERSION = 1
const PPT_TABLE_CLIPBOARD_JSON_MIME_TYPE =
  'application/vnd.interactive-os.ppt.table-export+json'
const PPT_TABLE_ROWS_IMPORT_MODEL = 'ppt-table-rows-import' as const
const PPT_TABLE_ROWS_JSON_IMPORT_FORMAT =
  'application-json-ppt-table-rows' as const
const PPT_TABLE_ROWS_JSON_MIME_TYPE =
  'application/vnd.interactive-os.ppt.table-rows+json'
const PPT_RICH_CLIPBOARD_FORMATS = [
  PPT_RICH_CLIPBOARD_JSON_MIME_TYPE,
  'text/html',
  'image/svg+xml',
  'text/plain',
] as const
type PPTRichClipboardImportFormat = PPTCanvasRichClipboardReadFormat
type PPTRichClipboardWriteMode = PPTCanvasRichClipboardWriteMode | 'pending'
type PPTRichClipboardExportPayload = {
  kind: typeof PPT_RICH_CLIPBOARD_KIND
  metadata: {
    objectCount: number
    selectedObjectIds: readonly string[]
    sourceSlideId: string
  }
  payload: PPTClipboardPayload
  version: typeof PPT_RICH_CLIPBOARD_VERSION
}
type PPTRichClipboardEffect = {
  formats: readonly string[]
  htmlLength: number
  importFormat?: PPTRichClipboardImportFormat
  imported?: boolean
  model: typeof PPT_RICH_CLIPBOARD_MODEL
  objectCount: number
  plainTextLength: number
  selectedObjectIds: readonly string[]
  sourceSlideId: string
  writeMode?: PPTRichClipboardWriteMode
}
type PPTRichClipboardFallback = {
  html: string
  plainText: string
}
type PPTSlideClipboardPayload = {
  slide: PPTSlide
  sourceSlideId: string
}
type PPTSlideClipboardFallbackHTMLSource = {
  htmlLength: number
  name: string
  sourceSlideId?: string
  svg: string
}
type PPTSlideClipboardExportPayload = {
  kind: typeof PPT_SLIDE_CLIPBOARD_KIND
  metadata: {
    elementCount: number
    sourceSlideId: string
  }
  payload: PPTSlideClipboardPayload
  version: typeof PPT_SLIDE_CLIPBOARD_VERSION
}
type PPTSlideClipboardEffect = {
  elementCount: number
  htmlLength: number
  importFormat?: PPTRichClipboardImportFormat | 'text-html-fallback'
  imported?: boolean
  jsonMimeType: typeof PPT_SLIDE_CLIPBOARD_JSON_MIME_TYPE
  model: typeof PPT_SLIDE_CLIPBOARD_MODEL
  slideName: string
  sourceSlideId: string
  targetSlideId?: string
  writeMode?: PPTRichClipboardWriteMode
}
type PPTHTMLClipboardEffect = {
  htmlLength: number
  jsonMimeType: typeof PPT_HTML_CLIPBOARD_JSON_MIME_TYPE
  model: typeof PPT_HTML_CLIPBOARD_MODEL
  sourceSlideId: string
  writeMode?: PPTRichClipboardWriteMode
}
type PPTDeckHTMLImportSource = {
  deck: PPTDeck
  htmlLength: number
}
type PPTDeckJSONImportSource = {
  deck: PPTDeck
  format:
    | typeof PPT_DECK_JSON_IMPORT_FORMAT
    | typeof PPT_DECK_JSON_TEXT_IMPORT_FORMAT
  jsonLength: number
}
type PPTSlideJSONImportSource = {
  format:
    | typeof PPT_SLIDE_JSON_IMPORT_FORMAT
    | typeof PPT_SLIDE_JSON_TEXT_IMPORT_FORMAT
  jsonLength: number
  slides: readonly PPTSlide[]
}
type PPTSlideNotesImportSource = {
  format:
    | typeof PPT_SLIDE_NOTES_JSON_IMPORT_FORMAT
    | typeof PPT_SLIDE_NOTES_MARKDOWN_IMPORT_FORMAT
    | typeof PPT_SLIDE_NOTES_TEXT_IMPORT_FORMAT
  notes: string
  textLength: number
}
type PPTSlideMetadataImportSource = {
  background?: PPTSlideMetadataBackgroundDescriptor
  format: typeof PPT_SLIDE_METADATA_JSON_IMPORT_FORMAT
  jsonLength: number
  name?: string
  notes?: string
}
type PPTSlideLayoutImportField =
  | 'hiddenPlaceholderIds'
  | 'layoutId'
  | 'themeId'
type PPTSlideLayoutImportSource = {
  fields: readonly PPTSlideLayoutImportField[]
  format: typeof PPT_SLIDE_LAYOUT_JSON_IMPORT_FORMAT
  hiddenPlaceholderIds?: readonly string[]
  jsonLength: number
  layoutId?: string
  themeId?: string
}
type PPTSlideTransitionImportField =
  | 'advanceAfterMs'
  | 'advanceOnClick'
  | 'durationMs'
  | 'type'
type PPTSlideTransitionImportSource = {
  fields: readonly PPTSlideTransitionImportField[]
  format: typeof PPT_SLIDE_TRANSITION_JSON_IMPORT_FORMAT
  jsonLength: number
  transition: Partial<PPTSlideTransition>
}
type PPTObjectAnimationImportField = keyof PPTElementAnimation
type PPTObjectAnimationImportSource = {
  animation: Partial<PPTElementAnimation>
  fields: readonly PPTObjectAnimationImportField[]
  format: typeof PPT_OBJECT_ANIMATION_JSON_IMPORT_FORMAT
  jsonLength: number
}
type PPTObjectStyleImportField =
  | 'opacity'
  | 'shadow'
type PPTObjectStyleImportSource = {
  fields: readonly PPTObjectStyleImportField[]
  format: typeof PPT_OBJECT_STYLE_JSON_IMPORT_FORMAT
  jsonLength: number
  object: {
    opacity?: number
    shadow?: PPTElementShadow | null
  }
}
type PPTObjectMetadataImportField =
  | 'altText'
  | 'hyperlinkUrl'
  | 'name'
type PPTObjectMetadataImportSource = {
  fields: readonly PPTObjectMetadataImportField[]
  format: typeof PPT_OBJECT_METADATA_JSON_IMPORT_FORMAT
  jsonLength: number
  metadata: {
    altText?: string | null
    hyperlinkUrl?: string | null
    name?: string
  }
}
type PPTObjectStateImportField =
  | 'locked'
  | 'visible'
type PPTObjectStateImportSource = {
  fields: readonly PPTObjectStateImportField[]
  format: typeof PPT_OBJECT_STATE_JSON_IMPORT_FORMAT
  jsonLength: number
  state: {
    locked?: boolean
    visible?: boolean
  }
}
type PPTObjectLayerImportPosition =
  | 'back'
  | 'backward'
  | 'forward'
  | 'front'
type PPTObjectLayerImportField =
  | 'position'
  | 'toIndex'
type PPTObjectLayerImportSource = {
  fields: readonly PPTObjectLayerImportField[]
  format: typeof PPT_OBJECT_LAYER_JSON_IMPORT_FORMAT
  jsonLength: number
  layer: {
    position?: PPTObjectLayerImportPosition
    toIndex?: number
  }
}
type PPTImageCropImportField =
  | 'fit'
  | 'x'
  | 'y'
type PPTImageCropImportSource = {
  fields: readonly PPTImageCropImportField[]
  format: typeof PPT_IMAGE_CROP_JSON_IMPORT_FORMAT
  imageCrop: {
    crop?: Partial<PPTImageCrop>
    fit?: PPTImageFit
  }
  jsonLength: number
}
type PPTImageReplaceImportField =
  | 'altText'
  | 'mimeType'
  | 'name'
  | 'naturalHeight'
  | 'naturalWidth'
  | 'src'
type PPTImageReplaceImportImageSource = PPTImageImportSource & {
  altText?: string
}
type PPTImageReplaceImportSource = {
  fields: readonly PPTImageReplaceImportField[]
  format: typeof PPT_IMAGE_REPLACE_JSON_IMPORT_FORMAT
  image: PPTImageReplaceImportImageSource
  jsonLength: number
  resolveNaturalSize: boolean
}
type PPTShapeStyleImportField =
  | 'cornerRadius'
  | 'fill'
  | 'stroke'
type PPTShapeStyleImportSource = {
  fields: readonly PPTShapeStyleImportField[]
  format: typeof PPT_SHAPE_STYLE_JSON_IMPORT_FORMAT
  jsonLength: number
  shape: {
    cornerRadius?: number
    fill?: PPTFill
    stroke?: PPTStroke
  }
}
type PPTTextStyleImportField =
  | 'color'
  | 'fontFamily'
  | 'fontSize'
  | 'fontWeight'
  | 'paragraphAlign'
  | 'paragraphBullet'
  | 'paragraphLineHeight'
  | 'paragraphSpacingAfter'
  | 'paragraphSpacingBefore'
  | 'textInset'
  | 'verticalAlign'
type PPTTextStyleImportText = Partial<Omit<PPTTextStyle, 'textInset'>> & {
  textInset?: Partial<NonNullable<PPTTextStyle['textInset']>>
}
type PPTTextStyleImportParagraph = {
  align?: PPTParagraph['align']
  bullet?: PPTParagraph['bullet'] | null
  lineHeight?: PPTParagraph['lineHeight']
  spacingAfter?: PPTParagraph['spacingAfter']
  spacingBefore?: PPTParagraph['spacingBefore']
}
type PPTTextStyleImportSource = {
  fields: readonly PPTTextStyleImportField[]
  format: typeof PPT_TEXT_STYLE_JSON_IMPORT_FORMAT
  jsonLength: number
  paragraph?: PPTTextStyleImportParagraph
  text?: PPTTextStyleImportText
}
type PPTTextBodyImportSource = {
  format: typeof PPT_TEXT_BODY_JSON_IMPORT_FORMAT
  jsonLength: number
  mode: 'plain-text' | 'text-body'
  textBody: PPTTextBody
}
type PPTTextAutoFitImportField =
  | 'handle'
  | 'mode'
type PPTTextAutoFitImportSource = {
  fields: readonly PPTTextAutoFitImportField[]
  format: typeof PPT_TEXT_AUTOFIT_JSON_IMPORT_FORMAT
  handle: ResizeHandle
  jsonLength: number
  mode: 'resize-to-fit'
}
type PPTLineStyleImportField =
  | 'color'
  | 'dash'
  | 'width'
type PPTLineStyleImportSource = {
  fields: readonly PPTLineStyleImportField[]
  format: typeof PPT_LINE_STYLE_JSON_IMPORT_FORMAT
  jsonLength: number
  stroke: Partial<PPTStroke>
}
type PPTObjectTransformImportField =
  | 'h'
  | 'rotation'
  | 'w'
  | 'x'
  | 'y'
type PPTObjectTransformImportSource = {
  fields: readonly PPTObjectTransformImportField[]
  format: typeof PPT_OBJECT_TRANSFORM_JSON_IMPORT_FORMAT
  jsonLength: number
  transform: Partial<Pick<Bounds, 'h' | 'w' | 'x' | 'y'> & {
    rotation: number
  }>
}
type PPTCommentImportField =
  | 'body'
  | 'createdAt'
  | 'resolved'
  | 'thread'
type PPTCommentImportSource = {
  comment: {
    body?: string
    createdAt?: string
    resolved?: boolean
    thread?: readonly PPTCommentThreadMessage[]
  }
  fields: readonly PPTCommentImportField[]
  format: typeof PPT_COMMENT_JSON_IMPORT_FORMAT
  jsonLength: number
}
type PPTMediaJSONImportField =
  | 'title'
  | 'url'
type PPTMediaJSONImportSource = {
  fields: readonly PPTMediaJSONImportField[]
  format: typeof PPT_MEDIA_JSON_IMPORT_FORMAT
  jsonLength: number
  source: PPTMediaImportSource
}
type PPTElementsJSONImportSource = {
  format:
    | typeof PPT_ELEMENTS_JSON_IMPORT_FORMAT
    | typeof PPT_ELEMENTS_JSON_TEXT_IMPORT_FORMAT
  jsonLength: number
  objects: readonly PPTElement[]
  selectedObjectIds: readonly string[]
  sourceSlideId: string
}
type PPTDeckHTMLFallbackSlideSource = {
  htmlLength: number
  name: string
  sourceSlideId?: string
  svg: string
}
type PPTDeckHTMLFallbackSource = {
  htmlLength: number
  slides: PPTDeckHTMLFallbackSlideSource[]
  title: string
}
type PPTDeckHTMLImportEffect = {
  firstImportedSlideId: string
  format:
    | typeof PPT_DECK_HTML_IMPORT_FORMAT
    | typeof PPT_DECK_HTML_FALLBACK_IMPORT_FORMAT
  htmlLength: number
  importedSlideCount: number
  model: typeof PPT_DECK_HTML_IMPORT_MODEL
  sourceDeckId: string
  sourceSlideCount: number
  sourceTitle: string
}
type PPTDeckMarkdownOutlineImportEffect = {
  firstImportedSlideId: string
  format: typeof PPT_DECK_MARKDOWN_OUTLINE_IMPORT_FORMAT
  importedSlideCount: number
  model: typeof PPT_DECK_MARKDOWN_OUTLINE_IMPORT_MODEL
  sourceSlideCount: number
  sourceTitle: string
  textLength: number
}
type PPTDeckJSONImportEffect = {
  firstImportedSlideId: string
  format:
    | typeof PPT_DECK_JSON_IMPORT_FORMAT
    | typeof PPT_DECK_JSON_TEXT_IMPORT_FORMAT
  importedSlideCount: number
  jsonLength: number
  model: typeof PPT_DECK_JSON_IMPORT_MODEL
  sourceDeckId: string
  sourceSlideCount: number
  sourceTitle: string
}
type PPTSlideJSONImportEffect = {
  firstImportedSlideId: string
  format:
    | typeof PPT_SLIDE_JSON_IMPORT_FORMAT
    | typeof PPT_SLIDE_JSON_TEXT_IMPORT_FORMAT
  importedSlideCount: number
  jsonLength: number
  model: typeof PPT_SLIDE_JSON_IMPORT_MODEL
  sourceSlideCount: number
  sourceSlideIds: string
  sourceSlideNames: string
}
type PPTSlideNotesImportEffect = {
  format:
    | typeof PPT_SLIDE_NOTES_JSON_IMPORT_FORMAT
    | typeof PPT_SLIDE_NOTES_MARKDOWN_IMPORT_FORMAT
    | typeof PPT_SLIDE_NOTES_TEXT_IMPORT_FORMAT
  model: typeof PPT_SLIDE_NOTES_IMPORT_MODEL
  notesLength: number
  slideId: string
  textLength: number
}
type PPTSlideMetadataImportEffect = {
  backgroundColor: string
  commandIds: string
  fieldIds: string
  format: typeof PPT_SLIDE_METADATA_JSON_IMPORT_FORMAT
  jsonLength: number
  model: typeof PPT_SLIDE_METADATA_IMPORT_MODEL
  name: string
  notesLength: number
  slideId: string
}
type PPTSlideLayoutApplyHostCommandEffect =
  SlideEditLayoutApplyHostCommandEffect<string, string, string, string>
type PPTSlideLayoutImportEffect = {
  commandFields: string
  commandIds: string
  commandTypes: string
  fields: string
  format: typeof PPT_SLIDE_LAYOUT_JSON_IMPORT_FORMAT
  hiddenPlaceholderIds: string
  jsonLength: number
  layoutId: string
  model: typeof PPT_SLIDE_LAYOUT_IMPORT_MODEL
  placeholderCommandCount: number
  slideId: string
  themeId: string
}
type PPTSlideTransitionImportEffect = {
  advanceAfterMs: string
  advanceOnClick: string
  commandFields: string
  commandIds: string
  durationMs: string
  fields: string
  format: typeof PPT_SLIDE_TRANSITION_JSON_IMPORT_FORMAT
  jsonLength: number
  model: typeof PPT_SLIDE_TRANSITION_IMPORT_MODEL
  slideId: string
  type: string
}
type PPTObjectAnimationImportEffect = {
  commandFields: string
  commandIds: string
  delayMs: string
  durationMs: string
  fields: string
  format: typeof PPT_OBJECT_ANIMATION_JSON_IMPORT_FORMAT
  jsonLength: number
  model: typeof PPT_OBJECT_ANIMATION_IMPORT_MODEL
  objectIds: string
  order: string
  slideId: string
  trigger: string
  type: string
}
type PPTObjectStyleImportEffect = {
  categories: string
  commandId: string
  commandTargets: string
  commandType: string
  fields: string
  format: typeof PPT_OBJECT_STYLE_JSON_IMPORT_FORMAT
  jsonLength: number
  model: typeof PPT_OBJECT_STYLE_IMPORT_MODEL
  objectIds: string
  opacity: string
  shadowAngle: string
  shadowBlur: string
  shadowColor: string
  shadowDistance: string
  shadowEnabled: string
  shadowOpacity: string
}
type PPTObjectMetadataImportEffect = {
  altTextLength: number
  altTextPresent: string
  commandFields: string
  commandIds: string
  commandTypes: string
  fields: string
  format: typeof PPT_OBJECT_METADATA_JSON_IMPORT_FORMAT
  hyperlinkUrl: string
  jsonLength: number
  model: typeof PPT_OBJECT_METADATA_IMPORT_MODEL
  name: string
  objectIds: string
  slideId: string
}
type PPTObjectStateImportEffect = {
  commandIds: string
  commandTypes: string
  fields: string
  format: typeof PPT_OBJECT_STATE_JSON_IMPORT_FORMAT
  jsonLength: number
  locked: string
  lockTargets: string
  model: typeof PPT_OBJECT_STATE_IMPORT_MODEL
  objectIds: string
  slideId: string
  visible: string
  visibilityTargets: string
}
type PPTObjectLayerImportEffect = {
  commandId: string
  commandType: string
  fields: string
  format: typeof PPT_OBJECT_LAYER_JSON_IMPORT_FORMAT
  fromIndex: number
  jsonLength: number
  model: typeof PPT_OBJECT_LAYER_IMPORT_MODEL
  objectId: string
  position: string
  slideId: string
  toIndex: number
}
type PPTImageCropImportEffect = {
  commandFields: string
  commandIds: string
  commandTypes: string
  fields: string
  fit: string
  format: typeof PPT_IMAGE_CROP_JSON_IMPORT_FORMAT
  jsonLength: number
  model: typeof PPT_IMAGE_CROP_IMPORT_MODEL
  objectIds: string
  slideId: string
  x: string
  y: string
}
type PPTImageReplaceImportEffect = {
  altTextLength: number
  commandId: string
  commandType: string
  fields: string
  format: typeof PPT_IMAGE_REPLACE_JSON_IMPORT_FORMAT
  jsonLength: number
  mimeType: string
  model: typeof PPT_IMAGE_REPLACE_IMPORT_MODEL
  name: string
  naturalHeight: string
  naturalWidth: string
  objectId: string
  slideId: string
  srcPrefix: string
}
type PPTShapeStyleImportEffect = {
  categories: string
  commandId: string
  commandTargets: string
  commandType: string
  cornerRadius: string
  fields: string
  fillColor: string
  fillOpacity: string
  format: typeof PPT_SHAPE_STYLE_JSON_IMPORT_FORMAT
  jsonLength: number
  model: typeof PPT_SHAPE_STYLE_IMPORT_MODEL
  objectIds: string
  strokeColor: string
  strokeDash: string
  strokeWidth: string
}
type PPTTextStyleImportEffect = {
  categories: string
  color: string
  commandId: string
  commandTargets: string
  commandType: string
  fields: string
  fontFamily: string
  fontSize: string
  fontWeight: string
  format: typeof PPT_TEXT_STYLE_JSON_IMPORT_FORMAT
  jsonLength: number
  model: typeof PPT_TEXT_STYLE_IMPORT_MODEL
  objectIds: string
  paragraphAlign: string
  paragraphBullet: string
  paragraphLineHeight: string
  paragraphSpacingAfter: string
  paragraphSpacingBefore: string
  textInset: string
  verticalAlign: string
}
type PPTTextBodyImportEffect = {
  commandTargets: string
  format: typeof PPT_TEXT_BODY_JSON_IMPORT_FORMAT
  jsonLength: number
  mode: PPTTextBodyImportSource['mode']
  model: typeof PPT_TEXT_BODY_IMPORT_MODEL
  objectIds: string
  paragraphCount: number
  runCount: number
  textLength: number
}
type PPTTextAutoFitImportEffect = {
  commandHandles: string
  commandIds: string
  commandTargets: string
  commandTypes: string
  fields: string
  format: typeof PPT_TEXT_AUTOFIT_JSON_IMPORT_FORMAT
  jsonLength: number
  mode: string
  model: typeof PPT_TEXT_AUTOFIT_IMPORT_MODEL
  objectIds: string
}
type PPTLineStyleImportEffect = {
  categories: string
  commandId: string
  commandTargets: string
  commandType: string
  fields: string
  format: typeof PPT_LINE_STYLE_JSON_IMPORT_FORMAT
  jsonLength: number
  model: typeof PPT_LINE_STYLE_IMPORT_MODEL
  objectIds: string
  strokeColor: string
  strokeDash: string
  strokeWidth: string
}
type PPTObjectTransformImportEffect = {
  commandTargets: string
  fields: string
  format: typeof PPT_OBJECT_TRANSFORM_JSON_IMPORT_FORMAT
  h: string
  jsonLength: number
  model: typeof PPT_OBJECT_TRANSFORM_IMPORT_MODEL
  objectIds: string
  rotation: string
  w: string
  x: string
  y: string
}
type PPTCommentImportEffect = {
  bodyLength: number
  commandTargets: string
  createdAt: string
  fields: string
  format: typeof PPT_COMMENT_JSON_IMPORT_FORMAT
  jsonLength: number
  messageCount: number
  model: typeof PPT_COMMENT_IMPORT_MODEL
  objectIds: string
  resolved: string
}
type PPTMediaJSONImportEffect = {
  fields: string
  format: typeof PPT_MEDIA_JSON_IMPORT_FORMAT
  importerId: string
  jsonLength: number
  model: typeof PPT_MEDIA_JSON_IMPORT_MODEL
  objectId: string
  title: string
  url: string
}
type PPTElementsJSONImportEffect = {
  format:
    | typeof PPT_ELEMENTS_JSON_IMPORT_FORMAT
    | typeof PPT_ELEMENTS_JSON_TEXT_IMPORT_FORMAT
  importedObjectCount: number
  jsonLength: number
  model: typeof PPT_ELEMENTS_JSON_IMPORT_MODEL
  selectedObjectIds: readonly string[]
  sourceSlideId: string
}
type PPTSlideSVGClipboardEffect = {
  jsonMimeType: typeof PPT_SLIDE_SVG_CLIPBOARD_JSON_MIME_TYPE
  model: typeof PPT_SLIDE_SVG_CLIPBOARD_MODEL
  sourceSlideId: string
  svgLength: number
  writeMode?: PPTRichClipboardWriteMode
}
type PPTSelectionSVGClipboardEffect = {
  jsonMimeType: typeof PPT_SELECTION_SVG_CLIPBOARD_JSON_MIME_TYPE
  model: typeof PPT_SELECTION_SVG_CLIPBOARD_MODEL
  selectedObjectIds: readonly string[]
  sourceSlideId: string
  svgLength: number
  writeMode?: PPTRichClipboardWriteMode
}
type PPTTableClipboardEffect = {
  columnCount: number
  htmlLength: number
  jsonMimeType: typeof PPT_TABLE_CLIPBOARD_JSON_MIME_TYPE
  model: typeof PPT_TABLE_CLIPBOARD_MODEL
  objectId: string
  plainTextLength: number
  rowCount: number
  sourceSlideId: string
  writeMode?: PPTRichClipboardWriteMode
}
type PPTTableRowsImportSource = {
  format: typeof PPT_TABLE_ROWS_JSON_IMPORT_FORMAT
  jsonLength: number
  rows: readonly (readonly string[])[]
}
type PPTTableRowsImportEffect = {
  columnCount: number
  commandTargets: string
  format: typeof PPT_TABLE_ROWS_JSON_IMPORT_FORMAT
  jsonLength: number
  model: typeof PPT_TABLE_ROWS_IMPORT_MODEL
  objectIds: string
  rowCount: number
}
type PPTStyleClipboardCategory =
  | 'object'
  | 'paragraph'
  | 'shape'
  | 'stroke'
  | 'text'
type PPTStyleClipboardParagraph = Pick<
  PPTParagraph,
  'align' | 'bullet' | 'lineHeight' | 'spacingAfter' | 'spacingBefore'
>
type PPTStyleClipboard = {
  categories: readonly PPTStyleClipboardCategory[]
  object: {
    opacity: number
    shadow: PPTElementShadow | null
  }
  paragraph?: PPTStyleClipboardParagraph
  shape?: {
    cornerRadius?: number
    fill: PPTFill
    stroke?: PPTStroke
  }
  sourceId: string
  sourceKind: PPTElement['kind']
  stroke?: PPTStroke
  text?: PPTTextStyle
  type: 'slide-style-clipboard'
}
type PPTStyleClipboardPackageCategory = SlideEditStyleClipboardBuiltInCategoryId
type PPTStyleClipboardPackageStyle =
  | {
    categoryId: 'line-style'
    value: PPTStroke
  }
  | {
    categoryId: 'object-effect'
    value: PPTStyleClipboard['object']
  }
  | {
    categoryId: 'shape-fill'
    value: Pick<NonNullable<PPTStyleClipboard['shape']>, 'cornerRadius' | 'fill'>
  }
  | {
    categoryId: 'shape-stroke'
    value: PPTStroke
  }
  | {
    categoryId: 'text-style'
    value: {
      paragraph?: PPTStyleClipboardParagraph
      text?: PPTTextStyle
    }
  }
type PPTStyleClipboardDescriptor = SlideEditStyleClipboardDescriptor<
  string,
  string,
  PPTElement['kind'],
  PPTStyleClipboardPackageCategory,
  unknown
>
type PPTStyleClipboardHostCommandEffect = SlideEditStyleClipboardHostCommandEffect<
  string,
  string,
  | SlideEditStyleClipboardCopyFormattingCommand<
    string,
    string,
    PPTElement['kind'],
    PPTStyleClipboardPackageCategory,
    unknown
  >
  | SlideEditStyleClipboardPasteFormattingCommand<
    string,
    string,
    PPTElement['kind'],
    PPTStyleClipboardPackageCategory,
    unknown
  >
>
type PPTColorSwatchChannel =
  | 'line-stroke'
  | 'shape-fill'
  | 'shape-stroke'
  | 'text-color'
type PPTColorSwatchPackageChannel = SlideEditColorSwatchBuiltInChannelId
type PPTColorSwatchDescriptor = SlideEditColorSwatchPaletteDescriptor<
  string,
  string,
  PPTColorSwatchPackageChannel,
  string
>
type PPTColorSwatchHostCommandEffect = SlideEditColorSwatchHostCommandEffect<
  string,
  string,
  PPTColorSwatchPackageChannel,
  string
>
type PPTColorSwatchSelection = SlideEditColorSwatchSelection<string>
const PPT_COLOR_SWATCH_CHANNEL_MAP = {
  'line-stroke': 'line-stroke',
  'shape-fill': 'fill',
  'shape-stroke': 'stroke',
  'text-color': 'text',
} as const satisfies Record<PPTColorSwatchChannel, PPTColorSwatchPackageChannel>
type PPTSurfaceCommand =
  | 'alignBottom'
  | 'alignCenter'
  | 'alignLeft'
  | 'alignMiddle'
  | 'alignRight'
  | 'alignTop'
  | 'bringForward'
  | 'bringToFront'
  | 'copyFormatting'
  | 'delete'
  | 'distributeHorizontal'
  | 'distributeVertical'
  | 'duplicate'
  | 'flipHorizontal'
  | 'flipVertical'
  | 'group'
  | 'lockSelection'
  | 'pasteFormatting'
  | 'selectSameType'
  | 'sendBackward'
  | 'sendToBack'
  | 'tidySelection'
  | 'ungroup'
  | 'unlockAll'
type PPTAlignmentPopoverCommand = Extract<
  PPTSurfaceCommand,
  | 'alignBottom'
  | 'alignCenter'
  | 'alignLeft'
  | 'alignMiddle'
  | 'alignRight'
  | 'alignTop'
  | 'distributeHorizontal'
  | 'distributeVertical'
>
type PPTCommandAvailability = ReturnType<typeof getPPTCanvasCommandAvailability> & {
  copyFormatting: boolean
  flipSelection: boolean
  pasteFormatting: boolean
  selectSameType: boolean
  tidySelection: boolean
}
type PPTCommandAvailabilityKey = keyof PPTCommandAvailability
type PPTSurfaceCommandDescriptor = {
  availability: PPTCommandAvailabilityKey
  command: PPTSurfaceCommand
  dataCommand: string
  label: string
  surfaces: readonly PPTCommandSurface[]
  title: string
}
type PPTSurfaceCommandGroup = {
  commands: readonly PPTSurfaceCommandDescriptor[]
  id: string
}
type PPTSurfaceCommandView = PPTSurfaceCommandDescriptor & {
  disabled: boolean
}
type PPTSurfaceCommandViewGroup = {
  commands: PPTSurfaceCommandView[]
  id: string
}
type PPTCommandPaletteItem = PPTCommandPaletteItemBase
type PPTShortcutHelpItem = {
  id: string
  section: string
  shortcut: string
  title: string
}
type PPTShortcutHelpSectionGroup = {
  items: PPTShortcutHelpItem[]
  section: string
}
type PPTSlideMetadataOrientation = 'landscape' | 'portrait'
type PPTSlideMetadataBackgroundDescriptor =
  | {
      kind: 'none'
    }
  | {
      color: string
      kind: 'solid-color'
      tokenId?: string
    }
type PPTSlideMetadataSizeDescriptor = {
  h: number
  w: number
}
type PPTSlideMetadataReadModel = {
  background: PPTSlideMetadataBackgroundDescriptor
  name: string
  notes: string
  orientation: PPTSlideMetadataOrientation
  size: PPTSlideMetadataSizeDescriptor
  slideId: string
}
type PPTSlideMetadataFieldId =
  | 'background'
  | 'name'
  | 'notes'
  | 'orientation'
  | 'size'
type PPTSlideMetadataCommandId =
  | 'update-slide-background'
  | 'update-slide-name'
  | 'update-slide-notes'
  | 'update-slide-orientation'
  | 'update-slide-size'
type PPTSlideMetadataFieldControl =
  | 'background-control'
  | 'multiline-text'
  | 'orientation-control'
  | 'size-control'
  | 'text'
type PPTSlideMetadataFieldDescriptor = {
  commandId: PPTSlideMetadataCommandId
  control: PPTSlideMetadataFieldControl
  id: PPTSlideMetadataFieldId
  isEditable: boolean
  isOptional: boolean
  requiredAdapterSlot: 'command-effect'
}
type PPTInspectorSurfaceId =
  | 'none'
  | 'object-selection-inspector'
  | 'slide-metadata-inspector'
type PPTInspectorTabId =
  | 'selection'
  | 'slide'
const PPT_INSPECTOR_TABS = [
  {
    id: 'slide',
    label: 'Slide',
    panelId: 'ppt-inspector-panel-slide',
    tabId: 'ppt-inspector-tab-slide',
  },
  {
    id: 'selection',
    label: 'Selection',
    panelId: 'ppt-inspector-panel-selection',
    tabId: 'ppt-inspector-tab-selection',
  },
] as const satisfies readonly {
  id: PPTInspectorTabId
  label: string
  panelId: string
  tabId: string
}[]
type PPTSlideMetadataInspectorDescriptor = {
  activeSlide: {
    index: number | null
    slideCount: number
    slideId: string
  }
  fields: readonly PPTSlideMetadataFieldDescriptor[]
  metadata: PPTSlideMetadataReadModel
  surface: 'slide-metadata-inspector'
}
type PPTSlideMetadataUpdateCommand =
  | {
      fieldId: 'background'
      id: 'update-slide-background'
      slideId: string
      value: PPTSlideMetadataBackgroundDescriptor
    }
  | {
      fieldId: 'name'
      id: 'update-slide-name'
      slideId: string
      value: string
    }
  | {
      fieldId: 'notes'
      id: 'update-slide-notes'
      slideId: string
      value: string
    }
  | {
      fieldId: 'orientation'
      id: 'update-slide-orientation'
      slideId: string
      value: PPTSlideMetadataOrientation
    }
  | {
      fieldId: 'size'
      id: 'update-slide-size'
      slideId: string
      value: PPTSlideMetadataSizeDescriptor
    }
type PPTSlideMetadataHostCommandEffect = {
  payload: PPTSlideMetadataUpdateCommand
  selection: {
    objectIds: readonly string[]
    slideId: string
  }
  type: 'slide-command-effect'
}
type PPTLayoutPlaceholderVisibilityCommand = {
  id: 'update-placeholder-visibility'
  isVisible: boolean
  placeholderId: string
  slideId: string
}
type PPTLayoutPlaceholderVisibilityHostCommandEffect = {
  payload: PPTLayoutPlaceholderVisibilityCommand
  selection: {
    objectIds: readonly string[]
    slideId: string
  }
  type: 'slide-command-effect'
}
type PPTSlideTransitionType = PPTSlideTransition['type']
type PPTSlideTransitionUpdateField =
  | 'advanceAfterMs'
  | 'advanceOnClick'
  | 'durationMs'
  | 'type'
type PPTSlideTransitionDescriptor =
  SlideEditSlideTransitionDescriptor<string, PPTSlideTransitionType>
type PPTSlideTransitionHostCommandEffect =
  SlideEditTransitionHostCommandEffect<string, PPTSlideTransitionType>
type PPTSlideTransitionUpdateCommand =
  SlideEditTransitionUpdateCommand<string, PPTSlideTransitionType>
type PPTElementAnimationType = PPTElementAnimation['type']
type PPTElementAnimationTrigger = PPTElementAnimation['trigger']
type PPTElementAnimationUpdateField =
  | 'delayMs'
  | 'durationMs'
  | 'order'
  | 'trigger'
  | 'type'
type PPTElementAnimationCSSStyle = ReturnType<
  typeof getSlideEditObjectAnimationCSSStyle
>
type PPTElementShadowUpdateField =
  | 'angle'
  | 'blur'
  | 'color'
  | 'distance'
  | 'enabled'
  | 'opacity'
type PPTParagraphSpacingField =
  | 'lineHeight'
  | 'spacingAfter'
  | 'spacingBefore'
type PPTParagraphCSSStyle = ReturnType<
  typeof getSlideEditTextParagraphSpacingCSSStyle
>
type PPTTextInset = NonNullable<PPTTextStyle['textInset']>
type PPTTextInsetField = keyof PPTTextInset
type PPTTextVerticalAlign = NonNullable<PPTTextStyle['verticalAlign']>
type PPTLayerPaneRowDescriptor = SlideEditLayerPaneRowDescriptor<string, string, string>
type PPTLayerPaneDescriptor = SlideEditLayerPaneDescriptor<string, string, string>
type PPTLayerPaneCommandDescriptor = SlideEditLayerPaneCommandDescriptor
type PPTLayerPaneHostCommandEffect = SlideEditLayerPaneHostCommandEffect<string, string>
type PPTLayerPaneIntent = SlideEditLayerPaneIntent<string>
type PPTLayerPaneKeyboardIntent = SlideEditLayerPaneKeyboardIntent<string>
type PPTObjectVisibilityDescriptor = SlideEditObjectVisibilityDescriptor<string, string, string>
type PPTObjectVisibilityHostCommandEffect = SlideEditObjectVisibilityHostCommandEffect<string, string>
type PPTCommentThreadCommand = {
  body: string
  id: 'add-comment-reply'
  messageCount: number
  objectId: string
  slideId: string
}
type PPTCommentThreadHostCommandEffect = {
  payload: PPTCommentThreadCommand
  selection: {
    objectIds: readonly string[]
    slideId: string
  }
  type: 'slide-command-effect'
}
type PPTInlineEditHistoryDirection = 'redo' | 'undo'
type PPTInlineEditEffect = {
  elementId: string
  historyDirection?: PPTInlineEditHistoryDirection
  inputType?: string
  lineBreak?: boolean
  model: typeof PPT_INLINE_EDIT_DOM_MODEL
  pasteText?: string
}
type PPTResizeHandleClickMemoryEffect = {
  handle: ResizeHandle
  id: string
  isDoubleClick: boolean
  model: typeof PPT_POINTER_CLICK_MEMORY_MODEL
  point: Point
}
type PPTLayerPaneGroupState = {
  collapsedGroupIds: readonly string[]
  focusedObjectId: string | null
  rangeAnchorObjectId: string | null
}
type PPTLayerPaneRenameState = {
  objectId: string
  value: string
}
type PPTLayerPaneDropPlacement = Exclude<SlideEditLayerPaneDropPlacement, 'none'>
type PPTLayerPaneDragState = {
  dropPlacement?: PPTLayerPaneDropPlacement
  dropTargetObjectId?: string
  dropToIndex?: number
  objectId: string
}
type PPTMinimapSize = PPTMinimapSizeBase
type PPTMinimapItemBounds = PPTMinimapItemBoundsBase
type PPTMinimapReadModel = PPTMinimapReadModelBase
type PPTContextMenuState = {
  x: number
  y: number
}
type PPTSlideDropPlacement = 'after' | 'before'
type PPTSlideDragState = {
  draggingSlideId: string
  dropPlacement?: PPTSlideDropPlacement
  dropTargetSlideId?: string
}
type PPTSelectionCommandAnchor = PPTCanvasFloatingAnchor
type PPTTextQuickFormatState = {
  align: NonNullable<PPTParagraph['align']>
  bullet: boolean
  color: string
  fontSize: number
  isBold: boolean
  isItalic: boolean
  isUnderline: boolean
  numbered: boolean
}
type PPTShapeQuickMenuState = {
  elementId: string
  shape: PPTShapeKind
}

const PPT_COMMAND_SURFACE_GROUPS: readonly PPTSurfaceCommandGroup[] = [{
  commands: [{
    availability: 'duplicate',
    command: 'duplicate',
    dataCommand: 'duplicate',
    label: 'Duplicate',
    surfaces: ['context-menu', 'selection-floating-bar'],
    title: PPT_COMMAND_AFFORDANCES.duplicate.title,
  }, {
    availability: 'copyFormatting',
    command: 'copyFormatting',
    dataCommand: 'copy-formatting',
    label: 'Copy formatting',
    surfaces: ['context-menu', 'selection-floating-bar'],
    title: 'Copy formatting',
  }, {
    availability: 'pasteFormatting',
    command: 'pasteFormatting',
    dataCommand: 'paste-formatting',
    label: 'Paste formatting',
    surfaces: ['context-menu', 'selection-floating-bar'],
    title: 'Paste formatting',
  }, {
    availability: 'selectSameType',
    command: 'selectSameType',
    dataCommand: 'select-same-type',
    label: 'Select same type',
    surfaces: ['context-menu', 'selection-floating-bar'],
    title: 'Select same type',
  }, {
    availability: 'tidySelection',
    command: 'tidySelection',
    dataCommand: 'tidy-selection',
    label: 'Tidy selection',
    surfaces: ['context-menu', 'selection-floating-bar'],
    title: 'Tidy selection',
  }, {
    availability: 'flipSelection',
    command: 'flipHorizontal',
    dataCommand: 'flip-horizontal',
    label: 'Flip horizontal',
    surfaces: ['context-menu'],
    title: 'Flip horizontal',
  }, {
    availability: 'flipSelection',
    command: 'flipVertical',
    dataCommand: 'flip-vertical',
    label: 'Flip vertical',
    surfaces: ['context-menu'],
    title: 'Flip vertical',
  }, {
    availability: 'delete',
    command: 'delete',
    dataCommand: 'delete',
    label: 'Delete',
    surfaces: ['context-menu', 'selection-floating-bar'],
    title: PPT_COMMAND_AFFORDANCES.delete.title,
  }],
  id: 'edit',
}, {
  commands: [{
    availability: 'alignLeft',
    command: 'alignLeft',
    dataCommand: 'align-left',
    label: 'Align left',
    surfaces: ['context-menu'],
    title: PPT_COMMAND_AFFORDANCES.alignLeft.title,
  }, {
    availability: 'alignCenter',
    command: 'alignCenter',
    dataCommand: 'align-center-x',
    label: 'Align center',
    surfaces: ['context-menu'],
    title: PPT_COMMAND_AFFORDANCES.alignCenter.title,
  }, {
    availability: 'alignRight',
    command: 'alignRight',
    dataCommand: 'align-right',
    label: 'Align right',
    surfaces: ['context-menu'],
    title: PPT_COMMAND_AFFORDANCES.alignRight.title,
  }],
  id: 'align',
}, {
  commands: [{
    availability: 'bringForward',
    command: 'bringForward',
    dataCommand: 'bring-forward',
    label: 'Bring forward',
    surfaces: ['context-menu'],
    title: PPT_COMMAND_AFFORDANCES.bringForward.title,
  }, {
    availability: 'bringToFront',
    command: 'bringToFront',
    dataCommand: 'bring-to-front',
    label: 'Bring to front',
    surfaces: ['context-menu', 'selection-floating-bar'],
    title: PPT_COMMAND_AFFORDANCES.bringToFront.title,
  }, {
    availability: 'sendBackward',
    command: 'sendBackward',
    dataCommand: 'send-backward',
    label: 'Send backward',
    surfaces: ['context-menu'],
    title: PPT_COMMAND_AFFORDANCES.sendBackward.title,
  }, {
    availability: 'sendToBack',
    command: 'sendToBack',
    dataCommand: 'send-to-back',
    label: 'Send to back',
    surfaces: ['context-menu', 'selection-floating-bar'],
    title: PPT_COMMAND_AFFORDANCES.sendToBack.title,
  }],
  id: 'order',
}, {
  commands: [{
    availability: 'group',
    command: 'group',
    dataCommand: 'group',
    label: 'Group',
    surfaces: ['context-menu', 'selection-floating-bar'],
    title: PPT_COMMAND_AFFORDANCES.group.title,
  }, {
    availability: 'ungroup',
    command: 'ungroup',
    dataCommand: 'ungroup',
    label: 'Ungroup',
    surfaces: ['context-menu', 'selection-floating-bar'],
    title: PPT_COMMAND_AFFORDANCES.ungroup.title,
  }],
  id: 'group',
}, {
  commands: [{
    availability: 'lockSelection',
    command: 'lockSelection',
    dataCommand: 'lock-selection',
    label: 'Lock',
    surfaces: ['context-menu', 'selection-floating-bar'],
    title: PPT_COMMAND_AFFORDANCES.lockSelection.title,
  }, {
    availability: 'unlockAll',
    command: 'unlockAll',
    dataCommand: 'unlock-all',
    label: 'Unlock all',
    surfaces: ['context-menu'],
    title: PPT_COMMAND_AFFORDANCES.unlockAll.title,
  }],
  id: 'lock',
}]
const PPT_ALIGNMENT_POPOVER_COMMANDS = [{
  command: 'alignLeft',
  dataCommand: 'align-left',
  label: 'Align left',
}, {
  command: 'alignCenter',
  dataCommand: 'align-center-x',
  label: 'Align center',
}, {
  command: 'alignRight',
  dataCommand: 'align-right',
  label: 'Align right',
}, {
  command: 'alignTop',
  dataCommand: 'align-top',
  label: 'Align top',
}, {
  command: 'alignMiddle',
  dataCommand: 'align-middle',
  label: 'Align middle',
}, {
  command: 'alignBottom',
  dataCommand: 'align-bottom',
  label: 'Align bottom',
}, {
  command: 'distributeHorizontal',
  dataCommand: 'distribute-horizontal',
  label: 'Distribute horizontal',
}, {
  command: 'distributeVertical',
  dataCommand: 'distribute-vertical',
  label: 'Distribute vertical',
}] as const satisfies readonly {
  command: PPTAlignmentPopoverCommand
  dataCommand: string
  label: string
}[]
const PPT_SHAPE_MENU_OPTIONS = [{
  label: 'Rectangle',
  shape: 'rect',
}, {
  label: 'Oval',
  shape: 'ellipse',
}, {
  label: 'Diamond',
  shape: 'diamond',
}] as const satisfies readonly {
  label: string
  shape: PPTShapeKind
}[]

const PPT_LINE_CONNECTION_DISTANCE = 36
const PPT_TIDY_GAP = 24
const PPT_COMMENT_DEFAULT_BODY = 'Comment'
const PPT_COMMENT_DEFAULT_AUTHOR = 'You'
const PPT_COMMENT_DEFAULT_CREATED_AT = 'Just now'
const PPT_COMMENT_BODY_MAX_LENGTH = 240
const PPT_COMMENT_REPLY_MAX_LENGTH = 240
const PPT_COMMENT_BOUNDS = {
  h: 132,
  w: 260,
}
const PPT_DEFAULT_TEXT_BOUNDS = {
  h: 76,
  w: 360,
}
const PPT_TEXT_FONT_SIZE_MIN = 8
const PPT_TEXT_FONT_SIZE_MAX = 120
const PPT_TEXT_FONT_SIZE_STEP = 2
const PPT_DEFAULT_TEXT_FONT_FAMILY = 'Inter'
const PPT_TEXT_FONT_FAMILY_OPTIONS = Object.freeze([
  { css: 'Inter, ui-sans-serif, system-ui, sans-serif', label: 'Inter', value: 'Inter' },
  { css: 'Arial, Helvetica, sans-serif', label: 'Arial', value: 'Arial' },
  { css: 'Georgia, serif', label: 'Georgia', value: 'Georgia' },
  { css: '"Courier New", monospace', label: 'Courier New', value: 'Courier New' },
] as const)
const PPT_DEFAULT_TEXT_VERTICAL_ALIGN: PPTTextVerticalAlign = 'top'
const PPT_TEXT_INSET_MIN = 0
const PPT_TEXT_INSET_MAX = 120
const PPT_TEXT_INSET_STEP = 2
const PPT_DEFAULT_TEXT_BOX_INSET = Object.freeze({
  bottom: 0,
  left: 0,
  right: 0,
  top: 0,
} as const satisfies PPTTextInset)
const PPT_DEFAULT_SHAPE_TEXT_INSET = Object.freeze({
  bottom: 18,
  left: 18,
  right: 18,
  top: 18,
} as const satisfies PPTTextInset)
const PPT_TEXT_AUTOFIT: PPTTextAutoFit = 'resizeShapeToFitText'
const PPT_TEXT_OVERFLOW_EPSILON = 1
const PPT_ELEMENT_OPACITY_MIN = 0
const PPT_ELEMENT_OPACITY_MAX = 1
const PPT_ELEMENT_OPACITY_STEP = 0.05
const PPT_DEFAULT_ELEMENT_SHADOW = Object.freeze({
  angle: 45,
  blur: 14,
  color: '#000000',
  distance: 8,
  opacity: 0.22,
} as const satisfies PPTElementShadow)
const PPT_ELEMENT_SHADOW_ANGLE_MIN = 0
const PPT_ELEMENT_SHADOW_ANGLE_MAX = 359
const PPT_ELEMENT_SHADOW_BLUR_MAX = 80
const PPT_ELEMENT_SHADOW_DISTANCE_MAX = 120
const PPT_ELEMENT_SHADOW_OPACITY_MIN = 0
const PPT_ELEMENT_SHADOW_OPACITY_MAX = 1
const PPT_ELEMENT_SHADOW_OPACITY_STEP = 0.05
const PPT_ALT_TEXT_MAX_LENGTH = 1000
const PPT_FILL_OPACITY_MIN = 0
const PPT_FILL_OPACITY_MAX = 1
const PPT_FILL_OPACITY_STEP = 0.05
const PPT_SHAPE_CORNER_RADIUS_DEFAULT = 24
const PPT_SHAPE_CORNER_RADIUS_MIN = 0
const PPT_SHAPE_CORNER_RADIUS_MAX = 120
const PPT_SHAPE_CORNER_RADIUS_STEP = 1
const PPT_STICKY_BOUNDS = Object.freeze({
  h: 168,
  w: 220,
} as const)
const PPT_SECTION_BOUNDS = Object.freeze({
  h: 240,
  w: 420,
} as const)
const PPT_HYPERLINK_URL_MAX_LENGTH = 2048
const PPT_STROKE_DASH_OPTIONS = Object.freeze(
  SLIDE_EDIT_OBJECT_STROKE_LINE_STYLE_OPTIONS.map((option) => ({
    id: option.id,
    label: option.label,
    value: option.id,
  })),
) as readonly {
  id: PPTStrokeDash
  label: string
  value: PPTStrokeDash
}[]
const PPT_SLIDE_TRANSITION_TYPES = Object.freeze(
  SLIDE_EDIT_TRANSITION_TYPES.map((option) => option.id),
) as readonly PPTSlideTransitionType[]
const PPT_DEFAULT_SLIDE_TRANSITION = Object.freeze({
  advanceAfterMs: null,
  advanceOnClick: SLIDE_EDIT_DEFAULT_TRANSITION.advance.onClick,
  durationMs: SLIDE_EDIT_DEFAULT_TRANSITION.durationMs,
  type: SLIDE_EDIT_DEFAULT_TRANSITION.type as PPTSlideTransitionType,
} as const satisfies PPTSlideTransition)
const PPT_SLIDE_TRANSITION_DURATION_MAX = SLIDE_EDIT_TRANSITION_TIMING_LIMITS.maxDurationMs
const PPT_SLIDE_TRANSITION_ADVANCE_AFTER_MAX = SLIDE_EDIT_TRANSITION_TIMING_LIMITS.maxAdvanceAfterMs
const PPT_ELEMENT_ANIMATION_TYPES = Object.freeze([
  'none',
  'fadeIn',
  'flyIn',
] as const satisfies readonly PPTElementAnimationType[])
const PPT_ELEMENT_ANIMATION_TRIGGERS = Object.freeze([
  'onClick',
  'withPrevious',
] as const satisfies readonly PPTElementAnimationTrigger[])
const PPT_DEFAULT_ELEMENT_ANIMATION = Object.freeze({
  delayMs: 0,
  durationMs: 400,
  order: 1,
  trigger: 'onClick',
  type: 'none',
} as const satisfies PPTElementAnimation)
const PPT_PARAGRAPH_LINE_HEIGHT_DEFAULT = 1.14
const PPT_PARAGRAPH_LINE_HEIGHT_MIN = 0.8
const PPT_PARAGRAPH_LINE_HEIGHT_MAX = 3
const PPT_PARAGRAPH_SPACING_MAX = 240
const PPT_SHORTCUT_HELP_SHORTCUT = 'Shift+/'
const PPT_SHORTCUT_HELP_SECTION_ORDER = [
  'Create',
  'Edit',
  'Arrange',
  'Format',
  'Slides',
  'View',
  'Export',
  'System',
]
const PPT_SLIDE_METADATA_FIELDS = Object.freeze([
  {
    commandId: 'update-slide-name',
    control: 'text',
    id: 'name',
    isEditable: true,
    isOptional: false,
    requiredAdapterSlot: 'command-effect',
  },
  {
    commandId: 'update-slide-background',
    control: 'background-control',
    id: 'background',
    isEditable: true,
    isOptional: false,
    requiredAdapterSlot: 'command-effect',
  },
  {
    commandId: 'update-slide-notes',
    control: 'multiline-text',
    id: 'notes',
    isEditable: true,
    isOptional: false,
    requiredAdapterSlot: 'command-effect',
  },
  {
    commandId: 'update-slide-size',
    control: 'size-control',
    id: 'size',
    isEditable: false,
    isOptional: true,
    requiredAdapterSlot: 'command-effect',
  },
  {
    commandId: 'update-slide-orientation',
    control: 'orientation-control',
    id: 'orientation',
    isEditable: false,
    isOptional: true,
    requiredAdapterSlot: 'command-effect',
  },
] as const satisfies readonly PPTSlideMetadataFieldDescriptor[])
const PPT_LAYER_PANE_COMMANDS = SLIDE_EDIT_LAYER_PANE_COMMANDS satisfies readonly PPTLayerPaneCommandDescriptor[]
const PPT_MINIMAP_SIZE: PPTMinimapSize = {
  h: 112,
  w: 176,
}
const PPT_MINIMAP_SLIDE_FRAME_ID = 'ppt-slide-frame'
const PPT_ERASER_POINT_DISTANCE = 4
const PPT_LAYER_PANE_GROUP_ROW_PREFIX = 'ppt-layer-group:'

type LineCreationMode = 'arrow' | 'line'
type PPTFlipAxis = 'horizontal' | 'vertical'
type PPTFreeformTool = 'highlight' | 'marker' | 'pen'
type PPTCreationTool =
  | {
      kind: 'shape'
      shape: PPTShapeKind
    }
  | {
      kind: 'text'
    }
  | {
      kind: 'sticky'
    }
  | {
      kind: 'section'
    }
  | {
      kind: 'comment'
    }
  | {
      kind: 'freeform'
      tool: PPTFreeformTool
    }
type PPTFindMatch = {
  elementId: string
  elementIndex: number
  end: number
  slideId: string
  slideIndex: number
  start: number
}

type PPTTextRunStyle = Omit<PPTRun, 'text'>
type PPTTextToken = {
  align?: PPTParagraph['align']
  bullet?: PPTParagraph['bullet']
  char: string
  lineHeight?: PPTParagraph['lineHeight']
  runStyle: PPTTextRunStyle
  spacingAfter?: PPTParagraph['spacingAfter']
  spacingBefore?: PPTParagraph['spacingBefore']
}

type Interaction =
  | {
      bounds: Bounds
      historyDeck?: PPTDeck
      kind: 'move'
      selection: string[]
      slideId: string
      snapGuides: PPTCanvasSnapGuides
      startDeck: PPTDeck
      startPoint: Point
    }
  | {
      bounds: Bounds
      handle: ResizeHandle
      kind: 'resize'
      selection: string[]
      slideId: string
      startDeck: PPTDeck
    }
  | {
      bounds: Bounds
      center: Point
      kind: 'rotate'
      selection: string[]
      slideId: string
      startAngle: number
      startDeck: PPTDeck
      startRotations: Array<{
        elementId: string
        rotation: number
      }>
    }
  | {
      elementId: string
      kind: 'element-create'
      slideId: string
      startDeck: PPTDeck
      startPoint: Point
      tool: PPTCreationTool
    }
  | {
      endpoint: 'end' | 'start'
      kind: 'line-endpoint'
      lineId: string
      slideId: string
      startDeck: PPTDeck
    }
  | {
      endMarker: PPTLineMarker
      kind: 'line-create'
      lineId: string
      slideId: string
      startDeck: PPTDeck
      startPoint: Point
    }
  | {
      elementId: string
      kind: 'freeform-create'
      points: Point[]
      slideId: string
      startDeck: PPTDeck
    }
  | {
      kind: 'line-route'
      lineId: string
      slideId: string
      startDeck: PPTDeck
    }
  | {
      additive: boolean
      baseSelection: string[]
      currentPoint: Point
      kind: 'marquee'
      slideId: string
      startPoint: Point
    }
  | {
      kind: 'pan'
    } & PPTPointerPanInteraction
  | {
      kind: 'laser'
    } & PPTPointerLaserInteraction
  | {
      currentPoint: Point
      erasedIds: string[]
      kind: 'erase'
      points: Point[]
      slideId: string
      startDeck: PPTDeck
      startPoint: Point
    }

function App() {
  const [deck, setDeck] = useState(SAMPLE_PPT_DECK)
  const [activeSlideId, setActiveSlideId] = useState(deck.slides[0].id)
  const [selection, setSelection] = useState<string[]>(['s1-title'])
  const [viewport, setViewport] = useState<Viewport>({ scale: 1, x: 0, y: 0 })
  const [editingId, setEditingId] = useState<string | null>(null)
  const [hoveredId, setHoveredId] = useState<string | null>(null)
  const [alignmentPreviewCommand, setAlignmentPreviewCommand] =
    useState<PPTAlignmentPopoverCommand | null>(null)
  const [interaction, setInteraction] = useState<Interaction | null>(null)
  const [clipboard, setClipboard] = useState<PPTClipboard | null>(null)
  const [slideClipboard, setSlideClipboard] =
    useState<PPTSlideClipboardPayload | null>(null)
  const [styleClipboard, setStyleClipboard] = useState<PPTStyleClipboard | null>(null)
  const [lastClipboardPasteEffect, setLastClipboardPasteEffect] = useState<PPTClipboardPasteHostCommandEffect | null>(null)
  const [lastClipboardPastePositionEffect, setLastClipboardPastePositionEffect] = useState<PPTClipboardPastePositionEffect | null>(null)
  const [lastRichClipboardEffect, setLastRichClipboardEffect] = useState<PPTRichClipboardEffect | null>(null)
  const [lastSlideClipboardEffect, setLastSlideClipboardEffect] =
    useState<PPTSlideClipboardEffect | null>(null)
  const [lastHTMLClipboardEffect, setLastHTMLClipboardEffect] = useState<PPTHTMLClipboardEffect | null>(null)
  const [lastDeckHTMLImportEffect, setLastDeckHTMLImportEffect] =
    useState<PPTDeckHTMLImportEffect | null>(null)
  const [
    lastDeckMarkdownOutlineImportEffect,
    setLastDeckMarkdownOutlineImportEffect,
  ] = useState<PPTDeckMarkdownOutlineImportEffect | null>(null)
  const [lastDeckJSONImportEffect, setLastDeckJSONImportEffect] =
    useState<PPTDeckJSONImportEffect | null>(null)
  const [lastSlideJSONImportEffect, setLastSlideJSONImportEffect] =
    useState<PPTSlideJSONImportEffect | null>(null)
  const [lastSlideNotesImportEffect, setLastSlideNotesImportEffect] =
    useState<PPTSlideNotesImportEffect | null>(null)
  const [lastSlideMetadataImportEffect, setLastSlideMetadataImportEffect] =
    useState<PPTSlideMetadataImportEffect | null>(null)
  const [lastSlideLayoutImportEffect, setLastSlideLayoutImportEffect] =
    useState<PPTSlideLayoutImportEffect | null>(null)
  const [lastSlideTransitionImportEffect, setLastSlideTransitionImportEffect] =
    useState<PPTSlideTransitionImportEffect | null>(null)
  const [lastObjectAnimationImportEffect, setLastObjectAnimationImportEffect] =
    useState<PPTObjectAnimationImportEffect | null>(null)
  const [lastObjectStyleImportEffect, setLastObjectStyleImportEffect] =
    useState<PPTObjectStyleImportEffect | null>(null)
  const [lastObjectMetadataImportEffect, setLastObjectMetadataImportEffect] =
    useState<PPTObjectMetadataImportEffect | null>(null)
  const [lastObjectStateImportEffect, setLastObjectStateImportEffect] =
    useState<PPTObjectStateImportEffect | null>(null)
  const [lastObjectLayerImportEffect, setLastObjectLayerImportEffect] =
    useState<PPTObjectLayerImportEffect | null>(null)
  const [lastImageCropImportEffect, setLastImageCropImportEffect] =
    useState<PPTImageCropImportEffect | null>(null)
  const [lastImageReplaceImportEffect, setLastImageReplaceImportEffect] =
    useState<PPTImageReplaceImportEffect | null>(null)
  const [lastShapeStyleImportEffect, setLastShapeStyleImportEffect] =
    useState<PPTShapeStyleImportEffect | null>(null)
  const [lastTextStyleImportEffect, setLastTextStyleImportEffect] =
    useState<PPTTextStyleImportEffect | null>(null)
  const [lastTextBodyImportEffect, setLastTextBodyImportEffect] =
    useState<PPTTextBodyImportEffect | null>(null)
  const [lastTextAutoFitImportEffect, setLastTextAutoFitImportEffect] =
    useState<PPTTextAutoFitImportEffect | null>(null)
  const [lastLineStyleImportEffect, setLastLineStyleImportEffect] =
    useState<PPTLineStyleImportEffect | null>(null)
  const [lastObjectTransformImportEffect, setLastObjectTransformImportEffect] =
    useState<PPTObjectTransformImportEffect | null>(null)
  const [lastCommentImportEffect, setLastCommentImportEffect] =
    useState<PPTCommentImportEffect | null>(null)
  const [lastElementsJSONImportEffect, setLastElementsJSONImportEffect] =
    useState<PPTElementsJSONImportEffect | null>(null)
  const [lastClipboardImportActionKinds, setLastClipboardImportActionKinds] =
    useState('')
  const [lastStageDropImportActionKind, setLastStageDropImportActionKind] =
    useState('')
  const [lastSlideSVGClipboardEffect, setLastSlideSVGClipboardEffect] =
    useState<PPTSlideSVGClipboardEffect | null>(null)
  const [lastSelectionSVGClipboardEffect, setLastSelectionSVGClipboardEffect] =
    useState<PPTSelectionSVGClipboardEffect | null>(null)
  const [lastTableClipboardEffect, setLastTableClipboardEffect] =
    useState<PPTTableClipboardEffect | null>(null)
  const [lastTableRowsImportEffect, setLastTableRowsImportEffect] =
    useState<PPTTableRowsImportEffect | null>(null)
  const [lastFallbackHTMLImportEffect, setLastFallbackHTMLImportEffect] =
    useState<PPTFallbackHTMLImportEffect | null>(null)
  const [lastImageImportEffect, setLastImageImportEffect] = useState<PPTImageImportEffect | null>(null)
  const [lastTableImportEffect, setLastTableImportEffect] = useState<PPTTableImportEffect | null>(null)
  const [lastStyleClipboardEffect, setLastStyleClipboardEffect] = useState<PPTStyleClipboardHostCommandEffect | null>(null)
  const [lastPlaceholderVisibilityEffect, setLastPlaceholderVisibilityEffect] = useState<PPTLayoutPlaceholderVisibilityHostCommandEffect | null>(null)
  const [lastSlideRailCommandEffect, setLastSlideRailCommandEffect] = useState<SlideEditRailHostCommandEffect<string> | null>(null)
  const [lastSlideTransitionEffect, setLastSlideTransitionEffect] = useState<PPTSlideTransitionHostCommandEffect | null>(null)
  const [lastAccessibilityEffect, setLastAccessibilityEffect] = useState<SlideEditObjectAccessibilityHostCommandEffect<string, string> | null>(null)
  const [lastColorSwatchEffect, setLastColorSwatchEffect] = useState<PPTColorSwatchHostCommandEffect | null>(null)
  const [lastCommentThreadEffect, setLastCommentThreadEffect] = useState<PPTCommentThreadHostCommandEffect | null>(null)
  const [lastCornerRadiusEffect, setLastCornerRadiusEffect] = useState<SlideEditObjectCornerRadiusHostCommandEffect<string, string> | null>(null)
  const [lastFillOpacityEffect, setLastFillOpacityEffect] = useState<SlideEditObjectFillOpacityHostCommandEffect<string, string> | null>(null)
  const [lastHyperlinkEffect, setLastHyperlinkEffect] = useState<SlideEditObjectHyperlinkHostCommandEffect<string, string> | null>(null)
  const [lastImageCropEffect, setLastImageCropEffect] = useState<SlideEditObjectImageCropHostCommandEffect<string, string> | null>(null)
  const [lastImageReplaceEffect, setLastImageReplaceEffect] = useState<SlideEditObjectImageReplaceHostCommandEffect<string, string> | null>(null)
  const [lastObjectAnimationEffect, setLastObjectAnimationEffect] = useState<SlideEditObjectAnimationHostCommandEffect<string, string> | null>(null)
  const [lastObjectOpacityEffect, setLastObjectOpacityEffect] = useState<SlideEditObjectOpacityHostCommandEffect<string, string> | null>(null)
  const [lastObjectVisibilityEffect, setLastObjectVisibilityEffect] = useState<PPTObjectVisibilityHostCommandEffect | null>(null)
  const [lastShadowEffect, setLastShadowEffect] = useState<SlideEditObjectShadowHostCommandEffect<string, string> | null>(null)
  const [lastStrokeLineStyleEffect, setLastStrokeLineStyleEffect] = useState<SlideEditObjectStrokeLineStyleHostCommandEffect<string, string> | null>(null)
  const [lastTextAutoFitEffect, setLastTextAutoFitEffect] = useState<SlideEditTextAutoFitHostCommandEffect<string, string> | null>(null)
  const [lastTextFontFamilyEffect, setLastTextFontFamilyEffect] = useState<SlideEditTextFontFamilyHostCommandEffect<string, string> | null>(null)
  const [lastTextFrameInsetEffect, setLastTextFrameInsetEffect] = useState<SlideEditTextFrameInsetHostCommandEffect<string, string> | null>(null)
  const [lastInlineEditEffect, setLastInlineEditEffect] = useState<PPTInlineEditEffect | null>(null)
  const [lastResizeHandleClickMemoryEffect, setLastResizeHandleClickMemoryEffect] = useState<PPTResizeHandleClickMemoryEffect | null>(null)
  const [lastMediaImport, setLastMediaImport] = useState<PPTMediaImportResult | null>(null)
  const [lastMediaJSONImportEffect, setLastMediaJSONImportEffect] =
    useState<PPTMediaJSONImportEffect | null>(null)
  const [lastTextPasteImport, setLastTextPasteImport] = useState<PPTTextPasteImportResult | null>(null)
  const [lastTextParagraphSpacingEffect, setLastTextParagraphSpacingEffect] = useState<SlideEditTextParagraphSpacingHostCommandEffect<string, string> | null>(null)
  const [lastTextVerticalAlignmentEffect, setLastTextVerticalAlignmentEffect] = useState<SlideEditTextVerticalAlignmentHostCommandEffect<string, string> | null>(null)
  const [slideDragState, setSlideDragState] = useState<PPTSlideDragState | null>(null)
  const [lineCreationMode, setLineCreationMode] = useState<LineCreationMode | null>(null)
  const [creationTool, setCreationTool] = useState<PPTCreationTool | null>(null)
  const [isTemporaryPanActive, setIsTemporaryPanActive] = useState(false)
  const [isPanToolActive, setIsPanToolActive] = useState(false)
  const [isLaserToolActive, setIsLaserToolActive] = useState(false)
  const [laserTrailPoints, setLaserTrailPoints] = useState<Point[]>([])
  const [isEraserToolActive, setIsEraserToolActive] = useState(false)
  const [commandPaletteOpen, setCommandPaletteOpen] = useState(false)
  const [shortcutHelpOpen, setShortcutHelpOpen] = useState(false)
  const [contextMenu, setContextMenu] = useState<PPTContextMenuState | null>(null)
  const [findOpen, setFindOpen] = useState(false)
  const [findQuery, setFindQuery] = useState('')
  const [replaceQuery, setReplaceQuery] = useState('')
  const [activeFindIndex, setActiveFindIndex] = useState(0)
  const [presentationSlideId, setPresentationSlideId] = useState<string | null>(null)
  const [showGrid, setShowGrid] = useState(true)
  const [showFrameGuides, setShowFrameGuides] = useState(true)
  const [showMinimap, setShowMinimap] = useState(true)
  const [theme, setTheme] = useState<'dark' | 'light'>('light')
  const [recentColors, setRecentColors] = useState<string[]>([])
  const [textOverflowById, setTextOverflowById] = useState<Record<string, boolean>>({})
  const [past, setPast] = useState<PPTDeck[]>([])
  const [future, setFuture] = useState<PPTDeck[]>([])
  const stageRef = useRef<HTMLDivElement | null>(null)
  const canvasStageElement = usePPTCanvasStageElement()
  const setPPTStageElementRef = useCallback((element: HTMLDivElement | null) => {
    stageRef.current = element
    canvasStageElement.mount.ref(element)
  }, [canvasStageElement.mount])
  const imageInputRef = useRef<HTMLInputElement | null>(null)
  const findInputRef = useRef<HTMLInputElement | null>(null)
  const {
    onFocus: handleTopbarToolbarFocus,
    onKeyDown: handleTopbarToolbarKeyDown,
    ref: setTopbarToolbarRoot,
  } = usePPTCanvasToolbarRovingFocus<HTMLElement>()
  const slideDragSuppressClickRef = useRef(false)
  const clipboardPastePositionMemoryRef = useRef<PPTClipboardPastePositionMemory | null>(null)
  const resizeHandleClickMemoryRef = useRef<PPTCanvasPointerClickMemory>(null)
  const deckRef = useRef(deck)

  useEffect(() => {
    deckRef.current = deck
  }, [deck])

  useEffect(() => {
    globalThis.window?.dispatchEvent(new Event('ppt-ready'))
  }, [])

  const updateTextOverflowState = useCallback((
    elementId: string,
    hasOverflow: boolean,
  ) => {
    setTextOverflowById((current) =>
      current[elementId] === hasOverflow
        ? current
        : {
            ...current,
            [elementId]: hasOverflow,
          })
  }, [])

  const rememberRecentColor = useCallback((color: string) => {
    const normalized = normalizePPTSwatchColor(color)

    if (!normalized) {
      return
    }

    setRecentColors((current) => [
      normalized,
      ...current.filter((item) => normalizePPTSwatchColor(item) !== normalized),
    ].slice(0, PPT_RECENT_COLOR_LIMIT))
  }, [])

  const activeSlide = findPPTSlide(deck, activeSlideId)
  const activeSlideIndex = deck.slides.findIndex((slide) => slide.id === activeSlide.id)
  const slideRailDescriptor = useMemo(() => createSlideEditRailDescriptor({
    activeSlideId: activeSlide.id,
    getThumbnailBounds: (_slideId, index) => ({
      h: PPT_SLIDE_RAIL_THUMB_HEIGHT,
      w: PPT_SLIDE_RAIL_THUMB_WIDTH,
      x: PPT_SLIDE_RAIL_THUMB_PADDING,
      y: PPT_SLIDE_RAIL_THUMB_PADDING +
        index * (PPT_SLIDE_RAIL_THUMB_HEIGHT + PPT_SLIDE_RAIL_THUMB_GAP),
    }),
    hitTargetPadding: PPT_SLIDE_RAIL_HIT_TARGET_PADDING,
    slideOrder: deck.slides.map((slide) => slide.id),
  }), [activeSlide.id, deck.slides])
  const presentationSlide = presentationSlideId
    ? deck.slides.find((slide) => slide.id === presentationSlideId) ?? activeSlide
    : null
  const presentationSlideIndex = presentationSlide
    ? Math.max(0, deck.slides.findIndex((slide) => slide.id === presentationSlide.id))
    : -1
  const scene = useMemo(() => createPPTCanvasScene(activeSlide), [activeSlide])
  const commandAdapter = useMemo(() => createPPTCanvasCommandAdapter(), [])
  const selectedBounds = scene.getBounds(selection)
  const selectedElement = findPPTElement(activeSlide, selection[0] ?? null)
  const selectedElementAnimation = selectedElement
    ? getPPTElementAnimation(selectedElement, activeSlide)
    : null
  const selectedElements = useMemo(
    () => getPPTCanvasSelectedItems({
      getItemId: (element) => element.id,
      items: activeSlide.elements,
      selection,
    }),
    [activeSlide.elements, selection],
  )
  const selectedTextElements = useMemo(
    () => selectedElements.filter(isPPTTextElement),
    [selectedElements],
  )
  const selectedTableElement = selectedElements.length === 1 &&
    selectedElements[0]?.kind === 'table'
    ? selectedElements[0]
    : null
  const selectedTextOverflow = selectedElement && isPPTTextElement(selectedElement)
    ? textOverflowById[selectedElement.id] === true
    : false
  const selectedTextAutoFitIndicator = selectedElement && isPPTTextElement(selectedElement)
    ? getPPTTextAutoFitIndicatorState(
        activeSlide.id,
        selectedElement,
        selectedTextOverflow,
      )
    : null
  const activeLayout = useMemo(
    () => getPPTLayoutDescriptor(activeSlide.layoutId),
    [activeSlide.layoutId],
  )
  const activeSlideTransition = useMemo(
    () => getPPTSlideTransition(activeSlide),
    [activeSlide],
  )
  const activeSlideAnimationBuildOrder = useMemo(
    () => getPPTSlideAnimationBuildOrder(activeSlide),
    [activeSlide],
  )
  const activeLayoutPlaceholders = useMemo(
    () => getPPTLayoutPlaceholders(activeLayout, activeSlide),
    [activeLayout, activeSlide],
  )
  const activeLayoutPlaceholderVisibilityDescriptors = useMemo(
    () => getPPTLayoutPlaceholderVisibilityDescriptors(activeLayout, activeSlide),
    [activeLayout, activeSlide],
  )
  const slideMetadataDescriptor = useMemo(
    () => createPPTSlideMetadataInspectorDescriptor({
      slide: activeSlide,
      slideCount: deck.slides.length,
      slideIndex: activeSlideIndex,
    }),
    [activeSlide, activeSlideIndex, deck.slides.length],
  )
  const inspectorSurface = getPPTInspectorSurface({
    activeSlideId: activeSlide.id,
    selectedObjectIds: selection,
  })
  const minimapItems = useMemo<PPTMinimapItemBounds[]>(
    () => [
      {
        bounds: {
          h: PPT_SLIDE_HEIGHT,
          w: PPT_SLIDE_WIDTH,
          x: 0,
          y: 0,
        },
        id: PPT_MINIMAP_SLIDE_FRAME_ID,
      },
      ...activeSlide.elements
        .filter((element) => element.visible !== false)
        .map((element) => ({
          bounds: pptGeometryToBounds(element.geometry),
          id: element.id,
        })),
    ],
    [activeSlide.elements],
  )
  const stageRect = canvasStageElement.getRect()
  const minimapModel = stageRect && showMinimap
    ? getPPTMinimapReadModel({
        items: minimapItems,
        size: PPT_MINIMAP_SIZE,
        stageRect,
        viewport,
      })
    : null

  const selectedLineElement = selection.length === 1 && selectedElement?.kind === 'line'
    ? selectedElement
    : null
  const findMatches = useMemo(() => getPPTDeckTextMatches(deck, findQuery), [deck, findQuery])
  const clampedFindIndex = findMatches.length === 0
    ? 0
    : Math.min(activeFindIndex, findMatches.length - 1)
  const activeFindMatch = findMatches[clampedFindIndex] ?? null
  const exportCode = useMemo(() => exportPPTDeckHTML(deck), [deck])
  const slideSvgCode = useMemo(() => exportPPTSlideSVG(activeSlide), [activeSlide])
  const selectionSvgCode = useMemo(
    () => exportPPTSelectionSVG(activeSlide, selection),
    [activeSlide, selection],
  )
  const frameGuideGeometry = useMemo(
    () => getSlideEditFrameGuideGeometry({
      config: PPT_FRAME_GUIDE_CONFIG,
      frameBounds: {
        h: PPT_SLIDE_HEIGHT,
        w: PPT_SLIDE_WIDTH,
        x: 0,
        y: 0,
      },
    }),
    [],
  )
  const canExportSelectionSVG = selectionSvgCode !== null
  const hasLockedItems = activeSlide.elements.some((element) => element.locked === true)
  const hasLockedSelection = selectedElements.some((element) => element.locked === true)
  const hasHiddenSelection = selectedElements.some((element) => element.visible === false)
  const hasGroupedSelection = selectedElements.some((element) => Boolean(element.groupId))
  const canSelectSameType = useMemo(
    () => canSelectSameTypePPTSelection(activeSlide.elements, selection),
    [activeSlide.elements, selection],
  )
  const canTidySelection = useMemo(
    () => canTidyPPTSelection(activeSlide.elements, selection),
    [activeSlide.elements, selection],
  )
  const canFlipSelection = useMemo(
    () => canFlipPPTSelection(activeSlide.elements, selection),
    [activeSlide.elements, selection],
  )
  const canCopyFormatting = selectedElements.length === 1 &&
    canCopyPPTElementFormatting(selectedElements[0])
  const styleClipboardPasteAvailability = useMemo(
    () => styleClipboard
      ? getPPTStyleClipboardPasteAvailability(activeSlide.id, selectedElements, styleClipboard)
      : null,
    [activeSlide.id, selectedElements, styleClipboard],
  )
  const canPasteFormatting = styleClipboardPasteAvailability?.canPaste ?? false
  const canFormatSelectedText = selectedTextElements.length > 0 &&
    selectedTextElements.length === selectedElements.length &&
    !hasLockedSelection &&
    !hasHiddenSelection
  const textQuickFormatState = canFormatSelectedText
    ? getPPTTextQuickFormatState(selectedTextElements)
    : null
  const shapeQuickMenuState = selectedElement?.kind === 'shape' &&
    selection.length === 1 &&
    selectedElement.locked !== true &&
    selectedElement.visible !== false
    ? {
        elementId: selectedElement.id,
        shape: selectedElement.shape,
      }
    : null
  const commandAvailability = useMemo<PPTCommandAvailability>(() => ({
    ...getPPTCanvasCommandAvailability({
      canPaste: (clipboard?.objects.length ?? 0) > 0,
      canRedo: future.length > 0,
      canUndo: past.length > 0,
      config: PPT_CANVAS_COMMAND_CONFIG,
      hasGroupedSelection,
      hasHiddenSelection,
      hasLockedItems,
      hasLockedSelection,
      selection,
    }),
    copyFormatting: canCopyFormatting,
    flipSelection: canFlipSelection,
    pasteFormatting: canPasteFormatting,
    selectSameType: canSelectSameType,
    tidySelection: canTidySelection,
  }), [
    canCopyFormatting,
    canFlipSelection,
    canPasteFormatting,
    clipboard?.objects.length,
    canSelectSameType,
    canTidySelection,
    future.length,
    hasGroupedSelection,
    hasHiddenSelection,
    hasLockedItems,
    hasLockedSelection,
    past.length,
    selection,
  ])
  const canDeleteSlide = deck.slides.length > 1
  const canMoveActiveSlideDown = activeSlideIndex >= 0 && activeSlideIndex < deck.slides.length - 1
  const canMoveActiveSlideUp = activeSlideIndex > 0
  const canPasteSlide = slideClipboard !== null
  const lastSlideRailReorderPayload = lastSlideRailCommandEffect?.payload.id === 'reorder-slide'
    ? lastSlideRailCommandEffect.payload
    : null
  const canResizeSelection = selectedBounds
    ? scene.canResizeSelection?.(selection) ?? true
    : false
  const canFitSelection = selectedBounds !== null

  const fitSlide = useCallback(() => {
    fitPPTCanvasViewportToBounds({
      bounds: {
        h: PPT_SLIDE_HEIGHT,
        w: PPT_SLIDE_WIDTH,
        x: 0,
        y: 0,
      },
      setViewport,
      stageElement: canvasStageElement,
    })
  }, [canvasStageElement])

  const fitSelection = useCallback(() => {
    fitPPTCanvasViewportToBounds({
      bounds: scene.getBounds(selection),
      setViewport,
      stageElement: canvasStageElement,
    })
  }, [canvasStageElement, scene, selection])

  const fitViewportToItems = useCallback((ids?: string[]) => {
    if (!ids || ids.length === 0) {
      fitSlide()
      return
    }

    fitPPTCanvasViewportToBounds({
      bounds: scene.getBounds(ids),
      setViewport,
      stageElement: canvasStageElement,
    })
  }, [canvasStageElement, fitSlide, scene])

  const resetZoom = useCallback(() => {
    resetPPTCanvasViewport({ setViewport })
  }, [])

  const navigateMinimapToWorldPoint = useCallback((point: Point) => {
    centerPPTCanvasViewportAtWorldPoint({
      point,
      setViewport,
      stageElement: canvasStageElement,
    })
  }, [canvasStageElement])

  function startPresentation(slideId = activeSlide.id) {
    setPresentationSlideId(slideId)
    setCommandPaletteOpen(false)
    setContextMenu(null)
    setIsPanToolActive(false)
    setIsLaserToolActive(false)
    setLaserTrailPoints([])
    setIsEraserToolActive(false)
  }

  function exitPresentation() {
    setPresentationSlideId(null)
  }

  function navigatePresentation(delta: -1 | 1) {
    setPresentationSlideId((current) => {
      const slides = deckRef.current.slides

      if (slides.length === 0) {
        return current
      }

      const currentId = current ?? activeSlide.id
      const currentIndex = Math.max(0, slides.findIndex((slide) => slide.id === currentId))
      const nextIndex = (currentIndex + delta + slides.length) % slides.length

      return slides[nextIndex]?.id ?? current
    })
  }

  useLayoutEffect(() => {
    const frame = schedulePPTCanvasAnimationFrameTask({
      task: () => fitSlide(),
    })

    return () => {
      cancelPPTCanvasAnimationFrameTask({ frame })
    }
  }, [activeSlideId, fitSlide])

  useEffect(() => {
    const cleanup = bindPPTCanvasEventListener({
      listener: () => fitSlide(),
      target: window,
      type: 'resize',
    })

    return () => {
      cleanup()
    }
  }, [fitSlide])

  useEffect(() => {
    return canvasStageElement.addWheelListener(handleStageWheel)
  })

  useEffect(() => {
    if (!contextMenu) {
      return
    }

    function onPointerDown(event: PointerEvent) {
      if (isPPTCanvasTargetWithinSelector({
        selectors: '[data-ppt-context-menu]',
        target: event.target,
      })) {
        return
      }

      setContextMenu(null)
    }

    const cleanup = bindPPTCanvasEventListener({
      listener: onPointerDown,
      target: document,
      type: 'pointerdown',
    })

    return () => {
      cleanup()
    }
  }, [contextMenu])

  useEffect(() => {
    if (!findOpen) {
      return
    }

    const frame = focusPPTCanvasElementOnNextFrame({
      resolveElement: () => findInputRef.current,
      select: true,
    })

    return () => {
      cancelPPTCanvasDeferredFocus({ frame })
    }
  }, [findOpen])

  useEffect(() => {
    if (!findOpen || !activeFindMatch) {
      return
    }

    focusPPTFindMatch(activeFindMatch)
  }, [
    activeFindMatch,
    findOpen,
  ])

  useEffect(() => {
    function onKeyDown(event: KeyboardEvent) {
      if (presentationSlideId) {
        const presentationKeyboardIntent = getPPTCanvasPresentationKeyboardIntent({
          key: event.key,
        })

        if (presentationKeyboardIntent.kind === 'exit') {
          if (presentationKeyboardIntent.preventDefault) {
            event.preventDefault()
          }
          exitPresentation()
          return
        }

        if (presentationKeyboardIntent.kind === 'navigate') {
          if (presentationKeyboardIntent.preventDefault) {
            event.preventDefault()
          }
          navigatePresentation(presentationKeyboardIntent.direction)
          return
        }
      }

      if (isPPTCanvasKeyboardTypingTarget(event.target)) {
        return
      }

      if (shortcutHelpOpen) {
        const shortcutHelpKeyboardIntent = getPPTCanvasModalKeyboardIntent({
          key: event.key,
        })

        if (shortcutHelpKeyboardIntent.kind === 'close') {
          if (shortcutHelpKeyboardIntent.preventDefault) {
            event.preventDefault()
          }
          closeShortcutHelp()
        }
        return
      }

      const mod = event.metaKey || event.ctrlKey
      const beforeTypingSystemShortcutIntent = getPPTCanvasKeyboardSystemShortcutIntent({
        config: PPT_CANVAS_COMMAND_CONFIG,
        event,
        key: event.key,
        mod,
        phase: 'before-typing-target',
      })

      if (beforeTypingSystemShortcutIntent?.kind === 'open-command-palette') {
        event.preventDefault()
        openCommandPalette()
        return
      }

      if (beforeTypingSystemShortcutIntent?.kind === 'open-find-replace') {
        event.preventDefault()
        openFindStrip()
        return
      }

      const systemShortcutIntent = getPPTCanvasKeyboardSystemShortcutIntent({
        config: PPT_CANVAS_COMMAND_CONFIG,
        event,
        key: event.key,
        mod,
        phase: 'after-typing-target',
      })

      if (
        systemShortcutIntent?.kind === 'temporary-pan' &&
        !event.altKey &&
        !event.ctrlKey &&
        !event.metaKey &&
        !isPPTTemporaryPanBlockedTarget(event.target)
      ) {
        event.preventDefault()
        setContextMenu(null)
        setIsTemporaryPanActive(true)
        return
      }

      if (systemShortcutIntent?.kind === 'open-shortcut-help') {
        event.preventDefault()
        openShortcutHelp()
        return
      }

      if (contextMenu && systemShortcutIntent?.kind === 'escape') {
        event.preventDefault()
        setContextMenu(null)
        return
      }

      const contextMenuKeyboardIntent = getPPTCanvasContextMenuKeyboardIntent({
        event,
        key: event.key,
      })

      if (contextMenuKeyboardIntent?.kind === 'open-context-menu') {
        if (openPPTContextMenuAtSelection()) {
          event.preventDefault()
        }
        return
      }

      const viewportIntent = getPPTCanvasKeyboardViewportShortcutIntent({
        config: PPT_CANVAS_COMMAND_CONFIG,
        event,
        key: event.key,
        mod,
        selection,
      })

      if (viewportIntent && isPPTCanvasKeyboardViewportIntent(viewportIntent)) {
        if (viewportIntent.preventDefault) {
          event.preventDefault()
        }

        runPPTCanvasKeyboardViewportIntent({
          handlers: {
            fitToItems: fitViewportToItems,
            resetViewport: resetZoom,
            zoom,
          },
          intent: viewportIntent,
        })
        return
      }

      const styleClipboardKeyboardIntent = getSlideEditStyleClipboardKeyboardIntent({
        event,
        key: event.key,
        mod,
      })

      if (styleClipboardKeyboardIntent?.kind === 'copy-formatting') {
        if (commandAvailability.copyFormatting) {
          event.preventDefault()
          copyFormatting()
        }
        return
      }

      if (styleClipboardKeyboardIntent?.kind === 'paste-formatting') {
        if (commandAvailability.pasteFormatting) {
          event.preventDefault()
          pasteFormatting()
        }
        return
      }

      const standardCommandIntent = getPPTCanvasKeyboardBuiltinCommandShortcutIntent({
        config: PPT_CANVAS_COMMAND_CONFIG,
        event,
        key: event.key,
        mod,
        selection,
      })

      if (
        standardCommandIntent &&
        isPPTCanvasKeyboardCommandIntent(standardCommandIntent) &&
        isPPTCanvasStandardCommandIntentKind(standardCommandIntent.kind)
      ) {
        if (standardCommandIntent.kind === 'paste-selection' && !commandAvailability.paste) {
          return
        }

        if (standardCommandIntent.preventDefault) {
          event.preventDefault()
        }

        runPPTCanvasKeyboardCommandIntent({
          handlers: {
            copySelection,
            cutSelection,
            deleteSelection,
            duplicateSelection,
            editSelection: noopPPTKeyboardCommandHandler,
            groupSelection,
            lockSelection: lockSelectedElements,
            moveSelection: nudgeSelection,
            pasteSelection,
            quickCreateSticky: noopPPTKeyboardCommandHandler,
            redoHistory: redo,
            reorderSelection,
            selectAll: selectAllElements,
            undoHistory: undo,
            ungroupSelection,
            unlockAll: unlockAllElements,
          },
          intent: standardCommandIntent,
        })
        return
      }

      const textFormattingKeyboardIntent = getSlideEditTextFormattingKeyboardIntent({
        altKey: event.altKey,
        key: event.key,
        mod,
        shiftKey: event.shiftKey,
      })

      if (textFormattingKeyboardIntent?.kind === 'toggle-bold') {
        if (canFormatSelectedText) {
          if (textFormattingKeyboardIntent.preventDefault) {
            event.preventDefault()
          }
          toggleSelectedTextBold()
        }
        return
      }

      const toolShortcutIntent = getPPTToolShortcutIntent(event)

      if (toolShortcutIntent && isPPTCanvasKeyboardToolIntent(toolShortcutIntent)) {
        runPPTCanvasKeyboardToolIntent({
          handlers: {
            setTool: activatePPTToolShortcut,
          },
          intent: toolShortcutIntent,
        })
        event.preventDefault()
        return
      }

      const deckNavigationKeyboardIntent = getSlideEditDeckNavigationKeyboardIntent({
        activeSlideId: activeSlide.id,
        altKey: event.altKey,
        ctrlKey: event.ctrlKey,
        key: event.key,
        metaKey: event.metaKey,
        shiftKey: event.shiftKey,
        slideOrder: slideRailDescriptor.slideOrder,
      })

      if (deckNavigationKeyboardIntent) {
        if (deckNavigationKeyboardIntent.preventDefault) {
          event.preventDefault()
        }
        selectSlide(deckNavigationKeyboardIntent.targetSlideId)
        return
      }

      if (systemShortcutIntent?.kind === 'escape') {
        event.preventDefault()
        setEditingId(null)
        setInteraction(null)
        setIsTemporaryPanActive(false)
        setLineCreationMode(null)
        setCreationTool(null)
        setIsLaserToolActive(false)
        setLaserTrailPoints([])
        setIsEraserToolActive(false)
        setContextMenu(null)
        setSelection([])
        return
      }

      const nudgeIntent = getPPTCanvasKeyboardNudgeShortcutIntent({
        config: PPT_CANVAS_COMMAND_CONFIG,
        event,
        key: event.key,
        mod,
        selection,
      })

      if (nudgeIntent?.kind === 'nudge-selection') {
        event.preventDefault()
        nudgeSelection(nudgeIntent.dx, nudgeIntent.dy)
        return
      }
    }

    function releaseTemporaryPan() {
      setIsTemporaryPanActive(false)
      setInteraction((current) => current?.kind === 'pan' ? null : current)
    }

    function onKeyUp(event: KeyboardEvent) {
      if (shouldReleasePPTCanvasKeyboardTemporaryPan({
        config: PPT_CANVAS_COMMAND_CONFIG,
        event,
      })) {
        releaseTemporaryPan()
      }
    }

    const cleanup = bindPPTCanvasEventListeners({
      listeners: [
        { listener: onKeyDown, target: window, type: 'keydown' },
        { listener: onKeyUp, target: window, type: 'keyup' },
        { listener: releaseTemporaryPan, target: window, type: 'blur' },
      ],
    })

    return () => {
      cleanup()
    }
  })

  useEffect(() => {
    function onPaste(event: ClipboardEvent) {
      if (isPPTCanvasKeyboardTypingTarget(event.target)) {
        return
      }

      const slideClipboardImport =
        getPPTSlideClipboardFromDataTransfer(event.clipboardData)

      if (slideClipboardImport && pasteSlideClipboardPayload(
        slideClipboardImport.payload,
        { importFormat: slideClipboardImport.format },
      )) {
        event.preventDefault()
        return
      }

      const slideFallbackHTMLSource =
        getPPTSlideFallbackHTMLSourceFromDataTransfer(event.clipboardData)

      if (
        slideFallbackHTMLSource &&
        pastePPTSlideFallbackHTMLSource(slideFallbackHTMLSource)
      ) {
        event.preventDefault()
        return
      }

      const slideJSONSource =
        getPPTSlideJSONSourceFromDataTransfer(event.clipboardData)

      if (slideJSONSource && pastePPTSlideJSONSource(slideJSONSource)) {
        event.preventDefault()
        return
      }

      const deckHTMLSource =
        getPPTDeckHTMLSourceFromDataTransfer(event.clipboardData)

      if (deckHTMLSource && pastePPTDeckHTMLSource(deckHTMLSource)) {
        event.preventDefault()
        return
      }

      const deckJSONSource =
        getPPTDeckJSONSourceFromDataTransfer(event.clipboardData)

      if (deckJSONSource && pastePPTDeckJSONSource(deckJSONSource)) {
        event.preventDefault()
        return
      }

      const elementJSONSource =
        getPPTElementsJSONSourceFromDataTransfer(event.clipboardData)

      if (elementJSONSource && pastePPTElementsJSONSource(elementJSONSource)) {
        event.preventDefault()
        return
      }

      const deckFallbackHTMLSource =
        getPPTDeckFallbackHTMLSourceFromDataTransfer(event.clipboardData)

      if (
        deckFallbackHTMLSource &&
        pastePPTDeckFallbackHTMLSource(deckFallbackHTMLSource)
      ) {
        event.preventDefault()
        return
      }

      const deckMarkdownOutlineSource =
        getPPTDeckMarkdownOutlineSourceFromDataTransfer(event.clipboardData)

      if (
        deckMarkdownOutlineSource &&
        pastePPTDeckMarkdownOutlineSource(deckMarkdownOutlineSource)
      ) {
        event.preventDefault()
        return
      }

      const slideMetadataSource =
        getPPTSlideMetadataSourceFromDataTransfer(event.clipboardData)

      if (
        slideMetadataSource &&
        pastePPTSlideMetadataSource(slideMetadataSource)
      ) {
        event.preventDefault()
        return
      }

      const slideLayoutSource =
        getPPTSlideLayoutSourceFromDataTransfer(event.clipboardData)

      if (
        slideLayoutSource &&
        pastePPTSlideLayoutSource(slideLayoutSource)
      ) {
        event.preventDefault()
        return
      }

      const slideTransitionSource =
        getPPTSlideTransitionSourceFromDataTransfer(event.clipboardData)

      if (
        slideTransitionSource &&
        pastePPTSlideTransitionSource(slideTransitionSource)
      ) {
        event.preventDefault()
        return
      }

      const objectAnimationSource =
        getPPTObjectAnimationSourceFromDataTransfer(event.clipboardData)

      if (
        objectAnimationSource &&
        pastePPTObjectAnimationSource(objectAnimationSource)
      ) {
        event.preventDefault()
        return
      }

      const objectStyleSource =
        getPPTObjectStyleSourceFromDataTransfer(event.clipboardData)

      if (
        objectStyleSource &&
        pastePPTObjectStyleSource(objectStyleSource)
      ) {
        event.preventDefault()
        return
      }

      const objectMetadataSource =
        getPPTObjectMetadataSourceFromDataTransfer(event.clipboardData)

      if (
        objectMetadataSource &&
        pastePPTObjectMetadataSource(objectMetadataSource)
      ) {
        event.preventDefault()
        return
      }

      const objectStateSource =
        getPPTObjectStateSourceFromDataTransfer(event.clipboardData)

      if (
        objectStateSource &&
        pastePPTObjectStateSource(objectStateSource)
      ) {
        event.preventDefault()
        return
      }

      const objectLayerSource =
        getPPTObjectLayerSourceFromDataTransfer(event.clipboardData)

      if (
        objectLayerSource &&
        pastePPTObjectLayerSource(objectLayerSource)
      ) {
        event.preventDefault()
        return
      }

      const objectTransformSource =
        getPPTObjectTransformSourceFromDataTransfer(event.clipboardData)

      if (
        objectTransformSource &&
        pastePPTObjectTransformSource(objectTransformSource)
      ) {
        event.preventDefault()
        return
      }

      const textAutoFitSource =
        getPPTTextAutoFitSourceFromDataTransfer(event.clipboardData)

      if (
        textAutoFitSource &&
        pastePPTTextAutoFitSource(textAutoFitSource)
      ) {
        event.preventDefault()
        return
      }

      const textBodySource =
        getPPTTextBodySourceFromDataTransfer(event.clipboardData)

      if (textBodySource && pastePPTTextBodySource(textBodySource)) {
        event.preventDefault()
        return
      }

      const tableRowsSource =
        getPPTTableRowsSourceFromDataTransfer(event.clipboardData)

      if (tableRowsSource && pastePPTTableRowsSource(tableRowsSource)) {
        event.preventDefault()
        return
      }

      const commentSource =
        getPPTCommentSourceFromDataTransfer(event.clipboardData)

      if (commentSource && pastePPTCommentSource(commentSource)) {
        event.preventDefault()
        return
      }

      const imageReplaceSource =
        getPPTImageReplaceSourceFromDataTransfer(event.clipboardData)

      if (
        imageReplaceSource &&
        pastePPTImageReplaceSource(imageReplaceSource)
      ) {
        event.preventDefault()
        return
      }

      const imageCropSource =
        getPPTImageCropSourceFromDataTransfer(event.clipboardData)

      if (imageCropSource && pastePPTImageCropSource(imageCropSource)) {
        event.preventDefault()
        return
      }

      const shapeStyleSource =
        getPPTShapeStyleSourceFromDataTransfer(event.clipboardData)

      if (shapeStyleSource && pastePPTShapeStyleSource(shapeStyleSource)) {
        event.preventDefault()
        return
      }

      const lineStyleSource =
        getPPTLineStyleSourceFromDataTransfer(event.clipboardData)

      if (lineStyleSource && pastePPTLineStyleSource(lineStyleSource)) {
        event.preventDefault()
        return
      }

      const textStyleSource =
        getPPTTextStyleSourceFromDataTransfer(event.clipboardData)

      if (textStyleSource && pastePPTTextStyleSource(textStyleSource)) {
        event.preventDefault()
        return
      }

      const slideNotesSource =
        getPPTSlideNotesSourceFromDataTransfer(event.clipboardData)

      if (slideNotesSource && pastePPTSlideNotesSource(slideNotesSource)) {
        event.preventDefault()
        return
      }

      const mediaJSONSource =
        getPPTMediaJSONSourceFromDataTransfer(event.clipboardData)

      if (mediaJSONSource && pastePPTMediaJSONSource(mediaJSONSource)) {
        event.preventDefault()
        return
      }

      const richClipboard = getPPTRichClipboardFromDataTransfer(event.clipboardData)

      if (richClipboard && pasteClipboardPayload(richClipboard.payload, {
        importFormat: richClipboard.format,
      })) {
        event.preventDefault()
        return
      }

      const importActions = getPPTClipboardImportActions(event.clipboardData)
      setLastClipboardImportActionKinds(
        importActions.map((action) => action.kind).join(' '),
      )

      for (const action of importActions) {
        if (runPPTClipboardImportAction(action)) {
          event.preventDefault()
          return
        }
      }
    }

    const cleanup = bindPPTCanvasEventListener({
      listener: onPaste,
      target: window,
      type: 'paste',
    })

    return () => {
      cleanup()
    }
  })

  function commitDeck(update: (current: PPTDeck) => PPTDeck) {
    setDeck((current) => {
      const next = update(current)

      if (JSON.stringify(next) === JSON.stringify(current)) {
        return current
      }

      setPast((history) => [...history.slice(-79), current])
      setFuture([])

      return next
    })
  }

  function commitElementCommand(
    run: (slide: PPTSlide) => PPTCanvasCommandItemsResult<PPTElement> | null,
  ) {
    commitDeck((current) => updatePPTDeckSlide(current, activeSlide.id, (slide) => {
      const result = run(slide)

      if (!result) {
        return slide
      }

      setSelection(result.selection)

      return {
        ...slide,
        elements: syncPPTLineConnections(
          result.items,
          getPPTSelectedLineIds(result.items, result.selection),
        ),
      }
    }))
  }

  function undo() {
    setPast((history) => {
      const previous = history.at(-1)

      if (!previous) {
        return history
      }

      setFuture((items) => [deckRef.current, ...items])
      setDeck(previous)

      return history.slice(0, -1)
    })
  }

  function redo() {
    setFuture((items) => {
      const next = items[0]

      if (!next) {
        return items
      }

      setPast((history) => [...history, deckRef.current])
      setDeck(next)

      return items.slice(1)
    })
  }

  function selectSlide(slideId: string) {
    setActiveSlideId(slideId)
    setSelection([])
    setEditingId(null)
    setInteraction(null)
    setIsPanToolActive(false)
    setIsLaserToolActive(false)
    setLaserTrailPoints([])
    setIsEraserToolActive(false)
    setContextMenu(null)
  }

  function focusPPTSlideThumb(slideId: string) {
    focusPPTCanvasElementBySelectorOnNextFrame<HTMLButtonElement>({
      match: ({ element }) =>
        element.getAttribute('data-ppt-slide-id') === slideId,
      root: document,
      selector: '.ppt-thumb',
    })
  }

  function handleSlideThumbSelect(
    slideId: string,
    event: ReactMouseEvent<HTMLButtonElement>,
  ) {
    if (slideDragSuppressClickRef.current) {
      event.preventDefault()
      event.stopPropagation()
      return
    }

    setLastSlideRailCommandEffect(toSlideEditRailHostCommandEffect({
      id: 'select-active-slide',
      slideId,
    }))
    selectSlide(slideId)
  }

  function handleSlideThumbKeyDown(
    slideId: string,
    event: ReactKeyboardEvent<HTMLButtonElement>,
  ) {
    if (event.altKey || event.ctrlKey || event.metaKey) {
      return
    }

    const slideOrder = deck.slides.map((slide) => slide.id)
    const railKey = event.key === 'ArrowRight'
      ? 'ArrowDown'
      : event.key === 'ArrowLeft'
        ? 'ArrowUp'
        : event.key
    const intent = getSlideEditRailListboxKeyboardIntent({
      activeSlideId: slideId,
      key: railKey,
      slideOrder,
    })

    if (!intent) {
      return
    }

    event.preventDefault()
    event.stopPropagation()

    const effect = getSlideEditRailKeyboardCommandEffect(intent)
    const targetSlideId = effect?.payload.id === 'select-active-slide'
      ? effect.payload.slideId
      : null

    if (!effect || !targetSlideId) {
      return
    }

    setLastSlideRailCommandEffect(effect)
    selectSlide(targetSlideId)
    focusPPTSlideThumb(targetSlideId)
  }

  function openFindStrip() {
    setFindOpen(true)
    setCommandPaletteOpen(false)
    setEditingId(null)
    setInteraction(null)
    setLineCreationMode(null)
    setCreationTool(null)
    setIsLaserToolActive(false)
    setLaserTrailPoints([])
    setIsEraserToolActive(false)
    setContextMenu(null)

    if (activeFindMatch) {
      focusPPTFindMatch(activeFindMatch)
    }
  }

  function closeFindStrip() {
    setFindOpen(false)
  }

  function openCommandPalette() {
    setCommandPaletteOpen(true)
    setShortcutHelpOpen(false)
    setFindOpen(false)
    setEditingId(null)
    setInteraction(null)
    setLineCreationMode(null)
    setCreationTool(null)
    setIsLaserToolActive(false)
    setLaserTrailPoints([])
    setIsEraserToolActive(false)
    setContextMenu(null)
  }

  function closeCommandPalette() {
    setCommandPaletteOpen(false)
  }

  function openShortcutHelp() {
    setShortcutHelpOpen(true)
    setCommandPaletteOpen(false)
    setFindOpen(false)
    setEditingId(null)
    setInteraction(null)
    setLineCreationMode(null)
    setCreationTool(null)
    setIsLaserToolActive(false)
    setLaserTrailPoints([])
    setIsEraserToolActive(false)
    setContextMenu(null)
  }

  function closeShortcutHelp() {
    setShortcutHelpOpen(false)
  }

  function focusPPTFindMatch(match: PPTFindMatch) {
    setActiveSlideId(match.slideId)
    setSelection([match.elementId])
    setEditingId(null)
    setInteraction(null)
    setLineCreationMode(null)
    setCreationTool(null)
    setIsLaserToolActive(false)
    setLaserTrailPoints([])
    setIsEraserToolActive(false)
    setContextMenu(null)
  }

  function updateFindQuery(query: string) {
    setFindQuery(query)
    setActiveFindIndex(0)

    const firstMatch = getPPTDeckTextMatches(deckRef.current, query)[0]

    if (firstMatch) {
      focusPPTFindMatch(firstMatch)
    }
  }

  function goToFindMatch(delta: -1 | 1) {
    if (findMatches.length === 0) {
      return
    }

    const nextIndex = (clampedFindIndex + delta + findMatches.length) % findMatches.length
    const nextMatch = findMatches[nextIndex]

    setActiveFindIndex(nextIndex)
    focusPPTFindMatch(nextMatch)
  }

  function replaceActiveFindMatch() {
    if (!activeFindMatch || findQuery.length === 0) {
      return
    }

    const match = activeFindMatch
    const nextMatchCount = Math.max(0, findMatches.length - 1)

    commitDeck((current) =>
      updatePPTDeckElement(current, match.slideId, match.elementId, (element) =>
        replacePPTElementTextRange(element, match.start, match.end, replaceQuery),
      ),
    )
    setActiveFindIndex((current) =>
      nextMatchCount === 0 ? 0 : Math.min(current, nextMatchCount - 1))
  }

  function replaceAllFindMatches() {
    if (findMatches.length === 0 || findQuery.length === 0) {
      return
    }

    commitDeck((current) => ({
      ...current,
      slides: current.slides.map((slide) => ({
        ...slide,
        elements: slide.elements.map((element) =>
          replaceAllPPTElementTextMatches(element, findQuery, replaceQuery)),
      })),
    }))
    setActiveFindIndex(0)
  }

  function handleFindKeyDown(event: ReactKeyboardEvent<HTMLInputElement>) {
    const intent = getPPTCanvasFindInputKeyboardIntent({
      key: event.key,
      shiftKey: event.shiftKey,
    })

    if (intent.preventDefault) {
      event.preventDefault()
    }

    if (intent.kind === 'find-match') {
      goToFindMatch(intent.direction)
      return
    }

    if (intent.kind === 'close-find') {
      closeFindStrip()
    }
  }

  function addSlide() {
    commitDeck((current) => {
      const nextIndex = current.slides.length + 1
      const id = createPPTSlideId(current)
      const slide: PPTSlide = {
        background: { color: '#ffffff' },
        elements: [{
          geometry: { h: 72, w: 720, x: 84, y: 82 },
          id: `${id}-title`,
          kind: 'textBox',
          name: 'Title',
          style: { color: '#111827', fontSize: 44, fontWeight: 'bold' },
          textBody: createPPTTextBody('Untitled slide'),
        }],
        id,
        name: `Slide ${nextIndex}`,
        notes: '',
      }

      setActiveSlideId(slide.id)
      setSelection([slide.elements[0].id])

      return {
        ...current,
        slides: [...current.slides, slide],
      }
    })
  }

  function duplicateActiveSlide() {
    commitDeck((current) => {
      const source = current.slides.find((slide) => slide.id === activeSlide.id)

      if (!source) {
        return current
      }

      const id = createPPTSlideId(current)
      const slide = clonePPTSlide(source, id)
      const result = insertPPTSlideAtTargetPlacement({
        placement: 'after',
        slide,
        slides: current.slides,
        targetSlideId: activeSlide.id,
      })

      if (!result) {
        return current
      }

      selectSlide(slide.id)

      return {
        ...current,
        slides: result.items,
      }
    })
  }

  function copyActiveSlide() {
    const payload = createPPTSlideClipboardPayload(activeSlide)
    const html = createPPTSlideClipboardHTML(payload)
    const effect = createPPTSlideClipboardEffect({
      html,
      payload,
      writeMode: 'pending',
    })

    setSlideClipboard(payload)
    setLastSlideClipboardEffect(effect)

    void writePPTSlideClipboardPayload({ payload }).then((writeMode) => {
      setLastSlideClipboardEffect((current) =>
        current &&
        current.sourceSlideId === effect.sourceSlideId &&
        current.htmlLength === effect.htmlLength
          ? {
              ...current,
              writeMode,
            }
          : current)
    })
  }

  function pasteCopiedSlide() {
    if (!slideClipboard) {
      return false
    }

    return pasteSlideClipboardPayload(slideClipboard)
  }

  function pasteSlideClipboardPayload(
    payload: PPTSlideClipboardPayload,
    options: { importFormat?: PPTRichClipboardImportFormat } = {},
  ) {
    let pastedSlide: PPTSlide | null = null

    commitDeck((current) => {
      const targetSlideId = current.slides.some((slide) =>
        slide.id === activeSlide.id)
        ? activeSlide.id
        : current.slides.at(-1)?.id

      if (!targetSlideId) {
        return current
      }

      const id = createPPTSlideId(current)
      const slide = clonePPTSlide(payload.slide, id)
      const result = insertPPTSlideAtTargetPlacement({
        placement: 'after',
        slide,
        slides: current.slides,
        targetSlideId,
      })

      if (!result) {
        return current
      }

      pastedSlide = slide
      setSlideClipboard(payload)
      setLastSlideClipboardEffect(createPPTSlideClipboardEffect({
        imported: Boolean(options.importFormat),
        importFormat: options.importFormat,
        payload,
        targetSlideId: slide.id,
      }))
      selectSlide(slide.id)

      return {
        ...current,
        slides: result.items,
      }
    })

    return pastedSlide !== null
  }

  function pastePPTSlideFallbackHTMLSource(
    source: PPTSlideClipboardFallbackHTMLSource,
  ) {
    let pastedSlide: PPTSlide | null = null

    commitDeck((current) => {
      const targetSlideId = current.slides.some((slide) =>
        slide.id === activeSlide.id)
        ? activeSlide.id
        : current.slides.at(-1)?.id

      if (!targetSlideId) {
        return current
      }

      const slide = createPPTSlideFromFallbackHTMLSource(current, source)
      const result = insertPPTSlideAtTargetPlacement({
        placement: 'after',
        slide,
        slides: current.slides,
        targetSlideId,
      })

      if (!result) {
        return current
      }

      pastedSlide = slide
      setLastSlideClipboardEffect(createPPTSlideClipboardFallbackHTMLEffect({
        source,
        targetSlideId: slide.id,
      }))
      selectSlide(slide.id)

      return {
        ...current,
        slides: result.items,
      }
    })

    return pastedSlide !== null
  }

  function pastePPTDeckHTMLSource(source: PPTDeckHTMLImportSource) {
    if (source.deck.slides.length === 0) {
      return false
    }

    commitDeck((current) => {
      const targetSlideId = current.slides.some((slide) =>
        slide.id === activeSlide.id)
        ? activeSlide.id
        : current.slides.at(-1)?.id

      if (!targetSlideId) {
        return current
      }

      const importedSlides = clonePPTDeckSlidesForImport(
        current,
        source.deck.slides,
      )

      if (importedSlides.length === 0) {
        return current
      }

      let slides = current.slides
      let anchorSlideId = targetSlideId

      for (const slide of importedSlides) {
        const result = insertPPTSlideAtTargetPlacement({
          placement: 'after',
          slide,
          slides,
          targetSlideId: anchorSlideId,
        })

        if (!result) {
          return current
        }

        slides = result.items
        anchorSlideId = slide.id
      }

      setLastDeckHTMLImportEffect(createPPTDeckHTMLImportEffect({
        importedSlides,
        source,
      }))
      selectSlide(importedSlides[0].id)

      return {
        ...current,
        slides,
      }
    })

    return true
  }

  function pastePPTDeckJSONSource(source: PPTDeckJSONImportSource) {
    if (source.deck.slides.length === 0) {
      return false
    }

    commitDeck((current) => {
      const targetSlideId = current.slides.some((slide) =>
        slide.id === activeSlide.id)
        ? activeSlide.id
        : current.slides.at(-1)?.id

      if (!targetSlideId) {
        return current
      }

      const importedSlides = clonePPTDeckSlidesForImport(
        current,
        source.deck.slides,
      )

      if (importedSlides.length === 0) {
        return current
      }

      let slides = current.slides
      let anchorSlideId = targetSlideId

      for (const slide of importedSlides) {
        const result = insertPPTSlideAtTargetPlacement({
          placement: 'after',
          slide,
          slides,
          targetSlideId: anchorSlideId,
        })

        if (!result) {
          return current
        }

        slides = result.items
        anchorSlideId = slide.id
      }

      setLastDeckJSONImportEffect(createPPTDeckJSONImportEffect({
        importedSlides,
        source,
      }))
      selectSlide(importedSlides[0].id)

      return {
        ...current,
        slides,
      }
    })

    return true
  }

  function pastePPTSlideJSONSource(source: PPTSlideJSONImportSource) {
    if (source.slides.length === 0) {
      return false
    }

    commitDeck((current) => {
      const targetSlideId = current.slides.some((slide) =>
        slide.id === activeSlide.id)
        ? activeSlide.id
        : current.slides.at(-1)?.id

      if (!targetSlideId) {
        return current
      }

      const importedSlides = clonePPTDeckSlidesForImport(
        current,
        source.slides,
      )

      if (importedSlides.length === 0) {
        return current
      }

      let slides = current.slides
      let anchorSlideId = targetSlideId

      for (const slide of importedSlides) {
        const result = insertPPTSlideAtTargetPlacement({
          placement: 'after',
          slide,
          slides,
          targetSlideId: anchorSlideId,
        })

        if (!result) {
          return current
        }

        slides = result.items
        anchorSlideId = slide.id
      }

      setLastSlideJSONImportEffect(createPPTSlideJSONImportEffect({
        importedSlides,
        source,
      }))
      selectSlide(importedSlides[0].id)

      return {
        ...current,
        slides,
      }
    })

    return true
  }

  function pastePPTElementsJSONSource(source: PPTElementsJSONImportSource) {
    if (source.objects.length === 0) {
      return false
    }

    const payload = createPPTClipboardPayload({
      objects: source.objects,
      operation: 'copy',
      selectedObjectIds: source.selectedObjectIds,
      sourceSlideId: source.sourceSlideId,
    })

    if (!pasteClipboardPayload(payload)) {
      return false
    }

    setLastElementsJSONImportEffect(createPPTElementsJSONImportEffect(source))
    return true
  }

  function pastePPTSlideNotesSource(source: PPTSlideNotesImportSource) {
    const effect = toPPTSlideMetadataHostCommandEffect({
      fieldId: 'notes',
      id: 'update-slide-notes',
      slideId: activeSlide.id,
      value: source.notes,
    })

    setLastSlideNotesImportEffect(createPPTSlideNotesImportEffect({
      effect,
      source,
    }))
    commitDeck((current) =>
      updatePPTDeckSlide(current, effect.selection.slideId, (slide) =>
        applyPPTSlideMetadataHostCommandEffect(slide, effect)))

    return true
  }

  function pastePPTSlideMetadataSource(source: PPTSlideMetadataImportSource) {
    const effects = createPPTSlideMetadataImportCommandEffects({
      slideId: activeSlide.id,
      source,
    })

    if (effects.length === 0) {
      return false
    }

    setLastSlideMetadataImportEffect(createPPTSlideMetadataImportEffect({
      effects,
      source,
    }))
    commitDeck((current) =>
      updatePPTDeckSlide(current, activeSlide.id, (slide) =>
        effects.reduce(
          (currentSlide, effect) =>
            applyPPTSlideMetadataHostCommandEffect(currentSlide, effect),
          slide,
        )))

    return true
  }

  function pastePPTSlideLayoutSource(source: PPTSlideLayoutImportSource) {
    const layout = getPPTLayoutDescriptor(source.layoutId ?? activeSlide.layoutId)
    const layoutEffect = source.layoutId === undefined
      ? null
      : getSlideEditLayoutApplyCommandEffect({
        existingObjectPolicy: 'preserve-existing-objects',
        layoutId: layout.layoutId,
        selectedObjectIds: selection,
        slideId: activeSlide.id,
      })
    const placeholderEffects = source.hiddenPlaceholderIds === undefined
      ? []
      : createPPTSlideLayoutPlaceholderImportCommandEffects({
        hiddenPlaceholderIds: source.hiddenPlaceholderIds,
        layout,
        selectedObjectIds: selection,
        slideId: activeSlide.id,
      })

    if (
      !layoutEffect &&
      placeholderEffects.length === 0 &&
      source.themeId === undefined
    ) {
      return false
    }

    if (placeholderEffects.length > 0) {
      setLastPlaceholderVisibilityEffect(
        placeholderEffects[placeholderEffects.length - 1],
      )
    }

    setLastSlideLayoutImportEffect(createPPTSlideLayoutImportEffect({
      layoutEffect,
      placeholderEffects,
      source,
    }))

    commitDeck((current) =>
      updatePPTDeckSlide(current, activeSlide.id, (slide) => {
        let nextSlide = slide

        if (layoutEffect) {
          nextSlide = {
            ...nextSlide,
            layoutId: layoutEffect.payload.layoutId,
            themeId: source.themeId ?? PPT_THEME_DESCRIPTOR.themeId,
          }
        } else if (source.themeId !== undefined) {
          nextSlide = {
            ...nextSlide,
            themeId: source.themeId,
          }
        }

        if (placeholderEffects.length > 0) {
          nextSlide = placeholderEffects.reduce(
            (currentSlide, effect) =>
              applyPPTLayoutPlaceholderVisibilityHostCommandEffect(
                currentSlide,
                effect,
              ),
            nextSlide,
          )
        }

        const nextLayout = getPPTLayoutDescriptor(nextSlide.layoutId)

        return {
          ...nextSlide,
          hiddenPlaceholderIds: normalizePPTSlideLayoutHiddenPlaceholderIds(
            nextSlide.hiddenPlaceholderIds ?? [],
            nextLayout,
          ),
        }
      }))

    return true
  }

  function pastePPTSlideTransitionSource(source: PPTSlideTransitionImportSource) {
    const effects = createPPTSlideTransitionImportCommandEffects({
      currentTransition: activeSlideTransition,
      slideId: activeSlide.id,
      source,
    })

    if (effects.length === 0) {
      return false
    }

    setLastSlideTransitionEffect(effects[effects.length - 1])
    setLastSlideTransitionImportEffect(createPPTSlideTransitionImportEffect({
      effects,
      source,
    }))
    commitDeck((current) =>
      updatePPTDeckSlide(current, activeSlide.id, (slide) => ({
        ...slide,
        transition: effects.reduce(
          (transition, effect) =>
            applyPPTSlideTransitionUpdateCommand(transition, effect.payload),
          getPPTSlideTransition(slide),
        ),
      })))

    return true
  }

  function pastePPTObjectAnimationSource(source: PPTObjectAnimationImportSource) {
    const objectIds = activeSlide.elements
      .filter((element) => selection.includes(element.id))
      .map((element) => element.id)

    if (objectIds.length === 0) {
      return false
    }

    const effects = createPPTObjectAnimationImportCommandEffects({
      objectIds,
      slideId: activeSlide.id,
      source,
    })

    if (effects.length === 0) {
      return false
    }

    setLastObjectAnimationEffect(effects[effects.length - 1])
    setLastObjectAnimationImportEffect(createPPTObjectAnimationImportEffect({
      effects,
      source,
    }))
    commitDeck((current) =>
      updatePPTDeckSlide(current, activeSlide.id, (slide) => ({
        ...slide,
        elements: slide.elements.map((element) => {
          const elementEffects = effects.filter((effect) =>
            effect.payload.objectId === element.id)

          if (elementEffects.length === 0) {
            return element
          }

          const animation = elementEffects.reduce(
            (currentAnimation, effect) => {
              const { field, value } = toPPTElementAnimationUpdate(effect.payload)

              return normalizePPTElementAnimation({
                ...currentAnimation,
                [field]: value,
              }, slide, element.id)
            },
            getPPTElementAnimation(element, slide),
          )

          return {
            ...element,
            animation,
          }
        }),
      })))

    return true
  }

  function pastePPTObjectStyleSource(source: PPTObjectStyleImportSource) {
    if (selectedElements.length === 0) {
      return false
    }

    const sourceElement = selectedElements[0]
    const styleClipboard: PPTStyleClipboard = {
      categories: ['object'],
      object: {
        opacity: source.object.opacity ?? getPPTElementOpacity(sourceElement),
        shadow: source.object.shadow === undefined
          ? hasPPTElementShadow(sourceElement)
            ? clonePPTElementShadow(getPPTElementShadow(sourceElement))
            : null
          : source.object.shadow
            ? clonePPTElementShadow(source.object.shadow)
            : null,
      },
      sourceId: 'ppt-object-style-json',
      sourceKind: sourceElement.kind,
      type: 'slide-style-clipboard',
    }
    const effect = createSlideEditStyleClipboardPasteCommandEffect({
      clipboard: createPPTStyleClipboardDescriptor(activeSlide.id, styleClipboard),
      targetSlideId: activeSlide.id,
      targets: getPPTStyleClipboardTargetInputs(selectedElements),
    })

    if (!effect) {
      return false
    }

    setStyleClipboard(styleClipboard)
    setLastStyleClipboardEffect(effect)
    setLastObjectStyleImportEffect(createPPTObjectStyleImportEffect({
      effect,
      source,
    }))

    const categoryApplicationsByObjectId = new Map(
      effect.payload.categoryApplications.map((application) => [
        application.objectId,
        application.appliedCategoryIds,
      ]),
    )

    commitDeck((current) =>
      updatePPTDeckSlide(current, activeSlide.id, (slide) => ({
        ...slide,
        elements: mapPPTElementsByIds(
          slide.elements,
          effect.payload.categoryApplications.map((application) =>
            application.objectId),
          (element) => {
            const appliedCategoryIds = categoryApplicationsByObjectId.get(element.id)

            return appliedCategoryIds
              ? applyPPTStyleClipboardToElement(
                  element,
                  styleClipboard,
                  appliedCategoryIds,
                )
              : element
          },
        ),
      })))

    return true
  }

  function pastePPTObjectMetadataSource(source: PPTObjectMetadataImportSource) {
    const objectIds = activeSlide.elements
      .filter((element) => selection.includes(element.id))
      .map((element) => element.id)

    if (objectIds.length === 0) {
      return false
    }

    const hyperlinkEffects = createPPTObjectMetadataHyperlinkEffects({
      objectIds,
      slideId: activeSlide.id,
      source,
    })
    const accessibilityEffects = createPPTObjectMetadataAccessibilityEffects({
      objectIds,
      slideId: activeSlide.id,
      source,
    })
    const renameEffects = createPPTObjectMetadataRenameEffects({
      objectIds,
      slide: activeSlide,
      source,
    })
    const effects = [
      ...hyperlinkEffects,
      ...accessibilityEffects,
      ...renameEffects,
    ]

    if (effects.length === 0) {
      return false
    }

    if (hyperlinkEffects.length > 0) {
      setLastHyperlinkEffect(hyperlinkEffects[hyperlinkEffects.length - 1])
    }

    if (accessibilityEffects.length > 0) {
      setLastAccessibilityEffect(
        accessibilityEffects[accessibilityEffects.length - 1],
      )
    }

    setLastObjectMetadataImportEffect(createPPTObjectMetadataImportEffect({
      accessibilityEffects,
      hyperlinkEffects,
      renameEffects,
      source,
    }))

    commitDeck((current) =>
      updatePPTDeckSlide(current, activeSlide.id, (slide) => ({
        ...slide,
        elements: slide.elements.map((element) => {
          const elementHyperlinkEffects = hyperlinkEffects.filter((effect) =>
            effect.payload.objectId === element.id)
          const elementAccessibilityEffects = accessibilityEffects.filter((effect) =>
            effect.payload.objectId === element.id)
          const elementRenameEffects = renameEffects.filter((effect) =>
            effect.payload.id === 'rename-object' &&
              effect.payload.objectId === element.id)

          if (
            elementHyperlinkEffects.length === 0 &&
            elementAccessibilityEffects.length === 0 &&
            elementRenameEffects.length === 0
          ) {
            return element
          }

          return elementRenameEffects.reduce(
            (currentElement, effect) =>
              applyPPTObjectMetadataRenameCommandEffectToElement(
                currentElement,
                effect,
              ),
            elementAccessibilityEffects.reduce(
              (currentElement, effect) =>
                applyPPTObjectAccessibilityCommandEffectToElement(
                  currentElement,
                  effect,
                ),
              elementHyperlinkEffects.reduce(
                (currentElement, effect) =>
                  applyPPTObjectHyperlinkCommandEffectToElement(
                    currentElement,
                    effect,
                  ),
                element,
              ),
            ),
          )
        }),
      })))

    return true
  }

  function pastePPTObjectStateSource(source: PPTObjectStateImportSource) {
    const objectIds = activeSlide.elements
      .filter((element) => selection.includes(element.id))
      .map((element) => element.id)

    if (objectIds.length === 0) {
      return false
    }

    const effects = createPPTObjectStateImportCommandEffects({
      objectIds,
      slide: activeSlide,
      source,
    })

    if (
      !effects.visibilityEffect &&
      effects.lockEffects.length === 0
    ) {
      return false
    }

    if (effects.visibilityEffect) {
      setLastObjectVisibilityEffect(effects.visibilityEffect)
    }

    setLastObjectStateImportEffect(createPPTObjectStateImportEffect({
      objectIds,
      source,
      ...effects,
    }))

    const visibilityObjectIds = new Set(
      effects.visibilityEffect?.payload.objectIds ?? [],
    )
    const lockObjectIds = new Set(effects.lockEffects.flatMap((effect) =>
      effect.payload.id === 'lock-objects' ||
        effect.payload.id === 'unlock-objects'
        ? effect.payload.objectIds
        : []))

    setSelection(objectIds)
    commitDeck((current) =>
      updatePPTDeckSlide(current, activeSlide.id, (slide) => ({
        ...slide,
        elements: slide.elements.map((element) => {
          if (!objectIds.includes(element.id)) {
            return element
          }

          return {
            ...element,
            ...(visibilityObjectIds.has(element.id) &&
                source.state.visible !== undefined
              ? { visible: source.state.visible }
              : {}),
            ...(lockObjectIds.has(element.id) &&
                source.state.locked !== undefined
              ? { locked: source.state.locked }
              : {}),
          }
        }),
      })))

    return true
  }

  function pastePPTObjectLayerSource(source: PPTObjectLayerImportSource) {
    const objectIds = activeSlide.elements
      .filter((element) => selection.includes(element.id))
      .map((element) => element.id)

    if (objectIds.length !== 1) {
      return false
    }

    const objectId = objectIds[0]
    const effect = createPPTObjectLayerImportCommandEffect({
      objectId,
      slide: activeSlide,
      source,
    })

    if (!effect || effect.payload.id !== 'reorder-object') {
      return false
    }

    const payload = effect.payload

    setLastObjectLayerImportEffect(createPPTObjectLayerImportEffect({
      effect,
      source,
    }))
    setSelection([objectId])
    commitDeck((current) =>
      updatePPTDeckSlide(current, effect.selection.slideId, (slide) => {
        const elements = reorderPPTLayerPaneElement(
          slide.elements,
          payload.objectId,
          payload.toIndex,
        )

        return elements
          ? { ...slide, elements }
          : slide
      }))

    return true
  }

  function pastePPTObjectTransformSource(
    source: PPTObjectTransformImportSource,
  ) {
    const objectIds = activeSlide.elements
      .filter((element) =>
        selection.includes(element.id) &&
        element.locked !== true &&
        element.visible !== false)
      .map((element) => element.id)

    if (objectIds.length === 0) {
      return false
    }

    setLastObjectTransformImportEffect(createPPTObjectTransformImportEffect({
      objectIds,
      source,
    }))

    commitDeck((current) =>
      updatePPTDeckSlide(current, activeSlide.id, (slide) => ({
        ...slide,
        elements: mapPPTElementsByIds(
          slide.elements,
          objectIds,
          (element) => applyPPTObjectTransformSourceToElement(element, source),
        ),
      })))

    return true
  }

  function pastePPTTextAutoFitSource(source: PPTTextAutoFitImportSource) {
    const effects = selectedElements
      .filter((element): element is PPTTextElement =>
        isPPTTextElement(element) &&
          element.locked !== true &&
          element.visible !== false)
      .map((element) =>
        getPPTTextAutoFitCommandEffect({
          element,
          handle: source.handle,
          hasOverflow: textOverflowById[element.id] === true,
          slideId: activeSlide.id,
        }))
      .filter((effect): effect is SlideEditTextAutoFitHostCommandEffect<string, string> =>
        effect !== null)

    if (effects.length === 0) {
      return false
    }

    const objectIds = effects.map((effect) => effect.payload.objectId)

    setLastTextAutoFitEffect(effects[effects.length - 1])
    setLastTextAutoFitImportEffect(createPPTTextAutoFitImportEffect({
      effects,
      source,
    }))
    commitDeck((current) =>
      updatePPTDeckSlide(current, activeSlide.id, (slide) => ({
        ...slide,
        elements: slide.elements.map((element) => {
          const effect = effects.find((effect) =>
            effect.payload.objectId === element.id)

          return effect && isPPTTextElement(element)
            ? applyPPTTextAutoFitCommandEffectToElement(element, effect)
            : element
        }),
      })))
    setTextOverflowById((current) => ({
      ...current,
      ...Object.fromEntries(objectIds.map((objectId) => [objectId, false])),
    }))

    return true
  }

  function pastePPTTextBodySource(source: PPTTextBodyImportSource) {
    const objectIds = activeSlide.elements
      .filter((element) =>
        selection.includes(element.id) &&
        isPPTTextElement(element) &&
        element.locked !== true &&
        element.visible !== false)
      .map((element) => element.id)

    if (objectIds.length === 0) {
      return false
    }

    setLastTextBodyImportEffect(createPPTTextBodyImportEffect({
      objectIds,
      source,
    }))

    commitDeck((current) =>
      updatePPTDeckSlide(current, activeSlide.id, (slide) => ({
        ...slide,
        elements: mapPPTElementsByIds(
          slide.elements,
          objectIds,
          (element) => isPPTTextElement(element)
            ? {
                ...element,
                textBody: clonePPTTextBody(source.textBody),
              }
            : element,
        ),
      })))

    return true
  }

  function pastePPTTableRowsSource(source: PPTTableRowsImportSource) {
    const rows = normalizePPTTableRows(source.rows)
    const objectIds = activeSlide.elements
      .filter((element) =>
        selection.includes(element.id) &&
        element.kind === 'table' &&
        element.locked !== true &&
        element.visible !== false)
      .map((element) => element.id)

    if (objectIds.length === 0) {
      return false
    }

    setLastTableRowsImportEffect(createPPTTableRowsImportEffect({
      objectIds,
      rows,
      source,
    }))

    commitDeck((current) =>
      updatePPTDeckSlide(current, activeSlide.id, (slide) => ({
        ...slide,
        elements: mapPPTElementsByIds(
          slide.elements,
          objectIds,
          (element) => element.kind === 'table'
            ? { ...element, rows }
            : element,
        ),
      })))

    return true
  }

  function pastePPTCommentSource(source: PPTCommentImportSource) {
    const objectIds = activeSlide.elements
      .filter((element) =>
        selection.includes(element.id) &&
        element.kind === 'comment' &&
        element.locked !== true &&
        element.visible !== false)
      .map((element) => element.id)

    if (objectIds.length === 0) {
      return false
    }

    setLastCommentImportEffect(createPPTCommentImportEffect({
      objectIds,
      source,
    }))

    commitDeck((current) =>
      updatePPTDeckSlide(current, activeSlide.id, (slide) => ({
        ...slide,
        elements: mapPPTElementsByIds(
          slide.elements,
          objectIds,
          (element) => element.kind === 'comment'
            ? applyPPTCommentImportSourceToElement(element, source)
            : element,
        ),
      })))

    return true
  }

  function pastePPTImageReplaceSource(source: PPTImageReplaceImportSource) {
    const imageElements = activeSlide.elements
      .filter((element): element is PPTImage =>
        selection.includes(element.id) &&
          element.kind === 'image' &&
          element.locked !== true &&
          element.visible !== false)

    if (imageElements.length !== 1) {
      return false
    }

    const elementId = imageElements[0].id
    const replaceImage = (image: PPTImageReplaceImportImageSource) => {
      const effect = createPPTImageReplaceCommandEffect({
        elementId,
        slideId: activeSlide.id,
        source: image,
      })

      setLastImageReplaceEffect(effect)
      setLastImageReplaceImportEffect(createPPTImageReplaceImportEffect({
        effect,
        source: {
          ...source,
          image,
        },
      }))
      commitPPTImageReplaceEffect(effect)
      finishPPTImageReplace(elementId)
    }

    if (source.resolveNaturalSize) {
      void resolvePPTImageSourceNaturalSize(source.image).then((image) => {
        replaceImage({
          ...source.image,
          ...image,
        })
      })
    } else {
      replaceImage(source.image)
    }

    return true
  }

  function pastePPTImageCropSource(source: PPTImageCropImportSource) {
    const objectIds = activeSlide.elements
      .filter((element) =>
        selection.includes(element.id) && element.kind === 'image')
      .map((element) => element.id)

    if (objectIds.length === 0) {
      return false
    }

    const effects = createPPTImageCropImportCommandEffects({
      objectIds,
      slideId: activeSlide.id,
      source,
    })

    if (effects.length === 0) {
      return false
    }

    setLastImageCropEffect(effects[effects.length - 1])
    setLastImageCropImportEffect(createPPTImageCropImportEffect({
      effects,
      source,
    }))

    commitDeck((current) =>
      updatePPTDeckSlide(current, activeSlide.id, (slide) => ({
        ...slide,
        elements: slide.elements.map((element) => {
          if (element.kind !== 'image') {
            return element
          }

          const elementEffects = effects.filter((effect) =>
            effect.payload.objectId === element.id)

          if (elementEffects.length === 0) {
            return element
          }

          return elementEffects.reduce(
            (currentElement, effect) =>
              applyPPTImageCropCommandEffectToElement(currentElement, effect),
            element,
          )
        }),
      })))

    return true
  }

  function pastePPTShapeStyleSource(source: PPTShapeStyleImportSource) {
    const shapeElements = selectedElements.filter((element): element is PPTShape =>
      element.kind === 'shape')

    if (shapeElements.length === 0) {
      return false
    }

    const sourceElement = shapeElements[0]
    const stroke = source.shape.stroke ?? sourceElement.stroke
    const categories: PPTStyleClipboardCategory[] = ['object', 'shape']

    if (stroke) {
      categories.push('stroke')
    }

    const styleClipboard: PPTStyleClipboard = {
      categories,
      object: {
        opacity: getPPTElementOpacity(sourceElement),
        shadow: hasPPTElementShadow(sourceElement)
          ? clonePPTElementShadow(getPPTElementShadow(sourceElement))
          : null,
      },
      shape: {
        cornerRadius: source.shape.cornerRadius ??
          getPPTShapeCornerRadius(sourceElement),
        fill: source.shape.fill
          ? clonePPTFill(source.shape.fill)
          : clonePPTFill(sourceElement.fill),
        ...(stroke
          ? { stroke: clonePPTStroke(stroke) }
          : {}),
      },
      sourceId: 'ppt-shape-style-json',
      sourceKind: sourceElement.kind,
      type: 'slide-style-clipboard',
    }

    if (styleClipboard.shape?.stroke) {
      styleClipboard.stroke = clonePPTStroke(styleClipboard.shape.stroke)
    }

    const effect = createSlideEditStyleClipboardPasteCommandEffect({
      clipboard: createPPTStyleClipboardDescriptor(activeSlide.id, styleClipboard),
      targetSlideId: activeSlide.id,
      targets: getPPTStyleClipboardTargetInputs(shapeElements),
    })

    if (!effect) {
      return false
    }

    setStyleClipboard(styleClipboard)
    setLastStyleClipboardEffect(effect)
    setLastShapeStyleImportEffect(createPPTShapeStyleImportEffect({
      effect,
      source,
    }))

    const categoryApplicationsByObjectId = new Map(
      effect.payload.categoryApplications.map((application) => [
        application.objectId,
        application.appliedCategoryIds,
      ]),
    )

    commitDeck((current) =>
      updatePPTDeckSlide(current, activeSlide.id, (slide) => ({
        ...slide,
        elements: mapPPTElementsByIds(
          slide.elements,
          effect.payload.categoryApplications.map((application) =>
            application.objectId),
          (element) => {
            const appliedCategoryIds = categoryApplicationsByObjectId.get(element.id)

            return appliedCategoryIds
              ? applyPPTStyleClipboardToElement(
                  element,
                  styleClipboard,
                  appliedCategoryIds,
                )
              : element
          },
        ),
      })))

    return true
  }

  function pastePPTLineStyleSource(source: PPTLineStyleImportSource) {
    const lineStyleElements = selectedElements.filter(isPPTLineStyleTargetElement)

    if (lineStyleElements.length === 0) {
      return false
    }

    const sourceElement = lineStyleElements[0]
    const currentStroke = getPPTElementStroke(sourceElement) ??
      normalizePPTStroke({})
    const stroke = normalizePPTStroke({
      ...currentStroke,
      ...source.stroke,
    })
    const styleClipboard: PPTStyleClipboard = {
      categories: ['object', 'stroke'],
      object: {
        opacity: getPPTElementOpacity(sourceElement),
        shadow: hasPPTElementShadow(sourceElement)
          ? clonePPTElementShadow(getPPTElementShadow(sourceElement))
          : null,
      },
      sourceId: 'ppt-line-style-json',
      sourceKind: sourceElement.kind,
      stroke,
      type: 'slide-style-clipboard',
    }

    const effect = createSlideEditStyleClipboardPasteCommandEffect({
      clipboard: createPPTStyleClipboardDescriptor(activeSlide.id, styleClipboard),
      targetSlideId: activeSlide.id,
      targets: getPPTStyleClipboardTargetInputs(lineStyleElements),
    })

    if (!effect) {
      return false
    }

    setStyleClipboard(styleClipboard)
    setLastStyleClipboardEffect(effect)
    setLastLineStyleImportEffect(createPPTLineStyleImportEffect({
      effect,
      source,
    }))

    const categoryApplicationsByObjectId = new Map(
      effect.payload.categoryApplications.map((application) => [
        application.objectId,
        application.appliedCategoryIds,
      ]),
    )

    commitDeck((current) =>
      updatePPTDeckSlide(current, activeSlide.id, (slide) => ({
        ...slide,
        elements: mapPPTElementsByIds(
          slide.elements,
          effect.payload.categoryApplications.map((application) =>
            application.objectId),
          (element) => {
            const appliedCategoryIds = categoryApplicationsByObjectId.get(element.id)

            return appliedCategoryIds
              ? applyPPTStyleClipboardToElement(
                  element,
                  styleClipboard,
                  appliedCategoryIds,
                )
              : element
          },
        ),
      })))

    return true
  }

  function pastePPTTextStyleSource(source: PPTTextStyleImportSource) {
    const textElements = selectedElements.filter(isPPTTextElement)

    if (textElements.length === 0) {
      return false
    }

    const sourceElement = textElements[0]
    const categories: PPTStyleClipboardCategory[] = ['object']
    const currentTextStyle = getPPTTextElementStyle(sourceElement)
    const currentTextInset = getPPTTextElementInset(sourceElement)
    let nextTextStyle: PPTTextStyle | undefined

    if (source.text) {
      const { textInset, ...textStyleFields } = source.text

      nextTextStyle = clonePPTTextStyle({
        ...currentTextStyle,
        ...textStyleFields,
        ...(textInset
          ? {
              textInset: {
                ...currentTextInset,
                ...textInset,
              },
            }
          : {}),
      })
    }

    const nextParagraph = getPPTTextStyleImportParagraph(
      sourceElement,
      source,
    )

    if (nextTextStyle) {
      categories.push('text')
    }

    if (nextParagraph) {
      categories.push('paragraph')
    }

    const styleClipboard: PPTStyleClipboard = {
      categories,
      object: {
        opacity: getPPTElementOpacity(sourceElement),
        shadow: hasPPTElementShadow(sourceElement)
          ? clonePPTElementShadow(getPPTElementShadow(sourceElement))
          : null,
      },
      ...(nextParagraph ? { paragraph: nextParagraph } : {}),
      sourceId: 'ppt-text-style-json',
      sourceKind: sourceElement.kind,
      ...(nextTextStyle ? { text: nextTextStyle } : {}),
      type: 'slide-style-clipboard',
    }

    const effect = createSlideEditStyleClipboardPasteCommandEffect({
      clipboard: createPPTStyleClipboardDescriptor(activeSlide.id, styleClipboard),
      targetSlideId: activeSlide.id,
      targets: getPPTStyleClipboardTargetInputs(textElements),
    })

    if (!effect) {
      return false
    }

    setStyleClipboard(styleClipboard)
    setLastStyleClipboardEffect(effect)
    setLastTextStyleImportEffect(createPPTTextStyleImportEffect({
      effect,
      source,
    }))

    const categoryApplicationsByObjectId = new Map(
      effect.payload.categoryApplications.map((application) => [
        application.objectId,
        application.appliedCategoryIds,
      ]),
    )

    commitDeck((current) =>
      updatePPTDeckSlide(current, activeSlide.id, (slide) => ({
        ...slide,
        elements: mapPPTElementsByIds(
          slide.elements,
          effect.payload.categoryApplications.map((application) =>
            application.objectId),
          (element) => {
            const appliedCategoryIds = categoryApplicationsByObjectId.get(element.id)

            return appliedCategoryIds
              ? applyPPTStyleClipboardToElement(
                  element,
                  styleClipboard,
                  appliedCategoryIds,
                )
              : element
          },
        ),
      })))

    return true
  }

  function pastePPTDeckFallbackHTMLSource(source: PPTDeckHTMLFallbackSource) {
    if (source.slides.length === 0) {
      return false
    }

    commitDeck((current) => {
      const targetSlideId = current.slides.some((slide) =>
        slide.id === activeSlide.id)
        ? activeSlide.id
        : current.slides.at(-1)?.id

      if (!targetSlideId) {
        return current
      }

      const importedSlides = createPPTSlidesFromDeckFallbackHTMLSource(
        current,
        source,
      )

      if (importedSlides.length === 0) {
        return current
      }

      let slides = current.slides
      let anchorSlideId = targetSlideId

      for (const slide of importedSlides) {
        const result = insertPPTSlideAtTargetPlacement({
          placement: 'after',
          slide,
          slides,
          targetSlideId: anchorSlideId,
        })

        if (!result) {
          return current
        }

        slides = result.items
        anchorSlideId = slide.id
      }

      setLastDeckHTMLImportEffect(createPPTDeckHTMLFallbackImportEffect({
        importedSlides,
        source,
      }))
      selectSlide(importedSlides[0].id)

      return {
        ...current,
        slides,
      }
    })

    return true
  }

  function pastePPTDeckMarkdownOutlineSource(
    source: PPTDeckMarkdownOutlineSource,
  ) {
    if (source.slides.length === 0) {
      return false
    }

    commitDeck((current) => {
      const targetSlideId = current.slides.some((slide) =>
        slide.id === activeSlide.id)
        ? activeSlide.id
        : current.slides.at(-1)?.id

      if (!targetSlideId) {
        return current
      }

      const createSlideId = createPPTCanvasSequentialIdFactory({
        existingIds: current.slides.map((slide) => slide.id),
        startIndex: current.slides.length + 1,
      })
      const importedSlides = createPPTDeckMarkdownOutlineSlides({
        createSlideId,
        source,
      })

      if (importedSlides.length === 0) {
        return current
      }

      let slides = current.slides
      let anchorSlideId = targetSlideId

      for (const slide of importedSlides) {
        const result = insertPPTSlideAtTargetPlacement({
          placement: 'after',
          slide,
          slides,
          targetSlideId: anchorSlideId,
        })

        if (!result) {
          return current
        }

        slides = result.items
        anchorSlideId = slide.id
      }

      setLastDeckMarkdownOutlineImportEffect(
        createPPTDeckMarkdownOutlineImportEffect({
          importedSlides,
          source,
        }),
      )
      selectSlide(importedSlides[0].id)

      return {
        ...current,
        slides,
      }
    })

    return true
  }

  function deleteActiveSlide() {
    if (deck.slides.length <= 1) {
      return
    }

    commitDeck((current) => {
      if (current.slides.length <= 1) {
        return current
      }

      const index = current.slides.findIndex((slide) => slide.id === activeSlide.id)
      const slides = deletePPTCanvasSelectionItems({
        getItemId: (slide) => slide.id,
        items: current.slides,
        selection: [activeSlide.id],
      })
      const nextSlide = slides[Math.min(Math.max(index, 0), slides.length - 1)]

      if (nextSlide) {
        selectSlide(nextSlide.id)
      }

      return {
        ...current,
        slides,
      }
    })
  }

  function moveActiveSlide(delta: -1 | 1) {
    commitDeck((current) => {
      const index = current.slides.findIndex((slide) => slide.id === activeSlide.id)
      const targetIndex = index + delta

      if (index < 0 || targetIndex < 0 || targetIndex >= current.slides.length) {
        return current
      }

      const targetSlide = current.slides[targetIndex]
      const result = targetSlide
        ? movePPTSlideToTargetPlacement({
            placement: delta > 0 ? 'after' : 'before',
            slideId: activeSlide.id,
            slides: current.slides,
            targetSlideId: targetSlide.id,
          })
        : null

      if (!result) {
        return current
      }

      return {
        ...current,
        slides: result.items,
      }
    })
  }

  function getSlideThumbDropPlacement(
    event: ReactDragEvent<HTMLButtonElement>,
  ): PPTSlideDropPlacement {
    const localGeometry = getPPTCanvasPointerLocalGeometry({
      event,
      target: event.currentTarget,
    })

    if (!localGeometry) {
      return 'after'
    }

    return localGeometry.point.y < localGeometry.rect.height / 2
      ? 'before'
      : 'after'
  }

  function clearSlideDragState({
    suppressClick = false,
  }: {
    suppressClick?: boolean
  } = {}) {
    if (suppressClick) {
      slideDragSuppressClickRef.current = true
      const timeout = schedulePPTCanvasTimeoutTask({
        delayMs: 120,
        task: () => {
          slideDragSuppressClickRef.current = false
        },
      })

      if (timeout === null) {
        slideDragSuppressClickRef.current = false
      }
    }

    setSlideDragState(null)
  }

  function handleSlideThumbDragStart(
    slideId: string,
    event: ReactDragEvent<HTMLButtonElement>,
  ) {
    setPPTCanvasDataTransferText({
      dataTransfer: event.dataTransfer,
      effectAllowed: 'move',
      text: slideId,
    })
    setContextMenu(null)
    setSlideDragState({ draggingSlideId: slideId })
  }

  function handleSlideThumbDragOver(
    targetSlideId: string,
    event: ReactDragEvent<HTMLButtonElement>,
  ) {
    const sourceSlideId = slideDragState?.draggingSlideId

    if (!sourceSlideId) {
      return
    }

    event.preventDefault()
    setPPTCanvasDataTransferDropEffect({
      dataTransfer: event.dataTransfer,
      dropEffect: 'move',
    })

    if (sourceSlideId === targetSlideId) {
      setSlideDragState({ draggingSlideId: sourceSlideId })
      return
    }

    setSlideDragState({
      draggingSlideId: sourceSlideId,
      dropPlacement: getSlideThumbDropPlacement(event),
      dropTargetSlideId: targetSlideId,
    })
  }

  function handleSlideThumbDrop(
    targetSlideId: string,
    event: ReactDragEvent<HTMLButtonElement>,
  ) {
    event.preventDefault()

    const sourceSlideId = slideDragState?.draggingSlideId ||
      getPPTCanvasDataTransferText({ dataTransfer: event.dataTransfer })
    const dropPlacement = slideDragState?.dropTargetSlideId === targetSlideId
      ? slideDragState.dropPlacement ?? getSlideThumbDropPlacement(event)
      : getSlideThumbDropPlacement(event)

    if (!sourceSlideId) {
      clearSlideDragState()
      return
    }

    reorderSlideByDrop(sourceSlideId, targetSlideId, dropPlacement)
    clearSlideDragState({ suppressClick: true })
  }

  function handleSlideThumbDragEnd() {
    clearSlideDragState({ suppressClick: true })
  }

  function reorderSlideByDrop(
    sourceSlideId: string,
    targetSlideId: string,
    placement: PPTSlideDropPlacement,
  ) {
    const currentDeck = deckRef.current
    const reorderResult = movePPTSlideToTargetPlacement({
      placement,
      slideId: sourceSlideId,
      slides: currentDeck.slides,
      targetSlideId,
    })

    if (!reorderResult) {
      return
    }

    const effect = getSlideEditRailPointerCommandEffect({
      slideId: sourceSlideId,
      slideOrder: currentDeck.slides.map((slide) => slide.id),
      toIndex: reorderResult.toIndex,
      type: 'thumbnail-drop',
    })

    if (!effect) {
      return
    }

    setLastSlideRailCommandEffect(effect)

    commitDeck((current) => {
      const result = movePPTSlideToTargetPlacement({
        placement,
        slideId: sourceSlideId,
        slides: current.slides,
        targetSlideId,
      })

      if (!result) {
        return current
      }

      return {
        ...current,
        slides: result.items,
      }
    })
  }

  function activatePPTCreationTool(tool: PPTCreationTool) {
    setCreationTool((current) => arePPTCreationToolsEqual(current, tool) ? null : tool)
    setLineCreationMode(null)
    setIsPanToolActive(false)
    setIsLaserToolActive(false)
    setLaserTrailPoints([])
    setIsEraserToolActive(false)
    setEditingId(null)
    setContextMenu(null)
  }

  function activateLineCreationMode(mode: LineCreationMode) {
    setLineCreationMode((current) => current === mode ? null : mode)
    setCreationTool(null)
    setIsPanToolActive(false)
    setIsLaserToolActive(false)
    setLaserTrailPoints([])
    setIsEraserToolActive(false)
    setEditingId(null)
    setContextMenu(null)
  }

  function activatePanTool() {
    setIsPanToolActive((current) => !current)
    setCreationTool(null)
    setLineCreationMode(null)
    setIsLaserToolActive(false)
    setLaserTrailPoints([])
    setIsEraserToolActive(false)
    setEditingId(null)
    setContextMenu(null)
  }

  function activateLaserTool() {
    const shouldActivate = !isLaserToolActive

    setIsLaserToolActive(shouldActivate)
    setLaserTrailPoints([])
    setInteraction((current) => current?.kind === 'laser' ? null : current)
    setCreationTool(null)
    setLineCreationMode(null)
    setIsPanToolActive(false)
    setIsEraserToolActive(false)
    setEditingId(null)
    setContextMenu(null)
  }

  function activateEraserTool() {
    const shouldActivate = !isEraserToolActive

    setIsEraserToolActive(shouldActivate)
    setInteraction((current) => current?.kind === 'erase' ? null : current)
    setCreationTool(null)
    setLineCreationMode(null)
    setIsPanToolActive(false)
    setIsLaserToolActive(false)
    setLaserTrailPoints([])
    setEditingId(null)
    setContextMenu(null)
  }

  function activatePPTToolShortcut(tool: Tool) {
    switch (tool) {
      case 'arrow':
        activateLineCreationMode('arrow')
        return true
      case 'comment':
        activatePPTCreationTool({ kind: 'comment' })
        return true
      case 'diamond':
      case 'ellipse':
      case 'rect':
        activatePPTCreationTool({ kind: 'shape', shape: tool })
        return true
      case 'eraser':
        activateEraserTool()
        return true
      case 'highlight':
      case 'marker':
      case 'pen':
        activatePPTCreationTool({ kind: 'freeform', tool })
        return true
      case 'laser':
        activateLaserTool()
        return true
      case 'pan':
        activatePanTool()
        return true
      case 'section':
        activatePPTCreationTool({ kind: 'section' })
        return true
      case 'select':
        activateSelectTool()
        return true
      case 'sticky':
        activatePPTCreationTool({ kind: 'sticky' })
        return true
      case 'text':
        activatePPTCreationTool({ kind: 'text' })
        return true
      default:
        return false
    }
  }

  function runPPTClipboardImportAction(action: PPTClipboardImportAction) {
    switch (action.kind) {
      case 'image-file-batch':
        void insertPPTImageFiles(action.files)
        return true
      case 'image-file':
        void insertPPTImageFile(action.file)
        return true
      case 'table-file-batch':
        void insertPPTTableFiles(action.files)
        return true
      case 'table-file':
        void insertPPTTableFile(action.file)
        return true
      case 'image-source':
        if (action.resolveNaturalSize) {
          void resolvePPTImageSourceNaturalSize(action.source).then((source) => {
            insertPPTImageSource(source)
          })
        } else {
          insertPPTImageSource(action.source)
        }
        return true
      case 'fallback-html-selection-source':
        insertPPTFallbackHTMLSelectionSource(action.source)
        return true
      case 'fallback-html-image-source':
        insertPPTFallbackHTMLImageSource(action.source)
        return true
      case 'fallback-html-shape-source':
        insertPPTFallbackHTMLShapeSource(action.source)
        return true
      case 'fallback-html-table-source':
        insertPPTFallbackHTMLTableSource(action.source)
        return true
      case 'fallback-html-text-source':
        insertPPTFallbackHTMLTextSource(action.source)
        return true
      case 'table-source':
        insertPPTTableSource(action.source)
        return true
      case 'media-source':
        return insertPPTMediaSource(action.source)
      case 'rich-text-source':
        return insertPPTRichTextPasteSource(action.source)
      case 'text-source':
        return insertPPTTextPasteSource(action.text)
    }
  }

  function runPPTStageDropImportAction(
    action: PPTStageDropImportAction,
    point: Point,
  ) {
    switch (action.kind) {
      case 'image-file-batch':
        void insertPPTImageFiles(action.files, point)
        return true
      case 'image-file':
        void insertPPTImageFile(action.file, point)
        return true
      case 'table-file-batch':
        void insertPPTTableFiles(action.files, point)
        return true
      case 'table-file':
        void insertPPTTableFile(action.file, point).then((inserted) => {
          if (!inserted && action.fallbackSource) {
            insertPPTTableSource(action.fallbackSource, point)
          }
        })
        return true
      case 'table-source':
        insertPPTTableSource(action.source, point)
        return true
      case 'media-source':
        insertPPTMediaSource(action.source, point)
        return true
    }
  }

  function insertPPTImageSource(
    source: PPTImageImportSource,
    center = getPPTViewportCenter(),
  ) {
    insertPPTImageSources([source], center)
  }

  function insertPPTImageSources(
    sources: readonly PPTImageImportSource[],
    center = getPPTViewportCenter(),
  ) {
    if (sources.length === 0) {
      return false
    }

    commitDeck((current) => updatePPTDeckSlide(current, activeSlide.id, (slide) => {
      const elements = createPPTImportedImageElements({
        center,
        createId: createPPTElementIdFactory(slide),
        sources,
      })
      const lastElement = elements[elements.length - 1]
      const lastSource = sources[sources.length - 1]

      if (!lastElement || !lastSource) {
        return slide
      }

      setLastImageImportEffect(createPPTImageImportEffect({
        batch: {
          count: elements.length,
          names: elements.map((element) => element.name),
        },
        element: lastElement,
        source: lastSource,
      }))
      setSelection(elements.map((element) => element.id))
      setEditingId(null)
      setLineCreationMode(null)
      setCreationTool(null)
      setIsPanToolActive(false)
      setIsLaserToolActive(false)
      setLaserTrailPoints([])
      setIsEraserToolActive(false)
      setContextMenu(null)

      return {
        ...slide,
        elements: [...slide.elements, ...elements],
      }
    }))

    return true
  }

  async function insertPPTImageFile(
    file: Blob & { name?: string },
    center = getPPTViewportCenter(),
  ) {
    const source = await readPPTImageFileSource(file)

    if (!source) {
      return false
    }

    insertPPTImageSource(source, center)
    return true
  }

  async function insertPPTImageFiles(
    files: readonly (Blob & { name?: string })[],
    center = getPPTViewportCenter(),
  ) {
    const sources: PPTImageImportSource[] = []

    for (const source of await Promise.all(files.map(readPPTImageFileSource))) {
      if (source) {
        sources.push(source)
      }
    }

    return insertPPTImageSources(sources, center)
  }

  async function pastePPTClipboardImage() {
    const source = await readPPTClipboardImageSource()

    if (!source) {
      return false
    }

    insertPPTImageSource(source)
    return true
  }

  function insertPPTFallbackHTMLSelectionSource(
    source: PPTFallbackHTMLSelectionSource,
    center = getPPTViewportCenter(),
  ) {
    commitDeck((current) => updatePPTDeckSlide(current, activeSlide.id, (slide) => {
      const elements = createPPTFallbackHTMLSelectionElements({
        center,
        createId: createPPTElementIdFactory(slide),
        source,
      })

      setLastFallbackHTMLImportEffect(createPPTFallbackHTMLSelectionImportEffect({
        elements,
        source,
      }))
      setSelection(elements.map((element) => element.id))
      setEditingId(null)
      setLineCreationMode(null)
      setCreationTool(null)
      setIsPanToolActive(false)
      setIsLaserToolActive(false)
      setLaserTrailPoints([])
      setIsEraserToolActive(false)
      setContextMenu(null)

      return {
        ...slide,
        elements: [...slide.elements, ...elements],
      }
    }))
  }

  function insertPPTFallbackHTMLImageSource(
    source: PPTFallbackHTMLImageSource,
    center = getPPTViewportCenter(),
  ) {
    commitDeck((current) => updatePPTDeckSlide(current, activeSlide.id, (slide) => {
      const element = createPPTFallbackHTMLImageElement({
        center,
        createId: createPPTElementIdFactory(slide),
        source,
      })

      setLastFallbackHTMLImportEffect(createPPTFallbackHTMLImportEffect({
        element,
        source,
      }))
      setSelection([element.id])
      setEditingId(null)
      setLineCreationMode(null)
      setCreationTool(null)
      setIsPanToolActive(false)
      setIsLaserToolActive(false)
      setLaserTrailPoints([])
      setIsEraserToolActive(false)
      setContextMenu(null)

      return {
        ...slide,
        elements: [...slide.elements, element],
      }
    }))
  }

  function insertPPTFallbackHTMLShapeSource(
    source: PPTFallbackHTMLShapeSource,
    center = getPPTViewportCenter(),
  ) {
    commitDeck((current) => updatePPTDeckSlide(current, activeSlide.id, (slide) => {
      const element = createPPTFallbackHTMLShapeElement({
        center,
        createId: createPPTElementIdFactory(slide),
        source,
      })

      setLastFallbackHTMLImportEffect(createPPTFallbackHTMLImportEffect({
        element,
        source,
      }))
      setSelection([element.id])
      setEditingId(null)
      setLineCreationMode(null)
      setCreationTool(null)
      setIsPanToolActive(false)
      setIsLaserToolActive(false)
      setLaserTrailPoints([])
      setIsEraserToolActive(false)
      setContextMenu(null)

      return {
        ...slide,
        elements: [...slide.elements, element],
      }
    }))
  }

  function insertPPTFallbackHTMLTableSource(
    source: PPTFallbackHTMLTableSource,
    center = getPPTViewportCenter(),
  ) {
    commitDeck((current) => updatePPTDeckSlide(current, activeSlide.id, (slide) => {
      const element = createPPTFallbackHTMLTableElement({
        center,
        createId: createPPTElementIdFactory(slide),
        source,
      })

      setLastFallbackHTMLImportEffect(createPPTFallbackHTMLImportEffect({
        element,
        source,
      }))
      setSelection([element.id])
      setEditingId(null)
      setLineCreationMode(null)
      setCreationTool(null)
      setIsPanToolActive(false)
      setIsLaserToolActive(false)
      setLaserTrailPoints([])
      setIsEraserToolActive(false)
      setContextMenu(null)

      return {
        ...slide,
        elements: [...slide.elements, element],
      }
    }))
  }

  function insertPPTFallbackHTMLTextSource(
    source: PPTFallbackHTMLTextSource,
    center = getPPTViewportCenter(),
  ) {
    commitDeck((current) => updatePPTDeckSlide(current, activeSlide.id, (slide) => {
      const element = createPPTFallbackHTMLTextElement({
        center,
        createId: createPPTElementIdFactory(slide),
        source,
      })

      setLastFallbackHTMLImportEffect(createPPTFallbackHTMLImportEffect({
        element,
        source,
      }))
      setSelection([element.id])
      setEditingId(null)
      setLineCreationMode(null)
      setCreationTool(null)
      setIsPanToolActive(false)
      setIsLaserToolActive(false)
      setLaserTrailPoints([])
      setIsEraserToolActive(false)
      setContextMenu(null)

      return {
        ...slide,
        elements: [...slide.elements, element],
      }
    }))
  }

  function replacePPTImageSource(
    elementId: string,
    source: PPTImageReplaceImportImageSource,
  ) {
    const effect = createPPTImageReplaceCommandEffect({
      elementId,
      slideId: activeSlide.id,
      source,
    })

    setLastImageReplaceEffect(effect)
    commitPPTImageReplaceEffect(effect)
    finishPPTImageReplace(elementId)
  }

  function createPPTImageReplaceCommandEffect({
    elementId,
    slideId,
    source,
  }: {
    elementId: string
    slideId: string
    source: PPTImageReplaceImportImageSource
  }) {
    return getSlideEditObjectImageReplaceCommandEffect({
      id: 'replace-object-image',
      objectId: elementId,
      slideId,
      source: {
        altText: source.altText ?? source.name,
        mimeType: source.mimeType,
        name: source.name,
        naturalHeight: source.naturalHeight,
        naturalWidth: source.naturalWidth,
        src: source.dataUrl,
      },
    })
  }

  function commitPPTImageReplaceEffect(
    effect: SlideEditObjectImageReplaceHostCommandEffect<string, string>,
  ) {
    commitDeck((current) => updatePPTDeckSlide(current, effect.selection.slideId, (slide) => ({
      ...slide,
      elements: mapPPTElementsByIds(slide.elements, [effect.payload.objectId], (element) => {
        if (element.kind !== 'image') {
          return element
        }

        return applyPPTImageReplaceCommandEffectToElement(element, effect)
      }),
    })))
  }

  function finishPPTImageReplace(elementId: string) {
    setSelection([elementId])
    setEditingId(null)
    setLineCreationMode(null)
    setCreationTool(null)
    setIsPanToolActive(false)
    setIsLaserToolActive(false)
    setLaserTrailPoints([])
    setIsEraserToolActive(false)
    setContextMenu(null)
  }

  async function replacePPTImageFile(
    elementId: string,
    file: Blob & { name?: string },
  ) {
    const source = await readPPTImageFileSource(file)

    if (!source) {
      return false
    }

    replacePPTImageSource(elementId, source)
    return true
  }

  function insertPPTTableSource(
    source: PPTTableImportSource = { format: 'default', rows: PPT_DEFAULT_TABLE_ROWS },
    center = getPPTViewportCenter(),
  ) {
    return insertPPTTableSources([source], center)
  }

  function insertPPTTableSources(
    sources: readonly PPTTableImportSource[],
    center = getPPTViewportCenter(),
  ) {
    if (sources.length === 0) {
      return false
    }

    commitDeck((current) => updatePPTDeckSlide(current, activeSlide.id, (slide) => {
      const elements = createPPTTableElements({
        center,
        createId: createPPTElementIdFactory(slide),
        sources,
      })
      const lastElement = elements[elements.length - 1]
      const lastSource = sources[sources.length - 1]

      if (!lastElement || !lastSource) {
        return slide
      }

      setLastTableImportEffect(createPPTTableImportEffect({
        batch: {
          count: elements.length,
          names: elements.map((element) => element.name),
        },
        element: lastElement,
        source: lastSource,
      }))
      setSelection(elements.map((element) => element.id))
      setEditingId(null)
      setLineCreationMode(null)
      setCreationTool(null)
      setIsPanToolActive(false)
      setIsLaserToolActive(false)
      setLaserTrailPoints([])
      setIsEraserToolActive(false)
      setContextMenu(null)

      return {
        ...slide,
        elements: [...slide.elements, ...elements],
      }
    }))

    return true
  }

  async function insertPPTTableFile(
    file: Blob & { name?: string },
    center = getPPTViewportCenter(),
  ) {
    const source = await readPPTTableFileSource(file)

    if (!source) {
      return false
    }

    insertPPTTableSource(source, center)
    return true
  }

  async function insertPPTTableFiles(
    files: readonly (Blob & { name?: string })[],
    center = getPPTViewportCenter(),
  ) {
    const sources: PPTTableImportSource[] = []

    for (const source of await Promise.all(files.map(readPPTTableFileSource))) {
      if (source) {
        sources.push(source)
      }
    }

    return insertPPTTableSources(sources, center)
  }

  function insertPPTMediaSource(
    source: PPTMediaImportSource,
    center = getPPTViewportCenter(),
    options: {
      onResult?: (result: PPTMediaImportResult) => void
    } = {},
  ) {
    commitDeck((current) => updatePPTDeckSlide(current, activeSlide.id, (slide) => {
      const result = createPPTMediaElement({
        createId: createPPTElementIdFactory(slide),
        position: center,
        source,
        viewport,
      })

      if (!result) {
        return slide
      }

      setLastMediaImport(result)
      options.onResult?.(result)
      setSelection([result.item.id])
      setEditingId(null)
      setLineCreationMode(null)
      setCreationTool(null)
      setIsPanToolActive(false)
      setIsLaserToolActive(false)
      setLaserTrailPoints([])
      setIsEraserToolActive(false)
      setContextMenu(null)

      return {
        ...slide,
        elements: [...slide.elements, result.item],
      }
    }))

    return true
  }

  function pastePPTMediaJSONSource(source: PPTMediaJSONImportSource) {
    return insertPPTMediaSource(source.source, getPPTViewportCenter(), {
      onResult: (result) => {
        setLastMediaJSONImportEffect(createPPTMediaJSONImportEffect({
          result,
          source,
        }))
      },
    })
  }

  function insertPPTTextPasteSource(
    text: string,
    center = getPPTViewportCenter(),
  ) {
    commitDeck((current) => updatePPTDeckSlide(current, activeSlide.id, (slide) => {
      const result = createPPTTextPasteElement({
        createId: createPPTElementIdFactory(slide),
        position: center,
        text,
        viewport,
      })

      if (!result) {
        return slide
      }

      setLastTextPasteImport(result)
      setSelection([result.item.id])
      setEditingId(null)
      setLineCreationMode(null)
      setCreationTool(null)
      setIsPanToolActive(false)
      setIsLaserToolActive(false)
      setLaserTrailPoints([])
      setIsEraserToolActive(false)
      setContextMenu(null)

      return {
        ...slide,
        elements: [...slide.elements, result.item],
      }
    }))

    return true
  }

  function insertPPTRichTextPasteSource(
    source: PPTRichTextPasteSource,
    center = getPPTViewportCenter(),
  ) {
    commitDeck((current) => updatePPTDeckSlide(current, activeSlide.id, (slide) => {
      const result = createPPTRichTextPasteElement({
        createId: createPPTElementIdFactory(slide),
        position: center,
        source,
        viewport,
      })

      if (!result) {
        return slide
      }

      setLastTextPasteImport(result)
      setSelection([result.item.id])
      setEditingId(null)
      setLineCreationMode(null)
      setCreationTool(null)
      setIsPanToolActive(false)
      setIsLaserToolActive(false)
      setLaserTrailPoints([])
      setIsEraserToolActive(false)
      setContextMenu(null)

      return {
        ...slide,
        elements: [...slide.elements, result.item],
      }
    }))

    return true
  }

  function handleImageInputChange(event: ReactChangeEvent<HTMLInputElement>) {
    const file = getPPTImageFileFromList(event.target.files)

    if (file) {
      void insertPPTImageFile(file)
    }

    event.target.value = ''
  }

  function deleteSelection() {
    if (!commandAvailability.delete) {
      return
    }

    commitElementCommand((slide) =>
      deletePPTCanvasCommand({
        adapter: commandAdapter,
        config: PPT_CANVAS_COMMAND_CONFIG,
        items: slide.elements,
        selection,
      }))
  }

  function duplicateSelection() {
    if (!commandAvailability.duplicate) {
      return
    }

    commitElementCommand((slide) =>
      duplicatePPTCanvasCommand({
        adapter: commandAdapter,
        config: PPT_CANVAS_COMMAND_CONFIG,
        createId: createPPTElementIdFactory(slide),
        items: slide.elements,
        selection,
      }))
  }

  function nudgeSelection(dx: number, dy: number) {
    if (!commandAvailability.nudge) {
      return
    }

    commitElementCommand((slide) => {
      const items = nudgePPTCanvasCommand({
        adapter: commandAdapter,
        config: PPT_CANVAS_COMMAND_CONFIG,
        dx,
        dy,
        items: slide.elements,
        selection,
      })

      return items ? { items, selection } : null
    })
  }

  function alignSelection(mode: PPTCanvasAlignMode) {
    if (!commandAvailability[canvasAlignModeAvailabilityKey[mode]]) {
      return
    }

    commitElementCommand((slide) => {
      if (selection.length === 1) {
        return {
          items: commandAdapter.alignSelection({
            items: slide.elements,
            mode,
            selection,
          }),
          selection,
        }
      }

      return alignPPTCanvasCommand({
        adapter: commandAdapter,
        config: PPT_CANVAS_COMMAND_CONFIG,
        items: slide.elements,
        mode,
        selection,
      })
    })
  }

  function distributeSelection(mode: PPTCanvasDistributeMode) {
    if (!commandAvailability[canvasDistributeModeAvailabilityKey[mode]]) {
      return
    }

    commitElementCommand((slide) =>
      distributePPTCanvasCommand({
        adapter: commandAdapter,
        config: PPT_CANVAS_COMMAND_CONFIG,
        items: slide.elements,
        mode,
        selection,
      }))
  }

  function groupSelection() {
    if (!commandAvailability.group) {
      return
    }

    commitElementCommand((slide) =>
      groupPPTCanvasCommand({
        adapter: commandAdapter,
        config: PPT_CANVAS_COMMAND_CONFIG,
        createId: createPPTElementIdFactory(slide),
        items: slide.elements,
        selection,
      }))
  }

  function ungroupSelection() {
    if (!commandAvailability.ungroup) {
      return
    }

    commitElementCommand((slide) =>
      ungroupPPTCanvasCommand({
        adapter: commandAdapter,
        config: PPT_CANVAS_COMMAND_CONFIG,
        items: slide.elements,
        selection,
      }))
  }

  function reorderSelection(mode: PPTCanvasReorderMode) {
    if (!commandAvailability[canvasReorderModeAvailabilityKey[mode]]) {
      return
    }

    commitElementCommand((slide) =>
      reorderPPTCanvasCommand({
        adapter: commandAdapter,
        config: PPT_CANVAS_COMMAND_CONFIG,
        items: slide.elements,
        mode,
        selection,
      }))
  }

  function lockSelectedElements() {
    if (!commandAvailability.lockSelection) {
      return
    }

    commitElementCommand((slide) =>
      lockPPTCanvasCommand({
        adapter: commandAdapter,
        config: PPT_CANVAS_COMMAND_CONFIG,
        items: slide.elements,
        selection,
      }))
  }

  function unlockAllElements() {
    if (!commandAvailability.unlockAll) {
      return
    }

    commitElementCommand((slide) =>
      unlockAllPPTCanvasCommand({
        adapter: commandAdapter,
        config: PPT_CANVAS_COMMAND_CONFIG,
        items: slide.elements,
        selection,
      }))
  }

  function copyFormatting() {
    if (!commandAvailability.copyFormatting) {
      return
    }

    const source = selectedElements[0]
    const next = source ? createPPTStyleClipboard(source) : null

    if (!next) {
      return
    }

    const effect = getSlideEditStyleClipboardCopyCommandEffect(
      createPPTStyleClipboardDescriptor(activeSlide.id, next),
    )

    setStyleClipboard(next)
    setLastStyleClipboardEffect(effect)
  }

  function pasteFormatting() {
    if (!commandAvailability.pasteFormatting || !styleClipboard) {
      return
    }

    const effect = createSlideEditStyleClipboardPasteCommandEffect({
      clipboard: createPPTStyleClipboardDescriptor(activeSlide.id, styleClipboard),
      targetSlideId: activeSlide.id,
      targets: getPPTStyleClipboardTargetInputs(selectedElements),
    })

    if (!effect) {
      return
    }

    setLastStyleClipboardEffect(effect)

    const categoryApplicationsByObjectId = new Map(
      effect.payload.categoryApplications.map((application) => [
        application.objectId,
        application.appliedCategoryIds,
      ]),
    )

    commitDeck((current) =>
      updatePPTDeckSlide(current, activeSlide.id, (slide) => ({
        ...slide,
        elements: mapPPTElementsByIds(
          slide.elements,
          effect.payload.categoryApplications.map((application) => application.objectId),
          (element) => {
            const appliedCategoryIds = categoryApplicationsByObjectId.get(element.id)

            return appliedCategoryIds
              ? applyPPTStyleClipboardToElement(
                  element,
                  styleClipboard,
                  appliedCategoryIds,
                )
              : element
          },
        ),
      })),
    )
  }

  function copySelection(operation: PPTClipboardOperation = 'copy') {
    const selected = getPPTCanvasSelectedItems({
      getItemId: (element) => element.id,
      items: activeSlide.elements,
      selection,
    })

    if (selected.length === 0) {
      return
    }

    const payload = createPPTClipboardPayload({
      objects: selected,
      operation,
      selectedObjectIds: selected.map((element) => element.id),
      sourceSlideId: activeSlide.id,
    })

    setClipboard(payload)
    setLastClipboardPasteEffect(null)
    setLastClipboardPastePositionEffect(null)
    clipboardPastePositionMemoryRef.current = {
      key: getPPTClipboardPastePositionKey(payload, activeSlide.id),
      pasteIndex: 0,
    }

    const richClipboardFallback = createPPTRichClipboardFallback({
      payload,
      selectionSvg: selectionSvgCode,
    })
    const richClipboardEffect = createPPTRichClipboardEffect(payload, {
      fallback: richClipboardFallback,
      writeMode: 'pending',
    })

    setLastRichClipboardEffect(richClipboardEffect)

    void writePPTRichClipboardPayload({
      fallback: richClipboardFallback,
      payload,
      selectionSvg: selectionSvgCode,
    }).then((writeMode) => {
      setLastRichClipboardEffect((current) =>
        current &&
        current.sourceSlideId === payload.sourceSlideId &&
        current.selectedObjectIds.join(' ') === payload.selectedObjectIds.join(' ')
          ? {
              ...current,
              writeMode,
            }
          : current)
    })
  }

  function cutSelection() {
    if (!commandAvailability.cut) {
      return
    }

    copySelection('cut')
    deleteSelection()
  }

  function pasteSelection() {
    if (!commandAvailability.paste || !clipboard) {
      return
    }

    pasteClipboardPayload(clipboard)
  }

  function pasteClipboardPayload(
    payload: PPTClipboardPayload,
    options: { importFormat?: PPTRichClipboardImportFormat } = {},
  ) {
    if (payload.objects.length === 0) {
      return false
    }

    const pasteKey = getPPTClipboardPastePositionKey(payload, activeSlide.id)
    const pasteSession = getPPTCanvasPastePositionSession({
      key: pasteKey,
      memory: clipboardPastePositionMemoryRef.current,
    })
    const pasteIndex = pasteSession.pasteIndex
    const viewportCenter = getPPTViewportCenter()
    const clipboardBounds = getPPTElementsBounds([...payload.objects])
    const pasteAnchor = getPPTCanvasPasteOffsetForBounds({
      clipboardBounds,
      pasteIndex,
      viewportCenter,
    })
    const pastePositionEffect: PPTClipboardPastePositionEffect = {
      anchor: pasteAnchor,
      clipboardBounds,
      clipboardObjectCount: payload.objects.length,
      model: PPT_PASTE_POSITION_MODEL,
      pasteIndex,
      viewportCenter,
    }
    const effect = createPPTClipboardPasteCommandEffect({
      createId: createPPTElementIdFactory(activeSlide),
      payload,
      slideFrame: {
        h: PPT_SLIDE_HEIGHT,
        w: PPT_SLIDE_WIDTH,
        x: 0,
        y: 0,
      },
      target: {
        kind: 'slide-frame-offset',
        offset: pasteAnchor,
        slideId: activeSlide.id,
      },
    })

    if (!effect) {
      return false
    }

    setClipboard(payload)
    setLastClipboardPasteEffect(effect)
    setLastClipboardPastePositionEffect(pastePositionEffect)
    if (options.importFormat) {
      const richClipboardFallback = createPPTRichClipboardFallback({
        payload,
        selectionSvg: null,
      })
      setLastRichClipboardEffect(createPPTRichClipboardEffect(payload, {
        fallback: richClipboardFallback,
        importFormat: options.importFormat,
        imported: true,
      }))
    }
    clipboardPastePositionMemoryRef.current = pasteSession.nextMemory
    setSelection([...effect.selection.objectIds])

    commitDeck((current) => updatePPTDeckSlide(current, activeSlide.id, (slide) => {
      return applyPPTClipboardPasteHostCommandEffect(slide, effect)
    }))
    return true
  }

  function selectAllElements() {
    const nextSelection = selectAllPPTCanvasCommand({
      adapter: commandAdapter,
      config: PPT_CANVAS_COMMAND_CONFIG,
      items: activeSlide.elements,
    })

    if (nextSelection) {
      setSelection(nextSelection)
    }
  }

  function selectSameTypeElements() {
    if (!commandAvailability.selectSameType) {
      return
    }

    setSelection(selectSameTypePPTSelection(activeSlide.elements, selection))
    setEditingId(null)
    setContextMenu(null)
  }

  function tidySelection() {
    if (!commandAvailability.tidySelection) {
      return
    }

    commitDeck((current) => updatePPTDeckSlide(current, activeSlide.id, (slide) => ({
      ...slide,
      elements: syncPPTLineConnections(
        tidyPPTSelectionElements(slide.elements, selection),
      ),
    })))
    setEditingId(null)
    setContextMenu(null)
  }

  function flipSelection(axis: PPTFlipAxis) {
    if (!commandAvailability.flipSelection) {
      return
    }

    commitDeck((current) => updatePPTDeckSlide(current, activeSlide.id, (slide) => {
      const elements = flipPPTSelectionElements(slide.elements, selection, axis)

      return {
        ...slide,
        elements: syncPPTLineConnections(
          elements,
          getPPTSelectedLineIds(elements, selection),
        ),
      }
    }))
    setEditingId(null)
    setContextMenu(null)
  }

  function activateSelectTool() {
    setCreationTool(null)
    setLineCreationMode(null)
    setIsPanToolActive(false)
    setIsLaserToolActive(false)
    setLaserTrailPoints([])
    setIsEraserToolActive(false)
    setEditingId(null)
    setContextMenu(null)
  }

  function runPPTSurfaceCommand(command: PPTSurfaceCommand) {
    switch (command) {
      case 'alignBottom':
        alignSelection('alignBottom')
        break
      case 'alignCenter':
        alignSelection('alignCenter')
        break
      case 'alignLeft':
        alignSelection('alignLeft')
        break
      case 'alignMiddle':
        alignSelection('alignMiddle')
        break
      case 'alignRight':
        alignSelection('alignRight')
        break
      case 'alignTop':
        alignSelection('alignTop')
        break
      case 'bringForward':
        reorderSelection('bringForward')
        break
      case 'bringToFront':
        reorderSelection('bringToFront')
        break
      case 'copyFormatting':
        copyFormatting()
        break
      case 'delete':
        deleteSelection()
        break
      case 'distributeHorizontal':
        distributeSelection('distributeHorizontal')
        break
      case 'distributeVertical':
        distributeSelection('distributeVertical')
        break
      case 'duplicate':
        duplicateSelection()
        break
      case 'flipHorizontal':
        flipSelection('horizontal')
        break
      case 'flipVertical':
        flipSelection('vertical')
        break
      case 'group':
        groupSelection()
        break
      case 'lockSelection':
        lockSelectedElements()
        break
      case 'pasteFormatting':
        pasteFormatting()
        break
      case 'selectSameType':
        selectSameTypeElements()
        break
      case 'sendBackward':
        reorderSelection('sendBackward')
        break
      case 'sendToBack':
        reorderSelection('sendToBack')
        break
      case 'tidySelection':
        tidySelection()
        break
      case 'ungroup':
        ungroupSelection()
        break
      case 'unlockAll':
        unlockAllElements()
        break
    }
  }

  function commitText(elementId: string, text: string) {
    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) =>
        replacePPTElementText(element, text),
      ),
    )
  }

  function autoFitTextElement(elementId: string, handle: ResizeHandle = 'se') {
    const element = findPPTElement(activeSlide, elementId)

    if (!element || !isPPTTextElement(element)) {
      return
    }

    const effect = getPPTTextAutoFitCommandEffect({
      element,
      handle,
      hasOverflow: textOverflowById[elementId] === true,
      slideId: activeSlide.id,
    })

    if (!effect) {
      return
    }

    setLastTextAutoFitEffect(effect)
    commitDeck((current) =>
      updatePPTDeckSlide(current, activeSlide.id, (slide) => ({
        ...slide,
        elements: mapPPTElementsByIds(slide.elements, [elementId], (element) => {
          if (
            !isPPTTextElement(element) ||
            element.locked === true ||
            element.visible === false
          ) {
            return element
          }

          return applyPPTTextAutoFitCommandEffectToElement(element, effect)
        }),
      })),
    )
    setTextOverflowById((current) => ({
      ...current,
      [elementId]: false,
    }))
  }

  function updateElementGeometry(
    elementId: string,
    field: 'h' | 'w' | 'x' | 'y',
    value: number,
  ) {
    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) => ({
        ...element,
        geometry: updatePPTElementBounds(element, {
          ...pptGeometryToBounds(element.geometry),
          [field]: value,
        }).geometry,
      })),
    )
  }

  function updateElementRotation(elementId: string, rotation: number) {
    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) => ({
        ...element,
        geometry: {
          ...element.geometry,
          rotation: normalizePPTCanvasRotationDegrees(rotation),
        },
      })),
    )
  }

  function updateElementOpacity(elementId: string, opacity: number) {
    const effect = getSlideEditObjectOpacityCommandEffect({
      fieldId: 'opacity',
      id: 'update-object-opacity',
      objectId: elementId,
      slideId: activeSlide.id,
      value: normalizeSlideEditObjectOpacity(opacity),
    })

    setLastObjectOpacityEffect(effect)

    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) => ({
        ...element,
        opacity: normalizePPTElementOpacity(effect.payload.value),
      })),
    )
  }

  function updateElementHyperlink(elementId: string, url: string) {
    const trimmedUrl = url.trim()
    const effect = getSlideEditObjectHyperlinkCommandEffect(
      trimmedUrl
        ? {
            fieldId: 'url',
            id: 'update-object-hyperlink',
            objectId: elementId,
            slideId: activeSlide.id,
            value: trimmedUrl,
          }
        : {
            id: 'remove-object-hyperlink',
            objectId: elementId,
            slideId: activeSlide.id,
          },
    )

    setLastHyperlinkEffect(effect)

    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) => {
        const hyperlink = effect.payload.id === 'update-object-hyperlink'
          ? normalizePPTElementHyperlink({ url: effect.payload.value ?? '' })
          : null

        return {
          ...element,
          hyperlink: hyperlink ?? undefined,
        }
      }),
    )
  }

  function updateElementAltText(elementId: string, altText: string) {
    const normalizedAltText = normalizePPTAltText(altText)
    const effect = getSlideEditObjectAccessibilityCommandEffect(
      normalizedAltText
        ? {
            fieldId: 'altText',
            id: 'update-object-accessibility',
            objectId: elementId,
            slideId: activeSlide.id,
            value: normalizedAltText,
          }
        : {
            id: 'remove-object-alt-text',
            objectId: elementId,
            slideId: activeSlide.id,
          },
    )

    setLastAccessibilityEffect(effect)

    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) => {
        const accessibility = effect.payload.id === 'update-object-accessibility'
          ? normalizePPTElementAccessibility({
              altText: typeof effect.payload.value === 'string'
                ? effect.payload.value
                : '',
            })
          : null

        return {
          ...element,
          accessibility: accessibility ?? undefined,
        }
      }),
    )
  }

  function updateElementShadow(
    elementId: string,
    field: PPTElementShadowUpdateField,
    value: boolean | number | string,
  ) {
    const effect = getSlideEditObjectShadowCommandEffect({
      fieldId: field,
      id: 'update-object-shadow',
      objectId: elementId,
      slideId: activeSlide.id,
      value,
    })

    setLastShadowEffect(effect)

    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) => {
        const { fieldId, value: fieldValue } = effect.payload

        if (fieldId === 'enabled') {
          return {
            ...element,
            shadow: fieldValue === true ? getPPTElementShadow(element) : undefined,
          }
        }

        return {
          ...element,
          shadow: normalizePPTElementShadow({
            ...getPPTElementShadow(element),
            [fieldId]: fieldValue,
          }),
        }
      }),
    )
  }

  function updateElementAnimation(
    elementId: string,
    field: PPTElementAnimationUpdateField,
    value: PPTElementAnimation[PPTElementAnimationUpdateField],
  ) {
    const effect = getSlideEditObjectAnimationUpdateCommandEffect(
      toSlideEditObjectAnimationCommand({
        elementId,
        field,
        slideId: activeSlide.id,
        value,
      }),
    )

    setLastObjectAnimationEffect(effect)

    commitDeck((current) =>
      updatePPTDeckSlide(current, activeSlide.id, (slide) => ({
        ...slide,
        elements: mapPPTElementsByIds(slide.elements, [elementId], (element) => {
          const { field: pptField, value: fieldValue } =
            toPPTElementAnimationUpdate(effect.payload)

          return {
            ...element,
            animation: normalizePPTElementAnimation({
              ...getPPTElementAnimation(element, slide),
              [pptField]: fieldValue,
            }, slide, element.id),
          }
        }),
      })),
    )
  }

  function updateElementTextStyle(
    elementId: string,
    field: keyof PPTTextStyle,
    value: string | number,
  ) {
    if (field === 'color' && typeof value === 'string') {
      rememberRecentColor(value)
    }

    const textFontFamilyEffect = field === 'fontFamily'
      ? getSlideEditTextFontFamilyCommandEffect({
        fieldId: 'fontFamily',
        id: 'update-text-font-family',
        objectId: elementId,
        slideId: activeSlide.id,
        value: String(value),
      }, {
        fallbackFontFamily: PPT_DEFAULT_TEXT_FONT_FAMILY,
        options: getPPTTextFontFamilyDescriptorOptions(),
      })
      : null
    const textVerticalAlignmentEffect = field === 'verticalAlign'
      ? getSlideEditTextVerticalAlignmentCommandEffect({
        fieldId: 'verticalAlignment',
        id: 'update-text-vertical-alignment',
        objectId: elementId,
        slideId: activeSlide.id,
        value: normalizeSlideEditTextVerticalAlignment(String(value)),
      })
      : null

    if (textFontFamilyEffect) {
      setLastTextFontFamilyEffect(textFontFamilyEffect)
    }

    if (textVerticalAlignmentEffect) {
      setLastTextVerticalAlignmentEffect(textVerticalAlignmentEffect)
    }

    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) => {
        if (!isPPTTextElement(element)) {
          return element
        }

        const style = {
          color: '#111827',
          fontFamily: PPT_DEFAULT_TEXT_FONT_FAMILY,
          fontSize: 24,
          verticalAlign: PPT_DEFAULT_TEXT_VERTICAL_ALIGN,
          ...element.style,
        }
        const nextValue =
          textFontFamilyEffect
            ? textFontFamilyEffect.payload.value
            : textVerticalAlignmentEffect
              ? textVerticalAlignmentEffect.payload.value
              : value

        return {
          ...element,
          style: {
            ...style,
            [field]: nextValue,
          },
        }
      }),
    )
  }

  function updateElementTextInset(
    elementId: string,
    field: PPTTextInsetField,
    value: number,
  ) {
    const effect = getSlideEditTextFrameInsetCommandEffect({
      fieldId: field,
      id: 'update-text-frame-inset',
      objectId: elementId,
      slideId: activeSlide.id,
      value,
    })

    setLastTextFrameInsetEffect(effect)

    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) => {
        if (!isPPTTextElement(element) || element.locked === true) {
          return element
        }

        const { fieldId, value: fieldValue } = effect.payload
        const inset = getPPTTextElementInset(element)

        return {
          ...element,
          style: {
            ...getPPTTextElementStyle(element),
            textInset: {
              ...inset,
              [fieldId]: normalizePPTTextInset(fieldValue),
            },
          },
        }
      }),
    )
  }

  function updateSelectedTextStyles(
    update: (style: PPTTextStyle) => PPTTextStyle,
  ) {
    if (!canFormatSelectedText) {
      return
    }

    commitDeck((current) =>
      updatePPTDeckSlide(current, activeSlide.id, (slide) => ({
        ...slide,
        elements: mapSelectedPPTTextElements(slide.elements, (element) => ({
          ...element,
          style: update(getPPTTextElementStyle(element)),
        })),
      })),
    )
  }

  function mapPPTElementsByIds(
    elements: PPTElement[],
    elementIds: readonly string[],
    mapElement: (element: PPTElement, index: number) => PPTElement,
  ) {
    return mapPPTCanvasSelectionItems({
      getItemId: (element) => element.id,
      items: elements,
      mapItem: mapElement,
      selection: elementIds,
    })
  }

  function mapSelectedPPTTextElements(
    elements: PPTElement[],
    mapTextElement: (element: PPTTextElement, index: number) => PPTTextElement,
  ) {
    return mapPPTCanvasSelectionItems({
      getItemId: (element) => element.id,
      isItemSelectable: (element) =>
        isPPTTextElement(element) &&
        element.locked !== true &&
        element.visible !== false,
      items: elements,
      mapItem: (element, index) =>
        isPPTTextElement(element) ? mapTextElement(element, index) : element,
      selection,
    })
  }

  function updateSelectedParagraphAlign(
    align: NonNullable<PPTParagraph['align']>,
  ) {
    if (!canFormatSelectedText) {
      return
    }

    commitDeck((current) =>
      updatePPTDeckSlide(current, activeSlide.id, (slide) => ({
        ...slide,
        elements: mapSelectedPPTTextElements(slide.elements, (element) => ({
          ...element,
          textBody: {
            paragraphs: element.textBody.paragraphs.map((paragraph) => ({
              ...paragraph,
              align,
            })),
          },
        })),
      })),
    )
  }

  function updateSelectedParagraphList(
    list: PPTParagraph['bullet'] | undefined,
  ) {
    if (!canFormatSelectedText) {
      return
    }

    commitDeck((current) =>
      updatePPTDeckSlide(current, activeSlide.id, (slide) => ({
        ...slide,
        elements: mapSelectedPPTTextElements(slide.elements, (element) => ({
          ...element,
          textBody: {
            paragraphs: element.textBody.paragraphs.map((paragraph) => ({
              ...paragraph,
              bullet: list,
            })),
          },
        })),
      })),
    )
  }

  function updateSelectedParagraphBullet(enabled: boolean) {
    updateSelectedParagraphList(enabled ? 'bullet' : undefined)
  }

  function toggleSelectedParagraphBullet() {
    const enabled = !areAllPPTTextElementsBulleted(selectedTextElements)

    updateSelectedParagraphBullet(enabled)
  }

  function toggleSelectedParagraphNumbered() {
    const enabled = !areAllPPTTextElementsNumbered(selectedTextElements)

    updateSelectedParagraphList(enabled ? 'numbered' : undefined)
  }

  function updateSelectedTextRunStyle(
    field: 'italic' | 'underline',
    enabled: boolean,
  ) {
    if (!canFormatSelectedText) {
      return
    }

    commitDeck((current) =>
      updatePPTDeckSlide(current, activeSlide.id, (slide) => ({
        ...slide,
        elements: mapSelectedPPTTextElements(slide.elements, (element) => ({
          ...element,
          textBody: {
            paragraphs: element.textBody.paragraphs.map((paragraph) => ({
              ...paragraph,
              runs: paragraph.runs.map((run) => ({
                ...run,
                [field]: enabled ? true : undefined,
              })),
            })),
          },
        })),
      })),
    )
  }

  function toggleSelectedTextItalic() {
    updateSelectedTextRunStyle(
      'italic',
      !areAllPPTTextRunsStyled(selectedTextElements, 'italic'),
    )
  }

  function toggleSelectedTextUnderline() {
    updateSelectedTextRunStyle(
      'underline',
      !areAllPPTTextRunsStyled(selectedTextElements, 'underline'),
    )
  }

  function toggleSelectedTextBold() {
    const isBold = selectedTextElements.length > 0 &&
      selectedTextElements.every((element) =>
        getPPTTextElementStyle(element).fontWeight === 'bold')

    updateSelectedTextStyles((style) => ({
      ...style,
      fontWeight: isBold ? 'regular' : 'bold',
    }))
  }

  function stepSelectedTextFontSize(delta: number) {
    updateSelectedTextStyles((style) => ({
      ...style,
      fontSize: clampPPTCanvasValue(
        style.fontSize + delta,
        PPT_TEXT_FONT_SIZE_MIN,
        PPT_TEXT_FONT_SIZE_MAX,
      ),
    }))
  }

  function updateSelectedTextColor(color: string) {
    rememberRecentColor(color)
    updateSelectedTextStyles((style) => ({
      ...style,
      color,
    }))
  }

  function updateElementName(elementId: string, name: string) {
    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) => ({
        ...element,
        name,
      })),
    )
  }

  function updateShapeFill(
    elementId: string,
    field: keyof PPTFill,
    value: number | string,
  ) {
    if (field === 'color' && typeof value === 'string') {
      rememberRecentColor(value)
    }

    const fillOpacityEffect = field === 'opacity'
      ? getSlideEditObjectFillOpacityCommandEffect({
        fieldId: 'fillOpacity',
        id: 'update-object-fill-opacity',
        objectId: elementId,
        slideId: activeSlide.id,
        value: normalizeSlideEditObjectFillOpacity(
          typeof value === 'number' ? value : Number(value),
        ),
      })
      : null

    if (fillOpacityEffect) {
      setLastFillOpacityEffect(fillOpacityEffect)
    }

    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) =>
        element.kind === 'shape'
          ? {
              ...element,
              fill: normalizePPTFill({
                ...element.fill,
                [field]: fillOpacityEffect?.payload.value ?? value,
              }),
            }
          : element),
    )
  }

  function updateShapeKind(elementId: string, shape: PPTShapeKind) {
    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) =>
        element.kind === 'shape'
          ? {
              ...element,
              cornerRadius: shape === 'rect' ? element.cornerRadius : undefined,
              shape,
            }
          : element),
    )
  }

  function updateShapeCornerRadius(elementId: string, cornerRadius: number) {
    const effect = getSlideEditObjectCornerRadiusCommandEffect({
      fieldId: 'cornerRadius',
      id: 'update-object-corner-radius',
      objectId: elementId,
      slideId: activeSlide.id,
      value: normalizeSlideEditObjectCornerRadius(cornerRadius),
    })

    setLastCornerRadiusEffect(effect)

    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) => {
        if (element.kind !== 'shape' || element.shape !== 'rect') {
          return element
        }

        const normalized = normalizePPTShapeCornerRadius(effect.payload.value)

        return {
          ...element,
          cornerRadius: normalized === PPT_SHAPE_CORNER_RADIUS_DEFAULT
            ? undefined
            : normalized,
        }
      }),
    )
  }

  function updateLineMarker(
    elementId: string,
    field: 'endMarker' | 'startMarker',
    value: PPTLineMarker,
  ) {
    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) => {
        if (element.kind !== 'line') {
          return element
        }

        return {
          ...element,
          [field]: value,
        }
      }),
    )
  }

  function updateLineRoute(
    elementId: string,
    route: PPTLineRoute,
  ) {
    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) => {
        if (element.kind !== 'line') {
          return element
        }

        return {
          ...element,
          route,
          routeBend: route === 'elbow'
            ? element.routeBend ?? 0.5
            : element.routeBend,
        }
      }),
    )
  }

  function updateImageFit(
    elementId: string,
    fit: PPTImageFit,
  ) {
    const effect = getSlideEditObjectImageCropCommandEffect({
      fieldId: 'fit',
      id: 'update-object-image-crop',
      objectId: elementId,
      slideId: activeSlide.id,
      value: normalizeSlideEditObjectImageCropFit(fit),
    })

    setLastImageCropEffect(effect)
    const payload = effect.payload

    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) => {
        if (element.kind !== 'image' || payload.id !== 'update-object-image-crop') {
          return element
        }

        return {
          ...element,
          fit: normalizePPTImageFit(String(payload.value)),
        }
      }),
    )
  }

  function updateImageCrop(
    elementId: string,
    field: keyof PPTImageCrop,
    value: number,
  ) {
    const effect = getSlideEditObjectImageCropCommandEffect({
      fieldId: field,
      id: 'update-object-image-crop',
      objectId: elementId,
      slideId: activeSlide.id,
      value: normalizeSlideEditObjectImageCropValue(value),
    })

    setLastImageCropEffect(effect)
    const payload = effect.payload

    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) => {
        if (element.kind !== 'image' || payload.id !== 'update-object-image-crop') {
          return element
        }

        return {
          ...element,
          crop: {
            ...getPPTImageCrop(element),
            [field]: Number(payload.value),
          },
        }
      }),
    )
  }

  function resetImageCrop(elementId: string) {
    const effect = getSlideEditObjectImageCropCommandEffect({
      id: 'reset-object-image-crop',
      objectId: elementId,
      slideId: activeSlide.id,
    })

    setLastImageCropEffect(effect)

    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) => {
        if (element.kind !== 'image' || effect.payload.id !== 'reset-object-image-crop') {
          return element
        }

        return {
          ...element,
          crop: effect.payload.crop,
          fit: normalizePPTImageFit(effect.payload.fit),
        }
      }),
    )
  }

  function updateTableRows(
    elementId: string,
    value: string,
  ) {
    const rows = normalizePPTTableRows(
      value.split(/\r?\n/).map((row) => row.split('\t')),
    )

    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) =>
        element.kind === 'table'
          ? { ...element, rows }
          : element),
    )
  }

  function updateCommentBody(
    elementId: string,
    value: string,
  ) {
    const body = normalizePPTCommentBody(value)

    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) =>
        element.kind === 'comment'
          ? {
              ...element,
              body,
              thread: getPPTCommentThreadWithBody(element, body),
            }
          : element),
    )
  }

  function updateCommentResolved(
    elementId: string,
    resolved: boolean,
  ) {
    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) =>
        element.kind === 'comment'
          ? { ...element, resolved }
          : element),
    )
  }

  function addCommentReply(
    elementId: string,
    value: string,
  ) {
    const body = normalizePPTCommentReplyBody(value)

    if (body.trim().length === 0) {
      return
    }

    const comment = findPPTElement(activeSlide, elementId)

    if (!comment || comment.kind !== 'comment') {
      return
    }

    const thread = getPPTCommentThread(comment)
    const message = createPPTCommentReplyMessage(comment, body, thread.length)
    const nextThread = [...thread, message]
    const effect = toPPTCommentThreadHostCommandEffect({
      body,
      id: 'add-comment-reply',
      messageCount: nextThread.length,
      objectId: elementId,
      slideId: activeSlide.id,
    })

    setLastCommentThreadEffect(effect)
    commitDeck((current) =>
      updatePPTDeckElement(current, effect.selection.slideId, elementId, (element) =>
        element.kind === 'comment'
          ? { ...element, thread: nextThread }
          : element),
    )
  }

  function applyObjectVisibilityCommandEffect(
    effect: PPTObjectVisibilityHostCommandEffect,
  ) {
    const objectIds = new Set(getPPTLayerPaneActualObjectIds(
      activeSlide,
      effect.payload.objectIds,
    ))

    if (objectIds.size === 0) {
      return
    }

    const visible = effect.payload.id === 'show-objects'

    setLastObjectVisibilityEffect(effect)
    setSelection([...objectIds])
    commitDeck((current) =>
      updatePPTDeckSlide(current, effect.selection.slideId, (slide) => ({
        ...slide,
        elements: mapPPTElementsByIds(slide.elements, [...objectIds], (element) => ({
          ...element,
          visible,
        })),
      })),
    )
  }

  function applyLayerPaneCommandEffect(effect: PPTLayerPaneHostCommandEffect) {
    const payload = effect.payload

    switch (payload.id) {
      case 'hide-objects':
      case 'show-objects':
      case 'lock-objects':
      case 'unlock-objects': {
        const objectIds = new Set(getPPTLayerPaneActualObjectIds(
          activeSlide,
          payload.objectIds,
        ))

        if (objectIds.size === 0) {
          return
        }

        const visible = payload.id === 'show-objects'
          ? true
          : payload.id === 'hide-objects'
            ? false
            : null
        const locked = payload.id === 'lock-objects'
          ? true
          : payload.id === 'unlock-objects'
            ? false
            : null

        setSelection([...objectIds])
        commitDeck((current) =>
          updatePPTDeckSlide(current, effect.selection.slideId, (slide) => ({
            ...slide,
            elements: mapPPTElementsByIds(slide.elements, [...objectIds], (element) => ({
              ...element,
              ...(visible === null ? {} : { visible }),
              ...(locked === null ? {} : { locked }),
            })),
          })),
        )
        return
      }
      case 'rename-object':
        updateElementName(payload.objectId, payload.name)
        return
      case 'reorder-object':
        commitDeck((current) =>
          updatePPTDeckSlide(current, effect.selection.slideId, (slide) => {
            const elements = reorderPPTLayerPaneElement(
              slide.elements,
              payload.objectId,
              payload.toIndex,
            )

            if (!elements) {
              return slide
            }

            return {
              ...slide,
              elements,
            }
          }),
        )
        setSelection(getPPTLayerPaneActualObjectIds(
          activeSlide,
          effect.selection.objectIds,
        ))
        return
      case 'select-objects': {
        const targetObjectId = payload.objectIds.at(-1)

        if (!targetObjectId) {
          return
        }

        setSelection((current) =>
          getPPTLayerPaneSelection({
            currentSelection: current,
            mode: payload.mode,
            objectIds: payload.objectIds,
            slide: activeSlide,
          }))
      }
    }
  }

  function updateElementStroke(
    elementId: string,
    field: keyof PPTStroke,
    value: string | number,
  ) {
    if (field === 'color' && typeof value === 'string') {
      rememberRecentColor(value)
    }

    const commandValue = field === 'dash'
      ? getSlideEditObjectStrokeLineStyleCommandEffect({
        fieldId: 'strokeLineStyle',
        id: 'update-object-stroke-line-style',
        objectId: elementId,
        slideId: activeSlide.id,
        value: normalizeSlideEditObjectStrokeLineStyle(
          typeof value === 'string' ? value : null,
        ),
      })
      : null

    if (commandValue) {
      setLastStrokeLineStyleEffect(commandValue)
    }

    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) => {
        if (
          element.kind !== 'shape' &&
          element.kind !== 'line' &&
          element.kind !== 'freeform'
        ) {
          return element
        }

        const stroke = normalizePPTStroke({
          color: '#111827',
          width: 2,
          ...element.stroke,
          [field]: commandValue?.payload.value ?? value,
        })

        return {
          ...element,
          stroke,
        }
      }),
    )
  }

  function applyColorSwatch(
    elementId: string,
    channel: PPTColorSwatchChannel,
    color: string,
    swatch: PPTColorSwatchSelection,
  ) {
    const normalizedColor = normalizePPTSwatchColor(color) ||
      normalizeSlideEditColorSwatchValue(color) ||
      ''
    const effect = getSlideEditColorSwatchCommandEffect({
      channelId: getPPTColorSwatchPackageChannel(channel),
      id: 'apply-color-swatch',
      objectIds: [elementId],
      slideId: activeSlide.id,
      swatch: {
        ...swatch,
        swatchId: swatch.swatchId || getSlideEditColorSwatchId({
          source: swatch.source,
          tokenId: swatch.tokenId,
          value: normalizedColor,
        }),
        value: swatch.value || normalizedColor,
      },
    })
    const nextColor = normalizePPTSwatchColor(effect.payload.swatch.value) ||
      effect.payload.swatch.value

    if (!nextColor) {
      return
    }

    setLastColorSwatchEffect(effect)

    if (channel === 'text-color') {
      updateElementTextStyle(elementId, 'color', nextColor)
      return
    }

    if (channel === 'shape-fill') {
      updateShapeFill(elementId, 'color', nextColor)
      return
    }

    updateElementStroke(elementId, 'color', nextColor)
  }

  function updateParagraphAlign(
    elementId: string,
    align: NonNullable<PPTParagraph['align']>,
  ) {
    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) => {
        if (!isPPTTextElement(element)) {
          return element
        }

        return {
          ...element,
          textBody: {
            paragraphs: element.textBody.paragraphs.map((paragraph) => ({
              ...paragraph,
              align,
            })),
          },
        }
      }),
    )
  }

  function updateParagraphList(
    elementId: string,
    list: PPTParagraph['bullet'] | undefined,
  ) {
    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) => {
        if (!isPPTTextElement(element) || element.locked === true) {
          return element
        }

        return {
          ...element,
          textBody: {
            paragraphs: element.textBody.paragraphs.map((paragraph) => ({
              ...paragraph,
              bullet: list,
            })),
          },
        }
      }),
    )
  }

  function updateParagraphBullet(
    elementId: string,
    enabled: boolean,
  ) {
    updateParagraphList(elementId, enabled ? 'bullet' : undefined)
  }

  function updateParagraphNumbered(
    elementId: string,
    enabled: boolean,
  ) {
    updateParagraphList(elementId, enabled ? 'numbered' : undefined)
  }

  function updateParagraphSpacing(
    elementId: string,
    field: PPTParagraphSpacingField,
    value: number,
  ) {
    const effect = getSlideEditTextParagraphSpacingCommandEffect(
      toSlideEditParagraphSpacingCommand({
        elementId,
        field,
        slideId: activeSlide.id,
        value,
      }),
    )

    setLastTextParagraphSpacingEffect(effect)

    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) => {
        if (!isPPTTextElement(element) || element.locked === true) {
          return element
        }

        const { field: pptField, value: fieldValue } =
          toPPTParagraphSpacingUpdate(effect.payload)

        return {
          ...element,
          textBody: {
            paragraphs: element.textBody.paragraphs.map((paragraph) => ({
              ...paragraph,
              [pptField]: fieldValue,
            })),
          },
        }
      }),
    )
  }

  function updateSlideName(name: string) {
    const effect = toPPTSlideMetadataHostCommandEffect({
      fieldId: 'name',
      id: 'update-slide-name',
      slideId: activeSlide.id,
      value: name,
    })

    commitDeck((current) =>
      updatePPTDeckSlide(current, effect.selection.slideId, (slide) =>
        applyPPTSlideMetadataHostCommandEffect(slide, effect)))
  }

  function updateSlideBackground(color: string) {
    const effect = toPPTSlideMetadataHostCommandEffect({
      fieldId: 'background',
      id: 'update-slide-background',
      slideId: activeSlide.id,
      value: {
        color,
        kind: 'solid-color',
      },
    })

    commitDeck((current) =>
      updatePPTDeckSlide(current, effect.selection.slideId, (slide) =>
        applyPPTSlideMetadataHostCommandEffect(slide, effect)))
  }

  function applySlideLayout(layoutId: string) {
    const layout = getPPTLayoutDescriptor(layoutId)
    const effect = getSlideEditLayoutApplyCommandEffect({
      existingObjectPolicy: 'preserve-existing-objects',
      layoutId: layout.layoutId,
      selectedObjectIds: selection,
      slideId: activeSlide.id,
    })

    commitDeck((current) => updatePPTDeckSlide(current, activeSlide.id, (slide) => ({
      ...slide,
      layoutId: effect.payload.layoutId,
      themeId: PPT_THEME_DESCRIPTOR.themeId,
    })))
  }

  function updateLayoutPlaceholderVisibility(
    placeholderId: string,
    isVisible: boolean,
  ) {
    const effect = toPPTLayoutPlaceholderVisibilityHostCommandEffect({
      id: 'update-placeholder-visibility',
      isVisible,
      placeholderId,
      slideId: activeSlide.id,
    }, selection)

    setLastPlaceholderVisibilityEffect(effect)
    commitDeck((current) =>
      updatePPTDeckSlide(current, effect.selection.slideId, (slide) =>
        applyPPTLayoutPlaceholderVisibilityHostCommandEffect(slide, effect)))
  }

  function updateSlideNotes(notes: string) {
    const effect = toPPTSlideMetadataHostCommandEffect({
      fieldId: 'notes',
      id: 'update-slide-notes',
      slideId: activeSlide.id,
      value: notes,
    })

    commitDeck((current) =>
      updatePPTDeckSlide(current, effect.selection.slideId, (slide) =>
        applyPPTSlideMetadataHostCommandEffect(slide, effect)))
  }

  function updateSlideTransition(
    field: PPTSlideTransitionUpdateField,
    value: PPTSlideTransition[PPTSlideTransitionUpdateField],
  ) {
    const effect = getSlideEditTransitionUpdateCommandEffect(
      getPPTSlideTransitionUpdateCommand({
        field,
        slideId: activeSlide.id,
        transition: activeSlideTransition,
        value,
      }),
    )

    applySlideTransitionCommandEffect(effect)
  }

  function applySlideTransitionCommandEffect(
    effect: PPTSlideTransitionHostCommandEffect,
  ) {
    setLastSlideTransitionEffect(effect)
    commitDeck((current) =>
      updatePPTDeckSlide(current, effect.selection.slideId, (slide) => ({
        ...slide,
        transition: applyPPTSlideTransitionUpdateCommand(
          getPPTSlideTransition(slide),
          effect.payload,
        ),
      })))
  }

  function copyHTML() {
    const effect = createPPTHTMLClipboardEffect({
      html: exportCode,
      sourceSlideId: activeSlide.id,
      writeMode: 'pending',
    })

    setLastHTMLClipboardEffect(effect)

    void writePPTHTMLClipboard({
      html: exportCode,
      sourceSlideId: activeSlide.id,
    }).then((writeMode) => {
      setLastHTMLClipboardEffect((current) =>
        current &&
        current.sourceSlideId === effect.sourceSlideId &&
        current.htmlLength === effect.htmlLength
          ? {
              ...current,
              writeMode,
            }
          : current)
    })
  }

  function downloadHTML() {
    downloadTextFile({
      content: exportCode,
      filename: 'ppt-subset.html',
      type: 'text/html;charset=utf-8',
    })
  }

  function downloadSlideSVG() {
    downloadTextFile({
      content: slideSvgCode,
      filename: `${activeSlide.id}.svg`,
      type: 'image/svg+xml;charset=utf-8',
    })
  }

  function copySlideSVG() {
    const effect = createPPTSlideSVGClipboardEffect({
      sourceSlideId: activeSlide.id,
      svg: slideSvgCode,
      writeMode: 'pending',
    })

    setLastSlideSVGClipboardEffect(effect)

    void writePPTSlideSVGClipboard({
      sourceSlideId: activeSlide.id,
      svg: slideSvgCode,
    }).then((writeMode) => {
      setLastSlideSVGClipboardEffect((current) =>
        current &&
        current.sourceSlideId === effect.sourceSlideId &&
        current.svgLength === effect.svgLength
          ? {
              ...current,
              writeMode,
            }
          : current)
    })
  }

  function copySelectionSVG() {
    if (!selectionSvgCode) {
      return
    }

    const effect = createPPTSelectionSVGClipboardEffect({
      selectedObjectIds: selection,
      sourceSlideId: activeSlide.id,
      svg: selectionSvgCode,
      writeMode: 'pending',
    })

    setLastSelectionSVGClipboardEffect(effect)

    void writePPTSelectionSVGClipboard({
      selectedObjectIds: selection,
      sourceSlideId: activeSlide.id,
      svg: selectionSvgCode,
    }).then((writeMode) => {
      setLastSelectionSVGClipboardEffect((current) =>
        current &&
        current.sourceSlideId === effect.sourceSlideId &&
        current.svgLength === effect.svgLength &&
        current.selectedObjectIds.join(' ') === effect.selectedObjectIds.join(' ')
          ? {
              ...current,
              writeMode,
            }
          : current)
    })
  }

  function copySelectedTable() {
    if (!selectedTableElement) {
      return
    }

    const html = createPPTTableClipboardHTML(selectedTableElement)
    const plainText = stringifyPPTTableRows(selectedTableElement.rows)
    const effect = createPPTTableClipboardEffect({
      html,
      objectId: selectedTableElement.id,
      plainText,
      rows: selectedTableElement.rows,
      sourceSlideId: activeSlide.id,
      writeMode: 'pending',
    })

    setLastTableClipboardEffect(effect)

    void writePPTTableClipboard({
      html,
      objectId: selectedTableElement.id,
      plainText,
      rows: selectedTableElement.rows,
      sourceSlideId: activeSlide.id,
    }).then((writeMode) => {
      setLastTableClipboardEffect((current) =>
        current &&
        current.sourceSlideId === effect.sourceSlideId &&
        current.objectId === effect.objectId &&
        current.htmlLength === effect.htmlLength &&
        current.plainTextLength === effect.plainTextLength
          ? {
              ...current,
              writeMode,
            }
          : current)
    })
  }

  function downloadSelectionSVG() {
    if (!selectionSvgCode) {
      return
    }

    downloadTextFile({
      content: selectionSvgCode,
      filename: `${activeSlide.id}-selection.svg`,
      type: 'image/svg+xml;charset=utf-8',
    })
  }

  function downloadTextFile({
    content,
    filename,
    type,
  }: {
    content: string
    filename: string
    type: string
  }) {
    downloadPPTCanvasTextFile({
      content,
      filename,
      type,
    })
  }

  function toggleTheme() {
    setTheme((current) => (current === 'dark' ? 'light' : 'dark'))
  }

  function screenToWorld(event: Pick<PointerEvent, 'clientX' | 'clientY'>) {
    return getPPTPointerProjection(event).startWorld
  }

  function getPPTPointerProjection(
    event: Pick<PointerEvent, 'clientX' | 'clientY'>,
  ) {
    return getPPTPointerStartProjection({
      event,
      stageElement: canvasStageElement,
      viewport,
    })
  }

  function getPPTPointerScreenPoint(
    event: Pick<PointerEvent, 'clientX' | 'clientY'>,
  ) {
    return canvasStageElement.getScreenPoint(event)
  }

  function getPPTViewportCenter() {
    return canvasStageElement.getViewportCenter(viewport) ?? {
      x: PPT_SLIDE_WIDTH / 2,
      y: PPT_SLIDE_HEIGHT / 2,
    }
  }

  function worldToScreen(point: Point) {
    return getPPTCanvasWorldClientPoint({
      point,
      stageElement: canvasStageElement,
      viewport,
    })
  }

  function openPPTContextMenu(x: number, y: number) {
    const viewportSize = getPPTCanvasClientViewportSize()
    const position = getPPTCanvasContextMenuPosition({
      menuSize: { height: 320, width: 220 },
      point: { x, y },
      viewportSize,
    })

    setContextMenu(position)
  }

  function openPPTContextMenuAtSelection() {
    if (!selectedBounds || selection.length === 0) {
      return false
    }

    const point = worldToScreen({
      x: selectedBounds.x + selectedBounds.w / 2,
      y: selectedBounds.y,
    })

    openPPTContextMenu(point.x, point.y)
    return true
  }

  function handleStageDragOver(event: ReactDragEvent<HTMLDivElement>) {
    if (canHandlePPTStageDropImport(event.dataTransfer)) {
      event.preventDefault()
      setPPTCanvasDataTransferDropEffect({
        dataTransfer: event.dataTransfer,
        dropEffect: 'copy',
      })
    }
  }

  function handleStageDrop(event: ReactDragEvent<HTMLDivElement>) {
    const action = getPPTStageDropImportAction(event.dataTransfer)
    setLastStageDropImportActionKind(action?.kind ?? '')

    if (!action) {
      return
    }

    const point = screenToWorld(event.nativeEvent)

    if (runPPTStageDropImportAction(action, point)) {
      event.preventDefault()
    }
  }

  function handleStageWheel(event: WheelEvent) {
    if (
      editingId ||
      isPPTCanvasKeyboardTypingTarget(event.target) ||
      isPPTWheelViewportPassthroughTarget(event.target)
    ) {
      return
    }

    const rect = canvasStageElement.getRect()

    if (!rect) {
      return
    }

    const nextViewport = getPPTCanvasWheelViewport({
      config: PPT_CANVAS_COMMAND_CONFIG,
      input: {
        ctrlKey: event.ctrlKey,
        deltaMode: event.deltaMode,
        deltaX: event.deltaX,
        deltaY: event.deltaY,
        metaKey: event.metaKey,
        shiftKey: event.shiftKey,
      },
      point: canvasStageElement.getScreenPoint(event),
      viewport,
    })

    if (!nextViewport) {
      return
    }

    event.preventDefault()
    setViewport(nextViewport)
  }

  function beginTemporaryPan(event: ReactPointerEvent<HTMLElement>) {
    if ((!isTemporaryPanActive && !isPanToolActive) || event.button !== 0) {
      return false
    }

    event.preventDefault()
    event.stopPropagation()
    capturePPTCanvasPointerFromEvent(event)
    const projection = getPPTPointerProjection(event.nativeEvent)
    const result = startPPTPointerPanInteraction({
      input: event.nativeEvent,
      startScreen: projection.startScreen,
      viewport,
    })

    setContextMenu(null)
    setInteraction(result.interaction)

    return true
  }

  function beginLaserPointer(event: ReactPointerEvent<HTMLElement>) {
    if (!isLaserToolActive || event.button !== 0) {
      return false
    }

    event.preventDefault()
    event.stopPropagation()
    capturePPTCanvasPointerFromEvent(event)

    const projection = getPPTPointerProjection(event.nativeEvent)
    const startWorld = clampPPTPointToSlide(projection.startWorld)
    const result = startPPTPointerLaserInteraction({
      config: PPT_CANVAS_COMMAND_CONFIG,
      input: event.nativeEvent,
      pointerGesture: 'laser',
      startScreen: projection.startScreen,
      startWorld,
    })

    if (!result || result.kind !== 'interaction') {
      return true
    }

    setContextMenu(null)
    setLaserTrailPoints(result.laserTrail.points)
    setInteraction(result.interaction)

    return true
  }

  function beginLineCreation(
    event: ReactPointerEvent<HTMLElement>,
    point: Point,
  ) {
    if (!lineCreationMode) {
      return false
    }

    event.preventDefault()
    event.stopPropagation()
    capturePPTCanvasPointerFromEvent(event)

    const startDeck = deckRef.current
    const startSlide = findPPTSlide(startDeck, activeSlide.id)
    const id = createPPTElementId(startSlide, 'line')
    const endMarker = lineCreationMode === 'arrow' ? 'arrow' : 'none'
    const element = createPPTLineElement({
      end: point,
      endMarker,
      id,
      name: lineCreationMode === 'arrow' ? 'Arrow' : 'Line',
      slide: startSlide,
      start: point,
    })
    const nextDeck = updatePPTDeckSlide(startDeck, activeSlide.id, (slide) => ({
      ...slide,
      elements: [...slide.elements, element],
    }))

    deckRef.current = nextDeck
    setDeck(nextDeck)
    setSelection([id])
    setInteraction({
      endMarker,
      kind: 'line-create',
      lineId: id,
      slideId: activeSlide.id,
      startDeck,
      startPoint: point,
    })

    return true
  }

  function beginFreeformCreation(
    event: ReactPointerEvent<HTMLElement>,
    point: Point,
  ) {
    if (creationTool?.kind !== 'freeform') {
      return false
    }

    event.preventDefault()
    event.stopPropagation()
    capturePPTCanvasPointerFromEvent(event)

    const startDeck = deckRef.current
    const startSlide = findPPTSlide(startDeck, activeSlide.id)
    const id = createPPTElementId(startSlide, 'freeform')
    const points = [clampPPTPointToSlide(point)]
    const style = getPPTFreeformToolStyle(creationTool.tool)
    const element = createPPTFreeformElement({
      id,
      name: style.name,
      opacity: style.opacity,
      points,
      stroke: style.stroke,
    })
    const nextDeck = updatePPTDeckSlide(startDeck, activeSlide.id, (slide) => ({
      ...slide,
      elements: [...slide.elements, element],
    }))

    deckRef.current = nextDeck
    setDeck(nextDeck)
    setSelection([id])
    setEditingId(null)
    setInteraction({
      elementId: id,
      kind: 'freeform-create',
      points,
      slideId: activeSlide.id,
      startDeck,
    })

    return true
  }

  function beginEraser(
    event: ReactPointerEvent<HTMLElement>,
    point: Point,
  ) {
    if (!isEraserToolActive) {
      return false
    }

    event.preventDefault()
    event.stopPropagation()
    capturePPTCanvasPointerFromEvent(event)

    const startDeck = deckRef.current
    const startSlide = findPPTSlide(startDeck, activeSlide.id)
    const points = [clampPPTPointToSlide(point)]
    const erasedIds = getPPTEraserHitElementIds({
      points,
      scene: createPPTCanvasScene(startSlide),
      slide: startSlide,
    })

    setContextMenu(null)
    setSelection((current) =>
      removePPTCanvasSelectionIds({ ids: erasedIds, selection: current }))
    setInteraction({
      currentPoint: points[0],
      erasedIds,
      kind: 'erase',
      points,
      slideId: activeSlide.id,
      startDeck,
      startPoint: points[0],
    })

    return true
  }

  function beginElementCreation(
    event: ReactPointerEvent<HTMLElement>,
    point: Point,
  ) {
    if (!creationTool) {
      return false
    }

    event.preventDefault()
    event.stopPropagation()
    capturePPTCanvasPointerFromEvent(event)

    const startDeck = deckRef.current
    const startSlide = findPPTSlide(startDeck, activeSlide.id)
    const id = createPPTElementId(
      startSlide,
      getPPTCreationToolIdPrefix(creationTool),
    )
    const element = createPPTElementFromCreationTool({
      current: point,
      id,
      start: point,
      tool: creationTool,
    })
    const nextDeck = updatePPTDeckSlide(startDeck, activeSlide.id, (slide) => ({
      ...slide,
      elements: [...slide.elements, element],
    }))

    deckRef.current = nextDeck
    setDeck(nextDeck)
    setSelection([id])
    setEditingId(null)
    setInteraction({
      elementId: id,
      kind: 'element-create',
      slideId: activeSlide.id,
      startDeck,
      startPoint: point,
      tool: creationTool,
    })

    return true
  }

  function focusStageShell() {
    focusPPTCanvasElement({ element: stageRef.current })
  }

  function handleElementPointerDown(
    event: ReactPointerEvent<HTMLDivElement>,
    elementId: string,
  ) {
    if (editingId || event.detail > 1) {
      return
    }

    if (event.button !== 0) {
      return
    }

    focusStageShell()

    if (beginTemporaryPan(event)) {
      return
    }

    if (beginEraser(event, screenToWorld(event.nativeEvent))) {
      return
    }

    if (beginLaserPointer(event)) {
      return
    }

    if (beginFreeformCreation(event, screenToWorld(event.nativeEvent))) {
      return
    }

    if (lineCreationMode && beginLineCreation(event, screenToWorld(event.nativeEvent))) {
      return
    }

    if (creationTool && beginElementCreation(event, screenToWorld(event.nativeEvent))) {
      return
    }

    event.preventDefault()
    event.stopPropagation()
    capturePPTCanvasPointerFromEvent(event)

    const additive = isAdditivePPTPointerInput(event)
    const pointerSelection = getPPTElementPointerSelection({
      additive,
      elementId,
      scene,
      selection,
    })
    const nextSelection = getPPTGroupPointerSelection({
      additive,
      fallbackSelection: pointerSelection.nextSelection,
      itemId: elementId,
      selection,
      slide: activeSlide,
    })
    const bounds = scene.getBounds(nextSelection)
    const hasLockedTarget = activeSlide.elements.some((element) =>
      nextSelection.includes(element.id) && element.locked === true)
    const hasHiddenTarget = activeSlide.elements.some((element) =>
      nextSelection.includes(element.id) && element.visible === false)

    if (!bounds || hasLockedTarget || hasHiddenTarget) {
      setSelection(nextSelection)
      return
    }

    let interactionBounds = bounds
    let interactionHistoryDeck: PPTDeck | undefined
    let interactionSelection = nextSelection
    let interactionStartDeck = deckRef.current

    if (event.altKey) {
      const sourceDeck = deckRef.current
      const sourceSlide = findPPTSlide(sourceDeck, activeSlide.id)
      const clones = commandAdapter.cloneSelection({
        createId: createPPTElementIdFactory(sourceSlide),
        ids: nextSelection,
        items: sourceSlide.elements,
        offset: { x: 0, y: 0 },
      })

      if (clones.length > 0) {
        const cloneIds = clones.map((clone) => clone.id)
        const liveDeck = updatePPTDeckSlide(sourceDeck, activeSlide.id, (slide) => ({
          ...slide,
          elements: [...slide.elements, ...clones],
        }))
        const liveSlide = findPPTSlide(liveDeck, activeSlide.id)
        const liveScene = createPPTCanvasScene(liveSlide)

        deckRef.current = liveDeck
        setDeck(liveDeck)
        interactionBounds = liveScene.getBounds(cloneIds) ?? bounds
        interactionHistoryDeck = sourceDeck
        interactionSelection = cloneIds
        interactionStartDeck = liveDeck
      }
    }

    setSelection(interactionSelection)
    setInteraction({
      bounds: interactionBounds,
      historyDeck: interactionHistoryDeck,
      kind: 'move',
      selection: interactionSelection,
      slideId: activeSlide.id,
      snapGuides: EMPTY_PPT_CANVAS_SNAP_GUIDES,
      startDeck: interactionStartDeck,
      startPoint: screenToWorld(event.nativeEvent),
    })
  }

  function handleElementContextMenu(
    event: ReactMouseEvent<HTMLDivElement>,
    elementId: string,
  ) {
    if (editingId || isPPTCanvasKeyboardTypingTarget(event.target)) {
      return
    }

    event.preventDefault()
    event.stopPropagation()
    setInteraction(null)
    setLineCreationMode(null)
    setCreationTool(null)
    setIsLaserToolActive(false)
    setLaserTrailPoints([])
    setIsEraserToolActive(false)

    if (!selection.includes(elementId)) {
      const nextSelection = getPPTGroupPointerSelection({
        additive: false,
        fallbackSelection: [elementId],
        itemId: elementId,
        selection: [],
        slide: activeSlide,
      })

      setSelection(nextSelection)
    }

    openPPTContextMenu(event.clientX, event.clientY)
  }

  function handleStagePointerDown(event: ReactPointerEvent<HTMLDivElement>) {
    if (editingId || event.button !== 0) {
      return
    }

    focusStageShell()
    capturePPTCanvasPointerFromEvent(event)
    const additive = isAdditivePPTPointerInput(event)
    const point = screenToWorld(event.nativeEvent)

    if (beginTemporaryPan(event)) {
      return
    }

    if (beginEraser(event, point)) {
      return
    }

    if (beginLaserPointer(event)) {
      return
    }

    if (beginFreeformCreation(event, point)) {
      return
    }

    if (lineCreationMode && beginLineCreation(event, point)) {
      return
    }

    if (creationTool && beginElementCreation(event, point)) {
      return
    }

    setInteraction({
      additive,
      baseSelection: additive ? selection : [],
      currentPoint: point,
      kind: 'marquee',
      slideId: activeSlide.id,
      startPoint: point,
    })

    if (!additive) {
      setSelection([])
    }
  }

  function handleStageContextMenu(event: ReactMouseEvent<HTMLDivElement>) {
    if (editingId || isPPTCanvasKeyboardTypingTarget(event.target)) {
      return
    }

    if (selection.length === 0) {
      setContextMenu(null)
      return
    }

    event.preventDefault()
    setInteraction(null)
    setLineCreationMode(null)
    setCreationTool(null)
    setIsLaserToolActive(false)
    setLaserTrailPoints([])
    setIsEraserToolActive(false)
    openPPTContextMenu(event.clientX, event.clientY)
  }

  function handleResizePointerDown(
    event: ReactPointerEvent<HTMLButtonElement>,
    handle: ResizeHandle,
  ) {
    event.preventDefault()
    event.stopPropagation()

    const clickId = `${activeSlide.id}:${selection.join(' ')}:${handle}`
    const point = {
      x: event.clientX,
      y: event.clientY,
    }
    const resizeHandleIntent = getPPTCanvasResizeHandleDoubleClickIntent({
      handle,
      handleId: clickId,
      lastClick: resizeHandleClickMemoryRef.current,
      point,
      time: event.timeStamp,
    })

    resizeHandleClickMemoryRef.current = resizeHandleIntent.nextClick
    setLastResizeHandleClickMemoryEffect({
      handle,
      id: clickId,
      isDoubleClick: resizeHandleIntent.isDoubleClick,
      model: PPT_POINTER_CLICK_MEMORY_MODEL,
      point,
    })

    if (resizeHandleIntent.intent?.kind === 'auto-size-selection') {
      autoSizeSelection(resizeHandleIntent.intent.handle)
      return
    }

    if (event.detail > 1) {
      return
    }

    if (!selectedBounds || !canResizeSelection) {
      return
    }

    capturePPTCanvasPointerFromEvent(event)

    setInteraction({
      bounds: selectedBounds,
      handle,
      kind: 'resize',
      selection,
      slideId: activeSlide.id,
      startDeck: deckRef.current,
    })
  }

  function handleRotatePointerDown(event: ReactPointerEvent<HTMLButtonElement>) {
    event.preventDefault()
    event.stopPropagation()
    capturePPTCanvasPointerFromEvent(event)

    if (!selectedBounds || !canResizeSelection) {
      return
    }

    const center = getPPTCanvasBoundsCenter(selectedBounds)
    const point = screenToWorld(event.nativeEvent)

    setInteraction({
      bounds: selectedBounds,
      center,
      kind: 'rotate',
      selection,
      slideId: activeSlide.id,
      startAngle: getPointAngle(center, point),
      startDeck: deckRef.current,
      startRotations: selectedElements.map((element) => ({
        elementId: element.id,
        rotation: element.geometry.rotation ?? 0,
      })),
    })
  }

  function handleLineEndpointPointerDown(
    event: ReactPointerEvent<HTMLButtonElement>,
    endpoint: 'end' | 'start',
  ) {
    event.preventDefault()
    event.stopPropagation()
    capturePPTCanvasPointerFromEvent(event)

    if (!selectedLineElement || !canResizeSelection) {
      return
    }

    setInteraction({
      endpoint,
      kind: 'line-endpoint',
      lineId: selectedLineElement.id,
      slideId: activeSlide.id,
      startDeck: deckRef.current,
    })
  }

  function handleLineRoutePointerDown(
    event: ReactPointerEvent<HTMLButtonElement>,
  ) {
    event.preventDefault()
    event.stopPropagation()
    capturePPTCanvasPointerFromEvent(event)

    if (!selectedLineElement || !canResizeSelection) {
      return
    }

    setInteraction({
      kind: 'line-route',
      lineId: selectedLineElement.id,
      slideId: activeSlide.id,
      startDeck: deckRef.current,
    })
  }

  function autoSizeSelection(handle: ResizeHandle) {
    if (selection.length === 0) {
      return
    }

    if (
      selection.length === 1 &&
      selectedElement &&
      isPPTTextElement(selectedElement)
    ) {
      autoFitTextElement(selectedElement.id, handle)
      return
    }

    commitDeck((current) => updatePPTDeckSlide(current, activeSlide.id, (slide) => ({
      ...slide,
      elements: mapPPTElementsByIds(slide.elements, selection, (element) => {
        const size = measurePPTElementAutoSize(element)

        if (!size) {
          return element
        }

        const next = pptGeometryToBounds(element.geometry)

        if (handle.includes('w') || handle.includes('e')) {
          next.w = size.w
        }

        if (handle.includes('n') || handle.includes('s')) {
          next.h = size.h
        }

        return {
          ...element,
          geometry: updatePPTElementBounds(element, next).geometry,
        }
      }),
    })))
  }

  function handlePointerMove(event: ReactPointerEvent<HTMLDivElement>) {
    if (!interaction) {
      return
    }

    if (interaction.kind === 'pan') {
      const preview = previewPPTPointerPanInteraction({
        config: PPT_CANVAS_COMMAND_CONFIG,
        currentScreen: getPPTPointerScreenPoint(event.nativeEvent),
        interaction,
      })

      if (preview.kind === 'preview') {
        setViewport(preview.viewport)
      }
      return
    }

    const point = screenToWorld(event.nativeEvent)

    if (interaction.kind === 'laser') {
      const preview = previewPPTPointerLaserInteraction({
        config: PPT_CANVAS_COMMAND_CONFIG,
        currentScreen: getPPTPointerScreenPoint(event.nativeEvent),
        currentWorld: clampPPTPointToSlide(point),
        interaction,
      })

      if (preview?.kind === 'preview') {
        setLaserTrailPoints(preview.laserTrail.points)
        setInteraction(preview.interaction)
      }
      return
    }

    if (interaction.kind === 'erase') {
      const points = getNextPPTEraserPoints(interaction.points, point)
      const startSlide = findPPTSlide(interaction.startDeck, interaction.slideId)
      const erasedIds = mergePPTEraserHitElementIds(
        interaction.erasedIds,
        getPPTEraserHitElementIds({
          points,
          scene: createPPTCanvasScene(startSlide),
          slide: startSlide,
        }),
      )

      setSelection((current) =>
        removePPTCanvasSelectionIds({ ids: erasedIds, selection: current }))
      setInteraction({
        ...interaction,
        currentPoint: clampPPTPointToSlide(point),
        erasedIds,
        points,
      })
      return
    }

    if (interaction.kind === 'marquee') {
      const bounds = normalizePPTCanvasBounds(interaction.startPoint, point)
      const nextSelection = getPPTCanvasMarqueeSelection({
        additive: interaction.additive,
        baseSelection: interaction.baseSelection,
        bounds,
        scene,
      })

      setInteraction({
        ...interaction,
        currentPoint: point,
      })
      setSelection(nextSelection)
      return
    }

    if (interaction.kind === 'line-create') {
      const currentSlide = findPPTSlide(deckRef.current, interaction.slideId)
      const elements = mapPPTElementsByIds(
        currentSlide.elements,
        [interaction.lineId],
        (element) =>
          element.kind === 'line'
            ? updatePPTLineEndpoint(
                element,
                'end',
                point,
                currentSlide,
                interaction.lineId,
              )
            : element,
      )
      const nextDeck = updatePPTDeckSlide(deckRef.current, interaction.slideId, (slide) => ({
        ...slide,
        elements,
      }))

      deckRef.current = nextDeck
      setDeck(nextDeck)
      return
    }

    if (interaction.kind === 'freeform-create') {
      const points = appendPPTFreeformPoint(interaction.points, point)
      const currentSlide = findPPTSlide(deckRef.current, interaction.slideId)
      const elements = mapPPTElementsByIds(
        currentSlide.elements,
        [interaction.elementId],
        (element) =>
          element.kind === 'freeform'
            ? createPPTFreeformElement({
                id: element.id,
                name: element.name,
                opacity: element.opacity ?? 1,
                points,
                stroke: element.stroke,
              })
            : element,
      )
      const nextDeck = updatePPTDeckSlide(deckRef.current, interaction.slideId, (slide) => ({
        ...slide,
        elements,
      }))

      deckRef.current = nextDeck
      setDeck(nextDeck)
      setInteraction({
        ...interaction,
        points,
      })
      return
    }

    if (interaction.kind === 'element-create') {
      const currentSlide = findPPTSlide(deckRef.current, interaction.slideId)
      const elements = mapPPTElementsByIds(
        currentSlide.elements,
        [interaction.elementId],
        () =>
          createPPTElementFromCreationTool({
            current: point,
            id: interaction.elementId,
            start: interaction.startPoint,
            tool: interaction.tool,
          }),
      )
      const nextDeck = updatePPTDeckSlide(deckRef.current, interaction.slideId, (slide) => ({
        ...slide,
        elements,
      }))

      deckRef.current = nextDeck
      setDeck(nextDeck)
      return
    }

    if (interaction.kind === 'line-route') {
      const currentSlide = findPPTSlide(deckRef.current, interaction.slideId)
      const elements = mapPPTElementsByIds(
        currentSlide.elements,
        [interaction.lineId],
        (element) =>
          element.kind === 'line'
            ? updatePPTLineRouteBend(element, point)
            : element,
      )
      const nextDeck = updatePPTDeckSlide(deckRef.current, interaction.slideId, (slide) => ({
        ...slide,
        elements,
      }))

      deckRef.current = nextDeck
      setDeck(nextDeck)
      return
    }

    const startSlide = findPPTSlide(interaction.startDeck, interaction.slideId)

    if (interaction.kind === 'move') {
      const startScene = createPPTCanvasScene(startSlide)
      const dx = point.x - interaction.startPoint.x
      const dy = point.y - interaction.startPoint.y
      const snap = getPPTCanvasMoveSnap({
        bounds: interaction.bounds,
        config: {
          gestures: {
            snapToAlignment: true,
            snapToGrid: true,
            snapToSpacing: true,
          },
        },
        dx,
        dy,
        scene: startScene,
        selection: interaction.selection,
        viewport,
      })
      const elements = movePPTCanvasSelection({
        adapter: pptCanvasTransformAdapter,
        dx: snap.dx,
        dy: snap.dy,
        items: startSlide.elements,
        selection: interaction.selection,
      })
      const syncedElements = syncPPTLineConnections(
        elements,
        getPPTSelectedLineIds(elements, interaction.selection),
      )

      const nextDeck = updatePPTDeckSlide(interaction.startDeck, interaction.slideId, (slide) => ({
        ...slide,
        elements: syncedElements,
      }))
      deckRef.current = nextDeck
      setDeck(nextDeck)
      setInteraction({
        ...interaction,
        snapGuides: {
          alignmentGuides: snap.alignmentGuides,
          spacingGuides: snap.spacingGuides,
        },
      })
      return
    }

    if (interaction.kind === 'resize') {
      const transformModifierState = getPPTCanvasPointerTransformModifierState(event)
      const elements = resizePPTCanvasSelection({
        adapter: pptCanvasTransformAdapter,
        bounds: interaction.bounds,
        handle: interaction.handle,
        items: startSlide.elements,
        point,
        preserveAspectRatio: transformModifierState.preserveAspectRatio,
        resizeFromCenter: transformModifierState.resizeFromCenter,
        selection: interaction.selection,
      })
      const syncedElements = syncPPTLineConnections(
        elements,
      )

      const nextDeck = updatePPTDeckSlide(interaction.startDeck, interaction.slideId, (slide) => ({
        ...slide,
        elements: syncedElements,
      }))
      deckRef.current = nextDeck
      setDeck(nextDeck)
      return
    }

    if (interaction.kind === 'line-endpoint') {
      const elements = mapPPTElementsByIds(
        startSlide.elements,
        [interaction.lineId],
        (element) =>
          element.kind === 'line'
            ? updatePPTLineEndpoint(
                element,
                interaction.endpoint,
                point,
                startSlide,
                interaction.lineId,
              )
            : element,
      )

      const nextDeck = updatePPTDeckSlide(interaction.startDeck, interaction.slideId, (slide) => ({
        ...slide,
        elements,
      }))
      deckRef.current = nextDeck
      setDeck(nextDeck)
      return
    }

    const delta = getPointAngle(interaction.center, point) - interaction.startAngle
    const rotationById = new Map(interaction.startRotations.map((item) => [
      item.elementId,
      item.rotation,
    ]))
    const transformModifierState = getPPTCanvasPointerTransformModifierState(event)
    const elements = mapPPTElementsByIds(
      startSlide.elements,
      interaction.startRotations.map((item) => item.elementId),
      (element) => {
        const startRotation = rotationById.get(element.id)

        if (startRotation === undefined || element.locked === true) {
          return element
        }

        const rawRotation = normalizePPTCanvasRotationDegrees(startRotation + delta)
        const rotation = transformModifierState.constrainAngle
          ? Math.round(rawRotation / 15) * 15
          : rawRotation

        return {
          ...element,
          geometry: {
            ...element.geometry,
            rotation: normalizePPTCanvasRotationDegrees(rotation),
          },
        }
      },
    )

    const nextDeck = updatePPTDeckSlide(interaction.startDeck, interaction.slideId, (slide) => ({
      ...slide,
      elements,
    }))
    deckRef.current = nextDeck
    setDeck(nextDeck)
  }

  function handlePointerUp() {
    if (!interaction) {
      return
    }

    if (interaction.kind === 'erase' && interaction.erasedIds.length > 0) {
      const erasedIds = new Set(interaction.erasedIds)
      const nextDeck = updatePPTDeckSlide(deckRef.current, interaction.slideId, (slide) => ({
        ...slide,
        elements: deletePPTCanvasSelectionItems({
          getItemId: (element) => element.id,
          items: slide.elements,
          selection: interaction.erasedIds,
        }),
      }))

      deckRef.current = nextDeck
      setDeck(nextDeck)
      setSelection((current) =>
        removePPTCanvasSelectionIds({ ids: erasedIds, selection: current }))
    }

    if (interaction.kind === 'line-create') {
      const currentSlide = findPPTSlide(deckRef.current, interaction.slideId)
      const createdLine = currentSlide.elements.find((element) =>
        element.id === interaction.lineId && element.kind === 'line')

      if (createdLine?.kind === 'line' && getPPTLineLength(createdLine) < 8) {
        const fallbackEnd = clampPPTPointToSlide({
          x: interaction.startPoint.x + 160,
          y: interaction.startPoint.y,
        })
        const elements = mapPPTElementsByIds(
          currentSlide.elements,
          [createdLine.id],
          (element) =>
            element.kind === 'line'
              ? updatePPTLineEndpoint(
                  element,
                  'end',
                  fallbackEnd,
                  currentSlide,
                  interaction.lineId,
                )
              : element,
        )
        const nextDeck = updatePPTDeckSlide(deckRef.current, interaction.slideId, (slide) => ({
          ...slide,
          elements,
        }))

        deckRef.current = nextDeck
        setDeck(nextDeck)
      }

      setLineCreationMode(null)
    }

    if (interaction.kind === 'freeform-create') {
      const currentSlide = findPPTSlide(deckRef.current, interaction.slideId)
      const created = currentSlide.elements.find((element) =>
        element.id === interaction.elementId && element.kind === 'freeform')

      if (created?.kind === 'freeform' && getPPTFreeformWorldLength(created) < 8) {
        const fallbackPoints = [
          interaction.points[0],
          clampPPTPointToSlide({
            x: interaction.points[0].x + 96,
            y: interaction.points[0].y + 36,
          }),
        ]
        const elements = mapPPTElementsByIds(
          currentSlide.elements,
          [interaction.elementId],
          (element) =>
            element.kind === 'freeform'
              ? createPPTFreeformElement({
                  id: element.id,
                  name: element.name,
                  opacity: element.opacity ?? 1,
                  points: fallbackPoints,
                  stroke: element.stroke,
                })
              : element,
        )
        const nextDeck = updatePPTDeckSlide(deckRef.current, interaction.slideId, (slide) => ({
          ...slide,
          elements,
        }))

        deckRef.current = nextDeck
        setDeck(nextDeck)
      }

      setCreationTool(null)
    }

    if (interaction.kind === 'element-create') {
      setCreationTool(null)

      if (interaction.tool.kind === 'sticky' || interaction.tool.kind === 'text') {
        setEditingId(interaction.elementId)
      }
    }

    const historyDeck = getPPTInteractionHistoryDeck(interaction)

    if (
      historyDeck &&
      JSON.stringify(deckRef.current) !== JSON.stringify(historyDeck)
    ) {
      setPast((history) => [...history.slice(-79), historyDeck])
      setFuture([])
    }

    setInteraction(null)
  }

  function getPPTInteractionHistoryDeck(current: Interaction) {
    if (current.kind === 'marquee' || current.kind === 'pan' || current.kind === 'laser') {
      return null
    }

    if (current.kind === 'move') {
      return current.historyDeck ?? current.startDeck
    }

    return current.startDeck
  }

  function zoom(direction: 'in' | 'out') {
    zoomPPTCanvasViewport({
      direction,
      setViewport,
      stageElement: canvasStageElement,
    })
  }

  const marqueeBounds = interaction?.kind === 'marquee'
    ? normalizePPTCanvasBounds(interaction.startPoint, interaction.currentPoint)
    : null
  const marqueeSelection = interaction?.kind === 'marquee'
    ? selection
    : null
  const snapGuides = interaction?.kind === 'move'
    ? interaction.snapGuides
    : EMPTY_PPT_CANVAS_SNAP_GUIDES
  const selectionCommandBarWidth = textQuickFormatState ? 516 : 332
  const selectionCommandAnchor = selectedBounds &&
    !editingId &&
    !interaction &&
    !commandPaletteOpen &&
    !contextMenu &&
    !creationTool &&
    !lineCreationMode &&
    !isLaserToolActive &&
    !isEraserToolActive
    ? getPPTCanvasFloatingAnchorForBounds({
        bounds: selectedBounds,
        floatingSize: {
          height: 40,
          width: selectionCommandBarWidth,
        },
        frameBounds: {
          h: PPT_SLIDE_HEIGHT,
          w: PPT_SLIDE_WIDTH,
          x: 0,
          y: 0,
        },
        stageRect: canvasStageElement.getRect(),
        viewport,
      })
    : null
  const eraserHitIds = interaction?.kind === 'erase'
    ? new Set(interaction.erasedIds)
    : null
  const selectionFloatingCommandGroups = getPPTCommandSurfaceGroups({
    availability: commandAvailability,
    surface: 'selection-floating-bar',
  })
  const contextCommandGroups = getPPTCommandSurfaceGroups({
    availability: commandAvailability,
    surface: 'context-menu',
  })
  const commandPaletteItems: PPTCommandPaletteItem[] = [{
    id: 'system:keyboard-shortcuts',
    onSelect: openShortcutHelp,
    section: 'System',
    shortcut: PPT_SHORTCUT_HELP_SHORTCUT,
    title: 'Keyboard shortcuts',
  }, {
    disabled: !commandAvailability.undo,
    id: 'command:undo',
    onSelect: undo,
    section: 'Edit',
    shortcut: 'Cmd/Ctrl+Z',
    title: PPT_COMMAND_AFFORDANCES.undo.title,
  }, {
    disabled: !commandAvailability.redo,
    id: 'command:redo',
    onSelect: redo,
    section: 'Edit',
    shortcut: 'Cmd/Ctrl+Y',
    title: PPT_COMMAND_AFFORDANCES.redo.title,
  }, {
    disabled: !commandAvailability.duplicate,
    id: 'command:duplicate',
    onSelect: duplicateSelection,
    section: 'Edit',
    shortcut: 'Cmd/Ctrl+D',
    title: PPT_COMMAND_AFFORDANCES.duplicate.title,
  }, {
    disabled: !commandAvailability.delete,
    id: 'command:delete',
    onSelect: deleteSelection,
    section: 'Edit',
    title: PPT_COMMAND_AFFORDANCES.delete.title,
  }, {
    disabled: !commandAvailability.cut,
    id: 'command:cut',
    onSelect: cutSelection,
    section: 'Edit',
    shortcut: 'Cmd/Ctrl+X',
    title: PPT_COMMAND_AFFORDANCES.cut.title,
  }, {
    disabled: selection.length === 0,
    id: 'command:copy',
    onSelect: copySelection,
    section: 'Edit',
    shortcut: 'Cmd/Ctrl+C',
    title: PPT_COMMAND_AFFORDANCES.copy.title,
  }, {
    disabled: !commandAvailability.paste,
    id: 'command:paste',
    onSelect: pasteSelection,
    section: 'Edit',
    shortcut: 'Cmd/Ctrl+V',
    title: PPT_COMMAND_AFFORDANCES.paste.title,
  }, {
    disabled: !commandAvailability.copyFormatting,
    id: 'command:copy-formatting',
    onSelect: copyFormatting,
    section: 'Edit',
    shortcut: SLIDE_EDIT_STYLE_CLIPBOARD_COPY_FORMATTING_SHORTCUT,
    title: 'Copy formatting',
  }, {
    disabled: !commandAvailability.pasteFormatting,
    id: 'command:paste-formatting',
    onSelect: pasteFormatting,
    section: 'Edit',
    shortcut: SLIDE_EDIT_STYLE_CLIPBOARD_PASTE_FORMATTING_SHORTCUT,
    title: 'Paste formatting',
  }, {
    id: 'command:select-all',
    onSelect: selectAllElements,
    section: 'Edit',
    shortcut: 'Cmd/Ctrl+A',
    title: PPT_COMMAND_AFFORDANCES.selectAll.title,
  }, {
    disabled: !commandAvailability.selectSameType,
    id: 'command:select-same-type',
    onSelect: selectSameTypeElements,
    section: 'Edit',
    title: 'Select same type',
  }, {
    disabled: !commandAvailability.alignLeft,
    id: 'command:align-left',
    onSelect: () => alignSelection('alignLeft'),
    section: 'Arrange',
    title: PPT_COMMAND_AFFORDANCES.alignLeft.title,
  }, {
    disabled: !commandAvailability.alignCenter,
    id: 'command:align-center',
    onSelect: () => alignSelection('alignCenter'),
    section: 'Arrange',
    title: PPT_COMMAND_AFFORDANCES.alignCenter.title,
  }, {
    disabled: !commandAvailability.alignRight,
    id: 'command:align-right',
    onSelect: () => alignSelection('alignRight'),
    section: 'Arrange',
    title: PPT_COMMAND_AFFORDANCES.alignRight.title,
  }, {
    disabled: !commandAvailability.alignTop,
    id: 'command:align-top',
    onSelect: () => alignSelection('alignTop'),
    section: 'Arrange',
    title: PPT_COMMAND_AFFORDANCES.alignTop.title,
  }, {
    disabled: !commandAvailability.alignMiddle,
    id: 'command:align-middle',
    onSelect: () => alignSelection('alignMiddle'),
    section: 'Arrange',
    title: PPT_COMMAND_AFFORDANCES.alignMiddle.title,
  }, {
    disabled: !commandAvailability.alignBottom,
    id: 'command:align-bottom',
    onSelect: () => alignSelection('alignBottom'),
    section: 'Arrange',
    title: PPT_COMMAND_AFFORDANCES.alignBottom.title,
  }, {
    disabled: !commandAvailability.distributeHorizontal,
    id: 'command:distribute-horizontal',
    onSelect: () => distributeSelection('distributeHorizontal'),
    section: 'Arrange',
    title: PPT_COMMAND_AFFORDANCES.distributeHorizontal.title,
  }, {
    disabled: !commandAvailability.distributeVertical,
    id: 'command:distribute-vertical',
    onSelect: () => distributeSelection('distributeVertical'),
    section: 'Arrange',
    title: PPT_COMMAND_AFFORDANCES.distributeVertical.title,
  }, {
    disabled: !commandAvailability.tidySelection,
    id: 'command:tidy-selection',
    onSelect: tidySelection,
    section: 'Arrange',
    title: 'Tidy selection',
  }, {
    disabled: !commandAvailability.flipSelection,
    id: 'command:flip-horizontal',
    onSelect: () => flipSelection('horizontal'),
    section: 'Arrange',
    title: 'Flip horizontal',
  }, {
    disabled: !commandAvailability.flipSelection,
    id: 'command:flip-vertical',
    onSelect: () => flipSelection('vertical'),
    section: 'Arrange',
    title: 'Flip vertical',
  }, {
    disabled: !commandAvailability.bringForward,
    id: 'command:bring-forward',
    onSelect: () => reorderSelection('bringForward'),
    section: 'Arrange',
    shortcut: 'Cmd/Ctrl+]',
    title: PPT_COMMAND_AFFORDANCES.bringForward.title,
  }, {
    disabled: !commandAvailability.bringToFront,
    id: 'command:bring-to-front',
    onSelect: () => reorderSelection('bringToFront'),
    section: 'Arrange',
    shortcut: 'Shift+Cmd/Ctrl+]',
    title: PPT_COMMAND_AFFORDANCES.bringToFront.title,
  }, {
    disabled: !commandAvailability.sendBackward,
    id: 'command:send-backward',
    onSelect: () => reorderSelection('sendBackward'),
    section: 'Arrange',
    shortcut: 'Cmd/Ctrl+[',
    title: PPT_COMMAND_AFFORDANCES.sendBackward.title,
  }, {
    disabled: !commandAvailability.sendToBack,
    id: 'command:send-to-back',
    onSelect: () => reorderSelection('sendToBack'),
    section: 'Arrange',
    shortcut: 'Shift+Cmd/Ctrl+[',
    title: PPT_COMMAND_AFFORDANCES.sendToBack.title,
  }, {
    disabled: !commandAvailability.group,
    id: 'command:group',
    onSelect: groupSelection,
    section: 'Arrange',
    shortcut: 'Cmd/Ctrl+G',
    title: PPT_COMMAND_AFFORDANCES.group.title,
  }, {
    disabled: !commandAvailability.ungroup,
    id: 'command:ungroup',
    onSelect: ungroupSelection,
    section: 'Arrange',
    shortcut: 'Shift+Cmd/Ctrl+G',
    title: PPT_COMMAND_AFFORDANCES.ungroup.title,
  }, {
    disabled: !commandAvailability.lockSelection,
    id: 'command:lock-selection',
    onSelect: lockSelectedElements,
    section: 'Arrange',
    shortcut: 'Cmd/Ctrl+L',
    title: PPT_COMMAND_AFFORDANCES.lockSelection.title,
  }, {
    disabled: !commandAvailability.unlockAll,
    id: 'command:unlock-all',
    onSelect: unlockAllElements,
    section: 'Arrange',
    shortcut: 'Shift+Cmd/Ctrl+L',
    title: PPT_COMMAND_AFFORDANCES.unlockAll.title,
  }, {
    id: 'tool:select',
    onSelect: activateSelectTool,
    section: 'Create',
    shortcut: PPT_TOOL_AFFORDANCES.select.shortcut,
    title: 'Select tool',
  }, {
    id: 'tool:pan',
    onSelect: activatePanTool,
    section: 'Create',
    shortcut: PPT_TOOL_AFFORDANCES.pan.shortcut,
    title: PPT_TOOL_AFFORDANCES.pan.ariaLabel,
  }, {
    id: 'tool:laser',
    onSelect: activateLaserTool,
    section: 'Create',
    shortcut: PPT_TOOL_AFFORDANCES.laser.shortcut,
    title: PPT_TOOL_AFFORDANCES.laser.ariaLabel,
  }, {
    id: 'tool:text',
    onSelect: () => activatePPTCreationTool({ kind: 'text' }),
    section: 'Create',
    shortcut: PPT_TOOL_AFFORDANCES.text.shortcut,
    title: PPT_TOOL_AFFORDANCES.text.ariaLabel,
  }, {
    id: 'tool:sticky',
    onSelect: () => activatePPTCreationTool({ kind: 'sticky' }),
    section: 'Create',
    shortcut: PPT_TOOL_AFFORDANCES.sticky.shortcut,
    title: PPT_TOOL_AFFORDANCES.sticky.ariaLabel,
  }, {
    id: 'tool:section',
    onSelect: () => activatePPTCreationTool({ kind: 'section' }),
    section: 'Create',
    shortcut: PPT_TOOL_AFFORDANCES.section.shortcut,
    title: PPT_TOOL_AFFORDANCES.section.ariaLabel,
  }, {
    id: 'tool:rect',
    onSelect: () => activatePPTCreationTool({ kind: 'shape', shape: 'rect' }),
    section: 'Create',
    shortcut: PPT_TOOL_AFFORDANCES.rect.shortcut,
    title: PPT_TOOL_AFFORDANCES.rect.ariaLabel,
  }, {
    id: 'tool:ellipse',
    onSelect: () => activatePPTCreationTool({ kind: 'shape', shape: 'ellipse' }),
    section: 'Create',
    shortcut: PPT_TOOL_AFFORDANCES.ellipse.shortcut,
    title: PPT_TOOL_AFFORDANCES.ellipse.ariaLabel,
  }, {
    id: 'tool:diamond',
    onSelect: () => activatePPTCreationTool({ kind: 'shape', shape: 'diamond' }),
    section: 'Create',
    shortcut: PPT_TOOL_AFFORDANCES.diamond.shortcut,
    title: PPT_TOOL_AFFORDANCES.diamond.ariaLabel,
  }, {
    id: 'tool:line',
    onSelect: () => activateLineCreationMode('line'),
    section: 'Create',
    title: 'Line tool',
  }, {
    id: 'tool:arrow',
    onSelect: () => activateLineCreationMode('arrow'),
    section: 'Create',
    shortcut: PPT_TOOL_AFFORDANCES.arrow.shortcut,
    title: PPT_TOOL_AFFORDANCES.arrow.ariaLabel,
  }, {
    id: 'tool:comment',
    onSelect: () => activatePPTCreationTool({ kind: 'comment' }),
    section: 'Create',
    shortcut: PPT_TOOL_AFFORDANCES.comment.shortcut,
    title: PPT_TOOL_AFFORDANCES.comment.ariaLabel,
  }, {
    id: 'tool:pen',
    onSelect: () => activatePPTCreationTool({ kind: 'freeform', tool: 'pen' }),
    section: 'Create',
    shortcut: PPT_TOOL_AFFORDANCES.pen.shortcut,
    title: PPT_TOOL_AFFORDANCES.pen.ariaLabel,
  }, {
    id: 'tool:marker',
    onSelect: () => activatePPTCreationTool({ kind: 'freeform', tool: 'marker' }),
    section: 'Create',
    shortcut: PPT_TOOL_AFFORDANCES.marker.shortcut,
    title: PPT_TOOL_AFFORDANCES.marker.ariaLabel,
  }, {
    id: 'tool:highlight',
    onSelect: () => activatePPTCreationTool({ kind: 'freeform', tool: 'highlight' }),
    section: 'Create',
    shortcut: PPT_TOOL_AFFORDANCES.highlight.shortcut,
    title: PPT_TOOL_AFFORDANCES.highlight.ariaLabel,
  }, {
    id: 'tool:eraser',
    onSelect: activateEraserTool,
    section: 'Create',
    shortcut: PPT_TOOL_AFFORDANCES.eraser.shortcut,
    title: PPT_TOOL_AFFORDANCES.eraser.ariaLabel,
  }, {
    id: 'tool:image',
    onSelect: () => imageInputRef.current?.click(),
    section: 'Create',
    title: 'Add image',
  }, {
    id: 'tool:paste-image',
    onSelect: () => {
      void pastePPTClipboardImage()
    },
    section: 'Create',
    title: 'Paste image',
  }, {
    id: 'tool:table',
    onSelect: () => insertPPTTableSource(),
    section: 'Create',
    title: 'Add table',
  }, {
    disabled: !selectedTableElement,
    id: 'table:copy',
    onSelect: copySelectedTable,
    section: 'Edit',
    title: 'Copy table',
  }, {
    id: 'slide:add',
    onSelect: addSlide,
    section: 'Slides',
    title: 'Add slide',
  }, {
    id: 'slide:duplicate',
    onSelect: duplicateActiveSlide,
    section: 'Slides',
    title: 'Duplicate slide',
  }, {
    id: 'slide:copy',
    onSelect: copyActiveSlide,
    section: 'Slides',
    title: 'Copy slide',
  }, {
    disabled: !canPasteSlide,
    id: 'slide:paste',
    onSelect: pasteCopiedSlide,
    section: 'Slides',
    title: 'Paste slide',
  }, {
    disabled: !canDeleteSlide,
    id: 'slide:delete',
    onSelect: deleteActiveSlide,
    section: 'Slides',
    title: 'Delete slide',
  }, {
    disabled: !canMoveActiveSlideUp,
    id: 'slide:move-up',
    onSelect: () => moveActiveSlide(-1),
    section: 'Slides',
    title: 'Move slide up',
  }, {
    disabled: !canMoveActiveSlideDown,
    id: 'slide:move-down',
    onSelect: () => moveActiveSlide(1),
    section: 'Slides',
    title: 'Move slide down',
  }, ...PPT_LAYOUT_DESCRIPTORS.map((layout) => ({
    disabled: activeLayout.layoutId === layout.layoutId,
    id: `slide:layout:${layout.layoutId}`,
    onSelect: () => applySlideLayout(layout.layoutId),
    section: 'Slides',
    title: `Apply ${layout.name}`,
  })), {
    id: 'export:copy-slide-svg',
    onSelect: copySlideSVG,
    section: 'Export',
    title: 'Copy slide SVG',
  }, {
    disabled: !canExportSelectionSVG,
    id: 'export:copy-selection-svg',
    onSelect: copySelectionSVG,
    section: 'Export',
    title: 'Copy selection SVG',
  }, {
    id: 'export:download-slide-svg',
    onSelect: downloadSlideSVG,
    section: 'Export',
    title: 'Download slide SVG',
  }, {
    disabled: !canExportSelectionSVG,
    id: 'export:download-selection-svg',
    onSelect: downloadSelectionSVG,
    section: 'Export',
    title: 'Download selection SVG',
  }, {
    id: 'view:find',
    onSelect: openFindStrip,
    section: 'View',
    shortcut: 'Cmd/Ctrl+F',
    title: 'Find text',
  }, {
    id: 'view:present',
    onSelect: () => startPresentation(),
    section: 'View',
    title: 'Start presentation',
  }, {
    id: 'view:fit-slide',
    onSelect: fitSlide,
    section: 'View',
    shortcut: '0',
    title: 'Fit slide',
  }, {
    disabled: !canFitSelection,
    id: 'view:fit-selection',
    onSelect: fitSelection,
    section: 'View',
    shortcut: '1',
    title: 'Fit selection',
  }, {
    id: 'view:reset-zoom',
    onSelect: resetZoom,
    section: 'View',
    shortcut: 'Cmd/Ctrl+0',
    title: PPT_COMMAND_AFFORDANCES.zoomReset.title,
  }, {
    id: 'view:zoom-in',
    onSelect: () => zoom('in'),
    section: 'View',
    shortcut: 'Cmd/Ctrl+=',
    title: PPT_COMMAND_AFFORDANCES.zoomIn.title,
  }, {
    id: 'view:zoom-out',
    onSelect: () => zoom('out'),
    section: 'View',
    shortcut: 'Cmd/Ctrl+-',
    title: PPT_COMMAND_AFFORDANCES.zoomOut.title,
  }, {
    id: 'view:toggle-grid',
    onSelect: () => setShowGrid((current) => !current),
    section: 'View',
    title: showGrid ? 'Hide grid' : 'Show grid',
  }, {
    id: 'view:toggle-minimap',
    onSelect: () => setShowMinimap((current) => !current),
    section: 'View',
    title: showMinimap ? 'Hide minimap' : 'Show minimap',
  }, {
    id: 'view:toggle-frame-guides',
    onSelect: () => setShowFrameGuides((current) => !current),
    section: 'View',
    title: showFrameGuides ? 'Hide frame guides' : 'Show frame guides',
  }, {
    id: 'view:toggle-theme',
    onSelect: toggleTheme,
    section: 'View',
    title: theme === 'dark' ? 'Switch to light theme' : 'Switch to dark theme',
  }, {
    disabled: !canFormatSelectedText,
    id: 'format:bold',
    onSelect: toggleSelectedTextBold,
    section: 'Format',
    shortcut: 'Cmd/Ctrl+B',
    title: 'Bold text',
  }, {
    disabled: !canFormatSelectedText,
    id: 'format:italic',
    onSelect: toggleSelectedTextItalic,
    section: 'Format',
    title: 'Italic text',
  }, {
    disabled: !canFormatSelectedText,
    id: 'format:underline',
    onSelect: toggleSelectedTextUnderline,
    section: 'Format',
    title: 'Underline text',
  }, {
    disabled: !canFormatSelectedText,
    id: 'format:bullet',
    onSelect: toggleSelectedParagraphBullet,
    section: 'Format',
    title: 'Toggle bullet list',
  }, {
    disabled: !canFormatSelectedText,
    id: 'format:numbered',
    onSelect: toggleSelectedParagraphNumbered,
    section: 'Format',
    title: 'Toggle numbered list',
  }, {
    disabled: !selectedElement || !isPPTTextElement(selectedElement) || !selectedTextOverflow,
    id: 'format:auto-fit-text',
    onSelect: () => {
      if (selectedElement && isPPTTextElement(selectedElement)) {
        autoFitTextElement(selectedElement.id)
      }
    },
    section: 'Format',
    title: 'Auto fit text',
  }]
  const shortcutHelpItems = getPPTShortcutHelpItems(commandPaletteItems)

  return (
    <main className="ppt-app" data-ppt-app data-theme={theme}>
      <header
        aria-label="PPT editor toolbar"
        aria-orientation="horizontal"
        className="ppt-topbar"
        data-ppt-toolbar
        data-ppt-toolbar-focus-model={PPT_TOOLBAR_FOCUS_MODEL}
        data-ppt-toolbar-keyboard-model={PPT_TOOLBAR_KEYBOARD_MODEL}
        data-ppt-toolbar-model={PPT_TOOLBAR_ROVING_FOCUS_MODEL}
        ref={setTopbarToolbarRoot}
        role="toolbar"
        onFocus={handleTopbarToolbarFocus}
        onKeyDown={handleTopbarToolbarKeyDown}
      >
        <div className="ppt-brand">
          <strong>PPT</strong>
          <span>{deck.title}</span>
        </div>
        <div className="ppt-toolbar-group">
          <button {...PPT_TOOLBAR_ITEM_PROPS} className="ppt-icon-button" disabled={!commandAvailability.undo} onClick={undo} title={PPT_COMMAND_AFFORDANCES.undo.title} type="button">
            <Undo2 size={17} />
          </button>
          <button {...PPT_TOOLBAR_ITEM_PROPS} className="ppt-icon-button" disabled={!commandAvailability.redo} onClick={redo} title={PPT_COMMAND_AFFORDANCES.redo.title} type="button">
            <Redo2 size={17} />
          </button>
          <button {...PPT_TOOLBAR_ITEM_PROPS} className="ppt-icon-button" data-ppt-find-open onClick={openFindStrip} title="Find text" type="button">
            <Search size={17} />
          </button>
          <button {...PPT_TOOLBAR_ITEM_PROPS} className="ppt-icon-button" data-ppt-command-palette-open onClick={openCommandPalette} title="Command palette" type="button">
            <Command size={17} />
          </button>
          <button {...PPT_TOOLBAR_ITEM_PROPS} className="ppt-icon-button" data-ppt-command="copy-formatting" disabled={!commandAvailability.copyFormatting} onClick={copyFormatting} title="Copy formatting" type="button">
            <Paintbrush size={17} />
          </button>
          <button {...PPT_TOOLBAR_ITEM_PROPS} className="ppt-icon-button" data-ppt-command="paste-formatting" disabled={!commandAvailability.pasteFormatting} onClick={pasteFormatting} title="Paste formatting" type="button">
            <Paintbrush size={17} />
          </button>
          <button {...PPT_TOOLBAR_ITEM_PROPS} className="ppt-icon-button" data-ppt-shortcut-help-open onClick={openShortcutHelp} title="Keyboard shortcuts" type="button">
            <Keyboard size={17} />
          </button>
        </div>
        {findOpen ? (
          <FindReplaceStrip
            activeIndex={clampedFindIndex}
            inputRef={findInputRef}
            matchCount={findMatches.length}
            query={findQuery}
            replaceQuery={replaceQuery}
            onClose={closeFindStrip}
            onFindKeyDown={handleFindKeyDown}
            onFindQueryChange={updateFindQuery}
            onNext={() => goToFindMatch(1)}
            onPrevious={() => goToFindMatch(-1)}
            onReplace={replaceActiveFindMatch}
            onReplaceAll={replaceAllFindMatches}
            onReplaceQueryChange={setReplaceQuery}
          />
        ) : null}
        <div className="ppt-toolbar-group">
          <button
            {...PPT_TOOLBAR_ITEM_PROPS}
            aria-label={PPT_TOOL_AFFORDANCES.pan.ariaLabel}
            aria-pressed={isPanToolActive}
            className="ppt-icon-button"
            data-ppt-pan-tool
            data-ppt-tool="pan"
            onClick={activatePanTool}
            title={PPT_TOOL_AFFORDANCES.pan.title}
            type="button"
          >
            <Hand size={17} />
          </button>
          <button
            {...PPT_TOOLBAR_ITEM_PROPS}
            aria-label={PPT_TOOL_AFFORDANCES.laser.ariaLabel}
            aria-pressed={isLaserToolActive}
            className="ppt-icon-button"
            data-ppt-laser-tool
            data-ppt-tool="laser"
            onClick={activateLaserTool}
            title={PPT_TOOL_AFFORDANCES.laser.title}
            type="button"
          >
            <MousePointer2 size={17} />
          </button>
          <button
            {...PPT_TOOLBAR_ITEM_PROPS}
            aria-label={PPT_TOOL_AFFORDANCES.text.ariaLabel}
            aria-pressed={creationTool?.kind === 'text'}
            className="ppt-icon-button"
            data-ppt-insert-tool="text"
            onClick={() => activatePPTCreationTool({ kind: 'text' })}
            title={PPT_TOOL_AFFORDANCES.text.title}
            type="button"
          >
            <Type size={17} />
          </button>
          <button
            {...PPT_TOOLBAR_ITEM_PROPS}
            aria-label={PPT_TOOL_AFFORDANCES.sticky.ariaLabel}
            aria-pressed={creationTool?.kind === 'sticky'}
            className="ppt-icon-button"
            data-ppt-insert-tool="sticky"
            onClick={() => activatePPTCreationTool({ kind: 'sticky' })}
            title={PPT_TOOL_AFFORDANCES.sticky.title}
            type="button"
          >
            <StickyNote size={17} />
          </button>
          <button
            {...PPT_TOOLBAR_ITEM_PROPS}
            aria-label={PPT_TOOL_AFFORDANCES.section.ariaLabel}
            aria-pressed={creationTool?.kind === 'section'}
            className="ppt-icon-button"
            data-ppt-insert-tool="section"
            onClick={() => activatePPTCreationTool({ kind: 'section' })}
            title={PPT_TOOL_AFFORDANCES.section.title}
            type="button"
          >
            <Frame size={17} />
          </button>
          <button
            {...PPT_TOOLBAR_ITEM_PROPS}
            aria-label={PPT_TOOL_AFFORDANCES.rect.ariaLabel}
            aria-pressed={isPPTShapeCreationTool(creationTool, 'rect')}
            className="ppt-icon-button"
            data-ppt-insert-shape="rect"
            data-ppt-insert-tool="rect"
            onClick={() => activatePPTCreationTool({ kind: 'shape', shape: 'rect' })}
            title={PPT_TOOL_AFFORDANCES.rect.title}
            type="button"
          >
            <Square size={17} />
          </button>
          <button
            {...PPT_TOOLBAR_ITEM_PROPS}
            aria-label={PPT_TOOL_AFFORDANCES.ellipse.ariaLabel}
            aria-pressed={isPPTShapeCreationTool(creationTool, 'ellipse')}
            className="ppt-icon-button"
            data-ppt-insert-shape="ellipse"
            data-ppt-insert-tool="ellipse"
            onClick={() => activatePPTCreationTool({ kind: 'shape', shape: 'ellipse' })}
            title={PPT_TOOL_AFFORDANCES.ellipse.title}
            type="button"
          >
            <Circle size={17} />
          </button>
          <button
            {...PPT_TOOLBAR_ITEM_PROPS}
            aria-label={PPT_TOOL_AFFORDANCES.diamond.ariaLabel}
            aria-pressed={isPPTShapeCreationTool(creationTool, 'diamond')}
            className="ppt-icon-button"
            data-ppt-insert-shape="diamond"
            data-ppt-insert-tool="diamond"
            onClick={() => activatePPTCreationTool({ kind: 'shape', shape: 'diamond' })}
            title={PPT_TOOL_AFFORDANCES.diamond.title}
            type="button"
          >
            <Diamond size={17} />
          </button>
          <button {...PPT_TOOLBAR_ITEM_PROPS} aria-pressed={lineCreationMode === 'line'} className="ppt-icon-button" data-ppt-insert-line="line" onClick={() => activateLineCreationMode('line')} title="Draw line" type="button">
            <Minus size={17} />
          </button>
          <button
            {...PPT_TOOLBAR_ITEM_PROPS}
            aria-label={PPT_TOOL_AFFORDANCES.arrow.ariaLabel}
            aria-pressed={lineCreationMode === 'arrow'}
            className="ppt-icon-button"
            data-ppt-insert-line="arrow"
            data-ppt-insert-tool="arrow"
            onClick={() => activateLineCreationMode('arrow')}
            title={PPT_TOOL_AFFORDANCES.arrow.title}
            type="button"
          >
            <ArrowRight size={17} />
          </button>
          <button
            {...PPT_TOOLBAR_ITEM_PROPS}
            aria-label={PPT_TOOL_AFFORDANCES.pen.ariaLabel}
            aria-pressed={isPPTFreeformCreationTool(creationTool, 'pen')}
            className="ppt-icon-button"
            data-ppt-insert-tool="pen"
            onClick={() => activatePPTCreationTool({ kind: 'freeform', tool: 'pen' })}
            title={PPT_TOOL_AFFORDANCES.pen.title}
            type="button"
          >
            <PenLine size={17} />
          </button>
          <button
            {...PPT_TOOLBAR_ITEM_PROPS}
            aria-label={PPT_TOOL_AFFORDANCES.marker.ariaLabel}
            aria-pressed={isPPTFreeformCreationTool(creationTool, 'marker')}
            className="ppt-icon-button"
            data-ppt-insert-tool="marker"
            onClick={() => activatePPTCreationTool({ kind: 'freeform', tool: 'marker' })}
            title={PPT_TOOL_AFFORDANCES.marker.title}
            type="button"
          >
            <PencilLine size={17} />
          </button>
          <button
            {...PPT_TOOLBAR_ITEM_PROPS}
            aria-label={PPT_TOOL_AFFORDANCES.highlight.ariaLabel}
            aria-pressed={isPPTFreeformCreationTool(creationTool, 'highlight')}
            className="ppt-icon-button"
            data-ppt-insert-tool="highlight"
            onClick={() => activatePPTCreationTool({ kind: 'freeform', tool: 'highlight' })}
            title={PPT_TOOL_AFFORDANCES.highlight.title}
            type="button"
          >
            <Highlighter size={17} />
          </button>
          <button
            {...PPT_TOOLBAR_ITEM_PROPS}
            aria-label={PPT_TOOL_AFFORDANCES.eraser.ariaLabel}
            aria-pressed={isEraserToolActive}
            className="ppt-icon-button"
            data-ppt-eraser-tool
            data-ppt-tool="eraser"
            onClick={activateEraserTool}
            title={PPT_TOOL_AFFORDANCES.eraser.title}
            type="button"
          >
            <Eraser size={17} />
          </button>
          <button
            {...PPT_TOOLBAR_ITEM_PROPS}
            aria-label={PPT_TOOL_AFFORDANCES.comment.ariaLabel}
            aria-pressed={creationTool?.kind === 'comment'}
            className="ppt-icon-button"
            data-ppt-insert-comment
            data-ppt-insert-tool="comment"
            onClick={() => activatePPTCreationTool({ kind: 'comment' })}
            title={PPT_TOOL_AFFORDANCES.comment.title}
            type="button"
          >
            <MessageSquare size={17} />
          </button>
          <button {...PPT_TOOLBAR_ITEM_PROPS} className="ppt-icon-button" data-ppt-insert-image onClick={() => imageInputRef.current?.click()} title="Add image" type="button">
            <ImagePlus size={17} />
          </button>
          <button
            {...PPT_TOOLBAR_ITEM_PROPS}
            className="ppt-icon-button"
            data-ppt-paste-image
            onClick={() => {
              void pastePPTClipboardImage()
            }}
            title="Paste image"
            type="button"
          >
            <ClipboardPaste size={17} />
          </button>
          <button {...PPT_TOOLBAR_ITEM_PROPS} className="ppt-icon-button" data-ppt-insert-table onClick={() => insertPPTTableSource()} title="Add table" type="button">
            <Table2 size={17} />
          </button>
          <button
            {...PPT_TOOLBAR_ITEM_PROPS}
            className="ppt-icon-button"
            data-ppt-copy-table
            disabled={!selectedTableElement}
            onClick={copySelectedTable}
            title="Copy table"
            type="button"
          >
            <Copy size={17} />
          </button>
          <input
            accept="image/*"
            className="ppt-file-input"
            data-ppt-image-upload-input
            ref={imageInputRef}
            tabIndex={-1}
            type="file"
            onChange={handleImageInputChange}
          />
          <button {...PPT_TOOLBAR_ITEM_PROPS} className="ppt-icon-button" disabled={!commandAvailability.delete} onClick={deleteSelection} title={PPT_COMMAND_AFFORDANCES.delete.title} type="button">
            <Trash2 size={17} />
          </button>
        </div>
        <div className="ppt-toolbar-group">
          <button {...PPT_TOOLBAR_ITEM_PROPS} className="ppt-icon-button" data-ppt-command="align-left" disabled={!commandAvailability.alignLeft} onClick={() => alignSelection('alignLeft')} title={PPT_COMMAND_AFFORDANCES.alignLeft.title} type="button">
            <AlignStartVertical size={17} />
          </button>
          <button {...PPT_TOOLBAR_ITEM_PROPS} className="ppt-icon-button" data-ppt-command="align-center-x" disabled={!commandAvailability.alignCenter} onClick={() => alignSelection('alignCenter')} title={PPT_COMMAND_AFFORDANCES.alignCenter.title} type="button">
            <AlignCenterVertical size={17} />
          </button>
          <button {...PPT_TOOLBAR_ITEM_PROPS} className="ppt-icon-button" data-ppt-command="align-right" disabled={!commandAvailability.alignRight} onClick={() => alignSelection('alignRight')} title={PPT_COMMAND_AFFORDANCES.alignRight.title} type="button">
            <AlignEndVertical size={17} />
          </button>
          <button {...PPT_TOOLBAR_ITEM_PROPS} className="ppt-icon-button" data-ppt-command="align-top" disabled={!commandAvailability.alignTop} onClick={() => alignSelection('alignTop')} title={PPT_COMMAND_AFFORDANCES.alignTop.title} type="button">
            <AlignStartHorizontal size={17} />
          </button>
          <button {...PPT_TOOLBAR_ITEM_PROPS} className="ppt-icon-button" data-ppt-command="align-middle" disabled={!commandAvailability.alignMiddle} onClick={() => alignSelection('alignMiddle')} title={PPT_COMMAND_AFFORDANCES.alignMiddle.title} type="button">
            <AlignCenterHorizontal size={17} />
          </button>
          <button {...PPT_TOOLBAR_ITEM_PROPS} className="ppt-icon-button" data-ppt-command="align-bottom" disabled={!commandAvailability.alignBottom} onClick={() => alignSelection('alignBottom')} title={PPT_COMMAND_AFFORDANCES.alignBottom.title} type="button">
            <AlignEndHorizontal size={17} />
          </button>
          <button {...PPT_TOOLBAR_ITEM_PROPS} className="ppt-icon-button" data-ppt-command="distribute-horizontal" disabled={!commandAvailability.distributeHorizontal} onClick={() => distributeSelection('distributeHorizontal')} title={PPT_COMMAND_AFFORDANCES.distributeHorizontal.title} type="button">
            <AlignHorizontalDistributeCenter size={17} />
          </button>
          <button {...PPT_TOOLBAR_ITEM_PROPS} className="ppt-icon-button" data-ppt-command="distribute-vertical" disabled={!commandAvailability.distributeVertical} onClick={() => distributeSelection('distributeVertical')} title={PPT_COMMAND_AFFORDANCES.distributeVertical.title} type="button">
            <AlignVerticalDistributeCenter size={17} />
          </button>
          <button {...PPT_TOOLBAR_ITEM_PROPS} className="ppt-icon-button" data-ppt-command="tidy-selection" disabled={!commandAvailability.tidySelection} onClick={tidySelection} title="Tidy selection" type="button">
            <Grid2X2 size={17} />
          </button>
          <button {...PPT_TOOLBAR_ITEM_PROPS} className="ppt-icon-button" data-ppt-command="flip-horizontal" disabled={!commandAvailability.flipSelection} onClick={() => flipSelection('horizontal')} title="Flip horizontal" type="button">
            <FlipHorizontal2 size={17} />
          </button>
          <button {...PPT_TOOLBAR_ITEM_PROPS} className="ppt-icon-button" data-ppt-command="flip-vertical" disabled={!commandAvailability.flipSelection} onClick={() => flipSelection('vertical')} title="Flip vertical" type="button">
            <FlipVertical2 size={17} />
          </button>
        </div>
        <div className="ppt-toolbar-group">
          <button {...PPT_TOOLBAR_ITEM_PROPS} className="ppt-icon-button" data-ppt-command="bring-forward" disabled={!commandAvailability.bringForward} onClick={() => reorderSelection('bringForward')} title={PPT_COMMAND_AFFORDANCES.bringForward.title} type="button">
            <MoveUp size={17} />
          </button>
          <button {...PPT_TOOLBAR_ITEM_PROPS} className="ppt-icon-button" data-ppt-command="bring-to-front" disabled={!commandAvailability.bringToFront} onClick={() => reorderSelection('bringToFront')} title={PPT_COMMAND_AFFORDANCES.bringToFront.title} type="button">
            <BringToFront size={17} />
          </button>
          <button {...PPT_TOOLBAR_ITEM_PROPS} className="ppt-icon-button" data-ppt-command="send-backward" disabled={!commandAvailability.sendBackward} onClick={() => reorderSelection('sendBackward')} title={PPT_COMMAND_AFFORDANCES.sendBackward.title} type="button">
            <MoveDown size={17} />
          </button>
          <button {...PPT_TOOLBAR_ITEM_PROPS} className="ppt-icon-button" data-ppt-command="send-to-back" disabled={!commandAvailability.sendToBack} onClick={() => reorderSelection('sendToBack')} title={PPT_COMMAND_AFFORDANCES.sendToBack.title} type="button">
            <SendToBack size={17} />
          </button>
        </div>
        <div className="ppt-toolbar-group">
          <button {...PPT_TOOLBAR_ITEM_PROPS} className="ppt-icon-button" data-ppt-command="group" disabled={!commandAvailability.group} onClick={groupSelection} title={PPT_COMMAND_AFFORDANCES.group.title} type="button">
            <Group size={17} />
          </button>
          <button {...PPT_TOOLBAR_ITEM_PROPS} className="ppt-icon-button" data-ppt-command="ungroup" disabled={!commandAvailability.ungroup} onClick={ungroupSelection} title={PPT_COMMAND_AFFORDANCES.ungroup.title} type="button">
            <Ungroup size={17} />
          </button>
          <button {...PPT_TOOLBAR_ITEM_PROPS} className="ppt-icon-button" data-ppt-command="lock-selection" disabled={!commandAvailability.lockSelection} onClick={lockSelectedElements} title={PPT_COMMAND_AFFORDANCES.lockSelection.title} type="button">
            <Lock size={17} />
          </button>
          <button {...PPT_TOOLBAR_ITEM_PROPS} className="ppt-icon-button" data-ppt-command="unlock-all" disabled={!commandAvailability.unlockAll} onClick={unlockAllElements} title={PPT_COMMAND_AFFORDANCES.unlockAll.title} type="button">
            <Unlock size={17} />
          </button>
        </div>
        <div className="ppt-toolbar-group">
          <button {...PPT_TOOLBAR_ITEM_PROPS} className="ppt-icon-button" data-ppt-present-start onClick={() => startPresentation()} title="Start presentation" type="button">
            <Play size={17} />
          </button>
          <button {...PPT_TOOLBAR_ITEM_PROPS} className="ppt-icon-button" onClick={() => zoom('out')} title="Zoom out" type="button">
            <ZoomOut size={17} />
          </button>
          <button {...PPT_TOOLBAR_ITEM_PROPS} className="ppt-icon-button" data-ppt-view-fit-slide onClick={fitSlide} title="Fit slide" type="button">
            <Maximize2 size={17} />
          </button>
          <button {...PPT_TOOLBAR_ITEM_PROPS} className="ppt-icon-button" data-ppt-view-fit-selection disabled={!canFitSelection} onClick={fitSelection} title="Fit selection" type="button">
            <Maximize2 size={17} />
          </button>
          <button {...PPT_TOOLBAR_ITEM_PROPS} className="ppt-icon-button" onClick={() => zoom('in')} title="Zoom in" type="button">
            <ZoomIn size={17} />
          </button>
          <button {...PPT_TOOLBAR_ITEM_PROPS} aria-pressed={showGrid} className="ppt-icon-button" data-ppt-view-grid onClick={() => setShowGrid((current) => !current)} title="Toggle grid" type="button">
            <Grid2X2 size={17} />
          </button>
          <button {...PPT_TOOLBAR_ITEM_PROPS} aria-pressed={showMinimap} className="ppt-icon-button" data-ppt-view-minimap onClick={() => setShowMinimap((current) => !current)} title="Toggle minimap" type="button">
            <MapIcon size={17} />
          </button>
          <button {...PPT_TOOLBAR_ITEM_PROPS} aria-pressed={showFrameGuides} className="ppt-icon-button" data-ppt-view-frame-guides onClick={() => setShowFrameGuides((current) => !current)} title="Toggle frame guides" type="button">
            <Ruler size={17} />
          </button>
          <button
            {...PPT_TOOLBAR_ITEM_PROPS}
            aria-label={theme === 'dark' ? 'Switch to light theme' : 'Switch to dark theme'}
            aria-pressed={theme === 'dark'}
            className="ppt-icon-button"
            data-ppt-theme-toggle
            onClick={toggleTheme}
            title={theme === 'dark' ? 'Switch to light theme' : 'Switch to dark theme'}
            type="button"
          >
            {theme === 'dark' ? <Sun size={17} /> : <Moon size={17} />}
          </button>
          <span className="ppt-zoom-label">{Math.round(viewport.scale * 100)}%</span>
        </div>
        <div className="ppt-toolbar-group">
          <button {...PPT_TOOLBAR_ITEM_PROPS} aria-label="Copy HTML" className="ppt-button" onClick={copyHTML} type="button">
            <Copy size={16} /> HTML
          </button>
          <button {...PPT_TOOLBAR_ITEM_PROPS} aria-label="Download HTML" className="ppt-button" onClick={downloadHTML} type="button">
            <Download size={16} /> HTML
          </button>
          <button {...PPT_TOOLBAR_ITEM_PROPS} aria-label="Copy slide SVG" className="ppt-button" data-ppt-copy-slide-svg onClick={copySlideSVG} type="button">
            <Copy size={16} /> SVG
          </button>
          <button {...PPT_TOOLBAR_ITEM_PROPS} aria-label="Download slide SVG" className="ppt-button" data-ppt-export-svg onClick={downloadSlideSVG} type="button">
            <Download size={16} /> SVG
          </button>
          <button {...PPT_TOOLBAR_ITEM_PROPS} aria-label="Copy selection SVG" className="ppt-button" data-ppt-copy-selection-svg disabled={!canExportSelectionSVG} onClick={copySelectionSVG} type="button">
            <Copy size={16} /> Sel SVG
          </button>
          <button {...PPT_TOOLBAR_ITEM_PROPS} aria-label="Download selection SVG" className="ppt-button" data-ppt-export-selection-svg disabled={!canExportSelectionSVG} onClick={downloadSelectionSVG} type="button">
            <Download size={16} /> Sel SVG
          </button>
        </div>
        <span className="ppt-selection-label">
          {selection.length > 0 ? `${selection.length} selected` : 'No selection'}
        </span>
      </header>

      <aside className="ppt-rail" aria-label="Slides">
        <div className="ppt-rail-header">
          <h2>Slides</h2>
          <div className="ppt-slide-actions">
            <button className="ppt-slide-action" data-ppt-slide-action="add" onClick={addSlide} title="Add slide" type="button">
              <FilePlus2 size={16} />
            </button>
            <button className="ppt-slide-action" data-ppt-slide-action="duplicate" onClick={duplicateActiveSlide} title="Duplicate slide" type="button">
              <CopyPlus size={16} />
            </button>
            <button className="ppt-slide-action" data-ppt-slide-action="copy" onClick={copyActiveSlide} title="Copy slide" type="button">
              <Copy size={16} />
            </button>
            <button className="ppt-slide-action" data-ppt-slide-action="paste" disabled={!canPasteSlide} onClick={pasteCopiedSlide} title="Paste slide" type="button">
              <ClipboardPaste size={16} />
            </button>
            <button className="ppt-slide-action" data-ppt-slide-action="move-up" disabled={!canMoveActiveSlideUp} onClick={() => moveActiveSlide(-1)} title="Move slide up" type="button">
              <ChevronUp size={16} />
            </button>
            <button className="ppt-slide-action" data-ppt-slide-action="move-down" disabled={!canMoveActiveSlideDown} onClick={() => moveActiveSlide(1)} title="Move slide down" type="button">
              <ChevronDown size={16} />
            </button>
            <button className="ppt-slide-action" data-ppt-slide-action="delete" disabled={!canDeleteSlide} onClick={deleteActiveSlide} title="Delete slide" type="button">
              <Trash2 size={16} />
            </button>
          </div>
        </div>
        <div
          aria-label="Slides"
          className="ppt-slide-list"
          data-ppt-slide-dragging={slideDragState?.draggingSlideId ?? undefined}
          data-ppt-slide-drop-placement={slideDragState?.dropPlacement ?? undefined}
          data-ppt-slide-drop-target={slideDragState?.dropTargetSlideId ?? undefined}
          data-ppt-slide-list
          data-ppt-slide-rail-active={slideRailDescriptor.activeSlideId ?? undefined}
          data-ppt-slide-rail-active-option={slideRailDescriptor.listbox.activeOptionId ?? undefined}
          data-ppt-slide-rail-command={lastSlideRailCommandEffect?.payload.id}
          data-ppt-slide-rail-command-from-index={lastSlideRailReorderPayload?.fromIndex}
          data-ppt-slide-rail-command-selection-slide={lastSlideRailCommandEffect?.selection?.slideId}
          data-ppt-slide-rail-command-slide={lastSlideRailReorderPayload?.slideId}
          data-ppt-slide-rail-command-to-index={lastSlideRailReorderPayload?.toIndex}
          data-ppt-slide-rail-command-type={lastSlideRailCommandEffect?.type}
          data-ppt-slide-rail-focusable-option={slideRailDescriptor.listbox.focusableOptionId ?? undefined}
          data-ppt-slide-rail-keyboard-keys={SLIDE_EDIT_RAIL_KEYBOARD_KEYS}
          data-ppt-slide-rail-keyboard-model={slideRailDescriptor.listbox.keyboardModel}
          data-ppt-slide-rail-model="slide-edit-rail-interactions"
          data-ppt-slide-rail-option-count={slideRailDescriptor.listbox.options.length}
          data-ppt-slide-rail-selection-mode={slideRailDescriptor.listbox.selectionMode}
          data-ppt-slide-rail-slide-order={slideRailDescriptor.slideOrder.join(' ')}
          data-ppt-slide-rail-thumbnail-count={slideRailDescriptor.thumbnails.length}
          role={slideRailDescriptor.listbox.role}
        >
          {deck.slides.map((slide, index) => (
            <SlideThumb
              active={slide.id === activeSlide.id}
              dragging={slideDragState?.draggingSlideId === slide.id}
              dropPlacement={slideDragState?.dropTargetSlideId === slide.id
                ? slideDragState.dropPlacement
                : undefined}
              index={index}
              key={slide.id}
              optionDescriptor={slideRailDescriptor.listbox.options[index]}
              slide={slide}
              thumbnailDescriptor={slideRailDescriptor.thumbnails[index]}
              onDragEnd={handleSlideThumbDragEnd}
              onDragOver={(event) => handleSlideThumbDragOver(slide.id, event)}
              onDragStart={(event) => handleSlideThumbDragStart(slide.id, event)}
              onDrop={(event) => handleSlideThumbDrop(slide.id, event)}
              onKeyDown={(event) => handleSlideThumbKeyDown(slide.id, event)}
              onSelect={(event) => handleSlideThumbSelect(slide.id, event)}
            />
          ))}
        </div>
      </aside>

      <section
        className="ppt-stage-shell"
        data-ppt-clipboard-count={clipboard?.objects.length ?? 0}
        data-ppt-clipboard-metadata-count={clipboard?.metadata.length ?? 0}
        data-ppt-clipboard-model="slide-edit-clipboard"
        data-ppt-clipboard-operation={clipboard?.operation ?? undefined}
        data-ppt-clipboard-paste-anchor={lastClipboardPasteEffect
          ? `${lastClipboardPasteEffect.payload.pastePlan.anchor.x},${lastClipboardPasteEffect.payload.pastePlan.anchor.y}`
          : undefined}
        data-ppt-clipboard-paste-command={lastClipboardPasteEffect?.payload.id}
        data-ppt-clipboard-paste-mapping-count={lastClipboardPasteEffect?.payload.pastePlan.mappings.length}
        data-ppt-clipboard-paste-operation={lastClipboardPasteEffect?.payload.pastePlan.operation}
        data-ppt-clipboard-paste-position-bounds-height={lastClipboardPastePositionEffect?.clipboardBounds?.h}
        data-ppt-clipboard-paste-position-bounds-width={lastClipboardPastePositionEffect?.clipboardBounds?.w}
        data-ppt-clipboard-paste-position-bounds-x={lastClipboardPastePositionEffect?.clipboardBounds?.x}
        data-ppt-clipboard-paste-position-bounds-y={lastClipboardPastePositionEffect?.clipboardBounds?.y}
        data-ppt-clipboard-paste-position-count={lastClipboardPastePositionEffect?.clipboardObjectCount}
        data-ppt-clipboard-paste-position-index={lastClipboardPastePositionEffect?.pasteIndex}
        data-ppt-clipboard-paste-position-model={lastClipboardPastePositionEffect?.model}
        data-ppt-clipboard-paste-position-viewport-x={lastClipboardPastePositionEffect?.viewportCenter?.x}
        data-ppt-clipboard-paste-position-viewport-y={lastClipboardPastePositionEffect?.viewportCenter?.y}
        data-ppt-clipboard-paste-selection={lastClipboardPasteEffect?.selection.objectIds.join(' ') ?? undefined}
        data-ppt-clipboard-paste-source-slide={lastClipboardPasteEffect?.payload.pastePlan.sourceSlideId}
        data-ppt-clipboard-paste-target-slide={lastClipboardPasteEffect?.payload.pastePlan.targetSlideId}
        data-ppt-clipboard-paste-type={lastClipboardPasteEffect?.type}
        data-ppt-clipboard-selected-object-ids={clipboard?.selectedObjectIds.join(' ') ?? undefined}
        data-ppt-clipboard-selection={clipboard?.selectedObjectIds.join(' ') ?? undefined}
        data-ppt-clipboard-source-slide={clipboard?.sourceSlideId ?? undefined}
        data-ppt-clipboard-type={clipboard?.type ?? undefined}
        data-ppt-slide-clipboard-element-count={lastSlideClipboardEffect?.elementCount}
        data-ppt-slide-clipboard-html-length={lastSlideClipboardEffect?.htmlLength}
        data-ppt-slide-clipboard-import-format={lastSlideClipboardEffect?.importFormat}
        data-ppt-slide-clipboard-imported={lastSlideClipboardEffect?.imported
          ? 'true'
          : undefined}
        data-ppt-slide-clipboard-json-mime-type={lastSlideClipboardEffect?.jsonMimeType}
        data-ppt-slide-clipboard-model={lastSlideClipboardEffect?.model}
        data-ppt-slide-clipboard-slide-name={lastSlideClipboardEffect?.slideName}
        data-ppt-slide-clipboard-source-slide={lastSlideClipboardEffect?.sourceSlideId}
        data-ppt-slide-clipboard-target-slide={lastSlideClipboardEffect?.targetSlideId}
        data-ppt-slide-clipboard-write-mode={lastSlideClipboardEffect?.writeMode}
        data-ppt-html-clipboard-html-length={lastHTMLClipboardEffect?.htmlLength}
        data-ppt-html-clipboard-json-mime-type={lastHTMLClipboardEffect?.jsonMimeType}
        data-ppt-html-clipboard-model={lastHTMLClipboardEffect?.model}
        data-ppt-html-clipboard-source-slide={lastHTMLClipboardEffect?.sourceSlideId}
        data-ppt-html-clipboard-write-mode={lastHTMLClipboardEffect?.writeMode}
        data-ppt-deck-html-import-first-slide={lastDeckHTMLImportEffect?.firstImportedSlideId}
        data-ppt-deck-html-import-format={lastDeckHTMLImportEffect?.format}
        data-ppt-deck-html-import-html-length={lastDeckHTMLImportEffect?.htmlLength}
        data-ppt-deck-html-import-imported-count={lastDeckHTMLImportEffect?.importedSlideCount}
        data-ppt-deck-html-import-model={lastDeckHTMLImportEffect?.model}
        data-ppt-deck-html-import-source-deck={lastDeckHTMLImportEffect?.sourceDeckId}
        data-ppt-deck-html-import-source-slide-count={lastDeckHTMLImportEffect?.sourceSlideCount}
        data-ppt-deck-html-import-source-title={lastDeckHTMLImportEffect?.sourceTitle}
        data-ppt-deck-json-import-first-slide={lastDeckJSONImportEffect?.firstImportedSlideId}
        data-ppt-deck-json-import-format={lastDeckJSONImportEffect?.format}
        data-ppt-deck-json-import-imported-count={lastDeckJSONImportEffect?.importedSlideCount}
        data-ppt-deck-json-import-json-length={lastDeckJSONImportEffect?.jsonLength}
        data-ppt-deck-json-import-model={lastDeckJSONImportEffect?.model}
        data-ppt-deck-json-import-source-deck={lastDeckJSONImportEffect?.sourceDeckId}
        data-ppt-deck-json-import-source-slide-count={lastDeckJSONImportEffect?.sourceSlideCount}
        data-ppt-deck-json-import-source-title={lastDeckJSONImportEffect?.sourceTitle}
        data-ppt-slide-json-import-first-slide={lastSlideJSONImportEffect?.firstImportedSlideId}
        data-ppt-slide-json-import-format={lastSlideJSONImportEffect?.format}
        data-ppt-slide-json-import-imported-count={lastSlideJSONImportEffect?.importedSlideCount}
        data-ppt-slide-json-import-json-length={lastSlideJSONImportEffect?.jsonLength}
        data-ppt-slide-json-import-model={lastSlideJSONImportEffect?.model}
        data-ppt-slide-json-import-source-slide-count={lastSlideJSONImportEffect?.sourceSlideCount}
        data-ppt-slide-json-import-source-slide-ids={lastSlideJSONImportEffect?.sourceSlideIds}
        data-ppt-slide-json-import-source-slide-names={lastSlideJSONImportEffect?.sourceSlideNames}
        data-ppt-slide-notes-import-format={lastSlideNotesImportEffect?.format}
        data-ppt-slide-notes-import-model={lastSlideNotesImportEffect?.model}
        data-ppt-slide-notes-import-notes-length={lastSlideNotesImportEffect?.notesLength}
        data-ppt-slide-notes-import-slide={lastSlideNotesImportEffect?.slideId}
        data-ppt-slide-notes-import-text-length={lastSlideNotesImportEffect?.textLength}
        data-ppt-slide-metadata-import-background={lastSlideMetadataImportEffect?.backgroundColor}
        data-ppt-slide-metadata-import-commands={lastSlideMetadataImportEffect?.commandIds}
        data-ppt-slide-metadata-import-fields={lastSlideMetadataImportEffect?.fieldIds}
        data-ppt-slide-metadata-import-format={lastSlideMetadataImportEffect?.format}
        data-ppt-slide-metadata-import-json-length={lastSlideMetadataImportEffect?.jsonLength}
        data-ppt-slide-metadata-import-model={lastSlideMetadataImportEffect?.model}
        data-ppt-slide-metadata-import-name={lastSlideMetadataImportEffect?.name}
        data-ppt-slide-metadata-import-notes-length={lastSlideMetadataImportEffect?.notesLength}
        data-ppt-slide-metadata-import-slide={lastSlideMetadataImportEffect?.slideId}
        data-ppt-slide-layout-import-command-fields={lastSlideLayoutImportEffect?.commandFields}
        data-ppt-slide-layout-import-command-types={lastSlideLayoutImportEffect?.commandTypes}
        data-ppt-slide-layout-import-commands={lastSlideLayoutImportEffect?.commandIds}
        data-ppt-slide-layout-import-fields={lastSlideLayoutImportEffect?.fields}
        data-ppt-slide-layout-import-format={lastSlideLayoutImportEffect?.format}
        data-ppt-slide-layout-import-hidden-placeholders={lastSlideLayoutImportEffect?.hiddenPlaceholderIds}
        data-ppt-slide-layout-import-json-length={lastSlideLayoutImportEffect?.jsonLength}
        data-ppt-slide-layout-import-layout={lastSlideLayoutImportEffect?.layoutId}
        data-ppt-slide-layout-import-model={lastSlideLayoutImportEffect?.model}
        data-ppt-slide-layout-import-placeholder-command-count={lastSlideLayoutImportEffect?.placeholderCommandCount}
        data-ppt-slide-layout-import-slide={lastSlideLayoutImportEffect?.slideId}
        data-ppt-slide-layout-import-theme={lastSlideLayoutImportEffect?.themeId}
        data-ppt-elements-json-import-count={lastElementsJSONImportEffect?.importedObjectCount}
        data-ppt-elements-json-import-format={lastElementsJSONImportEffect?.format}
        data-ppt-elements-json-import-json-length={lastElementsJSONImportEffect?.jsonLength}
        data-ppt-elements-json-import-model={lastElementsJSONImportEffect?.model}
        data-ppt-elements-json-import-selection={lastElementsJSONImportEffect?.selectedObjectIds.join(' ')}
        data-ppt-elements-json-import-source-slide={lastElementsJSONImportEffect?.sourceSlideId}
        data-ppt-deck-outline-import-first-slide={lastDeckMarkdownOutlineImportEffect?.firstImportedSlideId}
        data-ppt-deck-outline-import-format={lastDeckMarkdownOutlineImportEffect?.format}
        data-ppt-deck-outline-import-imported-count={lastDeckMarkdownOutlineImportEffect?.importedSlideCount}
        data-ppt-deck-outline-import-model={lastDeckMarkdownOutlineImportEffect?.model}
        data-ppt-deck-outline-import-source-slide-count={lastDeckMarkdownOutlineImportEffect?.sourceSlideCount}
        data-ppt-deck-outline-import-source-title={lastDeckMarkdownOutlineImportEffect?.sourceTitle}
        data-ppt-deck-outline-import-text-length={lastDeckMarkdownOutlineImportEffect?.textLength}
        data-ppt-slide-svg-clipboard-json-mime-type={lastSlideSVGClipboardEffect?.jsonMimeType}
        data-ppt-slide-svg-clipboard-model={lastSlideSVGClipboardEffect?.model}
        data-ppt-slide-svg-clipboard-source-slide={lastSlideSVGClipboardEffect?.sourceSlideId}
        data-ppt-slide-svg-clipboard-svg-length={lastSlideSVGClipboardEffect?.svgLength}
        data-ppt-slide-svg-clipboard-write-mode={lastSlideSVGClipboardEffect?.writeMode}
        data-ppt-selection-svg-clipboard-json-mime-type={lastSelectionSVGClipboardEffect?.jsonMimeType}
        data-ppt-selection-svg-clipboard-model={lastSelectionSVGClipboardEffect?.model}
        data-ppt-selection-svg-clipboard-selection={lastSelectionSVGClipboardEffect?.selectedObjectIds.join(' ')}
        data-ppt-selection-svg-clipboard-source-slide={lastSelectionSVGClipboardEffect?.sourceSlideId}
        data-ppt-selection-svg-clipboard-svg-length={lastSelectionSVGClipboardEffect?.svgLength}
        data-ppt-selection-svg-clipboard-write-mode={lastSelectionSVGClipboardEffect?.writeMode}
        data-ppt-table-clipboard-cols={lastTableClipboardEffect?.columnCount}
        data-ppt-table-clipboard-html-length={lastTableClipboardEffect?.htmlLength}
        data-ppt-table-clipboard-json-mime-type={lastTableClipboardEffect?.jsonMimeType}
        data-ppt-table-clipboard-model={lastTableClipboardEffect?.model}
        data-ppt-table-clipboard-object={lastTableClipboardEffect?.objectId}
        data-ppt-table-clipboard-plain-text-length={lastTableClipboardEffect?.plainTextLength}
        data-ppt-table-clipboard-rows={lastTableClipboardEffect?.rowCount}
        data-ppt-table-clipboard-source-slide={lastTableClipboardEffect?.sourceSlideId}
        data-ppt-table-clipboard-write-mode={lastTableClipboardEffect?.writeMode}
        data-ppt-table-rows-import-cols={lastTableRowsImportEffect?.columnCount}
        data-ppt-table-rows-import-command-targets={lastTableRowsImportEffect?.commandTargets}
        data-ppt-table-rows-import-format={lastTableRowsImportEffect?.format}
        data-ppt-table-rows-import-json-length={lastTableRowsImportEffect?.jsonLength}
        data-ppt-table-rows-import-model={lastTableRowsImportEffect?.model}
        data-ppt-table-rows-import-objects={lastTableRowsImportEffect?.objectIds}
        data-ppt-table-rows-import-rows={lastTableRowsImportEffect?.rowCount}
        data-ppt-rich-clipboard-formats={lastRichClipboardEffect?.formats.join(' ')}
        data-ppt-rich-clipboard-html-length={lastRichClipboardEffect?.htmlLength}
        data-ppt-rich-clipboard-import-format={lastRichClipboardEffect?.importFormat}
        data-ppt-rich-clipboard-imported={lastRichClipboardEffect?.imported ? 'true' : undefined}
        data-ppt-rich-clipboard-json-mime-type={PPT_RICH_CLIPBOARD_JSON_MIME_TYPE}
        data-ppt-rich-clipboard-model={lastRichClipboardEffect?.model}
        data-ppt-rich-clipboard-object-count={lastRichClipboardEffect?.objectCount}
        data-ppt-rich-clipboard-plain-text-length={lastRichClipboardEffect?.plainTextLength}
        data-ppt-rich-clipboard-selection={lastRichClipboardEffect?.selectedObjectIds.join(' ')}
        data-ppt-rich-clipboard-source-slide={lastRichClipboardEffect?.sourceSlideId}
        data-ppt-rich-clipboard-write-mode={lastRichClipboardEffect?.writeMode}
        data-ppt-style-clipboard-categories={styleClipboard?.categories.join(' ') ?? undefined}
        data-ppt-style-clipboard-command={lastStyleClipboardEffect?.payload.id}
        data-ppt-style-clipboard-command-applications={lastStyleClipboardEffect?.payload.id === 'paste-object-formatting'
          ? lastStyleClipboardEffect.payload.categoryApplications
              .map((application) =>
                `${application.objectId}:${application.appliedCategoryIds.join(',')}`)
              .join('|')
          : undefined}
        data-ppt-style-clipboard-command-selection={lastStyleClipboardEffect?.selection.objectIds.join(' ') ?? undefined}
        data-ppt-style-clipboard-command-slide={lastStyleClipboardEffect?.selection.slideId}
        data-ppt-style-clipboard-command-source-id={lastStyleClipboardEffect?.payload.clipboard.source.objectId}
        data-ppt-style-clipboard-command-targets={lastStyleClipboardEffect?.payload.id === 'paste-object-formatting'
          ? lastStyleClipboardEffect.payload.targetObjectIds.join(' ')
          : undefined}
        data-ppt-style-clipboard-command-type={lastStyleClipboardEffect?.type}
        data-ppt-style-clipboard-disabled-reason={styleClipboardPasteAvailability?.disabledReason}
        data-ppt-style-clipboard-model="slide-edit-style-clipboard"
        data-ppt-style-clipboard-package-categories={styleClipboard
          ? getPPTStyleClipboardPackageCategoryIds(styleClipboard).join(' ')
          : undefined}
        data-ppt-style-clipboard-source-id={styleClipboard?.sourceId ?? undefined}
        data-ppt-style-clipboard-source-kind={styleClipboard?.sourceKind ?? undefined}
        data-ppt-style-clipboard-supported-targets={styleClipboardPasteAvailability?.targets
          .filter((target) => target.isSupported)
          .map((target) => target.objectId)
          .join(' ')}
        data-ppt-style-clipboard-targets={styleClipboardPasteAvailability?.targetObjectIds.join(' ')}
        data-ppt-style-clipboard-type={styleClipboard?.type ?? undefined}
        data-ppt-object-style-import-categories={lastObjectStyleImportEffect?.categories}
        data-ppt-object-style-import-command={lastObjectStyleImportEffect?.commandId}
        data-ppt-object-style-import-command-targets={lastObjectStyleImportEffect?.commandTargets}
        data-ppt-object-style-import-command-type={lastObjectStyleImportEffect?.commandType}
        data-ppt-object-style-import-fields={lastObjectStyleImportEffect?.fields}
        data-ppt-object-style-import-format={lastObjectStyleImportEffect?.format}
        data-ppt-object-style-import-json-length={lastObjectStyleImportEffect?.jsonLength}
        data-ppt-object-style-import-model={lastObjectStyleImportEffect?.model}
        data-ppt-object-style-import-objects={lastObjectStyleImportEffect?.objectIds}
        data-ppt-object-style-import-opacity={lastObjectStyleImportEffect?.opacity}
        data-ppt-object-style-import-shadow-angle={lastObjectStyleImportEffect?.shadowAngle}
        data-ppt-object-style-import-shadow-blur={lastObjectStyleImportEffect?.shadowBlur}
        data-ppt-object-style-import-shadow-color={lastObjectStyleImportEffect?.shadowColor}
        data-ppt-object-style-import-shadow-distance={lastObjectStyleImportEffect?.shadowDistance}
        data-ppt-object-style-import-shadow-enabled={lastObjectStyleImportEffect?.shadowEnabled}
        data-ppt-object-style-import-shadow-opacity={lastObjectStyleImportEffect?.shadowOpacity}
        data-ppt-shape-style-import-categories={lastShapeStyleImportEffect?.categories}
        data-ppt-shape-style-import-command={lastShapeStyleImportEffect?.commandId}
        data-ppt-shape-style-import-command-targets={lastShapeStyleImportEffect?.commandTargets}
        data-ppt-shape-style-import-command-type={lastShapeStyleImportEffect?.commandType}
        data-ppt-shape-style-import-corner-radius={lastShapeStyleImportEffect?.cornerRadius}
        data-ppt-shape-style-import-fields={lastShapeStyleImportEffect?.fields}
        data-ppt-shape-style-import-fill-color={lastShapeStyleImportEffect?.fillColor}
        data-ppt-shape-style-import-fill-opacity={lastShapeStyleImportEffect?.fillOpacity}
        data-ppt-shape-style-import-format={lastShapeStyleImportEffect?.format}
        data-ppt-shape-style-import-json-length={lastShapeStyleImportEffect?.jsonLength}
        data-ppt-shape-style-import-model={lastShapeStyleImportEffect?.model}
        data-ppt-shape-style-import-objects={lastShapeStyleImportEffect?.objectIds}
        data-ppt-shape-style-import-stroke-color={lastShapeStyleImportEffect?.strokeColor}
        data-ppt-shape-style-import-stroke-dash={lastShapeStyleImportEffect?.strokeDash}
        data-ppt-shape-style-import-stroke-width={lastShapeStyleImportEffect?.strokeWidth}
        data-ppt-text-style-import-categories={lastTextStyleImportEffect?.categories}
        data-ppt-text-style-import-color={lastTextStyleImportEffect?.color}
        data-ppt-text-style-import-command={lastTextStyleImportEffect?.commandId}
        data-ppt-text-style-import-command-targets={lastTextStyleImportEffect?.commandTargets}
        data-ppt-text-style-import-command-type={lastTextStyleImportEffect?.commandType}
        data-ppt-text-style-import-fields={lastTextStyleImportEffect?.fields}
        data-ppt-text-style-import-font-family={lastTextStyleImportEffect?.fontFamily}
        data-ppt-text-style-import-font-size={lastTextStyleImportEffect?.fontSize}
        data-ppt-text-style-import-font-weight={lastTextStyleImportEffect?.fontWeight}
        data-ppt-text-style-import-format={lastTextStyleImportEffect?.format}
        data-ppt-text-style-import-json-length={lastTextStyleImportEffect?.jsonLength}
        data-ppt-text-style-import-model={lastTextStyleImportEffect?.model}
        data-ppt-text-style-import-objects={lastTextStyleImportEffect?.objectIds}
        data-ppt-text-style-import-paragraph-align={lastTextStyleImportEffect?.paragraphAlign}
        data-ppt-text-style-import-paragraph-bullet={lastTextStyleImportEffect?.paragraphBullet}
        data-ppt-text-style-import-paragraph-line-height={lastTextStyleImportEffect?.paragraphLineHeight}
        data-ppt-text-style-import-paragraph-spacing-after={lastTextStyleImportEffect?.paragraphSpacingAfter}
        data-ppt-text-style-import-paragraph-spacing-before={lastTextStyleImportEffect?.paragraphSpacingBefore}
        data-ppt-text-style-import-text-inset={lastTextStyleImportEffect?.textInset}
        data-ppt-text-style-import-vertical-align={lastTextStyleImportEffect?.verticalAlign}
        data-ppt-text-body-import-command-targets={lastTextBodyImportEffect?.commandTargets}
        data-ppt-text-body-import-format={lastTextBodyImportEffect?.format}
        data-ppt-text-body-import-json-length={lastTextBodyImportEffect?.jsonLength}
        data-ppt-text-body-import-mode={lastTextBodyImportEffect?.mode}
        data-ppt-text-body-import-model={lastTextBodyImportEffect?.model}
        data-ppt-text-body-import-objects={lastTextBodyImportEffect?.objectIds}
        data-ppt-text-body-import-paragraphs={lastTextBodyImportEffect?.paragraphCount}
        data-ppt-text-body-import-runs={lastTextBodyImportEffect?.runCount}
        data-ppt-text-body-import-text-length={lastTextBodyImportEffect?.textLength}
        data-ppt-line-style-import-categories={lastLineStyleImportEffect?.categories}
        data-ppt-line-style-import-command={lastLineStyleImportEffect?.commandId}
        data-ppt-line-style-import-command-targets={lastLineStyleImportEffect?.commandTargets}
        data-ppt-line-style-import-command-type={lastLineStyleImportEffect?.commandType}
        data-ppt-line-style-import-fields={lastLineStyleImportEffect?.fields}
        data-ppt-line-style-import-format={lastLineStyleImportEffect?.format}
        data-ppt-line-style-import-json-length={lastLineStyleImportEffect?.jsonLength}
        data-ppt-line-style-import-model={lastLineStyleImportEffect?.model}
        data-ppt-line-style-import-objects={lastLineStyleImportEffect?.objectIds}
        data-ppt-line-style-import-stroke-color={lastLineStyleImportEffect?.strokeColor}
        data-ppt-line-style-import-stroke-dash={lastLineStyleImportEffect?.strokeDash}
        data-ppt-line-style-import-stroke-width={lastLineStyleImportEffect?.strokeWidth}
        data-ppt-object-transform-import-command-targets={lastObjectTransformImportEffect?.commandTargets}
        data-ppt-object-transform-import-fields={lastObjectTransformImportEffect?.fields}
        data-ppt-object-transform-import-format={lastObjectTransformImportEffect?.format}
        data-ppt-object-transform-import-h={lastObjectTransformImportEffect?.h}
        data-ppt-object-transform-import-json-length={lastObjectTransformImportEffect?.jsonLength}
        data-ppt-object-transform-import-model={lastObjectTransformImportEffect?.model}
        data-ppt-object-transform-import-objects={lastObjectTransformImportEffect?.objectIds}
        data-ppt-object-transform-import-rotation={lastObjectTransformImportEffect?.rotation}
        data-ppt-object-transform-import-w={lastObjectTransformImportEffect?.w}
        data-ppt-object-transform-import-x={lastObjectTransformImportEffect?.x}
        data-ppt-object-transform-import-y={lastObjectTransformImportEffect?.y}
        data-ppt-comment-import-body-length={lastCommentImportEffect?.bodyLength}
        data-ppt-comment-import-command-targets={lastCommentImportEffect?.commandTargets}
        data-ppt-comment-import-created-at={lastCommentImportEffect?.createdAt}
        data-ppt-comment-import-fields={lastCommentImportEffect?.fields}
        data-ppt-comment-import-format={lastCommentImportEffect?.format}
        data-ppt-comment-import-json-length={lastCommentImportEffect?.jsonLength}
        data-ppt-comment-import-message-count={lastCommentImportEffect?.messageCount}
        data-ppt-comment-import-model={lastCommentImportEffect?.model}
        data-ppt-comment-import-objects={lastCommentImportEffect?.objectIds}
        data-ppt-comment-import-resolved={lastCommentImportEffect?.resolved}
        data-ppt-object-metadata-import-alt-text-length={lastObjectMetadataImportEffect?.altTextLength}
        data-ppt-object-metadata-import-alt-text-present={lastObjectMetadataImportEffect?.altTextPresent}
        data-ppt-object-metadata-import-command-fields={lastObjectMetadataImportEffect?.commandFields}
        data-ppt-object-metadata-import-command-types={lastObjectMetadataImportEffect?.commandTypes}
        data-ppt-object-metadata-import-commands={lastObjectMetadataImportEffect?.commandIds}
        data-ppt-object-metadata-import-fields={lastObjectMetadataImportEffect?.fields}
        data-ppt-object-metadata-import-format={lastObjectMetadataImportEffect?.format}
        data-ppt-object-metadata-import-hyperlink-url={lastObjectMetadataImportEffect?.hyperlinkUrl}
        data-ppt-object-metadata-import-json-length={lastObjectMetadataImportEffect?.jsonLength}
        data-ppt-object-metadata-import-model={lastObjectMetadataImportEffect?.model}
        data-ppt-object-metadata-import-name={lastObjectMetadataImportEffect?.name}
        data-ppt-object-metadata-import-objects={lastObjectMetadataImportEffect?.objectIds}
        data-ppt-object-metadata-import-slide={lastObjectMetadataImportEffect?.slideId}
        data-ppt-object-state-import-command-types={lastObjectStateImportEffect?.commandTypes}
        data-ppt-object-state-import-commands={lastObjectStateImportEffect?.commandIds}
        data-ppt-object-state-import-fields={lastObjectStateImportEffect?.fields}
        data-ppt-object-state-import-format={lastObjectStateImportEffect?.format}
        data-ppt-object-state-import-json-length={lastObjectStateImportEffect?.jsonLength}
        data-ppt-object-state-import-locked={lastObjectStateImportEffect?.locked}
        data-ppt-object-state-import-lock-targets={lastObjectStateImportEffect?.lockTargets}
        data-ppt-object-state-import-model={lastObjectStateImportEffect?.model}
        data-ppt-object-state-import-objects={lastObjectStateImportEffect?.objectIds}
        data-ppt-object-state-import-slide={lastObjectStateImportEffect?.slideId}
        data-ppt-object-state-import-visible={lastObjectStateImportEffect?.visible}
        data-ppt-object-state-import-visibility-targets={lastObjectStateImportEffect?.visibilityTargets}
        data-ppt-object-layer-import-command={lastObjectLayerImportEffect?.commandId}
        data-ppt-object-layer-import-command-type={lastObjectLayerImportEffect?.commandType}
        data-ppt-object-layer-import-fields={lastObjectLayerImportEffect?.fields}
        data-ppt-object-layer-import-format={lastObjectLayerImportEffect?.format}
        data-ppt-object-layer-import-from-index={lastObjectLayerImportEffect?.fromIndex}
        data-ppt-object-layer-import-json-length={lastObjectLayerImportEffect?.jsonLength}
        data-ppt-object-layer-import-model={lastObjectLayerImportEffect?.model}
        data-ppt-object-layer-import-object={lastObjectLayerImportEffect?.objectId}
        data-ppt-object-layer-import-position={lastObjectLayerImportEffect?.position}
        data-ppt-object-layer-import-slide={lastObjectLayerImportEffect?.slideId}
        data-ppt-object-layer-import-to-index={lastObjectLayerImportEffect?.toIndex}
        data-ppt-import-extension={PPT_IMPORT_EXTENSION.id}
        data-ppt-import-extension-last-clipboard-actions={lastClipboardImportActionKinds}
        data-ppt-import-extension-last-drop-action={lastStageDropImportActionKind}
        data-ppt-import-extension-clipboard-action-order={PPT_IMPORT_EXTENSION.clipboardActionOrder.join(' ')}
        data-ppt-import-extension-drop-action-order={PPT_IMPORT_EXTENSION.dropActionOrder.join(' ')}
        data-ppt-import-extension-install-unit={PPT_IMPORT_EXTENSION.installUnit}
        data-ppt-fallback-html-import-format={lastFallbackHTMLImportEffect?.format}
        data-ppt-fallback-html-import-kind={lastFallbackHTMLImportEffect?.kind}
        data-ppt-fallback-html-import-model={lastFallbackHTMLImportEffect?.model ?? PPT_FALLBACK_HTML_IMPORT_MODEL}
        data-ppt-fallback-html-import-name={lastFallbackHTMLImportEffect?.name}
        data-ppt-fallback-html-import-count={lastFallbackHTMLImportEffect?.objectCount}
        data-ppt-fallback-html-import-shape={lastFallbackHTMLImportEffect?.shape}
        data-ppt-fallback-html-import-source-object={lastFallbackHTMLImportEffect?.sourceObjectId}
        data-ppt-fallback-html-import-source-objects={lastFallbackHTMLImportEffect?.sourceObjectIds?.join(' ')}
        data-ppt-media-import-importer={lastMediaImport?.importerId}
        data-ppt-media-import-model={PPT_MEDIA_IMPORT_MODEL}
        data-ppt-media-import-selection={lastMediaImport?.item.id}
        data-ppt-media-import-url={lastMediaImport?.source.url}
        data-ppt-media-json-import-fields={lastMediaJSONImportEffect?.fields}
        data-ppt-media-json-import-format={lastMediaJSONImportEffect?.format}
        data-ppt-media-json-import-importer={lastMediaJSONImportEffect?.importerId}
        data-ppt-media-json-import-json-length={lastMediaJSONImportEffect?.jsonLength}
        data-ppt-media-json-import-model={lastMediaJSONImportEffect?.model}
        data-ppt-media-json-import-object={lastMediaJSONImportEffect?.objectId}
        data-ppt-media-json-import-title={lastMediaJSONImportEffect?.title}
        data-ppt-media-json-import-url={lastMediaJSONImportEffect?.url}
        data-ppt-inline-edit-element={lastInlineEditEffect?.elementId}
        data-ppt-inline-edit-history-direction={lastInlineEditEffect?.historyDirection}
        data-ppt-inline-edit-input-type={lastInlineEditEffect?.inputType}
        data-ppt-inline-edit-line-break={lastInlineEditEffect?.lineBreak ? 'true' : undefined}
        data-ppt-inline-edit-model={PPT_INLINE_EDIT_DOM_MODEL}
        data-ppt-inline-edit-paste-text={lastInlineEditEffect?.pasteText}
        data-ppt-text-paste-bold-runs={lastTextPasteImport?.boldRunCount}
        data-ppt-text-paste-bullet-paragraphs={lastTextPasteImport?.bulletParagraphCount}
        data-ppt-text-paste-format={lastTextPasteImport?.format}
        data-ppt-text-paste-hyperlink-url={lastTextPasteImport?.hyperlinkUrl}
        data-ppt-text-paste-importer={lastTextPasteImport?.importerId}
        data-ppt-text-paste-link-runs={lastTextPasteImport?.linkRunCount}
        data-ppt-text-paste-model={PPT_TEXT_PASTE_IMPORT_MODEL}
        data-ppt-text-paste-numbered-paragraphs={lastTextPasteImport?.numberedParagraphCount}
        data-ppt-text-paste-selection={lastTextPasteImport?.item.id}
        data-ppt-text-paste-underline-runs={lastTextPasteImport?.underlineRunCount}
        data-ppt-marquee-active={interaction?.kind === 'marquee' ? 'true' : 'false'}
        data-ppt-marquee-additive={interaction?.kind === 'marquee'
          ? String(interaction.additive)
          : undefined}
        data-ppt-marquee-h={marqueeBounds?.h}
        data-ppt-marquee-history="none"
        data-ppt-marquee-model={PPT_MARQUEE_SELECTION_MODEL}
        data-ppt-marquee-selection={marqueeSelection?.join(' ') ?? undefined}
        data-ppt-marquee-w={marqueeBounds?.w}
        data-ppt-marquee-x={marqueeBounds?.x}
        data-ppt-marquee-y={marqueeBounds?.y}
        data-ppt-color-swatch-command={lastColorSwatchEffect?.payload.id}
        data-ppt-color-swatch-command-channel={lastColorSwatchEffect?.payload.channelId}
        data-ppt-color-swatch-command-objects={lastColorSwatchEffect?.payload.objectIds.join(' ')}
        data-ppt-color-swatch-command-slide={lastColorSwatchEffect?.payload.slideId}
        data-ppt-color-swatch-command-source={lastColorSwatchEffect?.payload.swatch.source}
        data-ppt-color-swatch-command-swatch={lastColorSwatchEffect?.payload.swatch.swatchId}
        data-ppt-color-swatch-command-token={lastColorSwatchEffect?.payload.swatch.tokenId}
        data-ppt-color-swatch-command-type={lastColorSwatchEffect?.type}
        data-ppt-color-swatch-command-value={lastColorSwatchEffect?.payload.swatch.value}
        data-ppt-color-swatch-model="slide-edit-color-swatch-palette"
        data-ppt-placeholder-visibility-command={lastPlaceholderVisibilityEffect?.payload.id}
        data-ppt-placeholder-visibility-command-placeholder={lastPlaceholderVisibilityEffect?.payload.placeholderId}
        data-ppt-placeholder-visibility-command-selection={lastPlaceholderVisibilityEffect?.selection.objectIds.join(' ') ?? undefined}
        data-ppt-placeholder-visibility-command-slide={lastPlaceholderVisibilityEffect?.payload.slideId}
        data-ppt-placeholder-visibility-command-type={lastPlaceholderVisibilityEffect?.type}
        data-ppt-placeholder-visibility-command-visible={lastPlaceholderVisibilityEffect
          ? String(lastPlaceholderVisibilityEffect.payload.isVisible)
          : undefined}
        data-ppt-keyboard-nudge-enabled={commandAvailability.nudge ? 'true' : 'false'}
        data-ppt-keyboard-nudge-intent={PPT_KEYBOARD_NUDGE_INTENT_MODEL}
        data-ppt-keyboard-nudge-keys={PPT_KEYBOARD_NUDGE_KEYS}
        data-ppt-keyboard-nudge-large-step={String(PPT_KEYBOARD_NUDGE_LARGE_STEP)}
        data-ppt-keyboard-nudge-model={PPT_KEYBOARD_NUDGE_MODEL}
        data-ppt-keyboard-nudge-step={String(PPT_KEYBOARD_NUDGE_STEP)}
        data-ppt-arrow-tool-model={PPT_TOOL_AFFORDANCES.arrow.model}
        data-ppt-arrow-tool-shortcut={PPT_TOOL_AFFORDANCES.arrow.shortcut}
        data-ppt-drawing-tool={creationTool?.kind === 'freeform'
          ? creationTool.tool
          : undefined}
        data-ppt-eraser-hit-count={interaction?.kind === 'erase'
          ? interaction.erasedIds.length
          : undefined}
        data-ppt-eraser-tool-active={isEraserToolActive ? 'true' : 'false'}
        data-ppt-eraser-tool-model={PPT_TOOL_AFFORDANCES.eraser.model}
        data-ppt-eraser-tool-shortcut={PPT_TOOL_AFFORDANCES.eraser.shortcut}
        data-ppt-highlighter-tool-model={PPT_TOOL_AFFORDANCES.highlight.model}
        data-ppt-highlighter-tool-shortcut={PPT_TOOL_AFFORDANCES.highlight.shortcut}
        data-ppt-laser-tool-active={isLaserToolActive ? 'true' : 'false'}
        data-ppt-laser-tool-model={PPT_TOOL_AFFORDANCES.laser.model}
        data-ppt-laser-tool-shortcut={PPT_TOOL_AFFORDANCES.laser.shortcut}
        data-ppt-laser-trail-model={PPT_LASER_TRAIL_OVERLAY_MODEL}
        data-ppt-laser-trail-point-count={laserTrailPoints.length}
        data-ppt-laser-trail-state={interaction?.kind === 'laser'
          ? 'active'
          : laserTrailPoints.length > 0 ? 'idle' : 'empty'}
        data-ppt-pan-tool-active={isPanToolActive ? 'true' : 'false'}
        data-ppt-pan-tool-model={PPT_TOOL_AFFORDANCES.pan.model}
        data-ppt-pan-tool-shortcut={PPT_TOOL_AFFORDANCES.pan.shortcut}
        data-ppt-marker-tool-model={PPT_TOOL_AFFORDANCES.marker.model}
        data-ppt-marker-tool-shortcut={PPT_TOOL_AFFORDANCES.marker.shortcut}
        data-ppt-section-tool-model={PPT_TOOL_AFFORDANCES.section.model}
        data-ppt-section-tool-shortcut={PPT_TOOL_AFFORDANCES.section.shortcut}
        data-ppt-resize-aspect-ratio-modifier="Shift"
        data-ppt-resize-from-center-modifier="Alt"
        data-ppt-resize-modifier-model={PPT_RESIZE_POINTER_MODIFIERS_MODEL}
        data-ppt-alignment-popover-preview={alignmentPreviewCommand ?? undefined}
        data-ppt-alignment-popover-preview-model={CANVAS_DOM_ALIGNMENT_PREVIEW_GUIDE_MODEL}
        data-ppt-comment-thread-command={lastCommentThreadEffect?.payload.id}
        data-ppt-comment-thread-command-body={lastCommentThreadEffect?.payload.body}
        data-ppt-comment-thread-command-count={lastCommentThreadEffect?.payload.messageCount}
        data-ppt-comment-thread-command-object={lastCommentThreadEffect?.payload.objectId}
        data-ppt-comment-thread-command-selection={lastCommentThreadEffect?.selection.objectIds.join(' ') ?? undefined}
        data-ppt-comment-thread-command-slide={lastCommentThreadEffect?.payload.slideId}
        data-ppt-comment-thread-command-type={lastCommentThreadEffect?.type}
        data-ppt-comment-thread-model={PPT_COMMENT_THREAD_MODEL}
        data-ppt-accessibility-command={lastAccessibilityEffect?.payload.id}
        data-ppt-accessibility-command-field={lastAccessibilityEffect?.payload.id === 'update-object-accessibility'
          ? lastAccessibilityEffect.payload.fieldId
          : undefined}
        data-ppt-accessibility-command-object={lastAccessibilityEffect?.payload.objectId}
        data-ppt-accessibility-command-slide={lastAccessibilityEffect?.payload.slideId}
        data-ppt-accessibility-command-type={lastAccessibilityEffect?.type}
        data-ppt-accessibility-command-value={lastAccessibilityEffect?.payload.id === 'update-object-accessibility'
          ? String(lastAccessibilityEffect.payload.value)
          : undefined}
        data-ppt-accessibility-model="slide-edit-object-accessibility"
        data-ppt-corner-radius-command={lastCornerRadiusEffect?.payload.id}
        data-ppt-corner-radius-command-field={lastCornerRadiusEffect?.payload.fieldId}
        data-ppt-corner-radius-command-object={lastCornerRadiusEffect?.payload.objectId}
        data-ppt-corner-radius-command-slide={lastCornerRadiusEffect?.payload.slideId}
        data-ppt-corner-radius-command-type={lastCornerRadiusEffect?.type}
        data-ppt-corner-radius-command-value={lastCornerRadiusEffect?.payload.value}
        data-ppt-corner-radius-model="slide-edit-object-corner-radius"
        data-ppt-fill-opacity-command={lastFillOpacityEffect?.payload.id}
        data-ppt-fill-opacity-command-field={lastFillOpacityEffect?.payload.fieldId}
        data-ppt-fill-opacity-command-object={lastFillOpacityEffect?.payload.objectId}
        data-ppt-fill-opacity-command-slide={lastFillOpacityEffect?.payload.slideId}
        data-ppt-fill-opacity-command-type={lastFillOpacityEffect?.type}
        data-ppt-fill-opacity-command-value={lastFillOpacityEffect?.payload.value}
        data-ppt-fill-opacity-model="slide-edit-object-fill-opacity"
        data-ppt-hyperlink-command={lastHyperlinkEffect?.payload.id}
        data-ppt-hyperlink-command-field={lastHyperlinkEffect?.payload.id === 'update-object-hyperlink'
          ? lastHyperlinkEffect.payload.fieldId
          : undefined}
        data-ppt-hyperlink-command-object={lastHyperlinkEffect?.payload.objectId}
        data-ppt-hyperlink-command-slide={lastHyperlinkEffect?.payload.slideId}
        data-ppt-hyperlink-command-type={lastHyperlinkEffect?.type}
        data-ppt-hyperlink-command-value={lastHyperlinkEffect?.payload.id === 'update-object-hyperlink'
          ? lastHyperlinkEffect.payload.value
          : undefined}
        data-ppt-hyperlink-model="slide-edit-object-hyperlink"
        data-ppt-image-crop-command={lastImageCropEffect?.payload.id}
        data-ppt-image-crop-command-crop-x={lastImageCropEffect?.payload.id === 'reset-object-image-crop'
          ? lastImageCropEffect.payload.crop.x
          : undefined}
        data-ppt-image-crop-command-crop-y={lastImageCropEffect?.payload.id === 'reset-object-image-crop'
          ? lastImageCropEffect.payload.crop.y
          : undefined}
        data-ppt-image-crop-command-field={lastImageCropEffect?.payload.id === 'update-object-image-crop'
          ? lastImageCropEffect.payload.fieldId
          : undefined}
        data-ppt-image-crop-command-fit={lastImageCropEffect?.payload.id === 'reset-object-image-crop'
          ? lastImageCropEffect.payload.fit
          : undefined}
        data-ppt-image-crop-command-object={lastImageCropEffect?.payload.objectId}
        data-ppt-image-crop-command-slide={lastImageCropEffect?.payload.slideId}
        data-ppt-image-crop-command-type={lastImageCropEffect?.type}
        data-ppt-image-crop-command-value={lastImageCropEffect?.payload.id === 'update-object-image-crop'
          ? String(lastImageCropEffect.payload.value)
          : undefined}
        data-ppt-image-crop-import-command-fields={lastImageCropImportEffect?.commandFields}
        data-ppt-image-crop-import-command-types={lastImageCropImportEffect?.commandTypes}
        data-ppt-image-crop-import-commands={lastImageCropImportEffect?.commandIds}
        data-ppt-image-crop-import-fields={lastImageCropImportEffect?.fields}
        data-ppt-image-crop-import-fit={lastImageCropImportEffect?.fit}
        data-ppt-image-crop-import-format={lastImageCropImportEffect?.format}
        data-ppt-image-crop-import-json-length={lastImageCropImportEffect?.jsonLength}
        data-ppt-image-crop-import-model={lastImageCropImportEffect?.model}
        data-ppt-image-crop-import-objects={lastImageCropImportEffect?.objectIds}
        data-ppt-image-crop-import-slide={lastImageCropImportEffect?.slideId}
        data-ppt-image-crop-import-x={lastImageCropImportEffect?.x}
        data-ppt-image-crop-import-y={lastImageCropImportEffect?.y}
        data-ppt-image-crop-model="slide-edit-object-image-crop"
        data-ppt-image-import-count={lastImageImportEffect?.count}
        data-ppt-image-import-format={lastImageImportEffect?.format}
        data-ppt-image-import-mime={lastImageImportEffect?.mimeType}
        data-ppt-image-import-model={PPT_IMAGE_IMPORT_MODEL}
        data-ppt-image-import-name={lastImageImportEffect?.name}
        data-ppt-image-import-names={lastImageImportEffect?.names}
        data-ppt-image-import-natural-height={lastImageImportEffect?.naturalHeight}
        data-ppt-image-import-natural-width={lastImageImportEffect?.naturalWidth}
        data-ppt-image-replace-command={lastImageReplaceEffect?.payload.id}
        data-ppt-image-replace-command-mime={lastImageReplaceEffect?.payload.source.mimeType}
        data-ppt-image-replace-command-name={lastImageReplaceEffect?.payload.source.name}
        data-ppt-image-replace-command-natural-height={lastImageReplaceEffect?.payload.source.naturalHeight}
        data-ppt-image-replace-command-natural-width={lastImageReplaceEffect?.payload.source.naturalWidth}
        data-ppt-image-replace-command-object={lastImageReplaceEffect?.payload.objectId}
        data-ppt-image-replace-command-slide={lastImageReplaceEffect?.payload.slideId}
        data-ppt-image-replace-command-src-prefix={lastImageReplaceEffect?.payload.source.src.slice(0, 19)}
        data-ppt-image-replace-command-type={lastImageReplaceEffect?.type}
        data-ppt-image-replace-import-alt-text-length={lastImageReplaceImportEffect?.altTextLength}
        data-ppt-image-replace-import-command={lastImageReplaceImportEffect?.commandId}
        data-ppt-image-replace-import-command-type={lastImageReplaceImportEffect?.commandType}
        data-ppt-image-replace-import-fields={lastImageReplaceImportEffect?.fields}
        data-ppt-image-replace-import-format={lastImageReplaceImportEffect?.format}
        data-ppt-image-replace-import-json-length={lastImageReplaceImportEffect?.jsonLength}
        data-ppt-image-replace-import-mime={lastImageReplaceImportEffect?.mimeType}
        data-ppt-image-replace-import-model={lastImageReplaceImportEffect?.model}
        data-ppt-image-replace-import-name={lastImageReplaceImportEffect?.name}
        data-ppt-image-replace-import-natural-height={lastImageReplaceImportEffect?.naturalHeight}
        data-ppt-image-replace-import-natural-width={lastImageReplaceImportEffect?.naturalWidth}
        data-ppt-image-replace-import-object={lastImageReplaceImportEffect?.objectId}
        data-ppt-image-replace-import-slide={lastImageReplaceImportEffect?.slideId}
        data-ppt-image-replace-import-src-prefix={lastImageReplaceImportEffect?.srcPrefix}
        data-ppt-image-replace-model="slide-edit-object-image-replace"
        data-ppt-object-animation-command={lastObjectAnimationEffect?.payload.id}
        data-ppt-object-animation-command-field={lastObjectAnimationEffect?.payload.fieldId}
        data-ppt-object-animation-command-object={lastObjectAnimationEffect?.payload.objectId}
        data-ppt-object-animation-command-slide={lastObjectAnimationEffect?.payload.slideId}
        data-ppt-object-animation-command-type={lastObjectAnimationEffect?.type}
        data-ppt-object-animation-command-value={lastObjectAnimationEffect
          ? String(lastObjectAnimationEffect.payload.value)
          : undefined}
        data-ppt-object-animation-build-order={activeSlideAnimationBuildOrder.join(' ')}
        data-ppt-object-animation-build-order-model="slide-edit-object-animation-build-order"
        data-ppt-object-animation-import-command-fields={lastObjectAnimationImportEffect?.commandFields}
        data-ppt-object-animation-import-commands={lastObjectAnimationImportEffect?.commandIds}
        data-ppt-object-animation-import-delay={lastObjectAnimationImportEffect?.delayMs}
        data-ppt-object-animation-import-duration={lastObjectAnimationImportEffect?.durationMs}
        data-ppt-object-animation-import-fields={lastObjectAnimationImportEffect?.fields}
        data-ppt-object-animation-import-format={lastObjectAnimationImportEffect?.format}
        data-ppt-object-animation-import-json-length={lastObjectAnimationImportEffect?.jsonLength}
        data-ppt-object-animation-import-model={lastObjectAnimationImportEffect?.model}
        data-ppt-object-animation-import-objects={lastObjectAnimationImportEffect?.objectIds}
        data-ppt-object-animation-import-order={lastObjectAnimationImportEffect?.order}
        data-ppt-object-animation-import-slide={lastObjectAnimationImportEffect?.slideId}
        data-ppt-object-animation-import-trigger={lastObjectAnimationImportEffect?.trigger}
        data-ppt-object-animation-import-type={lastObjectAnimationImportEffect?.type}
        data-ppt-object-animation-model="slide-edit-object-animation"
        data-ppt-object-opacity-command={lastObjectOpacityEffect?.payload.id}
        data-ppt-object-opacity-command-field={lastObjectOpacityEffect?.payload.fieldId}
        data-ppt-object-opacity-command-object={lastObjectOpacityEffect?.payload.objectId}
        data-ppt-object-opacity-command-slide={lastObjectOpacityEffect?.payload.slideId}
        data-ppt-object-opacity-command-type={lastObjectOpacityEffect?.type}
        data-ppt-object-opacity-command-value={lastObjectOpacityEffect?.payload.value}
        data-ppt-object-opacity-model="slide-edit-object-opacity"
        data-ppt-object-visibility-command={lastObjectVisibilityEffect?.payload.id}
        data-ppt-object-visibility-command-objects={lastObjectVisibilityEffect?.payload.objectIds.join(' ') ?? undefined}
        data-ppt-object-visibility-command-selection={lastObjectVisibilityEffect?.selection.objectIds.join(' ') ?? undefined}
        data-ppt-object-visibility-command-slide={lastObjectVisibilityEffect?.selection.slideId}
        data-ppt-object-visibility-command-target-count={lastObjectVisibilityEffect?.payload.objectIds.length}
        data-ppt-object-visibility-command-type={lastObjectVisibilityEffect?.type}
        data-ppt-object-visibility-model="slide-edit-object-visibility"
        data-ppt-transition-command={lastSlideTransitionEffect?.payload.id}
        data-ppt-transition-command-field={lastSlideTransitionEffect?.payload.fieldId}
        data-ppt-transition-command-selection={lastSlideTransitionEffect?.selection.objectIds.join(' ') ?? undefined}
        data-ppt-transition-command-slide={lastSlideTransitionEffect?.selection.slideId}
        data-ppt-transition-command-type={lastSlideTransitionEffect?.type}
        data-ppt-transition-command-value={lastSlideTransitionEffect
          ? getPPTSlideTransitionCommandValue(lastSlideTransitionEffect.payload)
          : undefined}
        data-ppt-transition-import-advance-after={lastSlideTransitionImportEffect?.advanceAfterMs}
        data-ppt-transition-import-advance-on-click={lastSlideTransitionImportEffect?.advanceOnClick}
        data-ppt-transition-import-command-fields={lastSlideTransitionImportEffect?.commandFields}
        data-ppt-transition-import-commands={lastSlideTransitionImportEffect?.commandIds}
        data-ppt-transition-import-duration={lastSlideTransitionImportEffect?.durationMs}
        data-ppt-transition-import-fields={lastSlideTransitionImportEffect?.fields}
        data-ppt-transition-import-format={lastSlideTransitionImportEffect?.format}
        data-ppt-transition-import-json-length={lastSlideTransitionImportEffect?.jsonLength}
        data-ppt-transition-import-model={lastSlideTransitionImportEffect?.model}
        data-ppt-transition-import-slide={lastSlideTransitionImportEffect?.slideId}
        data-ppt-transition-import-type={lastSlideTransitionImportEffect?.type}
        data-ppt-transition-model="slide-edit-slide-transition-timing"
        data-ppt-shadow-command={lastShadowEffect?.payload.id}
        data-ppt-shadow-command-field={lastShadowEffect?.payload.fieldId}
        data-ppt-shadow-command-object={lastShadowEffect?.payload.objectId}
        data-ppt-shadow-command-slide={lastShadowEffect?.payload.slideId}
        data-ppt-shadow-command-type={lastShadowEffect?.type}
        data-ppt-shadow-command-value={lastShadowEffect
          ? String(lastShadowEffect.payload.value)
          : undefined}
        data-ppt-shadow-model="slide-edit-object-shadow"
        data-ppt-stroke-line-style-command={lastStrokeLineStyleEffect?.payload.id}
        data-ppt-stroke-line-style-command-field={lastStrokeLineStyleEffect?.payload.fieldId}
        data-ppt-stroke-line-style-command-object={lastStrokeLineStyleEffect?.payload.objectId}
        data-ppt-stroke-line-style-command-slide={lastStrokeLineStyleEffect?.payload.slideId}
        data-ppt-stroke-line-style-command-type={lastStrokeLineStyleEffect?.type}
        data-ppt-stroke-line-style-command-value={lastStrokeLineStyleEffect?.payload.value}
        data-ppt-stroke-line-style-model="slide-edit-object-stroke-line-style"
        data-ppt-text-font-family-command={lastTextFontFamilyEffect?.payload.id}
        data-ppt-text-font-family-command-field={lastTextFontFamilyEffect?.payload.fieldId}
        data-ppt-text-font-family-command-object={lastTextFontFamilyEffect?.payload.objectId}
        data-ppt-text-font-family-command-slide={lastTextFontFamilyEffect?.payload.slideId}
        data-ppt-text-font-family-command-type={lastTextFontFamilyEffect?.type}
        data-ppt-text-font-family-command-value={lastTextFontFamilyEffect?.payload.value}
        data-ppt-text-font-family-model="slide-edit-text-font-family"
        data-ppt-text-inset-command={lastTextFrameInsetEffect?.payload.id}
        data-ppt-text-inset-command-field={lastTextFrameInsetEffect?.payload.fieldId}
        data-ppt-text-inset-command-object={lastTextFrameInsetEffect?.payload.objectId}
        data-ppt-text-inset-command-slide={lastTextFrameInsetEffect?.payload.slideId}
        data-ppt-text-inset-command-type={lastTextFrameInsetEffect?.type}
        data-ppt-text-inset-command-value={lastTextFrameInsetEffect?.payload.value}
        data-ppt-text-inset-model="slide-edit-text-frame-inset"
        data-ppt-text-paragraph-spacing-command={lastTextParagraphSpacingEffect?.payload.id}
        data-ppt-text-paragraph-spacing-command-field={lastTextParagraphSpacingEffect?.payload.fieldId}
        data-ppt-text-paragraph-spacing-command-object={lastTextParagraphSpacingEffect?.payload.objectId}
        data-ppt-text-paragraph-spacing-command-slide={lastTextParagraphSpacingEffect?.payload.slideId}
        data-ppt-text-paragraph-spacing-command-type={lastTextParagraphSpacingEffect?.type}
        data-ppt-text-paragraph-spacing-command-unit={lastTextParagraphSpacingEffect?.payload.fieldId === 'lineHeightRatio'
          ? undefined
          : lastTextParagraphSpacingEffect?.payload.value.unit}
        data-ppt-text-paragraph-spacing-command-value={lastTextParagraphSpacingEffect
          ? lastTextParagraphSpacingEffect.payload.fieldId === 'lineHeightRatio'
            ? lastTextParagraphSpacingEffect.payload.value
            : lastTextParagraphSpacingEffect.payload.value.value
          : undefined}
        data-ppt-text-paragraph-spacing-model="slide-edit-text-paragraph-spacing"
        data-ppt-text-vertical-align-command={lastTextVerticalAlignmentEffect?.payload.id}
        data-ppt-text-vertical-align-command-field={lastTextVerticalAlignmentEffect?.payload.fieldId}
        data-ppt-text-vertical-align-command-object={lastTextVerticalAlignmentEffect?.payload.objectId}
        data-ppt-text-vertical-align-command-slide={lastTextVerticalAlignmentEffect?.payload.slideId}
        data-ppt-text-vertical-align-command-type={lastTextVerticalAlignmentEffect?.type}
        data-ppt-text-vertical-align-command-value={lastTextVerticalAlignmentEffect?.payload.value}
        data-ppt-text-vertical-align-model="slide-edit-text-vertical-alignment"
        data-ppt-text-autofit-command={lastTextAutoFitEffect?.payload.id}
        data-ppt-text-autofit-command-handle={lastTextAutoFitEffect?.payload.handle}
        data-ppt-text-autofit-command-height={lastTextAutoFitEffect?.payload.bounds.h}
        data-ppt-text-autofit-command-object={lastTextAutoFitEffect?.payload.objectId}
        data-ppt-text-autofit-command-selection={lastTextAutoFitEffect?.selection.objectIds.join(' ') ?? undefined}
        data-ppt-text-autofit-command-slide={lastTextAutoFitEffect?.selection.slideId}
        data-ppt-text-autofit-command-size-mode={lastTextAutoFitEffect?.payload.sizeMode}
        data-ppt-text-autofit-command-type={lastTextAutoFitEffect?.type}
        data-ppt-text-autofit-command-width={lastTextAutoFitEffect?.payload.bounds.w}
        data-ppt-text-autofit-model="slide-edit-text-box-auto-fit"
        data-ppt-text-autofit-size-modes={SLIDE_EDIT_TEXT_BOX_SIZE_MODES
          .map((mode) => mode.id)
          .join(' ')}
        data-ppt-text-autofit-import-command-handles={lastTextAutoFitImportEffect?.commandHandles}
        data-ppt-text-autofit-import-command-targets={lastTextAutoFitImportEffect?.commandTargets}
        data-ppt-text-autofit-import-command-types={lastTextAutoFitImportEffect?.commandTypes}
        data-ppt-text-autofit-import-commands={lastTextAutoFitImportEffect?.commandIds}
        data-ppt-text-autofit-import-fields={lastTextAutoFitImportEffect?.fields}
        data-ppt-text-autofit-import-format={lastTextAutoFitImportEffect?.format}
        data-ppt-text-autofit-import-json-length={lastTextAutoFitImportEffect?.jsonLength}
        data-ppt-text-autofit-import-mode={lastTextAutoFitImportEffect?.mode}
        data-ppt-text-autofit-import-model={lastTextAutoFitImportEffect?.model}
        data-ppt-text-autofit-import-objects={lastTextAutoFitImportEffect?.objectIds}
        data-ppt-resize-handle-click-double={lastResizeHandleClickMemoryEffect?.isDoubleClick ? 'true' : undefined}
        data-ppt-resize-handle-click-handle={lastResizeHandleClickMemoryEffect?.handle}
        data-ppt-resize-handle-click-id={lastResizeHandleClickMemoryEffect?.id}
        data-ppt-resize-handle-click-model={PPT_POINTER_CLICK_MEMORY_MODEL}
        data-ppt-resize-handle-click-x={lastResizeHandleClickMemoryEffect?.point.x}
        data-ppt-resize-handle-click-y={lastResizeHandleClickMemoryEffect?.point.y}
        data-ppt-table-import-cols={lastTableImportEffect?.columnCount}
        data-ppt-table-import-count={lastTableImportEffect?.count}
        data-ppt-table-import-format={lastTableImportEffect?.format}
        data-ppt-table-import-model={PPT_TABLE_IMPORT_MODEL}
        data-ppt-table-import-name={lastTableImportEffect?.name}
        data-ppt-table-import-names={lastTableImportEffect?.names}
        data-ppt-table-import-rows={lastTableImportEffect?.rowCount}
        data-ppt-text-overflow-indicator-anchor={selectedTextAutoFitIndicator?.anchor}
        data-ppt-text-overflow-indicator-axis={selectedTextAutoFitIndicator?.overflowAxis.join(' ')}
        data-ppt-text-overflow-indicator-height={selectedTextAutoFitIndicator?.bounds.h}
        data-ppt-text-overflow-indicator-line-count={selectedTextAutoFitIndicator?.lineCount}
        data-ppt-text-overflow-indicator-model="slide-edit-text-box-auto-fit"
        data-ppt-text-overflow-indicator-object={selectedTextAutoFitIndicator?.objectId}
        data-ppt-text-overflow-indicator-size-mode={selectedTextAutoFitIndicator?.sizeMode}
        data-ppt-text-overflow-indicator-slide={selectedTextAutoFitIndicator?.slideId}
        data-ppt-text-overflow-indicator-visible={selectedTextAutoFitIndicator
          ? String(selectedTextAutoFitIndicator.isVisible)
          : undefined}
        data-ppt-text-overflow-indicator-width={selectedTextAutoFitIndicator?.bounds.w}
        data-ppt-keyboard-command-dispatch={PPT_KEYBOARD_COMMAND_DISPATCH_MODEL}
        data-ppt-keyboard-command-intent={PPT_KEYBOARD_COMMAND_INTENT_MODEL}
        data-ppt-keyboard-tool-dispatch={PPT_KEYBOARD_TOOL_DISPATCH_MODEL}
        data-ppt-keyboard-viewport-intent={PPT_KEYBOARD_VIEWPORT_INTENT_MODEL}
        data-ppt-keyboard-viewport-model={PPT_KEYBOARD_VIEWPORT_MODEL}
        data-ppt-sticky-tool-model={PPT_TOOL_AFFORDANCES.sticky.model}
        data-ppt-sticky-tool-shortcut={PPT_TOOL_AFFORDANCES.sticky.shortcut}
        data-ppt-temporary-pan-active={isTemporaryPanActive ? 'true' : 'false'}
        data-ppt-temporary-pan-gesture={interaction?.kind === 'pan' ? 'true' : 'false'}
        data-ppt-temporary-pan-model={PPT_KEYBOARD_TEMPORARY_PAN_MODEL}
        data-ppt-temporary-pan-shortcut={PPT_KEYBOARD_TEMPORARY_PAN_SHORTCUT_LABEL}
        data-ppt-wheel-viewport-horizontal-pan-modifier={PPT_WHEEL_VIEWPORT_HORIZONTAL_PAN_MODIFIER}
        data-ppt-wheel-viewport-model={PPT_WHEEL_VIEWPORT_MODEL}
        data-ppt-wheel-viewport-pan={PPT_WHEEL_VIEWPORT_PAN_MODE}
        data-ppt-wheel-viewport-zoom-modifier={PPT_WHEEL_VIEWPORT_ZOOM_MODIFIER}
        data-ppt-recent-colors={recentColors.join(' ')}
        data-ppt-recent-color-count={recentColors.length}
        data-creation-tool={getPPTCreationToolDataValue(creationTool)}
        data-frame-guides={showFrameGuides ? 'true' : 'false'}
        data-grid={showGrid ? 'true' : 'false'}
        data-line-tool={lineCreationMode ?? undefined}
        data-minimap={showMinimap ? 'true' : 'false'}
        onContextMenu={handleStageContextMenu}
        onDragOver={handleStageDragOver}
        onDrop={handleStageDrop}
        onPointerDown={handleStagePointerDown}
        onPointerMove={handlePointerMove}
        onPointerUp={handlePointerUp}
        ref={setPPTStageElementRef}
        tabIndex={-1}
      >
        <div
          className="ppt-stage-world"
          style={{
            transform: `translate(${viewport.x}px, ${viewport.y}px) scale(${viewport.scale})`,
          }}
        >
          <div
            className="ppt-slide"
            data-ppt-hidden-placeholders={(activeSlide.hiddenPlaceholderIds ?? []).join(' ')}
            data-ppt-layout-id={activeLayout.layoutId}
            data-ppt-slide={activeSlide.id}
            data-ppt-theme-id={activeSlide.themeId ?? PPT_THEME_DESCRIPTOR.themeId}
            data-ppt-transition-advance-after={activeSlideTransition.advanceAfterMs ?? ''}
            data-ppt-transition-advance-on-click={activeSlideTransition.advanceOnClick ? 'true' : 'false'}
            data-ppt-transition-duration={activeSlideTransition.durationMs}
            data-ppt-transition-model="slide-edit-slide-transition-timing"
            data-ppt-transition-type={activeSlideTransition.type}
            style={{ background: activeSlide.background?.color ?? '#ffffff' }}
          >
            {activeSlide.elements.filter((element) => element.visible !== false).map((element) => (
              <PPTElementView
                editing={editingId === element.id}
                element={element}
                eraserHit={eraserHitIds?.has(element.id) === true}
                hovered={hoveredId === element.id}
                findActive={findOpen && activeFindMatch?.elementId === element.id}
                key={element.id}
                selected={selection.includes(element.id)}
                slideId={activeSlide.id}
                textAutoFitIndicator={selection.includes(element.id)
                  ? selectedTextAutoFitIndicator
                  : null}
                textOverflow={textOverflowById[element.id] === true}
                onCommitText={commitText}
                onEdit={() => {
                  if (isPPTTextElement(element)) {
                    setEditingId(element.id)
                    setSelection([element.id])
                    setContextMenu(null)
                  }
                }}
                onContextMenu={handleElementContextMenu}
                onPointerDown={handleElementPointerDown}
                onPointerEnter={() => setHoveredId(element.id)}
                onPointerLeave={() => setHoveredId((current) => current === element.id ? null : current)}
                onInlineEditEffect={setLastInlineEditEffect}
                onStopEdit={() => setEditingId(null)}
                onTextOverflowChange={updateTextOverflowState}
              />
            ))}
            {selection.length > 0 && !editingId && showFrameGuides ? (
              <FrameGuides geometry={frameGuideGeometry} />
            ) : null}
            {selectedBounds ? (
              <SelectionOverlay
                bounds={selectedBounds}
                canResize={canResizeSelection}
                scale={viewport.scale}
                selectedElements={selectedElements}
                textAutoFitIndicator={selectedTextAutoFitIndicator}
                textOverflow={selectedTextOverflow}
                onRotatePointerDown={handleRotatePointerDown}
                onResizePointerDown={handleResizePointerDown}
              />
            ) : null}
            <PPTSelectionFloatingBar
              anchor={selectionCommandAnchor}
              commandAvailability={commandAvailability}
              groups={selectionFloatingCommandGroups}
              scale={viewport.scale}
              shapeMenu={shapeQuickMenuState}
              textFormat={textQuickFormatState}
              onCommand={runPPTSurfaceCommand}
              onAlignmentPreviewChange={setAlignmentPreviewCommand}
              onFontSizeStep={stepSelectedTextFontSize}
              onParagraphBulletToggle={toggleSelectedParagraphBullet}
              onParagraphNumberedToggle={toggleSelectedParagraphNumbered}
              onParagraphAlign={updateSelectedParagraphAlign}
              onShapeKindChange={updateShapeKind}
              onTextBoldToggle={toggleSelectedTextBold}
              onTextColorChange={updateSelectedTextColor}
              onTextItalicToggle={toggleSelectedTextItalic}
              onTextUnderlineToggle={toggleSelectedTextUnderline}
            />
            {selectedLineElement && !editingId && canResizeSelection ? (
              <LineEndpointOverlay
                line={selectedLineElement}
                scale={viewport.scale}
                onPointerDown={handleLineEndpointPointerDown}
              />
            ) : null}
            {selectedLineElement &&
            (selectedLineElement.route ?? 'straight') === 'elbow' &&
            !editingId &&
            canResizeSelection ? (
              <LineRouteOverlay
                line={selectedLineElement}
                scale={viewport.scale}
                onPointerDown={handleLineRoutePointerDown}
              />
            ) : null}
            {marqueeBounds ? <Box className="ppt-marquee" bounds={marqueeBounds} /> : null}
            <Guides guides={snapGuides} scale={viewport.scale} />
            {laserTrailPoints.length > 0 ? (
              <PPTLaserTrailOverlay points={laserTrailPoints} scale={viewport.scale} />
            ) : null}
          </div>
        </div>
        <PPTContextCommandMenu
          groups={contextCommandGroups}
          menu={contextMenu}
          onClose={() => setContextMenu(null)}
          onCommand={runPPTSurfaceCommand}
        />
        <PPTMinimap
          model={minimapModel}
          onNavigateToWorldPoint={navigateMinimapToWorldPoint}
        />
      </section>

      <Inspector
        exportCode={exportCode}
        inspectorSurface={inspectorSurface}
        layoutDescriptors={PPT_LAYOUT_DESCRIPTORS}
        layoutPlaceholderVisibilityDescriptors={activeLayoutPlaceholderVisibilityDescriptors}
        layoutPlaceholders={activeLayoutPlaceholders}
        lastPlaceholderVisibilityEffect={lastPlaceholderVisibilityEffect}
        lastTextAutoFitEffect={lastTextAutoFitEffect}
        recentColors={recentColors}
        selection={selection}
        selectedElement={selectedElement}
        selectedElementAnimation={selectedElementAnimation}
        slide={activeSlide}
        slideMetadataDescriptor={slideMetadataDescriptor}
        slideLayoutId={activeLayout.layoutId}
        slideThemeId={activeSlide.themeId ?? PPT_THEME_DESCRIPTOR.themeId}
        slideTransition={activeSlideTransition}
        themeColorTokens={PPT_THEME_DESCRIPTOR.colorTokens}
        onCommentBodyChange={updateCommentBody}
        onCommentReplyAdd={addCommentReply}
        onCommentResolvedChange={updateCommentResolved}
        onCommitText={commitText}
        onCopyHTML={copyHTML}
        onColorSwatchApply={applyColorSwatch}
        onDownloadHTML={downloadHTML}
        onElementAltTextChange={updateElementAltText}
        onElementAnimationChange={updateElementAnimation}
        onElementGeometryChange={updateElementGeometry}
        onElementHyperlinkChange={updateElementHyperlink}
        onImageCropChange={updateImageCrop}
        onImageCropReset={resetImageCrop}
        onImageFitChange={updateImageFit}
        onImageReplaceFile={replacePPTImageFile}
        onElementNameChange={updateElementName}
        onElementOpacityChange={updateElementOpacity}
        onElementRotationChange={updateElementRotation}
        onElementShadowChange={updateElementShadow}
        onElementTextInsetChange={updateElementTextInset}
        onObjectVisibilityCommandEffect={applyObjectVisibilityCommandEffect}
        onLineMarkerChange={updateLineMarker}
        onLineRouteChange={updateLineRoute}
        onParagraphBulletChange={updateParagraphBullet}
        onParagraphNumberedChange={updateParagraphNumbered}
        onParagraphSpacingChange={updateParagraphSpacing}
        onElementTextStyleChange={updateElementTextStyle}
        onTextAutoFit={autoFitTextElement}
        onLayerPaneCommandEffect={applyLayerPaneCommandEffect}
        onLayoutPlaceholderVisibilityChange={updateLayoutPlaceholderVisibility}
        onParagraphAlignChange={updateParagraphAlign}
        onShapeCornerRadiusChange={updateShapeCornerRadius}
        onShapeKindChange={updateShapeKind}
        onSlideBackgroundChange={updateSlideBackground}
        onSlideLayoutChange={applySlideLayout}
        onShapeFillChange={updateShapeFill}
        onElementStrokeChange={updateElementStroke}
        onSlideNameChange={updateSlideName}
        onSlideNotesChange={updateSlideNotes}
        onSlideTransitionChange={updateSlideTransition}
        onTableRowsChange={updateTableRows}
        textAutoFitIndicator={selectedTextAutoFitIndicator}
        selectedTextOverflow={selectedTextOverflow}
      />
      <PPTPresentationOverlay
        slide={presentationSlide}
        slideCount={deck.slides.length}
        slideIndex={presentationSlideIndex}
        onExit={exitPresentation}
        onNext={() => navigatePresentation(1)}
        onPrevious={() => navigatePresentation(-1)}
      />
      <PPTCommandPalette
        items={commandPaletteItems}
        open={commandPaletteOpen}
        onClose={closeCommandPalette}
      />
      <PPTShortcutHelpOverlay
        items={shortcutHelpItems}
        open={shortcutHelpOpen}
        onClose={closeShortcutHelp}
      />
    </main>
  )
}

function PPTMinimap({
  model,
  onNavigateToWorldPoint,
}: {
  model: PPTMinimapReadModel | null
  onNavigateToWorldPoint: (point: Point) => void
}) {
  const svgRef = useRef<SVGSVGElement | null>(null)
  const [activePointerId, setActivePointerId] = useState<number | null>(null)

  if (!model) {
    return null
  }

  const readModel = model
  const itemRects = readModel.itemRects.filter(
    (item) => item.id !== PPT_MINIMAP_SLIDE_FRAME_ID,
  )

  function navigate(event: ReactPointerEvent<SVGSVGElement>) {
    const localGeometry = getPPTCanvasPointerLocalGeometry({
      event,
      target: svgRef.current,
    })

    if (!localGeometry) {
      return
    }

    const point = getPPTMinimapPointFromViewportOffset({
      model: readModel,
      offset: localGeometry.point,
      viewportSize: {
        h: localGeometry.rect.height,
        w: localGeometry.rect.width,
      },
    })

    if (!point) {
      return
    }

    onNavigateToWorldPoint(getPPTMinimapWorldPoint({ model: readModel, point }))
  }

  function handlePointerDown(event: ReactPointerEvent<SVGSVGElement>) {
    event.preventDefault()
    event.stopPropagation()
    capturePPTCanvasPointerFromEvent(event)
    setActivePointerId(event.pointerId)
    navigate(event)
  }

  function handlePointerMove(event: ReactPointerEvent<SVGSVGElement>) {
    if (activePointerId !== event.pointerId) {
      return
    }

    event.preventDefault()
    event.stopPropagation()
    navigate(event)
  }

  function clearPointer(event: ReactPointerEvent<SVGSVGElement>) {
    if (activePointerId === event.pointerId) {
      setActivePointerId(null)
    }
  }

  return (
    <div
      className="ppt-minimap"
      data-ppt-minimap
      data-ppt-minimap-item-count={itemRects.length}
      data-ppt-minimap-model={PPT_MINIMAP_READ_MODEL}
      data-ppt-minimap-scale={readModel.scale}
      data-ppt-minimap-viewport-h={readModel.viewportWorldBounds.h}
      data-ppt-minimap-viewport-w={readModel.viewportWorldBounds.w}
      data-ppt-minimap-viewport-x={readModel.viewportWorldBounds.x}
      data-ppt-minimap-viewport-y={readModel.viewportWorldBounds.y}
    >
      <svg
        aria-label="PPT minimap"
        className="ppt-minimap-map"
        data-ppt-minimap-map
        ref={svgRef}
        role="img"
        viewBox={`0 0 ${readModel.size.w} ${readModel.size.h}`}
        onPointerCancel={clearPointer}
        onPointerDown={handlePointerDown}
        onPointerMove={handlePointerMove}
        onPointerUp={clearPointer}
      >
        <rect
          className="ppt-minimap-world"
          data-ppt-minimap-world
          height={readModel.displayBounds.h}
          rx="3"
          width={readModel.displayBounds.w}
          x={readModel.displayBounds.x}
          y={readModel.displayBounds.y}
        />
        {itemRects.map((item) => (
          <rect
            className="ppt-minimap-item"
            data-ppt-minimap-item={item.id}
            height={item.rect.h}
            key={item.id}
            rx="1.5"
            width={item.rect.w}
            x={item.rect.x}
            y={item.rect.y}
          />
        ))}
        <rect
          className="ppt-minimap-viewport"
          data-ppt-minimap-viewport
          height={readModel.viewportRect.h}
          rx="2"
          width={readModel.viewportRect.w}
          x={readModel.viewportRect.x}
          y={readModel.viewportRect.y}
        />
      </svg>
    </div>
  )
}

function PPTShortcutHelpOverlay({
  items,
  onClose,
  open,
}: {
  items: readonly PPTShortcutHelpItem[]
  onClose: () => void
  open: boolean
}) {
  if (!open) {
    return null
  }

  return <PPTShortcutHelpDialog items={items} onClose={onClose} />
}

function PPTShortcutHelpDialog({
  items,
  onClose,
}: {
  items: readonly PPTShortcutHelpItem[]
  onClose: () => void
}) {
  const closeButtonRef = useRef<HTMLButtonElement | null>(null)
  const dialogRef = useRef<HTMLElement | null>(null)
  const groups = useMemo(() => groupPPTShortcutHelpItems(items), [items])

  usePPTCanvasModalFocusLifecycle({
    initialFocusRef: closeButtonRef,
  })

  function handleBackdropMouseDown(event: ReactMouseEvent<HTMLDivElement>) {
    const backdropPointerIntent = getPPTCanvasModalBackdropPointerIntent({
      currentTarget: event.currentTarget,
      target: event.target,
    })

    if (backdropPointerIntent.kind === 'dismiss') {
      onClose()
    }
  }

  function handleKeyDown(event: ReactKeyboardEvent<HTMLElement>) {
    const modalKeyboardIntent = getPPTCanvasModalKeyboardIntent({ key: event.key })

    if (modalKeyboardIntent.kind === 'close') {
      if (modalKeyboardIntent.preventDefault) {
        event.preventDefault()
      }
      if (modalKeyboardIntent.stopPropagation) {
        event.stopPropagation()
      }
      onClose()
      return
    }

    if (modalKeyboardIntent.kind === 'trap-focus') {
      trapPPTCanvasModalTabFocus({
        event,
        root: dialogRef.current,
      })
    }
  }

  return (
    <div
      className="ppt-shortcut-help-backdrop"
      data-ppt-shortcut-help-backdrop
      onMouseDown={handleBackdropMouseDown}
    >
      <section
        aria-label="Keyboard shortcuts"
        aria-modal="true"
        className="ppt-shortcut-help"
        data-ppt-shortcut-help
        data-ppt-shortcut-help-focus-lifecycle={PPT_MODAL_FOCUS_LIFECYCLE_MODEL}
        ref={dialogRef}
        role="dialog"
        onKeyDown={handleKeyDown}
      >
        <header className="ppt-shortcut-help-header">
          <h2>Keyboard shortcuts</h2>
          <button
            aria-label="Close keyboard shortcuts"
            className="ppt-icon-button"
            data-ppt-shortcut-help-close
            ref={closeButtonRef}
            title="Close"
            type="button"
            onClick={onClose}
          >
            <X size={16} />
          </button>
        </header>
        <div className="ppt-shortcut-help-sections">
          {groups.map((group) => (
            <section
              aria-label={group.section}
              className="ppt-shortcut-help-section"
              data-ppt-shortcut-help-section={group.section}
              key={group.section}
            >
              <h3>{group.section}</h3>
              <dl className="ppt-shortcut-help-list">
                {group.items.map((item) => (
                  <div
                    className="ppt-shortcut-help-row"
                    data-ppt-shortcut-help-item={item.id}
                    key={item.id}
                  >
                    <dt>{item.title}</dt>
                    <dd>
                      <kbd data-ppt-shortcut-help-shortcut={item.shortcut}>
                        {item.shortcut}
                      </kbd>
                    </dd>
                  </div>
                ))}
              </dl>
            </section>
          ))}
        </div>
      </section>
    </div>
  )
}

function PPTPresentationOverlay({
  onExit,
  onNext,
  onPrevious,
  slide,
  slideCount,
  slideIndex,
}: {
  onExit: () => void
  onNext: () => void
  onPrevious: () => void
  slide: PPTSlide | null
  slideCount: number
  slideIndex: number
}) {
  const scale = usePPTPresentationScale()

  if (!slide) {
    return null
  }

  const readableIndex = slideIndex + 1
  const transition = getPPTSlideTransition(slide)
  const transitionStyle = getSlideEditTransitionCSSStyle({
    durationMs: transition.durationMs,
  })
  const visibleAnimationBuildOrder = getPPTSlideAnimationBuildOrder(slide, {
    visibleOnly: true,
  })

  return (
    <div
      aria-label="Presentation preview"
      aria-modal="true"
      className="ppt-presentation"
      data-ppt-presentation
      data-ppt-presentation-advance-after={transition.advanceAfterMs ?? ''}
      data-ppt-presentation-advance-on-click={transition.advanceOnClick ? 'true' : 'false'}
      data-ppt-presentation-animation-build-order={visibleAnimationBuildOrder.join(' ')}
      data-ppt-presentation-animation-build-order-model="slide-edit-object-animation-build-order"
      data-ppt-presentation-index={`${readableIndex}/${slideCount}`}
      data-ppt-presentation-slide={slide.id}
      data-ppt-presentation-transition={transition.type}
      data-ppt-presentation-transition-duration={transition.durationMs}
      role="dialog"
    >
      <div className="ppt-presentation-header">
        <span data-ppt-presentation-title>{slide.name}</span>
        <span data-ppt-presentation-count>{readableIndex}/{slideCount}</span>
        <button
          className="ppt-icon-button"
          data-ppt-present-exit
          title="Exit presentation"
          type="button"
          onClick={onExit}
        >
          <X size={17} />
        </button>
      </div>
      <div className="ppt-presentation-stage">
        <div
          className="ppt-presentation-slide-shell"
          style={{
            height: PPT_SLIDE_HEIGHT * scale,
            width: PPT_SLIDE_WIDTH * scale,
          }}
        >
          <div
            key={slide.id}
            className="ppt-slide ppt-presentation-slide"
            data-ppt-presentation-slide-frame
            data-ppt-transition-advance-after={transition.advanceAfterMs ?? ''}
            data-ppt-transition-advance-on-click={transition.advanceOnClick ? 'true' : 'false'}
            data-ppt-transition-duration={transition.durationMs}
            data-ppt-transition-type={transition.type}
            style={{
              '--ppt-presentation-scale': String(scale),
              ...transitionStyle,
              background: slide.background?.color ?? '#ffffff',
              transform: `scale(${scale})`,
            } as CSSProperties}
          >
            {slide.elements.filter((element) => element.visible !== false).map((element) => (
              <PPTElementView
                editing={false}
                element={element}
                eraserHit={false}
                findActive={false}
                hovered={false}
                key={element.id}
                selected={false}
                slideId={slide.id}
                textAutoFitIndicator={null}
                textOverflow={false}
                onCommitText={() => undefined}
                onContextMenu={(event) => event.preventDefault()}
                onEdit={() => undefined}
                onInlineEditEffect={() => undefined}
                onPointerDown={() => undefined}
                onPointerEnter={() => undefined}
                onPointerLeave={() => undefined}
                onStopEdit={() => undefined}
                onTextOverflowChange={() => undefined}
              />
            ))}
          </div>
        </div>
      </div>
      <div className="ppt-presentation-controls" role="toolbar" aria-label="Presentation controls">
        <button
          className="ppt-button"
          data-ppt-present-prev
          type="button"
          onClick={onPrevious}
        >
          <ChevronLeft size={16} /> Previous
        </button>
        <button
          className="ppt-button"
          data-ppt-present-next
          type="button"
          onClick={onNext}
        >
          Next <ChevronRight size={16} />
        </button>
      </div>
    </div>
  )
}

function usePPTPresentationScale() {
  const [scale, setScale] = useState(() => getPPTPresentationScale())

  useEffect(() => {
    function updateScale() {
      setScale(getPPTPresentationScale())
    }

    updateScale()
    const cleanup = bindPPTCanvasEventListener({
      listener: updateScale,
      target: window,
      type: 'resize',
    })

    return () => {
      cleanup()
    }
  }, [])

  return scale
}

function getPPTPresentationScale() {
  const viewportSize = getPPTCanvasClientViewportSize()

  if (!viewportSize) {
    return 0.75
  }

  return clampPPTCanvasValue(
    Math.min(
      (viewportSize.width - 96) / PPT_SLIDE_WIDTH,
      (viewportSize.height - 168) / PPT_SLIDE_HEIGHT,
    ),
    0.2,
    1.4,
  )
}

function PPTCommandPalette({
  items,
  onClose,
  open,
}: {
  items: readonly PPTCommandPaletteItem[]
  onClose: () => void
  open: boolean
}) {
  if (!open) {
    return null
  }

  return <PPTCommandPaletteDialog items={items} onClose={onClose} />
}

function PPTCommandPaletteDialog({
  items,
  onClose,
}: {
  items: readonly PPTCommandPaletteItem[]
  onClose: () => void
}) {
  const [query, setQuery] = useState('')
  const [activeIndex, setActiveIndex] = useState(0)
  const dialogRef = useRef<HTMLElement>(null)
  const inputRef = useRef<HTMLInputElement>(null)
  const filteredItems = useMemo(
    () => filterPPTCommandPaletteItems(items, query).slice(0, 10),
    [items, query],
  )
  const maxActiveIndex = Math.max(0, filteredItems.length - 1)
  const activeItemIndex = Math.min(activeIndex, maxActiveIndex)
  const activeItem = filteredItems[activeItemIndex]
  const listboxId = 'ppt-command-palette-listbox'
  const activeOptionId = activeItem
    ? getPPTCommandPaletteOptionId(activeItem.id)
    : undefined

  usePPTCanvasModalFocusLifecycle({
    initialFocusRef: inputRef,
  })

  function runItem(item: PPTCommandPaletteItem | undefined) {
    if (!item || item.disabled) {
      return
    }

    item.onSelect()
    onClose()
  }

  function handleKeyDown(event: ReactKeyboardEvent<HTMLElement>) {
    const modalKeyboardIntent = getPPTCanvasModalKeyboardIntent({ key: event.key })

    if (modalKeyboardIntent.kind === 'close') {
      if (modalKeyboardIntent.preventDefault) {
        event.preventDefault()
      }
      if (modalKeyboardIntent.stopPropagation) {
        event.stopPropagation()
      }
      onClose()
      return
    }

    if (modalKeyboardIntent.kind === 'trap-focus') {
      trapPPTCanvasModalTabFocus({
        event,
        root: dialogRef.current,
      })
      return
    }

    const keyboardIntent = getPPTCanvasCommandPaletteKeyboardIntent({
      activeIndex: activeItemIndex,
      itemCount: filteredItems.length,
      key: event.key,
    })

    if (keyboardIntent.preventDefault) {
      event.preventDefault()
      event.stopPropagation()

      if (keyboardIntent.kind === 'move-active') {
        setActiveIndex(keyboardIntent.activeIndex)
        return
      }

      if (keyboardIntent.kind === 'run-active') {
        runItem(filteredItems[keyboardIntent.activeIndex])
      }
    }
  }

  function handleBackdropMouseDown(event: ReactMouseEvent<HTMLDivElement>) {
    const backdropPointerIntent = getPPTCanvasModalBackdropPointerIntent({
      currentTarget: event.currentTarget,
      target: event.target,
    })

    if (backdropPointerIntent.kind === 'dismiss') {
      onClose()
    }
  }

  return (
    <div
      className="ppt-command-palette-backdrop"
      data-ppt-command-palette-backdrop
      onMouseDown={handleBackdropMouseDown}
    >
      <section
        aria-label="Command palette"
        aria-modal="true"
        className="ppt-command-palette"
        data-ppt-command-palette
        data-ppt-command-palette-focus-lifecycle={PPT_MODAL_FOCUS_LIFECYCLE_MODEL}
        data-ppt-command-palette-focus-trap="true"
        data-ppt-command-palette-model={PPT_COMMAND_PALETTE_ITEMS_MODEL}
        data-ppt-command-palette-restore-focus="true"
        ref={dialogRef}
        role="dialog"
        onKeyDown={handleKeyDown}
      >
        <input
          aria-activedescendant={activeOptionId}
          aria-autocomplete="list"
          aria-controls={listboxId}
          aria-expanded="true"
          aria-label="Search commands"
          className="ppt-command-palette-input"
          data-ppt-command-palette-active-option={activeOptionId}
          data-ppt-command-palette-combobox="true"
          data-ppt-command-palette-controls={listboxId}
          data-ppt-command-palette-query
          placeholder="Search commands"
          ref={inputRef}
          role="combobox"
          value={query}
          onChange={(event) => {
            setQuery(event.target.value)
            setActiveIndex(0)
          }}
        />
        <div
          aria-label="Command results"
          className="ppt-command-palette-list"
          data-ppt-command-palette-active-option={activeOptionId}
          data-ppt-command-palette-listbox
          id={listboxId}
          role="listbox"
        >
          {filteredItems.length > 0 ? filteredItems.map((item, index) => (
            <button
              aria-disabled={item.disabled ? 'true' : undefined}
              aria-selected={index === activeItemIndex}
              className="ppt-command-palette-item"
              data-ppt-command-palette-active={index === activeItemIndex ? 'true' : undefined}
              data-ppt-command-palette-item={item.id}
              data-ppt-command-palette-option-id={getPPTCommandPaletteOptionId(item.id)}
              disabled={item.disabled}
              id={getPPTCommandPaletteOptionId(item.id)}
              key={item.id}
              role="option"
              type="button"
              onClick={() => runItem(item)}
              onMouseEnter={() => setActiveIndex(index)}
            >
              <span className="ppt-command-palette-item-main">
                <span className="ppt-command-palette-item-title">{item.title}</span>
                <span className="ppt-command-palette-item-section">{item.section}</span>
              </span>
              {item.shortcut ? (
                <kbd className="ppt-command-palette-shortcut">{item.shortcut}</kbd>
              ) : null}
            </button>
          )) : (
            <div className="ppt-command-palette-empty" data-ppt-command-palette-empty>
              No matches
            </div>
          )}
        </div>
      </section>
    </div>
  )
}

function getPPTCommandPaletteOptionId(itemId: string) {
  return `ppt-command-palette-option-${itemId.replace(/[^A-Za-z0-9_-]/g, '-')}`
}

function getPPTShortcutHelpItems(
  items: readonly PPTCommandPaletteItem[],
): PPTShortcutHelpItem[] {
  return items.flatMap((item) =>
    item.shortcut
      ? [{
          id: item.id,
          section: item.section,
          shortcut: item.shortcut,
          title: item.title,
        }]
      : [],
  )
}

function groupPPTShortcutHelpItems(
  items: readonly PPTShortcutHelpItem[],
): PPTShortcutHelpSectionGroup[] {
  const orderedSections = [
    ...PPT_SHORTCUT_HELP_SECTION_ORDER,
    ...items
      .map((item) => item.section)
      .filter((section) => !PPT_SHORTCUT_HELP_SECTION_ORDER.includes(section)),
  ]

  return orderedSections.flatMap((section) => {
    const sectionItems = items.filter((item) => item.section === section)

    return sectionItems.length > 0
      ? [{ items: sectionItems, section }]
      : []
  })
}

function getPPTSlideTransition(slide: PPTSlide): PPTSlideTransition {
  return normalizePPTSlideTransition(
    slide.transition ?? PPT_DEFAULT_SLIDE_TRANSITION,
    slide.id,
  )
}

function normalizePPTSlideTransition(
  transition: PPTSlideTransition,
  slideId = 'ppt-slide',
): PPTSlideTransition {
  return toPPTSlideTransition(
    getPPTSlideTransitionDescriptor(slideId, transition),
  )
}

function getPPTSlideTransitionDescriptor(
  slideId: string,
  transition: PPTSlideTransition,
): PPTSlideTransitionDescriptor {
  return createSlideEditTransitionDescriptor({
    advance: {
      afterMs: transition.advanceAfterMs ?? undefined,
      onClick: transition.advanceOnClick ?? true,
    },
    durationMs: transition.durationMs,
    slideId,
    type: isPPTSlideTransitionType(transition.type)
      ? transition.type
      : PPT_DEFAULT_SLIDE_TRANSITION.type,
  })
}

function toPPTSlideTransition(
  descriptor: PPTSlideTransitionDescriptor,
): PPTSlideTransition {
  return {
    advanceAfterMs: descriptor.advance.afterMs ?? null,
    advanceOnClick: descriptor.advance.onClick,
    durationMs: descriptor.durationMs,
    type: isPPTSlideTransitionType(descriptor.type)
      ? descriptor.type
      : PPT_DEFAULT_SLIDE_TRANSITION.type,
  }
}

function isPPTSlideTransitionType(
  value: string,
): value is PPTSlideTransitionType {
  return (PPT_SLIDE_TRANSITION_TYPES as readonly string[]).includes(value)
}

function getPPTSlideTransitionUpdateCommand({
  field,
  slideId,
  transition,
  value,
}: {
  field: PPTSlideTransitionUpdateField
  slideId: string
  transition: PPTSlideTransition
  value: PPTSlideTransition[PPTSlideTransitionUpdateField]
}): PPTSlideTransitionUpdateCommand {
  switch (field) {
    case 'type':
      return {
        fieldId: 'type',
        id: 'update-slide-transition',
        slideId,
        value: isPPTSlideTransitionType(String(value))
          ? String(value) as PPTSlideTransitionType
          : PPT_DEFAULT_SLIDE_TRANSITION.type,
      }
    case 'durationMs':
      return {
        fieldId: 'durationMs',
        id: 'update-slide-transition',
        slideId,
        value: Number(value),
      }
    case 'advanceOnClick':
      return {
        fieldId: 'advance',
        id: 'update-slide-transition',
        slideId,
        value: {
          afterMs: transition.advanceAfterMs ?? undefined,
          onClick: Boolean(value),
        },
      }
    case 'advanceAfterMs':
      return {
        fieldId: 'advance',
        id: 'update-slide-transition',
        slideId,
        value: {
          afterMs: value === null ? undefined : Number(value),
          onClick: transition.advanceOnClick ?? true,
        },
      }
  }
}

function applyPPTSlideTransitionUpdateCommand(
  transition: PPTSlideTransition,
  command: PPTSlideTransitionUpdateCommand,
): PPTSlideTransition {
  switch (command.fieldId) {
    case 'advance':
      return toPPTSlideTransition(createSlideEditTransitionDescriptor({
        advance: command.value,
        durationMs: transition.durationMs,
        slideId: command.slideId,
        type: transition.type,
      }))
    case 'durationMs':
      return toPPTSlideTransition(createSlideEditTransitionDescriptor({
        advance: {
          afterMs: transition.advanceAfterMs ?? undefined,
          onClick: transition.advanceOnClick,
        },
        durationMs: command.value,
        slideId: command.slideId,
        type: transition.type,
      }))
    case 'type':
      return toPPTSlideTransition(createSlideEditTransitionDescriptor({
        advance: {
          afterMs: transition.advanceAfterMs ?? undefined,
          onClick: transition.advanceOnClick,
        },
        durationMs: transition.durationMs,
        slideId: command.slideId,
        type: command.value,
      }))
  }
}

function getPPTSlideTransitionCommandValue(
  command: PPTSlideTransitionUpdateCommand,
) {
  switch (command.fieldId) {
    case 'advance':
      return [
        `onClick:${command.value.onClick ? 'true' : 'false'}`,
        command.value.afterMs === undefined
          ? 'afterMs:'
          : `afterMs:${command.value.afterMs}`,
      ].join(';')
    case 'durationMs':
    case 'type':
      return String(command.value)
  }
}

function clampPPTSlideTransitionDuration(value: number) {
  return createSlideEditTransitionDescriptor({
    durationMs: value,
    slideId: 'ppt-transition-parser',
  }).durationMs
}

function parsePPTSlideTransitionDuration(value: string) {
  return clampPPTSlideTransitionDuration(Number(value))
}

function parsePPTSlideTransitionAdvanceAfter(value: string) {
  return value.trim() === ''
    ? null
    : createSlideEditTransitionDescriptor({
        advance: {
          afterMs: Number(value),
          onClick: true,
        },
        slideId: 'ppt-transition-parser',
      }).advance.afterMs ?? null
}

function getPPTElementAnimation(
  element: PPTElement,
  slide?: PPTSlide,
): PPTElementAnimation {
  return normalizePPTElementAnimation(
    element.animation ?? {
      ...PPT_DEFAULT_ELEMENT_ANIMATION,
      order: getPPTElementDefaultAnimationOrder(element.id, slide),
    },
    slide,
    element.id,
  )
}

function normalizePPTElementAnimation(
  animation: PPTElementAnimation,
  slide?: PPTSlide,
  elementId?: string,
): PPTElementAnimation {
  return {
    delayMs: clampPPTElementAnimationTime(animation.delayMs),
    durationMs: clampPPTElementAnimationTime(animation.durationMs),
    order: clampPPTElementAnimationOrder(
      animation.order,
      slide?.elements.length,
      elementId ? getPPTElementDefaultAnimationOrder(elementId, slide) : undefined,
    ),
    trigger: PPT_ELEMENT_ANIMATION_TRIGGERS.includes(animation.trigger)
      ? animation.trigger
      : 'onClick',
    type: PPT_ELEMENT_ANIMATION_TYPES.includes(animation.type)
      ? animation.type
      : 'none',
  }
}

function getPPTObjectAnimationDescriptor(
  slide: PPTSlide,
  element: PPTElement,
): SlideEditObjectAnimationDescriptor<
  string,
  string,
  SlideEditBuiltInAnimationType,
  SlideEditBuiltInAnimationTrigger
> {
  const animation = getPPTElementAnimation(element, slide)

  return createSlideEditObjectAnimationDescriptor({
    delayMs: animation.delayMs,
    durationMs: animation.durationMs,
    objectId: element.id,
    order: animation.order,
    slideId: slide.id,
    trigger: toSlideEditObjectAnimationTrigger(animation.trigger),
    type: toSlideEditObjectAnimationType(animation.type),
  })
}

function getPPTSlideAnimationBuildOrder(
  slide: PPTSlide,
  options: { visibleOnly?: boolean } = {},
) {
  return getSlideEditObjectAnimationBuildOrder(
    slide.elements
      .filter((element) => !options.visibleOnly || element.visible !== false)
      .map((element) => getPPTObjectAnimationDescriptor(slide, element)),
  )
}

function toSlideEditObjectAnimationCommand({
  elementId,
  field,
  slideId,
  value,
}: {
  elementId: string
  field: PPTElementAnimationUpdateField
  slideId: string
  value: PPTElementAnimation[PPTElementAnimationUpdateField]
}): SlideEditObjectAnimationUpdateCommand<string, string> {
  if (field === 'type') {
    return {
      fieldId: 'type',
      id: 'update-object-animation',
      objectId: elementId,
      slideId,
      value: toSlideEditObjectAnimationType(value as PPTElementAnimationType),
    }
  }

  if (field === 'trigger') {
    return {
      fieldId: 'trigger',
      id: 'update-object-animation',
      objectId: elementId,
      slideId,
      value: toSlideEditObjectAnimationTrigger(value as PPTElementAnimationTrigger),
    }
  }

  switch (field) {
    case 'delayMs':
      return {
        fieldId: 'delayMs',
        id: 'update-object-animation',
        objectId: elementId,
        slideId,
        value: Number(value),
      }
    case 'durationMs':
      return {
        fieldId: 'durationMs',
        id: 'update-object-animation',
        objectId: elementId,
        slideId,
        value: Number(value),
      }
    case 'order':
      return {
        fieldId: 'order',
        id: 'update-object-animation',
        objectId: elementId,
        slideId,
        value: Number(value),
      }
  }
}

function toPPTElementAnimationUpdate(
  command: SlideEditObjectAnimationUpdateCommand<string, string>,
): {
  field: PPTElementAnimationUpdateField
  value: PPTElementAnimation[PPTElementAnimationUpdateField]
} {
  if (command.fieldId === 'type') {
    return {
      field: 'type',
      value: toPPTElementAnimationType(command.value),
    }
  }

  if (command.fieldId === 'trigger') {
    return {
      field: 'trigger',
      value: toPPTElementAnimationTrigger(command.value),
    }
  }

  switch (command.fieldId) {
    case 'delayMs':
      return {
        field: 'delayMs',
        value: normalizeSlideEditObjectAnimationDelayMs(command.value),
      }
    case 'durationMs':
      return {
        field: 'durationMs',
        value: normalizeSlideEditObjectAnimationDurationMs(command.value),
      }
    case 'order':
      return {
        field: 'order',
        value: normalizeSlideEditObjectAnimationOrder(command.value),
      }
  }
}

function toSlideEditObjectAnimationType(
  type: PPTElementAnimationType,
): SlideEditBuiltInAnimationType {
  switch (type) {
    case 'fadeIn':
      return 'fade-in'
    case 'flyIn':
      return 'fly-in'
    case 'none':
      return 'none'
  }
}

function toPPTElementAnimationType(type: string): PPTElementAnimationType {
  switch (type) {
    case 'fade-in':
      return 'fadeIn'
    case 'fly-in':
      return 'flyIn'
    default:
      return 'none'
  }
}

function toSlideEditObjectAnimationTrigger(
  trigger: PPTElementAnimationTrigger,
): SlideEditBuiltInAnimationTrigger {
  switch (trigger) {
    case 'onClick':
      return 'on-click'
    case 'withPrevious':
      return 'with-previous'
  }
}

function toPPTElementAnimationTrigger(trigger: string): PPTElementAnimationTrigger {
  return trigger === 'with-previous' ? 'withPrevious' : 'onClick'
}

function getPPTElementDefaultAnimationOrder(
  elementId: string,
  slide?: PPTSlide,
) {
  const index = slide?.elements.findIndex((element) => element.id === elementId) ?? -1

  return index >= 0 ? index + 1 : PPT_DEFAULT_ELEMENT_ANIMATION.order
}

function clampPPTElementAnimationTime(value: number) {
  return normalizeSlideEditObjectAnimationDurationMs(value)
}

function clampPPTElementAnimationOrder(
  value: number,
  elementCount: number = Number.MAX_SAFE_INTEGER,
  fallback: number = PPT_DEFAULT_ELEMENT_ANIMATION.order,
) {
  return clampPPTCanvasValue(
    Number.isFinite(value) ? Math.round(value) : fallback,
    1,
    Math.max(1, elementCount),
  )
}

function parsePPTElementAnimationTime(value: string) {
  return clampPPTElementAnimationTime(Number(value))
}

function parsePPTElementAnimationOrder(value: string, elementCount: number) {
  return clampPPTElementAnimationOrder(
    normalizeSlideEditObjectAnimationOrder(Number(value)),
    elementCount,
  )
}

function getDefaultPPTParagraphSpacing() {
  return {
    lineHeight: PPT_PARAGRAPH_LINE_HEIGHT_DEFAULT,
    spacingAfter: 0,
    spacingBefore: 0,
  }
}

function getPPTTextElementParagraphSpacing(element: PPTTextElement) {
  const paragraph = element.textBody.paragraphs[0]

  if (!paragraph) {
    return getDefaultPPTParagraphSpacing()
  }

  return {
    lineHeight: getPPTParagraphLineHeight(paragraph),
    spacingAfter: getPPTParagraphSpacingAfter(paragraph),
    spacingBefore: getPPTParagraphSpacingBefore(paragraph),
  }
}

function getPPTTextParagraphSpacingDescriptor(
  slideId: string,
  element: PPTTextElement,
): SlideEditTextParagraphSpacingDescriptor<string, string> {
  const spacing = getPPTTextElementParagraphSpacing(element)

  return createSlideEditTextParagraphSpacingDescriptor({
    lineHeightRatio: spacing.lineHeight,
    objectId: element.id,
    paragraphAfter: {
      unit: 'px',
      value: spacing.spacingAfter,
    },
    paragraphBefore: {
      unit: 'px',
      value: spacing.spacingBefore,
    },
    slideId,
  })
}

function getPPTTextParagraphSpacingField(
  descriptor: SlideEditTextParagraphSpacingDescriptor<string, string> | null,
  fieldId: SlideEditTextParagraphSpacingFieldId,
) {
  return descriptor?.fields.find((field) => field.id === fieldId)
}

function toSlideEditParagraphSpacingCommand({
  elementId,
  field,
  slideId,
  value,
}: {
  elementId: string
  field: PPTParagraphSpacingField
  slideId: string
  value: number
}): SlideEditTextParagraphSpacingUpdateCommand<string, string> {
  if (field === 'lineHeight') {
    return {
      fieldId: 'lineHeightRatio',
      id: 'update-text-paragraph-spacing',
      objectId: elementId,
      slideId,
      value,
    }
  }

  return {
    fieldId: field === 'spacingBefore' ? 'paragraphBefore' : 'paragraphAfter',
    id: 'update-text-paragraph-spacing',
    objectId: elementId,
    slideId,
    value: {
      unit: 'px',
      value,
    },
  }
}

function toPPTParagraphSpacingUpdate(
  command: SlideEditTextParagraphSpacingUpdateCommand<string, string>,
): {
  field: PPTParagraphSpacingField
  value: number
} {
  if (command.fieldId === 'lineHeightRatio') {
    return {
      field: 'lineHeight',
      value: normalizePPTParagraphLineHeight(command.value),
    }
  }

  return {
    field: command.fieldId === 'paragraphBefore' ? 'spacingBefore' : 'spacingAfter',
    value: normalizePPTParagraphSpacing(command.value.value),
  }
}

function normalizePPTParagraphLineHeight(value: number) {
  const finiteValue = Number.isFinite(value) ? value : PPT_PARAGRAPH_LINE_HEIGHT_DEFAULT

  return normalizeSlideEditTextLineHeightRatio(finiteValue)
}

function normalizePPTParagraphSpacing(value: number) {
  return normalizeSlideEditTextParagraphSpacingAmount({
    unit: 'px',
    value: Number.isFinite(value) ? value : 0,
  }).value
}

function parsePPTParagraphLineHeight(value: string) {
  return normalizePPTParagraphLineHeight(Number(value))
}

function parsePPTParagraphSpacing(value: string) {
  return normalizePPTParagraphSpacing(Number(value))
}

function getPPTParagraphLineHeight(paragraph: PPTParagraph) {
  return normalizePPTParagraphLineHeight(
    paragraph.lineHeight ?? PPT_PARAGRAPH_LINE_HEIGHT_DEFAULT,
  )
}

function getPPTParagraphSpacingAfter(paragraph: PPTParagraph) {
  return normalizePPTParagraphSpacing(paragraph.spacingAfter ?? 0)
}

function getPPTParagraphSpacingBefore(paragraph: PPTParagraph) {
  return normalizePPTParagraphSpacing(paragraph.spacingBefore ?? 0)
}

function getPPTParagraphStyle(paragraph: PPTParagraph): PPTParagraphCSSStyle {
  return getSlideEditTextParagraphSpacingCSSStyle({
    lineHeightRatio: getPPTParagraphLineHeight(paragraph),
    paragraphAfter: {
      unit: 'px',
      value: getPPTParagraphSpacingAfter(paragraph),
    },
    paragraphBefore: {
      unit: 'px',
      value: getPPTParagraphSpacingBefore(paragraph),
    },
  })
}

function getPPTElementAnimationStyle(
  animation: PPTElementAnimation,
): PPTElementAnimationCSSStyle {
  return getSlideEditObjectAnimationCSSStyle({
    delayMs: animation.delayMs,
    durationMs: animation.durationMs,
  })
}

function formatPPTElementAnimationType(type: PPTElementAnimationType) {
  switch (type) {
    case 'fadeIn':
      return 'Fade in'
    case 'flyIn':
      return 'Fly in'
    case 'none':
      return 'None'
  }
}

function formatPPTElementAnimationTrigger(trigger: PPTElementAnimationTrigger) {
  switch (trigger) {
    case 'onClick':
      return 'On click'
    case 'withPrevious':
      return 'With previous'
  }
}

function createPPTSlideMetadataInspectorDescriptor({
  slide,
  slideCount,
  slideIndex,
}: {
  slide: PPTSlide
  slideCount: number
  slideIndex: number
}): PPTSlideMetadataInspectorDescriptor {
  return {
    activeSlide: {
      index: slideIndex >= 0 ? slideIndex : null,
      slideCount,
      slideId: slide.id,
    },
    fields: PPT_SLIDE_METADATA_FIELDS,
    metadata: {
      background: getPPTSlideMetadataBackground(slide),
      name: slide.name,
      notes: slide.notes ?? '',
      orientation: getPPTSlideMetadataOrientation(),
      size: {
        h: PPT_SLIDE_HEIGHT,
        w: PPT_SLIDE_WIDTH,
      },
      slideId: slide.id,
    },
    surface: 'slide-metadata-inspector',
  }
}

function getPPTInspectorSurface({
  activeSlideId,
  selectedObjectIds,
}: {
  activeSlideId: string | null
  selectedObjectIds: readonly string[]
}): PPTInspectorSurfaceId {
  if (selectedObjectIds.length > 0) {
    return 'object-selection-inspector'
  }

  return activeSlideId ? 'slide-metadata-inspector' : 'none'
}

function getPPTSlideMetadataBackground(
  slide: PPTSlide,
): PPTSlideMetadataBackgroundDescriptor {
  if (!slide.background) {
    return {
      kind: 'none',
    }
  }

  return {
    color: slide.background.color,
    kind: 'solid-color',
  }
}

function getPPTSlideMetadataOrientation(): PPTSlideMetadataOrientation {
  return PPT_SLIDE_WIDTH >= PPT_SLIDE_HEIGHT ? 'landscape' : 'portrait'
}

function getPPTSlideMetadataField(
  descriptor: PPTSlideMetadataInspectorDescriptor,
  fieldId: PPTSlideMetadataFieldId,
): PPTSlideMetadataFieldDescriptor {
  const field = descriptor.fields.find((field) => field.id === fieldId)

  if (!field) {
    throw new Error(`Missing PPT slide metadata field: ${fieldId}`)
  }

  return field
}

function getPPTSlideMetadataFieldData(field: PPTSlideMetadataFieldDescriptor) {
  return {
    'data-ppt-slide-metadata-adapter-slot': field.requiredAdapterSlot,
    'data-ppt-slide-metadata-command': field.commandId,
    'data-ppt-slide-metadata-control': field.control,
    'data-ppt-slide-metadata-editable': String(field.isEditable),
    'data-ppt-slide-metadata-field': field.id,
    'data-ppt-slide-metadata-optional': String(field.isOptional),
  }
}

function toPPTSlideMetadataHostCommandEffect(
  command: PPTSlideMetadataUpdateCommand,
): PPTSlideMetadataHostCommandEffect {
  return {
    payload: command,
    selection: {
      objectIds: [],
      slideId: command.slideId,
    },
    type: 'slide-command-effect',
  }
}

function applyPPTSlideMetadataHostCommandEffect(
  slide: PPTSlide,
  effect: PPTSlideMetadataHostCommandEffect,
): PPTSlide {
  const command = effect.payload

  switch (command.fieldId) {
    case 'background':
      return {
        ...slide,
        background: command.value.kind === 'solid-color'
          ? { color: command.value.color }
          : { color: '#ffffff' },
      }
    case 'name':
      return {
        ...slide,
        name: command.value,
      }
    case 'notes':
      return {
        ...slide,
        notes: command.value,
      }
    case 'orientation':
    case 'size':
      return slide
  }
}

function toPPTLayoutPlaceholderVisibilityHostCommandEffect(
  command: PPTLayoutPlaceholderVisibilityCommand,
  selectedObjectIds: readonly string[],
): PPTLayoutPlaceholderVisibilityHostCommandEffect {
  return {
    payload: command,
    selection: {
      objectIds: selectedObjectIds,
      slideId: command.slideId,
    },
    type: 'slide-command-effect',
  }
}

function applyPPTLayoutPlaceholderVisibilityHostCommandEffect(
  slide: PPTSlide,
  effect: PPTLayoutPlaceholderVisibilityHostCommandEffect,
): PPTSlide {
  const hiddenPlaceholderIds = new Set(slide.hiddenPlaceholderIds ?? [])

  if (effect.payload.isVisible) {
    hiddenPlaceholderIds.delete(effect.payload.placeholderId)
  } else {
    hiddenPlaceholderIds.add(effect.payload.placeholderId)
  }

  return {
    ...slide,
    hiddenPlaceholderIds: [...hiddenPlaceholderIds],
  }
}

function createPPTClipboardPayload({
  objects,
  operation = 'copy',
  selectedObjectIds,
  sourceSlideId,
}: {
  objects: readonly PPTElement[]
  operation?: PPTClipboardOperation
  selectedObjectIds: readonly string[]
  sourceSlideId: string
}): PPTClipboardPayload {
  return createSlideEditClipboardPayload({
    metadata: objects.map((object): PPTClipboardObjectMetadata => ({
      groupId: object.groupId ?? null,
      objectId: object.id,
      placeholderId: null,
    })),
    objects,
    operation,
    selectedObjectIds,
    sourceSlideId,
  })
}

function createPPTRichClipboardEffect(
  payload: PPTClipboardPayload,
  options: {
    fallback: PPTRichClipboardFallback
    importFormat?: PPTRichClipboardImportFormat
    imported?: boolean
    writeMode?: PPTRichClipboardWriteMode
  },
): PPTRichClipboardEffect {
  return {
    formats: PPT_RICH_CLIPBOARD_FORMATS,
    htmlLength: options.fallback.html.length,
    importFormat: options.importFormat,
    imported: options.imported,
    model: PPT_RICH_CLIPBOARD_MODEL,
    objectCount: payload.objects.length,
    plainTextLength: options.fallback.plainText.length,
    selectedObjectIds: payload.selectedObjectIds,
    sourceSlideId: payload.sourceSlideId,
    writeMode: options.writeMode,
  }
}

function createPPTSlideClipboardPayload(slide: PPTSlide): PPTSlideClipboardPayload {
  return {
    slide,
    sourceSlideId: slide.id,
  }
}

function createPPTSlideClipboardEffect({
  html,
  importFormat,
  imported,
  payload,
  targetSlideId,
  writeMode,
}: {
  html?: string
  importFormat?: PPTRichClipboardImportFormat
  imported?: boolean
  payload: PPTSlideClipboardPayload
  targetSlideId?: string
  writeMode?: PPTRichClipboardWriteMode
}): PPTSlideClipboardEffect {
  return {
    elementCount: payload.slide.elements.length,
    htmlLength: html?.length ?? createPPTSlideClipboardHTML(payload).length,
    imported,
    importFormat,
    jsonMimeType: PPT_SLIDE_CLIPBOARD_JSON_MIME_TYPE,
    model: PPT_SLIDE_CLIPBOARD_MODEL,
    slideName: payload.slide.name,
    sourceSlideId: payload.sourceSlideId,
    targetSlideId,
    writeMode,
  }
}

function createPPTSlideClipboardFallbackHTMLEffect({
  source,
  targetSlideId,
}: {
  source: PPTSlideClipboardFallbackHTMLSource
  targetSlideId: string
}): PPTSlideClipboardEffect {
  return {
    elementCount: 1,
    htmlLength: source.htmlLength,
    importFormat: 'text-html-fallback',
    imported: true,
    jsonMimeType: PPT_SLIDE_CLIPBOARD_JSON_MIME_TYPE,
    model: PPT_SLIDE_CLIPBOARD_MODEL,
    slideName: source.name,
    sourceSlideId: source.sourceSlideId ?? '',
    targetSlideId,
  }
}

function createPPTHTMLClipboardEffect({
  html,
  sourceSlideId,
  writeMode,
}: {
  html: string
  sourceSlideId: string
  writeMode?: PPTRichClipboardWriteMode
}): PPTHTMLClipboardEffect {
  return {
    htmlLength: html.length,
    jsonMimeType: PPT_HTML_CLIPBOARD_JSON_MIME_TYPE,
    model: PPT_HTML_CLIPBOARD_MODEL,
    sourceSlideId,
    writeMode,
  }
}

function createPPTDeckHTMLImportEffect({
  importedSlides,
  source,
}: {
  importedSlides: readonly PPTSlide[]
  source: PPTDeckHTMLImportSource
}): PPTDeckHTMLImportEffect {
  return {
    firstImportedSlideId: importedSlides[0]?.id ?? '',
    format: PPT_DECK_HTML_IMPORT_FORMAT,
    htmlLength: source.htmlLength,
    importedSlideCount: importedSlides.length,
    model: PPT_DECK_HTML_IMPORT_MODEL,
    sourceDeckId: source.deck.id,
    sourceSlideCount: source.deck.slides.length,
    sourceTitle: source.deck.title,
  }
}

function createPPTDeckHTMLFallbackImportEffect({
  importedSlides,
  source,
}: {
  importedSlides: readonly PPTSlide[]
  source: PPTDeckHTMLFallbackSource
}): PPTDeckHTMLImportEffect {
  return {
    firstImportedSlideId: importedSlides[0]?.id ?? '',
    format: PPT_DECK_HTML_FALLBACK_IMPORT_FORMAT,
    htmlLength: source.htmlLength,
    importedSlideCount: importedSlides.length,
    model: PPT_DECK_HTML_IMPORT_MODEL,
    sourceDeckId: '',
    sourceSlideCount: source.slides.length,
    sourceTitle: source.title,
  }
}

function createPPTDeckJSONImportEffect({
  importedSlides,
  source,
}: {
  importedSlides: readonly PPTSlide[]
  source: PPTDeckJSONImportSource
}): PPTDeckJSONImportEffect {
  return {
    firstImportedSlideId: importedSlides[0]?.id ?? '',
    format: source.format,
    importedSlideCount: importedSlides.length,
    jsonLength: source.jsonLength,
    model: PPT_DECK_JSON_IMPORT_MODEL,
    sourceDeckId: source.deck.id,
    sourceSlideCount: source.deck.slides.length,
    sourceTitle: source.deck.title,
  }
}

function createPPTSlideJSONImportEffect({
  importedSlides,
  source,
}: {
  importedSlides: readonly PPTSlide[]
  source: PPTSlideJSONImportSource
}): PPTSlideJSONImportEffect {
  return {
    firstImportedSlideId: importedSlides[0]?.id ?? '',
    format: source.format,
    importedSlideCount: importedSlides.length,
    jsonLength: source.jsonLength,
    model: PPT_SLIDE_JSON_IMPORT_MODEL,
    sourceSlideCount: source.slides.length,
    sourceSlideIds: source.slides.map((slide) => slide.id).join(' '),
    sourceSlideNames: source.slides.map((slide) => slide.name).join(', '),
  }
}

function createPPTSlideNotesImportEffect({
  effect,
  source,
}: {
  effect: PPTSlideMetadataHostCommandEffect
  source: PPTSlideNotesImportSource
}): PPTSlideNotesImportEffect {
  return {
    format: source.format,
    model: PPT_SLIDE_NOTES_IMPORT_MODEL,
    notesLength: source.notes.length,
    slideId: effect.selection.slideId,
    textLength: source.textLength,
  }
}

function createPPTSlideMetadataImportEffect({
  effects,
  source,
}: {
  effects: readonly PPTSlideMetadataHostCommandEffect[]
  source: PPTSlideMetadataImportSource
}): PPTSlideMetadataImportEffect {
  return {
    backgroundColor: source.background?.kind === 'solid-color'
      ? source.background.color
      : '',
    commandIds: effects.map((effect) => effect.payload.id).join(' '),
    fieldIds: effects.map((effect) => effect.payload.fieldId).join(' '),
    format: source.format,
    jsonLength: source.jsonLength,
    model: PPT_SLIDE_METADATA_IMPORT_MODEL,
    name: source.name ?? '',
    notesLength: source.notes?.length ?? 0,
    slideId: effects[0]?.selection.slideId ?? '',
  }
}

function createPPTSlideMetadataImportCommandEffects({
  slideId,
  source,
}: {
  slideId: string
  source: PPTSlideMetadataImportSource
}): PPTSlideMetadataHostCommandEffect[] {
  const effects: PPTSlideMetadataHostCommandEffect[] = []

  if (source.name !== undefined) {
    effects.push(toPPTSlideMetadataHostCommandEffect({
      fieldId: 'name',
      id: 'update-slide-name',
      slideId,
      value: source.name,
    }))
  }

  if (source.background !== undefined) {
    effects.push(toPPTSlideMetadataHostCommandEffect({
      fieldId: 'background',
      id: 'update-slide-background',
      slideId,
      value: source.background,
    }))
  }

  if (source.notes !== undefined) {
    effects.push(toPPTSlideMetadataHostCommandEffect({
      fieldId: 'notes',
      id: 'update-slide-notes',
      slideId,
      value: source.notes,
    }))
  }

  return effects
}

function createPPTSlideLayoutImportEffect({
  layoutEffect,
  placeholderEffects,
  source,
}: {
  layoutEffect: PPTSlideLayoutApplyHostCommandEffect | null
  placeholderEffects: readonly PPTLayoutPlaceholderVisibilityHostCommandEffect[]
  source: PPTSlideLayoutImportSource
}): PPTSlideLayoutImportEffect {
  const effects = [
    ...(layoutEffect ? [layoutEffect] : []),
    ...placeholderEffects,
  ]

  return {
    commandFields: [
      ...(layoutEffect ? ['layoutId'] : []),
      ...placeholderEffects.map((effect) => effect.payload.placeholderId),
    ].join(' '),
    commandIds: effects.map((effect) => effect.payload.id).join(' '),
    commandTypes: effects.map((effect) => effect.type).join(' '),
    fields: source.fields.join(' '),
    format: source.format,
    hiddenPlaceholderIds: source.hiddenPlaceholderIds?.join(' ') ?? '',
    jsonLength: source.jsonLength,
    layoutId: source.layoutId ?? '',
    model: PPT_SLIDE_LAYOUT_IMPORT_MODEL,
    placeholderCommandCount: placeholderEffects.length,
    slideId: effects[0]?.selection.slideId ?? '',
    themeId: source.themeId ?? '',
  }
}

function createPPTSlideLayoutPlaceholderImportCommandEffects({
  hiddenPlaceholderIds,
  layout,
  selectedObjectIds,
  slideId,
}: {
  hiddenPlaceholderIds: readonly string[]
  layout: SlideEditLayoutDescriptor
  selectedObjectIds: readonly string[]
  slideId: string
}): PPTLayoutPlaceholderVisibilityHostCommandEffect[] {
  const hiddenPlaceholderIdSet = new Set(
    normalizePPTSlideLayoutHiddenPlaceholderIds(hiddenPlaceholderIds, layout),
  )

  return layout.placeholders.map((placeholder) =>
    toPPTLayoutPlaceholderVisibilityHostCommandEffect({
      id: 'update-placeholder-visibility',
      isVisible: !hiddenPlaceholderIdSet.has(placeholder.placeholderId),
      placeholderId: placeholder.placeholderId,
      slideId,
    }, selectedObjectIds))
}

function normalizePPTSlideLayoutHiddenPlaceholderIds(
  placeholderIds: readonly string[],
  layout: SlideEditLayoutDescriptor,
) {
  const hiddenPlaceholderIds = new Set(placeholderIds)

  return layout.placeholders
    .map((placeholder) => placeholder.placeholderId)
    .filter((placeholderId) => hiddenPlaceholderIds.has(placeholderId))
}

function createPPTSlideTransitionImportEffect({
  effects,
  source,
}: {
  effects: readonly PPTSlideTransitionHostCommandEffect[]
  source: PPTSlideTransitionImportSource
}): PPTSlideTransitionImportEffect {
  return {
    advanceAfterMs: source.transition.advanceAfterMs === undefined
      ? ''
      : String(source.transition.advanceAfterMs ?? ''),
    advanceOnClick: source.transition.advanceOnClick === undefined
      ? ''
      : String(source.transition.advanceOnClick),
    commandFields: effects.map((effect) => effect.payload.fieldId).join(' '),
    commandIds: effects.map((effect) => effect.payload.id).join(' '),
    durationMs: source.transition.durationMs === undefined
      ? ''
      : String(source.transition.durationMs),
    fields: source.fields.join(' '),
    format: source.format,
    jsonLength: source.jsonLength,
    model: PPT_SLIDE_TRANSITION_IMPORT_MODEL,
    slideId: effects[0]?.selection.slideId ?? '',
    type: source.transition.type ?? '',
  }
}

function createPPTSlideTransitionImportCommandEffects({
  currentTransition,
  slideId,
  source,
}: {
  currentTransition: PPTSlideTransition
  slideId: string
  source: PPTSlideTransitionImportSource
}): PPTSlideTransitionHostCommandEffect[] {
  let transition = currentTransition
  const effects: PPTSlideTransitionHostCommandEffect[] = []

  for (const field of source.fields) {
    const value = source.transition[field]

    if (value === undefined) {
      continue
    }

    const effect = getSlideEditTransitionUpdateCommandEffect(
      getPPTSlideTransitionUpdateCommand({
        field,
        slideId,
        transition,
        value,
      }),
    )

    effects.push(effect)
    transition = applyPPTSlideTransitionUpdateCommand(
      transition,
      effect.payload,
    )
  }

  return effects
}

function createPPTObjectAnimationImportEffect({
  effects,
  source,
}: {
  effects: readonly SlideEditObjectAnimationHostCommandEffect<string, string>[]
  source: PPTObjectAnimationImportSource
}): PPTObjectAnimationImportEffect {
  return {
    commandFields: effects.map((effect) => effect.payload.fieldId).join(' '),
    commandIds: effects.map((effect) => effect.payload.id).join(' '),
    delayMs: source.animation.delayMs === undefined
      ? ''
      : String(source.animation.delayMs),
    durationMs: source.animation.durationMs === undefined
      ? ''
      : String(source.animation.durationMs),
    fields: source.fields.join(' '),
    format: source.format,
    jsonLength: source.jsonLength,
    model: PPT_OBJECT_ANIMATION_IMPORT_MODEL,
    objectIds: [...new Set(effects.map((effect) => effect.payload.objectId))]
      .join(' '),
    order: source.animation.order === undefined
      ? ''
      : String(source.animation.order),
    slideId: effects[0]?.payload.slideId ?? '',
    trigger: source.animation.trigger ?? '',
    type: source.animation.type ?? '',
  }
}

function createPPTObjectAnimationImportCommandEffects({
  objectIds,
  slideId,
  source,
}: {
  objectIds: readonly string[]
  slideId: string
  source: PPTObjectAnimationImportSource
}): SlideEditObjectAnimationHostCommandEffect<string, string>[] {
  const effects: SlideEditObjectAnimationHostCommandEffect<string, string>[] = []

  for (const objectId of objectIds) {
    for (const field of source.fields) {
      const value = source.animation[field]

      if (value === undefined) {
        continue
      }

      effects.push(getSlideEditObjectAnimationUpdateCommandEffect(
        toSlideEditObjectAnimationCommand({
          elementId: objectId,
          field,
          slideId,
          value,
        }),
      ))
    }
  }

  return effects
}

function createPPTObjectStyleImportEffect({
  effect,
  source,
}: {
  effect: PPTStyleClipboardHostCommandEffect
  source: PPTObjectStyleImportSource
}): PPTObjectStyleImportEffect {
  const payload = effect.payload.id === 'paste-object-formatting'
    ? effect.payload
    : null
  const categories = payload
    ? uniquePPTCanvasValues(payload.categoryApplications.flatMap(
        (application) => application.appliedCategoryIds,
      )).join(' ')
    : ''
  const shadow = source.object.shadow

  return {
    categories,
    commandId: effect.payload.id,
    commandTargets: payload?.targetObjectIds.join(' ') ?? '',
    commandType: effect.type,
    fields: source.fields.join(' '),
    format: source.format,
    jsonLength: source.jsonLength,
    model: PPT_OBJECT_STYLE_IMPORT_MODEL,
    objectIds: payload?.categoryApplications
      .map((application) => application.objectId)
      .join(' ') ?? '',
    opacity: source.object.opacity === undefined
      ? ''
      : String(source.object.opacity),
    shadowAngle: shadow ? String(shadow.angle) : '',
    shadowBlur: shadow ? String(shadow.blur) : '',
    shadowColor: shadow?.color ?? '',
    shadowDistance: shadow ? String(shadow.distance) : '',
    shadowEnabled: shadow === undefined ? '' : String(shadow !== null),
    shadowOpacity: shadow ? String(shadow.opacity) : '',
  }
}

function createPPTObjectMetadataImportEffect({
  accessibilityEffects,
  hyperlinkEffects,
  renameEffects,
  source,
}: {
  accessibilityEffects: readonly SlideEditObjectAccessibilityHostCommandEffect<string, string>[]
  hyperlinkEffects: readonly SlideEditObjectHyperlinkHostCommandEffect<string, string>[]
  renameEffects: readonly PPTLayerPaneHostCommandEffect[]
  source: PPTObjectMetadataImportSource
}): PPTObjectMetadataImportEffect {
  const effects = [
    ...hyperlinkEffects,
    ...accessibilityEffects,
    ...renameEffects,
  ]
  const altText = source.metadata.altText

  return {
    altTextLength: typeof altText === 'string' ? altText.length : 0,
    altTextPresent: altText === undefined
      ? ''
      : String(typeof altText === 'string' && altText.length > 0),
    commandFields: effects.map((effect) =>
      getPPTObjectMetadataCommandField(effect)).join(' '),
    commandIds: effects.map((effect) => effect.payload.id).join(' '),
    commandTypes: effects.map((effect) => effect.type).join(' '),
    fields: source.fields.join(' '),
    format: source.format,
    hyperlinkUrl: source.metadata.hyperlinkUrl ?? '',
    jsonLength: source.jsonLength,
    model: PPT_OBJECT_METADATA_IMPORT_MODEL,
    name: source.metadata.name ?? '',
    objectIds: [...new Set(effects.flatMap((effect) =>
      getPPTObjectMetadataEffectObjectIds(effect)))]
      .join(' '),
    slideId: getPPTObjectMetadataEffectSlideId(effects[0]) ?? '',
  }
}

function createPPTObjectTransformImportEffect({
  objectIds,
  source,
}: {
  objectIds: readonly string[]
  source: PPTObjectTransformImportSource
}): PPTObjectTransformImportEffect {
  return {
    commandTargets: objectIds.join(' '),
    fields: source.fields.join(' '),
    format: source.format,
    h: source.transform.h === undefined ? '' : String(source.transform.h),
    jsonLength: source.jsonLength,
    model: PPT_OBJECT_TRANSFORM_IMPORT_MODEL,
    objectIds: objectIds.join(' '),
    rotation: source.transform.rotation === undefined
      ? ''
      : String(source.transform.rotation),
    w: source.transform.w === undefined ? '' : String(source.transform.w),
    x: source.transform.x === undefined ? '' : String(source.transform.x),
    y: source.transform.y === undefined ? '' : String(source.transform.y),
  }
}

function applyPPTObjectTransformSourceToElement(
  element: PPTElement,
  source: PPTObjectTransformImportSource,
): PPTElement {
  const bounds = clampPPTCanvasBoundsToFrame({
    bounds: {
      ...pptGeometryToBounds(element.geometry),
      ...(source.transform.h === undefined ? {} : { h: source.transform.h }),
      ...(source.transform.w === undefined ? {} : { w: source.transform.w }),
      ...(source.transform.x === undefined ? {} : { x: source.transform.x }),
      ...(source.transform.y === undefined ? {} : { y: source.transform.y }),
    },
    frame: {
      h: PPT_SLIDE_HEIGHT,
      w: PPT_SLIDE_WIDTH,
      x: 0,
      y: 0,
    },
    minHeight: 24,
    minWidth: 24,
  })
  const next = updatePPTElementBounds(element, bounds)

  return source.transform.rotation === undefined
    ? next
    : {
        ...next,
        geometry: {
          ...next.geometry,
          rotation: normalizePPTCanvasRotationDegrees(
            source.transform.rotation,
          ),
        },
      }
}

function getPPTObjectMetadataCommandField(
  effect:
    | SlideEditObjectAccessibilityHostCommandEffect<string, string>
    | SlideEditObjectHyperlinkHostCommandEffect<string, string>
    | PPTLayerPaneHostCommandEffect,
) {
  switch (effect.payload.id) {
    case 'remove-object-alt-text':
    case 'remove-object-hyperlink':
      return ''
    case 'rename-object':
      return 'name'
    case 'update-object-accessibility':
    case 'update-object-hyperlink':
      return effect.payload.fieldId
    default:
      return ''
  }
}

function getPPTObjectMetadataEffectObjectIds(
  effect:
    | SlideEditObjectAccessibilityHostCommandEffect<string, string>
    | SlideEditObjectHyperlinkHostCommandEffect<string, string>
    | PPTLayerPaneHostCommandEffect,
) {
  switch (effect.payload.id) {
    case 'hide-objects':
    case 'lock-objects':
    case 'select-objects':
    case 'show-objects':
    case 'unlock-objects':
      return effect.payload.objectIds
    case 'rename-object':
    case 'reorder-object':
    case 'remove-object-alt-text':
    case 'remove-object-hyperlink':
    case 'update-object-accessibility':
    case 'update-object-hyperlink':
      return [effect.payload.objectId]
  }
}

function getPPTObjectMetadataEffectSlideId(
  effect:
    | SlideEditObjectAccessibilityHostCommandEffect<string, string>
    | SlideEditObjectHyperlinkHostCommandEffect<string, string>
    | PPTLayerPaneHostCommandEffect
    | undefined,
) {
  if (!effect) {
    return undefined
  }

  return 'slideId' in effect.payload
    ? effect.payload.slideId
    : effect.selection.slideId
}

function createPPTObjectMetadataHyperlinkEffects({
  objectIds,
  slideId,
  source,
}: {
  objectIds: readonly string[]
  slideId: string
  source: PPTObjectMetadataImportSource
}): SlideEditObjectHyperlinkHostCommandEffect<string, string>[] {
  if (source.metadata.hyperlinkUrl === undefined) {
    return []
  }

  return objectIds.map((objectId) =>
    getSlideEditObjectHyperlinkCommandEffect(
      source.metadata.hyperlinkUrl
        ? {
            fieldId: 'url',
            id: 'update-object-hyperlink',
            objectId,
            slideId,
            value: source.metadata.hyperlinkUrl,
          }
        : {
            id: 'remove-object-hyperlink',
            objectId,
            slideId,
          },
    ))
}

function createPPTObjectMetadataAccessibilityEffects({
  objectIds,
  slideId,
  source,
}: {
  objectIds: readonly string[]
  slideId: string
  source: PPTObjectMetadataImportSource
}): SlideEditObjectAccessibilityHostCommandEffect<string, string>[] {
  if (source.metadata.altText === undefined) {
    return []
  }

  return objectIds.map((objectId) =>
    getSlideEditObjectAccessibilityCommandEffect(
      source.metadata.altText
        ? {
            fieldId: 'altText',
            id: 'update-object-accessibility',
            objectId,
            slideId,
            value: source.metadata.altText,
          }
        : {
            id: 'remove-object-alt-text',
            objectId,
            slideId,
          },
    ))
}

function createPPTObjectMetadataRenameEffects({
  objectIds,
  slide,
  source,
}: {
  objectIds: readonly string[]
  slide: PPTSlide
  source: PPTObjectMetadataImportSource
}): PPTLayerPaneHostCommandEffect[] {
  if (source.metadata.name === undefined) {
    return []
  }

  const descriptor = createPPTLayerPaneDescriptor({
    activeObjectId: null,
    collapsedGroupIds: new Set(),
    selectedObjectIds: objectIds,
    slide,
  })

  return objectIds
    .map((objectId) =>
      getSlideEditLayerPaneCommandEffect(descriptor, {
        name: source.metadata.name ?? '',
        objectId,
        type: 'rename-submit',
      }))
    .filter((effect): effect is PPTLayerPaneHostCommandEffect =>
      effect !== null)
}

function applyPPTObjectHyperlinkCommandEffectToElement(
  element: PPTElement,
  effect: SlideEditObjectHyperlinkHostCommandEffect<string, string>,
): PPTElement {
  const hyperlink = effect.payload.id === 'update-object-hyperlink'
    ? normalizePPTElementHyperlink({ url: effect.payload.value ?? '' })
    : null

  return {
    ...element,
    hyperlink: hyperlink ?? undefined,
  }
}

function applyPPTObjectMetadataRenameCommandEffectToElement(
  element: PPTElement,
  effect: PPTLayerPaneHostCommandEffect,
): PPTElement {
  return effect.payload.id === 'rename-object'
    ? {
        ...element,
        name: effect.payload.name,
      }
    : element
}

function applyPPTObjectAccessibilityCommandEffectToElement(
  element: PPTElement,
  effect: SlideEditObjectAccessibilityHostCommandEffect<string, string>,
): PPTElement {
  const accessibility = effect.payload.id === 'update-object-accessibility'
    ? normalizePPTElementAccessibility({
        altText: typeof effect.payload.value === 'string'
          ? effect.payload.value
          : '',
      })
    : null

  return {
    ...element,
    accessibility: accessibility ?? undefined,
  }
}

function createPPTObjectStateImportEffect({
  lockEffects,
  objectIds,
  source,
  visibilityEffect,
}: {
  lockEffects: readonly PPTLayerPaneHostCommandEffect[]
  objectIds: readonly string[]
  source: PPTObjectStateImportSource
  visibilityEffect: PPTObjectVisibilityHostCommandEffect | null
}): PPTObjectStateImportEffect {
  const visibilityEffects = visibilityEffect ? [visibilityEffect] : []
  const effects = source.state.locked === false
    ? [...lockEffects, ...visibilityEffects]
    : [...visibilityEffects, ...lockEffects]
  const lockTargets = uniquePPTCanvasValues(lockEffects.flatMap((effect) =>
    effect.payload.id === 'lock-objects' ||
      effect.payload.id === 'unlock-objects'
      ? effect.payload.objectIds
      : []))

  return {
    commandIds: effects.map((effect) => effect.payload.id).join(' '),
    commandTypes: effects.map((effect) => effect.type).join(' '),
    fields: source.fields.join(' '),
    format: source.format,
    jsonLength: source.jsonLength,
    locked: source.state.locked === undefined
      ? ''
      : String(source.state.locked),
    lockTargets: lockTargets.join(' '),
    model: PPT_OBJECT_STATE_IMPORT_MODEL,
    objectIds: objectIds.join(' '),
    slideId: effects[0]?.selection.slideId ?? '',
    visible: source.state.visible === undefined
      ? ''
      : String(source.state.visible),
    visibilityTargets: visibilityEffect?.payload.objectIds.join(' ') ?? '',
  }
}

function createPPTObjectStateImportCommandEffects({
  objectIds,
  slide,
  source,
}: {
  objectIds: readonly string[]
  slide: PPTSlide
  source: PPTObjectStateImportSource
}): {
  lockEffects: PPTLayerPaneHostCommandEffect[]
  visibilityEffect: PPTObjectVisibilityHostCommandEffect | null
} {
  const lockEffects = source.state.locked === undefined
    ? []
    : createPPTObjectStateLockCommandEffects({
      locked: source.state.locked,
      objectIds,
      slide,
    })
  const visibilitySlide = source.state.locked === false
    ? applyPPTObjectStateVirtualLock(slide, objectIds, false)
    : slide
  const visibilityEffect = source.state.visible === undefined
    ? null
    : createPPTObjectStateVisibilityCommandEffect({
      objectIds,
      slide: visibilitySlide,
      visible: source.state.visible,
    })

  return {
    lockEffects,
    visibilityEffect,
  }
}

function createPPTObjectStateVisibilityCommandEffect({
  objectIds,
  slide,
  visible,
}: {
  objectIds: readonly string[]
  slide: PPTSlide
  visible: boolean
}): PPTObjectVisibilityHostCommandEffect | null {
  const descriptor = createPPTLayerPaneDescriptor({
    activeObjectId: null,
    collapsedGroupIds: new Set(),
    selectedObjectIds: objectIds,
    slide,
  })

  return getSlideEditObjectVisibilityCommandEffect({
    commandId: visible ? 'show-objects' : 'hide-objects',
    objects: getPPTObjectVisibilityDescriptors(slide.id, descriptor.rows),
    selectedObjectIds: objectIds,
    slideId: slide.id,
  })
}

function createPPTObjectStateLockCommandEffects({
  locked,
  objectIds,
  slide,
}: {
  locked: boolean
  objectIds: readonly string[]
  slide: PPTSlide
}): PPTLayerPaneHostCommandEffect[] {
  const descriptor = createPPTLayerPaneDescriptor({
    activeObjectId: null,
    collapsedGroupIds: new Set(),
    selectedObjectIds: objectIds,
    slide,
  })

  return objectIds
    .map((objectId) => {
      const row = descriptor.rows.find((candidate) =>
        candidate.objectId === objectId)

      if (!row || row.isLocked === locked) {
        return null
      }

      return getSlideEditLayerPaneCommandEffect(descriptor, {
        objectId,
        type: 'lock-toggle',
      })
    })
    .filter((effect): effect is PPTLayerPaneHostCommandEffect =>
      effect !== null)
}

function applyPPTObjectStateVirtualLock(
  slide: PPTSlide,
  objectIds: readonly string[],
  locked: boolean,
): PPTSlide {
  const objectIdSet = new Set(objectIds)

  return {
    ...slide,
    elements: slide.elements.map((element) =>
      objectIdSet.has(element.id)
        ? { ...element, locked }
        : element),
  }
}

function createPPTObjectLayerImportEffect({
  effect,
  source,
}: {
  effect: PPTLayerPaneHostCommandEffect
  source: PPTObjectLayerImportSource
}): PPTObjectLayerImportEffect {
  const payload = effect.payload

  return {
    commandId: payload.id,
    commandType: effect.type,
    fields: source.fields.join(' '),
    format: source.format,
    fromIndex: payload.id === 'reorder-object' ? payload.fromIndex : -1,
    jsonLength: source.jsonLength,
    model: PPT_OBJECT_LAYER_IMPORT_MODEL,
    objectId: payload.id === 'reorder-object' ? payload.objectId : '',
    position: source.layer.position ?? '',
    slideId: effect.selection.slideId,
    toIndex: payload.id === 'reorder-object' ? payload.toIndex : -1,
  }
}

function createPPTObjectLayerImportCommandEffect({
  objectId,
  slide,
  source,
}: {
  objectId: string
  slide: PPTSlide
  source: PPTObjectLayerImportSource
}): PPTLayerPaneHostCommandEffect | null {
  const toIndex = getPPTObjectLayerImportTargetIndex({
    objectId,
    slide,
    source,
  })

  if (toIndex === null) {
    return null
  }

  const descriptor = createPPTLayerPaneDescriptor({
    activeObjectId: objectId,
    collapsedGroupIds: new Set(),
    selectedObjectIds: [objectId],
    slide,
  })

  return getSlideEditLayerPaneCommandEffect(descriptor, {
    objectId,
    toIndex,
    type: 'row-drop',
  })
}

function getPPTObjectLayerImportTargetIndex({
  objectId,
  slide,
  source,
}: {
  objectId: string
  slide: PPTSlide
  source: PPTObjectLayerImportSource
}) {
  const currentIndex = slide.elements.findIndex((element) =>
    element.id === objectId)

  if (currentIndex < 0) {
    return null
  }

  if (source.layer.toIndex !== undefined) {
    return clampPPTObjectLayerImportIndex(source.layer.toIndex, slide)
  }

  switch (source.layer.position) {
    case 'back':
      return 0
    case 'backward':
      return clampPPTObjectLayerImportIndex(currentIndex - 1, slide)
    case 'forward':
      return clampPPTObjectLayerImportIndex(currentIndex + 2, slide)
    case 'front':
      return slide.elements.length
    default:
      return null
  }
}

function clampPPTObjectLayerImportIndex(value: number, slide: PPTSlide) {
  return Math.max(0, Math.min(slide.elements.length, Math.trunc(value)))
}

function createPPTImageCropImportEffect({
  effects,
  source,
}: {
  effects: readonly SlideEditObjectImageCropHostCommandEffect<string, string>[]
  source: PPTImageCropImportSource
}): PPTImageCropImportEffect {
  return {
    commandFields: effects.map((effect) =>
      effect.payload.id === 'update-object-image-crop'
        ? effect.payload.fieldId
        : '',
    ).join(' '),
    commandIds: effects.map((effect) => effect.payload.id).join(' '),
    commandTypes: effects.map((effect) => effect.type).join(' '),
    fields: source.fields.join(' '),
    fit: source.imageCrop.fit ?? '',
    format: source.format,
    jsonLength: source.jsonLength,
    model: PPT_IMAGE_CROP_IMPORT_MODEL,
    objectIds: [...new Set(effects.map((effect) => effect.payload.objectId))]
      .join(' '),
    slideId: effects[0]?.payload.slideId ?? '',
    x: source.imageCrop.crop?.x === undefined
      ? ''
      : String(source.imageCrop.crop.x),
    y: source.imageCrop.crop?.y === undefined
      ? ''
      : String(source.imageCrop.crop.y),
  }
}

function createPPTImageReplaceImportEffect({
  effect,
  source,
}: {
  effect: SlideEditObjectImageReplaceHostCommandEffect<string, string>
  source: PPTImageReplaceImportSource
}): PPTImageReplaceImportEffect {
  const image = source.image
  const altText = image.altText ?? ''

  return {
    altTextLength: altText.length,
    commandId: effect.payload.id,
    commandType: effect.type,
    fields: source.fields.join(' '),
    format: source.format,
    jsonLength: source.jsonLength,
    mimeType: effect.payload.source.mimeType,
    model: PPT_IMAGE_REPLACE_IMPORT_MODEL,
    name: effect.payload.source.name ?? '',
    naturalHeight: effect.payload.source.naturalHeight === undefined
      ? ''
      : String(effect.payload.source.naturalHeight),
    naturalWidth: effect.payload.source.naturalWidth === undefined
      ? ''
      : String(effect.payload.source.naturalWidth),
    objectId: effect.payload.objectId,
    slideId: effect.payload.slideId,
    srcPrefix: effect.payload.source.src.slice(0, 19),
  }
}

function createPPTImageCropImportCommandEffects({
  objectIds,
  slideId,
  source,
}: {
  objectIds: readonly string[]
  slideId: string
  source: PPTImageCropImportSource
}): SlideEditObjectImageCropHostCommandEffect<string, string>[] {
  const effects: SlideEditObjectImageCropHostCommandEffect<string, string>[] = []

  for (const objectId of objectIds) {
    for (const field of source.fields) {
      if (field === 'fit') {
        const fit = source.imageCrop.fit

        if (fit === undefined) {
          continue
        }

        effects.push(getSlideEditObjectImageCropCommandEffect({
          fieldId: 'fit',
          id: 'update-object-image-crop',
          objectId,
          slideId,
          value: normalizeSlideEditObjectImageCropFit(fit),
        }))
        continue
      }

      const value = source.imageCrop.crop?.[field]

      if (value === undefined) {
        continue
      }

      effects.push(getSlideEditObjectImageCropCommandEffect({
        fieldId: field,
        id: 'update-object-image-crop',
        objectId,
        slideId,
        value: normalizeSlideEditObjectImageCropValue(value),
      }))
    }
  }

  return effects
}

function applyPPTImageCropCommandEffectToElement(
  element: PPTImage,
  effect: SlideEditObjectImageCropHostCommandEffect<string, string>,
): PPTImage {
  const payload = effect.payload

  if (payload.id === 'reset-object-image-crop') {
    return {
      ...element,
      crop: payload.crop,
      fit: normalizePPTImageFit(payload.fit),
    }
  }

  if (payload.fieldId === 'fit') {
    return {
      ...element,
      fit: normalizePPTImageFit(String(payload.value)),
    }
  }

  return {
    ...element,
    crop: {
      ...getPPTImageCrop(element),
      [payload.fieldId]: Number(payload.value),
    },
  }
}

function applyPPTImageReplaceCommandEffectToElement(
  element: PPTImage,
  effect: SlideEditObjectImageReplaceHostCommandEffect<string, string>,
): PPTImage {
  return {
    ...element,
    alt: effect.payload.source.altText ??
      effect.payload.source.name ??
      element.alt,
    name: effect.payload.source.name ?? element.name,
    src: effect.payload.source.src,
  }
}

function createPPTShapeStyleImportEffect({
  effect,
  source,
}: {
  effect: PPTStyleClipboardHostCommandEffect
  source: PPTShapeStyleImportSource
}): PPTShapeStyleImportEffect {
  const payload = effect.payload.id === 'paste-object-formatting'
    ? effect.payload
    : null
  const categories = payload
    ? uniquePPTCanvasValues(payload.categoryApplications.flatMap(
        (application) => application.appliedCategoryIds,
      )).join(' ')
    : ''

  return {
    categories,
    commandId: effect.payload.id,
    commandTargets: payload?.targetObjectIds.join(' ') ?? '',
    commandType: effect.type,
    cornerRadius: source.shape.cornerRadius === undefined
      ? ''
      : String(source.shape.cornerRadius),
    fields: source.fields.join(' '),
    fillColor: source.shape.fill?.color ?? '',
    fillOpacity: source.shape.fill?.opacity === undefined
      ? ''
      : String(source.shape.fill.opacity),
    format: source.format,
    jsonLength: source.jsonLength,
    model: PPT_SHAPE_STYLE_IMPORT_MODEL,
    objectIds: payload?.categoryApplications
      .map((application) => application.objectId)
      .join(' ') ?? '',
    strokeColor: source.shape.stroke?.color ?? '',
    strokeDash: source.shape.stroke ? getPPTStrokeDash(source.shape.stroke) : '',
    strokeWidth: source.shape.stroke?.width === undefined
      ? ''
      : String(source.shape.stroke.width),
  }
}

function createPPTTextStyleImportEffect({
  effect,
  source,
}: {
  effect: PPTStyleClipboardHostCommandEffect
  source: PPTTextStyleImportSource
}): PPTTextStyleImportEffect {
  const payload = effect.payload.id === 'paste-object-formatting'
    ? effect.payload
    : null
  const categories = payload
    ? uniquePPTCanvasValues(payload.categoryApplications.flatMap(
        (application) => application.appliedCategoryIds,
      )).join(' ')
    : ''
  const text = source.text
  const paragraph = source.paragraph

  return {
    categories,
    color: text?.color ?? '',
    commandId: effect.payload.id,
    commandTargets: payload?.targetObjectIds.join(' ') ?? '',
    commandType: effect.type,
    fields: source.fields.join(' '),
    fontFamily: text?.fontFamily ?? '',
    fontSize: text?.fontSize === undefined ? '' : String(text.fontSize),
    fontWeight: text?.fontWeight ?? '',
    format: source.format,
    jsonLength: source.jsonLength,
    model: PPT_TEXT_STYLE_IMPORT_MODEL,
    objectIds: payload?.categoryApplications
      .map((application) => application.objectId)
      .join(' ') ?? '',
    paragraphAlign: paragraph?.align ?? '',
    paragraphBullet: paragraph?.bullet ?? '',
    paragraphLineHeight: paragraph?.lineHeight === undefined
      ? ''
      : String(paragraph.lineHeight),
    paragraphSpacingAfter: paragraph?.spacingAfter === undefined
      ? ''
      : String(paragraph.spacingAfter),
    paragraphSpacingBefore: paragraph?.spacingBefore === undefined
      ? ''
      : String(paragraph.spacingBefore),
    textInset: text?.textInset
      ? formatPPTTextStyleImportInsetData(text.textInset)
      : '',
    verticalAlign: text?.verticalAlign ?? '',
  }
}

function formatPPTTextStyleImportInsetData(
  inset: Partial<NonNullable<PPTTextStyle['textInset']>>,
) {
  return (['top', 'right', 'bottom', 'left'] as const)
    .map((field) => inset[field] === undefined ? '' : String(inset[field]))
    .join(',')
}

function createPPTTextBodyImportEffect({
  objectIds,
  source,
}: {
  objectIds: readonly string[]
  source: PPTTextBodyImportSource
}): PPTTextBodyImportEffect {
  return {
    commandTargets: objectIds.join(' '),
    format: source.format,
    jsonLength: source.jsonLength,
    mode: source.mode,
    model: PPT_TEXT_BODY_IMPORT_MODEL,
    objectIds: objectIds.join(' '),
    paragraphCount: source.textBody.paragraphs.length,
    runCount: source.textBody.paragraphs.reduce(
      (count, paragraph) => count + paragraph.runs.length,
      0,
    ),
    textLength: readPPTText(source.textBody).length,
  }
}

function createPPTTextAutoFitImportEffect({
  effects,
  source,
}: {
  effects: readonly SlideEditTextAutoFitHostCommandEffect<string, string>[]
  source: PPTTextAutoFitImportSource
}): PPTTextAutoFitImportEffect {
  return {
    commandHandles: effects.map((effect) => effect.payload.handle).join(' '),
    commandIds: effects.map((effect) => effect.payload.id).join(' '),
    commandTargets: effects.map((effect) => effect.payload.objectId).join(' '),
    commandTypes: effects.map((effect) => effect.type).join(' '),
    fields: source.fields.join(' '),
    format: source.format,
    jsonLength: source.jsonLength,
    mode: source.mode,
    model: PPT_TEXT_AUTOFIT_IMPORT_MODEL,
    objectIds: effects.map((effect) => effect.payload.objectId).join(' '),
  }
}

function applyPPTTextAutoFitCommandEffectToElement(
  element: PPTTextElement,
  effect: SlideEditTextAutoFitHostCommandEffect<string, string>,
): PPTTextElement {
  return {
    ...element,
    geometry: updatePPTElementBounds(element, effect.payload.bounds).geometry,
    textAutoFit: PPT_TEXT_AUTOFIT,
  }
}

function createPPTCommentImportEffect({
  objectIds,
  source,
}: {
  objectIds: readonly string[]
  source: PPTCommentImportSource
}): PPTCommentImportEffect {
  return {
    bodyLength: source.comment.body?.length ?? 0,
    commandTargets: objectIds.join(' '),
    createdAt: source.comment.createdAt ?? '',
    fields: source.fields.join(' '),
    format: source.format,
    jsonLength: source.jsonLength,
    messageCount: source.comment.thread?.length ?? 0,
    model: PPT_COMMENT_IMPORT_MODEL,
    objectIds: objectIds.join(' '),
    resolved: source.comment.resolved === undefined
      ? ''
      : String(source.comment.resolved),
  }
}

function createPPTMediaJSONImportEffect({
  result,
  source,
}: {
  result: PPTMediaImportResult
  source: PPTMediaJSONImportSource
}): PPTMediaJSONImportEffect {
  return {
    fields: source.fields.join(' '),
    format: source.format,
    importerId: result.importerId,
    jsonLength: source.jsonLength,
    model: PPT_MEDIA_JSON_IMPORT_MODEL,
    objectId: result.item.id,
    title: source.source.title ?? '',
    url: result.source.url,
  }
}

function clonePPTTextBody(body: PPTTextBody): PPTTextBody {
  return {
    paragraphs: body.paragraphs.map((paragraph) => ({
      ...paragraph,
      runs: paragraph.runs.map((run) => ({ ...run })),
    })),
  }
}

function getPPTTextStyleImportParagraph(
  element: PPTTextElement,
  source: PPTTextStyleImportSource,
): PPTStyleClipboardParagraph | undefined {
  const importedParagraph = source.paragraph

  if (!importedParagraph) {
    return undefined
  }

  const paragraph = element.textBody.paragraphs[0]
  const bullet = source.fields.includes('paragraphBullet')
    ? importedParagraph.bullet ?? undefined
    : paragraph?.bullet

  return {
    align: importedParagraph.align ?? paragraph?.align ?? 'left',
    ...(bullet ? { bullet } : {}),
    lineHeight: importedParagraph.lineHeight ??
      (paragraph
        ? getPPTParagraphLineHeight(paragraph)
        : PPT_PARAGRAPH_LINE_HEIGHT_DEFAULT),
    spacingAfter: importedParagraph.spacingAfter ??
      (paragraph ? getPPTParagraphSpacingAfter(paragraph) : 0),
    spacingBefore: importedParagraph.spacingBefore ??
      (paragraph ? getPPTParagraphSpacingBefore(paragraph) : 0),
  }
}

function createPPTLineStyleImportEffect({
  effect,
  source,
}: {
  effect: PPTStyleClipboardHostCommandEffect
  source: PPTLineStyleImportSource
}): PPTLineStyleImportEffect {
  const payload = effect.payload.id === 'paste-object-formatting'
    ? effect.payload
    : null
  const categories = payload
    ? uniquePPTCanvasValues(payload.categoryApplications.flatMap(
        (application) => application.appliedCategoryIds,
      )).join(' ')
    : ''

  return {
    categories,
    commandId: effect.payload.id,
    commandTargets: payload?.targetObjectIds.join(' ') ?? '',
    commandType: effect.type,
    fields: source.fields.join(' '),
    format: source.format,
    jsonLength: source.jsonLength,
    model: PPT_LINE_STYLE_IMPORT_MODEL,
    objectIds: payload?.categoryApplications
      .map((application) => application.objectId)
      .join(' ') ?? '',
    strokeColor: source.stroke.color ?? '',
    strokeDash: source.stroke.dash ?? '',
    strokeWidth: source.stroke.width === undefined
      ? ''
      : String(source.stroke.width),
  }
}

function createPPTElementsJSONImportEffect(
  source: PPTElementsJSONImportSource,
): PPTElementsJSONImportEffect {
  return {
    format: source.format,
    importedObjectCount: source.objects.length,
    jsonLength: source.jsonLength,
    model: PPT_ELEMENTS_JSON_IMPORT_MODEL,
    selectedObjectIds: source.selectedObjectIds,
    sourceSlideId: source.sourceSlideId,
  }
}

function createPPTDeckMarkdownOutlineImportEffect({
  importedSlides,
  source,
}: {
  importedSlides: readonly PPTSlide[]
  source: PPTDeckMarkdownOutlineSource
}): PPTDeckMarkdownOutlineImportEffect {
  return {
    firstImportedSlideId: importedSlides[0]?.id ?? '',
    format: PPT_DECK_MARKDOWN_OUTLINE_IMPORT_FORMAT,
    importedSlideCount: importedSlides.length,
    model: PPT_DECK_MARKDOWN_OUTLINE_IMPORT_MODEL,
    sourceSlideCount: source.slideCount,
    sourceTitle: source.title,
    textLength: source.textLength,
  }
}

function createPPTSlideSVGClipboardEffect({
  sourceSlideId,
  svg,
  writeMode,
}: {
  sourceSlideId: string
  svg: string
  writeMode?: PPTRichClipboardWriteMode
}): PPTSlideSVGClipboardEffect {
  return {
    jsonMimeType: PPT_SLIDE_SVG_CLIPBOARD_JSON_MIME_TYPE,
    model: PPT_SLIDE_SVG_CLIPBOARD_MODEL,
    sourceSlideId,
    svgLength: svg.length,
    writeMode,
  }
}

function createPPTSelectionSVGClipboardEffect({
  selectedObjectIds,
  sourceSlideId,
  svg,
  writeMode,
}: {
  selectedObjectIds: readonly string[]
  sourceSlideId: string
  svg: string
  writeMode?: PPTRichClipboardWriteMode
}): PPTSelectionSVGClipboardEffect {
  return {
    jsonMimeType: PPT_SELECTION_SVG_CLIPBOARD_JSON_MIME_TYPE,
    model: PPT_SELECTION_SVG_CLIPBOARD_MODEL,
    selectedObjectIds,
    sourceSlideId,
    svgLength: svg.length,
    writeMode,
  }
}

function createPPTTableClipboardEffect({
  html,
  objectId,
  plainText,
  rows,
  sourceSlideId,
  writeMode,
}: {
  html: string
  objectId: string
  plainText: string
  rows: readonly (readonly string[])[]
  sourceSlideId: string
  writeMode?: PPTRichClipboardWriteMode
}): PPTTableClipboardEffect {
  return {
    columnCount: getPPTTableColumnCount(rows),
    htmlLength: html.length,
    jsonMimeType: PPT_TABLE_CLIPBOARD_JSON_MIME_TYPE,
    model: PPT_TABLE_CLIPBOARD_MODEL,
    objectId,
    plainTextLength: plainText.length,
    rowCount: rows.length,
    sourceSlideId,
    writeMode,
  }
}

function createPPTTableRowsImportEffect({
  objectIds,
  rows,
  source,
}: {
  objectIds: readonly string[]
  rows: readonly (readonly string[])[]
  source: PPTTableRowsImportSource
}): PPTTableRowsImportEffect {
  return {
    columnCount: getPPTTableColumnCount(rows),
    commandTargets: objectIds.join(' '),
    format: source.format,
    jsonLength: source.jsonLength,
    model: PPT_TABLE_ROWS_IMPORT_MODEL,
    objectIds: objectIds.join(' '),
    rowCount: rows.length,
  }
}

function createPPTRichClipboardExportPayload(
  payload: PPTClipboardPayload,
): PPTRichClipboardExportPayload {
  return {
    kind: PPT_RICH_CLIPBOARD_KIND,
    metadata: {
      objectCount: payload.objects.length,
      selectedObjectIds: payload.selectedObjectIds,
      sourceSlideId: payload.sourceSlideId,
    },
    payload,
    version: PPT_RICH_CLIPBOARD_VERSION,
  }
}

function stringifyPPTRichClipboardPayload(payload: PPTClipboardPayload) {
  return stringifyPPTCanvasRichClipboardPayload(
    createPPTRichClipboardExportPayload(payload),
  )
}

function createPPTSlideClipboardExportPayload(
  payload: PPTSlideClipboardPayload,
): PPTSlideClipboardExportPayload {
  return {
    kind: PPT_SLIDE_CLIPBOARD_KIND,
    metadata: {
      elementCount: payload.slide.elements.length,
      sourceSlideId: payload.sourceSlideId,
    },
    payload,
    version: PPT_SLIDE_CLIPBOARD_VERSION,
  }
}

function stringifyPPTSlideClipboardPayload(payload: PPTSlideClipboardPayload) {
  return stringifyPPTCanvasRichClipboardPayload(
    createPPTSlideClipboardExportPayload(payload),
  )
}

function createPPTSlideClipboardPlainText(slide: PPTSlide) {
  const body = slide.elements
    .map(createPPTElementClipboardPlainText)
    .map((text) => text.trim())
    .filter(Boolean)
    .join('\n\n')

  return [slide.name, body].filter(Boolean).join('\n\n') || 'PPT slide'
}

function createPPTSlideClipboardFallbackHTML(payload: PPTSlideClipboardPayload) {
  return [
    `<section data-ppt-slide-export="${escapePPTCanvasXmlAttribute(payload.slide.id)}" data-ppt-slide-name="${escapePPTCanvasXmlAttribute(payload.slide.name)}">`,
    exportPPTSlideSVG(payload.slide),
    '</section>',
  ].join('')
}

function createPPTSlideClipboardHTML(payload: PPTSlideClipboardPayload) {
  return createPPTCanvasRichClipboardHTML({
    fallbackHTML: createPPTSlideClipboardFallbackHTML(payload),
    json: stringifyPPTSlideClipboardPayload(payload),
    rootAttribute: PPT_SLIDE_CLIPBOARD_HTML_ROOT_ATTRIBUTE,
    scriptAttribute: PPT_SLIDE_CLIPBOARD_HTML_JSON_SCRIPT_ATTRIBUTE,
  })
}

function getPPTSlideFallbackHTMLSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  return getPPTSlideFallbackHTMLSourceFromHTML(
    dataTransfer?.getData('text/html') ?? '',
  )
}

function getPPTSlideFallbackHTMLSourceFromHTML(
  html: string,
): PPTSlideClipboardFallbackHTMLSource | null {
  if (!html || typeof DOMParser === 'undefined') {
    return null
  }

  const doc = new DOMParser().parseFromString(html, 'text/html')
  const element = doc.querySelector<HTMLElement>('[data-ppt-slide-export]')
  const svg = element?.querySelector<SVGSVGElement>('svg')

  if (!element || !svg) {
    return null
  }

  svg.querySelectorAll('script').forEach((node) => node.remove())

  const sourceSlideId =
    element.getAttribute('data-ppt-slide-export')?.trim() || undefined
  const name =
    element.getAttribute('data-ppt-slide-name')?.trim() ||
    svg.getAttribute('data-ppt-svg-slide')?.trim() ||
    'PPT Slide'

  return {
    htmlLength: html.length,
    name,
    ...(sourceSlideId ? { sourceSlideId } : {}),
    svg: svg.outerHTML,
  }
}

function getPPTDeckHTMLSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  const html = dataTransfer?.getData('text/html') ?? ''
  const plainText = dataTransfer?.getData('text/plain') ?? ''

  return getPPTDeckHTMLSourceFromHTML(html) ??
    getPPTDeckHTMLSourceFromHTML(plainText)
}

function getPPTDeckHTMLSourceFromHTML(
  html: string,
): PPTDeckHTMLImportSource | null {
  if (!html || typeof DOMParser === 'undefined') {
    return null
  }

  const doc = new DOMParser().parseFromString(html, 'text/html')
  const script = doc.querySelector<HTMLScriptElement>(
    'script[type="application/json"][data-ppt-deck],script[data-ppt-deck]',
  )
  const json = script?.textContent?.trim() ?? ''

  if (!json) {
    return null
  }

  try {
    const parsed = PPTDeckSchema.safeParse(JSON.parse(json))

    if (!parsed.success) {
      return null
    }

    return {
      deck: parsed.data,
      htmlLength: html.length,
    }
  } catch {
    return null
  }
}

function getPPTDeckJSONSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  if (!dataTransfer) {
    return null
  }

  const candidates: Array<{
    format: PPTDeckJSONImportSource['format']
    text: string
  }> = [
    {
      format: PPT_DECK_JSON_IMPORT_FORMAT,
      text: dataTransfer.getData(PPT_DECK_JSON_MIME_TYPE),
    },
    {
      format: PPT_DECK_JSON_IMPORT_FORMAT,
      text: dataTransfer.getData('application/json'),
    },
    {
      format: PPT_DECK_JSON_TEXT_IMPORT_FORMAT,
      text: dataTransfer.getData('text/json'),
    },
    {
      format: PPT_DECK_JSON_TEXT_IMPORT_FORMAT,
      text: dataTransfer.getData('text/markdown'),
    },
    {
      format: PPT_DECK_JSON_TEXT_IMPORT_FORMAT,
      text: dataTransfer.getData('text/plain'),
    },
  ]
  const seen = new Set<string>()

  for (const candidate of candidates) {
    const text = candidate.text.trim()

    if (!text || seen.has(text)) {
      continue
    }

    seen.add(text)

    const source = getPPTDeckJSONSourceFromText(text, candidate.format)

    if (source) {
      return source
    }
  }

  return null
}

function getPPTDeckJSONSourceFromText(
  text: string,
  format: PPTDeckJSONImportSource['format'],
): PPTDeckJSONImportSource | null {
  const json = getPPTImportJSONText(text)

  if (!json) {
    return null
  }

  try {
    const deck = getPPTDeckFromJSONValue(JSON.parse(json))

    return deck
      ? {
          deck,
          format,
          jsonLength: json.length,
        }
      : null
  } catch {
    return null
  }
}

function getPPTSlideJSONSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  if (!dataTransfer) {
    return null
  }

  const candidates: Array<{
    format: PPTSlideJSONImportSource['format']
    text: string
  }> = [
    {
      format: PPT_SLIDE_JSON_IMPORT_FORMAT,
      text: dataTransfer.getData(PPT_SLIDE_JSON_MIME_TYPE),
    },
    {
      format: PPT_SLIDE_JSON_IMPORT_FORMAT,
      text: dataTransfer.getData('application/json'),
    },
    {
      format: PPT_SLIDE_JSON_TEXT_IMPORT_FORMAT,
      text: dataTransfer.getData('text/json'),
    },
    {
      format: PPT_SLIDE_JSON_TEXT_IMPORT_FORMAT,
      text: dataTransfer.getData('text/markdown'),
    },
    {
      format: PPT_SLIDE_JSON_TEXT_IMPORT_FORMAT,
      text: dataTransfer.getData('text/plain'),
    },
  ]
  const seen = new Set<string>()

  for (const candidate of candidates) {
    const text = candidate.text.trim()

    if (!text || seen.has(text)) {
      continue
    }

    seen.add(text)

    const source = getPPTSlideJSONSourceFromText(text, candidate.format)

    if (source) {
      return source
    }
  }

  return null
}

function getPPTSlideJSONSourceFromText(
  text: string,
  format: PPTSlideJSONImportSource['format'],
): PPTSlideJSONImportSource | null {
  const json = getPPTImportJSONText(text)

  if (!json) {
    return null
  }

  try {
    const slides = getPPTSlidesFromJSONValue(JSON.parse(json))

    return slides.length > 0
      ? {
          format,
          jsonLength: json.length,
          slides,
        }
      : null
  } catch {
    return null
  }
}

function getPPTSlideMetadataSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  if (!dataTransfer) {
    return null
  }

  const candidates = [
    dataTransfer.getData(PPT_SLIDE_METADATA_JSON_MIME_TYPE),
    dataTransfer.getData('application/json'),
    dataTransfer.getData('text/json'),
    dataTransfer.getData('text/plain'),
  ]
  const seen = new Set<string>()

  for (const candidate of candidates) {
    const text = candidate.trim()

    if (!text || seen.has(text)) {
      continue
    }

    seen.add(text)

    const source = getPPTSlideMetadataSourceFromText(text)

    if (source) {
      return source
    }
  }

  return null
}

function getPPTSlideMetadataSourceFromText(
  text: string,
): PPTSlideMetadataImportSource | null {
  const json = getPPTImportJSONText(text)

  if (!json) {
    return null
  }

  try {
    return getPPTSlideMetadataSourceFromJSONValue(JSON.parse(json), json.length)
  } catch {
    return null
  }
}

function getPPTSlideMetadataSourceFromJSONValue(
  value: unknown,
  jsonLength: number,
): PPTSlideMetadataImportSource | null {
  const payloadValue = isPPTRecord(value) &&
    isPPTRecord(value.slideMetadata)
    ? value.slideMetadata
    : value

  if (!isPPTRecord(payloadValue)) {
    return null
  }

  const name = getPPTSlideMetadataNameFromJSONValue(payloadValue)
  const background = getPPTSlideMetadataBackgroundFromJSONValue(payloadValue)

  if (name === undefined && background === undefined) {
    return null
  }

  const notesValue = payloadValue.notes
  const notes = typeof notesValue === 'string'
    ? normalizePPTSlideNotesText(notesValue)
    : null

  return {
    ...(background !== undefined ? { background } : {}),
    format: PPT_SLIDE_METADATA_JSON_IMPORT_FORMAT,
    jsonLength,
    ...(name !== undefined ? { name } : {}),
    ...(notes !== null ? { notes } : {}),
  }
}

function getPPTSlideMetadataNameFromJSONValue(
  value: Record<string, unknown>,
) {
  const nameValue = value.name ?? value.slideName

  return typeof nameValue === 'string'
    ? normalizePPTSlideMetadataName(nameValue)
    : undefined
}

function getPPTSlideMetadataBackgroundFromJSONValue(
  value: Record<string, unknown>,
): PPTSlideMetadataBackgroundDescriptor | undefined {
  const backgroundValue = value.background ?? value.backgroundColor
  const colorValue = typeof backgroundValue === 'string'
    ? backgroundValue
    : isPPTRecord(backgroundValue)
      ? backgroundValue.color
      : undefined
  const color = typeof colorValue === 'string'
    ? normalizePPTSlideMetadataColor(colorValue)
    : null

  return color
    ? {
        color,
        kind: 'solid-color',
      }
    : undefined
}

function normalizePPTSlideMetadataName(name: string) {
  const normalized = name.replace(/\s+/g, ' ').trim().slice(0, 80).trim()

  return normalized || undefined
}

function normalizePPTSlideMetadataColor(color: string) {
  const normalized = color.trim()

  return /^#[\da-f]{6}$/i.test(normalized) ? normalized : null
}

function getPPTSlideLayoutSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  if (!dataTransfer) {
    return null
  }

  const candidates = [
    dataTransfer.getData(PPT_SLIDE_LAYOUT_JSON_MIME_TYPE),
    dataTransfer.getData('application/json'),
    dataTransfer.getData('text/json'),
    dataTransfer.getData('text/plain'),
  ]
  const seen = new Set<string>()

  for (const candidate of candidates) {
    const text = candidate.trim()

    if (!text || seen.has(text)) {
      continue
    }

    seen.add(text)

    const source = getPPTSlideLayoutSourceFromText(text)

    if (source) {
      return source
    }
  }

  return null
}

function getPPTSlideLayoutSourceFromText(
  text: string,
): PPTSlideLayoutImportSource | null {
  const json = getPPTImportJSONText(text)

  if (!json) {
    return null
  }

  try {
    return getPPTSlideLayoutSourceFromJSONValue(JSON.parse(json), json.length)
  } catch {
    return null
  }
}

function getPPTSlideLayoutSourceFromJSONValue(
  value: unknown,
  jsonLength: number,
): PPTSlideLayoutImportSource | null {
  const payloadValue = getPPTSlideLayoutPayloadValue(value)

  if (typeof payloadValue === 'string') {
    const layoutId = getPPTSlideLayoutIdFromJSONValue(payloadValue)

    return layoutId
      ? {
          fields: ['layoutId'],
          format: PPT_SLIDE_LAYOUT_JSON_IMPORT_FORMAT,
          jsonLength,
          layoutId,
        }
      : null
  }

  if (!isPPTRecord(payloadValue)) {
    return null
  }

  const layoutId = getPPTSlideLayoutIdFromJSONValue(
    payloadValue.layoutId ?? payloadValue.layout,
  )
  const themeId = getPPTSlideLayoutThemeIdFromJSONValue(
    payloadValue.themeId ?? payloadValue.theme,
  )
  const hiddenPlaceholderIds = getPPTSlideLayoutHiddenPlaceholderIdsFromJSONValue(
    payloadValue,
    layoutId,
  )
  const fields: PPTSlideLayoutImportField[] = []

  if (layoutId !== undefined) {
    fields.push('layoutId')
  }

  if (themeId !== undefined) {
    fields.push('themeId')
  }

  if (hiddenPlaceholderIds !== undefined) {
    fields.push('hiddenPlaceholderIds')
  }

  return fields.length > 0
    ? {
        fields,
        format: PPT_SLIDE_LAYOUT_JSON_IMPORT_FORMAT,
        ...(hiddenPlaceholderIds === undefined ? {} : { hiddenPlaceholderIds }),
        jsonLength,
        ...(layoutId === undefined ? {} : { layoutId }),
        ...(themeId === undefined ? {} : { themeId }),
      }
    : null
}

function getPPTSlideLayoutPayloadValue(value: unknown): unknown {
  if (!isPPTRecord(value)) {
    return value
  }

  if (isPPTRecord(value.slideLayout) || typeof value.slideLayout === 'string') {
    return value.slideLayout
  }

  if (isPPTRecord(value.layout)) {
    return value.layout
  }

  if (isPPTRecord(value.slide)) {
    return value.slide
  }

  return value
}

function getPPTSlideLayoutIdFromJSONValue(value: unknown) {
  if (typeof value !== 'string') {
    return undefined
  }

  const layoutId = value.trim()

  return PPT_LAYOUT_BY_ID.has(layoutId) ? layoutId : undefined
}

function getPPTSlideLayoutThemeIdFromJSONValue(value: unknown) {
  if (typeof value !== 'string') {
    return undefined
  }

  const themeId = value.trim()

  return themeId === PPT_THEME_DESCRIPTOR.themeId ? themeId : undefined
}

function getPPTSlideLayoutHiddenPlaceholderIdsFromJSONValue(
  value: Record<string, unknown>,
  layoutId: string | undefined,
): readonly string[] | undefined {
  const directHiddenPlaceholderIds = getPPTSlideLayoutPlaceholderIdListFromJSONValue(
    value.hiddenPlaceholderIds ?? value.hiddenPlaceholders,
  )

  if (directHiddenPlaceholderIds !== undefined) {
    return uniquePPTCanvasValues(directHiddenPlaceholderIds)
  }

  const placeholdersValue = value.placeholders
  const placeholderRecord = isPPTRecord(placeholdersValue)
    ? placeholdersValue
    : null
  const placeholderHiddenIds = placeholderRecord
    ? getPPTSlideLayoutPlaceholderIdListFromJSONValue(
      placeholderRecord.hidden ?? placeholderRecord.hiddenPlaceholderIds,
    )
    : undefined

  if (placeholderHiddenIds !== undefined) {
    return uniquePPTCanvasValues(placeholderHiddenIds)
  }

  const visibilityHiddenPlaceholderIds =
    getPPTSlideLayoutHiddenPlaceholderIdsFromVisibilityValue(
      value.placeholderVisibility ?? value.visibility ?? placeholdersValue,
    )

  if (visibilityHiddenPlaceholderIds !== undefined) {
    return uniquePPTCanvasValues(visibilityHiddenPlaceholderIds)
  }

  const visiblePlaceholderIds = getPPTSlideLayoutPlaceholderIdListFromJSONValue(
    value.visiblePlaceholderIds ?? value.visiblePlaceholders,
  ) ?? (
    placeholderRecord
      ? getPPTSlideLayoutPlaceholderIdListFromJSONValue(
        placeholderRecord.visible ?? placeholderRecord.visiblePlaceholderIds,
      )
      : undefined
  )

  if (visiblePlaceholderIds !== undefined && layoutId !== undefined) {
    return getPPTSlideLayoutHiddenPlaceholderIdsFromVisibleIds(
      visiblePlaceholderIds,
      getPPTLayoutDescriptor(layoutId),
    )
  }

  return undefined
}

function getPPTSlideLayoutPlaceholderIdListFromJSONValue(
  value: unknown,
): readonly string[] | undefined {
  if (!Array.isArray(value)) {
    return undefined
  }

  return value
    .map(getPPTSlideLayoutPlaceholderIdFromJSONValue)
    .filter((placeholderId): placeholderId is string =>
      placeholderId !== undefined)
}

function getPPTSlideLayoutHiddenPlaceholderIdsFromVisibilityValue(
  value: unknown,
): readonly string[] | undefined {
  if (Array.isArray(value)) {
    const hiddenPlaceholderIds = value
      .map(getPPTSlideLayoutHiddenPlaceholderIdFromVisibilityEntry)
      .filter((placeholderId): placeholderId is string =>
        placeholderId !== undefined)

    return hiddenPlaceholderIds.length > 0 ? hiddenPlaceholderIds : undefined
  }

  if (!isPPTRecord(value)) {
    return undefined
  }

  const hiddenPlaceholderIds = Object.entries(value)
    .map(([placeholderId, visibilityValue]) =>
      getPPTSlideLayoutHiddenPlaceholderIdFromVisibilityPair(
        placeholderId,
        visibilityValue,
      ))
    .filter((placeholderId): placeholderId is string =>
      placeholderId !== undefined)

  return hiddenPlaceholderIds.length > 0 ? hiddenPlaceholderIds : undefined
}

function getPPTSlideLayoutHiddenPlaceholderIdFromVisibilityEntry(
  value: unknown,
) {
  if (!isPPTRecord(value)) {
    return undefined
  }

  const placeholderId = getPPTSlideLayoutPlaceholderIdFromJSONValue(
    value.placeholderId ?? value.id,
  )

  if (!placeholderId) {
    return undefined
  }

  const visible = getPPTSlideLayoutPlaceholderVisibleFromJSONValue(
    value.visible ?? value.isVisible,
  )

  return visible === false ? placeholderId : undefined
}

function getPPTSlideLayoutHiddenPlaceholderIdFromVisibilityPair(
  key: string,
  value: unknown,
) {
  const placeholderId = getPPTSlideLayoutPlaceholderIdFromJSONValue(key)
  const visible = isPPTRecord(value)
    ? getPPTSlideLayoutPlaceholderVisibleFromJSONValue(
      value.visible ?? value.isVisible,
    )
    : getPPTSlideLayoutPlaceholderVisibleFromJSONValue(value)

  return placeholderId && visible === false ? placeholderId : undefined
}

function getPPTSlideLayoutPlaceholderVisibleFromJSONValue(value: unknown) {
  return typeof value === 'boolean' ? value : undefined
}

function getPPTSlideLayoutHiddenPlaceholderIdsFromVisibleIds(
  visiblePlaceholderIds: readonly string[],
  layout: SlideEditLayoutDescriptor,
) {
  const visiblePlaceholderIdSet = new Set(visiblePlaceholderIds)

  return layout.placeholders
    .map((placeholder) => placeholder.placeholderId)
    .filter((placeholderId) => !visiblePlaceholderIdSet.has(placeholderId))
}

function getPPTSlideLayoutPlaceholderIdFromJSONValue(value: unknown) {
  if (typeof value !== 'string') {
    return undefined
  }

  const placeholderId = value.trim()

  return placeholderId || undefined
}

function getPPTSlideTransitionSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  if (!dataTransfer) {
    return null
  }

  const candidates = [
    dataTransfer.getData(PPT_SLIDE_TRANSITION_JSON_MIME_TYPE),
    dataTransfer.getData('application/json'),
    dataTransfer.getData('text/json'),
    dataTransfer.getData('text/plain'),
  ]
  const seen = new Set<string>()

  for (const candidate of candidates) {
    const text = candidate.trim()

    if (!text || seen.has(text)) {
      continue
    }

    seen.add(text)

    const source = getPPTSlideTransitionSourceFromText(text)

    if (source) {
      return source
    }
  }

  return null
}

function getPPTSlideTransitionSourceFromText(
  text: string,
): PPTSlideTransitionImportSource | null {
  const json = getPPTImportJSONText(text)

  if (!json) {
    return null
  }

  try {
    return getPPTSlideTransitionSourceFromJSONValue(
      JSON.parse(json),
      json.length,
    )
  } catch {
    return null
  }
}

function getPPTSlideTransitionSourceFromJSONValue(
  value: unknown,
  jsonLength: number,
): PPTSlideTransitionImportSource | null {
  const payloadValue = isPPTRecord(value) &&
    isPPTRecord(value.transition)
    ? value.transition
    : value

  if (!isPPTRecord(payloadValue)) {
    return null
  }

  const transition: Partial<PPTSlideTransition> = {}
  const fields: PPTSlideTransitionImportField[] = []
  const type = getPPTSlideTransitionTypeFromJSONValue(payloadValue.type)
  const durationMs = getPPTSlideTransitionDurationFromJSONValue(
    payloadValue.durationMs ?? payloadValue.duration,
  )
  const advanceOnClick = typeof payloadValue.advanceOnClick === 'boolean'
    ? payloadValue.advanceOnClick
    : isPPTRecord(payloadValue.advance) &&
        typeof payloadValue.advance.onClick === 'boolean'
      ? payloadValue.advance.onClick
      : undefined
  const advanceAfterMs = getPPTSlideTransitionAdvanceAfterFromJSONValue(
    payloadValue.advanceAfterMs ??
      payloadValue.advanceAfter ??
      (isPPTRecord(payloadValue.advance) ? payloadValue.advance.afterMs : undefined),
  )

  if (type !== undefined) {
    transition.type = type
    fields.push('type')
  }

  if (durationMs !== undefined) {
    transition.durationMs = durationMs
    fields.push('durationMs')
  }

  if (advanceOnClick !== undefined) {
    transition.advanceOnClick = advanceOnClick
    fields.push('advanceOnClick')
  }

  if (advanceAfterMs !== undefined) {
    transition.advanceAfterMs = advanceAfterMs
    fields.push('advanceAfterMs')
  }

  return fields.length > 0
    ? {
        fields,
        format: PPT_SLIDE_TRANSITION_JSON_IMPORT_FORMAT,
        jsonLength,
        transition,
      }
    : null
}

function getPPTSlideTransitionTypeFromJSONValue(
  value: unknown,
): PPTSlideTransitionType | undefined {
  const type = typeof value === 'string' ? value.trim() : ''

  return isPPTSlideTransitionType(type) ? type : undefined
}

function getPPTSlideTransitionDurationFromJSONValue(value: unknown) {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    return undefined
  }

  return clampPPTSlideTransitionDuration(value)
}

function getPPTSlideTransitionAdvanceAfterFromJSONValue(value: unknown) {
  if (value === null) {
    return null
  }

  if (typeof value !== 'number' || !Number.isFinite(value)) {
    return undefined
  }

  return parsePPTSlideTransitionAdvanceAfter(String(value))
}

function getPPTObjectAnimationSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  if (!dataTransfer) {
    return null
  }

  const candidates: Array<{
    allowDirect: boolean
    text: string
  }> = [
    {
      allowDirect: true,
      text: dataTransfer.getData(PPT_OBJECT_ANIMATION_JSON_MIME_TYPE),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('application/json'),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('text/json'),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('text/plain'),
    },
  ]
  const seen = new Set<string>()

  for (const candidate of candidates) {
    const text = candidate.text.trim()

    if (!text || seen.has(text)) {
      continue
    }

    seen.add(text)

    const source = getPPTObjectAnimationSourceFromText(
      text,
      candidate.allowDirect,
    )

    if (source) {
      return source
    }
  }

  return null
}

function getPPTObjectAnimationSourceFromText(
  text: string,
  allowDirect: boolean,
): PPTObjectAnimationImportSource | null {
  const json = getPPTImportJSONText(text)

  if (!json) {
    return null
  }

  try {
    return getPPTObjectAnimationSourceFromJSONValue(
      JSON.parse(json),
      json.length,
      allowDirect,
    )
  } catch {
    return null
  }
}

function getPPTObjectAnimationSourceFromJSONValue(
  value: unknown,
  jsonLength: number,
  allowDirect: boolean,
): PPTObjectAnimationImportSource | null {
  const payloadValue = isPPTRecord(value) &&
    isPPTRecord(value.animation)
    ? value.animation
    : allowDirect
      ? value
      : null

  if (!isPPTRecord(payloadValue)) {
    return null
  }

  const animation: Partial<PPTElementAnimation> = {}
  const fields: PPTObjectAnimationImportField[] = []
  const type = getPPTElementAnimationTypeFromJSONValue(payloadValue.type)
  const trigger = getPPTElementAnimationTriggerFromJSONValue(
    payloadValue.trigger,
  )
  const durationMs = getPPTElementAnimationTimeFromJSONValue(
    payloadValue.durationMs ?? payloadValue.duration,
  )
  const delayMs = getPPTElementAnimationTimeFromJSONValue(
    payloadValue.delayMs ?? payloadValue.delay,
  )
  const order = getPPTElementAnimationOrderFromJSONValue(
    payloadValue.order ?? payloadValue.buildOrder,
  )

  if (type !== undefined) {
    animation.type = type
    fields.push('type')
  }

  if (trigger !== undefined) {
    animation.trigger = trigger
    fields.push('trigger')
  }

  if (durationMs !== undefined) {
    animation.durationMs = durationMs
    fields.push('durationMs')
  }

  if (delayMs !== undefined) {
    animation.delayMs = delayMs
    fields.push('delayMs')
  }

  if (order !== undefined) {
    animation.order = order
    fields.push('order')
  }

  return fields.length > 0
    ? {
        animation,
        fields,
        format: PPT_OBJECT_ANIMATION_JSON_IMPORT_FORMAT,
        jsonLength,
      }
    : null
}

function getPPTElementAnimationTypeFromJSONValue(
  value: unknown,
): PPTElementAnimationType | undefined {
  const type = typeof value === 'string' ? value.trim() : ''

  return (PPT_ELEMENT_ANIMATION_TYPES as readonly string[]).includes(type)
    ? type as PPTElementAnimationType
    : undefined
}

function getPPTElementAnimationTriggerFromJSONValue(
  value: unknown,
): PPTElementAnimationTrigger | undefined {
  const trigger = typeof value === 'string' ? value.trim() : ''

  return (PPT_ELEMENT_ANIMATION_TRIGGERS as readonly string[]).includes(trigger)
    ? trigger as PPTElementAnimationTrigger
    : undefined
}

function getPPTElementAnimationTimeFromJSONValue(value: unknown) {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    return undefined
  }

  return clampPPTElementAnimationTime(value)
}

function getPPTElementAnimationOrderFromJSONValue(value: unknown) {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    return undefined
  }

  return clampPPTElementAnimationOrder(value)
}

function getPPTObjectStyleSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  if (!dataTransfer) {
    return null
  }

  const candidates: Array<{
    allowDirect: boolean
    text: string
  }> = [
    {
      allowDirect: true,
      text: dataTransfer.getData(PPT_OBJECT_STYLE_JSON_MIME_TYPE),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('application/json'),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('text/json'),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('text/plain'),
    },
  ]
  const seen = new Set<string>()

  for (const candidate of candidates) {
    const text = candidate.text.trim()

    if (!text || seen.has(text)) {
      continue
    }

    seen.add(text)

    const source = getPPTObjectStyleSourceFromText(text, candidate.allowDirect)

    if (source) {
      return source
    }
  }

  return null
}

function getPPTObjectStyleSourceFromText(
  text: string,
  allowDirect: boolean,
): PPTObjectStyleImportSource | null {
  const json = getPPTImportJSONText(text)

  if (!json) {
    return null
  }

  try {
    return getPPTObjectStyleSourceFromJSONValue(
      JSON.parse(json),
      json.length,
      allowDirect,
    )
  } catch {
    return null
  }
}

function getPPTObjectStyleSourceFromJSONValue(
  value: unknown,
  jsonLength: number,
  allowDirect: boolean,
): PPTObjectStyleImportSource | null {
  const payloadValue = isPPTRecord(value) &&
    isPPTRecord(value.objectStyle)
    ? value.objectStyle
    : isPPTRecord(value) && isPPTRecord(value.objectEffect)
      ? value.objectEffect
      : allowDirect
        ? value
        : null

  if (!isPPTRecord(payloadValue)) {
    return null
  }

  const object: PPTObjectStyleImportSource['object'] = {}
  const fields: PPTObjectStyleImportField[] = []
  const opacity = getPPTObjectStyleOpacityFromJSONValue(payloadValue.opacity)
  const shadow = getPPTObjectStyleShadowFromJSONValue(payloadValue.shadow)

  if (opacity !== undefined) {
    object.opacity = opacity
    fields.push('opacity')
  }

  if (shadow !== undefined) {
    object.shadow = shadow
    fields.push('shadow')
  }

  return fields.length > 0
    ? {
        fields,
        format: PPT_OBJECT_STYLE_JSON_IMPORT_FORMAT,
        jsonLength,
        object,
      }
    : null
}

function getPPTObjectStyleOpacityFromJSONValue(value: unknown) {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    return undefined
  }

  return normalizePPTElementOpacity(value)
}

function getPPTObjectStyleShadowFromJSONValue(
  value: unknown,
): PPTElementShadow | null | undefined {
  if (value === null || value === false) {
    return null
  }

  if (value === true) {
    return normalizePPTElementShadow({})
  }

  if (!isPPTRecord(value)) {
    return undefined
  }

  if (value.enabled === false) {
    return null
  }

  return normalizePPTElementShadow({
    angle: typeof value.angle === 'number' ? value.angle : undefined,
    blur: typeof value.blur === 'number' ? value.blur : undefined,
    color: typeof value.color === 'string' ? value.color : undefined,
    distance: typeof value.distance === 'number' ? value.distance : undefined,
    opacity: typeof value.opacity === 'number' ? value.opacity : undefined,
  })
}

function getPPTObjectMetadataSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  if (!dataTransfer) {
    return null
  }

  const candidates: Array<{
    allowDirect: boolean
    text: string
  }> = [
    {
      allowDirect: true,
      text: dataTransfer.getData(PPT_OBJECT_METADATA_JSON_MIME_TYPE),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('application/json'),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('text/json'),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('text/plain'),
    },
  ]
  const seen = new Set<string>()

  for (const candidate of candidates) {
    const text = candidate.text.trim()

    if (!text || seen.has(text)) {
      continue
    }

    seen.add(text)

    const source = getPPTObjectMetadataSourceFromText(
      text,
      candidate.allowDirect,
    )

    if (source) {
      return source
    }
  }

  return null
}

function getPPTObjectMetadataSourceFromText(
  text: string,
  allowDirect: boolean,
): PPTObjectMetadataImportSource | null {
  const json = getPPTImportJSONText(text)

  if (!json) {
    return null
  }

  try {
    return getPPTObjectMetadataSourceFromJSONValue(
      JSON.parse(json),
      json.length,
      allowDirect,
    )
  } catch {
    return null
  }
}

function getPPTObjectMetadataSourceFromJSONValue(
  value: unknown,
  jsonLength: number,
  allowDirect: boolean,
): PPTObjectMetadataImportSource | null {
  const payloadValue = getPPTObjectMetadataPayloadValue(value, allowDirect)

  if (!isPPTRecord(payloadValue)) {
    return null
  }

  const metadata: PPTObjectMetadataImportSource['metadata'] = {}
  const fields: PPTObjectMetadataImportField[] = []
  const hyperlinkUrl = getPPTObjectMetadataHyperlinkURLFromJSONValue(
    payloadValue.hyperlink ??
      payloadValue.hyperlinkUrl ??
      payloadValue.url,
  )
  const altText = getPPTObjectMetadataAltTextFromJSONValue(
    isPPTRecord(payloadValue.accessibility)
      ? payloadValue.accessibility.altText
      : payloadValue.altText,
  )
  const name = getPPTObjectMetadataNameFromJSONValue(
    payloadValue.name ??
      payloadValue.objectName ??
      payloadValue.layerName,
  )

  if (hyperlinkUrl !== undefined) {
    metadata.hyperlinkUrl = hyperlinkUrl
    fields.push('hyperlinkUrl')
  }

  if (altText !== undefined) {
    metadata.altText = altText
    fields.push('altText')
  }

  if (name !== undefined) {
    metadata.name = name
    fields.push('name')
  }

  return fields.length > 0
    ? {
        fields,
        format: PPT_OBJECT_METADATA_JSON_IMPORT_FORMAT,
        jsonLength,
        metadata,
      }
    : null
}

function getPPTObjectMetadataPayloadValue(
  value: unknown,
  allowDirect: boolean,
): unknown {
  if (!isPPTRecord(value)) {
    return null
  }

  if (isPPTRecord(value.objectMetadata)) {
    return value.objectMetadata
  }

  if (
    typeof value.objectName === 'string' ||
    isPPTRecord(value.objectName)
  ) {
    return typeof value.objectName === 'string'
      ? { name: value.objectName }
      : value.objectName
  }

  if (
    typeof value.layerName === 'string' ||
    isPPTRecord(value.layerName)
  ) {
    return typeof value.layerName === 'string'
      ? { name: value.layerName }
      : value.layerName
  }

  return allowDirect ? value : null
}

function getPPTObjectMetadataHyperlinkURLFromJSONValue(
  value: unknown,
): string | null | undefined {
  if (value === null || value === false) {
    return null
  }

  const rawUrl = typeof value === 'string'
    ? value
    : isPPTRecord(value) && typeof value.url === 'string'
      ? value.url
      : null

  if (rawUrl === null) {
    return undefined
  }

  const trimmed = rawUrl.trim()

  if (!trimmed) {
    return null
  }

  return normalizePPTElementHyperlinkUrl(trimmed) || undefined
}

function getPPTObjectMetadataAltTextFromJSONValue(
  value: unknown,
): string | null | undefined {
  if (value === null || value === false) {
    return null
  }

  if (typeof value !== 'string') {
    return undefined
  }

  return normalizePPTAltText(value) || null
}

function getPPTObjectMetadataNameFromJSONValue(value: unknown) {
  if (typeof value !== 'string') {
    return undefined
  }

  const name = value.trim()

  return name.length > 0 ? name : undefined
}

function getPPTObjectStateSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  if (!dataTransfer) {
    return null
  }

  const candidates: Array<{
    allowDirect: boolean
    text: string
  }> = [
    {
      allowDirect: true,
      text: dataTransfer.getData(PPT_OBJECT_STATE_JSON_MIME_TYPE),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('application/json'),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('text/json'),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('text/plain'),
    },
  ]
  const seen = new Set<string>()

  for (const candidate of candidates) {
    const text = candidate.text.trim()

    if (!text || seen.has(text)) {
      continue
    }

    seen.add(text)

    const source = getPPTObjectStateSourceFromText(
      text,
      candidate.allowDirect,
    )

    if (source) {
      return source
    }
  }

  return null
}

function getPPTObjectStateSourceFromText(
  text: string,
  allowDirect: boolean,
): PPTObjectStateImportSource | null {
  const json = getPPTImportJSONText(text)

  if (!json) {
    return null
  }

  try {
    return getPPTObjectStateSourceFromJSONValue(
      JSON.parse(json),
      json.length,
      allowDirect,
    )
  } catch {
    return null
  }
}

function getPPTObjectStateSourceFromJSONValue(
  value: unknown,
  jsonLength: number,
  allowDirect: boolean,
): PPTObjectStateImportSource | null {
  const payloadValue = getPPTObjectStatePayloadValue(value, allowDirect)

  if (!isPPTRecord(payloadValue)) {
    return null
  }

  const state: PPTObjectStateImportSource['state'] = {}
  const fields: PPTObjectStateImportField[] = []
  const visible = getPPTObjectStateVisibleFromJSONValue(payloadValue)
  const locked = getPPTObjectStateLockedFromJSONValue(payloadValue.locked)

  if (visible !== undefined) {
    state.visible = visible
    fields.push('visible')
  }

  if (locked !== undefined) {
    state.locked = locked
    fields.push('locked')
  }

  return fields.length > 0
    ? {
        fields,
        format: PPT_OBJECT_STATE_JSON_IMPORT_FORMAT,
        jsonLength,
        state,
      }
    : null
}

function getPPTObjectStatePayloadValue(
  value: unknown,
  allowDirect: boolean,
): unknown {
  if (!isPPTRecord(value)) {
    return null
  }

  if (isPPTRecord(value.objectState)) {
    return value.objectState
  }

  if (isPPTRecord(value.objectVisibility)) {
    return value.objectVisibility
  }

  if (isPPTRecord(value.layerState)) {
    return value.layerState
  }

  if (isPPTRecord(value.selectionState)) {
    return value.selectionState
  }

  if (
    value.visible !== undefined ||
    value.hidden !== undefined ||
    value.locked !== undefined
  ) {
    return value
  }

  return allowDirect ? value : null
}

function getPPTObjectStateVisibleFromJSONValue(
  value: Record<string, unknown>,
) {
  if (typeof value.visible === 'boolean') {
    return value.visible
  }

  return typeof value.hidden === 'boolean'
    ? !value.hidden
    : undefined
}

function getPPTObjectStateLockedFromJSONValue(value: unknown) {
  return typeof value === 'boolean' ? value : undefined
}

function getPPTObjectLayerSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  if (!dataTransfer) {
    return null
  }

  const candidates: Array<{
    allowDirect: boolean
    text: string
  }> = [
    {
      allowDirect: true,
      text: dataTransfer.getData(PPT_OBJECT_LAYER_JSON_MIME_TYPE),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('application/json'),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('text/json'),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('text/plain'),
    },
  ]
  const seen = new Set<string>()

  for (const candidate of candidates) {
    const text = candidate.text.trim()

    if (!text || seen.has(text)) {
      continue
    }

    seen.add(text)

    const source = getPPTObjectLayerSourceFromText(
      text,
      candidate.allowDirect,
    )

    if (source) {
      return source
    }
  }

  return null
}

function getPPTObjectLayerSourceFromText(
  text: string,
  allowDirect: boolean,
): PPTObjectLayerImportSource | null {
  const json = getPPTImportJSONText(text)

  if (!json) {
    return null
  }

  try {
    return getPPTObjectLayerSourceFromJSONValue(
      JSON.parse(json),
      json.length,
      allowDirect,
    )
  } catch {
    return null
  }
}

function getPPTObjectLayerSourceFromJSONValue(
  value: unknown,
  jsonLength: number,
  allowDirect: boolean,
): PPTObjectLayerImportSource | null {
  const payloadValue = getPPTObjectLayerPayloadValue(value, allowDirect)

  if (!isPPTRecord(payloadValue)) {
    return null
  }

  const layer: PPTObjectLayerImportSource['layer'] = {}
  const fields: PPTObjectLayerImportField[] = []
  const position = getPPTObjectLayerPositionFromJSONValue(
    payloadValue.position ?? payloadValue.arrange ?? payloadValue.zOrder,
  )
  const toIndex = getPPTObjectLayerToIndexFromJSONValue(
    payloadValue.toIndex ?? payloadValue.index ?? payloadValue.order,
  )

  if (position !== undefined) {
    layer.position = position
    fields.push('position')
  }

  if (toIndex !== undefined) {
    layer.toIndex = toIndex
    fields.push('toIndex')
  }

  return fields.length > 0
    ? {
        fields,
        format: PPT_OBJECT_LAYER_JSON_IMPORT_FORMAT,
        jsonLength,
        layer,
      }
    : null
}

function getPPTObjectLayerPayloadValue(
  value: unknown,
  allowDirect: boolean,
): unknown {
  if (!isPPTRecord(value)) {
    return null
  }

  if (isPPTRecord(value.objectLayer)) {
    return value.objectLayer
  }

  if (isPPTRecord(value.layerOrder)) {
    return value.layerOrder
  }

  if (isPPTRecord(value.zOrder)) {
    return value.zOrder
  }

  if (
    value.position !== undefined ||
    value.arrange !== undefined ||
    value.toIndex !== undefined ||
    value.index !== undefined ||
    value.order !== undefined
  ) {
    return value
  }

  return allowDirect ? value : null
}

function getPPTObjectLayerPositionFromJSONValue(
  value: unknown,
): PPTObjectLayerImportPosition | undefined {
  if (typeof value !== 'string') {
    return undefined
  }

  switch (value.trim().toLowerCase().replace(/[\s_-]+/g, '')) {
    case 'back':
    case 'sendtoback':
      return 'back'
    case 'backward':
    case 'sendbackward':
      return 'backward'
    case 'forward':
    case 'bringforward':
      return 'forward'
    case 'front':
    case 'bringtofront':
      return 'front'
    default:
      return undefined
  }
}

function getPPTObjectLayerToIndexFromJSONValue(value: unknown) {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    return undefined
  }

  return Math.trunc(value)
}

function getPPTObjectTransformSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  if (!dataTransfer) {
    return null
  }

  const candidates: Array<{
    allowDirect: boolean
    text: string
  }> = [
    {
      allowDirect: true,
      text: dataTransfer.getData(PPT_OBJECT_TRANSFORM_JSON_MIME_TYPE),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('application/json'),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('text/json'),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('text/plain'),
    },
  ]
  const seen = new Set<string>()

  for (const candidate of candidates) {
    const text = candidate.text.trim()

    if (!text || seen.has(text)) {
      continue
    }

    seen.add(text)

    const source = getPPTObjectTransformSourceFromText(
      text,
      candidate.allowDirect,
    )

    if (source) {
      return source
    }
  }

  return null
}

function getPPTObjectTransformSourceFromText(
  text: string,
  allowDirect: boolean,
): PPTObjectTransformImportSource | null {
  const json = getPPTImportJSONText(text)

  if (!json) {
    return null
  }

  try {
    return getPPTObjectTransformSourceFromJSONValue(
      JSON.parse(json),
      json.length,
      allowDirect,
    )
  } catch {
    return null
  }
}

function getPPTObjectTransformSourceFromJSONValue(
  value: unknown,
  jsonLength: number,
  allowDirect: boolean,
): PPTObjectTransformImportSource | null {
  const payloadValue = isPPTRecord(value) &&
    isPPTRecord(value.objectTransform)
    ? value.objectTransform
    : isPPTRecord(value) && isPPTRecord(value.objectGeometry)
      ? value.objectGeometry
      : isPPTRecord(value) && isPPTRecord(value.transform)
        ? value.transform
        : isPPTRecord(value) && isPPTRecord(value.geometry)
          ? value.geometry
          : allowDirect
            ? value
            : null

  if (!isPPTRecord(payloadValue)) {
    return null
  }

  const transform: PPTObjectTransformImportSource['transform'] = {}
  const fields: PPTObjectTransformImportField[] = []
  const x = getPPTObjectTransformNumberFromJSONValue(payloadValue.x)
  const y = getPPTObjectTransformNumberFromJSONValue(payloadValue.y)
  const w = getPPTObjectTransformNumberFromJSONValue(
    payloadValue.w ?? payloadValue.width,
  )
  const h = getPPTObjectTransformNumberFromJSONValue(
    payloadValue.h ?? payloadValue.height,
  )
  const rotation = getPPTObjectTransformRotationFromJSONValue(
    payloadValue.rotation ?? payloadValue.rotate,
  )

  if (x !== undefined) {
    transform.x = x
    fields.push('x')
  }

  if (y !== undefined) {
    transform.y = y
    fields.push('y')
  }

  if (w !== undefined) {
    transform.w = w
    fields.push('w')
  }

  if (h !== undefined) {
    transform.h = h
    fields.push('h')
  }

  if (rotation !== undefined) {
    transform.rotation = rotation
    fields.push('rotation')
  }

  return fields.length > 0
    ? {
        fields,
        format: PPT_OBJECT_TRANSFORM_JSON_IMPORT_FORMAT,
        jsonLength,
        transform,
      }
    : null
}

function getPPTObjectTransformNumberFromJSONValue(value: unknown) {
  return typeof value === 'number' && Number.isFinite(value)
    ? value
    : undefined
}

function getPPTObjectTransformRotationFromJSONValue(value: unknown) {
  return typeof value === 'number' && Number.isFinite(value)
    ? normalizePPTCanvasRotationDegrees(value)
    : undefined
}

function getPPTImageReplaceSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  if (!dataTransfer) {
    return null
  }

  const candidates: Array<{
    allowDirect: boolean
    text: string
  }> = [
    {
      allowDirect: true,
      text: dataTransfer.getData(PPT_IMAGE_REPLACE_JSON_MIME_TYPE),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('application/json'),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('text/json'),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('text/plain'),
    },
  ]
  const seen = new Set<string>()

  for (const candidate of candidates) {
    const text = candidate.text.trim()

    if (!text || seen.has(text)) {
      continue
    }

    seen.add(text)

    const source = getPPTImageReplaceSourceFromText(
      text,
      candidate.allowDirect,
    )

    if (source) {
      return source
    }
  }

  return null
}

function getPPTImageReplaceSourceFromText(
  text: string,
  allowDirect: boolean,
): PPTImageReplaceImportSource | null {
  const json = getPPTImportJSONText(text)

  if (!json) {
    return null
  }

  try {
    return getPPTImageReplaceSourceFromJSONValue(
      JSON.parse(json),
      json.length,
      allowDirect,
    )
  } catch {
    return null
  }
}

function getPPTImageReplaceSourceFromJSONValue(
  value: unknown,
  jsonLength: number,
  allowDirect: boolean,
): PPTImageReplaceImportSource | null {
  const payloadValue = getPPTImageReplacePayloadValue(value, allowDirect)

  if (!isPPTRecord(payloadValue)) {
    return null
  }

  const dataUrl = getPPTImageReplaceDataUrlFromJSONValue(
    payloadValue.src ??
      payloadValue.dataUrl ??
      payloadValue.url,
  )

  if (!dataUrl) {
    return null
  }

  const mimeType = getPPTImageReplaceMimeTypeFromJSONValue(
    payloadValue.mimeType ??
      payloadValue.type,
    dataUrl,
  )

  if (!mimeType) {
    return null
  }

  const name = getPPTImageReplaceTextFromJSONValue(
    payloadValue.name ??
      payloadValue.fileName,
  )
  const altText = getPPTImageReplaceTextFromJSONValue(
    payloadValue.altText ??
      payloadValue.alt,
  )
  const naturalWidth = getPPTImageReplaceNaturalSizeFromJSONValue(
    payloadValue.naturalWidth ??
      payloadValue.width ??
      (isPPTRecord(payloadValue.naturalSize)
        ? payloadValue.naturalSize.width
        : undefined),
  )
  const naturalHeight = getPPTImageReplaceNaturalSizeFromJSONValue(
    payloadValue.naturalHeight ??
      payloadValue.height ??
      (isPPTRecord(payloadValue.naturalSize)
        ? payloadValue.naturalSize.height
        : undefined),
  )
  const fields: PPTImageReplaceImportField[] = [
    'src',
    'mimeType',
  ]

  if (name !== undefined) {
    fields.push('name')
  }

  if (altText !== undefined) {
    fields.push('altText')
  }

  if (naturalWidth !== undefined) {
    fields.push('naturalWidth')
  }

  if (naturalHeight !== undefined) {
    fields.push('naturalHeight')
  }

  return {
    fields,
    format: PPT_IMAGE_REPLACE_JSON_IMPORT_FORMAT,
    image: {
      ...(altText === undefined ? {} : { altText }),
      dataUrl,
      mimeType,
      ...(name === undefined ? {} : { name }),
      ...(naturalHeight === undefined ? {} : { naturalHeight }),
      ...(naturalWidth === undefined ? {} : { naturalWidth }),
    },
    jsonLength,
    resolveNaturalSize: naturalWidth === undefined || naturalHeight === undefined,
  }
}

function getPPTImageReplacePayloadValue(
  value: unknown,
  allowDirect: boolean,
): unknown {
  if (!isPPTRecord(value)) {
    return null
  }

  if (isPPTRecord(value.imageReplace)) {
    return value.imageReplace
  }

  if (isPPTRecord(value.imageSource)) {
    return value.imageSource
  }

  if (isPPTRecord(value.objectImage)) {
    return value.objectImage
  }

  if (isPPTRecord(value.replacementImage)) {
    return value.replacementImage
  }

  if (isPPTRecord(value.image)) {
    return value.image
  }

  return allowDirect ? value : null
}

function getPPTImageReplaceDataUrlFromJSONValue(value: unknown) {
  if (typeof value !== 'string') {
    return null
  }

  const dataUrl = value.trim()

  return /^data:image\/[a-z0-9.+-]+[;,]/i.test(dataUrl)
    ? dataUrl
    : null
}

function getPPTImageReplaceMimeTypeFromJSONValue(
  value: unknown,
  dataUrl: string,
) {
  const rawMimeType = typeof value === 'string'
    ? value
    : dataUrl.match(/^data:([^;,]+)/i)?.[1]
  const mimeType = rawMimeType?.trim().toLowerCase() ?? ''

  return /^image\/[a-z0-9.+-]+$/i.test(mimeType)
    ? mimeType
    : null
}

function getPPTImageReplaceTextFromJSONValue(value: unknown) {
  if (typeof value !== 'string') {
    return undefined
  }

  const text = value.replace(/\s+/g, ' ').trim()

  return text || undefined
}

function getPPTImageReplaceNaturalSizeFromJSONValue(value: unknown) {
  if (typeof value !== 'number' || !Number.isFinite(value) || value <= 0) {
    return undefined
  }

  return Math.round(value)
}

function getPPTImageCropSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  if (!dataTransfer) {
    return null
  }

  const candidates: Array<{
    allowDirect: boolean
    text: string
  }> = [
    {
      allowDirect: true,
      text: dataTransfer.getData(PPT_IMAGE_CROP_JSON_MIME_TYPE),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('application/json'),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('text/json'),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('text/plain'),
    },
  ]
  const seen = new Set<string>()

  for (const candidate of candidates) {
    const text = candidate.text.trim()

    if (!text || seen.has(text)) {
      continue
    }

    seen.add(text)

    const source = getPPTImageCropSourceFromText(
      text,
      candidate.allowDirect,
    )

    if (source) {
      return source
    }
  }

  return null
}

function getPPTImageCropSourceFromText(
  text: string,
  allowDirect: boolean,
): PPTImageCropImportSource | null {
  const json = getPPTImportJSONText(text)

  if (!json) {
    return null
  }

  try {
    return getPPTImageCropSourceFromJSONValue(
      JSON.parse(json),
      json.length,
      allowDirect,
    )
  } catch {
    return null
  }
}

function getPPTImageCropSourceFromJSONValue(
  value: unknown,
  jsonLength: number,
  allowDirect: boolean,
): PPTImageCropImportSource | null {
  const payloadValue = isPPTRecord(value) &&
    isPPTRecord(value.imageCrop)
    ? value.imageCrop
    : allowDirect
      ? value
      : null

  if (!isPPTRecord(payloadValue)) {
    return null
  }

  const imageCrop: PPTImageCropImportSource['imageCrop'] = {}
  const crop: Partial<PPTImageCrop> = {}
  const fields: PPTImageCropImportField[] = []
  const fit = getPPTImageCropFitFromJSONValue(payloadValue.fit)
  const cropValue = isPPTRecord(payloadValue.crop) ? payloadValue.crop : payloadValue
  const x = getPPTImageCropPositionFromJSONValue(cropValue.x)
  const y = getPPTImageCropPositionFromJSONValue(cropValue.y)

  if (fit !== undefined) {
    imageCrop.fit = fit
    fields.push('fit')
  }

  if (x !== undefined) {
    crop.x = x
    fields.push('x')
  }

  if (y !== undefined) {
    crop.y = y
    fields.push('y')
  }

  if (Object.keys(crop).length > 0) {
    imageCrop.crop = crop
  }

  return fields.length > 0
    ? {
        fields,
        format: PPT_IMAGE_CROP_JSON_IMPORT_FORMAT,
        imageCrop,
        jsonLength,
      }
    : null
}

function getPPTImageCropFitFromJSONValue(
  value: unknown,
): PPTImageFit | undefined {
  const fit = typeof value === 'string' ? value.trim() : ''

  return fit === 'contain' || fit === 'cover' ? fit : undefined
}

function getPPTImageCropPositionFromJSONValue(value: unknown) {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    return undefined
  }

  return normalizeSlideEditObjectImageCropValue(value)
}

function getPPTShapeStyleSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  if (!dataTransfer) {
    return null
  }

  const candidates: Array<{
    allowDirect: boolean
    text: string
  }> = [
    {
      allowDirect: true,
      text: dataTransfer.getData(PPT_SHAPE_STYLE_JSON_MIME_TYPE),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('application/json'),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('text/json'),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('text/plain'),
    },
  ]
  const seen = new Set<string>()

  for (const candidate of candidates) {
    const text = candidate.text.trim()

    if (!text || seen.has(text)) {
      continue
    }

    seen.add(text)

    const source = getPPTShapeStyleSourceFromText(
      text,
      candidate.allowDirect,
    )

    if (source) {
      return source
    }
  }

  return null
}

function getPPTShapeStyleSourceFromText(
  text: string,
  allowDirect: boolean,
): PPTShapeStyleImportSource | null {
  const json = getPPTImportJSONText(text)

  if (!json) {
    return null
  }

  try {
    return getPPTShapeStyleSourceFromJSONValue(
      JSON.parse(json),
      json.length,
      allowDirect,
    )
  } catch {
    return null
  }
}

function getPPTShapeStyleSourceFromJSONValue(
  value: unknown,
  jsonLength: number,
  allowDirect: boolean,
): PPTShapeStyleImportSource | null {
  const payloadValue = isPPTRecord(value) &&
    isPPTRecord(value.shapeStyle)
    ? value.shapeStyle
    : allowDirect
      ? value
      : null

  if (!isPPTRecord(payloadValue)) {
    return null
  }

  const shape: PPTShapeStyleImportSource['shape'] = {}
  const fields: PPTShapeStyleImportField[] = []
  const fill = getPPTShapeStyleFillFromJSONValue(
    payloadValue.fill ?? {
      color: payloadValue.fillColor,
      opacity: payloadValue.fillOpacity,
    },
  )
  const stroke = getPPTShapeStyleStrokeFromJSONValue(
    payloadValue.stroke ?? {
      color: payloadValue.strokeColor,
      dash: payloadValue.strokeDash,
      width: payloadValue.strokeWidth,
    },
  )
  const cornerRadius = getPPTShapeStyleCornerRadiusFromJSONValue(
    payloadValue.cornerRadius,
  )

  if (fill !== undefined) {
    shape.fill = fill
    fields.push('fill')
  }

  if (stroke !== undefined) {
    shape.stroke = stroke
    fields.push('stroke')
  }

  if (cornerRadius !== undefined) {
    shape.cornerRadius = cornerRadius
    fields.push('cornerRadius')
  }

  return fields.length > 0
    ? {
        fields,
        format: PPT_SHAPE_STYLE_JSON_IMPORT_FORMAT,
        jsonLength,
        shape,
      }
    : null
}

function getPPTShapeStyleFillFromJSONValue(
  value: unknown,
): PPTFill | undefined {
  if (typeof value === 'string') {
    const color = normalizePPTSwatchColor(value)

    return color ? normalizePPTFill({ color }) : undefined
  }

  if (!isPPTRecord(value)) {
    return undefined
  }

  const color = typeof value.color === 'string'
    ? normalizePPTSwatchColor(value.color)
    : ''
  const opacity = typeof value.opacity === 'number' &&
    Number.isFinite(value.opacity)
    ? normalizePPTFillOpacity(value.opacity)
    : undefined

  return color
    ? normalizePPTFill({ color, opacity })
    : undefined
}

function getPPTShapeStyleStrokeFromJSONValue(
  value: unknown,
): PPTStroke | undefined {
  if (!isPPTRecord(value)) {
    return undefined
  }

  const color = typeof value.color === 'string'
    ? normalizePPTSwatchColor(value.color)
    : ''
  const width = typeof value.width === 'number' && Number.isFinite(value.width)
    ? normalizePPTStrokeWidth(value.width)
    : undefined
  const dashValue = typeof value.dash === 'string' ? value.dash : ''
  const dash = isPPTStrokeDash(dashValue) ? normalizePPTStrokeDash(dashValue) : undefined

  return color || width !== undefined || dash
    ? normalizePPTStroke({
        ...(color ? { color } : {}),
        ...(dash ? { dash } : {}),
        ...(width !== undefined ? { width } : {}),
      })
    : undefined
}

function getPPTShapeStyleCornerRadiusFromJSONValue(value: unknown) {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    return undefined
  }

  return normalizePPTShapeCornerRadius(value)
}

function getPPTLineStyleSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  if (!dataTransfer) {
    return null
  }

  const candidates: Array<{
    allowDirect: boolean
    text: string
  }> = [
    {
      allowDirect: true,
      text: dataTransfer.getData(PPT_LINE_STYLE_JSON_MIME_TYPE),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('application/json'),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('text/json'),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('text/plain'),
    },
  ]
  const seen = new Set<string>()

  for (const candidate of candidates) {
    const text = candidate.text.trim()

    if (!text || seen.has(text)) {
      continue
    }

    seen.add(text)

    const source = getPPTLineStyleSourceFromText(
      text,
      candidate.allowDirect,
    )

    if (source) {
      return source
    }
  }

  return null
}

function getPPTLineStyleSourceFromText(
  text: string,
  allowDirect: boolean,
): PPTLineStyleImportSource | null {
  const json = getPPTImportJSONText(text)

  if (!json) {
    return null
  }

  try {
    return getPPTLineStyleSourceFromJSONValue(
      JSON.parse(json),
      json.length,
      allowDirect,
    )
  } catch {
    return null
  }
}

function getPPTLineStyleSourceFromJSONValue(
  value: unknown,
  jsonLength: number,
  allowDirect: boolean,
): PPTLineStyleImportSource | null {
  const payloadValue = isPPTRecord(value) &&
    isPPTRecord(value.lineStyle)
    ? value.lineStyle
    : isPPTRecord(value) && isPPTRecord(value.stroke)
      ? value.stroke
      : allowDirect
        ? value
        : null

  if (!isPPTRecord(payloadValue)) {
    return null
  }

  const strokeValue = isPPTRecord(payloadValue.stroke)
    ? payloadValue.stroke
    : payloadValue
  const stroke: Partial<PPTStroke> = {}
  const fields: PPTLineStyleImportField[] = []
  const color = getPPTLineStyleColorFromJSONValue(
    strokeValue.color ?? strokeValue.strokeColor,
  )
  const width = getPPTLineStyleWidthFromJSONValue(
    strokeValue.width ?? strokeValue.strokeWidth,
  )
  const dash = getPPTLineStyleDashFromJSONValue(
    strokeValue.dash ?? strokeValue.strokeDash,
  )

  if (color !== undefined) {
    stroke.color = color
    fields.push('color')
  }

  if (width !== undefined) {
    stroke.width = width
    fields.push('width')
  }

  if (dash !== undefined) {
    stroke.dash = dash
    fields.push('dash')
  }

  return fields.length > 0
    ? {
        fields,
        format: PPT_LINE_STYLE_JSON_IMPORT_FORMAT,
        jsonLength,
        stroke,
      }
    : null
}

function getPPTLineStyleColorFromJSONValue(value: unknown) {
  if (typeof value !== 'string') {
    return undefined
  }

  return normalizePPTSwatchColor(value) || undefined
}

function getPPTLineStyleWidthFromJSONValue(value: unknown) {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    return undefined
  }

  return normalizePPTStrokeWidth(value)
}

function getPPTLineStyleDashFromJSONValue(
  value: unknown,
): PPTStrokeDash | undefined {
  return typeof value === 'string' && isPPTStrokeDash(value)
    ? normalizePPTStrokeDash(value)
    : undefined
}

function getPPTTextStyleSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  if (!dataTransfer) {
    return null
  }

  const candidates: Array<{
    allowDirect: boolean
    text: string
  }> = [
    {
      allowDirect: true,
      text: dataTransfer.getData(PPT_TEXT_STYLE_JSON_MIME_TYPE),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('application/json'),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('text/json'),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('text/plain'),
    },
  ]
  const seen = new Set<string>()

  for (const candidate of candidates) {
    const text = candidate.text.trim()

    if (!text || seen.has(text)) {
      continue
    }

    seen.add(text)

    const source = getPPTTextStyleSourceFromText(
      text,
      candidate.allowDirect,
    )

    if (source) {
      return source
    }
  }

  return null
}

function getPPTTextStyleSourceFromText(
  text: string,
  allowDirect: boolean,
): PPTTextStyleImportSource | null {
  const json = getPPTImportJSONText(text)

  if (!json) {
    return null
  }

  try {
    return getPPTTextStyleSourceFromJSONValue(
      JSON.parse(json),
      json.length,
      allowDirect,
    )
  } catch {
    return null
  }
}

function getPPTTextStyleSourceFromJSONValue(
  value: unknown,
  jsonLength: number,
  allowDirect: boolean,
): PPTTextStyleImportSource | null {
  const payloadValue = isPPTRecord(value) &&
    isPPTRecord(value.textStyle)
    ? value.textStyle
    : allowDirect
      ? value
      : null

  if (!isPPTRecord(payloadValue)) {
    return null
  }

  const text: PPTTextStyleImportText = {}
  const paragraph: PPTTextStyleImportParagraph = {}
  const fields: PPTTextStyleImportField[] = []
  const paragraphValue = isPPTRecord(payloadValue.paragraph)
    ? payloadValue.paragraph
    : isPPTRecord(payloadValue.paragraphStyle)
      ? payloadValue.paragraphStyle
      : payloadValue
  const color = getPPTTextStyleColorFromJSONValue(payloadValue.color)
  const fontSize = getPPTTextStyleFontSizeFromJSONValue(
    payloadValue.fontSize ?? payloadValue.size,
  )
  const fontFamily = getPPTTextStyleFontFamilyFromJSONValue(
    payloadValue.fontFamily ?? payloadValue.font,
  )
  const fontWeight = getPPTTextStyleFontWeightFromJSONValue(
    payloadValue.fontWeight ?? payloadValue.weight,
  ) ?? getPPTTextStyleBoldFromJSONValue(payloadValue.bold)
  const verticalAlign = getPPTTextStyleVerticalAlignFromJSONValue(
    payloadValue.verticalAlign,
  )
  const textInset = getPPTTextStyleInsetFromJSONValue(
    payloadValue.textInset ?? payloadValue.inset,
  )
  const align = getPPTTextStyleParagraphAlignFromJSONValue(
    paragraphValue.align,
  )
  const bullet = getPPTTextStyleParagraphBulletFromJSONValue(
    paragraphValue.bullet ?? paragraphValue.list,
  )
  const lineHeight = getPPTTextStyleParagraphLineHeightFromJSONValue(
    paragraphValue.lineHeight,
  )
  const spacingBefore = getPPTTextStyleParagraphSpacingFromJSONValue(
    paragraphValue.spacingBefore,
  )
  const spacingAfter = getPPTTextStyleParagraphSpacingFromJSONValue(
    paragraphValue.spacingAfter,
  )

  if (color !== undefined) {
    text.color = color
    fields.push('color')
  }

  if (fontSize !== undefined) {
    text.fontSize = fontSize
    fields.push('fontSize')
  }

  if (fontFamily !== undefined) {
    text.fontFamily = fontFamily
    fields.push('fontFamily')
  }

  if (fontWeight !== undefined) {
    text.fontWeight = fontWeight
    fields.push('fontWeight')
  }

  if (verticalAlign !== undefined) {
    text.verticalAlign = verticalAlign
    fields.push('verticalAlign')
  }

  if (textInset !== undefined) {
    text.textInset = textInset
    fields.push('textInset')
  }

  if (align !== undefined) {
    paragraph.align = align
    fields.push('paragraphAlign')
  }

  if (bullet !== undefined) {
    paragraph.bullet = bullet
    fields.push('paragraphBullet')
  }

  if (lineHeight !== undefined) {
    paragraph.lineHeight = lineHeight
    fields.push('paragraphLineHeight')
  }

  if (spacingBefore !== undefined) {
    paragraph.spacingBefore = spacingBefore
    fields.push('paragraphSpacingBefore')
  }

  if (spacingAfter !== undefined) {
    paragraph.spacingAfter = spacingAfter
    fields.push('paragraphSpacingAfter')
  }

  return fields.length > 0
    ? {
        fields,
        format: PPT_TEXT_STYLE_JSON_IMPORT_FORMAT,
        jsonLength,
        ...(Object.keys(paragraph).length > 0 ? { paragraph } : {}),
        ...(Object.keys(text).length > 0 ? { text } : {}),
      }
    : null
}

function getPPTTextStyleColorFromJSONValue(value: unknown) {
  if (typeof value !== 'string') {
    return undefined
  }

  return normalizePPTSwatchColor(value) || undefined
}

function getPPTTextStyleFontSizeFromJSONValue(value: unknown) {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    return undefined
  }

  return clampPPTCanvasValue(
    value,
    PPT_TEXT_FONT_SIZE_MIN,
    PPT_TEXT_FONT_SIZE_MAX,
  )
}

function getPPTTextStyleFontFamilyFromJSONValue(value: unknown) {
  if (typeof value !== 'string' || !value.trim()) {
    return undefined
  }

  return normalizePPTTextFontFamily(value)
}

function getPPTTextStyleFontWeightFromJSONValue(
  value: unknown,
): PPTTextStyle['fontWeight'] | undefined {
  return value === 'regular' || value === 'semibold' || value === 'bold'
    ? value
    : undefined
}

function getPPTTextStyleBoldFromJSONValue(
  value: unknown,
): PPTTextStyle['fontWeight'] | undefined {
  if (typeof value !== 'boolean') {
    return undefined
  }

  return value ? 'bold' : 'regular'
}

function getPPTTextStyleVerticalAlignFromJSONValue(
  value: unknown,
): PPTTextVerticalAlign | undefined {
  return value === 'top' || value === 'middle' || value === 'bottom'
    ? value
    : undefined
}

function getPPTTextStyleInsetFromJSONValue(
  value: unknown,
): Partial<NonNullable<PPTTextStyle['textInset']>> | undefined {
  if (!isPPTRecord(value)) {
    return undefined
  }

  const inset: Partial<NonNullable<PPTTextStyle['textInset']>> = {}

  for (const field of ['top', 'right', 'bottom', 'left'] as const) {
    const fieldValue = value[field]

    if (typeof fieldValue === 'number' && Number.isFinite(fieldValue)) {
      inset[field] = normalizePPTTextInset(fieldValue)
    }
  }

  return Object.keys(inset).length > 0 ? inset : undefined
}

function getPPTTextStyleParagraphAlignFromJSONValue(
  value: unknown,
): PPTParagraph['align'] | undefined {
  return value === 'left' || value === 'center' || value === 'right'
    ? value
    : undefined
}

function getPPTTextStyleParagraphBulletFromJSONValue(
  value: unknown,
): PPTParagraph['bullet'] | null | undefined {
  if (value === null || value === false || value === 'none') {
    return null
  }

  if (value === true) {
    return 'bullet'
  }

  return value === 'bullet' || value === 'numbered'
    ? value
    : undefined
}

function getPPTTextStyleParagraphLineHeightFromJSONValue(value: unknown) {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    return undefined
  }

  return normalizePPTParagraphLineHeight(value)
}

function getPPTTextStyleParagraphSpacingFromJSONValue(value: unknown) {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    return undefined
  }

  return normalizePPTParagraphSpacing(value)
}

function getPPTTextBodySourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  if (!dataTransfer) {
    return null
  }

  const candidates: Array<{
    allowDirect: boolean
    text: string
  }> = [
    {
      allowDirect: true,
      text: dataTransfer.getData(PPT_TEXT_BODY_JSON_MIME_TYPE),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('application/json'),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('text/json'),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('text/plain'),
    },
  ]
  const seen = new Set<string>()

  for (const candidate of candidates) {
    const text = candidate.text.trim()

    if (!text || seen.has(text)) {
      continue
    }

    seen.add(text)

    const source = getPPTTextBodySourceFromText(
      text,
      candidate.allowDirect,
    )

    if (source) {
      return source
    }
  }

  return null
}

function getPPTTextBodySourceFromText(
  text: string,
  allowDirect: boolean,
): PPTTextBodyImportSource | null {
  const json = getPPTImportJSONText(text)

  if (!json) {
    return null
  }

  try {
    return getPPTTextBodySourceFromJSONValue(
      JSON.parse(json),
      json.length,
      allowDirect,
    )
  } catch {
    return null
  }
}

function getPPTTextBodySourceFromJSONValue(
  value: unknown,
  jsonLength: number,
  allowDirect: boolean,
): PPTTextBodyImportSource | null {
  const payloadValue = getPPTTextBodyPayloadValue(value, allowDirect)

  if (typeof payloadValue === 'string') {
    const text = payloadValue.replace(/\r\n?/g, '\n')

    return text
      ? {
          format: PPT_TEXT_BODY_JSON_IMPORT_FORMAT,
          jsonLength,
          mode: 'plain-text',
          textBody: createPPTTextBody(text),
        }
      : null
  }

  const textBody = getPPTTextBodyFromJSONValue(payloadValue)

  return textBody
    ? {
        format: PPT_TEXT_BODY_JSON_IMPORT_FORMAT,
        jsonLength,
        mode: 'text-body',
        textBody,
      }
    : null
}

function getPPTTextBodyPayloadValue(
  value: unknown,
  allowDirect: boolean,
): unknown {
  if (!isPPTRecord(value)) {
    return allowDirect ? value : null
  }

  if (isPPTRecord(value.textBody) || typeof value.textBody === 'string') {
    return value.textBody
  }

  if (isPPTRecord(value.body) || typeof value.body === 'string') {
    return value.body
  }

  if (isPPTRecord(value.content) || typeof value.content === 'string') {
    return value.content
  }

  if (typeof value.text === 'string') {
    return value.text
  }

  if (typeof value.plainText === 'string') {
    return value.plainText
  }

  return allowDirect ? value : null
}

function getPPTTextAutoFitSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  if (!dataTransfer) {
    return null
  }

  const candidates: Array<{
    allowDirect: boolean
    text: string
  }> = [
    {
      allowDirect: true,
      text: dataTransfer.getData(PPT_TEXT_AUTOFIT_JSON_MIME_TYPE),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('application/json'),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('text/json'),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('text/plain'),
    },
  ]
  const seen = new Set<string>()

  for (const candidate of candidates) {
    const text = candidate.text.trim()

    if (!text || seen.has(text)) {
      continue
    }

    seen.add(text)

    const source = getPPTTextAutoFitSourceFromText(
      text,
      candidate.allowDirect,
    )

    if (source) {
      return source
    }
  }

  return null
}

function getPPTTextAutoFitSourceFromText(
  text: string,
  allowDirect: boolean,
): PPTTextAutoFitImportSource | null {
  const json = getPPTImportJSONText(text)

  if (!json) {
    return null
  }

  try {
    return getPPTTextAutoFitSourceFromJSONValue(
      JSON.parse(json),
      json.length,
      allowDirect,
    )
  } catch {
    return null
  }
}

function getPPTTextAutoFitSourceFromJSONValue(
  value: unknown,
  jsonLength: number,
  allowDirect: boolean,
): PPTTextAutoFitImportSource | null {
  const payloadValue = getPPTTextAutoFitPayloadValue(value, allowDirect)
  const mode = getPPTTextAutoFitModeFromJSONValue(payloadValue)

  if (mode === null) {
    return null
  }

  const rawHandle = isPPTRecord(payloadValue)
    ? payloadValue.handle ?? payloadValue.resizeHandle
    : undefined
  const handle = getPPTTextAutoFitHandleFromJSONValue(rawHandle)

  if (rawHandle !== undefined && handle === undefined) {
    return null
  }

  return {
    fields: handle === undefined ? ['mode'] : ['mode', 'handle'],
    format: PPT_TEXT_AUTOFIT_JSON_IMPORT_FORMAT,
    handle: handle ?? 'se',
    jsonLength,
    mode,
  }
}

function getPPTTextAutoFitPayloadValue(
  value: unknown,
  allowDirect: boolean,
): unknown {
  if (!isPPTRecord(value)) {
    return allowDirect ? value : null
  }

  if (
    isPPTRecord(value.textAutoFit) ||
    typeof value.textAutoFit === 'boolean' ||
    typeof value.textAutoFit === 'string'
  ) {
    return value.textAutoFit
  }

  if (
    isPPTRecord(value.autoFit) ||
    typeof value.autoFit === 'boolean' ||
    typeof value.autoFit === 'string'
  ) {
    return value.autoFit
  }

  if (
    isPPTRecord(value.textBoxAutoFit) ||
    typeof value.textBoxAutoFit === 'boolean' ||
    typeof value.textBoxAutoFit === 'string'
  ) {
    return value.textBoxAutoFit
  }

  if (
    isPPTRecord(value.textOverflow) ||
    typeof value.textOverflow === 'boolean' ||
    typeof value.textOverflow === 'string'
  ) {
    return value.textOverflow
  }

  return allowDirect ? value : null
}

function getPPTTextAutoFitModeFromJSONValue(
  value: unknown,
): 'resize-to-fit' | null {
  if (value === true) {
    return 'resize-to-fit'
  }

  if (isPPTRecord(value)) {
    const modeValue = value.mode ??
      value.sizeMode ??
      value.textAutoFit ??
      value.autoFit ??
      value.textOverflow

    return getPPTTextAutoFitModeFromJSONValue(
      modeValue,
    )
  }

  if (typeof value !== 'string') {
    return null
  }

  const mode = value.replace(/[\s_-]/g, '').toLowerCase()

  return mode === 'auto' ||
    mode === 'fit' ||
    mode === 'resize' ||
    mode === 'resizetofit' ||
    mode === 'resizeshapetofittext'
    ? 'resize-to-fit'
    : null
}

function getPPTTextAutoFitHandleFromJSONValue(
  value: unknown,
): ResizeHandle | undefined {
  if (typeof value !== 'string') {
    return undefined
  }

  const handle = value.trim().toLowerCase()

  return PPT_RESIZE_HANDLES.includes(handle as ResizeHandle)
    ? handle as ResizeHandle
    : undefined
}

function getPPTTextBodyFromJSONValue(value: unknown): PPTTextBody | null {
  const parsed = PPTTextBodySchema.safeParse(value)

  if (!parsed.success) {
    return null
  }

  return normalizePPTImportedTextBody(parsed.data)
}

function normalizePPTImportedTextBody(body: PPTTextBody): PPTTextBody {
  const paragraphs = body.paragraphs.length > 0
    ? body.paragraphs
    : [{ runs: [{ text: '' }] }]

  return {
    paragraphs: paragraphs.map((paragraph) => ({
      ...paragraph,
      runs: paragraph.runs.length > 0
        ? paragraph.runs.map((run) => ({ ...run }))
        : [{ text: '' }],
    })),
  }
}

function getPPTTableRowsSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  if (!dataTransfer) {
    return null
  }

  const candidates: Array<{
    allowDirect: boolean
    text: string
  }> = [
    {
      allowDirect: true,
      text: dataTransfer.getData(PPT_TABLE_ROWS_JSON_MIME_TYPE),
    },
    {
      allowDirect: true,
      text: dataTransfer.getData(PPT_TABLE_CLIPBOARD_JSON_MIME_TYPE),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('application/json'),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('text/json'),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('text/plain'),
    },
  ]
  const seen = new Set<string>()

  for (const candidate of candidates) {
    const text = candidate.text.trim()

    if (!text || seen.has(text)) {
      continue
    }

    seen.add(text)

    const source = getPPTTableRowsSourceFromText(
      text,
      candidate.allowDirect,
    )

    if (source) {
      return source
    }
  }

  return null
}

function getPPTTableRowsSourceFromText(
  text: string,
  allowDirect: boolean,
): PPTTableRowsImportSource | null {
  const json = getPPTImportJSONText(text)

  if (!json) {
    return null
  }

  try {
    return getPPTTableRowsSourceFromJSONValue(
      JSON.parse(json),
      json.length,
      allowDirect,
    )
  } catch {
    return null
  }
}

function getPPTTableRowsSourceFromJSONValue(
  value: unknown,
  jsonLength: number,
  allowDirect: boolean,
): PPTTableRowsImportSource | null {
  const rows = getPPTTableRowsFromJSONValue(
    getPPTTableRowsPayloadValue(value, allowDirect),
  )

  return rows
    ? {
        format: PPT_TABLE_ROWS_JSON_IMPORT_FORMAT,
        jsonLength,
        rows: normalizePPTTableRows(rows),
      }
    : null
}

function getPPTTableRowsPayloadValue(
  value: unknown,
  allowDirect: boolean,
): unknown {
  if (!isPPTRecord(value)) {
    return allowDirect ? value : null
  }

  if (Array.isArray(value.tableRows) || isPPTRecord(value.tableRows)) {
    return value.tableRows
  }

  if (isPPTRecord(value.table) && Array.isArray(value.table.rows)) {
    return value.table
  }

  if (Array.isArray(value.rows)) {
    return value
  }

  return allowDirect ? value : null
}

function getPPTTableRowsFromJSONValue(
  value: unknown,
): readonly (readonly string[])[] | null {
  if (Array.isArray(value)) {
    return getPPTTableRowsFromJSONArray(value)
  }

  if (!isPPTRecord(value)) {
    return null
  }

  const rowsValue = value.rows
  const columns = getPPTTableColumnsFromJSONValue(
    value.columns ?? value.headers,
  )
  const rows = Array.isArray(rowsValue)
    ? getPPTTableRowsFromJSONArray(rowsValue, columns)
    : null

  if (!rows) {
    return null
  }

  return columns.length > 0
    ? [columns, ...rows]
    : rows
}

function getPPTTableRowsFromJSONArray(
  value: readonly unknown[],
  columns: readonly string[] = [],
): readonly (readonly string[])[] | null {
  const rows = value
    .map((row) => getPPTTableRowFromJSONValue(row, columns))
    .filter((row): row is string[] => row !== null)

  return isPPTTableRowsJSONValue(rows) ? rows : null
}

function getPPTTableColumnsFromJSONValue(value: unknown) {
  return Array.isArray(value)
    ? value
        .map(getPPTTableCellFromJSONValue)
        .filter((cell) => cell.trim().length > 0)
    : []
}

function getPPTTableRowFromJSONValue(
  value: unknown,
  columns: readonly string[],
) {
  if (Array.isArray(value)) {
    const row = value.map(getPPTTableCellFromJSONValue)

    return row.some((cell) => cell.trim().length > 0) ? row : null
  }

  if (isPPTRecord(value)) {
    const keys = columns.length > 0 ? columns : Object.keys(value)
    const row = keys.map((key) => getPPTTableCellFromJSONValue(value[key]))

    return row.some((cell) => cell.trim().length > 0) ? row : null
  }

  return null
}

function getPPTTableCellFromJSONValue(value: unknown) {
  if (value === null || value === undefined) {
    return ''
  }

  return typeof value === 'string' ||
    typeof value === 'number' ||
    typeof value === 'boolean'
    ? String(value)
    : ''
}

function isPPTTableRowsJSONValue(
  rows: readonly (readonly string[])[],
) {
  return rows.length > 0 &&
    rows.some((row) => row.some((cell) => cell.trim().length > 0)) &&
    getPPTTableColumnCount(rows) > 0
}

function getPPTMediaJSONSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  if (!dataTransfer) {
    return null
  }

  const candidates: Array<{
    allowDirect: boolean
    text: string
  }> = [
    {
      allowDirect: true,
      text: dataTransfer.getData(PPT_MEDIA_JSON_MIME_TYPE),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('application/json'),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('text/json'),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('text/plain'),
    },
  ]
  const seen = new Set<string>()

  for (const candidate of candidates) {
    const text = candidate.text.trim()

    if (!text || seen.has(text)) {
      continue
    }

    seen.add(text)

    const source = getPPTMediaJSONSourceFromText(
      text,
      candidate.allowDirect,
    )

    if (source) {
      return source
    }
  }

  return null
}

function getPPTMediaJSONSourceFromText(
  text: string,
  allowDirect: boolean,
): PPTMediaJSONImportSource | null {
  const json = getPPTImportJSONText(text)

  if (!json) {
    return null
  }

  try {
    return getPPTMediaJSONSourceFromJSONValue(
      JSON.parse(json),
      json.length,
      allowDirect,
    )
  } catch {
    return null
  }
}

function getPPTMediaJSONSourceFromJSONValue(
  value: unknown,
  jsonLength: number,
  allowDirect: boolean,
): PPTMediaJSONImportSource | null {
  const payloadValue = getPPTMediaJSONPayloadValue(value, allowDirect)

  if (!isPPTRecord(payloadValue)) {
    return null
  }

  const urlText = getPPTMediaJSONTextFromJSONValue(
    payloadValue.url ??
      payloadValue.href ??
      payloadValue.src,
  )

  if (!urlText) {
    return null
  }

  const mediaSource = getPPTMediaSourceFromText(urlText)

  if (!mediaSource) {
    return null
  }

  const title = getPPTMediaJSONTextFromJSONValue(
    payloadValue.title ??
      payloadValue.name ??
      payloadValue.label,
  )
  const fields: PPTMediaJSONImportField[] = ['url']

  if (title !== undefined) {
    fields.push('title')
  }

  return {
    fields,
    format: PPT_MEDIA_JSON_IMPORT_FORMAT,
    jsonLength,
    source: {
      ...mediaSource,
      ...(title === undefined ? {} : { title }),
    },
  }
}

function getPPTMediaJSONPayloadValue(
  value: unknown,
  allowDirect: boolean,
): unknown {
  if (!isPPTRecord(value)) {
    return null
  }

  if (typeof value.media === 'string') {
    return { url: value.media }
  }

  if (isPPTRecord(value.media)) {
    return value.media
  }

  if (typeof value.mediaSource === 'string') {
    return { url: value.mediaSource }
  }

  if (isPPTRecord(value.mediaSource)) {
    return value.mediaSource
  }

  if (isPPTRecord(value.linkCard)) {
    return value.linkCard
  }

  if (isPPTRecord(value.linkPreview)) {
    return value.linkPreview
  }

  if (isPPTRecord(value.embed)) {
    return value.embed
  }

  return allowDirect ? value : null
}

function getPPTMediaJSONTextFromJSONValue(value: unknown) {
  if (typeof value !== 'string') {
    return undefined
  }

  const text = value.replace(/\s+/g, ' ').trim()

  return text || undefined
}

function getPPTCommentSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  if (!dataTransfer) {
    return null
  }

  const candidates: Array<{
    allowDirect: boolean
    text: string
  }> = [
    {
      allowDirect: true,
      text: dataTransfer.getData(PPT_COMMENT_JSON_MIME_TYPE),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('application/json'),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('text/json'),
    },
    {
      allowDirect: false,
      text: dataTransfer.getData('text/plain'),
    },
  ]
  const seen = new Set<string>()

  for (const candidate of candidates) {
    const text = candidate.text.trim()

    if (!text || seen.has(text)) {
      continue
    }

    seen.add(text)

    const source = getPPTCommentSourceFromText(
      text,
      candidate.allowDirect,
    )

    if (source) {
      return source
    }
  }

  return null
}

function getPPTCommentSourceFromText(
  text: string,
  allowDirect: boolean,
): PPTCommentImportSource | null {
  const json = getPPTImportJSONText(text)

  if (!json) {
    return null
  }

  try {
    return getPPTCommentSourceFromJSONValue(
      JSON.parse(json),
      json.length,
      allowDirect,
    )
  } catch {
    return null
  }
}

function getPPTCommentSourceFromJSONValue(
  value: unknown,
  jsonLength: number,
  allowDirect: boolean,
): PPTCommentImportSource | null {
  const payloadValue = getPPTCommentPayloadValue(value, allowDirect)

  if (!isPPTRecord(payloadValue)) {
    return null
  }

  const comment: PPTCommentImportSource['comment'] = {}
  const fields: PPTCommentImportField[] = []
  const body = getPPTCommentBodyFromJSONValue(
    payloadValue.body ?? payloadValue.text,
  )
  const resolved = typeof payloadValue.resolved === 'boolean'
    ? payloadValue.resolved
    : undefined
  const createdAt = getPPTCommentCreatedAtFromJSONValue(
    payloadValue.createdAt,
  )
  const thread = getPPTCommentThreadFromJSONValue(
    payloadValue.thread ?? payloadValue.messages ?? payloadValue.replies,
    body,
  )

  if (body !== undefined) {
    comment.body = body
    fields.push('body')
  }

  if (resolved !== undefined) {
    comment.resolved = resolved
    fields.push('resolved')
  }

  if (createdAt !== undefined) {
    comment.createdAt = createdAt
    fields.push('createdAt')
  }

  if (thread !== undefined) {
    comment.thread = thread
    fields.push('thread')
  }

  return fields.length > 0
    ? {
        comment,
        fields,
        format: PPT_COMMENT_JSON_IMPORT_FORMAT,
        jsonLength,
      }
    : null
}

function getPPTCommentPayloadValue(
  value: unknown,
  allowDirect: boolean,
): unknown {
  if (!isPPTRecord(value)) {
    return allowDirect ? value : null
  }

  if (isPPTRecord(value.comment)) {
    return value.comment
  }

  if (isPPTRecord(value.commentThread)) {
    return value.commentThread
  }

  if (isPPTRecord(value.reviewComment)) {
    return value.reviewComment
  }

  return allowDirect ? value : null
}

function getPPTCommentBodyFromJSONValue(value: unknown) {
  return typeof value === 'string'
    ? normalizePPTCommentBody(value.replace(/\r\n?/g, '\n'))
    : undefined
}

function getPPTCommentCreatedAtFromJSONValue(value: unknown) {
  return typeof value === 'string' && value.trim()
    ? normalizePPTCommentCreatedAt(value)
    : undefined
}

function getPPTCommentThreadFromJSONValue(
  value: unknown,
  firstBody: string | undefined,
) {
  if (!Array.isArray(value)) {
    return undefined
  }

  const thread = value
    .map((message, index) =>
      getPPTCommentThreadMessageFromJSONValue(message, index))
    .filter((message): message is PPTCommentThreadMessage => message !== null)

  if (thread.length === 0) {
    return undefined
  }

  return firstBody === undefined
    ? thread
    : syncPPTCommentThreadWithBody(thread, firstBody)
}

function getPPTCommentThreadMessageFromJSONValue(
  value: unknown,
  index: number,
): PPTCommentThreadMessage | null {
  const body = typeof value === 'string'
    ? normalizePPTCommentReplyBody(value.replace(/\r\n?/g, '\n'))
    : isPPTRecord(value)
      ? normalizePPTCommentReplyBody(
          String(value.body ?? value.text ?? '').replace(/\r\n?/g, '\n'),
        )
      : ''

  if (!body.trim()) {
    return null
  }

  return {
    authorName: isPPTRecord(value) && typeof value.authorName === 'string'
      ? normalizePPTCommentAuthorName(value.authorName)
      : PPT_COMMENT_DEFAULT_AUTHOR,
    body,
    createdAt: isPPTRecord(value) && typeof value.createdAt === 'string'
      ? normalizePPTCommentCreatedAt(value.createdAt)
      : PPT_COMMENT_DEFAULT_CREATED_AT,
    id: isPPTRecord(value) && typeof value.id === 'string' && value.id.trim()
      ? normalizePPTCommentMessageId(value.id)
      : `json-comment:message-${index + 1}`,
  }
}

function getPPTSlideNotesSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  if (!dataTransfer) {
    return null
  }

  const candidates: Array<{
    format: PPTSlideNotesImportSource['format']
    text: string
  }> = [
    {
      format: PPT_SLIDE_NOTES_JSON_IMPORT_FORMAT,
      text: dataTransfer.getData(PPT_SLIDE_NOTES_JSON_MIME_TYPE),
    },
    {
      format: PPT_SLIDE_NOTES_JSON_IMPORT_FORMAT,
      text: dataTransfer.getData('application/json'),
    },
    {
      format: PPT_SLIDE_NOTES_JSON_IMPORT_FORMAT,
      text: dataTransfer.getData('text/json'),
    },
    {
      format: PPT_SLIDE_NOTES_MARKDOWN_IMPORT_FORMAT,
      text: dataTransfer.getData('text/markdown'),
    },
    {
      format: PPT_SLIDE_NOTES_TEXT_IMPORT_FORMAT,
      text: dataTransfer.getData('text/plain'),
    },
  ]
  const seen = new Set<string>()

  for (const candidate of candidates) {
    const text = candidate.text.trim()

    if (!text || seen.has(text)) {
      continue
    }

    seen.add(text)

    const source = getPPTSlideNotesSourceFromText(text, candidate.format)

    if (source) {
      return source
    }
  }

  return null
}

function getPPTSlideNotesSourceFromText(
  text: string,
  format: PPTSlideNotesImportSource['format'],
): PPTSlideNotesImportSource | null {
  const normalized = text.replace(/\r\n?/g, '\n').trim()

  if (!normalized) {
    return null
  }

  const notes = format === PPT_SLIDE_NOTES_JSON_IMPORT_FORMAT
    ? getPPTSlideNotesFromJSONText(normalized)
    : getPPTSlideNotesFromMarkedText(normalized)

  return notes
    ? {
        format,
        notes,
        textLength: normalized.length,
      }
    : null
}

function getPPTSlideNotesFromJSONText(text: string) {
  const json = getPPTImportJSONText(text)

  if (!json) {
    return null
  }

  try {
    return getPPTSlideNotesFromJSONValue(JSON.parse(json))
  } catch {
    return null
  }
}

function getPPTSlideNotesFromJSONValue(value: unknown): string | null {
  if (typeof value === 'string') {
    return normalizePPTSlideNotesText(value)
  }

  if (!isPPTRecord(value)) {
    return null
  }

  const notesValue =
    value.notes ??
      value.speakerNotes ??
      value.speaker_notes ??
      (isPPTRecord(value.slide) ? value.slide.notes : undefined)

  return typeof notesValue === 'string'
    ? normalizePPTSlideNotesText(notesValue)
    : null
}

function getPPTSlideNotesFromMarkedText(text: string) {
  const lines = text.split('\n')
  const firstContentIndex = lines.findIndex((line) => line.trim())

  if (firstContentIndex < 0) {
    return null
  }

  const firstLine = lines[firstContentIndex].trim()
  const headingMatch = firstLine.match(
    /^#{1,6}\s*(?:speaker\s+notes?|presenter\s+notes?|notes?)\s*#*$/i,
  )
  const prefixMatch = firstLine.match(
    /^(?:speaker\s+notes?|presenter\s+notes?|notes?)\s*:\s*(.*)$/i,
  )

  if (!headingMatch && !prefixMatch) {
    return null
  }

  const noteLines = [
    ...(prefixMatch?.[1]?.trim() ? [prefixMatch[1].trim()] : []),
    ...lines.slice(firstContentIndex + 1),
  ]

  return normalizePPTSlideNotesText(noteLines.join('\n'))
}

function normalizePPTSlideNotesText(text: string) {
  const normalized = text
    .replace(/\r\n?/g, '\n')
    .split('\n')
    .map((line) => line.trim())
    .join('\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim()

  return normalized || null
}

function getPPTImportJSONText(text: string) {
  const trimmed = text.trim()

  if (!trimmed) {
    return null
  }

  if (trimmed.startsWith('{') || trimmed.startsWith('[')) {
    return trimmed
  }

  const fenced = trimmed.match(/^```(?:json|ppt|ppt-json)?\s*\n([\s\S]*?)\n```$/i) ??
    trimmed.match(/```(?:json|ppt|ppt-json)?\s*\n([\s\S]*?)\n```/i)

  return fenced?.[1]?.trim() ?? null
}

function getPPTElementsJSONSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  if (!dataTransfer) {
    return null
  }

  const candidates: Array<{
    format: PPTElementsJSONImportSource['format']
    text: string
  }> = [
    {
      format: PPT_ELEMENTS_JSON_IMPORT_FORMAT,
      text: dataTransfer.getData(PPT_ELEMENTS_JSON_MIME_TYPE),
    },
    {
      format: PPT_ELEMENTS_JSON_IMPORT_FORMAT,
      text: dataTransfer.getData('application/json'),
    },
    {
      format: PPT_ELEMENTS_JSON_TEXT_IMPORT_FORMAT,
      text: dataTransfer.getData('text/json'),
    },
    {
      format: PPT_ELEMENTS_JSON_TEXT_IMPORT_FORMAT,
      text: dataTransfer.getData('text/markdown'),
    },
    {
      format: PPT_ELEMENTS_JSON_TEXT_IMPORT_FORMAT,
      text: dataTransfer.getData('text/plain'),
    },
  ]
  const seen = new Set<string>()

  for (const candidate of candidates) {
    const text = candidate.text.trim()

    if (!text || seen.has(text)) {
      continue
    }

    seen.add(text)

    const source = getPPTElementsJSONSourceFromText(text, candidate.format)

    if (source) {
      return source
    }
  }

  return null
}

function getPPTElementsJSONSourceFromText(
  text: string,
  format: PPTElementsJSONImportSource['format'],
): PPTElementsJSONImportSource | null {
  const json = getPPTImportJSONText(text)

  if (!json) {
    return null
  }

  try {
    const source = getPPTElementsJSONSourceFromValue(JSON.parse(json))

    return source
      ? {
          ...source,
          format,
          jsonLength: json.length,
        }
      : null
  } catch {
    return null
  }
}

function getPPTElementsJSONSourceFromValue(
  value: unknown,
): Omit<PPTElementsJSONImportSource, 'format' | 'jsonLength'> | null {
  const payloadValue = isPPTRecord(value) &&
    value.kind === PPT_RICH_CLIPBOARD_KIND &&
    value.version === PPT_RICH_CLIPBOARD_VERSION
    ? value.payload
    : value
  const sourceSlideId = isPPTRecord(payloadValue) &&
    typeof payloadValue.sourceSlideId === 'string'
    ? payloadValue.sourceSlideId
    : 'ai-elements-json'
  const rawObjects = getPPTElementsJSONObjectsValue(payloadValue)
  const objects = rawObjects.flatMap((item) => {
    const parsed = PPTElementSchema.safeParse(item)

    return parsed.success ? [parsed.data] : []
  })

  if (objects.length === 0) {
    return null
  }

  const selectedObjectIds = normalizeSlideEditClipboardSelectedObjectIds({
    getObjectId: (object) => object.id,
    objects,
    selectedObjectIds: isPPTRecord(payloadValue) &&
      Array.isArray(payloadValue.selectedObjectIds)
      ? payloadValue.selectedObjectIds
      : null,
  })

  return {
    objects,
    selectedObjectIds,
    sourceSlideId,
  }
}

function getPPTSlidesFromJSONValue(value: unknown) {
  const payloadValue = isPPTRecord(value) &&
    value.kind === PPT_SLIDE_CLIPBOARD_KIND &&
    value.version === PPT_SLIDE_CLIPBOARD_VERSION
    ? value.payload
    : value

  if (PPTDeckSchema.safeParse(payloadValue).success) {
    return []
  }

  const rawSlides = getPPTSlideJSONSlidesValue(payloadValue)

  return rawSlides.flatMap((item) => {
    const parsed = PPTSlideSchema.safeParse(item)

    return parsed.success ? [parsed.data] : []
  })
}

function getPPTSlideJSONSlidesValue(value: unknown): unknown[] {
  if (Array.isArray(value)) {
    return value
  }

  const parsed = PPTSlideSchema.safeParse(value)

  if (parsed.success) {
    return [parsed.data]
  }

  if (!isPPTRecord(value)) {
    return []
  }

  if (Array.isArray(value.slides)) {
    return value.slides
  }

  return value.slide === undefined ? [] : [value.slide]
}

function getPPTElementsJSONObjectsValue(value: unknown): unknown[] {
  if (Array.isArray(value)) {
    return value
  }

  const parsed = PPTElementSchema.safeParse(value)

  if (parsed.success) {
    return [parsed.data]
  }

  if (!isPPTRecord(value)) {
    return []
  }

  if (Array.isArray(value.objects)) {
    return value.objects
  }

  if (Array.isArray(value.elements)) {
    return value.elements
  }

  return value.element === undefined ? [] : [value.element]
}

function getPPTDeckFromJSONValue(value: unknown) {
  const parsed = PPTDeckSchema.safeParse(value)

  if (parsed.success) {
    return parsed.data
  }

  if (typeof value !== 'object' || value === null) {
    return null
  }

  const wrapped = PPTDeckSchema.safeParse((value as { deck?: unknown }).deck)

  return wrapped.success ? wrapped.data : null
}

function getPPTDeckFallbackHTMLSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  const html = dataTransfer?.getData('text/html') ?? ''
  const plainText = dataTransfer?.getData('text/plain') ?? ''

  return getPPTDeckFallbackHTMLSourceFromHTML(html) ??
    getPPTDeckFallbackHTMLSourceFromHTML(plainText)
}

function getPPTDeckFallbackHTMLSourceFromHTML(
  html: string,
): PPTDeckHTMLFallbackSource | null {
  if (
    !html ||
    typeof DOMParser === 'undefined' ||
    typeof XMLSerializer === 'undefined'
  ) {
    return null
  }

  const doc = new DOMParser().parseFromString(html, 'text/html')
  const slideElements = [...doc.querySelectorAll<HTMLElement>(
    'section.ppt-slide[data-ppt-slide], section[data-ppt-slide].ppt-slide',
  )].filter((slide) => slide.querySelector('[data-ppt-element], .ppt-element'))

  if (slideElements.length === 0) {
    return null
  }

  const css = createPPTDeckFallbackHTMLSnapshotCSS(doc)
  const title = doc.querySelector('title')?.textContent?.trim() ||
    'PPT HTML Deck'
  const slides = slideElements
    .map((slide, index): PPTDeckHTMLFallbackSlideSource | null => {
      const sourceSlideId =
        slide.getAttribute('data-ppt-slide')?.trim() || undefined
      const name =
        slide.getAttribute('data-ppt-slide-name')?.trim() ||
        (sourceSlideId ? `PPT ${sourceSlideId}` : `PPT Slide ${index + 1}`)
      const svg = createPPTDeckFallbackHTMLSlideSVG({
        css,
        slide,
        sourceSlideId,
      })

      return svg
        ? {
            htmlLength: slide.outerHTML.length,
            name,
            ...(sourceSlideId ? { sourceSlideId } : {}),
            svg,
          }
        : null
    })
    .filter((slide): slide is PPTDeckHTMLFallbackSlideSource =>
      slide !== null)

  return slides.length > 0
    ? {
        htmlLength: html.length,
        slides,
        title,
      }
    : null
}

function createPPTDeckFallbackHTMLSnapshotCSS(doc: Document) {
  const css = [...doc.querySelectorAll('style')]
    .map((style) => style.textContent ?? '')
    .filter(Boolean)
    .join('\n')
  const fallbackCSS = [
    `*{box-sizing:border-box;}`,
    `body{margin:0;}`,
    `.ppt-slide{position:relative;width:${PPT_SLIDE_WIDTH}px;height:${PPT_SLIDE_HEIGHT}px;overflow:hidden;background:#fff;}`,
    '.ppt-element{position:absolute;margin:0;overflow:hidden;white-space:pre-wrap;overflow-wrap:anywhere;}',
    '.ppt-image{display:block;padding:0;}',
  ].join('')

  return [css, fallbackCSS].filter(Boolean).join('\n')
}

function createPPTDeckFallbackHTMLSlideSVG({
  css,
  slide,
  sourceSlideId,
}: {
  css: string
  slide: HTMLElement
  sourceSlideId?: string
}) {
  const clone = slide.cloneNode(true) as HTMLElement

  sanitizePPTDeckFallbackHTMLElement(clone)
  clone.setAttribute('xmlns', 'http://www.w3.org/1999/xhtml')

  const serializedSlide = new XMLSerializer().serializeToString(clone)
  const sourceSlideAttr = sourceSlideId
    ? ` data-ppt-deck-fallback-source-slide="${
      escapePPTCanvasXmlAttribute(sourceSlideId)
    }"`
    : ''

  return [
    `<svg xmlns="http://www.w3.org/2000/svg" width="${PPT_SLIDE_WIDTH}" height="${PPT_SLIDE_HEIGHT}" viewBox="0 0 ${PPT_SLIDE_WIDTH} ${PPT_SLIDE_HEIGHT}" data-ppt-deck-fallback-slide="true"${sourceSlideAttr}>`,
    `<foreignObject width="${PPT_SLIDE_WIDTH}" height="${PPT_SLIDE_HEIGHT}">`,
    '<div xmlns="http://www.w3.org/1999/xhtml">',
    `<style><![CDATA[${escapePPTDeckFallbackStyleCDATA(css)}]]></style>`,
    serializedSlide,
    '</div>',
    '</foreignObject>',
    '</svg>',
  ].join('')
}

function sanitizePPTDeckFallbackHTMLElement(root: HTMLElement) {
  root
    .querySelectorAll('script, iframe, object, embed, link, meta')
    .forEach((element) => element.remove())

  for (const element of [root, ...root.querySelectorAll('*')]) {
    for (const attribute of [...element.attributes]) {
      const name = attribute.name.toLowerCase()
      const value = attribute.value.trim().toLowerCase()

      if (
        name.startsWith('on') ||
        name === 'srcdoc' ||
        ((name === 'href' || name === 'src') && value.startsWith('javascript:'))
      ) {
        element.removeAttribute(attribute.name)
      }
    }
  }
}

function escapePPTDeckFallbackStyleCDATA(value: string) {
  return value.replaceAll(']]>', ']]]]><![CDATA[>')
}

function createPPTSlideFromFallbackHTMLSource(
  deck: PPTDeck,
  source: PPTSlideClipboardFallbackHTMLSource,
): PPTSlide {
  const id = createPPTSlideId(deck)
  const imageId = `${id}-slide-snapshot`

  return {
    background: { color: '#ffffff' },
    elements: [{
      alt: source.name,
      fit: 'contain',
      geometry: {
        h: PPT_SLIDE_HEIGHT,
        w: PPT_SLIDE_WIDTH,
        x: 0,
        y: 0,
      },
      id: imageId,
      kind: 'image',
      name: source.name,
      src: `data:image/svg+xml;charset=utf-8,${
        encodeURIComponent(source.svg)
      }`,
    }],
    id,
    name: `${source.name} Snapshot`,
    notes: '',
  }
}

function createPPTSlidesFromDeckFallbackHTMLSource(
  deck: PPTDeck,
  source: PPTDeckHTMLFallbackSource,
): PPTSlide[] {
  const createId = createPPTCanvasSequentialIdFactory({
    existingIds: deck.slides.map((slide) => slide.id),
    startIndex: deck.slides.length + 1,
  })

  return source.slides.map((slideSource) => {
    const id = createId('slide')
    const imageId = `${id}-deck-snapshot`

    return {
      background: { color: '#ffffff' },
      elements: [{
        alt: slideSource.name,
        fit: 'contain',
        geometry: {
          h: PPT_SLIDE_HEIGHT,
          w: PPT_SLIDE_WIDTH,
          x: 0,
          y: 0,
        },
        id: imageId,
        kind: 'image',
        name: slideSource.name,
        src: `data:image/svg+xml;charset=utf-8,${
          encodeURIComponent(slideSource.svg)
        }`,
      }],
      id,
      name: `${slideSource.name} Snapshot`,
      notes: '',
    }
  })
}

function createPPTRichClipboardFallback({
  payload,
  selectionSvg,
}: {
  payload: PPTClipboardPayload
  selectionSvg: string | null
}): PPTRichClipboardFallback {
  const plainText = createPPTSelectionClipboardPlainText(payload.objects)

  return {
    html: createPPTSelectionClipboardFallbackHTML(payload.objects, selectionSvg),
    plainText: plainText || 'PPT selection',
  }
}

function createPPTRichClipboardHTML({
  fallbackHTML,
  payload,
}: {
  fallbackHTML: string
  payload: PPTClipboardPayload
}) {
  return createPPTCanvasRichClipboardHTML({
    fallbackHTML,
    json: stringifyPPTRichClipboardPayload(payload),
    rootAttribute: PPT_RICH_CLIPBOARD_HTML_ROOT_ATTRIBUTE,
    scriptAttribute: PPT_RICH_CLIPBOARD_HTML_JSON_SCRIPT_ATTRIBUTE,
  })
}

function createPPTSelectionClipboardPlainText(
  objects: readonly PPTElement[],
) {
  return objects
    .map(createPPTElementClipboardPlainText)
    .map((text) => text.trim())
    .filter(Boolean)
    .join('\n\n')
}

function createPPTElementClipboardPlainText(element: PPTElement) {
  switch (element.kind) {
    case 'comment':
      return [
        element.body,
        ...(element.thread ?? []).map((message) => message.body),
      ].join('\n')
    case 'freeform':
    case 'line':
      return element.name
    case 'image':
      return element.accessibility?.altText || element.alt || element.name
    case 'shape':
    case 'textBox':
      return readPPTTextForClipboard(element.textBody) || element.name
    case 'table':
      return stringifyPPTTableRows(element.rows)
  }
}

function readPPTTextForClipboard(body: PPTTextBody | undefined) {
  if (!body) {
    return ''
  }

  let numberedIndex = 0

  return body.paragraphs
    .map((paragraph) => {
      const listPrefix = paragraph.bullet === 'bullet'
        ? '\u2022 '
        : paragraph.bullet === 'numbered'
          ? `${numberedIndex + 1}. `
          : ''

      if (paragraph.bullet === 'numbered') {
        numberedIndex += 1
      }

      return `${listPrefix}${paragraph.runs.map((run) => run.text).join('')}`
    })
    .join('\n')
}

function createPPTSelectionClipboardFallbackHTML(
  objects: readonly PPTElement[],
  selectionSvg: string | null,
) {
  const contentHTML = objects
    .map(createPPTElementClipboardFallbackHTML)
    .filter(Boolean)
    .join('')
  const semanticHTML = contentHTML
    ? `<div data-ppt-selection-text="true">${contentHTML}</div>`
    : ''
  const previewHTML = selectionSvg
    ? `<div data-ppt-selection-preview="true">${selectionSvg}</div>`
    : ''

  if (!semanticHTML && !previewHTML) {
    return '<p>PPT selection</p>'
  }

  return `<section data-ppt-selection-export="true">${semanticHTML}${previewHTML}</section>`
}

function createPPTElementClipboardFallbackHTML(element: PPTElement) {
  if (element.kind === 'image') {
    return createPPTImageClipboardFallbackHTML(element)
  }

  if (element.kind === 'table') {
    return createPPTTableClipboardHTML(element)
  }

  if (element.kind === 'shape') {
    return createPPTShapeClipboardFallbackHTML(element)
  }

  if (isPPTTextElement(element)) {
    return createPPTTextElementClipboardFallbackHTML(element)
  }

  const text = createPPTElementClipboardPlainText(element).trim()

  if (!text) {
    return ''
  }

  return [
    `<p data-ppt-selection-object="${escapePPTCanvasXmlAttribute(element.id)}">`,
    escapePPTClipboardHTMLText(text).replace(/\n/g, '<br>'),
    '</p>',
  ].join('')
}

function createPPTShapeClipboardFallbackHTML(element: PPTShape) {
  const style = {
    ...getDefaultPPTTextStyle(),
    ...element.style,
  }
  const contentHTML = element.textBody
    ? createPPTTextBodyClipboardHTML(element.textBody)
    : ''
  const borderStyle = element.stroke
    ? `${element.stroke.width}px ${getPPTStrokeDashBorderStyle(element.stroke)} ${element.stroke.color}`
    : undefined
  const shapeStyle = createPPTClipboardStyleAttribute([
    ['align-items', getPPTTextVerticalAlignCSS(getPPTTextElementVerticalAlign(element))],
    ['background', getPPTFillColorCSS(element.fill)],
    ['border', borderStyle],
    ['border-radius', element.shape === 'rect'
      ? getSlideEditObjectCornerRadiusCSS(getPPTShapeCornerRadius(element))
      : element.shape === 'ellipse'
        ? '999px'
        : undefined],
    ['box-sizing', 'border-box'],
    ['clip-path', element.shape === 'diamond'
      ? 'polygon(50% 0,100% 50%,50% 100%,0 50%)'
      : undefined],
    ['color', style.color],
    ['display', 'flex'],
    ['font-family', getPPTTextFontFamilyCSS(style.fontFamily)],
    ['font-size', `${style.fontSize}px`],
    ['font-weight', getPPTClipboardFontWeightCSS(style.fontWeight)],
    ['height', `${element.geometry.h}px`],
    ['line-height', String(PPT_PARAGRAPH_LINE_HEIGHT_DEFAULT)],
    ['min-height', `${element.geometry.h}px`],
    ['padding', getPPTClipboardTextInsetCSS(getPPTTextElementInset(element))],
    ['text-align', getPPTElementParagraphAlign(element)],
    ['width', `${element.geometry.w}px`],
  ])

  return [
    `<section data-ppt-selection-object="${escapePPTCanvasXmlAttribute(element.id)}"${createPPTClipboardGeometryAttributes(element)} data-ppt-selection-shape="${escapePPTCanvasXmlAttribute(element.shape)}"${shapeStyle}>`,
    `<div data-ppt-selection-text-body="true">${contentHTML}</div>`,
    '</section>',
  ].join('')
}

function createPPTTextElementClipboardFallbackHTML(element: PPTTextElement) {
  const style = getPPTTextElementStyle(element)
  const contentHTML = createPPTTextBodyClipboardHTML(element.textBody)

  if (!contentHTML) {
    return ''
  }

  return [
    `<section data-ppt-selection-object="${escapePPTCanvasXmlAttribute(element.id)}"${createPPTClipboardGeometryAttributes(element)} data-ppt-selection-text-body="true"${createPPTClipboardStyleAttribute([
      ['align-items', getPPTTextVerticalAlignCSS(getPPTTextElementVerticalAlign(element))],
      ['box-sizing', 'border-box'],
      ['color', style.color],
      ['display', 'flex'],
      ['font-family', getPPTTextFontFamilyCSS(style.fontFamily)],
      ['font-size', `${style.fontSize}px`],
      ['font-weight', getPPTClipboardFontWeightCSS(style.fontWeight)],
      ['height', `${element.geometry.h}px`],
      ['line-height', String(PPT_PARAGRAPH_LINE_HEIGHT_DEFAULT)],
      ['min-height', `${element.geometry.h}px`],
      ['padding', getPPTClipboardTextInsetCSS(getPPTTextElementInset(element))],
      ['text-align', getPPTElementParagraphAlign(element)],
      ['width', `${element.geometry.w}px`],
    ])}>`,
    contentHTML,
    '</section>',
  ].join('')
}

function createPPTClipboardGeometryAttributes(element: PPTElement) {
  return [
    ['data-ppt-selection-x', element.geometry.x],
    ['data-ppt-selection-y', element.geometry.y],
    ['data-ppt-selection-w', element.geometry.w],
    ['data-ppt-selection-h', element.geometry.h],
  ]
    .map(([attribute, value]) =>
      ` ${attribute}="${escapePPTCanvasXmlAttribute(String(value))}"`)
    .join('')
}

function createPPTTextBodyClipboardHTML(body: PPTTextBody) {
  const parts: string[] = []
  let listKind: PPTParagraph['bullet'] | undefined
  let listItems: string[] = []

  function flushList() {
    if (listItems.length === 0) {
      return
    }

    const tag = listKind === 'numbered' ? 'ol' : 'ul'

    parts.push(
      `<${tag} data-ppt-selection-list="${listKind ?? 'bullet'}">${
        listItems.join('')
      }</${tag}>`,
    )
    listItems = []
    listKind = undefined
  }

  body.paragraphs.forEach((paragraph) => {
    const runsHTML = createPPTTextRunsClipboardHTML(paragraph.runs)

    if (paragraph.bullet === 'bullet' || paragraph.bullet === 'numbered') {
      if (listKind !== paragraph.bullet) {
        flushList()
        listKind = paragraph.bullet
      }

      listItems.push(
        `<li data-ppt-selection-paragraph="true"${createPPTParagraphClipboardStyleAttribute(paragraph)}>${runsHTML || '<br>'}</li>`,
      )
      return
    }

    flushList()
    parts.push(
      `<p data-ppt-selection-paragraph="true"${createPPTParagraphClipboardStyleAttribute(paragraph)}>${runsHTML || '<br>'}</p>`,
    )
  })

  flushList()

  return parts.join('')
}

function createPPTTextRunsClipboardHTML(runs: readonly PPTRun[]) {
  return runs
    .map((run) => createPPTRunClipboardHTML(run))
    .join('')
}

function createPPTRunClipboardHTML(run: PPTRun) {
  let content = escapePPTClipboardHTMLText(run.text).replace(/\n/g, '<br>')

  if (!content) {
    return ''
  }

  if (run.underline === true) {
    content = `<u>${content}</u>`
  }

  if (run.italic === true) {
    content = `<em>${content}</em>`
  }

  if (run.bold === true) {
    content = `<strong>${content}</strong>`
  }

  const styleAttribute = createPPTClipboardStyleAttribute([
    ['color', run.color],
    ['font-size', run.size === undefined ? undefined : `${run.size}px`],
  ])

  return styleAttribute ? `<span${styleAttribute}>${content}</span>` : content
}

function createPPTParagraphClipboardStyleAttribute(paragraph: PPTParagraph) {
  return createPPTClipboardStyleAttribute([
    ['text-align', paragraph.align],
    ['line-height', paragraph.lineHeight === undefined ? undefined : String(paragraph.lineHeight)],
    ['margin-bottom', paragraph.spacingAfter === undefined ? undefined : `${paragraph.spacingAfter}px`],
    ['margin-top', paragraph.spacingBefore === undefined ? undefined : `${paragraph.spacingBefore}px`],
  ])
}

function createPPTClipboardStyleAttribute(
  entries: readonly (readonly [string, string | undefined])[],
) {
  const style = entries
    .filter((entry): entry is readonly [string, string] =>
      entry[1] !== undefined && entry[1] !== '')
    .map(([property, value]) => `${property}:${value}`)
    .join(';')

  return style ? ` style="${escapePPTCanvasXmlAttribute(style)}"` : ''
}

function getPPTClipboardFontWeightCSS(fontWeight: PPTTextStyle['fontWeight']) {
  if (fontWeight === 'bold') {
    return '700'
  }

  if (fontWeight === 'semibold') {
    return '600'
  }

  return '400'
}

function getPPTClipboardTextInsetCSS(inset: PPTTextInset) {
  return `${inset.top}px ${inset.right}px ${inset.bottom}px ${inset.left}px`
}

function createPPTImageClipboardFallbackHTML(element: PPTImage) {
  const altText = createPPTElementClipboardPlainText(element).trim()
  const crop = getPPTImageCrop(element)
  const fit = getPPTImageFit(element)
  const captionHTML = altText
    ? `<figcaption>${escapePPTClipboardHTMLText(altText)}</figcaption>`
    : ''
  const figureStyle = createPPTClipboardStyleAttribute([
    ['height', `${element.geometry.h}px`],
    ['width', `${element.geometry.w}px`],
  ])
  const imageStyle = createPPTClipboardStyleAttribute([
    ['height', '100%'],
    ['object-fit', fit],
    ['object-position', `${crop.x}% ${crop.y}%`],
    ['width', '100%'],
  ])

  return [
    `<figure data-ppt-selection-object="${escapePPTCanvasXmlAttribute(element.id)}"${createPPTClipboardGeometryAttributes(element)} data-ppt-selection-image="true" data-ppt-selection-image-fit="${escapePPTCanvasXmlAttribute(fit)}" data-ppt-selection-image-crop-x="${escapePPTCanvasXmlAttribute(String(crop.x))}" data-ppt-selection-image-crop-y="${escapePPTCanvasXmlAttribute(String(crop.y))}"${figureStyle}>`,
    `<img alt="${escapePPTCanvasXmlAttribute(altText)}" src="${escapePPTCanvasXmlAttribute(element.src)}"${imageStyle}>`,
    captionHTML,
    '</figure>',
  ].join('')
}

function escapePPTClipboardHTMLText(value: string) {
  return escapePPTCanvasXmlAttribute(value)
}

async function writePPTHTMLClipboard({
  html,
  sourceSlideId,
}: {
  html: string
  sourceSlideId: string
}): Promise<PPTRichClipboardWriteMode> {
  const json = stringifyPPTCanvasRichClipboardPayload({
    kind: PPT_HTML_CLIPBOARD_KIND,
    metadata: {
      htmlLength: html.length,
      sourceSlideId,
    },
    version: PPT_HTML_CLIPBOARD_VERSION,
  })

  return writePPTCanvasRichClipboardPayload({
    html,
    json,
    jsonMimeType: PPT_HTML_CLIPBOARD_JSON_MIME_TYPE,
    plainText: html,
  })
}

async function writePPTSlideClipboardPayload({
  payload,
}: {
  payload: PPTSlideClipboardPayload
}): Promise<PPTRichClipboardWriteMode> {
  const json = stringifyPPTSlideClipboardPayload(payload)
  const html = createPPTSlideClipboardHTML(payload)
  const slideSvg = exportPPTSlideSVG(payload.slide)

  return writePPTCanvasRichClipboardPayload({
    html,
    json,
    jsonMimeType: PPT_SLIDE_CLIPBOARD_JSON_MIME_TYPE,
    plainText: createPPTSlideClipboardPlainText(payload.slide),
    selectionSvg: slideSvg,
  })
}

async function writePPTSlideSVGClipboard({
  sourceSlideId,
  svg,
}: {
  sourceSlideId: string
  svg: string
}): Promise<PPTRichClipboardWriteMode> {
  const json = stringifyPPTCanvasRichClipboardPayload({
    kind: PPT_SLIDE_SVG_CLIPBOARD_KIND,
    metadata: {
      sourceSlideId,
      svgLength: svg.length,
    },
    version: PPT_SLIDE_SVG_CLIPBOARD_VERSION,
  })

  return writePPTCanvasRichClipboardPayload({
    html: svg,
    json,
    jsonMimeType: PPT_SLIDE_SVG_CLIPBOARD_JSON_MIME_TYPE,
    plainText: svg,
    selectionSvg: svg,
  })
}

async function writePPTSelectionSVGClipboard({
  selectedObjectIds,
  sourceSlideId,
  svg,
}: {
  selectedObjectIds: readonly string[]
  sourceSlideId: string
  svg: string
}): Promise<PPTRichClipboardWriteMode> {
  const json = stringifyPPTCanvasRichClipboardPayload({
    kind: PPT_SELECTION_SVG_CLIPBOARD_KIND,
    metadata: {
      objectCount: selectedObjectIds.length,
      selectedObjectIds,
      sourceSlideId,
      svgLength: svg.length,
    },
    version: PPT_SELECTION_SVG_CLIPBOARD_VERSION,
  })

  return writePPTCanvasRichClipboardPayload({
    html: svg,
    json,
    jsonMimeType: PPT_SELECTION_SVG_CLIPBOARD_JSON_MIME_TYPE,
    plainText: svg,
    selectionSvg: svg,
  })
}

async function writePPTTableClipboard({
  html,
  objectId,
  plainText,
  rows,
  sourceSlideId,
}: {
  html: string
  objectId: string
  plainText: string
  rows: readonly (readonly string[])[]
  sourceSlideId: string
}): Promise<PPTRichClipboardWriteMode> {
  const json = stringifyPPTCanvasRichClipboardPayload({
    kind: PPT_TABLE_CLIPBOARD_KIND,
    metadata: {
      columnCount: getPPTTableColumnCount(rows),
      objectId,
      rowCount: rows.length,
      sourceSlideId,
    },
    rows,
    version: PPT_TABLE_CLIPBOARD_VERSION,
  })

  return writePPTCanvasRichClipboardPayload({
    html,
    json,
    jsonMimeType: PPT_TABLE_CLIPBOARD_JSON_MIME_TYPE,
    plainText,
  })
}

async function writePPTRichClipboardPayload({
  fallback,
  payload,
  selectionSvg,
}: {
  fallback: PPTRichClipboardFallback
  payload: PPTClipboardPayload
  selectionSvg: string | null
}): Promise<PPTRichClipboardWriteMode> {
  const json = stringifyPPTRichClipboardPayload(payload)
  const html = createPPTRichClipboardHTML({
    fallbackHTML: fallback.html,
    payload,
  })

  return writePPTCanvasRichClipboardPayload({
    html,
    json,
    jsonMimeType: PPT_RICH_CLIPBOARD_JSON_MIME_TYPE,
    plainText: fallback.plainText,
    selectionSvg: selectionSvg ?? '',
  })
}

function getPPTRichClipboardFromDataTransfer(dataTransfer: DataTransfer | null) {
  return readPPTCanvasRichClipboardFromDataTransfer({
    dataTransfer,
    jsonMimeType: PPT_RICH_CLIPBOARD_JSON_MIME_TYPE,
    parsePayload: normalizePPTRichClipboardPayload,
    scriptAttribute: PPT_RICH_CLIPBOARD_HTML_JSON_SCRIPT_ATTRIBUTE,
  })
}

function getPPTSlideClipboardFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  return readPPTCanvasRichClipboardFromDataTransfer({
    dataTransfer,
    jsonMimeType: PPT_SLIDE_CLIPBOARD_JSON_MIME_TYPE,
    parsePayload: normalizePPTSlideClipboardPayload,
    scriptAttribute: PPT_SLIDE_CLIPBOARD_HTML_JSON_SCRIPT_ATTRIBUTE,
  })
}

function normalizePPTSlideClipboardPayload(
  value: unknown,
): PPTSlideClipboardPayload | null {
  const payloadValue = isPPTRecord(value) &&
    value.kind === PPT_SLIDE_CLIPBOARD_KIND &&
    value.version === PPT_SLIDE_CLIPBOARD_VERSION
    ? value.payload
    : value

  if (!isPPTRecord(payloadValue)) {
    return null
  }

  const parsed = PPTSlideSchema.safeParse(payloadValue.slide)

  if (!parsed.success) {
    return null
  }

  return {
    slide: parsed.data,
    sourceSlideId: typeof payloadValue.sourceSlideId === 'string'
      ? payloadValue.sourceSlideId
      : parsed.data.id,
  }
}

function normalizePPTRichClipboardPayload(value: unknown): PPTClipboardPayload | null {
  const payloadValue = isPPTRecord(value) &&
    value.kind === PPT_RICH_CLIPBOARD_KIND &&
    value.version === PPT_RICH_CLIPBOARD_VERSION
    ? value.payload
    : value

  if (!isPPTRecord(payloadValue)) {
    return null
  }

  const sourceSlideId = typeof payloadValue.sourceSlideId === 'string'
    ? payloadValue.sourceSlideId
    : ''
  const rawObjects = Array.isArray(payloadValue.objects)
    ? payloadValue.objects
    : []
  const objects = rawObjects.flatMap((item) => {
    const parsed = PPTElementSchema.safeParse(item)

    return parsed.success ? [parsed.data] : []
  })

  if (!sourceSlideId || objects.length === 0) {
    return null
  }

  const selectedObjectIds = normalizeSlideEditClipboardSelectedObjectIds({
    getObjectId: (object) => object.id,
    objects,
    selectedObjectIds: Array.isArray(payloadValue.selectedObjectIds)
      ? payloadValue.selectedObjectIds
      : null,
  })

  return createPPTClipboardPayload({
    objects,
    operation: 'copy',
    selectedObjectIds,
    sourceSlideId,
  })
}

function isPPTRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function getPPTClipboardPastePositionKey(
  payload: PPTClipboardPayload,
  targetSlideId: string,
) {
  return createPPTCanvasPastePositionKey({
    segments: [
      payload.operation,
      payload.sourceSlideId,
      targetSlideId,
      payload.selectedObjectIds,
      payload.metadata.map((item) => item.objectId),
    ],
  })
}

function createPPTClipboardPasteCommandEffect({
  createId,
  payload,
  slideFrame,
  target,
}: {
  createId: (prefix: string) => string
  payload: PPTClipboardPayload
  slideFrame: Bounds
  target: PPTClipboardPasteTarget
}): PPTClipboardPasteHostCommandEffect | null {
  return createSlideEditClipboardPasteCommandEffect({
    payload,
    remapPolicy: createPPTClipboardRemapPolicy(createId, payload),
    slideFrame,
    target,
  })
}

function createPPTClipboardRemapPolicy(
  createId: (prefix: string) => string,
  payload: PPTClipboardPayload,
): SlideEditClipboardRemapPolicy<string, string, string> {
  const objectsById = new Map(payload.objects.map((object) => [object.id, object]))

  return {
    createGroupId: () => createId('group-copy'),
    createObjectId: (sourceObjectId, index) => {
      const object = objectsById.get(sourceObjectId)

      return createId(object
        ? getPPTElementIdPrefix(object)
        : `slide-object-copy-${index}`)
    },
    createPlaceholderId: () => null,
  }
}

function applyPPTClipboardPasteHostCommandEffect(
  slide: PPTSlide,
  effect: PPTClipboardPasteHostCommandEffect,
): PPTSlide {
  const plan = effect.payload.pastePlan
  const offset = {
    x: plan.anchor.x,
    y: plan.anchor.y,
  }
  const pasted = mapSlideEditClipboardPasteObjects({
    getObjectId: (object) => object.id,
    pastePlan: plan,
    payload: effect.payload.payload,
    transform: ({ mapping, source }) => {
      return clonePPTElementFromClipboardMapping(source, mapping, offset)
    },
  })

  if (pasted.length === 0) {
    return slide
  }

  return {
    ...slide,
    elements: [...slide.elements, ...pasted],
  }
}

function clonePPTElementFromClipboardMapping(
  source: PPTElement,
  mapping: PPTClipboardPasteObjectMapping,
  offset: Point,
): PPTElement {
  const geometry = clampPPTCanvasBoundsToFrame({
    bounds: {
      ...source.geometry,
      x: source.geometry.x + offset.x,
      y: source.geometry.y + offset.y,
    },
    frame: {
      h: PPT_SLIDE_HEIGHT,
      w: PPT_SLIDE_WIDTH,
      x: 0,
      y: 0,
    },
  })

  return {
    ...source,
    geometry: { ...source.geometry, ...geometry },
    groupId: mapping.targetGroupId ?? undefined,
    id: mapping.targetObjectId,
    name: `${source.name} Copy`,
  }
}

function createPPTLayerPaneDescriptor({
  activeObjectId = null,
  collapsedGroupIds,
  selectedObjectIds,
  slide,
}: {
  activeObjectId?: string | null
  collapsedGroupIds: ReadonlySet<string>
  selectedObjectIds: readonly string[]
  slide: PPTSlide
}): PPTLayerPaneDescriptor {
  const objects = getPPTLayerPaneObjectInputs({
    collapsedGroupIds,
    slide,
  })

  return createSlideEditLayerPaneDescriptor({
    activeObjectId,
    objects,
    selectedObjectIds: getPPTLayerPaneSelectedRowIds(slide, selectedObjectIds),
    slideId: slide.id,
  })
}

function getPPTLayerPaneObjectInputs({
  collapsedGroupIds,
  slide,
}: {
  collapsedGroupIds: ReadonlySet<string>
  slide: PPTSlide
}) {
  const groupedElements = new Map<string, PPTElement[]>()

  for (const element of slide.elements) {
    if (!element.groupId) {
      continue
    }

    const group = groupedElements.get(element.groupId) ?? []
    group.push(element)
    groupedElements.set(element.groupId, group)
  }

  const addedGroupIds = new Set<string>()
  const objects: Array<{
    displayName: string
    groupId?: string | null
    isExpanded?: boolean
    isGroup?: boolean
    isHidden?: boolean
    isLocked?: boolean
    isRenamable?: boolean
    isReorderable?: boolean
    isSelectable?: boolean
    kindLabel: string
    objectId: string
    order?: number
    parentObjectId?: string | null
  }> = []
  let order = 0

  for (const element of slide.elements) {
    if (!element.groupId) {
      objects.push(getPPTLayerPaneElementInput(element, order))
      order += 1
      continue
    }

    if (addedGroupIds.has(element.groupId)) {
      continue
    }

    addedGroupIds.add(element.groupId)

    const groupElements = groupedElements.get(element.groupId) ?? []
    const groupRowId = toPPTLayerPaneGroupRowId(element.groupId)
    const isExpanded = !collapsedGroupIds.has(element.groupId)

    objects.push({
      displayName: getPPTLayerPaneGroupDisplayName(addedGroupIds.size),
      groupId: element.groupId,
      isExpanded,
      isGroup: true,
      isHidden: groupElements.every((member) => member.visible === false),
      isLocked: groupElements.length > 0 &&
        groupElements.every((member) => member.locked === true),
      isRenamable: false,
      isReorderable: groupElements.length > 0 &&
        groupElements.every((member) => member.locked !== true),
      isSelectable: groupElements.length > 0,
      kindLabel: 'Group',
      objectId: groupRowId,
      order,
      parentObjectId: null,
    })
    order += 1

    if (!isExpanded) {
      continue
    }

    for (const groupElement of groupElements) {
      objects.push(getPPTLayerPaneElementInput(groupElement, order, groupRowId))
      order += 1
    }
  }

  return objects
}

function getPPTLayerPaneElementInput(
  element: PPTElement,
  order: number,
  parentObjectId: string | null = null,
) {
  return {
    displayName: element.name,
    groupId: element.groupId ?? null,
    isGroup: false,
    isHidden: element.visible === false,
    isLocked: element.locked === true,
    isRenamable: true,
    isReorderable: element.locked !== true,
    isSelectable: true,
    kindLabel: getPPTElementKindLabel(element),
    objectId: element.id,
    order,
    parentObjectId,
  }
}

function getPPTLayerPaneGroupDisplayName(index: number) {
  return `Group ${index}`
}

function toPPTLayerPaneGroupRowId(groupId: string) {
  return `${PPT_LAYER_PANE_GROUP_ROW_PREFIX}${groupId}`
}

function getPPTLayerPaneGroupIdFromRowId(objectId: string) {
  return objectId.startsWith(PPT_LAYER_PANE_GROUP_ROW_PREFIX)
    ? objectId.slice(PPT_LAYER_PANE_GROUP_ROW_PREFIX.length)
    : null
}

function getPPTLayerPaneFullySelectedGroupIds(
  slide: PPTSlide,
  selectedObjectIds: readonly string[],
) {
  return getPPTCanvasFullySelectedItemGroupIds({
    getItemGroupId: (element) => element.groupId,
    getItemId: (element) => element.id,
    items: slide.elements,
    selection: selectedObjectIds,
  })
}

function getPPTLayerPaneSelectedRowIds(
  slide: PPTSlide,
  selectedObjectIds: readonly string[],
) {
  const selectedRowIds = new Set(selectedObjectIds)

  for (const groupId of getPPTLayerPaneFullySelectedGroupIds(slide, selectedObjectIds)) {
    selectedRowIds.add(toPPTLayerPaneGroupRowId(groupId))
  }

  return [...selectedRowIds]
}

function getPPTLayerPaneDefaultFocusObjectId(
  slide: PPTSlide,
  selection: readonly string[],
) {
  const firstSelectedObjectId = selection[0]
  const firstSelectedElement = findPPTElement(slide, firstSelectedObjectId ?? null)

  if (firstSelectedElement?.groupId) {
    if (
      getPPTLayerPaneFullySelectedGroupIds(slide, selection)
        .includes(firstSelectedElement.groupId)
    ) {
      return toPPTLayerPaneGroupRowId(firstSelectedElement.groupId)
    }
  }

  return firstSelectedObjectId ?? null
}

function getPPTLayerPaneActualObjectIds(
  slide: PPTSlide,
  objectIds: readonly string[],
) {
  return uniquePPTCanvasValues(getPPTCanvasGroupExpandedSelectionIds({
    getItemGroupId: (element) => element.groupId,
    getItemId: (element) => element.id,
    getSelectionGroupId: getPPTLayerPaneGroupIdFromRowId,
    items: slide.elements,
    selection: objectIds,
  }))
}

function getPPTObjectVisibilityDescriptors(
  slideId: string,
  rows: readonly PPTLayerPaneRowDescriptor[],
): PPTObjectVisibilityDescriptor[] {
  return rows.map((row) => ({
    isHidden: row.isHidden,
    isLocked: row.isLocked,
    isSelectable: row.isSelectable,
    objectId: row.objectId,
    placeholderId: null,
    slideId,
  }))
}

function getPPTObjectVisibilityState(
  row: PPTLayerPaneRowDescriptor,
  selectionPolicy: SlideEditObjectSelectionPolicy = 'allow-hidden-selection',
): SlideEditObjectVisibilityState {
  return getSlideEditObjectVisibilityState({
    isHidden: row.isHidden,
    isLocked: row.isLocked,
    selectionPolicy,
  })
}

function getPPTObjectVisibilityCommandId(
  row: PPTLayerPaneRowDescriptor,
): SlideEditObjectVisibilityCommandId {
  return row.isHidden ? 'show-objects' : 'hide-objects'
}

function getPPTObjectVisibilityAvailability({
  descriptors,
  row,
}: {
  descriptors: readonly PPTObjectVisibilityDescriptor[]
  row: PPTLayerPaneRowDescriptor
}): SlideEditObjectVisibilityCommandAvailability<string> {
  return getSlideEditObjectVisibilityCommandAvailability({
    commandId: getPPTObjectVisibilityCommandId(row),
    objects: descriptors,
    selectedObjectIds: [row.objectId],
  })
}

function getPPTLayerPaneDropIndex(
  slide: PPTSlide,
  targetObjectId: string,
  placement: PPTLayerPaneDropPlacement,
) {
  const targetGroupId = getPPTLayerPaneGroupIdFromRowId(targetObjectId)

  if (targetGroupId) {
    const groupRange = getPPTElementGroupIndexRange({
      elements: slide.elements,
      groupId: targetGroupId,
    })

    if (!groupRange) {
      return null
    }

    return placement === 'before'
      ? groupRange.firstIndex
      : groupRange.lastIndex + 1
  }

  const targetIndex = slide.elements.findIndex((element) => element.id === targetObjectId)

  if (targetIndex < 0) {
    return null
  }

  return placement === 'before' ? targetIndex : targetIndex + 1
}

function getPPTLayerPaneKeyboardDropIndex(
  descriptor: PPTLayerPaneDescriptor,
  slide: PPTSlide,
  objectId: string,
  toIndex: number,
) {
  const fromIndex = descriptor.rows.findIndex((row) => row.objectId === objectId)

  if (fromIndex < 0) {
    return null
  }

  if (fromIndex < toIndex) {
    const targetRow = descriptor.rows[toIndex - 1]

    return targetRow
      ? getPPTLayerPaneDropIndex(slide, targetRow.objectId, 'after')
      : slide.elements.length
  }

  const targetRow = descriptor.rows[toIndex]

  return targetRow
    ? getPPTLayerPaneDropIndex(slide, targetRow.objectId, 'before')
    : slide.elements.length
}

function reorderPPTLayerPaneElement(
  elements: readonly PPTElement[],
  objectId: string,
  toIndex: number,
) {
  const groupId = getPPTLayerPaneGroupIdFromRowId(objectId)
  const selection = groupId
    ? getPPTElementGroupMemberIds({
      elements,
      groupId,
    })
    : [objectId]
  const result = movePPTElementsToIndex({
    elements,
    selection,
    toIndex,
  })

  return result.changed ? result.items : null
}

function getPPTLayerPaneSelection({
  currentSelection,
  mode,
  objectIds,
  slide,
}: {
  currentSelection: string[]
  mode: 'additive' | 'range' | 'replace'
  objectIds: readonly string[]
  slide: PPTSlide
}) {
  const targetObjectId = objectIds.at(-1)

  if (!targetObjectId) {
    return currentSelection
  }

  const targetGroupId = getPPTLayerPaneGroupIdFromRowId(targetObjectId)

  if (targetGroupId) {
    return getPPTCanvasGroupedItemSelection({
      additive: mode === 'additive',
      fallbackSelection: mode === 'additive' ? currentSelection : [],
      getItemGroupId: (element) => element.groupId,
      getItemId: (element) => element.id,
      groupId: targetGroupId,
      items: slide.elements,
      selection: currentSelection,
    })
  }

  if (mode === 'range') {
    return getPPTLayerPaneActualObjectIds(slide, objectIds)
  }

  return getPPTLayerSelection(
    currentSelection,
    targetObjectId,
    mode === 'additive',
    slide,
  )
}

function FindReplaceStrip({
  activeIndex,
  inputRef,
  matchCount,
  onClose,
  onFindKeyDown,
  onFindQueryChange,
  onNext,
  onPrevious,
  onReplace,
  onReplaceAll,
  onReplaceQueryChange,
  query,
  replaceQuery,
}: {
  activeIndex: number
  inputRef: RefObject<HTMLInputElement | null>
  matchCount: number
  onClose: () => void
  onFindKeyDown: (event: ReactKeyboardEvent<HTMLInputElement>) => void
  onFindQueryChange: (query: string) => void
  onNext: () => void
  onPrevious: () => void
  onReplace: () => void
  onReplaceAll: () => void
  onReplaceQueryChange: (query: string) => void
  query: string
  replaceQuery: string
}) {
  const hasQuery = query.length > 0
  const hasMatches = matchCount > 0

  return (
    <div className="ppt-find-strip" data-ppt-find-strip>
      <input
        aria-label="Find text"
        className="ppt-find-input"
        data-ppt-find-query
        placeholder="Find"
        ref={inputRef}
        value={query}
        onChange={(event) => onFindQueryChange(event.target.value)}
        onKeyDown={onFindKeyDown}
      />
      <span className="ppt-find-count" data-ppt-find-count>
        {hasQuery && hasMatches ? `${activeIndex + 1}/${matchCount}` : '0/0'}
      </span>
      <button
        aria-label="Previous match"
        className="ppt-icon-button"
        data-ppt-find-prev
        disabled={!hasMatches}
        title="Previous match"
        type="button"
        onClick={onPrevious}
      >
        <ChevronUp size={16} />
      </button>
      <button
        aria-label="Next match"
        className="ppt-icon-button"
        data-ppt-find-next
        disabled={!hasMatches}
        title="Next match"
        type="button"
        onClick={onNext}
      >
        <ChevronDown size={16} />
      </button>
      <input
        aria-label="Replace text"
        className="ppt-find-input"
        data-ppt-replace-query
        placeholder="Replace"
        value={replaceQuery}
        onChange={(event) => onReplaceQueryChange(event.target.value)}
      />
      <button
        className="ppt-button"
        data-ppt-find-replace
        disabled={!hasMatches}
        type="button"
        onClick={onReplace}
      >
        Replace
      </button>
      <button
        className="ppt-button"
        data-ppt-find-replace-all
        disabled={!hasMatches}
        type="button"
        onClick={onReplaceAll}
      >
        All
      </button>
      <button
        aria-label="Close find"
        className="ppt-icon-button"
        data-ppt-find-close
        title="Close find"
        type="button"
        onClick={onClose}
      >
        <X size={16} />
      </button>
    </div>
  )
}

function PPTSelectionFloatingBar({
  anchor,
  commandAvailability,
  groups,
  onAlignmentPreviewChange,
  onCommand,
  onFontSizeStep,
  onParagraphAlign,
  onParagraphBulletToggle,
  onParagraphNumberedToggle,
  onShapeKindChange,
  onTextBoldToggle,
  onTextColorChange,
  onTextItalicToggle,
  onTextUnderlineToggle,
  scale,
  shapeMenu,
  textFormat,
}: {
  anchor: PPTSelectionCommandAnchor | null
  commandAvailability: PPTCommandAvailability
  groups: readonly PPTSurfaceCommandViewGroup[]
  onAlignmentPreviewChange: (command: PPTAlignmentPopoverCommand | null) => void
  onCommand: (command: PPTSurfaceCommand) => void
  onFontSizeStep: (delta: number) => void
  onParagraphAlign: (align: NonNullable<PPTParagraph['align']>) => void
  onParagraphBulletToggle: () => void
  onParagraphNumberedToggle: () => void
  onShapeKindChange: (elementId: string, shape: PPTShapeKind) => void
  onTextBoldToggle: () => void
  onTextColorChange: (color: string) => void
  onTextItalicToggle: () => void
  onTextUnderlineToggle: () => void
  scale: number
  shapeMenu: PPTShapeQuickMenuState | null
  textFormat: PPTTextQuickFormatState | null
}) {
  if (!anchor || (groups.length === 0 && !shapeMenu && !textFormat)) {
    return null
  }

  return (
    <div
      aria-label="Selection actions"
      className="ppt-selection-floating-bar"
      data-placement={anchor.placement}
      data-ppt-selection-floating-bar
      role="toolbar"
      style={{
        '--ppt-command-scale': String(1 / scale),
        left: anchor.x,
        top: anchor.y,
      } as CSSProperties}
      onPointerDown={(event) => event.stopPropagation()}
    >
      {textFormat ? (
        <PPTTextQuickFormatControls
          state={textFormat}
          onFontSizeStep={onFontSizeStep}
          onParagraphAlign={onParagraphAlign}
          onParagraphBulletToggle={onParagraphBulletToggle}
          onParagraphNumberedToggle={onParagraphNumberedToggle}
          onTextBoldToggle={onTextBoldToggle}
          onTextColorChange={onTextColorChange}
          onTextItalicToggle={onTextItalicToggle}
          onTextUnderlineToggle={onTextUnderlineToggle}
        />
      ) : null}
      {textFormat ? <span className="ppt-command-divider" /> : null}
      {shapeMenu ? (
        <>
          <PPTShapeKindMenu
            state={shapeMenu}
            onShapeKindChange={onShapeKindChange}
          />
          <span className="ppt-command-divider" />
        </>
      ) : null}
      <PPTAlignmentPopover
        availability={commandAvailability}
        onCommand={onCommand}
        onPreviewChange={onAlignmentPreviewChange}
      />
      {groups.length > 0 ? <span className="ppt-command-divider" /> : null}
      {groups.map((group, groupIndex) => (
        <Fragment key={group.id}>
          {groupIndex > 0 ? <span className="ppt-command-divider" /> : null}
          {group.commands.map((command) => (
            <PPTSurfaceCommandButton
              command={command}
              key={command.command}
              surface="selection-floating-bar"
              onCommand={onCommand}
            />
          ))}
        </Fragment>
      ))}
    </div>
  )
}

function PPTShapeKindMenu({
  state,
  onShapeKindChange,
}: {
  state: PPTShapeQuickMenuState
  onShapeKindChange: (elementId: string, shape: PPTShapeKind) => void
}) {
  const [open, setOpen] = useState(false)
  const [activeShape, setActiveShape] = useState<PPTShapeKind>(state.shape)
  const [initialActiveShapeIndex, setInitialActiveShapeIndex] = useState(0)
  const triggerRef = useRef<HTMLButtonElement | null>(null)
  const activeOption = PPT_SHAPE_MENU_OPTIONS.find((option) =>
    option.shape === activeShape
  ) ?? PPT_SHAPE_MENU_OPTIONS[0]
  const {
    onFocus: handleMenuFocus,
    onKeyDown: handleMenuKeyDown,
    ref: setMenuRoot,
  } = usePPTCanvasMenuRovingFocus<HTMLDivElement>({
    initialActiveIndex: initialActiveShapeIndex,
    onClose: () => {
      closeMenu()
      focusTrigger()
    },
  })

  function focusTrigger() {
    focusPPTCanvasElementOnNextFrame({
      resolveElement: () => triggerRef.current,
    })
  }

  function openMenu(shape = state.shape) {
    setInitialActiveShapeIndex(Math.max(
      0,
      PPT_SHAPE_MENU_OPTIONS.findIndex((option) => option.shape === shape),
    ))
    setOpen(true)
    setActiveShape(shape)
  }

  function closeMenu() {
    setOpen(false)
  }

  function commitShape(shape: PPTShapeKind) {
    onShapeKindChange(state.elementId, shape)
    closeMenu()
    focusTrigger()
  }

  function handleTriggerKeyDown(event: ReactKeyboardEvent<HTMLButtonElement>) {
    const intent = getPPTCanvasMenuTriggerKeyboardIntent({ key: event.key })

    if (intent.preventDefault) {
      event.preventDefault()
    }

    if (intent.kind === 'open-menu') {
      event.stopPropagation()
      openMenu()
    }
  }

  return (
    <span
      className="ppt-floating-menu-wrap"
      data-ppt-shape-menu-open={open ? 'true' : 'false'}
    >
      <button
        aria-controls="ppt-shape-kind-menu"
        aria-expanded={open}
        aria-haspopup="menu"
        aria-label="Shape"
        className="ppt-floating-command"
        data-ppt-shape-menu-trigger
        ref={triggerRef}
        title="Shape"
        type="button"
        onClick={(event) => {
          event.preventDefault()
          event.stopPropagation()
          if (open) {
            closeMenu()
          } else {
            openMenu()
          }
        }}
        onKeyDown={handleTriggerKeyDown}
      >
        {renderPPTShapeMenuIcon(state.shape, 16)}
      </button>
      {open ? (
        <div
          aria-label="Shape"
          className="ppt-floating-menu"
          data-ppt-shape-menu
          data-ppt-shape-menu-active={activeOption.shape}
          data-ppt-shape-menu-model={PPT_SELECTION_TOOLBAR_DROPDOWN_MENU_MODEL}
          id="ppt-shape-kind-menu"
          ref={setMenuRoot}
          role="menu"
          onFocus={handleMenuFocus}
          onKeyDown={handleMenuKeyDown}
        >
          {PPT_SHAPE_MENU_OPTIONS.map((option) => (
            <button
              {...PPT_MENU_ITEM_PROPS}
              aria-checked={state.shape === option.shape}
              className="ppt-floating-menu-item"
              data-ppt-shape-menu-item={option.shape}
              key={option.shape}
              role="menuitemcheckbox"
              tabIndex={option.shape === activeOption.shape ? 0 : -1}
              type="button"
              onClick={(event) => {
                event.preventDefault()
                event.stopPropagation()
                commitShape(option.shape)
              }}
              onFocus={() => setActiveShape(option.shape)}
              onMouseEnter={() => setActiveShape(option.shape)}
            >
              {renderPPTShapeMenuIcon(option.shape, 15)}
              <span>{option.label}</span>
            </button>
          ))}
        </div>
      ) : null}
    </span>
  )
}

function renderPPTShapeMenuIcon(shape: PPTShapeKind, size: number) {
  switch (shape) {
    case 'ellipse':
      return <Circle size={size} />
    case 'diamond':
      return <Diamond size={size} />
    case 'rect':
      return <Square size={size} />
  }
}

function PPTAlignmentPopover({
  availability,
  onCommand,
  onPreviewChange,
}: {
  availability: PPTCommandAvailability
  onCommand: (command: PPTSurfaceCommand) => void
  onPreviewChange: (command: PPTAlignmentPopoverCommand | null) => void
}) {
  const [open, setOpen] = useState(false)
  const [activeCommand, setActiveCommand] =
    useState<PPTAlignmentPopoverCommand>('alignCenter')
  const [initialActiveCommandIndex, setInitialActiveCommandIndex] = useState(0)
  const triggerRef = useRef<HTMLButtonElement | null>(null)
  const commands = PPT_ALIGNMENT_POPOVER_COMMANDS.map((command) => ({
    ...command,
    disabled: !availability[command.command],
  }))
  const enabledCommands = commands.filter((command) => !command.disabled)
  const triggerDisabled = enabledCommands.length === 0
  const activeEnabledCommand = enabledCommands.find((command) =>
    command.command === activeCommand
  ) ?? enabledCommands[0] ?? null
  const activeEnabledCommandId = activeEnabledCommand?.command ?? null
  const {
    onFocus: handleMenuFocus,
    onKeyDown: handleMenuKeyDown,
    ref: setMenuRoot,
  } = usePPTCanvasMenuRovingFocus<HTMLDivElement>({
    initialActiveIndex: initialActiveCommandIndex,
    onClose: () => {
      closePopover()
      focusTrigger()
    },
  })

  useEffect(() => () => onPreviewChange(null), [onPreviewChange])

  function openPopover(command = activeEnabledCommand?.command) {
    if (!command) {
      return
    }

    setInitialActiveCommandIndex(Math.max(
      0,
      enabledCommands.findIndex((enabledCommand) =>
        enabledCommand.command === command
      ),
    ))
    setOpen(true)
    setActiveCommand(command)
    onPreviewChange(command)
  }

  function closePopover() {
    setOpen(false)
    onPreviewChange(null)
  }

  function focusTrigger() {
    focusPPTCanvasElementOnNextFrame({
      resolveElement: () => triggerRef.current,
    })
  }

  function handleTriggerKeyDown(event: ReactKeyboardEvent<HTMLButtonElement>) {
    const intent = getPPTCanvasMenuTriggerKeyboardIntent({ key: event.key })

    if (intent.preventDefault) {
      event.preventDefault()
    }

    if (intent.kind === 'open-menu') {
      event.stopPropagation()
      openPopover()
    }
  }

  return (
    <span
      className="ppt-alignment-popover-wrap"
      data-ppt-alignment-popover-open={open ? 'true' : 'false'}
    >
      <button
        aria-controls="ppt-alignment-popover-menu"
        aria-expanded={open}
        aria-haspopup="menu"
        aria-label="Alignment editor"
        className="ppt-floating-command"
        data-ppt-alignment-popover-trigger
        disabled={triggerDisabled}
        ref={triggerRef}
        title="Alignment editor"
        type="button"
        onClick={(event) => {
          event.preventDefault()
          event.stopPropagation()
          if (open) {
            closePopover()
          } else {
            openPopover()
          }
        }}
        onKeyDown={handleTriggerKeyDown}
      >
        <AlignCenterHorizontal size={16} />
      </button>
      {open ? (
        <div
          aria-label="Alignment editor"
          className="ppt-alignment-popover"
          data-ppt-alignment-popover
          data-ppt-alignment-popover-active={activeEnabledCommandId ?? activeCommand}
          data-ppt-alignment-popover-model={CANVAS_DOM_ALIGNMENT_POPOVER_MODEL}
          id="ppt-alignment-popover-menu"
          ref={setMenuRoot}
          role="menu"
          onFocus={handleMenuFocus}
          onKeyDown={handleMenuKeyDown}
          onMouseLeave={() => onPreviewChange(null)}
        >
          {commands.map((command) => (
            <button
              {...PPT_MENU_ITEM_PROPS}
              aria-disabled={command.disabled}
              className="ppt-alignment-popover-item"
              data-ppt-alignment-popover-command={command.dataCommand}
              data-ppt-alignment-popover-item
              disabled={command.disabled}
              key={command.command}
              role="menuitem"
              tabIndex={command.command === activeEnabledCommand?.command ? 0 : -1}
              type="button"
              onClick={(event) => {
                event.preventDefault()
                event.stopPropagation()
                if (command.disabled) {
                  return
                }

                onCommand(command.command)
                closePopover()
                focusTrigger()
              }}
              onFocus={() => {
                setActiveCommand(command.command)
                onPreviewChange(command.command)
              }}
              onMouseEnter={() => {
                setActiveCommand(command.command)
                onPreviewChange(command.command)
              }}
            >
              {renderPPTAlignmentPopoverIcon(command.command)}
              <span>{command.label}</span>
            </button>
          ))}
        </div>
      ) : null}
    </span>
  )
}

function renderPPTAlignmentPopoverIcon(command: PPTAlignmentPopoverCommand) {
  switch (command) {
    case 'alignLeft':
      return <AlignLeft size={15} />
    case 'alignCenter':
      return <AlignCenterHorizontal size={15} />
    case 'alignRight':
      return <AlignRight size={15} />
    case 'alignTop':
      return <AlignStartVertical size={15} />
    case 'alignMiddle':
      return <AlignCenterVertical size={15} />
    case 'alignBottom':
      return <AlignEndVertical size={15} />
    case 'distributeHorizontal':
      return <AlignHorizontalDistributeCenter size={15} />
    case 'distributeVertical':
      return <AlignVerticalDistributeCenter size={15} />
  }
}

function PPTTextQuickFormatControls({
  onFontSizeStep,
  onParagraphAlign,
  onParagraphBulletToggle,
  onParagraphNumberedToggle,
  onTextBoldToggle,
  onTextColorChange,
  onTextItalicToggle,
  onTextUnderlineToggle,
  state,
}: {
  onFontSizeStep: (delta: number) => void
  onParagraphAlign: (align: NonNullable<PPTParagraph['align']>) => void
  onParagraphBulletToggle: () => void
  onParagraphNumberedToggle: () => void
  onTextBoldToggle: () => void
  onTextColorChange: (color: string) => void
  onTextItalicToggle: () => void
  onTextUnderlineToggle: () => void
  state: PPTTextQuickFormatState
}) {
  return (
    <span className="ppt-text-quick-format" data-ppt-text-quick-bar>
      <button
        aria-label="Bold text"
        aria-pressed={state.isBold}
        className="ppt-floating-command"
        data-ppt-text-quick="bold"
        title="Bold text"
        type="button"
        onClick={(event) => {
          event.preventDefault()
          event.stopPropagation()
          onTextBoldToggle()
        }}
      >
        <Bold size={16} />
      </button>
      <button
        aria-label="Italic text"
        aria-pressed={state.isItalic}
        className="ppt-floating-command"
        data-ppt-text-quick="italic"
        title="Italic text"
        type="button"
        onClick={(event) => {
          event.preventDefault()
          event.stopPropagation()
          onTextItalicToggle()
        }}
      >
        <Italic size={16} />
      </button>
      <button
        aria-label="Underline text"
        aria-pressed={state.isUnderline}
        className="ppt-floating-command"
        data-ppt-text-quick="underline"
        title="Underline text"
        type="button"
        onClick={(event) => {
          event.preventDefault()
          event.stopPropagation()
          onTextUnderlineToggle()
        }}
      >
        <Underline size={16} />
      </button>
      <button
        aria-label="Decrease font size"
        className="ppt-floating-command"
        data-ppt-text-quick="font-size-down"
        title="Decrease font size"
        type="button"
        onClick={(event) => {
          event.preventDefault()
          event.stopPropagation()
          onFontSizeStep(-PPT_TEXT_FONT_SIZE_STEP)
        }}
      >
        <Minus size={16} />
      </button>
      <span className="ppt-font-size-chip" data-ppt-text-quick-size>
        <ALargeSmall size={15} />
        {state.fontSize}
      </span>
      <button
        aria-label="Increase font size"
        className="ppt-floating-command"
        data-ppt-text-quick="font-size-up"
        title="Increase font size"
        type="button"
        onClick={(event) => {
          event.preventDefault()
          event.stopPropagation()
          onFontSizeStep(PPT_TEXT_FONT_SIZE_STEP)
        }}
      >
        <Plus size={16} />
      </button>
      <label className="ppt-floating-color" title="Text color">
        <Baseline size={15} />
        <input
          aria-label="Text color"
          data-ppt-text-quick="color"
          type="color"
          value={state.color}
          onChange={(event) => onTextColorChange(event.target.value)}
          onPointerDown={(event) => event.stopPropagation()}
        />
      </label>
      <button
        aria-label="Toggle bullet list"
        aria-pressed={state.bullet}
        className="ppt-floating-command"
        data-ppt-text-quick="bullet"
        title="Toggle bullet list"
        type="button"
        onClick={(event) => {
          event.preventDefault()
          event.stopPropagation()
          onParagraphBulletToggle()
        }}
      >
        <List size={16} />
      </button>
      <button
        aria-label="Toggle numbered list"
        aria-pressed={state.numbered}
        className="ppt-floating-command"
        data-ppt-text-quick="numbered"
        title="Toggle numbered list"
        type="button"
        onClick={(event) => {
          event.preventDefault()
          event.stopPropagation()
          onParagraphNumberedToggle()
        }}
      >
        <ListOrdered size={16} />
      </button>
      <PPTParagraphAlignRadioGroup
        align={state.align}
        surface="quick"
        onAlignChange={onParagraphAlign}
      />
    </span>
  )
}

function PPTParagraphAlignRadioGroup({
  align,
  onAlignChange,
  surface,
}: {
  align: NonNullable<PPTParagraph['align']>
  onAlignChange: (align: NonNullable<PPTParagraph['align']>) => void
  surface: 'inspector' | 'quick'
}) {
  function selectAlign(
    nextAlign: NonNullable<PPTParagraph['align']>,
    container: HTMLElement | null,
  ) {
    onAlignChange(nextAlign)
    focusPPTCanvasElementBySelectorOnNextFrame<HTMLButtonElement>({
      match: ({ element }) =>
        element.getAttribute('data-ppt-paragraph-align') === nextAlign,
      root: container,
      selector: '[data-ppt-paragraph-align]',
    })
  }

  return (
    <span
      aria-label="Paragraph align"
      className={surface === 'quick'
        ? 'ppt-paragraph-align-radio-group ppt-paragraph-align-radio-group--quick'
        : 'ppt-segmented-control ppt-paragraph-align-radio-group'}
      data-ppt-paragraph-align-focus-model={PPT_RADIO_GROUP_FOCUS_MODEL}
      data-ppt-paragraph-align-keyboard-model={PPT_RADIO_GROUP_KEYBOARD_MODEL}
      data-ppt-paragraph-align-model={PPT_RADIO_GROUP_MODEL}
      data-ppt-paragraph-align-radiogroup={surface}
      role="radiogroup"
      onKeyDown={handlePPTCanvasRadioGroupKeyDown}
    >
      {PPT_PARAGRAPH_ALIGN_OPTIONS.map((option) => {
        const selected = align === option

        return (
          <button
            aria-checked={selected}
            aria-label={`Align text ${option}`}
            aria-pressed={selected}
            className={surface === 'quick' ? 'ppt-floating-command' : undefined}
            data-ppt-paragraph-align={option}
            data-ppt-paragraph-align-surface={surface}
            data-ppt-text-quick={surface === 'quick' ? `align-${option}` : undefined}
            key={option}
            role="radio"
            tabIndex={getPPTCanvasRadioTabIndex({
              checked: selected,
              disabled: false,
            })}
            title={`Align text ${option}`}
            type="button"
            onClick={(event) => {
              event.preventDefault()
              event.stopPropagation()
              selectAlign(option, event.currentTarget.closest('[role="radiogroup"]'))
            }}
          >
            {surface === 'quick'
              ? <PPTTextAlignIcon align={option} size={16} />
              : option}
          </button>
        )
      })}
    </span>
  )
}

function PPTTextAlignIcon({
  align,
  size,
}: {
  align: NonNullable<PPTParagraph['align']>
  size: number
}) {
  switch (align) {
    case 'center':
      return <AlignCenter size={size} />
    case 'right':
      return <AlignRight size={size} />
    case 'left':
      return <AlignLeft size={size} />
  }
}

function PPTContextCommandMenu({
  groups,
  menu,
  onClose,
  onCommand,
}: {
  groups: readonly PPTSurfaceCommandViewGroup[]
  menu: PPTContextMenuState | null
  onClose: () => void
  onCommand: (command: PPTSurfaceCommand) => void
}) {
  const {
    onFocus: handleMenuFocus,
    onKeyDown: handleMenuKeyDown,
    ref: setMenuRoot,
  } = usePPTCanvasMenuRovingFocus<HTMLDivElement>({ onClose })

  if (!menu || groups.length === 0) {
    return null
  }

  return (
    <div
      aria-label="Selection commands"
      className="ppt-context-menu"
      data-ppt-context-menu
      data-ppt-context-menu-focus-model={PPT_MENU_FOCUS_MODEL}
      data-ppt-context-menu-keyboard={PPT_MENU_KEYBOARD_KEYS}
      data-ppt-context-menu-model={PPT_MENU_ROVING_FOCUS_MODEL}
      ref={setMenuRoot}
      role="menu"
      style={{
        left: menu.x,
        top: menu.y,
      }}
      onContextMenu={(event) => event.preventDefault()}
      onFocus={handleMenuFocus}
      onKeyDown={handleMenuKeyDown}
      onPointerDown={(event) => event.stopPropagation()}
    >
      {groups.map((group) => (
        <div className="ppt-context-menu-group" key={group.id} role="group">
          {group.commands.map((command) => (
            <PPTSurfaceCommandButton
              command={command}
              key={command.command}
              surface="context-menu"
              onAfterCommand={onClose}
              onCommand={onCommand}
            />
          ))}
        </div>
      ))}
    </div>
  )
}

function PPTSurfaceCommandButton({
  command,
  onAfterCommand,
  onCommand,
  surface,
}: {
  command: PPTSurfaceCommandView
  onAfterCommand?: () => void
  onCommand: (command: PPTSurfaceCommand) => void
  surface: PPTCommandSurface
}) {
  const dataAttribute = surface === 'context-menu'
    ? {
        ...PPT_MENU_ITEM_PROPS,
        'data-ppt-context-command': command.dataCommand,
      }
    : { 'data-ppt-floating-command': command.dataCommand }

  return (
    <button
      aria-label={command.label}
      className={surface === 'context-menu'
        ? 'ppt-context-menu-item'
        : 'ppt-floating-command'}
      disabled={command.disabled}
      role={surface === 'context-menu' ? 'menuitem' : undefined}
      title={command.title}
      type="button"
      {...dataAttribute}
      onClick={(event) => {
        event.preventDefault()
        event.stopPropagation()

        if (command.disabled) {
          return
        }

        onCommand(command.command)
        onAfterCommand?.()
      }}
    >
      <PPTSurfaceCommandIcon command={command.command} size={16} />
      {surface === 'context-menu' ? <span>{command.label}</span> : null}
    </button>
  )
}

function PPTSurfaceCommandIcon({
  command,
  size,
}: {
  command: PPTSurfaceCommand
  size: number
}) {
  switch (command) {
    case 'alignCenter':
      return <AlignCenterVertical size={size} />
    case 'alignLeft':
      return <AlignStartVertical size={size} />
    case 'alignRight':
      return <AlignEndVertical size={size} />
    case 'bringForward':
      return <MoveUp size={size} />
    case 'bringToFront':
      return <BringToFront size={size} />
    case 'copyFormatting':
      return <Paintbrush size={size} />
    case 'delete':
      return <Trash2 size={size} />
    case 'duplicate':
      return <CopyPlus size={size} />
    case 'flipHorizontal':
      return <FlipHorizontal2 size={size} />
    case 'flipVertical':
      return <FlipVertical2 size={size} />
    case 'group':
      return <Group size={size} />
    case 'lockSelection':
      return <Lock size={size} />
    case 'pasteFormatting':
      return <Paintbrush size={size} />
    case 'selectSameType':
      return <Layers size={size} />
    case 'sendBackward':
      return <MoveDown size={size} />
    case 'sendToBack':
      return <SendToBack size={size} />
    case 'tidySelection':
      return <Grid2X2 size={size} />
    case 'ungroup':
      return <Ungroup size={size} />
    case 'unlockAll':
      return <Unlock size={size} />
  }
}

function SlideThumb({
  active,
  dragging,
  dropPlacement,
  index,
  optionDescriptor,
  onDragEnd,
  onDragOver,
  onDragStart,
  onDrop,
  onKeyDown,
  onSelect,
  slide,
  thumbnailDescriptor,
}: {
  active: boolean
  dragging: boolean
  dropPlacement?: PPTSlideDropPlacement
  index: number
  optionDescriptor?: SlideEditRailListboxOptionDescriptor<string>
  onDragEnd: (event: ReactDragEvent<HTMLButtonElement>) => void
  onDragOver: (event: ReactDragEvent<HTMLButtonElement>) => void
  onDragStart: (event: ReactDragEvent<HTMLButtonElement>) => void
  onDrop: (event: ReactDragEvent<HTMLButtonElement>) => void
  onKeyDown: (event: ReactKeyboardEvent<HTMLButtonElement>) => void
  onSelect: (event: ReactMouseEvent<HTMLButtonElement>) => void
  slide: PPTSlide
  thumbnailDescriptor?: SlideEditRailThumbnailDescriptor<string>
}) {
  const descriptorActive = thumbnailDescriptor?.isActive ?? active
  const descriptorIndex = thumbnailDescriptor?.index ?? index
  const tabIndex = optionDescriptor?.tabIndex ?? (descriptorActive ? 0 : -1)

  return (
    <button
      aria-current={descriptorActive ? 'page' : undefined}
      aria-label={`Open ${slide.name}`}
      aria-selected={optionDescriptor?.isSelected ?? descriptorActive}
      className="ppt-thumb"
      data-ppt-slide-active={descriptorActive ? 'true' : undefined}
      data-ppt-slide-dragging={dragging ? 'true' : undefined}
      data-ppt-slide-draggable="true"
      data-ppt-slide-drop-target={dropPlacement}
      data-ppt-slide-id={slide.id}
      data-ppt-slide-index={descriptorIndex}
      data-ppt-slide-rail-hit-h={thumbnailDescriptor?.hitTarget.h}
      data-ppt-slide-rail-hit-w={thumbnailDescriptor?.hitTarget.w}
      data-ppt-slide-rail-hit-x={thumbnailDescriptor?.hitTarget.x}
      data-ppt-slide-rail-hit-y={thumbnailDescriptor?.hitTarget.y}
      data-ppt-slide-rail-option-focusable={optionDescriptor?.isFocusable ? 'true' : 'false'}
      data-ppt-slide-rail-option-id={optionDescriptor?.id}
      data-ppt-slide-rail-option-index={optionDescriptor?.index}
      data-ppt-slide-rail-option-selected={optionDescriptor?.isSelected ? 'true' : 'false'}
      data-ppt-slide-rail-thumb-active={thumbnailDescriptor?.isActive ? 'true' : 'false'}
      data-ppt-slide-rail-thumb-h={thumbnailDescriptor?.bounds.h}
      data-ppt-slide-rail-thumb-w={thumbnailDescriptor?.bounds.w}
      data-ppt-slide-rail-thumb-x={thumbnailDescriptor?.bounds.x}
      data-ppt-slide-rail-thumb-y={thumbnailDescriptor?.bounds.y}
      data-ppt-slide-roving-tab-index={String(tabIndex)}
      draggable
      onDragEnd={onDragEnd}
      onDragOver={onDragOver}
      onDragStart={onDragStart}
      onDrop={onDrop}
      onClick={onSelect}
      onKeyDown={onKeyDown}
      role="option"
      tabIndex={tabIndex}
      type="button"
    >
      <span className="ppt-thumb-preview" style={{ background: slide.background?.color ?? '#fff' }}>
        {slide.elements.map((element) => (
          <span
            className={getPPTThumbElementClassName(element)}
            data-ppt-thumb-element={element.id}
            data-line-end-marker={element.kind === 'line' ? element.endMarker : undefined}
            data-line-start-marker={element.kind === 'line' ? element.startMarker : undefined}
            data-ppt-freeform-points={element.kind === 'freeform'
              ? element.points.length
              : undefined}
            data-ppt-thumb-comment-resolved={element.kind === 'comment' && element.resolved === true
              ? 'true'
              : undefined}
            data-ppt-thumb-table-cols={element.kind === 'table'
              ? getPPTTableColumnCount(element.rows)
              : undefined}
            data-ppt-thumb-table-rows={element.kind === 'table'
              ? element.rows.length
              : undefined}
            data-ppt-image-crop-x={element.kind === 'image'
              ? getPPTImageCrop(element).x
              : undefined}
            data-ppt-image-crop-y={element.kind === 'image'
              ? getPPTImageCrop(element).y
              : undefined}
            data-ppt-image-fit={element.kind === 'image'
              ? getPPTImageFit(element)
              : undefined}
            data-ppt-thumb-corner-radius={element.kind === 'shape' && element.shape === 'rect'
              ? formatPPTShapeCornerRadius(getPPTShapeCornerRadius(element))
              : undefined}
            data-ppt-thumb-fill-opacity={element.kind === 'shape'
              ? formatPPTFillOpacity(getPPTFillOpacity(element.fill))
              : undefined}
            data-ppt-flip-h={element.flipH === true ? 'true' : undefined}
            data-ppt-flip-v={element.flipV === true ? 'true' : undefined}
            data-ppt-thumb-alt-text={getPPTElementAltText(element)}
            data-ppt-thumb-hyperlink-url={getPPTElementHyperlink(element)?.url}
            data-ppt-thumb-stroke-dash={getPPTElementStrokeDash(element)}
            data-ppt-thumb-shadow={hasPPTElementShadow(element) ? 'true' : undefined}
            data-ppt-thumb-shadow-angle={hasPPTElementShadow(element)
              ? getPPTElementShadow(element).angle
              : undefined}
            data-ppt-thumb-shadow-blur={hasPPTElementShadow(element)
              ? getPPTElementShadow(element).blur
              : undefined}
            data-ppt-thumb-shadow-color={hasPPTElementShadow(element)
              ? getPPTElementShadow(element).color
              : undefined}
            data-ppt-thumb-shadow-distance={hasPPTElementShadow(element)
              ? getPPTElementShadow(element).distance
              : undefined}
            data-ppt-thumb-shadow-opacity={hasPPTElementShadow(element)
              ? formatPPTElementShadowOpacity(getPPTElementShadow(element).opacity)
              : undefined}
            data-ppt-thumb-opacity={formatPPTElementOpacity(getPPTElementOpacity(element))}
            data-ppt-thumb-bullet={isPPTTextElement(element) && hasPPTTextBodyBullet(element.textBody)
              ? 'true'
              : undefined}
            data-ppt-thumb-numbered={isPPTTextElement(element) && hasPPTTextBodyNumbered(element.textBody)
              ? 'true'
              : undefined}
            data-ppt-thumb-font-family={isPPTTextElement(element)
              ? normalizePPTTextFontFamily(getPPTTextElementStyle(element).fontFamily)
              : undefined}
            data-ppt-thumb-text-inset={isPPTTextElement(element)
              ? formatPPTTextInsetData(getPPTTextElementInset(element))
              : undefined}
            data-ppt-thumb-vertical-align={isPPTTextElement(element)
              ? getPPTTextElementVerticalAlign(element)
              : undefined}
            data-shape={element.kind === 'shape' ? element.shape : undefined}
            key={element.id}
            style={{
              background: element.kind === 'shape'
                ? getPPTFillColorCSS(element.fill)
                : element.kind === 'image'
                  ? undefined
                  : element.kind === 'table'
                    ? '#f8fafc'
                    : element.kind === 'comment'
                      ? '#fef3c7'
                      : element.kind === 'textBox'
                        ? '#cbd5e1'
                      : undefined,
              border: element.kind === 'shape' && element.stroke
                ? `${element.stroke.width}px solid ${element.stroke.color}`
                : undefined,
              borderStyle: element.kind === 'shape' && element.stroke
                ? getPPTStrokeDashBorderStyle(element.stroke)
                : undefined,
              borderRadius: element.kind === 'shape' && element.shape === 'rect'
                ? getSlideEditObjectCornerRadiusPreviewCSS({
                    h: element.geometry.h,
                    value: getPPTShapeCornerRadius(element),
                    w: element.geometry.w,
                  })
                : undefined,
              backgroundImage: element.kind === 'image'
                ? `url(${element.src})`
                : undefined,
              backgroundPosition: element.kind === 'image'
                ? getSlideEditObjectImageCropPositionCSS(getPPTImageCrop(element))
                : undefined,
              backgroundSize: element.kind === 'image'
                ? getPPTImageFit(element)
                : undefined,
              height: `${(element.geometry.h / PPT_SLIDE_HEIGHT) * 100}%`,
              fontFamily: isPPTTextElement(element)
                ? getPPTTextFontFamilyCSS(getPPTTextElementStyle(element).fontFamily)
                : undefined,
              alignItems: isPPTTextElement(element)
                ? getPPTTextVerticalAlignCSS(getPPTTextElementVerticalAlign(element))
                : undefined,
              left: `${(element.geometry.x / PPT_SLIDE_WIDTH) * 100}%`,
              filter: getPPTElementShadowFilter(element),
              opacity: getPPTElementOpacity(element),
              top: `${(element.geometry.y / PPT_SLIDE_HEIGHT) * 100}%`,
              transform: getPPTElementTransform(element),
              width: `${(element.geometry.w / PPT_SLIDE_WIDTH) * 100}%`,
              ...getPPTThumbLineDashStyle(element),
            }}
          />
        ))}
      </span>
      <span className="ppt-thumb-name">{descriptorIndex + 1}. {slide.name}</span>
    </button>
  )
}

function PPTElementView({
  editing,
  element,
  eraserHit,
  findActive,
  hovered,
  onCommitText,
  onContextMenu,
  onEdit,
  onInlineEditEffect,
  onPointerDown,
  onPointerEnter,
  onPointerLeave,
  onStopEdit,
  onTextOverflowChange,
  selected,
  slideId,
  textAutoFitIndicator,
  textOverflow,
}: {
  editing: boolean
  element: PPTElement
  eraserHit: boolean
  findActive: boolean
  hovered: boolean
  onCommitText: (elementId: string, text: string) => void
  onContextMenu: (event: ReactMouseEvent<HTMLDivElement>, elementId: string) => void
  onEdit: () => void
  onInlineEditEffect: (effect: PPTInlineEditEffect) => void
  onPointerDown: (event: ReactPointerEvent<HTMLDivElement>, elementId: string) => void
  onPointerEnter: () => void
  onPointerLeave: () => void
  onStopEdit: () => void
  onTextOverflowChange: (elementId: string, hasOverflow: boolean) => void
  selected: boolean
  slideId: string
  textAutoFitIndicator: SlideEditTextOverflowIndicatorState<string, string> | null
  textOverflow: boolean
}) {
  const animation = getPPTElementAnimation(element)
  const style = {
    ...pptElementStyle(element),
    ...getPPTElementAnimationStyle(animation),
  }
  const textBody = isPPTTextElement(element) ? element.textBody : null
  const textStyle = isPPTTextElement(element) ? element.style : undefined
  const text = isPPTTextElement(element) ? readPPTText(element.textBody) : ''
  const editorRef = useRef<HTMLDivElement | null>(null)

  useEffect(() => {
    if (!editing) {
      return
    }

    focusPPTCanvasElement({
      element: editorRef.current,
      preventScroll: false,
    })
  }, [editing])

  useLayoutEffect(() => {
    if (!textBody) {
      return
    }

    const editor = editorRef.current

    if (!editor) {
      return
    }

    const updateOverflow = () => {
      const measurement = measurePPTCanvasElementOverflow({
        container: editor.parentElement,
        element: editor,
        epsilon: PPT_TEXT_OVERFLOW_EPSILON,
      })

      onTextOverflowChange(element.id, measurement?.hasOverflow === true)
    }

    updateOverflow()
    const frame = schedulePPTCanvasAnimationFrameTask({
      task: updateOverflow,
    })

    return () => {
      cancelPPTCanvasAnimationFrameTask({ frame })
    }
  }, [
    editing,
    element.geometry.h,
    element.geometry.w,
    element.id,
    textStyle,
    onTextOverflowChange,
    text,
    textBody,
  ])

  function recordInlineEditEffect(effect: Omit<PPTInlineEditEffect, 'elementId' | 'model'>) {
    onInlineEditEffect({
      elementId: element.id,
      model: PPT_INLINE_EDIT_DOM_MODEL,
      ...effect,
    })
  }

  function handleInlineEditBeforeInput(event: ReactFormEvent<HTMLDivElement>) {
    if (!editing) {
      return
    }

    const inputType = getPPTInlineEditInputType(event)
    const historyDirection = getPPTInlineEditHistoryDirectionFromInputType(inputType)
    const lineBreak = isPPTInlineEditLineBreakInput(inputType)

    if (!historyDirection && !lineBreak) {
      return
    }

    recordInlineEditEffect({
      ...(historyDirection ? { historyDirection } : {}),
      inputType,
      ...(lineBreak ? { lineBreak: true } : {}),
    })
  }

  function handleInlineEditKeyDown(event: ReactKeyboardEvent<HTMLDivElement>) {
    if (!editing) {
      return
    }

    const intent = getPPTCanvasInlineEditKeyboardIntent({
      altKey: event.altKey,
      ctrlKey: event.ctrlKey,
      key: event.key,
      metaKey: event.metaKey,
      shiftKey: event.shiftKey,
    })

    if (intent.kind === 'history') {
      recordInlineEditEffect({
        historyDirection: intent.historyDirection,
      })
      return
    }

    if (intent.kind === 'line-break') {
      recordInlineEditEffect({
        inputType: intent.inputType,
        ...(isPPTInlineEditLineBreakInput(intent.inputType) ? { lineBreak: true } : {}),
      })
      return
    }

    if (intent.preventDefault) {
      event.preventDefault()
    }

    if (intent.kind === 'cancel') {
      event.currentTarget.innerText = text
      event.currentTarget.blur()
      onStopEdit()
      return
    }

    if (intent.kind === 'commit') {
      event.currentTarget.blur()
    }
  }

  function handleInlineEditPaste(event: ReactClipboardEvent<HTMLDivElement>) {
    if (!editing) {
      return
    }

    const pasteText = getPPTTextPasteSourcesFromDataTransfer(
      event.clipboardData,
    )[0]

    if (!pasteText) {
      return
    }

    event.preventDefault()
    insertPPTInlineEditText(event.currentTarget, pasteText)
    recordInlineEditEffect({ pasteText })
  }

  return (
    <div
      className="ppt-element"
      data-hovered={hovered ? 'true' : 'false'}
      data-group-id={element.groupId}
      data-kind={element.kind}
      data-ppt-comment-resolved={element.kind === 'comment' && element.resolved === true
        ? 'true'
        : undefined}
      data-ppt-comment-thread-count={element.kind === 'comment'
        ? getPPTCommentThread(element).length
        : undefined}
      data-ppt-eraser-hit={eraserHit ? 'true' : undefined}
      data-line-end-connection={element.kind === 'line'
        ? element.endConnection?.elementId
        : undefined}
      data-line-route={element.kind === 'line'
        ? element.route ?? 'straight'
        : undefined}
      data-line-end-x={element.kind === 'line'
        ? element.end.x
        : undefined}
      data-line-end-y={element.kind === 'line'
        ? element.end.y
        : undefined}
      data-line-start-x={element.kind === 'line'
        ? element.start.x
        : undefined}
      data-line-start-y={element.kind === 'line'
        ? element.start.y
        : undefined}
      data-line-start-connection={element.kind === 'line'
        ? element.startConnection?.elementId
        : undefined}
      data-ppt-freeform-points={element.kind === 'freeform'
        ? element.points.length
        : undefined}
      data-ppt-find-active={findActive ? 'true' : undefined}
      data-ppt-flip-h={element.flipH === true ? 'true' : undefined}
      data-ppt-flip-v={element.flipV === true ? 'true' : undefined}
      data-ppt-animation-delay={animation.delayMs}
      data-ppt-animation-duration={animation.durationMs}
      data-ppt-animation-order={animation.order}
      data-ppt-animation-trigger={animation.trigger}
      data-ppt-animation-type={animation.type}
      data-ppt-alt-text={getPPTElementAltText(element)}
      data-ppt-corner-radius={element.kind === 'shape' && element.shape === 'rect'
        ? formatPPTShapeCornerRadius(getPPTShapeCornerRadius(element))
        : undefined}
      data-ppt-fill-opacity={element.kind === 'shape'
        ? formatPPTFillOpacity(getPPTFillOpacity(element.fill))
        : undefined}
      data-ppt-hyperlink-url={getPPTElementHyperlink(element)?.url}
      data-ppt-opacity={formatPPTElementOpacity(getPPTElementOpacity(element))}
      data-ppt-stroke-dash={getPPTElementStrokeDash(element)}
      data-ppt-shadow={hasPPTElementShadow(element) ? 'true' : undefined}
      data-ppt-shadow-angle={hasPPTElementShadow(element)
        ? getPPTElementShadow(element).angle
        : undefined}
      data-ppt-shadow-blur={hasPPTElementShadow(element)
        ? getPPTElementShadow(element).blur
        : undefined}
      data-ppt-shadow-color={hasPPTElementShadow(element)
        ? getPPTElementShadow(element).color
        : undefined}
      data-ppt-shadow-distance={hasPPTElementShadow(element)
        ? getPPTElementShadow(element).distance
        : undefined}
      data-ppt-shadow-opacity={hasPPTElementShadow(element)
        ? formatPPTElementShadowOpacity(getPPTElementShadow(element).opacity)
        : undefined}
      data-ppt-font-family={isPPTTextElement(element)
        ? normalizePPTTextFontFamily(textStyle?.fontFamily)
        : undefined}
      data-ppt-text-inset={isPPTTextElement(element)
        ? formatPPTTextInsetData(getPPTTextElementInset(element))
        : undefined}
      data-ppt-vertical-align={isPPTTextElement(element)
        ? getPPTTextElementVerticalAlign(element)
        : undefined}
      data-ppt-text-autofit={isPPTTextElement(element) ? element.textAutoFit : undefined}
      data-ppt-text-autofit-model={isPPTTextElement(element)
        ? 'slide-edit-text-box-auto-fit'
        : undefined}
      data-ppt-text-autofit-size-mode={textAutoFitIndicator?.sizeMode ??
        (isPPTTextElement(element) ? getPPTTextAutoFitSizeMode(element) : undefined)}
      data-ppt-text-overflow={textOverflow ? 'true' : undefined}
      data-ppt-text-overflow-indicator-anchor={textAutoFitIndicator?.anchor}
      data-ppt-text-overflow-indicator-axis={textAutoFitIndicator?.overflowAxis.join(' ')}
      data-ppt-text-overflow-indicator-height={textAutoFitIndicator?.bounds.h}
      data-ppt-text-overflow-indicator-model={textAutoFitIndicator
        ? 'slide-edit-text-box-auto-fit'
        : undefined}
      data-ppt-text-overflow-indicator-slide={textAutoFitIndicator?.slideId ?? (isPPTTextElement(element) ? slideId : undefined)}
      data-ppt-text-overflow-indicator-visible={textAutoFitIndicator
        ? String(textAutoFitIndicator.isVisible)
        : undefined}
      data-ppt-text-overflow-indicator-width={textAutoFitIndicator?.bounds.w}
      data-ppt-image-crop-x={element.kind === 'image'
        ? getPPTImageCrop(element).x
        : undefined}
      data-ppt-image-crop-y={element.kind === 'image'
        ? getPPTImageCrop(element).y
        : undefined}
      data-ppt-image-fit={element.kind === 'image'
        ? getPPTImageFit(element)
        : undefined}
      data-ppt-table-cols={element.kind === 'table'
        ? getPPTTableColumnCount(element.rows)
        : undefined}
      data-ppt-table-rows={element.kind === 'table'
        ? element.rows.length
        : undefined}
      data-ppt-bullet-list={textBody && hasPPTTextBodyBullet(textBody) ? 'true' : undefined}
      data-ppt-numbered-list={textBody && hasPPTTextBodyNumbered(textBody) ? 'true' : undefined}
      data-locked={element.locked === true ? 'true' : 'false'}
      data-ppt-element={element.id}
      data-ppt-element-name={element.name}
      data-rotation={Math.round(element.geometry.rotation ?? 0)}
      data-selected={selected ? 'true' : 'false'}
      data-shape={element.kind === 'shape' ? element.shape : undefined}
      onContextMenu={(event) => onContextMenu(event, element.id)}
      onDoubleClick={onEdit}
      onPointerDown={(event) => onPointerDown(event, element.id)}
      onPointerEnter={onPointerEnter}
      onPointerLeave={onPointerLeave}
      style={style}
    >
      {element.kind === 'image' ? (
        <img
          alt={getPPTImageAltText(element)}
          draggable={false}
          src={element.src}
          style={{
            objectFit: getPPTImageFit(element),
            objectPosition: getSlideEditObjectImageCropPositionCSS(getPPTImageCrop(element)),
          }}
        />
      ) : element.kind === 'line' ? (
        <PPTLineSvg element={element} />
      ) : element.kind === 'freeform' ? (
        <PPTFreeformSvg element={element} />
      ) : element.kind === 'table' ? (
        <PPTTableView element={element} />
      ) : element.kind === 'comment' ? (
        <PPTCommentView element={element} />
      ) : (
        <div
          className="ppt-element-editor"
          contentEditable={editing}
          data-ppt-inline-edit-active={editing ? 'true' : 'false'}
          data-ppt-inline-edit-model={PPT_INLINE_EDIT_DOM_MODEL}
          ref={editorRef}
          suppressContentEditableWarning={true}
          onBeforeInput={handleInlineEditBeforeInput}
          onBlur={(event) => {
            onCommitText(element.id, event.currentTarget.innerText)
            onStopEdit()
          }}
          onKeyDown={handleInlineEditKeyDown}
          onPaste={handleInlineEditPaste}
        >
          {editing || !textBody ? text : <PPTTextBodyView body={textBody} />}
        </div>
      )}
    </div>
  )
}

function PPTCommentView({ element }: { element: PPTComment }) {
  const author = element.authorName ?? PPT_COMMENT_DEFAULT_AUTHOR
  const createdAt = element.createdAt ?? PPT_COMMENT_DEFAULT_CREATED_AT
  const thread = getPPTCommentThread(element)

  return (
    <div
      className="ppt-comment-card"
      data-ppt-comment-card
      data-ppt-comment-resolved={element.resolved === true ? 'true' : undefined}
      data-ppt-comment-thread-count={thread.length}
    >
      <div className="ppt-comment-meta">
        <MessageSquare size={15} />
        <span data-ppt-comment-author>{author}</span>
        <span data-ppt-comment-created>{createdAt}</span>
      </div>
      <p data-ppt-comment-body>{element.body || PPT_COMMENT_DEFAULT_BODY}</p>
    </div>
  )
}

function getPPTInlineEditInputType(event: ReactFormEvent<HTMLElement>) {
  const inputType = (event.nativeEvent as InputEvent).inputType

  return typeof inputType === 'string' ? inputType : ''
}

function PPTTableView({ element }: { element: PPTTable }) {
  const columnCount = getPPTTableColumnCount(element.rows)

  return (
    <div
      className="ppt-table-grid"
      data-ppt-table-cols={columnCount}
      data-ppt-table-rows={element.rows.length}
      style={{
        gridTemplateColumns: `repeat(${columnCount}, minmax(0, 1fr))`,
      }}
    >
      {element.rows.flatMap((row, rowIndex) =>
        row.map((cell, columnIndex) => (
          <div
            className="ppt-table-cell"
            data-ppt-table-cell={`${rowIndex}:${columnIndex}`}
            data-ppt-table-header={rowIndex === 0 ? 'true' : undefined}
            key={`${rowIndex}:${columnIndex}`}
          >
            {cell}
          </div>
        )),
      )}
    </div>
  )
}

function createPPTTableClipboardHTML(element: PPTTable) {
  const rows = normalizePPTTableRows(element.rows)
  const [header = [], ...bodyRows] = rows
  const headerHTML = header.length > 0
    ? `<thead><tr>${header.map((cell) =>
        `<th>${escapePPTCanvasXmlAttribute(cell)}</th>`).join('')}</tr></thead>`
    : ''
  const bodyHTML = bodyRows.length > 0
    ? `<tbody>${bodyRows.map((row) =>
        `<tr>${row.map((cell) =>
          `<td>${escapePPTCanvasXmlAttribute(cell)}</td>`).join('')}</tr>`).join('')}</tbody>`
    : ''

  return `<table data-ppt-selection-object="${escapePPTCanvasXmlAttribute(element.id)}"${createPPTClipboardGeometryAttributes(element)} data-ppt-selection-table="true" data-ppt-table-export="${escapePPTCanvasXmlAttribute(element.id)}"${createPPTClipboardStyleAttribute([
    ['height', `${element.geometry.h}px`],
    ['width', `${element.geometry.w}px`],
  ])}>${headerHTML}${bodyHTML}</table>`
}

function PPTTextBodyView({ body }: { body: PPTTextBody }) {
  return (
    <>
      {body.paragraphs.map((paragraph, index) => (
        <span
          className="ppt-text-paragraph"
          data-ppt-bullet={paragraph.bullet === 'bullet' ? 'true' : undefined}
          data-ppt-list={paragraph.bullet}
          data-ppt-numbered={paragraph.bullet === 'numbered' ? 'true' : undefined}
          data-ppt-line-height={getPPTParagraphLineHeight(paragraph)}
          data-ppt-spacing-after={getPPTParagraphSpacingAfter(paragraph)}
          data-ppt-spacing-before={getPPTParagraphSpacingBefore(paragraph)}
          key={index}
          style={getPPTParagraphStyle(paragraph)}
        >
          {paragraph.runs.map((run, runIndex) => (
            <span
              data-ppt-run-bold={run.bold === true ? 'true' : undefined}
              data-ppt-run-italic={run.italic === true ? 'true' : undefined}
              data-ppt-run-underline={run.underline === true ? 'true' : undefined}
              key={runIndex}
              style={pptTextRunStyle(run)}
            >
              {run.text}
            </span>
          ))}
        </span>
      ))}
    </>
  )
}

function PPTLineSvg({ element }: { element: PPTLine }) {
  const markerId = `${element.id}-arrow-marker`
  const markerEnd = element.endMarker === 'arrow' ? `url(#${markerId})` : undefined
  const markerStart = element.startMarker === 'arrow' ? `url(#${markerId})` : undefined

  return (
    <svg
      aria-hidden="true"
      focusable="false"
      preserveAspectRatio="none"
      viewBox={`0 0 ${element.geometry.w} ${element.geometry.h}`}
    >
      {(element.startMarker === 'arrow' || element.endMarker === 'arrow') ? (
        <defs>
          <marker
            id={markerId}
            markerHeight="8"
            markerUnits="strokeWidth"
            markerWidth="8"
            orient="auto-start-reverse"
            refX="7"
            refY="4"
            viewBox="0 0 8 8"
          >
            <path d="M 0 0 L 8 4 L 0 8 z" fill={element.stroke.color} />
          </marker>
        </defs>
      ) : null}
      {(element.route ?? 'straight') === 'elbow' ? (
        <path
          data-ppt-line-path
          d={getPPTLinePath(element)}
          fill="none"
          markerEnd={markerEnd}
          markerStart={markerStart}
          stroke={element.stroke.color}
          strokeDasharray={getPPTStrokeDashArray(element.stroke)}
          strokeLinecap="round"
          strokeLinejoin="round"
          strokeWidth={element.stroke.width}
        />
      ) : (
        <line
          markerEnd={markerEnd}
          markerStart={markerStart}
          stroke={element.stroke.color}
          strokeDasharray={getPPTStrokeDashArray(element.stroke)}
          strokeLinecap="round"
          strokeWidth={element.stroke.width}
          x1={element.start.x}
          x2={element.end.x}
          y1={element.start.y}
          y2={element.end.y}
        />
      )}
    </svg>
  )
}

function PPTFreeformSvg({ element }: { element: PPTFreeform }) {
  return (
    <svg
      aria-hidden="true"
      focusable="false"
      preserveAspectRatio="none"
      viewBox={`0 0 ${element.geometry.w} ${element.geometry.h}`}
    >
      <path
        data-ppt-freeform-path
        d={createPPTCanvasSvgFreehandPathData(element.points)}
        fill="none"
        stroke={element.stroke.color}
        strokeDasharray={getPPTStrokeDashArray(element.stroke)}
        strokeLinecap="round"
        strokeLinejoin="round"
        strokeWidth={element.stroke.width}
      />
    </svg>
  )
}

function SelectionOverlay({
  bounds,
  canResize,
  onRotatePointerDown,
  onResizePointerDown,
  scale,
  selectedElements,
  textAutoFitIndicator,
  textOverflow,
}: {
  bounds: Bounds
  canResize: boolean
  onRotatePointerDown: (event: ReactPointerEvent<HTMLButtonElement>) => void
  onResizePointerDown: (
    event: ReactPointerEvent<HTMLButtonElement>,
    handle: ResizeHandle,
  ) => void
  scale: number
  selectedElements: PPTElement[]
  textAutoFitIndicator: SlideEditTextOverflowIndicatorState<string, string> | null
  textOverflow: boolean
}) {
  return (
    <>
      <Box className="ppt-selection-box" bounds={bounds} />
      <div
        className="ppt-size-capsule"
        data-ppt-text-autofit-model={textAutoFitIndicator
          ? 'slide-edit-text-box-auto-fit'
          : undefined}
        data-ppt-text-autofit-size-mode={textAutoFitIndicator?.sizeMode}
        data-ppt-text-overflow={textOverflow ? 'true' : undefined}
        data-ppt-text-overflow-indicator-axis={textAutoFitIndicator?.overflowAxis.join(' ')}
        data-ppt-text-overflow-indicator-visible={textAutoFitIndicator
          ? String(textAutoFitIndicator.isVisible)
          : undefined}
        style={{
          left: bounds.x + bounds.w / 2,
          top: bounds.y + bounds.h + 8,
          transform: `translateX(-50%) scale(${1 / scale})`,
        }}
      >
        {selectedElements.length === 1
          ? `${Math.round(bounds.w)} x ${Math.round(bounds.h)}`
          : `${selectedElements.length} objects`}
        {textOverflow ? <span data-ppt-size-capsule-overflow>Overflow</span> : null}
      </div>
      {canResize ? (
        <button
          aria-label="Rotate selection"
          className="ppt-rotate-handle"
          data-ppt-rotate-handle
          onPointerDown={onRotatePointerDown}
          style={{
            left: bounds.x + bounds.w / 2,
            top: bounds.y - 34 / scale,
            transform: `translate(-50%, -50%) scale(${1 / scale})`,
          }}
          type="button"
        >
          <RotateCw size={14} />
        </button>
      ) : null}
      {canResize ? PPT_RESIZE_HANDLES.map((handle) => {
        const point = getPPTCanvasHandlePoint(bounds, handle)
        const size = 10 / scale

        return (
          <button
            aria-label={`Resize ${handle}`}
            className="ppt-resize-handle"
            key={handle}
            onPointerDown={(event) => onResizePointerDown(event, handle)}
            style={{
              cursor: `${handle}-resize`,
              height: size,
              left: point.x - size / 2,
              top: point.y - size / 2,
              width: size,
            }}
            type="button"
          />
        )
      }) : null}
    </>
  )
}

function LineEndpointOverlay({
  line,
  onPointerDown,
  scale,
}: {
  line: PPTLine
  onPointerDown: (
    event: ReactPointerEvent<HTMLButtonElement>,
    endpoint: 'end' | 'start',
  ) => void
  scale: number
}) {
  return (
    <>
      {(['start', 'end'] as const).map((endpoint) => {
        const point = getPPTLineEndpointPoint(line, endpoint)
        const size = 13 / scale

        return (
          <button
            aria-label={`Move line ${endpoint}`}
            className="ppt-line-endpoint-handle"
            data-ppt-line-endpoint={endpoint}
            key={endpoint}
            onPointerDown={(event) => onPointerDown(event, endpoint)}
            style={{
              height: size,
              left: point.x - size / 2,
              top: point.y - size / 2,
              width: size,
            }}
            type="button"
          />
        )
      })}
    </>
  )
}

function LineRouteOverlay({
  line,
  onPointerDown,
  scale,
}: {
  line: PPTLine
  onPointerDown: (event: ReactPointerEvent<HTMLButtonElement>) => void
  scale: number
}) {
  const point = getPPTLineBendPoint(line)
  const size = 15 / scale

  return (
    <button
      aria-label="Move line route"
      className="ppt-line-route-handle"
      data-ppt-line-route-handle
      onPointerDown={onPointerDown}
      style={{
        height: size,
        left: point.x - size / 2,
        top: point.y - size / 2,
        width: size,
      }}
      type="button"
    />
  )
}

function FrameGuides({ geometry }: { geometry: SlideEditFrameGuideGeometry }) {
  return (
    <div
      className="ppt-slide-frame-guides"
      data-ppt-frame-guide-columns={geometry.columns.length}
      data-ppt-frame-guide-lines={geometry.lines.length}
      data-ppt-frame-guide-regions={geometry.regions.length}
      data-ppt-frame-guides
    >
      {geometry.regions.map((region) => (
        <div
          className="ppt-frame-guide-region"
          data-ppt-frame-guide-kind={region.kind}
          data-ppt-frame-guide-region={region.id}
          key={region.id}
          style={{
            height: region.h,
            left: region.x,
            top: region.y,
            width: region.w,
          }}
        />
      ))}
      {geometry.columns.map((column) => (
        <div
          className="ppt-frame-guide-column"
          data-ppt-frame-guide-column={column.id}
          data-ppt-frame-guide-index={column.index}
          data-ppt-frame-guide-kind={column.kind}
          key={column.id}
          style={{
            height: column.h,
            left: column.x,
            top: column.y,
            width: column.w,
          }}
        />
      ))}
      {geometry.lines.map((line) => (
        <div
          className={[
            'ppt-frame-guide',
            `ppt-frame-guide-${line.orientation}`,
            `ppt-frame-guide-${line.kind}`,
          ].join(' ')}
          data-ppt-frame-guide-axis={line.axis}
          data-ppt-frame-guide-kind={line.kind}
          data-ppt-frame-guide-line={line.id}
          data-ppt-frame-guide-side={line.side}
          key={line.id}
          style={line.orientation === 'vertical'
            ? { height: line.length, left: line.x, top: line.y }
            : { left: line.x, top: line.y, width: line.length }}
        />
      ))}
    </div>
  )
}

function Box({ bounds, className }: { bounds: Bounds; className: string }) {
  return (
    <div
      className={className}
      style={{
        height: bounds.h,
        left: bounds.x,
        top: bounds.y,
        width: bounds.w,
      }}
    />
  )
}

function PPTLaserTrailOverlay({
  points,
  scale,
}: {
  points: Point[]
  scale: number
}) {
  const lastPoint = points.at(-1)
  const strokeWidth = Math.max(2, 5 / scale)
  const dotRadius = Math.max(4, 7 / scale)

  return (
    <svg
      aria-hidden="true"
      className="ppt-laser-trail"
      data-ppt-laser-trail
      data-ppt-laser-trail-point-count={points.length}
      height={PPT_SLIDE_HEIGHT}
      viewBox={`0 0 ${PPT_SLIDE_WIDTH} ${PPT_SLIDE_HEIGHT}`}
      width={PPT_SLIDE_WIDTH}
    >
      <path
        className="ppt-laser-trail-path"
        d={createPPTCanvasSvgPathData(points)}
        data-ppt-laser-trail-path
        strokeWidth={strokeWidth}
      />
      {lastPoint ? (
        <circle
          className="ppt-laser-trail-dot"
          cx={lastPoint.x}
          cy={lastPoint.y}
          data-ppt-laser-trail-dot
          r={dotRadius}
        />
      ) : null}
    </svg>
  )
}

function Guides({ guides, scale }: { guides: PPTCanvasSnapGuides; scale: number }) {
  return (
    <>
      {guides.alignmentGuides.map((guide, index) => (
        <div
          className={`ppt-guide ppt-guide-${guide.orientation}`}
          data-ppt-alignment-guide="true"
          data-ppt-alignment-guide-orientation={guide.orientation}
          data-ppt-alignment-guide-position={guide.position}
          key={`${guide.orientation}-${guide.position}-${index}`}
          style={guide.orientation === 'vertical'
            ? { height: guide.end - guide.start, left: guide.position, top: guide.start }
            : { left: guide.start, top: guide.position, width: guide.end - guide.start }}
        />
      ))}
      {guides.spacingGuides.map((guide, guideIndex) => {
        const labelPoint = getSpacingGuideLabelPoint(guide)

        return (
          <Fragment key={`${guide.orientation}-${guide.gap}-${guideIndex}`}>
            {guide.segments.map((segment, segmentIndex) => (
              <div
                className={`ppt-spacing-guide ppt-spacing-guide-${guide.orientation}`}
                data-ppt-spacing-guide="true"
                data-ppt-spacing-guide-gap={guide.gap}
                data-ppt-spacing-guide-orientation={guide.orientation}
                data-ppt-spacing-guide-segment-index={segmentIndex}
                key={`${segment.start.x}-${segment.start.y}-${segment.end.x}-${segment.end.y}-${segmentIndex}`}
                style={getSpacingGuideSegmentStyle(segment)}
              />
            ))}
            <span
              className="ppt-spacing-label"
              data-ppt-spacing-guide-label="true"
              data-ppt-spacing-guide-label-gap={guide.gap}
              data-ppt-spacing-guide-label-orientation={guide.orientation}
              style={{
                left: labelPoint.x,
                top: labelPoint.y,
                transform: `translate(-50%, -50%) scale(${1 / scale})`,
              }}
            >
              {guide.gap}
            </span>
          </Fragment>
        )
      })}
    </>
  )
}

function PPTColorSwatchStrip({
  channel,
  currentColor,
  objectIds,
  recentColors,
  slideId,
  themeColorTokens,
  onSelect,
}: {
  channel: PPTColorSwatchChannel
  currentColor: string
  objectIds: readonly string[]
  recentColors: readonly string[]
  slideId: string
  themeColorTokens: readonly SlideEditThemeColorToken[]
  onSelect: (color: string, swatch: PPTColorSwatchSelection) => void
}) {
  const descriptor = getPPTColorSwatchDescriptor({
    channel,
    currentColor,
    objectIds,
    recentColors,
    slideId,
    themeColorTokens,
  })
  const swatchCount = descriptor.sections.reduce(
    (count, section) => count + section.swatches.length,
    0,
  )
  const recentSection = descriptor.sections.find((section) => section.id === 'recent')

  return (
    <div
      className="ppt-color-swatch-strip"
      data-ppt-color-swatch-channel={channel}
      data-ppt-color-swatch-command={descriptor.field.commandId}
      data-ppt-color-swatch-control={descriptor.field.control}
      data-ppt-color-swatch-count={swatchCount}
      data-ppt-color-swatch-disabled={descriptor.state.isDisabled ? 'true' : 'false'}
      data-ppt-color-swatch-disabled-reason={descriptor.state.disabledReason}
      data-ppt-color-swatch-mixed={descriptor.state.isMixed ? 'true' : 'false'}
      data-ppt-color-swatch-model={descriptor.surface}
      data-ppt-color-swatch-object-ids={descriptor.objectIds.join(' ')}
      data-ppt-color-swatch-package-channel={descriptor.channel.id}
      data-ppt-color-swatch-palette={channel}
      data-ppt-color-swatch-selected-id={descriptor.state.selectedSwatchId}
      data-ppt-color-swatch-strip={channel}
      data-ppt-recent-color-count={recentSection?.swatches.length ?? 0}
    >
      {descriptor.sections.map((section) =>
        section.swatches.length > 0 ? (
          <div
            className="ppt-color-swatch-group"
            data-ppt-color-swatch-group={section.id}
            key={section.id}
          >
            {section.swatches.map((swatch) => (
              <button
                aria-label={`${swatch.label} ${channel}`}
                aria-pressed={swatch.selected}
                className="ppt-color-swatch"
                data-ppt-color-source={swatch.source}
                data-ppt-color-swatch={channel}
                data-ppt-color-swatch-id={swatch.id}
                data-ppt-color-swatch-selected={swatch.selected ? 'true' : undefined}
                data-ppt-color-token={swatch.tokenId}
                data-ppt-color-value={swatch.value}
                disabled={descriptor.state.isDisabled}
                key={swatch.id}
                style={{ backgroundColor: swatch.value }}
                title={swatch.label}
                type="button"
                onClick={() =>
                  onSelect(swatch.value, {
                    source: swatch.source,
                    swatchId: swatch.id,
                    tokenId: swatch.tokenId,
                    value: swatch.value,
                  })}
              />
            ))}
          </div>
        ) : null)}
    </div>
  )
}

function Inspector({
  exportCode,
  inspectorSurface,
  layoutDescriptors,
  layoutPlaceholderVisibilityDescriptors,
  layoutPlaceholders,
  lastPlaceholderVisibilityEffect,
  lastTextAutoFitEffect,
  onColorSwatchApply,
  onCommentBodyChange,
  onCommentReplyAdd,
  onCommentResolvedChange,
  onCommitText,
  onCopyHTML,
  onDownloadHTML,
  onElementAltTextChange,
  onElementAnimationChange,
  onElementGeometryChange,
  onElementHyperlinkChange,
  onElementNameChange,
  onElementOpacityChange,
  onElementRotationChange,
  onElementShadowChange,
  onElementStrokeChange,
  onElementTextInsetChange,
  onElementTextStyleChange,
  onImageCropChange,
  onImageCropReset,
  onImageFitChange,
  onImageReplaceFile,
  onLayerPaneCommandEffect,
  onObjectVisibilityCommandEffect,
  onLayoutPlaceholderVisibilityChange,
  onLineMarkerChange,
  onLineRouteChange,
  onParagraphBulletChange,
  onParagraphNumberedChange,
  onParagraphAlignChange,
  onParagraphSpacingChange,
  onShapeCornerRadiusChange,
  onShapeFillChange,
  onShapeKindChange,
  onSlideBackgroundChange,
  onSlideLayoutChange,
  onSlideNameChange,
  onSlideNotesChange,
  onSlideTransitionChange,
  onTableRowsChange,
  onTextAutoFit,
  recentColors,
  selection,
  selectedElement,
  selectedElementAnimation,
  selectedTextOverflow,
  textAutoFitIndicator,
  slide,
  slideMetadataDescriptor,
  slideLayoutId,
  slideThemeId,
  slideTransition,
  themeColorTokens,
}: {
  exportCode: string
  inspectorSurface: PPTInspectorSurfaceId
  layoutDescriptors: readonly SlideEditLayoutDescriptor[]
  layoutPlaceholderVisibilityDescriptors: readonly SlideEditPlaceholderDescriptor<string, string>[]
  layoutPlaceholders: readonly SlideEditResolvedLayoutPlaceholder[]
  lastPlaceholderVisibilityEffect: PPTLayoutPlaceholderVisibilityHostCommandEffect | null
  lastTextAutoFitEffect: SlideEditTextAutoFitHostCommandEffect<string, string> | null
  onColorSwatchApply: (
    elementId: string,
    channel: PPTColorSwatchChannel,
    color: string,
    swatch: PPTColorSwatchSelection,
  ) => void
  onCommentBodyChange: (elementId: string, value: string) => void
  onCommentReplyAdd: (elementId: string, value: string) => void
  onCommentResolvedChange: (elementId: string, resolved: boolean) => void
  onCommitText: (elementId: string, text: string) => void
  onCopyHTML: () => void
  onDownloadHTML: () => void
  onElementAltTextChange: (elementId: string, altText: string) => void
  onElementAnimationChange: (
    elementId: string,
    field: PPTElementAnimationUpdateField,
    value: PPTElementAnimation[PPTElementAnimationUpdateField],
  ) => void
  onElementGeometryChange: (
    elementId: string,
    field: 'h' | 'w' | 'x' | 'y',
    value: number,
  ) => void
  onElementHyperlinkChange: (elementId: string, url: string) => void
  onElementNameChange: (elementId: string, name: string) => void
  onElementOpacityChange: (elementId: string, opacity: number) => void
  onElementRotationChange: (elementId: string, rotation: number) => void
  onElementShadowChange: (
    elementId: string,
    field: PPTElementShadowUpdateField,
    value: boolean | number | string,
  ) => void
  onElementStrokeChange: (
    elementId: string,
    field: keyof PPTStroke,
    value: string | number,
  ) => void
  onElementTextStyleChange: (
    elementId: string,
    field: keyof PPTTextStyle,
    value: string | number,
  ) => void
  onElementTextInsetChange: (
    elementId: string,
    field: PPTTextInsetField,
    value: number,
  ) => void
  onImageCropChange: (
    elementId: string,
    field: keyof PPTImageCrop,
    value: number,
  ) => void
  onImageCropReset: (elementId: string) => void
  onImageFitChange: (
    elementId: string,
    fit: PPTImageFit,
  ) => void
  onImageReplaceFile: (
    elementId: string,
    file: Blob & { name?: string },
  ) => Promise<boolean>
  onLayerPaneCommandEffect: (effect: PPTLayerPaneHostCommandEffect) => void
  onObjectVisibilityCommandEffect: (
    effect: PPTObjectVisibilityHostCommandEffect,
  ) => void
  onLayoutPlaceholderVisibilityChange: (
    placeholderId: string,
    isVisible: boolean,
  ) => void
  onLineMarkerChange: (
    elementId: string,
    field: 'endMarker' | 'startMarker',
    value: PPTLineMarker,
  ) => void
  onLineRouteChange: (
    elementId: string,
    route: PPTLineRoute,
  ) => void
  onParagraphBulletChange: (
    elementId: string,
    enabled: boolean,
  ) => void
  onParagraphNumberedChange: (
    elementId: string,
    enabled: boolean,
  ) => void
  onParagraphAlignChange: (
    elementId: string,
    align: NonNullable<PPTParagraph['align']>,
  ) => void
  onParagraphSpacingChange: (
    elementId: string,
    field: PPTParagraphSpacingField,
    value: number,
  ) => void
  onShapeCornerRadiusChange: (
    elementId: string,
    cornerRadius: number,
  ) => void
  onShapeFillChange: (
    elementId: string,
    field: keyof PPTFill,
    value: number | string,
  ) => void
  onShapeKindChange: (elementId: string, shape: PPTShapeKind) => void
  onSlideBackgroundChange: (color: string) => void
  onSlideLayoutChange: (layoutId: string) => void
  onSlideNameChange: (name: string) => void
  onSlideNotesChange: (notes: string) => void
  onSlideTransitionChange: (
    field: PPTSlideTransitionUpdateField,
    value: PPTSlideTransition[PPTSlideTransitionUpdateField],
  ) => void
  onTableRowsChange: (elementId: string, value: string) => void
  onTextAutoFit: (elementId: string) => void
  recentColors: readonly string[]
  selection: string[]
  selectedElement: PPTElement | null
  selectedElementAnimation: PPTElementAnimation | null
  selectedTextOverflow: boolean
  textAutoFitIndicator: SlideEditTextOverflowIndicatorState<string, string> | null
  slide: PPTSlide
  slideMetadataDescriptor: PPTSlideMetadataInspectorDescriptor
  slideLayoutId: string
  slideThemeId: string
  slideTransition: PPTSlideTransition
  themeColorTokens: readonly SlideEditThemeColorToken[]
}) {
  const textStyle = selectedElement && isPPTTextElement(selectedElement)
    ? selectedElement.style
    : null
  const paragraphAlign = selectedElement && isPPTTextElement(selectedElement)
    ? selectedElement.textBody?.paragraphs[0]?.align ?? 'left'
    : 'left'
  const paragraphBullet = selectedElement && isPPTTextElement(selectedElement)
    ? hasPPTTextBodyBullet(selectedElement.textBody)
    : false
  const paragraphNumbered = selectedElement && isPPTTextElement(selectedElement)
    ? hasPPTTextBodyNumbered(selectedElement.textBody)
    : false
  const paragraphSpacing = selectedElement && isPPTTextElement(selectedElement)
    ? getPPTTextElementParagraphSpacing(selectedElement)
    : getDefaultPPTParagraphSpacing()
  const paragraphSpacingDescriptor = selectedElement && isPPTTextElement(selectedElement)
    ? getPPTTextParagraphSpacingDescriptor(slide.id, selectedElement)
    : null
  const paragraphLineHeightField = getPPTTextParagraphSpacingField(
    paragraphSpacingDescriptor,
    'lineHeightRatio',
  )
  const paragraphBeforeField = getPPTTextParagraphSpacingField(
    paragraphSpacingDescriptor,
    'paragraphBefore',
  )
  const paragraphAfterField = getPPTTextParagraphSpacingField(
    paragraphSpacingDescriptor,
    'paragraphAfter',
  )
  const strokeLineStyleDescriptor = selectedElement
    ? getPPTStrokeLineStyleDescriptor(slide.id, selectedElement)
    : null
  const cornerRadiusDescriptor = selectedElement?.kind === 'shape'
    ? getPPTCornerRadiusDescriptor(slide.id, selectedElement)
    : null
  const fillOpacityDescriptor = selectedElement?.kind === 'shape'
    ? getPPTFillOpacityDescriptor(slide.id, selectedElement)
    : null
  const imageCropDescriptor = selectedElement?.kind === 'image'
    ? getPPTImageCropDescriptor(slide.id, selectedElement)
    : null
  const imageReplaceDescriptor = selectedElement?.kind === 'image'
    ? getPPTImageReplaceDescriptor(slide.id, selectedElement)
    : null
  const objectOpacityDescriptor = selectedElement
    ? getPPTObjectOpacityDescriptor(slide.id, selectedElement)
    : null
  const objectHyperlinkDescriptor = selectedElement
    ? getPPTObjectHyperlinkDescriptor(slide.id, selectedElement)
    : null
  const objectHyperlinkUrlField = objectHyperlinkDescriptor?.fields.find((field) =>
    field.id === 'url'
  )
  const objectAccessibilityDescriptor = selectedElement
    ? getPPTObjectAccessibilityDescriptor(slide.id, selectedElement)
    : null
  const objectAccessibilityAltTextField = objectAccessibilityDescriptor?.fields.find((field) =>
    field.id === 'altText'
  )
  const objectShadowDescriptor = selectedElement
    ? getPPTObjectShadowDescriptor(slide.id, selectedElement)
    : null
  const objectShadowEnabledField = getPPTObjectShadowField(objectShadowDescriptor, 'enabled')
  const objectShadowColorField = getPPTObjectShadowField(objectShadowDescriptor, 'color')
  const objectShadowOpacityField = getPPTObjectShadowField(objectShadowDescriptor, 'opacity')
  const objectShadowBlurField = getPPTObjectShadowField(objectShadowDescriptor, 'blur')
  const objectShadowDistanceField = getPPTObjectShadowField(objectShadowDescriptor, 'distance')
  const objectShadowAngleField = getPPTObjectShadowField(objectShadowDescriptor, 'angle')
  const objectAnimationDescriptor = selectedElement
    ? getPPTObjectAnimationDescriptor(slide, selectedElement)
    : null
  const imageReplaceInputRef = useRef<HTMLInputElement | null>(null)
  const [commentReplyDraftById, setCommentReplyDraftById] =
    useState<Record<string, string>>({})
  const commentThread = selectedElement?.kind === 'comment'
    ? getPPTCommentThread(selectedElement)
    : []
  const commentReplyDraft = selectedElement?.kind === 'comment'
    ? commentReplyDraftById[selectedElement.id] ?? ''
    : ''
  const elementHyperlink = selectedElement
    ? getPPTElementHyperlink(selectedElement)
    : null
  const elementAltText = selectedElement
    ? getPPTElementAltText(selectedElement) ?? ''
    : ''
  const elementShadow = selectedElement
    ? getPPTElementShadow(selectedElement)
    : PPT_DEFAULT_ELEMENT_SHADOW
  const elementShadowEnabled = selectedElement
    ? hasPPTElementShadow(selectedElement)
    : false
  const objectShadowEnabled = objectShadowDescriptor?.metadata.isEnabled ?? elementShadowEnabled
  const textInset = selectedElement && isPPTTextElement(selectedElement)
    ? getPPTTextElementInset(selectedElement)
    : PPT_DEFAULT_TEXT_BOX_INSET
  const textFontFamilyDescriptor = selectedElement && isPPTTextElement(selectedElement)
    ? getPPTTextFontFamilyDescriptor(slide.id, selectedElement)
    : null
  const textFrameInsetDescriptor = selectedElement && isPPTTextElement(selectedElement)
    ? getPPTTextFrameInsetDescriptor(slide.id, selectedElement)
    : null
  const textVerticalAlignmentDescriptor = selectedElement && isPPTTextElement(selectedElement)
    ? getPPTTextVerticalAlignmentDescriptor(slide.id, selectedElement)
    : null
  const nameMetadataField = getPPTSlideMetadataField(slideMetadataDescriptor, 'name')
  const backgroundMetadataField = getPPTSlideMetadataField(slideMetadataDescriptor, 'background')
  const notesMetadataField = getPPTSlideMetadataField(slideMetadataDescriptor, 'notes')
  const sizeMetadataField = getPPTSlideMetadataField(slideMetadataDescriptor, 'size')
  const orientationMetadataField = getPPTSlideMetadataField(slideMetadataDescriptor, 'orientation')
  const slideTransitionDescriptor = getPPTSlideTransitionDescriptor(
    slide.id,
    slideTransition,
  )
  const [layerPaneGroupState, setLayerPaneGroupState] =
    useState<PPTLayerPaneGroupState>({
      collapsedGroupIds: [],
      focusedObjectId: null,
      rangeAnchorObjectId: null,
    })
  const [layerPaneRenameState, setLayerPaneRenameState] =
    useState<PPTLayerPaneRenameState | null>(null)
  const [layerPaneDragState, setLayerPaneDragState] =
    useState<PPTLayerPaneDragState | null>(null)
  const layerPaneRenameCommitSuppressedRef = useRef(false)
  const collapsedLayerPaneGroupIdSet = useMemo(
    () => new Set(layerPaneGroupState.collapsedGroupIds),
    [layerPaneGroupState.collapsedGroupIds],
  )
  const defaultLayerPaneFocusObjectId = getPPTLayerPaneDefaultFocusObjectId(
    slide,
    selection,
  )
  const layerPaneDescriptor = createPPTLayerPaneDescriptor({
    activeObjectId: layerPaneGroupState.focusedObjectId ??
      defaultLayerPaneFocusObjectId,
    collapsedGroupIds: collapsedLayerPaneGroupIdSet,
    selectedObjectIds: selection,
    slide,
  })
  const objectVisibilityDescriptors = getPPTObjectVisibilityDescriptors(
    slide.id,
    layerPaneDescriptor.rows,
  )
  const activeLayerPaneObjectId = getSlideEditLayerPaneResolvedFocusObjectId(
    layerPaneDescriptor,
    {
      defaultObjectId: defaultLayerPaneFocusObjectId,
      preferredObjectId: layerPaneGroupState.focusedObjectId,
    },
  )
  const layerPaneCommandIds = PPT_LAYER_PANE_COMMANDS.map((command) => command.id).join(' ')
  const layoutPlaceholderById = new Map(layoutPlaceholders.map((placeholder) => [
    placeholder.placeholderId,
    placeholder,
  ]))
  const hiddenPlaceholderCount = layoutPlaceholderVisibilityDescriptors
    .filter((placeholder) => !placeholder.isVisible).length
  const hasSelectedElement = selectedElement !== null
  const [activeInspectorTabId, setActiveInspectorTabId] = useState<PPTInspectorTabId>(
    hasSelectedElement ? 'selection' : 'slide',
  )
  const previousInspectorSelectionStateRef = useRef(hasSelectedElement)
  const inspectorTabsDescriptor = createPPTCanvasTabsDescriptor({
    activation: 'automatic',
    activeId: activeInspectorTabId,
    tabs: PPT_INSPECTOR_TABS,
  })
  const slideInspectorPanelAttributes = getPPTInspectorPanelAttributes(
    inspectorTabsDescriptor,
    'slide',
  )
  const selectionInspectorPanelAttributes = getPPTInspectorPanelAttributes(
    inspectorTabsDescriptor,
    'selection',
  )

  useEffect(() => {
    if (previousInspectorSelectionStateRef.current === hasSelectedElement) {
      return
    }

    previousInspectorSelectionStateRef.current = hasSelectedElement
    setActiveInspectorTabId(hasSelectedElement ? 'selection' : 'slide')
  }, [hasSelectedElement])

  function updateCommentReplyDraft(elementId: string, value: string) {
    setCommentReplyDraftById((current) => ({
      ...current,
      [elementId]: normalizePPTCommentReplyBody(value),
    }))
  }

  function commitCommentReplyDraft(elementId: string) {
    const value = commentReplyDraftById[elementId] ?? ''

    if (value.trim().length === 0) {
      return
    }

    onCommentReplyAdd(elementId, value)
    setCommentReplyDraftById((current) => ({
      ...current,
      [elementId]: '',
    }))
  }

  function runLayerPaneIntent(intent: PPTLayerPaneIntent) {
    if (intent.type === 'visibility-toggle') {
      const row = layerPaneDescriptor.rows.find((candidate) =>
        candidate.objectId === intent.objectId
      )

      if (!row) {
        return
      }

      const effect = getSlideEditObjectVisibilityCommandEffect({
        commandId: getPPTObjectVisibilityCommandId(row),
        objects: objectVisibilityDescriptors,
        selectedObjectIds: [row.objectId],
        slideId: slide.id,
      })

      if (effect) {
        onObjectVisibilityCommandEffect(effect)
      }

      return
    }

    const effect = getSlideEditLayerPaneCommandEffect(layerPaneDescriptor, intent)

    if (effect) {
      onLayerPaneCommandEffect(effect)
    }
  }

  function focusLayerPaneRow(objectId: string) {
    focusPPTCanvasElementBySelectorOnNextFrame<HTMLElement>({
      match: ({ element }) =>
        element.getAttribute('data-ppt-layer-pane-row') === objectId,
      root: document,
      selector: '[data-ppt-layer-pane-row]',
    })
  }

  function setLayerPaneFocusedObjectId(
    objectId: string,
    options?: { rangeAnchor?: boolean },
  ) {
    setLayerPaneGroupState((current) => ({
      ...current,
      focusedObjectId: objectId,
      rangeAnchorObjectId: options?.rangeAnchor
        ? objectId
        : current.rangeAnchorObjectId,
    }))
  }

  function setLayerPaneGroupExpanded(objectId: string, isExpanded: boolean) {
    const groupId = getPPTLayerPaneGroupIdFromRowId(objectId)

    if (!groupId) {
      return
    }

    setLayerPaneGroupState((current) => {
      const collapsedGroupIds = new Set(current.collapsedGroupIds)

      if (isExpanded) {
        collapsedGroupIds.delete(groupId)
      } else {
        collapsedGroupIds.add(groupId)
      }

      return {
        ...current,
        collapsedGroupIds: [...collapsedGroupIds],
        focusedObjectId: objectId,
      }
    })
  }

  function focusLayerPaneRenameInput(objectId: string) {
    focusPPTCanvasElementBySelectorOnNextFrame<HTMLInputElement>({
      match: ({ element }) =>
        element.getAttribute('data-ppt-layer-pane-rename-input') === objectId,
      root: document,
      select: true,
      selector: '[data-ppt-layer-pane-rename-input]',
    })
  }

  function startLayerPaneRename(row: PPTLayerPaneRowDescriptor) {
    if (!row.isRenamable) {
      return
    }

    layerPaneRenameCommitSuppressedRef.current = false
    setLayerPaneFocusedObjectId(row.objectId)
    setLayerPaneRenameState({
      objectId: row.objectId,
      value: row.displayName,
    })
    focusLayerPaneRenameInput(row.objectId)
  }

  function commitLayerPaneRename(objectId: string, value: string) {
    layerPaneRenameCommitSuppressedRef.current = true
    setLayerPaneRenameState(null)
    runLayerPaneIntent({
      name: value,
      objectId,
      type: 'rename-submit',
    })
    focusLayerPaneRow(objectId)
  }

  function cancelLayerPaneRename(objectId: string) {
    layerPaneRenameCommitSuppressedRef.current = true
    setLayerPaneRenameState(null)
    focusLayerPaneRow(objectId)
  }

  function handleLayerPaneRenameBlur(objectId: string, value: string) {
    if (layerPaneRenameCommitSuppressedRef.current) {
      layerPaneRenameCommitSuppressedRef.current = false
      return
    }

    commitLayerPaneRename(objectId, value)
  }

  function canDragLayerPaneRow(row: PPTLayerPaneRowDescriptor) {
    return row.isReorderable
  }

  function handleLayerPaneRowPress(
    row: PPTLayerPaneRowDescriptor,
    event: ReactMouseEvent<HTMLButtonElement>,
  ) {
    const rangeAnchorObjectId = event.shiftKey
      ? layerPaneGroupState.rangeAnchorObjectId ?? activeLayerPaneObjectId
      : null
    const selectionModifierState = getPPTCanvasSelectionListModifierState({
      ctrlKey: event.ctrlKey,
      hasRangeAnchor: Boolean(rangeAnchorObjectId),
      metaKey: event.metaKey,
      shiftKey: event.shiftKey,
    })

    setLayerPaneFocusedObjectId(row.objectId, {
      rangeAnchor: !selectionModifierState.range,
    })
    runLayerPaneIntent({
      ...(selectionModifierState.range && rangeAnchorObjectId
        ? { rangeAnchorObjectId }
        : { additive: selectionModifierState.additive }),
      objectId: row.objectId,
      type: 'row-press',
    })
  }

  function handleLayerPaneRowDragStart(
    row: PPTLayerPaneRowDescriptor,
    event: ReactDragEvent<HTMLElement>,
  ) {
    if (!canDragLayerPaneRow(row)) {
      event.preventDefault()
      return
    }

    setLayerPaneFocusedObjectId(row.objectId)
    setLayerPaneDragState({ objectId: row.objectId })
    setPPTCanvasDataTransferText({
      dataTransfer: event.dataTransfer,
      effectAllowed: 'move',
      text: row.objectId,
    })
  }

  function handleLayerPaneRowDragOver(
    row: PPTLayerPaneRowDescriptor,
    event: ReactDragEvent<HTMLElement>,
  ) {
    const draggedObjectId = layerPaneDragState?.objectId ||
      getPPTCanvasDataTransferText({ dataTransfer: event.dataTransfer })

    if (!draggedObjectId) {
      return
    }

    const rowGeometry = getPPTCanvasPointerLocalGeometry({
      event,
      target: event.currentTarget,
    })

    if (!rowGeometry) {
      setLayerPaneDragState({ objectId: draggedObjectId })
      return
    }

    const dropIndicator = getSlideEditLayerPaneDropIndicator(
      layerPaneDescriptor,
      {
        draggedObjectId,
        pointerOffsetY: rowGeometry.point.y,
        rowHeight: rowGeometry.rect.height,
        targetObjectId: row.objectId,
      },
    )

    if (dropIndicator.placement === 'none' || dropIndicator.targetObjectId === null) {
      setLayerPaneDragState({ objectId: draggedObjectId })
      return
    }

    const dropIndex = getPPTLayerPaneDropIndex(
      slide,
      dropIndicator.targetObjectId,
      dropIndicator.placement,
    )

    if (dropIndex === null) {
      setLayerPaneDragState({ objectId: draggedObjectId })
      return
    }

    event.preventDefault()
    setPPTCanvasDataTransferDropEffect({
      dataTransfer: event.dataTransfer,
      dropEffect: 'move',
    })
    setLayerPaneDragState({
      dropPlacement: dropIndicator.placement,
      dropTargetObjectId: dropIndicator.targetObjectId,
      dropToIndex: dropIndex,
      objectId: draggedObjectId,
    })
  }

  function handleLayerPaneRowDrop(
    row: PPTLayerPaneRowDescriptor,
    event: ReactDragEvent<HTMLElement>,
  ) {
    const draggedObjectId = layerPaneDragState?.objectId ||
      getPPTCanvasDataTransferText({ dataTransfer: event.dataTransfer })

    if (!draggedObjectId || draggedObjectId === row.objectId) {
      return
    }

    const activeDropIndex = layerPaneDragState?.dropTargetObjectId === row.objectId
      ? layerPaneDragState.dropToIndex
      : undefined

    const rowGeometry = getPPTCanvasPointerLocalGeometry({
      event,
      target: event.currentTarget,
    })

    if (!rowGeometry) {
      return
    }

    const dropIndicator = getSlideEditLayerPaneDropIndicator(
      layerPaneDescriptor,
      {
        draggedObjectId,
        pointerOffsetY: rowGeometry.point.y,
        rowHeight: rowGeometry.rect.height,
        targetObjectId: row.objectId,
      },
    )

    if (
      activeDropIndex === undefined &&
      (dropIndicator.placement === 'none' || dropIndicator.targetObjectId === null)
    ) {
      return
    }

    const dropIndex = activeDropIndex ?? getPPTLayerPaneDropIndex(
      slide,
      dropIndicator.targetObjectId ?? row.objectId,
      dropIndicator.placement === 'none' ? 'before' : dropIndicator.placement,
    )

    if (dropIndex === null) {
      return
    }

    event.preventDefault()
    event.stopPropagation()
    setLayerPaneDragState(null)
    setLayerPaneFocusedObjectId(draggedObjectId)
    runLayerPaneIntent({
      objectId: draggedObjectId,
      toIndex: dropIndex,
      type: 'row-drop',
    })
    focusLayerPaneRow(draggedObjectId)
  }

  function handleLayerPaneRowDragEnd() {
    setLayerPaneDragState(null)
  }

  function applyLayerPaneKeyboardIntent(intent: PPTLayerPaneKeyboardIntent) {
    switch (intent.type) {
      case 'focus-row':
      case 'focus-parent-row':
      case 'select-row':
        setLayerPaneFocusedObjectId(intent.objectId, {
          rangeAnchor: true,
        })
        runLayerPaneIntent({
          objectId: intent.objectId,
          type: 'row-press',
        })
        focusLayerPaneRow(intent.objectId)
        return
      case 'range-select-row':
        setLayerPaneGroupState((current) => ({
          ...current,
          focusedObjectId: intent.objectId,
          rangeAnchorObjectId: intent.rangeAnchorObjectId,
        }))
        runLayerPaneIntent({
          objectId: intent.objectId,
          rangeAnchorObjectId: intent.rangeAnchorObjectId,
          type: 'row-press',
        })
        focusLayerPaneRow(intent.objectId)
        return
      case 'reorder-row':
        {
          const dropIndex = getPPTLayerPaneKeyboardDropIndex(
            layerPaneDescriptor,
            slide,
            intent.objectId,
            intent.toIndex,
          )

          if (dropIndex === null) {
            return
          }

          setLayerPaneFocusedObjectId(intent.objectId)
          runLayerPaneIntent({
            objectId: intent.objectId,
            toIndex: dropIndex,
            type: 'row-drop',
          })
          focusLayerPaneRow(intent.objectId)
        }
        return
      case 'collapse-row':
        setLayerPaneGroupExpanded(intent.objectId, false)
        focusLayerPaneRow(intent.objectId)
        return
      case 'expand-row':
        setLayerPaneGroupExpanded(intent.objectId, true)
        focusLayerPaneRow(intent.objectId)
        return
      case 'rename-row':
        {
          const row = layerPaneDescriptor.rows.find((candidate) =>
            candidate.objectId === intent.objectId
          )

          if (row) {
            startLayerPaneRename(row)
          }
        }
        return
      case 'none':
        return
    }
  }

  function focusPPTInspectorTab(tabId: PPTInspectorTabId) {
    focusPPTCanvasElementBySelectorOnNextFrame<HTMLButtonElement>({
      match: ({ element }) =>
        element.getAttribute('data-ppt-inspector-tab') === tabId,
      root: document,
      selector: '[data-ppt-inspector-tab]',
    })
  }

  function selectPPTInspectorTab(tabId: PPTInspectorTabId) {
    setActiveInspectorTabId(tabId)
  }

  function handlePPTInspectorTabKeyDown(
    tabId: PPTInspectorTabId,
    event: ReactKeyboardEvent<HTMLButtonElement>,
  ) {
    if (
      event.target !== event.currentTarget ||
      event.ctrlKey ||
      event.metaKey ||
      (event.altKey && event.key !== 'ArrowDown' && event.key !== 'ArrowUp')
    ) {
      return
    }

    const intent = getPPTCanvasTabsKeyboardIntent({
      activation: inspectorTabsDescriptor.activation,
      currentId: tabId,
      key: event.key,
      tabs: PPT_INSPECTOR_TABS,
    })

    if (intent.kind === 'none') {
      return
    }

    event.preventDefault()
    event.stopPropagation()

    if (intent.activate) {
      setActiveInspectorTabId(intent.id)
    }
    focusPPTInspectorTab(intent.id)
  }

  function handleLayerPaneRowKeyDown(
    row: PPTLayerPaneRowDescriptor,
    event: ReactKeyboardEvent<HTMLElement>,
  ) {
    if (
      event.target !== event.currentTarget ||
      event.ctrlKey ||
      event.metaKey ||
      (event.altKey && event.key !== 'ArrowDown' && event.key !== 'ArrowUp')
    ) {
      return
    }

    const intent = getSlideEditLayerPaneKeyboardIntent(layerPaneDescriptor, {
      currentObjectId: row.objectId,
      key: event.key,
      altKey: event.altKey,
      rangeAnchorObjectId: layerPaneGroupState.rangeAnchorObjectId,
      shiftKey: event.shiftKey,
    })

    if (!intent.preventDefault) {
      return
    }

    event.preventDefault()
    event.stopPropagation()
    applyLayerPaneKeyboardIntent(intent)
  }

  return (
    <aside
      aria-label="Inspector"
      className="ppt-inspector"
      data-ppt-inspector-active-tab={activeInspectorTabId}
    >
      <div
        aria-label="Inspector panels"
        className="ppt-inspector-tabs"
        data-ppt-inspector-tabs
        data-ppt-inspector-tabs-activation={inspectorTabsDescriptor.activation}
        data-ppt-inspector-tabs-keyboard={inspectorTabsDescriptor.keyboardModel}
        data-ppt-inspector-tabs-model={PPT_TABS_ROVING_FOCUS_MODEL}
        role="tablist"
      >
        {inspectorTabsDescriptor.tabs.map((tab) => {
          return (
            <button
              {...tab.attributes}
              className="ppt-inspector-tab"
              data-ppt-inspector-tab={tab.id}
              data-ppt-inspector-tab-active={tab.isActive ? 'true' : 'false'}
              key={tab.id}
              type="button"
              onClick={() => selectPPTInspectorTab(tab.id)}
              onKeyDown={(event) => handlePPTInspectorTabKeyDown(tab.id, event)}
            >
              {tab.label}
            </button>
          )
        })}
      </div>
      <section
        {...slideInspectorPanelAttributes}
        className="ppt-panel-section"
        data-ppt-inspector-surface={inspectorSurface}
        data-ppt-inspector-tabpanel="slide"
        data-ppt-inspector-tabpanel-active={activeInspectorTabId === 'slide' ? 'true' : 'false'}
        data-ppt-slide-inspector-priority={inspectorSurface === 'slide-metadata-inspector' ? 'active' : 'secondary'}
        data-ppt-slide-metadata-active-index={slideMetadataDescriptor.activeSlide.index ?? ''}
        data-ppt-slide-metadata-command-slot="command-effect"
        data-ppt-slide-metadata-field-count={slideMetadataDescriptor.fields.length}
        data-ppt-slide-metadata-inspector
        data-ppt-slide-metadata-slide-count={slideMetadataDescriptor.activeSlide.slideCount}
        data-ppt-slide-metadata-slide-id={slideMetadataDescriptor.activeSlide.slideId}
        data-ppt-slide-metadata-surface={slideMetadataDescriptor.surface}
      >
        <label className="ppt-field" {...getPPTSlideMetadataFieldData(nameMetadataField)}>
          <span>Name</span>
          <input
            data-ppt-slide-field="name"
            value={slide.name}
            onChange={(event) => onSlideNameChange(event.target.value)}
          />
        </label>
        <label className="ppt-field" {...getPPTSlideMetadataFieldData(backgroundMetadataField)}>
          <span>Background</span>
          <input
            data-ppt-slide-field="background"
            type="color"
            value={slide.background?.color ?? '#ffffff'}
            onChange={(event) => onSlideBackgroundChange(event.target.value)}
          />
        </label>
        <label className="ppt-field" {...getPPTSlideMetadataFieldData(notesMetadataField)}>
          <span>Layout</span>
          <select
            data-ppt-slide-field="layout"
            value={slideLayoutId}
            onChange={(event) => onSlideLayoutChange(event.target.value)}
          >
            {layoutDescriptors.map((layout) => (
              <option key={layout.layoutId} value={layout.layoutId}>
                {layout.name}
              </option>
            ))}
          </select>
        </label>
        <div
          className="ppt-theme-token-strip"
          data-ppt-layout-id={slideLayoutId}
          data-ppt-theme-id={slideThemeId}
        >
          {themeColorTokens.map((token) => (
            <span
              className="ppt-theme-token"
              data-ppt-theme-token={token.tokenId}
              data-ppt-theme-token-role={token.role}
              key={token.tokenId}
              style={{ background: token.value }}
              title={token.label}
            />
          ))}
        </div>
        <div
          className="ppt-layout-placeholder-list"
          data-ppt-layout-placeholder-count={layoutPlaceholders.length}
          data-ppt-layout-placeholder-hidden-count={hiddenPlaceholderCount}
          data-ppt-placeholder-visibility-command={lastPlaceholderVisibilityEffect?.payload.id}
          data-ppt-placeholder-visibility-command-placeholder={lastPlaceholderVisibilityEffect?.payload.placeholderId}
          data-ppt-placeholder-visibility-command-slide={lastPlaceholderVisibilityEffect?.payload.slideId}
          data-ppt-placeholder-visibility-command-type={lastPlaceholderVisibilityEffect?.type}
          data-ppt-placeholder-visibility-command-visible={lastPlaceholderVisibilityEffect
            ? String(lastPlaceholderVisibilityEffect.payload.isVisible)
            : undefined}
        >
          {layoutPlaceholderVisibilityDescriptors.map((placeholder) => {
            const resolvedPlaceholder = layoutPlaceholderById.get(placeholder.placeholderId)

            return (
              <div
                className="ppt-layout-placeholder"
                data-ppt-layout-placeholder={placeholder.placeholderId}
                data-ppt-placeholder-bounds={`${placeholder.bounds.x},${placeholder.bounds.y},${placeholder.bounds.w},${placeholder.bounds.h}`}
                data-ppt-placeholder-layout={resolvedPlaceholder?.layoutId}
                data-ppt-placeholder-locked={placeholder.isLocked ? 'true' : 'false'}
                data-ppt-placeholder-master={resolvedPlaceholder?.masterId}
                data-ppt-placeholder-role={placeholder.role}
                data-ppt-placeholder-slide={placeholder.slideId}
                data-ppt-placeholder-visible={placeholder.isVisible ? 'true' : 'false'}
                key={placeholder.placeholderId}
              >
                <span>{placeholder.title}</span>
                <button
                  aria-label={`${placeholder.isVisible ? 'Hide' : 'Show'} ${placeholder.title}`}
                  className="ppt-placeholder-visibility-toggle"
                  data-ppt-placeholder-visibility-toggle={placeholder.placeholderId}
                  disabled={placeholder.isLocked}
                  title={placeholder.isVisible ? 'Hide placeholder' : 'Show placeholder'}
                  type="button"
                  onClick={() => onLayoutPlaceholderVisibilityChange(
                    placeholder.placeholderId,
                    !placeholder.isVisible,
                  )}
                >
                  {placeholder.isVisible ? <Eye size={14} /> : <EyeOff size={14} />}
                </button>
              </div>
            )
          })}
        </div>
        <label className="ppt-field">
          <span>Notes</span>
          <textarea
            data-ppt-slide-field="notes"
            value={slide.notes ?? ''}
            onChange={(event) => onSlideNotesChange(event.target.value)}
          />
        </label>
        <div
          className="ppt-slide-transition-fields"
          data-ppt-slide-transition
          data-ppt-transition-advance-after={slideTransitionDescriptor.advance.afterMs ?? ''}
          data-ppt-transition-advance-on-click={slideTransitionDescriptor.advance.onClick ? 'true' : 'false'}
          data-ppt-transition-duration={slideTransitionDescriptor.durationMs}
          data-ppt-transition-model="slide-edit-slide-transition-timing"
          data-ppt-transition-slide={slideTransitionDescriptor.slideId}
          data-ppt-transition-type={slideTransitionDescriptor.type}
          data-ppt-transition-types={SLIDE_EDIT_TRANSITION_TYPES.map((type) => type.id).join(' ')}
        >
          <label className="ppt-field">
            <span>Transition</span>
            <select
              data-ppt-slide-transition-field="type"
              value={slideTransition.type}
              onChange={(event) =>
                onSlideTransitionChange('type', event.target.value as PPTSlideTransitionType)}
            >
              {SLIDE_EDIT_TRANSITION_TYPES.map((type) => (
                <option key={type.id} value={type.id}>
                  {type.label}
                </option>
              ))}
            </select>
          </label>
          <label className="ppt-field">
            <span>Duration</span>
            <input
              data-ppt-slide-transition-field="durationMs"
              max={PPT_SLIDE_TRANSITION_DURATION_MAX}
              min={0}
              step={100}
              type="number"
              value={slideTransition.durationMs}
              onChange={(event) =>
                onSlideTransitionChange('durationMs', parsePPTSlideTransitionDuration(event.target.value))}
            />
          </label>
          <label className="ppt-checkbox-field">
            <input
              checked={slideTransition.advanceOnClick}
              data-ppt-slide-transition-field="advanceOnClick"
              type="checkbox"
              onChange={(event) =>
                onSlideTransitionChange('advanceOnClick', event.target.checked)}
            />
            <span>On click</span>
          </label>
          <label className="ppt-field">
            <span>After</span>
            <input
              data-ppt-slide-transition-field="advanceAfterMs"
              max={PPT_SLIDE_TRANSITION_ADVANCE_AFTER_MAX}
              min={0}
              placeholder="none"
              step={500}
              type="number"
              value={slideTransition.advanceAfterMs ?? ''}
              onChange={(event) =>
                onSlideTransitionChange('advanceAfterMs', parsePPTSlideTransitionAdvanceAfter(event.target.value))}
            />
          </label>
        </div>
        <div className="ppt-slide-metadata-readouts">
          <span
            className="ppt-slide-metadata-readout"
            data-ppt-slide-field="size"
            data-ppt-slide-metadata-value={`${slideMetadataDescriptor.metadata.size.w}x${slideMetadataDescriptor.metadata.size.h}`}
            {...getPPTSlideMetadataFieldData(sizeMetadataField)}
          >
            <span>Size</span>
            <strong>{slideMetadataDescriptor.metadata.size.w} x {slideMetadataDescriptor.metadata.size.h}</strong>
          </span>
          <span
            className="ppt-slide-metadata-readout"
            data-ppt-slide-field="orientation"
            data-ppt-slide-metadata-value={slideMetadataDescriptor.metadata.orientation}
            {...getPPTSlideMetadataFieldData(orientationMetadataField)}
          >
            <span>Orientation</span>
            <strong>{slideMetadataDescriptor.metadata.orientation}</strong>
          </span>
        </div>
      </section>

      <section
        {...selectionInspectorPanelAttributes}
        className="ppt-panel-section"
        data-ppt-inspector-surface={inspectorSurface}
        data-ppt-inspector-tabpanel="selection"
        data-ppt-inspector-tabpanel-active={activeInspectorTabId === 'selection' ? 'true' : 'false'}
        data-ppt-object-inspector
        data-ppt-object-inspector-active={selectedElement ? 'true' : 'false'}
        data-ppt-object-inspector-priority={inspectorSurface === 'object-selection-inspector' ? '0' : '1'}
      >
        {selectedElement ? (
          <>
            <label className="ppt-field">
              <span>Name</span>
              <input
                data-ppt-style-field="name"
                value={selectedElement.name}
                onChange={(event) =>
                  onElementNameChange(selectedElement.id, event.target.value)}
              />
            </label>
            <label className="ppt-field">
              <span>Opacity</span>
              <input
                data-ppt-style-field="opacity"
                data-ppt-object-opacity-attribute={objectOpacityDescriptor?.metadata.attribute}
                data-ppt-object-opacity-attribute-value={objectOpacityDescriptor?.metadata.attributeValue}
                data-ppt-object-opacity-command={objectOpacityDescriptor?.field.commandId}
                data-ppt-object-opacity-control={objectOpacityDescriptor?.field.control}
                data-ppt-object-opacity-surface={objectOpacityDescriptor?.surface}
                max={objectOpacityDescriptor?.field.max ?? PPT_ELEMENT_OPACITY_MAX}
                min={objectOpacityDescriptor?.field.min ?? PPT_ELEMENT_OPACITY_MIN}
                step={objectOpacityDescriptor?.field.step ?? PPT_ELEMENT_OPACITY_STEP}
                type="number"
                value={objectOpacityDescriptor?.value ?? getPPTElementOpacity(selectedElement)}
                onChange={(event) =>
                  onElementOpacityChange(
                    selectedElement.id,
                    parsePPTElementOpacity(event.target.value),
                  )}
              />
            </label>
            <label className="ppt-field">
              <span>Link</span>
              <input
                data-ppt-style-field="hyperlink"
                data-ppt-hyperlink-attribute={objectHyperlinkDescriptor?.metadata.attribute}
                data-ppt-hyperlink-attribute-value={objectHyperlinkDescriptor?.metadata.attributeValue}
                data-ppt-hyperlink-command={objectHyperlinkUrlField?.commandId}
                data-ppt-hyperlink-control={objectHyperlinkUrlField?.control}
                data-ppt-hyperlink-enabled={objectHyperlinkDescriptor?.metadata.isEnabled ? 'true' : 'false'}
                data-ppt-hyperlink-surface={objectHyperlinkDescriptor?.surface}
                data-ppt-hyperlink-validation={objectHyperlinkDescriptor?.metadata.validation.reason}
                maxLength={PPT_HYPERLINK_URL_MAX_LENGTH}
                placeholder="https://example.com"
                value={objectHyperlinkDescriptor?.hyperlink.url ?? elementHyperlink?.url ?? ''}
                onChange={(event) =>
                  onElementHyperlinkChange(
                    selectedElement.id,
                    event.target.value,
                  )}
              />
            </label>
            <label className="ppt-field">
              <span>Alt text</span>
              <textarea
                data-ppt-style-field="alt-text"
                data-ppt-accessibility-attribute={objectAccessibilityDescriptor?.metadata.attribute}
                data-ppt-accessibility-attribute-value={objectAccessibilityDescriptor?.metadata.attributeValue}
                data-ppt-accessibility-command={objectAccessibilityAltTextField?.commandId}
                data-ppt-accessibility-control={objectAccessibilityAltTextField?.control}
                data-ppt-accessibility-described={objectAccessibilityDescriptor?.metadata.isDescribed ? 'true' : 'false'}
                data-ppt-accessibility-surface={objectAccessibilityDescriptor?.surface}
                maxLength={PPT_ALT_TEXT_MAX_LENGTH}
                value={objectAccessibilityDescriptor?.value.altText ?? elementAltText}
                onChange={(event) =>
                  onElementAltTextChange(
                    selectedElement.id,
                    event.target.value,
                  )}
              />
            </label>
            <label className="ppt-checkbox-field">
              <input
                checked={objectShadowEnabled}
                data-ppt-shadow-attribute={objectShadowDescriptor?.metadata.attribute}
                data-ppt-shadow-attribute-value={objectShadowDescriptor?.metadata.attributeValue}
                data-ppt-shadow-command={objectShadowEnabledField?.commandId}
                data-ppt-shadow-control={objectShadowEnabledField?.control}
                data-ppt-shadow-descriptor-enabled={objectShadowDescriptor?.metadata.isEnabled ? 'true' : 'false'}
                data-ppt-shadow-field="enabled"
                data-ppt-shadow-surface={objectShadowDescriptor?.surface}
                type="checkbox"
                onChange={(event) =>
                  onElementShadowChange(
                    selectedElement.id,
                    'enabled',
                    event.target.checked,
                  )}
              />
              <span>Shadow</span>
            </label>
            <div
              className="ppt-geometry-grid"
              data-ppt-shadow-angle={objectShadowDescriptor?.shadow.angle ?? elementShadow.angle}
              data-ppt-shadow-blur={objectShadowDescriptor?.shadow.blur ?? elementShadow.blur}
              data-ppt-shadow-color={objectShadowDescriptor?.shadow.color ?? elementShadow.color}
              data-ppt-shadow-distance={objectShadowDescriptor?.shadow.distance ?? elementShadow.distance}
              data-ppt-shadow-enabled={objectShadowEnabled ? 'true' : 'false'}
              data-ppt-shadow-inspector
              data-ppt-shadow-opacity={objectShadowDescriptor?.shadow.opacity ?? elementShadow.opacity}
            >
              <label className="ppt-field">
                <span>Color</span>
                <input
                  data-ppt-shadow-attribute={objectShadowDescriptor?.metadata.attribute}
                  data-ppt-shadow-attribute-value={objectShadowDescriptor?.metadata.attributeValue}
                  data-ppt-shadow-command={objectShadowColorField?.commandId}
                  data-ppt-shadow-control={objectShadowColorField?.control}
                  data-ppt-shadow-descriptor-enabled={objectShadowDescriptor?.metadata.isEnabled ? 'true' : 'false'}
                  data-ppt-shadow-field="color"
                  data-ppt-shadow-surface={objectShadowDescriptor?.surface}
                  disabled={!objectShadowEnabled}
                  type="color"
                  value={objectShadowDescriptor?.shadow.color ?? elementShadow.color}
                  onChange={(event) =>
                    onElementShadowChange(
                      selectedElement.id,
                      'color',
                      event.target.value,
                    )}
                />
              </label>
              <label className="ppt-field">
                <span>Shadow opacity</span>
                <input
                  data-ppt-shadow-attribute={objectShadowDescriptor?.metadata.attribute}
                  data-ppt-shadow-attribute-value={objectShadowDescriptor?.metadata.attributeValue}
                  data-ppt-shadow-command={objectShadowOpacityField?.commandId}
                  data-ppt-shadow-control={objectShadowOpacityField?.control}
                  data-ppt-shadow-descriptor-enabled={objectShadowDescriptor?.metadata.isEnabled ? 'true' : 'false'}
                  data-ppt-shadow-field="opacity"
                  data-ppt-shadow-surface={objectShadowDescriptor?.surface}
                  data-ppt-shadow-unit={objectShadowOpacityField?.unit}
                  disabled={!objectShadowEnabled}
                  max={objectShadowOpacityField?.max ?? PPT_ELEMENT_SHADOW_OPACITY_MAX}
                  min={objectShadowOpacityField?.min ?? PPT_ELEMENT_SHADOW_OPACITY_MIN}
                  step={objectShadowOpacityField?.step ?? PPT_ELEMENT_SHADOW_OPACITY_STEP}
                  type="number"
                  value={objectShadowDescriptor?.shadow.opacity ?? elementShadow.opacity}
                  onChange={(event) =>
                    onElementShadowChange(
                      selectedElement.id,
                      'opacity',
                      parsePPTElementShadowOpacity(event.target.value),
                    )}
                />
              </label>
              <label className="ppt-field">
                <span>Blur</span>
                <input
                  data-ppt-shadow-attribute={objectShadowDescriptor?.metadata.attribute}
                  data-ppt-shadow-attribute-value={objectShadowDescriptor?.metadata.attributeValue}
                  data-ppt-shadow-command={objectShadowBlurField?.commandId}
                  data-ppt-shadow-control={objectShadowBlurField?.control}
                  data-ppt-shadow-descriptor-enabled={objectShadowDescriptor?.metadata.isEnabled ? 'true' : 'false'}
                  data-ppt-shadow-field="blur"
                  data-ppt-shadow-surface={objectShadowDescriptor?.surface}
                  data-ppt-shadow-unit={objectShadowBlurField?.unit}
                  disabled={!objectShadowEnabled}
                  max={objectShadowBlurField?.max ?? PPT_ELEMENT_SHADOW_BLUR_MAX}
                  min={objectShadowBlurField?.min ?? 0}
                  step={objectShadowBlurField?.step ?? 1}
                  type="number"
                  value={objectShadowDescriptor?.shadow.blur ?? elementShadow.blur}
                  onChange={(event) =>
                    onElementShadowChange(
                      selectedElement.id,
                      'blur',
                      parsePPTElementShadowBlur(event.target.value),
                    )}
                />
              </label>
              <label className="ppt-field">
                <span>Distance</span>
                <input
                  data-ppt-shadow-attribute={objectShadowDescriptor?.metadata.attribute}
                  data-ppt-shadow-attribute-value={objectShadowDescriptor?.metadata.attributeValue}
                  data-ppt-shadow-command={objectShadowDistanceField?.commandId}
                  data-ppt-shadow-control={objectShadowDistanceField?.control}
                  data-ppt-shadow-descriptor-enabled={objectShadowDescriptor?.metadata.isEnabled ? 'true' : 'false'}
                  data-ppt-shadow-field="distance"
                  data-ppt-shadow-surface={objectShadowDescriptor?.surface}
                  data-ppt-shadow-unit={objectShadowDistanceField?.unit}
                  disabled={!objectShadowEnabled}
                  max={objectShadowDistanceField?.max ?? PPT_ELEMENT_SHADOW_DISTANCE_MAX}
                  min={objectShadowDistanceField?.min ?? 0}
                  step={objectShadowDistanceField?.step ?? 1}
                  type="number"
                  value={objectShadowDescriptor?.shadow.distance ?? elementShadow.distance}
                  onChange={(event) =>
                    onElementShadowChange(
                      selectedElement.id,
                      'distance',
                      parsePPTElementShadowDistance(event.target.value),
                    )}
                />
              </label>
              <label className="ppt-field">
                <span>Angle</span>
                <input
                  data-ppt-shadow-attribute={objectShadowDescriptor?.metadata.attribute}
                  data-ppt-shadow-attribute-value={objectShadowDescriptor?.metadata.attributeValue}
                  data-ppt-shadow-command={objectShadowAngleField?.commandId}
                  data-ppt-shadow-control={objectShadowAngleField?.control}
                  data-ppt-shadow-descriptor-enabled={objectShadowDescriptor?.metadata.isEnabled ? 'true' : 'false'}
                  data-ppt-shadow-field="angle"
                  data-ppt-shadow-surface={objectShadowDescriptor?.surface}
                  data-ppt-shadow-unit={objectShadowAngleField?.unit}
                  disabled={!objectShadowEnabled}
                  max={objectShadowAngleField?.max ?? PPT_ELEMENT_SHADOW_ANGLE_MAX}
                  min={objectShadowAngleField?.min ?? PPT_ELEMENT_SHADOW_ANGLE_MIN}
                  step={objectShadowAngleField?.step ?? 1}
                  type="number"
                  value={objectShadowDescriptor?.shadow.angle ?? elementShadow.angle}
                  onChange={(event) =>
                    onElementShadowChange(
                      selectedElement.id,
                      'angle',
                      parsePPTElementShadowAngle(event.target.value),
                    )}
                />
              </label>
            </div>
            <div className="ppt-geometry-grid">
              {(['x', 'y', 'w', 'h'] as const).map((field) => (
                <label className="ppt-field" key={field}>
                  <span>{field.toUpperCase()}</span>
                  <input
                    data-ppt-geometry-field={field}
                    type="number"
                    value={Math.round(selectedElement.geometry[field])}
                    onChange={(event) =>
                      onElementGeometryChange(
                        selectedElement.id,
                        field,
                        Number(event.target.value),
                      )}
                  />
                </label>
              ))}
              <label className="ppt-field">
                <span>ROT</span>
                <input
                  data-ppt-geometry-field="rotation"
                  type="number"
                  value={Math.round(selectedElement.geometry.rotation ?? 0)}
                  onChange={(event) =>
                    onElementRotationChange(
                      selectedElement.id,
                      Number(event.target.value),
                    )}
                />
              </label>
            </div>
            {selectedElementAnimation ? (
              <div
                className="ppt-animation-grid"
                data-ppt-object-animation-inspector
                data-ppt-animation-delay={objectAnimationDescriptor?.delayMs ?? selectedElementAnimation.delayMs}
                data-ppt-animation-duration={objectAnimationDescriptor?.durationMs ?? selectedElementAnimation.durationMs}
                data-ppt-animation-limit-delay-max={SLIDE_EDIT_OBJECT_ANIMATION_LIMITS.maxDelayMs}
                data-ppt-animation-limit-duration-max={SLIDE_EDIT_OBJECT_ANIMATION_LIMITS.maxDurationMs}
                data-ppt-animation-limit-order-max={SLIDE_EDIT_OBJECT_ANIMATION_LIMITS.maxBuildOrder}
                data-ppt-animation-model="slide-edit-object-animation"
                data-ppt-animation-order={objectAnimationDescriptor?.order ?? selectedElementAnimation.order}
                data-ppt-animation-package-trigger={objectAnimationDescriptor?.trigger}
                data-ppt-animation-package-type={objectAnimationDescriptor?.type}
                data-ppt-animation-trigger={selectedElementAnimation.trigger}
                data-ppt-animation-trigger-options={SLIDE_EDIT_OBJECT_ANIMATION_TRIGGERS
                  .map((trigger) => trigger.id).join(' ')}
                data-ppt-animation-type={selectedElementAnimation.type}
                data-ppt-animation-type-options={SLIDE_EDIT_OBJECT_ANIMATION_TYPES
                  .map((type) => type.id).join(' ')}
              >
                <label className="ppt-field">
                  <span>Animation</span>
                  <select
                    data-ppt-animation-command="update-object-animation"
                    data-ppt-animation-field="type"
                    data-ppt-animation-package-value={objectAnimationDescriptor?.type}
                    value={selectedElementAnimation.type}
                    onChange={(event) =>
                      onElementAnimationChange(
                        selectedElement.id,
                        'type',
                        event.target.value as PPTElementAnimationType,
                      )}
                  >
                    {PPT_ELEMENT_ANIMATION_TYPES.map((type) => (
                      <option key={type} value={type}>
                        {formatPPTElementAnimationType(type)}
                      </option>
                    ))}
                  </select>
                </label>
                <label className="ppt-field">
                  <span>Trigger</span>
                  <select
                    data-ppt-animation-command="update-object-animation"
                    data-ppt-animation-field="trigger"
                    data-ppt-animation-package-value={objectAnimationDescriptor?.trigger}
                    value={selectedElementAnimation.trigger}
                    onChange={(event) =>
                      onElementAnimationChange(
                        selectedElement.id,
                        'trigger',
                        event.target.value as PPTElementAnimationTrigger,
                      )}
                  >
                    {PPT_ELEMENT_ANIMATION_TRIGGERS.map((trigger) => (
                      <option key={trigger} value={trigger}>
                        {formatPPTElementAnimationTrigger(trigger)}
                      </option>
                    ))}
                  </select>
                </label>
                <label className="ppt-field">
                  <span>Duration</span>
                  <input
                    data-ppt-animation-command="update-object-animation"
                    data-ppt-animation-field="durationMs"
                    max={SLIDE_EDIT_OBJECT_ANIMATION_LIMITS.maxDurationMs}
                    min={SLIDE_EDIT_OBJECT_ANIMATION_LIMITS.minDurationMs}
                    step={100}
                    type="number"
                    value={objectAnimationDescriptor?.durationMs ?? selectedElementAnimation.durationMs}
                    onChange={(event) =>
                      onElementAnimationChange(
                        selectedElement.id,
                        'durationMs',
                        parsePPTElementAnimationTime(event.target.value),
                      )}
                  />
                </label>
                <label className="ppt-field">
                  <span>Delay</span>
                  <input
                    data-ppt-animation-command="update-object-animation"
                    data-ppt-animation-field="delayMs"
                    max={SLIDE_EDIT_OBJECT_ANIMATION_LIMITS.maxDelayMs}
                    min={SLIDE_EDIT_OBJECT_ANIMATION_LIMITS.minDelayMs}
                    step={100}
                    type="number"
                    value={objectAnimationDescriptor?.delayMs ?? selectedElementAnimation.delayMs}
                    onChange={(event) =>
                      onElementAnimationChange(
                        selectedElement.id,
                        'delayMs',
                        parsePPTElementAnimationTime(event.target.value),
                      )}
                  />
                </label>
                <label className="ppt-field">
                  <span>Order</span>
                  <input
                    data-ppt-animation-command="update-object-animation"
                    data-ppt-animation-field="order"
                    max={SLIDE_EDIT_OBJECT_ANIMATION_LIMITS.maxBuildOrder}
                    min={Math.max(1, SLIDE_EDIT_OBJECT_ANIMATION_LIMITS.minBuildOrder)}
                    step={1}
                    type="number"
                    value={objectAnimationDescriptor?.order ?? selectedElementAnimation.order}
                    onChange={(event) =>
                      onElementAnimationChange(
                        selectedElement.id,
                        'order',
                        parsePPTElementAnimationOrder(event.target.value, slide.elements.length),
                      )}
                  />
                </label>
              </div>
            ) : null}
            {isPPTTextElement(selectedElement) ? (
              <>
                <label className="ppt-field">
                  <span>Text</span>
                  <textarea
                    data-ppt-style-field="text"
                    value={readPPTText(selectedElement.textBody)}
                    onChange={(event) => onCommitText(selectedElement.id, event.target.value)}
                  />
                </label>
                <div className="ppt-geometry-grid">
                  <div className="ppt-color-control" data-ppt-color-control="text-color">
                    <label className="ppt-field">
                      <span>Text color</span>
                      <input
                        data-ppt-style-field="text-color"
                        type="color"
                        value={textStyle?.color ?? '#111827'}
                        onChange={(event) =>
                          onElementTextStyleChange(
                            selectedElement.id,
                            'color',
                            event.target.value,
                          )}
                      />
                    </label>
                    <PPTColorSwatchStrip
                      channel="text-color"
                      currentColor={textStyle?.color ?? '#111827'}
                      objectIds={[selectedElement.id]}
                      recentColors={recentColors}
                      slideId={slide.id}
                      themeColorTokens={themeColorTokens}
                      onSelect={(color, swatch) =>
                        onColorSwatchApply(selectedElement.id, 'text-color', color, swatch)}
                    />
                  </div>
                  <label className="ppt-field">
                    <span>Font size</span>
                    <input
                      data-ppt-style-field="font-size"
                      type="number"
                      value={textStyle?.fontSize ?? 24}
                      onChange={(event) =>
                        onElementTextStyleChange(
                          selectedElement.id,
                          'fontSize',
                          Number(event.target.value),
                        )}
                    />
                  </label>
                </div>
                <label className="ppt-field">
                  <span>Font</span>
                  <select
                    data-ppt-style-field="font-family"
                    data-ppt-text-font-family-command={textFontFamilyDescriptor?.field.commandId}
                    data-ppt-text-font-family-control={textFontFamilyDescriptor?.field.control}
                    data-ppt-text-font-family-fallback={textFontFamilyDescriptor?.fallbackFontFamily}
                    data-ppt-text-font-family-options={textFontFamilyDescriptor?.options
                      .map((option) => option.family).join(' ')}
                    data-ppt-text-font-family-surface={textFontFamilyDescriptor?.surface}
                    value={textFontFamilyDescriptor?.fontFamily ?? normalizePPTTextFontFamily(textStyle?.fontFamily)}
                    onChange={(event) =>
                      onElementTextStyleChange(
                        selectedElement.id,
                        'fontFamily',
                        event.target.value,
                      )}
                  >
                    {PPT_TEXT_FONT_FAMILY_OPTIONS.map((option) => (
                      <option key={option.value} value={option.value}>
                        {option.label}
                      </option>
                    ))}
                  </select>
                </label>
                <label className="ppt-field">
                  <span>Vertical</span>
                  <select
                    data-ppt-style-field="vertical-align"
                    data-ppt-text-vertical-align-attribute={textVerticalAlignmentDescriptor?.metadata.attribute}
                    data-ppt-text-vertical-align-attribute-value={textVerticalAlignmentDescriptor?.metadata.value}
                    data-ppt-text-vertical-align-command={textVerticalAlignmentDescriptor?.field.commandId}
                    data-ppt-text-vertical-align-control={textVerticalAlignmentDescriptor?.field.control}
                    data-ppt-text-vertical-align-default-value={textVerticalAlignmentDescriptor?.metadata.defaultValue}
                    data-ppt-text-vertical-align-options={textVerticalAlignmentDescriptor?.field.options
                      .map((option) => option.id).join(' ')}
                    data-ppt-text-vertical-align-surface={textVerticalAlignmentDescriptor?.surface}
                    value={textVerticalAlignmentDescriptor?.value ?? getPPTTextElementVerticalAlign(selectedElement)}
                    onChange={(event) =>
                      onElementTextStyleChange(
                        selectedElement.id,
                        'verticalAlign',
                        event.target.value,
                      )}
                  >
                    {SLIDE_EDIT_TEXT_VERTICAL_ALIGNMENT_OPTIONS.map((option) => (
                      <option key={option.id} value={option.id}>
                        {option.label}
                      </option>
                    ))}
                  </select>
                </label>
                <div
                  className="ppt-paragraph-spacing-grid"
                  data-ppt-text-inset-attribute={textFrameInsetDescriptor?.metadata.attribute}
                  data-ppt-text-inset-attribute-value={textFrameInsetDescriptor?.metadata.value}
                  data-ppt-text-inset-bottom={textFrameInsetDescriptor?.inset.bottom ?? textInset.bottom}
                  data-ppt-text-inset-default-value={textFrameInsetDescriptor?.metadata.defaultValue}
                  data-ppt-text-inset-inspector
                  data-ppt-text-inset-left={textFrameInsetDescriptor?.inset.left ?? textInset.left}
                  data-ppt-text-inset-right={textFrameInsetDescriptor?.inset.right ?? textInset.right}
                  data-ppt-text-inset-surface={textFrameInsetDescriptor?.surface}
                  data-ppt-text-inset-top={textFrameInsetDescriptor?.inset.top ?? textInset.top}
                >
                  {(['top', 'right', 'bottom', 'left'] as const).map((field) => {
                    const textFrameInsetField = getPPTTextFrameInsetField(
                      textFrameInsetDescriptor,
                      field,
                    )

                    return (
                      <label className="ppt-field" key={field}>
                        <span>{field[0].toUpperCase() + field.slice(1)}</span>
                        <input
                          data-ppt-text-inset-command={textFrameInsetField?.commandId}
                          data-ppt-text-inset-control={textFrameInsetField?.control}
                          data-ppt-text-inset-field={field}
                          data-ppt-text-inset-unit={textFrameInsetField?.unit}
                          max={textFrameInsetField?.max ?? PPT_TEXT_INSET_MAX}
                          min={textFrameInsetField?.min ?? PPT_TEXT_INSET_MIN}
                          step={textFrameInsetField?.step ?? PPT_TEXT_INSET_STEP}
                          type="number"
                          value={textFrameInsetDescriptor?.inset[field] ?? textInset[field]}
                          onChange={(event) =>
                            onElementTextInsetChange(
                              selectedElement.id,
                              field,
                              parsePPTTextInset(event.target.value),
                            )}
                        />
                      </label>
                    )
                  })}
                </div>
                <label className="ppt-field">
                  <span>Weight</span>
                  <select
                    data-ppt-style-field="font-weight"
                    value={textStyle?.fontWeight ?? 'regular'}
                    onChange={(event) =>
                      onElementTextStyleChange(
                        selectedElement.id,
                        'fontWeight',
                        event.target.value,
                      )}
                  >
                    <option value="regular">Regular</option>
                    <option value="semibold">Semibold</option>
                    <option value="bold">Bold</option>
                  </select>
                </label>
                <div className="ppt-field">
                  <span>Paragraph</span>
                  <div className="ppt-paragraph-control-row">
                    <button
                      aria-pressed={paragraphBullet}
                      className="ppt-paragraph-bullet-button"
                      data-ppt-paragraph-bullet
                      type="button"
                      onClick={() =>
                        onParagraphBulletChange(selectedElement.id, !paragraphBullet)}
                    >
                      bullet
                    </button>
                    <button
                      aria-pressed={paragraphNumbered}
                      className="ppt-paragraph-bullet-button"
                      data-ppt-paragraph-numbered
                      type="button"
                      onClick={() =>
                        onParagraphNumberedChange(
                          selectedElement.id,
                          !paragraphNumbered,
                        )}
                    >
                      numbered
                    </button>
                    <PPTParagraphAlignRadioGroup
                      align={paragraphAlign}
                      surface="inspector"
                      onAlignChange={(align) =>
                        onParagraphAlignChange(selectedElement.id, align)}
                    />
                  </div>
                </div>
                <div
                  className="ppt-paragraph-spacing-grid"
                  data-ppt-paragraph-spacing-inspector
                  data-ppt-paragraph-spacing-surface={paragraphSpacingDescriptor?.surface}
                  data-ppt-paragraph-line-height={paragraphSpacingDescriptor?.values.lineHeightRatio ?? paragraphSpacing.lineHeight}
                  data-ppt-paragraph-spacing-after={paragraphSpacingDescriptor?.values.paragraphAfter.value ?? paragraphSpacing.spacingAfter}
                  data-ppt-paragraph-spacing-before={paragraphSpacingDescriptor?.values.paragraphBefore.value ?? paragraphSpacing.spacingBefore}
                >
                  <label className="ppt-field">
                    <span>Line height</span>
                    <input
                      data-ppt-paragraph-control={paragraphLineHeightField?.control}
                      data-ppt-paragraph-command={paragraphLineHeightField?.commandId}
                      data-ppt-paragraph-field="lineHeight"
                      max={paragraphLineHeightField?.max ?? PPT_PARAGRAPH_LINE_HEIGHT_MAX}
                      min={paragraphLineHeightField?.min ?? PPT_PARAGRAPH_LINE_HEIGHT_MIN}
                      step={paragraphLineHeightField?.step ?? 0.05}
                      type="number"
                      value={paragraphSpacingDescriptor?.values.lineHeightRatio ?? paragraphSpacing.lineHeight}
                      onChange={(event) =>
                        onParagraphSpacingChange(
                          selectedElement.id,
                          'lineHeight',
                          parsePPTParagraphLineHeight(event.target.value),
                        )}
                    />
                  </label>
                  <label className="ppt-field">
                    <span>Before</span>
                    <input
                      data-ppt-paragraph-command={paragraphBeforeField?.commandId}
                      data-ppt-paragraph-control={paragraphBeforeField?.control}
                      data-ppt-paragraph-field="spacingBefore"
                      data-ppt-paragraph-unit={paragraphBeforeField?.unit}
                      max={paragraphBeforeField?.max ?? PPT_PARAGRAPH_SPACING_MAX}
                      min={paragraphBeforeField?.min ?? 0}
                      step={paragraphBeforeField?.step ?? 2}
                      type="number"
                      value={paragraphSpacingDescriptor?.values.paragraphBefore.value ?? paragraphSpacing.spacingBefore}
                      onChange={(event) =>
                        onParagraphSpacingChange(
                          selectedElement.id,
                          'spacingBefore',
                          parsePPTParagraphSpacing(event.target.value),
                        )}
                    />
                  </label>
                  <label className="ppt-field">
                    <span>After</span>
                    <input
                      data-ppt-paragraph-command={paragraphAfterField?.commandId}
                      data-ppt-paragraph-control={paragraphAfterField?.control}
                      data-ppt-paragraph-field="spacingAfter"
                      data-ppt-paragraph-unit={paragraphAfterField?.unit}
                      max={paragraphAfterField?.max ?? PPT_PARAGRAPH_SPACING_MAX}
                      min={paragraphAfterField?.min ?? 0}
                      step={paragraphAfterField?.step ?? 2}
                      type="number"
                      value={paragraphSpacingDescriptor?.values.paragraphAfter.value ?? paragraphSpacing.spacingAfter}
                      onChange={(event) =>
                        onParagraphSpacingChange(
                          selectedElement.id,
                          'spacingAfter',
                          parsePPTParagraphSpacing(event.target.value),
                        )}
                    />
                  </label>
                </div>
                <div
                  className="ppt-text-overflow-control"
                  data-ppt-text-autofit={selectedElement.textAutoFit}
                  data-ppt-text-autofit-command={lastTextAutoFitEffect?.payload.id}
                  data-ppt-text-autofit-command-handle={lastTextAutoFitEffect?.payload.handle}
                  data-ppt-text-autofit-command-object={lastTextAutoFitEffect?.payload.objectId}
                  data-ppt-text-autofit-command-size-mode={lastTextAutoFitEffect?.payload.sizeMode}
                  data-ppt-text-autofit-command-type={lastTextAutoFitEffect?.type}
                  data-ppt-text-autofit-model="slide-edit-text-box-auto-fit"
                  data-ppt-text-autofit-size-mode={textAutoFitIndicator?.sizeMode ?? getPPTTextAutoFitSizeMode(selectedElement)}
                  data-ppt-text-overflow={selectedTextOverflow ? 'true' : 'false'}
                  data-ppt-text-overflow-indicator-axis={textAutoFitIndicator?.overflowAxis.join(' ')}
                  data-ppt-text-overflow-indicator-visible={textAutoFitIndicator
                    ? String(textAutoFitIndicator.isVisible)
                    : undefined}
                  data-ppt-text-overflow-inspector
                >
                  <span>{selectedTextOverflow ? 'Overflow' : 'Fits'}</span>
                  <button
                    className="ppt-button"
                    data-ppt-style-action="text-auto-fit"
                    disabled={!selectedTextOverflow}
                    type="button"
                    onClick={() => onTextAutoFit(selectedElement.id)}
                  >
                    <Maximize2 size={15} /> Auto fit
                  </button>
                </div>
              </>
            ) : null}
            {selectedElement.kind === 'shape' ? (
              <>
                <label className="ppt-field">
                  <span>Shape</span>
                  <select
                    data-ppt-style-field="shape"
                    value={selectedElement.shape}
                    onChange={(event) => {
                      if (isPPTShapeKind(event.target.value)) {
                        onShapeKindChange(selectedElement.id, event.target.value)
                      }
                    }}
                  >
                    <option value="rect">Rectangle</option>
                    <option value="ellipse">Oval</option>
                    <option value="diamond">Diamond</option>
                  </select>
                </label>
                {selectedElement.shape === 'rect' ? (
                  <label className="ppt-field">
                    <span>Corner radius</span>
                    <input
                      data-ppt-style-field="shape-corner-radius"
                      data-ppt-corner-radius-attribute={cornerRadiusDescriptor?.metadata.attribute}
                      data-ppt-corner-radius-attribute-value={cornerRadiusDescriptor?.metadata.attributeValue}
                      data-ppt-corner-radius-command={cornerRadiusDescriptor?.field.commandId}
                      data-ppt-corner-radius-control={cornerRadiusDescriptor?.field.control}
                      data-ppt-corner-radius-supported={cornerRadiusDescriptor?.isSupported ? 'true' : 'false'}
                      data-ppt-corner-radius-surface={cornerRadiusDescriptor?.surface}
                      max={cornerRadiusDescriptor?.field.max ?? PPT_SHAPE_CORNER_RADIUS_MAX}
                      min={cornerRadiusDescriptor?.field.min ?? PPT_SHAPE_CORNER_RADIUS_MIN}
                      step={cornerRadiusDescriptor?.field.step ?? PPT_SHAPE_CORNER_RADIUS_STEP}
                      type="number"
                      value={cornerRadiusDescriptor?.value ?? getPPTShapeCornerRadius(selectedElement)}
                      onChange={(event) =>
                        onShapeCornerRadiusChange(
                          selectedElement.id,
                          parsePPTShapeCornerRadius(event.target.value),
                        )}
                    />
                  </label>
                ) : null}
                <div className="ppt-geometry-grid">
                  <div className="ppt-color-control" data-ppt-color-control="shape-fill">
                    <label className="ppt-field">
                      <span>Fill</span>
                      <input
                        data-ppt-style-field="fill"
                        type="color"
                        value={selectedElement.fill.color}
                        onChange={(event) =>
                          onShapeFillChange(
                            selectedElement.id,
                            'color',
                            event.target.value,
                          )}
                      />
                    </label>
                    <PPTColorSwatchStrip
                      channel="shape-fill"
                      currentColor={selectedElement.fill.color}
                      objectIds={[selectedElement.id]}
                      recentColors={recentColors}
                      slideId={slide.id}
                      themeColorTokens={themeColorTokens}
                      onSelect={(color, swatch) =>
                        onColorSwatchApply(selectedElement.id, 'shape-fill', color, swatch)}
                    />
                  </div>
                  <div className="ppt-color-control" data-ppt-color-control="shape-stroke">
                    <label className="ppt-field">
                      <span>Stroke</span>
                      <input
                        data-ppt-style-field="stroke-color"
                        type="color"
                        value={selectedElement.stroke?.color ?? '#111827'}
                        onChange={(event) =>
                          onElementStrokeChange(
                            selectedElement.id,
                            'color',
                            event.target.value,
                          )}
                      />
                    </label>
                    <PPTColorSwatchStrip
                      channel="shape-stroke"
                      currentColor={selectedElement.stroke?.color ?? '#111827'}
                      objectIds={[selectedElement.id]}
                      recentColors={recentColors}
                      slideId={slide.id}
                      themeColorTokens={themeColorTokens}
                      onSelect={(color, swatch) =>
                        onColorSwatchApply(selectedElement.id, 'shape-stroke', color, swatch)}
                    />
                  </div>
                </div>
                <label className="ppt-field">
                  <span>Fill opacity</span>
                  <input
                    data-ppt-style-field="fill-opacity"
                    data-ppt-fill-opacity-attribute={fillOpacityDescriptor?.metadata.attribute}
                    data-ppt-fill-opacity-attribute-value={fillOpacityDescriptor?.metadata.attributeValue}
                    data-ppt-fill-opacity-command={fillOpacityDescriptor?.field.commandId}
                    data-ppt-fill-opacity-control={fillOpacityDescriptor?.field.control}
                    data-ppt-fill-opacity-surface={fillOpacityDescriptor?.surface}
                    max={fillOpacityDescriptor?.field.max ?? PPT_FILL_OPACITY_MAX}
                    min={fillOpacityDescriptor?.field.min ?? PPT_FILL_OPACITY_MIN}
                    step={fillOpacityDescriptor?.field.step ?? PPT_FILL_OPACITY_STEP}
                    type="number"
                    value={fillOpacityDescriptor?.value ?? getPPTFillOpacity(selectedElement.fill)}
                    onChange={(event) =>
                      onShapeFillChange(
                        selectedElement.id,
                        'opacity',
                        parsePPTFillOpacity(event.target.value),
                      )}
                  />
                </label>
                <label className="ppt-field">
                  <span>Stroke width</span>
                  <input
                    data-ppt-style-field="stroke-width"
                    type="number"
                    value={selectedElement.stroke?.width ?? 0}
                    onChange={(event) =>
                      onElementStrokeChange(
                        selectedElement.id,
                        'width',
                        Number(event.target.value),
                      )}
                  />
                </label>
                <label className="ppt-field">
                  <span>Dash</span>
                  <select
                    data-ppt-style-field="stroke-dash"
                    data-ppt-stroke-line-style-attribute={strokeLineStyleDescriptor?.metadata.attribute}
                    data-ppt-stroke-line-style-attribute-value={strokeLineStyleDescriptor?.metadata.attributeValue}
                    data-ppt-stroke-line-style-command={strokeLineStyleDescriptor?.field.commandId}
                    data-ppt-stroke-line-style-control={strokeLineStyleDescriptor?.field.control}
                    data-ppt-stroke-line-style-surface={strokeLineStyleDescriptor?.surface}
                    value={strokeLineStyleDescriptor?.value ?? getPPTStrokeDash(selectedElement.stroke)}
                    onChange={(event) => {
                      if (isPPTStrokeDash(event.target.value)) {
                        onElementStrokeChange(
                          selectedElement.id,
                          'dash',
                          event.target.value,
                        )
                      }
                    }}
                  >
                    {(strokeLineStyleDescriptor?.field.options ?? PPT_STROKE_DASH_OPTIONS).map((option) => (
                      <option key={option.id} value={option.id}>
                        {option.label}
                      </option>
                    ))}
                  </select>
                </label>
              </>
            ) : null}
            {selectedElement.kind === 'image' ? (
              <>
                <label className="ppt-field">
                  <span>Fit</span>
                  <select
                    data-ppt-image-crop-attribute={imageCropDescriptor?.metadata.attribute}
                    data-ppt-image-crop-attribute-value={imageCropDescriptor?.metadata.attributeValue}
                    data-ppt-image-crop-command={imageCropDescriptor?.fields.fit.commandId}
                    data-ppt-image-crop-control={imageCropDescriptor?.fields.fit.control}
                    data-ppt-image-crop-field="fit"
                    data-ppt-image-crop-supported={imageCropDescriptor?.isSupported ? 'true' : 'false'}
                    data-ppt-image-crop-surface={imageCropDescriptor?.surface}
                    data-ppt-style-field="image-fit"
                    value={imageCropDescriptor?.fit ?? getPPTImageFit(selectedElement)}
                    onChange={(event) => {
                      if (isPPTImageFit(event.target.value)) {
                        onImageFitChange(selectedElement.id, event.target.value)
                      }
                    }}
                  >
                    {(imageCropDescriptor?.fields.fit.options ?? [
                      { id: 'cover', label: 'Fill' },
                      { id: 'contain', label: 'Fit' },
                    ]).map((option) => (
                      <option key={option.id} value={option.id}>
                        {option.label}
                      </option>
                    ))}
                  </select>
                </label>
                <div className="ppt-geometry-grid">
                  {(['x', 'y'] as const).map((field) => {
                    const imageCropField = imageCropDescriptor?.fields[field]

                    return (
                      <label className="ppt-field" key={field}>
                        <span>Crop {field.toUpperCase()}</span>
                        <input
                          data-ppt-image-crop-attribute={imageCropDescriptor?.metadata.attribute}
                          data-ppt-image-crop-attribute-value={imageCropDescriptor?.metadata.attributeValue}
                          data-ppt-image-crop-command={imageCropField?.commandId}
                          data-ppt-image-crop-control={imageCropField?.control}
                          data-ppt-image-crop-field={field}
                          data-ppt-image-crop-supported={imageCropDescriptor?.isSupported ? 'true' : 'false'}
                          data-ppt-image-crop-surface={imageCropDescriptor?.surface}
                          data-ppt-image-crop-unit={imageCropField?.unit}
                          data-ppt-style-field={`image-crop-${field}`}
                          max={imageCropField?.max ?? 100}
                          min={imageCropField?.min ?? 0}
                          step={imageCropField?.step ?? 1}
                          type="number"
                          value={Math.round(imageCropDescriptor?.crop[field] ?? getPPTImageCrop(selectedElement)[field])}
                          onChange={(event) =>
                            onImageCropChange(
                              selectedElement.id,
                              field,
                              Number(event.target.value),
                            )}
                        />
                      </label>
                    )
                  })}
                </div>
                <button
                  className="ppt-button"
                  data-ppt-image-crop-command={imageCropDescriptor?.fields.reset.commandId}
                  data-ppt-image-crop-control={imageCropDescriptor?.fields.reset.control}
                  data-ppt-image-crop-field="reset"
                  data-ppt-image-crop-reset
                  data-ppt-image-crop-supported={imageCropDescriptor?.isSupported ? 'true' : 'false'}
                  data-ppt-image-crop-surface={imageCropDescriptor?.surface}
                  disabled={imageCropDescriptor?.isSupported === false}
                  type="button"
                  onClick={() => onImageCropReset(selectedElement.id)}
                >
                  <Undo2 size={15} /> Reset
                </button>
                <input
                  accept={imageReplaceDescriptor?.field.accept ?? 'image/*'}
                  className="ppt-file-input"
                  data-ppt-image-replace-attribute={imageReplaceDescriptor?.metadata.attribute}
                  data-ppt-image-replace-attribute-value={imageReplaceDescriptor?.metadata.attributeValue}
                  data-ppt-image-replace-command={imageReplaceDescriptor?.field.commandId}
                  data-ppt-image-replace-control={imageReplaceDescriptor?.field.control}
                  data-ppt-image-replace-field={imageReplaceDescriptor?.field.id}
                  data-ppt-image-replace-input
                  data-ppt-image-replace-source-name={imageReplaceDescriptor?.sourceName}
                  data-ppt-image-replace-supported={imageReplaceDescriptor?.isSupported ? 'true' : 'false'}
                  data-ppt-image-replace-surface={imageReplaceDescriptor?.surface}
                  ref={imageReplaceInputRef}
                  tabIndex={-1}
                  type="file"
                  onChange={(event) => {
                    const file = getPPTImageFileFromList(event.target.files)

                    if (file) {
                      void onImageReplaceFile(selectedElement.id, file)
                    }

                    event.target.value = ''
                  }}
                />
                <button
                  className="ppt-button"
                  data-ppt-image-replace-action
                  data-ppt-image-replace-command={imageReplaceDescriptor?.field.commandId}
                  data-ppt-image-replace-control={imageReplaceDescriptor?.field.control}
                  data-ppt-image-replace-field={imageReplaceDescriptor?.field.id}
                  data-ppt-image-replace-supported={imageReplaceDescriptor?.isSupported ? 'true' : 'false'}
                  data-ppt-image-replace-surface={imageReplaceDescriptor?.surface}
                  disabled={imageReplaceDescriptor?.isSupported === false}
                  type="button"
                  onClick={() => imageReplaceInputRef.current?.click()}
                >
                  <ImagePlus size={15} /> Change
                </button>
              </>
            ) : null}
            {selectedElement.kind === 'table' ? (
              <>
                <label className="ppt-field">
                  <span>Rows</span>
                  <textarea
                    data-ppt-style-field="table-data"
                    value={stringifyPPTTableRows(selectedElement.rows)}
                    onChange={(event) =>
                      onTableRowsChange(selectedElement.id, event.target.value)}
                  />
                </label>
                <span
                  className="ppt-muted"
                  data-ppt-table-inspector-size
                >
                  {selectedElement.rows.length} x {getPPTTableColumnCount(selectedElement.rows)}
                </span>
              </>
            ) : null}
            {selectedElement.kind === 'comment' ? (
              <>
                <label className="ppt-field">
                  <span>Comment</span>
                  <textarea
                    data-ppt-style-field="comment-body"
                    maxLength={PPT_COMMENT_BODY_MAX_LENGTH}
                    value={selectedElement.body}
                    onChange={(event) =>
                      onCommentBodyChange(selectedElement.id, event.target.value)}
                  />
                </label>
                <label className="ppt-checkbox-field">
                  <input
                    checked={selectedElement.resolved === true}
                    data-ppt-style-field="comment-resolved"
                    type="checkbox"
                    onChange={(event) =>
                      onCommentResolvedChange(selectedElement.id, event.target.checked)}
                  />
                  <span>Resolved</span>
                </label>
                <section
                  className="ppt-comment-thread"
                  data-ppt-comment-thread
                  data-ppt-comment-thread-count={commentThread.length}
                  data-ppt-comment-thread-model={PPT_COMMENT_THREAD_MODEL}
                  data-ppt-comment-thread-resolved={selectedElement.resolved === true ? 'true' : 'false'}
                >
                  <div className="ppt-comment-thread-header">
                    <span>Thread</span>
                    <span data-ppt-comment-thread-count-label>{commentThread.length}</span>
                  </div>
                  <div className="ppt-comment-thread-list">
                    {commentThread.map((message, index) => (
                      <article
                        className="ppt-comment-thread-message"
                        data-ppt-comment-thread-message={message.id}
                        data-ppt-comment-thread-message-index={index}
                        key={message.id}
                      >
                        <div className="ppt-comment-thread-meta">
                          <span data-ppt-comment-thread-author>{message.authorName}</span>
                          <span data-ppt-comment-thread-created>{message.createdAt}</span>
                        </div>
                        <p data-ppt-comment-thread-body>{message.body}</p>
                      </article>
                    ))}
                  </div>
                  <div className="ppt-comment-reply-row">
                    <label className="ppt-field">
                      <span>Reply</span>
                      <textarea
                        data-ppt-comment-reply-input
                        data-ppt-style-field="comment-reply"
                        maxLength={PPT_COMMENT_REPLY_MAX_LENGTH}
                        value={commentReplyDraft}
                        onChange={(event) =>
                          updateCommentReplyDraft(selectedElement.id, event.target.value)}
                      />
                    </label>
                    <button
                      className="ppt-button"
                      data-ppt-comment-reply-add
                      disabled={commentReplyDraft.trim().length === 0}
                      type="button"
                      onClick={() => commitCommentReplyDraft(selectedElement.id)}
                    >
                      <MessageSquare size={15} /> Add
                    </button>
                  </div>
                </section>
              </>
            ) : null}
            {selectedElement.kind === 'line' || selectedElement.kind === 'freeform' ? (
              <>
                <div className="ppt-geometry-grid">
                  <div className="ppt-color-control" data-ppt-color-control="line-stroke">
                    <label className="ppt-field">
                      <span>Stroke</span>
                      <input
                        data-ppt-style-field="line-stroke-color"
                        type="color"
                        value={selectedElement.stroke.color}
                        onChange={(event) =>
                          onElementStrokeChange(
                            selectedElement.id,
                            'color',
                            event.target.value,
                          )}
                      />
                    </label>
                    <PPTColorSwatchStrip
                      channel="line-stroke"
                      currentColor={selectedElement.stroke.color}
                      objectIds={[selectedElement.id]}
                      recentColors={recentColors}
                      slideId={slide.id}
                      themeColorTokens={themeColorTokens}
                      onSelect={(color, swatch) =>
                        onColorSwatchApply(selectedElement.id, 'line-stroke', color, swatch)}
                    />
                  </div>
                  <label className="ppt-field">
                    <span>Width</span>
                    <input
                      data-ppt-style-field="line-stroke-width"
                      type="number"
                      value={selectedElement.stroke.width}
                      onChange={(event) =>
                        onElementStrokeChange(
                          selectedElement.id,
                          'width',
                          Number(event.target.value),
                        )}
                    />
                  </label>
                </div>
                <label className="ppt-field">
                  <span>Dash</span>
                  <select
                    data-ppt-style-field="line-stroke-dash"
                    data-ppt-stroke-line-style-attribute={strokeLineStyleDescriptor?.metadata.attribute}
                    data-ppt-stroke-line-style-attribute-value={strokeLineStyleDescriptor?.metadata.attributeValue}
                    data-ppt-stroke-line-style-command={strokeLineStyleDescriptor?.field.commandId}
                    data-ppt-stroke-line-style-control={strokeLineStyleDescriptor?.field.control}
                    data-ppt-stroke-line-style-surface={strokeLineStyleDescriptor?.surface}
                    value={strokeLineStyleDescriptor?.value ?? getPPTStrokeDash(selectedElement.stroke)}
                    onChange={(event) => {
                      if (isPPTStrokeDash(event.target.value)) {
                        onElementStrokeChange(
                          selectedElement.id,
                          'dash',
                          event.target.value,
                        )
                      }
                    }}
                  >
                    {(strokeLineStyleDescriptor?.field.options ?? PPT_STROKE_DASH_OPTIONS).map((option) => (
                      <option key={option.id} value={option.id}>
                        {option.label}
                      </option>
                    ))}
                  </select>
                </label>
                {selectedElement.kind === 'line' ? (
                  <>
                    <label className="ppt-field">
                      <span>Route</span>
                      <select
                        data-ppt-style-field="line-route"
                        value={selectedElement.route ?? 'straight'}
                        onChange={(event) => {
                          if (isPPTLineRoute(event.target.value)) {
                            onLineRouteChange(selectedElement.id, event.target.value)
                          }
                        }}
                      >
                        <option value="straight">Straight</option>
                        <option value="elbow">Elbow</option>
                      </select>
                    </label>
                    <div className="ppt-geometry-grid">
                      <label className="ppt-field">
                        <span>Start</span>
                        <select
                          data-ppt-style-field="line-start-marker"
                          value={selectedElement.startMarker ?? 'none'}
                          onChange={(event) => {
                            if (isPPTLineMarker(event.target.value)) {
                              onLineMarkerChange(
                                selectedElement.id,
                                'startMarker',
                                event.target.value,
                              )
                            }
                          }}
                        >
                          <option value="none">None</option>
                          <option value="arrow">Arrow</option>
                        </select>
                      </label>
                      <label className="ppt-field">
                        <span>End</span>
                        <select
                          data-ppt-style-field="line-end-marker"
                          value={selectedElement.endMarker ?? 'none'}
                          onChange={(event) => {
                            if (isPPTLineMarker(event.target.value)) {
                              onLineMarkerChange(
                                selectedElement.id,
                                'endMarker',
                                event.target.value,
                              )
                            }
                          }}
                        >
                          <option value="none">None</option>
                          <option value="arrow">Arrow</option>
                        </select>
                      </label>
                    </div>
                  </>
                ) : null}
              </>
            ) : null}
          </>
        ) : (
          <span className="ppt-muted">None</span>
        )}
      </section>

      <div className="ppt-panel-header">
        <h2>Objects</h2>
      </div>
      <section
        className="ppt-panel-section"
        data-ppt-layer-pane
        data-ppt-layer-pane-active-object-id={activeLayerPaneObjectId ?? ''}
        data-ppt-layer-pane-command-count={PPT_LAYER_PANE_COMMANDS.length}
        data-ppt-layer-pane-commands={layerPaneCommandIds}
        data-ppt-layer-pane-command-slot="command-effect"
        data-ppt-layer-pane-model="slide-edit-object-layer-pane"
        data-ppt-layer-pane-row-count={layerPaneDescriptor.rows.length}
        data-ppt-layer-pane-slide-id={layerPaneDescriptor.slideId}
      >
        <div
          aria-label="Selection pane"
          className="ppt-layer-list"
          data-ppt-layer-pane-aria-container={layerPaneDescriptor.aria.containerRole}
          data-ppt-layer-pane-aria-row={layerPaneDescriptor.aria.rowRole}
          data-ppt-layer-pane-keyboard-intent-model={SLIDE_EDIT_LAYER_PANE_KEYBOARD_INTENT_MODEL}
          data-ppt-layer-pane-keyboard-keys={SLIDE_EDIT_LAYER_PANE_KEYBOARD_KEYS}
          data-ppt-layer-pane-keyboard-model={layerPaneDescriptor.aria.keyboardModel}
          data-ppt-layer-pane-range-anchor-object-id={layerPaneGroupState.rangeAnchorObjectId ?? ''}
          data-ppt-layer-pane-range-selection-model="row-press-range-anchor"
          data-ppt-layer-pane-selection-model={layerPaneDescriptor.aria.selectionModel}
          role={layerPaneDescriptor.aria.containerRole}
        >
          {layerPaneDescriptor.rows.map((row) => {
            const activeLayerPaneRename = layerPaneRenameState?.objectId === row.objectId
              ? layerPaneRenameState
              : null
            const isRenaming = activeLayerPaneRename !== null
            const renameValue = activeLayerPaneRename?.value ?? ''
            const isDraggable = canDragLayerPaneRow(row)
            const layerPaneDropPlacement = layerPaneDragState?.dropTargetObjectId === row.objectId
              ? layerPaneDragState.dropPlacement
              : undefined
            const layerPaneDropToIndex = layerPaneDragState?.dropTargetObjectId === row.objectId
              ? layerPaneDragState.dropToIndex
              : undefined
            const objectVisibilityState = getPPTObjectVisibilityState(row)
            const stageObjectVisibilityState = getPPTObjectVisibilityState(row, 'visible-only')
            const objectVisibilityAvailability = getPPTObjectVisibilityAvailability({
              descriptors: objectVisibilityDescriptors,
              row,
            })

            return (
              <div
              aria-disabled={!row.isSelectable}
              aria-expanded={row.ariaExpanded}
              aria-level={row.ariaLevel}
              aria-posinset={row.ariaPosInSet}
              aria-selected={row.isSelected}
              aria-setsize={row.ariaSetSize}
              className="ppt-layer-row"
              data-grouped={row.isGrouped ? 'true' : 'false'}
              data-hidden={row.isHidden ? 'true' : 'false'}
              data-locked={row.isLocked ? 'true' : 'false'}
              data-ppt-layer-pane-expanded={row.ariaExpanded === undefined ? '' : String(row.ariaExpanded)}
              data-ppt-layer-pane-grouped={row.isGrouped ? 'true' : 'false'}
              data-ppt-layer-pane-hidden={row.isHidden ? 'true' : 'false'}
              data-ppt-layer-pane-draggable={isDraggable ? 'true' : 'false'}
              data-ppt-layer-pane-dragging={layerPaneDragState?.objectId === row.objectId ? 'true' : 'false'}
              data-ppt-layer-pane-drop-indicator={layerPaneDropPlacement ?? ''}
              data-ppt-layer-pane-drop-indicator-model={SLIDE_EDIT_LAYER_PANE_DROP_INDICATOR_MODEL}
              data-ppt-layer-pane-drop-target={layerPaneDropPlacement ? 'true' : 'false'}
              data-ppt-layer-pane-drop-to-index={layerPaneDropToIndex ?? ''}
              data-ppt-layer-pane-is-group={row.isGroup ? 'true' : 'false'}
              data-ppt-layer-pane-kind={row.kindLabel}
              data-ppt-layer-pane-locked={row.isLocked ? 'true' : 'false'}
              data-ppt-layer-pane-order={row.order}
              data-ppt-layer-pane-parent-object-id={row.parentObjectId ?? ''}
              data-ppt-layer-pane-renamable={row.isRenamable ? 'true' : 'false'}
              data-ppt-layer-pane-reorderable={row.isReorderable ? 'true' : 'false'}
              data-ppt-layer-pane-row={row.objectId}
              data-ppt-layer-pane-row-type={row.isGroup ? 'group' : 'object'}
              data-ppt-layer-pane-selected={row.isSelected ? 'true' : 'false'}
              data-ppt-object-visibility-hidden={objectVisibilityState.isHidden ? 'true' : 'false'}
              data-ppt-object-visibility-layer-selection-block-reason={objectVisibilityState.selectionBlockReason ?? ''}
              data-ppt-object-visibility-model="slide-edit-object-visibility"
              data-ppt-object-visibility-selectable={objectVisibilityState.isSelectable ? 'true' : 'false'}
              data-ppt-object-visibility-selection-policy="allow-hidden-selection"
              data-ppt-object-visibility-stage-selection-block-reason={stageObjectVisibilityState.selectionBlockReason ?? ''}
              data-ppt-object-visibility-visible={objectVisibilityState.isVisible ? 'true' : 'false'}
              data-ppt-layer-row={row.objectId}
              draggable={isDraggable}
              key={row.objectId}
              role={layerPaneDescriptor.aria.rowRole}
              tabIndex={row.objectId === activeLayerPaneObjectId ? 0 : -1}
              onDragEnd={handleLayerPaneRowDragEnd}
              onDragOver={(event) => handleLayerPaneRowDragOver(row, event)}
              onDragStart={(event) => handleLayerPaneRowDragStart(row, event)}
              onDrop={(event) => handleLayerPaneRowDrop(row, event)}
              onKeyDown={(event) => handleLayerPaneRowKeyDown(row, event)}
            >
              {row.isGroup ? (
                <button
                  aria-label={row.ariaExpanded ? 'Collapse group' : 'Expand group'}
                  className="ppt-layer-disclosure"
                  data-ppt-layer-pane-disclosure={row.objectId}
                  data-ppt-layer-pane-intent={row.ariaExpanded ? 'collapse-row' : 'expand-row'}
                  style={{ paddingLeft: Math.max(0, row.ariaLevel - 1) * 12 }}
                  title={row.ariaExpanded ? 'Collapse group' : 'Expand group'}
                  type="button"
                  onClick={(event) => {
                    event.stopPropagation()
                    setLayerPaneFocusedObjectId(row.objectId)
                    setLayerPaneGroupExpanded(row.objectId, row.ariaExpanded !== true)
                    focusLayerPaneRow(row.objectId)
                  }}
                >
                  {row.ariaExpanded ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
                </button>
              ) : (
                <span
                  className="ppt-layer-disclosure ppt-layer-disclosure--spacer"
                  style={{ paddingLeft: Math.max(0, row.ariaLevel - 1) * 12 }}
                />
              )}
              {isRenaming ? (
                <div className="ppt-layer-select ppt-layer-select--editing">
                  <span className="ppt-layer-kind">
                    {row.isGroup ? <Group size={14} /> : <Layers size={14} />}
                  </span>
                  <input
                    aria-label="Rename object"
                    className="ppt-layer-rename-input"
                    data-ppt-layer-pane-rename-input={row.objectId}
                    value={renameValue}
                    onBlur={() =>
                      handleLayerPaneRenameBlur(
                        row.objectId,
                        renameValue,
                      )}
                    onChange={(event) =>
                      setLayerPaneRenameState({
                        objectId: row.objectId,
                        value: event.target.value,
                      })}
                    onClick={(event) => event.stopPropagation()}
                    onDoubleClick={(event) => event.stopPropagation()}
                    onKeyDown={(event) => {
                      event.stopPropagation()
                      const intent = getPPTCanvasEditableFieldKeyboardIntent({
                        key: event.key,
                      })

                      if (intent.preventDefault) {
                        event.preventDefault()
                      }

                      if (intent.kind === 'commit') {
                        commitLayerPaneRename(
                          row.objectId,
                          renameValue,
                        )
                        return
                      }

                      if (intent.kind === 'cancel') {
                        cancelLayerPaneRename(row.objectId)
                      }
                    }}
                  />
                </div>
              ) : (
                <button
                  className="ppt-layer-select"
                  data-ppt-layer-pane-intent="row-press"
                  data-ppt-layer-select={row.objectId}
                  type="button"
                  onDoubleClick={(event) => {
                    event.stopPropagation()
                    startLayerPaneRename(row)
                  }}
                  onClick={(event) => {
                    handleLayerPaneRowPress(row, event)
                  }}
                >
                  <span className="ppt-layer-kind">
                    {row.isGroup ? <Group size={14} /> : <Layers size={14} />}
                  </span>
                  <span className="ppt-layer-name">{row.displayName}</span>
                </button>
              )}
              <div className="ppt-layer-actions">
                <button
                  aria-label={row.isHidden ? 'Show object' : 'Hide object'}
                  className="ppt-layer-icon-button"
                  data-ppt-layer-pane-intent="visibility-toggle"
                  data-ppt-layer-visibility={row.objectId}
                  data-ppt-object-visibility-availability={objectVisibilityAvailability.isAvailable ? 'true' : 'false'}
                  data-ppt-object-visibility-command={objectVisibilityAvailability.commandId}
                  data-ppt-object-visibility-targets={objectVisibilityAvailability.targetObjectIds.join(' ')}
                  data-ppt-object-visibility-unavailable={objectVisibilityAvailability.unavailableReason ?? ''}
                  title={row.isHidden ? 'Show object' : 'Hide object'}
                  type="button"
                  onClick={(event) => {
                    event.stopPropagation()
                    runLayerPaneIntent({
                      objectId: row.objectId,
                      type: 'visibility-toggle',
                    })
                  }}
                >
                  {row.isHidden ? <EyeOff size={14} /> : <Eye size={14} />}
                </button>
                <button
                  aria-label={row.isLocked ? 'Unlock object' : 'Lock object'}
                  className="ppt-layer-icon-button"
                  data-ppt-layer-lock={row.objectId}
                  data-ppt-layer-pane-intent="lock-toggle"
                  title={row.isLocked ? 'Unlock object' : 'Lock object'}
                  type="button"
                  onClick={(event) => {
                    event.stopPropagation()
                    runLayerPaneIntent({
                      objectId: row.objectId,
                      type: 'lock-toggle',
                    })
                  }}
                >
                  {row.isLocked ? <Lock size={14} /> : <Unlock size={14} />}
                </button>
              </div>
              </div>
            )
          })}
        </div>
      </section>

      <div className="ppt-panel-header">
        <h2>Export</h2>
      </div>
      <section className="ppt-panel-section">
        <div className="ppt-toolbar-group">
          <button className="ppt-button" onClick={onCopyHTML} type="button">
            <Copy size={16} /> Copy
          </button>
          <button className="ppt-button" onClick={onDownloadHTML} type="button">
            <Download size={16} /> Download
          </button>
        </div>
        <textarea className="ppt-export-code" readOnly value={exportCode} />
      </section>
    </aside>
  )
}

function pptElementStyle(element: PPTElement): CSSProperties {
  const base: CSSProperties = {
    filter: getPPTElementShadowFilter(element),
    height: element.geometry.h,
    left: element.geometry.x,
    opacity: getPPTElementOpacity(element),
    top: element.geometry.y,
    transform: getPPTElementTransform(element),
    transformOrigin: 'center',
    width: element.geometry.w,
  }

  if (
    element.kind === 'comment' ||
    element.kind === 'freeform' ||
    element.kind === 'image' ||
    element.kind === 'line' ||
    element.kind === 'table'
  ) {
    return base
  }

  return {
    ...base,
    ...pptTextStyle(element.style),
    alignItems: getPPTTextVerticalAlignCSS(getPPTTextElementVerticalAlign(element)),
    background: element.kind === 'shape' ? getPPTFillColorCSS(element.fill) : 'transparent',
        border: element.kind === 'shape' && element.stroke
      ? `${element.stroke.width}px solid ${element.stroke.color}`
      : undefined,
    borderStyle: element.kind === 'shape' && element.stroke
      ? getPPTStrokeDashBorderStyle(element.stroke)
      : undefined,
    borderRadius: element.kind === 'shape' && element.shape === 'rect'
      ? getSlideEditObjectCornerRadiusCSS(getPPTShapeCornerRadius(element))
      : undefined,
    padding: getSlideEditTextFrameInsetPaddingCSS(getPPTTextElementInset(element)),
    textAlign: getPPTElementParagraphAlign(element),
  }
}

function getPPTElementTransform(element: PPTElement) {
  return createPPTCanvasCssBoundsTransform({
    flipX: element.flipH === true,
    flipY: element.flipV === true,
    rotation: element.geometry.rotation,
  }) || undefined
}

function pptTextStyle(style: PPTTextStyle | undefined): CSSProperties {
  if (!style) {
    return {}
  }

  return {
    color: style.color,
    fontFamily: getPPTTextFontFamilyCSS(style.fontFamily),
    fontSize: style.fontSize,
    fontWeight: style.fontWeight === 'bold'
      ? 700
      : style.fontWeight === 'semibold'
        ? 600
        : 400,
    lineHeight: PPT_PARAGRAPH_LINE_HEIGHT_DEFAULT,
  }
}

function pptTextRunStyle(run: PPTRun): CSSProperties {
  return {
    color: run.color,
    fontSize: run.size,
    fontStyle: run.italic === true ? 'italic' : undefined,
    fontWeight: run.bold === true ? 700 : undefined,
    textDecoration: run.underline === true ? 'underline' : undefined,
  }
}

const PPT_TEMPORARY_PAN_BLOCKED_TARGET_SELECTORS = [
  '[data-ppt-command-palette]',
  '[data-ppt-context-menu]',
  '[data-ppt-shortcut-help]',
] as const

const PPT_WHEEL_VIEWPORT_PASSTHROUGH_TARGET_SELECTORS = [
  '[data-ppt-command-palette]',
  '[data-ppt-context-menu]',
  '[data-ppt-minimap]',
  '[data-ppt-shortcut-help]',
  '[data-ppt-wheel-passthrough="true"]',
] as const

function isPPTTemporaryPanBlockedTarget(target: EventTarget | null) {
  return isPPTCanvasControlTarget({
    extraSelectors: PPT_TEMPORARY_PAN_BLOCKED_TARGET_SELECTORS,
    target,
  })
}

function isPPTWheelViewportPassthroughTarget(target: EventTarget | null) {
  return isPPTCanvasControlTarget({
    extraSelectors: PPT_WHEEL_VIEWPORT_PASSTHROUGH_TARGET_SELECTORS,
    target,
  }) || isPPTCanvasWheelPassthroughTarget(target)
}

function getPPTInspectorPanelAttributes(
  descriptor: PPTCanvasTabsDescriptor<PPTInspectorTabId>,
  tabId: PPTInspectorTabId,
) {
  const panel = descriptor.panels.find((item) => item.id === tabId)

  if (!panel) {
    throw new Error(`Missing PPT inspector panel for tab ${tabId}`)
  }

  return panel.attributes
}

function isPPTCanvasStandardCommandIntentKind(kind: string) {
  return PPT_CANVAS_STANDARD_COMMAND_INTENT_KINDS.has(kind)
}

function noopPPTKeyboardCommandHandler() {}

function selectSameTypePPTSelection(
  elements: readonly PPTElement[],
  selection: readonly string[],
): string[] {
  return selectSameTypePPTElements({
    elements,
    selection,
  })
}

function canSelectSameTypePPTSelection(
  elements: readonly PPTElement[],
  selection: readonly string[],
) {
  return canSelectSameTypePPTElements({
    elements,
    selection,
  })
}

function canFlipPPTSelection(
  elements: readonly PPTElement[],
  selection: readonly string[],
) {
  return canFlipPPTElements({
    elements,
    isElementSelectable: isPPTFlipSelectionElement,
    selection,
  })
}

function flipPPTSelectionElements(
  elements: PPTElement[],
  selection: readonly string[],
  axis: PPTFlipAxis,
) {
  return flipPPTElements({
    axis,
    elements,
    flipElement: ({ element, pivot, reflectedBounds }) =>
      element.kind === 'line'
        ? flipPPTLineElement(element, axis, pivot)
        : flipPPTElementBounds(element, axis, reflectedBounds),
    isElementSelectable: isPPTFlipSelectionElement,
    selection,
  })
}

function flipPPTElementBounds(
  element: PPTElement,
  axis: PPTFlipAxis,
  bounds: Bounds,
): PPTElement {
  const next = updatePPTElementBounds(element, bounds)

  return axis === 'horizontal'
    ? { ...next, flipH: next.flipH !== true }
    : { ...next, flipV: next.flipV !== true }
}

function flipPPTLineElement(
  line: PPTLine,
  axis: PPTFlipAxis,
  pivot: number,
) {
  const reflect = (point: Point): Point =>
    axis === 'horizontal'
      ? { x: 2 * pivot - point.x, y: point.y }
      : { x: point.x, y: 2 * pivot - point.y }

  return buildPPTLineFromWorldEndpoints({
    end: reflect(getPPTLineEndpointPoint(line, 'end')),
    endConnection: undefined,
    line,
    start: reflect(getPPTLineEndpointPoint(line, 'start')),
    startConnection: undefined,
  })
}

function isPPTFlipSelectionElement(element: PPTElement) {
  return element.visible !== false && element.locked !== true
}

function canTidyPPTSelection(
  elements: readonly PPTElement[],
  selection: readonly string[],
) {
  return canTidyPPTElements({
    elements,
    isElementSelectable: isPPTTidySelectionElement,
    selection,
  })
}

function tidyPPTSelectionElements(
  elements: PPTElement[],
  selection: readonly string[],
) {
  return tidyPPTElements({
    elements,
    gap: PPT_TIDY_GAP,
    isElementSelectable: isPPTTidySelectionElement,
    selection,
  })
}

function isPPTTidySelectionElement(element: PPTElement) {
  return isPPTFlipSelectionElement(element) && element.kind !== 'line'
}

function getPPTCommandSurfaceGroups({
  availability,
  surface,
}: {
  availability: PPTCommandAvailability
  surface: PPTCommandSurface
}): PPTSurfaceCommandViewGroup[] {
  return PPT_COMMAND_SURFACE_GROUPS.flatMap((group) => {
    const commands = group.commands
      .filter((command) => command.surfaces.includes(surface))
      .map((command) => ({
        ...command,
        disabled: !availability[command.availability],
      }))

    return commands.length > 0
      ? [{
          commands,
          id: group.id,
        }]
      : []
  })
}

function canCopyPPTElementFormatting(element: PPTElement | undefined) {
  return element !== undefined &&
    element.locked !== true &&
    element.visible !== false &&
    createPPTStyleClipboard(element) !== null
}

function isPPTLineStyleTargetElement(
  element: PPTElement,
): element is PPTShape | PPTLine | PPTFreeform {
  return element.kind === 'shape' ||
    element.kind === 'line' ||
    element.kind === 'freeform'
}

function createPPTStyleClipboard(element: PPTElement): PPTStyleClipboard | null {
  if (element.locked === true || element.visible === false) {
    return null
  }

  const categories: PPTStyleClipboardCategory[] = ['object']
  const stroke = getPPTElementStroke(element)
  const clipboard: PPTStyleClipboard = {
    categories,
    object: {
      opacity: getPPTElementOpacity(element),
      shadow: hasPPTElementShadow(element)
        ? clonePPTElementShadow(getPPTElementShadow(element))
        : null,
    },
    sourceId: element.id,
    sourceKind: element.kind,
    type: 'slide-style-clipboard',
  }

  if (element.kind === 'shape') {
    categories.push('shape')
    clipboard.shape = {
      fill: clonePPTFill(element.fill),
      ...(element.shape === 'rect'
        ? { cornerRadius: getPPTShapeCornerRadius(element) }
        : {}),
      ...(element.stroke ? { stroke: clonePPTStroke(element.stroke) } : {}),
    }
  }

  if (stroke) {
    categories.push('stroke')
    clipboard.stroke = clonePPTStroke(stroke)
  }

  if (isPPTTextElement(element)) {
    const paragraph = element.textBody.paragraphs[0]

    categories.push('text')
    clipboard.text = clonePPTTextStyle(getPPTTextElementStyle(element))

    if (paragraph) {
      categories.push('paragraph')
      clipboard.paragraph = {
        align: paragraph.align ?? 'left',
        bullet: paragraph.bullet,
        lineHeight: getPPTParagraphLineHeight(paragraph),
        spacingAfter: getPPTParagraphSpacingAfter(paragraph),
        spacingBefore: getPPTParagraphSpacingBefore(paragraph),
      }
    }
  }

  return clipboard
}

function createPPTStyleClipboardDescriptor(
  slideId: string,
  clipboard: PPTStyleClipboard,
): PPTStyleClipboardDescriptor {
  return createSlideEditStyleClipboardDescriptor<
    string,
    string,
    PPTElement['kind'],
    PPTStyleClipboardPackageCategory,
    unknown
  >({
    categories: getSlideEditStyleClipboardCategoryDescriptors({
      categoryIds: getPPTStyleClipboardPackageCategoryIds(clipboard),
    }),
    source: {
      kind: clipboard.sourceKind,
      objectId: clipboard.sourceId,
      slideId,
    },
    styles: getPPTStyleClipboardPackageStyles(clipboard),
  })
}

function getPPTStyleClipboardPackageCategoryIds(
  clipboard: PPTStyleClipboard,
): PPTStyleClipboardPackageCategory[] {
  const categoryIds: PPTStyleClipboardPackageCategory[] = ['object-effect']

  if (clipboard.shape) {
    categoryIds.push('shape-fill')

    if (clipboard.shape.stroke) {
      categoryIds.push('shape-stroke')
    }
  }

  if (clipboard.stroke) {
    categoryIds.push('line-style')
  }

  if (clipboard.text || clipboard.paragraph) {
    categoryIds.push('text-style')
  }

  return uniquePPTCanvasValues(categoryIds)
}

function getPPTStyleClipboardPackageStyles(
  clipboard: PPTStyleClipboard,
): PPTStyleClipboardPackageStyle[] {
  const styles: PPTStyleClipboardPackageStyle[] = [{
    categoryId: 'object-effect',
    value: clipboard.object,
  }]

  if (clipboard.shape) {
    styles.push({
      categoryId: 'shape-fill',
      value: {
        ...(clipboard.shape.cornerRadius !== undefined
          ? { cornerRadius: clipboard.shape.cornerRadius }
          : {}),
        fill: clipboard.shape.fill,
      },
    })

    if (clipboard.shape.stroke) {
      styles.push({
        categoryId: 'shape-stroke',
        value: clipboard.shape.stroke,
      })
    }
  }

  if (clipboard.stroke) {
    styles.push({
      categoryId: 'line-style',
      value: clipboard.stroke,
    })
  }

  if (clipboard.text || clipboard.paragraph) {
    styles.push({
      categoryId: 'text-style',
      value: {
        ...(clipboard.paragraph ? { paragraph: clipboard.paragraph } : {}),
        ...(clipboard.text ? { text: clipboard.text } : {}),
      },
    })
  }

  return styles
}

function getPPTStyleClipboardPasteAvailability(
  slideId: string,
  elements: readonly PPTElement[],
  clipboard: PPTStyleClipboard,
) {
  return getSlideEditStyleClipboardPasteAvailability({
    clipboard: createPPTStyleClipboardDescriptor(slideId, clipboard),
    targets: getPPTStyleClipboardTargetInputs(elements),
  })
}

function getPPTStyleClipboardTargetInputs(
  elements: readonly PPTElement[],
): SlideEditStyleClipboardTargetInput<string, PPTStyleClipboardPackageCategory>[] {
  return elements.map((element) => ({
    objectId: element.id,
    supportedCategoryIds: getPPTElementSupportedStyleClipboardCategoryIds(element),
  }))
}

function getPPTElementSupportedStyleClipboardCategoryIds(
  element: PPTElement,
): PPTStyleClipboardPackageCategory[] {
  if (element.locked === true || element.visible === false) {
    return []
  }

  const categoryIds: PPTStyleClipboardPackageCategory[] = ['object-effect']

  if (element.kind === 'shape') {
    categoryIds.push('shape-fill', 'shape-stroke', 'line-style')
  } else if (element.kind === 'line' || element.kind === 'freeform') {
    categoryIds.push('line-style')
  }

  if (isPPTTextElement(element)) {
    categoryIds.push('text-style')
  }

  return categoryIds
}

function applyPPTStyleClipboardToElement(
  element: PPTElement,
  clipboard: PPTStyleClipboard,
  appliedCategoryIds: readonly PPTStyleClipboardPackageCategory[],
): PPTElement {
  if (element.locked === true || element.visible === false) {
    return element
  }

  const appliedCategories = new Set(appliedCategoryIds)

  if (appliedCategories.size === 0) {
    return element
  }

  let next: PPTElement = element

  if (appliedCategories.has('object-effect')) {
    next = {
      ...next,
      opacity: clipboard.object.opacity === 1
        ? undefined
        : clipboard.object.opacity,
      shadow: clipboard.object.shadow
        ? clonePPTElementShadow(clipboard.object.shadow)
        : undefined,
    }
  }

  if (next.kind === 'shape') {
    if (clipboard.shape) {
      if (appliedCategories.has('shape-fill')) {
        const cornerRadius = clipboard.shape.cornerRadius ?? PPT_SHAPE_CORNER_RADIUS_DEFAULT
        next = {
          ...next,
          cornerRadius: next.shape === 'rect'
            ? getPPTShapeCornerRadiusModelValue(cornerRadius)
            : undefined,
          fill: clonePPTFill(clipboard.shape.fill),
        }
      }

      if (appliedCategories.has('shape-stroke')) {
        next = {
          ...next,
          stroke: clipboard.shape.stroke
            ? clonePPTStroke(clipboard.shape.stroke)
            : undefined,
        }
      }
    } else if (clipboard.stroke && appliedCategories.has('line-style')) {
      next = {
        ...next,
        stroke: clonePPTStroke(clipboard.stroke),
      }
    }
  } else if (
    (next.kind === 'line' || next.kind === 'freeform') &&
    clipboard.stroke &&
    appliedCategories.has('line-style')
  ) {
    next = {
      ...next,
      stroke: clonePPTStroke(clipboard.stroke),
    }
  }

  if (isPPTTextElement(next) && appliedCategories.has('text-style')) {
    const textElement = next

    if (clipboard.text) {
      next = {
        ...textElement,
        style: clonePPTTextStyle(clipboard.text),
      }
    }

    if (clipboard.paragraph) {
      next = {
        ...next,
        textBody: {
          paragraphs: textElement.textBody.paragraphs.map((paragraph) => ({
            ...paragraph,
            align: clipboard.paragraph?.align,
            bullet: clipboard.paragraph?.bullet,
            lineHeight: clipboard.paragraph?.lineHeight,
            spacingAfter: clipboard.paragraph?.spacingAfter,
            spacingBefore: clipboard.paragraph?.spacingBefore,
          })),
        },
      }
    }
  }

  return next
}

function clonePPTFill(fill: PPTFill): PPTFill {
  return normalizePPTFill(fill)
}

function clonePPTStroke(stroke: PPTStroke): PPTStroke {
  return normalizePPTStroke(stroke)
}

function clonePPTElementShadow(shadow: PPTElementShadow): PPTElementShadow {
  return normalizePPTElementShadow(shadow)
}

function clonePPTTextStyle(style: PPTTextStyle): PPTTextStyle {
  return {
    ...getDefaultPPTTextStyle(),
    ...style,
    ...(style.textInset ? { textInset: { ...style.textInset } } : {}),
  }
}

function getPPTTextQuickFormatState(
  elements: readonly PPTTextElement[],
): PPTTextQuickFormatState {
  const styles = elements.map(getPPTTextElementStyle)
  const firstStyle = styles[0] ?? getDefaultPPTTextStyle()
  const firstAlign = elements[0]?.textBody.paragraphs[0]?.align ?? 'left'

  return {
    align: elements.every((element) =>
      (element.textBody.paragraphs[0]?.align ?? 'left') === firstAlign)
      ? firstAlign
      : 'left',
    bullet: areAllPPTTextElementsBulleted(elements),
    color: styles.every((style) => style.color === firstStyle.color)
      ? firstStyle.color
      : '#111827',
    fontSize: styles.every((style) => style.fontSize === firstStyle.fontSize)
      ? firstStyle.fontSize
      : Math.round(styles.reduce((sum, style) => sum + style.fontSize, 0) / styles.length),
    isBold: styles.length > 0 &&
      styles.every((style) => style.fontWeight === 'bold'),
    isItalic: areAllPPTTextRunsStyled(elements, 'italic'),
    isUnderline: areAllPPTTextRunsStyled(elements, 'underline'),
    numbered: areAllPPTTextElementsNumbered(elements),
  }
}

function areAllPPTTextElementsBulleted(elements: readonly PPTTextElement[]) {
  return areAllPPTTextElementsListed(elements, 'bullet')
}

function areAllPPTTextElementsNumbered(elements: readonly PPTTextElement[]) {
  return areAllPPTTextElementsListed(elements, 'numbered')
}

function areAllPPTTextElementsListed(
  elements: readonly PPTTextElement[],
  list: NonNullable<PPTParagraph['bullet']>,
) {
  return elements.length > 0 &&
    elements.every((element) =>
      element.textBody.paragraphs.length > 0 &&
      element.textBody.paragraphs.every((paragraph) => paragraph.bullet === list))
}

function hasPPTTextBodyBullet(body: PPTTextBody) {
  return body.paragraphs.some((paragraph) => paragraph.bullet === 'bullet')
}

function hasPPTTextBodyNumbered(body: PPTTextBody) {
  return body.paragraphs.some((paragraph) => paragraph.bullet === 'numbered')
}

function areAllPPTTextRunsStyled(
  elements: readonly PPTTextElement[],
  field: 'italic' | 'underline',
) {
  return elements.length > 0 &&
    elements.every((element) =>
      element.textBody.paragraphs.length > 0 &&
      element.textBody.paragraphs.every((paragraph) =>
        paragraph.runs.length > 0 &&
        paragraph.runs.every((run) => run[field] === true)))
}

function getPPTTextElementStyle(element: PPTTextElement): PPTTextStyle {
  return {
    ...getDefaultPPTTextStyle(),
    ...element.style,
  }
}

function getPPTImageFit(element: PPTImage): PPTImageFit {
  return element.fit ?? 'cover'
}

function normalizePPTImageFit(value: string | null | undefined): PPTImageFit {
  return normalizeSlideEditObjectImageCropFit(value)
}

function getPPTImageCrop(element: PPTImage): PPTImageCrop {
  return element.crop ?? { x: 50, y: 50 }
}

function getPPTImageCropDescriptor(
  slideId: string,
  element: PPTImage,
): SlideEditObjectImageCropDescriptor<string, string> {
  return createSlideEditObjectImageCropDescriptor({
    crop: getPPTImageCrop(element),
    fit: getPPTImageFit(element),
    objectId: element.id,
    slideId,
  })
}

function getPPTImageReplaceDescriptor(
  slideId: string,
  element: PPTImage,
): SlideEditObjectImageReplaceDescriptor<string, string> {
  return createSlideEditObjectImageReplaceDescriptor({
    isSupported: element.locked !== true && element.visible !== false,
    objectId: element.id,
    slideId,
    sourceName: element.name,
    unsupportedReason: element.locked === true ? 'locked-object' : 'unsupported-object',
  })
}

function getDefaultPPTTextStyle(): PPTTextStyle {
  return {
    color: '#111827',
    fontFamily: PPT_DEFAULT_TEXT_FONT_FAMILY,
    fontSize: 24,
    fontWeight: 'regular',
    verticalAlign: PPT_DEFAULT_TEXT_VERTICAL_ALIGN,
  }
}

function getPPTTextFontFamilyDescriptorOptions() {
  return PPT_TEXT_FONT_FAMILY_OPTIONS.map((option) => ({
    cssFontFamily: option.css,
    family: option.value,
    isDefault: option.value === PPT_DEFAULT_TEXT_FONT_FAMILY,
    label: option.label,
    source: 'host' as const,
  }))
}

function getPPTTextFontFamilyDescriptor(
  slideId: string,
  element: PPTTextElement,
): SlideEditTextFontFamilyDescriptor<string, string> {
  return createSlideEditTextFontFamilyDescriptor({
    fallbackFontFamily: PPT_DEFAULT_TEXT_FONT_FAMILY,
    fontFamily: getPPTTextElementStyle(element).fontFamily,
    objectId: element.id,
    options: getPPTTextFontFamilyDescriptorOptions(),
    slideId,
  })
}

function normalizePPTTextFontFamily(fontFamily: string | undefined) {
  return normalizeSlideEditTextFontFamily({
    fallbackFontFamily: PPT_DEFAULT_TEXT_FONT_FAMILY,
    fontFamily,
    options: getPPTTextFontFamilyDescriptorOptions(),
  })
}

function getPPTTextFontFamilyCSS(fontFamily: string | undefined) {
  return getSlideEditTextFontFamilyCSS({
    fallbackFontFamily: PPT_DEFAULT_TEXT_FONT_FAMILY,
    fontFamily,
    options: getPPTTextFontFamilyDescriptorOptions(),
  })
}

function getPPTTextElementVerticalAlign(element: PPTElement) {
  const verticalAlign = element.kind === 'shape' || element.kind === 'textBox'
    ? element.style?.verticalAlign
    : undefined

  return normalizePPTTextVerticalAlign(
    verticalAlign,
    element.kind === 'shape' ? 'middle' : PPT_DEFAULT_TEXT_VERTICAL_ALIGN,
  )
}

function getPPTTextVerticalAlignmentDescriptor(
  slideId: string,
  element: PPTTextElement,
): SlideEditTextVerticalAlignmentDescriptor<string, string> {
  return createSlideEditTextVerticalAlignmentDescriptor({
    objectId: element.id,
    slideId,
    value: getPPTTextElementVerticalAlign(element),
  })
}

function normalizePPTTextVerticalAlign(
  verticalAlign: string | undefined,
  fallback: PPTTextVerticalAlign = PPT_DEFAULT_TEXT_VERTICAL_ALIGN,
) {
  const normalizedFallback = normalizeSlideEditTextVerticalAlignment(fallback)
  const normalizedValue = normalizeSlideEditTextVerticalAlignment(verticalAlign)

  return verticalAlign === undefined ? normalizedFallback : normalizedValue
}

function getPPTTextVerticalAlignCSS(verticalAlign: string | undefined) {
  return getSlideEditTextVerticalAlignmentFlexAlignItems(
    normalizePPTTextVerticalAlign(verticalAlign),
  )
}

function getPPTTextFrameInsetDescriptor(
  slideId: string,
  element: PPTTextElement,
): SlideEditTextFrameInsetDescriptor<string, string> {
  return createSlideEditTextFrameInsetDescriptor({
    inset: getPPTTextElementInset(element),
    objectId: element.id,
    slideId,
  })
}

function getPPTTextFrameInsetField(
  descriptor: SlideEditTextFrameInsetDescriptor<string, string> | null,
  fieldId: PPTTextInsetField,
) {
  return descriptor?.fields.find((field) => field.id === fieldId)
}

function getPPTTextElementInset(element: PPTElement): PPTTextInset {
  const fallback = element.kind === 'shape'
    ? PPT_DEFAULT_SHAPE_TEXT_INSET
    : PPT_DEFAULT_TEXT_BOX_INSET
  const inset = element.kind === 'shape' || element.kind === 'textBox'
    ? element.style?.textInset
    : undefined

  return {
    bottom: normalizePPTTextInset(inset?.bottom ?? fallback.bottom),
    left: normalizePPTTextInset(inset?.left ?? fallback.left),
    right: normalizePPTTextInset(inset?.right ?? fallback.right),
    top: normalizePPTTextInset(inset?.top ?? fallback.top),
  }
}

function parsePPTTextInset(value: string) {
  return normalizePPTTextInset(Number(value))
}

function normalizePPTTextInset(value: number) {
  const finiteValue = Number.isFinite(value) ? value : 0

  return normalizeSlideEditTextFrameInsetValue(finiteValue)
}

function formatPPTTextInsetData(inset: PPTTextInset) {
  return `${inset.top},${inset.right},${inset.bottom},${inset.left}`
}

function getPPTElementOpacity(element: PPTElement) {
  return normalizePPTElementOpacity(element.opacity ?? 1)
}

function parsePPTElementOpacity(value: string) {
  return normalizePPTElementOpacity(Number(value))
}

function normalizePPTElementOpacity(value: number) {
  return normalizeSlideEditObjectOpacity(value)
}

function formatPPTElementOpacity(value: number) {
  return toSlideEditObjectOpacityAttributeValue(value)
}

function getPPTObjectOpacityDescriptor(
  slideId: string,
  element: PPTElement,
): SlideEditObjectOpacityDescriptor<string, string> {
  return createSlideEditObjectOpacityDescriptor({
    objectId: element.id,
    slideId,
    value: getPPTElementOpacity(element),
  })
}

function getPPTElementHyperlink(element: PPTElement) {
  return element.hyperlink ? normalizePPTElementHyperlink(element.hyperlink) : null
}

function getPPTObjectHyperlinkDescriptor(
  slideId: string,
  element: PPTElement,
): SlideEditObjectHyperlinkDescriptor<string, string> {
  const hyperlink = getPPTElementHyperlink(element)

  return createSlideEditObjectHyperlinkDescriptor({
    hyperlink: hyperlink
      ? {
          target: 'same-context',
          title: '',
          url: hyperlink.url,
        }
      : null,
    objectId: element.id,
    slideId,
  })
}

function getPPTElementAltText(element: PPTElement) {
  return element.accessibility
    ? normalizePPTElementAccessibility(element.accessibility)?.altText
    : undefined
}

function getPPTObjectAccessibilityDescriptor(
  slideId: string,
  element: PPTElement,
): SlideEditObjectAccessibilityDescriptor<string, string> {
  return createSlideEditObjectAccessibilityDescriptor({
    objectId: element.id,
    slideId,
    value: {
      altText: getPPTElementAltText(element) ?? '',
      decorative: false,
    },
  })
}

function getPPTElementStrokeDash(element: PPTElement) {
  const stroke = getPPTElementStroke(element)

  return stroke ? getPPTStrokeDash(stroke) : undefined
}

function getPPTShapeCornerRadius(element: PPTShape) {
  return element.shape === 'rect'
    ? normalizePPTShapeCornerRadius(element.cornerRadius ?? PPT_SHAPE_CORNER_RADIUS_DEFAULT)
    : 0
}

function getPPTShapeCornerRadiusModelValue(value: number) {
  const normalized = normalizePPTShapeCornerRadius(value)

  return normalized === PPT_SHAPE_CORNER_RADIUS_DEFAULT
    ? undefined
    : normalized
}

function parsePPTShapeCornerRadius(value: string) {
  return normalizePPTShapeCornerRadius(Number(value))
}

function normalizePPTShapeCornerRadius(value: number) {
  return normalizeSlideEditObjectCornerRadius(value)
}

function formatPPTShapeCornerRadius(value: number) {
  return toSlideEditObjectCornerRadiusAttributeValue(value)
}

function getPPTCornerRadiusDescriptor(
  slideId: string,
  element: PPTShape,
): SlideEditObjectCornerRadiusDescriptor<string, string> {
  const isSupported = element.shape === 'rect'

  return createSlideEditObjectCornerRadiusDescriptor({
    isSupported,
    objectId: element.id,
    slideId,
    unsupportedReason: isSupported ? undefined : 'unsupported-shape',
    value: getPPTShapeCornerRadius(element),
  })
}

function getPPTFillOpacityDescriptor(
  slideId: string,
  element: PPTShape,
): SlideEditObjectFillOpacityDescriptor<string, string> {
  return createSlideEditObjectFillOpacityDescriptor({
    objectId: element.id,
    slideId,
    value: getPPTFillOpacity(element.fill),
  })
}

function normalizePPTFill(fill: Partial<PPTFill>): PPTFill {
  const opacity = normalizePPTFillOpacity(fill.opacity ?? 1)
  const normalized = {
    color: typeof fill.color === 'string' && fill.color
      ? fill.color
      : '#ffffff',
  }

  return opacity === 1
    ? normalized
    : { ...normalized, opacity }
}

function getPPTFillOpacity(fill: PPTFill) {
  return normalizePPTFillOpacity(fill.opacity ?? 1)
}

function parsePPTFillOpacity(value: string) {
  return normalizePPTFillOpacity(Number(value))
}

function normalizePPTFillOpacity(value: number) {
  return normalizeSlideEditObjectFillOpacity(value)
}

function formatPPTFillOpacity(value: number) {
  return toSlideEditObjectFillOpacityAttributeValue(value)
}

function getPPTFillColorCSS(fill: PPTFill) {
  const opacity = getPPTFillOpacity(fill)

  if (opacity === 1) {
    return fill.color
  }

  return getSlideEditColorWithAlphaCSS({
    color: fill.color,
    opacity,
  })
}

function getPPTColorSwatchPackageChannel(
  channel: PPTColorSwatchChannel,
): PPTColorSwatchPackageChannel {
  return PPT_COLOR_SWATCH_CHANNEL_MAP[channel]
}

function getPPTColorSwatchChannelDescriptor(
  channel: PPTColorSwatchChannel,
) {
  const packageChannel = getPPTColorSwatchPackageChannel(channel)

  return SLIDE_EDIT_COLOR_SWATCH_CHANNELS.find((item) =>
    item.id === packageChannel) ?? {
    id: packageChannel,
    label: packageChannel,
  }
}

function getPPTColorSwatchDescriptor({
  channel,
  currentColor,
  objectIds,
  recentColors,
  slideId,
  themeColorTokens,
}: {
  channel: PPTColorSwatchChannel
  currentColor: string
  objectIds: readonly string[]
  recentColors: readonly string[]
  slideId: string
  themeColorTokens: readonly SlideEditThemeColorToken[]
}): PPTColorSwatchDescriptor {
  const selectedValue = normalizePPTSwatchColor(currentColor) ||
    normalizeSlideEditColorSwatchValue(currentColor)
  const selectedTokenId = selectedValue
    ? themeColorTokens.find((token) =>
        normalizePPTSwatchColor(token.value) === selectedValue ||
        token.value === selectedValue,
      )?.tokenId ?? null
    : null

  return createSlideEditColorSwatchPaletteDescriptor({
    channel: getPPTColorSwatchChannelDescriptor(channel),
    disabledReason: objectIds.length > 0 ? undefined : 'no-selection',
    isDisabled: objectIds.length === 0,
    objectIds,
    recentColors,
    selectedTokenId,
    selectedValue,
    slideId,
    themeColorTokens,
  })
}

function normalizePPTSwatchColor(color: string) {
  return normalizeSlideEditColorHex(color) ?? ''
}

function getPPTElementStroke(element: PPTElement): PPTStroke | null {
  if (element.kind === 'shape') {
    return element.stroke ? normalizePPTStroke(element.stroke) : null
  }

  if (element.kind === 'line' || element.kind === 'freeform') {
    return normalizePPTStroke(element.stroke)
  }

  return null
}

function getPPTStrokeLineStyleDescriptor(
  slideId: string,
  element: PPTElement,
): SlideEditObjectStrokeLineStyleDescriptor<string, string> {
  const stroke = getPPTElementStroke(element)

  return createSlideEditObjectStrokeLineStyleDescriptor({
    isSupported: Boolean(stroke),
    objectId: element.id,
    slideId,
    unsupportedReason: stroke ? undefined : 'no-stroke',
    value: getPPTStrokeDash(stroke ?? undefined),
  })
}

function normalizePPTStroke(stroke: Partial<PPTStroke>): PPTStroke {
  const dash = normalizePPTStrokeDash(stroke.dash)
  const normalized = {
    color: typeof stroke.color === 'string' && stroke.color
      ? stroke.color
      : '#111827',
    width: normalizePPTStrokeWidth(stroke.width ?? 2),
  }

  return dash === 'solid'
    ? normalized
    : { ...normalized, dash }
}

function normalizePPTStrokeWidth(value: number) {
  const finiteValue = Number.isFinite(value) ? value : 2

  return Math.max(0, Math.min(40, finiteValue))
}

function getPPTStrokeDash(stroke: PPTStroke | undefined): PPTStrokeDash {
  return normalizePPTStrokeDash(stroke?.dash)
}

function normalizePPTStrokeDash(value: unknown): PPTStrokeDash {
  return normalizeSlideEditObjectStrokeLineStyle(
    typeof value === 'string' ? value : null,
  ) as PPTStrokeDash
}

function getPPTStrokeDashBorderStyle(stroke: PPTStroke | undefined) {
  return getSlideEditObjectStrokeLineStyleBorderStyle(getPPTStrokeDash(stroke))
}

function getPPTStrokeDashArray(stroke: PPTStroke | undefined) {
  return getSlideEditObjectStrokeLineStyleDashArray({
    strokeWidth: normalizePPTStrokeWidth(stroke?.width ?? 2),
    value: getPPTStrokeDash(stroke),
  })
}

function getPPTThumbLineDashStyle(element: PPTElement): CSSProperties {
  if (element.kind !== 'line') {
    return {}
  }

  return {
    '--ppt-thumb-line-dash': getPPTStrokeDashBorderStyle(element.stroke),
  } as CSSProperties
}

function getPPTImageAltText(element: PPTImage) {
  return getPPTElementAltText(element) ?? element.alt
}

function normalizePPTElementAccessibility(
  accessibility: Partial<PPTElementAccessibility>,
): PPTElementAccessibility | null {
  const altText = normalizePPTAltText(accessibility.altText ?? '')

  return altText ? { altText } : null
}

function normalizePPTAltText(value: string) {
  return normalizeSlideEditObjectAltTextStorageValue(value, {
    maxLength: PPT_ALT_TEXT_MAX_LENGTH,
  }) ?? ''
}

function normalizePPTElementHyperlink(
  hyperlink: Partial<PPTElementHyperlink>,
): PPTElementHyperlink | null {
  const url = normalizePPTElementHyperlinkUrl(hyperlink.url ?? '')

  return url ? { url } : null
}

function normalizePPTElementHyperlinkUrl(url: string) {
  return normalizeSlideEditObjectHyperlinkStorageUrl(url, {
    blockedSchemes: ['javascript', 'data', 'vbscript'],
    maxLength: PPT_HYPERLINK_URL_MAX_LENGTH,
  }) ?? ''
}

function hasPPTElementShadow(element: PPTElement) {
  return element.shadow !== undefined
}

function getPPTObjectShadowDescriptor(
  slideId: string,
  element: PPTElement,
): SlideEditObjectShadowDescriptor<string, string> {
  return createSlideEditObjectShadowDescriptor({
    objectId: element.id,
    shadow: {
      ...getPPTElementShadow(element),
      enabled: hasPPTElementShadow(element),
    },
    slideId,
  })
}

function getPPTObjectShadowField(
  descriptor: SlideEditObjectShadowDescriptor<string, string> | null,
  fieldId: PPTElementShadowUpdateField,
) {
  return descriptor?.fields.find((field) => field.id === fieldId)
}

function getPPTElementShadow(element: PPTElement): PPTElementShadow {
  return normalizePPTElementShadow(element.shadow ?? PPT_DEFAULT_ELEMENT_SHADOW)
}

function normalizePPTElementShadow(
  shadow: Partial<PPTElementShadow>,
): PPTElementShadow {
  return {
    angle: normalizePPTElementShadowAngle(shadow.angle ?? PPT_DEFAULT_ELEMENT_SHADOW.angle),
    blur: normalizePPTElementShadowBlur(shadow.blur ?? PPT_DEFAULT_ELEMENT_SHADOW.blur),
    color: normalizePPTElementShadowColor(shadow.color ?? PPT_DEFAULT_ELEMENT_SHADOW.color),
    distance: normalizePPTElementShadowDistance(shadow.distance ?? PPT_DEFAULT_ELEMENT_SHADOW.distance),
    opacity: normalizePPTElementShadowOpacity(shadow.opacity ?? PPT_DEFAULT_ELEMENT_SHADOW.opacity),
  }
}

function parsePPTElementShadowAngle(value: string) {
  return normalizePPTElementShadowAngle(Number(value))
}

function normalizePPTElementShadowAngle(value: number) {
  const finiteValue = Number.isFinite(value) ? value : PPT_DEFAULT_ELEMENT_SHADOW.angle

  return clampPPTCanvasValue(
    Math.round(finiteValue),
    PPT_ELEMENT_SHADOW_ANGLE_MIN,
    PPT_ELEMENT_SHADOW_ANGLE_MAX,
  )
}

function parsePPTElementShadowBlur(value: string) {
  return normalizePPTElementShadowBlur(Number(value))
}

function normalizePPTElementShadowBlur(value: number) {
  const finiteValue = Number.isFinite(value) ? value : PPT_DEFAULT_ELEMENT_SHADOW.blur

  return clampPPTCanvasValue(Math.round(finiteValue), 0, PPT_ELEMENT_SHADOW_BLUR_MAX)
}

function parsePPTElementShadowDistance(value: string) {
  return normalizePPTElementShadowDistance(Number(value))
}

function normalizePPTElementShadowDistance(value: number) {
  const finiteValue = Number.isFinite(value) ? value : PPT_DEFAULT_ELEMENT_SHADOW.distance

  return clampPPTCanvasValue(Math.round(finiteValue), 0, PPT_ELEMENT_SHADOW_DISTANCE_MAX)
}

function parsePPTElementShadowOpacity(value: string) {
  return normalizePPTElementShadowOpacity(Number(value))
}

function normalizePPTElementShadowOpacity(value: number) {
  const finiteValue = Number.isFinite(value) ? value : PPT_DEFAULT_ELEMENT_SHADOW.opacity
  const clamped = clampPPTCanvasValue(
    finiteValue,
    PPT_ELEMENT_SHADOW_OPACITY_MIN,
    PPT_ELEMENT_SHADOW_OPACITY_MAX,
  )

  return Math.round(clamped * 100) / 100
}

function formatPPTElementShadowOpacity(value: number) {
  return toSlideEditObjectOpacityAttributeValue(value)
}

function normalizePPTElementShadowColor(color: string) {
  return /^#[\da-f]{6}$/i.test(color) ? color : PPT_DEFAULT_ELEMENT_SHADOW.color
}

function getPPTElementShadowFilter(element: PPTElement) {
  return getSlideEditObjectShadowFilter({
    ...getPPTElementShadow(element),
    enabled: hasPPTElementShadow(element),
  })
}

function getPPTElementParagraphAlign(element: PPTElement) {
  if (!isPPTTextElement(element)) {
    return undefined
  }

  return element.textBody?.paragraphs[0]?.align ?? 'left'
}

const PPT_CANVAS_CREATION_ADAPTER: PPTCanvasCreationAdapter<PPTElement> = {
  createArrow: () => throwUnsupportedPPTCreationTool(),
  createHighlight: () => throwUnsupportedPPTCreationTool(),
  createMarker: () => throwUnsupportedPPTCreationTool(),
  createShape: ({ bounds, id, shapeType }) =>
    createPPTShapeElement({
      bounds,
      id,
      shape: toPPTShapeKind(shapeType),
    }),
  createText: ({ id, point }) => ({
    editValue: 'New text',
    item: createPPTTextElement({
      bounds: {
        ...PPT_DEFAULT_TEXT_BOUNDS,
        x: point.x,
        y: point.y,
      },
      id,
    }),
  }),
}

function throwUnsupportedPPTCreationTool(): never {
  throw new Error('Unsupported PPT creation tool')
}

function createPPTElementFromCreationTool({
  current,
  id,
  start,
  tool,
}: {
  current: Point
  id: string
  start: Point
  tool: PPTCreationTool
}): PPTElement {
  if (tool.kind === 'text') {
    const created = createPPTCanvasText({
      adapter: PPT_CANVAS_CREATION_ADAPTER,
      createId: () => id,
      point: start,
    })
    const item = created.item

    if (item.kind !== 'textBox') {
      return item
    }

    return {
      ...item,
      geometry: getPPTCreatedTextBounds({
        currentWorld: current,
        startWorld: start,
      }),
      textBody: createPPTTextBody(created.editValue),
    }
  }

  if (tool.kind === 'comment') {
    return createPPTCommentElement({
      id,
      point: start,
    })
  }

  if (tool.kind === 'sticky') {
    return createPPTStickyElement({
      bounds: getPPTCreatedStickyBounds({
        currentWorld: current,
        startWorld: start,
      }),
      id,
    })
  }

  if (tool.kind === 'section') {
    return createPPTSectionElement({
      bounds: getPPTCreatedSectionBounds({
        currentWorld: current,
        startWorld: start,
      }),
      id,
    })
  }

  if (tool.kind === 'freeform') {
    const style = getPPTFreeformToolStyle(tool.tool)

    return createPPTFreeformElement({
      id,
      name: style.name,
      opacity: style.opacity,
      points: [start, current],
      stroke: style.stroke,
    })
  }

  return createPPTCanvasShape({
    adapter: PPT_CANVAS_CREATION_ADAPTER,
    createId: () => id,
    currentWorld: current,
    shapeType: tool.shape,
    startWorld: start,
  })
}

function createPPTCommentElement({
  id,
  point,
}: {
  id: string
  point: Point
}): PPTComment {
  const body = PPT_COMMENT_DEFAULT_BODY

  return {
    authorName: PPT_COMMENT_DEFAULT_AUTHOR,
    body,
    createdAt: PPT_COMMENT_DEFAULT_CREATED_AT,
    geometry: clampPPTCreationBounds({
      ...PPT_COMMENT_BOUNDS,
      x: point.x,
      y: point.y,
    }),
    id,
    kind: 'comment',
    name: 'Comment',
    thread: [{
      authorName: PPT_COMMENT_DEFAULT_AUTHOR,
      body,
      createdAt: PPT_COMMENT_DEFAULT_CREATED_AT,
      id: `${id}:message-1`,
    }],
  }
}

function createPPTTextElement({
  bounds,
  id,
}: {
  bounds: Bounds
  id: string
}): PPTElement {
  return {
    geometry: clampPPTCreationBounds(bounds),
    id,
    kind: 'textBox',
    name: 'Text',
    style: { color: '#111827', fontSize: 30, fontWeight: 'semibold' },
    textBody: createPPTTextBody('New text'),
  }
}

function createPPTStickyElement({
  bounds,
  id,
}: {
  bounds: Bounds
  id: string
}): PPTShape {
  return {
    cornerRadius: 12,
    fill: { color: '#fef3c7' },
    geometry: clampPPTCreationBounds(bounds),
    id,
    kind: 'shape',
    name: 'Sticky note',
    shape: 'rect',
    stroke: { color: '#f59e0b', width: 2 },
    style: {
      color: '#713f12',
      fontSize: 22,
      fontWeight: 'semibold',
      textInset: {
        bottom: 18,
        left: 18,
        right: 18,
        top: 18,
      },
      verticalAlign: 'middle',
    },
    textBody: createPPTTextBody('Sticky note'),
  }
}

function createPPTSectionElement({
  bounds,
  id,
}: {
  bounds: Bounds
  id: string
}): PPTShape {
  return {
    cornerRadius: 16,
    fill: { color: '#dbeafe', opacity: 0.16 },
    geometry: clampPPTCreationBounds(bounds),
    id,
    kind: 'shape',
    name: 'Section',
    shape: 'rect',
    stroke: { color: '#2563eb', dash: 'dash', width: 2 },
    style: {
      color: '#1e3a8a',
      fontSize: 20,
      fontWeight: 'semibold',
      textInset: {
        bottom: 18,
        left: 18,
        right: 18,
        top: 18,
      },
      verticalAlign: 'top',
    },
    textBody: createPPTTextBody('Section'),
  }
}

function createPPTShapeElement({
  bounds,
  id,
  shape,
}: {
  bounds: Bounds
  id: string
  shape: PPTShapeKind
}): PPTShape {
  const label = getPPTShapeLabel(shape)

  return {
    fill: { color: '#eef2ff' },
    geometry: clampPPTCreationBounds(bounds),
    id,
    kind: 'shape',
    name: label,
    shape,
    stroke: { color: '#6366f1', width: 2 },
    style: { color: '#312e81', fontSize: 24, fontWeight: 'semibold' },
    textBody: createPPTTextBody(label),
  }
}

function getPPTCreatedTextBounds({
  currentWorld,
  startWorld,
}: {
  currentWorld: Point
  startWorld: Point
}): Bounds {
  const bounds = normalizePPTCanvasBounds(startWorld, currentWorld)

  if (bounds.w > 6 && bounds.h > 6) {
    return clampPPTCreationBounds(bounds)
  }

  return clampPPTCreationBounds({
    ...PPT_DEFAULT_TEXT_BOUNDS,
    x: startWorld.x,
    y: startWorld.y,
  })
}

function getPPTCreatedStickyBounds({
  currentWorld,
  startWorld,
}: {
  currentWorld: Point
  startWorld: Point
}): Bounds {
  const bounds = normalizePPTCanvasBounds(startWorld, currentWorld)

  if (bounds.w > 6 && bounds.h > 6) {
    return clampPPTCreationBounds(bounds)
  }

  return clampPPTCreationBounds({
    ...PPT_STICKY_BOUNDS,
    x: startWorld.x,
    y: startWorld.y,
  })
}

function getPPTCreatedSectionBounds({
  currentWorld,
  startWorld,
}: {
  currentWorld: Point
  startWorld: Point
}): Bounds {
  const bounds = normalizePPTCanvasBounds(startWorld, currentWorld)

  if (bounds.w > 12 && bounds.h > 12) {
    return clampPPTCreationBounds(bounds)
  }

  return clampPPTCreationBounds({
    ...PPT_SECTION_BOUNDS,
    x: startWorld.x,
    y: startWorld.y,
  })
}

function clampPPTCreationBounds(bounds: Bounds): Bounds {
  return clampPPTCanvasBoundsToFrame({
    bounds,
    frame: {
      h: PPT_SLIDE_HEIGHT,
      w: PPT_SLIDE_WIDTH,
      x: 0,
      y: 0,
    },
    minHeight: 24,
    minWidth: 24,
  })
}

function toPPTShapeKind(shape: PPTCanvasCreatedShapeKind): PPTShapeKind {
  if (shape === 'ellipse' || shape === 'diamond') {
    return shape
  }

  return 'rect'
}

function getPPTToolShortcutIntent(event: KeyboardEvent): PPTCanvasKeyboardToolIntent | null {
  if (event.metaKey || event.ctrlKey || event.altKey) {
    return null
  }

  const tool = getPPTCanvasKeyboardToolShortcutIntent({
    config: PPT_CANVAS_COMMAND_CONFIG,
    customCreationTools: [],
    event,
    key: event.key.toLowerCase(),
  })

  return tool
    ? {
        kind: 'set-tool',
        preventDefault: false,
        tool,
      }
    : null
}

function arePPTCreationToolsEqual(
  left: PPTCreationTool | null,
  right: PPTCreationTool,
) {
  if (!left || left.kind !== right.kind) {
    return false
  }

  if (left.kind === 'shape' && right.kind === 'shape') {
    return left.shape === right.shape
  }

  if (left.kind === 'freeform' && right.kind === 'freeform') {
    return left.tool === right.tool
  }

  return true
}

function isPPTShapeCreationTool(
  tool: PPTCreationTool | null,
  shape: PPTShapeKind,
) {
  return tool?.kind === 'shape' && tool.shape === shape
}

function isPPTFreeformCreationTool(
  tool: PPTCreationTool | null,
  freeformTool: PPTFreeformTool,
) {
  return tool?.kind === 'freeform' && tool.tool === freeformTool
}

function getPPTCreationToolDataValue(tool: PPTCreationTool | null) {
  if (!tool) {
    return undefined
  }

  if (tool.kind === 'freeform' && tool.tool !== 'pen') {
    return tool.tool
  }

  return getPPTCreationToolIdPrefix(tool)
}

function getPPTCreationToolIdPrefix(tool: PPTCreationTool) {
  if (tool.kind === 'shape') {
    return tool.shape
  }

  if (tool.kind === 'freeform') {
    return 'freeform'
  }

  return tool.kind
}

function getPPTDeckTextMatches(deck: PPTDeck, query: string): PPTFindMatch[] {
  if (query.length === 0) {
    return []
  }

  const matches: PPTFindMatch[] = []

  deck.slides.forEach((slide, slideIndex) => {
    slide.elements.forEach((element, elementIndex) => {
      if (!isPPTTextElement(element)) {
        return
      }

      const text = readPPTText(element.textBody)
      const ranges = getPPTTextMatchRanges(text, query)

      ranges.forEach((range) => {
        matches.push({
          elementId: element.id,
          elementIndex,
          end: range.end,
          slideId: slide.id,
          slideIndex,
          start: range.start,
        })
      })
    })
  })

  return matches
}

function getPPTTextMatchRanges(text: string, query: string) {
  if (query.length === 0) {
    return []
  }

  const ranges: Array<{ end: number; start: number }> = []
  const haystack = text.toLocaleLowerCase()
  const needle = query.toLocaleLowerCase()
  let start = haystack.indexOf(needle)

  while (start >= 0) {
    const end = start + query.length
    ranges.push({ end, start })
    start = haystack.indexOf(needle, end)
  }

  return ranges
}

function replaceAllPPTElementTextMatches(
  element: PPTElement,
  query: string,
  replacement: string,
): PPTElement {
  if (!isPPTTextElement(element)) {
    return element
  }

  const ranges = getPPTTextMatchRanges(readPPTText(element.textBody), query)

  return ranges
    .slice()
    .reverse()
    .reduce<PPTElement>(
      (nextElement, range) =>
        replacePPTElementTextRange(nextElement, range.start, range.end, replacement),
      element,
    )
}

function replacePPTElementTextRange(
  element: PPTElement,
  start: number,
  end: number,
  replacement: string,
): PPTElement {
  if (!isPPTTextElement(element)) {
    return element
  }

  return {
    ...element,
    textBody: replacePPTTextBodyRange(element.textBody, start, end, replacement),
  }
}

function replacePPTTextBodyRange(
  body: PPTTextBody,
  start: number,
  end: number,
  replacement: string,
): PPTTextBody {
  const tokens = tokenizePPTTextBody(body)
  const safeStart = clampPPTCanvasValue(Math.min(start, end), 0, tokens.length)
  const safeEnd = clampPPTCanvasValue(Math.max(start, end), safeStart, tokens.length)
  const anchor = tokens[safeStart] ?? tokens[safeStart - 1] ?? tokens[0]
  const replacementTokens = createPPTTextTokens(
    replacement,
    anchor?.runStyle ?? {},
    anchor?.align ?? body.paragraphs[0]?.align,
    anchor?.bullet ?? body.paragraphs[0]?.bullet,
    anchor?.lineHeight ?? body.paragraphs[0]?.lineHeight,
    anchor?.spacingBefore ?? body.paragraphs[0]?.spacingBefore,
    anchor?.spacingAfter ?? body.paragraphs[0]?.spacingAfter,
  )

  return buildPPTTextBodyFromTokens([
    ...tokens.slice(0, safeStart),
    ...replacementTokens,
    ...tokens.slice(safeEnd),
  ], {
    align: body.paragraphs[0]?.align,
    bullet: body.paragraphs[0]?.bullet,
    lineHeight: body.paragraphs[0]?.lineHeight,
    spacingAfter: body.paragraphs[0]?.spacingAfter,
    spacingBefore: body.paragraphs[0]?.spacingBefore,
  })
}

function tokenizePPTTextBody(body: PPTTextBody): PPTTextToken[] {
  const tokens: PPTTextToken[] = []

  body.paragraphs.forEach((paragraph, paragraphIndex) => {
    paragraph.runs.forEach((run) => {
      const { text, ...runStyle } = run

      for (let index = 0; index < text.length; index += 1) {
        tokens.push({
          align: paragraph.align,
          bullet: paragraph.bullet,
          char: text[index],
          lineHeight: paragraph.lineHeight,
          runStyle,
          spacingAfter: paragraph.spacingAfter,
          spacingBefore: paragraph.spacingBefore,
        })
      }
    })

    if (paragraphIndex < body.paragraphs.length - 1) {
      tokens.push({
        align: paragraph.align,
        bullet: paragraph.bullet,
        char: '\n',
        lineHeight: paragraph.lineHeight,
        runStyle: getPPTParagraphFallbackRunStyle(paragraph),
        spacingAfter: paragraph.spacingAfter,
        spacingBefore: paragraph.spacingBefore,
      })
    }
  })

  return tokens
}

function createPPTTextTokens(
  text: string,
  runStyle: PPTTextRunStyle,
  align: PPTParagraph['align'] | undefined,
  bullet: PPTParagraph['bullet'] | undefined,
  lineHeight: PPTParagraph['lineHeight'] | undefined,
  spacingBefore: PPTParagraph['spacingBefore'] | undefined,
  spacingAfter: PPTParagraph['spacingAfter'] | undefined,
): PPTTextToken[] {
  const tokens: PPTTextToken[] = []

  for (let index = 0; index < text.length; index += 1) {
    tokens.push({
      align,
      bullet,
      char: text[index],
      lineHeight,
      runStyle,
      spacingAfter,
      spacingBefore,
    })
  }

  return tokens
}

function buildPPTTextBodyFromTokens(
  tokens: PPTTextToken[],
  fallback: Pick<
    PPTParagraph,
    'align' | 'bullet' | 'lineHeight' | 'spacingAfter' | 'spacingBefore'
  >,
): PPTTextBody {
  const paragraphs: PPTParagraph[] = []
  let currentAlign = fallback.align
  let currentBullet = fallback.bullet
  let currentLineHeight = fallback.lineHeight
  let currentSpacingAfter = fallback.spacingAfter
  let currentSpacingBefore = fallback.spacingBefore
  let currentRuns: PPTRun[] = []
  let currentRunStyle: PPTTextRunStyle | null = null
  let currentText = ''

  function flushRun() {
    if (!currentRunStyle || currentText.length === 0) {
      return
    }

    currentRuns.push({
      ...currentRunStyle,
      text: currentText,
    })
    currentRunStyle = null
    currentText = ''
  }

  function flushParagraph() {
    flushRun()
    paragraphs.push(createPPTParagraph(currentRuns, currentAlign, currentBullet, {
      lineHeight: currentLineHeight,
      spacingAfter: currentSpacingAfter,
      spacingBefore: currentSpacingBefore,
    }))
    currentRuns = []
    currentRunStyle = null
    currentText = ''
    currentAlign = fallback.align
    currentBullet = fallback.bullet
    currentLineHeight = fallback.lineHeight
    currentSpacingAfter = fallback.spacingAfter
    currentSpacingBefore = fallback.spacingBefore
  }

  tokens.forEach((token) => {
    if (token.char === '\n') {
      flushParagraph()
      currentAlign = token.align ?? fallback.align
      currentBullet = token.bullet ?? fallback.bullet
      currentLineHeight = token.lineHeight ?? fallback.lineHeight
      currentSpacingAfter = token.spacingAfter ?? fallback.spacingAfter
      currentSpacingBefore = token.spacingBefore ?? fallback.spacingBefore
      return
    }

    if (!currentRunStyle && currentText.length === 0 && currentRuns.length === 0) {
      currentAlign = token.align ?? fallback.align
      currentBullet = token.bullet ?? fallback.bullet
      currentLineHeight = token.lineHeight ?? fallback.lineHeight
      currentSpacingAfter = token.spacingAfter ?? fallback.spacingAfter
      currentSpacingBefore = token.spacingBefore ?? fallback.spacingBefore
    }

    if (!currentRunStyle || !arePPTTextRunStylesEqual(currentRunStyle, token.runStyle)) {
      flushRun()
      currentRunStyle = token.runStyle
    }

    currentText += token.char
  })

  flushParagraph()

  return {
    paragraphs,
  }
}

function createPPTParagraph(
  runs: PPTRun[],
  align: PPTParagraph['align'] | undefined,
  bullet: PPTParagraph['bullet'] | undefined,
  spacing: Pick<PPTParagraph, 'lineHeight' | 'spacingAfter' | 'spacingBefore'> = {},
): PPTParagraph {
  return {
    ...(align ? { align } : {}),
    ...(bullet ? { bullet } : {}),
    ...(spacing.lineHeight === undefined ? {} : { lineHeight: spacing.lineHeight }),
    ...(spacing.spacingAfter === undefined ? {} : { spacingAfter: spacing.spacingAfter }),
    ...(spacing.spacingBefore === undefined ? {} : { spacingBefore: spacing.spacingBefore }),
    runs: runs.length > 0 ? runs : [{ text: '' }],
  }
}

function getPPTParagraphFallbackRunStyle(paragraph: PPTParagraph): PPTTextRunStyle {
  const firstRun = paragraph.runs[0]

  if (!firstRun) {
    return {}
  }

  return {
    ...(firstRun.bold === undefined ? {} : { bold: firstRun.bold }),
    ...(firstRun.color === undefined ? {} : { color: firstRun.color }),
    ...(firstRun.italic === undefined ? {} : { italic: firstRun.italic }),
    ...(firstRun.size === undefined ? {} : { size: firstRun.size }),
    ...(firstRun.underline === undefined ? {} : { underline: firstRun.underline }),
  }
}

function arePPTTextRunStylesEqual(
  left: PPTTextRunStyle,
  right: PPTTextRunStyle,
) {
  return left.bold === right.bold &&
    left.color === right.color &&
    left.italic === right.italic &&
    left.size === right.size &&
    left.underline === right.underline
}

function getPPTLineEndpointPoint(
  line: PPTLine,
  endpoint: 'end' | 'start',
): Point {
  const local = endpoint === 'start' ? line.start : line.end

  return {
    x: line.geometry.x + local.x,
    y: line.geometry.y + local.y,
  }
}

function getPPTLineBend(line: PPTLine) {
  return clampPPTCanvasValue(line.routeBend ?? 0.5, 0.08, 0.92)
}

function getPPTLineBendPoint(line: PPTLine): Point {
  const bendX = line.start.x + (line.end.x - line.start.x) * getPPTLineBend(line)

  return {
    x: line.geometry.x + bendX,
    y: line.geometry.y + (line.start.y + line.end.y) / 2,
  }
}

function getPPTLinePath(line: PPTLine) {
  const bendX = line.start.x + (line.end.x - line.start.x) * getPPTLineBend(line)

  return [
    `M ${line.start.x} ${line.start.y}`,
    `L ${bendX} ${line.start.y}`,
    `L ${bendX} ${line.end.y}`,
    `L ${line.end.x} ${line.end.y}`,
  ].join(' ')
}

function createPPTFreeformElement({
  id,
  name,
  opacity = 1,
  points,
  stroke = { color: '#2563eb', width: 5 },
}: {
  id: string
  name: string
  opacity?: number
  points: Point[]
  stroke?: PPTFreeform['stroke']
}): PPTFreeform {
  const normalized = normalizePPTFreeformWorldPoints(points)

  return {
    geometry: normalized.geometry,
    id,
    kind: 'freeform',
    name,
    opacity,
    points: normalized.points,
    stroke,
  }
}

function getPPTFreeformToolStyle(tool: PPTFreeformTool) {
  switch (tool) {
    case 'highlight':
      return {
        name: 'Highlighter',
        opacity: 0.42,
        stroke: { color: '#fde047', width: 18 },
      } satisfies Pick<PPTFreeform, 'name' | 'opacity' | 'stroke'>
    case 'marker':
      return {
        name: 'Marker',
        opacity: 1,
        stroke: { color: '#475569', width: 4 },
      } satisfies Pick<PPTFreeform, 'name' | 'opacity' | 'stroke'>
    case 'pen':
      return {
        name: 'Freeform',
        opacity: 1,
        stroke: { color: '#2563eb', width: 5 },
      } satisfies Pick<PPTFreeform, 'name' | 'opacity' | 'stroke'>
  }
}

function normalizePPTFreeformWorldPoints(points: Point[]) {
  const normalized = normalizePPTCanvasPointsToLocalBounds({
    fallbackPoint: {
      x: PPT_SLIDE_WIDTH / 2,
      y: PPT_SLIDE_HEIGHT / 2,
    },
    frame: {
      h: PPT_SLIDE_HEIGHT,
      w: PPT_SLIDE_WIDTH,
      x: 0,
      y: 0,
    },
    minHeight: 16,
    minWidth: 16,
    padding: 8,
    points,
  })

  return {
    geometry: normalized.bounds,
    points: normalized.points,
  }
}

function appendPPTFreeformPoint(points: Point[], point: Point) {
  const next = clampPPTPointToSlide(point)
  const start = points[0] ?? next

  return getNextPPTDrawingPoints({
    currentWorld: next,
    points,
    shiftKey: false,
    startWorld: start,
  })
}

function getNextPPTEraserPoints(points: Point[], point: Point) {
  const next = clampPPTPointToSlide(point)
  const last = points.at(-1)

  if (!last) {
    return [next]
  }

  const distance = getPPTCanvasPointDistance(last, next)

  if (distance < PPT_ERASER_POINT_DISTANCE) {
    return points
  }

  const steps = Math.max(1, Math.floor(distance / PPT_ERASER_POINT_DISTANCE))
  const sampled = Array.from({ length: steps }, (_, index) => {
    const ratio = (index + 1) / steps

    return {
      x: last.x + (next.x - last.x) * ratio,
      y: last.y + (next.y - last.y) * ratio,
    }
  })

  return [...points, ...sampled]
}

function clampPPTPointToSlide(point: Point) {
  return clampPPTCanvasPointToBounds(point, {
    h: PPT_SLIDE_HEIGHT,
    w: PPT_SLIDE_WIDTH,
    x: 0,
    y: 0,
  })
}

function getPPTFreeformWorldLength(element: PPTFreeform) {
  const points = getPPTFreeformWorldPoints(element)

  return points.slice(1).reduce((length, point, index) =>
    length + getPPTCanvasPointDistance(points[index], point), 0)
}

function getPPTFreeformWorldPoints(element: PPTFreeform) {
  return element.points.map((point) => ({
    x: element.geometry.x + point.x,
    y: element.geometry.y + point.y,
  }))
}

function createPPTLineElement({
  end,
  endMarker,
  id,
  name,
  slide,
  start,
}: {
  end: Point
  endMarker: PPTLineMarker
  id: string
  name: string
  slide: PPTSlide
  start: Point
}): PPTLine {
  const startAttachment = getPPTLineAttachment(slide, start, id)
  const endAttachment = getPPTLineAttachment(slide, end, id)
  const line: PPTLine = {
    end: { x: 24, y: 12 },
    endMarker,
    geometry: { h: 24, w: 24, x: start.x, y: start.y },
    id,
    kind: 'line',
    name,
    route: 'straight',
    start: { x: 0, y: 12 },
    stroke: { color: '#111827', width: 4 },
  }

  return buildPPTLineFromWorldEndpoints({
    end: endAttachment?.point ?? end,
    endConnection: endAttachment?.connection,
    line,
    start: startAttachment?.point ?? start,
    startConnection: startAttachment?.connection,
  })
}

function updatePPTLineEndpoint(
  line: PPTLine,
  endpoint: 'end' | 'start',
  point: Point,
  slide: PPTSlide,
  lineId: string,
): PPTLine {
  const currentStart = getPPTLineEndpointPoint(line, 'start')
  const currentEnd = getPPTLineEndpointPoint(line, 'end')
  const rawPoint = clampPPTPointToSlide(point)
  const attachment = getPPTLineAttachment(slide, rawPoint, lineId)
  const nextPoint = attachment?.point ?? rawPoint
  const start = endpoint === 'start' ? nextPoint : currentStart
  const end = endpoint === 'end' ? nextPoint : currentEnd

  return buildPPTLineFromWorldEndpoints({
    end,
    endConnection: endpoint === 'end'
      ? attachment?.connection
      : line.endConnection,
    line,
    start,
    startConnection: endpoint === 'start'
      ? attachment?.connection
      : line.startConnection,
  })
}

function updatePPTLineRouteBend(line: PPTLine, point: Point): PPTLine {
  const start = getPPTLineEndpointPoint(line, 'start')
  const end = getPPTLineEndpointPoint(line, 'end')
  const span = end.x - start.x
  const fallback = getPPTLineBend(line)
  const routeBend = Math.abs(span) < 1
    ? fallback
    : clampPPTCanvasValue((point.x - start.x) / span, 0.08, 0.92)

  return {
    ...line,
    route: 'elbow',
    routeBend,
  }
}

function buildPPTLineFromWorldEndpoints({
  end,
  endConnection,
  line,
  start,
  startConnection,
}: {
  end: Point
  endConnection: PPTLineConnection | undefined
  line: PPTLine
  start: Point
  startConnection: PPTLineConnection | undefined
}): PPTLine {
  const normalized = normalizePPTCanvasPointsToLocalBounds({
    frame: {
      h: PPT_SLIDE_HEIGHT,
      w: PPT_SLIDE_WIDTH,
      x: 0,
      y: 0,
    },
    minHeight: 24,
    minWidth: 24,
    points: [start, end],
  })
  const [
    localStart = { x: 0, y: 0 },
    localEnd = { x: 0, y: 0 },
  ] = normalized.points

  const next: PPTLine = {
    ...line,
    end: localEnd,
    geometry: {
      ...line.geometry,
      ...normalized.bounds,
    },
    start: localStart,
  }

  if (endConnection) {
    next.endConnection = endConnection
  } else {
    delete next.endConnection
  }

  if (startConnection) {
    next.startConnection = startConnection
  } else {
    delete next.startConnection
  }

  return next
}

function getPPTLineAttachment(
  slide: PPTSlide,
  point: Point,
  lineId: string,
): { connection: PPTLineConnection; point: Point } | null {
  let nearest: {
    connection: PPTLineConnection
    distance: number
    point: Point
  } | null = null

  for (const element of slide.elements) {
    if (
      element.id === lineId ||
      element.visible === false ||
      (element.kind !== 'shape' && element.kind !== 'image')
    ) {
      continue
    }

    for (const anchor of getPPTConnectorAnchors(element)) {
      const distance = getPPTCanvasPointDistance(point, anchor.point)

      if (
        distance <= PPT_LINE_CONNECTION_DISTANCE &&
        (!nearest || distance < nearest.distance)
      ) {
        nearest = {
          connection: {
            anchor: anchor.anchor,
            elementId: element.id,
          },
          distance,
          point: anchor.point,
        }
      }
    }
  }

  if (!nearest) {
    return null
  }

  return {
    connection: nearest.connection,
    point: nearest.point,
  }
}

function syncPPTLineConnections(
  elements: PPTElement[],
  detachedLineIds = new Set<string>(),
) {
  const byId = new Map(elements.map((element) => [element.id, element]))

  return elements.map((element) => {
    if (element.kind !== 'line') {
      return element
    }

    if (detachedLineIds.has(element.id)) {
      return clearPPTLineConnections(element)
    }

    const startConnection = getValidPPTLineConnection(byId, element.startConnection)
    const endConnection = getValidPPTLineConnection(byId, element.endConnection)

    if (!startConnection && !endConnection) {
      return clearPPTLineConnections(element)
    }

    return buildPPTLineFromWorldEndpoints({
      end: endConnection
        ? getPPTConnectionAnchorPoint(byId.get(endConnection.elementId), endConnection)
        : getPPTLineEndpointPoint(element, 'end'),
      endConnection,
      line: element,
      start: startConnection
        ? getPPTConnectionAnchorPoint(byId.get(startConnection.elementId), startConnection)
        : getPPTLineEndpointPoint(element, 'start'),
      startConnection,
    })
  })
}

function clearPPTLineConnections(line: PPTLine): PPTLine {
  if (!line.startConnection && !line.endConnection) {
    return line
  }

  const next = { ...line }
  delete next.startConnection
  delete next.endConnection

  return next
}

function getValidPPTLineConnection(
  elements: ReadonlyMap<string, PPTElement>,
  connection: PPTLineConnection | undefined,
) {
  const target = connection ? elements.get(connection.elementId) : undefined

  if (
    !connection ||
    !target ||
    target.visible === false ||
    (target.kind !== 'shape' && target.kind !== 'image')
  ) {
    return undefined
  }

  return connection
}

function getPPTSelectedLineIds(
  elements: PPTElement[],
  selection: string[],
) {
  return new Set(
    getPPTCanvasSelectedItems({
      getItemId: (element) => element.id,
      isItemSelectable: (element) => element.kind === 'line',
      items: elements,
      selection,
    }).map((element) => element.id),
  )
}

function getPPTConnectorAnchors(element: PPTElement): Array<{
  anchor: PPTLineConnection['anchor']
  point: Point
}> {
  const points = getPPTCanvasBoundsAnchorPoints(pptGeometryToBounds(element.geometry))

  return [
    { anchor: 'left', point: points.left },
    { anchor: 'right', point: points.right },
    { anchor: 'top', point: points.top },
    { anchor: 'bottom', point: points.bottom },
    { anchor: 'center', point: points.center },
  ]
}

function getPPTConnectionAnchorPoint(
  element: PPTElement | undefined,
  connection: PPTLineConnection,
) {
  const fallback = { x: 0, y: 0 }

  if (!element) {
    return fallback
  }

  return getPPTConnectorAnchors(element)
    .find((anchor) => anchor.anchor === connection.anchor)?.point ?? fallback
}

function getPPTLineLength(line: PPTLine) {
  return getPPTCanvasPointDistance(
    getPPTLineEndpointPoint(line, 'start'),
    getPPTLineEndpointPoint(line, 'end'),
  )
}

function getPointAngle(center: Point, point: Point) {
  return Math.atan2(point.y - center.y, point.x - center.x) * 180 / Math.PI
}

function measurePPTElementAutoSize(element: PPTElement) {
  if (!isPPTTextElement(element)) {
    return null
  }

  const size = measurePPTTextContentSize(element, {
    whiteSpace: 'pre',
    width: 'max-content',
  })
  const padding = getPPTTextMeasurementPadding(element)

  return {
    h: Math.ceil(size.h + padding.y),
    w: Math.ceil(size.w + padding.x),
  }
}

function measurePPTTextAutoFitSize(element: PPTTextElement) {
  const padding = getPPTTextMeasurementPadding(element)
  const maxWidth = Math.max(24, PPT_SLIDE_WIDTH - element.geometry.x)
  const maxHeight = Math.max(24, PPT_SLIDE_HEIGHT - element.geometry.y)
  const preferred = measurePPTTextContentSize(element, {
    whiteSpace: 'pre',
    width: 'max-content',
  })
  const width = clampPPTCanvasValue(Math.ceil(preferred.w + padding.x + 8), 24, maxWidth)
  const contentWidth = Math.max(1, width - padding.x)
  const wrapped = measurePPTTextContentSize(element, {
    whiteSpace: 'pre-wrap',
    width: `${contentWidth}px`,
  })

  return {
    h: clampPPTCanvasValue(Math.ceil(wrapped.h + padding.y + 8), 24, maxHeight),
    w: width,
  }
}

function getPPTTextAutoFitCommandEffect({
  element,
  handle,
  hasOverflow,
  slideId,
}: {
  element: PPTTextElement
  handle: ResizeHandle
  hasOverflow: boolean
  slideId: string
}) {
  return getSlideEditTextAutoFitGestureCommandEffect({
    bounds: pptGeometryToBounds(element.geometry),
    handle: handle as SlideEditTextResizeHandle,
    maxBounds: {
      h: PPT_SLIDE_HEIGHT,
      w: PPT_SLIDE_WIDTH,
      x: 0,
      y: 0,
    },
    measurement: getPPTTextAutoFitMeasurement(element, hasOverflow),
    minSize: {
      h: 24,
      w: 24,
    },
    objectId: element.id,
    sizeMode: getPPTTextAutoFitSizeMode(element),
    slideId,
    type: 'resize-handle-double-click',
  })
}

function getPPTTextAutoFitIndicatorState(
  slideId: string,
  element: PPTTextElement,
  hasOverflow: boolean,
): SlideEditTextOverflowIndicatorState<string, string> {
  return getSlideEditTextOverflowIndicatorState({
    bounds: pptGeometryToBounds(element.geometry),
    measurement: getPPTTextAutoFitMeasurement(element, hasOverflow),
    objectId: element.id,
    sizeMode: getPPTTextAutoFitSizeMode(element),
    slideId,
  })
}

function getPPTTextAutoFitMeasurement(
  element: PPTTextElement,
  hasOverflow: boolean,
): SlideEditTextBoxMeasurement {
  return {
    hasOverflow,
    measuredSize: measurePPTTextAutoFitSize(element),
  }
}

function getPPTTextAutoFitSizeMode(element: PPTTextElement): SlideEditTextBoxSizeMode {
  return element.textAutoFit === PPT_TEXT_AUTOFIT ? 'resize-to-fit' : 'fixed'
}

function measurePPTTextContentSize(
  element: PPTTextElement,
  options: {
    whiteSpace: string
    width: string
  },
) {
  const style = element.style
  let numberedIndex = 0
  const size = measurePPTCanvasTextBlocks({
    blocks: element.textBody.paragraphs.map((paragraph) => {
      const paragraphStyle = getPPTParagraphStyle(paragraph)
      const listPrefix = paragraph.bullet === 'bullet'
        ? '\u2022 '
        : paragraph.bullet === 'numbered'
          ? `${numberedIndex + 1}. `
          : ''

      if (paragraph.bullet === 'numbered') {
        numberedIndex += 1
      }

      return {
        lineHeight: paragraphStyle.lineHeight,
        marginBottom: paragraphStyle.marginBottom,
        marginTop: paragraphStyle.marginTop,
        text: `${listPrefix}${
          paragraph.runs.map((run) => run.text).join('') || ' '
        }`,
      }
    }),
    style: {
      fontFamily: getPPTTextFontFamilyCSS(style?.fontFamily),
      fontSize: `${style?.fontSize ?? 24}px`,
      fontWeight: style?.fontWeight === 'bold'
        ? '700'
        : style?.fontWeight === 'semibold'
          ? '600'
          : '400',
      overflowWrap: 'anywhere',
      whiteSpace: options.whiteSpace,
      width: options.width,
    },
  })

  return size ?? { h: 0, w: 0 }
}

function getPPTTextMeasurementPadding(element: PPTTextElement) {
  const inset = getPPTTextElementInset(element)

  return {
    x: inset.left + inset.right,
    y: inset.top + inset.bottom,
  }
}

function getSpacingGuideSegmentStyle(segment: {
  end: Point
  start: Point
}): CSSProperties {
  const left = Math.min(segment.start.x, segment.end.x)
  const top = Math.min(segment.start.y, segment.end.y)
  const width = Math.abs(segment.end.x - segment.start.x)
  const height = Math.abs(segment.end.y - segment.start.y)

  return width >= height
    ? {
        left,
        top,
        width,
      }
    : {
        height,
        left,
        top,
      }
}

function getSpacingGuideLabelPoint(guide: PPTCanvasSnapGuides['spacingGuides'][number]) {
  const points = guide.segments.flatMap((segment) => [segment.start, segment.end])
  const x = points.reduce((sum, point) => sum + point.x, 0) / points.length
  const y = points.reduce((sum, point) => sum + point.y, 0) / points.length

  return { x, y }
}

function createPPTSlideId(deck: PPTDeck) {
  return createPPTCanvasSequentialIdFactory({
    existingIds: deck.slides.map((slide) => slide.id),
    startIndex: deck.slides.length + 1,
  })('slide')
}

function clonePPTDeckSlidesForImport(
  deck: PPTDeck,
  slides: readonly PPTSlide[],
) {
  const createId = createPPTCanvasSequentialIdFactory({
    existingIds: deck.slides.map((slide) => slide.id),
    startIndex: deck.slides.length + 1,
  })

  return slides.map((slide) => clonePPTSlide(slide, createId('slide')))
}

function normalizePPTCommentBody(value: string) {
  return value.slice(0, PPT_COMMENT_BODY_MAX_LENGTH)
}

function normalizePPTCommentReplyBody(value: string) {
  return value.slice(0, PPT_COMMENT_REPLY_MAX_LENGTH)
}

function normalizePPTCommentAuthorName(value: string) {
  return value.trim().slice(0, 80) || PPT_COMMENT_DEFAULT_AUTHOR
}

function normalizePPTCommentCreatedAt(value: string) {
  return value.trim().slice(0, 80) || PPT_COMMENT_DEFAULT_CREATED_AT
}

function normalizePPTCommentMessageId(value: string) {
  return value.trim().slice(0, 120) || 'json-comment:message'
}

function getPPTCommentThread(comment: PPTComment): PPTCommentThreadMessage[] {
  if (comment.thread && comment.thread.length > 0) {
    return comment.thread
  }

  return [{
    authorName: comment.authorName ?? PPT_COMMENT_DEFAULT_AUTHOR,
    body: comment.body,
    createdAt: comment.createdAt ?? PPT_COMMENT_DEFAULT_CREATED_AT,
    id: `${comment.id}:message-1`,
  }]
}

function getPPTCommentThreadWithBody(
  comment: PPTComment,
  body: string,
): PPTComment['thread'] {
  const thread = getPPTCommentThread(comment)
  const [first, ...rest] = thread

  return [{
    ...first,
    body,
  }, ...rest]
}

function applyPPTCommentImportSourceToElement(
  element: PPTComment,
  source: PPTCommentImportSource,
): PPTComment {
  const importedThread = source.comment.thread?.length
    ? source.comment.thread
    : null
  const body = source.comment.body ??
    importedThread?.[0]?.body ??
    element.body
  const thread = importedThread
    ? syncPPTCommentThreadWithBody(importedThread, body)
    : getPPTCommentThreadWithBody(element, body)

  return {
    ...element,
    ...(source.comment.createdAt ? { createdAt: source.comment.createdAt } : {}),
    ...(source.comment.resolved === undefined ? {} : { resolved: source.comment.resolved }),
    body,
    thread,
  }
}

function syncPPTCommentThreadWithBody(
  thread: readonly PPTCommentThreadMessage[],
  body: string,
): PPTCommentThreadMessage[] {
  const [first, ...rest] = thread

  return first
    ? [{ ...first, body }, ...rest.map(clonePPTCommentThreadMessage)]
    : []
}

function clonePPTCommentThreadMessage(
  message: PPTCommentThreadMessage,
): PPTCommentThreadMessage {
  return { ...message }
}

function createPPTCommentReplyMessage(
  comment: PPTComment,
  body: string,
  currentThreadLength: number,
): PPTCommentThreadMessage {
  return {
    authorName: PPT_COMMENT_DEFAULT_AUTHOR,
    body,
    createdAt: PPT_COMMENT_DEFAULT_CREATED_AT,
    id: `${comment.id}:message-${currentThreadLength + 1}`,
  }
}

function toPPTCommentThreadHostCommandEffect(
  command: PPTCommentThreadCommand,
): PPTCommentThreadHostCommandEffect {
  return {
    payload: command,
    selection: {
      objectIds: [command.objectId],
      slideId: command.slideId,
    },
    type: 'slide-command-effect',
  }
}

function clonePPTSlide(slide: PPTSlide, id: string): PPTSlide {
  const elementIdBySourceId = new Map<string, string>()
  const elements = slide.elements.map((element, index) => {
    const nextId = createPPTSlideElementId(id, element, index)

    elementIdBySourceId.set(element.id, nextId)

    return {
      ...element,
      id: nextId,
      name: element.name,
    }
  })

  return {
    ...slide,
    elements: elements.map((element) =>
      remapPPTSlideCloneElementReferences(element, elementIdBySourceId)),
    id,
    name: `${slide.name} Copy`,
  }
}

function remapPPTSlideCloneElementReferences(
  element: PPTElement,
  elementIdBySourceId: ReadonlyMap<string, string>,
): PPTElement {
  if (element.kind !== 'line') {
    return element
  }

  const startConnection = remapPPTSlideCloneLineConnection(
    element.startConnection,
    elementIdBySourceId,
  )
  const endConnection = remapPPTSlideCloneLineConnection(
    element.endConnection,
    elementIdBySourceId,
  )
  const next: PPTLine = {
    ...element,
  }

  if (startConnection) {
    next.startConnection = startConnection
  } else {
    delete next.startConnection
  }

  if (endConnection) {
    next.endConnection = endConnection
  } else {
    delete next.endConnection
  }

  return next
}

function remapPPTSlideCloneLineConnection(
  connection: PPTLineConnection | undefined,
  elementIdBySourceId: ReadonlyMap<string, string>,
): PPTLineConnection | undefined {
  if (!connection) {
    return undefined
  }

  const elementId = elementIdBySourceId.get(connection.elementId)

  return elementId
    ? {
        ...connection,
        elementId,
      }
    : undefined
}

function createPPTSlideElementId(
  slideId: string,
  element: PPTElement,
  index: number,
) {
  const name = element.name
    .trim()
    .toLowerCase()
    .replaceAll(/[^a-z0-9]+/g, '-')
    .replaceAll(/^-|-$/g, '')
  const suffix = name || element.kind

  return `${slideId}-${suffix}-${index + 1}`
}

function getPPTElementKindLabel(element: PPTElement) {
  if (element.kind === 'shape') {
    return getPPTShapeLabel(element.shape)
  }

  if (element.kind === 'textBox') {
    return 'Text'
  }

  if (element.kind === 'image') {
    return 'Image'
  }

  if (element.kind === 'line') {
    return 'Line'
  }

  if (element.kind === 'freeform') {
    return 'Freeform'
  }

  if (element.kind === 'table') {
    return 'Table'
  }

  if (element.kind === 'comment') {
    return 'Comment'
  }

  return 'Object'
}

function getPPTShapeLabel(shape: PPTShapeKind) {
  if (shape === 'ellipse') {
    return 'Oval'
  }

  if (shape === 'diamond') {
    return 'Diamond'
  }

  return 'Rectangle'
}

function getPPTThumbElementClassName(element: PPTElement) {
  if (element.kind === 'textBox') {
    return 'ppt-thumb-text'
  }

  if (element.kind === 'image') {
    return 'ppt-thumb-image'
  }

  if (element.kind === 'line') {
    return 'ppt-thumb-line'
  }

  if (element.kind === 'freeform') {
    return 'ppt-thumb-freeform'
  }

  if (element.kind === 'table') {
    return 'ppt-thumb-table'
  }

  if (element.kind === 'comment') {
    return 'ppt-thumb-comment'
  }

  return 'ppt-thumb-shape'
}

function isPPTShapeKind(value: string): value is PPTShapeKind {
  return value === 'rect' || value === 'ellipse' || value === 'diamond'
}

function isPPTImageFit(value: string): value is PPTImageFit {
  return value === 'cover' || value === 'contain'
}

function isPPTLineMarker(value: string): value is PPTLineMarker {
  return value === 'none' || value === 'arrow'
}

function isPPTLineRoute(value: string): value is PPTLineRoute {
  return value === 'straight' || value === 'elbow'
}

function isPPTStrokeDash(value: string): value is PPTStrokeDash {
  return isSlideEditObjectStrokeLineStyleValue(value)
}

function getPPTLayerSelection(
  selection: string[],
  elementId: string,
  additive: boolean,
  slide: PPTSlide,
) {
  const fallbackSelection = getPPTSingleElementSelection(
    selection,
    elementId,
    additive,
  )

  return getPPTGroupPointerSelection({
    additive,
    fallbackSelection,
    includeHidden: true,
    itemId: elementId,
    selection,
    slide,
  })
}

function getPPTSingleElementSelection(
  selection: string[],
  elementId: string,
  additive: boolean,
) {
  return getPPTCanvasSingleItemSelection({
    additive,
    itemId: elementId,
    selection,
  })
}

function getPPTGroupPointerSelection({
  additive,
  fallbackSelection,
  includeHidden = false,
  itemId,
  selection,
  slide,
}: {
  additive: boolean
  fallbackSelection: string[]
  includeHidden?: boolean
  itemId: string
  selection: string[]
  slide: PPTSlide
}) {
  return getPPTCanvasGroupedItemPointerSelection({
    additive,
    fallbackSelection,
    getItemGroupId: (element) => element.groupId,
    getItemId: (element) => element.id,
    isItemSelectable: (element) => includeHidden || element.visible !== false,
    itemId,
    items: slide.elements,
    selection,
  })
}

export default App
