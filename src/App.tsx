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
  Grid2X2,
  Group,
  Hand,
  Highlighter,
  ImagePlus,
  Italic,
  Keyboard,
  Layers,
  List,
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
  type CSSProperties,
  type DragEvent as ReactDragEvent,
  type FocusEvent as ReactFocusEvent,
  type KeyboardEvent as ReactKeyboardEvent,
  type MouseEvent as ReactMouseEvent,
  type PointerEvent as ReactPointerEvent,
  type RefObject,
} from 'react'
import {
  createSlideEditLayoutPlaceholderDescriptor,
  createSlideEditThemeDescriptor,
  getSlideEditFrameGuideGeometry,
  getSlideEditLayoutApplyCommandEffect,
  getSlideEditLayoutPlaceholderVisibilityDescriptor,
  getSlideEditRailKeyboardCommandEffect,
  getSlideEditRailPointerCommandEffect,
  getSlideEditResolvedLayoutPlaceholder,
  toSlideEditRailHostCommandEffect,
  type SlideEditFrameGuideConfig,
  type SlideEditFrameGuideGeometry,
  type SlideEditLayoutDescriptor,
  type SlideEditMasterDescriptor,
  type SlideEditPlaceholderDescriptor,
  type SlideEditRailHostCommandEffect,
  type SlideEditResolvedLayoutPlaceholder,
  type SlideEditThemeColorToken,
} from '@interactive-os/slide-edit-affordance'
import {
  INITIAL_VIEWPORT,
  RESIZE_HANDLES,
  clamp,
  fitBoundsIntoViewport,
  getCanvasViewportWorldPoint,
  getCanvasViewportZoomStepMultiplier,
  handlePoint,
  normalizeBounds,
  type Bounds,
  type Point,
  type ResizeHandle,
  type Viewport,
  zoomViewport,
} from 'canvas/core'
import {
  EMPTY_CANVAS_SNAP_GUIDES,
  getCanvasItemPointerSelection,
  getCanvasMarqueeSelection,
  getCanvasMoveSnap,
  isAdditivePointerInput,
  moveCanvasSelection,
  resizeCanvasSelection,
  type CanvasSnapGuides,
} from 'canvas/foundation'
import {
  CANVAS_COMMAND_AFFORDANCES,
  CANVAS_TOOL_AFFORDANCES,
  alignCanvasCommand,
  createCanvasShape,
  createCanvasText,
  createCanvasAffordanceConfig,
  deleteCanvasCommand,
  distributeCanvasCommand,
  duplicateCanvasCommand,
  getCanvasWheelViewport,
  groupCanvasCommand,
  lockCanvasCommand,
  nudgeCanvasCommand,
  reorderCanvasCommand,
  selectAllCanvasCommand,
  ungroupCanvasCommand,
  unlockAllCanvasCommand,
  type CanvasAlignMode,
  type CanvasCommandItemsResult,
  type CanvasCreatedShapeKind,
  type CanvasCreationAdapter,
  type CanvasDistributeMode,
  type CanvasReorderMode,
} from 'canvas/engine'
import {
  PPT_DEFAULT_THEME_ID,
  PPT_SPLIT_LAYOUT_ID,
  PPT_SLIDE_HEIGHT,
  PPT_SLIDE_WIDTH,
  PPT_TITLE_BODY_LAYOUT_ID,
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
  pptGeometryToBounds,
  pptCanvasTransformAdapter,
} from './pptCanvasAdapter'
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
  createPPTImportedImageElement,
  getPPTImageFileFromDataTransfer,
  getPPTImageFileFromList,
  readPPTImageFileSource,
  type PPTImageImportSource,
} from './pptImageImport'
import {
  PPT_DEFAULT_TABLE_ROWS,
  createPPTTableElement,
  getPPTTableColumnCount,
  getPPTTableFileFromDataTransfer,
  getPPTTableSourceFromDataTransfer,
  normalizePPTTableRows,
  readPPTTableFileSource,
  stringifyPPTTableRows,
  type PPTTableImportSource,
} from './pptTableImport'
import './App.css'

const PPT_CANVAS_COMMAND_CONFIG = createCanvasAffordanceConfig({
  commands: {
    group: true,
    lockSelection: true,
    ungroup: true,
    unlockAll: true,
  },
})
const PPT_RECENT_COLOR_LIMIT = 8
const PPT_PARAGRAPH_ALIGN_OPTIONS = ['left', 'center', 'right'] as const

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
  CanvasAlignMode,
  keyof ReturnType<typeof getPPTCanvasCommandAvailability>
>

const canvasDistributeModeAvailabilityKey = {
  distributeHorizontal: 'distributeHorizontal',
  distributeVertical: 'distributeVertical',
} as const satisfies Record<
  CanvasDistributeMode,
  keyof ReturnType<typeof getPPTCanvasCommandAvailability>
>

const canvasReorderModeAvailabilityKey = {
  bringForward: 'bringForward',
  bringToFront: 'bringToFront',
  sendBackward: 'sendBackward',
  sendToBack: 'sendToBack',
} as const satisfies Record<
  CanvasReorderMode,
  keyof ReturnType<typeof getPPTCanvasCommandAvailability>
>

type PPTCommandSurface = 'context-menu' | 'selection-floating-bar'
type PPTClipboardOperation = 'copy' | 'cut'
type PPTClipboardObjectMetadata = {
  groupId?: string | null
  objectId: string
  placeholderId?: string | null
}
type PPTClipboardPayload = {
  elements: PPTElement[]
  metadata: readonly PPTClipboardObjectMetadata[]
  objects: PPTElement[]
  operation: PPTClipboardOperation
  selectedObjectIds: string[]
  selection: string[]
  sourceSlideId: string
  type: 'slide-object-clipboard'
}
type PPTClipboard = PPTClipboardPayload
type PPTClipboardPasteTarget =
  | {
      kind: 'active-slide'
      slideId: string
    }
  | {
      kind: 'pointer-position'
      pointerPosition: Point
      slideId: string
    }
  | {
      kind: 'slide-frame-offset'
      offset: Point
      slideId: string
    }
  | {
      kind: 'viewport-center'
      slideId: string
      viewportCenter: Point
    }
type PPTClipboardPasteObjectMapping = {
  sourceGroupId?: string | null
  sourceObjectId: string
  sourcePlaceholderId?: string | null
  targetGroupId?: string | null
  targetObjectId: string
  targetPlaceholderId?: string | null
}
type PPTClipboardPastePlan = {
  anchor: Point
  mappings: readonly PPTClipboardPasteObjectMapping[]
  operation: PPTClipboardOperation
  sourceSlideId: string
  targetSlideId: string
}
type PPTClipboardPasteCommand = {
  id: 'paste-slide-objects'
  pastePlan: PPTClipboardPastePlan
  payload: PPTClipboardPayload
}
type PPTClipboardPasteHostCommandEffect = {
  payload: PPTClipboardPasteCommand
  selection: {
    objectIds: readonly string[]
    slideId: string
  }
  type: 'slide-command-effect'
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
type PPTColorSwatchChannel =
  | 'line-stroke'
  | 'shape-fill'
  | 'shape-stroke'
  | 'text-color'
type PPTSurfaceCommand =
  | 'alignCenter'
  | 'alignLeft'
  | 'alignRight'
  | 'bringForward'
  | 'bringToFront'
  | 'copyFormatting'
  | 'delete'
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
type PPTCommandPaletteItem = {
  disabled?: boolean
  id: string
  run: () => void
  section: string
  shortcut?: string
  title: string
}
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
type PPTElementAnimationType = PPTElementAnimation['type']
type PPTElementAnimationTrigger = PPTElementAnimation['trigger']
type PPTElementAnimationUpdateField =
  | 'delayMs'
  | 'durationMs'
  | 'order'
  | 'trigger'
  | 'type'
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
type PPTTextInset = NonNullable<PPTTextStyle['textInset']>
type PPTTextInsetField = keyof PPTTextInset
type PPTTextVerticalAlign = NonNullable<PPTTextStyle['verticalAlign']>
type PPTLayerPaneAriaContract = {
  containerRole: 'tree'
  keyboardModel: 'roving-tabindex'
  rowRole: 'treeitem'
  selectionModel: 'host-controlled-multi-select'
}
type PPTLayerPaneRowDescriptor = {
  ariaLevel: number
  ariaPosInSet: number
  ariaSetSize: number
  displayName: string
  groupId: string | null
  isGrouped: boolean
  isGroup: boolean
  isHidden: boolean
  isLocked: boolean
  isRenamable: boolean
  isReorderable: boolean
  isSelectable: boolean
  isSelected: boolean
  kindLabel: string
  objectId: string
  order: number
  parentObjectId: string | null
  slideId: string
}
type PPTLayerPaneDescriptor = {
  activeObjectId: string | null
  aria: PPTLayerPaneAriaContract
  rows: readonly PPTLayerPaneRowDescriptor[]
  selectedObjectIds: readonly string[]
  slideId: string
}
type PPTLayerPaneCommandId =
  | 'hide-objects'
  | 'lock-objects'
  | 'rename-object'
  | 'reorder-object'
  | 'select-objects'
  | 'show-objects'
  | 'unlock-objects'
type PPTLayerPaneCommandDescriptor = {
  id: PPTLayerPaneCommandId
  requiredAdapterSlot: 'command-effect'
}
type PPTLayerPaneCommand =
  | {
      id: 'hide-objects' | 'lock-objects' | 'show-objects' | 'unlock-objects'
      objectIds: readonly string[]
    }
  | {
      id: 'rename-object'
      name: string
      objectId: string
    }
  | {
      fromIndex: number
      id: 'reorder-object'
      objectId: string
      toIndex: number
    }
  | {
      id: 'select-objects'
      mode: 'additive' | 'range' | 'replace'
      objectIds: readonly string[]
    }
type PPTLayerPaneHostCommandEffect = {
  payload: PPTLayerPaneCommand
  selection: {
    objectIds: readonly string[]
    slideId: string
  }
  type: 'slide-command-effect'
}
type PPTLayerPaneIntent =
  | {
      additive?: boolean
      objectId: string
      rangeAnchorObjectId?: string | null
      type: 'row-press'
    }
  | {
      name: string
      objectId: string
      type: 'rename-submit'
    }
  | {
      objectId: string
      type: 'visibility-toggle'
    }
  | {
      objectId: string
      type: 'lock-toggle'
    }
  | {
      objectId: string
      toIndex: number
      type: 'row-drop'
    }
type PPTMinimapSize = {
  h: number
  w: number
}
type PPTMinimapItemBounds = {
  bounds: Bounds
  id: string
}
type PPTMinimapItemRect = {
  id: string
  rect: Bounds
}
type PPTMinimapReadModel = {
  displayBounds: Bounds
  itemRects: PPTMinimapItemRect[]
  size: PPTMinimapSize
  viewportRect: Bounds
  viewportWorldBounds: Bounds
  worldBounds: Bounds
}
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
type PPTSelectionCommandAnchor = Point & {
  placement: 'above' | 'below'
}
type PPTTextQuickFormatState = {
  align: NonNullable<PPTParagraph['align']>
  bullet: boolean
  color: string
  fontSize: number
  isBold: boolean
  isItalic: boolean
  isUnderline: boolean
}

const PPT_COMMAND_SURFACE_GROUPS: readonly PPTSurfaceCommandGroup[] = [{
  commands: [{
    availability: 'duplicate',
    command: 'duplicate',
    dataCommand: 'duplicate',
    label: 'Duplicate',
    surfaces: ['context-menu', 'selection-floating-bar'],
    title: CANVAS_COMMAND_AFFORDANCES.duplicate.title,
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
    title: CANVAS_COMMAND_AFFORDANCES.delete.title,
  }],
  id: 'edit',
}, {
  commands: [{
    availability: 'alignLeft',
    command: 'alignLeft',
    dataCommand: 'align-left',
    label: 'Align left',
    surfaces: ['context-menu'],
    title: CANVAS_COMMAND_AFFORDANCES.alignLeft.title,
  }, {
    availability: 'alignCenter',
    command: 'alignCenter',
    dataCommand: 'align-center-x',
    label: 'Align center',
    surfaces: ['context-menu', 'selection-floating-bar'],
    title: CANVAS_COMMAND_AFFORDANCES.alignCenter.title,
  }, {
    availability: 'alignRight',
    command: 'alignRight',
    dataCommand: 'align-right',
    label: 'Align right',
    surfaces: ['context-menu'],
    title: CANVAS_COMMAND_AFFORDANCES.alignRight.title,
  }],
  id: 'align',
}, {
  commands: [{
    availability: 'bringForward',
    command: 'bringForward',
    dataCommand: 'bring-forward',
    label: 'Bring forward',
    surfaces: ['context-menu'],
    title: CANVAS_COMMAND_AFFORDANCES.bringForward.title,
  }, {
    availability: 'bringToFront',
    command: 'bringToFront',
    dataCommand: 'bring-to-front',
    label: 'Bring to front',
    surfaces: ['context-menu', 'selection-floating-bar'],
    title: CANVAS_COMMAND_AFFORDANCES.bringToFront.title,
  }, {
    availability: 'sendBackward',
    command: 'sendBackward',
    dataCommand: 'send-backward',
    label: 'Send backward',
    surfaces: ['context-menu'],
    title: CANVAS_COMMAND_AFFORDANCES.sendBackward.title,
  }, {
    availability: 'sendToBack',
    command: 'sendToBack',
    dataCommand: 'send-to-back',
    label: 'Send to back',
    surfaces: ['context-menu', 'selection-floating-bar'],
    title: CANVAS_COMMAND_AFFORDANCES.sendToBack.title,
  }],
  id: 'order',
}, {
  commands: [{
    availability: 'group',
    command: 'group',
    dataCommand: 'group',
    label: 'Group',
    surfaces: ['context-menu', 'selection-floating-bar'],
    title: CANVAS_COMMAND_AFFORDANCES.group.title,
  }, {
    availability: 'ungroup',
    command: 'ungroup',
    dataCommand: 'ungroup',
    label: 'Ungroup',
    surfaces: ['context-menu', 'selection-floating-bar'],
    title: CANVAS_COMMAND_AFFORDANCES.ungroup.title,
  }],
  id: 'group',
}, {
  commands: [{
    availability: 'lockSelection',
    command: 'lockSelection',
    dataCommand: 'lock-selection',
    label: 'Lock',
    surfaces: ['context-menu', 'selection-floating-bar'],
    title: CANVAS_COMMAND_AFFORDANCES.lockSelection.title,
  }, {
    availability: 'unlockAll',
    command: 'unlockAll',
    dataCommand: 'unlock-all',
    label: 'Unlock all',
    surfaces: ['context-menu'],
    title: CANVAS_COMMAND_AFFORDANCES.unlockAll.title,
  }],
  id: 'lock',
}]

const PPT_LINE_CONNECTION_DISTANCE = 36
const PPT_TIDY_GAP = 24
const PPT_COMMENT_DEFAULT_BODY = 'Comment'
const PPT_COMMENT_DEFAULT_AUTHOR = 'You'
const PPT_COMMENT_DEFAULT_CREATED_AT = 'Just now'
const PPT_COMMENT_BODY_MAX_LENGTH = 240
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
const PPT_TEXT_FONT_FAMILY_VALUES = new Set<string>(
  PPT_TEXT_FONT_FAMILY_OPTIONS.map((option) => option.value),
)
const PPT_DEFAULT_TEXT_VERTICAL_ALIGN: PPTTextVerticalAlign = 'top'
const PPT_TEXT_VERTICAL_ALIGN_OPTIONS = Object.freeze([
  { css: 'flex-start', label: 'Top', value: 'top' },
  { css: 'center', label: 'Middle', value: 'middle' },
  { css: 'flex-end', label: 'Bottom', value: 'bottom' },
] as const satisfies readonly {
  css: string
  label: string
  value: PPTTextVerticalAlign
}[])
const PPT_TEXT_VERTICAL_ALIGN_VALUES = new Set<string>(
  PPT_TEXT_VERTICAL_ALIGN_OPTIONS.map((option) => option.value),
)
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
const PPT_HYPERLINK_URL_MAX_LENGTH = 2048
const PPT_STROKE_DASH_OPTIONS = Object.freeze([
  { label: 'Solid', value: 'solid' },
  { label: 'Dash', value: 'dash' },
  { label: 'Dot', value: 'dot' },
] as const satisfies readonly {
  label: string
  value: PPTStrokeDash
}[])
const PPT_STROKE_DASH_VALUES = new Set<PPTStrokeDash>(
  PPT_STROKE_DASH_OPTIONS.map((option) => option.value),
)
const PPT_SLIDE_TRANSITION_TYPES = Object.freeze([
  'none',
  'fade',
  'push',
] as const satisfies readonly PPTSlideTransitionType[])
const PPT_DEFAULT_SLIDE_TRANSITION = Object.freeze({
  advanceAfterMs: null,
  advanceOnClick: true,
  durationMs: 0,
  type: 'none',
} as const satisfies PPTSlideTransition)
const PPT_SLIDE_TRANSITION_DURATION_MAX = 10000
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
const PPT_ELEMENT_ANIMATION_TIME_MAX = 10000
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
const PPT_LAYER_PANE_ARIA_CONTRACT = Object.freeze({
  containerRole: 'tree',
  keyboardModel: 'roving-tabindex',
  rowRole: 'treeitem',
  selectionModel: 'host-controlled-multi-select',
} as const satisfies PPTLayerPaneAriaContract)
const PPT_LAYER_PANE_COMMANDS = Object.freeze([
  { id: 'select-objects', requiredAdapterSlot: 'command-effect' },
  { id: 'rename-object', requiredAdapterSlot: 'command-effect' },
  { id: 'hide-objects', requiredAdapterSlot: 'command-effect' },
  { id: 'show-objects', requiredAdapterSlot: 'command-effect' },
  { id: 'lock-objects', requiredAdapterSlot: 'command-effect' },
  { id: 'unlock-objects', requiredAdapterSlot: 'command-effect' },
  { id: 'reorder-object', requiredAdapterSlot: 'command-effect' },
] as const satisfies readonly PPTLayerPaneCommandDescriptor[])
const PPT_MINIMAP_SIZE: PPTMinimapSize = {
  h: 112,
  w: 176,
}
const PPT_MINIMAP_PADDING = 8
const PPT_MINIMAP_MIN_WORLD_SIZE = 120
const PPT_LASER_POINT_DISTANCE = 3
const PPT_LASER_TRAIL_MAX_POINTS = 80
const PPT_ERASER_POINT_DISTANCE = 4
const PPT_ERASER_HIT_PADDING = 8

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
      snapGuides: CanvasSnapGuides
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
      startPoint: Point
      startViewport: Viewport
    }
  | {
      currentPoint: Point
      kind: 'laser'
      points: Point[]
      startPoint: Point
    }
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
  const [interaction, setInteraction] = useState<Interaction | null>(null)
  const [clipboard, setClipboard] = useState<PPTClipboard | null>(null)
  const [styleClipboard, setStyleClipboard] = useState<PPTStyleClipboard | null>(null)
  const [lastClipboardPasteEffect, setLastClipboardPasteEffect] = useState<PPTClipboardPasteHostCommandEffect | null>(null)
  const [lastPlaceholderVisibilityEffect, setLastPlaceholderVisibilityEffect] = useState<PPTLayoutPlaceholderVisibilityHostCommandEffect | null>(null)
  const [lastSlideRailCommandEffect, setLastSlideRailCommandEffect] = useState<SlideEditRailHostCommandEffect<string> | null>(null)
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
  const imageInputRef = useRef<HTMLInputElement | null>(null)
  const findInputRef = useRef<HTMLInputElement | null>(null)
  const commandPaletteRestoreFocusRef = useRef<HTMLElement | null>(null)
  const topbarToolbarRef = useRef<HTMLElement | null>(null)
  const [topbarToolbarFocusIndex, setTopbarToolbarFocusIndex] = useState(0)
  const slideDragSuppressClickRef = useRef(false)
  const deckRef = useRef(deck)

  useEffect(() => {
    deckRef.current = deck
  }, [deck])

  useEffect(() => {
    globalThis.window?.dispatchEvent(new Event('ppt-ready'))
  }, [])

  useLayoutEffect(() => {
    syncPPTTopbarToolbarRovingTabIndex(topbarToolbarRef.current, topbarToolbarFocusIndex)
  })

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
    () => activeSlide.elements.filter((element) => selection.includes(element.id)),
    [activeSlide.elements, selection],
  )
  const selectedTextElements = useMemo(
    () => selectedElements.filter(isPPTTextElement),
    [selectedElements],
  )
  const selectedTextOverflow = selectedElement && isPPTTextElement(selectedElement)
    ? textOverflowById[selectedElement.id] === true
    : false
  const activeLayout = useMemo(
    () => getPPTLayoutDescriptor(activeSlide.layoutId),
    [activeSlide.layoutId],
  )
  const activeSlideTransition = useMemo(
    () => getPPTSlideTransition(activeSlide),
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
    () => activeSlide.elements
      .filter((element) => element.visible !== false)
      .map((element) => ({
        bounds: pptGeometryToBounds(element.geometry),
        id: element.id,
      })),
    [activeSlide.elements],
  )
  const stageRect = stageRef.current?.getBoundingClientRect()
  const minimapModel = stageRect && showMinimap
    ? getPPTMinimapReadModel({
        items: minimapItems,
        stageSize: {
          h: stageRect.height,
          w: stageRect.width,
        },
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
  const canPasteFormatting = styleClipboard !== null &&
    selectedElements.some((element) =>
      canApplyPPTStyleClipboard(element, styleClipboard))
  const canFormatSelectedText = selectedTextElements.length > 0 &&
    selectedTextElements.length === selectedElements.length &&
    !hasLockedSelection &&
    !hasHiddenSelection
  const textQuickFormatState = canFormatSelectedText
    ? getPPTTextQuickFormatState(selectedTextElements)
    : null
  const commandAvailability = useMemo<PPTCommandAvailability>(() => ({
    ...getPPTCanvasCommandAvailability({
      canPaste: (clipboard?.objects.length ?? 0) > 0,
      canRedo: future.length > 0,
      canUndo: past.length > 0,
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
  const lastSlideRailReorderPayload = lastSlideRailCommandEffect?.payload.id === 'reorder-slide'
    ? lastSlideRailCommandEffect.payload
    : null
  const canResizeSelection = selectedBounds
    ? scene.canResizeSelection?.(selection) ?? true
    : false
  const canFitSelection = selectedBounds !== null

  const fitSlide = useCallback(() => {
    const rect = stageRef.current?.getBoundingClientRect()

    if (!rect) {
      return
    }

    setViewport(fitBoundsIntoViewport({
      h: PPT_SLIDE_HEIGHT,
      w: PPT_SLIDE_WIDTH,
      x: 0,
      y: 0,
    }, {
      height: rect.height,
      width: rect.width,
    }))
  }, [])

  const fitSelection = useCallback(() => {
    const rect = stageRef.current?.getBoundingClientRect()
    const bounds = scene.getBounds(selection)

    if (!rect || !bounds) {
      return
    }

    setViewport(fitBoundsIntoViewport(bounds, {
      height: rect.height,
      width: rect.width,
    }))
  }, [scene, selection])

  const resetZoom = useCallback(() => {
    setViewport(INITIAL_VIEWPORT)
  }, [])

  const navigateMinimapToWorldPoint = useCallback((point: Point) => {
    const rect = stageRef.current?.getBoundingClientRect()

    if (!rect) {
      return
    }

    setViewport((current) => getPPTMinimapViewportForWorldCenter({
      current,
      stageSize: {
        h: rect.height,
        w: rect.width,
      },
      worldCenter: point,
    }))
  }, [])

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
    fitSlide()
  }, [activeSlideId, fitSlide])

  useEffect(() => {
    window.addEventListener('resize', fitSlide)

    return () => window.removeEventListener('resize', fitSlide)
  }, [fitSlide])

  useEffect(() => {
    const stage = stageRef.current

    if (!stage) {
      return
    }

    stage.addEventListener('wheel', handleStageWheel, { passive: false })

    return () => stage.removeEventListener('wheel', handleStageWheel)
  })

  useEffect(() => {
    if (!contextMenu) {
      return
    }

    function onPointerDown(event: PointerEvent) {
      if (event.target instanceof Element &&
        event.target.closest('[data-ppt-context-menu]')) {
        return
      }

      setContextMenu(null)
    }

    document.addEventListener('pointerdown', onPointerDown)

    return () => document.removeEventListener('pointerdown', onPointerDown)
  }, [contextMenu])

  useEffect(() => {
    if (!findOpen) {
      return
    }

    const frame = window.requestAnimationFrame(() => {
      findInputRef.current?.focus()
      findInputRef.current?.select()
    })

    return () => window.cancelAnimationFrame(frame)
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
        if (event.key === 'Escape') {
          event.preventDefault()
          exitPresentation()
          return
        }

        if (event.key === 'ArrowRight' || event.key === 'PageDown' || event.key === ' ') {
          event.preventDefault()
          navigatePresentation(1)
          return
        }

        if (event.key === 'ArrowLeft' || event.key === 'PageUp') {
          event.preventDefault()
          navigatePresentation(-1)
          return
        }
      }

      if (isEditableTarget(event.target)) {
        return
      }

      if (shortcutHelpOpen) {
        if (event.key === 'Escape') {
          event.preventDefault()
          closeShortcutHelp()
        }
        return
      }

      if (isPPTTemporaryPanShortcut(event) && !isPPTTemporaryPanBlockedTarget(event.target)) {
        event.preventDefault()
        setContextMenu(null)
        setIsTemporaryPanActive(true)
        return
      }

      if (isPPTShortcutHelpShortcut(event)) {
        event.preventDefault()
        openShortcutHelp()
        return
      }

      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'k') {
        event.preventDefault()
        openCommandPalette()
        return
      }

      if (contextMenu && event.key === 'Escape') {
        event.preventDefault()
        setContextMenu(null)
        return
      }

      if (event.key === 'ContextMenu' || (event.shiftKey && event.key === 'F10')) {
        if (openPPTContextMenuAtSelection()) {
          event.preventDefault()
        }
        return
      }

      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'f') {
        event.preventDefault()
        openFindStrip()
        return
      }

      if ((event.metaKey || event.ctrlKey) && event.key === '0') {
        event.preventDefault()
        resetZoom()
        return
      }

      if ((event.metaKey || event.ctrlKey) && (event.key === '=' || event.key === '+')) {
        event.preventDefault()
        zoom('in')
        return
      }

      if ((event.metaKey || event.ctrlKey) && event.key === '-') {
        event.preventDefault()
        zoom('out')
        return
      }

      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'z') {
        event.preventDefault()
        if (event.shiftKey) {
          redo()
        } else {
          undo()
        }
      }

      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'y') {
        event.preventDefault()
        redo()
        return
      }

      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'a') {
        event.preventDefault()
        selectAllElements()
        return
      }

      if (
        (event.metaKey || event.ctrlKey) &&
        event.shiftKey &&
        event.key.toLowerCase() === 'c'
      ) {
        if (commandAvailability.copyFormatting) {
          event.preventDefault()
          copyFormatting()
        }
        return
      }

      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'c') {
        event.preventDefault()
        copySelection()
        return
      }

      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'x') {
        event.preventDefault()
        cutSelection()
        return
      }

      if (
        (event.metaKey || event.ctrlKey) &&
        event.shiftKey &&
        event.key.toLowerCase() === 'v'
      ) {
        if (commandAvailability.pasteFormatting) {
          event.preventDefault()
          pasteFormatting()
        }
        return
      }

      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'v') {
        if (commandAvailability.paste) {
          event.preventDefault()
          pasteSelection()
        }
        return
      }

      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'd') {
        event.preventDefault()
        duplicateSelection()
        return
      }

      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'g') {
        event.preventDefault()
        if (event.shiftKey) {
          ungroupSelection()
        } else {
          groupSelection()
        }
        return
      }

      if ((event.metaKey || event.ctrlKey) && event.key === ']') {
        event.preventDefault()
        reorderSelection(event.shiftKey ? 'bringToFront' : 'bringForward')
        return
      }

      if ((event.metaKey || event.ctrlKey) && event.key === '[') {
        event.preventDefault()
        reorderSelection(event.shiftKey ? 'sendToBack' : 'sendBackward')
        return
      }

      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'l') {
        event.preventDefault()
        if (event.shiftKey) {
          unlockAllElements()
        } else {
          lockSelectedElements()
        }
        return
      }

      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'b') {
        if (canFormatSelectedText) {
          event.preventDefault()
          toggleSelectedTextBold()
        }
        return
      }

      if (isPPTPanToolShortcut(event)) {
        event.preventDefault()
        activatePanTool()
        return
      }

      if (isPPTEraserToolShortcut(event)) {
        event.preventDefault()
        activateEraserTool()
        return
      }

      const shortcutTool = getPPTCreationToolForShortcut(event)

      if (shortcutTool) {
        event.preventDefault()
        activatePPTCreationTool(shortcutTool)
        return
      }

      if (isPPTLaserToolShortcut(event)) {
        event.preventDefault()
        activateLaserTool()
        return
      }

      if (isPPTSelectToolShortcut(event)) {
        event.preventDefault()
        activateSelectTool()
        return
      }

      if (!event.altKey && !event.ctrlKey && !event.metaKey && event.key === '0') {
        event.preventDefault()
        fitSlide()
        return
      }

      if (!event.altKey && !event.ctrlKey && !event.metaKey && event.key === '1') {
        event.preventDefault()
        fitSelection()
        return
      }

      if (event.key === 'PageUp') {
        event.preventDefault()
        activateRelativeSlide(-1)
        return
      }

      if (event.key === 'PageDown') {
        event.preventDefault()
        activateRelativeSlide(1)
        return
      }

      if (event.key === 'Escape') {
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

      if (event.key === 'Backspace' || event.key === 'Delete') {
        event.preventDefault()
        deleteSelection()
        return
      }

      if (isArrowKey(event.key)) {
        event.preventDefault()
        const distance = event.shiftKey ? 10 : 1
        const delta = getArrowNudgeDelta(event.key, distance)
        nudgeSelection(delta.dx, delta.dy)
        return
      }
    }

    function releaseTemporaryPan() {
      setIsTemporaryPanActive(false)
      setInteraction((current) => current?.kind === 'pan' ? null : current)
    }

    function onKeyUp(event: KeyboardEvent) {
      if (isPPTTemporaryPanKey(event)) {
        releaseTemporaryPan()
      }
    }

    window.addEventListener('keydown', onKeyDown)
    window.addEventListener('keyup', onKeyUp)
    window.addEventListener('blur', releaseTemporaryPan)

    return () => {
      window.removeEventListener('keydown', onKeyDown)
      window.removeEventListener('keyup', onKeyUp)
      window.removeEventListener('blur', releaseTemporaryPan)
    }
  })

  useEffect(() => {
    function onPaste(event: ClipboardEvent) {
      if (isEditableTarget(event.target)) {
        return
      }

      const file = getPPTImageFileFromDataTransfer(event.clipboardData)

      if (file) {
        event.preventDefault()
        void insertPPTImageFile(file)
        return
      }

      const tableSource = getPPTTableSourceFromDataTransfer(event.clipboardData)

      if (tableSource) {
        event.preventDefault()
        insertPPTTableSource(tableSource)
      }
    }

    window.addEventListener('paste', onPaste)

    return () => window.removeEventListener('paste', onPaste)
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
    run: (slide: PPTSlide) => CanvasCommandItemsResult<PPTElement> | null,
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
          new Set(result.selection.filter((id) =>
            result.items.some((element) => element.id === id && element.kind === 'line'))),
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
    window.requestAnimationFrame(() => {
      const thumb = [...document.querySelectorAll<HTMLButtonElement>('.ppt-thumb')]
        .find((item) => item.getAttribute('data-ppt-slide-id') === slideId)

      thumb?.focus({ preventScroll: true })
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
    let handled = false
    let targetSlideId: string | null = null

    if (event.key === 'ArrowDown' || event.key === 'ArrowRight') {
      handled = true
      const effect = getSlideEditRailKeyboardCommandEffect({
        activeSlideId: slideId,
        direction: 'next',
        slideOrder,
        type: 'select-relative',
      })

      targetSlideId = effect?.payload.id === 'select-active-slide'
        ? effect.payload.slideId
        : null
    } else if (event.key === 'ArrowUp' || event.key === 'ArrowLeft') {
      handled = true
      const effect = getSlideEditRailKeyboardCommandEffect({
        activeSlideId: slideId,
        direction: 'previous',
        slideOrder,
        type: 'select-relative',
      })

      targetSlideId = effect?.payload.id === 'select-active-slide'
        ? effect.payload.slideId
        : null
    } else if (event.key === 'Home') {
      handled = true
      targetSlideId = slideOrder[0] ?? null
    } else if (event.key === 'End') {
      handled = true
      targetSlideId = slideOrder.at(-1) ?? null
    } else if (event.key === 'Enter' || event.key === ' ') {
      handled = true
      targetSlideId = slideId
    }

    if (!handled) {
      return
    }

    event.preventDefault()
    event.stopPropagation()

    if (!targetSlideId) {
      return
    }

    setLastSlideRailCommandEffect(toSlideEditRailHostCommandEffect({
      id: 'select-active-slide',
      slideId: targetSlideId,
    }))
    selectSlide(targetSlideId)
    focusPPTSlideThumb(targetSlideId)
  }

  function activateRelativeSlide(delta: number) {
    const index = deck.slides.findIndex((slide) => slide.id === activeSlide.id)
    const nextSlide = deck.slides[index + delta]

    if (nextSlide) {
      selectSlide(nextSlide.id)
    }
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

  function handleTopbarToolbarFocus(event: ReactFocusEvent<HTMLElement>) {
    const toolbar = topbarToolbarRef.current
    const item = getPPTTopbarToolbarItemFromTarget(toolbar, event.target)

    if (!toolbar || !item) {
      return
    }

    const items = getPPTTopbarToolbarItems(toolbar)
    const index = items.indexOf(item)

    if (index >= 0) {
      setTopbarToolbarFocusIndex(index)
    }
  }

  function handleTopbarToolbarKeyDown(event: ReactKeyboardEvent<HTMLElement>) {
    if (event.altKey || event.ctrlKey || event.metaKey) {
      return
    }

    const toolbar = topbarToolbarRef.current
    const item = getPPTTopbarToolbarItemFromTarget(toolbar, event.target)

    if (!toolbar || !item) {
      return
    }

    const items = getPPTTopbarToolbarItems(toolbar)
    const currentIndex = items.indexOf(item)

    if (items.length === 0 || currentIndex < 0) {
      return
    }

    let nextIndex: number | null = null

    if (event.key === 'ArrowRight' || event.key === 'ArrowDown') {
      nextIndex = (currentIndex + 1) % items.length
    } else if (event.key === 'ArrowLeft' || event.key === 'ArrowUp') {
      nextIndex = (currentIndex - 1 + items.length) % items.length
    } else if (event.key === 'Home') {
      nextIndex = 0
    } else if (event.key === 'End') {
      nextIndex = items.length - 1
    }

    if (nextIndex === null) {
      return
    }

    event.preventDefault()
    event.stopPropagation()
    syncPPTTopbarToolbarRovingTabIndex(toolbar, nextIndex)
    setTopbarToolbarFocusIndex(nextIndex)
    items[nextIndex]?.focus({ preventScroll: true })
  }

  function openCommandPalette() {
    const activeElement = document.activeElement
    commandPaletteRestoreFocusRef.current =
      activeElement instanceof HTMLElement &&
        !activeElement.closest('[data-ppt-command-palette]')
        ? activeElement
        : null
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
    const restoreFocusElement = commandPaletteRestoreFocusRef.current

    setCommandPaletteOpen(false)
    commandPaletteRestoreFocusRef.current = null

    if (restoreFocusElement?.isConnected) {
      window.requestAnimationFrame(() => {
        restoreFocusElement.focus({ preventScroll: true })
      })
    }
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
    if (event.key === 'Enter') {
      event.preventDefault()
      goToFindMatch(event.shiftKey ? -1 : 1)
      return
    }

    if (event.key === 'Escape') {
      event.preventDefault()
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
      const index = current.slides.findIndex((slide) => slide.id === activeSlide.id)
      const source = current.slides[index]

      if (!source) {
        return current
      }

      const id = createPPTSlideId(current)
      const slide = clonePPTSlide(source, id)
      const slides = [...current.slides]
      slides.splice(index + 1, 0, slide)
      selectSlide(slide.id)

      return {
        ...current,
        slides,
      }
    })
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
      const slides = current.slides.filter((slide) => slide.id !== activeSlide.id)
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

      const slides = [...current.slides]
      const slide = slides[index]
      slides[index] = slides[targetIndex]
      slides[targetIndex] = slide

      return {
        ...current,
        slides,
      }
    })
  }

  function getSlideThumbDropPlacement(
    event: ReactDragEvent<HTMLButtonElement>,
  ): PPTSlideDropPlacement {
    const rect = event.currentTarget.getBoundingClientRect()

    return event.clientY < rect.top + rect.height / 2 ? 'before' : 'after'
  }

  function clearSlideDragState({
    suppressClick = false,
  }: {
    suppressClick?: boolean
  } = {}) {
    if (suppressClick) {
      slideDragSuppressClickRef.current = true
      window.setTimeout(() => {
        slideDragSuppressClickRef.current = false
      }, 120)
    }

    setSlideDragState(null)
  }

  function handleSlideThumbDragStart(
    slideId: string,
    event: ReactDragEvent<HTMLButtonElement>,
  ) {
    event.dataTransfer.effectAllowed = 'move'
    event.dataTransfer.setData('text/plain', slideId)
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
    event.dataTransfer.dropEffect = 'move'

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
      event.dataTransfer.getData('text/plain')
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
    const toIndex = getPPTSlideDropIndex(
      currentDeck.slides,
      sourceSlideId,
      targetSlideId,
      placement,
    )

    if (toIndex === null) {
      return
    }

    const effect = getSlideEditRailPointerCommandEffect({
      slideId: sourceSlideId,
      slideOrder: currentDeck.slides.map((slide) => slide.id),
      toIndex,
      type: 'thumbnail-drop',
    })

    if (!effect) {
      return
    }

    setLastSlideRailCommandEffect(effect)

    commitDeck((current) => {
      const dropIndex = getPPTSlideDropIndex(
        current.slides,
        sourceSlideId,
        targetSlideId,
        placement,
      )

      if (dropIndex === null) {
        return current
      }

      const sourceIndex = current.slides.findIndex((slide) => slide.id === sourceSlideId)

      if (sourceIndex < 0 || sourceIndex === dropIndex) {
        return current
      }

      const slides = [...current.slides]
      const [slide] = slides.splice(sourceIndex, 1)
      slides.splice(dropIndex, 0, slide)

      return {
        ...current,
        slides,
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

  function insertPPTImageSource(
    source: PPTImageImportSource,
    center = getPPTViewportCenter(),
  ) {
    commitDeck((current) => updatePPTDeckSlide(current, activeSlide.id, (slide) => {
      const element = createPPTImportedImageElement({
        center,
        createId: createPPTElementIdFactory(slide),
        source,
      })

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

  function insertPPTTableSource(
    source: PPTTableImportSource = { rows: PPT_DEFAULT_TABLE_ROWS },
    center = getPPTViewportCenter(),
  ) {
    commitDeck((current) => updatePPTDeckSlide(current, activeSlide.id, (slide) => {
      const id = createPPTElementId(slide, 'table')
      const element = createPPTTableElement({
        id,
        name: source.name ?? 'Table',
        point: center,
        rows: source.rows,
      })

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
      deleteCanvasCommand({
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
      duplicateCanvasCommand({
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
      const items = nudgeCanvasCommand({
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

  function alignSelection(mode: CanvasAlignMode) {
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

      return alignCanvasCommand({
        adapter: commandAdapter,
        config: PPT_CANVAS_COMMAND_CONFIG,
        items: slide.elements,
        mode,
        selection,
      })
    })
  }

  function distributeSelection(mode: CanvasDistributeMode) {
    if (!commandAvailability[canvasDistributeModeAvailabilityKey[mode]]) {
      return
    }

    commitElementCommand((slide) =>
      distributeCanvasCommand({
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
      groupCanvasCommand({
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
      ungroupCanvasCommand({
        adapter: commandAdapter,
        config: PPT_CANVAS_COMMAND_CONFIG,
        items: slide.elements,
        selection,
      }))
  }

  function reorderSelection(mode: CanvasReorderMode) {
    if (!commandAvailability[canvasReorderModeAvailabilityKey[mode]]) {
      return
    }

    commitElementCommand((slide) =>
      reorderCanvasCommand({
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
      lockCanvasCommand({
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
      unlockAllCanvasCommand({
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

    setStyleClipboard(next)
  }

  function pasteFormatting() {
    if (!commandAvailability.pasteFormatting || !styleClipboard) {
      return
    }

    const selectedIds = new Set(selection)

    commitDeck((current) =>
      updatePPTDeckSlide(current, activeSlide.id, (slide) => ({
        ...slide,
        elements: slide.elements.map((element) =>
          selectedIds.has(element.id)
            ? applyPPTStyleClipboardToElement(element, styleClipboard)
            : element),
      })),
    )
  }

  function copySelection(operation: PPTClipboardOperation = 'copy') {
    const selected = activeSlide.elements.filter((element) => selection.includes(element.id))

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
    void navigator.clipboard?.writeText(JSON.stringify({
      metadata: payload.metadata,
      objects: selected,
      operation: payload.operation,
      selectedObjectIds: payload.selectedObjectIds,
      sourceSlideId: payload.sourceSlideId,
      type: payload.type,
    })).catch(() => undefined)
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

    commitDeck((current) => updatePPTDeckSlide(current, activeSlide.id, (slide) => {
      const effect = createPPTClipboardPasteCommandEffect({
        createId: createPPTElementIdFactory(slide),
        payload: clipboard,
        slideFrame: {
          h: PPT_SLIDE_HEIGHT,
          w: PPT_SLIDE_WIDTH,
          x: 0,
          y: 0,
        },
        target: {
          kind: 'slide-frame-offset',
          offset: { x: 28, y: 28 },
          slideId: slide.id,
        },
      })

      if (!effect) {
        return slide
      }

      setLastClipboardPasteEffect(effect)
      setSelection([...effect.selection.objectIds])

      return applyPPTClipboardPasteHostCommandEffect(slide, effect)
    }))
  }

  function selectAllElements() {
    const nextSelection = selectAllCanvasCommand({
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
      case 'alignCenter':
        alignSelection('alignCenter')
        break
      case 'alignLeft':
        alignSelection('alignLeft')
        break
      case 'alignRight':
        alignSelection('alignRight')
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

  function autoFitTextElement(elementId: string) {
    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) => {
        if (
          !isPPTTextElement(element) ||
          element.locked === true ||
          element.visible === false
        ) {
          return element
        }

        const size = measurePPTTextAutoFitSize(element)

        if (!size) {
          return element
        }

        return {
          ...element,
          geometry: updatePPTElementBounds(element, {
            ...pptGeometryToBounds(element.geometry),
            ...size,
          }).geometry,
          textAutoFit: PPT_TEXT_AUTOFIT,
        }
      }),
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
          rotation: normalizePPTElementRotation(rotation),
        },
      })),
    )
  }

  function updateElementOpacity(elementId: string, opacity: number) {
    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) => ({
        ...element,
        opacity: normalizePPTElementOpacity(opacity),
      })),
    )
  }

  function updateElementHyperlink(elementId: string, url: string) {
    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) => {
        const hyperlink = normalizePPTElementHyperlink({ url })

        return {
          ...element,
          hyperlink: hyperlink ?? undefined,
        }
      }),
    )
  }

  function updateElementAltText(elementId: string, altText: string) {
    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) => {
        const accessibility = normalizePPTElementAccessibility({ altText })

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
    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) => {
        if (field === 'enabled') {
          return {
            ...element,
            shadow: value === true ? getPPTElementShadow(element) : undefined,
          }
        }

        return {
          ...element,
          shadow: normalizePPTElementShadow({
            ...getPPTElementShadow(element),
            [field]: value,
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
    commitDeck((current) =>
      updatePPTDeckSlide(current, activeSlide.id, (slide) => ({
        ...slide,
        elements: slide.elements.map((element) => {
          if (element.id !== elementId) {
            return element
          }

          return {
            ...element,
            animation: normalizePPTElementAnimation({
              ...getPPTElementAnimation(element, slide),
              [field]: value,
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
          field === 'fontFamily'
            ? normalizePPTTextFontFamily(String(value))
            : field === 'verticalAlign'
              ? normalizePPTTextVerticalAlign(String(value))
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
    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) => {
        if (!isPPTTextElement(element) || element.locked === true) {
          return element
        }

        const inset = getPPTTextElementInset(element)

        return {
          ...element,
          style: {
            ...getPPTTextElementStyle(element),
            textInset: {
              ...inset,
              [field]: normalizePPTTextInset(value),
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

    const selectedIds = new Set(selection)

    commitDeck((current) =>
      updatePPTDeckSlide(current, activeSlide.id, (slide) => ({
        ...slide,
        elements: slide.elements.map((element) => {
          if (!selectedIds.has(element.id) ||
            !isPPTTextElement(element) ||
            element.locked === true ||
            element.visible === false) {
            return element
          }

          return {
            ...element,
            style: update(getPPTTextElementStyle(element)),
          }
        }),
      })),
    )
  }

  function updateSelectedParagraphAlign(
    align: NonNullable<PPTParagraph['align']>,
  ) {
    if (!canFormatSelectedText) {
      return
    }

    const selectedIds = new Set(selection)

    commitDeck((current) =>
      updatePPTDeckSlide(current, activeSlide.id, (slide) => ({
        ...slide,
        elements: slide.elements.map((element) => {
          if (!selectedIds.has(element.id) ||
            !isPPTTextElement(element) ||
            element.locked === true ||
            element.visible === false) {
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
      })),
    )
  }

  function updateSelectedParagraphBullet(enabled: boolean) {
    if (!canFormatSelectedText) {
      return
    }

    const selectedIds = new Set(selection)

    commitDeck((current) =>
      updatePPTDeckSlide(current, activeSlide.id, (slide) => ({
        ...slide,
        elements: slide.elements.map((element) => {
          if (!selectedIds.has(element.id) ||
            !isPPTTextElement(element) ||
            element.locked === true ||
            element.visible === false) {
            return element
          }

          return {
            ...element,
            textBody: {
              paragraphs: element.textBody.paragraphs.map((paragraph) => ({
                ...paragraph,
                ...(enabled ? { bullet: 'bullet' as const } : { bullet: undefined }),
              })),
            },
          }
        }),
      })),
    )
  }

  function toggleSelectedParagraphBullet() {
    const enabled = !areAllPPTTextElementsBulleted(selectedTextElements)

    updateSelectedParagraphBullet(enabled)
  }

  function updateSelectedTextRunStyle(
    field: 'italic' | 'underline',
    enabled: boolean,
  ) {
    if (!canFormatSelectedText) {
      return
    }

    const selectedIds = new Set(selection)

    commitDeck((current) =>
      updatePPTDeckSlide(current, activeSlide.id, (slide) => ({
        ...slide,
        elements: slide.elements.map((element) => {
          if (!selectedIds.has(element.id) ||
            !isPPTTextElement(element) ||
            element.locked === true ||
            element.visible === false) {
            return element
          }

          return {
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
          }
        }),
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
      fontSize: clamp(
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

    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) =>
        element.kind === 'shape'
          ? {
              ...element,
              fill: normalizePPTFill({
                ...element.fill,
                [field]: value,
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
    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) => {
        if (element.kind !== 'shape' || element.shape !== 'rect') {
          return element
        }

        const normalized = normalizePPTShapeCornerRadius(cornerRadius)

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
    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) => {
        if (element.kind !== 'image') {
          return element
        }

        return {
          ...element,
          fit,
        }
      }),
    )
  }

  function updateImageCrop(
    elementId: string,
    field: keyof PPTImageCrop,
    value: number,
  ) {
    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) => {
        if (element.kind !== 'image') {
          return element
        }

        return {
          ...element,
          crop: {
            ...getPPTImageCrop(element),
            [field]: clamp(value, 0, 100),
          },
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

  function applyLayerPaneCommandEffect(effect: PPTLayerPaneHostCommandEffect) {
    const payload = effect.payload

    switch (payload.id) {
      case 'hide-objects':
      case 'show-objects':
      case 'lock-objects':
      case 'unlock-objects': {
        const objectIds = new Set(payload.objectIds)
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

        setSelection(effect.selection.objectIds.filter((objectId) => objectIds.has(objectId)))
        commitDeck((current) =>
          updatePPTDeckSlide(current, effect.selection.slideId, (slide) => ({
            ...slide,
            elements: slide.elements.map((element) =>
              objectIds.has(element.id)
                ? {
                    ...element,
                    ...(visible === null ? {} : { visible }),
                    ...(locked === null ? {} : { locked }),
                  }
                : element),
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
            const elements = [...slide.elements]
            const [element] = elements.splice(payload.fromIndex, 1)

            if (!element) {
              return slide
            }

            elements.splice(payload.toIndex, 0, element)

            return {
              ...slide,
              elements,
            }
          }),
        )
        setSelection([...effect.selection.objectIds])
        return
      case 'select-objects': {
        const targetObjectId = payload.objectIds.at(-1)

        if (!targetObjectId) {
          return
        }

        if (payload.mode === 'range') {
          setSelection([...payload.objectIds])
          return
        }

        setSelection((current) =>
          getPPTLayerSelection(
            current,
            targetObjectId,
            payload.mode === 'additive',
            activeSlide,
          ))
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
          [field]: value,
        })

        return {
          ...element,
          stroke,
        }
      }),
    )
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

  function updateParagraphBullet(
    elementId: string,
    enabled: boolean,
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
              ...(enabled ? { bullet: 'bullet' as const } : { bullet: undefined }),
            })),
          },
        }
      }),
    )
  }

  function updateParagraphSpacing(
    elementId: string,
    field: PPTParagraphSpacingField,
    value: number,
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
              [field]: field === 'lineHeight'
                ? normalizePPTParagraphLineHeight(value)
                : normalizePPTParagraphSpacing(value),
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
    commitDeck((current) =>
      updatePPTDeckSlide(current, activeSlide.id, (slide) => {
        const transition = getPPTSlideTransition(slide)

        return {
          ...slide,
          transition: normalizePPTSlideTransition({
            ...transition,
            [field]: value,
          }),
        }
      }))
  }

  function copyHTML() {
    void navigator.clipboard.writeText(exportCode).catch(() => undefined)
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
    const url = URL.createObjectURL(new Blob([content], {
      type,
    }))
    const anchor = document.createElement('a')
    anchor.href = url
    anchor.download = filename
    anchor.click()
    URL.revokeObjectURL(url)
  }

  function toggleTheme() {
    setTheme((current) => (current === 'dark' ? 'light' : 'dark'))
  }

  function screenToWorld(event: Pick<PointerEvent, 'clientX' | 'clientY'>) {
    const rect = stageRef.current?.getBoundingClientRect()

    return getCanvasViewportWorldPoint(viewport, {
      x: event.clientX - (rect?.left ?? 0),
      y: event.clientY - (rect?.top ?? 0),
    })
  }

  function getPointerClientPoint(event: Pick<PointerEvent, 'clientX' | 'clientY'>): Point {
    return {
      x: event.clientX,
      y: event.clientY,
    }
  }

  function getPPTViewportCenter() {
    const rect = stageRef.current?.getBoundingClientRect()

    if (!rect) {
      return {
        x: PPT_SLIDE_WIDTH / 2,
        y: PPT_SLIDE_HEIGHT / 2,
      }
    }

    return getCanvasViewportWorldPoint(viewport, {
      x: rect.width / 2,
      y: rect.height / 2,
    })
  }

  function worldToScreen(point: Point) {
    const rect = stageRef.current?.getBoundingClientRect()

    return {
      x: (rect?.left ?? 0) + viewport.x + point.x * viewport.scale,
      y: (rect?.top ?? 0) + viewport.y + point.y * viewport.scale,
    }
  }

  function openPPTContextMenu(x: number, y: number) {
    const menuWidth = 220
    const menuHeight = 320
    const margin = 8
    const viewportWidth = globalThis.innerWidth || x + menuWidth + margin
    const viewportHeight = globalThis.innerHeight || y + menuHeight + margin

    setContextMenu({
      x: clamp(x, margin, Math.max(margin, viewportWidth - menuWidth - margin)),
      y: clamp(y, margin, Math.max(margin, viewportHeight - menuHeight - margin)),
    })
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
    if (
      getPPTImageFileFromDataTransfer(event.dataTransfer) ||
      getPPTTableFileFromDataTransfer(event.dataTransfer) ||
      getPPTTableSourceFromDataTransfer(event.dataTransfer)
    ) {
      event.preventDefault()
    }
  }

  function handleStageDrop(event: ReactDragEvent<HTMLDivElement>) {
    const imageFile = getPPTImageFileFromDataTransfer(event.dataTransfer)

    if (imageFile) {
      event.preventDefault()
      void insertPPTImageFile(imageFile, screenToWorld(event.nativeEvent))
      return
    }

    const tableFile = getPPTTableFileFromDataTransfer(event.dataTransfer)
    const tableSource = getPPTTableSourceFromDataTransfer(event.dataTransfer)
    const point = screenToWorld(event.nativeEvent)

    if (tableFile) {
      event.preventDefault()
      void insertPPTTableFile(tableFile, point).then((inserted) => {
        if (!inserted && tableSource) {
          insertPPTTableSource(tableSource, point)
        }
      })
      return
    }

    if (tableSource) {
      event.preventDefault()
      insertPPTTableSource(tableSource, point)
    }
  }

  function handleStageWheel(event: WheelEvent) {
    if (
      editingId ||
      isEditableTarget(event.target) ||
      isPPTWheelViewportPassthroughTarget(event.target)
    ) {
      return
    }

    const rect = stageRef.current?.getBoundingClientRect()

    if (!rect) {
      return
    }

    const nextViewport = getCanvasWheelViewport({
      config: PPT_CANVAS_COMMAND_CONFIG,
      input: {
        ctrlKey: event.ctrlKey,
        deltaMode: event.deltaMode,
        deltaX: event.deltaX,
        deltaY: event.deltaY,
        metaKey: event.metaKey,
        shiftKey: event.shiftKey,
      },
      point: {
        x: event.clientX - rect.left,
        y: event.clientY - rect.top,
      },
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
    event.currentTarget.setPointerCapture(event.pointerId)
    setContextMenu(null)
    setInteraction({
      kind: 'pan',
      startPoint: getPointerClientPoint(event.nativeEvent),
      startViewport: viewport,
    })

    return true
  }

  function beginLaserPointer(
    event: ReactPointerEvent<HTMLElement>,
    point: Point,
  ) {
    if (!isLaserToolActive || event.button !== 0) {
      return false
    }

    event.preventDefault()
    event.stopPropagation()
    event.currentTarget.setPointerCapture(event.pointerId)

    const points = [clampPPTPointToSlide(point)]

    setContextMenu(null)
    setLaserTrailPoints(points)
    setInteraction({
      currentPoint: points[0],
      kind: 'laser',
      points,
      startPoint: points[0],
    })

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
    event.currentTarget.setPointerCapture(event.pointerId)

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
    event.currentTarget.setPointerCapture(event.pointerId)

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
    event.currentTarget.setPointerCapture(event.pointerId)

    const startDeck = deckRef.current
    const startSlide = findPPTSlide(startDeck, activeSlide.id)
    const points = [clampPPTPointToSlide(point)]
    const erasedIds = getPPTEraserHitElementIds(startSlide, points)

    setContextMenu(null)
    setSelection((current) => current.filter((id) => !erasedIds.includes(id)))
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
    event.currentTarget.setPointerCapture(event.pointerId)

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

    if (beginTemporaryPan(event)) {
      return
    }

    if (beginEraser(event, screenToWorld(event.nativeEvent))) {
      return
    }

    if (beginLaserPointer(event, screenToWorld(event.nativeEvent))) {
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
    event.currentTarget.setPointerCapture(event.pointerId)

    const additive = isAdditivePointerInput(event)
    const pointerSelection = getCanvasItemPointerSelection({
      additive,
      itemId: elementId,
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
      snapGuides: EMPTY_CANVAS_SNAP_GUIDES,
      startDeck: interactionStartDeck,
      startPoint: screenToWorld(event.nativeEvent),
    })
  }

  function handleElementContextMenu(
    event: ReactMouseEvent<HTMLDivElement>,
    elementId: string,
  ) {
    if (editingId || isEditableTarget(event.target)) {
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

    event.currentTarget.setPointerCapture(event.pointerId)
    const additive = isAdditivePointerInput(event)
    const point = screenToWorld(event.nativeEvent)

    if (beginTemporaryPan(event)) {
      return
    }

    if (beginEraser(event, point)) {
      return
    }

    if (beginLaserPointer(event, point)) {
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
    if (editingId || isEditableTarget(event.target)) {
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

    if (event.detail > 1) {
      return
    }

    event.currentTarget.setPointerCapture(event.pointerId)

    if (!selectedBounds || !canResizeSelection) {
      return
    }

    setInteraction({
      bounds: selectedBounds,
      handle,
      kind: 'resize',
      selection,
      slideId: activeSlide.id,
      startDeck: deckRef.current,
    })
  }

  function handleResizeHandleDoubleClick(
    event: ReactMouseEvent<HTMLButtonElement>,
    handle: ResizeHandle,
  ) {
    event.preventDefault()
    event.stopPropagation()
    autoSizeSelection(handle)
  }

  function handleRotatePointerDown(event: ReactPointerEvent<HTMLButtonElement>) {
    event.preventDefault()
    event.stopPropagation()
    event.currentTarget.setPointerCapture(event.pointerId)

    if (!selectedBounds || !canResizeSelection) {
      return
    }

    const center = getBoundsCenter(selectedBounds)
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
    event.currentTarget.setPointerCapture(event.pointerId)

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
    event.currentTarget.setPointerCapture(event.pointerId)

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

    commitDeck((current) => updatePPTDeckSlide(current, activeSlide.id, (slide) => ({
      ...slide,
      elements: slide.elements.map((element) => {
        if (!selection.includes(element.id)) {
          return element
        }

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
      const point = getPointerClientPoint(event.nativeEvent)

      setViewport({
        ...interaction.startViewport,
        x: interaction.startViewport.x + point.x - interaction.startPoint.x,
        y: interaction.startViewport.y + point.y - interaction.startPoint.y,
      })
      return
    }

    const point = screenToWorld(event.nativeEvent)

    if (interaction.kind === 'laser') {
      const points = getNextPPTLaserTrailPoints(interaction.points, point)

      setLaserTrailPoints(points)
      setInteraction({
        ...interaction,
        currentPoint: clampPPTPointToSlide(point),
        points,
      })
      return
    }

    if (interaction.kind === 'erase') {
      const points = getNextPPTEraserPoints(interaction.points, point)
      const startSlide = findPPTSlide(interaction.startDeck, interaction.slideId)
      const erasedIds = mergePPTElementIds(
        interaction.erasedIds,
        getPPTEraserHitElementIds(startSlide, points),
      )

      setSelection((current) => current.filter((id) => !erasedIds.includes(id)))
      setInteraction({
        ...interaction,
        currentPoint: clampPPTPointToSlide(point),
        erasedIds,
        points,
      })
      return
    }

    if (interaction.kind === 'marquee') {
      const bounds = normalizeBounds(interaction.startPoint, point)
      const nextSelection = getCanvasMarqueeSelection({
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
      const elements = currentSlide.elements.map((element) =>
        element.id === interaction.lineId && element.kind === 'line'
          ? updatePPTLineEndpoint(
              element,
              'end',
              point,
              currentSlide,
              interaction.lineId,
            )
          : element)
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
      const elements = currentSlide.elements.map((element) =>
        element.id === interaction.elementId && element.kind === 'freeform'
          ? createPPTFreeformElement({
              id: element.id,
              name: element.name,
              opacity: element.opacity ?? 1,
              points,
              stroke: element.stroke,
            })
          : element)
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
      const elements = currentSlide.elements.map((element) =>
        element.id === interaction.elementId
          ? createPPTElementFromCreationTool({
              current: point,
              id: interaction.elementId,
              start: interaction.startPoint,
              tool: interaction.tool,
            })
          : element)
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
      const elements = currentSlide.elements.map((element) =>
        element.id === interaction.lineId && element.kind === 'line'
          ? updatePPTLineRouteBend(element, point)
          : element)
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
      const snap = getCanvasMoveSnap({
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
      const elements = moveCanvasSelection({
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
      const elements = resizeCanvasSelection({
        adapter: pptCanvasTransformAdapter,
        bounds: interaction.bounds,
        handle: interaction.handle,
        items: startSlide.elements,
        point,
        preserveAspectRatio: event.shiftKey,
        resizeFromCenter: event.altKey,
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
      const elements = startSlide.elements.map((element) =>
        element.id === interaction.lineId && element.kind === 'line'
          ? updatePPTLineEndpoint(
              element,
              interaction.endpoint,
              point,
              startSlide,
              interaction.lineId,
            )
          : element)

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
    const elements = startSlide.elements.map((element) => {
      const startRotation = rotationById.get(element.id)

      if (startRotation === undefined || element.locked === true) {
        return element
      }

      const rawRotation = normalizePPTElementRotation(startRotation + delta)
      const rotation = event.shiftKey
        ? Math.round(rawRotation / 15) * 15
        : rawRotation

      return {
        ...element,
        geometry: {
          ...element.geometry,
          rotation: normalizePPTElementRotation(rotation),
        },
      }
    })

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
        elements: slide.elements.filter((element) => !erasedIds.has(element.id)),
      }))

      deckRef.current = nextDeck
      setDeck(nextDeck)
      setSelection((current) => current.filter((id) => !erasedIds.has(id)))
    }

    if (interaction.kind === 'line-create') {
      const currentSlide = findPPTSlide(deckRef.current, interaction.slideId)
      const createdLine = currentSlide.elements.find((element) =>
        element.id === interaction.lineId && element.kind === 'line')

      if (createdLine?.kind === 'line' && getPPTLineLength(createdLine) < 8) {
        const fallbackEnd = {
          x: Math.min(PPT_SLIDE_WIDTH, interaction.startPoint.x + 160),
          y: interaction.startPoint.y,
        }
        const elements = currentSlide.elements.map((element) =>
          element.id === createdLine.id && element.kind === 'line'
            ? updatePPTLineEndpoint(
                element,
                'end',
                fallbackEnd,
                currentSlide,
                interaction.lineId,
              )
            : element)
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
          {
            x: Math.min(PPT_SLIDE_WIDTH, interaction.points[0].x + 96),
            y: Math.min(PPT_SLIDE_HEIGHT, interaction.points[0].y + 36),
          },
        ]
        const elements = currentSlide.elements.map((element) =>
          element.id === interaction.elementId && element.kind === 'freeform'
            ? createPPTFreeformElement({
                id: element.id,
                name: element.name,
                opacity: element.opacity ?? 1,
                points: fallbackPoints,
                stroke: element.stroke,
              })
            : element)
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

      if (interaction.tool.kind === 'text') {
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
    const rect = stageRef.current?.getBoundingClientRect()
    const point = {
      x: (rect?.width ?? 0) / 2,
      y: (rect?.height ?? 0) / 2,
    }
    const multiplier = getCanvasViewportZoomStepMultiplier(viewport.scale, direction)

    setViewport((current) => zoomViewport(current, point, multiplier))
  }

  const marqueeBounds = interaction?.kind === 'marquee'
    ? normalizeBounds(interaction.startPoint, interaction.currentPoint)
    : null
  const snapGuides = interaction?.kind === 'move'
    ? interaction.snapGuides
    : EMPTY_CANVAS_SNAP_GUIDES
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
    ? getPPTSelectionCommandAnchor({
        barWidth: selectionCommandBarWidth,
        bounds: selectedBounds,
        stage: stageRef.current,
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
    run: openShortcutHelp,
    section: 'System',
    shortcut: PPT_SHORTCUT_HELP_SHORTCUT,
    title: 'Keyboard shortcuts',
  }, {
    disabled: !commandAvailability.undo,
    id: 'command:undo',
    run: undo,
    section: 'Edit',
    shortcut: 'Cmd/Ctrl+Z',
    title: CANVAS_COMMAND_AFFORDANCES.undo.title,
  }, {
    disabled: !commandAvailability.redo,
    id: 'command:redo',
    run: redo,
    section: 'Edit',
    shortcut: 'Cmd/Ctrl+Y',
    title: CANVAS_COMMAND_AFFORDANCES.redo.title,
  }, {
    disabled: !commandAvailability.duplicate,
    id: 'command:duplicate',
    run: duplicateSelection,
    section: 'Edit',
    shortcut: 'Cmd/Ctrl+D',
    title: CANVAS_COMMAND_AFFORDANCES.duplicate.title,
  }, {
    disabled: !commandAvailability.delete,
    id: 'command:delete',
    run: deleteSelection,
    section: 'Edit',
    title: CANVAS_COMMAND_AFFORDANCES.delete.title,
  }, {
    disabled: !commandAvailability.cut,
    id: 'command:cut',
    run: cutSelection,
    section: 'Edit',
    shortcut: 'Cmd/Ctrl+X',
    title: CANVAS_COMMAND_AFFORDANCES.cut.title,
  }, {
    disabled: selection.length === 0,
    id: 'command:copy',
    run: copySelection,
    section: 'Edit',
    shortcut: 'Cmd/Ctrl+C',
    title: CANVAS_COMMAND_AFFORDANCES.copy.title,
  }, {
    disabled: !commandAvailability.paste,
    id: 'command:paste',
    run: pasteSelection,
    section: 'Edit',
    shortcut: 'Cmd/Ctrl+V',
    title: CANVAS_COMMAND_AFFORDANCES.paste.title,
  }, {
    disabled: !commandAvailability.copyFormatting,
    id: 'command:copy-formatting',
    run: copyFormatting,
    section: 'Edit',
    shortcut: 'Shift+Cmd/Ctrl+C',
    title: 'Copy formatting',
  }, {
    disabled: !commandAvailability.pasteFormatting,
    id: 'command:paste-formatting',
    run: pasteFormatting,
    section: 'Edit',
    shortcut: 'Shift+Cmd/Ctrl+V',
    title: 'Paste formatting',
  }, {
    id: 'command:select-all',
    run: selectAllElements,
    section: 'Edit',
    shortcut: 'Cmd/Ctrl+A',
    title: CANVAS_COMMAND_AFFORDANCES.selectAll.title,
  }, {
    disabled: !commandAvailability.selectSameType,
    id: 'command:select-same-type',
    run: selectSameTypeElements,
    section: 'Edit',
    title: 'Select same type',
  }, {
    disabled: !commandAvailability.alignLeft,
    id: 'command:align-left',
    run: () => alignSelection('alignLeft'),
    section: 'Arrange',
    title: CANVAS_COMMAND_AFFORDANCES.alignLeft.title,
  }, {
    disabled: !commandAvailability.alignCenter,
    id: 'command:align-center',
    run: () => alignSelection('alignCenter'),
    section: 'Arrange',
    title: CANVAS_COMMAND_AFFORDANCES.alignCenter.title,
  }, {
    disabled: !commandAvailability.alignRight,
    id: 'command:align-right',
    run: () => alignSelection('alignRight'),
    section: 'Arrange',
    title: CANVAS_COMMAND_AFFORDANCES.alignRight.title,
  }, {
    disabled: !commandAvailability.alignTop,
    id: 'command:align-top',
    run: () => alignSelection('alignTop'),
    section: 'Arrange',
    title: CANVAS_COMMAND_AFFORDANCES.alignTop.title,
  }, {
    disabled: !commandAvailability.alignMiddle,
    id: 'command:align-middle',
    run: () => alignSelection('alignMiddle'),
    section: 'Arrange',
    title: CANVAS_COMMAND_AFFORDANCES.alignMiddle.title,
  }, {
    disabled: !commandAvailability.alignBottom,
    id: 'command:align-bottom',
    run: () => alignSelection('alignBottom'),
    section: 'Arrange',
    title: CANVAS_COMMAND_AFFORDANCES.alignBottom.title,
  }, {
    disabled: !commandAvailability.distributeHorizontal,
    id: 'command:distribute-horizontal',
    run: () => distributeSelection('distributeHorizontal'),
    section: 'Arrange',
    title: CANVAS_COMMAND_AFFORDANCES.distributeHorizontal.title,
  }, {
    disabled: !commandAvailability.distributeVertical,
    id: 'command:distribute-vertical',
    run: () => distributeSelection('distributeVertical'),
    section: 'Arrange',
    title: CANVAS_COMMAND_AFFORDANCES.distributeVertical.title,
  }, {
    disabled: !commandAvailability.tidySelection,
    id: 'command:tidy-selection',
    run: tidySelection,
    section: 'Arrange',
    title: 'Tidy selection',
  }, {
    disabled: !commandAvailability.flipSelection,
    id: 'command:flip-horizontal',
    run: () => flipSelection('horizontal'),
    section: 'Arrange',
    title: 'Flip horizontal',
  }, {
    disabled: !commandAvailability.flipSelection,
    id: 'command:flip-vertical',
    run: () => flipSelection('vertical'),
    section: 'Arrange',
    title: 'Flip vertical',
  }, {
    disabled: !commandAvailability.bringForward,
    id: 'command:bring-forward',
    run: () => reorderSelection('bringForward'),
    section: 'Arrange',
    shortcut: 'Cmd/Ctrl+]',
    title: CANVAS_COMMAND_AFFORDANCES.bringForward.title,
  }, {
    disabled: !commandAvailability.bringToFront,
    id: 'command:bring-to-front',
    run: () => reorderSelection('bringToFront'),
    section: 'Arrange',
    shortcut: 'Shift+Cmd/Ctrl+]',
    title: CANVAS_COMMAND_AFFORDANCES.bringToFront.title,
  }, {
    disabled: !commandAvailability.sendBackward,
    id: 'command:send-backward',
    run: () => reorderSelection('sendBackward'),
    section: 'Arrange',
    shortcut: 'Cmd/Ctrl+[',
    title: CANVAS_COMMAND_AFFORDANCES.sendBackward.title,
  }, {
    disabled: !commandAvailability.sendToBack,
    id: 'command:send-to-back',
    run: () => reorderSelection('sendToBack'),
    section: 'Arrange',
    shortcut: 'Shift+Cmd/Ctrl+[',
    title: CANVAS_COMMAND_AFFORDANCES.sendToBack.title,
  }, {
    disabled: !commandAvailability.group,
    id: 'command:group',
    run: groupSelection,
    section: 'Arrange',
    shortcut: 'Cmd/Ctrl+G',
    title: CANVAS_COMMAND_AFFORDANCES.group.title,
  }, {
    disabled: !commandAvailability.ungroup,
    id: 'command:ungroup',
    run: ungroupSelection,
    section: 'Arrange',
    shortcut: 'Shift+Cmd/Ctrl+G',
    title: CANVAS_COMMAND_AFFORDANCES.ungroup.title,
  }, {
    disabled: !commandAvailability.lockSelection,
    id: 'command:lock-selection',
    run: lockSelectedElements,
    section: 'Arrange',
    shortcut: 'Cmd/Ctrl+L',
    title: CANVAS_COMMAND_AFFORDANCES.lockSelection.title,
  }, {
    disabled: !commandAvailability.unlockAll,
    id: 'command:unlock-all',
    run: unlockAllElements,
    section: 'Arrange',
    shortcut: 'Shift+Cmd/Ctrl+L',
    title: CANVAS_COMMAND_AFFORDANCES.unlockAll.title,
  }, {
    id: 'tool:select',
    run: activateSelectTool,
    section: 'Create',
    shortcut: CANVAS_TOOL_AFFORDANCES.select.shortcut,
    title: 'Select tool',
  }, {
    id: 'tool:pan',
    run: activatePanTool,
    section: 'Create',
    shortcut: CANVAS_TOOL_AFFORDANCES.pan.shortcut,
    title: CANVAS_TOOL_AFFORDANCES.pan.ariaLabel,
  }, {
    id: 'tool:laser',
    run: activateLaserTool,
    section: 'Create',
    shortcut: CANVAS_TOOL_AFFORDANCES.laser.shortcut,
    title: CANVAS_TOOL_AFFORDANCES.laser.ariaLabel,
  }, {
    id: 'tool:text',
    run: () => activatePPTCreationTool({ kind: 'text' }),
    section: 'Create',
    shortcut: CANVAS_TOOL_AFFORDANCES.text.shortcut,
    title: CANVAS_TOOL_AFFORDANCES.text.ariaLabel,
  }, {
    id: 'tool:rect',
    run: () => activatePPTCreationTool({ kind: 'shape', shape: 'rect' }),
    section: 'Create',
    shortcut: CANVAS_TOOL_AFFORDANCES.rect.shortcut,
    title: CANVAS_TOOL_AFFORDANCES.rect.ariaLabel,
  }, {
    id: 'tool:ellipse',
    run: () => activatePPTCreationTool({ kind: 'shape', shape: 'ellipse' }),
    section: 'Create',
    shortcut: CANVAS_TOOL_AFFORDANCES.ellipse.shortcut,
    title: CANVAS_TOOL_AFFORDANCES.ellipse.ariaLabel,
  }, {
    id: 'tool:diamond',
    run: () => activatePPTCreationTool({ kind: 'shape', shape: 'diamond' }),
    section: 'Create',
    shortcut: CANVAS_TOOL_AFFORDANCES.diamond.shortcut,
    title: CANVAS_TOOL_AFFORDANCES.diamond.ariaLabel,
  }, {
    id: 'tool:line',
    run: () => activateLineCreationMode('line'),
    section: 'Create',
    title: 'Line tool',
  }, {
    id: 'tool:arrow',
    run: () => activateLineCreationMode('arrow'),
    section: 'Create',
    shortcut: CANVAS_TOOL_AFFORDANCES.arrow.shortcut,
    title: CANVAS_TOOL_AFFORDANCES.arrow.ariaLabel,
  }, {
    id: 'tool:comment',
    run: () => activatePPTCreationTool({ kind: 'comment' }),
    section: 'Create',
    shortcut: CANVAS_TOOL_AFFORDANCES.comment.shortcut,
    title: CANVAS_TOOL_AFFORDANCES.comment.ariaLabel,
  }, {
    id: 'tool:pen',
    run: () => activatePPTCreationTool({ kind: 'freeform', tool: 'pen' }),
    section: 'Create',
    shortcut: CANVAS_TOOL_AFFORDANCES.pen.shortcut,
    title: CANVAS_TOOL_AFFORDANCES.pen.ariaLabel,
  }, {
    id: 'tool:marker',
    run: () => activatePPTCreationTool({ kind: 'freeform', tool: 'marker' }),
    section: 'Create',
    shortcut: CANVAS_TOOL_AFFORDANCES.marker.shortcut,
    title: CANVAS_TOOL_AFFORDANCES.marker.ariaLabel,
  }, {
    id: 'tool:highlight',
    run: () => activatePPTCreationTool({ kind: 'freeform', tool: 'highlight' }),
    section: 'Create',
    shortcut: CANVAS_TOOL_AFFORDANCES.highlight.shortcut,
    title: CANVAS_TOOL_AFFORDANCES.highlight.ariaLabel,
  }, {
    id: 'tool:eraser',
    run: activateEraserTool,
    section: 'Create',
    shortcut: CANVAS_TOOL_AFFORDANCES.eraser.shortcut,
    title: CANVAS_TOOL_AFFORDANCES.eraser.ariaLabel,
  }, {
    id: 'tool:image',
    run: () => imageInputRef.current?.click(),
    section: 'Create',
    title: 'Add image',
  }, {
    id: 'tool:table',
    run: () => insertPPTTableSource(),
    section: 'Create',
    title: 'Add table',
  }, {
    id: 'slide:add',
    run: addSlide,
    section: 'Slides',
    title: 'Add slide',
  }, {
    id: 'slide:duplicate',
    run: duplicateActiveSlide,
    section: 'Slides',
    title: 'Duplicate slide',
  }, {
    disabled: !canDeleteSlide,
    id: 'slide:delete',
    run: deleteActiveSlide,
    section: 'Slides',
    title: 'Delete slide',
  }, {
    disabled: !canMoveActiveSlideUp,
    id: 'slide:move-up',
    run: () => moveActiveSlide(-1),
    section: 'Slides',
    title: 'Move slide up',
  }, {
    disabled: !canMoveActiveSlideDown,
    id: 'slide:move-down',
    run: () => moveActiveSlide(1),
    section: 'Slides',
    title: 'Move slide down',
  }, ...PPT_LAYOUT_DESCRIPTORS.map((layout) => ({
    disabled: activeLayout.layoutId === layout.layoutId,
    id: `slide:layout:${layout.layoutId}`,
    run: () => applySlideLayout(layout.layoutId),
    section: 'Slides',
    title: `Apply ${layout.name}`,
  })), {
    id: 'export:download-slide-svg',
    run: downloadSlideSVG,
    section: 'Export',
    title: 'Download slide SVG',
  }, {
    disabled: !canExportSelectionSVG,
    id: 'export:download-selection-svg',
    run: downloadSelectionSVG,
    section: 'Export',
    title: 'Download selection SVG',
  }, {
    id: 'view:find',
    run: openFindStrip,
    section: 'View',
    shortcut: 'Cmd/Ctrl+F',
    title: 'Find text',
  }, {
    id: 'view:present',
    run: () => startPresentation(),
    section: 'View',
    title: 'Start presentation',
  }, {
    id: 'view:fit-slide',
    run: fitSlide,
    section: 'View',
    shortcut: '0',
    title: 'Fit slide',
  }, {
    disabled: !canFitSelection,
    id: 'view:fit-selection',
    run: fitSelection,
    section: 'View',
    shortcut: '1',
    title: 'Fit selection',
  }, {
    id: 'view:reset-zoom',
    run: resetZoom,
    section: 'View',
    shortcut: 'Cmd/Ctrl+0',
    title: CANVAS_COMMAND_AFFORDANCES.zoomReset.title,
  }, {
    id: 'view:zoom-in',
    run: () => zoom('in'),
    section: 'View',
    shortcut: 'Cmd/Ctrl+=',
    title: CANVAS_COMMAND_AFFORDANCES.zoomIn.title,
  }, {
    id: 'view:zoom-out',
    run: () => zoom('out'),
    section: 'View',
    shortcut: 'Cmd/Ctrl+-',
    title: CANVAS_COMMAND_AFFORDANCES.zoomOut.title,
  }, {
    id: 'view:toggle-grid',
    run: () => setShowGrid((current) => !current),
    section: 'View',
    title: showGrid ? 'Hide grid' : 'Show grid',
  }, {
    id: 'view:toggle-minimap',
    run: () => setShowMinimap((current) => !current),
    section: 'View',
    title: showMinimap ? 'Hide minimap' : 'Show minimap',
  }, {
    id: 'view:toggle-frame-guides',
    run: () => setShowFrameGuides((current) => !current),
    section: 'View',
    title: showFrameGuides ? 'Hide frame guides' : 'Show frame guides',
  }, {
    id: 'view:toggle-theme',
    run: toggleTheme,
    section: 'View',
    title: theme === 'dark' ? 'Switch to light theme' : 'Switch to dark theme',
  }, {
    disabled: !canFormatSelectedText,
    id: 'format:bold',
    run: toggleSelectedTextBold,
    section: 'Format',
    shortcut: 'Cmd/Ctrl+B',
    title: 'Bold text',
  }, {
    disabled: !canFormatSelectedText,
    id: 'format:italic',
    run: toggleSelectedTextItalic,
    section: 'Format',
    title: 'Italic text',
  }, {
    disabled: !canFormatSelectedText,
    id: 'format:underline',
    run: toggleSelectedTextUnderline,
    section: 'Format',
    title: 'Underline text',
  }, {
    disabled: !canFormatSelectedText,
    id: 'format:bullet',
    run: toggleSelectedParagraphBullet,
    section: 'Format',
    title: 'Toggle bullet list',
  }, {
    disabled: !selectedElement || !isPPTTextElement(selectedElement) || !selectedTextOverflow,
    id: 'format:auto-fit-text',
    run: () => {
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
        data-ppt-toolbar-focus-model="roving-tabindex"
        data-ppt-toolbar-keyboard-model="arrow-home-end"
        ref={topbarToolbarRef}
        role="toolbar"
        onFocus={handleTopbarToolbarFocus}
        onKeyDown={handleTopbarToolbarKeyDown}
      >
        <div className="ppt-brand">
          <strong>PPT</strong>
          <span>{deck.title}</span>
        </div>
        <div className="ppt-toolbar-group">
          <button className="ppt-icon-button" disabled={!commandAvailability.undo} onClick={undo} title={CANVAS_COMMAND_AFFORDANCES.undo.title} type="button">
            <Undo2 size={17} />
          </button>
          <button className="ppt-icon-button" disabled={!commandAvailability.redo} onClick={redo} title={CANVAS_COMMAND_AFFORDANCES.redo.title} type="button">
            <Redo2 size={17} />
          </button>
          <button className="ppt-icon-button" data-ppt-find-open onClick={openFindStrip} title="Find text" type="button">
            <Search size={17} />
          </button>
          <button className="ppt-icon-button" data-ppt-command-palette-open onClick={openCommandPalette} title="Command palette" type="button">
            <Command size={17} />
          </button>
          <button className="ppt-icon-button" data-ppt-command="copy-formatting" disabled={!commandAvailability.copyFormatting} onClick={copyFormatting} title="Copy formatting" type="button">
            <Paintbrush size={17} />
          </button>
          <button className="ppt-icon-button" data-ppt-command="paste-formatting" disabled={!commandAvailability.pasteFormatting} onClick={pasteFormatting} title="Paste formatting" type="button">
            <Paintbrush size={17} />
          </button>
          <button className="ppt-icon-button" data-ppt-shortcut-help-open onClick={openShortcutHelp} title="Keyboard shortcuts" type="button">
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
            aria-label={CANVAS_TOOL_AFFORDANCES.pan.ariaLabel}
            aria-pressed={isPanToolActive}
            className="ppt-icon-button"
            data-ppt-pan-tool
            data-ppt-tool="pan"
            onClick={activatePanTool}
            title={CANVAS_TOOL_AFFORDANCES.pan.title}
            type="button"
          >
            <Hand size={17} />
          </button>
          <button
            aria-label={CANVAS_TOOL_AFFORDANCES.laser.ariaLabel}
            aria-pressed={isLaserToolActive}
            className="ppt-icon-button"
            data-ppt-laser-tool
            data-ppt-tool="laser"
            onClick={activateLaserTool}
            title={CANVAS_TOOL_AFFORDANCES.laser.title}
            type="button"
          >
            <MousePointer2 size={17} />
          </button>
          <button
            aria-label={CANVAS_TOOL_AFFORDANCES.text.ariaLabel}
            aria-pressed={creationTool?.kind === 'text'}
            className="ppt-icon-button"
            data-ppt-insert-tool="text"
            onClick={() => activatePPTCreationTool({ kind: 'text' })}
            title={CANVAS_TOOL_AFFORDANCES.text.title}
            type="button"
          >
            <Type size={17} />
          </button>
          <button
            aria-label={CANVAS_TOOL_AFFORDANCES.rect.ariaLabel}
            aria-pressed={isPPTShapeCreationTool(creationTool, 'rect')}
            className="ppt-icon-button"
            data-ppt-insert-shape="rect"
            data-ppt-insert-tool="rect"
            onClick={() => activatePPTCreationTool({ kind: 'shape', shape: 'rect' })}
            title={CANVAS_TOOL_AFFORDANCES.rect.title}
            type="button"
          >
            <Square size={17} />
          </button>
          <button
            aria-label={CANVAS_TOOL_AFFORDANCES.ellipse.ariaLabel}
            aria-pressed={isPPTShapeCreationTool(creationTool, 'ellipse')}
            className="ppt-icon-button"
            data-ppt-insert-shape="ellipse"
            data-ppt-insert-tool="ellipse"
            onClick={() => activatePPTCreationTool({ kind: 'shape', shape: 'ellipse' })}
            title={CANVAS_TOOL_AFFORDANCES.ellipse.title}
            type="button"
          >
            <Circle size={17} />
          </button>
          <button
            aria-label={CANVAS_TOOL_AFFORDANCES.diamond.ariaLabel}
            aria-pressed={isPPTShapeCreationTool(creationTool, 'diamond')}
            className="ppt-icon-button"
            data-ppt-insert-shape="diamond"
            data-ppt-insert-tool="diamond"
            onClick={() => activatePPTCreationTool({ kind: 'shape', shape: 'diamond' })}
            title={CANVAS_TOOL_AFFORDANCES.diamond.title}
            type="button"
          >
            <Diamond size={17} />
          </button>
          <button aria-pressed={lineCreationMode === 'line'} className="ppt-icon-button" data-ppt-insert-line="line" onClick={() => activateLineCreationMode('line')} title="Draw line" type="button">
            <Minus size={17} />
          </button>
          <button aria-pressed={lineCreationMode === 'arrow'} className="ppt-icon-button" data-ppt-insert-line="arrow" onClick={() => activateLineCreationMode('arrow')} title="Draw arrow" type="button">
            <ArrowRight size={17} />
          </button>
          <button
            aria-label={CANVAS_TOOL_AFFORDANCES.pen.ariaLabel}
            aria-pressed={isPPTFreeformCreationTool(creationTool, 'pen')}
            className="ppt-icon-button"
            data-ppt-insert-tool="pen"
            onClick={() => activatePPTCreationTool({ kind: 'freeform', tool: 'pen' })}
            title={CANVAS_TOOL_AFFORDANCES.pen.title}
            type="button"
          >
            <PenLine size={17} />
          </button>
          <button
            aria-label={CANVAS_TOOL_AFFORDANCES.marker.ariaLabel}
            aria-pressed={isPPTFreeformCreationTool(creationTool, 'marker')}
            className="ppt-icon-button"
            data-ppt-insert-tool="marker"
            onClick={() => activatePPTCreationTool({ kind: 'freeform', tool: 'marker' })}
            title={CANVAS_TOOL_AFFORDANCES.marker.title}
            type="button"
          >
            <PencilLine size={17} />
          </button>
          <button
            aria-label={CANVAS_TOOL_AFFORDANCES.highlight.ariaLabel}
            aria-pressed={isPPTFreeformCreationTool(creationTool, 'highlight')}
            className="ppt-icon-button"
            data-ppt-insert-tool="highlight"
            onClick={() => activatePPTCreationTool({ kind: 'freeform', tool: 'highlight' })}
            title={CANVAS_TOOL_AFFORDANCES.highlight.title}
            type="button"
          >
            <Highlighter size={17} />
          </button>
          <button
            aria-label={CANVAS_TOOL_AFFORDANCES.eraser.ariaLabel}
            aria-pressed={isEraserToolActive}
            className="ppt-icon-button"
            data-ppt-eraser-tool
            data-ppt-tool="eraser"
            onClick={activateEraserTool}
            title={CANVAS_TOOL_AFFORDANCES.eraser.title}
            type="button"
          >
            <Eraser size={17} />
          </button>
          <button
            aria-label={CANVAS_TOOL_AFFORDANCES.comment.ariaLabel}
            aria-pressed={creationTool?.kind === 'comment'}
            className="ppt-icon-button"
            data-ppt-insert-comment
            data-ppt-insert-tool="comment"
            onClick={() => activatePPTCreationTool({ kind: 'comment' })}
            title={CANVAS_TOOL_AFFORDANCES.comment.title}
            type="button"
          >
            <MessageSquare size={17} />
          </button>
          <button className="ppt-icon-button" data-ppt-insert-image onClick={() => imageInputRef.current?.click()} title="Add image" type="button">
            <ImagePlus size={17} />
          </button>
          <button className="ppt-icon-button" data-ppt-insert-table onClick={() => insertPPTTableSource()} title="Add table" type="button">
            <Table2 size={17} />
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
          <button className="ppt-icon-button" disabled={!commandAvailability.delete} onClick={deleteSelection} title={CANVAS_COMMAND_AFFORDANCES.delete.title} type="button">
            <Trash2 size={17} />
          </button>
        </div>
        <div className="ppt-toolbar-group">
          <button className="ppt-icon-button" data-ppt-command="align-left" disabled={!commandAvailability.alignLeft} onClick={() => alignSelection('alignLeft')} title={CANVAS_COMMAND_AFFORDANCES.alignLeft.title} type="button">
            <AlignStartVertical size={17} />
          </button>
          <button className="ppt-icon-button" data-ppt-command="align-center-x" disabled={!commandAvailability.alignCenter} onClick={() => alignSelection('alignCenter')} title={CANVAS_COMMAND_AFFORDANCES.alignCenter.title} type="button">
            <AlignCenterVertical size={17} />
          </button>
          <button className="ppt-icon-button" data-ppt-command="align-right" disabled={!commandAvailability.alignRight} onClick={() => alignSelection('alignRight')} title={CANVAS_COMMAND_AFFORDANCES.alignRight.title} type="button">
            <AlignEndVertical size={17} />
          </button>
          <button className="ppt-icon-button" data-ppt-command="align-top" disabled={!commandAvailability.alignTop} onClick={() => alignSelection('alignTop')} title={CANVAS_COMMAND_AFFORDANCES.alignTop.title} type="button">
            <AlignStartHorizontal size={17} />
          </button>
          <button className="ppt-icon-button" data-ppt-command="align-middle" disabled={!commandAvailability.alignMiddle} onClick={() => alignSelection('alignMiddle')} title={CANVAS_COMMAND_AFFORDANCES.alignMiddle.title} type="button">
            <AlignCenterHorizontal size={17} />
          </button>
          <button className="ppt-icon-button" data-ppt-command="align-bottom" disabled={!commandAvailability.alignBottom} onClick={() => alignSelection('alignBottom')} title={CANVAS_COMMAND_AFFORDANCES.alignBottom.title} type="button">
            <AlignEndHorizontal size={17} />
          </button>
          <button className="ppt-icon-button" data-ppt-command="distribute-horizontal" disabled={!commandAvailability.distributeHorizontal} onClick={() => distributeSelection('distributeHorizontal')} title={CANVAS_COMMAND_AFFORDANCES.distributeHorizontal.title} type="button">
            <AlignHorizontalDistributeCenter size={17} />
          </button>
          <button className="ppt-icon-button" data-ppt-command="distribute-vertical" disabled={!commandAvailability.distributeVertical} onClick={() => distributeSelection('distributeVertical')} title={CANVAS_COMMAND_AFFORDANCES.distributeVertical.title} type="button">
            <AlignVerticalDistributeCenter size={17} />
          </button>
          <button className="ppt-icon-button" data-ppt-command="tidy-selection" disabled={!commandAvailability.tidySelection} onClick={tidySelection} title="Tidy selection" type="button">
            <Grid2X2 size={17} />
          </button>
          <button className="ppt-icon-button" data-ppt-command="flip-horizontal" disabled={!commandAvailability.flipSelection} onClick={() => flipSelection('horizontal')} title="Flip horizontal" type="button">
            <FlipHorizontal2 size={17} />
          </button>
          <button className="ppt-icon-button" data-ppt-command="flip-vertical" disabled={!commandAvailability.flipSelection} onClick={() => flipSelection('vertical')} title="Flip vertical" type="button">
            <FlipVertical2 size={17} />
          </button>
        </div>
        <div className="ppt-toolbar-group">
          <button className="ppt-icon-button" data-ppt-command="bring-forward" disabled={!commandAvailability.bringForward} onClick={() => reorderSelection('bringForward')} title={CANVAS_COMMAND_AFFORDANCES.bringForward.title} type="button">
            <MoveUp size={17} />
          </button>
          <button className="ppt-icon-button" data-ppt-command="bring-to-front" disabled={!commandAvailability.bringToFront} onClick={() => reorderSelection('bringToFront')} title={CANVAS_COMMAND_AFFORDANCES.bringToFront.title} type="button">
            <BringToFront size={17} />
          </button>
          <button className="ppt-icon-button" data-ppt-command="send-backward" disabled={!commandAvailability.sendBackward} onClick={() => reorderSelection('sendBackward')} title={CANVAS_COMMAND_AFFORDANCES.sendBackward.title} type="button">
            <MoveDown size={17} />
          </button>
          <button className="ppt-icon-button" data-ppt-command="send-to-back" disabled={!commandAvailability.sendToBack} onClick={() => reorderSelection('sendToBack')} title={CANVAS_COMMAND_AFFORDANCES.sendToBack.title} type="button">
            <SendToBack size={17} />
          </button>
        </div>
        <div className="ppt-toolbar-group">
          <button className="ppt-icon-button" data-ppt-command="group" disabled={!commandAvailability.group} onClick={groupSelection} title={CANVAS_COMMAND_AFFORDANCES.group.title} type="button">
            <Group size={17} />
          </button>
          <button className="ppt-icon-button" data-ppt-command="ungroup" disabled={!commandAvailability.ungroup} onClick={ungroupSelection} title={CANVAS_COMMAND_AFFORDANCES.ungroup.title} type="button">
            <Ungroup size={17} />
          </button>
          <button className="ppt-icon-button" data-ppt-command="lock-selection" disabled={!commandAvailability.lockSelection} onClick={lockSelectedElements} title={CANVAS_COMMAND_AFFORDANCES.lockSelection.title} type="button">
            <Lock size={17} />
          </button>
          <button className="ppt-icon-button" data-ppt-command="unlock-all" disabled={!commandAvailability.unlockAll} onClick={unlockAllElements} title={CANVAS_COMMAND_AFFORDANCES.unlockAll.title} type="button">
            <Unlock size={17} />
          </button>
        </div>
        <div className="ppt-toolbar-group">
          <button className="ppt-icon-button" data-ppt-present-start onClick={() => startPresentation()} title="Start presentation" type="button">
            <Play size={17} />
          </button>
          <button className="ppt-icon-button" onClick={() => zoom('out')} title="Zoom out" type="button">
            <ZoomOut size={17} />
          </button>
          <button className="ppt-icon-button" data-ppt-view-fit-slide onClick={fitSlide} title="Fit slide" type="button">
            <Maximize2 size={17} />
          </button>
          <button className="ppt-icon-button" data-ppt-view-fit-selection disabled={!canFitSelection} onClick={fitSelection} title="Fit selection" type="button">
            <Maximize2 size={17} />
          </button>
          <button className="ppt-icon-button" onClick={() => zoom('in')} title="Zoom in" type="button">
            <ZoomIn size={17} />
          </button>
          <button aria-pressed={showGrid} className="ppt-icon-button" data-ppt-view-grid onClick={() => setShowGrid((current) => !current)} title="Toggle grid" type="button">
            <Grid2X2 size={17} />
          </button>
          <button aria-pressed={showMinimap} className="ppt-icon-button" data-ppt-view-minimap onClick={() => setShowMinimap((current) => !current)} title="Toggle minimap" type="button">
            <MapIcon size={17} />
          </button>
          <button aria-pressed={showFrameGuides} className="ppt-icon-button" data-ppt-view-frame-guides onClick={() => setShowFrameGuides((current) => !current)} title="Toggle frame guides" type="button">
            <Ruler size={17} />
          </button>
          <button
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
          <button aria-label="Copy HTML" className="ppt-button" onClick={copyHTML} type="button">
            <Copy size={16} /> HTML
          </button>
          <button aria-label="Download HTML" className="ppt-button" onClick={downloadHTML} type="button">
            <Download size={16} /> HTML
          </button>
          <button aria-label="Download slide SVG" className="ppt-button" data-ppt-export-svg onClick={downloadSlideSVG} type="button">
            <Download size={16} /> SVG
          </button>
          <button aria-label="Download selection SVG" className="ppt-button" data-ppt-export-selection-svg disabled={!canExportSelectionSVG} onClick={downloadSelectionSVG} type="button">
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
          data-ppt-slide-rail-command={lastSlideRailCommandEffect?.payload.id}
          data-ppt-slide-rail-command-from-index={lastSlideRailReorderPayload?.fromIndex}
          data-ppt-slide-rail-command-selection-slide={lastSlideRailCommandEffect?.selection?.slideId}
          data-ppt-slide-rail-command-slide={lastSlideRailReorderPayload?.slideId}
          data-ppt-slide-rail-command-to-index={lastSlideRailReorderPayload?.toIndex}
          data-ppt-slide-rail-command-type={lastSlideRailCommandEffect?.type}
          data-ppt-slide-rail-keyboard-keys="ArrowUp ArrowDown Home End Enter Space"
          data-ppt-slide-rail-keyboard-model="listbox-roving-focus"
          role="listbox"
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
              slide={slide}
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
        data-ppt-clipboard-operation={clipboard?.operation ?? undefined}
        data-ppt-clipboard-paste-anchor={lastClipboardPasteEffect
          ? `${lastClipboardPasteEffect.payload.pastePlan.anchor.x},${lastClipboardPasteEffect.payload.pastePlan.anchor.y}`
          : undefined}
        data-ppt-clipboard-paste-command={lastClipboardPasteEffect?.payload.id}
        data-ppt-clipboard-paste-mapping-count={lastClipboardPasteEffect?.payload.pastePlan.mappings.length}
        data-ppt-clipboard-paste-operation={lastClipboardPasteEffect?.payload.pastePlan.operation}
        data-ppt-clipboard-paste-selection={lastClipboardPasteEffect?.selection.objectIds.join(' ') ?? undefined}
        data-ppt-clipboard-paste-source-slide={lastClipboardPasteEffect?.payload.pastePlan.sourceSlideId}
        data-ppt-clipboard-paste-target-slide={lastClipboardPasteEffect?.payload.pastePlan.targetSlideId}
        data-ppt-clipboard-paste-type={lastClipboardPasteEffect?.type}
        data-ppt-clipboard-selected-object-ids={clipboard?.selectedObjectIds.join(' ') ?? undefined}
        data-ppt-clipboard-selection={clipboard?.selection.join(' ') ?? undefined}
        data-ppt-clipboard-source-slide={clipboard?.sourceSlideId ?? undefined}
        data-ppt-clipboard-type={clipboard?.type ?? undefined}
        data-ppt-style-clipboard-categories={styleClipboard?.categories.join(' ') ?? undefined}
        data-ppt-style-clipboard-source-id={styleClipboard?.sourceId ?? undefined}
        data-ppt-style-clipboard-source-kind={styleClipboard?.sourceKind ?? undefined}
        data-ppt-style-clipboard-type={styleClipboard?.type ?? undefined}
        data-ppt-placeholder-visibility-command={lastPlaceholderVisibilityEffect?.payload.id}
        data-ppt-placeholder-visibility-command-placeholder={lastPlaceholderVisibilityEffect?.payload.placeholderId}
        data-ppt-placeholder-visibility-command-selection={lastPlaceholderVisibilityEffect?.selection.objectIds.join(' ') ?? undefined}
        data-ppt-placeholder-visibility-command-slide={lastPlaceholderVisibilityEffect?.payload.slideId}
        data-ppt-placeholder-visibility-command-type={lastPlaceholderVisibilityEffect?.type}
        data-ppt-placeholder-visibility-command-visible={lastPlaceholderVisibilityEffect
          ? String(lastPlaceholderVisibilityEffect.payload.isVisible)
          : undefined}
        data-ppt-keyboard-nudge-enabled={commandAvailability.nudge ? 'true' : 'false'}
        data-ppt-keyboard-nudge-keys="ArrowLeft ArrowRight ArrowUp ArrowDown Shift+ArrowLeft Shift+ArrowRight Shift+ArrowUp Shift+ArrowDown"
        data-ppt-keyboard-nudge-large-step="10"
        data-ppt-keyboard-nudge-model="canvas-keyboard-nudge-shortcuts"
        data-ppt-keyboard-nudge-step="1"
        data-ppt-drawing-tool={creationTool?.kind === 'freeform'
          ? creationTool.tool
          : undefined}
        data-ppt-eraser-hit-count={interaction?.kind === 'erase'
          ? interaction.erasedIds.length
          : undefined}
        data-ppt-eraser-tool-active={isEraserToolActive ? 'true' : 'false'}
        data-ppt-eraser-tool-model="canvas-eraser-tool"
        data-ppt-eraser-tool-shortcut="E"
        data-ppt-highlighter-tool-model="canvas-highlighter-tool"
        data-ppt-highlighter-tool-shortcut="Shift+M"
        data-ppt-laser-tool-active={isLaserToolActive ? 'true' : 'false'}
        data-ppt-laser-tool-model="canvas-laser-pointer-tool"
        data-ppt-laser-tool-shortcut="P"
        data-ppt-laser-trail-model="canvas-laser-trail-overlay"
        data-ppt-laser-trail-point-count={laserTrailPoints.length}
        data-ppt-laser-trail-state={interaction?.kind === 'laser'
          ? 'active'
          : laserTrailPoints.length > 0 ? 'idle' : 'empty'}
        data-ppt-pan-tool-active={isPanToolActive ? 'true' : 'false'}
        data-ppt-pan-tool-model="canvas-pan-tool"
        data-ppt-pan-tool-shortcut="H"
        data-ppt-marker-tool-model="canvas-marker-tool"
        data-ppt-marker-tool-shortcut="M"
        data-ppt-resize-aspect-ratio-modifier="Shift"
        data-ppt-resize-from-center-modifier="Alt"
        data-ppt-resize-modifier-model="canvas-resize-pointer-modifiers"
        data-ppt-temporary-pan-active={isTemporaryPanActive ? 'true' : 'false'}
        data-ppt-temporary-pan-gesture={interaction?.kind === 'pan' ? 'true' : 'false'}
        data-ppt-temporary-pan-model="canvas-temporary-pan-shortcut"
        data-ppt-temporary-pan-shortcut="Space"
        data-ppt-wheel-viewport-horizontal-pan-modifier="Shift"
        data-ppt-wheel-viewport-model="canvas-wheel-viewport"
        data-ppt-wheel-viewport-pan="ordinary-wheel"
        data-ppt-wheel-viewport-zoom-modifier="Ctrl/Meta"
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
        ref={stageRef}
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
                textOverflow={selectedTextOverflow}
                onRotatePointerDown={handleRotatePointerDown}
                onResizeHandleDoubleClick={handleResizeHandleDoubleClick}
                onResizePointerDown={handleResizePointerDown}
              />
            ) : null}
            <PPTSelectionFloatingBar
              anchor={selectionCommandAnchor}
              groups={selectionFloatingCommandGroups}
              scale={viewport.scale}
              textFormat={textQuickFormatState}
              onCommand={runPPTSurfaceCommand}
              onFontSizeStep={stepSelectedTextFontSize}
              onParagraphBulletToggle={toggleSelectedParagraphBullet}
              onParagraphAlign={updateSelectedParagraphAlign}
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
        onCommentResolvedChange={updateCommentResolved}
        onCommitText={commitText}
        onCopyHTML={copyHTML}
        onDownloadHTML={downloadHTML}
        onElementAltTextChange={updateElementAltText}
        onElementAnimationChange={updateElementAnimation}
        onElementGeometryChange={updateElementGeometry}
        onElementHyperlinkChange={updateElementHyperlink}
        onImageCropChange={updateImageCrop}
        onImageFitChange={updateImageFit}
        onElementNameChange={updateElementName}
        onElementOpacityChange={updateElementOpacity}
        onElementRotationChange={updateElementRotation}
        onElementShadowChange={updateElementShadow}
        onElementTextInsetChange={updateElementTextInset}
        onLineMarkerChange={updateLineMarker}
        onLineRouteChange={updateLineRoute}
        onParagraphBulletChange={updateParagraphBullet}
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

  function navigate(event: ReactPointerEvent<SVGSVGElement>) {
    const point = getPPTMinimapSvgPoint(event, svgRef.current, readModel)

    if (!point) {
      return
    }

    onNavigateToWorldPoint(getPPTMinimapWorldPoint({ model: readModel, point }))
  }

  function handlePointerDown(event: ReactPointerEvent<SVGSVGElement>) {
    event.preventDefault()
    event.stopPropagation()
    event.currentTarget.setPointerCapture(event.pointerId)
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
      data-ppt-minimap-item-count={readModel.itemRects.length}
      data-ppt-minimap-scale={getPPTMinimapScale(readModel.worldBounds, readModel.displayBounds)}
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
        {readModel.itemRects.map((item) => (
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
  const previousFocusRef = useRef<HTMLElement | null>(null)
  const groups = useMemo(() => groupPPTShortcutHelpItems(items), [items])

  useEffect(() => {
    previousFocusRef.current = document.activeElement instanceof HTMLElement
      ? document.activeElement
      : null

    const focusTimer = window.setTimeout(() => {
      closeButtonRef.current?.focus()
    }, 0)

    return () => {
      window.clearTimeout(focusTimer)
      previousFocusRef.current?.focus()
    }
  }, [])

  function handleBackdropMouseDown(event: ReactMouseEvent<HTMLDivElement>) {
    if (event.target === event.currentTarget) {
      onClose()
    }
  }

  function handleKeyDown(event: ReactKeyboardEvent<HTMLElement>) {
    if (event.key === 'Escape') {
      event.preventDefault()
      event.stopPropagation()
      onClose()
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
  const transitionDuration = Math.max(1, transition.durationMs)

  return (
    <div
      aria-label="Presentation preview"
      aria-modal="true"
      className="ppt-presentation"
      data-ppt-presentation
      data-ppt-presentation-advance-after={transition.advanceAfterMs ?? ''}
      data-ppt-presentation-advance-on-click={transition.advanceOnClick ? 'true' : 'false'}
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
              '--ppt-transition-duration': `${transitionDuration}ms`,
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
                textOverflow={false}
                onCommitText={() => undefined}
                onContextMenu={(event) => event.preventDefault()}
                onEdit={() => undefined}
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
    window.addEventListener('resize', updateScale)

    return () => window.removeEventListener('resize', updateScale)
  }, [])

  return scale
}

function getPPTPresentationScale() {
  if (typeof window === 'undefined') {
    return 0.75
  }

  return clamp(
    Math.min(
      (window.innerWidth - 96) / PPT_SLIDE_WIDTH,
      (window.innerHeight - 168) / PPT_SLIDE_HEIGHT,
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

  useEffect(() => {
    const focusTimer = window.setTimeout(() => inputRef.current?.focus(), 0)

    return () => window.clearTimeout(focusTimer)
  }, [])

  function runItem(item: PPTCommandPaletteItem | undefined) {
    if (!item || item.disabled) {
      return
    }

    item.run()
    onClose()
  }

  function handleKeyDown(event: ReactKeyboardEvent<HTMLElement>) {
    if (event.key === 'Escape') {
      event.preventDefault()
      event.stopPropagation()
      onClose()
      return
    }

    if (event.key === 'Tab') {
      const focusables = getPPTCommandPaletteFocusables(dialogRef.current)

      if (focusables.length > 0) {
        const currentIndex = focusables.findIndex((item) => item === document.activeElement)
        const direction = event.shiftKey ? -1 : 1
        const nextIndex = currentIndex < 0
          ? 0
          : (currentIndex + direction + focusables.length) % focusables.length

        event.preventDefault()
        event.stopPropagation()
        focusables[nextIndex]?.focus()
      }
      return
    }

    if (event.key === 'ArrowDown') {
      event.preventDefault()
      event.stopPropagation()
      setActiveIndex(() => Math.min(activeItemIndex + 1, maxActiveIndex))
      return
    }

    if (event.key === 'ArrowUp') {
      event.preventDefault()
      event.stopPropagation()
      setActiveIndex(() => Math.max(0, activeItemIndex - 1))
      return
    }

    if (event.key === 'Enter') {
      event.preventDefault()
      event.stopPropagation()
      runItem(filteredItems[activeItemIndex])
    }
  }

  function handleBackdropMouseDown(event: ReactMouseEvent<HTMLDivElement>) {
    if (event.target === event.currentTarget) {
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
        data-ppt-command-palette-focus-trap="true"
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

function filterPPTCommandPaletteItems(
  items: readonly PPTCommandPaletteItem[],
  query: string,
) {
  const terms = query
    .trim()
    .toLowerCase()
    .split(/\s+/)
    .filter(Boolean)

  if (terms.length === 0) {
    return [...items]
  }

  return items.filter((item) => {
    const haystack = [
      item.title,
      item.section,
      item.shortcut ?? '',
    ].join(' ').toLowerCase()

    return terms.every((term) => haystack.includes(term))
  })
}

function getPPTCommandPaletteFocusables(dialog: HTMLElement | null) {
  if (!dialog) {
    return []
  }

  return [...dialog.querySelectorAll<HTMLElement>(
    'input, button:not(:disabled)',
  )]
}

function getPPTCommandPaletteOptionId(itemId: string) {
  return `ppt-command-palette-option-${itemId.replace(/[^A-Za-z0-9_-]/g, '-')}`
}

function syncPPTTopbarToolbarRovingTabIndex(
  toolbar: HTMLElement | null,
  focusIndex: number,
) {
  if (!toolbar) {
    return 0
  }

  const allButtons = [...toolbar.querySelectorAll<HTMLButtonElement>(
    '.ppt-toolbar-group button',
  )]
  const enabledButtons = allButtons.filter((button) => !button.disabled)
  const nextFocusIndex = enabledButtons.length === 0
    ? 0
    : clamp(focusIndex, 0, enabledButtons.length - 1)

  allButtons.forEach((button) => {
    const enabledIndex = enabledButtons.indexOf(button)

    if (enabledIndex < 0) {
      button.tabIndex = -1
      button.dataset.pptToolbarDisabledItem = 'true'
      button.removeAttribute('data-ppt-toolbar-item')
      button.removeAttribute('data-ppt-toolbar-focus-index')
      button.removeAttribute('data-ppt-toolbar-active')
      return
    }

    button.tabIndex = enabledIndex === nextFocusIndex ? 0 : -1
    button.dataset.pptToolbarItem = 'true'
    button.dataset.pptToolbarFocusIndex = String(enabledIndex)
    button.dataset.pptToolbarActive = enabledIndex === nextFocusIndex ? 'true' : 'false'
    delete button.dataset.pptToolbarDisabledItem
  })
  toolbar.dataset.pptToolbarItemCount = String(enabledButtons.length)
  toolbar.dataset.pptToolbarActiveIndex = String(nextFocusIndex)

  return nextFocusIndex
}

function getPPTTopbarToolbarItems(toolbar: HTMLElement | null) {
  if (!toolbar) {
    return []
  }

  return [...toolbar.querySelectorAll<HTMLButtonElement>(
    '.ppt-toolbar-group button:not(:disabled)',
  )]
}

function getPPTTopbarToolbarItemFromTarget(
  toolbar: HTMLElement | null,
  target: EventTarget | null,
) {
  if (!toolbar || !(target instanceof HTMLElement)) {
    return null
  }

  const button = target.closest<HTMLButtonElement>('.ppt-toolbar-group button')

  if (!button || button.disabled || !toolbar.contains(button)) {
    return null
  }

  return button
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
  return normalizePPTSlideTransition(slide.transition ?? PPT_DEFAULT_SLIDE_TRANSITION)
}

function normalizePPTSlideTransition(
  transition: PPTSlideTransition,
): PPTSlideTransition {
  return {
    advanceAfterMs: transition.advanceAfterMs === null || transition.advanceAfterMs === undefined
      ? null
      : clampPPTSlideTransitionDuration(transition.advanceAfterMs),
    advanceOnClick: transition.advanceOnClick ?? true,
    durationMs: clampPPTSlideTransitionDuration(transition.durationMs),
    type: PPT_SLIDE_TRANSITION_TYPES.includes(transition.type)
      ? transition.type
      : 'none',
  }
}

function clampPPTSlideTransitionDuration(value: number) {
  return clamp(
    Number.isFinite(value) ? Math.round(value) : 0,
    0,
    PPT_SLIDE_TRANSITION_DURATION_MAX,
  )
}

function parsePPTSlideTransitionDuration(value: string) {
  return clampPPTSlideTransitionDuration(Number(value))
}

function parsePPTSlideTransitionAdvanceAfter(value: string) {
  return value.trim() === ''
    ? null
    : clampPPTSlideTransitionDuration(Number(value))
}

function formatPPTSlideTransitionType(type: PPTSlideTransitionType) {
  switch (type) {
    case 'fade':
      return 'Fade'
    case 'push':
      return 'Push'
    case 'none':
      return 'None'
  }
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

function getPPTElementDefaultAnimationOrder(
  elementId: string,
  slide?: PPTSlide,
) {
  const index = slide?.elements.findIndex((element) => element.id === elementId) ?? -1

  return index >= 0 ? index + 1 : PPT_DEFAULT_ELEMENT_ANIMATION.order
}

function clampPPTElementAnimationTime(value: number) {
  return clamp(
    Number.isFinite(value) ? Math.round(value) : 0,
    0,
    PPT_ELEMENT_ANIMATION_TIME_MAX,
  )
}

function clampPPTElementAnimationOrder(
  value: number,
  elementCount: number = Number.MAX_SAFE_INTEGER,
  fallback: number = PPT_DEFAULT_ELEMENT_ANIMATION.order,
) {
  return clamp(
    Number.isFinite(value) ? Math.round(value) : fallback,
    1,
    Math.max(1, elementCount),
  )
}

function parsePPTElementAnimationTime(value: string) {
  return clampPPTElementAnimationTime(Number(value))
}

function parsePPTElementAnimationOrder(value: string, elementCount: number) {
  return clampPPTElementAnimationOrder(Number(value), elementCount)
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

function normalizePPTParagraphLineHeight(value: number) {
  const next = clamp(
    Number.isFinite(value) ? value : PPT_PARAGRAPH_LINE_HEIGHT_DEFAULT,
    PPT_PARAGRAPH_LINE_HEIGHT_MIN,
    PPT_PARAGRAPH_LINE_HEIGHT_MAX,
  )

  return Math.round(next * 100) / 100
}

function normalizePPTParagraphSpacing(value: number) {
  return clamp(
    Number.isFinite(value) ? Math.round(value) : 0,
    0,
    PPT_PARAGRAPH_SPACING_MAX,
  )
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

function getPPTParagraphStyle(paragraph: PPTParagraph): CSSProperties {
  return {
    lineHeight: getPPTParagraphLineHeight(paragraph),
    marginBottom: getPPTParagraphSpacingAfter(paragraph),
    marginTop: getPPTParagraphSpacingBefore(paragraph),
  }
}

function getPPTElementAnimationStyle(animation: PPTElementAnimation): CSSProperties {
  return {
    '--ppt-animation-delay': `${animation.delayMs}ms`,
    '--ppt-animation-duration': `${Math.max(1, animation.durationMs)}ms`,
  } as CSSProperties
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
  objects: PPTElement[]
  operation?: PPTClipboardOperation
  selectedObjectIds: string[]
  sourceSlideId: string
}): PPTClipboardPayload {
  return {
    elements: objects,
    metadata: objects.map((object) => ({
      groupId: object.groupId ?? null,
      objectId: object.id,
      placeholderId: null,
    })),
    objects,
    operation,
    selectedObjectIds,
    selection: selectedObjectIds,
    sourceSlideId,
    type: 'slide-object-clipboard',
  }
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
  const pastePlan = createPPTClipboardPastePlan({
    createId,
    payload,
    slideFrame,
    target,
  })

  if (!pastePlan) {
    return null
  }

  return {
    payload: {
      id: 'paste-slide-objects',
      pastePlan,
      payload,
    },
    selection: {
      objectIds: pastePlan.mappings.map((mapping) => mapping.targetObjectId),
      slideId: pastePlan.targetSlideId,
    },
    type: 'slide-command-effect',
  }
}

function createPPTClipboardPastePlan({
  createId,
  payload,
  slideFrame,
  target,
}: {
  createId: (prefix: string) => string
  payload: PPTClipboardPayload
  slideFrame: Bounds
  target: PPTClipboardPasteTarget
}): PPTClipboardPastePlan | null {
  if (payload.selectedObjectIds.length === 0 || payload.objects.length === 0) {
    return null
  }

  const mappings = getPPTClipboardPasteObjectMappings({
    createId,
    payload,
  })

  if (mappings.length === 0) {
    return null
  }

  return {
    anchor: getPPTClipboardPasteAnchor({
      slideFrame,
      target,
    }),
    mappings,
    operation: payload.operation,
    sourceSlideId: payload.sourceSlideId,
    targetSlideId: target.slideId,
  }
}

function getPPTClipboardPasteAnchor({
  slideFrame,
  target,
}: {
  slideFrame: Bounds
  target: PPTClipboardPasteTarget
}): Point {
  switch (target.kind) {
    case 'active-slide':
      return {
        x: slideFrame.x,
        y: slideFrame.y,
      }
    case 'pointer-position':
      return target.pointerPosition
    case 'slide-frame-offset':
      return {
        x: slideFrame.x + target.offset.x,
        y: slideFrame.y + target.offset.y,
      }
    case 'viewport-center':
      return target.viewportCenter
  }
}

function getPPTClipboardPasteObjectMappings({
  createId,
  payload,
}: {
  createId: (prefix: string) => string
  payload: PPTClipboardPayload
}): PPTClipboardPasteObjectMapping[] {
  const objectsById = new Map(payload.objects.map((object) => [object.id, object]))
  const metadataById = new Map(payload.metadata.map((metadata) => [metadata.objectId, metadata]))
  const groupIdBySource = new Map<string, string>()

  return payload.selectedObjectIds.flatMap((sourceObjectId) => {
    const object = objectsById.get(sourceObjectId)

    if (!object) {
      return []
    }

    const metadata = metadataById.get(sourceObjectId)
    const sourceGroupId = metadata?.groupId ?? object.groupId ?? null

    return [{
      sourceGroupId,
      sourceObjectId,
      sourcePlaceholderId: metadata?.placeholderId ?? null,
      targetGroupId: getPPTClipboardTargetGroupId({
        createId,
        groupIdBySource,
        sourceGroupId,
      }),
      targetObjectId: createId(getPPTElementIdPrefix(object)),
      targetPlaceholderId: null,
    }]
  })
}

function getPPTClipboardTargetGroupId({
  createId,
  groupIdBySource,
  sourceGroupId,
}: {
  createId: (prefix: string) => string
  groupIdBySource: Map<string, string>
  sourceGroupId: string | null
}) {
  if (!sourceGroupId) {
    return null
  }

  const existing = groupIdBySource.get(sourceGroupId)

  if (existing) {
    return existing
  }

  const targetGroupId = createId('group-copy')
  groupIdBySource.set(sourceGroupId, targetGroupId)

  return targetGroupId
}

function applyPPTClipboardPasteHostCommandEffect(
  slide: PPTSlide,
  effect: PPTClipboardPasteHostCommandEffect,
): PPTSlide {
  const plan = effect.payload.pastePlan
  const objectById = new Map(effect.payload.payload.objects.map((object) => [object.id, object]))
  const offset = {
    x: plan.anchor.x,
    y: plan.anchor.y,
  }
  const pasted = plan.mappings.flatMap((mapping) => {
    const source = objectById.get(mapping.sourceObjectId)

    return source
      ? [clonePPTElementFromClipboardMapping(source, mapping, offset)]
      : []
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
  return {
    ...source,
    geometry: {
      ...source.geometry,
      x: clamp(source.geometry.x + offset.x, 0, PPT_SLIDE_WIDTH - source.geometry.w),
      y: clamp(source.geometry.y + offset.y, 0, PPT_SLIDE_HEIGHT - source.geometry.h),
    },
    groupId: mapping.targetGroupId ?? undefined,
    id: mapping.targetObjectId,
    name: `${source.name} Copy`,
  }
}

function createPPTLayerPaneDescriptor({
  activeObjectId = null,
  selectedObjectIds,
  slide,
}: {
  activeObjectId?: string | null
  selectedObjectIds: readonly string[]
  slide: PPTSlide
}): PPTLayerPaneDescriptor {
  const selected = new Set(selectedObjectIds)
  const rowCount = slide.elements.length

  return {
    activeObjectId,
    aria: PPT_LAYER_PANE_ARIA_CONTRACT,
    rows: slide.elements.map((element, index) => ({
      ariaLevel: 1,
      ariaPosInSet: index + 1,
      ariaSetSize: rowCount,
      displayName: element.name,
      groupId: element.groupId ?? null,
      isGrouped: Boolean(element.groupId),
      isGroup: false,
      isHidden: element.visible === false,
      isLocked: element.locked === true,
      isRenamable: true,
      isReorderable: element.locked !== true,
      isSelectable: true,
      isSelected: selected.has(element.id),
      kindLabel: getPPTElementKindLabel(element),
      objectId: element.id,
      order: index,
      parentObjectId: null,
      slideId: slide.id,
    })),
    selectedObjectIds,
    slideId: slide.id,
  }
}

function getPPTLayerPaneCommandEffect(
  descriptor: PPTLayerPaneDescriptor,
  intent: PPTLayerPaneIntent,
): PPTLayerPaneHostCommandEffect | null {
  switch (intent.type) {
    case 'lock-toggle':
      return getPPTLayerPaneLockEffect(descriptor, intent.objectId)
    case 'rename-submit':
      return getPPTLayerPaneRenameEffect(descriptor, intent)
    case 'row-drop':
      return getPPTLayerPaneReorderEffect(descriptor, intent)
    case 'row-press':
      return getPPTLayerPaneSelectEffect(descriptor, intent)
    case 'visibility-toggle':
      return getPPTLayerPaneVisibilityEffect(descriptor, intent.objectId)
  }
}

function getPPTLayerPaneSelectEffect(
  descriptor: PPTLayerPaneDescriptor,
  intent: Extract<PPTLayerPaneIntent, { type: 'row-press' }>,
): PPTLayerPaneHostCommandEffect | null {
  const row = findPPTLayerPaneRow(descriptor, intent.objectId)

  if (!row?.isSelectable) {
    return null
  }

  const objectIds = intent.rangeAnchorObjectId
    ? getPPTLayerPaneRangeSelection(
        descriptor,
        intent.rangeAnchorObjectId,
        intent.objectId,
      )
    : intent.additive === true
      ? togglePPTLayerPaneSelection(
          descriptor.selectedObjectIds,
          intent.objectId,
        )
      : [intent.objectId]
  const mode = intent.rangeAnchorObjectId
    ? 'range'
    : intent.additive === true
      ? 'additive'
      : 'replace'

  return toPPTLayerPaneHostCommandEffect({
    descriptor,
    payload: {
      id: 'select-objects',
      mode,
      objectIds,
    },
    selectionObjectIds: objectIds,
  })
}

function getPPTLayerPaneRenameEffect(
  descriptor: PPTLayerPaneDescriptor,
  intent: Extract<PPTLayerPaneIntent, { type: 'rename-submit' }>,
): PPTLayerPaneHostCommandEffect | null {
  const row = findPPTLayerPaneRow(descriptor, intent.objectId)
  const name = intent.name.trim()

  if (!row?.isRenamable || name.length === 0) {
    return null
  }

  return toPPTLayerPaneHostCommandEffect({
    descriptor,
    payload: {
      id: 'rename-object',
      name,
      objectId: intent.objectId,
    },
    selectionObjectIds: [intent.objectId],
  })
}

function getPPTLayerPaneVisibilityEffect(
  descriptor: PPTLayerPaneDescriptor,
  objectId: string,
): PPTLayerPaneHostCommandEffect | null {
  const row = findPPTLayerPaneRow(descriptor, objectId)

  if (!row || row.isLocked) {
    return null
  }

  return toPPTLayerPaneHostCommandEffect({
    descriptor,
    payload: {
      id: row.isHidden ? 'show-objects' : 'hide-objects',
      objectIds: [objectId],
    },
    selectionObjectIds: [objectId],
  })
}

function getPPTLayerPaneLockEffect(
  descriptor: PPTLayerPaneDescriptor,
  objectId: string,
): PPTLayerPaneHostCommandEffect | null {
  const row = findPPTLayerPaneRow(descriptor, objectId)

  if (!row) {
    return null
  }

  return toPPTLayerPaneHostCommandEffect({
    descriptor,
    payload: {
      id: row.isLocked ? 'unlock-objects' : 'lock-objects',
      objectIds: [objectId],
    },
    selectionObjectIds: [objectId],
  })
}

function getPPTLayerPaneReorderEffect(
  descriptor: PPTLayerPaneDescriptor,
  intent: Extract<PPTLayerPaneIntent, { type: 'row-drop' }>,
): PPTLayerPaneHostCommandEffect | null {
  const row = findPPTLayerPaneRow(descriptor, intent.objectId)
  const fromIndex = descriptor.rows.findIndex((row) => row.objectId === intent.objectId)
  const toIndex = Math.max(0, Math.min(intent.toIndex, Math.max(0, descriptor.rows.length - 1)))

  if (!row?.isReorderable || fromIndex < 0 || fromIndex === toIndex) {
    return null
  }

  return toPPTLayerPaneHostCommandEffect({
    descriptor,
    payload: {
      fromIndex,
      id: 'reorder-object',
      objectId: intent.objectId,
      toIndex,
    },
    selectionObjectIds: [intent.objectId],
  })
}

function toPPTLayerPaneHostCommandEffect({
  descriptor,
  payload,
  selectionObjectIds,
}: {
  descriptor: PPTLayerPaneDescriptor
  payload: PPTLayerPaneCommand
  selectionObjectIds: readonly string[]
}): PPTLayerPaneHostCommandEffect {
  return {
    payload,
    selection: {
      objectIds: selectionObjectIds,
      slideId: descriptor.slideId,
    },
    type: 'slide-command-effect',
  }
}

function findPPTLayerPaneRow(
  descriptor: PPTLayerPaneDescriptor,
  objectId: string,
) {
  return descriptor.rows.find((row) => row.objectId === objectId) ?? null
}

function togglePPTLayerPaneSelection(
  selectedObjectIds: readonly string[],
  objectId: string,
) {
  return selectedObjectIds.includes(objectId)
    ? selectedObjectIds.filter((selectedId) => selectedId !== objectId)
    : [...selectedObjectIds, objectId]
}

function getPPTLayerPaneRangeSelection(
  descriptor: PPTLayerPaneDescriptor,
  anchorObjectId: string,
  objectId: string,
) {
  const anchorIndex = descriptor.rows.findIndex((row) => row.objectId === anchorObjectId)
  const targetIndex = descriptor.rows.findIndex((row) => row.objectId === objectId)

  if (anchorIndex < 0 || targetIndex < 0) {
    return [objectId]
  }

  const start = Math.min(anchorIndex, targetIndex)
  const end = Math.max(anchorIndex, targetIndex)

  return descriptor.rows
    .slice(start, end + 1)
    .filter((row) => row.isSelectable)
    .map((row) => row.objectId)
}

function getPPTMinimapReadModel({
  items,
  size = PPT_MINIMAP_SIZE,
  stageSize,
  viewport,
}: {
  items: readonly PPTMinimapItemBounds[]
  size?: PPTMinimapSize
  stageSize: PPTMinimapSize
  viewport: Viewport
}): PPTMinimapReadModel {
  const viewportWorldBounds = getPPTMinimapViewportWorldBounds({
    stageSize,
    viewport,
  })
  const contentBounds = items.reduce<Bounds>(
    (bounds, item) => unionPPTMinimapBounds(bounds, normalizePPTMinimapBounds(item.bounds)),
    {
      h: PPT_SLIDE_HEIGHT,
      w: PPT_SLIDE_WIDTH,
      x: 0,
      y: 0,
    },
  )
  const worldBounds = expandPPTMinimapWorldBounds(
    unionPPTMinimapBounds(contentBounds, viewportWorldBounds),
  )
  const displayBounds = getPPTMinimapDisplayBounds({
    size,
    worldBounds,
  })

  return {
    displayBounds,
    itemRects: items.map((item) => ({
      id: item.id,
      rect: pptWorldBoundsToMinimapRect({
        displayBounds,
        rect: item.bounds,
        worldBounds,
      }),
    })),
    size,
    viewportRect: pptWorldBoundsToMinimapRect({
      displayBounds,
      rect: viewportWorldBounds,
      worldBounds,
    }),
    viewportWorldBounds,
    worldBounds,
  }
}

function getPPTMinimapViewportForWorldCenter({
  current,
  stageSize,
  worldCenter,
}: {
  current: Viewport
  stageSize: PPTMinimapSize
  worldCenter: Point
}): Viewport {
  return {
    scale: current.scale,
    x: stageSize.w / 2 - worldCenter.x * current.scale,
    y: stageSize.h / 2 - worldCenter.y * current.scale,
  }
}

function getPPTMinimapWorldPoint({
  model,
  point,
}: {
  model: PPTMinimapReadModel
  point: Point
}): Point {
  const scale = getPPTMinimapScale(model.worldBounds, model.displayBounds)

  return {
    x: model.worldBounds.x + (point.x - model.displayBounds.x) / scale,
    y: model.worldBounds.y + (point.y - model.displayBounds.y) / scale,
  }
}

function getPPTMinimapSvgPoint(
  event: ReactPointerEvent<SVGSVGElement>,
  svg: SVGSVGElement | null,
  model: PPTMinimapReadModel,
): Point | null {
  const rect = svg?.getBoundingClientRect()

  if (!rect || rect.width <= 0 || rect.height <= 0) {
    return null
  }

  return {
    x: (event.clientX - rect.left) * model.size.w / rect.width,
    y: (event.clientY - rect.top) * model.size.h / rect.height,
  }
}

function getPPTMinimapViewportWorldBounds({
  stageSize,
  viewport,
}: {
  stageSize: PPTMinimapSize
  viewport: Viewport
}): Bounds {
  const scale = Math.max(Number.isFinite(viewport.scale) ? viewport.scale : 1, 0.001)
  const viewportX = Number.isFinite(viewport.x) ? viewport.x : 0
  const viewportY = Number.isFinite(viewport.y) ? viewport.y : 0

  return {
    h: Math.max(1, stageSize.h / scale),
    w: Math.max(1, stageSize.w / scale),
    x: -viewportX / scale,
    y: -viewportY / scale,
  }
}

function expandPPTMinimapWorldBounds(bounds: Bounds): Bounds {
  const width = Math.max(bounds.w, PPT_MINIMAP_MIN_WORLD_SIZE)
  const height = Math.max(bounds.h, PPT_MINIMAP_MIN_WORLD_SIZE)
  const centerX = bounds.x + bounds.w / 2
  const centerY = bounds.y + bounds.h / 2

  return {
    h: height,
    w: width,
    x: centerX - width / 2,
    y: centerY - height / 2,
  }
}

function getPPTMinimapDisplayBounds({
  size,
  worldBounds,
}: {
  size: PPTMinimapSize
  worldBounds: Bounds
}): Bounds {
  const availableWidth = Math.max(1, size.w - PPT_MINIMAP_PADDING * 2)
  const availableHeight = Math.max(1, size.h - PPT_MINIMAP_PADDING * 2)
  const scale = Math.min(
    availableWidth / Math.max(1, worldBounds.w),
    availableHeight / Math.max(1, worldBounds.h),
  )
  const width = worldBounds.w * scale
  const height = worldBounds.h * scale

  return {
    h: height,
    w: width,
    x: (size.w - width) / 2,
    y: (size.h - height) / 2,
  }
}

function pptWorldBoundsToMinimapRect({
  displayBounds,
  rect,
  worldBounds,
}: {
  displayBounds: Bounds
  rect: Bounds
  worldBounds: Bounds
}): Bounds {
  const scale = getPPTMinimapScale(worldBounds, displayBounds)

  return {
    h: Math.max(1, rect.h * scale),
    w: Math.max(1, rect.w * scale),
    x: displayBounds.x + (rect.x - worldBounds.x) * scale,
    y: displayBounds.y + (rect.y - worldBounds.y) * scale,
  }
}

function getPPTMinimapScale(worldBounds: Bounds, displayBounds: Bounds) {
  return Math.min(
    displayBounds.w / Math.max(1, worldBounds.w),
    displayBounds.h / Math.max(1, worldBounds.h),
  )
}

function unionPPTMinimapBounds(left: Bounds, right: Bounds): Bounds {
  const x = Math.min(left.x, right.x)
  const y = Math.min(left.y, right.y)
  const rightEdge = Math.max(left.x + left.w, right.x + right.w)
  const bottomEdge = Math.max(left.y + left.h, right.y + right.h)

  return {
    h: bottomEdge - y,
    w: rightEdge - x,
    x,
    y,
  }
}

function normalizePPTMinimapBounds(bounds: Bounds): Bounds {
  return {
    h: Math.max(1, bounds.h),
    w: Math.max(1, bounds.w),
    x: bounds.x,
    y: bounds.y,
  }
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
  groups,
  onCommand,
  onFontSizeStep,
  onParagraphAlign,
  onParagraphBulletToggle,
  onTextBoldToggle,
  onTextColorChange,
  onTextItalicToggle,
  onTextUnderlineToggle,
  scale,
  textFormat,
}: {
  anchor: PPTSelectionCommandAnchor | null
  groups: readonly PPTSurfaceCommandViewGroup[]
  onCommand: (command: PPTSurfaceCommand) => void
  onFontSizeStep: (delta: number) => void
  onParagraphAlign: (align: NonNullable<PPTParagraph['align']>) => void
  onParagraphBulletToggle: () => void
  onTextBoldToggle: () => void
  onTextColorChange: (color: string) => void
  onTextItalicToggle: () => void
  onTextUnderlineToggle: () => void
  scale: number
  textFormat: PPTTextQuickFormatState | null
}) {
  if (!anchor || (groups.length === 0 && !textFormat)) {
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
          onTextBoldToggle={onTextBoldToggle}
          onTextColorChange={onTextColorChange}
          onTextItalicToggle={onTextItalicToggle}
          onTextUnderlineToggle={onTextUnderlineToggle}
        />
      ) : null}
      {textFormat && groups.length > 0 ? <span className="ppt-command-divider" /> : null}
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

function PPTTextQuickFormatControls({
  onFontSizeStep,
  onParagraphAlign,
  onParagraphBulletToggle,
  onTextBoldToggle,
  onTextColorChange,
  onTextItalicToggle,
  onTextUnderlineToggle,
  state,
}: {
  onFontSizeStep: (delta: number) => void
  onParagraphAlign: (align: NonNullable<PPTParagraph['align']>) => void
  onParagraphBulletToggle: () => void
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
  const selectedIndex = Math.max(0, PPT_PARAGRAPH_ALIGN_OPTIONS.indexOf(align))

  function selectAlign(
    nextAlign: NonNullable<PPTParagraph['align']>,
    container: HTMLElement | null,
  ) {
    onAlignChange(nextAlign)
    window.requestAnimationFrame(() => {
      container
        ?.querySelector<HTMLButtonElement>(`[data-ppt-paragraph-align="${nextAlign}"]`)
        ?.focus({ preventScroll: true })
    })
  }

  function handleKeyDown(event: ReactKeyboardEvent<HTMLElement>) {
    if (event.altKey || event.ctrlKey || event.metaKey) {
      return
    }

    let nextIndex: number | null = null

    if (event.key === 'ArrowRight' || event.key === 'ArrowDown') {
      nextIndex = (selectedIndex + 1) % PPT_PARAGRAPH_ALIGN_OPTIONS.length
    } else if (event.key === 'ArrowLeft' || event.key === 'ArrowUp') {
      nextIndex = (selectedIndex - 1 + PPT_PARAGRAPH_ALIGN_OPTIONS.length) %
        PPT_PARAGRAPH_ALIGN_OPTIONS.length
    } else if (event.key === 'Home') {
      nextIndex = 0
    } else if (event.key === 'End') {
      nextIndex = PPT_PARAGRAPH_ALIGN_OPTIONS.length - 1
    }

    if (nextIndex === null) {
      return
    }

    event.preventDefault()
    event.stopPropagation()
    selectAlign(PPT_PARAGRAPH_ALIGN_OPTIONS[nextIndex], event.currentTarget)
  }

  return (
    <span
      aria-label="Paragraph align"
      className={surface === 'quick'
        ? 'ppt-paragraph-align-radio-group ppt-paragraph-align-radio-group--quick'
        : 'ppt-segmented-control ppt-paragraph-align-radio-group'}
      data-ppt-paragraph-align-focus-model="roving-tabindex"
      data-ppt-paragraph-align-keyboard-model="arrow-home-end"
      data-ppt-paragraph-align-radiogroup={surface}
      role="radiogroup"
      onKeyDown={handleKeyDown}
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
            tabIndex={selected ? 0 : -1}
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
  const menuRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    if (!menu) {
      return
    }

    const frame = window.requestAnimationFrame(() => {
      getPPTContextMenuItems(menuRef.current)[0]?.focus()
    })

    return () => window.cancelAnimationFrame(frame)
  }, [menu])

  function focusContextMenuItem(delta: number) {
    const items = getPPTContextMenuItems(menuRef.current)

    if (items.length === 0) {
      return
    }

    const currentIndex = items.findIndex((item) => item === document.activeElement)
    const nextIndex = currentIndex < 0
      ? 0
      : (currentIndex + delta + items.length) % items.length

    items[nextIndex]?.focus()
  }

  function focusContextMenuBoundary(position: 'first' | 'last') {
    const items = getPPTContextMenuItems(menuRef.current)
    const item = position === 'first' ? items[0] : items.at(-1)

    item?.focus()
  }

  function handleKeyDown(event: ReactKeyboardEvent<HTMLDivElement>) {
    if (event.key === 'Escape') {
      event.preventDefault()
      event.stopPropagation()
      onClose()
      return
    }

    if (event.key === 'ArrowDown') {
      event.preventDefault()
      focusContextMenuItem(1)
      return
    }

    if (event.key === 'ArrowUp') {
      event.preventDefault()
      focusContextMenuItem(-1)
      return
    }

    if (event.key === 'Home') {
      event.preventDefault()
      focusContextMenuBoundary('first')
      return
    }

    if (event.key === 'End') {
      event.preventDefault()
      focusContextMenuBoundary('last')
      return
    }

    if (event.key === 'Enter' || event.key === ' ') {
      const focused = document.activeElement

      if (focused instanceof HTMLButtonElement &&
        focused.closest('[data-ppt-context-menu]')) {
        event.preventDefault()
        focused.click()
      }
    }
  }

  if (!menu || groups.length === 0) {
    return null
  }

  return (
    <div
      aria-label="Selection commands"
      className="ppt-context-menu"
      data-ppt-context-menu
      data-ppt-context-menu-focus-model="enabled-menuitem-roving"
      data-ppt-context-menu-keyboard="arrow-home-end-enter-escape"
      ref={menuRef}
      role="menu"
      style={{
        left: menu.x,
        top: menu.y,
      }}
      onContextMenu={(event) => event.preventDefault()}
      onKeyDown={handleKeyDown}
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
    ? { 'data-ppt-context-command': command.dataCommand }
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

function getPPTContextMenuItems(menu: HTMLDivElement | null) {
  if (!menu) {
    return []
  }

  return [...menu.querySelectorAll<HTMLButtonElement>(
    '[role="menuitem"]:not(:disabled)',
  )]
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
  onDragEnd,
  onDragOver,
  onDragStart,
  onDrop,
  onKeyDown,
  onSelect,
  slide,
}: {
  active: boolean
  dragging: boolean
  dropPlacement?: PPTSlideDropPlacement
  index: number
  onDragEnd: (event: ReactDragEvent<HTMLButtonElement>) => void
  onDragOver: (event: ReactDragEvent<HTMLButtonElement>) => void
  onDragStart: (event: ReactDragEvent<HTMLButtonElement>) => void
  onDrop: (event: ReactDragEvent<HTMLButtonElement>) => void
  onKeyDown: (event: ReactKeyboardEvent<HTMLButtonElement>) => void
  onSelect: (event: ReactMouseEvent<HTMLButtonElement>) => void
  slide: PPTSlide
}) {
  return (
    <button
      aria-current={active ? 'page' : undefined}
      aria-label={`Open ${slide.name}`}
      aria-selected={active}
      className="ppt-thumb"
      data-ppt-slide-active={active ? 'true' : undefined}
      data-ppt-slide-dragging={dragging ? 'true' : undefined}
      data-ppt-slide-draggable="true"
      data-ppt-slide-drop-target={dropPlacement}
      data-ppt-slide-id={slide.id}
      data-ppt-slide-index={index}
      data-ppt-slide-roving-tab-index={active ? '0' : '-1'}
      draggable
      onDragEnd={onDragEnd}
      onDragOver={onDragOver}
      onDragStart={onDragStart}
      onDrop={onDrop}
      onClick={onSelect}
      onKeyDown={onKeyDown}
      role="option"
      tabIndex={active ? 0 : -1}
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
                ? getPPTThumbShapeCornerRadiusCSS(element)
                : undefined,
              backgroundImage: element.kind === 'image'
                ? `url(${element.src})`
                : undefined,
              backgroundPosition: element.kind === 'image'
                ? getPPTImageObjectPosition(element)
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
      <span className="ppt-thumb-name">{index + 1}. {slide.name}</span>
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
  onPointerDown,
  onPointerEnter,
  onPointerLeave,
  onStopEdit,
  onTextOverflowChange,
  selected,
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
  onPointerDown: (event: ReactPointerEvent<HTMLDivElement>, elementId: string) => void
  onPointerEnter: () => void
  onPointerLeave: () => void
  onStopEdit: () => void
  onTextOverflowChange: (elementId: string, hasOverflow: boolean) => void
  selected: boolean
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

    editorRef.current?.focus()
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
      const container = editor.parentElement
      const editorRect = editor.getBoundingClientRect()
      const containerRect = container?.getBoundingClientRect()
      const hasOverflow =
        editor.scrollWidth > editor.clientWidth + PPT_TEXT_OVERFLOW_EPSILON ||
        editor.scrollHeight > editor.clientHeight + PPT_TEXT_OVERFLOW_EPSILON ||
        (containerRect
          ? editorRect.width > containerRect.width + PPT_TEXT_OVERFLOW_EPSILON ||
            editorRect.height > containerRect.height + PPT_TEXT_OVERFLOW_EPSILON
          : false)

      onTextOverflowChange(element.id, hasOverflow)
    }

    updateOverflow()
    const frame = window.requestAnimationFrame(updateOverflow)

    return () => window.cancelAnimationFrame(frame)
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

  return (
    <div
      className="ppt-element"
      data-hovered={hovered ? 'true' : 'false'}
      data-group-id={element.groupId}
      data-kind={element.kind}
      data-ppt-comment-resolved={element.kind === 'comment' && element.resolved === true
        ? 'true'
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
      data-ppt-text-overflow={textOverflow ? 'true' : undefined}
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
      data-locked={element.locked === true ? 'true' : 'false'}
      data-ppt-element={element.id}
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
            objectPosition: getPPTImageObjectPosition(element),
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
          ref={editorRef}
          suppressContentEditableWarning={true}
          onBlur={(event) => {
            onCommitText(element.id, event.currentTarget.innerText)
            onStopEdit()
          }}
          onKeyDown={(event) => {
            if (event.key === 'Escape') {
              event.preventDefault()
              event.currentTarget.innerText = text
              event.currentTarget.blur()
              onStopEdit()
            }
            if ((event.metaKey || event.ctrlKey) && event.key === 'Enter') {
              event.preventDefault()
              event.currentTarget.blur()
            }
          }}
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

  return (
    <div
      className="ppt-comment-card"
      data-ppt-comment-card
      data-ppt-comment-resolved={element.resolved === true ? 'true' : undefined}
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

function PPTTextBodyView({ body }: { body: PPTTextBody }) {
  return (
    <>
      {body.paragraphs.map((paragraph, index) => (
        <span
          className="ppt-text-paragraph"
          data-ppt-bullet={paragraph.bullet === 'bullet' ? 'true' : undefined}
          data-ppt-line-height={getPPTParagraphLineHeight(paragraph)}
          data-ppt-spacing-after={getPPTParagraphSpacingAfter(paragraph)}
          data-ppt-spacing-before={getPPTParagraphSpacingBefore(paragraph)}
          key={index}
          style={getPPTParagraphStyle(paragraph)}
        >
          {paragraph.runs.map((run, runIndex) => (
            <span
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
        d={getPPTFreeformPathData(element.points)}
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
  onResizeHandleDoubleClick,
  onResizePointerDown,
  scale,
  selectedElements,
  textOverflow,
}: {
  bounds: Bounds
  canResize: boolean
  onRotatePointerDown: (event: ReactPointerEvent<HTMLButtonElement>) => void
  onResizeHandleDoubleClick: (
    event: ReactMouseEvent<HTMLButtonElement>,
    handle: ResizeHandle,
  ) => void
  onResizePointerDown: (
    event: ReactPointerEvent<HTMLButtonElement>,
    handle: ResizeHandle,
  ) => void
  scale: number
  selectedElements: PPTElement[]
  textOverflow: boolean
}) {
  return (
    <>
      <Box className="ppt-selection-box" bounds={bounds} />
      <div
        className="ppt-size-capsule"
        data-ppt-text-overflow={textOverflow ? 'true' : undefined}
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
      {canResize ? RESIZE_HANDLES.map((handle) => {
        const point = handlePoint(bounds, handle)
        const size = 10 / scale

        return (
          <button
            aria-label={`Resize ${handle}`}
            className="ppt-resize-handle"
            key={handle}
            onDoubleClick={(event) => onResizeHandleDoubleClick(event, handle)}
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
        d={getPPTLaserTrailPathData(points)}
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

function Guides({ guides, scale }: { guides: CanvasSnapGuides; scale: number }) {
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
  recentColors,
  themeColorTokens,
  onSelect,
}: {
  channel: PPTColorSwatchChannel
  currentColor: string
  recentColors: readonly string[]
  themeColorTokens: readonly SlideEditThemeColorToken[]
  onSelect: (color: string) => void
}) {
  const normalizedCurrent = normalizePPTSwatchColor(currentColor)
  const normalizedRecentColors = getPPTUniqueSwatchColors(recentColors)

  return (
    <div
      className="ppt-color-swatch-strip"
      data-ppt-color-swatch-channel={channel}
      data-ppt-color-swatch-count={themeColorTokens.length + normalizedRecentColors.length}
      data-ppt-color-swatch-strip={channel}
      data-ppt-recent-color-count={normalizedRecentColors.length}
    >
      <div className="ppt-color-swatch-group" data-ppt-color-swatch-group="theme">
        {themeColorTokens.map((token) => {
          const color = normalizePPTSwatchColor(token.value) || token.value
          const selected = normalizedCurrent === normalizePPTSwatchColor(color)

          return (
            <button
              aria-label={`${token.label} ${channel}`}
              aria-pressed={selected}
              className="ppt-color-swatch"
              data-ppt-color-source="theme"
              data-ppt-color-swatch={channel}
              data-ppt-color-swatch-selected={selected ? 'true' : undefined}
              data-ppt-color-token={token.tokenId}
              data-ppt-color-value={color}
              key={token.tokenId}
              style={{ backgroundColor: color }}
              title={token.label}
              type="button"
              onClick={() => onSelect(color)}
            />
          )
        })}
      </div>
      {normalizedRecentColors.length > 0 ? (
        <div className="ppt-color-swatch-group" data-ppt-color-swatch-group="recent">
          {normalizedRecentColors.map((color) => {
            const selected = normalizedCurrent === color

            return (
              <button
                aria-label={`Recent ${color} ${channel}`}
                aria-pressed={selected}
                className="ppt-color-swatch"
                data-ppt-color-source="recent"
                data-ppt-color-swatch={channel}
                data-ppt-color-swatch-selected={selected ? 'true' : undefined}
                data-ppt-color-value={color}
                key={color}
                style={{ backgroundColor: color }}
                title={color}
                type="button"
                onClick={() => onSelect(color)}
              />
            )
          })}
        </div>
      ) : null}
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
  onCommentBodyChange,
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
  onImageFitChange,
  onLayerPaneCommandEffect,
  onLayoutPlaceholderVisibilityChange,
  onLineMarkerChange,
  onLineRouteChange,
  onParagraphBulletChange,
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
  onCommentBodyChange: (elementId: string, value: string) => void
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
  onImageFitChange: (
    elementId: string,
    fit: PPTImageFit,
  ) => void
  onLayerPaneCommandEffect: (effect: PPTLayerPaneHostCommandEffect) => void
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
  const paragraphSpacing = selectedElement && isPPTTextElement(selectedElement)
    ? getPPTTextElementParagraphSpacing(selectedElement)
    : getDefaultPPTParagraphSpacing()
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
  const textInset = selectedElement && isPPTTextElement(selectedElement)
    ? getPPTTextElementInset(selectedElement)
    : PPT_DEFAULT_TEXT_BOX_INSET
  const nameMetadataField = getPPTSlideMetadataField(slideMetadataDescriptor, 'name')
  const backgroundMetadataField = getPPTSlideMetadataField(slideMetadataDescriptor, 'background')
  const notesMetadataField = getPPTSlideMetadataField(slideMetadataDescriptor, 'notes')
  const sizeMetadataField = getPPTSlideMetadataField(slideMetadataDescriptor, 'size')
  const orientationMetadataField = getPPTSlideMetadataField(slideMetadataDescriptor, 'orientation')
  const layerPaneDescriptor = createPPTLayerPaneDescriptor({
    activeObjectId: selectedElement?.id ?? null,
    selectedObjectIds: selection,
    slide,
  })
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
  const inspectorTabs: readonly {
    id: PPTInspectorTabId
    label: string
    panelId: string
    tabId: string
  }[] = [
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
  ]

  useEffect(() => {
    if (previousInspectorSelectionStateRef.current === hasSelectedElement) {
      return
    }

    previousInspectorSelectionStateRef.current = hasSelectedElement
    setActiveInspectorTabId(hasSelectedElement ? 'selection' : 'slide')
  }, [hasSelectedElement])

  function runLayerPaneIntent(intent: PPTLayerPaneIntent) {
    const effect = getPPTLayerPaneCommandEffect(layerPaneDescriptor, intent)

    if (effect) {
      onLayerPaneCommandEffect(effect)
    }
  }

  function focusLayerPaneRow(objectId: string) {
    window.requestAnimationFrame(() => {
      document
        .querySelector<HTMLElement>(`[data-ppt-layer-pane-row="${objectId}"]`)
        ?.focus({ preventScroll: true })
    })
  }

  function focusPPTInspectorTab(tabId: PPTInspectorTabId) {
    window.requestAnimationFrame(() => {
      document
        .querySelector<HTMLButtonElement>(`[data-ppt-inspector-tab="${tabId}"]`)
        ?.focus({ preventScroll: true })
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
      event.altKey ||
      event.ctrlKey ||
      event.metaKey
    ) {
      return
    }

    const currentIndex = inspectorTabs.findIndex((tab) => tab.id === tabId)

    if (currentIndex < 0) {
      return
    }

    let nextIndex: number | null = null

    if (event.key === 'ArrowRight' || event.key === 'ArrowDown') {
      nextIndex = (currentIndex + 1) % inspectorTabs.length
    } else if (event.key === 'ArrowLeft' || event.key === 'ArrowUp') {
      nextIndex = (currentIndex - 1 + inspectorTabs.length) % inspectorTabs.length
    } else if (event.key === 'Home') {
      nextIndex = 0
    } else if (event.key === 'End') {
      nextIndex = inspectorTabs.length - 1
    } else if (event.key === 'Enter' || event.key === ' ') {
      nextIndex = currentIndex
    }

    if (nextIndex === null) {
      return
    }

    event.preventDefault()
    event.stopPropagation()

    const nextTab = inspectorTabs[nextIndex]
    setActiveInspectorTabId(nextTab.id)
    focusPPTInspectorTab(nextTab.id)
  }

  function handleLayerPaneRowKeyDown(
    row: PPTLayerPaneRowDescriptor,
    event: ReactKeyboardEvent<HTMLElement>,
  ) {
    if (
      event.target !== event.currentTarget ||
      event.altKey ||
      event.ctrlKey ||
      event.metaKey
    ) {
      return
    }

    const rows = layerPaneDescriptor.rows
    const currentIndex = rows.findIndex((item) => item.objectId === row.objectId)

    if (currentIndex < 0) {
      return
    }

    let nextRow: PPTLayerPaneRowDescriptor | null = null

    if (event.key === 'ArrowDown') {
      nextRow = rows[Math.min(currentIndex + 1, rows.length - 1)] ?? null
    } else if (event.key === 'ArrowUp') {
      nextRow = rows[Math.max(currentIndex - 1, 0)] ?? null
    } else if (event.key === 'Home') {
      nextRow = rows[0] ?? null
    } else if (event.key === 'End') {
      nextRow = rows.at(-1) ?? null
    } else if (event.key === 'Enter' || event.key === ' ') {
      nextRow = row
    }

    if (!nextRow) {
      return
    }

    event.preventDefault()
    event.stopPropagation()
    runLayerPaneIntent({
      objectId: nextRow.objectId,
      type: 'row-press',
    })
    focusLayerPaneRow(nextRow.objectId)
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
        data-ppt-inspector-tabs-activation="automatic"
        data-ppt-inspector-tabs-keyboard="arrow-home-end-enter-space"
        role="tablist"
      >
        {inspectorTabs.map((tab) => {
          const active = activeInspectorTabId === tab.id

          return (
            <button
              aria-controls={tab.panelId}
              aria-selected={active}
              className="ppt-inspector-tab"
              data-ppt-inspector-tab={tab.id}
              data-ppt-inspector-tab-active={active ? 'true' : 'false'}
              id={tab.tabId}
              key={tab.id}
              role="tab"
              tabIndex={active ? 0 : -1}
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
        aria-labelledby="ppt-inspector-tab-slide"
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
        hidden={activeInspectorTabId !== 'slide'}
        id="ppt-inspector-panel-slide"
        role="tabpanel"
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
          data-ppt-transition-advance-after={slideTransition.advanceAfterMs ?? ''}
          data-ppt-transition-advance-on-click={slideTransition.advanceOnClick ? 'true' : 'false'}
          data-ppt-transition-duration={slideTransition.durationMs}
          data-ppt-transition-type={slideTransition.type}
        >
          <label className="ppt-field">
            <span>Transition</span>
            <select
              data-ppt-slide-transition-field="type"
              value={slideTransition.type}
              onChange={(event) =>
                onSlideTransitionChange('type', event.target.value as PPTSlideTransitionType)}
            >
              {PPT_SLIDE_TRANSITION_TYPES.map((type) => (
                <option key={type} value={type}>
                  {formatPPTSlideTransitionType(type)}
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
              max={PPT_SLIDE_TRANSITION_DURATION_MAX}
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
        aria-labelledby="ppt-inspector-tab-selection"
        className="ppt-panel-section"
        data-ppt-inspector-surface={inspectorSurface}
        data-ppt-inspector-tabpanel="selection"
        data-ppt-inspector-tabpanel-active={activeInspectorTabId === 'selection' ? 'true' : 'false'}
        data-ppt-object-inspector
        data-ppt-object-inspector-active={selectedElement ? 'true' : 'false'}
        data-ppt-object-inspector-priority={inspectorSurface === 'object-selection-inspector' ? '0' : '1'}
        hidden={activeInspectorTabId !== 'selection'}
        id="ppt-inspector-panel-selection"
        role="tabpanel"
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
                max={PPT_ELEMENT_OPACITY_MAX}
                min={PPT_ELEMENT_OPACITY_MIN}
                step={PPT_ELEMENT_OPACITY_STEP}
                type="number"
                value={getPPTElementOpacity(selectedElement)}
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
                maxLength={PPT_HYPERLINK_URL_MAX_LENGTH}
                placeholder="https://example.com"
                value={elementHyperlink?.url ?? ''}
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
                maxLength={PPT_ALT_TEXT_MAX_LENGTH}
                value={elementAltText}
                onChange={(event) =>
                  onElementAltTextChange(
                    selectedElement.id,
                    event.target.value,
                  )}
              />
            </label>
            <label className="ppt-checkbox-field">
              <input
                checked={elementShadowEnabled}
                data-ppt-shadow-field="enabled"
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
              data-ppt-shadow-angle={elementShadow.angle}
              data-ppt-shadow-blur={elementShadow.blur}
              data-ppt-shadow-color={elementShadow.color}
              data-ppt-shadow-distance={elementShadow.distance}
              data-ppt-shadow-enabled={elementShadowEnabled ? 'true' : 'false'}
              data-ppt-shadow-inspector
              data-ppt-shadow-opacity={elementShadow.opacity}
            >
              <label className="ppt-field">
                <span>Color</span>
                <input
                  data-ppt-shadow-field="color"
                  disabled={!elementShadowEnabled}
                  type="color"
                  value={elementShadow.color}
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
                  data-ppt-shadow-field="opacity"
                  disabled={!elementShadowEnabled}
                  max={PPT_ELEMENT_SHADOW_OPACITY_MAX}
                  min={PPT_ELEMENT_SHADOW_OPACITY_MIN}
                  step={PPT_ELEMENT_SHADOW_OPACITY_STEP}
                  type="number"
                  value={elementShadow.opacity}
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
                  data-ppt-shadow-field="blur"
                  disabled={!elementShadowEnabled}
                  max={PPT_ELEMENT_SHADOW_BLUR_MAX}
                  min={0}
                  step={1}
                  type="number"
                  value={elementShadow.blur}
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
                  data-ppt-shadow-field="distance"
                  disabled={!elementShadowEnabled}
                  max={PPT_ELEMENT_SHADOW_DISTANCE_MAX}
                  min={0}
                  step={1}
                  type="number"
                  value={elementShadow.distance}
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
                  data-ppt-shadow-field="angle"
                  disabled={!elementShadowEnabled}
                  max={PPT_ELEMENT_SHADOW_ANGLE_MAX}
                  min={PPT_ELEMENT_SHADOW_ANGLE_MIN}
                  step={1}
                  type="number"
                  value={elementShadow.angle}
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
                data-ppt-animation-delay={selectedElementAnimation.delayMs}
                data-ppt-animation-duration={selectedElementAnimation.durationMs}
                data-ppt-animation-order={selectedElementAnimation.order}
                data-ppt-animation-trigger={selectedElementAnimation.trigger}
                data-ppt-animation-type={selectedElementAnimation.type}
              >
                <label className="ppt-field">
                  <span>Animation</span>
                  <select
                    data-ppt-animation-field="type"
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
                    data-ppt-animation-field="trigger"
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
                    data-ppt-animation-field="durationMs"
                    max={PPT_ELEMENT_ANIMATION_TIME_MAX}
                    min={0}
                    step={100}
                    type="number"
                    value={selectedElementAnimation.durationMs}
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
                    data-ppt-animation-field="delayMs"
                    max={PPT_ELEMENT_ANIMATION_TIME_MAX}
                    min={0}
                    step={100}
                    type="number"
                    value={selectedElementAnimation.delayMs}
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
                    data-ppt-animation-field="order"
                    min={1}
                    step={1}
                    type="number"
                    value={selectedElementAnimation.order}
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
                      recentColors={recentColors}
                      themeColorTokens={themeColorTokens}
                      onSelect={(color) =>
                        onElementTextStyleChange(selectedElement.id, 'color', color)}
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
                    value={normalizePPTTextFontFamily(textStyle?.fontFamily)}
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
                    value={getPPTTextElementVerticalAlign(selectedElement)}
                    onChange={(event) =>
                      onElementTextStyleChange(
                        selectedElement.id,
                        'verticalAlign',
                        event.target.value,
                      )}
                  >
                    {PPT_TEXT_VERTICAL_ALIGN_OPTIONS.map((option) => (
                      <option key={option.value} value={option.value}>
                        {option.label}
                      </option>
                    ))}
                  </select>
                </label>
                <div
                  className="ppt-paragraph-spacing-grid"
                  data-ppt-text-inset-bottom={textInset.bottom}
                  data-ppt-text-inset-inspector
                  data-ppt-text-inset-left={textInset.left}
                  data-ppt-text-inset-right={textInset.right}
                  data-ppt-text-inset-top={textInset.top}
                >
                  {(['top', 'right', 'bottom', 'left'] as const).map((field) => (
                    <label className="ppt-field" key={field}>
                      <span>{field[0].toUpperCase() + field.slice(1)}</span>
                      <input
                        data-ppt-text-inset-field={field}
                        max={PPT_TEXT_INSET_MAX}
                        min={PPT_TEXT_INSET_MIN}
                        step={PPT_TEXT_INSET_STEP}
                        type="number"
                        value={textInset[field]}
                        onChange={(event) =>
                          onElementTextInsetChange(
                            selectedElement.id,
                            field,
                            parsePPTTextInset(event.target.value),
                          )}
                      />
                    </label>
                  ))}
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
                  data-ppt-paragraph-line-height={paragraphSpacing.lineHeight}
                  data-ppt-paragraph-spacing-after={paragraphSpacing.spacingAfter}
                  data-ppt-paragraph-spacing-before={paragraphSpacing.spacingBefore}
                >
                  <label className="ppt-field">
                    <span>Line height</span>
                    <input
                      data-ppt-paragraph-field="lineHeight"
                      max={PPT_PARAGRAPH_LINE_HEIGHT_MAX}
                      min={PPT_PARAGRAPH_LINE_HEIGHT_MIN}
                      step={0.05}
                      type="number"
                      value={paragraphSpacing.lineHeight}
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
                      data-ppt-paragraph-field="spacingBefore"
                      max={PPT_PARAGRAPH_SPACING_MAX}
                      min={0}
                      step={2}
                      type="number"
                      value={paragraphSpacing.spacingBefore}
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
                      data-ppt-paragraph-field="spacingAfter"
                      max={PPT_PARAGRAPH_SPACING_MAX}
                      min={0}
                      step={2}
                      type="number"
                      value={paragraphSpacing.spacingAfter}
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
                  data-ppt-text-overflow={selectedTextOverflow ? 'true' : 'false'}
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
                      max={PPT_SHAPE_CORNER_RADIUS_MAX}
                      min={PPT_SHAPE_CORNER_RADIUS_MIN}
                      step={PPT_SHAPE_CORNER_RADIUS_STEP}
                      type="number"
                      value={getPPTShapeCornerRadius(selectedElement)}
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
                      recentColors={recentColors}
                      themeColorTokens={themeColorTokens}
                      onSelect={(color) =>
                        onShapeFillChange(selectedElement.id, 'color', color)}
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
                      recentColors={recentColors}
                      themeColorTokens={themeColorTokens}
                      onSelect={(color) =>
                        onElementStrokeChange(selectedElement.id, 'color', color)}
                    />
                  </div>
                </div>
                <label className="ppt-field">
                  <span>Fill opacity</span>
                  <input
                    data-ppt-style-field="fill-opacity"
                    max={PPT_FILL_OPACITY_MAX}
                    min={PPT_FILL_OPACITY_MIN}
                    step={PPT_FILL_OPACITY_STEP}
                    type="number"
                    value={getPPTFillOpacity(selectedElement.fill)}
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
                    value={getPPTStrokeDash(selectedElement.stroke)}
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
                    {PPT_STROKE_DASH_OPTIONS.map((option) => (
                      <option key={option.value} value={option.value}>
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
                    data-ppt-style-field="image-fit"
                    value={getPPTImageFit(selectedElement)}
                    onChange={(event) => {
                      if (isPPTImageFit(event.target.value)) {
                        onImageFitChange(selectedElement.id, event.target.value)
                      }
                    }}
                  >
                    <option value="cover">Cover</option>
                    <option value="contain">Contain</option>
                  </select>
                </label>
                <div className="ppt-geometry-grid">
                  {(['x', 'y'] as const).map((field) => (
                    <label className="ppt-field" key={field}>
                      <span>Crop {field.toUpperCase()}</span>
                      <input
                        data-ppt-style-field={`image-crop-${field}`}
                        max={100}
                        min={0}
                        type="number"
                        value={Math.round(getPPTImageCrop(selectedElement)[field])}
                        onChange={(event) =>
                          onImageCropChange(
                            selectedElement.id,
                            field,
                            Number(event.target.value),
                          )}
                      />
                    </label>
                  ))}
                </div>
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
                      recentColors={recentColors}
                      themeColorTokens={themeColorTokens}
                      onSelect={(color) =>
                        onElementStrokeChange(selectedElement.id, 'color', color)}
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
                    value={getPPTStrokeDash(selectedElement.stroke)}
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
                    {PPT_STROKE_DASH_OPTIONS.map((option) => (
                      <option key={option.value} value={option.value}>
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
        data-ppt-layer-pane-active-object-id={layerPaneDescriptor.activeObjectId ?? ''}
        data-ppt-layer-pane-command-count={PPT_LAYER_PANE_COMMANDS.length}
        data-ppt-layer-pane-commands={layerPaneCommandIds}
        data-ppt-layer-pane-command-slot="command-effect"
        data-ppt-layer-pane-row-count={layerPaneDescriptor.rows.length}
        data-ppt-layer-pane-slide-id={layerPaneDescriptor.slideId}
      >
        <div
          aria-label="Selection pane"
          className="ppt-layer-list"
          data-ppt-layer-pane-aria-container={layerPaneDescriptor.aria.containerRole}
          data-ppt-layer-pane-aria-row={layerPaneDescriptor.aria.rowRole}
          data-ppt-layer-pane-keyboard-keys="arrow-home-end-enter-space"
          data-ppt-layer-pane-keyboard-model={layerPaneDescriptor.aria.keyboardModel}
          data-ppt-layer-pane-selection-model={layerPaneDescriptor.aria.selectionModel}
          role={layerPaneDescriptor.aria.containerRole}
        >
          {layerPaneDescriptor.rows.map((row) => (
            <div
              aria-disabled={!row.isSelectable}
              aria-level={row.ariaLevel}
              aria-posinset={row.ariaPosInSet}
              aria-selected={row.isSelected}
              aria-setsize={row.ariaSetSize}
              className="ppt-layer-row"
              data-grouped={row.isGrouped ? 'true' : 'false'}
              data-hidden={row.isHidden ? 'true' : 'false'}
              data-locked={row.isLocked ? 'true' : 'false'}
              data-ppt-layer-pane-grouped={row.isGrouped ? 'true' : 'false'}
              data-ppt-layer-pane-hidden={row.isHidden ? 'true' : 'false'}
              data-ppt-layer-pane-kind={row.kindLabel}
              data-ppt-layer-pane-locked={row.isLocked ? 'true' : 'false'}
              data-ppt-layer-pane-order={row.order}
              data-ppt-layer-pane-renamable={row.isRenamable ? 'true' : 'false'}
              data-ppt-layer-pane-reorderable={row.isReorderable ? 'true' : 'false'}
              data-ppt-layer-pane-row={row.objectId}
              data-ppt-layer-pane-selected={row.isSelected ? 'true' : 'false'}
              data-ppt-layer-row={row.objectId}
              key={row.objectId}
              role={layerPaneDescriptor.aria.rowRole}
              tabIndex={row.isSelected ? 0 : -1}
              onKeyDown={(event) => handleLayerPaneRowKeyDown(row, event)}
            >
              <button
                className="ppt-layer-select"
                data-ppt-layer-pane-intent="row-press"
                data-ppt-layer-select={row.objectId}
                type="button"
                onClick={(event) =>
                  runLayerPaneIntent({
                    additive: event.metaKey || event.ctrlKey || event.shiftKey,
                    objectId: row.objectId,
                    type: 'row-press',
                  })}
              >
                <span className="ppt-layer-kind">
                  <Layers size={14} />
                </span>
                <span className="ppt-layer-name">{row.displayName}</span>
              </button>
              <div className="ppt-layer-actions">
                <button
                  aria-label={row.isHidden ? 'Show object' : 'Hide object'}
                  className="ppt-layer-icon-button"
                  data-ppt-layer-pane-intent="visibility-toggle"
                  data-ppt-layer-visibility={row.objectId}
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
          ))}
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
      ? `${getPPTShapeCornerRadius(element)}px`
      : undefined,
    padding: getPPTTextInsetCSS(getPPTTextElementInset(element)),
    textAlign: getPPTElementParagraphAlign(element),
  }
}

function getPPTElementTransform(element: PPTElement) {
  const transforms = [
    element.geometry.rotation ? `rotate(${element.geometry.rotation}deg)` : '',
    element.flipH === true ? 'scaleX(-1)' : '',
    element.flipV === true ? 'scaleY(-1)' : '',
  ].filter(Boolean)

  return transforms.length > 0 ? transforms.join(' ') : undefined
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

function isEditableTarget(target: EventTarget | null) {
  return target instanceof HTMLElement &&
    (target.isContentEditable ||
      ['INPUT', 'TEXTAREA', 'SELECT'].includes(target.tagName))
}

function isPPTTemporaryPanKey(event: KeyboardEvent) {
  return event.code === 'Space' || event.key === ' ' || event.key === 'Spacebar'
}

function isPPTTemporaryPanShortcut(event: KeyboardEvent) {
  return isPPTTemporaryPanKey(event) &&
    !event.altKey &&
    !event.ctrlKey &&
    !event.metaKey
}

function isPPTTemporaryPanBlockedTarget(target: EventTarget | null) {
  return target instanceof Element &&
    Boolean(target.closest([
      'button',
      'input',
      'select',
      'textarea',
      '[contenteditable="true"]',
      '[data-ppt-command-palette]',
      '[data-ppt-context-menu]',
      '[data-ppt-shortcut-help]',
      '[role="button"]',
      '[role="menuitem"]',
      '[role="option"]',
      '[role="tab"]',
    ].join(',')))
}

function isPPTWheelViewportPassthroughTarget(target: EventTarget | null) {
  return target instanceof Element &&
    Boolean(target.closest([
      'button',
      'input',
      'select',
      'textarea',
      '[contenteditable="true"]',
      '[data-ppt-command-palette]',
      '[data-ppt-context-menu]',
      '[data-ppt-minimap]',
      '[data-ppt-shortcut-help]',
      '[data-ppt-wheel-passthrough="true"]',
      '[role="button"]',
      '[role="menuitem"]',
      '[role="option"]',
      '[role="tab"]',
    ].join(',')))
}

function isPPTShortcutHelpShortcut(event: KeyboardEvent) {
  return event.shiftKey &&
    !event.metaKey &&
    !event.ctrlKey &&
    !event.altKey &&
    (event.key === '?' || event.key === '/' || event.code === 'Slash')
}

function selectSameTypePPTSelection(
  elements: readonly PPTElement[],
  selection: readonly string[],
): string[] {
  const selected = new Set(selection)
  const selectedTypes = new Set(
    elements
      .filter((element) =>
        selected.has(element.id) && element.visible !== false)
      .map(getPPTElementTypeKey),
  )

  if (selectedTypes.size === 0) {
    return [...selection]
  }

  return elements
    .filter((element) =>
      element.visible !== false && selectedTypes.has(getPPTElementTypeKey(element)))
    .map((element) => element.id)
}

function canSelectSameTypePPTSelection(
  elements: readonly PPTElement[],
  selection: readonly string[],
) {
  if (selection.length === 0) {
    return false
  }

  const selected = new Set(selection)
  const nextSelection = selectSameTypePPTSelection(elements, selection)

  return nextSelection.some((id) => !selected.has(id))
}

function getPPTElementTypeKey(element: PPTElement) {
  if (element.kind === 'shape') {
    return `shape:${element.shape}`
  }

  return 'Object'
}

function canFlipPPTSelection(
  elements: readonly PPTElement[],
  selection: readonly string[],
) {
  if (selection.length === 0) {
    return false
  }

  return getPPTFlipSelectionElements(elements, selection).length === selection.length
}

function flipPPTSelectionElements(
  elements: PPTElement[],
  selection: readonly string[],
  axis: PPTFlipAxis,
) {
  if (!canFlipPPTSelection(elements, selection)) {
    return elements
  }

  const selected = new Set(selection)
  const selectedElements = getPPTFlipSelectionElements(elements, selection)
  const selectionBounds = getPPTElementsBounds(selectedElements)

  if (!selectionBounds) {
    return elements
  }

  const pivot = axis === 'horizontal'
    ? selectionBounds.x + selectionBounds.w / 2
    : selectionBounds.y + selectionBounds.h / 2

  return elements.map((element) =>
    selected.has(element.id) ? flipPPTElement(element, axis, pivot) : element)
}

function flipPPTElement(
  element: PPTElement,
  axis: PPTFlipAxis,
  pivot: number,
): PPTElement {
  if (element.kind === 'line') {
    return flipPPTLineElement(element, axis, pivot)
  }

  const bounds = pptGeometryToBounds(element.geometry)
  const next = updatePPTElementBounds(element, axis === 'horizontal'
    ? { ...bounds, x: 2 * pivot - (bounds.x + bounds.w) }
    : { ...bounds, y: 2 * pivot - (bounds.y + bounds.h) })

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

function getPPTFlipSelectionElements(
  elements: readonly PPTElement[],
  selection: readonly string[],
) {
  const selected = new Set(selection)

  return elements.filter((element) =>
    selected.has(element.id) &&
    element.visible !== false &&
    element.locked !== true)
}

function canTidyPPTSelection(
  elements: readonly PPTElement[],
  selection: readonly string[],
) {
  if (selection.length < 3) {
    return false
  }

  return getPPTTidySelectionElements(elements, selection).length === selection.length
}

function tidyPPTSelectionElements(
  elements: PPTElement[],
  selection: readonly string[],
) {
  if (!canTidyPPTSelection(elements, selection)) {
    return elements
  }

  const selected = getPPTTidySelectionElements(elements, selection)
  const selectionBounds = getPPTElementsBounds(selected)

  if (!selectionBounds) {
    return elements
  }

  const columnCount = Math.ceil(Math.sqrt(selected.length))
  const boundsList = selected.map((element) => pptGeometryToBounds(element.geometry))
  const cellWidth = Math.max(...boundsList.map((bounds) => bounds.w)) + PPT_TIDY_GAP
  const cellHeight = Math.max(...boundsList.map((bounds) => bounds.h)) + PPT_TIDY_GAP
  const sorted = [...selected].sort((a, b) => {
    const aBounds = pptGeometryToBounds(a.geometry)
    const bBounds = pptGeometryToBounds(b.geometry)

    return aBounds.y === bBounds.y
      ? aBounds.x - bBounds.x
      : aBounds.y - bBounds.y
  })
  const tidied = new Map<string, PPTElement>()

  sorted.forEach((element, index) => {
    const bounds = pptGeometryToBounds(element.geometry)
    const column = index % columnCount
    const row = Math.floor(index / columnCount)

    tidied.set(element.id, updatePPTElementBounds(element, {
      ...bounds,
      x: selectionBounds.x + column * cellWidth,
      y: selectionBounds.y + row * cellHeight,
    }))
  })

  return elements.map((element) => tidied.get(element.id) ?? element)
}

function getPPTTidySelectionElements(
  elements: readonly PPTElement[],
  selection: readonly string[],
) {
  const selected = new Set(selection)

  return elements.filter((element) =>
    selected.has(element.id) &&
    element.visible !== false &&
    element.locked !== true &&
    element.kind !== 'line')
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

function canApplyPPTStyleClipboard(
  element: PPTElement,
  clipboard: PPTStyleClipboard,
) {
  return getPPTStyleClipboardTargetCategories(element, clipboard).length > 0
}

function getPPTStyleClipboardTargetCategories(
  element: PPTElement,
  clipboard: PPTStyleClipboard,
) {
  if (element.locked === true || element.visible === false) {
    return []
  }

  const categories: PPTStyleClipboardCategory[] = ['object']

  if (element.kind === 'shape' && clipboard.shape) {
    categories.push('shape')
  }

  if (
    (element.kind === 'shape' ||
      element.kind === 'line' ||
      element.kind === 'freeform') &&
    clipboard.stroke
  ) {
    categories.push('stroke')
  }

  if (isPPTTextElement(element) && clipboard.text) {
    categories.push('text')
  }

  if (isPPTTextElement(element) && clipboard.paragraph) {
    categories.push('paragraph')
  }

  return categories
}

function applyPPTStyleClipboardToElement(
  element: PPTElement,
  clipboard: PPTStyleClipboard,
): PPTElement {
  if (!canApplyPPTStyleClipboard(element, clipboard)) {
    return element
  }

  let next: PPTElement = {
    ...element,
    opacity: clipboard.object.opacity === 1
      ? undefined
      : clipboard.object.opacity,
    shadow: clipboard.object.shadow
      ? clonePPTElementShadow(clipboard.object.shadow)
      : undefined,
  }

  if (next.kind === 'shape') {
    if (clipboard.shape) {
      const cornerRadius = clipboard.shape.cornerRadius ?? PPT_SHAPE_CORNER_RADIUS_DEFAULT
      next = {
        ...next,
        cornerRadius: next.shape === 'rect'
          ? getPPTShapeCornerRadiusModelValue(cornerRadius)
          : undefined,
        fill: clonePPTFill(clipboard.shape.fill),
        stroke: clipboard.shape.stroke
          ? clonePPTStroke(clipboard.shape.stroke)
          : undefined,
      }
    } else if (clipboard.stroke) {
      next = {
        ...next,
        stroke: clonePPTStroke(clipboard.stroke),
      }
    }
  } else if (
    (next.kind === 'line' || next.kind === 'freeform') &&
    clipboard.stroke
  ) {
    next = {
      ...next,
      stroke: clonePPTStroke(clipboard.stroke),
    }
  }

  if (isPPTTextElement(next)) {
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
  }
}

function areAllPPTTextElementsBulleted(elements: readonly PPTTextElement[]) {
  return elements.length > 0 &&
    elements.every((element) =>
      element.textBody.paragraphs.length > 0 &&
      element.textBody.paragraphs.every((paragraph) => paragraph.bullet === 'bullet'))
}

function hasPPTTextBodyBullet(body: PPTTextBody) {
  return body.paragraphs.some((paragraph) => paragraph.bullet === 'bullet')
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

function getPPTImageCrop(element: PPTImage): PPTImageCrop {
  return element.crop ?? { x: 50, y: 50 }
}

function getPPTImageObjectPosition(element: PPTImage) {
  const crop = getPPTImageCrop(element)

  return `${crop.x}% ${crop.y}%`
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

function normalizePPTTextFontFamily(fontFamily: string | undefined) {
  return PPT_TEXT_FONT_FAMILY_VALUES.has(fontFamily ?? '')
    ? fontFamily ?? PPT_DEFAULT_TEXT_FONT_FAMILY
    : PPT_DEFAULT_TEXT_FONT_FAMILY
}

function getPPTTextFontFamilyCSS(fontFamily: string | undefined) {
  const normalized = normalizePPTTextFontFamily(fontFamily)

  return PPT_TEXT_FONT_FAMILY_OPTIONS.find((option) => option.value === normalized)?.css ??
    PPT_TEXT_FONT_FAMILY_OPTIONS[0].css
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

function normalizePPTTextVerticalAlign(
  verticalAlign: string | undefined,
  fallback: PPTTextVerticalAlign = PPT_DEFAULT_TEXT_VERTICAL_ALIGN,
) {
  return PPT_TEXT_VERTICAL_ALIGN_VALUES.has(verticalAlign ?? '')
    ? verticalAlign as PPTTextVerticalAlign
    : fallback
}

function getPPTTextVerticalAlignCSS(verticalAlign: string | undefined) {
  const normalized = normalizePPTTextVerticalAlign(verticalAlign)

  return PPT_TEXT_VERTICAL_ALIGN_OPTIONS.find((option) => option.value === normalized)?.css ??
    PPT_TEXT_VERTICAL_ALIGN_OPTIONS[0].css
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
  return clamp(Math.round(Number.isFinite(value) ? value : 0), PPT_TEXT_INSET_MIN, PPT_TEXT_INSET_MAX)
}

function getPPTTextInsetCSS(inset: PPTTextInset) {
  return `${inset.top}px ${inset.right}px ${inset.bottom}px ${inset.left}px`
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
  const finiteValue = Number.isFinite(value) ? value : 1
  const clamped = clamp(finiteValue, PPT_ELEMENT_OPACITY_MIN, PPT_ELEMENT_OPACITY_MAX)

  return Math.round(clamped * 100) / 100
}

function formatPPTElementOpacity(value: number) {
  return String(normalizePPTElementOpacity(value))
}

function getPPTElementHyperlink(element: PPTElement) {
  return element.hyperlink ? normalizePPTElementHyperlink(element.hyperlink) : null
}

function getPPTElementAltText(element: PPTElement) {
  return element.accessibility
    ? normalizePPTElementAccessibility(element.accessibility)?.altText
    : undefined
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

function getPPTThumbShapeCornerRadiusCSS(element: PPTShape) {
  const radius = getPPTShapeCornerRadius(element)
  const minSize = Math.max(1, Math.min(element.geometry.w, element.geometry.h))
  const percent = Math.min(50, (radius / minSize) * 100)

  return `${Math.round(percent * 100) / 100}%`
}

function parsePPTShapeCornerRadius(value: string) {
  return normalizePPTShapeCornerRadius(Number(value))
}

function normalizePPTShapeCornerRadius(value: number) {
  const finiteValue = Number.isFinite(value)
    ? value
    : PPT_SHAPE_CORNER_RADIUS_DEFAULT
  const clamped = clamp(
    finiteValue,
    PPT_SHAPE_CORNER_RADIUS_MIN,
    PPT_SHAPE_CORNER_RADIUS_MAX,
  )

  return Math.round(clamped)
}

function formatPPTShapeCornerRadius(value: number) {
  return String(normalizePPTShapeCornerRadius(value))
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
  const finiteValue = Number.isFinite(value) ? value : 1
  const clamped = clamp(finiteValue, PPT_FILL_OPACITY_MIN, PPT_FILL_OPACITY_MAX)

  return Math.round(clamped * 100) / 100
}

function formatPPTFillOpacity(value: number) {
  return String(normalizePPTFillOpacity(value))
}

function getPPTFillColorCSS(fill: PPTFill) {
  const opacity = getPPTFillOpacity(fill)

  if (opacity === 1) {
    return fill.color
  }

  return getPPTColorWithAlpha(fill.color, opacity)
}

function getPPTColorWithAlpha(color: string, opacity: number) {
  const hex = normalizePPTColorHex(color)

  if (!hex) {
    return color
  }

  const red = Number.parseInt(hex.slice(0, 2), 16)
  const green = Number.parseInt(hex.slice(2, 4), 16)
  const blue = Number.parseInt(hex.slice(4, 6), 16)

  return `rgb(${red} ${green} ${blue} / ${formatPPTFillOpacity(opacity)})`
}

function normalizePPTColorHex(color: string) {
  const match = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(color.trim())

  if (!match) {
    return ''
  }

  const value = match[1]

  return value.length === 3
    ? [...value].map((char) => `${char}${char}`).join('').toLowerCase()
    : value.toLowerCase()
}

function normalizePPTSwatchColor(color: string) {
  const hex = normalizePPTColorHex(color)

  return hex ? `#${hex}` : ''
}

function getPPTUniqueSwatchColors(colors: readonly string[]) {
  const seen = new Set<string>()
  const unique: string[] = []

  for (const color of colors) {
    const normalized = normalizePPTSwatchColor(color)

    if (!normalized || seen.has(normalized)) {
      continue
    }

    seen.add(normalized)
    unique.push(normalized)
  }

  return unique
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
  return typeof value === 'string' && isPPTStrokeDash(value)
    ? value as PPTStrokeDash
    : 'solid'
}

function getPPTStrokeDashBorderStyle(stroke: PPTStroke | undefined) {
  const dash = getPPTStrokeDash(stroke)

  if (dash === 'dash') {
    return 'dashed'
  }

  if (dash === 'dot') {
    return 'dotted'
  }

  return 'solid'
}

function getPPTStrokeDashArray(stroke: PPTStroke | undefined) {
  const width = normalizePPTStrokeWidth(stroke?.width ?? 2)
  const dash = getPPTStrokeDash(stroke)

  if (dash === 'dash') {
    return `${Math.max(4, width * 3)} ${Math.max(3, width * 2)}`
  }

  if (dash === 'dot') {
    return `1 ${Math.max(3, width * 2)}`
  }

  return undefined
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
  const normalized = value.trim().slice(0, PPT_ALT_TEXT_MAX_LENGTH)

  if (!normalized || hasPPTControlCharacter(normalized)) {
    return ''
  }

  return normalized
}

function normalizePPTElementHyperlink(
  hyperlink: Partial<PPTElementHyperlink>,
): PPTElementHyperlink | null {
  const url = normalizePPTElementHyperlinkUrl(hyperlink.url ?? '')

  return url ? { url } : null
}

function normalizePPTElementHyperlinkUrl(url: string) {
  const normalized = url.trim().slice(0, PPT_HYPERLINK_URL_MAX_LENGTH)

  if (!normalized || !isPPTElementHyperlinkUrlAllowed(normalized)) {
    return ''
  }

  return normalized
}

function isPPTElementHyperlinkUrlAllowed(url: string) {
  return !hasPPTControlCharacter(url) &&
    !/^(javascript|data|vbscript):/i.test(url)
}

function hasPPTControlCharacter(value: string) {
  return [...value].some((char) => {
    const code = char.charCodeAt(0)

    return code <= 31 || code === 127
  })
}

function hasPPTElementShadow(element: PPTElement) {
  return element.shadow !== undefined
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

  return clamp(
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

  return clamp(Math.round(finiteValue), 0, PPT_ELEMENT_SHADOW_BLUR_MAX)
}

function parsePPTElementShadowDistance(value: string) {
  return normalizePPTElementShadowDistance(Number(value))
}

function normalizePPTElementShadowDistance(value: number) {
  const finiteValue = Number.isFinite(value) ? value : PPT_DEFAULT_ELEMENT_SHADOW.distance

  return clamp(Math.round(finiteValue), 0, PPT_ELEMENT_SHADOW_DISTANCE_MAX)
}

function parsePPTElementShadowOpacity(value: string) {
  return normalizePPTElementShadowOpacity(Number(value))
}

function normalizePPTElementShadowOpacity(value: number) {
  const finiteValue = Number.isFinite(value) ? value : PPT_DEFAULT_ELEMENT_SHADOW.opacity
  const clamped = clamp(
    finiteValue,
    PPT_ELEMENT_SHADOW_OPACITY_MIN,
    PPT_ELEMENT_SHADOW_OPACITY_MAX,
  )

  return Math.round(clamped * 100) / 100
}

function formatPPTElementShadowOpacity(value: number) {
  return String(normalizePPTElementShadowOpacity(value))
}

function normalizePPTElementShadowColor(color: string) {
  return /^#[\da-f]{6}$/i.test(color) ? color : PPT_DEFAULT_ELEMENT_SHADOW.color
}

function getPPTElementShadowFilter(element: PPTElement) {
  return hasPPTElementShadow(element)
    ? `drop-shadow(${getPPTElementShadowFilterCSS(getPPTElementShadow(element))})`
    : undefined
}

function getPPTElementShadowFilterCSS(shadow: PPTElementShadow) {
  const radians = (shadow.angle * Math.PI) / 180
  const offsetX = formatPPTPathNumber(Math.cos(radians) * shadow.distance)
  const offsetY = formatPPTPathNumber(Math.sin(radians) * shadow.distance)

  return `${offsetX}px ${offsetY}px ${shadow.blur}px ${getPPTElementShadowColorCSS(shadow)}`
}

function getPPTElementShadowColorCSS(shadow: PPTElementShadow) {
  const match = /^#([\da-f]{2})([\da-f]{2})([\da-f]{2})$/i.exec(shadow.color)

  if (!match) {
    return shadow.color
  }

  const [, red, green, blue] = match

  return `rgb(${Number.parseInt(red, 16)} ${Number.parseInt(green, 16)} ${Number.parseInt(blue, 16)} / ${formatPPTElementShadowOpacity(shadow.opacity)})`
}

function getPPTSelectionCommandAnchor({
  barWidth,
  bounds,
  stage,
  viewport,
}: {
  barWidth: number
  bounds: Bounds
  stage: HTMLElement | null
  viewport: Viewport
}): PPTSelectionCommandAnchor {
  const scale = viewport.scale
  const gap = 10 / scale
  const screenMargin = 8
  const barHeight = 40
  const barHalfWidth = barWidth / 2 / scale
  const centerX = bounds.x + bounds.w / 2

  if (!stage) {
    return {
      placement: 'above',
      x: clamp(centerX, barHalfWidth, PPT_SLIDE_WIDTH - barHalfWidth),
      y: bounds.y - gap,
    }
  }

  const rect = stage.getBoundingClientRect()
  const visibleLeft = clamp((screenMargin - viewport.x) / scale, 0, PPT_SLIDE_WIDTH)
  const visibleRight = clamp(
    (rect.width - screenMargin - viewport.x) / scale,
    0,
    PPT_SLIDE_WIDTH,
  )
  const minX = Math.min(visibleLeft + barHalfWidth, PPT_SLIDE_WIDTH - barHalfWidth)
  const maxX = Math.max(minX, visibleRight - barHalfWidth)
  const aboveY = bounds.y - gap
  const belowY = bounds.y + bounds.h + gap
  const aboveScreenY = viewport.y + aboveY * scale
  const belowScreenY = viewport.y + belowY * scale
  const aboveFits = aboveScreenY - barHeight >= screenMargin
  const belowFits = belowScreenY + barHeight <= rect.height - screenMargin
  const placement = aboveFits || !belowFits ? 'above' : 'below'

  return {
    placement,
    x: clamp(centerX, minX, maxX),
    y: placement === 'above' ? aboveY : belowY,
  }
}

function getPPTElementParagraphAlign(element: PPTElement) {
  if (!isPPTTextElement(element)) {
    return undefined
  }

  return element.textBody?.paragraphs[0]?.align ?? 'left'
}

const PPT_CANVAS_CREATION_ADAPTER: CanvasCreationAdapter<PPTElement> = {
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
    const created = createCanvasText({
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

  return createCanvasShape({
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
  const bounds = normalizeBounds(startWorld, currentWorld)

  if (bounds.w > 6 && bounds.h > 6) {
    return clampPPTCreationBounds(bounds)
  }

  return clampPPTCreationBounds({
    ...PPT_DEFAULT_TEXT_BOUNDS,
    x: startWorld.x,
    y: startWorld.y,
  })
}

function clampPPTCreationBounds(bounds: Bounds): Bounds {
  const w = clamp(bounds.w, 24, PPT_SLIDE_WIDTH)
  const h = clamp(bounds.h, 24, PPT_SLIDE_HEIGHT)

  return {
    h,
    w,
    x: clamp(bounds.x, 0, PPT_SLIDE_WIDTH - w),
    y: clamp(bounds.y, 0, PPT_SLIDE_HEIGHT - h),
  }
}

function toPPTShapeKind(shape: CanvasCreatedShapeKind): PPTShapeKind {
  if (shape === 'ellipse' || shape === 'diamond') {
    return shape
  }

  return 'rect'
}

function getPPTCreationToolForShortcut(event: KeyboardEvent): PPTCreationTool | null {
  if (event.metaKey || event.ctrlKey || event.altKey) {
    return null
  }

  if (doesEventMatchCanvasToolShortcut(event, CANVAS_TOOL_AFFORDANCES.text.keyboardShortcut)) {
    return { kind: 'text' }
  }

  if (doesEventMatchCanvasToolShortcut(event, CANVAS_TOOL_AFFORDANCES.rect.keyboardShortcut)) {
    return { kind: 'shape', shape: 'rect' }
  }

  if (doesEventMatchCanvasToolShortcut(event, CANVAS_TOOL_AFFORDANCES.ellipse.keyboardShortcut)) {
    return { kind: 'shape', shape: 'ellipse' }
  }

  if (doesEventMatchCanvasToolShortcut(event, CANVAS_TOOL_AFFORDANCES.comment.keyboardShortcut)) {
    return { kind: 'comment' }
  }

  if (doesEventMatchCanvasToolShortcut(event, CANVAS_TOOL_AFFORDANCES.marker.keyboardShortcut)) {
    return { kind: 'freeform', tool: 'marker' }
  }

  if (doesEventMatchCanvasToolShortcut(event, CANVAS_TOOL_AFFORDANCES.highlight.keyboardShortcut)) {
    return { kind: 'freeform', tool: 'highlight' }
  }

  if (doesEventMatchCanvasToolShortcut(event, CANVAS_TOOL_AFFORDANCES.pen.keyboardShortcut)) {
    return { kind: 'freeform', tool: 'pen' }
  }

  return null
}

function isPPTSelectToolShortcut(event: KeyboardEvent) {
  return !event.metaKey &&
    !event.ctrlKey &&
    !event.altKey &&
    doesEventMatchCanvasToolShortcut(event, CANVAS_TOOL_AFFORDANCES.select.keyboardShortcut)
}

function isPPTPanToolShortcut(event: KeyboardEvent) {
  return !event.metaKey &&
    !event.ctrlKey &&
    !event.altKey &&
    doesEventMatchCanvasToolShortcut(event, CANVAS_TOOL_AFFORDANCES.pan.keyboardShortcut)
}

function isPPTLaserToolShortcut(event: KeyboardEvent) {
  return !event.metaKey &&
    !event.ctrlKey &&
    !event.altKey &&
    doesEventMatchCanvasToolShortcut(event, CANVAS_TOOL_AFFORDANCES.laser.keyboardShortcut)
}

function isPPTEraserToolShortcut(event: KeyboardEvent) {
  return !event.metaKey &&
    !event.ctrlKey &&
    !event.altKey &&
    doesEventMatchCanvasToolShortcut(event, CANVAS_TOOL_AFFORDANCES.eraser.keyboardShortcut)
}

function doesEventMatchCanvasToolShortcut(
  event: KeyboardEvent,
  shortcut: {
    key: string
    shiftInsensitive?: boolean
    shiftKey?: boolean
  } | undefined,
) {
  if (!shortcut || event.key.toLowerCase() !== shortcut.key.toLowerCase()) {
    return false
  }

  return shortcut.shiftInsensitive || event.shiftKey === Boolean(shortcut.shiftKey)
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
  const safeStart = clamp(Math.min(start, end), 0, tokens.length)
  const safeEnd = clamp(Math.max(start, end), safeStart, tokens.length)
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

function isArrowKey(key: string) {
  return key === 'ArrowDown' ||
    key === 'ArrowLeft' ||
    key === 'ArrowRight' ||
    key === 'ArrowUp'
}

function getArrowNudgeDelta(key: string, distance: number) {
  if (key === 'ArrowLeft') {
    return { dx: -distance, dy: 0 }
  }

  if (key === 'ArrowRight') {
    return { dx: distance, dy: 0 }
  }

  if (key === 'ArrowUp') {
    return { dx: 0, dy: -distance }
  }

  return { dx: 0, dy: distance }
}

function getBoundsCenter(bounds: Bounds): Point {
  return {
    x: bounds.x + bounds.w / 2,
    y: bounds.y + bounds.h / 2,
  }
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
  return clamp(line.routeBend ?? 0.5, 0.08, 0.92)
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
  const safePoints = points.length > 0
    ? points.map(clampPPTPointToSlide)
    : [{ x: PPT_SLIDE_WIDTH / 2, y: PPT_SLIDE_HEIGHT / 2 }]
  const left = Math.min(...safePoints.map((point) => point.x))
  const top = Math.min(...safePoints.map((point) => point.y))
  const right = Math.max(...safePoints.map((point) => point.x))
  const bottom = Math.max(...safePoints.map((point) => point.y))
  const padding = 8
  const rawWidth = right - left
  const rawHeight = bottom - top
  const width = Math.max(16, rawWidth + padding * 2)
  const height = Math.max(16, rawHeight + padding * 2)
  const x = clamp(left - Math.max(padding, (width - rawWidth) / 2), 0, PPT_SLIDE_WIDTH - width)
  const y = clamp(top - Math.max(padding, (height - rawHeight) / 2), 0, PPT_SLIDE_HEIGHT - height)

  return {
    geometry: {
      h: height,
      w: width,
      x,
      y,
    },
    points: safePoints.map((point) => ({
      x: point.x - x,
      y: point.y - y,
    })),
  }
}

function appendPPTFreeformPoint(points: Point[], point: Point) {
  const next = clampPPTPointToSlide(point)
  const last = points.at(-1)

  if (last && getPointDistance(last, next) < 2) {
    return points
  }

  return [...points, next]
}

function getNextPPTLaserTrailPoints(points: Point[], point: Point) {
  const next = clampPPTPointToSlide(point)
  const last = points.at(-1)

  if (last && getPointDistance(last, next) < PPT_LASER_POINT_DISTANCE) {
    return points
  }

  return [...points, next].slice(-PPT_LASER_TRAIL_MAX_POINTS)
}

function getNextPPTEraserPoints(points: Point[], point: Point) {
  const next = clampPPTPointToSlide(point)
  const last = points.at(-1)

  if (!last) {
    return [next]
  }

  const distance = getPointDistance(last, next)

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

function getPPTEraserHitElementIds(slide: PPTSlide, eraserPoints: Point[]) {
  return slide.elements
    .filter((element): element is PPTFreeform =>
      element.kind === 'freeform' &&
      element.visible !== false &&
      element.locked !== true &&
      isPPTFreeformHitByEraser(element, eraserPoints))
    .map((element) => element.id)
}

function isPPTFreeformHitByEraser(element: PPTFreeform, eraserPoints: Point[]) {
  const freeformPoints = getPPTFreeformWorldPoints(element)
  const hitDistance = element.stroke.width / 2 + PPT_ERASER_HIT_PADDING

  return eraserPoints.some((point) =>
    isPPTPointNearPolyline(point, freeformPoints, hitDistance))
}

function isPPTPointNearPolyline(point: Point, polyline: Point[], distance: number) {
  if (polyline.length === 0) {
    return false
  }

  if (polyline.length === 1) {
    return getPointDistance(point, polyline[0]) <= distance
  }

  return polyline.slice(1).some((current, index) =>
    getPPTPointSegmentDistance(point, polyline[index], current) <= distance)
}

function getPPTPointSegmentDistance(point: Point, start: Point, end: Point) {
  const dx = end.x - start.x
  const dy = end.y - start.y
  const lengthSquared = dx * dx + dy * dy

  if (lengthSquared === 0) {
    return getPointDistance(point, start)
  }

  const ratio = clamp(
    ((point.x - start.x) * dx + (point.y - start.y) * dy) / lengthSquared,
    0,
    1,
  )

  return getPointDistance(point, {
    x: start.x + dx * ratio,
    y: start.y + dy * ratio,
  })
}

function mergePPTElementIds(left: string[], right: string[]) {
  const ids = new Set(left)

  right.forEach((id) => ids.add(id))

  return Array.from(ids)
}

function clampPPTPointToSlide(point: Point) {
  return {
    x: clamp(point.x, 0, PPT_SLIDE_WIDTH),
    y: clamp(point.y, 0, PPT_SLIDE_HEIGHT),
  }
}

function getPPTFreeformWorldLength(element: PPTFreeform) {
  const points = getPPTFreeformWorldPoints(element)

  return points.slice(1).reduce((length, point, index) =>
    length + getPointDistance(points[index], point), 0)
}

function getPPTFreeformWorldPoints(element: PPTFreeform) {
  return element.points.map((point) => ({
    x: element.geometry.x + point.x,
    y: element.geometry.y + point.y,
  }))
}

function getPPTFreeformPathData(points: readonly Point[]) {
  const [first, second, ...rest] = points

  if (!first) {
    return ''
  }

  if (!second) {
    return `M ${formatPPTPathNumber(first.x)} ${formatPPTPathNumber(first.y)}`
  }

  if (rest.length === 0) {
    return [
      `M ${formatPPTPathNumber(first.x)} ${formatPPTPathNumber(first.y)}`,
      `L ${formatPPTPathNumber(second.x)} ${formatPPTPathNumber(second.y)}`,
    ].join(' ')
  }

  return [
    `M ${formatPPTPathNumber(first.x)} ${formatPPTPathNumber(first.y)}`,
    `Q ${formatPPTPathNumber(second.x)} ${formatPPTPathNumber(second.y)} ${getPPTPathMidpoint(second, rest[0])}`,
    ...rest.slice(1).map((point, index) => {
      const control = rest[index]

      return `Q ${formatPPTPathNumber(control.x)} ${formatPPTPathNumber(control.y)} ${getPPTPathMidpoint(control, point)}`
    }),
    `L ${formatPPTPathNumber(rest[rest.length - 1].x)} ${formatPPTPathNumber(rest[rest.length - 1].y)}`,
  ].join(' ')
}

function getPPTLaserTrailPathData(points: readonly Point[]) {
  const [first, ...rest] = points

  if (!first) {
    return ''
  }

  return [
    `M ${formatPPTPathNumber(first.x)} ${formatPPTPathNumber(first.y)}`,
    ...rest.map((point) =>
      `L ${formatPPTPathNumber(point.x)} ${formatPPTPathNumber(point.y)}`),
  ].join(' ')
}

function getPPTPathMidpoint(a: Point, b: Point) {
  return `${formatPPTPathNumber((a.x + b.x) / 2)} ${formatPPTPathNumber((a.y + b.y) / 2)}`
}

function formatPPTPathNumber(value: number) {
  return Number(value.toFixed(2))
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
  const rawPoint = {
    x: clamp(point.x, 0, PPT_SLIDE_WIDTH),
    y: clamp(point.y, 0, PPT_SLIDE_HEIGHT),
  }
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
    : clamp((point.x - start.x) / span, 0.08, 0.92)

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
  const rawLeft = Math.min(start.x, end.x)
  const rawTop = Math.min(start.y, end.y)
  const rawWidth = Math.abs(end.x - start.x)
  const rawHeight = Math.abs(end.y - start.y)
  const width = Math.max(24, rawWidth)
  const height = Math.max(24, rawHeight)
  const x = clamp(
    rawLeft - Math.max(0, width - rawWidth) / 2,
    0,
    PPT_SLIDE_WIDTH - width,
  )
  const y = clamp(
    rawTop - Math.max(0, height - rawHeight) / 2,
    0,
    PPT_SLIDE_HEIGHT - height,
  )

  const next: PPTLine = {
    ...line,
    end: {
      x: end.x - x,
      y: end.y - y,
    },
    geometry: {
      ...line.geometry,
      h: height,
      w: width,
      x,
      y,
    },
    start: {
      x: start.x - x,
      y: start.y - y,
    },
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
      const distance = getPointDistance(point, anchor.point)

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
  const selected = new Set(selection)

  return new Set(
    elements
      .filter((element) => selected.has(element.id) && element.kind === 'line')
      .map((element) => element.id),
  )
}

function getPPTConnectorAnchors(element: PPTElement): Array<{
  anchor: PPTLineConnection['anchor']
  point: Point
}> {
  const left = element.geometry.x
  const top = element.geometry.y
  const right = element.geometry.x + element.geometry.w
  const bottom = element.geometry.y + element.geometry.h
  const centerX = element.geometry.x + element.geometry.w / 2
  const centerY = element.geometry.y + element.geometry.h / 2

  return [
    { anchor: 'left', point: { x: left, y: centerY } },
    { anchor: 'right', point: { x: right, y: centerY } },
    { anchor: 'top', point: { x: centerX, y: top } },
    { anchor: 'bottom', point: { x: centerX, y: bottom } },
    { anchor: 'center', point: { x: centerX, y: centerY } },
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
  return getPointDistance(
    getPPTLineEndpointPoint(line, 'start'),
    getPPTLineEndpointPoint(line, 'end'),
  )
}

function getPointDistance(a: Point, b: Point) {
  return Math.hypot(a.x - b.x, a.y - b.y)
}

function getPointAngle(center: Point, point: Point) {
  return Math.atan2(point.y - center.y, point.x - center.x) * 180 / Math.PI
}

function normalizePPTElementRotation(rotation: number) {
  const normalized = ((rotation % 360) + 360) % 360

  return Math.abs(normalized) < 0.001 ? 0 : Number(normalized.toFixed(3))
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
  const width = clamp(
    Math.ceil(Math.max(element.geometry.w, preferred.w + padding.x + 8)),
    24,
    maxWidth,
  )
  const contentWidth = Math.max(1, width - padding.x)
  const wrapped = measurePPTTextContentSize(element, {
    whiteSpace: 'pre-wrap',
    width: `${contentWidth}px`,
  })

  return {
    h: clamp(Math.ceil(wrapped.h + padding.y + 8), 24, maxHeight),
    w: width,
  }
}

function measurePPTTextContentSize(
  element: PPTTextElement,
  options: {
    whiteSpace: string
    width: string
  },
) {
  const style = element.style
  const measurer = document.createElement('div')

  measurer.style.position = 'fixed'
  measurer.style.left = '-10000px'
  measurer.style.top = '-10000px'
  measurer.style.width = options.width
  measurer.style.whiteSpace = options.whiteSpace ?? 'pre-wrap'
  measurer.style.overflowWrap = 'anywhere'
  measurer.style.fontFamily = getPPTTextFontFamilyCSS(style?.fontFamily)
  measurer.style.fontSize = `${style?.fontSize ?? 24}px`
  measurer.style.fontWeight = style?.fontWeight === 'bold'
    ? '700'
    : style?.fontWeight === 'semibold'
      ? '600'
      : '400'
  element.textBody.paragraphs.forEach((paragraph) => {
    const line = document.createElement('span')
    line.style.display = 'block'
    line.style.lineHeight = String(getPPTParagraphLineHeight(paragraph))
    line.style.marginBottom = `${getPPTParagraphSpacingAfter(paragraph)}px`
    line.style.marginTop = `${getPPTParagraphSpacingBefore(paragraph)}px`
    line.textContent = `${paragraph.bullet === 'bullet' ? '\u2022 ' : ''}${
      paragraph.runs.map((run) => run.text).join('') || ' '
    }`
    measurer.appendChild(line)
  })
  document.body.appendChild(measurer)

  const rect = measurer.getBoundingClientRect()
  measurer.remove()

  return {
    h: rect.height,
    w: rect.width,
  }
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

function getSpacingGuideLabelPoint(guide: CanvasSnapGuides['spacingGuides'][number]) {
  const points = guide.segments.flatMap((segment) => [segment.start, segment.end])
  const x = points.reduce((sum, point) => sum + point.x, 0) / points.length
  const y = points.reduce((sum, point) => sum + point.y, 0) / points.length

  return { x, y }
}

function createPPTSlideId(deck: PPTDeck) {
  const ids = new Set(deck.slides.map((slide) => slide.id))
  let next = deck.slides.length + 1
  let id = `slide-${next}`

  while (ids.has(id)) {
    next += 1
    id = `slide-${next}`
  }

  return id
}

function getPPTSlideDropIndex(
  slides: readonly PPTSlide[],
  sourceSlideId: string,
  targetSlideId: string,
  placement: PPTSlideDropPlacement,
) {
  const sourceIndex = slides.findIndex((slide) => slide.id === sourceSlideId)
  const targetIndex = slides.findIndex((slide) => slide.id === targetSlideId)

  if (sourceIndex < 0 || targetIndex < 0 || sourceIndex === targetIndex) {
    return null
  }

  let dropIndex = targetIndex + (placement === 'after' ? 1 : 0)

  if (sourceIndex < dropIndex) {
    dropIndex -= 1
  }

  const normalizedDropIndex = Math.max(0, Math.min(slides.length - 1, dropIndex))

  return normalizedDropIndex === sourceIndex ? null : normalizedDropIndex
}

function normalizePPTCommentBody(value: string) {
  return value.slice(0, PPT_COMMENT_BODY_MAX_LENGTH)
}

function getPPTCommentThreadWithBody(
  comment: PPTComment,
  body: string,
): PPTComment['thread'] {
  const thread = comment.thread && comment.thread.length > 0
    ? comment.thread
    : [{
        authorName: comment.authorName ?? PPT_COMMENT_DEFAULT_AUTHOR,
        body: comment.body,
        createdAt: comment.createdAt ?? PPT_COMMENT_DEFAULT_CREATED_AT,
        id: `${comment.id}:message-1`,
      }]
  const [first, ...rest] = thread

  return [{
    ...first,
    body,
  }, ...rest]
}

function clonePPTSlide(slide: PPTSlide, id: string): PPTSlide {
  return {
    ...slide,
    elements: slide.elements.map((element, index) => ({
      ...element,
      id: createPPTSlideElementId(id, element, index),
      name: element.name,
    })),
    id,
    name: `${slide.name} Copy`,
  }
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
  return PPT_STROKE_DASH_VALUES.has(value as PPTStrokeDash)
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
  if (!additive) {
    return [elementId]
  }

  return selection.includes(elementId)
    ? selection.filter((id) => id !== elementId)
    : [...selection, elementId]
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
  const memberIds = getPPTGroupMemberIds(slide, itemId, includeHidden)

  if (memberIds.length === 0) {
    return fallbackSelection
  }

  if (!additive) {
    return memberIds
  }

  const selected = new Set(selection)
  const allMembersSelected = memberIds.every((id) => selected.has(id))

  if (allMembersSelected) {
    return selection.filter((id) => !memberIds.includes(id))
  }

  return [
    ...selection,
    ...memberIds.filter((id) => !selected.has(id)),
  ]
}

function getPPTGroupMemberIds(
  slide: PPTSlide,
  elementId: string,
  includeHidden: boolean,
) {
  const element = findPPTElement(slide, elementId)

  if (!element?.groupId) {
    return []
  }

  return slide.elements
    .filter((candidate) =>
      candidate.groupId === element.groupId &&
      (includeHidden || candidate.visible !== false))
    .map((candidate) => candidate.id)
}

export default App
