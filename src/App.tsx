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
  ChevronUp,
  Circle,
  Copy,
  CopyPlus,
  Diamond,
  Download,
  Eye,
  EyeOff,
  FilePlus2,
  Grid2X2,
  Group,
  ImagePlus,
  Italic,
  Layers,
  List,
  Lock,
  Maximize2,
  Minus,
  MoveDown,
  MoveUp,
  Plus,
  Redo2,
  RotateCw,
  Search,
  SendToBack,
  Square,
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
  type KeyboardEvent as ReactKeyboardEvent,
  type MouseEvent as ReactMouseEvent,
  type PointerEvent as ReactPointerEvent,
  type RefObject,
} from 'react'
import {
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
  PPT_SLIDE_HEIGHT,
  PPT_SLIDE_WIDTH,
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
  type PPTElement,
  type PPTImage,
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
  type PPTTextBody,
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
  getPPTCanvasCommandAvailability,
  updatePPTElementBounds,
} from './pptCommandAdapter'
import { exportPPTDeckHTML } from './pptExport'
import {
  createPPTImportedImageElement,
  getPPTImageFileFromDataTransfer,
  getPPTImageFileFromList,
  readPPTImageFileSource,
  type PPTImageImportSource,
} from './pptImageImport'
import './App.css'

const PPT_CANVAS_COMMAND_CONFIG = createCanvasAffordanceConfig({
  commands: {
    group: true,
    lockSelection: true,
    ungroup: true,
    unlockAll: true,
  },
})

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
type PPTSurfaceCommand =
  | 'alignCenter'
  | 'alignLeft'
  | 'alignRight'
  | 'bringForward'
  | 'bringToFront'
  | 'delete'
  | 'duplicate'
  | 'group'
  | 'lockSelection'
  | 'sendBackward'
  | 'sendToBack'
  | 'ungroup'
  | 'unlockAll'
type PPTCommandAvailabilityKey =
  keyof ReturnType<typeof getPPTCanvasCommandAvailability>
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
type PPTContextMenuState = {
  x: number
  y: number
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
const PPT_DEFAULT_TEXT_BOUNDS = {
  h: 76,
  w: 360,
}
const PPT_TEXT_FONT_SIZE_MIN = 8
const PPT_TEXT_FONT_SIZE_MAX = 120
const PPT_TEXT_FONT_SIZE_STEP = 2

type LineCreationMode = 'arrow' | 'line'
type PPTCreationTool =
  | {
      kind: 'shape'
      shape: PPTShapeKind
    }
  | {
      kind: 'text'
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
  runStyle: PPTTextRunStyle
}

type Interaction =
  | {
      bounds: Bounds
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

function App() {
  const [deck, setDeck] = useState(SAMPLE_PPT_DECK)
  const [activeSlideId, setActiveSlideId] = useState(deck.slides[0].id)
  const [selection, setSelection] = useState<string[]>(['s1-title'])
  const [viewport, setViewport] = useState<Viewport>({ scale: 1, x: 0, y: 0 })
  const [editingId, setEditingId] = useState<string | null>(null)
  const [hoveredId, setHoveredId] = useState<string | null>(null)
  const [interaction, setInteraction] = useState<Interaction | null>(null)
  const [clipboard, setClipboard] = useState<PPTElement[]>([])
  const [lineCreationMode, setLineCreationMode] = useState<LineCreationMode | null>(null)
  const [creationTool, setCreationTool] = useState<PPTCreationTool | null>(null)
  const [contextMenu, setContextMenu] = useState<PPTContextMenuState | null>(null)
  const [findOpen, setFindOpen] = useState(false)
  const [findQuery, setFindQuery] = useState('')
  const [replaceQuery, setReplaceQuery] = useState('')
  const [activeFindIndex, setActiveFindIndex] = useState(0)
  const [showGrid, setShowGrid] = useState(true)
  const [past, setPast] = useState<PPTDeck[]>([])
  const [future, setFuture] = useState<PPTDeck[]>([])
  const stageRef = useRef<HTMLDivElement | null>(null)
  const imageInputRef = useRef<HTMLInputElement | null>(null)
  const findInputRef = useRef<HTMLInputElement | null>(null)
  const deckRef = useRef(deck)

  useEffect(() => {
    deckRef.current = deck
  }, [deck])

  useEffect(() => {
    globalThis.window?.dispatchEvent(new Event('ppt-ready'))
  }, [])

  const activeSlide = findPPTSlide(deck, activeSlideId)
  const activeSlideIndex = deck.slides.findIndex((slide) => slide.id === activeSlide.id)
  const scene = useMemo(() => createPPTCanvasScene(activeSlide), [activeSlide])
  const commandAdapter = useMemo(() => createPPTCanvasCommandAdapter(), [])
  const selectedBounds = scene.getBounds(selection)
  const selectedElement = findPPTElement(activeSlide, selection[0] ?? null)
  const selectedElements = useMemo(
    () => activeSlide.elements.filter((element) => selection.includes(element.id)),
    [activeSlide.elements, selection],
  )
  const selectedTextElements = useMemo(
    () => selectedElements.filter(isPPTTextElement),
    [selectedElements],
  )
  const selectedLineElement = selection.length === 1 && selectedElement?.kind === 'line'
    ? selectedElement
    : null
  const findMatches = useMemo(() => getPPTDeckTextMatches(deck, findQuery), [deck, findQuery])
  const clampedFindIndex = findMatches.length === 0
    ? 0
    : Math.min(activeFindIndex, findMatches.length - 1)
  const activeFindMatch = findMatches[clampedFindIndex] ?? null
  const exportCode = useMemo(() => exportPPTDeckHTML(deck), [deck])
  const hasLockedItems = activeSlide.elements.some((element) => element.locked === true)
  const hasLockedSelection = selectedElements.some((element) => element.locked === true)
  const hasHiddenSelection = selectedElements.some((element) => element.visible === false)
  const hasGroupedSelection = selectedElements.some((element) => Boolean(element.groupId))
  const canFormatSelectedText = selectedTextElements.length > 0 &&
    selectedTextElements.length === selectedElements.length &&
    !hasLockedSelection &&
    !hasHiddenSelection
  const textQuickFormatState = canFormatSelectedText
    ? getPPTTextQuickFormatState(selectedTextElements)
    : null
  const commandAvailability = useMemo(() => getPPTCanvasCommandAvailability({
    canPaste: clipboard.length > 0,
    canRedo: future.length > 0,
    canUndo: past.length > 0,
    hasGroupedSelection,
    hasHiddenSelection,
    hasLockedItems,
    hasLockedSelection,
    selection,
  }), [
    clipboard.length,
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
  const canResizeSelection = selectedBounds
    ? scene.canResizeSelection?.(selection) ?? true
    : false

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

  useLayoutEffect(() => {
    fitSlide()
  }, [activeSlideId, fitSlide])

  useEffect(() => {
    window.addEventListener('resize', fitSlide)

    return () => window.removeEventListener('resize', fitSlide)
  }, [fitSlide])

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
      if (isEditableTarget(event.target)) {
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

      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'b') {
        if (canFormatSelectedText) {
          event.preventDefault()
          toggleSelectedTextBold()
        }
        return
      }

      const shortcutTool = getPPTCreationToolForShortcut(event)

      if (shortcutTool) {
        event.preventDefault()
        activatePPTCreationTool(shortcutTool)
        return
      }

      if (isPPTSelectToolShortcut(event)) {
        event.preventDefault()
        setCreationTool(null)
        setLineCreationMode(null)
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
        setLineCreationMode(null)
        setCreationTool(null)
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

    window.addEventListener('keydown', onKeyDown)

    return () => window.removeEventListener('keydown', onKeyDown)
  })

  useEffect(() => {
    function onPaste(event: ClipboardEvent) {
      if (isEditableTarget(event.target)) {
        return
      }

      const file = getPPTImageFileFromDataTransfer(event.clipboardData)

      if (!file) {
        return
      }

      event.preventDefault()
      void insertPPTImageFile(file)
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
    setContextMenu(null)
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
    setEditingId(null)
    setInteraction(null)
    setLineCreationMode(null)
    setCreationTool(null)
    setContextMenu(null)

    if (activeFindMatch) {
      focusPPTFindMatch(activeFindMatch)
    }
  }

  function closeFindStrip() {
    setFindOpen(false)
  }

  function focusPPTFindMatch(match: PPTFindMatch) {
    setActiveSlideId(match.slideId)
    setSelection([match.elementId])
    setEditingId(null)
    setInteraction(null)
    setLineCreationMode(null)
    setCreationTool(null)
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

  function activatePPTCreationTool(tool: PPTCreationTool) {
    setCreationTool((current) => arePPTCreationToolsEqual(current, tool) ? null : tool)
    setLineCreationMode(null)
    setEditingId(null)
    setContextMenu(null)
  }

  function activateLineCreationMode(mode: LineCreationMode) {
    setLineCreationMode((current) => current === mode ? null : mode)
    setCreationTool(null)
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

  function copySelection() {
    const selected = activeSlide.elements.filter((element) => selection.includes(element.id))

    if (selected.length === 0) {
      return
    }

    setClipboard(selected)
    void navigator.clipboard?.writeText(JSON.stringify({
      elements: selected,
      type: 'application/ppt-elements+json',
    })).catch(() => undefined)
  }

  function cutSelection() {
    if (!commandAvailability.cut) {
      return
    }

    copySelection()
    deleteSelection()
  }

  function pasteSelection() {
    if (!commandAvailability.paste) {
      return
    }

    commitElementCommand((slide) => {
      const pasted = commandAdapter.pasteItems({
        clipboard,
        createId: createPPTElementIdFactory(slide),
        offset: { x: 28, y: 28 },
      })

      if (pasted.length === 0) {
        return null
      }

      return {
        items: [...slide.elements, ...pasted],
        selection: pasted.map((element) => element.id),
      }
    })
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
      case 'delete':
        deleteSelection()
        break
      case 'duplicate':
        duplicateSelection()
        break
      case 'group':
        groupSelection()
        break
      case 'lockSelection':
        lockSelectedElements()
        break
      case 'sendBackward':
        reorderSelection('sendBackward')
        break
      case 'sendToBack':
        reorderSelection('sendToBack')
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

  function updateElementTextStyle(
    elementId: string,
    field: keyof PPTTextStyle,
    value: string | number,
  ) {
    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) => {
        if (!isPPTTextElement(element)) {
          return element
        }

        const style = {
          color: '#111827',
          fontSize: 24,
          ...element.style,
        }

        return {
          ...element,
          style: {
            ...style,
            [field]: value,
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

  function updateShapeFill(elementId: string, color: string) {
    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) =>
        element.kind === 'shape'
          ? { ...element, fill: { color } }
          : element),
    )
  }

  function updateShapeKind(elementId: string, shape: PPTShapeKind) {
    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) =>
        element.kind === 'shape'
          ? { ...element, shape }
          : element),
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

  function toggleElementLocked(elementId: string) {
    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) => ({
        ...element,
        locked: element.locked === true ? false : true,
      })),
    )
  }

  function toggleElementVisibility(elementId: string) {
    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) => ({
        ...element,
        visible: element.visible === false,
      })),
    )
  }

  function updateElementStroke(
    elementId: string,
    field: 'color' | 'width',
    value: string | number,
  ) {
    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) => {
        if (element.kind !== 'shape' && element.kind !== 'line') {
          return element
        }

        const stroke = {
          color: '#111827',
          width: 2,
          ...element.stroke,
        }

        return {
          ...element,
          stroke: {
            ...stroke,
            [field]: value,
          },
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

  function updateSlideName(name: string) {
    commitDeck((current) => updatePPTDeckSlide(current, activeSlide.id, (slide) => ({
      ...slide,
      name,
    })))
  }

  function updateSlideBackground(color: string) {
    commitDeck((current) => updatePPTDeckSlide(current, activeSlide.id, (slide) => ({
      ...slide,
      background: { color },
    })))
  }

  function updateSlideNotes(notes: string) {
    commitDeck((current) => updatePPTDeckSlide(current, activeSlide.id, (slide) => ({
      ...slide,
      notes,
    })))
  }

  function copyHTML() {
    void navigator.clipboard.writeText(exportCode).catch(() => undefined)
  }

  function downloadHTML() {
    const url = URL.createObjectURL(new Blob([exportCode], {
      type: 'text/html;charset=utf-8',
    }))
    const anchor = document.createElement('a')
    anchor.href = url
    anchor.download = 'ppt-subset.html'
    anchor.click()
    URL.revokeObjectURL(url)
  }

  function screenToWorld(event: Pick<PointerEvent, 'clientX' | 'clientY'>) {
    const rect = stageRef.current?.getBoundingClientRect()

    return getCanvasViewportWorldPoint(viewport, {
      x: event.clientX - (rect?.left ?? 0),
      y: event.clientY - (rect?.top ?? 0),
    })
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
    if (getPPTImageFileFromDataTransfer(event.dataTransfer)) {
      event.preventDefault()
    }
  }

  function handleStageDrop(event: ReactDragEvent<HTMLDivElement>) {
    const file = getPPTImageFileFromDataTransfer(event.dataTransfer)

    if (!file) {
      return
    }

    event.preventDefault()
    void insertPPTImageFile(file, screenToWorld(event.nativeEvent))
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
      creationTool.kind === 'text' ? 'text' : creationTool.shape,
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

    setSelection(nextSelection)

    if (!bounds || hasLockedTarget) {
      return
    }

    setInteraction({
      bounds,
      kind: 'move',
      selection: nextSelection,
      slideId: activeSlide.id,
      snapGuides: EMPTY_CANVAS_SNAP_GUIDES,
      startDeck: deckRef.current,
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

    const point = screenToWorld(event.nativeEvent)

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

    if (interaction.kind === 'element-create') {
      setCreationTool(null)

      if (interaction.tool.kind === 'text') {
        setEditingId(interaction.elementId)
      }
    }

    if (
      interaction.kind !== 'marquee' &&
      JSON.stringify(deckRef.current) !== JSON.stringify(interaction.startDeck)
    ) {
      setPast((history) => [...history.slice(-79), interaction.startDeck])
      setFuture([])
    }

    setInteraction(null)
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
  const selectionCommandBarWidth = textQuickFormatState ? 420 : 236
  const selectionCommandAnchor = selectedBounds &&
    !editingId &&
    !interaction &&
    !contextMenu &&
    !creationTool &&
    !lineCreationMode
    ? getPPTSelectionCommandAnchor({
        barWidth: selectionCommandBarWidth,
        bounds: selectedBounds,
        stage: stageRef.current,
        viewport,
      })
    : null
  const selectionFloatingCommandGroups = getPPTCommandSurfaceGroups({
    availability: commandAvailability,
    surface: 'selection-floating-bar',
  })
  const contextCommandGroups = getPPTCommandSurfaceGroups({
    availability: commandAvailability,
    surface: 'context-menu',
  })

  return (
    <main className="ppt-app" data-ppt-app>
      <header className="ppt-topbar">
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
          <button className="ppt-icon-button" data-ppt-insert-image onClick={() => imageInputRef.current?.click()} title="Add image" type="button">
            <ImagePlus size={17} />
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
          <button className="ppt-icon-button" onClick={() => zoom('out')} title="Zoom out" type="button">
            <ZoomOut size={17} />
          </button>
          <button className="ppt-icon-button" onClick={fitSlide} title="Fit slide" type="button">
            <Maximize2 size={17} />
          </button>
          <button className="ppt-icon-button" onClick={() => zoom('in')} title="Zoom in" type="button">
            <ZoomIn size={17} />
          </button>
          <button aria-pressed={showGrid} className="ppt-icon-button" data-ppt-view-grid onClick={() => setShowGrid((current) => !current)} title="Toggle grid" type="button">
            <Grid2X2 size={17} />
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
        <div className="ppt-slide-list" role="listbox" aria-label="Slides">
          {deck.slides.map((slide, index) => (
            <SlideThumb
              active={slide.id === activeSlide.id}
              index={index}
              key={slide.id}
              slide={slide}
              onSelect={() => {
                selectSlide(slide.id)
              }}
            />
          ))}
        </div>
      </aside>

      <section
        className="ppt-stage-shell"
        data-creation-tool={getPPTCreationToolDataValue(creationTool)}
        data-grid={showGrid ? 'true' : 'false'}
        data-line-tool={lineCreationMode ?? undefined}
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
            data-ppt-slide={activeSlide.id}
            style={{ background: activeSlide.background?.color ?? '#ffffff' }}
          >
            {activeSlide.elements.filter((element) => element.visible !== false).map((element) => (
              <PPTElementView
                editing={editingId === element.id}
                element={element}
                hovered={hoveredId === element.id}
                findActive={findOpen && activeFindMatch?.elementId === element.id}
                key={element.id}
                selected={selection.includes(element.id)}
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
              />
            ))}
            {selection.length > 0 && !editingId ? <FrameGuides /> : null}
            {selectedBounds ? (
              <SelectionOverlay
                bounds={selectedBounds}
                canResize={canResizeSelection}
                scale={viewport.scale}
                selectedElements={selectedElements}
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
          </div>
        </div>
        <PPTContextCommandMenu
          groups={contextCommandGroups}
          menu={contextMenu}
          onClose={() => setContextMenu(null)}
          onCommand={runPPTSurfaceCommand}
        />
      </section>

      <Inspector
        exportCode={exportCode}
        selection={selection}
        selectedElement={selectedElement}
        slide={activeSlide}
        onCommitText={commitText}
        onCopyHTML={copyHTML}
        onDownloadHTML={downloadHTML}
        onElementGeometryChange={updateElementGeometry}
        onImageFitChange={updateImageFit}
        onElementLockToggle={toggleElementLocked}
        onElementNameChange={updateElementName}
        onElementRotationChange={updateElementRotation}
        onLineMarkerChange={updateLineMarker}
        onLineRouteChange={updateLineRoute}
        onParagraphBulletChange={updateParagraphBullet}
        onElementTextStyleChange={updateElementTextStyle}
        onElementVisibilityToggle={toggleElementVisibility}
        onLayerSelect={(elementId, additive) => {
          setSelection((current) =>
            getPPTLayerSelection(current, elementId, additive, activeSlide))
        }}
        onParagraphAlignChange={updateParagraphAlign}
        onShapeKindChange={updateShapeKind}
        onSlideBackgroundChange={updateSlideBackground}
        onShapeFillChange={updateShapeFill}
        onElementStrokeChange={updateElementStroke}
        onSlideNameChange={updateSlideName}
        onSlideNotesChange={updateSlideNotes}
      />
    </main>
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
      {(['left', 'center', 'right'] as const).map((align) => (
        <button
          aria-label={`Align text ${align}`}
          aria-pressed={state.align === align}
          className="ppt-floating-command"
          data-ppt-text-quick={`align-${align}`}
          key={align}
          title={`Align text ${align}`}
          type="button"
          onClick={(event) => {
            event.preventDefault()
            event.stopPropagation()
            onParagraphAlign(align)
          }}
        >
          <PPTTextAlignIcon align={align} size={16} />
        </button>
      ))}
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
  if (!menu || groups.length === 0) {
    return null
  }

  return (
    <div
      aria-label="Selection commands"
      className="ppt-context-menu"
      data-ppt-context-menu
      role="menu"
      style={{
        left: menu.x,
        top: menu.y,
      }}
      onContextMenu={(event) => event.preventDefault()}
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
    case 'delete':
      return <Trash2 size={size} />
    case 'duplicate':
      return <CopyPlus size={size} />
    case 'group':
      return <Group size={size} />
    case 'lockSelection':
      return <Lock size={size} />
    case 'sendBackward':
      return <MoveDown size={size} />
    case 'sendToBack':
      return <SendToBack size={size} />
    case 'ungroup':
      return <Ungroup size={size} />
    case 'unlockAll':
      return <Unlock size={size} />
  }
}

function SlideThumb({
  active,
  index,
  onSelect,
  slide,
}: {
  active: boolean
  index: number
  onSelect: () => void
  slide: PPTSlide
}) {
  return (
    <button
      aria-current={active ? 'page' : undefined}
      aria-label={`Open ${slide.name}`}
      className="ppt-thumb"
      onClick={onSelect}
      role="option"
      type="button"
    >
      <span className="ppt-thumb-preview" style={{ background: slide.background?.color ?? '#fff' }}>
        {slide.elements.map((element) => (
          <span
            className={getPPTThumbElementClassName(element)}
            data-line-end-marker={element.kind === 'line' ? element.endMarker : undefined}
            data-line-start-marker={element.kind === 'line' ? element.startMarker : undefined}
            data-ppt-image-fit={element.kind === 'image'
              ? getPPTImageFit(element)
              : undefined}
            data-ppt-thumb-bullet={isPPTTextElement(element) && hasPPTTextBodyBullet(element.textBody)
              ? 'true'
              : undefined}
            data-shape={element.kind === 'shape' ? element.shape : undefined}
            key={element.id}
            style={{
              background: element.kind === 'shape'
                ? element.fill.color
                : element.kind === 'image'
                  ? undefined
                  : element.kind === 'textBox'
                    ? '#cbd5e1'
                    : undefined,
              backgroundImage: element.kind === 'image'
                ? `url(${element.src})`
                : undefined,
              backgroundSize: element.kind === 'image'
                ? getPPTImageFit(element)
                : undefined,
              height: `${(element.geometry.h / PPT_SLIDE_HEIGHT) * 100}%`,
              left: `${(element.geometry.x / PPT_SLIDE_WIDTH) * 100}%`,
              top: `${(element.geometry.y / PPT_SLIDE_HEIGHT) * 100}%`,
              transform: element.geometry.rotation
                ? `rotate(${element.geometry.rotation}deg)`
                : undefined,
              width: `${(element.geometry.w / PPT_SLIDE_WIDTH) * 100}%`,
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
  findActive,
  hovered,
  onCommitText,
  onContextMenu,
  onEdit,
  onPointerDown,
  onPointerEnter,
  onPointerLeave,
  onStopEdit,
  selected,
}: {
  editing: boolean
  element: PPTElement
  findActive: boolean
  hovered: boolean
  onCommitText: (elementId: string, text: string) => void
  onContextMenu: (event: ReactMouseEvent<HTMLDivElement>, elementId: string) => void
  onEdit: () => void
  onPointerDown: (event: ReactPointerEvent<HTMLDivElement>, elementId: string) => void
  onPointerEnter: () => void
  onPointerLeave: () => void
  onStopEdit: () => void
  selected: boolean
}) {
  const style = pptElementStyle(element)
  const textBody = isPPTTextElement(element) ? element.textBody : null
  const text = isPPTTextElement(element) ? readPPTText(element.textBody) : ''
  const editorRef = useRef<HTMLDivElement | null>(null)

  useEffect(() => {
    if (!editing) {
      return
    }

    editorRef.current?.focus()
  }, [editing])

  return (
    <div
      className="ppt-element"
      data-hovered={hovered ? 'true' : 'false'}
      data-group-id={element.groupId}
      data-kind={element.kind}
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
      data-ppt-find-active={findActive ? 'true' : undefined}
      data-ppt-image-fit={element.kind === 'image'
        ? getPPTImageFit(element)
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
          alt={element.alt}
          draggable={false}
          src={element.src}
          style={{ objectFit: getPPTImageFit(element) }}
        />
      ) : element.kind === 'line' ? (
        <PPTLineSvg element={element} />
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

function PPTTextBodyView({ body }: { body: PPTTextBody }) {
  return (
    <>
      {body.paragraphs.map((paragraph, index) => (
        <span
          className="ppt-text-paragraph"
          data-ppt-bullet={paragraph.bullet === 'bullet' ? 'true' : undefined}
          key={index}
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
          strokeLinecap="round"
          strokeLinejoin="round"
          strokeWidth={element.stroke.width}
        />
      ) : (
        <line
          markerEnd={markerEnd}
          markerStart={markerStart}
          stroke={element.stroke.color}
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

function SelectionOverlay({
  bounds,
  canResize,
  onRotatePointerDown,
  onResizeHandleDoubleClick,
  onResizePointerDown,
  scale,
  selectedElements,
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
}) {
  return (
    <>
      <Box className="ppt-selection-box" bounds={bounds} />
      <div
        className="ppt-size-capsule"
        style={{
          left: bounds.x + bounds.w / 2,
          top: bounds.y + bounds.h + 8,
          transform: `translateX(-50%) scale(${1 / scale})`,
        }}
      >
        {selectedElements.length === 1
          ? `${Math.round(bounds.w)} x ${Math.round(bounds.h)}`
          : `${selectedElements.length} objects`}
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

function FrameGuides() {
  return (
    <>
      <div
        className="ppt-frame-guide ppt-frame-guide-vertical ppt-frame-guide-center"
        style={{ left: PPT_SLIDE_WIDTH / 2 }}
      />
      <div
        className="ppt-frame-guide ppt-frame-guide-horizontal ppt-frame-guide-center"
        style={{ top: PPT_SLIDE_HEIGHT / 2 }}
      />
      <div className="ppt-frame-guide ppt-frame-guide-vertical" style={{ left: 84 }} />
      <div
        className="ppt-frame-guide ppt-frame-guide-vertical"
        style={{ left: PPT_SLIDE_WIDTH - 84 }}
      />
      <div className="ppt-frame-guide ppt-frame-guide-horizontal" style={{ top: 64 }} />
      <div
        className="ppt-frame-guide ppt-frame-guide-horizontal"
        style={{ top: PPT_SLIDE_HEIGHT - 64 }}
      />
    </>
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

function Guides({ guides, scale }: { guides: CanvasSnapGuides; scale: number }) {
  return (
    <>
      {guides.alignmentGuides.map((guide, index) => (
        <div
          className={`ppt-guide ppt-guide-${guide.orientation}`}
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
                key={`${segment.start.x}-${segment.start.y}-${segment.end.x}-${segment.end.y}-${segmentIndex}`}
                style={getSpacingGuideSegmentStyle(segment)}
              />
            ))}
            <span
              className="ppt-spacing-label"
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

function Inspector({
  exportCode,
  onCommitText,
  onCopyHTML,
  onDownloadHTML,
  onElementGeometryChange,
  onElementLockToggle,
  onElementNameChange,
  onElementRotationChange,
  onElementStrokeChange,
  onElementTextStyleChange,
  onElementVisibilityToggle,
  onImageFitChange,
  onLayerSelect,
  onLineMarkerChange,
  onLineRouteChange,
  onParagraphBulletChange,
  onParagraphAlignChange,
  onShapeFillChange,
  onShapeKindChange,
  onSlideBackgroundChange,
  onSlideNameChange,
  onSlideNotesChange,
  selection,
  selectedElement,
  slide,
}: {
  exportCode: string
  onCommitText: (elementId: string, text: string) => void
  onCopyHTML: () => void
  onDownloadHTML: () => void
  onElementGeometryChange: (
    elementId: string,
    field: 'h' | 'w' | 'x' | 'y',
    value: number,
  ) => void
  onElementLockToggle: (elementId: string) => void
  onElementNameChange: (elementId: string, name: string) => void
  onElementRotationChange: (elementId: string, rotation: number) => void
  onElementStrokeChange: (
    elementId: string,
    field: 'color' | 'width',
    value: string | number,
  ) => void
  onElementTextStyleChange: (
    elementId: string,
    field: keyof PPTTextStyle,
    value: string | number,
  ) => void
  onElementVisibilityToggle: (elementId: string) => void
  onImageFitChange: (
    elementId: string,
    fit: PPTImageFit,
  ) => void
  onLayerSelect: (elementId: string, additive: boolean) => void
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
  onShapeFillChange: (elementId: string, color: string) => void
  onShapeKindChange: (elementId: string, shape: PPTShapeKind) => void
  onSlideBackgroundChange: (color: string) => void
  onSlideNameChange: (name: string) => void
  onSlideNotesChange: (notes: string) => void
  selection: string[]
  selectedElement: PPTElement | null
  slide: PPTSlide
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

  return (
    <aside className="ppt-inspector" aria-label="Inspector">
      <div className="ppt-panel-header">
        <h2>Slide</h2>
      </div>
      <section className="ppt-panel-section">
        <label className="ppt-field">
          <span>Name</span>
          <input
            value={slide.name}
            onChange={(event) => onSlideNameChange(event.target.value)}
          />
        </label>
        <label className="ppt-field">
          <span>Background</span>
          <input
            data-ppt-slide-field="background"
            type="color"
            value={slide.background?.color ?? '#ffffff'}
            onChange={(event) => onSlideBackgroundChange(event.target.value)}
          />
        </label>
        <label className="ppt-field">
          <span>Notes</span>
          <textarea
            value={slide.notes ?? ''}
            onChange={(event) => onSlideNotesChange(event.target.value)}
          />
        </label>
      </section>

      <div className="ppt-panel-header">
        <h2>Selection</h2>
      </div>
      <section className="ppt-panel-section">
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
            {isPPTTextElement(selectedElement) ? (
              <>
                <label className="ppt-field">
                  <span>Text</span>
                  <textarea
                    value={readPPTText(selectedElement.textBody)}
                    onChange={(event) => onCommitText(selectedElement.id, event.target.value)}
                  />
                </label>
                <div className="ppt-geometry-grid">
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
                  <div className="ppt-segmented-control" role="group" aria-label="Paragraph align">
                    <button
                      aria-pressed={paragraphBullet}
                      data-ppt-paragraph-bullet
                      type="button"
                      onClick={() =>
                        onParagraphBulletChange(selectedElement.id, !paragraphBullet)}
                    >
                      bullet
                    </button>
                    {(['left', 'center', 'right'] as const).map((align) => (
                      <button
                        aria-pressed={paragraphAlign === align}
                        data-ppt-paragraph-align={align}
                        key={align}
                        type="button"
                        onClick={() => onParagraphAlignChange(selectedElement.id, align)}
                      >
                        {align}
                      </button>
                    ))}
                  </div>
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
                <div className="ppt-geometry-grid">
                  <label className="ppt-field">
                    <span>Fill</span>
                    <input
                      data-ppt-style-field="fill"
                      type="color"
                      value={selectedElement.fill.color}
                      onChange={(event) =>
                        onShapeFillChange(selectedElement.id, event.target.value)}
                    />
                  </label>
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
                </div>
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
              </>
            ) : null}
            {selectedElement.kind === 'image' ? (
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
            ) : null}
            {selectedElement.kind === 'line' ? (
              <>
                <div className="ppt-geometry-grid">
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
        ) : (
          <span className="ppt-muted">None</span>
        )}
      </section>

      <div className="ppt-panel-header">
        <h2>Objects</h2>
      </div>
      <section className="ppt-panel-section">
        <div className="ppt-layer-list" aria-label="Selection pane">
          {slide.elements.map((element) => (
            <div
              aria-selected={selection.includes(element.id)}
              className="ppt-layer-row"
              data-grouped={element.groupId ? 'true' : 'false'}
              data-hidden={element.visible === false ? 'true' : 'false'}
              data-locked={element.locked === true ? 'true' : 'false'}
              data-ppt-layer-row={element.id}
              key={element.id}
            >
              <button
                className="ppt-layer-select"
                data-ppt-layer-select={element.id}
                type="button"
                onClick={(event) =>
                  onLayerSelect(element.id, event.metaKey || event.ctrlKey || event.shiftKey)}
              >
                <span className="ppt-layer-kind">
                  <Layers size={14} />
                </span>
                <span className="ppt-layer-name">{element.name}</span>
              </button>
              <div className="ppt-layer-actions">
                <button
                  aria-label={element.visible === false ? 'Show object' : 'Hide object'}
                  className="ppt-layer-icon-button"
                  data-ppt-layer-visibility={element.id}
                  title={element.visible === false ? 'Show object' : 'Hide object'}
                  type="button"
                  onClick={(event) => {
                    event.stopPropagation()
                    onElementVisibilityToggle(element.id)
                  }}
                >
                  {element.visible === false ? <EyeOff size={14} /> : <Eye size={14} />}
                </button>
                <button
                  aria-label={element.locked === true ? 'Unlock object' : 'Lock object'}
                  className="ppt-layer-icon-button"
                  data-ppt-layer-lock={element.id}
                  title={element.locked === true ? 'Unlock object' : 'Lock object'}
                  type="button"
                  onClick={(event) => {
                    event.stopPropagation()
                    onElementLockToggle(element.id)
                  }}
                >
                  {element.locked === true ? <Lock size={14} /> : <Unlock size={14} />}
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
    height: element.geometry.h,
    left: element.geometry.x,
    top: element.geometry.y,
    transform: element.geometry.rotation
      ? `rotate(${element.geometry.rotation}deg)`
      : undefined,
    transformOrigin: 'center',
    width: element.geometry.w,
  }

  if (element.kind === 'image' || element.kind === 'line') {
    return base
  }

  return {
    ...base,
    ...pptTextStyle(element.style),
    background: element.kind === 'shape' ? element.fill.color : 'transparent',
    border: element.kind === 'shape' && element.stroke
      ? `${element.stroke.width}px solid ${element.stroke.color}`
      : undefined,
    textAlign: getPPTElementParagraphAlign(element),
  }
}

function pptTextStyle(style: PPTTextStyle | undefined): CSSProperties {
  if (!style) {
    return {}
  }

  return {
    color: style.color,
    fontSize: style.fontSize,
    fontWeight: style.fontWeight === 'bold'
      ? 700
      : style.fontWeight === 'semibold'
        ? 600
        : 400,
    lineHeight: 1.14,
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

function getPPTCommandSurfaceGroups({
  availability,
  surface,
}: {
  availability: ReturnType<typeof getPPTCanvasCommandAvailability>
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

function getDefaultPPTTextStyle(): PPTTextStyle {
  return {
    color: '#111827',
    fontSize: 24,
    fontWeight: 'regular',
  }
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

  return createCanvasShape({
    adapter: PPT_CANVAS_CREATION_ADAPTER,
    createId: () => id,
    currentWorld: current,
    shapeType: tool.shape,
    startWorld: start,
  })
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

  return null
}

function isPPTSelectToolShortcut(event: KeyboardEvent) {
  return !event.metaKey &&
    !event.ctrlKey &&
    !event.altKey &&
    doesEventMatchCanvasToolShortcut(event, CANVAS_TOOL_AFFORDANCES.select.keyboardShortcut)
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
  return left?.kind === right.kind &&
    (left.kind === 'text' || right.kind === 'text' || left.shape === right.shape)
}

function isPPTShapeCreationTool(
  tool: PPTCreationTool | null,
  shape: PPTShapeKind,
) {
  return tool?.kind === 'shape' && tool.shape === shape
}

function getPPTCreationToolDataValue(tool: PPTCreationTool | null) {
  if (!tool) {
    return undefined
  }

  return tool.kind === 'text' ? 'text' : tool.shape
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
  )

  return buildPPTTextBodyFromTokens([
    ...tokens.slice(0, safeStart),
    ...replacementTokens,
    ...tokens.slice(safeEnd),
  ], body.paragraphs[0]?.align, body.paragraphs[0]?.bullet)
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
          runStyle,
        })
      }
    })

    if (paragraphIndex < body.paragraphs.length - 1) {
      tokens.push({
        align: paragraph.align,
        bullet: paragraph.bullet,
        char: '\n',
        runStyle: getPPTParagraphFallbackRunStyle(paragraph),
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
): PPTTextToken[] {
  const tokens: PPTTextToken[] = []

  for (let index = 0; index < text.length; index += 1) {
    tokens.push({
      align,
      bullet,
      char: text[index],
      runStyle,
    })
  }

  return tokens
}

function buildPPTTextBodyFromTokens(
  tokens: PPTTextToken[],
  fallbackAlign: PPTParagraph['align'] | undefined,
  fallbackBullet: PPTParagraph['bullet'] | undefined,
): PPTTextBody {
  const paragraphs: PPTParagraph[] = []
  let currentAlign = fallbackAlign
  let currentBullet = fallbackBullet
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
    paragraphs.push(createPPTParagraph(currentRuns, currentAlign, currentBullet))
    currentRuns = []
    currentRunStyle = null
    currentText = ''
    currentAlign = fallbackAlign
    currentBullet = fallbackBullet
  }

  tokens.forEach((token) => {
    if (token.char === '\n') {
      flushParagraph()
      currentAlign = token.align ?? fallbackAlign
      currentBullet = token.bullet ?? fallbackBullet
      return
    }

    if (!currentRunStyle && currentText.length === 0 && currentRuns.length === 0) {
      currentAlign = token.align ?? fallbackAlign
      currentBullet = token.bullet ?? fallbackBullet
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
): PPTParagraph {
  return {
    ...(align ? { align } : {}),
    ...(bullet ? { bullet } : {}),
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

  const text = readPPTText(element.textBody) || ' '
  const style = element.style
  const measurer = document.createElement('div')

  measurer.style.position = 'fixed'
  measurer.style.left = '-10000px'
  measurer.style.top = '-10000px'
  measurer.style.width = 'max-content'
  measurer.style.whiteSpace = 'pre'
  measurer.style.fontFamily = 'Inter, ui-sans-serif, system-ui, sans-serif'
  measurer.style.fontSize = `${style?.fontSize ?? 24}px`
  measurer.style.fontWeight = style?.fontWeight === 'bold'
    ? '700'
    : style?.fontWeight === 'semibold'
      ? '600'
      : '400'
  measurer.style.lineHeight = '1.14'
  measurer.textContent = text
  document.body.appendChild(measurer)

  const rect = measurer.getBoundingClientRect()
  measurer.remove()

  const padding = element.kind === 'shape' ? 36 : 4

  return {
    h: Math.ceil(rect.height + padding),
    w: Math.ceil(rect.width + padding),
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
