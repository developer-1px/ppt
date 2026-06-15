import {
  AlignCenterHorizontal,
  AlignCenterVertical,
  AlignEndHorizontal,
  AlignEndVertical,
  AlignHorizontalDistributeCenter,
  AlignStartHorizontal,
  AlignStartVertical,
  AlignVerticalDistributeCenter,
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
  Layers,
  Lock,
  Maximize2,
  MoveDown,
  MoveUp,
  Redo2,
  RotateCw,
  SendToBack,
  Square,
  Trash2,
  Type,
  Undo2,
  Unlock,
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
  type CSSProperties,
  type MouseEvent as ReactMouseEvent,
  type PointerEvent as ReactPointerEvent,
} from 'react'
import {
  RESIZE_HANDLES,
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
  alignCanvasCommand,
  createCanvasAffordanceConfig,
  deleteCanvasCommand,
  distributeCanvasCommand,
  duplicateCanvasCommand,
  lockCanvasCommand,
  nudgeCanvasCommand,
  reorderCanvasCommand,
  selectAllCanvasCommand,
  unlockAllCanvasCommand,
  type CanvasAlignMode,
  type CanvasCommandItemsResult,
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
  type PPTParagraph,
  type PPTShape,
  type PPTShapeKind,
  type PPTSlide,
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
import './App.css'

const PPT_CANVAS_COMMAND_CONFIG = createCanvasAffordanceConfig({
  commands: {
    group: false,
    lockSelection: true,
    ungroup: false,
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
  const [showGrid, setShowGrid] = useState(true)
  const [past, setPast] = useState<PPTDeck[]>([])
  const [future, setFuture] = useState<PPTDeck[]>([])
  const stageRef = useRef<HTMLDivElement | null>(null)
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
  const exportCode = useMemo(() => exportPPTDeckHTML(deck), [deck])
  const hasLockedItems = activeSlide.elements.some((element) => element.locked === true)
  const hasLockedSelection = selectedElements.some((element) => element.locked === true)
  const hasHiddenSelection = selectedElements.some((element) => element.visible === false)
  const commandAvailability = useMemo(() => getPPTCanvasCommandAvailability({
    canPaste: clipboard.length > 0,
    canRedo: future.length > 0,
    canUndo: past.length > 0,
    hasHiddenSelection,
    hasLockedItems,
    hasLockedSelection,
    selection,
  }), [
    clipboard.length,
    future.length,
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
    function onKeyDown(event: KeyboardEvent) {
      if (isEditableTarget(event.target)) {
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
        event.preventDefault()
        pasteSelection()
        return
      }

      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'd') {
        event.preventDefault()
        duplicateSelection()
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
        elements: result.items,
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
  }

  function activateRelativeSlide(delta: number) {
    const index = deck.slides.findIndex((slide) => slide.id === activeSlide.id)
    const nextSlide = deck.slides[index + delta]

    if (nextSlide) {
      selectSlide(nextSlide.id)
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

  function addTextBox() {
    commitDeck((current) => updatePPTDeckSlide(current, activeSlide.id, (slide) => {
      const id = createPPTElementId(slide, 'text')
      const element: PPTElement = {
        geometry: { h: 76, w: 360, x: 140, y: 150 },
        id,
        kind: 'textBox',
        name: 'Text',
        style: { color: '#111827', fontSize: 30, fontWeight: 'semibold' },
        textBody: createPPTTextBody('New text'),
      }

      setSelection([id])

      return {
        ...slide,
        elements: [...slide.elements, element],
      }
    }))
  }

  function addShape(shape: PPTShapeKind = 'rect') {
    commitDeck((current) => updatePPTDeckSlide(current, activeSlide.id, (slide) => {
      const id = createPPTElementId(slide, 'shape')
      const label = getPPTShapeLabel(shape)
      const element: PPTShape = {
        fill: { color: '#eef2ff' },
        geometry: { h: 150, w: 260, x: 160, y: 260 },
        id,
        kind: 'shape',
        name: label,
        shape,
        stroke: { color: '#6366f1', width: 2 },
        style: { color: '#312e81', fontSize: 24, fontWeight: 'semibold' },
        textBody: createPPTTextBody(label),
      }

      setSelection([id])

      return {
        ...slide,
        elements: [...slide.elements, element],
      }
    }))
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
        if (element.kind === 'image') {
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

  function updateShapeStroke(
    elementId: string,
    field: 'color' | 'width',
    value: string | number,
  ) {
    commitDeck((current) =>
      updatePPTDeckElement(current, activeSlide.id, elementId, (element) => {
        if (element.kind !== 'shape') {
          return element
        }

        const stroke = {
          color: '#111827',
          width: 1,
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
        if (element.kind === 'image' || !element.textBody) {
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

  function handleElementPointerDown(
    event: ReactPointerEvent<HTMLDivElement>,
    elementId: string,
  ) {
    if (editingId || event.detail > 1) {
      return
    }

    event.preventDefault()
    event.stopPropagation()
    event.currentTarget.setPointerCapture(event.pointerId)

    const pointerSelection = getCanvasItemPointerSelection({
      additive: isAdditivePointerInput(event),
      itemId: elementId,
      scene,
      selection,
    })
    const nextSelection = pointerSelection.nextSelection
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

  function handleStagePointerDown(event: ReactPointerEvent<HTMLDivElement>) {
    if (editingId || event.button !== 0) {
      return
    }

    event.currentTarget.setPointerCapture(event.pointerId)
    const additive = isAdditivePointerInput(event)
    const point = screenToWorld(event.nativeEvent)

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

      const nextDeck = updatePPTDeckSlide(interaction.startDeck, interaction.slideId, (slide) => ({
        ...slide,
        elements,
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
        selection: interaction.selection,
      })

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
        </div>
        <div className="ppt-toolbar-group">
          <button className="ppt-icon-button" onClick={addTextBox} title="Add text" type="button">
            <Type size={17} />
          </button>
          <button className="ppt-icon-button" data-ppt-insert-shape="rect" onClick={() => addShape('rect')} title="Add rectangle" type="button">
            <Square size={17} />
          </button>
          <button className="ppt-icon-button" data-ppt-insert-shape="ellipse" onClick={() => addShape('ellipse')} title="Add oval" type="button">
            <Circle size={17} />
          </button>
          <button className="ppt-icon-button" data-ppt-insert-shape="diamond" onClick={() => addShape('diamond')} title="Add diamond" type="button">
            <Diamond size={17} />
          </button>
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
        data-grid={showGrid ? 'true' : 'false'}
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
                key={element.id}
                selected={selection.includes(element.id)}
                onCommitText={commitText}
                onEdit={() => {
                  if (isPPTTextElement(element)) {
                    setEditingId(element.id)
                    setSelection([element.id])
                  }
                }}
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
            {marqueeBounds ? <Box className="ppt-marquee" bounds={marqueeBounds} /> : null}
            <Guides guides={snapGuides} scale={viewport.scale} />
          </div>
        </div>
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
        onElementLockToggle={toggleElementLocked}
        onElementNameChange={updateElementName}
        onElementRotationChange={updateElementRotation}
        onElementTextStyleChange={updateElementTextStyle}
        onElementVisibilityToggle={toggleElementVisibility}
        onLayerSelect={(elementId, additive) => {
          setSelection((current) =>
            getPPTLayerSelection(current, elementId, additive))
        }}
        onParagraphAlignChange={updateParagraphAlign}
        onShapeKindChange={updateShapeKind}
        onSlideBackgroundChange={updateSlideBackground}
        onShapeFillChange={updateShapeFill}
        onShapeStrokeChange={updateShapeStroke}
        onSlideNameChange={updateSlideName}
        onSlideNotesChange={updateSlideNotes}
      />
    </main>
  )
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
            className={element.kind === 'textBox' ? 'ppt-thumb-text' : 'ppt-thumb-shape'}
            data-shape={element.kind === 'shape' ? element.shape : undefined}
            key={element.id}
            style={{
              background: element.kind === 'shape' ? element.fill.color : '#cbd5e1',
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
  hovered,
  onCommitText,
  onEdit,
  onPointerDown,
  onPointerEnter,
  onPointerLeave,
  onStopEdit,
  selected,
}: {
  editing: boolean
  element: PPTElement
  hovered: boolean
  onCommitText: (elementId: string, text: string) => void
  onEdit: () => void
  onPointerDown: (event: ReactPointerEvent<HTMLDivElement>, elementId: string) => void
  onPointerEnter: () => void
  onPointerLeave: () => void
  onStopEdit: () => void
  selected: boolean
}) {
  const style = pptElementStyle(element)
  const text = element.kind === 'image' ? '' : readPPTText(element.textBody)
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
      data-kind={element.kind}
      data-locked={element.locked === true ? 'true' : 'false'}
      data-ppt-element={element.id}
      data-rotation={Math.round(element.geometry.rotation ?? 0)}
      data-selected={selected ? 'true' : 'false'}
      data-shape={element.kind === 'shape' ? element.shape : undefined}
      onDoubleClick={onEdit}
      onPointerDown={(event) => onPointerDown(event, element.id)}
      onPointerEnter={onPointerEnter}
      onPointerLeave={onPointerLeave}
      style={style}
    >
      {element.kind === 'image' ? (
        <img alt={element.alt} draggable={false} src={element.src} />
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
          {text}
        </div>
      )}
    </div>
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
  onElementTextStyleChange,
  onElementVisibilityToggle,
  onLayerSelect,
  onParagraphAlignChange,
  onShapeFillChange,
  onShapeKindChange,
  onShapeStrokeChange,
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
  onElementTextStyleChange: (
    elementId: string,
    field: keyof PPTTextStyle,
    value: string | number,
  ) => void
  onElementVisibilityToggle: (elementId: string) => void
  onLayerSelect: (elementId: string, additive: boolean) => void
  onParagraphAlignChange: (
    elementId: string,
    align: NonNullable<PPTParagraph['align']>,
  ) => void
  onShapeFillChange: (elementId: string, color: string) => void
  onShapeKindChange: (elementId: string, shape: PPTShapeKind) => void
  onShapeStrokeChange: (
    elementId: string,
    field: 'color' | 'width',
    value: string | number,
  ) => void
  onSlideBackgroundChange: (color: string) => void
  onSlideNameChange: (name: string) => void
  onSlideNotesChange: (notes: string) => void
  selection: string[]
  selectedElement: PPTElement | null
  slide: PPTSlide
}) {
  const textStyle = selectedElement && selectedElement.kind !== 'image'
    ? selectedElement.style
    : null
  const paragraphAlign = selectedElement && selectedElement.kind !== 'image'
    ? selectedElement.textBody?.paragraphs[0]?.align ?? 'left'
    : 'left'

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
            {selectedElement.kind !== 'image' ? (
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
                        onShapeStrokeChange(
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
                      onShapeStrokeChange(
                        selectedElement.id,
                        'width',
                        Number(event.target.value),
                      )}
                  />
                </label>
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

  if (element.kind === 'image') {
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

function isEditableTarget(target: EventTarget | null) {
  return target instanceof HTMLElement &&
    (target.isContentEditable ||
      ['INPUT', 'TEXTAREA', 'SELECT'].includes(target.tagName))
}

function getPPTElementParagraphAlign(element: PPTElement) {
  if (element.kind === 'image') {
    return undefined
  }

  return element.textBody?.paragraphs[0]?.align ?? 'left'
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

function getPointAngle(center: Point, point: Point) {
  return Math.atan2(point.y - center.y, point.x - center.x) * 180 / Math.PI
}

function normalizePPTElementRotation(rotation: number) {
  const normalized = ((rotation % 360) + 360) % 360

  return Math.abs(normalized) < 0.001 ? 0 : Number(normalized.toFixed(3))
}

function measurePPTElementAutoSize(element: PPTElement) {
  if (element.kind === 'image') {
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

function isPPTShapeKind(value: string): value is PPTShapeKind {
  return value === 'rect' || value === 'ellipse' || value === 'diamond'
}

function getPPTLayerSelection(
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

export default App
