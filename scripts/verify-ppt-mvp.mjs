import { spawn } from 'node:child_process'
import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import net from 'node:net'

const DEFAULT_APP_URL = 'http://127.0.0.1:5173/'
const EXTERNAL_APP_URL = process.env.APP_URL ?? null
let appUrl = EXTERNAL_APP_URL ?? DEFAULT_APP_URL
const CHROME_BIN =
  process.env.CHROME_BIN ??
  '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'
const CDP_COMMAND_TIMEOUT_MS = 10000
const PPT_SLIDE_WIDTH = 1280
const PPT_SLIDE_HEIGHT = 720
const PPT_TEST_IMAGE_WIDTH = 640
const PPT_TEST_IMAGE_HEIGHT = 360
const PPT_TIDY_GAP = 24
const PPT_OBJECT_ALT_TEXT = 'Revenue trend chart with highlighted AI cleanup'

const checks = []
const browserErrors = []
let devServer = null
let chrome = null
let chromeProfile = null

try {
  await ensureAppServer()
  const cdpPort = await getFreePort()
  await launchChrome(cdpPort)
  const page = await openPage(cdpPort, appUrl, {
    deviceScaleFactor: 1,
    height: 900,
    mobile: false,
    width: 1280,
  })

  await runFirstScreenScenario(page)
  await runTopToolbarRovingFocusScenario(page)
  await runTextEditingScenario(page)
  await runFindReplaceScenario(page)
  await runSpacingGuideScenario(page)
  await runSelectionAndDragScenario(page)
  await runMarqueeSelectionScenario(page)
  await runAltDragDuplicateScenario(page)
  await runAffordanceScenario(page)
  await runCommandSurfaceScenario(page)
  await runCrossSlideClipboardScenario(page)
  await runSelectSameTypeScenario(page)
  await runCommandPaletteScenario(page)
  await runShortcutHelpScenario(page)
  await runSlideMetadataScenario(page)
  await runThemeScenario(page)
  await runSlideTransitionScenario(page)
  await runObjectAnimationScenario(page)
  await runObjectOpacityScenario(page)
  await runObjectShadowScenario(page)
  await runObjectHyperlinkScenario(page)
  await runFitSelectionScenario(page)
  await runMinimapScenario(page)
  await runTidySelectionScenario(page)
  await runTextQuickFormatScenario(page)
  await runTextParagraphSpacingScenario(page)
  await runTextFontFamilyScenario(page)
  await runTextVerticalAlignScenario(page)
  await runTextFrameInsetScenario(page)
  await runViewAndShapeScenario(page)
  await runStickySectionScenario(page)
  await runLineAffordanceScenario(page)
  await runFreeformScenario(page)
  await runImageImportScenario(page)
  await runObjectAltTextScenario(page)
  await runTableImportScenario(page)
  await runTextPasteScenario(page)
  await runMediaImportScenario(page)
  await runCommentReviewScenario(page)
  await runFlipSelectionScenario(page)
  await runSelectionPaneScenario(page)
  await runTextOverflowScenario(page)
  await runPresentationScenario(page)
  await runExportScenario(page)
  await runAlignmentPopoverScenario(page)
  await runShapeMenuScenario(page)
  await runSlideManagementScenario(page)
  await runMobileScenario(cdpPort)

  await page.close()

  const failed = checks.filter((check) => !check.ok)
  const result = {
    browserErrors,
    checks,
    failed,
  }

  console.log(JSON.stringify(result, null, 2))
  process.exitCode = failed.length === 0 && browserErrors.length === 0 ? 0 : 1
} finally {
  if (chrome) await stopChrome()
  if (devServer) await stopDevServer()
  if (chromeProfile) {
    await rm(chromeProfile, {
      force: true,
      maxRetries: 5,
      recursive: true,
      retryDelay: 100,
    })
  }
}

async function runFirstScreenScenario(page) {
  const state = await page.eval(`(() => ({
    app: !!document.querySelector('[data-ppt-app]'),
    slideCount: document.querySelectorAll('.ppt-thumb').length,
    elementCount: document.querySelectorAll('[data-ppt-element]').length,
    selectedCount: document.querySelectorAll('[data-selected="true"]').length,
    hasPPTLabel: document.querySelector('.ppt-brand strong')?.textContent === 'PPT',
    hasRetouchShell: !!document.querySelector('.retouch-app, .slide-canvas'),
    hasStarterCopy: document.body.textContent.includes('React + TypeScript + Vite'),
    hasCanvasItemLeak: document.body.textContent.includes('CanvasItem'),
  }))()`)

  record('renders PPT app shell', state.app)
  record('renders multiple slides', state.slideCount >= 2, state)
  record('renders PPT elements', state.elementCount >= 5, state)
  record('has one selected element initially', state.selectedCount === 1, state)
  record('uses PPT label', state.hasPPTLabel, state)
  record('removes old retouch shell', !state.hasRetouchShell, state)
  record('removes Vite starter copy', !state.hasStarterCopy, state)
  record('does not expose CanvasItem as product text', !state.hasCanvasItemLeak, state)
}

async function runTopToolbarRovingFocusScenario(page) {
  const initial = await readPPTTopToolbarRovingFocusState(page)

  record(
    'exposes PPT top toolbar APG roving focus contract',
    initial.role === 'toolbar' &&
      initial.orientation === 'horizontal' &&
      initial.focusModel === 'roving-tabindex' &&
      initial.keyboardModel === 'arrow-home-end' &&
      initial.model === 'canvas-toolbar-roving-focus' &&
      initial.canvasEnabledItemCount === initial.enabledCount &&
      initial.canvasItemCount >= initial.enabledCount &&
      initial.enabledCount > 10,
    initial,
  )
  record(
    'keeps one enabled PPT toolbar tab stop',
    initial.zeroCount === 1 &&
      initial.minusCount === initial.enabledCount - 1 &&
      initial.disabledCount > 0 &&
      initial.disabledTabStopCount === 0,
    initial,
  )

  await page.eval(`(() => {
    const toolbar = document.querySelector('[data-ppt-toolbar]')
    const enabled = toolbar ? [...toolbar.querySelectorAll('.ppt-toolbar-group button:not(:disabled)')] : []
    const target = enabled.find((button) => button.tabIndex === 0)
    target?.focus()
  })()`)
  await delay(30)

  const beforeArrow = await readPPTTopToolbarRovingFocusState(page)

  await pressKey(page, {
    code: 'ArrowRight',
    key: 'ArrowRight',
    windowsVirtualKeyCode: 39,
  })
  await delay(30)
  const afterRight = await readPPTTopToolbarRovingFocusState(page)

  await pressKey(page, {
    code: 'ArrowLeft',
    key: 'ArrowLeft',
    windowsVirtualKeyCode: 37,
  })
  await delay(30)
  const afterLeft = await readPPTTopToolbarRovingFocusState(page)

  await pressKey(page, {
    code: 'End',
    key: 'End',
    windowsVirtualKeyCode: 35,
  })
  await delay(30)
  const afterEnd = await readPPTTopToolbarRovingFocusState(page)

  await pressKey(page, {
    code: 'ArrowRight',
    key: 'ArrowRight',
    windowsVirtualKeyCode: 39,
  })
  await delay(30)
  const afterWrap = await readPPTTopToolbarRovingFocusState(page)

  await pressKey(page, {
    code: 'Home',
    key: 'Home',
    windowsVirtualKeyCode: 36,
  })
  await delay(30)
  const afterHome = await readPPTTopToolbarRovingFocusState(page)

  record(
    'moves PPT top toolbar focus with Arrow/Home/End keys',
    beforeArrow.focusedIndex >= 0 &&
      afterRight.focusedIndex === (beforeArrow.focusedIndex + 1) % beforeArrow.enabledCount &&
      afterLeft.focusedIndex === beforeArrow.focusedIndex &&
      afterEnd.focusedIndex === beforeArrow.enabledCount - 1 &&
      afterWrap.focusedIndex === 0 &&
      afterHome.focusedIndex === 0,
    {
      afterEnd,
      afterHome,
      afterLeft,
      afterRight,
      afterWrap,
      beforeArrow,
    },
  )
}

async function readPPTTopToolbarRovingFocusState(page) {
  return page.eval(`(() => {
    const toolbar = document.querySelector('[data-ppt-toolbar]')
    const enabled = toolbar ? [...toolbar.querySelectorAll('.ppt-toolbar-group button:not(:disabled)')] : []
    const disabled = toolbar ? [...toolbar.querySelectorAll('.ppt-toolbar-group button:disabled')] : []
    const canvasItems = toolbar ? [...toolbar.querySelectorAll('[data-canvas-toolbar-item]')] : []
    const canvasEnabledItems = canvasItems.filter((button) =>
      button instanceof HTMLButtonElement
        ? !button.disabled && button.getAttribute('aria-disabled') !== 'true'
        : button.getAttribute('aria-disabled') !== 'true'
    )
    const focusedIndex = enabled.indexOf(document.activeElement)
    const activeIndex = enabled.findIndex((button) => button.tabIndex === 0)
    const tabStopIndex = enabled.findIndex((button) => button.tabIndex === 0)
    const itemTitles = enabled.map((button) =>
      button.getAttribute('title') ??
        button.getAttribute('aria-label') ??
        button.textContent.trim()
    )

    return {
      activeIndex,
      canvasEnabledItemCount: canvasEnabledItems.length,
      canvasItemCount: canvasItems.length,
      disabledCount: disabled.length,
      disabledTabStopCount: disabled.filter((button) => button.tabIndex >= 0).length,
      enabledCount: enabled.length,
      focusedIndex,
      focusedTitle: itemTitles[focusedIndex] ?? '',
      focusModel: toolbar?.getAttribute('data-ppt-toolbar-focus-model') ?? '',
      keyboardModel: toolbar?.getAttribute('data-ppt-toolbar-keyboard-model') ?? '',
      minusCount: enabled.filter((button) => button.tabIndex === -1).length,
      model: toolbar?.getAttribute('data-ppt-toolbar-model') ?? '',
      orientation: toolbar?.getAttribute('aria-orientation') ?? '',
      role: toolbar?.getAttribute('role') ?? '',
      tabStopIndex,
      tabStopTitle: itemTitles[tabStopIndex] ?? '',
      zeroCount: enabled.filter((button) => button.tabIndex === 0).length,
    }
  })()`)
}

async function runSelectionAndDragScenario(page) {
  const before = await page.eval(`(() => {
    const element = document.querySelector('[data-ppt-element="s1-card-1"]')
    const rect = element.getBoundingClientRect()
    return {
      left: element.style.left,
      x: rect.left + rect.width / 2,
      y: rect.top + rect.height / 2,
    }
  })()`)

  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mousePressed',
    x: before.x,
    y: before.y,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    type: 'mouseMoved',
    x: before.x + 5,
    y: before.y,
  })
  const duringDrag = await page.eval(`(() => ({
    alignmentGuideCount: document.querySelectorAll('[data-ppt-alignment-guide="true"]').length,
    spacingLabelCount: document.querySelectorAll('[data-ppt-spacing-guide-label="true"]').length,
  }))()`)
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    type: 'mouseMoved',
    x: before.x + 80,
    y: before.y + 34,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mouseReleased',
    x: before.x + 80,
    y: before.y + 34,
  })

  const after = await page.eval(`(() => {
    const element = document.querySelector('[data-ppt-element="s1-card-1"]')
    return {
      left: element.style.left,
      selected: element.getAttribute('data-selected'),
      undoEnabled: !document.querySelector('button[title="Undo"]').disabled,
    }
  })()`)

  record('selects dragged PPT element', after.selected === 'true', after)
  record(
    'renders canvas snap guides while dragging',
    duringDrag.alignmentGuideCount > 0 || duringDrag.spacingLabelCount > 0,
    duringDrag,
  )
  record('moves PPT element through canvas transform adapter', after.left !== before.left, {
    after,
    before,
  })
  record('records drag in undo history', after.undoEnabled, after)
}

async function runMarqueeSelectionScenario(page) {
  const replaceBox = await getPPTMarqueeDragBox(page, ['s1-card-1', 's1-card-2'])
  const before = await getPPTMarqueeState(page)

  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mousePressed',
    x: replaceBox.startX,
    y: replaceBox.startY,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    type: 'mouseMoved',
    x: replaceBox.endX,
    y: replaceBox.endY,
  })
  await delay(50)

  const duringReplace = await getPPTMarqueeState(page)

  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mouseReleased',
    x: replaceBox.endX,
    y: replaceBox.endY,
  })
  await delay(60)

  const afterReplace = await getPPTMarqueeState(page)

  record(
    'box-selects PPT objects through canvas marquee selection',
    duringReplace.marqueeModel === 'canvas-marquee-selection' &&
      duringReplace.marqueeActive === 'true' &&
      duringReplace.marqueeAdditive === 'false' &&
      duringReplace.marqueeHistory === 'none' &&
      duringReplace.marqueeCount === 1 &&
      duringReplace.marqueeWidth > 0 &&
      duringReplace.marqueeHeight > 0 &&
      duringReplace.marqueeSelection.includes('s1-card-1') &&
      duringReplace.marqueeSelection.includes('s1-card-2') &&
      afterReplace.selectedIds.includes('s1-card-1') &&
      afterReplace.selectedIds.includes('s1-card-2'),
    {
      afterReplace,
      before,
      duringReplace,
      replaceBox,
    },
  )

  await page.eval(`document.querySelector('button[title="Undo"]').click()`)
  await delay(80)

  const afterUndo = await getPPTMarqueeState(page)

  record(
    'keeps PPT marquee selection out of deck undo history',
    afterUndo.card1Left !== afterReplace.card1Left &&
      afterUndo.redoEnabled,
    {
      afterReplace,
      afterUndo,
    },
  )

  await page.eval(`document.querySelector('button[title="Redo"]').click()`)
  await delay(80)

  const afterRedo = await getPPTMarqueeState(page)

  record(
    'restores prior PPT deck history after marquee undo probe',
    afterRedo.card1Left === afterReplace.card1Left,
    {
      afterRedo,
      afterReplace,
    },
  )

  const titlePoint = await getElementCenter(page, 's1-title')

  await clickMouse(page, titlePoint.x, titlePoint.y, 1)
  await delay(50)

  const additiveBox = await getPPTMarqueeDragBox(page, ['s1-card-1'])

  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    modifiers: 8,
    type: 'mousePressed',
    x: additiveBox.startX,
    y: additiveBox.startY,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    modifiers: 8,
    type: 'mouseMoved',
    x: additiveBox.endX,
    y: additiveBox.endY,
  })
  await delay(50)

  const duringAdditive = await getPPTMarqueeState(page)

  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    modifiers: 8,
    type: 'mouseReleased',
    x: additiveBox.endX,
    y: additiveBox.endY,
  })
  await delay(60)

  const afterAdditive = await getPPTMarqueeState(page)

  record(
    'additively box-selects PPT objects with Shift marquee',
    duringAdditive.marqueeActive === 'true' &&
      duringAdditive.marqueeAdditive === 'true' &&
      duringAdditive.marqueeSelection.includes('s1-title') &&
      duringAdditive.marqueeSelection.includes('s1-card-1') &&
      afterAdditive.selectedIds.includes('s1-title') &&
      afterAdditive.selectedIds.includes('s1-card-1'),
    {
      additiveBox,
      afterAdditive,
      duringAdditive,
    },
  )

  await clickMouse(page, titlePoint.x, titlePoint.y, 2)
  await delay(80)

  const nativeBox = await getPPTMarqueeDragBox(page, ['s1-card-1'])
  const beforeNativeGuard = await getPPTMarqueeState(page)

  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mousePressed',
    x: nativeBox.startX,
    y: nativeBox.startY,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    type: 'mouseMoved',
    x: nativeBox.endX,
    y: nativeBox.endY,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mouseReleased',
    x: nativeBox.endX,
    y: nativeBox.endY,
  })
  await delay(60)

  const afterNativeGuard = await getPPTMarqueeState(page)

  record(
    'does not start PPT marquee while native text editing is active',
    beforeNativeGuard.editing &&
      afterNativeGuard.marqueeActive === 'false' &&
      afterNativeGuard.selectedIds.join(' ') === beforeNativeGuard.selectedIds.join(' '),
    {
      afterNativeGuard,
      beforeNativeGuard,
      nativeBox,
    },
  )

  await page.eval(`document.activeElement?.blur()`)
  await delay(50)
}

async function runSpacingGuideScenario(page) {
  const before = await page.eval(`(() => {
    const selected = document.querySelector('[data-ppt-element="s1-card-2"]')
    const left = document.querySelector('[data-ppt-element="s1-card-1"]')
    const right = document.querySelector('[data-ppt-element="s1-side-panel"]')
    const rect = selected.getBoundingClientRect()
    const selectedLeft = parseFloat(selected.style.left)
    const selectedWidth = parseFloat(selected.style.width)
    const leftRight = parseFloat(left.style.left) + parseFloat(left.style.width)
    const rightLeft = parseFloat(right.style.left)
    const targetLeft = (leftRight + rightLeft - selectedWidth) / 2
    const scale = rect.width / selectedWidth

    return {
      left: selected.style.left,
      targetX: rect.left + rect.width / 2 + (targetLeft - selectedLeft) * scale,
      x: rect.left + rect.width / 2,
      y: rect.top + rect.height / 2,
    }
  })()`)

  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mousePressed',
    x: before.x,
    y: before.y,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    type: 'mouseMoved',
    x: before.targetX,
    y: before.y,
  })
  await delay(50)

  const duringDrag = await page.eval(`(() => {
    const labels = [...document.querySelectorAll('[data-ppt-spacing-guide-label="true"]')]
    const segments = [...document.querySelectorAll('[data-ppt-spacing-guide="true"]')]

    return {
      spacingGaps: labels.map((label) => label.getAttribute('data-ppt-spacing-guide-label-gap') ?? ''),
      spacingLabelCount: labels.length,
      spacingOrientations: labels.map((label) => label.getAttribute('data-ppt-spacing-guide-label-orientation') ?? ''),
      spacingSegmentCount: segments.length,
      spacingSegmentGaps: segments.map((segment) => segment.getAttribute('data-ppt-spacing-guide-gap') ?? ''),
      spacingSegmentOrientations: segments.map((segment) => segment.getAttribute('data-ppt-spacing-guide-orientation') ?? ''),
    }
  })()`)

  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    type: 'mouseMoved',
    x: before.x,
    y: before.y,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mouseReleased',
    x: before.x,
    y: before.y,
  })

  const afterRelease = await page.eval(`(() => {
    const element = document.querySelector('[data-ppt-element="s1-card-2"]')

    return {
      left: element.style.left,
      selected: element.getAttribute('data-selected'),
      undoEnabled: !document.querySelector('button[title="Undo"]').disabled,
    }
  })()`)

  if (afterRelease.left !== before.left) {
    await page.eval(`document.querySelector('button[title="Undo"]').click()`)
    await delay(50)
  }

  const afterRestore = await page.eval(`(() => {
    const element = document.querySelector('[data-ppt-element="s1-card-2"]')

    return {
      left: element.style.left,
      selected: element.getAttribute('data-selected'),
    }
  })()`)

  record(
    'renders PPT spacing guide label metadata while dragging',
    duringDrag.spacingLabelCount > 0 &&
      duringDrag.spacingSegmentCount > 0 &&
      duringDrag.spacingGaps.every(Boolean) &&
      duringDrag.spacingOrientations.includes('horizontal') &&
      duringDrag.spacingSegmentGaps.every(Boolean) &&
      duringDrag.spacingSegmentOrientations.includes('horizontal'),
    duringDrag,
  )
  record('restores PPT spacing guide probe drag before continuing', afterRestore.left === before.left, {
    afterRelease,
    afterRestore,
    before,
  })
}

async function runAltDragDuplicateScenario(page) {
  const point = await getElementCenter(page, 's1-card-1')
  await clickMouse(page, point.x, point.y, 1)
  await delay(50)

  const before = await page.eval(`(() => {
    const element = document.querySelector('[data-ppt-element="s1-card-1"]')

    return {
      count: document.querySelectorAll('[data-ppt-element]').length,
      left: parseFloat(element.style.left),
      selectedId: document.querySelector('[data-selected="true"]')?.getAttribute('data-ppt-element') ?? '',
      top: parseFloat(element.style.top),
    }
  })()`)

  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    modifiers: 1,
    type: 'mousePressed',
    x: point.x,
    y: point.y,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    modifiers: 1,
    type: 'mouseMoved',
    x: point.x + 96,
    y: point.y + 28,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    modifiers: 1,
    type: 'mouseReleased',
    x: point.x + 96,
    y: point.y + 28,
  })
  await delay(100)

  const afterDuplicate = await page.eval(`(() => {
    const original = document.querySelector('[data-ppt-element="s1-card-1"]')
    const selected = document.querySelector('[data-selected="true"]')

    return {
      count: document.querySelectorAll('[data-ppt-element]').length,
      originalLeft: parseFloat(original.style.left),
      originalTop: parseFloat(original.style.top),
      selectedId: selected?.getAttribute('data-ppt-element') ?? '',
      selectedLeft: parseFloat(selected?.style.left ?? '0'),
      selectedName: document.querySelector('[data-ppt-style-field="name"]')?.value ?? '',
      selectedTop: parseFloat(selected?.style.top ?? '0'),
      undoEnabled: !document.querySelector('button[title="Undo"]').disabled,
    }
  })()`)

  record('duplicates selected PPT object with Alt drag', afterDuplicate.count === before.count + 1 && afterDuplicate.selectedId !== before.selectedId && afterDuplicate.selectedName.includes('Copy') && afterDuplicate.originalLeft === before.left && afterDuplicate.originalTop === before.top && afterDuplicate.selectedLeft !== before.left && afterDuplicate.undoEnabled, {
    afterDuplicate,
    before,
  })

  await page.eval(`document.querySelector('button[title="Undo"]').click()`)
  await delay(100)

  const afterUndo = await page.eval(`(() => {
    const original = document.querySelector('[data-ppt-element="s1-card-1"]')

    return {
      count: document.querySelectorAll('[data-ppt-element]').length,
      originalLeft: parseFloat(original.style.left),
      originalTop: parseFloat(original.style.top),
      redoEnabled: !document.querySelector('button[title="Redo"]').disabled,
    }
  })()`)

  record('undoes PPT Alt drag duplicate as one history step', afterUndo.count === before.count && afterUndo.originalLeft === before.left && afterUndo.originalTop === before.top && afterUndo.redoEnabled, {
    afterUndo,
    before,
  })

  const lockedPoint = await getElementCenter(page, 's1-card-1')
  await clickMouse(page, lockedPoint.x, lockedPoint.y, 1)
  await delay(50)
  await page.eval(`document.querySelector('[data-ppt-command="lock-selection"]').click()`)
  await delay(80)

  const beforeLockedDrag = await page.eval(`(() => ({
    count: document.querySelectorAll('[data-ppt-element]').length,
    locked: document.querySelector('[data-selected="true"]')?.getAttribute('data-locked') ?? '',
  }))()`)

  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    modifiers: 1,
    type: 'mousePressed',
    x: lockedPoint.x,
    y: lockedPoint.y,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    modifiers: 1,
    type: 'mouseMoved',
    x: lockedPoint.x + 80,
    y: lockedPoint.y + 20,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    modifiers: 1,
    type: 'mouseReleased',
    x: lockedPoint.x + 80,
    y: lockedPoint.y + 20,
  })
  await delay(80)

  const afterLockedDrag = await page.eval(`(() => ({
    count: document.querySelectorAll('[data-ppt-element]').length,
    locked: document.querySelector('[data-selected="true"]')?.getAttribute('data-locked') ?? '',
  }))()`)

  record('does not Alt drag duplicate locked PPT objects', beforeLockedDrag.locked === 'true' && afterLockedDrag.count === beforeLockedDrag.count && afterLockedDrag.locked === 'true', {
    afterLockedDrag,
    beforeLockedDrag,
  })

  await page.eval(`document.querySelector('[data-ppt-command="unlock-all"]').click()`)
  await delay(80)
  const restoredPoint = await getElementCenter(page, 's1-card-1')
  await clickMouse(page, restoredPoint.x, restoredPoint.y, 1)
  await delay(50)
}

async function runTextEditingScenario(page) {
  const titlePoint = await getElementCenter(page, 's1-title')

  await clickMouse(page, titlePoint.x, titlePoint.y, 2)
  await delay(50)
  await selectEditableContents(page, 's1-title')
  await page.send('Input.insertText', { text: 'Edited title' })
  await page.eval(`document.querySelector('[data-ppt-element="s1-title"] .ppt-element-editor').blur()`)
  await delay(50)

  const afterCommit = await page.eval(`(() => ({
    text: document.querySelector('[data-ppt-element="s1-title"]')?.textContent ?? '',
    undoEnabled: !document.querySelector('button[title="Undo"]').disabled,
  }))()`)

  record('commits inline PPT text edit on blur', afterCommit.text.includes('Edited title') && afterCommit.undoEnabled, afterCommit)

  const summaryPoint = await getElementCenter(page, 's1-summary')
  const beforeCancel = await page.eval(`document.querySelector('[data-ppt-element="s1-summary"]')?.textContent ?? ''`)

  await clickMouse(page, summaryPoint.x, summaryPoint.y, 2)
  await delay(50)
  await selectEditableContents(page, 's1-summary')
  await page.send('Input.insertText', { text: 'Should cancel' })
  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)

  const afterCancel = await page.eval(`document.querySelector('[data-ppt-element="s1-summary"]')?.textContent ?? ''`)

  record('cancels inline PPT text edit with Escape', afterCancel === beforeCancel, {
    afterCancel,
    beforeCancel,
  })

  await clickMouse(page, summaryPoint.x, summaryPoint.y, 2)
  await delay(50)
  await selectEditableContents(page, 's1-summary')

  const beforeInlinePaste = await getPPTInlineEditState(page)
  const inlinePasteText = 'Inline pasted text'

  await page.eval(`((text) => {
    const editor = document.querySelector('[data-ppt-element="s1-summary"] .ppt-element-editor')
    const dataTransfer = new DataTransfer()

    dataTransfer.setData('text/plain', text)
    editor?.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })(${JSON.stringify(inlinePasteText)})`)
  await delay(50)

  const afterInlinePaste = await getPPTInlineEditState(page)

  record(
    'pastes text inside PPT inline editor through canvas inline edit DOM contract',
    beforeInlinePaste.editorActive &&
      afterInlinePaste.editorActive &&
      afterInlinePaste.inlineEditModel === 'canvas-inline-edit-dom' &&
      afterInlinePaste.editorInlineEditModel === 'canvas-inline-edit-dom' &&
      afterInlinePaste.inlineEditElement === 's1-summary' &&
      afterInlinePaste.inlineEditPasteText === inlinePasteText &&
      afterInlinePaste.selectedText.includes(inlinePasteText) &&
      afterInlinePaste.elementCount === beforeInlinePaste.elementCount &&
      afterInlinePaste.textPasteSelection === beforeInlinePaste.textPasteSelection &&
      afterInlinePaste.mediaImportUrl === beforeInlinePaste.mediaImportUrl,
    {
      afterInlinePaste,
      beforeInlinePaste,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(50)

  const afterInlineUndoIntent = await getPPTInlineEditState(page)

  record(
    'records PPT inline editor undo shortcut as canvas inline edit history intent',
    afterInlineUndoIntent.inlineEditModel === 'canvas-inline-edit-dom' &&
      afterInlineUndoIntent.inlineEditElement === 's1-summary' &&
      afterInlineUndoIntent.inlineEditHistoryDirection === 'undo' &&
      afterInlineUndoIntent.elementCount === beforeInlinePaste.elementCount,
    afterInlineUndoIntent,
  )

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(50)

  const afterInlineRedoIntent = await getPPTInlineEditState(page)

  record(
    'records PPT inline editor redo shortcut as canvas inline edit history intent',
    afterInlineRedoIntent.inlineEditModel === 'canvas-inline-edit-dom' &&
      afterInlineRedoIntent.inlineEditElement === 's1-summary' &&
      afterInlineRedoIntent.inlineEditHistoryDirection === 'redo' &&
      afterInlineRedoIntent.elementCount === beforeInlinePaste.elementCount,
    afterInlineRedoIntent,
  )

  await pressKey(page, {
    code: 'Enter',
    key: 'Enter',
    windowsVirtualKeyCode: 13,
  })
  await delay(50)

  const afterInlineLineBreakIntent = await getPPTInlineEditState(page)

  record(
    'records PPT inline editor line break input through canvas inline edit DOM contract',
    afterInlineLineBreakIntent.inlineEditModel === 'canvas-inline-edit-dom' &&
      afterInlineLineBreakIntent.inlineEditElement === 's1-summary' &&
      afterInlineLineBreakIntent.inlineEditInputType === 'insertParagraph' &&
      afterInlineLineBreakIntent.inlineEditLineBreak === 'true',
    afterInlineLineBreakIntent,
  )

  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)

  const afterInlineCancel = await page.eval(`document.querySelector('[data-ppt-element="s1-summary"]')?.textContent ?? ''`)

  record('keeps PPT inline editor paste probe out of committed deck text on Escape', afterInlineCancel === beforeCancel, {
    afterInlineCancel,
    beforeCancel,
  })
}

async function runFindReplaceScenario(page) {
  await pressKey(page, {
    code: 'KeyF',
    key: 'f',
    modifiers: 2,
    windowsVirtualKeyCode: 70,
  })
  await waitUntil(
    () => page.eval(`!!document.querySelector('[data-ppt-find-query]')`),
    'Timed out waiting for PPT find query input',
  )
  await page.eval(`document.querySelector('[data-ppt-find-query]')?.focus()`)
  await delay(30)
  await page.send('Input.insertText', { text: 'PPT' })
  await delay(100)

  const afterFindMany = await page.eval(`(() => {
    const selected = document.querySelector('[data-selected="true"]')

    return {
      activeSlide: document.querySelector('.ppt-slide')?.getAttribute('data-ppt-slide') ?? '',
      count: document.querySelector('[data-ppt-find-count]')?.textContent ?? '',
      findActive: selected?.getAttribute('data-ppt-find-active') ?? '',
      selectedId: selected?.getAttribute('data-ppt-element') ?? '',
      stripOpen: !!document.querySelector('[data-ppt-find-strip]'),
      text: selected?.textContent ?? '',
    }
  })()`)

  record('finds multiple deck text results in slide order', afterFindMany.stripOpen && afterFindMany.count === '1/3' && afterFindMany.activeSlide === 'slide-1' && afterFindMany.selectedId === 's1-card-2', afterFindMany)

  await pressKey(page, {
    code: 'Enter',
    key: 'Enter',
    windowsVirtualKeyCode: 13,
  })
  await delay(100)

  const afterFindNext = await page.eval(`(() => {
    const selected = document.querySelector('[data-selected="true"]')

    return {
      activeSlide: document.querySelector('.ppt-slide')?.getAttribute('data-ppt-slide') ?? '',
      count: document.querySelector('[data-ppt-find-count]')?.textContent ?? '',
      selectedId: selected?.getAttribute('data-ppt-element') ?? '',
    }
  })()`)

  record('moves to next PPT find result with Enter', afterFindNext.count === '2/3' && afterFindNext.activeSlide === 'slide-2' && afterFindNext.selectedId === 's2-title', afterFindNext)

  await pressKey(page, {
    code: 'Enter',
    key: 'Enter',
    modifiers: 8,
    windowsVirtualKeyCode: 13,
  })
  await delay(100)

  const afterFindPrevious = await page.eval(`(() => {
    const selected = document.querySelector('[data-selected="true"]')

    return {
      activeSlide: document.querySelector('.ppt-slide')?.getAttribute('data-ppt-slide') ?? '',
      count: document.querySelector('[data-ppt-find-count]')?.textContent ?? '',
      selectedId: selected?.getAttribute('data-ppt-element') ?? '',
    }
  })()`)

  record('moves to previous PPT find result with Shift Enter', afterFindPrevious.count === '1/3' && afterFindPrevious.activeSlide === 'slide-1' && afterFindPrevious.selectedId === 's1-card-2', afterFindPrevious)

  await page.eval(`(() => {
    const input = document.querySelector('[data-ppt-find-query]')
    const valueSetter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set
    valueSetter.call(input, 'PPTX')
    input.dispatchEvent(new Event('input', { bubbles: true }))
    input.dispatchEvent(new Event('change', { bubbles: true }))
  })()`)
  await delay(100)

  const afterFind = await page.eval(`(() => {
    const selected = document.querySelector('[data-selected="true"]')

    return {
      activeSlide: document.querySelector('.ppt-slide')?.getAttribute('data-ppt-slide') ?? '',
      count: document.querySelector('[data-ppt-find-count]')?.textContent ?? '',
      findActive: selected?.getAttribute('data-ppt-find-active') ?? '',
      selectedId: selected?.getAttribute('data-ppt-element') ?? '',
      stripOpen: !!document.querySelector('[data-ppt-find-strip]'),
      text: selected?.textContent ?? '',
    }
  })()`)

  record('finds deck text across PPT slides', afterFind.stripOpen && afterFind.count === '1/1' && afterFind.activeSlide === 'slide-2' && afterFind.selectedId === 's2-title' && afterFind.findActive === 'true', afterFind)

  await page.eval(`(() => {
    const input = document.querySelector('[data-ppt-replace-query]')
    const valueSetter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set
    valueSetter.call(input, 'PPT-ready')
    input.dispatchEvent(new Event('input', { bubbles: true }))
    input.dispatchEvent(new Event('change', { bubbles: true }))
    document.querySelector('[data-ppt-find-replace]').click()
  })()`)
  await delay(100)

  const afterReplace = await page.eval(`(() => {
    const code = document.querySelector('.ppt-export-code')?.value ?? ''
    const selected = document.querySelector('[data-selected="true"]')

    return {
      count: document.querySelector('[data-ppt-find-count]')?.textContent ?? '',
      hasTextBodyRuns: code.includes('"textBody"') && code.includes('"paragraphs"') && code.includes('"runs"'),
      text: selected?.textContent ?? '',
      undoEnabled: !(document.querySelector('button[title="Undo"]')?.disabled ?? true),
    }
  })()`)

  record('replaces active PPT text match while preserving textBody structure', afterReplace.text.includes('PPT-ready later') && afterReplace.count === '0/0' && afterReplace.hasTextBodyRuns && afterReplace.undoEnabled, afterReplace)

  await page.eval(`document.querySelector('button[title="Undo"]').click()`)
  await delay(100)

  const afterUndo = await page.eval(`(() => {
    const selected = document.querySelector('[data-selected="true"]')

    return {
      count: document.querySelector('[data-ppt-find-count]')?.textContent ?? '',
      redoEnabled: !(document.querySelector('button[title="Redo"]')?.disabled ?? true),
      selectedId: selected?.getAttribute('data-ppt-element') ?? '',
      text: selected?.textContent ?? '',
    }
  })()`)

  record('undoes PPT find replacement', afterUndo.text.includes('PPTX later') && afterUndo.count === '1/1' && afterUndo.redoEnabled && afterUndo.selectedId === 's2-title', afterUndo)

  await page.eval(`document.querySelector('button[title="Redo"]').click()`)
  await delay(100)

  const afterRedo = await page.eval(`(() => {
    const selected = document.querySelector('[data-selected="true"]')

    return {
      count: document.querySelector('[data-ppt-find-count]')?.textContent ?? '',
      text: selected?.textContent ?? '',
    }
  })()`)

  record('redoes PPT find replacement', afterRedo.text.includes('PPT-ready later') && afterRedo.count === '0/0', afterRedo)

  await page.eval(`document.querySelector('[data-ppt-find-close]').click()`)
  await delay(50)

  const titlePoint = await getElementCenter(page, 's2-title')
  await clickMouse(page, titlePoint.x, titlePoint.y, 2)
  await delay(50)
  await page.eval(`(() => {
    const editor = document.querySelector('[data-ppt-element="s2-title"] .ppt-element-editor')
    editor?.dispatchEvent(new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      code: 'KeyF',
      ctrlKey: true,
      key: 'f',
    }))
  })()`)
  await delay(50)

  const afterGuard = await page.eval(`(() => {
    const editor = document.querySelector('[data-ppt-element="s2-title"] .ppt-element-editor')

    return {
      editing: editor?.isContentEditable === true && document.activeElement === editor,
      stripOpen: !!document.querySelector('[data-ppt-find-strip]'),
    }
  })()`)

  record('does not open PPT find while native text editing is active', afterGuard.editing && !afterGuard.stripOpen, afterGuard)

  const beforeNativeShortcutGuard = await page.eval(`(() => ({
    creationTool: document.querySelector('.ppt-stage-shell')?.getAttribute('data-creation-tool') ?? '',
    drawingTool: document.querySelector('.ppt-stage-shell')?.getAttribute('data-ppt-drawing-tool') ?? '',
    eraserToolActive: document.querySelector('.ppt-stage-shell')?.getAttribute('data-ppt-eraser-tool-active') ?? '',
    laserToolActive: document.querySelector('.ppt-stage-shell')?.getAttribute('data-ppt-laser-tool-active') ?? '',
    laserTrailPointCount: document.querySelector('.ppt-stage-shell')?.getAttribute('data-ppt-laser-trail-point-count') ?? '',
    lineTool: document.querySelector('.ppt-stage-shell')?.getAttribute('data-line-tool') ?? '',
    locked: document.querySelector('[data-ppt-element="s2-title"]')?.getAttribute('data-locked') ?? '',
    order: [...document.querySelectorAll('[data-ppt-element]')]
      .map((element) => element.getAttribute('data-ppt-element')).join(' '),
    panToolActive: document.querySelector('.ppt-stage-shell')?.getAttribute('data-ppt-pan-tool-active') ?? '',
    temporaryPanActive: document.querySelector('.ppt-stage-shell')?.getAttribute('data-ppt-temporary-pan-active') ?? '',
    temporaryPanGesture: document.querySelector('.ppt-stage-shell')?.getAttribute('data-ppt-temporary-pan-gesture') ?? '',
    viewportTransform: document.querySelector('.ppt-stage-world')?.style.transform ?? '',
  }))()`)

  await page.eval(`(() => {
    const editor = document.querySelector('[data-ppt-element="s2-title"] .ppt-element-editor')
    editor?.dispatchEvent(new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      code: 'KeyL',
      ctrlKey: true,
      key: 'l',
    }))
    editor?.dispatchEvent(new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      code: 'BracketRight',
      ctrlKey: true,
      key: ']',
    }))
    editor?.dispatchEvent(new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      code: 'Digit0',
      ctrlKey: true,
      key: '0',
    }))
    editor?.dispatchEvent(new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      code: 'Digit1',
      key: '1',
    }))
    editor?.dispatchEvent(new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      code: 'Space',
      key: ' ',
    }))
    editor?.dispatchEvent(new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      code: 'KeyH',
      key: 'h',
    }))
    editor?.dispatchEvent(new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      code: 'KeyP',
      key: 'p',
    }))
    editor?.dispatchEvent(new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      code: 'KeyM',
      key: 'm',
    }))
    editor?.dispatchEvent(new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      code: 'KeyM',
      key: 'M',
      shiftKey: true,
    }))
    editor?.dispatchEvent(new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      code: 'KeyE',
      key: 'e',
    }))
    editor?.dispatchEvent(new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      code: 'KeyS',
      key: 's',
    }))
    editor?.dispatchEvent(new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      code: 'KeyS',
      key: 'S',
      shiftKey: true,
    }))
    editor?.dispatchEvent(new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      code: 'KeyL',
      key: 'l',
    }))
  })()`)
  await delay(50)

  const afterNativeShortcutGuard = await page.eval(`(() => ({
    creationTool: document.querySelector('.ppt-stage-shell')?.getAttribute('data-creation-tool') ?? '',
    drawingTool: document.querySelector('.ppt-stage-shell')?.getAttribute('data-ppt-drawing-tool') ?? '',
    editing: document.activeElement?.matches('[data-ppt-element="s2-title"] .ppt-element-editor') === true,
    eraserToolActive: document.querySelector('.ppt-stage-shell')?.getAttribute('data-ppt-eraser-tool-active') ?? '',
    laserToolActive: document.querySelector('.ppt-stage-shell')?.getAttribute('data-ppt-laser-tool-active') ?? '',
    laserTrailPointCount: document.querySelector('.ppt-stage-shell')?.getAttribute('data-ppt-laser-trail-point-count') ?? '',
    lineTool: document.querySelector('.ppt-stage-shell')?.getAttribute('data-line-tool') ?? '',
    locked: document.querySelector('[data-ppt-element="s2-title"]')?.getAttribute('data-locked') ?? '',
    order: [...document.querySelectorAll('[data-ppt-element]')]
      .map((element) => element.getAttribute('data-ppt-element')).join(' '),
    panToolActive: document.querySelector('.ppt-stage-shell')?.getAttribute('data-ppt-pan-tool-active') ?? '',
    temporaryPanActive: document.querySelector('.ppt-stage-shell')?.getAttribute('data-ppt-temporary-pan-active') ?? '',
    temporaryPanGesture: document.querySelector('.ppt-stage-shell')?.getAttribute('data-ppt-temporary-pan-gesture') ?? '',
    viewportTransform: document.querySelector('.ppt-stage-world')?.style.transform ?? '',
  }))()`)

  record('does not run PPT arrange lock viewport pan laser or drawing shortcuts while native text editing is active', afterNativeShortcutGuard.editing && afterNativeShortcutGuard.locked === beforeNativeShortcutGuard.locked && afterNativeShortcutGuard.order === beforeNativeShortcutGuard.order && afterNativeShortcutGuard.viewportTransform === beforeNativeShortcutGuard.viewportTransform && afterNativeShortcutGuard.panToolActive === beforeNativeShortcutGuard.panToolActive && afterNativeShortcutGuard.temporaryPanActive === beforeNativeShortcutGuard.temporaryPanActive && afterNativeShortcutGuard.temporaryPanGesture === beforeNativeShortcutGuard.temporaryPanGesture && afterNativeShortcutGuard.laserToolActive === beforeNativeShortcutGuard.laserToolActive && afterNativeShortcutGuard.laserTrailPointCount === beforeNativeShortcutGuard.laserTrailPointCount && afterNativeShortcutGuard.creationTool === beforeNativeShortcutGuard.creationTool && afterNativeShortcutGuard.lineTool === beforeNativeShortcutGuard.lineTool && afterNativeShortcutGuard.drawingTool === beforeNativeShortcutGuard.drawingTool && afterNativeShortcutGuard.eraserToolActive === beforeNativeShortcutGuard.eraserToolActive, {
    afterNativeShortcutGuard,
    beforeNativeShortcutGuard,
  })

  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)
  await page.eval(`document.querySelector('.ppt-thumb[aria-label="Open Overview"]')?.click()`)
  await delay(50)
}

async function runAffordanceScenario(page) {
  const initial = await page.eval(`(() => {
    const element = document.querySelector('[data-ppt-element="s1-card-1"]')
    const frameGuides = document.querySelector('[data-ppt-frame-guides]')

    return {
      commandCount: document.querySelectorAll('[data-ppt-command]').length,
      frameGuideCount: document.querySelectorAll('.ppt-frame-guide').length,
      frameGuideState: {
        columnCount: document.querySelectorAll('[data-ppt-frame-guide-column]').length,
        lineCount: Number(frameGuides?.getAttribute('data-ppt-frame-guide-lines') ?? 0),
        pointerEvents: frameGuides ? getComputedStyle(frameGuides).pointerEvents : '',
        regionCount: document.querySelectorAll('[data-ppt-frame-guide-region]').length,
        rulerCount: document.querySelectorAll('[data-ppt-frame-guide-kind="ruler"]').length,
        safeAreaCount: document.querySelectorAll('[data-ppt-frame-guide-kind="safe-area"]').length,
      },
      geometryInputCount: document.querySelectorAll('[data-ppt-geometry-field]').length,
      hasRotateHandle: !!document.querySelector('[data-ppt-rotate-handle]'),
      hasSizeCapsule: !!document.querySelector('.ppt-size-capsule'),
      width: parseFloat(element.style.width),
    }
  })()`)

  record('renders PPT alignment and distribution commands', initial.commandCount >= 8, initial)
  record('renders PPT frame guides for selected object', initial.frameGuideCount >= 10, initial)
  record('renders PPT slide-edit frame guide contract geometry', initial.frameGuideState.lineCount >= 10 && initial.frameGuideState.columnCount === 4 && initial.frameGuideState.regionCount === 1 && initial.frameGuideState.safeAreaCount >= 5 && initial.frameGuideState.rulerCount === 2 && initial.frameGuideState.pointerEvents === 'none', initial.frameGuideState)
  record('renders PPT geometry inspector fields', initial.geometryInputCount === 5, initial)
  record('renders PPT selection size capsule', initial.hasSizeCapsule, initial)
  record('renders PPT rotation handle', initial.hasRotateHandle, initial)

  const firstResizeHandle = await page.eval(`(() => {
    const rect = document.querySelector('button[aria-label="Resize e"]').getBoundingClientRect()

    return {
      x: rect.left + rect.width / 2,
      y: rect.top + rect.height / 2,
    }
  })()`)

  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mousePressed',
    x: firstResizeHandle.x,
    y: firstResizeHandle.y,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    type: 'mouseMoved',
    x: firstResizeHandle.x + 42,
    y: firstResizeHandle.y,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mouseReleased',
    x: firstResizeHandle.x + 42,
    y: firstResizeHandle.y,
  })
  await delay(50)

  const afterResize = await page.eval(`(() => {
    const element = document.querySelector('[data-ppt-element="s1-card-1"]')

    return {
      width: parseFloat(element.style.width),
    }
  })()`)

  record('resizes selected object through canvas transform adapter', afterResize.width > initial.width, {
    afterResize,
    initial,
  })

  const resizeHandle = await page.eval(`(() => {
    const rect = document.querySelector('button[aria-label="Resize e"]').getBoundingClientRect()

    return {
      x: rect.left + rect.width / 2,
      y: rect.top + rect.height / 2,
    }
  })()`)

  await clickMouse(page, resizeHandle.x, resizeHandle.y, 1)
  await clickMouse(page, resizeHandle.x, resizeHandle.y, 2)
  await delay(100)

  const afterAuto = await page.eval(`(() => {
    const element = document.querySelector('[data-ppt-element="s1-card-1"]')
    const stage = document.querySelector('.ppt-stage-shell')

    return {
      command: stage?.getAttribute('data-ppt-text-autofit-command') ?? '',
      commandHandle: stage?.getAttribute('data-ppt-text-autofit-command-handle') ?? '',
      commandObject: stage?.getAttribute('data-ppt-text-autofit-command-object') ?? '',
      commandSelection: stage?.getAttribute('data-ppt-text-autofit-command-selection') ?? '',
      commandSizeMode: stage?.getAttribute('data-ppt-text-autofit-command-size-mode') ?? '',
      commandSlide: stage?.getAttribute('data-ppt-text-autofit-command-slide') ?? '',
      commandType: stage?.getAttribute('data-ppt-text-autofit-command-type') ?? '',
      height: parseFloat(element.style.height),
      model: stage?.getAttribute('data-ppt-text-autofit-model') ?? '',
      resizeClickDouble: stage?.getAttribute('data-ppt-resize-handle-click-double') ?? '',
      resizeClickHandle: stage?.getAttribute('data-ppt-resize-handle-click-handle') ?? '',
      resizeClickId: stage?.getAttribute('data-ppt-resize-handle-click-id') ?? '',
      resizeClickModel: stage?.getAttribute('data-ppt-resize-handle-click-model') ?? '',
      resizeClickX: stage?.getAttribute('data-ppt-resize-handle-click-x') ?? '',
      resizeClickY: stage?.getAttribute('data-ppt-resize-handle-click-y') ?? '',
      selectedAutoFit: element?.getAttribute('data-ppt-text-autofit') ?? '',
      width: parseFloat(element.style.width),
    }
  })()`)

  record('auto-sizes selected object from resize handle double-click', afterAuto.width !== afterResize.width, {
    afterAuto,
    afterResize,
  })
  record(
    'routes PPT resize handle double-click through canvas pointer click memory',
    afterAuto.resizeClickDouble === 'true' &&
      afterAuto.resizeClickHandle === 'e' &&
      afterAuto.resizeClickId === 'slide-1:s1-card-1:e' &&
      afterAuto.resizeClickModel === 'canvas-pointer-click-memory' &&
      afterAuto.resizeClickX !== '' &&
      afterAuto.resizeClickY !== '',
    afterAuto,
  )
  record('routes PPT text auto-size double-click through slide-edit command effect', afterAuto.command === 'resize-text-box-to-fit' && afterAuto.commandHandle === 'e' && afterAuto.commandObject === 's1-card-1' && afterAuto.commandSelection === 's1-card-1' && afterAuto.commandSizeMode === 'resize-to-fit' && afterAuto.commandSlide === 'slide-1' && afterAuto.commandType === 'slide-command-effect' && afterAuto.model === 'slide-edit-text-box-auto-fit' && afterAuto.selectedAutoFit === 'resizeShapeToFitText', afterAuto)

  const aspectResizeHandle = await page.eval(`(() => {
    const rect = document.querySelector('button[aria-label="Resize e"]').getBoundingClientRect()

    return {
      x: rect.left + rect.width / 2,
      y: rect.top + rect.height / 2,
    }
  })()`)

  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    modifiers: 8,
    type: 'mousePressed',
    x: aspectResizeHandle.x,
    y: aspectResizeHandle.y,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    modifiers: 8,
    type: 'mouseMoved',
    x: aspectResizeHandle.x + 46,
    y: aspectResizeHandle.y,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    modifiers: 8,
    type: 'mouseReleased',
    x: aspectResizeHandle.x + 46,
    y: aspectResizeHandle.y,
  })
  await delay(50)

  const afterAspectResize = await page.eval(`(() => {
    const element = document.querySelector('[data-ppt-element="s1-card-1"]')

    return {
      height: parseFloat(element.style.height),
      ratio: parseFloat(element.style.width) / parseFloat(element.style.height),
      width: parseFloat(element.style.width),
    }
  })()`)
  const expectedAspectRatio = afterAuto.width / afterAuto.height

  record('preserves aspect ratio when Shift-resizing PPT object', afterAspectResize.width > afterAuto.width && afterAspectResize.height > afterAuto.height && Math.abs(afterAspectResize.ratio - expectedAspectRatio) < 0.02, {
    afterAspectResize,
    expectedAspectRatio,
  })

  const centerResizeBefore = await page.eval(`(() => {
    const element = document.querySelector('[data-ppt-element="s1-card-1"]')
    const shell = document.querySelector('.ppt-stage-shell')
    const left = parseFloat(element.style.left)
    const top = parseFloat(element.style.top)
    const width = parseFloat(element.style.width)
    const height = parseFloat(element.style.height)
    const rect = document.querySelector('button[aria-label="Resize e"]').getBoundingClientRect()

    return {
      centerX: left + width / 2,
      centerY: top + height / 2,
      handleX: rect.left + rect.width / 2,
      handleY: rect.top + rect.height / 2,
      height,
      left,
      modifierModel: shell?.getAttribute('data-ppt-resize-modifier-model') ?? '',
      preserveAspectModifier: shell?.getAttribute('data-ppt-resize-aspect-ratio-modifier') ?? '',
      resizeFromCenterModifier: shell?.getAttribute('data-ppt-resize-from-center-modifier') ?? '',
      top,
      width,
    }
  })()`)

  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    modifiers: 1,
    type: 'mousePressed',
    x: centerResizeBefore.handleX,
    y: centerResizeBefore.handleY,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    modifiers: 1,
    type: 'mouseMoved',
    x: centerResizeBefore.handleX + 34,
    y: centerResizeBefore.handleY,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    modifiers: 1,
    type: 'mouseReleased',
    x: centerResizeBefore.handleX + 34,
    y: centerResizeBefore.handleY,
  })
  await delay(50)

  const afterCenterResize = await page.eval(`(() => {
    const element = document.querySelector('[data-ppt-element="s1-card-1"]')
    const left = parseFloat(element.style.left)
    const top = parseFloat(element.style.top)
    const width = parseFloat(element.style.width)
    const height = parseFloat(element.style.height)

    return {
      centerX: left + width / 2,
      centerY: top + height / 2,
      height,
      left,
      top,
      width,
    }
  })()`)

  record('resizes selected object from center with Alt resize drag', centerResizeBefore.modifierModel === 'canvas-resize-pointer-modifiers' && centerResizeBefore.preserveAspectModifier === 'Shift' && centerResizeBefore.resizeFromCenterModifier === 'Alt' && afterCenterResize.width > centerResizeBefore.width && afterCenterResize.left < centerResizeBefore.left && nearlyEqual(afterCenterResize.centerX, centerResizeBefore.centerX, 0.75) && nearlyEqual(afterCenterResize.centerY, centerResizeBefore.centerY, 0.75), {
    afterCenterResize,
    centerResizeBefore,
  })

  const rotateHandle = await page.eval(`(() => {
    const rect = document.querySelector('[data-ppt-rotate-handle]').getBoundingClientRect()

    return {
      x: rect.left + rect.width / 2,
      y: rect.top + rect.height / 2,
    }
  })()`)

  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mousePressed',
    x: rotateHandle.x,
    y: rotateHandle.y,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    type: 'mouseMoved',
    x: rotateHandle.x + 52,
    y: rotateHandle.y - 22,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mouseReleased',
    x: rotateHandle.x + 52,
    y: rotateHandle.y - 22,
  })
  await delay(50)

  const afterRotateHandle = await page.eval(`(() => {
    const element = document.querySelector('[data-ppt-element="s1-card-1"]')

    return {
      rotation: Number(element.getAttribute('data-rotation')),
      transform: element.style.transform,
      undoEnabled: !document.querySelector('button[title="Undo"]').disabled,
    }
  })()`)

  record('rotates selected object from PPT rotation handle', Math.abs(afterRotateHandle.rotation) > 0 && afterRotateHandle.transform.includes('rotate(') && afterRotateHandle.undoEnabled, afterRotateHandle)

  await page.eval(`(() => {
    const rotation = document.querySelector('[data-ppt-geometry-field="rotation"]')
    const valueSetter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set
    valueSetter.call(rotation, '45')
    rotation.dispatchEvent(new Event('input', { bubbles: true }))
    rotation.dispatchEvent(new Event('change', { bubbles: true }))
  })()`)
  await delay(50)

  const afterRotationInput = await page.eval(`(() => {
    const element = document.querySelector('[data-ppt-element="s1-card-1"]')

    return {
      rotation: element.getAttribute('data-rotation'),
      transform: element.style.transform,
    }
  })()`)

  record('updates PPT object rotation from inspector', afterRotationInput.rotation === '45' && afterRotationInput.transform.includes('rotate(45deg)'), afterRotationInput)

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      objectTransform: {
        h: 138,
        rotation: 45,
        w: 360,
        x: 214,
        y: 118,
      },
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(80)

  const afterTransformPaste = await page.eval(`(() => {
    const element = document.querySelector('[data-ppt-element="s1-card-1"]')
    const stage = document.querySelector('.ppt-stage-shell')

    return {
      fieldH: document.querySelector('[data-ppt-geometry-field="h"]')?.value ?? '',
      fieldRotation: document.querySelector('[data-ppt-geometry-field="rotation"]')?.value ?? '',
      fieldW: document.querySelector('[data-ppt-geometry-field="w"]')?.value ?? '',
      fieldX: document.querySelector('[data-ppt-geometry-field="x"]')?.value ?? '',
      fieldY: document.querySelector('[data-ppt-geometry-field="y"]')?.value ?? '',
      height: parseFloat(element.style.height),
      importFields: stage?.getAttribute('data-ppt-object-transform-import-fields') ?? '',
      importFormat: stage?.getAttribute('data-ppt-object-transform-import-format') ?? '',
      importH: stage?.getAttribute('data-ppt-object-transform-import-h') ?? '',
      importJsonLength: stage?.getAttribute('data-ppt-object-transform-import-json-length') ?? '',
      importModel: stage?.getAttribute('data-ppt-object-transform-import-model') ?? '',
      importObjects: stage?.getAttribute('data-ppt-object-transform-import-objects') ?? '',
      importRotation: stage?.getAttribute('data-ppt-object-transform-import-rotation') ?? '',
      importTargets: stage?.getAttribute('data-ppt-object-transform-import-command-targets') ?? '',
      importW: stage?.getAttribute('data-ppt-object-transform-import-w') ?? '',
      importX: stage?.getAttribute('data-ppt-object-transform-import-x') ?? '',
      importY: stage?.getAttribute('data-ppt-object-transform-import-y') ?? '',
      left: parseFloat(element.style.left),
      rotation: element.getAttribute('data-rotation'),
      top: parseFloat(element.style.top),
      transform: element.style.transform,
      width: parseFloat(element.style.width),
    }
  })()`)

  record(
    'pastes JSON object transform into selected PPT object',
    afterTransformPaste.importModel === 'ppt-object-transform-import' &&
      afterTransformPaste.importFormat === 'application-json-ppt-object-transform' &&
      afterTransformPaste.importTargets === 's1-card-1' &&
      afterTransformPaste.importObjects === 's1-card-1' &&
      afterTransformPaste.importFields === 'x y w h rotation' &&
      afterTransformPaste.importX === '214' &&
      afterTransformPaste.importY === '118' &&
      afterTransformPaste.importW === '360' &&
      afterTransformPaste.importH === '138' &&
      afterTransformPaste.importRotation === '45' &&
      Number(afterTransformPaste.importJsonLength) > 60 &&
      afterTransformPaste.left === 214 &&
      afterTransformPaste.top === 118 &&
      afterTransformPaste.width === 360 &&
      afterTransformPaste.height === 138 &&
      afterTransformPaste.rotation === '45' &&
      afterTransformPaste.transform.includes('rotate(45deg)') &&
      afterTransformPaste.fieldX === '214' &&
      afterTransformPaste.fieldY === '118' &&
      afterTransformPaste.fieldW === '360' &&
      afterTransformPaste.fieldH === '138' &&
      afterTransformPaste.fieldRotation === '45',
    {
      afterRotationInput,
      afterTransformPaste,
    },
  )

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      height: 120,
      rotate: 30,
      width: 300,
      x: 244,
      y: 148,
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(80)

  const afterStandaloneTransformPaste = await page.eval(`(() => {
    const element = document.querySelector('[data-ppt-element="s1-card-1"]')
    const stage = document.querySelector('.ppt-stage-shell')

    return {
      fieldH: document.querySelector('[data-ppt-geometry-field="h"]')?.value ?? '',
      fieldRotation: document.querySelector('[data-ppt-geometry-field="rotation"]')?.value ?? '',
      fieldW: document.querySelector('[data-ppt-geometry-field="w"]')?.value ?? '',
      fieldX: document.querySelector('[data-ppt-geometry-field="x"]')?.value ?? '',
      fieldY: document.querySelector('[data-ppt-geometry-field="y"]')?.value ?? '',
      height: parseFloat(element.style.height),
      importFields: stage?.getAttribute('data-ppt-object-transform-import-fields') ?? '',
      importFormat: stage?.getAttribute('data-ppt-object-transform-import-format') ?? '',
      importH: stage?.getAttribute('data-ppt-object-transform-import-h') ?? '',
      importJsonLength: stage?.getAttribute('data-ppt-object-transform-import-json-length') ?? '',
      importModel: stage?.getAttribute('data-ppt-object-transform-import-model') ?? '',
      importObjects: stage?.getAttribute('data-ppt-object-transform-import-objects') ?? '',
      importRotation: stage?.getAttribute('data-ppt-object-transform-import-rotation') ?? '',
      importTargets: stage?.getAttribute('data-ppt-object-transform-import-command-targets') ?? '',
      importW: stage?.getAttribute('data-ppt-object-transform-import-w') ?? '',
      importX: stage?.getAttribute('data-ppt-object-transform-import-x') ?? '',
      importY: stage?.getAttribute('data-ppt-object-transform-import-y') ?? '',
      left: parseFloat(element.style.left),
      rotation: element.getAttribute('data-rotation'),
      top: parseFloat(element.style.top),
      transform: element.style.transform,
      width: parseFloat(element.style.width),
    }
  })()`)

  record(
    'pastes standalone JSON object transform into selected PPT object',
    afterStandaloneTransformPaste.importModel === 'ppt-object-transform-import' &&
      afterStandaloneTransformPaste.importFormat === 'application-json-ppt-object-transform' &&
      afterStandaloneTransformPaste.importTargets === 's1-card-1' &&
      afterStandaloneTransformPaste.importObjects === 's1-card-1' &&
      afterStandaloneTransformPaste.importFields === 'x y w h rotation' &&
      afterStandaloneTransformPaste.importX === '244' &&
      afterStandaloneTransformPaste.importY === '148' &&
      afterStandaloneTransformPaste.importW === '300' &&
      afterStandaloneTransformPaste.importH === '120' &&
      afterStandaloneTransformPaste.importRotation === '30' &&
      Number(afterStandaloneTransformPaste.importJsonLength) > 50 &&
      afterStandaloneTransformPaste.left === 244 &&
      afterStandaloneTransformPaste.top === 148 &&
      afterStandaloneTransformPaste.width === 300 &&
      afterStandaloneTransformPaste.height === 120 &&
      afterStandaloneTransformPaste.rotation === '30' &&
      afterStandaloneTransformPaste.transform.includes('rotate(30deg)') &&
      afterStandaloneTransformPaste.fieldX === '244' &&
      afterStandaloneTransformPaste.fieldY === '148' &&
      afterStandaloneTransformPaste.fieldW === '300' &&
      afterStandaloneTransformPaste.fieldH === '120' &&
      afterStandaloneTransformPaste.fieldRotation === '30',
    {
      afterStandaloneTransformPaste,
      afterTransformPaste,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterStandaloneTransformUndo = await page.eval(`(() => {
    const element = document.querySelector('[data-ppt-element="s1-card-1"]')

    return {
      height: parseFloat(element.style.height),
      left: parseFloat(element.style.left),
      rotation: element.getAttribute('data-rotation'),
      top: parseFloat(element.style.top),
      transform: element.style.transform,
      width: parseFloat(element.style.width),
    }
  })()`)

  record(
    'undoes standalone JSON object transform as one history step',
    afterStandaloneTransformUndo.left === afterTransformPaste.left &&
      afterStandaloneTransformUndo.top === afterTransformPaste.top &&
      afterStandaloneTransformUndo.width === afterTransformPaste.width &&
      afterStandaloneTransformUndo.height === afterTransformPaste.height &&
      afterStandaloneTransformUndo.rotation === afterTransformPaste.rotation &&
      afterStandaloneTransformUndo.transform.includes('rotate(45deg)'),
    {
      afterStandaloneTransformPaste,
      afterStandaloneTransformUndo,
      afterTransformPaste,
    },
  )

  await page.eval(`document.querySelector('[data-ppt-command="align-center-x"]').click()`)
  await delay(50)

  const afterAlign = await page.eval(`(() => {
    const element = document.querySelector('[data-ppt-element="s1-card-1"]')
    const left = parseFloat(element.style.left)
    const width = parseFloat(element.style.width)

    return {
      expectedLeft: (1280 - width) / 2,
      left,
      width,
    }
  })()`)

  record('aligns selected object with PPT command', Math.abs(afterAlign.left - afterAlign.expectedLeft) <= 1, afterAlign)

  await pressKey(page, {
    code: 'ArrowRight',
    key: 'ArrowRight',
    windowsVirtualKeyCode: 39,
  })
  await delay(50)

  const afterNudge = await page.eval(`(() => {
    const shell = document.querySelector('.ppt-stage-shell')
    const element = document.querySelector('[data-ppt-element="s1-card-1"]')

    return {
      elementCount: document.querySelectorAll('[data-ppt-element]').length,
      nudgeEnabled: shell?.getAttribute('data-ppt-keyboard-nudge-enabled') ?? '',
      nudgeIntent: shell?.getAttribute('data-ppt-keyboard-nudge-intent') ?? '',
      nudgeKeys: shell?.getAttribute('data-ppt-keyboard-nudge-keys') ?? '',
      nudgeLargeStep: shell?.getAttribute('data-ppt-keyboard-nudge-large-step') ?? '',
      nudgeModel: shell?.getAttribute('data-ppt-keyboard-nudge-model') ?? '',
      nudgeStep: shell?.getAttribute('data-ppt-keyboard-nudge-step') ?? '',
      left: parseFloat(element.style.left),
      selectedCount: document.querySelectorAll('[data-selected="true"]').length,
    }
  })()`)

  record('exposes PPT canvas keyboard nudge shortcut metadata', afterNudge.nudgeEnabled === 'true' && afterNudge.nudgeModel === 'canvas-keyboard-nudge-shortcuts' && afterNudge.nudgeIntent === 'canvas-keyboard-nudge-shortcut-intent' && afterNudge.nudgeStep === '1' && afterNudge.nudgeLargeStep === '10' && afterNudge.nudgeKeys.includes('ArrowRight') && afterNudge.nudgeKeys.includes('Shift+ArrowRight'), {
    afterNudge,
  })

  record('nudges selected object with arrow key', afterNudge.left === afterAlign.left + 1, {
    afterAlign,
    afterNudge,
  })

  await pressKey(page, {
    code: 'ArrowRight',
    key: 'ArrowRight',
    modifiers: 8,
    windowsVirtualKeyCode: 39,
  })
  await delay(50)

  const afterLargeNudge = await page.eval(`(() => {
    const element = document.querySelector('[data-ppt-element="s1-card-1"]')

    return {
      hasElement: !!element,
      left: parseFloat(element?.style.left ?? 'NaN'),
      selectedCount: document.querySelectorAll('[data-selected="true"]').length,
    }
  })()`)

  record('large-nudges selected object with Shift Arrow key', nearlyEqual(afterLargeNudge.left, afterNudge.left + 10, 0.001) && afterLargeNudge.selectedCount === afterNudge.selectedCount, {
    afterLargeNudge,
    afterNudge,
  })

  const multiSelectTargets = await page.eval(`(() => {
    const ids = ['s1-card-2', 's1-side-panel']

    return ids.map((id) => {
      const rect = document.querySelector('[data-ppt-element="' + id + '"]').getBoundingClientRect()

      return {
        id,
        x: rect.left + rect.width / 2,
        y: rect.top + rect.height / 2,
      }
    })
  })()`)

  for (const target of multiSelectTargets) {
    await clickMouse(page, target.x, target.y, 1, 8)
  }
  await delay(50)

  const beforeDistribute = await page.eval(`(() => ({
    left: parseFloat(document.querySelector('[data-ppt-element="s1-card-1"]').style.left),
    selectedCount: document.querySelectorAll('[data-selected="true"]').length,
  }))()`)

  await page.eval(`document.querySelector('[data-ppt-command="distribute-horizontal"]').click()`)
  await delay(50)

  const afterDistribute = await page.eval(`(() => {
    const card = document.querySelector('[data-ppt-element="s1-card-1"]')
    const rect = card.getBoundingClientRect()
    const selected = [...document.querySelectorAll('[data-selected="true"]')]
      .map((element) => ({
        id: element.getAttribute('data-ppt-element'),
        left: parseFloat(element.style.left),
        right: parseFloat(element.style.left) + parseFloat(element.style.width),
      }))
      .sort((left, right) => left.left - right.left)
    const gaps = selected.slice(1).map((element, index) => element.left - selected[index].right)
    const gapDelta = Math.max(...gaps) - Math.min(...gaps)

    return {
      gapDelta,
      gaps,
      left: parseFloat(card.style.left),
      selectedCount: document.querySelectorAll('[data-selected="true"]').length,
      selectedIds: selected.map((element) => element.id),
      x: rect.left + rect.width / 2,
      y: rect.top + rect.height / 2,
    }
  })()`)

  record('distributes multi-selected objects with canvas command adapter', beforeDistribute.selectedCount === 3 && afterDistribute.selectedCount === 3 && afterDistribute.gaps.length === 2 && afterDistribute.gapDelta <= 1, {
    afterDistribute,
    beforeDistribute,
  })

  await page.eval(`document.querySelector('[data-ppt-command="group"]').click()`)
  await delay(50)

  const afterGroup = await page.eval(`(() => {
    const selected = [...document.querySelectorAll('[data-selected="true"]')]
    const groupIds = selected.map((element) => element.getAttribute('data-group-id'))
    const groupRow = document.querySelector('[data-ppt-layer-pane-row-type="group"]')
    const groupRowId = groupRow?.getAttribute('data-ppt-layer-pane-row') ?? ''
    const childRows = [...document.querySelectorAll(\`[data-ppt-layer-pane-parent-object-id="\${groupRowId}"]\`)]

    return {
      childLevels: childRows.map((row) => row.getAttribute('aria-level')),
      childRowCount: childRows.length,
      groupExpanded: groupRow?.getAttribute('aria-expanded') ?? '',
      groupedLayerCount: document.querySelectorAll('[data-ppt-layer-row][data-grouped="true"]').length,
      groupIds,
      groupRowCount: document.querySelectorAll('[data-ppt-layer-pane-row-type="group"]').length,
      groupRowId,
      groupRowLevel: groupRow?.getAttribute('aria-level') ?? '',
      groupRowType: groupRow?.getAttribute('data-ppt-layer-pane-row-type') ?? '',
      selectedCount: selected.length,
      ungroupDisabled: document.querySelector('[data-ppt-command="ungroup"]').disabled,
      uniqueGroupCount: new Set(groupIds).size,
    }
  })()`)

  record('groups multi-selected PPT objects with canvas command adapter', afterGroup.selectedCount === 3 && afterGroup.groupIds.every(Boolean) && afterGroup.uniqueGroupCount === 1 && afterGroup.groupedLayerCount >= 4 && !afterGroup.ungroupDisabled, afterGroup)
  record('renders grouped PPT objects as layer pane tree rows', afterGroup.groupRowCount === 1 && afterGroup.groupRowId.length > 0 && afterGroup.groupRowType === 'group' && afterGroup.groupExpanded === 'true' && afterGroup.groupRowLevel === '1' && afterGroup.childRowCount === 3 && afterGroup.childLevels.every((level) => level === '2'), afterGroup)

  await page.eval(`document.querySelector('[data-ppt-layer-pane-row-type="group"]')?.focus()`)
  await delay(30)
  await pressKey(page, {
    code: 'F2',
    key: 'F2',
    windowsVirtualKeyCode: 113,
  })
  await delay(50)
  const afterGroupRenameGuard = await page.eval(`(() => {
    const groupRow = document.querySelector('[data-ppt-layer-pane-row-type="group"]')

    return {
      focusedRowId: document.activeElement?.closest('[data-ppt-layer-pane-row]')?.getAttribute('data-ppt-layer-pane-row') ?? '',
      groupDraggable: groupRow?.getAttribute('data-ppt-layer-pane-draggable') ?? '',
      groupRenamable: groupRow?.getAttribute('data-ppt-layer-pane-renamable') ?? '',
      groupRowId: groupRow?.getAttribute('data-ppt-layer-pane-row') ?? '',
      renameInputCount: document.querySelectorAll('[data-ppt-layer-pane-rename-input]').length,
    }
  })()`)

  record(
    'does not open PPT layer pane rename for non-renamable group rows',
    afterGroupRenameGuard.groupDraggable === 'true' &&
      afterGroupRenameGuard.groupRenamable === 'false' &&
      afterGroupRenameGuard.renameInputCount === 0 &&
      afterGroupRenameGuard.focusedRowId === afterGroupRenameGuard.groupRowId,
    afterGroupRenameGuard,
  )

  const groupLayerKeyboardReorder = await reorderPPTLayerPaneGroupWithKeyboard(page)

  record(
    'reorders PPT group row with Alt Arrow keyboard intent',
    groupLayerKeyboardReorder.ok &&
      groupLayerKeyboardReorder.before.groupDraggable === 'true' &&
      groupLayerKeyboardReorder.before.childStageIds.length >= 2 &&
      groupLayerKeyboardReorder.after.groupMemberContiguous &&
      groupLayerKeyboardReorder.after.childStageIds.join(' ') === groupLayerKeyboardReorder.before.childStageIds.join(' ') &&
      groupLayerKeyboardReorder.after.groupStartIndex !== groupLayerKeyboardReorder.before.groupStartIndex &&
      groupLayerKeyboardReorder.after.selectedStageIds.join(' ') === groupLayerKeyboardReorder.before.childStageIds.join(' ') &&
      groupLayerKeyboardReorder.focusedRowId === groupLayerKeyboardReorder.groupRowId,
    groupLayerKeyboardReorder,
  )

  const beforeGroupLayerReorder = await readPPTLayerPaneGroupReorderState(page)
  const groupLayerReorderDrag = await dragPPTLayerPaneGroupRow(page)
  await delay(80)
  const afterGroupLayerReorder = await readPPTLayerPaneGroupReorderState(page)

  record(
    'reorders PPT group row as a layer pane member block',
    beforeGroupLayerReorder.groupDraggable === 'true' &&
      groupLayerReorderDrag.ok &&
      groupLayerReorderDrag.indicatorModel === 'slide-edit-layer-pane-drop-indicator' &&
      ['after', 'before'].includes(groupLayerReorderDrag.indicatorPlacement) &&
      groupLayerReorderDrag.indicatorTarget === 'true' &&
      beforeGroupLayerReorder.childStageIds.length >= 2 &&
      afterGroupLayerReorder.groupMemberContiguous &&
      afterGroupLayerReorder.childStageIds.join(' ') === beforeGroupLayerReorder.childStageIds.join(' ') &&
      afterGroupLayerReorder.groupStartIndex !== beforeGroupLayerReorder.groupStartIndex &&
      afterGroupLayerReorder.selectedStageIds.join(' ') === beforeGroupLayerReorder.childStageIds.join(' '),
    {
      afterGroupLayerReorder,
      beforeGroupLayerReorder,
      groupLayerReorderDrag,
    },
  )

  await pressKey(page, {
    code: 'ArrowLeft',
    key: 'ArrowLeft',
    windowsVirtualKeyCode: 37,
  })
  await delay(80)
  const afterGroupCollapse = await readPPTLayerPaneGroupTreeState(page)

  await pressKey(page, {
    code: 'ArrowRight',
    key: 'ArrowRight',
    windowsVirtualKeyCode: 39,
  })
  await delay(80)
  const afterGroupExpand = await readPPTLayerPaneGroupTreeState(page)

  await pressKey(page, {
    code: 'ArrowRight',
    key: 'ArrowRight',
    windowsVirtualKeyCode: 39,
  })
  await delay(80)
  const afterGroupFocusChild = await readPPTLayerPaneGroupTreeState(page)

  await pressKey(page, {
    code: 'ArrowLeft',
    key: 'ArrowLeft',
    windowsVirtualKeyCode: 37,
  })
  await delay(80)
  const afterGroupFocusParent = await readPPTLayerPaneGroupTreeState(page)

  record(
    'expands and collapses PPT group layer pane rows with slide-edit keyboard intent',
    afterGroupCollapse.keyboardIntentModel === 'slide-edit-layer-pane-keyboard-intent' &&
      afterGroupCollapse.groupExpanded === 'false' &&
      afterGroupCollapse.childRowCount === 0 &&
      afterGroupCollapse.focusedRowId === afterGroupCollapse.groupRowId &&
      afterGroupExpand.groupExpanded === 'true' &&
      afterGroupExpand.childRowCount === 3 &&
      afterGroupExpand.focusedRowId === afterGroupExpand.groupRowId &&
      afterGroupFocusChild.focusedParentId === afterGroupFocusChild.groupRowId &&
      afterGroupFocusChild.focusedRowId !== afterGroupFocusChild.groupRowId &&
      afterGroupFocusChild.selectedStageCount === 3 &&
      afterGroupFocusParent.focusedRowId === afterGroupFocusParent.groupRowId,
    {
      afterGroupCollapse,
      afterGroupExpand,
      afterGroupFocusChild,
      afterGroupFocusParent,
    },
  )

  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)
  await clickMouse(page, afterDistribute.x, afterDistribute.y, 1)
  await delay(50)

  const afterGroupedClick = await page.eval(`(() => ({
    selectedCount: document.querySelectorAll('[data-selected="true"]').length,
    selectedIds: [...document.querySelectorAll('[data-selected="true"]')]
      .map((element) => element.getAttribute('data-ppt-element')),
  }))()`)

  record('selects all grouped PPT objects from one grouped member click', afterGroupedClick.selectedCount === 3 && afterGroupedClick.selectedIds.includes('s1-card-1'), afterGroupedClick)

  const beforeGroupNudge = await page.eval(`(() => ({
    card1: parseFloat(document.querySelector('[data-ppt-element="s1-card-1"]').style.left),
    card2: parseFloat(document.querySelector('[data-ppt-element="s1-card-2"]').style.left),
    panel: parseFloat(document.querySelector('[data-ppt-element="s1-side-panel"]').style.left),
  }))()`)

  await pressKey(page, {
    code: 'ArrowRight',
    key: 'ArrowRight',
    windowsVirtualKeyCode: 39,
  })
  await delay(50)

  const afterGroupNudge = await page.eval(`(() => ({
    card1: parseFloat(document.querySelector('[data-ppt-element="s1-card-1"]').style.left),
    card2: parseFloat(document.querySelector('[data-ppt-element="s1-card-2"]').style.left),
    panel: parseFloat(document.querySelector('[data-ppt-element="s1-side-panel"]').style.left),
    selectedCount: document.querySelectorAll('[data-selected="true"]').length,
  }))()`)

  record('nudges grouped PPT objects together', afterGroupNudge.selectedCount === 3 && afterGroupNudge.card1 === beforeGroupNudge.card1 + 1 && afterGroupNudge.card2 === beforeGroupNudge.card2 + 1 && afterGroupNudge.panel === beforeGroupNudge.panel + 1, {
    afterGroupNudge,
    beforeGroupNudge,
  })

  await page.eval(`document.querySelector('[data-ppt-command="ungroup"]').click()`)
  await delay(50)

  const afterUngroup = await page.eval(`(() => {
    const selected = [...document.querySelectorAll('[data-selected="true"]')]

    return {
      groupedLayerCount: document.querySelectorAll('[data-ppt-layer-row][data-grouped="true"]').length,
      groupIds: selected.map((element) => element.getAttribute('data-group-id')),
      selectedCount: selected.length,
      ungroupDisabled: document.querySelector('[data-ppt-command="ungroup"]').disabled,
    }
  })()`)

  record('ungroups selected PPT objects with canvas command adapter', afterUngroup.selectedCount === 3 && afterUngroup.groupIds.every((groupId) => groupId === null) && afterUngroup.groupedLayerCount === 0 && afterUngroup.ungroupDisabled, afterUngroup)

  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)
  await clickMouse(page, afterDistribute.x, afterDistribute.y, 1)
  await delay(50)

  await page.eval(`document.querySelector('[data-ppt-command="send-to-back"]').click()`)
  await delay(50)

  const afterReorder = await page.eval(`(() => {
    const order = [...document.querySelectorAll('[data-ppt-element]')]
      .map((element) => element.getAttribute('data-ppt-element'))

    return {
      first: order[0],
      order,
    }
  })()`)

  record('reorders selected object with canvas command adapter', afterReorder.first === 's1-card-1', afterReorder)

  await pressKey(page, {
    code: 'BracketRight',
    key: ']',
    modifiers: 2,
    windowsVirtualKeyCode: 221,
  })
  await delay(50)
  const afterBringForwardShortcut = await readPPTElementLayerState(page, 's1-card-1')

  await pressKey(page, {
    code: 'BracketRight',
    key: ']',
    modifiers: 10,
    windowsVirtualKeyCode: 221,
  })
  await delay(50)
  const afterBringToFrontShortcut = await readPPTElementLayerState(page, 's1-card-1')

  await pressKey(page, {
    code: 'BracketLeft',
    key: '[',
    modifiers: 2,
    windowsVirtualKeyCode: 219,
  })
  await delay(50)
  const afterSendBackwardShortcut = await readPPTElementLayerState(page, 's1-card-1')

  await pressKey(page, {
    code: 'BracketLeft',
    key: '[',
    modifiers: 10,
    windowsVirtualKeyCode: 219,
  })
  await delay(50)
  const afterSendToBackShortcut = await readPPTElementLayerState(page, 's1-card-1')

  record('runs PPT layer order keyboard shortcuts from canvas command bindings', afterBringForwardShortcut.keyboardCommandIntent === 'canvas-keyboard-command-shortcut-intent' && afterBringForwardShortcut.keyboardCommandDispatch === 'canvas-keyboard-command-dispatch' && afterBringForwardShortcut.index > afterReorder.order.indexOf('s1-card-1') && afterBringToFrontShortcut.index === afterBringToFrontShortcut.order.length - 1 && afterSendBackwardShortcut.index === afterBringToFrontShortcut.index - 1 && afterSendToBackShortcut.index === 0, {
    afterBringForwardShortcut,
    afterBringToFrontShortcut,
    afterReorder,
    afterSendBackwardShortcut,
    afterSendToBackShortcut,
  })

  await pressKey(page, {
    code: 'KeyL',
    key: 'l',
    modifiers: 2,
    windowsVirtualKeyCode: 76,
  })
  await delay(50)
  const afterLockShortcut = await readPPTElementLayerState(page, 's1-card-1')

  await pressKey(page, {
    code: 'KeyL',
    key: 'l',
    modifiers: 10,
    windowsVirtualKeyCode: 76,
  })
  await delay(50)
  const afterUnlockShortcut = await readPPTElementLayerState(page, 's1-card-1')

  record('runs PPT lock and unlock keyboard shortcuts from canvas command bindings', afterLockShortcut.keyboardCommandIntent === 'canvas-keyboard-command-shortcut-intent' && afterLockShortcut.keyboardCommandDispatch === 'canvas-keyboard-command-dispatch' && afterLockShortcut.locked === 'true' && afterUnlockShortcut.locked === 'false' && afterUnlockShortcut.selected, {
    afterLockShortcut,
    afterUnlockShortcut,
  })

  await pressKey(page, {
    code: 'KeyD',
    key: 'd',
    modifiers: 2,
    windowsVirtualKeyCode: 68,
  })
  await delay(50)

  const afterDuplicate = await page.eval(`(() => ({
    elementCount: document.querySelectorAll('[data-ppt-element]').length,
    selectedCount: document.querySelectorAll('[data-selected="true"]').length,
  }))()`)

  record('duplicates selected object with keyboard command', afterDuplicate.elementCount > afterNudge.elementCount && afterDuplicate.selectedCount >= 1, {
    afterDuplicate,
    afterNudge,
  })

  await pressKey(page, {
    code: 'KeyC',
    key: 'c',
    modifiers: 2,
    windowsVirtualKeyCode: 67,
  })
  await pressKey(page, {
    code: 'KeyV',
    key: 'v',
    modifiers: 2,
    windowsVirtualKeyCode: 86,
  })
  await delay(50)

  const afterPaste = await page.eval(`(() => ({
    elementCount: document.querySelectorAll('[data-ppt-element]').length,
    selectedCount: document.querySelectorAll('[data-selected="true"]').length,
  }))()`)

  record('pastes copied PPT object with keyboard command', afterPaste.elementCount > afterDuplicate.elementCount && afterPaste.selectedCount === 1, {
    afterDuplicate,
    afterPaste,
  })

  await pressKey(page, {
    code: 'KeyX',
    key: 'x',
    modifiers: 2,
    windowsVirtualKeyCode: 88,
  })
  await delay(50)
  await pressKey(page, {
    code: 'KeyV',
    key: 'v',
    modifiers: 2,
    windowsVirtualKeyCode: 86,
  })
  await delay(50)

  const afterCutPaste = await page.eval(`(() => ({
    clipboardOperation: document.querySelector('.ppt-stage-shell')?.getAttribute('data-ppt-clipboard-operation') ?? '',
    elementCount: document.querySelectorAll('[data-ppt-element]').length,
    pasteOperation: document.querySelector('.ppt-stage-shell')?.getAttribute('data-ppt-clipboard-paste-operation') ?? '',
    selectedCount: document.querySelectorAll('[data-selected="true"]').length,
  }))()`)

  record('cuts and pastes selected PPT object with keyboard commands', afterCutPaste.elementCount === afterPaste.elementCount && afterCutPaste.selectedCount === 1 && afterCutPaste.clipboardOperation === 'cut' && afterCutPaste.pasteOperation === 'cut', {
    afterCutPaste,
    afterPaste,
  })

  await page.eval(`(() => {
    const fill = document.querySelector('[data-ppt-style-field="fill"]')
    const valueSetter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set
    valueSetter.call(fill, '#ff0000')
    fill.dispatchEvent(new Event('input', { bubbles: true }))
    fill.dispatchEvent(new Event('change', { bubbles: true }))
    const paragraph = document.querySelector('[data-ppt-paragraph-align="center"]')
    paragraph.click()
  })()`)
  await delay(50)

  const afterStyle = await page.eval(`(() => {
    const selected = document.querySelector('[data-selected="true"]')

    return {
      background: selected.style.background,
      textAlign: selected.style.textAlign,
    }
  })()`)

  record('updates fill and paragraph align through PPT inspector primitives', afterStyle.background === 'rgb(255, 0, 0)' && afterStyle.textAlign === 'center', afterStyle)

  await pressKey(page, {
    code: 'KeyA',
    key: 'a',
    modifiers: 2,
    windowsVirtualKeyCode: 65,
  })
  await delay(50)

  const afterSelectAll = await page.eval(`(() => ({
    elementCount: document.querySelectorAll('[data-ppt-element]').length,
    selectedCount: document.querySelectorAll('[data-selected="true"]').length,
  }))()`)

  record('selects all PPT objects with keyboard command', afterSelectAll.selectedCount === afterSelectAll.elementCount, afterSelectAll)
}

async function runCommandSurfaceScenario(page) {
  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)

  let point = await getElementCenter(page, 's1-card-1')
  await clickMouse(page, point.x, point.y, 1)
  await delay(80)

  const initial = await page.eval(`(() => ({
    deleteDisabled: document.querySelector('[data-ppt-floating-command="delete"]')?.disabled ?? true,
    duplicateDisabled: document.querySelector('[data-ppt-floating-command="duplicate"]')?.disabled ?? true,
    elementCount: document.querySelectorAll('[data-ppt-element]').length,
    floatingVisible: !!document.querySelector('[data-ppt-selection-floating-bar]'),
    groupDisabled: document.querySelector('[data-ppt-floating-command="group"]')?.disabled ?? false,
    selectedId: document.querySelector('[data-selected="true"]')?.getAttribute('data-ppt-element') ?? '',
  }))()`)

  record('renders PPT selection floating command bar from command availability', initial.floatingVisible && initial.selectedId.length > 0 && !initial.duplicateDisabled && !initial.deleteDisabled && initial.groupDisabled, initial)

  await page.eval(`document.querySelector('[data-ppt-floating-command="duplicate"]')?.click()`)
  await delay(80)

  const afterFloatingDuplicate = await page.eval(`(() => ({
    contextMenuOpen: !!document.querySelector('[data-ppt-context-menu]'),
    elementCount: document.querySelectorAll('[data-ppt-element]').length,
    selectedCount: document.querySelectorAll('[data-selected="true"]').length,
  }))()`)

  record('runs duplicate from PPT selection floating bar', afterFloatingDuplicate.elementCount === initial.elementCount + 1 && afterFloatingDuplicate.selectedCount >= 1 && !afterFloatingDuplicate.contextMenuOpen, {
    afterFloatingDuplicate,
    initial,
  })

  await page.eval(`document.querySelector('[data-ppt-floating-command="delete"]')?.click()`)
  await delay(80)

  const afterFloatingDelete = await page.eval(`(() => ({
    elementCount: document.querySelectorAll('[data-ppt-element]').length,
  }))()`)

  record('runs delete from PPT selection floating bar', afterFloatingDelete.elementCount === initial.elementCount, {
    afterFloatingDelete,
    initial,
  })

  point = await getElementCenter(page, 's1-card-1')
  await clickMouse(page, point.x, point.y, 1)
  await delay(50)
  await rightClickMouse(page, point.x, point.y)
  await delay(80)

  const afterContextOpen = await page.eval(`(() => ({
    activeCanvasMenuItem: document.activeElement?.hasAttribute('data-canvas-menu-item') ?? false,
    activeCommand: document.activeElement?.getAttribute('data-ppt-context-command') ?? '',
    activeRole: document.activeElement?.getAttribute('role') ?? '',
    canvasMenuItemCount: document.querySelectorAll('[data-ppt-context-menu] [data-canvas-menu-item]').length,
    commandItemCount: document.querySelectorAll('[data-ppt-context-command]').length,
    deleteDisabled: document.querySelector('[data-ppt-context-command="delete"]')?.disabled ?? true,
    duplicateDisabled: document.querySelector('[data-ppt-context-command="duplicate"]')?.disabled ?? true,
    enabledCommands: [...document.querySelectorAll('[data-ppt-context-command]:not(:disabled)')]
      .map((item) => item.getAttribute('data-ppt-context-command') ?? ''),
    floatingVisible: !!document.querySelector('[data-ppt-selection-floating-bar]'),
    focusModel: document.querySelector('[data-ppt-context-menu]')?.getAttribute('data-ppt-context-menu-focus-model') ?? '',
    groupDisabled: document.querySelector('[data-ppt-context-command="group"]')?.disabled ?? false,
    keyboard: document.querySelector('[data-ppt-context-menu]')?.getAttribute('data-ppt-context-menu-keyboard') ?? '',
    menuOpen: !!document.querySelector('[data-ppt-context-menu]'),
    model: document.querySelector('[data-ppt-context-menu]')?.getAttribute('data-ppt-context-menu-model') ?? '',
    menuRole: document.querySelector('[data-ppt-context-menu]')?.getAttribute('role') ?? '',
    selectedId: document.querySelector('[data-selected="true"]')?.getAttribute('data-ppt-element') ?? '',
  }))()`)

  record('opens PPT context menu with shared command availability', afterContextOpen.menuOpen && !afterContextOpen.floatingVisible && !afterContextOpen.duplicateDisabled && !afterContextOpen.deleteDisabled && afterContextOpen.groupDisabled, afterContextOpen)
  record(
    'opens PPT context menu with APG menu focus contract',
    afterContextOpen.menuRole === 'menu' &&
      afterContextOpen.keyboard === 'arrow-left-right-up-down-home-end-enter-space-escape' &&
      afterContextOpen.focusModel === 'enabled-menuitem-roving' &&
      afterContextOpen.model === 'canvas-menu-roving-focus' &&
      afterContextOpen.canvasMenuItemCount === afterContextOpen.commandItemCount &&
      afterContextOpen.canvasMenuItemCount >= afterContextOpen.enabledCommands.length &&
      afterContextOpen.activeCanvasMenuItem &&
      afterContextOpen.activeRole === 'menuitem' &&
      afterContextOpen.activeCommand === afterContextOpen.enabledCommands[0] &&
      afterContextOpen.activeCommand === 'duplicate',
    afterContextOpen,
  )

  await pressKey(page, {
    code: 'ArrowDown',
    key: 'ArrowDown',
    windowsVirtualKeyCode: 40,
  })
  await delay(20)
  const afterMenuArrowDown = await page.eval(`(() => ({
    activeCommand: document.activeElement?.getAttribute('data-ppt-context-command') ?? '',
    enabledCommands: [...document.querySelectorAll('[data-ppt-context-command]:not(:disabled)')]
      .map((item) => item.getAttribute('data-ppt-context-command') ?? ''),
  }))()`)

  await pressKey(page, {
    code: 'End',
    key: 'End',
    windowsVirtualKeyCode: 35,
  })
  await delay(20)
  const afterMenuEnd = await page.eval(`(() => ({
    activeCommand: document.activeElement?.getAttribute('data-ppt-context-command') ?? '',
    enabledCommands: [...document.querySelectorAll('[data-ppt-context-command]:not(:disabled)')]
      .map((item) => item.getAttribute('data-ppt-context-command') ?? ''),
  }))()`)

  await pressKey(page, {
    code: 'Home',
    key: 'Home',
    windowsVirtualKeyCode: 36,
  })
  await delay(20)
  const afterMenuHome = await page.eval(`(() => ({
    activeCommand: document.activeElement?.getAttribute('data-ppt-context-command') ?? '',
  }))()`)

  record(
    'moves PPT context menu focus with Arrow/Home/End keys',
    afterMenuArrowDown.activeCommand === afterContextOpen.enabledCommands[1] &&
      afterMenuEnd.activeCommand === afterMenuEnd.enabledCommands.at(-1) &&
      afterMenuHome.activeCommand === afterContextOpen.enabledCommands[0],
    {
      afterContextOpen,
      afterMenuArrowDown,
      afterMenuEnd,
      afterMenuHome,
    },
  )

  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)

  const afterMenuEscape = await page.eval(`(() => ({
    menuOpen: !!document.querySelector('[data-ppt-context-menu]'),
    selectedId: document.querySelector('[data-selected="true"]')?.getAttribute('data-ppt-element') ?? '',
  }))()`)

  record('closes PPT context menu with Escape while preserving selection', !afterMenuEscape.menuOpen && afterMenuEscape.selectedId === afterContextOpen.selectedId, {
    afterContextOpen,
    afterMenuEscape,
  })

  await rightClickMouse(page, point.x, point.y)
  await delay(80)
  await pressKey(page, {
    code: 'Enter',
    key: 'Enter',
    windowsVirtualKeyCode: 13,
  })
  await delay(80)

  const afterContextDuplicate = await page.eval(`(() => ({
    elementCount: document.querySelectorAll('[data-ppt-element]').length,
    menuOpen: !!document.querySelector('[data-ppt-context-menu]'),
    selectedCount: document.querySelectorAll('[data-selected="true"]').length,
  }))()`)

  record('runs duplicate from PPT context menu keyboard action', afterContextDuplicate.elementCount === initial.elementCount + 1 && afterContextDuplicate.selectedCount >= 1 && !afterContextDuplicate.menuOpen, {
    afterContextDuplicate,
    initial,
  })

  await page.eval(`document.querySelector('[data-ppt-floating-command="delete"]')?.click()`)
  await delay(80)

  point = await getElementCenter(page, 's1-title')
  await clickMouse(page, point.x, point.y, 2)
  await delay(80)
  await page.eval(`(() => {
    const editor = document.querySelector('[data-ppt-element="s1-title"] .ppt-element-editor')

    editor?.dispatchEvent(new MouseEvent('contextmenu', {
      bubbles: true,
      cancelable: true,
      clientX: 120,
      clientY: 120,
    }))
  })()`)
  await pressKey(page, {
    code: 'ContextMenu',
    key: 'ContextMenu',
    windowsVirtualKeyCode: 93,
  })
  await delay(80)

  const afterEditableGuard = await page.eval(`(() => {
    const editor = document.querySelector('[data-ppt-element="s1-title"] .ppt-element-editor')

    return {
      editing: editor?.isContentEditable === true && document.activeElement === editor,
      menuOpen: !!document.querySelector('[data-ppt-context-menu]'),
    }
  })()`)

  record('does not open PPT command surface while native text editing is active', afterEditableGuard.editing && !afterEditableGuard.menuOpen, afterEditableGuard)

  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)

  point = await getElementCenter(page, 's1-card-1')
  await clickMouse(page, point.x, point.y, 1)
  await delay(80)
  await page.eval(`document.querySelector('[data-ppt-floating-command="lock-selection"]')?.click()`)
  await delay(80)

  const afterLockFloating = await page.eval(`(() => ({
    deleteDisabled: document.querySelector('[data-ppt-floating-command="delete"]')?.disabled ?? false,
    duplicateDisabled: document.querySelector('[data-ppt-floating-command="duplicate"]')?.disabled ?? false,
    lockDisabled: document.querySelector('[data-ppt-floating-command="lock-selection"]')?.disabled ?? false,
    locked: document.querySelector('[data-selected="true"]')?.getAttribute('data-locked') ?? null,
    selectedId: document.querySelector('[data-selected="true"]')?.getAttribute('data-ppt-element') ?? '',
  }))()`)

  record('disables PPT floating commands for locked selection', afterLockFloating.locked === 'true' && afterLockFloating.deleteDisabled && afterLockFloating.duplicateDisabled && afterLockFloating.lockDisabled, afterLockFloating)

  const lockedPoint = await page.eval(`(() => {
    const selected = document.querySelector('[data-selected="true"]')
    const rect = selected.getBoundingClientRect()

    return {
      x: rect.left + rect.width / 2,
      y: rect.top + rect.height / 2,
    }
  })()`)

  await rightClickMouse(page, lockedPoint.x, lockedPoint.y)
  await delay(80)

  const afterLockContext = await page.eval(`(() => ({
    deleteDisabled: document.querySelector('[data-ppt-context-command="delete"]')?.disabled ?? false,
    duplicateDisabled: document.querySelector('[data-ppt-context-command="duplicate"]')?.disabled ?? false,
    lockDisabled: document.querySelector('[data-ppt-context-command="lock-selection"]')?.disabled ?? false,
    menuOpen: !!document.querySelector('[data-ppt-context-menu]'),
    unlockDisabled: document.querySelector('[data-ppt-context-command="unlock-all"]')?.disabled ?? true,
  }))()`)

  record('disables PPT context commands for locked selection', afterLockContext.menuOpen && afterLockContext.deleteDisabled && afterLockContext.duplicateDisabled && afterLockContext.lockDisabled && !afterLockContext.unlockDisabled, afterLockContext)

  await page.eval(`document.querySelector('[data-ppt-context-command="unlock-all"]')?.click()`)
  await delay(80)

  const afterUnlock = await page.eval(`(() => ({
    locked: document.querySelector('[data-selected="true"]')?.getAttribute('data-locked') ?? null,
  }))()`)

  record('unlocks locked PPT object from context menu', afterUnlock.locked === 'false', afterUnlock)
}

async function runCrossSlideClipboardScenario(page) {
  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)

  await page.eval(`document.querySelectorAll('.ppt-thumb')[0]?.click()`)
  await delay(80)

  const sourcePoint = await getElementCenter(page, 's1-title')
  await clickMouse(page, sourcePoint.x, sourcePoint.y, 1)
  await delay(80)

  const sourceBefore = await getPPTCrossSlideClipboardState(page)

  await page.eval(`(() => {
    window.__pptRichClipboardItemTypes = []
    window.__pptRichClipboardWriteCount = 0
    window.__pptRichClipboardHTML = ''
    window.__pptRichClipboardJSON = ''
    window.__pptRichClipboardPlainText = ''
    window.__pptRichClipboardSVG = ''
    window.__pptRichClipboardWriteText = ''

    window.ClipboardItem = class PPTRichClipboardItem {
      constructor(items) {
        this.items = items
        window.__pptRichClipboardItemTypes.push(Object.keys(items).sort())
      }
    }

    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: {
        write: async (items) => {
          window.__pptRichClipboardWriteCount = items.length
          const item = items[0]
          const mimeType = Object.keys(item.items)
            .find((type) => type !== 'text/html' && type !== 'text/plain' && type !== 'image/svg+xml') ?? ''

          window.__pptRichClipboardHTML = await item.items['text/html'].text()
          window.__pptRichClipboardPlainText = await item.items['text/plain'].text()
          window.__pptRichClipboardSVG = item.items['image/svg+xml']
            ? await item.items['image/svg+xml'].text()
            : ''
          window.__pptRichClipboardJSON = mimeType
            ? await item.items[mimeType].text()
            : ''
        },
        writeText: async (text) => {
          window.__pptRichClipboardWriteText = text
        },
      },
    })
  })()`)

  await pressKey(page, {
    code: 'KeyC',
    key: 'c',
    modifiers: 2,
    windowsVirtualKeyCode: 67,
  })
  await delay(120)

  const afterCopy = await getPPTCrossSlideClipboardState(page)
  const richClipboardWrite = await page.eval(`(() => ({
    html: window.__pptRichClipboardHTML ?? '',
    itemTypes: window.__pptRichClipboardItemTypes?.at(-1) ?? [],
    json: window.__pptRichClipboardJSON ?? '',
    plainText: window.__pptRichClipboardPlainText ?? '',
    svg: window.__pptRichClipboardSVG ?? '',
    writeCount: window.__pptRichClipboardWriteCount ?? 0,
    writeText: window.__pptRichClipboardWriteText ?? '',
  }))()`)

  record('stores PPT clipboard source slide metadata', afterCopy.keyboardCommandIntent === 'canvas-keyboard-command-shortcut-intent' && afterCopy.keyboardCommandDispatch === 'canvas-keyboard-command-dispatch' && afterCopy.clipboardModel === 'slide-edit-clipboard' && afterCopy.clipboardCount === 1 && afterCopy.clipboardSourceSlide === 'slide-1' && afterCopy.clipboardSelection === 's1-title' && afterCopy.clipboardType === 'slide-object-clipboard' && afterCopy.clipboardOperation === 'copy' && afterCopy.clipboardMetadataCount === 1 && afterCopy.clipboardSelectedObjectIds === 's1-title', {
    afterCopy,
    sourceBefore,
  })
  record('copies PPT selection as rich clipboard bundle', afterCopy.richClipboardModel === 'canvas-board-io-ppt-rich-clipboard' && afterCopy.richClipboardFormats.includes('application/vnd.interactive-os.ppt.selection+json') && afterCopy.richClipboardFormats.includes('text/html') && afterCopy.richClipboardFormats.includes('image/svg+xml') && afterCopy.richClipboardFormats.includes('text/plain') && afterCopy.richClipboardJsonMimeType === 'application/vnd.interactive-os.ppt.selection+json' && afterCopy.richClipboardObjectCount === 1 && afterCopy.richClipboardSelection === 's1-title' && afterCopy.richClipboardSourceSlide === 'slide-1' && afterCopy.richClipboardWriteMode === 'clipboard-item' && afterCopy.richClipboardPlainTextLength > 0 && afterCopy.richClipboardHTMLLength > afterCopy.richClipboardPlainTextLength && richClipboardWrite.writeCount === 1 && richClipboardWrite.itemTypes.includes(afterCopy.richClipboardJsonMimeType) && richClipboardWrite.itemTypes.includes('text/html') && richClipboardWrite.itemTypes.includes('text/plain') && richClipboardWrite.itemTypes.includes('image/svg+xml') && richClipboardWrite.plainText === afterCopy.selectedText && richClipboardWrite.html.includes(afterCopy.selectedText) && richClipboardWrite.html.includes('data-ppt-rich-clipboard-json') && richClipboardWrite.json.includes('"kind": "interactive-os.ppt.selection"') && richClipboardWrite.svg.includes('<svg'), {
    afterCopy,
    richClipboardWrite,
  })

  await page.eval(`((html, plainText) => {
    const dataTransfer = new DataTransfer()
    const fallbackHTML = html.replace(/<script\\b[\\s\\S]*?<\\/script>/gi, '')

    dataTransfer.setData('text/html', fallbackHTML)
    dataTransfer.setData('text/plain', plainText)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })(${JSON.stringify(richClipboardWrite.html)}, ${JSON.stringify(richClipboardWrite.plainText)})`)
  await delay(120)

  const afterTextFallbackHTMLPaste = await getPPTCrossSlideClipboardState(page)

  record(
    'pastes PPT text fallback HTML without embedded JSON as editable text box',
    afterTextFallbackHTMLPaste.stageCount === sourceBefore.stageCount + 1 &&
      afterTextFallbackHTMLPaste.selectedKind === 'textBox' &&
      afterTextFallbackHTMLPaste.selectedText === afterCopy.selectedText &&
      afterTextFallbackHTMLPaste.selectedFontSize === '56px' &&
      afterTextFallbackHTMLPaste.selectedWidth === sourceBefore.selectedWidth &&
      afterTextFallbackHTMLPaste.selectedHeight === sourceBefore.selectedHeight &&
      afterTextFallbackHTMLPaste.fallbackHTMLImportModel === 'ppt-fallback-html-import' &&
      afterTextFallbackHTMLPaste.fallbackHTMLImportFormat === 'text-html-ppt-fallback' &&
      afterTextFallbackHTMLPaste.fallbackHTMLImportKind === 'textBox' &&
      afterTextFallbackHTMLPaste.fallbackHTMLImportSourceObject === afterCopy.selectedId,
    {
      afterCopy,
      afterTextFallbackHTMLPaste,
      sourceBefore,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterTextFallbackUndo = await getPPTCrossSlideClipboardState(page)

  record(
    'undoes PPT text fallback HTML paste before cross-slide clipboard scenario continues',
    afterTextFallbackUndo.stageCount === sourceBefore.stageCount,
    {
      afterTextFallbackUndo,
      sourceBefore,
    },
  )

  await page.eval(`document.querySelectorAll('.ppt-thumb')[1]?.click()`)
  await delay(80)

  const targetBefore = await getPPTCrossSlideClipboardState(page)

  await pressKey(page, {
    code: 'KeyV',
    key: 'v',
    modifiers: 2,
    windowsVirtualKeyCode: 86,
  })
  await delay(80)

  const afterKeyboardPaste = await getPPTCrossSlideClipboardState(page)
  const firstPasteAnchor = parsePPTPoint(afterKeyboardPaste.pasteAnchor)
  const expectedFirstPasteAnchor = {
    x: afterKeyboardPaste.pastePositionViewportX - (sourceBefore.selectedX + sourceBefore.selectedWidth / 2),
    y: afterKeyboardPaste.pastePositionViewportY - (sourceBefore.selectedY + sourceBefore.selectedHeight / 2),
  }

  record('pastes copied PPT object onto another slide with target slide ids', afterKeyboardPaste.keyboardCommandIntent === 'canvas-keyboard-command-shortcut-intent' && afterKeyboardPaste.keyboardCommandDispatch === 'canvas-keyboard-command-dispatch' && afterKeyboardPaste.activeSlide === 'slide-2' && afterKeyboardPaste.stageCount === targetBefore.stageCount + 1 && afterKeyboardPaste.selectedCount === 1 && afterKeyboardPaste.selectedId.startsWith('slide-2-') && afterKeyboardPaste.selectedName.includes('Copy'), {
    afterKeyboardPaste,
    targetBefore,
  })
  record('creates PPT cross-slide paste command effect plan', afterKeyboardPaste.pasteCommand === 'paste-slide-objects' && afterKeyboardPaste.pasteType === 'slide-command-effect' && afterKeyboardPaste.pasteSourceSlide === 'slide-1' && afterKeyboardPaste.pasteTargetSlide === 'slide-2' && afterKeyboardPaste.pasteMappingCount === 1 && afterKeyboardPaste.pasteSelection === afterKeyboardPaste.selectedId && nearlyEqual(firstPasteAnchor.x, expectedFirstPasteAnchor.x, 0.001) && nearlyEqual(firstPasteAnchor.y, expectedFirstPasteAnchor.y, 0.001) && afterKeyboardPaste.pasteOperation === 'copy', {
    afterKeyboardPaste,
    expectedFirstPasteAnchor,
    firstPasteAnchor,
    targetBefore,
  })
  record('routes first PPT object paste anchor through canvas paste position', afterKeyboardPaste.pastePositionModel === 'canvas-paste-position' && afterKeyboardPaste.pastePositionIndex === 0 && afterKeyboardPaste.pastePositionCount === 1 && nearlyEqual(afterKeyboardPaste.pastePositionViewportX, afterKeyboardPaste.pastePositionBoundsX + afterKeyboardPaste.pastePositionBoundsWidth / 2 + firstPasteAnchor.x, 0.001) && nearlyEqual(afterKeyboardPaste.pastePositionViewportY, afterKeyboardPaste.pastePositionBoundsY + afterKeyboardPaste.pastePositionBoundsHeight / 2 + firstPasteAnchor.y, 0.001), {
    afterKeyboardPaste,
    firstPasteAnchor,
  })

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterUndo = await getPPTCrossSlideClipboardState(page)

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const afterRedo = await getPPTCrossSlideClipboardState(page)

  record('undoes and redoes PPT cross-slide paste as one history step', afterUndo.keyboardCommandIntent === 'canvas-keyboard-command-shortcut-intent' && afterUndo.keyboardCommandDispatch === 'canvas-keyboard-command-dispatch' && afterUndo.stageCount === targetBefore.stageCount && afterRedo.stageCount === afterKeyboardPaste.stageCount && afterRedo.selectedId.startsWith('slide-2-'), {
    afterKeyboardPaste,
    afterRedo,
    afterUndo,
    targetBefore,
  })

  await pressKey(page, {
    code: 'KeyK',
    key: 'k',
    modifiers: 2,
    windowsVirtualKeyCode: 75,
  })
  await delay(80)
  await page.send('Input.insertText', { text: 'paste' })
  await delay(80)

  const beforePalettePaste = await page.eval(`(() => ({
    disabled: document.querySelector('[data-ppt-command-palette-item="command:paste"]')?.disabled ?? true,
    itemPresent: !!document.querySelector('[data-ppt-command-palette-item="command:paste"]'),
    open: !!document.querySelector('[data-ppt-command-palette]'),
  }))()`)

  await pressKey(page, {
    code: 'Enter',
    key: 'Enter',
    windowsVirtualKeyCode: 13,
  })
  await delay(80)

  const afterPalettePaste = await getPPTCrossSlideClipboardState(page)

  record('runs PPT cross-slide paste from command palette', beforePalettePaste.open && beforePalettePaste.itemPresent && !beforePalettePaste.disabled && afterPalettePaste.stageCount === afterRedo.stageCount + 1 && afterPalettePaste.selectedId.startsWith('slide-2-'), {
    afterPalettePaste,
    afterRedo,
    beforePalettePaste,
  })
  record('routes repeated PPT object paste through canvas paste offset', afterPalettePaste.pastePositionModel === 'canvas-paste-position' && afterPalettePaste.pastePositionIndex === 1 && afterPalettePaste.pasteAnchor === '28,28' && afterPalettePaste.pasteSelection === afterPalettePaste.selectedId, {
    afterPalettePaste,
  })

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)
  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterCleanup = await getPPTCrossSlideClipboardState(page)

  await page.eval(`document.querySelectorAll('.ppt-thumb')[0]?.click()`)
  await delay(80)

  const sourceAfter = await getPPTCrossSlideClipboardState(page)

  record('keeps source PPT slide unchanged after cross-slide paste', afterCleanup.stageCount === afterRedo.stageCount && sourceAfter.activeSlide === 'slide-1' && sourceAfter.stageCount === sourceBefore.stageCount && sourceAfter.thumbCounts[0] === sourceBefore.thumbCounts[0], {
    afterCleanup,
    afterRedo,
    sourceAfter,
    sourceBefore,
    targetBefore,
  })

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const objects = [
      {
        geometry: { h: 88, w: 300, x: 180, y: 210 },
        id: 'ai-element-title',
        kind: 'textBox',
        name: 'AI Element Title',
        style: { color: '#111827', fontSize: 32, fontWeight: 'bold' },
        textBody: { paragraphs: [{ runs: [{ text: 'AI element JSON' }] }] },
      },
      {
        fill: { color: '#fef3c7' },
        geometry: { h: 104, w: 240, x: 540, y: 230 },
        hyperlink: { url: 'https://example.com/element-json' },
        id: 'ai-element-card',
        kind: 'shape',
        name: 'AI Element Card',
        shape: 'rect',
        stroke: { color: '#d97706', width: 2 },
        style: { color: '#78350f', fontSize: 24, fontWeight: 'semibold' },
        textBody: { paragraphs: [{ runs: [{ text: 'Retouch me' }] }] },
      },
    ]
    const json = JSON.stringify({
      objects,
      selectedObjectIds: ['ai-element-title', 'ai-element-card'],
      sourceSlideId: 'ai-json-elements',
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(100)

  const afterElementJSONPaste = await getPPTCrossSlideClipboardState(page)

  record(
    'pastes raw PPT element JSON through canvas clipboard paste pipeline',
    afterElementJSONPaste.activeSlide === 'slide-1' &&
      afterElementJSONPaste.stageCount === sourceAfter.stageCount + 2 &&
      afterElementJSONPaste.selectedCount === 2 &&
      afterElementJSONPaste.selectedKinds.includes('textBox') &&
      afterElementJSONPaste.selectedKinds.includes('shape') &&
      afterElementJSONPaste.selectedTexts.includes('AI element JSON') &&
      afterElementJSONPaste.selectedTexts.includes('Retouch me') &&
      afterElementJSONPaste.elementsJSONImportModel === 'ppt-elements-json-import' &&
      afterElementJSONPaste.elementsJSONImportFormat === 'application-json-ppt-elements' &&
      afterElementJSONPaste.elementsJSONImportCount === 2 &&
      afterElementJSONPaste.elementsJSONImportSelection === 'ai-element-title ai-element-card' &&
      afterElementJSONPaste.elementsJSONImportSourceSlide === 'ai-json-elements' &&
      afterElementJSONPaste.elementsJSONImportJsonLength > 100 &&
      afterElementJSONPaste.pasteCommand === 'paste-slide-objects' &&
      afterElementJSONPaste.pasteType === 'slide-command-effect' &&
      afterElementJSONPaste.pasteSourceSlide === 'ai-json-elements' &&
      afterElementJSONPaste.pasteTargetSlide === 'slide-1' &&
      afterElementJSONPaste.pasteMappingCount === 2 &&
      afterElementJSONPaste.pastePositionModel === 'canvas-paste-position',
    {
      afterElementJSONPaste,
      sourceAfter,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterElementJSONUndo = await getPPTCrossSlideClipboardState(page)

  record(
    'undoes raw PPT element JSON paste as one history step',
    afterElementJSONUndo.stageCount === sourceAfter.stageCount &&
      afterElementJSONUndo.activeSlide === 'slide-1',
    {
      afterElementJSONPaste,
      afterElementJSONUndo,
      sourceAfter,
    },
  )

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const payload = {
      kind: 'interactive-os.ppt.selection',
      metadata: {
        objectCount: 1,
        selectedObjectIds: ['external-rich-shape'],
        sourceSlideId: 'external-slide',
      },
      payload: {
        metadata: [{
          groupId: null,
          objectId: 'external-rich-shape',
          placeholderId: null,
        }],
        objects: [{
          fill: { color: '#22c55e', opacity: 1 },
          geometry: { h: 88, w: 180, x: 120, y: 160 },
          id: 'external-rich-shape',
          kind: 'shape',
          name: 'External Clipboard Shape',
          shape: 'rect',
          stroke: { color: '#15803d', width: 2 },
          style: { color: '#052e16', fontSize: 24 },
          textBody: { paragraphs: [{ runs: [{ text: 'Rich paste' }] }] },
        }],
        operation: 'copy',
        selectedObjectIds: ['external-rich-shape'],
        sourceSlideId: 'external-slide',
        type: 'slide-object-clipboard',
      },
      version: 1,
    }
    const json = JSON.stringify(payload)

    dataTransfer.setData('application/vnd.interactive-os.ppt.selection+json', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(80)

  const afterRichClipboardPaste = await getPPTCrossSlideClipboardState(page)

  record('pastes PPT rich clipboard custom JSON as object selection', afterRichClipboardPaste.activeSlide === 'slide-1' && afterRichClipboardPaste.stageCount === sourceAfter.stageCount + 1 && afterRichClipboardPaste.selectedCount === 1 && afterRichClipboardPaste.selectedKind === 'shape' && afterRichClipboardPaste.selectedName.includes('External Clipboard Shape Copy') && afterRichClipboardPaste.pasteSourceSlide === 'external-slide' && afterRichClipboardPaste.pasteTargetSlide === 'slide-1' && afterRichClipboardPaste.richClipboardImported === 'true' && afterRichClipboardPaste.richClipboardImportFormat === 'custom-json' && afterRichClipboardPaste.richClipboardModel === 'canvas-board-io-ppt-rich-clipboard', {
    afterRichClipboardPaste,
    sourceAfter,
  })

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const payload = {
      kind: 'interactive-os.ppt.selection',
      metadata: {
        objectCount: 1,
        selectedObjectIds: ['external-html-shape'],
        sourceSlideId: 'html-slide',
      },
      payload: {
        metadata: [{
          groupId: null,
          objectId: 'external-html-shape',
          placeholderId: null,
        }],
        objects: [{
          fill: { color: '#f97316', opacity: 1 },
          geometry: { h: 76, w: 190, x: 180, y: 210 },
          id: 'external-html-shape',
          kind: 'shape',
          name: 'HTML Clipboard Shape',
          shape: 'rect',
          stroke: { color: '#c2410c', width: 2 },
          style: { color: '#431407', fontSize: 22 },
          textBody: { paragraphs: [{ runs: [{ text: 'HTML paste' }] }] },
        }],
        operation: 'copy',
        selectedObjectIds: ['external-html-shape'],
        sourceSlideId: 'html-slide',
        type: 'slide-object-clipboard',
      },
      version: 1,
    }
    const json = JSON.stringify(payload).replace(/</g, '\\\\u003c')
    const html = '<section data-ppt-rich-clipboard="true"><script type="application/json" data-ppt-rich-clipboard-json>' + json + '</script></section>'

    dataTransfer.setData('text/html', html)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(80)

  const afterRichHTMLPaste = await getPPTCrossSlideClipboardState(page)

  record('pastes PPT rich clipboard embedded HTML JSON as object selection', afterRichHTMLPaste.activeSlide === 'slide-1' && afterRichHTMLPaste.stageCount === sourceAfter.stageCount + 1 && afterRichHTMLPaste.selectedKind === 'shape' && afterRichHTMLPaste.selectedName.includes('HTML Clipboard Shape Copy') && afterRichHTMLPaste.pasteSourceSlide === 'html-slide' && afterRichHTMLPaste.richClipboardImported === 'true' && afterRichHTMLPaste.richClipboardImportFormat === 'text-html', {
    afterRichHTMLPaste,
    sourceAfter,
  })

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const shapeCopyPoint = await getElementCenter(page, 's1-card-1')
  await clickMouse(page, shapeCopyPoint.x, shapeCopyPoint.y, 1)
  await delay(80)

  await page.eval(`(() => {
    window.__pptShapeRichClipboardItemTypes = []
    window.__pptShapeRichClipboardWriteCount = 0
    window.__pptShapeRichClipboardHTML = ''
    window.__pptShapeRichClipboardJSON = ''
    window.__pptShapeRichClipboardPlainText = ''
    window.__pptShapeRichClipboardSVG = ''

    window.ClipboardItem = class PPTShapeRichClipboardItem {
      constructor(items) {
        this.items = items
        window.__pptShapeRichClipboardItemTypes.push(Object.keys(items).sort())
      }
    }

    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: {
        write: async (items) => {
          window.__pptShapeRichClipboardWriteCount = items.length
          const item = items[0]
          const mimeType = Object.keys(item.items)
            .find((type) => type !== 'text/html' && type !== 'text/plain' && type !== 'image/svg+xml') ?? ''

          window.__pptShapeRichClipboardHTML = await item.items['text/html'].text()
          window.__pptShapeRichClipboardPlainText = await item.items['text/plain'].text()
          window.__pptShapeRichClipboardSVG = item.items['image/svg+xml']
            ? await item.items['image/svg+xml'].text()
            : ''
          window.__pptShapeRichClipboardJSON = mimeType
            ? await item.items[mimeType].text()
            : ''
        },
      },
    })
  })()`)

  await pressKey(page, {
    code: 'KeyC',
    key: 'c',
    modifiers: 2,
    windowsVirtualKeyCode: 67,
  })
  await delay(120)

  const afterShapeCopy = await getPPTCrossSlideClipboardState(page)
  const shapeRichClipboardWrite = await page.eval(`(() => ({
    html: window.__pptShapeRichClipboardHTML ?? '',
    itemTypes: window.__pptShapeRichClipboardItemTypes?.at(-1) ?? [],
    json: window.__pptShapeRichClipboardJSON ?? '',
    plainText: window.__pptShapeRichClipboardPlainText ?? '',
    svg: window.__pptShapeRichClipboardSVG ?? '',
    writeCount: window.__pptShapeRichClipboardWriteCount ?? 0,
  }))()`)

  record(
    'copies PPT shape selection with styled HTML clipboard fallback',
    afterShapeCopy.richClipboardModel === 'canvas-board-io-ppt-rich-clipboard' &&
      afterShapeCopy.richClipboardWriteMode === 'clipboard-item' &&
      afterShapeCopy.selectedKind === 'shape' &&
      afterShapeCopy.richClipboardSelection === afterShapeCopy.selectedId &&
      afterShapeCopy.richClipboardPlainTextLength > 0 &&
      afterShapeCopy.richClipboardHTMLLength > afterShapeCopy.richClipboardPlainTextLength &&
      shapeRichClipboardWrite.writeCount === 1 &&
      shapeRichClipboardWrite.itemTypes.includes(afterShapeCopy.richClipboardJsonMimeType) &&
      shapeRichClipboardWrite.itemTypes.includes('text/html') &&
      shapeRichClipboardWrite.itemTypes.includes('text/plain') &&
      shapeRichClipboardWrite.itemTypes.includes('image/svg+xml') &&
      shapeRichClipboardWrite.html.includes('data-ppt-selection-shape="rect"') &&
      shapeRichClipboardWrite.html.includes('background:#e0f2fe') &&
      shapeRichClipboardWrite.html.includes('border:2px solid #0ea5e9') &&
      shapeRichClipboardWrite.html.includes('border-radius:24px') &&
      shapeRichClipboardWrite.html.includes('width:') &&
      shapeRichClipboardWrite.html.includes('height:') &&
      shapeRichClipboardWrite.html.includes('Fast draft') &&
      shapeRichClipboardWrite.html.includes('Small edits') &&
      shapeRichClipboardWrite.plainText.includes('Fast draft') &&
      shapeRichClipboardWrite.plainText.includes('Small edits') &&
      shapeRichClipboardWrite.json.includes('"kind": "interactive-os.ppt.selection"') &&
      shapeRichClipboardWrite.svg.includes('<svg'),
    {
      afterShapeCopy,
      shapeRichClipboardWrite,
    },
  )

  await page.eval(`((html, plainText) => {
    const dataTransfer = new DataTransfer()
    const fallbackHTML = html.replace(/<script\\b[\\s\\S]*?<\\/script>/gi, '')

    dataTransfer.setData('text/html', fallbackHTML)
    dataTransfer.setData('text/plain', plainText)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })(${JSON.stringify(shapeRichClipboardWrite.html)}, ${JSON.stringify(shapeRichClipboardWrite.plainText)})`)
  await delay(120)

  const afterShapeFallbackHTMLPaste = await getPPTCrossSlideClipboardState(page)

  record(
    'pastes PPT shape fallback HTML without embedded JSON as editable shape',
    afterShapeFallbackHTMLPaste.stageCount === afterShapeCopy.stageCount + 1 &&
      afterShapeFallbackHTMLPaste.selectedKind === 'shape' &&
      afterShapeFallbackHTMLPaste.selectedShape === 'rect' &&
      afterShapeFallbackHTMLPaste.selectedText.includes('Fast draft') &&
      afterShapeFallbackHTMLPaste.selectedText.includes('Small edits') &&
      afterShapeFallbackHTMLPaste.selectedFill.includes('224, 242, 254') &&
      afterShapeFallbackHTMLPaste.selectedBorderColor === 'rgb(14, 165, 233)' &&
      afterShapeFallbackHTMLPaste.selectedCornerRadius === '24' &&
      afterShapeFallbackHTMLPaste.selectedWidth > 0 &&
      afterShapeFallbackHTMLPaste.selectedHeight > 0 &&
      afterShapeFallbackHTMLPaste.fallbackHTMLImportModel === 'ppt-fallback-html-import' &&
      afterShapeFallbackHTMLPaste.fallbackHTMLImportFormat === 'text-html-ppt-fallback' &&
      afterShapeFallbackHTMLPaste.fallbackHTMLImportShape === 'rect' &&
      afterShapeFallbackHTMLPaste.fallbackHTMLImportSourceObject === afterShapeCopy.selectedId,
    {
      afterShapeCopy,
      afterShapeFallbackHTMLPaste,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const multiTextPoint = await getElementCenter(page, 's1-title')
  const multiShapePoint = await getElementCenter(page, 's1-card-1')

  await clickMouse(page, multiTextPoint.x, multiTextPoint.y, 1)
  await delay(80)
  await clickMouse(page, multiShapePoint.x, multiShapePoint.y, 1, 8)
  await delay(80)

  await page.eval(`(() => {
    window.__pptMultiRichClipboardItemTypes = []
    window.__pptMultiRichClipboardWriteCount = 0
    window.__pptMultiRichClipboardHTML = ''
    window.__pptMultiRichClipboardJSON = ''
    window.__pptMultiRichClipboardPlainText = ''
    window.__pptMultiRichClipboardSVG = ''

    window.ClipboardItem = class PPTMultiRichClipboardItem {
      constructor(items) {
        this.items = items
        window.__pptMultiRichClipboardItemTypes.push(Object.keys(items).sort())
      }
    }

    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: {
        write: async (items) => {
          window.__pptMultiRichClipboardWriteCount = items.length
          const item = items[0]
          const mimeType = Object.keys(item.items)
            .find((type) => type !== 'text/html' && type !== 'text/plain' && type !== 'image/svg+xml') ?? ''

          window.__pptMultiRichClipboardHTML = await item.items['text/html'].text()
          window.__pptMultiRichClipboardPlainText = await item.items['text/plain'].text()
          window.__pptMultiRichClipboardSVG = item.items['image/svg+xml']
            ? await item.items['image/svg+xml'].text()
            : ''
          window.__pptMultiRichClipboardJSON = mimeType
            ? await item.items[mimeType].text()
            : ''
        },
      },
    })
  })()`)

  await pressKey(page, {
    code: 'KeyC',
    key: 'c',
    modifiers: 2,
    windowsVirtualKeyCode: 67,
  })
  await delay(120)

  const afterMultiCopy = await getPPTCrossSlideClipboardState(page)
  const multiRichClipboardWrite = await page.eval(`(() => ({
    html: window.__pptMultiRichClipboardHTML ?? '',
    itemTypes: window.__pptMultiRichClipboardItemTypes?.at(-1) ?? [],
    json: window.__pptMultiRichClipboardJSON ?? '',
    plainText: window.__pptMultiRichClipboardPlainText ?? '',
    svg: window.__pptMultiRichClipboardSVG ?? '',
    writeCount: window.__pptMultiRichClipboardWriteCount ?? 0,
  }))()`)
  const multiCopiedShapeId = afterMultiCopy.selectedIds.find((id, index) =>
    afterMultiCopy.selectedKinds[index] === 'shape') ?? ''

  record(
    'copies multi-selected PPT text and shape with positioned fallback HTML',
    afterMultiCopy.selectedCount === 2 &&
      afterMultiCopy.richClipboardObjectCount === 2 &&
      Boolean(multiCopiedShapeId) &&
      afterMultiCopy.selectedIds.includes('s1-title') &&
      afterMultiCopy.selectedKinds.includes('shape') &&
      afterMultiCopy.richClipboardSelection.includes('s1-title') &&
      afterMultiCopy.richClipboardSelection.includes(multiCopiedShapeId) &&
      multiRichClipboardWrite.writeCount === 1 &&
      multiRichClipboardWrite.itemTypes.includes(afterMultiCopy.richClipboardJsonMimeType) &&
      multiRichClipboardWrite.html.includes(`data-ppt-selection-object="${multiCopiedShapeId}"`) &&
      multiRichClipboardWrite.html.includes('data-ppt-selection-text-body="true"') &&
      multiRichClipboardWrite.html.includes('data-ppt-selection-shape="rect"') &&
      multiRichClipboardWrite.html.includes('data-ppt-selection-x="') &&
      multiRichClipboardWrite.html.includes('data-ppt-selection-y="') &&
      multiRichClipboardWrite.html.includes('Edited title') &&
      multiRichClipboardWrite.html.includes('Fast draft'),
    {
      afterMultiCopy,
      multiCopiedShapeId,
      multiRichClipboardWrite,
    },
  )

  await page.eval(`((html, plainText) => {
    const dataTransfer = new DataTransfer()
    const fallbackHTML = html.replace(/<script\\b[\\s\\S]*?<\\/script>/gi, '')

    dataTransfer.setData('text/html', fallbackHTML)
    dataTransfer.setData('text/plain', plainText)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })(${JSON.stringify(multiRichClipboardWrite.html)}, ${JSON.stringify(multiRichClipboardWrite.plainText)})`)
  await delay(120)

  const afterMultiFallbackHTMLPaste = await getPPTCrossSlideClipboardState(page)

  record(
    'pastes multi-object PPT fallback HTML as editable selection',
    afterMultiFallbackHTMLPaste.stageCount === afterMultiCopy.stageCount + 2 &&
      afterMultiFallbackHTMLPaste.selectedCount === 2 &&
      afterMultiFallbackHTMLPaste.selectedKinds.includes('textBox') &&
      afterMultiFallbackHTMLPaste.selectedKinds.includes('shape') &&
      afterMultiFallbackHTMLPaste.selectedTexts.some((text) =>
        text.includes('Edited title')) &&
      afterMultiFallbackHTMLPaste.selectedTexts.some((text) =>
        text.includes('Fast draft')) &&
      afterMultiFallbackHTMLPaste.fallbackHTMLImportModel === 'ppt-fallback-html-import' &&
      afterMultiFallbackHTMLPaste.fallbackHTMLImportFormat === 'text-html-ppt-fallback' &&
      afterMultiFallbackHTMLPaste.fallbackHTMLImportKind === 'selection' &&
      afterMultiFallbackHTMLPaste.fallbackHTMLImportCount === 2 &&
      afterMultiFallbackHTMLPaste.fallbackHTMLImportSourceObjects.includes('s1-title') &&
      afterMultiFallbackHTMLPaste.fallbackHTMLImportSourceObjects.includes(multiCopiedShapeId),
    {
      afterMultiCopy,
      afterMultiFallbackHTMLPaste,
      multiCopiedShapeId,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterMultiFallbackUndo = await getPPTCrossSlideClipboardState(page)

  record(
    'undoes multi-object PPT fallback HTML paste as one history step',
    afterMultiFallbackUndo.stageCount === afterMultiCopy.stageCount,
    {
      afterMultiCopy,
      afterMultiFallbackHTMLPaste,
      afterMultiFallbackUndo,
    },
  )

  const beforeExternalHTMLPaste = await getPPTCrossSlideClipboardState(page)

  await page.eval(`(() => {
    const createImage = (color, alt) => {
      const canvas = document.createElement('canvas')
      const context = canvas.getContext('2d')

      canvas.width = 120
      canvas.height = 72
      context.fillStyle = color
      context.fillRect(0, 0, 120, 72)
      context.fillStyle = '#ffffff'
      context.fillRect(18, 22, 84, 28)

      return '<figure><img alt="' + alt + '" src="' +
        canvas.toDataURL('image/png') + '"></figure>'
    }
    const dataTransfer = new DataTransfer()
    const html = '<section>' +
      createImage('#0f766e', 'North chart') +
      '<table><tr><th>Metric</th><th>Score</th></tr><tr><td>North</td><td>81</td></tr></table>' +
      createImage('#7c3aed', 'South chart') +
      '</section>'

    dataTransfer.setData('text/html', html)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(180)

  const afterExternalHTMLPaste = await getPPTCrossSlideClipboardState(page)

  record(
    'pastes external HTML data images and table as editable selection',
    afterExternalHTMLPaste.stageCount === beforeExternalHTMLPaste.stageCount + 3 &&
      afterExternalHTMLPaste.selectedCount === 3 &&
      afterExternalHTMLPaste.selectedKinds.filter((kind) => kind === 'image').length === 2 &&
      afterExternalHTMLPaste.selectedKinds.includes('table') &&
      afterExternalHTMLPaste.selectedTexts.some((text) => text.includes('North chart')) &&
      afterExternalHTMLPaste.selectedTexts.some((text) => text.includes('South chart')) &&
      afterExternalHTMLPaste.selectedTexts.some((text) => text.includes('Metric')) &&
      afterExternalHTMLPaste.fallbackHTMLImportModel === 'ppt-fallback-html-import' &&
      afterExternalHTMLPaste.fallbackHTMLImportFormat === 'text-html-ppt-fallback' &&
      afterExternalHTMLPaste.fallbackHTMLImportKind === 'selection' &&
      afterExternalHTMLPaste.fallbackHTMLImportCount === 3,
    {
      afterExternalHTMLPaste,
      beforeExternalHTMLPaste,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterExternalHTMLUndo = await getPPTCrossSlideClipboardState(page)

  record(
    'undoes external HTML fallback selection paste as one history step',
    afterExternalHTMLUndo.stageCount === beforeExternalHTMLPaste.stageCount,
    {
      afterExternalHTMLPaste,
      afterExternalHTMLUndo,
      beforeExternalHTMLPaste,
    },
  )
}

async function runSelectSameTypeScenario(page) {
  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)

  let point = await getElementCenter(page, 's1-card-1')
  await clickMouse(page, point.x, point.y, 1)
  await delay(80)

  const initialSurface = await page.eval(`(() => ({
    floatingDisabled: document.querySelector('[data-ppt-floating-command="select-same-type"]')?.disabled ?? true,
    floatingVisible: !!document.querySelector('[data-ppt-selection-floating-bar]'),
    selectedKind: document.querySelector('[data-selected="true"]')?.getAttribute('data-kind') ?? '',
    selectedShape: document.querySelector('[data-selected="true"]')?.getAttribute('data-shape') ?? '',
    selectedId: document.querySelector('[data-selected="true"]')?.getAttribute('data-ppt-element') ?? '',
  }))()`)

  record('enables PPT select same type on selection floating bar', initialSurface.floatingVisible && initialSurface.selectedKind === 'shape' && initialSurface.selectedShape === 'rect' && !initialSurface.floatingDisabled, initialSurface)

  await page.eval(`document.querySelector('[data-ppt-floating-command="select-same-type"]')?.click()`)
  await delay(80)

  const afterFloatingSelect = await page.eval(`(() => {
    const selected = [...document.querySelectorAll('[data-selected="true"]')]

    return {
      selectedIds: selected.map((element) => element.getAttribute('data-ppt-element')),
      selectedKinds: selected.map((element) => element.getAttribute('data-kind')),
      selectedShapes: selected.map((element) => element.getAttribute('data-shape')),
    }
  })()`)

  record('selects same PPT shape type from floating bar', afterFloatingSelect.selectedIds.includes('s1-card-1') && afterFloatingSelect.selectedIds.includes('s1-card-2') && afterFloatingSelect.selectedIds.includes('s1-side-panel') && afterFloatingSelect.selectedIds.length >= 3 && afterFloatingSelect.selectedKinds.every((kind) => kind === 'shape') && afterFloatingSelect.selectedShapes.every((shape) => shape === 'rect'), afterFloatingSelect)

  await page.eval(`document.querySelector('[data-ppt-layer-row="s1-card-2"] [data-ppt-layer-visibility]')?.click()`)
  await delay(80)
  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)

  point = await getElementCenter(page, 's1-card-1')
  await clickMouse(page, point.x, point.y, 1)
  await delay(60)
  await rightClickMouse(page, point.x, point.y)
  await delay(80)

  const contextSurface = await page.eval(`(() => ({
    contextDisabled: document.querySelector('[data-ppt-context-command="select-same-type"]')?.disabled ?? true,
    menuOpen: !!document.querySelector('[data-ppt-context-menu]'),
  }))()`)

  record('enables PPT select same type in context menu', contextSurface.menuOpen && !contextSurface.contextDisabled, contextSurface)

  await page.eval(`document.querySelector('[data-ppt-context-command="select-same-type"]')?.click()`)
  await delay(80)

  const afterHiddenSelect = await page.eval(`(() => ({
    card2Hidden: document.querySelector('[data-ppt-layer-row="s1-card-2"]')?.getAttribute('data-hidden') ?? '',
    card2OnStage: !!document.querySelector('[data-ppt-element="s1-card-2"]'),
    selectedIds: [...document.querySelectorAll('[data-selected="true"]')]
      .map((element) => element.getAttribute('data-ppt-element')),
  }))()`)

  record('excludes hidden PPT objects from select same type', afterHiddenSelect.card2Hidden === 'true' && !afterHiddenSelect.card2OnStage && afterHiddenSelect.selectedIds.includes('s1-card-1') && afterHiddenSelect.selectedIds.includes('s1-side-panel') && !afterHiddenSelect.selectedIds.includes('s1-card-2'), afterHiddenSelect)

  await page.eval(`document.querySelector('[data-ppt-layer-row="s1-card-2"] [data-ppt-layer-visibility]')?.click()`)
  await delay(80)
  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)

  await page.eval(`document.querySelector('[data-ppt-layer-row="s1-card-2"] [data-ppt-layer-lock]')?.click()`)
  await delay(80)
  point = await getElementCenter(page, 's1-card-1')
  await clickMouse(page, point.x, point.y, 1)
  await delay(60)
  await page.eval(`document.querySelector('[data-ppt-floating-command="select-same-type"]')?.click()`)
  await delay(80)

  const afterLockedSelect = await page.eval(`(() => ({
    card2Locked: document.querySelector('[data-ppt-layer-row="s1-card-2"]')?.getAttribute('data-locked') ?? '',
    deleteDisabled: document.querySelector('button[title="Delete"]')?.disabled ?? false,
    selectedIds: [...document.querySelectorAll('[data-selected="true"]')]
      .map((element) => element.getAttribute('data-ppt-element')),
  }))()`)

  record('includes locked visible PPT objects in select same type while preserving transform guards', afterLockedSelect.card2Locked === 'true' && afterLockedSelect.selectedIds.includes('s1-card-2') && afterLockedSelect.deleteDisabled, afterLockedSelect)

  await page.eval(`document.querySelector('[data-ppt-command="unlock-all"]')?.click()`)
  await delay(80)
  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)

  await pressKey(page, {
    code: 'KeyK',
    key: 'k',
    modifiers: 2,
    windowsVirtualKeyCode: 75,
  })
  await delay(80)
  await page.send('Input.insertText', { text: 'same type' })
  await delay(80)

  const paletteDisabled = await page.eval(`(() => ({
    disabled: document.querySelector('[data-ppt-command-palette-item="command:select-same-type"]')?.disabled ?? false,
    itemPresent: !!document.querySelector('[data-ppt-command-palette-item="command:select-same-type"]'),
    open: !!document.querySelector('[data-ppt-command-palette]'),
    selectedCount: document.querySelectorAll('[data-selected="true"]').length,
  }))()`)

  record('disables PPT select same type palette item without selection', paletteDisabled.itemPresent && paletteDisabled.open && paletteDisabled.selectedCount === 0 && paletteDisabled.disabled, paletteDisabled)

  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(80)

  point = await getElementCenter(page, 's1-card-1')
  await clickMouse(page, point.x, point.y, 1)
  await delay(60)
  await pressKey(page, {
    code: 'KeyK',
    key: 'k',
    modifiers: 2,
    windowsVirtualKeyCode: 75,
  })
  await delay(80)
  await page.send('Input.insertText', { text: 'same type' })
  await delay(80)

  const paletteEnabled = await page.eval(`(() => ({
    disabled: document.querySelector('[data-ppt-command-palette-item="command:select-same-type"]')?.disabled ?? true,
    itemPresent: !!document.querySelector('[data-ppt-command-palette-item="command:select-same-type"]'),
  }))()`)

  await pressKey(page, {
    code: 'Enter',
    key: 'Enter',
    windowsVirtualKeyCode: 13,
  })
  await delay(100)

  const afterPaletteSelect = await page.eval(`(() => ({
    open: !!document.querySelector('[data-ppt-command-palette]'),
    selectedIds: [...document.querySelectorAll('[data-selected="true"]')]
      .map((element) => element.getAttribute('data-ppt-element')),
  }))()`)

  record('runs PPT select same type from command palette', paletteEnabled.itemPresent && !paletteEnabled.disabled && !afterPaletteSelect.open && afterPaletteSelect.selectedIds.includes('s1-card-1') && afterPaletteSelect.selectedIds.includes('s1-card-2') && afterPaletteSelect.selectedIds.includes('s1-side-panel'), {
    afterPaletteSelect,
    paletteEnabled,
  })
}

async function runCommandPaletteScenario(page) {
  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)

  const point = await getElementCenter(page, 's1-card-1')
  await clickMouse(page, point.x, point.y, 1)
  await delay(80)

  const initial = await page.eval(`(() => ({
    elementCount: document.querySelectorAll('[data-ppt-element]').length,
    selectedId: document.querySelector('[data-selected="true"]')?.getAttribute('data-ppt-element') ?? '',
  }))()`)

  const paletteButton = await page.eval(`(() => {
    const button = document.querySelector('[data-ppt-command-palette-open]')
    const rect = button?.getBoundingClientRect()

    return {
      exists: !!button,
      x: rect ? rect.left + rect.width / 2 : 0,
      y: rect ? rect.top + rect.height / 2 : 0,
    }
  })()`)
  await clickMouse(page, paletteButton.x, paletteButton.y, 1)
  await delay(100)

  const afterToolbarOpen = await page.eval(`(() => ({
    focusLifecycle: document.querySelector('[data-ppt-command-palette]')?.getAttribute('data-ppt-command-palette-focus-lifecycle') ?? '',
    focusedQuery: document.activeElement?.matches('[data-ppt-command-palette-query]') === true,
    focusTrap: document.querySelector('[data-ppt-command-palette]')?.getAttribute('data-ppt-command-palette-focus-trap') ?? '',
    itemCount: document.querySelectorAll('[data-ppt-command-palette-item]').length,
    model: document.querySelector('[data-ppt-command-palette]')?.getAttribute('data-ppt-command-palette-model') ?? '',
    open: !!document.querySelector('[data-ppt-command-palette]'),
    openerExists: !!document.querySelector('[data-ppt-command-palette-open]'),
    restoreFocus: document.querySelector('[data-ppt-command-palette]')?.getAttribute('data-ppt-command-palette-restore-focus') ?? '',
  }))()`)

  record(
    'opens PPT command palette with focus trap metadata',
    paletteButton.exists &&
      afterToolbarOpen.open &&
      afterToolbarOpen.focusedQuery &&
      afterToolbarOpen.focusLifecycle === 'canvas-modal-focus-lifecycle' &&
      afterToolbarOpen.focusTrap === 'true' &&
      afterToolbarOpen.model === 'canvas-command-palette-items' &&
      afterToolbarOpen.restoreFocus === 'true' &&
      afterToolbarOpen.itemCount > 0,
    {
      afterToolbarOpen,
      paletteButton,
    },
  )

  await pressKey(page, {
    code: 'Tab',
    key: 'Tab',
    modifiers: 8,
    windowsVirtualKeyCode: 9,
  })
  await delay(30)
  const afterShiftTab = await page.eval(`(() => ({
    activeInside: document.activeElement?.closest('[data-ppt-command-palette]') !== null,
    activeItem: document.activeElement?.getAttribute('data-ppt-command-palette-item') ?? '',
    focusedQuery: document.activeElement?.matches('[data-ppt-command-palette-query]') === true,
  }))()`)

  await pressKey(page, {
    code: 'Tab',
    key: 'Tab',
    windowsVirtualKeyCode: 9,
  })
  await delay(30)
  const afterTabWrap = await page.eval(`(() => ({
    activeInside: document.activeElement?.closest('[data-ppt-command-palette]') !== null,
    activeItem: document.activeElement?.getAttribute('data-ppt-command-palette-item') ?? '',
    focusedQuery: document.activeElement?.matches('[data-ppt-command-palette-query]') === true,
  }))()`)

  record(
    'traps Tab focus inside PPT command palette',
    afterShiftTab.activeInside &&
      afterShiftTab.activeItem.length > 0 &&
      !afterShiftTab.focusedQuery &&
      afterTabWrap.activeInside &&
      afterTabWrap.focusedQuery,
    {
      afterShiftTab,
      afterTabWrap,
    },
  )

  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(80)

  const afterToolbarEscape = await page.eval(`(() => ({
    focusedOpener: document.activeElement?.matches('[data-ppt-command-palette-open]') === true,
    open: !!document.querySelector('[data-ppt-command-palette]'),
    selectedId: document.querySelector('[data-selected="true"]')?.getAttribute('data-ppt-element') ?? '',
  }))()`)

  record(
    'restores focus to PPT command palette opener on Escape',
    !afterToolbarEscape.open &&
      afterToolbarEscape.focusedOpener &&
      afterToolbarEscape.selectedId === initial.selectedId,
    afterToolbarEscape,
  )

  await pressKey(page, {
    code: 'KeyK',
    key: 'k',
    modifiers: 2,
    windowsVirtualKeyCode: 75,
  })
  await delay(100)

  const afterOpen = await page.eval(`(() => ({
    activeDescendant: document.querySelector('[data-ppt-command-palette-query]')?.getAttribute('aria-activedescendant') ?? '',
    activeOptionExists: !!document.getElementById(document.querySelector('[data-ppt-command-palette-query]')?.getAttribute('aria-activedescendant') ?? ''),
    activeOptionId: document.querySelector('[data-ppt-command-palette-active="true"]')?.id ?? '',
    activeOptionItem: document.querySelector('[data-ppt-command-palette-active="true"]')?.getAttribute('data-ppt-command-palette-item') ?? '',
    activeOptionRole: document.querySelector('[data-ppt-command-palette-active="true"]')?.getAttribute('role') ?? '',
    activeOptionSelected: document.querySelector('[data-ppt-command-palette-active="true"]')?.getAttribute('aria-selected') ?? '',
    combobox: document.querySelector('[data-ppt-command-palette-query]')?.getAttribute('role') ?? '',
    controls: document.querySelector('[data-ppt-command-palette-query]')?.getAttribute('aria-controls') ?? '',
    expanded: document.querySelector('[data-ppt-command-palette-query]')?.getAttribute('aria-expanded') ?? '',
    focused: document.activeElement?.matches('[data-ppt-command-palette-query]') === true,
    itemCount: document.querySelectorAll('[data-ppt-command-palette-item]').length,
    listboxExists: !!document.getElementById(document.querySelector('[data-ppt-command-palette-query]')?.getAttribute('aria-controls') ?? ''),
    listboxRole: document.getElementById(document.querySelector('[data-ppt-command-palette-query]')?.getAttribute('aria-controls') ?? '')?.getAttribute('role') ?? '',
    model: document.querySelector('[data-ppt-command-palette]')?.getAttribute('data-ppt-command-palette-model') ?? '',
    open: !!document.querySelector('[data-ppt-command-palette]'),
  }))()`)

  record('opens PPT command palette from keyboard shortcut', afterOpen.open && afterOpen.focused && afterOpen.itemCount >= 10, afterOpen)
  record(
    'exposes PPT command palette combobox/listbox active option contract',
    afterOpen.combobox === 'combobox' &&
      afterOpen.expanded === 'true' &&
      afterOpen.controls.length > 0 &&
      afterOpen.listboxExists &&
      afterOpen.listboxRole === 'listbox' &&
      afterOpen.model === 'canvas-command-palette-items' &&
      afterOpen.activeDescendant.length > 0 &&
      afterOpen.activeDescendant === afterOpen.activeOptionId &&
      afterOpen.activeOptionExists &&
      afterOpen.activeOptionRole === 'option' &&
      afterOpen.activeOptionSelected === 'true' &&
      afterOpen.activeOptionItem.length > 0,
    afterOpen,
  )

  await pressKey(page, {
    code: 'ArrowDown',
    key: 'ArrowDown',
    windowsVirtualKeyCode: 40,
  })
  await delay(30)
  const afterComboboxArrowDown = await readPPTCommandPaletteComboboxState(page)

  await pressKey(page, {
    code: 'ArrowUp',
    key: 'ArrowUp',
    windowsVirtualKeyCode: 38,
  })
  await delay(30)
  const afterComboboxArrowUp = await readPPTCommandPaletteComboboxState(page)

  record(
    'updates PPT command palette aria-activedescendant with Arrow keys',
    afterComboboxArrowDown.activeDescendant.length > 0 &&
      afterComboboxArrowDown.activeDescendant === afterComboboxArrowDown.activeOptionId &&
      afterComboboxArrowDown.activeOptionItem !== afterOpen.activeOptionItem &&
      afterComboboxArrowUp.activeDescendant === afterComboboxArrowUp.activeOptionId &&
      afterComboboxArrowUp.activeOptionItem === afterOpen.activeOptionItem,
    {
      afterComboboxArrowDown,
      afterComboboxArrowUp,
      afterOpen,
    },
  )

  await page.send('Input.insertText', { text: 'duplicate' })
  await delay(80)

  const afterFilter = await page.eval(`(() => ({
    firstItem: document.querySelector('[data-ppt-command-palette-item]')?.getAttribute('data-ppt-command-palette-item') ?? '',
    itemIds: [...document.querySelectorAll('[data-ppt-command-palette-item]')]
      .map((item) => item.getAttribute('data-ppt-command-palette-item')),
    itemCount: document.querySelectorAll('[data-ppt-command-palette-item]').length,
  }))()`)

  record('filters PPT command palette items by query', afterFilter.firstItem === 'command:duplicate' && afterFilter.itemIds.includes('slide:duplicate'), afterFilter)

  await pressKey(page, {
    code: 'Enter',
    key: 'Enter',
    windowsVirtualKeyCode: 13,
  })
  await delay(100)

  const afterDuplicate = await page.eval(`(() => ({
    elementCount: document.querySelectorAll('[data-ppt-element]').length,
    open: !!document.querySelector('[data-ppt-command-palette]'),
    selectedCount: document.querySelectorAll('[data-selected="true"]').length,
  }))()`)

  record('runs PPT command palette item with Enter', afterDuplicate.elementCount === initial.elementCount + 1 && afterDuplicate.selectedCount >= 1 && !afterDuplicate.open, {
    afterDuplicate,
    initial,
  })

  await pressKey(page, {
    code: 'KeyK',
    key: 'k',
    modifiers: 2,
    windowsVirtualKeyCode: 75,
  })
  await delay(80)
  await page.send('Input.insertText', { text: 'delete' })
  await pressKey(page, {
    code: 'Enter',
    key: 'Enter',
    windowsVirtualKeyCode: 13,
  })
  await delay(100)

  const afterCleanupDelete = await page.eval(`(() => ({
    elementCount: document.querySelectorAll('[data-ppt-element]').length,
    open: !!document.querySelector('[data-ppt-command-palette]'),
  }))()`)

  record('cleans up PPT command palette duplicate through delete command', afterCleanupDelete.elementCount === initial.elementCount && !afterCleanupDelete.open, afterCleanupDelete)

  const pointAfterDelete = await getElementCenter(page, 's1-card-1')
  await clickMouse(page, pointAfterDelete.x, pointAfterDelete.y, 1)
  await delay(50)
  await pressKey(page, {
    code: 'KeyK',
    key: 'k',
    modifiers: 2,
    windowsVirtualKeyCode: 75,
  })
  await delay(80)
  await page.send('Input.insertText', { text: 'group' })
  await delay(80)

  const beforeDisabled = await page.eval(`(() => ({
    activeDescendant: document.querySelector('[data-ppt-command-palette-query]')?.getAttribute('aria-activedescendant') ?? '',
    groupAriaDisabled: document.querySelector('[data-ppt-command-palette-item="command:group"]')?.getAttribute('aria-disabled') ?? '',
    groupDisabled: document.querySelector('[data-ppt-command-palette-item="command:group"]')?.disabled ?? false,
    groupOptionId: document.querySelector('[data-ppt-command-palette-item="command:group"]')?.id ?? '',
    groupId: document.querySelector('[data-selected="true"]')?.getAttribute('data-group-id') ?? null,
    selectedCount: document.querySelectorAll('[data-selected="true"]').length,
  }))()`)

  await pressKey(page, {
    code: 'Enter',
    key: 'Enter',
    windowsVirtualKeyCode: 13,
  })
  await delay(80)

  const afterDisabled = await page.eval(`(() => ({
    groupId: document.querySelector('[data-selected="true"]')?.getAttribute('data-group-id') ?? null,
    open: !!document.querySelector('[data-ppt-command-palette]'),
    selectedCount: document.querySelectorAll('[data-selected="true"]').length,
  }))()`)

  record(
    'does not run disabled PPT command palette item',
    beforeDisabled.groupDisabled &&
      beforeDisabled.groupAriaDisabled === 'true' &&
      beforeDisabled.activeDescendant === beforeDisabled.groupOptionId &&
      beforeDisabled.selectedCount === 1 &&
      afterDisabled.selectedCount === 1 &&
      afterDisabled.groupId === beforeDisabled.groupId &&
      afterDisabled.open,
    {
      afterDisabled,
      beforeDisabled,
    },
  )

  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(80)

  await pressKey(page, {
    code: 'KeyK',
    key: 'k',
    modifiers: 2,
    windowsVirtualKeyCode: 75,
  })
  await delay(80)

  const alignIds = await readCommandPaletteIds(page, 'align')
  const toolIds = await readCommandPaletteIds(page, 'tool')
  const arrowToolIds = await readCommandPaletteIds(page, 'arrow tool')
  const panToolIds = await readCommandPaletteIds(page, 'pan tool')
  const laserToolIds = await readCommandPaletteIds(page, 'laser pointer')
  const stickyToolIds = await readCommandPaletteIds(page, 'sticky')
  const sectionToolIds = await readCommandPaletteIds(page, 'section tool')
  const penToolIds = await readCommandPaletteIds(page, 'pen tool')
  const markerToolIds = await readCommandPaletteIds(page, 'marker')
  const highlighterToolIds = await readCommandPaletteIds(page, 'highlighter')
  const eraserToolIds = await readCommandPaletteIds(page, 'eraser')
  const findIds = await readCommandPaletteIds(page, 'find')
  const groupIds = await readCommandPaletteIds(page, 'group')
  const lockIds = await readCommandPaletteIds(page, 'lock')
  const frontIds = await readCommandPaletteIds(page, 'front')
  const backIds = await readCommandPaletteIds(page, 'back')
  const tidyIds = await readCommandPaletteIds(page, 'tidy')
  const flipIds = await readCommandPaletteIds(page, 'flip')
  const fitIds = await readCommandPaletteIds(page, 'fit')
  const resetZoomIds = await readCommandPaletteIds(page, 'Cmd/Ctrl+0')
  const zoomInShortcutIds = await readCommandPaletteIds(page, 'Cmd/Ctrl+=')
  const zoomOutShortcutIds = await readCommandPaletteIds(page, 'Cmd/Ctrl+-')
  const fitSlideShortcutIds = await readCommandPaletteIds(page, '0')
  const fitSelectionShortcutIds = await readCommandPaletteIds(page, '1')
  const gridIds = await readCommandPaletteIds(page, 'grid')
  const guideIds = await readCommandPaletteIds(page, 'guide')
  const presentIds = await readCommandPaletteIds(page, 'present')
  const bringForwardShortcutIds = await readCommandPaletteIds(page, 'Cmd/Ctrl+]')
  const bringToFrontShortcutIds = await readCommandPaletteIds(page, 'Shift+Cmd/Ctrl+]')
  const sendBackwardShortcutIds = await readCommandPaletteIds(page, 'Cmd/Ctrl+[')
  const sendToBackShortcutIds = await readCommandPaletteIds(page, 'Shift+Cmd/Ctrl+[')
  const lockShortcutIds = await readCommandPaletteIds(page, 'Cmd/Ctrl+L')
  const unlockShortcutIds = await readCommandPaletteIds(page, 'Shift+Cmd/Ctrl+L')
  const exposed = {
    hasAlign: alignIds.includes('command:align-left'),
    hasCreate: toolIds.includes('tool:text') &&
      arrowToolIds.includes('tool:arrow') &&
      panToolIds.includes('tool:pan') &&
      laserToolIds.includes('tool:laser') &&
      stickyToolIds.includes('tool:sticky') &&
      sectionToolIds.includes('tool:section') &&
      penToolIds.includes('tool:pen') &&
      markerToolIds.includes('tool:marker') &&
      highlighterToolIds.includes('tool:highlight') &&
      eraserToolIds.includes('tool:eraser'),
    hasFind: findIds.includes('view:find'),
    hasFlip: flipIds.includes('command:flip-horizontal') &&
      flipIds.includes('command:flip-vertical'),
    hasGroup: groupIds.includes('command:group') && groupIds.includes('command:ungroup'),
    hasLock: lockIds.includes('command:lock-selection') && lockIds.includes('command:unlock-all'),
    hasLockShortcuts: lockShortcutIds.includes('command:lock-selection') &&
      unlockShortcutIds.includes('command:unlock-all'),
    hasReorder: frontIds.includes('command:bring-to-front') && backIds.includes('command:send-to-back'),
    hasReorderShortcuts: bringForwardShortcutIds.includes('command:bring-forward') &&
      bringToFrontShortcutIds.includes('command:bring-to-front') &&
      sendBackwardShortcutIds.includes('command:send-backward') &&
      sendToBackShortcutIds.includes('command:send-to-back'),
    hasTidy: tidyIds.includes('command:tidy-selection'),
    hasView: fitIds.includes('view:fit-slide') &&
      fitIds.includes('view:fit-selection') &&
      resetZoomIds.includes('view:reset-zoom') &&
      zoomInShortcutIds.includes('view:zoom-in') &&
      zoomOutShortcutIds.includes('view:zoom-out') &&
      gridIds.includes('view:toggle-grid') &&
      guideIds.includes('view:toggle-frame-guides') &&
      presentIds.includes('view:present'),
    hasViewportShortcuts: fitSlideShortcutIds.includes('view:fit-slide') &&
      fitSelectionShortcutIds.includes('view:fit-selection') &&
      resetZoomIds.includes('view:reset-zoom') &&
      zoomInShortcutIds.includes('view:zoom-in') &&
      zoomOutShortcutIds.includes('view:zoom-out'),
    shortcuts: {
      bringForward: bringForwardShortcutIds,
      bringToFront: bringToFrontShortcutIds,
      fitSelection: fitSelectionShortcutIds,
      fitSlide: fitSlideShortcutIds,
      lockSelection: lockShortcutIds,
      resetZoom: resetZoomIds,
      sendBackward: sendBackwardShortcutIds,
      sendToBack: sendToBackShortcutIds,
      unlockAll: unlockShortcutIds,
      zoomIn: zoomInShortcutIds,
      zoomOut: zoomOutShortcutIds,
    },
    visibleCounts: {
      align: alignIds.length,
      back: backIds.length,
      find: findIds.length,
      fit: fitIds.length,
      flip: flipIds.length,
      front: frontIds.length,
      grid: gridIds.length,
      guide: guideIds.length,
      group: groupIds.length,
      eraserTool: eraserToolIds.length,
      arrowTool: arrowToolIds.length,
      highlighterTool: highlighterToolIds.length,
      lock: lockIds.length,
      laserTool: laserToolIds.length,
      markerTool: markerToolIds.length,
      present: presentIds.length,
      panTool: panToolIds.length,
      sectionTool: sectionToolIds.length,
      stickyTool: stickyToolIds.length,
      penTool: penToolIds.length,
      tidy: tidyIds.length,
      tool: toolIds.length,
    },
  }

  record('exposes PPT create view and arrange commands in command palette', exposed.hasAlign && exposed.hasCreate && exposed.hasFind && exposed.hasFlip && exposed.hasGroup && exposed.hasLock && exposed.hasLockShortcuts && exposed.hasReorder && exposed.hasReorderShortcuts && exposed.hasTidy && exposed.hasView && exposed.hasViewportShortcuts, exposed)

  const guideToggleIds = await readCommandPaletteIds(page, 'frame guides')
  await pressKey(page, {
    code: 'Enter',
    key: 'Enter',
    windowsVirtualKeyCode: 13,
  })
  await delay(80)

  const afterPaletteGuideToggle = await page.eval(`(() => ({
    frameGuideAttr: document.querySelector('.ppt-stage-shell')?.getAttribute('data-frame-guides'),
    frameGuideCount: document.querySelectorAll('[data-ppt-frame-guides]').length,
    itemPresent: ${JSON.stringify(guideToggleIds.includes('view:toggle-frame-guides'))},
    open: !!document.querySelector('[data-ppt-command-palette]'),
    pressed: document.querySelector('[data-ppt-view-frame-guides]')?.getAttribute('aria-pressed'),
  }))()`)

  record('toggles PPT frame guides from command palette', afterPaletteGuideToggle.itemPresent && !afterPaletteGuideToggle.open && afterPaletteGuideToggle.frameGuideAttr === 'false' && afterPaletteGuideToggle.pressed === 'false' && afterPaletteGuideToggle.frameGuideCount === 0, afterPaletteGuideToggle)

  await page.eval(`document.querySelector('[data-ppt-view-frame-guides]')?.click()`)
  await delay(50)

  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(80)

  const titlePoint = await getElementCenter(page, 's1-title')
  await clickMouse(page, titlePoint.x, titlePoint.y, 2)
  await delay(80)
  await page.eval(`(() => {
    const editor = document.querySelector('[data-ppt-element="s1-title"] .ppt-element-editor')

    editor?.dispatchEvent(new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      code: 'KeyK',
      ctrlKey: true,
      key: 'k',
    }))
  })()`)
  await delay(80)

  const afterNativeGuard = await page.eval(`(() => {
    const editor = document.querySelector('[data-ppt-element="s1-title"] .ppt-element-editor')

    return {
      editing: editor?.isContentEditable === true && document.activeElement === editor,
      open: !!document.querySelector('[data-ppt-command-palette]'),
    }
  })()`)

  record('does not open PPT command palette while native text editing is active', afterNativeGuard.editing && !afterNativeGuard.open, afterNativeGuard)

  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)
}

async function runShortcutHelpScenario(page) {
  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)

  await pressKey(page, {
    code: 'Slash',
    key: '?',
    modifiers: 8,
    windowsVirtualKeyCode: 191,
  })
  await delay(100)

  const afterShortcutOpen = await page.eval(`(() => {
    const sectionNames = [...document.querySelectorAll('[data-ppt-shortcut-help-section]')]
      .map((section) => section.getAttribute('data-ppt-shortcut-help-section'))
    const itemIds = [...document.querySelectorAll('[data-ppt-shortcut-help-item]')]
      .map((item) => item.getAttribute('data-ppt-shortcut-help-item'))
    const shortcuts = [...document.querySelectorAll('[data-ppt-shortcut-help-shortcut]')]
      .map((item) => item.getAttribute('data-ppt-shortcut-help-shortcut'))

    return {
      closeFocused: document.activeElement?.matches('[data-ppt-shortcut-help-close]') === true,
      focusLifecycle: document.querySelector('[data-ppt-shortcut-help]')?.getAttribute('data-ppt-shortcut-help-focus-lifecycle') ?? '',
      itemCount: itemIds.length,
      itemIds,
      open: !!document.querySelector('[data-ppt-shortcut-help]'),
      sectionNames,
      shortcuts,
    }
  })()`)

  record('opens PPT keyboard shortcut help from Shift+/ shortcut', afterShortcutOpen.open && afterShortcutOpen.closeFocused && afterShortcutOpen.focusLifecycle === 'canvas-modal-focus-lifecycle' && afterShortcutOpen.itemCount >= 12, afterShortcutOpen)
  record('groups PPT keyboard shortcut help items by command section', afterShortcutOpen.sectionNames.includes('Create') && afterShortcutOpen.sectionNames.includes('Edit') && afterShortcutOpen.sectionNames.includes('Arrange') && afterShortcutOpen.sectionNames.includes('View') && afterShortcutOpen.sectionNames.includes('Format'), afterShortcutOpen)
  record('derives PPT keyboard shortcut help from command palette shortcuts', afterShortcutOpen.itemIds.includes('system:keyboard-shortcuts') && afterShortcutOpen.itemIds.includes('command:duplicate') && afterShortcutOpen.itemIds.includes('command:bring-forward') && afterShortcutOpen.itemIds.includes('command:lock-selection') && afterShortcutOpen.itemIds.includes('view:fit-slide') && afterShortcutOpen.itemIds.includes('view:reset-zoom') && afterShortcutOpen.itemIds.includes('view:zoom-in') && afterShortcutOpen.itemIds.includes('tool:pan') && afterShortcutOpen.itemIds.includes('tool:laser') && afterShortcutOpen.itemIds.includes('tool:text') && afterShortcutOpen.itemIds.includes('tool:sticky') && afterShortcutOpen.itemIds.includes('tool:section') && afterShortcutOpen.itemIds.includes('tool:arrow') && afterShortcutOpen.itemIds.includes('tool:marker') && afterShortcutOpen.itemIds.includes('tool:highlight') && afterShortcutOpen.itemIds.includes('tool:eraser') && afterShortcutOpen.itemIds.includes('format:bold') && afterShortcutOpen.shortcuts.includes('Shift+/') && afterShortcutOpen.shortcuts.includes('Cmd/Ctrl+D') && afterShortcutOpen.shortcuts.includes('Cmd/Ctrl+]') && afterShortcutOpen.shortcuts.includes('Shift+Cmd/Ctrl+L') && afterShortcutOpen.shortcuts.includes('0') && afterShortcutOpen.shortcuts.includes('1') && afterShortcutOpen.shortcuts.includes('Cmd/Ctrl+0') && afterShortcutOpen.shortcuts.includes('Cmd/Ctrl+=') && afterShortcutOpen.shortcuts.includes('H') && afterShortcutOpen.shortcuts.includes('P') && afterShortcutOpen.shortcuts.includes('S') && afterShortcutOpen.shortcuts.includes('Shift+S') && afterShortcutOpen.shortcuts.includes('L') && afterShortcutOpen.shortcuts.includes('M') && afterShortcutOpen.shortcuts.includes('Shift+M') && afterShortcutOpen.shortcuts.includes('E'), afterShortcutOpen)

  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(80)

  const afterEscapeClose = await page.eval(`(() => ({
    open: !!document.querySelector('[data-ppt-shortcut-help]'),
  }))()`)

  record('closes PPT keyboard shortcut help with Escape', !afterEscapeClose.open, afterEscapeClose)

  await pressKey(page, {
    code: 'KeyK',
    key: 'k',
    modifiers: 2,
    windowsVirtualKeyCode: 75,
  })
  await delay(80)
  await page.send('Input.insertText', { text: 'keyboard' })
  await delay(80)

  const beforePaletteRun = await page.eval(`(() => ({
    itemPresent: !!document.querySelector('[data-ppt-command-palette-item="system:keyboard-shortcuts"]'),
    open: !!document.querySelector('[data-ppt-command-palette]'),
  }))()`)

  await pressKey(page, {
    code: 'Enter',
    key: 'Enter',
    windowsVirtualKeyCode: 13,
  })
  await delay(100)

  const afterPaletteRun = await page.eval(`(() => ({
    helpOpen: !!document.querySelector('[data-ppt-shortcut-help]'),
    paletteOpen: !!document.querySelector('[data-ppt-command-palette]'),
    systemItem: !!document.querySelector('[data-ppt-shortcut-help-item="system:keyboard-shortcuts"]'),
  }))()`)

  record('opens PPT keyboard shortcut help from command palette entry', beforePaletteRun.open && beforePaletteRun.itemPresent && afterPaletteRun.helpOpen && !afterPaletteRun.paletteOpen && afterPaletteRun.systemItem, {
    afterPaletteRun,
    beforePaletteRun,
  })

  await clickMouse(page, 12, 12, 1)
  await delay(80)

  const afterBackdropClose = await page.eval(`(() => ({
    open: !!document.querySelector('[data-ppt-shortcut-help]'),
  }))()`)

  record('closes PPT keyboard shortcut help with backdrop click', !afterBackdropClose.open, afterBackdropClose)

  const titlePoint = await getElementCenter(page, 's1-title')
  await clickMouse(page, titlePoint.x, titlePoint.y, 2)
  await delay(80)
  await page.eval(`(() => {
    const editor = document.querySelector('[data-ppt-element="s1-title"] .ppt-element-editor')

    editor?.dispatchEvent(new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      code: 'Slash',
      key: '?',
      shiftKey: true,
    }))
  })()`)
  await delay(80)

  const afterNativeGuard = await page.eval(`(() => {
    const editor = document.querySelector('[data-ppt-element="s1-title"] .ppt-element-editor')

    return {
      editing: editor?.isContentEditable === true && document.activeElement === editor,
      open: !!document.querySelector('[data-ppt-shortcut-help]'),
    }
  })()`)

  record('does not open PPT keyboard shortcut help while native text editing is active', afterNativeGuard.editing && !afterNativeGuard.open, afterNativeGuard)

  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)
}

async function runSlideMetadataScenario(page) {
  const titlePoint = await getElementCenter(page, 's1-title')
  await clickMouse(page, titlePoint.x, titlePoint.y, 1)
  await delay(50)

  const initial = await getPPTSlideMetadataState(page)
  const initialTabs = await getPPTInspectorTabsState(page)

  record(
    'renders PPT slide metadata inspector descriptor',
    initial.inspector &&
      initial.surface === 'slide-metadata-inspector' &&
      initial.slideId === 'slide-1' &&
      initial.slideCount >= 2 &&
      initial.fieldCount === 5 &&
      ['name', 'background', 'notes', 'size', 'orientation'].every((field) => initial.fields.includes(field)) &&
      initial.sizeValue === '1280x720' &&
      initial.orientationValue === 'landscape' &&
      initial.editableByField.name === 'true' &&
      initial.editableByField.background === 'true' &&
      initial.editableByField.notes === 'true' &&
      initial.editableByField.size === 'false' &&
      initial.editableByField.orientation === 'false',
    initial,
  )
  record(
    'keeps PPT object inspector priority over slide metadata inspector',
    initial.inspectorSurface === 'object-selection-inspector' &&
      initial.objectActive === 'true' &&
      initial.objectPriority === '0' &&
      initial.slidePriority === 'secondary',
    initial,
  )
  record(
    'exposes PPT inspector APG tabs contract',
    initialTabs.tablistRole === 'tablist' &&
      initialTabs.model === 'canvas-tabs-roving-focus' &&
      initialTabs.activation === 'automatic' &&
      initialTabs.keyboard === 'arrow-home-end-enter-space' &&
      initialTabs.tabCount === 2 &&
      initialTabs.panelCount === 2 &&
      initialTabs.selectedTabIds.length === 1 &&
      initialTabs.tabStopIds.length === 1 &&
      initialTabs.relationshipsValid &&
      initialTabs.activePanelIds.length === 1 &&
      initialTabs.activePanelIds[0] === initialTabs.selectedTabIds[0],
    initialTabs,
  )

  await page.eval(`document.querySelector('[data-ppt-inspector-tab="slide"]')?.focus()`)
  await pressKey(page, {
    code: 'ArrowRight',
    key: 'ArrowRight',
    windowsVirtualKeyCode: 39,
  })
  await delay(50)

  const afterTabArrowRight = await getPPTInspectorTabsState(page)

  await pressKey(page, {
    code: 'Home',
    key: 'Home',
    windowsVirtualKeyCode: 36,
  })
  await delay(50)

  const afterTabHome = await getPPTInspectorTabsState(page)

  await pressKey(page, {
    code: 'End',
    key: 'End',
    windowsVirtualKeyCode: 35,
  })
  await delay(50)

  const afterTabEnd = await getPPTInspectorTabsState(page)

  record(
    'moves PPT inspector tab focus with Arrow and Home End keys',
    afterTabArrowRight.activeTab === 'selection' &&
      afterTabArrowRight.focusedTab === 'selection' &&
      afterTabHome.activeTab === 'slide' &&
      afterTabHome.focusedTab === 'slide' &&
      afterTabEnd.activeTab === 'selection' &&
      afterTabEnd.focusedTab === 'selection',
    {
      afterTabArrowRight,
      afterTabEnd,
      afterTabHome,
    },
  )

  await page.eval(`document.querySelector('[data-ppt-inspector-tab="slide"]')?.focus()`)
  await pressKey(page, {
    code: 'Space',
    key: ' ',
    windowsVirtualKeyCode: 32,
  })
  await delay(50)

  const afterTabSpace = await getPPTInspectorTabsState(page)

  await page.eval(`document.querySelector('[data-ppt-inspector-tab="selection"]')?.focus()`)
  await pressKey(page, {
    code: 'Enter',
    key: 'Enter',
    windowsVirtualKeyCode: 13,
  })
  await delay(50)

  const afterTabEnter = await getPPTInspectorTabsState(page)

  record(
    'activates PPT inspector tabs with Enter and Space',
    afterTabSpace.activeTab === 'slide' &&
      afterTabSpace.focusedTab === 'slide' &&
      afterTabEnter.activeTab === 'selection' &&
      afterTabEnter.focusedTab === 'selection',
    {
      afterTabEnter,
      afterTabSpace,
    },
  )

  await page.eval(`(() => {
    const name = document.querySelector('[data-ppt-slide-field="name"]')
    const background = document.querySelector('[data-ppt-slide-field="background"]')
    const notes = document.querySelector('[data-ppt-slide-field="notes"]')
    const inputSetter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set
    const textAreaSetter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set

    inputSetter.call(name, 'Metadata Retouch')
    name.dispatchEvent(new Event('input', { bubbles: true }))
    name.dispatchEvent(new Event('change', { bubbles: true }))

    inputSetter.call(background, '#dbeafe')
    background.dispatchEvent(new Event('input', { bubbles: true }))
    background.dispatchEvent(new Event('change', { bubbles: true }))

    textAreaSetter.call(notes, 'Metadata cue from slide inspector.')
    notes.dispatchEvent(new Event('input', { bubbles: true }))
    notes.dispatchEvent(new Event('change', { bubbles: true }))
  })()`)
  await delay(80)

  const afterUpdate = await getPPTSlideMetadataState(page)

  record(
    'updates PPT slide metadata through slide-command-effect projection',
    afterUpdate.commandSlot === 'command-effect' &&
      afterUpdate.commands.includes('update-slide-name') &&
      afterUpdate.commands.includes('update-slide-background') &&
      afterUpdate.commands.includes('update-slide-notes') &&
      afterUpdate.name === 'Metadata Retouch' &&
      afterUpdate.thumbName.includes('Metadata Retouch') &&
      afterUpdate.notes === 'Metadata cue from slide inspector.' &&
      afterUpdate.slideBackground === 'rgb(219, 234, 254)',
    {
      afterUpdate,
      initial,
    },
  )

  await page.eval(`(() => {
    const name = document.querySelector('[data-ppt-slide-field="name"]')
    const background = document.querySelector('[data-ppt-slide-field="background"]')
    const notes = document.querySelector('[data-ppt-slide-field="notes"]')
    const inputSetter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set
    const textAreaSetter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set

    inputSetter.call(name, ${JSON.stringify(initial.name)})
    name.dispatchEvent(new Event('input', { bubbles: true }))
    name.dispatchEvent(new Event('change', { bubbles: true }))

    inputSetter.call(background, ${JSON.stringify(initial.backgroundValue)})
    background.dispatchEvent(new Event('input', { bubbles: true }))
    background.dispatchEvent(new Event('change', { bubbles: true }))

    textAreaSetter.call(notes, ${JSON.stringify(initial.notes)})
    notes.dispatchEvent(new Event('input', { bubbles: true }))
    notes.dispatchEvent(new Event('change', { bubbles: true }))
  })()`)
  await delay(80)

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      background: { color: '#f8fafc' },
      name: 'AI Metadata Slide',
      notes: 'AI metadata note for final retouch.',
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(100)

  const afterMetadataPaste = await getPPTSlideMetadataState(page)

  record(
    'pastes JSON slide metadata into active PPT slide',
    afterMetadataPaste.slideMetadataImportModel === 'ppt-slide-metadata-import' &&
      afterMetadataPaste.slideMetadataImportFormat === 'application-json-ppt-slide-metadata' &&
      afterMetadataPaste.slideMetadataImportSlide === 'slide-1' &&
      afterMetadataPaste.slideMetadataImportFields === 'name background notes' &&
      afterMetadataPaste.slideMetadataImportCommands ===
        'update-slide-name update-slide-background update-slide-notes' &&
      afterMetadataPaste.slideMetadataImportName === 'AI Metadata Slide' &&
      afterMetadataPaste.slideMetadataImportBackground === '#f8fafc' &&
      afterMetadataPaste.slideMetadataImportNotesLength > 0 &&
      afterMetadataPaste.slideMetadataImportJsonLength > 100 &&
      afterMetadataPaste.name === 'AI Metadata Slide' &&
      afterMetadataPaste.thumbName.includes('AI Metadata Slide') &&
      afterMetadataPaste.notes === 'AI metadata note for final retouch.' &&
      afterMetadataPaste.slideBackground === 'rgb(248, 250, 252)',
    {
      afterMetadataPaste,
    },
  )

  await page.eval(`(() => {
    const name = document.querySelector('[data-ppt-slide-field="name"]')
    const background = document.querySelector('[data-ppt-slide-field="background"]')
    const notes = document.querySelector('[data-ppt-slide-field="notes"]')
    const inputSetter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set
    const textAreaSetter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set

    inputSetter.call(name, ${JSON.stringify(initial.name)})
    name.dispatchEvent(new Event('input', { bubbles: true }))
    name.dispatchEvent(new Event('change', { bubbles: true }))

    inputSetter.call(background, ${JSON.stringify(initial.backgroundValue)})
    background.dispatchEvent(new Event('input', { bubbles: true }))
    background.dispatchEvent(new Event('change', { bubbles: true }))

    textAreaSetter.call(notes, ${JSON.stringify(initial.notes)})
    notes.dispatchEvent(new Event('input', { bubbles: true }))
    notes.dispatchEvent(new Event('change', { bubbles: true }))
  })()`)
  await delay(80)

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      notes: 'AI notes from JSON clipboard.\\nConfirm the ask before presenting.',
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(100)

  const afterJSONNotesPaste = await getPPTSlideMetadataState(page)

  record(
    'pastes JSON speaker notes into active PPT slide metadata',
    afterJSONNotesPaste.slideNotesImportModel === 'ppt-slide-notes-import' &&
      afterJSONNotesPaste.slideNotesImportFormat === 'application-json-ppt-notes' &&
      afterJSONNotesPaste.slideNotesImportSlide === 'slide-1' &&
      afterJSONNotesPaste.notes.includes('AI notes from JSON clipboard.') &&
      afterJSONNotesPaste.notes.includes('Confirm the ask before presenting.') &&
      afterJSONNotesPaste.slideNotesImportNotesLength ===
        afterJSONNotesPaste.notes.length &&
      afterJSONNotesPaste.slideNotesImportTextLength > afterJSONNotesPaste.notes.length,
    {
      afterJSONNotesPaste,
    },
  )

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const markdown = [
      'Speaker notes:',
      '- Rehearse the imported AI draft.',
      '- Leave the final wording editable.',
    ].join('\\n')

    dataTransfer.setData('text/markdown', markdown)
    dataTransfer.setData('text/plain', markdown)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(100)

  const afterMarkdownNotesPaste = await getPPTSlideMetadataState(page)

  record(
    'pastes Markdown speaker notes into active PPT slide metadata',
    afterMarkdownNotesPaste.slideNotesImportModel === 'ppt-slide-notes-import' &&
      afterMarkdownNotesPaste.slideNotesImportFormat === 'text-markdown-ppt-notes' &&
      afterMarkdownNotesPaste.slideNotesImportSlide === 'slide-1' &&
      afterMarkdownNotesPaste.notes.includes('Rehearse the imported AI draft.') &&
      afterMarkdownNotesPaste.notes.includes('Leave the final wording editable.') &&
      afterMarkdownNotesPaste.slideNotesImportNotesLength ===
        afterMarkdownNotesPaste.notes.length &&
      afterMarkdownNotesPaste.slideNotesImportTextLength > afterMarkdownNotesPaste.notes.length,
    {
      afterMarkdownNotesPaste,
    },
  )

  await page.eval(`(() => {
    const notes = document.querySelector('[data-ppt-slide-field="notes"]')
    const textAreaSetter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set

    textAreaSetter.call(notes, ${JSON.stringify(initial.notes)})
    notes.dispatchEvent(new Event('input', { bubbles: true }))
    notes.dispatchEvent(new Event('change', { bubbles: true }))
  })()`)
  await delay(80)
}

async function runThemeScenario(page) {
  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)

  const initial = await readPPTThemeState(page)

  await page.eval(`document.querySelector('[data-ppt-theme-toggle]')?.click()`)
  await delay(80)

  const afterToolbarToggle = await readPPTThemeState(page)

  record('toggles PPT dark mode from toolbar', initial.theme === 'light' && afterToolbarToggle.theme === 'dark' && afterToolbarToggle.themePressed === 'true' && afterToolbarToggle.topbarBg !== initial.topbarBg && afterToolbarToggle.inspectorBg !== initial.inspectorBg && afterToolbarToggle.stageBg !== initial.stageBg, {
    afterToolbarToggle,
    initial,
  })
  record('keeps PPT slide model background while toggling dark mode', afterToolbarToggle.slideBg === initial.slideBg && afterToolbarToggle.slideInlineBg === initial.slideInlineBg, {
    afterToolbarToggle,
    initial,
  })

  await pressKey(page, {
    code: 'KeyK',
    key: 'k',
    modifiers: 2,
    windowsVirtualKeyCode: 75,
  })
  await delay(80)
  await page.send('Input.insertText', { text: 'light theme' })
  await delay(80)

  const paletteState = await page.eval(`(() => ({
    itemPresent: !!document.querySelector('[data-ppt-command-palette-item="view:toggle-theme"]'),
    open: !!document.querySelector('[data-ppt-command-palette]'),
  }))()`)

  await pressKey(page, {
    code: 'Enter',
    key: 'Enter',
    windowsVirtualKeyCode: 13,
  })
  await delay(80)

  const afterPaletteToggle = await readPPTThemeState(page)

  record('toggles PPT theme from command palette', paletteState.open && paletteState.itemPresent && afterPaletteToggle.theme === 'light' && afterPaletteToggle.themePressed === 'false' && afterPaletteToggle.topbarBg === initial.topbarBg && afterPaletteToggle.stageBg === initial.stageBg, {
    afterPaletteToggle,
    initial,
    paletteState,
  })
}

async function readPPTThemeState(page) {
  return page.eval(`(() => {
    const app = document.querySelector('[data-ppt-app]')
    const topbar = document.querySelector('.ppt-topbar')
    const inspector = document.querySelector('.ppt-inspector')
    const stage = document.querySelector('.ppt-stage-shell')
    const slide = document.querySelector('.ppt-slide')

    return {
      appBg: getComputedStyle(app).backgroundColor,
      inspectorBg: getComputedStyle(inspector).backgroundColor,
      slideBg: getComputedStyle(slide).backgroundColor,
      slideInlineBg: slide?.style.background ?? '',
      stageBg: getComputedStyle(stage).backgroundColor,
      theme: app?.getAttribute('data-theme') ?? '',
      themePressed: document.querySelector('[data-ppt-theme-toggle]')?.getAttribute('aria-pressed') ?? '',
      topbarBg: getComputedStyle(topbar).backgroundColor,
    }
  })()`)
}

async function runFitSelectionScenario(page) {
  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)

  await page.eval(`document.querySelector('[data-ppt-view-fit-slide]')?.click()`)
  await delay(80)

  const afterFitSlide = await readViewportState(page)
  const point = await getElementCenter(page, 's1-card-1')
  await clickMouse(page, point.x, point.y, 1)
  await delay(80)

  const beforeToolbarFit = await page.eval(`(() => ({
    fitSelectionDisabled: document.querySelector('[data-ppt-view-fit-selection]')?.disabled ?? true,
    selectedCount: document.querySelectorAll('[data-selected="true"]').length,
  }))()`)

  await page.eval(`document.querySelector('[data-ppt-view-fit-selection]')?.click()`)
  await delay(120)

  const afterToolbarFit = await readViewportState(page)

  record('fits PPT viewport to selected object from toolbar', beforeToolbarFit.selectedCount === 1 && !beforeToolbarFit.fitSelectionDisabled && afterToolbarFit.scale > afterFitSlide.scale, {
    afterFitSlide,
    afterToolbarFit,
    beforeToolbarFit,
  })

  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)
  await pressKey(page, {
    code: 'KeyK',
    key: 'k',
    modifiers: 2,
    windowsVirtualKeyCode: 75,
  })
  await delay(80)
  await page.send('Input.insertText', { text: 'fit selection' })
  await delay(80)

  const paletteDisabled = await page.eval(`(() => ({
    disabled: document.querySelector('[data-ppt-command-palette-item="view:fit-selection"]')?.disabled ?? false,
    fitIds: [...document.querySelectorAll('[data-ppt-command-palette-item]')]
      .map((item) => item.getAttribute('data-ppt-command-palette-item')),
    open: !!document.querySelector('[data-ppt-command-palette]'),
    selectedCount: document.querySelectorAll('[data-selected="true"]').length,
  }))()`)

  record('disables PPT fit selection command palette item without selection', paletteDisabled.open && paletteDisabled.selectedCount === 0 && paletteDisabled.fitIds.includes('view:fit-selection') && paletteDisabled.disabled, paletteDisabled)

  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(80)
  await page.eval(`document.querySelector('[data-ppt-view-fit-slide]')?.click()`)
  await delay(80)

  await selectPPTLayerRows(page, ['s1-card-1', 's1-side-panel'])
  await delay(80)

  const beforePaletteFit = await readViewportState(page)
  await pressKey(page, {
    code: 'KeyK',
    key: 'k',
    modifiers: 2,
    windowsVirtualKeyCode: 75,
  })
  await delay(80)
  const fitQueryIds = await readCommandPaletteIds(page, 'fit')
  await readCommandPaletteIds(page, 'fit selection')

  const paletteEnabled = await page.eval(`(() => ({
    disabled: document.querySelector('[data-ppt-command-palette-item="view:fit-selection"]')?.disabled ?? true,
    fitIds: [...document.querySelectorAll('[data-ppt-command-palette-item]')]
      .map((item) => item.getAttribute('data-ppt-command-palette-item')),
    selectedCount: document.querySelectorAll('[data-selected="true"]').length,
  }))()`)

  await pressKey(page, {
    code: 'Enter',
    key: 'Enter',
    windowsVirtualKeyCode: 13,
  })
  await delay(120)

  const afterPaletteFit = await readViewportState(page)

  record('fits PPT viewport to multi-selection from command palette', fitQueryIds.includes('view:fit-slide') && fitQueryIds.includes('view:fit-selection') && paletteEnabled.fitIds.includes('view:fit-selection') && !paletteEnabled.disabled && paletteEnabled.selectedCount >= 2 && afterPaletteFit.scale > beforePaletteFit.scale && !afterPaletteFit.paletteOpen, {
    afterPaletteFit,
    beforePaletteFit,
    fitQueryIds,
    paletteEnabled,
  })

  await page.eval(`document.querySelector('[data-ppt-view-fit-slide]')?.click()`)
  await delay(80)

  const afterRestoreFitSlide = await readViewportState(page)

  record('keeps PPT fit slide command on full slide bounds', afterRestoreFitSlide.scale < afterPaletteFit.scale && afterRestoreFitSlide.label !== afterPaletteFit.label, {
    afterPaletteFit,
    afterRestoreFitSlide,
  })

  await pressKey(page, {
    code: 'Digit1',
    key: '1',
    windowsVirtualKeyCode: 49,
  })
  await delay(120)
  const afterShortcutFitSelection = await readViewportState(page)

  await page.eval(`document.activeElement instanceof HTMLElement && document.activeElement.blur()`)
  const beforeTemporaryPan = await readPPTTemporaryPanState(page)
  const panStart = await page.eval(`(() => {
    const rect = document.querySelector('.ppt-stage-shell')?.getBoundingClientRect()

    return {
      x: (rect?.left ?? 0) + (rect?.width ?? 0) / 2,
      y: (rect?.top ?? 0) + (rect?.height ?? 0) / 2,
    }
  })()`)

  await page.send('Input.dispatchKeyEvent', {
    code: 'Space',
    key: ' ',
    type: 'keyDown',
    windowsVirtualKeyCode: 32,
  })
  await delay(50)
  const duringTemporaryPan = await readPPTTemporaryPanState(page)
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mousePressed',
    x: panStart.x,
    y: panStart.y,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    type: 'mouseMoved',
    x: panStart.x + 86,
    y: panStart.y + 34,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mouseReleased',
    x: panStart.x + 86,
    y: panStart.y + 34,
  })
  await delay(80)
  const afterTemporaryPanDrag = await readPPTTemporaryPanState(page)
  await page.send('Input.dispatchKeyEvent', {
    code: 'Space',
    key: ' ',
    type: 'keyUp',
    windowsVirtualKeyCode: 32,
  })
  await delay(50)
  const afterTemporaryPanRelease = await readPPTTemporaryPanState(page)

  const beforeWheelViewport = await readPPTWheelViewportState(page)
  await page.send('Input.dispatchMouseEvent', {
    deltaX: 14,
    deltaY: 28,
    type: 'mouseWheel',
    x: panStart.x,
    y: panStart.y,
  })
  await delay(80)
  const afterWheelPan = await readPPTWheelViewportState(page)
  await page.send('Input.dispatchMouseEvent', {
    deltaX: 0,
    deltaY: 32,
    modifiers: 8,
    type: 'mouseWheel',
    x: panStart.x,
    y: panStart.y,
  })
  await delay(80)
  const afterShiftWheelPan = await readPPTWheelViewportState(page)
  await page.send('Input.dispatchMouseEvent', {
    deltaX: 0,
    deltaY: -120,
    modifiers: 2,
    type: 'mouseWheel',
    x: panStart.x,
    y: panStart.y,
  })
  await delay(80)
  const afterWheelZoom = await readPPTWheelViewportState(page)

  await page.eval(`document.querySelector('[data-ppt-pan-tool]')?.click()`)
  await delay(80)
  const afterToolbarPanTool = await readPPTPanToolState(page)
  await pressKey(page, {
    code: 'KeyV',
    key: 'v',
    windowsVirtualKeyCode: 86,
  })
  await delay(50)

  await pressKey(page, {
    code: 'KeyK',
    key: 'k',
    modifiers: 2,
    windowsVirtualKeyCode: 75,
  })
  await delay(80)
  const panToolPaletteIds = await readCommandPaletteIds(page, 'pan tool')
  await pressKey(page, {
    code: 'Enter',
    key: 'Enter',
    windowsVirtualKeyCode: 13,
  })
  await delay(80)
  const afterPalettePanTool = await readPPTPanToolState(page)
  await pressKey(page, {
    code: 'KeyV',
    key: 'v',
    windowsVirtualKeyCode: 86,
  })
  await delay(50)

  const beforePanToolDrag = await readPPTPanToolState(page)
  await pressKey(page, {
    code: 'KeyH',
    key: 'h',
    windowsVirtualKeyCode: 72,
  })
  await delay(80)
  const afterPanToolShortcut = await readPPTPanToolState(page)
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mousePressed',
    x: panStart.x,
    y: panStart.y,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    type: 'mouseMoved',
    x: panStart.x + 76,
    y: panStart.y + 29,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mouseReleased',
    x: panStart.x + 76,
    y: panStart.y + 29,
  })
  await delay(80)
  const afterPanToolDrag = await readPPTPanToolState(page)
  await pressKey(page, {
    code: 'KeyV',
    key: 'v',
    windowsVirtualKeyCode: 86,
  })
  await delay(50)
  const afterPanToolSelect = await readPPTPanToolState(page)

  await pressKey(page, {
    code: 'Digit0',
    key: '0',
    windowsVirtualKeyCode: 48,
  })
  await delay(120)
  const afterShortcutFitSlide = await readViewportState(page)

  record('runs PPT viewport fit keyboard shortcuts from canvas bindings', afterShortcutFitSelection.keyboardIntent === 'canvas-keyboard-viewport-shortcut-intent' && afterShortcutFitSelection.keyboardModel === 'canvas-keyboard-viewport-shortcuts' && afterShortcutFitSelection.scale > afterRestoreFitSlide.scale && afterShortcutFitSlide.scale < afterShortcutFitSelection.scale && afterShortcutFitSlide.label === afterRestoreFitSlide.label, {
    afterRestoreFitSlide,
    afterShortcutFitSelection,
    afterShortcutFitSlide,
  })

  record('pans PPT viewport with canvas Space temporary pan shortcut', beforeTemporaryPan.model === 'canvas-temporary-pan-shortcut' && beforeTemporaryPan.shortcut === 'Space' && beforeTemporaryPan.active === 'false' && duringTemporaryPan.active === 'true' && duringTemporaryPan.cursor === 'grab' && afterTemporaryPanDrag.gesture === 'false' && nearlyEqual(afterTemporaryPanDrag.scale, beforeTemporaryPan.scale, 0.001) && Math.abs(afterTemporaryPanDrag.x - beforeTemporaryPan.x) >= 40 && Math.abs(afterTemporaryPanDrag.y - beforeTemporaryPan.y) >= 20 && afterTemporaryPanDrag.selectedIds === beforeTemporaryPan.selectedIds && afterTemporaryPanRelease.active === 'false' && afterTemporaryPanRelease.gesture === 'false', {
    afterTemporaryPanDrag,
    afterTemporaryPanRelease,
    beforeTemporaryPan,
    duringTemporaryPan,
  })

  record('pans and zooms PPT viewport with canvas wheel affordance', beforeWheelViewport.model === 'canvas-wheel-viewport' && beforeWheelViewport.pan === 'ordinary-wheel' && beforeWheelViewport.horizontalPanModifier === 'Shift' && beforeWheelViewport.zoomModifier === 'Ctrl/Meta' && nearlyEqual(afterWheelPan.scale, beforeWheelViewport.scale, 0.001) && afterWheelPan.x < beforeWheelViewport.x && afterWheelPan.y < beforeWheelViewport.y && nearlyEqual(afterShiftWheelPan.scale, afterWheelPan.scale, 0.001) && afterShiftWheelPan.x < afterWheelPan.x && nearlyEqual(afterShiftWheelPan.y, afterWheelPan.y, 0.001) && afterWheelZoom.scale > afterShiftWheelPan.scale, {
    afterShiftWheelPan,
    afterWheelPan,
    afterWheelZoom,
    beforeWheelViewport,
  })

  record('activates and drags PPT viewport with canvas pan tool affordance', afterToolbarPanTool.active === 'true' && afterToolbarPanTool.toolbarPressed === 'true' && panToolPaletteIds.includes('tool:pan') && afterPalettePanTool.active === 'true' && !afterPalettePanTool.paletteOpen && beforePanToolDrag.active === 'false' && afterPanToolShortcut.model === 'canvas-pan-tool' && afterPanToolShortcut.shortcut === 'H' && afterPanToolShortcut.active === 'true' && afterPanToolShortcut.cursor === 'grab' && nearlyEqual(afterPanToolDrag.scale, beforePanToolDrag.scale, 0.001) && Math.abs(afterPanToolDrag.x - beforePanToolDrag.x) >= 40 && Math.abs(afterPanToolDrag.y - beforePanToolDrag.y) >= 20 && afterPanToolDrag.selectedIds === beforePanToolDrag.selectedIds && afterPanToolSelect.active === 'false', {
    afterPalettePanTool,
    afterPanToolDrag,
    afterPanToolSelect,
    afterPanToolShortcut,
    afterToolbarPanTool,
    beforePanToolDrag,
    panToolPaletteIds,
  })

  await page.eval(`document.querySelector('[data-ppt-laser-tool]')?.click()`)
  await delay(80)
  const afterToolbarLaserTool = await readPPTLaserToolState(page)
  await pressKey(page, {
    code: 'KeyV',
    key: 'v',
    windowsVirtualKeyCode: 86,
  })
  await delay(50)
  const afterLaserSelect = await readPPTLaserToolState(page)

  await pressKey(page, {
    code: 'KeyK',
    key: 'k',
    modifiers: 2,
    windowsVirtualKeyCode: 75,
  })
  await delay(80)
  const laserToolPaletteIds = await readCommandPaletteIds(page, 'laser pointer')
  await pressKey(page, {
    code: 'Enter',
    key: 'Enter',
    windowsVirtualKeyCode: 13,
  })
  await delay(80)
  const afterPaletteLaserTool = await readPPTLaserToolState(page)
  await pressKey(page, {
    code: 'KeyV',
    key: 'v',
    windowsVirtualKeyCode: 86,
  })
  await delay(50)

  const beforeLaserDrag = await readPPTLaserToolState(page)
  const laserStart = await getElementCenter(page, 's1-card-1')
  await pressKey(page, {
    code: 'KeyP',
    key: 'p',
    windowsVirtualKeyCode: 80,
  })
  await delay(80)
  const afterLaserShortcut = await readPPTLaserToolState(page)
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mousePressed',
    x: laserStart.x,
    y: laserStart.y,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    type: 'mouseMoved',
    x: laserStart.x + 44,
    y: laserStart.y + 22,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    type: 'mouseMoved',
    x: laserStart.x + 92,
    y: laserStart.y + 38,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mouseReleased',
    x: laserStart.x + 92,
    y: laserStart.y + 38,
  })
  await delay(80)
  const afterLaserDrag = await readPPTLaserToolState(page)
  await pressKey(page, {
    code: 'KeyV',
    key: 'v',
    windowsVirtualKeyCode: 86,
  })
  await delay(50)
  const afterLaserFinalSelect = await readPPTLaserToolState(page)

  record('activates and previews PPT laser pointer with canvas laser affordance', afterToolbarLaserTool.active === 'true' && afterToolbarLaserTool.toolbarPressed === 'true' && afterLaserSelect.active === 'false' && afterLaserSelect.trailPointCount === 0 && laserToolPaletteIds.includes('tool:laser') && afterPaletteLaserTool.active === 'true' && !afterPaletteLaserTool.paletteOpen && beforeLaserDrag.active === 'false' && afterLaserShortcut.model === 'canvas-laser-pointer-tool' && afterLaserShortcut.shortcut === 'P' && afterLaserShortcut.trailModel === 'canvas-laser-trail-overlay' && afterLaserShortcut.active === 'true' && afterLaserShortcut.cursor === 'crosshair' && afterLaserDrag.trailState === 'idle' && afterLaserDrag.trailPointCount >= 2 && afterLaserDrag.trailRenderedPointCount === afterLaserDrag.trailPointCount && afterLaserDrag.trailPath && afterLaserDrag.trailDot && afterLaserDrag.selectedIds === beforeLaserDrag.selectedIds && afterLaserDrag.elementCount === beforeLaserDrag.elementCount && afterLaserDrag.transform === beforeLaserDrag.transform && afterLaserFinalSelect.active === 'false' && afterLaserFinalSelect.trailPointCount === 0, {
    afterLaserDrag,
    afterLaserFinalSelect,
    afterLaserSelect,
    afterLaserShortcut,
    afterPaletteLaserTool,
    afterToolbarLaserTool,
    beforeLaserDrag,
    laserToolPaletteIds,
  })

  await pressKey(page, {
    code: 'Equal',
    key: '=',
    modifiers: 2,
    windowsVirtualKeyCode: 187,
  })
  await delay(80)
  const afterShortcutZoomIn = await readViewportState(page)

  await pressKey(page, {
    code: 'Minus',
    key: '-',
    modifiers: 2,
    windowsVirtualKeyCode: 189,
  })
  await delay(80)
  const afterShortcutZoomOut = await readViewportState(page)

  await pressKey(page, {
    code: 'Digit0',
    key: '0',
    modifiers: 2,
    windowsVirtualKeyCode: 48,
  })
  await delay(80)
  const afterShortcutResetZoom = await readViewportState(page)

  record('runs PPT viewport zoom keyboard shortcuts from canvas bindings', afterShortcutZoomIn.keyboardIntent === 'canvas-keyboard-viewport-shortcut-intent' && afterShortcutZoomIn.keyboardModel === 'canvas-keyboard-viewport-shortcuts' && afterShortcutZoomIn.scale > afterShortcutFitSlide.scale && afterShortcutZoomOut.scale < afterShortcutZoomIn.scale && nearlyEqual(afterShortcutResetZoom.scale, 1, 0.001) && nearlyEqual(afterShortcutResetZoom.x, 0, 0.001) && nearlyEqual(afterShortcutResetZoom.y, 0, 0.001), {
    afterShortcutFitSlide,
    afterShortcutResetZoom,
    afterShortcutZoomIn,
    afterShortcutZoomOut,
  })
}

async function runMinimapScenario(page) {
  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)
  await page.eval(`document.querySelector('[data-ppt-view-fit-slide]')?.click()`)
  await delay(100)

  const initial = await readPPTMinimapState(page)

  record('renders PPT minimap viewport overview', initial.open && initial.model === 'canvas-minimap-read-model' && initial.itemCount >= 5 && initial.hasWorld && initial.hasViewport && initial.togglePressed === 'true', initial)

  const point = await getElementCenter(page, 's1-card-1')
  await clickMouse(page, point.x, point.y, 1)
  await delay(80)
  await page.eval(`document.querySelector('[data-ppt-view-fit-selection]')?.click()`)
  await delay(120)

  const afterFitSelection = await readPPTMinimapState(page)

  record('syncs PPT minimap viewport rect with zoomed viewport', afterFitSelection.open && afterFitSelection.scale > initial.scale && afterFitSelection.viewportW < initial.viewportW && afterFitSelection.viewportH < initial.viewportH, {
    afterFitSelection,
    initial,
  })

  const clickPoint = await page.eval(`(() => {
    const world = document.querySelector('[data-ppt-minimap-world]')
    const map = document.querySelector('[data-ppt-minimap-map]')
    const worldRect = world.getBoundingClientRect()
    const mapRect = map.getBoundingClientRect()

    return {
      x: Math.min(mapRect.right - 6, worldRect.left + worldRect.width * 0.86),
      y: Math.min(mapRect.bottom - 6, worldRect.top + worldRect.height * 0.52),
    }
  })()`)

  await clickMouse(page, clickPoint.x, clickPoint.y, 1)
  await delay(120)

  const afterMinimapClick = await readPPTMinimapState(page)

  record('moves PPT viewport from minimap click without changing zoom', afterMinimapClick.open && Math.abs(afterMinimapClick.scale - afterFitSelection.scale) < 0.001 && (Math.abs(afterMinimapClick.viewportX - afterFitSelection.viewportX) > 1 || Math.abs(afterMinimapClick.viewportY - afterFitSelection.viewportY) > 1), {
    afterFitSelection,
    afterMinimapClick,
  })

  await page.eval(`document.querySelector('[data-ppt-view-minimap]')?.click()`)
  await delay(80)

  const afterToolbarHide = await readPPTMinimapState(page)

  record('hides PPT minimap from toolbar toggle', !afterToolbarHide.open && afterToolbarHide.stageAttr === 'false' && afterToolbarHide.togglePressed === 'false', afterToolbarHide)

  await pressKey(page, {
    code: 'KeyK',
    key: 'k',
    modifiers: 2,
    windowsVirtualKeyCode: 75,
  })
  await delay(80)
  await page.send('Input.insertText', { text: 'minimap' })
  await delay(80)

  const beforePaletteShow = await page.eval(`(() => ({
    itemPresent: !!document.querySelector('[data-ppt-command-palette-item="view:toggle-minimap"]'),
    open: !!document.querySelector('[data-ppt-command-palette]'),
  }))()`)

  await pressKey(page, {
    code: 'Enter',
    key: 'Enter',
    windowsVirtualKeyCode: 13,
  })
  await delay(100)

  const afterPaletteShow = await readPPTMinimapState(page)

  record('shows PPT minimap from command palette toggle', beforePaletteShow.open && beforePaletteShow.itemPresent && afterPaletteShow.open && afterPaletteShow.stageAttr === 'true' && afterPaletteShow.togglePressed === 'true' && !afterPaletteShow.paletteOpen, {
    afterPaletteShow,
    beforePaletteShow,
  })

  await page.eval(`document.querySelector('[data-ppt-view-fit-slide]')?.click()`)
  await delay(80)
}

async function runTidySelectionScenario(page) {
  const ids = ['s1-card-1', 's1-card-2', 's1-side-panel']

  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)
  await page.eval(`document.querySelector('[data-ppt-view-fit-slide]')?.click()`)
  await delay(80)

  await selectPPTLayerRows(page, ids)
  await delay(80)

  const beforeTidy = await readPPTTidyState(page, ids)

  record('enables PPT tidy selection for three visible unlocked objects', beforeTidy.selectedCount === 3 && !beforeTidy.toolbarDisabled && !beforeTidy.floatingDisabled, beforeTidy)

  await page.eval(`document.querySelector('[data-ppt-command="tidy-selection"]')?.click()`)
  await delay(120)

  const afterTidy = await readPPTTidyState(page, ids)

  record('tidies selected PPT objects into canvas-style grid', afterTidy.elementCount === beforeTidy.elementCount && afterTidy.selectedCount === 3 && positionsChanged(beforeTidy.positions, afterTidy.positions, ids) && positionsMatch(afterTidy.positions, beforeTidy.expected, ids), {
    afterTidy,
    beforeTidy,
  })

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(100)

  const afterUndo = await readPPTTidyState(page, ids)

  record('undoes PPT tidy selection in one step', afterUndo.selectedCount === 3 && positionsMatch(afterUndo.positions, beforeTidy.positions, ids), {
    afterUndo,
    beforeTidy,
  })

  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)

  await selectPPTLayerRows(page, [ids[0]])
  await delay(60)
  await pressKey(page, {
    code: 'KeyK',
    key: 'k',
    modifiers: 2,
    windowsVirtualKeyCode: 75,
  })
  await delay(80)
  await readCommandPaletteIds(page, 'tidy')

  const paletteDisabled = await page.eval(`(() => ({
    disabled: document.querySelector('[data-ppt-command-palette-item="command:tidy-selection"]')?.disabled ?? false,
    itemPresent: !!document.querySelector('[data-ppt-command-palette-item="command:tidy-selection"]'),
    open: !!document.querySelector('[data-ppt-command-palette]'),
    selectedCount: document.querySelectorAll('[data-selected="true"]').length,
  }))()`)

  record('disables PPT tidy selection palette item below minimum selection', paletteDisabled.itemPresent && paletteDisabled.open && paletteDisabled.selectedCount === 1 && paletteDisabled.disabled, paletteDisabled)

  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)
}

async function readPPTCommandPaletteComboboxState(page) {
  return page.eval(`(() => {
    const input = document.querySelector('[data-ppt-command-palette-query]')
    const activeDescendant = input?.getAttribute('aria-activedescendant') ?? ''
    const activeOption = activeDescendant ? document.getElementById(activeDescendant) : null

    return {
      activeDescendant,
      activeOptionExists: !!activeOption,
      activeOptionId: document.querySelector('[data-ppt-command-palette-active="true"]')?.id ?? '',
      activeOptionItem: activeOption?.getAttribute('data-ppt-command-palette-item') ?? '',
      activeOptionRole: activeOption?.getAttribute('role') ?? '',
      activeOptionSelected: activeOption?.getAttribute('aria-selected') ?? '',
      focusedQuery: document.activeElement === input,
    }
  })()`)
}

async function readCommandPaletteIds(page, query) {
  await page.eval(`(() => {
    const input = document.querySelector('[data-ppt-command-palette-query]')
    const valueSetter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set

    valueSetter.call(input, ${JSON.stringify(query)})
    input.dispatchEvent(new Event('input', { bubbles: true }))
    input.dispatchEvent(new Event('change', { bubbles: true }))
  })()`)
  await delay(50)

  return page.eval(`(() => [...document.querySelectorAll('[data-ppt-command-palette-item]')]
    .map((item) => item.getAttribute('data-ppt-command-palette-item'))
  )()`)
}

async function readPPTElementLayerState(page, elementId) {
  return page.eval(`(() => {
    const elementId = ${JSON.stringify(elementId)}
    const shell = document.querySelector('.ppt-stage-shell')
    const order = [...document.querySelectorAll('[data-ppt-element]')]
      .map((element) => element.getAttribute('data-ppt-element'))
    const element = document.querySelector('[data-ppt-element="' + elementId + '"]')

    return {
      index: order.indexOf(elementId),
      keyboardCommandDispatch: shell?.getAttribute('data-ppt-keyboard-command-dispatch') ?? '',
      keyboardCommandIntent: shell?.getAttribute('data-ppt-keyboard-command-intent') ?? '',
      locked: element?.getAttribute('data-locked') ?? '',
      order,
      selected: element?.getAttribute('data-selected') === 'true',
    }
  })()`)
}

async function readPPTObjectStateImportState(page, elementId) {
  return page.eval(`(() => {
    const elementId = ${JSON.stringify(elementId)}
    const shell = document.querySelector('.ppt-stage-shell')
    const row = document.querySelector('[data-ppt-layer-row="' + elementId + '"]')
    const stageElement = document.querySelector('[data-ppt-element="' + elementId + '"]')

    return {
      commandTypes: shell?.getAttribute('data-ppt-object-state-import-command-types') ?? '',
      commands: shell?.getAttribute('data-ppt-object-state-import-commands') ?? '',
      fields: shell?.getAttribute('data-ppt-object-state-import-fields') ?? '',
      format: shell?.getAttribute('data-ppt-object-state-import-format') ?? '',
      hidden: row?.getAttribute('data-ppt-layer-pane-hidden') ?? '',
      importLocked: shell?.getAttribute('data-ppt-object-state-import-locked') ?? '',
      importVisible: shell?.getAttribute('data-ppt-object-state-import-visible') ?? '',
      jsonLength: Number(shell?.getAttribute('data-ppt-object-state-import-json-length') ?? 0),
      locked: row?.getAttribute('data-ppt-layer-pane-locked') ?? '',
      lockTargets: shell?.getAttribute('data-ppt-object-state-import-lock-targets') ?? '',
      model: shell?.getAttribute('data-ppt-object-state-import-model') ?? '',
      objectIds: shell?.getAttribute('data-ppt-object-state-import-objects') ?? '',
      rowSelected: row?.getAttribute('aria-selected') ?? '',
      slide: shell?.getAttribute('data-ppt-object-state-import-slide') ?? '',
      stageElementExists: !!stageElement,
      stageSelected: stageElement?.getAttribute('data-selected') ?? '',
      visibilityCommand: shell?.getAttribute('data-ppt-object-visibility-command') ?? '',
      visibilityTargets: shell?.getAttribute('data-ppt-object-state-import-visibility-targets') ?? '',
    }
  })()`)
}

async function readPPTObjectLayerImportState(page, elementId) {
  return page.eval(`(() => {
    const elementId = ${JSON.stringify(elementId)}
    const shell = document.querySelector('.ppt-stage-shell')
    const layerRows = [...document.querySelectorAll('[data-ppt-layer-row][data-ppt-layer-pane-row-type="object"]')]
    const stageElements = [...document.querySelectorAll('[data-ppt-element]')]
    const row = document.querySelector('[data-ppt-layer-row="' + elementId + '"]')
    const stageElement = document.querySelector('[data-ppt-element="' + elementId + '"]')
    const layerOrder = layerRows.map((candidate) => candidate.getAttribute('data-ppt-layer-row') ?? '')
    const stageOrder = stageElements.map((candidate) => candidate.getAttribute('data-ppt-element') ?? '')

    return {
      command: shell?.getAttribute('data-ppt-object-layer-import-command') ?? '',
      commandType: shell?.getAttribute('data-ppt-object-layer-import-command-type') ?? '',
      fields: shell?.getAttribute('data-ppt-object-layer-import-fields') ?? '',
      format: shell?.getAttribute('data-ppt-object-layer-import-format') ?? '',
      fromIndex: Number(shell?.getAttribute('data-ppt-object-layer-import-from-index') ?? -1),
      jsonLength: Number(shell?.getAttribute('data-ppt-object-layer-import-json-length') ?? 0),
      layerIndex: layerOrder.indexOf(elementId),
      layerOrder,
      model: shell?.getAttribute('data-ppt-object-layer-import-model') ?? '',
      objectId: shell?.getAttribute('data-ppt-object-layer-import-object') ?? '',
      position: shell?.getAttribute('data-ppt-object-layer-import-position') ?? '',
      rowOrder: Number(row?.getAttribute('data-ppt-layer-pane-order') ?? -1),
      rowSelected: row?.getAttribute('aria-selected') ?? '',
      slide: shell?.getAttribute('data-ppt-object-layer-import-slide') ?? '',
      stageIndex: stageOrder.indexOf(elementId),
      stageOrder,
      stageSelected: stageElement?.getAttribute('data-selected') ?? '',
      toIndex: Number(shell?.getAttribute('data-ppt-object-layer-import-to-index') ?? -1),
    }
  })()`)
}

async function readPPTTidyState(page, ids) {
  return page.eval(`(() => {
    const ids = ${JSON.stringify(ids)}
    const entries = ids.map((id) => {
      const element = document.querySelector(\`[data-ppt-element="\${id}"]\`)

      return {
        h: parseFloat(element?.style.height ?? '0'),
        id,
        w: parseFloat(element?.style.width ?? '0'),
        x: parseFloat(element?.style.left ?? '0'),
        y: parseFloat(element?.style.top ?? '0'),
      }
    })
    const minX = Math.min(...entries.map((entry) => entry.x))
    const minY = Math.min(...entries.map((entry) => entry.y))
    const columnCount = Math.ceil(Math.sqrt(entries.length))
    const cellWidth = Math.max(...entries.map((entry) => entry.w)) + ${PPT_TIDY_GAP}
    const cellHeight = Math.max(...entries.map((entry) => entry.h)) + ${PPT_TIDY_GAP}
    const expectedEntries = [...entries]
      .sort((a, b) => a.y === b.y ? a.x - b.x : a.y - b.y)
      .map((entry, index) => {
        const column = index % columnCount
        const row = Math.floor(index / columnCount)
        const targetX = minX + column * cellWidth
        const targetY = minY + row * cellHeight

        return [entry.id, {
          ...entry,
          x: Math.min(${PPT_SLIDE_WIDTH} - entry.w, Math.max(0, targetX)),
          y: Math.min(${PPT_SLIDE_HEIGHT} - entry.h, Math.max(0, targetY)),
        }]
      })

    return {
      elementCount: document.querySelectorAll('[data-ppt-element]').length,
      expected: Object.fromEntries(expectedEntries),
      floatingDisabled: document.querySelector('[data-ppt-floating-command="tidy-selection"]')?.disabled ?? true,
      positions: Object.fromEntries(entries.map((entry) => [entry.id, entry])),
      selectedCount: document.querySelectorAll('[data-selected="true"]').length,
      selectedIds: [...document.querySelectorAll('[data-selected="true"]')]
        .map((element) => element.getAttribute('data-ppt-element')),
      toolbarDisabled: document.querySelector('[data-ppt-command="tidy-selection"]')?.disabled ?? true,
    }
  })()`)
}

async function readPPTFlipState(page, ids) {
  return page.eval(`(() => {
    const ids = ${JSON.stringify(ids)}
    const entries = ids.map((id) => {
      const element = document.querySelector(\`[data-ppt-element="\${id}"]\`)

      return {
        flipH: element?.getAttribute('data-ppt-flip-h') ?? '',
        flipV: element?.getAttribute('data-ppt-flip-v') ?? '',
        h: parseFloat(element?.style.height ?? '0'),
        id,
        transform: element?.style.transform ?? '',
        w: parseFloat(element?.style.width ?? '0'),
        x: parseFloat(element?.style.left ?? '0'),
        y: parseFloat(element?.style.top ?? '0'),
      }
    })
    const minX = Math.min(...entries.map((entry) => entry.x))
    const minY = Math.min(...entries.map((entry) => entry.y))
    const maxX = Math.max(...entries.map((entry) => entry.x + entry.w))
    const maxY = Math.max(...entries.map((entry) => entry.y + entry.h))
    const pivotX = minX + (maxX - minX) / 2
    const pivotY = minY + (maxY - minY) / 2
    const expectedHorizontal = Object.fromEntries(entries.map((entry) => {
      const targetX = 2 * pivotX - (entry.x + entry.w)

      return [entry.id, {
        ...entry,
        x: Math.min(${PPT_SLIDE_WIDTH} - entry.w, Math.max(0, targetX)),
      }]
    }))
    const expectedVertical = Object.fromEntries(entries.map((entry) => {
      const targetY = 2 * pivotY - (entry.y + entry.h)

      return [entry.id, {
        ...entry,
        y: Math.min(${PPT_SLIDE_HEIGHT} - entry.h, Math.max(0, targetY)),
      }]
    }))

    return {
      expectedHorizontal,
      expectedVertical,
      flipHorizontalDisabled: document.querySelector('[data-ppt-command="flip-horizontal"]')?.disabled ?? true,
      flipVerticalDisabled: document.querySelector('[data-ppt-command="flip-vertical"]')?.disabled ?? true,
      positions: Object.fromEntries(entries.map((entry) => [entry.id, entry])),
      selectedCount: document.querySelectorAll('[data-selected="true"]').length,
      selectedIds: [...document.querySelectorAll('[data-selected="true"]')]
        .map((element) => element.getAttribute('data-ppt-element')),
    }
  })()`)
}

async function selectPPTLayerRows(page, ids) {
  await page.eval(`(() => {
    const ids = ${JSON.stringify(ids)}

    ids.forEach((id, index) => {
      document.querySelector(\`[data-ppt-layer-select="\${id}"]\`)?.dispatchEvent(new MouseEvent('click', {
        bubbles: true,
        cancelable: true,
        ctrlKey: index > 0,
      }))
    })
  })()`)
}

function getPPTColorSwatchState(page, channel, elementId = '') {
  return page.eval(`((channel, elementId) => {
    const selected = elementId
      ? document.querySelector(\`[data-ppt-element="\${elementId}"]\`)
      : document.querySelector('[data-selected="true"]')
    const selectedId = selected?.getAttribute('data-ppt-element') ?? ''
    const strokeElement = selected?.querySelector('line, [data-ppt-line-path], [data-ppt-freeform-path]')
    const thumb = selectedId
      ? document.querySelector(\`.ppt-thumb[aria-current="page"] [data-ppt-thumb-element="\${selectedId}"]\`)
      : null
    const exportCode = document.querySelector('.ppt-export-code')?.value ?? ''
    const exportIndex = selectedId ? exportCode.indexOf(\`"id": "\${selectedId}"\`) : -1
    const recentColors = (document.querySelector('.ppt-stage-shell')?.getAttribute('data-ppt-recent-colors') ?? '')
      .split(/\\s+/)
      .filter(Boolean)
    const palette = document.querySelector(\`[data-ppt-color-swatch-palette="\${channel}"]\`)
    const stage = document.querySelector('.ppt-stage-shell')
    const selectedTheme = document.querySelector(\`[data-ppt-color-swatch="\${channel}"][data-ppt-color-source="theme"][aria-pressed="true"]\`)
    const inputSelector = channel === 'shape-fill'
      ? '[data-ppt-style-field="fill"]'
      : channel === 'shape-stroke'
        ? '[data-ppt-style-field="stroke-color"]'
        : channel === 'text-color'
          ? '[data-ppt-style-field="text-color"]'
          : '[data-ppt-style-field="line-stroke-color"]'

    return {
      background: selected?.style.background ?? '',
      borderColor: selected?.style.borderColor ?? '',
      command: stage?.getAttribute('data-ppt-color-swatch-command') ?? '',
      commandChannel: stage?.getAttribute('data-ppt-color-swatch-command-channel') ?? '',
      commandObjects: stage?.getAttribute('data-ppt-color-swatch-command-objects') ?? '',
      commandSlide: stage?.getAttribute('data-ppt-color-swatch-command-slide') ?? '',
      commandSource: stage?.getAttribute('data-ppt-color-swatch-command-source') ?? '',
      commandSwatch: stage?.getAttribute('data-ppt-color-swatch-command-swatch') ?? '',
      commandToken: stage?.getAttribute('data-ppt-color-swatch-command-token') ?? '',
      commandType: stage?.getAttribute('data-ppt-color-swatch-command-type') ?? '',
      commandValue: stage?.getAttribute('data-ppt-color-swatch-command-value') ?? '',
      importCommandChannels: stage?.getAttribute('data-ppt-color-swatch-import-command-channels') ?? '',
      importCommandSources: stage?.getAttribute('data-ppt-color-swatch-import-command-sources') ?? '',
      importCommandSwatches: stage?.getAttribute('data-ppt-color-swatch-import-command-swatches') ?? '',
      importCommandTargets: stage?.getAttribute('data-ppt-color-swatch-import-command-targets') ?? '',
      importCommandTokens: stage?.getAttribute('data-ppt-color-swatch-import-command-tokens') ?? '',
      importCommandTypes: stage?.getAttribute('data-ppt-color-swatch-import-command-types') ?? '',
      importCommandValues: stage?.getAttribute('data-ppt-color-swatch-import-command-values') ?? '',
      importCommands: stage?.getAttribute('data-ppt-color-swatch-import-commands') ?? '',
      importFields: stage?.getAttribute('data-ppt-color-swatch-import-fields') ?? '',
      importFormat: stage?.getAttribute('data-ppt-color-swatch-import-format') ?? '',
      importJsonLength: Number(stage?.getAttribute('data-ppt-color-swatch-import-json-length') ?? 0),
      importModel: stage?.getAttribute('data-ppt-color-swatch-import-model') ?? '',
      importObjects: stage?.getAttribute('data-ppt-color-swatch-import-objects') ?? '',
      importSource: stage?.getAttribute('data-ppt-color-swatch-import-source') ?? '',
      importSwatch: stage?.getAttribute('data-ppt-color-swatch-import-swatch') ?? '',
      importToken: stage?.getAttribute('data-ppt-color-swatch-import-token') ?? '',
      importValue: stage?.getAttribute('data-ppt-color-swatch-import-value') ?? '',
      descriptorCommand: palette?.getAttribute('data-ppt-color-swatch-command') ?? '',
      descriptorControl: palette?.getAttribute('data-ppt-color-swatch-control') ?? '',
      descriptorDisabled: palette?.getAttribute('data-ppt-color-swatch-disabled') ?? '',
      descriptorMixed: palette?.getAttribute('data-ppt-color-swatch-mixed') ?? '',
      descriptorModel: palette?.getAttribute('data-ppt-color-swatch-model') ?? '',
      descriptorObjectIds: palette?.getAttribute('data-ppt-color-swatch-object-ids') ?? '',
      descriptorSelectedId: palette?.getAttribute('data-ppt-color-swatch-selected-id') ?? '',
      exportSlice: exportIndex >= 0
        ? exportCode.slice(Math.max(0, exportIndex - 700), exportIndex + 1200)
        : '',
      fillValue: document.querySelector('[data-ppt-style-field="fill"]')?.value ?? '',
      inputValue: document.querySelector(inputSelector)?.value ?? '',
      model: stage?.getAttribute('data-ppt-color-swatch-model') ?? '',
      packageChannel: palette?.getAttribute('data-ppt-color-swatch-package-channel') ?? '',
      recentAccentCount: recentColors.filter((color) => color === '#2563eb').length,
      recentColors,
      recentCount: document.querySelectorAll(\`[data-ppt-color-swatch="\${channel}"][data-ppt-color-source="recent"]\`).length,
      recentUnique: new Set(recentColors).size === recentColors.length,
      selectedId,
      selectedKind: selected?.getAttribute('data-kind') ?? '',
      selectedToken: selectedTheme?.getAttribute('data-ppt-color-token') ?? '',
      stroke: strokeElement?.getAttribute('stroke') ?? '',
      strokeValue: document.querySelector('[data-ppt-style-field="stroke-color"]')?.value ?? '',
      styleColor: selected?.style.color ?? '',
      themeCount: document.querySelectorAll(\`[data-ppt-color-swatch="\${channel}"][data-ppt-color-source="theme"]\`).length,
      thumbBackground: thumb?.style.background ?? '',
      thumbBorder: thumb?.style.border ?? '',
    }
  })(${JSON.stringify(channel)}, ${JSON.stringify(elementId)})`)
}

function getPPTTextFormatPainterState(page, elementId) {
  return page.eval(`((id) => {
    const element = document.querySelector(\`[data-ppt-element="\${id}"]\`)
    const layerName = document.querySelector(\`[data-ppt-layer-row="\${id}"] .ppt-layer-name\`)
    const paragraph = element?.querySelector('.ppt-text-paragraph')
    const stage = document.querySelector('.ppt-stage-shell')

    return {
      boldRun: element?.querySelector('[data-ppt-run-bold="true"]')?.style.fontWeight ?? '',
      boldRunCount: element?.querySelectorAll('[data-ppt-run-bold="true"]').length ?? 0,
      bulletList: element?.getAttribute('data-ppt-bullet-list') ?? '',
      color: element?.style.color ?? '',
      fontFamily: element?.getAttribute('data-ppt-font-family') ?? '',
      fontSize: element?.style.fontSize ?? '',
      fontWeight: element?.style.fontWeight ?? '',
      height: element?.style.height ?? '',
      italicRun: element?.querySelector('[data-ppt-run-italic="true"]')?.style.fontStyle ?? '',
      italicRunCount: element?.querySelectorAll('[data-ppt-run-italic="true"]').length ?? 0,
      left: element?.style.left ?? '',
      name: layerName?.textContent ?? '',
      numberedList: element?.getAttribute('data-ppt-numbered-list') ?? '',
      paragraphBullet: paragraph?.getAttribute('data-ppt-bullet') === 'true'
        ? paragraph.textContent ?? ''
        : '',
      paragraphLineHeight: paragraph?.getAttribute('data-ppt-line-height') ?? '',
      paragraphList: paragraph?.getAttribute('data-ppt-list') ?? '',
      paragraphSpacingAfter: paragraph?.getAttribute('data-ppt-spacing-after') ?? '',
      paragraphSpacingBefore: paragraph?.getAttribute('data-ppt-spacing-before') ?? '',
      selected: element?.getAttribute('data-selected') ?? '',
      styleAlignItems: element?.style.alignItems ?? '',
      styleClipboardCategories: stage?.getAttribute('data-ppt-style-clipboard-categories') ?? '',
      styleClipboardCommand: stage?.getAttribute('data-ppt-style-clipboard-command') ?? '',
      styleClipboardCommandApplications: stage?.getAttribute('data-ppt-style-clipboard-command-applications') ?? '',
      styleClipboardCommandTargets: stage?.getAttribute('data-ppt-style-clipboard-command-targets') ?? '',
      styleClipboardCommandType: stage?.getAttribute('data-ppt-style-clipboard-command-type') ?? '',
      styleFontFamily: element?.style.fontFamily ?? '',
      styleClipboardPackageCategories: stage?.getAttribute('data-ppt-style-clipboard-package-categories') ?? '',
      styleClipboardRunItalic: stage?.getAttribute('data-ppt-style-clipboard-run-italic') ?? '',
      styleClipboardRunUnderline: stage?.getAttribute('data-ppt-style-clipboard-run-underline') ?? '',
      stylePadding: element?.style.padding ?? '',
      text: element?.textContent ?? '',
      textAlign: element?.style.textAlign ?? '',
      textInset: element?.getAttribute('data-ppt-text-inset') ?? '',
      textFontSizeImportCategories: stage?.getAttribute('data-ppt-text-font-size-import-categories') ?? '',
      textFontSizeImportCommand: stage?.getAttribute('data-ppt-text-font-size-import-command') ?? '',
      textFontSizeImportCommandTargets: stage?.getAttribute('data-ppt-text-font-size-import-command-targets') ?? '',
      textFontSizeImportCommandType: stage?.getAttribute('data-ppt-text-font-size-import-command-type') ?? '',
      textFontSizeImportFields: stage?.getAttribute('data-ppt-text-font-size-import-fields') ?? '',
      textFontSizeImportFormat: stage?.getAttribute('data-ppt-text-font-size-import-format') ?? '',
      textFontSizeImportJsonLength: Number(stage?.getAttribute('data-ppt-text-font-size-import-json-length') ?? 0),
      textFontSizeImportModel: stage?.getAttribute('data-ppt-text-font-size-import-model') ?? '',
      textFontSizeImportObjects: stage?.getAttribute('data-ppt-text-font-size-import-objects') ?? '',
      textFontSizeImportValue: stage?.getAttribute('data-ppt-text-font-size-import-value') ?? '',
      textFontWeightImportCategories: stage?.getAttribute('data-ppt-text-font-weight-import-categories') ?? '',
      textFontWeightImportCommand: stage?.getAttribute('data-ppt-text-font-weight-import-command') ?? '',
      textFontWeightImportCommandTargets: stage?.getAttribute('data-ppt-text-font-weight-import-command-targets') ?? '',
      textFontWeightImportCommandType: stage?.getAttribute('data-ppt-text-font-weight-import-command-type') ?? '',
      textFontWeightImportFields: stage?.getAttribute('data-ppt-text-font-weight-import-fields') ?? '',
      textFontWeightImportFormat: stage?.getAttribute('data-ppt-text-font-weight-import-format') ?? '',
      textFontWeightImportJsonLength: Number(stage?.getAttribute('data-ppt-text-font-weight-import-json-length') ?? 0),
      textFontWeightImportModel: stage?.getAttribute('data-ppt-text-font-weight-import-model') ?? '',
      textFontWeightImportObjects: stage?.getAttribute('data-ppt-text-font-weight-import-objects') ?? '',
      textFontWeightImportValue: stage?.getAttribute('data-ppt-text-font-weight-import-value') ?? '',
      runSize: element?.querySelector('[data-ppt-run-size]')?.style.fontSize ?? '',
      runSizeCount: element?.querySelectorAll('[data-ppt-run-size]').length ?? 0,
      runSizeValue: element?.querySelector('[data-ppt-run-size]')?.getAttribute('data-ppt-run-size') ?? '',
      textRunSizeImportCommandFields: stage?.getAttribute('data-ppt-text-run-size-import-command-fields') ?? '',
      textRunSizeImportCommandIds: stage?.getAttribute('data-ppt-text-run-size-import-command-ids') ?? '',
      textRunSizeImportCommandTargets: stage?.getAttribute('data-ppt-text-run-size-import-command-targets') ?? '',
      textRunSizeImportCommandTypes: stage?.getAttribute('data-ppt-text-run-size-import-command-types') ?? '',
      textRunSizeImportCommandValues: stage?.getAttribute('data-ppt-text-run-size-import-command-values') ?? '',
      textRunSizeImportFields: stage?.getAttribute('data-ppt-text-run-size-import-fields') ?? '',
      textRunSizeImportFormat: stage?.getAttribute('data-ppt-text-run-size-import-format') ?? '',
      textRunSizeImportJsonLength: Number(stage?.getAttribute('data-ppt-text-run-size-import-json-length') ?? 0),
      textRunSizeImportModel: stage?.getAttribute('data-ppt-text-run-size-import-model') ?? '',
      textRunSizeImportObjects: stage?.getAttribute('data-ppt-text-run-size-import-objects') ?? '',
      textRunSizeImportRuns: Number(stage?.getAttribute('data-ppt-text-run-size-import-runs') ?? 0),
      textRunSizeImportSlide: stage?.getAttribute('data-ppt-text-run-size-import-slide') ?? '',
      textRunSizeImportValue: stage?.getAttribute('data-ppt-text-run-size-import-value') ?? '',
      runColor: element?.querySelector('[data-ppt-run-color]')?.style.color ?? '',
      runColorCount: element?.querySelectorAll('[data-ppt-run-color]').length ?? 0,
      runColorValue: element?.querySelector('[data-ppt-run-color]')?.getAttribute('data-ppt-run-color') ?? '',
      textRunColorImportCommandFields: stage?.getAttribute('data-ppt-text-run-color-import-command-fields') ?? '',
      textRunColorImportCommandIds: stage?.getAttribute('data-ppt-text-run-color-import-command-ids') ?? '',
      textRunColorImportCommandTargets: stage?.getAttribute('data-ppt-text-run-color-import-command-targets') ?? '',
      textRunColorImportCommandTypes: stage?.getAttribute('data-ppt-text-run-color-import-command-types') ?? '',
      textRunColorImportCommandValues: stage?.getAttribute('data-ppt-text-run-color-import-command-values') ?? '',
      textRunColorImportFields: stage?.getAttribute('data-ppt-text-run-color-import-fields') ?? '',
      textRunColorImportFormat: stage?.getAttribute('data-ppt-text-run-color-import-format') ?? '',
      textRunColorImportJsonLength: Number(stage?.getAttribute('data-ppt-text-run-color-import-json-length') ?? 0),
      textRunColorImportModel: stage?.getAttribute('data-ppt-text-run-color-import-model') ?? '',
      textRunColorImportObjects: stage?.getAttribute('data-ppt-text-run-color-import-objects') ?? '',
      textRunColorImportRuns: Number(stage?.getAttribute('data-ppt-text-run-color-import-runs') ?? 0),
      textRunColorImportSlide: stage?.getAttribute('data-ppt-text-run-color-import-slide') ?? '',
      textRunColorImportValue: stage?.getAttribute('data-ppt-text-run-color-import-value') ?? '',
      textRunBoldImportCommandFields: stage?.getAttribute('data-ppt-text-run-bold-import-command-fields') ?? '',
      textRunBoldImportCommandIds: stage?.getAttribute('data-ppt-text-run-bold-import-command-ids') ?? '',
      textRunBoldImportCommandTargets: stage?.getAttribute('data-ppt-text-run-bold-import-command-targets') ?? '',
      textRunBoldImportCommandTypes: stage?.getAttribute('data-ppt-text-run-bold-import-command-types') ?? '',
      textRunBoldImportCommandValues: stage?.getAttribute('data-ppt-text-run-bold-import-command-values') ?? '',
      textRunBoldImportFields: stage?.getAttribute('data-ppt-text-run-bold-import-fields') ?? '',
      textRunBoldImportFormat: stage?.getAttribute('data-ppt-text-run-bold-import-format') ?? '',
      textRunBoldImportJsonLength: Number(stage?.getAttribute('data-ppt-text-run-bold-import-json-length') ?? 0),
      textRunBoldImportModel: stage?.getAttribute('data-ppt-text-run-bold-import-model') ?? '',
      textRunBoldImportObjects: stage?.getAttribute('data-ppt-text-run-bold-import-objects') ?? '',
      textRunBoldImportRuns: Number(stage?.getAttribute('data-ppt-text-run-bold-import-runs') ?? 0),
      textRunBoldImportSlide: stage?.getAttribute('data-ppt-text-run-bold-import-slide') ?? '',
      textRunBoldImportValue: stage?.getAttribute('data-ppt-text-run-bold-import-value') ?? '',
      textRunItalicImportCommandFields: stage?.getAttribute('data-ppt-text-run-italic-import-command-fields') ?? '',
      textRunItalicImportCommandIds: stage?.getAttribute('data-ppt-text-run-italic-import-command-ids') ?? '',
      textRunItalicImportCommandTargets: stage?.getAttribute('data-ppt-text-run-italic-import-command-targets') ?? '',
      textRunItalicImportCommandTypes: stage?.getAttribute('data-ppt-text-run-italic-import-command-types') ?? '',
      textRunItalicImportCommandValues: stage?.getAttribute('data-ppt-text-run-italic-import-command-values') ?? '',
      textRunItalicImportFields: stage?.getAttribute('data-ppt-text-run-italic-import-fields') ?? '',
      textRunItalicImportFormat: stage?.getAttribute('data-ppt-text-run-italic-import-format') ?? '',
      textRunItalicImportJsonLength: Number(stage?.getAttribute('data-ppt-text-run-italic-import-json-length') ?? 0),
      textRunItalicImportModel: stage?.getAttribute('data-ppt-text-run-italic-import-model') ?? '',
      textRunItalicImportObjects: stage?.getAttribute('data-ppt-text-run-italic-import-objects') ?? '',
      textRunItalicImportRuns: Number(stage?.getAttribute('data-ppt-text-run-italic-import-runs') ?? 0),
      textRunItalicImportSlide: stage?.getAttribute('data-ppt-text-run-italic-import-slide') ?? '',
      textRunItalicImportValue: stage?.getAttribute('data-ppt-text-run-italic-import-value') ?? '',
      textRunUnderlineImportCommandFields: stage?.getAttribute('data-ppt-text-run-underline-import-command-fields') ?? '',
      textRunUnderlineImportCommandIds: stage?.getAttribute('data-ppt-text-run-underline-import-command-ids') ?? '',
      textRunUnderlineImportCommandTargets: stage?.getAttribute('data-ppt-text-run-underline-import-command-targets') ?? '',
      textRunUnderlineImportCommandTypes: stage?.getAttribute('data-ppt-text-run-underline-import-command-types') ?? '',
      textRunUnderlineImportCommandValues: stage?.getAttribute('data-ppt-text-run-underline-import-command-values') ?? '',
      textRunUnderlineImportFields: stage?.getAttribute('data-ppt-text-run-underline-import-fields') ?? '',
      textRunUnderlineImportFormat: stage?.getAttribute('data-ppt-text-run-underline-import-format') ?? '',
      textRunUnderlineImportJsonLength: Number(stage?.getAttribute('data-ppt-text-run-underline-import-json-length') ?? 0),
      textRunUnderlineImportModel: stage?.getAttribute('data-ppt-text-run-underline-import-model') ?? '',
      textRunUnderlineImportObjects: stage?.getAttribute('data-ppt-text-run-underline-import-objects') ?? '',
      textRunUnderlineImportRuns: Number(stage?.getAttribute('data-ppt-text-run-underline-import-runs') ?? 0),
      textRunUnderlineImportSlide: stage?.getAttribute('data-ppt-text-run-underline-import-slide') ?? '',
      textRunUnderlineImportValue: stage?.getAttribute('data-ppt-text-run-underline-import-value') ?? '',
      textParagraphAlignImportCategories: stage?.getAttribute('data-ppt-text-paragraph-align-import-categories') ?? '',
      textParagraphAlignImportCommand: stage?.getAttribute('data-ppt-text-paragraph-align-import-command') ?? '',
      textParagraphAlignImportCommandTargets: stage?.getAttribute('data-ppt-text-paragraph-align-import-command-targets') ?? '',
      textParagraphAlignImportCommandType: stage?.getAttribute('data-ppt-text-paragraph-align-import-command-type') ?? '',
      textParagraphAlignImportFields: stage?.getAttribute('data-ppt-text-paragraph-align-import-fields') ?? '',
      textParagraphAlignImportFormat: stage?.getAttribute('data-ppt-text-paragraph-align-import-format') ?? '',
      textParagraphAlignImportJsonLength: Number(stage?.getAttribute('data-ppt-text-paragraph-align-import-json-length') ?? 0),
      textParagraphAlignImportModel: stage?.getAttribute('data-ppt-text-paragraph-align-import-model') ?? '',
      textParagraphAlignImportObjects: stage?.getAttribute('data-ppt-text-paragraph-align-import-objects') ?? '',
      textParagraphAlignImportValue: stage?.getAttribute('data-ppt-text-paragraph-align-import-value') ?? '',
      textParagraphBulletImportCategories: stage?.getAttribute('data-ppt-text-paragraph-bullet-import-categories') ?? '',
      textParagraphBulletImportCommand: stage?.getAttribute('data-ppt-text-paragraph-bullet-import-command') ?? '',
      textParagraphBulletImportCommandTargets: stage?.getAttribute('data-ppt-text-paragraph-bullet-import-command-targets') ?? '',
      textParagraphBulletImportCommandType: stage?.getAttribute('data-ppt-text-paragraph-bullet-import-command-type') ?? '',
      textParagraphBulletImportFields: stage?.getAttribute('data-ppt-text-paragraph-bullet-import-fields') ?? '',
      textParagraphBulletImportFormat: stage?.getAttribute('data-ppt-text-paragraph-bullet-import-format') ?? '',
      textParagraphBulletImportJsonLength: Number(stage?.getAttribute('data-ppt-text-paragraph-bullet-import-json-length') ?? 0),
      textParagraphBulletImportModel: stage?.getAttribute('data-ppt-text-paragraph-bullet-import-model') ?? '',
      textParagraphBulletImportObjects: stage?.getAttribute('data-ppt-text-paragraph-bullet-import-objects') ?? '',
      textParagraphBulletImportValue: stage?.getAttribute('data-ppt-text-paragraph-bullet-import-value') ?? '',
      textStyleImportCategories: stage?.getAttribute('data-ppt-text-style-import-categories') ?? '',
      textStyleImportColor: stage?.getAttribute('data-ppt-text-style-import-color') ?? '',
      textStyleImportCommand: stage?.getAttribute('data-ppt-text-style-import-command') ?? '',
      textStyleImportCommandTargets: stage?.getAttribute('data-ppt-text-style-import-command-targets') ?? '',
      textStyleImportCommandType: stage?.getAttribute('data-ppt-text-style-import-command-type') ?? '',
      textStyleImportFields: stage?.getAttribute('data-ppt-text-style-import-fields') ?? '',
      textStyleImportFontFamily: stage?.getAttribute('data-ppt-text-style-import-font-family') ?? '',
      textStyleImportFontSize: stage?.getAttribute('data-ppt-text-style-import-font-size') ?? '',
      textStyleImportFontWeight: stage?.getAttribute('data-ppt-text-style-import-font-weight') ?? '',
      textStyleImportFormat: stage?.getAttribute('data-ppt-text-style-import-format') ?? '',
      textStyleImportJsonLength: stage?.getAttribute('data-ppt-text-style-import-json-length') ?? '',
      textStyleImportModel: stage?.getAttribute('data-ppt-text-style-import-model') ?? '',
      textStyleImportObjects: stage?.getAttribute('data-ppt-text-style-import-objects') ?? '',
      textStyleImportParagraphAlign: stage?.getAttribute('data-ppt-text-style-import-paragraph-align') ?? '',
      textStyleImportParagraphBullet: stage?.getAttribute('data-ppt-text-style-import-paragraph-bullet') ?? '',
      textStyleImportParagraphLineHeight: stage?.getAttribute('data-ppt-text-style-import-paragraph-line-height') ?? '',
      textStyleImportParagraphSpacingAfter: stage?.getAttribute('data-ppt-text-style-import-paragraph-spacing-after') ?? '',
      textStyleImportParagraphSpacingBefore: stage?.getAttribute('data-ppt-text-style-import-paragraph-spacing-before') ?? '',
      textStyleImportRunBold: stage?.getAttribute('data-ppt-text-style-import-run-bold') ?? '',
      textStyleImportRunColor: stage?.getAttribute('data-ppt-text-style-import-run-color') ?? '',
      textStyleImportRunItalic: stage?.getAttribute('data-ppt-text-style-import-run-italic') ?? '',
      textStyleImportRunSize: stage?.getAttribute('data-ppt-text-style-import-run-size') ?? '',
      textStyleImportRunUnderline: stage?.getAttribute('data-ppt-text-style-import-run-underline') ?? '',
      textStyleImportTextInset: stage?.getAttribute('data-ppt-text-style-import-text-inset') ?? '',
      textStyleImportVerticalAlign: stage?.getAttribute('data-ppt-text-style-import-vertical-align') ?? '',
      top: element?.style.top ?? '',
      underlineRun: element?.querySelector('[data-ppt-run-underline="true"]')?.style.textDecoration ?? '',
      underlineRunCount: element?.querySelectorAll('[data-ppt-run-underline="true"]').length ?? 0,
      verticalAlign: element?.getAttribute('data-ppt-vertical-align') ?? '',
      width: element?.style.width ?? '',
    }
  })(${JSON.stringify(elementId)})`)
}

async function readViewportState(page) {
  return page.eval(`(() => {
    const shell = document.querySelector('.ppt-stage-shell')
    const transform = document.querySelector('.ppt-stage-world')?.style.transform ?? ''
    const scale = Number(transform.match(/scale\\(([^)]+)\\)/)?.[1] ?? 0)
    const translate = transform.match(/translate\\(([^p]+)px, ([^p]+)px\\)/)

    return {
      keyboardIntent: shell?.getAttribute('data-ppt-keyboard-viewport-intent') ?? '',
      keyboardModel: shell?.getAttribute('data-ppt-keyboard-viewport-model') ?? '',
      label: document.querySelector('.ppt-zoom-label')?.textContent ?? '',
      paletteOpen: !!document.querySelector('[data-ppt-command-palette]'),
      scale,
      transform,
      x: Number(translate?.[1] ?? 0),
      y: Number(translate?.[2] ?? 0),
    }
  })()`)
}

async function readPPTTemporaryPanState(page) {
  return page.eval(`(() => {
    const shell = document.querySelector('.ppt-stage-shell')
    const transform = document.querySelector('.ppt-stage-world')?.style.transform ?? ''
    const scale = Number(transform.match(/scale\\(([^)]+)\\)/)?.[1] ?? 0)
    const translate = transform.match(/translate\\(([^p]+)px, ([^p]+)px\\)/)
    const selectedIds = [...document.querySelectorAll('[data-selected="true"]')]
      .map((element) => element.getAttribute('data-ppt-element') ?? '')
      .filter(Boolean)
      .join(' ')

    return {
      active: shell?.getAttribute('data-ppt-temporary-pan-active') ?? '',
      cursor: shell ? getComputedStyle(shell).cursor : '',
      gesture: shell?.getAttribute('data-ppt-temporary-pan-gesture') ?? '',
      model: shell?.getAttribute('data-ppt-temporary-pan-model') ?? '',
      scale,
      selectedCount: selectedIds ? selectedIds.split(' ').length : 0,
      selectedIds,
      shortcut: shell?.getAttribute('data-ppt-temporary-pan-shortcut') ?? '',
      transform,
      x: Number(translate?.[1] ?? 0),
      y: Number(translate?.[2] ?? 0),
    }
  })()`)
}

async function readPPTWheelViewportState(page) {
  return page.eval(`(() => {
    const shell = document.querySelector('.ppt-stage-shell')
    const transform = document.querySelector('.ppt-stage-world')?.style.transform ?? ''
    const scale = Number(transform.match(/scale\\(([^)]+)\\)/)?.[1] ?? 0)
    const translate = transform.match(/translate\\(([^p]+)px, ([^p]+)px\\)/)

    return {
      horizontalPanModifier: shell?.getAttribute('data-ppt-wheel-viewport-horizontal-pan-modifier') ?? '',
      model: shell?.getAttribute('data-ppt-wheel-viewport-model') ?? '',
      pan: shell?.getAttribute('data-ppt-wheel-viewport-pan') ?? '',
      scale,
      transform,
      x: Number(translate?.[1] ?? 0),
      y: Number(translate?.[2] ?? 0),
      zoomModifier: shell?.getAttribute('data-ppt-wheel-viewport-zoom-modifier') ?? '',
    }
  })()`)
}

async function readPPTPanToolState(page) {
  return page.eval(`(() => {
    const shell = document.querySelector('.ppt-stage-shell')
    const transform = document.querySelector('.ppt-stage-world')?.style.transform ?? ''
    const scale = Number(transform.match(/scale\\(([^)]+)\\)/)?.[1] ?? 0)
    const translate = transform.match(/translate\\(([^p]+)px, ([^p]+)px\\)/)
    const selectedIds = [...document.querySelectorAll('[data-selected="true"]')]
      .map((element) => element.getAttribute('data-ppt-element') ?? '')
      .filter(Boolean)
      .join(' ')

    return {
      active: shell?.getAttribute('data-ppt-pan-tool-active') ?? '',
      cursor: shell ? getComputedStyle(shell).cursor : '',
      model: shell?.getAttribute('data-ppt-pan-tool-model') ?? '',
      paletteOpen: !!document.querySelector('[data-ppt-command-palette]'),
      scale,
      selectedIds,
      shortcut: shell?.getAttribute('data-ppt-pan-tool-shortcut') ?? '',
      toolbarPressed: document.querySelector('[data-ppt-pan-tool]')?.getAttribute('aria-pressed') ?? '',
      transform,
      x: Number(translate?.[1] ?? 0),
      y: Number(translate?.[2] ?? 0),
    }
  })()`)
}

async function readPPTLaserToolState(page) {
  return page.eval(`(() => {
    const shell = document.querySelector('.ppt-stage-shell')
    const transform = document.querySelector('.ppt-stage-world')?.style.transform ?? ''
    const scale = Number(transform.match(/scale\\(([^)]+)\\)/)?.[1] ?? 0)
    const translate = transform.match(/translate\\(([^p]+)px, ([^p]+)px\\)/)
    const selectedIds = [...document.querySelectorAll('[data-selected="true"]')]
      .map((element) => element.getAttribute('data-ppt-element') ?? '')
      .filter(Boolean)
      .join(' ')
    const trail = document.querySelector('[data-ppt-laser-trail]')

    return {
      active: shell?.getAttribute('data-ppt-laser-tool-active') ?? '',
      cursor: shell ? getComputedStyle(shell).cursor : '',
      elementCount: document.querySelectorAll('[data-ppt-element]').length,
      model: shell?.getAttribute('data-ppt-laser-tool-model') ?? '',
      paletteOpen: !!document.querySelector('[data-ppt-command-palette]'),
      scale,
      selectedIds,
      shortcut: shell?.getAttribute('data-ppt-laser-tool-shortcut') ?? '',
      toolbarPressed: document.querySelector('[data-ppt-laser-tool]')?.getAttribute('aria-pressed') ?? '',
      trailDot: !!document.querySelector('[data-ppt-laser-trail-dot]'),
      trailModel: shell?.getAttribute('data-ppt-laser-trail-model') ?? '',
      trailPath: !!document.querySelector('[data-ppt-laser-trail-path]'),
      trailPointCount: Number(shell?.getAttribute('data-ppt-laser-trail-point-count') ?? 0),
      trailRenderedPointCount: Number(trail?.getAttribute('data-ppt-laser-trail-point-count') ?? 0),
      trailState: shell?.getAttribute('data-ppt-laser-trail-state') ?? '',
      transform,
      x: Number(translate?.[1] ?? 0),
      y: Number(translate?.[2] ?? 0),
    }
  })()`)
}

async function readPPTMinimapState(page) {
  return page.eval(`(() => {
    const minimap = document.querySelector('[data-ppt-minimap]')
    const viewport = document.querySelector('[data-ppt-minimap-viewport]')
    const viewportRect = viewport?.getBoundingClientRect()
    const world = document.querySelector('[data-ppt-minimap-world]')
    const stage = document.querySelector('.ppt-stage-shell')
    const view = {
      h: Number(viewport?.getAttribute('height') ?? 0),
      w: Number(viewport?.getAttribute('width') ?? 0),
      x: Number(viewport?.getAttribute('x') ?? 0),
      y: Number(viewport?.getAttribute('y') ?? 0),
    }

    return {
      hasViewport: !!viewport,
      hasWorld: !!world,
      itemCount: Number(minimap?.getAttribute('data-ppt-minimap-item-count') ?? 0),
      model: minimap?.getAttribute('data-ppt-minimap-model') ?? '',
      open: !!minimap,
      paletteOpen: !!document.querySelector('[data-ppt-command-palette]'),
      scale: Number(document.querySelector('.ppt-stage-world')?.style.transform.match(/scale\\(([^)]+)\\)/)?.[1] ?? 0),
      stageAttr: stage?.getAttribute('data-minimap') ?? '',
      togglePressed: document.querySelector('[data-ppt-view-minimap]')?.getAttribute('aria-pressed') ?? '',
      view,
      viewportBoxWidth: viewportRect?.width ?? 0,
      viewportH: Number(minimap?.getAttribute('data-ppt-minimap-viewport-h') ?? 0),
      viewportW: Number(minimap?.getAttribute('data-ppt-minimap-viewport-w') ?? 0),
      viewportX: Number(minimap?.getAttribute('data-ppt-minimap-viewport-x') ?? 0),
      viewportY: Number(minimap?.getAttribute('data-ppt-minimap-viewport-y') ?? 0),
    }
  })()`)
}

async function runTextQuickFormatScenario(page) {
  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)

  let titlePoint = await getElementCenter(page, 's1-title')
  await clickMouse(page, titlePoint.x, titlePoint.y, 1)
  await delay(80)

  const initial = await page.eval(`(() => ({
    bulletPressed: document.querySelector('[data-ppt-text-quick="bullet"]')?.getAttribute('aria-pressed') ?? '',
    boldPressed: document.querySelector('[data-ppt-text-quick="bold"]')?.getAttribute('aria-pressed') ?? '',
    fontSize: Number(document.querySelector('[data-ppt-style-field="font-size"]')?.value ?? 0),
    italicPressed: document.querySelector('[data-ppt-text-quick="italic"]')?.getAttribute('aria-pressed') ?? '',
    numberedPressed: document.querySelector('[data-ppt-text-quick="numbered"]')?.getAttribute('aria-pressed') ?? '',
    quickBarVisible: !!document.querySelector('[data-ppt-text-quick-bar]'),
    selectedId: document.querySelector('[data-selected="true"]')?.getAttribute('data-ppt-element') ?? '',
    underlinePressed: document.querySelector('[data-ppt-text-quick="underline"]')?.getAttribute('aria-pressed') ?? '',
  }))()`)

  record('renders PPT text quick format bar for selected text', initial.quickBarVisible && initial.selectedId === 's1-title' && initial.boldPressed === 'true' && initial.bulletPressed === 'false' && initial.numberedPressed === 'false' && initial.italicPressed === 'false' && initial.underlinePressed === 'false' && initial.fontSize > 0, initial)

  const initialAlignRadio = await readPPTParagraphAlignRadioGroupState(page)

  record(
    'exposes PPT paragraph align as APG radiogroups',
    initialAlignRadio.quick.role === 'radiogroup' &&
      initialAlignRadio.quick.focusModel === 'roving-tabindex' &&
      initialAlignRadio.quick.keyboardModel === 'arrow-home-end' &&
      initialAlignRadio.quick.model === 'canvas-radio-group' &&
      initialAlignRadio.quick.radioCount === 3 &&
      initialAlignRadio.quick.checkedValues.includes('left') &&
      initialAlignRadio.quick.tabStopValues.includes('left') &&
      initialAlignRadio.inspector.role === 'radiogroup' &&
      initialAlignRadio.inspector.focusModel === 'roving-tabindex' &&
      initialAlignRadio.inspector.keyboardModel === 'arrow-home-end' &&
      initialAlignRadio.inspector.model === 'canvas-radio-group' &&
      initialAlignRadio.inspector.radioCount === 3 &&
      initialAlignRadio.inspector.checkedValues.includes('left') &&
      initialAlignRadio.inspector.tabStopValues.includes('left'),
    initialAlignRadio,
  )

  await page.eval(`document.querySelector('[data-ppt-paragraph-align-radiogroup="quick"] [data-ppt-paragraph-align="left"]')?.focus()`)
  await pressKey(page, {
    code: 'ArrowRight',
    key: 'ArrowRight',
    windowsVirtualKeyCode: 39,
  })
  await delay(80)

  const afterAlignArrow = await readPPTParagraphAlignRadioGroupState(page)

  await pressKey(page, {
    code: 'Home',
    key: 'Home',
    windowsVirtualKeyCode: 36,
  })
  await delay(80)

  const afterAlignHome = await readPPTParagraphAlignRadioGroupState(page)

  record(
    'updates PPT paragraph align radiogroup with Arrow/Home keys',
    afterAlignArrow.quick.focusedValue === 'center' &&
      afterAlignArrow.quick.checkedValues.includes('center') &&
      afterAlignArrow.inspector.checkedValues.includes('center') &&
      afterAlignArrow.selectedTextAlign === 'center' &&
      afterAlignHome.quick.focusedValue === 'left' &&
      afterAlignHome.quick.checkedValues.includes('left') &&
      afterAlignHome.inspector.checkedValues.includes('left') &&
      afterAlignHome.selectedTextAlign === 'left',
    {
      afterAlignArrow,
      afterAlignHome,
      initialAlignRadio,
    },
  )

  const beforeTextColorSwatch = await getPPTColorSwatchState(page, 'text-color', 's1-title')

  await page.eval(`document.querySelector('[data-ppt-color-swatch="text-color"][data-ppt-color-token="ppt-color-accent"]')?.click()`)
  await delay(100)

  const afterTextColorSwatch = await getPPTColorSwatchState(page, 'text-color', 's1-title')

  record(
    'applies PPT text theme color swatch to text color model',
    beforeTextColorSwatch.themeCount >= 4 &&
      afterTextColorSwatch.model === 'slide-edit-color-swatch-palette' &&
      afterTextColorSwatch.descriptorModel === 'color-swatch-palette' &&
      afterTextColorSwatch.descriptorCommand === 'apply-color-swatch' &&
      afterTextColorSwatch.descriptorControl === 'color-swatch-palette' &&
      afterTextColorSwatch.descriptorDisabled === 'false' &&
      afterTextColorSwatch.packageChannel === 'text' &&
      afterTextColorSwatch.command === 'apply-color-swatch' &&
      afterTextColorSwatch.commandChannel === 'text' &&
      afterTextColorSwatch.commandObjects.includes('s1-title') &&
      afterTextColorSwatch.commandSource === 'theme' &&
      afterTextColorSwatch.commandSwatch === 'theme:ppt-color-accent' &&
      afterTextColorSwatch.commandToken === 'ppt-color-accent' &&
      afterTextColorSwatch.commandType === 'slide-command-effect' &&
      afterTextColorSwatch.commandValue === '#2563eb' &&
      afterTextColorSwatch.inputValue === '#2563eb' &&
      afterTextColorSwatch.styleColor === 'rgb(37, 99, 235)' &&
      afterTextColorSwatch.selectedToken === 'ppt-color-accent' &&
      afterTextColorSwatch.recentColors.includes('#2563eb') &&
      afterTextColorSwatch.recentUnique &&
      afterTextColorSwatch.exportSlice.includes('"color": "#2563eb"'),
    {
      afterTextColorSwatch,
      beforeTextColorSwatch,
    },
  )

  await page.eval(`document.querySelector('[data-ppt-present-start]')?.click()`)
  await delay(120)

  const textColorPreview = await page.eval(`(() => {
    const overlay = document.querySelector('[data-ppt-presentation]')
    const element = overlay?.querySelector('[data-ppt-element="s1-title"]')

    return {
      color: element?.style.color ?? '',
      open: !!overlay,
    }
  })()`)

  record(
    'keeps PPT text theme color swatch in presentation preview',
    textColorPreview.open && textColorPreview.color === 'rgb(37, 99, 235)',
    textColorPreview,
  )

  await page.eval(`document.querySelector('[data-ppt-present-exit]')?.click()`)
  await delay(80)

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      colorSwatch: {
        channel: 'text-color',
        color: '#0f766e',
        source: 'recent',
      },
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const afterTextColorJSONPaste = await getPPTColorSwatchState(page, 'text-color', 's1-title')

  record(
    'pastes PPT color swatch JSON through slide-edit color command effect',
    afterTextColorJSONPaste.model === 'slide-edit-color-swatch-palette' &&
      afterTextColorJSONPaste.importModel === 'ppt-color-swatch-import' &&
      afterTextColorJSONPaste.importFormat === 'application-json-ppt-color-swatch' &&
      afterTextColorJSONPaste.importFields === 'channel color source' &&
      afterTextColorJSONPaste.importCommands === 'apply-color-swatch' &&
      afterTextColorJSONPaste.importCommandChannels === 'text' &&
      afterTextColorJSONPaste.importCommandTargets === 's1-title' &&
      afterTextColorJSONPaste.importCommandSources === 'recent' &&
      afterTextColorJSONPaste.importCommandSwatches === 'recent:#0f766e' &&
      afterTextColorJSONPaste.importCommandTypes === 'slide-command-effect' &&
      afterTextColorJSONPaste.importCommandValues === '#0f766e' &&
      afterTextColorJSONPaste.importObjects === 's1-title' &&
      afterTextColorJSONPaste.importSource === 'recent' &&
      afterTextColorJSONPaste.importValue === '#0f766e' &&
      afterTextColorJSONPaste.importJsonLength > 60 &&
      afterTextColorJSONPaste.command === 'apply-color-swatch' &&
      afterTextColorJSONPaste.commandChannel === 'text' &&
      afterTextColorJSONPaste.commandObjects.includes('s1-title') &&
      afterTextColorJSONPaste.commandSource === 'recent' &&
      afterTextColorJSONPaste.commandSwatch === 'recent:#0f766e' &&
      afterTextColorJSONPaste.commandType === 'slide-command-effect' &&
      afterTextColorJSONPaste.commandValue === '#0f766e' &&
      afterTextColorJSONPaste.inputValue === '#0f766e' &&
      afterTextColorJSONPaste.styleColor === 'rgb(15, 118, 110)' &&
      afterTextColorJSONPaste.recentColors.includes('#0f766e') &&
      afterTextColorJSONPaste.recentUnique &&
      afterTextColorJSONPaste.exportSlice.includes('"color": "#0f766e"'),
    {
      afterTextColorJSONPaste,
      afterTextColorSwatch,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(100)

  const afterTextColorJSONUndo = await getPPTColorSwatchState(page, 'text-color', 's1-title')

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(100)

  const afterTextColorJSONRedo = await getPPTColorSwatchState(page, 'text-color', 's1-title')

  record(
    'undoes and redoes PPT color swatch JSON as one history step',
    afterTextColorJSONUndo.styleColor === 'rgb(37, 99, 235)' &&
      afterTextColorJSONUndo.inputValue === '#2563eb' &&
      afterTextColorJSONRedo.styleColor === afterTextColorJSONPaste.styleColor &&
      afterTextColorJSONRedo.inputValue === '#0f766e',
    {
      afterTextColorJSONPaste,
      afterTextColorJSONRedo,
      afterTextColorJSONUndo,
    },
  )

  await page.eval(`document.querySelector('[data-ppt-text-quick="bold"]')?.click()`)
  await page.eval(`document.querySelector('[data-ppt-text-quick="italic"]')?.click()`)
  await page.eval(`document.querySelector('[data-ppt-text-quick="underline"]')?.click()`)
  await page.eval(`document.querySelector('[data-ppt-text-quick="font-size-up"]')?.click()`)
  await setTextQuickColor(page, '#0055ff')
  await page.eval(`document.querySelector('[data-ppt-text-quick="align-right"]')?.click()`)
  await page.eval(`document.querySelector('[data-ppt-text-quick="bullet"]')?.click()`)
  await delay(100)

  const afterSingleFormat = await page.eval(`(() => {
    const title = document.querySelector('[data-ppt-element="s1-title"]')

    return {
      bulletList: title?.getAttribute('data-ppt-bullet-list') ?? '',
      bulletPressed: document.querySelector('[data-ppt-text-quick="bullet"]')?.getAttribute('aria-pressed') ?? '',
      color: title?.style.color ?? '',
      fontSize: Number(document.querySelector('[data-ppt-style-field="font-size"]')?.value ?? 0),
      fontWeight: document.querySelector('[data-ppt-style-field="font-weight"]')?.value ?? '',
      inspectorBulletPressed: document.querySelector('[data-ppt-paragraph-bullet]')?.getAttribute('aria-pressed') ?? '',
      inspectorNumberedPressed: document.querySelector('[data-ppt-paragraph-numbered]')?.getAttribute('aria-pressed') ?? '',
      italicPressed: document.querySelector('[data-ppt-text-quick="italic"]')?.getAttribute('aria-pressed') ?? '',
      italicRun: title?.querySelector('[data-ppt-run-italic="true"]')?.style.fontStyle ?? '',
      paragraphBullet: title?.querySelector('[data-ppt-bullet="true"]')?.textContent ?? '',
      numberedList: title?.getAttribute('data-ppt-numbered-list') ?? '',
      numberedPressed: document.querySelector('[data-ppt-text-quick="numbered"]')?.getAttribute('aria-pressed') ?? '',
      rightPressed: document.querySelector('[data-ppt-paragraph-align="right"]')?.getAttribute('aria-pressed') ?? '',
      textAlign: title?.style.textAlign ?? '',
      thumbBulletCount: document.querySelectorAll('[data-ppt-thumb-bullet="true"]').length,
      underlinePressed: document.querySelector('[data-ppt-text-quick="underline"]')?.getAttribute('aria-pressed') ?? '',
      underlineRun: title?.querySelector('[data-ppt-run-underline="true"]')?.style.textDecoration ?? '',
    }
  })()`)

  record('applies PPT text quick formatting to selected text model', afterSingleFormat.color === 'rgb(0, 85, 255)' && afterSingleFormat.fontSize === initial.fontSize + 2 && afterSingleFormat.fontWeight === 'regular' && afterSingleFormat.textAlign === 'right' && afterSingleFormat.rightPressed === 'true' && afterSingleFormat.bulletList === 'true' && afterSingleFormat.numberedList === '' && afterSingleFormat.bulletPressed === 'true' && afterSingleFormat.numberedPressed === 'false' && afterSingleFormat.inspectorBulletPressed === 'true' && afterSingleFormat.inspectorNumberedPressed === 'false' && afterSingleFormat.italicPressed === 'true' && afterSingleFormat.underlinePressed === 'true' && afterSingleFormat.italicRun === 'italic' && afterSingleFormat.underlineRun.includes('underline') && afterSingleFormat.paragraphBullet.length > 0 && afterSingleFormat.thumbBulletCount > 0, {
    afterSingleFormat,
    initial,
  })

  await page.eval(`document.querySelector('[data-ppt-command="copy-formatting"]')?.click()`)
  await delay(80)

  const summaryBeforeFormatPaste = await getPPTTextFormatPainterState(page, 's1-summary')
  const summaryFormatPoint = await getElementCenter(page, 's1-summary')

  await clickMouse(page, summaryFormatPoint.x, summaryFormatPoint.y, 1)
  await delay(80)
  await page.eval(`document.querySelector('[data-ppt-command="paste-formatting"]')?.click()`)
  await delay(100)

  const summaryAfterFormatPaste = await getPPTTextFormatPainterState(page, 's1-summary')

  record(
    'pastes PPT text formatting into text target without changing content geometry or name',
    summaryBeforeFormatPaste.text === summaryAfterFormatPaste.text &&
      summaryBeforeFormatPaste.name === summaryAfterFormatPaste.name &&
      summaryBeforeFormatPaste.left === summaryAfterFormatPaste.left &&
      summaryBeforeFormatPaste.top === summaryAfterFormatPaste.top &&
      summaryBeforeFormatPaste.width === summaryAfterFormatPaste.width &&
      summaryBeforeFormatPaste.height === summaryAfterFormatPaste.height &&
      summaryAfterFormatPaste.color === 'rgb(0, 85, 255)' &&
      summaryAfterFormatPaste.fontSize === `${afterSingleFormat.fontSize}px` &&
      summaryAfterFormatPaste.textAlign === 'right' &&
      summaryAfterFormatPaste.bulletList === 'true' &&
      summaryAfterFormatPaste.paragraphBullet.length > 0 &&
      summaryAfterFormatPaste.italicRun === 'italic' &&
      summaryAfterFormatPaste.italicRunCount > 0 &&
      summaryAfterFormatPaste.underlineRun.includes('underline') &&
      summaryAfterFormatPaste.underlineRunCount > 0 &&
      summaryAfterFormatPaste.styleClipboardCategories.includes('text-run') &&
      summaryAfterFormatPaste.styleClipboardPackageCategories.includes('text-run-style') &&
      summaryAfterFormatPaste.styleClipboardCommandApplications.includes('text-run-style') &&
      summaryAfterFormatPaste.styleClipboardRunItalic === 'true' &&
      summaryAfterFormatPaste.styleClipboardRunUnderline === 'true',
    {
      afterSingleFormat,
      summaryAfterFormatPaste,
      summaryBeforeFormatPaste,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const summaryBeforeTextStylePaste = await getPPTTextFormatPainterState(page, 's1-summary')

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      textStyle: {
        color: '#7c3aed',
        fontFamily: 'Georgia',
        fontSize: 34,
        fontWeight: 'semibold',
        paragraph: {
          align: 'center',
          bullet: 'numbered',
          lineHeight: 1.32,
          spacingAfter: 10,
          spacingBefore: 4,
        },
        runStyle: {
          bold: false,
          italic: false,
          underline: false,
        },
        textInset: {
          bottom: 12,
          left: 14,
          right: 10,
          top: 8,
        },
        verticalAlign: 'middle',
      },
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const summaryAfterTextStylePaste = await getPPTTextFormatPainterState(page, 's1-summary')

  record(
    'pastes JSON text style into selected PPT text object',
    summaryBeforeTextStylePaste.text === summaryAfterTextStylePaste.text &&
      summaryBeforeTextStylePaste.name === summaryAfterTextStylePaste.name &&
      summaryBeforeTextStylePaste.left === summaryAfterTextStylePaste.left &&
      summaryBeforeTextStylePaste.top === summaryAfterTextStylePaste.top &&
      summaryBeforeTextStylePaste.width === summaryAfterTextStylePaste.width &&
      summaryBeforeTextStylePaste.height === summaryAfterTextStylePaste.height &&
      summaryAfterTextStylePaste.selected === 'true' &&
      summaryAfterTextStylePaste.textStyleImportModel === 'ppt-text-style-import' &&
      summaryAfterTextStylePaste.textStyleImportFormat === 'application-json-ppt-text-style' &&
      summaryAfterTextStylePaste.textStyleImportCommand === 'paste-object-formatting' &&
      summaryAfterTextStylePaste.textStyleImportCommandTargets === 's1-summary' &&
      summaryAfterTextStylePaste.textStyleImportCommandType === 'slide-command-effect' &&
      summaryAfterTextStylePaste.textStyleImportObjects === 's1-summary' &&
      summaryAfterTextStylePaste.textStyleImportCategories.includes('object-effect') &&
      summaryAfterTextStylePaste.textStyleImportCategories.includes('text-style') &&
      summaryAfterTextStylePaste.textStyleImportCategories.includes('text-run-style') &&
      summaryAfterTextStylePaste.textStyleImportFields === 'color fontSize fontFamily fontWeight verticalAlign textInset paragraphAlign paragraphBullet paragraphLineHeight paragraphSpacingBefore paragraphSpacingAfter runBold runItalic runUnderline' &&
      summaryAfterTextStylePaste.textStyleImportColor === '#7c3aed' &&
      summaryAfterTextStylePaste.textStyleImportFontFamily === 'Georgia' &&
      summaryAfterTextStylePaste.textStyleImportFontSize === '34' &&
      summaryAfterTextStylePaste.textStyleImportFontWeight === 'semibold' &&
      summaryAfterTextStylePaste.textStyleImportVerticalAlign === 'middle' &&
      summaryAfterTextStylePaste.textStyleImportTextInset === '8,10,12,14' &&
      summaryAfterTextStylePaste.textStyleImportParagraphAlign === 'center' &&
      summaryAfterTextStylePaste.textStyleImportParagraphBullet === 'numbered' &&
      summaryAfterTextStylePaste.textStyleImportParagraphLineHeight === '1.32' &&
      summaryAfterTextStylePaste.textStyleImportParagraphSpacingAfter === '10' &&
      summaryAfterTextStylePaste.textStyleImportParagraphSpacingBefore === '4' &&
      summaryAfterTextStylePaste.textStyleImportRunBold === 'false' &&
      summaryAfterTextStylePaste.textStyleImportRunColor === '' &&
      summaryAfterTextStylePaste.textStyleImportRunItalic === 'false' &&
      summaryAfterTextStylePaste.textStyleImportRunSize === '' &&
      summaryAfterTextStylePaste.textStyleImportRunUnderline === 'false' &&
      Number(summaryAfterTextStylePaste.textStyleImportJsonLength) > 220 &&
      summaryAfterTextStylePaste.color === 'rgb(124, 58, 237)' &&
      summaryAfterTextStylePaste.fontFamily === 'Georgia' &&
      summaryAfterTextStylePaste.styleFontFamily.includes('Georgia') &&
      summaryAfterTextStylePaste.fontSize === '34px' &&
      summaryAfterTextStylePaste.fontWeight === '600' &&
      summaryAfterTextStylePaste.boldRunCount === 0 &&
      summaryAfterTextStylePaste.italicRunCount === 0 &&
      summaryAfterTextStylePaste.runColorCount === 0 &&
      summaryAfterTextStylePaste.runSizeCount === 0 &&
      summaryAfterTextStylePaste.underlineRunCount === 0 &&
      summaryAfterTextStylePaste.verticalAlign === 'middle' &&
      summaryAfterTextStylePaste.styleAlignItems === 'center' &&
      summaryAfterTextStylePaste.textInset === '8,10,12,14' &&
      summaryAfterTextStylePaste.stylePadding === '8px 10px 12px 14px' &&
      summaryAfterTextStylePaste.textAlign === 'center' &&
      summaryAfterTextStylePaste.numberedList === 'true' &&
      summaryAfterTextStylePaste.paragraphList === 'numbered' &&
      summaryAfterTextStylePaste.paragraphLineHeight === '1.32' &&
      summaryAfterTextStylePaste.paragraphSpacingAfter === '10' &&
      summaryAfterTextStylePaste.paragraphSpacingBefore === '4' &&
      summaryAfterTextStylePaste.styleClipboardCommand === 'paste-object-formatting' &&
      summaryAfterTextStylePaste.styleClipboardCommandTargets === 's1-summary' &&
      summaryAfterTextStylePaste.styleClipboardCommandApplications.includes('s1-summary') &&
      summaryAfterTextStylePaste.styleClipboardCommandApplications.includes('text-style') &&
      summaryAfterTextStylePaste.styleClipboardCommandApplications.includes('text-run-style') &&
      summaryAfterTextStylePaste.styleClipboardPackageCategories.includes('text-run-style') &&
      summaryAfterTextStylePaste.styleClipboardRunItalic === 'false' &&
      summaryAfterTextStylePaste.styleClipboardRunUnderline === 'false',
    {
      summaryAfterTextStylePaste,
      summaryBeforeTextStylePaste,
    },
  )

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      textRunStyle: {
        color: '#0f766e',
        italic: true,
        size: '31px',
      },
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const summaryAfterStandaloneTextRunStylePaste = await getPPTTextFormatPainterState(page, 's1-summary')

  record(
    'pastes standalone PPT textRunStyle JSON through style clipboard run category',
    summaryAfterStandaloneTextRunStylePaste.selected === 'true' &&
      summaryAfterStandaloneTextRunStylePaste.text === summaryAfterTextStylePaste.text &&
      summaryAfterStandaloneTextRunStylePaste.textStyleImportModel === 'ppt-text-style-import' &&
      summaryAfterStandaloneTextRunStylePaste.textStyleImportFormat === 'application-json-ppt-text-style' &&
      summaryAfterStandaloneTextRunStylePaste.textStyleImportCommand === 'paste-object-formatting' &&
      summaryAfterStandaloneTextRunStylePaste.textStyleImportCommandTargets === 's1-summary' &&
      summaryAfterStandaloneTextRunStylePaste.textStyleImportCommandType === 'slide-command-effect' &&
      summaryAfterStandaloneTextRunStylePaste.textStyleImportCategories.includes('object-effect') &&
      summaryAfterStandaloneTextRunStylePaste.textStyleImportCategories.includes('text-run-style') &&
      !summaryAfterStandaloneTextRunStylePaste.textStyleImportCategories.includes('text-style') &&
      summaryAfterStandaloneTextRunStylePaste.textStyleImportFields === 'runColor runItalic runSize' &&
      summaryAfterStandaloneTextRunStylePaste.textStyleImportRunColor === '#0f766e' &&
      summaryAfterStandaloneTextRunStylePaste.textStyleImportRunItalic === 'true' &&
      summaryAfterStandaloneTextRunStylePaste.textStyleImportRunSize === '31' &&
      summaryAfterStandaloneTextRunStylePaste.runColor === 'rgb(15, 118, 110)' &&
      summaryAfterStandaloneTextRunStylePaste.runColorCount > 0 &&
      summaryAfterStandaloneTextRunStylePaste.runColorValue === '#0f766e' &&
      summaryAfterStandaloneTextRunStylePaste.italicRun === 'italic' &&
      summaryAfterStandaloneTextRunStylePaste.italicRunCount > 0 &&
      summaryAfterStandaloneTextRunStylePaste.runSize === '31px' &&
      summaryAfterStandaloneTextRunStylePaste.runSizeCount > 0 &&
      summaryAfterStandaloneTextRunStylePaste.runSizeValue === '31' &&
      summaryAfterStandaloneTextRunStylePaste.styleClipboardCategories.includes('text-run') &&
      summaryAfterStandaloneTextRunStylePaste.styleClipboardPackageCategories.includes('text-run-style') &&
      summaryAfterStandaloneTextRunStylePaste.styleClipboardCommandApplications.includes('text-run-style') &&
      summaryAfterStandaloneTextRunStylePaste.styleClipboardRunItalic === 'true',
    {
      summaryAfterStandaloneTextRunStylePaste,
      summaryAfterTextStylePaste,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const summaryAfterStandaloneTextRunStyleUndo = await getPPTTextFormatPainterState(page, 's1-summary')

  record(
    'undoes standalone PPT textRunStyle JSON as one history step',
    summaryAfterStandaloneTextRunStyleUndo.runColorCount === 0 &&
      summaryAfterStandaloneTextRunStyleUndo.italicRunCount === 0 &&
      summaryAfterStandaloneTextRunStyleUndo.runSizeCount === 0 &&
      summaryAfterStandaloneTextRunStyleUndo.textAlign === summaryAfterTextStylePaste.textAlign &&
      summaryAfterStandaloneTextRunStyleUndo.fontSize === summaryAfterTextStylePaste.fontSize,
    {
      summaryAfterStandaloneTextRunStylePaste,
      summaryAfterStandaloneTextRunStyleUndo,
      summaryAfterTextStylePaste,
    },
  )

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      textParagraphStyle: {
        align: 'left',
        bullet: null,
        lineHeight: 1.08,
        spacingAfter: 6,
        spacingBefore: 2,
      },
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const summaryAfterStandaloneParagraphStylePaste = await getPPTTextFormatPainterState(page, 's1-summary')

  record(
    'pastes standalone PPT textParagraphStyle JSON through style clipboard text category',
    summaryAfterStandaloneParagraphStylePaste.selected === 'true' &&
      summaryAfterStandaloneParagraphStylePaste.text === summaryAfterTextStylePaste.text &&
      summaryAfterStandaloneParagraphStylePaste.textStyleImportModel === 'ppt-text-style-import' &&
      summaryAfterStandaloneParagraphStylePaste.textStyleImportFormat === 'application-json-ppt-text-style' &&
      summaryAfterStandaloneParagraphStylePaste.textStyleImportCommand === 'paste-object-formatting' &&
      summaryAfterStandaloneParagraphStylePaste.textStyleImportCommandTargets === 's1-summary' &&
      summaryAfterStandaloneParagraphStylePaste.textStyleImportCommandType === 'slide-command-effect' &&
      summaryAfterStandaloneParagraphStylePaste.textStyleImportCategories.includes('object-effect') &&
      summaryAfterStandaloneParagraphStylePaste.textStyleImportCategories.includes('text-style') &&
      !summaryAfterStandaloneParagraphStylePaste.textStyleImportCategories.includes('text-run-style') &&
      summaryAfterStandaloneParagraphStylePaste.textStyleImportFields === 'paragraphAlign paragraphBullet paragraphLineHeight paragraphSpacingBefore paragraphSpacingAfter' &&
      summaryAfterStandaloneParagraphStylePaste.textStyleImportParagraphAlign === 'left' &&
      summaryAfterStandaloneParagraphStylePaste.textStyleImportParagraphBullet === '' &&
      summaryAfterStandaloneParagraphStylePaste.textStyleImportParagraphLineHeight === '1.08' &&
      summaryAfterStandaloneParagraphStylePaste.textStyleImportParagraphSpacingAfter === '6' &&
      summaryAfterStandaloneParagraphStylePaste.textStyleImportParagraphSpacingBefore === '2' &&
      summaryAfterStandaloneParagraphStylePaste.textAlign === 'left' &&
      summaryAfterStandaloneParagraphStylePaste.numberedList === '' &&
      summaryAfterStandaloneParagraphStylePaste.bulletList === '' &&
      summaryAfterStandaloneParagraphStylePaste.paragraphList === '' &&
      summaryAfterStandaloneParagraphStylePaste.paragraphLineHeight === '1.08' &&
      summaryAfterStandaloneParagraphStylePaste.paragraphSpacingAfter === '6' &&
      summaryAfterStandaloneParagraphStylePaste.paragraphSpacingBefore === '2' &&
      summaryAfterStandaloneParagraphStylePaste.styleClipboardCategories.includes('paragraph') &&
      summaryAfterStandaloneParagraphStylePaste.styleClipboardPackageCategories.includes('text-style') &&
      summaryAfterStandaloneParagraphStylePaste.styleClipboardCommandApplications.includes('text-style') &&
      !summaryAfterStandaloneParagraphStylePaste.styleClipboardPackageCategories.includes('text-run-style'),
    {
      summaryAfterStandaloneParagraphStylePaste,
      summaryAfterTextStylePaste,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const summaryAfterStandaloneParagraphStyleUndo = await getPPTTextFormatPainterState(page, 's1-summary')

  record(
    'undoes standalone PPT textParagraphStyle JSON as one history step',
    summaryAfterStandaloneParagraphStyleUndo.textAlign === summaryAfterTextStylePaste.textAlign &&
      summaryAfterStandaloneParagraphStyleUndo.paragraphList === summaryAfterTextStylePaste.paragraphList &&
      summaryAfterStandaloneParagraphStyleUndo.paragraphLineHeight === summaryAfterTextStylePaste.paragraphLineHeight &&
      summaryAfterStandaloneParagraphStyleUndo.paragraphSpacingAfter === summaryAfterTextStylePaste.paragraphSpacingAfter &&
      summaryAfterStandaloneParagraphStyleUndo.paragraphSpacingBefore === summaryAfterTextStylePaste.paragraphSpacingBefore,
    {
      summaryAfterStandaloneParagraphStylePaste,
      summaryAfterStandaloneParagraphStyleUndo,
      summaryAfterTextStylePaste,
    },
  )

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify(42)

    dataTransfer.setData(
      'application/vnd.interactive-os.ppt.text-font-size+json',
      json,
    )
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const summaryAfterTextFontSizePaste = await getPPTTextFormatPainterState(page, 's1-summary')

  record(
    'pastes PPT text font size JSON through slide-edit style clipboard effect',
    summaryAfterTextFontSizePaste.selected === 'true' &&
      summaryAfterTextFontSizePaste.text === summaryAfterTextStylePaste.text &&
      summaryAfterTextFontSizePaste.fontSize === '42px' &&
      summaryAfterTextFontSizePaste.textFontSizeImportModel === 'ppt-text-font-size-import' &&
      summaryAfterTextFontSizePaste.textFontSizeImportFormat === 'application-json-ppt-text-font-size' &&
      summaryAfterTextFontSizePaste.textFontSizeImportCommand === 'paste-object-formatting' &&
      summaryAfterTextFontSizePaste.textFontSizeImportCommandTargets === 's1-summary' &&
      summaryAfterTextFontSizePaste.textFontSizeImportCommandType === 'slide-command-effect' &&
      summaryAfterTextFontSizePaste.textFontSizeImportObjects === 's1-summary' &&
      summaryAfterTextFontSizePaste.textFontSizeImportCategories.includes('object-effect') &&
      summaryAfterTextFontSizePaste.textFontSizeImportCategories.includes('text-style') &&
      summaryAfterTextFontSizePaste.textFontSizeImportFields === 'value' &&
      summaryAfterTextFontSizePaste.textFontSizeImportJsonLength >= 2 &&
      summaryAfterTextFontSizePaste.textFontSizeImportValue === '42' &&
      summaryAfterTextFontSizePaste.styleClipboardCommand === 'paste-object-formatting' &&
      summaryAfterTextFontSizePaste.styleClipboardCommandTargets === 's1-summary' &&
      summaryAfterTextFontSizePaste.styleClipboardCommandApplications.includes('s1-summary') &&
      summaryAfterTextFontSizePaste.styleClipboardCommandApplications.includes('text-style'),
    {
      summaryAfterTextFontSizePaste,
      summaryAfterTextStylePaste,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const summaryAfterTextFontSizeUndo = await getPPTTextFormatPainterState(page, 's1-summary')

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const summaryAfterTextFontSizeRedo = await getPPTTextFormatPainterState(page, 's1-summary')

  record(
    'undoes and redoes PPT text font size JSON as one history step',
    summaryAfterTextFontSizeUndo.fontSize === '34px' &&
      summaryAfterTextFontSizeRedo.fontSize === '42px',
    {
      summaryAfterTextFontSizePaste,
      summaryAfterTextFontSizeRedo,
      summaryAfterTextFontSizeUndo,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const summaryAfterTextFontSizeRestore = await getPPTTextFormatPainterState(page, 's1-summary')

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify('bold')

    dataTransfer.setData(
      'application/vnd.interactive-os.ppt.text-font-weight+json',
      json,
    )
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const summaryAfterTextFontWeightPaste = await getPPTTextFormatPainterState(page, 's1-summary')

  record(
    'pastes PPT text font weight JSON through slide-edit style clipboard effect',
    summaryAfterTextFontSizeRestore.fontSize === '34px' &&
      summaryAfterTextFontWeightPaste.selected === 'true' &&
      summaryAfterTextFontWeightPaste.text === summaryAfterTextStylePaste.text &&
      summaryAfterTextFontWeightPaste.fontSize === '34px' &&
      summaryAfterTextFontWeightPaste.fontWeight === '700' &&
      summaryAfterTextFontWeightPaste.textFontWeightImportModel === 'ppt-text-font-weight-import' &&
      summaryAfterTextFontWeightPaste.textFontWeightImportFormat === 'application-json-ppt-text-font-weight' &&
      summaryAfterTextFontWeightPaste.textFontWeightImportCommand === 'paste-object-formatting' &&
      summaryAfterTextFontWeightPaste.textFontWeightImportCommandTargets === 's1-summary' &&
      summaryAfterTextFontWeightPaste.textFontWeightImportCommandType === 'slide-command-effect' &&
      summaryAfterTextFontWeightPaste.textFontWeightImportObjects === 's1-summary' &&
      summaryAfterTextFontWeightPaste.textFontWeightImportCategories.includes('object-effect') &&
      summaryAfterTextFontWeightPaste.textFontWeightImportCategories.includes('text-style') &&
      summaryAfterTextFontWeightPaste.textFontWeightImportFields === 'value' &&
      summaryAfterTextFontWeightPaste.textFontWeightImportJsonLength > 5 &&
      summaryAfterTextFontWeightPaste.textFontWeightImportValue === 'bold' &&
      summaryAfterTextFontWeightPaste.styleClipboardCommand === 'paste-object-formatting' &&
      summaryAfterTextFontWeightPaste.styleClipboardCommandTargets === 's1-summary' &&
      summaryAfterTextFontWeightPaste.styleClipboardCommandApplications.includes('s1-summary') &&
      summaryAfterTextFontWeightPaste.styleClipboardCommandApplications.includes('text-style'),
    {
      summaryAfterTextFontSizeRestore,
      summaryAfterTextFontWeightPaste,
      summaryAfterTextStylePaste,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const summaryAfterTextFontWeightUndo = await getPPTTextFormatPainterState(page, 's1-summary')

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const summaryAfterTextFontWeightRedo = await getPPTTextFormatPainterState(page, 's1-summary')

  record(
    'undoes and redoes PPT text font weight JSON as one history step',
    summaryAfterTextFontWeightUndo.fontWeight === '600' &&
      summaryAfterTextFontWeightRedo.fontWeight === '700',
    {
      summaryAfterTextFontWeightPaste,
      summaryAfterTextFontWeightRedo,
      summaryAfterTextFontWeightUndo,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const summaryAfterTextFontWeightRestore = await getPPTTextFormatPainterState(page, 's1-summary')

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify('right')

    dataTransfer.setData(
      'application/vnd.interactive-os.ppt.text-paragraph-align+json',
      json,
    )
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const summaryAfterTextParagraphAlignPaste = await getPPTTextFormatPainterState(page, 's1-summary')

  record(
    'pastes PPT text paragraph align JSON through slide-edit style clipboard effect',
    summaryAfterTextFontWeightRestore.fontWeight === '600' &&
      summaryAfterTextFontWeightRestore.textAlign === 'center' &&
      summaryAfterTextParagraphAlignPaste.selected === 'true' &&
      summaryAfterTextParagraphAlignPaste.text === summaryAfterTextStylePaste.text &&
      summaryAfterTextParagraphAlignPaste.textAlign === 'right' &&
      summaryAfterTextParagraphAlignPaste.textParagraphAlignImportModel === 'ppt-text-paragraph-align-import' &&
      summaryAfterTextParagraphAlignPaste.textParagraphAlignImportFormat === 'application-json-ppt-text-paragraph-align' &&
      summaryAfterTextParagraphAlignPaste.textParagraphAlignImportCommand === 'paste-object-formatting' &&
      summaryAfterTextParagraphAlignPaste.textParagraphAlignImportCommandTargets === 's1-summary' &&
      summaryAfterTextParagraphAlignPaste.textParagraphAlignImportCommandType === 'slide-command-effect' &&
      summaryAfterTextParagraphAlignPaste.textParagraphAlignImportObjects === 's1-summary' &&
      summaryAfterTextParagraphAlignPaste.textParagraphAlignImportCategories.includes('object-effect') &&
      summaryAfterTextParagraphAlignPaste.textParagraphAlignImportCategories.includes('text-style') &&
      summaryAfterTextParagraphAlignPaste.textParagraphAlignImportFields === 'value' &&
      summaryAfterTextParagraphAlignPaste.textParagraphAlignImportJsonLength > 6 &&
      summaryAfterTextParagraphAlignPaste.textParagraphAlignImportValue === 'right' &&
      summaryAfterTextParagraphAlignPaste.styleClipboardCommand === 'paste-object-formatting' &&
      summaryAfterTextParagraphAlignPaste.styleClipboardCommandTargets === 's1-summary' &&
      summaryAfterTextParagraphAlignPaste.styleClipboardCommandApplications.includes('s1-summary') &&
      summaryAfterTextParagraphAlignPaste.styleClipboardCommandApplications.includes('text-style'),
    {
      summaryAfterTextFontWeightRestore,
      summaryAfterTextParagraphAlignPaste,
      summaryAfterTextStylePaste,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const summaryAfterTextParagraphAlignUndo = await getPPTTextFormatPainterState(page, 's1-summary')

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const summaryAfterTextParagraphAlignRedo = await getPPTTextFormatPainterState(page, 's1-summary')

  record(
    'undoes and redoes PPT text paragraph align JSON as one history step',
    summaryAfterTextParagraphAlignUndo.textAlign === 'center' &&
      summaryAfterTextParagraphAlignRedo.textAlign === 'right',
    {
      summaryAfterTextParagraphAlignPaste,
      summaryAfterTextParagraphAlignRedo,
      summaryAfterTextParagraphAlignUndo,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const summaryAfterTextParagraphAlignRestore = await getPPTTextFormatPainterState(page, 's1-summary')

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify('bullet')

    dataTransfer.setData(
      'application/vnd.interactive-os.ppt.text-paragraph-bullet+json',
      json,
    )
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const summaryAfterTextParagraphBulletPaste = await getPPTTextFormatPainterState(page, 's1-summary')

  record(
    'pastes PPT text paragraph bullet JSON through slide-edit style clipboard effect',
    summaryAfterTextParagraphAlignRestore.textAlign === 'center' &&
      summaryAfterTextParagraphAlignRestore.paragraphList === 'numbered' &&
      summaryAfterTextParagraphBulletPaste.selected === 'true' &&
      summaryAfterTextParagraphBulletPaste.text === summaryAfterTextStylePaste.text &&
      summaryAfterTextParagraphBulletPaste.textAlign === 'center' &&
      summaryAfterTextParagraphBulletPaste.bulletList === 'true' &&
      summaryAfterTextParagraphBulletPaste.numberedList === '' &&
      summaryAfterTextParagraphBulletPaste.paragraphList === 'bullet' &&
      summaryAfterTextParagraphBulletPaste.paragraphBullet.length > 0 &&
      summaryAfterTextParagraphBulletPaste.textParagraphBulletImportModel === 'ppt-text-paragraph-bullet-import' &&
      summaryAfterTextParagraphBulletPaste.textParagraphBulletImportFormat === 'application-json-ppt-text-paragraph-bullet' &&
      summaryAfterTextParagraphBulletPaste.textParagraphBulletImportCommand === 'paste-object-formatting' &&
      summaryAfterTextParagraphBulletPaste.textParagraphBulletImportCommandTargets === 's1-summary' &&
      summaryAfterTextParagraphBulletPaste.textParagraphBulletImportCommandType === 'slide-command-effect' &&
      summaryAfterTextParagraphBulletPaste.textParagraphBulletImportObjects === 's1-summary' &&
      summaryAfterTextParagraphBulletPaste.textParagraphBulletImportCategories.includes('object-effect') &&
      summaryAfterTextParagraphBulletPaste.textParagraphBulletImportCategories.includes('text-style') &&
      summaryAfterTextParagraphBulletPaste.textParagraphBulletImportFields === 'value' &&
      summaryAfterTextParagraphBulletPaste.textParagraphBulletImportJsonLength > 7 &&
      summaryAfterTextParagraphBulletPaste.textParagraphBulletImportValue === 'bullet' &&
      summaryAfterTextParagraphBulletPaste.styleClipboardCommand === 'paste-object-formatting' &&
      summaryAfterTextParagraphBulletPaste.styleClipboardCommandTargets === 's1-summary' &&
      summaryAfterTextParagraphBulletPaste.styleClipboardCommandApplications.includes('s1-summary') &&
      summaryAfterTextParagraphBulletPaste.styleClipboardCommandApplications.includes('text-style'),
    {
      summaryAfterTextParagraphAlignRestore,
      summaryAfterTextParagraphBulletPaste,
      summaryAfterTextStylePaste,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const summaryAfterTextParagraphBulletUndo = await getPPTTextFormatPainterState(page, 's1-summary')

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const summaryAfterTextParagraphBulletRedo = await getPPTTextFormatPainterState(page, 's1-summary')

  record(
    'undoes and redoes PPT text paragraph bullet JSON as one history step',
    summaryAfterTextParagraphBulletUndo.paragraphList === 'numbered' &&
      summaryAfterTextParagraphBulletRedo.paragraphList === 'bullet',
    {
      summaryAfterTextParagraphBulletPaste,
      summaryAfterTextParagraphBulletRedo,
      summaryAfterTextParagraphBulletUndo,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const summaryAfterTextParagraphBulletRestore = await getPPTTextFormatPainterState(page, 's1-summary')

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify('28px')

    dataTransfer.setData(
      'application/vnd.interactive-os.ppt.text-run-size+json',
      json,
    )
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const summaryAfterTextRunSizePaste = await getPPTTextFormatPainterState(page, 's1-summary')

  record(
    'pastes PPT text run size JSON through run style command effect',
    summaryAfterTextParagraphBulletRestore.paragraphList === 'numbered' &&
      summaryAfterTextParagraphBulletRestore.runSizeCount === 0 &&
      summaryAfterTextRunSizePaste.selected === 'true' &&
      summaryAfterTextRunSizePaste.text === summaryAfterTextStylePaste.text &&
      summaryAfterTextRunSizePaste.runSize === '28px' &&
      summaryAfterTextRunSizePaste.runSizeCount > 0 &&
      summaryAfterTextRunSizePaste.runSizeValue === '28' &&
      summaryAfterTextRunSizePaste.textRunSizeImportModel === 'ppt-text-run-size-import' &&
      summaryAfterTextRunSizePaste.textRunSizeImportFormat === 'application-json-ppt-text-run-size' &&
      summaryAfterTextRunSizePaste.textRunSizeImportCommandIds === 'update-text-run-style' &&
      summaryAfterTextRunSizePaste.textRunSizeImportCommandFields === 'size' &&
      summaryAfterTextRunSizePaste.textRunSizeImportCommandTargets === 's1-summary' &&
      summaryAfterTextRunSizePaste.textRunSizeImportCommandTypes === 'slide-command-effect' &&
      summaryAfterTextRunSizePaste.textRunSizeImportCommandValues === '28' &&
      summaryAfterTextRunSizePaste.textRunSizeImportFields === 'value' &&
      summaryAfterTextRunSizePaste.textRunSizeImportJsonLength >= 6 &&
      summaryAfterTextRunSizePaste.textRunSizeImportObjects === 's1-summary' &&
      summaryAfterTextRunSizePaste.textRunSizeImportRuns > 0 &&
      summaryAfterTextRunSizePaste.textRunSizeImportSlide === 'slide-1' &&
      summaryAfterTextRunSizePaste.textRunSizeImportValue === '28',
    {
      summaryAfterTextParagraphBulletRestore,
      summaryAfterTextRunSizePaste,
      summaryAfterTextStylePaste,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const summaryAfterTextRunSizeUndo = await getPPTTextFormatPainterState(page, 's1-summary')

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const summaryAfterTextRunSizeRedo = await getPPTTextFormatPainterState(page, 's1-summary')

  record(
    'undoes and redoes PPT text run size JSON as one history step',
    summaryAfterTextRunSizeUndo.runSizeCount === 0 &&
      summaryAfterTextRunSizeRedo.runSize === '28px' &&
      summaryAfterTextRunSizeRedo.runSizeCount > 0,
    {
      summaryAfterTextRunSizePaste,
      summaryAfterTextRunSizeRedo,
      summaryAfterTextRunSizeUndo,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify('#0f766e')

    dataTransfer.setData(
      'application/vnd.interactive-os.ppt.text-run-color+json',
      json,
    )
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const summaryAfterTextRunColorPaste = await getPPTTextFormatPainterState(page, 's1-summary')

  record(
    'pastes PPT text run color JSON through run style command effect',
    summaryAfterTextParagraphBulletRestore.paragraphList === 'numbered' &&
      summaryAfterTextParagraphBulletRestore.runColorCount === 0 &&
      summaryAfterTextRunColorPaste.selected === 'true' &&
      summaryAfterTextRunColorPaste.text === summaryAfterTextStylePaste.text &&
      summaryAfterTextRunColorPaste.runColor === 'rgb(15, 118, 110)' &&
      summaryAfterTextRunColorPaste.runColorCount > 0 &&
      summaryAfterTextRunColorPaste.runColorValue === '#0f766e' &&
      summaryAfterTextRunColorPaste.textRunColorImportModel === 'ppt-text-run-color-import' &&
      summaryAfterTextRunColorPaste.textRunColorImportFormat === 'application-json-ppt-text-run-color' &&
      summaryAfterTextRunColorPaste.textRunColorImportCommandIds === 'update-text-run-style' &&
      summaryAfterTextRunColorPaste.textRunColorImportCommandFields === 'color' &&
      summaryAfterTextRunColorPaste.textRunColorImportCommandTargets === 's1-summary' &&
      summaryAfterTextRunColorPaste.textRunColorImportCommandTypes === 'slide-command-effect' &&
      summaryAfterTextRunColorPaste.textRunColorImportCommandValues === '#0f766e' &&
      summaryAfterTextRunColorPaste.textRunColorImportFields === 'value' &&
      summaryAfterTextRunColorPaste.textRunColorImportJsonLength >= 9 &&
      summaryAfterTextRunColorPaste.textRunColorImportObjects === 's1-summary' &&
      summaryAfterTextRunColorPaste.textRunColorImportRuns > 0 &&
      summaryAfterTextRunColorPaste.textRunColorImportSlide === 'slide-1' &&
      summaryAfterTextRunColorPaste.textRunColorImportValue === '#0f766e',
    {
      summaryAfterTextParagraphBulletRestore,
      summaryAfterTextRunColorPaste,
      summaryAfterTextStylePaste,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const summaryAfterTextRunColorUndo = await getPPTTextFormatPainterState(page, 's1-summary')

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const summaryAfterTextRunColorRedo = await getPPTTextFormatPainterState(page, 's1-summary')

  record(
    'undoes and redoes PPT text run color JSON as one history step',
    summaryAfterTextRunColorUndo.runColorCount === 0 &&
      summaryAfterTextRunColorRedo.runColor === 'rgb(15, 118, 110)' &&
      summaryAfterTextRunColorRedo.runColorCount > 0,
    {
      summaryAfterTextRunColorPaste,
      summaryAfterTextRunColorRedo,
      summaryAfterTextRunColorUndo,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify(true)

    dataTransfer.setData(
      'application/vnd.interactive-os.ppt.text-run-bold+json',
      json,
    )
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const summaryAfterTextRunBoldPaste = await getPPTTextFormatPainterState(page, 's1-summary')

  record(
    'pastes PPT text run bold JSON through run style command effect',
    summaryAfterTextParagraphBulletRestore.paragraphList === 'numbered' &&
      summaryAfterTextParagraphBulletRestore.boldRunCount === 0 &&
      summaryAfterTextRunBoldPaste.selected === 'true' &&
      summaryAfterTextRunBoldPaste.text === summaryAfterTextStylePaste.text &&
      ['700', 'bold'].includes(summaryAfterTextRunBoldPaste.boldRun) &&
      summaryAfterTextRunBoldPaste.boldRunCount > 0 &&
      summaryAfterTextRunBoldPaste.textRunBoldImportModel === 'ppt-text-run-bold-import' &&
      summaryAfterTextRunBoldPaste.textRunBoldImportFormat === 'application-json-ppt-text-run-bold' &&
      summaryAfterTextRunBoldPaste.textRunBoldImportCommandIds === 'update-text-run-style' &&
      summaryAfterTextRunBoldPaste.textRunBoldImportCommandFields === 'bold' &&
      summaryAfterTextRunBoldPaste.textRunBoldImportCommandTargets === 's1-summary' &&
      summaryAfterTextRunBoldPaste.textRunBoldImportCommandTypes === 'slide-command-effect' &&
      summaryAfterTextRunBoldPaste.textRunBoldImportCommandValues === 'true' &&
      summaryAfterTextRunBoldPaste.textRunBoldImportFields === 'value' &&
      summaryAfterTextRunBoldPaste.textRunBoldImportJsonLength >= 4 &&
      summaryAfterTextRunBoldPaste.textRunBoldImportObjects === 's1-summary' &&
      summaryAfterTextRunBoldPaste.textRunBoldImportRuns > 0 &&
      summaryAfterTextRunBoldPaste.textRunBoldImportSlide === 'slide-1' &&
      summaryAfterTextRunBoldPaste.textRunBoldImportValue === 'true',
    {
      summaryAfterTextParagraphBulletRestore,
      summaryAfterTextRunBoldPaste,
      summaryAfterTextStylePaste,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const summaryAfterTextRunBoldUndo = await getPPTTextFormatPainterState(page, 's1-summary')

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const summaryAfterTextRunBoldRedo = await getPPTTextFormatPainterState(page, 's1-summary')

  record(
    'undoes and redoes PPT text run bold JSON as one history step',
    summaryAfterTextRunBoldUndo.boldRunCount === 0 &&
      ['700', 'bold'].includes(summaryAfterTextRunBoldRedo.boldRun) &&
      summaryAfterTextRunBoldRedo.boldRunCount > 0,
    {
      summaryAfterTextRunBoldPaste,
      summaryAfterTextRunBoldRedo,
      summaryAfterTextRunBoldUndo,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify(true)

    dataTransfer.setData(
      'application/vnd.interactive-os.ppt.text-run-italic+json',
      json,
    )
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const summaryAfterTextRunItalicPaste = await getPPTTextFormatPainterState(page, 's1-summary')

  record(
    'pastes PPT text run italic JSON through run style command effect',
    summaryAfterTextParagraphBulletRestore.paragraphList === 'numbered' &&
      summaryAfterTextParagraphBulletRestore.italicRunCount === 0 &&
      summaryAfterTextRunItalicPaste.selected === 'true' &&
      summaryAfterTextRunItalicPaste.text === summaryAfterTextStylePaste.text &&
      summaryAfterTextRunItalicPaste.italicRun === 'italic' &&
      summaryAfterTextRunItalicPaste.italicRunCount > 0 &&
      summaryAfterTextRunItalicPaste.textRunItalicImportModel === 'ppt-text-run-italic-import' &&
      summaryAfterTextRunItalicPaste.textRunItalicImportFormat === 'application-json-ppt-text-run-italic' &&
      summaryAfterTextRunItalicPaste.textRunItalicImportCommandIds === 'update-text-run-style' &&
      summaryAfterTextRunItalicPaste.textRunItalicImportCommandFields === 'italic' &&
      summaryAfterTextRunItalicPaste.textRunItalicImportCommandTargets === 's1-summary' &&
      summaryAfterTextRunItalicPaste.textRunItalicImportCommandTypes === 'slide-command-effect' &&
      summaryAfterTextRunItalicPaste.textRunItalicImportCommandValues === 'true' &&
      summaryAfterTextRunItalicPaste.textRunItalicImportFields === 'value' &&
      summaryAfterTextRunItalicPaste.textRunItalicImportJsonLength >= 4 &&
      summaryAfterTextRunItalicPaste.textRunItalicImportObjects === 's1-summary' &&
      summaryAfterTextRunItalicPaste.textRunItalicImportRuns > 0 &&
      summaryAfterTextRunItalicPaste.textRunItalicImportSlide === 'slide-1' &&
      summaryAfterTextRunItalicPaste.textRunItalicImportValue === 'true',
    {
      summaryAfterTextParagraphBulletRestore,
      summaryAfterTextRunItalicPaste,
      summaryAfterTextStylePaste,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const summaryAfterTextRunItalicUndo = await getPPTTextFormatPainterState(page, 's1-summary')

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const summaryAfterTextRunItalicRedo = await getPPTTextFormatPainterState(page, 's1-summary')

  record(
    'undoes and redoes PPT text run italic JSON as one history step',
    summaryAfterTextRunItalicUndo.italicRunCount === 0 &&
      summaryAfterTextRunItalicRedo.italicRun === 'italic' &&
      summaryAfterTextRunItalicRedo.italicRunCount > 0,
    {
      summaryAfterTextRunItalicPaste,
      summaryAfterTextRunItalicRedo,
      summaryAfterTextRunItalicUndo,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const summaryAfterTextRunItalicRestore = await getPPTTextFormatPainterState(page, 's1-summary')

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify('underline')

    dataTransfer.setData(
      'application/vnd.interactive-os.ppt.text-run-underline+json',
      json,
    )
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const summaryAfterTextRunUnderlinePaste = await getPPTTextFormatPainterState(page, 's1-summary')

  record(
    'pastes PPT text run underline JSON through run style command effect',
    summaryAfterTextRunItalicRestore.italicRunCount === 0 &&
      summaryAfterTextRunItalicRestore.underlineRunCount === 0 &&
      summaryAfterTextRunUnderlinePaste.selected === 'true' &&
      summaryAfterTextRunUnderlinePaste.text === summaryAfterTextStylePaste.text &&
      summaryAfterTextRunUnderlinePaste.underlineRun.includes('underline') &&
      summaryAfterTextRunUnderlinePaste.underlineRunCount > 0 &&
      summaryAfterTextRunUnderlinePaste.textRunUnderlineImportModel === 'ppt-text-run-underline-import' &&
      summaryAfterTextRunUnderlinePaste.textRunUnderlineImportFormat === 'application-json-ppt-text-run-underline' &&
      summaryAfterTextRunUnderlinePaste.textRunUnderlineImportCommandIds === 'update-text-run-style' &&
      summaryAfterTextRunUnderlinePaste.textRunUnderlineImportCommandFields === 'underline' &&
      summaryAfterTextRunUnderlinePaste.textRunUnderlineImportCommandTargets === 's1-summary' &&
      summaryAfterTextRunUnderlinePaste.textRunUnderlineImportCommandTypes === 'slide-command-effect' &&
      summaryAfterTextRunUnderlinePaste.textRunUnderlineImportCommandValues === 'true' &&
      summaryAfterTextRunUnderlinePaste.textRunUnderlineImportFields === 'value' &&
      summaryAfterTextRunUnderlinePaste.textRunUnderlineImportJsonLength >= 10 &&
      summaryAfterTextRunUnderlinePaste.textRunUnderlineImportObjects === 's1-summary' &&
      summaryAfterTextRunUnderlinePaste.textRunUnderlineImportRuns > 0 &&
      summaryAfterTextRunUnderlinePaste.textRunUnderlineImportSlide === 'slide-1' &&
      summaryAfterTextRunUnderlinePaste.textRunUnderlineImportValue === 'true',
    {
      summaryAfterTextRunItalicRestore,
      summaryAfterTextRunUnderlinePaste,
      summaryAfterTextStylePaste,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const summaryAfterTextRunUnderlineUndo = await getPPTTextFormatPainterState(page, 's1-summary')

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const summaryAfterTextRunUnderlineRedo = await getPPTTextFormatPainterState(page, 's1-summary')

  record(
    'undoes and redoes PPT text run underline JSON as one history step',
    summaryAfterTextRunUnderlineUndo.underlineRunCount === 0 &&
      summaryAfterTextRunUnderlineRedo.underlineRun.includes('underline') &&
      summaryAfterTextRunUnderlineRedo.underlineRunCount > 0,
    {
      summaryAfterTextRunUnderlinePaste,
      summaryAfterTextRunUnderlineRedo,
      summaryAfterTextRunUnderlineUndo,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)
  await waitUntil(
    () => page.eval(`!!document.querySelector('[data-ppt-element="s1-title"]') && !!document.querySelector('[data-ppt-element="s1-summary"]')`),
    'Expected PPT title and summary to render before multi-text selection',
  )
  titlePoint = await getElementCenter(page, 's1-title')
  const summaryPoint = await getElementCenter(page, 's1-summary')
  await clickMouse(page, titlePoint.x, titlePoint.y, 1)
  await delay(50)
  await clickMouse(page, summaryPoint.x, summaryPoint.y, 1, 8)
  await delay(80)

  const beforeMulti = await page.eval(`(() => ({
    quickBarVisible: !!document.querySelector('[data-ppt-text-quick-bar]'),
    selectedCount: document.querySelectorAll('[data-selected="true"]').length,
  }))()`)

  await setTextQuickColor(page, '#008060')
  await page.eval(`document.querySelector('[data-ppt-text-quick="align-center"]')?.click()`)
  await page.eval(`document.querySelector('[data-ppt-text-quick="bullet"]')?.click()`)
  await page.eval(`document.querySelector('[data-ppt-text-quick="italic"]')?.click()`)
  await page.eval(`document.querySelector('[data-ppt-text-quick="underline"]')?.click()`)
  await delay(100)

  const afterMultiFormat = await page.eval(`(() => {
    const title = document.querySelector('[data-ppt-element="s1-title"]')
    const summary = document.querySelector('[data-ppt-element="s1-summary"]')

    return {
      summaryBulletList: summary?.getAttribute('data-ppt-bullet-list') ?? '',
      summaryColor: summary?.style.color ?? '',
      summaryItalicRun: summary?.querySelector('[data-ppt-run-italic="true"]')?.style.fontStyle ?? '',
      summaryTextAlign: summary?.style.textAlign ?? '',
      summaryUnderlineRun: summary?.querySelector('[data-ppt-run-underline="true"]')?.style.textDecoration ?? '',
      titleBulletList: title?.getAttribute('data-ppt-bullet-list') ?? '',
      titleColor: title?.style.color ?? '',
      titleItalicRun: title?.querySelector('[data-ppt-run-italic="true"]')?.style.fontStyle ?? '',
      titleTextAlign: title?.style.textAlign ?? '',
      titleUnderlineRun: title?.querySelector('[data-ppt-run-underline="true"]')?.style.textDecoration ?? '',
    }
  })()`)

  record('applies PPT text quick formatting to multi-selected text objects', beforeMulti.quickBarVisible && beforeMulti.selectedCount === 2 && afterMultiFormat.titleColor === 'rgb(0, 128, 96)' && afterMultiFormat.summaryColor === 'rgb(0, 128, 96)' && afterMultiFormat.titleTextAlign === 'center' && afterMultiFormat.summaryTextAlign === 'center' && afterMultiFormat.titleBulletList === 'true' && afterMultiFormat.summaryBulletList === 'true' && afterMultiFormat.titleItalicRun === 'italic' && afterMultiFormat.summaryItalicRun === 'italic' && afterMultiFormat.titleUnderlineRun.includes('underline') && afterMultiFormat.summaryUnderlineRun.includes('underline'), {
    afterMultiFormat,
    beforeMulti,
  })

  await page.eval(`document.querySelector('[data-ppt-text-quick="numbered"]')?.click()`)
  await delay(80)

  const afterNumberedToggle = await page.eval(`(() => {
    const title = document.querySelector('[data-ppt-element="s1-title"]')
    const summary = document.querySelector('[data-ppt-element="s1-summary"]')
    const exportCode = document.querySelector('.ppt-export-code')?.value ?? ''

    return {
      exportHasNumberedMarkup: exportCode.includes('data-ppt-numbered-list="true"') && exportCode.includes('data-ppt-numbered="true"'),
      exportHasNumberedModel: exportCode.includes('"bullet": "numbered"'),
      inspectorBulletPressed: document.querySelector('[data-ppt-paragraph-bullet]')?.getAttribute('aria-pressed') ?? '',
      inspectorNumberedPressed: document.querySelector('[data-ppt-paragraph-numbered]')?.getAttribute('aria-pressed') ?? '',
      numberedPressed: document.querySelector('[data-ppt-text-quick="numbered"]')?.getAttribute('aria-pressed') ?? '',
      summaryBulletList: summary?.getAttribute('data-ppt-bullet-list') ?? '',
      summaryNumberedCount: summary?.querySelectorAll('[data-ppt-numbered="true"]').length ?? 0,
      summaryNumberedList: summary?.getAttribute('data-ppt-numbered-list') ?? '',
      thumbNumberedCount: document.querySelectorAll('[data-ppt-thumb-numbered="true"]').length,
      titleBulletList: title?.getAttribute('data-ppt-bullet-list') ?? '',
      titleNumberedCount: title?.querySelectorAll('[data-ppt-numbered="true"]').length ?? 0,
      titleNumberedList: title?.getAttribute('data-ppt-numbered-list') ?? '',
    }
  })()`)

  record(
    'toggles PPT numbered list formatting on multi-selected text objects',
    afterNumberedToggle.numberedPressed === 'true' &&
      afterNumberedToggle.inspectorBulletPressed === 'false' &&
      afterNumberedToggle.inspectorNumberedPressed === 'true' &&
      afterNumberedToggle.titleBulletList === '' &&
      afterNumberedToggle.summaryBulletList === '' &&
      afterNumberedToggle.titleNumberedList === 'true' &&
      afterNumberedToggle.summaryNumberedList === 'true' &&
      afterNumberedToggle.exportHasNumberedMarkup &&
      afterNumberedToggle.exportHasNumberedModel &&
      afterNumberedToggle.titleNumberedCount > 0 &&
      afterNumberedToggle.summaryNumberedCount > 0 &&
      afterNumberedToggle.thumbNumberedCount > 0,
    afterNumberedToggle,
  )

  await page.eval(`document.querySelector('[data-ppt-text-quick="bullet"]')?.click()`)
  await delay(80)

  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)
  titlePoint = await getElementCenter(page, 's1-title')
  await clickMouse(page, titlePoint.x, titlePoint.y, 2)
  await delay(80)

  const afterEditingGuard = await page.eval(`(() => {
    const editor = document.querySelector('[data-ppt-element="s1-title"] .ppt-element-editor')

    return {
      editing: editor?.isContentEditable === true && document.activeElement === editor,
      quickBarVisible: !!document.querySelector('[data-ppt-text-quick-bar]'),
    }
  })()`)

  record('hides PPT text quick format bar while native text editing is active', afterEditingGuard.editing && !afterEditingGuard.quickBarVisible, afterEditingGuard)

  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)
}

async function runTextParagraphSpacingScenario(page) {
  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)
  await page.eval(`document.querySelector('.ppt-thumb[aria-label="Open Overview"]')?.click()`)
  await delay(80)

  const titlePoint = await getElementCenter(page, 's1-title')
  await clickMouse(page, titlePoint.x, titlePoint.y, 1)
  await delay(80)

  const initial = await getPPTTextParagraphSpacingState(page)

  record(
    'renders PPT paragraph spacing controls in text inspector',
    initial.inspector &&
      initial.descriptorSurface === 'text-paragraph-spacing' &&
      initial.lineHeightControl === 'line-height-ratio' &&
      initial.lineHeightCommand === 'update-text-paragraph-spacing' &&
      initial.model === 'slide-edit-text-paragraph-spacing' &&
      initial.selectedId === 's1-title' &&
      initial.lineHeight === '1.14' &&
      initial.spacingBefore === '0' &&
      initial.spacingBeforeControl === 'paragraph-spacing' &&
      initial.spacingBeforeUnit === 'px' &&
      initial.spacingAfter === '0' &&
      initial.spacingAfterControl === 'paragraph-spacing' &&
      initial.spacingAfterUnit === 'px' &&
      initial.selectedLineHeight === '1.14',
    initial,
  )

  await page.eval(`(() => {
    const lineHeight = document.querySelector('[data-ppt-paragraph-field="lineHeight"]')
    const spacingBefore = document.querySelector('[data-ppt-paragraph-field="spacingBefore"]')
    const spacingAfter = document.querySelector('[data-ppt-paragraph-field="spacingAfter"]')
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set

    setter.call(lineHeight, '1.4')
    lineHeight.dispatchEvent(new Event('input', { bubbles: true }))
    lineHeight.dispatchEvent(new Event('change', { bubbles: true }))

    setter.call(spacingBefore, '6')
    spacingBefore.dispatchEvent(new Event('input', { bubbles: true }))
    spacingBefore.dispatchEvent(new Event('change', { bubbles: true }))

    setter.call(spacingAfter, '12')
    spacingAfter.dispatchEvent(new Event('input', { bubbles: true }))
    spacingAfter.dispatchEvent(new Event('change', { bubbles: true }))
  })()`)
  await delay(120)

  const afterSpacing = await getPPTTextParagraphSpacingState(page)

  record(
    'updates PPT paragraph line height and spacing metadata from inspector',
    afterSpacing.command === 'update-text-paragraph-spacing' &&
      afterSpacing.commandField === 'paragraphAfter' &&
      afterSpacing.commandObject === 's1-title' &&
      afterSpacing.commandSlide === 'slide-1' &&
      afterSpacing.commandType === 'slide-command-effect' &&
      afterSpacing.commandUnit === 'px' &&
      afterSpacing.commandValue === '12' &&
      afterSpacing.inspectorLineHeight === '1.4' &&
      afterSpacing.inspectorSpacingBefore === '6' &&
      afterSpacing.inspectorSpacingAfter === '12' &&
      afterSpacing.lineHeight === '1.4' &&
      afterSpacing.spacingBefore === '6' &&
      afterSpacing.spacingAfter === '12' &&
      afterSpacing.selectedLineHeight === '1.4' &&
      afterSpacing.selectedSpacingBefore === '6' &&
      afterSpacing.selectedSpacingAfter === '12' &&
      afterSpacing.selectedStyleLineHeight === '1.4' &&
      afterSpacing.selectedStyleMarginTop === '6px' &&
      afterSpacing.selectedStyleMarginBottom === '12px',
    {
      afterSpacing,
      initial,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterUndo = await getPPTTextParagraphSpacingState(page)

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const afterRedo = await getPPTTextParagraphSpacingState(page)

  record(
    'undoes and redoes PPT paragraph spacing field as one history step',
    afterUndo.lineHeight === '1.4' &&
      afterUndo.spacingBefore === '6' &&
      afterUndo.spacingAfter === '0' &&
      afterRedo.lineHeight === '1.4' &&
      afterRedo.spacingBefore === '6' &&
      afterRedo.spacingAfter === '12',
    {
      afterRedo,
      afterSpacing,
      afterUndo,
    },
  )

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      textParagraphSpacing: {
        lineHeight: 1.28,
        spacingAfter: 15,
        spacingBefore: 9,
      },
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const afterJSONPaste = await getPPTTextParagraphSpacingState(page)

  record(
    'pastes PPT paragraph spacing JSON through slide-edit command effects',
    afterJSONPaste.command === 'update-text-paragraph-spacing' &&
      afterJSONPaste.commandField === 'paragraphAfter' &&
      afterJSONPaste.commandObject === 's1-title' &&
      afterJSONPaste.commandType === 'slide-command-effect' &&
      afterJSONPaste.commandUnit === 'px' &&
      afterJSONPaste.commandValue === '15' &&
      afterJSONPaste.importCommandFields === 'lineHeightRatio paragraphBefore paragraphAfter' &&
      afterJSONPaste.importCommandTargets === 's1-title s1-title s1-title' &&
      afterJSONPaste.importCommandTypes === 'slide-command-effect slide-command-effect slide-command-effect' &&
      afterJSONPaste.importCommandUnits === 'ratio px px' &&
      afterJSONPaste.importCommandValues === '1.28 9 15' &&
      afterJSONPaste.importCommands === 'update-text-paragraph-spacing update-text-paragraph-spacing update-text-paragraph-spacing' &&
      afterJSONPaste.importFields === 'lineHeight spacingBefore spacingAfter' &&
      afterJSONPaste.importFormat === 'application-json-ppt-text-paragraph-spacing' &&
      afterJSONPaste.importJsonLength > 60 &&
      afterJSONPaste.importLineHeight === '1.28' &&
      afterJSONPaste.importModel === 'ppt-text-paragraph-spacing-import' &&
      afterJSONPaste.importObjects === 's1-title' &&
      afterJSONPaste.importSpacingAfter === '15' &&
      afterJSONPaste.importSpacingBefore === '9' &&
      afterJSONPaste.lineHeight === '1.28' &&
      afterJSONPaste.spacingBefore === '9' &&
      afterJSONPaste.spacingAfter === '15' &&
      afterJSONPaste.selectedLineHeight === '1.28' &&
      afterJSONPaste.selectedSpacingBefore === '9' &&
      afterJSONPaste.selectedSpacingAfter === '15' &&
      afterJSONPaste.selectedStyleLineHeight === '1.28' &&
      afterJSONPaste.selectedStyleMarginTop === '9px' &&
      afterJSONPaste.selectedStyleMarginBottom === '15px',
    {
      afterJSONPaste,
      afterRedo,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterJSONUndo = await getPPTTextParagraphSpacingState(page)

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const afterJSONRedo = await getPPTTextParagraphSpacingState(page)

  record(
    'undoes and redoes PPT paragraph spacing JSON as one history step',
    afterJSONUndo.lineHeight === '1.4' &&
      afterJSONUndo.spacingBefore === '6' &&
      afterJSONUndo.spacingAfter === '12' &&
      afterJSONRedo.lineHeight === '1.28' &&
      afterJSONRedo.spacingBefore === '9' &&
      afterJSONRedo.spacingAfter === '15',
    {
      afterJSONPaste,
      afterJSONRedo,
      afterJSONUndo,
    },
  )
}

async function runTextFontFamilyScenario(page) {
  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)
  await page.eval(`document.querySelector('.ppt-thumb[aria-label="Open Overview"]')?.click()`)
  await delay(80)

  const titlePoint = await getElementCenter(page, 's1-title')
  await clickMouse(page, titlePoint.x, titlePoint.y, 1)
  await delay(80)

  const initial = await getPPTTextFontFamilyState(page)

  record(
    'renders PPT font family control in text inspector',
    initial.selectedId === 's1-title' &&
      initial.descriptorCommand === 'update-text-font-family' &&
      initial.descriptorControl === 'font-family-select' &&
      initial.descriptorFallback === 'Inter' &&
      initial.descriptorOptions === 'Inter Arial Georgia Courier New' &&
      initial.descriptorSurface === 'text-font-family' &&
      initial.fontFamily === 'Inter' &&
      initial.model === 'slide-edit-text-font-family' &&
      initial.selectedFontFamily === 'Inter' &&
      initial.thumbFontFamily === 'Inter',
    initial,
  )

  await page.eval(`(() => {
    const fontFamily = document.querySelector('[data-ppt-style-field="font-family"]')
    const setter = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set

    setter.call(fontFamily, 'Georgia')
    fontFamily.dispatchEvent(new Event('input', { bubbles: true }))
    fontFamily.dispatchEvent(new Event('change', { bubbles: true }))
  })()`)
  await delay(120)

  const afterGeorgia = await getPPTTextFontFamilyState(page)

  record(
    'updates PPT text font family metadata from inspector',
    afterGeorgia.command === 'update-text-font-family' &&
      afterGeorgia.commandField === 'fontFamily' &&
      afterGeorgia.commandObject === 's1-title' &&
      afterGeorgia.commandSlide === 'slide-1' &&
      afterGeorgia.commandType === 'slide-command-effect' &&
      afterGeorgia.commandValue === 'Georgia' &&
      afterGeorgia.fontFamily === 'Georgia' &&
      afterGeorgia.selectedFontFamily === 'Georgia' &&
      afterGeorgia.selectedStyleFontFamily.includes('Georgia') &&
      afterGeorgia.thumbFontFamily === 'Georgia' &&
      afterGeorgia.thumbStyleFontFamily.includes('Georgia'),
    {
      afterGeorgia,
      initial,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterUndo = await getPPTTextFontFamilyState(page)

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const afterRedo = await getPPTTextFontFamilyState(page)

  record(
    'undoes and redoes PPT text font family as one history step',
    afterUndo.fontFamily === 'Inter' &&
      afterUndo.selectedFontFamily === 'Inter' &&
      afterRedo.fontFamily === 'Georgia' &&
      afterRedo.selectedFontFamily === 'Georgia',
    {
      afterGeorgia,
      afterRedo,
      afterUndo,
    },
  )

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify('Courier New')

    dataTransfer.setData(
      'application/vnd.interactive-os.ppt.text-font-family+json',
      json,
    )
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const afterJSONPaste = await getPPTTextFontFamilyState(page)

  record(
    'pastes PPT text font family JSON through slide-edit command effect',
    afterJSONPaste.command === 'update-text-font-family' &&
      afterJSONPaste.commandField === 'fontFamily' &&
      afterJSONPaste.commandObject === 's1-title' &&
      afterJSONPaste.commandSlide === 'slide-1' &&
      afterJSONPaste.commandType === 'slide-command-effect' &&
      afterJSONPaste.commandValue === 'Courier New' &&
      afterJSONPaste.fontFamily === 'Courier New' &&
      afterJSONPaste.importCommandFields === 'fontFamily' &&
      afterJSONPaste.importCommandTargets === 's1-title' &&
      afterJSONPaste.importCommandTypes === 'slide-command-effect' &&
      afterJSONPaste.importCommandValues === 'Courier New' &&
      afterJSONPaste.importCommands === 'update-text-font-family' &&
      afterJSONPaste.importFields === 'value' &&
      afterJSONPaste.importFormat === 'application-json-ppt-text-font-family' &&
      afterJSONPaste.importJsonLength > 10 &&
      afterJSONPaste.importModel === 'ppt-text-font-family-import' &&
      afterJSONPaste.importObjects === 's1-title' &&
      afterJSONPaste.importSlide === 'slide-1' &&
      afterJSONPaste.importValue === 'Courier New' &&
      afterJSONPaste.selectedFontFamily === 'Courier New' &&
      afterJSONPaste.selectedStyleFontFamily.includes('Courier New') &&
      afterJSONPaste.thumbFontFamily === 'Courier New' &&
      afterJSONPaste.thumbStyleFontFamily.includes('Courier New'),
    {
      afterGeorgia,
      afterJSONPaste,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterJSONUndo = await getPPTTextFontFamilyState(page)

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const afterJSONRedo = await getPPTTextFontFamilyState(page)

  record(
    'undoes and redoes PPT text font family JSON as one history step',
    afterJSONUndo.fontFamily === 'Georgia' &&
      afterJSONUndo.selectedFontFamily === 'Georgia' &&
      afterJSONRedo.fontFamily === 'Courier New' &&
      afterJSONRedo.selectedFontFamily === 'Courier New',
    {
      afterJSONPaste,
      afterJSONRedo,
      afterJSONUndo,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  await page.eval(`document.querySelector('[data-ppt-present-start]')?.click()`)
  await delay(120)

  const preview = await page.eval(`(() => {
    const overlay = document.querySelector('[data-ppt-presentation]')
    const element = overlay?.querySelector('[data-ppt-element="s1-title"]')

    return {
      fontFamily: element?.getAttribute('data-ppt-font-family') ?? '',
      open: !!overlay,
      styleFontFamily: element?.style.fontFamily ?? '',
    }
  })()`)

  record(
    'keeps PPT font family metadata in presentation preview',
    preview.open &&
      preview.fontFamily === 'Georgia' &&
      preview.styleFontFamily.includes('Georgia'),
    preview,
  )

  await page.eval(`document.querySelector('[data-ppt-present-exit]')?.click()`)
  await delay(80)
}

async function runTextVerticalAlignScenario(page) {
  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)
  await page.eval(`document.querySelector('.ppt-thumb[aria-label="Open Overview"]')?.click()`)
  await delay(80)

  const titlePoint = await getElementCenter(page, 's1-title')
  await clickMouse(page, titlePoint.x, titlePoint.y, 1)
  await delay(80)

  const initial = await getPPTTextVerticalAlignState(page)

  record(
    'renders PPT text vertical alignment control in text inspector',
    initial.selectedId === 's1-title' &&
      initial.descriptorAttribute === 'data-slide-text-vertical-align' &&
      initial.descriptorAttributeValue === 'top' &&
      initial.descriptorCommand === 'update-text-vertical-alignment' &&
      initial.descriptorControl === 'vertical-alignment-segmented-control' &&
      initial.descriptorDefaultValue === 'top' &&
      initial.descriptorOptions === 'top middle bottom' &&
      initial.descriptorSurface === 'text-vertical-alignment' &&
      initial.model === 'slide-edit-text-vertical-alignment' &&
      initial.verticalAlign === 'top' &&
      initial.selectedVerticalAlign === 'top' &&
      initial.thumbVerticalAlign === 'top',
    initial,
  )

  await page.eval(`(() => {
    const verticalAlign = document.querySelector('[data-ppt-style-field="vertical-align"]')
    const setter = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set

    setter.call(verticalAlign, 'middle')
    verticalAlign.dispatchEvent(new Event('input', { bubbles: true }))
    verticalAlign.dispatchEvent(new Event('change', { bubbles: true }))
  })()`)
  await delay(120)

  const afterMiddle = await getPPTTextVerticalAlignState(page)

  record(
    'updates PPT text vertical alignment metadata from inspector',
    afterMiddle.command === 'update-text-vertical-alignment' &&
      afterMiddle.commandField === 'verticalAlignment' &&
      afterMiddle.commandObject === 's1-title' &&
      afterMiddle.commandSlide === 'slide-1' &&
      afterMiddle.commandType === 'slide-command-effect' &&
      afterMiddle.commandValue === 'middle' &&
      afterMiddle.descriptorAttributeValue === 'middle' &&
      afterMiddle.verticalAlign === 'middle' &&
      afterMiddle.selectedVerticalAlign === 'middle' &&
      afterMiddle.selectedStyleAlignItems === 'center' &&
      afterMiddle.thumbVerticalAlign === 'middle' &&
      afterMiddle.thumbStyleAlignItems === 'center',
    {
      afterMiddle,
      initial,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterUndo = await getPPTTextVerticalAlignState(page)

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const afterRedo = await getPPTTextVerticalAlignState(page)

  record(
    'undoes and redoes PPT text vertical alignment as one history step',
    afterUndo.verticalAlign === 'top' &&
      afterUndo.selectedVerticalAlign === 'top' &&
      afterRedo.verticalAlign === 'middle' &&
      afterRedo.selectedVerticalAlign === 'middle',
    {
      afterMiddle,
      afterRedo,
      afterUndo,
    },
  )

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      textVerticalAlign: 'bottom',
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const afterJSONPaste = await getPPTTextVerticalAlignState(page)

  record(
    'pastes PPT text vertical alignment JSON through slide-edit command effect',
    afterJSONPaste.command === 'update-text-vertical-alignment' &&
      afterJSONPaste.commandField === 'verticalAlignment' &&
      afterJSONPaste.commandObject === 's1-title' &&
      afterJSONPaste.commandType === 'slide-command-effect' &&
      afterJSONPaste.commandValue === 'bottom' &&
      afterJSONPaste.importAlignItems === 'flex-end' &&
      afterJSONPaste.importCommandFields === 'verticalAlignment' &&
      afterJSONPaste.importCommandTargets === 's1-title' &&
      afterJSONPaste.importCommandTypes === 'slide-command-effect' &&
      afterJSONPaste.importCommands === 'update-text-vertical-alignment' &&
      afterJSONPaste.importFields === 'textVerticalAlign' &&
      afterJSONPaste.importFormat === 'application-json-ppt-text-vertical-align' &&
      afterJSONPaste.importJsonLength > 25 &&
      afterJSONPaste.importModel === 'ppt-text-vertical-align-import' &&
      afterJSONPaste.importObjects === 's1-title' &&
      afterJSONPaste.importValue === 'bottom' &&
      afterJSONPaste.selectedStyleAlignItems === 'flex-end' &&
      afterJSONPaste.selectedVerticalAlign === 'bottom' &&
      afterJSONPaste.thumbStyleAlignItems === 'flex-end' &&
      afterJSONPaste.thumbVerticalAlign === 'bottom' &&
      afterJSONPaste.verticalAlign === 'bottom',
    {
      afterJSONPaste,
      afterRedo,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterJSONUndo = await getPPTTextVerticalAlignState(page)

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const afterJSONRedo = await getPPTTextVerticalAlignState(page)

  record(
    'undoes and redoes PPT text vertical alignment JSON as one history step',
    afterJSONUndo.selectedVerticalAlign === 'middle' &&
      afterJSONUndo.selectedStyleAlignItems === 'center' &&
      afterJSONRedo.selectedVerticalAlign === 'bottom' &&
      afterJSONRedo.selectedStyleAlignItems === 'flex-end',
    {
      afterJSONPaste,
      afterJSONRedo,
      afterJSONUndo,
    },
  )

  await page.eval(`document.querySelector('[data-ppt-present-start]')?.click()`)
  await delay(120)

  const preview = await page.eval(`(() => {
    const overlay = document.querySelector('[data-ppt-presentation]')
    const element = overlay?.querySelector('[data-ppt-element="s1-title"]')

    return {
      alignItems: element?.style.alignItems ?? '',
      open: !!overlay,
      verticalAlign: element?.getAttribute('data-ppt-vertical-align') ?? '',
    }
  })()`)

  record(
    'keeps PPT text vertical alignment metadata in presentation preview',
    preview.open &&
      preview.verticalAlign === 'bottom' &&
      preview.alignItems === 'flex-end',
    preview,
  )

  await page.eval(`document.querySelector('[data-ppt-present-exit]')?.click()`)
  await delay(80)
}

async function runTextFrameInsetScenario(page) {
  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)
  await page.eval(`document.querySelector('.ppt-thumb[aria-label="Open Overview"]')?.click()`)
  await delay(80)

  const titlePoint = await getElementCenter(page, 's1-title')
  await clickMouse(page, titlePoint.x, titlePoint.y, 1)
  await delay(80)

  const initial = await getPPTTextFrameInsetState(page)

  record(
    'renders PPT text frame inset controls in text inspector',
    initial.selectedId === 's1-title' &&
      initial.descriptorAttribute === 'data-slide-text-frame-inset' &&
      initial.descriptorAttributeValue === '0 0 0 0' &&
      initial.descriptorDefaultValue === '0 0 0 0' &&
      initial.descriptorSurface === 'text-frame-inset' &&
      initial.leftCommand === 'update-text-frame-inset' &&
      initial.leftControl === 'inset-number' &&
      initial.leftUnit === 'px' &&
      initial.model === 'slide-edit-text-frame-inset' &&
      initial.top === '0' &&
      initial.topControl === 'inset-number' &&
      initial.topUnit === 'px' &&
      initial.right === '0' &&
      initial.rightControl === 'inset-number' &&
      initial.rightUnit === 'px' &&
      initial.bottom === '0' &&
      initial.bottomControl === 'inset-number' &&
      initial.bottomUnit === 'px' &&
      initial.left === '0' &&
      initial.selectedTextInset === '0,0,0,0' &&
      initial.thumbTextInset === '0,0,0,0',
    initial,
  )

  await page.eval(`(() => {
    const values = {
      top: '10',
      right: '14',
      bottom: '18',
      left: '22',
    }
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set

    Object.entries(values).forEach(([field, value]) => {
      const input = document.querySelector(\`[data-ppt-text-inset-field="\${field}"]\`)
      setter.call(input, value)
      input.dispatchEvent(new Event('input', { bubbles: true }))
      input.dispatchEvent(new Event('change', { bubbles: true }))
    })
  })()`)
  await delay(120)

  const afterInset = await getPPTTextFrameInsetState(page)

  record(
    'updates PPT text frame inset metadata from inspector',
    afterInset.command === 'update-text-frame-inset' &&
      afterInset.commandField === 'left' &&
      afterInset.commandObject === 's1-title' &&
      afterInset.commandSlide === 'slide-1' &&
      afterInset.commandType === 'slide-command-effect' &&
      afterInset.commandValue === '22' &&
      afterInset.descriptorAttribute === 'data-slide-text-frame-inset' &&
      afterInset.descriptorAttributeValue === '10 14 18 22' &&
      afterInset.descriptorDefaultValue === '0 0 0 0' &&
      afterInset.descriptorSurface === 'text-frame-inset' &&
      afterInset.top === '10' &&
      afterInset.right === '14' &&
      afterInset.bottom === '18' &&
      afterInset.left === '22' &&
      afterInset.inspectorTextInset === '10,14,18,22' &&
      afterInset.selectedTextInset === '10,14,18,22' &&
      afterInset.selectedStylePadding === '10px 14px 18px 22px' &&
      afterInset.thumbTextInset === '10,14,18,22',
    {
      afterInset,
      initial,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterUndo = await getPPTTextFrameInsetState(page)

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const afterRedo = await getPPTTextFrameInsetState(page)

  record(
    'undoes and redoes PPT text frame inset field as one history step',
    afterUndo.top === '10' &&
      afterUndo.right === '14' &&
      afterUndo.bottom === '18' &&
      afterUndo.left === '0' &&
      afterRedo.selectedTextInset === '10,14,18,22',
    {
      afterInset,
      afterRedo,
      afterUndo,
    },
  )

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      textFrameInset: {
        bottom: 16,
        left: 20,
        right: 12,
        top: 8,
      },
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const afterJSONPaste = await getPPTTextFrameInsetState(page)

  record(
    'pastes PPT text frame inset JSON through slide-edit command effects',
    afterJSONPaste.importBottom === '16' &&
      afterJSONPaste.importCommandFields === 'top right bottom left' &&
      afterJSONPaste.importCommands === 'update-text-frame-inset update-text-frame-inset update-text-frame-inset update-text-frame-inset' &&
      afterJSONPaste.importCommandTargets === 's1-title s1-title s1-title s1-title' &&
      afterJSONPaste.importCommandTypes === 'slide-command-effect slide-command-effect slide-command-effect slide-command-effect' &&
      afterJSONPaste.importFields === 'top right bottom left' &&
      afterJSONPaste.importFormat === 'application-json-ppt-text-frame-inset' &&
      afterJSONPaste.importJsonLength > 50 &&
      afterJSONPaste.importLeft === '20' &&
      afterJSONPaste.importModel === 'ppt-text-frame-inset-import' &&
      afterJSONPaste.importObjects === 's1-title' &&
      afterJSONPaste.importRight === '12' &&
      afterJSONPaste.importTop === '8' &&
      afterJSONPaste.inspectorTextInset === '8,12,16,20' &&
      afterJSONPaste.selectedStylePadding === '8px 12px 16px 20px' &&
      afterJSONPaste.selectedTextInset === '8,12,16,20' &&
      afterJSONPaste.thumbTextInset === '8,12,16,20' &&
      afterJSONPaste.top === '8' &&
      afterJSONPaste.right === '12' &&
      afterJSONPaste.bottom === '16' &&
      afterJSONPaste.left === '20',
    {
      afterJSONPaste,
      afterRedo,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterJSONUndo = await getPPTTextFrameInsetState(page)

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const afterJSONRedo = await getPPTTextFrameInsetState(page)

  record(
    'undoes and redoes PPT text frame inset JSON as one history step',
    afterJSONUndo.selectedTextInset === '10,14,18,22' &&
      afterJSONUndo.selectedStylePadding === '10px 14px 18px 22px' &&
      afterJSONRedo.selectedTextInset === '8,12,16,20' &&
      afterJSONRedo.selectedStylePadding === '8px 12px 16px 20px',
    {
      afterJSONPaste,
      afterJSONRedo,
      afterJSONUndo,
    },
  )

  await page.eval(`document.querySelector('[data-ppt-present-start]')?.click()`)
  await delay(120)

  const preview = await page.eval(`(() => {
    const overlay = document.querySelector('[data-ppt-presentation]')
    const element = overlay?.querySelector('[data-ppt-element="s1-title"]')

    return {
      open: !!overlay,
      padding: element?.style.padding ?? '',
      textInset: element?.getAttribute('data-ppt-text-inset') ?? '',
    }
  })()`)

  record(
    'keeps PPT text frame inset metadata in presentation preview',
    preview.open &&
      preview.textInset === '8,12,16,20' &&
      preview.padding === '8px 12px 16px 20px',
    preview,
  )

  await page.eval(`document.querySelector('[data-ppt-present-exit]')?.click()`)
  await delay(80)
}

async function runExportScenario(page) {
  await installPPTDownloadCapture(page)

  const state = await page.eval(`(() => {
    const code = document.querySelector('.ppt-export-code')?.value ?? ''
    const objectAltText = ${JSON.stringify(PPT_OBJECT_ALT_TEXT)}

    return {
      hasDeckJson: code.includes('data-ppt-deck'),
      hasSlideMarkup: code.includes('data-ppt-slide="slide-1"'),
      hasElementMarkup: code.includes('data-ppt-element="s1-title"'),
      hasAnimationMarkup: code.includes('data-ppt-animation-type="flyIn"') && code.includes('data-ppt-animation-trigger="withPrevious"') && code.includes('data-ppt-animation-duration="800"') && code.includes('data-ppt-animation-delay="200"') && code.includes('data-ppt-animation-order="3"'),
      hasAnimationModel: code.includes('"animation"') && code.includes('"type": "flyIn"') && code.includes('"trigger": "withPrevious"') && code.includes('"durationMs": 800') && code.includes('"delayMs": 200') && code.includes('"order": 3'),
      hasBulletMarkup: code.includes('data-ppt-bullet-list="true"') && code.includes('data-ppt-bullet="true"'),
      hasBulletModel: code.includes('"bullet": "bullet"'),
      hasCommentMarkup: code.includes('class="ppt-element ppt-comment"') && code.includes('data-ppt-comment-resolved="true"') && code.includes('Review CTA wording'),
      hasCommentModel: code.includes('"kind": "comment"') && code.includes('"resolved": true') && code.includes('"body": "Review CTA wording"'),
      hasCommentThreadMarkup: code.includes('data-ppt-comment-thread-count="2"') && code.includes('data-ppt-comment-thread-body') && code.includes('Looks good after headline edit.'),
      hasCommentThreadModel: code.includes('"thread"') && code.includes('"body": "Looks good after headline edit."'),
      hasItalicMarkup: code.includes('data-ppt-run-italic="true"') && code.includes('font-style:italic'),
      hasItalicModel: code.includes('"italic": true'),
      hasObjectOpacityMarkup: code.includes('data-ppt-opacity="0.42"') && code.includes('opacity:0.42'),
      hasObjectOpacityModel: code.includes('"opacity": 0.42'),
      hasObjectShadowMarkup: code.includes('data-ppt-shadow="true"') && code.includes('data-ppt-shadow-color="#334155"') && code.includes('data-ppt-shadow-opacity="0.36"') && code.includes('filter:drop-shadow'),
      hasObjectShadowModel: code.includes('"shadow"') && code.includes('"color": "#334155"') && code.includes('"opacity": 0.36') && code.includes('"blur": 18') && code.includes('"distance": 12') && code.includes('"angle": 60'),
      hasObjectHyperlinkMarkup: code.includes('data-ppt-hyperlink-url="https://example.com/ppt"'),
      hasObjectHyperlinkModel: code.includes('"hyperlink"') && code.includes('"url": "https://example.com/ppt"'),
      hasObjectAltTextMarkup: code.includes('data-ppt-alt-text="' + objectAltText + '"') && code.includes('alt="' + objectAltText + '"'),
      hasObjectAltTextModel: code.includes('"accessibility"') && code.includes('"altText": "' + objectAltText + '"'),
      hasCornerRadiusMarkup: code.includes('data-ppt-corner-radius="36"') && code.includes('border-radius:36px'),
      hasCornerRadiusModel: code.includes('"cornerRadius": 36'),
      hasFillOpacityMarkup: code.includes('data-ppt-fill-opacity="0.35"') && code.includes('background:rgb(') && code.includes('/ 0.35'),
      hasFillOpacityModel: code.includes('"fill"') && code.includes('"opacity": 0.35'),
      hasStrokeDashMarkup: code.includes('data-ppt-stroke-dash="dash"') && code.includes('border-style:dashed') && code.includes('data-ppt-stroke-dash="dot"') && code.includes('stroke-dasharray='),
      hasStrokeDashModel: code.includes('"dash": "dash"') && code.includes('"dash": "dot"'),
      hasFontFamilyMarkup: code.includes('data-ppt-font-family="Georgia"') && code.includes('font-family:Georgia, serif'),
      hasFontFamilyModel: code.includes('"fontFamily": "Georgia"'),
      hasParagraphSpacingMarkup: code.includes('data-ppt-line-height="1.28"') && code.includes('data-ppt-spacing-before="9"') && code.includes('data-ppt-spacing-after="15"') && code.includes('line-height:1.28') && code.includes('margin-top:9px') && code.includes('margin-bottom:15px'),
      hasParagraphSpacingModel: code.includes('"lineHeight": 1.28') && code.includes('"spacingBefore": 9') && code.includes('"spacingAfter": 15'),
      hasTextFrameInsetMarkup: code.includes('data-ppt-text-inset="8,12,16,20"') && code.includes('padding:8px 12px 16px 20px'),
      hasTextFrameInsetModel: code.includes('"textInset"') && code.includes('"top": 8') && code.includes('"right": 12') && code.includes('"bottom": 16') && code.includes('"left": 20'),
      hasTextVerticalAlignMarkup: code.includes('data-ppt-vertical-align="bottom"') && code.includes('align-items:flex-end'),
      hasTextVerticalAlignModel: code.includes('"verticalAlign": "bottom"'),
      hasImageMarkup: code.includes('class="ppt-element ppt-image"') && code.includes('data:image/svg+xml'),
      hasImageFitMarkup: code.includes('data-ppt-image-fit="contain"') && code.includes('object-fit:contain'),
      hasImageFitModel: code.includes('"fit": "contain"'),
      hasImageCropMarkup: code.includes('data-ppt-image-crop-x="25"') && code.includes('data-ppt-image-crop-y="70"') && code.includes('object-position:25% 70%'),
      hasImageCropModel: code.includes('"crop"') && code.includes('"x": 25') && code.includes('"y": 70'),
      hasImageFlipMarkup: code.includes('data-ppt-flip-h="true"') && code.includes('scaleX(-1)'),
      hasImageFlipModel: code.includes('"flipH": true'),
      hasImageModel: code.includes('"kind": "image"') && code.includes('"src": "data:image/svg+xml'),
      hasFreeformMarkup: code.includes('class="ppt-element ppt-freeform"') && code.includes('data-ppt-kind="freeform"') && code.includes('data-ppt-freeform-path'),
      hasFreeformModel: code.includes('"kind": "freeform"') && code.includes('"points"') && code.includes('"stroke"'),
      hasLineConnectionMarkup: code.includes('data-ppt-start-connection="'),
      hasLineConnectionModel: code.includes('"startConnection"') && code.includes('"anchor"'),
      hasLineMarkup: code.includes('class="ppt-element ppt-line"') && code.includes('<line '),
      hasLineModel: code.includes('"kind": "line"') && code.includes('"endMarker": "arrow"'),
      hasLineRouteMarkup: code.includes('data-ppt-line-route="elbow"') && code.includes('<path '),
      hasLineRouteModel: code.includes('"route": "elbow"') && code.includes('"routeBend"'),
      hasLayoutMarkup: code.includes('data-ppt-layout-id="ppt-layout-split"'),
      hasLayoutModel: code.includes('"layoutId": "ppt-layout-split"'),
      hasPlaceholderVisibilityMarkup: code.includes('data-ppt-hidden-placeholders="media"'),
      hasPlaceholderVisibilityModel: code.includes('"hiddenPlaceholderIds"') && code.includes('"media"'),
      hasPPTDeckModel: code.includes('"slides"') && code.includes('"elements"'),
      hasRotationStyle: code.includes('transform:rotate(45deg)'),
      hasSpeakerNotesMarkup: code.includes('class="ppt-notes"') && code.includes('data-ppt-notes-for="slide-1"') && code.includes('Presenter cue: review image crop and final CTA.'),
      hasSpeakerNotesModel: code.includes('"notes": "Presenter cue: review image crop and final CTA."'),
      hasTableMarkup: code.includes('class="ppt-element ppt-table"') && code.includes('data-ppt-table-rows="') && code.includes('<th data-ppt-table-cell="0">'),
      hasTableModel: code.includes('"kind": "table"') && code.includes('"rows"') && code.includes('"Region"'),
      hasTextAutoFitMarkup: code.includes('data-ppt-text-autofit="resizeShapeToFitText"'),
      hasTextAutoFitModel: code.includes('"textAutoFit": "resizeShapeToFitText"'),
      hasThemeMarkup: code.includes('data-ppt-theme-id="ppt-theme-default"'),
      hasThemeModel: code.includes('"themeId": "ppt-theme-default"'),
      hasTransitionMarkup: code.includes('data-ppt-transition-type="push"') && code.includes('data-ppt-transition-duration="650"') && code.includes('data-ppt-transition-advance-on-click="false"') && code.includes('data-ppt-transition-advance-after="3000"'),
      hasTransitionModel: code.includes('"transition"') && code.includes('"type": "push"') && code.includes('"durationMs": 650') && code.includes('"advanceOnClick": false') && code.includes('"advanceAfterMs": 3000'),
      hasUnderlineMarkup: code.includes('data-ppt-run-underline="true"') && code.includes('text-decoration:underline'),
      hasUnderlineModel: code.includes('"underline": true'),
    }
  })()`)

  record('exports HTML slide markup', state.hasSlideMarkup, state)
  record('exports PPT element markup', state.hasElementMarkup, state)
  record('exports embedded PPT deck JSON', state.hasDeckJson && state.hasPPTDeckModel, state)
  record('exports PPT object animation metadata', state.hasAnimationMarkup && state.hasAnimationModel, state)
  record('exports PPT object opacity metadata', state.hasObjectOpacityMarkup && state.hasObjectOpacityModel, state)
  record('exports PPT object shadow metadata', state.hasObjectShadowMarkup && state.hasObjectShadowModel, state)
  record('exports PPT object hyperlink metadata', state.hasObjectHyperlinkMarkup && state.hasObjectHyperlinkModel, state)
  record('exports PPT object alt text metadata', state.hasObjectAltTextMarkup && state.hasObjectAltTextModel, state)
  record('exports PPT shape corner radius metadata', state.hasCornerRadiusMarkup && state.hasCornerRadiusModel, state)
  record('exports PPT shape fill opacity metadata', state.hasFillOpacityMarkup && state.hasFillOpacityModel, state)
  record('exports PPT stroke dash style metadata', state.hasStrokeDashMarkup && state.hasStrokeDashModel, state)
  record('exports PPT bullet list markup and model data', state.hasBulletMarkup && state.hasBulletModel, state)
  record('exports PPT font family markup and model data', state.hasFontFamilyMarkup && state.hasFontFamilyModel, state)
  record('exports PPT paragraph spacing markup and model data', state.hasParagraphSpacingMarkup && state.hasParagraphSpacingModel, state)
  record('exports PPT text frame inset markup and model data', state.hasTextFrameInsetMarkup && state.hasTextFrameInsetModel, state)
  record('exports PPT text vertical alignment markup and model data', state.hasTextVerticalAlignMarkup && state.hasTextVerticalAlignModel, state)
  record('exports PPT comment markup and model data', state.hasCommentMarkup && state.hasCommentModel, state)
  record('exports PPT comment thread markup and model data', state.hasCommentThreadMarkup && state.hasCommentThreadModel, state)
  record('exports PPT italic and underline run markup and model data', state.hasItalicMarkup && state.hasItalicModel && state.hasUnderlineMarkup && state.hasUnderlineModel, state)
  record('exports inserted PPT image markup and model data', state.hasImageMarkup && state.hasImageModel, state)
  record('exports PPT image fit markup and model data', state.hasImageFitMarkup && state.hasImageFitModel, state)
  record('exports PPT image crop position markup and model data', state.hasImageCropMarkup && state.hasImageCropModel, state)
  record('exports PPT image flip markup and model data', state.hasImageFlipMarkup && state.hasImageFlipModel, state)
  record('exports PPT freeform path markup and model data', state.hasFreeformMarkup && state.hasFreeformModel, state)
  record('exports inserted PPT line and arrow model data', state.hasLineMarkup && state.hasLineModel, state)
  record('exports PPT connector attachment metadata', state.hasLineConnectionMarkup && state.hasLineConnectionModel, state)
  record('exports PPT connector route metadata', state.hasLineRouteMarkup && state.hasLineRouteModel, state)
  record('exports PPT layout/theme metadata', state.hasLayoutMarkup && state.hasLayoutModel && state.hasThemeMarkup && state.hasThemeModel, state)
  record('exports PPT placeholder visibility metadata', state.hasPlaceholderVisibilityMarkup && state.hasPlaceholderVisibilityModel, state)
  record('exports PPT slide transition metadata', state.hasTransitionMarkup && state.hasTransitionModel, state)
  record('exports inserted PPT table markup and model data', state.hasTableMarkup && state.hasTableModel, state)
  record('exports PPT text auto-fit markup and model data', state.hasTextAutoFitMarkup && state.hasTextAutoFitModel, state)
  record('exports PPT speaker notes markup and model data', state.hasSpeakerNotesMarkup && state.hasSpeakerNotesModel, state)
  record('exports PPT object rotation style', state.hasRotationStyle, state)

  await page.eval(`(() => {
    window.__pptClipboardItemTypes = []
    window.__pptClipboardWriteCount = 0
    window.__pptClipboardWriteText = ''

    window.ClipboardItem = class PPTClipboardItem {
      constructor(items) {
        this.items = items
        window.__pptClipboardItemTypes.push(Object.keys(items).sort())
      }
    }

    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: {
        write: async (items) => {
          window.__pptClipboardWriteCount = items.length
        },
        writeText: async (text) => {
          window.__pptClipboardWriteText = text
        },
      },
    })
  })()`)
  await page.eval(`document.querySelector('[aria-label="Copy HTML"]')?.click()`)
  await delay(80)

  const htmlClipboardState = await page.eval(`(() => {
    const stage = document.querySelector('[data-ppt-app] .ppt-stage-shell')
    const mimeType = stage?.getAttribute('data-ppt-html-clipboard-json-mime-type') ?? ''
    const itemTypes = window.__pptClipboardItemTypes?.at(-1) ?? []

    return {
      hasCustomJson: itemTypes.includes(mimeType),
      hasHTML: itemTypes.includes('text/html'),
      hasPlainText: itemTypes.includes('text/plain'),
      htmlLength: Number(stage?.getAttribute('data-ppt-html-clipboard-html-length') ?? 0),
      itemTypes,
      mimeType,
      model: stage?.getAttribute('data-ppt-html-clipboard-model') ?? '',
      sourceSlide: stage?.getAttribute('data-ppt-html-clipboard-source-slide') ?? '',
      writeCount: window.__pptClipboardWriteCount ?? 0,
      writeMode: stage?.getAttribute('data-ppt-html-clipboard-write-mode') ?? '',
      writeTextLength: (window.__pptClipboardWriteText ?? '').length,
    }
  })()`)

  record(
    'copies PPT HTML through canvas rich clipboard writer',
    htmlClipboardState.model === 'canvas-rich-html-clipboard' &&
      htmlClipboardState.sourceSlide === 'slide-1' &&
      htmlClipboardState.writeMode === 'clipboard-item' &&
      htmlClipboardState.writeCount === 1 &&
      htmlClipboardState.htmlLength > 1000 &&
      htmlClipboardState.hasCustomJson &&
      htmlClipboardState.hasHTML &&
      htmlClipboardState.hasPlainText &&
      htmlClipboardState.writeTextLength === 0,
    htmlClipboardState,
  )

  const beforeDeckHTMLPaste = await page.eval(`(() => ({
    slideCount: document.querySelectorAll('.ppt-thumb').length,
  }))()`)

  await page.eval(`(() => {
    const html = document.querySelector('.ppt-export-code')?.value ?? ''
    const dataTransfer = new DataTransfer()

    dataTransfer.setData('text/html', html)
    dataTransfer.setData('text/plain', html)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const deckHTMLImportState = await page.eval(`(() => {
    const stage = document.querySelector('[data-ppt-app] .ppt-stage-shell')
    const activeSlide = document.querySelector('.ppt-slide')?.getAttribute('data-ppt-slide') ?? ''
    const activeThumb = document.querySelector('.ppt-thumb[aria-current="page"]')

    return {
      activeName: activeThumb?.querySelector('.ppt-thumb-name')?.textContent ?? '',
      activeSlide,
      firstImportedSlideId: stage?.getAttribute('data-ppt-deck-html-import-first-slide') ?? '',
      format: stage?.getAttribute('data-ppt-deck-html-import-format') ?? '',
      htmlLength: Number(stage?.getAttribute('data-ppt-deck-html-import-html-length') ?? 0),
      importedCount: Number(stage?.getAttribute('data-ppt-deck-html-import-imported-count') ?? 0),
      model: stage?.getAttribute('data-ppt-deck-html-import-model') ?? '',
      slideCount: document.querySelectorAll('.ppt-thumb').length,
      sourceDeck: stage?.getAttribute('data-ppt-deck-html-import-source-deck') ?? '',
      sourceSlideCount: Number(stage?.getAttribute('data-ppt-deck-html-import-source-slide-count') ?? 0),
      sourceTitle: stage?.getAttribute('data-ppt-deck-html-import-source-title') ?? '',
    }
  })()`)

  record(
    'pastes PPT HTML export back as imported deck slides',
    deckHTMLImportState.model === 'ppt-deck-html-import' &&
      deckHTMLImportState.format === 'text-html-ppt-deck' &&
      deckHTMLImportState.sourceDeck === 'deck-ai-retouch' &&
      deckHTMLImportState.sourceTitle === 'AI Retouch Demo' &&
      deckHTMLImportState.sourceSlideCount === beforeDeckHTMLPaste.slideCount &&
      deckHTMLImportState.importedCount === beforeDeckHTMLPaste.slideCount &&
      deckHTMLImportState.slideCount === beforeDeckHTMLPaste.slideCount * 2 &&
      deckHTMLImportState.activeSlide === deckHTMLImportState.firstImportedSlideId &&
      deckHTMLImportState.activeName.includes('Copy') &&
      deckHTMLImportState.htmlLength > 1000,
    {
      beforeDeckHTMLPaste,
      deckHTMLImportState,
    },
  )

  await deletePPTSlidesByThumbNameIncludes(page, ['Copy'])
  await page.eval(`document.querySelector('.ppt-thumb[aria-label="Open Overview"]')?.click()`)
  await delay(80)

  const afterDeckHTMLPasteCleanup = await page.eval(`(() => ({
    activeSlide: document.querySelector('.ppt-slide')?.getAttribute('data-ppt-slide') ?? '',
    slideCount: document.querySelectorAll('.ppt-thumb').length,
  }))()`)

  record(
    'removes PPT HTML deck paste probes before export scenario continues',
    afterDeckHTMLPasteCleanup.activeSlide === 'slide-1' &&
      afterDeckHTMLPasteCleanup.slideCount === beforeDeckHTMLPaste.slideCount,
    {
      afterDeckHTMLPasteCleanup,
      beforeDeckHTMLPaste,
    },
  )

  await page.eval(`(() => {
    const deck = {
      id: 'deck-ai-json-draft',
      size: { h: 720, w: 1280 },
      slides: [
        {
          background: { color: '#ffffff' },
          elements: [
            {
              geometry: { h: 82, w: 1040, x: 92, y: 72 },
              id: 'ai-json-title',
              kind: 'textBox',
              name: 'Title',
              style: { color: '#111827', fontSize: 42, fontWeight: 'bold' },
              textBody: { paragraphs: [{ runs: [{ text: 'AI JSON Draft' }] }] },
              textAutoFit: 'resizeShapeToFitText',
            },
            {
              geometry: { h: 210, w: 940, x: 104, y: 184 },
              id: 'ai-json-body',
              kind: 'textBox',
              name: 'Body',
              style: { color: '#1f2937', fontSize: 28 },
              textBody: {
                paragraphs: [
                  { bullet: 'bullet', runs: [{ text: 'Editable generated slide' }] },
                  { bullet: 'numbered', runs: [{ text: 'Retouch the final wording' }] },
                ],
              },
            },
          ],
          id: 'ai-json-slide-1',
          layoutId: 'ppt-layout-title-body',
          name: 'AI JSON Opening',
          notes: 'AI generated speaker note',
          themeId: 'ppt-theme-default',
        },
        {
          background: { color: '#f8fafc' },
          elements: [
            {
              fill: { color: '#eff6ff' },
              geometry: { h: 136, w: 520, x: 120, y: 124 },
              hyperlink: { url: 'https://example.com/json-source' },
              id: 'ai-json-link-card',
              kind: 'shape',
              name: 'JSON Link',
              shape: 'rect',
              stroke: { color: '#2563eb', width: 2 },
              style: { color: '#1e3a8a', fontSize: 26, fontWeight: 'semibold' },
              textBody: { paragraphs: [{ runs: [{ text: 'JSON source link' }] }] },
            },
          ],
          id: 'ai-json-slide-2',
          name: 'AI JSON Source',
          themeId: 'ppt-theme-default',
        },
      ],
      title: 'AI JSON Deck',
    }
    const json = JSON.stringify(deck)
    const dataTransfer = new DataTransfer()

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(140)

  const deckJSONImportState = await page.eval(`(() => {
    const stage = document.querySelector('[data-ppt-app] .ppt-stage-shell')
    const activeSlide = document.querySelector('.ppt-slide')?.getAttribute('data-ppt-slide') ?? ''
    const activeThumb = document.querySelector('.ppt-thumb[aria-current="page"]')
    const body = document.querySelector('.ppt-slide [data-ppt-element-name="Body"]')
    const title = document.querySelector('.ppt-slide [data-ppt-element-name="Title"]')
    const exportCode = document.querySelector('.ppt-export-code')?.value ?? ''

    return {
      activeName: activeThumb?.querySelector('.ppt-thumb-name')?.textContent ?? '',
      activeSlide,
      bodyText: body?.textContent ?? '',
      exportHasHyperlink: exportCode.includes('https://example.com/json-source') && exportCode.includes('data-ppt-hyperlink-url='),
      exportHasNotes: exportCode.includes('AI generated speaker note'),
      firstImportedSlideId: stage?.getAttribute('data-ppt-deck-json-import-first-slide') ?? '',
      format: stage?.getAttribute('data-ppt-deck-json-import-format') ?? '',
      importedCount: Number(stage?.getAttribute('data-ppt-deck-json-import-imported-count') ?? 0),
      jsonLength: Number(stage?.getAttribute('data-ppt-deck-json-import-json-length') ?? 0),
      model: stage?.getAttribute('data-ppt-deck-json-import-model') ?? '',
      slideCount: document.querySelectorAll('.ppt-thumb').length,
      sourceDeck: stage?.getAttribute('data-ppt-deck-json-import-source-deck') ?? '',
      sourceSlideCount: Number(stage?.getAttribute('data-ppt-deck-json-import-source-slide-count') ?? 0),
      sourceTitle: stage?.getAttribute('data-ppt-deck-json-import-source-title') ?? '',
      titleText: title?.textContent ?? '',
    }
  })()`)

  record(
    'pastes raw PPT deck JSON as editable slides',
    deckJSONImportState.model === 'ppt-deck-json-import' &&
      deckJSONImportState.format === 'application-json-ppt-deck' &&
      deckJSONImportState.sourceDeck === 'deck-ai-json-draft' &&
      deckJSONImportState.sourceTitle === 'AI JSON Deck' &&
      deckJSONImportState.sourceSlideCount === 2 &&
      deckJSONImportState.importedCount === 2 &&
      deckJSONImportState.slideCount === beforeDeckHTMLPaste.slideCount + 2 &&
      deckJSONImportState.activeSlide === deckJSONImportState.firstImportedSlideId &&
      deckJSONImportState.activeName.includes('AI JSON Opening Copy') &&
      deckJSONImportState.titleText.includes('AI JSON Draft') &&
      deckJSONImportState.bodyText.includes('Editable generated slide') &&
      deckJSONImportState.bodyText.includes('Retouch the final wording') &&
      deckJSONImportState.exportHasNotes &&
      deckJSONImportState.exportHasHyperlink &&
      deckJSONImportState.jsonLength > 100,
    {
      beforeDeckHTMLPaste,
      deckJSONImportState,
    },
  )

  await deletePPTSlidesByThumbNameIncludes(page, [
    'AI JSON Source Copy',
    'AI JSON Opening Copy',
  ])
  await page.eval(`document.querySelector('.ppt-thumb[aria-label="Open Overview"]')?.click()`)
  await delay(80)

  const afterDeckJSONPasteCleanup = await page.eval(`(() => ({
    activeSlide: document.querySelector('.ppt-slide')?.getAttribute('data-ppt-slide') ?? '',
    slideCount: document.querySelectorAll('.ppt-thumb').length,
  }))()`)

  record(
    'removes PPT deck JSON paste probes before export scenario continues',
    afterDeckJSONPasteCleanup.activeSlide === 'slide-1' &&
      afterDeckJSONPasteCleanup.slideCount === beforeDeckHTMLPaste.slideCount,
    {
      afterDeckJSONPasteCleanup,
      beforeDeckHTMLPaste,
    },
  )

  await page.eval(`(() => {
    const slide = {
      background: { color: '#ffffff' },
      elements: [
        {
          geometry: { h: 82, w: 1040, x: 92, y: 72 },
          id: 'ai-single-title',
          kind: 'textBox',
          name: 'Title',
          style: { color: '#111827', fontSize: 42, fontWeight: 'bold' },
          textBody: { paragraphs: [{ runs: [{ text: 'AI Single Slide' }] }] },
          textAutoFit: 'resizeShapeToFitText',
        },
        {
          geometry: { h: 190, w: 860, x: 104, y: 184 },
          id: 'ai-single-body',
          kind: 'textBox',
          name: 'Body',
          style: { color: '#1f2937', fontSize: 28 },
          textBody: {
            paragraphs: [
              { bullet: 'bullet', runs: [{ text: 'Paste one AI generated slide' }] },
              { bullet: 'numbered', runs: [{ text: 'Retouch as editable PPT content' }] },
            ],
          },
        },
        {
          fill: { color: '#eff6ff' },
          geometry: { h: 72, w: 340, x: 104, y: 426 },
          hyperlink: { url: 'https://example.com/single-slide-json' },
          id: 'ai-single-link',
          kind: 'shape',
          name: 'Source Link',
          shape: 'rect',
          stroke: { color: '#2563eb', width: 2 },
          style: { color: '#1e3a8a', fontSize: 22, fontWeight: 'semibold' },
          textBody: { paragraphs: [{ runs: [{ text: 'Source link' }] }] },
        },
      ],
      id: 'ai-single-slide',
      layoutId: 'ppt-layout-title-body',
      name: 'AI Single Slide',
      notes: 'AI single slide note',
      themeId: 'ppt-theme-default',
    }
    const json = JSON.stringify(slide)
    const dataTransfer = new DataTransfer()

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(140)

  const slideJSONImportState = await page.eval(`(() => {
    const stage = document.querySelector('[data-ppt-app] .ppt-stage-shell')
    const activeSlide = document.querySelector('.ppt-slide')?.getAttribute('data-ppt-slide') ?? ''
    const activeThumb = document.querySelector('.ppt-thumb[aria-current="page"]')
    const body = document.querySelector('.ppt-slide [data-ppt-element-name="Body"]')
    const title = document.querySelector('.ppt-slide [data-ppt-element-name="Title"]')
    const exportCode = document.querySelector('.ppt-export-code')?.value ?? ''

    return {
      activeName: activeThumb?.querySelector('.ppt-thumb-name')?.textContent ?? '',
      activeSlide,
      bodyText: body?.textContent ?? '',
      exportHasHyperlink: exportCode.includes('https://example.com/single-slide-json') && exportCode.includes('data-ppt-hyperlink-url='),
      exportHasNotes: exportCode.includes('AI single slide note'),
      firstImportedSlideId: stage?.getAttribute('data-ppt-slide-json-import-first-slide') ?? '',
      format: stage?.getAttribute('data-ppt-slide-json-import-format') ?? '',
      importedCount: Number(stage?.getAttribute('data-ppt-slide-json-import-imported-count') ?? 0),
      jsonLength: Number(stage?.getAttribute('data-ppt-slide-json-import-json-length') ?? 0),
      model: stage?.getAttribute('data-ppt-slide-json-import-model') ?? '',
      slideCount: document.querySelectorAll('.ppt-thumb').length,
      sourceSlideCount: Number(stage?.getAttribute('data-ppt-slide-json-import-source-slide-count') ?? 0),
      sourceSlideIds: stage?.getAttribute('data-ppt-slide-json-import-source-slide-ids') ?? '',
      sourceSlideNames: stage?.getAttribute('data-ppt-slide-json-import-source-slide-names') ?? '',
      titleText: title?.textContent ?? '',
    }
  })()`)

  record(
    'pastes raw PPT slide JSON as editable slide',
    slideJSONImportState.model === 'ppt-slide-json-import' &&
      slideJSONImportState.format === 'application-json-ppt-slide' &&
      slideJSONImportState.sourceSlideCount === 1 &&
      slideJSONImportState.sourceSlideIds === 'ai-single-slide' &&
      slideJSONImportState.sourceSlideNames === 'AI Single Slide' &&
      slideJSONImportState.importedCount === 1 &&
      slideJSONImportState.slideCount === beforeDeckHTMLPaste.slideCount + 1 &&
      slideJSONImportState.activeSlide === slideJSONImportState.firstImportedSlideId &&
      slideJSONImportState.activeName.includes('AI Single Slide Copy') &&
      slideJSONImportState.titleText.includes('AI Single Slide') &&
      slideJSONImportState.bodyText.includes('Paste one AI generated slide') &&
      slideJSONImportState.bodyText.includes('Retouch as editable PPT content') &&
      slideJSONImportState.exportHasNotes &&
      slideJSONImportState.exportHasHyperlink &&
      slideJSONImportState.jsonLength > 100,
    {
      beforeDeckHTMLPaste,
      slideJSONImportState,
    },
  )

  await deletePPTSlideByThumbName(page, 'AI Single Slide Copy')
  await page.eval(`document.querySelector('.ppt-thumb[aria-label="Open Overview"]')?.click()`)
  await delay(80)

  const afterSlideJSONPasteCleanup = await page.eval(`(() => ({
    activeSlide: document.querySelector('.ppt-slide')?.getAttribute('data-ppt-slide') ?? '',
    slideCount: document.querySelectorAll('.ppt-thumb').length,
  }))()`)

  record(
    'removes PPT slide JSON paste probe before export scenario continues',
    afterSlideJSONPasteCleanup.activeSlide === 'slide-1' &&
      afterSlideJSONPasteCleanup.slideCount === beforeDeckHTMLPaste.slideCount,
    {
      afterSlideJSONPasteCleanup,
      beforeDeckHTMLPaste,
    },
  )

  await page.eval(`(() => {
    const payload = {
      slides: [
        {
          background: { color: '#ffffff' },
          elements: [
            {
              geometry: { h: 86, w: 960, x: 108, y: 86 },
              id: 'ai-wrapper-title-a',
              kind: 'textBox',
              name: 'Title',
              style: { color: '#111827', fontSize: 42, fontWeight: 'bold' },
              textBody: { paragraphs: [{ runs: [{ text: 'Wrapper Slide A' }] }] },
            },
          ],
          id: 'ai-wrapper-slide-a',
          name: 'Wrapper Slide A',
          themeId: 'ppt-theme-default',
        },
        {
          background: { color: '#f8fafc' },
          elements: [
            {
              geometry: { h: 86, w: 960, x: 108, y: 86 },
              id: 'ai-wrapper-title-b',
              kind: 'textBox',
              name: 'Title',
              style: { color: '#111827', fontSize: 42, fontWeight: 'bold' },
              textBody: { paragraphs: [{ runs: [{ text: 'Wrapper Slide B' }] }] },
            },
          ],
          id: 'ai-wrapper-slide-b',
          name: 'Wrapper Slide B',
          themeId: 'ppt-theme-default',
        },
      ],
    }
    const fence = String.fromCharCode(96, 96, 96)
    const markdown = fence + 'json\\n' + JSON.stringify(payload, null, 2) +
      '\\n' + fence
    const dataTransfer = new DataTransfer()

    dataTransfer.setData('text/markdown', markdown)
    dataTransfer.setData('text/plain', markdown)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(140)

  const fencedSlideJSONImportState = await page.eval(`(() => {
    const stage = document.querySelector('[data-ppt-app] .ppt-stage-shell')
    const activeSlide = document.querySelector('.ppt-slide')?.getAttribute('data-ppt-slide') ?? ''
    const activeThumb = document.querySelector('.ppt-thumb[aria-current="page"]')
    const title = document.querySelector('.ppt-slide [data-ppt-element-name="Title"]')

    return {
      activeName: activeThumb?.querySelector('.ppt-thumb-name')?.textContent ?? '',
      activeSlide,
      firstImportedSlideId: stage?.getAttribute('data-ppt-slide-json-import-first-slide') ?? '',
      format: stage?.getAttribute('data-ppt-slide-json-import-format') ?? '',
      importedCount: Number(stage?.getAttribute('data-ppt-slide-json-import-imported-count') ?? 0),
      model: stage?.getAttribute('data-ppt-slide-json-import-model') ?? '',
      slideCount: document.querySelectorAll('.ppt-thumb').length,
      sourceSlideCount: Number(stage?.getAttribute('data-ppt-slide-json-import-source-slide-count') ?? 0),
      sourceSlideIds: stage?.getAttribute('data-ppt-slide-json-import-source-slide-ids') ?? '',
      sourceSlideNames: stage?.getAttribute('data-ppt-slide-json-import-source-slide-names') ?? '',
      titleText: title?.textContent ?? '',
    }
  })()`)

  record(
    'pastes fenced PPT slide JSON wrapper as editable slides',
    fencedSlideJSONImportState.model === 'ppt-slide-json-import' &&
      fencedSlideJSONImportState.format === 'text-json-ppt-slide' &&
      fencedSlideJSONImportState.sourceSlideCount === 2 &&
      fencedSlideJSONImportState.sourceSlideIds === 'ai-wrapper-slide-a ai-wrapper-slide-b' &&
      fencedSlideJSONImportState.sourceSlideNames === 'Wrapper Slide A, Wrapper Slide B' &&
      fencedSlideJSONImportState.importedCount === 2 &&
      fencedSlideJSONImportState.slideCount === beforeDeckHTMLPaste.slideCount + 2 &&
      fencedSlideJSONImportState.activeSlide === fencedSlideJSONImportState.firstImportedSlideId &&
      fencedSlideJSONImportState.activeName.includes('Wrapper Slide A Copy') &&
      fencedSlideJSONImportState.titleText.includes('Wrapper Slide A'),
    {
      beforeDeckHTMLPaste,
      fencedSlideJSONImportState,
    },
  )

  await deletePPTSlidesByThumbNameIncludes(page, [
    'Wrapper Slide B Copy',
    'Wrapper Slide A Copy',
  ])
  await page.eval(`document.querySelector('.ppt-thumb[aria-label="Open Overview"]')?.click()`)
  await delay(80)

  const afterFencedSlideJSONPasteCleanup = await page.eval(`(() => ({
    activeSlide: document.querySelector('.ppt-slide')?.getAttribute('data-ppt-slide') ?? '',
    slideCount: document.querySelectorAll('.ppt-thumb').length,
  }))()`)

  record(
    'removes fenced PPT slide JSON paste probes before export scenario continues',
    afterFencedSlideJSONPasteCleanup.activeSlide === 'slide-1' &&
      afterFencedSlideJSONPasteCleanup.slideCount === beforeDeckHTMLPaste.slideCount,
    {
      afterFencedSlideJSONPasteCleanup,
      beforeDeckHTMLPaste,
    },
  )

  await page.eval(`(() => {
    const payload = {
      deck: {
        id: 'deck-ai-fenced-json',
        size: { h: 720, w: 1280 },
        slides: [
          {
            background: { color: '#ffffff' },
            elements: [
              {
                geometry: { h: 96, w: 980, x: 112, y: 86 },
                id: 'ai-fenced-title',
                kind: 'textBox',
                name: 'Title',
                style: { color: '#111827', fontSize: 44, fontWeight: 'bold' },
                textBody: { paragraphs: [{ runs: [{ text: 'Fenced JSON Draft' }] }] },
              },
            ],
            id: 'ai-fenced-slide-1',
            name: 'Fenced JSON',
            themeId: 'ppt-theme-default',
          },
        ],
        title: 'AI Fenced JSON Deck',
      },
    }
    const fence = String.fromCharCode(96, 96, 96)
    const markdown = fence + 'json\\n' + JSON.stringify(payload, null, 2) +
      '\\n' + fence
    const dataTransfer = new DataTransfer()

    dataTransfer.setData('text/markdown', markdown)
    dataTransfer.setData('text/plain', markdown)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(140)

  const deckFencedJSONImportState = await page.eval(`(() => {
    const stage = document.querySelector('[data-ppt-app] .ppt-stage-shell')
    const activeSlide = document.querySelector('.ppt-slide')?.getAttribute('data-ppt-slide') ?? ''
    const activeThumb = document.querySelector('.ppt-thumb[aria-current="page"]')
    const title = document.querySelector('.ppt-slide [data-ppt-element-name="Title"]')

    return {
      activeName: activeThumb?.querySelector('.ppt-thumb-name')?.textContent ?? '',
      activeSlide,
      firstImportedSlideId: stage?.getAttribute('data-ppt-deck-json-import-first-slide') ?? '',
      format: stage?.getAttribute('data-ppt-deck-json-import-format') ?? '',
      importedCount: Number(stage?.getAttribute('data-ppt-deck-json-import-imported-count') ?? 0),
      model: stage?.getAttribute('data-ppt-deck-json-import-model') ?? '',
      slideCount: document.querySelectorAll('.ppt-thumb').length,
      sourceDeck: stage?.getAttribute('data-ppt-deck-json-import-source-deck') ?? '',
      sourceSlideCount: Number(stage?.getAttribute('data-ppt-deck-json-import-source-slide-count') ?? 0),
      sourceTitle: stage?.getAttribute('data-ppt-deck-json-import-source-title') ?? '',
      titleText: title?.textContent ?? '',
    }
  })()`)

  record(
    'pastes fenced AI PPT deck JSON as editable slide',
    deckFencedJSONImportState.model === 'ppt-deck-json-import' &&
      deckFencedJSONImportState.format === 'text-json-ppt-deck' &&
      deckFencedJSONImportState.sourceDeck === 'deck-ai-fenced-json' &&
      deckFencedJSONImportState.sourceTitle === 'AI Fenced JSON Deck' &&
      deckFencedJSONImportState.sourceSlideCount === 1 &&
      deckFencedJSONImportState.importedCount === 1 &&
      deckFencedJSONImportState.slideCount === beforeDeckHTMLPaste.slideCount + 1 &&
      deckFencedJSONImportState.activeSlide === deckFencedJSONImportState.firstImportedSlideId &&
      deckFencedJSONImportState.activeName.includes('Fenced JSON Copy') &&
      deckFencedJSONImportState.titleText.includes('Fenced JSON Draft'),
    {
      beforeDeckHTMLPaste,
      deckFencedJSONImportState,
    },
  )

  await deletePPTSlideByThumbName(page, 'Fenced JSON Copy')
  await page.eval(`document.querySelector('.ppt-thumb[aria-label="Open Overview"]')?.click()`)
  await delay(80)

  const afterDeckFencedJSONPasteCleanup = await page.eval(`(() => ({
    activeSlide: document.querySelector('.ppt-slide')?.getAttribute('data-ppt-slide') ?? '',
    slideCount: document.querySelectorAll('.ppt-thumb').length,
  }))()`)

  record(
    'removes PPT fenced deck JSON paste probe before export scenario continues',
    afterDeckFencedJSONPasteCleanup.activeSlide === 'slide-1' &&
      afterDeckFencedJSONPasteCleanup.slideCount === beforeDeckHTMLPaste.slideCount,
    {
      afterDeckFencedJSONPasteCleanup,
      beforeDeckHTMLPaste,
    },
  )

  await page.eval(`(() => {
    const html = document.querySelector('.ppt-export-code')?.value ?? ''
    const fallbackHTML = html.replace(/<script\\b[\\s\\S]*?<\\/script>/gi, '')
    const dataTransfer = new DataTransfer()

    dataTransfer.setData('text/html', fallbackHTML)
    dataTransfer.setData('text/plain', fallbackHTML)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(140)

  const deckHTMLFallbackImportState = await page.eval(`(() => {
    const stage = document.querySelector('[data-ppt-app] .ppt-stage-shell')
    const activeSlide = document.querySelector('.ppt-slide')?.getAttribute('data-ppt-slide') ?? ''
    const activeThumb = document.querySelector('.ppt-thumb[aria-current="page"]')
    const selectedImage = document.querySelector('[data-ppt-element][data-kind="image"] img')
    const selectedImageSrc = selectedImage?.getAttribute('src') ?? ''
    const decodedImageSrc = selectedImageSrc.startsWith('data:image/svg+xml')
      ? decodeURIComponent(selectedImageSrc.split(',').slice(1).join(','))
      : ''

    return {
      activeName: activeThumb?.querySelector('.ppt-thumb-name')?.textContent ?? '',
      activeSlide,
      decodedImageSrc,
      firstImportedSlideId: stage?.getAttribute('data-ppt-deck-html-import-first-slide') ?? '',
      format: stage?.getAttribute('data-ppt-deck-html-import-format') ?? '',
      htmlLength: Number(stage?.getAttribute('data-ppt-deck-html-import-html-length') ?? 0),
      imageCount: document.querySelectorAll('[data-ppt-element][data-kind="image"]').length,
      importedCount: Number(stage?.getAttribute('data-ppt-deck-html-import-imported-count') ?? 0),
      model: stage?.getAttribute('data-ppt-deck-html-import-model') ?? '',
      slideCount: document.querySelectorAll('.ppt-thumb').length,
      sourceDeck: stage?.getAttribute('data-ppt-deck-html-import-source-deck') ?? '',
      sourceSlideCount: Number(stage?.getAttribute('data-ppt-deck-html-import-source-slide-count') ?? 0),
      sourceTitle: stage?.getAttribute('data-ppt-deck-html-import-source-title') ?? '',
    }
  })()`)

  record(
    'pastes script-stripped PPT HTML export as snapshot slides',
    deckHTMLFallbackImportState.model === 'ppt-deck-html-import' &&
      deckHTMLFallbackImportState.format === 'text-html-ppt-deck-fallback' &&
      deckHTMLFallbackImportState.sourceDeck === '' &&
      deckHTMLFallbackImportState.sourceTitle === 'AI Retouch Demo' &&
      deckHTMLFallbackImportState.sourceSlideCount === beforeDeckHTMLPaste.slideCount &&
      deckHTMLFallbackImportState.importedCount === beforeDeckHTMLPaste.slideCount &&
      deckHTMLFallbackImportState.slideCount === beforeDeckHTMLPaste.slideCount * 2 &&
      deckHTMLFallbackImportState.activeSlide === deckHTMLFallbackImportState.firstImportedSlideId &&
      deckHTMLFallbackImportState.activeName.includes('Snapshot') &&
      deckHTMLFallbackImportState.imageCount === 1 &&
      deckHTMLFallbackImportState.decodedImageSrc.includes('<foreignObject') &&
      !deckHTMLFallbackImportState.decodedImageSrc.includes('<script') &&
      !deckHTMLFallbackImportState.decodedImageSrc.includes('onload='),
    {
      beforeDeckHTMLPaste,
      deckHTMLFallbackImportState,
    },
  )

  await deletePPTSlidesByThumbNameIncludes(page, ['Snapshot'])
  await page.eval(`document.querySelector('.ppt-thumb[aria-label="Open Overview"]')?.click()`)
  await delay(80)

  const afterDeckHTMLFallbackPasteCleanup = await page.eval(`(() => ({
    activeSlide: document.querySelector('.ppt-slide')?.getAttribute('data-ppt-slide') ?? '',
    slideCount: document.querySelectorAll('.ppt-thumb').length,
  }))()`)

  record(
    'removes PPT HTML deck fallback paste probes before export scenario continues',
    afterDeckHTMLFallbackPasteCleanup.activeSlide === 'slide-1' &&
      afterDeckHTMLFallbackPasteCleanup.slideCount === beforeDeckHTMLPaste.slideCount,
    {
      afterDeckHTMLFallbackPasteCleanup,
      beforeDeckHTMLPaste,
    },
  )

  await page.eval(`(() => {
    const markdown = [
      '# AI Draft Outline',
      '',
      '## Market context',
      '- Current demand is fragmented',
      '- **Risk** needs tighter framing',
      'Notes: Mention the source assumptions.',
      '',
      '## Retouch plan',
      '1. Draft compact slides',
      '2. Review final copy',
    ].join('\\n')
    const dataTransfer = new DataTransfer()

    dataTransfer.setData('text/markdown', markdown)
    dataTransfer.setData('text/plain', markdown)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(140)

  const deckOutlineImportState = await page.eval(`(() => {
    const stage = document.querySelector('[data-ppt-app] .ppt-stage-shell')
    const activeSlide = document.querySelector('.ppt-slide')?.getAttribute('data-ppt-slide') ?? ''
    const activeThumb = document.querySelector('.ppt-thumb[aria-current="page"]')
    const selectedBody = document.querySelector('.ppt-slide [data-ppt-element$="-body"]')
    const selectedTitle = document.querySelector('.ppt-slide [data-ppt-element$="-title"]')
    const exportCode = document.querySelector('.ppt-export-code')?.value ?? ''

    return {
      activeName: activeThumb?.querySelector('.ppt-thumb-name')?.textContent ?? '',
      activeSlide,
      bodyText: selectedBody?.textContent ?? '',
      bulletCount: selectedBody?.querySelectorAll('[data-ppt-bullet="true"]').length ?? 0,
      exportHasLayout: exportCode.includes('"layoutId": "ppt-layout-title-body"'),
      exportHasNotes: exportCode.includes('Mention the source assumptions.'),
      exportHasNumbered: exportCode.includes('"bullet": "numbered"') && exportCode.includes('Draft compact slides'),
      firstImportedSlideId: stage?.getAttribute('data-ppt-deck-outline-import-first-slide') ?? '',
      format: stage?.getAttribute('data-ppt-deck-outline-import-format') ?? '',
      importedCount: Number(stage?.getAttribute('data-ppt-deck-outline-import-imported-count') ?? 0),
      model: stage?.getAttribute('data-ppt-deck-outline-import-model') ?? '',
      slideCount: document.querySelectorAll('.ppt-thumb').length,
      sourceSlideCount: Number(stage?.getAttribute('data-ppt-deck-outline-import-source-slide-count') ?? 0),
      sourceTitle: stage?.getAttribute('data-ppt-deck-outline-import-source-title') ?? '',
      textLength: Number(stage?.getAttribute('data-ppt-deck-outline-import-text-length') ?? 0),
      titleText: selectedTitle?.textContent ?? '',
    }
  })()`)

  record(
    'pastes Markdown outline as editable PPT slides',
    deckOutlineImportState.model === 'ppt-deck-markdown-outline-import' &&
      deckOutlineImportState.format === 'text-markdown-ppt-outline' &&
      deckOutlineImportState.sourceTitle === 'AI Draft Outline' &&
      deckOutlineImportState.sourceSlideCount === 2 &&
      deckOutlineImportState.importedCount === 2 &&
      deckOutlineImportState.slideCount === beforeDeckHTMLPaste.slideCount + 2 &&
      deckOutlineImportState.activeSlide === deckOutlineImportState.firstImportedSlideId &&
      deckOutlineImportState.activeName.includes('Market context') &&
      deckOutlineImportState.titleText.includes('Market context') &&
      deckOutlineImportState.bodyText.includes('Current demand is fragmented') &&
      deckOutlineImportState.bodyText.includes('Risk needs tighter framing') &&
      deckOutlineImportState.bulletCount === 2 &&
      deckOutlineImportState.textLength > 0 &&
      deckOutlineImportState.exportHasNotes &&
      deckOutlineImportState.exportHasNumbered &&
      deckOutlineImportState.exportHasLayout,
    {
      beforeDeckHTMLPaste,
      deckOutlineImportState,
    },
  )

  await deletePPTSlidesByThumbNameIncludes(page, [
    'Retouch plan',
    'Market context',
  ])
  await page.eval(`document.querySelector('.ppt-thumb[aria-label="Open Overview"]')?.click()`)
  await delay(80)

  const afterDeckOutlinePasteCleanup = await page.eval(`(() => ({
    activeSlide: document.querySelector('.ppt-slide')?.getAttribute('data-ppt-slide') ?? '',
    slideCount: document.querySelectorAll('.ppt-thumb').length,
  }))()`)

  record(
    'removes PPT Markdown outline paste probes before export scenario continues',
    afterDeckOutlinePasteCleanup.activeSlide === 'slide-1' &&
      afterDeckOutlinePasteCleanup.slideCount === beforeDeckHTMLPaste.slideCount,
    {
      afterDeckOutlinePasteCleanup,
      beforeDeckHTMLPaste,
    },
  )

  await page.eval(`(() => {
    window.__pptClipboardItemTypes = []
    window.__pptClipboardWriteCount = 0
    window.__pptClipboardWriteText = ''
  })()`)
  await page.eval(`document.querySelector('[data-ppt-copy-slide-svg]')?.click()`)
  await delay(80)

  const slideSvgClipboardState = await page.eval(`(() => {
    const stage = document.querySelector('[data-ppt-app] .ppt-stage-shell')
    const mimeType = stage?.getAttribute('data-ppt-slide-svg-clipboard-json-mime-type') ?? ''
    const itemTypes = window.__pptClipboardItemTypes?.at(-1) ?? []

    return {
      hasCustomJson: itemTypes.includes(mimeType),
      hasHTML: itemTypes.includes('text/html'),
      hasPlainText: itemTypes.includes('text/plain'),
      hasSVG: itemTypes.includes('image/svg+xml'),
      itemTypes,
      mimeType,
      model: stage?.getAttribute('data-ppt-slide-svg-clipboard-model') ?? '',
      sourceSlide: stage?.getAttribute('data-ppt-slide-svg-clipboard-source-slide') ?? '',
      svgLength: Number(stage?.getAttribute('data-ppt-slide-svg-clipboard-svg-length') ?? 0),
      writeCount: window.__pptClipboardWriteCount ?? 0,
      writeMode: stage?.getAttribute('data-ppt-slide-svg-clipboard-write-mode') ?? '',
      writeTextLength: (window.__pptClipboardWriteText ?? '').length,
    }
  })()`)

  record(
    'copies active PPT slide as SVG through canvas rich clipboard writer',
    slideSvgClipboardState.model === 'canvas-rich-slide-svg-clipboard' &&
      slideSvgClipboardState.sourceSlide === 'slide-1' &&
      slideSvgClipboardState.writeMode === 'clipboard-item' &&
      slideSvgClipboardState.writeCount === 1 &&
      slideSvgClipboardState.svgLength > 1000 &&
      slideSvgClipboardState.hasCustomJson &&
      slideSvgClipboardState.hasHTML &&
      slideSvgClipboardState.hasPlainText &&
      slideSvgClipboardState.hasSVG &&
      slideSvgClipboardState.writeTextLength === 0,
    slideSvgClipboardState,
  )

  await page.eval(`document.querySelector('[data-ppt-export-svg]')?.click()`)
  await delay(80)

  const slideSvgState = await page.eval(`(() => {
    const download = (window.__pptDownloads ?? [])
      .find((entry) => entry.download === 'slide-1.svg') ?? {}
    const text = download.text ?? ''
    const objectAltText = ${JSON.stringify(PPT_OBJECT_ALT_TEXT)}

    return {
      download: download.download ?? '',
      hasAnimation: text.includes('data-ppt-animation-type="flyIn"') && text.includes('data-ppt-animation-trigger="withPrevious"') && text.includes('data-ppt-animation-duration="800"') && text.includes('data-ppt-animation-delay="200"') && text.includes('data-ppt-animation-order="3"'),
      hasBackground: text.includes('data-ppt-svg-background="true"'),
      hasObjectOpacity: text.includes('data-ppt-opacity="0.42"') && text.includes('opacity="0.42"'),
      hasObjectShadow: text.includes('data-ppt-shadow="true"') && text.includes('data-ppt-shadow-color="#334155"') && text.includes('data-ppt-shadow-opacity="0.36"') && text.includes('filter:drop-shadow'),
      hasObjectHyperlink: text.includes('data-ppt-hyperlink-url="https://example.com/ppt"'),
      hasObjectAltText: text.includes('data-ppt-alt-text="' + objectAltText + '"') && text.includes('<title>' + objectAltText + '</title>'),
      hasCornerRadius: text.includes('data-ppt-corner-radius="36"') && text.includes('rx="36"'),
      hasFillOpacity: text.includes('data-ppt-fill-opacity="0.35"') && text.includes('fill-opacity="0.35"'),
      hasStrokeDash: text.includes('data-ppt-stroke-dash="dash"') && text.includes('data-ppt-stroke-dash="dot"') && text.includes('stroke-dasharray='),
      hasFontFamily: text.includes('data-ppt-font-family="Georgia"') && text.includes('font-family="Georgia, serif"'),
      hasParagraphSpacing: text.includes('data-ppt-line-height="1.28"') && text.includes('data-ppt-spacing-before="9"') && text.includes('data-ppt-spacing-after="15"'),
      hasTextFrameInset: text.includes('data-ppt-text-inset="8,12,16,20"'),
      hasTextVerticalAlign: text.includes('data-ppt-vertical-align="bottom"'),
      hasComment: text.includes('data-ppt-kind="comment"') && text.includes('data-ppt-comment-body="true"'),
      hasFreeform: text.includes('data-ppt-kind="freeform"') && text.includes('data-ppt-freeform-path'),
      hasImage: text.includes('data-ppt-kind="image"') && text.includes('href="data:image/svg+xml'),
      hasLayout: text.includes('data-ppt-svg-layout-id="ppt-layout-split"'),
      hasLine: text.includes('data-ppt-kind="line"') && (text.includes('<line ') || text.includes('data-ppt-line-path')),
      hasPlaceholderVisibility: text.includes('data-ppt-svg-hidden-placeholders="media"'),
      hasScope: text.includes('data-ppt-svg-scope="slide"'),
      hasShape: text.includes('data-ppt-kind="shape"'),
      hasSlide: text.includes('data-ppt-svg-slide="slide-1"'),
      hasSvg: text.includes('<svg xmlns="http://www.w3.org/2000/svg"'),
      hasTable: text.includes('data-ppt-kind="table"') && text.includes('data-ppt-table-cell=') && text.includes('data-ppt-table-text='),
      hasText: text.includes('data-ppt-kind="textBox"') && text.includes('<text '),
      hasTextAutoFit: text.includes('data-ppt-text-autofit="resizeShapeToFitText"'),
      hasTheme: text.includes('data-ppt-svg-theme-id="ppt-theme-default"'),
      hasTransition: text.includes('data-ppt-svg-transition-type="push"') && text.includes('data-ppt-svg-transition-duration="650"') && text.includes('data-ppt-svg-transition-advance-on-click="false"') && text.includes('data-ppt-svg-transition-advance-after="3000"'),
      type: download.type ?? '',
    }
  })()`)

  record('downloads active PPT slide as SVG', slideSvgState.download === 'slide-1.svg' && slideSvgState.type.includes('image/svg+xml') && slideSvgState.hasSvg && slideSvgState.hasSlide && slideSvgState.hasScope && slideSvgState.hasBackground, slideSvgState)
  record('exports PPT image/shape/text/line/freeform/table/comment into slide SVG', slideSvgState.hasImage && slideSvgState.hasShape && slideSvgState.hasText && slideSvgState.hasLine && slideSvgState.hasFreeform && slideSvgState.hasTable && slideSvgState.hasComment, slideSvgState)
  record('exports PPT object animation metadata into slide SVG', slideSvgState.hasAnimation, slideSvgState)
  record('exports PPT object opacity metadata into slide SVG', slideSvgState.hasObjectOpacity, slideSvgState)
  record('exports PPT object shadow metadata into slide SVG', slideSvgState.hasObjectShadow, slideSvgState)
  record('exports PPT object hyperlink metadata into slide SVG', slideSvgState.hasObjectHyperlink, slideSvgState)
  record('exports PPT object alt text metadata into slide SVG', slideSvgState.hasObjectAltText, slideSvgState)
  record('exports PPT shape corner radius metadata into slide SVG', slideSvgState.hasCornerRadius, slideSvgState)
  record('exports PPT shape fill opacity metadata into slide SVG', slideSvgState.hasFillOpacity, slideSvgState)
  record('exports PPT stroke dash style metadata into slide SVG', slideSvgState.hasStrokeDash, slideSvgState)
  record('exports PPT font family metadata into slide SVG', slideSvgState.hasFontFamily, slideSvgState)
  record('exports PPT paragraph spacing metadata into slide SVG', slideSvgState.hasParagraphSpacing, slideSvgState)
  record('exports PPT text frame inset metadata into slide SVG', slideSvgState.hasTextFrameInset, slideSvgState)
  record('exports PPT text vertical alignment metadata into slide SVG', slideSvgState.hasTextVerticalAlign, slideSvgState)
  record('exports PPT text auto-fit metadata into slide SVG', slideSvgState.hasTextAutoFit, slideSvgState)
  record('exports PPT layout/theme metadata into slide SVG', slideSvgState.hasLayout && slideSvgState.hasTheme, slideSvgState)
  record('exports PPT placeholder visibility metadata into slide SVG', slideSvgState.hasPlaceholderVisibility, slideSvgState)
  record('exports PPT slide transition metadata into slide SVG', slideSvgState.hasTransition, slideSvgState)

  const imageId = await page.eval(`(() => {
    const image =
      document.querySelector('[data-kind="image"][data-ppt-image-fit="contain"][data-ppt-image-crop-x="25"][data-ppt-image-crop-y="70"][data-ppt-flip-h="true"]') ??
      [...document.querySelectorAll('[data-kind="image"]')].at(-1)

    return image?.getAttribute('data-ppt-element') ?? ''
  })()`)
  await selectPPTLayerRows(page, [imageId])
  await delay(80)

  const beforeSelectionSvg = await page.eval(`(() => ({
    copyDisabled: document.querySelector('[data-ppt-copy-selection-svg]')?.disabled ?? true,
    disabled: document.querySelector('[data-ppt-export-selection-svg]')?.disabled ?? true,
    selectedId: document.querySelector('[data-selected="true"]')?.getAttribute('data-ppt-element') ?? '',
  }))()`)

  await page.eval(`(() => {
    window.__pptClipboardItemTypes = []
    window.__pptClipboardWriteCount = 0
    window.__pptClipboardWriteText = ''
  })()`)
  await page.eval(`document.querySelector('[data-ppt-copy-selection-svg]')?.click()`)
  await delay(80)

  const selectionSvgClipboardState = await page.eval(`(() => {
    const stage = document.querySelector('[data-ppt-app] .ppt-stage-shell')
    const mimeType = stage?.getAttribute('data-ppt-selection-svg-clipboard-json-mime-type') ?? ''
    const itemTypes = window.__pptClipboardItemTypes?.at(-1) ?? []

    return {
      hasCustomJson: itemTypes.includes(mimeType),
      hasHTML: itemTypes.includes('text/html'),
      hasPlainText: itemTypes.includes('text/plain'),
      hasSVG: itemTypes.includes('image/svg+xml'),
      itemTypes,
      mimeType,
      model: stage?.getAttribute('data-ppt-selection-svg-clipboard-model') ?? '',
      selection: stage?.getAttribute('data-ppt-selection-svg-clipboard-selection') ?? '',
      sourceSlide: stage?.getAttribute('data-ppt-selection-svg-clipboard-source-slide') ?? '',
      svgLength: Number(stage?.getAttribute('data-ppt-selection-svg-clipboard-svg-length') ?? 0),
      writeCount: window.__pptClipboardWriteCount ?? 0,
      writeMode: stage?.getAttribute('data-ppt-selection-svg-clipboard-write-mode') ?? '',
      writeTextLength: (window.__pptClipboardWriteText ?? '').length,
    }
  })()`)

  await page.eval(`document.querySelector('[data-ppt-export-selection-svg]')?.click()`)
  await delay(80)

  const selectionSvgState = await page.eval(`(() => {
    const selectedId = document.querySelector('[data-selected="true"]')?.getAttribute('data-ppt-element') ?? ''
    const download = (window.__pptDownloads ?? [])
      .find((entry) => entry.download === 'slide-1-selection.svg') ?? {}
    const text = download.text ?? ''

    return {
      download: download.download ?? '',
      hasCrop: text.includes('data-ppt-image-crop-x="25"') && text.includes('data-ppt-image-crop-y="70"'),
      hasFit: text.includes('data-ppt-image-fit="contain"'),
      hasFlip: text.includes('data-ppt-flip-h="true"') && text.includes('scale(-1 1)'),
      hasOnlySelectedImage: text.includes(\`data-ppt-element="\${selectedId}"\`) && !text.includes('data-ppt-element="s1-title"'),
      hasScope: text.includes('data-ppt-svg-scope="selection"'),
      hasSvg: text.includes('<svg xmlns="http://www.w3.org/2000/svg"'),
      selectedId,
      type: download.type ?? '',
    }
  })()`)

  record('enables selected-object PPT SVG export', beforeSelectionSvg.selectedId === imageId && !beforeSelectionSvg.disabled && !beforeSelectionSvg.copyDisabled, beforeSelectionSvg)
  record(
    'copies selected PPT objects as SVG through canvas rich clipboard writer',
    selectionSvgClipboardState.model === 'canvas-rich-selection-svg-clipboard' &&
      selectionSvgClipboardState.sourceSlide === 'slide-1' &&
      selectionSvgClipboardState.selection === imageId &&
      selectionSvgClipboardState.writeMode === 'clipboard-item' &&
      selectionSvgClipboardState.writeCount === 1 &&
      selectionSvgClipboardState.svgLength > 100 &&
      selectionSvgClipboardState.hasCustomJson &&
      selectionSvgClipboardState.hasHTML &&
      selectionSvgClipboardState.hasPlainText &&
      selectionSvgClipboardState.hasSVG &&
      selectionSvgClipboardState.writeTextLength === 0,
    selectionSvgClipboardState,
  )
  record('downloads selected PPT objects as SVG', selectionSvgState.download === 'slide-1-selection.svg' && selectionSvgState.type.includes('image/svg+xml') && selectionSvgState.hasSvg && selectionSvgState.hasScope && selectionSvgState.hasOnlySelectedImage, selectionSvgState)
  record('preserves PPT image fit/crop/flip metadata in selection SVG', selectionSvgState.hasFit && selectionSvgState.hasCrop && selectionSvgState.hasFlip, selectionSvgState)

  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)
  await pressKey(page, {
    code: 'KeyK',
    key: 'k',
    modifiers: 2,
    windowsVirtualKeyCode: 75,
  })
  await delay(80)
  await page.send('Input.insertText', { text: 'selection svg' })
  await delay(80)

  const disabledState = await page.eval(`(() => ({
    itemPresent: !!document.querySelector('[data-ppt-command-palette-item="export:download-selection-svg"]'),
    paletteDisabled: document.querySelector('[data-ppt-command-palette-item="export:download-selection-svg"]')?.disabled ?? false,
    paletteOpen: !!document.querySelector('[data-ppt-command-palette]'),
    selectedCount: document.querySelectorAll('[data-selected="true"]').length,
    toolbarDisabled: document.querySelector('[data-ppt-export-selection-svg]')?.disabled ?? false,
  }))()`)

  record('disables PPT selection SVG export without selection', disabledState.paletteOpen && disabledState.itemPresent && disabledState.selectedCount === 0 && disabledState.toolbarDisabled && disabledState.paletteDisabled, disabledState)

  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)
}

async function runAlignmentPopoverScenario(page) {
  await page.eval(`document.querySelector('.ppt-thumb[aria-label="Open Overview"]')?.click()`)
  await delay(80)

  const alignmentTargetState = await page.eval(`(() => ({
    activeSlide: document.querySelector('.ppt-slide')?.getAttribute('data-ppt-slide') ?? '',
    elementIds: [...document.querySelectorAll('.ppt-slide [data-ppt-element]')]
      .map((element) => element.getAttribute('data-ppt-element') ?? ''),
    hasS1Card1: !!document.querySelector('[data-ppt-element="s1-card-1"]'),
    thumbLabels: [...document.querySelectorAll('.ppt-thumb')]
      .map((thumb) => thumb.getAttribute('aria-label') ?? ''),
    thumbNames: [...document.querySelectorAll('.ppt-thumb .ppt-thumb-name')]
      .map((name) => name.textContent ?? ''),
  }))()`)

  if (!alignmentTargetState.hasS1Card1) {
    throw new Error(JSON.stringify({
      message: 'PPT alignment target is missing',
      state: alignmentTargetState,
    }))
  }

  const point = await page.eval(`(() => {
    const rect = document.querySelector('[data-ppt-element="s1-card-1"]').getBoundingClientRect()

    return {
      x: rect.left + rect.width / 2,
      y: rect.top + rect.height / 2,
    }
  })()`)

  await clickMouse(page, point.x, point.y, 1)
  await delay(80)

  await page.eval(`document.querySelector('[data-ppt-alignment-popover-trigger]')?.click()`)
  await delay(50)

  const alignmentPopoverOpen = await page.eval(`(() => {
    const trigger = document.querySelector('[data-ppt-alignment-popover-trigger]')
    const popover = document.querySelector('[data-ppt-alignment-popover]')
    const items = [...document.querySelectorAll('[data-ppt-alignment-popover-item]')]

    return {
      active: popover?.getAttribute('data-ppt-alignment-popover-active') ?? '',
      canvasMenuItemCount: items.filter((item) => item.hasAttribute('data-canvas-menu-item')).length,
      expanded: trigger?.getAttribute('aria-expanded') ?? '',
      itemCommands: items.map((item) => item.getAttribute('data-ppt-alignment-popover-command') ?? ''),
      itemCount: items.length,
      model: popover?.getAttribute('data-ppt-alignment-popover-model') ?? '',
      role: popover?.getAttribute('role') ?? '',
      selectedId: document.querySelector('[data-selected="true"]')?.getAttribute('data-ppt-element') ?? '',
      triggerHasPopup: trigger?.getAttribute('aria-haspopup') ?? '',
    }
  })()`)

  record('opens PPT selection alignment popover from floating bar', alignmentPopoverOpen.selectedId !== '' && alignmentPopoverOpen.expanded === 'true' && alignmentPopoverOpen.role === 'menu' && alignmentPopoverOpen.model === 'canvas-dom-alignment-popover' && alignmentPopoverOpen.triggerHasPopup === 'menu' && alignmentPopoverOpen.itemCount === 8 && alignmentPopoverOpen.canvasMenuItemCount === alignmentPopoverOpen.itemCount && alignmentPopoverOpen.itemCommands.includes('align-center-x') && alignmentPopoverOpen.itemCommands.includes('distribute-horizontal'), alignmentPopoverOpen)

  await page.eval(`document.querySelector('[data-ppt-alignment-popover-command="align-center-x"]')?.focus()`)
  await delay(40)

  const alignmentPreview = await page.eval(`(() => {
    const stage = document.querySelector('.ppt-stage-shell')

    return {
      active: document.querySelector('[data-ppt-alignment-popover]')?.getAttribute('data-ppt-alignment-popover-active') ?? '',
      focusedCommand: document.activeElement?.getAttribute('data-ppt-alignment-popover-command') ?? '',
      preview: stage?.getAttribute('data-ppt-alignment-popover-preview') ?? '',
      previewModel: stage?.getAttribute('data-ppt-alignment-popover-preview-model') ?? '',
    }
  })()`)

  record('exposes PPT alignment popover focus preview metadata', alignmentPreview.focusedCommand === 'align-center-x' && alignmentPreview.active === 'alignCenter' && alignmentPreview.preview === 'alignCenter' && alignmentPreview.previewModel === 'canvas-dom-alignment-preview-guide', alignmentPreview)

  await pressKey(page, {
    code: 'ArrowDown',
    key: 'ArrowDown',
    windowsVirtualKeyCode: 40,
  })
  await delay(40)
  await pressKey(page, {
    code: 'End',
    key: 'End',
    windowsVirtualKeyCode: 35,
  })
  await delay(40)
  await pressKey(page, {
    code: 'Home',
    key: 'Home',
    windowsVirtualKeyCode: 36,
  })
  await delay(40)

  const alignmentKeyboard = await page.eval(`(() => {
    const stage = document.querySelector('.ppt-stage-shell')

    return {
      active: document.querySelector('[data-ppt-alignment-popover]')?.getAttribute('data-ppt-alignment-popover-active') ?? '',
      focusedCommand: document.activeElement?.getAttribute('data-ppt-alignment-popover-command') ?? '',
      preview: stage?.getAttribute('data-ppt-alignment-popover-preview') ?? '',
    }
  })()`)

  record('moves PPT alignment popover focus with Arrow Home End keys', alignmentKeyboard.focusedCommand === 'align-left' && alignmentKeyboard.active === 'alignLeft' && alignmentKeyboard.preview === 'alignLeft', alignmentKeyboard)

  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(40)

  const alignmentEscape = await page.eval(`(() => ({
    expanded: document.querySelector('[data-ppt-alignment-popover-trigger]')?.getAttribute('aria-expanded') ?? '',
    focusedTrigger: document.activeElement?.hasAttribute('data-ppt-alignment-popover-trigger') ?? false,
    open: !!document.querySelector('[data-ppt-alignment-popover]'),
    preview: document.querySelector('.ppt-stage-shell')?.getAttribute('data-ppt-alignment-popover-preview') ?? '',
  }))()`)

  record('closes PPT alignment popover with Escape and restores trigger focus', alignmentEscape.expanded === 'false' && alignmentEscape.focusedTrigger && !alignmentEscape.open && alignmentEscape.preview === '', alignmentEscape)

  await page.eval(`document.querySelector('[data-ppt-alignment-popover-trigger]')?.click()`)
  await delay(40)
  await page.eval(`document.querySelector('[data-ppt-alignment-popover-command="align-left"]')?.click()`)
  await delay(50)

  const afterPopoverAlign = await page.eval(`(() => {
    const selected = document.querySelector('[data-selected="true"]')

    return {
      expanded: document.querySelector('[data-ppt-alignment-popover-trigger]')?.getAttribute('aria-expanded') ?? '',
      left: parseFloat(selected?.style.left ?? '0'),
      open: !!document.querySelector('[data-ppt-alignment-popover]'),
      selectedId: selected?.getAttribute('data-ppt-element') ?? '',
    }
  })()`)

  record('runs PPT alignment command from selection popover', afterPopoverAlign.selectedId === alignmentPopoverOpen.selectedId && afterPopoverAlign.expanded === 'false' && !afterPopoverAlign.open && afterPopoverAlign.left <= 1, {
    afterPopoverAlign,
    alignmentPopoverOpen,
  })
}

async function runShapeMenuScenario(page) {
  await page.eval(`document.querySelector('.ppt-thumb[aria-label="Open Overview"]')?.click()`)
  await delay(80)

  const point = await page.eval(`(() => {
    const element =
      document.querySelector('[data-kind="shape"][data-shape="rect"]:not([data-locked="true"])') ??
      document.querySelector('[data-kind="shape"]:not([data-locked="true"])')

    if (!(element instanceof HTMLElement)) {
      return {
        id: '',
        shape: '',
        x: 0,
        y: 0,
      }
    }

    return {
      id: element.getAttribute('data-ppt-element'),
      shape: element.getAttribute('data-shape') ?? '',
    }
  })()`)

  await selectPPTLayerRows(page, [point.id])
  await delay(80)

  await page.eval(`document.querySelector('[data-ppt-shape-menu-trigger]')?.click()`)
  await delay(50)

  const shapeMenuOpen = await page.eval(`(() => {
    const trigger = document.querySelector('[data-ppt-shape-menu-trigger]')
    const menu = document.querySelector('[data-ppt-shape-menu]')
    const items = [...document.querySelectorAll('[data-ppt-shape-menu-item]')]

    return {
      active: menu?.getAttribute('data-ppt-shape-menu-active') ?? '',
      checked: items.map((item) => item.getAttribute('aria-checked') ?? ''),
      checkedShape: items.find((item) => item.getAttribute('aria-checked') === 'true')?.getAttribute('data-ppt-shape-menu-item') ?? '',
      controls: trigger?.getAttribute('aria-controls') ?? '',
      canvasMenuItemCount: items.filter((item) => item.hasAttribute('data-canvas-menu-item')).length,
      expanded: trigger?.getAttribute('aria-expanded') ?? '',
      hasPopup: trigger?.getAttribute('aria-haspopup') ?? '',
      itemCount: items.length,
      itemRoles: items.map((item) => item.getAttribute('role') ?? ''),
      itemShapes: items.map((item) => item.getAttribute('data-ppt-shape-menu-item') ?? ''),
      model: menu?.getAttribute('data-ppt-shape-menu-model') ?? '',
      role: menu?.getAttribute('role') ?? '',
      selectedId: document.querySelector('[data-selected="true"]')?.getAttribute('data-ppt-element') ?? '',
    }
  })()`)

  record('opens PPT Shape menu from selection floating bar', point.id !== '' && shapeMenuOpen.selectedId === point.id && shapeMenuOpen.expanded === 'true' && shapeMenuOpen.hasPopup === 'menu' && shapeMenuOpen.role === 'menu' && shapeMenuOpen.model === 'canvas-selection-toolbar-dropdown-menu' && shapeMenuOpen.controls === 'ppt-shape-kind-menu' && shapeMenuOpen.itemCount === 3 && shapeMenuOpen.canvasMenuItemCount === shapeMenuOpen.itemCount && shapeMenuOpen.itemRoles.every((role) => role === 'menuitemcheckbox') && shapeMenuOpen.itemShapes.join(' ') === 'rect ellipse diamond' && shapeMenuOpen.checkedShape === point.shape, {
    point,
    shapeMenuOpen,
  })

  await page.eval(`document.querySelector('[data-ppt-shape-menu-item="${point.shape}"]')?.focus()`)
  await delay(40)
  await pressKey(page, {
    code: 'ArrowDown',
    key: 'ArrowDown',
    windowsVirtualKeyCode: 40,
  })
  await delay(40)
  await pressKey(page, {
    code: 'End',
    key: 'End',
    windowsVirtualKeyCode: 35,
  })
  await delay(40)
  await pressKey(page, {
    code: 'Home',
    key: 'Home',
    windowsVirtualKeyCode: 36,
  })
  await delay(40)

  const shapeMenuKeyboard = await page.eval(`(() => ({
    active: document.querySelector('[data-ppt-shape-menu]')?.getAttribute('data-ppt-shape-menu-active') ?? '',
    focusedShape: document.activeElement?.getAttribute('data-ppt-shape-menu-item') ?? '',
  }))()`)

  record('moves PPT Shape menu focus with Arrow Home End keys', shapeMenuKeyboard.active === 'rect' && shapeMenuKeyboard.focusedShape === 'rect', shapeMenuKeyboard)

  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(40)

  const shapeMenuEscape = await page.eval(`(() => ({
    expanded: document.querySelector('[data-ppt-shape-menu-trigger]')?.getAttribute('aria-expanded') ?? '',
    focusedTrigger: document.activeElement?.hasAttribute('data-ppt-shape-menu-trigger') ?? false,
    open: !!document.querySelector('[data-ppt-shape-menu]'),
    selectedId: document.querySelector('[data-selected="true"]')?.getAttribute('data-ppt-element') ?? '',
  }))()`)

  record('closes PPT Shape menu with Escape and restores trigger focus', shapeMenuEscape.expanded === 'false' && shapeMenuEscape.focusedTrigger && !shapeMenuEscape.open && shapeMenuEscape.selectedId === point.id, {
    point,
    shapeMenuEscape,
  })

  const targetShape = point.shape === 'ellipse' ? 'diamond' : 'ellipse'

  await page.eval(`document.querySelector('[data-ppt-shape-menu-trigger]')?.click()`)
  await delay(40)
  await page.eval(`document.querySelector('[data-ppt-shape-menu-item="${targetShape}"]')?.click()`)
  await delay(80)

  const afterShapeMenuChange = await page.eval(`(() => {
    const selected = document.querySelector('[data-selected="true"]')
    const code = document.querySelector('.ppt-export-code')?.value ?? ''

    return {
      expanded: document.querySelector('[data-ppt-shape-menu-trigger]')?.getAttribute('aria-expanded') ?? '',
      exportHasShapeClass: code.includes('ppt-shape-${targetShape}'),
      exportHasShapeModel: code.includes('"shape": "${targetShape}"'),
      inspectorShape: document.querySelector('[data-ppt-style-field="shape"]')?.value ?? '',
      open: !!document.querySelector('[data-ppt-shape-menu]'),
      selectedId: selected?.getAttribute('data-ppt-element') ?? '',
      selectedShape: selected?.getAttribute('data-shape') ?? '',
    }
  })()`)

  record('changes selected PPT shape kind from Shape menu', afterShapeMenuChange.selectedId === point.id && afterShapeMenuChange.expanded === 'false' && !afterShapeMenuChange.open && afterShapeMenuChange.selectedShape === targetShape && afterShapeMenuChange.inspectorShape === targetShape && afterShapeMenuChange.exportHasShapeClass && afterShapeMenuChange.exportHasShapeModel, {
    afterShapeMenuChange,
    point,
    targetShape,
  })
}

async function runViewAndShapeScenario(page) {
  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)

  await page.eval(`document.querySelector('[data-ppt-view-grid]').click()`)
  await delay(50)

  const afterGridToggle = await page.eval(`(() => ({
    grid: document.querySelector('.ppt-stage-shell')?.getAttribute('data-grid'),
    pressed: document.querySelector('[data-ppt-view-grid]')?.getAttribute('aria-pressed'),
    zoomLabel: document.querySelector('.ppt-zoom-label')?.textContent ?? '',
  }))()`)

  record('toggles PPT editing grid visibility', afterGridToggle.grid === 'false' && afterGridToggle.pressed === 'false', afterGridToggle)
  record('shows PPT zoom percentage', /\d+%/.test(afterGridToggle.zoomLabel), afterGridToggle)

  await page.eval(`document.querySelector('[data-ppt-view-frame-guides]').click()`)
  await delay(50)

  const afterFrameGuideToggle = await page.eval(`(() => ({
    frameGuides: document.querySelector('.ppt-stage-shell')?.getAttribute('data-frame-guides'),
    guideLayerCount: document.querySelectorAll('[data-ppt-frame-guides]').length,
    pressed: document.querySelector('[data-ppt-view-frame-guides]')?.getAttribute('aria-pressed'),
  }))()`)

  record('toggles PPT frame guide visibility from toolbar', afterFrameGuideToggle.frameGuides === 'false' && afterFrameGuideToggle.pressed === 'false' && afterFrameGuideToggle.guideLayerCount === 0, afterFrameGuideToggle)

  await page.eval(`document.querySelector('[data-ppt-view-frame-guides]').click()`)
  await delay(50)

  await page.eval(`document.querySelector('[data-ppt-insert-shape="ellipse"]').click()`)
  await delay(20)

  const createEllipse = await page.eval(`(() => {
    const slide = document.querySelector('.ppt-slide').getBoundingClientRect()

    return {
      pressed: document.querySelector('[data-ppt-insert-shape="ellipse"]')?.getAttribute('aria-pressed'),
      x: slide.left + slide.width * 0.18,
      y: slide.top + slide.height * 0.3,
    }
  })()`)

  await clickMouse(page, createEllipse.x, createEllipse.y, 1)
  await delay(80)

  const afterInsertShape = await page.eval(`(() => {
    const selected = document.querySelector('[data-selected="true"]')

    return {
      height: parseFloat(selected?.style.height ?? '0'),
      selectedCount: document.querySelectorAll('[data-selected="true"]').length,
      shape: selected?.getAttribute('data-shape') ?? null,
      text: selected?.textContent ?? '',
      width: parseFloat(selected?.style.width ?? '0'),
    }
  })()`)

  record('creates PPT oval shape from canvas tool click', createEllipse.pressed === 'true' && afterInsertShape.selectedCount === 1 && afterInsertShape.shape === 'ellipse' && afterInsertShape.width > 100 && afterInsertShape.height > 80, {
    afterInsertShape,
    createEllipse,
  })

  await page.eval(`(() => {
    const shape = document.querySelector('[data-ppt-style-field="shape"]')
    shape.value = 'diamond'
    shape.dispatchEvent(new Event('change', { bubbles: true }))
  })()`)
  await delay(50)

  const afterShapeChange = await page.eval(`(() => {
    const selected = document.querySelector('[data-selected="true"]')

    return {
      shape: selected?.getAttribute('data-shape') ?? null,
      thumbDiamondCount: document.querySelectorAll('.ppt-thumb-shape[data-shape="diamond"]').length,
    }
  })()`)

  record('changes selected PPT shape kind in inspector', afterShapeChange.shape === 'diamond' && afterShapeChange.thumbDiamondCount > 0, afterShapeChange)

  await pressKey(page, {
    code: 'KeyR',
    key: 'r',
    windowsVirtualKeyCode: 82,
  })
  await delay(20)

  const dragRect = await page.eval(`(() => {
    const slide = document.querySelector('.ppt-slide').getBoundingClientRect()

    return {
      endX: slide.left + slide.width * 0.58,
      endY: slide.top + slide.height * 0.38,
      pressed: document.querySelector('[data-ppt-insert-tool="rect"]')?.getAttribute('aria-pressed'),
      startX: slide.left + slide.width * 0.42,
      startY: slide.top + slide.height * 0.24,
    }
  })()`)

  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mousePressed',
    x: dragRect.startX,
    y: dragRect.startY,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    type: 'mouseMoved',
    x: dragRect.endX,
    y: dragRect.endY,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mouseReleased',
    x: dragRect.endX,
    y: dragRect.endY,
  })
  await delay(80)

  const afterDragRect = await page.eval(`(() => {
    const selected = document.querySelector('[data-selected="true"]')

    return {
      height: parseFloat(selected?.style.height ?? '0'),
      selectedKind: selected?.getAttribute('data-kind') ?? '',
      shape: selected?.getAttribute('data-shape') ?? null,
      width: parseFloat(selected?.style.width ?? '0'),
    }
  })()`)

  record('creates PPT rectangle from canvas shortcut drag', dragRect.pressed === 'true' && afterDragRect.selectedKind === 'shape' && afterDragRect.shape === 'rect' && afterDragRect.width > 140 && afterDragRect.height > 90, {
    afterDragRect,
    dragRect,
  })

  const beforeShapeColorSwatch = await getPPTColorSwatchState(page, 'shape-fill')

  await page.eval(`document.querySelector('[data-ppt-color-swatch="shape-fill"][data-ppt-color-token="ppt-color-accent"]')?.click()`)
  await delay(80)
  await page.eval(`document.querySelector('[data-ppt-color-swatch="shape-stroke"][data-ppt-color-token="ppt-color-background"]')?.click()`)
  await delay(80)

  const afterShapeColorSwatch = await getPPTColorSwatchState(page, 'shape-stroke')

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterShapeColorSwatchUndo = await getPPTColorSwatchState(page, 'shape-stroke')

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const afterShapeColorSwatchRedo = await getPPTColorSwatchState(page, 'shape-stroke')

  record(
    'applies PPT shape fill and stroke theme color swatches with undo redo',
    beforeShapeColorSwatch.themeCount >= 4 &&
      afterShapeColorSwatch.model === 'slide-edit-color-swatch-palette' &&
      afterShapeColorSwatch.descriptorModel === 'color-swatch-palette' &&
      afterShapeColorSwatch.descriptorCommand === 'apply-color-swatch' &&
      afterShapeColorSwatch.descriptorControl === 'color-swatch-palette' &&
      afterShapeColorSwatch.packageChannel === 'stroke' &&
      afterShapeColorSwatch.command === 'apply-color-swatch' &&
      afterShapeColorSwatch.commandChannel === 'stroke' &&
      afterShapeColorSwatch.commandObjects.includes(afterShapeColorSwatch.selectedId) &&
      afterShapeColorSwatch.commandSource === 'theme' &&
      afterShapeColorSwatch.commandSwatch === 'theme:ppt-color-background' &&
      afterShapeColorSwatch.commandToken === 'ppt-color-background' &&
      afterShapeColorSwatch.commandType === 'slide-command-effect' &&
      afterShapeColorSwatch.commandValue === '#f8fafc' &&
      afterShapeColorSwatch.fillValue === '#2563eb' &&
      afterShapeColorSwatch.strokeValue === '#f8fafc' &&
      afterShapeColorSwatch.background.includes('37, 99, 235') &&
      afterShapeColorSwatch.borderColor === 'rgb(248, 250, 252)' &&
      afterShapeColorSwatch.thumbBackground.includes('37, 99, 235') &&
      afterShapeColorSwatch.exportSlice.includes('"fill"') &&
      afterShapeColorSwatch.exportSlice.includes('"color": "#2563eb"') &&
      afterShapeColorSwatch.exportSlice.includes('"stroke"') &&
      afterShapeColorSwatch.exportSlice.includes('"color": "#f8fafc"') &&
      afterShapeColorSwatch.recentColors.includes('#2563eb') &&
      afterShapeColorSwatch.recentAccentCount === 1 &&
      afterShapeColorSwatch.recentUnique &&
      afterShapeColorSwatchUndo.strokeValue === '#6366f1' &&
      afterShapeColorSwatchUndo.fillValue === '#2563eb' &&
      afterShapeColorSwatchRedo.strokeValue === '#f8fafc',
    {
      afterShapeColorSwatch,
      afterShapeColorSwatchRedo,
      afterShapeColorSwatchUndo,
      beforeShapeColorSwatch,
    },
  )

  await page.eval(`(() => {
    const dash = document.querySelector('[data-ppt-style-field="stroke-dash"]')
    dash.value = 'dash'
    dash.dispatchEvent(new Event('change', { bubbles: true }))
  })()`)
  await delay(80)

  const afterShapeDash = await getPPTShapeStrokeDashState(page)

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterShapeDashUndo = await getPPTShapeStrokeDashState(page)

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const afterShapeDashRedo = await getPPTShapeStrokeDashState(page)

  record(
    'updates and restores PPT shape outline dash style from inspector',
    afterShapeDash.inspectorDash === 'dash' &&
      afterShapeDash.descriptorSurface === 'object-stroke-line-style' &&
      afterShapeDash.descriptorCommand === 'update-object-stroke-line-style' &&
      afterShapeDash.descriptorControl === 'stroke-line-style-segmented-control' &&
      afterShapeDash.descriptorAttribute === 'data-slide-object-stroke-line-style' &&
      afterShapeDash.descriptorAttributeValue === 'dash' &&
      afterShapeDash.model === 'slide-edit-object-stroke-line-style' &&
      afterShapeDash.command === 'update-object-stroke-line-style' &&
      afterShapeDash.commandField === 'strokeLineStyle' &&
      afterShapeDash.commandObject === afterShapeDash.selectedId &&
      afterShapeDash.commandSlide === 'slide-1' &&
      afterShapeDash.commandType === 'slide-command-effect' &&
      afterShapeDash.commandValue === 'dash' &&
      afterShapeDash.selectedDash === 'dash' &&
      afterShapeDash.selectedBorderStyle === 'dashed' &&
      afterShapeDash.thumbDash === 'dash' &&
      afterShapeDash.thumbBorderStyle === 'dashed' &&
      afterShapeDashUndo.inspectorDash === 'solid' &&
      afterShapeDashUndo.selectedDash === 'solid' &&
      afterShapeDashRedo.inspectorDash === 'dash' &&
      afterShapeDashRedo.selectedDash === 'dash',
    {
      afterShapeDash,
      afterShapeDashRedo,
      afterShapeDashUndo,
    },
  )

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      objectStrokeLineStyle: { value: 'dot' },
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const afterShapeDashJSONPaste = await getPPTShapeStrokeDashState(page)

  record(
    'pastes JSON object stroke line style through slide-edit command effect',
    afterShapeDashJSONPaste.importModel === 'ppt-object-stroke-line-style-import' &&
      afterShapeDashJSONPaste.importFormat === 'application-json-ppt-object-stroke-line-style' &&
      afterShapeDashJSONPaste.importSlide === 'slide-1' &&
      afterShapeDashJSONPaste.importObjects === afterShapeDashJSONPaste.selectedId &&
      afterShapeDashJSONPaste.importFields === 'strokeLineStyle' &&
      afterShapeDashJSONPaste.importCommands === 'update-object-stroke-line-style' &&
      afterShapeDashJSONPaste.importCommandFields === 'strokeLineStyle' &&
      afterShapeDashJSONPaste.importCommandTargets === afterShapeDashJSONPaste.selectedId &&
      afterShapeDashJSONPaste.importCommandTypes === 'slide-command-effect' &&
      afterShapeDashJSONPaste.importCommandValues === 'dot' &&
      afterShapeDashJSONPaste.importValue === 'dot' &&
      afterShapeDashJSONPaste.importJsonLength > 40 &&
      afterShapeDashJSONPaste.command === 'update-object-stroke-line-style' &&
      afterShapeDashJSONPaste.commandField === 'strokeLineStyle' &&
      afterShapeDashJSONPaste.commandValue === 'dot' &&
      afterShapeDashJSONPaste.inspectorDash === 'dot' &&
      afterShapeDashJSONPaste.selectedDash === 'dot' &&
      afterShapeDashJSONPaste.selectedBorderStyle === 'dotted' &&
      afterShapeDashJSONPaste.thumbDash === 'dot' &&
      afterShapeDashJSONPaste.thumbBorderStyle === 'dotted',
    {
      afterShapeDashJSONPaste,
      afterShapeDashRedo,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterShapeDashJSONUndo = await getPPTShapeStrokeDashState(page)

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const afterShapeDashJSONRedo = await getPPTShapeStrokeDashState(page)

  record(
    'undoes and redoes PPT object stroke line style JSON as one history step',
    afterShapeDashJSONUndo.inspectorDash === 'dash' &&
      afterShapeDashJSONUndo.selectedDash === 'dash' &&
      afterShapeDashJSONRedo.inspectorDash === 'dot' &&
      afterShapeDashJSONRedo.selectedDash === 'dot',
    {
      afterShapeDashJSONPaste,
      afterShapeDashJSONRedo,
      afterShapeDashJSONUndo,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  await page.eval(`(() => {
    const input = document.querySelector('[data-ppt-style-field="fill-opacity"]')
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set

    setter.call(input, '0.35')
    input.dispatchEvent(new Event('input', { bubbles: true }))
    input.dispatchEvent(new Event('change', { bubbles: true }))
  })()`)
  await delay(80)

  const afterFillOpacity = await getPPTShapeFillOpacityState(page)

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterFillOpacityUndo = await getPPTShapeFillOpacityState(page)

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const afterFillOpacityRedo = await getPPTShapeFillOpacityState(page)

  record(
    'updates and restores PPT shape fill opacity without fading object stroke/text',
    afterFillOpacity.inspectorOpacity === '0.35' &&
      afterFillOpacity.descriptorSurface === 'object-fill-opacity' &&
      afterFillOpacity.descriptorCommand === 'update-object-fill-opacity' &&
      afterFillOpacity.descriptorControl === 'fill-opacity-slider' &&
      afterFillOpacity.descriptorAttribute === 'data-slide-object-fill-opacity' &&
      afterFillOpacity.descriptorAttributeValue === '0.35' &&
      afterFillOpacity.model === 'slide-edit-object-fill-opacity' &&
      afterFillOpacity.command === 'update-object-fill-opacity' &&
      afterFillOpacity.commandField === 'fillOpacity' &&
      afterFillOpacity.commandObject === afterFillOpacity.selectedId &&
      afterFillOpacity.commandSlide === 'slide-1' &&
      afterFillOpacity.commandType === 'slide-command-effect' &&
      afterFillOpacity.commandValue === '0.35' &&
      afterFillOpacity.selectedFillOpacity === '0.35' &&
      afterFillOpacity.selectedBackground.includes('0.35') &&
      afterFillOpacity.selectedObjectOpacity === '1' &&
      afterFillOpacity.selectedBorderStyle === 'dashed' &&
      afterFillOpacity.thumbFillOpacity === '0.35' &&
      afterFillOpacity.thumbBackground.includes('0.35') &&
      afterFillOpacityUndo.inspectorOpacity === '1' &&
      afterFillOpacityUndo.selectedFillOpacity === '1' &&
      afterFillOpacityRedo.inspectorOpacity === '0.35' &&
      afterFillOpacityRedo.selectedFillOpacity === '0.35',
    {
      afterFillOpacity,
      afterFillOpacityRedo,
      afterFillOpacityUndo,
    },
  )

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      objectFillOpacity: { value: '0.62' },
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const afterFillOpacityJSONPaste = await getPPTShapeFillOpacityState(page)

  record(
    'pastes JSON object fill opacity through slide-edit command effect',
    afterFillOpacityJSONPaste.importModel === 'ppt-object-fill-opacity-import' &&
      afterFillOpacityJSONPaste.importFormat === 'application-json-ppt-object-fill-opacity' &&
      afterFillOpacityJSONPaste.importSlide === 'slide-1' &&
      afterFillOpacityJSONPaste.importObjects === afterFillOpacityJSONPaste.selectedId &&
      afterFillOpacityJSONPaste.importFields === 'fillOpacity' &&
      afterFillOpacityJSONPaste.importCommands === 'update-object-fill-opacity' &&
      afterFillOpacityJSONPaste.importCommandFields === 'fillOpacity' &&
      afterFillOpacityJSONPaste.importCommandTargets === afterFillOpacityJSONPaste.selectedId &&
      afterFillOpacityJSONPaste.importCommandTypes === 'slide-command-effect' &&
      afterFillOpacityJSONPaste.importCommandValues === '0.62' &&
      afterFillOpacityJSONPaste.importValue === '0.62' &&
      afterFillOpacityJSONPaste.importJsonLength > 35 &&
      afterFillOpacityJSONPaste.command === 'update-object-fill-opacity' &&
      afterFillOpacityJSONPaste.commandField === 'fillOpacity' &&
      afterFillOpacityJSONPaste.commandValue === '0.62' &&
      afterFillOpacityJSONPaste.inspectorOpacity === '0.62' &&
      afterFillOpacityJSONPaste.selectedFillOpacity === '0.62' &&
      afterFillOpacityJSONPaste.selectedBackground.includes('0.62') &&
      afterFillOpacityJSONPaste.selectedObjectOpacity === '1' &&
      afterFillOpacityJSONPaste.thumbFillOpacity === '0.62' &&
      afterFillOpacityJSONPaste.thumbBackground.includes('0.62'),
    {
      afterFillOpacityJSONPaste,
      afterFillOpacityRedo,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterFillOpacityJSONUndo = await getPPTShapeFillOpacityState(page)

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const afterFillOpacityJSONRedo = await getPPTShapeFillOpacityState(page)

  record(
    'undoes and redoes PPT object fill opacity JSON as one history step',
    afterFillOpacityJSONUndo.inspectorOpacity === '0.35' &&
      afterFillOpacityJSONUndo.selectedFillOpacity === '0.35' &&
      afterFillOpacityJSONRedo.inspectorOpacity === '0.62' &&
      afterFillOpacityJSONRedo.selectedFillOpacity === '0.62',
    {
      afterFillOpacityJSONPaste,
      afterFillOpacityJSONRedo,
      afterFillOpacityJSONUndo,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  await page.eval(`(() => {
    const input = document.querySelector('[data-ppt-style-field="shape-corner-radius"]')
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set

    setter.call(input, '36')
    input.dispatchEvent(new Event('input', { bubbles: true }))
    input.dispatchEvent(new Event('change', { bubbles: true }))
  })()`)
  await delay(80)

  const afterCornerRadius = await getPPTShapeCornerRadiusState(page)

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterCornerRadiusUndo = await getPPTShapeCornerRadiusState(page)

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const afterCornerRadiusRedo = await getPPTShapeCornerRadiusState(page)

  record(
    'updates and restores PPT shape corner radius from inspector',
    afterCornerRadius.inspectorRadius === '36' &&
      afterCornerRadius.descriptorSurface === 'object-corner-radius' &&
      afterCornerRadius.descriptorCommand === 'update-object-corner-radius' &&
      afterCornerRadius.descriptorControl === 'corner-radius-slider' &&
      afterCornerRadius.descriptorAttribute === 'data-slide-object-corner-radius' &&
      afterCornerRadius.descriptorAttributeValue === '36' &&
      afterCornerRadius.descriptorSupported === 'true' &&
      afterCornerRadius.model === 'slide-edit-object-corner-radius' &&
      afterCornerRadius.command === 'update-object-corner-radius' &&
      afterCornerRadius.commandField === 'cornerRadius' &&
      afterCornerRadius.commandObject === afterCornerRadius.selectedId &&
      afterCornerRadius.commandSlide === 'slide-1' &&
      afterCornerRadius.commandType === 'slide-command-effect' &&
      afterCornerRadius.commandValue === '36' &&
      afterCornerRadius.selectedCornerRadius === '36' &&
      afterCornerRadius.selectedBorderRadius === '36px' &&
      afterCornerRadius.thumbCornerRadius === '36' &&
      afterCornerRadius.thumbBorderRadius !== '' &&
      afterCornerRadiusUndo.inspectorRadius === '24' &&
      afterCornerRadiusUndo.selectedCornerRadius === '24' &&
      afterCornerRadiusRedo.inspectorRadius === '36' &&
      afterCornerRadiusRedo.selectedCornerRadius === '36',
    {
      afterCornerRadius,
      afterCornerRadiusRedo,
      afterCornerRadiusUndo,
    },
  )

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      objectCornerRadius: { value: '18' },
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const afterCornerRadiusJSONPaste = await getPPTShapeCornerRadiusState(page)

  record(
    'pastes JSON object corner radius through slide-edit command effect',
    afterCornerRadiusJSONPaste.importModel === 'ppt-object-corner-radius-import' &&
      afterCornerRadiusJSONPaste.importFormat === 'application-json-ppt-object-corner-radius' &&
      afterCornerRadiusJSONPaste.importSlide === 'slide-1' &&
      afterCornerRadiusJSONPaste.importObjects === afterCornerRadiusJSONPaste.selectedId &&
      afterCornerRadiusJSONPaste.importFields === 'cornerRadius' &&
      afterCornerRadiusJSONPaste.importCommands === 'update-object-corner-radius' &&
      afterCornerRadiusJSONPaste.importCommandFields === 'cornerRadius' &&
      afterCornerRadiusJSONPaste.importCommandTargets === afterCornerRadiusJSONPaste.selectedId &&
      afterCornerRadiusJSONPaste.importCommandTypes === 'slide-command-effect' &&
      afterCornerRadiusJSONPaste.importCommandValues === '18' &&
      afterCornerRadiusJSONPaste.importValue === '18' &&
      afterCornerRadiusJSONPaste.importJsonLength > 35 &&
      afterCornerRadiusJSONPaste.command === 'update-object-corner-radius' &&
      afterCornerRadiusJSONPaste.commandField === 'cornerRadius' &&
      afterCornerRadiusJSONPaste.commandValue === '18' &&
      afterCornerRadiusJSONPaste.inspectorRadius === '18' &&
      afterCornerRadiusJSONPaste.selectedCornerRadius === '18' &&
      afterCornerRadiusJSONPaste.selectedBorderRadius === '18px' &&
      afterCornerRadiusJSONPaste.thumbCornerRadius === '18' &&
      afterCornerRadiusJSONPaste.thumbBorderRadius !== '',
    {
      afterCornerRadiusJSONPaste,
      afterCornerRadiusRedo,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterCornerRadiusJSONUndo = await getPPTShapeCornerRadiusState(page)

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const afterCornerRadiusJSONRedo = await getPPTShapeCornerRadiusState(page)

  record(
    'undoes and redoes PPT object corner radius JSON as one history step',
    afterCornerRadiusJSONUndo.inspectorRadius === '36' &&
      afterCornerRadiusJSONUndo.selectedCornerRadius === '36' &&
      afterCornerRadiusJSONRedo.inspectorRadius === '18' &&
      afterCornerRadiusJSONRedo.selectedCornerRadius === '18',
    {
      afterCornerRadiusJSONPaste,
      afterCornerRadiusJSONRedo,
      afterCornerRadiusJSONUndo,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  await page.eval(`document.querySelector('[data-ppt-present-start]')?.click()`)
  await delay(120)

  const fillPreview = await page.eval(`((id) => {
    const overlay = document.querySelector('[data-ppt-presentation]')
    const element = overlay?.querySelector(\`[data-ppt-element="\${id}"]\`)

    return {
      background: element?.style.background ?? '',
      borderRadius: element?.style.borderRadius ?? '',
      cornerRadius: element?.getAttribute('data-ppt-corner-radius') ?? '',
      fillOpacity: element?.getAttribute('data-ppt-fill-opacity') ?? '',
      objectOpacity: element?.style.opacity ?? '',
      open: !!overlay,
    }
  })(${JSON.stringify(afterFillOpacity.selectedId)})`)

  record(
    'keeps PPT shape fill opacity in presentation preview',
    fillPreview.open &&
      fillPreview.fillOpacity === '0.35' &&
      fillPreview.background.includes('0.35') &&
      fillPreview.objectOpacity === '1',
    fillPreview,
  )

  record(
    'keeps PPT shape corner radius in presentation preview',
    fillPreview.open &&
      fillPreview.cornerRadius === '36' &&
      fillPreview.borderRadius === '36px',
    fillPreview,
  )

  await page.eval(`document.querySelector('[data-ppt-present-exit]')?.click()`)
  await delay(80)

  await page.eval(`(() => {
    const opacity = document.querySelector('[data-ppt-style-field="opacity"]')
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set

    setter.call(opacity, '0.42')
    opacity.dispatchEvent(new Event('input', { bubbles: true }))
    opacity.dispatchEvent(new Event('change', { bubbles: true }))
  })()`)
  await delay(80)

  await page.eval(`document.querySelector('[data-ppt-shadow-field="enabled"]')?.click()`)
  await delay(80)

  await page.eval(`document.querySelector('[data-ppt-command="copy-formatting"]')?.click()`)
  await delay(80)

  const afterCopyFormatting = await page.eval(`(() => {
    const shell = document.querySelector('.ppt-stage-shell')

    return {
      categories: shell?.getAttribute('data-ppt-style-clipboard-categories') ?? '',
      command: shell?.getAttribute('data-ppt-style-clipboard-command') ?? '',
      commandSelection: shell?.getAttribute('data-ppt-style-clipboard-command-selection') ?? '',
      commandSlide: shell?.getAttribute('data-ppt-style-clipboard-command-slide') ?? '',
      commandSourceId: shell?.getAttribute('data-ppt-style-clipboard-command-source-id') ?? '',
      commandType: shell?.getAttribute('data-ppt-style-clipboard-command-type') ?? '',
      copyDisabled: document.querySelector('[data-ppt-command="copy-formatting"]')?.disabled ?? true,
      model: shell?.getAttribute('data-ppt-style-clipboard-model') ?? '',
      packageCategories: shell?.getAttribute('data-ppt-style-clipboard-package-categories') ?? '',
      pasteDisabled: document.querySelector('[data-ppt-command="paste-formatting"]')?.disabled ?? true,
      sourceId: shell?.getAttribute('data-ppt-style-clipboard-source-id') ?? '',
      sourceKind: shell?.getAttribute('data-ppt-style-clipboard-source-kind') ?? '',
      supportedTargets: shell?.getAttribute('data-ppt-style-clipboard-supported-targets') ?? '',
      type: shell?.getAttribute('data-ppt-style-clipboard-type') ?? '',
    }
  })()`)

  record(
    'copies PPT formatting into style clipboard from toolbar',
    !afterCopyFormatting.copyDisabled &&
      !afterCopyFormatting.pasteDisabled &&
      afterCopyFormatting.type === 'slide-style-clipboard' &&
      afterCopyFormatting.model === 'slide-edit-style-clipboard' &&
      afterCopyFormatting.sourceId === afterCornerRadius.selectedId &&
      afterCopyFormatting.sourceKind === 'shape' &&
      afterCopyFormatting.command === 'copy-object-formatting' &&
      afterCopyFormatting.commandSelection === afterCornerRadius.selectedId &&
      afterCopyFormatting.commandSlide === 'slide-1' &&
      afterCopyFormatting.commandSourceId === afterCornerRadius.selectedId &&
      afterCopyFormatting.commandType === 'slide-command-effect' &&
      afterCopyFormatting.categories.includes('shape') &&
      afterCopyFormatting.categories.includes('stroke') &&
      afterCopyFormatting.packageCategories.includes('object-effect') &&
      afterCopyFormatting.packageCategories.includes('shape-fill') &&
      afterCopyFormatting.packageCategories.includes('shape-stroke') &&
      afterCopyFormatting.packageCategories.includes('line-style'),
    afterCopyFormatting,
  )

  await pressKey(page, {
    code: 'KeyR',
    key: 'r',
    windowsVirtualKeyCode: 82,
  })
  await delay(20)

  const firstFormatTargetDrag = await page.eval(`(() => {
    const slide = document.querySelector('.ppt-slide').getBoundingClientRect()

    return {
      endX: slide.left + slide.width * 0.82,
      endY: slide.top + slide.height * 0.28,
      startX: slide.left + slide.width * 0.66,
      startY: slide.top + slide.height * 0.14,
    }
  })()`)

  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mousePressed',
    x: firstFormatTargetDrag.startX,
    y: firstFormatTargetDrag.startY,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    type: 'mouseMoved',
    x: firstFormatTargetDrag.endX,
    y: firstFormatTargetDrag.endY,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mouseReleased',
    x: firstFormatTargetDrag.endX,
    y: firstFormatTargetDrag.endY,
  })
  await delay(100)

  const beforeShortcutPasteFormatting = await getPPTFormatPainterSelectedShapeState(page)

  await pressKey(page, {
    code: 'KeyV',
    key: 'v',
    modifiers: 10,
    windowsVirtualKeyCode: 86,
  })
  await delay(100)

  const afterShortcutPasteFormatting = await getPPTFormatPainterSelectedShapeState(page)

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterShortcutPasteFormattingUndo = await getPPTFormatPainterSelectedShapeState(page)

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const afterShortcutPasteFormattingRedo = await getPPTFormatPainterSelectedShapeState(page)

  record(
    'pastes PPT shape formatting from keyboard without changing content geometry or name',
    beforeShortcutPasteFormatting.selectedKind === 'shape' &&
      beforeShortcutPasteFormatting.shape === 'rect' &&
      beforeShortcutPasteFormatting.fillOpacity === '1' &&
      beforeShortcutPasteFormatting.objectOpacity === '1' &&
      beforeShortcutPasteFormatting.shadow === '' &&
      beforeShortcutPasteFormatting.borderStyle === 'solid' &&
      afterShortcutPasteFormatting.selectedId === beforeShortcutPasteFormatting.selectedId &&
      afterShortcutPasteFormatting.name === beforeShortcutPasteFormatting.name &&
      afterShortcutPasteFormatting.left === beforeShortcutPasteFormatting.left &&
      afterShortcutPasteFormatting.top === beforeShortcutPasteFormatting.top &&
      afterShortcutPasteFormatting.width === beforeShortcutPasteFormatting.width &&
      afterShortcutPasteFormatting.height === beforeShortcutPasteFormatting.height &&
      afterShortcutPasteFormatting.fillOpacity === '0.35' &&
      afterShortcutPasteFormatting.background.includes('0.35') &&
      afterShortcutPasteFormatting.objectOpacity === '0.42' &&
      afterShortcutPasteFormatting.styleOpacity === '0.42' &&
      afterShortcutPasteFormatting.shadow === 'true' &&
      afterShortcutPasteFormatting.shadowOpacity === '0.22' &&
      afterShortcutPasteFormatting.filter.includes('drop-shadow') &&
      afterShortcutPasteFormatting.borderStyle === 'dashed' &&
      afterShortcutPasteFormatting.strokeDash === 'dash' &&
      afterShortcutPasteFormatting.cornerRadius === '36' &&
      afterShortcutPasteFormatting.borderRadius === '36px' &&
      afterShortcutPasteFormatting.styleClipboardCommand === 'paste-object-formatting' &&
      afterShortcutPasteFormatting.styleClipboardCommandApplications.includes(afterShortcutPasteFormatting.selectedId) &&
      afterShortcutPasteFormatting.styleClipboardCommandApplications.includes('shape-fill') &&
      afterShortcutPasteFormatting.styleClipboardCommandApplications.includes('object-effect') &&
      afterShortcutPasteFormatting.styleClipboardCommandSelection.split(' ').includes(afterShortcutPasteFormatting.selectedId) &&
      afterShortcutPasteFormatting.styleClipboardCommandTargets.split(' ').includes(afterShortcutPasteFormatting.selectedId) &&
      afterShortcutPasteFormattingUndo.fillOpacity === '1' &&
      afterShortcutPasteFormattingUndo.objectOpacity === '1' &&
      afterShortcutPasteFormattingUndo.shadow === '' &&
      afterShortcutPasteFormattingUndo.cornerRadius === '24' &&
      afterShortcutPasteFormattingRedo.fillOpacity === '0.35' &&
      afterShortcutPasteFormattingRedo.objectOpacity === '0.42' &&
      afterShortcutPasteFormattingRedo.shadow === 'true' &&
      afterShortcutPasteFormattingRedo.cornerRadius === '36',
    {
      afterShortcutPasteFormatting,
      afterShortcutPasteFormattingRedo,
      afterShortcutPasteFormattingUndo,
      beforeShortcutPasteFormatting,
    },
  )

  await pressKey(page, {
    code: 'KeyR',
    key: 'r',
    windowsVirtualKeyCode: 82,
  })
  await delay(20)

  const secondFormatTargetDrag = await page.eval(`(() => {
    const slide = document.querySelector('.ppt-slide').getBoundingClientRect()

    return {
      endX: slide.left + slide.width * 0.82,
      endY: slide.top + slide.height * 0.48,
      startX: slide.left + slide.width * 0.66,
      startY: slide.top + slide.height * 0.34,
    }
  })()`)

  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mousePressed',
    x: secondFormatTargetDrag.startX,
    y: secondFormatTargetDrag.startY,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    type: 'mouseMoved',
    x: secondFormatTargetDrag.endX,
    y: secondFormatTargetDrag.endY,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mouseReleased',
    x: secondFormatTargetDrag.endX,
    y: secondFormatTargetDrag.endY,
  })
  await delay(100)

  await pressKey(page, {
    code: 'KeyK',
    key: 'k',
    modifiers: 2,
    windowsVirtualKeyCode: 75,
  })
  await delay(80)

  await page.send('Input.insertText', { text: 'paste formatting' })
  await delay(80)

  const beforePalettePasteFormatting = await page.eval(`(() => ({
    disabled: document.querySelector('[data-ppt-command-palette-item="command:paste-formatting"]')?.disabled ?? true,
    itemPresent: !!document.querySelector('[data-ppt-command-palette-item="command:paste-formatting"]'),
    open: !!document.querySelector('[data-ppt-command-palette]'),
  }))()`)

  await pressKey(page, {
    code: 'Enter',
    key: 'Enter',
    windowsVirtualKeyCode: 13,
  })
  await delay(100)

  const afterPalettePasteFormatting = await getPPTFormatPainterSelectedShapeState(page)

  record(
    'pastes PPT shape formatting from command palette',
    beforePalettePasteFormatting.open &&
      beforePalettePasteFormatting.itemPresent &&
      !beforePalettePasteFormatting.disabled &&
      !afterPalettePasteFormatting.paletteOpen &&
      afterPalettePasteFormatting.fillOpacity === '0.35' &&
      afterPalettePasteFormatting.objectOpacity === '0.42' &&
      afterPalettePasteFormatting.shadow === 'true' &&
      afterPalettePasteFormatting.strokeDash === 'dash' &&
      afterPalettePasteFormatting.cornerRadius === '36',
    {
      afterPalettePasteFormatting,
      beforePalettePasteFormatting,
    },
  )

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      shapeStyle: {
        cornerRadius: 18,
        fill: {
          color: '#ef4444',
          opacity: 0.62,
        },
        stroke: {
          color: '#111827',
          dash: 'dot',
          width: 6,
        },
      },
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const afterShapeStylePaste = await getPPTFormatPainterSelectedShapeState(page)

  record(
    'pastes JSON shape style into selected PPT shape',
    afterShapeStylePaste.shapeStyleImportModel === 'ppt-shape-style-import' &&
      afterShapeStylePaste.shapeStyleImportFormat === 'application-json-ppt-shape-style' &&
      afterShapeStylePaste.shapeStyleImportCommand === 'paste-object-formatting' &&
      afterShapeStylePaste.shapeStyleImportCommandTargets === afterShapeStylePaste.selectedId &&
      afterShapeStylePaste.shapeStyleImportCommandType === 'slide-command-effect' &&
      afterShapeStylePaste.shapeStyleImportObjects === afterShapeStylePaste.selectedId &&
      afterShapeStylePaste.shapeStyleImportCategories.includes('object-effect') &&
      afterShapeStylePaste.shapeStyleImportCategories.includes('shape-fill') &&
      afterShapeStylePaste.shapeStyleImportCategories.includes('shape-stroke') &&
      afterShapeStylePaste.shapeStyleImportCategories.includes('line-style') &&
      afterShapeStylePaste.shapeStyleImportFields === 'fill stroke cornerRadius' &&
      afterShapeStylePaste.shapeStyleImportFillColor === '#ef4444' &&
      afterShapeStylePaste.shapeStyleImportFillOpacity === '0.62' &&
      afterShapeStylePaste.shapeStyleImportStrokeColor === '#111827' &&
      afterShapeStylePaste.shapeStyleImportStrokeDash === 'dot' &&
      afterShapeStylePaste.shapeStyleImportStrokeWidth === '6' &&
      afterShapeStylePaste.shapeStyleImportCornerRadius === '18' &&
      Number(afterShapeStylePaste.shapeStyleImportJsonLength) > 100 &&
      afterShapeStylePaste.fillOpacity === '0.62' &&
      afterShapeStylePaste.background.includes('0.62') &&
      afterShapeStylePaste.borderColor === 'rgb(17, 24, 39)' &&
      afterShapeStylePaste.borderStyle === 'dotted' &&
      afterShapeStylePaste.borderWidth === '6px' &&
      afterShapeStylePaste.strokeDash === 'dot' &&
      afterShapeStylePaste.cornerRadius === '18' &&
      afterShapeStylePaste.borderRadius === '18px' &&
      afterShapeStylePaste.objectOpacity === '0.42' &&
      afterShapeStylePaste.shadow === 'true' &&
      afterShapeStylePaste.styleClipboardCommand === 'paste-object-formatting' &&
      afterShapeStylePaste.styleClipboardCommandApplications.includes(afterShapeStylePaste.selectedId) &&
      afterShapeStylePaste.styleClipboardCommandApplications.includes('shape-fill') &&
      afterShapeStylePaste.styleClipboardCommandApplications.includes('shape-stroke') &&
      afterShapeStylePaste.styleClipboardCommandApplications.includes('line-style'),
    {
      afterPalettePasteFormatting,
      afterShapeStylePaste,
    },
  )

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      cornerRadius: 12,
      fillColor: '#22c55e',
      fillOpacity: 0.48,
      strokeColor: '#1f2937',
      strokeDash: 'dash',
      strokeWidth: 4,
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const afterStandaloneShapeStylePaste = await getPPTFormatPainterSelectedShapeState(page)

  record(
    'pastes standalone PPT shape style fields through style clipboard categories',
    afterStandaloneShapeStylePaste.shapeStyleImportModel === 'ppt-shape-style-import' &&
      afterStandaloneShapeStylePaste.shapeStyleImportFormat === 'application-json-ppt-shape-style' &&
      afterStandaloneShapeStylePaste.shapeStyleImportCommand === 'paste-object-formatting' &&
      afterStandaloneShapeStylePaste.shapeStyleImportCommandTargets === afterStandaloneShapeStylePaste.selectedId &&
      afterStandaloneShapeStylePaste.shapeStyleImportCommandType === 'slide-command-effect' &&
      afterStandaloneShapeStylePaste.shapeStyleImportObjects === afterStandaloneShapeStylePaste.selectedId &&
      afterStandaloneShapeStylePaste.shapeStyleImportCategories.includes('object-effect') &&
      afterStandaloneShapeStylePaste.shapeStyleImportCategories.includes('shape-fill') &&
      afterStandaloneShapeStylePaste.shapeStyleImportCategories.includes('shape-stroke') &&
      afterStandaloneShapeStylePaste.shapeStyleImportCategories.includes('line-style') &&
      afterStandaloneShapeStylePaste.shapeStyleImportFields === 'fill stroke cornerRadius' &&
      afterStandaloneShapeStylePaste.shapeStyleImportFillColor === '#22c55e' &&
      afterStandaloneShapeStylePaste.shapeStyleImportFillOpacity === '0.48' &&
      afterStandaloneShapeStylePaste.shapeStyleImportStrokeColor === '#1f2937' &&
      afterStandaloneShapeStylePaste.shapeStyleImportStrokeDash === 'dash' &&
      afterStandaloneShapeStylePaste.shapeStyleImportStrokeWidth === '4' &&
      afterStandaloneShapeStylePaste.shapeStyleImportCornerRadius === '12' &&
      Number(afterStandaloneShapeStylePaste.shapeStyleImportJsonLength) > 100 &&
      afterStandaloneShapeStylePaste.fillOpacity === '0.48' &&
      afterStandaloneShapeStylePaste.background.includes('0.48') &&
      afterStandaloneShapeStylePaste.borderColor === 'rgb(31, 41, 55)' &&
      afterStandaloneShapeStylePaste.borderStyle === 'dashed' &&
      afterStandaloneShapeStylePaste.borderWidth === '4px' &&
      afterStandaloneShapeStylePaste.strokeDash === 'dash' &&
      afterStandaloneShapeStylePaste.cornerRadius === '12' &&
      afterStandaloneShapeStylePaste.borderRadius === '12px' &&
      afterStandaloneShapeStylePaste.objectOpacity === afterShapeStylePaste.objectOpacity &&
      afterStandaloneShapeStylePaste.shadow === afterShapeStylePaste.shadow &&
      afterStandaloneShapeStylePaste.styleClipboardCommand === 'paste-object-formatting' &&
      afterStandaloneShapeStylePaste.styleClipboardCommandApplications.includes(afterStandaloneShapeStylePaste.selectedId) &&
      afterStandaloneShapeStylePaste.styleClipboardCommandApplications.includes('shape-fill') &&
      afterStandaloneShapeStylePaste.styleClipboardCommandApplications.includes('shape-stroke') &&
      afterStandaloneShapeStylePaste.styleClipboardCommandApplications.includes('line-style'),
    {
      afterShapeStylePaste,
      afterStandaloneShapeStylePaste,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterStandaloneShapeStyleUndo = await getPPTFormatPainterSelectedShapeState(page)

  record(
    'undoes standalone PPT shape style fields as one history step',
    afterStandaloneShapeStyleUndo.selectedId === afterShapeStylePaste.selectedId &&
      afterStandaloneShapeStyleUndo.fillOpacity === afterShapeStylePaste.fillOpacity &&
      afterStandaloneShapeStyleUndo.borderColor === afterShapeStylePaste.borderColor &&
      afterStandaloneShapeStyleUndo.borderStyle === afterShapeStylePaste.borderStyle &&
      afterStandaloneShapeStyleUndo.borderWidth === afterShapeStylePaste.borderWidth &&
      afterStandaloneShapeStyleUndo.strokeDash === afterShapeStylePaste.strokeDash &&
      afterStandaloneShapeStyleUndo.cornerRadius === afterShapeStylePaste.cornerRadius,
    {
      afterShapeStylePaste,
      afterStandaloneShapeStylePaste,
      afterStandaloneShapeStyleUndo,
    },
  )

  await pressKey(page, {
    code: 'KeyT',
    key: 't',
    windowsVirtualKeyCode: 84,
  })
  await delay(20)

  const createText = await page.eval(`(() => {
    const slide = document.querySelector('.ppt-slide').getBoundingClientRect()

    return {
      pressed: document.querySelector('[data-ppt-insert-tool="text"]')?.getAttribute('aria-pressed'),
      x: slide.left + slide.width * 0.22,
      y: slide.top + slide.height * 0.68,
    }
  })()`)

  await clickMouse(page, createText.x, createText.y, 1)
  await delay(100)

  const afterCreateText = await page.eval(`(() => {
    const selected = document.querySelector('[data-selected="true"]')
    const editor = selected?.querySelector('.ppt-element-editor')

    return {
      editing: editor?.isContentEditable === true && document.activeElement === editor,
      selectedKind: selected?.getAttribute('data-kind') ?? '',
      text: selected?.textContent ?? '',
    }
  })()`)

  record('creates PPT text from canvas shortcut click and enters inline edit', createText.pressed === 'true' && afterCreateText.selectedKind === 'textBox' && afterCreateText.editing && afterCreateText.text.includes('New text'), {
    afterCreateText,
    createText,
  })

  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)

  await page.eval(`(() => {
    const background = document.querySelector('[data-ppt-slide-field="background"]')
    const valueSetter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set
    valueSetter.call(background, '#fef3c7')
    background.dispatchEvent(new Event('input', { bubbles: true }))
    background.dispatchEvent(new Event('change', { bubbles: true }))
  })()`)
  await delay(50)

  const afterBackground = await page.eval(`(() => ({
    slideBackground: document.querySelector('.ppt-slide')?.style.background ?? '',
    thumbBackground: document.querySelector('.ppt-thumb[aria-current="page"] .ppt-thumb-preview')?.style.background ?? '',
  }))()`)

  record('updates PPT slide background in inspector', afterBackground.slideBackground === 'rgb(254, 243, 199)' && afterBackground.thumbBackground === 'rgb(254, 243, 199)', afterBackground)

  const beforeLayout = await page.eval(`(() => ({
    elementCount: document.querySelectorAll('[data-kind]').length,
    layout: document.querySelector('[data-ppt-slide-field="layout"]')?.value ?? '',
    placeholderCount: Number(document.querySelector('[data-ppt-layout-placeholder-count]')?.getAttribute('data-ppt-layout-placeholder-count') ?? 0),
    slideLayout: document.querySelector('.ppt-slide')?.getAttribute('data-ppt-layout-id') ?? '',
    slideTheme: document.querySelector('.ppt-slide')?.getAttribute('data-ppt-theme-id') ?? '',
    themeTokenCount: document.querySelectorAll('[data-ppt-theme-token]').length,
  }))()`)

  record(
    'renders PPT layout/theme token controls in inspector',
    beforeLayout.layout === 'ppt-layout-title-body' &&
      beforeLayout.slideLayout === 'ppt-layout-title-body' &&
      beforeLayout.slideTheme === 'ppt-theme-default' &&
      beforeLayout.themeTokenCount >= 4 &&
      beforeLayout.placeholderCount >= 2,
    beforeLayout,
  )

  await page.eval(`(() => {
    const layout = document.querySelector('[data-ppt-slide-field="layout"]')
    const valueSetter = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set
    valueSetter.call(layout, 'ppt-layout-split')
    layout.dispatchEvent(new Event('input', { bubbles: true }))
    layout.dispatchEvent(new Event('change', { bubbles: true }))
  })()`)
  await delay(50)

  const afterLayout = await page.eval(`(() => ({
    elementCount: document.querySelectorAll('[data-kind]').length,
    layout: document.querySelector('[data-ppt-slide-field="layout"]')?.value ?? '',
    placeholderCount: Number(document.querySelector('[data-ppt-layout-placeholder-count]')?.getAttribute('data-ppt-layout-placeholder-count') ?? 0),
    slideLayout: document.querySelector('.ppt-slide')?.getAttribute('data-ppt-layout-id') ?? '',
    slideTheme: document.querySelector('.ppt-slide')?.getAttribute('data-ppt-theme-id') ?? '',
  }))()`)

  record(
    'updates PPT slide layout through slide-edit layout/theme command effect',
    afterLayout.layout === 'ppt-layout-split' &&
      afterLayout.slideLayout === 'ppt-layout-split' &&
      afterLayout.slideTheme === 'ppt-theme-default' &&
      afterLayout.placeholderCount >= 3 &&
      afterLayout.elementCount === beforeLayout.elementCount,
    {
      afterLayout,
      beforeLayout,
    },
  )

  const placeholderSelectionId = await page.eval(`document.querySelector('.ppt-slide [data-ppt-element]')?.getAttribute('data-ppt-element') ?? ''`)

  if (placeholderSelectionId) {
    const point = await getElementCenter(page, placeholderSelectionId)
    await clickMouse(page, point.x, point.y, 1)
    await delay(50)
  }

  const beforePlaceholderVisibility = await getPPTPlaceholderVisibilityState(page)

  record(
    'renders PPT placeholder visibility metadata in inspector',
    beforePlaceholderVisibility.count >= 3 &&
      beforePlaceholderVisibility.placeholderId === 'media' &&
      beforePlaceholderVisibility.role === 'media' &&
      beforePlaceholderVisibility.bounds.length > 0 &&
      beforePlaceholderVisibility.layout === 'ppt-layout-split' &&
      beforePlaceholderVisibility.master === 'ppt-master-default' &&
      beforePlaceholderVisibility.locked === 'false' &&
      beforePlaceholderVisibility.visible === 'true' &&
      beforePlaceholderVisibility.toggleDisabled === false,
    beforePlaceholderVisibility,
  )

  await page.eval(`document.querySelector('[data-ppt-placeholder-visibility-toggle="media"]')?.click()`)
  await delay(80)

  const afterPlaceholderHide = await getPPTPlaceholderVisibilityState(page)

  record(
    'hides PPT layout placeholder from inspector',
    afterPlaceholderHide.visible === 'false' &&
      afterPlaceholderHide.hiddenCount === beforePlaceholderVisibility.hiddenCount + 1 &&
      afterPlaceholderHide.slideHiddenPlaceholders.split(' ').includes('media') &&
      afterPlaceholderHide.selectedIds === beforePlaceholderVisibility.selectedIds &&
      afterPlaceholderHide.command === 'update-placeholder-visibility' &&
      afterPlaceholderHide.commandPlaceholder === 'media' &&
      afterPlaceholderHide.commandSlide === beforePlaceholderVisibility.slideId &&
      afterPlaceholderHide.commandType === 'slide-command-effect' &&
      afterPlaceholderHide.commandVisible === 'false',
    {
      afterPlaceholderHide,
      beforePlaceholderVisibility,
    },
  )

  await page.eval(`document.querySelector('button[title="Undo"]')?.click()`)
  await delay(80)

  const afterPlaceholderUndo = await getPPTPlaceholderVisibilityState(page)

  await page.eval(`document.querySelector('button[title="Redo"]')?.click()`)
  await delay(80)

  const afterPlaceholderRedo = await getPPTPlaceholderVisibilityState(page)

  record(
    'undoes and redoes PPT placeholder visibility as one history step',
    afterPlaceholderUndo.visible === 'true' &&
      !afterPlaceholderUndo.slideHiddenPlaceholders.split(' ').includes('media') &&
      afterPlaceholderUndo.selectedIds === beforePlaceholderVisibility.selectedIds &&
      afterPlaceholderRedo.visible === 'false' &&
      afterPlaceholderRedo.slideHiddenPlaceholders.split(' ').includes('media') &&
      afterPlaceholderRedo.selectedIds === beforePlaceholderVisibility.selectedIds,
    {
      afterPlaceholderRedo,
      afterPlaceholderUndo,
      beforePlaceholderVisibility,
    },
  )

  await page.eval(`document.querySelector('[data-ppt-placeholder-visibility-toggle="media"]')?.click()`)
  await delay(80)

  const afterPlaceholderShow = await getPPTPlaceholderVisibilityState(page)

  record(
    'shows PPT layout placeholder from inspector',
    afterPlaceholderShow.visible === 'true' &&
      !afterPlaceholderShow.slideHiddenPlaceholders.split(' ').includes('media') &&
      afterPlaceholderShow.selectedIds === beforePlaceholderVisibility.selectedIds &&
      afterPlaceholderShow.command === 'update-placeholder-visibility' &&
      afterPlaceholderShow.commandPlaceholder === 'media' &&
      afterPlaceholderShow.commandVisible === 'true',
    {
      afterPlaceholderShow,
      beforePlaceholderVisibility,
    },
  )

  await page.eval(`document.querySelector('[data-ppt-placeholder-visibility-toggle="media"]')?.click()`)
  await delay(80)

  const afterPlaceholderRehide = await getPPTPlaceholderVisibilityState(page)

  record(
    'rehides PPT layout placeholder for export metadata',
    afterPlaceholderRehide.visible === 'false' &&
      afterPlaceholderRehide.slideHiddenPlaceholders.split(' ').includes('media') &&
      afterPlaceholderRehide.selectedIds === beforePlaceholderVisibility.selectedIds,
    {
      afterPlaceholderRehide,
      beforePlaceholderVisibility,
    },
  )

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      slideLayout: {
        layoutId: 'ppt-layout-title-body',
        themeId: 'ppt-theme-default',
        visiblePlaceholderIds: ['title'],
      },
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const afterSlideLayoutPaste = await getPPTSlideLayoutImportState(page, 'body')

  record(
    'pastes JSON slide layout and placeholder visibility into active PPT slide',
    afterSlideLayoutPaste.importModel === 'ppt-slide-layout-import' &&
      afterSlideLayoutPaste.importFormat === 'application-json-ppt-slide-layout' &&
      afterSlideLayoutPaste.importSlide === 'slide-1' &&
      afterSlideLayoutPaste.importFields === 'layoutId themeId hiddenPlaceholderIds' &&
      afterSlideLayoutPaste.importCommands ===
        'apply-layout update-placeholder-visibility update-placeholder-visibility' &&
      afterSlideLayoutPaste.commandFields === 'layoutId title body' &&
      afterSlideLayoutPaste.commandTypes ===
        'slide-command-effect slide-command-effect slide-command-effect' &&
      afterSlideLayoutPaste.importLayout === 'ppt-layout-title-body' &&
      afterSlideLayoutPaste.importTheme === 'ppt-theme-default' &&
      afterSlideLayoutPaste.importHiddenPlaceholders === 'body' &&
      afterSlideLayoutPaste.importPlaceholderCommandCount === 2 &&
      afterSlideLayoutPaste.importJsonLength > 100 &&
      afterSlideLayoutPaste.layout === 'ppt-layout-title-body' &&
      afterSlideLayoutPaste.theme === 'ppt-theme-default' &&
      afterSlideLayoutPaste.placeholderId === 'body' &&
      afterSlideLayoutPaste.visible === 'false' &&
      afterSlideLayoutPaste.hiddenCount === 1 &&
      afterSlideLayoutPaste.slideHiddenPlaceholders === 'body' &&
      afterSlideLayoutPaste.command === 'update-placeholder-visibility' &&
      afterSlideLayoutPaste.commandPlaceholder === 'body' &&
      afterSlideLayoutPaste.commandVisible === 'false',
    {
      afterPlaceholderRehide,
      afterSlideLayoutPaste,
    },
  )

  await page.eval(`document.querySelector('button[title="Undo"]')?.click()`)
  await delay(100)

  const afterSlideLayoutPasteUndo = await getPPTPlaceholderVisibilityState(page)

  record(
    'undoes PPT slide layout JSON paste as one history step',
    afterSlideLayoutPasteUndo.layout === 'ppt-layout-split' &&
      afterSlideLayoutPasteUndo.visible === 'false' &&
      afterSlideLayoutPasteUndo.slideHiddenPlaceholders.split(' ').includes('media') &&
      afterSlideLayoutPasteUndo.selectedIds === beforePlaceholderVisibility.selectedIds,
    {
      afterSlideLayoutPaste,
      afterSlideLayoutPasteUndo,
      beforePlaceholderVisibility,
    },
  )

  await page.eval(`(() => {
    const notes = document.querySelector('[data-ppt-slide-field="notes"]')
    const valueSetter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set

    valueSetter.call(notes, 'Presenter cue: review image crop and final CTA.')
    notes.dispatchEvent(new Event('input', { bubbles: true }))
    notes.dispatchEvent(new Event('change', { bubbles: true }))
  })()`)
  await delay(50)

  const afterNotes = await page.eval(`(() => ({
    notes: document.querySelector('[data-ppt-slide-field="notes"]')?.value ?? '',
  }))()`)

  record('updates PPT speaker notes in inspector', afterNotes.notes === 'Presenter cue: review image crop and final CTA.', afterNotes)
}

async function runSlideTransitionScenario(page) {
  await page.eval(`document.querySelector('.ppt-thumb[aria-label="Open Overview"]')?.click()`)
  await delay(80)

  const initial = await getPPTSlideTransitionState(page)

  record(
    'renders PPT slide transition/timing controls in inspector',
    initial.inspector &&
      initial.type === 'none' &&
      initial.duration === '0' &&
      initial.advanceOnClick === 'true' &&
      initial.advanceAfter === '' &&
      initial.inspectorModel === 'slide-edit-slide-transition-timing' &&
      initial.stageModel === 'slide-edit-slide-transition-timing' &&
      initial.stageType === 'none',
    initial,
  )

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      transition: {
        advanceAfterMs: 2500,
        advanceOnClick: false,
        durationMs: 720,
        type: 'push',
      },
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const afterTransitionPaste = await getPPTSlideTransitionState(page)

  record(
    'pastes JSON slide transition into active PPT slide',
    afterTransitionPaste.importModel === 'ppt-slide-transition-import' &&
      afterTransitionPaste.importFormat === 'application-json-ppt-slide-transition' &&
      afterTransitionPaste.importSlide === 'slide-1' &&
      afterTransitionPaste.importFields ===
        'type durationMs advanceOnClick advanceAfterMs' &&
      afterTransitionPaste.importCommands ===
        'update-slide-transition update-slide-transition update-slide-transition update-slide-transition' &&
      afterTransitionPaste.importCommandFields === 'type durationMs advance advance' &&
      afterTransitionPaste.importType === 'push' &&
      afterTransitionPaste.importDuration === '720' &&
      afterTransitionPaste.importAdvanceOnClick === 'false' &&
      afterTransitionPaste.importAdvanceAfter === '2500' &&
      afterTransitionPaste.importJsonLength > 80 &&
      afterTransitionPaste.type === 'push' &&
      afterTransitionPaste.duration === '720' &&
      afterTransitionPaste.advanceOnClick === 'false' &&
      afterTransitionPaste.advanceAfter === '2500' &&
      afterTransitionPaste.stageType === 'push' &&
      afterTransitionPaste.stageDuration === '720' &&
      afterTransitionPaste.stageAdvanceOnClick === 'false' &&
      afterTransitionPaste.stageAdvanceAfter === '2500',
    {
      afterTransitionPaste,
    },
  )

  await page.eval(`(() => {
    const type = document.querySelector('[data-ppt-slide-transition-field="type"]')
    const duration = document.querySelector('[data-ppt-slide-transition-field="durationMs"]')
    const advanceOnClick = document.querySelector('[data-ppt-slide-transition-field="advanceOnClick"]')
    const advanceAfter = document.querySelector('[data-ppt-slide-transition-field="advanceAfterMs"]')
    const selectSetter = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set
    const inputSetter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set

    selectSetter.call(type, 'fade')
    type.dispatchEvent(new Event('input', { bubbles: true }))
    type.dispatchEvent(new Event('change', { bubbles: true }))

    inputSetter.call(duration, '650')
    duration.dispatchEvent(new Event('input', { bubbles: true }))
    duration.dispatchEvent(new Event('change', { bubbles: true }))

    if (advanceOnClick.checked) {
      advanceOnClick.click()
    }

    inputSetter.call(advanceAfter, '3000')
    advanceAfter.dispatchEvent(new Event('input', { bubbles: true }))
    advanceAfter.dispatchEvent(new Event('change', { bubbles: true }))
  })()`)
  await delay(120)

  const afterFade = await getPPTSlideTransitionState(page)

  record(
    'updates PPT slide transition/timing metadata from inspector',
    afterFade.type === 'fade' &&
      afterFade.duration === '650' &&
      afterFade.advanceOnClick === 'false' &&
      afterFade.advanceAfter === '3000' &&
      afterFade.stageType === 'fade' &&
      afterFade.stageDuration === '650' &&
      afterFade.stageAdvanceOnClick === 'false' &&
      afterFade.stageAdvanceAfter === '3000' &&
      afterFade.command === 'update-slide-transition' &&
      afterFade.commandField === 'advance' &&
      afterFade.commandSlide === 'slide-1' &&
      afterFade.commandType === 'slide-command-effect' &&
      afterFade.commandValue === 'onClick:false;afterMs:3000',
    {
      afterFade,
      initial,
    },
  )

  await page.eval(`(() => {
    const type = document.querySelector('[data-ppt-slide-transition-field="type"]')
    const selectSetter = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set

    selectSetter.call(type, 'push')
    type.dispatchEvent(new Event('input', { bubbles: true }))
    type.dispatchEvent(new Event('change', { bubbles: true }))
  })()`)
  await delay(80)

  const afterPush = await getPPTSlideTransitionState(page)

  record(
    'routes PPT slide transition update through slide-edit command effect',
    afterPush.command === 'update-slide-transition' &&
      afterPush.commandField === 'type' &&
      afterPush.commandSlide === 'slide-1' &&
      afterPush.commandType === 'slide-command-effect' &&
      afterPush.commandValue === 'push' &&
      afterPush.inspectorTypes === 'none fade push',
    afterPush,
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterUndo = await getPPTSlideTransitionState(page)

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const afterRedo = await getPPTSlideTransitionState(page)

  record(
    'undoes and redoes PPT slide transition type as one history step',
    afterPush.type === 'push' &&
      afterUndo.type === 'fade' &&
      afterRedo.type === 'push' &&
      afterRedo.duration === '650' &&
      afterRedo.advanceOnClick === 'false' &&
      afterRedo.advanceAfter === '3000',
    {
      afterFade,
      afterPush,
      afterRedo,
      afterUndo,
    },
  )
}

async function runObjectAnimationScenario(page) {
  await page.eval(`document.querySelector('.ppt-thumb[aria-label="Open Overview"]')?.click()`)
  await delay(80)

  const titlePoint = await getElementCenter(page, 's1-title')
  await clickMouse(page, titlePoint.x, titlePoint.y, 1)
  await delay(80)

  const initial = await getPPTObjectAnimationState(page)

  record(
    'renders PPT object animation/build order controls in inspector',
    initial.inspector &&
      initial.selectedId === 's1-title' &&
      initial.model === 'slide-edit-object-animation' &&
      initial.descriptorModel === 'slide-edit-object-animation' &&
      initial.type === 'none' &&
      initial.trigger === 'onClick' &&
      initial.descriptorPackageType === 'none' &&
      initial.descriptorPackageTrigger === 'on-click' &&
      initial.descriptorTypeOptions === 'none fade-in fly-in' &&
      initial.descriptorTriggerOptions === 'on-click with-previous' &&
      initial.duration === '400' &&
      initial.delay === '0' &&
      initial.typeCommand === 'update-object-animation' &&
      initial.triggerCommand === 'update-object-animation' &&
      initial.durationCommand === 'update-object-animation' &&
      initial.delayCommand === 'update-object-animation' &&
      initial.orderCommand === 'update-object-animation' &&
      initial.buildOrderModel === 'slide-edit-object-animation-build-order' &&
      initial.buildOrder.split(' ').includes('s1-title') &&
      Number(initial.order) > 0 &&
      initial.selectedType === 'none',
    initial,
  )

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      animation: {
        delayMs: 180,
        durationMs: 640,
        order: 2,
        trigger: 'withPrevious',
        type: 'fadeIn',
      },
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const afterAnimationPaste = await getPPTObjectAnimationState(page)

  record(
    'pastes JSON object animation into selected PPT object',
    afterAnimationPaste.importModel === 'ppt-object-animation-import' &&
      afterAnimationPaste.importFormat === 'application-json-ppt-object-animation' &&
      afterAnimationPaste.importSlide === 'slide-1' &&
      afterAnimationPaste.importObjects === 's1-title' &&
      afterAnimationPaste.importFields ===
        'type trigger durationMs delayMs order' &&
      afterAnimationPaste.importCommands ===
        'update-object-animation update-object-animation update-object-animation update-object-animation update-object-animation' &&
      afterAnimationPaste.importCommandFields ===
        'type trigger durationMs delayMs order' &&
      afterAnimationPaste.importType === 'fadeIn' &&
      afterAnimationPaste.importTrigger === 'withPrevious' &&
      afterAnimationPaste.importDuration === '640' &&
      afterAnimationPaste.importDelay === '180' &&
      afterAnimationPaste.importOrder === '2' &&
      afterAnimationPaste.importJsonLength > 80 &&
      afterAnimationPaste.type === 'fadeIn' &&
      afterAnimationPaste.trigger === 'withPrevious' &&
      afterAnimationPaste.duration === '640' &&
      afterAnimationPaste.delay === '180' &&
      afterAnimationPaste.order === '2' &&
      afterAnimationPaste.selectedType === 'fadeIn' &&
      afterAnimationPaste.selectedTrigger === 'withPrevious' &&
      afterAnimationPaste.selectedDuration === '640' &&
      afterAnimationPaste.selectedDelay === '180' &&
      afterAnimationPaste.selectedOrder === '2' &&
      afterAnimationPaste.command === 'update-object-animation' &&
      afterAnimationPaste.commandField === 'order' &&
      afterAnimationPaste.commandObject === 's1-title' &&
      afterAnimationPaste.commandSlide === 'slide-1' &&
      afterAnimationPaste.commandType === 'slide-command-effect' &&
      afterAnimationPaste.commandValue === '2' &&
      afterAnimationPaste.buildOrder.split(' ').includes('s1-title'),
    {
      afterAnimationPaste,
      initial,
    },
  )

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      delayMs: 90,
      durationMs: 520,
      order: 1,
      trigger: 'onClick',
      type: 'flyIn',
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const afterStandaloneAnimationPaste = await getPPTObjectAnimationState(page)

  record(
    'pastes standalone JSON object animation into selected PPT object through slide-edit command-effects',
    afterStandaloneAnimationPaste.importModel === 'ppt-object-animation-import' &&
      afterStandaloneAnimationPaste.importFormat === 'application-json-ppt-object-animation' &&
      afterStandaloneAnimationPaste.importSlide === 'slide-1' &&
      afterStandaloneAnimationPaste.importObjects === 's1-title' &&
      afterStandaloneAnimationPaste.importFields ===
        'type trigger durationMs delayMs order' &&
      afterStandaloneAnimationPaste.importCommands ===
        'update-object-animation update-object-animation update-object-animation update-object-animation update-object-animation' &&
      afterStandaloneAnimationPaste.importCommandFields ===
        'type trigger durationMs delayMs order' &&
      afterStandaloneAnimationPaste.importType === 'flyIn' &&
      afterStandaloneAnimationPaste.importTrigger === 'onClick' &&
      afterStandaloneAnimationPaste.importDuration === '520' &&
      afterStandaloneAnimationPaste.importDelay === '90' &&
      afterStandaloneAnimationPaste.importOrder === '1' &&
      afterStandaloneAnimationPaste.importJsonLength > 70 &&
      afterStandaloneAnimationPaste.type === 'flyIn' &&
      afterStandaloneAnimationPaste.trigger === 'onClick' &&
      afterStandaloneAnimationPaste.duration === '520' &&
      afterStandaloneAnimationPaste.delay === '90' &&
      afterStandaloneAnimationPaste.order === '1' &&
      afterStandaloneAnimationPaste.selectedType === 'flyIn' &&
      afterStandaloneAnimationPaste.selectedTrigger === 'onClick' &&
      afterStandaloneAnimationPaste.selectedDuration === '520' &&
      afterStandaloneAnimationPaste.selectedDelay === '90' &&
      afterStandaloneAnimationPaste.selectedOrder === '1' &&
      afterStandaloneAnimationPaste.command === 'update-object-animation' &&
      afterStandaloneAnimationPaste.commandField === 'order' &&
      afterStandaloneAnimationPaste.commandObject === 's1-title' &&
      afterStandaloneAnimationPaste.commandSlide === 'slide-1' &&
      afterStandaloneAnimationPaste.commandType === 'slide-command-effect' &&
      afterStandaloneAnimationPaste.commandValue === '1' &&
      afterStandaloneAnimationPaste.buildOrder.split(' ').includes('s1-title'),
    {
      afterAnimationPaste,
      afterStandaloneAnimationPaste,
    },
  )

  await page.eval(`(() => {
    const type = document.querySelector('[data-ppt-animation-field="type"]')
    const trigger = document.querySelector('[data-ppt-animation-field="trigger"]')
    const duration = document.querySelector('[data-ppt-animation-field="durationMs"]')
    const delay = document.querySelector('[data-ppt-animation-field="delayMs"]')
    const order = document.querySelector('[data-ppt-animation-field="order"]')
    const selectSetter = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set
    const inputSetter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set

    selectSetter.call(type, 'fadeIn')
    type.dispatchEvent(new Event('input', { bubbles: true }))
    type.dispatchEvent(new Event('change', { bubbles: true }))

    selectSetter.call(trigger, 'withPrevious')
    trigger.dispatchEvent(new Event('input', { bubbles: true }))
    trigger.dispatchEvent(new Event('change', { bubbles: true }))

    inputSetter.call(duration, '800')
    duration.dispatchEvent(new Event('input', { bubbles: true }))
    duration.dispatchEvent(new Event('change', { bubbles: true }))

    inputSetter.call(delay, '200')
    delay.dispatchEvent(new Event('input', { bubbles: true }))
    delay.dispatchEvent(new Event('change', { bubbles: true }))

    inputSetter.call(order, '3')
    order.dispatchEvent(new Event('input', { bubbles: true }))
    order.dispatchEvent(new Event('change', { bubbles: true }))
  })()`)
  await delay(120)

  const afterFade = await getPPTObjectAnimationState(page)

  record(
    'updates PPT object animation/build order metadata from inspector',
    afterFade.type === 'fadeIn' &&
      afterFade.trigger === 'withPrevious' &&
      afterFade.duration === '800' &&
      afterFade.delay === '200' &&
      afterFade.order === '3' &&
      afterFade.selectedType === 'fadeIn' &&
      afterFade.selectedTrigger === 'withPrevious' &&
      afterFade.selectedDuration === '800' &&
      afterFade.selectedDelay === '200' &&
      afterFade.selectedOrder === '3' &&
      afterFade.command === 'update-object-animation' &&
      afterFade.commandField === 'order' &&
      afterFade.commandObject === 's1-title' &&
      afterFade.commandSlide === 'slide-1' &&
      afterFade.commandType === 'slide-command-effect' &&
      afterFade.commandValue === '3' &&
      afterFade.descriptorPackageType === 'fade-in' &&
      afterFade.descriptorPackageTrigger === 'with-previous' &&
      afterFade.buildOrderModel === 'slide-edit-object-animation-build-order' &&
      afterFade.buildOrder.split(' ').includes('s1-title') &&
      afterFade.buildOrder.split(' ')[0] !== 's1-title',
    {
      afterFade,
      initial,
    },
  )

  await page.eval(`(() => {
    const type = document.querySelector('[data-ppt-animation-field="type"]')
    const selectSetter = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set

    selectSetter.call(type, 'flyIn')
    type.dispatchEvent(new Event('input', { bubbles: true }))
    type.dispatchEvent(new Event('change', { bubbles: true }))
  })()`)
  await delay(80)

  const afterFly = await getPPTObjectAnimationState(page)

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterUndo = await getPPTObjectAnimationState(page)

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const afterRedo = await getPPTObjectAnimationState(page)

  record(
    'undoes and redoes PPT object animation type as one history step',
    afterFly.type === 'flyIn' &&
      afterFly.commandField === 'type' &&
      afterFly.commandValue === 'fly-in' &&
      afterUndo.type === 'fadeIn' &&
      afterRedo.type === 'flyIn' &&
      afterRedo.trigger === 'withPrevious' &&
      afterRedo.duration === '800' &&
      afterRedo.delay === '200' &&
      afterRedo.order === '3',
    {
      afterFade,
      afterFly,
      afterRedo,
      afterUndo,
    },
  )

  await page.eval(`document.querySelector('[data-ppt-present-start]')?.click()`)
  await delay(120)

  const preview = await getPPTObjectAnimationPreviewState(page, 's1-title')

  record(
    'applies PPT object animation metadata in presentation preview',
    preview.open &&
      preview.type === 'flyIn' &&
      preview.trigger === 'withPrevious' &&
      preview.duration === '800' &&
      preview.delay === '200' &&
      preview.order === '3' &&
      preview.buildOrderModel === 'slide-edit-object-animation-build-order' &&
      preview.buildOrder.split(' ').includes('s1-title'),
    preview,
  )

  await page.eval(`document.querySelector('[data-ppt-present-exit]')?.click()`)
  await delay(80)
}

async function runObjectOpacityScenario(page) {
  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)
  await page.eval(`document.querySelector('.ppt-thumb[aria-label="Open Overview"]')?.click()`)
  await delay(80)
  await waitUntil(
    () => page.eval(`!!document.querySelector('[data-ppt-element="s1-card-1"]')`),
    'Expected overview card to render before object opacity selection',
  )

  const cardPoint = await getElementCenter(page, 's1-card-1')
  await clickMouse(page, cardPoint.x, cardPoint.y, 1)
  await delay(80)

  const initial = await getPPTObjectOpacityState(page)
  const targetId = initial.selectedId

  record(
    'renders PPT object opacity control in inspector',
    targetId.length > 0 &&
      initial.opacity === '1' &&
      initial.selectedOpacity === '1' &&
      initial.thumbOpacity === '1',
    initial,
  )

  await page.eval(`(() => {
    const opacity = document.querySelector('[data-ppt-style-field="opacity"]')
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set

    setter.call(opacity, '0.42')
    opacity.dispatchEvent(new Event('input', { bubbles: true }))
    opacity.dispatchEvent(new Event('change', { bubbles: true }))
  })()`)
  await delay(120)

  const afterOpacity = await getPPTObjectOpacityState(page, targetId)

  record(
    'updates PPT object opacity metadata from inspector',
    afterOpacity.opacity === '0.42' &&
      afterOpacity.descriptorSurface === 'object-opacity' &&
      afterOpacity.descriptorCommand === 'update-object-opacity' &&
      afterOpacity.descriptorControl === 'opacity-slider' &&
      afterOpacity.descriptorAttribute === 'data-slide-object-opacity' &&
      afterOpacity.descriptorAttributeValue === '0.42' &&
      afterOpacity.model === 'slide-edit-object-opacity' &&
      afterOpacity.command === 'update-object-opacity' &&
      afterOpacity.commandField === 'opacity' &&
      afterOpacity.commandObject === afterOpacity.selectedId &&
      afterOpacity.commandSlide === 'slide-1' &&
      afterOpacity.commandType === 'slide-command-effect' &&
      afterOpacity.commandValue === '0.42' &&
      afterOpacity.selectedOpacity === '0.42' &&
      afterOpacity.selectedStyleOpacity === '0.42' &&
      afterOpacity.thumbOpacity === '0.42' &&
      afterOpacity.thumbStyleOpacity === '0.42',
    {
      afterOpacity,
      initial,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterUndo = await getPPTObjectOpacityState(page, targetId)

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const afterRedo = await getPPTObjectOpacityState(page, targetId)

  record(
    'undoes and redoes PPT object opacity as one history step',
    afterUndo.opacity === '1' &&
      afterUndo.selectedOpacity === '1' &&
      afterRedo.opacity === '0.42' &&
      afterRedo.selectedOpacity === '0.42',
    {
      afterOpacity,
      afterRedo,
      afterUndo,
    },
  )

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      objectOpacity: { value: 0.67 },
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const afterOpacityJSONPaste = await getPPTObjectOpacityState(page, targetId)

  record(
    'pastes JSON object opacity through slide-edit command effect',
    afterOpacityJSONPaste.opacityImportModel === 'ppt-object-opacity-import' &&
      afterOpacityJSONPaste.opacityImportFormat === 'application-json-ppt-object-opacity' &&
      afterOpacityJSONPaste.opacityImportSlide === 'slide-1' &&
      afterOpacityJSONPaste.opacityImportObjects === targetId &&
      afterOpacityJSONPaste.opacityImportFields === 'opacity' &&
      afterOpacityJSONPaste.opacityImportCommands === 'update-object-opacity' &&
      afterOpacityJSONPaste.opacityImportCommandFields === 'opacity' &&
      afterOpacityJSONPaste.opacityImportCommandTargets === targetId &&
      afterOpacityJSONPaste.opacityImportCommandTypes === 'slide-command-effect' &&
      afterOpacityJSONPaste.opacityImportCommandValues === '0.67' &&
      afterOpacityJSONPaste.opacityImportValue === '0.67' &&
      afterOpacityJSONPaste.opacityImportJsonLength > 30 &&
      afterOpacityJSONPaste.command === 'update-object-opacity' &&
      afterOpacityJSONPaste.commandField === 'opacity' &&
      afterOpacityJSONPaste.commandValue === '0.67' &&
      afterOpacityJSONPaste.opacity === '0.67' &&
      afterOpacityJSONPaste.selectedOpacity === '0.67' &&
      afterOpacityJSONPaste.selectedStyleOpacity === '0.67' &&
      afterOpacityJSONPaste.thumbOpacity === '0.67' &&
      afterOpacityJSONPaste.thumbStyleOpacity === '0.67',
    {
      afterOpacityJSONPaste,
      afterRedo,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterOpacityJSONUndo = await getPPTObjectOpacityState(page, targetId)

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const afterOpacityJSONRedo = await getPPTObjectOpacityState(page, targetId)

  record(
    'undoes and redoes PPT object opacity JSON as one history step',
    afterOpacityJSONUndo.opacity === '0.42' &&
      afterOpacityJSONUndo.selectedOpacity === '0.42' &&
      afterOpacityJSONRedo.opacity === '0.67' &&
      afterOpacityJSONRedo.selectedOpacity === '0.67',
    {
      afterOpacityJSONPaste,
      afterOpacityJSONRedo,
      afterOpacityJSONUndo,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({ opacity: 0.64 })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const afterStandaloneOpacityJSONPaste = await getPPTObjectOpacityState(page, targetId)

  record(
    'pastes standalone JSON object opacity through slide-edit command effect',
    afterStandaloneOpacityJSONPaste.opacityImportModel === 'ppt-object-opacity-import' &&
      afterStandaloneOpacityJSONPaste.opacityImportFormat === 'application-json-ppt-object-opacity' &&
      afterStandaloneOpacityJSONPaste.opacityImportSlide === 'slide-1' &&
      afterStandaloneOpacityJSONPaste.opacityImportObjects === targetId &&
      afterStandaloneOpacityJSONPaste.opacityImportFields === 'opacity' &&
      afterStandaloneOpacityJSONPaste.opacityImportCommands === 'update-object-opacity' &&
      afterStandaloneOpacityJSONPaste.opacityImportCommandFields === 'opacity' &&
      afterStandaloneOpacityJSONPaste.opacityImportCommandTargets === targetId &&
      afterStandaloneOpacityJSONPaste.opacityImportCommandTypes === 'slide-command-effect' &&
      afterStandaloneOpacityJSONPaste.opacityImportCommandValues === '0.64' &&
      afterStandaloneOpacityJSONPaste.opacityImportValue === '0.64' &&
      afterStandaloneOpacityJSONPaste.opacityImportJsonLength > 10 &&
      afterStandaloneOpacityJSONPaste.command === 'update-object-opacity' &&
      afterStandaloneOpacityJSONPaste.commandField === 'opacity' &&
      afterStandaloneOpacityJSONPaste.commandValue === '0.64' &&
      afterStandaloneOpacityJSONPaste.opacity === '0.64' &&
      afterStandaloneOpacityJSONPaste.selectedOpacity === '0.64' &&
      afterStandaloneOpacityJSONPaste.selectedStyleOpacity === '0.64' &&
      afterStandaloneOpacityJSONPaste.thumbOpacity === '0.64' &&
      afterStandaloneOpacityJSONPaste.thumbStyleOpacity === '0.64',
    {
      afterOpacityJSONRedo,
      afterStandaloneOpacityJSONPaste,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterStandaloneOpacityJSONUndo = await getPPTObjectOpacityState(page, targetId)

  record(
    'undoes standalone JSON object opacity as one history step',
    afterStandaloneOpacityJSONUndo.opacity === '0.42' &&
      afterStandaloneOpacityJSONUndo.selectedOpacity === '0.42' &&
      afterStandaloneOpacityJSONUndo.selectedStyleOpacity === '0.42' &&
      afterStandaloneOpacityJSONUndo.thumbOpacity === '0.42',
    {
      afterStandaloneOpacityJSONPaste,
      afterStandaloneOpacityJSONUndo,
    },
  )

  await page.eval(`document.querySelector('[data-ppt-present-start]')?.click()`)
  await delay(120)

  const preview = await page.eval(`(() => {
    const overlay = document.querySelector('[data-ppt-presentation]')
    const element = overlay?.querySelector(${JSON.stringify(`[data-ppt-element="${targetId}"]`)})

    return {
      opacity: element?.getAttribute('data-ppt-opacity') ?? '',
      open: !!overlay,
      styleOpacity: element?.style.opacity ?? '',
    }
  })()`)

  record(
    'keeps PPT object opacity metadata in presentation preview',
    preview.open &&
      preview.opacity === '0.42' &&
      preview.styleOpacity === '0.42',
    preview,
  )

  await page.eval(`document.querySelector('[data-ppt-present-exit]')?.click()`)
  await delay(80)
}

async function runObjectShadowScenario(page) {
  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)
  await page.eval(`document.querySelector('.ppt-thumb[aria-label="Open Overview"]')?.click()`)
  await delay(80)
  await waitUntil(
    () => page.eval(`!!document.querySelector('[data-ppt-element="s1-card-1"]')`),
    'Expected overview card to render before object hyperlink selection',
  )

  const cardPoint = await getElementCenter(page, 's1-card-1')
  await clickMouse(page, cardPoint.x, cardPoint.y, 1)
  await delay(80)

  const initial = await getPPTObjectShadowState(page)
  const targetId = initial.selectedId

  record(
    'renders PPT object shadow controls in inspector',
    targetId.length > 0 &&
      initial.descriptorAttribute === 'data-slide-object-shadow' &&
      initial.descriptorAttributeValue === 'none' &&
      initial.descriptorCommand === 'update-object-shadow' &&
      initial.descriptorControl === 'shadow-toggle' &&
      initial.descriptorEnabled === 'false' &&
      initial.descriptorSurface === 'object-shadow' &&
      initial.inspectorEnabled === 'false' &&
      initial.enabled === false &&
      initial.model === 'slide-edit-object-shadow' &&
      initial.opacity === '0.22' &&
      initial.opacityDisabled === true &&
      initial.selectedShadow === '' &&
      initial.thumbShadow === '',
    initial,
  )

  await page.eval(`document.querySelector('[data-ppt-shadow-field="enabled"]')?.click()`)
  await delay(120)

  const afterEnable = await getPPTObjectShadowState(page, targetId)

  record(
    'enables PPT object shadow metadata from inspector',
    afterEnable.command === 'update-object-shadow' &&
      afterEnable.commandField === 'enabled' &&
      afterEnable.commandObject === targetId &&
      afterEnable.commandSlide === 'slide-1' &&
      afterEnable.commandType === 'slide-command-effect' &&
      afterEnable.commandValue === 'true' &&
      afterEnable.descriptorAttribute === 'data-slide-object-shadow' &&
      afterEnable.descriptorCommand === 'update-object-shadow' &&
      afterEnable.descriptorControl === 'shadow-toggle' &&
      afterEnable.descriptorEnabled === 'true' &&
      afterEnable.descriptorOpacity === '0.22' &&
      afterEnable.descriptorSurface === 'object-shadow' &&
      afterEnable.enabled === true &&
      afterEnable.inspectorEnabled === 'true' &&
      afterEnable.model === 'slide-edit-object-shadow' &&
      afterEnable.selectedShadow === 'true' &&
      afterEnable.selectedOpacity === '0.22' &&
      afterEnable.selectedFilter.includes('drop-shadow') &&
      afterEnable.thumbShadow === 'true' &&
      afterEnable.thumbFilter.includes('drop-shadow'),
    {
      afterEnable,
      initial,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterEnableUndo = await getPPTObjectShadowState(page, targetId)

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const afterEnableRedo = await getPPTObjectShadowState(page, targetId)

  record(
    'undoes and redoes PPT object shadow enable as one history step',
    afterEnableUndo.enabled === false &&
      afterEnableUndo.selectedShadow === '' &&
      afterEnableRedo.enabled === true &&
      afterEnableRedo.selectedShadow === 'true',
    {
      afterEnableRedo,
      afterEnableUndo,
    },
  )

  await page.eval(`(() => {
    const setInputValue = (selector, value) => {
      const input = document.querySelector(selector)
      const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set

      setter.call(input, value)
      input.dispatchEvent(new Event('input', { bubbles: true }))
      input.dispatchEvent(new Event('change', { bubbles: true }))
    }

    setInputValue('[data-ppt-shadow-field="color"]', '#334155')
    setInputValue('[data-ppt-shadow-field="opacity"]', '0.36')
    setInputValue('[data-ppt-shadow-field="blur"]', '18')
    setInputValue('[data-ppt-shadow-field="distance"]', '12')
    setInputValue('[data-ppt-shadow-field="angle"]', '60')
  })()`)
  await delay(160)

  const afterFields = await getPPTObjectShadowState(page, targetId)

  record(
    'updates PPT object shadow fields from inspector',
    afterFields.color === '#334155' &&
      afterFields.command === 'update-object-shadow' &&
      afterFields.commandField === 'angle' &&
      afterFields.commandObject === targetId &&
      afterFields.commandSlide === 'slide-1' &&
      afterFields.commandType === 'slide-command-effect' &&
      afterFields.commandValue === '60' &&
      afterFields.descriptorAngle === '60' &&
      afterFields.descriptorBlur === '18' &&
      afterFields.descriptorColor === '#334155' &&
      afterFields.descriptorDistance === '12' &&
      afterFields.descriptorEnabled === 'true' &&
      afterFields.descriptorOpacity === '0.36' &&
      afterFields.opacity === '0.36' &&
      afterFields.opacityControl === 'slider' &&
      afterFields.opacityUnit === 'ratio' &&
      afterFields.blur === '18' &&
      afterFields.blurControl === 'number' &&
      afterFields.blurUnit === 'px' &&
      afterFields.distance === '12' &&
      afterFields.distanceControl === 'number' &&
      afterFields.distanceUnit === 'px' &&
      afterFields.angle === '60' &&
      afterFields.angleControl === 'number' &&
      afterFields.angleUnit === 'deg' &&
      afterFields.selectedColor === '#334155' &&
      afterFields.selectedOpacity === '0.36' &&
      afterFields.selectedBlur === '18' &&
      afterFields.selectedDistance === '12' &&
      afterFields.selectedAngle === '60' &&
      afterFields.thumbOpacity === '0.36',
    {
      afterEnableRedo,
      afterFields,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterFieldUndo = await getPPTObjectShadowState(page, targetId)

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const afterFieldRedo = await getPPTObjectShadowState(page, targetId)

  record(
    'undoes and redoes PPT object shadow field as one history step',
    afterFieldUndo.angle === '45' &&
      afterFieldUndo.selectedAngle === '45' &&
      afterFieldRedo.angle === '60' &&
      afterFieldRedo.selectedAngle === '60',
    {
      afterFieldRedo,
      afterFieldUndo,
    },
  )

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      objectStyle: {
        opacity: 0.42,
        shadow: {
          angle: 60,
          blur: 18,
          color: '#334155',
          distance: 12,
          enabled: true,
          opacity: 0.36,
        },
      },
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const afterStylePaste = await getPPTObjectShadowState(page, targetId)

  record(
    'pastes JSON object visual style into selected PPT object',
    afterStylePaste.objectStyleImportModel === 'ppt-object-style-import' &&
      afterStylePaste.objectStyleImportFormat === 'application-json-ppt-object-style' &&
      afterStylePaste.objectStyleImportCommand === 'paste-object-formatting' &&
      afterStylePaste.objectStyleImportCommandTargets === targetId &&
      afterStylePaste.objectStyleImportCommandType === 'slide-command-effect' &&
      afterStylePaste.objectStyleImportObjects === targetId &&
      afterStylePaste.objectStyleImportCategories === 'object-effect' &&
      afterStylePaste.objectStyleImportFields === 'opacity shadow' &&
      afterStylePaste.objectStyleImportOpacity === '0.42' &&
      afterStylePaste.objectStyleImportShadowEnabled === 'true' &&
      afterStylePaste.objectStyleImportShadowColor === '#334155' &&
      afterStylePaste.objectStyleImportShadowOpacity === '0.36' &&
      afterStylePaste.objectStyleImportShadowBlur === '18' &&
      afterStylePaste.objectStyleImportShadowDistance === '12' &&
      afterStylePaste.objectStyleImportShadowAngle === '60' &&
      afterStylePaste.objectStyleImportJsonLength > 100 &&
      afterStylePaste.selectedObjectOpacity === '0.42' &&
      afterStylePaste.selectedShadow === 'true' &&
      afterStylePaste.selectedColor === '#334155' &&
      afterStylePaste.selectedOpacity === '0.36' &&
      afterStylePaste.selectedBlur === '18' &&
      afterStylePaste.selectedDistance === '12' &&
      afterStylePaste.selectedAngle === '60' &&
      afterStylePaste.thumbOpacity === '0.36',
    {
      afterFieldRedo,
      afterStylePaste,
    },
  )

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      opacity: 0.58,
      shadow: {
        angle: 30,
        blur: 20,
        color: '#0f766e',
        distance: 8,
        enabled: true,
        opacity: 0.44,
      },
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const afterStandaloneObjectStylePaste = await getPPTObjectShadowState(page, targetId)

  record(
    'pastes standalone PPT object style fields through style clipboard object category',
    afterStandaloneObjectStylePaste.objectStyleImportModel === 'ppt-object-style-import' &&
      afterStandaloneObjectStylePaste.objectStyleImportFormat === 'application-json-ppt-object-style' &&
      afterStandaloneObjectStylePaste.objectStyleImportCommand === 'paste-object-formatting' &&
      afterStandaloneObjectStylePaste.objectStyleImportCommandTargets === targetId &&
      afterStandaloneObjectStylePaste.objectStyleImportCommandType === 'slide-command-effect' &&
      afterStandaloneObjectStylePaste.objectStyleImportObjects === targetId &&
      afterStandaloneObjectStylePaste.objectStyleImportCategories === 'object-effect' &&
      afterStandaloneObjectStylePaste.objectStyleImportFields === 'opacity shadow' &&
      afterStandaloneObjectStylePaste.objectStyleImportOpacity === '0.58' &&
      afterStandaloneObjectStylePaste.objectStyleImportShadowEnabled === 'true' &&
      afterStandaloneObjectStylePaste.objectStyleImportShadowColor === '#0f766e' &&
      afterStandaloneObjectStylePaste.objectStyleImportShadowOpacity === '0.44' &&
      afterStandaloneObjectStylePaste.objectStyleImportShadowBlur === '20' &&
      afterStandaloneObjectStylePaste.objectStyleImportShadowDistance === '8' &&
      afterStandaloneObjectStylePaste.objectStyleImportShadowAngle === '30' &&
      afterStandaloneObjectStylePaste.objectStyleImportJsonLength > 100 &&
      afterStandaloneObjectStylePaste.selectedObjectOpacity === '0.58' &&
      afterStandaloneObjectStylePaste.selectedShadow === 'true' &&
      afterStandaloneObjectStylePaste.selectedColor === '#0f766e' &&
      afterStandaloneObjectStylePaste.selectedOpacity === '0.44' &&
      afterStandaloneObjectStylePaste.selectedBlur === '20' &&
      afterStandaloneObjectStylePaste.selectedDistance === '8' &&
      afterStandaloneObjectStylePaste.selectedAngle === '30' &&
      afterStandaloneObjectStylePaste.thumbOpacity === '0.44',
    {
      afterStandaloneObjectStylePaste,
      afterStylePaste,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterStandaloneObjectStyleUndo = await getPPTObjectShadowState(page, targetId)

  record(
    'undoes standalone PPT object style fields as one history step',
    afterStandaloneObjectStyleUndo.selectedObjectOpacity === afterStylePaste.selectedObjectOpacity &&
      afterStandaloneObjectStyleUndo.selectedColor === afterStylePaste.selectedColor &&
      afterStandaloneObjectStyleUndo.selectedOpacity === afterStylePaste.selectedOpacity &&
      afterStandaloneObjectStyleUndo.selectedBlur === afterStylePaste.selectedBlur &&
      afterStandaloneObjectStyleUndo.selectedDistance === afterStylePaste.selectedDistance &&
      afterStandaloneObjectStyleUndo.selectedAngle === afterStylePaste.selectedAngle,
    {
      afterStandaloneObjectStylePaste,
      afterStandaloneObjectStyleUndo,
      afterStylePaste,
    },
  )

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      objectShadow: {
        angle: 120,
        blur: 24,
        color: '#475569',
        distance: 14,
        enabled: true,
        opacity: 0.48,
      },
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const afterShadowJSONPaste = await getPPTObjectShadowState(page, targetId)

  record(
    'pastes JSON object shadow through slide-edit command effects',
    afterShadowJSONPaste.shadowImportModel === 'ppt-object-shadow-import' &&
      afterShadowJSONPaste.shadowImportFormat === 'application-json-ppt-object-shadow' &&
      afterShadowJSONPaste.shadowImportSlide === 'slide-1' &&
      afterShadowJSONPaste.shadowImportObjects === targetId &&
      afterShadowJSONPaste.shadowImportFields === 'enabled color opacity blur distance angle' &&
      afterShadowJSONPaste.shadowImportCommands === [
        'update-object-shadow',
        'update-object-shadow',
        'update-object-shadow',
        'update-object-shadow',
        'update-object-shadow',
        'update-object-shadow',
      ].join(' ') &&
      afterShadowJSONPaste.shadowImportCommandFields === 'enabled color opacity blur distance angle' &&
      afterShadowJSONPaste.shadowImportCommandTargets === [
        targetId,
        targetId,
        targetId,
        targetId,
        targetId,
        targetId,
      ].join(' ') &&
      afterShadowJSONPaste.shadowImportCommandTypes === [
        'slide-command-effect',
        'slide-command-effect',
        'slide-command-effect',
        'slide-command-effect',
        'slide-command-effect',
        'slide-command-effect',
      ].join(' ') &&
      afterShadowJSONPaste.shadowImportCommandValues === 'true #475569 0.48 24 14 120' &&
      afterShadowJSONPaste.shadowImportEnabled === 'true' &&
      afterShadowJSONPaste.shadowImportColor === '#475569' &&
      afterShadowJSONPaste.shadowImportOpacity === '0.48' &&
      afterShadowJSONPaste.shadowImportBlur === '24' &&
      afterShadowJSONPaste.shadowImportDistance === '14' &&
      afterShadowJSONPaste.shadowImportAngle === '120' &&
      afterShadowJSONPaste.shadowImportJsonLength > 100 &&
      afterShadowJSONPaste.command === 'update-object-shadow' &&
      afterShadowJSONPaste.commandField === 'angle' &&
      afterShadowJSONPaste.commandObject === targetId &&
      afterShadowJSONPaste.commandSlide === 'slide-1' &&
      afterShadowJSONPaste.commandType === 'slide-command-effect' &&
      afterShadowJSONPaste.commandValue === '120' &&
      afterShadowJSONPaste.selectedShadow === 'true' &&
      afterShadowJSONPaste.selectedColor === '#475569' &&
      afterShadowJSONPaste.selectedOpacity === '0.48' &&
      afterShadowJSONPaste.selectedBlur === '24' &&
      afterShadowJSONPaste.selectedDistance === '14' &&
      afterShadowJSONPaste.selectedAngle === '120' &&
      afterShadowJSONPaste.thumbOpacity === '0.48',
    {
      afterShadowJSONPaste,
      afterStylePaste,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterShadowJSONUndo = await getPPTObjectShadowState(page, targetId)

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const afterShadowJSONRedo = await getPPTObjectShadowState(page, targetId)

  record(
    'undoes and redoes PPT object shadow JSON as one history step',
    afterShadowJSONUndo.selectedColor === '#334155' &&
      afterShadowJSONUndo.selectedOpacity === '0.36' &&
      afterShadowJSONUndo.selectedBlur === '18' &&
      afterShadowJSONUndo.selectedDistance === '12' &&
      afterShadowJSONUndo.selectedAngle === '60' &&
      afterShadowJSONRedo.selectedColor === '#475569' &&
      afterShadowJSONRedo.selectedOpacity === '0.48' &&
      afterShadowJSONRedo.selectedBlur === '24' &&
      afterShadowJSONRedo.selectedDistance === '14' &&
      afterShadowJSONRedo.selectedAngle === '120',
    {
      afterShadowJSONRedo,
      afterShadowJSONUndo,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      angle: 30,
      blur: 20,
      color: '#0f766e',
      distance: 8,
      enabled: true,
      opacity: 0.44,
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const afterStandaloneShadowJSONPaste = await getPPTObjectShadowState(page, targetId)

  record(
    'pastes standalone JSON object shadow through slide-edit command effects',
    afterStandaloneShadowJSONPaste.shadowImportModel === 'ppt-object-shadow-import' &&
      afterStandaloneShadowJSONPaste.shadowImportFormat === 'application-json-ppt-object-shadow' &&
      afterStandaloneShadowJSONPaste.shadowImportSlide === 'slide-1' &&
      afterStandaloneShadowJSONPaste.shadowImportObjects === targetId &&
      afterStandaloneShadowJSONPaste.shadowImportFields === 'enabled color opacity blur distance angle' &&
      afterStandaloneShadowJSONPaste.shadowImportCommands === [
        'update-object-shadow',
        'update-object-shadow',
        'update-object-shadow',
        'update-object-shadow',
        'update-object-shadow',
        'update-object-shadow',
      ].join(' ') &&
      afterStandaloneShadowJSONPaste.shadowImportCommandFields === 'enabled color opacity blur distance angle' &&
      afterStandaloneShadowJSONPaste.shadowImportCommandTargets === [
        targetId,
        targetId,
        targetId,
        targetId,
        targetId,
        targetId,
      ].join(' ') &&
      afterStandaloneShadowJSONPaste.shadowImportCommandTypes === [
        'slide-command-effect',
        'slide-command-effect',
        'slide-command-effect',
        'slide-command-effect',
        'slide-command-effect',
        'slide-command-effect',
      ].join(' ') &&
      afterStandaloneShadowJSONPaste.shadowImportCommandValues === 'true #0f766e 0.44 20 8 30' &&
      afterStandaloneShadowJSONPaste.shadowImportEnabled === 'true' &&
      afterStandaloneShadowJSONPaste.shadowImportColor === '#0f766e' &&
      afterStandaloneShadowJSONPaste.shadowImportOpacity === '0.44' &&
      afterStandaloneShadowJSONPaste.shadowImportBlur === '20' &&
      afterStandaloneShadowJSONPaste.shadowImportDistance === '8' &&
      afterStandaloneShadowJSONPaste.shadowImportAngle === '30' &&
      afterStandaloneShadowJSONPaste.shadowImportJsonLength > 70 &&
      afterStandaloneShadowJSONPaste.command === 'update-object-shadow' &&
      afterStandaloneShadowJSONPaste.commandField === 'angle' &&
      afterStandaloneShadowJSONPaste.commandObject === targetId &&
      afterStandaloneShadowJSONPaste.commandSlide === 'slide-1' &&
      afterStandaloneShadowJSONPaste.commandType === 'slide-command-effect' &&
      afterStandaloneShadowJSONPaste.commandValue === '30' &&
      afterStandaloneShadowJSONPaste.selectedShadow === 'true' &&
      afterStandaloneShadowJSONPaste.selectedColor === '#0f766e' &&
      afterStandaloneShadowJSONPaste.selectedOpacity === '0.44' &&
      afterStandaloneShadowJSONPaste.selectedBlur === '20' &&
      afterStandaloneShadowJSONPaste.selectedDistance === '8' &&
      afterStandaloneShadowJSONPaste.selectedAngle === '30' &&
      afterStandaloneShadowJSONPaste.thumbOpacity === '0.44',
    {
      afterShadowJSONRedo,
      afterStandaloneShadowJSONPaste,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterStandaloneShadowJSONUndo = await getPPTObjectShadowState(page, targetId)

  record(
    'undoes standalone JSON object shadow as one history step',
    afterStandaloneShadowJSONUndo.selectedColor === '#334155' &&
      afterStandaloneShadowJSONUndo.selectedOpacity === '0.36' &&
      afterStandaloneShadowJSONUndo.selectedBlur === '18' &&
      afterStandaloneShadowJSONUndo.selectedDistance === '12' &&
      afterStandaloneShadowJSONUndo.selectedAngle === '60',
    {
      afterStandaloneShadowJSONPaste,
      afterStandaloneShadowJSONUndo,
    },
  )

  await page.eval(`document.querySelector('[data-ppt-present-start]')?.click()`)
  await delay(120)

  const preview = await page.eval(`(() => {
    const overlay = document.querySelector('[data-ppt-presentation]')
    const element = overlay?.querySelector(${JSON.stringify(`[data-ppt-element="${targetId}"]`)})

    return {
      angle: element?.getAttribute('data-ppt-shadow-angle') ?? '',
      color: element?.getAttribute('data-ppt-shadow-color') ?? '',
      distance: element?.getAttribute('data-ppt-shadow-distance') ?? '',
      open: !!overlay,
      opacity: element?.getAttribute('data-ppt-shadow-opacity') ?? '',
      shadow: element?.getAttribute('data-ppt-shadow') ?? '',
      styleFilter: element?.style.filter ?? '',
    }
  })()`)

  record(
    'keeps PPT object shadow metadata in presentation preview',
    preview.open &&
      preview.shadow === 'true' &&
      preview.color === '#334155' &&
      preview.opacity === '0.36' &&
      preview.distance === '12' &&
      preview.angle === '60' &&
      preview.styleFilter.includes('drop-shadow'),
    preview,
  )

  await page.eval(`document.querySelector('[data-ppt-present-exit]')?.click()`)
  await delay(80)
}

async function runObjectHyperlinkScenario(page) {
  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)
  await page.eval(`document.querySelector('.ppt-thumb[aria-label="Open Overview"]')?.click()`)
  await delay(80)

  const cardPoint = await getElementCenter(page, 's1-card-1')
  await clickMouse(page, cardPoint.x, cardPoint.y, 1)
  await delay(80)

  const initial = await getPPTObjectHyperlinkState(page)
  const targetId = initial.selectedId
  const url = 'https://example.com/ppt'

  record(
    'renders PPT object hyperlink control in inspector',
    targetId.length > 0 &&
      initial.url === '' &&
      initial.selectedUrl === '' &&
      initial.thumbUrl === '',
    initial,
  )

  await page.eval(`((url) => {
    const input = document.querySelector('[data-ppt-style-field="hyperlink"]')
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set

    setter.call(input, url)
    input.dispatchEvent(new Event('input', { bubbles: true }))
    input.dispatchEvent(new Event('change', { bubbles: true }))
  })(${JSON.stringify(url)})`)
  await delay(120)

  const afterUrl = await getPPTObjectHyperlinkState(page, targetId)

  record(
    'updates PPT object hyperlink metadata from inspector',
    afterUrl.url === url &&
      afterUrl.descriptorSurface === 'object-hyperlink' &&
      afterUrl.descriptorCommand === 'update-object-hyperlink' &&
      afterUrl.descriptorControl === 'url' &&
      afterUrl.descriptorAttribute === 'data-slide-object-hyperlink' &&
      afterUrl.descriptorEnabled === 'true' &&
      afterUrl.descriptorTarget === 'same-context' &&
      afterUrl.descriptorUrl === url &&
      afterUrl.model === 'slide-edit-object-hyperlink' &&
      afterUrl.command === 'update-object-hyperlink' &&
      afterUrl.commandField === 'url' &&
      afterUrl.commandObject === afterUrl.selectedId &&
      afterUrl.commandSlide === 'slide-1' &&
      afterUrl.commandType === 'slide-command-effect' &&
      afterUrl.commandValue === url &&
      afterUrl.selectedUrl === url &&
      afterUrl.thumbUrl === url,
    {
      afterUrl,
      initial,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterUndo = await getPPTObjectHyperlinkState(page, targetId)

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const afterRedo = await getPPTObjectHyperlinkState(page, targetId)

  record(
    'undoes and redoes PPT object hyperlink URL as one history step',
    afterUndo.url === '' &&
      afterUndo.selectedUrl === '' &&
      afterRedo.url === url &&
      afterRedo.selectedUrl === url,
    {
      afterRedo,
      afterUndo,
    },
  )

  await page.eval(`(() => {
    const input = document.querySelector('[data-ppt-style-field="hyperlink"]')
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set

    setter.call(input, '')
    input.dispatchEvent(new Event('input', { bubbles: true }))
    input.dispatchEvent(new Event('change', { bubbles: true }))
  })()`)
  await delay(120)

  const afterClear = await getPPTObjectHyperlinkState(page, targetId)

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterClearUndo = await getPPTObjectHyperlinkState(page, targetId)

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const afterClearRedo = await getPPTObjectHyperlinkState(page, targetId)

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterRestore = await getPPTObjectHyperlinkState(page, targetId)

  record(
    'removes and restores PPT object hyperlink as one history step',
    afterClear.url === '' &&
      afterClear.command === 'remove-object-hyperlink' &&
      afterClear.commandObject === afterClear.selectedId &&
      afterClear.commandSlide === 'slide-1' &&
      afterClear.commandType === 'slide-command-effect' &&
      afterClear.selectedUrl === '' &&
      afterClearUndo.url === url &&
      afterClearUndo.selectedUrl === url &&
      afterClearRedo.url === '' &&
      afterClearRedo.selectedUrl === '' &&
      afterRestore.url === url &&
      afterRestore.selectedUrl === url,
    {
      afterClear,
      afterClearRedo,
      afterClearUndo,
      afterRestore,
    },
  )

  const importedUrl = 'https://example.com/imported-ppt-link'

  await page.eval(`((url) => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      objectHyperlink: { url },
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })(${JSON.stringify(importedUrl)})`)
  await delay(120)

  const afterHyperlinkJSONPaste = await getPPTObjectHyperlinkState(page, targetId)

  record(
    'pastes JSON object hyperlink through slide-edit command effect',
    afterHyperlinkJSONPaste.hyperlinkImportModel === 'ppt-object-hyperlink-import' &&
      afterHyperlinkJSONPaste.hyperlinkImportFormat === 'application-json-ppt-object-hyperlink' &&
      afterHyperlinkJSONPaste.hyperlinkImportSlide === 'slide-1' &&
      afterHyperlinkJSONPaste.hyperlinkImportObjects === targetId &&
      afterHyperlinkJSONPaste.hyperlinkImportFields === 'url' &&
      afterHyperlinkJSONPaste.hyperlinkImportCommands === 'update-object-hyperlink' &&
      afterHyperlinkJSONPaste.hyperlinkImportCommandFields === 'url' &&
      afterHyperlinkJSONPaste.hyperlinkImportCommandTargets === targetId &&
      afterHyperlinkJSONPaste.hyperlinkImportCommandTypes === 'slide-command-effect' &&
      afterHyperlinkJSONPaste.hyperlinkImportEnabled === 'true' &&
      afterHyperlinkJSONPaste.hyperlinkImportUrl === importedUrl &&
      afterHyperlinkJSONPaste.hyperlinkImportJsonLength > 40 &&
      afterHyperlinkJSONPaste.command === 'update-object-hyperlink' &&
      afterHyperlinkJSONPaste.commandField === 'url' &&
      afterHyperlinkJSONPaste.commandValue === importedUrl &&
      afterHyperlinkJSONPaste.selectedUrl === importedUrl &&
      afterHyperlinkJSONPaste.thumbUrl === importedUrl,
    {
      afterHyperlinkJSONPaste,
      afterRestore,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterHyperlinkJSONUndo = await getPPTObjectHyperlinkState(page, targetId)

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const afterHyperlinkJSONRedo = await getPPTObjectHyperlinkState(page, targetId)

  record(
    'undoes and redoes PPT object hyperlink JSON as one history step',
    afterHyperlinkJSONUndo.selectedUrl === url &&
      afterHyperlinkJSONRedo.selectedUrl === importedUrl,
    {
      afterHyperlinkJSONPaste,
      afterHyperlinkJSONRedo,
      afterHyperlinkJSONUndo,
    },
  )

  const standaloneUrl = 'https://example.com/standalone-ppt-link'

  await page.eval(`((url) => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({ url })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })(${JSON.stringify(standaloneUrl)})`)
  await delay(120)

  const afterStandaloneHyperlinkPaste = await getPPTObjectHyperlinkState(page, targetId)

  record(
    'pastes standalone JSON object hyperlink through slide-edit command effect',
    afterStandaloneHyperlinkPaste.hyperlinkImportModel === 'ppt-object-hyperlink-import' &&
      afterStandaloneHyperlinkPaste.hyperlinkImportFormat === 'application-json-ppt-object-hyperlink' &&
      afterStandaloneHyperlinkPaste.hyperlinkImportSlide === 'slide-1' &&
      afterStandaloneHyperlinkPaste.hyperlinkImportObjects === targetId &&
      afterStandaloneHyperlinkPaste.hyperlinkImportFields === 'url' &&
      afterStandaloneHyperlinkPaste.hyperlinkImportCommands === 'update-object-hyperlink' &&
      afterStandaloneHyperlinkPaste.hyperlinkImportCommandFields === 'url' &&
      afterStandaloneHyperlinkPaste.hyperlinkImportCommandTargets === targetId &&
      afterStandaloneHyperlinkPaste.hyperlinkImportCommandTypes === 'slide-command-effect' &&
      afterStandaloneHyperlinkPaste.hyperlinkImportEnabled === 'true' &&
      afterStandaloneHyperlinkPaste.hyperlinkImportUrl === standaloneUrl &&
      afterStandaloneHyperlinkPaste.hyperlinkImportJsonLength > 30 &&
      afterStandaloneHyperlinkPaste.command === 'update-object-hyperlink' &&
      afterStandaloneHyperlinkPaste.commandField === 'url' &&
      afterStandaloneHyperlinkPaste.commandValue === standaloneUrl &&
      afterStandaloneHyperlinkPaste.selectedUrl === standaloneUrl &&
      afterStandaloneHyperlinkPaste.thumbUrl === standaloneUrl,
    {
      afterHyperlinkJSONRedo,
      afterStandaloneHyperlinkPaste,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterStandaloneHyperlinkUndo = await getPPTObjectHyperlinkState(page, targetId)

  record(
    'undoes standalone JSON object hyperlink as one history step',
    afterStandaloneHyperlinkUndo.selectedUrl === importedUrl &&
      afterStandaloneHyperlinkUndo.thumbUrl === importedUrl,
    {
      afterStandaloneHyperlinkPaste,
      afterStandaloneHyperlinkUndo,
    },
  )

  const metadataAltText = 'AI generated card linking to PPT reference.'

  await page.eval(`((url, altText) => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      objectMetadata: {
        accessibility: { altText },
        hyperlink: { url },
      },
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })(${JSON.stringify(url)}, ${JSON.stringify(metadataAltText)})`)
  await delay(120)

  const afterMetadataPaste = await getPPTObjectHyperlinkState(page, targetId)

  record(
    'pastes JSON object metadata into selected PPT object',
    afterMetadataPaste.objectMetadataImportModel === 'ppt-object-metadata-import' &&
      afterMetadataPaste.objectMetadataImportFormat === 'application-json-ppt-object-metadata' &&
      afterMetadataPaste.objectMetadataImportSlide === 'slide-1' &&
      afterMetadataPaste.objectMetadataImportObjects === targetId &&
      afterMetadataPaste.objectMetadataImportFields === 'hyperlinkUrl altText' &&
      afterMetadataPaste.objectMetadataImportCommands ===
        'update-object-hyperlink update-object-accessibility' &&
      afterMetadataPaste.objectMetadataImportCommandFields === 'url altText' &&
      afterMetadataPaste.objectMetadataImportCommandTypes ===
        'slide-command-effect slide-command-effect' &&
      afterMetadataPaste.objectMetadataImportHyperlinkUrl === url &&
      afterMetadataPaste.objectMetadataImportAltTextPresent === 'true' &&
      afterMetadataPaste.objectMetadataImportAltTextLength === metadataAltText.length &&
      afterMetadataPaste.objectMetadataImportJsonLength > 90 &&
      afterMetadataPaste.command === 'update-object-hyperlink' &&
      afterMetadataPaste.commandField === 'url' &&
      afterMetadataPaste.commandValue === url &&
      afterMetadataPaste.accessibilityCommand === 'update-object-accessibility' &&
      afterMetadataPaste.accessibilityCommandField === 'altText' &&
      afterMetadataPaste.accessibilityCommandValue === metadataAltText &&
      afterMetadataPaste.selectedUrl === url &&
      afterMetadataPaste.selectedAltText === metadataAltText &&
      afterMetadataPaste.thumbUrl === url &&
      afterMetadataPaste.thumbAltText === metadataAltText,
    {
      afterMetadataPaste,
      afterRestore,
    },
  )

  const standaloneMetadataUrl = 'https://example.com/standalone-ppt-metadata'
  const standaloneMetadataAltText = 'Standalone metadata card description.'

  await page.eval(`((url, altText) => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      altText,
      url,
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })(${JSON.stringify(standaloneMetadataUrl)}, ${JSON.stringify(standaloneMetadataAltText)})`)
  await delay(120)

  const afterStandaloneMetadataPaste = await getPPTObjectHyperlinkState(page, targetId)

  record(
    'pastes standalone JSON object metadata into selected PPT object',
    afterStandaloneMetadataPaste.objectMetadataImportModel === 'ppt-object-metadata-import' &&
      afterStandaloneMetadataPaste.objectMetadataImportFormat === 'application-json-ppt-object-metadata' &&
      afterStandaloneMetadataPaste.objectMetadataImportSlide === 'slide-1' &&
      afterStandaloneMetadataPaste.objectMetadataImportObjects === targetId &&
      afterStandaloneMetadataPaste.objectMetadataImportFields === 'hyperlinkUrl altText' &&
      afterStandaloneMetadataPaste.objectMetadataImportCommands ===
        'update-object-hyperlink update-object-accessibility' &&
      afterStandaloneMetadataPaste.objectMetadataImportCommandFields === 'url altText' &&
      afterStandaloneMetadataPaste.objectMetadataImportCommandTypes ===
        'slide-command-effect slide-command-effect' &&
      afterStandaloneMetadataPaste.objectMetadataImportHyperlinkUrl === standaloneMetadataUrl &&
      afterStandaloneMetadataPaste.objectMetadataImportAltTextPresent === 'true' &&
      afterStandaloneMetadataPaste.objectMetadataImportAltTextLength === standaloneMetadataAltText.length &&
      afterStandaloneMetadataPaste.objectMetadataImportJsonLength > 80 &&
      afterStandaloneMetadataPaste.command === 'update-object-hyperlink' &&
      afterStandaloneMetadataPaste.commandField === 'url' &&
      afterStandaloneMetadataPaste.commandValue === standaloneMetadataUrl &&
      afterStandaloneMetadataPaste.accessibilityCommand === 'update-object-accessibility' &&
      afterStandaloneMetadataPaste.accessibilityCommandField === 'altText' &&
      afterStandaloneMetadataPaste.accessibilityCommandValue === standaloneMetadataAltText &&
      afterStandaloneMetadataPaste.selectedUrl === standaloneMetadataUrl &&
      afterStandaloneMetadataPaste.selectedAltText === standaloneMetadataAltText &&
      afterStandaloneMetadataPaste.thumbUrl === standaloneMetadataUrl &&
      afterStandaloneMetadataPaste.thumbAltText === standaloneMetadataAltText,
    {
      afterMetadataPaste,
      afterStandaloneMetadataPaste,
    },
  )

  await page.eval(`document.querySelector('button[title="Undo"]')?.click()`)
  await delay(80)

  const afterStandaloneMetadataUndo = await getPPTObjectHyperlinkState(page, targetId)

  record(
    'undoes standalone JSON object metadata as one history step',
    afterStandaloneMetadataUndo.selectedUrl === url &&
      afterStandaloneMetadataUndo.selectedAltText === metadataAltText &&
      afterStandaloneMetadataUndo.thumbUrl === url &&
      afterStandaloneMetadataUndo.thumbAltText === metadataAltText,
    {
      afterMetadataPaste,
      afterStandaloneMetadataPaste,
      afterStandaloneMetadataUndo,
    },
  )

  const metadataName = 'AI renamed layer'

  await page.eval(`((name) => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      objectName: { name },
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })(${JSON.stringify(metadataName)})`)
  await delay(120)

  const afterMetadataNamePaste = await getPPTObjectHyperlinkState(page, targetId)

  await page.eval(`document.querySelector('button[title="Undo"]')?.click()`)
  await delay(80)

  const afterMetadataNameUndo = await getPPTObjectHyperlinkState(page, targetId)

  await page.eval(`document.querySelector('button[title="Redo"]')?.click()`)
  await delay(80)

  const afterMetadataNameRedo = await getPPTObjectHyperlinkState(page, targetId)

  record(
    'pastes JSON object name through slide-edit layer pane rename command effect',
    afterMetadataNamePaste.objectMetadataImportModel === 'ppt-object-metadata-import' &&
      afterMetadataNamePaste.objectMetadataImportFormat === 'application-json-ppt-object-metadata' &&
      afterMetadataNamePaste.objectMetadataImportSlide === 'slide-1' &&
      afterMetadataNamePaste.objectMetadataImportObjects === targetId &&
      afterMetadataNamePaste.objectMetadataImportFields === 'name' &&
      afterMetadataNamePaste.objectMetadataImportCommands === 'rename-object' &&
      afterMetadataNamePaste.objectMetadataImportCommandFields === 'name' &&
      afterMetadataNamePaste.objectMetadataImportCommandTypes === 'slide-command-effect' &&
      afterMetadataNamePaste.objectMetadataImportName === metadataName &&
      afterMetadataNamePaste.objectMetadataImportJsonLength > 30 &&
      afterMetadataNamePaste.selectedName === metadataName &&
      afterMetadataNamePaste.rowName === metadataName &&
      afterMetadataNameUndo.selectedName === afterMetadataPaste.selectedName &&
      afterMetadataNameUndo.rowName === afterMetadataPaste.rowName &&
      afterMetadataNameRedo.selectedName === metadataName &&
      afterMetadataNameRedo.rowName === metadataName,
    {
      afterMetadataNamePaste,
      afterMetadataNameRedo,
      afterMetadataNameUndo,
      afterMetadataPaste,
    },
  )

  await page.eval(`document.querySelector('[data-ppt-present-start]')?.click()`)
  await delay(120)

  const preview = await page.eval(`(() => {
    const overlay = document.querySelector('[data-ppt-presentation]')
    const element = overlay?.querySelector(${JSON.stringify(`[data-ppt-element="${targetId}"]`)})

    return {
      open: !!overlay,
      url: element?.getAttribute('data-ppt-hyperlink-url') ?? '',
    }
  })()`)

  record(
    'keeps PPT object hyperlink metadata in presentation preview',
    preview.open && preview.url === url,
    preview,
  )

  await page.eval(`document.querySelector('[data-ppt-present-exit]')?.click()`)
  await delay(80)
}

async function runImageImportScenario(page) {
  const before = await getPPTImageImportState(page)

  await page.eval(`(() => {
    const input = document.querySelector('[data-ppt-image-upload-input]')
    const filesSetter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'files').set
    const dataTransfer = new DataTransfer()
    dataTransfer.items.add(${createPPTTestImageFileExpression('upload.svg', '#2563eb')})
    filesSetter.call(input, dataTransfer.files)
    input.dispatchEvent(new Event('change', { bubbles: true }))
  })()`)
  await delay(150)

  const afterUpload = await getPPTImageImportState(page)

  record(
    'inserts PPT image from file picker affordance',
      afterUpload.importExtension === 'ppt-import-extension' &&
      afterUpload.importExtensionInstallUnit === 'src/pptImportExtension' &&
      afterUpload.importExtensionClipboardActionOrder ===
        'image-file-batch image-file table-file-batch table-file fallback-html-selection-source fallback-html-image-source fallback-html-shape-source fallback-html-table-source fallback-html-text-source image-source table-source rich-text-source media-source text-source' &&
      afterUpload.importExtensionDropActionOrder ===
        'image-file-batch image-file table-file-batch table-file table-source media-source' &&
      afterUpload.imageImportModel === 'canvas-image-import' &&
      afterUpload.imageCount === before.imageCount + 1 &&
      afterUpload.selectedKind === 'image' &&
      afterUpload.selectedImageSrc.startsWith('data:image/svg+xml'),
    {
      afterUpload,
      before,
    },
  )

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    dataTransfer.items.add(${createPPTTestImageFileExpression('paste.svg', '#16a34a')})
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(150)

  const afterPaste = await getPPTImageImportState(page)

  record('pastes image file into PPT slide from clipboard event', afterPaste.imageImportModel === 'canvas-image-import' && afterPaste.imageCount === afterUpload.imageCount + 1 && afterPaste.selectedKind === 'image' && afterPaste.selectedName === 'paste.svg', {
    afterPaste,
    afterUpload,
  })

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    dataTransfer.items.add(${createPPTTestImageFileExpression('batch-a.svg', '#f97316')})
    dataTransfer.items.add(${createPPTTestImageFileExpression('batch-b.svg', '#0f766e')})
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(180)

  const afterMultiPaste = await getPPTImageImportState(page)

  record('pastes multiple image files into PPT slide from one clipboard event', afterMultiPaste.imageImportModel === 'canvas-image-import' && afterMultiPaste.imageImportFormat === 'file' && afterMultiPaste.imageImportCount === 2 && afterMultiPaste.imageImportNames.includes('batch-a.svg') && afterMultiPaste.imageImportNames.includes('batch-b.svg') && afterMultiPaste.imageCount === afterPaste.imageCount + 2 && afterMultiPaste.selectedCount === 2 && afterMultiPaste.selectedKinds === 'image image' && afterMultiPaste.selectedNames.includes('batch-a.svg') && afterMultiPaste.selectedNames.includes('batch-b.svg'), {
    afterMultiPaste,
    afterPaste,
  })

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    dataTransfer.setData('image/svg+xml', '<svg width="120" height="80" viewBox="0 0 120 80" onload="alert(1)" xmlns="http://www.w3.org/2000/svg"><script>alert(1)</script><rect width="120" height="80" fill="#7c3aed"/></svg>')
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(100)

  const afterSvgMimePaste = await getPPTImageImportState(page)

  record('pastes SVG MIME clipboard payload into PPT image element', afterSvgMimePaste.imageImportModel === 'canvas-image-import' && afterSvgMimePaste.imageImportFormat === 'svg-mime' && afterSvgMimePaste.imageImportMime === 'image/svg+xml' && afterSvgMimePaste.imageImportNaturalWidth === 120 && afterSvgMimePaste.imageImportNaturalHeight === 80 && afterSvgMimePaste.imageCount === afterMultiPaste.imageCount + 1 && afterSvgMimePaste.selectedKind === 'image' && afterSvgMimePaste.selectedName === 'clipboard.svg' && afterSvgMimePaste.selectedImageSrc.startsWith('data:image/svg+xml') && !afterSvgMimePaste.selectedImageDecoded.includes('<script') && !afterSvgMimePaste.selectedImageDecoded.includes('onload='), {
    afterMultiPaste,
    afterSvgMimePaste,
  })

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    dataTransfer.setData('text/html', '<section><svg viewBox="0 0 90 60" xmlns="http://www.w3.org/2000/svg"><circle cx="30" cy="30" r="20" fill="#0ea5e9"/></svg></section>')
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(100)

  const afterSvgHTMLPaste = await getPPTImageImportState(page)

  record('pastes HTML inline SVG clipboard payload into PPT image element', afterSvgHTMLPaste.imageImportModel === 'canvas-image-import' && afterSvgHTMLPaste.imageImportFormat === 'svg-html-inline' && afterSvgHTMLPaste.imageImportNaturalWidth === 90 && afterSvgHTMLPaste.imageImportNaturalHeight === 60 && afterSvgHTMLPaste.imageCount === afterSvgMimePaste.imageCount + 1 && afterSvgHTMLPaste.selectedKind === 'image' && afterSvgHTMLPaste.selectedImageSrc.startsWith('data:image/svg+xml') && afterSvgHTMLPaste.selectedWidth === 90 && afterSvgHTMLPaste.selectedHeight === 60, {
    afterSvgHTMLPaste,
    afterSvgMimePaste,
  })

  await page.eval(`(() => {
    const canvas = document.createElement('canvas')
    const context = canvas.getContext('2d')
    const dataTransfer = new DataTransfer()

    canvas.width = 96
    canvas.height = 48
    context.fillStyle = '#f59e0b'
    context.fillRect(0, 0, 96, 48)
    context.fillStyle = '#111827'
    context.fillRect(12, 12, 72, 24)
    dataTransfer.setData('text/html', '<figure><img alt="Copied Chart" src="' + canvas.toDataURL('image/png') + '"></figure>')
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(180)

  const afterDataImagePaste = await getPPTImageImportState(page)

  record('pastes HTML data image clipboard payload into PPT image element', afterDataImagePaste.imageImportModel === 'canvas-image-import' && afterDataImagePaste.imageImportFormat === 'data-url-html-img' && afterDataImagePaste.imageImportMime === 'image/png' && afterDataImagePaste.imageImportNaturalWidth === 96 && afterDataImagePaste.imageImportNaturalHeight === 48 && afterDataImagePaste.imageCount === afterSvgHTMLPaste.imageCount + 1 && afterDataImagePaste.selectedKind === 'image' && afterDataImagePaste.selectedName === 'Copied Chart.png' && afterDataImagePaste.selectedImageSrc.startsWith('data:image/png') && afterDataImagePaste.selectedWidth === 96 && afterDataImagePaste.selectedHeight === 48, {
    afterDataImagePaste,
    afterSvgHTMLPaste,
  })

  await page.eval(`(() => {
    window.__pptImageRichClipboardItemTypes = []
    window.__pptImageRichClipboardWriteCount = 0
    window.__pptImageRichClipboardHTML = ''
    window.__pptImageRichClipboardJSON = ''
    window.__pptImageRichClipboardPlainText = ''
    window.__pptImageRichClipboardSVG = ''

    window.ClipboardItem = class PPTImageRichClipboardItem {
      constructor(items) {
        this.items = items
        window.__pptImageRichClipboardItemTypes.push(Object.keys(items).sort())
      }
    }

    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: {
        write: async (items) => {
          window.__pptImageRichClipboardWriteCount = items.length
          const item = items[0]
          const mimeType = Object.keys(item.items)
            .find((type) => type !== 'text/html' && type !== 'text/plain' && type !== 'image/svg+xml') ?? ''

          window.__pptImageRichClipboardHTML = await item.items['text/html'].text()
          window.__pptImageRichClipboardPlainText = await item.items['text/plain'].text()
          window.__pptImageRichClipboardSVG = item.items['image/svg+xml']
            ? await item.items['image/svg+xml'].text()
            : ''
          window.__pptImageRichClipboardJSON = mimeType
            ? await item.items[mimeType].text()
            : ''
        },
      },
    })
  })()`)

  await pressKey(page, {
    code: 'KeyC',
    key: 'c',
    modifiers: 2,
    windowsVirtualKeyCode: 67,
  })
  await delay(120)

  const afterDataImageCopy = await getPPTImageImportState(page)
  const imageRichClipboardWrite = await page.eval(`(() => ({
    html: window.__pptImageRichClipboardHTML ?? '',
    itemTypes: window.__pptImageRichClipboardItemTypes?.at(-1) ?? [],
    json: window.__pptImageRichClipboardJSON ?? '',
    plainText: window.__pptImageRichClipboardPlainText ?? '',
    svg: window.__pptImageRichClipboardSVG ?? '',
    writeCount: window.__pptImageRichClipboardWriteCount ?? 0,
  }))()`)

  record(
    'copies selected PPT image with HTML image fallback through canvas rich clipboard writer',
    afterDataImageCopy.richClipboardModel === 'canvas-board-io-ppt-rich-clipboard' &&
      afterDataImageCopy.richClipboardWriteMode === 'clipboard-item' &&
      afterDataImageCopy.richClipboardSelection === afterDataImagePaste.selectedId &&
      afterDataImageCopy.richClipboardPlainTextLength > 0 &&
      afterDataImageCopy.richClipboardHTMLLength > afterDataImageCopy.richClipboardPlainTextLength &&
      imageRichClipboardWrite.writeCount === 1 &&
      imageRichClipboardWrite.itemTypes.includes(afterDataImageCopy.richClipboardJsonMimeType) &&
      imageRichClipboardWrite.itemTypes.includes('text/html') &&
      imageRichClipboardWrite.itemTypes.includes('text/plain') &&
      imageRichClipboardWrite.itemTypes.includes('image/svg+xml') &&
      imageRichClipboardWrite.html.includes('data-ppt-selection-image="true"') &&
      imageRichClipboardWrite.html.includes('<img') &&
      imageRichClipboardWrite.html.includes('src="data:image/png') &&
      imageRichClipboardWrite.html.includes(afterDataImageCopy.selectedAltText) &&
      imageRichClipboardWrite.plainText === afterDataImageCopy.selectedAltText &&
      imageRichClipboardWrite.json.includes('"kind": "interactive-os.ppt.selection"') &&
      imageRichClipboardWrite.svg.includes('<svg'),
    {
      afterDataImageCopy,
      afterDataImagePaste,
      imageRichClipboardWrite,
    },
  )

  await page.eval(`(() => {
    const svg = '<svg width="140" height="70" viewBox="0 0 140 70" xmlns="http://www.w3.org/2000/svg"><rect width="140" height="70" fill="#0891b2"/></svg>'
    window.__pptClipboardImageReadCount = 0
    window.__pptClipboardImageReadTypes = []

    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: {
        read: async () => {
          window.__pptClipboardImageReadCount += 1
          window.__pptClipboardImageReadTypes = ['image/svg+xml']

          return [{
            types: ['image/svg+xml'],
            getType: async (type) => new Blob([svg], { type }),
          }]
        },
      },
    })

    document.querySelector('[data-ppt-paste-image]')?.click()
  })()`)
  await delay(220)

  const afterClipboardImagePaste = await getPPTImageImportState(page)
  const clipboardImageReadState = await page.eval(`(() => ({
    readCount: window.__pptClipboardImageReadCount ?? 0,
    readTypes: window.__pptClipboardImageReadTypes ?? [],
  }))()`)

  record('pastes image from navigator clipboard through canvas image reader', afterClipboardImagePaste.imageImportModel === 'canvas-image-import' && afterClipboardImagePaste.imageImportFormat === 'file' && afterClipboardImagePaste.imageImportMime === 'image/svg+xml' && afterClipboardImagePaste.imageImportNaturalWidth === 140 && afterClipboardImagePaste.imageImportNaturalHeight === 70 && afterClipboardImagePaste.imageCount === afterDataImagePaste.imageCount + 1 && afterClipboardImagePaste.selectedKind === 'image' && afterClipboardImagePaste.selectedName === 'Image' && afterClipboardImagePaste.selectedImageSrc.startsWith('data:image/svg+xml') && clipboardImageReadState.readCount === 1 && clipboardImageReadState.readTypes.includes('image/svg+xml'), {
    afterClipboardImagePaste,
    afterDataImagePaste,
    clipboardImageReadState,
  })

  await page.eval(`(() => {
    const stage = document.querySelector('.ppt-stage-shell')
    if (!stage) return

    const rect = stage.getBoundingClientRect()
    const dataTransfer = new DataTransfer()
    dataTransfer.items.add(${createPPTTestImageFileExpression('drop.svg', '#dc2626')})
    stage.dispatchEvent(new DragEvent('drop', {
      bubbles: true,
      cancelable: true,
      clientX: rect.left + rect.width * 0.7,
      clientY: rect.top + rect.height * 0.4,
      dataTransfer,
    }))
  })()`)
  await delay(150)

  const afterDrop = await getPPTImageImportState(page)

  record('drops image file onto PPT stage at pointer position', afterDrop.imageImportModel === 'canvas-image-import' && afterDrop.imageImportFormat === 'file' && afterDrop.imageCount === afterClipboardImagePaste.imageCount + 1 && afterDrop.selectedKind === 'image' && afterDrop.selectedName === 'drop.svg' && afterDrop.selectedLeft > 0 && afterDrop.selectedTop >= 0, {
    afterDrop,
    afterClipboardImagePaste,
  })

  await page.eval(`(() => {
    const input = document.querySelector('[data-ppt-style-field="image-fit"]')
    input.value = 'contain'
    input.dispatchEvent(new Event('change', { bubbles: true }))
  })()`)
  await delay(50)

  const afterFit = await getPPTImageImportState(page)

  record('changes selected PPT image fit in inspector', afterFit.selectedImageFit === 'contain' && afterFit.inspectorImageFit === 'contain' && afterFit.thumbContainCount > 0, {
    afterDrop,
    afterFit,
  })
  record(
    'routes PPT image fit through slide-edit crop command-effect',
    afterFit.imageCropModel === 'slide-edit-object-image-crop' &&
      afterFit.imageCropCommand === 'update-object-image-crop' &&
      afterFit.imageCropCommandField === 'fit' &&
      afterFit.imageCropCommandObject === afterFit.selectedId &&
      afterFit.imageCropCommandType === 'slide-command-effect' &&
      afterFit.imageCropCommandValue === 'contain' &&
      afterFit.imageCropFitDescriptorAttribute === 'data-slide-object-image-crop' &&
      afterFit.imageCropFitDescriptorAttributeValue === 'contain:50,50' &&
      afterFit.imageCropFitDescriptorCommand === 'update-object-image-crop' &&
      afterFit.imageCropFitDescriptorControl === 'image-fit-select' &&
      afterFit.imageCropFitDescriptorSurface === 'object-image-crop',
    afterFit,
  )

  await page.eval(`(() => {
    const valueSetter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set
    const x = document.querySelector('[data-ppt-style-field="image-crop-x"]')
    const y = document.querySelector('[data-ppt-style-field="image-crop-y"]')

    valueSetter.call(x, '25')
    x.dispatchEvent(new Event('input', { bubbles: true }))
    x.dispatchEvent(new Event('change', { bubbles: true }))

    valueSetter.call(y, '70')
    y.dispatchEvent(new Event('input', { bubbles: true }))
    y.dispatchEvent(new Event('change', { bubbles: true }))
  })()`)
  await delay(50)

  const afterCrop = await getPPTImageImportState(page)

  record('changes selected PPT image crop position in inspector', afterCrop.inspectorCropX === 25 && afterCrop.inspectorCropY === 70 && afterCrop.selectedImagePosition === '25% 70%' && afterCrop.thumbCropXCount > 0 && afterCrop.thumbCropYCount > 0, {
    afterCrop,
    afterFit,
  })
  record(
    'routes PPT image crop position through slide-edit crop command-effect',
    afterCrop.imageCropModel === 'slide-edit-object-image-crop' &&
      afterCrop.imageCropCommand === 'update-object-image-crop' &&
      afterCrop.imageCropCommandField === 'y' &&
      afterCrop.imageCropCommandObject === afterCrop.selectedId &&
      afterCrop.imageCropCommandType === 'slide-command-effect' &&
      afterCrop.imageCropCommandValue === '70' &&
      afterCrop.imageCropXDescriptorAttributeValue === 'contain:25,70' &&
      afterCrop.imageCropYDescriptorAttributeValue === 'contain:25,70' &&
      afterCrop.imageCropXDescriptorCommand === 'update-object-image-crop' &&
      afterCrop.imageCropYDescriptorCommand === 'update-object-image-crop' &&
      afterCrop.imageCropXDescriptorControl === 'crop-position-input' &&
      afterCrop.imageCropYDescriptorControl === 'crop-position-input' &&
      afterCrop.imageCropXDescriptorSurface === 'object-image-crop' &&
      afterCrop.imageCropYDescriptorSurface === 'object-image-crop',
    afterCrop,
  )

  await page.eval(`document.querySelector('[data-ppt-image-crop-reset]')?.click()`)
  await delay(50)

  const afterReset = await getPPTImageImportState(page)

  record(
    'resets selected PPT image crop through inspector action',
    afterReset.inspectorImageFit === 'cover' &&
      afterReset.inspectorCropX === 50 &&
      afterReset.inspectorCropY === 50 &&
      afterReset.selectedImageFit === 'cover' &&
      afterReset.selectedImagePosition === '50% 50%' &&
      afterReset.imageCropResetDescriptorCommand === 'reset-object-image-crop' &&
      afterReset.imageCropResetDescriptorControl === 'image-crop-reset-button' &&
      afterReset.imageCropResetDescriptorSurface === 'object-image-crop',
    afterReset,
  )
  record(
    'routes PPT image crop reset through slide-edit crop command-effect',
    afterReset.imageCropModel === 'slide-edit-object-image-crop' &&
      afterReset.imageCropCommand === 'reset-object-image-crop' &&
      afterReset.imageCropCommandObject === afterReset.selectedId &&
      afterReset.imageCropCommandType === 'slide-command-effect' &&
      afterReset.imageCropCommandFit === 'cover' &&
      afterReset.imageCropCommandCropX === '50' &&
      afterReset.imageCropCommandCropY === '50',
    afterReset,
  )

  await page.eval(`document.querySelector('button[title="Undo"]')?.click()`)
  await delay(50)

  const afterResetUndo = await getPPTImageImportState(page)

  record(
    'undoes PPT image crop reset before export checks',
    afterResetUndo.inspectorImageFit === 'contain' &&
      afterResetUndo.inspectorCropX === 25 &&
      afterResetUndo.inspectorCropY === 70 &&
      afterResetUndo.selectedImagePosition === '25% 70%',
    {
      afterReset,
      afterResetUndo,
    },
  )

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      imageCrop: {
        crop: { x: 25, y: 70 },
        fit: 'contain',
      },
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(80)

  const afterCropPaste = await getPPTImageImportState(page)

  record(
    'pastes JSON image crop into selected PPT image',
    afterCropPaste.imageCropImportModel === 'ppt-image-crop-import' &&
      afterCropPaste.imageCropImportFormat === 'application-json-ppt-image-crop' &&
      afterCropPaste.imageCropImportSlide === 'slide-1' &&
      afterCropPaste.imageCropImportObjects === afterCropPaste.selectedId &&
      afterCropPaste.imageCropImportFields === 'fit x y' &&
      afterCropPaste.imageCropImportCommands ===
        'update-object-image-crop update-object-image-crop update-object-image-crop' &&
      afterCropPaste.imageCropImportCommandFields === 'fit x y' &&
      afterCropPaste.imageCropImportCommandTypes ===
        'slide-command-effect slide-command-effect slide-command-effect' &&
      afterCropPaste.imageCropImportFit === 'contain' &&
      afterCropPaste.imageCropImportX === '25' &&
      afterCropPaste.imageCropImportY === '70' &&
      afterCropPaste.imageCropImportJsonLength > 50 &&
      afterCropPaste.imageCropCommand === 'update-object-image-crop' &&
      afterCropPaste.imageCropCommandField === 'y' &&
      afterCropPaste.imageCropCommandObject === afterCropPaste.selectedId &&
      afterCropPaste.imageCropCommandValue === '70' &&
      afterCropPaste.inspectorImageFit === 'contain' &&
      afterCropPaste.inspectorCropX === 25 &&
      afterCropPaste.inspectorCropY === 70 &&
      afterCropPaste.selectedImageFit === 'contain' &&
      afterCropPaste.selectedImagePosition === '25% 70%',
    {
      afterCropPaste,
      afterResetUndo,
    },
  )

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      crop: { x: 40, y: 60 },
      fit: 'cover',
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(80)

  const afterStandaloneCropPaste = await getPPTImageImportState(page)

  record(
    'pastes standalone JSON image crop into selected PPT image',
    afterStandaloneCropPaste.imageCropImportModel === 'ppt-image-crop-import' &&
      afterStandaloneCropPaste.imageCropImportFormat === 'application-json-ppt-image-crop' &&
      afterStandaloneCropPaste.imageCropImportSlide === 'slide-1' &&
      afterStandaloneCropPaste.imageCropImportObjects === afterStandaloneCropPaste.selectedId &&
      afterStandaloneCropPaste.imageCropImportFields === 'fit x y' &&
      afterStandaloneCropPaste.imageCropImportCommands ===
        'update-object-image-crop update-object-image-crop update-object-image-crop' &&
      afterStandaloneCropPaste.imageCropImportCommandFields === 'fit x y' &&
      afterStandaloneCropPaste.imageCropImportCommandTypes ===
        'slide-command-effect slide-command-effect slide-command-effect' &&
      afterStandaloneCropPaste.imageCropImportFit === 'cover' &&
      afterStandaloneCropPaste.imageCropImportX === '40' &&
      afterStandaloneCropPaste.imageCropImportY === '60' &&
      afterStandaloneCropPaste.imageCropImportJsonLength > 35 &&
      afterStandaloneCropPaste.imageCropCommand === 'update-object-image-crop' &&
      afterStandaloneCropPaste.imageCropCommandField === 'y' &&
      afterStandaloneCropPaste.imageCropCommandObject === afterStandaloneCropPaste.selectedId &&
      afterStandaloneCropPaste.imageCropCommandValue === '60' &&
      afterStandaloneCropPaste.inspectorImageFit === 'cover' &&
      afterStandaloneCropPaste.inspectorCropX === 40 &&
      afterStandaloneCropPaste.inspectorCropY === 60 &&
      afterStandaloneCropPaste.selectedImageFit === 'cover' &&
      afterStandaloneCropPaste.selectedImagePosition === '40% 60%',
    {
      afterCropPaste,
      afterStandaloneCropPaste,
    },
  )

  await page.eval(`document.querySelector('button[title="Undo"]')?.click()`)
  await delay(80)

  const afterStandaloneCropUndo = await getPPTImageImportState(page)

  record(
    'undoes standalone JSON image crop as one history step',
    afterStandaloneCropUndo.inspectorImageFit === 'contain' &&
      afterStandaloneCropUndo.inspectorCropX === 25 &&
      afterStandaloneCropUndo.inspectorCropY === 70 &&
      afterStandaloneCropUndo.selectedImageFit === 'contain' &&
      afterStandaloneCropUndo.selectedImagePosition === '25% 70%',
    {
      afterCropPaste,
      afterStandaloneCropPaste,
      afterStandaloneCropUndo,
    },
  )

  await page.eval(`(() => {
    const input = document.querySelector('[data-ppt-image-replace-input]')
    const filesSetter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'files').set
    const dataTransfer = new DataTransfer()
    dataTransfer.items.add(${createPPTTestImageFileExpression('replacement.svg', '#7c3aed')})
    filesSetter.call(input, dataTransfer.files)
    input.dispatchEvent(new Event('change', { bubbles: true }))
  })()`)
  await delay(150)

  const afterReplace = await getPPTImageImportState(page)

  record(
    'replaces selected PPT image source while preserving retouch state',
    afterReplace.imageCount === afterResetUndo.imageCount &&
      afterReplace.selectedId === afterResetUndo.selectedId &&
      afterReplace.selectedName === 'replacement.svg' &&
      afterReplace.selectedAltText === 'replacement.svg' &&
      afterReplace.selectedImageSrc !== afterResetUndo.selectedImageSrc &&
      afterReplace.selectedImageSrc.startsWith('data:image/svg+xml') &&
      afterReplace.selectedLeft === afterResetUndo.selectedLeft &&
      afterReplace.selectedTop === afterResetUndo.selectedTop &&
      afterReplace.selectedWidth === afterResetUndo.selectedWidth &&
      afterReplace.selectedImageFit === 'contain' &&
      afterReplace.selectedImagePosition === '25% 70%' &&
      afterReplace.inspectorImageFit === 'contain' &&
      afterReplace.inspectorCropX === 25 &&
      afterReplace.inspectorCropY === 70,
    {
      afterReplace,
      afterResetUndo,
    },
  )
  record(
    'routes PPT image replacement through slide-edit replace command-effect',
    afterReplace.imageReplaceModel === 'slide-edit-object-image-replace' &&
      afterReplace.imageReplaceCommand === 'replace-object-image' &&
      afterReplace.imageReplaceCommandMime === 'image/svg+xml' &&
      afterReplace.imageReplaceCommandName === 'replacement.svg' &&
      afterReplace.imageReplaceCommandObject === afterReplace.selectedId &&
      afterReplace.imageReplaceCommandSrcPrefix.startsWith('data:image/svg+xml') &&
      afterReplace.imageReplaceCommandType === 'slide-command-effect' &&
      afterReplace.imageReplaceDescriptorAttribute === 'data-slide-object-image-replace' &&
      afterReplace.imageReplaceDescriptorAttributeValue === 'ready' &&
      afterReplace.imageReplaceDescriptorCommand === 'replace-object-image' &&
      afterReplace.imageReplaceDescriptorControl === 'image-source-file-input' &&
      afterReplace.imageReplaceDescriptorField === 'source' &&
      afterReplace.imageReplaceDescriptorSourceName === 'replacement.svg' &&
      afterReplace.imageReplaceDescriptorSurface === 'object-image-replace',
    afterReplace,
  )

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const svg = '<svg width="160" height="90" viewBox="0 0 160 90" xmlns="http://www.w3.org/2000/svg"><rect width="160" height="90" fill="#059669"/><text x="18" y="52" font-family="Arial" font-size="22" fill="white">JSON</text></svg>'
    const json = JSON.stringify({
      imageReplace: {
        altText: 'AI chart replacement',
        mimeType: 'image/svg+xml',
        name: 'json-replacement.svg',
        naturalHeight: 90,
        naturalWidth: 160,
        src: 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg),
      },
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const afterJSONReplace = await getPPTImageImportState(page)

  await page.eval(`document.querySelector('button[title="Undo"]')?.click()`)
  await delay(80)

  const afterJSONReplaceUndo = await getPPTImageImportState(page)

  await page.eval(`document.querySelector('button[title="Redo"]')?.click()`)
  await delay(80)

  const afterJSONReplaceRedo = await getPPTImageImportState(page)

  record(
    'pastes JSON image source into selected PPT image through slide-edit replace command-effect',
    afterJSONReplace.imageReplaceImportModel === 'ppt-image-replace-import' &&
      afterJSONReplace.imageReplaceImportFormat === 'application-json-ppt-image-replace' &&
      afterJSONReplace.imageReplaceImportSlide === 'slide-1' &&
      afterJSONReplace.imageReplaceImportObject === afterReplace.selectedId &&
      afterJSONReplace.imageReplaceImportFields ===
        'src mimeType name altText naturalWidth naturalHeight' &&
      afterJSONReplace.imageReplaceImportCommand === 'replace-object-image' &&
      afterJSONReplace.imageReplaceImportCommandType === 'slide-command-effect' &&
      afterJSONReplace.imageReplaceImportMime === 'image/svg+xml' &&
      afterJSONReplace.imageReplaceImportName === 'json-replacement.svg' &&
      afterJSONReplace.imageReplaceImportAltTextLength === 'AI chart replacement'.length &&
      afterJSONReplace.imageReplaceImportNaturalWidth === '160' &&
      afterJSONReplace.imageReplaceImportNaturalHeight === '90' &&
      afterJSONReplace.imageReplaceImportSrcPrefix.startsWith('data:image/svg+xml') &&
      afterJSONReplace.imageReplaceCommand === 'replace-object-image' &&
      afterJSONReplace.imageReplaceCommandName === 'json-replacement.svg' &&
      afterJSONReplace.imageReplaceCommandMime === 'image/svg+xml' &&
      afterJSONReplace.imageReplaceCommandObject === afterReplace.selectedId &&
      afterJSONReplace.imageCount === afterReplace.imageCount &&
      afterJSONReplace.selectedName === 'json-replacement.svg' &&
      afterJSONReplace.selectedAltText === 'AI chart replacement' &&
      afterJSONReplace.selectedImageSrc !== afterReplace.selectedImageSrc &&
      afterJSONReplace.selectedImageDecoded.includes('#059669') &&
      afterJSONReplace.selectedImageFit === 'contain' &&
      afterJSONReplace.selectedImagePosition === '25% 70%' &&
      afterJSONReplaceUndo.selectedName === afterReplace.selectedName &&
      afterJSONReplaceUndo.selectedAltText === afterReplace.selectedAltText &&
      afterJSONReplaceUndo.selectedImageSrc === afterReplace.selectedImageSrc &&
      afterJSONReplaceRedo.selectedName === 'json-replacement.svg' &&
      afterJSONReplaceRedo.selectedAltText === 'AI chart replacement' &&
      afterJSONReplaceRedo.selectedImageDecoded.includes('#059669'),
    {
      afterJSONReplace,
      afterJSONReplaceRedo,
      afterJSONReplaceUndo,
      afterReplace,
    },
  )

  await page.eval(`document.querySelector('button[title="Undo"]')?.click()`)
  await delay(80)

  const afterJSONReplaceRestore = await getPPTImageImportState(page)

  record(
    'restores PPT image source after JSON image replacement probe',
    afterJSONReplaceRestore.selectedName === afterReplace.selectedName &&
      afterJSONReplaceRestore.selectedAltText === afterReplace.selectedAltText &&
      afterJSONReplaceRestore.selectedImageSrc === afterReplace.selectedImageSrc &&
      afterJSONReplaceRestore.selectedImageFit === afterReplace.selectedImageFit &&
      afterJSONReplaceRestore.selectedImagePosition === afterReplace.selectedImagePosition,
    {
      afterJSONReplaceRedo,
      afterJSONReplaceRestore,
      afterReplace,
    },
  )

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const svg = '<svg width="140" height="80" viewBox="0 0 140 80" xmlns="http://www.w3.org/2000/svg"><rect width="140" height="80" fill="#dc2626"/><text x="14" y="47" font-family="Arial" font-size="20" fill="white">DIRECT</text></svg>'
    const json = JSON.stringify({
      altText: 'Standalone image source',
      dataUrl: 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg),
      mimeType: 'image/svg+xml',
      name: 'standalone-json-replacement.svg',
      naturalHeight: 80,
      naturalWidth: 140,
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const afterStandaloneJSONReplace = await getPPTImageImportState(page)

  await page.eval(`document.querySelector('button[title="Undo"]')?.click()`)
  await delay(80)

  const afterStandaloneJSONReplaceUndo = await getPPTImageImportState(page)

  await page.eval(`document.querySelector('button[title="Redo"]')?.click()`)
  await delay(80)

  const afterStandaloneJSONReplaceRedo = await getPPTImageImportState(page)

  record(
    'pastes standalone JSON image source into selected PPT image through slide-edit replace command-effect',
    afterStandaloneJSONReplace.imageReplaceImportModel === 'ppt-image-replace-import' &&
      afterStandaloneJSONReplace.imageReplaceImportFormat === 'application-json-ppt-image-replace' &&
      afterStandaloneJSONReplace.imageReplaceImportSlide === 'slide-1' &&
      afterStandaloneJSONReplace.imageReplaceImportObject === afterReplace.selectedId &&
      afterStandaloneJSONReplace.imageReplaceImportFields ===
        'src mimeType name altText naturalWidth naturalHeight' &&
      afterStandaloneJSONReplace.imageReplaceImportCommand === 'replace-object-image' &&
      afterStandaloneJSONReplace.imageReplaceImportCommandType === 'slide-command-effect' &&
      afterStandaloneJSONReplace.imageReplaceImportMime === 'image/svg+xml' &&
      afterStandaloneJSONReplace.imageReplaceImportName === 'standalone-json-replacement.svg' &&
      afterStandaloneJSONReplace.imageReplaceImportAltTextLength === 'Standalone image source'.length &&
      afterStandaloneJSONReplace.imageReplaceImportNaturalWidth === '140' &&
      afterStandaloneJSONReplace.imageReplaceImportNaturalHeight === '80' &&
      afterStandaloneJSONReplace.imageReplaceImportSrcPrefix.startsWith('data:image/svg+xml') &&
      afterStandaloneJSONReplace.imageReplaceCommand === 'replace-object-image' &&
      afterStandaloneJSONReplace.imageReplaceCommandName === 'standalone-json-replacement.svg' &&
      afterStandaloneJSONReplace.imageReplaceCommandMime === 'image/svg+xml' &&
      afterStandaloneJSONReplace.imageReplaceCommandObject === afterReplace.selectedId &&
      afterStandaloneJSONReplace.imageCount === afterReplace.imageCount &&
      afterStandaloneJSONReplace.selectedName === 'standalone-json-replacement.svg' &&
      afterStandaloneJSONReplace.selectedAltText === 'Standalone image source' &&
      afterStandaloneJSONReplace.selectedImageSrc !== afterReplace.selectedImageSrc &&
      afterStandaloneJSONReplace.selectedImageDecoded.includes('#dc2626') &&
      afterStandaloneJSONReplace.selectedImageFit === 'contain' &&
      afterStandaloneJSONReplace.selectedImagePosition === '25% 70%' &&
      afterStandaloneJSONReplaceUndo.selectedName === afterReplace.selectedName &&
      afterStandaloneJSONReplaceUndo.selectedAltText === afterReplace.selectedAltText &&
      afterStandaloneJSONReplaceUndo.selectedImageSrc === afterReplace.selectedImageSrc &&
      afterStandaloneJSONReplaceRedo.selectedName === 'standalone-json-replacement.svg' &&
      afterStandaloneJSONReplaceRedo.selectedAltText === 'Standalone image source' &&
      afterStandaloneJSONReplaceRedo.selectedImageDecoded.includes('#dc2626'),
    {
      afterReplace,
      afterStandaloneJSONReplace,
      afterStandaloneJSONReplaceRedo,
      afterStandaloneJSONReplaceUndo,
    },
  )

  await page.eval(`document.querySelector('button[title="Undo"]')?.click()`)
  await delay(80)

  const afterStandaloneJSONReplaceRestore = await getPPTImageImportState(page)

  record(
    'restores PPT image source after standalone JSON image replacement probe',
    afterStandaloneJSONReplaceRestore.selectedName === afterReplace.selectedName &&
      afterStandaloneJSONReplaceRestore.selectedAltText === afterReplace.selectedAltText &&
      afterStandaloneJSONReplaceRestore.selectedImageSrc === afterReplace.selectedImageSrc &&
      afterStandaloneJSONReplaceRestore.selectedImageFit === afterReplace.selectedImageFit &&
      afterStandaloneJSONReplaceRestore.selectedImagePosition === afterReplace.selectedImagePosition,
    {
      afterReplace,
      afterStandaloneJSONReplaceRedo,
      afterStandaloneJSONReplaceRestore,
    },
  )

  await page.eval(`(() => {
    window.__pptRetouchedImageRichClipboardItemTypes = []
    window.__pptRetouchedImageRichClipboardWriteCount = 0
    window.__pptRetouchedImageRichClipboardHTML = ''
    window.__pptRetouchedImageRichClipboardPlainText = ''
    window.__pptRetouchedImageRichClipboardJSON = ''

    window.ClipboardItem = class PPTRetouchedImageRichClipboardItem {
      constructor(items) {
        this.items = items
        window.__pptRetouchedImageRichClipboardItemTypes.push(Object.keys(items).sort())
      }
    }

    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: {
        write: async (items) => {
          window.__pptRetouchedImageRichClipboardWriteCount = items.length
          const item = items[0]
          const mimeType = Object.keys(item.items)
            .find((type) => type !== 'text/html' && type !== 'text/plain' && type !== 'image/svg+xml') ?? ''

          window.__pptRetouchedImageRichClipboardHTML = await item.items['text/html'].text()
          window.__pptRetouchedImageRichClipboardPlainText = await item.items['text/plain'].text()
          window.__pptRetouchedImageRichClipboardJSON = mimeType
            ? await item.items[mimeType].text()
            : ''
        },
      },
    })
  })()`)

  await pressKey(page, {
    code: 'KeyC',
    key: 'c',
    modifiers: 2,
    windowsVirtualKeyCode: 67,
  })
  await delay(120)

  const afterRetouchedImageCopy = await getPPTImageImportState(page)
  const retouchedImageRichClipboardWrite = await page.eval(`(() => ({
    html: window.__pptRetouchedImageRichClipboardHTML ?? '',
    itemTypes: window.__pptRetouchedImageRichClipboardItemTypes?.at(-1) ?? [],
    json: window.__pptRetouchedImageRichClipboardJSON ?? '',
    plainText: window.__pptRetouchedImageRichClipboardPlainText ?? '',
    writeCount: window.__pptRetouchedImageRichClipboardWriteCount ?? 0,
  }))()`)

  record(
    'copies retouched PPT image with positioned fit crop fallback HTML',
    afterRetouchedImageCopy.richClipboardSelection === afterReplace.selectedId &&
      retouchedImageRichClipboardWrite.writeCount === 1 &&
      retouchedImageRichClipboardWrite.itemTypes.includes(afterRetouchedImageCopy.richClipboardJsonMimeType) &&
      retouchedImageRichClipboardWrite.html.includes('data-ppt-selection-image="true"') &&
      retouchedImageRichClipboardWrite.html.includes('data-ppt-selection-image-fit="contain"') &&
      retouchedImageRichClipboardWrite.html.includes('data-ppt-selection-image-crop-x="25"') &&
      retouchedImageRichClipboardWrite.html.includes('data-ppt-selection-image-crop-y="70"') &&
      retouchedImageRichClipboardWrite.html.includes('data-ppt-selection-x="') &&
      retouchedImageRichClipboardWrite.html.includes('data-ppt-selection-w="') &&
      retouchedImageRichClipboardWrite.html.includes('replacement.svg') &&
      retouchedImageRichClipboardWrite.json.includes('"kind": "interactive-os.ppt.selection"'),
    {
      afterReplace,
      afterRetouchedImageCopy,
      retouchedImageRichClipboardWrite,
    },
  )

  await page.eval(`((html, plainText) => {
    const dataTransfer = new DataTransfer()
    const fallbackHTML = html.replace(/<script\\b[\\s\\S]*?<\\/script>/gi, '')

    dataTransfer.setData('text/html', fallbackHTML)
    dataTransfer.setData('text/plain', plainText)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })(${JSON.stringify(retouchedImageRichClipboardWrite.html)}, ${JSON.stringify(retouchedImageRichClipboardWrite.plainText)})`)
  await delay(180)

  const afterImageFallbackHTMLPaste = await getPPTImageImportState(page)

  record(
    'pastes PPT image fallback HTML as editable image preserving fit crop',
    afterImageFallbackHTMLPaste.imageCount === afterRetouchedImageCopy.imageCount + 1 &&
      afterImageFallbackHTMLPaste.selectedKind === 'image' &&
      afterImageFallbackHTMLPaste.selectedName === 'replacement.svg' &&
      afterImageFallbackHTMLPaste.selectedAltText === 'replacement.svg' &&
      afterImageFallbackHTMLPaste.selectedImageFit === 'contain' &&
      afterImageFallbackHTMLPaste.selectedImagePosition === '25% 70%' &&
      afterImageFallbackHTMLPaste.selectedWidth === afterRetouchedImageCopy.selectedWidth &&
      afterImageFallbackHTMLPaste.selectedHeight === afterRetouchedImageCopy.selectedHeight &&
      afterImageFallbackHTMLPaste.fallbackHTMLImportModel === 'ppt-fallback-html-import' &&
      afterImageFallbackHTMLPaste.fallbackHTMLImportFormat === 'text-html-ppt-fallback' &&
      afterImageFallbackHTMLPaste.fallbackHTMLImportKind === 'image' &&
      afterImageFallbackHTMLPaste.fallbackHTMLImportSourceObject === afterReplace.selectedId,
    {
      afterImageFallbackHTMLPaste,
      afterRetouchedImageCopy,
    },
  )

  const beforeResize = await page.eval(`(() => {
    const selected = document.querySelector('[data-selected="true"]')
    const handle = document.querySelector('button[aria-label="Resize e"]').getBoundingClientRect()

    return {
      handleX: handle.left + handle.width / 2,
      handleY: handle.top + handle.height / 2,
      width: parseFloat(selected.style.width),
    }
  })()`)

  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mousePressed',
    x: beforeResize.handleX,
    y: beforeResize.handleY,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    type: 'mouseMoved',
    x: beforeResize.handleX + 40,
    y: beforeResize.handleY,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mouseReleased',
    x: beforeResize.handleX + 40,
    y: beforeResize.handleY,
  })
  await delay(50)

  const afterResize = await getPPTImageImportState(page)

  record('resizes inserted PPT image with existing selection handles', afterResize.selectedWidth > beforeResize.width && afterResize.selectedKind === 'image', {
    afterResize,
    beforeResize,
  })
}

async function runObjectAltTextScenario(page) {
  const imageId = await page.eval(`(() => [...document.querySelectorAll('[data-kind="image"]')].at(-1)?.getAttribute('data-ppt-element') ?? '')()`)
  await selectPPTLayerRows(page, [imageId])
  await delay(80)

  const initial = await getPPTObjectAltTextState(page, imageId)

  record(
    'renders PPT object alt text control in inspector',
    imageId.length > 0 &&
      initial.selectedId === imageId &&
      initial.selectedKind === 'image' &&
      initial.altText === '' &&
      initial.selectedAltText === '' &&
      initial.thumbAltText === '' &&
      initial.imageAlt.length > 0,
    initial,
  )

  await page.eval(`((altText) => {
    const textarea = document.querySelector('[data-ppt-style-field="alt-text"]')
    const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set

    setter.call(textarea, altText)
    textarea.dispatchEvent(new Event('input', { bubbles: true }))
    textarea.dispatchEvent(new Event('change', { bubbles: true }))
    textarea.blur()
  })(${JSON.stringify(PPT_OBJECT_ALT_TEXT)})`)
  await delay(120)

  const afterAltText = await getPPTObjectAltTextState(page, imageId)

  record(
    'updates PPT object alt text metadata from inspector',
    afterAltText.altText === PPT_OBJECT_ALT_TEXT &&
      afterAltText.descriptorSurface === 'object-accessibility' &&
      afterAltText.descriptorCommand === 'update-object-accessibility' &&
      afterAltText.descriptorControl === 'multiline-text' &&
      afterAltText.descriptorAttribute === 'data-slide-object-accessibility' &&
      afterAltText.descriptorAltText === PPT_OBJECT_ALT_TEXT &&
      afterAltText.descriptorDescribed === 'true' &&
      afterAltText.model === 'slide-edit-object-accessibility' &&
      afterAltText.command === 'update-object-accessibility' &&
      afterAltText.commandField === 'altText' &&
      afterAltText.commandObject === imageId &&
      afterAltText.commandSlide === 'slide-1' &&
      afterAltText.commandType === 'slide-command-effect' &&
      afterAltText.commandValue === PPT_OBJECT_ALT_TEXT &&
      afterAltText.selectedAltText === PPT_OBJECT_ALT_TEXT &&
      afterAltText.thumbAltText === PPT_OBJECT_ALT_TEXT &&
      afterAltText.imageAlt === PPT_OBJECT_ALT_TEXT,
    {
      afterAltText,
      initial,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterUndo = await getPPTObjectAltTextState(page, imageId)

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const afterRedo = await getPPTObjectAltTextState(page, imageId)

  record(
    'undoes and redoes PPT object alt text as one history step',
    afterUndo.altText === '' &&
      afterUndo.selectedAltText === '' &&
      afterUndo.imageAlt === initial.imageAlt &&
      afterRedo.altText === PPT_OBJECT_ALT_TEXT &&
      afterRedo.selectedAltText === PPT_OBJECT_ALT_TEXT &&
      afterRedo.imageAlt === PPT_OBJECT_ALT_TEXT,
    {
      afterRedo,
      afterUndo,
    },
  )

  await page.eval(`(() => {
    const textarea = document.querySelector('[data-ppt-style-field="alt-text"]')
    const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set

    setter.call(textarea, '')
    textarea.dispatchEvent(new Event('input', { bubbles: true }))
    textarea.dispatchEvent(new Event('change', { bubbles: true }))
    textarea.blur()
  })()`)
  await delay(120)

  const afterClear = await getPPTObjectAltTextState(page, imageId)

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterClearUndo = await getPPTObjectAltTextState(page, imageId)

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const afterClearRedo = await getPPTObjectAltTextState(page, imageId)

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterRestore = await getPPTObjectAltTextState(page, imageId)

  record(
    'removes and restores PPT object alt text as one history step',
    afterClear.altText === '' &&
      afterClear.command === 'remove-object-alt-text' &&
      afterClear.commandObject === imageId &&
      afterClear.commandSlide === 'slide-1' &&
      afterClear.commandType === 'slide-command-effect' &&
      afterClear.selectedAltText === '' &&
      afterClear.imageAlt === initial.imageAlt &&
      afterClearUndo.altText === PPT_OBJECT_ALT_TEXT &&
      afterClearUndo.selectedAltText === PPT_OBJECT_ALT_TEXT &&
      afterClearRedo.altText === '' &&
      afterClearRedo.selectedAltText === '' &&
      afterRestore.altText === PPT_OBJECT_ALT_TEXT &&
      afterRestore.selectedAltText === PPT_OBJECT_ALT_TEXT &&
      afterRestore.imageAlt === PPT_OBJECT_ALT_TEXT,
    {
      afterClear,
      afterClearRedo,
      afterClearUndo,
      afterRestore,
    },
  )

  const jsonAltText = 'JSON chart replacement description'

  await page.eval(`((altText) => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      objectAccessibility: {
        altText,
      },
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })(${JSON.stringify(jsonAltText)})`)
  await delay(120)

  const afterJSONPaste = await getPPTObjectAltTextState(page, imageId)

  record(
    'pastes JSON object accessibility alt text through slide-edit command effect',
    afterJSONPaste.importModel === 'ppt-object-accessibility-import' &&
      afterJSONPaste.importFormat === 'application-json-ppt-object-accessibility' &&
      afterJSONPaste.importSlide === 'slide-1' &&
      afterJSONPaste.importObjects === imageId &&
      afterJSONPaste.importFields === 'altText' &&
      afterJSONPaste.importCommands === 'update-object-accessibility' &&
      afterJSONPaste.importCommandFields === 'altText' &&
      afterJSONPaste.importCommandTargets === imageId &&
      afterJSONPaste.importCommandTypes === 'slide-command-effect' &&
      afterJSONPaste.importCommandValues === jsonAltText &&
      afterJSONPaste.importAltTextLength === jsonAltText.length &&
      afterJSONPaste.importAltTextPresent === 'true' &&
      afterJSONPaste.importJsonLength > 60 &&
      afterJSONPaste.command === 'update-object-accessibility' &&
      afterJSONPaste.commandField === 'altText' &&
      afterJSONPaste.commandObject === imageId &&
      afterJSONPaste.commandSlide === 'slide-1' &&
      afterJSONPaste.commandType === 'slide-command-effect' &&
      afterJSONPaste.commandValue === jsonAltText &&
      afterJSONPaste.altText === jsonAltText &&
      afterJSONPaste.descriptorAltText === jsonAltText &&
      afterJSONPaste.selectedAltText === jsonAltText &&
      afterJSONPaste.thumbAltText === jsonAltText &&
      afterJSONPaste.imageAlt === jsonAltText,
    {
      afterJSONPaste,
      afterRestore,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterJSONUndo = await getPPTObjectAltTextState(page, imageId)

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(80)

  const afterJSONRedo = await getPPTObjectAltTextState(page, imageId)

  record(
    'undoes and redoes PPT object accessibility JSON as one history step',
    afterJSONUndo.altText === PPT_OBJECT_ALT_TEXT &&
      afterJSONUndo.selectedAltText === PPT_OBJECT_ALT_TEXT &&
      afterJSONRedo.altText === jsonAltText &&
      afterJSONRedo.selectedAltText === jsonAltText,
    {
      afterJSONRedo,
      afterJSONUndo,
    },
  )

  const standaloneAltText = 'Standalone JSON object alt text'

  await page.eval(`((altText) => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({ altText })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })(${JSON.stringify(standaloneAltText)})`)
  await delay(120)

  const afterStandaloneAltTextPaste = await getPPTObjectAltTextState(page, imageId)

  record(
    'pastes standalone JSON object accessibility alt text through slide-edit command effect',
    afterStandaloneAltTextPaste.importModel === 'ppt-object-accessibility-import' &&
      afterStandaloneAltTextPaste.importFormat === 'application-json-ppt-object-accessibility' &&
      afterStandaloneAltTextPaste.importSlide === 'slide-1' &&
      afterStandaloneAltTextPaste.importObjects === imageId &&
      afterStandaloneAltTextPaste.importFields === 'altText' &&
      afterStandaloneAltTextPaste.importCommands === 'update-object-accessibility' &&
      afterStandaloneAltTextPaste.importCommandFields === 'altText' &&
      afterStandaloneAltTextPaste.importCommandTargets === imageId &&
      afterStandaloneAltTextPaste.importCommandTypes === 'slide-command-effect' &&
      afterStandaloneAltTextPaste.importCommandValues === standaloneAltText &&
      afterStandaloneAltTextPaste.importAltTextLength === standaloneAltText.length &&
      afterStandaloneAltTextPaste.importAltTextPresent === 'true' &&
      afterStandaloneAltTextPaste.importJsonLength > 20 &&
      afterStandaloneAltTextPaste.command === 'update-object-accessibility' &&
      afterStandaloneAltTextPaste.commandField === 'altText' &&
      afterStandaloneAltTextPaste.commandObject === imageId &&
      afterStandaloneAltTextPaste.commandSlide === 'slide-1' &&
      afterStandaloneAltTextPaste.commandType === 'slide-command-effect' &&
      afterStandaloneAltTextPaste.commandValue === standaloneAltText &&
      afterStandaloneAltTextPaste.altText === standaloneAltText &&
      afterStandaloneAltTextPaste.descriptorAltText === standaloneAltText &&
      afterStandaloneAltTextPaste.selectedAltText === standaloneAltText &&
      afterStandaloneAltTextPaste.thumbAltText === standaloneAltText &&
      afterStandaloneAltTextPaste.imageAlt === standaloneAltText,
    {
      afterJSONRedo,
      afterStandaloneAltTextPaste,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterStandaloneAltTextUndo = await getPPTObjectAltTextState(page, imageId)

  record(
    'undoes standalone JSON object accessibility alt text as one history step',
    afterStandaloneAltTextUndo.altText === jsonAltText &&
      afterStandaloneAltTextUndo.selectedAltText === jsonAltText &&
      afterStandaloneAltTextUndo.imageAlt === jsonAltText,
    {
      afterStandaloneAltTextPaste,
      afterStandaloneAltTextUndo,
    },
  )

  await page.eval(`((altText) => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      objectAccessibility: {
        altText,
      },
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })(${JSON.stringify(PPT_OBJECT_ALT_TEXT)})`)
  await delay(120)

  await page.eval(`document.querySelector('[data-ppt-present-start]')?.click()`)
  await delay(120)

  const preview = await page.eval(`((id) => {
    const overlay = document.querySelector('[data-ppt-presentation]')
    const element = overlay?.querySelector(\`[data-ppt-element="\${id}"]\`)
    const image = element?.querySelector('img')

    return {
      altText: element?.getAttribute('data-ppt-alt-text') ?? '',
      imageAlt: image?.getAttribute('alt') ?? '',
      open: !!overlay,
    }
  })(${JSON.stringify(imageId)})`)

  record(
    'keeps PPT object alt text metadata in presentation preview',
    preview.open &&
      preview.altText === PPT_OBJECT_ALT_TEXT &&
      preview.imageAlt === PPT_OBJECT_ALT_TEXT,
    preview,
  )

  await page.eval(`document.querySelector('[data-ppt-present-exit]')?.click()`)
  await delay(80)
}

async function runTableImportScenario(page) {
  const before = await getPPTTableState(page)

  await page.eval(`document.querySelector('[data-ppt-insert-table]')?.click()`)
  await delay(80)

  const afterToolbarInsert = await getPPTTableState(page)

  record('inserts PPT table from toolbar affordance', afterToolbarInsert.tableCount === before.tableCount + 1 && afterToolbarInsert.selectedKind === 'table' && afterToolbarInsert.selectedRows === 3 && afterToolbarInsert.selectedCols === 3 && afterToolbarInsert.inspectorValue.includes('Metric'), {
    afterToolbarInsert,
    before,
  })

  await page.eval(`(() => {
    const textarea = document.querySelector('[data-ppt-style-field="table-data"]')
    const valueSetter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set

    valueSetter.call(textarea, 'Metric\\tQ1\\tQ2\\nRevenue\\t10\\t12\\nMargin\\t42%\\t45%')
    textarea.dispatchEvent(new Event('input', { bubbles: true }))
    textarea.dispatchEvent(new Event('change', { bubbles: true }))
    textarea.blur()
  })()`)
  await delay(80)

  const afterInspectorEdit = await getPPTTableState(page)

  record('retouches selected PPT table data in inspector TSV field', afterInspectorEdit.selectedKind === 'table' && afterInspectorEdit.selectedRows === 3 && afterInspectorEdit.selectedCols === 3 && afterInspectorEdit.cellTexts.includes('Revenue') && afterInspectorEdit.cellTexts.includes('45%') && afterInspectorEdit.inspectorSize === '3 x 3' && afterInspectorEdit.thumbTableCount > before.thumbTableCount, {
    afterInspectorEdit,
    afterToolbarInsert,
  })

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      tableRows: {
        columns: ['Metric', 'Actual', 'Target'],
        rows: [
          { Metric: 'Revenue', Actual: '12', Target: '15' },
          { Metric: 'Margin', Actual: '45%', Target: '50%' },
          { Metric: 'Retention', Actual: '91%', Target: '94%' },
        ],
      },
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(100)

  const afterJSONRowsPaste = await getPPTTableState(page)

  record(
    'pastes JSON table rows into selected PPT table',
    afterJSONRowsPaste.tableRowsImportModel === 'ppt-table-rows-import' &&
      afterJSONRowsPaste.tableRowsImportFormat === 'application-json-ppt-table-rows' &&
      afterJSONRowsPaste.tableRowsImportObjects === afterInspectorEdit.selectedId &&
      afterJSONRowsPaste.tableRowsImportTargets === afterInspectorEdit.selectedId &&
      afterJSONRowsPaste.tableRowsImportRows === 4 &&
      afterJSONRowsPaste.tableRowsImportCols === 3 &&
      afterJSONRowsPaste.tableRowsImportJsonLength > 160 &&
      afterJSONRowsPaste.tableCount === afterInspectorEdit.tableCount &&
      afterJSONRowsPaste.selectedId === afterInspectorEdit.selectedId &&
      afterJSONRowsPaste.selectedRows === 4 &&
      afterJSONRowsPaste.selectedCols === 3 &&
      afterJSONRowsPaste.inspectorSize === '4 x 3' &&
      afterJSONRowsPaste.cellTexts.includes('Retention') &&
      afterJSONRowsPaste.cellTexts.includes('94%'),
    {
      afterInspectorEdit,
      afterJSONRowsPaste,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterJSONRowsUndo = await getPPTTableState(page)

  record(
    'undoes PPT table rows JSON paste as one history step',
    afterJSONRowsUndo.tableCount === afterInspectorEdit.tableCount &&
      afterJSONRowsUndo.selectedId === afterInspectorEdit.selectedId &&
      afterJSONRowsUndo.selectedRows === 3 &&
      afterJSONRowsUndo.selectedCols === 3 &&
      afterJSONRowsUndo.cellTexts.includes('Revenue') &&
      afterJSONRowsUndo.cellTexts.includes('45%') &&
      !afterJSONRowsUndo.cellTexts.includes('Retention'),
    {
      afterInspectorEdit,
      afterJSONRowsPaste,
      afterJSONRowsUndo,
    },
  )

  await page.eval(`(() => {
    window.__pptTableClipboardItemTypes = []
    window.__pptTableClipboardWriteCount = 0
    window.__pptTableClipboardHTML = ''
    window.__pptTableClipboardJSON = ''
    window.__pptTableClipboardPlainText = ''

    window.ClipboardItem = class PPTTableClipboardItem {
      constructor(items) {
        this.items = items
        window.__pptTableClipboardItemTypes.push(Object.keys(items).sort())
      }
    }

    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: {
        write: async (items) => {
          window.__pptTableClipboardWriteCount = items.length
          const item = items[0]
          const mimeType = Object.keys(item.items)
            .find((type) => type !== 'text/html' && type !== 'text/plain') ?? ''

          window.__pptTableClipboardHTML = await item.items['text/html'].text()
          window.__pptTableClipboardPlainText = await item.items['text/plain'].text()
          window.__pptTableClipboardJSON = mimeType
            ? await item.items[mimeType].text()
            : ''
        },
      },
    })
  })()`)
  await page.eval(`document.querySelector('[data-ppt-copy-table]')?.click()`)
  await delay(120)

  const afterCopyTable = await getPPTTableState(page)
  const tableClipboardWrite = await page.eval(`(() => ({
    html: window.__pptTableClipboardHTML ?? '',
    itemTypes: window.__pptTableClipboardItemTypes?.at(-1) ?? [],
    json: window.__pptTableClipboardJSON ?? '',
    plainText: window.__pptTableClipboardPlainText ?? '',
    writeCount: window.__pptTableClipboardWriteCount ?? 0,
  }))()`)

  record(
    'copies selected PPT table as HTML and TSV through canvas rich clipboard writer',
    afterCopyTable.tableClipboardModel === 'canvas-rich-table-clipboard' &&
      afterCopyTable.tableClipboardObject === afterInspectorEdit.selectedId &&
      afterCopyTable.tableClipboardRows === 3 &&
      afterCopyTable.tableClipboardCols === 3 &&
      afterCopyTable.tableClipboardWriteMode === 'clipboard-item' &&
      tableClipboardWrite.writeCount === 1 &&
      tableClipboardWrite.itemTypes.includes(afterCopyTable.tableClipboardJsonMimeType) &&
      tableClipboardWrite.itemTypes.includes('text/html') &&
      tableClipboardWrite.itemTypes.includes('text/plain') &&
      tableClipboardWrite.html.includes('<table') &&
      tableClipboardWrite.html.includes('data-ppt-selection-table="true"') &&
      tableClipboardWrite.html.includes('data-ppt-selection-x="') &&
      tableClipboardWrite.html.includes('data-ppt-selection-w="') &&
      tableClipboardWrite.html.includes('<th>Metric</th>') &&
      tableClipboardWrite.html.includes('<td>45%</td>') &&
      tableClipboardWrite.plainText.includes('Revenue\t10\t12') &&
      tableClipboardWrite.plainText.includes('Margin\t42%\t45%') &&
      tableClipboardWrite.json.includes('"kind": "interactive-os.ppt.table-export"'),
    {
      afterCopyTable,
      afterInspectorEdit,
      tableClipboardWrite,
    },
  )

  await page.eval(`((html, plainText) => {
    const dataTransfer = new DataTransfer()
    const fallbackHTML = html.replace(/<script\\b[\\s\\S]*?<\\/script>/gi, '')

    dataTransfer.setData('text/html', fallbackHTML)
    dataTransfer.setData('text/plain', plainText)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })(${JSON.stringify(tableClipboardWrite.html)}, ${JSON.stringify(tableClipboardWrite.plainText)})`)
  await delay(120)

  const afterTableFallbackHTMLPaste = await getPPTTableState(page)

  record(
    'pastes PPT table fallback HTML as editable table preserving size',
    afterTableFallbackHTMLPaste.tableCount === afterCopyTable.tableCount + 1 &&
      afterTableFallbackHTMLPaste.selectedKind === 'table' &&
      afterTableFallbackHTMLPaste.selectedRows === 3 &&
      afterTableFallbackHTMLPaste.selectedCols === 3 &&
      afterTableFallbackHTMLPaste.selectedWidth === afterCopyTable.selectedWidth &&
      afterTableFallbackHTMLPaste.selectedHeight === afterCopyTable.selectedHeight &&
      afterTableFallbackHTMLPaste.cellTexts.includes('Revenue') &&
      afterTableFallbackHTMLPaste.cellTexts.includes('45%') &&
      afterTableFallbackHTMLPaste.fallbackHTMLImportModel === 'ppt-fallback-html-import' &&
      afterTableFallbackHTMLPaste.fallbackHTMLImportFormat === 'text-html-ppt-fallback' &&
      afterTableFallbackHTMLPaste.fallbackHTMLImportKind === 'table' &&
      afterTableFallbackHTMLPaste.fallbackHTMLImportSourceObject === afterInspectorEdit.selectedId,
    {
      afterCopyTable,
      afterInspectorEdit,
      afterTableFallbackHTMLPaste,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterTableFallbackUndo = await getPPTTableState(page)

  record(
    'undoes PPT table fallback HTML paste before table scenario continues',
    afterTableFallbackUndo.tableCount === afterCopyTable.tableCount,
    {
      afterCopyTable,
      afterTableFallbackHTMLPaste,
      afterTableFallbackUndo,
    },
  )

  await page.eval(`document.querySelector('[data-ppt-insert-table]')?.click()`)
  await delay(100)

  const afterPaletteInsert = await getPPTTableState(page)

  record('inserts PPT table from toolbar affordance', afterPaletteInsert.tableCount === afterInspectorEdit.tableCount + 1 && afterPaletteInsert.selectedKind === 'table' && afterPaletteInsert.selectedRows === 3 && afterPaletteInsert.selectedCols === 3 && !afterPaletteInsert.paletteOpen, {
    afterInspectorEdit,
    afterPaletteInsert,
  })

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    dataTransfer.setData('text/tab-separated-values', 'Name\\tValue\\nUsers\\t120\\nARR\\t$1M')
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(100)

  const afterPaste = await getPPTTableState(page)

  record('pastes TSV clipboard data into PPT table element', afterPaste.tableImportModel === 'canvas-table-import' && afterPaste.tableImportFormat === 'text-tsv' && afterPaste.tableCount === afterPaletteInsert.tableCount + 1 && afterPaste.selectedKind === 'table' && afterPaste.selectedRows === 3 && afterPaste.selectedCols === 2 && afterPaste.cellTexts.includes('Users') && afterPaste.cellTexts.includes('$1M'), {
    afterPaletteInsert,
    afterPaste,
  })

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()

    dataTransfer.setData('text/html', '<table><thead><tr><th>Plan</th><th>Owner</th><th>Status</th></tr></thead><tbody><tr><td>Outline</td><td>AI</td><td>Draft</td></tr><tr><td>Retouch</td><td>Human</td><td>Ready</td></tr></tbody></table>')
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(100)

  const afterHtmlPaste = await getPPTTableState(page)

  record('pastes HTML table clipboard data into PPT table element', afterHtmlPaste.tableImportModel === 'canvas-table-import' && afterHtmlPaste.tableImportFormat === 'text-html' && afterHtmlPaste.tableCount === afterPaste.tableCount + 1 && afterHtmlPaste.selectedKind === 'table' && afterHtmlPaste.selectedRows === 3 && afterHtmlPaste.selectedCols === 3 && afterHtmlPaste.tableImportRows === 3 && afterHtmlPaste.tableImportCols === 3 && afterHtmlPaste.cellTexts.includes('Retouch') && afterHtmlPaste.cellTexts.includes('Human'), {
    afterHtmlPaste,
    afterPaste,
  })

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()

    dataTransfer.setData('text/markdown', '| Workstream | Owner | Status |\\n| --- | --- | --- |\\n| Outline | AI | Draft |\\n| Retouch | Human | Ready |')
    dataTransfer.setData('text/plain', '| Workstream | Owner | Status |\\n| --- | --- | --- |\\n| Outline | AI | Draft |\\n| Retouch | Human | Ready |')
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(100)

  const afterMarkdownPaste = await getPPTTableState(page)

  record('pastes Markdown table clipboard data into PPT table element', afterMarkdownPaste.tableImportModel === 'canvas-table-import' && afterMarkdownPaste.tableImportFormat === 'text-markdown' && afterMarkdownPaste.tableCount === afterHtmlPaste.tableCount + 1 && afterMarkdownPaste.importExtensionLastClipboardActions === 'table-source' && afterMarkdownPaste.selectedKind === 'table' && afterMarkdownPaste.selectedRows === 3 && afterMarkdownPaste.selectedCols === 3 && afterMarkdownPaste.tableImportRows === 3 && afterMarkdownPaste.tableImportCols === 3 && afterMarkdownPaste.cellTexts.includes('Workstream') && afterMarkdownPaste.cellTexts.includes('Ready'), {
    afterHtmlPaste,
    afterMarkdownPaste,
  })

  await page.eval(`(() => {
    const files = [
      ${createPPTTestTableFileExpression('pipeline.csv', 'Stage,Owner\nDraft,AI\nRetouch,Human')},
      ${createPPTTestTableFileExpression('scores.tsv', 'Name\tScore\nFit\t92\nTone\t88')},
    ]
    const dataTransfer = {
      files,
      getData: () => '',
      items: files.map((file) => ({
        getAsFile: () => file,
        kind: 'file',
        type: file.type,
      })),
      types: [],
    }
    const event = new Event('paste', {
      bubbles: true,
      cancelable: true,
    })

    Object.defineProperty(event, 'clipboardData', { value: dataTransfer })
    window.dispatchEvent(event)
  })()`)
  await delay(220)

  const afterFilePaste = await getPPTTableState(page)

  record('pastes multiple CSV/TSV files into PPT tables from one clipboard event', afterFilePaste.tableImportModel === 'canvas-table-import' && afterFilePaste.tableImportCount === 2 && afterFilePaste.tableImportFormat === 'text-tsv' && afterFilePaste.tableImportNames.includes('pipeline') && afterFilePaste.tableImportNames.includes('scores') && afterFilePaste.tableCount === afterMarkdownPaste.tableCount + 2 && afterFilePaste.selectedCount === 2 && afterFilePaste.selectedKinds === 'table table' && afterFilePaste.selectedNames.includes('pipeline') && afterFilePaste.selectedNames.includes('scores'), {
    afterFilePaste,
    afterMarkdownPaste,
  })

  await page.eval(`(() => {
    const stage = document.querySelector('.ppt-stage-shell')
    const rect = stage.getBoundingClientRect()
    const dataTransfer = new DataTransfer()

    dataTransfer.items.add(${createPPTTestTableFileExpression('metrics.csv', 'Region,Score\nNA,88\nEU,91')})
    dataTransfer.setData('text/csv', 'Region,Score\\nNA,88\\nEU,91')
    stage.dispatchEvent(new DragEvent('drop', {
      bubbles: true,
      cancelable: true,
      clientX: rect.left + rect.width * 0.28,
      clientY: rect.top + rect.height * 0.64,
      dataTransfer,
    }))
  })()`)
  await delay(180)

  const afterDrop = await getPPTTableState(page)

  record('drops CSV file onto PPT stage as table element', afterDrop.tableImportModel === 'canvas-table-import' && afterDrop.tableImportFormat === 'canvas-csv' && afterDrop.tableCount === afterFilePaste.tableCount + 1 && afterDrop.selectedKind === 'table' && ['metrics', 'Table'].includes(afterDrop.selectedName) && afterDrop.selectedRows === 3 && afterDrop.selectedCols === 2 && afterDrop.cellTexts.includes('Region') && afterDrop.cellTexts.includes('EU') && afterDrop.selectedLeft > 0 && afterDrop.selectedTop >= 0, {
    afterDrop,
    afterFilePaste,
  })

  await page.eval(`(() => {
    const stage = document.querySelector('.ppt-stage-shell')
    if (!stage) return

    const rect = stage.getBoundingClientRect()
    const files = [
      ${createPPTTestTableFileExpression('north.csv', 'Region,Score\nNorth,81\nSouth,79')},
      ${createPPTTestTableFileExpression('quality.tsv', 'Metric\tValue\nPass\t97\nWarn\t3')},
    ]
    const dataTransfer = {
      files,
      getData: () => '',
      items: files.map((file) => ({
        getAsFile: () => file,
        kind: 'file',
        type: file.type,
      })),
      types: [],
    }
    const event = new DragEvent('drop', {
      bubbles: true,
      cancelable: true,
      clientX: rect.left + rect.width * 0.62,
      clientY: rect.top + rect.height * 0.58,
    })

    Object.defineProperty(event, 'dataTransfer', { value: dataTransfer })
    stage.dispatchEvent(event)
  })()`)
  await delay(220)

  const afterMultiDrop = await getPPTTableState(page)

  record('drops multiple CSV/TSV files onto PPT stage as selected tables', afterMultiDrop.tableImportModel === 'canvas-table-import' && afterMultiDrop.tableImportCount === 2 && afterMultiDrop.tableImportFormat === 'text-tsv' && afterMultiDrop.tableImportNames.includes('north') && afterMultiDrop.tableImportNames.includes('quality') && afterMultiDrop.tableCount === afterDrop.tableCount + 2 && afterMultiDrop.selectedCount === 2 && afterMultiDrop.selectedKinds === 'table table' && afterMultiDrop.selectedNames.includes('north') && afterMultiDrop.selectedNames.includes('quality'), {
    afterDrop,
    afterMultiDrop,
  })
}

async function runTextPasteScenario(page) {
  await page.eval(`document.querySelector('.ppt-thumb[aria-label="Open Overview"]')?.click()`)
  await delay(80)

  const before = await getPPTTextPasteState(page)

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()

    dataTransfer.setData('text/html', '<article><p><strong>Bold plan</strong> and <em>italic note</em> with <u>underline</u> and <a href="https://example.com">link</a></p><ul><li>First bullet</li><li><strong>Second bullet</strong></li></ul><ol><li>First step</li><li><strong>Second step</strong></li></ol></article>')
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const afterRichPaste = await getPPTTextPasteState(page)

  record(
    'pastes HTML rich text clipboard data into PPT text body',
      afterRichPaste.textPasteModel === 'canvas-text-paste-import' &&
      afterRichPaste.textPasteImporter === 'ppt-rich-html-text' &&
      afterRichPaste.textPasteFormat === 'text-html-rich' &&
      afterRichPaste.textPasteRichFallback === '' &&
      afterRichPaste.textPasteBoldRuns >= 2 &&
      afterRichPaste.textPasteUnderlineRuns >= 2 &&
      afterRichPaste.textPasteBulletParagraphs === 2 &&
      afterRichPaste.textPasteNumberedParagraphs === 2 &&
      afterRichPaste.textPasteLinkRuns === 1 &&
      afterRichPaste.textBoxCount === before.textBoxCount + 1 &&
      afterRichPaste.selectedKind === 'textBox' &&
      afterRichPaste.selectedName === 'Rich Text' &&
      afterRichPaste.selectedBoldRunCount >= 2 &&
      afterRichPaste.selectedItalicRunCount >= 1 &&
      afterRichPaste.selectedUnderlineRunCount >= 2 &&
      afterRichPaste.selectedBulletParagraphCount === 2 &&
      afterRichPaste.selectedNumberedParagraphCount === 2 &&
      afterRichPaste.selectedText.includes('Bold plan') &&
      afterRichPaste.selectedText.includes('Second bullet') &&
      afterRichPaste.selectedText.includes('Second step'),
    {
      afterRichPaste,
      before,
    },
  )

  await page.eval(`(() => {
    window.__pptTextRichClipboardItemTypes = []
    window.__pptTextRichClipboardWriteCount = 0
    window.__pptTextRichClipboardHTML = ''
    window.__pptTextRichClipboardJSON = ''
    window.__pptTextRichClipboardPlainText = ''
    window.__pptTextRichClipboardSVG = ''

    window.ClipboardItem = class PPTTextRichClipboardItem {
      constructor(items) {
        this.items = items
        window.__pptTextRichClipboardItemTypes.push(Object.keys(items).sort())
      }
    }

    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: {
        write: async (items) => {
          window.__pptTextRichClipboardWriteCount = items.length
          const item = items[0]
          const mimeType = Object.keys(item.items)
            .find((type) => type !== 'text/html' && type !== 'text/plain' && type !== 'image/svg+xml') ?? ''

          window.__pptTextRichClipboardHTML = await item.items['text/html'].text()
          window.__pptTextRichClipboardPlainText = await item.items['text/plain'].text()
          window.__pptTextRichClipboardSVG = item.items['image/svg+xml']
            ? await item.items['image/svg+xml'].text()
            : ''
          window.__pptTextRichClipboardJSON = mimeType
            ? await item.items[mimeType].text()
            : ''
        },
      },
    })
  })()`)

  await pressKey(page, {
    code: 'KeyC',
    key: 'c',
    modifiers: 2,
    windowsVirtualKeyCode: 67,
  })
  await delay(120)

  const afterRichCopy = await getPPTTextPasteState(page)
  const textRichClipboardWrite = await page.eval(`(() => ({
    html: window.__pptTextRichClipboardHTML ?? '',
    itemTypes: window.__pptTextRichClipboardItemTypes?.at(-1) ?? [],
    json: window.__pptTextRichClipboardJSON ?? '',
    plainText: window.__pptTextRichClipboardPlainText ?? '',
    svg: window.__pptTextRichClipboardSVG ?? '',
    writeCount: window.__pptTextRichClipboardWriteCount ?? 0,
  }))()`)

  record(
    'copies PPT rich text selection with semantic HTML clipboard fallback',
    afterRichCopy.richClipboardModel === 'canvas-board-io-ppt-rich-clipboard' &&
      afterRichCopy.richClipboardWriteMode === 'clipboard-item' &&
      afterRichCopy.richClipboardSelection === afterRichPaste.selectedId &&
      afterRichCopy.richClipboardPlainTextLength > 0 &&
      afterRichCopy.richClipboardHTMLLength > afterRichCopy.richClipboardPlainTextLength &&
      textRichClipboardWrite.writeCount === 1 &&
      textRichClipboardWrite.itemTypes.includes(afterRichCopy.richClipboardJsonMimeType) &&
      textRichClipboardWrite.itemTypes.includes('text/html') &&
      textRichClipboardWrite.itemTypes.includes('text/plain') &&
      textRichClipboardWrite.itemTypes.includes('image/svg+xml') &&
      textRichClipboardWrite.html.includes('data-ppt-selection-text-body="true"') &&
      textRichClipboardWrite.html.includes('<strong>Bold plan</strong>') &&
      textRichClipboardWrite.html.includes('<em>italic note</em>') &&
      textRichClipboardWrite.html.includes('<u>underline</u>') &&
      textRichClipboardWrite.html.includes('data-ppt-selection-list="bullet"') &&
      textRichClipboardWrite.html.includes('data-ppt-selection-list="numbered"') &&
      textRichClipboardWrite.html.includes('<ol') &&
      textRichClipboardWrite.html.includes('<li') &&
      textRichClipboardWrite.plainText.includes('Bold plan') &&
      textRichClipboardWrite.plainText.includes('Second bullet') &&
      textRichClipboardWrite.plainText.includes('2. Second step') &&
      textRichClipboardWrite.json.includes('"kind": "interactive-os.ppt.selection"') &&
      textRichClipboardWrite.svg.includes('<svg'),
    {
      afterRichCopy,
      afterRichPaste,
      textRichClipboardWrite,
    },
  )

  await page.eval(`document.querySelector('button[title="Undo"]').click()`)
  await delay(80)

  const afterRichUndo = await getPPTTextPasteState(page)

  record(
    'undoes PPT HTML rich text paste as one history step',
    afterRichUndo.textBoxCount === before.textBoxCount &&
      afterRichUndo.thumbTextCount === before.thumbTextCount &&
      afterRichUndo.redoEnabled,
    {
      afterRichPaste,
      afterRichUndo,
      before,
    },
  )

  const htmlLinkUrl = 'https://example.com/source-brief'

  await page.eval(`((url) => {
    const dataTransfer = new DataTransfer()

    dataTransfer.setData('text/html', '<p><a href="' + url + '">Source brief</a></p>')
    dataTransfer.setData('text/plain', 'Source brief')
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })(${JSON.stringify(htmlLinkUrl)})`)
  await delay(120)

  const afterHTMLLinkPaste = await getPPTTextPasteState(page)

  record(
    'pastes single HTML link as PPT linked text box',
    afterHTMLLinkPaste.textPasteModel === 'canvas-text-paste-import' &&
      afterHTMLLinkPaste.textPasteImporter === 'ppt-rich-html-text' &&
      afterHTMLLinkPaste.textPasteFormat === 'text-html-rich' &&
      afterHTMLLinkPaste.textPasteHyperlinkUrl === htmlLinkUrl &&
      afterHTMLLinkPaste.selectedHyperlink === htmlLinkUrl &&
      afterHTMLLinkPaste.textPasteLinkRuns === 1 &&
      afterHTMLLinkPaste.textBoxCount === before.textBoxCount + 1 &&
      afterHTMLLinkPaste.selectedKind === 'textBox' &&
      afterHTMLLinkPaste.selectedName === 'Rich Text' &&
      afterHTMLLinkPaste.selectedText.includes('Source brief') &&
      afterHTMLLinkPaste.selectedUnderlineRunCount >= 1 &&
      afterHTMLLinkPaste.exportCode.includes(`data-ppt-hyperlink-url="${htmlLinkUrl}"`) &&
      afterHTMLLinkPaste.exportCode.includes(`"url": "${htmlLinkUrl}"`),
    {
      afterHTMLLinkPaste,
      before,
    },
  )

  await page.eval(`document.querySelector('button[title="Undo"]').click()`)
  await delay(80)

  const markdownLinkUrl = 'https://example.com/review-brief'

  await page.eval(`((url) => {
    const dataTransfer = new DataTransfer()
    const markdown = '[Review brief](' + url + ')'

    dataTransfer.setData('text/markdown', markdown)
    dataTransfer.setData('text/plain', markdown)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })(${JSON.stringify(markdownLinkUrl)})`)
  await delay(120)

  const afterMarkdownLinkPaste = await getPPTTextPasteState(page)

  record(
    'pastes single Markdown link as PPT linked text box',
    afterMarkdownLinkPaste.textPasteModel === 'canvas-text-paste-import' &&
      afterMarkdownLinkPaste.importExtensionLastClipboardActions === 'rich-text-source' &&
      afterMarkdownLinkPaste.textPasteImporter === 'ppt-rich-markdown-text' &&
      afterMarkdownLinkPaste.textPasteFormat === 'text-markdown-rich' &&
      afterMarkdownLinkPaste.textPasteHyperlinkUrl === markdownLinkUrl &&
      afterMarkdownLinkPaste.selectedHyperlink === markdownLinkUrl &&
      afterMarkdownLinkPaste.textPasteLinkRuns === 1 &&
      afterMarkdownLinkPaste.textBoxCount === before.textBoxCount + 1 &&
      afterMarkdownLinkPaste.selectedKind === 'textBox' &&
      afterMarkdownLinkPaste.selectedName === 'Markdown Text' &&
      afterMarkdownLinkPaste.selectedText.includes('Review brief') &&
      afterMarkdownLinkPaste.exportCode.includes(`data-ppt-hyperlink-url="${markdownLinkUrl}"`) &&
      afterMarkdownLinkPaste.exportCode.includes(`"url": "${markdownLinkUrl}"`),
    {
      afterMarkdownLinkPaste,
      before,
    },
  )

  await page.eval(`document.querySelector('button[title="Undo"]').click()`)
  await delay(80)

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const markdown = '[Trap](javascript:alert)'

    dataTransfer.setData('text/markdown', markdown)
    dataTransfer.setData('text/plain', markdown)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const afterUnsafeMarkdownLinkPaste = await getPPTTextPasteState(page)

  record(
    'blocks unsafe Markdown link paste from PPT hyperlink metadata',
    afterUnsafeMarkdownLinkPaste.textPasteModel === 'canvas-text-paste-import' &&
      afterUnsafeMarkdownLinkPaste.textPasteImporter === 'ppt-rich-markdown-text' &&
      afterUnsafeMarkdownLinkPaste.textPasteFormat === 'text-markdown-rich' &&
      afterUnsafeMarkdownLinkPaste.textPasteHyperlinkUrl === '' &&
      afterUnsafeMarkdownLinkPaste.selectedHyperlink === '' &&
      afterUnsafeMarkdownLinkPaste.textPasteLinkRuns === 1 &&
      afterUnsafeMarkdownLinkPaste.textBoxCount === before.textBoxCount + 1 &&
      afterUnsafeMarkdownLinkPaste.selectedKind === 'textBox' &&
      afterUnsafeMarkdownLinkPaste.selectedName === 'Markdown Text' &&
      afterUnsafeMarkdownLinkPaste.selectedText.includes('Trap') &&
      !afterUnsafeMarkdownLinkPaste.exportCode.includes('javascript:alert'),
    {
      afterUnsafeMarkdownLinkPaste,
      before,
    },
  )

  await page.eval(`document.querySelector('button[title="Undo"]').click()`)
  await delay(80)

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const markdown = '# Launch plan\\n- **Draft** with _notes_\\n1. [Review](https://example.com/review) handoff'

    dataTransfer.setData('text/markdown', markdown)
    dataTransfer.setData('text/plain', markdown)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const afterMarkdownPaste = await getPPTTextPasteState(page)

  record(
    'pastes Markdown rich text clipboard data into PPT text body',
    afterMarkdownPaste.textPasteModel === 'canvas-text-paste-import' &&
      afterMarkdownPaste.importExtensionLastClipboardActions === 'rich-text-source' &&
      afterMarkdownPaste.textPasteImporter === 'ppt-rich-markdown-text' &&
      afterMarkdownPaste.textPasteFormat === 'text-markdown-rich' &&
      afterMarkdownPaste.textPasteBoldRuns >= 2 &&
      afterMarkdownPaste.textPasteBulletParagraphs === 1 &&
      afterMarkdownPaste.textPasteNumberedParagraphs === 1 &&
      afterMarkdownPaste.textPasteLinkRuns === 1 &&
      afterMarkdownPaste.textPasteUnderlineRuns >= 1 &&
      afterMarkdownPaste.textBoxCount === before.textBoxCount + 1 &&
      afterMarkdownPaste.selectedKind === 'textBox' &&
      afterMarkdownPaste.selectedName === 'Markdown Text' &&
      afterMarkdownPaste.selectedBoldRunCount >= 2 &&
      afterMarkdownPaste.selectedItalicRunCount >= 1 &&
      afterMarkdownPaste.selectedUnderlineRunCount >= 1 &&
      afterMarkdownPaste.selectedBulletParagraphCount === 1 &&
      afterMarkdownPaste.selectedNumberedParagraphCount === 1 &&
      afterMarkdownPaste.selectedText.includes('Launch plan') &&
      afterMarkdownPaste.selectedText.includes('Draft with notes') &&
      afterMarkdownPaste.selectedText.includes('Review handoff'),
    {
      afterMarkdownPaste,
      before,
    },
  )

  await page.eval(`document.querySelector('button[title="Undo"]').click()`)
  await delay(80)

  const afterMarkdownUndo = await getPPTTextPasteState(page)

  record(
    'undoes PPT Markdown rich text paste as one history step',
    afterMarkdownUndo.textBoxCount === before.textBoxCount &&
      afterMarkdownUndo.thumbTextCount === before.thumbTextCount &&
      afterMarkdownUndo.redoEnabled,
    {
      afterMarkdownPaste,
      afterMarkdownUndo,
      before,
    },
  )

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()

    dataTransfer.setData('text/plain', 'Pasted plain text\\nfrom clipboard')
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const afterPaste = await getPPTTextPasteState(page)

  record(
    'pastes plain text clipboard data into PPT text box',
    afterPaste.textPasteModel === 'canvas-text-paste-import' &&
      afterPaste.textPasteImporter === 'ppt-plain-text' &&
      afterPaste.textPasteSelection === afterPaste.selectedId &&
      afterPaste.textBoxCount === before.textBoxCount + 1 &&
      afterPaste.thumbTextCount === before.thumbTextCount + 1 &&
      afterPaste.selectedKind === 'textBox' &&
      afterPaste.selectedName === 'Text' &&
      afterPaste.selectedText.includes('Pasted plain text') &&
      afterPaste.selectedText.includes('from clipboard') &&
      afterPaste.selectedLeft >= 0 &&
      afterPaste.selectedTop >= 0 &&
      afterPaste.selectedWidth > 0 &&
      afterPaste.selectedHeight > 0 &&
      afterPaste.undoEnabled,
    {
      afterPaste,
      before,
    },
  )

  await page.eval(`document.querySelector('button[title="Undo"]').click()`)
  await delay(80)

  const afterUndo = await getPPTTextPasteState(page)

  record(
    'undoes PPT plain text paste as one history step',
    afterUndo.textBoxCount === before.textBoxCount &&
      afterUndo.thumbTextCount === before.thumbTextCount &&
      afterUndo.redoEnabled,
    {
      afterPaste,
      afterUndo,
      before,
    },
  )

  await page.eval(`document.querySelector('button[title="Redo"]').click()`)
  await delay(80)

  const afterRedo = await getPPTTextPasteState(page)

  record(
    'redoes PPT plain text paste with textBody content',
    afterRedo.textBoxCount === before.textBoxCount + 1 &&
      afterRedo.selectedId === afterPaste.selectedId &&
      afterRedo.selectedText.includes('Pasted plain text') &&
      afterRedo.selectedText.includes('from clipboard'),
    {
      afterPaste,
      afterRedo,
    },
  )

  const textPoint = await getElementCenter(page, afterRedo.selectedId)

  await clickMouse(page, textPoint.x, textPoint.y, 2)
  await delay(80)

  const nativeGuard = await page.eval(`((elementId) => {
    const selected = document.querySelector(\`[data-ppt-element="\${elementId}"]\`)
    const editor = selected?.querySelector('.ppt-element-editor')
    const dataTransfer = new DataTransfer()

    dataTransfer.setData('text/plain', 'Native editor paste')

    const beforeCount = document.querySelectorAll('[data-kind="textBox"]').length
    const allowed = editor?.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))

    return {
      afterCount: document.querySelectorAll('[data-kind="textBox"]').length,
      beforeCount,
      editorEditable: editor?.isContentEditable ?? false,
      eventAllowed: allowed ?? false,
      selectedText: selected?.textContent ?? '',
      textPasteImporter: document.querySelector('.ppt-stage-shell')?.getAttribute('data-ppt-text-paste-importer') ?? '',
    }
  })(${JSON.stringify(afterRedo.selectedId)})`)

  record(
    'keeps native PPT text editor paste out of global text import',
    nativeGuard.editorEditable &&
      !nativeGuard.eventAllowed &&
      nativeGuard.afterCount === nativeGuard.beforeCount &&
      nativeGuard.selectedText.includes('Native editor paste') &&
      nativeGuard.selectedText.includes('Pasted plain text') &&
      nativeGuard.textPasteImporter === 'ppt-plain-text',
    nativeGuard,
  )

  await page.eval(`document.activeElement?.blur()`)
  await delay(50)

  const beforeTextBodyJSONPaste = await getPPTTextPasteState(page)

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      textBody: {
        paragraphs: [
          {
            align: 'center',
            runs: [
              { bold: true, color: '#dc2626', text: 'AI revised headline ' },
              { italic: true, size: 28, text: 'needs final polish' },
            ],
            spacingAfter: 8,
            spacingBefore: 4,
          },
          {
            bullet: 'bullet',
            runs: [{ text: 'Retouch copy as PPT text' }],
          },
          {
            bullet: 'numbered',
            runs: [{ text: 'Keep structure editable', underline: true }],
          },
        ],
      },
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const afterTextBodyJSONPaste = await getPPTTextPasteState(page)

  record(
    'pastes JSON textBody into selected PPT text object',
    afterTextBodyJSONPaste.textBodyImportModel === 'ppt-text-body-import' &&
      afterTextBodyJSONPaste.textBodyImportFormat === 'application-json-ppt-text-body' &&
      afterTextBodyJSONPaste.textBodyImportMode === 'text-body' &&
      afterTextBodyJSONPaste.textBodyImportObjects === beforeTextBodyJSONPaste.selectedId &&
      afterTextBodyJSONPaste.textBodyImportTargets === beforeTextBodyJSONPaste.selectedId &&
      afterTextBodyJSONPaste.textBodyImportParagraphs === 3 &&
      afterTextBodyJSONPaste.textBodyImportRuns === 4 &&
      afterTextBodyJSONPaste.textBodyImportTextLength > 70 &&
      afterTextBodyJSONPaste.textBodyImportJsonLength > 180 &&
      afterTextBodyJSONPaste.textBoxCount === beforeTextBodyJSONPaste.textBoxCount &&
      afterTextBodyJSONPaste.selectedId === beforeTextBodyJSONPaste.selectedId &&
      afterTextBodyJSONPaste.selectedText.includes('AI revised headline') &&
      afterTextBodyJSONPaste.selectedText.includes('Retouch copy as PPT text') &&
      afterTextBodyJSONPaste.selectedText.includes('Keep structure editable') &&
      afterTextBodyJSONPaste.selectedBoldRunCount >= 1 &&
      afterTextBodyJSONPaste.selectedItalicRunCount >= 1 &&
      afterTextBodyJSONPaste.selectedUnderlineRunCount >= 1 &&
      afterTextBodyJSONPaste.selectedBulletParagraphCount === 1 &&
      afterTextBodyJSONPaste.selectedNumberedParagraphCount === 1,
    {
      afterTextBodyJSONPaste,
      beforeTextBodyJSONPaste,
    },
  )

  await page.eval(`document.querySelector('button[title="Undo"]').click()`)
  await delay(80)

  const afterTextBodyJSONUndo = await getPPTTextPasteState(page)

  record(
    'undoes PPT textBody JSON paste as one history step',
    afterTextBodyJSONUndo.textBoxCount === beforeTextBodyJSONPaste.textBoxCount &&
      afterTextBodyJSONUndo.selectedId === beforeTextBodyJSONPaste.selectedId &&
      afterTextBodyJSONUndo.selectedText === beforeTextBodyJSONPaste.selectedText &&
      afterTextBodyJSONUndo.redoEnabled,
    {
      afterTextBodyJSONPaste,
      afterTextBodyJSONUndo,
      beforeTextBodyJSONPaste,
    },
  )

  const beforeStyledFallbackPaste = await getPPTTextPasteState(page)

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()

    dataTransfer.setData('text/html', [
      '<section data-ppt-selection-object="styled-html-text" data-ppt-selection-text-body="true" style="width:360px;height:126px;color:#111827;font-family:Inter;font-size:24px;font-weight:400;text-align:center;align-items:center">',
      '<p data-ppt-selection-paragraph="true" style="text-align:center;line-height:1.35;margin-top:7px;margin-bottom:11px">',
      '<span style="color:#dc2626;font-size:34px"><strong>Scaled</strong></span>',
      ' ',
      '<span style="font-size:22px"><em>note</em></span>',
      '</p>',
      '</section>',
    ].join(''))
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const afterStyledFallbackPaste = await getPPTTextPasteState(page)

  record(
    'pastes PPT fallback HTML text styles into editable text model',
    afterStyledFallbackPaste.fallbackHTMLImportModel === 'ppt-fallback-html-import' &&
      afterStyledFallbackPaste.fallbackHTMLImportFormat === 'text-html-ppt-fallback' &&
      afterStyledFallbackPaste.fallbackHTMLImportKind === 'textBox' &&
      afterStyledFallbackPaste.fallbackHTMLImportSourceObject === 'styled-html-text' &&
      afterStyledFallbackPaste.textBoxCount === beforeStyledFallbackPaste.textBoxCount + 1 &&
      afterStyledFallbackPaste.selectedKind === 'textBox' &&
      afterStyledFallbackPaste.selectedName === 'PPT HTML Text' &&
      afterStyledFallbackPaste.selectedText.includes('Scaled note') &&
      afterStyledFallbackPaste.selectedTextAlign === 'center' &&
      afterStyledFallbackPaste.selectedFontSize === '24px' &&
      afterStyledFallbackPaste.selectedRunFontSizes.includes('34px') &&
      afterStyledFallbackPaste.selectedRunFontSizes.includes('22px') &&
      afterStyledFallbackPaste.selectedRunColors.includes('rgb(220, 38, 38)') &&
      afterStyledFallbackPaste.selectedParagraphLineHeights.includes('1.35') &&
      afterStyledFallbackPaste.selectedParagraphSpacingBefore.includes('7') &&
      afterStyledFallbackPaste.selectedParagraphSpacingAfter.includes('11'),
    {
      afterStyledFallbackPaste,
      beforeStyledFallbackPaste,
    },
  )

  await page.eval(`document.activeElement?.blur()`)
  await delay(50)
}

async function runMediaImportScenario(page) {
  await page.eval(`document.querySelector('.ppt-thumb[aria-label="Open Overview"]')?.click()`)
  await delay(80)

  const before = await getPPTMediaImportState(page)
  const pasteUrl = 'https://example.com/ppt-link-card'

  await page.eval(`((url) => {
    const dataTransfer = new DataTransfer()

    dataTransfer.setData('text/uri-list', url)
    dataTransfer.setData('text/plain', url)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })(${JSON.stringify(pasteUrl)})`)
  await delay(120)

  const afterPaste = await getPPTMediaImportState(page)

  record(
    'pastes URL media source into PPT link card via canvas media import',
    afterPaste.mediaImportModel === 'canvas-media-import' &&
      afterPaste.mediaImportImporter === 'ppt-link-card' &&
      afterPaste.mediaImportUrl === pasteUrl &&
      afterPaste.mediaImportSelection === afterPaste.selectedId &&
      afterPaste.shapeCount === before.shapeCount + 1 &&
      afterPaste.textBoxCount === before.textBoxCount &&
      afterPaste.selectedKind === 'shape' &&
      afterPaste.selectedName === 'Link card' &&
      afterPaste.selectedHyperlink === pasteUrl &&
      afterPaste.selectedText.includes(pasteUrl) &&
      afterPaste.thumbHyperlinkCount >= before.thumbHyperlinkCount + 1 &&
      afterPaste.exportCode.includes(`data-ppt-hyperlink-url="${pasteUrl}"`) &&
      afterPaste.exportCode.includes('"name": "Link card"'),
    {
      afterPaste,
      before,
    },
  )

  const jsonUrl = 'https://example.com/json-media-card'
  const jsonTitle = 'Quarterly media brief'

  await page.eval(`((url, title) => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      mediaSource: {
        title,
        url,
      },
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })(${JSON.stringify(jsonUrl)}, ${JSON.stringify(jsonTitle)})`)
  await delay(120)

  const afterJSONPaste = await getPPTMediaImportState(page)

  record(
    'pastes JSON media source into PPT link card via canvas media import',
    afterJSONPaste.mediaImportModel === 'canvas-media-import' &&
      afterJSONPaste.mediaImportImporter === 'ppt-link-card' &&
      afterJSONPaste.mediaImportUrl === jsonUrl &&
      afterJSONPaste.mediaImportSelection === afterJSONPaste.selectedId &&
      afterJSONPaste.mediaJSONImportModel === 'ppt-media-json-import' &&
      afterJSONPaste.mediaJSONImportFormat === 'application-json-ppt-media' &&
      afterJSONPaste.mediaJSONImportFields === 'url title' &&
      afterJSONPaste.mediaJSONImportImporter === 'ppt-link-card' &&
      afterJSONPaste.mediaJSONImportObject === afterJSONPaste.selectedId &&
      afterJSONPaste.mediaJSONImportTitle === jsonTitle &&
      afterJSONPaste.mediaJSONImportUrl === jsonUrl &&
      afterJSONPaste.mediaJSONImportJsonLength > 70 &&
      afterJSONPaste.shapeCount === afterPaste.shapeCount + 1 &&
      afterJSONPaste.selectedKind === 'shape' &&
      afterJSONPaste.selectedName === 'Link card' &&
      afterJSONPaste.selectedHyperlink === jsonUrl &&
      afterJSONPaste.selectedText.includes(jsonTitle) &&
      afterJSONPaste.selectedText.includes(jsonUrl) &&
      afterJSONPaste.exportCode.includes(`data-ppt-hyperlink-url="${jsonUrl}"`),
    {
      afterJSONPaste,
      afterPaste,
    },
  )

  const dropUrl = 'https://example.com/dropped-resource'

  const dropDispatch = await page.eval(`((url) => {
    const stage = document.querySelector('.ppt-stage-shell')
    const rect = stage.getBoundingClientRect()
    const dataTransfer = new DataTransfer()
    const point = {
      x: rect.left + rect.width * 0.72,
      y: rect.top + rect.height * 0.62,
    }
    let captured = null

    stage.addEventListener('drop', (event) => {
      captured = {
        plain: event.dataTransfer?.getData('text/plain') ?? null,
        types: Array.from(event.dataTransfer?.types ?? []),
        uri: event.dataTransfer?.getData('text/uri-list') ?? null,
      }
    }, {
      capture: true,
      once: true,
    })

    dataTransfer.setData('text/uri-list', url)
    dataTransfer.setData('text/plain', url)
    const event = new DragEvent('drop', {
      bubbles: true,
      cancelable: true,
      clientX: point.x,
      clientY: point.y,
      dataTransfer,
    })
    const allowed = stage.dispatchEvent(event)

    return {
      allowed,
      captured,
      defaultPrevented: event.defaultPrevented,
      point,
    }
  })(${JSON.stringify(dropUrl)})`)
  await delay(120)

  const afterDrop = await getPPTMediaImportState(page)

  record(
    'drops URL media source onto PPT stage as link card',
    afterDrop.shapeCount === afterJSONPaste.shapeCount + 1 &&
      afterDrop.selectedKind === 'shape' &&
      afterDrop.selectedName === 'Link card' &&
      afterDrop.selectedHyperlink === dropUrl &&
      afterDrop.mediaImportUrl === dropUrl &&
      dropDispatch.defaultPrevented &&
      dropDispatch.captured?.uri === dropUrl &&
      afterDrop.selectedLeft > afterJSONPaste.selectedLeft &&
      afterDrop.selectedTop > afterJSONPaste.selectedTop,
    {
      afterDrop,
      afterJSONPaste,
      dropDispatch,
    },
  )

  const titlePoint = await getElementCenter(page, 's1-title')
  await clickMouse(page, titlePoint.x, titlePoint.y, 2)
  await delay(80)

  const beforeNativeGuard = await getPPTMediaImportState(page)

  await page.eval(`((url) => {
    const editor = document.querySelector('[data-ppt-element="s1-title"] .ppt-element-editor')
    const dataTransfer = new DataTransfer()

    dataTransfer.setData('text/plain', url)
    editor?.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })(${JSON.stringify('https://example.com/native-editor')})`)
  await delay(80)

  const afterNativeGuard = await getPPTMediaImportState(page)

  record(
    'keeps native PPT text editor URL paste out of global media import',
    beforeNativeGuard.editing &&
      afterNativeGuard.shapeCount === beforeNativeGuard.shapeCount &&
      afterNativeGuard.mediaImportUrl === beforeNativeGuard.mediaImportUrl,
    {
      afterNativeGuard,
      beforeNativeGuard,
    },
  )

  await page.eval(`document.activeElement?.blur()`)
  await delay(50)
}

async function runCommentReviewScenario(page) {
  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)

  const before = await getPPTCommentState(page)

  await page.eval(`document.querySelector('[data-ppt-insert-comment]')?.click()`)
  await delay(40)

  const toolbarPoint = await page.eval(`(() => {
    const slide = document.querySelector('.ppt-slide').getBoundingClientRect()

    return {
      pressed: document.querySelector('[data-ppt-insert-comment]')?.getAttribute('aria-pressed'),
      tool: document.querySelector('.ppt-stage-shell')?.getAttribute('data-creation-tool') ?? '',
      x: slide.left + slide.width * 0.72,
      y: slide.top + slide.height * 0.18,
    }
  })()`)

  await clickMouse(page, toolbarPoint.x, toolbarPoint.y, 1)
  await delay(120)

  const afterToolbarInsert = await getPPTCommentState(page)

  record('creates PPT comment from canvas comment toolbar tool', toolbarPoint.pressed === 'true' && toolbarPoint.tool === 'comment' && afterToolbarInsert.commentCount === before.commentCount + 1 && afterToolbarInsert.selectedKind === 'comment' && afterToolbarInsert.selectedBody === 'Comment' && afterToolbarInsert.thumbCommentCount === before.thumbCommentCount + 1, {
    afterToolbarInsert,
    before,
    toolbarPoint,
  })

  await page.eval(`(() => {
    const textarea = document.querySelector('[data-ppt-style-field="comment-body"]')
    const checkbox = document.querySelector('[data-ppt-style-field="comment-resolved"]')
    const valueSetter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set

    valueSetter.call(textarea, 'Review CTA wording')
    textarea.dispatchEvent(new Event('input', { bubbles: true }))
    textarea.dispatchEvent(new Event('change', { bubbles: true }))

    if (!checkbox.checked) {
      checkbox.click()
    }
  })()`)
  await delay(100)

  const afterInspectorEdit = await getPPTCommentState(page)

  record('retouches PPT comment body and resolved state in inspector', afterInspectorEdit.selectedKind === 'comment' && afterInspectorEdit.selectedBody === 'Review CTA wording' && afterInspectorEdit.inspectorBody === 'Review CTA wording' && afterInspectorEdit.selectedResolved === 'true' && afterInspectorEdit.inspectorResolved, {
    afterInspectorEdit,
    afterToolbarInsert,
  })

  record('shows PPT comment thread in inspector', afterInspectorEdit.commentThreadModel === 'canvas-comment-thread' && afterInspectorEdit.commentThreadCount === 1 && afterInspectorEdit.commentThreadMessageCount === 1 && afterInspectorEdit.commentThreadFirstBody === 'Review CTA wording', {
    afterInspectorEdit,
    afterToolbarInsert,
  })

  await page.eval(`(() => {
    const textarea = document.querySelector('[data-ppt-style-field="comment-reply"]')
    const valueSetter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set

    valueSetter.call(textarea, 'Looks good after headline edit.')
    textarea.dispatchEvent(new Event('input', { bubbles: true }))
    textarea.dispatchEvent(new Event('change', { bubbles: true }))
  })()`)
  await delay(40)
  await page.eval(`document.querySelector('[data-ppt-comment-reply-add]')?.click()`)
  await delay(100)

  const afterThreadReply = await getPPTCommentState(page)

  record('adds PPT comment thread reply from inspector', afterThreadReply.selectedKind === 'comment' && afterThreadReply.commentThreadCount === 2 && afterThreadReply.commentThreadMessageCount === 2 && afterThreadReply.commentThreadBodies.includes('Looks good after headline edit.') && afterThreadReply.commentThreadCommand === 'add-comment-reply' && afterThreadReply.commentThreadCommandBody === 'Looks good after headline edit.' && afterThreadReply.commentThreadCommandCount === 2 && afterThreadReply.commentThreadCommandObject === afterThreadReply.selectedId && afterThreadReply.commentThreadCommandType === 'slide-command-effect' && afterThreadReply.replyInputValue === '', {
    afterInspectorEdit,
    afterThreadReply,
  })

  record('keeps PPT comment body synced with first thread message', afterThreadReply.selectedBody === 'Review CTA wording' && afterThreadReply.inspectorBody === 'Review CTA wording' && afterThreadReply.commentThreadFirstBody === 'Review CTA wording', {
    afterInspectorEdit,
    afterThreadReply,
  })

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      comment: {
        body: 'AI review: tighten final claim.',
        createdAt: 'AI review',
        resolved: false,
        thread: [
          {
            authorName: 'AI',
            body: 'Draft headline overpromises the metric.',
            createdAt: 'AI review',
            id: 'ai-review-1',
          },
          {
            authorName: 'Editor',
            body: 'Use a smaller claim before export.',
            createdAt: 'Human pass',
            id: 'human-review-1',
          },
        ],
      },
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(100)

  const afterCommentJSONPaste = await getPPTCommentState(page)

  record(
    'pastes JSON comment thread into selected PPT comment',
    afterCommentJSONPaste.commentImportModel === 'ppt-comment-import' &&
      afterCommentJSONPaste.commentImportFormat === 'application-json-ppt-comment' &&
      afterCommentJSONPaste.commentImportObjects === afterThreadReply.selectedId &&
      afterCommentJSONPaste.commentImportTargets === afterThreadReply.selectedId &&
      afterCommentJSONPaste.commentImportFields === 'body resolved createdAt thread' &&
      afterCommentJSONPaste.commentImportMessageCount === 2 &&
      afterCommentJSONPaste.commentImportResolved === 'false' &&
      afterCommentJSONPaste.commentImportBodyLength > 20 &&
      afterCommentJSONPaste.commentImportJsonLength > 220 &&
      afterCommentJSONPaste.selectedKind === 'comment' &&
      afterCommentJSONPaste.selectedId === afterThreadReply.selectedId &&
      afterCommentJSONPaste.selectedBody === 'AI review: tighten final claim.' &&
      afterCommentJSONPaste.selectedResolved === '' &&
      afterCommentJSONPaste.inspectorBody === 'AI review: tighten final claim.' &&
      !afterCommentJSONPaste.inspectorResolved &&
      afterCommentJSONPaste.commentThreadCount === 2 &&
      afterCommentJSONPaste.commentThreadMessageCount === 2 &&
      afterCommentJSONPaste.commentThreadFirstBody === 'AI review: tighten final claim.' &&
      afterCommentJSONPaste.commentThreadBodies.includes('Use a smaller claim before export.'),
    {
      afterCommentJSONPaste,
      afterThreadReply,
    },
  )

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(80)

  const afterCommentJSONUndo = await getPPTCommentState(page)

  record(
    'undoes PPT comment JSON paste as one history step',
    afterCommentJSONUndo.selectedId === afterThreadReply.selectedId &&
      afterCommentJSONUndo.selectedBody === 'Review CTA wording' &&
      afterCommentJSONUndo.selectedResolved === 'true' &&
      afterCommentJSONUndo.commentThreadCount === 2 &&
      afterCommentJSONUndo.commentThreadBodies.includes('Looks good after headline edit.'),
    {
      afterCommentJSONPaste,
      afterCommentJSONUndo,
      afterThreadReply,
    },
  )

  const beforeMove = await page.eval(`(() => {
    const selected = document.querySelector('[data-selected="true"]')
    const rect = selected.getBoundingClientRect()

    return {
      left: parseFloat(selected.style.left),
      x: rect.left + rect.width / 2,
      y: rect.top + rect.height / 2,
    }
  })()`)

  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mousePressed',
    x: beforeMove.x,
    y: beforeMove.y,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    type: 'mouseMoved',
    x: beforeMove.x + 42,
    y: beforeMove.y + 18,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mouseReleased',
    x: beforeMove.x + 42,
    y: beforeMove.y + 18,
  })
  await delay(100)

  const afterMove = await getPPTCommentState(page)

  record('moves PPT comment through existing selection transform flow', afterMove.selectedKind === 'comment' && afterMove.selectedLeft !== beforeMove.left, {
    afterMove,
    beforeMove,
  })

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(100)

  const afterUndoMove = await getPPTCommentState(page)

  record('undoes PPT comment move as one history step', afterUndoMove.selectedKind === 'comment' && afterUndoMove.selectedLeft === beforeMove.left && afterUndoMove.selectedBody === 'Review CTA wording', {
    afterUndoMove,
    beforeMove,
  })

  await pressKey(page, {
    code: 'KeyC',
    key: 'c',
    windowsVirtualKeyCode: 67,
  })
  await delay(50)

  const shortcutTool = await getPPTCommentState(page)

  record('activates PPT comment tool from canvas C shortcut', shortcutTool.creationTool === 'comment' && shortcutTool.toolbarPressed === 'true', shortcutTool)

  const shortcutPoint = await page.eval(`(() => {
    const slide = document.querySelector('.ppt-slide').getBoundingClientRect()

    return {
      x: slide.left + slide.width * 0.62,
      y: slide.top + slide.height * 0.68,
    }
  })()`)

  await clickMouse(page, shortcutPoint.x, shortcutPoint.y, 1)
  await delay(100)

  const afterShortcutInsert = await getPPTCommentState(page)

  record('creates PPT comment from keyboard shortcut tool', afterShortcutInsert.commentCount === afterUndoMove.commentCount + 1 && afterShortcutInsert.selectedKind === 'comment', {
    afterShortcutInsert,
    afterUndoMove,
  })

  await pressKey(page, {
    code: 'KeyK',
    key: 'k',
    modifiers: 2,
    windowsVirtualKeyCode: 75,
  })
  await delay(80)
  await page.send('Input.insertText', { text: 'comment' })
  await delay(80)
  await pressKey(page, {
    code: 'Enter',
    key: 'Enter',
    windowsVirtualKeyCode: 13,
  })
  await delay(80)

  const paletteTool = await getPPTCommentState(page)

  record('activates PPT comment tool from command palette', paletteTool.creationTool === 'comment' && !paletteTool.paletteOpen, paletteTool)

  const palettePoint = await page.eval(`(() => {
    const slide = document.querySelector('.ppt-slide').getBoundingClientRect()

    return {
      x: slide.left + slide.width * 0.16,
      y: slide.top + slide.height * 0.72,
    }
  })()`)

  await clickMouse(page, palettePoint.x, palettePoint.y, 1)
  await delay(100)

  const afterPaletteInsert = await getPPTCommentState(page)

  record('creates PPT comment from command palette tool', afterPaletteInsert.commentCount === afterShortcutInsert.commentCount + 1 && afterPaletteInsert.selectedKind === 'comment', {
    afterPaletteInsert,
    afterShortcutInsert,
  })
}

async function runStickySectionScenario(page) {
  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)
  await page.eval(`document.querySelector('.ppt-thumb[aria-label="Open Overview"]')?.click()`)
  await delay(80)

  await pressKey(page, {
    code: 'KeyS',
    key: 's',
    windowsVirtualKeyCode: 83,
  })
  await delay(80)

  const afterStickyShortcut = await getPPTStickySectionState(page)

  record('starts PPT sticky note tool from canvas S shortcut', afterStickyShortcut.creationTool === 'sticky' && afterStickyShortcut.stickyToolbarPressed === 'true' && afterStickyShortcut.stickyToolModel === 'canvas-sticky-note-tool' && afterStickyShortcut.stickyToolShortcut === 'S', afterStickyShortcut)

  const stickyPoint = await page.eval(`(() => {
    const slide = document.querySelector('.ppt-slide').getBoundingClientRect()

    return {
      x: slide.left + slide.width * 0.74,
      y: slide.top + slide.height * 0.28,
    }
  })()`)

  const beforeStickyCreate = await getPPTStickySectionState(page)

  await clickMouse(page, stickyPoint.x, stickyPoint.y, 1)
  await delay(100)

  const afterStickyCreate = await getPPTStickySectionState(page)
  const stickyId = afterStickyCreate.selectedId

  record('creates editable PPT sticky note from canvas sticky tool', afterStickyCreate.elementCount === beforeStickyCreate.elementCount + 1 && afterStickyCreate.selectedKind === 'shape' && afterStickyCreate.selectedShape === 'rect' && afterStickyCreate.selectedName === 'Sticky note' && afterStickyCreate.selectedFillOpacity === '1' && afterStickyCreate.selectedStrokeDash === 'solid' && afterStickyCreate.selectedText.includes('Sticky note') && afterStickyCreate.editingSticky, {
    afterStickyCreate,
    beforeStickyCreate,
  })

  await selectEditableContents(page, stickyId)
  await page.send('Input.insertText', { text: 'Review assumption' })
  await page.eval(`document.querySelector('[data-ppt-element="${stickyId}"] .ppt-element-editor')?.blur()`)
  await delay(100)

  const afterStickyEdit = await getPPTStickySectionState(page)

  record('edits PPT sticky note through existing text editing flow', afterStickyEdit.selectedId === stickyId && afterStickyEdit.selectedText.includes('Review assumption') && afterStickyEdit.exportCode.includes('Review assumption') && afterStickyEdit.undoEnabled, afterStickyEdit)

  await pressKey(page, {
    code: 'KeyS',
    key: 'S',
    modifiers: 8,
    windowsVirtualKeyCode: 83,
  })
  await delay(80)

  const afterSectionShortcut = await getPPTStickySectionState(page)

  record('starts PPT section tool from canvas Shift+S shortcut', afterSectionShortcut.creationTool === 'section' && afterSectionShortcut.sectionToolbarPressed === 'true' && afterSectionShortcut.sectionToolModel === 'canvas-section-tool' && afterSectionShortcut.sectionToolShortcut === 'Shift+S', afterSectionShortcut)

  const sectionDrag = await page.eval(`(() => {
    const slide = document.querySelector('.ppt-slide').getBoundingClientRect()

    return {
      endX: slide.left + slide.width * 0.88,
      endY: slide.top + slide.height * 0.76,
      startX: slide.left + slide.width * 0.52,
      startY: slide.top + slide.height * 0.55,
    }
  })()`)

  const beforeSectionCreate = await getPPTStickySectionState(page)

  await dragMouse(page, [{
    x: sectionDrag.startX,
    y: sectionDrag.startY,
  }, {
    x: sectionDrag.endX,
    y: sectionDrag.endY,
  }])
  await delay(100)

  const afterSectionCreate = await getPPTStickySectionState(page)
  const beforeSectionResizeWidth = afterSectionCreate.selectedWidth

  record('creates PPT section frame from canvas section tool drag', afterSectionCreate.elementCount === beforeSectionCreate.elementCount + 1 && afterSectionCreate.selectedKind === 'shape' && afterSectionCreate.selectedShape === 'rect' && afterSectionCreate.selectedName === 'Section' && afterSectionCreate.selectedStrokeDash === 'dash' && afterSectionCreate.selectedFillOpacity === '0.16' && afterSectionCreate.selectedWidth > 250 && afterSectionCreate.selectedHeight > 100 && afterSectionCreate.exportCode.includes('Section'), {
    afterSectionCreate,
    beforeSectionCreate,
  })

  const resizeHandle = await page.eval(`(() => {
    const rect = document.querySelector('button[aria-label="Resize e"]').getBoundingClientRect()

    return {
      x: rect.left + rect.width / 2,
      y: rect.top + rect.height / 2,
    }
  })()`)

  await dragMouse(page, [{
    x: resizeHandle.x,
    y: resizeHandle.y,
  }, {
    x: resizeHandle.x + 48,
    y: resizeHandle.y,
  }])
  await delay(100)

  const afterSectionResize = await getPPTStickySectionState(page)

  record('resizes PPT section frame through existing selection handles', afterSectionResize.selectedName === 'Section' && afterSectionResize.selectedWidth > beforeSectionResizeWidth, {
    afterSectionResize,
    beforeSectionResizeWidth,
  })

  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(60)
}

async function runLineAffordanceScenario(page) {
  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)

  const before = await getPPTLineState(page)

  await page.eval(`document.querySelector('[data-ppt-insert-line="line"]').click()`)
  await delay(20)

  const drawLine = await page.eval(`(() => {
    const start = document.querySelector('[data-ppt-element="s1-card-1"]').getBoundingClientRect()
    const end = document.querySelector('[data-ppt-element="s1-side-panel"]').getBoundingClientRect()

    return {
      endX: end.left,
      endY: end.top + end.height / 2,
      pressed: document.querySelector('[data-ppt-insert-line="line"]').getAttribute('aria-pressed'),
      startX: start.right,
      startY: start.top + start.height / 2,
    }
  })()`)

  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mousePressed',
    x: drawLine.startX,
    y: drawLine.startY,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    type: 'mouseMoved',
    x: drawLine.endX,
    y: drawLine.endY,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mouseReleased',
    x: drawLine.endX,
    y: drawLine.endY,
  })
  await delay(50)

  const afterLine = await getPPTLineState(page)

  record('draws PPT line element from toolbar creation mode', drawLine.pressed === 'true' && afterLine.lineCount === before.lineCount + 1 && afterLine.selectedKind === 'line' && afterLine.thumbLineCount === before.thumbLineCount + 1, {
    afterLine,
    before,
    drawLine,
  })
  record('attaches drawn PPT line endpoints to shape anchors', afterLine.startConnection !== '' && afterLine.endConnection === 's1-side-panel', {
    afterLine,
  })

  const lineId = afterLine.selectedId
  const beforeConnectionFollow = await getPPTLineState(page, lineId)

  await page.eval(`document.querySelector('[data-ppt-layer-select="s1-side-panel"]')?.click()`)
  await delay(50)
  await pressKey(page, {
    code: 'ArrowRight',
    key: 'ArrowRight',
    windowsVirtualKeyCode: 39,
  })
  await delay(50)

  const afterConnectionFollow = await getPPTLineState(page, lineId)

  record('keeps attached PPT line endpoint connected when target moves', afterConnectionFollow.worldX2 > beforeConnectionFollow.worldX2 && afterConnectionFollow.endConnection === 's1-side-panel', {
    afterConnectionFollow,
    beforeConnectionFollow,
  })

  await page.eval(`document.querySelector(${JSON.stringify(`[data-ppt-layer-select="${lineId}"]`)})?.click()`)
  await delay(50)

  const afterReselectLine = await getPPTLineState(page)

  record('reselects created PPT line from selection pane', afterReselectLine.selectedId === lineId && afterReselectLine.selectedKind === 'line', {
    afterReselectLine,
    lineId,
  })

  const beforeLineColorSwatch = await getPPTColorSwatchState(page, 'line-stroke', lineId)

  await page.eval(`document.querySelector('[data-ppt-color-swatch="line-stroke"][data-ppt-color-token="ppt-color-accent"]')?.click()`)
  await delay(80)

  const afterLineColorSwatch = await getPPTColorSwatchState(page, 'line-stroke', lineId)

  record(
    'applies PPT line stroke theme color swatch',
    beforeLineColorSwatch.themeCount >= 4 &&
      afterLineColorSwatch.model === 'slide-edit-color-swatch-palette' &&
      afterLineColorSwatch.descriptorModel === 'color-swatch-palette' &&
      afterLineColorSwatch.descriptorCommand === 'apply-color-swatch' &&
      afterLineColorSwatch.descriptorControl === 'color-swatch-palette' &&
      afterLineColorSwatch.packageChannel === 'line-stroke' &&
      afterLineColorSwatch.command === 'apply-color-swatch' &&
      afterLineColorSwatch.commandChannel === 'line-stroke' &&
      afterLineColorSwatch.commandObjects.includes(lineId) &&
      afterLineColorSwatch.commandSource === 'theme' &&
      afterLineColorSwatch.commandSwatch === 'theme:ppt-color-accent' &&
      afterLineColorSwatch.commandToken === 'ppt-color-accent' &&
      afterLineColorSwatch.commandType === 'slide-command-effect' &&
      afterLineColorSwatch.commandValue === '#2563eb' &&
      afterLineColorSwatch.stroke === '#2563eb' &&
      afterLineColorSwatch.inputValue === '#2563eb' &&
      afterLineColorSwatch.recentAccentCount === 1 &&
      afterLineColorSwatch.recentUnique &&
      afterLineColorSwatch.exportSlice.includes('"stroke"') &&
      afterLineColorSwatch.exportSlice.includes('"color": "#2563eb"'),
    {
      afterLineColorSwatch,
      beforeLineColorSwatch,
    },
  )

  await page.eval(`(() => {
    const color = document.querySelector('[data-ppt-style-field="line-stroke-color"]')
    const width = document.querySelector('[data-ppt-style-field="line-stroke-width"]')
    const dash = document.querySelector('[data-ppt-style-field="line-stroke-dash"]')
    const marker = document.querySelector('[data-ppt-style-field="line-end-marker"]')
    const colorSetter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set
    const widthSetter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set
    colorSetter.call(color, '#dc2626')
    color.dispatchEvent(new Event('input', { bubbles: true }))
    color.dispatchEvent(new Event('change', { bubbles: true }))
    widthSetter.call(width, '7')
    width.dispatchEvent(new Event('input', { bubbles: true }))
    width.dispatchEvent(new Event('change', { bubbles: true }))
    dash.value = 'dot'
    dash.dispatchEvent(new Event('change', { bubbles: true }))
    marker.value = 'arrow'
    marker.dispatchEvent(new Event('change', { bubbles: true }))
  })()`)
  await delay(50)

  const afterStyle = await getPPTLineState(page)

  record('updates PPT line stroke, dash style, and arrow marker from inspector', afterStyle.stroke === '#dc2626' && afterStyle.strokeWidth === '7' && afterStyle.inspectorDash === 'dot' && afterStyle.descriptorSurface === 'object-stroke-line-style' && afterStyle.descriptorCommand === 'update-object-stroke-line-style' && afterStyle.descriptorControl === 'stroke-line-style-segmented-control' && afterStyle.descriptorAttribute === 'data-slide-object-stroke-line-style' && afterStyle.descriptorAttributeValue === 'dot' && afterStyle.command === 'update-object-stroke-line-style' && afterStyle.commandField === 'strokeLineStyle' && afterStyle.commandObject === afterStyle.selectedId && afterStyle.commandSlide === 'slide-1' && afterStyle.commandType === 'slide-command-effect' && afterStyle.commandValue === 'dot' && afterStyle.selectedDash === 'dot' && afterStyle.strokeDasharray !== '' && afterStyle.thumbDash === 'dot' && afterStyle.markerEnd.includes('url('), afterStyle)

  await page.eval(`document.querySelector('[data-ppt-present-start]')?.click()`)
  await delay(120)

  const preview = await page.eval(`((id) => {
    const overlay = document.querySelector('[data-ppt-presentation]')
    const element = overlay?.querySelector(\`[data-ppt-element="\${id}"]\`)
    const stroke = element?.querySelector('line, [data-ppt-line-path]')

    return {
      dash: element?.getAttribute('data-ppt-stroke-dash') ?? '',
      dasharray: stroke?.getAttribute('stroke-dasharray') ?? '',
      open: !!overlay,
    }
  })(${JSON.stringify(lineId)})`)

  record(
    'keeps PPT line dash style in presentation preview',
    preview.open && preview.dash === 'dot' && preview.dasharray !== '',
    preview,
  )

  await page.eval(`document.querySelector('[data-ppt-present-exit]')?.click()`)
  await delay(80)

  await page.eval(`document.querySelector('[data-ppt-command="copy-formatting"]')?.click()`)
  await delay(80)

  await page.eval(`document.querySelector('[data-ppt-insert-line="line"]').click()`)
  await delay(20)

  const formatTargetLine = await page.eval(`(() => {
    const slide = document.querySelector('.ppt-slide').getBoundingClientRect()

    return {
      endX: slide.left + slide.width * 0.34,
      endY: slide.top + slide.height * 0.82,
      startX: slide.left + slide.width * 0.14,
      startY: slide.top + slide.height * 0.74,
    }
  })()`)

  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mousePressed',
    x: formatTargetLine.startX,
    y: formatTargetLine.startY,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    type: 'mouseMoved',
    x: formatTargetLine.endX,
    y: formatTargetLine.endY,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mouseReleased',
    x: formatTargetLine.endX,
    y: formatTargetLine.endY,
  })
  await delay(80)

  const beforeLineFormatPaste = await getPPTLineState(page)

  await page.eval(`document.querySelector('[data-ppt-command="paste-formatting"]')?.click()`)
  await delay(80)

  const afterLineFormatPaste = await getPPTLineState(page)

  record(
    'pastes PPT line formatting into line target',
    beforeLineFormatPaste.selectedKind === 'line' &&
      beforeLineFormatPaste.strokeWidth !== '7' &&
      beforeLineFormatPaste.selectedDash === 'solid' &&
      afterLineFormatPaste.selectedId === beforeLineFormatPaste.selectedId &&
      afterLineFormatPaste.selectedName === beforeLineFormatPaste.selectedName &&
      afterLineFormatPaste.stroke === '#dc2626' &&
      afterLineFormatPaste.strokeWidth === '7' &&
      afterLineFormatPaste.selectedDash === 'dot' &&
      afterLineFormatPaste.strokeDasharray !== '',
    {
      afterLineFormatPaste,
      beforeLineFormatPaste,
    },
  )

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      lineStyle: {
        stroke: {
          color: '#0f766e',
          dash: 'dash',
          width: 5,
        },
      },
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(100)

  const afterLineStylePaste = await getPPTLineState(page)

  record(
    'pastes JSON line style into selected PPT line',
    afterLineStylePaste.selectedId === beforeLineFormatPaste.selectedId &&
      afterLineStylePaste.selectedName === beforeLineFormatPaste.selectedName &&
      afterLineStylePaste.lineStyleImportModel === 'ppt-line-style-import' &&
      afterLineStylePaste.lineStyleImportFormat === 'application-json-ppt-line-style' &&
      afterLineStylePaste.lineStyleImportCommand === 'paste-object-formatting' &&
      afterLineStylePaste.lineStyleImportCommandTargets === afterLineStylePaste.selectedId &&
      afterLineStylePaste.lineStyleImportCommandType === 'slide-command-effect' &&
      afterLineStylePaste.lineStyleImportObjects === afterLineStylePaste.selectedId &&
      afterLineStylePaste.lineStyleImportCategories.includes('object-effect') &&
      afterLineStylePaste.lineStyleImportCategories.includes('line-style') &&
      afterLineStylePaste.lineStyleImportFields === 'color width dash' &&
      afterLineStylePaste.lineStyleImportStrokeColor === '#0f766e' &&
      afterLineStylePaste.lineStyleImportStrokeDash === 'dash' &&
      afterLineStylePaste.lineStyleImportStrokeWidth === '5' &&
      Number(afterLineStylePaste.lineStyleImportJsonLength) > 60 &&
      afterLineStylePaste.stroke === '#0f766e' &&
      afterLineStylePaste.strokeWidth === '5' &&
      afterLineStylePaste.selectedDash === 'dash' &&
      afterLineStylePaste.strokeDasharray !== '' &&
      afterLineStylePaste.styleClipboardCommand === 'paste-object-formatting' &&
      afterLineStylePaste.styleClipboardCommandTargets === afterLineStylePaste.selectedId &&
      afterLineStylePaste.styleClipboardCommandType === 'slide-command-effect' &&
      afterLineStylePaste.styleClipboardCommandApplications.includes(afterLineStylePaste.selectedId) &&
      afterLineStylePaste.styleClipboardCommandApplications.includes('line-style'),
    {
      afterLineFormatPaste,
      afterLineStylePaste,
      beforeLineFormatPaste,
    },
  )

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      strokeColor: '#7c3aed',
      strokeDash: 'dot',
      strokeWidth: 3,
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(100)

  const afterStandaloneLineStylePaste = await getPPTLineState(page)

  record(
    'pastes standalone PPT line style fields through style clipboard line category',
    afterStandaloneLineStylePaste.selectedId === beforeLineFormatPaste.selectedId &&
      afterStandaloneLineStylePaste.selectedName === beforeLineFormatPaste.selectedName &&
      afterStandaloneLineStylePaste.lineStyleImportModel === 'ppt-line-style-import' &&
      afterStandaloneLineStylePaste.lineStyleImportFormat === 'application-json-ppt-line-style' &&
      afterStandaloneLineStylePaste.lineStyleImportCommand === 'paste-object-formatting' &&
      afterStandaloneLineStylePaste.lineStyleImportCommandTargets === afterStandaloneLineStylePaste.selectedId &&
      afterStandaloneLineStylePaste.lineStyleImportCommandType === 'slide-command-effect' &&
      afterStandaloneLineStylePaste.lineStyleImportObjects === afterStandaloneLineStylePaste.selectedId &&
      afterStandaloneLineStylePaste.lineStyleImportCategories.includes('object-effect') &&
      afterStandaloneLineStylePaste.lineStyleImportCategories.includes('line-style') &&
      afterStandaloneLineStylePaste.lineStyleImportFields === 'color width dash' &&
      afterStandaloneLineStylePaste.lineStyleImportStrokeColor === '#7c3aed' &&
      afterStandaloneLineStylePaste.lineStyleImportStrokeDash === 'dot' &&
      afterStandaloneLineStylePaste.lineStyleImportStrokeWidth === '3' &&
      Number(afterStandaloneLineStylePaste.lineStyleImportJsonLength) > 50 &&
      afterStandaloneLineStylePaste.stroke === '#7c3aed' &&
      afterStandaloneLineStylePaste.strokeWidth === '3' &&
      afterStandaloneLineStylePaste.selectedDash === 'dot' &&
      afterStandaloneLineStylePaste.strokeDasharray !== '' &&
      afterStandaloneLineStylePaste.styleClipboardCommand === 'paste-object-formatting' &&
      afterStandaloneLineStylePaste.styleClipboardCommandTargets === afterStandaloneLineStylePaste.selectedId &&
      afterStandaloneLineStylePaste.styleClipboardCommandType === 'slide-command-effect' &&
      afterStandaloneLineStylePaste.styleClipboardCommandApplications.includes(afterStandaloneLineStylePaste.selectedId) &&
      afterStandaloneLineStylePaste.styleClipboardCommandApplications.includes('line-style'),
    {
      afterLineStylePaste,
      afterStandaloneLineStylePaste,
      beforeLineFormatPaste,
    },
  )

  await pressKey(page, {
    code: 'Delete',
    key: 'Delete',
    windowsVirtualKeyCode: 46,
  })
  await delay(80)

  await page.eval(`document.querySelector(${JSON.stringify(`[data-ppt-layer-select="${lineId}"]`)})?.click()`)
  await delay(50)
  await page.eval(`document.querySelector('[data-ppt-command="copy-formatting"]')?.click()`)
  await delay(80)

  await page.eval(`(() => {
    const route = document.querySelector('[data-ppt-style-field="line-route"]')
    route.value = 'elbow'
    route.dispatchEvent(new Event('change', { bubbles: true }))
  })()`)
  await delay(50)

  const afterRoute = await getPPTLineState(page)

  record('changes PPT connector route to elbow in inspector', afterRoute.route === 'elbow' && afterRoute.hasPath && afterRoute.routeHandleCount === 1, afterRoute)

  const beforeRouteDrag = await page.eval(`(() => {
    const handle = document.querySelector('[data-ppt-line-route-handle]').getBoundingClientRect()
    const path = document.querySelector('[data-selected="true"] [data-ppt-line-path]')

    return {
      d: path.getAttribute('d'),
      handleX: handle.left + handle.width / 2,
      handleY: handle.top + handle.height / 2,
    }
  })()`)

  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mousePressed',
    x: beforeRouteDrag.handleX,
    y: beforeRouteDrag.handleY,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    type: 'mouseMoved',
    x: beforeRouteDrag.handleX - 34,
    y: beforeRouteDrag.handleY,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mouseReleased',
    x: beforeRouteDrag.handleX - 34,
    y: beforeRouteDrag.handleY,
  })
  await delay(50)

  const afterRouteDrag = await getPPTLineState(page)

  record('moves PPT elbow connector route handle', afterRouteDrag.route === 'elbow' && afterRouteDrag.pathD !== beforeRouteDrag.d, {
    afterRouteDrag,
    beforeRouteDrag,
  })

  const beforeEndpoint = await page.eval(`(() => {
    const handle = document.querySelector('[data-ppt-line-endpoint="end"]').getBoundingClientRect()
    const shape = document.querySelector('[data-selected="true"]')

    return {
      handleX: handle.left + handle.width / 2,
      handleY: handle.top + handle.height / 2,
      worldX2: parseFloat(shape.style.left) + Number(shape.getAttribute('data-line-end-x') ?? '0'),
      worldY2: parseFloat(shape.style.top) + Number(shape.getAttribute('data-line-end-y') ?? '0'),
    }
  })()`)

  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mousePressed',
    x: beforeEndpoint.handleX,
    y: beforeEndpoint.handleY,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    type: 'mouseMoved',
    x: beforeEndpoint.handleX + 40,
    y: beforeEndpoint.handleY + 32,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mouseReleased',
    x: beforeEndpoint.handleX + 40,
    y: beforeEndpoint.handleY + 32,
  })
  await delay(50)

  const afterEndpoint = await getPPTLineState(page)

  record('moves PPT line endpoint with endpoint handle', afterEndpoint.endpointHandleCount === 2 && afterEndpoint.worldX2 !== beforeEndpoint.worldX2 && afterEndpoint.worldY2 !== beforeEndpoint.worldY2 && afterEndpoint.endConnection === '', {
    afterEndpoint,
    beforeEndpoint,
  })

  await pressKey(page, {
    code: 'KeyL',
    key: 'l',
    windowsVirtualKeyCode: 76,
  })
  await delay(50)

  const afterArrowShortcut = await getPPTLineState(page)

  record('starts PPT arrow tool from canvas L shortcut', afterArrowShortcut.lineTool === 'arrow' && afterArrowShortcut.arrowToolPressed === 'true' && afterArrowShortcut.arrowToolModel === 'canvas-arrow-tool' && afterArrowShortcut.arrowToolShortcut === 'L', afterArrowShortcut)

  const drawArrow = await page.eval(`(() => {
    const stage = document.querySelector('.ppt-stage-shell').getBoundingClientRect()

    return {
      endX: stage.left + stage.width * 0.66,
      endY: stage.top + stage.height * 0.34,
      pressed: document.querySelector('[data-ppt-insert-line="arrow"]').getAttribute('aria-pressed'),
      startX: stage.left + stage.width * 0.46,
      startY: stage.top + stage.height * 0.34,
    }
  })()`)

  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mousePressed',
    x: drawArrow.startX,
    y: drawArrow.startY,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    type: 'mouseMoved',
    x: drawArrow.endX,
    y: drawArrow.endY,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mouseReleased',
    x: drawArrow.endX,
    y: drawArrow.endY,
  })
  await delay(50)

  const afterArrow = await getPPTLineState(page)

  record('draws PPT arrow element from toolbar creation mode', drawArrow.pressed === 'true' && afterArrow.lineCount === afterLine.lineCount + 1 && afterArrow.selectedKind === 'line' && afterArrow.markerEnd.includes('url(') && afterArrow.selectedName === 'Arrow', {
    afterArrow,
    afterLine,
    drawArrow,
  })

  const beforeResize = await page.eval(`(() => {
    const selected = document.querySelector('[data-selected="true"]')
    const handle = document.querySelector('button[aria-label="Resize e"]').getBoundingClientRect()

    return {
      handleX: handle.left + handle.width / 2,
      handleY: handle.top + handle.height / 2,
      width: parseFloat(selected.style.width),
    }
  })()`)

  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mousePressed',
    x: beforeResize.handleX,
    y: beforeResize.handleY,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    type: 'mouseMoved',
    x: beforeResize.handleX + 48,
    y: beforeResize.handleY,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mouseReleased',
    x: beforeResize.handleX + 48,
    y: beforeResize.handleY,
  })
  await delay(50)

  const afterResize = await getPPTLineState(page)

  record('resizes PPT line with existing selection handles', afterResize.selectedWidth > beforeResize.width && afterResize.selectedKind === 'line', {
    afterResize,
    beforeResize,
  })
}

async function runFreeformScenario(page) {
  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)
  await page.eval(`document.querySelector('.ppt-thumb[aria-label="Open Overview"]')?.click()`)
  await delay(80)

  const before = await getPPTFreeformState(page)

  await page.eval(`document.querySelector('[data-ppt-insert-tool="pen"]')?.click()`)
  await delay(40)

  const draw = await page.eval(`(() => {
    const slide = document.querySelector('.ppt-slide').getBoundingClientRect()

    return {
      midX: slide.left + slide.width * 0.47,
      midY: slide.top + slide.height * 0.21,
      pressed: document.querySelector('[data-ppt-insert-tool="pen"]')?.getAttribute('aria-pressed') ?? '',
      startX: slide.left + slide.width * 0.33,
      startY: slide.top + slide.height * 0.28,
      endX: slide.left + slide.width * 0.58,
      endY: slide.top + slide.height * 0.31,
    }
  })()`)

  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mousePressed',
    x: draw.startX,
    y: draw.startY,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    type: 'mouseMoved',
    x: draw.midX,
    y: draw.midY,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    type: 'mouseMoved',
    x: draw.endX,
    y: draw.endY,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mouseReleased',
    x: draw.endX,
    y: draw.endY,
  })
  await delay(100)

  const afterCreate = await getPPTFreeformState(page)

  record('draws PPT freeform path from toolbar pen tool', draw.pressed === 'true' && afterCreate.freeformCount === before.freeformCount + 1 && afterCreate.selectedKind === 'freeform' && afterCreate.pathD.includes('M ') && afterCreate.pointCount >= 2 && afterCreate.thumbFreeformCount === before.thumbFreeformCount + 1, {
    afterCreate,
    before,
    draw,
  })

  const freeformId = afterCreate.selectedId
  const beforeFreeformFormatPaste = await getPPTFreeformState(page)

  await page.eval(`document.querySelector('[data-ppt-command="paste-formatting"]')?.click()`)
  await delay(80)

  const afterFreeformFormatPaste = await getPPTFreeformState(page)

  record(
    'pastes PPT line formatting into freeform target',
    beforeFreeformFormatPaste.selectedKind === 'freeform' &&
      beforeFreeformFormatPaste.strokeWidth !== '7' &&
      beforeFreeformFormatPaste.selectedDash === 'solid' &&
      !beforeFreeformFormatPaste.pasteDisabled &&
      afterFreeformFormatPaste.selectedId === beforeFreeformFormatPaste.selectedId &&
      afterFreeformFormatPaste.selectedName === beforeFreeformFormatPaste.selectedName &&
      afterFreeformFormatPaste.stroke === '#dc2626' &&
      afterFreeformFormatPaste.strokeWidth === '7' &&
      afterFreeformFormatPaste.selectedDash === 'dot' &&
      afterFreeformFormatPaste.strokeDasharray !== '',
    {
      afterFreeformFormatPaste,
      beforeFreeformFormatPaste,
    },
  )

  await page.eval(`document.querySelector('[data-ppt-color-swatch="line-stroke"][data-ppt-color-token="ppt-color-accent"]')?.click()`)
  await delay(80)

  const afterFreeformColorSwatch = await getPPTColorSwatchState(page, 'line-stroke', freeformId)

  record(
    'applies PPT freeform stroke theme color swatch',
    afterFreeformColorSwatch.selectedKind === 'freeform' &&
      afterFreeformColorSwatch.model === 'slide-edit-color-swatch-palette' &&
      afterFreeformColorSwatch.descriptorModel === 'color-swatch-palette' &&
      afterFreeformColorSwatch.descriptorCommand === 'apply-color-swatch' &&
      afterFreeformColorSwatch.descriptorControl === 'color-swatch-palette' &&
      afterFreeformColorSwatch.packageChannel === 'line-stroke' &&
      afterFreeformColorSwatch.command === 'apply-color-swatch' &&
      afterFreeformColorSwatch.commandChannel === 'line-stroke' &&
      afterFreeformColorSwatch.commandObjects.includes(freeformId) &&
      afterFreeformColorSwatch.commandSource === 'theme' &&
      afterFreeformColorSwatch.commandSwatch === 'theme:ppt-color-accent' &&
      afterFreeformColorSwatch.commandToken === 'ppt-color-accent' &&
      afterFreeformColorSwatch.commandType === 'slide-command-effect' &&
      afterFreeformColorSwatch.commandValue === '#2563eb' &&
      afterFreeformColorSwatch.stroke === '#2563eb' &&
      afterFreeformColorSwatch.inputValue === '#2563eb' &&
      afterFreeformColorSwatch.recentAccentCount === 1 &&
      afterFreeformColorSwatch.recentUnique &&
      afterFreeformColorSwatch.exportSlice.includes('"kind": "freeform"') &&
      afterFreeformColorSwatch.exportSlice.includes('"stroke"') &&
      afterFreeformColorSwatch.exportSlice.includes('"color": "#2563eb"'),
    afterFreeformColorSwatch,
  )

  const beforeMovePoint = await getElementCenter(page, freeformId)

  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mousePressed',
    x: beforeMovePoint.x,
    y: beforeMovePoint.y,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    type: 'mouseMoved',
    x: beforeMovePoint.x + 46,
    y: beforeMovePoint.y + 28,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mouseReleased',
    x: beforeMovePoint.x + 46,
    y: beforeMovePoint.y + 28,
  })
  await delay(100)

  const afterMove = await getPPTFreeformState(page)

  record('moves PPT freeform path with existing selection transform flow', afterMove.selectedId === freeformId && afterMove.selectedKind === 'freeform' && afterMove.selectedLeft !== afterCreate.selectedLeft && afterMove.selectedTop !== afterCreate.selectedTop, {
    afterCreate,
    afterMove,
  })

  const beforeResize = await page.eval(`(() => {
    const selected = document.querySelector('[data-selected="true"]')
    const handle = document.querySelector('button[aria-label="Resize e"]').getBoundingClientRect()

    return {
      handleX: handle.left + handle.width / 2,
      handleY: handle.top + handle.height / 2,
      pathD: selected?.querySelector('[data-ppt-freeform-path]')?.getAttribute('d') ?? '',
      width: parseFloat(selected?.style.width ?? '0'),
    }
  })()`)

  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mousePressed',
    x: beforeResize.handleX,
    y: beforeResize.handleY,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    type: 'mouseMoved',
    x: beforeResize.handleX + 52,
    y: beforeResize.handleY,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mouseReleased',
    x: beforeResize.handleX + 52,
    y: beforeResize.handleY,
  })
  await delay(100)

  const afterResize = await getPPTFreeformState(page)

  record('resizes PPT freeform path with existing handles', afterResize.selectedKind === 'freeform' && afterResize.selectedWidth > beforeResize.width && afterResize.pathD !== beforeResize.pathD, {
    afterResize,
    beforeResize,
  })

  const beforeRotate = await page.eval(`(() => {
    const handle = document.querySelector('.ppt-rotate-handle').getBoundingClientRect()
    const selected = document.querySelector('[data-selected="true"]')

    return {
      handleX: handle.left + handle.width / 2,
      handleY: handle.top + handle.height / 2,
      rotation: selected?.getAttribute('data-rotation') ?? '',
    }
  })()`)

  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mousePressed',
    x: beforeRotate.handleX,
    y: beforeRotate.handleY,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    type: 'mouseMoved',
    x: beforeRotate.handleX + 42,
    y: beforeRotate.handleY - 26,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mouseReleased',
    x: beforeRotate.handleX + 42,
    y: beforeRotate.handleY - 26,
  })
  await delay(100)

  const afterRotate = await getPPTFreeformState(page)

  record('rotates PPT freeform path with existing rotation handle', afterRotate.selectedKind === 'freeform' && afterRotate.selectedRotation !== beforeRotate.rotation && afterRotate.selectedTransform.includes('rotate'), {
    afterRotate,
    beforeRotate,
  })

  await pressKey(page, {
    code: 'KeyD',
    key: 'd',
    modifiers: 2,
    windowsVirtualKeyCode: 68,
  })
  await delay(120)

  const afterDuplicate = await getPPTFreeformState(page)

  record('duplicates PPT freeform path through existing command flow', afterDuplicate.freeformCount === afterResize.freeformCount + 1 && afterDuplicate.selectedKind === 'freeform' && afterDuplicate.selectedId !== freeformId, {
    afterDuplicate,
    afterResize,
  })

  await pressKey(page, {
    code: 'Delete',
    key: 'Delete',
    windowsVirtualKeyCode: 46,
  })
  await delay(120)

  const afterDelete = await getPPTFreeformState(page)

  record('deletes PPT freeform path through existing command flow', afterDelete.freeformCount === afterResize.freeformCount, {
    afterDelete,
    afterDuplicate,
    afterResize,
  })

  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(60)
  await pressKey(page, {
    code: 'KeyP',
    key: 'P',
    modifiers: 8,
    windowsVirtualKeyCode: 80,
  })
  await delay(80)

  const afterShortcut = await getPPTFreeformState(page)

  record('starts PPT pen tool from canvas Shift+P shortcut', afterShortcut.creationTool === 'freeform' && afterShortcut.toolbarPressed === 'true', afterShortcut)

  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(60)
  await pressKey(page, {
    code: 'KeyK',
    key: 'k',
    modifiers: 2,
    windowsVirtualKeyCode: 75,
  })
  await delay(80)
  await page.send('Input.insertText', { text: 'pen' })
  await delay(80)

  const beforePalette = await getPPTFreeformState(page)

  await pressKey(page, {
    code: 'Enter',
    key: 'Enter',
    windowsVirtualKeyCode: 13,
  })
  await delay(100)

  const afterPalette = await getPPTFreeformState(page)

  record('starts PPT pen tool from command palette', beforePalette.paletteOpen && beforePalette.palettePen && afterPalette.creationTool === 'freeform' && !afterPalette.paletteOpen, {
    afterPalette,
    beforePalette,
  })

  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(60)
  await pressKey(page, {
    code: 'KeyM',
    key: 'm',
    windowsVirtualKeyCode: 77,
  })
  await delay(80)

  const afterMarkerShortcut = await getPPTFreeformState(page)

  record('starts PPT marker tool from canvas M shortcut', afterMarkerShortcut.creationTool === 'marker' && afterMarkerShortcut.drawingTool === 'marker' && afterMarkerShortcut.markerToolbarPressed === 'true', afterMarkerShortcut)

  const markerDraw = await page.eval(`(() => {
    const slide = document.querySelector('.ppt-slide').getBoundingClientRect()

    return {
      endX: slide.left + slide.width * 0.42,
      endY: slide.top + slide.height * 0.52,
      midX: slide.left + slide.width * 0.34,
      midY: slide.top + slide.height * 0.47,
      startX: slide.left + slide.width * 0.25,
      startY: slide.top + slide.height * 0.50,
    }
  })()`)

  await dragMouse(page, [{
    x: markerDraw.startX,
    y: markerDraw.startY,
  }, {
    x: markerDraw.midX,
    y: markerDraw.midY,
  }, {
    x: markerDraw.endX,
    y: markerDraw.endY,
  }])
  await delay(100)

  const afterMarkerCreate = await getPPTFreeformState(page)

  record('draws PPT marker path with canvas marker stroke style', afterMarkerCreate.freeformCount === afterPalette.freeformCount + 1 && afterMarkerCreate.selectedKind === 'freeform' && afterMarkerCreate.selectedName === 'Marker' && afterMarkerCreate.stroke === '#475569' && afterMarkerCreate.strokeWidth === '4' && afterMarkerCreate.selectedOpacity === '1', {
    afterMarkerCreate,
    afterPalette,
  })

  await pressKey(page, {
    code: 'KeyM',
    key: 'M',
    modifiers: 8,
    windowsVirtualKeyCode: 77,
  })
  await delay(80)

  const afterHighlighterShortcut = await getPPTFreeformState(page)

  record('starts PPT highlighter tool from canvas Shift+M shortcut', afterHighlighterShortcut.creationTool === 'highlight' && afterHighlighterShortcut.drawingTool === 'highlight' && afterHighlighterShortcut.highlighterToolbarPressed === 'true', afterHighlighterShortcut)

  const highlighterDraw = await page.eval(`(() => {
    const slide = document.querySelector('.ppt-slide').getBoundingClientRect()

    return {
      endX: slide.left + slide.width * 0.73,
      endY: slide.top + slide.height * 0.49,
      midX: slide.left + slide.width * 0.63,
      midY: slide.top + slide.height * 0.45,
      startX: slide.left + slide.width * 0.53,
      startY: slide.top + slide.height * 0.48,
    }
  })()`)

  await dragMouse(page, [{
    x: highlighterDraw.startX,
    y: highlighterDraw.startY,
  }, {
    x: highlighterDraw.midX,
    y: highlighterDraw.midY,
  }, {
    x: highlighterDraw.endX,
    y: highlighterDraw.endY,
  }])
  await delay(100)

  const afterHighlighterCreate = await getPPTFreeformState(page)
  const highlighterId = afterHighlighterCreate.selectedId

  record('draws PPT highlighter path with canvas highlighter stroke style', afterHighlighterCreate.freeformCount === afterMarkerCreate.freeformCount + 1 && afterHighlighterCreate.selectedKind === 'freeform' && afterHighlighterCreate.selectedName === 'Highlighter' && afterHighlighterCreate.stroke === '#fde047' && afterHighlighterCreate.strokeWidth === '18' && afterHighlighterCreate.selectedOpacity === '0.42' && afterHighlighterCreate.selectedStyleOpacity === '0.42', {
    afterHighlighterCreate,
    afterMarkerCreate,
  })

  await pressKey(page, {
    code: 'KeyE',
    key: 'e',
    windowsVirtualKeyCode: 69,
  })
  await delay(80)

  const afterEraserShortcut = await getPPTFreeformState(page)

  record('starts PPT eraser tool from canvas E shortcut', afterEraserShortcut.eraserActive === 'true' && afterEraserShortcut.eraserToolbarPressed === 'true' && afterEraserShortcut.creationTool === '', afterEraserShortcut)

  const nonDrawingEraseTrace = await page.eval(`(() => {
    const shapes = [...document.querySelectorAll('[data-kind="shape"]')]
      .map((element) => {
        const rect = element.getBoundingClientRect()

        return {
          h: rect.height,
          id: element.getAttribute('data-ppt-element') ?? '',
          w: rect.width,
          x: rect.left,
          y: rect.top,
        }
      })
      .filter((rect) => rect.w > 24 && rect.h > 24)
    const freeforms = [...document.querySelectorAll('[data-kind="freeform"]')]
      .map((element) => {
        const rect = element.getBoundingClientRect()

        return {
          bottom: rect.bottom,
          left: rect.left,
          right: rect.right,
          top: rect.top,
        }
      })
    const candidates = [
      [0.16, 0.18],
      [0.84, 0.18],
      [0.16, 0.82],
      [0.84, 0.82],
      [0.5, 0.16],
      [0.5, 0.84],
      [0.5, 0.5],
    ]
    const intersectsFreeform = (x1, y1, x2, y2) => {
      const pad = 24
      const left = Math.min(x1, x2) - pad
      const right = Math.max(x1, x2) + pad
      const top = Math.min(y1, y2) - pad
      const bottom = Math.max(y1, y2) + pad

      return freeforms.some((rect) =>
        right >= rect.left &&
          left <= rect.right &&
          bottom >= rect.top &&
          top <= rect.bottom)
    }

    for (const shape of shapes) {
      const delta = Math.min(12, shape.w * 0.08, shape.h * 0.08)

      for (const [fx, fy] of candidates) {
        const centerX = shape.x + shape.w * fx
        const centerY = shape.y + shape.h * fy
        const startX = centerX - delta
        const startY = centerY - delta
        const endX = centerX + delta
        const endY = centerY + delta

        if (!intersectsFreeform(startX, startY, endX, endY)) {
          return {
            endX,
            endY,
            shapeId: shape.id,
            startX,
            startY,
          }
        }
      }
    }

    const fallback = shapes[0]
    const centerX = fallback.x + fallback.w / 2
    const centerY = fallback.y + fallback.h / 2

    return {
      endX: centerX + 8,
      endY: centerY + 8,
      shapeId: fallback.id,
      startX: centerX - 8,
      startY: centerY - 8,
    }
  })()`)
  const beforeNonDrawingErase = await getPPTFreeformState(page)

  await dragMouse(page, [{
    x: nonDrawingEraseTrace.startX,
    y: nonDrawingEraseTrace.startY,
  }, {
    x: nonDrawingEraseTrace.endX,
    y: nonDrawingEraseTrace.endY,
  }])
  await delay(100)

  const afterNonDrawingErase = await getPPTFreeformState(page)

  record('keeps non-drawing PPT objects when eraser crosses shapes', afterNonDrawingErase.eraserActive === 'true' && afterNonDrawingErase.shapeCount === beforeNonDrawingErase.shapeCount && afterNonDrawingErase.freeformCount === beforeNonDrawingErase.freeformCount && afterNonDrawingErase.elementCount === beforeNonDrawingErase.elementCount, {
    afterNonDrawingErase,
    beforeNonDrawingErase,
    nonDrawingEraseTrace,
  })

  const highlighterTrace = await page.eval(`(() => {
    const id = ${JSON.stringify(highlighterId)}
    const rect = document.querySelector('[data-ppt-element="' + id + '"]').getBoundingClientRect()

    return {
      endX: rect.left + rect.width * 0.75,
      endY: rect.top + rect.height * 0.52,
      midX: rect.left + rect.width / 2,
      midY: rect.top + rect.height / 2,
      startX: rect.left + rect.width * 0.25,
      startY: rect.top + rect.height * 0.48,
    }
  })()`)

  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mousePressed',
    x: highlighterTrace.startX,
    y: highlighterTrace.startY,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    type: 'mouseMoved',
    x: highlighterTrace.midX,
    y: highlighterTrace.midY,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    type: 'mouseMoved',
    x: highlighterTrace.endX,
    y: highlighterTrace.endY,
  })
  await delay(50)

  const duringEraserDrag = await getPPTFreeformState(page)

  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    type: 'mouseReleased',
    x: highlighterTrace.endX,
    y: highlighterTrace.endY,
  })
  await delay(120)

  const afterEraserDelete = await getPPTFreeformState(page)

  record('previews and erases PPT freeform stroke with canvas eraser semantics', duringEraserDrag.eraserHitCount > 0 && duringEraserDrag.eraserHitIds.includes(highlighterId) && afterEraserDelete.freeformCount === afterNonDrawingErase.freeformCount - 1 && !afterEraserDelete.freeformIds.includes(highlighterId) && afterEraserDelete.shapeCount === afterNonDrawingErase.shapeCount, {
    afterEraserDelete,
    afterNonDrawingErase,
    duringEraserDrag,
  })

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(120)

  const afterEraserUndo = await getPPTFreeformState(page)

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(120)

  const afterEraserRedo = await getPPTFreeformState(page)

  record('records PPT eraser stroke deletion as undoable history step', afterEraserUndo.freeformCount === afterNonDrawingErase.freeformCount && afterEraserUndo.freeformIds.includes(highlighterId) && afterEraserRedo.freeformCount === afterEraserDelete.freeformCount && !afterEraserRedo.freeformIds.includes(highlighterId), {
    afterEraserDelete,
    afterEraserRedo,
    afterEraserUndo,
    afterNonDrawingErase,
  })

  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(60)
}

async function runFlipSelectionScenario(page) {
  const shapeIds = ['s1-card-1', 's1-card-2']

  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)
  await page.eval(`document.querySelector('[data-ppt-view-fit-slide]')?.click()`)
  await delay(80)

  await selectPPTLayerRows(page, shapeIds)
  await delay(80)

  const beforeShapeFlip = await readPPTFlipState(page, shapeIds)

  record('enables PPT flip commands for visible unlocked selection', beforeShapeFlip.selectedCount === 2 && !beforeShapeFlip.flipHorizontalDisabled && !beforeShapeFlip.flipVerticalDisabled, beforeShapeFlip)

  await page.eval(`document.querySelector('[data-ppt-command="flip-horizontal"]')?.click()`)
  await delay(100)

  const afterShapeFlip = await readPPTFlipState(page, shapeIds)

  record('flips selected PPT shapes horizontally around selection bounds', afterShapeFlip.selectedCount === 2 && positionsChanged(beforeShapeFlip.positions, afterShapeFlip.positions, shapeIds) && positionsMatch(afterShapeFlip.positions, beforeShapeFlip.expectedHorizontal, shapeIds) && shapeIds.every((id) => afterShapeFlip.positions[id]?.flipH === 'true'), {
    afterShapeFlip,
    beforeShapeFlip,
  })

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(100)

  const afterShapeUndo = await readPPTFlipState(page, shapeIds)

  record('undoes PPT shape flip in one step', positionsMatch(afterShapeUndo.positions, beforeShapeFlip.positions, shapeIds) && shapeIds.every((id) => afterShapeUndo.positions[id]?.flipH !== 'true'), {
    afterShapeUndo,
    beforeShapeFlip,
  })

  const textIds = ['s1-title', 's1-summary']

  await selectPPTLayerRows(page, textIds)
  await delay(80)

  const beforeTextFlip = await readPPTFlipState(page, textIds)

  await pressKey(page, {
    code: 'KeyK',
    key: 'k',
    modifiers: 2,
    windowsVirtualKeyCode: 75,
  })
  await delay(80)
  await readCommandPaletteIds(page, 'flip vertical')
  await pressKey(page, {
    code: 'Enter',
    key: 'Enter',
    windowsVirtualKeyCode: 13,
  })
  await delay(120)

  const afterTextFlip = await readPPTFlipState(page, textIds)

  record('flips selected PPT text boxes from command palette', afterTextFlip.selectedCount === 2 && positionsMatch(afterTextFlip.positions, beforeTextFlip.expectedVertical, textIds) && textIds.every((id) => afterTextFlip.positions[id]?.flipV === 'true'), {
    afterTextFlip,
    beforeTextFlip,
  })

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(100)

  const lineId = await page.eval(`(() => document.querySelector('[data-kind="line"]')?.getAttribute('data-ppt-element') ?? '')()`)

  await selectPPTLayerRows(page, [lineId])
  await delay(80)

  const beforeLineFlip = await getPPTLineState(page, lineId)
  const linePivotX = (Math.min(beforeLineFlip.worldX1, beforeLineFlip.worldX2) +
    Math.max(beforeLineFlip.worldX1, beforeLineFlip.worldX2)) / 2

  await page.eval(`document.querySelector('[data-ppt-command="flip-horizontal"]')?.click()`)
  await delay(100)

  const afterLineFlip = await getPPTLineState(page, lineId)

  record('flips PPT line endpoints horizontally', afterLineFlip.selectedKind === 'line' && nearlyEqual(afterLineFlip.worldX1, 2 * linePivotX - beforeLineFlip.worldX1) && nearlyEqual(afterLineFlip.worldX2, 2 * linePivotX - beforeLineFlip.worldX2) && afterLineFlip.startConnection === '' && afterLineFlip.endConnection === '', {
    afterLineFlip,
    beforeLineFlip,
  })

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(100)

  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)
  await pressKey(page, {
    code: 'KeyK',
    key: 'k',
    modifiers: 2,
    windowsVirtualKeyCode: 75,
  })
  await delay(80)
  await readCommandPaletteIds(page, 'flip')

  const paletteDisabled = await page.eval(`(() => ({
    horizontalDisabled: document.querySelector('[data-ppt-command-palette-item="command:flip-horizontal"]')?.disabled ?? false,
    itemPresent: !!document.querySelector('[data-ppt-command-palette-item="command:flip-horizontal"]'),
    open: !!document.querySelector('[data-ppt-command-palette]'),
    selectedCount: document.querySelectorAll('[data-selected="true"]').length,
    verticalDisabled: document.querySelector('[data-ppt-command-palette-item="command:flip-vertical"]')?.disabled ?? false,
  }))()`)

  record('disables PPT flip palette items without selection', paletteDisabled.itemPresent && paletteDisabled.open && paletteDisabled.selectedCount === 0 && paletteDisabled.horizontalDisabled && paletteDisabled.verticalDisabled, paletteDisabled)

  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)

  const imageId = await page.eval(`(() => [...document.querySelectorAll('[data-kind="image"]')].at(-1)?.getAttribute('data-ppt-element') ?? '')()`)

  await selectPPTLayerRows(page, [imageId])
  await delay(80)
  await page.eval(`document.querySelector('[data-ppt-command="flip-horizontal"]')?.click()`)
  await delay(100)

  const afterImageFlip = await getPPTImageImportState(page)

  record('persists PPT image horizontal flip for export', afterImageFlip.selectedKind === 'image' && afterImageFlip.selectedFlipH === 'true' && afterImageFlip.selectedTransform.includes('scaleX(-1)') && afterImageFlip.thumbFlipHCount > 0, afterImageFlip)
}

async function runSelectionPaneScenario(page) {
  const initial = await page.eval(`(() => {
    const pane = document.querySelector('[data-ppt-layer-pane]')
    const tree = document.querySelector('.ppt-layer-list')
    const rows = [...document.querySelectorAll('[data-ppt-layer-pane-row]')]
    const selected = document.querySelector('[data-selected="true"]')
    const selectedId = selected?.getAttribute('data-ppt-element') ?? ''
    const alternateRow = rows.find((row) => row.getAttribute('data-ppt-layer-pane-row') !== selectedId) ?? null
    const selectedRow = document.querySelector('[data-ppt-layer-pane-row][aria-selected="true"]')

    return {
      alternateRowId: alternateRow?.getAttribute('data-ppt-layer-pane-row') ?? '',
      commandCount: Number(pane?.getAttribute('data-ppt-layer-pane-command-count') ?? 0),
      commandSlot: pane?.getAttribute('data-ppt-layer-pane-command-slot') ?? '',
      commands: pane?.getAttribute('data-ppt-layer-pane-commands') ?? '',
      contractModel: pane?.getAttribute('data-ppt-layer-pane-model') ?? '',
      firstRowKind: rows[0]?.getAttribute('data-ppt-layer-pane-kind') ?? '',
      firstRowOrder: rows[0]?.getAttribute('data-ppt-layer-pane-order') ?? '',
      keyboardIntentModel: tree?.getAttribute('data-ppt-layer-pane-keyboard-intent-model') ?? '',
      keyboardKeys: tree?.getAttribute('data-ppt-layer-pane-keyboard-keys') ?? '',
      keyboardModel: tree?.getAttribute('data-ppt-layer-pane-keyboard-model') ?? '',
      layerCount: document.querySelectorAll('[data-ppt-layer-row]').length,
      pane: !!pane,
      rangeSelectionModel: tree?.getAttribute('data-ppt-layer-pane-range-selection-model') ?? '',
      rowIds: rows.map((row) => row.getAttribute('data-ppt-layer-pane-row') ?? ''),
      rowRole: rows[0]?.getAttribute('role') ?? '',
      selectedId,
      selectedRowId: selectedRow?.getAttribute('data-ppt-layer-pane-row') ?? '',
      selectionModel: tree?.getAttribute('data-ppt-layer-pane-selection-model') ?? '',
      stageCount: document.querySelectorAll('[data-ppt-element]').length,
      treeRole: tree?.getAttribute('role') ?? '',
    }
  })()`)

  record('renders PPT object selection pane', initial.layerCount === initial.stageCount && initial.layerCount > 0, initial)
  record(
    'renders PPT object layer pane descriptor contract',
    initial.pane &&
      initial.contractModel === 'slide-edit-object-layer-pane' &&
      initial.treeRole === 'tree' &&
      initial.rowRole === 'treeitem' &&
      initial.keyboardModel === 'roving-tabindex' &&
      initial.keyboardIntentModel === 'slide-edit-layer-pane-keyboard-intent' &&
      initial.keyboardKeys === 'arrow-left-right-up-down-home-end-enter-space-shift-range-alt-reorder-f2-rename' &&
      initial.rangeSelectionModel === 'row-press-range-anchor' &&
      initial.selectionModel === 'host-controlled-multi-select' &&
      initial.selectedRowId === initial.selectedId &&
      initial.rowIds.length === initial.layerCount &&
      initial.firstRowOrder === '0' &&
      initial.firstRowKind.length > 0,
    initial,
  )
  record(
    'exposes PPT object layer pane command-effect descriptors',
    initial.commandSlot === 'command-effect' &&
      initial.commandCount === 7 &&
      ['select-objects', 'rename-object', 'hide-objects', 'show-objects', 'lock-objects', 'unlock-objects', 'reorder-object'].every((command) => initial.commands.includes(command)),
    initial,
  )

  const layerPaneRangeSelection = await selectPPTLayerPaneRange(page)
  record(
    'range-selects PPT object layer pane rows with Shift click command-effect',
    layerPaneRangeSelection.ok &&
      layerPaneRangeSelection.firstClickOk &&
      layerPaneRangeSelection.shiftClickOk &&
      layerPaneRangeSelection.anchorAfterFirstClick === layerPaneRangeSelection.anchorId &&
      layerPaneRangeSelection.anchorAfterShiftClick === layerPaneRangeSelection.anchorId &&
      layerPaneRangeSelection.expectedActualIds.length >= 2 &&
      layerPaneRangeSelection.selectedStageIds.length === layerPaneRangeSelection.expectedActualIds.length &&
      layerPaneRangeSelection.expectedActualIds.every((objectId) =>
        layerPaneRangeSelection.selectedStageIds.includes(objectId)) &&
      layerPaneRangeSelection.expectedRowIds.every((rowId) =>
        layerPaneRangeSelection.selectedRowIds.includes(rowId)),
    layerPaneRangeSelection,
  )

  const layerPaneAdditiveSelection = await selectPPTLayerPaneAdditive(page)
  record(
    'keeps PPT object layer pane Ctrl click additive selection',
    layerPaneAdditiveSelection.ok &&
      layerPaneAdditiveSelection.firstClickOk &&
      layerPaneAdditiveSelection.additiveClickOk &&
      layerPaneAdditiveSelection.anchorAfterAdditive === layerPaneAdditiveSelection.additiveId &&
      layerPaneAdditiveSelection.selectedStageIds.includes(layerPaneAdditiveSelection.anchorId) &&
      layerPaneAdditiveSelection.selectedStageIds.includes(layerPaneAdditiveSelection.additiveId),
    layerPaneAdditiveSelection,
  )

  const layerPaneKeyboardRangeSelection = await selectPPTLayerPaneKeyboardRange(page)
  record(
    'range-selects PPT object layer pane rows with Shift Arrow keyboard intent',
    layerPaneKeyboardRangeSelection.ok &&
      layerPaneKeyboardRangeSelection.anchorAfterFirstClick === layerPaneKeyboardRangeSelection.anchorId &&
      layerPaneKeyboardRangeSelection.anchorAfterShiftArrow === layerPaneKeyboardRangeSelection.anchorId &&
      layerPaneKeyboardRangeSelection.focusedRowId === layerPaneKeyboardRangeSelection.targetId &&
      layerPaneKeyboardRangeSelection.expectedActualIds.length >= 2 &&
      layerPaneKeyboardRangeSelection.selectedStageIds.length === layerPaneKeyboardRangeSelection.expectedActualIds.length &&
      layerPaneKeyboardRangeSelection.expectedActualIds.every((objectId) =>
        layerPaneKeyboardRangeSelection.selectedStageIds.includes(objectId)) &&
      layerPaneKeyboardRangeSelection.expectedRowIds.every((rowId) =>
        layerPaneKeyboardRangeSelection.selectedRowIds.includes(rowId)),
    layerPaneKeyboardRangeSelection,
  )

  await page.eval(`document.querySelector('[data-ppt-layer-select="${initial.selectedId}"]')?.click()`)
  await delay(50)

  await page.eval(`(() => {
    const row = document.querySelector('[data-ppt-layer-pane-row][aria-selected="true"]')
    row?.focus()
  })()`)
  await delay(30)

  await pressKey(page, {
    code: 'End',
    key: 'End',
    windowsVirtualKeyCode: 35,
  })
  await delay(80)
  const afterLayerEnd = await readPPTLayerPaneKeyboardState(page)

  await pressKey(page, {
    code: 'Home',
    key: 'Home',
    windowsVirtualKeyCode: 36,
  })
  await delay(80)
  const afterLayerHome = await readPPTLayerPaneKeyboardState(page)

  await pressKey(page, {
    code: 'ArrowDown',
    key: 'ArrowDown',
    windowsVirtualKeyCode: 40,
  })
  await delay(80)
  const afterLayerArrowDown = await readPPTLayerPaneKeyboardState(page)

  record(
    'moves PPT object layer pane treeitem selection with Arrow/Home/End keys',
    initial.rowIds.length > 2 &&
      afterLayerEnd.selectedRowId === initial.rowIds.at(-1) &&
      afterLayerEnd.focusedRowId === initial.rowIds.at(-1) &&
      afterLayerEnd.tabStopIds.length === 1 &&
      afterLayerEnd.tabStopIds[0] === initial.rowIds.at(-1) &&
      afterLayerHome.selectedRowId === initial.rowIds[0] &&
      afterLayerHome.focusedRowId === initial.rowIds[0] &&
      afterLayerArrowDown.selectedRowId === initial.rowIds[1] &&
      afterLayerArrowDown.focusedRowId === initial.rowIds[1],
    {
      afterLayerArrowDown,
      afterLayerEnd,
      afterLayerHome,
      initial,
    },
  )

  const layerPaneKeyboardReorder = await reorderPPTLayerPaneObjectWithKeyboard(page)

  record(
    'reorders PPT object layer pane row with Alt Arrow keyboard intent',
    layerPaneKeyboardReorder.ok &&
      layerPaneKeyboardReorder.before.objectRowIds[layerPaneKeyboardReorder.sourceIndex] === layerPaneKeyboardReorder.sourceId &&
      layerPaneKeyboardReorder.before.objectRowIds[layerPaneKeyboardReorder.targetIndex] === layerPaneKeyboardReorder.targetId &&
      layerPaneKeyboardReorder.after.objectRowIds[layerPaneKeyboardReorder.sourceIndex] === layerPaneKeyboardReorder.targetId &&
      layerPaneKeyboardReorder.after.objectRowIds[layerPaneKeyboardReorder.targetIndex] === layerPaneKeyboardReorder.sourceId &&
      layerPaneKeyboardReorder.after.stageOrder[layerPaneKeyboardReorder.targetIndex] === layerPaneKeyboardReorder.sourceId &&
      layerPaneKeyboardReorder.after.selectedId === layerPaneKeyboardReorder.sourceId &&
      layerPaneKeyboardReorder.focusedRowId === layerPaneKeyboardReorder.sourceId,
    layerPaneKeyboardReorder,
  )

  await page.eval(`(() => {
    const rows = [...document.querySelectorAll('[data-ppt-layer-pane-row]')]
    rows.at(-1)?.focus()
  })()`)
  await delay(30)
  await pressKey(page, {
    code: 'Space',
    key: ' ',
    windowsVirtualKeyCode: 32,
  })
  await delay(80)
  const afterLayerSpace = await readPPTLayerPaneKeyboardState(page)

  await page.eval(`(() => {
    const rows = [...document.querySelectorAll('[data-ppt-layer-pane-row]')]
    rows[0]?.focus()
  })()`)
  await delay(30)
  await pressKey(page, {
    code: 'Enter',
    key: 'Enter',
    windowsVirtualKeyCode: 13,
  })
  await delay(80)
  const afterLayerEnter = await readPPTLayerPaneKeyboardState(page)

  record(
    'selects PPT object layer pane treeitem with Space and Enter',
    afterLayerSpace.selectedRowId === afterLayerSpace.rowIds.at(-1) &&
      afterLayerSpace.focusedRowId === afterLayerSpace.rowIds.at(-1) &&
      afterLayerEnter.selectedRowId === afterLayerEnter.rowIds[0] &&
      afterLayerEnter.focusedRowId === afterLayerEnter.rowIds[0] &&
      afterLayerEnter.tabStopIds.length === 1 &&
      afterLayerEnter.tabStopIds[0] === afterLayerEnter.rowIds[0],
    {
      afterLayerEnter,
      afterLayerSpace,
      initial,
    },
  )

  await page.eval(`document.querySelector('[data-ppt-layer-select="${initial.alternateRowId}"]')?.click()`)
  await delay(50)

  const afterLayerSelect = await page.eval(`(() => {
    const selected = document.querySelector('[data-selected="true"]')
    const selectedId = selected?.getAttribute('data-ppt-element') ?? ''
    const row = document.querySelector('[data-ppt-layer-pane-row][aria-selected="true"]')

    return {
      selectedId,
      rowId: row?.getAttribute('data-ppt-layer-pane-row') ?? '',
      rowSelected: row?.getAttribute('data-ppt-layer-pane-selected') ?? '',
    }
  })()`)
  const layerTargetId = afterLayerSelect.selectedId || initial.selectedId

  record(
    'selects PPT object from layer pane command-effect',
    afterLayerSelect.selectedId === initial.alternateRowId &&
      afterLayerSelect.rowId === initial.alternateRowId &&
      afterLayerSelect.rowSelected === 'true',
    {
      afterLayerSelect,
      initial,
    },
  )

  await page.eval(`document.querySelector('[data-ppt-layer-pane-row="${layerTargetId}"]')?.focus()`)
  await delay(30)
  await pressKey(page, {
    code: 'F2',
    key: 'F2',
    windowsVirtualKeyCode: 113,
  })
  await delay(50)
  const afterRenameOpen = await readPPTLayerPaneRenameState(page, layerTargetId)

  await page.eval(`(() => {
    const input = document.querySelector('[data-ppt-layer-pane-rename-input="${layerTargetId}"]')
    const valueSetter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set
    valueSetter.call(input, 'Layer renamed object')
    input.dispatchEvent(new Event('input', { bubbles: true }))
  })()`)
  await delay(30)
  await pressKey(page, {
    code: 'Enter',
    key: 'Enter',
    windowsVirtualKeyCode: 13,
  })
  await delay(80)
  const afterRenameCommit = await readPPTLayerPaneRenameState(page, layerTargetId)

  await page.eval(`document.querySelector('[data-ppt-layer-pane-row="${layerTargetId}"]')?.focus()`)
  await delay(30)
  await pressKey(page, {
    code: 'F2',
    key: 'F2',
    windowsVirtualKeyCode: 113,
  })
  await delay(50)
  await page.eval(`(() => {
    const input = document.querySelector('[data-ppt-layer-pane-rename-input="${layerTargetId}"]')
    const valueSetter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set
    valueSetter.call(input, 'Cancelled layer name')
    input.dispatchEvent(new Event('input', { bubbles: true }))
  })()`)
  await delay(30)
  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(80)
  const afterRenameCancel = await readPPTLayerPaneRenameState(page, layerTargetId)

  record(
    'renames PPT object from layer pane rename command-effect',
    afterRenameOpen.inputOpen &&
      afterRenameOpen.inputValue.length > 0 &&
      afterRenameCommit.inputOpen === false &&
      afterRenameCommit.rowName === 'Layer renamed object' &&
      afterRenameCommit.inspectorName === 'Layer renamed object' &&
      afterRenameCommit.stageName === 'Layer renamed object' &&
      afterRenameCancel.inputOpen === false &&
      afterRenameCancel.rowName === 'Layer renamed object',
    {
      afterRenameCancel,
      afterRenameCommit,
      afterRenameOpen,
    },
  )

  const beforeLayerReorder = await readPPTLayerPaneReorderState(page)
  const layerReorderDrag = await dragPPTLayerPaneRow(page)
  await delay(80)
  const afterLayerReorder = await readPPTLayerPaneReorderState(page)

  record(
    'reorders PPT object from layer pane drag command-effect with drop indicator',
    layerReorderDrag.ok &&
      layerReorderDrag.indicatorModel === 'slide-edit-layer-pane-drop-indicator' &&
      layerReorderDrag.indicatorPlacement === 'before' &&
      layerReorderDrag.indicatorTarget === 'true' &&
      layerReorderDrag.indicatorToIndex === '0' &&
      beforeLayerReorder.draggableObjectRowCount >= 2 &&
      beforeLayerReorder.objectRowIds[0] === layerReorderDrag.targetId &&
      beforeLayerReorder.objectRowIds.at(-1) === layerReorderDrag.sourceId &&
      afterLayerReorder.objectRowIds[0] === layerReorderDrag.sourceId &&
      afterLayerReorder.stageOrder[0] === layerReorderDrag.sourceId &&
      afterLayerReorder.selectedId === layerReorderDrag.sourceId,
    {
      afterLayerReorder,
      beforeLayerReorder,
      layerReorderDrag,
    },
  )

  const beforeLayerReorderToEnd = await readPPTLayerPaneReorderState(page)
  const layerReorderToEndDrag = await dragPPTLayerPaneRow(page, 'after-end')
  await delay(80)
  const afterLayerReorderToEnd = await readPPTLayerPaneReorderState(page)

  record(
    'moves PPT object to layer pane end with after drop indicator',
    layerReorderToEndDrag.ok &&
      layerReorderToEndDrag.indicatorModel === 'slide-edit-layer-pane-drop-indicator' &&
      layerReorderToEndDrag.indicatorPlacement === 'after' &&
      layerReorderToEndDrag.indicatorTarget === 'true' &&
      Number(layerReorderToEndDrag.indicatorToIndex) === beforeLayerReorderToEnd.objectRowIds.length &&
      beforeLayerReorderToEnd.objectRowIds[0] === layerReorderToEndDrag.sourceId &&
      beforeLayerReorderToEnd.objectRowIds.at(-1) === layerReorderToEndDrag.targetId &&
      afterLayerReorderToEnd.objectRowIds.at(-1) === layerReorderToEndDrag.sourceId &&
      afterLayerReorderToEnd.stageOrder.at(-1) === layerReorderToEndDrag.sourceId &&
      afterLayerReorderToEnd.selectedId === layerReorderToEndDrag.sourceId,
    {
      afterLayerReorderToEnd,
      beforeLayerReorderToEnd,
      layerReorderToEndDrag,
    },
  )
  await page.eval(`document.querySelector('[data-ppt-layer-select="${layerTargetId}"]')?.click()`)
  await delay(50)

  await page.eval(`(() => {
    const name = document.querySelector('[data-ppt-style-field="name"]')
    const valueSetter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set
    valueSetter.call(name, 'Locked accent')
    name.dispatchEvent(new Event('input', { bubbles: true }))
    name.dispatchEvent(new Event('change', { bubbles: true }))
  })()`)
  await delay(50)

  const afterRename = await page.eval(`(() => {
    const row = document.querySelector('[data-ppt-layer-row][aria-selected="true"]')

    return {
      inspectorName: document.querySelector('[data-ppt-style-field="name"]')?.value ?? '',
      rowName: row?.querySelector('.ppt-layer-name')?.textContent ?? '',
    }
  })()`)

  record('renames selected PPT object from inspector', afterRename.inspectorName === 'Locked accent' && afterRename.rowName === 'Locked accent', afterRename)

  await page.eval(`document.querySelector('[data-ppt-layer-row][aria-selected="true"] [data-ppt-layer-lock]').click()`)
  await delay(50)

  const afterLayerLock = await page.eval(`(() => {
    const selected = document.querySelector('[data-ppt-element="${layerTargetId}"]')
    const row = document.querySelector('[data-ppt-layer-row="${layerTargetId}"]')
    const visibilityButton = row?.querySelector('[data-ppt-layer-visibility]')

    return {
      deleteDisabled: document.querySelector('button[title="Delete"]')?.disabled ?? false,
      lockDisabled: document.querySelector('[data-ppt-command="lock-selection"]')?.disabled ?? false,
      locked: selected?.getAttribute('data-locked') ?? null,
      rowLayerBlockReason: row?.getAttribute('data-ppt-object-visibility-layer-selection-block-reason') ?? '',
      rowPaneLocked: row?.getAttribute('data-ppt-layer-pane-locked') ?? null,
      resizeHandleCount: document.querySelectorAll('.ppt-resize-handle').length,
      rowLocked: row?.getAttribute('data-locked') ?? null,
      rowStageBlockReason: row?.getAttribute('data-ppt-object-visibility-stage-selection-block-reason') ?? '',
      rowVisibilityAvailability: visibilityButton?.getAttribute('data-ppt-object-visibility-availability') ?? '',
      rowVisibilityModel: row?.getAttribute('data-ppt-object-visibility-model') ?? '',
      rowVisibilitySelectable: row?.getAttribute('data-ppt-object-visibility-selectable') ?? '',
      rowVisibilityUnavailable: visibilityButton?.getAttribute('data-ppt-object-visibility-unavailable') ?? '',
      unlockDisabled: document.querySelector('[data-ppt-command="unlock-all"]')?.disabled ?? true,
    }
  })()`)

  record('locks selected PPT object from selection pane', afterLayerLock.locked === 'true' && afterLayerLock.rowLocked === 'true' && afterLayerLock.rowPaneLocked === 'true' && afterLayerLock.resizeHandleCount === 0, afterLayerLock)
  record('disables transform commands for locked PPT object', afterLayerLock.deleteDisabled && afterLayerLock.lockDisabled && !afterLayerLock.unlockDisabled, afterLayerLock)
  record('exposes locked PPT object visibility unavailable metadata',
    afterLayerLock.rowVisibilityModel === 'slide-edit-object-visibility' &&
      afterLayerLock.rowVisibilityAvailability === 'false' &&
      afterLayerLock.rowVisibilityUnavailable === 'locked-selection' &&
      afterLayerLock.rowLayerBlockReason === 'locked' &&
      afterLayerLock.rowStageBlockReason === 'locked' &&
      afterLayerLock.rowVisibilitySelectable === 'false',
    afterLayerLock)

  await page.eval(`document.querySelector('[data-ppt-command="unlock-all"]').click()`)
  await delay(50)

  const afterUnlockAll = await page.eval(`(() => {
    const selected = document.querySelector('[data-ppt-element="${layerTargetId}"]')
    const row = document.querySelector('[data-ppt-layer-row="${layerTargetId}"]')

    return {
      locked: selected?.getAttribute('data-locked') ?? null,
      rowPaneLocked: row?.getAttribute('data-ppt-layer-pane-locked') ?? null,
      resizeHandleCount: document.querySelectorAll('.ppt-resize-handle').length,
      rowLocked: row?.getAttribute('data-locked') ?? null,
    }
  })()`)

  record('unlocks all PPT objects from toolbar', afterUnlockAll.locked === 'false' && afterUnlockAll.rowLocked === 'false' && afterUnlockAll.rowPaneLocked === 'false' && afterUnlockAll.resizeHandleCount > 0, afterUnlockAll)

  await page.eval(`document.querySelector('[data-ppt-layer-row="${layerTargetId}"] [data-ppt-layer-visibility]').click()`)
  await delay(50)

  const afterHide = await page.eval(`(() => {
    const stage = document.querySelector('.ppt-stage-shell')
    const stageElement = document.querySelector('[data-ppt-element="${layerTargetId}"]')
    const row = document.querySelector('[data-ppt-layer-row="${layerTargetId}"]')
    const visibilityButton = row?.querySelector('[data-ppt-layer-visibility]')

    return {
      command: stage?.getAttribute('data-ppt-object-visibility-command') ?? '',
      commandObjects: stage?.getAttribute('data-ppt-object-visibility-command-objects') ?? '',
      commandSelection: stage?.getAttribute('data-ppt-object-visibility-command-selection') ?? '',
      commandSlide: stage?.getAttribute('data-ppt-object-visibility-command-slide') ?? '',
      commandTargetCount: stage?.getAttribute('data-ppt-object-visibility-command-target-count') ?? '',
      commandType: stage?.getAttribute('data-ppt-object-visibility-command-type') ?? '',
      hidden: row?.getAttribute('data-hidden') ?? null,
      model: stage?.getAttribute('data-ppt-object-visibility-model') ?? '',
      rowAvailability: visibilityButton?.getAttribute('data-ppt-object-visibility-availability') ?? '',
      rowCommand: visibilityButton?.getAttribute('data-ppt-object-visibility-command') ?? '',
      rowModel: row?.getAttribute('data-ppt-object-visibility-model') ?? '',
      rowPaneHidden: row?.getAttribute('data-ppt-layer-pane-hidden') ?? null,
      rowSelected: row?.getAttribute('aria-selected') ?? null,
      rowStageBlockReason: row?.getAttribute('data-ppt-object-visibility-stage-selection-block-reason') ?? '',
      rowTargets: visibilityButton?.getAttribute('data-ppt-object-visibility-targets') ?? '',
      rowVisible: row?.getAttribute('data-ppt-object-visibility-visible') ?? '',
      stageElementExists: !!stageElement,
    }
  })()`)

  record('hides selected PPT object from selection pane', afterHide.hidden === 'true' && afterHide.rowPaneHidden === 'true' && afterHide.rowSelected === 'true' && !afterHide.stageElementExists, afterHide)
  record('routes PPT object hide through slide-edit visibility command effect',
    afterHide.command === 'hide-objects' &&
      afterHide.commandObjects === layerTargetId &&
      afterHide.commandSelection === layerTargetId &&
      afterHide.commandSlide === 'slide-1' &&
      afterHide.commandTargetCount === '1' &&
      afterHide.commandType === 'slide-command-effect' &&
      afterHide.model === 'slide-edit-object-visibility' &&
      afterHide.rowModel === 'slide-edit-object-visibility' &&
      afterHide.rowCommand === 'show-objects' &&
      afterHide.rowAvailability === 'true' &&
      afterHide.rowTargets === layerTargetId &&
      afterHide.rowVisible === 'false' &&
      afterHide.rowStageBlockReason === 'hidden',
    afterHide)

  await page.eval(`document.querySelector('[data-ppt-layer-row="${layerTargetId}"] [data-ppt-layer-visibility]').click()`)
  await delay(50)

  const afterShow = await page.eval(`(() => {
    const stage = document.querySelector('.ppt-stage-shell')
    const stageElement = document.querySelector('[data-ppt-element="${layerTargetId}"]')
    const row = document.querySelector('[data-ppt-layer-row="${layerTargetId}"]')
    const visibilityButton = row?.querySelector('[data-ppt-layer-visibility]')

    return {
      command: stage?.getAttribute('data-ppt-object-visibility-command') ?? '',
      commandObjects: stage?.getAttribute('data-ppt-object-visibility-command-objects') ?? '',
      commandSelection: stage?.getAttribute('data-ppt-object-visibility-command-selection') ?? '',
      commandSlide: stage?.getAttribute('data-ppt-object-visibility-command-slide') ?? '',
      commandTargetCount: stage?.getAttribute('data-ppt-object-visibility-command-target-count') ?? '',
      commandType: stage?.getAttribute('data-ppt-object-visibility-command-type') ?? '',
      hidden: row?.getAttribute('data-hidden') ?? null,
      model: stage?.getAttribute('data-ppt-object-visibility-model') ?? '',
      rowAvailability: visibilityButton?.getAttribute('data-ppt-object-visibility-availability') ?? '',
      rowCommand: visibilityButton?.getAttribute('data-ppt-object-visibility-command') ?? '',
      rowModel: row?.getAttribute('data-ppt-object-visibility-model') ?? '',
      rowPaneHidden: row?.getAttribute('data-ppt-layer-pane-hidden') ?? null,
      rowStageBlockReason: row?.getAttribute('data-ppt-object-visibility-stage-selection-block-reason') ?? '',
      rowTargets: visibilityButton?.getAttribute('data-ppt-object-visibility-targets') ?? '',
      rowVisible: row?.getAttribute('data-ppt-object-visibility-visible') ?? '',
      stageElementExists: !!stageElement,
    }
  })()`)

  record('shows hidden PPT object from selection pane', afterShow.hidden === 'false' && afterShow.rowPaneHidden === 'false' && afterShow.stageElementExists, afterShow)
  record('routes PPT object show through slide-edit visibility command effect',
    afterShow.command === 'show-objects' &&
      afterShow.commandObjects === layerTargetId &&
      afterShow.commandSelection === layerTargetId &&
      afterShow.commandSlide === 'slide-1' &&
      afterShow.commandTargetCount === '1' &&
      afterShow.commandType === 'slide-command-effect' &&
      afterShow.model === 'slide-edit-object-visibility' &&
      afterShow.rowModel === 'slide-edit-object-visibility' &&
      afterShow.rowCommand === 'hide-objects' &&
      afterShow.rowAvailability === 'true' &&
      afterShow.rowTargets === layerTargetId &&
      afterShow.rowVisible === 'true' &&
      afterShow.rowStageBlockReason === '',
    afterShow)

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      objectState: {
        locked: true,
        visible: false,
      },
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const afterObjectStateHideLock = await readPPTObjectStateImportState(
    page,
    layerTargetId,
  )

  record(
    'pastes JSON object visible and locked state into selected PPT object',
    afterObjectStateHideLock.model === 'ppt-object-state-import' &&
      afterObjectStateHideLock.format === 'application-json-ppt-object-state' &&
      afterObjectStateHideLock.fields === 'visible locked' &&
      afterObjectStateHideLock.commands === 'hide-objects lock-objects' &&
      afterObjectStateHideLock.commandTypes ===
        'slide-command-effect slide-command-effect' &&
      afterObjectStateHideLock.importVisible === 'false' &&
      afterObjectStateHideLock.importLocked === 'true' &&
      afterObjectStateHideLock.objectIds === layerTargetId &&
      afterObjectStateHideLock.visibilityTargets === layerTargetId &&
      afterObjectStateHideLock.lockTargets === layerTargetId &&
      afterObjectStateHideLock.slide === 'slide-1' &&
      afterObjectStateHideLock.jsonLength > 40 &&
      afterObjectStateHideLock.hidden === 'true' &&
      afterObjectStateHideLock.locked === 'true' &&
      afterObjectStateHideLock.rowSelected === 'true' &&
      !afterObjectStateHideLock.stageElementExists &&
      afterObjectStateHideLock.visibilityCommand === 'hide-objects',
    {
      afterObjectStateHideLock,
      afterShow,
    },
  )

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      objectState: {
        locked: false,
        visible: true,
      },
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const afterObjectStateShowUnlock = await readPPTObjectStateImportState(
    page,
    layerTargetId,
  )

  await page.eval(`document.querySelector('button[title="Undo"]')?.click()`)
  await delay(100)

  const afterObjectStateShowUnlockUndo = await readPPTObjectStateImportState(
    page,
    layerTargetId,
  )

  await page.eval(`document.querySelector('button[title="Redo"]')?.click()`)
  await delay(100)

  const afterObjectStateShowUnlockRedo = await readPPTObjectStateImportState(
    page,
    layerTargetId,
  )

  record(
    'pastes JSON object unlock and show state through one PPT history step',
    afterObjectStateShowUnlock.model === 'ppt-object-state-import' &&
      afterObjectStateShowUnlock.commands === 'unlock-objects show-objects' &&
      afterObjectStateShowUnlock.importVisible === 'true' &&
      afterObjectStateShowUnlock.importLocked === 'false' &&
      afterObjectStateShowUnlock.visibilityTargets === layerTargetId &&
      afterObjectStateShowUnlock.lockTargets === layerTargetId &&
      afterObjectStateShowUnlock.hidden === 'false' &&
      afterObjectStateShowUnlock.locked === 'false' &&
      afterObjectStateShowUnlock.stageElementExists &&
      afterObjectStateShowUnlock.stageSelected === 'true' &&
      afterObjectStateShowUnlockUndo.hidden === 'true' &&
      afterObjectStateShowUnlockUndo.locked === 'true' &&
      !afterObjectStateShowUnlockUndo.stageElementExists &&
      afterObjectStateShowUnlockRedo.hidden === 'false' &&
      afterObjectStateShowUnlockRedo.locked === 'false' &&
      afterObjectStateShowUnlockRedo.stageElementExists,
    {
      afterObjectStateHideLock,
      afterObjectStateShowUnlock,
      afterObjectStateShowUnlockRedo,
      afterObjectStateShowUnlockUndo,
    },
  )

  const beforeObjectLayerImport = await readPPTObjectLayerImportState(
    page,
    layerTargetId,
  )

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      objectLayer: {
        position: 'front',
      },
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const afterObjectLayerFront = await readPPTObjectLayerImportState(
    page,
    layerTargetId,
  )

  record(
    'pastes JSON object layer front state through slide-edit layer pane reorder effect',
    beforeObjectLayerImport.layerOrder.length >= 2 &&
      afterObjectLayerFront.model === 'ppt-object-layer-import' &&
      afterObjectLayerFront.format === 'application-json-ppt-object-layer' &&
      afterObjectLayerFront.fields === 'position' &&
      afterObjectLayerFront.command === 'reorder-object' &&
      afterObjectLayerFront.commandType === 'slide-command-effect' &&
      afterObjectLayerFront.fromIndex === beforeObjectLayerImport.layerIndex &&
      afterObjectLayerFront.toIndex === beforeObjectLayerImport.layerOrder.length &&
      afterObjectLayerFront.objectId === layerTargetId &&
      afterObjectLayerFront.position === 'front' &&
      afterObjectLayerFront.slide === 'slide-1' &&
      afterObjectLayerFront.jsonLength > 30 &&
      afterObjectLayerFront.layerOrder.at(-1) === layerTargetId &&
      afterObjectLayerFront.stageOrder.at(-1) === layerTargetId &&
      afterObjectLayerFront.rowSelected === 'true' &&
      afterObjectLayerFront.stageSelected === 'true',
    {
      afterObjectLayerFront,
      beforeObjectLayerImport,
    },
  )

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      zOrder: {
        position: 'send-to-back',
      },
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(120)

  const afterObjectLayerBack = await readPPTObjectLayerImportState(
    page,
    layerTargetId,
  )

  await page.eval(`document.querySelector('button[title="Undo"]')?.click()`)
  await delay(100)

  const afterObjectLayerBackUndo = await readPPTObjectLayerImportState(
    page,
    layerTargetId,
  )

  await page.eval(`document.querySelector('button[title="Redo"]')?.click()`)
  await delay(100)

  const afterObjectLayerBackRedo = await readPPTObjectLayerImportState(
    page,
    layerTargetId,
  )

  record(
    'pastes JSON object layer back state through one PPT history step',
    afterObjectLayerBack.model === 'ppt-object-layer-import' &&
      afterObjectLayerBack.fields === 'position' &&
      afterObjectLayerBack.command === 'reorder-object' &&
      afterObjectLayerBack.fromIndex === afterObjectLayerFront.layerOrder.length - 1 &&
      afterObjectLayerBack.toIndex === 0 &&
      afterObjectLayerBack.objectId === layerTargetId &&
      afterObjectLayerBack.position === 'back' &&
      afterObjectLayerBack.layerOrder[0] === layerTargetId &&
      afterObjectLayerBack.stageOrder[0] === layerTargetId &&
      afterObjectLayerBackUndo.layerOrder.at(-1) === layerTargetId &&
      afterObjectLayerBackUndo.stageOrder.at(-1) === layerTargetId &&
      afterObjectLayerBackRedo.layerOrder[0] === layerTargetId &&
      afterObjectLayerBackRedo.stageOrder[0] === layerTargetId,
    {
      afterObjectLayerBack,
      afterObjectLayerBackRedo,
      afterObjectLayerBackUndo,
      afterObjectLayerFront,
    },
  )
}

async function readPPTLayerPaneKeyboardState(page) {
  return page.eval(`(() => {
    const rows = [...document.querySelectorAll('[data-ppt-layer-pane-row]')]
    const focusedRow = document.activeElement?.closest('[data-ppt-layer-pane-row]')
    const selectedRows = rows.filter((row) => row.getAttribute('aria-selected') === 'true')
    const tabStopRows = rows.filter((row) => row.tabIndex === 0)

    return {
      focusedRowId: focusedRow?.getAttribute('data-ppt-layer-pane-row') ?? '',
      rowIds: rows.map((row) => row.getAttribute('data-ppt-layer-pane-row') ?? ''),
      selectedRowId: selectedRows[0]?.getAttribute('data-ppt-layer-pane-row') ?? '',
      selectedRowIds: selectedRows.map((row) => row.getAttribute('data-ppt-layer-pane-row') ?? ''),
      tabStopIds: tabStopRows.map((row) => row.getAttribute('data-ppt-layer-pane-row') ?? ''),
      treeKeyboardIntentModel: document.querySelector('.ppt-layer-list')?.getAttribute('data-ppt-layer-pane-keyboard-intent-model') ?? '',
      treeKeyboardKeys: document.querySelector('.ppt-layer-list')?.getAttribute('data-ppt-layer-pane-keyboard-keys') ?? '',
    }
  })()`)
}

async function clickPPTLayerPaneSelect(page, objectId, modifiers = {}) {
  return page.eval(`((objectId, modifiers) => {
    const button = document.querySelector(\`[data-ppt-layer-select="\${objectId}"]\`)

    if (!(button instanceof HTMLElement)) {
      return false
    }

    button.dispatchEvent(new MouseEvent('click', {
      bubbles: true,
      cancelable: true,
      ctrlKey: Boolean(modifiers.ctrlKey),
      metaKey: Boolean(modifiers.metaKey),
      shiftKey: Boolean(modifiers.shiftKey),
    }))

    return true
  })(${JSON.stringify(objectId)}, ${JSON.stringify(modifiers)})`)
}

async function readPPTLayerPaneClickSelectionState(page) {
  return page.eval(`(() => {
    const rows = [...document.querySelectorAll('[data-ppt-layer-pane-row]')]

    return {
      rangeAnchorId: document.querySelector('.ppt-layer-list')?.getAttribute('data-ppt-layer-pane-range-anchor-object-id') ?? '',
      selectedRowIds: rows
        .filter((row) => row.getAttribute('aria-selected') === 'true')
        .map((row) => row.getAttribute('data-ppt-layer-pane-row') ?? ''),
      selectedStageIds: [...document.querySelectorAll('[data-selected="true"]')]
        .map((element) => element.getAttribute('data-ppt-element') ?? ''),
    }
  })()`)
}

async function getPPTLayerPaneRangeTargets(page) {
  return page.eval(`(() => {
    const rows = [...document.querySelectorAll('[data-ppt-layer-pane-row]')]
      .filter((row) => row.getAttribute('aria-disabled') !== 'true')
    const visibleObjectRows = rows.filter((row) =>
      row.getAttribute('data-ppt-layer-pane-row-type') === 'object' &&
        row.getAttribute('data-ppt-layer-pane-hidden') !== 'true')
    const anchor = visibleObjectRows[0]
    const target = visibleObjectRows[Math.min(2, visibleObjectRows.length - 1)]

    if (!(anchor instanceof HTMLElement) || !(target instanceof HTMLElement) || anchor === target) {
      return {
        ok: false,
        rowCount: rows.length,
        visibleObjectRowCount: visibleObjectRows.length,
      }
    }

    const anchorIndex = rows.indexOf(anchor)
    const targetIndex = rows.indexOf(target)
    const [fromIndex, toIndex] = anchorIndex < targetIndex
      ? [anchorIndex, targetIndex]
      : [targetIndex, anchorIndex]
    const expectedRows = rows.slice(fromIndex, toIndex + 1)
    const expectedActualIds = new Set()

    for (const row of expectedRows) {
      const rowId = row.getAttribute('data-ppt-layer-pane-row') ?? ''

      if (row.getAttribute('data-ppt-layer-pane-row-type') === 'group') {
        rows
          .filter((candidate) =>
            candidate.getAttribute('data-ppt-layer-pane-parent-object-id') === rowId &&
              candidate.getAttribute('data-ppt-layer-pane-row-type') === 'object')
          .forEach((child) => expectedActualIds.add(child.getAttribute('data-ppt-layer-pane-row') ?? ''))
        continue
      }

      expectedActualIds.add(rowId)
    }

    return {
      anchorId: anchor.getAttribute('data-ppt-layer-pane-row') ?? '',
      expectedActualIds: [...expectedActualIds].filter(Boolean),
      expectedRowIds: expectedRows.map((row) => row.getAttribute('data-ppt-layer-pane-row') ?? ''),
      ok: true,
      targetId: target.getAttribute('data-ppt-layer-pane-row') ?? '',
    }
  })()`)
}

async function selectPPTLayerPaneRange(page) {
  const targets = await getPPTLayerPaneRangeTargets(page)

  if (!targets.ok) {
    return targets
  }

  const firstClickOk = await clickPPTLayerPaneSelect(page, targets.anchorId)
  await delay(50)
  const afterFirstClick = await readPPTLayerPaneClickSelectionState(page)
  const shiftClickOk = await clickPPTLayerPaneSelect(page, targets.targetId, {
    shiftKey: true,
  })
  await delay(80)
  const afterShiftClick = await readPPTLayerPaneClickSelectionState(page)

  return {
    ...targets,
    anchorAfterFirstClick: afterFirstClick.rangeAnchorId,
    anchorAfterShiftClick: afterShiftClick.rangeAnchorId,
    firstClickOk,
    selectedRowIds: afterShiftClick.selectedRowIds,
    selectedStageIds: afterShiftClick.selectedStageIds,
    shiftClickOk,
  }
}

async function selectPPTLayerPaneAdditive(page) {
  const targets = await page.eval(`(() => {
    const objectRows = [...document.querySelectorAll('[data-ppt-layer-pane-row-type="object"]')]
      .filter((row) =>
        row.getAttribute('aria-disabled') !== 'true' &&
          row.getAttribute('data-ppt-layer-pane-hidden') !== 'true')
    const anchor = objectRows[0]
    const additive = objectRows[1]

    if (!(anchor instanceof HTMLElement) || !(additive instanceof HTMLElement)) {
      return {
        ok: false,
        objectRowCount: objectRows.length,
      }
    }

    return {
      additiveId: additive.getAttribute('data-ppt-layer-pane-row') ?? '',
      anchorId: anchor.getAttribute('data-ppt-layer-pane-row') ?? '',
      ok: true,
    }
  })()`)

  if (!targets.ok) {
    return targets
  }

  const firstClickOk = await clickPPTLayerPaneSelect(page, targets.anchorId)
  await delay(50)
  const additiveClickOk = await clickPPTLayerPaneSelect(page, targets.additiveId, {
    ctrlKey: true,
  })
  await delay(80)
  const afterAdditive = await readPPTLayerPaneClickSelectionState(page)

  return {
    ...targets,
    additiveClickOk,
    anchorAfterAdditive: afterAdditive.rangeAnchorId,
    firstClickOk,
    selectedRowIds: afterAdditive.selectedRowIds,
    selectedStageIds: afterAdditive.selectedStageIds,
  }
}

async function getPPTLayerPaneKeyboardRangeTargets(page) {
  return page.eval(`(() => {
    const rows = [...document.querySelectorAll('[data-ppt-layer-pane-row]')]
      .filter((row) => row.getAttribute('aria-disabled') !== 'true')
    let anchor = null
    let target = null

    for (let index = 0; index < rows.length - 1; index += 1) {
      const current = rows[index]
      const next = rows[index + 1]

      if (
        current.getAttribute('data-ppt-layer-pane-row-type') === 'object' &&
        next.getAttribute('data-ppt-layer-pane-row-type') === 'object' &&
        current.getAttribute('data-ppt-layer-pane-hidden') !== 'true' &&
        next.getAttribute('data-ppt-layer-pane-hidden') !== 'true'
      ) {
        anchor = current
        target = next
        break
      }
    }

    if (!(anchor instanceof HTMLElement) || !(target instanceof HTMLElement)) {
      return {
        ok: false,
        rowCount: rows.length,
      }
    }

    return {
      anchorId: anchor.getAttribute('data-ppt-layer-pane-row') ?? '',
      expectedActualIds: [
        anchor.getAttribute('data-ppt-layer-pane-row') ?? '',
        target.getAttribute('data-ppt-layer-pane-row') ?? '',
      ].filter(Boolean),
      expectedRowIds: [
        anchor.getAttribute('data-ppt-layer-pane-row') ?? '',
        target.getAttribute('data-ppt-layer-pane-row') ?? '',
      ].filter(Boolean),
      ok: true,
      targetId: target.getAttribute('data-ppt-layer-pane-row') ?? '',
    }
  })()`)
}

async function selectPPTLayerPaneKeyboardRange(page) {
  const targets = await getPPTLayerPaneKeyboardRangeTargets(page)

  if (!targets.ok) {
    return targets
  }

  const firstClickOk = await clickPPTLayerPaneSelect(page, targets.anchorId)
  await delay(50)
  await page.eval(`((anchorId) => {
    document.querySelector(\`[data-ppt-layer-pane-row="\${anchorId}"]\`)?.focus()
  })(${JSON.stringify(targets.anchorId)})`)
  await delay(30)
  const afterFirstClick = await readPPTLayerPaneClickSelectionState(page)
  await pressKey(page, {
    code: 'ArrowDown',
    key: 'ArrowDown',
    modifiers: 8,
    windowsVirtualKeyCode: 40,
  })
  await delay(80)
  const afterShiftArrow = await readPPTLayerPaneClickSelectionState(page)
  const focusedRowId = await page.eval(`(() =>
    document.activeElement?.closest('[data-ppt-layer-pane-row]')?.getAttribute('data-ppt-layer-pane-row') ?? ''
  )()`)

  return {
    ...targets,
    anchorAfterFirstClick: afterFirstClick.rangeAnchorId,
    anchorAfterShiftArrow: afterShiftArrow.rangeAnchorId,
    firstClickOk,
    focusedRowId,
    selectedRowIds: afterShiftArrow.selectedRowIds,
    selectedStageIds: afterShiftArrow.selectedStageIds,
  }
}

async function readPPTLayerPaneRenameState(page, objectId) {
  return page.eval(`((objectId) => {
    const row = document.querySelector(\`[data-ppt-layer-pane-row="\${objectId}"]\`)
    const input = document.querySelector(\`[data-ppt-layer-pane-rename-input="\${objectId}"]\`)
    const inspectorName = document.querySelector('[data-ppt-style-field="name"]')
    const stageElement = document.querySelector(\`[data-ppt-element="\${objectId}"]\`)

    return {
      focusedRowId: document.activeElement?.closest('[data-ppt-layer-pane-row]')?.getAttribute('data-ppt-layer-pane-row') ?? '',
      inputOpen: !!input,
      inputValue: input?.value ?? '',
      inspectorName: inspectorName?.value ?? '',
      renamable: row?.getAttribute('data-ppt-layer-pane-renamable') ?? '',
      rowName: row?.querySelector('.ppt-layer-name')?.textContent ?? '',
      stageName: stageElement?.getAttribute('data-ppt-element-name') ?? '',
    }
  })(${JSON.stringify(objectId)})`)
}

async function reorderPPTLayerPaneObjectWithKeyboard(page) {
  const before = await readPPTLayerPaneReorderState(page)
  const target = await page.eval(`(() => {
    const rows = [...document.querySelectorAll('[data-ppt-layer-pane-row-type="object"][data-ppt-layer-pane-draggable="true"]')]
    const source = rows.find((row, index) => index < rows.length - 1)
    const sourceIndex = source ? rows.indexOf(source) : -1
    const next = sourceIndex >= 0 ? rows[sourceIndex + 1] : null

    if (!(source instanceof HTMLElement) || !(next instanceof HTMLElement)) {
      return {
        ok: false,
        rowCount: rows.length,
        sourceFound: source instanceof HTMLElement,
        targetFound: next instanceof HTMLElement,
      }
    }

    source.querySelector('[data-ppt-layer-select]')?.click()
    source.focus()

    return {
      ok: true,
      sourceId: source.getAttribute('data-ppt-layer-pane-row') ?? '',
      sourceIndex,
      targetId: next.getAttribute('data-ppt-layer-pane-row') ?? '',
      targetIndex: sourceIndex + 1,
    }
  })()`)

  if (!target.ok) {
    return {
      ...target,
      before,
    }
  }

  await delay(50)
  await pressKey(page, {
    code: 'ArrowDown',
    key: 'ArrowDown',
    modifiers: 1,
    windowsVirtualKeyCode: 40,
  })
  await delay(100)
  const after = await readPPTLayerPaneReorderState(page)
  const focusedRowId = await page.eval(`(() =>
    document.activeElement?.closest('[data-ppt-layer-pane-row]')?.getAttribute('data-ppt-layer-pane-row') ?? ''
  )()`)

  return {
    ...target,
    after,
    before,
    focusedRowId,
  }
}

async function readPPTLayerPaneReorderState(page) {
  return page.eval(`(() => {
    const objectRows = [...document.querySelectorAll('[data-ppt-layer-pane-row-type="object"]')]
    const draggableObjectRows = objectRows.filter((row) =>
      row.getAttribute('data-ppt-layer-pane-draggable') === 'true')

    return {
      draggableObjectRowCount: draggableObjectRows.length,
      objectRowIds: objectRows.map((row) => row.getAttribute('data-ppt-layer-pane-row') ?? ''),
      selectedId: document.querySelector('[data-selected="true"]')?.getAttribute('data-ppt-element') ?? '',
      stageOrder: [...document.querySelectorAll('[data-ppt-element]')]
        .map((element) => element.getAttribute('data-ppt-element') ?? ''),
    }
  })()`)
}

async function dragPPTLayerPaneRow(page, mode = 'before-start') {
  const started = await page.eval(`((mode) => {
    const rows = [...document.querySelectorAll('[data-ppt-layer-pane-row-type="object"][data-ppt-layer-pane-draggable="true"]')]
    const source = mode === 'after-end' ? rows[0] : rows.at(-1)
    const target = mode === 'after-end' ? rows.at(-1) : rows[0]

    if (!(source instanceof HTMLElement) || !(target instanceof HTMLElement) || source === target) {
      return {
        ok: false,
        mode,
        rowCount: rows.length,
        sourceFound: source instanceof HTMLElement,
        targetFound: target instanceof HTMLElement,
      }
    }

    source.querySelector('[data-ppt-layer-select]')?.click()

    const rect = target.getBoundingClientRect()
    const clientX = rect.left + rect.width / 2
    const clientY = mode === 'after-end' ? rect.bottom - 2 : rect.top + 2
    const dataTransfer = typeof DataTransfer === 'function'
      ? new DataTransfer()
      : {
          data: new Map(),
          dropEffect: 'move',
          effectAllowed: 'move',
          getData(type) {
            return this.data.get(type) ?? ''
          },
          setData(type, value) {
            this.data.set(type, value)
          },
        }

    function createDragEvent(type) {
      let event

      try {
        event = new DragEvent(type, {
          bubbles: true,
          cancelable: true,
          clientX,
          clientY,
          dataTransfer,
        })
      } catch {
        event = new Event(type, {
          bubbles: true,
          cancelable: true,
        })
      }

      if (!event.dataTransfer) {
        Object.defineProperty(event, 'dataTransfer', {
          configurable: true,
          value: dataTransfer,
        })
      }
      if (event.clientX !== clientX) {
        Object.defineProperty(event, 'clientX', {
          configurable: true,
          value: clientX,
        })
      }
      if (event.clientY !== clientY) {
        Object.defineProperty(event, 'clientY', {
          configurable: true,
          value: clientY,
        })
      }

      return event
    }

    const sourceId = source.getAttribute('data-ppt-layer-pane-row') ?? ''
    const targetId = target.getAttribute('data-ppt-layer-pane-row') ?? ''

    source.dispatchEvent(createDragEvent('dragstart'))
    target.dispatchEvent(createDragEvent('dragover'))

    window.__pptLayerPaneObjectDrag = {
      clientX,
      clientY,
      dataTransfer,
      mode,
      sourceId,
      targetId,
    }

    return {
      mode,
      ok: true,
      pending: true,
      sourceId,
      targetId,
    }
  })(${JSON.stringify(mode)})`)

  if (!started.ok) {
    return started
  }

  await delay(30)

  return page.eval(`(() => {
    const drag = window.__pptLayerPaneObjectDrag

    if (!drag) {
      return {
        ok: false,
        pendingFound: false,
      }
    }

    const source = document.querySelector(\`[data-ppt-layer-pane-row="\${drag.sourceId}"]\`)
    const target = document.querySelector(\`[data-ppt-layer-pane-row="\${drag.targetId}"]\`)

    if (!(source instanceof HTMLElement) || !(target instanceof HTMLElement)) {
      delete window.__pptLayerPaneObjectDrag
      return {
        ok: false,
        sourceFound: source instanceof HTMLElement,
        targetFound: target instanceof HTMLElement,
      }
    }

    function createDragEvent(type) {
      let event

      try {
        event = new DragEvent(type, {
          bubbles: true,
          cancelable: true,
          clientX: drag.clientX,
          clientY: drag.clientY,
          dataTransfer: drag.dataTransfer,
        })
      } catch {
        event = new Event(type, {
          bubbles: true,
          cancelable: true,
        })
      }

      if (!event.dataTransfer) {
        Object.defineProperty(event, 'dataTransfer', {
          configurable: true,
          value: drag.dataTransfer,
        })
      }
      if (event.clientX !== drag.clientX) {
        Object.defineProperty(event, 'clientX', {
          configurable: true,
          value: drag.clientX,
        })
      }
      if (event.clientY !== drag.clientY) {
        Object.defineProperty(event, 'clientY', {
          configurable: true,
          value: drag.clientY,
        })
      }

      return event
    }

    const indicatorModel = target.getAttribute('data-ppt-layer-pane-drop-indicator-model') ?? ''
    const indicatorPlacement = target.getAttribute('data-ppt-layer-pane-drop-indicator') ?? ''
    const indicatorTarget = target.getAttribute('data-ppt-layer-pane-drop-target') ?? ''
    const indicatorToIndex = target.getAttribute('data-ppt-layer-pane-drop-to-index') ?? ''
    target.dispatchEvent(createDragEvent('drop'))
    source.dispatchEvent(createDragEvent('dragend'))
    delete window.__pptLayerPaneObjectDrag

    return {
      indicatorModel,
      indicatorPlacement,
      indicatorTarget,
      indicatorToIndex,
      mode: drag.mode,
      ok: true,
      sourceId: drag.sourceId,
      targetId: drag.targetId,
    }
  })()`)
}

async function dragPPTLayerPaneGroupRow(page) {
  const started = await page.eval(`(() => {
    const group = document.querySelector('[data-ppt-layer-pane-row-type="group"]')

    if (!(group instanceof HTMLElement)) {
      return {
        groupFound: false,
        ok: false,
      }
    }

    const groupRowId = group.getAttribute('data-ppt-layer-pane-row') ?? ''
    const rows = [...document.querySelectorAll('[data-ppt-layer-pane-row]')]
    const groupIndex = rows.indexOf(group)
    const targets = rows.filter((row) =>
      row instanceof HTMLElement &&
        row.getAttribute('data-ppt-layer-pane-draggable') === 'true' &&
        row.getAttribute('data-ppt-layer-pane-row') !== groupRowId &&
        row.getAttribute('data-ppt-layer-pane-parent-object-id') !== groupRowId)
    const afterTarget = targets.find((row) => rows.indexOf(row) > groupIndex)
    const beforeTarget = [...targets].reverse().find((row) => rows.indexOf(row) < groupIndex)
    const target = afterTarget ?? beforeTarget
    const placement = afterTarget ? 'after' : 'before'

    if (!(target instanceof HTMLElement)) {
      return {
        groupFound: true,
        ok: false,
        targetFound: false,
        targetCount: targets.length,
      }
    }

    const rect = target.getBoundingClientRect()
    const clientX = rect.left + rect.width / 2
    const clientY = placement === 'after' ? rect.bottom - 2 : rect.top + 2
    const dataTransfer = typeof DataTransfer === 'function'
      ? new DataTransfer()
      : {
          data: new Map(),
          dropEffect: 'move',
          effectAllowed: 'move',
          getData(type) {
            return this.data.get(type) ?? ''
          },
          setData(type, value) {
            this.data.set(type, value)
          },
        }

    function createDragEvent(type) {
      let event

      try {
        event = new DragEvent(type, {
          bubbles: true,
          cancelable: true,
          clientX,
          clientY,
          dataTransfer,
        })
      } catch {
        event = new Event(type, {
          bubbles: true,
          cancelable: true,
        })
      }

      if (!event.dataTransfer) {
        Object.defineProperty(event, 'dataTransfer', {
          configurable: true,
          value: dataTransfer,
        })
      }
      if (event.clientX !== clientX) {
        Object.defineProperty(event, 'clientX', {
          configurable: true,
          value: clientX,
        })
      }
      if (event.clientY !== clientY) {
        Object.defineProperty(event, 'clientY', {
          configurable: true,
          value: clientY,
        })
      }

      return event
    }

    const sourceId = groupRowId
    const targetId = target.getAttribute('data-ppt-layer-pane-row') ?? ''

    group.dispatchEvent(createDragEvent('dragstart'))
    target.dispatchEvent(createDragEvent('dragover'))

    window.__pptLayerPaneGroupDrag = {
      clientX,
      clientY,
      dataTransfer,
      placement,
      sourceId,
      targetId,
    }

    return {
      ok: true,
      pending: true,
      placement,
      sourceId,
      targetId,
    }
  })()`)

  if (!started.ok) {
    return started
  }

  await delay(30)

  return page.eval(`(() => {
    const drag = window.__pptLayerPaneGroupDrag

    if (!drag) {
      return {
        ok: false,
        pendingFound: false,
      }
    }

    const group = document.querySelector(\`[data-ppt-layer-pane-row="\${drag.sourceId}"]\`)
    const target = document.querySelector(\`[data-ppt-layer-pane-row="\${drag.targetId}"]\`)

    if (!(group instanceof HTMLElement) || !(target instanceof HTMLElement)) {
      delete window.__pptLayerPaneGroupDrag
      return {
        groupFound: group instanceof HTMLElement,
        ok: false,
        targetFound: target instanceof HTMLElement,
      }
    }

    function createDragEvent(type) {
      let event

      try {
        event = new DragEvent(type, {
          bubbles: true,
          cancelable: true,
          clientX: drag.clientX,
          clientY: drag.clientY,
          dataTransfer: drag.dataTransfer,
        })
      } catch {
        event = new Event(type, {
          bubbles: true,
          cancelable: true,
        })
      }

      if (!event.dataTransfer) {
        Object.defineProperty(event, 'dataTransfer', {
          configurable: true,
          value: drag.dataTransfer,
        })
      }
      if (event.clientX !== drag.clientX) {
        Object.defineProperty(event, 'clientX', {
          configurable: true,
          value: drag.clientX,
        })
      }
      if (event.clientY !== drag.clientY) {
        Object.defineProperty(event, 'clientY', {
          configurable: true,
          value: drag.clientY,
        })
      }

      return event
    }

    const indicatorModel = target.getAttribute('data-ppt-layer-pane-drop-indicator-model') ?? ''
    const indicatorPlacement = target.getAttribute('data-ppt-layer-pane-drop-indicator') ?? ''
    const indicatorTarget = target.getAttribute('data-ppt-layer-pane-drop-target') ?? ''
    const indicatorToIndex = target.getAttribute('data-ppt-layer-pane-drop-to-index') ?? ''
    target.dispatchEvent(createDragEvent('drop'))
    group.dispatchEvent(createDragEvent('dragend'))
    delete window.__pptLayerPaneGroupDrag

    return {
      indicatorModel,
      indicatorPlacement,
      indicatorTarget,
      indicatorToIndex,
      ok: true,
      placement: drag.placement,
      sourceId: drag.sourceId,
      targetId: drag.targetId,
    }
  })()`)
}

async function reorderPPTLayerPaneGroupWithKeyboard(page) {
  const before = await readPPTLayerPaneGroupReorderState(page)
  const target = await page.eval(`(() => {
    const group = document.querySelector('[data-ppt-layer-pane-row-type="group"]')

    if (!(group instanceof HTMLElement)) {
      return {
        groupFound: false,
        ok: false,
      }
    }

    group.focus()

    const groupRowId = group.getAttribute('data-ppt-layer-pane-row') ?? ''
    const rows = [...document.querySelectorAll('[data-ppt-layer-pane-row]')]
    const groupIndex = rows.indexOf(group)
    const hasTargetAfter = rows.some((row, index) =>
      index > groupIndex &&
        row.getAttribute('data-ppt-layer-pane-draggable') === 'true' &&
        row.getAttribute('data-ppt-layer-pane-parent-object-id') !== groupRowId)
    const hasTargetBefore = rows.some((row, index) =>
      index < groupIndex &&
        row.getAttribute('data-ppt-layer-pane-draggable') === 'true' &&
        row.getAttribute('data-ppt-layer-pane-parent-object-id') !== groupRowId)

    if (!hasTargetAfter && !hasTargetBefore) {
      return {
        groupRowId,
        ok: false,
        targetFound: false,
      }
    }

    return {
      direction: hasTargetAfter ? 'down' : 'up',
      groupRowId,
      ok: true,
    }
  })()`)

  if (!target.ok) {
    return {
      ...target,
      before,
    }
  }

  await delay(50)
  await pressKey(page, {
    code: target.direction === 'down' ? 'ArrowDown' : 'ArrowUp',
    key: target.direction === 'down' ? 'ArrowDown' : 'ArrowUp',
    modifiers: 1,
    windowsVirtualKeyCode: target.direction === 'down' ? 40 : 38,
  })
  await delay(100)
  const after = await readPPTLayerPaneGroupReorderState(page)
  const focusedRowId = await page.eval(`(() =>
    document.activeElement?.closest('[data-ppt-layer-pane-row]')?.getAttribute('data-ppt-layer-pane-row') ?? ''
  )()`)

  return {
    ...target,
    after,
    before,
    focusedRowId,
    groupRowId: target.groupRowId,
  }
}

async function readPPTLayerPaneGroupReorderState(page) {
  return page.eval(`(() => {
    const group = document.querySelector('[data-ppt-layer-pane-row-type="group"]')
    const groupRowId = group?.getAttribute('data-ppt-layer-pane-row') ?? ''
    const childRows = [...document.querySelectorAll(\`[data-ppt-layer-pane-parent-object-id="\${groupRowId}"][data-ppt-layer-pane-row-type="object"]\`)]
    const childRowIds = childRows.map((row) => row.getAttribute('data-ppt-layer-pane-row') ?? '')
    const stageOrder = [...document.querySelectorAll('[data-ppt-element]')]
      .map((element) => element.getAttribute('data-ppt-element') ?? '')
    const childStageIds = stageOrder.filter((id) => childRowIds.includes(id))
    const childStageIndexes = childStageIds.map((id) => stageOrder.indexOf(id))

    return {
      childRowIds,
      childStageIds,
      childStageIndexes,
      groupDraggable: group?.getAttribute('data-ppt-layer-pane-draggable') ?? '',
      groupMemberContiguous: childStageIndexes.every((index, offset) =>
        offset === 0 || index === childStageIndexes[offset - 1] + 1),
      groupRowId,
      groupStartIndex: childStageIndexes[0] ?? -1,
      selectedStageIds: [...document.querySelectorAll('[data-selected="true"]')]
        .map((element) => element.getAttribute('data-ppt-element') ?? ''),
      stageOrder,
    }
  })()`)
}

async function readPPTLayerPaneGroupTreeState(page) {
  return page.eval(`(() => {
    const tree = document.querySelector('.ppt-layer-list')
    const groupRow = document.querySelector('[data-ppt-layer-pane-row-type="group"]')
    const groupRowId = groupRow?.getAttribute('data-ppt-layer-pane-row') ?? ''
    const childRows = [...document.querySelectorAll(\`[data-ppt-layer-pane-parent-object-id="\${groupRowId}"]\`)]
    const focusedRow = document.activeElement?.closest('[data-ppt-layer-pane-row]')
    const focusedRowId = focusedRow?.getAttribute('data-ppt-layer-pane-row') ?? ''

    return {
      childParentIds: childRows.map((row) => row.getAttribute('data-ppt-layer-pane-parent-object-id') ?? ''),
      childRowCount: childRows.length,
      childRowIds: childRows.map((row) => row.getAttribute('data-ppt-layer-pane-row') ?? ''),
      focusedParentId: focusedRow?.getAttribute('data-ppt-layer-pane-parent-object-id') ?? '',
      focusedRowId,
      groupExpanded: groupRow?.getAttribute('aria-expanded') ?? '',
      groupRowCount: document.querySelectorAll('[data-ppt-layer-pane-row-type="group"]').length,
      groupRowId,
      keyboardIntentModel: tree?.getAttribute('data-ppt-layer-pane-keyboard-intent-model') ?? '',
      selectedRowIds: [...document.querySelectorAll('[data-ppt-layer-pane-row][aria-selected="true"]')]
        .map((row) => row.getAttribute('data-ppt-layer-pane-row') ?? ''),
      selectedStageCount: document.querySelectorAll('[data-selected="true"]').length,
      tabStopIds: [...document.querySelectorAll('[data-ppt-layer-pane-row]')]
        .filter((row) => row.tabIndex === 0)
        .map((row) => row.getAttribute('data-ppt-layer-pane-row') ?? ''),
    }
  })()`)
}

async function runTextOverflowScenario(page) {
  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)
  await page.eval(`document.querySelector('.ppt-thumb[aria-label="Open Overview"]')?.click()`)
  await delay(80)

  const before = await getPPTTextOverflowState(page)

  await page.eval(`document.querySelector('[data-ppt-insert-tool="text"]')?.click()`)
  await delay(40)

  const point = await page.eval(`(() => {
    const slide = document.querySelector('.ppt-slide').getBoundingClientRect()

    return {
      x: slide.left + slide.width * 0.12,
      y: slide.top + slide.height * 0.13,
    }
  })()`)

  await clickMouse(page, point.x, point.y, 1)
  await delay(120)
  await page.eval(`document.activeElement?.blur()`)
  await delay(80)

  const afterCreate = await getPPTTextOverflowState(page)

  record('creates selected PPT text box for overflow retouch', afterCreate.textCount === before.textCount + 1 && afterCreate.selectedKind === 'textBox' && afterCreate.selectedText.includes('New text'), {
    afterCreate,
    before,
  })

  await page.eval(`(() => {
    const text = document.querySelector('[data-ppt-style-field="text"]')
    const width = document.querySelector('[data-ppt-geometry-field="w"]')
    const height = document.querySelector('[data-ppt-geometry-field="h"]')
    const textSetter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set
    const inputSetter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set

    textSetter.call(text, 'Overflowing AI generated takeaway that needs quick auto fit before PPTX compatible export.')
    text.dispatchEvent(new Event('input', { bubbles: true }))
    text.dispatchEvent(new Event('change', { bubbles: true }))

    inputSetter.call(width, '132')
    width.dispatchEvent(new Event('input', { bubbles: true }))
    width.dispatchEvent(new Event('change', { bubbles: true }))

    inputSetter.call(height, '34')
    height.dispatchEvent(new Event('input', { bubbles: true }))
    height.dispatchEvent(new Event('change', { bubbles: true }))
  })()`)
  await delay(220)

  const afterOverflow = await getPPTTextOverflowState(page)

  record('detects selected PPT text overflow on stage, capsule, and inspector', afterOverflow.selectedKind === 'textBox' && afterOverflow.selectedOverflow === 'true' && afterOverflow.capsuleOverflow === 'true' && afterOverflow.capsuleText.includes('Overflow') && afterOverflow.inspectorOverflow === 'true' && !afterOverflow.autoFitDisabled, {
    afterCreate,
    afterOverflow,
  })
  record('exposes PPT text overflow through slide-edit auto-fit indicator metadata', afterOverflow.stageAutoFitModel === 'slide-edit-text-box-auto-fit' && afterOverflow.stageAutoFitSizeModes === 'fixed resize-to-fit shrink-text' && afterOverflow.selectedAutoFitModel === 'slide-edit-text-box-auto-fit' && afterOverflow.selectedAutoFitSizeMode === 'fixed' && afterOverflow.selectedOverflowIndicatorModel === 'slide-edit-text-box-auto-fit' && afterOverflow.selectedOverflowIndicatorSlide === 'slide-1' && afterOverflow.selectedOverflowIndicatorVisible === 'true' && afterOverflow.stageOverflowIndicatorModel === 'slide-edit-text-box-auto-fit' && afterOverflow.stageOverflowIndicatorObject === afterOverflow.selectedId && afterOverflow.stageOverflowIndicatorSizeMode === 'fixed' && afterOverflow.stageOverflowIndicatorSlide === 'slide-1' && afterOverflow.stageOverflowIndicatorVisible === 'true' && afterOverflow.capsuleSizeMode === 'fixed' && afterOverflow.capsuleOverflowIndicatorVisible === 'true' && afterOverflow.inspectorAutoFitModel === 'slide-edit-text-box-auto-fit' && afterOverflow.inspectorAutoFitSizeMode === 'fixed' && afterOverflow.inspectorOverflowIndicatorVisible === 'true' && afterOverflow.selectedOverflowAxis.includes('vertical') && afterOverflow.stageOverflowAxis.includes('vertical') && afterOverflow.capsuleOverflowAxis.includes('vertical') && afterOverflow.inspectorOverflowAxis.includes('vertical'), afterOverflow)

  await page.eval(`document.querySelector('[data-ppt-style-action="text-auto-fit"]')?.click()`)
  await delay(220)

  const afterAutoFit = await getPPTTextOverflowState(page)

  record('auto-fits overflowing PPT text from inspector', afterAutoFit.selectedKind === 'textBox' && afterAutoFit.selectedOverflow !== 'true' && afterAutoFit.selectedAutoFit === 'resizeShapeToFitText' && afterAutoFit.inspectorAutoFit === 'resizeShapeToFitText' && (afterAutoFit.selectedWidth > afterOverflow.selectedWidth || afterAutoFit.selectedHeight > afterOverflow.selectedHeight), {
    afterAutoFit,
    afterOverflow,
  })
  record('routes PPT inspector text auto-fit through slide-edit command effect', afterAutoFit.stageAutoFitCommand === 'resize-text-box-to-fit' && afterAutoFit.stageAutoFitCommandHandle === 'se' && afterAutoFit.stageAutoFitCommandObject === afterAutoFit.selectedId && afterAutoFit.stageAutoFitCommandSelection === afterAutoFit.selectedId && afterAutoFit.stageAutoFitCommandSizeMode === 'resize-to-fit' && afterAutoFit.stageAutoFitCommandSlide === 'slide-1' && afterAutoFit.stageAutoFitCommandType === 'slide-command-effect' && afterAutoFit.inspectorAutoFitCommand === 'resize-text-box-to-fit' && afterAutoFit.inspectorAutoFitCommandHandle === 'se' && afterAutoFit.inspectorAutoFitCommandObject === afterAutoFit.selectedId && afterAutoFit.inspectorAutoFitCommandSizeMode === 'resize-to-fit' && afterAutoFit.inspectorAutoFitCommandType === 'slide-command-effect' && afterAutoFit.exportHasAutoFit, {
    afterAutoFit,
    afterOverflow,
  })

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(160)

  const afterUndo = await getPPTTextOverflowState(page)

  record('undoes PPT text auto-fit as one history step', afterUndo.selectedId === afterAutoFit.selectedId && afterUndo.selectedOverflow === 'true' && afterUndo.selectedAutoFit === '' && afterUndo.selectedWidth === afterOverflow.selectedWidth && afterUndo.selectedHeight === afterOverflow.selectedHeight, {
    afterAutoFit,
    afterOverflow,
    afterUndo,
  })

  await page.eval(`(() => {
    const dataTransfer = new DataTransfer()
    const json = JSON.stringify({
      textAutoFit: {
        handle: 'se',
        mode: 'resize-to-fit',
      },
    })

    dataTransfer.setData('application/json', json)
    dataTransfer.setData('text/plain', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })()`)
  await delay(220)

  const afterJSONAutoFit = await getPPTTextOverflowState(page)

  record('pastes PPT text auto-fit JSON through slide-edit command effect', afterJSONAutoFit.selectedId === afterUndo.selectedId && afterJSONAutoFit.selectedOverflow !== 'true' && afterJSONAutoFit.selectedAutoFit === 'resizeShapeToFitText' && afterJSONAutoFit.stageAutoFitImportModel === 'ppt-text-autofit-import' && afterJSONAutoFit.stageAutoFitImportFormat === 'application-json-ppt-text-autofit' && afterJSONAutoFit.stageAutoFitImportFields === 'mode handle' && afterJSONAutoFit.stageAutoFitImportCommands === 'resize-text-box-to-fit' && afterJSONAutoFit.stageAutoFitImportCommandHandles === 'se' && afterJSONAutoFit.stageAutoFitImportCommandTargets === afterJSONAutoFit.selectedId && afterJSONAutoFit.stageAutoFitImportCommandTypes === 'slide-command-effect' && afterJSONAutoFit.stageAutoFitImportMode === 'resize-to-fit' && afterJSONAutoFit.stageAutoFitImportObjects === afterJSONAutoFit.selectedId && afterJSONAutoFit.stageAutoFitImportJsonLength > 40 && (afterJSONAutoFit.selectedWidth > afterUndo.selectedWidth || afterJSONAutoFit.selectedHeight > afterUndo.selectedHeight), {
    afterJSONAutoFit,
    afterUndo,
  })

  await pressKey(page, {
    code: 'KeyZ',
    key: 'z',
    modifiers: 2,
    windowsVirtualKeyCode: 90,
  })
  await delay(160)

  const afterJSONUndo = await getPPTTextOverflowState(page)

  record('undoes pasted PPT text auto-fit JSON as one history step', afterJSONUndo.selectedId === afterJSONAutoFit.selectedId && afterJSONUndo.selectedOverflow === 'true' && afterJSONUndo.selectedAutoFit === '' && afterJSONUndo.selectedWidth === afterUndo.selectedWidth && afterJSONUndo.selectedHeight === afterUndo.selectedHeight, {
    afterJSONAutoFit,
    afterJSONUndo,
    afterUndo,
  })

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(160)

  const afterRedo = await getPPTTextOverflowState(page)

  record('redoes PPT text auto-fit with model metadata', afterRedo.selectedId === afterJSONAutoFit.selectedId && afterRedo.selectedOverflow !== 'true' && afterRedo.selectedAutoFit === 'resizeShapeToFitText' && afterRedo.inspectorAutoFit === 'resizeShapeToFitText' && afterRedo.selectedAutoFitSizeMode === 'resize-to-fit' && afterRedo.inspectorAutoFitSizeMode === 'resize-to-fit' && afterRedo.selectedWidth === afterJSONAutoFit.selectedWidth && afterRedo.selectedHeight === afterJSONAutoFit.selectedHeight && afterRedo.exportHasAutoFit, {
    afterJSONAutoFit,
    afterRedo,
  })
}

async function runPresentationScenario(page) {
  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)
  await page.eval(`document.querySelector('.ppt-thumb[aria-label="Open Overview"]')?.click()`)
  await delay(80)

  const before = await getPPTPresentationState(page)

  await page.eval(`document.querySelector('[data-ppt-present-start]')?.click()`)
  await delay(120)

  const afterToolbarStart = await getPPTPresentationState(page)

  record('starts PPT presentation preview from current slide toolbar', afterToolbarStart.open && afterToolbarStart.slideId === 'slide-1' && afterToolbarStart.index === '1/2' && afterToolbarStart.title === 'Overview' && afterToolbarStart.activeSlide === before.activeSlide && afterToolbarStart.selectedIds === before.selectedIds && afterToolbarStart.visibleElementCount > 0, {
    afterToolbarStart,
    before,
  })

  record('applies PPT slide transition metadata in presentation preview', afterToolbarStart.transitionType === 'push' && afterToolbarStart.transitionDuration === '650' && afterToolbarStart.advanceOnClick === 'false' && afterToolbarStart.advanceAfter === '3000', {
    afterToolbarStart,
  })

  await page.eval(`document.querySelector('[data-ppt-present-next]')?.click()`)
  await delay(80)

  const afterNext = await getPPTPresentationState(page)

  record('moves PPT presentation preview to next slide button', afterNext.open && afterNext.slideId === 'slide-2' && afterNext.index === '2/2' && afterNext.activeSlide === before.activeSlide && afterNext.selectedIds === before.selectedIds && afterNext.transitionType === 'none', {
    afterNext,
    before,
  })

  await pressKey(page, {
    code: 'ArrowRight',
    key: 'ArrowRight',
    windowsVirtualKeyCode: 39,
  })
  await delay(80)

  const afterArrowRight = await getPPTPresentationState(page)

  await pressKey(page, {
    code: 'PageUp',
    key: 'PageUp',
    windowsVirtualKeyCode: 33,
  })
  await delay(80)

  const afterPageUp = await getPPTPresentationState(page)

  record('cycles PPT presentation preview with keyboard navigation', afterArrowRight.open && afterArrowRight.slideId === 'slide-1' && afterArrowRight.index === '1/2' && afterPageUp.open && afterPageUp.slideId === 'slide-2' && afterPageUp.index === '2/2', {
    afterArrowRight,
    afterNext,
    afterPageUp,
  })

  const previewPoint = await page.eval(`(() => {
    const frame = document.querySelector('[data-ppt-presentation-slide-frame]')?.getBoundingClientRect()

    return frame
      ? {
          x: frame.left + frame.width / 2,
          y: frame.top + frame.height / 2,
        }
      : {
          x: 0,
          y: 0,
        }
  })()`)
  await clickMouse(page, previewPoint.x, previewPoint.y, 1)
  await delay(80)

  const afterPreviewClick = await getPPTPresentationState(page)

  record('keeps PPT presentation preview read-only', afterPreviewClick.open && afterPreviewClick.activeSlide === before.activeSlide && afterPreviewClick.selectedIds === before.selectedIds, {
    afterPreviewClick,
    before,
  })

  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(100)

  const afterEscape = await getPPTPresentationState(page)

  record('exits PPT presentation preview with Escape', !afterEscape.open && afterEscape.activeSlide === before.activeSlide && afterEscape.selectedIds === before.selectedIds, {
    afterEscape,
    before,
  })

  await pressKey(page, {
    code: 'KeyK',
    key: 'k',
    modifiers: 2,
    windowsVirtualKeyCode: 75,
  })
  await delay(80)
  await page.send('Input.insertText', { text: 'present' })
  await delay(80)

  const beforePaletteStart = await getPPTPresentationState(page)

  await pressKey(page, {
    code: 'Enter',
    key: 'Enter',
    windowsVirtualKeyCode: 13,
  })
  await delay(120)

  const afterPaletteStart = await getPPTPresentationState(page)

  record('starts PPT presentation preview from command palette', beforePaletteStart.paletteOpen && beforePaletteStart.palettePresent && afterPaletteStart.open && !afterPaletteStart.paletteOpen && afterPaletteStart.slideId === 'slide-1' && afterPaletteStart.activeSlide === before.activeSlide && afterPaletteStart.selectedIds === before.selectedIds, {
    afterPaletteStart,
    before,
    beforePaletteStart,
  })

  await page.eval(`document.querySelector('[data-ppt-present-exit]')?.click()`)
  await delay(100)

  const afterExitButton = await getPPTPresentationState(page)

  record('exits PPT presentation preview with explicit control', !afterExitButton.open && afterExitButton.activeSlide === before.activeSlide && afterExitButton.selectedIds === before.selectedIds, {
    afterExitButton,
    before,
  })
}

async function getPPTPresentationState(page) {
  return page.eval(`(() => {
    const overlay = document.querySelector('[data-ppt-presentation]')

    return {
      activeSlide: document.querySelector('.ppt-thumb[aria-current="page"] .ppt-thumb-name')?.textContent ?? '',
      advanceAfter: overlay?.getAttribute('data-ppt-presentation-advance-after') ?? '',
      advanceOnClick: overlay?.getAttribute('data-ppt-presentation-advance-on-click') ?? '',
      index: overlay?.getAttribute('data-ppt-presentation-index') ?? '',
      open: !!overlay,
      paletteOpen: !!document.querySelector('[data-ppt-command-palette]'),
      palettePresent: !!document.querySelector('[data-ppt-command-palette-item="view:present"]'),
      selectedIds: [...document.querySelectorAll('[data-selected="true"]')]
        .map((element) => element.getAttribute('data-ppt-element'))
        .filter(Boolean)
        .join(','),
      slideId: overlay?.getAttribute('data-ppt-presentation-slide') ?? '',
      title: overlay?.querySelector('[data-ppt-presentation-title]')?.textContent ?? '',
      transitionDuration: overlay?.getAttribute('data-ppt-presentation-transition-duration') ?? '',
      transitionType: overlay?.getAttribute('data-ppt-presentation-transition') ?? '',
      visibleElementCount: overlay?.querySelectorAll('[data-ppt-element]').length ?? 0,
    }
  })()`)
}

async function runSlideManagementScenario(page) {
  const before = await getSlideRailState(page)

  await page.eval(`document.querySelector('[data-ppt-slide-action="duplicate"]').click()`)
  await delay(50)

  const afterDuplicate = await getSlideRailState(page)

  record('duplicates active PPT slide in rail', afterDuplicate.count === before.count + 1 && afterDuplicate.activeName.includes('Copy'), {
    afterDuplicate,
    before,
  })

  await page.eval(`(() => {
    window.__pptSlideClipboardItemTypes = []
    window.__pptSlideClipboardWriteCount = 0
    window.__pptSlideClipboardHTML = ''
    window.__pptSlideClipboardJSON = ''
    window.__pptSlideClipboardPlainText = ''
    window.__pptSlideClipboardSVG = ''
    window.ClipboardItem = class PPTSlideClipboardItem {
      constructor(items) {
        this.items = items
        window.__pptSlideClipboardItemTypes.push(Object.keys(items).sort())
      }
    }
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: {
        async write(items) {
          window.__pptSlideClipboardWriteCount = items.length
          const item = items[0]
          const jsonMimeType = Object.keys(item.items)
            .find((type) => type.includes('ppt.slide+json')) ?? ''

          window.__pptSlideClipboardHTML = await item.items['text/html'].text()
          window.__pptSlideClipboardPlainText = await item.items['text/plain'].text()
          window.__pptSlideClipboardSVG = await item.items['image/svg+xml'].text()
          window.__pptSlideClipboardJSON = jsonMimeType
            ? await item.items[jsonMimeType].text()
            : ''
        },
        async writeText(text) {
          window.__pptSlideClipboardPlainText = text
        },
      },
    })
  })()`)

  await page.eval(`document.querySelector('[data-ppt-slide-action="copy"]').click()`)
  await delay(80)

  const afterCopySlide = await getSlideRailState(page)
  const slideClipboardWrite = await page.eval(`(() => ({
    html: window.__pptSlideClipboardHTML ?? '',
    itemTypes: window.__pptSlideClipboardItemTypes?.at(-1) ?? [],
    json: window.__pptSlideClipboardJSON ?? '',
    plainText: window.__pptSlideClipboardPlainText ?? '',
    svg: window.__pptSlideClipboardSVG ?? '',
    writeCount: window.__pptSlideClipboardWriteCount ?? 0,
  }))()`)

  record(
    'copies active PPT slide as rich slide clipboard bundle',
    afterCopySlide.slideClipboardModel === 'canvas-board-io-ppt-slide-clipboard' &&
      afterCopySlide.slideClipboardJsonMimeType === 'application/vnd.interactive-os.ppt.slide+json' &&
      afterCopySlide.slideClipboardSourceSlide === afterDuplicate.activeId &&
      afterCopySlide.slideClipboardElementCount > 0 &&
      afterCopySlide.slideClipboardHTMLLength > 0 &&
      afterCopySlide.slideClipboardWriteMode === 'clipboard-item' &&
      slideClipboardWrite.writeCount === 1 &&
      slideClipboardWrite.itemTypes.includes('application/vnd.interactive-os.ppt.slide+json') &&
      slideClipboardWrite.itemTypes.includes('text/html') &&
      slideClipboardWrite.itemTypes.includes('text/plain') &&
      slideClipboardWrite.itemTypes.includes('image/svg+xml') &&
      slideClipboardWrite.html.includes('data-ppt-slide-clipboard-json') &&
      slideClipboardWrite.json.includes('"kind": "interactive-os.ppt.slide"') &&
      slideClipboardWrite.plainText.includes('Overview') &&
      slideClipboardWrite.svg.includes('<svg'),
    {
      afterCopySlide,
      afterDuplicate,
      slideClipboardWrite,
    },
  )

  await page.eval(`document.querySelector('[data-ppt-slide-action="paste"]').click()`)
  await delay(80)

  const afterPasteSlideButton = await getSlideRailState(page)

  record(
    'pastes copied PPT slide from slide rail action',
    afterPasteSlideButton.count === afterDuplicate.count + 1 &&
      afterPasteSlideButton.activeId !== afterDuplicate.activeId &&
      afterPasteSlideButton.activeName.includes('Copy') &&
      afterPasteSlideButton.slideClipboardModel === 'canvas-board-io-ppt-slide-clipboard' &&
      afterPasteSlideButton.slideClipboardSourceSlide === afterDuplicate.activeId &&
      afterPasteSlideButton.slideClipboardTargetSlide === afterPasteSlideButton.activeId &&
      afterPasteSlideButton.slideClipboardImported !== 'true',
    {
      afterDuplicate,
      afterPasteSlideButton,
    },
  )

  await page.eval(`document.querySelector('[data-ppt-slide-action="delete"]')?.click()`)
  await delay(60)
  await page.eval(`((slideId) => {
    const thumb = [...document.querySelectorAll('.ppt-thumb')]
      .find((item) => item.getAttribute('data-ppt-slide-id') === slideId)
    thumb?.click()
  })(${JSON.stringify(afterDuplicate.activeId)})`)
  await delay(60)

  await page.eval(`((json) => {
    const dataTransfer = new DataTransfer()
    dataTransfer.setData('application/vnd.interactive-os.ppt.slide+json', json)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })(${JSON.stringify(slideClipboardWrite.json)})`)
  await delay(80)

  const afterPasteSlideCustomJSON = await getSlideRailState(page)

  record(
    'pastes external PPT slide clipboard custom JSON as slide',
    afterPasteSlideCustomJSON.count === afterDuplicate.count + 1 &&
      afterPasteSlideCustomJSON.activeId !== afterDuplicate.activeId &&
      afterPasteSlideCustomJSON.slideClipboardImported === 'true' &&
      afterPasteSlideCustomJSON.slideClipboardImportFormat === 'custom-json' &&
      afterPasteSlideCustomJSON.slideClipboardSourceSlide === afterDuplicate.activeId &&
      afterPasteSlideCustomJSON.slideClipboardTargetSlide === afterPasteSlideCustomJSON.activeId,
    {
      afterDuplicate,
      afterPasteSlideCustomJSON,
    },
  )

  await page.eval(`document.querySelector('[data-ppt-slide-action="delete"]')?.click()`)
  await delay(60)
  await page.eval(`((slideId) => {
    const thumb = [...document.querySelectorAll('.ppt-thumb')]
      .find((item) => item.getAttribute('data-ppt-slide-id') === slideId)
    thumb?.click()
  })(${JSON.stringify(afterDuplicate.activeId)})`)
  await delay(60)

  await page.eval(`((html, plainText) => {
    const dataTransfer = new DataTransfer()
    const fallbackHTML = html.replace(/<script\\b[\\s\\S]*?<\\/script>/gi, '')

    dataTransfer.setData('text/html', fallbackHTML)
    dataTransfer.setData('text/plain', plainText)
    window.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer,
    }))
  })(${JSON.stringify(slideClipboardWrite.html)}, ${JSON.stringify(slideClipboardWrite.plainText)})`)
  await delay(80)

  const afterPasteSlideFallbackHTML = await getSlideRailState(page)

  record(
    'pastes PPT slide fallback HTML as snapshot slide when JSON is stripped',
    afterPasteSlideFallbackHTML.count === afterDuplicate.count + 1 &&
      afterPasteSlideFallbackHTML.activeId !== afterDuplicate.activeId &&
      afterPasteSlideFallbackHTML.activeName.includes('Snapshot') &&
      afterPasteSlideFallbackHTML.slideClipboardImported === 'true' &&
      afterPasteSlideFallbackHTML.slideClipboardImportFormat === 'text-html-fallback' &&
      afterPasteSlideFallbackHTML.slideClipboardElementCount === 1 &&
      afterPasteSlideFallbackHTML.slideClipboardSourceSlide === afterDuplicate.activeId &&
      afterPasteSlideFallbackHTML.slideClipboardTargetSlide === afterPasteSlideFallbackHTML.activeId,
    {
      afterDuplicate,
      afterPasteSlideFallbackHTML,
    },
  )

  await page.eval(`document.querySelector('[data-ppt-slide-action="delete"]')?.click()`)
  await delay(60)
  await page.eval(`((slideId) => {
    const thumb = [...document.querySelectorAll('.ppt-thumb')]
      .find((item) => item.getAttribute('data-ppt-slide-id') === slideId)
    thumb?.click()
  })(${JSON.stringify(afterDuplicate.activeId)})`)
  await delay(60)

  const afterSlideClipboardCleanup = await getSlideRailState(page)

  record(
    'restores PPT slide rail after slide clipboard paste probes',
    afterSlideClipboardCleanup.count === afterDuplicate.count &&
      afterSlideClipboardCleanup.activeId === afterDuplicate.activeId,
    {
      afterDuplicate,
      afterSlideClipboardCleanup,
    },
  )

  await page.eval(`document.querySelector('[data-ppt-slide-action="move-down"]').click()`)
  await delay(50)

  const afterMoveDown = await getSlideRailState(page)

  record('moves active PPT slide down in rail', afterMoveDown.activeIndex === afterDuplicate.activeIndex + 1 && afterMoveDown.activeName.includes('Copy'), {
    afterDuplicate,
    afterMoveDown,
  })

  await pressKey(page, {
    code: 'PageUp',
    key: 'PageUp',
    windowsVirtualKeyCode: 33,
  })
  await delay(50)

  const afterPageUp = await getSlideRailState(page)

  await pressKey(page, {
    code: 'PageDown',
    key: 'PageDown',
    windowsVirtualKeyCode: 34,
  })
  await delay(50)

  const afterPageDown = await getSlideRailState(page)

  record('navigates PPT slides with PageUp and PageDown', afterPageUp.activeIndex === afterMoveDown.activeIndex - 1 && afterPageDown.activeIndex === afterMoveDown.activeIndex, {
    afterMoveDown,
    afterPageDown,
    afterPageUp,
  })

  const selectedElementId = await page.eval(`document.querySelector('.ppt-slide [data-ppt-element]')?.getAttribute('data-ppt-element') ?? ''`)

  if (selectedElementId) {
    const point = await getElementCenter(page, selectedElementId)
    await clickMouse(page, point.x, point.y, 1)
    await delay(50)
  }

  const beforeDrag = await getSlideRailState(page)
  const dragDetails = await dragPPTSlideThumbnail(page)
  await delay(80)

  const afterDrag = await getSlideRailState(page)

  record('drags PPT slide thumbnail to reorder rail', dragDetails.ok && afterDrag.model === 'slide-edit-rail-interactions' && afterDrag.activeIndex === 0 && afterDrag.activeId === beforeDrag.activeId && afterDrag.activeName.startsWith('1. ') && afterDrag.activeName.includes('Copy') && afterDrag.selectedIds === beforeDrag.selectedIds && afterDrag.draggableCount === afterDrag.count && afterDrag.command === 'reorder-slide' && afterDrag.commandFromIndex === String(beforeDrag.activeIndex) && afterDrag.commandToIndex === '0' && afterDrag.commandSlide === beforeDrag.activeId && afterDrag.commandSelectionSlide === beforeDrag.activeId && afterDrag.commandType === 'slide-command-effect', {
    afterDrag,
    beforeDrag,
    dragDetails,
  })

  await page.eval(`document.querySelector('button[title="Undo"]').click()`)
  await delay(80)

  const afterDragUndo = await getSlideRailState(page)

  await page.eval(`document.querySelector('button[title="Redo"]').click()`)
  await delay(80)

  const afterDragRedo = await getSlideRailState(page)

  record('undoes and redoes PPT slide drag reorder as one history step', afterDragUndo.activeId === beforeDrag.activeId && afterDragUndo.activeIndex === beforeDrag.activeIndex && afterDragUndo.selectedIds === beforeDrag.selectedIds && afterDragRedo.activeId === beforeDrag.activeId && afterDragRedo.activeIndex === 0 && afterDragRedo.selectedIds === beforeDrag.selectedIds, {
    afterDragRedo,
    afterDragUndo,
    beforeDrag,
  })

  await page.eval(`document.querySelector('[data-ppt-slide-action="delete"]')?.click()`)
  await delay(50)

  const afterDelete = await getSlideRailState(page)

  record('deletes active PPT slide without removing final slide', afterDelete.count === before.count && afterDelete.count >= 1, {
    afterDelete,
    before,
  })

  await focusPPTSlideThumb(page, afterDelete.activeId)
  await delay(50)

  const initialKeyboard = await getSlideRailState(page)
  const expectedActiveOptionId = `slide-rail-option-${initialKeyboard.activeIndex}`
  const expectedSlideOrder = initialKeyboard.ids.join(' ')

  record(
    'exposes PPT slide rail listbox keyboard affordance',
    initialKeyboard.model === 'slide-edit-rail-interactions' &&
      initialKeyboard.listRole === 'listbox' &&
      initialKeyboard.keyboardModel === 'aria-listbox-roving-focus' &&
      initialKeyboard.keyboardKeys === 'ArrowUp ArrowDown Home End Enter Space' &&
      initialKeyboard.selectionMode === 'single' &&
      initialKeyboard.activeAttr === initialKeyboard.activeId &&
      initialKeyboard.activeOption === expectedActiveOptionId &&
      initialKeyboard.focusableOption === expectedActiveOptionId &&
      initialKeyboard.slideOrder === expectedSlideOrder &&
      initialKeyboard.optionCount === initialKeyboard.count &&
      initialKeyboard.optionCountAttr === String(initialKeyboard.count) &&
      initialKeyboard.thumbnailCount === String(initialKeyboard.count) &&
      initialKeyboard.optionIds.length === initialKeyboard.count &&
      initialKeyboard.optionIds.every((id, index) => id === `slide-rail-option-${index}`) &&
      initialKeyboard.optionIndexes.every((value, index) => value === String(index)) &&
      initialKeyboard.optionFocusableIds.length === 1 &&
      initialKeyboard.optionFocusableIds[0] === initialKeyboard.activeId &&
      initialKeyboard.activeThumbW === '112' &&
      initialKeyboard.activeThumbH === '86' &&
      initialKeyboard.activeHitW === '124' &&
      initialKeyboard.activeHitH === '98' &&
      initialKeyboard.selectedOptionIds.length === 1 &&
      initialKeyboard.selectedOptionIds[0] === initialKeyboard.activeId &&
      initialKeyboard.tabStopIds.length === 1 &&
      initialKeyboard.tabStopIds[0] === initialKeyboard.activeId &&
      initialKeyboard.focusedId === initialKeyboard.activeId,
    {
      initialKeyboard,
    },
  )

  await pressKey(page, {
    code: 'Home',
    key: 'Home',
    windowsVirtualKeyCode: 36,
  })
  await delay(50)

  const afterRailHome = await getSlideRailState(page)

  await pressKey(page, {
    code: 'ArrowDown',
    key: 'ArrowDown',
    windowsVirtualKeyCode: 40,
  })
  await delay(50)

  const afterRailArrowDown = await getSlideRailState(page)

  await pressKey(page, {
    code: 'End',
    key: 'End',
    windowsVirtualKeyCode: 35,
  })
  await delay(50)

  const afterRailEnd = await getSlideRailState(page)

  await pressKey(page, {
    code: 'ArrowUp',
    key: 'ArrowUp',
    windowsVirtualKeyCode: 38,
  })
  await delay(50)

  const afterRailArrowUp = await getSlideRailState(page)

  record('navigates PPT slide rail listbox with Arrow and Home End keys', afterRailHome.activeIndex === 0 && afterRailHome.focusedId === afterRailHome.activeId && afterRailArrowDown.activeIndex === Math.min(1, afterRailArrowDown.count - 1) && afterRailArrowDown.focusedId === afterRailArrowDown.activeId && afterRailEnd.activeIndex === afterRailEnd.count - 1 && afterRailEnd.focusedId === afterRailEnd.activeId && afterRailArrowUp.activeIndex === Math.max(0, afterRailEnd.count - 2) && afterRailArrowUp.focusedId === afterRailArrowUp.activeId, {
    afterRailArrowDown,
    afterRailArrowUp,
    afterRailEnd,
    afterRailHome,
  })

  const spaceTargetId = afterRailArrowUp.ids.find((id) => id !== afterRailArrowUp.activeId) ?? afterRailArrowUp.activeId
  await focusPPTSlideThumb(page, spaceTargetId)
  await pressKey(page, {
    code: 'Space',
    key: ' ',
    windowsVirtualKeyCode: 32,
  })
  await delay(50)

  const afterRailSpace = await getSlideRailState(page)
  const enterTargetId = afterRailSpace.ids.find((id) => id !== afterRailSpace.activeId) ?? afterRailSpace.activeId
  await focusPPTSlideThumb(page, enterTargetId)
  await pressKey(page, {
    code: 'Enter',
    key: 'Enter',
    windowsVirtualKeyCode: 13,
  })
  await delay(50)

  const afterRailEnter = await getSlideRailState(page)

  record('selects focused PPT slide thumbnail with Enter and Space', afterRailSpace.activeId === spaceTargetId && afterRailSpace.command === 'select-active-slide' && afterRailSpace.commandSelectionSlide === spaceTargetId && afterRailSpace.focusedId === spaceTargetId && afterRailEnter.activeId === enterTargetId && afterRailEnter.command === 'select-active-slide' && afterRailEnter.commandSelectionSlide === enterTargetId && afterRailEnter.focusedId === enterTargetId, {
    afterRailEnter,
    afterRailSpace,
    enterTargetId,
    spaceTargetId,
  })
}

function createPPTTestImageFileExpression(name, color) {
  const svg = [
    `<svg xmlns="http://www.w3.org/2000/svg" width="${PPT_TEST_IMAGE_WIDTH}" height="${PPT_TEST_IMAGE_HEIGHT}" viewBox="0 0 ${PPT_TEST_IMAGE_WIDTH} ${PPT_TEST_IMAGE_HEIGHT}">`,
    `<rect width="${PPT_TEST_IMAGE_WIDTH}" height="${PPT_TEST_IMAGE_HEIGHT}" rx="24" fill="${color}"/>`,
    '<circle cx="500" cy="110" r="72" fill="white" fill-opacity="0.45"/>',
    '<path d="M80 270h360" stroke="white" stroke-width="42" stroke-linecap="round" opacity="0.72"/>',
    '</svg>',
  ].join('')

  return `new File([${JSON.stringify(svg)}], ${JSON.stringify(name)}, { type: 'image/svg+xml' })`
}

function createPPTTestTableFileExpression(name, text) {
  const type = name.toLowerCase().endsWith('.tsv')
    ? 'text/tab-separated-values'
    : 'text/csv'

  return `new File([${JSON.stringify(text)}], ${JSON.stringify(name)}, { type: ${JSON.stringify(type)} })`
}

function getPPTImageImportState(page) {
  return page.eval(`(() => {
    const selected = document.querySelector('[data-selected="true"]')
    const selectedElements = [...document.querySelectorAll('[data-selected="true"]')]
    const selectedLayerNames = [...document.querySelectorAll('[data-ppt-layer-row][aria-selected="true"] .ppt-layer-name')]
      .map((element) => element.textContent ?? '')
    const selectedImage = selected?.querySelector('img') ?? null
    const stage = document.querySelector('.ppt-stage-shell')
    const fitField = document.querySelector('[data-ppt-style-field="image-fit"]')
    const cropXField = document.querySelector('[data-ppt-style-field="image-crop-x"]')
    const cropYField = document.querySelector('[data-ppt-style-field="image-crop-y"]')
    const selectedImageSrc = selectedImage?.getAttribute('src') ?? ''
    const selectedImageDecoded = (() => {
      try {
        return decodeURIComponent(selectedImageSrc.split(',')[1] ?? '')
      } catch {
        return ''
      }
    })()

    return {
      imageCropCommand: stage?.getAttribute('data-ppt-image-crop-command') ?? '',
      imageCropCommandCropX: stage?.getAttribute('data-ppt-image-crop-command-crop-x') ?? '',
      imageCropCommandCropY: stage?.getAttribute('data-ppt-image-crop-command-crop-y') ?? '',
      imageCropCommandField: stage?.getAttribute('data-ppt-image-crop-command-field') ?? '',
      imageCropCommandFit: stage?.getAttribute('data-ppt-image-crop-command-fit') ?? '',
      imageCropCommandObject: stage?.getAttribute('data-ppt-image-crop-command-object') ?? '',
      imageCropCommandSlide: stage?.getAttribute('data-ppt-image-crop-command-slide') ?? '',
      imageCropCommandType: stage?.getAttribute('data-ppt-image-crop-command-type') ?? '',
      imageCropCommandValue: stage?.getAttribute('data-ppt-image-crop-command-value') ?? '',
      imageCropImportCommandFields: stage?.getAttribute('data-ppt-image-crop-import-command-fields') ?? '',
      imageCropImportCommandTypes: stage?.getAttribute('data-ppt-image-crop-import-command-types') ?? '',
      imageCropImportCommands: stage?.getAttribute('data-ppt-image-crop-import-commands') ?? '',
      imageCropImportFields: stage?.getAttribute('data-ppt-image-crop-import-fields') ?? '',
      imageCropImportFit: stage?.getAttribute('data-ppt-image-crop-import-fit') ?? '',
      imageCropImportFormat: stage?.getAttribute('data-ppt-image-crop-import-format') ?? '',
      imageCropImportJsonLength: Number(stage?.getAttribute('data-ppt-image-crop-import-json-length') ?? 0),
      imageCropImportModel: stage?.getAttribute('data-ppt-image-crop-import-model') ?? '',
      imageCropImportObjects: stage?.getAttribute('data-ppt-image-crop-import-objects') ?? '',
      imageCropImportSlide: stage?.getAttribute('data-ppt-image-crop-import-slide') ?? '',
      imageCropImportX: stage?.getAttribute('data-ppt-image-crop-import-x') ?? '',
      imageCropImportY: stage?.getAttribute('data-ppt-image-crop-import-y') ?? '',
      fallbackHTMLImportFormat: stage?.getAttribute('data-ppt-fallback-html-import-format') ?? '',
      fallbackHTMLImportKind: stage?.getAttribute('data-ppt-fallback-html-import-kind') ?? '',
      fallbackHTMLImportModel: stage?.getAttribute('data-ppt-fallback-html-import-model') ?? '',
      fallbackHTMLImportSourceObject: stage?.getAttribute('data-ppt-fallback-html-import-source-object') ?? '',
      importExtension: stage?.getAttribute('data-ppt-import-extension') ?? '',
      importExtensionClipboardActionOrder: stage?.getAttribute('data-ppt-import-extension-clipboard-action-order') ?? '',
      importExtensionDropActionOrder: stage?.getAttribute('data-ppt-import-extension-drop-action-order') ?? '',
      importExtensionInstallUnit: stage?.getAttribute('data-ppt-import-extension-install-unit') ?? '',
      imageImportCount: Number(stage?.getAttribute('data-ppt-image-import-count') ?? 0),
      imageImportFormat: stage?.getAttribute('data-ppt-image-import-format') ?? '',
      imageImportMime: stage?.getAttribute('data-ppt-image-import-mime') ?? '',
      imageImportModel: stage?.getAttribute('data-ppt-image-import-model') ?? '',
      imageImportName: stage?.getAttribute('data-ppt-image-import-name') ?? '',
      imageImportNames: stage?.getAttribute('data-ppt-image-import-names') ?? '',
      imageImportNaturalHeight: Number(stage?.getAttribute('data-ppt-image-import-natural-height') ?? 0),
      imageImportNaturalWidth: Number(stage?.getAttribute('data-ppt-image-import-natural-width') ?? 0),
      imageCropFitDescriptorAttribute: fitField?.getAttribute('data-ppt-image-crop-attribute') ?? '',
      imageCropFitDescriptorAttributeValue: fitField?.getAttribute('data-ppt-image-crop-attribute-value') ?? '',
      imageCropFitDescriptorCommand: fitField?.getAttribute('data-ppt-image-crop-command') ?? '',
      imageCropFitDescriptorControl: fitField?.getAttribute('data-ppt-image-crop-control') ?? '',
      imageCropFitDescriptorSurface: fitField?.getAttribute('data-ppt-image-crop-surface') ?? '',
      imageCropModel: stage?.getAttribute('data-ppt-image-crop-model') ?? '',
      imageCropResetDescriptorCommand: document.querySelector('[data-ppt-image-crop-reset]')?.getAttribute('data-ppt-image-crop-command') ?? '',
      imageCropResetDescriptorControl: document.querySelector('[data-ppt-image-crop-reset]')?.getAttribute('data-ppt-image-crop-control') ?? '',
      imageCropResetDescriptorSurface: document.querySelector('[data-ppt-image-crop-reset]')?.getAttribute('data-ppt-image-crop-surface') ?? '',
      imageCropXDescriptorAttributeValue: cropXField?.getAttribute('data-ppt-image-crop-attribute-value') ?? '',
      imageCropXDescriptorCommand: cropXField?.getAttribute('data-ppt-image-crop-command') ?? '',
      imageCropXDescriptorControl: cropXField?.getAttribute('data-ppt-image-crop-control') ?? '',
      imageCropXDescriptorSurface: cropXField?.getAttribute('data-ppt-image-crop-surface') ?? '',
      imageCropYDescriptorAttributeValue: cropYField?.getAttribute('data-ppt-image-crop-attribute-value') ?? '',
      imageCropYDescriptorCommand: cropYField?.getAttribute('data-ppt-image-crop-command') ?? '',
      imageCropYDescriptorControl: cropYField?.getAttribute('data-ppt-image-crop-control') ?? '',
      imageCropYDescriptorSurface: cropYField?.getAttribute('data-ppt-image-crop-surface') ?? '',
      imageReplaceCommand: stage?.getAttribute('data-ppt-image-replace-command') ?? '',
      imageReplaceCommandMime: stage?.getAttribute('data-ppt-image-replace-command-mime') ?? '',
      imageReplaceCommandName: stage?.getAttribute('data-ppt-image-replace-command-name') ?? '',
      imageReplaceCommandObject: stage?.getAttribute('data-ppt-image-replace-command-object') ?? '',
      imageReplaceCommandSlide: stage?.getAttribute('data-ppt-image-replace-command-slide') ?? '',
      imageReplaceCommandSrcPrefix: stage?.getAttribute('data-ppt-image-replace-command-src-prefix') ?? '',
      imageReplaceCommandType: stage?.getAttribute('data-ppt-image-replace-command-type') ?? '',
      imageReplaceImportAltTextLength: Number(stage?.getAttribute('data-ppt-image-replace-import-alt-text-length') ?? 0),
      imageReplaceImportCommand: stage?.getAttribute('data-ppt-image-replace-import-command') ?? '',
      imageReplaceImportCommandType: stage?.getAttribute('data-ppt-image-replace-import-command-type') ?? '',
      imageReplaceImportFields: stage?.getAttribute('data-ppt-image-replace-import-fields') ?? '',
      imageReplaceImportFormat: stage?.getAttribute('data-ppt-image-replace-import-format') ?? '',
      imageReplaceImportJsonLength: Number(stage?.getAttribute('data-ppt-image-replace-import-json-length') ?? 0),
      imageReplaceImportMime: stage?.getAttribute('data-ppt-image-replace-import-mime') ?? '',
      imageReplaceImportModel: stage?.getAttribute('data-ppt-image-replace-import-model') ?? '',
      imageReplaceImportName: stage?.getAttribute('data-ppt-image-replace-import-name') ?? '',
      imageReplaceImportNaturalHeight: stage?.getAttribute('data-ppt-image-replace-import-natural-height') ?? '',
      imageReplaceImportNaturalWidth: stage?.getAttribute('data-ppt-image-replace-import-natural-width') ?? '',
      imageReplaceImportObject: stage?.getAttribute('data-ppt-image-replace-import-object') ?? '',
      imageReplaceImportSlide: stage?.getAttribute('data-ppt-image-replace-import-slide') ?? '',
      imageReplaceImportSrcPrefix: stage?.getAttribute('data-ppt-image-replace-import-src-prefix') ?? '',
      imageReplaceDescriptorAttribute: document.querySelector('[data-ppt-image-replace-input]')?.getAttribute('data-ppt-image-replace-attribute') ?? '',
      imageReplaceDescriptorAttributeValue: document.querySelector('[data-ppt-image-replace-input]')?.getAttribute('data-ppt-image-replace-attribute-value') ?? '',
      imageReplaceDescriptorCommand: document.querySelector('[data-ppt-image-replace-action]')?.getAttribute('data-ppt-image-replace-command') ?? '',
      imageReplaceDescriptorControl: document.querySelector('[data-ppt-image-replace-action]')?.getAttribute('data-ppt-image-replace-control') ?? '',
      imageReplaceDescriptorField: document.querySelector('[data-ppt-image-replace-action]')?.getAttribute('data-ppt-image-replace-field') ?? '',
      imageReplaceDescriptorSourceName: document.querySelector('[data-ppt-image-replace-input]')?.getAttribute('data-ppt-image-replace-source-name') ?? '',
      imageReplaceDescriptorSurface: document.querySelector('[data-ppt-image-replace-action]')?.getAttribute('data-ppt-image-replace-surface') ?? '',
      imageReplaceModel: stage?.getAttribute('data-ppt-image-replace-model') ?? '',
      inspectorCropX: Number(cropXField?.value ?? 0),
      inspectorCropY: Number(cropYField?.value ?? 0),
      inspectorImageFit: fitField?.value ?? '',
      imageCount: document.querySelectorAll('[data-kind="image"]').length,
      richClipboardHTMLLength: Number(stage?.getAttribute('data-ppt-rich-clipboard-html-length') ?? 0),
      richClipboardJsonMimeType: stage?.getAttribute('data-ppt-rich-clipboard-json-mime-type') ?? '',
      richClipboardModel: stage?.getAttribute('data-ppt-rich-clipboard-model') ?? '',
      richClipboardPlainTextLength: Number(stage?.getAttribute('data-ppt-rich-clipboard-plain-text-length') ?? 0),
      richClipboardSelection: stage?.getAttribute('data-ppt-rich-clipboard-selection') ?? '',
      richClipboardWriteMode: stage?.getAttribute('data-ppt-rich-clipboard-write-mode') ?? '',
      selectedAltText: selectedImage?.getAttribute('alt') ?? '',
      selectedCount: selectedElements.length,
      selectedFlipH: selected?.getAttribute('data-ppt-flip-h') ?? '',
      selectedFlipV: selected?.getAttribute('data-ppt-flip-v') ?? '',
      selectedId: selected?.getAttribute('data-ppt-element') ?? '',
      selectedImageDecoded,
      selectedImageFit: selectedImage?.style.objectFit ?? '',
      selectedImagePosition: selectedImage?.style.objectPosition ?? '',
      selectedImageSrc,
      selectedKind: selected?.getAttribute('data-kind') ?? '',
      selectedKinds: selectedElements.map((element) => element.getAttribute('data-kind') ?? '').join(' '),
      selectedHeight: parseFloat(selected?.style.height ?? '0'),
      selectedLeft: parseFloat(selected?.style.left ?? '0'),
      selectedName: document.querySelector('[data-ppt-layer-row][aria-selected="true"] .ppt-layer-name')?.textContent ?? '',
      selectedNames: selectedLayerNames.join(' '),
      selectedTop: parseFloat(selected?.style.top ?? '0'),
      selectedTransform: selected?.style.transform ?? '',
      selectedWidth: parseFloat(selected?.style.width ?? '0'),
      thumbCropXCount: document.querySelectorAll('.ppt-thumb-image[data-ppt-image-crop-x="25"]').length,
      thumbCropYCount: document.querySelectorAll('.ppt-thumb-image[data-ppt-image-crop-y="70"]').length,
      thumbFlipHCount: document.querySelectorAll('.ppt-thumb-image[data-ppt-flip-h="true"]').length,
      thumbFlipVCount: document.querySelectorAll('.ppt-thumb-image[data-ppt-flip-v="true"]').length,
      thumbContainCount: document.querySelectorAll('.ppt-thumb-image[data-ppt-image-fit="contain"]').length,
      thumbImageCount: document.querySelectorAll('.ppt-thumb-image').length,
    }
  })()`)
}

function getPPTTableState(page) {
  return page.eval(`(() => {
    const selected = document.querySelector('[data-selected="true"]')
    const selectedElements = [...document.querySelectorAll('[data-selected="true"]')]
    const selectedLayerNames = [...document.querySelectorAll('[data-ppt-layer-row][aria-selected="true"] .ppt-layer-name')]
      .map((element) => element.textContent ?? '')
    const stage = document.querySelector('.ppt-stage-shell')
    const tableCells = [...selected?.querySelectorAll('[data-ppt-table-cell]') ?? []]

    return {
      cellTexts: tableCells.map((cell) => cell.textContent ?? ''),
      inspectorSize: document.querySelector('[data-ppt-table-inspector-size]')?.textContent?.trim() ?? '',
      inspectorValue: document.querySelector('[data-ppt-style-field="table-data"]')?.value ?? '',
      fallbackHTMLImportFormat: stage?.getAttribute('data-ppt-fallback-html-import-format') ?? '',
      fallbackHTMLImportKind: stage?.getAttribute('data-ppt-fallback-html-import-kind') ?? '',
      fallbackHTMLImportModel: stage?.getAttribute('data-ppt-fallback-html-import-model') ?? '',
      fallbackHTMLImportSourceObject: stage?.getAttribute('data-ppt-fallback-html-import-source-object') ?? '',
      importExtensionLastClipboardActions: stage?.getAttribute('data-ppt-import-extension-last-clipboard-actions') ?? '',
      importExtensionLastDropAction: stage?.getAttribute('data-ppt-import-extension-last-drop-action') ?? '',
      paletteOpen: !!document.querySelector('[data-ppt-command-palette]'),
      selectedCols: Number(selected?.getAttribute('data-ppt-table-cols') ?? 0),
      selectedCount: selectedElements.length,
      selectedHeight: parseFloat(selected?.style.height ?? '0'),
      selectedId: selected?.getAttribute('data-ppt-element') ?? '',
      selectedKind: selected?.getAttribute('data-kind') ?? '',
      selectedKinds: selectedElements.map((element) => element.getAttribute('data-kind') ?? '').join(' '),
      selectedLeft: parseFloat(selected?.style.left ?? '0'),
      selectedName: document.querySelector('[data-ppt-layer-row][aria-selected="true"] .ppt-layer-name')?.textContent ?? '',
      selectedNames: selectedLayerNames.join(' '),
      selectedRows: Number(selected?.getAttribute('data-ppt-table-rows') ?? 0),
      tableImportCols: Number(stage?.getAttribute('data-ppt-table-import-cols') ?? 0),
      tableImportCount: Number(stage?.getAttribute('data-ppt-table-import-count') ?? 0),
      tableImportFormat: stage?.getAttribute('data-ppt-table-import-format') ?? '',
      tableImportModel: stage?.getAttribute('data-ppt-table-import-model') ?? '',
      tableImportName: stage?.getAttribute('data-ppt-table-import-name') ?? '',
      tableImportNames: stage?.getAttribute('data-ppt-table-import-names') ?? '',
      tableImportRows: Number(stage?.getAttribute('data-ppt-table-import-rows') ?? 0),
      tableClipboardCols: Number(stage?.getAttribute('data-ppt-table-clipboard-cols') ?? 0),
      tableClipboardHtmlLength: Number(stage?.getAttribute('data-ppt-table-clipboard-html-length') ?? 0),
      tableClipboardJsonMimeType: stage?.getAttribute('data-ppt-table-clipboard-json-mime-type') ?? '',
      tableClipboardModel: stage?.getAttribute('data-ppt-table-clipboard-model') ?? '',
      tableClipboardObject: stage?.getAttribute('data-ppt-table-clipboard-object') ?? '',
      tableClipboardPlainTextLength: Number(stage?.getAttribute('data-ppt-table-clipboard-plain-text-length') ?? 0),
      tableClipboardRows: Number(stage?.getAttribute('data-ppt-table-clipboard-rows') ?? 0),
      tableClipboardSourceSlide: stage?.getAttribute('data-ppt-table-clipboard-source-slide') ?? '',
      tableClipboardWriteMode: stage?.getAttribute('data-ppt-table-clipboard-write-mode') ?? '',
      tableRowsImportCols: Number(stage?.getAttribute('data-ppt-table-rows-import-cols') ?? 0),
      tableRowsImportFormat: stage?.getAttribute('data-ppt-table-rows-import-format') ?? '',
      tableRowsImportJsonLength: Number(stage?.getAttribute('data-ppt-table-rows-import-json-length') ?? 0),
      tableRowsImportModel: stage?.getAttribute('data-ppt-table-rows-import-model') ?? '',
      tableRowsImportObjects: stage?.getAttribute('data-ppt-table-rows-import-objects') ?? '',
      tableRowsImportRows: Number(stage?.getAttribute('data-ppt-table-rows-import-rows') ?? 0),
      tableRowsImportTargets: stage?.getAttribute('data-ppt-table-rows-import-command-targets') ?? '',
      selectedTop: parseFloat(selected?.style.top ?? '0'),
      selectedWidth: parseFloat(selected?.style.width ?? '0'),
      tableCount: document.querySelectorAll('[data-kind="table"]').length,
      thumbTableCount: document.querySelectorAll('.ppt-thumb-table').length,
    }
  })()`)
}

function getPPTTextPasteState(page) {
  return page.eval(`(() => {
    const selected = document.querySelector('[data-selected="true"]')
    const stage = document.querySelector('.ppt-stage-shell')
    const selectedParagraphs = selected
      ? [...selected.querySelectorAll('.ppt-text-paragraph')]
      : []
    const selectedRuns = selected
      ? [...selected.querySelectorAll('.ppt-text-paragraph > span')]
      : []

    return {
      exportCode: document.querySelector('.ppt-export-code')?.value ?? '',
      fallbackHTMLImportFormat: stage?.getAttribute('data-ppt-fallback-html-import-format') ?? '',
      fallbackHTMLImportKind: stage?.getAttribute('data-ppt-fallback-html-import-kind') ?? '',
      fallbackHTMLImportModel: stage?.getAttribute('data-ppt-fallback-html-import-model') ?? '',
      fallbackHTMLImportSourceObject: stage?.getAttribute('data-ppt-fallback-html-import-source-object') ?? '',
      importExtensionLastClipboardActions: stage?.getAttribute('data-ppt-import-extension-last-clipboard-actions') ?? '',
      redoEnabled: !document.querySelector('button[title="Redo"]')?.disabled,
      selectedBoldRunCount: selected?.querySelectorAll('[data-ppt-run-bold="true"]').length ?? 0,
      selectedBulletParagraphCount: selected?.querySelectorAll('[data-ppt-bullet="true"]').length ?? 0,
      selectedFontSize: selected ? getComputedStyle(selected).fontSize : '',
      selectedHeight: parseFloat(selected?.style.height ?? '0'),
      selectedHyperlink: selected?.getAttribute('data-ppt-hyperlink-url') ?? '',
      selectedId: selected?.getAttribute('data-ppt-element') ?? '',
      selectedItalicRunCount: selected?.querySelectorAll('[data-ppt-run-italic="true"]').length ?? 0,
      selectedKind: selected?.getAttribute('data-kind') ?? '',
      selectedLeft: parseFloat(selected?.style.left ?? '0'),
      selectedName: selected?.getAttribute('data-ppt-element-name') ?? '',
      selectedNumberedParagraphCount: selected?.querySelectorAll('[data-ppt-numbered="true"]').length ?? 0,
      selectedParagraphLineHeights: selectedParagraphs.map((paragraph) => paragraph.getAttribute('data-ppt-line-height') ?? ''),
      selectedParagraphSpacingAfter: selectedParagraphs.map((paragraph) => paragraph.getAttribute('data-ppt-spacing-after') ?? ''),
      selectedParagraphSpacingBefore: selectedParagraphs.map((paragraph) => paragraph.getAttribute('data-ppt-spacing-before') ?? ''),
      selectedRunColors: selectedRuns.map((run) => getComputedStyle(run).color),
      selectedRunFontSizes: selectedRuns.map((run) => getComputedStyle(run).fontSize),
      selectedText: selected?.textContent ?? '',
      selectedTextAlign: selected ? getComputedStyle(selected).textAlign : '',
      selectedTop: parseFloat(selected?.style.top ?? '0'),
      selectedUnderlineRunCount: selected?.querySelectorAll('[data-ppt-run-underline="true"]').length ?? 0,
      selectedWidth: parseFloat(selected?.style.width ?? '0'),
      richClipboardHTMLLength: Number(stage?.getAttribute('data-ppt-rich-clipboard-html-length') ?? 0),
      richClipboardJsonMimeType: stage?.getAttribute('data-ppt-rich-clipboard-json-mime-type') ?? '',
      richClipboardModel: stage?.getAttribute('data-ppt-rich-clipboard-model') ?? '',
      richClipboardPlainTextLength: Number(stage?.getAttribute('data-ppt-rich-clipboard-plain-text-length') ?? 0),
      richClipboardSelection: stage?.getAttribute('data-ppt-rich-clipboard-selection') ?? '',
      richClipboardWriteMode: stage?.getAttribute('data-ppt-rich-clipboard-write-mode') ?? '',
      textBoxCount: document.querySelectorAll('[data-kind="textBox"]').length,
      textBodyImportFormat: stage?.getAttribute('data-ppt-text-body-import-format') ?? '',
      textBodyImportJsonLength: Number(stage?.getAttribute('data-ppt-text-body-import-json-length') ?? 0),
      textBodyImportMode: stage?.getAttribute('data-ppt-text-body-import-mode') ?? '',
      textBodyImportModel: stage?.getAttribute('data-ppt-text-body-import-model') ?? '',
      textBodyImportObjects: stage?.getAttribute('data-ppt-text-body-import-objects') ?? '',
      textBodyImportParagraphs: Number(stage?.getAttribute('data-ppt-text-body-import-paragraphs') ?? 0),
      textBodyImportRuns: Number(stage?.getAttribute('data-ppt-text-body-import-runs') ?? 0),
      textBodyImportTargets: stage?.getAttribute('data-ppt-text-body-import-command-targets') ?? '',
      textBodyImportTextLength: Number(stage?.getAttribute('data-ppt-text-body-import-text-length') ?? 0),
      textPasteBoldRuns: Number(stage?.getAttribute('data-ppt-text-paste-bold-runs') ?? 0),
      textPasteBulletParagraphs: Number(stage?.getAttribute('data-ppt-text-paste-bullet-paragraphs') ?? 0),
      textPasteFormat: stage?.getAttribute('data-ppt-text-paste-format') ?? '',
      textPasteHyperlinkUrl: stage?.getAttribute('data-ppt-text-paste-hyperlink-url') ?? '',
      textPasteImporter: stage?.getAttribute('data-ppt-text-paste-importer') ?? '',
      textPasteLinkRuns: Number(stage?.getAttribute('data-ppt-text-paste-link-runs') ?? 0),
      textPasteModel: stage?.getAttribute('data-ppt-text-paste-model') ?? '',
      textPasteNumberedParagraphs: Number(stage?.getAttribute('data-ppt-text-paste-numbered-paragraphs') ?? 0),
      textPasteRichFallback: stage?.getAttribute('data-ppt-text-paste-rich-fallback') ?? '',
      textPasteSelection: stage?.getAttribute('data-ppt-text-paste-selection') ?? '',
      textPasteUnderlineRuns: Number(stage?.getAttribute('data-ppt-text-paste-underline-runs') ?? 0),
      thumbTextCount: document.querySelectorAll('.ppt-thumb-text').length,
      undoEnabled: !document.querySelector('button[title="Undo"]')?.disabled,
    }
  })()`)
}

function getPPTInlineEditState(page) {
  return page.eval(`(() => {
    const selected = document.querySelector('[data-selected="true"]')
    const editor = selected?.querySelector('.ppt-element-editor')
    const stage = document.querySelector('.ppt-stage-shell')

    return {
      editorActive: document.activeElement === editor,
      editorInlineEditModel: editor?.getAttribute('data-ppt-inline-edit-model') ?? '',
      elementCount: document.querySelectorAll('[data-ppt-element]').length,
      inlineEditElement: stage?.getAttribute('data-ppt-inline-edit-element') ?? '',
      inlineEditHistoryDirection: stage?.getAttribute('data-ppt-inline-edit-history-direction') ?? '',
      inlineEditInputType: stage?.getAttribute('data-ppt-inline-edit-input-type') ?? '',
      inlineEditLineBreak: stage?.getAttribute('data-ppt-inline-edit-line-break') ?? '',
      inlineEditModel: stage?.getAttribute('data-ppt-inline-edit-model') ?? '',
      inlineEditPasteText: stage?.getAttribute('data-ppt-inline-edit-paste-text') ?? '',
      mediaImportUrl: stage?.getAttribute('data-ppt-media-import-url') ?? '',
      selectedId: selected?.getAttribute('data-ppt-element') ?? '',
      selectedText: selected?.textContent ?? '',
      textPasteSelection: stage?.getAttribute('data-ppt-text-paste-selection') ?? '',
    }
  })()`)
}

function getPPTMediaImportState(page) {
  return page.eval(`(() => {
    const selected = document.querySelector('[data-selected="true"]')
    const stage = document.querySelector('.ppt-stage-shell')

    return {
      editing: !!document.querySelector('.ppt-element-editor[contenteditable="true"]'),
      exportCode: document.querySelector('.ppt-export-code')?.value ?? '',
      importExtensionLastClipboardActions: stage?.getAttribute('data-ppt-import-extension-last-clipboard-actions') ?? '',
      mediaImportImporter: stage?.getAttribute('data-ppt-media-import-importer') ?? '',
      mediaImportModel: stage?.getAttribute('data-ppt-media-import-model') ?? '',
      mediaImportSelection: stage?.getAttribute('data-ppt-media-import-selection') ?? '',
      mediaImportUrl: stage?.getAttribute('data-ppt-media-import-url') ?? '',
      mediaJSONImportFields: stage?.getAttribute('data-ppt-media-json-import-fields') ?? '',
      mediaJSONImportFormat: stage?.getAttribute('data-ppt-media-json-import-format') ?? '',
      mediaJSONImportImporter: stage?.getAttribute('data-ppt-media-json-import-importer') ?? '',
      mediaJSONImportJsonLength: Number(stage?.getAttribute('data-ppt-media-json-import-json-length') ?? 0),
      mediaJSONImportModel: stage?.getAttribute('data-ppt-media-json-import-model') ?? '',
      mediaJSONImportObject: stage?.getAttribute('data-ppt-media-json-import-object') ?? '',
      mediaJSONImportTitle: stage?.getAttribute('data-ppt-media-json-import-title') ?? '',
      mediaJSONImportUrl: stage?.getAttribute('data-ppt-media-json-import-url') ?? '',
      selectedHyperlink: selected?.getAttribute('data-ppt-hyperlink-url') ?? '',
      selectedId: selected?.getAttribute('data-ppt-element') ?? '',
      selectedKind: selected?.getAttribute('data-kind') ?? '',
      selectedLeft: parseFloat(selected?.style.left ?? '0'),
      selectedName: selected?.getAttribute('data-ppt-element-name') ?? '',
      selectedText: selected?.textContent ?? '',
      selectedTop: parseFloat(selected?.style.top ?? '0'),
      shapeCount: document.querySelectorAll('[data-kind="shape"]').length,
      textBoxCount: document.querySelectorAll('[data-kind="textBox"]').length,
      thumbHyperlinkCount: document.querySelectorAll('.ppt-thumb-shape[data-ppt-thumb-hyperlink-url]').length,
    }
  })()`)
}

function getPPTMarqueeDragBox(page, elementIds) {
  return page.eval(`((elementIds) => {
    const slide = document.querySelector('.ppt-slide')?.getBoundingClientRect()
    const rects = elementIds
      .map((id) => document.querySelector(\`[data-ppt-element="\${id}"]\`)?.getBoundingClientRect())
      .filter(Boolean)

    if (!slide || rects.length === 0) {
      return {
        endX: 0,
        endY: 0,
        startX: 0,
        startY: 0,
      }
    }

    const left = Math.min(...rects.map((rect) => rect.left))
    const right = Math.max(...rects.map((rect) => rect.right))
    const top = Math.min(...rects.map((rect) => rect.top))
    const bottom = Math.max(...rects.map((rect) => rect.bottom))

    return {
      endX: Math.min(slide.right - 8, right + 24),
      endY: Math.min(slide.bottom - 8, bottom + 24),
      startX: Math.max(slide.left + 8, left - 24),
      startY: Math.max(slide.top + 8, top - 24),
    }
  })(${JSON.stringify(elementIds)})`)
}

function getPPTMarqueeState(page) {
  return page.eval(`(() => {
    const stage = document.querySelector('.ppt-stage-shell')
    const marqueeSelection = (stage?.getAttribute('data-ppt-marquee-selection') ?? '')
      .split(' ')
      .filter(Boolean)

    return {
      card1Left: parseFloat(document.querySelector('[data-ppt-element="s1-card-1"]')?.style.left ?? '0'),
      editing: !!document.querySelector('.ppt-element-editor[contenteditable="true"]'),
      marqueeActive: stage?.getAttribute('data-ppt-marquee-active') ?? '',
      marqueeAdditive: stage?.getAttribute('data-ppt-marquee-additive') ?? '',
      marqueeCount: document.querySelectorAll('.ppt-marquee').length,
      marqueeHeight: Number(stage?.getAttribute('data-ppt-marquee-h') ?? 0),
      marqueeHistory: stage?.getAttribute('data-ppt-marquee-history') ?? '',
      marqueeModel: stage?.getAttribute('data-ppt-marquee-model') ?? '',
      marqueeSelection,
      marqueeWidth: Number(stage?.getAttribute('data-ppt-marquee-w') ?? 0),
      redoEnabled: !document.querySelector('button[title="Redo"]')?.disabled,
      selectedIds: [...document.querySelectorAll('[data-selected="true"]')]
        .map((element) => element.getAttribute('data-ppt-element'))
        .filter(Boolean),
      undoEnabled: !document.querySelector('button[title="Undo"]')?.disabled,
    }
  })()`)
}

function getPPTCommentState(page) {
  return page.eval(`(() => {
    const selected = document.querySelector('[data-selected="true"]')
    const selectedCard = selected?.querySelector('[data-ppt-comment-card]') ?? null
    const stage = document.querySelector('.ppt-stage-shell')
    const thread = document.querySelector('[data-ppt-comment-thread]')
    const threadBodies = Array.from(document.querySelectorAll('[data-ppt-comment-thread] [data-ppt-comment-thread-body]')).map((node) => node.textContent ?? '')

    return {
      commentCount: document.querySelectorAll('[data-kind="comment"]').length,
      commentImportBodyLength: Number(stage?.getAttribute('data-ppt-comment-import-body-length') ?? 0),
      commentImportCreatedAt: stage?.getAttribute('data-ppt-comment-import-created-at') ?? '',
      commentImportFields: stage?.getAttribute('data-ppt-comment-import-fields') ?? '',
      commentImportFormat: stage?.getAttribute('data-ppt-comment-import-format') ?? '',
      commentImportJsonLength: Number(stage?.getAttribute('data-ppt-comment-import-json-length') ?? 0),
      commentImportMessageCount: Number(stage?.getAttribute('data-ppt-comment-import-message-count') ?? 0),
      commentImportModel: stage?.getAttribute('data-ppt-comment-import-model') ?? '',
      commentImportObjects: stage?.getAttribute('data-ppt-comment-import-objects') ?? '',
      commentImportResolved: stage?.getAttribute('data-ppt-comment-import-resolved') ?? '',
      commentImportTargets: stage?.getAttribute('data-ppt-comment-import-command-targets') ?? '',
      creationTool: document.querySelector('.ppt-stage-shell')?.getAttribute('data-creation-tool') ?? '',
      commentThreadCommand: stage?.getAttribute('data-ppt-comment-thread-command') ?? '',
      commentThreadCommandBody: stage?.getAttribute('data-ppt-comment-thread-command-body') ?? '',
      commentThreadCommandCount: Number(stage?.getAttribute('data-ppt-comment-thread-command-count') ?? 0),
      commentThreadCommandObject: stage?.getAttribute('data-ppt-comment-thread-command-object') ?? '',
      commentThreadCommandSlide: stage?.getAttribute('data-ppt-comment-thread-command-slide') ?? '',
      commentThreadCommandType: stage?.getAttribute('data-ppt-comment-thread-command-type') ?? '',
      commentThreadCount: Number(thread?.getAttribute('data-ppt-comment-thread-count') ?? 0),
      commentThreadFirstBody: threadBodies[0] ?? '',
      commentThreadBodies: threadBodies,
      commentThreadMessageCount: document.querySelectorAll('[data-ppt-comment-thread] [data-ppt-comment-thread-message]').length,
      commentThreadModel: thread?.getAttribute('data-ppt-comment-thread-model') ?? '',
      inspectorBody: document.querySelector('[data-ppt-style-field="comment-body"]')?.value ?? '',
      inspectorResolved: document.querySelector('[data-ppt-style-field="comment-resolved"]')?.checked ?? false,
      paletteOpen: !!document.querySelector('[data-ppt-command-palette]'),
      replyInputValue: document.querySelector('[data-ppt-style-field="comment-reply"]')?.value ?? '',
      selectedBody: selectedCard?.querySelector('[data-ppt-comment-body]')?.textContent ?? '',
      selectedId: selected?.getAttribute('data-ppt-element') ?? '',
      selectedKind: selected?.getAttribute('data-kind') ?? '',
      selectedLeft: parseFloat(selected?.style.left ?? '0'),
      selectedName: document.querySelector('[data-ppt-layer-row][aria-selected="true"] .ppt-layer-name')?.textContent ?? '',
      selectedResolved: selected?.getAttribute('data-ppt-comment-resolved') ?? '',
      selectedTop: parseFloat(selected?.style.top ?? '0'),
      thumbCommentCount: document.querySelectorAll('.ppt-thumb-comment').length,
      toolbarPressed: document.querySelector('[data-ppt-insert-comment]')?.getAttribute('aria-pressed') ?? '',
    }
  })()`)
}

function getPPTLineState(page, elementId = null) {
  const selector = elementId
    ? `[data-ppt-element="${elementId}"]`
    : '[data-selected="true"]'

  return page.eval(`(() => {
    const selected = document.querySelector(${JSON.stringify(selector)})
    const selectedLine = selected?.querySelector('line') ?? null
    const selectedPath = selected?.querySelector('[data-ppt-line-path]') ?? null
    const selectedStrokeElement = selectedLine ?? selectedPath
    const left = parseFloat(selected?.style.left ?? '0')
    const top = parseFloat(selected?.style.top ?? '0')
    const x1 = Number(selected?.getAttribute('data-line-start-x') ?? selectedLine?.getAttribute('x1') ?? 0)
    const x2 = Number(selected?.getAttribute('data-line-end-x') ?? selectedLine?.getAttribute('x2') ?? 0)
    const y1 = Number(selected?.getAttribute('data-line-start-y') ?? selectedLine?.getAttribute('y1') ?? 0)
    const y2 = Number(selected?.getAttribute('data-line-end-y') ?? selectedLine?.getAttribute('y2') ?? 0)
    const stage = document.querySelector('.ppt-stage-shell')
    const lineStrokeDashField = document.querySelector('[data-ppt-style-field="line-stroke-dash"]')

    return {
      arrowToolModel: stage?.getAttribute('data-ppt-arrow-tool-model') ?? '',
      arrowToolPressed: document.querySelector('[data-ppt-insert-line="arrow"]')?.getAttribute('aria-pressed') ?? '',
      arrowToolShortcut: stage?.getAttribute('data-ppt-arrow-tool-shortcut') ?? '',
      command: stage?.getAttribute('data-ppt-stroke-line-style-command') ?? '',
      commandField: stage?.getAttribute('data-ppt-stroke-line-style-command-field') ?? '',
      commandObject: stage?.getAttribute('data-ppt-stroke-line-style-command-object') ?? '',
      commandSlide: stage?.getAttribute('data-ppt-stroke-line-style-command-slide') ?? '',
      commandType: stage?.getAttribute('data-ppt-stroke-line-style-command-type') ?? '',
      commandValue: stage?.getAttribute('data-ppt-stroke-line-style-command-value') ?? '',
      descriptorAttribute: lineStrokeDashField?.getAttribute('data-ppt-stroke-line-style-attribute') ?? '',
      descriptorAttributeValue: lineStrokeDashField?.getAttribute('data-ppt-stroke-line-style-attribute-value') ?? '',
      descriptorCommand: lineStrokeDashField?.getAttribute('data-ppt-stroke-line-style-command') ?? '',
      descriptorControl: lineStrokeDashField?.getAttribute('data-ppt-stroke-line-style-control') ?? '',
      descriptorSurface: lineStrokeDashField?.getAttribute('data-ppt-stroke-line-style-surface') ?? '',
      lineStyleImportCategories: stage?.getAttribute('data-ppt-line-style-import-categories') ?? '',
      lineStyleImportCommand: stage?.getAttribute('data-ppt-line-style-import-command') ?? '',
      lineStyleImportCommandTargets: stage?.getAttribute('data-ppt-line-style-import-command-targets') ?? '',
      lineStyleImportCommandType: stage?.getAttribute('data-ppt-line-style-import-command-type') ?? '',
      lineStyleImportFields: stage?.getAttribute('data-ppt-line-style-import-fields') ?? '',
      lineStyleImportFormat: stage?.getAttribute('data-ppt-line-style-import-format') ?? '',
      lineStyleImportJsonLength: stage?.getAttribute('data-ppt-line-style-import-json-length') ?? '',
      lineStyleImportModel: stage?.getAttribute('data-ppt-line-style-import-model') ?? '',
      lineStyleImportObjects: stage?.getAttribute('data-ppt-line-style-import-objects') ?? '',
      lineStyleImportStrokeColor: stage?.getAttribute('data-ppt-line-style-import-stroke-color') ?? '',
      lineStyleImportStrokeDash: stage?.getAttribute('data-ppt-line-style-import-stroke-dash') ?? '',
      lineStyleImportStrokeWidth: stage?.getAttribute('data-ppt-line-style-import-stroke-width') ?? '',
      lineCount: document.querySelectorAll('[data-kind="line"]').length,
      lineTool: stage?.getAttribute('data-line-tool') ?? '',
      endConnection: selected?.getAttribute('data-line-end-connection') ?? '',
      hasPath: !!selectedPath,
      inspectorDash: document.querySelector('[data-ppt-style-field="line-stroke-dash"]')?.value ?? '',
      markerEnd: selectedStrokeElement?.getAttribute('marker-end') ?? '',
      pathD: selectedPath?.getAttribute('d') ?? '',
      route: selected?.getAttribute('data-line-route') ?? '',
      routeHandleCount: document.querySelectorAll('[data-ppt-line-route-handle]').length,
      selectedId: selected?.getAttribute('data-ppt-element') ?? '',
      selectedDash: selected?.getAttribute('data-ppt-stroke-dash') ?? '',
      selectedKind: selected?.getAttribute('data-kind') ?? '',
      selectedName: document.querySelector('[data-ppt-layer-row][aria-selected="true"] .ppt-layer-name')?.textContent ?? '',
      selectedWidth: parseFloat(selected?.style.width ?? '0'),
      startConnection: selected?.getAttribute('data-line-start-connection') ?? '',
      stroke: selectedStrokeElement?.getAttribute('stroke') ?? '',
      strokeDasharray: selectedStrokeElement?.getAttribute('stroke-dasharray') ?? '',
      strokeWidth: selectedStrokeElement?.getAttribute('stroke-width') ?? '',
      styleClipboardCommand: stage?.getAttribute('data-ppt-style-clipboard-command') ?? '',
      styleClipboardCommandApplications: stage?.getAttribute('data-ppt-style-clipboard-command-applications') ?? '',
      styleClipboardCommandTargets: stage?.getAttribute('data-ppt-style-clipboard-command-targets') ?? '',
      styleClipboardCommandType: stage?.getAttribute('data-ppt-style-clipboard-command-type') ?? '',
      thumbDash: document.querySelector(\`.ppt-thumb[aria-current="page"] [data-ppt-thumb-element="\${selected?.getAttribute('data-ppt-element') ?? ''}"]\`)?.getAttribute('data-ppt-thumb-stroke-dash') ?? '',
      thumbLineCount: document.querySelectorAll('.ppt-thumb-line').length,
      endpointHandleCount: document.querySelectorAll('[data-ppt-line-endpoint]').length,
      x1,
      x2,
      y1,
      y2,
      worldX1: left + x1,
      worldX2: left + x2,
      worldY1: top + y1,
      worldY2: top + y2,
    }
  })()`)
}

function getPPTShapeStrokeDashState(page) {
  return page.eval(`(() => {
    const selected = document.querySelector('[data-selected="true"]')
    const targetId = selected?.getAttribute('data-ppt-element') ?? ''
    const thumb = document.querySelector(\`.ppt-thumb[aria-current="page"] [data-ppt-thumb-element="\${targetId}"]\`)
    const field = document.querySelector('[data-ppt-style-field="stroke-dash"]')
    const stage = document.querySelector('.ppt-stage-shell')

    return {
      command: stage?.getAttribute('data-ppt-stroke-line-style-command') ?? '',
      commandField: stage?.getAttribute('data-ppt-stroke-line-style-command-field') ?? '',
      commandObject: stage?.getAttribute('data-ppt-stroke-line-style-command-object') ?? '',
      commandSlide: stage?.getAttribute('data-ppt-stroke-line-style-command-slide') ?? '',
      commandType: stage?.getAttribute('data-ppt-stroke-line-style-command-type') ?? '',
      commandValue: stage?.getAttribute('data-ppt-stroke-line-style-command-value') ?? '',
      descriptorAttribute: field?.getAttribute('data-ppt-stroke-line-style-attribute') ?? '',
      descriptorAttributeValue: field?.getAttribute('data-ppt-stroke-line-style-attribute-value') ?? '',
      descriptorCommand: field?.getAttribute('data-ppt-stroke-line-style-command') ?? '',
      descriptorControl: field?.getAttribute('data-ppt-stroke-line-style-control') ?? '',
      descriptorSurface: field?.getAttribute('data-ppt-stroke-line-style-surface') ?? '',
      importCommandFields: stage?.getAttribute('data-ppt-stroke-line-style-import-command-fields') ?? '',
      importCommandTargets: stage?.getAttribute('data-ppt-stroke-line-style-import-command-targets') ?? '',
      importCommandTypes: stage?.getAttribute('data-ppt-stroke-line-style-import-command-types') ?? '',
      importCommandValues: stage?.getAttribute('data-ppt-stroke-line-style-import-command-values') ?? '',
      importCommands: stage?.getAttribute('data-ppt-stroke-line-style-import-commands') ?? '',
      importFields: stage?.getAttribute('data-ppt-stroke-line-style-import-fields') ?? '',
      importFormat: stage?.getAttribute('data-ppt-stroke-line-style-import-format') ?? '',
      importJsonLength: Number(stage?.getAttribute('data-ppt-stroke-line-style-import-json-length') ?? 0),
      importModel: stage?.getAttribute('data-ppt-stroke-line-style-import-model') ?? '',
      importObjects: stage?.getAttribute('data-ppt-stroke-line-style-import-objects') ?? '',
      importSlide: stage?.getAttribute('data-ppt-stroke-line-style-import-slide') ?? '',
      importValue: stage?.getAttribute('data-ppt-stroke-line-style-import-value') ?? '',
      inspectorDash: field?.value ?? '',
      model: stage?.getAttribute('data-ppt-stroke-line-style-model') ?? '',
      selectedBorderStyle: selected?.style.borderStyle ?? '',
      selectedDash: selected?.getAttribute('data-ppt-stroke-dash') ?? '',
      selectedId: targetId,
      selectedKind: selected?.getAttribute('data-kind') ?? '',
      thumbBorderStyle: thumb?.style.borderStyle ?? '',
      thumbDash: thumb?.getAttribute('data-ppt-thumb-stroke-dash') ?? '',
    }
  })()`)
}

function getPPTShapeFillOpacityState(page) {
  return page.eval(`(() => {
    const selected = document.querySelector('[data-selected="true"]')
    const targetId = selected?.getAttribute('data-ppt-element') ?? ''
    const thumb = document.querySelector(\`.ppt-thumb[aria-current="page"] [data-ppt-thumb-element="\${targetId}"]\`)
    const field = document.querySelector('[data-ppt-style-field="fill-opacity"]')
    const stage = document.querySelector('.ppt-stage-shell')

    return {
      command: stage?.getAttribute('data-ppt-fill-opacity-command') ?? '',
      commandField: stage?.getAttribute('data-ppt-fill-opacity-command-field') ?? '',
      commandObject: stage?.getAttribute('data-ppt-fill-opacity-command-object') ?? '',
      commandSlide: stage?.getAttribute('data-ppt-fill-opacity-command-slide') ?? '',
      commandType: stage?.getAttribute('data-ppt-fill-opacity-command-type') ?? '',
      commandValue: stage?.getAttribute('data-ppt-fill-opacity-command-value') ?? '',
      descriptorAttribute: field?.getAttribute('data-ppt-fill-opacity-attribute') ?? '',
      descriptorAttributeValue: field?.getAttribute('data-ppt-fill-opacity-attribute-value') ?? '',
      descriptorCommand: field?.getAttribute('data-ppt-fill-opacity-command') ?? '',
      descriptorControl: field?.getAttribute('data-ppt-fill-opacity-control') ?? '',
      descriptorSurface: field?.getAttribute('data-ppt-fill-opacity-surface') ?? '',
      importCommandFields: stage?.getAttribute('data-ppt-fill-opacity-import-command-fields') ?? '',
      importCommandTargets: stage?.getAttribute('data-ppt-fill-opacity-import-command-targets') ?? '',
      importCommandTypes: stage?.getAttribute('data-ppt-fill-opacity-import-command-types') ?? '',
      importCommandValues: stage?.getAttribute('data-ppt-fill-opacity-import-command-values') ?? '',
      importCommands: stage?.getAttribute('data-ppt-fill-opacity-import-commands') ?? '',
      importFields: stage?.getAttribute('data-ppt-fill-opacity-import-fields') ?? '',
      importFormat: stage?.getAttribute('data-ppt-fill-opacity-import-format') ?? '',
      importJsonLength: Number(stage?.getAttribute('data-ppt-fill-opacity-import-json-length') ?? 0),
      importModel: stage?.getAttribute('data-ppt-fill-opacity-import-model') ?? '',
      importObjects: stage?.getAttribute('data-ppt-fill-opacity-import-objects') ?? '',
      importSlide: stage?.getAttribute('data-ppt-fill-opacity-import-slide') ?? '',
      importValue: stage?.getAttribute('data-ppt-fill-opacity-import-value') ?? '',
      inspectorOpacity: field?.value ?? '',
      model: stage?.getAttribute('data-ppt-fill-opacity-model') ?? '',
      selectedBackground: selected?.style.background ?? '',
      selectedBorderStyle: selected?.style.borderStyle ?? '',
      selectedFillOpacity: selected?.getAttribute('data-ppt-fill-opacity') ?? '',
      selectedId: targetId,
      selectedKind: selected?.getAttribute('data-kind') ?? '',
      selectedObjectOpacity: selected?.style.opacity ?? '',
      thumbBackground: thumb?.style.background ?? '',
      thumbFillOpacity: thumb?.getAttribute('data-ppt-thumb-fill-opacity') ?? '',
    }
  })()`)
}

function getPPTShapeCornerRadiusState(page) {
  return page.eval(`(() => {
    const selected = document.querySelector('[data-selected="true"]')
    const targetId = selected?.getAttribute('data-ppt-element') ?? ''
    const thumb = document.querySelector(\`.ppt-thumb[aria-current="page"] [data-ppt-thumb-element="\${targetId}"]\`)
    const field = document.querySelector('[data-ppt-style-field="shape-corner-radius"]')
    const stage = document.querySelector('.ppt-stage-shell')

    return {
      command: stage?.getAttribute('data-ppt-corner-radius-command') ?? '',
      commandField: stage?.getAttribute('data-ppt-corner-radius-command-field') ?? '',
      commandObject: stage?.getAttribute('data-ppt-corner-radius-command-object') ?? '',
      commandSlide: stage?.getAttribute('data-ppt-corner-radius-command-slide') ?? '',
      commandType: stage?.getAttribute('data-ppt-corner-radius-command-type') ?? '',
      commandValue: stage?.getAttribute('data-ppt-corner-radius-command-value') ?? '',
      descriptorAttribute: field?.getAttribute('data-ppt-corner-radius-attribute') ?? '',
      descriptorAttributeValue: field?.getAttribute('data-ppt-corner-radius-attribute-value') ?? '',
      descriptorCommand: field?.getAttribute('data-ppt-corner-radius-command') ?? '',
      descriptorControl: field?.getAttribute('data-ppt-corner-radius-control') ?? '',
      descriptorSupported: field?.getAttribute('data-ppt-corner-radius-supported') ?? '',
      descriptorSurface: field?.getAttribute('data-ppt-corner-radius-surface') ?? '',
      importCommandFields: stage?.getAttribute('data-ppt-corner-radius-import-command-fields') ?? '',
      importCommandTargets: stage?.getAttribute('data-ppt-corner-radius-import-command-targets') ?? '',
      importCommandTypes: stage?.getAttribute('data-ppt-corner-radius-import-command-types') ?? '',
      importCommandValues: stage?.getAttribute('data-ppt-corner-radius-import-command-values') ?? '',
      importCommands: stage?.getAttribute('data-ppt-corner-radius-import-commands') ?? '',
      importFields: stage?.getAttribute('data-ppt-corner-radius-import-fields') ?? '',
      importFormat: stage?.getAttribute('data-ppt-corner-radius-import-format') ?? '',
      importJsonLength: Number(stage?.getAttribute('data-ppt-corner-radius-import-json-length') ?? 0),
      importModel: stage?.getAttribute('data-ppt-corner-radius-import-model') ?? '',
      importObjects: stage?.getAttribute('data-ppt-corner-radius-import-objects') ?? '',
      importSlide: stage?.getAttribute('data-ppt-corner-radius-import-slide') ?? '',
      importValue: stage?.getAttribute('data-ppt-corner-radius-import-value') ?? '',
      inspectorRadius: field?.value ?? '',
      model: stage?.getAttribute('data-ppt-corner-radius-model') ?? '',
      selectedBorderRadius: selected?.style.borderRadius ?? '',
      selectedCornerRadius: selected?.getAttribute('data-ppt-corner-radius') ?? '',
      selectedId: targetId,
      selectedKind: selected?.getAttribute('data-kind') ?? '',
      thumbBorderRadius: thumb?.style.borderRadius ?? '',
      thumbCornerRadius: thumb?.getAttribute('data-ppt-thumb-corner-radius') ?? '',
    }
  })()`)
}

function getPPTFormatPainterSelectedShapeState(page) {
  return page.eval(`(() => {
    const selected = document.querySelector('[data-selected="true"]')
    const targetId = selected?.getAttribute('data-ppt-element') ?? ''
    const layerName = document.querySelector('[data-ppt-layer-row][aria-selected="true"] .ppt-layer-name')
    const shell = document.querySelector('.ppt-stage-shell')

    return {
      background: selected?.style.background ?? '',
      borderColor: selected?.style.borderColor ?? '',
      borderRadius: selected?.style.borderRadius ?? '',
      borderStyle: selected?.style.borderStyle ?? '',
      borderWidth: selected?.style.borderWidth ?? '',
      cornerRadius: selected?.getAttribute('data-ppt-corner-radius') ?? '',
      fillOpacity: selected?.getAttribute('data-ppt-fill-opacity') ?? '',
      filter: selected?.style.filter ?? '',
      height: selected?.style.height ?? '',
      left: selected?.style.left ?? '',
      name: layerName?.textContent ?? '',
      objectOpacity: selected?.getAttribute('data-ppt-opacity') ?? '',
      paletteOpen: !!document.querySelector('[data-ppt-command-palette]'),
      selectedId: targetId,
      selectedKind: selected?.getAttribute('data-kind') ?? '',
      shadow: selected?.getAttribute('data-ppt-shadow') ?? '',
      shadowOpacity: selected?.getAttribute('data-ppt-shadow-opacity') ?? '',
      shape: selected?.getAttribute('data-shape') ?? '',
      styleOpacity: selected?.style.opacity ?? '',
      styleClipboardCommand: shell?.getAttribute('data-ppt-style-clipboard-command') ?? '',
      styleClipboardCommandApplications: shell?.getAttribute('data-ppt-style-clipboard-command-applications') ?? '',
      styleClipboardCommandSelection: shell?.getAttribute('data-ppt-style-clipboard-command-selection') ?? '',
      styleClipboardCommandTargets: shell?.getAttribute('data-ppt-style-clipboard-command-targets') ?? '',
      styleClipboardCommandType: shell?.getAttribute('data-ppt-style-clipboard-command-type') ?? '',
      shapeStyleImportCategories: shell?.getAttribute('data-ppt-shape-style-import-categories') ?? '',
      shapeStyleImportCommand: shell?.getAttribute('data-ppt-shape-style-import-command') ?? '',
      shapeStyleImportCommandTargets: shell?.getAttribute('data-ppt-shape-style-import-command-targets') ?? '',
      shapeStyleImportCommandType: shell?.getAttribute('data-ppt-shape-style-import-command-type') ?? '',
      shapeStyleImportCornerRadius: shell?.getAttribute('data-ppt-shape-style-import-corner-radius') ?? '',
      shapeStyleImportFields: shell?.getAttribute('data-ppt-shape-style-import-fields') ?? '',
      shapeStyleImportFillColor: shell?.getAttribute('data-ppt-shape-style-import-fill-color') ?? '',
      shapeStyleImportFillOpacity: shell?.getAttribute('data-ppt-shape-style-import-fill-opacity') ?? '',
      shapeStyleImportFormat: shell?.getAttribute('data-ppt-shape-style-import-format') ?? '',
      shapeStyleImportJsonLength: shell?.getAttribute('data-ppt-shape-style-import-json-length') ?? '',
      shapeStyleImportModel: shell?.getAttribute('data-ppt-shape-style-import-model') ?? '',
      shapeStyleImportObjects: shell?.getAttribute('data-ppt-shape-style-import-objects') ?? '',
      shapeStyleImportStrokeColor: shell?.getAttribute('data-ppt-shape-style-import-stroke-color') ?? '',
      shapeStyleImportStrokeDash: shell?.getAttribute('data-ppt-shape-style-import-stroke-dash') ?? '',
      shapeStyleImportStrokeWidth: shell?.getAttribute('data-ppt-shape-style-import-stroke-width') ?? '',
      strokeDash: selected?.getAttribute('data-ppt-stroke-dash') ?? '',
      top: selected?.style.top ?? '',
      width: selected?.style.width ?? '',
    }
  })()`)
}

function getPPTFreeformState(page) {
  return page.eval(`(() => {
    const selected = document.querySelector('[data-selected="true"]')
    const selectedId = selected?.getAttribute('data-ppt-element') ?? ''
    const selectedPath = selected?.querySelector('[data-ppt-freeform-path]') ?? null
    const layerName = selectedId
      ? document.querySelector(\`[data-ppt-layer-row="\${selectedId}"] .ppt-layer-name\`)
      : null

    return {
      creationTool: document.querySelector('.ppt-stage-shell')?.getAttribute('data-creation-tool') ?? '',
      drawingTool: document.querySelector('.ppt-stage-shell')?.getAttribute('data-ppt-drawing-tool') ?? '',
      elementCount: document.querySelectorAll('[data-ppt-element]').length,
      eraserActive: document.querySelector('.ppt-stage-shell')?.getAttribute('data-ppt-eraser-tool-active') ?? '',
      eraserHitCount: Number(document.querySelector('.ppt-stage-shell')?.getAttribute('data-ppt-eraser-hit-count') ?? 0),
      eraserHitIds: [...document.querySelectorAll('[data-ppt-eraser-hit="true"]')]
        .map((element) => element.getAttribute('data-ppt-element')),
      eraserToolbarPressed: document.querySelector('[data-ppt-eraser-tool]')?.getAttribute('aria-pressed') ?? '',
      freeformCount: document.querySelectorAll('[data-kind="freeform"]').length,
      freeformIds: [...document.querySelectorAll('[data-kind="freeform"]')]
        .map((element) => element.getAttribute('data-ppt-element')),
      highlighterToolbarPressed: document.querySelector('[data-ppt-insert-tool="highlight"]')?.getAttribute('aria-pressed') ?? '',
      markerToolbarPressed: document.querySelector('[data-ppt-insert-tool="marker"]')?.getAttribute('aria-pressed') ?? '',
      paletteOpen: !!document.querySelector('[data-ppt-command-palette]'),
      paletteEraser: !!document.querySelector('[data-ppt-command-palette-item="tool:eraser"]'),
      paletteHighlight: !!document.querySelector('[data-ppt-command-palette-item="tool:highlight"]'),
      paletteMarker: !!document.querySelector('[data-ppt-command-palette-item="tool:marker"]'),
      palettePen: !!document.querySelector('[data-ppt-command-palette-item="tool:pen"]'),
      pasteDisabled: document.querySelector('[data-ppt-command="paste-formatting"]')?.disabled ?? true,
      pathD: selectedPath?.getAttribute('d') ?? '',
      pointCount: Number(selected?.getAttribute('data-ppt-freeform-points') ?? 0),
      selectedDash: selected?.getAttribute('data-ppt-stroke-dash') ?? '',
      selectedId,
      selectedKind: selected?.getAttribute('data-kind') ?? '',
      selectedLeft: parseFloat(selected?.style.left ?? '0'),
      selectedName: layerName?.textContent ?? '',
      selectedOpacity: selected?.getAttribute('data-ppt-opacity') ?? '',
      selectedRotation: selected?.getAttribute('data-rotation') ?? '',
      selectedStyleOpacity: selected?.style.opacity ?? '',
      selectedTop: parseFloat(selected?.style.top ?? '0'),
      selectedTransform: selected?.style.transform ?? '',
      selectedWidth: parseFloat(selected?.style.width ?? '0'),
      shapeCount: document.querySelectorAll('[data-kind="shape"]').length,
      stroke: selectedPath?.getAttribute('stroke') ?? '',
      strokeDasharray: selectedPath?.getAttribute('stroke-dasharray') ?? '',
      strokeWidth: selectedPath?.getAttribute('stroke-width') ?? '',
      thumbFreeformCount: document.querySelectorAll('.ppt-thumb-freeform').length,
      toolbarPressed: document.querySelector('[data-ppt-insert-tool="pen"]')?.getAttribute('aria-pressed') ?? '',
    }
  })()`)
}

function getPPTStickySectionState(page) {
  return page.eval(`(() => {
    const selected = document.querySelector('[data-selected="true"]')
    const selectedId = selected?.getAttribute('data-ppt-element') ?? ''
    const layerName = selectedId
      ? document.querySelector(\`[data-ppt-layer-row="\${selectedId}"] .ppt-layer-name\`)
      : null
    const stage = document.querySelector('.ppt-stage-shell')

    return {
      creationTool: stage?.getAttribute('data-creation-tool') ?? '',
      editingSticky: selectedId.length > 0 &&
        document.activeElement?.matches(\`[data-ppt-element="\${selectedId}"] .ppt-element-editor\`) === true,
      elementCount: document.querySelectorAll('[data-ppt-element]').length,
      exportCode: document.querySelector('.ppt-export-code')?.value ?? '',
      sectionToolbarPressed: document.querySelector('[data-ppt-insert-tool="section"]')?.getAttribute('aria-pressed') ?? '',
      sectionToolModel: stage?.getAttribute('data-ppt-section-tool-model') ?? '',
      sectionToolShortcut: stage?.getAttribute('data-ppt-section-tool-shortcut') ?? '',
      selectedFillOpacity: selected?.getAttribute('data-ppt-fill-opacity') ?? '',
      selectedHeight: parseFloat(selected?.style.height ?? '0'),
      selectedId,
      selectedKind: selected?.getAttribute('data-kind') ?? '',
      selectedName: layerName?.textContent ?? '',
      selectedShape: selected?.getAttribute('data-shape') ?? '',
      selectedStrokeDash: selected?.getAttribute('data-ppt-stroke-dash') ?? '',
      selectedText: selected?.textContent ?? '',
      selectedWidth: parseFloat(selected?.style.width ?? '0'),
      stickyToolbarPressed: document.querySelector('[data-ppt-insert-tool="sticky"]')?.getAttribute('aria-pressed') ?? '',
      stickyToolModel: stage?.getAttribute('data-ppt-sticky-tool-model') ?? '',
      stickyToolShortcut: stage?.getAttribute('data-ppt-sticky-tool-shortcut') ?? '',
      undoEnabled: !document.querySelector('button[title="Undo"]')?.disabled,
    }
  })()`)
}

function getPPTTextOverflowState(page) {
  return page.eval(`(() => {
    const selected = document.querySelector('[data-selected="true"]')
    const capsule = document.querySelector('.ppt-size-capsule')
    const inspector = document.querySelector('[data-ppt-text-overflow-inspector]')
    const autoFitButton = document.querySelector('[data-ppt-style-action="text-auto-fit"]')
    const stage = document.querySelector('.ppt-stage-shell')
    const selectedId = selected?.getAttribute('data-ppt-element') ?? ''
    const exportCode = document.querySelector('.ppt-export-code')?.value ?? ''

    return {
      autoFitDisabled: autoFitButton?.disabled ?? true,
      capsuleOverflow: capsule?.getAttribute('data-ppt-text-overflow') ?? '',
      capsuleOverflowAxis: capsule?.getAttribute('data-ppt-text-overflow-indicator-axis') ?? '',
      capsuleOverflowIndicatorVisible: capsule?.getAttribute('data-ppt-text-overflow-indicator-visible') ?? '',
      capsuleSizeMode: capsule?.getAttribute('data-ppt-text-autofit-size-mode') ?? '',
      capsuleText: capsule?.textContent ?? '',
      inspectorAutoFitCommand: inspector?.getAttribute('data-ppt-text-autofit-command') ?? '',
      inspectorAutoFitCommandHandle: inspector?.getAttribute('data-ppt-text-autofit-command-handle') ?? '',
      inspectorAutoFitCommandObject: inspector?.getAttribute('data-ppt-text-autofit-command-object') ?? '',
      inspectorAutoFitCommandSizeMode: inspector?.getAttribute('data-ppt-text-autofit-command-size-mode') ?? '',
      inspectorAutoFitCommandType: inspector?.getAttribute('data-ppt-text-autofit-command-type') ?? '',
      inspectorAutoFitModel: inspector?.getAttribute('data-ppt-text-autofit-model') ?? '',
      inspectorAutoFitSizeMode: inspector?.getAttribute('data-ppt-text-autofit-size-mode') ?? '',
      inspectorAutoFit: inspector?.getAttribute('data-ppt-text-autofit') ?? '',
      inspectorOverflow: inspector?.getAttribute('data-ppt-text-overflow') ?? '',
      inspectorOverflowAxis: inspector?.getAttribute('data-ppt-text-overflow-indicator-axis') ?? '',
      inspectorOverflowIndicatorVisible: inspector?.getAttribute('data-ppt-text-overflow-indicator-visible') ?? '',
      inspectorText: document.querySelector('[data-ppt-style-field="text"]')?.value ?? '',
      exportHasAutoFit: selectedId !== '' &&
        exportCode.includes(\`data-ppt-element="\${selectedId}"\`) &&
        exportCode.includes('data-ppt-text-autofit="resizeShapeToFitText"'),
      selectedAutoFit: selected?.getAttribute('data-ppt-text-autofit') ?? '',
      selectedAutoFitModel: selected?.getAttribute('data-ppt-text-autofit-model') ?? '',
      selectedAutoFitSizeMode: selected?.getAttribute('data-ppt-text-autofit-size-mode') ?? '',
      selectedHeight: parseFloat(selected?.style.height ?? '0'),
      selectedId,
      selectedKind: selected?.getAttribute('data-kind') ?? '',
      selectedOverflow: selected?.getAttribute('data-ppt-text-overflow') ?? '',
      selectedOverflowAxis: selected?.getAttribute('data-ppt-text-overflow-indicator-axis') ?? '',
      selectedOverflowIndicatorModel: selected?.getAttribute('data-ppt-text-overflow-indicator-model') ?? '',
      selectedOverflowIndicatorSlide: selected?.getAttribute('data-ppt-text-overflow-indicator-slide') ?? '',
      selectedOverflowIndicatorVisible: selected?.getAttribute('data-ppt-text-overflow-indicator-visible') ?? '',
      selectedText: selected?.querySelector('.ppt-element-editor')?.textContent ?? '',
      selectedWidth: parseFloat(selected?.style.width ?? '0'),
      stageAutoFitCommand: stage?.getAttribute('data-ppt-text-autofit-command') ?? '',
      stageAutoFitCommandHandle: stage?.getAttribute('data-ppt-text-autofit-command-handle') ?? '',
      stageAutoFitCommandObject: stage?.getAttribute('data-ppt-text-autofit-command-object') ?? '',
      stageAutoFitCommandSelection: stage?.getAttribute('data-ppt-text-autofit-command-selection') ?? '',
      stageAutoFitCommandSizeMode: stage?.getAttribute('data-ppt-text-autofit-command-size-mode') ?? '',
      stageAutoFitCommandSlide: stage?.getAttribute('data-ppt-text-autofit-command-slide') ?? '',
      stageAutoFitCommandType: stage?.getAttribute('data-ppt-text-autofit-command-type') ?? '',
      stageAutoFitImportCommandHandles: stage?.getAttribute('data-ppt-text-autofit-import-command-handles') ?? '',
      stageAutoFitImportCommandTargets: stage?.getAttribute('data-ppt-text-autofit-import-command-targets') ?? '',
      stageAutoFitImportCommandTypes: stage?.getAttribute('data-ppt-text-autofit-import-command-types') ?? '',
      stageAutoFitImportCommands: stage?.getAttribute('data-ppt-text-autofit-import-commands') ?? '',
      stageAutoFitImportFields: stage?.getAttribute('data-ppt-text-autofit-import-fields') ?? '',
      stageAutoFitImportFormat: stage?.getAttribute('data-ppt-text-autofit-import-format') ?? '',
      stageAutoFitImportJsonLength: Number(stage?.getAttribute('data-ppt-text-autofit-import-json-length') ?? 0),
      stageAutoFitImportMode: stage?.getAttribute('data-ppt-text-autofit-import-mode') ?? '',
      stageAutoFitImportModel: stage?.getAttribute('data-ppt-text-autofit-import-model') ?? '',
      stageAutoFitImportObjects: stage?.getAttribute('data-ppt-text-autofit-import-objects') ?? '',
      stageAutoFitModel: stage?.getAttribute('data-ppt-text-autofit-model') ?? '',
      stageAutoFitSizeModes: stage?.getAttribute('data-ppt-text-autofit-size-modes') ?? '',
      stageOverflowAxis: stage?.getAttribute('data-ppt-text-overflow-indicator-axis') ?? '',
      stageOverflowIndicatorModel: stage?.getAttribute('data-ppt-text-overflow-indicator-model') ?? '',
      stageOverflowIndicatorObject: stage?.getAttribute('data-ppt-text-overflow-indicator-object') ?? '',
      stageOverflowIndicatorSizeMode: stage?.getAttribute('data-ppt-text-overflow-indicator-size-mode') ?? '',
      stageOverflowIndicatorSlide: stage?.getAttribute('data-ppt-text-overflow-indicator-slide') ?? '',
      stageOverflowIndicatorVisible: stage?.getAttribute('data-ppt-text-overflow-indicator-visible') ?? '',
      textCount: document.querySelectorAll('[data-kind="textBox"]').length,
    }
  })()`)
}

function getPPTSlideMetadataState(page) {
  return page.eval(`(() => {
    const inspector = document.querySelector('[data-ppt-slide-metadata-inspector]')
    const objectInspector = document.querySelector('[data-ppt-object-inspector]')
    const stage = document.querySelector('.ppt-stage-shell')
    const fields = [...document.querySelectorAll('[data-ppt-slide-metadata-field]')]
    const editableByField = Object.fromEntries(fields.map((field) => [
      field.getAttribute('data-ppt-slide-metadata-field') ?? '',
      field.getAttribute('data-ppt-slide-metadata-editable') ?? '',
    ]))

    return {
      commandSlot: inspector?.getAttribute('data-ppt-slide-metadata-command-slot') ?? '',
      commands: fields.map((field) => field.getAttribute('data-ppt-slide-metadata-command') ?? ''),
      editableByField,
      backgroundValue: document.querySelector('[data-ppt-slide-field="background"]')?.value ?? '',
      fieldCount: Number(inspector?.getAttribute('data-ppt-slide-metadata-field-count') ?? 0),
      fields: fields.map((field) => field.getAttribute('data-ppt-slide-metadata-field') ?? ''),
      inspector: !!inspector,
      inspectorSurface: inspector?.getAttribute('data-ppt-inspector-surface') ?? '',
      name: document.querySelector('[data-ppt-slide-field="name"]')?.value ?? '',
      notes: document.querySelector('[data-ppt-slide-field="notes"]')?.value ?? '',
      objectActive: objectInspector?.getAttribute('data-ppt-object-inspector-active') ?? '',
      objectPriority: objectInspector?.getAttribute('data-ppt-object-inspector-priority') ?? '',
      orientationValue: document.querySelector('[data-ppt-slide-metadata-field="orientation"]')?.getAttribute('data-ppt-slide-metadata-value') ?? '',
      slideMetadataImportBackground: stage?.getAttribute('data-ppt-slide-metadata-import-background') ?? '',
      slideMetadataImportCommands: stage?.getAttribute('data-ppt-slide-metadata-import-commands') ?? '',
      slideMetadataImportFields: stage?.getAttribute('data-ppt-slide-metadata-import-fields') ?? '',
      slideMetadataImportFormat: stage?.getAttribute('data-ppt-slide-metadata-import-format') ?? '',
      slideMetadataImportJsonLength: Number(stage?.getAttribute('data-ppt-slide-metadata-import-json-length') ?? 0),
      slideMetadataImportModel: stage?.getAttribute('data-ppt-slide-metadata-import-model') ?? '',
      slideMetadataImportName: stage?.getAttribute('data-ppt-slide-metadata-import-name') ?? '',
      slideMetadataImportNotesLength: Number(stage?.getAttribute('data-ppt-slide-metadata-import-notes-length') ?? 0),
      slideMetadataImportSlide: stage?.getAttribute('data-ppt-slide-metadata-import-slide') ?? '',
      slideNotesImportFormat: stage?.getAttribute('data-ppt-slide-notes-import-format') ?? '',
      slideNotesImportModel: stage?.getAttribute('data-ppt-slide-notes-import-model') ?? '',
      slideNotesImportNotesLength: Number(stage?.getAttribute('data-ppt-slide-notes-import-notes-length') ?? 0),
      slideNotesImportSlide: stage?.getAttribute('data-ppt-slide-notes-import-slide') ?? '',
      slideNotesImportTextLength: Number(stage?.getAttribute('data-ppt-slide-notes-import-text-length') ?? 0),
      sizeValue: document.querySelector('[data-ppt-slide-metadata-field="size"]')?.getAttribute('data-ppt-slide-metadata-value') ?? '',
      slideBackground: document.querySelector('.ppt-slide')?.style.background ?? '',
      slideCount: Number(inspector?.getAttribute('data-ppt-slide-metadata-slide-count') ?? 0),
      slideId: inspector?.getAttribute('data-ppt-slide-metadata-slide-id') ?? '',
      slidePriority: inspector?.getAttribute('data-ppt-slide-inspector-priority') ?? '',
      surface: inspector?.getAttribute('data-ppt-slide-metadata-surface') ?? '',
      thumbName: document.querySelector('.ppt-thumb[aria-current="page"] .ppt-thumb-name')?.textContent ?? '',
    }
  })()`)
}

function getPPTInspectorTabsState(page) {
  return page.eval(`(() => {
    const inspector = document.querySelector('.ppt-inspector')
    const tablist = document.querySelector('[data-ppt-inspector-tabs]')
    const tabs = [...document.querySelectorAll('[data-ppt-inspector-tab]')]
    const panels = [...document.querySelectorAll('[data-ppt-inspector-tabpanel]')]
    const tabDetails = tabs.map((tab) => ({
      controls: tab.getAttribute('aria-controls') ?? '',
      focused: document.activeElement === tab,
      id: tab.getAttribute('data-ppt-inspector-tab') ?? '',
      role: tab.getAttribute('role') ?? '',
      selected: tab.getAttribute('aria-selected') ?? '',
      tabId: tab.id,
      tabIndex: tab.tabIndex,
    }))
    const panelDetails = panels.map((panel) => ({
      active: panel.getAttribute('data-ppt-inspector-tabpanel-active') ?? '',
      hidden: panel.hasAttribute('hidden'),
      id: panel.getAttribute('data-ppt-inspector-tabpanel') ?? '',
      labelledBy: panel.getAttribute('aria-labelledby') ?? '',
      panelId: panel.id,
      role: panel.getAttribute('role') ?? '',
    }))

    return {
      activation: tablist?.getAttribute('data-ppt-inspector-tabs-activation') ?? '',
      activePanelIds: panelDetails
        .filter((panel) => panel.active === 'true' && !panel.hidden)
        .map((panel) => panel.id),
      activeTab: inspector?.getAttribute('data-ppt-inspector-active-tab') ?? '',
      focusedTab: tabDetails.find((tab) => tab.focused)?.id ?? '',
      keyboard: tablist?.getAttribute('data-ppt-inspector-tabs-keyboard') ?? '',
      model: tablist?.getAttribute('data-ppt-inspector-tabs-model') ?? '',
      panelCount: panels.length,
      panels: panelDetails,
      relationshipsValid: tabDetails.every((tab) =>
        panelDetails.some((panel) =>
          panel.panelId === tab.controls &&
            panel.labelledBy === tab.tabId &&
            panel.role === 'tabpanel',
        ),
      ),
      selectedTabIds: tabDetails
        .filter((tab) => tab.selected === 'true')
        .map((tab) => tab.id),
      tabCount: tabs.length,
      tablistRole: tablist?.getAttribute('role') ?? '',
      tabs: tabDetails,
      tabStopIds: tabDetails
        .filter((tab) => tab.tabIndex === 0)
        .map((tab) => tab.id),
    }
  })()`)
}

function getPPTSlideTransitionState(page) {
  return page.eval(`(() => {
    const inspector = document.querySelector('[data-ppt-slide-transition]')
    const shell = document.querySelector('.ppt-stage-shell')
    const stage = document.querySelector('.ppt-stage-world .ppt-slide')

    return {
      advanceAfter: document.querySelector('[data-ppt-slide-transition-field="advanceAfterMs"]')?.value ?? '',
      advanceOnClick: document.querySelector('[data-ppt-slide-transition-field="advanceOnClick"]')?.checked ? 'true' : 'false',
      command: shell?.getAttribute('data-ppt-transition-command') ?? '',
      commandField: shell?.getAttribute('data-ppt-transition-command-field') ?? '',
      commandSelection: shell?.getAttribute('data-ppt-transition-command-selection') ?? '',
      commandSlide: shell?.getAttribute('data-ppt-transition-command-slide') ?? '',
      commandType: shell?.getAttribute('data-ppt-transition-command-type') ?? '',
      commandValue: shell?.getAttribute('data-ppt-transition-command-value') ?? '',
      duration: document.querySelector('[data-ppt-slide-transition-field="durationMs"]')?.value ?? '',
      inspector: !!inspector,
      importAdvanceAfter: shell?.getAttribute('data-ppt-transition-import-advance-after') ?? '',
      importAdvanceOnClick: shell?.getAttribute('data-ppt-transition-import-advance-on-click') ?? '',
      importCommandFields: shell?.getAttribute('data-ppt-transition-import-command-fields') ?? '',
      importCommands: shell?.getAttribute('data-ppt-transition-import-commands') ?? '',
      importDuration: shell?.getAttribute('data-ppt-transition-import-duration') ?? '',
      importFields: shell?.getAttribute('data-ppt-transition-import-fields') ?? '',
      importFormat: shell?.getAttribute('data-ppt-transition-import-format') ?? '',
      importJsonLength: Number(shell?.getAttribute('data-ppt-transition-import-json-length') ?? 0),
      importModel: shell?.getAttribute('data-ppt-transition-import-model') ?? '',
      importSlide: shell?.getAttribute('data-ppt-transition-import-slide') ?? '',
      importType: shell?.getAttribute('data-ppt-transition-import-type') ?? '',
      inspectorAdvanceAfter: inspector?.getAttribute('data-ppt-transition-advance-after') ?? '',
      inspectorAdvanceOnClick: inspector?.getAttribute('data-ppt-transition-advance-on-click') ?? '',
      inspectorDuration: inspector?.getAttribute('data-ppt-transition-duration') ?? '',
      inspectorModel: inspector?.getAttribute('data-ppt-transition-model') ?? '',
      inspectorSlide: inspector?.getAttribute('data-ppt-transition-slide') ?? '',
      inspectorType: inspector?.getAttribute('data-ppt-transition-type') ?? '',
      inspectorTypes: inspector?.getAttribute('data-ppt-transition-types') ?? '',
      stageAdvanceAfter: stage?.getAttribute('data-ppt-transition-advance-after') ?? '',
      stageAdvanceOnClick: stage?.getAttribute('data-ppt-transition-advance-on-click') ?? '',
      stageDuration: stage?.getAttribute('data-ppt-transition-duration') ?? '',
      stageModel: stage?.getAttribute('data-ppt-transition-model') ?? '',
      stageType: stage?.getAttribute('data-ppt-transition-type') ?? '',
      type: document.querySelector('[data-ppt-slide-transition-field="type"]')?.value ?? '',
    }
  })()`)
}

function getPPTObjectAnimationState(page) {
  return page.eval(`(() => {
    const inspector = document.querySelector('[data-ppt-object-animation-inspector]')
    const selected = document.querySelector('[data-selected="true"]')
    const stage = document.querySelector('.ppt-stage-shell')
    const type = document.querySelector('[data-ppt-animation-field="type"]')
    const trigger = document.querySelector('[data-ppt-animation-field="trigger"]')
    const duration = document.querySelector('[data-ppt-animation-field="durationMs"]')
    const delay = document.querySelector('[data-ppt-animation-field="delayMs"]')
    const order = document.querySelector('[data-ppt-animation-field="order"]')

    return {
      command: stage?.getAttribute('data-ppt-object-animation-command') ?? '',
      commandField: stage?.getAttribute('data-ppt-object-animation-command-field') ?? '',
      commandObject: stage?.getAttribute('data-ppt-object-animation-command-object') ?? '',
      commandSlide: stage?.getAttribute('data-ppt-object-animation-command-slide') ?? '',
      commandType: stage?.getAttribute('data-ppt-object-animation-command-type') ?? '',
      commandValue: stage?.getAttribute('data-ppt-object-animation-command-value') ?? '',
      buildOrder: stage?.getAttribute('data-ppt-object-animation-build-order') ?? '',
      buildOrderModel: stage?.getAttribute('data-ppt-object-animation-build-order-model') ?? '',
      delay: delay?.value ?? '',
      delayCommand: delay?.getAttribute('data-ppt-animation-command') ?? '',
      descriptorDelayLimit: inspector?.getAttribute('data-ppt-animation-limit-delay-max') ?? '',
      descriptorDurationLimit: inspector?.getAttribute('data-ppt-animation-limit-duration-max') ?? '',
      descriptorModel: inspector?.getAttribute('data-ppt-animation-model') ?? '',
      descriptorOrderLimit: inspector?.getAttribute('data-ppt-animation-limit-order-max') ?? '',
      descriptorPackageTrigger: inspector?.getAttribute('data-ppt-animation-package-trigger') ?? '',
      descriptorPackageType: inspector?.getAttribute('data-ppt-animation-package-type') ?? '',
      descriptorTriggerOptions: inspector?.getAttribute('data-ppt-animation-trigger-options') ?? '',
      descriptorTypeOptions: inspector?.getAttribute('data-ppt-animation-type-options') ?? '',
      duration: duration?.value ?? '',
      durationCommand: duration?.getAttribute('data-ppt-animation-command') ?? '',
      inspector: !!inspector,
      inspectorDelay: inspector?.getAttribute('data-ppt-animation-delay') ?? '',
      inspectorDuration: inspector?.getAttribute('data-ppt-animation-duration') ?? '',
      inspectorOrder: inspector?.getAttribute('data-ppt-animation-order') ?? '',
      inspectorTrigger: inspector?.getAttribute('data-ppt-animation-trigger') ?? '',
      inspectorType: inspector?.getAttribute('data-ppt-animation-type') ?? '',
      importCommandFields: stage?.getAttribute('data-ppt-object-animation-import-command-fields') ?? '',
      importCommands: stage?.getAttribute('data-ppt-object-animation-import-commands') ?? '',
      importDelay: stage?.getAttribute('data-ppt-object-animation-import-delay') ?? '',
      importDuration: stage?.getAttribute('data-ppt-object-animation-import-duration') ?? '',
      importFields: stage?.getAttribute('data-ppt-object-animation-import-fields') ?? '',
      importFormat: stage?.getAttribute('data-ppt-object-animation-import-format') ?? '',
      importJsonLength: Number(stage?.getAttribute('data-ppt-object-animation-import-json-length') ?? 0),
      importModel: stage?.getAttribute('data-ppt-object-animation-import-model') ?? '',
      importObjects: stage?.getAttribute('data-ppt-object-animation-import-objects') ?? '',
      importOrder: stage?.getAttribute('data-ppt-object-animation-import-order') ?? '',
      importSlide: stage?.getAttribute('data-ppt-object-animation-import-slide') ?? '',
      importTrigger: stage?.getAttribute('data-ppt-object-animation-import-trigger') ?? '',
      importType: stage?.getAttribute('data-ppt-object-animation-import-type') ?? '',
      model: stage?.getAttribute('data-ppt-object-animation-model') ?? '',
      order: order?.value ?? '',
      orderCommand: order?.getAttribute('data-ppt-animation-command') ?? '',
      selectedDelay: selected?.getAttribute('data-ppt-animation-delay') ?? '',
      selectedDuration: selected?.getAttribute('data-ppt-animation-duration') ?? '',
      selectedId: selected?.getAttribute('data-ppt-element') ?? '',
      selectedOrder: selected?.getAttribute('data-ppt-animation-order') ?? '',
      selectedTrigger: selected?.getAttribute('data-ppt-animation-trigger') ?? '',
      selectedType: selected?.getAttribute('data-ppt-animation-type') ?? '',
      trigger: trigger?.value ?? '',
      triggerCommand: trigger?.getAttribute('data-ppt-animation-command') ?? '',
      triggerPackageValue: trigger?.getAttribute('data-ppt-animation-package-value') ?? '',
      type: type?.value ?? '',
      typeCommand: type?.getAttribute('data-ppt-animation-command') ?? '',
      typePackageValue: type?.getAttribute('data-ppt-animation-package-value') ?? '',
    }
  })()`)
}

function getPPTObjectAnimationPreviewState(page, elementId) {
  return page.eval(`((id) => {
    const overlay = document.querySelector('[data-ppt-presentation]')
    const element = overlay?.querySelector(\`[data-ppt-element="\${id}"]\`)

    return {
      buildOrder: overlay?.getAttribute('data-ppt-presentation-animation-build-order') ?? '',
      buildOrderModel: overlay?.getAttribute('data-ppt-presentation-animation-build-order-model') ?? '',
      delay: element?.getAttribute('data-ppt-animation-delay') ?? '',
      duration: element?.getAttribute('data-ppt-animation-duration') ?? '',
      elementId: element?.getAttribute('data-ppt-element') ?? '',
      open: !!overlay,
      order: element?.getAttribute('data-ppt-animation-order') ?? '',
      trigger: element?.getAttribute('data-ppt-animation-trigger') ?? '',
      type: element?.getAttribute('data-ppt-animation-type') ?? '',
    }
  })(${JSON.stringify(elementId)})`)
}

function getPPTObjectOpacityState(page, elementId) {
  return page.eval(`((id) => {
    const selected = document.querySelector('[data-selected="true"]')
    const targetId = id || selected?.getAttribute('data-ppt-element') || ''
    const thumb = document.querySelector(\`.ppt-thumb[aria-current="page"] [data-ppt-thumb-element="\${targetId}"]\`)
    const field = document.querySelector('[data-ppt-style-field="opacity"]')
    const stage = document.querySelector('.ppt-stage-shell')

    return {
      command: stage?.getAttribute('data-ppt-object-opacity-command') ?? '',
      commandField: stage?.getAttribute('data-ppt-object-opacity-command-field') ?? '',
      commandObject: stage?.getAttribute('data-ppt-object-opacity-command-object') ?? '',
      commandSlide: stage?.getAttribute('data-ppt-object-opacity-command-slide') ?? '',
      commandType: stage?.getAttribute('data-ppt-object-opacity-command-type') ?? '',
      commandValue: stage?.getAttribute('data-ppt-object-opacity-command-value') ?? '',
      descriptorAttribute: field?.getAttribute('data-ppt-object-opacity-attribute') ?? '',
      descriptorAttributeValue: field?.getAttribute('data-ppt-object-opacity-attribute-value') ?? '',
      descriptorCommand: field?.getAttribute('data-ppt-object-opacity-command') ?? '',
      descriptorControl: field?.getAttribute('data-ppt-object-opacity-control') ?? '',
      descriptorSurface: field?.getAttribute('data-ppt-object-opacity-surface') ?? '',
      model: stage?.getAttribute('data-ppt-object-opacity-model') ?? '',
      opacityImportCommandFields: stage?.getAttribute('data-ppt-object-opacity-import-command-fields') ?? '',
      opacityImportCommandTargets: stage?.getAttribute('data-ppt-object-opacity-import-command-targets') ?? '',
      opacityImportCommandTypes: stage?.getAttribute('data-ppt-object-opacity-import-command-types') ?? '',
      opacityImportCommandValues: stage?.getAttribute('data-ppt-object-opacity-import-command-values') ?? '',
      opacityImportCommands: stage?.getAttribute('data-ppt-object-opacity-import-commands') ?? '',
      opacityImportFields: stage?.getAttribute('data-ppt-object-opacity-import-fields') ?? '',
      opacityImportFormat: stage?.getAttribute('data-ppt-object-opacity-import-format') ?? '',
      opacityImportJsonLength: Number(stage?.getAttribute('data-ppt-object-opacity-import-json-length') ?? 0),
      opacityImportModel: stage?.getAttribute('data-ppt-object-opacity-import-model') ?? '',
      opacityImportObjects: stage?.getAttribute('data-ppt-object-opacity-import-objects') ?? '',
      opacityImportSlide: stage?.getAttribute('data-ppt-object-opacity-import-slide') ?? '',
      opacityImportValue: stage?.getAttribute('data-ppt-object-opacity-import-value') ?? '',
      opacity: field?.value ?? '',
      selectedId: selected?.getAttribute('data-ppt-element') ?? '',
      selectedOpacity: selected?.getAttribute('data-ppt-opacity') ?? '',
      selectedStyleOpacity: selected?.style.opacity ?? '',
      thumbOpacity: thumb?.getAttribute('data-ppt-thumb-opacity') ?? '',
      thumbStyleOpacity: thumb?.style.opacity ?? '',
    }
  })(${JSON.stringify(elementId)})`)
}

function getPPTObjectShadowState(page, elementId) {
  return page.eval(`((id) => {
    const selected = document.querySelector('[data-selected="true"]')
    const targetId = id || selected?.getAttribute('data-ppt-element') || ''
    const thumb = document.querySelector(\`.ppt-thumb[aria-current="page"] [data-ppt-thumb-element="\${targetId}"]\`)
    const stage = document.querySelector('.ppt-stage-shell')
    const enabled = document.querySelector('[data-ppt-shadow-field="enabled"]')
    const color = document.querySelector('[data-ppt-shadow-field="color"]')
    const opacity = document.querySelector('[data-ppt-shadow-field="opacity"]')
    const blur = document.querySelector('[data-ppt-shadow-field="blur"]')
    const distance = document.querySelector('[data-ppt-shadow-field="distance"]')
    const angle = document.querySelector('[data-ppt-shadow-field="angle"]')
    const descriptorAttributeValue = enabled?.getAttribute('data-ppt-shadow-attribute-value') ?? ''
    let descriptor = null

    try {
      descriptor = descriptorAttributeValue && descriptorAttributeValue !== 'none'
        ? JSON.parse(descriptorAttributeValue)
        : null
    } catch {
      descriptor = null
    }

    return {
      angle: angle?.value ?? '',
      angleControl: angle?.getAttribute('data-ppt-shadow-control') ?? '',
      angleUnit: angle?.getAttribute('data-ppt-shadow-unit') ?? '',
      blur: blur?.value ?? '',
      blurControl: blur?.getAttribute('data-ppt-shadow-control') ?? '',
      blurUnit: blur?.getAttribute('data-ppt-shadow-unit') ?? '',
      color: color?.value ?? '',
      colorControl: color?.getAttribute('data-ppt-shadow-control') ?? '',
      command: stage?.getAttribute('data-ppt-shadow-command') ?? '',
      commandField: stage?.getAttribute('data-ppt-shadow-command-field') ?? '',
      commandObject: stage?.getAttribute('data-ppt-shadow-command-object') ?? '',
      commandSlide: stage?.getAttribute('data-ppt-shadow-command-slide') ?? '',
      commandType: stage?.getAttribute('data-ppt-shadow-command-type') ?? '',
      commandValue: stage?.getAttribute('data-ppt-shadow-command-value') ?? '',
      descriptorAngle: descriptor?.angle === undefined ? '' : String(descriptor.angle),
      descriptorAttribute: enabled?.getAttribute('data-ppt-shadow-attribute') ?? '',
      descriptorAttributeValue,
      descriptorBlur: descriptor?.blur === undefined ? '' : String(descriptor.blur),
      descriptorColor: descriptor?.color ?? '',
      descriptorCommand: enabled?.getAttribute('data-ppt-shadow-command') ?? '',
      descriptorControl: enabled?.getAttribute('data-ppt-shadow-control') ?? '',
      descriptorDistance: descriptor?.distance === undefined ? '' : String(descriptor.distance),
      descriptorEnabled: enabled?.getAttribute('data-ppt-shadow-descriptor-enabled') ?? '',
      descriptorOpacity: descriptor?.opacity === undefined ? '' : String(descriptor.opacity),
      descriptorSurface: enabled?.getAttribute('data-ppt-shadow-surface') ?? '',
      distance: distance?.value ?? '',
      distanceControl: distance?.getAttribute('data-ppt-shadow-control') ?? '',
      distanceUnit: distance?.getAttribute('data-ppt-shadow-unit') ?? '',
      enabled: enabled?.checked ?? false,
      inspectorEnabled: document.querySelector('[data-ppt-shadow-inspector]')?.getAttribute('data-ppt-shadow-enabled') ?? '',
      model: stage?.getAttribute('data-ppt-shadow-model') ?? '',
      objectStyleImportCategories: stage?.getAttribute('data-ppt-object-style-import-categories') ?? '',
      objectStyleImportCommand: stage?.getAttribute('data-ppt-object-style-import-command') ?? '',
      objectStyleImportCommandTargets: stage?.getAttribute('data-ppt-object-style-import-command-targets') ?? '',
      objectStyleImportCommandType: stage?.getAttribute('data-ppt-object-style-import-command-type') ?? '',
      objectStyleImportFields: stage?.getAttribute('data-ppt-object-style-import-fields') ?? '',
      objectStyleImportFormat: stage?.getAttribute('data-ppt-object-style-import-format') ?? '',
      objectStyleImportJsonLength: Number(stage?.getAttribute('data-ppt-object-style-import-json-length') ?? 0),
      objectStyleImportModel: stage?.getAttribute('data-ppt-object-style-import-model') ?? '',
      objectStyleImportObjects: stage?.getAttribute('data-ppt-object-style-import-objects') ?? '',
      objectStyleImportOpacity: stage?.getAttribute('data-ppt-object-style-import-opacity') ?? '',
      objectStyleImportShadowAngle: stage?.getAttribute('data-ppt-object-style-import-shadow-angle') ?? '',
      objectStyleImportShadowBlur: stage?.getAttribute('data-ppt-object-style-import-shadow-blur') ?? '',
      objectStyleImportShadowColor: stage?.getAttribute('data-ppt-object-style-import-shadow-color') ?? '',
      objectStyleImportShadowDistance: stage?.getAttribute('data-ppt-object-style-import-shadow-distance') ?? '',
      objectStyleImportShadowEnabled: stage?.getAttribute('data-ppt-object-style-import-shadow-enabled') ?? '',
      objectStyleImportShadowOpacity: stage?.getAttribute('data-ppt-object-style-import-shadow-opacity') ?? '',
      opacity: opacity?.value ?? '',
      opacityControl: opacity?.getAttribute('data-ppt-shadow-control') ?? '',
      opacityDisabled: opacity?.disabled ?? false,
      opacityUnit: opacity?.getAttribute('data-ppt-shadow-unit') ?? '',
      shadowImportAngle: stage?.getAttribute('data-ppt-shadow-import-angle') ?? '',
      shadowImportBlur: stage?.getAttribute('data-ppt-shadow-import-blur') ?? '',
      shadowImportColor: stage?.getAttribute('data-ppt-shadow-import-color') ?? '',
      shadowImportCommandFields: stage?.getAttribute('data-ppt-shadow-import-command-fields') ?? '',
      shadowImportCommandTargets: stage?.getAttribute('data-ppt-shadow-import-command-targets') ?? '',
      shadowImportCommandTypes: stage?.getAttribute('data-ppt-shadow-import-command-types') ?? '',
      shadowImportCommandValues: stage?.getAttribute('data-ppt-shadow-import-command-values') ?? '',
      shadowImportCommands: stage?.getAttribute('data-ppt-shadow-import-commands') ?? '',
      shadowImportDistance: stage?.getAttribute('data-ppt-shadow-import-distance') ?? '',
      shadowImportEnabled: stage?.getAttribute('data-ppt-shadow-import-enabled') ?? '',
      shadowImportFields: stage?.getAttribute('data-ppt-shadow-import-fields') ?? '',
      shadowImportFormat: stage?.getAttribute('data-ppt-shadow-import-format') ?? '',
      shadowImportJsonLength: Number(stage?.getAttribute('data-ppt-shadow-import-json-length') ?? 0),
      shadowImportModel: stage?.getAttribute('data-ppt-shadow-import-model') ?? '',
      shadowImportObjects: stage?.getAttribute('data-ppt-shadow-import-objects') ?? '',
      shadowImportOpacity: stage?.getAttribute('data-ppt-shadow-import-opacity') ?? '',
      shadowImportSlide: stage?.getAttribute('data-ppt-shadow-import-slide') ?? '',
      selectedAngle: selected?.getAttribute('data-ppt-shadow-angle') ?? '',
      selectedBlur: selected?.getAttribute('data-ppt-shadow-blur') ?? '',
      selectedColor: selected?.getAttribute('data-ppt-shadow-color') ?? '',
      selectedDistance: selected?.getAttribute('data-ppt-shadow-distance') ?? '',
      selectedFilter: selected?.style.filter ?? '',
      selectedId: selected?.getAttribute('data-ppt-element') ?? '',
      selectedObjectOpacity: selected?.getAttribute('data-ppt-opacity') ?? '',
      selectedOpacity: selected?.getAttribute('data-ppt-shadow-opacity') ?? '',
      selectedShadow: selected?.getAttribute('data-ppt-shadow') ?? '',
      thumbFilter: thumb?.style.filter ?? '',
      thumbOpacity: thumb?.getAttribute('data-ppt-thumb-shadow-opacity') ?? '',
      thumbShadow: thumb?.getAttribute('data-ppt-thumb-shadow') ?? '',
    }
  })(${JSON.stringify(elementId)})`)
}

function getPPTObjectHyperlinkState(page, elementId) {
  return page.eval(`((id) => {
    const selected = document.querySelector('[data-selected="true"]')
    const targetId = id || selected?.getAttribute('data-ppt-element') || ''
    const thumb = document.querySelector(\`.ppt-thumb[aria-current="page"] [data-ppt-thumb-element="\${targetId}"]\`)
    const field = document.querySelector('[data-ppt-style-field="hyperlink"]')
    const stage = document.querySelector('.ppt-stage-shell')
    const descriptorAttributeValue = field?.getAttribute('data-ppt-hyperlink-attribute-value') ?? ''
    let descriptor = null

    try {
      descriptor = descriptorAttributeValue && descriptorAttributeValue !== 'none'
        ? JSON.parse(descriptorAttributeValue)
        : null
    } catch {
      descriptor = null
    }

    return {
      accessibilityCommand: stage?.getAttribute('data-ppt-accessibility-command') ?? '',
      accessibilityCommandField: stage?.getAttribute('data-ppt-accessibility-command-field') ?? '',
      accessibilityCommandObject: stage?.getAttribute('data-ppt-accessibility-command-object') ?? '',
      accessibilityCommandSlide: stage?.getAttribute('data-ppt-accessibility-command-slide') ?? '',
      accessibilityCommandType: stage?.getAttribute('data-ppt-accessibility-command-type') ?? '',
      accessibilityCommandValue: stage?.getAttribute('data-ppt-accessibility-command-value') ?? '',
      command: stage?.getAttribute('data-ppt-hyperlink-command') ?? '',
      commandField: stage?.getAttribute('data-ppt-hyperlink-command-field') ?? '',
      commandObject: stage?.getAttribute('data-ppt-hyperlink-command-object') ?? '',
      commandSlide: stage?.getAttribute('data-ppt-hyperlink-command-slide') ?? '',
      commandType: stage?.getAttribute('data-ppt-hyperlink-command-type') ?? '',
      commandValue: stage?.getAttribute('data-ppt-hyperlink-command-value') ?? '',
      descriptorAttribute: field?.getAttribute('data-ppt-hyperlink-attribute') ?? '',
      descriptorAttributeValue,
      descriptorCommand: field?.getAttribute('data-ppt-hyperlink-command') ?? '',
      descriptorControl: field?.getAttribute('data-ppt-hyperlink-control') ?? '',
      descriptorEnabled: field?.getAttribute('data-ppt-hyperlink-enabled') ?? '',
      descriptorSurface: field?.getAttribute('data-ppt-hyperlink-surface') ?? '',
      descriptorTarget: descriptor?.target ?? '',
      descriptorUrl: descriptor?.url ?? '',
      descriptorValidation: field?.getAttribute('data-ppt-hyperlink-validation') ?? '',
      hyperlinkImportCommandFields: stage?.getAttribute('data-ppt-hyperlink-import-command-fields') ?? '',
      hyperlinkImportCommandTargets: stage?.getAttribute('data-ppt-hyperlink-import-command-targets') ?? '',
      hyperlinkImportCommandTypes: stage?.getAttribute('data-ppt-hyperlink-import-command-types') ?? '',
      hyperlinkImportCommands: stage?.getAttribute('data-ppt-hyperlink-import-commands') ?? '',
      hyperlinkImportEnabled: stage?.getAttribute('data-ppt-hyperlink-import-enabled') ?? '',
      hyperlinkImportFields: stage?.getAttribute('data-ppt-hyperlink-import-fields') ?? '',
      hyperlinkImportFormat: stage?.getAttribute('data-ppt-hyperlink-import-format') ?? '',
      hyperlinkImportJsonLength: Number(stage?.getAttribute('data-ppt-hyperlink-import-json-length') ?? 0),
      hyperlinkImportModel: stage?.getAttribute('data-ppt-hyperlink-import-model') ?? '',
      hyperlinkImportObjects: stage?.getAttribute('data-ppt-hyperlink-import-objects') ?? '',
      hyperlinkImportSlide: stage?.getAttribute('data-ppt-hyperlink-import-slide') ?? '',
      hyperlinkImportUrl: stage?.getAttribute('data-ppt-hyperlink-import-url') ?? '',
      model: stage?.getAttribute('data-ppt-hyperlink-model') ?? '',
      objectMetadataImportAltTextLength: Number(stage?.getAttribute('data-ppt-object-metadata-import-alt-text-length') ?? 0),
      objectMetadataImportAltTextPresent: stage?.getAttribute('data-ppt-object-metadata-import-alt-text-present') ?? '',
      objectMetadataImportCommandFields: stage?.getAttribute('data-ppt-object-metadata-import-command-fields') ?? '',
      objectMetadataImportCommandTypes: stage?.getAttribute('data-ppt-object-metadata-import-command-types') ?? '',
      objectMetadataImportCommands: stage?.getAttribute('data-ppt-object-metadata-import-commands') ?? '',
      objectMetadataImportFields: stage?.getAttribute('data-ppt-object-metadata-import-fields') ?? '',
      objectMetadataImportFormat: stage?.getAttribute('data-ppt-object-metadata-import-format') ?? '',
      objectMetadataImportHyperlinkUrl: stage?.getAttribute('data-ppt-object-metadata-import-hyperlink-url') ?? '',
      objectMetadataImportJsonLength: Number(stage?.getAttribute('data-ppt-object-metadata-import-json-length') ?? 0),
      objectMetadataImportModel: stage?.getAttribute('data-ppt-object-metadata-import-model') ?? '',
      objectMetadataImportName: stage?.getAttribute('data-ppt-object-metadata-import-name') ?? '',
      objectMetadataImportObjects: stage?.getAttribute('data-ppt-object-metadata-import-objects') ?? '',
      objectMetadataImportSlide: stage?.getAttribute('data-ppt-object-metadata-import-slide') ?? '',
      rowName: targetId
        ? document.querySelector(\`[data-ppt-layer-row="\${targetId}"] .ppt-layer-name\`)?.textContent ?? ''
        : '',
      selectedAltText: selected?.getAttribute('data-ppt-alt-text') ?? '',
      selectedId: selected?.getAttribute('data-ppt-element') ?? '',
      selectedName: selected?.getAttribute('data-ppt-element-name') ?? '',
      selectedUrl: selected?.getAttribute('data-ppt-hyperlink-url') ?? '',
      thumbAltText: thumb?.getAttribute('data-ppt-thumb-alt-text') ?? '',
      thumbUrl: thumb?.getAttribute('data-ppt-thumb-hyperlink-url') ?? '',
      url: field?.value ?? '',
    }
  })(${JSON.stringify(elementId)})`)
}

function getPPTObjectAltTextState(page, elementId) {
  return page.eval(`((id) => {
    const selected = document.querySelector('[data-selected="true"]')
    const targetId = id || selected?.getAttribute('data-ppt-element') || ''
    const target = targetId
      ? document.querySelector(\`[data-ppt-element="\${targetId}"]\`)
      : selected
    const thumb = document.querySelector(\`.ppt-thumb[aria-current="page"] [data-ppt-thumb-element="\${targetId}"]\`)
    const field = document.querySelector('[data-ppt-style-field="alt-text"]')
    const image = target?.querySelector('img') ?? null
    const stage = document.querySelector('.ppt-stage-shell')
    const descriptorAttributeValue = field?.getAttribute('data-ppt-accessibility-attribute-value') ?? ''
    let descriptor = null

    try {
      descriptor = descriptorAttributeValue && descriptorAttributeValue !== 'none'
        ? JSON.parse(descriptorAttributeValue)
        : null
    } catch {
      descriptor = null
    }

    return {
      altText: field?.value ?? '',
      command: stage?.getAttribute('data-ppt-accessibility-command') ?? '',
      commandField: stage?.getAttribute('data-ppt-accessibility-command-field') ?? '',
      commandObject: stage?.getAttribute('data-ppt-accessibility-command-object') ?? '',
      commandSlide: stage?.getAttribute('data-ppt-accessibility-command-slide') ?? '',
      commandType: stage?.getAttribute('data-ppt-accessibility-command-type') ?? '',
      commandValue: stage?.getAttribute('data-ppt-accessibility-command-value') ?? '',
      descriptorAltText: descriptor?.altText ?? '',
      descriptorAttribute: field?.getAttribute('data-ppt-accessibility-attribute') ?? '',
      descriptorAttributeValue,
      descriptorCommand: field?.getAttribute('data-ppt-accessibility-command') ?? '',
      descriptorControl: field?.getAttribute('data-ppt-accessibility-control') ?? '',
      descriptorDescribed: field?.getAttribute('data-ppt-accessibility-described') ?? '',
      descriptorSurface: field?.getAttribute('data-ppt-accessibility-surface') ?? '',
      imageAlt: image?.getAttribute('alt') ?? '',
      importAltTextLength: Number(stage?.getAttribute('data-ppt-accessibility-import-alt-text-length') ?? 0),
      importAltTextPresent: stage?.getAttribute('data-ppt-accessibility-import-alt-text-present') ?? '',
      importCommandFields: stage?.getAttribute('data-ppt-accessibility-import-command-fields') ?? '',
      importCommandTargets: stage?.getAttribute('data-ppt-accessibility-import-command-targets') ?? '',
      importCommandTypes: stage?.getAttribute('data-ppt-accessibility-import-command-types') ?? '',
      importCommandValues: stage?.getAttribute('data-ppt-accessibility-import-command-values') ?? '',
      importCommands: stage?.getAttribute('data-ppt-accessibility-import-commands') ?? '',
      importFields: stage?.getAttribute('data-ppt-accessibility-import-fields') ?? '',
      importFormat: stage?.getAttribute('data-ppt-accessibility-import-format') ?? '',
      importJsonLength: Number(stage?.getAttribute('data-ppt-accessibility-import-json-length') ?? 0),
      importModel: stage?.getAttribute('data-ppt-accessibility-import-model') ?? '',
      importObjects: stage?.getAttribute('data-ppt-accessibility-import-objects') ?? '',
      importSlide: stage?.getAttribute('data-ppt-accessibility-import-slide') ?? '',
      model: stage?.getAttribute('data-ppt-accessibility-model') ?? '',
      selectedAltText: target?.getAttribute('data-ppt-alt-text') ?? '',
      selectedId: selected?.getAttribute('data-ppt-element') ?? '',
      selectedKind: selected?.getAttribute('data-kind') ?? '',
      thumbAltText: thumb?.getAttribute('data-ppt-thumb-alt-text') ?? '',
    }
  })(${JSON.stringify(elementId)})`)
}

function getPPTTextParagraphSpacingState(page) {
  return page.eval(`(() => {
    const inspector = document.querySelector('[data-ppt-paragraph-spacing-inspector]')
    const lineHeight = document.querySelector('[data-ppt-paragraph-field="lineHeight"]')
    const selected = document.querySelector('[data-selected="true"]')
    const paragraph = selected?.querySelector('.ppt-text-paragraph')
    const spacingAfter = document.querySelector('[data-ppt-paragraph-field="spacingAfter"]')
    const spacingBefore = document.querySelector('[data-ppt-paragraph-field="spacingBefore"]')
    const stage = document.querySelector('.ppt-stage-shell')

    return {
      command: stage?.getAttribute('data-ppt-text-paragraph-spacing-command') ?? '',
      commandField: stage?.getAttribute('data-ppt-text-paragraph-spacing-command-field') ?? '',
      commandObject: stage?.getAttribute('data-ppt-text-paragraph-spacing-command-object') ?? '',
      commandSlide: stage?.getAttribute('data-ppt-text-paragraph-spacing-command-slide') ?? '',
      commandType: stage?.getAttribute('data-ppt-text-paragraph-spacing-command-type') ?? '',
      commandUnit: stage?.getAttribute('data-ppt-text-paragraph-spacing-command-unit') ?? '',
      commandValue: stage?.getAttribute('data-ppt-text-paragraph-spacing-command-value') ?? '',
      descriptorSurface: inspector?.getAttribute('data-ppt-paragraph-spacing-surface') ?? '',
      inspector: !!inspector,
      inspectorLineHeight: inspector?.getAttribute('data-ppt-paragraph-line-height') ?? '',
      inspectorSpacingAfter: inspector?.getAttribute('data-ppt-paragraph-spacing-after') ?? '',
      inspectorSpacingBefore: inspector?.getAttribute('data-ppt-paragraph-spacing-before') ?? '',
      importCommandFields: stage?.getAttribute('data-ppt-text-paragraph-spacing-import-command-fields') ?? '',
      importCommandTargets: stage?.getAttribute('data-ppt-text-paragraph-spacing-import-command-targets') ?? '',
      importCommandTypes: stage?.getAttribute('data-ppt-text-paragraph-spacing-import-command-types') ?? '',
      importCommandUnits: stage?.getAttribute('data-ppt-text-paragraph-spacing-import-command-units') ?? '',
      importCommandValues: stage?.getAttribute('data-ppt-text-paragraph-spacing-import-command-values') ?? '',
      importCommands: stage?.getAttribute('data-ppt-text-paragraph-spacing-import-commands') ?? '',
      importFields: stage?.getAttribute('data-ppt-text-paragraph-spacing-import-fields') ?? '',
      importFormat: stage?.getAttribute('data-ppt-text-paragraph-spacing-import-format') ?? '',
      importJsonLength: Number(stage?.getAttribute('data-ppt-text-paragraph-spacing-import-json-length') ?? '0'),
      importLineHeight: stage?.getAttribute('data-ppt-text-paragraph-spacing-import-line-height') ?? '',
      importModel: stage?.getAttribute('data-ppt-text-paragraph-spacing-import-model') ?? '',
      importObjects: stage?.getAttribute('data-ppt-text-paragraph-spacing-import-objects') ?? '',
      importSpacingAfter: stage?.getAttribute('data-ppt-text-paragraph-spacing-import-spacing-after') ?? '',
      importSpacingBefore: stage?.getAttribute('data-ppt-text-paragraph-spacing-import-spacing-before') ?? '',
      lineHeight: lineHeight?.value ?? '',
      lineHeightCommand: lineHeight?.getAttribute('data-ppt-paragraph-command') ?? '',
      lineHeightControl: lineHeight?.getAttribute('data-ppt-paragraph-control') ?? '',
      model: stage?.getAttribute('data-ppt-text-paragraph-spacing-model') ?? '',
      selectedId: selected?.getAttribute('data-ppt-element') ?? '',
      selectedLineHeight: paragraph?.getAttribute('data-ppt-line-height') ?? '',
      selectedSpacingAfter: paragraph?.getAttribute('data-ppt-spacing-after') ?? '',
      selectedSpacingBefore: paragraph?.getAttribute('data-ppt-spacing-before') ?? '',
      selectedStyleLineHeight: paragraph?.style.lineHeight ?? '',
      selectedStyleMarginBottom: paragraph?.style.marginBottom ?? '',
      selectedStyleMarginTop: paragraph?.style.marginTop ?? '',
      spacingAfter: spacingAfter?.value ?? '',
      spacingAfterCommand: spacingAfter?.getAttribute('data-ppt-paragraph-command') ?? '',
      spacingAfterControl: spacingAfter?.getAttribute('data-ppt-paragraph-control') ?? '',
      spacingAfterUnit: spacingAfter?.getAttribute('data-ppt-paragraph-unit') ?? '',
      spacingBefore: spacingBefore?.value ?? '',
      spacingBeforeCommand: spacingBefore?.getAttribute('data-ppt-paragraph-command') ?? '',
      spacingBeforeControl: spacingBefore?.getAttribute('data-ppt-paragraph-control') ?? '',
      spacingBeforeUnit: spacingBefore?.getAttribute('data-ppt-paragraph-unit') ?? '',
    }
  })()`)
}

function getPPTTextFontFamilyState(page) {
  return page.eval(`(() => {
    const field = document.querySelector('[data-ppt-style-field="font-family"]')
    const selected = document.querySelector('[data-selected="true"]')
    const stage = document.querySelector('.ppt-stage-shell')
    const thumb = document.querySelector('.ppt-thumb[aria-current="page"] [data-ppt-thumb-element="s1-title"]')

    return {
      command: stage?.getAttribute('data-ppt-text-font-family-command') ?? '',
      commandField: stage?.getAttribute('data-ppt-text-font-family-command-field') ?? '',
      commandObject: stage?.getAttribute('data-ppt-text-font-family-command-object') ?? '',
      commandSlide: stage?.getAttribute('data-ppt-text-font-family-command-slide') ?? '',
      commandType: stage?.getAttribute('data-ppt-text-font-family-command-type') ?? '',
      commandValue: stage?.getAttribute('data-ppt-text-font-family-command-value') ?? '',
      descriptorCommand: field?.getAttribute('data-ppt-text-font-family-command') ?? '',
      descriptorControl: field?.getAttribute('data-ppt-text-font-family-control') ?? '',
      descriptorFallback: field?.getAttribute('data-ppt-text-font-family-fallback') ?? '',
      descriptorOptions: field?.getAttribute('data-ppt-text-font-family-options') ?? '',
      descriptorSurface: field?.getAttribute('data-ppt-text-font-family-surface') ?? '',
      fontFamily: field?.value ?? '',
      importCommandFields: stage?.getAttribute('data-ppt-text-font-family-import-command-fields') ?? '',
      importCommandTargets: stage?.getAttribute('data-ppt-text-font-family-import-command-targets') ?? '',
      importCommandTypes: stage?.getAttribute('data-ppt-text-font-family-import-command-types') ?? '',
      importCommandValues: stage?.getAttribute('data-ppt-text-font-family-import-command-values') ?? '',
      importCommands: stage?.getAttribute('data-ppt-text-font-family-import-commands') ?? '',
      importFields: stage?.getAttribute('data-ppt-text-font-family-import-fields') ?? '',
      importFormat: stage?.getAttribute('data-ppt-text-font-family-import-format') ?? '',
      importJsonLength: Number(stage?.getAttribute('data-ppt-text-font-family-import-json-length') ?? 0),
      importModel: stage?.getAttribute('data-ppt-text-font-family-import-model') ?? '',
      importObjects: stage?.getAttribute('data-ppt-text-font-family-import-objects') ?? '',
      importSlide: stage?.getAttribute('data-ppt-text-font-family-import-slide') ?? '',
      importValue: stage?.getAttribute('data-ppt-text-font-family-import-value') ?? '',
      model: stage?.getAttribute('data-ppt-text-font-family-model') ?? '',
      selectedFontFamily: selected?.getAttribute('data-ppt-font-family') ?? '',
      selectedId: selected?.getAttribute('data-ppt-element') ?? '',
      selectedStyleFontFamily: selected?.style.fontFamily ?? '',
      thumbFontFamily: thumb?.getAttribute('data-ppt-thumb-font-family') ?? '',
      thumbStyleFontFamily: thumb?.style.fontFamily ?? '',
    }
  })()`)
}

function getPPTTextVerticalAlignState(page) {
  return page.eval(`(() => {
    const selected = document.querySelector('[data-selected="true"]')
    const field = document.querySelector('[data-ppt-style-field="vertical-align"]')
    const stage = document.querySelector('.ppt-stage-shell')
    const thumb = document.querySelector('.ppt-thumb[aria-current="page"] [data-ppt-thumb-element="s1-title"]')

    return {
      command: stage?.getAttribute('data-ppt-text-vertical-align-command') ?? '',
      commandField: stage?.getAttribute('data-ppt-text-vertical-align-command-field') ?? '',
      commandObject: stage?.getAttribute('data-ppt-text-vertical-align-command-object') ?? '',
      commandSlide: stage?.getAttribute('data-ppt-text-vertical-align-command-slide') ?? '',
      commandType: stage?.getAttribute('data-ppt-text-vertical-align-command-type') ?? '',
      commandValue: stage?.getAttribute('data-ppt-text-vertical-align-command-value') ?? '',
      descriptorAttribute: field?.getAttribute('data-ppt-text-vertical-align-attribute') ?? '',
      descriptorAttributeValue: field?.getAttribute('data-ppt-text-vertical-align-attribute-value') ?? '',
      descriptorCommand: field?.getAttribute('data-ppt-text-vertical-align-command') ?? '',
      descriptorControl: field?.getAttribute('data-ppt-text-vertical-align-control') ?? '',
      descriptorDefaultValue: field?.getAttribute('data-ppt-text-vertical-align-default-value') ?? '',
      descriptorOptions: field?.getAttribute('data-ppt-text-vertical-align-options') ?? '',
      descriptorSurface: field?.getAttribute('data-ppt-text-vertical-align-surface') ?? '',
      importAlignItems: stage?.getAttribute('data-ppt-text-vertical-align-import-align-items') ?? '',
      importCommandFields: stage?.getAttribute('data-ppt-text-vertical-align-import-command-fields') ?? '',
      importCommandTargets: stage?.getAttribute('data-ppt-text-vertical-align-import-command-targets') ?? '',
      importCommandTypes: stage?.getAttribute('data-ppt-text-vertical-align-import-command-types') ?? '',
      importCommands: stage?.getAttribute('data-ppt-text-vertical-align-import-commands') ?? '',
      importFields: stage?.getAttribute('data-ppt-text-vertical-align-import-fields') ?? '',
      importFormat: stage?.getAttribute('data-ppt-text-vertical-align-import-format') ?? '',
      importJsonLength: Number(stage?.getAttribute('data-ppt-text-vertical-align-import-json-length') ?? '0'),
      importModel: stage?.getAttribute('data-ppt-text-vertical-align-import-model') ?? '',
      importObjects: stage?.getAttribute('data-ppt-text-vertical-align-import-objects') ?? '',
      importValue: stage?.getAttribute('data-ppt-text-vertical-align-import-value') ?? '',
      model: stage?.getAttribute('data-ppt-text-vertical-align-model') ?? '',
      selectedId: selected?.getAttribute('data-ppt-element') ?? '',
      selectedStyleAlignItems: selected?.style.alignItems ?? '',
      selectedVerticalAlign: selected?.getAttribute('data-ppt-vertical-align') ?? '',
      thumbStyleAlignItems: thumb?.style.alignItems ?? '',
      thumbVerticalAlign: thumb?.getAttribute('data-ppt-thumb-vertical-align') ?? '',
      verticalAlign: field?.value ?? '',
    }
  })()`)
}

function getPPTTextFrameInsetState(page) {
  return page.eval(`(() => {
    const inspector = document.querySelector('[data-ppt-text-inset-inspector]')
    const selected = document.querySelector('[data-selected="true"]')
    const stage = document.querySelector('.ppt-stage-shell')
    const thumb = document.querySelector('.ppt-thumb[aria-current="page"] [data-ppt-thumb-element="s1-title"]')
    const top = document.querySelector('[data-ppt-text-inset-field="top"]')
    const right = document.querySelector('[data-ppt-text-inset-field="right"]')
    const bottom = document.querySelector('[data-ppt-text-inset-field="bottom"]')
    const left = document.querySelector('[data-ppt-text-inset-field="left"]')

    return {
      bottom: bottom?.value ?? '',
      bottomControl: bottom?.getAttribute('data-ppt-text-inset-control') ?? '',
      bottomUnit: bottom?.getAttribute('data-ppt-text-inset-unit') ?? '',
      command: stage?.getAttribute('data-ppt-text-inset-command') ?? '',
      commandField: stage?.getAttribute('data-ppt-text-inset-command-field') ?? '',
      commandObject: stage?.getAttribute('data-ppt-text-inset-command-object') ?? '',
      commandSlide: stage?.getAttribute('data-ppt-text-inset-command-slide') ?? '',
      commandType: stage?.getAttribute('data-ppt-text-inset-command-type') ?? '',
      commandValue: stage?.getAttribute('data-ppt-text-inset-command-value') ?? '',
      descriptorAttribute: inspector?.getAttribute('data-ppt-text-inset-attribute') ?? '',
      descriptorAttributeValue: inspector?.getAttribute('data-ppt-text-inset-attribute-value') ?? '',
      descriptorDefaultValue: inspector?.getAttribute('data-ppt-text-inset-default-value') ?? '',
      descriptorSurface: inspector?.getAttribute('data-ppt-text-inset-surface') ?? '',
      inspectorTextInset: [
        inspector?.getAttribute('data-ppt-text-inset-top') ?? '',
        inspector?.getAttribute('data-ppt-text-inset-right') ?? '',
        inspector?.getAttribute('data-ppt-text-inset-bottom') ?? '',
        inspector?.getAttribute('data-ppt-text-inset-left') ?? '',
      ].join(','),
      importBottom: stage?.getAttribute('data-ppt-text-inset-import-bottom') ?? '',
      importCommandFields: stage?.getAttribute('data-ppt-text-inset-import-command-fields') ?? '',
      importCommandTargets: stage?.getAttribute('data-ppt-text-inset-import-command-targets') ?? '',
      importCommandTypes: stage?.getAttribute('data-ppt-text-inset-import-command-types') ?? '',
      importCommands: stage?.getAttribute('data-ppt-text-inset-import-commands') ?? '',
      importFields: stage?.getAttribute('data-ppt-text-inset-import-fields') ?? '',
      importFormat: stage?.getAttribute('data-ppt-text-inset-import-format') ?? '',
      importJsonLength: Number(stage?.getAttribute('data-ppt-text-inset-import-json-length') ?? '0'),
      importLeft: stage?.getAttribute('data-ppt-text-inset-import-left') ?? '',
      importModel: stage?.getAttribute('data-ppt-text-inset-import-model') ?? '',
      importObjects: stage?.getAttribute('data-ppt-text-inset-import-objects') ?? '',
      importRight: stage?.getAttribute('data-ppt-text-inset-import-right') ?? '',
      importTop: stage?.getAttribute('data-ppt-text-inset-import-top') ?? '',
      left: left?.value ?? '',
      leftCommand: left?.getAttribute('data-ppt-text-inset-command') ?? '',
      leftControl: left?.getAttribute('data-ppt-text-inset-control') ?? '',
      leftUnit: left?.getAttribute('data-ppt-text-inset-unit') ?? '',
      model: stage?.getAttribute('data-ppt-text-inset-model') ?? '',
      right: right?.value ?? '',
      rightControl: right?.getAttribute('data-ppt-text-inset-control') ?? '',
      rightUnit: right?.getAttribute('data-ppt-text-inset-unit') ?? '',
      selectedId: selected?.getAttribute('data-ppt-element') ?? '',
      selectedStylePadding: selected?.style.padding ?? '',
      selectedTextInset: selected?.getAttribute('data-ppt-text-inset') ?? '',
      thumbTextInset: thumb?.getAttribute('data-ppt-thumb-text-inset') ?? '',
      top: top?.value ?? '',
      topControl: top?.getAttribute('data-ppt-text-inset-control') ?? '',
      topUnit: top?.getAttribute('data-ppt-text-inset-unit') ?? '',
    }
  })()`)
}

async function clickMouse(page, x, y, clickCount, modifiers = 0) {
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount,
    modifiers,
    type: 'mousePressed',
    x,
    y,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount,
    modifiers,
    type: 'mouseReleased',
    x,
    y,
  })
}

async function dragMouse(page, points, modifiers = 0) {
  const [start, ...rest] = points
  const end = rest.at(-1) ?? start

  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    modifiers,
    type: 'mousePressed',
    x: start.x,
    y: start.y,
  })

  for (const point of rest) {
    await page.send('Input.dispatchMouseEvent', {
      button: 'left',
      modifiers,
      type: 'mouseMoved',
      x: point.x,
      y: point.y,
    })
  }

  await page.send('Input.dispatchMouseEvent', {
    button: 'left',
    clickCount: 1,
    modifiers,
    type: 'mouseReleased',
    x: end.x,
    y: end.y,
  })
}

async function installPPTDownloadCapture(page) {
  await page.eval(`(() => {
    window.__pptDownloads = []
    let downloadIndex = 0

    URL.createObjectURL = (blob) => {
      const url = \`blob:ppt-download-\${downloadIndex++}\`
      const entry = {
        download: '',
        text: '',
        type: blob.type,
        url,
      }

      window.__pptDownloads.push(entry)
      void blob.text().then((text) => {
        entry.text = text
      })

      return url
    }
    URL.revokeObjectURL = () => {}
    HTMLAnchorElement.prototype.click = function click() {
      const entry = window.__pptDownloads
        .find((download) => download.url === this.href)

      if (entry) {
        entry.download = this.download
      }
    }
  })()`)
}

async function rightClickMouse(page, x, y) {
  await page.send('Input.dispatchMouseEvent', {
    button: 'right',
    clickCount: 1,
    type: 'mousePressed',
    x,
    y,
  })
  await page.send('Input.dispatchMouseEvent', {
    button: 'right',
    clickCount: 1,
    type: 'mouseReleased',
    x,
    y,
  })
}

async function setTextQuickColor(page, color) {
  await page.eval(`((color) => {
    const input = document.querySelector('[data-ppt-text-quick="color"]')
    const valueSetter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set

    valueSetter.call(input, color)
    input.dispatchEvent(new Event('input', { bubbles: true }))
    input.dispatchEvent(new Event('change', { bubbles: true }))
  })(${JSON.stringify(color)})`)
}

async function readPPTParagraphAlignRadioGroupState(page) {
  return page.eval(`(() => {
    const readGroup = (surface) => {
      const group = document.querySelector(\`[data-ppt-paragraph-align-radiogroup="\${surface}"]\`)
      const radios = group ? [...group.querySelectorAll('[role="radio"]')] : []

      return {
        checkedValues: radios
          .filter((radio) => radio.getAttribute('aria-checked') === 'true')
          .map((radio) => radio.getAttribute('data-ppt-paragraph-align') ?? ''),
        focusModel: group?.getAttribute('data-ppt-paragraph-align-focus-model') ?? '',
        focusedValue: document.activeElement?.closest(\`[data-ppt-paragraph-align-radiogroup="\${surface}"]\`)
          ? document.activeElement?.getAttribute('data-ppt-paragraph-align') ?? ''
          : '',
        keyboardModel: group?.getAttribute('data-ppt-paragraph-align-keyboard-model') ?? '',
        model: group?.getAttribute('data-ppt-paragraph-align-model') ?? '',
        radioCount: radios.length,
        role: group?.getAttribute('role') ?? '',
        tabStopValues: radios
          .filter((radio) => radio.tabIndex === 0)
          .map((radio) => radio.getAttribute('data-ppt-paragraph-align') ?? ''),
      }
    }

    return {
      inspector: readGroup('inspector'),
      quick: readGroup('quick'),
      selectedTextAlign: document.querySelector('[data-selected="true"]')?.style.textAlign ?? '',
    }
  })()`)
}

function getElementCenter(page, elementId) {
  return page.eval(`(() => {
    const rect = document.querySelector('[data-ppt-element="${elementId}"]').getBoundingClientRect()

    return {
      x: rect.left + rect.width / 2,
      y: rect.top + rect.height / 2,
    }
  })()`)
}

function positionsChanged(before, after, ids, tolerance = 0.5) {
  return ids.some((id) =>
    Math.abs((before[id]?.x ?? 0) - (after[id]?.x ?? 0)) > tolerance ||
    Math.abs((before[id]?.y ?? 0) - (after[id]?.y ?? 0)) > tolerance)
}

function positionsMatch(actual, expected, ids, tolerance = 0.5) {
  return ids.every((id) =>
    Math.abs((actual[id]?.x ?? 0) - (expected[id]?.x ?? 0)) <= tolerance &&
    Math.abs((actual[id]?.y ?? 0) - (expected[id]?.y ?? 0)) <= tolerance)
}

function nearlyEqual(left, right, tolerance = 0.75) {
  return Math.abs(left - right) <= tolerance
}

function parsePPTPoint(value) {
  const [x = 0, y = 0] = value.split(',').map((part) => Number(part))

  return { x, y }
}

function selectEditableContents(page, elementId) {
  return page.eval(`(() => {
    const editor = document.querySelector('[data-ppt-element="${elementId}"] .ppt-element-editor')
    const range = document.createRange()
    const selection = window.getSelection()

    range.selectNodeContents(editor)
    selection.removeAllRanges()
    selection.addRange(range)
  })()`)
}

function getSlideRailState(page) {
  return page.eval(`(() => {
    const rail = document.querySelector('[data-ppt-slide-list]')
    const stage = document.querySelector('.ppt-stage-shell')
    const selectedIds = [...document.querySelectorAll('[data-selected="true"]')]
      .map((element) => element.getAttribute('data-ppt-element'))
      .filter(Boolean)
    const thumbs = [...document.querySelectorAll('.ppt-thumb')]
    const activeIndex = thumbs.findIndex((thumb) => thumb.getAttribute('aria-current') === 'page')
    const activeThumb = thumbs[activeIndex] ?? null
    const focusedThumb = document.activeElement?.closest('.ppt-thumb')

    return {
      activeId: activeThumb?.getAttribute('data-ppt-slide-id') ?? '',
      activeAttr: rail?.getAttribute('data-ppt-slide-rail-active') ?? '',
      activeHitH: activeThumb?.getAttribute('data-ppt-slide-rail-hit-h') ?? '',
      activeHitW: activeThumb?.getAttribute('data-ppt-slide-rail-hit-w') ?? '',
      activeHitX: activeThumb?.getAttribute('data-ppt-slide-rail-hit-x') ?? '',
      activeHitY: activeThumb?.getAttribute('data-ppt-slide-rail-hit-y') ?? '',
      activeIndex,
      activeName: activeThumb?.querySelector('.ppt-thumb-name')?.textContent ?? '',
      activeOption: rail?.getAttribute('data-ppt-slide-rail-active-option') ?? '',
      activeThumbH: activeThumb?.getAttribute('data-ppt-slide-rail-thumb-h') ?? '',
      activeThumbW: activeThumb?.getAttribute('data-ppt-slide-rail-thumb-w') ?? '',
      activeThumbX: activeThumb?.getAttribute('data-ppt-slide-rail-thumb-x') ?? '',
      activeThumbY: activeThumb?.getAttribute('data-ppt-slide-rail-thumb-y') ?? '',
      command: rail?.getAttribute('data-ppt-slide-rail-command') ?? '',
      commandFromIndex: rail?.getAttribute('data-ppt-slide-rail-command-from-index') ?? '',
      commandSelectionSlide: rail?.getAttribute('data-ppt-slide-rail-command-selection-slide') ?? '',
      commandSlide: rail?.getAttribute('data-ppt-slide-rail-command-slide') ?? '',
      commandToIndex: rail?.getAttribute('data-ppt-slide-rail-command-to-index') ?? '',
      commandType: rail?.getAttribute('data-ppt-slide-rail-command-type') ?? '',
      count: thumbs.length,
      draggableCount: thumbs.filter((thumb) => thumb.getAttribute('data-ppt-slide-draggable') === 'true').length,
      focusableOption: rail?.getAttribute('data-ppt-slide-rail-focusable-option') ?? '',
      focusedId: focusedThumb?.getAttribute('data-ppt-slide-id') ?? '',
      ids: thumbs.map((thumb) => thumb.getAttribute('data-ppt-slide-id') ?? ''),
      keyboardKeys: rail?.getAttribute('data-ppt-slide-rail-keyboard-keys') ?? '',
      keyboardModel: rail?.getAttribute('data-ppt-slide-rail-keyboard-model') ?? '',
      listRole: rail?.getAttribute('role') ?? '',
      model: rail?.getAttribute('data-ppt-slide-rail-model') ?? '',
      names: thumbs.map((thumb) => thumb.querySelector('.ppt-thumb-name')?.textContent ?? ''),
      optionCountAttr: rail?.getAttribute('data-ppt-slide-rail-option-count') ?? '',
      optionCount: thumbs.filter((thumb) => thumb.getAttribute('role') === 'option').length,
      optionFocusableIds: thumbs
        .filter((thumb) => thumb.getAttribute('data-ppt-slide-rail-option-focusable') === 'true')
        .map((thumb) => thumb.getAttribute('data-ppt-slide-id') ?? ''),
      optionIds: thumbs.map((thumb) => thumb.getAttribute('data-ppt-slide-rail-option-id') ?? ''),
      optionIndexes: thumbs.map((thumb) => thumb.getAttribute('data-ppt-slide-rail-option-index') ?? ''),
      rovingTabIndexes: thumbs.map((thumb) => thumb.getAttribute('data-ppt-slide-roving-tab-index') ?? ''),
      selectionMode: rail?.getAttribute('data-ppt-slide-rail-selection-mode') ?? '',
      selectedOptionIds: thumbs
        .filter((thumb) => thumb.getAttribute('aria-selected') === 'true')
        .map((thumb) => thumb.getAttribute('data-ppt-slide-id') ?? ''),
      selectedIds: selectedIds.join(','),
      slideClipboardElementCount: Number(stage?.getAttribute('data-ppt-slide-clipboard-element-count') ?? 0),
      slideClipboardHTMLLength: Number(stage?.getAttribute('data-ppt-slide-clipboard-html-length') ?? 0),
      slideClipboardImportFormat: stage?.getAttribute('data-ppt-slide-clipboard-import-format') ?? '',
      slideClipboardImported: stage?.getAttribute('data-ppt-slide-clipboard-imported') ?? '',
      slideClipboardJsonMimeType: stage?.getAttribute('data-ppt-slide-clipboard-json-mime-type') ?? '',
      slideClipboardModel: stage?.getAttribute('data-ppt-slide-clipboard-model') ?? '',
      slideClipboardSlideName: stage?.getAttribute('data-ppt-slide-clipboard-slide-name') ?? '',
      slideClipboardSourceSlide: stage?.getAttribute('data-ppt-slide-clipboard-source-slide') ?? '',
      slideClipboardTargetSlide: stage?.getAttribute('data-ppt-slide-clipboard-target-slide') ?? '',
      slideClipboardWriteMode: stage?.getAttribute('data-ppt-slide-clipboard-write-mode') ?? '',
      slideOrder: rail?.getAttribute('data-ppt-slide-rail-slide-order') ?? '',
      tabStopIds: thumbs
        .filter((thumb) => thumb.tabIndex === 0)
        .map((thumb) => thumb.getAttribute('data-ppt-slide-id') ?? ''),
      thumbnailCount: rail?.getAttribute('data-ppt-slide-rail-thumbnail-count') ?? '',
    }
  })()`)
}

function focusPPTSlideThumb(page, slideId) {
  return page.eval(`((slideId) => {
    const thumb = [...document.querySelectorAll('.ppt-thumb')]
      .find((item) => item.getAttribute('data-ppt-slide-id') === slideId)

    thumb?.focus()
  })(${JSON.stringify(slideId)})`)
}

function getPPTPlaceholderVisibilityState(page, placeholderId = 'media') {
  return page.eval(`((placeholderId) => {
    const list = document.querySelector('[data-ppt-layout-placeholder-count]')
    const placeholder = document.querySelector(\`[data-ppt-layout-placeholder="\${placeholderId}"]\`)
    const stage = document.querySelector('.ppt-stage-shell')
    const slide = document.querySelector('.ppt-slide')
    const toggle = document.querySelector(\`[data-ppt-placeholder-visibility-toggle="\${placeholderId}"]\`)
    const selectedIds = [...document.querySelectorAll('[data-selected="true"]')]
      .map((element) => element.getAttribute('data-ppt-element'))
      .filter(Boolean)

    return {
      bounds: placeholder?.getAttribute('data-ppt-placeholder-bounds') ?? '',
      command: stage?.getAttribute('data-ppt-placeholder-visibility-command') ?? '',
      commandPlaceholder: stage?.getAttribute('data-ppt-placeholder-visibility-command-placeholder') ?? '',
      commandSelection: stage?.getAttribute('data-ppt-placeholder-visibility-command-selection') ?? '',
      commandSlide: stage?.getAttribute('data-ppt-placeholder-visibility-command-slide') ?? '',
      commandType: stage?.getAttribute('data-ppt-placeholder-visibility-command-type') ?? '',
      commandVisible: stage?.getAttribute('data-ppt-placeholder-visibility-command-visible') ?? '',
      count: Number(list?.getAttribute('data-ppt-layout-placeholder-count') ?? 0),
      hiddenCount: Number(list?.getAttribute('data-ppt-layout-placeholder-hidden-count') ?? 0),
      layout: placeholder?.getAttribute('data-ppt-placeholder-layout') ?? '',
      locked: placeholder?.getAttribute('data-ppt-placeholder-locked') ?? '',
      master: placeholder?.getAttribute('data-ppt-placeholder-master') ?? '',
      placeholderId: placeholder?.getAttribute('data-ppt-layout-placeholder') ?? '',
      role: placeholder?.getAttribute('data-ppt-placeholder-role') ?? '',
      selectedIds: selectedIds.join(','),
      slideHiddenPlaceholders: slide?.getAttribute('data-ppt-hidden-placeholders') ?? '',
      slideId: placeholder?.getAttribute('data-ppt-placeholder-slide') ?? '',
      toggleDisabled: toggle?.disabled ?? null,
      visible: placeholder?.getAttribute('data-ppt-placeholder-visible') ?? '',
    }
  })(${JSON.stringify(placeholderId)})`)
}

function getPPTSlideLayoutImportState(page, placeholderId = 'body') {
  return page.eval(`((placeholderId) => {
    const list = document.querySelector('[data-ppt-layout-placeholder-count]')
    const placeholder = document.querySelector(\`[data-ppt-layout-placeholder="\${placeholderId}"]\`)
    const stage = document.querySelector('.ppt-stage-shell')
    const slide = document.querySelector('.ppt-slide')

    return {
      command: stage?.getAttribute('data-ppt-placeholder-visibility-command') ?? '',
      commandFields: stage?.getAttribute('data-ppt-slide-layout-import-command-fields') ?? '',
      commandPlaceholder: stage?.getAttribute('data-ppt-placeholder-visibility-command-placeholder') ?? '',
      commandTypes: stage?.getAttribute('data-ppt-slide-layout-import-command-types') ?? '',
      commandVisible: stage?.getAttribute('data-ppt-placeholder-visibility-command-visible') ?? '',
      count: Number(list?.getAttribute('data-ppt-layout-placeholder-count') ?? 0),
      hiddenCount: Number(list?.getAttribute('data-ppt-layout-placeholder-hidden-count') ?? 0),
      importCommands: stage?.getAttribute('data-ppt-slide-layout-import-commands') ?? '',
      importFields: stage?.getAttribute('data-ppt-slide-layout-import-fields') ?? '',
      importFormat: stage?.getAttribute('data-ppt-slide-layout-import-format') ?? '',
      importHiddenPlaceholders: stage?.getAttribute('data-ppt-slide-layout-import-hidden-placeholders') ?? '',
      importJsonLength: Number(stage?.getAttribute('data-ppt-slide-layout-import-json-length') ?? 0),
      importLayout: stage?.getAttribute('data-ppt-slide-layout-import-layout') ?? '',
      importModel: stage?.getAttribute('data-ppt-slide-layout-import-model') ?? '',
      importPlaceholderCommandCount: Number(stage?.getAttribute('data-ppt-slide-layout-import-placeholder-command-count') ?? 0),
      importSlide: stage?.getAttribute('data-ppt-slide-layout-import-slide') ?? '',
      importTheme: stage?.getAttribute('data-ppt-slide-layout-import-theme') ?? '',
      layout: slide?.getAttribute('data-ppt-layout-id') ?? '',
      placeholderId: placeholder?.getAttribute('data-ppt-layout-placeholder') ?? '',
      slideHiddenPlaceholders: slide?.getAttribute('data-ppt-hidden-placeholders') ?? '',
      theme: slide?.getAttribute('data-ppt-theme-id') ?? '',
      visible: placeholder?.getAttribute('data-ppt-placeholder-visible') ?? '',
    }
  })(${JSON.stringify(placeholderId)})`)
}

function dragPPTSlideThumbnail(page, {
  placement = 'before',
  sourceSelector = '.ppt-thumb[aria-current="page"]',
  targetSelector = '.ppt-slide-list .ppt-thumb:first-child',
} = {}) {
  return page.eval(`((input) => {
    const source = document.querySelector(input.sourceSelector)
    const target = document.querySelector(input.targetSelector)

    if (!(source instanceof HTMLElement) || !(target instanceof HTMLElement)) {
      return {
        ok: false,
        sourceFound: source instanceof HTMLElement,
        targetFound: target instanceof HTMLElement,
      }
    }

    const rect = target.getBoundingClientRect()
    const clientX = rect.left + rect.width / 2
    const clientY = input.placement === 'before'
      ? rect.top + 2
      : rect.bottom - 2
    const dataTransfer = typeof DataTransfer === 'function'
      ? new DataTransfer()
      : {
          data: new Map(),
          dropEffect: 'move',
          effectAllowed: 'move',
          getData(type) {
            return this.data.get(type) ?? ''
          },
          setData(type, value) {
            this.data.set(type, value)
          },
        }

    function createDragEvent(type) {
      let event

      try {
        event = new DragEvent(type, {
          bubbles: true,
          cancelable: true,
          clientX,
          clientY,
          dataTransfer,
        })
      } catch {
        event = new Event(type, {
          bubbles: true,
          cancelable: true,
        })
      }

      if (!event.dataTransfer) {
        Object.defineProperty(event, 'dataTransfer', {
          configurable: true,
          value: dataTransfer,
        })
      }
      if (event.clientX !== clientX) {
        Object.defineProperty(event, 'clientX', {
          configurable: true,
          value: clientX,
        })
      }
      if (event.clientY !== clientY) {
        Object.defineProperty(event, 'clientY', {
          configurable: true,
          value: clientY,
        })
      }

      return event
    }

    const sourceId = source.getAttribute('data-ppt-slide-id') ?? ''
    const sourceIndex = source.getAttribute('data-ppt-slide-index') ?? ''
    const targetId = target.getAttribute('data-ppt-slide-id') ?? ''
    const targetIndex = target.getAttribute('data-ppt-slide-index') ?? ''

    source.dispatchEvent(createDragEvent('dragstart'))
    target.dispatchEvent(createDragEvent('dragover'))
    target.dispatchEvent(createDragEvent('drop'))
    source.dispatchEvent(createDragEvent('dragend'))

    return {
      ok: true,
      sourceId,
      sourceIndex,
      targetId,
      targetIndex,
    }
  })(${JSON.stringify({ placement, sourceSelector, targetSelector })})`)
}

async function deletePPTSlideByThumbName(page, name) {
  const selected = await page.eval(`((name) => {
    const thumb = [...document.querySelectorAll('.ppt-thumb')]
      .find((item) =>
        item.querySelector('.ppt-thumb-name')?.textContent?.includes(name))

    if (!(thumb instanceof HTMLElement)) {
      return false
    }

    thumb.click()
    return true
  })(${JSON.stringify(name)})`)

  await delay(80)

  if (!selected) {
    return false
  }

  await page.eval(`document.querySelector('[data-ppt-slide-action="delete"]')?.click()`)
  await delay(100)

  return true
}

async function deletePPTSlidesByThumbNameIncludes(page, names) {
  for (const name of names) {
    let guard = 0

    while (guard < 20 && await deletePPTSlideByThumbName(page, name)) {
      guard += 1
    }
  }
}

function getPPTCrossSlideClipboardState(page) {
  return page.eval(`(() => {
    const stage = document.querySelector('.ppt-stage-shell')
    const slide = document.querySelector('.ppt-slide')
    const selected = document.querySelector('[data-selected="true"]')
    const selectedElements = [...document.querySelectorAll('[data-selected="true"]')]
    const thumbs = [...document.querySelectorAll('.ppt-thumb')]

    return {
      activeSlide: slide?.getAttribute('data-ppt-slide') ?? '',
      clipboardCount: Number(stage?.getAttribute('data-ppt-clipboard-count') ?? 0),
      clipboardMetadataCount: Number(stage?.getAttribute('data-ppt-clipboard-metadata-count') ?? 0),
      clipboardModel: stage?.getAttribute('data-ppt-clipboard-model') ?? '',
      clipboardOperation: stage?.getAttribute('data-ppt-clipboard-operation') ?? '',
      clipboardSelection: stage?.getAttribute('data-ppt-clipboard-selection') ?? '',
      clipboardSelectedObjectIds: stage?.getAttribute('data-ppt-clipboard-selected-object-ids') ?? '',
      clipboardSourceSlide: stage?.getAttribute('data-ppt-clipboard-source-slide') ?? '',
      clipboardType: stage?.getAttribute('data-ppt-clipboard-type') ?? '',
      fallbackHTMLImportCount: Number(stage?.getAttribute('data-ppt-fallback-html-import-count') ?? 0),
      fallbackHTMLImportFormat: stage?.getAttribute('data-ppt-fallback-html-import-format') ?? '',
      fallbackHTMLImportKind: stage?.getAttribute('data-ppt-fallback-html-import-kind') ?? '',
      fallbackHTMLImportModel: stage?.getAttribute('data-ppt-fallback-html-import-model') ?? '',
      fallbackHTMLImportName: stage?.getAttribute('data-ppt-fallback-html-import-name') ?? '',
      fallbackHTMLImportShape: stage?.getAttribute('data-ppt-fallback-html-import-shape') ?? '',
      fallbackHTMLImportSourceObject: stage?.getAttribute('data-ppt-fallback-html-import-source-object') ?? '',
      fallbackHTMLImportSourceObjects: stage?.getAttribute('data-ppt-fallback-html-import-source-objects') ?? '',
      elementsJSONImportCount: Number(stage?.getAttribute('data-ppt-elements-json-import-count') ?? 0),
      elementsJSONImportFormat: stage?.getAttribute('data-ppt-elements-json-import-format') ?? '',
      elementsJSONImportJsonLength: Number(stage?.getAttribute('data-ppt-elements-json-import-json-length') ?? 0),
      elementsJSONImportModel: stage?.getAttribute('data-ppt-elements-json-import-model') ?? '',
      elementsJSONImportSelection: stage?.getAttribute('data-ppt-elements-json-import-selection') ?? '',
      elementsJSONImportSourceSlide: stage?.getAttribute('data-ppt-elements-json-import-source-slide') ?? '',
      keyboardCommandDispatch: stage?.getAttribute('data-ppt-keyboard-command-dispatch') ?? '',
      keyboardCommandIntent: stage?.getAttribute('data-ppt-keyboard-command-intent') ?? '',
      pasteAnchor: stage?.getAttribute('data-ppt-clipboard-paste-anchor') ?? '',
      pasteCommand: stage?.getAttribute('data-ppt-clipboard-paste-command') ?? '',
      pasteMappingCount: Number(stage?.getAttribute('data-ppt-clipboard-paste-mapping-count') ?? 0),
      pasteOperation: stage?.getAttribute('data-ppt-clipboard-paste-operation') ?? '',
      pastePositionBoundsHeight: Number(stage?.getAttribute('data-ppt-clipboard-paste-position-bounds-height') ?? 0),
      pastePositionBoundsWidth: Number(stage?.getAttribute('data-ppt-clipboard-paste-position-bounds-width') ?? 0),
      pastePositionBoundsX: Number(stage?.getAttribute('data-ppt-clipboard-paste-position-bounds-x') ?? 0),
      pastePositionBoundsY: Number(stage?.getAttribute('data-ppt-clipboard-paste-position-bounds-y') ?? 0),
      pastePositionCount: Number(stage?.getAttribute('data-ppt-clipboard-paste-position-count') ?? 0),
      pastePositionIndex: Number(stage?.getAttribute('data-ppt-clipboard-paste-position-index') ?? 0),
      pastePositionModel: stage?.getAttribute('data-ppt-clipboard-paste-position-model') ?? '',
      pastePositionViewportX: Number(stage?.getAttribute('data-ppt-clipboard-paste-position-viewport-x') ?? 0),
      pastePositionViewportY: Number(stage?.getAttribute('data-ppt-clipboard-paste-position-viewport-y') ?? 0),
      pasteSelection: stage?.getAttribute('data-ppt-clipboard-paste-selection') ?? '',
      pasteSourceSlide: stage?.getAttribute('data-ppt-clipboard-paste-source-slide') ?? '',
      pasteTargetSlide: stage?.getAttribute('data-ppt-clipboard-paste-target-slide') ?? '',
      pasteType: stage?.getAttribute('data-ppt-clipboard-paste-type') ?? '',
      richClipboardFormats: stage?.getAttribute('data-ppt-rich-clipboard-formats') ?? '',
      richClipboardHTMLLength: Number(stage?.getAttribute('data-ppt-rich-clipboard-html-length') ?? 0),
      richClipboardImported: stage?.getAttribute('data-ppt-rich-clipboard-imported') ?? '',
      richClipboardImportFormat: stage?.getAttribute('data-ppt-rich-clipboard-import-format') ?? '',
      richClipboardJsonMimeType: stage?.getAttribute('data-ppt-rich-clipboard-json-mime-type') ?? '',
      richClipboardModel: stage?.getAttribute('data-ppt-rich-clipboard-model') ?? '',
      richClipboardObjectCount: Number(stage?.getAttribute('data-ppt-rich-clipboard-object-count') ?? 0),
      richClipboardPlainTextLength: Number(stage?.getAttribute('data-ppt-rich-clipboard-plain-text-length') ?? 0),
      richClipboardSelection: stage?.getAttribute('data-ppt-rich-clipboard-selection') ?? '',
      richClipboardSourceSlide: stage?.getAttribute('data-ppt-rich-clipboard-source-slide') ?? '',
      richClipboardWriteMode: stage?.getAttribute('data-ppt-rich-clipboard-write-mode') ?? '',
      selectedCount: document.querySelectorAll('[data-selected="true"]').length,
      selectedBorderColor: selected ? getComputedStyle(selected).borderColor : '',
      selectedCornerRadius: selected?.getAttribute('data-ppt-corner-radius') ?? '',
      selectedFill: selected ? getComputedStyle(selected).background : '',
      selectedFontSize: selected ? getComputedStyle(selected).fontSize : '',
      selectedHeight: Number.parseFloat(selected?.style.height ?? '0'),
      selectedId: selected?.getAttribute('data-ppt-element') ?? '',
      selectedIds: selectedElements.map((element) =>
        element.getAttribute('data-ppt-element') ?? ''),
      selectedKind: selected?.getAttribute('data-kind') ?? '',
      selectedKinds: selectedElements.map((element) =>
        element.getAttribute('data-kind') ?? ''),
      selectedName: document.querySelector('[data-ppt-layer-row][aria-selected="true"] .ppt-layer-name')?.textContent ?? '',
      selectedShape: selected?.getAttribute('data-shape') ?? '',
      selectedText: selected?.querySelector('.ppt-element-editor, [data-ppt-comment-body], .ppt-table-grid')?.textContent?.trim() ?? selected?.getAttribute('data-ppt-alt-text') ?? selected?.getAttribute('data-ppt-element-name') ?? '',
      selectedTexts: selectedElements.map((element) =>
        element.querySelector('.ppt-element-editor, [data-ppt-comment-body], .ppt-table-grid')?.textContent?.trim() ?? element.getAttribute('data-ppt-alt-text') ?? element.getAttribute('data-ppt-element-name') ?? ''),
      selectedWidth: Number.parseFloat(selected?.style.width ?? '0'),
      selectedX: Number.parseFloat(selected?.style.left ?? '0'),
      selectedY: Number.parseFloat(selected?.style.top ?? '0'),
      stageCount: document.querySelectorAll('[data-ppt-element]').length,
      thumbCounts: thumbs.map((thumb) => thumb.querySelectorAll('.ppt-thumb-preview > span').length),
    }
  })()`)
}

async function pressKey(page, {
  code,
  key,
  modifiers = 0,
  windowsVirtualKeyCode,
}) {
  await page.send('Input.dispatchKeyEvent', {
    code,
    key,
    modifiers,
    type: 'keyDown',
    windowsVirtualKeyCode,
  })
  await page.send('Input.dispatchKeyEvent', {
    code,
    key,
    modifiers,
    type: 'keyUp',
    windowsVirtualKeyCode,
  })
}

async function runMobileScenario(cdpPort) {
  const page = await openPage(cdpPort, appUrl, {
    deviceScaleFactor: 2,
    height: 780,
    mobile: true,
    width: 390,
  })
  const state = await page.eval(`(() => ({
    app: !!document.querySelector('[data-ppt-app]'),
    railWidth: Math.round(document.querySelector('.ppt-rail')?.getBoundingClientRect().width ?? 0),
    stageVisible: (document.querySelector('.ppt-stage-shell')?.getBoundingClientRect().width ?? 0) > 120,
  }))()`)

  record('mobile shell renders', state.app && state.stageVisible && state.railWidth > 0, state)
  await page.close()
}

async function ensureAppServer() {
  if (EXTERNAL_APP_URL) {
    if (await isHttpReady(appUrl)) return

    throw new Error(`APP_URL is not reachable: ${appUrl}`)
  }

  const vitePort = await getFreePort()
  appUrl = `http://127.0.0.1:${vitePort}/`
  devServer = spawn('pnpm', [
    'dev',
    '--host',
    '127.0.0.1',
    '--port',
    String(vitePort),
    '--strictPort',
  ], {
    cwd: process.cwd(),
    detached: true,
    stdio: ['ignore', 'pipe', 'pipe'],
  })
  devServer.stdout.on('data', () => undefined)
  devServer.stderr.on('data', () => undefined)

  await waitUntil(async () => isHttpReady(appUrl), 'Vite dev server did not start', 15000)
}

async function launchChrome(port) {
  chromeProfile = await mkdtemp(join(tmpdir(), 'ppt-subset-chrome-'))
  chrome = spawn(CHROME_BIN, [
    '--headless=new',
    `--remote-debugging-port=${port}`,
    `--user-data-dir=${chromeProfile}`,
    '--no-default-browser-check',
    '--no-first-run',
    'about:blank',
  ], {
    detached: true,
    stdio: ['ignore', 'pipe', 'pipe'],
  })
  chrome.stderr.on('data', () => undefined)
  chrome.stdout.on('data', () => undefined)

  await waitUntil(
    async () => {
      try {
        const response = await fetch(`http://127.0.0.1:${port}/json/version`)
        return response.ok
      } catch {
        return false
      }
    },
    'Chrome CDP did not start',
    15000,
  )
}

async function openPage(cdpPort, url, viewport = null) {
  const target = await fetch(
    `http://127.0.0.1:${cdpPort}/json/new?${encodeURIComponent(url)}`,
    { method: 'PUT' },
  ).then((response) => response.json())
  const ws = new WebSocket(target.webSocketDebuggerUrl)
  let id = 0
  const pending = new Map()

  ws.addEventListener('message', (event) => {
    const message = JSON.parse(event.data)

    if (message.method === 'Runtime.exceptionThrown') {
      browserErrors.push(
        message.params.exceptionDetails.exception?.description ??
          message.params.exceptionDetails.text,
      )
    }
    if (message.method === 'Log.entryAdded' && message.params.entry.level === 'error') {
      browserErrors.push(message.params.entry.text)
    }

    if (!message.id || !pending.has(message.id)) return

    const { resolve, reject, timeout } = pending.get(message.id)
    pending.delete(message.id)
    clearTimeout(timeout)
    if (message.error) reject(new Error(JSON.stringify(message.error)))
    else resolve(message.result)
  })

  await new Promise((resolve, reject) => {
    ws.addEventListener('open', resolve, { once: true })
    ws.addEventListener('error', reject, { once: true })
  })

  const page = {
    send(method, params = {}) {
      const callId = ++id
      ws.send(JSON.stringify({ id: callId, method, params }))
      return new Promise((resolve, reject) => {
        const timeout = setTimeout(() => {
          pending.delete(callId)
          reject(new Error(`Timed out CDP command: ${method}`))
        }, CDP_COMMAND_TIMEOUT_MS)

        pending.set(callId, { resolve, reject, timeout })
      })
    },
    async eval(expression) {
      const result = await page.send('Runtime.evaluate', {
        awaitPromise: true,
        expression,
        returnByValue: true,
      })
      if (result.exceptionDetails) {
        throw new Error(JSON.stringify(result.exceptionDetails))
      }
      return result.result.value
    },
    close() {
      const closeTarget = fetch(
        `http://127.0.0.1:${cdpPort}/json/close/${target.id}`,
      ).catch(() => undefined)

      ws.close()
      return closeTarget
    },
  }

  await page.send('Runtime.enable')
  await page.send('Log.enable')
  await page.send('Page.enable')
  if (viewport) {
    await page.send('Emulation.setDeviceMetricsOverride', viewport)
  }
  await page.send('Page.navigate', { url })
  await waitUntil(
    async () =>
      page.eval(
        `location.href === ${JSON.stringify(url)} && document.readyState === 'complete' && !!document.querySelector('[data-ppt-app]')`,
      ),
    `Timed out loading page: ${url}`,
    15000,
  )

  return page
}

async function stopChrome() {
  if (!chrome) return

  const exited = new Promise((resolve) => chrome.once('exit', resolve))

  if (chrome.pid) {
    try {
      process.kill(-chrome.pid, 'SIGTERM')
    } catch {
      chrome.kill()
    }
  } else {
    chrome.kill()
  }

  await Promise.race([exited, delay(1000)])

  if (chrome.exitCode === null && chrome.signalCode === null && chrome.pid) {
    try {
      process.kill(-chrome.pid, 'SIGKILL')
    } catch {
      chrome.kill('SIGKILL')
    }
  }
}

async function stopDevServer() {
  if (!devServer) return

  const exited = new Promise((resolve) => devServer.once('exit', resolve))

  if (devServer.pid) {
    try {
      process.kill(-devServer.pid, 'SIGTERM')
    } catch {
      devServer.kill()
    }
  } else {
    devServer.kill()
  }

  await Promise.race([exited, delay(1000)])

  if (devServer.exitCode === null && devServer.signalCode === null && devServer.pid) {
    try {
      process.kill(-devServer.pid, 'SIGKILL')
    } catch {
      devServer.kill('SIGKILL')
    }
  }
}

function record(name, ok, details = undefined) {
  checks.push({
    details,
    name,
    ok: Boolean(ok),
  })
}

async function getFreePort() {
  const server = net.createServer()

  await new Promise((resolve, reject) => {
    server.once('error', reject)
    server.listen(0, '127.0.0.1', resolve)
  })

  const address = server.address()
  const port = typeof address === 'object' && address ? address.port : null

  await new Promise((resolve, reject) => {
    server.once('error', reject)
    server.close(resolve)
  })

  if (!port) {
    throw new Error('Could not allocate a free port')
  }

  return port
}

async function isHttpReady(url) {
  try {
    const response = await fetch(url)
    return response.ok
  } catch {
    return false
  }
}

async function waitUntil(check, message, timeout = 5000) {
  const started = Date.now()

  while (Date.now() - started < timeout) {
    if (await check()) {
      return
    }
    await delay(100)
  }

  throw new Error(message)
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms))
}
