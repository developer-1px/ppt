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
      undoEnabled: !document.querySelector('button[title="Undo"]').disabled,
    }
  })()`)

  record('replaces active PPT text match while preserving textBody structure', afterReplace.text.includes('PPT-ready later') && afterReplace.count === '0/0' && afterReplace.hasTextBodyRuns && afterReplace.undoEnabled, afterReplace)

  await page.eval(`document.querySelector('button[title="Undo"]').click()`)
  await delay(100)

  const afterUndo = await page.eval(`(() => {
    const selected = document.querySelector('[data-selected="true"]')

    return {
      count: document.querySelector('[data-ppt-find-count]')?.textContent ?? '',
      redoEnabled: !document.querySelector('button[title="Redo"]').disabled,
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
      selectedAutoFit: element?.getAttribute('data-ppt-text-autofit') ?? '',
      width: parseFloat(element.style.width),
    }
  })()`)

  record('auto-sizes selected object from resize handle double-click', afterAuto.width !== afterResize.width, {
    afterAuto,
    afterResize,
  })
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
      left: parseFloat(element.style.left),
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
      afterContextOpen.keyboard === 'arrow-home-end-enter-escape' &&
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

  await pressKey(page, {
    code: 'KeyC',
    key: 'c',
    modifiers: 2,
    windowsVirtualKeyCode: 67,
  })
  await delay(50)

  const afterCopy = await getPPTCrossSlideClipboardState(page)

  record('stores PPT clipboard source slide metadata', afterCopy.keyboardCommandIntent === 'canvas-keyboard-command-shortcut-intent' && afterCopy.keyboardCommandDispatch === 'canvas-keyboard-command-dispatch' && afterCopy.clipboardModel === 'slide-edit-clipboard' && afterCopy.clipboardCount === 1 && afterCopy.clipboardSourceSlide === 'slide-1' && afterCopy.clipboardSelection === 's1-title' && afterCopy.clipboardType === 'slide-object-clipboard' && afterCopy.clipboardOperation === 'copy' && afterCopy.clipboardMetadataCount === 1 && afterCopy.clipboardSelectedObjectIds === 's1-title', {
    afterCopy,
    sourceBefore,
  })

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

  record('pastes copied PPT object onto another slide with target slide ids', afterKeyboardPaste.keyboardCommandIntent === 'canvas-keyboard-command-shortcut-intent' && afterKeyboardPaste.keyboardCommandDispatch === 'canvas-keyboard-command-dispatch' && afterKeyboardPaste.activeSlide === 'slide-2' && afterKeyboardPaste.stageCount === targetBefore.stageCount + 1 && afterKeyboardPaste.selectedCount === 1 && afterKeyboardPaste.selectedId.startsWith('slide-2-') && afterKeyboardPaste.selectedName.includes('Copy'), {
    afterKeyboardPaste,
    targetBefore,
  })
  record('creates PPT cross-slide paste command effect plan', afterKeyboardPaste.pasteCommand === 'paste-slide-objects' && afterKeyboardPaste.pasteType === 'slide-command-effect' && afterKeyboardPaste.pasteSourceSlide === 'slide-1' && afterKeyboardPaste.pasteTargetSlide === 'slide-2' && afterKeyboardPaste.pasteMappingCount === 1 && afterKeyboardPaste.pasteSelection === afterKeyboardPaste.selectedId && afterKeyboardPaste.pasteAnchor === '28,28' && afterKeyboardPaste.pasteOperation === 'copy', {
    afterKeyboardPaste,
    targetBefore,
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

    return {
      bulletList: element?.getAttribute('data-ppt-bullet-list') ?? '',
      color: element?.style.color ?? '',
      fontSize: element?.style.fontSize ?? '',
      height: element?.style.height ?? '',
      left: element?.style.left ?? '',
      name: layerName?.textContent ?? '',
      paragraphBullet: paragraph?.getAttribute('data-ppt-bullet') === 'true'
        ? paragraph.textContent ?? ''
        : '',
      selected: element?.getAttribute('data-selected') ?? '',
      text: element?.textContent ?? '',
      textAlign: element?.style.textAlign ?? '',
      top: element?.style.top ?? '',
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
    quickBarVisible: !!document.querySelector('[data-ppt-text-quick-bar]'),
    selectedId: document.querySelector('[data-selected="true"]')?.getAttribute('data-ppt-element') ?? '',
    underlinePressed: document.querySelector('[data-ppt-text-quick="underline"]')?.getAttribute('aria-pressed') ?? '',
  }))()`)

  record('renders PPT text quick format bar for selected text', initial.quickBarVisible && initial.selectedId === 's1-title' && initial.boldPressed === 'true' && initial.bulletPressed === 'false' && initial.italicPressed === 'false' && initial.underlinePressed === 'false' && initial.fontSize > 0, initial)

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
      italicPressed: document.querySelector('[data-ppt-text-quick="italic"]')?.getAttribute('aria-pressed') ?? '',
      italicRun: title?.querySelector('[data-ppt-run-italic="true"]')?.style.fontStyle ?? '',
      paragraphBullet: title?.querySelector('[data-ppt-bullet="true"]')?.textContent ?? '',
      rightPressed: document.querySelector('[data-ppt-paragraph-align="right"]')?.getAttribute('aria-pressed') ?? '',
      textAlign: title?.style.textAlign ?? '',
      thumbBulletCount: document.querySelectorAll('[data-ppt-thumb-bullet="true"]').length,
      underlinePressed: document.querySelector('[data-ppt-text-quick="underline"]')?.getAttribute('aria-pressed') ?? '',
      underlineRun: title?.querySelector('[data-ppt-run-underline="true"]')?.style.textDecoration ?? '',
    }
  })()`)

  record('applies PPT text quick formatting to selected text model', afterSingleFormat.color === 'rgb(0, 85, 255)' && afterSingleFormat.fontSize === initial.fontSize + 2 && afterSingleFormat.fontWeight === 'regular' && afterSingleFormat.textAlign === 'right' && afterSingleFormat.rightPressed === 'true' && afterSingleFormat.bulletList === 'true' && afterSingleFormat.bulletPressed === 'true' && afterSingleFormat.inspectorBulletPressed === 'true' && afterSingleFormat.italicPressed === 'true' && afterSingleFormat.underlinePressed === 'true' && afterSingleFormat.italicRun === 'italic' && afterSingleFormat.underlineRun.includes('underline') && afterSingleFormat.paragraphBullet.length > 0 && afterSingleFormat.thumbBulletCount > 0, {
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
      summaryAfterFormatPaste.paragraphBullet.length > 0,
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

  await pressKey(page, {
    code: 'Escape',
    key: 'Escape',
    windowsVirtualKeyCode: 27,
  })
  await delay(50)
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
      preview.verticalAlign === 'middle' &&
      preview.alignItems === 'center',
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
      preview.textInset === '10,14,18,22' &&
      preview.padding === '10px 14px 18px 22px',
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
      hasParagraphSpacingMarkup: code.includes('data-ppt-line-height="1.4"') && code.includes('data-ppt-spacing-before="6"') && code.includes('data-ppt-spacing-after="12"') && code.includes('line-height:1.4') && code.includes('margin-top:6px') && code.includes('margin-bottom:12px'),
      hasParagraphSpacingModel: code.includes('"lineHeight": 1.4') && code.includes('"spacingBefore": 6') && code.includes('"spacingAfter": 12'),
      hasTextFrameInsetMarkup: code.includes('data-ppt-text-inset="10,14,18,22"') && code.includes('padding:10px 14px 18px 22px'),
      hasTextFrameInsetModel: code.includes('"textInset"') && code.includes('"top": 10') && code.includes('"right": 14') && code.includes('"bottom": 18') && code.includes('"left": 22'),
      hasTextVerticalAlignMarkup: code.includes('data-ppt-vertical-align="middle"') && code.includes('align-items:center'),
      hasTextVerticalAlignModel: code.includes('"verticalAlign": "middle"'),
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
      hasParagraphSpacing: text.includes('data-ppt-line-height="1.4"') && text.includes('data-ppt-spacing-before="6"') && text.includes('data-ppt-spacing-after="12"'),
      hasTextFrameInset: text.includes('data-ppt-text-inset="10,14,18,22"'),
      hasTextVerticalAlign: text.includes('data-ppt-vertical-align="middle"'),
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

  const imageId = await page.eval(`(() => [...document.querySelectorAll('[data-kind="image"]')].at(-1)?.getAttribute('data-ppt-element') ?? '')()`)
  await selectPPTLayerRows(page, [imageId])
  await delay(80)

  const beforeSelectionSvg = await page.eval(`(() => ({
    disabled: document.querySelector('[data-ppt-export-selection-svg]')?.disabled ?? true,
    selectedId: document.querySelector('[data-selected="true"]')?.getAttribute('data-ppt-element') ?? '',
  }))()`)

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

  record('enables selected-object PPT SVG export', beforeSelectionSvg.selectedId === imageId && !beforeSelectionSvg.disabled, beforeSelectionSvg)
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
      expanded: trigger?.getAttribute('aria-expanded') ?? '',
      itemCommands: items.map((item) => item.getAttribute('data-ppt-alignment-popover-command') ?? ''),
      itemCount: items.length,
      model: popover?.getAttribute('data-ppt-alignment-popover-model') ?? '',
      role: popover?.getAttribute('role') ?? '',
      selectedId: document.querySelector('[data-selected="true"]')?.getAttribute('data-ppt-element') ?? '',
      triggerHasPopup: trigger?.getAttribute('aria-haspopup') ?? '',
    }
  })()`)

  record('opens PPT selection alignment popover from floating bar', alignmentPopoverOpen.selectedId !== '' && alignmentPopoverOpen.expanded === 'true' && alignmentPopoverOpen.role === 'menu' && alignmentPopoverOpen.model === 'canvas-dom-alignment-popover' && alignmentPopoverOpen.triggerHasPopup === 'menu' && alignmentPopoverOpen.itemCount === 8 && alignmentPopoverOpen.itemCommands.includes('align-center-x') && alignmentPopoverOpen.itemCommands.includes('distribute-horizontal'), alignmentPopoverOpen)

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

  record('opens PPT Shape menu from selection floating bar', point.id !== '' && shapeMenuOpen.selectedId === point.id && shapeMenuOpen.expanded === 'true' && shapeMenuOpen.hasPopup === 'menu' && shapeMenuOpen.role === 'menu' && shapeMenuOpen.model === 'canvas-selection-toolbar-dropdown-menu' && shapeMenuOpen.controls === 'ppt-shape-kind-menu' && shapeMenuOpen.itemCount === 3 && shapeMenuOpen.itemRoles.every((role) => role === 'menuitemcheckbox') && shapeMenuOpen.itemShapes.join(' ') === 'rect ellipse diamond' && shapeMenuOpen.checkedShape === point.shape, {
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
      initial.stageType === 'none',
    initial,
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
      afterFade.stageAdvanceAfter === '3000',
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
      Number(initial.order) > 0 &&
      initial.selectedType === 'none',
    initial,
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
      afterFade.descriptorPackageTrigger === 'with-previous',
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
      preview.order === '3',
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

  record('inserts PPT image from file picker affordance', afterUpload.imageCount === before.imageCount + 1 && afterUpload.selectedKind === 'image' && afterUpload.selectedImageSrc.startsWith('data:image/svg+xml'), {
    afterUpload,
    before,
  })

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

  record('pastes image file into PPT slide from clipboard event', afterPaste.imageCount === afterUpload.imageCount + 1 && afterPaste.selectedKind === 'image' && afterPaste.selectedName === 'paste.svg', {
    afterPaste,
    afterUpload,
  })

  await page.eval(`(() => {
    const stage = document.querySelector('.ppt-stage-shell')
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

  record('drops image file onto PPT stage at pointer position', afterDrop.imageCount === afterPaste.imageCount + 1 && afterDrop.selectedKind === 'image' && afterDrop.selectedName === 'drop.svg' && afterDrop.selectedLeft > 0 && afterDrop.selectedTop >= 0, {
    afterDrop,
    afterPaste,
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

  await pressKey(page, {
    code: 'KeyK',
    key: 'k',
    modifiers: 2,
    windowsVirtualKeyCode: 75,
  })
  await delay(80)
  await page.send('Input.insertText', { text: 'table' })
  await delay(80)
  await pressKey(page, {
    code: 'Enter',
    key: 'Enter',
    windowsVirtualKeyCode: 13,
  })
  await delay(100)

  const afterPaletteInsert = await getPPTTableState(page)

  record('inserts PPT table from command palette', afterPaletteInsert.tableCount === afterInspectorEdit.tableCount + 1 && afterPaletteInsert.selectedKind === 'table' && afterPaletteInsert.selectedRows === 3 && afterPaletteInsert.selectedCols === 3 && !afterPaletteInsert.paletteOpen, {
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

  record('pastes TSV clipboard data into PPT table element', afterPaste.tableCount === afterPaletteInsert.tableCount + 1 && afterPaste.selectedKind === 'table' && afterPaste.selectedRows === 3 && afterPaste.selectedCols === 2 && afterPaste.cellTexts.includes('Users') && afterPaste.cellTexts.includes('$1M'), {
    afterPaletteInsert,
    afterPaste,
  })

  await page.eval(`(() => {
    const stage = document.querySelector('.ppt-stage-shell')
    const rect = stage.getBoundingClientRect()
    const dataTransfer = new DataTransfer()

    dataTransfer.items.add(${createPPTTestTableFileExpression('metrics.csv', 'Region,Score\\nNA,88\\nEU,91')})
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

  record('drops CSV file onto PPT stage as table element', afterDrop.tableCount === afterPaste.tableCount + 1 && afterDrop.selectedKind === 'table' && ['metrics', 'Table'].includes(afterDrop.selectedName) && afterDrop.selectedRows === 3 && afterDrop.selectedCols === 2 && afterDrop.cellTexts.includes('Region') && afterDrop.cellTexts.includes('EU') && afterDrop.selectedLeft > 0 && afterDrop.selectedTop >= 0, {
    afterDrop,
    afterPaste,
  })
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

  await pressKey(page, {
    code: 'Delete',
    key: 'Delete',
    windowsVirtualKeyCode: 46,
  })
  await delay(80)

  await page.eval(`document.querySelector(${JSON.stringify(`[data-ppt-layer-select="${lineId}"]`)})?.click()`)
  await delay(50)

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

  const shapePoint = await getElementCenter(page, 's1-card-1')
  const beforeNonDrawingErase = await getPPTFreeformState(page)

  await dragMouse(page, [{
    x: shapePoint.x - 24,
    y: shapePoint.y - 18,
  }, {
    x: shapePoint.x + 24,
    y: shapePoint.y + 18,
  }])
  await delay(100)

  const afterNonDrawingErase = await getPPTFreeformState(page)

  record('keeps non-drawing PPT objects when eraser crosses shapes', afterNonDrawingErase.eraserActive === 'true' && afterNonDrawingErase.shapeCount === beforeNonDrawingErase.shapeCount && afterNonDrawingErase.freeformCount === beforeNonDrawingErase.freeformCount && afterNonDrawingErase.elementCount === beforeNonDrawingErase.elementCount, {
    afterNonDrawingErase,
    beforeNonDrawingErase,
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
      initial.keyboardKeys === 'arrow-left-right-home-end-enter-space-shift-range-alt-reorder' &&
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

  await pressKey(page, {
    code: 'KeyY',
    key: 'y',
    modifiers: 2,
    windowsVirtualKeyCode: 89,
  })
  await delay(160)

  const afterRedo = await getPPTTextOverflowState(page)

  record('redoes PPT text auto-fit with model metadata', afterRedo.selectedId === afterAutoFit.selectedId && afterRedo.selectedOverflow !== 'true' && afterRedo.selectedAutoFit === 'resizeShapeToFitText' && afterRedo.inspectorAutoFit === 'resizeShapeToFitText' && afterRedo.selectedAutoFitSizeMode === 'resize-to-fit' && afterRedo.inspectorAutoFitSizeMode === 'resize-to-fit' && afterRedo.exportHasAutoFit, {
    afterAutoFit,
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

  await page.eval(`document.querySelector('[data-ppt-slide-action="delete"]').click()`)
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

  record('exposes PPT slide rail listbox keyboard affordance', initialKeyboard.model === 'slide-edit-rail-interactions' && initialKeyboard.listRole === 'listbox' && initialKeyboard.keyboardModel === 'aria-listbox-roving-focus' && initialKeyboard.keyboardKeys === 'ArrowUp ArrowDown Home End Enter Space' && initialKeyboard.selectionMode === 'single' && initialKeyboard.activeAttr === initialKeyboard.activeId && initialKeyboard.activeOption === expectedActiveOptionId && initialKeyboard.focusableOption === expectedActiveOptionId && initialKeyboard.slideOrder === expectedSlideOrder && initialKeyboard.optionCount === initialKeyboard.count && initialKeyboard.optionCountAttr === String(initialKeyboard.count) && initialKeyboard.thumbnailCount === String(initialKeyboard.count) && initialKeyboard.optionIds.length === initialKeyboard.count && initialKeyboard.optionIds.every((id, index) => id === `slide-rail-option-${index}`) && initialKeyboard.optionIndexes.every((value, index) => value === String(index)) && initialKeyboard.optionFocusableIds.length === 1 && initialKeyboard.optionFocusableIds[0] === initialKeyboard.activeId && initialKeyboard.activeThumbW === '112' && initialKeyboard.activeThumbH === '86' && initialKeyboard.activeHitW === '124' && initialKeyboard.activeHitH === '98' && initialKeyboard.selectedOptionIds.length === 1 && initialKeyboard.selectedOptionIds[0] === initialKeyboard.activeId && initialKeyboard.tabStopIds.length === 1 && initialKeyboard.tabStopIds[0] === initialKeyboard.activeId && initialKeyboard.focusedId === initialKeyboard.activeId, {
    initialKeyboard,
  })

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
    const selectedImage = selected?.querySelector('img') ?? null
    const stage = document.querySelector('.ppt-stage-shell')
    const fitField = document.querySelector('[data-ppt-style-field="image-fit"]')
    const cropXField = document.querySelector('[data-ppt-style-field="image-crop-x"]')
    const cropYField = document.querySelector('[data-ppt-style-field="image-crop-y"]')

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
      selectedAltText: selectedImage?.getAttribute('alt') ?? '',
      selectedFlipH: selected?.getAttribute('data-ppt-flip-h') ?? '',
      selectedFlipV: selected?.getAttribute('data-ppt-flip-v') ?? '',
      selectedId: selected?.getAttribute('data-ppt-element') ?? '',
      selectedImageFit: selectedImage?.style.objectFit ?? '',
      selectedImagePosition: selectedImage?.style.objectPosition ?? '',
      selectedImageSrc: selectedImage?.getAttribute('src') ?? '',
      selectedKind: selected?.getAttribute('data-kind') ?? '',
      selectedHeight: parseFloat(selected?.style.height ?? '0'),
      selectedLeft: parseFloat(selected?.style.left ?? '0'),
      selectedName: document.querySelector('[data-ppt-layer-row][aria-selected="true"] .ppt-layer-name')?.textContent ?? '',
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
    const tableCells = [...selected?.querySelectorAll('[data-ppt-table-cell]') ?? []]

    return {
      cellTexts: tableCells.map((cell) => cell.textContent ?? ''),
      inspectorSize: document.querySelector('[data-ppt-table-inspector-size]')?.textContent?.trim() ?? '',
      inspectorValue: document.querySelector('[data-ppt-style-field="table-data"]')?.value ?? '',
      paletteOpen: !!document.querySelector('[data-ppt-command-palette]'),
      selectedCols: Number(selected?.getAttribute('data-ppt-table-cols') ?? 0),
      selectedId: selected?.getAttribute('data-ppt-element') ?? '',
      selectedKind: selected?.getAttribute('data-kind') ?? '',
      selectedLeft: parseFloat(selected?.style.left ?? '0'),
      selectedName: document.querySelector('[data-ppt-layer-row][aria-selected="true"] .ppt-layer-name')?.textContent ?? '',
      selectedRows: Number(selected?.getAttribute('data-ppt-table-rows') ?? 0),
      selectedTop: parseFloat(selected?.style.top ?? '0'),
      tableCount: document.querySelectorAll('[data-kind="table"]').length,
      thumbTableCount: document.querySelectorAll('.ppt-thumb-table').length,
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
      borderRadius: selected?.style.borderRadius ?? '',
      borderStyle: selected?.style.borderStyle ?? '',
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
    const stage = document.querySelector('.ppt-stage-world .ppt-slide')

    return {
      advanceAfter: document.querySelector('[data-ppt-slide-transition-field="advanceAfterMs"]')?.value ?? '',
      advanceOnClick: document.querySelector('[data-ppt-slide-transition-field="advanceOnClick"]')?.checked ? 'true' : 'false',
      duration: document.querySelector('[data-ppt-slide-transition-field="durationMs"]')?.value ?? '',
      inspector: !!inspector,
      inspectorAdvanceAfter: inspector?.getAttribute('data-ppt-transition-advance-after') ?? '',
      inspectorAdvanceOnClick: inspector?.getAttribute('data-ppt-transition-advance-on-click') ?? '',
      inspectorDuration: inspector?.getAttribute('data-ppt-transition-duration') ?? '',
      inspectorType: inspector?.getAttribute('data-ppt-transition-type') ?? '',
      stageAdvanceAfter: stage?.getAttribute('data-ppt-transition-advance-after') ?? '',
      stageAdvanceOnClick: stage?.getAttribute('data-ppt-transition-advance-on-click') ?? '',
      stageDuration: stage?.getAttribute('data-ppt-transition-duration') ?? '',
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
      opacity: opacity?.value ?? '',
      opacityControl: opacity?.getAttribute('data-ppt-shadow-control') ?? '',
      opacityDisabled: opacity?.disabled ?? false,
      opacityUnit: opacity?.getAttribute('data-ppt-shadow-unit') ?? '',
      selectedAngle: selected?.getAttribute('data-ppt-shadow-angle') ?? '',
      selectedBlur: selected?.getAttribute('data-ppt-shadow-blur') ?? '',
      selectedColor: selected?.getAttribute('data-ppt-shadow-color') ?? '',
      selectedDistance: selected?.getAttribute('data-ppt-shadow-distance') ?? '',
      selectedFilter: selected?.style.filter ?? '',
      selectedId: selected?.getAttribute('data-ppt-element') ?? '',
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
      model: stage?.getAttribute('data-ppt-hyperlink-model') ?? '',
      selectedId: selected?.getAttribute('data-ppt-element') ?? '',
      selectedUrl: selected?.getAttribute('data-ppt-hyperlink-url') ?? '',
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

function getPPTCrossSlideClipboardState(page) {
  return page.eval(`(() => {
    const stage = document.querySelector('.ppt-stage-shell')
    const slide = document.querySelector('.ppt-slide')
    const selected = document.querySelector('[data-selected="true"]')
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
      keyboardCommandDispatch: stage?.getAttribute('data-ppt-keyboard-command-dispatch') ?? '',
      keyboardCommandIntent: stage?.getAttribute('data-ppt-keyboard-command-intent') ?? '',
      pasteAnchor: stage?.getAttribute('data-ppt-clipboard-paste-anchor') ?? '',
      pasteCommand: stage?.getAttribute('data-ppt-clipboard-paste-command') ?? '',
      pasteMappingCount: Number(stage?.getAttribute('data-ppt-clipboard-paste-mapping-count') ?? 0),
      pasteOperation: stage?.getAttribute('data-ppt-clipboard-paste-operation') ?? '',
      pasteSelection: stage?.getAttribute('data-ppt-clipboard-paste-selection') ?? '',
      pasteSourceSlide: stage?.getAttribute('data-ppt-clipboard-paste-source-slide') ?? '',
      pasteTargetSlide: stage?.getAttribute('data-ppt-clipboard-paste-target-slide') ?? '',
      pasteType: stage?.getAttribute('data-ppt-clipboard-paste-type') ?? '',
      selectedCount: document.querySelectorAll('[data-selected="true"]').length,
      selectedId: selected?.getAttribute('data-ppt-element') ?? '',
      selectedKind: selected?.getAttribute('data-kind') ?? '',
      selectedName: document.querySelector('[data-ppt-layer-row][aria-selected="true"] .ppt-layer-name')?.textContent ?? '',
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
