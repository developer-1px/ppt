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
  await runTextEditingScenario(page)
  await runSelectionAndDragScenario(page)
  await runAffordanceScenario(page)
  await runViewAndShapeScenario(page)
  await runSelectionPaneScenario(page)
  await runExportScenario(page)
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
    alignmentGuideCount: document.querySelectorAll('.ppt-guide').length,
    spacingLabelCount: document.querySelectorAll('.ppt-spacing-label').length,
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
  record('renders canvas snap guides while dragging', duringDrag.alignmentGuideCount > 0 || duringDrag.spacingLabelCount > 0, duringDrag)
  record('moves PPT element through canvas transform adapter', after.left !== before.left, {
    after,
    before,
  })
  record('records drag in undo history', after.undoEnabled, after)
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

async function runAffordanceScenario(page) {
  const initial = await page.eval(`(() => {
    const element = document.querySelector('[data-ppt-element="s1-card-1"]')

    return {
      commandCount: document.querySelectorAll('[data-ppt-command]').length,
      frameGuideCount: document.querySelectorAll('.ppt-frame-guide').length,
      geometryInputCount: document.querySelectorAll('[data-ppt-geometry-field]').length,
      hasRotateHandle: !!document.querySelector('[data-ppt-rotate-handle]'),
      hasSizeCapsule: !!document.querySelector('.ppt-size-capsule'),
      width: parseFloat(element.style.width),
    }
  })()`)

  record('renders PPT alignment and distribution commands', initial.commandCount >= 8, initial)
  record('renders PPT frame guides for selected object', initial.frameGuideCount >= 6, initial)
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

    return {
      height: parseFloat(element.style.height),
      width: parseFloat(element.style.width),
    }
  })()`)

  record('auto-sizes selected object from resize handle double-click', afterAuto.width !== afterResize.width, {
    afterAuto,
    afterResize,
  })

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
    const element = document.querySelector('[data-ppt-element="s1-card-1"]')

    return {
      elementCount: document.querySelectorAll('[data-ppt-element]').length,
      left: parseFloat(element.style.left),
      selectedCount: document.querySelectorAll('[data-selected="true"]').length,
    }
  })()`)

  record('nudges selected object with arrow key', afterNudge.left === afterAlign.left + 1, {
    afterAlign,
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

    return {
      left: parseFloat(card.style.left),
      selectedCount: document.querySelectorAll('[data-selected="true"]').length,
      x: rect.left + rect.width / 2,
      y: rect.top + rect.height / 2,
    }
  })()`)

  record('distributes multi-selected objects with canvas command adapter', beforeDistribute.selectedCount === 3 && afterDistribute.selectedCount === 3 && afterDistribute.left !== beforeDistribute.left, {
    afterDistribute,
    beforeDistribute,
  })

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
    elementCount: document.querySelectorAll('[data-ppt-element]').length,
    selectedCount: document.querySelectorAll('[data-selected="true"]').length,
  }))()`)

  record('cuts and pastes selected PPT object with keyboard commands', afterCutPaste.elementCount === afterPaste.elementCount && afterCutPaste.selectedCount === 1, {
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

async function runExportScenario(page) {
  const state = await page.eval(`(() => {
    const code = document.querySelector('.ppt-export-code')?.value ?? ''
    return {
      hasDeckJson: code.includes('data-ppt-deck'),
      hasSlideMarkup: code.includes('data-ppt-slide="slide-1"'),
      hasElementMarkup: code.includes('data-ppt-element="s1-title"'),
      hasPPTDeckModel: code.includes('"slides"') && code.includes('"elements"'),
      hasRotationStyle: code.includes('transform:rotate(45deg)'),
    }
  })()`)

  record('exports HTML slide markup', state.hasSlideMarkup, state)
  record('exports PPT element markup', state.hasElementMarkup, state)
  record('exports embedded PPT deck JSON', state.hasDeckJson && state.hasPPTDeckModel, state)
  record('exports PPT object rotation style', state.hasRotationStyle, state)
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

  await page.eval(`document.querySelector('[data-ppt-insert-shape="ellipse"]').click()`)
  await delay(50)

  const afterInsertShape = await page.eval(`(() => {
    const selected = document.querySelector('[data-selected="true"]')

    return {
      selectedCount: document.querySelectorAll('[data-selected="true"]').length,
      shape: selected?.getAttribute('data-shape') ?? null,
      text: selected?.textContent ?? '',
    }
  })()`)

  record('inserts PPT oval shape from toolbar', afterInsertShape.selectedCount === 1 && afterInsertShape.shape === 'ellipse', afterInsertShape)

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
}

async function runSelectionPaneScenario(page) {
  const initial = await page.eval(`(() => {
    const selected = document.querySelector('[data-selected="true"]')
    const selectedId = selected?.getAttribute('data-ppt-element') ?? ''

    return {
      layerCount: document.querySelectorAll('[data-ppt-layer-row]').length,
      selectedId,
      stageCount: document.querySelectorAll('[data-ppt-element]').length,
    }
  })()`)

  record('renders PPT object selection pane', initial.layerCount === initial.stageCount && initial.layerCount > 0, initial)

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
    const selected = document.querySelector('[data-ppt-element="${initial.selectedId}"]')
    const row = document.querySelector('[data-ppt-layer-row="${initial.selectedId}"]')

    return {
      deleteDisabled: document.querySelector('button[title="Delete"]')?.disabled ?? false,
      lockDisabled: document.querySelector('[data-ppt-command="lock-selection"]')?.disabled ?? false,
      locked: selected?.getAttribute('data-locked') ?? null,
      resizeHandleCount: document.querySelectorAll('.ppt-resize-handle').length,
      rowLocked: row?.getAttribute('data-locked') ?? null,
      unlockDisabled: document.querySelector('[data-ppt-command="unlock-all"]')?.disabled ?? true,
    }
  })()`)

  record('locks selected PPT object from selection pane', afterLayerLock.locked === 'true' && afterLayerLock.rowLocked === 'true' && afterLayerLock.resizeHandleCount === 0, afterLayerLock)
  record('disables transform commands for locked PPT object', afterLayerLock.deleteDisabled && afterLayerLock.lockDisabled && !afterLayerLock.unlockDisabled, afterLayerLock)

  await page.eval(`document.querySelector('[data-ppt-command="unlock-all"]').click()`)
  await delay(50)

  const afterUnlockAll = await page.eval(`(() => {
    const selected = document.querySelector('[data-ppt-element="${initial.selectedId}"]')
    const row = document.querySelector('[data-ppt-layer-row="${initial.selectedId}"]')

    return {
      locked: selected?.getAttribute('data-locked') ?? null,
      resizeHandleCount: document.querySelectorAll('.ppt-resize-handle').length,
      rowLocked: row?.getAttribute('data-locked') ?? null,
    }
  })()`)

  record('unlocks all PPT objects from toolbar', afterUnlockAll.locked === 'false' && afterUnlockAll.rowLocked === 'false' && afterUnlockAll.resizeHandleCount > 0, afterUnlockAll)

  await page.eval(`document.querySelector('[data-ppt-layer-row="${initial.selectedId}"] [data-ppt-layer-visibility]').click()`)
  await delay(50)

  const afterHide = await page.eval(`(() => {
    const stageElement = document.querySelector('[data-ppt-element="${initial.selectedId}"]')
    const row = document.querySelector('[data-ppt-layer-row="${initial.selectedId}"]')

    return {
      hidden: row?.getAttribute('data-hidden') ?? null,
      rowSelected: row?.getAttribute('aria-selected') ?? null,
      stageElementExists: !!stageElement,
    }
  })()`)

  record('hides selected PPT object from selection pane', afterHide.hidden === 'true' && afterHide.rowSelected === 'true' && !afterHide.stageElementExists, afterHide)

  await page.eval(`document.querySelector('[data-ppt-layer-row="${initial.selectedId}"] [data-ppt-layer-visibility]').click()`)
  await delay(50)

  const afterShow = await page.eval(`(() => {
    const stageElement = document.querySelector('[data-ppt-element="${initial.selectedId}"]')
    const row = document.querySelector('[data-ppt-layer-row="${initial.selectedId}"]')

    return {
      hidden: row?.getAttribute('data-hidden') ?? null,
      stageElementExists: !!stageElement,
    }
  })()`)

  record('shows hidden PPT object from selection pane', afterShow.hidden === 'false' && afterShow.stageElementExists, afterShow)
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

  await page.eval(`document.querySelector('[data-ppt-slide-action="delete"]').click()`)
  await delay(50)

  const afterDelete = await getSlideRailState(page)

  record('deletes active PPT slide without removing final slide', afterDelete.count === before.count && afterDelete.count >= 1, {
    afterDelete,
    before,
  })
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

function getElementCenter(page, elementId) {
  return page.eval(`(() => {
    const rect = document.querySelector('[data-ppt-element="${elementId}"]').getBoundingClientRect()

    return {
      x: rect.left + rect.width / 2,
      y: rect.top + rect.height / 2,
    }
  })()`)
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
    const thumbs = [...document.querySelectorAll('.ppt-thumb')]
    const activeIndex = thumbs.findIndex((thumb) => thumb.getAttribute('aria-current') === 'page')
    const activeThumb = thumbs[activeIndex] ?? null

    return {
      activeIndex,
      activeName: activeThumb?.querySelector('.ppt-thumb-name')?.textContent ?? '',
      count: thumbs.length,
      names: thumbs.map((thumb) => thumb.querySelector('.ppt-thumb-name')?.textContent ?? ''),
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
