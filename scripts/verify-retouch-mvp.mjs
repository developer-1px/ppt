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
const SELECTOR = {
  copyButton: 'button[aria-label="Copy HTML"]',
  downloadButton: 'button[aria-label="Download HTML"]',
  editor: '[data-editing="true"]',
  editorEditable: '[data-editing="true"][contenteditable]',
  modeButton: '.mode-button',
  slideCanvas: '.slide-canvas',
  slideThumb: '.slide-thumb',
  stageShell: '.stage-shell',
}

const checks = []
const browserErrors = []
const CDP_COMMAND_TIMEOUT_MS = 10000
let devServer = null
let chrome = null
let chromeProfile = null

try {
  await ensureAppServer()
  const cdpPort = await getFreePort()
  await launchChrome(cdpPort)
  const page = await openPage(cdpPort, appUrl, {
    width: 1280,
    height: 900,
    deviceScaleFactor: 1,
    mobile: false,
  })

  await runFirstScreenScenario(page)
  await runEditSurfaceParityScenario(page)
  await runCompactEditSurfaceScenario(cdpPort)
  await runTextScenario(page)
  await runLayoutScenario(page)
  await runExportScenario(page, cdpPort)
  await runPersistenceScenario(page)
  await runMobileScenario(cdpPort)

  await page.close()

  const failed = checks.filter((check) => !check.ok)
  const result = {
    checks,
    browserErrors,
    failed,
  }

  console.log(JSON.stringify(result, null, 2))
  process.exitCode = failed.length === 0 && browserErrors.length === 0 ? 0 : 1
} finally {
  if (chrome) await stopChrome()
  if (devServer) await stopDevServer()
  if (chromeProfile) {
    await rm(chromeProfile, {
      recursive: true,
      force: true,
      maxRetries: 5,
      retryDelay: 100,
    })
  }
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
    await Promise.race([exited, delay(1000)])
  }
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

async function launchChrome(port) {
  chromeProfile = await mkdtemp(join(tmpdir(), 'ppt-retouch-chrome-'))
  chrome = spawn(CHROME_BIN, [
    '--headless=new',
    `--remote-debugging-port=${port}`,
    `--user-data-dir=${chromeProfile}`,
    '--no-first-run',
    '--no-default-browser-check',
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
    ws,
    send(method, params = {}) {
      const callId = ++id
      ws.send(JSON.stringify({ id: callId, method, params }))
      return new Promise((resolve, reject) => {
        const timeout = setTimeout(() => {
          pending.delete(callId)
          reject(
            new Error(
              `Timed out CDP command: ${method} ${JSON.stringify(params)}`,
            ),
          )
        }, CDP_COMMAND_TIMEOUT_MS)

        pending.set(callId, { resolve, reject, timeout })
      })
    },
    async eval(expression) {
      const result = await page.send('Runtime.evaluate', {
        expression,
        awaitPromise: true,
        returnByValue: true,
      })
      if (result.exceptionDetails) {
        throw new Error(JSON.stringify(result.exceptionDetails))
      }
      return result.result.value
    },
    async waitFor(expression, timeout = 5000) {
      return waitUntil(async () => page.eval(expression), `Timed out: ${expression}`, timeout)
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
    await page.send('Page.navigate', { url })
  }
  await waitUntil(
    async () =>
      page.eval(
        `location.href === ${JSON.stringify(url)} && document.readyState === 'complete'`,
      ),
    `Timed out loading page: ${url}`,
    15000,
  )
  await page.waitFor(presentSelectorExpression(blockSelector('s1-title')))

  return page
}

async function runFirstScreenScenario(page) {
  const state = await page.eval(`(() => ({
    hasEditorShell: !!document.querySelector('.retouch-app'),
    slideCount: document.querySelectorAll('.slide-thumb').length,
    changedSlideCount: document.querySelectorAll('.slide-thumb[data-changed="true"]').length,
    hasTextMode: !!Array.from(document.querySelectorAll('.mode-button')).find((button) => button.textContent?.trim() === 'Text'),
    hasArrangeMode: !!Array.from(document.querySelectorAll('.mode-button')).find((button) => button.textContent?.trim() === 'Arrange'),
    canvasBackgroundImage: getComputedStyle(document.querySelector(${JSON.stringify(SELECTOR.slideCanvas)})).backgroundImage,
    hasStarterCopy: document.body.textContent.includes('React + TypeScript + Vite'),
    hasStoredDeck: localStorage.getItem('ppt-retouch:v3:deck') !== null,
    hasMainSlideBlock: !!document.querySelector(${JSON.stringify(blockSelector('s1-title'))}),
    ...(() => {
      const viewTabs = Array.from(document.querySelectorAll('.view-toggle [role="tab"]'))
      const selectedViewTabs = viewTabs.filter((tab) => tab.getAttribute('aria-selected') === 'true')
      const selectedViewTab = selectedViewTabs[0] ?? null
      const selectedViewPanel = selectedViewTab?.getAttribute('aria-controls')
        ? document.getElementById(selectedViewTab.getAttribute('aria-controls'))
        : null

      return {
        selectedViewTabCount: selectedViewTabs.length,
        selectedViewTabId: selectedViewTab?.id ?? null,
        viewPanelLabelledBy: selectedViewPanel?.getAttribute('aria-labelledby') ?? null,
        viewPanelRole: selectedViewPanel?.getAttribute('role') ?? null,
        viewTabCount: viewTabs.length,
        viewTablistRole: document.querySelector('.view-toggle')?.getAttribute('role') ?? null,
      }
    })(),
    ...(() => {
      const commandButtons = Array.from(document.querySelectorAll([
        '.zoom-controls button',
        '.toolbar button[data-action="undo"]',
        '.toolbar button[data-action="redo"]',
        '.toolbar button[data-action="reset"]',
        '.toolbar button[data-action="add-text"]',
        '.toolbar button[data-action="present"]',
        '.rail-actions button',
      ].join(',')))
      const pressedLeaks = commandButtons
        .filter((button) => button.hasAttribute('aria-pressed'))
        .map((button) => button.getAttribute('aria-label') ?? button.textContent?.trim() ?? '')

      return {
        actionToolbarLabel: document.querySelector('.toolbar')?.getAttribute('aria-label') ?? null,
        actionToolbarRole: document.querySelector('.toolbar')?.getAttribute('role') ?? null,
        commandPressedLeaks: pressedLeaks,
        copyPressed: document.querySelector('button[data-action="copy-html"]')?.getAttribute('aria-pressed') ?? null,
        downloadPressed: document.querySelector('button[data-action="download-html"]')?.getAttribute('aria-pressed') ?? null,
        modePressedValues: Array.from(document.querySelectorAll('.mode-button'))
          .map((button) => button.getAttribute('aria-pressed')),
        modeToolbarLabel: document.querySelector('.mode-toggle')?.getAttribute('aria-label') ?? null,
        modeToolbarRole: document.querySelector('.mode-toggle')?.getAttribute('role') ?? null,
        railActionToolbarLabel: document.querySelector('.rail-actions')?.getAttribute('aria-label') ?? null,
        railActionToolbarRole: document.querySelector('.rail-actions')?.getAttribute('role') ?? null,
        zoomToolbarLabel: document.querySelector('.zoom-controls')?.getAttribute('aria-label') ?? null,
        zoomToolbarRole: document.querySelector('.zoom-controls')?.getAttribute('role') ?? null,
      }
    })(),
  }))()`)

  check('first screen is retouch editor', state.hasEditorShell && state.hasMainSlideBlock, state)
  check('slide thumbnails are available', state.slideCount >= 3, state)
  check('slide thumbnails start without modified marks', state.changedSlideCount === 0, state)
  check('clean initial deck does not create autosave', !state.hasStoredDeck, state)
  check('mode toggle is available', state.hasTextMode && state.hasArrangeMode, state)
  check(
    'sidebar view switch follows APG tabs',
    state.viewTablistRole === 'tablist' &&
      state.viewTabCount === 2 &&
      state.selectedViewTabCount === 1 &&
      state.viewPanelRole === 'tabpanel' &&
      state.viewPanelLabelledBy === state.selectedViewTabId,
    state,
  )
  check(
    'toolbar groups follow APG toolbar semantics',
    state.modeToolbarRole === 'toolbar' &&
      state.modeToolbarLabel === 'Mode' &&
      state.zoomToolbarRole === 'toolbar' &&
      state.zoomToolbarLabel === 'Canvas zoom' &&
      state.actionToolbarRole === 'toolbar' &&
      state.actionToolbarLabel === 'Actions' &&
      state.railActionToolbarRole === 'toolbar' &&
      state.railActionToolbarLabel === 'Slide actions',
    state,
  )
  check(
    'command toolbar buttons do not expose toggle state',
    state.commandPressedLeaks.length === 0 &&
      state.modePressedValues.includes('true') &&
      state.modePressedValues.includes('false') &&
      state.copyPressed === 'false' &&
      state.downloadPressed === 'false',
    state,
  )
  check('Text Mode starts as clean slide preview', state.canvasBackgroundImage === 'none', state)
  check('Vite starter copy is removed', !state.hasStarterCopy, state)
}

async function runEditSurfaceParityScenario(page) {
  const deltas = []

  for (const slideName of ['Overview', 'Agenda', 'Decision']) {
    await clickSlide(page, slideName)
    const blockIds = await page.eval(
      "Array.from(document.querySelectorAll('[data-block]')).map((block) => block.dataset.block)",
    )

    for (const blockId of blockIds) {
      const preview = await textRangeMetrics(page, `[data-block="${blockId}"]`)
      await page.eval(`document.querySelector('[data-block="${blockId}"]').click()`)
      await waitForEditor(page)
      const blockAfter = await textRangeMetrics(page, `[data-block="${blockId}"]`)
      const editor = await textRangeMetrics(page, SELECTOR.editor)
      const editorChrome = await page.eval(`(() => {
        const style = getComputedStyle(document.querySelector(${JSON.stringify(SELECTOR.editor)}))

        return {
          outlineOffset: style.outlineOffset,
          outlineStyle: style.outlineStyle,
          outlineWidth: style.outlineWidth,
        }
      })()`)

      deltas.push({
        blockId,
        slideName,
        boxHeightDelta: blockAfter.boxHeight - preview.boxHeight,
        boxLeftDelta: blockAfter.boxLeft - preview.boxLeft,
        boxTopDelta: blockAfter.boxTop - preview.boxTop,
        canvasLeftDelta: blockAfter.canvasLeft - preview.canvasLeft,
        canvasTopDelta: blockAfter.canvasTop - preview.canvasTop,
        outlineOffset: editorChrome.outlineOffset,
        outlineStyle: editorChrome.outlineStyle,
        outlineWidth: editorChrome.outlineWidth,
        previewTextInsideBox: preview.textInsideBox,
        editorTextInsideBox: editor.textInsideBox,
        stageScrollLeftDelta: blockAfter.stageScrollLeft - preview.stageScrollLeft,
        stageScrollTopDelta: blockAfter.stageScrollTop - preview.stageScrollTop,
        textHeightDelta: editor.textHeight - preview.textHeight,
        textLeftDelta: editor.textLeft - preview.textLeft,
        textTopDelta: editor.textTop - preview.textTop,
      })

      await cancelTextEditor(page)
    }
  }

  await clickSlide(page, 'Overview')

  check(
    'Text Mode editor opens on the same text without boxed chrome',
    deltas.every(
      (delta) =>
        Math.abs(delta.boxHeightDelta) < 1 &&
        Math.abs(delta.boxLeftDelta) < 1 &&
        Math.abs(delta.boxTopDelta) < 1 &&
        Math.abs(delta.canvasLeftDelta) < 1 &&
        Math.abs(delta.canvasTopDelta) < 1 &&
        delta.outlineOffset === '0px' &&
        delta.outlineStyle === 'none' &&
        delta.outlineWidth === '0px' &&
        delta.previewTextInsideBox &&
        delta.editorTextInsideBox &&
        Math.abs(delta.stageScrollLeftDelta) < 1 &&
        Math.abs(delta.stageScrollTopDelta) < 1 &&
        Math.abs(delta.textHeightDelta) < 1 &&
        Math.abs(delta.textLeftDelta) < 1 &&
        Math.abs(delta.textTopDelta) < 1,
    ),
    deltas,
  )
}

async function runCompactEditSurfaceScenario(cdpPort) {
  const page = await openPage(cdpPort, appUrl, {
    width: 1188,
    height: 635,
    deviceScaleFactor: 1,
    mobile: false,
  })
  const deltas = []

  for (const blockId of ['s1-title', 's1-subtitle', 's1-note']) {
    const preview = await textRangeMetrics(page, `[data-block="${blockId}"]`)
    await clickAt(page, {
      x: preview.textLeft + preview.textWidth / 2,
      y: preview.textTop + preview.textHeight / 2,
    })
    await page.waitFor(`document.querySelector('[data-editing=\"true\"]')?.dataset.editingBlock === '${blockId}'`)
    const blockAfter = await textRangeMetrics(page, `[data-block="${blockId}"]`)
    const editor = await textRangeMetrics(page, '[data-editing=\"true\"]')

    deltas.push({
      blockId,
      boxTopDelta: blockAfter.boxTop - preview.boxTop,
      canvasTopDelta: blockAfter.canvasTop - preview.canvasTop,
      previewTextInsideBox: preview.textInsideBox,
      editorTextInsideBox: editor.textInsideBox,
      stageScrollLeftDelta: blockAfter.stageScrollLeft - preview.stageScrollLeft,
      stageScrollTopDelta: blockAfter.stageScrollTop - preview.stageScrollTop,
      textLeftDelta: editor.textLeft - preview.textLeft,
      textTopDelta: editor.textTop - preview.textTop,
    })

    await cancelTextEditor(page)
  }

  check(
    'Text Mode compact viewport edit entry has no visual gap',
    deltas.every(
      (delta) =>
        Math.abs(delta.boxTopDelta) < 1 &&
        Math.abs(delta.canvasTopDelta) < 1 &&
        delta.previewTextInsideBox &&
        delta.editorTextInsideBox &&
        Math.abs(delta.stageScrollLeftDelta) < 1 &&
        Math.abs(delta.stageScrollTopDelta) < 1 &&
        Math.abs(delta.textLeftDelta) < 1 &&
        Math.abs(delta.textTopDelta) < 1,
    ),
    deltas,
  )
  await page.close()
}

async function runTextScenario(page) {
  const titleBefore = await blockState(page, 's1-title')
  const titlePreviewText = await textRangeMetrics(page, '[data-block="s1-title"]')
  const titleHoverPoint = await blockCenter(page, 's1-title')

  await page.send('Input.dispatchMouseEvent', {
    type: 'mouseMoved',
    x: titleHoverPoint.x,
    y: titleHoverPoint.y,
    button: 'none',
  })
  await delay(50)
  const titleHoverChrome = await blockChrome(page, 's1-title')
  check(
    'Text Mode hover does not draw a text box',
    titleHoverChrome.outlineStyle === 'none',
    titleHoverChrome,
  )
  await movePointerAway(page)

  await page.eval(`document.querySelector('[data-block="s1-title"]').click()`)
  await waitForEditor(page)
  await commitTextEditor(page)
  const titleNoOp = await blockState(page, 's1-title')
  check('no-op text edit does not create history', titleNoOp.text === titleBefore.text && titleNoOp.undoDisabled === true, titleNoOp)

  const titleClickText = await textRangeMetrics(page, '[data-block="s1-title"]')
  await clickAt(page, {
    x: titleClickText.textLeft + 2,
    y: titleClickText.textTop + titleClickText.textHeight / 2,
  })
  await waitForEditor(page)
  await typeEditorText(page, 'Q')
  const clickCaretText = await page.eval("document.querySelector('[data-editing=\"true\"]')?.textContent ?? ''")
  const clickCaretIndex = clickCaretText.indexOf('Q')
  check(
    'Text Mode starts editing at clicked text position',
    clickCaretIndex >= 0 && clickCaretIndex < titleBefore.text.length,
    { clickCaretText, clickCaretIndex },
  )
  const koreanBeforeInput = await insertEditorTextCommand(page, 'ㅇ')
  check(
    'Text Mode Korean insertText inserts one character',
    koreanBeforeInput.after === `${koreanBeforeInput.before}ㅇ`,
    koreanBeforeInput,
  )
  await cancelTextEditor(page)
  await movePointerAway(page)

  await focusBlockAndPress(page, 's1-title', 'Enter')
  await waitForEditor(page)
  const keyboardEditState = await page.eval(`(() => ({
    editorOpen: !!document.querySelector('[data-editing=\"true\"]'),
    editorText: document.querySelector('[data-editing=\"true\"]')?.textContent ?? null,
  }))()`)
  check('keyboard Enter starts text edit', keyboardEditState.editorOpen && keyboardEditState.editorText === titleBefore.text, keyboardEditState)
  await cancelTextEditor(page)

  await page.eval(`document.querySelector('[data-block="s1-title"]').click()`)
  await waitForEditor(page)
  const titleEditorText = await textRangeMetrics(page, '[data-editing=\"true\"]')
  const titleEditorBox = await editorMetrics(page)
  const editingTitle = await page.eval(`(() => ({
    originalBlockCount: document.querySelectorAll('[data-block="s1-title"]').length,
    editorCount: document.querySelectorAll('[data-editing=\"true\"][contenteditable]').length,
    editsOriginalBlockText:
      document.querySelector('[data-editing=\"true\"]')?.closest('[data-block="s1-title"]') ===
      document.querySelector('[data-block="s1-title"]'),
    editorText: document.querySelector('[data-editing=\"true\"]')?.textContent,
    contentEditable: document.querySelector('[data-editing=\"true\"]')?.contentEditable,
    hasMarkdownChrome:
      document.body.textContent.includes('**') ||
      document.body.textContent.includes('H1') ||
      document.body.textContent.includes('MD') ||
      !!document.querySelector('.ProseMirror, .nano, .block-picker, .command-palette'),
  }))()`)
  check(
    'Text Mode edits the original block directly',
    editingTitle.originalBlockCount === 1 &&
      editingTitle.editorCount === 1 &&
      editingTitle.editsOriginalBlockText &&
      editingTitle.contentEditable === 'plaintext-only',
    editingTitle,
  )
  check('Text Mode does not show markdown editor chrome', !editingTitle.hasMarkdownChrome, editingTitle)
  check('Text Mode non-empty editor does not inherit source min-height', titleEditorBox.minHeight === '0px', titleEditorBox)
  check(
    'Text Mode editor keeps preview text position',
    Math.abs(titleEditorText.textTop - titlePreviewText.textTop) < 1 &&
      Math.abs(titleEditorText.textLeft - titlePreviewText.textLeft) < 1 &&
      Math.abs(titleEditorText.textHeight - titlePreviewText.textHeight) < 1,
    { before: titlePreviewText, after: titleEditorText },
  )
  check(
    'Text Mode edit surface hugs visible text',
    Math.abs(titleEditorText.boxLeft - titleEditorText.textLeft) < 1 &&
      Math.abs(titleEditorText.boxTop - titleEditorText.textTop) < 1 &&
      Math.abs(titleEditorText.boxWidth - titleEditorText.textWidth) < 1 &&
      Math.abs(titleEditorText.boxHeight - titleEditorText.textHeight) < 1,
    titleEditorText,
  )

  await pressEditorEnter(page)
  const titleLineBreakEdit = await page.eval(`(() => ({
    editorOpen: !!document.querySelector('[data-editing=\"true\"]'),
    text: document.querySelector('[data-editing=\"true\"]')?.textContent ?? '',
  }))()`)
  check(
    'plain Enter inserts a text line break without closing editor',
    titleLineBreakEdit.editorOpen && titleLineBreakEdit.text.includes('\n'),
    titleLineBreakEdit,
  )
  await typeEditorText(page, 'Next line')
  await commitTextEditor(page)
  const titleLineBreakCommit = await blockState(page, 's1-title')
  check(
    'committed preview keeps plain Enter line break',
    titleLineBreakCommit.text === `${titleBefore.text}\nNext line` &&
      !titleLineBreakCommit.editorOpen,
    titleLineBreakCommit,
  )
  await clickToolbar(page, 'Undo')
  await page.waitFor(`document.querySelector('[data-block="s1-title"]')?.textContent === ${JSON.stringify(titleBefore.text)}`)

  await page.eval(`document.querySelector('[data-block="s1-title"]').click()`)
  await waitForEditor(page)
  await typeEditorText(page, ' Approved')
  await commitTextEditor(page)
  await movePointerAway(page)
  const titleCommitted = await blockState(page, 's1-title')
  check('title text commit works', titleCommitted.text === `${titleBefore.text} Approved`, titleCommitted)
  check('keyboard shortcut commits text edit', !titleCommitted.editorOpen, titleCommitted)
  const titleChromeAfterCommit = await blockChrome(page, 's1-title')
  check('Text Mode returns to clean preview after commit', titleChromeAfterCommit.outlineStyle === 'none', titleChromeAfterCommit)
  check('text commit enables undo', titleCommitted.undoDisabled === false, titleCommitted)

  await clickToolbar(page, 'Undo')
  await delay(150)
  const titleUndone = await blockState(page, 's1-title')
  check('undo restores title text', titleUndone.text === titleBefore.text, titleUndone)

  await clickToolbar(page, 'Redo')
  await delay(150)
  const titleRedone = await blockState(page, 's1-title')
  check('redo restores title edit', titleRedone.text === `${titleBefore.text} Approved`, titleRedone)

  await pressHistoryShortcut(page, 'z')
  await delay(150)
  const titleKeyboardUndone = await blockState(page, 's1-title')
  check('keyboard undo restores title text', titleKeyboardUndone.text === titleBefore.text, titleKeyboardUndone)

  await pressHistoryShortcut(page, 'y')
  await delay(150)
  const titleKeyboardRedone = await blockState(page, 's1-title')
  check('keyboard redo restores title edit', titleKeyboardRedone.text === `${titleBefore.text} Approved`, titleKeyboardRedone)

  await clickSlide(page, 'Agenda')
  const bodyBefore = await blockState(page, 's2-footer')
  const bodyText = await textRangeMetrics(page, '[data-block="s2-footer"]')
  const shortenedBodyText = 'Decide faster.'
  await clickAt(page, {
    x: bodyText.textLeft + bodyText.textWidth / 2,
    y: bodyText.textTop + bodyText.textHeight / 2,
  })
  await page.waitFor("document.querySelector('[data-editing=\"true\"]')?.dataset.editingBlock === 's2-footer'")
  await replaceEditorText(page, shortenedBodyText)
  await commitTextEditor(page)
  const bodyShortened = await blockState(page, 's2-footer')
  check(
    'Text Mode shortens a body sentence without moving it',
    bodyBefore.role === 'body' &&
      bodyShortened.text === shortenedBodyText &&
      bodyShortened.text.length < bodyBefore.text.length &&
      Math.abs(bodyShortened.x - bodyBefore.x) < 1 &&
      Math.abs(bodyShortened.y - bodyBefore.y) < 1 &&
      Math.abs(bodyShortened.width - bodyBefore.width) < 1 &&
      bodyShortened.height <= bodyBefore.height + 1,
    { before: bodyBefore, after: bodyShortened },
  )
  await clickToolbar(page, 'Undo')
  await page.waitFor(`document.querySelector('[data-block="s2-footer"]')?.textContent === ${JSON.stringify(bodyBefore.text)}`)
  const bodyShortenUndone = await blockState(page, 's2-footer')
  check('undo restores shortened body sentence', bodyShortenUndone.text === bodyBefore.text, bodyShortenUndone)
  await clickToolbar(page, 'Redo')
  await page.waitFor(`document.querySelector('[data-block="s2-footer"]')?.textContent === ${JSON.stringify(shortenedBodyText)}`)
  const bodyShortenRedone = await blockState(page, 's2-footer')
  check('redo reapplies shortened body sentence', bodyShortenRedone.text === shortenedBodyText, bodyShortenRedone)
  await clickSlide(page, 'Overview')

  await page.eval(`document.querySelector('[data-block="s1-title"]').click()`)
  await waitForEditor(page)
  await typeEditorText(page, ' ToolbarUndo')
  await clickToolbar(page, 'Undo')
  await page.waitFor(`!document.querySelector('[data-editing=\"true\"]') && document.querySelector('[data-block="s1-title"]')?.textContent === ${JSON.stringify(`${titleBefore.text} Approved`)}`)
  const toolbarUndoState = await blockState(page, 's1-title')
  check(
    'toolbar Undo commits live edit before undoing it',
    toolbarUndoState.text === `${titleBefore.text} Approved` &&
      !toolbarUndoState.editorOpen &&
      toolbarUndoState.redoDisabled === false,
    toolbarUndoState,
  )

  await clickToolbar(page, 'Redo')
  await page.waitFor(`document.querySelector('[data-block="s1-title"]')?.textContent === ${JSON.stringify(`${titleBefore.text} Approved ToolbarUndo`)}`)
  const toolbarRedoState = await blockState(page, 's1-title')
  check(
    'toolbar Redo restores text undone from a live edit',
    toolbarRedoState.text === `${titleBefore.text} Approved ToolbarUndo` &&
      !toolbarRedoState.editorOpen,
    toolbarRedoState,
  )
  await clickToolbar(page, 'Undo')
  await page.waitFor(`document.querySelector('[data-block="s1-title"]')?.textContent === ${JSON.stringify(`${titleBefore.text} Approved`)}`)

  await focusBlockAndPress(page, 's1-title', 'Enter')
  await page.waitFor("document.querySelector('[data-editing=\"true\"]')?.dataset.editingBlock === 's1-title'")
  await typeEditorText(page, ' ChainTitle')
  const chainSubtitleBefore = await blockState(page, 's1-subtitle')
  const chainSubtitleText = await textRangeMetrics(page, '[data-block="s1-subtitle"]')
  await clickAt(page, {
    x: chainSubtitleText.textLeft + chainSubtitleText.textWidth / 2,
    y: chainSubtitleText.textTop + chainSubtitleText.textHeight / 2,
  })
  await page.waitFor("document.querySelector('[data-editing=\"true\"]')?.dataset.editingBlock === 's1-subtitle'")
  const chainedTitleCommit = await blockState(page, 's1-title')
  const chainedEditor = await page.eval(`(() => ({
    blockId: document.querySelector('[data-editing=\"true\"]')?.dataset.editingBlock ?? null,
    text: document.querySelector('[data-editing=\"true\"]')?.textContent ?? null,
  }))()`)
  check(
    'Text Mode commits current block before editing another text block',
    chainedTitleCommit.text === `${titleBefore.text} Approved ChainTitle` &&
      chainedEditor.blockId === 's1-subtitle' &&
      chainedEditor.text === chainSubtitleBefore.text,
    { title: chainedTitleCommit, editor: chainedEditor },
  )

  await typeEditorText(page, ' ChainSubtitle')
  await commitTextEditor(page)
  const chainedSubtitleCommit = await blockState(page, 's1-subtitle')
  check(
    'Text Mode commits the second block in consecutive edits',
    chainedSubtitleCommit.text.includes('ChainSubtitle') &&
      chainedSubtitleCommit.text !== chainSubtitleBefore.text,
    { before: chainSubtitleBefore, after: chainedSubtitleCommit },
  )
  await clickToolbar(page, 'Undo')
  await page.waitFor(`!document.querySelector('[data-block="s1-subtitle"]')?.textContent.includes('ChainSubtitle')`)
  await clickToolbar(page, 'Undo')
  await page.waitFor(`document.querySelector('[data-block="s1-title"]')?.textContent === ${JSON.stringify(`${titleBefore.text} Approved`)}`)

  await page.eval(`document.querySelector('[data-block="s1-title"]').click()`)
  await waitForEditor(page)
  await typeEditorText(page, ' SlideSwitch')
  await clickSlide(page, 'Agenda')
  await clickSlide(page, 'Overview')
  const slideSwitchCommit = await blockState(page, 's1-title')
  check(
    'Text Mode commits live edit before slide switch',
    slideSwitchCommit.text === `${titleBefore.text} Approved SlideSwitch` &&
      !slideSwitchCommit.editorOpen,
    slideSwitchCommit,
  )
  await clickToolbar(page, 'Undo')
  await page.waitFor(`document.querySelector('[data-block="s1-title"]')?.textContent === ${JSON.stringify(`${titleBefore.text} Approved`)}`)

  await page.eval(`document.querySelector('[data-block="s1-title"]').click()`)
  await waitForEditor(page)
  await typeEditorText(page, ' ModeSwitch')
  await clickMode(page, 'Arrange')
  const modeSwitchCommit = await blockState(page, 's1-title')
  const modeAfterLiveEdit = await page.eval("document.querySelector('.retouch-app')?.dataset.mode ?? null")
  check(
    'Text Mode commits live edit before Arrange switch',
    modeSwitchCommit.text === `${titleBefore.text} Approved ModeSwitch` &&
      modeAfterLiveEdit === 'layout' &&
      !modeSwitchCommit.editorOpen,
    { state: modeSwitchCommit, mode: modeAfterLiveEdit },
  )
  await clickToolbar(page, 'Undo')
  await page.waitFor(`document.querySelector('[data-block="s1-title"]')?.textContent === ${JSON.stringify(`${titleBefore.text} Approved`)}`)
  await clickMode(page, 'Text')

  const subtitleBefore = await blockState(page, 's1-subtitle')
  await page.eval(`document.querySelector('[data-block="s1-subtitle"]').click()`)
  await waitForEditor(page)
  const editBefore = await editorMetrics(page)
  await typeEditorText(
    page,
    ' Extra context that wraps onto more lines so this fixed-width text box grows vertically instead of clipping or scrolling. Add owner names, renewal risk, partner follow-up, timing, decision path, and executive review notes so the content clearly exceeds the original text box height.',
  )
  await page.waitFor(`(() => {
    const editor = document.querySelector('[data-editing=\"true\"]')
    return editor && editor.getBoundingClientRect().height > ${editBefore.height + 10}
  })()`)
  const autoHeight = await editorMetrics(page)
  const autoHeightText = await textRangeMetrics(page, '[data-editing="true"]')
  check('Text Mode autoheight grows box, not scroll/clip', autoHeight.height > editBefore.height + 10 && autoHeight.wrapperOverflowY === 'visible' && autoHeight.editorOverflowY === 'visible', autoHeight)
  check('Text Mode autoheight keeps content top stable while typing', Math.abs(autoHeight.top - editBefore.top) < 1 && Math.abs(autoHeight.textTop - editBefore.textTop) < 1, { before: editBefore, after: autoHeight })

  await commitTextEditor(page)
  const subtitleGrown = await blockState(page, 's1-subtitle')
  const subtitleGrownText = await textRangeMetrics(page, '[data-block="s1-subtitle"]')
  check('autoheight persists after commit', subtitleGrown.height > subtitleBefore.height + 10, { before: subtitleBefore, after: subtitleGrown })
  check(
    'committed preview matches live editor text box',
    Math.abs(subtitleGrownText.boxTop - autoHeightText.boxTop) < 1 &&
      Math.abs(subtitleGrownText.boxHeight - autoHeightText.boxHeight) < 1 &&
      Math.abs(subtitleGrownText.textTop - autoHeightText.textTop) < 1 &&
      Math.abs(subtitleGrownText.textLeft - autoHeightText.textLeft) < 1 &&
      Math.abs(subtitleGrownText.textHeight - autoHeightText.textHeight) < 1,
    { editing: autoHeightText, committed: subtitleGrownText },
  )

  await clickToolbar(page, 'Undo')
  await delay(150)
  const subtitleUndone = await blockState(page, 's1-subtitle')
  check('undo restores text and autoheight', subtitleUndone.text === subtitleBefore.text && Math.abs(subtitleUndone.height - subtitleBefore.height) < 1, { before: subtitleBefore, after: subtitleUndone })

  await clickToolbar(page, 'Redo')
  await delay(150)
  const subtitleRedone = await blockState(page, 's1-subtitle')
  check('redo restores autoheight growth', subtitleRedone.text === subtitleGrown.text && Math.abs(subtitleRedone.height - subtitleGrown.height) < 1, { before: subtitleGrown, after: subtitleRedone })

  await clickMode(page, 'Arrange')
  const autoheightOnlyLayoutState = await blockState(page, 's1-subtitle')
  check('Arrange Reset ignores text autoheight only changes', autoheightOnlyLayoutState.resetDisabled === true, autoheightOnlyLayoutState)
  await clickMode(page, 'Text')

  await page.eval(`document.querySelector('[data-block="s1-subtitle"]').click()`)
  await waitForEditor(page)
  const grownEditHeight = await editorHeight(page)
  await replaceEditorText(page, 'Short follow-up.')
  await page.waitFor(`(() => {
    const editor = document.querySelector('[data-editing=\"true\"]')
    return editor && editor.getBoundingClientRect().height < ${grownEditHeight - 10}
  })()`)
  const shrinkHeight = await editorHeight(page)
  check('Text Mode autoheight shrinks with shorter content', shrinkHeight < grownEditHeight - 10, { grownEditHeight, shrinkHeight })

  await commitTextEditor(page)
  const subtitleShrunk = await blockState(page, 's1-subtitle')
  check('autoheight shrink persists after commit', subtitleShrunk.text === 'Short follow-up.' && subtitleShrunk.height < subtitleGrown.height - 10, { before: subtitleGrown, after: subtitleShrunk })

  await page.eval(`document.querySelector('[data-block="s1-subtitle"]').click()`)
  await waitForEditor(page)
  await replaceEditorText(
    page,
    'enterprise-renewal-risk-account-owner-review-decision-path-follow-up-needed-before-month-close',
  )
  const longEditorFit = await inlineFits(page, '[data-editing=\"true\"]')
  check('Text Mode wraps long unbroken text while editing', longEditorFit.fits, longEditorFit)
  await commitTextEditor(page)
  const longPreviewFit = await inlineFits(page, '[data-block="s1-subtitle"]')
  check('preview wraps long unbroken text after commit', longPreviewFit.fits, longPreviewFit)

  await page.eval(`document.querySelector('[data-block="s1-subtitle"]').click()`)
  await waitForEditor(page)
  await setEditorText(page, '')
  await commitTextEditor(page)
  const emptySubtitle = await page.eval(`(() => {
    const block = document.querySelector('[data-block="s1-subtitle"]')
    const rect = block.getBoundingClientRect()
    return {
      text: block.textContent,
      empty: block.dataset.empty,
      width: rect.width,
      height: rect.height,
    }
  })()`)
  check('empty text block remains findable', emptySubtitle.text === '' && emptySubtitle.empty === 'true' && emptySubtitle.width > 0 && emptySubtitle.height > 0, emptySubtitle)
  await focusBlockAndPress(page, 's1-subtitle', 'Enter')
  await waitForEditor(page)
  const emptyReopen = await page.eval(`document.querySelector('[data-editing=\"true\"]')?.textContent ?? null`)
  check('empty text block can be reopened from keyboard', emptyReopen === '', { emptyReopen })
  await cancelTextEditor(page)

  await page.eval(`document.querySelector('[data-block="s1-note"]').click()`)
  await waitForEditor(page)
  const noteBeforeCancel = await page.eval(`document.querySelector('[data-editing=\"true\"]')?.textContent ?? ''`)
  await typeEditorText(page, ' Cancelled draft')
  await cancelTextEditor(page)
  const noteAfterCancel = await blockState(page, 's1-note')
  check('Escape cancels text draft', noteAfterCancel.text === noteBeforeCancel, noteAfterCancel)

  await page.eval(`document.querySelector('[data-block="s1-note"]').click()`)
  await waitForEditor(page)
  await typeEditorText(page, ' Mistyped')
  await pressHistoryShortcut(page, 'z')
  const noteAfterLiveUndo = await page.eval(`(() => {
    const editor = document.querySelector('[data-editing=\"true\"]')

    return {
      editorOpen: !!editor,
      text: editor?.textContent ?? '',
    }
  })()`)
  check(
    'Text Mode keyboard undo reverts live text draft',
    noteAfterLiveUndo.editorOpen && noteAfterLiveUndo.text === noteBeforeCancel,
    noteAfterLiveUndo,
  )
  await cancelTextEditor(page)

  const subtitleBeforeTextDrag = await blockState(page, 's1-subtitle')
  const textModeResizeHandles = await page.eval(
    "document.querySelectorAll('.resize-handle').length",
  )
  await dragBlock(page, 's1-subtitle', 48, 24)
  const subtitleAfterTextDrag = await blockState(page, 's1-subtitle')
  const editorAfterTextDrag = await page.eval(
    "!!document.querySelector('[data-editing=\"true\"]')",
  )
  check(
    'Text Mode drag attempt does not change layout',
    textModeResizeHandles === 0 &&
      !rectChanged(subtitleBeforeTextDrag, subtitleAfterTextDrag) &&
      subtitleAfterTextDrag.text === subtitleBeforeTextDrag.text,
    {
      before: subtitleBeforeTextDrag,
      after: subtitleAfterTextDrag,
      editorOpen: editorAfterTextDrag,
      resizeHandles: textModeResizeHandles,
    },
  )
  if (editorAfterTextDrag) {
    await cancelTextEditor(page)
  }

  await clickSlide(page, 'Decision')
  const bottomNoteBefore = await blockState(page, 's3-note')
  await page.eval(`document.querySelector('[data-block="s3-note"]').click()`)
  await waitForEditor(page)
  const bottomNoteDraftText =
    'Decision needed: approve the retention sprint with owner, timing, customer impact, renewal coverage, executive sponsor, next action, weekly check-in, risk review, onboarding help, and discount guidance visible in the slide.'
  await replaceEditorText(
    page,
    bottomNoteDraftText,
  )
  await page.waitFor(`(() => {
    const editor = document.querySelector('[data-editing=\"true\"]')
    return editor && editor.getBoundingClientRect().height > ${bottomNoteBefore.height + 10}
  })()`)
  const bottomEditorFit = await editorFitsSlide(page)
  check('autoheight keeps bottom editor inside slide while typing', bottomEditorFit.fits, bottomEditorFit)
  await commitTextEditor(page)
  const bottomNoteAfter = await blockState(page, 's3-note')
  const bottomNoteFit = await blockFitsSlide(page, 's3-note')
  check('autoheight keeps bottom text box inside slide', bottomNoteFit.fits && bottomNoteAfter.height > bottomNoteBefore.height + 10, { before: bottomNoteBefore, after: bottomNoteAfter, fit: bottomNoteFit })

  await clickSlide(page, 'Overview')
  const stageScrollTop = await page.eval("document.querySelector('.stage-shell')?.scrollTop ?? null")
  check('slide switch resets stage scroll', stageScrollTop === 0, { stageScrollTop })
}

async function runLayoutScenario(page) {
  await clickMode(page, 'Arrange')
  const arrangeSurface = await page.eval(`(() => ({
    canvasBackgroundImage: getComputedStyle(document.querySelector('.slide-canvas')).backgroundImage,
  }))()`)
  check('Arrange Mode shows layout grid only when arranging', arrangeSurface.canvasBackgroundImage !== 'none', arrangeSurface)

  await clickSlide(page, 'Agenda')
  const centerSnapProbe = await dragBlockBySlideUnitsWithProbe(page, 's2-step-1', 337, 0)
  const centerSnapped = await blockCenterAlignment(page, 's2-step-1')
  check(
    'Arrange Mode snaps blocks to the slide center',
    centerSnapProbe.hasCenterGuideX && centerSnapped.centeredX,
    { probe: centerSnapProbe, alignment: centerSnapped },
  )
  await clickToolbar(page, 'Undo')
  await delay(150)

  const siblingSnapProbe = await dragBlockBySlideUnitsWithProbe(page, 's2-step-1', 599, 0)
  const siblingEdgeAligned = await blockEdgeAlignment(page, 's2-step-1', 's2-step-2')
  check(
    'Arrange Mode snaps blocks to sibling edges',
    siblingSnapProbe.hasGuideXAt770 && siblingEdgeAligned.leftToPeerRight,
    { probe: siblingSnapProbe, alignment: siblingEdgeAligned },
  )
  await clickToolbar(page, 'Undo')
  await delay(150)

  const stepBeforeNudge = await blockState(page, 's2-step-1')
  await clickBlock(page, 's2-step-1')
  await pressKey(page, 'ArrowRight')
  await delay(150)
  const stepNudged = await blockState(page, 's2-step-1')
  check(
    'Arrange Mode arrow keys nudge the selected block',
    stepNudged.x > stepBeforeNudge.x + 4 &&
      Math.abs(stepNudged.y - stepBeforeNudge.y) < 1 &&
      stepNudged.text === stepBeforeNudge.text,
    { before: stepBeforeNudge, after: stepNudged },
  )
  await clickToolbar(page, 'Undo')
  await delay(150)
  const stepNudgeUndone = await blockState(page, 's2-step-1')
  check(
    'undo restores keyboard nudged layout',
    !rectChanged(stepBeforeNudge, stepNudgeUndone) &&
      stepNudgeUndone.text === stepBeforeNudge.text,
    { before: stepBeforeNudge, after: stepNudgeUndone },
  )

  await clickBlock(page, 's2-step-1')
  await pressKey(page, 'ArrowRight', { shift: true })
  await delay(150)
  const stepLargeNudged = await blockState(page, 's2-step-1')
  check(
    'Arrange Mode Shift+Arrow nudges the selected block by a larger step',
    stepLargeNudged.x > stepBeforeNudge.x + 20 &&
      Math.abs(stepLargeNudged.y - stepBeforeNudge.y) < 1 &&
      stepLargeNudged.text === stepBeforeNudge.text,
    { before: stepBeforeNudge, after: stepLargeNudged },
  )
  await clickToolbar(page, 'Undo')
  await delay(150)
  const stepLargeNudgeUndone = await blockState(page, 's2-step-1')
  check(
    'undo restores large keyboard nudged layout',
    !rectChanged(stepBeforeNudge, stepLargeNudgeUndone) &&
      stepLargeNudgeUndone.text === stepBeforeNudge.text,
    { before: stepBeforeNudge, after: stepLargeNudgeUndone },
  )

  const multiStepOneBefore = await blockState(page, 's2-step-1')
  const multiStepTwoBefore = await blockState(page, 's2-step-2')
  await clickBlock(page, 's2-step-1')
  await clickBlock(page, 's2-step-2', { shift: true })
  await delay(100)
  const multiSelectState = await page.eval(`(() => ({
    selectedBlocks: document.querySelectorAll('[data-selected="true"]').length,
    resizeHandles: document.querySelectorAll('.resize-handle').length,
    stepOneSelected:
      document.querySelector('[data-block="s2-step-1"]')?.dataset.selected,
    stepTwoSelected:
      document.querySelector('[data-block="s2-step-2"]')?.dataset.selected,
  }))()`)
  check(
    'Arrange Mode Shift-click selects multiple blocks',
    multiSelectState.selectedBlocks === 2 &&
      multiSelectState.resizeHandles === 0 &&
      multiSelectState.stepOneSelected === 'true' &&
      multiSelectState.stepTwoSelected === 'true',
    multiSelectState,
  )

  await dragBlock(page, 's2-step-1', 42, 24)
  const multiStepOneMoved = await blockState(page, 's2-step-1')
  const multiStepTwoMoved = await blockState(page, 's2-step-2')
  const multiMoveSelection = await page.eval(
    `document.querySelectorAll('[data-selected="true"]').length`,
  )
  check(
    'Arrange Mode drag moves multi-selected blocks together',
    rectChanged(multiStepOneBefore, multiStepOneMoved) &&
      rectChanged(multiStepTwoBefore, multiStepTwoMoved) &&
      multiStepOneMoved.text === multiStepOneBefore.text &&
      multiStepTwoMoved.text === multiStepTwoBefore.text &&
      multiMoveSelection === 2,
    {
      before: { one: multiStepOneBefore, two: multiStepTwoBefore },
      after: { one: multiStepOneMoved, two: multiStepTwoMoved },
      selectedBlocks: multiMoveSelection,
    },
  )
  await clickToolbar(page, 'Undo')
  await delay(150)
  const multiStepOneUndone = await blockState(page, 's2-step-1')
  const multiStepTwoUndone = await blockState(page, 's2-step-2')
  check(
    'undo restores multi-selected drag',
    !rectChanged(multiStepOneBefore, multiStepOneUndone) &&
      !rectChanged(multiStepTwoBefore, multiStepTwoUndone),
    {
      before: { one: multiStepOneBefore, two: multiStepTwoBefore },
      after: { one: multiStepOneUndone, two: multiStepTwoUndone },
    },
  )
  await clickSlide(page, 'Overview')

  const layerOrderBefore = await slideBlockOrder(page)
  const layerOrderTarget = layerOrderBefore[0]
  await clickBlock(page, layerOrderTarget)
  await clickToolbar(page, 'Bring to front')
  await page.waitFor(
    `Array.from(document.querySelectorAll('.slide-canvas [data-block]')).at(-1)?.dataset.block === ${JSON.stringify(layerOrderTarget)}`,
  )
  const layerOrderAfter = await slideBlockOrder(page)
  check(
    'Arrange Mode brings selected block to front',
    layerOrderBefore.length === layerOrderAfter.length &&
      layerOrderBefore[0] === layerOrderTarget &&
      layerOrderAfter.at(-1) === layerOrderTarget,
    { before: layerOrderBefore, after: layerOrderAfter, target: layerOrderTarget },
  )
  await clickToolbar(page, 'Undo')
  await delay(150)
  const layerOrderUndone = await slideBlockOrder(page)
  check(
    'undo restores layer order',
    sameArray(layerOrderBefore, layerOrderUndone),
    { before: layerOrderBefore, after: layerOrderUndone },
  )

  const noteBefore = await blockState(page, 's1-note')

  await dragBlock(page, 's1-note', 5, 5)
  const noteAfterSmallDrag = await blockState(page, 's1-note')
  check('Arrange Mode ignores tiny accidental drag', !rectChanged(noteBefore, noteAfterSmallDrag), { before: noteBefore, after: noteAfterSmallDrag })

  await dragBlock(page, 's1-note', 42, 24)
  const noteMoved = await blockState(page, 's1-note')
  check('Arrange Mode drag moves block only', rectChanged(noteBefore, noteMoved) && noteMoved.text === noteBefore.text, { before: noteBefore, after: noteMoved })
  check(
    'Arrange Mode reset target is explicit',
    noteMoved.resetDisabled === false &&
      noteMoved.resetScope === 'layout' &&
      noteMoved.resetAriaLabel === 'Reset layout' &&
      noteMoved.resetTitle === 'Reset layout',
    noteMoved,
  )

  await clickToolbar(page, 'Undo')
  await delay(150)
  const noteMoveUndone = await blockState(page, 's1-note')
  check('undo restores dragged layout', !rectChanged(noteBefore, noteMoveUndone) && noteMoveUndone.text === noteBefore.text, { before: noteBefore, after: noteMoveUndone })

  await clickToolbar(page, 'Redo')
  await delay(150)
  const noteMoveRedone = await blockState(page, 's1-note')
  check('redo restores dragged layout', rectChanged(noteBefore, noteMoveRedone), { before: noteBefore, after: noteMoveRedone })

  await clickToolbar(page, 'Reset')
  await delay(150)
  const noteReset = await blockState(page, 's1-note')
  check('reset restores selected layout', !rectChanged(noteBefore, noteReset), { before: noteBefore, after: noteReset })

  const metricBefore = await blockState(page, 's1-metric')
  await clickBlock(page, 's1-metric')
  await page.waitFor(`document.querySelector('[data-block="s1-metric"]')?.dataset.selected === 'true'`)
  await page.waitFor("!!document.querySelector('.resize-handle[data-handle=\"e\"]')")
  const handle = await resizeHandleCenter(page, 'e')
  await dragFromTo(page, handle, { x: handle.x + 44, y: handle.y })
  const metricResized = await blockState(page, 's1-metric')
  check(
    'Arrange Mode east resize changes width only',
    metricResized.width > metricBefore.width + 10 &&
      Math.abs(metricResized.height - metricBefore.height) < 1 &&
      metricResized.text === metricBefore.text,
    { before: metricBefore, after: metricResized },
  )
  await page.waitFor(`document.querySelector('[data-block="s1-metric"]')?.dataset.selected === 'true' && !!document.querySelector('.selection-overlay')`)
  const overlayFit = await selectionOverlayFitsBlock(page, 's1-metric')
  check('Arrange Mode selection follows autoheight block', overlayFit.fits, overlayFit)

  await clickSlide(page, 'Agenda')
  const cardBeforeResize = await blockState(page, 's2-step-2')
  await clickBlock(page, 's2-step-2')
  await page.waitFor(`document.querySelector('[data-block="s2-step-2"]')?.dataset.selected === 'true'`)
  await page.waitFor("!!document.querySelector('.resize-handle[data-handle=\"se\"]')")
  const cornerHandle = await resizeHandleCenter(page, 'se')
  await dragFromTo(page, cornerHandle, {
    x: cornerHandle.x + 48,
    y: cornerHandle.y + 48,
  })
  const cardCornerResized = await blockState(page, 's2-step-2')
  check(
    'Arrange Mode corner resize grows a card in both axes',
    cardCornerResized.width > cardBeforeResize.width + 10 &&
      cardCornerResized.height > cardBeforeResize.height + 10 &&
      cardCornerResized.text === cardBeforeResize.text,
    { before: cardBeforeResize, after: cardCornerResized },
  )
  await clickToolbar(page, 'Reset')
  await delay(150)
  const cardResizeReset = await blockState(page, 's2-step-2')
  check(
    'Arrange Mode reset restores card width and height',
    !rectChanged(cardBeforeResize, cardResizeReset) &&
      cardResizeReset.text === cardBeforeResize.text,
    { before: cardBeforeResize, after: cardResizeReset },
  )
  await clickSlide(page, 'Overview')

  await clickBlock(page, 's1-title')
  const noEditorInLayout = await page.eval("!document.querySelector('[data-editing=\"true\"]')")
  check('Arrange Mode click does not edit text', noEditorInLayout, null)

  await pressKey(page, 'Escape')
  await delay(100)
  const layoutEscapeState = await page.eval(`(() => {
    const reset = document.querySelector('button[data-action="reset"]')

    return {
      mode: document.querySelector('.retouch-app')?.dataset.mode,
      overlayCount: document.querySelectorAll('.selection-overlay').length,
      resetDisabled: reset?.disabled ?? null,
      resetScope: reset?.dataset.resetScope ?? null,
      selectedBlocks: document.querySelectorAll('[data-selected="true"]').length,
    }
  })()`)
  check(
    'Arrange Mode Escape clears selection',
    layoutEscapeState.mode === 'layout' &&
      layoutEscapeState.overlayCount === 0 &&
      layoutEscapeState.selectedBlocks === 0 &&
      layoutEscapeState.resetDisabled === true,
    layoutEscapeState,
  )

  await clickBlock(page, 's1-title')
  await page.waitFor(`document.querySelector('[data-block="s1-title"]')?.dataset.selected === 'true'`)
  await clickMode(page, 'Text')
  const textModeAfterArrange = await page.eval(`(() => {
    const reset = document.querySelector('button[data-action="reset"]')

    return {
      mode: document.querySelector('.retouch-app')?.dataset.mode,
      overlayCount: document.querySelectorAll('.selection-overlay').length,
      resetDisabled: reset?.disabled ?? null,
      resetAriaLabel: reset?.getAttribute('aria-label'),
      resetScope: reset?.dataset.resetScope ?? null,
      resetTitle: reset?.getAttribute('title'),
      selectedBlocks: document.querySelectorAll('[data-selected="true"]').length,
    }
  })()`)
  check(
    'Text Mode clears hidden Arrange selection on mode switch',
    textModeAfterArrange.mode === 'text' &&
      textModeAfterArrange.overlayCount === 0 &&
      textModeAfterArrange.selectedBlocks === 0 &&
      textModeAfterArrange.resetDisabled === false &&
      textModeAfterArrange.resetAriaLabel === 'Reset deck' &&
      textModeAfterArrange.resetScope === 'deck' &&
      textModeAfterArrange.resetTitle === 'Reset deck',
    textModeAfterArrange,
  )
}

async function runExportScenario(page, cdpPort) {
  await clickMode(page, 'Text')
  const titleBefore = await blockState(page, 's1-title')
  const expectedTitle = `${titleBefore.text} Export`
  await page.eval(`document.querySelector('[data-block="s1-title"]').click()`)
  await waitForEditor(page)
  await typeEditorText(page, ' Export')
  await commitTextEditor(page)

  await clickMode(page, 'Arrange')
  const exportNoteBeforeMove = await blockState(page, 's1-note')
  await dragBlockBySlideUnits(page, 's1-note', 48, 32)
  await page.waitFor(`(() => {
    const block = document.querySelector('[data-block="s1-note"]')
    const rect = block.getBoundingClientRect()

    return Math.abs(rect.x - ${exportNoteBeforeMove.x}) > 1 ||
      Math.abs(rect.y - ${exportNoteBeforeMove.y}) > 1
  })()`)
  const exportNoteMoved = await blockState(page, 's1-note')
  check(
    'Export scenario moves note layout before exporting',
    rectChanged(exportNoteBeforeMove, exportNoteMoved),
    { before: exportNoteBeforeMove, after: exportNoteMoved },
  )

  await clickSlide(page, 'Agenda')
  const exportCardBeforeResize = await blockState(page, 's2-step-2')
  await clickBlock(page, 's2-step-2')
  await page.waitFor("!!document.querySelector('.resize-handle[data-handle=\"se\"]')")
  const exportCardResizeHandle = await resizeHandleCenter(page, 'se')
  await dragFromTo(page, exportCardResizeHandle, {
    x: exportCardResizeHandle.x + 48,
    y: exportCardResizeHandle.y + 48,
  })
  const exportCardResized = await blockState(page, 's2-step-2')
  check(
    'Export scenario resizes a card before exporting',
    exportCardResized.width > exportCardBeforeResize.width + 10 &&
      exportCardResized.height > exportCardBeforeResize.height + 10 &&
      exportCardResized.text === exportCardBeforeResize.text,
    { before: exportCardBeforeResize, after: exportCardResized },
  )
  const previewAgendaExportMetrics = await exportedPreviewMetrics(page, 'slide-2', [
    's2-step-2',
  ])
  await clickSlide(page, 'Overview')
  const previewOverviewExportMetrics = await exportedPreviewMetrics(
    page,
    'slide-1',
    ['s1-title', 's1-note'],
  )

  await clickToolbar(page, 'Copy HTML')
  await page.waitFor("!!document.querySelector('.export-buffer')")
  const exportState = await page.eval(`(() => {
    const value = document.querySelector('.export-buffer')?.value ?? ''
    const copyButton = document.querySelector('button[aria-label="Copy HTML"]')
    const downloadButton = document.querySelector('button[aria-label="Download HTML"]')
    const parsed = new DOMParser().parseFromString(value, 'text/html')
    const patchText =
      parsed.querySelector('script[type="application/json"][data-retouch-patch]')
        ?.textContent ?? ''
    let patch = null

    try {
      patch = JSON.parse(patchText)
    } catch {
      patch = null
    }

    return {
      hasEditedTitle: value.includes(${JSON.stringify(expectedTitle)}),
      hasDataSlide: value.includes('data-slide="slide-1"'),
      hasDataBlock: value.includes('data-block="s1-note"'),
      hasStyleCoordinates: /left:\\d/.test(value) && /top:\\d/.test(value),
      resizedCardStyle:
        parsed.querySelector('[data-block="s2-step-2"]')?.getAttribute('style') ?? '',
      isCompleteDocument:
        value.trimStart().toLowerCase().startsWith('<!doctype html>') &&
        !!parsed.querySelector('html head style') &&
        !!parsed.querySelector('html body main.deck'),
      hasPresentationPrintCss:
        value.includes('@page{size:16in 9in;margin:0;}') &&
        value.includes('@media print') &&
        value.includes('break-after:page'),
      hasSharedSlideTheme:
        value.includes('container-type:inline-size') &&
        value.includes('font-size: clamp(10px, 4.5313cqw, 58px);') &&
        value.includes('padding: clamp(3px, 1.7188cqw, 22px);') &&
        value.includes('.block-step.strong'),
      parsedEditedTitle:
        parsed.querySelector('[data-block="s1-title"]')?.textContent ?? '',
      hasRawEditorChrome:
        value.includes('data-editing') ||
        value.includes('plain-text-editor') ||
        value.includes('resize-handle'),
      hasCopyAction: !!copyButton,
      hasDownloadAction: !!downloadButton,
      hasVisibleRawCodePanel: !!document.querySelector('.export-panel'),
      patchVersion: patch?.version ?? null,
      patchTextCount: patch?.text?.length ?? null,
      patchLayoutCount: patch?.layout?.length ?? null,
      patchTitle:
        patch?.text?.find((entry) => entry.blockId === 's1-title')?.text ?? null,
      patchMovedNote:
        patch?.layout?.find((entry) => entry.blockId === 's1-note')?.rect ?? null,
      patchResizedCard:
        patch?.layout?.find((entry) => entry.blockId === 's2-step-2')?.rect ?? null,
      html: value,
    }
  })()`)
  check('export reflects current text and layout state', exportState.hasEditedTitle && exportState.hasDataSlide && exportState.hasDataBlock && exportState.hasStyleCoordinates && !exportState.hasRawEditorChrome, exportState)
  check('export is a standalone HTML document', exportState.isCompleteDocument && exportState.parsedEditedTitle === expectedTitle, exportState)
  check('export includes print-ready slide CSS', exportState.hasPresentationPrintCss, exportState)
  check('export uses the same slide theme tokens as preview', exportState.hasSharedSlideTheme, exportState)
  check(
    'export carries a structured retouch patch manifest',
    exportState.patchVersion === 1 &&
      exportState.patchTextCount >= 1 &&
      exportState.patchLayoutCount >= 1 &&
      exportState.patchTitle === expectedTitle &&
      exportState.patchMovedNote?.x !== undefined &&
      exportState.patchMovedNote?.y !== undefined &&
      exportState.patchMovedNote?.width !== undefined &&
      exportState.patchResizedCard?.width !== undefined &&
      exportState.patchResizedCard?.height !== undefined &&
      exportState.resizedCardStyle.includes('min-height:'),
    exportState,
  )
  check('export has a clear copy action', exportState.hasCopyAction, exportState)
  check('export has a direct HTML download action', exportState.hasDownloadAction, exportState)
  check('export does not expose raw code panel by default', !exportState.hasVisibleRawCodePanel, exportState)

  const exportedPage = await openExportedHtmlPage(cdpPort, exportState.html)
  const renderedOverviewExportMetrics = await exportedPreviewMetrics(
    exportedPage,
    'slide-1',
    ['s1-title', 's1-note'],
  )
  const renderedAgendaExportMetrics = await exportedPreviewMetrics(
    exportedPage,
    'slide-2',
    ['s2-step-2'],
  )
  check(
    'exported HTML renders like the current preview',
    renderedOverviewExportMetrics['s1-title']?.text ===
      previewOverviewExportMetrics['s1-title']?.text &&
      renderedOverviewExportMetrics['s1-note']?.text ===
        previewOverviewExportMetrics['s1-note']?.text &&
      renderedAgendaExportMetrics['s2-step-2']?.text ===
        previewAgendaExportMetrics['s2-step-2']?.text &&
      exportedBlockMetricsMatch(
        renderedOverviewExportMetrics['s1-title'],
        previewOverviewExportMetrics['s1-title'],
      ) &&
      exportedBlockMetricsMatch(
        renderedOverviewExportMetrics['s1-note'],
        previewOverviewExportMetrics['s1-note'],
      ) &&
      exportedBlockMetricsMatch(
        renderedAgendaExportMetrics['s2-step-2'],
        previewAgendaExportMetrics['s2-step-2'],
      ),
    {
      preview: {
        ...previewOverviewExportMetrics,
        ...previewAgendaExportMetrics,
      },
      exported: {
        ...renderedOverviewExportMetrics,
        ...renderedAgendaExportMetrics,
      },
    },
  )
  await exportedPage.close()

  await page.eval(`navigator.clipboard.writeText = async (value) => {
    window.__pptRetouchCopiedHtml = value
  }`)
  await clickSelector(page, SELECTOR.copyButton)
  await page.waitFor(`${querySelectorExpression(SELECTOR.copyButton)}?.getAttribute('aria-pressed') === 'true'`)
  const copied = await page.eval(`${querySelectorExpression(SELECTOR.copyButton)}?.getAttribute('aria-pressed') === 'true'`)
  check('copy action gives completion feedback', copied, null)
  const copiedState = await page.eval(`(() => {
    const copyButton = document.querySelector('button[aria-label="Copy HTML"]')

    return {
      copyState: copyButton?.dataset.copyState,
      title: copyButton?.getAttribute('title'),
    }
  })()`)
  check('copy action switches to copied state', copiedState.copyState === 'copied' && copiedState.title === 'Copied', copiedState)

  await clickMode(page, 'Text')
  await page.eval(`document.querySelector('[data-block="s1-title"]').click()`)
  await waitForEditor(page)
  await typeEditorText(page, ' LiveDraft')
  const liveDraftExportFeedback = await page.eval(`(() => {
    const copyButton = document.querySelector('button[aria-label="Copy HTML"]')

    return {
      editorOpen: !!document.querySelector('[data-editing=\"true\"]'),
      copyPressed: copyButton?.getAttribute('aria-pressed'),
      copyState: copyButton?.dataset.copyState,
      title: copyButton?.getAttribute('title'),
    }
  })()`)
  check(
    'copy feedback clears while a live text draft is visible',
    liveDraftExportFeedback.editorOpen &&
      liveDraftExportFeedback.copyPressed === 'false' &&
      liveDraftExportFeedback.copyState === 'idle' &&
      liveDraftExportFeedback.title === 'Copy HTML',
    liveDraftExportFeedback,
  )
  await cancelTextEditor(page)
  const canceledDraftExportFeedback = await page.eval(`(() => {
    const copyButton = document.querySelector('button[aria-label="Copy HTML"]')

    return {
      editorOpen: !!document.querySelector('[data-editing=\"true\"]'),
      copyPressed: copyButton?.getAttribute('aria-pressed'),
      copyState: copyButton?.dataset.copyState,
      title: copyButton?.getAttribute('title'),
    }
  })()`)
  check(
    'copy feedback returns after canceling an uncommitted draft',
    !canceledDraftExportFeedback.editorOpen &&
      canceledDraftExportFeedback.copyPressed === 'true' &&
      canceledDraftExportFeedback.copyState === 'copied' &&
      canceledDraftExportFeedback.title === 'Copied',
    canceledDraftExportFeedback,
  )

  await page.eval(`(() => {
    window.__pptRetouchOriginalExecCommand = document.execCommand
    navigator.clipboard.writeText = async () => {
      throw new Error('clipboard unavailable')
    }
    document.execCommand = () => false
  })()`)
  await clickSelector(page, SELECTOR.copyButton)
  await page.waitFor(`${querySelectorExpression(SELECTOR.copyButton)}?.dataset.copyState === 'failed'`)
  const failedCopyFeedback = await page.eval(`(() => {
    const copyButton = document.querySelector('button[aria-label="Copy HTML"]')
    const style = getComputedStyle(copyButton)

    document.execCommand = window.__pptRetouchOriginalExecCommand

    return {
      backgroundColor: style.backgroundColor,
      color: style.color,
      copyPressed: copyButton?.getAttribute('aria-pressed'),
      copyState: copyButton?.dataset.copyState,
      title: copyButton?.getAttribute('title'),
    }
  })()`)
  check(
    'copy action does not show copied state when clipboard fails',
    failedCopyFeedback.copyPressed === 'false' &&
      failedCopyFeedback.copyState === 'failed' &&
      failedCopyFeedback.title === 'Copy failed',
    failedCopyFeedback,
  )
  check(
    'copy failure is visibly distinguishable',
    failedCopyFeedback.color === 'rgb(185, 28, 28)' &&
      failedCopyFeedback.backgroundColor === 'rgb(254, 242, 242)',
    failedCopyFeedback,
  )

  await page.eval(`(() => {
    window.__pptRetouchDownloadAfterCopyFailure = null
    const originalClick = HTMLAnchorElement.prototype.click
    HTMLAnchorElement.prototype.click = function patchedClick() {
      window.__pptRetouchDownloadAfterCopyFailure = {
        download: this.download,
        href: this.href,
      }
      HTMLAnchorElement.prototype.click = originalClick
    }
  })()`)
  await clickSelector(page, SELECTOR.downloadButton)
  await page.waitFor(`(() => {
    const copyButton = document.querySelector('button[aria-label="Copy HTML"]')
    const downloadButton = document.querySelector('button[aria-label="Download HTML"]')

    return copyButton?.dataset.copyState === 'idle' &&
      downloadButton?.dataset.downloadState === 'downloaded'
  })()`)
  const downloadAfterCopyFailure = await page.eval(`(() => {
    const copyButton = document.querySelector('button[aria-label="Copy HTML"]')
    const downloadButton = document.querySelector('button[aria-label="Download HTML"]')

    return {
      copyState: copyButton?.dataset.copyState,
      copyTitle: copyButton?.getAttribute('title'),
      download: window.__pptRetouchDownloadAfterCopyFailure ?? null,
      downloadState: downloadButton?.dataset.downloadState,
      downloadTitle: downloadButton?.getAttribute('title'),
    }
  })()`)
  check(
    'download clears prior copy failure state',
    downloadAfterCopyFailure.copyState === 'idle' &&
      downloadAfterCopyFailure.copyTitle === 'Copy HTML' &&
      downloadAfterCopyFailure.downloadState === 'downloaded' &&
      downloadAfterCopyFailure.downloadTitle === 'Downloaded' &&
      downloadAfterCopyFailure.download?.download === 'retouched-slides.html',
    downloadAfterCopyFailure,
  )

  await page.eval(`(() => {
    navigator.clipboard.writeText = async () => {
      throw new Error('clipboard unavailable')
    }
    document.execCommand = () => false
  })()`)
  await clickSelector(page, SELECTOR.copyButton)
  await page.waitFor(`${querySelectorExpression(SELECTOR.copyButton)}?.dataset.copyState === 'failed'`)
  await page.eval(`document.execCommand = window.__pptRetouchOriginalExecCommand`)

  await page.eval(`document.querySelector('[data-block="s1-title"]').click()`)
  await waitForEditor(page)
  await typeEditorText(page, ' AfterFail')
  const failedDuringDraft = await page.eval(`(() => {
    const copyButton = document.querySelector('button[aria-label="Copy HTML"]')

    return {
      editorOpen: !!document.querySelector('[data-editing="true"]'),
      copyState: copyButton?.dataset.copyState,
      title: copyButton?.getAttribute('title'),
    }
  })()`)
  check(
    'copy failure clears while a live text draft is visible',
    failedDuringDraft.editorOpen &&
      failedDuringDraft.copyState === 'idle' &&
      failedDuringDraft.title === 'Copy HTML',
    failedDuringDraft,
  )
  await commitTextEditor(page)
  await page.waitFor(`(() => {
    const copyButton = document.querySelector('button[aria-label="Copy HTML"]')
    const title = document.querySelector('[data-block="s1-title"]')?.textContent ?? ''

    return title.endsWith('AfterFail') && copyButton?.dataset.copyState === 'idle'
  })()`)
  const failedAfterCommit = await page.eval(`(() => {
    const copyButton = document.querySelector('button[aria-label="Copy HTML"]')

    return {
      copyState: copyButton?.dataset.copyState,
      text: document.querySelector('[data-block="s1-title"]')?.textContent ?? '',
      title: copyButton?.getAttribute('title'),
    }
  })()`)
  check(
    'copy failure does not stick to a new export state',
    failedAfterCommit.text.endsWith('AfterFail') &&
      failedAfterCommit.copyState === 'idle' &&
      failedAfterCommit.title === 'Copy HTML',
    failedAfterCommit,
  )
  await clickToolbar(page, 'Undo')
  await page.waitFor(`document.querySelector('[data-block="s1-title"]')?.textContent === ${JSON.stringify(expectedTitle)}`)

  await page.eval(`(() => {
    window.__pptRetouchDownload = null
    const originalClick = HTMLAnchorElement.prototype.click
    HTMLAnchorElement.prototype.click = function patchedClick() {
      window.__pptRetouchDownload = {
        download: this.download,
        href: this.href,
      }
      HTMLAnchorElement.prototype.click = originalClick
    }
  })()`)
  await clickSelector(page, SELECTOR.downloadButton)
  await page.waitFor(`${querySelectorExpression(SELECTOR.downloadButton)}?.getAttribute('aria-pressed') === 'true'`)
  const downloadState = await page.eval(`(() => {
    const downloadButton = document.querySelector('button[aria-label="Download HTML"]')

    return {
      downloadState: downloadButton?.dataset.downloadState,
      title: downloadButton?.getAttribute('title'),
      download: window.__pptRetouchDownload ?? null,
    }
  })()`)
  check(
    'download action emits an HTML file',
    downloadState.downloadState === 'downloaded' &&
      downloadState.title === 'Downloaded' &&
      downloadState.download?.download === 'retouched-slides.html' &&
      downloadState.download?.href?.startsWith('blob:'),
    downloadState,
  )

  await clickMode(page, 'Text')
  await page.eval(`navigator.clipboard.writeText = async (value) => {
    window.__pptRetouchCopiedHtml = value
  }`)
  await page.eval(`document.querySelector('[data-block="s1-title"]').click()`)
  await waitForEditor(page)
  await typeEditorText(page, ' DraftCopy')
  await clickSelector(page, SELECTOR.copyButton)
  await waitForNoEditor(page)
  const immediateCopy = await page.eval(`(() => {
    const copied = window.__pptRetouchCopiedHtml ?? ''
    const parsed = new DOMParser().parseFromString(copied, 'text/html')

    return {
      blockText: document.querySelector('[data-block="s1-title"]')?.textContent ?? '',
      copiedTitle: parsed.querySelector('[data-block="s1-title"]')?.textContent ?? '',
      copiedHasDraft: copied.includes('DraftCopy'),
    }
  })()`)
  check(
    'Copy HTML commits live text before exporting',
    immediateCopy.blockText.endsWith('DraftCopy') &&
      immediateCopy.copiedTitle.endsWith('DraftCopy') &&
      immediateCopy.copiedHasDraft,
    immediateCopy,
  )
  await clickToolbar(page, 'Undo')
  await page.waitFor(`document.querySelector('[data-block="s1-title"]')?.textContent === ${JSON.stringify(expectedTitle)}`)

  await page.eval(`(() => {
    window.__pptRetouchBlobText = null
    const originalCreateObjectURL = URL.createObjectURL

    URL.createObjectURL = function patchedCreateObjectURL(blob) {
      blob.text().then((text) => {
        window.__pptRetouchBlobText = text
      })
      URL.createObjectURL = originalCreateObjectURL

      return originalCreateObjectURL.call(URL, blob)
    }
  })()`)
  await page.eval(`document.querySelector('[data-block="s1-title"]').click()`)
  await waitForEditor(page)
  await typeEditorText(page, ' DraftDownload')
  await clickSelector(page, SELECTOR.downloadButton)
  await page.waitFor("!document.querySelector('[data-editing=\"true\"]') && window.__pptRetouchBlobText !== null")
  const immediateDownload = await page.eval(`(() => {
    const html = window.__pptRetouchBlobText ?? ''
    const parsed = new DOMParser().parseFromString(html, 'text/html')

    return {
      blockText: document.querySelector('[data-block="s1-title"]')?.textContent ?? '',
      downloadedTitle:
        parsed.querySelector('[data-block="s1-title"]')?.textContent ?? '',
      downloadedHasDraft: html.includes('DraftDownload'),
    }
  })()`)
  check(
    'Download HTML commits live text before exporting',
    immediateDownload.blockText.endsWith('DraftDownload') &&
      immediateDownload.downloadedTitle.endsWith('DraftDownload') &&
      immediateDownload.downloadedHasDraft,
    immediateDownload,
  )
  await clickToolbar(page, 'Undo')
  await page.waitFor(`document.querySelector('[data-block="s1-title"]')?.textContent === ${JSON.stringify(expectedTitle)}`)
}

async function runPersistenceScenario(page) {
  const beforeReload = await blockState(page, 's1-title')
  const changedThumbBeforeReload = await slideThumbState(page, 'Overview')
  await page.send('Page.reload', { ignoreCache: true })
  await delay(300)
  await page.waitFor(
    "document.readyState === 'complete' && !!document.querySelector('[data-block=\"s1-title\"]')",
  )
  const afterReload = await blockState(page, 's1-title')
  const changedThumbAfterReload = await slideThumbState(page, 'Overview')
  const stored = await page.eval(`(() => {
    const raw = localStorage.getItem('ppt-retouch:v3:deck')
    const parsed = raw ? JSON.parse(raw) : null

    return {
      hasStoredDeck: !!parsed?.deck?.slides?.length,
      version: parsed?.version ?? null,
      title:
        parsed?.deck?.slides?.[0]?.blocks?.find((block) => block.id === 's1-title')
          ?.text ?? null,
    }
  })()`)

  check(
    'edits persist after reload',
    afterReload.text === beforeReload.text && stored.title === beforeReload.text,
    { beforeReload, afterReload, stored },
  )
  check('autosave stores a versioned deck', stored.version === 1 && stored.hasStoredDeck, stored)
  check(
    'slide thumbnails mark edited slides',
    changedThumbBeforeReload.changed === 'true' &&
      changedThumbAfterReload.changed === 'true' &&
      changedThumbAfterReload.ariaLabel.includes('modified') &&
      changedThumbAfterReload.hasChangeDot,
    { beforeReload: changedThumbBeforeReload, afterReload: changedThumbAfterReload },
  )

  const resetDraftTitle = `${beforeReload.text} DraftReset`
  const noteBeforeTextReset = await blockState(page, 's1-note')
  await page.eval(`document.querySelector('[data-block="s1-title"]').click()`)
  await waitForEditor(page)
  await typeEditorText(page, ' DraftReset')
  const textResetReady = await page.eval(`(() => {
    const reset = document.querySelector('button[data-action="reset"]')

    return {
      resetAriaLabel: reset?.getAttribute('aria-label') ?? null,
      resetDisabled: reset?.disabled ?? null,
      resetScope: reset?.dataset.resetScope ?? null,
      resetTitle: reset?.getAttribute('title') ?? null,
    }
  })()`)
  check(
    'Text Mode reset target is explicit',
    textResetReady.resetDisabled === false &&
      textResetReady.resetAriaLabel === 'Reset text' &&
      textResetReady.resetScope === 'text' &&
      textResetReady.resetTitle === 'Reset text',
    textResetReady,
  )
  await clickToolbar(page, 'Reset')
  await page.waitFor(`document.querySelector('[data-block="s1-title"]')?.textContent === 'PPT Retouch'`)
  await page.waitFor(`document.querySelector('button[data-action="reset"]')?.disabled === true`)
  await waitForNoEditor(page)
  await page.waitFor(`(() => {
    const raw = localStorage.getItem('ppt-retouch:v3:deck')
    const parsed = raw ? JSON.parse(raw) : null

    return parsed?.deck?.slides?.[0]?.blocks?.find((block) => block.id === 's1-title')
      ?.text === 'PPT Retouch'
  })()`)
  const resetState = await blockState(page, 's1-title')
  const noteAfterTextReset = await blockState(page, 's1-note')
  const resetThumbState = await slideThumbState(page, 'Overview')
  const resetStorage = await page.eval(`(() => {
    const raw = localStorage.getItem('ppt-retouch:v3:deck')
    const parsed = raw ? JSON.parse(raw) : null

    return {
      title:
        parsed?.deck?.slides?.[0]?.blocks?.find((block) => block.id === 's1-title')
          ?.text ?? null,
      version: parsed?.version ?? null,
    }
  })()`)

  check(
    'Text Mode reset restores selected text only',
    resetState.text === 'PPT Retouch' &&
      resetState.undoDisabled === false &&
      resetState.resetDisabled === true &&
      !rectChanged(noteBeforeTextReset, noteAfterTextReset),
    {
      resetState,
      noteBefore: noteBeforeTextReset,
      noteAfter: noteAfterTextReset,
    },
  )
  check('Text Mode reset is undoable', resetState.undoDisabled === false, resetState)
  check('Text Mode reset updates autosave', resetStorage.title === 'PPT Retouch' && resetStorage.version === 1, resetStorage)
  check(
    'selected text reset keeps other slide changes marked',
    resetThumbState.changed === 'true' && resetThumbState.hasChangeDot,
    resetThumbState,
  )

  await clickToolbar(page, 'Undo')
  await page.waitFor(`document.querySelector('[data-block="s1-title"]')?.textContent === ${JSON.stringify(resetDraftTitle)}`)
  const undoResetState = await blockState(page, 's1-title')
  const undoResetThumbState = await slideThumbState(page, 'Overview')
  check(
    'undo restores selected text reset draft',
    undoResetState.text === resetDraftTitle &&
      undoResetState.resetDisabled === false &&
      undoResetThumbState.changed === 'true',
    { state: undoResetState, thumb: undoResetThumbState },
  )

  await clickToolbar(page, 'Redo')
  await page.waitFor(`document.querySelector('[data-block="s1-title"]')?.textContent === 'PPT Retouch'`)
  const redoResetState = await blockState(page, 's1-title')
  const redoResetThumbState = await slideThumbState(page, 'Overview')
  check(
    'redo reapplies selected text reset',
    redoResetState.text === 'PPT Retouch' &&
      redoResetState.resetDisabled === true &&
      redoResetThumbState.changed === 'true',
    { state: redoResetState, thumb: redoResetThumbState },
  )

  await pressKey(page, 'Escape')
  await delay(100)
  const textEscapeResetReady = await page.eval(`(() => {
    const reset = document.querySelector('button[data-action="reset"]')

    return {
      resetDisabled: reset?.disabled ?? null,
      resetAriaLabel: reset?.getAttribute('aria-label') ?? null,
      resetScope: reset?.dataset.resetScope ?? null,
      resetTitle: reset?.getAttribute('title'),
      selectedBlocks: document.querySelectorAll('[data-selected="true"]').length,
    }
  })()`)
  check(
    'Text Mode Escape clears selected text and exposes deck reset',
    textEscapeResetReady.resetDisabled === false &&
      textEscapeResetReady.resetAriaLabel === 'Reset deck' &&
      textEscapeResetReady.resetScope === 'deck' &&
      textEscapeResetReady.resetTitle === 'Reset deck' &&
      textEscapeResetReady.selectedBlocks === 0,
    textEscapeResetReady,
  )

  await clickStageBackground(page)
  const deckResetReady = await page.eval(`(() => {
    const reset = document.querySelector('button[data-action="reset"]')

    return {
      resetDisabled: reset?.disabled ?? null,
      resetAriaLabel: reset?.getAttribute('aria-label') ?? null,
      resetScope: reset?.dataset.resetScope ?? null,
      resetTitle: reset?.getAttribute('title'),
      selectedBlocks: document.querySelectorAll('[data-selected="true"]').length,
    }
  })()`)
  check(
    'Text Mode empty stage click exposes deck reset',
    deckResetReady.resetDisabled === false &&
      deckResetReady.resetAriaLabel === 'Reset deck' &&
      deckResetReady.resetScope === 'deck' &&
      deckResetReady.resetTitle === 'Reset deck' &&
      deckResetReady.selectedBlocks === 0,
    deckResetReady,
  )

  await clickToolbar(page, 'Reset')
  await page.waitFor(`document.querySelector('.slide-thumb[aria-current="page"]')?.dataset.changed === 'false'`)
  const deckResetState = await blockState(page, 's1-title')
  const noteAfterDeckReset = await blockState(page, 's1-note')
  const deckResetThumbState = await slideThumbState(page, 'Overview')
  const deckResetStorage = await page.eval(`localStorage.getItem('ppt-retouch:v3:deck')`)
  check(
    'Text Mode deck reset restores all slide changes after selection is cleared',
    deckResetState.text === 'PPT Retouch' &&
      deckResetState.resetDisabled === true &&
      deckResetThumbState.changed === 'false' &&
      deckResetStorage === null &&
      !deckResetThumbState.hasChangeDot &&
      rectChanged(noteBeforeTextReset, noteAfterDeckReset),
    {
      storage: deckResetStorage,
      state: deckResetState,
      thumb: deckResetThumbState,
      noteBefore: noteBeforeTextReset,
      noteAfter: noteAfterDeckReset,
    },
  )
}

async function runMobileScenario(cdpPort) {
  const page = await openPage(cdpPort, appUrl, {
    width: 390,
    height: 844,
    deviceScaleFactor: 1,
    mobile: true,
  })
  const mobile = await page.eval(`(() => ({
    clientWidth: document.documentElement.clientWidth,
    scrollWidth: document.documentElement.scrollWidth,
    thumbs: document.querySelectorAll('.slide-thumb').length,
    hasArrangeMode: !!Array.from(document.querySelectorAll('.mode-button')).find((button) => button.textContent?.trim() === 'Arrange'),
    slideWidth: document.querySelector('.slide-canvas').getBoundingClientRect().width,
    slideHeight: document.querySelector('.slide-canvas').getBoundingClientRect().height,
    stageClientWidth: document.querySelector('.stage-shell').clientWidth,
    stageScrollWidth: document.querySelector('.stage-shell').scrollWidth,
    topbarHeight: document.querySelector('.topbar').getBoundingClientRect().height,
    railHeight: document.querySelector('.slide-rail').getBoundingClientRect().height,
  }))()`)
  check('mobile viewport has no horizontal page overflow', mobile.scrollWidth <= mobile.clientWidth + 1, mobile)
  check('mobile keeps core controls reachable', mobile.thumbs >= 3 && mobile.hasArrangeMode, mobile)
  check('mobile keeps slide readable with internal stage pan', mobile.slideWidth >= 560 && mobile.slideHeight >= 315 && mobile.stageScrollWidth > mobile.stageClientWidth, mobile)
  check('mobile chrome stays compact', mobile.topbarHeight <= 56 && mobile.railHeight <= 96, mobile)
  await page.close()
}

async function blockState(page, blockId) {
  return page.eval(`(() => {
    const block = document.querySelector('[data-block="${blockId}"]')
    const rect = block.getBoundingClientRect()
    const undo = document.querySelector('button[aria-label="Undo"]')
    const redo = document.querySelector('button[aria-label="Redo"]')
    const reset = document.querySelector('button[data-action="reset"]')
    return {
      text: block.textContent,
      role: block.dataset.role,
      x: rect.x,
      y: rect.y,
      width: rect.width,
      height: rect.height,
      selected: block.dataset.selected,
      editorOpen: !!document.querySelector('[data-editing=\"true\"]'),
      undoDisabled: undo?.disabled ?? null,
      redoDisabled: redo?.disabled ?? null,
      resetAriaLabel: reset?.getAttribute('aria-label') ?? null,
      resetDisabled: reset?.disabled ?? null,
      resetScope: reset?.dataset.resetScope ?? null,
      resetTitle: reset?.getAttribute('title') ?? null,
    }
  })()`)
}

async function blockChrome(page, blockId) {
  return page.eval(`(() => {
    const block = document.querySelector('[data-block="${blockId}"]')
    const style = getComputedStyle(block)

    return {
      outlineColor: style.outlineColor,
      outlineStyle: style.outlineStyle,
      outlineWidth: style.outlineWidth,
    }
  })()`)
}

async function blockFitsSlide(page, blockId) {
  return page.eval(`(() => {
    const block = document.querySelector('[data-block="${blockId}"]')
    const slide = document.querySelector('.slide-canvas')
    const blockRect = block.getBoundingClientRect()
    const slideRect = slide.getBoundingClientRect()
    const style = getComputedStyle(block)

    return {
      blockTop: blockRect.top,
      blockBottom: blockRect.bottom,
      slideTop: slideRect.top,
      slideBottom: slideRect.bottom,
      scrollHeight: block.scrollHeight,
      clientHeight: block.clientHeight,
      overflowY: style.overflowY,
      fits:
        blockRect.top >= slideRect.top - 1 &&
        blockRect.bottom <= slideRect.bottom + 1 &&
        block.scrollHeight <= block.clientHeight + 1,
    }
  })()`)
}

async function editorFitsSlide(page) {
  return page.eval(`(() => {
    const editor = document.querySelector('[data-editing=\"true\"]')
    const slide = document.querySelector('.slide-canvas')
    const editorRect = editor.getBoundingClientRect()
    const slideRect = slide.getBoundingClientRect()
    const style = getComputedStyle(editor)

    return {
      editorTop: editorRect.top,
      editorBottom: editorRect.bottom,
      slideTop: slideRect.top,
      slideBottom: slideRect.bottom,
      scrollHeight: editor.scrollHeight,
      clientHeight: editor.clientHeight,
      overflowY: style.overflowY,
      fits:
        editorRect.top >= slideRect.top - 1 &&
        editorRect.bottom <= slideRect.bottom + 1 &&
        editor.scrollHeight <= editor.clientHeight + 1,
    }
  })()`)
}

async function selectionOverlayFitsBlock(page, blockId) {
  return page.eval(`(() => {
    const block = document.querySelector('[data-block="${blockId}"]')
    const overlay = document.querySelector('.selection-overlay')

    if (!block || !overlay) {
      return {
        fits: false,
        hasBlock: !!block,
        hasOverlay: !!overlay,
      }
    }

    const blockRect = block.getBoundingClientRect()
    const overlayRect = overlay.getBoundingClientRect()

    return {
      blockHeight: blockRect.height,
      overlayHeight: overlayRect.height,
      blockWidth: blockRect.width,
      overlayWidth: overlayRect.width,
      fits:
        Math.abs(blockRect.height - overlayRect.height) < 1 &&
        Math.abs(blockRect.width - overlayRect.width) < 1,
    }
  })()`)
}

async function inlineFits(page, selector) {
  return page.eval(`(() => {
    const element = document.querySelector(${JSON.stringify(selector)})
    const rect = element.getBoundingClientRect()
    const style = getComputedStyle(element)

    return {
      width: rect.width,
      scrollWidth: element.scrollWidth,
      clientWidth: element.clientWidth,
      overflowWrap: style.overflowWrap,
      fits: element.scrollWidth <= element.clientWidth + 1,
    }
  })()`)
}

async function openExportedHtmlPage(cdpPort, html) {
  return openPage(
    cdpPort,
    `data:text/html;charset=utf-8,${encodeURIComponent(html)}`,
    {
      width: 1500,
      height: 900,
      deviceScaleFactor: 1,
      mobile: false,
    },
  )
}

async function exportedPreviewMetrics(page, slideId, blockIds) {
  return page.eval(`(() => {
    const slide = document.querySelector('[data-slide="${slideId}"]')
    const ids = ${JSON.stringify(blockIds)}
    const result = {}

    if (!slide) {
      for (const blockId of ids) {
        result[blockId] = null
      }

      return result
    }

    const slideRect = slide.getBoundingClientRect()

    for (const blockId of ids) {
      const block = slide.querySelector('[data-block="' + blockId + '"]')

      if (!block) {
        result[blockId] = null
        continue
      }

      const textElement = block.querySelector('.slide-block-text') ?? block
      const range = document.createRange()
      range.selectNodeContents(textElement)
      const blockRect = block.getBoundingClientRect()
      const textRect = range.getBoundingClientRect()
      const xScale = 1280 / slideRect.width
      const yScale = 720 / slideRect.height

      result[blockId] = {
        text: block.textContent,
        x: (blockRect.left - slideRect.left) * xScale,
        y: (blockRect.top - slideRect.top) * yScale,
        width: blockRect.width * xScale,
        height: blockRect.height * yScale,
        textX: (textRect.left - slideRect.left) * xScale,
        textY: (textRect.top - slideRect.top) * yScale,
        textWidth: textRect.width * xScale,
        textHeight: textRect.height * yScale,
      }
    }

    return result
  })()`)
}

function exportedBlockMetricsMatch(exported, preview) {
  if (!exported || !preview) {
    return false
  }

  const blockTolerance = 2
  const textPositionTolerance = 18
  const textSizeTolerance = Math.max(18, preview.textWidth * 0.05)

  return (
    Math.abs(exported.x - preview.x) < blockTolerance &&
    Math.abs(exported.y - preview.y) < blockTolerance &&
    Math.abs(exported.width - preview.width) < blockTolerance &&
    Math.abs(exported.height - preview.height) < blockTolerance &&
    Math.abs(exported.textX - preview.textX) < textPositionTolerance &&
    Math.abs(exported.textY - preview.textY) < textPositionTolerance &&
    Math.abs(exported.textWidth - preview.textWidth) < textSizeTolerance &&
    Math.abs(exported.textHeight - preview.textHeight) < textPositionTolerance
  )
}

async function editorHeight(page) {
  return page.eval(`${querySelectorExpression(SELECTOR.editor)}?.getBoundingClientRect().height ?? 0`)
}

async function textRangeMetrics(page, selector) {
  return page.eval(`(() => {
    const element = document.querySelector(${JSON.stringify(selector)})
    const textElement =
      element.matches('[data-block]')
        ? element.querySelector('.slide-block-text') ?? element
        : element
    const range = document.createRange()
    range.selectNodeContents(textElement)
    const textRect = range.getBoundingClientRect()
    const boxRect = element.getBoundingClientRect()
    const canvasRect = document.querySelector(${JSON.stringify(SELECTOR.slideCanvas)})?.getBoundingClientRect()
    const stage = document.querySelector(${JSON.stringify(SELECTOR.stageShell)})

    return {
      boxTop: boxRect.top,
      boxLeft: boxRect.left,
      boxWidth: boxRect.width,
      boxHeight: boxRect.height,
      boxRight: boxRect.right,
      boxBottom: boxRect.bottom,
      canvasTop: canvasRect?.top ?? 0,
      canvasLeft: canvasRect?.left ?? 0,
      stageScrollTop: stage?.scrollTop ?? 0,
      stageScrollLeft: stage?.scrollLeft ?? 0,
      textTop: textRect.top,
      textLeft: textRect.left,
      textWidth: textRect.width,
      textHeight: textRect.height,
      textRight: textRect.right,
      textBottom: textRect.bottom,
      textInsideBox:
        textRect.top >= boxRect.top - 1 &&
        textRect.left >= boxRect.left - 1 &&
        textRect.right <= boxRect.right + 1 &&
        textRect.bottom <= boxRect.bottom + 1,
    }
  })()`)
}

async function editorMetrics(page) {
  return page.eval(`(() => {
    const wrapper = document.querySelector(${JSON.stringify(SELECTOR.editor)})
    const wrapperRect = wrapper?.getBoundingClientRect()

    return {
      top: wrapperRect?.top ?? 0,
      height: wrapperRect?.height ?? 0,
      textTop: wrapperRect?.top ?? 0,
      textHeight: wrapperRect?.height ?? 0,
      wrapperDisplay: wrapper ? getComputedStyle(wrapper).display : null,
      wrapperAlignItems: wrapper ? getComputedStyle(wrapper).alignItems : null,
      wrapperOverflowY: wrapper ? getComputedStyle(wrapper).overflowY : null,
      editorOverflowY: wrapper ? getComputedStyle(wrapper).overflowY : null,
      minHeight: wrapper ? getComputedStyle(wrapper).minHeight : null,
    }
  })()`)
}

async function clickMode(page, label) {
  await page.eval(`Array.from(document.querySelectorAll(${JSON.stringify(SELECTOR.modeButton)})).find((button) => button.textContent?.trim() === ${JSON.stringify(label)})?.click()`)
  await delay(100)
}

async function clickToolbar(page, label) {
  const actionByLabel = {
    'Copy HTML': 'copy-html',
    'Download HTML': 'download-html',
    Redo: 'redo',
    Reset: 'reset',
    Undo: 'undo',
  }
  const action = actionByLabel[label]
  const selector = action
    ? `button[data-action="${action}"]`
    : `button[aria-label="${label}"]`

  await clickSelector(page, selector)
}

async function clickSlide(page, label) {
  const activeSlideThumbSelector = `${SELECTOR.slideThumb}[aria-current="page"]`

  await page.eval(`Array.from(document.querySelectorAll(${JSON.stringify(SELECTOR.slideThumb)})).find((button) => button.textContent?.includes(${JSON.stringify(label)}))?.click()`)
  await page.waitFor(`${querySelectorExpression(activeSlideThumbSelector)}?.textContent?.includes(${JSON.stringify(label)})`)
  await delay(100)
}

async function slideBlockOrder(page) {
  return page.eval(
    `Array.from(document.querySelectorAll(${JSON.stringify(`${SELECTOR.slideCanvas} [data-block]`)})).map((block) => block.dataset.block)`,
  )
}

function blockSelector(blockId) {
  return `[data-block="${blockId}"]`
}

function querySelectorExpression(selector) {
  return `document.querySelector(${JSON.stringify(selector)})`
}

function presentSelectorExpression(selector) {
  return `!!${querySelectorExpression(selector)}`
}

async function waitForEditor(page) {
  await page.waitFor(presentSelectorExpression(SELECTOR.editorEditable))
}

async function waitForNoEditor(page) {
  await page.waitFor(`!${querySelectorExpression(SELECTOR.editor)}`)
}

async function focusEditor(page) {
  await page.eval(`${querySelectorExpression(SELECTOR.editor)}?.focus()`)
}

async function clickSelector(page, selector) {
  await page.eval(`${querySelectorExpression(selector)}?.click()`)
}

function sameArray(a, b) {
  return a.length === b.length && a.every((value, index) => value === b[index])
}

async function slideThumbState(page, label) {
  return page.eval(`(() => {
    const thumb = Array.from(document.querySelectorAll('.slide-thumb'))
      .find((button) => button.textContent?.includes('${label}'))

    return {
      ariaLabel: thumb?.getAttribute('aria-label') ?? '',
      changed: thumb?.dataset.changed ?? null,
      hasChangeDot: !!thumb?.querySelector('.thumb-change'),
    }
  })()`)
}

async function focusBlockAndPress(page, blockId, key) {
  await page.eval(`${querySelectorExpression(blockSelector(blockId))}?.focus()`)
  await pressKey(page, key)
}

async function pressKey(page, key, options = {}) {
  const keyCodeByKey = {
    ArrowDown: 40,
    ArrowLeft: 37,
    ArrowRight: 39,
    ArrowUp: 38,
    Enter: 13,
    Escape: 27,
  }
  const code = key.startsWith('Arrow') || key === 'Enter' || key === 'Escape'
    ? key
    : `Key${key.toUpperCase()}`
  const keyCode = keyCodeByKey[key] ?? key.toUpperCase().charCodeAt(0)
  const modifiers = options.shift ? 8 : 0

  await page.send('Input.dispatchKeyEvent', {
    type: 'keyDown',
    key,
    code,
    modifiers,
    windowsVirtualKeyCode: keyCode,
    nativeVirtualKeyCode: keyCode,
  })
  await page.send('Input.dispatchKeyEvent', {
    type: 'keyUp',
    key,
    code,
    modifiers,
    windowsVirtualKeyCode: keyCode,
    nativeVirtualKeyCode: keyCode,
  })
}

async function pressHistoryShortcut(page, key) {
  const normalizedKey = key.toLowerCase()
  await page.send('Input.dispatchKeyEvent', {
    type: 'keyDown',
    key: normalizedKey,
    code: `Key${normalizedKey.toUpperCase()}`,
    windowsVirtualKeyCode: normalizedKey.toUpperCase().charCodeAt(0),
    nativeVirtualKeyCode: normalizedKey.toUpperCase().charCodeAt(0),
    modifiers: 2,
  })
  await page.send('Input.dispatchKeyEvent', {
    type: 'keyUp',
    key: normalizedKey,
    code: `Key${normalizedKey.toUpperCase()}`,
    windowsVirtualKeyCode: normalizedKey.toUpperCase().charCodeAt(0),
    nativeVirtualKeyCode: normalizedKey.toUpperCase().charCodeAt(0),
    modifiers: 2,
  })
}

async function commitTextEditor(page) {
  await focusEditor(page)
  await page.send('Input.dispatchKeyEvent', {
    type: 'keyDown',
    key: 'Enter',
    code: 'Enter',
    windowsVirtualKeyCode: 13,
    nativeVirtualKeyCode: 13,
    modifiers: 2,
  })
  await page.send('Input.dispatchKeyEvent', {
    type: 'keyUp',
    key: 'Enter',
    code: 'Enter',
    windowsVirtualKeyCode: 13,
    nativeVirtualKeyCode: 13,
    modifiers: 2,
  })
  await waitForNoEditor(page)
}

async function pressEditorEnter(page) {
  await focusEditor(page)
  await page.send('Input.dispatchKeyEvent', {
    type: 'keyDown',
    key: 'Enter',
    code: 'Enter',
    windowsVirtualKeyCode: 13,
    nativeVirtualKeyCode: 13,
  })
  await page.send('Input.dispatchKeyEvent', {
    type: 'keyUp',
    key: 'Enter',
    code: 'Enter',
    windowsVirtualKeyCode: 13,
    nativeVirtualKeyCode: 13,
  })
}

async function cancelTextEditor(page) {
  await focusEditor(page)
  await page.send('Input.dispatchKeyEvent', {
    type: 'keyDown',
    key: 'Escape',
    code: 'Escape',
    windowsVirtualKeyCode: 27,
    nativeVirtualKeyCode: 27,
  })
  await page.send('Input.dispatchKeyEvent', {
    type: 'keyUp',
    key: 'Escape',
    code: 'Escape',
    windowsVirtualKeyCode: 27,
    nativeVirtualKeyCode: 27,
  })
  await waitForNoEditor(page)
}

async function replaceEditorText(page, text) {
  await page.eval(`(() => {
    const editor = document.querySelector(${JSON.stringify(SELECTOR.editor)})
    editor.focus()
    const selection = window.getSelection()
    const range = document.createRange()
    range.selectNodeContents(editor)
    selection.removeAllRanges()
    selection.addRange(range)
  })()`)
  await typeEditorText(page, text)
}

async function typeEditorText(page, text) {
  await focusEditor(page)

  for (const char of text) {
    if (char === '\n') {
      await pressEditorEnter(page)
      continue
    }

    const { code, keyCode } = keyDescriptorForChar(char)

    await page.send('Input.dispatchKeyEvent', {
      type: 'keyDown',
      key: char,
      code,
      windowsVirtualKeyCode: keyCode,
      nativeVirtualKeyCode: keyCode,
      text: char,
      unmodifiedText: char,
    })
    await page.send('Input.dispatchKeyEvent', {
      type: 'keyUp',
      key: char,
      code,
      windowsVirtualKeyCode: keyCode,
      nativeVirtualKeyCode: keyCode,
    })
  }
}

async function insertEditorTextCommand(page, text) {
  const before = await page.eval(`(() => {
    const editor = document.querySelector(${JSON.stringify(SELECTOR.editor)})

    if (!editor) {
      return null
    }

    editor.focus()
    const selection = window.getSelection()
    const range = document.createRange()

    range.selectNodeContents(editor)
    range.collapse(false)
    selection.removeAllRanges()
    selection.addRange(range)

    return editor.textContent ?? ''
  })()`)
  await page.send('Input.insertText', { text })
  const after = await page.eval(
    `${querySelectorExpression(SELECTOR.editor)}?.textContent ?? null`,
  )

  return { before, after }
}

function keyDescriptorForChar(char) {
  const punctuation = {
    ' ': ['Space', 32],
    '.': ['Period', 190],
    ',': ['Comma', 188],
    '-': ['Minus', 189],
    "'": ['Quote', 222],
    ':': ['Semicolon', 186],
  }

  if (punctuation[char]) {
    const [code, keyCode] = punctuation[char]

    return { code, keyCode }
  }

  if (/^[0-9]$/.test(char)) {
    return { code: `Digit${char}`, keyCode: char.charCodeAt(0) }
  }

  return {
    code: /^[a-z]$/i.test(char) ? `Key${char.toUpperCase()}` : '',
    keyCode: char.toUpperCase().charCodeAt(0),
  }
}

async function setEditorText(page, text) {
  await page.eval(`(() => {
    const editor = document.querySelector(${JSON.stringify(SELECTOR.editor)})
    editor.focus()
    editor.textContent = ${JSON.stringify(text)}
    editor.dispatchEvent(new InputEvent('input', {
      bubbles: true,
      data: ${JSON.stringify(text)},
      inputType: 'insertText',
    }))
  })()`)
}

async function clickBlock(page, blockId, options = {}) {
  const center = await blockCenter(page, blockId)
  await clickAt(page, center, options)
}

async function clickStageBackground(page) {
  const point = await page.eval(`(() => {
    const stage = document.querySelector('.stage-shell')
    const rect = stage.getBoundingClientRect()

    return { x: rect.left + 8, y: rect.top + 8 }
  })()`)

  await clickAt(page, point)
}

async function resizeHandleCenter(page, handle) {
  return page.eval(`(() => {
    const handle = document.querySelector('.resize-handle[data-handle="${handle}"]')
    const rect = handle.getBoundingClientRect()

    return { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 }
  })()`)
}

async function dragBlock(page, blockId, dx, dy) {
  const center = await blockCenter(page, blockId)
  await dragFromTo(page, center, { x: center.x + dx, y: center.y + dy })
}

async function dragBlockBySlideUnits(page, blockId, dx, dy) {
  const center = await blockCenter(page, blockId)
  const delta = await slideUnitsToScreenDelta(page, dx, dy)
  await dragFromTo(page, center, { x: center.x + delta.x, y: center.y + delta.y })
}

async function dragBlockBySlideUnitsWithProbe(page, blockId, dx, dy) {
  const center = await blockCenter(page, blockId)
  const delta = await slideUnitsToScreenDelta(page, dx, dy)
  const target = { x: center.x + delta.x, y: center.y + delta.y }

  await page.send('Input.dispatchMouseEvent', {
    type: 'mouseMoved',
    x: center.x,
    y: center.y,
    button: 'none',
  })
  await page.send('Input.dispatchMouseEvent', {
    type: 'mousePressed',
    x: center.x,
    y: center.y,
    button: 'left',
    buttons: 1,
    clickCount: 1,
  })
  await page.send('Input.dispatchMouseEvent', {
    type: 'mouseMoved',
    x: target.x,
    y: target.y,
    button: 'left',
    buttons: 1,
  })
  await delay(100)

  const probe = await page.eval(`(() => {
    const slide = document.querySelector('.slide-canvas')
    const guideX = document.querySelector('.snap-guide-x')
    const guideY = document.querySelector('.snap-guide-y')
    const slideRect = slide.getBoundingClientRect()
    const guideXRect = guideX?.getBoundingClientRect()
    const guideYRect = guideY?.getBoundingClientRect()

    return {
      hasCenterGuideX:
        !!guideXRect &&
        Math.abs(guideXRect.left - (slideRect.left + slideRect.width / 2)) < 1,
      hasCenterGuideY:
        !!guideYRect &&
        Math.abs(guideYRect.top - (slideRect.top + slideRect.height / 2)) < 1,
      guideXSlideUnits: guideXRect
        ? ((guideXRect.left - slideRect.left) / slideRect.width) * 1280
        : null,
      hasGuideXAt770: guideXRect
        ? Math.abs(((guideXRect.left - slideRect.left) / slideRect.width) * 1280 - 770) < 1
        : false,
    }
  })()`)

  await page.send('Input.dispatchMouseEvent', {
    type: 'mouseReleased',
    x: target.x,
    y: target.y,
    button: 'left',
    buttons: 0,
    clickCount: 1,
  })
  await delay(150)

  return probe
}

async function slideUnitsToScreenDelta(page, dx, dy) {
  return page.eval(`(() => {
    const rect = document.querySelector('.slide-canvas').getBoundingClientRect()

    return {
      x: (${dx} / 1280) * rect.width,
      y: (${dy} / 720) * rect.height,
    }
  })()`)
}

async function blockCenterAlignment(page, blockId) {
  return page.eval(`(() => {
    const block = document.querySelector('[data-block="${blockId}"]')
    const slide = document.querySelector('.slide-canvas')
    const blockRect = block.getBoundingClientRect()
    const slideRect = slide.getBoundingClientRect()
    const blockCenterX = blockRect.left + blockRect.width / 2
    const slideCenterX = slideRect.left + slideRect.width / 2

    return {
      blockCenterX,
      slideCenterX,
      centeredX: Math.abs(blockCenterX - slideCenterX) < 1,
    }
  })()`)
}

async function blockEdgeAlignment(page, blockId, peerBlockId) {
  return page.eval(`(() => {
    const block = document.querySelector('[data-block="${blockId}"]')
    const peer = document.querySelector('[data-block="${peerBlockId}"]')
    const blockRect = block.getBoundingClientRect()
    const peerRect = peer.getBoundingClientRect()

    return {
      blockLeft: blockRect.left,
      peerRight: peerRect.right,
      leftToPeerRight: Math.abs(blockRect.left - peerRect.right) < 1,
    }
  })()`)
}

async function blockCenter(page, blockId) {
  return page.eval(`(() => {
    const block = document.querySelector('[data-block="${blockId}"]')
    const rect = block.getBoundingClientRect()
    return { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 }
  })()`)
}

async function clickAt(page, point, options = {}) {
  const modifiers = options.shift ? 8 : 0
  const modifierParams = modifiers === 0 ? {} : { modifiers }

  await page.send('Input.dispatchMouseEvent', {
    type: 'mouseMoved',
    x: point.x,
    y: point.y,
    button: 'none',
    ...modifierParams,
  })
  await page.send('Input.dispatchMouseEvent', {
    type: 'mousePressed',
    x: point.x,
    y: point.y,
    button: 'left',
    buttons: 1,
    clickCount: 1,
    ...modifierParams,
  })
  await delay(50)
  await page.send('Input.dispatchMouseEvent', {
    type: 'mouseReleased',
    x: point.x,
    y: point.y,
    button: 'left',
    buttons: 0,
    clickCount: 1,
    ...modifierParams,
  })
}

async function movePointerAway(page) {
  await page.send('Input.dispatchMouseEvent', {
    type: 'mouseMoved',
    x: 4,
    y: 4,
    button: 'none',
  })
  await delay(50)
}

async function dragFromTo(page, from, to) {
  await page.send('Input.dispatchMouseEvent', {
    type: 'mouseMoved',
    x: from.x,
    y: from.y,
    button: 'none',
  })
  await page.send('Input.dispatchMouseEvent', {
    type: 'mousePressed',
    x: from.x,
    y: from.y,
    button: 'left',
    buttons: 1,
    clickCount: 1,
  })
  await page.send('Input.dispatchMouseEvent', {
    type: 'mouseMoved',
    x: to.x,
    y: to.y,
    button: 'left',
    buttons: 1,
  })
  await page.send('Input.dispatchMouseEvent', {
    type: 'mouseReleased',
    x: to.x,
    y: to.y,
    button: 'left',
    buttons: 0,
    clickCount: 1,
  })
  await delay(150)
}

function rectChanged(a, b) {
  return (
    Math.abs(a.x - b.x) > 0.5 ||
    Math.abs(a.y - b.y) > 0.5 ||
    Math.abs(a.width - b.width) > 0.5 ||
    Math.abs(a.height - b.height) > 0.5
  )
}

function check(name, ok, detail) {
  checks.push({ name, ok: Boolean(ok), detail })
}

async function isHttpReady(url) {
  try {
    const response = await fetch(url)
    return response.ok
  } catch {
    return false
  }
}

async function getFreePort() {
  return new Promise((resolve, reject) => {
    const server = net.createServer()
    server.listen(0, '127.0.0.1', () => {
      const address = server.address()
      server.close(() => {
        if (address && typeof address === 'object') resolve(address.port)
        else reject(new Error('Failed to allocate port'))
      })
    })
    server.on('error', reject)
  })
}

async function waitUntil(predicate, message, timeout = 5000) {
  const started = Date.now()
  let last = null
  while (Date.now() - started < timeout) {
    last = await predicate()
    if (last) return last
    await delay(50)
  }
  throw new Error(`${message}; last=${JSON.stringify(last)}`)
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms))
}
