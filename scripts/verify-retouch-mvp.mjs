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
    width: 1280,
    height: 900,
    deviceScaleFactor: 1,
    mobile: false,
  })

  await runFirstScreenScenario(page)
  await runEditSurfaceParityScenario(page)
  await runTextScenario(page)
  await runLayoutScenario(page)
  await runExportScenario(page)
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
  if (chrome) {
    const exited = new Promise((resolve) => chrome.once('exit', resolve))
    chrome.kill()
    await Promise.race([exited, delay(1000)])
  }
  if (devServer) devServer.kill()
  if (chromeProfile) {
    await rm(chromeProfile, {
      recursive: true,
      force: true,
      maxRetries: 5,
      retryDelay: 100,
    })
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
    stdio: ['ignore', 'pipe', 'pipe'],
  })
  devServer.stdout.on('data', () => undefined)
  devServer.stderr.on('data', () => undefined)

  await waitUntil(async () => isHttpReady(appUrl), 'Vite dev server did not start', 15000)
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

    const { resolve, reject } = pending.get(message.id)
    pending.delete(message.id)
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
        pending.set(callId, { resolve, reject })
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
      ws.close()
    },
  }

  await page.send('Runtime.enable')
  await page.send('Log.enable')
  await page.send('Page.enable')
  if (viewport) {
    await page.send('Emulation.setDeviceMetricsOverride', viewport)
    await page.send('Page.navigate', { url })
  }
  await page.eval(`
    new Promise((resolve) => {
      if (document.readyState === 'complete') resolve()
      else addEventListener('load', resolve, { once: true })
    })
  `)
  await page.waitFor("!!document.querySelector('[data-block=\"s1-title\"]')")

  return page
}

async function runFirstScreenScenario(page) {
  const state = await page.eval(`(() => ({
    hasEditorShell: !!document.querySelector('.retouch-app'),
    slideCount: document.querySelectorAll('.slide-thumb').length,
    hasTextMode: !!Array.from(document.querySelectorAll('.mode-button')).find((button) => button.textContent?.trim() === 'Text'),
    hasArrangeMode: !!Array.from(document.querySelectorAll('.mode-button')).find((button) => button.textContent?.trim() === 'Arrange'),
    canvasBackgroundImage: getComputedStyle(document.querySelector('.slide-canvas')).backgroundImage,
    hasStarterCopy: document.body.textContent.includes('React + TypeScript + Vite'),
    hasMainSlideBlock: !!document.querySelector('[data-block="s1-title"]'),
  }))()`)

  check('first screen is retouch editor', state.hasEditorShell && state.hasMainSlideBlock, state)
  check('slide thumbnails are available', state.slideCount >= 3, state)
  check('mode toggle is available', state.hasTextMode && state.hasArrangeMode, state)
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
      await page.waitFor("!!document.querySelector('.plain-text-editor[contenteditable]')")
      const editor = await textRangeMetrics(page, '.plain-text-editor')

      deltas.push({
        blockId,
        slideName,
        boxHeightDelta: editor.boxHeight - preview.boxHeight,
        textHeightDelta: editor.textHeight - preview.textHeight,
        textLeftDelta: editor.textLeft - preview.textLeft,
        textTopDelta: editor.textTop - preview.textTop,
      })

      await cancelTextEditor(page)
    }
  }

  await clickSlide(page, 'Overview')

  check(
    'Text Mode editor opens on the same visual text box',
    deltas.every(
      (delta) =>
        Math.abs(delta.boxHeightDelta) < 1 &&
        Math.abs(delta.textHeightDelta) < 1 &&
        Math.abs(delta.textLeftDelta) < 1 &&
        Math.abs(delta.textTopDelta) < 1,
    ),
    deltas,
  )
}

async function runTextScenario(page) {
  const titleBefore = await blockState(page, 's1-title')
  const titlePreviewText = await textRangeMetrics(page, '[data-block="s1-title"]')

  await page.eval(`document.querySelector('[data-block="s1-title"]').click()`)
  await page.waitFor("!!document.querySelector('.plain-text-editor[contenteditable]')")
  await commitTextEditor(page)
  const titleNoOp = await blockState(page, 's1-title')
  check('no-op text edit does not create history', titleNoOp.text === titleBefore.text && titleNoOp.undoDisabled === true, titleNoOp)

  await page.eval(`document.querySelector('[data-block="s1-title"]').click()`)
  await page.waitFor("!!document.querySelector('.plain-text-editor[contenteditable]')")
  const titleEditorText = await textRangeMetrics(page, '.plain-text-editor')
  const editingTitle = await page.eval(`(() => ({
    originalBlockCount: document.querySelectorAll('[data-block="s1-title"]').length,
    editorCount: document.querySelectorAll('.plain-text-editor').length,
    editorText: document.querySelector('.plain-text-editor')?.textContent,
    contentEditable: document.querySelector('.plain-text-editor')?.contentEditable,
    hasMarkdownChrome:
      document.body.textContent.includes('**') ||
      document.body.textContent.includes('H1') ||
      document.body.textContent.includes('MD') ||
      !!document.querySelector('.ProseMirror, .nano, .block-picker, .command-palette'),
  }))()`)
  check('Text Mode replaces original block with one plaintext editor', editingTitle.originalBlockCount === 0 && editingTitle.editorCount === 1 && editingTitle.contentEditable === 'plaintext-only', editingTitle)
  check('Text Mode does not show markdown editor chrome', !editingTitle.hasMarkdownChrome, editingTitle)
  check(
    'Text Mode editor keeps preview text position',
    Math.abs(titleEditorText.textTop - titlePreviewText.textTop) < 1 &&
      Math.abs(titleEditorText.textLeft - titlePreviewText.textLeft) < 1 &&
      Math.abs(titleEditorText.textHeight - titlePreviewText.textHeight) < 1,
    { before: titlePreviewText, after: titleEditorText },
  )

  await page.send('Input.insertText', { text: ' Approved' })
  await commitTextEditor(page)
  const titleCommitted = await blockState(page, 's1-title')
  check('title text commit works', titleCommitted.text === `${titleBefore.text} Approved`, titleCommitted)
  check('plain Enter commits text edit', !titleCommitted.editorOpen, titleCommitted)
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

  const subtitleBefore = await blockState(page, 's1-subtitle')
  await page.eval(`document.querySelector('[data-block="s1-subtitle"]').click()`)
  await page.waitFor("!!document.querySelector('.plain-text-editor[contenteditable]')")
  const editBefore = await editorMetrics(page)
  await page.send('Input.insertText', {
    text: ' Extra context that wraps onto more lines so this fixed-width text box grows vertically instead of clipping or scrolling. Add owner names, renewal risk, partner follow-up, timing, decision path, and executive review notes so the content clearly exceeds the original text box height.',
  })
  await page.waitFor(`(() => {
    const editor = document.querySelector('.plain-text-editor')
    return editor && editor.getBoundingClientRect().height > ${editBefore.height + 10}
  })()`)
  const autoHeight = await editorMetrics(page)
  check('Text Mode autoheight grows box, not scroll/clip', autoHeight.height > editBefore.height + 10 && autoHeight.wrapperOverflowY === 'visible' && autoHeight.editorOverflowY === 'visible', autoHeight)
  check('Text Mode autoheight keeps content top stable while typing', Math.abs(autoHeight.top - editBefore.top) < 1 && Math.abs(autoHeight.textTop - editBefore.textTop) < 1, { before: editBefore, after: autoHeight })

  await commitTextEditor(page)
  const subtitleGrown = await blockState(page, 's1-subtitle')
  check('autoheight persists after commit', subtitleGrown.height > subtitleBefore.height + 10, { before: subtitleBefore, after: subtitleGrown })

  await clickToolbar(page, 'Undo')
  await delay(150)
  const subtitleUndone = await blockState(page, 's1-subtitle')
  check('undo restores text and autoheight', subtitleUndone.text === subtitleBefore.text && Math.abs(subtitleUndone.height - subtitleBefore.height) < 1, { before: subtitleBefore, after: subtitleUndone })

  await clickToolbar(page, 'Redo')
  await delay(150)
  const subtitleRedone = await blockState(page, 's1-subtitle')
  check('redo restores autoheight growth', subtitleRedone.text === subtitleGrown.text && Math.abs(subtitleRedone.height - subtitleGrown.height) < 1, { before: subtitleGrown, after: subtitleRedone })

  await page.eval(`document.querySelector('[data-block="s1-subtitle"]').click()`)
  await page.waitFor("!!document.querySelector('.plain-text-editor[contenteditable]')")
  const grownEditHeight = await editorHeight(page)
  await replaceEditorText(page, 'Short follow-up.')
  await page.waitFor(`(() => {
    const editor = document.querySelector('.plain-text-editor')
    return editor && editor.getBoundingClientRect().height < ${grownEditHeight - 10}
  })()`)
  const shrinkHeight = await editorHeight(page)
  check('Text Mode autoheight shrinks with shorter content', shrinkHeight < grownEditHeight - 10, { grownEditHeight, shrinkHeight })

  await commitTextEditor(page)
  const subtitleShrunk = await blockState(page, 's1-subtitle')
  check('autoheight shrink persists after commit', subtitleShrunk.text === 'Short follow-up.' && subtitleShrunk.height < subtitleGrown.height - 10, { before: subtitleGrown, after: subtitleShrunk })

  await page.eval(`document.querySelector('[data-block="s1-subtitle"]').click()`)
  await page.waitFor("!!document.querySelector('.plain-text-editor[contenteditable]')")
  await replaceEditorText(
    page,
    'enterprise-renewal-risk-account-owner-review-decision-path-follow-up-needed-before-month-close',
  )
  const longEditorFit = await inlineFits(page, '.plain-text-editor')
  check('Text Mode wraps long unbroken text while editing', longEditorFit.fits, longEditorFit)
  await commitTextEditor(page)
  const longPreviewFit = await inlineFits(page, '[data-block="s1-subtitle"]')
  check('preview wraps long unbroken text after commit', longPreviewFit.fits, longPreviewFit)

  await page.eval(`document.querySelector('[data-block="s1-note"]').click()`)
  await page.waitFor("!!document.querySelector('.plain-text-editor[contenteditable]')")
  const noteBeforeCancel = await page.eval(`document.querySelector('.plain-text-editor')?.textContent ?? ''`)
  await page.send('Input.insertText', { text: ' Cancelled draft' })
  await cancelTextEditor(page)
  const noteAfterCancel = await blockState(page, 's1-note')
  check('Escape cancels text draft', noteAfterCancel.text === noteBeforeCancel, noteAfterCancel)

  await clickSlide(page, 'Decision')
  const bottomNoteBefore = await blockState(page, 's3-note')
  await page.eval(`document.querySelector('[data-block="s3-note"]').click()`)
  await page.waitFor("!!document.querySelector('.plain-text-editor[contenteditable]')")
  const bottomNoteDraftText =
    'Decision needed: approve the retention sprint with owner, timing, customer impact, renewal coverage, executive sponsor, next action, weekly check-in, risk review, onboarding help, and discount guidance visible in the slide.'
  await replaceEditorText(
    page,
    bottomNoteDraftText,
  )
  await page.waitFor(`(() => {
    const editor = document.querySelector('.plain-text-editor')
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
  const noteBefore = await blockState(page, 's1-note')

  await dragBlock(page, 's1-note', 42, 24)
  const noteMoved = await blockState(page, 's1-note')
  check('Arrange Mode drag moves block only', rectChanged(noteBefore, noteMoved) && noteMoved.text === noteBefore.text, { before: noteBefore, after: noteMoved })
  check('drag enables reset', noteMoved.resetDisabled === false, noteMoved)

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
  const handle = await page.eval(`(() => {
    const handle = document.querySelector('.resize-handle[data-handle="e"]')
    const rect = handle.getBoundingClientRect()
    return { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 }
  })()`)
  await dragFromTo(page, handle, { x: handle.x + 44, y: handle.y })
  const metricResized = await blockState(page, 's1-metric')
  check('Arrange Mode resize changes width only', metricResized.width > metricBefore.width + 10 && metricResized.text === metricBefore.text, { before: metricBefore, after: metricResized })
  const overlayFit = await selectionOverlayFitsBlock(page, 's1-metric')
  check('Arrange Mode selection follows autoheight block', overlayFit.fits, overlayFit)

  await clickBlock(page, 's1-title')
  const noEditorInLayout = await page.eval("!document.querySelector('.plain-text-editor')")
  check('Arrange Mode click does not edit text', noEditorInLayout, null)
}

async function runExportScenario(page) {
  await clickMode(page, 'Text')
  const titleBefore = await blockState(page, 's1-title')
  const expectedTitle = `${titleBefore.text} Export`
  await page.eval(`document.querySelector('[data-block="s1-title"]').click()`)
  await page.waitFor("!!document.querySelector('.plain-text-editor[contenteditable]')")
  await page.send('Input.insertText', { text: ' Export' })
  await commitTextEditor(page)

  await clickMode(page, 'Arrange')
  await dragBlock(page, 's1-note', 24, 16)

  await clickToolbar(page, 'Copy HTML')
  await page.waitFor("!!document.querySelector('.export-buffer')")
  const exportState = await page.eval(`(() => {
    const value = document.querySelector('.export-buffer')?.value ?? ''
    const copyButton = document.querySelector('button[aria-label="Copy HTML"]')
    return {
      hasEditedTitle: value.includes(${JSON.stringify(expectedTitle)}),
      hasDataSlide: value.includes('data-slide="slide-1"'),
      hasDataBlock: value.includes('data-block="s1-note"'),
      hasStyleCoordinates: /left:\\d/.test(value) && /top:\\d/.test(value),
      hasRawEditorChrome: value.includes('plain-text-editor') || value.includes('resize-handle'),
      hasCopyAction: !!copyButton,
      hasVisibleRawCodePanel: !!document.querySelector('.export-panel'),
    }
  })()`)
  check('export reflects current text and layout state', exportState.hasEditedTitle && exportState.hasDataSlide && exportState.hasDataBlock && exportState.hasStyleCoordinates && !exportState.hasRawEditorChrome, exportState)
  check('export has a clear copy action', exportState.hasCopyAction, exportState)
  check('export does not expose raw code panel by default', !exportState.hasVisibleRawCodePanel, exportState)

  await page.eval(`document.querySelector('button[aria-label="Copy HTML"]')?.click()`)
  await page.waitFor(`document.querySelector('button[aria-label="Copy HTML"]')?.getAttribute('aria-pressed') === 'true'`)
  const copied = await page.eval(`document.querySelector('button[aria-label="Copy HTML"]')?.getAttribute('aria-pressed') === 'true'`)
  check('copy action gives completion feedback', copied, null)
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
  page.close()
}

async function blockState(page, blockId) {
  return page.eval(`(() => {
    const block = document.querySelector('[data-block="${blockId}"]')
    const rect = block.getBoundingClientRect()
    const undo = document.querySelector('button[aria-label="Undo"]')
    const redo = document.querySelector('button[aria-label="Redo"]')
    const reset = document.querySelector('button[aria-label="Reset"]')
    return {
      text: block.textContent,
      x: rect.x,
      y: rect.y,
      width: rect.width,
      height: rect.height,
      selected: block.dataset.selected,
      editorOpen: !!document.querySelector('.plain-text-editor'),
      undoDisabled: undo?.disabled ?? null,
      redoDisabled: redo?.disabled ?? null,
      resetDisabled: reset?.disabled ?? null,
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
    const editor = document.querySelector('.plain-text-editor')
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

async function editorHeight(page) {
  return page.eval("document.querySelector('.plain-text-editor')?.getBoundingClientRect().height ?? 0")
}

async function textRangeMetrics(page, selector) {
  return page.eval(`(() => {
    const element = document.querySelector(${JSON.stringify(selector)})
    const range = document.createRange()
    range.selectNodeContents(element)
    const textRect = range.getBoundingClientRect()
    const boxRect = element.getBoundingClientRect()

    return {
      boxTop: boxRect.top,
      boxLeft: boxRect.left,
      boxWidth: boxRect.width,
      boxHeight: boxRect.height,
      textTop: textRect.top,
      textLeft: textRect.left,
      textWidth: textRect.width,
      textHeight: textRect.height,
    }
  })()`)
}

async function editorMetrics(page) {
  return page.eval(`(() => {
    const wrapper = document.querySelector('.plain-text-editor')
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
  await page.eval(`Array.from(document.querySelectorAll('.mode-button')).find((button) => button.textContent?.trim() === '${label}')?.click()`)
  await delay(100)
}

async function clickToolbar(page, label) {
  await page.eval(`document.querySelector('button[aria-label="${label}"]')?.click()`)
}

async function clickSlide(page, label) {
  await page.eval(`Array.from(document.querySelectorAll('.slide-thumb')).find((button) => button.textContent?.includes('${label}'))?.click()`)
  await delay(150)
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
  await page.eval("document.querySelector('.plain-text-editor')?.focus()")
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
  await page.waitFor("!document.querySelector('.plain-text-editor')")
}

async function cancelTextEditor(page) {
  await page.eval("document.querySelector('.plain-text-editor')?.focus()")
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
  await page.waitFor("!document.querySelector('.plain-text-editor')")
}

async function replaceEditorText(page, text) {
  await page.eval(`(() => {
    const editor = document.querySelector('.plain-text-editor')
    editor.focus()
    const selection = window.getSelection()
    const range = document.createRange()
    range.selectNodeContents(editor)
    selection.removeAllRanges()
    selection.addRange(range)
  })()`)
  await page.send('Input.insertText', { text })
}

async function clickBlock(page, blockId) {
  const center = await blockCenter(page, blockId)
  await clickAt(page, center)
}

async function dragBlock(page, blockId, dx, dy) {
  const center = await blockCenter(page, blockId)
  await dragFromTo(page, center, { x: center.x + dx, y: center.y + dy })
}

async function blockCenter(page, blockId) {
  return page.eval(`(() => {
    const block = document.querySelector('[data-block="${blockId}"]')
    const rect = block.getBoundingClientRect()
    return { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 }
  })()`)
}

async function clickAt(page, point) {
  await page.send('Input.dispatchMouseEvent', {
    type: 'mouseMoved',
    x: point.x,
    y: point.y,
    button: 'none',
  })
  await page.send('Input.dispatchMouseEvent', {
    type: 'mousePressed',
    x: point.x,
    y: point.y,
    button: 'left',
    buttons: 1,
    clickCount: 1,
  })
  await delay(50)
  await page.send('Input.dispatchMouseEvent', {
    type: 'mouseReleased',
    x: point.x,
    y: point.y,
    button: 'left',
    buttons: 0,
    clickCount: 1,
  })
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
