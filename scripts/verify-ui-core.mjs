import assert from 'node:assert/strict'
import { readFileSync, readdirSync } from 'node:fs'
import { join } from 'node:path'

const root = new URL('..', import.meta.url)
const read = (path) => readFileSync(new URL(path, root), 'utf8')

const app = read('src/App.tsx')
const appCSS = read('src/App.css')
const commandSurface = read('src/ui/command-surface/PPTSurfaceCommands.tsx')
const commandSurfaceModel = read('src/ui/command-surface/PPTSurfaceCommandModel.ts')
const controls = read('src/ui/core/controls.tsx')
const controlsCSS = read('src/ui/core/controls.css')
const editorChrome = read('src/ui/core/useEditorChrome.ts')
const editorShell = read('src/ui/shell/EditorShell.tsx')
const inspectorActionDispatcher = read('src/ui/inspector/PPTInspectorActionDispatcher.ts')
const inspectorContract = read('src/ui/inspector/PPTInspectorContract.ts')
const inspectorCSS = read('src/ui/inspector/inspector.css')
const inspectorIndex = read('src/ui/inspector/index.ts')
const layerPane = read('src/ui/inspector/PPTLayerPane.tsx')
const layerPaneAdapter = read('src/pptLayerPaneAdapter.ts')
const slideInspectorPanel = read('src/ui/inspector/PPTSlideInspectorPanel.tsx')
const indexCSS = read('src/index.css')
const packageJSON = read('package.json')
const alignmentPopover = read('src/ui/selection-toolbar/PPTAlignmentPopover.tsx')
const paragraphAlign = read('src/ui/text-formatting/PPTParagraphAlignRadioGroup.tsx')
const selectionToolbar = read('src/ui/selection-toolbar/PPTSelectionToolbar.tsx')
const selectionToolbarCSS = read('src/ui/selection-toolbar/selection-toolbar.css')
const shellCSS = read('src/ui/shell/editor-shell.css')
const shapeKindMenu = read('src/ui/selection-toolbar/PPTShapeKindMenu.tsx')
const textFormattingCSS = read('src/ui/text-formatting/text-formatting.css')
const textQuickFormat = read('src/ui/selection-toolbar/PPTTextQuickFormatControls.tsx')
const tokensCSS = read('src/ui/core/tokens.css')
const viteConfig = read('vite.config.ts')

const expectedImports = [
  "@import './ui/core/tokens.css';",
  "@import './ui/core/controls.css';",
  "@import './ui/shell/editor-shell.css';",
]

assert.deepEqual(
  appCSS.split('\n').slice(0, expectedImports.length),
  expectedImports,
  'App.css must load UI tokens, primitives, and shell styles in that order',
)

for (const token of [
  '--ppt-ui-control-size',
  '--ppt-ui-control-size-compact',
  '--ppt-ui-control-gap',
  '--ppt-ui-radius-md',
  '--ppt-ui-shadow-floating',
  '--ppt-ui-z-selection-toolbar',
  '--ppt-ui-z-toolbar',
]) {
  assert.match(tokensCSS, new RegExp(token), `Missing UI token: ${token}`)
}

assert.match(controls, /label: string/, 'IconButton must require an accessible label')
assert.match(controlsCSS, /\.ppt-icon-button/, 'Control styles must own icon buttons')
assert.doesNotMatch(appCSS, /^\.ppt-icon-button\s*[,\{]/m, 'App.css must not redefine core icon buttons')
assert.doesNotMatch(appCSS, /^\.ppt-floating-command\s*[,\{]/m, 'App.css must not redefine floating controls')
assert.doesNotMatch(appCSS, /\.ppt-topbar\s*\{/, 'App.css must not redefine the editor shell')
assert.match(shellCSS, /\.ppt-topbar\s*\{/, 'Shell styles must own the editor toolbar')
assert.match(app, /<EditorShell/, 'PPT must compose its root through the editor shell seam')
assert.match(app, /<EditorToolbar/, 'PPT must compose its toolbar through the editor shell seam')
const inspectorCall = app.match(/<PPTInspector[\s\S]*?\/>/)?.[0] ?? ''
assert.match(inspectorCall, /model=\{\{/, 'App must pass one cohesive Inspector model')
assert.match(inspectorCall, /onAction=\{handleInspectorAction\}/, 'App must translate one Inspector action stream')
assert.doesNotMatch(inspectorCall, /on(?:Comment|Commit|Copy|Color|Download|Element|Image|Layer|Layout|Line|Object|Paragraph|Shape|Slide|Table|Text)[A-Z]/, 'App must not regress to field-level Inspector callbacks')
assert.match(app, /function PPTInspector\(\{ model, onAction \}: PPTInspectorProps\)/, 'PPT Inspector must consume its model and action contract')
assert.match(app, /function handleInspectorAction\(action: PPTInspectorAction\)/, 'App must own Inspector action translation')
assert.match(inspectorContract, /type PPTInspectorProps = \{[\s\S]*model: PPTInspectorModel[\s\S]*onAction: \(action: PPTInspectorAction\) => void[\s\S]*\}/, 'Inspector must expose one model and one action interface')
assert.match(inspectorContract, /type PPTInspectorAction =[\s\S]*type: 'set-slide-transition'/, 'Inspector action contract must cover slide and selection editing')
assert.match(inspectorActionDispatcher, /function createPPTInspectorActionDispatcher/, 'Inspector dispatcher must own field event to action translation')
assert.doesNotMatch(inspectorContract, /function createPPTInspectorActionDispatcher/, 'Inspector contract must remain a pure type surface')
assert.doesNotMatch(app, /function onElement(?:Geometry|TextStyle|Shadow)Change/, 'App must not own Inspector field event adapters')
assert.match(app, /data-ppt-inspector-tabpanel="slide"[\s\S]*<PPTSlideInspectorPanel model=\{model\} onAction=\{onAction\} \/>/, 'PPT Inspector must compose the Slide panel through its model and action seam')
assert.doesNotMatch(app, /data-ppt-slide-field="(?:name|background|layout|notes)"/, 'App must not own Slide inspector fields')
assert.doesNotMatch(app, /onSlide(?:Name|Background|Layout|Notes|Transition)Change/, 'App must not own Slide inspector field event adapters')
assert.match(slideInspectorPanel, /type PPTSlideInspectorPanelProps = Pick<[\s\S]*'model' \| 'onAction'/, 'Slide Inspector panel must expose only the Inspector model and action seam')
assert.match(slideInspectorPanel, /createPPTInspectorActionDispatcher\(onAction\)/, 'Slide Inspector panel must translate controls through the Inspector action dispatcher')
assert.match(slideInspectorPanel, /createSlideEditTransitionDescriptor/, 'Slide Inspector panel must reuse the Canvas-selected transition affordance')
assert.match(slideInspectorPanel, /data-ppt-slide-metadata-field/, 'Slide Inspector panel must own slide metadata field rendering')
assert.match(app, /<PPTLayerPane model=\{model\} onAction=\{onAction\} \/>/, 'PPT Inspector must compose the Layer Pane through its model and action seam')
assert.doesNotMatch(app, /data-ppt-layer-pane(?:=|\s*\n)/, 'App must not own Layer Pane rendering')
assert.doesNotMatch(app, /function (?:runLayerPaneIntent|handleLayerPaneRow|startLayerPaneRename)/, 'App must not own Layer Pane interactions')
assert.doesNotMatch(app, /layerPaneGroupState|layerPaneRenameState|layerPaneDragState/, 'App must not own Layer Pane local state')
assert.match(layerPane, /type PPTLayerPaneProps = Pick<PPTInspectorProps, 'model' \| 'onAction'>/, 'Layer Pane must expose only the Inspector model and action seam')
assert.match(layerPane, /createPPTInspectorActionDispatcher\(onAction\)/, 'Layer Pane must translate interactions through the Inspector action dispatcher')
assert.match(layerPane, /getSlideEditLayerPaneKeyboardIntent/, 'Layer Pane must reuse the Canvas-selected keyboard affordance')
assert.match(layerPane, /getSlideEditLayerPaneDropIndicator/, 'Layer Pane must reuse the Canvas-selected drag affordance')
assert.match(layerPane, /getPPTCanvasSelectionListModifierState/, 'Layer Pane must reuse the Canvas selection modifier affordance')
assert.match(layerPaneAdapter, /export function createPPTLayerPaneDescriptor/, 'Layer Pane adapter must own PPT row projection')
assert.match(layerPaneAdapter, /export function getPPTLayerPaneSelection/, 'Layer Pane adapter must own PPT selection translation')
assert.match(layerPaneAdapter, /export function reorderPPTLayerPaneElement/, 'Layer Pane adapter must own PPT reorder translation')
assert.match(inspectorIndex, /import '\.\/inspector\.css'/, 'Inspector public module must load its co-located styles')
for (const selector of [
  '.ppt-panel-section',
  '.ppt-field',
  '.ppt-slide-transition-fields',
  '.ppt-layer-list',
  '.ppt-export-code',
]) {
  const escapedSelector = selector.replace('.', '\\.')
  assert.match(inspectorCSS, new RegExp(`^${escapedSelector}(?:\\s|,|\\{)`, 'm'), `Inspector styles must own ${selector}`)
  assert.doesNotMatch(appCSS, new RegExp(`^${escapedSelector}(?:\\s|,|\\{)`, 'm'), `App.css must not own ${selector}`)
}
assert.match(app, /<PPTSelectionToolbar[\s\S]*model=\{\{[\s\S]*onAction=\{handleSelectionToolbarAction\}/, 'App must cross the selection toolbar seam through model and action')
assert.doesNotMatch(app, /<PPTAlignmentPopover(?:\s|\/)|<PPTShapeKindMenu(?:\s|\/)|<PPTTextQuickFormatControls(?:\s|\/)/, 'App must not compose selection toolbar internals')
assert.doesNotMatch(app, /function PPTSelectionFloatingBar|function PPTTextQuickFormatControls|function PPTParagraphAlignRadioGroup/, 'App must not own selection toolbar or text control implementation')
assert.doesNotMatch(app, /function PPTAlignmentPopover/, 'App must not own alignment menu implementation')
assert.doesNotMatch(app, /function PPTShapeKindMenu/, 'App must not own shape menu implementation')
assert.doesNotMatch(app, /PPT_ALIGNMENT_POPOVER_COMMANDS|PPT_SHAPE_MENU_OPTIONS|PPT_COMMAND_SURFACE_GROUPS/, 'App must not own UI command registries')
assert.match(app, /useEditorChrome/, 'PPT must use the shared editor chrome module')
assert.doesNotMatch(app, /useExclusiveDisclosure/, 'PPT must not compose chrome from a shallow disclosure hook')
assert.doesNotMatch(app, /setToolShelfOpen|setViewOptionsOpen|setExportOptionsOpen|setShowGrid|setShowFrameGuides|setShowMinimap|setInspectorOpen/, 'Editor chrome must not drift back to independent booleans')
assert.doesNotMatch(app, /className="ppt-(?:button|floating-command|icon-button|slide-action)"/, 'App controls must use UI core primitives')
assert.match(editorChrome, /useReducer/, 'Editor chrome policy must be centralized in one reducer')
assert.match(editorChrome, /activeTransientSurface: null/, 'Transient editor chrome must start closed')
assert.match(editorChrome, /inspectorOpen: false/, 'Inspector must start closed')
assert.match(editorChrome, /frameGuides: false[\s\S]*grid: false[\s\S]*minimap: false/, 'Secondary view features must start hidden')
assert.match(editorChrome, /event\.key === 'Escape'/, 'Escape must dismiss transient editor chrome')
assert.match(editorShell, /data-editor-inspector-open/, 'Editor shell must own inspector layout state')
assert.match(editorShell, /data-editor-transient-surface/, 'Editor toolbar must expose one transient surface state')
assert.match(shellCSS, /data-editor-inspector-open/, 'Shell CSS must consume the generic inspector state')
assert.match(shellCSS, /data-editor-transient-surface/, 'Shell CSS must consume the generic transient surface state')
assert.match(shellCSS, /\.ppt-app\s*\{[^}]*position: relative/, 'Editor shell must establish the mobile Inspector containing block')
assert.match(shellCSS, /@media \(max-width: 920px\)[\s\S]*?\.ppt-inspector\s*\{[^}]*position: absolute/, 'Mobile Inspector must overlay the editor inside its shell')
assert.doesNotMatch(shellCSS, /data-ppt-(?:inspector-open|tool-shelf-open|view-options-open|export-options-open)/, 'Shell layout must not depend on PPT diagnostic state')
assert.match(selectionToolbar, /type PPTSelectionToolbarProps = \{[\s\S]*model: PPTSelectionToolbarModel[\s\S]*onAction: \(action: PPTSelectionToolbarAction\) => void[\s\S]*\}/, 'Selection toolbar must expose one model and one action interface')
assert.match(selectionToolbar, /<PPTAlignmentPopover/, 'Selection toolbar must compose the alignment popover')
assert.match(selectionToolbar, /<PPTShapeKindMenu/, 'Selection toolbar must compose the shape menu')
assert.match(selectionToolbar, /<PPTTextQuickFormatControls/, 'Selection toolbar must compose quick text formatting')
assert.doesNotMatch(selectionToolbar, /onTextBoldToggle|onTextItalicToggle|onFontSizeStep|onParagraphAlign/, 'Selection toolbar must not regress to callback-per-control')
assert.match(selectionToolbar, /import '\.\/selection-toolbar\.css'/, 'Selection toolbar must load its co-located styles')
assert.match(selectionToolbarCSS, /\.ppt-selection-floating-bar\s*\{/, 'Selection toolbar styles must own the floating bar')
assert.doesNotMatch(appCSS, /\.ppt-selection-floating-bar\s*\{/, 'App.css must not own selection toolbar styles')
assert.match(commandSurfaceModel, /const COMMAND_GROUPS:/, 'Command surface model must own the command registry')
assert.match(commandSurfaceModel, /export function getPPTCommandSurfaceGroups/, 'Command surface model must derive surface-specific command views')
assert.match(commandSurface, /export function PPTSurfaceCommandButton/, 'Command surface must own command rendering')
assert.doesNotMatch(app, /function PPTSurfaceCommandButton|function getPPTCommandSurfaceGroups/, 'App must consume the command surface module')
assert.match(textQuickFormat, /onAction: \(action: PPTTextQuickFormatAction\) => void/, 'Quick text formatting must emit one action stream')
assert.match(paragraphAlign, /PPT_RADIO_GROUP_FOCUS_MODEL/, 'Paragraph alignment must own its Canvas focus affordance')
assert.match(paragraphAlign, /handlePPTCanvasRadioGroupKeyDown/, 'Paragraph alignment must own keyboard behavior')
assert.match(textFormattingCSS, /grid-template-columns: repeat\(4, minmax\(0, 1fr\)\)/, 'Paragraph alignment must reserve one stable column per option')
assert.doesNotMatch(appCSS, /\.ppt-segmented-control\s*\{/, 'App.css must not own shared text formatting controls')
assert.match(alignmentPopover, /availability: Readonly<Record<PPTAlignmentPopoverCommand, boolean>>/, 'Alignment menu must receive availability through its public interface')
assert.match(alignmentPopover, /onPreviewChange: \(command: PPTAlignmentPopoverCommand \| null\) => void/, 'Alignment menu must own preview lifecycle behind one callback')
assert.match(shapeKindMenu, /state: PPTShapeQuickMenuState/, 'Shape menu must receive a cohesive state value')

for (const [name, source] of [
  ['alignment popover', alignmentPopover],
  ['shape kind menu', shapeKindMenu],
]) {
  assert.match(source, /usePPTCanvasMenuRovingFocus/, `${name} must own roving menu focus`)
  assert.match(source, /focusPPTCanvasElementOnNextFrame/, `${name} must restore trigger focus`)
  assert.match(source, /getPPTCanvasMenuTriggerKeyboardIntent/, `${name} must reuse the Canvas menu trigger affordance`)
}

for (const [name, source] of [
  ['package.json', packageJSON],
  ['src/index.css', indexCSS],
  ['vite.config.ts', viteConfig],
]) {
  assert.doesNotMatch(source, /tailwind/i, `${name} must not introduce a second styling system`)
}

const coreDirectory = new URL('src/ui/core/', root)
for (const fileName of readdirSync(coreDirectory)) {
  if (!/\.(ts|tsx)$/.test(fileName)) {
    continue
  }

  const source = readFileSync(join(coreDirectory.pathname, fileName), 'utf8')
  assert.doesNotMatch(
    source,
    /from ['"](?:\.\.\/){2}|from ['"].*(?:ppt|canvas)/i,
    `UI core must not depend on PPT or canvas implementation: ${fileName}`,
  )
}

console.log('UI core structure verified')
