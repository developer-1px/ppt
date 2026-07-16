import assert from 'node:assert/strict'
import { readFileSync, readdirSync } from 'node:fs'
import { join } from 'node:path'

const root = new URL('..', import.meta.url)
const read = (path) => readFileSync(new URL(path, root), 'utf8')

const app = read('src/App.tsx')
const appCSS = read('src/App.css')
const controls = read('src/ui/core/controls.tsx')
const controlsCSS = read('src/ui/core/controls.css')
const indexCSS = read('src/index.css')
const packageJSON = read('package.json')
const shellCSS = read('src/ui/shell/editor-shell.css')
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
  '--ppt-ui-radius-md',
  '--ppt-ui-shadow-floating',
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
assert.match(app, /useExclusiveDisclosure/, 'PPT chrome must use the shared disclosure model')
assert.doesNotMatch(app, /setToolShelfOpen|setViewOptionsOpen|setExportOptionsOpen/, 'Toolbar disclosures must not drift back to independent booleans')
assert.doesNotMatch(app, /className="ppt-(?:button|floating-command|icon-button|slide-action)"/, 'App controls must use UI core primitives')

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
