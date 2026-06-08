import { useManagedToolbarPattern } from './apgPatternAdapter'
import type { RetouchMode } from './retouchViewState'

type ModeToolbarKey = RetouchMode

const MODE_TOOLBAR_KEYS = ['text', 'layout'] as const

export function TopbarModeControls({
  mode,
  onChangeMode,
}: {
  mode: RetouchMode
  onChangeMode: (mode: RetouchMode) => void
}) {
  const modeToolbar = useManagedToolbarPattern<ModeToolbarKey>({
    activeKey: mode,
    elementIdPrefix: 'mode-tool-',
    handlers: {
      layout: () => onChangeMode('layout'),
      text: () => onChangeMode('text'),
    },
    items: {
      layout: { label: 'Arrange' },
      text: { label: 'Text' },
    },
    label: 'Mode',
    rootKeys: MODE_TOOLBAR_KEYS,
    selectedKeys: [mode],
  })
  const modeToolbarProps = modeToolbar.itemProps

  return (
    <div {...modeToolbar.rootProps} className="mode-toggle">
      <button
        {...modeToolbarProps.text}
        className="mode-button"
        type="button"
      >
        Text
      </button>
      <button
        {...modeToolbarProps.layout}
        className="mode-button"
        type="button"
      >
        Arrange
      </button>
    </div>
  )
}
