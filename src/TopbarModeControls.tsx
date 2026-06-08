import { useActionToolbarPattern } from './apgPatternAdapter'
import type { RetouchMode } from './retouchViewState'

const MODE_TOOLBAR_ACTIONS = [
  { action: 'text', label: 'Text' },
  { action: 'layout', label: 'Arrange' },
] as const

type ModeToolbarKey = (typeof MODE_TOOLBAR_ACTIONS)[number]['action']

export function TopbarModeControls({
  mode,
  onChangeMode,
}: {
  mode: RetouchMode
  onChangeMode: (mode: RetouchMode) => void
}) {
  const modeToolbar = useActionToolbarPattern<ModeToolbarKey>({
    activeKey: mode,
    actions: MODE_TOOLBAR_ACTIONS,
    elementIdPrefix: 'mode-tool-',
    label: 'Mode',
    onSelect: onChangeMode,
    omitPressed: false,
    selectedKeys: [mode],
  })
  const modeToolbarProps = modeToolbar.itemProps

  return (
    <div {...modeToolbar.rootProps} className="mode-toggle">
      {MODE_TOOLBAR_ACTIONS.map(({ action, label }) => (
        <button
          {...modeToolbarProps[action]}
          className="mode-button"
          key={action}
          type="button"
        >
          {label}
        </button>
      ))}
    </div>
  )
}
