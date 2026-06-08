import {
  Maximize2,
  ZoomIn,
  ZoomOut,
} from 'lucide-react'
import { Fragment } from 'react'
import {
  disabledToolbarKeys,
  useActionToolbarPattern,
} from './apgPatternAdapter'

const ZOOM_TOOLBAR_ACTIONS = [
  { action: 'out', icon: ZoomOut, label: 'Zoom out' },
  { action: 'in', icon: ZoomIn, label: 'Zoom in' },
  { action: 'fit', icon: Maximize2, label: 'Fit canvas' },
] as const

type ZoomToolbarKey = (typeof ZOOM_TOOLBAR_ACTIONS)[number]['action']

export function TopbarZoomControls({
  canZoomIn,
  canZoomOut,
  onZoomFit,
  onZoomIn,
  onZoomOut,
  zoomLabel,
}: {
  canZoomIn: boolean
  canZoomOut: boolean
  onZoomFit: () => void
  onZoomIn: () => void
  onZoomOut: () => void
  zoomLabel: string
}) {
  const zoomCommands = {
    fit: onZoomFit,
    in: onZoomIn,
    out: onZoomOut,
  } satisfies Record<ZoomToolbarKey, () => void>
  const zoomDisabled = {
    fit: false,
    in: !canZoomIn,
    out: !canZoomOut,
  } satisfies Record<ZoomToolbarKey, boolean>
  const zoomToolbar = useActionToolbarPattern<ZoomToolbarKey>({
    actions: ZOOM_TOOLBAR_ACTIONS,
    disabledKeys: disabledToolbarKeys<ZoomToolbarKey>(zoomDisabled),
    elementIdPrefix: 'zoom-tool-',
    label: 'Canvas zoom',
    onSelect: (action) => zoomCommands[action](),
  })
  const zoomToolbarProps = zoomToolbar.itemProps

  return (
    <div {...zoomToolbar.rootProps} className="zoom-controls">
      {ZOOM_TOOLBAR_ACTIONS.map(({ action, icon: Icon, label }, index) => (
        <Fragment key={action}>
          <button
            {...zoomToolbarProps[action]}
            aria-label={label}
            disabled={zoomDisabled[action]}
            title={label}
            type="button"
          >
            <Icon aria-hidden="true" size={16} strokeWidth={2.2} />
          </button>
          {index === 0 ? (
            <span aria-live="polite" className="zoom-value">
              {zoomLabel}
            </span>
          ) : null}
        </Fragment>
      ))}
    </div>
  )
}
