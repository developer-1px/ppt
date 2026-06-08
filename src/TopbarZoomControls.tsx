import {
  Maximize2,
  ZoomIn,
  ZoomOut,
} from 'lucide-react'
import {
  disabledToolbarKeys,
  useManagedToolbarPattern,
} from './apgPatternAdapter'

type ZoomToolbarKey = 'fit' | 'in' | 'out'

const ZOOM_TOOLBAR_KEYS = ['out', 'in', 'fit'] as const

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
  const zoomToolbar = useManagedToolbarPattern<ZoomToolbarKey>({
    disabledKeys: disabledToolbarKeys<ZoomToolbarKey>([
      ['out', !canZoomOut],
      ['in', !canZoomIn],
    ]),
    elementIdPrefix: 'zoom-tool-',
    handlers: {
      fit: onZoomFit,
      in: onZoomIn,
      out: onZoomOut,
    },
    items: {
      fit: { label: 'Fit canvas' },
      in: { label: 'Zoom in' },
      out: { label: 'Zoom out' },
    },
    label: 'Canvas zoom',
    omitPressed: true,
    rootKeys: ZOOM_TOOLBAR_KEYS,
  })
  const zoomToolbarProps = zoomToolbar.itemProps

  return (
    <div {...zoomToolbar.rootProps} className="zoom-controls">
      <button
        {...zoomToolbarProps.out}
        aria-label="Zoom out"
        disabled={!canZoomOut}
        title="Zoom out"
        type="button"
      >
        <ZoomOut aria-hidden="true" size={16} strokeWidth={2.2} />
      </button>
      <span aria-live="polite" className="zoom-value">
        {zoomLabel}
      </span>
      <button
        {...zoomToolbarProps.in}
        aria-label="Zoom in"
        disabled={!canZoomIn}
        title="Zoom in"
        type="button"
      >
        <ZoomIn aria-hidden="true" size={16} strokeWidth={2.2} />
      </button>
      <button
        {...zoomToolbarProps.fit}
        aria-label="Fit canvas"
        title="Fit canvas"
        type="button"
      >
        <Maximize2 aria-hidden="true" size={16} strokeWidth={2.2} />
      </button>
    </div>
  )
}
