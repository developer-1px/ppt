import {
  Check,
  Code2,
  Download,
  Maximize2,
  Play,
  Redo2,
  RotateCcw,
  Type,
  Undo2,
  ZoomIn,
  ZoomOut,
} from 'lucide-react'
import {
  disabledToolbarKeys,
  useManagedToolbarPattern,
} from './apgPatternAdapter'
import type { ResetScope, RetouchMode } from './retouchViewState'
import './Topbar.css'

type ModeToolbarKey = RetouchMode
type ZoomToolbarKey = 'fit' | 'in' | 'out'
type ActionToolbarKey =
  | 'add-text'
  | 'copy-html'
  | 'download-html'
  | 'present'
  | 'redo'
  | 'reset'
  | 'undo'

const MODE_TOOLBAR_KEYS = ['text', 'layout'] as const
const ZOOM_TOOLBAR_KEYS = ['out', 'in', 'fit'] as const
const ACTION_TOOLBAR_KEYS = [
  'undo',
  'redo',
  'reset',
  'add-text',
  'present',
  'copy-html',
  'download-html',
] as const

type TopbarProps = {
  canRedo: boolean
  canReset: boolean
  canUndo: boolean
  canZoomIn: boolean
  canZoomOut: boolean
  copyState: 'copied' | 'failed' | 'idle'
  copyTitle: string
  exportCopied: boolean
  exportDownloaded: boolean
  mode: RetouchMode
  onChangeMode: (mode: RetouchMode) => void
  onCopyExport: () => void
  onDownloadExport: () => void
  onInsertTextBlock: () => void
  onPresent: () => void
  onRedo: () => void
  onReset: () => void
  onUndo: () => void
  onZoomFit: () => void
  onZoomIn: () => void
  onZoomOut: () => void
  resetScope: ResetScope
  resetTitle: string
  zoomLabel: string
}

export function Topbar({
  canRedo,
  canReset,
  canUndo,
  canZoomIn,
  canZoomOut,
  copyState,
  copyTitle,
  exportCopied,
  exportDownloaded,
  mode,
  onChangeMode,
  onCopyExport,
  onDownloadExport,
  onInsertTextBlock,
  onPresent,
  onRedo,
  onReset,
  onUndo,
  onZoomFit,
  onZoomIn,
  onZoomOut,
  resetScope,
  resetTitle,
  zoomLabel,
}: TopbarProps) {
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
  const actionToolbar = useManagedToolbarPattern<ActionToolbarKey>({
    disabledKeys: disabledToolbarKeys<ActionToolbarKey>([
      ['undo', !canUndo],
      ['redo', !canRedo],
      ['reset', !canReset],
    ]),
    elementIdPrefix: 'action-tool-',
    handlers: {
      'add-text': onInsertTextBlock,
      'copy-html': onCopyExport,
      'download-html': onDownloadExport,
      present: onPresent,
      redo: onRedo,
      reset: onReset,
      undo: onUndo,
    },
    items: {
      'add-text': { label: 'Add text box' },
      'copy-html': { label: 'Copy HTML' },
      'download-html': { label: 'Download HTML' },
      present: { label: 'Present' },
      redo: { label: 'Redo' },
      reset: { label: resetTitle },
      undo: { label: 'Undo' },
    },
    label: 'Actions',
    omitPressed: true,
    rootKeys: ACTION_TOOLBAR_KEYS,
  })
  const actionToolbarProps = actionToolbar.itemProps

  return (
    <header className="topbar">
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

      <div {...actionToolbar.rootProps} className="toolbar">
        <button
          {...actionToolbarProps.undo}
          aria-label="Undo"
          data-action="undo"
          disabled={!canUndo}
          title="Undo"
          type="button"
        >
          <Undo2 aria-hidden="true" size={16} strokeWidth={2.2} />
        </button>
        <button
          {...actionToolbarProps.redo}
          aria-label="Redo"
          data-action="redo"
          disabled={!canRedo}
          title="Redo"
          type="button"
        >
          <Redo2 aria-hidden="true" size={16} strokeWidth={2.2} />
        </button>
        <button
          {...actionToolbarProps.reset}
          aria-label={resetTitle}
          data-action="reset"
          data-reset-scope={resetScope}
          disabled={!canReset}
          title={resetTitle}
          type="button"
        >
          <RotateCcw aria-hidden="true" size={16} strokeWidth={2.2} />
        </button>
        <button
          {...actionToolbarProps['add-text']}
          aria-label="Add text box"
          data-action="add-text"
          title="Add text box"
          type="button"
        >
          <Type aria-hidden="true" size={16} strokeWidth={2.2} />
        </button>
        <button
          {...actionToolbarProps.present}
          aria-label="Present"
          data-action="present"
          title="Present"
          type="button"
        >
          <Play aria-hidden="true" size={16} strokeWidth={2.2} />
        </button>
        <button
          {...actionToolbarProps['copy-html']}
          aria-label="Copy HTML"
          aria-pressed={exportCopied}
          data-action="copy-html"
          data-copy-state={copyState}
          title={copyTitle}
          type="button"
        >
          {exportCopied ? (
            <Check aria-hidden="true" size={16} strokeWidth={2.4} />
          ) : (
            <Code2 aria-hidden="true" size={16} strokeWidth={2.2} />
          )}
        </button>
        <button
          {...actionToolbarProps['download-html']}
          aria-label="Download HTML"
          aria-pressed={exportDownloaded}
          data-action="download-html"
          data-download-state={exportDownloaded ? 'downloaded' : 'idle'}
          title={exportDownloaded ? 'Downloaded' : 'Download HTML'}
          type="button"
        >
          {exportDownloaded ? (
            <Check aria-hidden="true" size={16} strokeWidth={2.4} />
          ) : (
            <Download aria-hidden="true" size={16} strokeWidth={2.2} />
          )}
        </button>
      </div>
    </header>
  )
}
