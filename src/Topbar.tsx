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

type Mode = 'text' | 'layout'

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
  mode: Mode
  onChangeMode: (mode: Mode) => void
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
  resetScope: 'deck' | 'layout' | 'text'
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
  return (
    <header className="topbar">
      <div className="mode-toggle" role="tablist" aria-label="Mode">
        <button
          aria-selected={mode === 'text'}
          className="mode-button"
          onClick={() => onChangeMode('text')}
          role="tab"
          type="button"
        >
          Text
        </button>
        <button
          aria-selected={mode === 'layout'}
          className="mode-button"
          onClick={() => onChangeMode('layout')}
          role="tab"
          type="button"
        >
          Arrange
        </button>
      </div>

      <div className="zoom-controls" role="toolbar" aria-label="Canvas zoom">
        <button
          aria-label="Zoom out"
          disabled={!canZoomOut}
          onClick={onZoomOut}
          title="Zoom out"
          type="button"
        >
          <ZoomOut aria-hidden="true" size={16} strokeWidth={2.2} />
        </button>
        <span aria-live="polite" className="zoom-value">
          {zoomLabel}
        </span>
        <button
          aria-label="Zoom in"
          disabled={!canZoomIn}
          onClick={onZoomIn}
          title="Zoom in"
          type="button"
        >
          <ZoomIn aria-hidden="true" size={16} strokeWidth={2.2} />
        </button>
        <button
          aria-label="Fit canvas"
          onClick={onZoomFit}
          title="Fit canvas"
          type="button"
        >
          <Maximize2 aria-hidden="true" size={16} strokeWidth={2.2} />
        </button>
      </div>

      <div className="toolbar" role="toolbar" aria-label="Actions">
        <button
          aria-label="Undo"
          data-action="undo"
          disabled={!canUndo}
          onClick={onUndo}
          title="Undo"
          type="button"
        >
          <Undo2 aria-hidden="true" size={16} strokeWidth={2.2} />
        </button>
        <button
          aria-label="Redo"
          data-action="redo"
          disabled={!canRedo}
          onClick={onRedo}
          title="Redo"
          type="button"
        >
          <Redo2 aria-hidden="true" size={16} strokeWidth={2.2} />
        </button>
        <button
          aria-label={resetTitle}
          data-action="reset"
          data-reset-scope={resetScope}
          disabled={!canReset}
          onClick={onReset}
          title={resetTitle}
          type="button"
        >
          <RotateCcw aria-hidden="true" size={16} strokeWidth={2.2} />
        </button>
        <button
          aria-label="Add text box"
          data-action="add-text"
          onClick={onInsertTextBlock}
          title="Add text box"
          type="button"
        >
          <Type aria-hidden="true" size={16} strokeWidth={2.2} />
        </button>
        <button
          aria-label="Present"
          data-action="present"
          onClick={onPresent}
          title="Present"
          type="button"
        >
          <Play aria-hidden="true" size={16} strokeWidth={2.2} />
        </button>
        <button
          aria-label="Copy HTML"
          aria-pressed={exportCopied}
          data-action="copy-html"
          data-copy-state={copyState}
          onClick={onCopyExport}
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
          aria-label="Download HTML"
          aria-pressed={exportDownloaded}
          data-action="download-html"
          data-download-state={exportDownloaded ? 'downloaded' : 'idle'}
          onClick={onDownloadExport}
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
