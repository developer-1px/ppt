import type { ResetScope, RetouchMode } from './retouchViewState'
import { TopbarActionControls } from './TopbarActionControls'
import { TopbarModeControls } from './TopbarModeControls'
import { TopbarZoomControls } from './TopbarZoomControls'
import './Topbar.css'

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
  return (
    <header className="topbar">
      <TopbarModeControls mode={mode} onChangeMode={onChangeMode} />
      <TopbarZoomControls
        canZoomIn={canZoomIn}
        canZoomOut={canZoomOut}
        onZoomFit={onZoomFit}
        onZoomIn={onZoomIn}
        onZoomOut={onZoomOut}
        zoomLabel={zoomLabel}
      />
      <TopbarActionControls
        canRedo={canRedo}
        canReset={canReset}
        canUndo={canUndo}
        copyState={copyState}
        copyTitle={copyTitle}
        exportCopied={exportCopied}
        exportDownloaded={exportDownloaded}
        onCopyExport={onCopyExport}
        onDownloadExport={onDownloadExport}
        onInsertTextBlock={onInsertTextBlock}
        onPresent={onPresent}
        onRedo={onRedo}
        onReset={onReset}
        onUndo={onUndo}
        resetScope={resetScope}
        resetTitle={resetTitle}
      />
    </header>
  )
}
