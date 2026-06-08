import {
  Check,
  Code2,
  Download,
  Play,
  Redo2,
  RotateCcw,
  Type,
  Undo2,
} from 'lucide-react'
import {
  disabledToolbarKeys,
  useManagedToolbarPattern,
} from './apgPatternAdapter'
import type { ResetScope } from './retouchViewState'

type ActionToolbarKey =
  | 'add-text'
  | 'copy-html'
  | 'download-html'
  | 'present'
  | 'redo'
  | 'reset'
  | 'undo'

const ACTION_TOOLBAR_KEYS = [
  'undo',
  'redo',
  'reset',
  'add-text',
  'present',
  'copy-html',
  'download-html',
] as const

type TopbarActionControlsProps = {
  canRedo: boolean
  canReset: boolean
  canUndo: boolean
  copyState: 'copied' | 'failed' | 'idle'
  copyTitle: string
  exportCopied: boolean
  exportDownloaded: boolean
  onCopyExport: () => void
  onDownloadExport: () => void
  onInsertTextBlock: () => void
  onPresent: () => void
  onRedo: () => void
  onReset: () => void
  onUndo: () => void
  resetScope: ResetScope
  resetTitle: string
}

export function TopbarActionControls({
  canRedo,
  canReset,
  canUndo,
  copyState,
  copyTitle,
  exportCopied,
  exportDownloaded,
  onCopyExport,
  onDownloadExport,
  onInsertTextBlock,
  onPresent,
  onRedo,
  onReset,
  onUndo,
  resetScope,
  resetTitle,
}: TopbarActionControlsProps) {
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
  )
}
