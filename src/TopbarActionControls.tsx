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
  useActionToolbarPattern,
} from './apgPatternAdapter'
import type { ResetScope } from './retouchViewState'

const ACTION_TOOLBAR_ACTIONS = [
  { action: 'undo', icon: Undo2, label: 'Undo' },
  { action: 'redo', icon: Redo2, label: 'Redo' },
  { action: 'reset', icon: RotateCcw, label: 'Reset' },
  { action: 'add-text', icon: Type, label: 'Add text box' },
  { action: 'present', icon: Play, label: 'Present' },
  { action: 'copy-html', icon: Code2, label: 'Copy HTML' },
  { action: 'download-html', icon: Download, label: 'Download HTML' },
] as const

type ActionToolbarKey = (typeof ACTION_TOOLBAR_ACTIONS)[number]['action']

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
  const actionViews = ACTION_TOOLBAR_ACTIONS.map(({
    action,
    icon: DefaultIcon,
    label: defaultLabel,
  }) => {
    const label = action === 'reset' ? resetTitle : defaultLabel
    const Icon =
      (action === 'copy-html' && exportCopied) ||
      (action === 'download-html' && exportDownloaded)
        ? Check
        : DefaultIcon
    const title =
      action === 'copy-html'
        ? copyTitle
        : action === 'download-html' && exportDownloaded
          ? 'Downloaded'
          : label

    return {
      action,
      Icon,
      label,
      title,
    }
  })
  const actionCommands = {
    'add-text': onInsertTextBlock,
    'copy-html': onCopyExport,
    'download-html': onDownloadExport,
    present: onPresent,
    redo: onRedo,
    reset: onReset,
    undo: onUndo,
  } satisfies Record<ActionToolbarKey, () => void>
  const actionDisabled = {
    'add-text': false,
    'copy-html': false,
    'download-html': false,
    present: false,
    redo: !canRedo,
    reset: !canReset,
    undo: !canUndo,
  } satisfies Record<ActionToolbarKey, boolean>
  const actionToolbar = useActionToolbarPattern<ActionToolbarKey>({
    actions: actionViews.map(({ action, label }) => ({
      action,
      label,
    })),
    disabledKeys: disabledToolbarKeys<ActionToolbarKey>(actionDisabled),
    elementIdPrefix: 'action-tool-',
    label: 'Actions',
    onSelect: (action) => actionCommands[action](),
  })
  const actionToolbarProps = actionToolbar.itemProps

  return (
    <div {...actionToolbar.rootProps} className="toolbar">
      {actionViews.map(({ action, Icon, label, title }) => (
        <button
          {...actionToolbarProps[action]}
          aria-label={label}
          aria-pressed={
            action === 'copy-html'
              ? exportCopied
              : action === 'download-html'
                ? exportDownloaded
                : undefined
          }
          data-action={action}
          data-copy-state={action === 'copy-html' ? copyState : undefined}
          data-download-state={
            action === 'download-html'
              ? exportDownloaded ? 'downloaded' : 'idle'
              : undefined
          }
          data-reset-scope={action === 'reset' ? resetScope : undefined}
          disabled={actionDisabled[action]}
          key={action}
          title={title}
          type="button"
        >
          <Icon
            aria-hidden="true"
            size={16}
            strokeWidth={Icon === Check ? 2.4 : 2.2}
          />
        </button>
      ))}
    </div>
  )
}
