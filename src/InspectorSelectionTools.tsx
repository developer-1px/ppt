import {
  AlignCenterHorizontal,
  AlignCenterVertical,
  AlignEndHorizontal,
  AlignEndVertical,
  AlignHorizontalSpaceBetween,
  AlignStartHorizontal,
  AlignStartVertical,
  AlignVerticalSpaceBetween,
  BringToFront,
  MoveDown,
  MoveUp,
  SendToBack,
  type LucideIcon,
} from 'lucide-react'
import type {
  AlignSelectionAction,
  DistributeSelectionAction,
} from './selectionAlignment'
import type { LayerOrderAction } from './selectionLayerOrder'
import { useActionToolbarPattern } from './apgPatternAdapter'

type SelectionToolAction<TKey extends string> = {
  action: TKey
  icon: LucideIcon
  label: string
}

const ALIGNMENT_ACTIONS: SelectionToolAction<AlignSelectionAction>[] = [
  { action: 'left', icon: AlignStartVertical, label: 'Align left' },
  { action: 'center-x', icon: AlignCenterVertical, label: 'Align horizontal center' },
  { action: 'right', icon: AlignEndVertical, label: 'Align right' },
  { action: 'top', icon: AlignStartHorizontal, label: 'Align top' },
  { action: 'middle-y', icon: AlignCenterHorizontal, label: 'Align vertical middle' },
  { action: 'bottom', icon: AlignEndHorizontal, label: 'Align bottom' },
]

const LAYER_ORDER_ACTIONS: SelectionToolAction<LayerOrderAction>[] = [
  { action: 'front', icon: BringToFront, label: 'Bring to front' },
  { action: 'forward', icon: MoveUp, label: 'Bring forward' },
  { action: 'backward', icon: MoveDown, label: 'Send backward' },
  { action: 'back', icon: SendToBack, label: 'Send to back' },
]

const DISTRIBUTION_ACTIONS: SelectionToolAction<DistributeSelectionAction>[] = [
  {
    action: 'horizontal',
    icon: AlignHorizontalSpaceBetween,
    label: 'Distribute horizontal spacing',
  },
  {
    action: 'vertical',
    icon: AlignVerticalSpaceBetween,
    label: 'Distribute vertical spacing',
  },
]

function SelectionToolGroup<TKey extends string>({
  actions,
  className,
  disabled = false,
  elementIdPrefix,
  label,
  onSelect,
}: {
  actions: readonly SelectionToolAction<TKey>[]
  className: string
  disabled?: boolean
  elementIdPrefix: string
  label: string
  onSelect: (action: TKey) => void
}) {
  const toolbar = useActionToolbarPattern<TKey>({
    actions,
    disabledKeys: disabled ? actions.map(({ action }) => action) : undefined,
    elementIdPrefix,
    label,
    onSelect,
  })

  return (
    <div {...toolbar.rootProps} className={className}>
      {actions.map(({ action, icon: Icon, label: actionLabel }) => (
        <button
          {...toolbar.itemProps[action]}
          aria-label={actionLabel}
          disabled={disabled}
          key={action}
          title={actionLabel}
          type="button"
        >
          <Icon aria-hidden="true" size={15} strokeWidth={2.2} />
        </button>
      ))}
    </div>
  )
}

export function AlignmentTools({
  onAlignSelection,
}: {
  onAlignSelection: (action: AlignSelectionAction) => void
}) {
  return (
    <SelectionToolGroup
      actions={ALIGNMENT_ACTIONS}
      className="alignment-tools"
      elementIdPrefix="alignment-tool-"
      label="Alignment"
      onSelect={onAlignSelection}
    />
  )
}

export function DistributionTools({
  disabled,
  onDistributeSelection,
}: {
  disabled: boolean
  onDistributeSelection: (action: DistributeSelectionAction) => void
}) {
  return (
    <SelectionToolGroup
      actions={DISTRIBUTION_ACTIONS}
      className="distribution-tools"
      disabled={disabled}
      elementIdPrefix="distribution-tool-"
      label="Distribution"
      onSelect={onDistributeSelection}
    />
  )
}

export function LayerOrderTools({
  onLayerOrderChange,
}: {
  onLayerOrderChange: (action: LayerOrderAction) => void
}) {
  return (
    <SelectionToolGroup
      actions={LAYER_ORDER_ACTIONS}
      className="layer-tools"
      elementIdPrefix="layer-tool-"
      label="Layer order"
      onSelect={onLayerOrderChange}
    />
  )
}
