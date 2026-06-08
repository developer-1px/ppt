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
} from 'lucide-react'
import type {
  AlignSelectionAction,
  DistributeSelectionAction,
} from './selectionAlignment'
import type { LayerOrderAction } from './selectionLayerOrder'
import { useActionToolbarPattern } from './apgPatternAdapter'

const ALIGNMENT_ACTIONS: {
  action: AlignSelectionAction
  icon: typeof AlignStartVertical
  label: string
}[] = [
  { action: 'left', icon: AlignStartVertical, label: 'Align left' },
  { action: 'center-x', icon: AlignCenterVertical, label: 'Align horizontal center' },
  { action: 'right', icon: AlignEndVertical, label: 'Align right' },
  { action: 'top', icon: AlignStartHorizontal, label: 'Align top' },
  { action: 'middle-y', icon: AlignCenterHorizontal, label: 'Align vertical middle' },
  { action: 'bottom', icon: AlignEndHorizontal, label: 'Align bottom' },
]

const LAYER_ORDER_ACTIONS: {
  action: LayerOrderAction
  icon: typeof BringToFront
  label: string
}[] = [
  { action: 'front', icon: BringToFront, label: 'Bring to front' },
  { action: 'forward', icon: MoveUp, label: 'Bring forward' },
  { action: 'backward', icon: MoveDown, label: 'Send backward' },
  { action: 'back', icon: SendToBack, label: 'Send to back' },
]

const DISTRIBUTION_ACTIONS: {
  action: DistributeSelectionAction
  icon: typeof AlignHorizontalSpaceBetween
  label: string
}[] = [
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

export function AlignmentTools({
  onAlignSelection,
}: {
  onAlignSelection: (action: AlignSelectionAction) => void
}) {
  const toolbar = useActionToolbarPattern<AlignSelectionAction>({
    actions: ALIGNMENT_ACTIONS,
    elementIdPrefix: 'alignment-tool-',
    label: 'Alignment',
    onSelect: onAlignSelection,
  })

  return (
    <div {...toolbar.rootProps} className="alignment-tools">
      {ALIGNMENT_ACTIONS.map(({ action, icon: Icon, label }) => (
        <button
          {...toolbar.itemProps[action]}
          aria-label={label}
          key={action}
          title={label}
          type="button"
        >
          <Icon aria-hidden="true" size={15} strokeWidth={2.2} />
        </button>
      ))}
    </div>
  )
}

export function DistributionTools({
  disabled,
  onDistributeSelection,
}: {
  disabled: boolean
  onDistributeSelection: (action: DistributeSelectionAction) => void
}) {
  const toolbar = useActionToolbarPattern<DistributeSelectionAction>({
    actions: DISTRIBUTION_ACTIONS,
    disabledKeys: disabled ? ['horizontal', 'vertical'] : undefined,
    elementIdPrefix: 'distribution-tool-',
    label: 'Distribution',
    onSelect: onDistributeSelection,
  })

  return (
    <div {...toolbar.rootProps} className="distribution-tools">
      {DISTRIBUTION_ACTIONS.map(({ action, icon: Icon, label }) => (
        <button
          {...toolbar.itemProps[action]}
          aria-label={label}
          disabled={disabled}
          key={action}
          title={label}
          type="button"
        >
          <Icon aria-hidden="true" size={15} strokeWidth={2.2} />
        </button>
      ))}
    </div>
  )
}

export function LayerOrderTools({
  onLayerOrderChange,
}: {
  onLayerOrderChange: (action: LayerOrderAction) => void
}) {
  const toolbar = useActionToolbarPattern<LayerOrderAction>({
    actions: LAYER_ORDER_ACTIONS,
    elementIdPrefix: 'layer-tool-',
    label: 'Layer order',
    onSelect: onLayerOrderChange,
  })

  return (
    <div {...toolbar.rootProps} className="layer-tools">
      {LAYER_ORDER_ACTIONS.map(({ action, icon: Icon, label }) => (
        <button
          {...toolbar.itemProps[action]}
          aria-label={label}
          key={action}
          title={label}
          type="button"
        >
          <Icon aria-hidden="true" size={15} strokeWidth={2.2} />
        </button>
      ))}
    </div>
  )
}
