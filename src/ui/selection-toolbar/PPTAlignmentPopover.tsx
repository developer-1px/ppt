import {
  AlignCenterHorizontal,
  AlignCenterVertical,
  AlignEndVertical,
  AlignHorizontalDistributeCenter,
  AlignLeft,
  AlignRight,
  AlignStartVertical,
  AlignVerticalDistributeCenter,
} from 'lucide-react'
import {
  useEffect,
  useRef,
  useState,
  type KeyboardEvent as ReactKeyboardEvent,
} from 'react'
import {
  PPT_MENU_ITEM_PROPS,
  focusPPTCanvasElementOnNextFrame,
  getPPTCanvasMenuTriggerKeyboardIntent,
  usePPTCanvasMenuRovingFocus,
} from '../../pptCanvasAppAffordanceAdapter'
import { CANVAS_DOM_ALIGNMENT_POPOVER_MODEL } from '../../pptDomEditAffordanceAdapter'
import { IconButton } from '../core'

export type PPTAlignmentPopoverCommand =
  | 'alignBottom'
  | 'alignCenter'
  | 'alignLeft'
  | 'alignMiddle'
  | 'alignRight'
  | 'alignTop'
  | 'distributeHorizontal'
  | 'distributeVertical'

export type PPTAlignmentPopoverProps = {
  availability: Readonly<Record<PPTAlignmentPopoverCommand, boolean>>
  onCommand: (command: PPTAlignmentPopoverCommand) => void
  onPreviewChange: (command: PPTAlignmentPopoverCommand | null) => void
}

const ALIGNMENT_COMMANDS = [{
  command: 'alignLeft',
  dataCommand: 'align-left',
  label: 'Align left',
}, {
  command: 'alignCenter',
  dataCommand: 'align-center-x',
  label: 'Align center',
}, {
  command: 'alignRight',
  dataCommand: 'align-right',
  label: 'Align right',
}, {
  command: 'alignTop',
  dataCommand: 'align-top',
  label: 'Align top',
}, {
  command: 'alignMiddle',
  dataCommand: 'align-middle',
  label: 'Align middle',
}, {
  command: 'alignBottom',
  dataCommand: 'align-bottom',
  label: 'Align bottom',
}, {
  command: 'distributeHorizontal',
  dataCommand: 'distribute-horizontal',
  label: 'Distribute horizontal',
}, {
  command: 'distributeVertical',
  dataCommand: 'distribute-vertical',
  label: 'Distribute vertical',
}] as const satisfies readonly {
  command: PPTAlignmentPopoverCommand
  dataCommand: string
  label: string
}[]

export function PPTAlignmentPopover({
  availability,
  onCommand,
  onPreviewChange,
}: PPTAlignmentPopoverProps) {
  const [open, setOpen] = useState(false)
  const [activeCommand, setActiveCommand] =
    useState<PPTAlignmentPopoverCommand>('alignCenter')
  const [initialActiveCommandIndex, setInitialActiveCommandIndex] = useState(0)
  const triggerRef = useRef<HTMLButtonElement | null>(null)
  const commands = ALIGNMENT_COMMANDS.map((command) => ({
    ...command,
    disabled: !availability[command.command],
  }))
  const enabledCommands = commands.filter((command) => !command.disabled)
  const triggerDisabled = enabledCommands.length === 0
  const activeEnabledCommand = enabledCommands.find((command) =>
    command.command === activeCommand
  ) ?? enabledCommands[0] ?? null
  const activeEnabledCommandId = activeEnabledCommand?.command ?? null
  const {
    onFocus: handleMenuFocus,
    onKeyDown: handleMenuKeyDown,
    ref: setMenuRoot,
  } = usePPTCanvasMenuRovingFocus<HTMLDivElement>({
    initialActiveIndex: initialActiveCommandIndex,
    onClose: () => {
      closePopover()
      focusTrigger()
    },
  })

  useEffect(() => () => onPreviewChange(null), [onPreviewChange])

  function openPopover(command = activeEnabledCommand?.command) {
    if (!command) {
      return
    }

    setInitialActiveCommandIndex(Math.max(
      0,
      enabledCommands.findIndex((enabledCommand) =>
        enabledCommand.command === command
      ),
    ))
    setOpen(true)
    setActiveCommand(command)
    onPreviewChange(command)
  }

  function closePopover() {
    setOpen(false)
    onPreviewChange(null)
  }

  function focusTrigger() {
    focusPPTCanvasElementOnNextFrame({
      resolveElement: () => triggerRef.current,
    })
  }

  function handleTriggerKeyDown(event: ReactKeyboardEvent<HTMLButtonElement>) {
    const intent = getPPTCanvasMenuTriggerKeyboardIntent({ key: event.key })

    if (intent.preventDefault) {
      event.preventDefault()
    }

    if (intent.kind === 'open-menu') {
      event.stopPropagation()
      openPopover()
    }
  }

  return (
    <span
      className="ppt-alignment-popover-wrap"
      data-ppt-alignment-popover-open={open ? 'true' : 'false'}
    >
      <IconButton
        aria-controls="ppt-alignment-popover-menu"
        aria-expanded={open}
        aria-haspopup="menu"
        data-ppt-alignment-popover-trigger
        disabled={triggerDisabled}
        label="Alignment editor"
        ref={triggerRef}
        variant="floating"
        onClick={(event) => {
          event.preventDefault()
          event.stopPropagation()
          if (open) {
            closePopover()
          } else {
            openPopover()
          }
        }}
        onKeyDown={handleTriggerKeyDown}
      >
        <AlignCenterHorizontal size={16} />
      </IconButton>
      {open ? (
        <div
          aria-label="Alignment editor"
          className="ppt-alignment-popover"
          data-ppt-alignment-popover
          data-ppt-alignment-popover-active={activeEnabledCommandId ?? activeCommand}
          data-ppt-alignment-popover-model={CANVAS_DOM_ALIGNMENT_POPOVER_MODEL}
          id="ppt-alignment-popover-menu"
          ref={setMenuRoot}
          role="menu"
          onFocus={handleMenuFocus}
          onKeyDown={handleMenuKeyDown}
          onMouseLeave={() => onPreviewChange(null)}
        >
          {commands.map((command) => (
            <button
              {...PPT_MENU_ITEM_PROPS}
              aria-disabled={command.disabled}
              className="ppt-alignment-popover-item"
              data-ppt-alignment-popover-command={command.dataCommand}
              data-ppt-alignment-popover-item
              disabled={command.disabled}
              key={command.command}
              role="menuitem"
              tabIndex={command.command === activeEnabledCommand?.command ? 0 : -1}
              type="button"
              onClick={(event) => {
                event.preventDefault()
                event.stopPropagation()
                if (command.disabled) {
                  return
                }

                onCommand(command.command)
                closePopover()
                focusTrigger()
              }}
              onFocus={() => {
                setActiveCommand(command.command)
                onPreviewChange(command.command)
              }}
              onMouseEnter={() => {
                setActiveCommand(command.command)
                onPreviewChange(command.command)
              }}
            >
              <PPTAlignmentIcon command={command.command} />
              <span>{command.label}</span>
            </button>
          ))}
        </div>
      ) : null}
    </span>
  )
}

function PPTAlignmentIcon({ command }: { command: PPTAlignmentPopoverCommand }) {
  switch (command) {
    case 'alignLeft':
      return <AlignLeft size={15} />
    case 'alignCenter':
      return <AlignCenterHorizontal size={15} />
    case 'alignRight':
      return <AlignRight size={15} />
    case 'alignTop':
      return <AlignStartVertical size={15} />
    case 'alignMiddle':
      return <AlignCenterVertical size={15} />
    case 'alignBottom':
      return <AlignEndVertical size={15} />
    case 'distributeHorizontal':
      return <AlignHorizontalDistributeCenter size={15} />
    case 'distributeVertical':
      return <AlignVerticalDistributeCenter size={15} />
  }
}
