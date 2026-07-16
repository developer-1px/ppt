import {
  AlignCenterVertical,
  AlignEndVertical,
  AlignStartVertical,
  BringToFront,
  CopyPlus,
  Eraser,
  FlipHorizontal2,
  FlipVertical2,
  Grid2X2,
  Group,
  Layers,
  Lock,
  MoveDown,
  MoveUp,
  Paintbrush,
  SendToBack,
  Trash2,
  Ungroup,
  Unlock,
} from 'lucide-react'
import { PPT_MENU_ITEM_PROPS } from '../../pptCanvasAppAffordanceAdapter'
import type {
  PPTCommandSurface,
  PPTSurfaceCommand,
  PPTSurfaceCommandView,
} from './PPTSurfaceCommandModel'

export function PPTSurfaceCommandButton({
  command,
  onAfterCommand,
  onCommand,
  surface,
}: {
  command: PPTSurfaceCommandView
  onAfterCommand?: () => void
  onCommand: (command: PPTSurfaceCommand) => void
  surface: PPTCommandSurface
}) {
  const dataAttribute = surface === 'context-menu'
    ? {
        ...PPT_MENU_ITEM_PROPS,
        'data-ppt-context-command': command.dataCommand,
      }
    : { 'data-ppt-floating-command': command.dataCommand }

  return (
    <button
      aria-label={command.label}
      className={surface === 'context-menu'
        ? 'ppt-context-menu-item'
        : 'ppt-floating-command'}
      disabled={command.disabled}
      role={surface === 'context-menu' ? 'menuitem' : undefined}
      title={command.title}
      type="button"
      {...dataAttribute}
      onClick={(event) => {
        event.preventDefault()
        event.stopPropagation()

        if (command.disabled) {
          return
        }

        onCommand(command.command)
        onAfterCommand?.()
      }}
    >
      <PPTSurfaceCommandIcon command={command.command} size={16} />
      {surface === 'context-menu' ? <span>{command.label}</span> : null}
    </button>
  )
}

function PPTSurfaceCommandIcon({
  command,
  size,
}: {
  command: PPTSurfaceCommand
  size: number
}) {
  switch (command) {
    case 'alignCenter':
      return <AlignCenterVertical size={size} />
    case 'alignLeft':
      return <AlignStartVertical size={size} />
    case 'alignRight':
      return <AlignEndVertical size={size} />
    case 'bringForward':
      return <MoveUp size={size} />
    case 'bringToFront':
      return <BringToFront size={size} />
    case 'clearTextFormatting':
      return <Eraser size={size} />
    case 'copyFormatting':
      return <Paintbrush size={size} />
    case 'delete':
      return <Trash2 size={size} />
    case 'duplicate':
      return <CopyPlus size={size} />
    case 'flipHorizontal':
      return <FlipHorizontal2 size={size} />
    case 'flipVertical':
      return <FlipVertical2 size={size} />
    case 'group':
      return <Group size={size} />
    case 'lockSelection':
      return <Lock size={size} />
    case 'pasteFormatting':
      return <Paintbrush size={size} />
    case 'selectSameType':
      return <Layers size={size} />
    case 'sendBackward':
      return <MoveDown size={size} />
    case 'sendToBack':
      return <SendToBack size={size} />
    case 'tidySelection':
      return <Grid2X2 size={size} />
    case 'ungroup':
      return <Ungroup size={size} />
    case 'unlockAll':
      return <Unlock size={size} />
    case 'alignBottom':
    case 'alignMiddle':
    case 'alignTop':
    case 'distributeHorizontal':
    case 'distributeVertical':
      return null
  }
}
