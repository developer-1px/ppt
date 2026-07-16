import { PPT_COMMAND_AFFORDANCES } from '../../pptCanvasEngineAdapter'
import type { PPTCommandPaletteItemBase } from '../../pptCanvasAppAffordanceAdapter'
import { getPPTCanvasCommandAvailability } from '../../pptCommandAdapter'

export type PPTCommandPaletteItem = PPTCommandPaletteItemBase

export type PPTCommandSurface = 'context-menu' | 'selection-floating-bar'

export type PPTSurfaceCommand =
  | 'alignBottom'
  | 'alignCenter'
  | 'alignLeft'
  | 'alignMiddle'
  | 'alignRight'
  | 'alignTop'
  | 'bringForward'
  | 'bringToFront'
  | 'clearTextFormatting'
  | 'copyFormatting'
  | 'delete'
  | 'distributeHorizontal'
  | 'distributeVertical'
  | 'duplicate'
  | 'flipHorizontal'
  | 'flipVertical'
  | 'group'
  | 'lockSelection'
  | 'pasteFormatting'
  | 'selectSameType'
  | 'sendBackward'
  | 'sendToBack'
  | 'tidySelection'
  | 'ungroup'
  | 'unlockAll'

export type PPTCommandAvailability = ReturnType<
  typeof getPPTCanvasCommandAvailability
> & {
  clearTextFormatting: boolean
  copyFormatting: boolean
  flipSelection: boolean
  pasteFormatting: boolean
  selectSameType: boolean
  tidySelection: boolean
}

type PPTCommandAvailabilityKey = keyof PPTCommandAvailability

type PPTSurfaceCommandDescriptor = {
  availability: PPTCommandAvailabilityKey
  command: PPTSurfaceCommand
  dataCommand: string
  label: string
  surfaces: readonly PPTCommandSurface[]
  title: string
}

type PPTSurfaceCommandGroup = {
  commands: readonly PPTSurfaceCommandDescriptor[]
  id: string
}

export type PPTSurfaceCommandView = PPTSurfaceCommandDescriptor & {
  disabled: boolean
}

export type PPTSurfaceCommandViewGroup = {
  commands: PPTSurfaceCommandView[]
  id: string
}

const COMMAND_GROUPS: readonly PPTSurfaceCommandGroup[] = [{
  commands: [{
    availability: 'duplicate',
    command: 'duplicate',
    dataCommand: 'duplicate',
    label: 'Duplicate',
    surfaces: ['context-menu', 'selection-floating-bar'],
    title: PPT_COMMAND_AFFORDANCES.duplicate.title,
  }, {
    availability: 'copyFormatting',
    command: 'copyFormatting',
    dataCommand: 'copy-formatting',
    label: 'Copy formatting',
    surfaces: ['context-menu', 'selection-floating-bar'],
    title: 'Copy formatting',
  }, {
    availability: 'pasteFormatting',
    command: 'pasteFormatting',
    dataCommand: 'paste-formatting',
    label: 'Paste formatting',
    surfaces: ['context-menu', 'selection-floating-bar'],
    title: 'Paste formatting',
  }, {
    availability: 'clearTextFormatting',
    command: 'clearTextFormatting',
    dataCommand: 'clear-text-formatting',
    label: 'Clear formatting',
    surfaces: ['context-menu', 'selection-floating-bar'],
    title: 'Clear formatting',
  }, {
    availability: 'selectSameType',
    command: 'selectSameType',
    dataCommand: 'select-same-type',
    label: 'Select same type',
    surfaces: ['context-menu', 'selection-floating-bar'],
    title: 'Select same type',
  }, {
    availability: 'tidySelection',
    command: 'tidySelection',
    dataCommand: 'tidy-selection',
    label: 'Tidy selection',
    surfaces: ['context-menu', 'selection-floating-bar'],
    title: 'Tidy selection',
  }, {
    availability: 'flipSelection',
    command: 'flipHorizontal',
    dataCommand: 'flip-horizontal',
    label: 'Flip horizontal',
    surfaces: ['context-menu'],
    title: 'Flip horizontal',
  }, {
    availability: 'flipSelection',
    command: 'flipVertical',
    dataCommand: 'flip-vertical',
    label: 'Flip vertical',
    surfaces: ['context-menu'],
    title: 'Flip vertical',
  }, {
    availability: 'delete',
    command: 'delete',
    dataCommand: 'delete',
    label: 'Delete',
    surfaces: ['context-menu', 'selection-floating-bar'],
    title: PPT_COMMAND_AFFORDANCES.delete.title,
  }],
  id: 'edit',
}, {
  commands: [{
    availability: 'alignLeft',
    command: 'alignLeft',
    dataCommand: 'align-left',
    label: 'Align left',
    surfaces: ['context-menu'],
    title: PPT_COMMAND_AFFORDANCES.alignLeft.title,
  }, {
    availability: 'alignCenter',
    command: 'alignCenter',
    dataCommand: 'align-center-x',
    label: 'Align center',
    surfaces: ['context-menu'],
    title: PPT_COMMAND_AFFORDANCES.alignCenter.title,
  }, {
    availability: 'alignRight',
    command: 'alignRight',
    dataCommand: 'align-right',
    label: 'Align right',
    surfaces: ['context-menu'],
    title: PPT_COMMAND_AFFORDANCES.alignRight.title,
  }],
  id: 'align',
}, {
  commands: [{
    availability: 'bringForward',
    command: 'bringForward',
    dataCommand: 'bring-forward',
    label: 'Bring forward',
    surfaces: ['context-menu'],
    title: PPT_COMMAND_AFFORDANCES.bringForward.title,
  }, {
    availability: 'bringToFront',
    command: 'bringToFront',
    dataCommand: 'bring-to-front',
    label: 'Bring to front',
    surfaces: ['context-menu', 'selection-floating-bar'],
    title: PPT_COMMAND_AFFORDANCES.bringToFront.title,
  }, {
    availability: 'sendBackward',
    command: 'sendBackward',
    dataCommand: 'send-backward',
    label: 'Send backward',
    surfaces: ['context-menu'],
    title: PPT_COMMAND_AFFORDANCES.sendBackward.title,
  }, {
    availability: 'sendToBack',
    command: 'sendToBack',
    dataCommand: 'send-to-back',
    label: 'Send to back',
    surfaces: ['context-menu', 'selection-floating-bar'],
    title: PPT_COMMAND_AFFORDANCES.sendToBack.title,
  }],
  id: 'order',
}, {
  commands: [{
    availability: 'group',
    command: 'group',
    dataCommand: 'group',
    label: 'Group',
    surfaces: ['context-menu', 'selection-floating-bar'],
    title: PPT_COMMAND_AFFORDANCES.group.title,
  }, {
    availability: 'ungroup',
    command: 'ungroup',
    dataCommand: 'ungroup',
    label: 'Ungroup',
    surfaces: ['context-menu', 'selection-floating-bar'],
    title: PPT_COMMAND_AFFORDANCES.ungroup.title,
  }],
  id: 'group',
}, {
  commands: [{
    availability: 'lockSelection',
    command: 'lockSelection',
    dataCommand: 'lock-selection',
    label: 'Lock',
    surfaces: ['context-menu', 'selection-floating-bar'],
    title: PPT_COMMAND_AFFORDANCES.lockSelection.title,
  }, {
    availability: 'unlockAll',
    command: 'unlockAll',
    dataCommand: 'unlock-all',
    label: 'Unlock all',
    surfaces: ['context-menu'],
    title: PPT_COMMAND_AFFORDANCES.unlockAll.title,
  }],
  id: 'lock',
}]

export function getPPTCommandSurfaceGroups({
  availability,
  surface,
}: {
  availability: PPTCommandAvailability
  surface: PPTCommandSurface
}): PPTSurfaceCommandViewGroup[] {
  return COMMAND_GROUPS.flatMap((group) => {
    const commands = group.commands
      .filter((command) => command.surfaces.includes(surface))
      .map((command) => ({
        ...command,
        disabled: !availability[command.availability],
      }))

    return commands.length > 0
      ? [{
          commands,
          id: group.id,
        }]
      : []
  })
}
