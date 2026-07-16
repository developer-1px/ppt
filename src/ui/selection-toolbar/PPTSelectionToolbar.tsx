import { MoreHorizontal } from 'lucide-react'
import { Fragment, useState, type CSSProperties } from 'react'
import './selection-toolbar.css'
import type { PPTCanvasFloatingAnchor } from '../../pptCanvasAppAffordanceAdapter'
import type { PPTShapeKind } from '../../pptModel'
import {
  PPTSurfaceCommandButton,
  type PPTSurfaceCommand,
  type PPTSurfaceCommandViewGroup,
} from '../command-surface'
import { IconButton } from '../core'
import {
  PPTAlignmentPopover,
  type PPTAlignmentPopoverCommand,
  type PPTAlignmentPopoverProps,
} from './PPTAlignmentPopover'
import {
  PPTShapeKindMenu,
  type PPTShapeQuickMenuState,
} from './PPTShapeKindMenu'
import {
  PPTTextQuickFormatControls,
  type PPTTextQuickFormatAction,
  type PPTTextQuickFormatState,
} from './PPTTextQuickFormatControls'

export type PPTSelectionToolbarModel = {
  alignmentAvailability: PPTAlignmentPopoverProps['availability']
  anchor: PPTCanvasFloatingAnchor | null
  commandGroups: readonly PPTSurfaceCommandViewGroup[]
  scale: number
  shape: PPTShapeQuickMenuState | null
  text: PPTTextQuickFormatState | null
}

export type PPTSelectionToolbarAction =
  | {
      command: PPTAlignmentPopoverCommand | null
      type: 'alignment-preview-change'
    }
  | {
      command: PPTSurfaceCommand
      type: 'run-command'
    }
  | {
      elementId: string
      shape: PPTShapeKind
      type: 'shape-kind-change'
    }
  | {
      action: PPTTextQuickFormatAction
      type: 'text-format'
    }

export type PPTSelectionToolbarProps = {
  model: PPTSelectionToolbarModel
  onAction: (action: PPTSelectionToolbarAction) => void
}

export function PPTSelectionToolbar({
  model,
  onAction,
}: PPTSelectionToolbarProps) {
  const [expanded, setExpanded] = useState(false)
  const {
    alignmentAvailability,
    anchor,
    commandGroups,
    scale,
    shape,
    text,
  } = model

  if (!anchor || (commandGroups.length === 0 && !shape && !text)) {
    return null
  }

  return (
    <div
      aria-label="Selection actions"
      className="ppt-selection-floating-bar"
      data-placement={anchor.placement}
      data-ppt-selection-floating-bar
      data-ppt-selection-floating-expanded={expanded ? 'true' : 'false'}
      role="toolbar"
      style={{
        '--ppt-command-scale': String(1 / scale),
        left: anchor.x,
        top: anchor.y,
      } as CSSProperties}
      onPointerDown={(event) => event.stopPropagation()}
    >
      {text ? (
        <PPTTextQuickFormatControls
          expanded={expanded}
          state={text}
          onAction={(action) => onAction({
            action,
            type: 'text-format',
          })}
        />
      ) : null}
      {text ? <span className="ppt-command-divider" /> : null}
      {shape ? (
        <>
          <PPTShapeKindMenu
            state={shape}
            onShapeKindChange={(elementId, nextShape) => onAction({
              elementId,
              shape: nextShape,
              type: 'shape-kind-change',
            })}
          />
          <span className="ppt-command-divider" />
        </>
      ) : null}
      <span className="ppt-selection-floating-details" hidden={!expanded}>
        <PPTAlignmentPopover
          availability={alignmentAvailability}
          onCommand={(command) => onAction({
            command,
            type: 'run-command',
          })}
          onPreviewChange={(command) => onAction({
            command,
            type: 'alignment-preview-change',
          })}
        />
        {commandGroups.length > 0
          ? <span className="ppt-command-divider" />
          : null}
        {commandGroups.map((group, groupIndex) => (
          <Fragment key={group.id}>
            {groupIndex > 0 ? <span className="ppt-command-divider" /> : null}
            {group.commands.map((command) => (
              <PPTSurfaceCommandButton
                command={command}
                key={command.command}
                surface="selection-floating-bar"
                onCommand={(nextCommand) => onAction({
                  command: nextCommand,
                  type: 'run-command',
                })}
              />
            ))}
          </Fragment>
        ))}
      </span>
      <IconButton
        aria-pressed={expanded}
        data-ppt-selection-floating-more
        label={expanded ? 'Fewer selection actions' : 'More selection actions'}
        tooltip={expanded ? 'Fewer actions' : 'More actions'}
        variant="floating"
        onClick={(event) => {
          event.preventDefault()
          event.stopPropagation()
          setExpanded((current) => !current)
        }}
      >
        <MoreHorizontal size={16} />
      </IconButton>
    </div>
  )
}
