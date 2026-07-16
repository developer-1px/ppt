import {
  ChevronDown,
  ChevronRight,
  Eye,
  EyeOff,
  Group,
  Layers,
  Lock,
  Unlock,
} from 'lucide-react'
import {
  useMemo,
  useRef,
  useState,
  type DragEvent as ReactDragEvent,
  type KeyboardEvent as ReactKeyboardEvent,
  type MouseEvent as ReactMouseEvent,
} from 'react'
import {
  getSlideEditLayerPaneCommandEffect,
  getSlideEditLayerPaneDropIndicator,
  getSlideEditLayerPaneKeyboardIntent,
  getSlideEditLayerPaneResolvedFocusObjectId,
  getSlideEditObjectVisibilityCommandEffect,
  SLIDE_EDIT_LAYER_PANE_COMMANDS,
  SLIDE_EDIT_LAYER_PANE_DROP_INDICATOR_MODEL,
  SLIDE_EDIT_LAYER_PANE_KEYBOARD_INTENT_MODEL,
  SLIDE_EDIT_LAYER_PANE_KEYBOARD_KEYS,
} from '../../pptSlideEditAffordanceAdapter'
import {
  focusPPTCanvasElementBySelectorOnNextFrame,
  getPPTCanvasDataTransferText,
  getPPTCanvasEditableFieldKeyboardIntent,
  getPPTCanvasPointerLocalGeometry,
  getPPTCanvasSelectionListModifierState,
  setPPTCanvasDataTransferDropEffect,
  setPPTCanvasDataTransferText,
} from '../../pptCanvasAppAffordanceAdapter'
import {
  createPPTLayerPaneDescriptor,
  getPPTLayerPaneDefaultFocusObjectId,
  getPPTLayerPaneDropIndex,
  getPPTLayerPaneGroupIdFromRowId,
  getPPTLayerPaneKeyboardDropIndex,
  getPPTObjectVisibilityAvailability,
  getPPTObjectVisibilityCommandId,
  getPPTObjectVisibilityDescriptors,
  getPPTObjectVisibilityState,
  type PPTLayerPaneDropPlacement,
  type PPTLayerPaneIntent,
  type PPTLayerPaneKeyboardIntent,
  type PPTLayerPaneRowDescriptor,
} from '../../pptLayerPaneAdapter'
import { createPPTInspectorActionDispatcher } from './PPTInspectorActionDispatcher'
import type { PPTInspectorProps } from './PPTInspectorContract'

type PPTLayerPaneProps = Pick<PPTInspectorProps, 'model' | 'onAction'>
type PPTLayerPaneGroupState = {
  collapsedGroupIds: readonly string[]
  focusedObjectId: string | null
  rangeAnchorObjectId: string | null
}
type PPTLayerPaneRenameState = {
  objectId: string
  value: string
}
type PPTLayerPaneDragState = {
  dropPlacement?: PPTLayerPaneDropPlacement
  dropTargetObjectId?: string
  dropToIndex?: number
  objectId: string
}

const PPT_LAYER_PANE_COMMANDS = SLIDE_EDIT_LAYER_PANE_COMMANDS

export function PPTLayerPane({ model, onAction }: PPTLayerPaneProps) {
  const { selection, slide } = model
  const {
    onLayerPaneCommandEffect,
    onObjectVisibilityCommandEffect,
  } = createPPTInspectorActionDispatcher(onAction)
  const [layerPaneGroupState, setLayerPaneGroupState] =
    useState<PPTLayerPaneGroupState>({
      collapsedGroupIds: [],
      focusedObjectId: null,
      rangeAnchorObjectId: null,
    })
  const [layerPaneRenameState, setLayerPaneRenameState] =
    useState<PPTLayerPaneRenameState | null>(null)
  const [layerPaneDragState, setLayerPaneDragState] =
    useState<PPTLayerPaneDragState | null>(null)
  const layerPaneRenameCommitSuppressedRef = useRef(false)
  const collapsedLayerPaneGroupIdSet = useMemo(
    () => new Set(layerPaneGroupState.collapsedGroupIds),
    [layerPaneGroupState.collapsedGroupIds],
  )
  const defaultLayerPaneFocusObjectId = getPPTLayerPaneDefaultFocusObjectId(
    slide,
    selection,
  )
  const layerPaneDescriptor = createPPTLayerPaneDescriptor({
    activeObjectId: layerPaneGroupState.focusedObjectId ??
      defaultLayerPaneFocusObjectId,
    collapsedGroupIds: collapsedLayerPaneGroupIdSet,
    selectedObjectIds: selection,
    slide,
  })
  const objectVisibilityDescriptors = getPPTObjectVisibilityDescriptors(
    slide.id,
    layerPaneDescriptor.rows,
  )
  const activeLayerPaneObjectId = getSlideEditLayerPaneResolvedFocusObjectId(
    layerPaneDescriptor,
    {
      defaultObjectId: defaultLayerPaneFocusObjectId,
      preferredObjectId: layerPaneGroupState.focusedObjectId,
    },
  )
  const layerPaneCommandIds = PPT_LAYER_PANE_COMMANDS.map((command) => command.id).join(' ')
  function runLayerPaneIntent(intent: PPTLayerPaneIntent) {
    if (intent.type === 'visibility-toggle') {
      const row = layerPaneDescriptor.rows.find((candidate) =>
        candidate.objectId === intent.objectId
      )

      if (!row) {
        return
      }

      const effect = getSlideEditObjectVisibilityCommandEffect({
        commandId: getPPTObjectVisibilityCommandId(row),
        objects: objectVisibilityDescriptors,
        selectedObjectIds: [row.objectId],
        slideId: slide.id,
      })

      if (effect) {
        onObjectVisibilityCommandEffect(effect)
      }

      return
    }

    const effect = getSlideEditLayerPaneCommandEffect(layerPaneDescriptor, intent)

    if (effect) {
      onLayerPaneCommandEffect(effect)
    }
  }

  function focusLayerPaneRow(objectId: string) {
    focusPPTCanvasElementBySelectorOnNextFrame<HTMLElement>({
      match: ({ element }) =>
        element.getAttribute('data-ppt-layer-pane-row') === objectId,
      root: document,
      selector: '[data-ppt-layer-pane-row]',
    })
  }

  function setLayerPaneFocusedObjectId(
    objectId: string,
    options?: { rangeAnchor?: boolean },
  ) {
    setLayerPaneGroupState((current) => ({
      ...current,
      focusedObjectId: objectId,
      rangeAnchorObjectId: options?.rangeAnchor
        ? objectId
        : current.rangeAnchorObjectId,
    }))
  }

  function setLayerPaneGroupExpanded(objectId: string, isExpanded: boolean) {
    const groupId = getPPTLayerPaneGroupIdFromRowId(objectId)

    if (!groupId) {
      return
    }

    setLayerPaneGroupState((current) => {
      const collapsedGroupIds = new Set(current.collapsedGroupIds)

      if (isExpanded) {
        collapsedGroupIds.delete(groupId)
      } else {
        collapsedGroupIds.add(groupId)
      }

      return {
        ...current,
        collapsedGroupIds: [...collapsedGroupIds],
        focusedObjectId: objectId,
      }
    })
  }

  function focusLayerPaneRenameInput(objectId: string) {
    focusPPTCanvasElementBySelectorOnNextFrame<HTMLInputElement>({
      match: ({ element }) =>
        element.getAttribute('data-ppt-layer-pane-rename-input') === objectId,
      root: document,
      select: true,
      selector: '[data-ppt-layer-pane-rename-input]',
    })
  }

  function startLayerPaneRename(row: PPTLayerPaneRowDescriptor) {
    if (!row.isRenamable) {
      return
    }

    layerPaneRenameCommitSuppressedRef.current = false
    setLayerPaneFocusedObjectId(row.objectId)
    setLayerPaneRenameState({
      objectId: row.objectId,
      value: row.displayName,
    })
    focusLayerPaneRenameInput(row.objectId)
  }

  function commitLayerPaneRename(objectId: string, value: string) {
    layerPaneRenameCommitSuppressedRef.current = true
    setLayerPaneRenameState(null)
    runLayerPaneIntent({
      name: value,
      objectId,
      type: 'rename-submit',
    })
    focusLayerPaneRow(objectId)
  }

  function cancelLayerPaneRename(objectId: string) {
    layerPaneRenameCommitSuppressedRef.current = true
    setLayerPaneRenameState(null)
    focusLayerPaneRow(objectId)
  }

  function handleLayerPaneRenameBlur(objectId: string, value: string) {
    if (layerPaneRenameCommitSuppressedRef.current) {
      layerPaneRenameCommitSuppressedRef.current = false
      return
    }

    commitLayerPaneRename(objectId, value)
  }

  function canDragLayerPaneRow(row: PPTLayerPaneRowDescriptor) {
    return row.isReorderable
  }

  function handleLayerPaneRowPress(
    row: PPTLayerPaneRowDescriptor,
    event: ReactMouseEvent<HTMLButtonElement>,
  ) {
    const rangeAnchorObjectId = event.shiftKey
      ? layerPaneGroupState.rangeAnchorObjectId ?? activeLayerPaneObjectId
      : null
    const selectionModifierState = getPPTCanvasSelectionListModifierState({
      ctrlKey: event.ctrlKey,
      hasRangeAnchor: Boolean(rangeAnchorObjectId),
      metaKey: event.metaKey,
      shiftKey: event.shiftKey,
    })

    setLayerPaneFocusedObjectId(row.objectId, {
      rangeAnchor: !selectionModifierState.range,
    })
    runLayerPaneIntent({
      ...(selectionModifierState.range && rangeAnchorObjectId
        ? { rangeAnchorObjectId }
        : { additive: selectionModifierState.additive }),
      objectId: row.objectId,
      type: 'row-press',
    })
  }

  function handleLayerPaneRowDragStart(
    row: PPTLayerPaneRowDescriptor,
    event: ReactDragEvent<HTMLElement>,
  ) {
    if (!canDragLayerPaneRow(row)) {
      event.preventDefault()
      return
    }

    setLayerPaneFocusedObjectId(row.objectId)
    setLayerPaneDragState({ objectId: row.objectId })
    setPPTCanvasDataTransferText({
      dataTransfer: event.dataTransfer,
      effectAllowed: 'move',
      text: row.objectId,
    })
  }

  function handleLayerPaneRowDragOver(
    row: PPTLayerPaneRowDescriptor,
    event: ReactDragEvent<HTMLElement>,
  ) {
    const draggedObjectId = layerPaneDragState?.objectId ||
      getPPTCanvasDataTransferText({ dataTransfer: event.dataTransfer })

    if (!draggedObjectId) {
      return
    }

    const rowGeometry = getPPTCanvasPointerLocalGeometry({
      event,
      target: event.currentTarget,
    })

    if (!rowGeometry) {
      setLayerPaneDragState({ objectId: draggedObjectId })
      return
    }

    const dropIndicator = getSlideEditLayerPaneDropIndicator(
      layerPaneDescriptor,
      {
        draggedObjectId,
        pointerOffsetY: rowGeometry.point.y,
        rowHeight: rowGeometry.rect.height,
        targetObjectId: row.objectId,
      },
    )

    if (dropIndicator.placement === 'none' || dropIndicator.targetObjectId === null) {
      setLayerPaneDragState({ objectId: draggedObjectId })
      return
    }

    const dropIndex = getPPTLayerPaneDropIndex(
      slide,
      dropIndicator.targetObjectId,
      dropIndicator.placement,
    )

    if (dropIndex === null) {
      setLayerPaneDragState({ objectId: draggedObjectId })
      return
    }

    event.preventDefault()
    setPPTCanvasDataTransferDropEffect({
      dataTransfer: event.dataTransfer,
      dropEffect: 'move',
    })
    setLayerPaneDragState({
      dropPlacement: dropIndicator.placement,
      dropTargetObjectId: dropIndicator.targetObjectId,
      dropToIndex: dropIndex,
      objectId: draggedObjectId,
    })
  }

  function handleLayerPaneRowDrop(
    row: PPTLayerPaneRowDescriptor,
    event: ReactDragEvent<HTMLElement>,
  ) {
    const draggedObjectId = layerPaneDragState?.objectId ||
      getPPTCanvasDataTransferText({ dataTransfer: event.dataTransfer })

    if (!draggedObjectId || draggedObjectId === row.objectId) {
      return
    }

    const activeDropIndex = layerPaneDragState?.dropTargetObjectId === row.objectId
      ? layerPaneDragState.dropToIndex
      : undefined

    const rowGeometry = getPPTCanvasPointerLocalGeometry({
      event,
      target: event.currentTarget,
    })

    if (!rowGeometry) {
      return
    }

    const dropIndicator = getSlideEditLayerPaneDropIndicator(
      layerPaneDescriptor,
      {
        draggedObjectId,
        pointerOffsetY: rowGeometry.point.y,
        rowHeight: rowGeometry.rect.height,
        targetObjectId: row.objectId,
      },
    )

    if (
      activeDropIndex === undefined &&
      (dropIndicator.placement === 'none' || dropIndicator.targetObjectId === null)
    ) {
      return
    }

    const dropIndex = activeDropIndex ?? getPPTLayerPaneDropIndex(
      slide,
      dropIndicator.targetObjectId ?? row.objectId,
      dropIndicator.placement === 'none' ? 'before' : dropIndicator.placement,
    )

    if (dropIndex === null) {
      return
    }

    event.preventDefault()
    event.stopPropagation()
    setLayerPaneDragState(null)
    setLayerPaneFocusedObjectId(draggedObjectId)
    runLayerPaneIntent({
      objectId: draggedObjectId,
      toIndex: dropIndex,
      type: 'row-drop',
    })
    focusLayerPaneRow(draggedObjectId)
  }

  function handleLayerPaneRowDragEnd() {
    setLayerPaneDragState(null)
  }

  function applyLayerPaneKeyboardIntent(intent: PPTLayerPaneKeyboardIntent) {
    switch (intent.type) {
      case 'focus-row':
      case 'focus-parent-row':
      case 'select-row':
        setLayerPaneFocusedObjectId(intent.objectId, {
          rangeAnchor: true,
        })
        runLayerPaneIntent({
          objectId: intent.objectId,
          type: 'row-press',
        })
        focusLayerPaneRow(intent.objectId)
        return
      case 'range-select-row':
        setLayerPaneGroupState((current) => ({
          ...current,
          focusedObjectId: intent.objectId,
          rangeAnchorObjectId: intent.rangeAnchorObjectId,
        }))
        runLayerPaneIntent({
          objectId: intent.objectId,
          rangeAnchorObjectId: intent.rangeAnchorObjectId,
          type: 'row-press',
        })
        focusLayerPaneRow(intent.objectId)
        return
      case 'reorder-row':
        {
          const dropIndex = getPPTLayerPaneKeyboardDropIndex(
            layerPaneDescriptor,
            slide,
            intent.objectId,
            intent.toIndex,
          )

          if (dropIndex === null) {
            return
          }

          setLayerPaneFocusedObjectId(intent.objectId)
          runLayerPaneIntent({
            objectId: intent.objectId,
            toIndex: dropIndex,
            type: 'row-drop',
          })
          focusLayerPaneRow(intent.objectId)
        }
        return
      case 'collapse-row':
        setLayerPaneGroupExpanded(intent.objectId, false)
        focusLayerPaneRow(intent.objectId)
        return
      case 'expand-row':
        setLayerPaneGroupExpanded(intent.objectId, true)
        focusLayerPaneRow(intent.objectId)
        return
      case 'rename-row':
        {
          const row = layerPaneDescriptor.rows.find((candidate) =>
            candidate.objectId === intent.objectId
          )

          if (row) {
            startLayerPaneRename(row)
          }
        }
        return
      case 'none':
        return
    }
  }

  function handleLayerPaneRowKeyDown(
    row: PPTLayerPaneRowDescriptor,
    event: ReactKeyboardEvent<HTMLElement>,
  ) {
    if (
      event.target !== event.currentTarget ||
      event.ctrlKey ||
      event.metaKey ||
      (event.altKey && event.key !== 'ArrowDown' && event.key !== 'ArrowUp')
    ) {
      return
    }

    const intent = getSlideEditLayerPaneKeyboardIntent(layerPaneDescriptor, {
      currentObjectId: row.objectId,
      key: event.key,
      altKey: event.altKey,
      rangeAnchorObjectId: layerPaneGroupState.rangeAnchorObjectId,
      shiftKey: event.shiftKey,
    })

    if (!intent.preventDefault) {
      return
    }

    event.preventDefault()
    event.stopPropagation()
    applyLayerPaneKeyboardIntent(intent)
  }

  return (
    <>
      <div className="ppt-panel-header ppt-inspector-advanced-only">
        <h2>Objects</h2>
      </div>
      <section
        className="ppt-panel-section"
        data-ppt-layer-pane
        data-ppt-layer-pane-active-object-id={activeLayerPaneObjectId ?? ''}
        data-ppt-layer-pane-command-count={PPT_LAYER_PANE_COMMANDS.length}
        data-ppt-layer-pane-commands={layerPaneCommandIds}
        data-ppt-layer-pane-command-slot="command-effect"
        data-ppt-layer-pane-model="slide-edit-object-layer-pane"
        data-ppt-layer-pane-row-count={layerPaneDescriptor.rows.length}
        data-ppt-layer-pane-slide-id={layerPaneDescriptor.slideId}
      >
        <div
          aria-label="Selection pane"
          className="ppt-layer-list"
          data-ppt-layer-pane-aria-container={layerPaneDescriptor.aria.containerRole}
          data-ppt-layer-pane-aria-row={layerPaneDescriptor.aria.rowRole}
          data-ppt-layer-pane-keyboard-intent-model={SLIDE_EDIT_LAYER_PANE_KEYBOARD_INTENT_MODEL}
          data-ppt-layer-pane-keyboard-keys={SLIDE_EDIT_LAYER_PANE_KEYBOARD_KEYS}
          data-ppt-layer-pane-keyboard-model={layerPaneDescriptor.aria.keyboardModel}
          data-ppt-layer-pane-range-anchor-object-id={layerPaneGroupState.rangeAnchorObjectId ?? ''}
          data-ppt-layer-pane-range-selection-model="row-press-range-anchor"
          data-ppt-layer-pane-selection-model={layerPaneDescriptor.aria.selectionModel}
          role={layerPaneDescriptor.aria.containerRole}
        >
          {layerPaneDescriptor.rows.map((row) => {
            const activeLayerPaneRename =
              layerPaneRenameState?.objectId === row.objectId
                ? layerPaneRenameState
                : null
            const isRenaming = activeLayerPaneRename !== null
            const renameValue = activeLayerPaneRename?.value ?? ''
            const isDraggable = canDragLayerPaneRow(row)
            const layerPaneDropPlacement =
              layerPaneDragState?.dropTargetObjectId === row.objectId
                ? layerPaneDragState.dropPlacement
                : undefined
            const layerPaneDropToIndex =
              layerPaneDragState?.dropTargetObjectId === row.objectId
                ? layerPaneDragState.dropToIndex
                : undefined
            const objectVisibilityState = getPPTObjectVisibilityState(row)
            const stageObjectVisibilityState = getPPTObjectVisibilityState(
              row,
              'visible-only',
            )
            const objectVisibilityAvailability =
              getPPTObjectVisibilityAvailability({
                descriptors: objectVisibilityDescriptors,
                row,
              })

            return (
              <div
                aria-disabled={!row.isSelectable}
                aria-expanded={row.ariaExpanded}
                aria-level={row.ariaLevel}
                aria-posinset={row.ariaPosInSet}
                aria-selected={row.isSelected}
                aria-setsize={row.ariaSetSize}
                className="ppt-layer-row"
                data-grouped={row.isGrouped ? 'true' : 'false'}
                data-hidden={row.isHidden ? 'true' : 'false'}
                data-locked={row.isLocked ? 'true' : 'false'}
                data-ppt-layer-pane-expanded={
                  row.ariaExpanded === undefined
                    ? ''
                    : String(row.ariaExpanded)
                }
                data-ppt-layer-pane-grouped={row.isGrouped ? 'true' : 'false'}
                data-ppt-layer-pane-hidden={row.isHidden ? 'true' : 'false'}
                data-ppt-layer-pane-draggable={isDraggable ? 'true' : 'false'}
                data-ppt-layer-pane-dragging={
                  layerPaneDragState?.objectId === row.objectId ? 'true' : 'false'
                }
                data-ppt-layer-pane-drop-indicator={
                  layerPaneDropPlacement ?? ''
                }
                data-ppt-layer-pane-drop-indicator-model={
                  SLIDE_EDIT_LAYER_PANE_DROP_INDICATOR_MODEL
                }
                data-ppt-layer-pane-drop-target={
                  layerPaneDropPlacement ? 'true' : 'false'
                }
                data-ppt-layer-pane-drop-to-index={layerPaneDropToIndex ?? ''}
                data-ppt-layer-pane-is-group={row.isGroup ? 'true' : 'false'}
                data-ppt-layer-pane-kind={row.kindLabel}
                data-ppt-layer-pane-locked={row.isLocked ? 'true' : 'false'}
                data-ppt-layer-pane-order={row.order}
                data-ppt-layer-pane-parent-object-id={row.parentObjectId ?? ''}
                data-ppt-layer-pane-renamable={
                  row.isRenamable ? 'true' : 'false'
                }
                data-ppt-layer-pane-reorderable={
                  row.isReorderable ? 'true' : 'false'
                }
                data-ppt-layer-pane-row={row.objectId}
                data-ppt-layer-pane-row-type={
                  row.isGroup ? 'group' : 'object'
                }
                data-ppt-layer-pane-selected={row.isSelected ? 'true' : 'false'}
                data-ppt-object-visibility-hidden={
                  objectVisibilityState.isHidden ? 'true' : 'false'
                }
                data-ppt-object-visibility-layer-selection-block-reason={
                  objectVisibilityState.selectionBlockReason ?? ''
                }
                data-ppt-object-visibility-model="slide-edit-object-visibility"
                data-ppt-object-visibility-selectable={
                  objectVisibilityState.isSelectable ? 'true' : 'false'
                }
                data-ppt-object-visibility-selection-policy="allow-hidden-selection"
                data-ppt-object-visibility-stage-selection-block-reason={
                  stageObjectVisibilityState.selectionBlockReason ?? ''
                }
                data-ppt-object-visibility-visible={
                  objectVisibilityState.isVisible ? 'true' : 'false'
                }
                data-ppt-layer-row={row.objectId}
                draggable={isDraggable}
                key={row.objectId}
                role={layerPaneDescriptor.aria.rowRole}
                tabIndex={row.objectId === activeLayerPaneObjectId ? 0 : -1}
                onDragEnd={handleLayerPaneRowDragEnd}
                onDragOver={(event) => handleLayerPaneRowDragOver(row, event)}
                onDragStart={(event) => handleLayerPaneRowDragStart(row, event)}
                onDrop={(event) => handleLayerPaneRowDrop(row, event)}
                onKeyDown={(event) => handleLayerPaneRowKeyDown(row, event)}
              >
                {row.isGroup ? (
                  <button
                    aria-label={
                      row.ariaExpanded ? 'Collapse group' : 'Expand group'
                    }
                    className="ppt-layer-disclosure"
                    data-ppt-layer-pane-disclosure={row.objectId}
                    data-ppt-layer-pane-intent={
                      row.ariaExpanded ? 'collapse-row' : 'expand-row'
                    }
                    style={{
                      paddingLeft: Math.max(0, row.ariaLevel - 1) * 12,
                    }}
                    title={
                      row.ariaExpanded ? 'Collapse group' : 'Expand group'
                    }
                    type="button"
                    onClick={(event) => {
                      event.stopPropagation()
                      setLayerPaneFocusedObjectId(row.objectId)
                      setLayerPaneGroupExpanded(
                        row.objectId,
                        row.ariaExpanded !== true,
                      )
                      focusLayerPaneRow(row.objectId)
                    }}
                  >
                    {row.ariaExpanded ? (
                      <ChevronDown size={14} />
                    ) : (
                      <ChevronRight size={14} />
                    )}
                  </button>
                ) : (
                  <span
                    className="ppt-layer-disclosure ppt-layer-disclosure--spacer"
                    style={{
                      paddingLeft: Math.max(0, row.ariaLevel - 1) * 12,
                    }}
                  />
                )}
                {isRenaming ? (
                  <div className="ppt-layer-select ppt-layer-select--editing">
                    <span className="ppt-layer-kind">
                      {row.isGroup ? <Group size={14} /> : <Layers size={14} />}
                    </span>
                    <input
                      aria-label="Rename object"
                      className="ppt-layer-rename-input"
                      data-ppt-layer-pane-rename-input={row.objectId}
                      value={renameValue}
                      onBlur={() =>
                        handleLayerPaneRenameBlur(row.objectId, renameValue)}
                      onChange={(event) =>
                        setLayerPaneRenameState({
                          objectId: row.objectId,
                          value: event.target.value,
                        })}
                      onClick={(event) => event.stopPropagation()}
                      onDoubleClick={(event) => event.stopPropagation()}
                      onKeyDown={(event) => {
                        event.stopPropagation()
                        const intent = getPPTCanvasEditableFieldKeyboardIntent({
                          key: event.key,
                        })

                        if (intent.preventDefault) {
                          event.preventDefault()
                        }

                        if (intent.kind === 'commit') {
                          commitLayerPaneRename(row.objectId, renameValue)
                          return
                        }

                        if (intent.kind === 'cancel') {
                          cancelLayerPaneRename(row.objectId)
                        }
                      }}
                    />
                  </div>
                ) : (
                  <button
                    className="ppt-layer-select"
                    data-ppt-layer-pane-intent="row-press"
                    data-ppt-layer-select={row.objectId}
                    type="button"
                    onDoubleClick={(event) => {
                      event.stopPropagation()
                      startLayerPaneRename(row)
                    }}
                    onClick={(event) => {
                      handleLayerPaneRowPress(row, event)
                    }}
                  >
                    <span className="ppt-layer-kind">
                      {row.isGroup ? <Group size={14} /> : <Layers size={14} />}
                    </span>
                    <span className="ppt-layer-name">{row.displayName}</span>
                  </button>
                )}
                <div className="ppt-layer-actions">
                  <button
                    aria-label={row.isHidden ? 'Show object' : 'Hide object'}
                    className="ppt-layer-icon-button"
                    data-ppt-layer-pane-intent="visibility-toggle"
                    data-ppt-layer-visibility={row.objectId}
                    data-ppt-object-visibility-availability={
                      objectVisibilityAvailability.isAvailable
                        ? 'true'
                        : 'false'
                    }
                    data-ppt-object-visibility-command={
                      objectVisibilityAvailability.commandId
                    }
                    data-ppt-object-visibility-targets={
                      objectVisibilityAvailability.targetObjectIds.join(' ')
                    }
                    data-ppt-object-visibility-unavailable={
                      objectVisibilityAvailability.unavailableReason ?? ''
                    }
                    title={row.isHidden ? 'Show object' : 'Hide object'}
                    type="button"
                    onClick={(event) => {
                      event.stopPropagation()
                      runLayerPaneIntent({
                        objectId: row.objectId,
                        type: 'visibility-toggle',
                      })
                    }}
                  >
                    {row.isHidden ? <EyeOff size={14} /> : <Eye size={14} />}
                  </button>
                  <button
                    aria-label={row.isLocked ? 'Unlock object' : 'Lock object'}
                    className="ppt-layer-icon-button"
                    data-ppt-layer-lock={row.objectId}
                    data-ppt-layer-pane-intent="lock-toggle"
                    title={row.isLocked ? 'Unlock object' : 'Lock object'}
                    type="button"
                    onClick={(event) => {
                      event.stopPropagation()
                      runLayerPaneIntent({
                        objectId: row.objectId,
                        type: 'lock-toggle',
                      })
                    }}
                  >
                    {row.isLocked ? <Lock size={14} /> : <Unlock size={14} />}
                  </button>
                </div>
              </div>
            )
          })}
        </div>
      </section>
    </>
  )
}
