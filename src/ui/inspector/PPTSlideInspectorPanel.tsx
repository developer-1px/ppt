import { Eye, EyeOff } from 'lucide-react'
import {
  createSlideEditTransitionDescriptor,
  SLIDE_EDIT_TRANSITION_TIMING_LIMITS,
  SLIDE_EDIT_TRANSITION_TYPES,
  type SlideEditSlideMetadataFieldDescriptor,
  type SlideEditSlideMetadataFieldId,
} from '../../pptSlideEditAffordanceAdapter'
import type { PPTSlideTransition } from '../../pptModel'
import { createPPTInspectorActionDispatcher } from './PPTInspectorActionDispatcher'
import type {
  PPTInspectorModel,
  PPTInspectorProps,
} from './PPTInspectorContract'

type PPTSlideInspectorPanelProps = Pick<
  PPTInspectorProps,
  'model' | 'onAction'
>

export function PPTSlideInspectorPanel({
  model,
  onAction,
}: PPTSlideInspectorPanelProps) {
  const {
    layoutDescriptors,
    layoutPlaceholderVisibilityDescriptors,
    layoutPlaceholders,
    lastPlaceholderVisibilityEffect,
    slide,
    slideLayoutId,
    slideMetadataDescriptor,
    slideThemeId,
    slideTransition,
    themeColorTokens,
  } = model
  const {
    onLayoutPlaceholderVisibilityChange,
    onSlideBackgroundChange,
    onSlideLayoutChange,
    onSlideNameChange,
    onSlideNotesChange,
    onSlideTransitionChange,
  } = createPPTInspectorActionDispatcher(onAction)
  const nameMetadataField = getPPTSlideMetadataField(
    slideMetadataDescriptor,
    'name',
  )
  const backgroundMetadataField = getPPTSlideMetadataField(
    slideMetadataDescriptor,
    'background',
  )
  const notesMetadataField = getPPTSlideMetadataField(
    slideMetadataDescriptor,
    'notes',
  )
  const sizeMetadataField = getPPTSlideMetadataField(
    slideMetadataDescriptor,
    'size',
  )
  const orientationMetadataField = getPPTSlideMetadataField(
    slideMetadataDescriptor,
    'orientation',
  )
  const slideTransitionDescriptor = createSlideEditTransitionDescriptor({
    advance: {
      afterMs: slideTransition.advanceAfterMs ?? undefined,
      onClick: slideTransition.advanceOnClick,
    },
    durationMs: slideTransition.durationMs,
    slideId: slide.id,
    type: slideTransition.type,
  })
  const layoutPlaceholderById = new Map(layoutPlaceholders.map((placeholder) => [
    placeholder.placeholderId,
    placeholder,
  ]))
  const hiddenPlaceholderCount = layoutPlaceholderVisibilityDescriptors
    .filter((placeholder) => !placeholder.isVisible).length

  return (
    <>
      <label className="ppt-field" {...getPPTSlideMetadataFieldData(nameMetadataField)}>
        <span>Name</span>
        <input
          data-ppt-slide-field="name"
          value={slide.name}
          onChange={(event) => onSlideNameChange(event.target.value)}
        />
      </label>
      <label className="ppt-field" {...getPPTSlideMetadataFieldData(backgroundMetadataField)}>
        <span>Background</span>
        <input
          data-ppt-slide-field="background"
          type="color"
          value={slide.background?.color ?? '#ffffff'}
          onChange={(event) => onSlideBackgroundChange(event.target.value)}
        />
      </label>
      <label className="ppt-field">
        <span>Layout</span>
        <select
          data-ppt-slide-field="layout"
          value={slideLayoutId}
          onChange={(event) => onSlideLayoutChange(event.target.value)}
        >
          {layoutDescriptors.map((layout) => (
            <option key={layout.layoutId} value={layout.layoutId}>
              {layout.name}
            </option>
          ))}
        </select>
      </label>
      <div
        className="ppt-theme-token-strip"
        data-ppt-layout-id={slideLayoutId}
        data-ppt-theme-id={slideThemeId}
      >
        {themeColorTokens.map((token) => (
          <span
            className="ppt-theme-token"
            data-ppt-theme-token={token.tokenId}
            data-ppt-theme-token-role={token.role}
            key={token.tokenId}
            style={{ background: token.value }}
            title={token.label}
          />
        ))}
      </div>
      <div
        className="ppt-layout-placeholder-list"
        data-ppt-layout-placeholder-count={layoutPlaceholders.length}
        data-ppt-layout-placeholder-hidden-count={hiddenPlaceholderCount}
        data-ppt-placeholder-visibility-command={lastPlaceholderVisibilityEffect?.payload.id}
        data-ppt-placeholder-visibility-command-placeholder={lastPlaceholderVisibilityEffect?.payload.placeholderId}
        data-ppt-placeholder-visibility-command-slide={lastPlaceholderVisibilityEffect?.selection.slideId}
        data-ppt-placeholder-visibility-command-type={lastPlaceholderVisibilityEffect?.type}
        data-ppt-placeholder-visibility-command-visible={lastPlaceholderVisibilityEffect
          ? String(lastPlaceholderVisibilityEffect.payload.isVisible)
          : undefined}
      >
        {layoutPlaceholderVisibilityDescriptors.map((placeholder) => {
          const resolvedPlaceholder = layoutPlaceholderById.get(placeholder.placeholderId)

          return (
            <div
              className="ppt-layout-placeholder"
              data-ppt-layout-placeholder={placeholder.placeholderId}
              data-ppt-placeholder-bounds={`${placeholder.bounds.x},${placeholder.bounds.y},${placeholder.bounds.w},${placeholder.bounds.h}`}
              data-ppt-placeholder-layout={resolvedPlaceholder?.layoutId}
              data-ppt-placeholder-locked={placeholder.isLocked ? 'true' : 'false'}
              data-ppt-placeholder-master={resolvedPlaceholder?.masterId}
              data-ppt-placeholder-role={placeholder.role}
              data-ppt-placeholder-slide={placeholder.slideId}
              data-ppt-placeholder-visible={placeholder.isVisible ? 'true' : 'false'}
              key={placeholder.placeholderId}
            >
              <span>{placeholder.title}</span>
              <button
                aria-label={`${placeholder.isVisible ? 'Hide' : 'Show'} ${placeholder.title}`}
                className="ppt-placeholder-visibility-toggle"
                data-ppt-placeholder-visibility-toggle={placeholder.placeholderId}
                disabled={placeholder.isLocked}
                title={placeholder.isVisible ? 'Hide placeholder' : 'Show placeholder'}
                type="button"
                onClick={() => onLayoutPlaceholderVisibilityChange(
                  placeholder.placeholderId,
                  !placeholder.isVisible,
                )}
              >
                {placeholder.isVisible ? <Eye size={14} /> : <EyeOff size={14} />}
              </button>
            </div>
          )
        })}
      </div>
      <label className="ppt-field" {...getPPTSlideMetadataFieldData(notesMetadataField)}>
        <span>Notes</span>
        <textarea
          data-ppt-slide-field="notes"
          value={slide.notes ?? ''}
          onChange={(event) => onSlideNotesChange(event.target.value)}
        />
      </label>
      <div
        className="ppt-slide-transition-fields"
        data-ppt-slide-transition
        data-ppt-transition-advance-after={slideTransitionDescriptor.advance.afterMs ?? ''}
        data-ppt-transition-advance-on-click={slideTransitionDescriptor.advance.onClick ? 'true' : 'false'}
        data-ppt-transition-duration={slideTransitionDescriptor.durationMs}
        data-ppt-transition-model="slide-edit-slide-transition-timing"
        data-ppt-transition-slide={slideTransitionDescriptor.slideId}
        data-ppt-transition-type={slideTransitionDescriptor.type}
        data-ppt-transition-types={SLIDE_EDIT_TRANSITION_TYPES.map((type) => type.id).join(' ')}
      >
        <label className="ppt-field">
          <span>Transition</span>
          <select
            data-ppt-slide-transition-field="type"
            value={slideTransition.type}
            onChange={(event) => onSlideTransitionChange(
              'type',
              event.target.value as PPTSlideTransition['type'],
            )}
          >
            {SLIDE_EDIT_TRANSITION_TYPES.map((type) => (
              <option key={type.id} value={type.id}>
                {type.label}
              </option>
            ))}
          </select>
        </label>
        <label className="ppt-field">
          <span>Duration</span>
          <input
            data-ppt-slide-transition-field="durationMs"
            max={SLIDE_EDIT_TRANSITION_TIMING_LIMITS.maxDurationMs}
            min={0}
            step={100}
            type="number"
            value={slideTransition.durationMs}
            onChange={(event) => onSlideTransitionChange(
              'durationMs',
              Number(event.target.value),
            )}
          />
        </label>
        <label className="ppt-checkbox-field">
          <input
            checked={slideTransition.advanceOnClick}
            data-ppt-slide-transition-field="advanceOnClick"
            type="checkbox"
            onChange={(event) => onSlideTransitionChange(
              'advanceOnClick',
              event.target.checked,
            )}
          />
          <span>On click</span>
        </label>
        <label className="ppt-field">
          <span>After</span>
          <input
            data-ppt-slide-transition-field="advanceAfterMs"
            max={SLIDE_EDIT_TRANSITION_TIMING_LIMITS.maxAdvanceAfterMs}
            min={0}
            placeholder="none"
            step={500}
            type="number"
            value={slideTransition.advanceAfterMs ?? ''}
            onChange={(event) => onSlideTransitionChange(
              'advanceAfterMs',
              parsePPTSlideTransitionAdvanceAfter(event.target.value),
            )}
          />
        </label>
      </div>
      <div className="ppt-slide-metadata-readouts">
        <span
          className="ppt-slide-metadata-readout"
          data-ppt-slide-field="size"
          data-ppt-slide-metadata-value={`${slideMetadataDescriptor.metadata.size.w}x${slideMetadataDescriptor.metadata.size.h}`}
          {...getPPTSlideMetadataFieldData(sizeMetadataField)}
        >
          <span>Size</span>
          <strong>{slideMetadataDescriptor.metadata.size.w} x {slideMetadataDescriptor.metadata.size.h}</strong>
        </span>
        <span
          className="ppt-slide-metadata-readout"
          data-ppt-slide-field="orientation"
          data-ppt-slide-metadata-value={slideMetadataDescriptor.metadata.orientation}
          {...getPPTSlideMetadataFieldData(orientationMetadataField)}
        >
          <span>Orientation</span>
          <strong>{slideMetadataDescriptor.metadata.orientation}</strong>
        </span>
      </div>
    </>
  )
}

function getPPTSlideMetadataField(
  descriptor: PPTInspectorModel['slideMetadataDescriptor'],
  fieldId: SlideEditSlideMetadataFieldId,
): SlideEditSlideMetadataFieldDescriptor {
  const field = descriptor.fields.find((candidate) => candidate.id === fieldId)

  if (!field) {
    throw new Error(`Missing PPT slide metadata field: ${fieldId}`)
  }

  return field
}

function getPPTSlideMetadataFieldData(
  field: SlideEditSlideMetadataFieldDescriptor,
) {
  return {
    'data-ppt-slide-metadata-adapter-slot': field.requiredAdapterSlot,
    'data-ppt-slide-metadata-command': field.commandId,
    'data-ppt-slide-metadata-control': field.control,
    'data-ppt-slide-metadata-editable': String(field.isEditable),
    'data-ppt-slide-metadata-field': field.id,
    'data-ppt-slide-metadata-optional': String(field.isOptional),
  }
}

function parsePPTSlideTransitionAdvanceAfter(value: string) {
  return value.trim() === '' ? null : Number(value)
}
