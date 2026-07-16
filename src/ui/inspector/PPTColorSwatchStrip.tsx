import {
  getPPTColorSwatchDescriptor,
  type PPTColorSwatchTarget,
} from '../../pptColorSwatchAdapter'
import { createPPTInspectorActionDispatcher } from './PPTInspectorActionDispatcher'
import type { PPTInspectorProps } from './PPTInspectorContract'

type PPTColorSwatchStripProps = Pick<
  PPTInspectorProps,
  'model' | 'onAction'
> & {
  target: PPTColorSwatchTarget
}

export function PPTColorSwatchStrip({
  model,
  onAction,
  target,
}: PPTColorSwatchStripProps) {
  const { onColorSwatchApply } = createPPTInspectorActionDispatcher(onAction)
  const descriptor = getPPTColorSwatchDescriptor({
    channel: target.channel,
    currentColor: target.color,
    objectIds: [target.elementId],
    recentColors: model.recentColors,
    slideId: model.slide.id,
    themeColorTokens: model.themeColorTokens,
  })
  const swatchCount = descriptor.sections.reduce(
    (count, section) => count + section.swatches.length,
    0,
  )
  const recentSection = descriptor.sections.find(
    (section) => section.id === 'recent',
  )

  return (
    <div
      className="ppt-color-swatch-strip"
      data-ppt-color-swatch-channel={target.channel}
      data-ppt-color-swatch-command={descriptor.field.commandId}
      data-ppt-color-swatch-control={descriptor.field.control}
      data-ppt-color-swatch-count={swatchCount}
      data-ppt-color-swatch-disabled={descriptor.state.isDisabled ? 'true' : 'false'}
      data-ppt-color-swatch-disabled-reason={descriptor.state.disabledReason}
      data-ppt-color-swatch-mixed={descriptor.state.isMixed ? 'true' : 'false'}
      data-ppt-color-swatch-model={descriptor.surface}
      data-ppt-color-swatch-object-ids={descriptor.objectIds.join(' ')}
      data-ppt-color-swatch-package-channel={descriptor.channel.id}
      data-ppt-color-swatch-palette={target.channel}
      data-ppt-color-swatch-selected-id={descriptor.state.selectedSwatchId}
      data-ppt-color-swatch-strip={target.channel}
      data-ppt-recent-color-count={recentSection?.swatches.length ?? 0}
    >
      {descriptor.sections.map((section) =>
        section.swatches.length > 0 ? (
          <div
            className="ppt-color-swatch-group"
            data-ppt-color-swatch-group={section.id}
            key={section.id}
          >
            {section.swatches.map((swatch) => (
              <button
                aria-label={`${swatch.label} ${target.channel}`}
                aria-pressed={swatch.selected}
                className="ppt-color-swatch"
                data-ppt-color-source={swatch.source}
                data-ppt-color-swatch={target.channel}
                data-ppt-color-swatch-id={swatch.id}
                data-ppt-color-swatch-selected={swatch.selected ? 'true' : undefined}
                data-ppt-color-token={swatch.tokenId}
                data-ppt-color-value={swatch.value}
                disabled={descriptor.state.isDisabled}
                key={swatch.id}
                style={{ backgroundColor: swatch.value }}
                title={swatch.label}
                type="button"
                onClick={() =>
                  onColorSwatchApply(
                    target.elementId,
                    target.channel,
                    swatch.value,
                    {
                      source: swatch.source,
                      swatchId: swatch.id,
                      tokenId: swatch.tokenId,
                      value: swatch.value,
                    },
                  )}
              />
            ))}
          </div>
        ) : null)}
    </div>
  )
}
