import { ListIndentDecrease, ListIndentIncrease, Maximize2 } from 'lucide-react'
import { isPPTTextElement, readPPTText } from '../../pptModel'
import {
  getPPTParagraphListLevel,
  getPPTTextAutoFitSizeMode,
  getPPTTextElementInset,
  getPPTTextElementParagraphSpacing,
  getPPTTextElementStyle,
  getPPTTextElementVerticalAlign,
  getPPTTextFontFamilyDescriptor,
  getPPTTextFrameInsetDescriptor,
  getPPTTextFrameInsetField,
  getPPTTextParagraphListLevelDescriptor,
  getPPTTextParagraphSpacingDescriptor,
  getPPTTextParagraphSpacingField,
  getPPTTextVerticalAlignmentDescriptor,
  hasPPTTextBodyBullet,
  hasPPTTextBodyNumbered,
  normalizePPTTextFontFamily,
  parsePPTParagraphLineHeight,
  parsePPTParagraphSpacing,
  parsePPTTextInset,
  PPT_PARAGRAPH_LINE_HEIGHT_MAX,
  PPT_PARAGRAPH_LINE_HEIGHT_MIN,
  PPT_PARAGRAPH_LIST_LEVEL_MAX,
  PPT_PARAGRAPH_LIST_LEVEL_MIN,
  PPT_PARAGRAPH_SPACING_MAX,
  PPT_TEXT_FONT_FAMILY_OPTIONS,
  PPT_TEXT_INSET_MAX,
  PPT_TEXT_INSET_MIN,
  PPT_TEXT_INSET_STEP,
  PPT_TEXT_VERTICAL_ALIGNMENT_OPTIONS,
} from '../../pptTextAdapter'
import { Button } from '../core'
import { PPTParagraphAlignRadioGroup } from '../text-formatting'
import { createPPTInspectorActionDispatcher } from './PPTInspectorActionDispatcher'
import type { PPTInspectorProps } from './PPTInspectorContract'
import { PPTColorSwatchStrip } from './PPTColorSwatchStrip'

type PPTTextInspectorFieldsProps = Pick<
  PPTInspectorProps,
  'model' | 'onAction'
>

export function PPTTextInspectorFields({
  model,
  onAction,
}: PPTTextInspectorFieldsProps) {
  const element = model.selectedElement
  const {
    onCommitText,
    onElementTextInsetChange,
    onElementTextStyleChange,
    onParagraphAlignChange,
    onParagraphBulletChange,
    onParagraphListLevelStep,
    onParagraphNumberedChange,
    onParagraphSpacingChange,
    onTextAutoFit,
  } = createPPTInspectorActionDispatcher(onAction)

  if (!element || !isPPTTextElement(element)) {
    return null
  }

  const textStyle = getPPTTextElementStyle(element)
  const paragraph = element.textBody.paragraphs[0] ?? { runs: [] }
  const paragraphAlign = paragraph.align ?? 'left'
  const paragraphBullet = hasPPTTextBodyBullet(element.textBody)
  const paragraphNumbered = hasPPTTextBodyNumbered(element.textBody)
  const paragraphListLevel = getPPTParagraphListLevel(paragraph)
  const paragraphSpacing = getPPTTextElementParagraphSpacing(element)
  const paragraphSpacingDescriptor = getPPTTextParagraphSpacingDescriptor(
    model.slide.id,
    element,
  )
  const paragraphListLevelDescriptor = getPPTTextParagraphListLevelDescriptor(
    model.slide.id,
    element,
  )
  const paragraphLineHeightField = getPPTTextParagraphSpacingField(
    paragraphSpacingDescriptor,
    'lineHeightRatio',
  )
  const paragraphBeforeField = getPPTTextParagraphSpacingField(
    paragraphSpacingDescriptor,
    'paragraphBefore',
  )
  const paragraphAfterField = getPPTTextParagraphSpacingField(
    paragraphSpacingDescriptor,
    'paragraphAfter',
  )
  const textInset = getPPTTextElementInset(element)
  const textFontFamilyDescriptor = getPPTTextFontFamilyDescriptor(
    model.slide.id,
    element,
  )
  const textFrameInsetDescriptor = getPPTTextFrameInsetDescriptor(
    model.slide.id,
    element,
  )
  const textVerticalAlignmentDescriptor =
    getPPTTextVerticalAlignmentDescriptor(model.slide.id, element)

  return (
    <>
      <label className="ppt-field">
        <span>Text</span>
        <textarea
          data-ppt-style-field="text"
          value={readPPTText(element.textBody)}
          onChange={(event) => onCommitText(element.id, event.target.value)}
        />
      </label>
      <div className="ppt-geometry-grid">
        <div className="ppt-color-control" data-ppt-color-control="text-color">
          <label className="ppt-field">
            <span>Text color</span>
            <input
              data-ppt-style-field="text-color"
              type="color"
              value={textStyle.color ?? '#111827'}
              onChange={(event) =>
                onElementTextStyleChange(
                  element.id,
                  'color',
                  event.target.value,
                )}
            />
          </label>
          <PPTColorSwatchStrip
            model={model}
            onAction={onAction}
            target={{
              channel: 'text-color',
              color: textStyle.color ?? '#111827',
              elementId: element.id,
            }}
          />
        </div>
        <label className="ppt-field">
          <span>Font size</span>
          <input
            data-ppt-style-field="font-size"
            type="number"
            value={textStyle.fontSize ?? 24}
            onChange={(event) =>
              onElementTextStyleChange(
                element.id,
                'fontSize',
                Number(event.target.value),
              )}
          />
        </label>
      </div>
      <label className="ppt-field">
        <span>Font</span>
        <select
          data-ppt-style-field="font-family"
          data-ppt-text-font-family-command={textFontFamilyDescriptor.field.commandId}
          data-ppt-text-font-family-control={textFontFamilyDescriptor.field.control}
          data-ppt-text-font-family-fallback={textFontFamilyDescriptor.fallbackFontFamily}
          data-ppt-text-font-family-options={textFontFamilyDescriptor.options
            .map((option) => option.family).join(' ')}
          data-ppt-text-font-family-surface={textFontFamilyDescriptor.surface}
          value={
            textFontFamilyDescriptor.fontFamily ??
            normalizePPTTextFontFamily(textStyle.fontFamily)
          }
          onChange={(event) =>
            onElementTextStyleChange(
              element.id,
              'fontFamily',
              event.target.value,
            )}
        >
          {PPT_TEXT_FONT_FAMILY_OPTIONS.map((option) => (
            <option key={option.value} value={option.value}>
              {option.label}
            </option>
          ))}
        </select>
      </label>
      <label className="ppt-field">
        <span>Vertical</span>
        <select
          data-ppt-style-field="vertical-align"
          data-ppt-text-vertical-align-attribute={
            textVerticalAlignmentDescriptor.metadata.attribute
          }
          data-ppt-text-vertical-align-attribute-value={
            textVerticalAlignmentDescriptor.metadata.value
          }
          data-ppt-text-vertical-align-command={textVerticalAlignmentDescriptor.field.commandId}
          data-ppt-text-vertical-align-control={textVerticalAlignmentDescriptor.field.control}
          data-ppt-text-vertical-align-default-value={
            textVerticalAlignmentDescriptor.metadata.defaultValue
          }
          data-ppt-text-vertical-align-options={textVerticalAlignmentDescriptor.field.options
            .map((option) => option.id).join(' ')}
          data-ppt-text-vertical-align-surface={textVerticalAlignmentDescriptor.surface}
          value={textVerticalAlignmentDescriptor.value ?? getPPTTextElementVerticalAlign(element)}
          onChange={(event) =>
            onElementTextStyleChange(
              element.id,
              'verticalAlign',
              event.target.value,
            )}
        >
          {PPT_TEXT_VERTICAL_ALIGNMENT_OPTIONS.map((option) => (
            <option key={option.id} value={option.id}>
              {option.label}
            </option>
          ))}
        </select>
      </label>
      <div
        className="ppt-paragraph-spacing-grid"
        data-ppt-text-inset-attribute={textFrameInsetDescriptor.metadata.attribute}
        data-ppt-text-inset-attribute-value={textFrameInsetDescriptor.metadata.value}
        data-ppt-text-inset-bottom={textFrameInsetDescriptor.inset.bottom ?? textInset.bottom}
        data-ppt-text-inset-default-value={textFrameInsetDescriptor.metadata.defaultValue}
        data-ppt-text-inset-inspector
        data-ppt-text-inset-left={textFrameInsetDescriptor.inset.left ?? textInset.left}
        data-ppt-text-inset-right={textFrameInsetDescriptor.inset.right ?? textInset.right}
        data-ppt-text-inset-surface={textFrameInsetDescriptor.surface}
        data-ppt-text-inset-top={textFrameInsetDescriptor.inset.top ?? textInset.top}
      >
        {(['top', 'right', 'bottom', 'left'] as const).map((field) => {
          const textFrameInsetField = getPPTTextFrameInsetField(
            textFrameInsetDescriptor,
            field,
          )

          return (
            <label className="ppt-field" key={field}>
              <span>{field[0].toUpperCase() + field.slice(1)}</span>
              <input
                data-ppt-text-inset-command={textFrameInsetField?.commandId}
                data-ppt-text-inset-control={textFrameInsetField?.control}
                data-ppt-text-inset-field={field}
                data-ppt-text-inset-unit={textFrameInsetField?.unit}
                max={textFrameInsetField?.max ?? PPT_TEXT_INSET_MAX}
                min={textFrameInsetField?.min ?? PPT_TEXT_INSET_MIN}
                step={textFrameInsetField?.step ?? PPT_TEXT_INSET_STEP}
                type="number"
                value={textFrameInsetDescriptor.inset[field] ?? textInset[field]}
                onChange={(event) =>
                  onElementTextInsetChange(
                    element.id,
                    field,
                    parsePPTTextInset(event.target.value),
                  )}
              />
            </label>
          )
        })}
      </div>
      <label className="ppt-field">
        <span>Weight</span>
        <select
          data-ppt-style-field="font-weight"
          value={textStyle.fontWeight ?? 'regular'}
          onChange={(event) =>
            onElementTextStyleChange(
              element.id,
              'fontWeight',
              event.target.value,
            )}
        >
          <option value="regular">Regular</option>
          <option value="semibold">Semibold</option>
          <option value="bold">Bold</option>
        </select>
      </label>
      <div className="ppt-field">
        <span>Paragraph</span>
        <div className="ppt-paragraph-control-row">
          <button
            aria-pressed={paragraphBullet}
            className="ppt-paragraph-bullet-button"
            data-ppt-paragraph-bullet
            type="button"
            onClick={() =>
              onParagraphBulletChange(element.id, !paragraphBullet)}
          >
            bullet
          </button>
          <button
            aria-pressed={paragraphNumbered}
            className="ppt-paragraph-bullet-button"
            data-ppt-paragraph-numbered
            type="button"
            onClick={() =>
              onParagraphNumberedChange(element.id, !paragraphNumbered)}
          >
            numbered
          </button>
          <button
            aria-label="Decrease list level"
            className="ppt-paragraph-bullet-button"
            data-ppt-paragraph-list-level-command={
              paragraphListLevelDescriptor.field.commandIds.decrease
            }
            data-ppt-paragraph-list-level-control={paragraphListLevelDescriptor.field.control}
            data-ppt-paragraph-list-level-down
            data-ppt-paragraph-list-level={paragraphListLevel}
            data-ppt-paragraph-list-level-can-decrease={
              paragraphListLevelDescriptor.canDecrease ? 'true' : 'false'
            }
            data-ppt-paragraph-list-level-can-increase={
              paragraphListLevelDescriptor.canIncrease ? 'true' : 'false'
            }
            data-ppt-paragraph-list-level-indent={paragraphListLevelDescriptor.indent.cssValue}
            data-ppt-paragraph-list-level-max={paragraphListLevelDescriptor.field.max}
            data-ppt-paragraph-list-level-min={paragraphListLevelDescriptor.field.min}
            data-ppt-paragraph-list-level-step={paragraphListLevelDescriptor.field.step}
            data-ppt-paragraph-list-level-surface={paragraphListLevelDescriptor.surface}
            disabled={paragraphListLevel <= PPT_PARAGRAPH_LIST_LEVEL_MIN}
            title="Decrease list level"
            type="button"
            onClick={() => onParagraphListLevelStep(element.id, -1)}
          >
            <ListIndentDecrease size={15} />
          </button>
          <button
            aria-label="Increase list level"
            className="ppt-paragraph-bullet-button"
            data-ppt-paragraph-list-level-command={
              paragraphListLevelDescriptor.field.commandIds.increase
            }
            data-ppt-paragraph-list-level-control={paragraphListLevelDescriptor.field.control}
            data-ppt-paragraph-list-level-up
            data-ppt-paragraph-list-level={paragraphListLevel}
            data-ppt-paragraph-list-level-can-decrease={
              paragraphListLevelDescriptor.canDecrease ? 'true' : 'false'
            }
            data-ppt-paragraph-list-level-can-increase={
              paragraphListLevelDescriptor.canIncrease ? 'true' : 'false'
            }
            data-ppt-paragraph-list-level-indent={paragraphListLevelDescriptor.indent.cssValue}
            data-ppt-paragraph-list-level-max={paragraphListLevelDescriptor.field.max}
            data-ppt-paragraph-list-level-min={paragraphListLevelDescriptor.field.min}
            data-ppt-paragraph-list-level-step={paragraphListLevelDescriptor.field.step}
            data-ppt-paragraph-list-level-surface={paragraphListLevelDescriptor.surface}
            disabled={paragraphListLevel >= PPT_PARAGRAPH_LIST_LEVEL_MAX}
            title="Increase list level"
            type="button"
            onClick={() => onParagraphListLevelStep(element.id, 1)}
          >
            <ListIndentIncrease size={15} />
          </button>
          <PPTParagraphAlignRadioGroup
            align={paragraphAlign}
            surface="inspector"
            onAlignChange={(align) => onParagraphAlignChange(element.id, align)}
          />
        </div>
      </div>
      <div
        className="ppt-paragraph-spacing-grid"
        data-ppt-paragraph-spacing-inspector
        data-ppt-paragraph-spacing-surface={paragraphSpacingDescriptor.surface}
        data-ppt-paragraph-line-height={
          paragraphSpacingDescriptor.values.lineHeightRatio ??
          paragraphSpacing.lineHeight
        }
        data-ppt-paragraph-spacing-after={
          paragraphSpacingDescriptor.values.paragraphAfter.value ??
          paragraphSpacing.spacingAfter
        }
        data-ppt-paragraph-spacing-before={
          paragraphSpacingDescriptor.values.paragraphBefore.value ??
          paragraphSpacing.spacingBefore
        }
      >
        <label className="ppt-field">
          <span>Line height</span>
          <input
            data-ppt-paragraph-control={paragraphLineHeightField?.control}
            data-ppt-paragraph-command={paragraphLineHeightField?.commandId}
            data-ppt-paragraph-field="lineHeight"
            max={paragraphLineHeightField?.max ?? PPT_PARAGRAPH_LINE_HEIGHT_MAX}
            min={paragraphLineHeightField?.min ?? PPT_PARAGRAPH_LINE_HEIGHT_MIN}
            step={paragraphLineHeightField?.step ?? 0.05}
            type="number"
            value={paragraphSpacingDescriptor.values.lineHeightRatio ?? paragraphSpacing.lineHeight}
            onChange={(event) =>
              onParagraphSpacingChange(
                element.id,
                'lineHeight',
                parsePPTParagraphLineHeight(event.target.value),
              )}
          />
        </label>
        <label className="ppt-field">
          <span>Before</span>
          <input
            data-ppt-paragraph-command={paragraphBeforeField?.commandId}
            data-ppt-paragraph-control={paragraphBeforeField?.control}
            data-ppt-paragraph-field="spacingBefore"
            data-ppt-paragraph-unit={paragraphBeforeField?.unit}
            max={paragraphBeforeField?.max ?? PPT_PARAGRAPH_SPACING_MAX}
            min={paragraphBeforeField?.min ?? 0}
            step={paragraphBeforeField?.step ?? 2}
            type="number"
            value={
              paragraphSpacingDescriptor.values.paragraphBefore.value ??
              paragraphSpacing.spacingBefore
            }
            onChange={(event) =>
              onParagraphSpacingChange(
                element.id,
                'spacingBefore',
                parsePPTParagraphSpacing(event.target.value),
              )}
          />
        </label>
        <label className="ppt-field">
          <span>After</span>
          <input
            data-ppt-paragraph-command={paragraphAfterField?.commandId}
            data-ppt-paragraph-control={paragraphAfterField?.control}
            data-ppt-paragraph-field="spacingAfter"
            data-ppt-paragraph-unit={paragraphAfterField?.unit}
            max={paragraphAfterField?.max ?? PPT_PARAGRAPH_SPACING_MAX}
            min={paragraphAfterField?.min ?? 0}
            step={paragraphAfterField?.step ?? 2}
            type="number"
            value={
              paragraphSpacingDescriptor.values.paragraphAfter.value ??
              paragraphSpacing.spacingAfter
            }
            onChange={(event) =>
              onParagraphSpacingChange(
                element.id,
                'spacingAfter',
                parsePPTParagraphSpacing(event.target.value),
              )}
          />
        </label>
      </div>
      <div
        className="ppt-text-overflow-control"
        data-ppt-text-autofit={element.textAutoFit}
        data-ppt-text-autofit-command={model.lastTextAutoFitEffect?.payload.id}
        data-ppt-text-autofit-command-handle={model.lastTextAutoFitEffect?.payload.handle}
        data-ppt-text-autofit-command-object={model.lastTextAutoFitEffect?.payload.objectId}
        data-ppt-text-autofit-command-size-mode={model.lastTextAutoFitEffect?.payload.sizeMode}
        data-ppt-text-autofit-command-type={model.lastTextAutoFitEffect?.type}
        data-ppt-text-autofit-model="slide-edit-text-box-auto-fit"
        data-ppt-text-autofit-size-mode={
          model.textAutoFitIndicator?.sizeMode ??
          getPPTTextAutoFitSizeMode(element)
        }
        data-ppt-text-overflow={model.selectedTextOverflow ? 'true' : 'false'}
        data-ppt-text-overflow-indicator-axis={model.textAutoFitIndicator?.overflowAxis.join(' ')}
        data-ppt-text-overflow-indicator-visible={model.textAutoFitIndicator
          ? String(model.textAutoFitIndicator.isVisible)
          : undefined}
        data-ppt-text-overflow-inspector
      >
        <span>{model.selectedTextOverflow ? 'Overflow' : 'Fits'}</span>
        <Button
          data-ppt-style-action="text-auto-fit"
          disabled={!model.selectedTextOverflow}
          onClick={() => onTextAutoFit(element.id)}
        >
          <Maximize2 size={15} /> Auto fit
        </Button>
      </div>
    </>
  )
}
