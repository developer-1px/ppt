import {
  AlignCenter,
  AlignJustify,
  AlignLeft,
  AlignRight,
} from 'lucide-react'
import {
  PPT_RADIO_GROUP_FOCUS_MODEL,
  PPT_RADIO_GROUP_KEYBOARD_MODEL,
  PPT_RADIO_GROUP_MODEL,
  focusPPTCanvasElementBySelectorOnNextFrame,
  getPPTCanvasRadioTabIndex,
  handlePPTCanvasRadioGroupKeyDown,
} from '../../pptCanvasAppAffordanceAdapter'
import type { PPTParagraph } from '../../pptModel'

export type PPTParagraphAlign = NonNullable<PPTParagraph['align']>

export type PPTParagraphAlignRadioGroupProps = {
  align: PPTParagraphAlign
  onAlignChange: (align: PPTParagraphAlign) => void
  surface: 'inspector' | 'quick'
}

const ALIGN_OPTIONS = [
  'left',
  'center',
  'right',
  'justify',
] as const satisfies readonly PPTParagraphAlign[]

export function PPTParagraphAlignRadioGroup({
  align,
  onAlignChange,
  surface,
}: PPTParagraphAlignRadioGroupProps) {
  function selectAlign(
    nextAlign: PPTParagraphAlign,
    container: HTMLElement | null,
  ) {
    onAlignChange(nextAlign)
    focusPPTCanvasElementBySelectorOnNextFrame<HTMLButtonElement>({
      match: ({ element }) =>
        element.getAttribute('data-ppt-paragraph-align') === nextAlign,
      root: container,
      selector: '[data-ppt-paragraph-align]',
    })
  }

  return (
    <span
      aria-label="Paragraph align"
      className={surface === 'quick'
        ? 'ppt-paragraph-align-radio-group ppt-paragraph-align-radio-group--quick'
        : 'ppt-segmented-control ppt-paragraph-align-radio-group'}
      data-ppt-paragraph-align-focus-model={PPT_RADIO_GROUP_FOCUS_MODEL}
      data-ppt-paragraph-align-keyboard-model={PPT_RADIO_GROUP_KEYBOARD_MODEL}
      data-ppt-paragraph-align-model={PPT_RADIO_GROUP_MODEL}
      data-ppt-paragraph-align-radiogroup={surface}
      role="radiogroup"
      onKeyDown={handlePPTCanvasRadioGroupKeyDown}
    >
      {ALIGN_OPTIONS.map((option) => {
        const selected = align === option

        return (
          <button
            aria-checked={selected}
            aria-label={`Align text ${option}`}
            aria-pressed={selected}
            className={surface === 'quick' ? 'ppt-floating-command' : undefined}
            data-ppt-paragraph-align={option}
            data-ppt-paragraph-align-surface={surface}
            data-ppt-text-quick={surface === 'quick' ? `align-${option}` : undefined}
            key={option}
            role="radio"
            tabIndex={getPPTCanvasRadioTabIndex({
              checked: selected,
              disabled: false,
            })}
            title={`Align text ${option}`}
            type="button"
            onClick={(event) => {
              event.preventDefault()
              event.stopPropagation()
              selectAlign(option, event.currentTarget.closest('[role="radiogroup"]'))
            }}
          >
            {surface === 'quick'
              ? <PPTTextAlignIcon align={option} size={16} />
              : option}
          </button>
        )
      })}
    </span>
  )
}

function PPTTextAlignIcon({
  align,
  size,
}: {
  align: PPTParagraphAlign
  size: number
}) {
  switch (align) {
    case 'center':
      return <AlignCenter size={size} />
    case 'justify':
      return <AlignJustify size={size} />
    case 'right':
      return <AlignRight size={size} />
    case 'left':
      return <AlignLeft size={size} />
  }
}
