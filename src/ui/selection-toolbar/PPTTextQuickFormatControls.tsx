import {
  ALargeSmall,
  Baseline,
  Bold,
  Eraser,
  Highlighter,
  Italic,
  List,
  ListIndentDecrease,
  ListIndentIncrease,
  ListOrdered,
  Minus,
  Plus,
  Strikethrough,
  Underline,
} from 'lucide-react'
import { PPT_TEXT_FONT_SIZE_STEP } from '../../pptCanvasAppAffordanceAdapter'
import { SLIDE_EDIT_TEXT_PARAGRAPH_LIST_LEVEL_FIELD } from '../../pptSlideEditAffordanceAdapter'
import { IconButton } from '../core'
import {
  PPTParagraphAlignRadioGroup,
  type PPTParagraphAlign,
} from '../text-formatting'

export type PPTTextQuickFormatState = {
  align: PPTParagraphAlign
  bullet: boolean
  canDecreaseListLevel: boolean
  canIncreaseListLevel: boolean
  color: string
  fontSize: number
  highlight: string
  isBold: boolean
  isItalic: boolean
  isStrikethrough: boolean
  isUnderline: boolean
  listLevel: number
  numbered: boolean
}

export type PPTTextQuickFormatAction =
  | { type: 'clear-formatting' }
  | { type: 'set-color'; color: string }
  | { type: 'set-highlight'; color: string }
  | { type: 'set-paragraph-align'; align: PPTParagraphAlign }
  | { type: 'step-font-size'; delta: number }
  | { type: 'step-list-level'; delta: -1 | 1 }
  | { type: 'toggle-bold' }
  | { type: 'toggle-bullet' }
  | { type: 'toggle-italic' }
  | { type: 'toggle-numbered' }
  | { type: 'toggle-strikethrough' }
  | { type: 'toggle-underline' }

export type PPTTextQuickFormatControlsProps = {
  expanded: boolean
  state: PPTTextQuickFormatState
  onAction: (action: PPTTextQuickFormatAction) => void
}

export function PPTTextQuickFormatControls({
  expanded,
  onAction,
  state,
}: PPTTextQuickFormatControlsProps) {
  return (
    <span
      className="ppt-text-quick-format"
      data-ppt-text-quick-bar
      data-ppt-text-quick-expanded={expanded ? 'true' : 'false'}
    >
      <IconButton
        aria-pressed={state.isBold}
        data-ppt-text-quick="bold"
        label="Bold text"
        variant="floating"
        onClick={(event) => {
          event.preventDefault()
          event.stopPropagation()
          onAction({ type: 'toggle-bold' })
        }}
      >
        <Bold size={16} />
      </IconButton>
      <IconButton
        aria-pressed={state.isItalic}
        data-ppt-text-quick="italic"
        label="Italic text"
        variant="floating"
        onClick={(event) => {
          event.preventDefault()
          event.stopPropagation()
          onAction({ type: 'toggle-italic' })
        }}
      >
        <Italic size={16} />
      </IconButton>
      <IconButton
        aria-pressed={state.isUnderline}
        data-ppt-text-quick="underline"
        label="Underline text"
        variant="floating"
        onClick={(event) => {
          event.preventDefault()
          event.stopPropagation()
          onAction({ type: 'toggle-underline' })
        }}
      >
        <Underline size={16} />
      </IconButton>
      <IconButton
        aria-pressed={state.isStrikethrough}
        data-ppt-text-quick="strikethrough"
        label="Strikethrough text"
        variant="floating"
        onClick={(event) => {
          event.preventDefault()
          event.stopPropagation()
          onAction({ type: 'toggle-strikethrough' })
        }}
      >
        <Strikethrough size={16} />
      </IconButton>
      <IconButton
        data-ppt-text-quick="clear-formatting"
        label="Clear formatting"
        variant="floating"
        onClick={(event) => {
          event.preventDefault()
          event.stopPropagation()
          onAction({ type: 'clear-formatting' })
        }}
      >
        <Eraser size={16} />
      </IconButton>
      <IconButton
        data-ppt-text-quick="font-size-down"
        label="Decrease font size"
        variant="floating"
        onClick={(event) => {
          event.preventDefault()
          event.stopPropagation()
          onAction({
            delta: -PPT_TEXT_FONT_SIZE_STEP,
            type: 'step-font-size',
          })
        }}
      >
        <Minus size={16} />
      </IconButton>
      <span className="ppt-font-size-chip" data-ppt-text-quick-size>
        <ALargeSmall size={15} />
        {state.fontSize}
      </span>
      <IconButton
        data-ppt-text-quick="font-size-up"
        label="Increase font size"
        variant="floating"
        onClick={(event) => {
          event.preventDefault()
          event.stopPropagation()
          onAction({
            delta: PPT_TEXT_FONT_SIZE_STEP,
            type: 'step-font-size',
          })
        }}
      >
        <Plus size={16} />
      </IconButton>
      <label className="ppt-floating-color" title="Text color">
        <Baseline size={15} />
        <input
          aria-label="Text color"
          data-ppt-text-quick="color"
          type="color"
          value={state.color}
          onChange={(event) => onAction({
            color: event.target.value,
            type: 'set-color',
          })}
          onPointerDown={(event) => event.stopPropagation()}
        />
      </label>
      <label className="ppt-floating-color" title="Highlight color">
        <Highlighter size={15} />
        <input
          aria-label="Highlight color"
          data-ppt-text-quick="highlight"
          type="color"
          value={state.highlight}
          onChange={(event) => onAction({
            color: event.target.value,
            type: 'set-highlight',
          })}
          onPointerDown={(event) => event.stopPropagation()}
        />
      </label>
      <IconButton
        aria-pressed={state.bullet}
        data-ppt-text-quick="bullet"
        label="Toggle bullet list"
        variant="floating"
        onClick={(event) => {
          event.preventDefault()
          event.stopPropagation()
          onAction({ type: 'toggle-bullet' })
        }}
      >
        <List size={16} />
      </IconButton>
      <IconButton
        aria-pressed={state.numbered}
        data-ppt-text-quick="numbered"
        label="Toggle numbered list"
        variant="floating"
        onClick={(event) => {
          event.preventDefault()
          event.stopPropagation()
          onAction({ type: 'toggle-numbered' })
        }}
      >
        <ListOrdered size={16} />
      </IconButton>
      <IconButton
        data-ppt-text-quick="list-level-down"
        data-ppt-text-quick-list-level-command={SLIDE_EDIT_TEXT_PARAGRAPH_LIST_LEVEL_FIELD.commandIds.decrease}
        data-ppt-text-quick-list-level-control={SLIDE_EDIT_TEXT_PARAGRAPH_LIST_LEVEL_FIELD.control}
        data-ppt-text-quick-list-level={state.listLevel}
        data-ppt-text-quick-list-level-max={SLIDE_EDIT_TEXT_PARAGRAPH_LIST_LEVEL_FIELD.max}
        data-ppt-text-quick-list-level-min={SLIDE_EDIT_TEXT_PARAGRAPH_LIST_LEVEL_FIELD.min}
        data-ppt-text-quick-list-level-step={SLIDE_EDIT_TEXT_PARAGRAPH_LIST_LEVEL_FIELD.step}
        data-ppt-text-quick-list-level-surface="text-paragraph-list-level"
        disabled={!state.canDecreaseListLevel}
        label="Decrease list level"
        variant="floating"
        onClick={(event) => {
          event.preventDefault()
          event.stopPropagation()
          onAction({ delta: -1, type: 'step-list-level' })
        }}
      >
        <ListIndentDecrease size={16} />
      </IconButton>
      <IconButton
        data-ppt-text-quick="list-level-up"
        data-ppt-text-quick-list-level-command={SLIDE_EDIT_TEXT_PARAGRAPH_LIST_LEVEL_FIELD.commandIds.increase}
        data-ppt-text-quick-list-level-control={SLIDE_EDIT_TEXT_PARAGRAPH_LIST_LEVEL_FIELD.control}
        data-ppt-text-quick-list-level={state.listLevel}
        data-ppt-text-quick-list-level-max={SLIDE_EDIT_TEXT_PARAGRAPH_LIST_LEVEL_FIELD.max}
        data-ppt-text-quick-list-level-min={SLIDE_EDIT_TEXT_PARAGRAPH_LIST_LEVEL_FIELD.min}
        data-ppt-text-quick-list-level-step={SLIDE_EDIT_TEXT_PARAGRAPH_LIST_LEVEL_FIELD.step}
        data-ppt-text-quick-list-level-surface="text-paragraph-list-level"
        disabled={!state.canIncreaseListLevel}
        label="Increase list level"
        variant="floating"
        onClick={(event) => {
          event.preventDefault()
          event.stopPropagation()
          onAction({ delta: 1, type: 'step-list-level' })
        }}
      >
        <ListIndentIncrease size={16} />
      </IconButton>
      <PPTParagraphAlignRadioGroup
        align={state.align}
        surface="quick"
        onAlignChange={(align) => onAction({
          align,
          type: 'set-paragraph-align',
        })}
      />
    </span>
  )
}
