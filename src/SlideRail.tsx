import {
  type ButtonHTMLAttributes,
  type CSSProperties,
  type HTMLAttributes,
  useMemo,
} from 'react'
import {
  ArrowDown,
  ArrowUp,
  Copy,
  Plus,
  Trash2,
} from 'lucide-react'
import {
  listboxDefinition,
  reducePatternData,
  useListboxPattern,
  useToolbarPattern,
  type PatternData,
  type PatternEvent,
  type ReactToolbarRenderItem,
} from '@interactive-os/aria/react'
import { getRect, rectToStyle, type RetouchSlide } from './retouchModel'

type SlideRailProps = {
  activeSlideId: string
  canDeleteSlide: boolean
  canMoveSlideDown: boolean
  canMoveSlideUp: boolean
  canvasViewTablistProps: HTMLAttributes<HTMLDivElement>
  canvasViewTabProps: Record<'slide' | 'grid', ButtonHTMLAttributes<HTMLButtonElement>>
  changedSlideIds: Set<string>
  onAddSlide: () => void
  onDeleteSlide: () => void
  onDuplicateSlide: () => void
  onMoveSlideDown: () => void
  onMoveSlideUp: () => void
  onSelectSlide: (slideId: string) => void
  slides: RetouchSlide[]
}

type SlideRailActionKey =
  | 'add'
  | 'delete'
  | 'duplicate'
  | 'move-down'
  | 'move-up'

const SLIDE_RAIL_ACTION_KEYS = [
  'add',
  'duplicate',
  'move-up',
  'move-down',
  'delete',
] as const

export function SlideRail({
  activeSlideId,
  canDeleteSlide,
  canMoveSlideDown,
  canMoveSlideUp,
  canvasViewTablistProps,
  canvasViewTabProps,
  changedSlideIds,
  onAddSlide,
  onDeleteSlide,
  onDuplicateSlide,
  onMoveSlideDown,
  onMoveSlideUp,
  onSelectSlide,
  slides,
}: SlideRailProps) {
  const slideRailData = useMemo<PatternData>(() => {
    const slideIds = slides.map((slide) => slide.id)

    return {
      items: Object.fromEntries(
        slides.map((slide) => [
          slide.id,
          {
            label: changedSlideIds.has(slide.id)
              ? `${slide.name}, modified`
              : slide.name,
            textValue: slide.name,
          },
        ]),
      ),
      relations: { rootKeys: slideIds },
      state: {
        activeKey: activeSlideId,
        selectedKeys: [activeSlideId],
      },
      refs: { label: 'Slides' },
    }
  }, [activeSlideId, changedSlideIds, slides])

  const slideRailListbox = useListboxPattern(
    slideRailData,
    (event: PatternEvent) => {
      if (event.type === 'select') {
        const selectedSlideId = event.keys[0]

        if (selectedSlideId) {
          onSelectSlide(selectedSlideId)
        }
        return
      }

      if (event.type === 'navigate') {
        const nextSlideId = reducePatternData(
          listboxDefinition,
          slideRailData,
          event,
        ).state?.activeKey

        if (nextSlideId) {
          onSelectSlide(nextSlideId)
        }
      }
    },
    {
      elementIdPrefix: 'slide-thumb-',
      focusStrategy: 'rovingTabIndex',
      orientation: 'vertical',
      selectionMode: 'single',
      typeaheadEnabled: true,
    },
  )
  const railActionDisabledKeys = useMemo(() => [
    ...(!canMoveSlideUp ? ['move-up' as const] : []),
    ...(!canMoveSlideDown ? ['move-down' as const] : []),
    ...(!canDeleteSlide ? ['delete' as const] : []),
  ], [canDeleteSlide, canMoveSlideDown, canMoveSlideUp])
  const railActionData = useMemo<PatternData>(() => ({
    items: {
      add: { label: 'Add slide' },
      delete: { label: 'Delete slide' },
      duplicate: { label: 'Duplicate slide' },
      'move-down': { label: 'Move slide down' },
      'move-up': { label: 'Move slide up' },
    },
    relations: { rootKeys: SLIDE_RAIL_ACTION_KEYS },
    refs: { label: 'Slide actions' },
    state: {
      activeKey: firstEnabledToolbarKey(
        SLIDE_RAIL_ACTION_KEYS,
        railActionDisabledKeys,
      ),
      disabledKeys: railActionDisabledKeys,
    },
  }), [railActionDisabledKeys])
  const railActionToolbar = useToolbarPattern(
    railActionData,
    (event) => {
      for (const key of selectedToolbarKeys(event)) {
        if (key === 'add') {
          onAddSlide()
        }
        if (key === 'duplicate') {
          onDuplicateSlide()
        }
        if (key === 'move-up') {
          onMoveSlideUp()
        }
        if (key === 'move-down') {
          onMoveSlideDown()
        }
        if (key === 'delete') {
          onDeleteSlide()
        }
      }
    },
    {
      elementIdPrefix: 'slide-action-',
      orientation: 'horizontal',
    },
  )
  const railActionProps = toolbarItemPropsByKey<SlideRailActionKey>(
    railActionToolbar.renderItems,
  )

  return (
    <aside className="slide-rail">
      <div {...canvasViewTablistProps} className="view-toggle">
        <button
          {...canvasViewTabProps.slide}
          type="button"
        >
          Slide
        </button>
        <button
          {...canvasViewTabProps.grid}
          type="button"
        >
          Grid
        </button>
      </div>

      <div {...railActionToolbar.rootProps} className="rail-actions">
        <button
          {...railActionProps.add}
          aria-label="Add slide"
          title="Add slide"
          type="button"
        >
          <Plus aria-hidden="true" size={15} strokeWidth={2.2} />
        </button>
        <button
          {...railActionProps.duplicate}
          aria-label="Duplicate slide"
          title="Duplicate slide"
          type="button"
        >
          <Copy aria-hidden="true" size={15} strokeWidth={2.2} />
        </button>
        <button
          {...railActionProps['move-up']}
          aria-label="Move slide up"
          disabled={!canMoveSlideUp}
          title="Move slide up"
          type="button"
        >
          <ArrowUp aria-hidden="true" size={15} strokeWidth={2.2} />
        </button>
        <button
          {...railActionProps['move-down']}
          aria-label="Move slide down"
          disabled={!canMoveSlideDown}
          title="Move slide down"
          type="button"
        >
          <ArrowDown aria-hidden="true" size={15} strokeWidth={2.2} />
        </button>
        <button
          {...railActionProps.delete}
          aria-label="Delete slide"
          disabled={!canDeleteSlide}
          title="Delete slide"
          type="button"
        >
          <Trash2 aria-hidden="true" size={15} strokeWidth={2.2} />
        </button>
      </div>

      <div {...slideRailListbox.rootProps} className="slide-list">
        {slideRailListbox.renderItems.map((item, index) => {
          const slide = slides.find((candidate) => candidate.id === item.key)

          if (!slide) {
            return null
          }

          const changed = changedSlideIds.has(slide.id)
          const optionProps =
            item.optionProps as ButtonHTMLAttributes<HTMLButtonElement>

          return (
            <button
              {...optionProps}
              aria-current={slide.id === activeSlideId ? 'page' : undefined}
              aria-label={changed ? `${slide.name}, modified` : slide.name}
              className="slide-thumb"
              data-changed={changed ? 'true' : 'false'}
              key={slide.id}
              type="button"
            >
              <span className="thumb-number">{index + 1}</span>
              <MiniSlide slide={slide} />
              <span className="thumb-name">{slide.name}</span>
              {changed ? <span aria-hidden="true" className="thumb-change" /> : null}
            </button>
          )
        })}
      </div>
    </aside>
  )
}

function selectedToolbarKeys(event: PatternEvent) {
  return event.type === 'select' ? event.keys : []
}

function toolbarItemPropsByKey<TKey extends string>(
  items: readonly ReactToolbarRenderItem[],
) {
  return Object.fromEntries(
    items.map((item) => [
      item.key,
      item.itemProps as ButtonHTMLAttributes<HTMLButtonElement>,
    ]),
  ) as Record<TKey, ButtonHTMLAttributes<HTMLButtonElement>>
}

function firstEnabledToolbarKey<TKey extends string>(
  keys: readonly TKey[],
  disabledKeys: readonly string[],
) {
  return keys.find((key) => !disabledKeys.includes(key)) ?? keys[0] ?? null
}

function MiniSlide({ slide }: { slide: RetouchSlide }) {
  return (
    <span className="mini-slide" style={{ '--accent': slide.accent } as CSSProperties}>
      {slide.blocks.map((block) => (
        <span
          className={`mini-block ${block.role}`}
          key={block.id}
          style={rectToStyle(getRect(block))}
        />
      ))}
    </span>
  )
}
