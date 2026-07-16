import {
  createSlideEditColorSwatchPaletteDescriptor,
  normalizeSlideEditColorHex,
  normalizeSlideEditColorSwatchValue,
  SLIDE_EDIT_COLOR_SWATCH_CHANNELS,
  type SlideEditColorSwatchBuiltInChannelId,
  type SlideEditColorSwatchHostCommandEffect,
  type SlideEditColorSwatchPaletteDescriptor,
  type SlideEditColorSwatchSelection,
  type SlideEditThemeColorToken,
} from './pptSlideEditAffordanceAdapter'

export type PPTColorSwatchChannel =
  | 'line-stroke'
  | 'shape-fill'
  | 'shape-stroke'
  | 'text-color'

export type PPTColorSwatchPackageChannel =
  SlideEditColorSwatchBuiltInChannelId

type PPTColorSwatchDescriptor = SlideEditColorSwatchPaletteDescriptor<
  string,
  string,
  PPTColorSwatchPackageChannel,
  string
>

export type PPTColorSwatchHostCommandEffect =
  SlideEditColorSwatchHostCommandEffect<
    string,
    string,
    PPTColorSwatchPackageChannel,
    string
  >

export type PPTColorSwatchSelection = SlideEditColorSwatchSelection<string>

export type PPTColorSwatchTarget = {
  channel: PPTColorSwatchChannel
  color: string
  elementId: string
}

const PPT_COLOR_SWATCH_CHANNEL_MAP = {
  'line-stroke': 'line-stroke',
  'shape-fill': 'fill',
  'shape-stroke': 'stroke',
  'text-color': 'text',
} as const satisfies Record<
  PPTColorSwatchChannel,
  PPTColorSwatchPackageChannel
>

export function getPPTColorSwatchPackageChannel(
  channel: PPTColorSwatchChannel,
): PPTColorSwatchPackageChannel {
  return PPT_COLOR_SWATCH_CHANNEL_MAP[channel]
}

function getPPTColorSwatchChannelDescriptor(
  channel: PPTColorSwatchChannel,
) {
  const packageChannel = getPPTColorSwatchPackageChannel(channel)

  return SLIDE_EDIT_COLOR_SWATCH_CHANNELS.find((item) =>
    item.id === packageChannel) ?? {
    id: packageChannel,
    label: packageChannel,
  }
}

export function getPPTColorSwatchDescriptor({
  channel,
  currentColor,
  objectIds,
  recentColors,
  slideId,
  themeColorTokens,
}: {
  channel: PPTColorSwatchChannel
  currentColor: string
  objectIds: readonly string[]
  recentColors: readonly string[]
  slideId: string
  themeColorTokens: readonly SlideEditThemeColorToken[]
}): PPTColorSwatchDescriptor {
  const selectedValue = normalizePPTSwatchColor(currentColor) ||
    normalizeSlideEditColorSwatchValue(currentColor)
  const selectedTokenId = selectedValue
    ? themeColorTokens.find((token) =>
        normalizePPTSwatchColor(token.value) === selectedValue ||
        token.value === selectedValue,
      )?.tokenId ?? null
    : null

  return createSlideEditColorSwatchPaletteDescriptor({
    channel: getPPTColorSwatchChannelDescriptor(channel),
    disabledReason: objectIds.length > 0 ? undefined : 'no-selection',
    isDisabled: objectIds.length === 0,
    objectIds,
    recentColors,
    selectedTokenId,
    selectedValue,
    slideId,
    themeColorTokens,
  })
}

export function normalizePPTSwatchColor(color: string) {
  return normalizeSlideEditColorHex(color) ?? ''
}
