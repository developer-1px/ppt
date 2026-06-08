import {
  RetouchDeckSchema,
  type RetouchDeck,
  type RetouchSlide,
} from './retouchModel'
import { SAMPLE_DECK } from './sampleDeck'

const BASE_DECK = RetouchDeckSchema.parse(SAMPLE_DECK)
const BASE_SLIDES_BY_ID = new Map(
  BASE_DECK.slides.map((slide) => [slide.id, slide]),
)

export function retouchDeckEquals(a: unknown, b: unknown) {
  const parsedA = parseRetouchDeck(a)
  const parsedB = parseRetouchDeck(b)

  return Boolean(parsedA && parsedB && jsonEquals(parsedA, parsedB))
}

export function changedSlideIdsFromBaseline(deck: unknown) {
  const parsed = parseRetouchDeck(deck)

  if (!parsed) {
    return new Set<string>()
  }

  return new Set(
    parsed.slides
      .filter((slide) => !retouchSlideEquals(slide, BASE_SLIDES_BY_ID.get(slide.id)))
      .map((slide) => slide.id),
  )
}

function parseRetouchDeck(value: unknown): RetouchDeck | null {
  const parsed = RetouchDeckSchema.safeParse(value)

  return parsed.success ? parsed.data : null
}

function retouchSlideEquals(
  slide: RetouchSlide,
  baselineSlide: RetouchSlide | undefined,
) {
  return Boolean(baselineSlide && jsonEquals(slide, baselineSlide))
}

function jsonEquals(a: unknown, b: unknown) {
  return JSON.stringify(a) === JSON.stringify(b)
}
