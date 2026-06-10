import {
  RetouchDeckSchema,
  type RetouchDeck,
} from './retouchModel'
import { SAMPLE_DECK } from './sampleDeck'

const BASE_DECK = RetouchDeckSchema.parse(SAMPLE_DECK)
const BASE_SLIDES_BY_ID = new Map(
  BASE_DECK.slides.map((slide) => [slide.id, slide]),
)

export function retouchDeckEquals(a: unknown, b: unknown) {
  const parsedA = parseRetouchDeck(a)
  const parsedB = parseRetouchDeck(b)

  if (parsedA === null || parsedB === null) {
    return false
  }

  return jsonEquals(parsedA, parsedB)
}

export function changedSlideIdsFromBaseline(deck: unknown) {
  const parsed = parseRetouchDeck(deck)

  if (parsed === null) {
    return new Set<string>()
  }

  return new Set(
    parsed.slides
      .filter((slide) => {
        const baselineSlide = BASE_SLIDES_BY_ID.get(slide.id)

        return baselineSlide === undefined || !jsonEquals(slide, baselineSlide)
      })
      .map((slide) => slide.id),
  )
}

function parseRetouchDeck(value: unknown): RetouchDeck | null {
  const parsed = RetouchDeckSchema.safeParse(value)

  return parsed.success ? parsed.data : null
}

function jsonEquals(a: unknown, b: unknown) {
  return JSON.stringify(a) === JSON.stringify(b)
}
