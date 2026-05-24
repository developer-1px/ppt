import { SAMPLE_DECK, RetouchDeckSchema } from './retouchModel'

const STORAGE_KEY = 'ppt-retouch:v3:deck'
const STORAGE_VERSION = 1

export function deckEquals(a: unknown, b: unknown) {
  const parsedA = RetouchDeckSchema.safeParse(a)
  const parsedB = RetouchDeckSchema.safeParse(b)

  if (!parsedA.success || !parsedB.success) {
    return false
  }

  return JSON.stringify(parsedA.data) === JSON.stringify(parsedB.data)
}

export function changedSlides(deck: unknown) {
  const parsed = RetouchDeckSchema.safeParse(deck)
  const parsedBase = RetouchDeckSchema.safeParse(SAMPLE_DECK)

  if (!parsed.success || !parsedBase.success) {
    return new Set<string>()
  }

  return new Set(
    parsed.data.slides
      .filter((slide) => {
        const baseSlide = parsedBase.data.slides.find(
          (candidate) => candidate.id === slide.id,
        )

        return !baseSlide || JSON.stringify(slide) !== JSON.stringify(baseSlide)
      })
      .map((slide) => slide.id),
  )
}

export function readInitialDeck() {
  if (typeof window === 'undefined') {
    return SAMPLE_DECK
  }

  try {
    const raw = window.localStorage.getItem(STORAGE_KEY)

    if (!raw) {
      return SAMPLE_DECK
    }

    const payload: unknown = JSON.parse(raw)

    if (
      !payload ||
      typeof payload !== 'object' ||
      !('version' in payload) ||
      payload.version !== STORAGE_VERSION ||
      !('deck' in payload)
    ) {
      window.localStorage.removeItem(STORAGE_KEY)
      return SAMPLE_DECK
    }

    const parsed = RetouchDeckSchema.safeParse(payload.deck)

    if (!parsed.success) {
      window.localStorage.removeItem(STORAGE_KEY)
      return SAMPLE_DECK
    }

    return parsed.data
  } catch {
    window.localStorage.removeItem(STORAGE_KEY)
    return SAMPLE_DECK
  }
}

export function persistDeck(deck: unknown) {
  if (typeof window === 'undefined') {
    return
  }

  try {
    if (deckEquals(deck, SAMPLE_DECK)) {
      window.localStorage.removeItem(STORAGE_KEY)
      return
    }

    window.localStorage.setItem(
      STORAGE_KEY,
      JSON.stringify({
        version: STORAGE_VERSION,
        deck,
      }),
    )
  } catch {
    // Autosave is best-effort; editing must keep working if storage is unavailable.
  }
}
