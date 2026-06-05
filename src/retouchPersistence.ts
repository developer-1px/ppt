import { useEffect, useMemo, useState } from 'react'
import { createDirtyState } from '@zod-crud/dirty-state'
import { createDocumentPersistence } from '@zod-crud/persist-web'
import type { DocumentPersistenceCodec } from '@zod-crud/persist-web'
import type { JSONDocument } from 'zod-crud'
import {
  SAMPLE_DECK,
  RetouchDeckSchema,
  type RetouchDeck,
} from './retouchModel'

const STORAGE_KEY = 'ppt-retouch:v3:deck'
const STORAGE_VERSION = 1

const retouchDraftCodec: DocumentPersistenceCodec = {
  encode(input) {
    return JSON.stringify({
      version: STORAGE_VERSION,
      deck: input.value,
      selection: input.selection,
      savedAt: input.savedAt,
    })
  },

  decode(text) {
    const value: unknown = JSON.parse(text)

    if (
      value &&
      typeof value === 'object' &&
      'version' in value &&
      value.version === STORAGE_VERSION &&
      'deck' in value
    ) {
      const candidate = value as {
        deck: unknown
        savedAt?: unknown
        selection?: unknown
      }

      return {
        value: candidate.deck,
        selection: null,
        savedAt: typeof candidate.savedAt === 'string' ? candidate.savedAt : null,
      }
    }

    return {
      value,
      selection: null,
      savedAt: null,
    }
  },
}

function deckEquals(a: unknown, b: unknown) {
  const parsedA = RetouchDeckSchema.safeParse(a)
  const parsedB = RetouchDeckSchema.safeParse(b)

  if (!parsedA.success || !parsedB.success) {
    return false
  }

  return JSON.stringify(parsedA.data) === JSON.stringify(parsedB.data)
}

export function createRetouchDocumentPersistence(doc: JSONDocument<RetouchDeck>) {
  return createDocumentPersistence(doc, {
    codec: retouchDraftCodec,
    key: STORAGE_KEY,
  })
}

export function createRetouchDirtyState(doc: JSONDocument<RetouchDeck>) {
  return createDirtyState(doc, {
    equals: deckEquals,
  })
}

export function useRetouchDraftPersistence(doc: JSONDocument<RetouchDeck>) {
  const persistence = useMemo(() => createRetouchDocumentPersistence(doc), [doc])
  const dirtyState = useMemo(() => createRetouchDirtyState(doc), [doc])
  const [hasDeckChanges, setHasDeckChanges] = useState(() =>
    dirtyState.isDirty(),
  )
  const [persistenceReady, setPersistenceReady] = useState(false)

  useEffect(() => {
    const unsubscribe = dirtyState.subscribe((snapshot) => {
      setHasDeckChanges(snapshot.dirty)
    })

    return () => {
      unsubscribe()
    }
  }, [dirtyState])

  useEffect(() => {
    let cancelled = false

    void persistence
      .restore({
        restoreSelection: true,
      })
      .then(() => {
        if (!cancelled) {
          setHasDeckChanges(dirtyState.isDirty())
          setPersistenceReady(true)
        }
      })

    return () => {
      cancelled = true
    }
  }, [dirtyState, persistence])

  useEffect(() => {
    if (!persistenceReady) {
      return
    }

    void (dirtyState.isDirty() ? persistence.save() : persistence.clear())
  }, [dirtyState, doc.value, persistence, persistenceReady])

  return {
    hasDeckChanges,
    persistenceReady,
  }
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
  return SAMPLE_DECK
}
