import { useCallback, useEffect, useState } from 'react'

export function useExclusiveDisclosure<DisclosureId extends string>() {
  const [activeDisclosure, setActiveDisclosure] =
    useState<DisclosureId | null>(null)

  const close = useCallback(() => {
    setActiveDisclosure(null)
  }, [])

  const isOpen = useCallback(
    (id: DisclosureId) => activeDisclosure === id,
    [activeDisclosure],
  )

  const toggle = useCallback((id: DisclosureId) => {
    setActiveDisclosure((current) => current === id ? null : id)
  }, [])

  useEffect(() => {
    if (activeDisclosure === null) {
      return
    }

    const closeOnEscape = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        close()
      }
    }

    window.addEventListener('keydown', closeOnEscape)
    return () => window.removeEventListener('keydown', closeOnEscape)
  }, [activeDisclosure, close])

  return {
    activeDisclosure,
    close,
    isOpen,
    toggle,
  }
}
