import type { CSSProperties } from 'react'

type CSSVariableName = `--${string}`
type CSSVariableValue = string | number | undefined

export function cssVariables(
  variables: Partial<Record<CSSVariableName, CSSVariableValue>>,
): CSSProperties {
  return variables as CSSProperties
}
