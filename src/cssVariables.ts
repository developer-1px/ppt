import type { CSSProperties } from 'react'

type CSSVariableName = `--${string}`
type CSSVariableValue = string | number | undefined
type CSSVariableProperties = CSSProperties &
  Partial<Record<CSSVariableName, CSSVariableValue>>

export function cssVariables(
  variables: CSSVariableProperties,
): CSSVariableProperties {
  return variables
}
