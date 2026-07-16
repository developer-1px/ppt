import type {
  PPTLine,
  PPTLineMarker,
  PPTLineRoute,
} from './pptModel'

export const PPT_LINE_ROUTE_OPTIONS = Object.freeze([
  { label: 'Straight', value: 'straight' },
  { label: 'Elbow', value: 'elbow' },
] as const satisfies readonly {
  label: string
  value: PPTLineRoute
}[])

export const PPT_LINE_MARKER_OPTIONS = Object.freeze([
  { label: 'None', value: 'none' },
  { label: 'Arrow', value: 'arrow' },
] as const satisfies readonly {
  label: string
  value: PPTLineMarker
}[])

export function getPPTLineRoute(line: PPTLine): PPTLineRoute {
  return line.route ?? 'straight'
}

export function getPPTLineMarker(
  line: PPTLine,
  field: 'endMarker' | 'startMarker',
): PPTLineMarker {
  return line[field] ?? 'none'
}

export function isPPTLineMarker(value: string): value is PPTLineMarker {
  return value === 'none' || value === 'arrow'
}

export function isPPTLineRoute(value: string): value is PPTLineRoute {
  return value === 'straight' || value === 'elbow'
}
