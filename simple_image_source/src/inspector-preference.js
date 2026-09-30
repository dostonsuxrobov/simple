export const INSPECTOR_STORAGE_KEY = 'simple-image:inspector-open'

export function inspectorOpenFromStored(value) {
  return value !== 'closed'
}

export function storedInspectorOpen(value) {
  return value ? 'open' : 'closed'
}
