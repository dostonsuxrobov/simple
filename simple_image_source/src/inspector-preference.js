export const INSPECTOR_STORAGE_KEY = 'simple-image:inspector-open'

export function inspectorOpenFromStored(value) {
  // Closed until someone opens it: the window is a viewer first.
  return value === 'open'
}

export function storedInspectorOpen(value) {
  return value ? 'open' : 'closed'
}
