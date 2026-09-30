export function resolveTextFit(edit: { textFit?: 'fit' | 'wrap'; originalText?: string; text?: string }): 'fit' | 'wrap'
export function layoutText(text: string, width: number, measure: (text: string) => number, mode?: 'fit' | 'wrap'): { lines: string[]; fitScale: number }
