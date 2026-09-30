export type PageDropEdge = 'before' | 'after'

export function pageIndicesForTransfer(selectedPages: ReadonlySet<number>, draggedPage: number) {
  const indices = selectedPages.has(draggedPage) ? [...selectedPages] : [draggedPage]
  return [...new Set(indices)]
    .filter((index) => Number.isInteger(index) && index >= 0)
    .sort((a, b) => a - b)
}

export function pageDropEdge(clientY: number, top: number, height: number): PageDropEdge {
  return clientY < top + Math.max(0, height) / 2 ? 'before' : 'after'
}

export function pageDropInsertIndex(pageIndex: number, edge: PageDropEdge) {
  return pageIndex + (edge === 'after' ? 1 : 0)
}

export function pageReorderDestination(draggedPage: number, insertIndex: number) {
  return insertIndex - (draggedPage < insertIndex ? 1 : 0)
}

export function isPdfTransferFile(file: Pick<File, 'name' | 'type'>) {
  return file.type.toLowerCase() === 'application/pdf' || /\.pdf$/i.test(file.name)
}

const IMPORTABLE_TRANSFER_TYPES = new Set(['application/pdf', 'image/png', 'image/jpeg'])

export function isImportableTransferFile(file: Pick<File, 'name' | 'type'>) {
  return IMPORTABLE_TRANSFER_TYPES.has(file.type.toLowerCase()) || /\.(pdf|png|jpe?g|gif|webp|bmp|avif|svg|txt|md|docx?)$/i.test(file.name) || /^image\//i.test(file.type)
}
