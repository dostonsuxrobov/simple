// Browser entry for scripts/qa-print-dialog.cjs: the real Print and Export dialogs with the
// props App.tsx passes, the real print engine rendering their previews, and a sheet whose
// page setup the dialog edits through applyPageSetupPatch. Calls are recorded on
// window.harness; nothing is printed or saved.
import { useCallback, useMemo, useState } from 'react'
import { createRoot } from 'react-dom/client'
import '../src/styles.css'
import { SpreadsheetPrintDialog } from '../src/components/SpreadsheetPrintDialog'
import type { SpreadsheetPreviewPage, SpreadsheetPrinterInfo } from '../src/components/SpreadsheetPrintDialog'
import { SpreadsheetExportDialog } from '../src/components/SpreadsheetExportDialog'
import { applyPageSetupPatch, pageSetupStateFor } from '../src/lib/print-page-setup'
import type { SpreadsheetPageSetupPatch } from '../src/lib/print-page-setup'
import type { SpreadsheetPrintPreview } from '../src/spreadsheet-types'
import { visualPrintInput } from './qa-print-visuals-fixture'
// The print engine is plain JavaScript; main.cjs runs this same module for the preview.
import { createSpreadsheetPrintDocument } from '../electron/spreadsheet-print.cjs'

interface Harness {
  previews: unknown[]
  pages: SpreadsheetPreviewPage[][]
  printed: Array<Record<string, unknown>>
  pdf: Array<Record<string, unknown>>
  patches: SpreadsheetPageSetupPatch[]
  mode: string
}

declare global {
  interface Window {
    harness: Harness
  }
}

window.harness = { previews: [], pages: [], printed: [], pdf: [], patches: [], mode: new URLSearchParams(location.search).get('mode') || 'print' }

const base = visualPrintInput()
const sheet = base.workbook.sheets[0]
for (let row = 7; row <= 140; row += 1) sheet.cells[`A${row}`] = { value: `Row ${row}` }
sheet.rowCount = 140
sheet.frozen = { rows: 1, columns: 1 }
const selection = { top: 19, bottom: 29, left: 1, right: 3 }
const printers: SpreadsheetPrinterInfo[] = [{ name: 'Office Laser', displayName: 'Office Laser' }, { name: 'Label', displayName: 'Label Printer', description: 'Thermal' }]

function Harness() {
  const [revision, setRevision] = useState(0)
  const [open, setOpen] = useState(true)
  const mode = window.harness.mode
  // eslint-disable-next-line react-hooks/exhaustive-deps
  const pageSetup = useMemo(() => pageSetupStateFor(sheet, selection), [revision])
  const render = useCallback(async (options: unknown): Promise<SpreadsheetPrintPreview> => {
    window.harness.previews.push(options)
    const document = createSpreadsheetPrintDocument({ ...base, selection, options })
    return {
      html: document.html, title: document.title, sheets: document.sheetCount, cells: document.printedCells, pages: document.pageCount,
      pageBreaks: document.pageBreaks, minimumScale: document.minimumScale, oversizedDimensions: document.oversizedDimensions,
      paper: document.paper, options: document.options, warnings: document.warnings,
    }
  }, [])
  const listPrinters = useCallback(async () => printers, [])
  if (!open) return <div id="closed">closed</div>
  if (mode === 'export') {
    return <SpreadsheetExportDialog sheetCount={1} selectionLabel="B20:D30" defaultGridlines compatibilityWarning={false} onClose={() => setOpen(false)} onExport={async (format, options) => { window.harness.pdf.push({ format, options }); return true }} />
  }
  const wired = mode !== 'noprinters'
  return (
    <SpreadsheetPrintDialog
      sheetCount={1}
      selectionLabel="B20:D30"
      defaultGridlines
      compatibilityWarning={false}
      onClose={() => setOpen(false)}
      onRenderPreview={render}
      onPrint={async (options) => { window.harness.printed.push(options as unknown as Record<string, unknown>); return { printed: true, canceled: false, sheets: 1, cells: 1, pages: 1 } }}
      listPrinters={wired ? listPrinters : async () => { throw new Error('No handler registered for workbook:list-printers') }}
      onSavePdf={wired ? async (options) => { window.harness.pdf.push(options as unknown as Record<string, unknown>); return false } : undefined}
      pageSetup={wired ? pageSetup : undefined}
      onPageSetupChange={wired ? (patch) => { window.harness.patches.push(patch); applyPageSetupPatch(sheet, patch); setRevision((value) => value + 1) } : undefined}
      onPreviewPages={wired ? (pages) => { window.harness.pages.push(pages) } : undefined}
    />
  )
}

createRoot(document.getElementById('root')!).render(<Harness />)
