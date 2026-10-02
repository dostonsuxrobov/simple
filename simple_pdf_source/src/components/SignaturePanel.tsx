import { useEffect, useRef, useState } from 'react'
import type { KeyboardEvent as ReactKeyboardEvent, PointerEvent as ReactPointerEvent } from 'react'
import { Check, ImagePlus, Signature, Trash2, X } from 'lucide-react'
import { makeId } from '../lib/pdf'
import { cx, errorMessage } from '../lib/utils'
import { Button, IconButton } from './ui'

export interface SignatureImage {
  dataUrl: string
  width: number
  height: number
}

interface StoredSignature extends SignatureImage {
  id: string
  createdAt: number
}

interface SignaturePanelProps {
  onUse: (signature: SignatureImage) => void
  onClose: () => void
}

type CreateMode = 'draw' | 'type'

const SIGNATURES_KEY = 'folio:signatures:v1'
const LEGACY_SIGNATURES_KEY = 'simple.pdf.signatures'
const MAX_SIGNATURES = 10
const DRAW_WIDTH = 400
const DRAW_HEIGHT = 140
const DRAW_SCALE = 2

const INK_COLORS = [
  { label: 'Ink black', value: '#1b1b1f' },
  { label: 'Blue', value: '#1d4ed8' },
]

const SCRIPT_FONTS = [
  { label: 'Segoe Script', family: '"Segoe Script", "Brush Script MT", cursive' },
  { label: 'Brush Script', family: '"Brush Script MT", "Segoe Script", cursive' },
  { label: 'Lucida Handwriting', family: '"Lucida Handwriting", "Segoe Script", cursive' },
]

interface SignatureDraft {
  mode: CreateMode
  inkColor: string
  typedName: string
  fontFamily: string
  strokes: Array<Array<{ x: number; y: number }>>
}

// A signature drawn or typed but not used yet survives closing the panel
// (Escape, a click outside, Cancel) for the rest of the session.
let signatureDraft: SignatureDraft | null = null

function readStoredSignatures(): StoredSignature[] {
  try {
    let raw = localStorage.getItem(SIGNATURES_KEY)
    if (raw === null) {
      // One-time migration from the key shipped before the app-wide
      // 'folio:' prefix convention was applied here.
      raw = localStorage.getItem(LEGACY_SIGNATURES_KEY)
      if (raw !== null) {
        localStorage.setItem(SIGNATURES_KEY, raw)
        localStorage.removeItem(LEGACY_SIGNATURES_KEY)
      }
    }
    const value = JSON.parse(raw || '[]')
    return Array.isArray(value)
      ? value
        .filter((item) => item
          && typeof item.id === 'string'
          && typeof item.dataUrl === 'string'
          && item.dataUrl.startsWith('data:image/')
          && Number.isFinite(item.width) && item.width > 0
          && Number.isFinite(item.height) && item.height > 0)
        .slice(0, MAX_SIGNATURES)
      : []
  } catch {
    return []
  }
}

function writeStoredSignatures(next: StoredSignature[]) {
  try {
    localStorage.setItem(SIGNATURES_KEY, JSON.stringify(next))
  } catch {
    // Persisting signatures is best-effort only.
  }
}

/** Crop the canvas to its drawn content, dropping transparent and near-white margins. */
function trimmedSignature(source: HTMLCanvasElement): SignatureImage | null {
  const context = source.getContext('2d')
  if (!context || !source.width || !source.height) return null
  const data = context.getImageData(0, 0, source.width, source.height).data
  let minX = source.width
  let minY = source.height
  let maxX = -1
  let maxY = -1
  for (let y = 0; y < source.height; y += 1) {
    for (let x = 0; x < source.width; x += 1) {
      const offset = (y * source.width + x) * 4
      if (data[offset + 3] < 16) continue
      if (data[offset] > 244 && data[offset + 1] > 244 && data[offset + 2] > 244) continue
      if (x < minX) minX = x
      if (x > maxX) maxX = x
      if (y < minY) minY = y
      if (y > maxY) maxY = y
    }
  }
  if (maxX < minX || maxY < minY) return null
  const pad = 6
  minX = Math.max(0, minX - pad)
  minY = Math.max(0, minY - pad)
  maxX = Math.min(source.width - 1, maxX + pad)
  maxY = Math.min(source.height - 1, maxY + pad)
  const output = document.createElement('canvas')
  output.width = maxX - minX + 1
  output.height = maxY - minY + 1
  const outputContext = output.getContext('2d')
  if (!outputContext) return null
  outputContext.drawImage(source, minX, minY, output.width, output.height, 0, 0, output.width, output.height)
  return { dataUrl: output.toDataURL('image/png'), width: output.width, height: output.height }
}

function typedSignatureCanvas(text: string, family: string, color: string) {
  const canvas = document.createElement('canvas')
  const measureContext = canvas.getContext('2d')
  if (!measureContext) return null
  const font = `64px ${family}`
  measureContext.font = font
  canvas.width = Math.ceil(measureContext.measureText(text).width) + 56
  canvas.height = 170
  const context = canvas.getContext('2d')
  if (!context) return null
  context.font = font
  context.fillStyle = color
  context.textBaseline = 'middle'
  context.fillText(text, 28, canvas.height / 2)
  return canvas
}

async function importedSignatureCanvas(dataUrl: string) {
  const image = new Image()
  await new Promise<void>((resolve, reject) => {
    image.onload = () => resolve()
    image.onerror = () => reject(new Error('The image could not be decoded.'))
    image.src = dataUrl
  })
  const sourceWidth = image.naturalWidth || image.width
  const sourceHeight = image.naturalHeight || image.height
  if (!sourceWidth || !sourceHeight) throw new Error('The image could not be decoded.')
  const scale = Math.min(1, 1600 / Math.max(sourceWidth, sourceHeight))
  const canvas = document.createElement('canvas')
  canvas.width = Math.max(1, Math.round(sourceWidth * scale))
  canvas.height = Math.max(1, Math.round(sourceHeight * scale))
  const context = canvas.getContext('2d')
  if (!context) throw new Error('The image could not be decoded.')
  context.drawImage(image, 0, 0, canvas.width, canvas.height)
  return canvas
}

export function SignaturePanel({ onUse, onClose }: SignaturePanelProps) {
  const [signatures, setSignatures] = useState<StoredSignature[]>(readStoredSignatures)
  const [mode, setMode] = useState<CreateMode>(() => signatureDraft?.mode ?? 'draw')
  const [inkColor, setInkColor] = useState(() => signatureDraft?.inkColor ?? INK_COLORS[0].value)
  const [typedName, setTypedName] = useState(() => signatureDraft?.typedName ?? '')
  const [fontFamily, setFontFamily] = useState(() => signatureDraft?.fontFamily ?? SCRIPT_FONTS[0].family)
  const [hasDrawing, setHasDrawing] = useState(() => Boolean(signatureDraft?.strokes.length))
  const [createError, setCreateError] = useState('')
  const drawCanvasRef = useRef<HTMLCanvasElement>(null)
  const strokesRef = useRef<Array<Array<{ x: number; y: number }>>>(signatureDraft?.strokes ?? [])
  const typeInputRef = useRef<HTMLInputElement>(null)
  const draftRef = useRef({ mode, inkColor, typedName, fontFamily })
  draftRef.current = { mode, inkColor, typedName, fontFamily }
  // Set once the signature is used: nothing is left to keep.
  const usedRef = useRef(false)

  useEffect(() => () => {
    const draft = { ...draftRef.current, strokes: strokesRef.current }
    signatureDraft = usedRef.current || (!draft.strokes.length && !draft.typedName.trim()) ? null : draft
  }, [])

  function redrawStrokes() {
    const canvas = drawCanvasRef.current
    const context = canvas?.getContext('2d')
    if (!canvas || !context) return
    context.setTransform(1, 0, 0, 1, 0, 0)
    context.clearRect(0, 0, canvas.width, canvas.height)
    context.setTransform(DRAW_SCALE, 0, 0, DRAW_SCALE, 0, 0)
    context.strokeStyle = inkColor
    context.fillStyle = inkColor
    context.lineWidth = 2.4
    context.lineCap = 'round'
    context.lineJoin = 'round'
    for (const stroke of strokesRef.current) {
      if (!stroke.length) continue
      if (stroke.length === 1) {
        context.beginPath()
        context.arc(stroke[0].x, stroke[0].y, 1.2, 0, Math.PI * 2)
        context.fill()
        continue
      }
      context.beginPath()
      context.moveTo(stroke[0].x, stroke[0].y)
      for (let index = 1; index < stroke.length - 1; index += 1) {
        const midX = (stroke[index].x + stroke[index + 1].x) / 2
        const midY = (stroke[index].y + stroke[index + 1].y) / 2
        context.quadraticCurveTo(stroke[index].x, stroke[index].y, midX, midY)
      }
      const last = stroke[stroke.length - 1]
      context.lineTo(last.x, last.y)
      context.stroke()
    }
  }

  useEffect(() => {
    redrawStrokes()
  }, [inkColor, mode])

  useEffect(() => {
    if (mode === 'type') typeInputRef.current?.focus()
  }, [mode])

  function drawPoint(event: ReactPointerEvent<HTMLCanvasElement>) {
    const bounds = event.currentTarget.getBoundingClientRect()
    return {
      x: Math.min(DRAW_WIDTH, Math.max(0, event.clientX - bounds.left)),
      y: Math.min(DRAW_HEIGHT, Math.max(0, event.clientY - bounds.top)),
    }
  }

  function beginStroke(event: ReactPointerEvent<HTMLCanvasElement>) {
    if (event.button !== 0) return
    event.preventDefault()
    event.currentTarget.setPointerCapture(event.pointerId)
    strokesRef.current.push([drawPoint(event)])
    redrawStrokes()
    if (!hasDrawing) setHasDrawing(true)
  }

  function extendStroke(event: ReactPointerEvent<HTMLCanvasElement>) {
    if (!event.currentTarget.hasPointerCapture(event.pointerId)) return
    const stroke = strokesRef.current.at(-1)
    if (!stroke) return
    const point = drawPoint(event)
    const previous = stroke.at(-1)
    if (previous && Math.hypot(point.x - previous.x, point.y - previous.y) < 1.2) return
    stroke.push(point)
    redrawStrokes()
  }

  function endStroke(event: ReactPointerEvent<HTMLCanvasElement>) {
    if (event.currentTarget.hasPointerCapture(event.pointerId)) event.currentTarget.releasePointerCapture(event.pointerId)
  }

  function clearDrawing() {
    strokesRef.current = []
    setHasDrawing(false)
    setCreateError('')
    redrawStrokes()
  }

  function storeAndUse(signature: SignatureImage) {
    const stored: StoredSignature = { id: makeId('signature'), ...signature, createdAt: Date.now() }
    const next = [stored, ...signatures].slice(0, MAX_SIGNATURES)
    setSignatures(next)
    writeStoredSignatures(next)
    usedRef.current = true
    onUse(signature)
  }

  function removeSignature(id: string) {
    const next = signatures.filter((signature) => signature.id !== id)
    setSignatures(next)
    writeStoredSignatures(next)
  }

  const canCreate = mode === 'draw' ? hasDrawing : Boolean(typedName.trim())

  function createSignature() {
    if (!canCreate) return
    const source = mode === 'draw'
      ? drawCanvasRef.current
      : typedSignatureCanvas(typedName.trim(), fontFamily, inkColor)
    const signature = source ? trimmedSignature(source) : null
    if (!signature) {
      setCreateError('The signature is empty — draw or type something first.')
      return
    }
    storeAndUse(signature)
  }

  async function importImage() {
    try {
      const picked = await window.simple.pickImage()
      if (!picked) return
      const signature = trimmedSignature(await importedSignatureCanvas(picked.dataUrl))
      if (!signature) {
        setCreateError('That image looks blank — pick one with a visible signature.')
        return
      }
      storeAndUse(signature)
    } catch (error) {
      setCreateError(errorMessage(error))
    }
  }

  function handleKeyDown(event: ReactKeyboardEvent<HTMLDivElement>) {
    if (event.key === 'Escape') {
      event.preventDefault()
      event.stopPropagation()
      onClose()
      return
    }
    if (event.key === 'Enter' && !(event.target instanceof HTMLButtonElement)) {
      event.preventDefault()
      event.stopPropagation()
      createSignature()
      return
    }
    if (event.key === 'Tab') {
      const focusable = [...event.currentTarget.querySelectorAll<HTMLElement>('button, input')]
        .filter((element) => !element.hasAttribute('disabled'))
      if (!focusable.length) return
      const first = focusable[0]
      const last = focusable[focusable.length - 1]
      if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last.focus() }
      else if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first.focus() }
    }
  }

  return (
    <div className="print-dialog-overlay" onPointerDown={(event) => event.target === event.currentTarget && onClose()}>
      <div className="print-dialog signature-dialog" role="dialog" aria-modal="true" aria-label="Sign" onKeyDown={handleKeyDown}>
        <header className="print-dialog-header">
          <Signature size={15} strokeWidth={1.8} aria-hidden="true" />
          <strong>Sign</strong>
          <span className="print-dialog-document">Pick a signature, then click the page to place it</span>
          <IconButton icon={X} label="Close signing" compact onClick={onClose} />
        </header>

        <div className="print-dialog-body">
          {signatures.length > 0 && (
            <div className="print-field">
              <span>Saved signatures</span>
              <div className="signature-grid">
                {signatures.map((signature) => (
                  <div key={signature.id} className="signature-item">
                    <button type="button" className="signature-choice" title="Use this signature" onClick={() => onUse(signature)}>
                      <img src={signature.dataUrl} alt="Saved signature" draggable={false} />
                    </button>
                    <button type="button" className="signature-delete" aria-label="Delete signature" title="Delete signature" onClick={() => removeSignature(signature.id)}>
                      <Trash2 size={11} />
                    </button>
                  </div>
                ))}
              </div>
            </div>
          )}

          <div className="print-field">
            <span>{signatures.length ? 'New signature' : 'Create your signature'}</span>
            <div className="signature-tabs" role="tablist" aria-label="Signature style">
              <button type="button" role="tab" aria-selected={mode === 'draw'} className={cx('signature-tab', mode === 'draw' && 'is-active')} onClick={() => { setMode('draw'); setCreateError('') }}>Draw</button>
              <button type="button" role="tab" aria-selected={mode === 'type'} className={cx('signature-tab', mode === 'type' && 'is-active')} onClick={() => { setMode('type'); setCreateError('') }}>Type</button>
            </div>

            {mode === 'draw' ? (
              <div className="signature-draw">
                <canvas
                  ref={drawCanvasRef}
                  className="signature-draw-canvas"
                  width={DRAW_WIDTH * DRAW_SCALE}
                  height={DRAW_HEIGHT * DRAW_SCALE}
                  style={{ width: DRAW_WIDTH, height: DRAW_HEIGHT }}
                  aria-label="Draw your signature"
                  onPointerDown={beginStroke}
                  onPointerMove={extendStroke}
                  onPointerUp={endStroke}
                  onPointerCancel={endStroke}
                />
                {!hasDrawing && <span className="signature-draw-hint" aria-hidden="true">Draw your signature here</span>}
              </div>
            ) : (
              <div className="signature-type">
                <input
                  ref={typeInputRef}
                  type="text"
                  aria-label="Type your name"
                  placeholder="Type your name"
                  maxLength={60}
                  value={typedName}
                  onChange={(event) => { setTypedName(event.target.value); setCreateError('') }}
                />
                <div className="signature-font-choices" role="radiogroup" aria-label="Signature font">
                  {SCRIPT_FONTS.map((font) => (
                    <button
                      key={font.label}
                      type="button"
                      role="radio"
                      aria-checked={fontFamily === font.family}
                      className={cx('signature-font-choice', fontFamily === font.family && 'is-active')}
                      style={{ fontFamily: font.family, color: inkColor }}
                      title={font.label}
                      onClick={() => setFontFamily(font.family)}
                    >
                      {typedName.trim() || 'Signature'}
                    </button>
                  ))}
                </div>
              </div>
            )}

            <div className="signature-options">
              <div className="signature-colors" role="radiogroup" aria-label="Ink color">
                {INK_COLORS.map((color) => (
                  <button
                    key={color.value}
                    type="button"
                    role="radio"
                    aria-checked={inkColor === color.value}
                    aria-label={color.label}
                    title={color.label}
                    className={cx('signature-color', inkColor === color.value && 'is-active')}
                    style={{ background: color.value }}
                    onClick={() => setInkColor(color.value)}
                  />
                ))}
              </div>
              {mode === 'draw' && <Button variant="ghost" disabled={!hasDrawing} onClick={clearDrawing}>Clear</Button>}
            </div>
            {createError && <small className="print-range-error">{createError}</small>}
          </div>
        </div>

        <footer className="print-dialog-footer">
          <Button icon={ImagePlus} variant="ghost" onClick={() => { void importImage() }}>Import image</Button>
          <Button variant="secondary" onClick={onClose}>Cancel</Button>
          <Button icon={Check} variant="primary" disabled={!canCreate} onClick={createSignature}>Use signature</Button>
        </footer>
      </div>
    </div>
  )
}
