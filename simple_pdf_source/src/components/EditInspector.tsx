import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  Bold,
  Check,
  ClipboardPaste,
  Copy,
  FlipHorizontal2,
  FlipVertical2,
  ImagePlus,
  Italic,
  MousePointer2,
  RotateCcw,
  RotateCw,
  ScanSearch,
  Trash2,
  Type,
  X,
} from 'lucide-react'
import type { PageObjectEdit, PageTextEdit } from '../types'
import { clamp, cx } from '../lib/utils'
import { Button, IconButton } from './ui'

interface EditInspectorProps {
  pageHasNativeText: boolean | null
  textEdit: PageTextEdit | null
  objectEdit: PageObjectEdit | null
  selectingObjectRegion: boolean
  onTextChange: (edit: PageTextEdit) => void
  onObjectChange: (edit: PageObjectEdit) => void
  onCommitText: () => void
  onCommitObject: () => void
  onCancelSelection: () => void
  onDeleteSelection: () => void
  onCopySelection: () => void
  onPasteSelection: () => void
  canPasteSelection: boolean
  onDuplicateText: () => void
  onAddText: () => void
  onAddImage: () => void
  onReplaceImage: () => void
  onRotateObject: (direction: 'left' | 'right') => void
  onFlipObject: (direction: 'horizontal' | 'vertical') => void
  onDuplicateObject: () => void
  onSelectObjectRegion: () => void
  onClose: () => void
}

function rgbToHex(color: [number, number, number]) {
  return `#${color.map((channel) => Math.round(clamp(channel, 0, 1) * 255).toString(16).padStart(2, '0')).join('')}`
}

function hexToRgb(value: string): [number, number, number] {
  const normalized = value.replace('#', '')
  return [0, 2, 4].map((offset) => parseInt(normalized.slice(offset, offset + 2), 16) / 255) as [number, number, number]
}

export function EditInspector({
  textEdit, objectEdit, pageHasNativeText, selectingObjectRegion, onTextChange, onObjectChange,
  onCommitText, onCommitObject, onCancelSelection, onDeleteSelection, onAddText,
  onCopySelection, onPasteSelection, canPasteSelection, onDuplicateText, onAddImage, onReplaceImage, onRotateObject, onFlipObject, onDuplicateObject,
  onSelectObjectRegion, onClose,
}: EditInspectorProps) {
  const selected = textEdit || objectEdit
  const fontFamilies = ['Segoe UI', 'Arial', 'Times New Roman', 'Georgia', 'Courier New']
  const sourceFontLabel = textEdit?.fontFamily.split(',')
    .map((family) => family.replace(/["']/g, '').trim())
    .find((family) => !/^(g_|pdfjs|sans-serif$|serif$|monospace$)/i.test(family))

  return (
    <aside className="edit-inspector" aria-label="Edit PDF properties">
      <header className="edit-inspector-header">
        <div><strong>Edit PDF</strong><span>Direct page editing</span></div>
        <IconButton icon={X} label="Close edit mode" compact onClick={onClose} />
      </header>

      {!selected && (
        <div className="edit-inspector-empty">
          <span className="inspector-empty-icon"><MousePointer2 size={18} /></span>
          <strong>{pageHasNativeText === false ? 'This page has no editable text' : 'Select content on the page'}</strong>
          <p>{pageHasNativeText === false
            ? 'Scanned words are part of the picture. You can add text and notes, or edit the picture. Recognizing and changing the printed words is not available yet.'
            : 'Click text to place the caret inside it. Click an image to move, resize, replace, or remove it.'}</p>
          <div className="inspector-add-actions">
            <Button icon={Type} variant="secondary" onClick={onAddText}>Add text</Button>
            <Button icon={ImagePlus} variant="secondary" onClick={onAddImage}>Add image</Button>
            <Button icon={ClipboardPaste} variant="secondary" disabled={!canPasteSelection} onClick={onPasteSelection}>Paste</Button>
          </div>
          <button
            type="button"
            className={cx('region-select-button', selectingObjectRegion && 'is-active')}
            onClick={onSelectObjectRegion}
          >
            <ScanSearch size={16} />
            <span><strong>{selectingObjectRegion ? 'Draw on the page…' : 'Select artwork area'}</strong><small>For charts, logos, and grouped vector content</small></span>
          </button>
          <div className="edit-scope-note">
            <strong>How complex content is handled</strong>
            <p>Text stays searchable. Images remain images. A selected artwork area is reconstructed as one movable image so its appearance stays intact.</p>
          </div>
        </div>
      )}

      {textEdit && (
        <div className="inspector-content">
          <div className="selection-summary"><span className="selection-kind"><Type size={14} /> Text</span><small>Page {textEdit.pageIndex + 1}</small></div>
          <section className="inspector-section">
            <h3>Format</h3>
            <label className="inspector-field">
              <span>Font</span>
              <select aria-label="Font family" value={textEdit.fontFamily} onChange={(event) => onTextChange({
                ...textEdit,
                fontFamily: event.target.value,
                // A retained embedded subset represents the source face. Once
                // the user explicitly chooses another family, let the saver
                // resolve that requested family instead of silently reusing
                // the original embedded font.
                fontKey: undefined,
                fontData: undefined,
                preserveSourceMetrics: false,
                modified: true,
              })}>
                {!fontFamilies.includes(textEdit.fontFamily) && <option value={textEdit.fontFamily}>{sourceFontLabel || 'Original font'} (document)</option>}
                {fontFamilies.map((family) => <option key={family} value={family}>{family}</option>)}
              </select>
            </label>
            <div className="inspector-field-row">
              <label className="inspector-field">
                <span>Size</span>
                <div className="unit-input"><input aria-label="Font size" type="number" min="4" max="96" step="0.5" value={Math.round(textEdit.fontSize * 10) / 10} onChange={(event) => {
                  const size = clamp(Number(event.target.value) || 11, 4, 96)
                  onTextChange({ ...textEdit, fontSize: size, lineHeight: (textEdit.lineHeight || textEdit.fontSize * 1.18) * size / textEdit.fontSize, preserveSourceMetrics: false, modified: true })
                }} /><small>pt</small></div>
              </label>
              <label className="inspector-field color-field">
                <span>Color</span>
                <input type="color" value={rgbToHex(textEdit.color)} onChange={(event) => onTextChange({ ...textEdit, color: hexToRgb(event.target.value), modified: true })} />
              </label>
            </div>
            <div className="inspector-field">
              <span>Style</span>
              <div className="segmented-control">
                <button
                  type="button"
                  className={(textEdit.fontWeight || 400) >= 600 ? 'is-active' : ''}
                  title="Bold"
                  aria-pressed={(textEdit.fontWeight || 400) >= 600}
                  onClick={() => onTextChange({ ...textEdit, fontWeight: (textEdit.fontWeight || 400) >= 600 ? 400 : 700, fontKey: undefined, fontData: undefined, preserveSourceMetrics: false, modified: true })}
                ><Bold size={15} /></button>
                <button
                  type="button"
                  className={textEdit.fontStyle === 'italic' ? 'is-active' : ''}
                  title="Italic"
                  aria-pressed={textEdit.fontStyle === 'italic'}
                  onClick={() => onTextChange({ ...textEdit, fontStyle: textEdit.fontStyle === 'italic' ? 'normal' : 'italic', fontKey: undefined, fontData: undefined, preserveSourceMetrics: false, modified: true })}
                ><Italic size={15} /></button>
              </div>
            </div>
            <div className="inspector-field">
              <span>Alignment</span>
              <div className="segmented-control">
                {([
                  ['left', AlignLeft, 'Align left'],
                  ['center', AlignCenter, 'Align center'],
                  ['right', AlignRight, 'Align right'],
                ] as const).map(([alignment, Icon, label]) => (
                  <button key={alignment} type="button" className={textEdit.align === alignment ? 'is-active' : ''} title={label} aria-label={label} aria-pressed={textEdit.align === alignment} onClick={() => onTextChange({ ...textEdit, align: alignment, modified: true })}><Icon size={15} /></button>
                ))}
              </div>
            </div>
            <div className="inspector-field-row">
              <label className="inspector-field">
                <span>Character spacing</span>
                <div className="unit-input"><input aria-label="Character spacing" type="number" min="-4" max="20" step="0.1" value={Math.round((textEdit.letterSpacing || 0) * 10) / 10} onChange={(event) => onTextChange({ ...textEdit, letterSpacing: clamp(Number(event.target.value) || 0, -4, 20), preserveSourceMetrics: false, modified: true })} /><small>pt</small></div>
              </label>
              <label className="inspector-field">
                <span>Rotation</span>
                <div className="unit-input"><input type="number" min="-180" max="180" step="1" value={Math.round((textEdit.angle || 0) * 180 / Math.PI)} onChange={(event) => onTextChange({ ...textEdit, angle: clamp(Number(event.target.value) || 0, -180, 180) * Math.PI / 180, modified: true })} /><small>°</small></div>
              </label>
            </div>
            <label className="inspector-field">
              <span>Line spacing</span>
              <div className="unit-input"><input aria-label="Line spacing" type="number" min={Math.round(textEdit.fontSize * 0.8 * 10) / 10} max="200" step="0.5" value={Math.round((textEdit.lineHeight || textEdit.fontSize * 1.18) * 10) / 10} onChange={(event) => onTextChange({ ...textEdit, lineHeight: clamp(Number(event.target.value) || textEdit.fontSize * 1.18, textEdit.fontSize * 0.8, 200), modified: true })} /><small>pt</small></div>
            </label>
          </section>
          <section className="inspector-section">
            <h3>Text box</h3>
            <label className="inspector-field">
              <span>Text fitting</span>
              <select aria-label="Text fitting" value={textEdit.textFit || (textEdit.originalText && !textEdit.originalText.includes('\n') && !textEdit.text.includes('\n') ? 'fit' : 'wrap')} onChange={(event) => onTextChange({ ...textEdit, textFit: event.target.value as 'fit' | 'wrap', modified: true })}>
                <option value="fit">Fit to one line</option>
                <option value="wrap">Wrap inside box</option>
              </select>
            </label>
            <div className="inspector-button-grid">
              <Button icon={Copy} variant="secondary" onClick={onCopySelection}>Copy</Button>
              <Button icon={ClipboardPaste} variant="secondary" disabled={!canPasteSelection} onClick={onPasteSelection}>Paste</Button>
              <Button icon={Copy} variant="secondary" onClick={onDuplicateText}>Duplicate</Button>
            </div>
            <div className="geometry-grid">
              {(['x', 'y', 'width', 'height'] as const).map((key) => (
                <label key={key}><span>{key === 'width' ? 'W' : key === 'height' ? 'H' : key.toUpperCase()}</span><input type="number" min={key === 'width' || key === 'height' ? 4 : undefined} step="1" value={Math.round(textEdit.rect[key] * 10) / 10} onChange={(event) => onTextChange({ ...textEdit, rect: { ...textEdit.rect, [key]: Math.max(key === 'width' || key === 'height' ? 4 : 0, Number(event.target.value) || 0) }, modified: true })} /></label>
              ))}
            </div>
            <p className="field-help">Drag the top grip to move or the handles to resize. Wrap keeps the font size; fit compresses a longer line. Ctrl+Enter applies your edit.</p>
          </section>
          <div className="inspector-footer">
            <Button icon={Trash2} variant="ghost" className="inspector-delete" onClick={onDeleteSelection}>Delete</Button>
            <span />
            <Button variant="ghost" onClick={onCancelSelection}>Cancel</Button>
            <Button icon={Check} variant="primary" onClick={onCommitText}>Done</Button>
          </div>
        </div>
      )}

      {objectEdit && (
        <div className="inspector-content">
          <div className="selection-summary"><span className="selection-kind"><ImagePlus size={14} /> {objectEdit.label}</span><small>Page {objectEdit.pageIndex + 1}</small></div>
          <section className="inspector-section">
            <h3>Object</h3>
            <div className="inspector-button-grid">
              <Button icon={Copy} variant="secondary" onClick={onCopySelection}>Copy</Button>
              <Button icon={ClipboardPaste} variant="secondary" disabled={!canPasteSelection} onClick={onPasteSelection}>Paste</Button>
              <Button icon={ImagePlus} variant="secondary" onClick={onReplaceImage}>Replace</Button>
              <Button icon={Copy} variant="secondary" onClick={onDuplicateObject}>Duplicate</Button>
              <Button icon={RotateCcw} variant="secondary" onClick={() => onRotateObject('left')}>Rotate left</Button>
              <Button icon={RotateCw} variant="secondary" onClick={() => onRotateObject('right')}>Rotate right</Button>
              <Button icon={FlipHorizontal2} variant="secondary" onClick={() => onFlipObject('horizontal')}>Flip horizontal</Button>
              <Button icon={FlipVertical2} variant="secondary" onClick={() => onFlipObject('vertical')}>Flip vertical</Button>
            </div>
            <label className="inspector-field range-field">
              <span>Opacity <small>{Math.round(objectEdit.opacity * 100)}%</small></span>
              <input type="range" min="5" max="100" step="1" value={Math.round(objectEdit.opacity * 100)} onChange={(event) => onObjectChange({ ...objectEdit, opacity: Number(event.target.value) / 100, modified: true })} />
            </label>
          </section>
          <section className="inspector-section">
            <h3>Position & size</h3>
            <div className="geometry-grid">
              {(['x', 'y', 'width', 'height'] as const).map((key) => (
                <label key={key}><span>{key === 'width' ? 'W' : key === 'height' ? 'H' : key.toUpperCase()}</span><input type="number" min={key === 'width' || key === 'height' ? 4 : undefined} step="1" value={Math.round(objectEdit.rect[key] * 10) / 10} onChange={(event) => onObjectChange({ ...objectEdit, rect: { ...objectEdit.rect, [key]: Math.max(key === 'width' || key === 'height' ? 4 : 0, Number(event.target.value) || 0) }, modified: true })} /></label>
              ))}
            </div>
            <p className="field-help">Drag the object or use arrow keys to move by 1 pt; hold Shift for 10 pt. Ctrl+D duplicates the selection.</p>
          </section>
          <div className="inspector-footer">
            <Button icon={Trash2} variant="ghost" className="inspector-delete" onClick={onDeleteSelection}>Delete</Button>
            <span />
            <Button variant="ghost" onClick={onCancelSelection}>Cancel</Button>
            <Button icon={Check} variant="primary" onClick={onCommitObject}>Done</Button>
          </div>
        </div>
      )}
    </aside>
  )
}
