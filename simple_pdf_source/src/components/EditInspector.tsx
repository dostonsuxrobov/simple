import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  Check,
  ClipboardPaste,
  Copy,
  ImagePlus,
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
  textEdit, objectEdit, selectingObjectRegion, onTextChange, onObjectChange,
  onCommitText, onCommitObject, onCancelSelection, onDeleteSelection, onAddText,
  onCopySelection, onPasteSelection, canPasteSelection, onDuplicateText, onAddImage, onReplaceImage, onRotateObject, onDuplicateObject,
  onSelectObjectRegion, onClose,
}: EditInspectorProps) {
  const selected = textEdit || objectEdit

  return (
    <aside className="edit-inspector" aria-label="Edit PDF properties">
      <header className="edit-inspector-header">
        <div><strong>Edit PDF</strong><span>Direct page editing</span></div>
        <IconButton icon={X} label="Close edit mode" compact onClick={onClose} />
      </header>

      {!selected && (
        <div className="edit-inspector-empty">
          <span className="inspector-empty-icon"><MousePointer2 size={18} /></span>
          <strong>Select content on the page</strong>
          <p>Click text to place the caret inside it. Click an image to move, resize, replace, or remove it.</p>
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
              <select value={textEdit.fontFamily} onChange={(event) => onTextChange({ ...textEdit, fontFamily: event.target.value, modified: true })}>
                <option value="Segoe UI">Segoe UI</option>
                <option value="Arial">Arial</option>
                <option value="Times New Roman">Times New Roman</option>
                <option value="Georgia">Georgia</option>
                <option value="Courier New">Courier New</option>
              </select>
            </label>
            <div className="inspector-field-row">
              <label className="inspector-field">
                <span>Size</span>
                <div className="unit-input"><input type="number" min="4" max="96" step="0.5" value={Math.round(textEdit.fontSize * 10) / 10} onChange={(event) => onTextChange({ ...textEdit, fontSize: clamp(Number(event.target.value) || 11, 4, 96), modified: true })} /><small>pt</small></div>
              </label>
              <label className="inspector-field color-field">
                <span>Color</span>
                <input type="color" value={rgbToHex(textEdit.color)} onChange={(event) => onTextChange({ ...textEdit, color: hexToRgb(event.target.value), modified: true })} />
              </label>
            </div>
            <div className="inspector-field">
              <span>Alignment</span>
              <div className="segmented-control">
                {([
                  ['left', AlignLeft, 'Align left'],
                  ['center', AlignCenter, 'Align center'],
                  ['right', AlignRight, 'Align right'],
                ] as const).map(([alignment, Icon, label]) => (
                  <button key={alignment} type="button" className={textEdit.align === alignment ? 'is-active' : ''} title={label} onClick={() => onTextChange({ ...textEdit, align: alignment, modified: true })}><Icon size={15} /></button>
                ))}
              </div>
            </div>
          </section>
          <section className="inspector-section">
            <h3>Text box</h3>
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
            <p className="field-help">Drag the page border to move the box or its handles to resize. Text reflows only inside this box.</p>
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
            <p className="field-help">Drag the object directly on the page. Use the eight handles for precise resizing.</p>
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
