"""Independent PDF/XML verification after run-editor-stress.mjs (PyMuPDF/Pillow)."""
import json
from pathlib import Path
import xml.etree.ElementTree as ET
import zipfile
import pymupdf
from PIL import Image, ImageChops

root = Path(__file__).resolve().parent.parent / 'tmp' / 'editor-stress'
before = pymupdf.open(root / 'before-reopen.pdf')
after = pymupdf.open(root / 'after-reopen.pdf')
assert len(before) == len(after) == 6
report = {'pageCounts': [len(before), len(after)], 'textEqual': [], 'pixelsEqual': [], 'changedPixels': [], 'maxTextPositionShiftPt': 0}

def spans(page):
    return [span for block in page.get_text('dict')['blocks'] if 'lines' in block for line in block['lines'] for span in line['spans']]

for first, second in zip(before, after):
    report['textEqual'].append(first.get_text() == second.get_text())
    a, b = first.get_pixmap(), second.get_pixmap()
    assert (a.width, a.height) == (b.width, b.height)
    report['pixelsEqual'].append(a.samples == b.samples)
    diff = ImageChops.difference(Image.frombytes('RGB', [a.width, a.height], a.samples), Image.frombytes('RGB', [b.width, b.height], b.samples))
    report['changedPixels'].append(sum(max(pixel) > 0 for pixel in diff.get_flattened_data()))
    x, y = spans(first), spans(second)
    assert len(x) == len(y)
    delta = max(abs(v - u) for s, t in zip(x, y) for u, v in zip(s['bbox'], t['bbox']))
    report['maxTextPositionShiftPt'] = max(report['maxTextPositionShiftPt'], delta)

text = ''.join(page.get_text() for page in after)
assert 'FOOTNOTE_CONTENT' in text
assert all(char in text for char in '中文测试日本語の文章한국어문장')
assert all(report['textEqual'])
assert report['maxTextPositionShiftPt'] < .04
assert any(span['color'] == 0x003388 for span in spans(after[0]))
with zipfile.ZipFile(root / 'created.docx') as package:
    for name in package.namelist():
        if name.endswith('.xml'):
            ET.fromstring(package.read(name))
report.update(strictXmlValid=True, footnoteRetained=True, mixedScriptsRetained=True)
(root / 'verified-roundtrip.json').write_text(json.dumps(report, indent=2), encoding='utf8')
for page, name in [(after[0], 'tables-pdf'), (after[-2], 'multilingual-pdf'), (after[-1], 'landscape-reopened-pdf')]:
    page.get_pixmap().save(root / f'{name}.png')
print(json.dumps(report))
