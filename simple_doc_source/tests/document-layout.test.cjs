const test = require('node:test');
const assert = require('node:assert/strict');
const {needsOfficeLayout} = require('../electron/document-layout.cjs');
const flow = '<wp:anchor behindDoc="0"><wp:positionH relativeFrom="margin"><wp:align>left</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:wrapSquare wrapText="bothSides"/><pic:pic></pic:pic></wp:anchor>';
test('only supported square-wrap body images bypass Office page conversion', () => {
  assert.equal(needsOfficeLayout(flow), false);
  assert.equal(needsOfficeLayout(flow, 'word/header1.xml'), true);
  assert.equal(needsOfficeLayout(flow.replace('behindDoc="0"', 'behindDoc="1"')), true);
  assert.equal(needsOfficeLayout(flow.replace('<wp:posOffset>0', '<wp:posOffset>42')), true);
  assert.equal(needsOfficeLayout(flow.replace('relativeFrom="paragraph"', 'relativeFrom="page"')), true);
  assert.equal(needsOfficeLayout(flow.replace('<wp:wrapSquare wrapText="bothSides"/>', '<wp:wrapNone/>')), true);
  assert.equal(needsOfficeLayout('<w:tblpPr/>' + flow), true);
  assert.equal(needsOfficeLayout(flow.replace('</wp:anchor>', '')), true);
  assert.equal(needsOfficeLayout('<w:p><w:r><w:t>ordinary text</w:t></w:r></w:p>'), false);
});
