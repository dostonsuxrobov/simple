// Keep complex Word layout in the compatible Office page view. The editor's
// own margin-aligned square-wrap image is a supported flowing object, so it can
// reopen directly without starting another renderer or changing pagination.
function needsOfficeLayout(xml, partName = 'word/document.xml') {
  if (/<(?:\w+:)?(?:tblpPr|txbxContent|pict)\b/.test(xml)) return true;
  const starts = xml.match(/<(?:\w+:)?anchor\b/g) || [];
  if (!starts.length) return false;
  if (partName !== 'word/document.xml') return true;
  const anchors = [...xml.matchAll(/<wp:anchor\b[^>]*>[\s\S]*?<\/wp:anchor>/g)].map(match => match[0]);
  if (anchors.length !== starts.length) return true;
  return anchors.some(anchor => !(
    !/behindDoc=["']1["']/.test(anchor)
    && /<pic:pic>/.test(anchor)
    && /<wp:wrapSquare wrapText="bothSides"\s*\/>/.test(anchor)
    && /<wp:positionH relativeFrom="margin"><wp:align>(?:left|center|right)<\/wp:align><\/wp:positionH>/.test(anchor)
    && /<wp:positionV relativeFrom="paragraph"><wp:posOffset>0<\/wp:posOffset><\/wp:positionV>/.test(anchor)
  ));
}
module.exports = { needsOfficeLayout };
