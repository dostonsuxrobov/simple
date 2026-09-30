// Read mapped Unicode characters from standard SFNT cmap tables. Only ranges
// with a nonzero glyph are returned; no font bytes or outlines enter the model.
function fontCoverage(value) {
  const bytes = Buffer.from(value);
  try {
    if (bytes.length < 12) return [];
    let cmap;
    for (let i = 0, count = bytes.readUInt16BE(4); i < Math.min(count, 256); i++) {
      const at = 12 + i * 16;
      if (bytes.toString('ascii', at, at + 4) === 'cmap') cmap = bytes.subarray(bytes.readUInt32BE(at + 8), bytes.readUInt32BE(at + 8) + bytes.readUInt32BE(at + 12));
    }
    if (!cmap) return [];
    const points = new Set();
    for (let i = 0, count = cmap.readUInt16BE(2); i < Math.min(count, 64); i++) {
      const at = 4 + i * 8, platform = cmap.readUInt16BE(at), encoding = cmap.readUInt16BE(at + 2);
      if (platform !== 0 && !(platform === 3 && [1, 10].includes(encoding))) continue;
      const offset = cmap.readUInt32BE(at + 4), format = cmap.readUInt16BE(offset);
      if (format === 12) {
        const groups = cmap.readUInt32BE(offset + 12);
        if (groups > 20000 || offset + 16 + groups * 12 > cmap.length) continue;
        for (let n = 0; n < groups; n++) {
          const pos = offset + 16 + n * 12, start = cmap.readUInt32BE(pos), end = Math.min(0x10ffff, cmap.readUInt32BE(pos + 4));
          for (let cp = start + (cmap.readUInt32BE(pos + 8) === 0 ? 1 : 0); cp <= end; cp++) points.add(cp);
        }
      } else if (format === 4) {
        const count = cmap.readUInt16BE(offset + 6) / 2, endAt = offset + 14, startAt = endAt + count * 2 + 2, deltaAt = startAt + count * 2, rangeAt = deltaAt + count * 2;
        if (count > 8192 || rangeAt + count * 2 > cmap.length) continue;
        for (let n = 0; n < count; n++) {
          const start = cmap.readUInt16BE(startAt + n * 2), end = cmap.readUInt16BE(endAt + n * 2), delta = cmap.readInt16BE(deltaAt + n * 2), range = cmap.readUInt16BE(rangeAt + n * 2);
          for (let cp = start; cp <= end && cp < 0xffff; cp++) {
            let glyph = range ? cmap.readUInt16BE(rangeAt + n * 2 + range + (cp - start) * 2) : cp;
            if (!range || glyph) glyph = (glyph + delta) & 0xffff;
            if (glyph) points.add(cp);
          }
        }
      }
    }
    const ranges = [];
    for (const cp of [...points].sort((a, b) => a - b)) {
      const last = ranges.at(-1);
      if (last && cp === last[1] + 1) last[1] = cp;
      else ranges.push([cp, cp]);
    }
    return ranges;
  } catch { return []; }
}
module.exports = { fontCoverage };
