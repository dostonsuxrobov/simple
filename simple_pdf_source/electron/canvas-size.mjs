export function boundedCanvasSize(width, height, desiredScale, maxPixels = 24_000_000, maxEdge = 16_384) {
  if (![width, height, desiredScale].every(n => Number.isFinite(n) && n > 0)) throw new Error('Invalid page canvas dimensions.')
  const scale = Math.min(desiredScale, Math.sqrt(maxPixels / width / height), maxEdge / width, maxEdge / height)
  return { width: Math.max(1, Math.floor(width * scale)), height: Math.max(1, Math.floor(height * scale)) }
}
