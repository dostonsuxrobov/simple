const NATIVE_IMAGE = /\.(png|jpe?g)$/i
const CONVERTIBLE_IMAGE = /\.(gif|webp|bmp|avif|svg|tiff?|ico)$/i

export function isConvertibleImage(file: Pick<File, 'name' | 'type'>) {
  return !NATIVE_IMAGE.test(file.name) && (CONVERTIBLE_IMAGE.test(file.name) || /^image\/(gif|webp|bmp|avif|svg\+xml|tiff|x-icon)$/i.test(file.type))
}

/** Rasterise formats the PDF writer cannot embed directly into a PNG the main process accepts. */
export async function convertImageToPng(file: File): Promise<{ name: string; data: ArrayBuffer }> {
  const url = URL.createObjectURL(file)
  try {
    const image = new Image()
    image.src = url
    await image.decode()
    const width = image.naturalWidth || 1024
    const height = image.naturalHeight || 768
    const scale = Math.min(1, 8000 / Math.max(width, height))
    const canvas = document.createElement('canvas')
    canvas.width = Math.max(1, Math.round(width * scale))
    canvas.height = Math.max(1, Math.round(height * scale))
    const context = canvas.getContext('2d')
    if (!context) throw new Error('Could not convert the image.')
    context.drawImage(image, 0, 0, canvas.width, canvas.height)
    const blob = await new Promise<Blob | null>((resolve) => canvas.toBlob(resolve, 'image/png'))
    if (!blob) throw new Error(`“${file.name}” could not be converted to an image page.`)
    return { name: file.name.replace(/\.[^.]+$/, '') + '.png', data: await blob.arrayBuffer() }
  } catch (error) {
    throw new Error(error instanceof Error && error.message ? error.message : `“${file.name}” could not be read as an image.`)
  } finally {
    URL.revokeObjectURL(url)
  }
}
