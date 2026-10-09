import JSZip from 'jszip'
import { fileToCanvas, canvasToBlob, saveBlob } from './pdfTools.js'

export async function convertImages(files, options = {}) {
  const outputs = []
  for (const file of files) {
    if (!file.type.startsWith('image/')) continue
    const canvas = await fileToCanvas(file, {
      maxSide: Number(options.maxSide ?? 0),
      crop: options.crop,
      grayscale: Boolean(options.grayscale),
      contrast: Number(options.contrast ?? 1),
      threshold: Boolean(options.threshold)
    })
    const mime = options.format || 'image/webp'
    const blob = await canvasToBlob(canvas, mime, Number(options.quality ?? 0.88))
    const extension = mime === 'image/jpeg' ? 'jpg' : mime === 'image/png' ? 'png' : 'webp'
    const baseName = file.name.replace(/\.[^.]+$/, '')
    outputs.push({ blob, filename: `${baseName}-toolkitini.${extension}`, size: blob.size })
  }
  if (!outputs.length) throw new Error('Pilih minimal satu file gambar.')
  return outputs
}

export async function saveMany(outputs, archiveName = 'toolkitini-images.zip') {
  if (!outputs.length) throw new Error('Tidak ada gambar hasil yang dapat disimpan.')
  const zip = new JSZip()
  for (const output of outputs) zip.file(output.filename, await output.blob.arrayBuffer())
  const archive = await zip.generateAsync({ type: 'blob' })
  saveBlob(archive, archiveName)
  return { count: outputs.length, size: archive.size, blob: archive }
}
