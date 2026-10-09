import { PDFDocument, StandardFonts, rgb, degrees } from 'pdf-lib'
import * as pdfjsLib from 'pdfjs-dist'
import pdfWorkerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url'
import JSZip from 'jszip'
import { Document, Packer, Paragraph } from 'docx'
import { parsePageRanges, parsePageOrder } from './utils.js'

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorkerUrl

export function saveBlob(blob, filename) {
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = filename
  document.body.appendChild(anchor)
  anchor.click()
  anchor.remove()
  setTimeout(() => URL.revokeObjectURL(url), 60_000)
}

export function fileToCanvas(file, options = {}) {
  return createImageBitmap(file).then((bitmap) => {
    const crop = options.crop ?? { x: 0, y: 0, width: 100, height: 100 }
    const sx = Math.round(bitmap.width * Number(crop.x ?? 0) / 100)
    const sy = Math.round(bitmap.height * Number(crop.y ?? 0) / 100)
    const sw = Math.max(1, Math.min(bitmap.width - sx, Math.round(bitmap.width * Number(crop.width ?? 100) / 100)))
    const sh = Math.max(1, Math.min(bitmap.height - sy, Math.round(bitmap.height * Number(crop.height ?? 100) / 100)))
    const maxSide = Number(options.maxSide ?? 0)
    const factor = maxSide > 0 ? Math.min(1, maxSide / Math.max(sw, sh)) : 1
    const canvas = document.createElement('canvas')
    canvas.width = Math.max(1, Math.round(sw * factor))
    canvas.height = Math.max(1, Math.round(sh * factor))
    const context = canvas.getContext('2d', { willReadFrequently: true })
    if (!context) throw new Error('Canvas tidak dapat dibuat pada perangkat ini.')
    context.drawImage(bitmap, sx, sy, sw, sh, 0, 0, canvas.width, canvas.height)
    bitmap.close()
    if (options.grayscale || options.contrast !== undefined || options.threshold) {
      const imageData = context.getImageData(0, 0, canvas.width, canvas.height)
      const contrast = Number(options.contrast ?? 1)
      for (let index = 0; index < imageData.data.length; index += 4) {
        let r = imageData.data[index]
        let g = imageData.data[index + 1]
        let b = imageData.data[index + 2]
        if (options.grayscale || options.threshold) {
          const gray = 0.299 * r + 0.587 * g + 0.114 * b
          r = g = b = gray
        }
        r = Math.max(0, Math.min(255, (r - 128) * contrast + 128))
        g = Math.max(0, Math.min(255, (g - 128) * contrast + 128))
        b = Math.max(0, Math.min(255, (b - 128) * contrast + 128))
        if (options.threshold) r = g = b = (r > 160 ? 255 : 0)
        imageData.data[index] = r
        imageData.data[index + 1] = g
        imageData.data[index + 2] = b
      }
      context.putImageData(imageData, 0, 0)
    }
    return canvas
  })
}

export function canvasToBlob(canvas, type = 'image/png', quality = 0.9) {
  return new Promise((resolve, reject) => canvas.toBlob((blob) => blob ? resolve(blob) : reject(new Error('Gagal membuat file gambar.')), type, quality))
}

export async function imageFilesToPdf(files) {
  const pdf = await PDFDocument.create()
  for (const file of files) {
    if (!file.type.startsWith('image/')) continue
    const canvas = await fileToCanvas(file, { maxSide: 2800 })
    const blob = await canvasToBlob(canvas, 'image/png')
    const bytes = await blob.arrayBuffer()
    const image = await pdf.embedPng(bytes)
    // Put each photo on an A4 page, preserving its aspect ratio and leaving a clean margin.
    const portrait = { width: 595.28, height: 841.89 }
    const landscape = { width: 841.89, height: 595.28 }
    const pageSize = image.width > image.height ? landscape : portrait
    const margin = 28
    const scale = Math.min((pageSize.width - margin * 2) / image.width, (pageSize.height - margin * 2) / image.height)
    const width = image.width * scale
    const height = image.height * scale
    const page = pdf.addPage([pageSize.width, pageSize.height])
    page.drawImage(image, { x: (pageSize.width - width) / 2, y: (pageSize.height - height) / 2, width, height })
  }
  if (!pdf.getPageCount()) throw new Error('Pilih minimal satu gambar yang didukung.')
  return new Blob([await pdf.save()], { type: 'application/pdf' })
}

async function loadPdfJs(file) {
  const data = new Uint8Array(await file.arrayBuffer())
  return pdfjsLib.getDocument({ data }).promise
}

export async function openPdfDocument(file) {
  return loadPdfJs(file)
}

export async function getPdfText(file) {
  const pdf = await loadPdfJs(file)
  const pages = []
  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber)
    const content = await page.getTextContent()
    pages.push(content.items.map((item) => item.str).join(' '))
  }
  return pages.join('\n\n')
}

export async function renderPdfPageToCanvas(pdf, pageNumber, scale = 1.5) {
  const page = await pdf.getPage(pageNumber)
  const viewport = page.getViewport({ scale })
  const canvas = document.createElement('canvas')
  canvas.width = Math.ceil(viewport.width)
  canvas.height = Math.ceil(viewport.height)
  const context = canvas.getContext('2d', { alpha: false })
  if (!context) throw new Error('Canvas tidak tersedia.')
  context.fillStyle = '#ffffff'
  context.fillRect(0, 0, canvas.width, canvas.height)
  await page.render({ canvasContext: context, viewport }).promise
  return { canvas, page, viewport }
}

export async function mergePdfs(files) {
  const output = await PDFDocument.create()
  for (const file of files) {
    if (file.type !== 'application/pdf' && !file.name.toLowerCase().endsWith('.pdf')) continue
    const source = await PDFDocument.load(await file.arrayBuffer())
    const pages = await output.copyPages(source, source.getPageIndices())
    pages.forEach((page) => output.addPage(page))
  }
  if (!output.getPageCount()) throw new Error('Pilih minimal satu file PDF.')
  return new Blob([await output.save()], { type: 'application/pdf' })
}

export async function extractPdfPages(file, range) {
  const source = await PDFDocument.load(await file.arrayBuffer())
  const indices = parsePageRanges(range, source.getPageCount())
  const output = await PDFDocument.create()
  const pages = await output.copyPages(source, indices)
  pages.forEach((page) => output.addPage(page))
  return new Blob([await output.save()], { type: 'application/pdf' })
}

export async function organizePdf(file, orderText, rotationText, angle) {
  const source = await PDFDocument.load(await file.arrayBuffer())
  const order = parsePageOrder(orderText, source.getPageCount())
  const rotationSet = new Set(String(rotationText ?? '').trim() ? parsePageRanges(rotationText, source.getPageCount()) : [])
  const output = await PDFDocument.create()
  const pages = await output.copyPages(source, order)
  pages.forEach((page, index) => {
    const sourceIndex = order[index]
    if (rotationSet.has(sourceIndex)) page.setRotation(degrees(page.getRotation().angle + Number(angle)))
    output.addPage(page)
  })
  return new Blob([await output.save()], { type: 'application/pdf' })
}

export async function compressPdf(file, scale = 0.7, quality = 0.65) {
  const source = await loadPdfJs(file)
  const output = await PDFDocument.create()
  for (let pageNumber = 1; pageNumber <= source.numPages; pageNumber += 1) {
    const basePage = await source.getPage(pageNumber)
    const baseViewport = basePage.getViewport({ scale: 1 })
    const rendered = await renderPdfPageToCanvas(source, pageNumber, Number(scale))
    const jpg = await canvasToBlob(rendered.canvas, 'image/jpeg', Number(quality))
    const embedded = await output.embedJpg(await jpg.arrayBuffer())
    const page = output.addPage([baseViewport.width, baseViewport.height])
    page.drawImage(embedded, { x: 0, y: 0, width: baseViewport.width, height: baseViewport.height })
  }
  return new Blob([await output.save({ useObjectStreams: true })], { type: 'application/pdf' })
}

export async function pdfToImagesZip(file, scale = 1.5, type = 'image/png') {
  const pdf = await loadPdfJs(file)
  const zip = new JSZip()
  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const { canvas } = await renderPdfPageToCanvas(pdf, pageNumber, Number(scale))
    const blob = await canvasToBlob(canvas, type, 0.92)
    const extension = type === 'image/jpeg' ? 'jpg' : 'png'
    zip.file(`halaman-${String(pageNumber).padStart(3, '0')}.${extension}`, await blob.arrayBuffer())
  }
  return zip.generateAsync({ type: 'blob' })
}

export async function pdfToDocx(file) {
  const text = await getPdfText(file)
  if (!text.trim()) throw new Error('PDF tidak memiliki teks yang dapat diekstrak. Untuk PDF hasil scan, jalankan OCR terlebih dahulu.')
  const paragraphs = text.split(/\n{1,}/).map((line) => new Paragraph({ text: line.trim() })).filter((paragraph) => paragraph)
  const document = new Document({ sections: [{ properties: {}, children: paragraphs.length ? paragraphs : [new Paragraph({ text })] }] })
  return Packer.toBlob(document)
}

export async function addPdfAnnotations(file, options) {
  const pdf = await PDFDocument.load(await file.arrayBuffer())
  const font = await pdf.embedFont(StandardFonts.Helvetica)
  const indices = parsePageRanges(options.pages, pdf.getPageCount())
  for (const index of indices) {
    const page = pdf.getPage(index)
    if (options.mode === 'highlight') {
      page.drawRectangle({ x: 36, y: page.getHeight() * 0.35, width: page.getWidth() - 72, height: 26, color: rgb(1, 0.88, 0.25), opacity: 0.35, borderOpacity: 0 })
      if (options.text) page.drawText(safePdfText(options.text), { x: 42, y: page.getHeight() * 0.35 + 8, size: 10, font, color: rgb(0.22, 0.19, 0.08), maxWidth: page.getWidth() - 90 })
    } else {
      page.drawText(safePdfText(options.text || 'Catatan'), { x: 36, y: 52, size: 12, font, color: rgb(0.18, 0.29, 0.48), maxWidth: page.getWidth() - 72, lineHeight: 16 })
    }
  }
  return new Blob([await pdf.save()], { type: 'application/pdf' })
}

function safePdfText(value) {
  return String(value ?? '').replace(/[^\x20-\xFF\n\r\t]/g, '?').slice(0, 2000)
}

export async function addPdfWatermarkAndNumbers(file, options) {
  const pdf = await PDFDocument.load(await file.arrayBuffer())
  const font = await pdf.embedFont(StandardFonts.Helvetica)
  const text = safePdfText(options.watermark ?? '')
  const start = Math.max(1, Number(options.start ?? 1))
  pdf.getPages().forEach((page, index) => {
    const { width, height } = page.getSize()
    if (text) page.drawText(text, { x: width * 0.18, y: height * 0.5, size: Math.min(38, width / 12), font, color: rgb(0.45, 0.48, 0.55), opacity: 0.22, rotate: degrees(-28), maxWidth: width * 0.75 })
    if (options.numbering) page.drawText(String(start + index), { x: width / 2 - 8, y: 22, size: 10, font, color: rgb(0.25, 0.27, 0.31) })
  })
  return new Blob([await pdf.save()], { type: 'application/pdf' })
}
