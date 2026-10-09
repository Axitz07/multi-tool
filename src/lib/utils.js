export function parsePageRanges(input, pageCount) {
  if (!Number.isInteger(pageCount) || pageCount < 1) throw new Error('PDF tidak memiliki halaman yang valid.')
  const value = String(input ?? '').trim()
  if (!value) return Array.from({ length: pageCount }, (_, index) => index)

  const pages = new Set()
  for (const partRaw of value.split(',')) {
    const part = partRaw.trim()
    if (!part) continue
    const match = part.match(/^(\d+)(?:\s*-\s*(\d+))?$/)
    if (!match) throw new Error(`Rentang halaman "${part}" tidak valid. Contoh: 1,3-5.`)
    const first = Number(match[1])
    const last = match[2] ? Number(match[2]) : first
    if (first < 1 || last < 1 || first > pageCount || last > pageCount) {
      throw new Error(`Halaman harus berada di antara 1 dan ${pageCount}.`)
    }
    if (last < first) throw new Error(`Rentang "${part}" terbalik. Gunakan urutan awal-akhir.`)
    for (let page = first; page <= last; page += 1) pages.add(page - 1)
  }
  if (!pages.size) throw new Error('Pilih minimal satu halaman.')
  return [...pages].sort((a, b) => a - b)
}

export function parsePageOrder(input, pageCount) {
  const value = String(input ?? '').trim()
  if (!value) return Array.from({ length: pageCount }, (_, index) => index)
  const order = value.split(',').map((part) => Number(part.trim()))
  if (order.some((page) => !Number.isInteger(page) || page < 1 || page > pageCount)) {
    throw new Error(`Urutan harus berisi nomor halaman antara 1 dan ${pageCount}.`)
  }
  if (new Set(order).size !== order.length) throw new Error('Nomor halaman tidak boleh duplikat pada urutan baru.')
  return order.map((page) => page - 1)
}

export function formatBytes(bytes) {
  if (!Number.isFinite(bytes) || bytes < 0) return '0 B'
  if (bytes < 1024) return `${bytes} B`
  const units = ['KB', 'MB', 'GB', 'TB']
  let value = bytes / 1024
  let unit = units[0]
  for (let index = 1; value >= 1024 && index < units.length - 1; index += 1) {
    value /= 1024
    unit = units[index]
  }
  return `${value.toFixed(value >= 10 ? 1 : 2)} ${unit}`
}

export function calculateGpa(rows) {
  const scale = { A: 4, AB: 3.5, B: 3, BC: 2.5, C: 2, D: 1, E: 0 }
  let weighted = 0
  let credits = 0
  for (const row of rows) {
    const sks = Number(row.credits)
    if (!Number.isFinite(sks) || sks <= 0) continue
    const grade = scale[row.grade]
    if (grade === undefined) continue
    weighted += sks * grade
    credits += sks
  }
  return { gpa: credits ? weighted / credits : 0, credits, weighted }
}

function clean(value) {
  return String(value ?? '').trim()
}

export function buildCitation(data, style = 'apa') {
  const authors = clean(data.authors).split(';').map(clean).filter(Boolean)
  const title = clean(data.title)
  const year = clean(data.year) || 'n.d.'
  const source = clean(data.source)
  const doi = clean(data.doi)
  const url = clean(data.url)
  if (!title) throw new Error('Judul sumber wajib diisi.')
  if (style === 'ieee') {
    const authorText = authors.length ? `${authors.join(', ')}, ` : ''
    const sourceText = source ? `, ${source}` : ''
    const doiText = doi ? `, doi: ${doi}` : ''
    const urlText = !doi && url ? `, [Online]. Available: ${url}` : ''
    return `${authorText}"${title},"${sourceText}, ${year}${doiText}${urlText}.`.replace(/,\s*,/g, ',').replace(/\s+\./g, '.')
  }
  let authorText = ''
  if (authors.length === 1) authorText = authors[0]
  else if (authors.length === 2) authorText = `${authors[0]}, & ${authors[1]}`
  else if (authors.length > 2) authorText = `${authors.slice(0, 19).join(', ')}, … ${authors.at(-1)}`
  const sourceText = source ? ` *${source}*.` : ''
  const doiText = doi ? ` https://doi.org/${doi.replace(/^https?:\/\/(dx\.)?doi\.org\//i, '')}` : (url ? ` ${url}` : '')
  const authorPrefix = authorText ? `${authorText} (${year}).` : `(${year}).`
  return `${authorPrefix} ${title}.${sourceText}${doiText}`.replace(/\s+/g, ' ').trim()
}

export function makeStudyNotes(rawText) {
  const text = clean(rawText).replace(/\s+/g, ' ')
  if (text.length < 40) throw new Error('Teks materi terlalu pendek untuk diringkas. Coba gunakan dokumen yang berisi lebih banyak teks.')
  const sentences = text.match(/[^.!?]+[.!?]?/g)?.map((sentence) => sentence.trim()).filter((sentence) => sentence.length > 25) ?? []
  const stop = new Set(['yang','dan','atau','dari','dengan','untuk','pada','dalam','adalah','itu','ini','sebagai','akan','dapat','oleh','the','and','for','with','that','this','from','are','was','were','have','has','not','but','into','their','they'])
  const words = text.toLowerCase().match(/[a-zA-ZÀ-ÿ]{4,}/g) ?? []
  const frequencies = new Map()
  for (const word of words) if (!stop.has(word)) frequencies.set(word, (frequencies.get(word) ?? 0) + 1)
  const keywords = [...frequencies.entries()].sort((a, b) => b[1] - a[1]).slice(0, 8).map(([word]) => word)
  const summary = sentences.length <= 5 ? sentences : [sentences[0], ...sentences.slice(1).sort((a, b) => b.length - a.length).slice(0, 4)]
  const questions = sentences.slice(0, 6).map((sentence, index) => `${index + 1}. Jelaskan gagasan utama dari pernyataan berikut: "${sentence.slice(0, 180)}${sentence.length > 180 ? '…' : ''}"`)
  return { summary, keywords, questions }
}
