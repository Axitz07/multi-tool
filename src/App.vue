<script setup>
import { computed, onBeforeUnmount, reactive, ref, watch } from 'vue'
import {
  ArrowDownToLine, ArrowLeftRight, BookOpen, BrainCircuit, CalendarClock, Check, CheckCircle2,
  ChevronDown, CircleHelp, Clock3, FileArchive, FileImage, FilePlus2, FileScan, FileText,
  Files, Filter, GraduationCap, Highlighter, Image, Layers3, Menu, Moon, Plus, Quote,
  RotateCw, ScanText, Search, Settings2, ShieldCheck, Sparkles, Stamp, Trash2, Upload,
  X, Zap
} from 'lucide-vue-next'
import { createWorker } from 'tesseract.js'
import { Document, Packer, Paragraph } from 'docx'
import {
  addPdfAnnotations, addPdfWatermarkAndNumbers, compressPdf, extractPdfPages,
  getPdfText, imageFilesToPdf, mergePdfs, openPdfDocument, organizePdf,
  pdfToDocx, pdfToImagesZip, renderPdfPageToCanvas, saveBlob
} from './lib/pdfTools.js'
import { convertImages, saveMany } from './lib/imageTools.js'
import { buildCitation, calculateGpa, formatBytes, makeStudyNotes } from './lib/utils.js'

const iconMap = { FileImage, ScanText, FileArchive, Files, Layers3, FileText, Image, FileScan, Highlighter, Stamp, GraduationCap, CalendarClock, Quote, Clock3, BrainCircuit, ArrowLeftRight, RotateCw }
const toolCatalog = [
  { id: 'image-to-pdf', title: 'Foto ke PDF', description: 'Gabungkan foto tugas menjadi dokumen PDF.', group: 'Dokumen', icon: 'FileImage', tag: 'Sering dipakai', accept: 'image/*', multiple: true, action: 'Buat PDF', detail: 'Pilih foto satu atau beberapa. Urutan file yang dipilih menjadi urutan halaman PDF.' },
  { id: 'ocr', title: 'OCR · Foto/PDF ke Teks', description: 'Ekstrak teks cetak dari scan atau foto.', group: 'Dokumen', icon: 'ScanText', tag: 'Produktif', accept: 'image/*,.pdf', multiple: true, action: 'Ekstrak teks', detail: 'OCR memakai model bahasa Indonesia dan Inggris. Pemakaian pertama perlu mengunduh data bahasa.' },
  { id: 'pdf-compress', title: 'Kompres PDF', description: 'Perkecil file sebelum diunggah ke portal kampus.', group: 'Dokumen', icon: 'FileArchive', accept: '.pdf,application/pdf', multiple: false, action: 'Kompres PDF', detail: 'Mode ini merender halaman menjadi gambar. Ukuran bisa turun, tetapi teks tidak lagi bisa diseleksi dan kualitas bisa berubah.' },
  { id: 'pdf-merge', title: 'Gabungkan PDF', description: 'Satukan cover, isi laporan, dan lampiran.', group: 'Dokumen', icon: 'Files', accept: '.pdf,application/pdf', multiple: true, action: 'Gabungkan file', detail: 'Pilih PDF sesuai urutan yang diinginkan. File sumber tidak diubah.' },
  { id: 'pdf-extract', title: 'Split & Extract Pages', description: 'Ambil halaman tertentu dari PDF tebal.', group: 'Dokumen', icon: 'Layers3', accept: '.pdf,application/pdf', multiple: false, action: 'Ekstrak halaman', detail: 'Contoh format: 1,3-5,8. Nomor halaman dimulai dari 1.' },
  { id: 'pdf-organize', title: 'Atur Halaman PDF', description: 'Susun ulang, hapus, atau putar halaman.', group: 'Dokumen', icon: 'RotateCw', accept: '.pdf,application/pdf', multiple: false, action: 'Simpan susunan baru', detail: 'Masukkan nomor halaman sesuai urutan baru. Halaman yang tidak dicantumkan akan dihapus.' },
  { id: 'pdf-to-docx', title: 'PDF ke Word / TXT', description: 'Ambil teks PDF atau ekspor ke DOCX.', group: 'Dokumen', icon: 'FileText', accept: '.pdf,application/pdf', multiple: false, action: 'Konversi dokumen', detail: 'DOCX berisi teks hasil ekstraksi, bukan salinan tata letak PDF. PDF scan perlu OCR terlebih dahulu.' },
  { id: 'pdf-to-images', title: 'PDF ke Gambar', description: 'Ekspor halaman PDF menjadi PNG/JPG.', group: 'Dokumen', icon: 'Image', accept: '.pdf,application/pdf', multiple: false, action: 'Ekspor ke ZIP', detail: 'Setiap halaman disimpan sebagai gambar dan dikemas menjadi ZIP.' },
  { id: 'image-tools', title: 'Konverter & Resize Gambar', description: 'Ubah JPG, PNG, WebP, dimensi, dan kualitas.', group: 'Gambar', icon: 'ArrowLeftRight', accept: 'image/*', multiple: true, action: 'Proses gambar', detail: 'Bisa memproses beberapa gambar sekaligus. Dimensi maksimum 0 berarti ukuran asli.' },
  { id: 'scanner', title: 'Smart Document Scanner', description: 'Crop, luruskan area, grayscale, dan tingkatkan kontras.', group: 'Gambar', icon: 'FileScan', accept: 'image/*', multiple: true, action: 'Bersihkan scan', detail: 'Versi awal menyediakan crop manual dan peningkatan kontras. Deteksi tepi otomatis belum disertakan.' },
  { id: 'pdf-annotate', title: 'Anotasi PDF', description: 'Tambahkan catatan atau highlight ke halaman.', group: 'Dokumen', icon: 'Highlighter', accept: '.pdf,application/pdf', multiple: false, action: 'Tambahkan anotasi', detail: 'Anotasi disematkan ke file PDF baru. File asli tetap aman.' },
  { id: 'pdf-watermark', title: 'Nomor Halaman & Watermark', description: 'Tambahkan nomor halaman dan teks watermark.', group: 'Dokumen', icon: 'Stamp', accept: '.pdf,application/pdf', multiple: false, action: 'Terapkan ke PDF', detail: 'Nomor halaman berada di bagian bawah tengah. Watermark menggunakan teks diagonal tipis.' },
  { id: 'gpa', title: 'Kalkulator IPS / IPK', description: 'Hitung nilai berbobot SKS dan simulasikan target.', group: 'Akademik', icon: 'GraduationCap', noFiles: true, action: 'Hitung IPK' },
  { id: 'assignments', title: 'Tugas & Deadline', description: 'Catat tugas, mata kuliah, prioritas, dan tenggat.', group: 'Akademik', icon: 'CalendarClock', noFiles: true, action: 'Tambah tugas' },
  { id: 'citation', title: 'Generator Daftar Pustaka', description: 'Susun referensi dasar format APA 7 dan IEEE.', group: 'Akademik', icon: 'Quote', noFiles: true, action: 'Buat sitasi' },
  { id: 'pomodoro', title: 'Focus Timer', description: 'Timer belajar dan istirahat ala Pomodoro.', group: 'Produktivitas', icon: 'Clock3', noFiles: true, action: 'Mulai fokus' },
  { id: 'study-notes', title: 'Ringkas Materi & Kuis', description: 'Susun catatan dan pertanyaan latihan dari teks.', group: 'Akademik', icon: 'BrainCircuit', accept: 'image/*,.pdf', multiple: false, action: 'Buat catatan', detail: 'Fitur ini berjalan dengan ekstraksi teks dan heuristik lokal, bukan model AI generatif. Hasil harus ditinjau ulang.' }
]

const groups = ['Semua alat', 'Dokumen', 'Gambar', 'Akademik', 'Produktivitas']
const currentToolId = ref('home')
const selectedGroup = ref('Semua alat')
const searchQuery = ref('')
const selectedFiles = ref([])
const busy = ref(false)
const status = ref('')
const statusType = ref('info')
const dropActive = ref(false)
const lastResult = ref('')
const fileInputKey = ref(0)

const currentTool = computed(() => toolCatalog.find((tool) => tool.id === currentToolId.value))
const filteredTools = computed(() => toolCatalog.filter((tool) => {
  const matchesGroup = selectedGroup.value === 'Semua alat' || tool.group === selectedGroup.value
  const query = searchQuery.value.trim().toLowerCase()
  const matchesSearch = !query || `${tool.title} ${tool.description} ${tool.group}`.toLowerCase().includes(query)
  return matchesGroup && matchesSearch
}))
const homeTools = computed(() => toolCatalog.filter((tool) => ['image-to-pdf','ocr','pdf-compress','pdf-merge','gpa','assignments'].includes(tool.id)))
const fileTotal = computed(() => selectedFiles.value.reduce((sum, file) => sum + file.size, 0))
const requiresFiles = computed(() => currentTool.value && !currentTool.value.noFiles)

function toast(message, type = 'info') {
  status.value = message
  statusType.value = type
  if (typeof window !== 'undefined') {
    window.clearTimeout(toast.timeout)
    toast.timeout = window.setTimeout(() => { status.value = '' }, 6500)
  }
}
function openTool(id) {
  currentToolId.value = id
  selectedFiles.value = []
  lastResult.value = ''
  status.value = ''
  fileInputKey.value += 1
  if (id !== 'home') selectedGroup.value = 'Semua alat'
  window.scrollTo({ top: 0, behavior: 'smooth' })
}
function chooseGroup(group) {
  selectedGroup.value = group
  currentToolId.value = 'home'
}
function addFiles(event) {
  selectedFiles.value = Array.from(event.target.files ?? [])
  lastResult.value = ''
  if (selectedFiles.value.length) toast(`${selectedFiles.value.length} file dipilih · ${formatBytes(fileTotal.value)}`)
}
function onDrop(event) {
  dropActive.value = false
  const dropped = Array.from(event.dataTransfer?.files ?? [])
  if (dropped.length) {
    selectedFiles.value = dropped
    toast(`${dropped.length} file siap diproses.`)
  }
}
function removeFile(index) {
  selectedFiles.value.splice(index, 1)
}
function deliver(blob, filename, message = 'File berhasil dibuat.') {
  saveBlob(blob, filename)
  lastResult.value = `${filename} · ${formatBytes(blob.size)}`
  toast(`${message} ${lastResult.value}`, 'success')
}
function textBlob(text, type = 'text/plain;charset=utf-8') {
  return new Blob([text], { type })
}
function safeBaseName(file) {
  return file.name.replace(/\.[^.]+$/, '').replace(/[^\p{L}\p{N}_-]+/gu, '-').slice(0, 70) || 'toolkitini'
}

// PDF and image tool options
const compressScale = ref(0.7)
const compressQuality = ref(0.65)
const pageRange = ref('1-3')
const pageOrder = ref('1,2,3')
const rotatePages = ref('1')
const rotateAngle = ref('90')
const pdfOutputFormat = ref('txt')
const imageFormat = ref('image/webp')
const imageMaxSide = ref(1600)
const imageQuality = ref(0.88)
const scanCrop = reactive({ x: 0, y: 0, width: 100, height: 100 })
const scanGrayscale = ref(true)
const scanThreshold = ref(false)
const scanContrast = ref(1.25)
const annotationMode = ref('highlight')
const annotationText = ref('Materi penting')
const annotationPages = ref('')
const watermarkText = ref('ToolkitIni')
const addPageNumbers = ref(true)
const numberStart = ref(1)
const watermarkPages = ref('')

async function ocrFile(file) {
  const worker = await createWorker('ind+eng')
  const output = []
  try {
    if (file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf')) {
      const pdf = await openPdfDocument(file)
      for (let page = 1; page <= pdf.numPages; page += 1) {
        toast(`OCR halaman ${page} dari ${pdf.numPages}...`)
        const rendered = await renderPdfPageToCanvas(pdf, page, 1.65)
        const result = await worker.recognize(rendered.canvas)
        output.push(`--- Halaman ${page} ---\n${result.data.text.trim()}`)
      }
    } else {
      const result = await worker.recognize(file)
      output.push(result.data.text.trim())
    }
  } finally {
    await worker.terminate()
  }
  return output.join('\n\n')
}

async function readableText(file) {
  const text = await getPdfText(file)
  if (text.trim().length > 30) return text
  return ocrFile(file)
}

async function runCurrentTool() {
  if (!currentTool.value) return
  const id = currentToolId.value
  if (requiresFiles.value && !selectedFiles.value.length) {
    toast('Pilih file terlebih dahulu. Aplikasi belum bisa membaca pikiran, syukurlah.', 'error')
    return
  }
  busy.value = true
  status.value = ''
  lastResult.value = ''
  try {
    switch (id) {
      case 'image-to-pdf': {
        const blob = await imageFilesToPdf(selectedFiles.value)
        deliver(blob, 'toolkitini-foto.pdf', 'PDF berhasil dibuat.')
        break
      }
      case 'ocr': {
        const parts = []
        for (const file of selectedFiles.value) {
          const text = file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf') ? await readableText(file) : await ocrFile(file)
          parts.push(`===== ${file.name} =====\n${text}`)
        }
        const combined = parts.join('\n\n')
        deliver(textBlob(combined), 'hasil-ocr.txt', 'Teks hasil OCR berhasil diekspor.')
        break
      }
      case 'pdf-compress': {
        const file = selectedFiles.value[0]
        const blob = await compressPdf(file, compressScale.value, compressQuality.value)
        const change = file.size ? ((file.size - blob.size) / file.size) * 100 : 0
        const comparison = change > 0 ? `Ukuran turun ${change.toFixed(1)}%.` : change < 0 ? `Ukuran naik ${Math.abs(change).toFixed(1)}%; coba kualitas yang lebih rendah.` : 'Ukuran file hampir tidak berubah.'
        deliver(blob, `${safeBaseName(file)}-compressed.pdf`, `Proses selesai. ${comparison} Karena halaman dirasterisasi, periksa ketajaman teks sebelum mengumpulkan.`)
        break
      }
      case 'pdf-merge': {
        const blob = await mergePdfs(selectedFiles.value)
        deliver(blob, 'toolkitini-gabungan.pdf', 'PDF berhasil digabung.')
        break
      }
      case 'pdf-extract': {
        const file = selectedFiles.value[0]
        const blob = await extractPdfPages(file, pageRange.value)
        deliver(blob, `${safeBaseName(file)}-halaman-terpilih.pdf`, 'Halaman berhasil diekstrak.')
        break
      }
      case 'pdf-organize': {
        const file = selectedFiles.value[0]
        const blob = await organizePdf(file, pageOrder.value, rotatePages.value, rotateAngle.value)
        deliver(blob, `${safeBaseName(file)}-diatur.pdf`, 'Susunan halaman baru berhasil dibuat.')
        break
      }
      case 'pdf-to-docx': {
        const file = selectedFiles.value[0]
        if (pdfOutputFormat.value === 'txt') deliver(textBlob(await getPdfText(file)), `${safeBaseName(file)}.txt`, 'Teks PDF berhasil diekstrak.')
        else deliver(await pdfToDocx(file), `${safeBaseName(file)}.docx`, 'Dokumen Word berhasil dibuat dengan teks hasil ekstraksi.')
        break
      }
      case 'pdf-to-images': {
        const file = selectedFiles.value[0]
        const blob = await pdfToImagesZip(file, 1.5, imageFormat.value === 'image/jpeg' ? 'image/jpeg' : 'image/png')
        deliver(blob, `${safeBaseName(file)}-halaman.zip`, 'Halaman PDF berhasil diekspor menjadi gambar.')
        break
      }
      case 'image-tools': {
        const outputs = await convertImages(selectedFiles.value, { format: imageFormat.value, maxSide: imageMaxSide.value, quality: imageQuality.value })
        const archive = await saveMany(outputs, 'toolkitini-images.zip')
        lastResult.value = `${archive.count} gambar diproses · ZIP ${formatBytes(archive.size)}`
        toast(`Gambar berhasil diproses dan dikemas dalam ZIP. ${lastResult.value}`, 'success')
        break
      }
      case 'scanner': {
        const outputs = await convertImages(selectedFiles.value, { format: 'image/jpeg', quality: 0.92, maxSide: 2800, crop: { ...scanCrop }, grayscale: scanGrayscale.value, threshold: scanThreshold.value, contrast: scanContrast.value })
        const archive = await saveMany(outputs, 'toolkitini-hasil-scan.zip')
        lastResult.value = `${archive.count} hasil scan · ZIP ${formatBytes(archive.size)}`
        toast(`Hasil scan berhasil dikemas dalam ZIP. ${lastResult.value}`, 'success')
        break
      }
      case 'pdf-annotate': {
        const file = selectedFiles.value[0]
        const blob = await addPdfAnnotations(file, { mode: annotationMode.value, text: annotationText.value, pages: annotationPages.value })
        deliver(blob, `${safeBaseName(file)}-anotasi.pdf`, 'Anotasi berhasil ditambahkan.')
        break
      }
      case 'pdf-watermark': {
        const file = selectedFiles.value[0]
        const blob = await addPdfWatermarkAndNumbers(file, { watermark: watermarkText.value, numbering: addPageNumbers.value, start: numberStart.value })
        deliver(blob, `${safeBaseName(file)}-watermark.pdf`, 'Watermark dan nomor halaman berhasil diterapkan.')
        break
      }
      case 'gpa': {
        const result = calculateGpa(courseRows.value)
        if (!result.credits) throw new Error('Isi setidaknya satu mata kuliah dengan SKS dan nilai yang valid.')
        lastResult.value = `IPS simulasi ${result.gpa.toFixed(2)} dari ${result.credits} SKS.`
        toast(lastResult.value, 'success')
        break
      }
      case 'assignments': {
        addAssignment()
        break
      }
      case 'citation': {
        const citation = buildCitation(citationData, citationStyle.value)
        citationResult.value = citation
        lastResult.value = 'Sitasi selesai dibuat. Periksa ejaan dan metadata sumber.'
        toast('Sitasi berhasil dibuat.', 'success')
        break
      }
      case 'pomodoro': {
        toggleTimer()
        break
      }
      case 'study-notes': {
        const file = selectedFiles.value[0]
        const text = file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf') ? await readableText(file) : await ocrFile(file)
        studyNotes.value = makeStudyNotes(text)
        lastResult.value = 'Catatan dan pertanyaan latihan dibuat dari hasil ekstraksi teks.'
        toast('Draf catatan berhasil dibuat. Tinjau ulang sebelum digunakan belajar.', 'success')
        break
      }
      default: break
    }
  } catch (error) {
    console.error(error)
    toast(error instanceof Error ? error.message : 'Terjadi kesalahan saat memproses file.', 'error')
  } finally {
    busy.value = false
  }
}

// GPA calculator
const gradeOptions = ['A', 'AB', 'B', 'BC', 'C', 'D', 'E']
const courseRows = ref([{ name: 'Pengantar Sistem Informasi', credits: 3, grade: 'A' }, { name: 'Matematika', credits: 3, grade: 'B' }])
const gpaResult = computed(() => calculateGpa(courseRows.value))
function addCourse() { courseRows.value.push({ name: '', credits: 3, grade: 'A' }) }
function removeCourse(index) { courseRows.value.splice(index, 1) }

// Assignment tracker, saved locally in this browser/device
const assignments = ref([])
const assignmentForm = reactive({ title: '', course: '', due: '', priority: 'Sedang' })
function loadAssignments() {
  try { assignments.value = JSON.parse(localStorage.getItem('toolkitini-assignments') || '[]') }
  catch { assignments.value = [] }
}
watch(assignments, (value) => localStorage.setItem('toolkitini-assignments', JSON.stringify(value)), { deep: true })
function addAssignment() {
  if (!assignmentForm.title.trim()) throw new Error('Nama tugas wajib diisi.')
  assignments.value.unshift({ id: crypto.randomUUID?.() ?? `${Date.now()}-${Math.random()}`, title: assignmentForm.title.trim(), course: assignmentForm.course.trim() || 'Umum', due: assignmentForm.due, priority: assignmentForm.priority, done: false, createdAt: new Date().toISOString() })
  assignmentForm.title = ''
  assignmentForm.course = ''
  assignmentForm.due = ''
  assignmentForm.priority = 'Sedang'
  lastResult.value = 'Tugas berhasil disimpan di perangkat ini.'
  toast(lastResult.value, 'success')
}
function removeAssignment(id) { assignments.value = assignments.value.filter((task) => task.id !== id) }
const sortedAssignments = computed(() => [...assignments.value].sort((a, b) => Number(a.done) - Number(b.done) || (a.due || '9999').localeCompare(b.due || '9999')))
const openAssignments = computed(() => assignments.value.filter((task) => !task.done).length)
function daysUntil(date) {
  if (!date) return 'Tanpa tenggat'
  const due = new Date(`${date}T23:59:59`)
  const today = new Date(); today.setHours(0, 0, 0, 0)
  const days = Math.ceil((due - today) / 86400000)
  if (days < 0) return `Terlambat ${Math.abs(days)} hari`
  if (days === 0) return 'Hari ini'
  if (days === 1) return 'Besok'
  return `${days} hari lagi`
}

// Citation generator
const citationStyle = ref('apa')
const citationData = reactive({ authors: '', title: '', year: String(new Date().getFullYear()), source: '', doi: '', url: '' })
const citationResult = ref('')
async function copyText(text) {
  try { await navigator.clipboard.writeText(text); toast('Teks disalin ke clipboard.', 'success') }
  catch { toast('Clipboard tidak tersedia. Pilih lalu salin teks secara manual.', 'error') }
}

// Local extractive notes and quiz prompts
const studyNotes = ref(null)

// Focus timer
const focusLength = ref(25)
const breakLength = ref(5)
const timerSeconds = ref(25 * 60)
const timerMode = ref('focus')
const timerRunning = ref(false)
let timerInterval = null
const timerLabel = computed(() => `${String(Math.floor(timerSeconds.value / 60)).padStart(2, '0')}:${String(timerSeconds.value % 60).padStart(2, '0')}`)
const timerProgress = computed(() => {
  const total = (timerMode.value === 'focus' ? focusLength.value : breakLength.value) * 60
  return total ? Math.max(0, Math.min(100, ((total - timerSeconds.value) / total) * 100)) : 0
})
function stopTimerInterval() { if (timerInterval) clearInterval(timerInterval); timerInterval = null }
function resetTimer() { stopTimerInterval(); timerRunning.value = false; timerMode.value = 'focus'; timerSeconds.value = focusLength.value * 60 }
function switchTimerMode() {
  timerMode.value = timerMode.value === 'focus' ? 'break' : 'focus'
  timerSeconds.value = (timerMode.value === 'focus' ? focusLength.value : breakLength.value) * 60
  toast(timerMode.value === 'break' ? 'Sesi fokus selesai. Waktunya istirahat.' : 'Istirahat selesai. Kembali fokus.', 'success')
}
function toggleTimer() {
  if (timerRunning.value) { stopTimerInterval(); timerRunning.value = false; return }
  if (timerSeconds.value <= 0) timerSeconds.value = (timerMode.value === 'focus' ? focusLength.value : breakLength.value) * 60
  timerRunning.value = true
  timerInterval = setInterval(() => {
    if (timerSeconds.value > 0) timerSeconds.value -= 1
    else { stopTimerInterval(); timerRunning.value = false; switchTimerMode() }
  }, 1000)
}
watch([focusLength, breakLength], () => { if (!timerRunning.value) resetTimer() })
onBeforeUnmount(() => stopTimerInterval())
loadAssignments()
</script>

<template>
  <div class="app-shell">
    <aside class="sidebar">
      <button class="brand" @click="openTool('home')" aria-label="Kembali ke beranda">
        <span class="brand-mark"><Zap :size="21" :stroke-width="2.4" /></span>
        <span class="brand-copy"><strong>ToolkitIni</strong><small>STUDENT TOOLKIT</small></span>
      </button>

      <div class="side-section-label">WORKSPACE</div>
      <button class="nav-link" :class="{ active: currentToolId === 'home' }" @click="openTool('home')">
        <Layers3 :size="18" /><span>Dashboard</span><span class="nav-chevron">⌘ 1</span>
      </button>
      <div class="side-section-label nav-section-label">KATEGORI</div>
      <button v-for="group in groups" :key="group" class="nav-link" :class="{ active: selectedGroup === group && currentToolId === 'home' }" @click="chooseGroup(group)">
        <Files v-if="group === 'Semua alat' || group === 'Dokumen'" :size="18" />
        <Image v-else-if="group === 'Gambar'" :size="18" />
        <GraduationCap v-else-if="group === 'Akademik'" :size="18" />
        <Clock3 v-else :size="18" />
        <span>{{ group }}</span>
        <span v-if="group === 'Semua alat'" class="count-chip">{{ toolCatalog.length }}</span>
      </button>

      <div class="sidebar-spacer"></div>
      <div class="privacy-card">
        <div class="privacy-icon"><ShieldCheck :size="18" /></div>
        <div><strong>Local-first tools</strong><p>File diproses di perangkat jika memungkinkan.</p></div>
      </div>
      <div class="sidebar-footer"><span class="status-dot"></span> ToolkitIni Next <span class="version">v0.2</span></div>
    </aside>

    <main class="main-area">
      <header class="topbar">
        <div class="breadcrumb"><span>ToolkitIni</span><span class="crumb-slash">/</span><strong>{{ currentTool?.title || (selectedGroup === 'Semua alat' ? 'Overview' : selectedGroup) }}</strong></div>
        <div class="topbar-actions">
          <label class="global-search"><Search :size="16"/><input v-model="searchQuery" placeholder="Cari alat..." @focus="currentToolId = 'home'"/><kbd>Ctrl K</kbd></label>
          <div class="avatar">A</div>
        </div>
      </header>

      <div class="content-wrap">
        <template v-if="currentToolId === 'home'">
          <section class="hero-section">
            <div class="hero-copy">
              <div class="eyebrow"><Sparkles :size="14" /> YOUR STUDY WORKSPACE</div>
              <h1>Kerja kuliah, <span>lebih teratur.</span></h1>
              <p>Tools praktis buat dokumen, materi, dan deadline. Satu tempat, lebih sedikit tab browser yang bikin laptop menyerah.</p>
              <div class="hero-buttons"><button class="btn-primary" @click="openTool('image-to-pdf')"><FilePlus2 :size="17"/> Mulai dengan Foto ke PDF <ArrowDownToLine :size="15" class="arrow-side"/></button><button class="btn-quiet" @click="selectedGroup = 'Semua alat'; searchQuery = ''">Jelajahi semua alat <ArrowLeftRight :size="15"/></button></div>
              <div class="hero-proof"><span><CheckCircle2 :size="15"/> 17 tools terintegrasi</span><span><ShieldCheck :size="15"/> Fokus privasi</span><span><Zap :size="15"/> Tanpa akun</span></div>
            </div>
            <div class="hero-art" aria-hidden="true">
              <div class="art-orbit orbit-a"></div><div class="art-orbit orbit-b"></div>
              <div class="art-card art-card-back"><FileText :size="25"/><span>Assignment.pdf</span><small>2.4 MB</small></div>
              <div class="art-card art-card-front"><div class="art-pdf">PDF</div><div class="art-lines"><i></i><i></i><i></i><i></i></div><span>Ready to submit</span><CheckCircle2 :size="17" class="art-check"/></div>
              <div class="art-float float-one"><ScanText :size="17"/><span>OCR ready</span></div><div class="art-float float-two"><GraduationCap :size="17"/><span>Study smarter</span></div>
            </div>
          </section>

          <section class="stats-grid">
            <div class="stat-card"><div class="stat-symbol lilac"><Files :size="18"/></div><div><span class="stat-label">SEMUA TOOLS</span><strong>{{ toolCatalog.length }} <small>fitur</small></strong><p>Dokumen sampai akademik</p></div></div>
            <div class="stat-card"><div class="stat-symbol mint"><CalendarClock :size="18"/></div><div><span class="stat-label">TUGAS AKTIF</span><strong>{{ openAssignments }} <small>tugas</small></strong><p>Tersimpan di perangkat ini</p></div></div>
            <div class="stat-card"><div class="stat-symbol amber"><ShieldCheck :size="18"/></div><div><span class="stat-label">MODE PRIVASI</span><strong>Local <small>first</small></strong><p>Tanpa upload ke server sendiri</p></div></div>
          </section>

          <section class="section-block">
            <div class="section-heading"><div><div class="eyebrow muted">GET STARTED</div><h2>Tools yang paling kepakai</h2><p>Mulai dari kebutuhan yang biasanya muncul saat ngerjain tugas.</p></div><button class="text-button" @click="selectedGroup = 'Semua alat'; searchQuery = ''">Lihat semua <ArrowLeftRight :size="15"/></button></div>
            <div class="tool-grid featured-grid">
              <button v-for="tool in homeTools" :key="tool.id" class="tool-card" @click="openTool(tool.id)">
                <span class="tool-icon" :class="'tone-' + tool.group.toLowerCase()"><component :is="iconMap[tool.icon] || FileText" :size="20"/></span>
                <span v-if="tool.tag" class="tool-tag">{{ tool.tag }}</span>
                <strong>{{ tool.title }}</strong><p>{{ tool.description }}</p><span class="tool-card-footer">Buka tool <span>↗</span></span>
              </button>
            </div>
          </section>

          <section class="section-block all-tools-section">
            <div class="section-heading"><div><div class="eyebrow muted">TOOL LIBRARY</div><h2>{{ selectedGroup === 'Semua alat' ? 'Semua alat produktivitas' : selectedGroup }}</h2><p>{{ filteredTools.length }} alat tersedia dalam workspace ToolkitIni.</p></div><div class="filter-indicator"><Filter :size="15"/> <span>{{ selectedGroup }}</span><ChevronDown :size="14"/></div></div>
            <div v-if="filteredTools.length" class="tool-grid">
              <button v-for="tool in filteredTools" :key="tool.id" class="tool-card compact-card" @click="openTool(tool.id)">
                <span class="tool-icon" :class="'tone-' + tool.group.toLowerCase()"><component :is="iconMap[tool.icon] || FileText" :size="20"/></span>
                <strong>{{ tool.title }}</strong><p>{{ tool.description }}</p><span class="tool-card-footer">Buka tool <span>↗</span></span>
              </button>
            </div>
            <div v-else class="empty-state"><Search :size="24"/><strong>Nggak ada tool yang cocok.</strong><p>Coba kata kunci lain. Mesin pencari juga punya hari buruk.</p></div>
          </section>
          <footer class="main-footer"><span>ToolkitIni · Built for everyday campus life.</span><span>Privasi dulu. Produktivitas menyusul.</span></footer>
        </template>

        <template v-else-if="currentTool">
          <section class="tool-page-heading">
            <button class="back-link" @click="openTool('home')">← Kembali ke dashboard</button>
            <div class="tool-heading-row"><span class="tool-icon large" :class="'tone-' + currentTool.group.toLowerCase()"><component :is="iconMap[currentTool.icon] || FileText" :size="25"/></span><div><div class="eyebrow muted">{{ currentTool.group.toUpperCase() }} TOOL</div><h1>{{ currentTool.title }}</h1><p>{{ currentTool.description }}</p></div></div>
          </section>

          <section class="workspace-grid">
            <div class="workspace-main">
              <div v-if="requiresFiles" class="panel file-panel">
                <div class="panel-heading"><div><h3>Pilih file</h3><p>{{ currentTool.detail }}</p></div><span class="step-badge">01</span></div>
                <label class="dropzone" :class="{ 'dropzone-active': dropActive }" @dragenter.prevent="dropActive = true" @dragover.prevent="dropActive = true" @dragleave.prevent="dropActive = false" @drop.prevent="onDrop">
                  <input :key="fileInputKey" type="file" :accept="currentTool.accept || '*/*'" :multiple="currentTool.multiple !== false" @change="addFiles" />
                  <span class="upload-icon"><Upload :size="22"/></span><strong>Tarik file ke sini atau <em>pilih file</em></strong><span class="upload-hint">{{ currentTool.multiple === false ? 'Satu file untuk satu proses' : 'Bisa pilih beberapa file sekaligus' }} · Pemrosesan lokal jika didukung</span>
                </label>
                <div v-if="selectedFiles.length" class="selected-files">
                  <div class="selected-head"><span>{{ selectedFiles.length }} file dipilih</span><span>{{ formatBytes(fileTotal) }}</span></div>
                  <div v-for="(file, index) in selectedFiles" :key="file.name + file.size + index" class="file-row"><span class="file-type-icon"><FileText :size="17"/></span><div class="file-meta"><strong>{{ file.name }}</strong><small>{{ formatBytes(file.size) }}</small></div><button class="icon-button danger-hover" @click="removeFile(index)" aria-label="Hapus file"><X :size="16"/></button></div>
                </div>
              </div>

              <div v-if="currentToolId === 'pdf-compress'" class="panel options-panel"><div class="panel-heading"><div><h3>Pengaturan kompresi</h3><p>Semakin kecil skala dan kualitas, biasanya semakin kecil file hasilnya.</p></div><span class="step-badge">02</span></div><label class="field-label">Resolusi render <strong>{{ Math.round(compressScale * 100) }}%</strong></label><input v-model.number="compressScale" class="range-input" type="range" min="0.35" max="1" step="0.05"/><div class="range-ends"><span>File lebih kecil</span><span>Lebih tajam</span></div><label class="field-label spaced-label">Kualitas JPEG <strong>{{ Math.round(compressQuality * 100) }}%</strong></label><input v-model.number="compressQuality" class="range-input" type="range" min="0.35" max="0.95" step="0.05"/><div class="warning-note"><CircleHelp :size="17"/><span>Kompresi ini merasterisasi halaman. Gunakan salinan dan cek keterbacaan sebelum submit.</span></div></div>

              <div v-if="currentToolId === 'pdf-extract'" class="panel options-panel"><div class="panel-heading"><div><h3>Halaman yang diambil</h3><p>Gunakan nomor halaman dan rentang dipisahkan koma.</p></div><span class="step-badge">02</span></div><label class="field-label">Rentang halaman</label><input v-model="pageRange" class="text-input" placeholder="Contoh: 1,3-5,8"/><p class="helper-text">Contoh <code>1,3-5,8</code> akan mengambil halaman 1, 3, 4, 5, dan 8.</p></div>

              <div v-if="currentToolId === 'pdf-organize'" class="panel options-panel"><div class="panel-heading"><div><h3>Susunan halaman baru</h3><p>Masukkan urutan halaman sumber. Halaman yang tidak ditulis akan dihapus.</p></div><span class="step-badge">02</span></div><label class="field-label">Urutan halaman</label><input v-model="pageOrder" class="text-input" placeholder="Misalnya: 3,1,2,5"/><label class="field-label spaced-label">Halaman yang diputar</label><input v-model="rotatePages" class="text-input" placeholder="Misalnya: 2,4"/><div class="inline-field"><label class="radio-pill"><input v-model="rotateAngle" type="radio" value="90"/> Putar 90°</label><label class="radio-pill"><input v-model="rotateAngle" type="radio" value="180"/> Putar 180°</label><label class="radio-pill"><input v-model="rotateAngle" type="radio" value="270"/> Putar 270°</label></div><p class="helper-text">Nomor halaman pada kolom putar mengacu ke nomor halaman sumber.</p></div>

              <div v-if="currentToolId === 'pdf-to-docx'" class="panel options-panel"><div class="panel-heading"><div><h3>Format keluaran</h3><p>Pilih format yang akan dihasilkan.</p></div><span class="step-badge">02</span></div><div class="choice-grid"><label class="choice-card" :class="{ chosen: pdfOutputFormat === 'txt' }"><input v-model="pdfOutputFormat" type="radio" value="txt"/><FileText :size="21"/><strong>TXT</strong><small>Teks sederhana, cepat, dan ringan</small></label><label class="choice-card" :class="{ chosen: pdfOutputFormat === 'docx' }"><input v-model="pdfOutputFormat" type="radio" value="docx"/><Files :size="21"/><strong>Word DOCX</strong><small>Teks dalam dokumen yang bisa diedit</small></label></div></div>

              <div v-if="currentToolId === 'pdf-to-images' || currentToolId === 'image-tools'" class="panel options-panel"><div class="panel-heading"><div><h3>Format & ukuran</h3><p>Sesuaikan kebutuhan file hasil.</p></div><span class="step-badge">02</span></div><label class="field-label">Format output</label><select v-model="imageFormat" class="text-input"><option value="image/png">PNG · kualitas lossless</option><option value="image/jpeg">JPG · ukuran lebih kecil</option><option v-if="currentToolId === 'image-tools'" value="image/webp">WebP · cocok untuk web</option></select><template v-if="currentToolId === 'image-tools'"><label class="field-label spaced-label">Dimensi maksimum (px)</label><input v-model.number="imageMaxSide" type="number" min="0" max="12000" class="text-input"/><label class="field-label spaced-label">Kualitas <strong>{{ Math.round(imageQuality * 100) }}%</strong></label><input v-model.number="imageQuality" class="range-input" type="range" min="0.35" max="1" step="0.05"/></template></div>

              <div v-if="currentToolId === 'scanner'" class="panel options-panel"><div class="panel-heading"><div><h3>Perbaiki hasil scan</h3><p>Crop menggunakan persentase dari gambar sumber.</p></div><span class="step-badge">02</span></div><div class="form-grid two"><label class="form-field"><span>Crop X (%)</span><input v-model.number="scanCrop.x" type="number" min="0" max="99" class="text-input"/></label><label class="form-field"><span>Crop Y (%)</span><input v-model.number="scanCrop.y" type="number" min="0" max="99" class="text-input"/></label><label class="form-field"><span>Lebar crop (%)</span><input v-model.number="scanCrop.width" type="number" min="1" max="100" class="text-input"/></label><label class="form-field"><span>Tinggi crop (%)</span><input v-model.number="scanCrop.height" type="number" min="1" max="100" class="text-input"/></label></div><label class="field-label spaced-label">Kontras <strong>{{ scanContrast.toFixed(2) }}×</strong></label><input v-model.number="scanContrast" class="range-input" type="range" min="0.7" max="2.5" step="0.05"/><div class="inline-field"><label class="check-pill"><input v-model="scanGrayscale" type="checkbox"/> Grayscale</label><label class="check-pill"><input v-model="scanThreshold" type="checkbox"/> Hitam-putih tajam</label></div><p class="helper-text">Bila opsi hitam-putih diaktifkan, hasil akan memakai ambang kontras tinggi. Periksa teks kecil agar tidak hilang.</p></div>

              <div v-if="currentToolId === 'pdf-annotate'" class="panel options-panel"><div class="panel-heading"><div><h3>Atur anotasi</h3><p>Catatan akan ditambahkan ke semua halaman yang dipilih.</p></div><span class="step-badge">02</span></div><label class="field-label">Jenis anotasi</label><div class="inline-field"><label class="radio-pill"><input v-model="annotationMode" type="radio" value="highlight"/> Highlight</label><label class="radio-pill"><input v-model="annotationMode" type="radio" value="text"/> Catatan teks</label></div><label class="field-label spaced-label">Teks anotasi (opsional untuk highlight)</label><textarea v-model="annotationText" class="text-input text-area" rows="3" placeholder="Tulis catatan singkat..."/><label class="field-label spaced-label">Halaman (kosong = semua halaman)</label><input v-model="annotationPages" class="text-input" placeholder="Contoh: 1,3-5"/></div>

              <div v-if="currentToolId === 'pdf-watermark'" class="panel options-panel"><div class="panel-heading"><div><h3>Watermark & nomor halaman</h3><p>Atur identitas dan penomoran laporan.</p></div><span class="step-badge">02</span></div><label class="field-label">Teks watermark</label><input v-model="watermarkText" class="text-input" placeholder="Nama / Draft / ToolkitIni"/><label class="check-pill spaced-label"><input v-model="addPageNumbers" type="checkbox"/> Tambahkan nomor halaman</label><label class="field-label spaced-label">Nomor halaman dimulai dari</label><input v-model.number="numberStart" type="number" min="1" class="text-input"/></div>

              <div v-if="currentToolId === 'gpa'" class="panel academic-panel"><div class="panel-heading"><div><h3>Mata kuliah semester ini</h3><p>Nilai berbobot SKS untuk simulasi IPS. Skala nilai memakai A=4, AB=3,5, B=3, BC=2,5, C=2, D=1, E=0.</p></div><span class="step-badge">01</span></div><div class="course-list"><div v-for="(row, index) in courseRows" :key="index" class="course-row"><input v-model="row.name" class="text-input course-name" placeholder="Nama mata kuliah"/><input v-model.number="row.credits" class="text-input credits-input" type="number" min="1" max="12" aria-label="SKS"/><select v-model="row.grade" class="text-input grade-input"><option v-for="grade in gradeOptions" :key="grade" :value="grade">{{ grade }}</option></select><button class="icon-button danger-hover" @click="removeCourse(index)" aria-label="Hapus mata kuliah"><Trash2 :size="16"/></button></div></div><div class="course-add-row"><button class="btn-outline" @click="addCourse"><Plus :size="16"/> Tambah mata kuliah</button><span>{{ gpaResult.credits }} SKS dihitung</span></div><div class="gpa-result"><div><span>Estimasi IPS</span><strong>{{ gpaResult.gpa.toFixed(2) }}</strong><small>{{ gpaResult.credits ? 'Berdasarkan nilai dan SKS yang diisi' : 'Masukkan data mata kuliah' }}</small></div><div class="gpa-meter"><div class="gpa-meter-track"><i :style="{ width: `${Math.max(0, Math.min(4, gpaResult.gpa)) / 4 * 100}%` }"></i></div><div class="range-ends"><span>0.00</span><span>4.00</span></div></div></div></div>

              <div v-if="currentToolId === 'assignments'" class="panel academic-panel"><div class="panel-heading"><div><h3>Tambah tugas baru</h3><p>Data disimpan di localStorage browser/perangkat ini, tanpa sinkronisasi cloud.</p></div><span class="step-badge">01</span></div><div class="form-grid two"><label class="form-field span-two"><span>Nama tugas *</span><input v-model="assignmentForm.title" class="text-input" placeholder="Misalnya: Makalah Sistem Informasi" @keydown.enter.prevent="addAssignment"/></label><label class="form-field"><span>Mata kuliah</span><input v-model="assignmentForm.course" class="text-input" placeholder="Contoh: Analisis Proses Bisnis"/></label><label class="form-field"><span>Tenggat</span><input v-model="assignmentForm.due" class="text-input" type="date"/></label><label class="form-field span-two"><span>Prioritas</span><select v-model="assignmentForm.priority" class="text-input"><option>Rendah</option><option>Sedang</option><option>Tinggi</option></select></label></div><button class="btn-primary full-button" :disabled="!assignmentForm.title.trim()" @click="addAssignment"><Plus :size="16"/> Simpan tugas</button><div class="assignment-list-header"><h3>Daftar tugas</h3><span>{{ assignments.length }} total · {{ openAssignments }} belum selesai</span></div><div v-if="sortedAssignments.length" class="assignment-list"><div v-for="task in sortedAssignments" :key="task.id" class="assignment-item" :class="{ completed: task.done }"><label class="task-check"><input v-model="task.done" type="checkbox"/><span></span></label><div class="assignment-copy"><strong>{{ task.title }}</strong><p>{{ task.course }} <span>·</span> {{ task.due || 'Tanpa tenggat' }}</p><div class="task-meta"><span class="priority" :class="'priority-' + task.priority.toLowerCase()">{{ task.priority }}</span><span :class="{ overdue: !task.done && task.due && new Date(`${task.due}T23:59:59`) < new Date() }">{{ task.done ? 'Selesai' : daysUntil(task.due) }}</span></div></div><button class="icon-button danger-hover" @click="removeAssignment(task.id)" aria-label="Hapus tugas"><Trash2 :size="16"/></button></div></div><div v-else class="empty-mini"><CalendarClock :size="22"/><strong>Belum ada tugas.</strong><span>Catat sekarang, panik belakangan semoga berkurang.</span></div></div>

              <div v-if="currentToolId === 'citation'" class="panel academic-panel"><div class="panel-heading"><div><h3>Data sumber referensi</h3><p>Masukkan metadata yang benar. Hasil generator tetap perlu diperiksa.</p></div><span class="step-badge">01</span></div><label class="field-label">Gaya sitasi</label><div class="inline-field"><label class="radio-pill"><input v-model="citationStyle" type="radio" value="apa"/> APA 7</label><label class="radio-pill"><input v-model="citationStyle" type="radio" value="ieee"/> IEEE</label></div><div class="form-grid two citation-fields"><label class="form-field span-two"><span>Penulis (pisahkan dengan titik koma)</span><input v-model="citationData.authors" class="text-input" placeholder="Nama Belakang, Inisial; Penulis Kedua"/></label><label class="form-field span-two"><span>Judul *</span><input v-model="citationData.title" class="text-input" placeholder="Judul artikel / buku"/></label><label class="form-field"><span>Tahun</span><input v-model="citationData.year" class="text-input" placeholder="2026"/></label><label class="form-field"><span>Nama jurnal / penerbit</span><input v-model="citationData.source" class="text-input" placeholder="Nama sumber"/></label><label class="form-field"><span>DOI (opsional)</span><input v-model="citationData.doi" class="text-input" placeholder="10.xxxx/xxxxx"/></label><label class="form-field"><span>URL (opsional)</span><input v-model="citationData.url" class="text-input" placeholder="https://..."/></label></div><button class="btn-primary full-button" :disabled="!citationData.title.trim()" @click="runCurrentTool"><Quote :size="16"/> Buat sitasi</button><div v-if="citationResult" class="citation-output"><div class="output-top"><strong>{{ citationStyle.toUpperCase() }} · HASIL</strong><button class="text-button" @click="copyText(citationResult)">Salin <FileText :size="14"/></button></div><p>{{ citationResult }}</p></div></div>

              <div v-if="currentToolId === 'pomodoro'" class="panel timer-panel"><div class="panel-heading"><div><h3>Waktu fokus</h3><p>Kerjakan satu hal dulu. Multitasking sering cuma membuka banyak tab dengan penuh percaya diri.</p></div><span class="step-badge">01</span></div><div class="timer-tabs"><button :class="{ active: timerMode === 'focus' }" :disabled="timerRunning" @click="timerMode = 'focus'; timerSeconds = focusLength * 60">Fokus</button><button :class="{ active: timerMode === 'break' }" :disabled="timerRunning" @click="timerMode = 'break'; timerSeconds = breakLength * 60">Istirahat</button></div><div class="timer-display"><div class="timer-ring"><svg viewBox="0 0 220 220" aria-hidden="true"><circle class="ring-track" cx="110" cy="110" r="96"/><circle class="ring-progress" cx="110" cy="110" r="96" :style="{ strokeDasharray: `${2 * Math.PI * 96}`, strokeDashoffset: `${2 * Math.PI * 96 * (1 - timerProgress / 100)}` }"/></svg><div class="timer-center"><span>{{ timerMode === 'focus' ? 'FOCUS SESSION' : 'TAKE A BREAK' }}</span><strong>{{ timerLabel }}</strong><small>{{ timerRunning ? 'Sedang berjalan' : 'Siap saat kamu siap' }}</small></div></div></div><div class="timer-controls"><button class="btn-primary" @click="toggleTimer">{{ timerRunning ? 'Jeda timer' : 'Mulai timer' }} <span>{{ timerRunning ? 'Ⅱ' : '▶' }}</span></button><button class="btn-outline" @click="resetTimer">Reset</button></div><div class="form-grid two timer-settings"><label class="form-field"><span>Durasi fokus (menit)</span><input v-model.number="focusLength" class="text-input" type="number" min="1" max="180" :disabled="timerRunning"/></label><label class="form-field"><span>Durasi istirahat (menit)</span><input v-model.number="breakLength" class="text-input" type="number" min="1" max="60" :disabled="timerRunning"/></label></div><p class="helper-text centered">Timer berjalan selama halaman aplikasi terbuka. Notifikasi sistem belum digunakan.</p></div>

              <div v-if="currentToolId === 'study-notes'" class="panel academic-panel"><div class="panel-heading"><div><h3>Buat catatan belajar</h3><p>Ekstrak teks dari file, susun gagasan utama, kata kunci, dan pertanyaan latihan.</p></div><span class="step-badge">02</span></div><div class="warning-note"><BrainCircuit :size="17"/><span>Ini bukan AI generatif. Ringkasan berbasis heuristik, sehingga cocok sebagai draf awal dan bukan pengganti membaca materi.</span></div><div v-if="studyNotes" class="study-notes-output"><div class="notes-section"><h4><BookOpen :size="16"/> Draf ringkasan</h4><ul><li v-for="(sentence, i) in studyNotes.summary" :key="i">{{ sentence }}</li></ul></div><div class="notes-section"><h4><Sparkles :size="16"/> Kata kunci</h4><div class="keyword-list"><span v-for="word in studyNotes.keywords" :key="word">{{ word }}</span></div></div><div class="notes-section"><h4><CircleHelp :size="16"/> Pertanyaan latihan</h4><ol><li v-for="question in studyNotes.questions" :key="question">{{ question.replace(/^\d+\.\s*/, '') }}</li></ol></div><button class="btn-outline" @click="copyText(['RINGKASAN', ...studyNotes.summary, '', 'KATA KUNCI', studyNotes.keywords.join(', '), '', 'PERTANYAAN LATIHAN', ...studyNotes.questions].join('\n'))"><FileText :size="15"/> Salin semua catatan</button></div><div v-else class="empty-mini"><BookOpen :size="22"/><strong>Belum ada catatan.</strong><span>Pilih PDF yang punya teks atau foto materi di bagian atas, lalu jalankan prosesnya.</span></div></div>

              <div class="run-panel"><div class="run-summary"><span class="run-summary-icon"><Zap :size="17"/></span><div><strong>{{ currentTool.action }}</strong><p>{{ requiresFiles ? 'File sumber tidak akan ditimpa.' : 'Hasil akan ditampilkan di halaman ini.' }}</p></div></div><button class="btn-primary run-button" :disabled="busy || (requiresFiles && !selectedFiles.length)" @click="runCurrentTool"><span v-if="busy" class="spinner"></span><component v-else :is="iconMap[currentTool.icon] || Zap" :size="17"/>{{ busy ? 'Sedang memproses...' : currentTool.action }} <ArrowDownToLine v-if="!busy" :size="15"/></button><div v-if="lastResult" class="result-success"><CheckCircle2 :size="17"/><span>{{ lastResult }}</span></div></div>
            </div>

            <aside class="workspace-aside"><div class="panel aside-panel"><div class="aside-heading"><ShieldCheck :size="18"/><strong>Privasi & keamanan</strong></div><ul class="check-list"><li><Check :size="15"/> File sumber tidak ditimpa</li><li><Check :size="15"/> Pemrosesan lokal jika didukung</li><li><Check :size="15"/> Tidak perlu akun ToolkitIni</li></ul><p class="aside-note">OCR memerlukan pengunduhan model bahasa saat pertama kali digunakan. Fitur API eksternal tidak aktif di versi ini.</p></div><div class="panel aside-panel help-panel"><div class="aside-heading"><CircleHelp :size="18"/><strong>Tips penggunaan</strong></div><div class="tip-item"><span>01</span><p>Gunakan salinan file asli untuk operasi kompresi, anotasi, dan perubahan halaman.</p></div><div class="tip-item"><span>02</span><p>Periksa hasil unduhan sebelum mengumpulkan tugas di portal kampus.</p></div><div class="tip-item"><span>03</span><p>Untuk dokumen rahasia, hindari mengunggah file ke layanan web yang tidak jelas.</p></div></div><div class="mini-promo"><div class="mini-promo-icon"><Sparkles :size="18"/></div><strong>Satu tempat. Lebih fokus.</strong><p>ToolkitIni dirancang untuk kebutuhan kuliah sehari-hari.</p><button @click="openTool('home')">Jelajahi tool lain <span>↗</span></button></div></aside>
          </section>
          <footer class="main-footer"><span>ToolkitIni · Built for everyday campus life.</span><span>Hasil otomatis tetap perlu ditinjau.</span></footer>
        </template>
      </div>
    </main>

    <transition name="toast"><div v-if="status" class="toast-message" :class="'toast-' + statusType"><CheckCircle2 v-if="statusType === 'success'" :size="18"/><CircleHelp v-else :size="18"/><span>{{ status }}</span><button @click="status = ''" aria-label="Tutup pemberitahuan"><X :size="15"/></button></div></transition>
  </div>
</template>
