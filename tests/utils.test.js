import test from 'node:test'
import assert from 'node:assert/strict'
import { parsePageRanges, parsePageOrder, formatBytes, calculateGpa, buildCitation, makeStudyNotes } from '../src/lib/utils.js'

test('parsePageRanges expands, sorts, and de-duplicates ranges', () => {
  assert.deepEqual(parsePageRanges('3,1-2,2', 5), [0, 1, 2])
})

test('parsePageRanges rejects pages outside document', () => {
  assert.throws(() => parsePageRanges('1,9', 5), /antara 1 dan 5/)
})

test('parsePageOrder preserves requested order and allows omitted pages', () => {
  assert.deepEqual(parsePageOrder('3,1', 3), [2, 0])
})

test('parsePageOrder rejects duplicate pages', () => {
  assert.throws(() => parsePageOrder('1,1', 3), /duplikat/)
})

test('formatBytes formats byte counts', () => {
  assert.equal(formatBytes(1024), '1.00 KB')
  assert.equal(formatBytes(0), '0 B')
})

test('calculateGpa weights grades by credits', () => {
  const result = calculateGpa([{ credits: 3, grade: 'A' }, { credits: 1, grade: 'B' }])
  assert.equal(result.credits, 4)
  assert.equal(result.gpa, 3.75)
})

test('buildCitation includes essential APA data', () => {
  assert.match(buildCitation({ authors: 'A. Author', title: 'Judul', year: '2024', source: 'Jurnal X' }, 'apa'), /A\. Author \(2024\)\. Judul\. \*Jurnal X\*\./)
})

test('makeStudyNotes rejects very short text', () => {
  assert.throws(() => makeStudyNotes('Teks pendek'), /terlalu pendek/)
})
