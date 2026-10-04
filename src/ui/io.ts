/** Delimited text (CSV/TSV) and file helpers. */

export function parseDelimited (text: string, delim = detectDelimiter(text)): string[][] {
  const rows: string[][] = []
  let row: string[] = []
  let field = ''
  let i = 0
  const n = text.length
  let quoted = false
  /** True once the current row has a field in progress (so a lone "" row is kept). */
  let started = false
  while (i < n) {
    const ch = text[i]
    if (quoted) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"'
          i += 2
          continue
        }
        quoted = false
        i++
        continue
      }
      field += ch
      i++
      continue
    }
    // A quote opens a quoted field only at the start of a field, and only if
    // it's closed somewhere; otherwise it's an ordinary character.
    if (ch === '"' && field === '' && text.indexOf('"', i + 1) !== -1) {
      quoted = true
      started = true
      i++
      continue
    }
    if (ch === delim) {
      row.push(field)
      field = ''
      started = true
      i++
      continue
    }
    if (ch === '\r' || ch === '\n') {
      row.push(field)
      rows.push(row)
      row = []
      field = ''
      started = false
      i += ch === '\r' && text[i + 1] === '\n' ? 2 : 1
      continue
    }
    field += ch
    started = true
    i++
  }
  if (started || field !== '' || row.length) {
    row.push(field)
    rows.push(row)
  }
  return rows
}

export function detectDelimiter (text: string): string {
  const nl = text.search(/\r|\n/)
  const firstLine = text.slice(0, nl < 0 ? 1000 : Math.min(nl, 10000))
  const tabs = firstLine.split('\t').length
  const commas = firstLine.split(',').length
  const semis = firstLine.split(';').length
  if (tabs >= commas && tabs >= semis && tabs > 1) return '\t'
  if (semis > commas) return ';'
  return ','
}

export function toDelimited (rows: string[][], delim: string): string {
  const quote = (s: string): string =>
    s.includes(delim) || s.includes('"') || s.includes('\n') || s.includes('\r') ? '"' + s.replace(/"/g, '""') + '"' : s
  return rows.map((r) => r.map(quote).join(delim)).join(delim === '\t' ? '\n' : '\r\n')
}

export function escapeHtml (s: string): string {
  return s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]!)
}

export function download (filename: string, text: string, mime: string): void {
  const blob = new Blob([text], { type: mime })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  document.body.append(a)
  a.click()
  a.remove()
  setTimeout(() => URL.revokeObjectURL(url), 1000)
}

/**
 * Neutralise text that another spreadsheet would run as a formula when the
 * CSV is opened there (CSV injection): =, +, -, @, tab and CR.
 */
export function csvSafe (text: string): string {
  return /^[=+\-@\t\r]/.test(text) && !/^[+-]?\d/.test(text) ? "'" + text : text
}
