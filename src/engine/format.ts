import { formatGeneral, isError, parseNumber, type Scalar } from './values.js'

/** Excel serial date ↔ JS date (1900 date system, epoch 1899-12-30). */
const EPOCH = Date.UTC(1899, 11, 30)
const DAY_MS = 86_400_000

export function dateToSerial (y: number, m: number, d: number, hh = 0, mm = 0, ss = 0): number {
  return (Date.UTC(y, m - 1, d, hh, mm, ss) - EPOCH) / DAY_MS
}

export function serialToDate (serial: number): Date {
  return new Date(EPOCH + Math.round(serial * DAY_MS))
}

export const FORMATS = {
  general: 'General',
  number: '#,##0.00',
  integer: '0',
  currency: '$#,##0.00',
  percent: '0.00%',
  scientific: '0.00E+00',
  date: 'yyyy-mm-dd',
  datetime: 'yyyy-mm-dd hh:mm',
  time: 'hh:mm:ss',
  text: '@'
} as const

export interface ParsedInput {
  kind: 'empty' | 'formula' | 'value'
  value: Scalar
  /** A number format implied by how the value was typed ("12%", "$5"). */
  format?: string
}

const ISO_DATE_RE = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/
const US_DATE_RE = /^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/
const TIME_RE = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/

/** Interpret raw cell input the way a spreadsheet does when you type it. */
export function parseInput (raw: string): ParsedInput {
  if (raw === '') return { kind: 'empty', value: null }
  if (raw.length > 1 && raw[0] === '=') return { kind: 'formula', value: null }
  if (raw[0] === "'") return { kind: 'value', value: raw.slice(1) }
  const t = raw.trim()
  const upper = t.toUpperCase()
  if (upper === 'TRUE') return { kind: 'value', value: true }
  if (upper === 'FALSE') return { kind: 'value', value: false }

  const n = parseNumber(t)
  if (n !== undefined) {
    if (t.endsWith('%')) return { kind: 'value', value: n, format: t.includes('.') ? '0.00%' : '0%' }
    if (t.includes(',')) return { kind: 'value', value: n, format: t.includes('.') ? '#,##0.00' : '#,##0' }
    return { kind: 'value', value: n }
  }
  const money = /^([+-]?)\$\s?([\d,]*\.?\d+)$/.exec(t)
  if (money) {
    const v = parseNumber(money[2])
    if (v !== undefined) return { kind: 'value', value: money[1] === '-' ? -v : v, format: FORMATS.currency }
  }
  let m = ISO_DATE_RE.exec(t)
  if (m && !validTime(m[4], m[5], m[6])) m = null
  if (m) {
    const serial = dateToSerial(+m[1], +m[2], +m[3], +(m[4] ?? 0), +(m[5] ?? 0), +(m[6] ?? 0))
    if (validDate(+m[1], +m[2], +m[3])) return { kind: 'value', value: serial, format: m[4] ? FORMATS.datetime : FORMATS.date }
  }
  m = US_DATE_RE.exec(t)
  if (m && !validTime(m[4], m[5], m[6])) m = null
  if (m) {
    const serial = dateToSerial(+m[3], +m[1], +m[2], +(m[4] ?? 0), +(m[5] ?? 0), +(m[6] ?? 0))
    if (validDate(+m[3], +m[1], +m[2])) return { kind: 'value', value: serial, format: m[4] ? 'm/d/yyyy hh:mm' : 'm/d/yyyy' }
  }
  m = TIME_RE.exec(t)
  if (m && +m[1] < 24 && +m[2] < 60) {
    return { kind: 'value', value: (+m[1] * 3600 + +m[2] * 60 + +(m[3] ?? 0)) / 86400, format: m[3] ? FORMATS.time : 'hh:mm' }
  }
  return { kind: 'value', value: raw }
}

function validTime (h?: string, m?: string, s?: string): boolean {
  return (h === undefined || +h < 24) && (m === undefined || +m < 60) && (s === undefined || +s < 60)
}

function validDate (y: number, m: number, d: number): boolean {
  if (y < 1900 || m < 1 || m > 12 || d < 1) return false
  return d <= new Date(Date.UTC(y, m, 0)).getUTCDate()
}

const formatterCache = new Map<string, (n: number) => string>()

/** Render a value for display under a number format. */
export function formatValue (v: Scalar, format: string | undefined): string {
  if (v === null) return ''
  if (typeof v === 'string') return v
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
  if (isError(v)) return v.code
  if (!format || format === 'General' || format === '@') return formatGeneral(v)
  let fn = formatterCache.get(format)
  if (!fn) {
    fn = compileFormat(format)
    formatterCache.set(format, fn)
  }
  return fn(v)
}

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']

function compileFormat (format: string): (n: number) => string {
  // Sections: positive;negative;zero
  const sections = splitSections(format)
  if (sections.length > 1) {
    const [pos, neg, zero] = sections.map(compileSection)
    return (n) => (n > 0 ? pos(n) : n < 0 ? (neg ?? pos)(-n) : (zero ?? pos)(n))
  }
  const only = compileSection(sections[0])
  return only
}

function splitSections (format: string): string[] {
  const out: string[] = []
  let cur = ''
  let quoted = false
  for (const ch of format) {
    if (ch === '"') quoted = !quoted
    if (ch === ';' && !quoted) {
      out.push(cur)
      cur = ''
    } else cur += ch
  }
  out.push(cur)
  return out
}

function compileSection (section: string): (n: number) => string {
  // [Red], [>100] and similar tags: colours and conditions aren't rendered.
  section = section.replace(/\[[^\]]*\]/g, '')
  const bare = section.replace(/"[^"]*"/g, '').replace(/\\./g, '')
  // Dates may contain 0s only as fractional seconds (ss.000).
  if (/[ymdhs]/i.test(bare) && !/[0#]/.test(bare.replace(/s+\.0+/gi, ''))) return compileDateFormat(section)
  // No digit placeholders at all (e.g. the "-" zero section of an accounting format): literal text.
  if (!/[0#?]/.test(bare)) {
    const text = section.replace(/"([^"]*)"/g, '$1').replace(/\\(.)/g, '$1').replace(/_./g, ' ').replace(/\*./g, '')
    return () => text
  }
  return compileNumberFormat(section)
}

function compileNumberFormat (section: string): (n: number) => string {
  // _x leaves room for a character (alignment padding) and *x repeats one; render a space / nothing.
  const literal = (s: string): string => s.replace(/"([^"]*)"/g, '$1').replace(/\\(.)/g, '$1').replace(/_./g, ' ').replace(/\*./g, '')
  const m = /^([^0#?]*)([0#?,]*)(?:\.([0#?]+))?(E[+-]0+)?([^0#?]*)$/i.exec(section)
  if (!m) return formatGeneral
  const prefix = literal(m[1])
  const intPart = m[2]
  const decimals = (m[3] ?? '').length
  // "0" decimals are always shown; "#" decimals only when needed (0.## → 1.5).
  const minDecimals = ((m[3] ?? '').match(/0/g) ?? []).length
  const sci = m[4]
  let suffix = literal(m[5])
  const percent = section.includes('%')
  const thousands = intPart.includes(',')
  const minInt = (intPart.match(/0/g) ?? []).length
  if (percent) suffix = suffix.includes('%') ? suffix : suffix + '%'
  const nf = new Intl.NumberFormat('en-US', {
    minimumFractionDigits: minDecimals,
    maximumFractionDigits: decimals,
    useGrouping: thousands,
    minimumIntegerDigits: Math.max(1, minInt)
  })
  return (n: number): string => {
    if (sci) {
      const exp = sci.length - 2
      const [mant, e] = n.toExponential(decimals).split('e')
      const ev = Number(e)
      const sign = ev < 0 ? '-' : sci[1] === '+' ? '+' : ''
      return prefix + mant + 'E' + sign + String(Math.abs(ev)).padStart(exp, '0') + suffix
    }
    const value = percent ? n * 100 : n
    let s = nf.format(Math.abs(value))
    if (minInt === 0 && s.startsWith('0') && decimals > 0) s = s.slice(1)
    const neg = value < 0 && Number(s.replace(/[^\d.]/g, '')) !== 0
    return (neg ? '-' : '') + prefix + s + suffix
  }
}

function compileDateFormat (section: string): (n: number) => string {
  // Tokenize: yyyy, yy, mmmm, mmm, mm, m, dddd, ddd, dd, d, hh, h, ss, s, AM/PM, literals
  const re = /(yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|am\/pm|a\/p|\.0+|"[^"]*"|\\.|.)/gi
  const parts = section.match(re) ?? []
  const ampm = /am\/pm|a\/p/i.test(section)
  // "m" right after an hour or right before seconds means minutes.
  const tokens = parts.map((p, i) => {
    const low = p.toLowerCase()
    if (low === 'm' || low === 'mm') {
      const prev = parts.slice(0, i).reverse().find((x) => /[a-z]/i.test(x))
      const nextTok = parts.slice(i + 1).find((x) => /[a-z]/i.test(x))
      if ((prev && /^h/i.test(prev)) || (nextTok && /^s/i.test(nextTok))) return low === 'm' ? 'MIN' : 'MMIN'
    }
    return p
  })
  const fracDigits = Math.min(3, Math.max(0, ...tokens.map((t) => (/^\.0+$/.test(t) ? t.length - 1 : 0))))
  const unit = 1000 / 10 ** fracDigits
  return (serial: number): string => {
    // Round to what's displayed: 23:59:59.9 shown without fractions is the next day at 00:00:00.
    const d = new Date(Math.round((EPOCH + serial * DAY_MS) / unit) * unit)
    const frac = d.getUTCMilliseconds()
    const Y = d.getUTCFullYear()
    const M = d.getUTCMonth()
    const D = d.getUTCDate()
    let H = d.getUTCHours()
    const Mi = d.getUTCMinutes()
    const S = d.getUTCSeconds()
    const pm = H >= 12
    if (ampm) H = H % 12 || 12
    let out = ''
    for (const tok of tokens) {
      switch (tok.toLowerCase()) {
        case 'yyyy': out += Y; break
        case 'yy': out += String(Y % 100).padStart(2, '0'); break
        case 'mmmm': out += MONTHS[M]; break
        case 'mmm': out += MONTHS[M].slice(0, 3); break
        case 'mm': out += String(M + 1).padStart(2, '0'); break
        case 'm': out += M + 1; break
        case 'dddd': out += DAYS[d.getUTCDay()]; break
        case 'ddd': out += DAYS[d.getUTCDay()].slice(0, 3); break
        case 'dd': out += String(D).padStart(2, '0'); break
        case 'd': out += D; break
        case 'hh': out += String(H).padStart(2, '0'); break
        case 'h': out += H; break
        case 'mmin': out += String(Mi).padStart(2, '0'); break
        case 'min': out += Mi; break
        case 'ss': out += String(S).padStart(2, '0'); break
        case 's': out += S; break
        case 'am/pm': out += pm ? 'PM' : 'AM'; break
        case 'a/p': out += pm ? 'P' : 'A'; break
        default:
          if (/^\.0+$/.test(tok)) out += '.' + String(frac).padStart(3, '0').slice(0, tok.length - 1)
          else if (tok.startsWith('"')) out += tok.slice(1, -1)
          else if (tok.startsWith('\\')) out += tok.slice(1)
          else out += tok
      }
    }
    return out
  }
}

/** Format a number with TEXT()-style format codes. */
export function textFormat (v: number, format: string): string {
  if (format === '' || format.toLowerCase() === 'general') return formatGeneral(v)
  return formatValue(v, format)
}

/** Raw input that reproduces a value exactly (for "paste values"). */
export function valueToRaw (v: Scalar): string {
  if (v === null) return ''
  if (typeof v === 'number') return String(v)
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
  if (isError(v)) return v.code
  const parsed = parseInput(v)
  return parsed.kind === 'value' && parsed.value === v ? v : "'" + v
}
