import { useMemo, useState } from 'react'
import readXlsxFile, { readSheetNames } from 'read-excel-file'

const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 6)
const download = (blob, name) => {
  const a = document.createElement('a')
  a.href = URL.createObjectURL(blob)
  a.download = name
  a.click()
  setTimeout(() => URL.revokeObjectURL(a.href), 1000)
}

const BAD = 'Could not read that file. Use an .xlsx file (in Excel: File > Save As > Excel Workbook).'
const NUM = /^-?(0|[1-9]\d*|[1-9]\d{0,2}(,\d{3})+)(\.\d+)?$/
const ISO = /^\d{4}-\d{2}-\d{2}/
const blank = (v) => v === null || v === undefined || String(v).trim() === ''
const iso = (d) => d.toISOString().slice(0, 10)

const OPS = {
  text: [['contains', 'contains'], ['ncontains', 'does not contain'], ['eq', 'is exactly'], ['starts', 'starts with'], ['ends', 'ends with'], ['empty', 'is blank'], ['notempty', 'is not blank']],
  number: [['eq', 'equals'], ['ne', 'does not equal'], ['gt', 'is greater than'], ['gte', 'is at least'], ['lt', 'is less than'], ['lte', 'is at most'], ['between', 'is between'], ['empty', 'is blank'], ['notempty', 'is not blank']],
  date: [['eq', 'is on'], ['before', 'is before'], ['after', 'is after'], ['between', 'is between'], ['empty', 'is blank'], ['notempty', 'is not blank']],
  boolean: [['istrue', 'is Yes'], ['isfalse', 'is No'], ['empty', 'is blank']],
}
const PICK = [['any', 'is any of'], ['none', 'is none of']]
const DIR = { text: ['A to Z', 'Z to A'], number: ['Smallest first', 'Largest first'], date: ['Oldest first', 'Newest first'], boolean: ['No first', 'Yes first'] }
const NOVAL = ['empty', 'notempty', 'istrue', 'isfalse']

// ---- reading and "researching" the columns ----
const infer = (vals) => {
  const f = vals.filter((v) => !blank(v))
  const all = (fn) => f.every(fn)
  if (!f.length) return 'empty'
  if (all((v) => typeof v === 'boolean')) return 'boolean'
  if (all((v) => (v instanceof Date && !isNaN(v)) || (typeof v === 'string' && ISO.test(v.trim()) && !isNaN(Date.parse(v))))) return 'date'
  if (all((v) => typeof v === 'number' || (typeof v === 'string' && NUM.test(v.trim())))) return 'number'
  return 'text'
}
const norm = (v, type) => {
  if (blank(v)) return null
  if (type === 'number') { const n = typeof v === 'number' ? v : parseFloat(String(v).replace(/,/g, '')); return isNaN(n) ? null : n }
  if (type === 'date') { const d = v instanceof Date ? v : new Date(v); return isNaN(d) ? null : iso(d) }
  if (type === 'boolean') return typeof v === 'boolean' ? v : null
  return String(v).trim()
}
const guessHeader = (raw) => {
  const top = raw.slice(0, 10)
  const w = top.reduce((m, r) => Math.max(m, r.filter((v) => v != null).length), 0)
  const i = top.findIndex((r) => r.filter((v) => typeof v === 'string' && v.trim()).length >= Math.max(2, w * 0.6))
  return i < 0 ? 1 : i + 1
}
const buildTable = (raw, hr, ov) => {
  const head = raw[hr - 1] || []
  const body = raw.slice(hr).filter((r) => r.some((v) => !blank(v)))
  const width = body.reduce((m, r) => Math.max(m, r.length), head.length)
  const seen = {}
  const cols = Array.from({ length: width }, (_, i) => {
    let name = blank(head[i]) ? `Column ${i + 1}` : String(head[i]).trim()
    seen[name] = (seen[name] || 0) + 1
    if (seen[name] > 1) name += ` (${seen[name]})`
    const auto = infer(body.map((r) => r[i]))
    return { i, name, auto, type: ov[i] || auto }
  })
  const rows = body.map((r) => cols.map((c) => norm(r[c.i], c.type)))
  for (const c of cols) {
    const counts = new Map()
    let min = null, max = null, filled = 0
    for (const r of rows) {
      const v = r[c.i]
      if (v === null) continue
      filled++
      counts.set(v, (counts.get(v) || 0) + 1)
      if (c.type === 'number' || c.type === 'date') { if (min === null || v < min) min = v; if (max === null || v > max) max = v }
    }
    const d = counts.size
    Object.assign(c, { filled, distinct: d, min, max, pick: c.type === 'text' && (d <= 8 || (d <= 30 && d * 2 <= filled)) })
    c.choices = c.pick ? [...counts].sort((a, b) => b[1] - a[1]) : []
  }
  return { cols, rows }
}

// ---- filtering and sorting ----
const ready = (r) => NOVAL.includes(r.op) || (['any', 'none'].includes(r.op) ? r.list.length > 0 : r.op === 'between' ? r.a !== '' && r.b !== '' : r.a !== '')
const test = (r, t, v) => {
  if (r.op === 'empty') return v === null
  if (r.op === 'notempty') return v !== null
  if (v === null) return ['ncontains', 'none', 'ne'].includes(r.op)
  if (t === 'boolean') return r.op === 'istrue' ? v === true : v === false
  if (r.op === 'any' || r.op === 'none') return r.list.includes(v) === (r.op === 'any')
  const cv = (s) => (t === 'number' ? parseFloat(s) : t === 'text' ? s.toLowerCase() : s)
  const x = t === 'text' ? v.toLowerCase() : v
  const a = cv(r.a), b = cv(r.b || '')
  switch (r.op) {
    case 'contains': return x.includes(a)
    case 'ncontains': return !x.includes(a)
    case 'starts': return x.startsWith(a)
    case 'ends': return x.endsWith(a)
    case 'eq': return x === a
    case 'ne': return x !== a
    case 'gt': case 'after': return x > a
    case 'gte': return x >= a
    case 'lt': case 'before': return x < a
    case 'lte': return x <= a
    case 'between': return x >= (a < b ? a : b) && x <= (a < b ? b : a)
    default: return true
  }
}
const run = (table, rules, mode, sorts) => {
  const act = rules.filter(ready)
  const pass = (row) => (r) => test(r, table.cols[r.col].type, row[r.col])
  const rows = act.length ? table.rows.filter((row) => (mode === 'all' ? act.every(pass(row)) : act.some(pass(row)))) : [...table.rows]
  rows.sort((p, q) => {
    for (const s of sorts) {
      const a = p[s.col], b = q[s.col]
      if (a === b) continue
      if (a === null) return 1
      if (b === null) return -1
      const c = typeof a === 'string' ? a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }) : a < b ? -1 : 1
      return s.dir === 'desc' ? -c : c
    }
    return 0
  })
  const sum = {}, cnt = {}
  for (const c of table.cols) if (c.type === 'number') { sum[c.i] = 0; cnt[c.i] = 0 }
  for (const r of rows) for (const k in sum) if (r[k] !== null) { sum[k] += r[k]; cnt[k]++ }
  return { rows, sum, cnt }
}

const typeLabel = (c, t = c.type) => ({ number: 'Number', date: 'Date', boolean: 'Yes/No', empty: 'Empty' }[t] || (t === c.type && c.pick ? `Text, ${c.distinct} choices` : 'Text'))
const show = (v, c) => (v === null ? '' : c.type === 'boolean' ? (v ? 'Yes' : 'No') : String(v))
const n2 = (x) => x.toLocaleString(undefined, { maximumFractionDigits: 2 })
const csvCell = (x) => { const s = x === null ? '' : String(x); return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s }

export default function DataTool() {
  const [file, setFile] = useState(null)
  const [sheets, setSheets] = useState([])
  const [sheet, setSheet] = useState(1)
  const [raw, setRaw] = useState(null)
  const [hr, setHr] = useState(1)
  const [ov, setOv] = useState({})
  const [rules, setRules] = useState([])
  const [mode, setMode] = useState('all')
  const [sorts, setSorts] = useState([])
  const [hidden, setHidden] = useState([])
  const [limit, setLimit] = useState(100)
  const [err, setErr] = useState('')

  const table = useMemo(() => raw && buildTable(raw, hr, ov), [raw, hr, ov])
  const result = useMemo(() => table && run(table, rules, mode, sorts), [table, rules, mode, sorts])

  const reset = () => { setRules([]); setSorts([]); setHidden([]); setLimit(100) }
  const openSheet = async (f, idx) => {
    try {
      const rows = await readXlsxFile(f, { sheet: idx })
      setRaw(rows); setSheet(idx); setHr(guessHeader(rows)); setOv({}); reset(); setErr('')
    } catch { setErr(BAD) }
  }
  const onFile = async (e) => {
    const f = e.target.files[0]
    e.target.value = ''
    if (!f) return
    try { setSheets(await readSheetNames(f)); setFile(f); await openSheet(f, 1) } catch { setErr(BAD) }
  }

  const opsFor = (c) => [...(c.pick ? PICK : []), ...OPS[c.type === 'empty' ? 'text' : c.type]]
  const fresh = (c) => ({ col: c.i, op: opsFor(c)[0][0], a: '', b: '', list: [] })
  const setRule = (id, p) => setRules(rules.map((r) => (r.id === id ? { ...r, ...p } : r)))
  const setSort = (k, p) => setSorts(sorts.map((s, j) => (j === k ? { ...s, ...p } : s)))
  const vis = table ? table.cols.filter((c) => !hidden.includes(c.i)) : []

  const exportCsv = () => {
    const lines = [vis.map((c) => c.name), ...result.rows.map((r) => vis.map((c) => show(r[c.i], c)))]
    download(new Blob(['\ufeff' + lines.map((l) => l.map(csvCell).join(',')).join('\r\n')], { type: 'text/csv;charset=utf-8' }), 'filtered-data.csv')
  }

  return (
    <div>
      <section>
        <h2>Your file</h2>
        <label className="file">Choose Excel file (.xlsx)<input type="file" accept=".xlsx" onChange={onFile} /></label>
        {file && <span className="muted"> {file.name}</span>}
        <p className="muted">The file is read on this computer only. Nothing is uploaded anywhere.</p>
        {err && <p className="err" role="alert">{err}</p>}
        {table && (
          <div className="meta">
            {sheets.length > 1 && (
              <label>Sheet
                <select value={sheet} onChange={(e) => openSheet(file, +e.target.value)}>{sheets.map((s, i) => <option key={i} value={i + 1}>{s}</option>)}</select>
              </label>
            )}
            <label>Column names are on row
              <input type="number" min="1" max={raw.length} value={hr} onChange={(e) => { setHr(Math.max(1, +e.target.value || 1)); setOv({}); reset() }} />
            </label>
          </div>
        )}
      </section>

      {table && (
        <>
          <section>
            <h2>What we found in each column</h2>
            <p className="muted">{table.rows.length} rows, {table.cols.length} columns. If a column was understood wrongly, change its type.</p>
            <div className="tbl">
              <table>
                <thead><tr><th>Column</th><th>Type</th><th>Filled</th><th>Contains</th></tr></thead>
                <tbody>
                  {table.cols.map((c) => (
                    <tr key={c.i}>
                      <td>{c.name}</td>
                      <td>
                        <select value={ov[c.i] || 'auto'} onChange={(e) => { const o = { ...ov }; if (e.target.value === 'auto') delete o[c.i]; else o[c.i] = e.target.value; setOv(o); reset() }}>
                          <option value="auto">{typeLabel(c, c.auto)} (detected)</option>
                          <option value="text">Text</option><option value="number">Number</option><option value="date">Date</option>
                        </select>
                      </td>
                      <td>{c.filled} of {table.rows.length}</td>
                      <td>{c.min !== null ? `${c.min} to ${c.max}` : `${c.distinct} different values`}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>

          <section>
            <h2>Show only rows that match</h2>
            {rules.length > 1 && (
              <p>
                <select value={mode} onChange={(e) => setMode(e.target.value)} aria-label="Match mode">
                  <option value="all">Match ALL of these</option><option value="any">Match ANY of these</option>
                </select>
              </p>
            )}
            {rules.map((r) => {
              const c = table.cols[r.col]
              const t = c.type === 'empty' ? 'text' : c.type
              const kind = t === 'number' ? 'number' : t === 'date' ? 'date' : 'text'
              return (
                <div key={r.id} className="rule">
                  <select value={r.col} onChange={(e) => setRule(r.id, fresh(table.cols[+e.target.value]))} aria-label="Column">
                    {table.cols.map((x) => <option key={x.i} value={x.i}>{x.name}</option>)}
                  </select>
                  <select value={r.op} onChange={(e) => setRule(r.id, { op: e.target.value })} aria-label="Condition">
                    {opsFor(c).map(([k, l]) => <option key={k} value={k}>{l}</option>)}
                  </select>
                  {['any', 'none'].includes(r.op) ? (
                    <div className="picks">
                      {c.choices.map(([v, k]) => (
                        <label key={v} className="inline">
                          <input type="checkbox" checked={r.list.includes(v)} onChange={(e) => setRule(r.id, { list: e.target.checked ? [...r.list, v] : r.list.filter((x) => x !== v) })} />
                          {v} <span className="muted">({k})</span>
                        </label>
                      ))}
                    </div>
                  ) : !NOVAL.includes(r.op) && (
                    <>
                      <input type={kind} step="any" value={r.a} onChange={(e) => setRule(r.id, { a: e.target.value })} aria-label="Value" />
                      {r.op === 'between' && <> and <input type={kind} step="any" value={r.b} onChange={(e) => setRule(r.id, { b: e.target.value })} aria-label="Second value" /></>}
                    </>
                  )}
                  <button onClick={() => setRules(rules.filter((x) => x.id !== r.id))}>Remove</button>
                </div>
              )
            })}
            <div className="actions">
              <button onClick={() => setRules([...rules, { id: uid(), ...fresh(table.cols[0]) }])}>Add a filter</button>
              {rules.length > 0 && <button onClick={() => setRules([])}>Clear filters</button>}
            </div>
          </section>

          <section>
            <h2>Order the rows</h2>
            {sorts.map((s, k) => {
              const c = table.cols[s.col]
              const [up, down] = DIR[c.type] || DIR.text
              return (
                <div key={k} className="rule">
                  <select value={s.col} onChange={(e) => setSort(k, { col: +e.target.value })} aria-label="Sort column">
                    {table.cols.map((x) => <option key={x.i} value={x.i}>{x.name}</option>)}
                  </select>
                  <select value={s.dir} onChange={(e) => setSort(k, { dir: e.target.value })} aria-label="Direction">
                    <option value="asc">{up}</option><option value="desc">{down}</option>
                  </select>
                  <button onClick={() => setSorts(sorts.filter((_, j) => j !== k))}>Remove</button>
                </div>
              )
            })}
            <button onClick={() => setSorts([...sorts, { col: 0, dir: 'asc' }])}>{sorts.length ? 'Then by another column' : 'Sort by a column'}</button>
          </section>

          <section>
            <h2>Columns to show</h2>
            <div className="chips">
              {table.cols.map((c) => (
                <label key={c.i} className="inline">
                  <input type="checkbox" checked={!hidden.includes(c.i)} onChange={(e) => setHidden(e.target.checked ? hidden.filter((x) => x !== c.i) : [...hidden, c.i])} /> {c.name}
                </label>
              ))}
            </div>
          </section>

          <section>
            <div className="actions" style={{ justifyContent: 'space-between' }}>
              <h2 style={{ margin: 0 }}>Result: {result.rows.length} of {table.rows.length} rows</h2>
              <button className="primary" disabled={!vis.length} onClick={exportCsv}>Download result (CSV for Excel)</button>
            </div>
            {result.rows.length === 0 ? <p className="muted">No rows match. Remove or loosen a filter.</p> : (
              <div className="tbl" style={{ marginTop: 12 }}>
                <table>
                  <thead><tr>{vis.map((c) => <th key={c.i}>{c.name}</th>)}</tr></thead>
                  <tbody>
                    {result.rows.slice(0, limit).map((r, k) => (
                      <tr key={k}>{vis.map((c) => <td key={c.i} className={c.type === 'number' ? 'num' : ''}>{show(r[c.i], c)}</td>)}</tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
            {result.rows.length > limit && <p><button onClick={() => setLimit(limit + 100)}>Show 100 more</button> <span className="muted">Showing {limit}. The download includes all {result.rows.length}.</span></p>}
            {vis.some((c) => c.type === 'number') && result.rows.length > 0 && (
              <>
                <h2>Totals for the result</h2>
                <div className="tbl">
                  <table>
                    <thead><tr><th>Column</th><th>Rows with a value</th><th>Sum</th><th>Average</th></tr></thead>
                    <tbody>
                      {vis.filter((c) => c.type === 'number').map((c) => (
                        <tr key={c.i}><td>{c.name}</td><td>{result.cnt[c.i]}</td><td>{n2(result.sum[c.i])}</td><td>{result.cnt[c.i] ? n2(result.sum[c.i] / result.cnt[c.i]) : ''}</td></tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </>
            )}
          </section>
        </>
      )}
    </div>
  )
}
