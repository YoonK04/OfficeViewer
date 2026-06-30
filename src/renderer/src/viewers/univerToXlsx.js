// Univer IWorkbookData 스냅샷 → xlsx(ArrayBuffer) 변환 (ExcelJS, 서식 보존 저장)
import ExcelJS from 'exceljs'

// Univer 정렬 enum → ExcelJS 문자열
const H = { 1: 'left', 2: 'center', 3: 'right' }
const V = { 1: 'top', 2: 'middle', 3: 'bottom' }

// '#rrggbb' | 'rgb(r,g,b)' → 'FFRRGGBB'
function toARGB(color) {
  if (!color) return undefined
  let rgb = typeof color === 'string' ? color : color.rgb
  if (!rgb) return undefined
  rgb = String(rgb).trim()
  if (rgb.startsWith('#')) {
    let hex = rgb.slice(1)
    if (hex.length === 3) hex = hex.split('').map((c) => c + c).join('')
    if (hex.length === 6) return 'FF' + hex.toUpperCase()
    if (hex.length === 8) return hex.toUpperCase()
    return undefined
  }
  const m = rgb.match(/rgba?\(([^)]+)\)/)
  if (m) {
    const parts = m[1].split(',').map((s) => parseFloat(s.trim()))
    const [r, g, b] = parts
    const h = (n) => Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, '0')
    return ('FF' + h(r) + h(g) + h(b)).toUpperCase()
  }
  return undefined
}

function resolveStyle(s, stylesMap) {
  if (!s) return null
  if (typeof s === 'string') return stylesMap[s] || null
  return s // 인라인 스타일 객체
}

function applyStyle(cell, st) {
  if (!st) return

  const font = {}
  if (st.ff) font.name = st.ff
  if (st.fs) font.size = st.fs
  if (st.bl) font.bold = true
  if (st.it) font.italic = true
  if (st.ul && (st.ul.s || st.ul === 1)) font.underline = true
  if (st.st && (st.st.s || st.st === 1)) font.strike = true
  const fc = toARGB(st.cl)
  if (fc) font.color = { argb: fc }
  if (Object.keys(font).length) cell.font = font

  const bg = toARGB(st.bg)
  if (bg) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } }

  const align = {}
  if (st.ht && H[st.ht]) align.horizontal = H[st.ht]
  if (st.vt && V[st.vt]) align.vertical = V[st.vt]
  if (st.tb === 3) align.wrapText = true
  if (Object.keys(align).length) cell.alignment = align

  if (st.n && st.n.pattern) cell.numFmt = st.n.pattern

  if (st.bd) {
    const side = (b) => (b ? { style: 'thin', color: { argb: toARGB(b.cl) || 'FFBFBFBF' } } : undefined)
    const border = {}
    if (st.bd.t) border.top = side(st.bd.t)
    if (st.bd.b) border.bottom = side(st.bd.b)
    if (st.bd.l) border.left = side(st.bd.l)
    if (st.bd.r) border.right = side(st.bd.r)
    if (Object.keys(border).length) cell.border = border
  }
}

function cellValue(c) {
  if (c.f) {
    const formula = String(c.f).replace(/^=/, '')
    return { formula, result: c.v }
  }
  if (c.v === undefined || c.v === null) return null
  return c.v
}

export async function univerSnapshotToXlsx(snapshot) {
  const wb = new ExcelJS.Workbook()
  const stylesMap = snapshot.styles || {}
  const order = snapshot.sheetOrder || Object.keys(snapshot.sheets || {})

  for (const sheetId of order) {
    const sheet = snapshot.sheets[sheetId]
    if (!sheet) continue
    const ws = wb.addWorksheet(sheet.name || 'Sheet')

    const cellData = sheet.cellData || {}
    for (const rKey of Object.keys(cellData)) {
      const r = parseInt(rKey, 10)
      const rowObj = cellData[rKey]
      for (const cKey of Object.keys(rowObj)) {
        const c = parseInt(cKey, 10)
        const cd = rowObj[cKey]
        if (cd == null) continue
        const cell = ws.getCell(r + 1, c + 1)
        const val = cellValue(cd)
        if (val !== null) cell.value = val
        applyStyle(cell, resolveStyle(cd.s, stylesMap))
      }
    }

    // 병합
    for (const m of sheet.mergeData || []) {
      try {
        ws.mergeCells(m.startRow + 1, m.startColumn + 1, m.endRow + 1, m.endColumn + 1)
      } catch (e) {
        /* 겹치는 병합 무시 */
      }
    }

    // 열 너비 (px → Excel 문자단위 근사)
    const columnData = sheet.columnData || {}
    for (const cKey of Object.keys(columnData)) {
      const w = columnData[cKey].w
      if (w) ws.getColumn(parseInt(cKey, 10) + 1).width = Math.max(2, w / 7)
    }

    // 행 높이 (px → pt)
    const rowData = sheet.rowData || {}
    for (const rKey of Object.keys(rowData)) {
      const h = rowData[rKey].h
      if (h) ws.getRow(parseInt(rKey, 10) + 1).height = h * 0.75
    }
  }

  if (wb.worksheets.length === 0) wb.addWorksheet('Sheet1')

  const buf = await wb.xlsx.writeBuffer()
  return buf // ArrayBuffer
}
