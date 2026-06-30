// luckyexcel(Luckysheet 포맷) → Univer IWorkbookData 변환기
// 값 / 수식 / 병합 / 열너비·행높이 / 서식(글꼴·색·정렬·숫자서식) / 테두리를 보존한다.

// Luckysheet 정렬 enum → Univer enum 매핑
const H_ALIGN = { 0: 2, 1: 1, 2: 3 } // lucky(0중앙,1좌,2우) → univer(1좌,2중앙,3우)
const V_ALIGN = { 0: 2, 1: 1, 2: 3 } // lucky(0중간,1상,2하) → univer(1상,2중간,3하)

// Excel 테두리 style 문자열 → Univer BorderStyleTypes
const BORDER_STYLE = {
  thin: 1, hair: 2, dotted: 3, dashed: 4, dashDot: 5, dashDotDot: 6,
  double: 7, medium: 8, mediumDashed: 9, mediumDashDot: 10,
  mediumDashDotDot: 11, slantDashDot: 12, thick: 13
}

function borderSide(side) {
  if (!side || !side.color) return undefined
  const s = BORDER_STYLE[side.style] || 1
  return { s, cl: { rgb: side.color } }
}

function buildStyle(v) {
  const s = {}
  if (v.bl) s.bl = 1
  if (v.it) s.it = 1
  if (v.un) s.ul = { s: 1 } // underline
  if (v.cl) s.st = { s: 1 } // strikethrough
  if (v.ff) s.ff = typeof v.ff === 'string' ? v.ff : undefined
  if (v.fs) s.fs = v.fs
  if (v.fc) s.cl = { rgb: v.fc }
  if (v.bg) s.bg = { rgb: v.bg }
  if (v.ht !== undefined && H_ALIGN[v.ht] !== undefined) s.ht = H_ALIGN[v.ht]
  if (v.vt !== undefined && V_ALIGN[v.vt] !== undefined) s.vt = V_ALIGN[v.vt]
  if (v.tb === 2) s.tb = 3 // 자동 줄바꿈
  if (v.ct && v.ct.fa && v.ct.fa !== 'General') s.n = { pattern: v.ct.fa }
  // 텍스트 회전: tr=3 세로쓰기, 그 외는 rt(각도). Excel 91~180° → 음수각으로 변환
  if (v.tr === 3) {
    s.tr = { a: 0, v: 1 }
  } else if (typeof v.rt === 'number' && v.rt !== 0) {
    s.tr = { a: v.rt <= 90 ? v.rt : 90 - v.rt, v: 0 }
  }
  return s
}

function cellValueAndType(v) {
  let value = v.v
  if (value === undefined || value === null) value = v.m
  let t // Univer CellValueType: 1=string,2=number,3=boolean
  if (v.ct && v.ct.t === 'n') t = 2
  else if (v.ct && v.ct.t === 'b') t = 3
  else if (typeof value === 'number') t = 2
  else t = 1
  return { value: value === undefined ? null : value, t }
}

export function luckyToUniver(luckyJson, fallbackName = 'Workbook') {
  const styles = {}
  const styleKeyToId = new Map()
  let styleSeq = 0

  function registerStyle(styleObj) {
    if (Object.keys(styleObj).length === 0) return undefined
    const key = JSON.stringify(styleObj)
    if (styleKeyToId.has(key)) return styleKeyToId.get(key)
    const id = 'st-' + styleSeq++
    styles[id] = styleObj
    styleKeyToId.set(key, id)
    return id
  }

  const sheets = {}
  const sheetOrder = []
  const luckySheets = (luckyJson.sheets || []).slice().sort((a, b) => (a.order || 0) - (b.order || 0))

  luckySheets.forEach((ls, idx) => {
    const sheetId = 'sheet-' + idx
    sheetOrder.push(sheetId)

    let maxRow = ls.row || 0
    let maxCol = ls.column || 0
    const cellMap = new Map() // "r_c" → { r, c, v, t, f, style }

    const ensureCell = (r, c) => {
      const k = r + '_' + c
      let cell = cellMap.get(k)
      if (!cell) {
        cell = { r, c, style: {} }
        cellMap.set(k, cell)
      }
      if (r + 1 > maxRow) maxRow = r + 1
      if (c + 1 > maxCol) maxCol = c + 1
      return cell
    }

    // 값 + 서식
    for (const cd of ls.celldata || []) {
      const { r, c, v } = cd
      if (v === null || v === undefined) continue
      const cell = ensureCell(r, c)
      const { value, t } = cellValueAndType(v)
      cell.v = value
      cell.t = t
      if (v.f) cell.f = v.f
      Object.assign(cell.style, buildStyle(v))
    }

    const config = ls.config || {}

    // 테두리 (셀 단위)
    for (const bi of config.borderInfo || []) {
      if (bi.rangeType === 'cell' && bi.value) {
        const val = bi.value
        const cell = ensureCell(val.row_index, val.col_index)
        const bd = {}
        const t = borderSide(val.t)
        const b = borderSide(val.b)
        const l = borderSide(val.l)
        const r = borderSide(val.r)
        if (t) bd.t = t
        if (b) bd.b = b
        if (l) bd.l = l
        if (r) bd.r = r
        if (Object.keys(bd).length) cell.style.bd = bd
      } else if (bi.rangeType === 'range' && Array.isArray(bi.range)) {
        applyRangeBorder(bi, ensureCell)
      }
    }

    // cellData 조립
    const cellData = {}
    for (const cell of cellMap.values()) {
      const out = {}
      if (cell.v !== undefined) out.v = cell.v
      if (cell.t !== undefined) out.t = cell.t
      if (cell.f) out.f = cell.f
      const sid = registerStyle(cell.style)
      if (sid) out.s = sid
      if (out.v === undefined && out.f === undefined && out.s === undefined) continue
      if (!cellData[cell.r]) cellData[cell.r] = {}
      cellData[cell.r][cell.c] = out
    }

    // 병합
    const mergeData = []
    if (config.merge) {
      for (const key of Object.keys(config.merge)) {
        const m = config.merge[key]
        mergeData.push({
          startRow: m.r,
          startColumn: m.c,
          endRow: m.r + (m.rs || 1) - 1,
          endColumn: m.c + (m.cs || 1) - 1
        })
      }
    }

    // 열 너비 / 행 높이 (luckyexcel가 픽셀로 제공)
    const columnData = {}
    if (config.columnlen) {
      for (const k of Object.keys(config.columnlen)) columnData[k] = { w: config.columnlen[k] }
    }
    const rowData = {}
    if (config.rowlen) {
      for (const k of Object.keys(config.rowlen)) rowData[k] = { h: config.rowlen[k] }
    }

    // 숨김 행/열 (luckyexcel: rowhidden/colhidden 의 키가 숨김 대상)
    if (config.rowhidden) {
      for (const k of Object.keys(config.rowhidden)) {
        rowData[k] = rowData[k] || {}
        rowData[k].hd = 1
      }
    }
    if (config.colhidden) {
      for (const k of Object.keys(config.colhidden)) {
        columnData[k] = columnData[k] || {}
        columnData[k].hd = 1
      }
    }

    const defColW = Math.round(ls.defaultColWidth || 72)
    const defRowH = Math.round(ls.defaultRowHeight || 19)
    const gridlines = Number(ls.showGridLines === undefined ? 1 : ls.showGridLines) ? 1 : 0

    sheets[sheetId] = {
      id: sheetId,
      name: ls.name || `Sheet${idx + 1}`,
      rowCount: Math.max(maxRow + 50, 100),
      columnCount: Math.max(maxCol + 10, 26),
      defaultColumnWidth: defColW,
      defaultRowHeight: defRowH,
      showGridlines: gridlines,
      cellData,
      mergeData,
      columnData,
      rowData,
      freeze: { xSplit: 0, ySplit: 0, startRow: -1, startColumn: -1 }
    }
  })

  if (sheetOrder.length === 0) {
    sheets['sheet-0'] = { id: 'sheet-0', name: 'Sheet1', rowCount: 100, columnCount: 26, cellData: {} }
    sheetOrder.push('sheet-0')
  }

  return {
    id: 'wb-' + Math.abs(hashString(fallbackName)),
    name: (luckyJson.info && luckyJson.info.name) || fallbackName,
    appVersion: '0.1.0',
    locale: 'koKR',
    sheetOrder,
    styles,
    sheets
  }
}

// 범위 테두리(border-all/outside/top/...) → 각 셀 변에 적용
function applyRangeBorder(bi, ensureCell) {
  const side = borderSide({ style: bi.style === undefined ? 'thin' : bi.style, color: bi.color || '#000000' })
  if (!side) return
  const type = bi.borderType || 'border-all'
  for (const rng of bi.range) {
    const [r1, r2] = rng.row
    const [c1, c2] = rng.column
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        const cell = ensureCell(r, c)
        const bd = cell.style.bd || (cell.style.bd = {})
        const all = type === 'border-all'
        const outside = type === 'border-outside' || type === 'border-none'
        if (all || type === 'border-top' || (outside && r === r1)) bd.t = side
        if (all || type === 'border-bottom' || (outside && r === r2)) bd.b = side
        if (all || type === 'border-left' || (outside && c === c1)) bd.l = side
        if (all || type === 'border-right' || (outside && c === c2)) bd.r = side
      }
    }
  }
}

function hashString(str) {
  let h = 0
  for (let i = 0; i < str.length; i++) {
    h = (h << 5) - h + str.charCodeAt(i)
    h |= 0
  }
  return h
}
