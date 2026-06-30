// luckyexcel(Luckysheet 포맷) → Univer IWorkbookData 변환기
// 값 / 수식 / 병합 / 열너비·행높이 / 기본 서식(글꼴, 색, 정렬, 숫자서식)을 보존한다.

// Luckysheet 정렬 enum → Univer enum 매핑
const H_ALIGN = { 0: 2, 1: 1, 2: 3 } // lucky(0중앙,1좌,2우) → univer(1좌,2중앙,3우)
const V_ALIGN = { 0: 2, 1: 1, 2: 3 } // lucky(0중간,1상,2하) → univer(1상,2중간,3하)

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
    const keys = Object.keys(styleObj)
    if (keys.length === 0) return undefined
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

    const cellData = {}
    let maxRow = ls.row || 0
    let maxCol = ls.column || 0

    for (const cd of ls.celldata || []) {
      const { r, c, v } = cd
      if (v === null || v === undefined) continue
      if (r + 1 > maxRow) maxRow = r + 1
      if (c + 1 > maxCol) maxCol = c + 1
      const cell = {}
      const { value, t } = cellValueAndType(v)
      cell.v = value
      cell.t = t
      if (v.f) cell.f = v.f // 수식
      const styleObj = buildStyle(v)
      const sid = registerStyle(styleObj)
      if (sid) cell.s = sid
      if (!cellData[r]) cellData[r] = {}
      cellData[r][c] = cell
    }

    const config = ls.config || {}

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

    // 열 너비
    const columnData = {}
    if (config.columnlen) {
      for (const k of Object.keys(config.columnlen)) {
        columnData[k] = { w: config.columnlen[k] }
      }
    }

    // 행 높이
    const rowData = {}
    if (config.rowlen) {
      for (const k of Object.keys(config.rowlen)) {
        rowData[k] = { h: config.rowlen[k] }
      }
    }

    sheets[sheetId] = {
      id: sheetId,
      name: ls.name || `Sheet${idx + 1}`,
      rowCount: Math.max(maxRow + 50, 100),
      columnCount: Math.max(maxCol + 10, 26),
      defaultColumnWidth: 88,
      defaultRowHeight: 24,
      cellData,
      mergeData,
      columnData,
      rowData,
      freeze: { xSplit: 0, ySplit: 0, startRow: -1, startColumn: -1 }
    }
  })

  if (sheetOrder.length === 0) {
    // 빈 워크북 방지
    sheets['sheet-0'] = {
      id: 'sheet-0',
      name: 'Sheet1',
      rowCount: 100,
      columnCount: 26,
      cellData: {}
    }
    sheetOrder.push('sheet-0')
  }

  return {
    id: 'wb-' + Math.abs(hashString(fallbackName)),
    name: (luckyJson.info && luckyJson.info.name) || fallbackName,
    appVersion: '0.1.0',
    locale: 'enUS',
    sheetOrder,
    styles,
    sheets
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
