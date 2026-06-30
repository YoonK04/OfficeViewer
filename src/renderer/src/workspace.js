// 워크스페이스: 분할 가능한 패널 + 패널별 탭 + 지연 로딩 뷰어
import { createExcelViewer } from './viewers/excel.js'
import { createWordViewer } from './viewers/word.js'
import { createPptViewer } from './viewers/ppt.js'
import { createFallbackViewer } from './viewers/fallback.js'

const EXCEL_EXTS = ['.xlsx', '.xlsm']
const WORD_EXTS = ['.docx']
const PPT_EXTS = ['.pptx']

let workspaceEl = null
const panes = []
let activePaneId = null
let seq = 0
const uid = (p) => `${p}-${seq++}`

export function initWorkspace(el) {
  workspaceEl = el
  addPane() // 시작 패널 1개
}

function addPane(afterPane) {
  const pane = {
    id: uid('pane'),
    el: null,
    tabbarEl: null,
    bodyEl: null,
    tabs: [],
    activeTabId: null
  }

  const paneEl = document.createElement('div')
  paneEl.className = 'pane'
  paneEl.dataset.paneId = pane.id

  const tabbar = document.createElement('div')
  tabbar.className = 'tabbar'

  const body = document.createElement('div')
  body.className = 'pane-body'

  paneEl.appendChild(tabbar)
  paneEl.appendChild(body)

  pane.el = paneEl
  pane.tabbarEl = tabbar
  pane.bodyEl = body

  paneEl.addEventListener('mousedown', () => setActivePane(pane.id))

  paneEl.style.flex = '1 1 0'

  // afterPane이 주어지면 그 패널 바로 오른쪽에 삽입, 아니면 맨 끝에 추가
  if (afterPane && afterPane.el) {
    afterPane.el.after(paneEl)
    const i = panes.indexOf(afterPane)
    panes.splice(i + 1, 0, pane)
  } else {
    workspaceEl.appendChild(paneEl)
    panes.push(pane)
  }
  setActivePane(pane.id)
  renderTabbar(pane)
  renderBody(pane)
  layoutSplitters()
  return pane
}

// 패널 사이에 드래그 가능한 경계를 (재)배치
function layoutSplitters() {
  workspaceEl.querySelectorAll('.pane-splitter').forEach((s) => s.remove())
  for (let i = 0; i < panes.length - 1; i++) {
    const left = panes[i]
    const right = panes[i + 1]
    const sp = document.createElement('div')
    sp.className = 'pane-splitter'
    left.el.after(sp)
    attachPaneDrag(sp, left, right)
  }
}

function attachPaneDrag(sp, left, right) {
  sp.addEventListener('mousedown', (e) => {
    e.preventDefault()
    const startX = e.clientX
    const wL = left.el.getBoundingClientRect().width
    const wR = right.el.getBoundingClientRect().width
    const gL = parseFloat(left.el.style.flexGrow || '1')
    const gR = parseFloat(right.el.style.flexGrow || '1')
    const totalG = gL + gR
    const totalW = wL + wR
    document.body.style.cursor = 'col-resize'
    document.body.style.userSelect = 'none'

    const onMove = (ev) => {
      const dx = ev.clientX - startX
      let newL = Math.min(totalW - 120, Math.max(120, wL + dx))
      const ratio = newL / totalW
      left.el.style.flexGrow = String(totalG * ratio)
      right.el.style.flexGrow = String(totalG * (1 - ratio))
    }
    const onUp = () => {
      document.body.style.cursor = ''
      document.body.style.userSelect = ''
      window.removeEventListener('mousemove', onMove)
      window.removeEventListener('mouseup', onUp)
    }
    window.addEventListener('mousemove', onMove)
    window.addEventListener('mouseup', onUp)
  })
}

function getPane(id) {
  return panes.find((p) => p.id === id)
}
function getActivePane() {
  return getPane(activePaneId) || panes[0]
}

function setActivePane(id) {
  activePaneId = id
  panes.forEach((p) => p.el.classList.toggle('active', p.id === id))
}

export function splitActivePane() {
  const newPane = addPane(getActivePane())
  return newPane
}

// 활성 패널의 활성 탭 뷰어 (개발/디버그용)
export function getActiveViewer() {
  const pane = getActivePane()
  const tab = pane.tabs.find((t) => t.id === pane.activeTabId)
  return tab ? tab.viewer : null
}

// 활성 패널의 활성 탭 저장 (엑셀만 지원)
export async function saveActiveTab() {
  const pane = getActivePane()
  const tab = pane.tabs.find((t) => t.id === pane.activeTabId)
  if (!tab || !tab.viewer || typeof tab.viewer.save !== 'function') return
  try {
    toast('저장 중…', true)
    const buf = await tab.viewer.save()
    await window.api.writeFile(tab.filePath, buf)
    tab.dirty = false
    renderTabbar(pane)
    toast('저장됨: ' + tab.fileName)
  } catch (err) {
    console.error('[workspace] 저장 실패:', err)
    toast('저장 실패: ' + (err.message || err))
  }
}

let toastTimer = null
function toast(msg, persist) {
  let el = document.getElementById('ov-toast')
  if (!el) {
    el = document.createElement('div')
    el.id = 'ov-toast'
    el.className = 'toast'
    document.body.appendChild(el)
  }
  el.textContent = msg
  el.classList.add('show')
  if (toastTimer) clearTimeout(toastTimer)
  if (!persist) toastTimer = setTimeout(() => el.classList.remove('show'), 1800)
}

function closePane(pane) {
  if (panes.length <= 1) return // 마지막 패널은 유지
  pane.tabs.forEach((t) => disposeTab(t))
  pane.el.remove()
  const idx = panes.indexOf(pane)
  panes.splice(idx, 1)
  panes.forEach((p) => (p.el.style.flexGrow = '1'))
  setActivePane(panes[Math.max(0, idx - 1)].id)
  layoutSplitters()
}

// 파일을 활성 패널에 연다 (이미 열려 있으면 해당 탭으로 이동)
export async function openFile({ path, name, ext }) {
  const pane = getActivePane()
  const existing = pane.tabs.find((t) => t.filePath === path)
  if (existing) {
    activateTab(pane, existing.id)
    return
  }
  const tab = {
    id: uid('tab'),
    filePath: path,
    fileName: name,
    ext,
    viewer: null,
    contentEl: null,
    loaded: false
  }
  pane.tabs.push(tab)
  renderTabbar(pane)
  await activateTab(pane, tab.id)
}

async function activateTab(pane, tabId) {
  pane.activeTabId = tabId
  renderTabbar(pane)
  // 모든 탭 콘텐츠 숨기고 활성 탭만 표시
  pane.tabs.forEach((t) => {
    if (t.contentEl) t.contentEl.classList.toggle('hidden', t.id !== tabId)
  })
  renderBody(pane)
  const tab = pane.tabs.find((t) => t.id === tabId)
  if (tab && !tab.loaded) {
    await loadViewer(pane, tab)
  }
}

async function loadViewer(pane, tab) {
  // 콘텐츠 컨테이너 생성 (보이는 상태로 — Univer 사이즈 계산 위해)
  const content = document.createElement('div')
  content.className = 'viewer'
  pane.bodyEl.appendChild(content)
  tab.contentEl = content

  // 로딩 스피너
  const loading = document.createElement('div')
  loading.className = 'placeholder'
  loading.innerHTML = '<div class="spinner"></div><div>여는 중…</div>'
  content.appendChild(loading)

  try {
    if (EXCEL_EXTS.includes(tab.ext)) {
      const arrayBuffer = await window.api.readFile(tab.filePath)
      loading.remove()
      tab.viewer = createExcelViewer(content, {
        arrayBuffer,
        fileName: tab.fileName,
        onDirtyChange: (isDirty) => {
          tab.dirty = isDirty
          renderTabbar(pane)
        }
      })
    } else if (WORD_EXTS.includes(tab.ext)) {
      const arrayBuffer = await window.api.readFile(tab.filePath)
      loading.remove()
      tab.viewer = await createWordViewer(content, { arrayBuffer, fileName: tab.fileName })
    } else if (PPT_EXTS.includes(tab.ext)) {
      const arrayBuffer = await window.api.readFile(tab.filePath)
      loading.remove()
      tab.viewer = await createPptViewer(content, { arrayBuffer, fileName: tab.fileName })
    } else {
      loading.remove()
      tab.viewer = createFallbackViewer(content, {
        filePath: tab.filePath,
        fileName: tab.fileName,
        ext: tab.ext
      })
    }
    tab.loaded = true
  } catch (err) {
    console.error('[workspace] 뷰어 로드 실패:', err)
    loading.innerHTML = `<div class="big">⚠️</div><div>열기 실패: ${err.message || err}</div>`
  }
}

function disposeTab(tab) {
  if (tab.viewer && tab.viewer.dispose) {
    try {
      tab.viewer.dispose()
    } catch (e) {
      /* noop */
    }
  }
  if (tab.contentEl) tab.contentEl.remove()
}

function closeTab(pane, tabId) {
  const idx = pane.tabs.findIndex((t) => t.id === tabId)
  if (idx === -1) return
  const tab = pane.tabs[idx]
  disposeTab(tab)
  pane.tabs.splice(idx, 1)
  if (pane.activeTabId === tabId) {
    const next = pane.tabs[idx] || pane.tabs[idx - 1]
    pane.activeTabId = next ? next.id : null
    if (next) {
      pane.tabs.forEach((t) => {
        if (t.contentEl) t.contentEl.classList.toggle('hidden', t.id !== next.id)
      })
    }
  }
  renderTabbar(pane)
  renderBody(pane)
}

// ── 탭 드래그 (이동 / 우측 분할) ──────────────
const DRAG_THRESHOLD = 5

function beginTabDrag(e, fromPane, tab, tabEl) {
  const startX = e.clientX
  const startY = e.clientY
  let started = false
  let ghost = null
  let indicator = null
  let lastTarget = null

  const onMove = (ev) => {
    if (!started) {
      if (Math.abs(ev.clientX - startX) < DRAG_THRESHOLD && Math.abs(ev.clientY - startY) < DRAG_THRESHOLD) return
      started = true
      document.body.style.cursor = 'grabbing'
      document.body.style.userSelect = 'none'
      tabEl.classList.add('dragging')
      ghost = document.createElement('div')
      ghost.className = 'tab-ghost'
      ghost.textContent = tab.fileName
      document.body.appendChild(ghost)
      indicator = document.createElement('div')
      indicator.className = 'drop-indicator'
      document.body.appendChild(indicator)
    }
    ghost.style.left = ev.clientX + 12 + 'px'
    ghost.style.top = ev.clientY + 12 + 'px'
    lastTarget = resolveDropTarget(ev.clientX, ev.clientY)
    paintIndicator(indicator, lastTarget)
  }

  const onUp = () => {
    window.removeEventListener('mousemove', onMove)
    window.removeEventListener('mouseup', onUp)
    document.body.style.cursor = ''
    document.body.style.userSelect = ''
    tabEl.classList.remove('dragging')
    if (ghost) ghost.remove()
    if (indicator) indicator.remove()

    if (!started) {
      // 단순 클릭 → 활성화
      setActivePane(fromPane.id)
      activateTab(fromPane, tab.id)
      return
    }
    if (!lastTarget) return
    if (lastTarget.mode === 'split-right') {
      const np = addPane(lastTarget.pane)
      moveTab(tab, fromPane, np)
    } else if (lastTarget.pane !== fromPane) {
      moveTab(tab, fromPane, lastTarget.pane)
    } else {
      setActivePane(fromPane.id)
      activateTab(fromPane, tab.id)
    }
  }

  window.addEventListener('mousemove', onMove)
  window.addEventListener('mouseup', onUp)
}

// 커서 위치 → 어느 패널의 어떤 동작인지 판정
function resolveDropTarget(x, y) {
  const el = document.elementFromPoint(x, y)
  if (!el) return null
  const paneEl = el.closest('.pane')
  if (!paneEl) return null
  const pane = panes.find((p) => p.id === paneEl.dataset.paneId)
  if (!pane) return null
  const paneRect = paneEl.getBoundingClientRect()
  const tabbarRect = pane.tabbarEl.getBoundingClientRect()
  if (y <= tabbarRect.bottom) return { pane, mode: 'merge', rect: paneRect, tabbarRect }
  if (x >= paneRect.right - paneRect.width * 0.35) return { pane, mode: 'split-right', rect: paneRect }
  return { pane, mode: 'merge', rect: paneRect, tabbarRect }
}

function paintIndicator(indicator, target) {
  if (!target) {
    indicator.style.display = 'none'
    return
  }
  indicator.style.display = 'block'
  const r = target.rect
  if (target.mode === 'split-right') {
    indicator.style.left = r.left + r.width / 2 + 'px'
    indicator.style.top = r.top + 'px'
    indicator.style.width = r.width / 2 + 'px'
    indicator.style.height = r.height + 'px'
  } else {
    indicator.style.left = r.left + 'px'
    indicator.style.top = r.top + 'px'
    indicator.style.width = r.width + 'px'
    indicator.style.height = r.height + 'px'
  }
}

// 탭을 from 패널에서 to 패널로 이동 (뷰어 DOM 보존)
function moveTab(tab, from, to) {
  if (from === to) return
  const idx = from.tabs.indexOf(tab)
  if (idx === -1) return
  from.tabs.splice(idx, 1)
  if (tab.contentEl) to.bodyEl.appendChild(tab.contentEl)
  to.tabs.push(tab)

  // 원래 패널의 활성 탭 정리
  if (from.activeTabId === tab.id) {
    const n = from.tabs[idx] || from.tabs[idx - 1]
    from.activeTabId = n ? n.id : null
  }
  from.tabs.forEach((t) => {
    if (t.contentEl) t.contentEl.classList.toggle('hidden', t.id !== from.activeTabId)
  })
  renderTabbar(from)
  renderBody(from)

  setActivePane(to.id)
  activateTab(to, tab.id) // 대상 패널에서 표시(+ 미로딩 시 로드)

  if (from.tabs.length === 0 && panes.length > 1) closePane(from)
}

function renderTabbar(pane) {
  const bar = pane.tabbarEl
  bar.innerHTML = ''

  pane.tabs.forEach((tab) => {
    const tabEl = document.createElement('div')
    tabEl.className = 'tab' + (tab.id === pane.activeTabId ? ' active' : '')
    tabEl.title = tab.filePath

    const label = document.createElement('span')
    label.className = 'tab-label'
    label.textContent = (tab.dirty ? '● ' : '') + tab.fileName

    const close = document.createElement('span')
    close.className = 'tab-close'
    close.textContent = '×'
    close.addEventListener('click', (e) => {
      e.stopPropagation()
      closeTab(pane, tab.id)
    })

    tabEl.appendChild(label)
    tabEl.appendChild(close)
    // 마우스다운 → (이동 적으면)탭 활성화 / (드래그)탭 이동·분할
    tabEl.addEventListener('mousedown', (e) => {
      if (e.button !== 0) return
      if (e.target.classList.contains('tab-close')) return
      beginTabDrag(e, pane, tab, tabEl)
    })
    bar.appendChild(tabEl)
  })

  // 우측 도구: 분할 / 패널 닫기
  const spacer = document.createElement('div')
  spacer.className = 'tab-spacer'
  bar.appendChild(spacer)

  const tools = document.createElement('div')
  tools.className = 'pane-tools'

  const splitBtn = document.createElement('div')
  splitBtn.className = 'pane-tool'
  splitBtn.title = '오른쪽으로 분할'
  splitBtn.textContent = '⊞'
  splitBtn.addEventListener('click', (e) => {
    e.stopPropagation()
    splitActivePane()
  })
  tools.appendChild(splitBtn)

  if (panes.length > 1) {
    const closePaneBtn = document.createElement('div')
    closePaneBtn.className = 'pane-tool'
    closePaneBtn.title = '이 패널 닫기'
    closePaneBtn.textContent = '⊟'
    closePaneBtn.addEventListener('click', (e) => {
      e.stopPropagation()
      closePane(pane)
    })
    tools.appendChild(closePaneBtn)
  }

  bar.appendChild(tools)
}

function renderBody(pane) {
  // 탭이 없으면 빈 안내 표시
  const existingEmpty = pane.bodyEl.querySelector('.empty-pane')
  if (pane.tabs.length === 0) {
    if (!existingEmpty) {
      const empty = document.createElement('div')
      empty.className = 'empty-pane'
      empty.textContent = '왼쪽에서 파일을 선택하세요.'
      pane.bodyEl.appendChild(empty)
    }
  } else if (existingEmpty) {
    existingEmpty.remove()
  }
}
