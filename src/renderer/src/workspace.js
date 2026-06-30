// 워크스페이스: 분할 가능한 패널 + 패널별 탭 + 지연 로딩 뷰어
import { createExcelViewer } from './viewers/excel.js'
import { createFallbackViewer } from './viewers/fallback.js'

const EXCEL_EXTS = ['.xlsx', '.xlsm']

let workspaceEl = null
const panes = []
let activePaneId = null
let seq = 0
const uid = (p) => `${p}-${seq++}`

export function initWorkspace(el) {
  workspaceEl = el
  addPane() // 시작 패널 1개
}

function addPane() {
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
  workspaceEl.appendChild(paneEl)

  pane.el = paneEl
  pane.tabbarEl = tabbar
  pane.bodyEl = body

  paneEl.addEventListener('mousedown', () => setActivePane(pane.id))

  panes.push(pane)
  setActivePane(pane.id)
  renderTabbar(pane)
  renderBody(pane)
  return pane
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
  const newPane = addPane()
  return newPane
}

function closePane(pane) {
  if (panes.length <= 1) return // 마지막 패널은 유지
  pane.tabs.forEach((t) => disposeTab(t))
  pane.el.remove()
  const idx = panes.indexOf(pane)
  panes.splice(idx, 1)
  setActivePane(panes[Math.max(0, idx - 1)].id)
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
      tab.viewer = createExcelViewer(content, { arrayBuffer, fileName: tab.fileName })
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

function renderTabbar(pane) {
  const bar = pane.tabbarEl
  bar.innerHTML = ''

  pane.tabs.forEach((tab) => {
    const tabEl = document.createElement('div')
    tabEl.className = 'tab' + (tab.id === pane.activeTabId ? ' active' : '')
    tabEl.title = tab.filePath

    const label = document.createElement('span')
    label.className = 'tab-label'
    label.textContent = tab.fileName

    const close = document.createElement('span')
    close.className = 'tab-close'
    close.textContent = '×'
    close.addEventListener('click', (e) => {
      e.stopPropagation()
      closeTab(pane, tab.id)
    })

    tabEl.appendChild(label)
    tabEl.appendChild(close)
    tabEl.addEventListener('click', () => {
      setActivePane(pane.id)
      activateTab(pane, tab.id)
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
