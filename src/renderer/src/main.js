import './style.css'
import { initExplorer, setRoot } from './explorer.js'
import {
  initWorkspace,
  openFile,
  splitActivePane,
  saveActiveTab,
  getActiveViewer,
  openDroppedFiles,
  notify
} from './workspace.js'

const workspaceEl = document.getElementById('workspace')
const treeEl = document.getElementById('tree')
const btnOpenFolder = document.getElementById('btn-open-folder')

initWorkspace(workspaceEl)
initExplorer(treeEl, (file) => openFile(file))

// ── 파일 탐색기에서 드래그&드롭으로 열기 ──────────
const SUPPORTED_EXTS = [
  '.xlsx', '.xlsm', '.xls', '.csv',
  '.docx', '.doc', '.pptx', '.ppt', '.hwp', '.hwpx', '.pdf'
]

const dropOverlay = document.createElement('div')
dropOverlay.id = 'drop-overlay'
dropOverlay.innerHTML = '<div class="drop-box">📂 여기에 놓으면 열립니다</div>'
document.body.appendChild(dropOverlay)

let dragDepth = 0
const extOf = (name) => {
  const i = name.lastIndexOf('.')
  return i < 0 ? '' : name.slice(i).toLowerCase()
}

window.addEventListener('dragenter', (e) => {
  if (!e.dataTransfer || !Array.from(e.dataTransfer.types).includes('Files')) return
  e.preventDefault()
  dragDepth++
  dropOverlay.classList.add('show')
})
window.addEventListener('dragover', (e) => {
  if (e.dataTransfer && Array.from(e.dataTransfer.types).includes('Files')) {
    e.preventDefault()
    e.dataTransfer.dropEffect = 'copy'
  }
})
window.addEventListener('dragleave', (e) => {
  e.preventDefault()
  dragDepth = Math.max(0, dragDepth - 1)
  if (dragDepth === 0) dropOverlay.classList.remove('show')
})
window.addEventListener('drop', async (e) => {
  e.preventDefault()
  dragDepth = 0
  dropOverlay.classList.remove('show')
  const files = Array.from((e.dataTransfer && e.dataTransfer.files) || [])
  if (files.length === 0) return

  const list = []
  let skipped = 0
  for (const file of files) {
    const ext = extOf(file.name)
    if (!SUPPORTED_EXTS.includes(ext)) {
      skipped++
      continue
    }
    const path = window.api.getPathForFile(file)
    if (path) list.push({ path, name: file.name, ext })
  }
  if (list.length) await openDroppedFiles(list, e.clientX, e.clientY)
  if (skipped) notify(`${skipped}개 파일은 지원하지 않는 형식이라 제외했습니다.`)
})

// Ctrl+S → 활성 탭 저장
window.addEventListener('keydown', (e) => {
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
    e.preventDefault()
    saveActiveTab()
  }
})

// 개발용 스모크 테스트 훅
window.__ov = {
  setRoot,
  openFile,
  splitActivePane,
  saveActiveTab,
  getActiveViewer,
  openDroppedFiles
}

btnOpenFolder.addEventListener('click', async () => {
  const folder = await window.api.openFolder()
  if (folder) await setRoot(folder)
})

// ── 사이드바 너비 조절 ──────────────────
const splitter = document.getElementById('splitter-v')
const sidebar = document.getElementById('sidebar')
let dragging = false
splitter.addEventListener('mousedown', () => {
  dragging = true
  document.body.style.cursor = 'col-resize'
  document.body.style.userSelect = 'none'
})
window.addEventListener('mousemove', (e) => {
  if (!dragging) return
  const w = Math.min(560, Math.max(160, e.clientX))
  sidebar.style.width = w + 'px'
})
window.addEventListener('mouseup', () => {
  dragging = false
  document.body.style.cursor = ''
  document.body.style.userSelect = ''
})
