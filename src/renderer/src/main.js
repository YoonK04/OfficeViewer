import './style.css'
import { initExplorer, setRoot } from './explorer.js'
import {
  initWorkspace,
  openFile,
  splitActivePane,
  saveActiveTab,
  getActiveViewer
} from './workspace.js'

const workspaceEl = document.getElementById('workspace')
const treeEl = document.getElementById('tree')
const btnOpenFolder = document.getElementById('btn-open-folder')

initWorkspace(workspaceEl)
initExplorer(treeEl, (file) => openFile(file))

// Ctrl+S → 활성 탭 저장
window.addEventListener('keydown', (e) => {
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
    e.preventDefault()
    saveActiveTab()
  }
})

// 개발용 스모크 테스트 훅
window.__ov = { setRoot, openFile, splitActivePane, saveActiveTab, getActiveViewer }

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
