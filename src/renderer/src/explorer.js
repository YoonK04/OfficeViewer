// 왼쪽 파일 탐색기 (지연 로딩 트리)
const FILE_ICONS = {
  '.xlsx': '📗',
  '.xlsm': '📗',
  '.xls': '📗',
  '.csv': '📗',
  '.docx': '📄',
  '.doc': '📄',
  '.pptx': '📊',
  '.ppt': '📊'
}

let onOpenFile = null
let treeEl = null
let selectedRow = null

export function initExplorer(el, openFileCb) {
  treeEl = el
  onOpenFile = openFileCb
}

export async function setRoot(rootPath) {
  treeEl.innerHTML = ''
  const rootName = rootPath.split(/[\\/]/).filter(Boolean).pop() || rootPath
  const node = buildDirNode({ name: rootName, path: rootPath }, 0)
  treeEl.appendChild(node.el)
  await expandDir(node)
}

function buildDirNode(entry, depth) {
  const node = { entry, depth, expanded: false, loaded: false, childrenEl: null }

  const wrap = document.createElement('div')
  wrap.className = 'tree-node'

  const row = document.createElement('div')
  row.className = 'tree-row'
  row.style.paddingLeft = 6 + depth * 14 + 'px'

  const twisty = document.createElement('span')
  twisty.className = 'tree-twisty'
  twisty.textContent = '▶'

  const icon = document.createElement('span')
  icon.className = 'tree-icon'
  icon.textContent = '📁'

  const label = document.createElement('span')
  label.className = 'tree-label'
  label.textContent = entry.name

  row.appendChild(twisty)
  row.appendChild(icon)
  row.appendChild(label)

  const children = document.createElement('div')
  children.className = 'tree-children'
  children.style.display = 'none'

  wrap.appendChild(row)
  wrap.appendChild(children)

  node.el = wrap
  node.childrenEl = children
  node.twisty = twisty
  node.icon = icon

  row.addEventListener('click', () => toggleDir(node))
  return node
}

function buildFileNode(entry, depth) {
  const wrap = document.createElement('div')
  wrap.className = 'tree-node'

  const row = document.createElement('div')
  row.className = 'tree-row'
  row.style.paddingLeft = 6 + depth * 14 + 'px'

  const twisty = document.createElement('span')
  twisty.className = 'tree-twisty'
  twisty.textContent = ''

  const icon = document.createElement('span')
  icon.className = 'tree-icon'
  icon.textContent = FILE_ICONS[entry.ext] || '📄'

  const label = document.createElement('span')
  label.className = 'tree-label'
  label.textContent = entry.name

  row.appendChild(twisty)
  row.appendChild(icon)
  row.appendChild(label)
  wrap.appendChild(row)

  row.addEventListener('click', () => {
    if (selectedRow) selectedRow.classList.remove('selected')
    row.classList.add('selected')
    selectedRow = row
    onOpenFile({ path: entry.path, name: entry.name, ext: entry.ext })
  })
  return { el: wrap }
}

async function toggleDir(node) {
  if (node.expanded) {
    node.expanded = false
    node.twisty.textContent = '▶'
    node.icon.textContent = '📁'
    node.childrenEl.style.display = 'none'
  } else {
    await expandDir(node)
  }
}

async function expandDir(node) {
  node.expanded = true
  node.twisty.textContent = '▼'
  node.icon.textContent = '📂'
  node.childrenEl.style.display = ''
  if (node.loaded) return

  try {
    const entries = await window.api.readDir(node.entry.path)
    for (const e of entries) {
      if (e.type === 'dir') {
        const child = buildDirNode(e, node.depth + 1)
        node.childrenEl.appendChild(child.el)
      } else {
        const child = buildFileNode(e, node.depth + 1)
        node.childrenEl.appendChild(child.el)
      }
    }
    if (entries.length === 0) {
      const empty = document.createElement('div')
      empty.className = 'tree-empty'
      empty.style.paddingLeft = 6 + (node.depth + 1) * 14 + 'px'
      empty.textContent = '(오피스 문서 없음)'
      node.childrenEl.appendChild(empty)
    }
    node.loaded = true
  } catch (err) {
    console.error('[explorer] 디렉터리 읽기 실패:', err)
  }
}
