// 워드/PPT 등 아직 앱 내 렌더링을 지원하지 않는 형식용 폴백 뷰어
const ICONS = {
  '.docx': '📄',
  '.doc': '📄',
  '.pptx': '📊',
  '.ppt': '📊',
  '.xls': '📗',
  '.csv': '📗'
}

const LABELS = {
  '.docx': 'Word 문서',
  '.doc': 'Word 문서(구버전)',
  '.pptx': 'PowerPoint 문서',
  '.ppt': 'PowerPoint 문서(구버전)',
  '.xls': 'Excel(구버전 .xls)',
  '.csv': 'CSV'
}

export function createFallbackViewer(container, { filePath, fileName, ext }) {
  const wrap = document.createElement('div')
  wrap.className = 'placeholder'
  wrap.innerHTML = `
    <div class="big">${ICONS[ext] || '📁'}</div>
    <div>${LABELS[ext] || ext} — 앱 내 미리보기는 준비 중입니다.</div>
    <div style="color:#777;font-size:12px;">${fileName}</div>
    <button class="btn">원래 프로그램으로 열기</button>
  `
  wrap.querySelector('button').addEventListener('click', () => {
    window.api.openPath(filePath)
  })
  container.appendChild(wrap)

  return {
    dispose() {
      wrap.remove()
    }
  }
}
