// PPT(.pptx) 미리보기 — pptx-preview (보기 전용, 오프라인 베스트에포트)
import { init as initPptx } from 'pptx-preview'

export async function createPptViewer(container, { arrayBuffer }) {
  const scroll = document.createElement('div')
  scroll.className = 'doc-scroll ppt-scroll'
  container.appendChild(scroll)

  // 컨테이너 너비에 맞춰 16:9 기준 크기 산정
  const cw = container.clientWidth || 960
  const width = Math.min(1280, Math.max(480, cw - 48))
  const height = Math.round((width * 9) / 16)

  try {
    const previewer = initPptx(scroll, { width, height })
    await previewer.preview(arrayBuffer)
  } catch (err) {
    console.error('[ppt] 렌더 실패:', err)
    scroll.innerHTML =
      '<div class="placeholder"><div class="big">⚠️</div><div>PPT 미리보기를 만들 수 없습니다.</div></div>'
  }

  return {
    dispose() {
      scroll.remove()
    }
  }
}
