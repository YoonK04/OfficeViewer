// 워드(.docx) 미리보기 — docx-preview (보기 전용)
import { renderAsync } from 'docx-preview'

export async function createWordViewer(container, { arrayBuffer }) {
  const scroll = document.createElement('div')
  scroll.className = 'doc-scroll'
  container.appendChild(scroll)

  try {
    await renderAsync(arrayBuffer, scroll, undefined, {
      className: 'docx',
      inWrapper: true,
      ignoreWidth: false,
      ignoreHeight: false,
      breakPages: true,
      experimental: true,
      useBase64URL: true
    })
  } catch (err) {
    console.error('[word] 렌더 실패:', err)
    scroll.innerHTML =
      '<div class="placeholder"><div class="big">⚠️</div><div>워드 미리보기를 만들 수 없습니다.</div></div>'
  }

  return {
    dispose() {
      scroll.remove()
    }
  }
}
