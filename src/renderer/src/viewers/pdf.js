// PDF 뷰어 — pdf.js(pdfjs-dist)로 각 페이지를 캔버스에 렌더링.
// 오프라인·자체 렌더(플러그인 의존 없음), 화면에 보이는 페이지만 렌더(메모리 안전).
// legacy 빌드: 구형 Chromium(Electron 31)용 폴리필(Promise.try 등) 포함
import * as pdfjsLib from 'pdfjs-dist/legacy/build/pdf.mjs'
import PdfWorker from 'pdfjs-dist/legacy/build/pdf.worker.mjs?worker'

pdfjsLib.GlobalWorkerOptions.workerPort = new PdfWorker()

export function createPdfViewer(container, { arrayBuffer }) {
  const scroll = document.createElement('div')
  scroll.className = 'doc-scroll pdf-scroll'
  container.appendChild(scroll)

  const loading = document.createElement('div')
  loading.className = 'placeholder'
  loading.innerHTML = '<div class="spinner"></div><div>PDF 여는 중…</div>'
  container.appendChild(loading)

  let pdf = null
  let disposed = false
  let observer = null
  const rendered = new Map() // pageNum → canvas

  ;(async () => {
    try {
      // arrayBuffer가 detach될 수 있으니 복사본 사용
      const data = new Uint8Array(arrayBuffer.slice(0))
      pdf = await pdfjsLib.getDocument({ data }).promise
      if (disposed) return
      loading.remove()

      const width = Math.min(1100, Math.max(320, (container.clientWidth || 900) - 48))

      const pageEls = []
      for (let n = 1; n <= pdf.numPages; n++) {
        const page = await pdf.getPage(n)
        if (disposed) return
        const base = page.getViewport({ scale: 1 })
        const scale = width / base.width
        const vp = page.getViewport({ scale })
        const holder = document.createElement('div')
        holder.className = 'pdf-page'
        holder.style.width = vp.width + 'px'
        holder.style.height = vp.height + 'px'
        holder.dataset.page = String(n)
        scroll.appendChild(holder)
        pageEls.push({ holder, page, scale })
      }

      // 보이는(근처) 페이지만 렌더
      observer = new IntersectionObserver(
        (entries) => {
          for (const e of entries) {
            const n = parseInt(e.target.dataset.page, 10)
            if (e.isIntersecting) renderPage(n)
          }
        },
        { root: scroll, rootMargin: '600px 0px' }
      )
      pageEls.forEach(({ holder }) => observer.observe(holder))

      async function renderPage(n) {
        if (rendered.has(n) || disposed) return
        const info = pageEls[n - 1]
        if (!info) return
        rendered.set(n, true)
        const dpr = window.devicePixelRatio || 1
        const vp = info.page.getViewport({ scale: info.scale * dpr })
        const canvas = document.createElement('canvas')
        canvas.width = vp.width
        canvas.height = vp.height
        canvas.style.width = '100%'
        canvas.style.height = '100%'
        info.holder.appendChild(canvas)
        try {
          await info.page.render({ canvasContext: canvas.getContext('2d'), viewport: vp }).promise
        } catch (err) {
          if (!disposed) console.error('[pdf] 페이지 렌더 실패:', n, err)
        }
      }
    } catch (err) {
      console.error('[pdf] 로드 실패:', err)
      if (!disposed) {
        loading.innerHTML = '<div class="big">⚠️</div><div>PDF를 열 수 없습니다.</div>'
      }
    }
  })()

  return {
    dispose() {
      disposed = true
      if (observer) observer.disconnect()
      try {
        if (pdf) pdf.destroy()
      } catch (e) {
        /* noop */
      }
      scroll.remove()
      loading.remove()
    }
  }
}
