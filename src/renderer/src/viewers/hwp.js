// HWP(.hwp 5.0) 미리보기 — hwp.js (보기 전용, 베스트에포트)
// hwp.js는 일부 Node 의존성이 있어 렌더러에서 실패할 수 있으므로
// 동적 import + try/catch로 격리한다. 실패하면 "원래 프로그램으로 열기" 폴백.
// HWPX(zip형)·복잡한 표/이미지/정밀 레이아웃은 지원이 약하다.

export async function createHwpViewer(container, { arrayBuffer, filePath, fileName, ext }) {
  if (ext === '.hwpx') {
    return fallback(container, filePath, fileName, 'HWPX 형식은 앱 내 미리보기를 지원하지 않습니다.')
  }

  const host = document.createElement('div')
  host.className = 'doc-scroll hwp-scroll'
  container.appendChild(host)

  try {
    const mod = await import('hwp.js')
    const HwpViewer = mod.Viewer || (mod.default && mod.default.Viewer)
    if (!HwpViewer) throw new Error('hwp.js Viewer 로드 실패')
    const data = new Uint8Array(arrayBuffer)
    // eslint-disable-next-line no-new
    new HwpViewer(host, data, { type: 'binary' })
    return { dispose() { host.remove() } }
  } catch (err) {
    console.error('[hwp] 렌더 실패:', err)
    host.remove()
    return fallback(
      container,
      filePath,
      fileName,
      'HWP 미리보기를 만들 수 없습니다(복잡한 문서이거나 미지원).'
    )
  }
}

function fallback(container, filePath, fileName, msg) {
  const wrap = document.createElement('div')
  wrap.className = 'placeholder'
  wrap.innerHTML = `
    <div class="big">📕</div>
    <div>${msg}</div>
    <div style="color:#777;font-size:12px;">${fileName}</div>
    <button class="btn">원래 프로그램으로 열기</button>
  `
  wrap.querySelector('button').addEventListener('click', () => window.api.openPath(filePath))
  container.appendChild(wrap)
  return { dispose() { wrap.remove() } }
}
