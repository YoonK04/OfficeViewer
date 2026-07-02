// Web Worker: xlsx/xlsm 바이트 → Luckysheet exportJson (luckyexcel)
// 파싱이 무거워 UI 스레드를 막으므로 워커에서 처리한다.
// luckyexcel은 window/document를 참조하는 코드 경로가 있으나(렌더/다운로드용)
// transformExcelToLucky(file) 경로에서는 쓰지 않으므로 최소 shim만 제공한다.

self.window = self
if (typeof self.document === 'undefined') {
  self.document = {
    createElement: () => ({ style: {}, getContext: () => null, appendChild: () => {} }),
    getElementById: () => null,
    body: { appendChild: () => {} }
  }
}

self.onmessage = async (e) => {
  const { arrayBuffer, fileName } = e.data
  try {
    const mod = await import('luckyexcel')
    const LuckyExcel = mod.default || mod
    const file = new File([arrayBuffer], fileName)
    LuckyExcel.transformExcelToLucky(file, (exportJson) => {
      self.postMessage({ ok: true, exportJson })
    })
  } catch (err) {
    self.postMessage({ ok: false, error: String((err && err.stack) || err) })
  }
}
