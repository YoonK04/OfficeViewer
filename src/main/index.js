import { app, BrowserWindow, ipcMain, dialog, shell } from 'electron'
import { join } from 'path'
import { promises as fs } from 'fs'
import { fileURLToPath } from 'url'
import { dirname } from 'path'

const __dirname = dirname(fileURLToPath(import.meta.url))

const SUPPORTED = ['.xlsx', '.xlsm', '.xls', '.csv', '.docx', '.doc', '.pptx', '.ppt', '.hwp', '.hwpx', '.pdf']

function createWindow() {
  const win = new BrowserWindow({
    width: 1440,
    height: 900,
    title: 'Office Viewer',
    backgroundColor: '#1e1e1e',
    webPreferences: {
      preload: join(__dirname, '../preload/index.js'),
      sandbox: false,
      contextIsolation: true,
      nodeIntegration: false
    }
  })

  // 렌더러 콘솔 → 메인 stdout (디버깅용)
  win.webContents.on('console-message', (_e, level, message) => {
    console.log(`[renderer:${level}] ${message}`)
  })

  // 파일을 창에 드롭했을 때 그 파일로 페이지가 이동하는 기본 동작 차단
  win.webContents.on('will-navigate', (e) => e.preventDefault())
  win.webContents.setWindowOpenHandler(() => ({ action: 'deny' }))

  if (process.env['ELECTRON_RENDERER_URL']) {
    win.loadURL(process.env['ELECTRON_RENDERER_URL'])
    if (!process.env['OV_SMOKE']) win.webContents.openDevTools()
  } else {
    win.loadFile(join(__dirname, '../renderer/index.html'))
  }

  if (process.env['OV_SMOKE']) {
    win.webContents.on('did-finish-load', () => runSmoke(win))
  }
}

// 스모크 테스트: 샘플 폴더를 열고 각 형식을 띄운 뒤 화면을 PNG로 캡처
async function runSmoke(win) {
  const sampleDir = join(__dirname, '../../sample-data')
  const wait = (ms) => new Promise((r) => setTimeout(r, ms))
  const run = (js) => win.webContents.executeJavaScript(js)
  const cap = async (name) => {
    const img = await win.webContents.capturePage()
    await fs.writeFile(join(sampleDir, `_smoke_${name}.png`), img.toPNG())
    console.log('SMOKE_CAPTURED ' + name)
  }
  const F = (n, ext) => ({ path: join(sampleDir, n), name: n, ext })
  try {
    await wait(1500)
    await run(`window.__ov.setRoot(${JSON.stringify(sampleDir)})`)
    await win.webContents.capturePage() // 워밍업(첫 캡처 빈이미지 방지)

    // 0-e) #2: 분할 패널의 탭을 모두 닫으면 패널 자동 제거
    await run(`window.__ov.splitActivePane()`)
    await run(`window.__ov.openFile(${JSON.stringify(F('감사테스트.xlsx', '.xlsx'))})`)
    await wait(2500)
    await run(`(() => {
      const before = document.querySelectorAll('.pane').length;
      const active = document.querySelector('.pane.active');
      active.querySelector('.tab.active .tab-close').click();
      return before + '->' + document.querySelectorAll('.pane').length;
    })()`).then((n) => console.log('SMOKE_EMPTYPANE ' + n))
    await wait(500)

    // 0-perf) 대용량 xlsm 로드 시간 측정
    await run(`window.__ov.openFile(${JSON.stringify(F('대용량.xlsm', '.xlsm'))})`)
    await wait(8000)
    await cap('heavy')

    // 0-pdf) PDF 뷰어(pdf.js 캔버스)
    await run(`window.__ov.openFile(${JSON.stringify(F('샘플문서.pdf', '.pdf'))})`)
    await wait(3000)
    await run(`(() => {
      const canvases = document.querySelectorAll('.pdf-page canvas');
      return JSON.stringify({ pages: document.querySelectorAll('.pdf-page').length, rendered: canvases.length });
    })()`).then((r) => console.log('SMOKE_PDF_STATE ' + r))
    await cap('pdf')

    // 0) 공정도(테두리/병합/열너비 충실도 확인)
    await run(`window.__ov.openFile(${JSON.stringify(F('공정도.xlsx', '.xlsx'))})`)
    await wait(3500)
    await cap('flowchart')

    // 0-b) 감사테스트(숨김 행·열 / 회전 / 격자선 끄기)
    await run(`window.__ov.openFile(${JSON.stringify(F('감사테스트.xlsx', '.xlsx'))})`)
    await wait(2000)
    await run(`(() => {
      try {
        const api = window.__ov.getActiveViewer().getUniverAPI();
        const ws = api.getActiveWorkbook().getActiveSheet();
        ws.getRange('A1').activate();
        if (ws.scrollToCell) ws.scrollToCell(0, 0);
      } catch(e) { return 'scroll-err:'+e; }
    })()`)
    await wait(1500)
    await cap('audit')

    // 0-c) 드래그&드롭 오버레이(합성 dragenter)
    await run(`(() => {
      const dt = new DataTransfer();
      dt.items.add(new File(['x'], 'test.xlsx'));
      window.dispatchEvent(new DragEvent('dragenter', { bubbles: true, dataTransfer: dt }));
      return document.getElementById('drop-overlay').classList.contains('show');
    })()`).then((v) => console.log('SMOKE_DROP_OVERLAY ' + v))
    await wait(700)
    await cap('dnd-overlay')
    await run(`window.dispatchEvent(new DragEvent('dragleave', { bubbles: true }))`)

    // 0-d) 드롭 파이프라인: 실제 경로로 openDroppedFiles → 우측 좌표에 열기
    await run(`window.__ov.splitActivePane()`)
    await run(`(() => {
      const pane = document.querySelectorAll('.pane')[1];
      const r = pane.getBoundingClientRect();
      return window.__ov.openDroppedFiles(
        [${JSON.stringify(F('감사테스트.xlsx', '.xlsx'))}],
        r.left + r.width/2, r.top + r.height/2
      );
    })()`)
    await wait(3000)
    await cap('dnd-open')

    // 1) 엑셀
    await run(`window.__ov.openFile(${JSON.stringify(F('매출요약.xlsx', '.xlsx'))})`)
    await wait(3500)
    await cap('excel')

    // 2) 저장 라운드트립: 셀 편집 → 저장 → 재오픈
    try {
      const diag = await run(`(() => {
        const v = window.__ov.getActiveViewer();
        return { type: typeof v, keys: v ? Object.keys(v) : null };
      })()`)
      console.log('SMOKE_VIEWER ' + JSON.stringify(diag))
      await run(`(() => {
        const v = window.__ov.getActiveViewer();
        const api = v.getUniverAPI();
        api.getActiveWorkbook().getActiveSheet().getRange('A10').setValue('저장 라운드트립 OK');
      })()`)
      await wait(800)
      await run(`window.__ov.saveActiveTab()`)
      await wait(2500)
      console.log('SMOKE_SAVED')
    } catch (e) {
      console.log('SMOKE_SAVE_ERROR ' + (e && e.message ? e.message : e))
    }

    // 3) 워드
    try {
      await run(`window.__ov.openFile(${JSON.stringify(F('보고서.docx', '.docx'))})`)
      await wait(2500)
      await cap('word')
    } catch (e) {
      console.log('SMOKE_WORD_ERROR ' + e)
    }

    // 4) PPT
    try {
      await run(`window.__ov.openFile(${JSON.stringify(F('발표자료.pptx', '.pptx'))})`)
      await wait(3000)
      await cap('ppt')
    } catch (e) {
      console.log('SMOKE_PPT_ERROR ' + e)
    }

    // 5) 분할 + 저장된 엑셀 재오픈(라운드트립 결과 확인)
    try {
      await run(`window.__ov.splitActivePane()`)
      await run(`window.__ov.openFile(${JSON.stringify(F('매출요약.xlsx', '.xlsx'))})`)
      await wait(3500)
      await cap('split')
    } catch (e) {
      console.log('SMOKE_SPLIT_ERROR ' + e)
    }

    // 6) 탭 드래그 → 우측 분할 (다중 탭이 있는 첫 패널의 탭을 빼내 분할)
    try {
      await run(`(() => {
        const pane = document.querySelectorAll('.pane')[0];
        pane.dispatchEvent(new MouseEvent('mousedown',{bubbles:true}));
        const tab = pane.querySelector('.tab.active') || pane.querySelector('.tab');
        const tr = tab.getBoundingClientRect(), pr = pane.getBoundingClientRect();
        const sx = tr.left + tr.width/2, sy = tr.top + tr.height/2;
        const tx = pr.right - 20, ty = pr.top + pr.height/2;
        const fire = (t,x,y,el) => (el||window).dispatchEvent(new MouseEvent(t,{bubbles:true,clientX:x,clientY:y,button:0}));
        fire('mousedown',sx,sy,tab); fire('mousemove',sx+10,sy);
        fire('mousemove',(sx+tx)/2,(sy+ty)/2); fire('mousemove',tx,ty); fire('mouseup',tx,ty);
        return [...document.querySelectorAll('.pane')].map(p => p.querySelectorAll('.tab').length).join(',');
      })()`).then((n) => console.log('SMOKE_PANES_AFTER_SPLIT tabs=[' + n + ']'))
      await wait(2500)
      await cap('drag-split')
    } catch (e) {
      console.log('SMOKE_DRAGSPLIT_ERROR ' + e)
    }

    // 7) 탭 드래그 → 첫 패널 탭바에 합치기(merge)
    try {
      await run(`(() => {
        const src = document.querySelector('.pane.active') || document.querySelector('.pane');
        const tab = src.querySelector('.tab.active') || src.querySelector('.tab');
        const first = document.querySelectorAll('.pane')[0];
        const tb = first.querySelector('.tabbar').getBoundingClientRect();
        const tr = tab.getBoundingClientRect();
        const sx = tr.left + tr.width/2, sy = tr.top + tr.height/2;
        const tx = tb.left + tb.width/2, ty = tb.top + tb.height/2;
        const fire = (t,x,y,el) => (el||window).dispatchEvent(new MouseEvent(t,{bubbles:true,clientX:x,clientY:y,button:0}));
        fire('mousedown',sx,sy,tab); fire('mousemove',sx+10,sy);
        fire('mousemove',(sx+tx)/2,(sy+ty)/2); fire('mousemove',tx,ty); fire('mouseup',tx,ty);
        return document.querySelectorAll('.pane').length;
      })()`).then((n) => console.log('SMOKE_PANES_AFTER_MERGE ' + n))
      await wait(2500)
      await cap('drag-merge')
    } catch (e) {
      console.log('SMOKE_DRAGMERGE_ERROR ' + e)
    }

    console.log('SMOKE_DONE')
  } catch (err) {
    console.log('SMOKE_ERROR ' + (err && err.stack ? err.stack : err))
  }
}

// 폴더 선택 대화상자
ipcMain.handle('dialog:openFolder', async () => {
  const res = await dialog.showOpenDialog({ properties: ['openDirectory'] })
  if (res.canceled || res.filePaths.length === 0) return null
  return res.filePaths[0]
})

// 디렉터리 한 단계 읽기 (지연 로딩용)
ipcMain.handle('fs:readDir', async (_evt, dirPath) => {
  const entries = await fs.readdir(dirPath, { withFileTypes: true })
  const result = []
  for (const e of entries) {
    if (e.name.startsWith('~$') || e.name.startsWith('.')) continue // Office 임시파일/숨김 제외
    const full = join(dirPath, e.name)
    if (e.isDirectory()) {
      result.push({ name: e.name, path: full, type: 'dir' })
    } else {
      const ext = e.name.slice(e.name.lastIndexOf('.')).toLowerCase()
      if (SUPPORTED.includes(ext)) {
        result.push({ name: e.name, path: full, type: 'file', ext })
      }
    }
  }
  // 폴더 먼저, 그다음 파일 (이름순)
  result.sort((a, b) => {
    if (a.type !== b.type) return a.type === 'dir' ? -1 : 1
    return a.name.localeCompare(b.name, 'ko')
  })
  return result
})

// 파일 바이너리 읽기 (렌더러에서 파싱)
ipcMain.handle('fs:readFile', async (_evt, filePath) => {
  const buf = await fs.readFile(filePath)
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength)
})

// 파일 저장
ipcMain.handle('fs:writeFile', async (_evt, filePath, arrayBuffer) => {
  await fs.writeFile(filePath, Buffer.from(arrayBuffer))
  return true
})

// 원래 연결 프로그램으로 열기
ipcMain.handle('shell:openPath', async (_evt, filePath) => {
  return shell.openPath(filePath)
})

app.whenReady().then(() => {
  createWindow()
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow()
  })
})

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit()
})
