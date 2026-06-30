import { app, BrowserWindow, ipcMain, dialog, shell } from 'electron'
import { join } from 'path'
import { promises as fs } from 'fs'
import { fileURLToPath } from 'url'
import { dirname } from 'path'

const __dirname = dirname(fileURLToPath(import.meta.url))

const SUPPORTED = ['.xlsx', '.xlsm', '.xls', '.csv', '.docx', '.doc', '.pptx', '.ppt']

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

// 스모크 테스트: 샘플 폴더를 열고 파일을 띄운 뒤 화면을 PNG로 캡처
async function runSmoke(win) {
  const sampleDir = join(__dirname, '../../sample-data')
  const file1 = { path: join(sampleDir, '매출요약.xlsx'), name: '매출요약.xlsx', ext: '.xlsx' }
  const file2 = { path: join(sampleDir, '재고현황.xlsx'), name: '재고현황.xlsx', ext: '.xlsx' }
  try {
    await new Promise((r) => setTimeout(r, 1500))
    await win.webContents.executeJavaScript(
      `window.__ov.setRoot(${JSON.stringify(sampleDir)})`
    )
    await win.webContents.executeJavaScript(
      `window.__ov.openFile(${JSON.stringify(file1)})`
    )
    await new Promise((r) => setTimeout(r, 3500))
    const img1 = await win.webContents.capturePage()
    await fs.writeFile(join(sampleDir, '_smoke_1.png'), img1.toPNG())
    console.log('SMOKE_CAPTURED 1')

    // 분할 + 두 번째 파일
    await win.webContents.executeJavaScript(`window.__ov.splitActivePane()`)
    await win.webContents.executeJavaScript(
      `window.__ov.openFile(${JSON.stringify(file2)})`
    )
    await new Promise((r) => setTimeout(r, 3000))
    const img2 = await win.webContents.capturePage()
    await fs.writeFile(join(sampleDir, '_smoke_2.png'), img2.toPNG())
    console.log('SMOKE_CAPTURED 2')
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
