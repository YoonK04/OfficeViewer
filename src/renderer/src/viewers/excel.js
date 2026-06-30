// 엑셀 뷰어: Univer 인스턴스(컨테이너별 격리) + luckyexcel 가져오기
import { createUniver, defaultTheme, LocaleType, merge } from '@univerjs/presets'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import sheetsCoreKoKR from '@univerjs/preset-sheets-core/locales/ko-KR'
import '@univerjs/preset-sheets-core/lib/index.css'
import LuckyExcel from 'luckyexcel'
import { luckyToUniver } from './luckyToUniver.js'

function emptyWorkbook(name) {
  return {
    id: 'wb-empty',
    name,
    sheetOrder: ['s1'],
    styles: {},
    sheets: { s1: { id: 's1', name: 'Sheet1', rowCount: 100, columnCount: 26, cellData: {} } }
  }
}

// container: 마운트할 DOM 요소
// 반환: { dispose(), getUniverAPI() }
export function createExcelViewer(container, { arrayBuffer, fileName }) {
  const { univer, univerAPI } = createUniver({
    locale: LocaleType.KO_KR,
    theme: defaultTheme,
    locales: { [LocaleType.KO_KR]: merge({}, sheetsCoreKoKR) },
    presets: [UniverSheetsCorePreset({ container })]
  })

  const file = new File([arrayBuffer], fileName, {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  })

  try {
    LuckyExcel.transformExcelToLucky(file, (exportJson) => {
      try {
        if (!exportJson || !exportJson.sheets || exportJson.sheets.length === 0) {
          univerAPI.createWorkbook(emptyWorkbook(fileName))
          return
        }
        const wb = luckyToUniver(exportJson, fileName)
        univerAPI.createWorkbook(wb)
      } catch (err) {
        console.error('[excel] 변환 실패:', err)
        univerAPI.createWorkbook(emptyWorkbook(fileName))
      }
    })
  } catch (err) {
    console.error('[excel] luckyexcel 파싱 실패:', err)
    univerAPI.createWorkbook(emptyWorkbook(fileName))
  }

  return {
    dispose() {
      try {
        univer.dispose()
      } catch (e) {
        /* noop */
      }
    },
    getUniverAPI: () => univerAPI
  }
}
