// 엑셀 뷰어: Univer 인스턴스(컨테이너별 격리) + luckyexcel 가져오기
import { createUniver, defaultTheme, LocaleType, merge } from '@univerjs/presets'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import sheetsCoreKoKR from '@univerjs/preset-sheets-core/locales/ko-KR'
import '@univerjs/preset-sheets-core/lib/index.css'
import LuckyExcel from 'luckyexcel'
import { luckyToUniver } from './luckyToUniver.js'
import { univerSnapshotToXlsx } from './univerToXlsx.js'

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
// opts.onDirtyChange: (isDirty) => void  편집 상태 변화 알림
// 반환: { dispose(), getUniverAPI(), save(), isDirty() }
export function createExcelViewer(container, { arrayBuffer, fileName, onDirtyChange }) {
  const { univer, univerAPI } = createUniver({
    locale: LocaleType.KO_KR,
    theme: defaultTheme,
    locales: { [LocaleType.KO_KR]: merge({}, sheetsCoreKoKR) },
    presets: [UniverSheetsCorePreset({ container })]
  })

  let dirty = false
  let ready = false
  const setDirty = (v) => {
    if (dirty === v) return
    dirty = v
    if (onDirtyChange) onDirtyChange(v)
  }

  // 사용자 편집 커맨드만 감지 (로드 시 자동 재계산 mutation은 제외)
  const EDIT_CMD = /^sheet\.command\.(set-range-values|set-style|set-bold|set-italic|set-underline|set-overline|set-strike|set-font|set-background|set-text-color|set-text-wrap|set-text-rotation|set-border|set-horizontal|set-vertical|clear-selection|insert-row|insert-col|insert-multi|remove-row|remove-col|delete-range|move-range|auto-fill|auto-clear|copy-down|copy-right|paste|delta-(column-width|row-height)|set-(col|row)-data|set-row-height|set-worksheet-col-width|(add|remove)-worksheet-merge|set-tab-color|reset-(background-color|text-color))/
  try {
    univerAPI.onCommandExecuted((command) => {
      if (!ready) return
      const id = command && command.id ? String(command.id) : ''
      if (EDIT_CMD.test(id)) setDirty(true)
    })
  } catch (e) {
    /* 이벤트 미지원 시 무시 */
  }

  const file = new File([arrayBuffer], fileName, {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  })

  try {
    LuckyExcel.transformExcelToLucky(file, (exportJson) => {
      try {
        if (!exportJson || !exportJson.sheets || exportJson.sheets.length === 0) {
          univerAPI.createWorkbook(emptyWorkbook(fileName))
        } else {
          const wb = luckyToUniver(exportJson, fileName)
          univerAPI.createWorkbook(wb)
        }
      } catch (err) {
        console.error('[excel] 변환 실패:', err)
        univerAPI.createWorkbook(emptyWorkbook(fileName))
      }
      // 초기 로드 mutation이 dirty로 잡히지 않도록 다음 틱부터 감지
      setTimeout(() => {
        ready = true
      }, 300)
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
    getUniverAPI: () => univerAPI,
    isDirty: () => dirty,
    // 현재 워크북을 xlsx ArrayBuffer로 직렬화
    async save() {
      const wb = univerAPI.getActiveWorkbook()
      const snapshot = wb.save()
      const buf = await univerSnapshotToXlsx(snapshot)
      setDirty(false)
      return buf
    }
  }
}
