# Office Viewer

로컬 Office 문서를 **한 창에서** 보는 통합 뷰어. 왼쪽 파일 탐색기 + 오른쪽 탭/분할 문서 영역.

여러 엑셀을 매번 새 창으로 여는 불편을 없애기 위해 만든 데스크톱 앱(Electron).

## 현재 상태 (MVP)

- ✅ 왼쪽 폴더 탐색기 (지연 로딩 트리, 오피스 문서만 표시)
- ✅ 오른쪽 탭 — 파일을 탭으로 누적해서 열기
- ✅ 분할 보기 — 패널을 좌우로 나눠 동시에 비교 (각 패널 독립)
- ✅ 엑셀(.xlsx/.xlsm) 렌더링 — Univer 엔진, **보기 + 편집**
  - 글꼴/색/병합/숫자서식/수식/다중 시트까지 보존 (luckyexcel 가져오기)
- ✅ 한국어 UI
- ⏳ 워드(.docx)·PPT(.pptx) — 현재는 "원래 프로그램으로 열기" 폴백
- ⏳ 편집 내용 .xlsx로 저장(내보내기)
- ⏳ 패널 경계 드래그 리사이즈

## 기술 구성

| 영역 | 사용 기술 |
|------|-----------|
| 셸 | Electron + electron-vite |
| 스프레드시트 | @univerjs/presets (Univer 0.25) — 다중 인스턴스 |
| xlsx 가져오기 | luckyexcel → Univer 변환기 (`src/renderer/src/viewers/luckyToUniver.js`) |

## 실행

```bash
npm install
npm run dev      # 개발 모드 (창 + DevTools)
npm run build    # 프로덕션 번들 빌드
```

### 스모크 테스트

```bash
OV_SMOKE=1 npm run dev   # 샘플을 자동으로 열고 sample-data/_smoke_*.png 캡처
```

## 폴더 구조

```
src/
  main/index.js        Electron 메인 — 창 생성, 파일 IPC(폴더/읽기/쓰기/열기)
  preload/index.js     contextBridge로 window.api 노출
  renderer/
    index.html
    src/
      main.js          진입점 — 탐색기 + 워크스페이스 연결
      explorer.js      왼쪽 파일 트리
      workspace.js     패널/탭/분할 로직
      viewers/
        excel.js          Univer 인스턴스 + luckyexcel 가져오기
        luckyToUniver.js  Luckysheet 포맷 → Univer IWorkbookData 변환
        fallback.js       워드/PPT 폴백
      style.css
sample-data/           테스트용 샘플 엑셀
```
