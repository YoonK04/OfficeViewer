# Office Viewer

로컬 Office 문서를 **한 창에서** 보는 통합 뷰어. 왼쪽 파일 탐색기 + 오른쪽 탭/분할 문서 영역.

엑셀·워드·PPT를 매번 새 창으로 여는 불편을 없애기 위해 만든 데스크톱 앱입니다. 파일을 탭으로 쌓아 열고, 화면을 좌우로 나눠 동시에 비교할 수 있습니다.

> 모든 처리는 **로컬에서 오프라인**으로 이뤄집니다. 파일이 외부로 전송되지 않습니다.

## 주요 기능

| 기능 | 상태 | 설명 |
|------|:---:|------|
| 파일 탐색기 | ✅ | 폴더를 열어 트리로 탐색 (지연 로딩, 오피스 문서만 표시) |
| 탭 | ✅ | 파일을 탭으로 누적해서 열기 / 닫기 |
| 좌우 분할 | ✅ | 패널을 나눠 동시에 비교, 패널 경계 드래그로 너비 조절 |
| 엑셀 (.xlsx/.xlsm) | ✅ | **보기 + 편집** — 글꼴·색·병합·숫자서식·수식·다중 시트 보존 |
| 엑셀 저장 | ✅ | `Ctrl+S` 로 원본 .xlsx에 서식 보존 저장 (편집 시 탭에 ● 표시) |
| 워드 (.docx) | ✅ | 앱 내 미리보기 (제목·서식·표·목록 렌더링, 보기 전용) |
| PPT (.pptx) | ✅ | 앱 내 슬라이드 미리보기 (보기 전용, 오프라인 베스트에포트) |
| 구버전 (.doc/.ppt/.xls) | ↗ | "원래 프로그램으로 열기" 폴백 |
| 한국어 UI | ✅ | 엑셀 편집 리본 한국어 |

## 화면 구성

```
┌──────────┬─────────────────────────────────────┐
│ 탐색기    │  [탭1] [탭2] [탭3]            ⊞  ⊟   │
│  📁 폴더  ├──────────────────┬──────────────────┤
│  ├ 📗 a   │                  │                  │
│  ├ 📄 b   │   문서 보기/편집   │   분할된 두 번째   │
│  └ 📊 c   │     (패널 A)      │   패널 (패널 B)    │
│          │                  ║◄ 경계 드래그       │
└──────────┴──────────────────┴──────────────────┘
```

- 탭바 우측 **⊞** : 현재 패널을 오른쪽으로 분할
- 탭바 우측 **⊟** : 해당 패널 닫기 (패널이 2개 이상일 때)
- `Ctrl+S` : 활성 엑셀 탭 저장

## 기술 구성

| 영역 | 사용 기술 |
|------|-----------|
| 셸 | Electron + electron-vite + Vite |
| 스프레드시트 | [@univerjs/presets](https://univer.ai) (Univer 0.25) — 다중 인스턴스 |
| xlsx 가져오기 | luckyexcel → Univer 변환 (서식 보존, `viewers/luckyToUniver.js`) |
| xlsx 저장 | Univer 스냅샷 → [ExcelJS](https://github.com/exceljs/exceljs) (`viewers/univerToXlsx.js`) |
| 워드 미리보기 | [docx-preview](https://github.com/VolodymyrBaydalka/docxjs) |
| PPT 미리보기 | [pptx-preview](https://github.com/501351981/pptx-preview) |

## 실행

```bash
npm install
npm run dev        # 개발 모드 (창 + DevTools)
npm run build      # 렌더러/메인 번들 빌드
```

### 설치형 .exe 만들기 (Windows)

```bash
npm run dist       # release/ 에 NSIS 설치 파일 생성
npm run dist:dir   # 설치 없이 압축 해제형(폴더) 빌드만
```

### 스모크 테스트

샘플을 자동으로 열고 각 형식을 렌더링해 `sample-data/_smoke_*.png` 로 캡처합니다.

```bash
# 샘플 생성 (python + openpyxl/python-docx/python-pptx 필요)
python sample-data/make_samples.py
python sample-data/make_docs.py

OV_SMOKE=1 npm run dev
```

## 폴더 구조

```
src/
  main/index.js          Electron 메인 — 창 생성, 파일 IPC(폴더/읽기/쓰기/열기), 스모크
  preload/index.js       contextBridge 로 window.api 노출
  renderer/
    index.html
    src/
      main.js            진입점 — 탐색기 + 워크스페이스 연결, Ctrl+S
      explorer.js        왼쪽 파일 트리
      workspace.js       패널/탭/분할/리사이즈/저장 로직
      viewers/
        excel.js            Univer 인스턴스 + 가져오기 + dirty 감지 + save()
        luckyToUniver.js    Luckysheet 포맷 → Univer IWorkbookData
        univerToXlsx.js     Univer 스냅샷 → xlsx(ExcelJS, 서식 보존)
        word.js             docx-preview 워드 미리보기
        ppt.js              pptx-preview PPT 미리보기
        fallback.js         구버전 형식 폴백(원래 프로그램으로 열기)
      style.css
sample-data/             테스트용 샘플 + 생성 스크립트
```

## 알려진 한계

- 엑셀 저장은 앱이 추적하는 서식(값·수식·글꼴·색·정렬·병합·숫자서식·열너비·행높이)을 보존합니다. 차트·조건부서식·이미지 등 고급 요소는 보기에서는 표시되지 않을 수 있고 저장 시 유지되지 않습니다.
- 워드·PPT는 **보기 전용**입니다. 편집은 "원래 프로그램으로 열기"를 사용하세요.
- PPT 미리보기는 오프라인 렌더링 특성상 폰트·정밀 레이아웃에서 원본과 차이가 있을 수 있습니다.

## 라이선스

MIT (앱 코드). 의존 라이브러리는 각자의 라이선스를 따릅니다.
