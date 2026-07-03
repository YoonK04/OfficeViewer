# Office Viewer

로컬 Office 문서를 **한 창에서** 보는 통합 뷰어. 왼쪽 파일 탐색기 + 오른쪽 탭/분할 문서 영역.

엑셀·워드·PPT를 매번 새 창으로 여는 불편을 없애기 위해 만든 데스크톱 앱입니다. 파일을 탭으로 쌓아 열고, 화면을 좌우로 나눠 동시에 비교할 수 있습니다.

> 모든 처리는 **로컬에서 오프라인**으로 이뤄집니다. 파일이 외부로 전송되지 않습니다.

📖 **[사용 방법 가이드 (이미지 포함)](docs/사용방법.md)** · 오프라인용 단일 파일: [`docs/사용방법.html`](docs/사용방법.html)

## 주요 기능

| 기능 | 상태 | 설명 |
|------|:---:|------|
| 파일 탐색기 | ✅ | 폴더를 열어 트리로 탐색 (지연 로딩, 오피스 문서만 표시) |
| **드래그&드롭** | ✅ | Windows 탐색기에서 파일을 **끌어다 놓으면** 놓은 패널에서 바로 열림 |
| 탭 | ✅ | 파일을 탭으로 누적해서 열기 / 닫기 |
| **탭 드래그** | ✅ | 탭을 **우측 가장자리에 끌어다 놓으면 분할**, **다른 패널 탭바에 놓으면 합치기** |
| 좌우 분할 | ✅ | 패널을 나눠 동시에 비교, 패널 경계 드래그로 너비 조절 |
| 엑셀 (.xlsx/.xlsm) | ✅ | **보기 + 편집** — 글꼴·색·**테두리**·병합·숫자서식·수식·열너비·다중 시트 보존 |
| 엑셀 저장 | ✅ | `Ctrl+S` 로 원본 .xlsx에 서식 보존 저장 (편집 시 탭에 ● 표시) |
| 워드 (.docx) | ✅ | 앱 내 미리보기 (제목·서식·표·목록 렌더링, 보기 전용) |
| PPT (.pptx) | ✅ | 앱 내 슬라이드 미리보기 (보기 전용, 오프라인 베스트에포트) |
| PDF (.pdf) | ✅ | 앱 내 미리보기 (pdf.js, 페이지 스크롤, 원본 그대로) |
| 한글 (.hwp) | 🧪 | 실험적 미리보기(hwp.js, 단순 문서 위주) — 실패 시 "원래 프로그램으로 열기" 자동 폴백 |
| 구버전·기타 (.doc/.ppt/.xls/.hwpx) | ↗ | "원래 프로그램으로 열기" 폴백 |
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
npm run dist       # release/ 에 NSIS 설치 파일(.exe) 생성
npm run dist:dir   # 설치 없이 압축 해제형(폴더) 빌드만 → release/win-unpacked/
```

`npm run dist` 를 돌리면 `release/Office Viewer Setup <버전>.exe` 설치 파일이 만들어집니다.
이 파일 하나만 전달하면 받는 사람이 더블클릭으로 설치해 쓸 수 있습니다.

> **코드 서명은 비활성화**되어 있습니다(`win.signAndEditExecutable: false`).
> 이렇게 한 이유: electron-builder가 서명 도구(`winCodeSign`)를 풀 때 macOS용 심볼릭 링크가 포함돼,
> **개발자 모드/관리자 권한이 없는 Windows에서는** `Cannot create symbolic link` 오류로 빌드가 실패합니다.
> 서명을 끄면 그 단계를 건너뛰어 일반 사용자 권한으로도 설치 파일이 만들어집니다.
> (서명되지 않은 앱이므로 최초 실행 시 SmartScreen "추가 정보 → 실행" 안내가 나올 수 있습니다.)
>
> 미리 빌드된 설치 파일/포터블은 [Releases](https://github.com/YoonK04/OfficeViewer/releases)에서 받을 수 있습니다.

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

## 엑셀 충실도 (원본 대비)

가져오기/저장에서 보존하는 요소:

- 값·수식, 다중 시트
- 글꼴(이름·크기·굵게·기울임·밑줄·취소선)·글자색·배경색
- 가로/세로 정렬·자동 줄바꿈·**텍스트 회전**
- **셀 테두리**(스타일·색)
- 병합 셀, 숫자/날짜 서식
- 열 너비·행 높이·**숨김 행/열**·**격자선 표시 여부**

보존되지 **않는** 요소(가져오기 라이브러리 한계 — 데이터 자체가 추출되지 않음):

- **도형·화살표·텍스트박스 등 드로잉 객체**, 이미지, 차트
- 틀 고정(freeze panes), 조건부 서식, 데이터 유효성(드롭다운), 시트 보호

> 셀 테두리로 그린 표/박스는 정상 복원되지만, **도형(autoshape)·선·화살표로 그린 도식**은 표시되지 않습니다.
> 이런 객체가 많은 파일은 "원래 프로그램으로 열기"로 보는 것을 권장합니다.

## 그 밖의 한계

- 워드·PPT는 **보기 전용**입니다. 편집은 "원래 프로그램으로 열기"를 사용하세요.
- PPT 미리보기는 오프라인 렌더링 특성상 폰트·정밀 레이아웃에서 원본과 차이가 있을 수 있습니다.
- **HWP는 실험적**입니다. `.hwp`(5.0)만 대상이며 표·이미지가 많은 복잡한 문서는 렌더가 실패할 수 있고, 그럴 때는 자동으로 "원래 프로그램으로 열기"(한/글) 버튼으로 넘어갑니다. `.hwpx`는 앱 내 미리보기를 지원하지 않습니다.

## 라이선스

MIT (앱 코드). 의존 라이브러리는 각자의 라이선스를 따릅니다.
