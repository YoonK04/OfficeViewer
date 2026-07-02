# 사용방법 가이드를 이미지 내장 단일 HTML 파일로 생성 (오프라인 전달용)
import os, base64

base = os.path.dirname(os.path.abspath(__file__))
img_dir = os.path.join(base, "images")

def img(name):
    with open(os.path.join(img_dir, name), "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
    return f'<img src="data:image/png;base64,{b64}" alt="{name}">'

html = f"""<!doctype html>
<html lang="ko"><head><meta charset="utf-8">
<title>Office Viewer 사용 방법</title>
<style>
  body {{ font-family:"Malgun Gothic","Segoe UI",sans-serif; max-width:860px; margin:0 auto;
         padding:40px 24px; color:#222; line-height:1.7; }}
  h1 {{ border-bottom:3px solid #0e639c; padding-bottom:10px; }}
  h2 {{ margin-top:38px; color:#0e639c; }}
  img {{ width:100%; border:1px solid #ddd; border-radius:6px; margin:10px 0;
        box-shadow:0 2px 10px #0002; }}
  table {{ border-collapse:collapse; width:100%; margin:12px 0; }}
  th,td {{ border:1px solid #ccc; padding:8px 12px; text-align:left; }}
  th {{ background:#f0f4f8; }}
  blockquote {{ background:#fff8e1; border-left:4px solid #f0ad4e; margin:14px 0;
               padding:10px 16px; color:#555; }}
  code {{ background:#eef1f4; padding:2px 6px; border-radius:4px; }}
  kbd {{ background:#333; color:#fff; padding:2px 7px; border-radius:4px; font-size:.9em; }}
</style></head><body>

<h1>Office Viewer 사용 방법</h1>
<p>엑셀·워드·PPT를 <b>한 창에서</b> 탭으로 열고, 화면을 나눠 비교하는 통합 뷰어입니다.
모든 작업은 <b>내 컴퓨터에서 오프라인</b>으로 이뤄지며 파일이 외부로 전송되지 않습니다.</p>

<h2>1. 설치 / 실행</h2>
<ul>
<li><b>설치형:</b> <code>Office Viewer Setup 0.1.0.exe</code> 더블클릭 → 설치 → 바탕화면 아이콘 실행</li>
<li><b>포터블:</b> zip을 풀고 <code>Office Viewer.exe</code> 더블클릭(설치 불필요)</li>
</ul>
<blockquote>서명되지 않은 앱이라 처음 실행 시 파란 SmartScreen 창이 뜰 수 있습니다.
<b>추가 정보 → 실행</b>을 누르면 됩니다. (안전한 로컬 앱입니다.)</blockquote>

<h2>2. 파일 여는 방법 (두 가지)</h2>
<p><b>방법 A — 폴더 열기:</b> 왼쪽 위 <b>📁 폴더 열기</b>로 폴더를 선택하면 그 안의 오피스 문서만
왼쪽 트리에 표시됩니다. 파일을 클릭하면 열립니다. 폴더(▶)를 눌러 하위 폴더도 펼칠 수 있습니다.</p>
<p><b>방법 B — 드래그&드롭:</b> Windows 파일 탐색기에서 파일을 <b>끌어다가 뷰어 위에 놓기만</b> 하면
바로 열립니다. 분할해 둔 상태면 <b>놓은 쪽 패널</b>에서 열리고, 여러 개를 한 번에 놓아도 됩니다.</p>

<h2>3. 엑셀 열기 · 편집 · 저장</h2>
<p>엑셀 파일을 클릭하면 오른쪽에 <b>탭으로</b> 열립니다. 셀 색·병합·천단위·수식·여러 시트가
원본 그대로 표시되고, <b>셀을 클릭해 바로 편집</b>할 수 있습니다.</p>
{img("01-excel.png")}
<ul>
<li>편집하면 탭 이름 앞에 <b>●</b> 표시(저장 안 됨)가 생깁니다.</li>
<li><kbd>Ctrl</kbd>+<kbd>S</kbd> 로 원본 .xlsx에 <b>서식을 유지한 채 저장</b>됩니다.</li>
<li>하단 시트 탭으로 시트를 전환합니다.</li>
</ul>

<h2>4. 워드(.docx) 미리보기</h2>
<p>워드 문서를 클릭하면 페이지 형태 미리보기가 열립니다(제목·서식·표·목록 표시, <b>보기 전용</b>).</p>
{img("02-word.png")}

<h2>5. PPT(.pptx) 미리보기</h2>
<p>발표자료를 클릭하면 슬라이드가 차례로 표시됩니다(<b>보기 전용</b>).</p>
{img("03-ppt.png")}

<h2>6. 화면 나눠 보기 (분할)</h2>
<p><b>가장 쉬운 방법 — 탭을 끌어다 놓기:</b></p>
<ul>
<li>탭을 <b>화면 오른쪽 가장자리로 드래그</b>하면 그 자리에서 <b>좌우로 분할</b>됩니다.</li>
<li>분할된 탭을 <b>다른 패널의 탭 줄로 드래그</b>하면 다시 <b>한 화면으로 합쳐집니다</b>.</li>
</ul>
<p>버튼으로도 가능합니다:</p>
<ol>
<li>탭바 <b>오른쪽 끝 ⊞ 버튼</b>을 누르면 화면이 좌우로 나뉩니다.</li>
<li>새 패널이 활성화된 상태에서 다른 파일을 클릭하면 그쪽에 열립니다.</li>
<li>두 패널 <b>사이 경계선을 드래그</b>해 너비를 조절합니다.</li>
<li>패널을 닫으려면 그 패널의 <b>⊟ 버튼</b>을 누릅니다.</li>
</ol>
{img("04-split.png")}

<h2>7. 단축키 / 버튼 요약</h2>
<table>
<tr><th>동작</th><th>방법</th></tr>
<tr><td>폴더 열기</td><td>왼쪽 위 📁 폴더 열기</td></tr>
<tr><td>파일 열기</td><td>탐색기에서 파일 클릭(탭으로 열림)</td></tr>
<tr><td>탭 닫기</td><td>탭의 ×</td></tr>
<tr><td>엑셀 저장</td><td>Ctrl + S</td></tr>
<tr><td>화면 분할</td><td>탭바 우측 ⊞</td></tr>
<tr><td>패널 닫기</td><td>탭바 우측 ⊟</td></tr>
<tr><td>패널 너비 조절</td><td>패널 사이 경계선 드래그</td></tr>
</table>

<h2>지원 형식</h2>
<table>
<tr><th>형식</th><th>동작</th></tr>
<tr><td>.xlsx, .xlsm</td><td>보기 + 편집 + 저장</td></tr>
<tr><td>.docx</td><td>미리보기(보기 전용)</td></tr>
<tr><td>.pptx</td><td>미리보기(보기 전용)</td></tr>
<tr><td>.hwp</td><td>실험적 미리보기(단순 문서), 안 되면 원래 프로그램으로 열기</td></tr>
<tr><td>.doc, .ppt, .xls, .csv, .hwpx</td><td>"원래 프로그램으로 열기"</td></tr>
</table>
<blockquote>차트·조건부서식·이미지 등 일부 고급 요소는 표시되지 않을 수 있습니다.</blockquote>

</body></html>"""

out = os.path.join(base, "사용방법.html")
with open(out, "w", encoding="utf-8") as f:
    f.write(html)
print("written:", out, f"({len(html)//1024} KB)")
