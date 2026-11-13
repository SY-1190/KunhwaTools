# KunhwaTools

브라우저 문서·이미지 도구와 로컬 PDF 편집·변환 작업 공간입니다.

## 포함 기능 (업데이트)

1. PDF -> 이미지 변환기
2. 이미지 -> PDF 변환기
3. PDF 순서 변경(썸네일 드래그) 및 페이지 나누기
4. PDF 합치기
5. 이미지 크기 조정기
6. 이미지 확장자 포맷 변환기
7. 공정 작업시간 측정기 (공정명/고객사/메모/사진·동영상 첨부 및 촬영 + JSON/CSV 내보내기)
8. QR코드 생성기
9. PDF 글자·이미지 편집 및 새 PDF 다운로드 (브라우저 엔진)
10. PDF → Word(.docx) (로컬 엔진)
11. HWP/HWPX 포맷 변환 및 별도 웹 에디터
12. PDF → Excel, PDF 압축, 파일명 일괄 변경

## 실행 방법 (로컬)

PDF 편집은 GitHub Pages와 로컬 웹 모두에서 브라우저로 처리합니다. Word 변환은 아래 로컬 서버가 필요합니다.

PDF → Excel은 정적 웹에서도 브라우저 자동 분석으로 작동합니다. Python 기반 `pdfplumber` 정밀 엔진을 사용하려면 로컬 실행기를 시작합니다.

```powershell
.\start-kunhwa-tools.ps1
```

최초 실행 또는 PDF 엔진 업데이트 시 `.venv-pdf-excel`에 필요한 패키지를 설치하고 `http://127.0.0.1:8878/`에서 앱을 시작합니다. 로컬 변환 파일은 서버 디스크에 저장하지 않고 메모리에서 처리합니다. PDF → Excel은 정밀 엔진에 연결할 수 없으면 기존 브라우저 변환으로 전환됩니다. Word 변환에는 브라우저 대체 엔진이 없습니다.

## PDF 편집과 Word 변환

- PDF 편집에서 기존 글자를 두 번 클릭하거나, 선택 후 연필 버튼을 누르면 페이지 위에서 바로 수정할 수 있습니다. 적용한 뒤 새 PDF로 저장합니다.
- PDF 순서 변경에서는 PDF를 한 번에 여러 개 선택하거나 추가로 끌어다 놓을 수 있습니다. 파일이 달라도 페이지를 섞어 배치하고, 삭제·복원·회전한 뒤 하나의 PDF로 저장하거나 나눌 수 있습니다. 텍스트 분할 입력은 썸네일의 전체 페이지 번호(`p.1`, `p.2` 등)를 사용합니다.

- 글자는 원본 줄 단위로 선택합니다. 새 글자는 한글을 지원하는 기본 글꼴을 PDF에 넣어 검색·복사할 수 있습니다. 복잡한 글꼴·배치의 수정 결과는 원본과 다를 수 있습니다.
- 기존 글자·이미지는 PDF redaction으로 제거/교체합니다. 배경색이 있는 문서는 글자를 지울 영역의 배경색을 지정할 수 있습니다. 저장은 새 다운로드이며 원본 파일을 덮어쓰지 않습니다.
- 암호/편집 제한/서명 필드가 있는 PDF는 브라우저 편집에서 지원하지 않습니다. PDF는 150페이지 이하, 55MB 이하를 지원합니다.
- 스캔 문서의 글자는 직접 편집 가능한 텍스트가 아닙니다. Word 편집 모드는 PDF 본문·선 기반 표·이미지를 재구성하며 스캔 페이지는 이미지로 유지합니다. OCR은 적용되지 않으며 다단·복잡한 배치의 완전한 재현을 보장하지 않습니다.
- 원본 모양 유지 모드는 페이지를 이미지로 Word에 넣으므로 글자 편집은 불가능합니다.

## 숨김 게임

브랜드 로고를 두 번 클릭하거나 로고에 포커스 후 Shift+Enter로 Orbit Courier에 진입합니다. 방향키/WASD 또는 마우스·터치로 이동하고 Space 또는 하단 버튼으로 일시 정지합니다. 결정 수집, 운석 회피, 에너지 3회와 최고 점수 저장을 지원합니다. 기존 게임 코드는 제거했습니다.

## GitHub Pages 배포

1. 이 폴더를 GitHub 저장소에 push
2. GitHub 저장소 설정(Settings) -> Pages
3. `Deploy from a branch` 선택
4. Branch: `main` / Folder: `/ (root)` 선택 후 저장
5. 배포 URL 접속

## 기술 구성

- HTML/CSS/Vanilla JS
- pdf-lib (PDF 생성/편집)
- pdf.js (PDF 렌더링)
- JSZip (ZIP 다운로드)
- qrcodejs (QR 생성)
- pdfplumber 0.11.9 (PDF 문자·선·표 구조 분석, MIT License)
- openpyxl 3.1.5 (Excel 문서 생성)
- PyMuPDF 1.26.7 (PDF 렌더링·redaction·편집, AGPL 또는 상용 라이선스)
- MuPDF.js 1.28.6 (브라우저 PDF 편집, AGPL-3.0-or-later, 최초 파일 열기 시 로드, 약 10MB WASM)
- python-docx 1.2.0 (Word 문서 생성, MIT)
- Phaser 3.90.0 (게임 물리·렌더링, MIT; 로컬 vendor)
- Bootstrap Icons 1.13.1 (MIT; 로컬 SVG sprite)

PyMuPDF를 포함한 배포/공개 서비스는 AGPL 조건 또는 상용 라이선스를 검토해야 합니다. Phaser와 Bootstrap Icons 라이선스는 `assets`에 보존합니다. 새 UI는 공식 shadcn Sidebar의 분류·헤더 패턴을 디자인 참고로만 사용하며 React/shadcn 패키지는 도입하지 않았습니다.

브라우저 PDF 편집의 [라이선스와 대응 소스](./assets/vendor/mupdf-1.28.6/SOURCE.html)는 편집 화면에서도 확인할 수 있습니다. 이 PDF 통합 코드(`pdf-browser-engine.mjs`, `pdf-browser-worker.mjs`, `pdf-studio.js`)는 AGPL-3.0-or-later로 제공하며, 원본 MuPDF.js 배포 파일을 수정하지 않았습니다.

## PDF 엔진 검증

`node --test tests/test-pdf-browser.mjs`로 브라우저 엔진의 한글 교체, 이미지 교체·추가, 회전과 오류 입력을 검증합니다.

```powershell
.\.venv-pdf-excel\Scripts\python.exe -m unittest discover -s tests -v
```

글자 실제 삭제·한글 교체·이미지 교체, 편집 가능한 Word 본문·표·이미지, 병합 셀, 스캔/모양 유지 모드, 회전과 잘못된 입력을 검사합니다.

## UX 강화 사항

- 대용량 처리 진행률 바 표시
- 대용량 처리 취소 버튼 제공
- 기능별 에러/가이드 안내
- PDF 페이지 썸네일 Drag & Drop 정렬 UI
- PDF 분할 블록(구간 선택) UI
- 진행률 기반 ETA(남은 예상시간) 표시

## 참고

- `tools.mytory.net` 참고 방향으로 상단 안내 + 도구 목록 + 도구 실행 섹션 흐름을 적용했습니다.
- 상단 Kunhwa 브랜드와 분류형 도구 메뉴를 사용합니다.
- GitHub Pages 단일 HTML 제약을 고려해 해시 라우팅(`index.html#tool-id`)으로 도구별 화면 전환을 구현했습니다.
