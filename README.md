# Kunhwa Tools Lab (MVP)

GitHub Pages에 바로 배포 가능한 정적 웹앱 초안입니다.

## 포함 기능 (업데이트)

1. PDF -> 이미지 변환기
2. 이미지 -> PDF 변환기
3. PDF 순서 변경(썸네일 드래그) 및 페이지 나누기
4. PDF 합치기
5. 이미지 크기 조정기
6. 이미지 확장자 포맷 변환기
7. 공정 작업시간 측정기 (공정명/고객사/메모/사진·동영상 첨부 및 촬영 + JSON/CSV 내보내기)
8. QR코드 생성기

## 실행 방법 (로컬)

`index.html`을 브라우저로 열고, 허브에서 각 도구 페이지로 이동해 사용합니다.

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

## UX 강화 사항

- 대용량 처리 진행률 바 표시
- 대용량 처리 취소 버튼 제공
- 기능별 에러/가이드 안내
- PDF 페이지 썸네일 Drag & Drop 정렬 UI
- PDF 분할 블록(구간 선택) UI
- 진행률 기반 ETA(남은 예상시간) 표시

## 참고

- `tools.mytory.net`처럼 도구별 개별 페이지 구조로 분리했습니다.
- 상단 공통 헤더에 고양이 마스코트를 추가했습니다.
