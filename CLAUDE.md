# JY 업무노트 – 작업 규칙

갤럭시 탭 S8(S펜) + PC에서 쓰는 업무용 PWA. 빌드 없이 정적 파일로 GitHub Pages(`main` / root)에서 배포됨.
사이트: https://jyseong0402-design.github.io/JY-Work/

## 업데이트할 때마다 꼭
1. `js/version.js` — `APP_VERSION`을 올리고 `CHANGES` 맨 위에 사용자용 변경 내역(한국어, 짧게)을 추가
   (앱이 이를 보고 '새 버전 적용됨' 알림, 좌측 하단 버전, 실행 중 '지금 적용' 안내를 띄움)
2. `sw.js` — `CACHE` 이름 숫자 올리기, 새 파일이 생기면 `ASSETS`에 추가
3. 작업 브랜치에 커밋·푸시 후 **`main`에도 바로 푸시** (사용자가 main 직접 반영을 허락함, 2026-11-10)

## 구조
- `js/app.js` 화면 동작 / `js/store.js` 로컬 저장(+변경 추적) / `js/sync.js` Firebase 동기화
- `js/ink.js` S펜 필기 엔진 / `js/inkdb.js` 필기 IndexedDB / `js/summarize.js` 오프라인 요약 / `js/ai.js` Claude API
- `js/vendor/` 번들된 라이브러리(Firebase, SheetJS, Anthropic SDK) – 외부 CDN 사용 안 함
- Firestore 경로 `users/{uid}/items`, `users/{uid}/ink` — 규칙은 `firestore.rules`(본인만 접근)

## 사용자
- 한국어로 답변. 비개발자이므로 GitHub/Firebase 조작은 화면 기준으로 단계별 안내.
