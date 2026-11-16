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
- `js/freight.js` 운반비 검증 엔진 / `js/freight-ui.js` 화면 — 지난 검증파일 xlsx를 JSZip으로 열어 XML을 직접 고침(서식 보존)
  · 원장: 운반보관비 3계정(1510424001~3), 차변만, 오아시스 차감(차변 음수)=0 / 집계 '현황' Total을 회사·납품처·운송사로 매칭
  · 채널: 집계 '납품처' 시트(B채널)와 현황 열을 금액으로 매칭 → 운반비 G=채널, J=분류(간선_외부매입→매입운반비, 간선_*→간선) / 검증 '납품처' 시트는 전월 vs 당월로 재작성
  · 오른쪽 표: Y열=1월…(Y+M-1)=M월. 후행물류비 표는 당월 열 추가(수식 끝 +수기값은 제거·경고), 선행물류비 표는 지난달 열 수식을 당월 열로 이동
  · 운반비N월 = 지난달 시트 복사 + H열 금액 교체, U열 '제안)' 사유, 계정원장에 N월 행(+Q~V 수식) 추가, mapping 추가 지원

- 메뉴(2026-11-16~): 홈(dashboard: 달력 #calMode 달력/목록, #homeSearch→search 뷰, #homeClosing=board.homeClosingHTML) · 메모(meetings) · 월마감 현황(board) · RD 보관함(rd) · 자료생성(auto) · 설정. tasks/search 뷰는 메뉴 없이 [data-go-view]로 이동
- `js/rdvault.js` 📥 RD 보관함: IndexedDB 'jywork-rd'(기기별, 요금 없음) 월별 보관, classify로 거래처원장/계정별원장/기간별입고현황/운반비 집계 인식(날짜 있는 RD는 그 달로), 지난달 파일=자료생성과 같은 저장소(jywork-closing/kind, jywork-freight/'last'), 파일명 _YYMM이 보관함 달 이상이면 사용 안 함, 한꺼번에 만들기+zip
- `js/rdask.js` 💬 RD에 물어보기(RD 보관함 아래 #aqCard): 질문→AI(aiJson, SPEC_SCHEMA)가 조회 조건 JSON만 생성(보내는 것: 질문·열 이름·계정코드/계정명·사업장 목록, 금액 X)→브라우저에서 runSpec으로 필터·정렬·소계/요약→원래 RD 열 그대로 + 맨 끝 '법인' 열(corpOf: 사업장→대흥/군자/송이애/중부·이천팜/SPC). 14xxx=제조, 15xxx=판관. 엑셀은 SheetJS(제목·머리글·소계·합계, 법인별 시트 옵션)
- 화면: 「📊 자료생성」(view-auto) 한 메뉴 + 드롭다운(#autoType). 운반비=freight-ui(.auto-panel[data-panel=freight]), 나머지=closing-ui의 TYPES 목록
  · 새 자료 추가: closing.js에 build 함수 + closing-ui.js TYPES 한 줄 + index.html <option> 하나
  · 원칙: 기존 서식·수식 절대 유지, 이번 달만 추가(손대지 않은 시트는 원본 그대로 저장). 검증은 rd/integrity.py 방식(원래 셀 서식·수식 비교)
- `js/board.js` 📋 월마감 현황(view-board): store kind 'closing' = 업무 {name,file,timing,auto,needs[{id,name,dept,person}](예전 producer는 needOf가 사람/조직으로 나눠 읽음),order,marks{'YYYY-MM':{done,recv{needId:date}}}}
  · 기본 목록 js/closing-seed.js(고정 id seedNN, 사용자가 공개 저장소 노출을 알고 넣기로 함 2026-11-15): 기기마다 처음 한 번 store.seed(오래된 updatedAt → 다른 기기 수정·삭제가 이김). 📥/💾 목록 파일(JSON)은 백업용. 담당자·target·done은 사이트에 두지 않음
  · 📊 파일 생성 = #dlgMake 팝업: #view-auto의 .auto-pick/.auto-panel을 팝업으로 옮겼다가 닫으면 되돌림, 생성 완료 시 window 'backup:made' 이벤트
- `js/closing.js` 월마감 엔진 / `js/closing-ui.js` 화면 / `js/xlsxkit.js` 공용 XML 편집(열·행 삽입 시 수식 참조 이동, 시트 추가)
  · 채권채무: 거래처원장 11110401·12100101 → 'YY년MM월매출채권/매입채무' + 매출비교·매입비교 열 삽입(차액 열 이동, 새 거래처 '신규' 행)
  · 거래처원장 범위 확인: RD 기초(거래처별·합계) vs 지난달 월 시트 I열 잔액 대조 → 다르면 warnings 맨 앞 '범위 확인' 경고, 비교표 A1의 조회 범위 문구를 info로 안내, RD에 사업장 열 있으면 사업장 목록 표시
  · 선급금: 11110800 → 'YY년MM월선급금' + 선급금증감 열 / 내부거래: 비용·수익 최근 6개월 계정 중 상대가 계열사인 행, D·상대회사는 이력 우선
  · 포장소모품비: 기간별입고현황→수불o, Mapping P열 계정→수불x, 회사별 시트 3칸 묶음+요약 열(예산정리 시트는 수기)

## 새 백업자료 추가 요청이 오면 (사용자가 '생성파일 + RD'를 줌)
1. 파일 구조 파악: 매달 무엇이 늘어나는지(월 시트 추가 / 행 추가 / 열 삽입), RD의 어떤 계정·열을 쓰는지
2. 애매한 규칙(포함·제외 기준, 수기 칸)은 먼저 질문
3. closing.js에 build 함수(기존 Book/appendRows/insertColumns 재사용) + detectBook 인식 + closing-ui TYPES + index.html <option>
4. 지난달 상태로 RD를 흉내 내 테스트: LibreOffice 재계산 값 비교 + 원래 셀 서식·수식 무결성 점검(0건이어야 함) + 브라우저 실행
5. 버전·캐시 올리고 브랜치·main 푸시, 사용자에게 무엇이 자동이고 무엇이 수기인지 안내

## 사용자
- 한국어로 답변. 비개발자이므로 GitHub/Firebase 조작은 화면 기준으로 단계별 안내.
