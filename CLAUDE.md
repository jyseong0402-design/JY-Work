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

- 메모(2026-11-18~): 누적/낱장 구분 없음. 예전 series 기록은 migrateSeries(빈 기록 삭제·나머지 series='' 일반 메모), 일정→회의메모는 eventId 같은 메모 하나를 이어서 씀. 필기 못 불러오면 inkBlocked로 저장 멈춤, inkdb '<id>~prev'에 이전 판, 설정 ✍️ 필기 점검(inkReport/restoreInk), 클라우드 ink는 700KB 조각(part)
- 메모 사진(2026-11-19~): 메모 item.photos=[{id,name,at}](목록만 items로 동기화), 사진 자체는 `js/photos.js` IndexedDB 'jywork-photos'(긴 변 1600px JPEG dataURL) + Firestore users/{uid}/photos(문서 id=사진 id, 700KB 넘으면 `_pN` 조각, pid로 조회). sync.js savePhoto/getPhoto(없으면 클라우드)/removePhoto, 로그인 때 uploadPendingPhotos. 📷 버튼·붙여넣기·끌어놓기 → 필기 쪽 안에 { t:'img', pid, x,y,w,h } 획으로 넣음(ink.js drawAll: 사진 먼저·필기 위, setImageSource, 지우개 제외, 'select' 도구로 옮기기·크기·🗑, undo 'imgmove'). 메모 닫을 때 쪽에 없는 사진은 removePhoto(inkBlocked면 안 함), 1.20.0 사진 줄 사진은 placeLoosePhotos로 쪽 안에 옮김. #dlgPhoto 크게 보기·저장
- 메뉴(2026-11-16~): 홈(dashboard: 달력 #calMode 달력/목록, #homeSearch→search 뷰, #homeClosing=board.homeClosingHTML) · 메모(meetings) · 월마감 현황(board) · RD 보관함(rd) · 자료생성(auto) · 설정. tasks/search 뷰는 메뉴 없이 [data-go-view]로 이동
- `js/kpi.js` 📈 KPI 현황(view-kpi, 2026-11-20~): KPI 엑셀의 'Summary…' 시트만 SheetJS로 읽음(bookSheets→sheets 옵션, !ref를 A1부터로 바꿔 열 번호 맞춤). 'YYYYA'/'YYYYB' 줄+아래 월 줄로 실적·예산 열(처음 나온 열만), B3=결산월, D열=대항목(grp)·E열=세부. store kind 'kpi' {id:k연도_제목hash, title, year, month, file, rows[{label,top,grp,a[12],b[12]}], notes{월:메모}} → items로 동기화. model()이 라벨 정규식으로 매출(첫 매출액=품목, 둘째=거래처)·물량·단가·원가·GP·CM·관리비·EBIT·조정EBITDA를 찾음. 특이사항 findings(): 품목합≠거래처합 점검, 매출±10%, 품목/거래처 ±20%&1백만, 신규/중단, 단가±7%, GP%±5%p, 판매비·관리비±15%, 흑자/적자 전환, 예산 달성률 90~110% 밖, 연중 최고 매출. 그래프는 직접 그린 SVG
  · (1.22.0~) 메뉴 이름 '📈 베트남', 화면 표기 예산→BGT(사용자 요청). 🎯 Target: 'Target … PL' 시트(parseTarget: 연도 줄+'N월' 줄=월 실적 열, 그 오른쪽 첫 머리글 묶음=시나리오 'Target 3-1/3-2/10톤 COOP+WINN/20톤…' = 월 기준 목표치, B/C/D열=lv1/2/3, p=상위 라벨) → item.target{sheet,scen,rows[{label,lv,p,a[12],t[]}]}. tmodel·targetTable(당월·최근3개월 평균 vs 선택 시나리오, ui.tgt='jy-kpitgt')·targetFindings. 사용자: Target 대비가 가장 중요. 누계 보고 Ⅱ=Target 대비
  · 📝 경영회의 코멘트(#kpiMode cmt): 경영회의 장표 '베트남법인 실적' 표 = KPI 'PPT (2)' 시트 128행 묶음(parsePpt: 머리글 줄 Budget·Target + Budget 바로 왼쪽 Actual, 월 머리글은 날짜값일 수 있음, 표 바로 왼쪽 '(단위' 열=라벨, 숨긴 행 h=1 — 장표엔 숨긴 행 제외, cellStyles:true 필요). 원장 parseGl: 'GL(원장)' Account/Debit/Credit/year/month, 손익계정 5~8xx 월별 VND 합계 + 당월·전월 100만VND 이상 거래 400개(911 상대·profit and loss 제외), VND÷18.5=원. 코멘트 A매출액 B단위원가 C판매비 D관리비 E조정EBITDA: cmHeadline(단기 Target·BGT 대비), cmFacts(하위 행 전월 대비, 운반비 kg당, 거래처), cmGl(계정 증감·주요 거래), cmQuestions(현지 문의), cmAi(aiJson A~E). 메모 notes['cm{월}{A~E}']. 파일 읽기 약 12초
  · (1.22.1~) AI 없이 장표 형식 자동 초안(사용자: API 못 씀): 9월 장표(경영회의_2026년09월_v1p_R.pptx 37쪽) 문구를 따름 — cmHeadline "매출액 단기 Target 초과 달성 +15백만원, +37%," / "계획 대비 …"(장표는 BGT가 아니라 '계획'), 차이는 표에 보이는 반올림 값끼리 뺌, E는 Target 한 줄. cmBullets: A=거래처별 당월 매출·전월 대비+물량+단가, B·C·D=하위 행 전월 대비 큰 순 3개(B 원/kg 부호, C·D 절감은 부호 없이, 운반비는 원/kg), E=Target 대비 GP·판매비·관리비 영향. [사유] 자리는 현지 확인 후 채움. 메모가 비면 textarea에 초안 표시(.auto, 수정 시 저장). AI 버튼은 키 있을 때만
  · 💰 지원금 수익(subsidyCard, 월 분석 Target 카드 아래): 'Target … PL' 시트에서 라벨에 '지원' 있는 lv≤2 행(9월 파일 C154 'AT지원사업 수익', 7월~) + 하위 운반비·라벨비·창고임차료. 당월·전월·누계·매출 대비·판매비 지원율(지원/(판매비+지원))·지원 없을 때 조정 EBITDA. 원장 5·7xx 중 711 또는 적요 hỗ trợ/support 거래도 표시. aT 지원율(운송 80%, 창고·라벨 90%)은 사용자 aT 정산서 기준
  · 💬 물어보기(#kpiChat details): 사용자가 숫자 전송 허락(2026-11-20). kpiContext()=모든 kpi 항목의 Summary 행(실적·BGT)+Target 행+메모를 system에 넣고 aiChat(ai.js, messages 대화 최근 12개). 약 3만 자/질문
  · (1.21.2~) 누계 보고는 보고서 형식 article.rp 한 열: Ⅰ 핵심 요약(rp-tiles + periodFindings를 개선/부진/참고로 묶음 + 보고 메모) Ⅱ 손익 요약(누계·예산·달성률·월평균·당월·전월) Ⅲ 월별 손익(막대 값 표시 vlabels + 월×항목 표) Ⅳ 분기 Ⅴ 품목별 표+구성 Ⅵ 거래처별 표(거래 기간)+구성 Ⅶ 수익성·EBIT. 표 클래스 .rt(key/sub/tot/hl), 적자·음수 예산은 달성률 '-'
  · 📑 누계 보고(#kpiMode ytd, localStorage 'jy-kpimode'): 월 버튼=1~N월. renderYtd = 누계 카드(예산·달성률·월평균), periodFindings(월매출 흐름, 마지막 분기 vs 직전 분기, 누계 예산, 1Q→마지막Q GP%·관리비·판매비율·품목 비중·평균 단가, CM/EBIT 흑자 유지 시작월, 누계 3백만↑ 거래처 시작/중단, 품목합≠거래처합 달), 월별 손익표(PPT 형식, 누계 비율·단가는 합계로 재계산), 분기 비교(월평균), 품목·거래처 누적 막대, 수익성 선. 보고 메모 notes['ytdN'], 🖨 인쇄(@media print로 메뉴·버튼 숨김)
- `js/rdvault.js` 📥 RD 보관함: IndexedDB 'jywork-rd'(기기별, 요금 없음) 월별 보관, classify로 거래처원장/계정별원장/기간별입고현황/운반비 집계 인식(날짜 있는 RD는 그 달로), 지난달 파일=자료생성과 같은 저장소(jywork-closing/kind, jywork-freight/'last'), 파일명 _YYMM이 보관함 달 이상이면 사용 안 함, 한꺼번에 만들기+zip
- RD 보관함 📁 폴더 연결(rdvault.js, File System Access, PC 크롬·엣지): 연도 폴더 핸들을 IndexedDB 'jywork-folders'에 보관, 'N월_RD'→RD(그 달로 강제 보관), 'N-1월'→지난달 완성본(detectBook·운반비 시트로 종류 판단, 최근 수정본)→putSaved, 결과는 'N월' 폴더에 renameMonth(지난달 이름의 달 표기 교체) 이름으로 저장(같은 이름 있으면 _자동)
- `js/rdask.js` 💬 RD에 물어보기(RD 보관함 아래 #aqCard): 질문→AI(aiJson, SPEC_SCHEMA)가 조회 조건 JSON만 생성(보내는 것: 질문·열 이름·계정코드/계정명·사업장 목록, 금액 X)→브라우저에서 runSpec으로 필터·정렬·소계/요약→원래 RD 열 그대로 + 맨 끝 '법인' 열(corpOf: 사업장→대흥/군자/송이애/중부·이천팜/SPC). 14xxx=제조, 15xxx=판관. 엑셀(원장 행)은 toStyledExcel: 첫 RD 파일을 본보기로 다른 시트 삭제(Book.removeSheet)·머리글 줄 유지·데이터 줄 서식은 날짜 있는 첫 줄, 소계는 원본 합계 줄 서식(없으면 데이터 줄), '법인' 가상 열은 엑셀에 안 넣음. 요약(summary)만 SheetJS. API 키 없거나 AI 실패 시 localSpec(계정명 앞부분 매칭+앞의 제조/판관, 법인, 잔액 n억 이상, 합계/요약, 이어서 고치기)
- 화면: 「📊 자료생성」(view-auto) 한 메뉴 + 드롭다운(#autoType). 운반비=freight-ui(.auto-panel[data-panel=freight]), 나머지=closing-ui의 TYPES 목록
  · 새 자료 추가: closing.js에 build 함수 + closing-ui.js TYPES 한 줄 + index.html <option> 하나
  · 원칙: 기존 서식·수식 절대 유지, 이번 달만 추가(손대지 않은 시트는 원본 그대로 저장). 검증은 rd/integrity.py 방식(원래 셀 서식·수식 비교)
- `js/board.js` 📋 월마감 현황(view-board): store kind 'closing' = 업무 {name,file,timing,auto,needs[{id,name,dept,person}](예전 producer는 needOf가 사람/조직으로 나눠 읽음),order,marks{'YYYY-MM':{done,recv{needId:date}}}}
  · 기본 목록 js/closing-seed.js(고정 id seedNN, 사용자가 공개 저장소 노출을 알고 넣기로 함 2026-11-15): 기기마다 처음 한 번 store.seed(오래된 updatedAt → 다른 기기 수정·삭제가 이김). 📥/💾 목록 파일(JSON)은 백업용. 담당자·target·done은 사이트에 두지 않음
  · (2026-11-19~) file=파일명 규칙({YYYY}{YY}{MM}{M}{YYMM}{END}{Mon}, {MM+1} 가능), folder=저장 위치, note=작업 메모, needs[].rd=직접 뽑는 로우데이터. 보는 달=일하는 달, 이름은 전달(결산월)로 채움(closeOf). 기존 기기는 'jywork-closing-v3' 한 번 upgradeSeed(SEED_OLD/NEED_OLD와 같을 때만 교체, seed26~ 추가)
  · `js/forms.js` 📄 양식: IndexedDB 'jywork-forms'(files/meta, key=업무 id, 기기별). 📂 한꺼번에 등록 = matchFiles(파일명 규칙 정규식, 뒤에 붙은 말 적은 것→최근 달→최근 수정본). 📊 파일 생성 열 때 xlsx 양식이 있으면 setClosingPrev/setFreightPrev로 ① 지난달 파일에 넣음
  · 📁 폴더에서 양식 자동 등록(autoFromFolder): 업무 폴더 핸들 'jywork-folders'/'work'(읽기 전용, 위 폴더를 골라도 '업무' 하위 탐색), 업무별 folder를 결산월로 채워 openPath → matchFiles(target=결산월, 없으면 최근 파일+표시), 같은 이름·수정시각이면 건너뜀
  · 📊 파일 생성 = #dlgMake 팝업: #view-auto의 .auto-pick/.auto-panel을 팝업으로 옮겼다가 닫으면 되돌림, 생성 완료 시 window 'backup:made' 이벤트
- `js/closing.js` 월마감 엔진 / `js/closing-ui.js` 화면 / `js/xlsxkit.js` 공용 XML 편집(열·행 삽입 시 수식 참조 이동, 시트 추가)
  · 채권채무: 거래처원장 11110401·12100101 → 'YY년MM월매출채권/매입채무' + 매출비교·매입비교 열 삽입(차액 열 이동, 새 거래처 '신규' 행)
  · 거래처원장 범위 확인: RD 기초(거래처별·합계) vs 지난달 월 시트 I열 잔액 대조 → 다르면 warnings 맨 앞 '범위 확인' 경고, 비교표 A1의 조회 범위 문구를 info로 안내, RD에 사업장 열 있으면 사업장 목록 표시
  · 채권채무 원장형(송이애, kind arapList): '매출채권'·'매입채무'·'맵핑' 시트, A열 'YYYY.M.' 달 묶음으로 RD 행(B~I) 아래에 추가, J·K 수식은 마지막 행 복사, 코드 열 타입은 지난 행 따라감, 맵핑 A/H열에 없는 거래처 경고. buildArAp가 자동 위임. RD 여러 개면 지난달 잔액=기초 일치율로 파일 선택(대흥 비교표형도 동일). 보관함 칸 'arapList'
  · CAPEX 재분류 검증(capex): 원본=대흥 '통합 Capex Actual'의 'CAPEX 닥터애그…' 시트(C 구분1·D 구분2·F 투자내역·투자내역 다음 12칸=1~12월, 원). 검증 '닥터애그(CAPEX Actual)' 맨 아래에 A 연도·B 월·D(본보기)·E·F·I + '{Y} Actual'(3행) 블록의 M월 열(천원). 구분은 지난 달들의 (투자내역·금액) 짝으로 원본→검증 재분류를 배워 적용, 없으면 기존 분류/앞부분 일치. 송이애·군자 Actual은 KPI 보고 수기, 새해 열 수기
  · 선급금: 11110800 → 'YY년MM월선급금' + 선급금증감 열 / 내부거래: 비용·수익 최근 6개월 계정 중 상대가 계열사인 행, D·상대회사는 이력 우선
  · 전력비·수도광열비(energy): '전력비'·'가스 등유 벙커c유 등' 시트, 블록(사용량/단가/요금)마다 년·월 행. RD 사용현황 2개(대흥 / 송이애·새송이·용인)의 계량기 열을 지난 12개월 값 일치로 짝지음(없으면 머리글+순번), 수식 칸은 유지, 달 행 없으면 모든 블록에 먼저 끼우고 위 행 수식 복사
  · Raw Material(rawmat): 'raw materials P' 1행 Act·2행 물량/매입액·6/7행 년월 열. 구분 '원재료'만, D열 구역→입고공장(RM_SECTIONS), E열 품목명(별칭 콘콥=콘컵, 케이포박=케이폭), '포' 단위는 C열 16kg로 =포×16, 표고=미배양 배지(생산입고+구매) 물량만·금액 수기. 새해 열은 수기
  · 포장소모품비 수불o는 기존 (구분·입고공장) 조합만, 여러 RD에 같은 조합이면 첫 파일만
  · 포장소모품비: 기간별입고현황→수불o, Mapping P열 계정→수불x, 회사별 시트 3칸 묶음+요약 열(예산정리 시트는 수기)

## 새 백업자료 추가 요청이 오면 (사용자가 '생성파일 + RD'를 줌)
1. 파일 구조 파악: 매달 무엇이 늘어나는지(월 시트 추가 / 행 추가 / 열 삽입), RD의 어떤 계정·열을 쓰는지
2. 애매한 규칙(포함·제외 기준, 수기 칸)은 먼저 질문
3. closing.js에 build 함수(기존 Book/appendRows/insertColumns 재사용) + detectBook 인식 + closing-ui TYPES + index.html <option>
4. 지난달 상태로 RD를 흉내 내 테스트: LibreOffice 재계산 값 비교 + 원래 셀 서식·수식 무결성 점검(0건이어야 함) + 브라우저 실행
5. 버전·캐시 올리고 브랜치·main 푸시, 사용자에게 무엇이 자동이고 무엇이 수기인지 안내

## 사용자
- 한국어로 답변. 비개발자이므로 GitHub/Firebase 조작은 화면 기준으로 단계별 안내.
