// 월마감 현황 기본 목록 (업무 일정표 + 2026년 9·10월 부속자료 폴더 기준). 처음 열 때 한 번 들어가고, 이후 수정·삭제는 사이트에서.
// id가 고정이라 여러 기기에서 넣어도 하나로 합쳐짐
// file·folder 의 {YYYY} {YY} {MM} {M} {YYMM} {END}(말일) 는 보고 있는 달로 바뀜. {MM+1}처럼 달을 더할 수 있음
// needs[].rd = true → ERP 등에서 내가 뽑는 로우데이터(RD), 아니면 다른 사람에게 받는 파일
const P = '4.경영회의\\경영회의 부속자료\\{YYYY}년\\{M}월';
const DH_DATA = '5.KPI\\대흥\\대흥KPI Data\\{YYMM}';

export const CLOSING_SEED = [
 { id: 'seed01', order: 1, name: '거래처 마스터', file: '거래처 Master_{YY}.{MM}.{END}_원본_v1a', folder: P, timing: '전사 매출마감 이후', auto: '',
  note: '발송용 파일도 같이 저장: 거래처 Master_{YY}.{MM}.{END}_원본_v1a(발송용) — check·Summary·매출통합 등 내부 시트를 뺀 본\n시트: 합산·품목별(팽이·느타리·새송이·표고·양송이·기타)·vs.B·매핑·단위원가·매출통합',
  needs: [
   { id: 'seed01n01', name: '데일리리포트 마감 원본 ((Dr.AG) Daily Report_{YYYY}.{MM}_마감_원본)', dept: '전략기획', person: '' },
  ] },
 { id: 'seed02', order: 2, name: '데일리리포트 원본', file: '(Dr.AG) Daily Report_{YYYY}.{MM}_마감_원본', folder: P, timing: '전사 매출마감 이후', auto: '',
  note: '일별 원본은 Daily Report\\daily report 원본\\{M}월, 로우데이터는 Daily Raw 폴더(daily report_매출통합_집계, DSG Raw)',
  needs: [
   { id: 'seed02n01', name: 'ERP 매출·매입 RD (매출통합 · DSG_PRD)', dept: 'ERP', person: '', rd: true },
  ] },
 { id: 'seed03', order: 3, name: '대흥KPI_CF미반영', file: 'DH_KPI_{YYYY}년{MM}월_v1', folder: '5.KPI\\대흥\\{YYMM}', timing: '전사 매출마감 이후', auto: '',
  note: `받은 자료·판관비는 ${DH_DATA} 폴더에 모음`,
  needs: [
   { id: 'seed03n01', name: '판관비 (DH_KPI_{YYYY}년{MM}월 판관비_v1)', dept: '전략기획', person: '' },
   { id: 'seed03n02', name: 'ERP 손익·계정별원장 RD', dept: 'ERP', person: '', rd: true },
  ] },
 { id: 'seed04', order: 4, name: '대흥KPI_CF반영', file: 'DH_KPI_{YYYY}년{MM}월_v1b', folder: '5.KPI\\대흥\\{YYMM}', timing: 'CF자료 수취 이후', auto: '',
  note: `받은 자료는 ${DH_DATA} · 김해정차장님 자료는 부속자료 '수신_김해정차장님_Capex,CF' 폴더`,
  needs: [
   { id: 'seed04n01', name: 'FDA청소 ({M}월 FDA청소)', dept: '', person: '김도연차장님' },
   { id: 'seed04n02', name: 'HR자료 (2026경영회의 HR 자료_1월~12월)', dept: '', person: '조기민부장님' },
   { id: 'seed04n03', name: '일근태현황 (일근태현황_월별 잔업시간 집계_대흥_아팜머({M}월))', dept: '전략기획', person: '' },
   { id: 'seed04n04', name: '선급금/선급비용 ({YY}년 {M}월 선급금선급비용)', dept: '전략기획', person: '' },
   { id: 'seed04n05', name: '연결조정내역 ({YY}년 {M}월 내부거래)', dept: '', person: '김해정차장님' },
   { id: 'seed04n06', name: 'CF ({YY}년{M}월CF)', dept: '', person: '김해정차장님' },
   { id: 'seed04n07', name: 'Sales estimate ({YY}년 {MM}월 Sales estimate_v1)', dept: '전략기획', person: '' },
   { id: 'seed04n08', name: '판관비 (DH_KPI_{YYYY}년{MM}월 판관비_v1)', dept: '전략기획', person: '' },
   { id: 'seed04n09', name: 'CAPEX (0. 통합 Capex Actual {YYYY}. v4({M}월))', dept: '', person: '김해정차장님' },
   { id: 'seed04n10', name: '조정사항 (BI_조정사항_…_{YYMM}, 스마트팜)', dept: '', person: '조기민부장님' },
   { id: 'seed04n11', name: '전사 인원현황 (drag_전사인원현황_월_{YY}{MM}{END})', dept: '', person: '조기민부장님' },
  ] },
 { id: 'seed05', order: 5, name: '베트남KPI', file: 'KPI_베트남_{YYMM}_v1a', folder: '6.전략기획\\베트남\\베트남\\베트남 KPI\\{YYYY}_{MM}', timing: '삼도자료', auto: '',
  note: 'FS를 받으면 {YYMM}_베트남 로우데이터 통합_자동 파일에 붙여 회차별·수불·폐기를 만든 뒤 KPI에 반영\n수기: 허브 B3 당월, aT 지원액(운반비·라벨비·창고·인건비), GL 수기태그',
  needs: [
   { id: 'seed05n01', name: '삼도 결산자료 (DR. AG_FS_{Mon} {YYYY}.xls)', dept: '삼도회계법인', person: '', rd: true },
   { id: 'seed05n02', name: 'aT 지원액·인원현황', dept: '베트남법인', person: '' },
  ] },
 { id: 'seed06', order: 6, name: '운반비_집계', file: '운반비_집계_{YYMM}_v1a_검증', folder: P, timing: '운영팀자료', auto: 'freight',
  needs: [
   { id: 'seed06n01', name: '운반비 집계 (운반비_집계_{YYMM})', dept: '운영팀', person: '오수찬 부장님' },
   { id: 'seed06n02', name: 'ERP 계정별원장 – 운반보관비 1510424001~3', dept: 'ERP', person: '', rd: true },
  ] },
 { id: 'seed07', order: 7, name: '내부거래_업무위탁', file: '내부거래_업무위탁_{YYYY}.{MM}_v1', folder: P, timing: '재무팀 자료', auto: 'internal',
  needs: [
   { id: 'seed07n01', name: 'ERP 계정별원장 – 비용·수익 (최근 6개월)', dept: 'ERP', person: '', rd: true },
  ] },
 { id: 'seed08', order: 8, name: '전력비 및 에너지 사용현황', file: '닥터애그_전력비 및 에너지 사용 현황_{YYMM}_v1a', folder: P, timing: '전사 전력비 자료', auto: 'energy',
  note: "완성본은 전략기획·운영기획·권윤주상무님께 발신 (부속자료 '발신_전력비' 폴더)\n담당: 이형철 대리님, 정연주 차장님, 양창모 차장님",
  needs: [
   { id: 'seed08n01', name: '(그룹사 이천지역) {M}월 전력비 사용 현황(집계)', dept: '', person: '' },
   { id: 'seed08n02', name: '{YYMM} 전력비 및 에너지 사용 현황_청도', dept: '', person: '' },
   { id: 'seed08n03', name: '{YY}년{MM}월_재배용_수도광열비_사용현황', dept: '', person: '' },
  ] },
 { id: 'seed09', order: 9, name: '원재료비', file: '{M}월 Raw Material_v1a', folder: P, timing: '전사 마감 이후', auto: 'rawmat',
  note: '새해 열은 수기 · 표고 금액 수기',
  needs: [
   { id: 'seed09n01', name: 'ERP 기간별입고현황 – 원재료 (대흥·송이애·군자 3개)', dept: 'ERP', person: '', rd: true },
  ] },
 { id: 'seed10', order: 10, name: 'CF&NWC', file: '{M}월 Cashflow&NWC_v1a', folder: P, timing: '전사 마감 이후', auto: '',
  note: "시트 1개 'Cashflow&NWC_12개월' — 김해정차장님 자료는 '수신_김해정차장님_Capex,CF' 폴더",
  needs: [
   { id: 'seed10n01', name: 'ERP 재무상태표·계정 잔액 RD', dept: 'ERP', person: '', rd: true },
   { id: 'seed10n02', name: '자금일보 (자금일보(ver.{YYYY}.{MM}.{END}))', dept: '', person: '김해정차장님' },
   { id: 'seed10n03', name: 'Net debt (Drag {YYYY}년 KPI_net debt)', dept: '', person: '김해정차장님' },
  ] },
 { id: 'seed11', order: 11, name: '포장소모품비', file: '{M}월_소모품비 확인_v1a', folder: P, timing: '전사 마감 이후', auto: 'packing',
  note: '포장소모품비 예산정리 시트는 수기',
  needs: [
   { id: 'seed11n01', name: 'ERP 기간별입고현황 (수불o)', dept: 'ERP', person: '', rd: true },
   { id: 'seed11n02', name: 'ERP 계정별원장 – 포장소모품비 (수불x)', dept: 'ERP', person: '', rd: true },
  ] },
 { id: 'seed12', order: 12, name: '제조원가명세서 군자(이천)', file: 'BI_제조원가명세서_군자(이천)_{YYYY}.{MM}_v1a', folder: P, timing: 'KPI 이후', auto: '',
  note: '군자 KPI 파일(5.KPI\\군자\\{YYMM})을 외부 연결로 참조 · 시트: vs. bgt·COGM·Monthly P&L·INV.L·GL',
  needs: [
   { id: 'seed12n01', name: '군자 KPI ((군자){YYYY}.{MM}_KPI_v1a)', dept: '', person: '양창모차장님' },
   { id: 'seed12n02', name: 'ERP 군자 계정별원장(GL)·재고(INV.L) RD', dept: 'ERP', person: '', rd: true },
  ] },
 { id: 'seed13', order: 13, name: 'AR/AP', file: '(대흥) {YY}년{MM}월 거래처별 외상매입매출', folder: P, timing: 'KPI 이후', auto: 'arap',
  note: '송이애 원장형도 자동 생성(📊 파일 생성 → 채권·채무) · 군자는 받은 파일 그대로',
  needs: [
   { id: 'seed13n03', name: 'ERP 거래처원장 – 외상매출금 11110401·외상매입금 12100101', dept: 'ERP', person: '', rd: true },
   { id: 'seed13n01', name: '송이애 AR/AP ((송이애) {M}월 외상매출매입_{YYYY}.{MM})', dept: '', person: '전호종대리님' },
   { id: 'seed13n02', name: '군자 AR/AP (군자_{YY}년 {M}월 채권채무 등)', dept: '', person: '양창모차장님' },
  ] },
 { id: 'seed14', order: 14, name: '수출지원금', file: '{YYYY}년_수출지원금_계획대비 실적_{MM}{END}', folder: '6.전략기획\\수출', timing: '', auto: '',
  needs: [
   { id: 'seed14n01', name: 'ERP 계정별원장 – 수출지원금 (계정원장 시트)', dept: 'ERP', person: '', rd: true },
  ] },
 { id: 'seed15', order: 15, name: 'CAPEX 재분류 검증', file: '{YYYY}년{MM}월_Dr.AG_Capex_재분류_검증_v1a', folder: P, timing: 'CAPEX 자료 받으면 바로 가능', auto: 'capex',
  note: '송이애·군자 CAPEX Actual은 KPI 파일 보고 수기',
  needs: [
   { id: 'seed15n01', name: '통합 Capex Actual (0. 통합 Capex Actual {YYYY}. v4({M}월))', dept: '', person: '김해정차장님' },
  ] },
 { id: 'seed16', order: 16, name: 'Sales estimate', file: '{YY}년 {MM+1}월 Sales estimate_v1', folder: '6.전략기획\\영업추정', timing: 'PPT', auto: '',
  note: '파일명 뒤에 영업 회의 날짜가 붙음 (예: _영업1021) · 완성본은 부속자료 폴더에도 복사',
  needs: [
   { id: 'seed16n01', name: '영업 추정 (YYMMDD_Montly estmate_drag_v1)', dept: '', person: '강준모 상무님, 김성수 상무님' },
   { id: 'seed16n02', name: '매출통합 (데일리리포트 매출 Master)', dept: '전략기획', person: '', rd: true },
  ] },
 { id: 'seed17', order: 17, name: '익월 생산추정', file: '', folder: '', timing: 'PPT', auto: '',
  needs: [
   { id: 'seed17n01', name: '생산회의자료, 새송이공간관리파일, 각 담당자 확인', dept: '', person: '최민석과장님(새송이), 김기중부장님(팽이,느타리), 황교창부장님(표고)' },
  ] },
 { id: 'seed18', order: 18, name: '기타품목매출 break down', file: '기타품목매출 break down_{YYYY}.{MM}_v1', folder: P, timing: 'PPT', auto: '',
  needs: [
   { id: 'seed18n01', name: 'ERP 기타매출 RD (기타매출RD 시트)', dept: 'ERP', person: '', rd: true },
   { id: 'seed18n02', name: '매출 Master (데일리리포트 6. 매출 Master)', dept: '전략기획', person: '' },
   { id: 'seed18n03', name: '중부 수불', dept: '', person: '' },
  ] },
 { id: 'seed19', order: 19, name: '통합KPI', file: 'Drag {YYYY}년 {M}월 KPI_v1a', folder: '5.KPI\\Drag KPI', timing: '각사 KPI 수령 후', auto: '',
  note: '각사 KPI 폴더: 5.KPI\\대흥|군자|송이애\\{YYMM} · raw materials P 시트는 원재료비 완성본에서',
  needs: [
   { id: 'seed19n01', name: '대흥 KPI (DH_KPI_{YYYY}년{MM}월)', dept: '', person: '성재윤' },
   { id: 'seed19n02', name: '군자 KPI ((군자){YYYY}.{MM}_KPI)', dept: '', person: '양창모차장님' },
   { id: 'seed19n03', name: '송이애 KPI ((송이애) new KPI_{YYYY}년 {M}월)', dept: '', person: '전호종대리님' },
   { id: 'seed19n04', name: '중부/이천팜 KPI', dept: '', person: '전호종대리님' },
  ] },
 { id: 'seed20', order: 20, name: 'AR/AP 증감사유 분석', file: 'ARAP 증감분석_{YYMM}_v1a', folder: P, timing: '통합 KPI 이후', auto: '',
  note: '시트: 당월 · 대흥/송이애/군자농원 맵핑',
  needs: [
   { id: 'seed20n01', name: '대흥·송이애·군자 AR/AP 완성본', dept: '전략기획', person: '' },
  ] },
 { id: 'seed21', order: 21, name: 'CFO 기타 증감사유 분석', file: '', folder: '', timing: '통합 KPI 이후', auto: '', needs: [] },
 { id: 'seed22', order: 22, name: '전사 지급수수료 증감사유 분석', file: '지급수수료_비교_{YYMM}', folder: P, timing: '통합 KPI 이후', auto: '',
  note: '판관 버전도 같이 저장: 판관_지급수수료_비교_{YYMM} · 시트: 법인별 제조/판관 (대흥·송이애·군자·중부이천팜)',
  needs: [
   { id: 'seed22n01', name: 'ERP 계정별원장 – 지급수수료 (4개 법인 제조·판관)', dept: 'ERP', person: '', rd: true },
  ] },
 { id: 'seed23', order: 23, name: '경영회의 PPT 초안', file: 'Dr.AG 경영회의_{YYYY}년{MM}월_v1a', folder: '4.경영회의\\경영회의\\{M}월', timing: '통합 KPI 이후', auto: '',
  needs: [
   { id: 'seed23n01', name: '기타품목매출 break down (기타품목매출 break down_{YYYY}.{MM})', dept: '전략기획', person: '' },
  ] },
 { id: 'seed24', order: 24, name: '품목별 CM PPT', file: 'Dr.AG 경영회의_{YYYY}년{M}월_품목별 CM_v1a', folder: '4.경영회의\\경영회의\\{M}월', timing: '통합 KPI 이후', auto: '',
  needs: [
   { id: 'seed24n01', name: '통합 KPI (품목별 CM 시트) · 거래처 Master', dept: '전략기획', person: '' },
  ] },
 { id: 'seed25', order: 25, name: '경매단가 업데이트(오아시스)', file: '', folder: '', timing: '', auto: '', needs: [] },
 // ↓ 2026-11-19 추가: 목록엔 없었지만 매달 부속자료 폴더에 만들어지는 자료
 { id: 'seed26', order: 26, name: '선급금·선급비용', file: '{YY}년 {M}월 선급금선급비용', folder: DH_DATA, timing: '전사 마감 이후', auto: 'prepay',
  note: '대흥 KPI(CF반영)에 들어가는 자료',
  needs: [
   { id: 'seed26n01', name: 'ERP 거래처원장 – 선급금 11110800', dept: 'ERP', person: '', rd: true },
  ] },
 { id: 'seed27', order: 27, name: '내부거래 검증', file: '11.▲내부거래 검증({MM}월)', folder: P, timing: '전사 마감 이후', auto: '',
  note: '시트: 닥터애그·송이애·군자농원 매출/입고 · 현황판 · 단위매칭올인원',
  needs: [
   { id: 'seed27n01', name: 'ERP 매출현황·입고현황 (닥터애그·송이애·군자농원)', dept: 'ERP', person: '', rd: true },
   { id: 'seed27n02', name: '내부거래 ({YY}년 {M}월 내부거래)', dept: '', person: '김해정차장님' },
  ] },
 { id: 'seed28', order: 28, name: '전사 매출매입', file: '전사 매출매입_{YYYY}.{MM}_v1a', folder: P, timing: '전사 매출마감 이후', auto: '',
  needs: [
   { id: 'seed28n01', name: 'ERP 매출·매입 RD (매출통합 · DSG_PRD)', dept: 'ERP', person: '', rd: true },
  ] },
 { id: 'seed29', order: 29, name: '전사 월 수불현황', file: '전사_월 수불현황_{YY}_v1', folder: P, timing: '전사 마감 이후', auto: '',
  note: '한 해 파일 하나에 매달 추가 (버전만 올림)',
  needs: [
   { id: 'seed29n01', name: 'ERP 수불 RD (대흥 RD · 군자RD 시트)', dept: 'ERP', person: '', rd: true },
   { id: 'seed29n02', name: '건표고버섯 생산현황', dept: '', person: '' },
  ] },
 { id: 'seed30', order: 30, name: '제조소모품비 비교', file: '제조소모품비_비교_{YYMM}', folder: P, timing: '전사 마감 이후', auto: '',
  needs: [
   { id: 'seed30n01', name: 'ERP 계정별원장 – 소모품비 (대흥·송이애·군자·중부이천팜)', dept: 'ERP', person: '', rd: true },
  ] },
 { id: 'seed31', order: 31, name: '단위 운반비', file: '단위 운반비_{YYYY}{MM}', folder: P, timing: '운반비 검증 이후', auto: '',
  note: '시트: 채널 대시보드 · kg당 운반비 · 매출물량 · 운반비 · 납품처',
  needs: [
   { id: 'seed31n01', name: '운반비 검증 완성본 (운반비_집계_{YYMM}_v1a_검증)', dept: '전략기획', person: '' },
   { id: 'seed31n02', name: '운반비 현황 (청도·군자 운송정보)', dept: '운영팀', person: '' },
   { id: 'seed31n03', name: '매출물량 (데일리리포트 매출통합)', dept: '전략기획', person: '', rd: true },
  ] },
 { id: 'seed32', order: 32, name: '포장소모품비 Target', file: 'Dr.AG_포장소모품비 Target_{YYYY}.{MM}_v1a', folder: P, timing: '영업기획 자료 받은 후', auto: '',
  needs: [
   { id: 'seed32n01', name: '상품 포장비 (상품_포장비_Monthly_{YYYY})', dept: '영업기획팀', person: '' },
   { id: 'seed32n02', name: '제품 포장비 (제품_포장비_월간_{YYYY}{MM})', dept: '영업기획팀', person: '' },
  ] },
 { id: 'seed33', order: 33, name: '전력비·수도광열비 계획대비 증감', file: '계획대비계정증감_전력비, 수도광열비_{YY}년{MM}월_v1a', folder: P, timing: '전력비 사용현황 이후', auto: '',
  note: '시트: 품목비교 · 법인비교 · 전력비/수도광열비 bgt · RD · 매핑',
  needs: [
   { id: 'seed33n01', name: 'ERP 계정별원장 – 전력비·수도광열비 (전력비RD·수도광열비RD 시트)', dept: 'ERP', person: '', rd: true },
   { id: 'seed33n02', name: '사업계획 수도광열비 (05-1 {YYYY} 사업계획_재배용 수도광열비)', dept: '전략기획', person: '' },
  ] },
];

// 예전 기본 목록에서 바뀐 항목 (이미 목록이 있는 기기에 한 번 반영할 때, 사용자가 손대지 않은 값만 바꿈)
export const SEED_OLD = {
  seed01: { file: '거래처 Master_26.09.30_원본_v1a' },
  seed02: { file: '(Dr.AG) Daily Report_2026.10.00_원본_v1' },
  seed03: { file: 'DH_KPI_2026년00월_v1' },
  seed04: { file: '' },
  seed05: { file: 'KPI_베트남_2600_v1a' },
  seed06: { file: '운반비_집계_2600_v1a_검증' },
  seed07: { file: '내부거래_업무위탁_2026.2_v1' },
  seed08: { name: '전력비, 수도광열비', file: '계획대비계정증감_전력비, 수도광열비_26년07월_v1a 2607 전력비 및 에너지 사용 현황' },
  seed09: { file: '0월 Raw Material_v1' },
  seed10: { file: '0월 Cashflow&NWC-v1' },
  seed11: { file: '0월_소모품비 확인_v1a' },
  seed12: { file: 'BI_제조원가명세서_군자(이천)_202600_v1' },
  seed13: { file: '(대흥) 26년00월 거래처별 외상매입매출' },
  seed14: { file: '' }, seed15: { file: '2026년00월_Dr.AG_Capex_재분류_검증_v1' },
  seed16: { file: '26년 00월 Sales estimate_v1_영업0923' },
  seed18: { file: '기타품목매출 break down_2026.00_v1' },
  seed19: { file: '' }, seed20: { file: '' }, seed22: { file: '' }, seed23: { file: '' }, seed24: { file: '' },
};
// 예전 필요파일 이름 → 지금 이름 (같은 id면 이름만 손대지 않았을 때 바꿈)
export const NEED_OLD = {
  seed01n01: '데일리리포트 마감 원본 (DailyReport_2026.00)',
  seed02n01: 'ERP',
  seed04n01: 'FDA청소 (0월 FDA청소)',
  seed04n02: 'HR자료 (2026경영회의 HR자료_1월~12월)',
  seed04n03: '일근태현황 (일근태현황_월별 잔업시간 집계_대흥_아팜머(0월))',
  seed04n04: '선급금/선급비용 (00년 0월 선급금선급비용)',
  seed04n05: '연결조정내역 (00년 0월 내부거래)',
  seed04n06: 'CF (00년 0월 CF)',
  seed04n07: 'Sales estimate (00년 0월 Sales estimate_v1)',
  seed04n08: '판관비 (DH_KPI_2000년 0월 판관비_v1)',
  seed04n09: 'CAPEX (통합 Capex Actual 2026.)',
  seed05n01: 'DR. AG_FS_Jul 2026',
  seed06n01: '운반비_집계_2608',
  seed07n01: 'ERP',
  seed08n01: '계획대비계정증감_전력비, 수도광열비_26년07월_v1a 2607 전력비 및 에너지 사용 현황',
  seed09n01: 'ERP', seed10n01: 'ERP', seed11n01: 'ERP',
  seed12n01: '군자 KPI',
  seed13n01: '송이애 AR/AP', seed13n02: '군자 AR/AP',
  seed16n01: '260923_Montly estmate_drag_v1',
  seed18n01: 'ERP',
  seed23n01: '기타품목매출 break down (기타품목매출 break down_2000.00)',
};
