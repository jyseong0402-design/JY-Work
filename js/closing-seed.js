// 월마감 현황 기본 목록 (업무 일정표 기준). 처음 열 때 한 번 들어가고, 이후 수정·삭제는 사이트에서.
// id가 고정이라 여러 기기에서 넣어도 하나로 합쳐짐
export const CLOSING_SEED = [
 {
  "id": "seed01",
  "order": 1,
  "name": "거래처 마스터",
  "file": "거래처 Master_26.09.30_원본_v1a",
  "timing": "전사 매출마감 이후",
  "auto": "",
  "needs": [
   {
    "id": "seed01n01",
    "name": "데일리리포트 마감 원본 (DailyReport_2026.00)",
    "producer": "전략기획"
   }
  ]
 },
 {
  "id": "seed02",
  "order": 2,
  "name": "데일리리포트 원본",
  "file": "(Dr.AG) Daily Report_2026.10.00_원본_v1",
  "timing": "전사 매출마감 이후",
  "auto": "",
  "needs": [
   {
    "id": "seed02n01",
    "name": "ERP",
    "producer": "전략기획"
   }
  ]
 },
 {
  "id": "seed03",
  "order": 3,
  "name": "대흥KPI_CF미반영",
  "file": "DH_KPI_2026년00월_v1",
  "timing": "전사 매출마감 이후",
  "auto": "",
  "needs": []
 },
 {
  "id": "seed04",
  "order": 4,
  "name": "대흥KPI_CF반영",
  "file": "",
  "timing": "CF자료 수취 이후",
  "auto": "",
  "needs": [
   {
    "id": "seed04n01",
    "name": "FDA청소 (0월 FDA청소)",
    "producer": "김도연차장님"
   },
   {
    "id": "seed04n02",
    "name": "HR자료 (2026경영회의 HR자료_1월~12월)",
    "producer": "조기민부장님"
   },
   {
    "id": "seed04n03",
    "name": "일근태현황 (일근태현황_월별 잔업시간 집계_대흥_아팜머(0월))",
    "producer": "전략기획"
   },
   {
    "id": "seed04n04",
    "name": "선급금/선급비용 (00년 0월 선급금선급비용)",
    "producer": "전략기획"
   },
   {
    "id": "seed04n05",
    "name": "연결조정내역 (00년 0월 내부거래)",
    "producer": "김해정차장님"
   },
   {
    "id": "seed04n06",
    "name": "CF (00년 0월 CF)",
    "producer": "김해정차장님"
   },
   {
    "id": "seed04n07",
    "name": "Sales estimate (00년 0월 Sales estimate_v1)",
    "producer": "전략기획"
   },
   {
    "id": "seed04n08",
    "name": "판관비 (DH_KPI_2000년 0월 판관비_v1)",
    "producer": "전략기획"
   },
   {
    "id": "seed04n09",
    "name": "CAPEX (통합 Capex Actual 2026.)",
    "producer": "김해정차장님"
   }
  ]
 },
 {
  "id": "seed05",
  "order": 5,
  "name": "베트남KPI",
  "file": "KPI_베트남_2600_v1a",
  "timing": "삼도자료",
  "auto": "",
  "needs": [
   {
    "id": "seed05n01",
    "name": "DR. AG_FS_Jul 2026",
    "producer": "삼도회계법인"
   }
  ]
 },
 {
  "id": "seed06",
  "order": 6,
  "name": "운반비_집계",
  "file": "운반비_집계_2600_v1a_검증",
  "timing": "운영팀자료",
  "auto": "freight",
  "needs": [
   {
    "id": "seed06n01",
    "name": "운반비_집계_2608",
    "producer": "오수찬 부장님"
   }
  ]
 },
 {
  "id": "seed07",
  "order": 7,
  "name": "내부거래_업무위탁",
  "file": "내부거래_업무위탁_2026.2_v1",
  "timing": "재무팀 자료",
  "auto": "internal",
  "needs": [
   {
    "id": "seed07n01",
    "name": "ERP",
    "producer": "ERP"
   }
  ]
 },
 {
  "id": "seed08",
  "order": 8,
  "name": "전력비, 수도광열비",
  "file": "계획대비계정증감_전력비, 수도광열비_26년07월_v1a 2607 전력비 및 에너지 사용 현황",
  "timing": "전사 전력비 자료",
  "auto": "",
  "needs": [
   {
    "id": "seed08n01",
    "name": "계획대비계정증감_전력비, 수도광열비_26년07월_v1a 2607 전력비 및 에너지 사용 현황",
    "producer": "이형철 대리님, 정연주 차장님, 양창모 차장님"
   }
  ]
 },
 {
  "id": "seed09",
  "order": 9,
  "name": "원재료비",
  "file": "0월 Raw Material_v1",
  "timing": "전사 마감 이후",
  "auto": "",
  "needs": [
   {
    "id": "seed09n01",
    "name": "ERP",
    "producer": "ERP"
   }
  ]
 },
 {
  "id": "seed10",
  "order": 10,
  "name": "CF&NWC",
  "file": "0월 Cashflow&NWC-v1",
  "timing": "전사 마감 이후",
  "auto": "",
  "needs": [
   {
    "id": "seed10n01",
    "name": "ERP",
    "producer": "ERP"
   }
  ]
 },
 {
  "id": "seed11",
  "order": 11,
  "name": "포장소모품비",
  "file": "0월_소모품비 확인_v1a",
  "timing": "전사 마감 이후",
  "auto": "packing",
  "needs": [
   {
    "id": "seed11n01",
    "name": "ERP",
    "producer": "ERP"
   }
  ]
 },
 {
  "id": "seed12",
  "order": 12,
  "name": "제조원가명세서 군자(이천)",
  "file": "BI_제조원가명세서_군자(이천)_202600_v1",
  "timing": "KPI 이후",
  "auto": "",
  "needs": [
   {
    "id": "seed12n01",
    "name": "군자 KPI",
    "producer": "양창모차장님"
   }
  ]
 },
 {
  "id": "seed13",
  "order": 13,
  "name": "AR/AP",
  "file": "(대흥) 26년00월 거래처별 외상매입매출",
  "timing": "KPI 이후",
  "auto": "arap",
  "needs": [
   {
    "id": "seed13n01",
    "name": "송이애 AR/AP",
    "producer": "전호종대리님"
   },
   {
    "id": "seed13n02",
    "name": "군자 AR/AP",
    "producer": "양창모차장님"
   }
  ]
 },
 {
  "id": "seed14",
  "order": 14,
  "name": "수출지원금",
  "file": "",
  "timing": "",
  "auto": "",
  "needs": []
 },
 {
  "id": "seed15",
  "order": 15,
  "name": "CAPEX 재분류 검증",
  "file": "2026년00월_Dr.AG_Capex_재분류_검증_v1",
  "timing": "CAPEX 자료 받으면 바로 가능",
  "auto": "",
  "needs": []
 },
 {
  "id": "seed16",
  "order": 16,
  "name": "Sales estimate",
  "file": "26년 00월 Sales estimate_v1_영업0923",
  "timing": "PPT",
  "auto": "",
  "needs": [
   {
    "id": "seed16n01",
    "name": "260923_Montly estmate_drag_v1",
    "producer": "강준모 상무님, 김성수 상무님"
   }
  ]
 },
 {
  "id": "seed17",
  "order": 17,
  "name": "익월 생산추정",
  "file": "",
  "timing": "PPT",
  "auto": "",
  "needs": [
   {
    "id": "seed17n01",
    "name": "생산회의자료, 새송이공간관리파일, 각 담당자 확인",
    "producer": "최민석과장님(새송이), 김기중부장님(팽이,느타리), 황교창부장님(표고)"
   }
  ]
 },
 {
  "id": "seed18",
  "order": 18,
  "name": "기타품목매출 break down",
  "file": "기타품목매출 break down_2026.00_v1",
  "timing": "PPT",
  "auto": "",
  "needs": [
   {
    "id": "seed18n01",
    "name": "ERP",
    "producer": "ERP"
   }
  ]
 },
 {
  "id": "seed19",
  "order": 19,
  "name": "통합KPI",
  "file": "",
  "timing": "각사 KPI 수령 후",
  "auto": "",
  "needs": [
   {
    "id": "seed19n01",
    "name": "대흥 KPI",
    "producer": "성재윤"
   },
   {
    "id": "seed19n02",
    "name": "군자 KPI",
    "producer": "양창모차장님"
   },
   {
    "id": "seed19n03",
    "name": "송이애 KPI",
    "producer": "전호종대리님"
   },
   {
    "id": "seed19n04",
    "name": "중부/이천팜 KPI",
    "producer": "전호종대리님"
   }
  ]
 },
 {
  "id": "seed20",
  "order": 20,
  "name": "AR/AP 증감사유 분석",
  "file": "",
  "timing": "통합 KPI 이후",
  "auto": "",
  "needs": []
 },
 {
  "id": "seed21",
  "order": 21,
  "name": "CFO 기타 증감사유 분석",
  "file": "",
  "timing": "통합 KPI 이후",
  "auto": "",
  "needs": []
 },
 {
  "id": "seed22",
  "order": 22,
  "name": "전사 지급수수료 증감사유 분석",
  "file": "",
  "timing": "통합 KPI 이후",
  "auto": "",
  "needs": []
 },
 {
  "id": "seed23",
  "order": 23,
  "name": "경영회의 PPT 초안",
  "file": "",
  "timing": "통합 KPI 이후",
  "auto": "",
  "needs": [
   {
    "id": "seed23n01",
    "name": "기타품목매출 break down (기타품목매출 break down_2000.00)",
    "producer": "전략기획"
   }
  ]
 },
 {
  "id": "seed24",
  "order": 24,
  "name": "품목별 CM PPT",
  "file": "",
  "timing": "통합 KPI 이후",
  "auto": "",
  "needs": []
 },
 {
  "id": "seed25",
  "order": 25,
  "name": "경매단가 업데이트(오아시스)",
  "file": "",
  "timing": "",
  "auto": "",
  "needs": []
 }
];
