"""
엑셀 보고서 자동화 프로그램

폴더 안의 엑셀/CSV 파일을 모두 합치고, 원하는 기준으로 집계한 요약 보고서를
서식이 적용된 엑셀 파일 하나로 만들어 줍니다.

사용 예:
    python excel_report.py 입력폴더
    python excel_report.py 입력폴더 --group 지점 --value 매출액
    python excel_report.py 입력폴더 --group 지점 --column 품목 --value 매출액 --agg mean
    python excel_report.py --sample 샘플폴더      # 연습용 샘플 파일 생성

결과 파일 시트 구성:
    요약      : 파일별 행 수, 집계 조건
    합친데이터 : 모든 파일을 이어 붙인 표 (맨 앞 '원본파일' 열 포함)
    집계      : --group / --column / --value 기준 피벗 표 (합계 행·열 포함)
    중복행    : 모든 열 값이 같은 중복 행 목록 (있을 때만)

필요 패키지: pip install pandas openpyxl
"""
from __future__ import annotations

import argparse
import sys
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

EXTS = {".xlsx", ".xlsm", ".xls", ".csv"}
AGG_LABEL = {"sum": "합계", "mean": "평균", "count": "개수", "max": "최대", "min": "최소"}

HEADER_FILL = PatternFill("solid", fgColor="1F6F50")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TOTAL_FILL = PatternFill("solid", fgColor="DCEFE6")
THIN = Side(style="thin", color="D9E0DD")


def read_any(path: Path) -> list[tuple[str, pd.DataFrame]]:
    """파일 하나를 읽어 (이름, 표) 목록을 돌려준다. 엑셀은 시트마다 하나씩."""
    if path.suffix.lower() == ".csv":
        for enc in ("utf-8-sig", "cp949"):
            try:
                return [(path.name, pd.read_csv(path, encoding=enc))]
            except UnicodeDecodeError:
                continue
        raise ValueError("CSV 인코딩을 알 수 없습니다 (UTF-8, CP949 아님)")
    sheets = pd.read_excel(path, sheet_name=None)
    multi = len(sheets) > 1
    return [(f"{path.name}:{name}" if multi else path.name, df) for name, df in sheets.items()]


def collect(folder: Path, out: Path) -> tuple[pd.DataFrame, list[dict]]:
    files = sorted(p for p in folder.iterdir()
                   if p.suffix.lower() in EXTS and not p.name.startswith(("~$", "보고서_")) and p.resolve() != out.resolve())
    if not files:
        sys.exit(f"'{folder}' 폴더에 엑셀/CSV 파일이 없습니다.")
    frames, log = [], []
    for f in files:
        try:
            for name, df in read_any(f):
                df = df.dropna(how="all")
                df.columns = [str(c).strip() for c in df.columns]
                df.insert(0, "원본파일", name)
                frames.append(df)
                log.append({"원본파일": name, "행수": len(df), "열수": df.shape[1] - 1, "상태": "읽음"})
        except Exception as e:  # 한 파일이 깨져도 나머지는 계속 처리
            log.append({"원본파일": f.name, "행수": 0, "열수": 0, "상태": f"실패: {e}"})
    if not frames:
        sys.exit("읽을 수 있는 파일이 없습니다.")
    return pd.concat(frames, ignore_index=True, sort=False), log


def make_pivot(df: pd.DataFrame, group: str, column: str | None, value: str | None, agg: str) -> pd.DataFrame:
    for c in filter(None, [group, column, value]):
        if c not in df.columns:
            sys.exit(f"'{c}' 열이 없습니다. 사용 가능한 열: {', '.join(df.columns)}")
    if agg == "count" or value is None:
        data, agg = df.assign(_n=1), "count"
        value = "_n"
    else:
        data = df.copy()
        data[value] = pd.to_numeric(data[value].astype(str).str.replace(r"[,\s₩원]", "", regex=True), errors="coerce")
    if column:
        pv = pd.pivot_table(data, index=group, columns=column, values=value, aggfunc=agg,
                            margins=True, margins_name="전체", observed=True)
    else:
        pv = data.groupby(group, dropna=False)[value].agg(agg).to_frame(
            f"{value if value != '_n' else ''} {AGG_LABEL[agg]}".strip())
        total = data[value].agg(agg)
        pv = pd.concat([pv.sort_values(pv.columns[0], ascending=False),
                        pd.DataFrame({pv.columns[0]: [total]}, index=["전체"])])
    pv.index.name = group
    return pv.reset_index()


def style_sheet(ws, money_cols: set[int] | None = None, total_row: bool = False) -> None:
    for cell in ws[1]:
        cell.fill, cell.font = HEADER_FILL, HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions
    for col in ws.iter_cols(min_row=1, max_row=min(ws.max_row, 500)):
        width = max(len(str(c.value or "")) + sum(ord(ch) > 255 for ch in str(c.value or "")) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(width + 2, 8), 45)
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.border = Border(bottom=THIN)
            if isinstance(c.value, (int, float)):
                c.number_format = "#,##0.##" if isinstance(c.value, float) else "#,##0"
            elif isinstance(c.value, (datetime, date)):
                c.number_format = "yyyy-mm-dd"
    if total_row and ws.max_row > 1:
        for c in ws[ws.max_row]:
            c.fill, c.font = TOTAL_FILL, Font(bold=True)


def build_report(folder: Path, out: Path, group: str | None, column: str | None, value: str | None, agg: str) -> None:
    df, log = collect(folder, out)
    dups = df[df.drop(columns="원본파일").duplicated(keep=False)]

    summary = pd.DataFrame(log)
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        info = pd.DataFrame({
            "항목": ["생성일시", "입력 폴더", "읽은 시트 수", "전체 행 수", "중복 의심 행", "집계 기준"],
            "값": [datetime.now().strftime("%Y-%m-%d %H:%M"), str(folder.resolve()), len(log), len(df), len(dups),
                  f"{group} × {column or '-'} / {value or '행'} {AGG_LABEL[agg]}" if group else "(지정 안 함)"],
        })
        info.to_excel(xw, sheet_name="요약", index=False)
        summary.to_excel(xw, sheet_name="요약", index=False, startrow=len(info) + 2)
        df.to_excel(xw, sheet_name="합친데이터", index=False)
        if group:
            make_pivot(df, group, column, value, agg).to_excel(xw, sheet_name="집계", index=False)
        if len(dups):
            dups.to_excel(xw, sheet_name="중복행", index=False)

    wb = load_workbook(out)
    for ws in wb.worksheets:
        style_sheet(ws, total_row=ws.title == "집계")
    # 요약 시트의 두 번째 표 머리글도 꾸미기
    ws = wb["요약"]
    for cell in ws[len(info) + 3]:
        cell.fill, cell.font = HEADER_FILL, HEADER_FONT
    ws.auto_filter.ref = None
    wb.save(out)

    print(f"완료: {out}")
    print(f"  시트 {len(log)}개, 총 {len(df):,}행 합침 / 중복 의심 {len(dups):,}행")
    for row in log:
        print(f"  - {row['원본파일']}: {row['행수']:,}행 ({row['상태']})")


def make_sample(folder: Path) -> None:
    """월별 판매 파일 3개를 만든다 (연습용 가상 데이터)."""
    import random
    random.seed(7)
    folder.mkdir(parents=True, exist_ok=True)
    items = {"노트북": 1290000, "모니터": 289000, "키보드": 59000, "마우스": 29000}
    for month in (1, 2, 3):
        rows = []
        for _ in range(60):
            it = random.choice(list(items))
            q = random.randint(1, 10)
            rows.append({"일자": date(2026, month, random.randint(1, 28)), "지점": random.choice(["서울", "부산", "대구", "광주"]),
                         "품목": it, "수량": q, "단가": items[it], "매출액": q * items[it]})
        pd.DataFrame(rows).sort_values("일자").to_excel(folder / f"판매_2026-{month:02d}.xlsx", index=False)
    print(f"샘플 파일 3개를 만들었습니다: {folder}")


def main() -> None:
    ap = argparse.ArgumentParser(description="폴더 안 엑셀 파일을 합치고 집계 보고서를 만듭니다.")
    ap.add_argument("folder", type=Path, help="엑셀/CSV 파일이 들어 있는 폴더")
    ap.add_argument("-o", "--out", type=Path, help="결과 파일 경로 (기본: 폴더/보고서_YYYYMMDD.xlsx)")
    ap.add_argument("-g", "--group", help="행 기준 열 이름 (예: 지점)")
    ap.add_argument("-c", "--column", help="열 기준 열 이름 (선택, 예: 품목)")
    ap.add_argument("-v", "--value", help="값 열 이름 (예: 매출액). 없으면 개수를 셉니다")
    ap.add_argument("-a", "--agg", choices=list(AGG_LABEL), default="sum", help="계산 방법 (기본 sum)")
    ap.add_argument("--sample", action="store_true", help="폴더에 연습용 샘플 파일을 만들고 끝냅니다")
    args = ap.parse_args()

    if args.sample:
        make_sample(args.folder)
        return
    if not args.folder.is_dir():
        sys.exit(f"폴더를 찾을 수 없습니다: {args.folder}")
    out = args.out or args.folder / f"보고서_{date.today():%Y%m%d}.xlsx"
    agg = "count" if not args.value else args.agg
    build_report(args.folder, out, args.group, args.column, args.value, agg)


if __name__ == "__main__":
    main()
