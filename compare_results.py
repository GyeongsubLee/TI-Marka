"""기존 프로그램 결과와 새 프로그램 결과(INSULATION_통합정리_*.xlsx)가 같은지 비교한다.

사용법
  python compare_results.py 기존결과.xlsx 새결과.xlsx
  (인자 없이 실행하면 이 폴더의 INSULATION_통합정리_*.xlsx 중 가장 최근 2개를 비교: 먼저 만든 파일 = 기존 결과)

행 순서가 아니라 (Source File, Issue Type, Source Start Row) 기준으로 같은 항목끼리 비교하므로
중간에 한 건이 빠지거나 늘어도 그 건만 보고된다.
"""
from pathlib import Path
import sys

import openpyxl

KEY_COLUMNS = ["Source File", "Issue Type", "Source Start Row"]  # 시트에 있는 것만 사용
IGNORE_COLUMNS = {"Source Path", "Traceback"}  # 실행 위치/코드 줄번호에 따라 달라지는 값
MAX_SHOW = 10  # 항목별로 보여줄 최대 건수


def read_sheets(path):
    """시트별로 {키: {컬럼: 값}} 을 만든다."""
    wb = openpyxl.load_workbook(path, read_only=True)
    try:
        sheets = {}
        for ws in wb.worksheets:
            rows = [[None if v == "" else v for v in row] for row in ws.iter_rows(values_only=True)]
            records = {}
            if rows and rows[0][0] != "데이터 없음":
                header = rows[0]
                key_cols = [c for c in KEY_COLUMNS if c in header]
                for row in rows[1:]:
                    rec = dict(zip(header, row))
                    key = tuple(rec[c] for c in key_cols)
                    while key in records:  # 같은 키가 또 있으면 순번을 붙여 구분
                        key = key + ("#",)
                    records[key] = rec
            sheets[ws.title] = records
        return sheets
    finally:
        wb.close()


def describe(rec):
    label = rec.get("Item Name") or rec.get("Row Snapshot") or rec.get("Message") or ""
    return str(label)[:80]


def compare(old_path, new_path):
    print(f"기존 결과: {old_path}")
    print(f"새 결과  : {new_path}\n")
    old, new = read_sheets(old_path), read_sheets(new_path)
    same_all = True

    for title, old_recs in old.items():
        new_recs = new.get(title, {})
        only_old = [k for k in old_recs if k not in new_recs]
        only_new = [k for k in new_recs if k not in old_recs]
        changed = []
        for key in old_recs.keys() & new_recs.keys():
            a, b = old_recs[key], new_recs[key]
            diffs = [(c, a.get(c), b.get(c)) for c in dict.fromkeys(list(a) + list(b))
                     if c not in IGNORE_COLUMNS and a.get(c) != b.get(c)]
            if diffs:
                changed.append((key, diffs))

        print(f"[{title}] 기존 {len(old_recs)}건 / 새 {len(new_recs)}건")
        if not (only_old or only_new or changed):
            print("  동일")
            continue

        same_all = False
        for name, keys, recs in (("기존 결과에만 있음", only_old, old_recs), ("새 결과에만 있음", only_new, new_recs)):
            if keys:
                print(f"  {name}: {len(keys)}건")
                for key in keys[:MAX_SHOW]:
                    print(f"    {key} {describe(recs[key])}")
        if changed:
            print(f"  값이 다른 항목: {len(changed)}건")
            for key, diffs in sorted(changed, key=str)[:MAX_SHOW]:
                print(f"    {key} " + ", ".join(f"{c}: {x!r} -> {y!r}" for c, x, y in diffs))

    print("\n결과: " + ("두 파일 내용이 같습니다." if same_all else "차이가 있습니다. 위 내용을 확인하세요."))


def main():
    if len(sys.argv) == 3:
        old_path, new_path = Path(sys.argv[1]), Path(sys.argv[2])
    else:
        base_dir = Path(__file__).resolve().parent
        results = sorted(base_dir.glob("INSULATION_통합정리_*.xlsx"))
        if len(results) < 2:
            print("비교할 결과 파일(INSULATION_통합정리_*.xlsx)이 2개 이상 필요합니다.")
            return
        old_path, new_path = results[-2], results[-1]
    compare(old_path, new_path)


if __name__ == "__main__":
    main()
    if len(sys.argv) != 3:
        input("엔터를 누르면 종료합니다...")
