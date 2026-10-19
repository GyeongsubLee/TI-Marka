from pathlib import Path
import re
import traceback
from datetime import datetime
import warnings

# 경고 숨김 (데이터 처리에는 영향 없는 openpyxl 경고)
warnings.filterwarnings("ignore", message=r"Cannot parse header or footer.*")
warnings.filterwarnings("ignore", message=r"DrawingML support is incomplete.*")

import openpyxl
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# =========================
# 사용자 설정
# =========================
RECURSIVE = False   # True: 하위 폴더까지 검색 / False: 현재 폴더만 검색
SKIP_KEYWORDS = ["가공본", "통합정리", "~$"]  # 결과물/임시파일 제외
DATA_START_ROW = 5   # 실제 데이터 시작 추정 행
VALIDATION_CHECK_MAX_COL = 14  # A:N 기준 검토


# =========================
# Pipe OD -> Nominal Diameter (DN) 변환표
# =========================
OD_TO_DN_MAP = {
    10.29: "DN 6", 13.72: "DN 8", 17.14: "DN 10",
    18.00: "DN 15", 21.34: "DN 15",
    25.00: "DN 20", 26.67: "DN 20",
    32.00: "DN 25", 33.40: "DN 25",
    38.00: "DN 32", 42.16: "DN 32",
    45.00: "DN 40", 48.26: "DN 40",
    57.00: "DN 50", 60.32: "DN 50",
    73.02: "DN 65", 76.00: "DN 65",
    88.90: "DN 80", 89.00: "DN 80",
    101.60: "DN 90",
    108.00: "DN 100", 114.30: "DN 100",
    141.30: "DN 125",
    159.00: "DN 150", 168.28: "DN 150",
    219.00: "DN 200", 219.08: "DN 200",
    273.00: "DN 250",
    323.80: "DN 300", 325.00: "DN 300",
    355.60: "DN 350", 377.00: "DN 350",
    406.40: "DN 400", 426.00: "DN 400",
    457.20: "DN 450",
    508.00: "DN 500", 530.00: "DN 500",
    558.80: "DN 550",
    609.60: "DN 600", 630.00: "DN 600",
    660.40: "DN 650",
    711.20: "DN 700",
    762.00: "DN 750",
    813.00: "DN 800", 820.00: "DN 800",
    864.00: "DN 850",
    914.00: "DN 900",
    965.00: "DN 950",
    1016.00: "DN 1000", 1020.00: "DN 1000",
    1067.00: "DN 1050",
    1118.00: "DN 1100",
    1168.00: "DN 1150",
    1219.00: "DN 1200", 1220.00: "DN 1200",
    1321.00: "DN 1300",
    1420.00: "DN 1400", 1422.00: "DN 1400",
    1524.00: "DN 1500",
    1626.00: "DN 1600",
    1727.00: "DN 1700",
    1829.00: "DN 1800",
    1930.00: "DN 1900",
    2032.00: "DN 2000",
}


def clean_text(x):
    if x is None:
        return ""
    if isinstance(x, str):
        return " ".join(x.replace("\n", " ").replace("\r", " ").split())
    return str(x)


def get_target_sheet(workbook):
    exact_name = "Техномонтажная ведомость"
    if exact_name in workbook.sheetnames:
        return workbook[exact_name]

    for name in workbook.sheetnames:
        lname = name.lower()
        if "ведомость" in lname:
            return workbook[name]

    for name in workbook.sheetnames:
        lname = name.lower()
        if "assembly" in lname or "insulation" in lname:
            return workbook[name]

    return None


def build_expanded_matrix(ws):
    rows = ws.max_row
    cols = ws.max_column
    mat = [[ws.cell(r, c).value for c in range(1, cols + 1)] for r in range(1, rows + 1)]

    for mr in list(ws.merged_cells.ranges):
        min_col, min_row, max_col, max_row = mr.bounds
        value = ws.cell(min_row, min_col).value
        for rr in range(min_row, max_row + 1):
            for cc in range(min_col, max_col + 1):
                mat[rr - 1][cc - 1] = value

    return mat


def get_row_values(mat, row_no, max_cols=VALIDATION_CHECK_MAX_COL):
    row = [clean_text(v) for v in mat[row_no - 1][:max_cols]]
    if len(row) < max_cols:
        row.extend([""] * (max_cols - len(row)))
    return row


def is_pipeline_header(text):
    t = clean_text(text)
    return "Pipeline / Трубопровод -" in t


def is_footer_row(text):
    t = clean_text(text)
    return t.startswith("Sheet /") or t.startswith("Лист")


def is_meaningful_row(row_vals):
    return any(clean_text(v) for v in row_vals)


def parse_pipeline_header(text):
    t = clean_text(text)

    marker = "Pipeline / Трубопровод -"
    idx = t.find(marker)
    if idx >= 0:
        t = t[idx:]

    result = {
        "Line Header": t,
        "LINE NO": "",
        "ISO NO": "",
        "INDOOR/OUTDOOR": "",
        "Excluded": "N",
        "Exclusion Letter": "",
    }

    m = re.search(
        r"Pipeline\s*/\s*Трубопровод\s*-\s*(.+?)\s+([A-Z]+)\s*\((INDOOR|OUTDOOR)\)",
        t,
        flags=re.IGNORECASE,
    )
    if m:
        base_line_no = m.group(1).strip()
        symbol = m.group(2).strip().upper()
        inout = m.group(3).strip().upper()
        result["LINE NO"] = f"{base_line_no} ({symbol})"
        result["INDOOR/OUTDOOR"] = inout
    else:
        m2 = re.search(
            r"Pipeline\s*/\s*Трубопровод\s*-\s*(.+?)\s*\((GCC-[^)]+ISO-\d+)\)",
            t,
            flags=re.IGNORECASE,
        )
        if m2:
            result["LINE NO"] = m2.group(1).strip()

    m_iso = re.search(r"(GCC-[A-Z0-9-]*TKM-ISO-\d+)", t, flags=re.IGNORECASE)
    if m_iso:
        result["ISO NO"] = m_iso.group(1)

    if "excluded according to the letter" in t.lower() or "исключено согласно письму" in t.lower():
        result["Excluded"] = "Y"
        matches = re.findall(r"(GCC-[A-Z0-9-]+)", t, flags=re.IGNORECASE)
        candidates = [m for m in matches if "ISO" not in m.upper()]
        if candidates:
            result["Exclusion Letter"] = candidates[-1]

    return result


def parse_size_to_float(size_text):
    s = clean_text(size_text)
    if not s:
        return None

    s = s.replace(" ", "")

    if re.fullmatch(r"\d{1,3}(,\d{3})+", s):
        s = s.replace(",", "")
        try:
            return float(s)
        except Exception:
            return None

    s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None


def get_nominal_diameter(item_name, od_text, tolerance=0.3):
    name = clean_text(item_name).lower()
    if not name:
        return ""

    if not ("pipe" in name or "труба" in name):
        return ""

    od_val = parse_size_to_float(od_text)
    if od_val is None:
        return ""

    for ref_od, dn in OD_TO_DN_MAP.items():
        if abs(od_val - ref_od) <= tolerance:
            return dn

    return ""


def autosize_worksheet(ws, min_width=10, max_width=45):
    for col in range(1, ws.max_column + 1):
        letter = get_column_letter(col)
        max_len = 0
        for row in range(1, ws.max_row + 1):
            v = ws.cell(row, col).value
            if v is None:
                continue
            length = len(str(v).replace("\n", " "))
            if length > max_len:
                max_len = length
        ws.column_dimensions[letter].width = min(max(max_len * 0.95, min_width), max_width)


def style_header_row(ws, row_no=1, color="1F4E78"):
    for cell in ws[row_no]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=color)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def write_dataframe(ws, df):
    if df.empty:
        ws["A1"] = "데이터 없음"
        return

    for c_idx, col_name in enumerate(df.columns, start=1):
        ws.cell(1, c_idx, col_name)

    style_header_row(ws, 1)

    for r_idx, (_, row) in enumerate(df.iterrows(), start=2):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(r_idx, c_idx, value)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    autosize_worksheet(ws)


def parse_insulation_workbook(file_path):
    wb = openpyxl.load_workbook(file_path, data_only=False)
    ws = get_target_sheet(wb)

    if ws is None:
        raise ValueError("대상 시트를 찾지 못했습니다. ('Техномонтажная ведомость' 등)")

    mat = build_expanded_matrix(ws)
    rows = ws.max_row
    cols = ws.max_column

    # ----- row 4가 Section인지, 바로 Pipeline Header인지 동적 판별 -----
    section_value = ""
    row4_a = clean_text(mat[3][0]) if rows >= 4 and cols >= 1 else ""

    if row4_a and not is_pipeline_header(row4_a):
        # 일반 케이스: row4 = Section
        section_value = row4_a
        start_row = 5
        covered_rows = set([1, 2, 3, 4])
    else:
        # 예외 케이스: row4 = 첫 Pipeline Header
        section_value = ""
        start_row = 4
        covered_rows = set([1, 2, 3])

    item_records = []
    excluded_headers = []
    parse_issues = []

    current_header = {
        "Line Header": "",
        "LINE NO": "",
        "ISO NO": "",
        "INDOOR/OUTDOOR": "",
        "Excluded": "",
        "Exclusion Letter": "",
    }

    r = start_row
    while r <= rows:
        row = get_row_values(mat, r)
        a, b, c, d, e, f, g, h, i, j, k, l, m, n = row

        if is_footer_row(a):
            covered_rows.add(r)
            r += 1
            continue

        if not is_meaningful_row(row):
            covered_rows.add(r)
            r += 1
            continue

        if is_pipeline_header(a):
            current_header = parse_pipeline_header(a)
            covered_rows.add(r)

            if current_header["Excluded"] == "Y":
                excluded_headers.append({
                    "Source File": file_path.name,
                    "Source Path": str(file_path),
                    "Sheet Name": ws.title,
                    "Section": section_value,
                    **current_header,
                    "Source Start Row": r,
                    "Source End Row": r,
                })

            r += 1
            continue

        # 품목행
        if b and not b.startswith("Name/") and not b.startswith("Pipeline /"):
            rec = {
                "Source File": file_path.name,
                "Source Path": str(file_path),
                "Sheet Name": ws.title,
                "Section": section_value,

                "Line Header": current_header.get("Line Header", ""),
                "LINE NO": current_header.get("LINE NO", ""),
                "ISO NO": current_header.get("ISO NO", ""),
                "INDOOR/OUTDOOR": current_header.get("INDOOR/OUTDOOR", ""),
                "Excluded": current_header.get("Excluded", ""),
                "Exclusion Letter": current_header.get("Exclusion Letter", ""),

                "Brand/Position": a,
                "Item Name": b,
                "Size (mm)": c,
                "Nominal Diameter": get_nominal_diameter(b, c),
                "Length/Height (m)": d,
                "Qty": e,
                "Temperature (°C)": f,
                "Purpose/Location": g,
                "Insulation Material": h,
                "Insulation Thickness (mm)": i,
                "Covering Thickness (mm)": j,
                "Surface (m²)": k,
                "Volume of Insulation Layer (m³)": l,
                "Document Designation": m,
                "Note": n,

                "Insulation Detail": "",
                "Insulation Detail Thickness (mm)": "",
                "Covering Material": "",
                "Covering Surface (m²)": "",

                "Source Start Row": r,
                "Source End Row": r,
            }

            rr = r + 1
            while rr <= rows:
                next_row = get_row_values(mat, rr)
                na, nb, nc, nd, ne, nf, ng, nh, ni, nj, nk, nl, nm, nn = next_row

                if is_pipeline_header(na) or nb or is_footer_row(na):
                    break

                if not is_meaningful_row(next_row):
                    # 완전 공란 행이면 item 범위에 넣지 않고 넘어감
                    break

                if any(next_row):
                    low_h = nh.lower() if nh else ""

                    if nh and ("thickness" in low_h or "толщина" in low_h):
                        if not rec["Insulation Detail"]:
                            rec["Insulation Detail"] = nh
                        if ni and not rec["Insulation Detail Thickness (mm)"]:
                            rec["Insulation Detail Thickness (mm)"] = ni
                        if nl and not rec["Volume of Insulation Layer (m³)"]:
                            rec["Volume of Insulation Layer (m³)"] = nl
                    elif nh:
                        if (nj or nk) and not rec["Covering Material"]:
                            rec["Covering Material"] = nh
                            if nj and not rec["Covering Thickness (mm)"]:
                                rec["Covering Thickness (mm)"] = nj
                            if nk and not rec["Covering Surface (m²)"]:
                                rec["Covering Surface (m²)"] = nk
                        elif not rec["Insulation Detail"]:
                            rec["Insulation Detail"] = nh

                    if ni and not rec["Insulation Detail Thickness (mm)"] and rec["Insulation Detail"]:
                        rec["Insulation Detail Thickness (mm)"] = ni

                rec["Source End Row"] = rr
                rr += 1

            # Source Start/End 범위 전체를 covered 처리
            for rr_cov in range(rec["Source Start Row"], rec["Source End Row"] + 1):
                covered_rows.add(rr_cov)

            # 기본 검토: ITEM인데 LINE/ISO가 비어 있으면 이슈 기록
            if not rec["LINE NO"] or not rec["ISO NO"]:
                parse_issues.append({
                    "Source File": file_path.name,
                    "Source Path": str(file_path),
                    "Sheet Name": ws.title,
                    "Issue Type": "MISSING_LINE_OR_ISO_ON_ITEM",
                    "Message": "ITEM은 추출되었으나 LINE NO 또는 ISO NO가 비어 있습니다.",
                    "Row No": rec["Source Start Row"],
                    "Source Start Row": rec["Source Start Row"],
                    "Source End Row": rec["Source End Row"],
                    "LINE NO": rec["LINE NO"],
                    "ISO NO": rec["ISO NO"],
                    "Item Name": rec["Item Name"],
                    "Row Snapshot": " | ".join(get_row_values(mat, rec["Source Start Row"])),
                })

            item_records.append(rec)
            r = rr
            continue

        # 여기까지 왔는데 의미 있는 행이면 파서가 놓친 행
        parse_issues.append({
            "Source File": file_path.name,
            "Source Path": str(file_path),
            "Sheet Name": ws.title,
            "Issue Type": "UNPARSED_ROW_DURING_SCAN",
            "Message": "의미 있는 행이지만 ITEM/HEADER/FOOTER로 분류되지 않았습니다.",
            "Row No": r,
            "Source Start Row": r,
            "Source End Row": r,
            "LINE NO": current_header.get("LINE NO", ""),
            "ISO NO": current_header.get("ISO NO", ""),
            "Item Name": b,
            "Row Snapshot": " | ".join(row),
        })
        r += 1

    # 후검증: Source Start/End 범위를 기준으로 커버되지 않은 의미 있는 행 전부 찾기
    for chk_row in range(start_row, rows + 1):
        row_vals = get_row_values(mat, chk_row)
        a = row_vals[0]

        if chk_row in covered_rows:
            continue
        if not is_meaningful_row(row_vals):
            continue
        if is_footer_row(a):
            continue

        parse_issues.append({
            "Source File": file_path.name,
            "Source Path": str(file_path),
            "Sheet Name": ws.title,
            "Issue Type": "UNCOVERED_MEANINGFUL_ROW",
            "Message": "통합정리데이터의 Source Start/End 범위 어떤 곳에도 포함되지 않은 의미 있는 행입니다.",
            "Row No": chk_row,
            "Source Start Row": chk_row,
            "Source End Row": chk_row,
            "LINE NO": "",
            "ISO NO": "",
            "Item Name": row_vals[1] if len(row_vals) > 1 else "",
            "Row Snapshot": " | ".join(row_vals),
        })

    unique_lines = len(set([x["LINE NO"] for x in item_records if x["LINE NO"]]))
    unique_iso = len(set([x["ISO NO"] for x in item_records if x["ISO NO"]]))

    summary = {
        "Source File": file_path.name,
        "Source Path": str(file_path),
        "Sheet Name": ws.title,
        "Rows": rows,
        "Columns": cols,
        "Item Count": len(item_records),
        "Unique LINE NO": unique_lines,
        "Unique ISO NO": unique_iso,
        "Excluded Header Count": len(excluded_headers),
        "Coverage Issue Count": len(parse_issues),
        "Status": "OK",
        "Message": "",
    }

    return item_records, excluded_headers, summary, parse_issues


def save_output_excel(output_path, all_items, all_excluded, summaries, errors):
    wb = Workbook()
    wb.remove(wb.active)

    ws_items = wb.create_sheet("통합정리데이터")
    df_items = pd.DataFrame(all_items)
    if not df_items.empty:
        preferred_cols = [
            "Source File", "Section",
            "LINE NO", "ISO NO", "INDOOR/OUTDOOR", "Excluded", "Exclusion Letter",
            "Item Name", "Brand/Position",
            "Size (mm)", "Nominal Diameter", "Length/Height (m)", "Qty", "Temperature (°C)",
            "Purpose/Location",
            "Insulation Material", "Insulation Thickness (mm)",
            "Insulation Detail", "Insulation Detail Thickness (mm)",
            "Covering Material", "Covering Thickness (mm)", "Covering Surface (m²)",
            "Surface (m²)", "Volume of Insulation Layer (m³)",
            "Document Designation", "Note",
            "Line Header",
            "Sheet Name", "Source Path",
            "Source Start Row", "Source End Row",
        ]
        existing_cols = [c for c in preferred_cols if c in df_items.columns]
        other_cols = [c for c in df_items.columns if c not in existing_cols]
        df_items = df_items[existing_cols + other_cols]
    write_dataframe(ws_items, df_items)

    ws_sum = wb.create_sheet("파일별요약")
    df_sum = pd.DataFrame(summaries)
    write_dataframe(ws_sum, df_sum)

    ws_exc = wb.create_sheet("제외라인목록")
    df_exc = pd.DataFrame(all_excluded)
    if not df_exc.empty:
        preferred_cols = [
            "Source File", "Section", "LINE NO", "ISO NO", "INDOOR/OUTDOOR",
            "Excluded", "Exclusion Letter", "Line Header", "Sheet Name", "Source Path",
            "Source Start Row", "Source End Row",
        ]
        existing_cols = [c for c in preferred_cols if c in df_exc.columns]
        other_cols = [c for c in df_exc.columns if c not in existing_cols]
        df_exc = df_exc[existing_cols + other_cols]
    write_dataframe(ws_exc, df_exc)

    ws_err = wb.create_sheet("에러로그")
    df_err = pd.DataFrame(errors)
    if not df_err.empty:
        preferred_cols = [
            "Source File", "Source Path", "Sheet Name", "Issue Type", "Message",
            "Row No", "Source Start Row", "Source End Row",
            "LINE NO", "ISO NO", "Item Name", "Row Snapshot",
            "Error", "Traceback",
        ]
        existing_cols = [c for c in preferred_cols if c in df_err.columns]
        other_cols = [c for c in df_err.columns if c not in existing_cols]
        df_err = df_err[existing_cols + other_cols]
    write_dataframe(ws_err, df_err)

    wb.save(output_path)


def main():
    print("=== START PROGRAM ===")

    base_dir = Path(__file__).resolve().parent
    print(f"작업 폴더: {base_dir}")

    files = list(base_dir.rglob("*.xlsx")) if RECURSIVE else list(base_dir.glob("*.xlsx"))

    target_files = []
    for f in files:
        name = f.name
        if any(keyword in name for keyword in SKIP_KEYWORDS):
            continue
        target_files.append(f)

    print(f"검색된 xlsx 파일 수: {len(files)}")
    print(f"실제 처리 대상 파일 수: {len(target_files)}")

    if not target_files:
        print("처리할 .xlsx 파일이 없습니다.")
        input("엔터를 누르면 종료합니다...")
        return

    all_items = []
    all_excluded = []
    summaries = []
    errors = []

    for idx, file_path in enumerate(sorted(target_files), start=1):
        print(f"[{idx}/{len(target_files)}] 처리 중: {file_path.name}")

        try:
            items, excluded, summary, parse_issues = parse_insulation_workbook(file_path)

            all_items.extend(items)
            all_excluded.extend(excluded)
            summaries.append(summary)
            errors.extend(parse_issues)

            print(f"    -> ITEM {len(items)}건, 제외라인 {len(excluded)}건, 검토이슈 {len(parse_issues)}건")

        except Exception as e:
            print(f"    !! 오류 발생: {file_path.name}")
            print(f"       {e}")

            summaries.append({
                "Source File": file_path.name,
                "Source Path": str(file_path),
                "Sheet Name": "",
                "Rows": "",
                "Columns": "",
                "Item Count": 0,
                "Unique LINE NO": 0,
                "Unique ISO NO": 0,
                "Excluded Header Count": 0,
                "Coverage Issue Count": 0,
                "Status": "ERROR",
                "Message": str(e),
            })

            errors.append({
                "Source File": file_path.name,
                "Source Path": str(file_path),
                "Sheet Name": "",
                "Issue Type": "FILE_PROCESSING_ERROR",
                "Message": "파일 처리 중 예외 발생",
                "Row No": "",
                "Source Start Row": "",
                "Source End Row": "",
                "LINE NO": "",
                "ISO NO": "",
                "Item Name": "",
                "Row Snapshot": "",
                "Error": str(e),
                "Traceback": traceback.format_exc(),
            })

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = base_dir / f"INSULATION_통합정리_{timestamp}.xlsx"

    print("결과 파일 저장 중...")
    save_output_excel(
        output_path=output_path,
        all_items=all_items,
        all_excluded=all_excluded,
        summaries=summaries,
        errors=errors,
    )

    print("\n=== 작업 완료 ===")
    print(f"통합 파일: {output_path}")
    print(f"총 아이템 수: {len(all_items)}")
    print(f"제외라인 수: {len(all_excluded)}")
    print(f"에러/검토이슈 수: {len(errors)}")

    input("엔터를 누르면 종료합니다...")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("### 프로그램 치명적 오류 ###")
        print(e)
        print(traceback.format_exc())
        input("엔터를 누르면 종료합니다...")
