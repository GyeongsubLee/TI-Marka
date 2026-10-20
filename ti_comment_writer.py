"""EHT비교 시트에 적은 COMMENT 를 개별 TI-LST 파일에 옮겨 적는다.

사용법
  1. ti_lst_reader.py 결과 파일의 'EHT비교' 시트 COMMENT 열(AC)에 코멘트를 적고,
     파일명에 COMMENT 가 들어가게 저장한다 (예: EHT비교_COMMENT.xlsx, 그 시트만 따로 저장해도 됨).
  2. 그 파일을 TI-LST 파일들이 있는 폴더(ti_lst_reader.py 와 같은 폴더)에 넣고 이 프로그램을 실행한다.
     (COMMENT 파일을 이 프로그램 위로 끌어다 놓아도 됨)
  3. 'OUTPUT TI COMMENT' 폴더에 COMMENT 가 들어간 TI-LST 가 '원래이름_COMMENT.xlsx' 로 저장된다.
     원본 TI-LST 는 바꾸지 않는다. 어디에 적었는지는 같은 폴더의 COMMENT_반영결과_*.xlsx 에 남는다.

어디에 적나
  - EHT비교에서 TI 와 비교한 행: 비교한 TI 품목(같은 ISO 시트·DN 의 파이프·밸브)의 품목 행 O열
  - 붙일 TI 품목이 없는 행(TI에 ISO 없음 등): 그 ISO 시트가 있는 TI-LST, 없으면 같은 구역에서 ISO 번호가
    가장 가까운 TI-LST 의 데이터가 끝난 다음 행부터 O열에 'ISO NO / 라인 / DN: COMMENT' 로 차례로 적음
  원본 파일의 다른 내용(서식, 그림, 머리글/바닥글 등)은 그대로 두고 O열 칸만 넣는다.
"""
from pathlib import Path
import os
import posixpath
import re
import sys
import time
import traceback
import xml.etree.ElementTree as ET
import zipfile
from collections import defaultdict
from datetime import datetime
from multiprocessing import freeze_support

import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

import ti_lst_reader as ti  # 같은 폴더의 ti_lst_reader.py (TI-LST 읽기·비교 규칙을 그대로 씀)


# =========================
# 사용자 설정
# =========================
COMMENT_FILE_KEYWORD = "COMMENT"     # 파일명에 이 글자가 있는 엑셀(EHT비교 시트)에서 COMMENT 를 읽음
OUTPUT_FOLDER = "OUTPUT TI COMMENT"  # 결과 폴더 (이 프로그램이 있는 폴더 안)
OUTPUT_SUFFIX = "_COMMENT"           # 결과 파일명: 원래이름_COMMENT.xlsx
COMMENT_COLUMN = "O"                 # TI-LST 에 COMMENT 를 적을 열 (N열 Note 바로 오른쪽)
APPEND_FORMAT = "{iso} / {line} / DN {dn}: {comment}"  # 데이터 끝에 추가하는 COMMENT 형식


REPORT_KEY_COLUMNS = ("ISO NO", "DN", "TI LINE NO", "TI 파일")  # EHT비교 시트인지 확인하는 머리글
ISO_SERIAL_RE = re.compile(r"^(.*-ISO)-(\d+)(?:-\d{2})?$")     # ...-TKM-ISO-00007-01 -> (...-TKM-ISO, 7)

LOG_COLUMNS = [
    "결과", "COMMENT", "TI 파일", "TI 행", "품명", "Size (mm)", "Nominal Diameter", "두께 (mm)",
    "LINE NO", "ISO NO", "DN", "판정", "EHT 라인", "COMMENT 파일", "COMMENT 행", "저장 파일",
]


# =========================
# COMMENT 읽기
# =========================
def comment_text(value):
    if value is None:
        return ""
    return str(value).strip()


def read_comment_file(path):
    """EHT비교 시트(머리글로 찾음)에서 COMMENT 가 적힌 행을 읽는다. EHT비교 시트가 없으면 None.

    COMMENT 열: 머리글에 COMMENT 가 들어간 열, 없으면 'EHT 행' 다음 열(AC)."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        for ws in wb.worksheets:
            ws.reset_dimensions()
            rows = ws.iter_rows(values_only=True)
            header = [ti.clean_text(v) for v in next(rows, ())]
            if not all(col in header for col in REPORT_KEY_COLUMNS):
                continue
            comment_col = next((i for i, h in enumerate(header) if "comment" in h.lower()), None)
            if comment_col is None:
                comment_col = header.index("EHT 행") + 1 if "EHT 행" in header else len(header)
            columns = {h: i for i, h in enumerate(header) if h and i != comment_col}

            records = []
            for row_no, row in enumerate(rows, start=2):
                text = comment_text(row[comment_col] if comment_col < len(row) else None)
                if not text:
                    continue
                rec = {h: (row[i] if i < len(row) else None) for h, i in columns.items()}
                rec.update({"COMMENT": text, "COMMENT 파일": path.name, "COMMENT 행": row_no})
                records.append(rec)
            return records, ws.title, get_column_letter(comment_col + 1)
    finally:
        wb.close()
    return None


def split_names(value, known=()):
    """'a, b' -> ['a', 'b'] (통째로 아는 이름이면 그대로)"""
    text = ti.clean_text(value)
    if not text:
        return []
    if text in known:
        return [text]
    return [part.strip() for part in text.split(", ") if part.strip()]


# =========================
# COMMENT 를 적을 TI 품목 찾기 (ti_lst_reader 의 EHT 비교 규칙 그대로)
# =========================
def comment_iso(rec):
    iso = ti.normalize_iso(rec.get("ISO NO"))
    if iso and not ti.has_sheet_no(iso):  # EHT ISO 에 시트번호가 없으면 라인번호 끝 숫자로 보충 (비교 때와 같음)
        sheet_no = ti.get_sheet_no(rec.get("EHT 라인"))
        if sheet_no:
            iso = ti.normalize_iso(f"{iso}-{sheet_no}")
    return iso


def comment_targets(rec, groups, known_files):
    """EHT비교 한 행에서 비교했던 TI 품목들.

    EHT비교의 TI LINE NO / TI 보온구분 / TI 파일 은 비교에 쓴 품목에서 나온 값이므로 그 값으로 다시 고르고,
    비교 때처럼 제외(Excluded) 라인은 빼고(그것만 있으면 그대로), 전기히팅 표시 품목이 있으면 그 품목만 쓴다."""
    entries = groups.get((comment_iso(rec), ti.dn_number(rec.get("DN"))), [])
    files = set(split_names(rec.get("TI 파일"), known_files))
    lines = set(split_names(rec.get("TI LINE NO")))
    types = set(split_names(rec.get("TI 보온구분")))
    picked = [e for e in entries
              if e["file"] in files and e["line"] in lines and (not types or e["type"] in types)]
    active = [e for e in picked if not e["excluded"]]
    if active:
        picked = [e for e in active if e["heating"]] or active
    return picked


def choose_append_file(iso, file_isos):
    """붙일 품목이 없는 COMMENT 를 적을 TI 파일: 그 ISO 시트가 있는 파일,
    없으면 같은 구역(ISO 번호 앞부분이 같은) 파일 중 ISO 번호가 가장 가까운 파일."""
    exact = sorted(name for name, isos in file_isos.items() if iso in isos)
    if exact:
        return exact[0]
    m = ISO_SERIAL_RE.match(iso or "")
    if not m:
        return None
    base, serial = m.group(1), int(m.group(2))
    best = None
    for name in sorted(file_isos):
        distances = [abs(int(k.group(2)) - serial) for k in map(ISO_SERIAL_RE.match, file_isos[name])
                     if k and k.group(1) == base]
        if distances and (best is None or min(distances) < best[0]):
            best = (min(distances), name)
    return best[1] if best else None


# =========================
# TI-LST 에 적기 (시트 XML 에 칸만 추가 - 다른 내용은 그대로)
# =========================
def read_rels(zf, part):
    """part 의 관계 파일 -> {Id: (Type, 대상 경로)}"""
    folder, name = posixpath.split(part)
    root = ET.fromstring(zf.read(posixpath.join(folder, "_rels", name + ".rels")))
    rels = {}
    for rel in root:
        if rel.get("TargetMode") == "External":
            continue
        target = rel.get("Target", "")
        path = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(folder, target))
        rels[rel.get("Id")] = (rel.get("Type", ""), path)
    return rels


def find_sheet_part(zf, sheet_name):
    """시트 이름 -> 시트 XML 경로 (예: xl/worksheets/sheet2.xml)"""
    workbook = next(path for kind, path in read_rels(zf, "").values() if kind.endswith("/officeDocument"))
    sheet_rels = read_rels(zf, workbook)
    for sheet in ET.fromstring(zf.read(workbook)).iter():
        if sheet.tag.split("}")[-1] == "sheet" and sheet.get("name") == sheet_name:
            rel_id = next(v for k, v in sheet.attrib.items() if k.endswith("}id"))
            return sheet_rels[rel_id][1]
    raise ValueError(f"시트를 찾지 못했습니다: {sheet_name}")


def decode_xml(data):
    m = re.match(rb'\s*<\?xml[^>]*encoding=["\']([\w-]+)["\']', data)
    encoding = m.group(1).decode() if m else "utf-8"
    if encoding.lower() in ("utf-8", "utf8") and data.startswith(b"\xef\xbb\xbf"):
        encoding = "utf-8-sig"
    return data.decode(encoding), encoding


class SheetXml:
    """시트 XML 에서 sheetData 의 행 위치를 찾아 칸을 넣는다 (정규식으로 필요한 곳만 고침)."""

    def __init__(self, xml):
        self.xml = xml
        m = re.search(r"<(\w+:)?sheetData\b[^>]*?(/?)>", xml)
        if not m:
            raise ValueError("sheetData 가 없습니다")
        self.prefix = m.group(1) or ""
        p = re.escape(self.prefix)
        self.row_re = re.compile(rf"<{p}row\b([^>]*?)(/>|>(.*?)</{p}row>)", re.S)
        self.cell_re = re.compile(rf"<{p}c\b([^>]*?)(/>|>.*?</{p}c>)", re.S)
        self.value_re = re.compile(rf"<{p}(?:v|is|f)\b")
        if m.group(2):  # <sheetData/>
            self.head, self.body, self.tail = xml[:m.start()], "", xml[m.end():]
            self.open_tag, self.close_tag = f"<{self.prefix}sheetData>", f"</{self.prefix}sheetData>"
        else:
            end = xml.index(f"</{self.prefix}sheetData>", m.end())
            self.head, self.body, self.tail = xml[:m.end()], xml[m.end():end], xml[end:]
            self.open_tag = self.close_tag = ""

    def rows(self):
        """(행 번호, 매치) - r 이 없는 행은 앞 행 + 1"""
        current = 0
        for rm in self.row_re.finditer(self.body):
            r = re.search(r'\br\s*=\s*["\'](\d+)["\']', rm.group(1))
            current = int(r.group(1)) if r else current + 1
            yield current, rm

    def last_value_row(self):
        """값이 들어 있는 마지막 행 번호 (모든 열 기준)"""
        last = 0
        for row_no, rm in self.rows():
            if rm.group(3) and self.value_re.search(rm.group(3)):
                last = row_no
        return last

    def cell(self, ref, style, text):
        p = self.prefix
        text = ti.ILLEGAL_XML_CHARS_RE.sub("", text)[:ti.EXCEL_MAX_CELL_CHARS]
        return (f'<{p}c r="{ref}"{style} t="inlineStr"><{p}is>'
                f'<{p}t xml:space="preserve">{ti.xml_escape(text)}</{p}t></{p}is></{p}c>')

    def row_with_cell(self, attrs, body, row_no, col, text):
        """행 안의 col 칸을 text 로 (있으면 서식 유지하고 바꾸고, 없으면 열 순서에 맞게 끼워 넣음)"""
        current = 0
        replace = insert_at = None
        cells = list(self.cell_re.finditer(body))
        for cm in cells:
            ref = re.search(r'\br\s*=\s*["\']([A-Za-z]+)\d+["\']', cm.group(1))
            current = column_index_from_string(ref.group(1).upper()) if ref else current + 1
            if current == col:
                replace = cm
                break
            if current > col:
                insert_at = cm.start()
                break
        style = ""
        if replace:
            s = re.search(r'\bs\s*=\s*["\'](\d+)["\']', replace.group(1))
            style = f' s="{s.group(1)}"' if s else ""
        new_cell = self.cell(f"{get_column_letter(col)}{row_no}", style, text)
        if replace:
            body = body[:replace.start()] + new_cell + body[replace.end():]
        else:
            at = insert_at if insert_at is not None else (cells[-1].end() if cells else 0)
            body = body[:at] + new_cell + body[at:]

        # spans(선택 속성)가 새 칸을 포함하지 않으면 지움
        sm = re.search(r'\s+spans\s*=\s*(["\'])([^"\']*)\1', attrs)
        if sm:
            try:
                covered = any(int(a) <= col <= int(b) for a, b in (r.split(":") for r in sm.group(2).split()))
            except ValueError:
                covered = False
            if not covered:
                attrs = attrs[:sm.start()] + attrs[sm.end():]
        return f"<{self.prefix}row{attrs}>{body}</{self.prefix}row>"

    def set_cells(self, col, values):
        """values: {행 번호: 글자} 를 col 열에 넣는다. 없는 행은 순서에 맞게 새로 만든다."""
        targets = sorted(values)
        pieces, pos, ti_ = [], 0, 0
        for row_no, rm in self.rows():
            new_rows = []
            while ti_ < len(targets) and targets[ti_] < row_no:
                n = targets[ti_]
                new_rows.append(f'<{self.prefix}row r="{n}">{self.cell(f"{get_column_letter(col)}{n}", "", values[n])}'
                                f'</{self.prefix}row>')
                ti_ += 1
            if new_rows:
                pieces += [self.body[pos:rm.start()], *new_rows]
                pos = rm.start()
            if ti_ < len(targets) and targets[ti_] == row_no:
                pieces += [self.body[pos:rm.start()],
                           self.row_with_cell(rm.group(1), rm.group(3) or "", row_no, col, values[row_no])]
                pos = rm.end()
                ti_ += 1
        pieces.append(self.body[pos:])
        for n in targets[ti_:]:
            pieces.append(f'<{self.prefix}row r="{n}">{self.cell(f"{get_column_letter(col)}{n}", "", values[n])}'
                          f'</{self.prefix}row>')
        self.body = "".join(pieces)
        self.widen_dimension(col, max(targets) if targets else 0)

    def widen_dimension(self, col, row):
        p = re.escape(self.prefix)
        m = re.search(rf'<{p}dimension\b[^>]*?\bref\s*=\s*(["\'])([^"\']+)\1', self.head)
        if not m or ":" not in m.group(2):
            return
        start, end = m.group(2).split(":", 1)
        em = re.fullmatch(r"([A-Za-z]+)(\d+)", end)
        if not em:
            return
        new_end = (f"{get_column_letter(max(column_index_from_string(em.group(1).upper()), col))}"
                   f"{max(int(em.group(2)), row)}")
        self.head = self.head[:m.start(2)] + f"{start}:{new_end}" + self.head[m.end(2):]

    def text(self):
        return self.head + self.open_tag + self.body + self.close_tag + self.tail


def existing_values(path, sheet_name, col):
    """원본 TI-LST 의 COMMENT 열에 이미 있는 값 {행: 값}"""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb[sheet_name]
        ws.reset_dimensions()
        return {r: v for r, (v,) in enumerate(ws.iter_rows(min_col=col, max_col=col, values_only=True), start=1)
                if v not in (None, "")}
    finally:
        wb.close()


def write_ti_comments(src, sheet_name, out_path, row_texts, append_texts):
    """원본 TI-LST 를 복사하며 대상 시트에 COMMENT 칸만 넣는다.

    row_texts: {품목 행: [COMMENT...]} / append_texts: 데이터 끝 다음 행부터 차례로 적을 글자
    돌려주는 값: ({품목 행: 기존 값 있었는지}, [추가한 행 번호])"""
    col = column_index_from_string(COMMENT_COLUMN)
    old = existing_values(src, sheet_name, col) if row_texts else {}
    values, had_old = {}, {}
    for row, texts in row_texts.items():
        text = " / ".join(dict.fromkeys(texts))
        before = comment_text(old.get(row))
        had_old[row] = bool(before) and text not in before
        values[row] = f"{before} / {text}" if had_old[row] else (before or text)

    with zipfile.ZipFile(src) as zin:
        part = find_sheet_part(zin, sheet_name)
        xml, encoding = decode_xml(zin.read(part))
        sheet = SheetXml(xml)
        first = max(sheet.last_value_row(), max(values, default=0)) + 1  # 원본 데이터가 끝난 다음 행
        appended = list(range(first, first + len(append_texts)))
        values.update(zip(appended, append_texts))
        sheet.set_cells(col, values)
        new_data = sheet.text().encode(encoding)
        ET.fromstring(new_data)  # 고친 XML 이 깨지지 않았는지 확인 (깨졌으면 여기서 오류 -> 저장 안 함)

        # 다른 파일(서식, 그림, 공유 문자열 등)은 원본 그대로 복사하고 시트 XML 만 바꿈
        tmp_path = out_path.with_name(out_path.name + ".tmp")
        try:
            with zipfile.ZipFile(tmp_path, "w") as zout:
                for info in zin.infolist():
                    copy = zipfile.ZipInfo(info.filename, info.date_time)
                    copy.compress_type = info.compress_type
                    copy.external_attr = info.external_attr
                    zout.writestr(copy, new_data if info.filename == part else zin.read(info.filename))
            os.replace(tmp_path, out_path)
        finally:
            if tmp_path.exists():
                tmp_path.unlink()
    return had_old, appended


# =========================
# 실행
# =========================
def is_eht_list(path):
    """ti_lst_reader 처럼 파일명에 EHT 가 있고 EHT LIST 머리글이 있으면 TI-LST 가 아님"""
    if not ti.EHT_FILE_KEYWORD or ti.EHT_FILE_KEYWORD.lower() not in path.name.lower():
        return False
    try:
        return ti.read_eht_list(path) is not None
    except Exception:
        return False


def find_ti_files(base_dir, out_dir):
    """ti_lst_reader 와 같은 규칙으로 TI-LST 파일 찾기 (결과 폴더, COMMENT/EHT LIST 파일 제외)"""
    files = base_dir.rglob("*.xlsx") if ti.RECURSIVE else base_dir.glob("*.xlsx")
    return sorted(f for f in files
                  if not any(k.lower() in f.name.lower() for k in ti.SKIP_KEYWORDS) and out_dir not in f.parents
                  and not is_eht_list(f))


def parse_ti_files(paths):
    """TI-LST 읽기 (ti_lst_reader 와 같은 방식, 병렬) -> 품목 목록"""
    items = []
    for file_items, _, summary, _ in ti.process_files(paths):
        if summary["Status"] == "ERROR":
            continue
        items.extend(file_items)
    return items


def log_row(result, rec, **extra):
    row = {
        "결과": result, "COMMENT": rec["COMMENT"], "ISO NO": rec.get("ISO NO"), "DN": rec.get("DN"),
        "판정": rec.get("판정"), "EHT 라인": rec.get("EHT 라인"),
        "COMMENT 파일": rec["COMMENT 파일"], "COMMENT 행": rec["COMMENT 행"],
    }
    row.update(extra)
    return row


def main():
    print("=== COMMENT -> TI-LST ===")
    started = time.perf_counter()
    base_dir = Path(__file__).resolve().parent
    out_dir = base_dir / OUTPUT_FOLDER
    print(f"작업 폴더: {base_dir}")

    # 1) COMMENT 파일 읽기 (끌어다 놓은 파일, 없으면 폴더에서 파일명에 COMMENT 가 있는 엑셀)
    comment_files = [Path(a) for a in sys.argv[1:] if a.lower().endswith(".xlsx")] or sorted(
        f for f in base_dir.glob("*.xlsx")
        if COMMENT_FILE_KEYWORD.lower() in f.name.lower() and not f.name.startswith("~$"))
    comments = []
    for path in comment_files:
        found = read_comment_file(path)
        if found is None:
            print(f"!! EHT비교 시트를 찾지 못함: {path.name}")
            continue
        records, sheet, col = found
        print(f"COMMENT 파일: {path.name} ('{sheet}' 시트 {col}열, COMMENT {len(records)}건)")
        comments.extend(records)
    if not comments:
        print(f"COMMENT 가 없습니다. 파일명에 '{COMMENT_FILE_KEYWORD}' 가 들어간 EHT비교 파일을 이 폴더에 넣어 주세요.")
        input("엔터를 누르면 종료합니다...")
        return

    # 2) TI-LST 읽기: COMMENT 에 나온 파일 먼저, 붙일 품목이 없는 COMMENT 가 있으면 나머지도 (같은 구역 파일 찾기)
    out_dir.mkdir(exist_ok=True)
    ti_files = find_ti_files(base_dir, out_dir)
    by_name = {}
    for f in ti_files:
        by_name.setdefault(f.name, f)
    named = {name for rec in comments for name in split_names(rec.get("TI 파일"), by_name)}
    first = [by_name[name] for name in sorted(named) if name in by_name]
    print(f"TI-LST 읽는 중... (COMMENT 에 나온 파일 {len(first)}개)")
    items = parse_ti_files(first) if first else []

    groups, _ = ti.build_ti_groups(items)
    targets = [comment_targets(rec, groups, by_name) for rec in comments]
    if any(not picked and not split_names(rec.get("TI 파일")) for rec, picked in zip(comments, targets)):
        rest = [f for f in ti_files if f.name not in named]
        if rest:
            print(f"TI 파일이 없는 COMMENT(TI에 ISO 없음 등)가 있어 같은 구역 TI-LST 를 찾으려고 나머지도 읽습니다 "
                  f"({len(rest)}개)")
            items += parse_ti_files(rest)

    item_of = {(rec["Source File"], rec["Source Start Row"]): rec for rec in items}
    path_of = {rec["Source File"]: Path(rec["Source Path"]) for rec in items}
    sheet_of = {rec["Source File"]: rec["Sheet Name"] for rec in items}
    file_isos = defaultdict(set)
    for rec in items:
        if rec["ISO NO"]:
            file_isos[rec["Source File"]].add(ti.normalize_iso(rec["ISO NO"]))

    # 3) 파일별로 적을 내용 모으기
    row_texts = defaultdict(lambda: defaultdict(list))   # 파일 -> 품목 행 -> [COMMENT]
    row_comments = defaultdict(list)                      # (파일, 행) -> [COMMENT 행 레코드]
    appends = defaultdict(list)                           # 파일 -> [(글자, COMMENT 행 레코드)]
    log = []
    for rec, picked in zip(comments, targets):
        if picked:
            for e in picked:
                row_texts[e["file"]][e["row"]].append(rec["COMMENT"])
                row_comments[(e["file"], e["row"])].append(rec)
            continue
        # 붙일 품목이 없으면 데이터 끝에 추가: TI 파일이 적혀 있으면 그 파일, 없으면 같은 구역 파일
        names = split_names(rec.get("TI 파일"), by_name)
        if names:
            name = next((n for n in names if n in sheet_of), None)
            if name is None:
                log.append(log_row("TI 파일 없음 (폴더에 없거나 읽지 못함)", rec, **{"TI 파일": ", ".join(names)}))
                continue
        else:
            iso = comment_iso(rec)
            name = choose_append_file(iso, file_isos) if iso else None
            if name is None:
                log.append(log_row("TI 파일 없음 (같은 구역 TI-LST 를 찾지 못함)", rec))
                continue
        line = ti.clean_text(rec.get("EHT 라인")) or ti.clean_text(rec.get("TI LINE NO"))
        text = APPEND_FORMAT.format(iso=ti.clean_text(rec.get("ISO NO")), line=line,
                                    dn=ti.clean_text(rec.get("DN")), comment=rec["COMMENT"])
        appends[name].append((text, rec))

    # 4) TI-LST 저장 (원본은 그대로, OUTPUT 폴더에 '원래이름_COMMENT.xlsx')
    saved = 0
    for name in sorted(set(row_texts) | set(appends)):
        src = path_of[name]
        out_path = out_dir / f"{src.stem}{OUTPUT_SUFFIX}{src.suffix}"
        try:
            had_old, appended = write_ti_comments(src, sheet_of[name], out_path, row_texts.get(name, {}),
                                                  [text for text, _ in appends.get(name, [])])
        except Exception as e:
            print(f"!! 저장 실패: {name} ({e})")
            for row in row_texts.get(name, {}):
                for rec in row_comments[(name, row)]:
                    log.append(log_row(f"저장 실패: {e}", rec, **{"TI 파일": name, "TI 행": row}))
            for _, rec in appends.get(name, []):
                log.append(log_row(f"저장 실패: {e}", rec, **{"TI 파일": name}))
            continue
        saved += 1
        print(f"저장: {out_path.name} (품목 {len(row_texts.get(name, {}))}행, 데이터 끝에 추가 {len(appended)}행)")
        for row in sorted(row_texts.get(name, {})):
            item = item_of[(name, row)]
            result = "품목 행에 적음 (기존 값 뒤에 붙임)" if had_old[row] else "품목 행에 적음"
            for rec in row_comments[(name, row)]:
                log.append(log_row(result, rec, **{
                    "TI 파일": name, "TI 행": row, "품명": item["Item Name"], "Size (mm)": item["Size (mm)"],
                    "Nominal Diameter": item["Nominal Diameter"], "두께 (mm)": item["Insulation Thickness (mm)"],
                    "LINE NO": item["LINE NO"], "저장 파일": out_path.name,
                }))
        for row, (text, rec) in zip(appended, appends.get(name, [])):
            log.append(log_row("데이터 끝에 추가", rec, **{
                "TI 파일": name, "TI 행": row, "저장 파일": out_path.name}))

    # 5) 어디에 적었는지 기록
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = out_dir / f"COMMENT_반영결과_{timestamp}.xlsx"
    columns = ti.order_columns(log, LOG_COLUMNS)
    ti.write_xlsx(log_path, [("반영결과", columns, [[rec.get(c) for c in columns] for rec in log])])

    print("\n=== 작업 완료 ===")
    print(f"COMMENT {len(comments)}건 -> TI-LST {saved}개 저장 ({out_dir})")
    failed = sum(1 for rec in log if not rec["결과"].startswith(("품목 행", "데이터 끝")))
    if failed:
        print(f"!! 적지 못한 COMMENT {failed}건 - {log_path.name} 의 '결과' 열 확인")
    print(f"반영 결과: {log_path}")
    print(f"소요 시간: {time.perf_counter() - started:.1f}초")
    input("엔터를 누르면 종료합니다...")


if __name__ == "__main__":
    freeze_support()  # exe로 묶어 실행할 때 병렬 처리용
    try:
        main()
    except Exception as e:
        print("### 프로그램 치명적 오류 ###")
        print(e)
        print(traceback.format_exc())
        input("엔터를 누르면 종료합니다...")
