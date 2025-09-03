# build_data_dictionary_hardcoded.py
# N5 schema: Source | Parameter Name | Data Type | Content | Description

import json
import math
import re
from pathlib import Path
from typing import List, Dict, Tuple
import pandas as pd

# ───────── 하드코딩 경로 설정 ─────────
MANIFEST_PATH = Path(r"code_path_TESTER.json")     # ← 여기를 실제 경로로 바꾸세요
OUT_XLSX      = Path(r"data_dictionary.xlsx")

INCLUDE_VARIABLES = True
ENCODINGS         = ("utf-8", "utf-8-sig", "cp949", "euc-kr")
DESC_DELIM        = "!\n"  # Description에 줄 단위 구분자(LLM 투입 용이 & 엑셀 줄바꿈)

# 엑셀 열 너비(문자 단위). set_column과 일치시켜야 줄 수 추정 정확.
COL_WIDTHS = [20, 28, 24, 90, 44]
BASE_LINE_HEIGHT = 15  # 포인트 단위 대략 값(엑셀 기본 15pt 근사)
MAX_ROW_HEIGHT = 300   # 안전 캡


def make_source_label(top: str, sub: str) -> str:
    if top == "SharedHeader":
        if sub == "GlobalHeader":
            return "Global Header"
        if "LayerHeader" in sub:
            return "Layer Header"
    return f"{top}.{sub}"


def read_text(path: Path) -> str:
    for enc in ENCODINGS:
        try:
            return path.read_text(encoding=enc)
        except Exception:
            continue
    return path.read_bytes().decode("utf-8", errors="ignore")


def normalize_newlines(text: str) -> str:
    return "\n".join(ln.rstrip() for ln in text.splitlines())


DEFINE_RE = re.compile(r'^\s*#\s*define\s+([A-Za-z_]\w+)\s+(.+?)\s*$', re.M)
ENUM_START_RE   = re.compile(r'^\s*(?:typedef\s+)?enum(?:\s+class)?\s+([A-Za-z_]\w+)?\s*\{', re.M)
STRUCT_START_RE = re.compile(r'^\s*(?:typedef\s+)?struct\s+([A-Za-z_]\w+)?\s*\{', re.M)
VAR_LINE_RE = re.compile(r'^[^\n;]+;\s*$', re.M)


def looks_like_var_decl(line: str) -> bool:
    s = line.strip()
    if not s.endswith(';'): return False
    if s.startswith('#'): return False
    if s.startswith(("typedef ", "using ", "enum ", "struct ", "union ")): return False
    if "(" in s and ")" in s: return False
    return True


def split_type_name(decl: str) -> Tuple[str, str]:
    s = decl.strip().rstrip(';').strip()
    if "," in s: return "", ""
    toks = s.split()
    if not toks: return "", ""
    name = toks[-1]
    m = re.match(r'([A-Za-z_]\w*)(.*)', name)
    if m: name = m.group(1)
    cut = s.rfind(name)
    if cut == -1: return "", ""
    typ = s[:cut].strip()
    return typ, name


def find_brace_block(text: str, start_pos: int) -> Tuple[int, int]:
    depth = 0
    i = start_pos
    while i < len(text):
        ch = text[i]
        if ch == '{':
            depth += 1
        elif ch == '}':
            depth -= 1
            if depth == 0:
                j = i + 1
                while j < len(text) and text[j].isspace(): j += 1
                if j < len(text) and text[j] == ';': j += 1
                return start_pos, j
        i += 1
    return start_pos, start_pos


# Description 생성 로직

def make_description_from_content(content: str) -> str:
    lines = [ln.strip() for ln in content.splitlines() if ln.strip()]
    # define 은 한줄 그대로
    if content.strip().startswith("#define"):
        return content.strip()
    # enum/struct 내부만 추출 후 각 줄을 DESC_DELIM으로 구분
    if "{" in content and "}" in content:
        inner = content[content.find("{")+1:content.rfind("}")]
        inner_lines = [l.strip().rstrip(',') for l in inner.splitlines() if l.strip()]
        return DESC_DELIM.join(inner_lines)
    # 변수 선언 등은 라인 그대로 넣어도 됨(필요시 공란 유지 가능)
    return content.strip()


def parse_header_text(source_label: str, text: str) -> List[Dict]:
    rows: List[Dict] = []
    # 1) #define (함수형 매크로 제외)
    for m in DEFINE_RE.finditer(text):
        name = m.group(1)
        if "(" in name and ")" in name:
            continue
        body = m.group(0).strip()
        rows.append({
            "Source": source_label,
            "Parameter Name": name,
            "Data Type": "Macro",
            "Content": body,
            "Description": make_description_from_content(body)
        })
    # 2) enum/struct 블록
    for kind_re in (ENUM_START_RE, STRUCT_START_RE):
        start = 0
        while True:
            m = kind_re.search(text, pos=start)
            if not m: break
            brace_open = text.find("{", m.start())
            if brace_open == -1: break
            b0, b1 = find_brace_block(text, brace_open)
            header_start = m.start()
            content = text[header_start:b1].strip()
            rows.append({
                "Source": source_label,
                "Parameter Name": "-",
                "Data Type": "-",
                "Content": content,
                "Description": make_description_from_content(content)
            })
            start = b1
    # 3) 변수 선언(옵션)
    if INCLUDE_VARIABLES:
        for vm in VAR_LINE_RE.finditer(text):
            line = vm.group(0)
            if not looks_like_var_decl(line): continue
            typ, name = split_type_name(line)
            if not typ or not name: continue
            rows.append({
                "Source": source_label,
                "Parameter Name": name,
                "Data Type": typ,
                "Content": line.strip(),
                "Description": ""  # 변수는 설명 비워둠
            })
    return rows


# ===== 동적 행 높이 계산 =====

def count_wrapped_lines(text: str, width_chars: int) -> int:
    if not text:
        return 1
    # 줄바꿈 기준으로 쪼개서 각 줄의 예상 랩 줄 수 합산
    total = 0
    for para in str(text).split("\n"):
        if not para:
            total += 1
            continue
        # 공백 기반 대략 길이
        est = max(1, math.ceil(len(para) / max(1, width_chars)))
        total += est
    return total


def compute_row_height(row: Dict) -> int:
    # 각 컬럼별 텍스트와 그 열 너비로 예상 줄 수 계산, 최댓값을 가져와 높이 환산
    texts = [
        row.get("Source", ""),
        row.get("Parameter Name", ""),
        row.get("Data Type", ""),
        row.get("Content", ""),
        row.get("Description", ""),
    ]
    lines_needed = 1
    for txt, w in zip(texts, COL_WIDTHS):
        lines_needed = max(lines_needed, count_wrapped_lines(txt, w))
    height = min(MAX_ROW_HEIGHT, max(BASE_LINE_HEIGHT, lines_needed * BASE_LINE_HEIGHT))
    return height


def build_from_manifest(manifest_path: Path, out_excel: Path) -> None:
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    all_rows: List[Dict] = []

    for top, submap in manifest.items():
        if not isinstance(submap, dict): continue
        for sub, files in submap.items():
            source_label = make_source_label(top, sub)
            for f in files:
                p = Path(f)
                if not p.exists():
                    all_rows.append({
                        "Source": source_label,
                        "Parameter Name": "-",
                        "Data Type": "-",
                        "Content": f"// FILE NOT FOUND: {p}",
                        "Description": ""
                    })
                    continue
                text = normalize_newlines(read_text(p))
                rows = parse_header_text(source_label, text)
                all_rows.extend(rows)

    # 중복 제거
    dedup, seen = [], set()
    for r in all_rows:
        key = (r["Source"], r["Parameter Name"], r["Data Type"], r["Content"])
        if key in seen: continue
        seen.add(key)
        dedup.append(r)

    df = pd.DataFrame(dedup, columns=["Source", "Parameter Name", "Data Type", "Content", "Description"])

    out_excel.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_excel, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="DataDictionary")
        wb, ws = writer.book, writer.sheets["DataDictionary"]
        wrap = wb.add_format({"text_wrap": True, "valign": "top"})
        for col, width in enumerate(COL_WIDTHS):
            ws.set_column(col, col, width, wrap)
        # 헤더 높이 고정
        ws.set_row(0, 22)
        # 각 행별로 내용 기반 높이 계산하여 설정
        for r_idx, row in enumerate(dedup, start=1):  # 1-based for Excel (0은 헤더)
            ws.set_row(r_idx, compute_row_height(row))

    print(f"OK: {out_excel}")


def main():
    build_from_manifest(MANIFEST_PATH, OUT_XLSX)


if __name__ == "__main__":
    main()
