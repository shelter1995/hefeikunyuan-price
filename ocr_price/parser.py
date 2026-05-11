from __future__ import annotations

import re
from dataclasses import dataclass
from html import unescape
from typing import Any


LOCATION_KEYS = ("地点", "市场", "地区", "区域", "城市")
REBAR_KEYS = ("螺纹", "螺纹钢", "hrb", "hrb400", "hrb400e")
COIL_KEYS = ("盘螺", "盘圆", "线材")

COMPANY_RE = re.compile(r"([\u4e00-\u9fa5A-Za-z0-9（）()·\-_]{2,}(公司|集团|钢铁|钢厂|贸易|金属))")
DATE_RE = re.compile(
    r"(?P<y>\d{4})?[年/\-.]?(?P<m>\d{1,2})[月/\-.](?P<d>\d{1,2})日?"
)
PRICE_RE = re.compile(r"(?<!\d)(\d{3,5})(?!\d)")
LINE_PRICE_SIMPLE_RE = re.compile(
    r"(?P<location>[\u4e00-\u9fa5]{2,})\D{0,4}(?P<rebar>\d{3,5})\D{0,4}(?P<coil>\d{3,5})"
)


@dataclass
class Cell:
    text: str
    x: float
    y: float


def _normalize_text(s: str) -> str:
    return re.sub(r"\s+", "", s).lower()


def _contains_any(text: str, keys: tuple[str, ...]) -> bool:
    t = _normalize_text(text)
    return any(k in t for k in keys)


def _bbox_center(bbox: list[tuple[float, float]] | None) -> tuple[float, float] | None:
    if not bbox:
        return None
    xs = [p[0] for p in bbox]
    ys = [p[1] for p in bbox]
    return (sum(xs) / len(xs), sum(ys) / len(ys))


def _build_cells(text_boxes: list[dict[str, Any]]) -> list[Cell]:
    cells: list[Cell] = []
    for item in text_boxes:
        text = (item.get("text") or "").strip()
        if not text:
            continue
        center = _bbox_center(item.get("bbox"))
        if not center:
            continue
        cells.append(Cell(text=text, x=center[0], y=center[1]))
    return cells


def _cluster_rows(cells: list[Cell], y_threshold: float = 14.0) -> list[list[Cell]]:
    if not cells:
        return []
    ordered = sorted(cells, key=lambda c: (c.y, c.x))
    rows: list[list[Cell]] = [[ordered[0]]]
    for cell in ordered[1:]:
        if abs(cell.y - rows[-1][-1].y) <= y_threshold:
            rows[-1].append(cell)
        else:
            rows.append([cell])
    for row in rows:
        row.sort(key=lambda c: c.x)
    return rows


def _header_row_index(rows: list[list[Cell]]) -> int:
    best_idx = -1
    best_score = -1
    for idx, row in enumerate(rows):
        score = 0
        for cell in row:
            txt = cell.text
            if _contains_any(txt, LOCATION_KEYS):
                score += 3
            if _contains_any(txt, REBAR_KEYS):
                score += 2
            if _contains_any(txt, COIL_KEYS):
                score += 2
        if len(row) >= 3:
            score += 1
        if score > best_score:
            best_score = score
            best_idx = idx
    if best_idx < 0:
        raise ValueError("Unable to find a table header row from OCR text boxes.")
    return best_idx


def _column_anchors(row: list[Cell]) -> list[float]:
    return [c.x for c in row]


def _row_to_columns(row: list[Cell], anchors: list[float]) -> dict[int, str]:
    mapped: dict[int, str] = {}
    for cell in row:
        idx = min(range(len(anchors)), key=lambda i: abs(anchors[i] - cell.x))
        if idx in mapped:
            mapped[idx] = f"{mapped[idx]} {cell.text}".strip()
        else:
            mapped[idx] = cell.text
    return mapped


def _detect_column_groups(header_cols: dict[int, str], anchor_count: int) -> list[tuple[int, int, int]]:
    groups = _detect_column_groups_strict(header_cols, anchor_count)
    if groups:
        return groups

    # Fallback: strict 3-column loop when header OCR quality is poor.
    return [(i, i + 1, i + 2) for i in range(0, anchor_count - 2, 3)]


def _detect_column_groups_strict(
    header_cols: dict[int, str], anchor_count: int
) -> list[tuple[int, int, int]]:
    groups: list[tuple[int, int, int]] = []
    i = 0
    while i + 2 < anchor_count:
        t0 = header_cols.get(i, "")
        t1 = header_cols.get(i + 1, "")
        t2 = header_cols.get(i + 2, "")
        if (
            _contains_any(t0, LOCATION_KEYS)
            and _contains_any(t1, REBAR_KEYS)
            and _contains_any(t2, COIL_KEYS)
        ):
            groups.append((i, i + 1, i + 2))
            i += 3
            continue
        i += 1
    return groups


def _parse_price(text: str) -> int | None:
    values = [int(m.group(1)) for m in PRICE_RE.finditer(text or "")]
    if not values:
        return None
    return values[0]


def _is_valid_price_value(value: int | None) -> bool:
    return value is not None and 2000 <= value <= 6000


def _is_location_like(text: str) -> bool:
    val = (text or "").strip()
    if not val:
        return False
    if PRICE_RE.search(val):
        return False
    if _contains_any(val, REBAR_KEYS + COIL_KEYS):
        return False
    if len(val) > 16:
        return False
    return bool(re.search(r"[\u4e00-\u9fa5A-Za-z]", val))


def _extract_triplet_records_from_tokens(
    token_rows: list[list[str]],
    target_location: str | None = None,
    row_offset: int = 0,
) -> list[dict[str, Any]]:
    target_norm = _normalize_text(target_location) if target_location else None
    records: list[dict[str, Any]] = []

    for local_row_idx, row in enumerate(token_rows):
        tokens = [str(x or "").strip() for x in row if str(x or "").strip()]
        if len(tokens) < 3:
            continue
        i = 0
        while i + 2 < len(tokens):
            location = tokens[i].strip("：:|/ ")
            rebar_raw = tokens[i + 1]
            coil_raw = tokens[i + 2]
            rebar_price = _parse_price(rebar_raw)
            coil_price = _parse_price(coil_raw)
            if (
                _is_location_like(location)
                and _is_valid_price_value(rebar_price)
                and _is_valid_price_value(coil_price)
            ):
                if target_norm and target_norm not in _normalize_text(location):
                    i += 1
                    continue
                records.append(
                    {
                        "region_title": None,
                        "location": location,
                        "rebar_price": rebar_price,
                        "coil_price": coil_price,
                        "rebar_raw": rebar_raw,
                        "coil_raw": coil_raw,
                        "group_index": None,
                        "source_row_index": row_offset + local_row_idx,
                        "header_row_index": None,
                    }
                )
                i += 3
                continue
            i += 1
    return records


def _extract_company_and_date(lines: list[str]) -> tuple[str | None, str | None]:
    company = None
    date_str = None

    for line in lines:
        if not company:
            m = COMPANY_RE.search(line)
            if m:
                company = m.group(1)
        if not date_str:
            m = DATE_RE.search(line)
            if m:
                y = m.group("y")
                mth = int(m.group("m"))
                day = int(m.group("d"))
                if y:
                    date_str = f"{int(y):04d}-{mth:02d}-{day:02d}"
                else:
                    date_str = f"{mth:02d}-{day:02d}"
        if company and date_str:
            break
    return company, date_str


def _group_title(
    rows: list[list[Cell]],
    header_idx: int,
    col_idx: int,
    anchors: list[float],
    header_text: str,
) -> str | None:
    left = anchors[col_idx - 1] if col_idx > 0 else anchors[col_idx] - 40
    right = anchors[col_idx + 1] if col_idx + 1 < len(anchors) else anchors[col_idx] + 40

    for r in range(header_idx - 1, -1, -1):
        for cell in rows[r]:
            if left <= cell.x <= right:
                if not _contains_any(cell.text, LOCATION_KEYS + REBAR_KEYS + COIL_KEYS):
                    if not PRICE_RE.search(cell.text):
                        return cell.text
    if header_text and not _contains_any(header_text, LOCATION_KEYS):
        return header_text
    return None


def parse_table_price_records(
    text_boxes: list[dict[str, Any]],
    target_location: str | None = None,
) -> dict[str, Any]:
    cells = _build_cells(text_boxes)
    rows = _cluster_rows(cells)
    if not rows:
        raise ValueError("No OCR text cells found with coordinates.")

    header_idx = _header_row_index(rows)
    header_row = rows[header_idx]
    anchors = _column_anchors(header_row)
    header_cols = _row_to_columns(header_row, anchors)
    groups = _detect_column_groups(header_cols, len(anchors))

    all_lines = [" ".join(c.text for c in row) for row in rows[: max(header_idx, 1) + 3]]
    company, quote_date = _extract_company_and_date(all_lines)

    records: list[dict[str, Any]] = []
    last_location: dict[int, str] = {}
    target_norm = _normalize_text(target_location) if target_location else None

    for row_idx in range(header_idx + 1, len(rows)):
        mapped = _row_to_columns(rows[row_idx], anchors)
        for g_idx, (loc_col, rebar_col, coil_col) in enumerate(groups):
            location = (mapped.get(loc_col) or "").strip()
            rebar_raw = (mapped.get(rebar_col) or "").strip()
            coil_raw = (mapped.get(coil_col) or "").strip()

            if not location and g_idx in last_location and (rebar_raw or coil_raw):
                location = last_location[g_idx]
            if location:
                last_location[g_idx] = location

            if not (location or rebar_raw or coil_raw):
                continue

            rebar_price = _parse_price(rebar_raw)
            coil_price = _parse_price(coil_raw)
            if rebar_price is None and coil_price is None:
                continue

            if target_norm and target_norm not in _normalize_text(location):
                continue

            header_text = header_cols.get(loc_col, "")
            records.append(
                {
                    "region_title": _group_title(rows, header_idx, loc_col, anchors, header_text),
                    "location": location,
                    "rebar_price": rebar_price,
                    "coil_price": coil_price,
                    "rebar_raw": rebar_raw,
                    "coil_raw": coil_raw,
                    "group_index": g_idx,
                    "source_row_index": row_idx,
                    "header_row_index": header_idx,
                }
            )

    if not records:
        fallback_records = _extract_triplet_records_from_tokens(
            token_rows=[[c.text for c in row] for row in rows[max(header_idx + 1, 0) :]],
            target_location=target_location,
            row_offset=max(header_idx + 1, 0),
        )
        if fallback_records:
            records = fallback_records

    return {
        "company": company,
        "quote_date": quote_date,
        "header_row_index": header_idx,
        "group_count": len(groups),
        "records": records,
    }


LINE_PRICE_RE = re.compile(
    r"(?P<location>[\u4e00-\u9fa5A-Za-z0-9]{2,})\D{0,8}(?:螺纹|螺纹钢)\D*(?P<rebar>\d{3,5}).*?(?:盘螺|盘圆)\D*(?P<coil>\d{3,5})"
)

# New patterns for diverse formats
LINE_PRICE_RE_V2 = re.compile(
    r"(?P<location>[\u4e00-\u9fa5A-Za-z0-9]{2,})[：:]\s*(?P<rebar>\d{3,5})(?:螺纹|螺纹钢).*?[、,，]\s*(?P<coil>\d{3,5})(?:盘螺|盘圆)"
)
LINE_PRICE_RE_V3 = re.compile(
    r"(?P<location>[\u4e00-\u9fa5A-Za-z0-9]{2,})\s+(?:螺纹|螺纹钢)\s*(?P<rebar>\d{3,5})[，,、]\s*(?:盘螺|盘圆)\s*(?P<coil>\d{3,5})"
)
LINE_PRICE_RE_V4 = re.compile(
    r"(?P<location>[\u4e00-\u9fa5A-Za-z0-9]{2,}).*?(?:螺纹|螺纹钢)\s*(?P<rebar>\d{3,5}).*?(?:盘螺|盘圆)\s*(?P<coil>\d{3,5})"
)


def _try_parse_price_line(line: str) -> dict[str, str] | None:
    """Try multiple patterns to parse a price line."""
    for pattern in (LINE_PRICE_RE, LINE_PRICE_RE_V2, LINE_PRICE_RE_V3, LINE_PRICE_RE_V4, LINE_PRICE_SIMPLE_RE):
        m = pattern.search(line)
        if m and m.group("rebar") and m.group("coil"):
            return {
                "location": m.group("location"),
                "rebar": m.group("rebar"),
                "coil": m.group("coil"),
            }
    return None


def parse_price_lines_from_text(raw_text: str, target_location: str | None = None) -> dict[str, Any]:
    lines = [line.strip() for line in raw_text.splitlines() if line.strip()]
    company, quote_date = _extract_company_and_date(lines)
    target_norm = _normalize_text(target_location) if target_location else None

    records: list[dict[str, Any]] = []
    for idx, line in enumerate(lines):
        parsed = _try_parse_price_line(line)
        if not parsed:
            continue
        location = parsed["location"]
        if target_norm and target_norm not in _normalize_text(location):
            continue
        records.append(
            {
                "region_title": None,
                "location": location,
                "rebar_price": int(parsed["rebar"]),
                "coil_price": int(parsed["coil"]),
                "rebar_raw": parsed["rebar"],
                "coil_raw": parsed["coil"],
                "group_index": None,
                "source_row_index": idx,
                "header_row_index": None,
            }
        )

    return {
        "company": company,
        "quote_date": quote_date,
        "header_row_index": None,
        "group_count": None,
        "records": records,
    }


def _parse_markdown_tables(md_text: str) -> list[list[list[str]]]:
    """
    Returns tables as list[table], where table is list[row], row is list[cell_text].
    """
    lines = [ln.rstrip() for ln in md_text.splitlines()]
    tables: list[list[list[str]]] = []
    current_rows: list[list[str]] = []

    def flush() -> None:
        nonlocal current_rows
        if current_rows:
            tables.append(current_rows)
            current_rows = []

    for line in lines:
        striped = line.strip()
        if not striped or "|" not in striped:
            flush()
            continue
        if re.fullmatch(r"[\|\-\:\s]+", striped):
            continue
        cells = [c.strip() for c in striped.strip("|").split("|")]
        current_rows.append(cells)
    flush()
    return tables


def parse_markdown_price_records(md_text: str, target_location: str | None = None) -> dict[str, Any]:
    tables = _parse_markdown_tables(md_text)
    lines = [line.strip() for line in md_text.splitlines() if line.strip()]
    company, quote_date = _extract_company_and_date(lines[:60])
    parsed = _parse_records_from_tables(tables, target_location=target_location)
    return {
        "company": company,
        "quote_date": quote_date,
        "header_row_index": parsed["header_row_index"],
        "group_count": parsed["group_count"],
        "records": parsed["records"],
    }


HTML_TABLE_RE = re.compile(r"<table[^>]*>.*?</table>", re.I | re.S)
HTML_TR_RE = re.compile(r"<tr[^>]*>(.*?)</tr>", re.I | re.S)
HTML_CELL_RE = re.compile(r"<t[dh]([^>]*)>(.*?)</t[dh]>", re.I | re.S)
HTML_COLSPAN_RE = re.compile(r'colspan\s*=\s*"?(?P<n>\d+)"?', re.I)
HTML_TAG_RE = re.compile(r"<[^>]+>", re.S)


def _strip_html_text(raw: str) -> str:
    text = HTML_TAG_RE.sub("", raw or "")
    text = text.replace("&nbsp;", " ")
    return unescape(text).strip()


def _parse_html_tables(html_text: str) -> list[list[list[str]]]:
    tables: list[list[list[str]]] = []
    for table_html in HTML_TABLE_RE.findall(html_text or ""):
        rows: list[list[str]] = []
        for tr in HTML_TR_RE.findall(table_html):
            row_cells: list[str] = []
            for attr, content in HTML_CELL_RE.findall(tr):
                txt = _strip_html_text(content)
                m = HTML_COLSPAN_RE.search(attr or "")
                span = int(m.group("n")) if m else 1
                span = max(span, 1)
                for _ in range(span):
                    row_cells.append(txt)
            if row_cells:
                rows.append(row_cells)
        if rows:
            tables.append(rows)
    return tables


def _parse_records_from_tables(
    tables: list[list[list[str]]],
    target_location: str | None = None,
) -> dict[str, Any]:
    target_norm = _normalize_text(target_location) if target_location else None
    records: list[dict[str, Any]] = []
    group_count = 0
    header_row_index = None

    for table in tables:
        if len(table) < 2:
            continue
        max_col = max(len(r) for r in table)
        padded = [r + [""] * (max_col - len(r)) for r in table]

        found_header = -1
        best_groups: list[tuple[int, int, int]] = []
        for idx, row in enumerate(padded[: min(len(padded), 12)]):
            header_cols = {i: v for i, v in enumerate(row)}
            groups = _detect_column_groups_strict(header_cols, len(row))
            score = len(groups)
            if score > len(best_groups):
                best_groups = groups
                found_header = idx
        if not best_groups or found_header < 0:
            fallback = _extract_triplet_records_from_tokens(
                token_rows=padded,
                target_location=target_location,
                row_offset=0,
            )
            if fallback:
                records.extend(fallback)
            continue

        table_before = len(records)
        group_count = max(group_count, len(best_groups))
        if header_row_index is None:
            header_row_index = found_header
        headers = padded[found_header]
        last_locations: dict[int, str] = {}

        for row_idx in range(found_header + 1, len(padded)):
            row = padded[row_idx]
            for g_idx, (loc_col, rebar_col, coil_col) in enumerate(best_groups):
                location = row[loc_col].strip()
                rebar_raw = row[rebar_col].strip()
                coil_raw = row[coil_col].strip()

                if not location and g_idx in last_locations and (rebar_raw or coil_raw):
                    location = last_locations[g_idx]
                if location:
                    last_locations[g_idx] = location

                if not (location or rebar_raw or coil_raw):
                    continue

                rebar_price = _parse_price(rebar_raw)
                coil_price = _parse_price(coil_raw)
                if rebar_price is None and coil_price is None:
                    continue

                if target_norm and target_norm not in _normalize_text(location):
                    continue

                records.append(
                    {
                        "region_title": None,
                        "location": location,
                        "rebar_price": rebar_price,
                        "coil_price": coil_price,
                        "rebar_raw": rebar_raw,
                        "coil_raw": coil_raw,
                        "group_index": g_idx,
                        "source_row_index": row_idx,
                        "header_row_index": found_header,
                        "group_header_location": headers[loc_col] if loc_col < len(headers) else None,
                        "group_header_rebar": headers[rebar_col] if rebar_col < len(headers) else None,
                        "group_header_coil": headers[coil_col] if coil_col < len(headers) else None,
                    }
                )

        if len(records) == table_before:
            fallback = _extract_triplet_records_from_tokens(
                token_rows=padded[found_header + 1 :],
                target_location=target_location,
                row_offset=found_header + 1,
            )
            if fallback:
                records.extend(fallback)

    return {
        "records": records,
        "group_count": group_count if group_count else None,
        "header_row_index": header_row_index,
    }


def parse_html_price_records(html_text: str, target_location: str | None = None) -> dict[str, Any]:
    tables = _parse_html_tables(html_text)
    lines = [_strip_html_text(line) for line in html_text.splitlines() if line.strip()]
    company, quote_date = _extract_company_and_date(lines[:80])
    parsed = _parse_records_from_tables(tables, target_location=target_location)
    return {
        "company": company,
        "quote_date": quote_date,
        "header_row_index": parsed["header_row_index"],
        "group_count": parsed["group_count"],
        "records": parsed["records"],
    }


# ============================================================================
# Inventory extraction from text documents
# ============================================================================

INVENTORY_STATUS_KEYWORDS = {
    "缺货": ("无货", "无", "缺货", "没货", "暂无", "等生产", "停产"),
    "告警": ("极少", "少", "少量", "紧张", "配"),
}


def _detect_inventory_status(spec_text: str) -> tuple[str, str]:
    """Detect inventory status from spec description text."""
    text = spec_text.strip()
    # Check for shortage keywords
    for kw in INVENTORY_STATUS_KEYWORDS["缺货"]:
        if kw in text:
            return "缺货", kw
    # Check for alert keywords
    for kw in INVENTORY_STATUS_KEYWORDS["告警"]:
        if kw in text:
            return "告警", kw
    # Check for quantity note like (3件), (22件), (36件)
    m = re.search(r"\((\d+件?)\)", text)
    if m:
        return "告警", m.group(1)
    # Check for numeric quantity without parentheses
    m = re.search(r"(\d+件)", text)
    if m:
        return "告警", m.group(1)
    return "充足", ""


def _extract_inventory_specs(specs_text: str) -> list[dict[str, str]]:
    """Extract inventory specs from text like '12、14（少）、16、18（3件）'."""
    results: list[dict[str, str]] = []
    # Split by common separators
    parts = re.split(r"[、，,；;]", specs_text)
    for part in parts:
        part = part.strip()
        if not part:
            continue
        # Match spec number with optional note: "12", "14（少）", "16(3件)", "22E配"
        m = re.match(r"(\d+[eE]?)\s*(?:\(([^)]+)\))?", part)
        if m:
            spec = m.group(1)
            note = m.group(2) or ""
            status, detected_note = _detect_inventory_status(part)
            results.append({
                "规格": spec,
                "状态": status,
                "原始描述": detected_note or note or part,
            })
    return results


def _parse_inventory_line(line: str) -> list[dict[str, str]]:
    """Parse a single line for inventory information."""
    items: list[dict[str, str]] = []

    # Pattern 1: "...规格有：specs"
    # Examples: "9米HRB400E规格有：12、16、18"
    #           "铁标12米HRB400E规格有：12、14（少）、16"
    if "规格有" in line or "规格有" in line:
        prefix, _, specs_text = line.partition("规格有")
        specs_text = specs_text.lstrip("：:").strip()
        if not specs_text:
            return items

        # Extract length from prefix
        length = None
        m_len = re.search(r"(\d+)米", prefix)
        if m_len:
            length = m_len.group(1)

        # Extract product type from prefix
        product = "螺纹"
        for p in ("盘螺", "线材", "高线"):
            if p in prefix:
                product = p
                break

        # Build spec prefix
        spec_prefix = ""
        if length:
            spec_prefix += f"{length}米"
        spec_prefix += product

        for spec_item in _extract_inventory_specs(specs_text):
            items.append({
                "规格": f"{spec_prefix}{spec_item['规格']}",
                "状态": spec_item["状态"],
                "原始描述": spec_item["原始描述"],
            })
        return items

    # Pattern 2: "PRODUCT specs" (simple list)
    # Examples: "螺纹 抗震9米：10无货、12、14、16"
    m = re.search(r"(螺纹|盘螺|线材|高线).*?[：:]\s*(.+)", line)
    if m:
        product = m.group(1)
        specs_text = m.group(2)
        for spec_item in _extract_inventory_specs(specs_text):
            items.append({
                "规格": f"{product}{spec_item['规格']}",
                "状态": spec_item["状态"],
                "原始描述": spec_item["原始描述"],
            })
        return items

    # Pattern 3: "spec规格明天生产..." or "spec规格等生产"
    m = re.search(r"(\d+)(?:规格)?.*?等生产|(\d+)(?:规格)?.*?明天生产", line)
    if m:
        spec = m.group(1) or m.group(2)
        items.append({
            "规格": f"螺纹{spec}",
            "状态": "缺货",
            "原始描述": "等生产",
        })
        return items

    return items


def parse_inventory_from_text(raw_text: str) -> list[dict[str, str]]:
    """Parse inventory description from offline quote text.
    
    Returns list of inventory items with format:
    [{"规格": "9米螺纹12", "状态": "充足/告警/缺货", "原始描述": "..."}, ...]
    """
    items: list[dict[str, str]] = []
    lines = raw_text.splitlines()

    for line in lines:
        line = line.strip()
        if not line:
            continue
        items.extend(_parse_inventory_line(line))

    return items
