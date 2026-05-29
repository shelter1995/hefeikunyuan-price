from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from openpyxl.styles import PatternFill

from .xlsx_utils import load_workbook_safe


@dataclass
class InventoryItem:
    product: str  # 螺纹, 盘螺, 线材/高线
    spec: str  # 6, 8, 10, 12, 14, etc.
    length: str | None  # 9, 12
    material: str | None  # HRB400E, HRB500E, HPB300
    status: str  # 充足, 告警, 缺货
    note: str = ""


# Color fills
FILL_BLUE = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
FILL_YELLOW = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
FILL_RED = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
CLEAR_FILL = PatternFill(fill_type=None)

CONFIRMED_MAPPING_STATUSES = {"已确认匹配", "已确认不更新"}


SIMPLE_SPEC_RE = re.compile(
    r"(?P<spec>\d+)[eE]?\s*(?:\((?P<note>[^)]+)\))?"
)

STATUS_ALERT_KEYWORDS = ("极少", "少", "少量", "紧张")
STATUS_SHORTAGE_KEYWORDS = ("无货", "缺货", "没货", "暂无", "等生产", "停产")


def _detect_status(spec_text: str) -> tuple[str, str]:
    text = spec_text.strip()
    # Check for shortage keywords
    for kw in STATUS_SHORTAGE_KEYWORDS:
        if kw in text:
            return "缺货", kw
    # Check for alert keywords
    for kw in STATUS_ALERT_KEYWORDS:
        if kw in text:
            return "告警", kw
    # Check for quantity note like (3件), (22件), (36件)
    if re.search(r"\(\d+件?\)", text):
        note = re.search(r"\(([^)]+)\)", text).group(1)
        return "告警", note
    # Check for numeric quantity without parentheses
    if re.search(r"\d+件", text):
        return "告警", re.search(r"(\d+件)", text).group(1)
    return "充足", ""


def _extract_specs(specs_text: str) -> list[tuple[str, str, str]]:
    """Extract (spec, status, note) tuples from specs text."""
    results: list[tuple[str, str, str]] = []
    # Split by common separators
    parts = re.split(r"[、，,；;]", specs_text)
    for part in parts:
        part = part.strip()
        if not part:
            continue
        m = SIMPLE_SPEC_RE.match(part)
        if m:
            spec = m.group("spec")
            status, note = _detect_status(part)
            results.append((spec, status, note))
    return results


def _parse_inventory_line(line: str) -> list[InventoryItem]:
    """Parse a single line for inventory information."""
    items: list[InventoryItem] = []

    # Pattern 1: "...规格有：specs"
    # Examples: "9米HRB400E规格有：12、16、18"
    #           "铁标12米HRB400E规格有：12、14（少）、16"
    #           "12米HRB500E规格有：12（3件）、20、22"
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

        # Extract material from prefix
        material = None
        m_mat = re.search(r"(HRB\d+E?|HPB\d+)", prefix)
        if m_mat:
            material = m_mat.group(1)

        # Extract product from prefix
        product = "螺纹"
        for p in ("盘螺", "线材", "高线"):
            if p in prefix:
                product = p
                break

        for spec, status, note in _extract_specs(specs_text):
            items.append(
                InventoryItem(
                    product=product,
                    spec=spec,
                    length=length,
                    material=material,
                    status=status,
                    note=note,
                )
            )
        return items

    # Pattern 2: "MATERIAL产品大包：specs"
    # Examples: "HRB400E盘螺大包：12（22件）、6、8"
    #           "HPB300线材大包：10"
    m = re.search(r"(HRB\d+E?|HPB\d+)?\s*(盘螺|线材|高线)(?:大包)?[：:]\s*(.+)", line)
    if m:
        material = m.group(1)
        product = m.group(2)
        specs_text = m.group(3)
        for spec, status, note in _extract_specs(specs_text):
            items.append(
                InventoryItem(
                    product=product,
                    spec=spec,
                    length=None,
                    material=material,
                    status=status,
                    note=note,
                )
            )
        return items

    # Pattern 3: "spec规格明天生产..." or "spec规格等生产"
    m = re.search(r"(\d+)(?:规格)?.*?等生产|(\d+)(?:规格)?.*?明天生产", line)
    if m:
        spec = m.group(1) or m.group(2)
        items.append(
            InventoryItem(
                product="螺纹",
                spec=spec,
                length=None,
                material=None,
                status="缺货",
                note="等生产",
            )
        )
        return items

    return items


def parse_inventory_text(text: str) -> list[InventoryItem]:
    """Parse inventory description from offline quote text."""
    items: list[InventoryItem] = []
    lines = text.splitlines()

    for line in lines:
        line = line.strip()
        if not line:
            continue
        items.extend(_parse_inventory_line(line))

    return items


def _company_match(file_company: str, target_company: str) -> bool:
    """Check if file company matches target company (fuzzy match)."""
    if not file_company or not target_company:
        return False
    # Exact match
    if file_company == target_company:
        return True
    # One contains the other
    if target_company in file_company or file_company in target_company:
        return True
    # Normalize and compare first 2 chars (common practice for Chinese company names)
    norm_file = re.sub(r"[报价钢厂集团]+$", "", file_company)
    norm_target = re.sub(r"[报价钢厂集团]+$", "", target_company)
    if norm_file == norm_target:
        return True
    if norm_target in norm_file or norm_file in norm_target:
        return True
    return False


def load_inventory_from_sources(
    source_json_paths: list[Path], company_name: str
) -> list[InventoryItem]:
    """Load inventory from source files based on company name."""
    items: list[InventoryItem] = []
    for path in source_json_paths:
        data = _load_json_safe(path)
        if not data:
            continue
        meta = data.get("meta") or {}
        input_file = str(meta.get("input_file") or "").strip()
        file_company = _extract_company_from_filename(input_file or path.name)
        if not _company_match(file_company, company_name):
            continue
        structured_items = _inventory_items_from_vision_result(data)
        if structured_items:
            items.extend(structured_items)
            continue
        # Try to read raw text from original file (pass OCR JSON for image files)
        text = _read_source_text(input_file or str(path), ocr_json=data)
        if text:
            items.extend(parse_inventory_text(text))
    return items


def _load_json_safe(path: Path) -> dict[str, Any] | None:
    try:
        import json
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None


def _extract_company_from_filename(filename: str) -> str:
    """Extract company name from filename or JSON data."""
    stem = Path(filename).stem
    # Remove common prefixes
    stem = re.sub(r"^ocr价格提取[_\-]?", "", stem)
    stem = re.sub(r"^test[_\-]?", "", stem)
    stem = re.sub(r"[_\-]?minimax$", "", stem)
    stem = re.sub(r"[_\-]?text$", "", stem)
    # Remove date suffixes
    stem = re.sub(r"[\-_ ]?\d{4}[\-_.]\d{1,2}[\-_.]\d{1,2}$", "", stem)
    stem = re.sub(r"[\-_ ]?\d{1,2}[\-_.]\d{1,2}$", "", stem)
    stem = re.sub(r"[\-_ ]+$", "", stem)
    return stem.strip()


def _read_source_text(input_file: str, ocr_json: dict[str, Any] | None = None) -> str | None:
    path = Path(input_file)
    if not path.exists():
        return None
    suffix = path.suffix.lower()
    if suffix == ".txt":
        return path.read_text(encoding="utf-8", errors="ignore")
    if suffix in {".jpg", ".jpeg", ".png", ".pdf", ".webp"}:
        # For image files, try to extract inventory from MiniMax vision result
        if ocr_json:
            # Try _vision_result first (MiniMax format)
            vision_result = ocr_json.get("_vision_result", {})
            if vision_result and "库存情况" in vision_result:
                inventory = vision_result["库存情况"]
                if isinstance(inventory, list):
                    # Convert inventory list to text lines in format compatible with parser
                    # Format: "location长度product规格有：spec1、spec2（status）"
                    lines = []
                    for item in inventory:
                        if isinstance(item, dict):
                            spec = item.get("规格", "")
                            status = item.get("状态", "充足")
                            # Parse spec to extract components
                            # e.g., "蚌埠9米螺纹12E" -> location="蚌埠", length="9米", product="螺纹", spec="12E"
                            m = re.match(r"([^\d]*)(\d+)米(螺纹|盘螺|线材|高线)(.+)", spec)
                            if m:
                                location = m.group(1)
                                length = m.group(2)
                                product = m.group(3)
                                spec_detail = m.group(4)
                                # Build line in expected format
                                line = f"{length}米规格有：{spec_detail}"
                                if status != "充足":
                                    line += f"（{status}）"
                                lines.append(line)
                            else:
                                # Fallback: just append as-is
                                lines.append(spec)
                        elif isinstance(item, str):
                            lines.append(item)
                    if lines:
                        return "\n".join(lines)
                elif isinstance(inventory, str):
                    return inventory
            # Fallback: try to reconstruct from records
            records = ocr_json.get("records", [])
            if records:
                lines = []
                for record in records:
                    if isinstance(record, dict):
                        line = record.get("规格", "")
                        if record.get("状态"):
                            line += f"（{record['状态']}）"
                        if line:
                            lines.append(line)
                if lines:
                    return "\n".join(lines)
        return None
    return None


def _inventory_items_from_vision_result(data: dict[str, Any]) -> list[InventoryItem]:
    vision_result = data.get("_vision_result")
    if not isinstance(vision_result, dict):
        return []
    inventory = vision_result.get("库存情况")
    if not isinstance(inventory, list):
        return []

    items: list[InventoryItem] = []
    for raw_item in inventory:
        if not isinstance(raw_item, dict):
            continue
        parsed = _parse_structured_inventory_item(raw_item)
        if parsed:
            items.append(parsed)
    return items


def _parse_structured_inventory_item(item: dict[str, Any]) -> InventoryItem | None:
    spec_text = str(item.get("规格") or "").strip()
    if not spec_text:
        return None

    status = str(item.get("状态") or "").strip()
    note = str(item.get("原始描述") or "").strip()
    if status not in {"充足", "告警", "缺货"}:
        status, detected_note = _detect_status(note or spec_text)
        note = note or detected_note

    length = None
    m_len = re.search(r"(\d+)米", spec_text)
    if m_len:
        length = m_len.group(1)

    product = "螺纹"
    for candidate in ("盘螺", "线材", "高线", "圆钢", "螺纹"):
        if candidate in spec_text:
            product = candidate
            break
    if product == "高线":
        product = "线材"

    tail = spec_text
    if product in spec_text:
        tail = spec_text.split(product, 1)[1]
    m_spec = re.search(r"(\d+)[eE]?", tail)
    if not m_spec:
        numbers = re.findall(r"\d+", spec_text)
        if not numbers:
            return None
        spec = numbers[-1]
    else:
        spec = m_spec.group(1)

    return InventoryItem(
        product=product,
        spec=spec,
        length=length,
        material=None,
        status=status,
        note=note,
    )


def _norm_text(s: str) -> str:
    return re.sub(r"\s+", "", s).upper()


def _match_product(item_product: str, row_product: str) -> bool:
    """Match inventory product with row product type."""
    item_norm = _norm_text(item_product)
    row_norm = _norm_text(row_product)
    if "螺纹" in item_norm and "螺纹" in row_norm:
        return True
    if "盘螺" in item_norm and "盘螺" in row_norm:
        return True
    if ("线材" in item_norm or "高线" in item_norm) and ("线材" in row_norm or "高线" in row_norm):
        return True
    return item_norm == row_norm


def _match_material(item_material: str | None, row_material: str) -> bool:
    if not item_material:
        return True
    item_norm = _norm_text(item_material)
    row_norm = _norm_text(row_material)
    return item_norm == row_norm or item_norm in row_norm or row_norm in item_norm


def _match_length(item_length: str | None, row_length: str | None) -> bool:
    if not item_length:
        return True
    return item_length == (row_length or "")


def _match_spec(item_spec: str, row_spec: str) -> bool:
    return item_spec == str(row_spec).strip()


def _detect_product_type(row: int, ws: Any) -> str:
    """Detect product type from A column (等级)."""
    length = ws.cell(row=row, column=3).value
    spec = ws.cell(row=row, column=2).value
    val = ws.cell(row=row, column=1).value
    if val:
        text = str(val).upper()
        if "一级钢" in text:
            return "线材"  # 一级钢 usually corresponds to 高线/线材
    # Check D column material
    mat = ws.cell(row=row, column=4).value
    if mat:
        mat_text = str(mat).upper()
        if "HPB" in mat_text:
            return "线材"
        if "HRB" in mat_text and not length and str(spec or "").strip() in {"6", "8", "10", "12"}:
            return "盘螺"
    # Default based on row context: rows 9-11 are 一级钢 (line), 12+ are 抗震三级钢 (rebar)
    if row <= 11:
        return "线材"
    return "螺纹"


def apply_inventory_to_project(
    project_excel: Path,
    mill_inventories: dict[str, list[InventoryItem]],
    sheet_name: str = "报价表",
    mapping_json_path: Path | None = None,
    clear_existing_colors: bool = False,
) -> dict[str, Any]:
    """
    Apply inventory colors to the quote sheet.
    mill_inventories: {mill_name: [InventoryItem, ...]}
    """
    wb = load_workbook_safe(project_excel)
    if sheet_name not in wb.sheetnames:
        return {"status": "skipped", "reason": f"{sheet_name} not found"}

    ws = wb[sheet_name]

    # Build mill -> column mapping from row 1 and row 8
    mill_to_col: dict[str, int] = {}
    for row in (1, 8):
        for col in range(1, ws.max_column + 1):
            val = ws.cell(row=row, column=col).value
            if val and isinstance(val, str):
                mill = val.strip()
                if mill and mill not in (
                    "钢厂",
                    "预备发货厂家",
                    "合肥鲲源贸易有限公司钢材报价单",
                    "等级",
                    "规格(mm)",
                    "长度（米）",
                    "材质",
                    "数量",
                    "金额",
                    "提货网差",
                    "数量（吨）",
                    "网差",
                ):
                    mill_to_col[mill] = col

    # Build row mapping
    row_map: list[dict[str, Any]] = []
    for r in range(9, ws.max_row + 1):
        spec = ws.cell(row=r, column=2).value
        length = ws.cell(row=r, column=3).value
        material = ws.cell(row=r, column=4).value
        if spec is None:
            continue
        product = _detect_product_type(r, ws)
        row_map.append({
            "row": r,
            "spec": str(spec).strip(),
            "length": str(length).strip() if length else None,
            "material": str(material).strip() if material else None,
            "product": product,
        })

    confirmed_mapping = _load_confirmed_mill_mapping(mapping_json_path)
    source_to_sheet_mill = {
        source_mill: _resolve_sheet_mill(source_mill, confirmed_mapping)
        for source_mill in mill_inventories
    }

    cleared_count = 0
    if clear_existing_colors and row_map:
        target_cols: set[int] = set()
        for resolved_mill in source_to_sheet_mill.values():
            for mapped_mill, col in mill_to_col.items():
                if _mill_match(mapped_mill, resolved_mill):
                    target_cols.add(col)
                    break

        for row_info in row_map:
            row = row_info["row"]
            for col in target_cols:
                ws.cell(row=row, column=col).fill = CLEAR_FILL
                cleared_count += 1

    applied: list[dict[str, Any]] = []

    for source_mill, items in mill_inventories.items():
        resolved_mill = source_to_sheet_mill.get(source_mill, source_mill)
        # Find column for this mill
        target_col = None
        for mapped_mill, col in mill_to_col.items():
            if _mill_match(mapped_mill, resolved_mill):
                target_col = col
                break

        if target_col is None:
            continue

        for item in items:
            # Find matching row
            for row_info in row_map:
                if (
                    _match_spec(item.spec, row_info["spec"])
                    and _match_material(item.material, row_info["material"] or "")
                    and _match_length(item.length, row_info["length"])
                    and _match_product(item.product, row_info["product"])
                ):
                    cell = ws.cell(row=row_info["row"], column=target_col)
                    if item.status == "充足":
                        cell.fill = FILL_BLUE
                    elif item.status == "告警":
                        cell.fill = FILL_YELLOW
                    elif item.status == "缺货":
                        cell.fill = FILL_RED
                    applied.append({
                        "mill": source_mill,
                        "sheet_mill": resolved_mill,
                        "row": row_info["row"],
                        "col": target_col,
                        "cell": ws.cell(row=row_info["row"], column=target_col).coordinate,
                        "product": item.product,
                        "spec": item.spec,
                        "length": item.length,
                        "material": item.material,
                        "status": item.status,
                    })
                    break

    wb.save(project_excel)
    return {
        "status": "ok",
        "applied_count": len(applied),
        "cleared_count": cleared_count,
        "applied": applied,
    }


def _mill_match(mapped: str, target: str) -> bool:
    """Check if mapped mill name matches target mill name."""
    mapped_norm = _norm_text(mapped)
    target_norm = _norm_text(target)
    if mapped_norm == target_norm:
        return True
    if mapped_norm in target_norm or target_norm in mapped_norm:
        return True
    return False


def _load_confirmed_mill_mapping(mapping_json_path: Path | None) -> dict[str, str]:
    """Load confirmed source->sheet mill mapping from mapping json."""
    if not mapping_json_path:
        return {}

    data = _load_json_safe(mapping_json_path)
    if not isinstance(data, list):
        return {}

    mapping: dict[str, str] = {}
    for row in data:
        if not isinstance(row, dict):
            continue
        status = str(row.get("状态") or "").strip()
        if status not in CONFIRMED_MAPPING_STATUSES:
            continue

        source_mill = str(row.get("最新清单厂家Sheet") or "").strip()
        sheet_mill = str(row.get("项目文件Sheet") or "").strip()
        if not source_mill or not sheet_mill:
            continue

        key = _norm_text(source_mill)
        old = mapping.get(key)
        if old is None or _mill_match(old, sheet_mill):
            mapping[key] = sheet_mill

    return mapping


def _resolve_sheet_mill(source_mill: str, confirmed_mapping: dict[str, str]) -> str:
    return confirmed_mapping.get(_norm_text(source_mill), source_mill)
