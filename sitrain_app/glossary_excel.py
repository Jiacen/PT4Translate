from __future__ import annotations

from io import BytesIO
from pathlib import PurePosixPath
import zipfile

from fastapi import HTTPException
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill


HEADERS = ("英文术语", "中文标准译法", "英文别名（逗号分隔）", "中文别名（逗号分隔，可选）")
LEGACY_HEADERS = ("源术语", "强制译法", "源术语别名（逗号分隔）")
SHEET_NAME = "术语表"
MAX_IMPORT_BYTES = 5 * 1024 * 1024
MAX_IMPORT_ROWS = 5_000
MAX_ARCHIVE_ENTRIES = 500
MAX_ARCHIVE_UNCOMPRESSED = 20 * 1024 * 1024


def _excel_text(value: object) -> str:
    text = "" if value is None else str(value).strip()
    # Avoid turning a source/target term into an Excel formula on export.
    return f"'{text}" if text[:1] in {"=", "+", "-", "@"} else text


def _import_text(value: object) -> str:
    text = "" if value is None else str(value).strip()
    return text[1:] if len(text) > 1 and text.startswith("'") and text[1] in "=+-@" else text


def export_global_terms(terms: list[dict]) -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = SHEET_NAME
    sheet.append(HEADERS)
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="008A96")
        cell.alignment = Alignment(horizontal="center")
    for term in terms:
        sheet.append((_excel_text(term["source"]), _excel_text(term["target"]),
                      _excel_text(", ".join(term.get("aliases") or [])),
                      _excel_text(", ".join(term.get("chinese_aliases") or []))))
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:D{max(1, sheet.max_row)}"
    sheet.column_dimensions["A"].width = 32
    sheet.column_dimensions["B"].width = 32
    sheet.column_dimensions["C"].width = 42
    sheet.column_dimensions["D"].width = 42

    instructions = workbook.create_sheet("填写说明")
    instructions.append(("使用说明",))
    instructions["A1"].font = Font(bold=True, color="FFFFFF")
    instructions["A1"].fill = PatternFill("solid", fgColor="008A96")
    instructions.append(("请在“术语表”中填写；英文术语和中文标准译法为必填项。",))
    instructions.append(("英文别名用于识别同一英文术语的不同写法；中文别名用于中文→英文翻译时识别不同中文写法。多个别名以英文或中文逗号分隔。",))
    instructions.append(("导入时重复的英文术语不会覆盖现有全局术语，而会被跳过。",))
    instructions.column_dimensions["A"].width = 90

    output = BytesIO()
    workbook.save(output)
    return output.getvalue()


def _validate_xlsx(content: bytes) -> None:
    if len(content) == 0:
        raise HTTPException(400, "术语表为空")
    if len(content) > MAX_IMPORT_BYTES:
        raise HTTPException(413, "术语表不能超过 5 MB")
    try:
        with zipfile.ZipFile(BytesIO(content)) as archive:
            entries = archive.infolist()
            if len(entries) > MAX_ARCHIVE_ENTRIES:
                raise HTTPException(400, "术语表包含过多内部条目")
            total = 0
            for entry in entries:
                part = PurePosixPath(entry.filename)
                if part.is_absolute() or ".." in part.parts:
                    raise HTTPException(400, "术语表包含不安全路径")
                total += entry.file_size
                if total > MAX_ARCHIVE_UNCOMPRESSED:
                    raise HTTPException(400, "术语表解压规模超限")
    except zipfile.BadZipFile as exc:
        raise HTTPException(400, "请上传有效的 .xlsx 术语表") from exc


def parse_global_terms(content: bytes) -> tuple[list[dict[str, object]], list[str]]:
    _validate_xlsx(content)
    try:
        workbook = load_workbook(BytesIO(content), read_only=True, data_only=False, keep_links=False)
    except Exception as exc:
        raise HTTPException(400, "无法读取术语表，请使用下载的 .xlsx 模板") from exc
    try:
        sheet = workbook[SHEET_NAME] if SHEET_NAME in workbook.sheetnames else workbook.active
        header = tuple(_import_text(cell.value) for cell in next(sheet.iter_rows(min_row=1, max_row=1)))
        is_legacy = header[:len(LEGACY_HEADERS)] == LEGACY_HEADERS
        if header[:len(HEADERS)] != HEADERS and not is_legacy:
            raise HTTPException(400, "术语表表头必须为：英文术语、中文标准译法、英文别名（逗号分隔）、中文别名（逗号分隔，可选）")
        terms: list[dict[str, object]] = []
        errors: list[str] = []
        seen: set[str] = set()
        for row_number, row in enumerate(sheet.iter_rows(min_row=2, max_col=3 if is_legacy else 4), start=2):
            if row_number > MAX_IMPORT_ROWS + 1:
                raise HTTPException(400, f"单次最多导入 {MAX_IMPORT_ROWS} 条术语")
            if any(cell.data_type == "f" for cell in row):
                errors.append(f"第 {row_number} 行：不允许使用 Excel 公式")
                continue
            values = tuple(_import_text(cell.value) for cell in row)
            source, target, aliases = values[:3]
            chinese_aliases = "" if is_legacy else values[3]
            if not source and not target and not aliases and not chinese_aliases:
                continue
            if not source or not target:
                errors.append(f"第 {row_number} 行：英文术语和中文标准译法均为必填项")
                continue
            if len(source) > 500 or len(target) > 500:
                errors.append(f"第 {row_number} 行：英文术语或中文标准译法超过 500 个字符")
                continue
            key = source.casefold()
            if key in seen:
                errors.append(f"第 {row_number} 行：与本表中的“{source}”重复")
                continue
            seen.add(key)
            alias_items = tuple(dict.fromkeys(item.strip() for item in aliases.replace("，", ",").split(",") if item.strip()))
            chinese_alias_items = tuple(dict.fromkeys(item.strip() for item in chinese_aliases.replace("，", ",").split(",") if item.strip()))
            terms.append({"source": source, "target": target, "aliases": alias_items, "chinese_aliases": chinese_alias_items})
        return terms, errors
    finally:
        workbook.close()
