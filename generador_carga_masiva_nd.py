# -*- coding: utf-8 -*-
"""Generador de Carga Masiva ND - Colegio de Notarios de Lima."""
from __future__ import annotations

import os
import re
import subprocess
import sys
import zipfile
from dataclasses import dataclass, field
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from xml.sax.saxutils import escape as xml_escape
import xml.etree.ElementTree as ET

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except Exception:
    tk = None
    filedialog = messagebox = ttk = None

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore
    DND_AVAILABLE = True
except Exception:
    DND_FILES = None
    TkinterDnD = None
    DND_AVAILABLE = False

APP_TITLE = "Generador de Carga Masiva ND"
APP_SUBTITLE = "Colegio de Notarios de Lima"
APP_VERSION = "4.0"
BG = "#f6f1e6"
CARD = "#ffffff"
ACCENT = "#7b1010"
ACCENT2 = "#b08b2b"
TEXT = "#2b2b2b"
MUTED = "#666666"
SUCCESS = "#176b3a"
WARNING = "#a86600"
DANGER = "#a12d2d"
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
FORBIDDEN_INVOICE = re.compile(r'[;,\$%&@\(\)ÑñüÜ¿\?¡!\"“”‘’áéíóúÁÉÍÓÚàèìòùÀÈÌÒÙäëïöüÄËÏÖÜ]')
MONTHS_ES = {
    "ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4,
    "MAYO": 5, "JUNIO": 6, "JULIO": 7, "AGOSTO": 8,
    "SEPTIEMBRE": 9, "SETIEMBRE": 9, "OCTUBRE": 10,
    "NOVIEMBRE": 11, "DICIEMBRE": 12,
}


def resource_path(name: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / name


def open_path(path: str):
    if not path:
        return
    p = Path(path)
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(p))
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(p)])
        else:
            subprocess.Popen(["xdg-open", str(p)])
    except Exception:
        pass


def col_letters_to_index(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch.upper()) - ord("A") + 1)
    return n - 1


def cell_ref_to_coords(ref: str) -> Tuple[int, int]:
    m = re.match(r"([A-Z]+)(\d+)", ref.upper())
    if not m:
        raise ValueError(f"Referencia de celda inválida: {ref}")
    return int(m.group(2)) - 1, col_letters_to_index(m.group(1))


def normalize_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def decimal_from(value) -> Decimal:
    txt = normalize_text(value).replace(" ", "")
    if not txt:
        return Decimal("0")
    if "," in txt and "." not in txt:
        txt = txt.replace(",", ".")
    try:
        return Decimal(txt)
    except InvalidOperation:
        return Decimal("0")


def money2(value: Decimal) -> str:
    q = value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return format(q, ".2f")


class SimpleXlsxReader:
    """Lector XLSX mínimo: extrae los valores de la primera hoja."""

    def __init__(self, path: str):
        self.path = Path(path)
        self.shared_strings: List[str] = []

    def _read_shared_strings(self, zf: zipfile.ZipFile):
        name = "xl/sharedStrings.xml"
        if name not in zf.namelist():
            self.shared_strings = []
            return
        root = ET.fromstring(zf.read(name))
        self.shared_strings = [
            "".join((t.text or "") for t in si.iter(f"{{{NS_MAIN}}}t"))
            for si in root.findall(f"{{{NS_MAIN}}}si")
        ]

    def _first_sheet_path(self, zf: zipfile.ZipFile) -> str:
        wb_root = ET.fromstring(zf.read("xl/workbook.xml"))
        sheets = wb_root.find(f"{{{NS_MAIN}}}sheets")
        if sheets is None or len(sheets) == 0:
            raise ValueError("El archivo XLSX no contiene hojas.")
        rid = list(sheets)[0].attrib.get(f"{{{NS_REL}}}id")
        rel_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        for rel in rel_root:
            if rel.attrib.get("Id") == rid:
                target = rel.attrib.get("Target", "").lstrip("/")
                return target if target.startswith("xl/") else "xl/" + target
        raise ValueError("No se pudo ubicar la primera hoja del XLSX.")

    def read_first_sheet(self) -> Dict[Tuple[int, int], object]:
        if not self.path.exists():
            raise FileNotFoundError(self.path)
        with zipfile.ZipFile(self.path, "r") as zf:
            self._read_shared_strings(zf)
            root = ET.fromstring(zf.read(self._first_sheet_path(zf)))
            data: Dict[Tuple[int, int], object] = {}
            for c in root.iter(f"{{{NS_MAIN}}}c"):
                ref = c.attrib.get("r")
                if not ref:
                    continue
                row, col = cell_ref_to_coords(ref)
                ctype = c.attrib.get("t", "")
                value: object = ""
                if ctype == "inlineStr":
                    is_node = c.find(f"{{{NS_MAIN}}}is")
                    if is_node is not None:
                        value = "".join((t.text or "") for t in is_node.iter(f"{{{NS_MAIN}}}t"))
                else:
                    v = c.find(f"{{{NS_MAIN}}}v")
                    raw = v.text if v is not None else None
                    if raw is None:
                        value = ""
                    elif ctype == "s":
                        try:
                            value = self.shared_strings[int(raw)]
                        except Exception:
                            value = raw
                    elif ctype == "b":
                        value = raw == "1"
                    elif ctype == "str":
                        value = raw
                    else:
                        try:
                            num = float(raw)
                            value = int(num) if num.is_integer() else num
                        except Exception:
                            value = raw
                data[(row, col)] = value
            return data


@dataclass
class Record:
    excel_row: int
    serie: str
    numero: str
    proveedor: str
    igv: Decimal
    no_gravadas: Decimal
    issues: List[str] = field(default_factory=list)

    @property
    def status(self) -> str:
        return "OK" if not self.issues else "OBSERVADO"


@dataclass
class Analysis:
    source_path: str
    source_name: str
    ruc: str
    period: str
    records: List[Record]
    global_issues: List[str] = field(default_factory=list)

    @property
    def total(self) -> int:
        return len(self.records)

    @property
    def observed(self) -> int:
        return sum(1 for r in self.records if r.issues)

    @property
    def valid(self) -> int:
        return self.total - self.observed


@dataclass
class FileResult:
    source_name: str
    source_path: str
    ruc: str
    period: str
    found: int
    valid: int
    observed: int
    status: str
    output_folder: str
    correlativo: int = 1
    txt_path: str = ""              # TXT oficial: solo registros válidos
    zip_path: str = ""              # ZIP oficial: solo registros válidos
    valid_txt_path: str = ""        # Vista/control de válidos
    observed_txt_path: str = ""     # Vista/control de observados
    report_path: str = ""
    messages: List[str] = field(default_factory=list)
    detail_rows: List[Record] = field(default_factory=list)

def detect_ruc(cells: Dict[Tuple[int, int], object]) -> str:
    for value in cells.values():
        m = re.search(r"\bRUC\s*:?\s*(\d{11})\b", normalize_text(value), flags=re.I)
        if m:
            return m.group(1)
    return ""


def detect_period(cells: Dict[Tuple[int, int], object]) -> str:
    for value in cells.values():
        text = normalize_text(value).upper()
        for name, month in MONTHS_ES.items():
            if name in text:
                m = re.search(r"\b(20\d{2})\b", text)
                if m:
                    return f"{int(m.group(1)):04d}{month:02d}"
    for value in cells.values():
        text = normalize_text(value)
        m = re.search(r"\b(20\d{2}(?:0[1-9]|1[0-2]))\b", text)
        if m:
            return m.group(1)
    return ""


def find_header_row(cells: Dict[Tuple[int, int], object]) -> Optional[int]:
    max_row = max((r for r, _ in cells), default=-1)
    for r in range(max_row + 1):
        row_values = {c: normalize_text(v).upper() for (rr, c), v in cells.items() if rr == r}
        if row_values.get(4) in {"T.D", "TD"} and row_values.get(5) == "SERIE":
            return r
    return None


def analyze_xlsx(path: str, igv_mode: str, rate_pct: Decimal = Decimal("18")) -> Analysis:
    cells = SimpleXlsxReader(path).read_first_sheet()
    ruc = detect_ruc(cells)
    period = detect_period(cells)
    header_row = find_header_row(cells)
    global_issues: List[str] = []
    if not ruc:
        global_issues.append("No se pudo detectar el RUC de 11 dígitos en el encabezado.")
    if not period:
        global_issues.append("No se pudo detectar el período aaaamm en el encabezado.")
    if header_row is None:
        global_issues.append("No se encontró la fila de cabecera con T.D. y Serie.")
        return Analysis(path, Path(path).name, ruc, period, [], global_issues)

    max_row = max((r for r, _ in cells), default=header_row)
    records: List[Record] = []
    for r in range(header_row + 1, max_row + 1):
        if normalize_text(cells.get((r, 4), "")) != "91":
            continue
        serie = normalize_text(cells.get((r, 5), ""))
        numero = normalize_text(cells.get((r, 7), ""))
        proveedor = normalize_text(cells.get((r, 10), ""))
        igv_registered = (
            decimal_from(cells.get((r, 12), 0))
            + decimal_from(cells.get((r, 14), 0))
            + decimal_from(cells.get((r, 16), 0))
        )
        no_gravadas = decimal_from(cells.get((r, 17), 0))
        igv = no_gravadas * rate_pct / Decimal("100") if igv_mode == "calculated" else igv_registered
        records.append(Record(r + 1, serie, numero, proveedor, igv, no_gravadas))

    if not records:
        global_issues.append("No se encontraron registros con T.D. = 91.")
    return Analysis(path, Path(path).name, ruc, period, records, global_issues)


def validate_analysis(analysis: Analysis, ruc: str, period: str, fill_nnnn: bool = True):
    analysis.global_issues = list(dict.fromkeys(analysis.global_issues))
    for rec in analysis.records:
        rec.issues.clear()

    if not re.fullmatch(r"\d{11}", ruc or ""):
        analysis.global_issues.append("El RUC de salida debe tener exactamente 11 dígitos.")
    if not re.fullmatch(r"\d{6}", period or ""):
        analysis.global_issues.append("El período de salida debe tener formato aaaamm.")
    elif not (1 <= int(period[4:6]) <= 12):
        analysis.global_issues.append("El mes del período debe estar entre 01 y 12.")

    seen = set()
    for rec in analysis.records:
        if fill_nnnn and not rec.serie:
            rec.serie = "NNNN"
        if not rec.serie:
            rec.issues.append("P101 Serie obligatoria.")
        if len(rec.serie) > 10:
            rec.issues.append("P102 Serie excede 10 caracteres.")
        if not rec.numero:
            rec.issues.append("P101 Número de invoice obligatorio.")
        if len(rec.numero) > 15:
            rec.issues.append(f"P102 Número excede 15 caracteres ({len(rec.numero)}).")
        if rec.numero and set(rec.numero) <= {"0"}:
            rec.issues.append("P103 Número de invoice no puede ser todo cero.")
        if FORBIDDEN_INVOICE.search(rec.numero):
            rec.issues.append("P109 Número contiene caracteres no permitidos.")
        if rec.igv <= 0:
            rec.issues.append("P105 IGV debe ser mayor que cero.")

        key = (ruc, period, rec.serie, rec.numero)
        if key in seen:
            rec.issues.append("P107 Invoice duplicado en el archivo.")
        seen.add(key)


def output_lines(records: List[Record], period: str) -> List[str]:
    return [f"{period}|{r.serie}|{r.numero}|{money2(r.igv)}|" for r in records]


def _write_txt(path: Path, records: List[Record], period: str):
    lines = output_lines(records, period)
    path.write_text(("\n".join(lines) + "\n") if lines else "", encoding="utf-8", newline="\n")


def generate_output_for_analysis(analysis: Analysis, output_root: str, ruc: str, period: str, correlativo: int):
    """Genera siempre TXT separados para válidos y observados.

    Salidas por archivo:
    - *_VALIDOS.txt: registros que pasan todas las validaciones.
    - *_OBSERVADOS.txt: registros observados, conservando el mismo formato pipe del requerimiento.
    - <stem>.txt + <stem>.zip: archivo oficial de carga con SOLO los válidos.

    No se genera CSV: el entregable operativo es TXT/ZIP, tal como pide el requerimiento.
    Las causas de observación se muestran dentro de la interfaz y en los reportes auxiliares Excel/PDF.
    """
    out_folder = Path(output_root) / f"SALIDA_ND_{Path(analysis.source_name).stem}"
    out_folder.mkdir(parents=True, exist_ok=True)

    stem = f"{ruc}_ND_{period}_{correlativo:02d}"
    valid_records = [r for r in analysis.records if not r.issues]
    observed_records = [r for r in analysis.records if r.issues]

    valid_preview = out_folder / f"{stem}_VALIDOS.txt"
    observed_preview = out_folder / f"{stem}_OBSERVADOS.txt"
    _write_txt(valid_preview, valid_records, period)
    _write_txt(observed_preview, observed_records, period)

    txt_path = ""
    zip_path = ""
    if not analysis.global_issues and valid_records:
        official_txt = out_folder / f"{stem}.txt"
        official_zip = out_folder / f"{stem}.zip"
        _write_txt(official_txt, valid_records, period)
        with zipfile.ZipFile(official_zip, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            zf.write(official_txt, arcname=official_txt.name)
        txt_path = str(official_txt)
        zip_path = str(official_zip)

    return str(out_folder), txt_path, zip_path, str(valid_preview), str(observed_preview)

def refresh_result_files(res: FileResult):
    """Revalida un resultado editado en la interfaz y regenera sus salidas."""
    analysis = Analysis(
        source_path=res.source_path,
        source_name=res.source_name,
        ruc=res.ruc,
        period=res.period,
        records=res.detail_rows,
        global_issues=list(res.messages),
    )
    # Las observaciones globales detectadas originalmente que ya no aplican a RUC/período
    # se recalculan limpiando las de formato de salida.
    analysis.global_issues = [m for m in analysis.global_issues if not (
        m.startswith("El RUC de salida") or m.startswith("El período de salida") or m.startswith("El mes del período")
    )]
    validate_analysis(analysis, res.ruc, res.period, True)
    folder, txt, zipp, valid_txt, observed_txt = generate_output_for_analysis(
        analysis, str(Path(res.output_folder).parent), res.ruc, res.period, res.correlativo
    )
    res.output_folder = folder
    res.txt_path = txt
    res.zip_path = zipp
    res.valid_txt_path = valid_txt
    res.observed_txt_path = observed_txt
    res.report_path = ""
    res.messages = list(dict.fromkeys(analysis.global_issues))
    res.found = analysis.total
    res.valid = analysis.valid
    res.observed = analysis.observed
    if res.valid and res.observed:
        res.status = "PARCIAL"
    elif res.valid and not res.observed:
        res.status = "COMPLETADO"
    elif res.observed:
        res.status = "OBSERVADO"
    else:
        res.status = "ERROR"
    return res


def process_files(
    input_paths: List[str], output_root: str, base_correlativo: int,
    igv_mode: str, rate_pct: Decimal, fill_nnnn: bool,
    override_ruc: str = "", override_period: str = "",
) -> List[FileResult]:
    results: List[FileResult] = []
    correlativo = base_correlativo
    for path in input_paths:
        try:
            analysis = analyze_xlsx(path, igv_mode, rate_pct)
            final_ruc = override_ruc.strip() or analysis.ruc
            final_period = override_period.strip() or analysis.period
            validate_analysis(analysis, final_ruc, final_period, fill_nnnn)
            folder, txt_path, zip_path, valid_txt, observed_txt = generate_output_for_analysis(
                analysis, output_root, final_ruc, final_period, correlativo
            )
            if analysis.valid and analysis.observed:
                status = "PARCIAL"
            elif analysis.valid and not analysis.observed:
                status = "COMPLETADO"
            elif analysis.observed:
                status = "OBSERVADO"
            else:
                status = "ERROR"
            results.append(FileResult(
                source_name=analysis.source_name,
                source_path=analysis.source_path,
                ruc=final_ruc,
                period=final_period,
                found=analysis.total,
                valid=analysis.valid,
                observed=analysis.observed,
                status=status,
                output_folder=folder,
                correlativo=correlativo,
                txt_path=txt_path,
                zip_path=zip_path,
                valid_txt_path=valid_txt,
                observed_txt_path=observed_txt,
                report_path="",
                messages=list(dict.fromkeys(analysis.global_issues)),
                detail_rows=analysis.records,
            ))
        except Exception as exc:
            results.append(FileResult(
                source_name=Path(path).name,
                source_path=path,
                ruc="", period="", found=0, valid=0, observed=0,
                status="ERROR", output_folder=output_root, correlativo=correlativo,
                messages=[str(exc)], detail_rows=[],
            ))
        correlativo += 1
    return results

# ---------- Exportador XLSX de resumen (sin dependencias externas) ----------
def _xlsx_col_letter(n: int) -> str:
    out = ""
    while n:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


def _xlsx_cell(ref: str, value, style: int = 0) -> str:
    sattr = f' s="{style}"' if style else ""
    if value is None:
        return f'<c r="{ref}"{sattr}/>'
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f'<c r="{ref}"{sattr}><v>{value}</v></c>'
    text = xml_escape(str(value))
    return f'<c r="{ref}" t="inlineStr"{sattr}><is><t>{text}</t></is></c>'


def _xlsx_sheet_xml(rows, widths=None, freeze_header=True):
    cols = ""
    if widths:
        parts = [f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>' for i, w in enumerate(widths, 1)]
        cols = "<cols>" + "".join(parts) + "</cols>"
    row_xml = []
    for r_idx, row in enumerate(rows, 1):
        cells = []
        for c_idx, item in enumerate(row, 1):
            value, style = (item if isinstance(item, tuple) and len(item) == 2 else (item, 0))
            cells.append(_xlsx_cell(f"{_xlsx_col_letter(c_idx)}{r_idx}", value, style))
        row_xml.append(f'<row r="{r_idx}">' + "".join(cells) + "</row>")
    pane = '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>' if freeze_header else ""
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<sheetViews><sheetView workbookViewId="0">{pane}</sheetView></sheetViews>'
        f'{cols}<sheetData>{"".join(row_xml)}</sheetData></worksheet>'
    )


def export_summary_xlsx(results: List[FileResult], output_root: str) -> str:
    out = Path(output_root)
    out.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = out / f"RESUMEN_CARGA_MASIVA_ND_{stamp}.xlsx"

    resumen = [[
        ("Archivo", 1), ("RUC", 1), ("Período", 1), ("ND encontrados", 1),
        ("Válidos", 1), ("Observados", 1), ("Estado", 1),
        ("ZIP oficial", 1), ("TXT válidos", 1), ("TXT observados", 1),
    ]]
    detalle = [[
        ("Archivo", 1), ("Fila Excel", 1), ("Proveedor", 1), ("Serie", 1),
        ("Invoice", 1), ("No gravadas", 1), ("IGV salida", 1),
        ("Estado", 1), ("Detalle", 1),
    ]]

    for res in results:
        resumen.append([
            res.source_name, res.ruc, res.period, res.found, res.valid,
            res.observed, res.status, res.zip_path or "", res.valid_txt_path or "", res.observed_txt_path or "",
        ])
        for r in res.detail_rows:
            detalle.append([
                res.source_name, r.excel_row, r.proveedor, r.serie, r.numero,
                float(r.no_gravadas), float(r.igv), r.status,
                " | ".join(r.issues) if r.issues else "Sin observaciones",
            ])

    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>'''
    root_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''
    workbook = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Resumen" sheetId="1" r:id="rId1"/><sheet name="Detalle" sheetId="2" r:id="rId2"/></sheets></workbook>'''
    wb_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''
    styles = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2"><font><sz val="10"/><name val="Calibri"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="10"/><name val="Calibri"/></font></fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF7B1010"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFill="1" applyFont="1"/></cellXfs>
</styleSheet>'''

    sheet1 = _xlsx_sheet_xml(resumen, [28, 14, 11, 14, 10, 12, 14, 44, 44, 44])
    sheet2 = _xlsx_sheet_xml(detalle, [24, 10, 28, 12, 22, 14, 12, 14, 50])
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", root_rels)
        zf.writestr("xl/workbook.xml", workbook)
        zf.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        zf.writestr("xl/styles.xml", styles)
        zf.writestr("xl/worksheets/sheet1.xml", sheet1)
        zf.writestr("xl/worksheets/sheet2.xml", sheet2)
    return str(path)


# ---------- Exportador PDF de resumen (sin dependencias externas) ----------
def _pdf_escape(text: str) -> str:
    return str(text).replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


def _pdf_text_line(text, x, y, size=9, bold=False) -> str:
    font = "F2" if bold else "F1"
    return f"BT /{font} {size} Tf {x:.1f} {y:.1f} Td ({_pdf_escape(text)}) Tj ET\n"


def _build_simple_pdf(pages: List[str]) -> bytes:
    objects: List[bytes] = []
    objects.append(b"<< /Type /Catalog /Pages 2 0 R >>")
    kids = [f"{5 + i * 2} 0 R" for i in range(len(pages))]
    objects.append(f"<< /Type /Pages /Kids [{' '.join(kids)}] /Count {len(pages)} >>".encode("ascii"))
    objects.append(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>")
    objects.append(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>")

    for i, content in enumerate(pages):
        content_obj = 6 + i * 2
        page = (
            f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] "
            f"/Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> /Contents {content_obj} 0 R >>"
        ).encode("ascii")
        content_bytes = content.encode("cp1252", errors="replace")
        stream = f"<< /Length {len(content_bytes)} >>\nstream\n".encode("ascii") + content_bytes + b"endstream"
        objects.extend([page, stream])

    out = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    offsets = [0]
    for idx, obj in enumerate(objects, start=1):
        offsets.append(len(out))
        out.extend(f"{idx} 0 obj\n".encode("ascii"))
        out.extend(obj)
        out.extend(b"\nendobj\n")
    xref = len(out)
    out.extend(f"xref\n0 {len(objects) + 1}\n".encode("ascii"))
    out.extend(b"0000000000 65535 f \n")
    for off in offsets[1:]:
        out.extend(f"{off:010d} 00000 n \n".encode("ascii"))
    out.extend(f"trailer << /Size {len(objects) + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF".encode("ascii"))
    return bytes(out)


def export_summary_pdf(results: List[FileResult], output_root: str) -> str:
    out = Path(output_root)
    out.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = out / f"RESUMEN_CARGA_MASIVA_ND_{stamp}.pdf"

    total_files = len(results)
    total_nd = sum(r.found for r in results)
    total_valid = sum(r.valid for r in results)
    total_obs = sum(r.observed for r in results)
    rows_per_page = 28
    pages: List[str] = []

    for start in range(0, max(len(results), 1), rows_per_page):
        y = 800
        content = ""
        content += _pdf_text_line("COLEGIO DE NOTARIOS DE LIMA", 40, y, 16, True); y -= 22
        content += _pdf_text_line("Resumen - Generador de Carga Masiva ND", 40, y, 12, True); y -= 18
        content += _pdf_text_line(datetime.now().strftime("Fecha de generación: %d/%m/%Y %H:%M"), 40, y, 9); y -= 22
        if start == 0:
            content += _pdf_text_line(
                f"Archivos procesados: {total_files}    ND encontrados: {total_nd}    Válidos: {total_valid}    Observados: {total_obs}",
                40, y, 9, True,
            ); y -= 28
        content += _pdf_text_line("Archivo", 40, y, 8, True)
        content += _pdf_text_line("RUC", 230, y, 8, True)
        content += _pdf_text_line("Periodo", 315, y, 8, True)
        content += _pdf_text_line("ND", 370, y, 8, True)
        content += _pdf_text_line("Validos", 410, y, 8, True)
        content += _pdf_text_line("Obs.", 465, y, 8, True)
        content += _pdf_text_line("Estado", 505, y, 8, True)
        y -= 14

        for res in results[start:start + rows_per_page]:
            name = (res.source_name[:31] + "...") if len(res.source_name) > 34 else res.source_name
            content += _pdf_text_line(name, 40, y, 7)
            content += _pdf_text_line(res.ruc, 230, y, 7)
            content += _pdf_text_line(res.period, 315, y, 7)
            content += _pdf_text_line(str(res.found), 375, y, 7)
            content += _pdf_text_line(str(res.valid), 420, y, 7)
            content += _pdf_text_line(str(res.observed), 470, y, 7)
            content += _pdf_text_line(res.status, 505, y, 7)
            y -= 17

        y -= 14
        content += _pdf_text_line("Los archivos observados cuentan con un CSV de validación para revisión.", 40, y, 8)
        pages.append(content)

    path.write_bytes(_build_simple_pdf(pages))
    return str(path)

if tk is not None:
    class SummaryCard(tk.Frame):
        def __init__(self, master, title: str, value_color=ACCENT):
            super().__init__(master, bg=CARD, highlightthickness=1, highlightbackground="#ddd8ce")
            tk.Label(self, text=title, font=("Segoe UI", 10, "bold"), bg=CARD, fg=MUTED).pack(anchor="w", padx=14, pady=(10, 0))
            self.var = tk.StringVar(value="0")
            tk.Label(self, textvariable=self.var, font=("Segoe UI", 24, "bold"), bg=CARD, fg=value_color).pack(anchor="w", padx=14, pady=(0, 10))

        def set(self, value):
            self.var.set(str(value))


    BaseTk = TkinterDnD.Tk if DND_AVAILABLE else tk.Tk


    class App(BaseTk):
        def __init__(self):
            super().__init__()
            self.withdraw()
            self.title(f"{APP_TITLE} - {APP_SUBTITLE}")
            self.configure(bg=BG)
            self.geometry("1260x850")
            self.minsize(1120, 760)
            self.selected_files: List[str] = []
            self.results_map: Dict[str, FileResult] = {}
            self.last_results: List[FileResult] = []
            self.last_summary_xlsx = ""
            self.last_summary_pdf = ""
            self.logo_img = None
            self.icon_img = None
            self.output_var = tk.StringVar(value=str(Path.home() / "Desktop" / "Salida_Carga_Masiva_ND"))
            self.ruc_override_var = tk.StringVar(value="")
            self.period_override_var = tk.StringVar(value="")
            self.base_corr_var = tk.StringVar(value="01")
            self.igv_mode_var = tk.StringVar(value="registered")
            self.rate_var = tk.StringVar(value="18")
            self.fill_nnnn_var = tk.BooleanVar(value=True)
            self.status_var = tk.StringVar(value="Agregue uno o varios archivos Excel y luego presione Procesar.")
            self._apply_icon()
            self._build_ui()
            self._show_splash()

        def _apply_icon(self):
            ico = resource_path("cnl_nd.ico")
            png = resource_path("logo_cnl.png")
            try:
                if sys.platform.startswith("win") and ico.exists():
                    self.iconbitmap(str(ico))
                elif png.exists():
                    self.icon_img = tk.PhotoImage(file=str(png))
                    self.iconphoto(True, self.icon_img)
            except Exception:
                pass

        def _show_splash(self):
            splash = tk.Toplevel(self)
            splash.overrideredirect(True)
            splash.configure(bg=BG)
            w, h = 560, 300
            x = (splash.winfo_screenwidth() - w) // 2
            y = (splash.winfo_screenheight() - h) // 2
            splash.geometry(f"{w}x{h}+{x}+{y}")
            card = tk.Frame(splash, bg=CARD, highlightthickness=2, highlightbackground=ACCENT2)
            card.pack(fill="both", expand=True, padx=10, pady=10)
            logo_path = resource_path("logo_cnl.png")
            try:
                self.logo_img = tk.PhotoImage(file=str(logo_path))
                tk.Label(card, image=self.logo_img, bg=CARD).pack(pady=(28, 10))
            except Exception:
                pass
            tk.Label(card, text=APP_TITLE, bg=CARD, fg=ACCENT, font=("Segoe UI", 18, "bold")).pack()
            tk.Label(card, text=APP_SUBTITLE, bg=CARD, fg=TEXT, font=("Segoe UI", 11)).pack(pady=(4, 0))
            tk.Label(card, text=f"Versión {APP_VERSION}", bg=CARD, fg=MUTED, font=("Segoe UI", 9)).pack(pady=(10, 0))
            self.after(950, lambda: (splash.destroy(), self.deiconify(), self.lift()))

        def _build_ui(self):
            self._build_header()
            self._build_input()
            self._build_options()
            self._build_summary()
            self._build_results()
            self._build_footer()

        def _build_header(self):
            frame = tk.Frame(self, bg=BG)
            frame.pack(fill="x", padx=18, pady=(16, 8))
            left = tk.Frame(frame, bg=BG)
            left.pack(side="left", fill="x", expand=True)
            logo_path = resource_path("logo_cnl.png")
            if logo_path.exists():
                try:
                    if self.logo_img is None:
                        self.logo_img = tk.PhotoImage(file=str(logo_path))
                    tk.Label(left, image=self.logo_img, bg=BG).pack(side="left", padx=(0, 16))
                except Exception:
                    pass
            text_wrap = tk.Frame(left, bg=BG)
            text_wrap.pack(side="left", fill="y")
            tk.Label(text_wrap, text=APP_SUBTITLE, bg=BG, fg=ACCENT, font=("Segoe UI", 24, "bold")).pack(anchor="w")
            tk.Label(text_wrap, text=f"{APP_TITLE}  ·  v{APP_VERSION}", bg=BG, fg=TEXT, font=("Segoe UI", 13)).pack(anchor="w", pady=(2, 0))
            tk.Label(
                text_wrap,
                text="Procesamiento masivo de Registros de Compras para No Domiciliados: valida, genera TXT/ZIP y emite reportes de control.",
                bg=BG, fg=MUTED, font=("Segoe UI", 10), wraplength=760, justify="left",
            ).pack(anchor="w", pady=(6, 0))
            self.process_btn = tk.Button(
                frame, text="Procesar Registro(s) de Compras", command=self.process_all,
                bg=ACCENT, fg="white", activebackground=ACCENT, activeforeground="white",
                font=("Segoe UI", 12, "bold"), padx=18, pady=12, relief="flat", cursor="hand2",
            )
            self.process_btn.pack(side="right")

        def _build_input(self):
            frame = tk.Frame(self, bg=CARD, highlightthickness=1, highlightbackground="#ddd8ce")
            frame.pack(fill="x", padx=18, pady=8)
            top = tk.Frame(frame, bg=CARD)
            top.pack(fill="x", padx=14, pady=(12, 6))
            tk.Label(top, text="1. Archivos de entrada", bg=CARD, fg=TEXT, font=("Segoe UI", 12, "bold")).pack(side="left")
            tk.Button(top, text="Agregar Excel(es)", command=self.add_files, bg=ACCENT2, fg="white", relief="flat", padx=10, pady=6, cursor="hand2").pack(side="right", padx=(6, 0))
            tk.Button(top, text="Limpiar lista", command=self.clear_files, bg="#6b6b6b", fg="white", relief="flat", padx=10, pady=6, cursor="hand2").pack(side="right")
            msg = ("Arrastra y suelta aquí los archivos .xlsx\n" if DND_AVAILABLE else "Selecciona uno o varios archivos .xlsx\n") + "o usa el botón 'Agregar Excel(es)'."
            self.drop_label = tk.Label(
                frame, text=msg, bg="#fbfaf7", fg=MUTED, font=("Segoe UI", 10), pady=12,
                highlightthickness=1, highlightbackground="#d6d1c7",
            )
            self.drop_label.pack(fill="x", padx=14, pady=(0, 8))
            if DND_AVAILABLE:
                self.drop_label.drop_target_register(DND_FILES)
                self.drop_label.dnd_bind("<<Drop>>", self.on_drop)
            wrap = tk.Frame(frame, bg=CARD)
            wrap.pack(fill="x", padx=14, pady=(0, 14))
            self.files_listbox = tk.Listbox(wrap, height=5, selectmode=tk.EXTENDED)
            self.files_listbox.pack(side="left", fill="x", expand=True)
            sb = ttk.Scrollbar(wrap, orient="vertical", command=self.files_listbox.yview)
            sb.pack(side="right", fill="y")
            self.files_listbox.configure(yscrollcommand=sb.set)

        def _build_options(self):
            frame = tk.Frame(self, bg=CARD, highlightthickness=1, highlightbackground="#ddd8ce")
            frame.pack(fill="x", padx=18, pady=8)
            tk.Label(frame, text="2. Configuración de procesamiento", bg=CARD, fg=TEXT, font=("Segoe UI", 12, "bold")).grid(row=0, column=0, columnspan=6, sticky="w", padx=14, pady=(12, 8))
            tk.Label(frame, text="Carpeta de salida", bg=CARD, fg=TEXT).grid(row=1, column=0, sticky="e", padx=8, pady=6)
            tk.Entry(frame, textvariable=self.output_var).grid(row=1, column=1, columnspan=4, sticky="ew", padx=8, pady=6)
            tk.Button(frame, text="Elegir", command=self.pick_output, bg="#8f8f8f", fg="white", relief="flat", padx=10, pady=5).grid(row=1, column=5, sticky="w", padx=8, pady=6)
            tk.Label(frame, text="RUC (opcional)", bg=CARD, fg=TEXT).grid(row=2, column=0, sticky="e", padx=8, pady=6)
            tk.Entry(frame, textvariable=self.ruc_override_var, width=18).grid(row=2, column=1, sticky="w", padx=8, pady=6)
            tk.Label(frame, text="Período (opcional)", bg=CARD, fg=TEXT).grid(row=2, column=2, sticky="e", padx=8, pady=6)
            tk.Entry(frame, textvariable=self.period_override_var, width=12).grid(row=2, column=3, sticky="w", padx=8, pady=6)
            tk.Label(frame, text="Correlativo inicial", bg=CARD, fg=TEXT).grid(row=2, column=4, sticky="e", padx=8, pady=6)
            tk.Spinbox(frame, from_=1, to=99, textvariable=self.base_corr_var, width=5, format="%02.0f").grid(row=2, column=5, sticky="w", padx=8, pady=6)
            tk.Label(frame, text="Regla de IGV", bg=CARD, fg=TEXT).grid(row=3, column=0, sticky="e", padx=8, pady=6)
            self.igv_combo = ttk.Combobox(frame, state="readonly", width=46, values=[
                "Usar IGV registrado (columnas I.G.V.)",
                "Calcular porcentaje sobre Adquisiciones no Gravadas",
            ])
            self.igv_combo.current(0)
            self.igv_combo.grid(row=3, column=1, columnspan=2, sticky="w", padx=8, pady=6)
            self.igv_combo.bind("<<ComboboxSelected>>", self.on_igv_mode_change)
            tk.Label(frame, text="Tasa %", bg=CARD, fg=TEXT).grid(row=3, column=3, sticky="e", padx=8, pady=6)
            self.rate_entry = tk.Entry(frame, textvariable=self.rate_var, width=8, state="disabled")
            self.rate_entry.grid(row=3, column=4, sticky="w", padx=8, pady=6)
            tk.Checkbutton(frame, text='Completar serie vacía con "NNNN"', variable=self.fill_nnnn_var, bg=CARD, fg=TEXT, activebackground=CARD).grid(row=3, column=5, sticky="w", padx=8, pady=6)
            tk.Label(frame, text="Si procesa varios Excel, el correlativo se incrementa automáticamente por cada archivo.", bg=CARD, fg=MUTED, font=("Segoe UI", 9)).grid(row=4, column=0, columnspan=6, sticky="w", padx=14, pady=(0, 12))
            frame.grid_columnconfigure(1, weight=1)
            frame.grid_columnconfigure(2, weight=1)
            frame.grid_columnconfigure(4, weight=1)

        def _build_summary(self):
            wrap = tk.Frame(self, bg=BG)
            wrap.pack(fill="x", padx=18, pady=8)
            self.card_files = SummaryCard(wrap, "Archivos procesados", ACCENT)
            self.card_nd = SummaryCard(wrap, "Comprobantes ND encontrados", ACCENT2)
            self.card_ok = SummaryCard(wrap, "Registros válidos", SUCCESS)
            self.card_obs = SummaryCard(wrap, "Registros observados", DANGER)
            cards = [self.card_files, self.card_nd, self.card_ok, self.card_obs]
            for i, card in enumerate(cards):
                card.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 12, 0))
                wrap.grid_columnconfigure(i, weight=1)

        def _build_results(self):
            paned = tk.PanedWindow(self, bg=BG, sashwidth=8, orient="vertical")
            paned.pack(fill="both", expand=True, padx=18, pady=8)
            top = tk.Frame(paned, bg=CARD, highlightthickness=1, highlightbackground="#ddd8ce")
            bottom = tk.Frame(paned, bg=CARD, highlightthickness=1, highlightbackground="#ddd8ce")
            paned.add(top, stretch="always", minsize=220)
            paned.add(bottom, stretch="always", minsize=220)

            tk.Label(top, text="3. Resumen por archivo", bg=CARD, fg=TEXT, font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=14, pady=(12, 8))
            cols = ("archivo", "ruc", "periodo", "encontrados", "validos", "observados", "estado", "txt_validos", "txt_observados", "zip")
            self.results_tree = ttk.Treeview(top, columns=cols, show="headings", height=9)
            heads = {
                "archivo": "Archivo", "ruc": "RUC", "periodo": "Período", "encontrados": "ND",
                "validos": "Válidos", "observados": "Observados", "estado": "Estado",
                "txt_validos": "TXT válidos", "txt_observados": "TXT observados", "zip": "ZIP oficial",
            }
            widths = {"archivo": 200, "ruc": 110, "periodo": 80, "encontrados": 60, "validos": 70, "observados": 90, "estado": 95, "txt_validos": 220, "txt_observados": 220, "zip": 220}
            for c in cols:
                self.results_tree.heading(c, text=heads[c])
                self.results_tree.column(c, width=widths[c], anchor="w")
            self.results_tree.pack(fill="both", expand=True, padx=14, pady=(0, 6))
            self.results_tree.bind("<<TreeviewSelect>>", self.on_result_selected)

            actions = tk.Frame(top, bg=CARD)
            actions.pack(fill="x", padx=14, pady=(0, 12))
            tk.Button(actions, text="Previsualizar TXT", command=self.preview_selected_txt, bg=ACCENT2, fg="white", relief="flat", padx=12, pady=6).pack(side="left")
            tk.Button(actions, text="Corregir / reenviar observados", command=self.remediate_selected, bg=ACCENT, fg="white", relief="flat", padx=12, pady=6).pack(side="left", padx=(8, 0))
            tk.Button(actions, text="Abrir salida del archivo", command=self.open_selected_output, bg="#777777", fg="white", relief="flat", padx=12, pady=6).pack(side="left", padx=(8, 0))
            tk.Label(actions, text="Se generan siempre TXT separados de válidos y observados; el ZIP oficial contiene solo los válidos.", bg=CARD, fg=MUTED, font=("Segoe UI", 9)).pack(side="right")

            tk.Label(bottom, text="4. Detalle del archivo seleccionado", bg=CARD, fg=TEXT, font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=14, pady=(12, 8))
            self.detail_msg = tk.StringVar(value="Seleccione un archivo procesado para ver el detalle.")
            tk.Label(bottom, textvariable=self.detail_msg, bg=CARD, fg=MUTED, wraplength=1180, justify="left").pack(anchor="w", padx=14, pady=(0, 8))
            dcols = ("fila", "proveedor", "serie", "invoice", "no_gravadas", "igv", "estado", "detalle")
            self.detail_tree = ttk.Treeview(bottom, columns=dcols, show="headings", height=10)
            dheads = {"fila": "Fila", "proveedor": "Proveedor", "serie": "Serie", "invoice": "Invoice", "no_gravadas": "No gravadas", "igv": "IGV salida", "estado": "Estado", "detalle": "Validación"}
            dwidths = {"fila": 55, "proveedor": 200, "serie": 90, "invoice": 170, "no_gravadas": 100, "igv": 90, "estado": 90, "detalle": 430}
            for c in dcols:
                self.detail_tree.heading(c, text=dheads[c])
                self.detail_tree.column(c, width=dwidths[c], anchor="w")
            self.detail_tree.pack(fill="both", expand=True, padx=14, pady=(0, 12))

        def _build_footer(self):
            frame = tk.Frame(self, bg=BG)
            frame.pack(fill="x", padx=18, pady=(0, 14))
            tk.Label(frame, textvariable=self.status_var, bg=BG, fg=MUTED, font=("Segoe UI", 10)).pack(side="left")
            tk.Button(frame, text="Abrir carpeta de salida", command=lambda: open_path(self.output_var.get()), bg="#777777", fg="white", relief="flat", padx=10, pady=6).pack(side="right", padx=(6, 0))
            tk.Button(frame, text="Abrir resumen PDF", command=self.open_summary_pdf, bg=ACCENT, fg="white", relief="flat", padx=10, pady=6).pack(side="right", padx=(6, 0))
            tk.Button(frame, text="Abrir resumen Excel", command=self.open_summary_xlsx, bg=SUCCESS, fg="white", relief="flat", padx=10, pady=6).pack(side="right")

        def add_files(self):
            paths = filedialog.askopenfilenames(
                title="Seleccione uno o varios Registros de Compras",
                filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")],
            )
            if paths:
                self._append_files(list(paths))

        def clear_files(self):
            self.selected_files.clear()
            self.files_listbox.delete(0, tk.END)
            self.results_tree.delete(*self.results_tree.get_children())
            self.detail_tree.delete(*self.detail_tree.get_children())
            self.results_map.clear()
            self.last_results = []
            self.last_summary_xlsx = ""
            self.last_summary_pdf = ""
            self.detail_msg.set("Seleccione un archivo procesado para ver el detalle.")
            self._update_summary([])
            self.status_var.set("Lista limpiada. Agregue uno o varios archivos Excel.")

        def _append_files(self, paths: List[str]):
            existing = set(self.selected_files)
            added = 0
            for p in paths:
                p = p.strip().strip("{}")
                if p and p.lower().endswith(".xlsx") and p not in existing:
                    self.selected_files.append(p)
                    self.files_listbox.insert(tk.END, p)
                    existing.add(p)
                    added += 1
            self.status_var.set(f"{len(self.selected_files)} archivo(s) en cola. Se agregaron {added}.")

        def on_drop(self, event):
            try:
                paths = self.tk.splitlist(event.data)
            except Exception:
                paths = [event.data]
            self._append_files(list(paths))

        def pick_output(self):
            path = filedialog.askdirectory(title="Seleccione la carpeta de salida")
            if path:
                self.output_var.set(path)

        def on_igv_mode_change(self, _event=None):
            if self.igv_combo.current() == 1:
                self.igv_mode_var.set("calculated")
                self.rate_entry.configure(state="normal")
            else:
                self.igv_mode_var.set("registered")
                self.rate_entry.configure(state="disabled")

        def _get_rate(self) -> Decimal:
            try:
                return Decimal(self.rate_var.get().replace(",", "."))
            except Exception:
                raise ValueError("La tasa de IGV debe ser un número válido.")

        def process_all(self):
            if not self.selected_files:
                messagebox.showwarning("Sin archivos", "Primero agregue uno o varios archivos .xlsx.")
                return
            try:
                correlativo = int(self.base_corr_var.get())
                if not 1 <= correlativo <= 99:
                    raise ValueError
            except Exception:
                messagebox.showerror("Correlativo inválido", "El correlativo inicial debe estar entre 01 y 99.")
                return
            out = self.output_var.get().strip()
            if not out:
                messagebox.showwarning("Falta carpeta", "Seleccione una carpeta de salida.")
                return

            self.config(cursor="watch")
            self.process_btn.config(state="disabled")
            self.update_idletasks()
            try:
                results = process_files(
                    self.selected_files, out, correlativo,
                    self.igv_mode_var.get(), self._get_rate(), self.fill_nnnn_var.get(),
                    self.ruc_override_var.get(), self.period_override_var.get(),
                )
                self.last_results = results
                self.last_summary_xlsx = export_summary_xlsx(results, out)
                self.last_summary_pdf = export_summary_pdf(results, out)
                self._render_results(results)
            except Exception as exc:
                messagebox.showerror("Error", str(exc))
            finally:
                self.config(cursor="")
                self.process_btn.config(state="normal")

        def _render_results(self, results: List[FileResult]):
            self.results_tree.delete(*self.results_tree.get_children())
            self.results_map.clear()
            for i, res in enumerate(results, 1):
                iid = f"res_{i}"
                self.results_map[iid] = res
                self.results_tree.insert("", "end", iid=iid, values=(
                    res.source_name, res.ruc, res.period, res.found,
                    res.valid, res.observed, res.status,
                    res.valid_txt_path or "—", res.observed_txt_path or "—", res.zip_path or "—",
                ))
            self._update_summary(results)
            completed = sum(1 for r in results if r.status == "COMPLETADO")
            partial = sum(1 for r in results if r.status == "PARCIAL")
            observed = sum(1 for r in results if r.status == "OBSERVADO")
            errors = sum(1 for r in results if r.status == "ERROR")
            self.status_var.set(
                f"Proceso terminado. Completados: {completed}. Parciales: {partial}. Observados: {observed}. Errores: {errors}. TXT separados y reportes generados."
            )
            items = self.results_tree.get_children()
            if items:
                self.results_tree.selection_set(items[0])
                self.on_result_selected()

        def _update_summary(self, results: List[FileResult]):
            self.card_files.set(len(results))
            self.card_nd.set(sum(r.found for r in results))
            self.card_ok.set(sum(r.valid for r in results))
            self.card_obs.set(sum(r.observed for r in results))

        def on_result_selected(self, _event=None):
            sel = self.results_tree.selection()
            if not sel:
                return
            res = self.results_map.get(sel[0])
            if not res:
                return
            self.detail_tree.delete(*self.detail_tree.get_children())
            for r in res.detail_rows:
                self.detail_tree.insert("", "end", values=(
                    r.excel_row, r.proveedor, r.serie, r.numero,
                    money2(r.no_gravadas), money2(r.igv), r.status,
                    " | ".join(r.issues) if r.issues else "Sin observaciones",
                ))
            parts = [
                f"Archivo: {res.source_name}", f"Estado: {res.status}",
                f"RUC: {res.ruc or 'No detectado'}", f"Período: {res.period or 'No detectado'}",
                f"Salida: {res.output_folder}",
            ]
            if res.messages:
                parts.append("Observaciones globales: " + " | ".join(res.messages))
            elif res.zip_path:
                parts.append(f"ZIP generado: {res.zip_path}")
            self.detail_msg.set("    •    ".join(parts))

        def _selected_result(self) -> Optional[FileResult]:
            sel = self.results_tree.selection()
            return self.results_map.get(sel[0]) if sel else None

        def open_selected_output(self):
            res = self._selected_result()
            if not res:
                messagebox.showinfo("Sin selección", "Seleccione primero un archivo del resumen.")
                return
            open_path(res.output_folder)

        def preview_selected_txt(self):
            res = self._selected_result()
            if not res:
                messagebox.showinfo("Sin selección", "Seleccione primero un archivo del resumen.")
                return

            win = tk.Toplevel(self)
            win.title(f"Previsualización TXT - {res.source_name}")
            win.geometry("900x600")
            win.configure(bg=BG)
            tk.Label(win, text="Previsualización de archivos TXT", bg=BG, fg=ACCENT, font=("Segoe UI", 16, "bold")).pack(anchor="w", padx=16, pady=(14, 4))
            tk.Label(win, text="Seleccione qué conjunto desea revisar. El TXT oficial y el ZIP nunca incluyen registros observados.", bg=BG, fg=MUTED, font=("Segoe UI", 9)).pack(anchor="w", padx=16, pady=(0, 10))

            selector = tk.StringVar(value="Válidos")
            options = ["Válidos", "Observados", "TXT oficial de carga"]
            combo = ttk.Combobox(win, textvariable=selector, values=options, state="readonly", width=28)
            combo.pack(anchor="w", padx=16, pady=(0, 8))

            text = tk.Text(win, wrap="none", font=("Consolas", 10), bg="white", fg=TEXT)
            text.pack(fill="both", expand=True, padx=16, pady=(0, 10))
            bottom = tk.Frame(win, bg=BG)
            bottom.pack(fill="x", padx=16, pady=(0, 14))
            count_var = tk.StringVar(value="")
            tk.Label(bottom, textvariable=count_var, bg=BG, fg=MUTED).pack(side="left")

            def render(_event=None):
                mode = selector.get()
                if mode == "Válidos":
                    rows = [r for r in res.detail_rows if not r.issues]
                    path = res.valid_txt_path
                elif mode == "Observados":
                    rows = [r for r in res.detail_rows if r.issues]
                    path = res.observed_txt_path
                else:
                    rows = [r for r in res.detail_rows if not r.issues]
                    path = res.txt_path
                content = "\n".join(output_lines(rows, res.period))
                if content:
                    content += "\n"
                text.configure(state="normal")
                text.delete("1.0", tk.END)
                text.insert("1.0", content or "(sin registros para este conjunto)")
                text.configure(state="disabled")
                count_var.set(f"{len(rows)} registro(s)  •  Archivo: {path or 'no generado'}")

            combo.bind("<<ComboboxSelected>>", render)
            tk.Button(bottom, text="Abrir archivo", command=lambda: open_path(
                res.valid_txt_path if selector.get()=="Válidos" else res.observed_txt_path if selector.get()=="Observados" else res.txt_path
            ), bg="#777777", fg="white", relief="flat", padx=10, pady=5).pack(side="right")
            render()

        def remediate_selected(self):
            res = self._selected_result()
            if not res:
                messagebox.showinfo("Sin selección", "Seleccione primero un archivo del resumen.")
                return
            observed_rows = [r for r in res.detail_rows if r.issues]
            if not observed_rows:
                messagebox.showinfo("Sin observados", "Este archivo ya no tiene registros observados.")
                return

            win = tk.Toplevel(self)
            win.title(f"Corregir observados - {res.source_name}")
            win.geometry("1100x650")
            win.configure(bg=BG)
            tk.Label(win, text="Corrección y reenvío de observados", bg=BG, fg=ACCENT, font=("Segoe UI", 16, "bold")).pack(anchor="w", padx=16, pady=(14, 4))
            tk.Label(win, text="Seleccione un registro, corrija Serie / Invoice / IGV y guarde. El sistema revalida todo el archivo y mueve automáticamente el registro a Válidos cuando supera las reglas.", bg=BG, fg=MUTED, font=("Segoe UI", 9), wraplength=1040, justify="left").pack(anchor="w", padx=16, pady=(0, 10))

            cols=("fila","proveedor","serie","invoice","igv","observacion")
            tree=ttk.Treeview(win,columns=cols,show="headings",height=11)
            heads={"fila":"Fila","proveedor":"Proveedor","serie":"Serie","invoice":"Invoice","igv":"IGV","observacion":"Observación"}
            widths={"fila":55,"proveedor":220,"serie":100,"invoice":190,"igv":90,"observacion":420}
            for c in cols:
                tree.heading(c,text=heads[c]); tree.column(c,width=widths[c],anchor="w")
            tree.pack(fill="both",expand=True,padx=16,pady=(0,10))

            edit=tk.Frame(win,bg=CARD,highlightthickness=1,highlightbackground="#ddd8ce"); edit.pack(fill="x",padx=16,pady=(0,10))
            serie_var=tk.StringVar(); invoice_var=tk.StringVar(); igv_var=tk.StringVar(); selected_row=tk.StringVar(value="")
            tk.Label(edit,textvariable=selected_row,bg=CARD,fg=TEXT,font=("Segoe UI",10,"bold")).grid(row=0,column=0,columnspan=6,sticky="w",padx=10,pady=(8,4))
            tk.Label(edit,text="Serie",bg=CARD).grid(row=1,column=0,sticky="e",padx=6,pady=8); tk.Entry(edit,textvariable=serie_var,width=16).grid(row=1,column=1,sticky="w",padx=6,pady=8)
            tk.Label(edit,text="Invoice",bg=CARD).grid(row=1,column=2,sticky="e",padx=6,pady=8); tk.Entry(edit,textvariable=invoice_var,width=28).grid(row=1,column=3,sticky="w",padx=6,pady=8)
            tk.Label(edit,text="IGV",bg=CARD).grid(row=1,column=4,sticky="e",padx=6,pady=8); tk.Entry(edit,textvariable=igv_var,width=14).grid(row=1,column=5,sticky="w",padx=6,pady=8)

            rec_map={}
            def populate():
                tree.delete(*tree.get_children()); rec_map.clear()
                for i,r in enumerate([x for x in res.detail_rows if x.issues],1):
                    iid=f"obs_{i}"; rec_map[iid]=r
                    tree.insert("","end",iid=iid,values=(r.excel_row,r.proveedor,r.serie,r.numero,money2(r.igv)," | ".join(r.issues)))
                if not rec_map:
                    selected_row.set("Todos los observados fueron corregidos.")

            def select_rec(_event=None):
                sel=tree.selection()
                if not sel: return
                r=rec_map.get(sel[0])
                if not r: return
                selected_row.set(f"Editando fila Excel {r.excel_row} — {r.proveedor}")
                serie_var.set(r.serie); invoice_var.set(r.numero); igv_var.set(money2(r.igv))

            def save_change():
                sel=tree.selection()
                if not sel:
                    messagebox.showwarning("Seleccione un registro", "Seleccione una fila observada para corregirla.", parent=win); return
                r=rec_map.get(sel[0])
                try:
                    new_igv=Decimal(igv_var.get().replace(",","."))
                except Exception:
                    messagebox.showerror("IGV inválido", "Ingrese un IGV numérico válido.", parent=win); return
                r.serie=serie_var.get().strip(); r.numero=invoice_var.get().strip(); r.igv=new_igv
                refresh_result_files(res)
                # Actualiza los resúmenes globales y reportes ejecutivos.
                self.last_summary_xlsx = export_summary_xlsx(self.last_results, self.output_var.get().strip())
                self.last_summary_pdf = export_summary_pdf(self.last_results, self.output_var.get().strip())
                self._render_results(self.last_results)
                populate()
                messagebox.showinfo("Revalidación completada", f"Ahora: {res.valid} válidos y {res.observed} observados.\nSe regeneraron ambos TXT y el ZIP oficial si corresponde.", parent=win)

            tree.bind("<<TreeviewSelect>>",select_rec)
            buttons=tk.Frame(win,bg=BG); buttons.pack(fill="x",padx=16,pady=(0,14))
            tk.Button(buttons,text="Guardar corrección y revalidar",command=save_change,bg=SUCCESS,fg="white",relief="flat",padx=12,pady=7).pack(side="left")
            tk.Button(buttons,text="Previsualizar TXT actualizados",command=self.preview_selected_txt,bg=ACCENT2,fg="white",relief="flat",padx=12,pady=7).pack(side="left",padx=(8,0))
            populate()

        def open_summary_xlsx(self):
            if self.last_summary_xlsx and Path(self.last_summary_xlsx).exists():
                open_path(self.last_summary_xlsx)
            else:
                messagebox.showinfo("Sin resumen", "Procese primero uno o varios archivos para generar el resumen Excel.")

        def open_summary_pdf(self):
            if self.last_summary_pdf and Path(self.last_summary_pdf).exists():
                open_path(self.last_summary_pdf)
            else:
                messagebox.showinfo("Sin resumen", "Procese primero uno o varios archivos para generar el resumen PDF.")

else:
    App = None


def cli_main(argv: List[str]) -> int:
    import argparse
    p = argparse.ArgumentParser(description="Generador de carga masiva ND")
    p.add_argument("xlsx", nargs="+")
    p.add_argument("--salida", default=".")
    p.add_argument("--ruc", default="")
    p.add_argument("--periodo", default="")
    p.add_argument("--correlativo", type=int, default=1)
    p.add_argument("--igv", choices=["registered", "calculated"], default="registered")
    p.add_argument("--tasa", default="18")
    args = p.parse_args(argv)
    results = process_files(args.xlsx, args.salida, args.correlativo, args.igv, Decimal(args.tasa), True, args.ruc, args.periodo)
    xlsx = export_summary_xlsx(results, args.salida)
    pdf = export_summary_pdf(results, args.salida)
    for r in results:
        print(f"[{r.status}] {r.source_name} | RUC={r.ruc} | PERIODO={r.period} | ND={r.found} | VALIDOS={r.valid} | OBS={r.observed}")
        if r.messages:
            print("  ", " | ".join(r.messages))
        if r.zip_path:
            print("  ZIP:", r.zip_path)
        print("  TXT VALIDOS:", r.valid_txt_path)
        print("  TXT OBSERVADOS:", r.observed_txt_path)
    print("RESUMEN XLSX:", xlsx)
    print("RESUMEN PDF:", pdf)
    return 0


if __name__ == "__main__":
    if len(sys.argv) > 1:
        raise SystemExit(cli_main(sys.argv[1:]))
    if App is None:
        print("Tkinter no está disponible. Use el modo CLI.")
        raise SystemExit(1)
    App().mainloop()
