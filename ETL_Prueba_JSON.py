# -*- coding: utf-8 -*-

import json
import os
import re
import sys
import unicodedata
from datetime import datetime
from io import BytesIO
from typing import Dict, Optional, List

import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.units import pixels_to_EMU


# =========================================================
# CONFIG
# =========================================================

AUTO_ROW_HEIGHT = False
AUTO_ROW_MIN_HEIGHT = 34
AUTO_ROW_MAX_HEIGHT = 260
AUTO_LINE_HEIGHT = 23

EXCEL_WEB_MODE = True
EXCEL_WEB_MIN_ROW_HEIGHT = 38


# =========================================================
# Heurísticas
# =========================================================

TITLECASE_HINTS = (
    "estado", "municip", "colonia", "localizacion", "localización",
    "direccion", "dirección", "compania", "compañia", "nombre",
    "paterno", "materno", "puesto", "del_cd_mun", "cd_mun", "oficina",
    "region", "región"
)

SENTENCECASE_HINTS = (
    "proyecto", "descripcion", "descripción", "observaciones", "acabados"
)

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
CODE_RE  = re.compile(r"^[A-Z]{1,6}\d{2,}$")

CITY_FIXES = {
    "ciudad de mexico": "Ciudad de México",
    "estado de mexico": "Estado de México",
    "nuevo leon": "Nuevo León",
    "san luis potosi": "San Luis Potosí",
    "michoacan": "Michoacán",
    "queretaro": "Querétaro",
    "yucatan": "Yucatán",
    "aguascalientes": "Aguascalientes",
    "baja california": "Baja California",
    "baja california sur": "Baja California Sur",
    "campeche": "Campeche",
    "chihuahua": "Chihuahua",
    "coahuila": "Coahuila",
    "colima": "Colima",
    "durango": "Durango",
    "guanajuato": "Guanajuato",
    "guerrero": "Guerrero",
    "hidalgo": "Hidalgo",
    "jalisco": "Jalisco",
    "morelos": "Morelos",
    "nayarit": "Nayarit",
    "oaxaca": "Oaxaca",
    "puebla": "Puebla",
    "quintana roo": "Quintana Roo",
    "sinaloa": "Sinaloa",
    "sonora": "Sonora",
    "tabasco": "Tabasco",
    "tamaulipas": "Tamaulipas",
    "tlaxcala": "Tlaxcala",
    "veracruz": "Veracruz",
    "zacatecas": "Zacatecas",
}


COLOR_ORANGE = "ED7D31"
COLOR_GRAY = "E7E6E6"

_GROUP_FILLS = {
    "Información General": PatternFill("solid", COLOR_ORANGE),
    "Ubicación del Proyecto": PatternFill("solid", COLOR_GRAY),
    "Caracteristicas del Proyecto": PatternFill("solid", COLOR_ORANGE),
    "Datos de la Compañia": PatternFill("solid", COLOR_GRAY),
    "Participantes y Contactos": PatternFill("solid", COLOR_ORANGE),
    "Detalles Adicionales": PatternFill("solid", COLOR_GRAY),
    "Otros": PatternFill("solid", COLOR_GRAY),
}

_GROUP_FONT = Font(name="Poppins", size=11, bold=True, color="000000")
_GROUP_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)


# =========================================================
# ✅ SOLUCIÓN DE CARACTERES: mojibake repair + unescape
# =========================================================

_REPL_CHAR = "�"

def _score_spanish(s: str) -> int:
    if not isinstance(s, str) or not s:
        return -10
    good = "áéíóúüñÁÉÍÓÚÜÑ"
    bad  = "ÃÂ�µ¥¢¤ðÐþÞ¿¡"
    score = 0
    score += sum(2 for ch in s if ch in good)
    score -= sum(2 for ch in s if ch in bad)
    score -= sum(3 for ch in s if ord(ch) < 32 and ch not in "\n\t\r")
    return score

def _try_redecode(s: str, from_enc: str, to_enc: str) -> Optional[str]:
    try:
        b = s.encode(from_enc)
        return b.decode(to_enc)
    except Exception:
        return None

def _repair_mojibake(s: str) -> str:
    """
    Repara mojibake común ES:
    - UTF-8 leído como latin1/cp1252 => 'MÃ©xico'
    - CP850/CP437 vs latin1 => 'Ni¥os', 'µREA'
    - ✅ NUEVO: C1 controls (U+0080–U+009F) típicos cuando CP850 se interpretó como texto “plano”
      Ej: 'inter\x90s' donde 0x90 (CP850) = 'É'
    """
    if not isinstance(s, str) or not s:
        return s

    # Detecta C1 controls (U+0080..U+009F)
    has_c1 = any(0x80 <= ord(ch) <= 0x9F for ch in s)

    symptoms = ("Ã", "Â", "�", "µ", "¥", "¤", "¢", "à", "¨")
    if (not has_c1) and (not any(x in s for x in symptoms)):
        return s

    candidates = [s]

    def try_redecode_latin1_to(target_enc: str):
        try:
            b = s.encode("latin1", errors="strict")   # latin1 preserva 0x00..0xFF tal cual
            return b.decode(target_enc, errors="strict")
        except Exception:
            return None

    def try_redecode(from_enc: str, to_enc: str):
        try:
            b = s.encode(from_enc)
            return b.decode(to_enc)
        except Exception:
            return None

    # Caso clásico: UTF-8 mal leído como Latin1/CP1252 => 'MÃ©xico'
    for enc in ("latin1", "cp1252"):
        cand = try_redecode(enc, "utf-8")
        if cand:
            candidates.append(cand)

    # Caso CP850/CP437 mal tratado como texto => aparecen ¥, µ, o C1 controls (\x90 etc)
    for target in ("cp850", "cp437"):
        cand = try_redecode_latin1_to(target)
        if cand:
            candidates.append(cand)

    # Tu ruta anterior (cp1252 <-> cp850/cp437)
    for target in ("cp850", "cp437"):
        cand = try_redecode("cp1252", target)
        if cand:
            candidates.append(cand)

    for src in ("cp850", "cp437"):
        cand = try_redecode(src, "cp1252")
        if cand:
            candidates.append(cand)

    # Scoring “español” (reusa tu _score_spanish si ya existe)
    def score_spanish(x: str) -> int:
        good = "áéíóúüñÁÉÍÓÚÜÑ"
        bad  = "ÃÂ�µ¥¢¤ðÐþÞ¿¡"
        sc = 0
        sc += sum(2 for ch in x if ch in good)
        sc -= sum(2 for ch in x if ch in bad)
        # ✅ penaliza C1 controls
        sc -= sum(4 for ch in x if 0x80 <= ord(ch) <= 0x9F)
        return sc

    best = max(candidates, key=score_spanish)

    # Si aún quedan C1 controls, al menos elimínalos para que no “corten” el texto
    best = "".join(ch for ch in best if not (0x80 <= ord(ch) <= 0x9F))

    return best
def _unescape_quotes_backslashes(s: str) -> str:

    if not isinstance(s, str):
        return s

    # quitar escapes comunes
    s = s.replace(r'\\"', '"')
    s = s.replace(r"\"", '"')

    # quitar backslashes sobrantes
    s = s.replace("\\", "")

    # limpiar saltos de línea escapados
    s = s.replace(r"\n", "\n").replace(r"\t", "\t")

    # quitar comillas envolventes
    s = s.strip()

    if len(s) >= 2 and s.startswith('"') and s.endswith('"'):
        s = s[1:-1]

    return s

def _repair_all_strings_df(df: pd.DataFrame) -> pd.DataFrame:

    # caracteres típicos de mojibake
    pattern = re.compile(r"[ÃÂ�µ¥¤¢à¨]")

    def fix(v):
        if isinstance(v, str):
            v = _repair_mojibake(v)
            v = _unescape_quotes_backslashes(v)
        return v

    text_columns = list(df.select_dtypes(include="object").columns)

    for col in text_columns:

        series = df[col]

        # normalizar nombre de columna
        norm = _norm_colkey(col)

        # columnas específicas donde sabemos que vienen \"texto\"
        ESCAPE_TARGETS = {"nombre", "nombre_del_proyecto", "proyecto"}

        # si la columna es una de las problemáticas, limpiar directamente
        if norm in ESCAPE_TARGETS:
            df[col] = series.map(
                lambda v: _unescape_quotes_backslashes(_repair_mojibake(v)) if isinstance(v, str) else v
            )
            continue

        # para el resto solo intentar reparar mojibake si hay síntomas
        try:
            if not series.astype(str).str.contains(pattern).any():
                continue
        except Exception:
            continue

        df[col] = series.map(lambda v: _repair_mojibake(v) if isinstance(v, str) else v)

    return df


# =========================================================
# Helpers texto
# =========================================================

def _norm_noaccents_lower(s: str) -> str:
    s = (s or "").strip().lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"\s+", " ", s).strip()
    return s

def _strip_wrapping_quotes(s: str) -> str:
    if len(s) >= 2 and s[0] == '"' and s[-1] == '"':
        return s[1:-1]
    return s

def _sentence_case_spanish(s: str) -> str:
    s = s.lower()
    out = []
    cap_next = True
    for ch in s:
        if cap_next and ch.isalpha():
            out.append(ch.upper())
            cap_next = False
        else:
            out.append(ch)
        if ch in ".!?":
            cap_next = True
        if ch == "\n":
            cap_next = True
    return "".join(out)

def _title_case_spanish(s: str) -> str:
    lower_words = {"de","del","la","las","el","los","y","e","o","u","a","al","en","con","por","para","sin"}
    parts = s.split()
    out = []
    for i, w in enumerate(parts):
        wl = w.lower()
        if not any(ch.isalpha() for ch in w):
            out.append(w); continue
        if i > 0 and wl in lower_words:
            out.append(wl)
        else:
            out.append(wl[:1].upper() + wl[1:])
    return " ".join(out)

def _choose_case_strategy(col_name: str) -> str:
    c = (col_name or "").lower()
    if any(h in c for h in SENTENCECASE_HINTS):
        return "sentence"
    if any(h in c for h in TITLECASE_HINTS):
        return "title"
    return "none"

def _uppercase_inside_quotes(text: str) -> str:
    if not isinstance(text, str) or '"' not in text:
        return text
    return re.sub(r'"([^"]+)"', lambda m: '"' + m.group(1).upper() + '"', text)

def _fix_shouty_caps_mixed(text: str) -> str:
    if not isinstance(text, str):
        return text
    keep_upper = {"C.P", "C.P.", "ID", "PDF", "MX", "M2", "M²", "SAT", "IMEI", "NSS"}

    def fix_word(w: str) -> str:
        letters = [ch for ch in w if ch.isalpha()]
        if not letters:
            return w
        if w.replace(":", "").replace(".", "").upper() in {k.replace(".", "") for k in keep_upper}:
            return w
        if len([ch for ch in w if ch.isalpha()]) >= 3 and all(ch.isupper() for ch in letters):
            low = w.lower()
            return low[:1].upper() + low[1:]
        return w

    tokens = re.split(r'(\s+)', text)
    out = []
    for t in tokens:
        if t.isspace() or t == "":
            out.append(t); continue
        subtoks = re.split(r'([,:;.\-()])', t)
        fixed = []
        for st in subtoks:
            if st in {",", ":", ";", ".", "-", "(", ")"} or st == "":
                fixed.append(st)
            else:
                fixed.append(fix_word(st))
        out.append("".join(fixed))
    return "".join(out)

def _normalize_free_text(col: str, s: str, force_upper: bool = False) -> str:
    if not isinstance(s, str):
        return s
    s = _fix_shouty_caps_mixed(s)
    if force_upper:
        return s.upper()

    col_l = (col or "").lower()
    if col_l == "localizacion1":
        return _title_case_spanish(s)

    s = _sentence_case_spanish(s)
    if col_l == "proyecto":
        s = _uppercase_inside_quotes(s)
    return s

def _fix_clasico_observaciones_codes(s: str) -> str:
    if not isinstance(s, str):
        return s

    # Detect patterns like oc123, Oc123, pp123, etc. and normalize to uppercase prefix
    return re.sub(
        r"\b([a-zA-Z]{2})(\d{3,})\b",
        lambda m: m.group(1).upper() + m.group(2),
        s
    )

def _smart_text_format(v, col_name: str):
    if not isinstance(v, str):
        return v

    s = v.strip()
    s = _strip_wrapping_quotes(s)
    if not s:
        return None

    k = _norm_noaccents_lower(s)
    if k in CITY_FIXES:
        return CITY_FIXES[k]

    if EMAIL_RE.match(s):
        return s.lower()

    if CODE_RE.match(s):
        return s

    if " " not in s and any(ch.isdigit() for ch in s) and any(ch.isalpha() for ch in s):
        return s

    letters = [ch for ch in s if ch.isalpha()]
    all_caps = bool(letters) and all(ch.isupper() for ch in letters)
    if all_caps:
        strat = _choose_case_strategy(col_name)
        if strat == "sentence":
            return _sentence_case_spanish(s)
        if strat == "title":
            return _title_case_spanish(s)
        return _title_case_spanish(s)

    return s


# =========================================================
# Helpers DF / Excel
# =========================================================

def _export_headers_with_spaces(cols):
    return [str(c).replace("_", " ").strip() for c in cols]

def _compute_widths_from_df(df: pd.DataFrame, padding: int = 4, max_width: int = 60) -> dict:
    if df.empty:
        return {}
    sample = df if len(df) <= 2000 else df.head(2000)
    widths = {}
    for col in sample.columns:
        max_len = len(str(col))
        for v in sample[col]:
            if v is None:
                continue
            max_len = max(max_len, len(str(v)))
        widths[col] = min(max_len + padding, max_width)
    return widths

def _norm_colkey(name: str) -> str:
    if name is None:
        return ""
    s = str(name).strip().lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.replace(" ", "_").replace("-", "_")
    s = re.sub(r"_+", "_", s)
    return s

def _apply_width_overrides(ws, df_export: pd.DataFrame):
    TARGET_WIDTHS = {
        "email_1": 42,
        "email_2": 42,
        "email_3": 55,
        "proyecto": 95,
        "descripcion": 120,
        "descripcionextra": 120,
        "observaciones": 95,
        "descripcion_extra": 120,
        "localizacion": 85,
        "localizacion1": 85,
        "localizacion_del_proyecto": 95,
        "acabados": 120,
    }

    for idx, col_name in enumerate(df_export.columns, start=1):
        k = _norm_colkey(col_name)
        if k == "email3": k = "email_3"
        if k == "email2": k = "email_2"
        if k == "email1": k = "email_1"
        if k in ("descripcion_extra_del_proyecto",): k = "descripcion_extra"
        if k in TARGET_WIDTHS:
            ws.column_dimensions[get_column_letter(idx)].width = TARGET_WIDTHS[k]

def _resolve_resource_path(filename: str) -> Optional[str]:
    base_path = getattr(sys, "_MEIPASS", None)
    if base_path:
        candidate = os.path.join(base_path, filename)
        if os.path.exists(candidate):
            return candidate
    candidate = os.path.join(os.path.dirname(__file__), filename)
    if os.path.exists(candidate):
        return candidate
    candidate = os.path.join(os.getcwd(), filename)
    if os.path.exists(candidate):
        return candidate
    return None


# =========================================================
# Estilos / formatos
# =========================================================

def _apply_styles_excel_and_sheets(ws, header_row: int, first_data_row: int, nrows: int, ncols: int, orig_headers=None):
    header_font = Font(name="Poppins", size=11, bold=True, color="FFFFFF")
    body_font = Font(name="Poppins", size=11)

    header_fill = PatternFill("solid", COLOR_ORANGE)
    zebra_a = PatternFill("solid", "F2F2F2")
    zebra_b = PatternFill("solid", "FFFFFF")

    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    body_align_center = Alignment(wrap_text=True, vertical="center", horizontal="center")
    body_align_left = Alignment(wrap_text=True, vertical="center", horizontal="left")

    LEFT_ALIGN_COLUMNS = {
        "proyecto",
        "localizacion1",
        "descripcion",
        "acabados",
        "observaciones"
    }

    ws.freeze_panes = f"A{first_data_row}"
    ws.sheet_view.showGridLines = False
    ws.auto_filter.ref = f"A{header_row}:{get_column_letter(ncols)}{nrows}"

    # IMPORTANT: For CLASICO reports the template already contains the correct
    # header styling and colors. We must NEVER override them.
    for cell in ws[header_row]:
        cell.alignment = header_align
        cell.fill = header_fill
        cell.font = header_font
    else:
        # Non-clasico reports still apply programmatic styling
        for cell in ws[header_row]:
            cell.alignment = header_align
            cell.fill = header_fill
            cell.font = header_font

    for r in range(first_data_row, nrows + 1):

        fill = zebra_a if (r % 2 == 0) else zebra_b

        for c in range(1, ncols + 1):

            cell = ws.cell(row=r, column=c)
            header_name = str(ws.cell(row=header_row, column=c).value)

            norm = header_name.lower().replace(" ", "_")

            cell.font = body_font
            cell.fill = fill

            if norm in LEFT_ALIGN_COLUMNS:
                # Special handling for long descriptive fields like 'acabados'
                if norm == "acabados":
                    cell.alignment = Alignment(
                        wrap_text=True,
                        vertical="top",
                        horizontal="left"
                    )
                else:
                    cell.alignment = body_align_left
            else:
                cell.alignment = body_align_center


def _apply_row_borders(ws, first_data_row: int, nrows: int, ncols: int):

    row_side = Side(style="thin", color="b0b0b0")

    for row in ws.iter_rows(
        min_row=first_data_row,
        max_row=nrows,
        min_col=1,
        max_col=ncols
    ):

        for cell in row:

            existing = cell.border

            cell.border = Border(
                left=existing.left,
                right=existing.right,
                top=existing.top,
                bottom=row_side
            )

def _apply_fixed_row_height(ws, first_data_row: int, nrows: int, height: int = 38):

    for r in range(first_data_row, nrows + 1):
        ws.row_dimensions[r].height = height

def _format_date_columns_no_time(ws, df_orig: pd.DataFrame, first_data_row: int):
    date_cols_idx = [i for i, col in enumerate(df_orig.columns, start=1) if "fecha" in str(col).lower()]
    if not date_cols_idx:
        return
    date_fmt = "yyyy-mm-dd"
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    for r in range(first_data_row, ws.max_row + 1):
        for c in date_cols_idx:
            cell = ws.cell(row=r, column=c)
            if isinstance(cell.value, datetime):
                cell.number_format = date_fmt
                cell.alignment = left_align

def _format_numeric_columns(ws, df_orig: pd.DataFrame, first_data_row: int):
    formats = {}
    for i, col in enumerate(df_orig.columns, start=1):
        cname = _norm_noaccents_lower(str(col)).replace(" ", "_")

        if "inversion" in cname:
            formats[i] = '$#,##0'
        elif col in ("Sup_Construida", "Sup_Urbanizada"):
            formats[i] = '#,##0'
        elif "sup_" in cname or "superficie" in cname:
            formats[i] = '#,##0'
        elif cname in ("numero_unidades", "num_niveles"):
            formats[i] = '#,##0'
        elif cname in ("latitud", "longitud", "latitude", "longitude", "lat", "lng"):
            formats[i] = '0.000000'
        elif cname in ("dia_publicado", "mes_publicado", "ano_publicado", "anio_publicado",
                       "dia_inicio", "mes_inicio", "ano_inicio", "anio_inicio"):
            formats[i] = '0'
        else:
            if col in df_orig.columns and pd.api.types.is_numeric_dtype(df_orig[col]):
                formats[i] = '#,##0'

    for r in range(first_data_row, ws.max_row + 1):
        for cidx, fmt in formats.items():
            cell = ws.cell(row=r, column=cidx)
            if isinstance(cell.value, (int, float)) and cell.value is not None:
                cell.number_format = fmt


# =========================================================
# Branding / Clasificación
# =========================================================

def _apply_branding_row(
    ws,
    ncols: int,
    empresa: str,
    usuario: str,
    report_label: str,
    logo_filename: str = "logo_bimsa.jpg",
    logo_path: Optional[str] = None,
):
    ws.row_dimensions[1].height = 50
    ws.column_dimensions["A"].width = 28

    white = PatternFill("solid", "FFFFFF")
    bottom_side = Side(style="thin", color="000000")

    for c in range(1, ncols + 1):
        cell = ws.cell(1, c)
        cell.fill = white

        existing = cell.border

        cell.border = Border(
            left=existing.left,
            right=existing.right,
            top=existing.top,
            bottom=bottom_side
        )

    # Only insert a logo if an explicit logo_path is provided.
    # If the template already contains a logo, we leave it untouched.
    if logo_path and os.path.exists(logo_path):

        img = XLImage(logo_path)
        img.width = 170
        img.height = 40

        # Center image inside cell A1
        cell_height_emu = pixels_to_EMU(55)  # row height ~80px
        img_height_emu = pixels_to_EMU(img.height)

        vertical_offset = int((cell_height_emu - img_height_emu) / 2)
        horizontal_offset = pixels_to_EMU(10)  # small left padding


        marker = AnchorMarker(
            colOff=horizontal_offset,
            rowOff=vertical_offset
        )

        img.anchor = OneCellAnchor(
            _from=marker,
            ext=XDRPositiveSize2D(
                pixels_to_EMU(img.width),
                pixels_to_EMU(img.height)
            )
        )

        ws.add_image(img)

    ws.merge_cells(start_row=1, start_column=2, end_row=1, end_column=4)

    info_cell = ws.cell(1, 2)
    info_lines = [
        f"Empresa: {empresa}",
        f"Usuario: {usuario}",
    ]

    info_cell.alignment = Alignment(
        horizontal="left",
        vertical="center",
        wrap_text=True
    )

    info_cell.value = "\n".join(info_lines)
    info_cell.font = Font(name="Poppins", size=11, bold=False, color="000000")
    info_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    right_col = ncols
    ws.merge_cells(start_row=1, start_column=max(1, right_col - 1), end_row=1, end_column=right_col)
    tag_cell = ws.cell(1, max(1, right_col - 1))
    tag_cell.value = report_label
    tag_cell.font = Font(name="Poppins", size=14, bold=True, color="FFFFFF")
    tag_cell.fill = PatternFill("solid", "ED7D31")
    tag_cell.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)

    for c in range(1, ncols + 1):
        ws.cell(1, c).fill = white

# =========================================================
# Auto-height (wrap por palabras)
# =========================================================

def _safe_col_width(ws, col_letter: str) -> float:
    w = ws.column_dimensions[col_letter].width
    return float(w) if (w is not None and w > 0) else 12.0

def _estimate_wrapped_lines(text: str, col_width_chars: float) -> int:
    if not text:
        return 1
    SHEETS_WIDTH_FACTOR = 0.88
    effective = max(8, int(col_width_chars * SHEETS_WIDTH_FACTOR) - 1)

    total_lines = 0
    for raw_line in str(text).split("\n"):
        raw_line = raw_line.strip()
        if not raw_line:
            total_lines += 1
            continue

        words = raw_line.split()
        line_len = 0
        lines = 1

        for w in words:
            wlen = len(w)
            if line_len == 0:
                line_len = wlen
            elif line_len + 1 + wlen <= effective:
                line_len += 1 + wlen
            else:
                lines += 1
                line_len = wlen

        total_lines += lines

    return max(1, total_lines)

def _apply_auto_row_heights(ws, first_data_row: int, min_h: int, max_h: int, line_h: int):
    for r in range(first_data_row, ws.max_row + 1):
        max_lines = 1
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            if cell.value is None:
                continue
            col_letter = get_column_letter(c)
            col_w = _safe_col_width(ws, col_letter)
            lines = _estimate_wrapped_lines(cell.value, col_w)
            max_lines = max(max_lines, lines)

        ws.row_dimensions[r].height = max(min_h, min(max_h, int(max_lines * line_h)))


# =========================================================
# ETL principal
# =========================================================

def ETL_BIMSA(
    ruta_json: str,
    tipo_reporte: str,
    return_mode: str = "file",
    carpeta_excel: str = ".",
    empresa: str = "",
    usuario: str = "",
    tipo_fecha: Optional[str] = None,
    fecha_inicio: Optional[str] = None,
    fecha_fin: Optional[str] = None,
    report_label: Optional[str] = None,
    logo_path: Optional[str] = None,
):
    print("[BIMSA_ETL] Iniciando ETL BIMSA...")

    tipo_upper = str(tipo_reporte).strip().upper()
    es_clasico = False

    now = datetime.now()
    nombre_excel = f"BIMSA_{tipo_upper}_{now.strftime('%Y%m%d_%H%M%S')}.xlsx"

    if report_label is None:
        report_label = tipo_upper.lower()

    with open(ruta_json, "r", encoding="utf-8") as f:
        data = json.load(f)

    if not isinstance(data, list) or not data:
        raise ValueError("El JSON debe ser una lista con al menos un registro")

    df = pd.DataFrame(data)

    # ---------------------------------------------------------
    # Normalizar columnas que vienen del backend como "Anio_*"
    # (ej: Anio_inicio, Anio_publicado) para evitar problemas
    # con plantillas que esperan "Ano_*" o variantes similares
    # ---------------------------------------------------------
    df = df.rename(columns=lambda c: (
        c.replace("Anio_", "Año_")
         .replace("anio_", "año_")
    ) if isinstance(c, str) else c)

    # ✅ FIX GLOBAL DE CARACTERES (a TODO texto)
    df = _repair_all_strings_df(df)

    # Columnas catálogo (no deben procesarse)
    CATALOGO_COLUMNS = {
        "tipo_proyecto",
        "etapa",
        "tipo_desarrollo",
        "region",
        "estado_proyecto",
        "Nombre_del_Proyecto",
        "nombre_del_proyecto",
        "estado",
        "genero",
        "subgenero",
        "tipo_obra",
        "sector",
        "rol_compania",
        "estado_compania",
        "puesto",
        "puesto_1",
        "puesto_2",
        "puesto_3",
        "clave_tipo_obra",
    }

    # Precompute text columns once (avoids repeated dtype scanning)
    text_columns = df.select_dtypes(include="object").columns

    for col in text_columns:

        norm_col = _norm_colkey(col)

        if norm_col in CATALOGO_COLUMNS:
            continue

        series = df[col]

        # Skip empty columns quickly
        if series.isna().all():
            continue

        # Only process non-null values
        mask = series.notna()

        df.loc[mask, col] = series[mask].map(
            lambda v, c=col: _smart_text_format(v, c) if isinstance(v, str) else v
        )

    # Reglas por tipo
    force_upper_proyecto_mapas = (tipo_upper == "MAPAS")

    ALIAS_NOMBRE_PROY = {"nombre_del_proyecto"}
    ALIAS_DESC_PROY   = {"descripcion_del_proyecto"}
    ALIAS_PROYECTO    = {"proyecto", "nombre"}
    ALIAS_LOCALIZ     = {"localizacion1", "localizacion", "localizacion_del_proyecto"}
    ALIAS_OBSERV      = {"observaciones"}
    ALIAS_DESC_EXTRA  = {"descripcion_extra", "descripcionextra", "descripcion_extra_del_proyecto"}

    for col in list(df.columns):
        nk2 = _norm_noaccents_lower(col).replace(" ", "_").replace("__", "_")

        if tipo_upper == "CONTACTOS":
            if nk2 in ALIAS_NOMBRE_PROY:
                df[col] = df[col].map(lambda x: _normalize_free_text("Proyecto", x, force_upper=True))
                continue
            if nk2 in ALIAS_DESC_PROY:
                df[col] = df[col].map(lambda x: _normalize_free_text("Proyecto", x, force_upper=False))
                df[col] = df[col].map(_uppercase_inside_quotes)
                continue

        if tipo_upper == "MAPAS" and nk2 in (ALIAS_NOMBRE_PROY | ALIAS_DESC_PROY | ALIAS_PROYECTO):
            df[col] = df[col].map(lambda x: _normalize_free_text("Proyecto", x, force_upper=True))
            continue

        if nk2 in ALIAS_NOMBRE_PROY:
            df[col] = df[col].map(lambda x: x.upper() if isinstance(x, str) else x)
            continue

        if nk2 in (ALIAS_DESC_PROY | ALIAS_PROYECTO):
            df[col] = df[col].map(lambda x: _normalize_free_text("Proyecto", x, force_upper=False))
            df[col] = df[col].map(_uppercase_inside_quotes)

        elif nk2 in ALIAS_LOCALIZ:
            df[col] = df[col].map(lambda x: _normalize_free_text("Localizacion1", x, force_upper=False))

        elif nk2 in ALIAS_OBSERV:
            df[col] = df[col].map(lambda x: _normalize_free_text("Observaciones", x, force_upper=False))
            df[col] = df[col].map(_fix_clasico_observaciones_codes)

        elif nk2 in ALIAS_DESC_EXTRA:
            df[col] = df[col].map(lambda x: _normalize_free_text("Descripcion_Extra", x, force_upper=False))

    # Descripcion / Acabados
    for c in ("Descripcion", "Acabados"):
        if c in df.columns:
            df[c] = df[c].map(_fix_shouty_caps_mixed)
            df[c] = df[c].map(lambda x: _sentence_case_spanish(x) if isinstance(x, str) else x)

    # Fechas
    for col in df.columns:
        if "fecha" in str(col).lower():
            s = df[col]
            has_slash = False
            try:
                sample = s.dropna().astype(str).head(50)
                has_slash = any("/" in v for v in sample)
            except Exception:
                pass
            df[col] = pd.to_datetime(s, errors="coerce", dayfirst=has_slash)

    # Numéricos generales
    for col in df.columns:
        cname = str(col).lower()
        if "inversion" in cname:
            df[col] = pd.to_numeric(df[col], errors="coerce")
        if "sup_" in cname or "superficie" in cname:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    for c in ("Numero_Unidades", "Num_Niveles"):
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    # MAPAS: lat/long + dia/mes/año numéricos
    if tipo_upper == "MAPAS":
        def _k(name: str) -> str:
            return _norm_noaccents_lower(name).replace(" ", "_").replace("__", "_")
        for col in df.columns:
            kk = _k(col)
            if kk in ("latitud", "longitud", "latitude", "longitude", "lat", "lng"):
                df[col] = pd.to_numeric(df[col], errors="coerce")
            if kk in (
                "dia_publicado", "mes_publicado", "año_publicado", "Anio_publicado", "Mes_inicio", "Año_inicio"
                "dia_inicio", "mes_inicio", "anio_inicio", "Anio_inicio", "Celular", "Extension", "Lada"
            ):
                # Convertir a número entero REAL (no string)
                df[col] = pd.to_numeric(df[col], errors="coerce")


    df_export = df.copy()

    df_export.columns = _export_headers_with_spaces(df_export.columns)

    # OUTPUT
    if return_mode == "bytes":
        output = BytesIO()
        writer_target = output
        ruta_final = None
    else:
        os.makedirs(carpeta_excel, exist_ok=True)
        ruta_final = f"{carpeta_excel.rstrip('/')}/{nombre_excel}"
        writer_target = ruta_final

    # =========================================================
    # GENERACIÓN DIRECTA DE EXCEL (SIN PLANTILLA)
    # =========================================================

    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Reporte"

    # Definir filas
    header_row = 2
    first_data_row = 3

    # Insertar headers dinámicos
    for col_idx, col_name in enumerate(df_export.columns, start=1):
        ws.cell(row=header_row, column=col_idx, value=str(col_name) if col_name else "")


    # Insertar datos
    data_matrix = df_export.to_numpy()

    for r_offset, row in enumerate(data_matrix):
        excel_row = first_data_row + r_offset
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=excel_row, column=c_idx, value=value)

    ncols = len(df_export.columns)
    nrows = first_data_row + len(df_export) - 1

    # Aplicar anchos de columna
    widths = _compute_widths_from_df(df_export, padding=4, max_width=60)

    MIN_HEADER_WIDTH = 18  # 👈 mínimo para evitar wrap en headers

    for idx, col_name in enumerate(df_export.columns, start=1):
        header_len = len(str(col_name))

        # ancho base calculado
        base_width = widths.get(col_name, 12)

        # asegurar que el header quepa en una línea
        header_width = header_len + 4

        final_width = max(base_width, header_width, MIN_HEADER_WIDTH)

        ws.column_dimensions[get_column_letter(idx)].width = final_width

    _apply_width_overrides(ws, df_export)

    # =========================================================
    # Aplicar estilos (tipografía + patrón de cebra)
    # =========================================================

    _apply_styles_excel_and_sheets(
        ws,
        header_row=header_row,
        first_data_row=first_data_row,
        nrows=nrows,
        ncols=ncols,
        orig_headers=None
    )

    # Borde inferior para headers (look más limpio tipo tabla)
    header_border = Border(bottom=Side(style="thin", color="000000"))

    for c in range(1, ncols + 1):
        cell = ws.cell(header_row, c)
        existing = cell.border
        cell.border = Border(
            left=existing.left,
            right=existing.right,
            top=existing.top,
            bottom=header_border.bottom
        )

    # (Opcional) Altura de header
    ws.row_dimensions[header_row].height = 38

    # =========================================================
    # Aplicar branding
    # =========================================================
    if not logo_path:
        logo_path = _resolve_resource_path("logo_bimsa.jpg")

    _apply_branding_row(
        ws,
        ncols=ncols,
        empresa=empresa,
        usuario=usuario,
        report_label=report_label,
        logo_filename="logo_bimsa.jpg",
        logo_path=logo_path,
    )

    # =========================================================
    # FORMATEO INTELIGENTE (fechas, dinero, texto largo)
    # =========================================================

    for col_idx, col_name in enumerate(df.columns, start=1):

        col_series = df[col_name]
        col_lower = str(col_name).lower()
        if any(k in col_lower for k in ["latitud", "longitud", "latitude", "longitude", "lat", "lng"]):
            for r in range(first_data_row, ws.max_row + 1):
                cell = ws.cell(row=r, column=col_idx)
                try:
                    val = float(cell.value)
                    cell.value = val
                    cell.number_format = '0.00000'
                except (TypeError, ValueError):
                    pass
            continue

        # --- FORZAR FORMATO NUMÉRICO PARA DIA/MES/AÑO ---
        elif any(k in col_lower for k in [
            "dia_publicado", "mes_publicado", "ano_publicado", "anio_publicado",
            "dia_inicio", "mes_inicio", "ano_inicio", "anio_inicio"
        ]):
            for r in range(first_data_row, ws.max_row + 1):
                cell = ws.cell(row=r, column=col_idx)
                try:
                    val = int(cell.value)
                    cell.value = val
                    cell.number_format = '0'
                except (TypeError, ValueError):
                    pass

        # --- DETECCIÓN DE FECHAS ---
        if "fecha" in col_lower or pd.api.types.is_datetime64_any_dtype(col_series):
            for r in range(first_data_row, ws.max_row + 1):
                cell = ws.cell(row=r, column=col_idx)
                if isinstance(cell.value, datetime):
                    cell.number_format = "yyyy-mm-dd"

        # --- DETECCIÓN DE DINERO ---
        elif "inversion" in col_lower or "precio" in col_lower or "monto" in col_lower:
            for r in range(first_data_row, ws.max_row + 1):
                cell = ws.cell(row=r, column=col_idx)
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '$#,##0'

        # --- DETECCIÓN NUMÉRICA GENERAL ---
        elif pd.api.types.is_numeric_dtype(col_series):
            for r in range(first_data_row, ws.max_row + 1):
                cell = ws.cell(row=r, column=col_idx)
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0'

        # --- TEXTO LARGO (wrap + left align más elegante) ---
        if any(k in col_lower for k in ["descripcion", "observaciones", "proyecto"]):
            for r in range(first_data_row, ws.max_row + 1):
                cell = ws.cell(row=r, column=col_idx)
                cell.alignment = Alignment(
                    wrap_text=True,
                    vertical="center",
                    horizontal="left"
                )

    # =========================================================
    # Altura de filas
    # =========================================================

    if AUTO_ROW_HEIGHT:
        min_h = AUTO_ROW_MIN_HEIGHT

        if EXCEL_WEB_MODE:
            min_h = max(min_h, EXCEL_WEB_MIN_ROW_HEIGHT)

        _apply_auto_row_heights(
            ws,
            first_data_row=first_data_row,
            min_h=min_h,
            max_h=AUTO_ROW_MAX_HEIGHT,
            line_h=AUTO_LINE_HEIGHT,
        )
    else:
        _apply_fixed_row_height(
            ws,
            first_data_row=first_data_row,
            nrows=nrows,
            height=60
        )

    # =========================================================
    # Guardar archivo
    # =========================================================

    if return_mode == "bytes":

        output = BytesIO()
        wb.save(output)
        output.seek(0)

        print(f"[BIMSA_ETL] ETL terminado correctamente: {nombre_excel}")

        return nombre_excel, output.getvalue()

    else:

        os.makedirs(carpeta_excel, exist_ok=True)

        ruta_final = f"{carpeta_excel.rstrip('/')}/{nombre_excel}"

        wb.save(ruta_final)

        print(f"[BIMSA_ETL] ETL terminado correctamente: {nombre_excel}")

        return ruta_final
