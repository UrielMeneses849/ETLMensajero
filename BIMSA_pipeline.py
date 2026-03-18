# BIMSA_pipeline.py
from __future__ import annotations

from pathlib import Path
from datetime import datetime
from typing import Tuple, Optional, Dict, Any

from ETL_XML_to_JSON import xml_a_json
from ETL_Prueba_JSON import ETL_BIMSA


def _safe_token(s: str) -> str:
    s = (s or "").strip().replace(" ", "_")
    return "".join(ch for ch in s if ch.isalnum() or ch in ("_", "-")).strip("_") or "REPORTE"


def _sanitize_etl_opts(etl_opts: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Evita romper ETL_BIMSA si el backend manda flags que tu ETL no soporta.
    """
    opts = dict(etl_opts or {})

    # ⛔️ Flags que NO deben llegar a ETL_BIMSA (causan errores como audit_dir)
    for bad_key in (
        "audit_dir",
        "enable_encoding_audit",
        "encoding_audit",
        "encoding_audit_dir",
        "audit",
    ):
        opts.pop(bad_key, None)

    return opts


def process_bimsa(
    xml_string: str,
    tipo_reporte: str,
    *,
    output_root: str = "bimsa_runs",
    guardar_json: bool = True,
    guardar_excel_en_disco: bool = False,
    etl_opts: Optional[Dict[str, Any]] = None,
) -> Tuple[str, bytes]:
    """
    Punto de entrada único (API) para el backend de BIMSA:

    - Recibe XML (string)
    - Guarda JSON internamente (opcional)
    - Genera Excel en memoria
    - Retorna (nombre_excel, excel_bytes)
    """

    tipo = _safe_token(tipo_reporte).upper()
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    run_dir = Path(output_root) / f"{ts}_{tipo.lower()}"

    json_dir = run_dir / "json"
    excel_dir = run_dir / "excel"

    if guardar_json:
        json_dir.mkdir(parents=True, exist_ok=True)
    if guardar_excel_en_disco:
        excel_dir.mkdir(parents=True, exist_ok=True)

    # 1) XML -> JSON
    if guardar_json:
        ruta_json = xml_a_json(
            xml_string,
            tipo,
            carpeta_salida=str(json_dir),
            silent=True,
        )
    else:
        temp_dir = run_dir / "_temp"
        temp_dir.mkdir(parents=True, exist_ok=True)
        ruta_json = xml_a_json(
            xml_string,
            tipo,
            carpeta_salida=str(temp_dir),
            silent=True,
        )

    # 2) JSON -> Excel (bytes)
    safe_opts = _sanitize_etl_opts(etl_opts)
    nombre_excel, excel_bytes = ETL_BIMSA(
        ruta_json,
        tipo,
        return_mode="bytes",
        **safe_opts,
    )

    # 3) (Opcional) guardar Excel en disco
    if guardar_excel_en_disco:
        out_path = excel_dir / nombre_excel
        out_path.write_bytes(excel_bytes)

    return nombre_excel, excel_bytes