# ETL_XML_to_JSON.py
import re
import os
import json
import xml.etree.ElementTree as ET
from datetime import datetime
from typing import Union, Any, Dict, List



def _elem_to_dict(elem: ET.Element):
    children = list(elem)
    if not children:
        return elem.text or ""

    grouped = {}
    for ch in children:
        tag = ch.tag.split("}", 1)[-1]
        grouped.setdefault(tag, []).append(_elem_to_dict(ch))

    out = {}
    for k, v in grouped.items():
        out[k] = v[0] if len(v) == 1 else v
    return out


def _force_list_payload(payload: Any) -> List[Dict[str, Any]]:
    """
    Garantiza que el JSON final SIEMPRE sea una LISTA de registros (list[dict]).
    """
    if payload is None:
        return []

    # Si ya es lista:
    if isinstance(payload, list):
        # si hay strings sueltos, envuélvelos
        out = []
        for x in payload:
            out.append(x if isinstance(x, dict) else {"value": x})
        return out

    # Si es dict:
    if isinstance(payload, dict):
        # Caso típico: { "row": [ {...}, {...} ] } o { "registro": [...] }
        list_candidates = []
        for v in payload.values():
            if isinstance(v, list) and v and all(isinstance(i, dict) for i in v):
                list_candidates.append(v)

        if list_candidates:
            # toma la lista más larga (más probable que sean los registros)
            return max(list_candidates, key=len)

        # Caso: { "datos": { "row": [...] } } ya lo habrás desempacado antes, pero por si acaso:
        for v in payload.values():
            if isinstance(v, dict):
                inner = _force_list_payload(v)
                if inner:
                    return inner

        # Si es un registro único (dict plano), lo volvemos lista
        return [payload]

    # Cualquier otro tipo:
    return [{"value": payload}]


def xml_a_json(
    xml_input: Union[str, bytes],
    tipo_reporte: str,
    carpeta_salida: str = ".",
    silent: bool = True,
) -> str:

    # 1) Normaliza entrada
    if isinstance(xml_input, bytes):
        for enc in ("utf-8-sig", "utf-8", "cp1252", "cp850", "latin1"):
            try:
                xml_string = xml_input.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        else:
            xml_string = xml_input.decode("utf-8", errors="replace")
    else:
        xml_string = xml_input or ""

    # 2) Limpieza segura
    xml_limpio = xml_string.strip().lstrip("\ufeff")

    # 3) (ETLV3) Sin procesamiento de texto para maximizar performance

    # 5) Parse
    try:
        root = ET.fromstring(xml_limpio)
    except ET.ParseError as e:
        raise ValueError(f"XML inválido tras limpieza: {e}")

    payload = _elem_to_dict(root)

    # Si viene envuelto:
    if isinstance(payload, dict) and "datos" in payload:
        payload = payload["datos"]

    payload_list = _force_list_payload(payload)

    # 6) Guardar JSON UTF-8 real
    os.makedirs(carpeta_salida, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = os.path.join(carpeta_salida, f"{tipo_reporte}_{ts}.json")

    with open(out_path, "w", encoding="utf-8", newline="\n") as f:
        json.dump(payload_list, f, ensure_ascii=False, indent=2)

    if not silent:
        print(f"[XML2JSON] OK -> {out_path}")

    return out_path