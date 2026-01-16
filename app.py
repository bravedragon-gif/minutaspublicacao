# app.py
# Streamlit app para:
# 1) Upload de uma minuta DOCX (template) + planilha XLSX
# 2) Preencher campos {{...}} e (opcionalmente) repetir linhas/blocos via docxtpl (Jinja2)
# 3) Ordenar registros (portarias/atestos) em ordem crescente
# 4) (Opcional) Inserir um bloco “lista de portarias” em um marcador de texto no DOCX
#
# Requisitos (requirements.txt):
# streamlit==1.41.1
# pandas==2.2.3
# openpyxl==3.1.5
# python-docx==1.1.2
# docxtpl==0.17.0
# Jinja2==3.1.4

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from io import BytesIO
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Pt
from docxtpl import DocxTemplate


# ---------------------------
# Utilitários de normalização
# ---------------------------

def normalize_key(s: str) -> str:
    """
    Normaliza cabeçalhos da planilha para casar com placeholders:
    - remove acentos
    - troca espaços/pontuação por _
    - deixa em MAIÚSCULAS
    Ex.: "Nº Atesto" -> "N_ATESTO", "Matrícula" -> "MATRICULA"
    """
    s = str(s).strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"[^A-Za-z0-9]+", "_", s).strip("_")
    return s.upper()


def safe_str(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    return str(v).strip()


# --------------------------------
# Parsing para ordenação crescente
# --------------------------------

_NUM_YEAR = re.compile(r"(\d+)\s*/\s*(\d{4})")      # 123/2025
_FIRST_INT = re.compile(r"(\d+)")                  # primeiro inteiro


def parse_order_value(text: str) -> Tuple[int, int]:
    """
    Retorna (ano, numero) quando achar padrão N/AAAA; caso contrário (0, N).
    Se não achar número, retorna (10**9, 10**9) para ir ao final.
    """
    text = safe_str(text)
    m = _NUM_YEAR.search(text)
    if m:
        n = int(m.group(1))
        y = int(m.group(2))
        return (y, n)
    m2 = _FIRST_INT.search(text)
    if m2:
        n = int(m2.group(1))
        return (0, n)
    return (10**9, 10**9)


def guess_order_column(d
