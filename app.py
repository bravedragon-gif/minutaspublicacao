import re
import unicodedata
from io import BytesIO
from typing import Any, Dict, List, Tuple

import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Pt


# =========================
# Config
# =========================
DEFAULT_MARKER_PORTARIAS = "INSERIR CAMPO PORTARIAS"


# =========================
# Normalização / util
# =========================
def normalize_key(s: str) -> str:
    """
    Normaliza cabeçalhos para chaves:
    - remove acento
    - troca pontuação/espaço por _
    - upper
    """
    s = str(s).strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"[^A-Za-z0-9]+", "_", s).strip("_")
    return s.upper()


def sstr(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    return str(v).strip()


def parse_order_value(text: str) -> Tuple[int, int]:
    """
    Ordenação para número (com suporte a N/AAAA):
      - se achar 123/2025 -> (2025, 123)
      - se achar 123 -> (0, 123)
      - se não achar -> (inf, inf)
    """
    text = sstr(text)
    m = re.search(r"(\d+)\s*/\s*(\d{4})", text)
    if m:
        return (int(m.group(2)), int(m.group(1)))
    m2 = re.search(r"(\d+)", text)
    if m2:
        return (0, int(m2.group(1)))
    return (10**9, 10**9)


# =========================
# Aliases (Planilha -> Campos canônicos)
# =========================
ALIASES: Dict[str, List[str]] = {
    # Campos básicos do PM
    "POSTO": ["POSTO", "POSTO_GRADUACAO", "POSTO_GRADUAÇÃO", "GRAD", "GRADUACAO", "GRADUAÇÃO"],
    "QUADRO": ["QUADRO"],
    "MATRICULA": ["MATRICULA", "MATRÍCULA", "MATR", "MAT"],
    "NOME": ["NOME", "NOME_DO_REQUERENTE", "NOME_REQUERENTE"],

    "DIAS": ["DIAS", "QUANTOS_DIAS", "QUANTIDADE_DIAS"],
    "INICIO": ["INICIO", "INÍCIO", "DATA_DE_INICIO", "DATA_INICIO", "DT_INICIO"],
    "TERMINO": ["TERMINO", "TÉRMINO", "DATA_DE_TERMINO", "DATA_TERMINO", "DT_TERMINO", "FIM"],
    "D_REST": ["D_REST", "DIAS_RESTANTES", "DIAS_REST", "RESTANTES"],

    "ATESTO": ["ATESTO", "N_ATESTO", "Nº_DO_ATESTO", "N_DO_ATESTO", "NUM_ATESTO", "NR_ATESTO"],

    # Campos SEI / processo (se você colocar placeholders depois)
    "REQUERIMENTO_DOC_SEI": ["REQUERIMENTO_DOC_SEI", "REQUERIMENTO_DOC_SEI_"],
    "DEFERIMENTO_DOC_SEI": ["DEFERIMENTO_DOC_SEI", "DEFERIMENTO_DOC_SEI_"],
    "PROCESSO_SEI": ["PROCESSO_SEI"],

    # Portarias (o que você pediu)
    "NUMERO_PORTARIA": ["NUMERO_PORTARIA", "NUMERO_DA_PORTARIA", "N_PORTARIA", "NUM_PORTARIA"],
    "TEXTO_PORTARIA": [
        "TEXTO_PORTARIA",
        "TEXTO_DA_PORTARIA",
        "INSIRA_ABAIXO_O_TEXTO_DA_PORTARIA",
        "INSIRA_ABAIXO_O_TEXTO_DA_PORTARIA_",
    ],
}


def apply_aliases(df_norm: pd.DataFrame) -> pd.DataFrame:
    """
    Renomeia colunas normalizadas para chaves canônicas quando existir correspondência.
    """
    cols = set(df_norm.columns)
    rename_map = {}

    for canonical, variants in ALIASES.items():
        if canonical in cols:
            continue
        for v in variants:
            vv = normalize_key(v)
            if vv in cols:
                rename_map[vv] = canonical
                break

    if rename_map:
        df_norm = df_norm.rename(columns=rename_map)

    return df_norm


# =========================
# DOCX helpers
# =========================
def iter_all_paragraphs(doc: Document):
    # Corpo
    for p in doc.paragraphs:
        yield p

    # Tabelas (células)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p

    # Cabeçalho/rodapé
    for sec in doc.sections:
        for p in sec.header.paragraphs:
            yield p
        for t in sec.header.tables:
            for row in t.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        yield p

        for p in sec.footer.paragraphs:
            yield p
        for t in sec.footer.tables:
            for row in t.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        yield p


def replace_placeholders(doc: Document, mapping: Dict[str, str]) -> None:
    """
    Substitui tokens {{CHAVE}} por valores em todo o documento.
    Observação: usa paragraph.text (pode simplificar formatação do parágrafo onde houver substituição).
    """
    for p in iter_all_paragraphs(doc):
        if not p.text:
            continue
        txt = p.text
        new_txt = txt
        for k, v in mapping.items():
            token = "{{" + k + "}}"
            if token in new_txt:
                new_txt = new_txt.replace(token, v)
        if new_txt != txt:
            p.text = new_txt


def insert_portarias_at_marker(
    doc: Document,
    marker: str,
    portarias_sorted: List[Dict[str, str]],
    num_key: str = "NUMERO_PORTARIA",
    text_key: str = "TEXTO_PORTARIA",
    space_after_pt: int = 12,
) -> bool:
    """
    Encontra o parágrafo com o marcador e insere:
      Portaria nº <NUMERO_PORTARIA>
      <TEXTO_PORTARIA completo, preservando linhas>
    Retorna True se encontrou marcador.
    """
    for p in doc.paragraphs:
        if marker in (p.text or ""):
            # limpa o marcador
            p.text = ""

            for r in portarias_sorted:
                num = sstr(r.get(num_key, ""))
                txt = r.get(text_key, "")
                txt = "" if txt is None else str(txt)  # texto completo, sem resumo

                if not (num or txt.strip()):
