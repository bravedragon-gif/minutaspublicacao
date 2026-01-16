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


def guess_order_column(df_cols_norm: List[str]) -> Optional[str]:
    """
    Tenta adivinhar uma coluna de ordenação (portaria/atesto/número).
    """
    priorities = [
        "PORTARIA",
        "N_PORTARIA",
        "NUM_PORTARIA",
        "N_ATESTO",
        "ATESTO",
        "NUMERO",
        "N_NUMERO",
        "N",
    ]
    for p in priorities:
        if p in df_cols_norm:
            return p

    # fallback: qualquer coluna contendo essas palavras
    for c in df_cols_norm:
        if any(k in c for k in ("PORTARIA", "ATESTO", "NUMER", "N_")):
            return c
    return None


# ---------------------------
# Leitura da planilha (XLSX)
# ---------------------------

@dataclass
class SheetData:
    df_raw: pd.DataFrame
    df_norm: pd.DataFrame
    records: List[Dict[str, str]]          # registros com chaves normalizadas (MAIÚSCULAS)
    globals_first_row: Dict[str, str]      # campos globais (primeira linha)


def load_sheet(xlsx_bytes: bytes) -> SheetData:
    df = pd.read_excel(BytesIO(xlsx_bytes), dtype=str).fillna("")
    if df.empty:
        raise ValueError("A planilha está vazia (sem linhas de dados).")

    # Normaliza colunas
    col_map = {c: normalize_key(c) for c in df.columns}
    df_norm = df.rename(columns=col_map).copy()

    # Converte tudo para string “limpa”
    for c in df_norm.columns:
        df_norm[c] = df_norm[c].map(safe_str)

    records = df_norm.to_dict(orient="records")
    globals_first = records[0].copy() if records else {}

    return SheetData(df_raw=df, df_norm=df_norm, records=records, globals_first_row=globals_first)


def sort_records(records: List[Dict[str, str]], order_col: str) -> List[Dict[str, str]]:
    order_col = normalize_key(order_col)
    return sorted(records, key=lambda r: parse_order_value(r.get(order_col, "")))


# ------------------------------------
# Renderização DOCX com docxtpl (Jinja)
# ------------------------------------

def render_template_docx(template_bytes: bytes, context: Dict[str, Any]) -> bytes:
    tpl = DocxTemplate(BytesIO(template_bytes))
    tpl.render(context)
    out = BytesIO()
    tpl.save(out)
    return out.getvalue()


# ----------------------------------------------------
# Inserção opcional de bloco no marcador (python-docx)
# ----------------------------------------------------

def insert_block_at_marker(docx_bytes: bytes, marker: str, lines: List[str], space_after_pt: int = 12) -> bytes:
    """
    Insere 'lines' no local de um marcador de texto exato em parágrafos do corpo.
    Observação: se o marcador estiver dentro de tabela/cabeçalho, pode exigir ajuste adicional.
    """
    doc = Document(BytesIO(docx_bytes))
    found = False

    for p in doc.paragraphs:
        if marker in p.text:
            found = True
            p.text = ""  # limpa o marcador
            for line in lines:
                if not line.strip():
                    continue
                np = p.insert_paragraph_before(line)
                np.style = doc.styles["Normal"]
                np.paragraph_format.space_after = Pt(space_after_pt)
            break

    if not found:
        raise ValueError(f"Marcador '{marker}' não encontrado no corpo do documento.")

    out = BytesIO()
    doc.save(out)
    return out.getvalue()


def build_portaria_lines(
    records: List[Dict[str, str]],
    include_cols: List[str],
    label_cols: bool = True,
    sep: str = " - ",
) -> List[str]:
    cols = [normalize_key(c) for c in include_cols]
    lines: List[str] = []
    for r in records:
        parts = []
        for c in cols:
            v = safe_str(r.get(c, ""))
            if not v:
                continue
            parts.append(f"{c}: {v}" if label_cols else v)
        lines.append(sep.join(parts))
    return lines


# ---------------
# Interface Streamlit
# ---------------

st.set_page_config(page_title="Minutas — Preenchimento por Planilha", page_icon="📄", layout="centered")

st.title("Preenchimento de minuta (DOCX) por planilha (XLSX)")
st.write(
    "Fluxo: faça upload da minuta (DOCX template) e da planilha (XLSX). "
    "O app ordena os registros em ordem crescente e preenche os campos."
)

with st.expander("Como deve estar o template (DOCX)", expanded=False):
    st.markdown(
        """
**1) Campos simples (texto corrido):** use `{{CAMPO}}`  
Ex.: `O {{POSTO}} {{QUADRO}} {{NOME}}, matr. {{MATRICULA}} ...`

**2) Tabela com repetição (várias linhas):** use loop Jinja (docxtpl).  
Na linha de dados da tabela, faça algo assim (exemplo conceitual):

- 1ª célula da linha: `{% for r in registros %}{{r.POSTO}}`  
- Demais células: `{{r.QUADRO}}`, `{{r.MATRICULA}}`, `{{r.NOME}}`...  
- Última célula da mesma linha: `{{r.ATESTO}}{% endfor %}`

Isso faz o Word repetir a linha para cada item em `registros`.
"""
    )

col1, col2 = st.columns(2)
with col1:
    docx_file = st.file_uploader("Minuta (DOCX template)", type=["docx"])
with col2:
    xlsx_file = st.file_uploader("Planilha (XLSX)", type=["xlsx"])

if not docx_file or not xlsx_file:
    st.stop()

# Carrega planilha
try:
    sheet = load_sheet(xlsx_file.getvalue())
except Exception as e:
    st.error(f"Erro ao ler planilha: {e}")
    st.stop()

df_norm = sheet.df_norm
cols_norm = list(df_norm.columns)

with st.expander("Configurações", expanded=True):
    # Ordenação
    guessed = guess_order_column(cols_norm) or (cols_norm[0] if cols_norm else "")
    order_col = st.selectbox(
        "Ordenar crescente por (coluna da planilha)",
        options=cols_norm,
        index=cols_norm.index(guessed) if guessed in cols_norm else 0,
    )

    # Campos globais extras (opcional)
    st.caption("Campos globais opcionais (se você usar no DOCX como {{ANO}}, {{MES}}, etc.).")
    extras = {
        "ANO": st.text_input("ANO (opcional)", value=""),
        "MES": st.text_input("MES (opcional)", value=""),
        "EXERCICIO": st.text_input("EXERCICIO (opcional)", value=""),
    }
    # Remove vazios
    extras = {k: v for k, v in extras.items() if safe_str(v)}

    # Opção: inserir bloco de portarias em marcador
    st.divider()
    use_marker_block = st.checkbox("Também inserir bloco (lista) em um marcador de texto", value=False)

    marker_text = ""
    marker_cols: List[str] = []
    marker_label_cols = True
    marker_space = 12

    if use_marker_block:
        marker_text = st.text_input("Texto do marcador no DOCX", value="INSERIR CAMPO PORTARIAS")
        marker_cols = st.multiselect(
            "Colunas para compor cada linha do bloco",
            options=cols_norm,
            default=[c for c in cols_norm if c in ("PORTARIA", "ATESTO", "N_ATESTO", "NUMERO", "NOME", "MATRICULA")][:4]
            or cols_norm[:4],
        )
        marker_label_cols = st.checkbox("Incluir rótulo da coluna (COL: valor)", value=True)
        marker_space = int(st.number_input("Espaçamento após cada linha (pt)", min_value=0, max_value=48, value=12, step=1))

# Processa e gera
if st.button("Gerar minuta preenchida", type="primary"):
    try:
        # 1) Ordena registros
        records_sorted = sort_records(sheet.records, order_col=order_col)

        # 2) Contexto do docxtpl:
        #    - registros: lista para loops
        #    - globals: primeira linha para {{POSTO}}, {{NOME}} etc (se você usar sem r.)
        context: Dict[str, Any] = {}
        context.update(sheet.globals_first_row)  # permite {{POSTO}} sem loop (pega 1ª linha)
        context.update(extras)                   # campos globais adicionais
        context["registros"] = records_sorted    # para loops: {% for r in registros %}

        # 3) Renderiza template
        rendered = render_template_docx(docx_file.getvalue(), context)

        # 4) Opcional: inserir bloco em marcador
        if use_marker_block:
            if not marker_cols:
                raise ValueError("Selecione ao menos uma coluna para compor o bloco.")
            lines = build_portaria_lines(records_sorted, marker_cols, label_cols=marker_label_cols)
            rendered = insert_block_at_marker(rendered, marker_text, lines, space_after_pt=marker_space)

        st.success("Minuta gerada com sucesso.")
        st.download_button(
            "Baixar DOCX preenchido",
            data=rendered,
            file_name="minuta_preenchida.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

        with st.expander("Prévia das colunas normalizadas (para casar com {{...}})", expanded=False):
            st.write("Use estes nomes (MAIÚSCULOS) nos placeholders do template:")
            st.code(", ".join(cols_norm))

    except Exception as e:
        st.error(f"Erro ao gerar minuta: {e}")
        st.info(
            "Se a tabela não estiver repetindo linhas, verifique se você inseriu o loop "
            "`{% for r in registros %}` e o `{% endfor %}` na linha de dados da tabela."
        )
