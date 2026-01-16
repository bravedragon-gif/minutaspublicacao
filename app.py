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
                    continue

                # Título com número
                head = p.insert_paragraph_before(f"Portaria nº {num}".strip())
                head.style = doc.styles["Normal"]
                head.paragraph_format.space_after = Pt(0)

                # Corpo com texto completo (linhas viram parágrafos)
                lines = txt.splitlines() if txt else [""]
                for line in lines:
                    body = p.insert_paragraph_before(line)
                    body.style = doc.styles["Normal"]
                    body.paragraph_format.space_after = Pt(0)

                # Espaço entre portarias
                sep = p.insert_paragraph_before("")
                sep.paragraph_format.space_after = Pt(space_after_pt)

            return True

    return False


# =========================
# Streamlit UI
# =========================
st.set_page_config(page_title="Minuta — Preenchimento + Portarias", page_icon="📄", layout="centered")
st.title("Preencher Minuta (DOCX) com Planilha (XLSX) + Portarias em ordem crescente")

docx_file = st.file_uploader("Minuta em branco (DOCX)", type=["docx"])
xlsx_file = st.file_uploader("Dados (XLSX)", type=["xlsx"])

with st.expander("Configurações", expanded=True):
    marker_portarias = st.text_input(
        "Marcador do bloco de Portarias no DOCX",
        value=DEFAULT_MARKER_PORTARIAS,
        help="Na sua minuta está como 'INSERIR CAMPO PORTARIAS'.",
    )
    space_after = st.number_input("Espaçamento entre portarias (pt)", min_value=0, max_value=48, value=12, step=1)

if not docx_file or not xlsx_file:
    st.stop()

try:
    # 1) Ler planilha
    df = pd.read_excel(BytesIO(xlsx_file.getvalue()), dtype=str).fillna("")
    df.columns = df.columns.astype(str).str.strip()

    # 2) Normalizar colunas
    df_norm = df.copy()
    df_norm.columns = [normalize_key(c) for c in df_norm.columns]
    for c in df_norm.columns:
        df_norm[c] = df_norm[c].map(sstr)

    # 3) Aplicar aliases (para garantir NUMERO_PORTARIA / TEXTO_PORTARIA etc.)
    df_norm = apply_aliases(df_norm)
    cols_norm = list(df_norm.columns)

    # 4) Registros
    records = df_norm.to_dict(orient="records")
    if not records:
        raise ValueError("Planilha sem registros.")

    # 5) Mapping para placeholders gerais: usa a primeira linha
    mapping = {k: sstr(v) for k, v in records[0].items()}

    # 6) Portarias: filtra e ordena por NUMERO_PORTARIA
    if "NUMERO_PORTARIA" in cols_norm:
        portarias = [r for r in records if sstr(r.get("NUMERO_PORTARIA", "")) or sstr(r.get("TEXTO_PORTARIA", ""))]
        portarias_sorted = sorted(portarias, key=lambda r: parse_order_value(r.get("NUMERO_PORTARIA", "")))
    else:
        portarias_sorted = []

    # 7) Abrir DOCX e preencher
    doc = Document(BytesIO(docx_file.getvalue()))

    replace_placeholders(doc, mapping)

    # 8) Inserir portarias (se houver colunas e marcador existir)
    inserted = False
    if portarias_sorted and ("TEXTO_PORTARIA" in cols_norm):
        inserted = insert_portarias_at_marker(
            doc,
            marker=marker_portarias,
            portarias_sorted=portarias_sorted,
            num_key="NUMERO_PORTARIA",
            text_key="TEXTO_PORTARIA",
            space_after_pt=int(space_after),
        )

    # 9) Salvar saída
    out = BytesIO()
    doc.save(out)
    out_bytes = out.getvalue()

    st.success("Minuta gerada com sucesso.")

    if portarias_sorted:
        st.caption(f"Portarias detectadas: {len(portarias_sorted)} (ordenadas crescente por NUMERO_PORTARIA).")
    else:
        st.warning("Nenhuma portaria detectada na planilha (coluna NUMERO_PORTARIA/TEXTO_PORTARIA vazias ou ausentes).")

    if not inserted:
        st.info(
            f"Não encontrei o marcador '{marker_portarias}' no corpo do documento. "
            f"Na sua minuta ele existe como texto literal. Verifique se está idêntico."
        )

    with st.expander("Colunas detectadas (normalizadas)", expanded=False):
        st.code(", ".join(cols_norm))

    st.download_button(
        "Baixar DOCX preenchido",
        data=out_bytes,
        file_name="Minuta_Preenchida.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

except Exception as e:
    st.error(f"Erro: {e}")
