import re
from io import BytesIO

import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Pt

st.set_page_config(page_title="Minutas", page_icon="📄")

MARCADOR_PORTARIAS = "INSERIR CAMPO PORTARIAS"


def extrair_numero(valor: str):
    m = re.search(r"(\d+)", str(valor))
    return int(m.group(1)) if m else None


def detectar_coluna_numero(df: pd.DataFrame) -> str:
    cols = [c for c in df.columns if "número" in c.lower() or "numero" in c.lower()]
    return cols[0] if cols else df.columns[0]


def iter_paragrafos_doc(doc: Document):
    # Corpo
    for p in doc.paragraphs:
        yield p

    # Tabelas (células)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p

    # Cabeçalho/rodapé (todas as seções)
    for section in doc.sections:
        for p in section.header.paragraphs:
            yield p
        for table in section.header.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        yield p

        for p in section.footer.paragraphs:
            yield p
        for table in section.footer.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        yield p


def substituir_campos(doc: Document, mapping: dict[str, str]):
    """
    Substitui placeholders no formato {{CAMPO}} por valores.
    Observação: usa paragraph.text (pode simplificar estilos no trecho substituído).
    """
    for p in iter_paragrafos_doc(doc):
        txt = p.text
        if not txt:
            continue

        novo = txt
        for k, v in mapping.items():
            # Substitui {{COLUNA}}
            token = "{{" + k + "}}"
            if token in novo:
                novo = novo.replace(token, v)

        if novo != txt:
            p.text = novo


def inserir_portarias(doc: Document, df_sorted: pd.DataFrame, marcador: str, espaco_apos_pt: int = 12):
    """
    Insere a lista de portarias (uma linha por registro) no local do marcador.
    """
    achou = False

    for p in doc.paragraphs:
        if marcador in p.text:
            achou = True
            p.text = ""

            # monta texto por linha (coluna: valor)
            cols = list(df_sorted.columns)
            for _, row in df_sorted.iterrows():
                partes = []
                for c in cols:
                    v = str(row[c]).strip()
                    if v:
                        partes.append(f"{c}: {v}")
                texto = " - ".join(partes).strip()
                if not texto:
                    continue

                np = p.insert_paragraph_before(texto)
                np.style = doc.styles["Normal"]
                np.paragraph_format.space_after = Pt(espaco_apos_pt)

            break

    if not achou:
        raise ValueError(f"Não encontrei o marcador '{marcador}' no DOCX.")


def processar(docx_bytes: bytes, xlsx_bytes: bytes, marcador_portarias: str, espaco_apos_pt: int = 12) -> bytes:
    df = pd.read_excel(BytesIO(xlsx_bytes), dtype=str).fillna("")
    df.columns = df.columns.astype(str).str.strip()

    if df.empty:
        raise ValueError("A planilha está vazia (sem linhas de dados).")

    # 1) ordena portarias
    col_num = detectar_coluna_numero(df)
    df["__ord"] = df[col_num].apply(extrair_numero)
    df_sorted = df.sort_values("__ord", ascending=True, na_position="last").drop(columns="__ord")

    # 2) abre doc
    doc = Document(BytesIO(docx_bytes))

    # 3) substitui campos usando a PRIMEIRA LINHA como “dados fixos”
    primeira = df_sorted.iloc[0].to_dict()
    mapping = {str(k).strip(): str(v).strip() for k, v in primeira.items()}

    substituir_campos(doc, mapping)

    # 4) insere portarias no marcador
    inserir_portarias(doc, df_sorted, marcador_portarias, espaco_apos_pt=espaco_apos_pt)

    out = BytesIO()
    doc.save(out)
    return out.getvalue()


st.title("Preencher minuta por planilha (campos + portarias)")

docx = st.file_uploader("Minuta em branco (DOCX)", type=["docx"])
xlsx = st.file_uploader("Planilha (XLSX)", type=["xlsx"])

with st.expander("Configurações", expanded=True):
    marcador = st.text_input("Marcador do bloco de portarias", value=MARCADOR_PORTARIAS)
    espaco_apos = st.number_input("Espaçamento após cada portaria (pt)", 0, 48, 12, 1)

if st.button("Gerar", type="primary", disabled=(docx is None or xlsx is None)):
    try:
        result = processar(docx.getvalue(), xlsx.getvalue(), marcador, espaco_apos_pt=int(espaco_apos))
        st.success("Minuta gerada com sucesso.")
        st.download_button(
            "Baixar DOCX preenchido",
            data=result,
            file_name="Minuta_Preenchida.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except Exception as e:
        st.error(f"Erro: {e}")
