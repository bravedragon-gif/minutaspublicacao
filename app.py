import re
from io import BytesIO

import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Pt

st.set_page_config(page_title="Minutas", page_icon="📄")

MARCADOR = "INSERIR CAMPO PORTARIAS"

def extrair_numero(valor: str):
    m = re.search(r"(\d+)", str(valor))
    return int(m.group(1)) if m else None

def detectar_coluna_numero(df: pd.DataFrame) -> str:
    cols = [c for c in df.columns if "número" in c.lower() or "numero" in c.lower()]
    return cols[0] if cols else df.columns[0]

def preencher(docx_bytes: bytes, xlsx_bytes: bytes, marcador: str) -> bytes:
    df = pd.read_excel(BytesIO(xlsx_bytes), dtype=str).fillna("")
    df.columns = df.columns.astype(str).str.strip()

    col_num = detectar_coluna_numero(df)
    df["__ord"] = df[col_num].apply(extrair_numero)
    df = df.sort_values("__ord", ascending=True, na_position="last").drop(columns="__ord")

    doc = Document(BytesIO(docx_bytes))

    achou = False
    for p in doc.paragraphs:
        if marcador in p.text:
            achou = True
            p.text = ""
            for _, row in df.iterrows():
                texto = " - ".join([f"{c}: {str(row[c]).strip()}" for c in df.columns if str(row[c]).strip()])
                if texto.strip():
                    np = p.insert_paragraph_before(texto)
                    np.paragraph_format.space_after = Pt(12)
            break

    if not achou:
        raise ValueError(f"Marcador '{marcador}' não encontrado no DOCX.")

    out = BytesIO()
    doc.save(out)
    return out.getvalue()

st.title("Preencher minuta por planilha")
docx = st.file_uploader("Minuta em branco (DOCX)", type=["docx"])
xlsx = st.file_uploader("Planilha (XLSX)", type=["xlsx"])
marcador = st.text_input("Marcador no DOCX", value=MARCADOR)

if st.button("Gerar", type="primary", disabled=(docx is None or xlsx is None)):
    result = preencher(docx.getvalue(), xlsx.getvalue(), marcador)
    st.download_button(
        "Baixar DOCX preenchido",
        data=result,
        file_name="Minuta_Preenchida.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
