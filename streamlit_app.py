import streamlit as st
import pandas as pd
import fitz
import re
import unicodedata
from difflib import SequenceMatcher
from io import BytesIO

# =========================================================
# CONFIGURAÇÃO
# =========================================================

st.set_page_config(
    page_title="Conferência de Baixas",
    page_icon="📋",
    layout="wide"
)

st.title("📋 Conferência de Baixas")
st.caption("Confronto de remessas Excel com relatórios PDF de baixas diretas.")

# =========================================================
# FUNÇÕES
# =========================================================

def normalizar(texto):
    texto = str(texto).strip().lower()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    return texto


def somente_numeros(texto):
    return re.sub(r"\D", "", str(texto))


def ler_excel(arquivo):
    wb = pd.ExcelFile(arquivo)

    if "FILIAL" not in wb.sheet_names:
        raise ValueError("A planilha não possui uma aba chamada FILIAL.")

    df = pd.read_excel(arquivo, sheet_name="FILIAL", header=None)

    documentos = []

    # Lê a primeira coluna preservando a ordem
    for valor in df.iloc[:, 0].tolist():

        if pd.isna(valor):
            continue

        texto = str(valor).strip()

        if texto.lower() in ["nan", "documento", "doc", "nota"]:
            continue

        if "total" in normalizar(texto):
            continue

        numero = somente_numeros(texto)

        if not numero:
            continue

        documentos.append(texto)

    return documentos


def ler_pdf(arquivo):
    pdf = fitz.open(stream=arquivo.read(), filetype="pdf")

    texto_completo = ""

    for pagina in pdf:
        texto_completo += pagina.get_text() + "\n"

    return texto_completo


def localizar_documentos(texto_pdf, documento):
    numero = somente_numeros(documento)

    if not numero:
        return []

    linhas = texto_pdf.splitlines()
    encontrados = []

    for i, linha in enumerate(linhas):

        numeros_linha = somente_numeros(linha)

        if numero in numeros_linha and len(numeros_linha) <= len(numero) + 10:
            contexto = "\n".join(linhas[max(0, i-2):min(len(linhas), i+5)])
