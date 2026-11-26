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

            if "confecções elite" in normalizar(contexto):
                continue

            encontrados.append(contexto)

    return encontrados


def buscar_numero_pdf(texto_pdf, documento):
    numero = somente_numeros(documento)

    if not numero:
        return False

    # Procura o número em linhas do relatório
    for linha in texto_pdf.splitlines():

        linha_num = somente_numeros(linha)

        if numero == linha_num:
            return True

        # Considera documento com sufixo /001 ou /002
        if linha_num.startswith(numero) and len(linha_num) <= len(numero) + 8:
            return True

    return False


def encontrar_similares(texto_pdf, documento, limite=0.70):
    numero = somente_numeros(documento)
    candidatos = []

    if not numero:
        return candidatos

    # Extrai números que aparecem no PDF
    encontrados = re.findall(r"\b\d{3,}(?:/\d{1,3})?\b", texto_pdf)

    vistos = set()

    for candidato in encontrados:

        cand_num = somente_numeros(candidato)

        if not cand_num or cand_num in vistos:
            continue

        vistos.add(cand_num)

        similaridade = SequenceMatcher(
            None,
            numero,
            cand_num
        ).ratio()

        if similaridade >= limite and cand_num != numero:
            candidatos.append((candidato, similaridade))

    candidatos.sort(key=lambda x: x[1], reverse=True)

    return candidatos[:5]


def obter_contexto(texto_pdf, documento):
    linhas = texto_pdf.splitlines()
    numero = somente_numeros(documento)
    resultados = []

    for i, linha in enumerate(linhas):

        linha_num = somente_numeros(linha)

        if numero and numero in linha_num and len(linha_num) <= len(numero) + 8:

            trecho = linhas[max(0, i-2):min(len(linhas), i+8)]

            resultados.append("\n".join(trecho))

    return resultados


# =========================================================
# UPLOAD
# =========================================================

st.subheader("1. Envie os arquivos")

col1, col2 = st.columns(2)

with col1:
    excel_file = st.file_uploader(
        "Remessa Excel",
        type=["xlsx", "xls"]
    )

with col2:
    pdf_file = st.file_uploader(
        "Relatório diário PDF",
        type=["pdf"]
    )

st.info(
    "A primeira versão utiliza a aba FILIAL e a primeira coluna "
    "da planilha para identificar os documentos."
)

# =========================================================
# PROCESSAMENTO
# =========================================================

if excel_file and pdf_file:

    if st.button("🔎 CONFRONTAR DOCUMENTOS", type="primary", use_container_width=True):

        try:

            documentos = ler_excel(excel_file)
            texto_pdf = ler_pdf(pdf_file)

            resultados = []

            for documento in documentos:

                correspondencias = localizar_documentos(
                    texto_pdf,
                    documento
                )

                if correspondencias:

                    resultados.append({
                        "documento": documento,
                        "status": "Encontrado",
                        "correspondencia": "Localizado no relatório",
                        "destino": "Conferir contexto abaixo",
                        "contexto": correspondencias
                    })

                else:

                    similares = encontrar_similares(
                        texto_pdf,
                        documento
                    )

                    if similares:

                        resultados.append({
                            "documento": documento,
                            "status": "Divergência possível",
                            "correspondencia": similares,
                            "destino": "Não confirmado",
                            "contexto": []
                        })

                    else:

                        resultados.append({
                            "documento": documento,
                            "status": "Compra para confirmar",
                            "correspondencia": "Não localizado",
                            "destino": "Confirmar origem",
                            "contexto": []
                        })

            st.session_state["resultados"] = resultados
            st.session_state["texto_pdf"] = texto_pdf

        except Exception as e:
            st.error(f"Erro ao processar os arquivos: {e}")


# =========================================================
# RESULTADOS
# =========================================================

if "resultados" in st.session_state:

    resultados = st.session_state["resultados"]

    st.divider()
    st.subheader("2. Conferência dos documentos")

    st.caption(
        "A ordem abaixo segue a ordem da remessa Excel. "
        "Use os controles para confirmar ou sinalizar correções."
    )

    for i, resultado in enumerate(resultados):

        documento = resultado["documento"]

        with st.container(border=True):

            st.markdown(f"### {i+1}. Documento {documento}")

            status = resultado["status"]

            if status == "Encontrado":

                st.success("Documento localizado no relatório")

                st.write("**Correspondência:**", resultado["correspondencia"])

                contextos = resultado["contexto"]

                if contextos:

                    st.markdown("**Destino / trecho da baixa no relatório:**")

                    for contexto in contextos:
                        st.text(contexto)

            elif status == "Divergência possível":

                st.warning("Possível divergência de numeração")

                st.write("**Número da remessa:**", documento)

                st.write("**Possíveis correspondências:**")

                for candidato, percentual in resultado["correspondencia"]:

                    st.write(
                        f"- {candidato} — "
                        f"{percentual*100:.1f}% de semelhança"
                    )

                st.error(
                    "A semelhança NÃO confirma que seja o mesmo documento."
                )

            else:

                st.warning("Compra para confirmar")

                st.write(
                    "Documento não localizado no relatório PDF enviado."
                )

            # Único controle individual: checkbox de confirmação
            chave = f"conferencia_{i}"

            st.checkbox(
                "Confirmar",
                key=chave
            )

# =========================================================
# RESUMO
# =========================================================

if "resultados" in st.session_state:

    st.divider()
    st.subheader("3. Resumo")

    resultados = st.session_state["resultados"]

    total = len(resultados)
    confirmados = sum(
        bool(st.session_state.get(f"conferencia_{i}", False))
        for i in range(total)
    )

    st.metric("Confirmados", confirmados)
    st.metric("Ainda não confirmados", total - confirmados)

    # =====================================================
    # EXPORTAÇÃO DO RESUMO
    # =====================================================

    linhas_resumo = []

    for i, resultado in enumerate(resultados):

        linhas_resumo.append({
            "Ordem": i+1,
            "Documento": resultado["documento"],
            "Status automático": resultado["status"],
            "Confirmado": "SIM" if st.session_state.get(f"conferencia_{i}", False) else "NÃO"
        })

    df_resultado = pd.DataFrame(linhas_resumo)

    st.subheader("4. Resultado final")

    st.dataframe(
        df_resultado,
        use_container_width=True,
        hide_index=True
    )

    csv = df_resultado.to_csv(
        index=False,
        sep=";",
        encoding="utf-8-sig"
    )

    st.download_button(
        "📥 Baixar resultado CSV",
        data=csv,
        file_name="conferencia_baixas.csv",
        mime="text/csv",
        use_container_width=True
    )

    texto_copia = "\n".join(
        [
            f"{linha['Ordem']}. Documento {linha['Documento']} | "
            f"{linha['Status automático']} | "
            f"Confirmado: {linha['Confirmado']}"
            for linha in linhas_resumo
        ]
    )

    st.text_area(
        "Resumo para copiar",
        texto_copia,
        height=250
    )

# =========================================================
# RODAPÉ
# =========================================================

st.divider()

st.caption(
    "Conferência de Baixas | Ferramenta de apoio. "
    "A confirmação final deve ser realizada pelo responsável."
)
