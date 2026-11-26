import streamlit as st
import pandas as pd
import fitz
import re
import unicodedata
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

st.set_page_config(page_title="Conferência de Baixas", page_icon="📋", layout="wide")
st.title("📋 Conferência de Baixas")
st.caption("Confronto da sequência Excel com o relatório PDF de baixas diretas.")


def normalizar(txt):
    txt = unicodedata.normalize("NFD", str(txt).strip().lower())
    return "".join(c for c in txt if unicodedata.category(c) != "Mn")


def numeros(txt):
    return re.sub(r"\D", "", str(txt))


def ler_remessa(arquivo):
    df = pd.read_excel(arquivo, sheet_name="FILIAL", header=None)
    # Modelo simplificado: número do documento na coluna B. Também aceita
    # arquivos antigos com números na coluna A, caso a coluna B esteja vazia.
    coluna = 1 if df.shape[1] > 1 and df.iloc[:, 1].notna().any() else 0
    docs = []
    for v in df.iloc[:, coluna].tolist():
        if pd.isna(v):
            continue
        s = str(v).strip()
        n = numeros(s)
        if not n or normalizar(s) in ("documento", "numero do documento", "n° do doc.", "n do doc.", "doc", "nota"):
            continue
        docs.append(s)
    return docs


def quantidade_pt(txt):
    txt = str(txt).strip().replace(".", "").replace(",", ".")
    try:
        return float(txt)
    except ValueError:
        return None


def extrair_pdf(arquivo):
    pdf = fitz.open(stream=arquivo.read(), filetype="pdf")
    linhas = [l.strip() for pg in pdf for l in pg.get_text().splitlines() if l.strip()]
    registros = {}
    fornecedor = ""
    nota = ""
    i = 0
    while i < len(linhas):
        linha = linhas[i]
        if "Fornecedor :" in linha:
            parte = linha.split("Fornecedor :", 1)[1].strip()
            fornecedor = re.split(r"\s+Vencimento\s*:", parte, flags=re.I)[0].strip()
        if "Nº da Nota" in linha or "N° da Nota" in linha or "No da Nota" in linha:
            candidatos = re.findall(r"\b\d{3,}(?:/\d{1,3})?\b", linha)
            if i + 1 < len(linhas):
                candidatos += re.findall(r"\b\d{3,}(?:/\d{1,3})?\b", linhas[i + 1])
            com_barra = [c for c in candidatos if "/" in c]
            nota = com_barra[0] if com_barra else (candidatos[0] if candidatos else "")
            if nota:
                chave = numeros(nota)
                registros.setdefault(chave, {"nota": nota, "empresa": fornecedor, "destinos": []})
                registros[chave]["empresa"] = fornecedor or registros[chave]["empresa"]
        if linha.lower().startswith("destino") and nota:
            desc = linhas[i + 1] if i + 1 < len(linhas) else ""
            desc = re.sub(r"^(Estoque:|Centro de Custo:)", "", desc, flags=re.I).strip()
            qty = None
            for j in range(i + 2, min(i + 6, len(linhas))):
                m = re.match(r"^(-?\d+[.,]\d+)\s*(?:[A-Za-z]{1,6})?$", linhas[j])
                if m:
                    qty = quantidade_pt(m.group(1))
                    break
            # Regra: zero e negativos não são destinos efetivos.
            if qty is not None and qty > 0 and desc:
                chave = numeros(nota)
                registros.setdefault(chave, {"nota": nota, "empresa": fornecedor, "destinos": []})
                entrada = f"{desc} — {qty:g}"
                if entrada not in registros[chave]["destinos"]:
                    registros[chave]["destinos"].append(entrada)
                registros[chave]["empresa"] = fornecedor or registros[chave]["empresa"]
        i += 1
    return registros


def encontrar(documento, registros):
    alvo = numeros(documento)
    if not alvo:
        return None
    # Correspondência exata. Sem prefixos/parciais para evitar associações indevidas.
    candidatos = [r for chave, r in registros.items() if chave == alvo]
    return candidatos[0] if len(candidatos) == 1 else None


def pdf_confirmados(resultados):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    estilos = getSampleStyleSheet()
    titulo = ParagraphStyle("Titulo", parent=estilos["Title"], fontSize=16, leading=20, alignment=1, spaceAfter=12)
    normal = ParagraphStyle("NormalNovo", parent=estilos["Normal"], fontSize=9, leading=12)
    story = [Paragraph("CONFERÊNCIA DE BAIXAS", titulo), Paragraph("Documentos confirmados", normal), Spacer(1, 12)]
    dados = [["Documento", "Empresa / Fornecedor", "Destino(s) da baixa"]]
    for r in resultados:
        dest = "<br/>".join(r["destinos"]) if r["destinos"] else "Destino não identificado"
        dados.append([Paragraph(str(r["documento"]), normal), Paragraph(r["empresa"] or "Não identificado", normal), Paragraph(dest, normal)])
    tabela = Table(dados, colWidths=[75, 190, 245], repeatRows=1)
    tabela.setStyle(TableStyle([("BACKGROUND", (0,0), (-1,0), colors.HexColor("#D9E2F3")), ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"), ("GRID", (0,0), (-1,-1), 0.5, colors.grey), ("VALIGN", (0,0), (-1,-1), "TOP"), ("LEFTPADDING", (0,0), (-1,-1), 6), ("RIGHTPADDING", (0,0), (-1,-1), 6), ("TOPPADDING", (0,0), (-1,-1), 6), ("BOTTOMPADDING", (0,0), (-1,-1), 6)]))
    story.append(tabela)
    doc.build(story)
    buffer.seek(0)
    return buffer


col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader("Remessa Excel", type=["xlsx", "xls"])
with col2:
    pdf_file = st.file_uploader("Relatório diário PDF", type=["pdf"])
st.info("Use o modelo simplificado: número do documento na coluna B. O PDF é a fonte de fornecedor e destinos. Destinos com quantidade 0,00 são ignorados.")

if excel_file and pdf_file:
    if st.button("🔎 CONFRONTAR DOCUMENTOS", type="primary", use_container_width=True):
        try:
            documentos = ler_remessa(excel_file)
            registros = extrair_pdf(pdf_file)
            resultados = []
            for idx, docnum in enumerate(documentos):
                registro = encontrar(docnum, registros)
                resultados.append({"documento": docnum, "empresa": registro["empresa"] if registro else "Não encontrado no PDF", "destinos": registro["destinos"][:] if registro else [], "encontrado": bool(registro), "confirmado": bool(registro)})
            st.session_state.resultados = resultados
            st.session_state.registros_pdf = registros
            st.session_state.manual_map = {}
            for i, r in enumerate(resultados):
                st.session_state[f"confirmar_{i}"] = bool(r["encontrado"])
        except Exception as e:
            st.error(f"Erro ao processar arquivos: {e}")

if st.session_state.get("resultados"):
    resultados = st.session_state.resultados
    registros_pdf = st.session_state.get("registros_pdf", {})
    encontrados = sum(1 for r in resultados if r["encontrado"])
    confirmados_atual = sum(bool(st.session_state.get(f"confirmar_{i}", r.get("confirmado", False))) for i, r in enumerate(resultados))
    pendentes = len(resultados) - confirmados_atual
    problemas = sum(1 for r in resultados if not r["encontrado"] or not r["destinos"])

    st.subheader("Resumo da conferência")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total de documentos", len(resultados))
    m2.metric("Encontrados", encontrados)
    m3.metric("Confirmados", confirmados_atual)
    m4.metric("Pendentes", pendentes)
    if problemas:
        st.warning(f"{problemas} documento(s) precisam de atenção: não encontrados ou sem destino positivo.")
    else:
        st.success("Todos os documentos foram encontrados com destino positivo.")

    b1, b2 = st.columns(2)
    with b1:
        if st.button("✅ Confirmar encontrados", use_container_width=True):
            for i, r in enumerate(resultados):
                if r["encontrado"] and r["destinos"]:
                    st.session_state[f"confirmar_{i}"] = True
            st.rerun()
    with b2:
        if st.button("☑️ Desmarcar todos", use_container_width=True):
            for i in range(len(resultados)):
                st.session_state[f"confirmar_{i}"] = False
            st.rerun()

    st.subheader("Conferência na ordem da planilha")
    for i, r in enumerate(resultados):
        with st.container(border=True):
            st.markdown(f"**Documento: {r['documento']}**")
            st.markdown(f"**Empresa:** {r['empresa']}")
            if r["destinos"]:
                st.markdown("**Baixado para:**")
                for destino in r["destinos"]:
                    st.write(f"- {destino}")
            else:
                st.warning("Destino não localizado com quantidade positiva. Verifique manualmente.")
            # Se não houve correspondência exata, permitir seleção manual de nota do PDF.
            if not r["encontrado"] and registros_pdf:
                opcoes = ["— Selecione uma nota do PDF —"] + [f"{v['nota']} | {v['empresa']}" for v in registros_pdf.values()]
                escolha = st.selectbox("Vincular manualmente a uma nota", opcoes, key=f"manual_{i}")
                if escolha != opcoes[0]:
                    nota_escolhida = escolha.split(" | ", 1)[0]
                    reg_escolhido = next((v for v in registros_pdf.values() if v["nota"] == nota_escolhida), None)
                    if reg_escolhido:
                        r["empresa"] = reg_escolhido["empresa"]
                        r["destinos"] = reg_escolhido["destinos"][:]
                        r["encontrado"] = True
                        r["manual"] = True
                        st.session_state.resultados[i] = r
                        st.success("Vínculo manual aplicado. Confirme o documento para incluí-lo no PDF final.")
            r["confirmado"] = st.checkbox("Confirmar", key=f"confirmar_{i}")
            st.session_state.resultados[i] = r

    confirmados = [r for i, r in enumerate(st.session_state.resultados) if st.session_state.get(f"confirmar_{i}", False)]
    st.metric("Documentos selecionados para o PDF", len(confirmados))
    if confirmados:
        pdf_saida = pdf_confirmados(confirmados)
        st.download_button("📄 Baixar PDF dos confirmados", data=pdf_saida, file_name="conferencia_baixas_confirmadas.pdf", mime="application/pdf", use_container_width=True)
    else:
        st.caption("Marque os documentos conferidos para habilitar o PDF final.")
