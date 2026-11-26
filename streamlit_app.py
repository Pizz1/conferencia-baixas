
        if "Nº da Nota" in linha or "N° da Nota" in linha:
            candidatos = re.findall(r"\b\d{3,}(?:/\d{1,3})?\b", linha)
            # O número da nota frequentemente está na linha seguinte e não na linha do rótulo.
            if i + 1 < len(linhas):
                candidatos += re.findall(r"\b\d{3,}(?:/\d{1,3})?\b", linhas[i + 1])
            com_barra = [c for c in candidatos if "/" in c]
            if com_barra:
                nota = com_barra[0]
            elif candidatos:
                nota = candidatos[0]
            if nota:
                chave = numeros(nota)
                registros.setdefault(chave, {"nota": nota, "empresa": fornecedor, "destinos": []})
                registros[chave]["empresa"] = fornecedor or registros[chave]["empresa"]

        # Ao encontrar Destino, lê o nome e a quantidade alocada logo abaixo.
        # Quantidade 0,00 é ignorada por completo.
        if linha.lower().startswith("destino") and nota:
            desc = linhas[i + 1] if i + 1 < len(linhas) else ""
            desc = re.sub(r"^(Estoque:|Centro de Custo:)", "", desc, flags=re.I).strip()
            qty = None
            for j in range(i + 2, min(i + 5, len(linhas))):
                # Quantidade costuma estar antes da unidade (CX, KG, SC etc.).
                m = re.match(r"^(-?\d+[.,]\d+)\s*(?:[A-Za-z]{1,6})?$", linhas[j])
                if m:
                    qty = quantidade_pt(m.group(1))
                    break
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
    candidatos = [r for chave, r in registros.items() if chave == alvo or chave.startswith(alvo)]
    # Se o número na planilha for o número-base da nota, aceita correspondência por prefixo.
    # Se houver mais de uma nota candidata, não escolhe arbitrariamente.
    if len(candidatos) == 1:
        return candidatos[0]
    if len(candidatos) > 1:
        exatos = [r for chave, r in registros.items() if chave == alvo]
        return exatos[0] if len(exatos) == 1 else None
    # Não usa correspondência parcial reversa para evitar associações incorretas.
    return None


def pdf_confirmados(resultados):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    estilos = getSampleStyleSheet()
    titulo = ParagraphStyle("Titulo", parent=estilos["Title"], fontSize=16, leading=20, alignment=1, spaceAfter=12)
    normal = ParagraphStyle("NormalNovo", parent=estilos["Normal"], fontSize=9, leading=12)
    story = [Paragraph("CONFERÊNCIA DE BAIXAS", titulo), Paragraph("Documentos confirmados na conferência", normal), Spacer(1, 12)]
    dados = [["Documento", "Empresa / Fornecedor", "Destino(s) da baixa"]]
    for r in resultados:
        dest = "<br/>".join(r["destinos"]) if r["destinos"] else "Destino não identificado no PDF"
        dados.append([Paragraph(str(r["documento"]), normal), Paragraph(r["empresa"] or "Não identificado", normal), Paragraph(dest, normal)])
    tabela = Table(dados, colWidths=[75, 190, 245], repeatRows=1)
    tabela.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#D9E2F3")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.black),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), 0.5, colors.grey),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING", (0,0), (-1,-1), 6), ("RIGHTPADDING", (0,0), (-1,-1), 6),
        ("TOPPADDING", (0,0), (-1,-1), 6), ("BOTTOMPADDING", (0,0), (-1,-1), 6),
    ]))
    story.append(tabela)
    doc.build(story)
    buffer.seek(0)
    return buffer


st.subheader("1. Envie os arquivos")
col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader("Remessa Excel", type=["xlsx", "xls"])
with col2:
    pdf_file = st.file_uploader("Relatório diário PDF", type=["pdf"])
st.info("A aba FILIAL da planilha define apenas a sequência dos documentos. Empresa e destino são extraídos do relatório PDF. Destinos com quantidade 0,00 são ignorados.")

if excel_file and pdf_file:
    if st.button("🔎 CONFRONTAR DOCUMENTOS", type="primary", use_container_width=True):
        try:
            documentos = ler_remessa(excel_file)
            registros = extrair_pdf(pdf_file)
            st.session_state.resultados = []
            for docnum in documentos:
                registro = encontrar(docnum, registros)
                st.session_state.resultados.append({
                    "documento": docnum,
                    "empresa": registro["empresa"] if registro else "Não encontrado no PDF",
                    "destinos": registro["destinos"][:] if registro else [],
                    "confirmado": False,
                    "encontrado": bool(registro)
                })
        except Exception as e:
            st.error(f"Erro ao processar arquivos: {e}")

if "resultados" in st.session_state and st.session_state.resultados:
    st.subheader("2. Conferência (ordem da planilha)")
    for i, r in enumerate(st.session_state.resultados):
        with st.container(border=True):
            st.markdown(f"**Documento: {r['documento']}**")
            st.markdown(f"**Empresa:** {r['empresa']}")
            if r["destinos"]:
                st.markdown("**Baixado para:**")
                for destino in r["destinos"]:
                    st.write(f"- {destino}")
            else:
                st.warning("Destino não localizado com quantidade positiva no PDF. Verifique manualmente.")
            r["confirmado"] = st.checkbox("Confirmar", key=f"confirmar_{i}")

    confirmados = [r for r in st.session_state.resultados if r["confirmado"]]
    st.metric("Documentos confirmados", len(confirmados))
    if confirmados:
        pdf_saida = pdf_confirmados(confirmados)
        st.download_button("📄 Baixar resumo em PDF (somente confirmados)", data=pdf_saida, file_name="conferencia_baixas_confirmadas.pdf", mime="application/pdf", use_container_width=True)
    else:
        st.caption("Marque os documentos conferidos para habilitar o PDF final.")
