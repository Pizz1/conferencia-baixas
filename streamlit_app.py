import streamlit as st
import pandas as pd
import fitz
import re
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

st.set_page_config(page_title="Conferência de Baixas", page_icon="📋", layout="wide")
st.title("📋 Conferência de Baixas")
st.caption("Confira somente o destino das baixas, na sequência dos documentos da planilha.")

def digits(s):
    return re.sub(r"\D", "", str(s))

def clean(s):
    return str(s).strip()

def read_excel(upload):
    df = pd.read_excel(upload, sheet_name="FILIAL", header=None, dtype=str)
    # Modelo: coluna B (index 1). Compatibilidade com planilhas com números na coluna A.
    col = 1 if df.shape[1] > 1 and df.iloc[:,1].notna().any() else 0
    docs=[]
    for value in df.iloc[:,col].tolist():
        if pd.isna(value): continue
        value=clean(value)
        if not digits(value): continue
        if value.lower() in ("numero do documento","número do documento","documento","sequência","sequencia"): continue
        docs.append(value)
    return docs

def parse_pt_num(s):
    s=str(s).strip().replace(".","").replace(",",".")
    try: return float(s)
    except: return None

def parse_pdf(upload):
    pdf=fitz.open(stream=upload.read(),filetype="pdf")
    lines=[]
    for page in pdf:
        lines.extend([clean(x) for x in page.get_text().splitlines() if clean(x)])
    records=[]
    records_by_note={}
    current=None
    i=0
    # A linha do campo Nº da Nota é seguida, em geral, pela linha com o número/valor.
    while i<len(lines):
        line=lines[i]
        if re.search(r"N[º°o]\s*da\s*Nota\s*:", line, re.I) or line.strip().lower()=="nº da nota :":
            candidate=""
            # Procura nas linhas imediatamente seguintes pelo número da nota (com sufixo /xxx)
            for j in range(i+1,min(i+5,len(lines))):
                m=re.search(r"\b(\d{2,}(?:/\d{1,4})?)\b",lines[j])
                if m:
                    candidate=m.group(1); break
            if candidate:
                current=records_by_note.setdefault(candidate,{"nota":candidate,"destinos":[]})
                if current not in records: records.append(current)
        # algumas extrações trazem o campo e o valor na mesma linha
        m_inline=re.search(r"N[º°o]\s*da\s*Nota\s*:\s*(\d{2,}(?:/\d{1,4})?)",line,re.I)
        if m_inline:
            current=records_by_note.setdefault(m_inline.group(1),{"nota":m_inline.group(1),"destinos":[]})
            if current not in records: records.append(current)
        if line.lower()=="destino" and current:
            destline=lines[i+1] if i+1<len(lines) else ""
            destline=re.sub(r"^(Estoque:|Centro de Custo:)\s*","",destline,flags=re.I).strip()
            qty=None
            # O número da alocação aparece logo após a descrição do destino,
            # normalmente antes da unidade (KG, CX, UN etc.).
            for j in range(i+2,min(i+5,len(lines))):
                m=re.match(r"^\s*(-?\d[\d.,]*)\s*(?:[A-Za-z]{1,8})?\s*$",lines[j])
                if m:
                    qty=parse_pt_num(m.group(1)); break
            if destline and qty is not None and qty>0:
                entry=f"{destline} — {qty:g}"
                if entry not in current["destinos"]:
                    current["destinos"].append(entry)
        i+=1
    return records

def match_doc(doc, records):
    # Excel pode conter 16789, enquanto PDF contém 16789/001.
    # Comparar o número-base (parte antes da barra) de forma exata.
    target=digits(str(doc).split("/",1)[0])
    matches=[]
    for r in records:
        base=digits(r["nota"].split("/",1)[0])
        if target and base==target:
            matches.append(r)
    # Se houver mais de uma nota com o mesmo número-base, não escolher silenciosamente.
    return matches[0] if len(matches)==1 else None

def make_pdf(rows):
    buf=BytesIO()
    pdf=SimpleDocTemplate(buf,pagesize=A4,rightMargin=36,leftMargin=36,topMargin=36,bottomMargin=36)
    styles=getSampleStyleSheet()
    body=ParagraphStyle("body",parent=styles["Normal"],fontSize=10,leading=14)
    title=ParagraphStyle("title2",parent=styles["Title"],fontSize=16,leading=20,alignment=1)
    story=[Paragraph("CONFERÊNCIA DE BAIXAS",title),Spacer(1,12)]
    data=[["Documento","Destino(s) da baixa"]]
    for r in rows:
        d="<br/>".join(r["destinos"]) if r["destinos"] else "Destino não identificado"
        data.append([Paragraph(str(r["documento"]),body),Paragraph(d,body)])
    table=Table(data,colWidths=[100,410],repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#D9E2F3")),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("GRID",(0,0),(-1,-1),0.5,colors.grey),("VALIGN",(0,0),(-1,-1),"TOP"),
        ("LEFTPADDING",(0,0),(-1,-1),6),("RIGHTPADDING",(0,0),(-1,-1),6),
        ("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6)
    ]))
    story.append(table); pdf.build(story); buf.seek(0)
    return buf

excel_col,pdf_col=st.columns(2)
with excel_col: xfile=st.file_uploader("Planilha Excel",type=["xlsx","xls"])
with pdf_col: pfile=st.file_uploader("Relatório de baixas PDF",type=["pdf"])
st.info("O aplicativo segue a sequência do Excel. Busca o número-base da nota no PDF, exibe apenas destinos com quantidade positiva e não mostra fornecedor.")

if xfile and pfile and st.button("🔎 CONFRONTAR",type="primary",use_container_width=True):
    try:
        docs=read_excel(xfile)
        records=parse_pdf(pfile)
        rows=[]
        for docnum in docs:
            reg=match_doc(docnum,records)
            rows.append({"documento":docnum,"destinos":reg["destinos"][:] if reg else [],
                         "encontrado":bool(reg),"confirmado":bool(reg and reg["destinos"])})
        st.session_state["rows"]=rows
    except Exception as e:
        st.error(f"Erro na leitura dos arquivos: {e}")

if st.session_state.get("rows"):
    rows=st.session_state["rows"]
    total=len(rows); found=sum(r["encontrado"] for r in rows)
    confirmed=sum(bool(st.session_state.get(f"confirm_{i}",r["confirmado"])) for i,r in enumerate(rows))
    pending=total-confirmed
    issues=sum(not r["encontrado"] or not r["destinos"] for r in rows)
    st.subheader("Progresso")
    c1,c2,c3,c4=st.columns(4)
    c1.metric("Total",total); c2.metric("Encontrados",found); c3.metric("Confirmados",confirmed); c4.metric("Pendentes",pending)
    if issues: st.warning(f"{issues} documento(s) não encontrados ou sem destino positivo.")
    left,right=st.columns(2)
    with left:
        if st.button("✅ Confirmar encontrados",use_container_width=True):
            for i,r in enumerate(rows): st.session_state[f"confirm_{i}"]=bool(r["encontrado"] and r["destinos"])
            st.rerun()
    with right:
        if st.button("☑️ Desmarcar todos",use_container_width=True):
            for i in range(total): st.session_state[f"confirm_{i}"]=False
            st.rerun()
    st.subheader("Conferência — ordem da planilha")
    for i,r in enumerate(rows):
        with st.container(border=True):
            st.markdown(f"**Documento: {r['documento']}**")
            if r["destinos"]:
                st.markdown("**Baixado para:**")
                for d in r["destinos"]: st.write(f"- {d}")
            elif not r["encontrado"]:
                st.warning("Documento não localizado automaticamente no PDF.")
            else:
                st.warning("Documento encontrado, mas não há destino com quantidade positiva.")
            st.checkbox("Confirmar",key=f"confirm_{i}",value=bool(st.session_state.get(f"confirm_{i}",r["confirmado"])))
    selected=[]
    for i,r in enumerate(rows):
        if st.session_state.get(f"confirm_{i}",False):
            selected.append(r)
    if selected:
        st.download_button("📄 Baixar PDF dos confirmados",data=make_pdf(selected),
                          file_name="conferencia_baixas_confirmadas.pdf",mime="application/pdf",use_container_width=True)
