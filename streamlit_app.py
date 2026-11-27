import streamlit as st
import pandas as pd
import fitz
import re
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

st.set_page_config(page_title="Conferência de Baixas", page_icon="📋", layout="wide")
st.markdown("""
<style>
.block-container {padding-top:1rem; padding-bottom:1rem; max-width:1400px}
h1 {font-size:1.65rem!important; margin-bottom:.15rem}
.smalltext {font-size: .82rem}
div[data-testid="stVerticalBlockBorderWrapper"] {border-radius:12px}
</style>
""", unsafe_allow_html=True)
st.title("📋 Conferência de Baixas")
st.caption("Os documentos seguem a ordem da planilha. As informações exibidas são extraídas do PDF.")

def norm_num(v):
    return re.sub(r"\D", "", str(v))

def read_excel(upload):
    df = pd.read_excel(upload, sheet_name="FILIAL", header=None, dtype=str)
    # Modelo simples: coluna B contém o número. Compatível também com número na coluna A.
    col = 1 if df.shape[1] > 1 and df.iloc[:, 1].notna().any() else 0
    docs = []
    for v in df.iloc[:, col].tolist():
        if pd.isna(v): continue
        v = str(v).strip()
        if not v or v.lower().startswith(("número do documento", "numero do documento", "documento")): continue
        if norm_num(v): docs.append(v)
    return docs

def pt_float(s):
    try: return float(str(s).replace(".", "").replace(",", "."))
    except: return None

def parse_report(upload):
    pdf = fitz.open(stream=upload.read(), filetype="pdf")
    lines = []
    for page in pdf:
        lines.extend([x.strip() for x in page.get_text().splitlines() if x.strip()])
    invoices = []
    current_invoice = None
    current_item = None
    i = 0
    while i < len(lines):
        line = lines[i]
        # fornecedor/note header
        if line.lower().startswith("fornecedor"):
            vendor = re.sub(r"^Fornecedor\s*:\s*", "", line, flags=re.I).strip()
            vendor = re.split(r"\s+Vencimento\s*:", vendor, flags=re.I)[0].strip()
            if vendor and vendor.lower() not in ("produto",):
                # header lines often contain supplier and due date together
                current_vendor = vendor
        if re.match(r"^N[º°o]\s*da\s*Nota\s*:", line, re.I):
            note = ""
            # Number sometimes appears on following line, combined with Valor
            for j in range(i+1, min(i+4, len(lines))):
                m = re.search(r"(\d{2,}(?:/\d{1,4})?)", lines[j])
                if m:
                    note = m.group(1); break
            if note:
                current_invoice = {"note": note, "vendor": current_vendor if "current_vendor" in locals() else "", "items": []}
                invoices.append(current_invoice)
                current_item = None
        # Product item rows: code, description, then four numeric rows (unit, qty, total, stock)
        if current_invoice and re.fullmatch(r"\d{4,}", line) and i+5 < len(lines):
            code = line
            desc = lines[i+1]
            unit_price = pt_float(lines[i+2])
            qty = pt_float(lines[i+3])
            total = pt_float(lines[i+4])
            stock = pt_float(lines[i+5])
            if None not in (unit_price, qty, total, stock):
                current_item = {"code":code, "description":desc, "unit_price":unit_price,
                                "quantity":qty, "total":total, "stock_value":stock, "destinations":[]}
                current_invoice["items"].append(current_item)
                i += 5
        # Destino line binds to most recently read item; multiple allocation blocks may repeat product lines
        if line.lower() == "destino" and current_invoice and current_item:
            destline = lines[i+1] if i+1 < len(lines) else ""
            dest = re.sub(r"^(Estoque:|Centro de Custo:)\s*", "", destline, flags=re.I).strip()
            allocation = None
            unit = ""
            if i+2 < len(lines):
                allocation = pt_float(lines[i+2])
            if i+3 < len(lines):
                unit = lines[i+3]
            if allocation is not None and allocation > 0 and dest:
                current_item["destinations"].append({"name":dest, "quantity":allocation, "unit":unit})
        i += 1
    return invoices

def invoice_base(note):
    return norm_num(str(note).split("/",1)[0])

def match_doc(doc, invoices):
    key = norm_num(str(doc).split("/",1)[0])
    matches = [x for x in invoices if key and invoice_base(x["note"]) == key]
    return matches[0] if len(matches)==1 else (matches if matches else None)

def money(v):
    return "R$ " + f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def make_pdf(rows):
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=28,rightMargin=28,topMargin=28,bottomMargin=28)
    styles = getSampleStyleSheet()
    small = ParagraphStyle("small", parent=styles["Normal"],fontSize=8.5,leading=11)
    head = ParagraphStyle("head",parent=styles["Heading2"],fontSize=11,leading=13)
    story=[Paragraph("CONFERÊNCIA DE BAIXAS — DOCUMENTOS CONFIRMADOS",styles["Title"]),Spacer(1,10)]
    for row in rows:
        elems=[Paragraph(f"<b>Documento: {row['documento']}</b> | Nota: {row.get('note','')}",head)]
        if row.get("vendor"): elems.append(Paragraph(f"<b>Fornecedor:</b> {row['vendor']}",small))
        data=[["Código","Produto","Qtd. documento","V. unitário","V. total","V. estoque"]]
        for item in row.get("items",[]):
            data.append([item["code"],Paragraph(item["description"],small),str(item["quantity"]),money(item["unit_price"]),money(item["total"]),money(item["stock_value"])])
        t=Table(data,colWidths=[45,180,65,65,65,65],repeatRows=1)
        t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#e3e9e6")),("GRID",(0,0),(-1,-1),.3,colors.grey),("VALIGN",(0,0),(-1,-1),"TOP"),("FONTSIZE",(0,0),(-1,-1),8),("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4),("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4)]))
        elems.append(t)
        for item in row.get("items",[]):
            for d in item["destinations"]:
                elems.append(Paragraph(f"<b>Baixado para:</b> {d['name']} — {d['quantity']:g} {d['unit']}",small))
        story.append(KeepTogether(elems)); story.append(Spacer(1,9))
    doc.build(story); buf.seek(0); return buf

def card_html(row, confirmed):
    border = "#16a34a" if confirmed else "#d6d9de"
    bg = "#eaf8ed" if confirmed else "#ffffff"
    return f"""<div style='border:2px solid {border};background:{bg};border-radius:12px;padding:12px 14px;margin:0 0 8px 0'>
      <div style='font-weight:700;font-size:1.05rem;color:{'#166534' if confirmed else '#111827'}'>{"✓ " if confirmed else ""}Documento: {row['documento']} <span style='font-weight:400;font-size:.9rem'>| Nota: {row.get('note','')}</span></div>
      <div style='font-size:.88rem;margin:4px 0'><b>Fornecedor:</b> {row.get('vendor','')}</div></div>"""

excel_col,pdf_col=st.columns(2)
with excel_col: xfile=st.file_uploader("Planilha Excel",type=["xlsx","xls"])
with pdf_col: pfile=st.file_uploader("Relatório PDF",type=["pdf"])
st.caption("O cartão apresenta os dados do documento e seus destinos. Destinos com quantidade 0,00 são ignorados.")

if xfile and pfile and st.button("🔎 CONFERIR DOCUMENTOS",type="primary",use_container_width=True):
    try:
        docs=read_excel(xfile)
        invoices=parse_report(pfile)
        result=[]
        for d in docs:
            m=match_doc(d,invoices)
            if isinstance(m,list):
                inv=None # ambiguous match; don't auto-pick
                status="ambiguous"
            else:
                inv=m
                status="found" if inv else "missing"
            items=inv["items"] if inv else []
            result.append({"documento":d,"note":inv["note"] if inv else "", "vendor":inv["vendor"] if inv else "",
                           "items":items, "encontrado":bool(inv), "status":status,
                           "confirmed":bool(inv and any(it["destinations"] for it in items))})
        st.session_state["records"]=result
        for idx,r in enumerate(result):
            st.session_state[f"confirm_{idx}"]=r["confirmed"]
    except Exception as e:
        st.error(f"Erro ao ler os arquivos: {e}")

if st.session_state.get("records"):
    records=st.session_state["records"]
    total=len(records)
    confirmed=sum(st.session_state.get(f"confirm_{i}",False) for i in range(total))
    missing=sum(not r["encontrado"] for r in records)
    a,b,c=st.columns(3)
    a.metric("Documentos",total); b.metric("Confirmados",confirmed); c.metric("Pendentes",total-confirmed)
    left,right=st.columns([1,1])
    with left:
        if st.button("✓ Confirmar encontrados",use_container_width=True):
            for i,r in enumerate(records): st.session_state[f"confirm_{i}"]=bool(r["encontrado"] and any(it["destinations"] for it in r["items"]))
            st.rerun()
    with right:
        if st.button("↺ Desmarcar todos",use_container_width=True):
            for i in range(total): st.session_state[f"confirm_{i}"]=False
            st.rerun()
    st.markdown("### Documentos — sequência da planilha")
    for i,r in enumerate(records):
        confirmed=st.session_state.get(f"confirm_{i}",False)
        with st.container(border=True):
            st.markdown(card_html(r,confirmed),unsafe_allow_html=True)
            if not r["encontrado"]:
                st.warning("Documento não encontrado de forma automática no PDF.")
            elif not r["items"]:
                st.warning("Nota encontrada, mas sem produtos extraídos.")
            else:
                for item in r["items"]:
                    tabledata=[["Código","Produto","Quantidade","V. unitário","V. total","V. estoque"],
                        [item["code"],item["description"],str(item["quantity"]),money(item["unit_price"]),money(item["total"]),money(item["stock_value"])]]
                    t=st.dataframe(pd.DataFrame([{"Código":item["code"],"Produto":item["description"],"Quantidade":item["quantity"],"Valor unitário":money(item["unit_price"]),"Valor total":money(item["total"]),"Valor estoque":money(item["stock_value"])}]),hide_index=True,use_container_width=True)
                    if item["destinations"]:
                        for d in item["destinations"]:
                            st.markdown(f"📍 **Baixado para:** {d['name']} — **{d['quantity']:g} {d['unit']}**")
                    else:
                        st.caption("Nenhum destino com quantidade positiva identificado para este item.")
            # Checkbox is retained for reliability; styling makes confirmed cards green.
            st.checkbox("Confirmar documento (clique para marcar/desmarcar)",key=f"confirm_{i}")
    selected=[r for i,r in enumerate(records) if st.session_state.get(f"confirm_{i}",False)]
    if selected:
        st.download_button("📄 Gerar PDF dos confirmados",data=make_pdf(selected),file_name="conferencia_baixas_confirmadas.pdf",mime="application/pdf",use_container_width=True)
