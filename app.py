import streamlit as st
import io, re
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pathlib import Path

st.set_page_config(page_title="Statement PDF → Excel", page_icon="📄", layout="centered")

st.markdown("""
<style>
.hero{background:linear-gradient(135deg,#1e3a5f,#2563eb);border-radius:16px;
      padding:36px 32px 28px;margin-bottom:28px;color:white;text-align:center}
.hero h1{font-size:2rem;margin:0 0 8px;font-weight:700}
.hero p{font-size:1rem;opacity:.85;margin:0}
.badge{display:inline-block;background:rgba(255,255,255,.18);border-radius:20px;
       padding:3px 14px;font-size:.8rem;margin:10px 4px 0}
.ok{background:#f0fdf4;border:1px solid #86efac;border-radius:10px;
    padding:16px 20px;color:#166534;font-weight:500}
.stDownloadButton>button{background:#2563eb!important;color:white!important;
    border-radius:8px!important;padding:10px 24px!important;font-size:1rem!important;
    font-weight:600!important;width:100%!important;border:none!important}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
  <h1>📄 Statement → Excel</h1>
  <p>Upload a customer statement PDF and download a clean formatted Excel file instantly.</p>
  <span class="badge">✅ Savannah Cement</span>
  <span class="badge">✅ National Cement</span>
  <span class="badge">🔒 Files never stored</span>
</div>
""", unsafe_allow_html=True)

with st.expander("📋 Supported formats"):
    st.markdown("""
**🟠 Savannah Cement** — Doc No, Ref/Vehicle, Date, Due Date, Details, Debit, Credit, Balance

**🔵 National Cement (Rhino/Simba)** — Date, Description, Vehicle, LPO, Qty, Invoice, CU Invoice, Location, Debit, Credit, Balance + separate **Cheques on Hand** sheet
    """)

# ── Shared constants ───────────────────────────────────────────────────────────
C = dict(savannah_orange="E8540A", national_blue="1565C0", header_dark="2C3E50",
         col_hdr_fg="FFFFFF", alt_row="F2F6FA", credit_row="E8F5E9", info_bg="ECF0F1",
         debit_fg="1565C0", credit_fg="2E7D32", neg_bal="C62828", border="BDBDBD")
MFMT = '#,##0.00;(#,##0.00);"-"'

def thin(sides="all"):
    s = Side(style="thin", color=C["border"])
    kw = dict(left=s,right=s,top=s,bottom=s) if sides=="all" else dict(bottom=s)
    return Border(**kw)

def pn(s):
    try: return float(str(s).strip().replace(",",""))
    except: return None

def wc(ws, r, c, v=None, bold=False, fg="000000", bg=None, align="left",
       fmt=None, bdr=None, sz=9, ind=0):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = Font(name="Arial", size=sz, bold=bold, color=fg)
    if bg: cell.fill = PatternFill("solid", start_color=bg)
    cell.alignment = Alignment(horizontal=align, vertical="center", indent=ind)
    if fmt: cell.number_format = fmt
    if bdr: cell.border = bdr
    return cell

# ── Detection ──────────────────────────────────────────────────────────────────
def detect(lines):
    t = "\n".join(lines[:40])
    if "NATIONAL CEMENT" in t.upper() or "CU Invoice Number" in t: return "national"
    if "Savannah Cement" in t or "A/R Invoices" in t:              return "savannah"
    if "LPO No" in t or "Cheque no" in t:                          return "national"
    if "Due Date" in t and "Document" in t:                        return "savannah"
    return "unknown"

# ── National parser ────────────────────────────────────────────────────────────
DR4 = re.compile(r"^\d{2}\.\d{2}\.\d{4}$")
MR  = re.compile(r"^-?[\d,]+\.\d{2,3}$")
QR  = re.compile(r"^\d+\.\d{3}$")

def parse_national(lines):
    hdr = {}
    for ln in lines[:20]:
        m = re.search(r"Statement As At:\s*([\d.]+)", ln)
        if m: hdr["statement_date"] = m.group(1)
        m = re.search(r"PIN No:\s*(\S+)", ln)
        if m: hdr["pin"] = m.group(1)
    for ln in lines[:5]:
        if "CEMENT" in ln.upper() and "LTD" in ln.upper():
            hdr["supplier"] = ln.strip().split("  ")[0].strip(); break
    hdr.setdefault("supplier", "National Cement Company Ltd")
    for i, ln in enumerate(lines):
        if "Customer Statement" in ln and i > 0:
            hdr["customer"] = lines[i-1].strip(); break
    hdr.setdefault("customer", "")

    txs, chqs, in_chq = [], [], False
    i = 0
    while i < len(lines):
        ln = lines[i].strip()
        if "Cheques on Hand" in ln:  in_chq=True;  i+=1; continue
        if "Cheque No" in ln and "Due On" in ln:    i+=1; continue
        if in_chq:
            p = ln.split()
            if len(p)>=3:
                try:
                    amt=pn(p[-1]); dt=p[-2] if DR4.match(p[-2]) else ""
                    if amt: chqs.append({"cheque_no":p[0],"due_on":dt,"amount":amt})
                except: pass
            i+=1; continue
        if "Total" in ln and "Closing" in ln: i+=1; continue
        p = ln.split()
        if not p or not DR4.match(p[0]): i+=1; continue

        date=p[0]; rest=p[1:]
        j=i+1
        while j<len(lines):
            nx=lines[j].strip()
            if not nx or DR4.match(nx.split()[0] if nx.split() else "") \
               or any(k in nx for k in ("Total","Ageing","Cheque")): break
            if re.match(r"^\d{5,}", nx) and rest: rest[-1]+=nx
            j+=1
        i=j

        amts, nonamts, dc = [], [], ""
        for x in reversed(rest):
            if x in ("DR","CR"): dc=x
            elif MR.match(x.replace(",","")): amts.insert(0, pn(x))
            else:
                nonamts.insert(0,x)
                if len(amts)>=3: break

        na=nonamts
        desc    = na[0] if len(na)>0 else ""
        vehicle = na[1] if len(na)>1 else ""
        lpo     = na[2] if len(na)>2 else ""
        qty     = pn(na[3]) if len(na)>3 and QR.match(na[3]) else None
        inv     = na[4] if len(na)>4 else ""
        cu      = na[5] if len(na)>5 else ""
        loc     = na[6] if len(na)>6 else ""
        curr    = na[7] if len(na)>7 else ""

        dbt=crd=bal=None
        if len(amts)>=3: dbt,crd,bal=amts[-3],amts[-2],amts[-1]
        elif len(amts)==2: dbt,bal=amts[0],amts[1]
        elif len(amts)==1: bal=amts[0]

        if desc=="Balance": desc="Balance B/F"
        txs.append({"date":date,"description":desc,"cheque_no":"","vehicle":vehicle,
                    "lpo_no":lpo,"quantity":qty,"invoice_no":inv,"cu_invoice":cu,
                    "location":loc,"currency":curr,"debit":dbt,"credit":crd,"balance":bal})
    return hdr, txs, chqs

# ── Savannah parser ────────────────────────────────────────────────────────────
DR2 = re.compile(r"^\d{2}\.\d{2}\.\d{2}$")
MR2 = re.compile(r"^-?[\d,]+\.\d{2}$")
PFX = ("IN ","CN ","RC ","DN ")

def parse_savannah(lines):
    hdr = {"customer_name":"New Muthokinju Hardware Ltd","customer_code":"LC000072",
           "credit_limit":2_000_000.00,"sales_employee":"Fred Oguttu",
           "contact_person":"David","currency":"KES","printed_on":""}
    for ln in lines:
        m = re.search(r"Printed On:\s*\S+\s+(\S+)", ln)
        if m: hdr["printed_on"]=m.group(1)

    txs=[]; i=0
    while i<len(lines):
        ln=lines[i].strip()
        if not any(ln.startswith(p) for p in PFX): i+=1; continue
        p=ln.split(); dt=p[0]; dn=p[1] if len(p)>1 else ""; rest=p[2:]
        di=[j for j,x in enumerate(rest) if DR2.match(x)]
        mi=[j for j,x in enumerate(rest) if MR2.match(x)]
        if len(di)>=2:
            rn=" ".join(rest[:di[0]]); d=rest[di[0]]; dd=rest[di[1]]; af=di[1]+1
            if mi:
                fm=min(m for m in mi if m>di[1]); dps=rest[af:fm]; amts=rest[fm:]
            else: dps=rest[af:]; amts=[]
        else: rn=d=dd=""; dps=rest; amts=[]

        rex=[]; j=i+1
        while j<len(lines):
            nx=lines[j].strip()
            if any(nx.startswith(x) for x in PFX): break
            if nx.startswith("Page:") or nx.startswith("Total") or nx.startswith("Balance Due"): break
            if DR2.match(nx) or MR2.match(nx): break
            if any(k in nx for k in ("Savannah","Customer Statement","Currency:","Credit Limit:",
                "Sales Employee:","Contact Person:","Address:","Phone #:",
                "Document Ref.","Prior Period","Debit Credit")): break
            rex.append(nx); j+=1
        i=j

        dbt=crd=bal=None
        if len(amts)==2:
            if dt in ("CN","RC"): crd,bal=pn(amts[0]),pn(amts[1])
            else: dbt,bal=pn(amts[0]),pn(amts[1])
        elif len(amts)==3: dbt,crd,bal=pn(amts[0]),pn(amts[1]),pn(amts[2])
        elif len(amts)==1: bal=pn(amts[0])

        txs.append({"doc_type":dt,"doc_no":dn,"ref_no":(" ".join([rn]+rex)).strip(),
                    "date":d,"due_date":dd,"details":" ".join(dps).strip() or "A/R Invoices",
                    "debit":dbt,"credit":crd,"balance":bal})
    return hdr, txs

# ── Excel: Savannah ────────────────────────────────────────────────────────────
def excel_savannah(hdr, txs):
    wb=openpyxl.Workbook(); ws=wb.active; ws.title="Statement"
    ws.sheet_properties.tabColor=C["savannah_orange"]
    for i,w in enumerate([6,8,30,11,11,32,16,16,18],1):
        ws.column_dimensions[get_column_letter(i)].width=w

    ws.merge_cells("A1:I1")
    wc(ws,1,1,"SAVANNAH CEMENT 2025 LIMITED  |  CUSTOMER STATEMENT",
       bold=True,fg=C["col_hdr_fg"],bg=C["header_dark"],align="center",sz=14)
    ws.row_dimensions[1].height=30

    lf=Font(name="Arial",bold=True,size=9); vf=Font(name="Arial",size=9)
    lb=PatternFill("solid",start_color=C["info_bg"])
    al=Alignment(horizontal="left",vertical="center",indent=1)
    info=[("Customer Name",hdr["customer_name"],"Currency",hdr["currency"]),
          ("Customer Code",hdr["customer_code"],"Credit Limit",f"KES {hdr['credit_limit']:,.2f}"),
          ("Sales Employee",hdr["sales_employee"],"Contact",hdr["contact_person"]),
          ("Printed On",hdr["printed_on"],"","")]
    for r,(l1,v1,l2,v2) in enumerate(info,start=2):
        for ci,val,bold in [(1,l1,True),(3,v1,False),(6,l2,True),(8,v2,False)]:
            ws.merge_cells(start_row=r,start_column=ci,end_row=r,end_column=ci+1)
            c=ws.cell(row=r,column=ci,value=val)
            c.font=lf if bold else vf; c.alignment=al
            if bold: c.fill=lb
        ws.row_dimensions[r].height=16

    H=6
    for ci,h in enumerate(["Type","Doc No.","Ref/Vehicle","Date","Due Date",
                            "Details","Debit (KES)","Credit (KES)","Balance (KES)"],1):
        wc(ws,H,ci,h,bold=True,fg=C["col_hdr_fg"],bg=C["savannah_orange"],align="center",bdr=thin())
    ws.row_dimensions[H].height=22

    D=H+1
    for ri,tx in enumerate(txs,start=D):
        cr=tx["doc_type"] in ("CN","RC")
        bg=C["credit_row"] if cr else (C["alt_row"] if (ri-D)%2==1 else None)
        for ci,(v,a,f,fg) in enumerate(zip(
            [tx["doc_type"],tx["doc_no"],tx["ref_no"],tx["date"],tx["due_date"],
             tx["details"],tx["debit"],tx["credit"],tx["balance"]],
            ["center","center","left","center","center","left","right","right","right"],
            [None,None,None,None,None,None,MFMT,MFMT,MFMT],
            ["000000","000000","000000","000000","000000","000000",
             C["debit_fg"],C["credit_fg"],"000000"]),1):
            if ci==9 and v is not None: fg=C["neg_bal"] if v<0 else C["credit_fg"]
            wc(ws,ri,ci,v,align=a,fmt=f,fg=fg,bg=bg,bdr=thin("bottom"))
        ws.row_dimensions[ri].height=15

    T=D+len(txs)
    ws.merge_cells(f"A{T}:F{T}")
    wc(ws,T,1,"TOTAL",bold=True,fg=C["col_hdr_fg"],bg=C["header_dark"],align="right",bdr=thin(),ind=1)
    for ci,col in [(7,"G"),(8,"H"),(9,"I")]:
        c=ws.cell(row=T,column=ci,value=f"=SUM({col}{D}:{col}{T-1})")
        c.font=Font(name="Arial",bold=True,size=9,color=C["col_hdr_fg"])
        c.fill=PatternFill("solid",start_color=C["header_dark"])
        c.number_format=MFMT; c.alignment=Alignment(horizontal="right",vertical="center")
        c.border=thin()
    ws.row_dimensions[T].height=18
    ws.freeze_panes=f"A{D}"; ws.auto_filter.ref=f"A{H}:I{T-1}"
    ws.page_setup.orientation="landscape"; ws.page_setup.fitToPage=True; ws.page_setup.fitToWidth=1
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf

# ── Excel: National ────────────────────────────────────────────────────────────
def excel_national(hdr, txs, chqs):
    wb=openpyxl.Workbook(); ws=wb.active; ws.title="Statement"
    ws.sheet_properties.tabColor=C["national_blue"]
    for i,w in enumerate([12,12,14,20,18,10,14,26,12,6,16,16,18],1):
        ws.column_dimensions[get_column_letter(i)].width=w

    ws.merge_cells("A1:M1")
    wc(ws,1,1,"NATIONAL CEMENT COMPANY LTD  |  CUSTOMER STATEMENT",
       bold=True,fg=C["col_hdr_fg"],bg=C["national_blue"],align="center",sz=14)
    ws.row_dimensions[1].height=30

    lf=Font(name="Arial",bold=True,size=9); vf=Font(name="Arial",size=9)
    lb=PatternFill("solid",start_color=C["info_bg"])
    al=Alignment(horizontal="left",vertical="center",indent=1)
    info=[("Customer",hdr.get("customer",""),"Supplier",hdr.get("supplier","")),
          ("Statement Date",hdr.get("statement_date",""),"PIN No",hdr.get("pin",""))]
    for r,(l1,v1,l2,v2) in enumerate(info,start=2):
        for ci,lbl,val in [(1,l1,v1),(8,l2,v2)]:
            ws.merge_cells(start_row=r,start_column=ci,end_row=r,end_column=ci+2)
            ws.merge_cells(start_row=r,start_column=ci+3,end_row=r,end_column=ci+5)
            lc=ws.cell(row=r,column=ci,value=lbl); lc.font=lf; lc.fill=lb; lc.alignment=al
            vc=ws.cell(row=r,column=ci+3,value=val); vc.font=vf; vc.alignment=al
        ws.row_dimensions[r].height=16

    H=4
    for ci,h in enumerate(["Date","Description","Cheque No","Vehicle No.","LPO No",
                            "Qty","Invoice No.","CU Invoice No.","Location","Curr.",
                            "Debit (KES)","Credit (KES)","Balance (KES)"],1):
        wc(ws,H,ci,h,bold=True,fg=C["col_hdr_fg"],bg=C["national_blue"],align="center",bdr=thin())
    ws.row_dimensions[H].height=22

    D=H+1
    for ri,tx in enumerate(txs,start=D):
        bf="B/F" in str(tx.get("description",""))
        bg=C["info_bg"] if bf else (C["alt_row"] if (ri-D)%2==1 else None)
        vals=[tx["date"],tx["description"],tx["cheque_no"],tx["vehicle"],tx["lpo_no"],
              tx["quantity"],tx["invoice_no"],tx["cu_invoice"],tx["location"],tx["currency"],
              tx["debit"],tx["credit"],tx["balance"]]
        aligns=["center","left","center","center","center","right","center","center",
                "center","center","right","right","right"]
        fmts=[None,None,None,None,None,"#,##0.000",None,None,None,None,MFMT,MFMT,MFMT]
        fgs=["000000"]*10+[C["debit_fg"],C["credit_fg"],"000000"]
        for ci,(v,a,f,fg) in enumerate(zip(vals,aligns,fmts,fgs),1):
            if ci==13 and v is not None: fg=C["neg_bal"] if v<0 else C["credit_fg"]
            wc(ws,ri,ci,v,align=a,fmt=f,fg=fg,bg=bg,bdr=thin("bottom"))
        ws.row_dimensions[ri].height=15

    T=D+len(txs)
    ws.merge_cells(f"A{T}:J{T}")
    wc(ws,T,1,"TOTAL / CLOSING BALANCE",bold=True,fg=C["col_hdr_fg"],
       bg=C["national_blue"],align="right",bdr=thin(),ind=1)
    for ci,col in [(11,"K"),(12,"L"),(13,"M")]:
        c=ws.cell(row=T,column=ci,value=f"=SUM({col}{D}:{col}{T-1})")
        c.font=Font(name="Arial",bold=True,size=9,color=C["col_hdr_fg"])
        c.fill=PatternFill("solid",start_color=C["national_blue"])
        c.number_format=MFMT; c.alignment=Alignment(horizontal="right",vertical="center")
        c.border=thin()
    ws.row_dimensions[T].height=18
    ws.freeze_panes=f"A{D}"; ws.auto_filter.ref=f"A{H}:M{T-1}"
    ws.page_setup.orientation="landscape"; ws.page_setup.fitToPage=True; ws.page_setup.fitToWidth=1

    if chqs:
        ws2=wb.create_sheet("Cheques on Hand"); ws2.sheet_properties.tabColor="2E7D32"
        for col,w in [("A",16),("B",16),("C",18)]: ws2.column_dimensions[col].width=w
        ws2.merge_cells("A1:C1")
        wc(ws2,1,1,"CHEQUES ON HAND",bold=True,fg=C["col_hdr_fg"],bg="2E7D32",align="center",sz=12)
        ws2.row_dimensions[1].height=26
        for ci,h in enumerate(["Cheque No","Due On","Amount (KES)"],1):
            wc(ws2,2,ci,h,bold=True,fg=C["col_hdr_fg"],bg=C["national_blue"],align="center",bdr=thin())
        ws2.row_dimensions[2].height=18
        for ri,chq in enumerate(chqs,start=3):
            bg=C["alt_row"] if (ri-3)%2==1 else None
            wc(ws2,ri,1,chq["cheque_no"],align="center",bg=bg,bdr=thin("bottom"))
            wc(ws2,ri,2,chq["due_on"],   align="center",bg=bg,bdr=thin("bottom"))
            wc(ws2,ri,3,chq["amount"],   align="right", bg=bg,bdr=thin("bottom"),
               fmt=MFMT,fg=C["national_blue"])
            ws2.row_dimensions[ri].height=15
        CR=3+len(chqs)
        ws2.merge_cells(f"A{CR}:B{CR}")
        wc(ws2,CR,1,"TOTAL",bold=True,fg=C["col_hdr_fg"],bg=C["national_blue"],align="right",bdr=thin())
        c=ws2.cell(row=CR,column=3,value=f"=SUM(C3:C{CR-1})")
        c.font=Font(name="Arial",bold=True,size=9,color=C["col_hdr_fg"])
        c.fill=PatternFill("solid",start_color=C["national_blue"])
        c.number_format=MFMT; c.alignment=Alignment(horizontal="right",vertical="center"); c.border=thin()
        ws2.row_dimensions[CR].height=18

    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf

# ── Main convert ───────────────────────────────────────────────────────────────
def convert(pdf_bytes):
    lines=[]
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        pages=len(pdf.pages)
        for page in pdf.pages:
            t=page.extract_text()
            if t: lines.extend(t.split("\n"))
    fmt=detect(lines)
    if fmt=="national":
        hdr,txs,chqs=parse_national(lines)
        return fmt,len(txs),len(chqs),pages,excel_national(hdr,txs,chqs)
    elif fmt=="savannah":
        hdr,txs=parse_savannah(lines)
        return fmt,len(txs),0,pages,excel_savannah(hdr,txs)
    return "unknown",0,0,pages,None

# ── UI ─────────────────────────────────────────────────────────────────────────
uploaded = st.file_uploader("Drop your PDF here", type=["pdf","PDF"],
                             label_visibility="collapsed")

if uploaded:
    st.markdown(f"**Uploaded:** `{uploaded.name}` &nbsp;({uploaded.size/1024:.1f} KB)")
    with st.spinner("Converting…"):
        try:
            fmt, tx_count, chq_count, pages, buf = convert(uploaded.read())
        except Exception as e:
            st.error(f"❌ Error: {e}"); st.stop()

    if fmt=="unknown":
        st.warning("⚠️ Format not recognised. Supported: Savannah Cement, National Cement.")
    else:
        label = "Savannah Cement" if fmt=="savannah" else "National Cement"
        chq_txt = f" · {chq_count} cheques" if chq_count else ""
        st.markdown(f"""
        <div class="ok">✅ &nbsp;<b>{label}</b> converted &nbsp;·&nbsp;
        {pages} page(s) &nbsp;·&nbsp; {tx_count} transactions{chq_txt}</div>
        """, unsafe_allow_html=True)
        st.write("")
        st.download_button("⬇️  Download Excel File", data=buf,
            file_name=Path(uploaded.name).stem+"_converted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.markdown("---")
st.markdown("<div style='text-align:center;color:#94a3b8;font-size:.8rem'>"
            "Files processed in memory · never stored · free to use</div>",
            unsafe_allow_html=True)
