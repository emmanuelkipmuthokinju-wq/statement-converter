import streamlit as st
import io, re, json, os
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pathlib import Path
st.set_page_config(page_title="Statement PDF → Excel", page_icon="📄", layout="centered")
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;600;700;800&display=swap');
html, body, [class*="css"] { font-family:'Poppins',sans-serif; }
.stApp { background: linear-gradient(160deg,#fdf2ff 0%,#eef4ff 45%,#fff8ec 100%); }
.hero{
  background: linear-gradient(120deg,#ff5f6d 0%,#a06be0 35%,#3b82f6 70%,#06b6d4 100%);
  background-size: 200% 200%;
  animation: gradientshift 8s ease infinite;
  border-radius: 24px;
  padding: 42px 32px 34px;
  margin-bottom: 26px;
  color: white;
  text-align: center;
  box-shadow: 0 18px 40px -12px rgba(124,58,237,.45);
}
@keyframes gradientshift{0%{background-position:0% 50%}50%{background-position:100% 50%}100%{background-position:0% 50%}}
.hero h1{font-size:2.3rem;margin:0 0 10px;font-weight:800;letter-spacing:-.5px;
  text-shadow:0 2px 12px rgba(0,0,0,.15)}
.hero p{font-size:1.05rem;opacity:.95;margin:0 0 4px;font-weight:400}
.badge{display:inline-block;background:rgba(255,255,255,.22);backdrop-filter:blur(6px);
       border:1px solid rgba(255,255,255,.35);border-radius:24px;
       padding:5px 16px;font-size:.82rem;font-weight:600;margin:10px 5px 0}
.ok{
  background: linear-gradient(120deg,#d1fae5,#ecfeff);
  border: 1.5px solid #34d399;
  border-radius: 16px;
  padding: 18px 22px;
  color:#065f46;
  font-weight:600;
  margin-top:14px;
  box-shadow: 0 6px 18px -8px rgba(16,185,129,.35);
}
div[data-testid="stExpander"]{
  border-radius:16px!important; border:1.5px solid #e9d5ff!important;
  background:white!important; box-shadow:0 4px 14px -8px rgba(124,58,237,.25)!important;
}
section[data-testid="stFileUploaderDropzone"]{
  border-radius:18px!important; border:2px dashed #a855f7!important;
  background: linear-gradient(135deg,#fdf4ff,#eef2ff)!important;
}
.stDownloadButton>button{
  background: linear-gradient(120deg,#f97316,#ec4899,#8b5cf6)!important;
  background-size:200% 100%!important;
  color:white!important; border-radius:14px!important;
  padding:13px 24px!important; font-size:1.05rem!important;
  font-weight:700!important; width:100%!important; border:none!important;
  box-shadow:0 10px 24px -8px rgba(236,72,153,.55)!important;
  transition:all .25s ease!important;
}
.stDownloadButton>button:hover{
  background-position:100% 0!important; transform:translateY(-2px)!important;
  box-shadow:0 14px 30px -8px rgba(236,72,153,.7)!important;
}
.footer-note{
  text-align:center; color:#9333ea; font-size:.85rem; font-weight:600;
  background:linear-gradient(90deg,#fdf4ff,#eef2ff); border-radius:12px;
  padding:12px; margin-top:8px;
}
</style>
""", unsafe_allow_html=True)
st.markdown("""
<div class="hero">
  <h1>📄✨ Statement → Excel</h1>
  <p>Upload a customer statement PDF and get a beautifully formatted Excel file in seconds.</p>
  <span class="badge">🟠 Savannah Cement</span>
  <span class="badge">🔵 National Cement</span>
  <span class="badge">🟢 Mombasa Cement</span>
  <span class="badge">🟣 Karsan Ramji</span>
  <span class="badge">🟦 Skylife</span>
  <span class="badge">🟪 Springtech</span>
  <span class="badge">🔴 Scania</span>
  <span class="badge">🟩 Trans-Solutions</span>
  <span class="badge">🟧 AutoXpress</span>
  <span class="badge">🟥 Gathani</span>
  <span class="badge">🩵 Anco Trading</span>
  <span class="badge">🟤 Doshi / Ultra (COA)</span>
  <span class="badge">🟢 Bamburi Cement</span>
  <span class="badge">🔴 Crown Paints</span>
  <span class="badge">🔷 Mapei East Africa</span>
  <span class="badge">🔵 Basco Paints</span>
  <span class="badge">🐬 Dolphin Ricambi</span>
  <span class="badge">🔒 Files never stored</span>
</div>
""", unsafe_allow_html=True)
with st.expander("🎨 Supported formats"):
    st.markdown("""
| Supplier | Key Columns |
|---|---|
| 🟠 **Savannah Cement** | Doc No · Ref/Vehicle · Date · Due Date · Details · Debit · Credit · Balance |
| 🔵 **National Cement** | Date · Description · Vehicle · LPO · Qty · Invoice · CU Invoice · Location · Debit · Credit · Balance + Cheques sheet |
| 🟢 **Mombasa Cement** | Doc Date · Reference · Cheque No · Description · Amount · Balance |
| 🟣 **Karsan Ramji (Ndovu)** | Date · Doc No · BP Ref · CU Invoice · Line Memo · Debit · Credit · Balance |
| 🟦 **Skylife Kenya** | Tally-style ledger: Date · Particulars · Vch No · Debit · Credit · Running Balance |
| 🟪 **Springtech (K) Ltd** | Date · Transaction · Amount · Balance + Ageing Summary sheet |
| 🔴 **Scania East Africa** | Date · Invoice · Description · Due Date · Debit · Credit · Balance |
| 🟩 **Trans-Solutions (E.A)** | Date · Transaction · Amount · Balance + Ageing Summary sheet |
| 🟧 **AutoXpress** | Date · Tax Invoice · GP Invoice · LPO · Opening Bal · Invoice · Payment · Cummulative + Ageing Summary sheet |
| 🟥 **Gathani Limited** | Date · Particulars · Doc Ref · Debit · Credit · Nett Amount + Ageing Summary sheet |
| 🩵 **Anco Trading Ltd** | Transaction · Doc No · Ref No · Doc Date · DR/CR · Debit · Credit · Balance + Ageing Analysis sheet |
| 🟤 **Doshi / Ultra (COA)** | Two-column layout: Debit (Date · Particulars · CU Invoice · Amount) + Credit (Date · Particulars · Amount) + Ageing Summary |
| 🟢 **Bamburi Cement PLC** | Doc No · Delivery Note · Posting Date · Doc Type · Item Text/Ref · Qty · Unit · Amount · Running Balance |
| 🔴 **Crown Paints Kenya PLC** | Date · Description (A/R Invoice/Credit Memo/Incoming Payment) · Doc No · Reference · Debit · Credit · Cumulative Bal. + Ageing Summary + Product Summary sheets |
| 🔷 **Mapei East Africa Ltd** | Date · Invoice · Description · Due Date · Debit · Credit · Balance (credits shown as negatives; wrapped descriptions handled) |
| 🔵 **Basco Products (K) Ltd** | Doc Date · Invoice No · Type (Invoice/Paymnt/Credit/Adjmnt) · Reference/Order No · Debit · Credit · Balance — multi-account merged statements + Ageing Summary sheet |
| 🐬 **Dolphin Ricambi Ltd** | Tally-style ledger: Date · Particulars · Vch Type · Vch No · Excise Inv No · Narration · Debit · Credit · Running Balance |
    """)
# ── Colours & helpers ──────────────────────────────────────────────────────────
C = dict(savannah_orange="E8540A", national_blue="1565C0", mombasa_green="1B5E20",
         karsan_purple="4A148C", header_dark="2C3E50", col_hdr_fg="FFFFFF",
         alt_row="F2F6FA", credit_row="E8F5E9", receipt_row="FFF3E0",
         info_bg="ECF0F1", debit_fg="1565C0", credit_fg="2E7D32",
         neg_bal="C62828", border="BDBDBD",
         skylife_teal="00838F", springtech_indigo="3F51B5", scania_red="C8102E",
         transsolutions_cyan="0097A7", autoxpress_amber="F57C00",
         gathani_maroon="7B1C1C", anco_teal="00695C", bamburi_green="00843D",
         crown_red="C8102E", mapei_blue="005CA9", basco_blue="0D47A1",
         dolphin_navy="0D3B66")
MFMT = '#,##0.00;(#,##0.00);"-"'
IFMT = '#,##0;(#,##0);"-"'
def pnf(s):
    """General money parser: handles leading/trailing '-', commas, currency prefixes."""
    if s is None: return None
    s = str(s).strip()
    if not s: return None
    neg = False
    if s.startswith("-"): neg = True; s = s[1:]
    if s.endswith("-"): neg = True; s = s[:-1]
    s = s.replace(",", "").replace("KES", "").replace("KShs", "").strip()
    try: v = float(s)
    except: return None
    return -v if neg else v
def thin(sides="all"):
    s = Side(style="thin", color=C["border"])
    kw = dict(left=s,right=s,top=s,bottom=s) if sides=="all" else dict(bottom=s)
    return Border(**kw)
def pn(s):
    if s is None: return None
    try: return float(str(s).strip().replace(",","").replace("-","").strip()) * (-1 if str(s).strip().endswith("-") else 1)
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
def banner(ws, row, cols, text, bg):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=cols)
    wc(ws, row, 1, text, bold=True, fg=C["col_hdr_fg"], bg=bg, align="center", sz=14)
    ws.row_dimensions[row].height = 30
def col_headers(ws, row, headers, bg):
    for ci, h in enumerate(headers, 1):
        wc(ws, row, ci, h, bold=True, fg=C["col_hdr_fg"], bg=bg, align="center", bdr=thin())
    ws.row_dimensions[row].height = 22
# ── FORMAT DETECTION ───────────────────────────────────────────────────────────
def detect(lines):
    t = "\n".join(lines[:50])
    tfull = "\n".join(lines[:150])
    if "DOLPHIN RICAMBI" in tfull.upper() or "DOLPHIN" in t.upper():     return "dolphin"
    if "MAPEI" in tfull.upper():                                       return "mapei"
    if "BASCO" in tfull.upper():                                       return "basco"
    if "CROWN PAINTS" in tfull.upper():                                return "crown"
    if "SKYLIFE" in t.upper():                                        return "skylife"
    if "SPRINGTECH" in t.upper():                                      return "springtech"
    if "TRANS-SOLUTIONS" in t.upper() or "TRANSSOLUTIONS" in t.upper(): return "transsolutions"
    if "AUTOXPRESS" in t.upper() or "AUTO XPRESS" in t.upper():        return "autoxpress"
    if "SCANIA" in tfull.upper():                                      return "scania"
    if "MOMBASA CEMENT" in t.upper():                                  return "mombasa"
    if "Karsan Ramji" in t or "NDOVU" in t.upper() or "BP REF" in t:  return "karsan"
    if "NATIONAL CEMENT" in t.upper() or "CU Invoice Number" in t:    return "national"
    if "Savannah Cement" in t or "A/R Invoices" in t:                 return "savannah"
    if "LPO No" in t or "Cheque no" in t:                             return "national"
    if "Due Date" in t and "Document" in t:                           return "savannah"
    if "AR Invoice" in t and "BP REF" in t:                           return "karsan"
    if "BAMBURI CEMENT" in t.upper():                                  return "bamburi"
    if "GATHANI" in t.upper():                                         return "gathani"
    if "ANCO TRADING" in t.upper():                                    return "anco"
    if "Confirmation of Accounts" in t or "Doshi" in t or "Ultra Multiware" in t: return "coa"
    if re.search(r"Date\s+Particulars\s+Debit Amount", t):                 return "coa"
    return "unknown"
# ══════════════════════════════════════════════════════════════════════════════
# MOMBASA CEMENT PARSER
# ══════════════════════════════════════════════════════════════════════════════
DATE_MCL = re.compile(r"^\d{2}\.\d{2}\.\d{4}$")
MONEY_RE  = re.compile(r"^-?[\d,]+\.\d{2}-?$")
def parse_mombasa(lines):
    hdr = {"supplier": "Mombasa Cement Limited", "account": "", "account_no": "",
           "currency": "KES", "period": ""}
    for ln in lines[:15]:
        m = re.search(r"ACCOUNT:\s*(.+?)\s{2,}", ln)
        if m: hdr["account"] = m.group(1).strip()
        m = re.search(r"ACCOUNT NO\s+(\S+)", ln)
        if m: hdr["account_no"] = m.group(1)
        m = re.search(r"CURRENCY\s+(\S+)", ln)
        if m: hdr["currency"] = m.group(1)
    for ln in lines:
        m = re.search(r"PERIOD\s+([\d.]+)\s+TO\s+([\d.]+)", ln)
        if m: hdr["period"] = f"{m.group(1)} to {m.group(2)}"; break
    txs = []
    skip = {"DOC. DATE","OUR","REFERENCE","CHEQUE NO","DESCRIPTION","AMOUNT",
            "BALANCE","MOMBASA CEMENT","P. O. BOX","MOMBASA,KENYA","Tel:",
            "Mobile:","Email:","STATEMENT","ACCOUNT:","NAIROBI","KENYA","PAGE",
            "NET TOTAL","CLOSING BALANCE","AGEING","UNALLOCATED","AMOUNT DUE",
            "1 TO","31 TO","61 TO","91 TO","OVER","We consider","writing within",
            "Please note","TRANSACTION DURING","OPENING BALANCE"}
    section = "opening"
    for ln in lines:
        ln = ln.strip()
        if "TRANSACTION DURING PERIOD" in ln: section = "current"; continue
        if "CLOSING BALANCE" in ln:
            p = ln.split()
            amts = [x for x in p if MONEY_RE.match(x)]
            if amts: hdr["closing_balance"] = pn(amts[-1])
            continue
        if any(ln.startswith(s) or s in ln for s in skip if len(s) > 3): continue
        parts = ln.split()
        if not parts or not DATE_MCL.match(parts[0]): continue
        date = parts[0]; rest = parts[1:]
        amts_idx = []
        for j in range(len(rest)-1, -1, -1):
            if MONEY_RE.match(rest[j]): amts_idx.insert(0, j)
            elif amts_idx: break
        if not amts_idx: continue
        first_amt = amts_idx[0]
        desc_parts = rest[:first_amt]
        amt_parts  = rest[first_amt:]
        ref_no = desc_parts[0] if desc_parts else ""
        cheque = ""
        desc   = ""
        if len(desc_parts) >= 3:
            if re.match(r'^[A-Z0-9]+$', desc_parts[1]) and not MONEY_RE.match(desc_parts[1]):
                cheque = desc_parts[1]; desc = " ".join(desc_parts[2:])
            else:
                desc = " ".join(desc_parts[1:])
        elif len(desc_parts) == 2:
            desc = desc_parts[1]
        amount = pn(amt_parts[0]) if amt_parts else None
        balance= pn(amt_parts[1]) if len(amt_parts) > 1 else None
        is_receipt = "Receipt" in desc or "receipt" in desc
        txs.append({"date": date, "ref_no": ref_no, "cheque_no": cheque,
                    "description": desc, "amount": amount, "balance": balance,
                    "section": section, "is_receipt": is_receipt})
    return hdr, txs
def excel_mombasa(hdr, txs):
    wb = openpyxl.Workbook(); ws = wb.active
    ws.title = "Statement"; ws.sheet_properties.tabColor = C["mombasa_green"]
    for i, w in enumerate([14, 14, 16, 14, 28, 18, 18], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    banner(ws, 1, 7, "MOMBASA CEMENT LIMITED  |  CUSTOMER STATEMENT", C["mombasa_green"])
    lf = Font(name="Arial", bold=True, size=9)
    vf = Font(name="Arial", size=9)
    lb = PatternFill("solid", start_color=C["info_bg"])
    al = Alignment(horizontal="left", vertical="center", indent=1)
    info = [("Account", hdr.get("account",""), "Account No", hdr.get("account_no","")),
            ("Currency", hdr.get("currency","KES"), "Period", hdr.get("period",""))]
    for r, (l1,v1,l2,v2) in enumerate(info, start=2):
        for ci,lbl,val in [(1,l1,v1),(5,l2,v2)]:
            ws.merge_cells(start_row=r,start_column=ci,end_row=r,end_column=ci+1)
            ws.merge_cells(start_row=r,start_column=ci+2,end_row=r,end_column=ci+2)
            lc=ws.cell(row=r,column=ci,value=lbl); lc.font=lf; lc.fill=lb; lc.alignment=al
            vc=ws.cell(row=r,column=ci+2,value=val); vc.font=vf; vc.alignment=al
        ws.row_dimensions[r].height=16
    H = 4
    col_headers(ws, H, ["Doc Date","Our Reference","Cheque No","Section",
                         "Description","Amount (KES)","Balance (KES)"], C["mombasa_green"])
    D = H + 1
    for ri, tx in enumerate(txs, start=D):
        is_rec = tx["is_receipt"]
        is_open= tx["section"] == "opening"
        bg = C["receipt_row"] if is_rec else (C["alt_row"] if (ri-D)%2==1 else None)
        sec_label = "Opening Bal." if is_open else "Current Period"
        vals   = [tx["date"], tx["ref_no"], tx["cheque_no"], sec_label,
                  tx["description"], tx["amount"], tx["balance"]]
        aligns = ["center","center","center","center","left","right","right"]
        fmts   = [None,None,None,None,None,MFMT,MFMT]
        for ci,(v,a,f) in enumerate(zip(vals,aligns,fmts),1):
            fg="000000"
            if ci==6 and v is not None: fg=C["credit_fg"] if is_rec else C["debit_fg"]
            if ci==7 and v is not None: fg=C["neg_bal"] if v<0 else C["credit_fg"]
            wc(ws,ri,ci,v,align=a,fmt=f,fg=fg,bg=bg,bdr=thin("bottom"))
        ws.row_dimensions[ri].height=15
    T = D + len(txs)
    ws.merge_cells(f"A{T}:E{T}")
    wc(ws,T,1,"TOTAL / CLOSING BALANCE",bold=True,fg=C["col_hdr_fg"],
       bg=C["mombasa_green"],align="right",bdr=thin(),ind=1)
    for ci,col in [(6,"F"),(7,"G")]:
        c=ws.cell(row=T,column=ci,value=f"=SUM({col}{D}:{col}{T-1})")
        c.font=Font(name="Arial",bold=True,size=9,color=C["col_hdr_fg"])
        c.fill=PatternFill("solid",start_color=C["mombasa_green"])
        c.number_format=MFMT; c.alignment=Alignment(horizontal="right",vertical="center")
        c.border=thin()
    ws.row_dimensions[T].height=18
    ws.freeze_panes=f"A{D}"; ws.auto_filter.ref=f"A{H}:G{T-1}"
    ws.page_setup.orientation="landscape"; ws.page_setup.fitToPage=True; ws.page_setup.fitToWidth=1
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf
# ══════════════════════════════════════════════════════════════════════════════
# DOLPHIN RICAMBI LIMITED PARSER  (Tally-style ledger: To/By debit/credit,
# multi-period opening & closing balances, invoice detail wrapped on next line)
# ══════════════════════════════════════════════════════════════════════════════
DATE_DOL  = re.compile(r"^\d{1,2}-\d{1,2}-\d{4}$")      # 1-5-2026 / 15-5-2026
MONEY_DOL = re.compile(r"^[\d,]+\.\d{2}$")
VTYPE_SINGLE = {"Sales", "Receipt", "Journal", "Payment", "Contra", "Purchase"}
def parse_dolphin(lines):
    hdr = {"supplier": "Dolphin Ricambi Limited", "customer": "", "period": "",
           "po_box": "", "tel": "", "opening_balance": None, "closing_balance": None}
    for ln in lines[:12]:
        s = ln.strip()
        m = re.search(r"POST BOX NO\s+(.+)", s, re.I)
        if m: hdr["po_box"] = m.group(1).strip()
        m = re.search(r"TEL NO\s+(.+)", s, re.I)
        if m: hdr["tel"] = m.group(1).strip()
        if "Muthokinju" in s and "Ledger" not in s: hdr["customer"] = s
        m = re.match(r"^(\d{1,2}-[A-Za-z]{3}-\d{4})\s+to\s+(\d{1,2}-[A-Za-z]{3}-\d{4})$", s)
        if m: hdr["period"] = f"{m.group(1)} to {m.group(2)}"
    SKIP = ("DOLPHIN RICAMBI", "POST BOX", "NAIROBI", "TEL NO", "Ledger Account",
            "Post Box No", "Page", "Date Particulars", "continued",
            "Carried Over", "Brought Forward")
    txs = []; cur_date = ""; opening_seen = False
    for ln in lines:
        s = ln.strip()
        if not s: continue
        if any(k in s for k in SKIP): continue
        parts = s.split()
        if all(MONEY_DOL.match(p) for p in parts): continue
        if "Closing Balance" in s:
            m = re.search(r"([\d,]+\.\d{2})", s)
            if m: hdr["closing_balance"] = pnf(m.group(1))
            continue
        if DATE_DOL.match(parts[0]):
            cur_date = parts[0]; rest = parts[1:]
        else:
            rest = parts
        if not rest: continue
        if rest[0] not in ("To", "By"):
            if txs:
                m = re.search(r"inv no\s+(\S+)", s, re.I)
                if m and not txs[-1]["inv_no"]:
                    txs[-1]["inv_no"] = m.group(1)
                    extra = re.sub(r"inv no\s+\S+", "", s, flags=re.I).strip()
                else:
                    extra = s
                if extra:
                    txs[-1]["narration"] = (txs[-1]["narration"] + " " + extra).strip()
            continue
        side = rest[0]; body = rest[1:]
        if not body or not MONEY_DOL.match(body[-1]): continue
        amount = pnf(body[-1]); mid = body[:-1]
        if "Opening Balance" in " ".join(mid):
            if opening_seen: continue
            opening_seen = True
            if hdr["opening_balance"] is None: hdr["opening_balance"] = amount
            txs.append({"date": cur_date, "particulars": "Opening Balance", "vch_type": "",
                        "vch_no": "", "inv_no": "", "narration": "",
                        "debit": amount if side == "To" else None,
                        "credit": amount if side == "By" else None})
            continue
        vch_no = ""
        if mid and re.match(r"^\d+$", mid[-1]):
            vch_no = mid.pop()
        vch_type = ""
        if len(mid) >= 2 and mid[-2] in ("Credit", "Debit") and mid[-1] == "Note":
            vch_type = mid[-2] + " " + mid[-1]; mid = mid[:-2]
        elif mid and mid[-1] in VTYPE_SINGLE:
            vch_type = mid.pop()
        particulars = " ".join(mid).strip() or "—"
        txs.append({"date": cur_date, "particulars": particulars, "vch_type": vch_type,
                    "vch_no": vch_no, "inv_no": "", "narration": "",
                    "debit": amount if side == "To" else None,
                    "credit": amount if side == "By" else None})
    return hdr, txs
def excel_dolphin(hdr, txs):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Statement"
    ws.sheet_properties.tabColor = C["dolphin_navy"]
    for i, w in enumerate([12, 30, 13, 12, 14, 26, 15, 15, 18], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    banner(ws, 1, 9, "DOLPHIN RICAMBI LIMITED  |  LEDGER ACCOUNT STATEMENT", C["dolphin_navy"])
    lf = Font(name="Arial", bold=True, size=9); vf = Font(name="Arial", size=9)
    lb = PatternFill("solid", start_color=C["info_bg"])
    al = Alignment(horizontal="left", vertical="center", indent=1)
    info = [("Customer", hdr.get("customer", ""), "Period", hdr.get("period", "")),
            ("PO Box", hdr.get("po_box", ""), "Tel", hdr.get("tel", ""))]
    for r, (l1, v1, l2, v2) in enumerate(info, start=2):
        for ci, lbl, val in [(1, l1, v1), (6, l2, v2)]:
            ws.merge_cells(start_row=r, start_column=ci, end_row=r, end_column=ci + 1)
            ws.merge_cells(start_row=r, start_column=ci + 2, end_row=r, end_column=ci + 2 if ci == 6 else ci + 3)
            lc = ws.cell(row=r, column=ci, value=lbl); lc.font = lf; lc.fill = lb; lc.alignment = al
            vc = ws.cell(row=r, column=ci + 2, value=val); vc.font = vf; vc.alignment = al
        ws.row_dimensions[r].height = 16
    H = 4
    col_headers(ws, H, ["Date", "Particulars", "Vch Type", "Vch No.", "Excise Inv. No.",
                        "Narration", "Debit (KES)", "Credit (KES)", "Running Balance (KES)"],
                C["dolphin_navy"])
    D = H + 1
    for ri, tx in enumerate(txs, start=D):
        is_open = "Opening Balance" in tx["particulars"]
        is_cr = tx["credit"] is not None
        bg = C["info_bg"] if is_open else (C["credit_row"] if is_cr else (C["alt_row"] if (ri - D) % 2 == 1 else None))
        if ri == D:
            c = ws.cell(row=ri, column=9, value=(tx["debit"] or 0) - (tx["credit"] or 0))
        else:
            c = ws.cell(row=ri, column=9, value=f"=I{ri-1}+G{ri}-H{ri}")
        c.number_format = MFMT; c.font = Font(name="Arial", size=9, bold=is_open)
        c.alignment = Alignment(horizontal="right", vertical="center"); c.border = thin("bottom")
        vals = [tx["date"], tx["particulars"], tx["vch_type"], tx["vch_no"], tx["inv_no"],
                tx["narration"], tx["debit"], tx["credit"]]
        aligns = ["center", "left", "center", "center", "center", "left", "right", "right"]
        fmts = [None, None, None, None, None, None, MFMT, MFMT]
        fgs = ["000000", "000000", "000000", "000000", "555555", "777777", C["debit_fg"], C["credit_fg"]]
        for ci, (v, a, f, fg) in enumerate(zip(vals, aligns, fmts, fgs), 1):
            wc(ws, ri, ci, v, align=a, fmt=f, fg=fg, bg=bg, bdr=thin("bottom"), sz=9 if ci != 5 else 8)
        ws.row_dimensions[ri].height = 15
    T = D + len(txs)
    ws.merge_cells(f"A{T}:F{T}")
    wc(ws, T, 1, "TOTAL", bold=True, fg=C["col_hdr_fg"], bg=C["dolphin_navy"], align="right", bdr=thin(), ind=1)
    for ci, col in [(7, "G"), (8, "H")]:
        c = ws.cell(row=T, column=ci, value=f"=SUM({col}{D}:{col}{T-1})")
        c.font = Font(name="Arial", bold=True, size=9, color=C["col_hdr_fg"])
        c.fill = PatternFill("solid", start_color=C["dolphin_navy"])
        c.number_format = MFMT; c.alignment = Alignment(horizontal="right", vertical="center"); c.border = thin()
    cb = hdr.get("closing_balance")
    c = ws.cell(row=T, column=9, value=cb)
    c.font = Font(name="Arial", bold=True, size=9, color=C["col_hdr_fg"])
    c.fill = PatternFill("solid", start_color=C["dolphin_navy"])
    c.number_format = MFMT; c.alignment = Alignment(horizontal="right", vertical="center"); c.border = thin()
    ws.row_dimensions[T].height = 18
    ws.freeze_panes = f"A{D}"; ws.auto_filter.ref = f"A{H}:I{T-1}"
    ws.page_setup.orientation = "landscape"; ws.page_setup.fitToPage = True; ws.page_setup.fitToWidth = 1
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf
# ══════════════════════════════════════════════════════════════════════════════
# KARSAN RAMJI PARSER
# ══════════════════════════════════════════════════════════════════════════════
DATE_KRS = re.compile(r"^\d{2}/\d{2}/\d{4}$")
MONEY_KRS= re.compile(r"^-?[\d,]+\.\d{2}$")
def parse_karsan(lines):
    hdr = {"supplier":"Karsan Ramji and Sons Ltd","customer":"","customer_code":"",
           "phone":"","email":"","period":""}
    for i, ln in enumerate(lines[:10]):
        m = re.search(r"(N\d+)\s+(.+)", ln)
        if m: hdr["customer_code"]=m.group(1); hdr["customer"]=m.group(2).strip()
        m = re.search(r"Phone#?:\s*(\S+)", ln)
        if m: hdr["phone"]=m.group(1)
        m = re.search(r"Email:\s*(\S+)", ln)
        if m: hdr["email"]=m.group(1)
        m = re.search(r"From\s+([\d/]+).+To\s+([\d/]+)", ln)
        if m: hdr["period"]=f"{m.group(1)} to {m.group(2)}"
    txs = []
    skip_kw = {"DATE","DOC. NO.","BP REF.","CU INVOICE","LINE MEMO","DEBIT","CREDIT",
               "BALANCE","Customer Statement","Document Date","Karsan","Phone","Address",
               "Email","Current Balance","Total","Aging","E & O","Page"}
    for ln in lines:
        ln = ln.strip()
        if any(k in ln for k in skip_kw): continue
        if "Balance B/F" in ln:
            p = ln.split()
            amts = [x for x in p if MONEY_KRS.match(x)]
            bal = pn(amts[-1]) if amts else None
            txs.append({"date":"--","doc_no":"--","bp_ref":"--","cu_invoice":"--",
                        "line_memo":"Balance B/F","debit":None,"credit":None,
                        "balance":bal,"is_bf":True})
            continue
        parts = ln.split()
        if not parts or not DATE_KRS.match(parts[0]): continue
        date=parts[0]; rest=parts[1:]
        amts=[]; nonamts=[]
        for x in reversed(rest):
            if MONEY_KRS.match(x) and len(amts)<3: amts.insert(0,x)
            else: nonamts.insert(0,x)
        dbt = pn(amts[0]) if len(amts)>=3 else None
        crd = pn(amts[1]) if len(amts)>=3 else None
        bal = pn(amts[2]) if len(amts)>=3 else (pn(amts[-1]) if amts else None)
        na=nonamts
        doc_no  = na[0] if len(na)>0 else ""
        bp_ref  = na[1] if len(na)>1 else ""
        cu_inv  = na[2] if len(na)>2 else ""
        memo    = " ".join(na[3:]) if len(na)>3 else ""
        txs.append({"date":date,"doc_no":doc_no,"bp_ref":bp_ref,"cu_invoice":cu_inv,
                    "line_memo":memo,"debit":dbt,"credit":crd,"balance":bal,"is_bf":False})
    return hdr, txs
def excel_karsan(hdr, txs):
    wb=openpyxl.Workbook(); ws=wb.active
    ws.title="Statement"; ws.sheet_properties.tabColor=C["karsan_purple"]
    for i,w in enumerate([12,14,12,22,14,16,16,18],1):
        ws.column_dimensions[get_column_letter(i)].width=w
    banner(ws,1,8,"KARSAN RAMJI AND SONS LTD (NDOVU CEMENT)  |  CUSTOMER STATEMENT",C["karsan_purple"])
    lf=Font(name="Arial",bold=True,size=9); vf=Font(name="Arial",size=9)
    lb=PatternFill("solid",start_color=C["info_bg"])
    al=Alignment(horizontal="left",vertical="center",indent=1)
    info=[("Customer",hdr.get("customer",""),"Code",hdr.get("customer_code","")),
          ("Period",hdr.get("period",""),"Phone",hdr.get("phone",""))]
    for r,(l1,v1,l2,v2) in enumerate(info,start=2):
        for ci,lbl,val in [(1,l1,v1),(5,l2,v2)]:
            ws.merge_cells(start_row=r,start_column=ci,end_row=r,end_column=ci+1)
            ws.merge_cells(start_row=r,start_column=ci+2,end_row=r,end_column=ci+2)
            lc=ws.cell(row=r,column=ci,value=lbl); lc.font=lf; lc.fill=lb; lc.alignment=al
            vc=ws.cell(row=r,column=ci+2,value=val); vc.font=vf; vc.alignment=al
        ws.row_dimensions[r].height=16
    H=4
    col_headers(ws,H,["Date","Doc No.","BP Ref No.","CU Invoice No.",
                       "Line Memo","Debit (KES)","Credit (KES)","Balance (KES)"],C["karsan_purple"])
    D=H+1
    for ri,tx in enumerate(txs,start=D):
        is_bf=tx.get("is_bf",False)
        bg=C["info_bg"] if is_bf else (C["alt_row"] if (ri-D)%2==1 else None)
        vals=[tx["date"],tx["doc_no"],tx["bp_ref"],tx["cu_invoice"],
              tx["line_memo"],tx["debit"],tx["credit"],tx["balance"]]
        aligns=["center","center","center","center","left","right","right","right"]
        fmts=[None,None,None,None,None,MFMT,MFMT,MFMT]
        for ci,(v,a,f) in enumerate(zip(vals,aligns,fmts),1):
            fg="000000"
            if ci==6 and v: fg=C["debit_fg"]
            if ci==7 and v: fg=C["credit_fg"]
            if ci==8 and v is not None: fg=C["neg_bal"] if v<0 else C["credit_fg"]
            wc(ws,ri,ci,v,align=a,fmt=f,fg=fg,bg=bg,bdr=thin("bottom"))
        ws.row_dimensions[ri].height=15
    T=D+len(txs)
    ws.merge_cells(f"A{T}:E{T}")
    wc(ws,T,1,"CURRENT BALANCE",bold=True,fg=C["col_hdr_fg"],bg=C["karsan_purple"],
       align="right",bdr=thin(),ind=1)
    for ci,col in [(6,"F"),(7,"G"),(8,"H")]:
        c=ws.cell(row=T,column=ci,value=f"=SUM({col}{D}:{col}{T-1})")
        c.font=Font(name="Arial",bold=True,size=9,color=C["col_hdr_fg"])
        c.fill=PatternFill("solid",start_color=C["karsan_purple"])
        c.number_format=MFMT; c.alignment=Alignment(horizontal="right",vertical="center")
        c.border=thin()
    ws.row_dimensions[T].height=18
    ws.freeze_panes=f"A{D}"; ws.auto_filter.ref=f"A{H}:H{T-1}"
    ws.page_setup.orientation="landscape"; ws.page_setup.fitToPage=True; ws.page_setup.fitToWidth=1
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf
# ══════════════════════════════════════════════════════════════════════════════
# NATIONAL CEMENT PARSER  (precise column-aware parser)
# ══════════════════════════════════════════════════════════════════════════════
DR4     = re.compile(r"^\d{2}\.\d{2}\.\d{4}$")
MONEY_N = re.compile(r"^[\d,]+\.\d{2}$")
QTY_N   = re.compile(r"^\d+\.\d{3}$")
LPOSUF  = re.compile(r"^\d{7,8}$")
def parse_national(lines):
    hdr = {}
    for ln in lines[:20]:
        m = re.search(r"Statement As At:\s*([\d.]+)", ln)
        if m: hdr["statement_date"] = m.group(1)
        m = re.search(r"PIN No:\s*(\S+)", ln)
        if m: hdr["pin"] = m.group(1)
        if "NATIONAL CEMENT" in ln.upper(): hdr["supplier"] = "National Cement Company Ltd"
        if "New Muthokinju" in ln: hdr["customer"] = "New Muthokinju Hardware Ltd"
    hdr.setdefault("supplier", "National Cement Company Ltd")
    hdr.setdefault("customer", "New Muthokinju Hardware Ltd")
    txs = []; chqs = []; in_chq = False
    i = 0
    while i < len(lines):
        ln = lines[i].strip()
        if "Cheques on Hand" in ln: in_chq = True; i+=1; continue
        if "Cheque No" in ln and "Due On" in ln: i+=1; continue
        if "Total/Closing" in ln or "Ageing" in ln or "Total" == ln.split()[0] if ln.split() else False:
            i+=1; continue
        if in_chq:
            p = ln.split()
            if len(p) == 3 and DR4.match(p[1]):
                try: chqs.append({"cheque_no":p[0],"due_on":p[1],"amount":float(p[2].replace(",",""))})
                except: pass
            i+=1; continue
        parts = ln.split()
        if not parts or not DR4.match(parts[0]): i+=1; continue
        date = parts[0]
        rest = list(parts[1:])
        lpo_suffix = ""
        if i+1 < len(lines):
            nxt = lines[i+1].strip()
            if LPOSUF.match(nxt):
                lpo_suffix = nxt
                i += 1
        if rest and rest[0] == "Balance":
            dr_cr = rest[-1] if rest[-1] in ("DR","CR") else ""
            nums = [x for x in rest if MONEY_N.match(x)]
            bal = float(nums[-1].replace(",","")) if nums else None
            bal_disp = f"{bal:,.2f} {dr_cr}".strip() if bal is not None else None
            txs.append({"date":date,"description":"Balance B/F","cheque_no":"",
                        "vehicle":"","lpo_no":"","quantity":None,
                        "invoice_no":"","cu_invoice":"","location":"",
                        "currency":"KES","debit":0.0,"credit":0.0,
                        "balance":bal,"balance_disp":bal_disp,"is_bf":True})
            i+=1; continue
        desc = rest[0] if rest else "Invoice"
        rest = rest[1:]
        qty_idx = next((j for j,x in enumerate(rest) if QTY_N.match(x)), None)
        if qty_idx is None: i+=1; continue
        before_qty = rest[:qty_idx]
        after_qty  = rest[qty_idx:]
        qty = float(after_qty[0].replace(",","")) if after_qty else None
        tail = list(after_qty[1:])
        dr_cr = ""
        if tail and tail[-1] in ("DR","CR"): dr_cr = tail.pop()
        balance = float(tail.pop().replace(",","")) if tail and MONEY_N.match(tail[-1]) else None
        balance_disp = f"{balance:,.2f} {dr_cr}".strip() if balance is not None else None
        credit  = float(tail.pop().replace(",","")) if tail and MONEY_N.match(tail[-1]) else None
        debit   = float(tail.pop().replace(",","")) if tail and MONEY_N.match(tail[-1]) else None
        curr = tail.pop() if tail else "KES"
        inv_no   = tail[0] if len(tail) >= 1 else ""
        cu_inv   = tail[1] if len(tail) >= 2 else ""
        location = " ".join(tail[2:]) if len(tail) >= 3 else ""
        vehicle = ""; lpo_prefix = ""
        if before_qty:
            last = before_qty[-1]
            if re.match(r"^POHH\d+$", last) or re.match(r"^\d{4,5}$", last):
                lpo_prefix = last
                vehicle = " ".join(before_qty[:-1])
            else:
                vehicle = " ".join(before_qty)
        if lpo_prefix and lpo_suffix:
            lpo_no = lpo_prefix + lpo_suffix
        elif lpo_prefix:
            lpo_no = lpo_prefix
        else:
            lpo_no = lpo_suffix
        txs.append({"date":date,"description":desc,"cheque_no":"",
                    "vehicle":vehicle,"lpo_no":lpo_no,"quantity":qty,
                    "invoice_no":inv_no,"cu_invoice":cu_inv,"location":location,
                    "currency":curr,"debit":debit,"credit":credit,"balance":balance,
                    "balance_disp":balance_disp,"is_bf":False})
        i += 1
    return hdr, txs, chqs
def excel_national(hdr,txs,chqs):
    wb=openpyxl.Workbook(); ws=wb.active; ws.title="Statement"
    ws.sheet_properties.tabColor=C["national_blue"]
    for i,w in enumerate([12,12,14,20,18,10,14,26,12,6,16,16,18],1):
        ws.column_dimensions[get_column_letter(i)].width=w
    banner(ws,1,13,"NATIONAL CEMENT COMPANY LTD  |  CUSTOMER STATEMENT",C["national_blue"])
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
    col_headers(ws,H,["Date","Description","Cheque No","Vehicle No.","LPO No","Qty",
                       "Invoice No.","CU Invoice No.","Location","Curr.",
                       "Debit (KES)","Credit (KES)","Balance (KES)"],C["national_blue"])
    D=H+1
    for ri,tx in enumerate(txs,start=D):
        bf="B/F" in str(tx.get("description",""))
        bg=C["info_bg"] if bf else (C["alt_row"] if (ri-D)%2==1 else None)
        vals=[tx["date"],tx["description"],tx["cheque_no"],tx["vehicle"],tx["lpo_no"],
              tx["quantity"],tx["invoice_no"],tx["cu_invoice"],tx["location"],tx["currency"],
              tx["debit"],tx["credit"],tx.get("balance_disp", tx["balance"])]
        aligns=["center","left","center","center","center","right","center","center",
                "center","center","right","right","right"]
        fmts=[None,None,None,None,None,"#,##0.000",None,None,None,None,MFMT,MFMT,None]
        fgs=["000000"]*10+[C["debit_fg"],C["credit_fg"],"000000"]
        for ci,(v,a,f,fg) in enumerate(zip(vals,aligns,fmts,fgs),1):
            wc(ws,ri,ci,v,align=a,fmt=f,fg=fg,bg=bg,bdr=thin("bottom"))
        ws.row_dimensions[ri].height=15
    T=D+len(txs)
    ws.merge_cells(f"A{T}:J{T}")
    wc(ws,T,1,"TOTAL",bold=True,fg=C["col_hdr_fg"],bg=C["national_blue"],align="right",bdr=thin(),ind=1)
    for ci,col in [(11,"K"),(12,"L")]:
        c=ws.cell(row=T,column=ci,value=f"=SUM({col}{D}:{col}{T-1})")
        c.font=Font(name="Arial",bold=True,size=9,color=C["col_hdr_fg"])
        c.fill=PatternFill("solid",start_color=C["national_blue"])
        c.number_format=MFMT; c.alignment=Alignment(horizontal="right",vertical="center"); c.border=thin()
    last_bal = txs[-1].get("balance_disp","") if txs else ""
    c=ws.cell(row=T,column=13,value=last_bal)
    c.font=Font(name="Arial",bold=True,size=9,color=C["col_hdr_fg"])
    c.fill=PatternFill("solid",start_color=C["national_blue"])
    c.alignment=Alignment(horizontal="right",vertical="center"); c.border=thin()
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
            wc(ws2,ri,2,chq["due_on"],align="center",bg=bg,bdr=thin("bottom"))
            wc(ws2,ri,3,chq["amount"],align="right",bg=bg,bdr=thin("bottom"),fmt=MFMT,fg=C["national_blue"])
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
# ══════════════════════════════════════════════════════════════════════════════
# SAVANNAH CEMENT PARSER
# ══════════════════════════════════════════════════════════════════════════════
DR2=re.compile(r"^\d{2}\.\d{2}\.\d{2}$"); MR2=re.compile(r"^-?[\d,]+\.\d{2}$")
PFX=("IN ","CN ","RC ","DN ")
def parse_savannah(lines):
    hdr={"customer_name":"New Muthokinju Hardware Ltd","customer_code":"LC000072",
         "credit_limit":2_000_000.00,"sales_employee":"Fred Oguttu",
         "contact_person":"David","currency":"KES","printed_on":""}
    for ln in lines:
        m=re.search(r"Printed On:\s*\S+\s+(\S+)",ln)
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
    return hdr,txs
def excel_savannah(hdr,txs):
    wb=openpyxl.Workbook(); ws=wb.active; ws.title="Statement"
    ws.sheet_properties.tabColor=C["savannah_orange"]
    for i,w in enumerate([6,8,30,11,11,32,16,16,18],1):
        ws.column_dimensions[get_column_letter(i)].width=w
    banner(ws,1,9,"SAVANNAH CEMENT 2025 LIMITED  |  CUSTOMER STATEMENT",C["savannah_orange"])
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
    col_headers(ws,H,["Type","Doc No.","Ref/Vehicle","Date","Due Date",
                       "Details","Debit (KES)","Credit (KES)","Balance (KES)"],C["savannah_orange"])
    D=H+1
    for ri,tx in enumerate(txs,start=D):
        cr=tx["doc_type"] in ("CN","RC")
        bg=C["credit_row"] if cr else (C["alt_row"] if (ri-D)%2==1 else None)
        vals=[tx["doc_type"],tx["doc_no"],tx["ref_no"],tx["date"],tx["due_date"],
              tx["details"],tx["debit"],tx["credit"],tx["balance"]]
        aligns=["center","center","left","center","center","left","right","right","right"]
        fmts=[None,None,None,None,None,None,MFMT,MFMT,MFMT]
        fgs=["000000","000000","000000","000000","000000","000000",
             C["debit_fg"],C["credit_fg"],"000000"]
        for ci,(v,a,f,fg) in enumerate(zip(vals,aligns,fmts,fgs),1):
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
        c.number_format=MFMT; c.alignment=Alignment(horizontal="right",vertical="center"); c.border=thin()
    ws.row_dimensions[T].height=18
    ws.freeze_panes=f"A{D}"; ws.auto_filter.ref=f"A{H}:I{T-1}"
    ws.page_setup.orientation="landscape"; ws.page_setup.fitToPage=True; ws.page_setup.fitToWidth=1
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf
# ══════════════════════════════════════════════════════════════════════════════
# SKYLIFE KENYA LIMITED PARSER  (Tally-style ledger: To/By, Vch Type, Vch No.)
# ══════════════════════════════════════════════════════════════════════════════
DATE_SKY = re.compile(r"^\d{1,2}-[A-Za-z]{3}-\d{2}$")
MONEY_SKY = re.compile(r"^[\d,]+\.\d{2}$")
VCHNO_SKY = re.compile(r"^[A-Z0-9]{6,}$")
SKIP_SKY = ("Skylife Kenya Limited", "P.O.Box No.", "Nairobi, Kenya",
            "NEW MUTHOKINJU HARDWARE LTD", "Ledger Account", "Page", "continued",
            "Date Particulars")
def parse_skylife(lines):
    hdr = {"supplier": "Skylife Kenya Limited", "customer": "New Muthokinju Hardware Ltd",
           "period": "", "closing_balance": None}
    for ln in lines[:10]:
        m = re.search(r"^(\d{1,2}-[A-Za-z]{3}-\d{2})\s+to\s+(\d{1,2}-[A-Za-z]{3}-\d{2})$", ln.strip())
        if m: hdr["period"] = f"{m.group(1)} to {m.group(2)}"
    txs = []; current_date = ""
    for ln in lines:
        ln = ln.strip()
        if not ln: continue
        if any(s in ln for s in SKIP_SKY): continue
        if "Carried Over" in ln or "Brought Forward" in ln: continue
        if re.match(r"^[\d,]+\.\d{2}\s+[\d,]+\.\d{2}$", ln): continue
        parts = ln.split()
        if not parts: continue
        if "Closing Balance" in ln:
            m = re.search(r"([\d,]+\.\d{2})$", ln)
            if m: hdr["closing_balance"] = pnf(m.group(1))
            continue
        if DATE_SKY.match(parts[0]):
            current_date = parts[0]; rest = parts[1:]
        else:
            rest = parts
        if not rest or rest[0] not in ("To", "By"): continue
        side = "Debit" if rest[0] == "To" else "Credit"
        body = rest[1:]
        if not body or not MONEY_SKY.match(body[-1]): continue
        amount = pnf(body[-1])
        remainder = body[:-1]
        vch_no = ""
        if remainder and VCHNO_SKY.match(remainder[-1]) and any(ch.isdigit() for ch in remainder[-1]):
            vch_no = remainder.pop()
        particulars = " ".join(remainder).strip() or "—"
        txs.append({"date": current_date, "particulars": particulars, "vch_no": vch_no,
                    "debit": amount if side == "Debit" else None,
                    "credit": amount if side == "Credit" else None})
    return hdr, txs
def excel_skylife(hdr, txs):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Statement"
    ws.sheet_properties.tabColor = C["skylife_teal"]
    for i, w in enumerate([12, 40, 22, 16, 16, 18], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    banner(ws, 1, 6, "SKYLIFE KENYA LIMITED  |  LEDGER ACCOUNT STATEMENT", C["skylife_teal"])
    lf = Font(name="Arial", bold=True, size=9); vf = Font(name="Arial", size=9)
    lb = PatternFill("solid", start_color=C["info_bg"])
    al = Alignment(horizontal="left", vertical="center", indent=1)
    info = [("Customer", hdr.get("customer", ""), "Supplier", hdr.get("supplier", "")),
            ("Period", hdr.get("period", ""), "Currency", "KES")]
    for r, (l1, v1, l2, v2) in enumerate(info, start=2):
        for ci, lbl, val in [(1, l1, v1), (4, l2, v2)]:
            ws.merge_cells(start_row=r, start_column=ci, end_row=r, end_column=ci)
            ws.merge_cells(start_row=r, start_column=ci+1, end_row=r, end_column=ci+2 if ci==4 else ci+2)
            lc = ws.cell(row=r, column=ci, value=lbl); lc.font=lf; lc.fill=lb; lc.alignment=al
            vc = ws.cell(row=r, column=ci+1, value=val); vc.font=vf; vc.alignment=al
        ws.row_dimensions[r].height = 16
    H = 4
    col_headers(ws, H, ["Date", "Particulars", "Vch No.", "Debit (KES)", "Credit (KES)", "Running Balance (KES)"],
                C["skylife_teal"])
    D = H + 1
    for ri, tx in enumerate(txs, start=D):
        is_open = "Opening Balance" in tx["particulars"]
        bg = C["info_bg"] if is_open else (C["credit_row"] if tx["credit"] else (C["alt_row"] if (ri-D) % 2 == 1 else None))
        if ri == D:
            bal_formula = tx["debit"] or 0
            c = ws.cell(row=ri, column=6, value=bal_formula)
        else:
            c = ws.cell(row=ri, column=6, value=f"=F{ri-1}+D{ri}-E{ri}")
        c.number_format = MFMT; c.font = Font(name="Arial", size=9); c.alignment = Alignment(horizontal="right", vertical="center")
        c.border = thin("bottom")
        vals = [tx["date"], tx["particulars"], tx["vch_no"], tx["debit"], tx["credit"]]
        aligns = ["center", "left", "center", "right", "right"]
        fmts = [None, None, None, MFMT, MFMT]
        fgs = ["000000", "000000", "000000", C["debit_fg"], C["credit_fg"]]
        for ci, (v, a, f, fg) in enumerate(zip(vals, aligns, fmts, fgs), 1):
            wc(ws, ri, ci, v, align=a, fmt=f, fg=fg, bg=bg, bdr=thin("bottom"))
        ws.row_dimensions[ri].height = 15
    T = D + len(txs)
    ws.merge_cells(f"A{T}:C{T}")
    wc(ws, T, 1, "TOTAL", bold=True, fg=C["col_hdr_fg"], bg=C["header_dark"], align="right", bdr=thin(), ind=1)
    for ci, col in [(4, "D"), (5, "E")]:
        c = ws.cell(row=T, column=ci, value=f"=SUM({col}{D}:{col}{T-1})")
        c.font = Font(name="Arial", bold=True, size=9, color=C["col_hdr_fg"])
        c.fill = PatternFill("solid", start_color=C["header_dark"])
        c.number_format = MFMT; c.alignment = Alignment(horizontal="right", vertical="center"); c.border = thin()
    cb = hdr.get("closing_balance")
    c = ws.cell(row=T, column=6, value=cb if cb is not None else "")
    c.font = Font(name="Arial", bold=True, size=9, color=C["col_hdr_fg"])
    c.fill = PatternFill("solid", start_color=C["header_dark"])
    c.number_format = MFMT; c.alignment = Alignment(horizontal="right", vertical="center"); c.border = thin()
    ws.row_dimensions[T].height = 18
    ws.freeze_panes = f"A{D}"; ws.auto_filter.ref = f"A{H}:F{T-1}"
    ws.page_setup.orientation = "landscape"; ws.page_setup.fitToPage = True; ws.page_setup.fitToWidth = 1
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf
# ══════════════════════════════════════════════════════════════════════════════
# GENERIC "AGING STATEMENT" PARSER  (Springtech and Trans-Solutions)
# ══════════════════════════════════════════════════════════════════════════════
DATE_AG = re.compile(r"^\d{2}[-/]\d{2}[-/]\d{4}$")
MONEY_AG = re.compile(r"^-?[\d,]+\.\d{2}$")
def parse_aging_rows(lines):
    txs = []
    for ln in lines:
        ln = ln.strip()
        parts = ln.split()
        if not parts or not DATE_AG.match(parts[0]): continue
        rest = parts[1:]
        if not rest: continue
        money_idxs = [i for i, x in enumerate(rest) if MONEY_AG.match(x)]
        if len(money_idxs) >= 2:
            a_idx, b_idx = money_idxs[-2], money_idxs[-1]
            desc = " ".join(rest[:a_idx])
            amount = pnf(rest[a_idx]); balance = pnf(rest[b_idx])
        elif len(money_idxs) == 1:
            desc = " ".join(rest[:money_idxs[0]])
            amount = None; balance = pnf(rest[money_idxs[0]])
        else:
            continue
        txs.append({"date": parts[0], "description": desc.strip() or "—",
                    "amount": amount, "balance": balance})
    return txs
def parse_aging_summary(full_text):
    m = re.search(r"(-?[\d,]+\.\d{2})\s+(-?[\d,]+\.\d{2})\s+(-?[\d,]+\.\d{2})\s+"
                   r"(-?[\d,]+\.\d{2})\s+(-?[\d,]+\.\d{2})", full_text)
    buckets = {"current": None, "d1_30": None, "d31_60": None, "d61_90": None, "over90": None}
    if m:
        buckets = {"current": pnf(m.group(1)), "d1_30": pnf(m.group(2)), "d31_60": pnf(m.group(3)),
                   "d61_90": pnf(m.group(4)), "over90": pnf(m.group(5))}
    amt = None
    m2 = re.search(r"Amount Due[^\d\-]*(-?(?:KES|KShs)?\s*[\d,]+\.\d{2})", full_text)
    if m2: amt = pnf(m2.group(1))
    return buckets, amt
def parse_springtech(lines):
    hdr = {"supplier": "Springtech (K) Ltd", "address": "P.O. Box 94089 - 80107, Mombasa, Kenya",
           "customer": "New Muthokinju Hardware Ltd", "currency": "KES"}
    full = "\n".join(lines)
    hdr["buckets"], hdr["amount_due"] = parse_aging_summary(full)
    m = re.search(r"Date\s*\n?\s*(\d{2}-\d{2}-\d{4})", full)
    if m: hdr["statement_date"] = m.group(1)
    return hdr, parse_aging_rows(lines)
def parse_transsolutions(lines):
    hdr = {"supplier": "Trans-Solutions (E.A) Ltd", "address": "P.O.Box 17634, Kenya 00500",
           "customer": "New Muthokinju Hardware Limited", "currency": "KES"}
    full = "\n".join(lines)
    hdr["buckets"], hdr["amount_due"] = parse_aging_summary(full)
    m = re.search(r"(\d{2}/\d{2}/\d{4})", full)
    if m: hdr["statement_date"] = m.group(1)
    m = re.search(r"PAYBILL NO:\s*([\d\s]+)", full)
    if m: hdr["paybill"] = re.sub(r"\s+", "", m.group(1))
    return hdr, parse_aging_rows(lines)
def excel_aging_statement(hdr, txs, tab_color, banner_text):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Statement"
    ws.sheet_properties.tabColor = tab_color
    for i, w in enumerate([14, 46, 16, 16], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    banner(ws, 1, 4, banner_text, tab_color)
    lf = Font(name="Arial", bold=True, size=9); vf = Font(name="Arial", size=9)
    lb = PatternFill("solid", start_color=C["info_bg"])
    al = Alignment(horizontal="left", vertical="center", indent=1)
    info = [("Customer", hdr.get("customer", ""), "Statement Date", hdr.get("statement_date", "")),
            ("Supplier", hdr.get("supplier", ""), "Currency", hdr.get("currency", "KES"))]
    for r, (l1, v1, l2, v2) in enumerate(info, start=2):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=1)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=2)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=3)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=4)
        lc = ws.cell(row=r, column=1, value=l1); lc.font=lf; lc.fill=lb; lc.alignment=al
        vc = ws.cell(row=r, column=2, value=v1); vc.font=vf; vc.alignment=al
        lc2 = ws.cell(row=r, column=3, value=l2); lc2.font=lf; lc2.fill=lb; lc2.alignment=al
        vc2 = ws.cell(row=r, column=4, value=v2); vc2.font=vf; vc2.alignment=al
        ws.row_dimensions[r].height = 16
    H = 5
    col_headers(ws, H, ["Date", "Transaction", "Amount (KES)", "Balance (KES)"], tab_color)
    D = H + 1
    for ri, tx in enumerate(txs, start=D):
        is_bf = "Balance forward" in tx["description"]
        is_credit = tx["amount"] is not None and tx["amount"] < 0
        bg = C["info_bg"] if is_bf else (C["credit_row"] if is_credit else (C["alt_row"] if (ri-D) % 2 == 1 else None))
        vals = [tx["date"], tx["description"], tx["amount"], tx["balance"]]
        aligns = ["center", "left", "right", "right"]
        fmts = [None, None, MFMT, MFMT]
        for ci, (v, a, f) in enumerate(zip(vals, aligns, fmts), 1):
            fg = "000000"
            if ci == 3 and v is not None: fg = C["credit_fg"] if v < 0 else C["debit_fg"]
            if ci == 4 and v is not None: fg = C["neg_bal"] if v < 0 else "000000"
            wc(ws, ri, ci, v, align=a, fmt=f, fg=fg, bg=bg, bdr=thin("bottom"))
        ws.row_dimensions[ri].height = 15
    T = D + len(txs)
    ws.merge_cells(f"A{T}:B{T}")
    wc(ws, T, 1, "CLOSING BALANCE", bold=True, fg=C["col_hdr_fg"], bg=tab_color, align="right", bdr=thin(), ind=1)
    c = ws.cell(row=T, column=3, value="")
    c.fill = PatternFill("solid", start_color=tab_color); c.border = thin()
    last_bal = txs[-1]["balance"] if txs else None
    c = ws.cell(row=T, column=4, value=last_bal)
    c.font = Font(name="Arial", bold=True, size=9, color=C["col_hdr_fg"])
    c.fill = PatternFill("solid", start_color=tab_color)
    c.number_format = MFMT; c.alignment = Alignment(horizontal="right", vertical="center"); c.border = thin()
    ws.row_dimensions[T].height = 18
    ws.freeze_panes = f"A{D}"; ws.auto_filter.ref = f"A{H}:D{T-1}"
    ws.page_setup.fitToPage = True; ws.page_setup.fitToWidth = 1
    b = hdr.get("buckets") or {}
    if any(v is not None for v in b.values()) or hdr.get("amount_due") is not None:
        ws2 = wb.create_sheet("Ageing Summary"); ws2.sheet_properties.tabColor = tab_color
        for col, w in [("A", 20), ("B", 18)]: ws2.column_dimensions[col].width = w
        ws2.merge_cells("A1:B1")
        wc(ws2, 1, 1, "AGEING SUMMARY", bold=True, fg=C["col_hdr_fg"], bg=tab_color, align="center", sz=12)
        ws2.row_dimensions[1].height = 26
        rows2 = [("Current", b.get("current")), ("1–30 Days Past Due", b.get("d1_30")),
                 ("31–60 Days Past Due", b.get("d31_60")), ("61–90 Days Past Due", b.get("d61_90")),
                 ("Over 90 Days Past Due", b.get("over90")), ("Amount Due", hdr.get("amount_due"))]
        for ri, (lbl, val) in enumerate(rows2, start=2):
            bold = lbl == "Amount Due"
            wc(ws2, ri, 1, lbl, bold=bold, align="left", bdr=thin("bottom"), bg=C["info_bg"] if bold else None)
            wc(ws2, ri, 2, val, align="right", fmt=MFMT, bdr=thin("bottom"),
               fg=C["neg_bal"] if (val or 0) < 0 else "000000", bg=C["info_bg"] if bold else None)
            ws2.row_dimensions[ri].height = 16
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf
# ══════════════════════════════════════════════════════════════════════════════
# SCANIA EAST AFRICA LIMITED PARSER
# ══════════════════════════════════════════════════════════════════════════════
DATE_SC = re.compile(r"^\d{2}/\d{2}/\d{4}$")
MONEY_SC = re.compile(r"^-?[\d,]+\.\d{2}$")
SKIP_SC = ("Scania East Africa", "VAT Registration", "Mombasa Rd", "P.O.Box", "Nairobi",
           "Nakuru", "Eldoret", "Account statement", "Page Page", "Date and time", "Phone",
           "Fax", "BACS/CHAPS", "Tax registration number", "Customer account",
           "Terms of payment", "Currency", "From date", "To date", "New Muthokinju",
           "P O Box", "Bonje", "Harvester Road", "Tairi Mbili", "Date Invoice Description",
           "+254")
def parse_scania_row(toks):
    if toks[0] in ("Opening", "Closing"):
        try: kes_idx = toks.index("KES")
        except ValueError: kes_idx = None
        bal = pnf(toks[kes_idx+1]) if kes_idx is not None and kes_idx+1 < len(toks) else None
        return {"label": toks[0], "date": "", "invoice": "", "description": "", "due_date": "",
                "debit": None, "credit": None, "balance": bal, "is_summary": True}
    date = toks[0]
    invoice = toks[1] if len(toks) > 1 else ""
    try: kes_idx = toks.index("KES")
    except ValueError: kes_idx = len(toks)
    date_positions = [j for j in range(2, kes_idx) if DATE_SC.match(toks[j])]
    due_idx = date_positions[-1] if date_positions else None
    if due_idx is not None:
        due_date = toks[due_idx]; desc = " ".join(toks[2:due_idx])
    else:
        due_date = ""; desc = " ".join(toks[2:kes_idx])
    money = toks[kes_idx+1:]
    debit = pnf(money[0]) if len(money) > 0 else None
    credit = pnf(money[1]) if len(money) > 1 else None
    balance = pnf(money[2]) if len(money) > 2 else None
    return {"label": "", "date": date, "invoice": invoice, "description": desc, "due_date": due_date,
            "debit": debit, "credit": credit, "balance": balance, "is_summary": False}
def parse_scania(lines):
    hdr = {"customer": "New Muthokinju", "account": "", "terms": "", "currency": "KES",
           "from_date": "", "to_date": "", "tax_reg": ""}
    for ln in lines:
        m = re.search(r"Customer account\s+(\d+)", ln)
        if m: hdr["account"] = m.group(1)
        m = re.search(r"Terms of payment\s+(\d+\s+days\s+net)", ln)
        if m: hdr["terms"] = m.group(1)
        m = re.search(r"From date\s+([\d/]+)", ln)
        if m: hdr["from_date"] = m.group(1)
        m = re.search(r"To date\s+([\d/]+)", ln)
        if m: hdr["to_date"] = m.group(1)
        m = re.search(r"Tax registration number\s+(\S+)", ln)
        if m: hdr["tax_reg"] = m.group(1)
    raw = [l.strip() for l in lines if l.strip()]
    txs = []; i = 0
    while i < len(raw):
        toks = raw[i].split()
        if not toks or not (DATE_SC.match(toks[0]) or toks[0] in ("Opening", "Closing")):
            i += 1; continue
        buf = toks[:]
        while "KES" not in buf and i+1 < len(raw):
            i += 1; buf.extend(raw[i].split())
        tx = parse_scania_row(buf)
        if i+1 < len(raw):
            nxt = raw[i+1]; ntoks = nxt.split()
            if (ntoks and not (DATE_SC.match(ntoks[0]) or ntoks[0] in ("Opening", "Closing"))
                    and not any(k in nxt for k in SKIP_SC) and not MONEY_SC.match(ntoks[0])):
                if not tx["is_summary"]:
                    tx["description"] = (tx["description"] + " " + nxt).strip()
                i += 1
        txs.append(tx); i += 1
    return hdr, txs
def excel_scania(hdr, txs):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Statement"
    ws.sheet_properties.tabColor = C["scania_red"]
    for i, w in enumerate([14, 16, 24, 14, 16, 16, 18], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    banner(ws, 1, 7, "SCANIA EAST AFRICA LIMITED  |  ACCOUNT STATEMENT", C["scania_red"])
    lf = Font(name="Arial", bold=True, size=9); vf = Font(name="Arial", size=9)
    lb = PatternFill("solid", start_color=C["info_bg"])
    al = Alignment(horizontal="left", vertical="center", indent=1)
    info = [("Customer", hdr.get("customer", ""), "Account No.", hdr.get("account", "")),
            ("Terms", hdr.get("terms", ""), "Tax Reg.", hdr.get("tax_reg", "")),
            ("From Date", hdr.get("from_date", ""), "To Date", hdr.get("to_date", ""))]
    for r, (l1, v1, l2, v2) in enumerate(info, start=2):
        for ci, lbl, val in [(1, l1, v1), (4, l2, v2)]:
            ws.merge_cells(start_row=r, start_column=ci, end_row=r, end_column=ci)
            ws.merge_cells(start_row=r, start_column=ci+1, end_row=r, end_column=ci+2)
            lc = ws.cell(row=r, column=ci, value=lbl); lc.font=lf; lc.fill=lb; lc.alignment=al
            vc = ws.cell(row=r, column=ci+1, value=val); vc.font=vf; vc.alignment=al
        ws.row_dimensions[r].height = 16
    H = 5
    col_headers(ws, H, ["Date", "Invoice", "Description", "Due Date", "Debit (KES)", "Credit (KES)", "Balance (KES)"],
                C["scania_red"])
    D = H + 1
    for ri, tx in enumerate(txs, start=D):
        is_sum = tx["is_summary"]
        bg = C["info_bg"] if is_sum else (C["alt_row"] if (ri-D) % 2 == 1 else None)
        if is_sum:
            ws.merge_cells(start_row=ri, start_column=1, end_row=ri, end_column=4)
            wc(ws, ri, 1, tx["label"], bold=True, align="right", bg=bg, bdr=thin("bottom"), ind=1)
            wc(ws, ri, 5, None, bg=bg, bdr=thin("bottom"))
            wc(ws, ri, 6, None, bg=bg, bdr=thin("bottom"))
            wc(ws, ri, 7, tx["balance"], bold=True, align="right", fmt=MFMT, bg=bg, bdr=thin("bottom"))
        else:
            vals = [tx["date"], tx["invoice"], tx["description"], tx["due_date"], tx["debit"], tx["credit"], tx["balance"]]
            aligns = ["center", "center", "left", "center", "right", "right", "right"]
            fmts = [None, None, None, None, MFMT, MFMT, MFMT]
            fgs = ["000000"]*4 + [C["debit_fg"], C["credit_fg"], "000000"]
            for ci, (v, a, f, fg) in enumerate(zip(vals, aligns, fmts, fgs), 1):
                if ci == 7 and v is not None: fg = C["neg_bal"] if v < 0 else "000000"
                wc(ws, ri, ci, v, align=a, fmt=f, fg=fg, bg=bg, bdr=thin("bottom"))
        ws.row_dimensions[ri].height = 15
    T = D + len(txs)
    ws.merge_cells(f"A{T}:D{T}")
    wc(ws, T, 1, "TOTAL", bold=True, fg=C["col_hdr_fg"], bg=C["scania_red"], align="right", bdr=thin(), ind=1)
    for ci, col in [(5, "E"), (6, "F")]:
        c = ws.cell(row=T, column=ci, value=f"=SUM({col}{D}:{col}{T-1})")
        c.font = Font(name="Arial", bold=True, size=9, color=C["col_hdr_fg"])
        c.fill = PatternFill("solid", start_color=C["scania_red"])
        c.number_format = MFMT; c.alignment = Alignment(horizontal="right", vertical="center"); c.border = thin()
    closing = next((t["balance"] for t in reversed(txs) if t["is_summary"] and t["label"] == "Closing"), None)
    c = ws.cell(row=T, column=7, value=closing)
    c.font = Font(name="Arial", bold=True, size=9, color=C["col_hdr_fg"])
    c.fill = PatternFill("solid", start_color=C["scania_red"])
    c.number_format = MFMT; c.alignment = Alignment(horizontal="right", vertical="center"); c.border = thin()
    ws.row_dimensions[T].height = 18
    ws.freeze_panes = f"A{D}"; ws.auto_filter.ref = f"A{H}:G{T-1}"
    ws.page_setup.orientation = "landscape"; ws.page_setup.fitToPage = True; ws.page_setup.fitToWidth = 1
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf
# ══════════════════════════════════════════════════════════════════════════════
# MAPEI EAST AFRICA LTD PARSER
# ══════════════════════════════════════════════════════════════════════════════
DATE_MPI  = re.compile(r"^\d{2}/\d{2}/\d{4}$")
MONEY_MPI = re.compile(r"^-?[\d,]+\.\d{2}$")
INV_MPI   = re.compile(r"^\d{2}[A-Z]{2}\d{4,}$")
SKIP_MPI  = ("Mapei East Africa", "Account statement", "Riverside", "Wing A",
             "Page Page", "Date and time", "Telephone", "Fax", "BACS / CHAPS",
             "Tax registration number", "Customer account", "Terms of payment",
             "Currency KES", "From date", "To date", "Credit limit",
             "Credit available", "P.O. BOX", "P.O.BOX",
             "Date Invoice Description")
def parse_mapei(lines):
    hdr = {"supplier": "Mapei East Africa Ltd.", "customer": "", "account": "",
           "terms": "", "currency": "KES", "credit_limit": None,
           "credit_available": None, "from_date": "", "to_date": "",
           "tax_reg": "", "statement_datetime": "",
           "opening_balance": None, "closing_balance": None}
    for ln in lines:
        m = re.search(r"^(.*?)\s*Customer account\s+(\S+)", ln)
        if m:
            hdr["account"] = m.group(2)
            if m.group(1).strip(): hdr["customer"] = m.group(1).strip()
        m = re.search(r"Terms of payment\s+(.+)$", ln)
        if m: hdr["terms"] = m.group(1).strip()
        m = re.search(r"From date\s+([\d/]+)", ln)
        if m: hdr["from_date"] = m.group(1)
        m = re.search(r"To date\s+([\d/]+)", ln)
        if m: hdr["to_date"] = m.group(1)
        m = re.search(r"Credit limit\s+([\d,]+\.\d{2})", ln)
        if m: hdr["credit_limit"] = pnf(m.group(1))
        m = re.search(r"Credit available\s+([\d,]+\.\d{2})", ln)
        if m: hdr["credit_available"] = pnf(m.group(1))
        m = re.search(r"Tax registration number\s+(\S+)", ln)
        if m: hdr["tax_reg"] = m.group(1)
        m = re.search(r"Date and time\s+([\d/]+\s+[\d:]+)", ln)
        if m: hdr["statement_datetime"] = m.group(1)
    txs = []; pending_desc = []
    for ln in lines:
        ln = ln.strip()
        if not ln: continue
        if ln.startswith("Opening balance"):
            m = re.search(r"(-?[\d,]+\.\d{2})\s*$", ln)
            if m:
                hdr["opening_balance"] = pnf(m.group(1))
                txs.append({"date": "--", "invoice": "", "description": "Opening balance",
                            "due_date": "", "debit": None, "credit": None,
                            "balance": hdr["opening_balance"], "is_summary": True})
            pending_desc = []; continue
        if ln.startswith("Closing balance"):
            m = re.search(r"(-?[\d,]+\.\d{2})\s*$", ln)
            if m: hdr["closing_balance"] = pnf(m.group(1))
            pending_desc = []; continue
        if any(k in ln for k in SKIP_MPI):
            pending_desc = []; continue
        if ln in ("Nairobi", "Kenya", "NAIROBI", "Original"):
            pending_desc = []; continue
        parts = ln.split()
        if not DATE_MPI.match(parts[0]):
            if not any(MONEY_MPI.match(p) for p in parts):
                pending_desc.append(ln)
            else:
                pending_desc = []
            continue
        date = parts[0]; rest = parts[1:]
        try:
            kes_idx = rest.index("KES")
        except ValueError:
            pending_desc = []; continue
        money = rest[kes_idx+1:]
        debit   = pnf(money[0]) if len(money) > 0 else None
        credit  = pnf(money[1]) if len(money) > 1 else None
        balance = pnf(money[2]) if len(money) > 2 else None
        mid = rest[:kes_idx]
        due_date = ""
        if mid and DATE_MPI.match(mid[-1]):
            due_date = mid[-1]; mid = mid[:-1]
        invoice = ""
        if mid and INV_MPI.match(mid[0]):
            invoice = mid[0]; mid = mid[1:]
        desc = " ".join(pending_desc + mid).strip()
        pending_desc = []
        if not desc:
            desc = "Credit Note" if (credit or 0) != 0 else "Invoice"
        txs.append({"date": date, "invoice": invoice, "description": desc,
                    "due_date": due_date, "debit": debit, "credit": credit,
                    "balance": balance, "is_summary": False})
    return hdr, txs
def excel_mapei(hdr, txs):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Statement"
    ws.sheet_properties.tabColor = C["mapei_blue"]
    for i, w in enumerate([13, 14, 34, 13, 16, 16, 18], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    banner(ws, 1, 7, "MAPEI EAST AFRICA LTD  |  ACCOUNT STATEMENT", C["mapei_blue"])
    lf = Font(name="Arial", bold=True, size=9); vf = Font(name="Arial", size=9)
    lb = PatternFill("solid", start_color=C["info_bg"])
    al = Alignment(horizontal="left", vertical="center", indent=1)
    cl = hdr.get("credit_limit"); ca = hdr.get("credit_available")
    info = [("Customer", hdr.get("customer", ""), "Account No.", hdr.get("account", "")),
            ("Terms", hdr.get("terms", ""), "Tax Reg.", hdr.get("tax_reg", "")),
            ("From Date", hdr.get("from_date", ""), "To Date", hdr.get("to_date", "")),
            ("Credit Limit", f"KES {cl:,.2f}" if cl is not None else "",
             "Credit Available", f"KES {ca:,.2f}" if ca is not None else ""),
            ("Statement Date", hdr.get("statement_datetime", ""), "Currency", hdr.get("currency", "KES"))]
    for r, (l1, v1, l2, v2) in enumerate(info, start=2):
        for ci, lbl, val in [(1, l1, v1), (4, l2, v2)]:
            ws.merge_cells(start_row=r, start_column=ci, end_row=r, end_column=ci)
            ws.merge_cells(start_row=r, start_column=ci+1, end_row=r, end_column=ci+2)
            lc = ws.cell(row=r, column=ci, value=lbl); lc.font=lf; lc.fill=lb; lc.alignment=al
            vc = ws.cell(row=r, column=ci+1, value=val); vc.font=vf; vc.alignment=al
        ws.row_dimensions[r].height = 16
    H = 7
    col_headers(ws, H, ["Date", "Invoice", "Description", "Due Date",
                         "Debit (KES)", "Credit (KES)", "Balance (KES)"], C["mapei_blue"])
    D = H + 1
    for ri, tx in enumerate(txs, start=D):
        is_sum = tx["is_summary"]
        is_cr  = (tx["credit"] or 0) != 0
        bg = C["info_bg"] if is_sum else (C["credit_row"] if is_cr else (C["alt_row"] if (ri-D) % 2 == 1 else None))
        if is_sum:
            ws.merge_cells(start_row=ri, start_column=1, end_row=ri, end_column=4)
            wc(ws, ri, 1, tx["description"], bold=True, align="right", bg=bg, bdr=thin("bottom"), ind=1)
            wc(ws, ri, 5, None, bg=bg, bdr=thin("bottom"))
            wc(ws, ri, 6, None, bg=bg, bdr=thin("bottom"))
            wc(ws, ri, 7, tx["balance"], bold=True, align="right", fmt=MFMT, bg=bg, bdr=thin("bottom"))
        else:
            vals = [tx["date"], tx["invoice"], tx["description"], tx["due_date"],
                    tx["debit"], tx["credit"], tx["balance"]]
            aligns = ["center", "center", "left", "center", "right", "right", "right"]
            fmts = [None, None, None, None, MFMT, MFMT, MFMT]
            fgs = ["000000"]*4 + [C["debit_fg"], C["credit_fg"], "000000"]
            for ci, (v, a, f, fg) in enumerate(zip(vals, aligns, fmts, fgs), 1):
                if ci == 7 and v is not None: fg = C["neg_bal"] if v < 0 else "000000"
                wc(ws, ri, ci, v, align=a, fmt=f, fg=fg, bg=bg, bdr=thin("bottom"))
        ws.row_dimensions[ri].height = 15
    T = D + len(txs)
    ws.merge_cells(f"A{T}:D{T}")
    wc(ws, T, 1, "TOTAL / CLOSING BALANCE", bold=True, fg=C["col_hdr_fg"], bg=C["mapei_blue"],
       align="right", bdr=thin(), ind=1)
    for ci, col in [(5, "E"), (6, "F")]:
        c = ws.cell(row=T, column=ci, value=f"=SUM({col}{D}:{col}{T-1})")
        c.font = Font(name="Arial", bold=True, size=9, color=C["col_hdr_fg"])
        c.fill = PatternFill("solid", start_color=C["mapei_blue"])
        c.number_format = MFMT; c.alignment = Alignment(horizontal="right", vertical="center"); c.border = thin()
    closing = hdr.get("closing_balance")
    if closing is None and txs: closing = txs[-1]["balance"]
    c = ws.cell(row=T, column=7, value=closing)
    c.font = Font(name="Arial", bold=True, size=9, color=C["col_hdr_fg"])
    c.fill = PatternFill("solid", start_color=C["mapei_blue"])
    c.number_format = MFMT; c.alignment = Alignment(horizontal="right", vertical="center"); c.border = thin()
    ws.row_dimensions[T].height = 18
    ws.freeze_panes = f"A{D}"; ws.auto_filter.ref = f"A{H}:G{T-1}"
    ws.page_setup.orientation = "landscape"; ws.page_setup.fitToPage = True; ws.page_setup.fitToWidth = 1
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf
# ══════════════════════════════════════════════════════════════════════════════
# BAMBURI CEMENT PLC PARSER
# ══════════════════════════════════════════════════════════════════════════════
MONTHS_BC = {"Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"}
MONEY_BC  = re.compile(r"^-?[\d,]+(\.\d{1,2})?$")
DOCTYPE_BC = {
    ("Data","Transfer"):        "Data Transfer",
    ("Billing","doc.transfer"): "Billing doc.transfer",
    ("Accounting","Docum"):     "Accounting Document",
    ("Customer","Paymen"):      "Customer Payment",
}
def is_money_bc(tok):
    return bool(tok) and bool(MONEY_BC.match(tok)) and any(ch.isdigit() for ch in tok)
def parse_bamburi_row(toks):
    if len(toks) < 4: return None
    bal_tok, amt_tok = toks[-1], toks[-2]
    if not (is_money_bc(bal_tok) and is_money_bc(amt_tok)): return None
    balance, amount = pnf(bal_tok), pnf(amt_tok)
    if balance is None or amount is None: return None
    rest = toks[:-2]
    unit, qty = "", None
    if rest and rest[-1] in ("TO", "EA"):
        unit = rest[-1]
        if len(rest) >= 2:
            try:
                qty = float(rest[-2]); rest = rest[:-2]
            except ValueError:
                rest = rest[:-1]
        else:
            rest = rest[:-1]
    month_idx = next((i for i, tk in enumerate(rest[:4]) if tk in MONTHS_BC), None)
    if month_idx is None or month_idx + 2 >= len(rest): return None
    doc_fields = rest[:month_idx]
    date_str = f"{rest[month_idx]} {rest[month_idx+1]} {rest[month_idx+2]}"
    after_date = rest[month_idx+3:]
    doctype, item_rest = "", after_date
    if len(after_date) >= 3 and (after_date[0], after_date[1]) == ("Customer", "Credit"):
        doctype = "Customer Credit Memo"; item_rest = after_date[3:]
    elif len(after_date) >= 2 and (after_date[0], after_date[1]) in DOCTYPE_BC:
        doctype = DOCTYPE_BC[(after_date[0], after_date[1])]; item_rest = after_date[2:]
    elif len(after_date) >= 2:
        doctype = " ".join(after_date[:2]); item_rest = after_date[2:]
    elif after_date:
        doctype = after_date[0]; item_rest = []
    if len(doc_fields) >= 2 and re.match(r"^\d+$", doc_fields[1]):
        doc_number, delivery_note = doc_fields[0], doc_fields[1]
    elif doc_fields:
        doc_number, delivery_note = " ".join(doc_fields), ""
    else:
        doc_number, delivery_note = "", ""
    return {"doc_number": doc_number, "delivery_note": delivery_note, "date": date_str,
            "doc_type": doctype, "item_text": " ".join(item_rest), "qty": qty, "unit": unit,
            "amount": amount, "balance": balance}
def parse_bamburi(lines):
    hdr = {"customer": "", "po_box": "", "location": "", "statement_date": "",
           "period": "", "opening_balance": None, "closing_balance": None, "pin": ""}
    SKIP_STARTS = ("BAMBURI CEMENT", "Kitui Road", "P.O. Box", "Tel:", "Mobile:",
                   "Ke-customer.service", "PIN:", "Company", "Doc Delivery",
                   "Number Note", "Account statement", "Balance carried forward",
                   "Line Items", "No postings", "Final balance", "KES ")
    for i, ln in enumerate(lines):
        s = ln.strip()
        if s.startswith("PIN:"): hdr["pin"] = s.replace("PIN:", "").strip()
        if s.startswith("Date:"): hdr["statement_date"] = s.replace("Date:", "").strip()
        if s == "Company":
            block = [x.strip() for x in lines[i+1:i+8]]
            block = [x for x in block if not x.startswith(("Date:", "Our Clerk", "Telephone"))]
            if len(block) > 0: hdr["customer"] = block[0]
            if len(block) > 1 and block[1].upper().startswith("PO BOX"): hdr["po_box"] = block[1]
            if len(block) > 2: hdr["location"] = block[2]
        if s.startswith("Account statement for the period"):
            m = re.search(r"from\s+(.+?)\s+to\s+(.+)$", s)
            if m: hdr["period"] = f"{m.group(1)} to {m.group(2)}"
        if s.startswith("Balance carried forward"):
            for j in range(i+1, min(i+3, len(lines))):
                v = pnf(lines[j].strip())
                if v is not None: hdr["opening_balance"] = v; break
        if s.startswith("Final balance"):
            for j in range(i+1, min(i+3, len(lines))):
                m = re.search(r"([\d,]+\.\d{2})", lines[j])
                if m: hdr["closing_balance"] = pnf(m.group(1)); break
    txs = []
    for ln in lines:
        s = ln.strip()
        if not s or any(s.startswith(p) for p in SKIP_STARTS): continue
        toks = s.split()
        if not toks: continue
        tx = parse_bamburi_row(toks)
        if tx is not None:
            txs.append(tx)
        elif txs and not is_money_bc(toks[-1]):
            txs[-1]["item_text"] = (txs[-1]["item_text"] + " " + s).strip()
    return hdr, txs
def excel_bamburi(hdr, txs):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Statement"
    ws.sheet_properties.tabColor = C["bamburi_green"]
    for i, w in enumerate([14, 14, 13, 18, 38, 9, 7, 15, 16], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    banner(ws, 1, 9, "BAMBURI CEMENT PLC  |  ACCOUNT STATEMENT", C["bamburi_green"])
    lf = Font(name="Arial", bold=True, size=9); vf = Font(name="Arial", size=9)
    lb = PatternFill("solid", start_color=C["info_bg"])
    al = Alignment(horizontal="left", vertical="center", indent=1)
    info = [("Customer", hdr.get("customer", ""), "PIN", hdr.get("pin", "")),
            ("PO Box", hdr.get("po_box", ""), "Location", hdr.get("location", "")),
            ("Period", hdr.get("period", ""), "Statement Date", hdr.get("statement_date", ""))]
    for r, (l1, v1, l2, v2) in enumerate(info, start=2):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=1)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=6)
        ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=9)
        lc = ws.cell(row=r, column=1, value=l1); lc.font=lf; lc.fill=lb; lc.alignment=al
        vc = ws.cell(row=r, column=2, value=v1); vc.font=vf; vc.alignment=al
        lc2 = ws.cell(row=r, column=5, value=l2); lc2.font=lf; lc2.fill=lb; lc2.alignment=al
        vc2 = ws.cell(row=r, column=7, value=v2); vc2.font=vf; vc2.alignment=al
        ws.row_dimensions[r].height = 16
    H = 6
    col_headers(ws, H, ["Doc Number", "Delivery Note", "Posting Date", "Doc Type",
                         "Item Text / Ref", "Qty", "Unit", "Amount (KES)", "Running Balance (KES)"],
                C["bamburi_green"])
    D = H + 1
    for ri, tx in enumerate(txs, start=D):
        bg = C["alt_row"] if (ri - D) % 2 == 1 else None
        amt = tx["amount"]
        amt_fg = C["credit_fg"] if (amt is not None and amt < 0) else C["debit_fg"]
        vals   = [tx["doc_number"], tx["delivery_note"], tx["date"], tx["doc_type"],
                  tx["item_text"], tx["qty"], tx["unit"], tx["amount"], tx["balance"]]
        aligns = ["left", "left", "center", "left", "left", "right", "center", "right", "right"]
        fmts   = [None, None, None, None, None, "#,##0.000;(#,##0.000)", None, MFMT, MFMT]
        fgs    = ["000000"]*7 + [amt_fg, "000000"]
        for ci, (v, a, f, fg) in enumerate(zip(vals, aligns, fmts, fgs), 1):
            if ci == 9 and v is not None: fg = C["neg_bal"] if v < 0 else "000000"
            wc(ws, ri, ci, v, align=a, fmt=f, fg=fg, bg=bg, bdr=thin("bottom"))
        ws.row_dimensions[ri].height = 15
    T = D + len(txs)
    ws.merge_cells(f"A{T}:G{T}")
    wc(ws, T, 1, "TOTAL", bold=True, fg=C["col_hdr_fg"], bg=C["bamburi_green"], align="right", bdr=thin(), ind=1)
    c = ws.cell(row=T, column=8, value=f"=SUM(H{D}:H{T-1})")
    c.font = Font(name="Arial", bold=True, size=9, color=C["col_hdr_fg"])
    c.fill = PatternFill("solid", start_color=C["bamburi_green"])
    c.number_format = MFMT; c.alignment = Alignment(horizontal="right", vertical="center"); c.border = thin()
    closing = hdr.get("closing_balance")
    if closing is None and txs: closing = txs[-1]["balance"]
    c = ws.cell(row=T, column=9, value=closing)
    c.font = Font(name="Arial", bold=True, size=9, color=C["col_hdr_fg"])
    c.fill = PatternFill("solid", start_color=C["bamburi_green"])
    c.number_format = MFMT; c.alignment = Alignment(horizontal="right", vertical="center"); c.border = thin()
    ws.row_dimensions[T].height = 18
    ws.freeze_panes = f"A{D}"; ws.auto_filter.ref = f"A{H}:I{T-1}"
    ws.page_setup.orientation = "landscape"; ws.page_setup.fitToPage = True; ws.page_setup.fitToWidth = 1
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf
# ══════════════════════════════════════════════════════════════════════════════
# AUTOXPRESS LIMITED PARSER  (date wraps across 3 lines: "DD-MM-", data, "YYYY")
# ══════════════════════════════════════════════════════════════════════════════
DATE_PREFIX_AX = re.compile(r"^\d{2}-\d{2}-$")
YEAR_AX = re.compile(r"^\d{4}$")
TAXINV_AX = re.compile(r"^\d{15,}$")
LPO_AX = re.compile(r"^(POHH|XNB)")
def parse_autoxpress(lines):
    hdr = {"account_no": "", "customer": "New Muthokinju Hardware", "email": "", "tpin": "",
           "sales_group": "", "pay_term": "", "currency": "KES", "credit_limit": None,
           "statement_date": "", "buckets": {}, "total_balance_due": None}
    full = "\n".join(lines)
    m = re.search(r"A/C No:\s*(\S+)", full)
    if m: hdr["account_no"] = m.group(1)
    m = re.search(r"financemanager@\S+", full)
    if m: hdr["email"] = m.group(0)
    m = re.search(r"TPIN:\s*(\S+)", full)
    if m: hdr["tpin"] = m.group(1)
    m = re.search(r"SALES GROUP:\s*(.+)", full)
    if m: hdr["sales_group"] = m.group(1).strip()
    m = re.search(r"PAY TERM:\s*(.+)", full)
    if m: hdr["pay_term"] = m.group(1).strip()
    m = re.search(r"CREDIT LIMIT:\s*([\d,]+)", full)
    if m: hdr["credit_limit"] = pnf(m.group(1))
    m = re.search(r"STATEMENT DATE:\s*(\S+)", full)
    if m: hdr["statement_date"] = m.group(1)
    m = re.search(r"(-?[\d,]+)\s+(-?[\d,]+)\s+(-?[\d,]+)\s+(-?[\d,]+)\s+(-?[\d,]+)\s+(-?[\d,]+)\s*\n\s*[\d.]+%",
                  full)
    if m:
        hdr["buckets"] = {"120plus": pnf(m.group(1)), "d90": pnf(m.group(2)), "d60": pnf(m.group(3)),
                          "d30": pnf(m.group(4)), "current_balance": pnf(m.group(5)), "balance": pnf(m.group(6))}
    m = re.search(r"TOTAL BALANCE DUE\s*:\s*([\d,\-]+)", full)
    if m: hdr["total_balance_due"] = pnf(m.group(1))
    clean = [l.strip() for l in lines if l.strip()]
    txs = []; i = 0; pending_date = None
    while i < len(clean):
        ln = clean[i]
        if DATE_PREFIX_AX.match(ln):
            pending_date = ln; i += 1; continue
        if YEAR_AX.match(ln) and pending_date and txs:
            txs[-1]["date"] = pending_date + ln
            pending_date = None; i += 1; continue
        if pending_date is not None:
            toks = ln.split()
            if len(toks) >= 4:
                rest, last4 = toks[:-4], toks[-4:]
                tax_invoice = ""; lpo = ""; gp_parts = []
                for tok in rest:
                    if TAXINV_AX.match(tok): tax_invoice = tok
                    elif LPO_AX.match(tok): lpo = (lpo + " " + tok).strip()
                    else: gp_parts.append(tok)
                opening, invoice, payment, cumulative = (pnf(x) for x in last4)
                txs.append({"date": None, "tax_invoice": tax_invoice, "gp_invoice": " ".join(gp_parts),
                            "lpo_no": lpo, "opening_bal": opening, "invoice": invoice,
                            "payment": payment, "cumulative": cumulative})
        i += 1
    return hdr, txs
def excel_autoxpress(hdr, txs):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Statement"
    ws.sheet_properties.tabColor = C["autoxpress_amber"]
    for i, w in enumerate([12, 22, 16, 20, 14, 12, 12, 14], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    banner(ws, 1, 8, "AUTOXPRESS LIMITED  |  ACCOUNT STATEMENT", C["autoxpress_amber"])
    lf = Font(name="Arial", bold=True, size=9); vf = Font(name="Arial", size=9)
    lb = PatternFill("solid", start_color=C["info_bg"])
    al = Alignment(horizontal="left", vertical="center", indent=1)
    info = [("Customer", hdr.get("customer", ""), "A/C No.", hdr.get("account_no", "")),
            ("TPIN", hdr.get("tpin", ""), "Sales Group", hdr.get("sales_group", "")),
            ("Pay Term", hdr.get("pay_term", ""), "Credit Limit",
             f"KES {hdr['credit_limit']:,.0f}" if hdr.get("credit_limit") is not None else ""),
            ("Statement Date", hdr.get("statement_date", ""), "Currency", hdr.get("currency", "KES"))]
    for r, (l1, v1, l2, v2) in enumerate(info, start=2):
        for ci, lbl, val in [(1, l1, v1), (5, l2, v2)]:
            ws.merge_cells(start_row=r, start_column=ci, end_row=r, end_column=ci+1)
            ws.merge_cells(start_row=r, start_column=ci+2, end_row=r, end_column=ci+2)
            lc = ws.cell(row=r, column=ci, value=lbl); lc.font=lf; lc.fill=lb; lc.alignment=al
            vc = ws.cell(row=r, column=ci+2, value=val); vc.font=vf; vc.alignment=al
        ws.row_dimensions[r].height = 16
    H = 7
    col_headers(ws, H, ["Date", "Tax Invoice", "GP Invoice", "LPO Number", "Opening Bal. (KES)",
                        "Invoice (KES)", "Payment (KES)", "Cummulative (KES)"], C["autoxpress_amber"])
    D = H + 1
    for ri, tx in enumerate(txs, start=D):
        is_credit = (tx["payment"] or 0) > 0
        bg = C["credit_row"] if is_credit else (C["alt_row"] if (ri-D) % 2 == 1 else None)
        vals = [tx["date"], tx["tax_invoice"], tx["gp_invoice"], tx["lpo_no"],
                tx["opening_bal"], tx["invoice"], tx["payment"], tx["cumulative"]]
        aligns = ["center", "center", "center", "center", "right", "right", "right", "right"]
        fmts = [None, None, None, None, IFMT, IFMT, IFMT, IFMT]
        fgs = ["000000"]*4 + ["000000", C["debit_fg"], C["credit_fg"], "000000"]
        for ci, (v, a, f, fg) in enumerate(zip(vals, aligns, fmts, fgs), 1):
            wc(ws, ri, ci, v, align=a, fmt=f, fg=fg, bg=bg, bdr=thin("bottom"))
        ws.row_dimensions[ri].height = 15
    T = D + len(txs)
    ws.merge_cells(f"A{T}:D{T}")
    wc(ws, T, 1, "TOTAL", bold=True, fg=C["col_hdr_fg"], bg=C["autoxpress_amber"], align="right", bdr=thin(), ind=1)
    for ci, col in [(6, "F"), (7, "G")]:
        c = ws.cell(row=T, column=ci, value=f"=SUM({col}{D}:{col}{T-1})")
        c.font = Font(name="Arial", bold=True, size=9, color=C["col_hdr_fg"])
        c.fill = PatternFill("solid", start_color=C["autoxpress_amber"])
        c.number_format = IFMT; c.alignment = Alignment(horizontal="right", vertical="center"); c.border = thin()
    last_cum = txs[-1]["cumulative"] if txs else None
    for ci in (5, 8):
        c = ws.cell(row=T, column=ci, value=last_cum if ci == 8 else "")
        c.font = Font(name="Arial", bold=True, size=9, color=C["col_hdr_fg"])
        c.fill = PatternFill("solid", start_color=C["autoxpress_amber"])
        c.number_format = IFMT; c.alignment = Alignment(horizontal="right", vertical="center"); c.border = thin()
    ws.row_dimensions[T].height = 18
    ws.freeze_panes = f"A{D}"; ws.auto_filter.ref = f"A{H}:H{T-1}"
    ws.page_setup.orientation = "landscape"; ws.page_setup.fitToPage = True; ws.page_setup.fitToWidth = 1
    b = hdr.get("buckets") or {}
    if any(v is not None for v in b.values()):
        ws2 = wb.create_sheet("Ageing Summary"); ws2.sheet_properties.tabColor = C["autoxpress_amber"]
        for col, w in [("A", 22), ("B", 18)]: ws2.column_dimensions[col].width = w
        ws2.merge_cells("A1:B1")
        wc(ws2, 1, 1, "AGEING SUMMARY", bold=True, fg=C["col_hdr_fg"], bg=C["autoxpress_amber"], align="center", sz=12)
        ws2.row_dimensions[1].height = 26
        rows2 = [("120+ Days", b.get("120plus")), ("90 Days", b.get("d90")), ("60 Days", b.get("d60")),
                 ("30 Days", b.get("d30")), ("Current Balance", b.get("current_balance")),
                 ("Balance", b.get("balance")), ("Total Balance Due", hdr.get("total_balance_due"))]
        for ri, (lbl, val) in enumerate(rows2, start=2):
            bold = lbl == "Total Balance Due"
            wc(ws2, ri, 1, lbl, bold=bold, align="left", bdr=thin("bottom"), bg=C["info_bg"] if bold else None)
            wc(ws2, ri, 2, val, align="right", fmt=IFMT, bdr=thin("bottom"),
               fg=C["neg_bal"] if (val or 0) < 0 else "000000", bg=C["info_bg"] if bold else None)
            ws2.row_dimensions[ri].height = 16
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf
# ══════════════════════════════════════════════════════════════════════════════
# GATHANI LIMITED PARSER
# ══════════════════════════════════════════════════════════════════════════════
DATE_GTH = re.compile(r"^\d{1,2}-[A-Za-z]{3}-\d{2}$")
MONEY_GTH = re.compile(r"^[\d,]+\.\d{2}$")
def parse_gathani(lines):
    hdr = {"supplier": "Gathani Limited", "customer": "", "address": "", "contact": "",
           "as_at": "", "pin": "", "buckets": {}}
    for ln in lines:
        m = re.search(r"NEW MUTHOKINJU", ln, re.I)
        if m: hdr["customer"] = "New Muthokinju Hardware Ltd"
        m = re.search(r"As At\s*:\s*(\S+)", ln)
        if m: hdr["as_at"] = m.group(1)
        m = re.search(r"PIN NO\s*:\s*(\S+)", ln)
        if m: hdr["pin"] = m.group(1)
        m = re.search(r"MARIANA ROSE[:\s]+(\S+)", ln, re.I)
        if m: hdr["contact"] = m.group(1)
        if re.search(r"\d{1,3}(?:,\d{3})*\.\d{2}\s+Dr", ln) and "Days" not in ln:
            nums = re.findall(r"([\d,]+\.\d{2})\s*(Dr|Cr)?", ln)
            keys = ["0_30","30_60","60_90","90_120","above_120","balance_due"]
            for ki, (amt, drcr) in enumerate(nums):
                if ki < len(keys):
                    v = pnf(amt)
                    if drcr and drcr.upper() == "CR" and v: v = -v
                    hdr["buckets"][keys[ki]] = v
    txs = []
    skip_kw = {"Date","Particulars","Document","Ref","Nett",
               "Amount","GATHANI","Gathani","Page","STATEMENT","PO BOX",
               "Tel:","EMAIL","PIN","Pls","Industrial","0-30","Bank","867386",
               "KYANGOMBE","MARIANA"}
    for ln in lines:
        ln = ln.strip()
        if not ln: continue
        if any(k in ln for k in skip_kw): continue
        parts = ln.split()
        if not parts: continue
        if parts[0] == "Opening" and "Balance" in ln:
            nums = [(pnf(p), parts[i+1] if i+1 < len(parts) else "") for i,p in enumerate(parts) if MONEY_GTH.match(p)]
            if nums:
                amt, drcr = nums[0]
                txs.append({
                    "date": "--", "particulars": "Opening Balance",
                    "doc_ref": "", "debit": None, "credit": None,
                    "nett_amount": amt, "drcr": drcr, "is_opening": True
                })
            continue
        if not DATE_GTH.match(parts[0]): continue
        date = parts[0]
        rest = parts[1:]
        money_vals = []
        desc_parts = []
        i = 0
        while i < len(rest):
            if MONEY_GTH.match(rest[i]):
                drcr = rest[i+1] if i+1 < len(rest) and rest[i+1] in ("Dr","Cr","DR","CR") else ""
                money_vals.append((pnf(rest[i]), drcr))
                if drcr: i += 2
                else: i += 1
            else:
                desc_parts.append(rest[i])
                i += 1
        doc_ref = ""
        particulars = ""
        if desc_parts:
            if re.match(r"^\d+$", desc_parts[-1]):
                doc_ref = desc_parts[-1]
                particulars = " ".join(desc_parts[:-1])
            else:
                particulars = " ".join(desc_parts)
        is_credit = "Credit" in particulars or "credit" in particulars
        debit = credit = nett = None
        if len(money_vals) >= 2:
            nett, nett_drcr = money_vals[-1]
            tx_amt, tx_drcr = money_vals[-2]
            if is_credit or tx_drcr.upper() == "CR":
                credit = tx_amt
            else:
                debit = tx_amt
        elif len(money_vals) == 1:
            nett, nett_drcr = money_vals[0]
        else:
            continue
        txs.append({
            "date": date, "particulars": particulars, "doc_ref": doc_ref,
            "debit": debit, "credit": credit, "nett_amount": nett,
            "drcr": nett_drcr if len(money_vals) >= 1 else "", "is_opening": False
        })
    return hdr, txs
def excel_gathani(hdr, txs):
    wb = openpyxl.Workbook(); ws = wb.active
    ws.title = "Statement"; ws.sheet_properties.tabColor = C["gathani_maroon"]
    for i, w in enumerate([14, 20, 14, 18, 18, 20], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    banner(ws, 1, 6, "GATHANI LIMITED  |  STATEMENT OF ACCOUNT", C["gathani_maroon"])
    lf = Font(name="Arial", bold=True, size=9)
    vf = Font(name="Arial", size=9)
    lb = PatternFill("solid", start_color=C["info_bg"])
    al = Alignment(horizontal="left", vertical="center", indent=1)
    info = [
        ("Customer", hdr.get("customer",""), "As At", hdr.get("as_at","")),
        ("PIN No.", hdr.get("pin",""),        "Contact", hdr.get("contact","")),
    ]
    for r, (l1,v1,l2,v2) in enumerate(info, start=2):
        for ci,lbl,val in [(1,l1,v1),(4,l2,v2)]:
            ws.merge_cells(start_row=r,start_column=ci,end_row=r,end_column=ci+1)
            lc = ws.cell(row=r, column=ci, value=lbl); lc.font=lf; lc.fill=lb; lc.alignment=al
            vc = ws.cell(row=r, column=ci+2, value=val); vc.font=vf; vc.alignment=al
        ws.row_dimensions[r].height = 16
    H = 4
    col_headers(ws, H, ["Date","Particulars","Doc / Ref No.",
                         "Debit (KES)","Credit (KES)","Nett Amount (KES)"], C["gathani_maroon"])
    D = H + 1
    for ri, tx in enumerate(txs, start=D):
        is_open = tx.get("is_opening", False)
        is_cr   = tx["credit"] is not None and not is_open
        bg = C["info_bg"] if is_open else (C["credit_row"] if is_cr else (C["alt_row"] if (ri-D)%2==1 else None))
        vals   = [tx["date"], tx["particulars"], tx["doc_ref"],
                  tx["debit"], tx["credit"], tx["nett_amount"]]
        aligns = ["center","left","center","right","right","right"]
        fmts   = [None, None, None, MFMT, MFMT, MFMT]
        fgs    = ["000000","000000","000000",C["debit_fg"],C["credit_fg"],"000000"]
        for ci,(v,a,f,fg) in enumerate(zip(vals,aligns,fmts,fgs),1):
            wc(ws, ri, ci, v, align=a, fmt=f, fg=fg, bg=bg, bdr=thin("bottom"))
        ws.row_dimensions[ri].height = 15
    T = D + len(txs)
    ws.merge_cells(f"A{T}:C{T}")
    wc(ws,T,1,"TOTAL",bold=True,fg=C["col_hdr_fg"],bg=C["gathani_maroon"],align="right",bdr=thin(),ind=1)
    for ci, col in [(4,"D"),(5,"E")]:
        c = ws.cell(row=T,column=ci,value=f"=SUM({col}{D}:{col}{T-1})")
        c.font=Font(name="Arial",bold=True,size=9,color=C["col_hdr_fg"])
        c.fill=PatternFill("solid",start_color=C["gathani_maroon"])
        c.number_format=MFMT; c.alignment=Alignment(horizontal="right",vertical="center"); c.border=thin()
    last_nett = txs[-1]["nett_amount"] if txs else None
    c = ws.cell(row=T,column=6,value=last_nett)
    c.font=Font(name="Arial",bold=True,size=9,color=C["col_hdr_fg"])
    c.fill=PatternFill("solid",start_color=C["gathani_maroon"])
    c.number_format=MFMT; c.alignment=Alignment(horizontal="right",vertical="center"); c.border=thin()
    ws.row_dimensions[T].height = 18
    b = hdr.get("buckets", {})
    if any(v is not None for v in b.values()):
        ws2 = wb.create_sheet("Ageing Summary"); ws2.sheet_properties.tabColor = C["gathani_maroon"]
        for col,w in [("A",22),("B",18)]: ws2.column_dimensions[col].width = w
        ws2.merge_cells("A1:B1")
        wc(ws2,1,1,"AGEING SUMMARY",bold=True,fg=C["col_hdr_fg"],bg=C["gathani_maroon"],align="center",sz=12)
        ws2.row_dimensions[1].height=26
        rows2 = [
            ("0–30 Days",    b.get("0_30")),
            ("30–60 Days",   b.get("30_60")),
            ("60–90 Days",   b.get("60_90")),
            ("90–120 Days",  b.get("90_120")),
            ("Above 120 Days",b.get("above_120")),
            ("Balance Due",  b.get("balance_due")),
        ]
        for ri,(lbl,val) in enumerate(rows2,start=2):
            bold = lbl == "Balance Due"
            bg2  = C["info_bg"] if bold else None
            wc(ws2,ri,1,lbl,bold=bold,align="left",bdr=thin("bottom"),bg=bg2)
            wc(ws2,ri,2,val,align="right",fmt=MFMT,bdr=thin("bottom"),
               fg=C["neg_bal"] if (val or 0)<0 else "000000",bg=bg2)
            ws2.row_dimensions[ri].height=16
    ws.freeze_panes=f"A{D}"; ws.auto_filter.ref=f"A{H}:F{T-1}"
    ws.page_setup.orientation="landscape"; ws.page_setup.fitToPage=True; ws.page_setup.fitToWidth=1
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf
# ══════════════════════════════════════════════════════════════════════════════
# ANCO TRADING LTD PARSER
# ══════════════════════════════════════════════════════════════════════════════
DATE_ANCO  = re.compile(r"^\d{2}/\d{2}/\d{4}$")
MONEY_ANCO = re.compile(r"^[\d,]+\.\d{2}$")
def parse_anco(lines):
    hdr = {"supplier": "Anco Trading Ltd", "customer": "", "client_code": "", "email": "",
           "pin": "", "credit_days": "", "credit_limit": "",
           "period": "", "total_outstanding": None, "buckets": {}}
    for ln in lines:
        m = re.search(r"Client Name\s*:\s*(.+?)(?:\s{3,}|Client Adress|Client Postal|$)", ln)
        if m:
            name = m.group(1).strip()
            name = re.split(r'\s{2,}', name)[0].strip()
            if name: hdr["customer"] = name
        if ln.strip() == "LIMITED" and hdr.get("customer","").endswith("HARDWARE"):
            hdr["customer"] += " LIMITED"
        m = re.search(r"Client Code\s*:\s*(\S+)", ln)
        if m: hdr["client_code"] = m.group(1)
        m = re.search(r"Client Email\s*:\s*(\S+)", ln)
        if m: hdr["email"] = m.group(1)
        m = re.search(r"Client Pin No\s*:\s*(\S+)", ln)
        if m: hdr["pin"] = m.group(1)
        m = re.search(r"Client Credit Days\s+([\d.]+)", ln)
        if m: hdr["credit_days"] = m.group(1)
        m = re.search(r"Client Credit Limit\s+([\d,]+(?:\.\d+)?)", ln)
        if m: hdr["credit_limit"] = m.group(1)
        m = re.search(r"BETWEEN\s+(\S+)\s+AND\s+(\S+)", ln)
        if m: hdr["period"] = f"{m.group(1)} to {m.group(2)}"
        m = re.search(r"TOTAL OUTSTANDING\s+([\d,]+\.\d{2})", ln)
        if m: hdr["total_outstanding"] = pnf(m.group(1))
        m = re.search(r"0-30 days\s+30-60 days\s+60-90 days\s+<90 days\s+Total", ln, re.I)
        if m: _ageing_next = True
    for i, ln in enumerate(lines):
        if re.search(r"0-30 days", ln, re.I) and i+1 < len(lines):
            nums = re.findall(r"[\d,]+\.\d{2}", lines[i+1])
            keys = ["0_30","30_60","60_90","lt90","total"]
            for ki,n in enumerate(nums):
                if ki < len(keys): hdr["buckets"][keys[ki]] = pnf(n)
    txs = []
    skip_kw = {"TRANSACTION","DOCNO","REF NO","DOC DATE","DRCR","DEBIT","CREDIT",
               "ANCO","Client","BETWEEN","TOTAL","AGEING","0-30","days","OUTSTANDING"}
    for ln in lines:
        ln = ln.strip()
        if not ln: continue
        if any(k in ln for k in skip_kw): continue
        parts = ln.split()
        if not parts: continue
        if parts[0] == "OPENING" and len(parts) >= 3:
            nums = [p for p in parts if MONEY_ANCO.match(p)]
            date_m = [p for p in parts if DATE_ANCO.match(p)]
            drcr = "DR" if "DR" in parts else "CR"
            debit = pnf(nums[0]) if len(nums) >= 3 else None
            credit = pnf(nums[1]) if len(nums) >= 3 else None
            balance = pnf(nums[2]) if len(nums) >= 3 else pnf(nums[0]) if nums else None
            txs.append({
                "transaction": "OPENING BALANCE", "doc_no": "",
                "ref_no": "Opening Balance", "doc_date": date_m[0] if date_m else "--",
                "drcr": drcr, "debit": debit, "credit": credit, "balance": balance,
                "is_opening": True
            })
            continue
        tx_types = ("SINV","RECIEPT","RECEIPT","JOURNAL","CREDIT","DEBIT")
        if not any(parts[0].upper().startswith(t) for t in tx_types): continue
        money_idx = [i for i,p in enumerate(parts) if MONEY_ANCO.match(p)]
        date_idx  = [i for i,p in enumerate(parts) if DATE_ANCO.match(p)]
        if len(money_idx) < 2: continue
        balance = pnf(parts[money_idx[-1]])
        credit  = pnf(parts[money_idx[-2]])
        debit   = pnf(parts[money_idx[-3]]) if len(money_idx) >= 3 else None
        drcr_idx = next((i for i,p in enumerate(parts) if p in ("DR","CR")), None)
        drcr = parts[drcr_idx] if drcr_idx is not None else ("CR" if (credit or 0)>0 else "DR")
        doc_date = parts[date_idx[-1]] if date_idx else ""
        ref_no = parts[date_idx[0]] if len(date_idx) >= 2 else ""
        pivot = min((date_idx + money_idx + ([drcr_idx] if drcr_idx else [])), default=len(parts))
        mid_parts = parts[1:pivot]
        doc_no  = mid_parts[0] if len(mid_parts) >= 1 else ""
        ref_note= " ".join(mid_parts[1:]) if len(mid_parts) > 1 else ""
        txs.append({
            "transaction": parts[0], "doc_no": doc_no,
            "ref_no": ref_note or ref_no, "doc_date": doc_date,
            "drcr": drcr, "debit": debit, "credit": credit, "balance": balance,
            "is_opening": False
        })
    return hdr, txs
def excel_anco(hdr, txs):
    wb = openpyxl.Workbook(); ws = wb.active
    ws.title = "Statement"; ws.sheet_properties.tabColor = C["anco_teal"]
    for i,w in enumerate([14,10,20,14,6,16,16,18],1):
        ws.column_dimensions[get_column_letter(i)].width = w
    banner(ws,1,8,"ANCO TRADING LTD  |  CUSTOMER STATEMENT OF ACCOUNTS",C["anco_teal"])
    lf=Font(name="Arial",bold=True,size=9); vf=Font(name="Arial",size=9)
    lb=PatternFill("solid",start_color=C["info_bg"])
    al=Alignment(horizontal="left",vertical="center",indent=1)
    info=[
        ("Customer",  hdr.get("customer",""),      "Period",        hdr.get("period","")),
        ("Client Code",hdr.get("client_code",""),  "PIN No.",       hdr.get("pin","")),
        ("Email",     hdr.get("email",""),          "Credit Limit",  hdr.get("credit_limit","")),
        ("Credit Days",hdr.get("credit_days",""),  "Outstanding",
         f"KES {hdr['total_outstanding']:,.2f}" if hdr.get("total_outstanding") else ""),
    ]
    for r,(l1,v1,l2,v2) in enumerate(info,start=2):
        for ci,lbl,val in [(1,l1,v1),(5,l2,v2)]:
            ws.merge_cells(start_row=r,start_column=ci,end_row=r,end_column=ci+1)
            lc=ws.cell(row=r,column=ci,value=lbl); lc.font=lf; lc.fill=lb; lc.alignment=al
            vc=ws.cell(row=r,column=ci+2,value=val); vc.font=vf; vc.alignment=al
        ws.row_dimensions[r].height=16
    H=6
    col_headers(ws,H,["Transaction","Doc No.","Ref / Note","Doc Date",
                       "DR/CR","Debit (KES)","Credit (KES)","Balance (KES)"],C["anco_teal"])
    D=H+1
    for ri,tx in enumerate(txs,start=D):
        is_open = tx.get("is_opening",False)
        is_cr   = tx["drcr"]=="CR" and not is_open
        is_jnl  = tx["transaction"].upper()=="JOURNAL"
        bg = (C["info_bg"]   if is_open else
              C["credit_row"] if is_cr and not is_jnl else
              C["receipt_row"] if is_jnl else
              (C["alt_row"] if (ri-D)%2==1 else None))
        vals   = [tx["transaction"],tx["doc_no"],tx["ref_no"],tx["doc_date"],
                  tx["drcr"],tx["debit"],tx["credit"],tx["balance"]]
        aligns = ["center","center","left","center","center","right","right","right"]
        fmts   = [None,None,None,None,None,MFMT,MFMT,MFMT]
        fgs    = ["000000","000000","000000","000000",
                  C["debit_fg"] if tx["drcr"]=="DR" else C["credit_fg"],
                  C["debit_fg"],C["credit_fg"],"000000"]
        for ci,(v,a,f,fg) in enumerate(zip(vals,aligns,fmts,fgs),1):
            if ci==8 and v is not None: fg=C["neg_bal"] if v<0 else C["credit_fg"]
            wc(ws,ri,ci,v,align=a,fmt=f,fg=fg,bg=bg,bdr=thin("bottom"))
        ws.row_dimensions[ri].height=15
    T=D+len(txs)
    ws.merge_cells(f"A{T}:E{T}")
    wc(ws,T,1,"TOTAL OUTSTANDING",bold=True,fg=C["col_hdr_fg"],bg=C["anco_teal"],align="right",bdr=thin(),ind=1)
    for ci,col in [(6,"F"),(7,"G")]:
        c=ws.cell(row=T,column=ci,value=f"=SUM({col}{D}:{col}{T-1})")
        c.font=Font(name="Arial",bold=True,size=9,color=C["col_hdr_fg"])
        c.fill=PatternFill("solid",start_color=C["anco_teal"])
        c.number_format=MFMT; c.alignment=Alignment(horizontal="right",vertical="center"); c.border=thin()
    outstanding = hdr.get("total_outstanding")
    c=ws.cell(row=T,column=8,value=outstanding)
    c.font=Font(name="Arial",bold=True,size=9,color=C["col_hdr_fg"])
    c.fill=PatternFill("solid",start_color=C["anco_teal"])
    c.number_format=MFMT; c.alignment=Alignment(horizontal="right",vertical="center"); c.border=thin()
    ws.row_dimensions[T].height=18
    b=hdr.get("buckets",{})
    if any(v is not None for v in b.values()):
        ws2=wb.create_sheet("Ageing Summary"); ws2.sheet_properties.tabColor=C["anco_teal"]
        for col,w in [("A",22),("B",18)]: ws2.column_dimensions[col].width=w
        ws2.merge_cells("A1:B1")
        wc(ws2,1,1,"AGEING ANALYSIS",bold=True,fg=C["col_hdr_fg"],bg=C["anco_teal"],align="center",sz=12)
        ws2.row_dimensions[1].height=26
        rows2=[("0–30 Days",b.get("0_30")),("30–60 Days",b.get("30_60")),
               ("60–90 Days",b.get("60_90")),("<90 Days",b.get("lt90")),
               ("Total Outstanding",hdr.get("total_outstanding"))]
        for ri,(lbl,val) in enumerate(rows2,start=2):
            bold=lbl=="Total Outstanding"
            bg2=C["info_bg"] if bold else None
            wc(ws2,ri,1,lbl,bold=bold,align="left",bdr=thin("bottom"),bg=bg2)
            wc(ws2,ri,2,val,align="right",fmt=MFMT,bdr=thin("bottom"),
               fg=C["neg_bal"] if (val or 0)<0 else "000000",bg=bg2)
            ws2.row_dimensions[ri].height=16
    ws.freeze_panes=f"A{D}"; ws.auto_filter.ref=f"A{H}:H{T-1}"
    ws.page_setup.orientation="landscape"; ws.page_setup.fitToPage=True; ws.page_setup.fitToWidth=1
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf
# ══════════════════════════════════════════════════════════════════════════════
# CROWN PAINTS KENYA PLC PARSER
# ══════════════════════════════════════════════════════════════════════════════
DATE_CRN  = re.compile(r"^\d{1,2}-[A-Za-z]{3}-\d{2}$")
MONEY_CRN = re.compile(r"^-?[\d,]+\.\d{2}$")
DESC_CRN  = ("A/R Credit Memo", "A/R Invoice", "Incoming Payment")
def parse_crown(lines):
    full = "\n".join(lines)
    hdr = {"supplier": "Crown Paints Kenya PLC", "account_code": "", "customer": "",
           "period_from": "", "period_to": "", "credit_limit": None, "credit_days": "",
           "opening_balance": None, "total_debit": None, "total_credit": None,
           "total_outstanding": None, "ageing": {}}
    m = re.search(r"\b(RCNE\d+)\b", full)
    if m: hdr["account_code"] = m.group(1)
    m = re.search(r"\b(?:RCNE\d+\s+)?([A-Z][A-Z0-9 &/.,'()-]{5,}(?:LTD|LIMITED))\b", full)
    if m: hdr["customer"] = m.group(1).strip()
    m = re.search(r"(\d{1,2}-[A-Za-z]{3}-\d{4})\s*(?:TO)?\s*[\r\n]*\s*(?:TO\s*)?(\d{1,2}-[A-Za-z]{3}-\d{4})", full)
    if m: hdr["period_from"], hdr["period_to"] = m.group(1), m.group(2)
    m = re.search(r"Credit Limit\s+([\d,]+\.\d{2})", full)
    if not m:
        m = re.search(r"([\d,]+\.\d{2})\s*[\r\n]+\s*Credit Limit", full)
    if m: hdr["credit_limit"] = pnf(m.group(1))
    m = re.search(r"(\d+\s+Days From Statement)", full)
    if m: hdr["credit_days"] = m.group(1)
    m = re.search(r"Opening Balance\s+([\d,]+\.\d{2})", full)
    if m: hdr["opening_balance"] = pnf(m.group(1))
    m = re.search(r"Total\s*[\r\n]+\s*([\d,]+\.\d{2})\s+([\d,]+\.\d{2})", full)
    if m: hdr["total_debit"], hdr["total_credit"] = pnf(m.group(1)), pnf(m.group(2))
    m = re.search(r"Total Outstanding Balance:\s*[\r\n]*\s*([\d,]+\.\d{2})", full)
    if m: hdr["total_outstanding"] = pnf(m.group(1))
    m = re.search(
        r"0\s*-\s*30 Days\s+31\s*-\s*60 Days\s+61\s*-\s*90 Days\s+91-120 Days\s+121\s*\+\s*Days\s+Amount Due\s*"
        r"([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})",
        full)
    if m:
        keys = ["0_30", "31_60", "61_90", "91_120", "121_plus", "amount_due"]
        for k, v in zip(keys, m.groups()): hdr["ageing"][k] = pnf(v)
    products = []
    for ln in lines:
        pm = re.match(
            r"^([A-Z]{2,6}-[A-Z0-9 /]+?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+"
            r"([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})$", ln.strip())
        if pm:
            products.append({
                "category": pm.group(1).strip(),
                "sales_volume": pnf(pm.group(2)), "return_volume": pnf(pm.group(3)),
                "net_volume": pnf(pm.group(4)), "sales_amount": pnf(pm.group(5)),
                "return_amount": pnf(pm.group(6)), "net_amount": pnf(pm.group(7)),
            })
    txs = []
    if hdr["opening_balance"] is not None:
        txs.append({"date": "--", "desc": "Opening Balance", "doc_no": "", "reference": "",
                    "debit": None, "credit": None, "balance": hdr["opening_balance"], "is_opening": True})
    def trailing_money_count(tokens):
        idx = len(tokens)
        while idx > 0 and MONEY_CRN.match(tokens[idx - 1]):
            idx -= 1
        return len(tokens) - idx
    n = len(lines)
    i = 0
    while i < n:
        ln = lines[i].strip()
        if not ln:
            i += 1; continue
        parts = ln.split()
        if not parts or not DATE_CRN.match(parts[0]):
            i += 1; continue
        date = parts[0]
        rest_line = ln[len(parts[0]):].strip()
        desc = None
        rest_after_desc = ""
        for kw in DESC_CRN:
            if rest_line.startswith(kw):
                desc = kw
                rest_after_desc = rest_line[len(kw):].strip()
                break
        if desc is None:
            i += 1; continue
        tok = rest_after_desc.split()
        doc_no = tok[0] if tok else ""
        collected = tok[1:]
        j = i + 1
        while trailing_money_count(collected) < 3 and j < n:
            nxt = lines[j].strip()
            if not nxt:
                j += 1; continue
            nparts = nxt.split()
            if DATE_CRN.match(nparts[0]) or nparts[0] == "Total" or "Total Outstanding" in nxt:
                break
            collected.extend(nparts)
            j += 1
        if len(collected) >= 3 and all(MONEY_CRN.match(x) for x in collected[-3:]):
            money_tokens = collected[-3:]
            reference_tokens = collected[:-3]
        else:
            money_tokens = []
            reference_tokens = collected
        reference = " ".join(reference_tokens).strip()
        debit = credit = balance = None
        if len(money_tokens) == 3:
            debit, credit, balance = pnf(money_tokens[0]), pnf(money_tokens[1]), pnf(money_tokens[2])
        txs.append({"date": date, "desc": desc, "doc_no": doc_no, "reference": reference,
                    "debit": debit, "credit": credit, "balance": balance, "is_opening": False})
        i = j if j > i + 1 else i + 1
    return hdr, txs, products
def excel_crown(hdr, txs, products):
    wb = openpyxl.Workbook(); ws = wb.active
    ws.title = "Statement"; ws.sheet_properties.tabColor = C["crown_red"]
    for i, w in enumerate([12, 16, 12, 32, 16, 14, 18], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    banner(ws, 1, 7, "CROWN PAINTS KENYA PLC  |  CUSTOMER STATEMENT", C["crown_red"])
    lf = Font(name="Arial", bold=True, size=9); vf = Font(name="Arial", size=9)
    lb = PatternFill("solid", start_color=C["info_bg"])
    al = Alignment(horizontal="left", vertical="center", indent=1)
    period = f"{hdr.get('period_from','')} to {hdr.get('period_to','')}".strip(" to")
    info = [
        ("Customer", hdr.get("customer", ""), "Account Code", hdr.get("account_code", "")),
        ("Period", period, "Credit Days", hdr.get("credit_days", "")),
        ("Credit Limit",
         f"KES {hdr['credit_limit']:,.2f}" if hdr.get("credit_limit") is not None else "",
         "Opening Balance",
         f"{hdr['opening_balance']:,.2f}" if hdr.get("opening_balance") is not None else ""),
    ]
    for r, (l1, v1, l2, v2) in enumerate(info, start=2):
        for ci, lbl, val in [(1, l1, v1), (5, l2, v2)]:
            ws.merge_cells(start_row=r, start_column=ci, end_row=r, end_column=ci + 1)
            lc = ws.cell(row=r, column=ci, value=lbl); lc.font = lf; lc.fill = lb; lc.alignment = al
            vc = ws.cell(row=r, column=ci + 2, value=val); vc.font = vf; vc.alignment = al
        ws.row_dimensions[r].height = 16
    H = 6
    col_headers(ws, H, ["Date", "Description", "Doc No", "Reference",
                         "Debit (KES)", "Credit (KES)", "Cumulative Bal. (KES)"], C["crown_red"])
    D = H + 1
    for ri, tx in enumerate(txs, start=D):
        is_open = tx.get("is_opening", False)
        is_cr = tx["desc"] in ("A/R Credit Memo", "Incoming Payment") and not is_open
        bg = C["info_bg"] if is_open else (C["credit_row"] if is_cr else (C["alt_row"] if (ri - D) % 2 == 1 else None))
        vals = [tx["date"], tx["desc"], tx["doc_no"], tx["reference"], tx["debit"], tx["credit"], tx["balance"]]
        aligns = ["center", "center", "center", "left", "right", "right", "right"]
        fmts = [None, None, None, None, MFMT, MFMT, MFMT]
        for ci, (v, a, f) in enumerate(zip(vals, aligns, fmts), 1):
            fg = "000000"
            if ci == 5 and v is not None: fg = C["debit_fg"]
            if ci == 6 and v is not None: fg = C["credit_fg"]
            wc(ws, ri, ci, v, align=a, fmt=f, fg=fg, bg=bg, bdr=thin("bottom"))
        ws.row_dimensions[ri].height = 15
    T = D + len(txs)
    ws.merge_cells(f"A{T}:D{T}")
    wc(ws, T, 1, "TOTAL", bold=True, fg=C["col_hdr_fg"], bg=C["crown_red"], align="right", bdr=thin(), ind=1)
    for ci, col in [(5, "E"), (6, "F")]:
        c = ws.cell(row=T, column=ci, value=f"=SUM({col}{D}:{col}{T - 1})")
        c.font = Font(name="Arial", bold=True, size=9, color=C["col_hdr_fg"])
        c.fill = PatternFill("solid", start_color=C["crown_red"])
        c.number_format = MFMT; c.alignment = Alignment(horizontal="right", vertical="center"); c.border = thin()
    last_bal = txs[-1]["balance"] if txs else hdr.get("total_outstanding")
    c = ws.cell(row=T, column=7, value=last_bal)
    c.font = Font(name="Arial", bold=True, size=9, color=C["col_hdr_fg"])
    c.fill = PatternFill("solid", start_color=C["crown_red"])
    c.number_format = MFMT; c.alignment = Alignment(horizontal="right", vertical="center"); c.border = thin()
    ws.row_dimensions[T].height = 18
    ws.freeze_panes = f"A{D}"; ws.auto_filter.ref = f"A{H}:G{T - 1}"
    ws.page_setup.orientation = "landscape"; ws.page_setup.fitToPage = True; ws.page_setup.fitToWidth = 1
    ag = hdr.get("ageing", {})
    if any(v is not None for v in ag.values()):
        ws2 = wb.create_sheet("Ageing Summary"); ws2.sheet_properties.tabColor = C["crown_red"]
        for col, w in [("A", 22), ("B", 18)]: ws2.column_dimensions[col].width = w
        ws2.merge_cells("A1:B1")
        wc(ws2, 1, 1, "AGEING SUMMARY", bold=True, fg=C["col_hdr_fg"], bg=C["crown_red"], align="center", sz=12)
        ws2.row_dimensions[1].height = 26
        rows2 = [
            ("0 - 30 Days", ag.get("0_30")),
            ("31 - 60 Days", ag.get("31_60")),
            ("61 - 90 Days", ag.get("61_90")),
            ("91 - 120 Days", ag.get("91_120")),
            ("121+ Days", ag.get("121_plus")),
            ("Amount Due", ag.get("amount_due")),
        ]
        for ri, (lbl, val) in enumerate(rows2, start=2):
            bold = lbl == "Amount Due"
            bg2 = C["info_bg"] if bold else None
            wc(ws2, ri, 1, lbl, bold=bold, align="left", bdr=thin("bottom"), bg=bg2)
            wc(ws2, ri, 2, val, align="right", fmt=MFMT, bdr=thin("bottom"),
               fg=C["neg_bal"] if (val or 0) < 0 else "000000", bg=bg2)
            ws2.row_dimensions[ri].height = 16
    if products:
        ws3 = wb.create_sheet("Product Summary"); ws3.sheet_properties.tabColor = C["crown_red"]
        for i, w in enumerate([26, 14, 14, 14, 16, 16, 16], 1):
            ws3.column_dimensions[get_column_letter(i)].width = w
        banner(ws3, 1, 7, "PRODUCT-WISE GROUP SUMMARY", C["crown_red"])
        col_headers(ws3, 2, ["Category", "Sales Volume", "Return Volume", "Net Volume",
                              "Sales Amount (KES)", "Return Amount (KES)", "Net Amount (KES)"], C["crown_red"])
        for ri, p in enumerate(products, start=3):
            bg = C["alt_row"] if (ri - 3) % 2 == 1 else None
            vals = [p["category"], p["sales_volume"], p["return_volume"], p["net_volume"],
                    p["sales_amount"], p["return_amount"], p["net_amount"]]
            aligns = ["left", "right", "right", "right", "right", "right", "right"]
            fmts = [None, IFMT, IFMT, IFMT, MFMT, MFMT, MFMT]
            for ci, (v, a, f) in enumerate(zip(vals, aligns, fmts), 1):
                wc(ws3, ri, ci, v, align=a, fmt=f, bdr=thin("bottom"), bg=bg)
            ws3.row_dimensions[ri].height = 15
        T3 = 3 + len(products)
        wc(ws3, T3, 1, "TOTAL", bold=True, fg=C["col_hdr_fg"], bg=C["crown_red"], align="right", bdr=thin())
        for ci, col in [(2, "B"), (3, "C"), (4, "D"), (5, "E"), (6, "F"), (7, "G")]:
            c = ws3.cell(row=T3, column=ci, value=f"=SUM({col}3:{col}{T3 - 1})")
            c.font = Font(name="Arial", bold=True, size=9, color=C["col_hdr_fg"])
            c.fill = PatternFill("solid", start_color=C["crown_red"])
            c.number_format = IFMT if ci in (2, 3, 4) else MFMT
            c.alignment = Alignment(horizontal="right", vertical="center"); c.border = thin()
        ws3.freeze_panes = "A3"; ws3.auto_filter.ref = f"A2:G{T3 - 1}"
        ws3.page_setup.orientation = "landscape"; ws3.page_setup.fitToPage = True; ws3.page_setup.fitToWidth = 1
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf
# ══════════════════════════════════════════════════════════════════════════════
# CONFIRMATION OF ACCOUNTS PARSER  (Doshi / Ultra Multiware — two-column layout)
# ══════════════════════════════════════════════════════════════════════════════
DATE_COA  = re.compile(r"^\d{1,2}-[A-Za-z]{3}-\d{2}$")
MONEY_COA = re.compile(r"^[\d,]+\.\d{2}$")
CU_COA    = re.compile(r"^\d{16,19}$")
SKIP_COA = {"Date","Particulars","Debit","Amount","Credit","Dear","Sub:",
            "Given","Kindly","Please","period.","I/We","P.O","P.O.BOX",
            "NAIROBI","To","From:","continued","Page"}
def parse_coa(lines):
    hdr = {"supplier": "", "customer": "New Muthokinju Hardware Limited",
           "date": "", "period": "", "currency": "KES", "closing_balance": None}
    for ln in lines[:8]:
        m = re.search(r"From:\s*(.+)", ln)
        if m:
            candidate = m.group(1).strip()
            candidate = re.split(r"P\.O\.?\s*Box|P\.O\.?BOX", candidate)[0].strip()
            if candidate: hdr["supplier"] = candidate
        m = re.search(r"Date\s*:\s*([\d\w-]+)", ln)
        if m: hdr["date"] = m.group(1)
    for ln in lines:
        m = re.search(r"(\d{1,2}-[A-Za-z]{3}-\d{2})\s+to\s+(\d{1,2}-[A-Za-z]{3}-\d{2})", ln)
        if m: hdr["period"] = f"{m.group(1)} to {m.group(2)}"; break
    debits  = []
    credits = []
    ageing  = {}
    for ln in lines:
        ln = ln.strip()
        if not ln: continue
        parts = ln.split()
        if not parts: continue
        if "Closing" in ln and "Balance" in ln:
            nums = [p for p in parts if MONEY_COA.match(p)]
            if nums: hdr["closing_balance"] = pnf(nums[-1])
            continue
        money_count = sum(1 for p in parts if MONEY_COA.match(p))
        if money_count >= 4 and not DATE_COA.match(parts[0]):
            nums = [pnf(p) for p in parts if MONEY_COA.match(p)]
            if len(nums) >= 5:
                keys = ["120_days","90_days","60_days","30_days","current","total_due"]
                for ki, v in enumerate(nums):
                    if ki < len(keys): ageing[keys[ki]] = v
            continue
        if any(k in ln for k in SKIP_COA) and not DATE_COA.match(parts[0]):
            continue
        if not parts or not DATE_COA.match(parts[0]):
            continue
        date_positions = [i for i,p in enumerate(parts) if DATE_COA.match(p)]
        money_positions= [i for i,p in enumerate(parts) if MONEY_COA.match(p)]
        if not money_positions: continue
        first_money = money_positions[0]
        debit_date  = parts[0]
        debit_amount= pnf(parts[first_money])
        debit_mid   = parts[1:first_money]
        cu_inv = next((p for p in debit_mid if CU_COA.match(p)), "")
        debit_desc  = " ".join(p for p in debit_mid if p not in (cu_inv, "-"))
        debits.append({"date": debit_date, "particulars": debit_desc,
                       "cu_invoice": cu_inv, "amount": debit_amount})
        if len(date_positions) >= 2:
            cr_date_pos  = date_positions[1]
            cr_money_pos = next((i for i in money_positions if i > cr_date_pos), None)
            if cr_money_pos is not None:
                cr_date   = parts[cr_date_pos]
                cr_amount = pnf(parts[cr_money_pos])
                cr_desc   = " ".join(parts[cr_date_pos+1:cr_money_pos])
                credits.append({"date": cr_date, "particulars": cr_desc, "amount": cr_amount})
    return hdr, debits, credits, ageing
def excel_coa(hdr, debits, credits, ageing):
    supplier = hdr.get("supplier", "Supplier Statement")
    COA_COLOR = "1A237E"
    COA_DR    = "BF360C"
    COA_CR    = "1B5E20"
    wb = openpyxl.Workbook(); ws = wb.active
    ws.title = "Statement"; ws.sheet_properties.tabColor = COA_COLOR
    for i, w in enumerate([13, 28, 22, 16, 2, 13, 28, 16], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.merge_cells("A1:H1")
    c = ws["A1"]
    c.value = f"{supplier.upper()}  |  CONFIRMATION OF ACCOUNTS"
    c.font  = Font(name="Arial", bold=True, size=13, color="FFFFFF")
    c.fill  = PatternFill("solid", start_color=COA_COLOR)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30
    lf = Font(name="Arial", bold=True, size=9); vf = Font(name="Arial", size=9)
    lb = PatternFill("solid", start_color=C["info_bg"])
    al = Alignment(horizontal="left", vertical="center", indent=1)
    info = [
        ("Customer",   hdr.get("customer",""),  "Statement Date", hdr.get("date","")),
        ("Period",     hdr.get("period",""),     "Currency",       hdr.get("currency","KES")),
        ("Closing Balance", f"KES {hdr['closing_balance']:,.2f}" if hdr.get("closing_balance") else "", "", ""),
    ]
    for r,(l1,v1,l2,v2) in enumerate(info, start=2):
        for ci,lbl,val in [(1,l1,v1),(5,l2,v2)]:
            ws.merge_cells(start_row=r,start_column=ci,end_row=r,end_column=ci+1)
            lc=ws.cell(row=r,column=ci,value=lbl); lc.font=lf; lc.fill=lb; lc.alignment=al
            vc=ws.cell(row=r,column=ci+2,value=val); vc.font=vf; vc.alignment=al
        ws.row_dimensions[r].height = 16
    H = 5
    for ci, h in enumerate(["Date","Particulars","CU Invoice No.","Debit (KES)"], 1):
        wc(ws, H, ci, h, bold=True, fg="FFFFFF", bg=COA_DR, align="center", bdr=thin())
    ws.cell(row=H, column=5).fill = PatternFill("solid", start_color="EEEEEE")
    for ci, h in zip([6,7,8], ["Date","Particulars","Credit (KES)"]):
        wc(ws, H, ci, h, bold=True, fg="FFFFFF", bg=COA_CR, align="center", bdr=thin())
    ws.row_dimensions[H].height = 22
    D = H + 1
    max_rows = max(len(debits), len(credits))
    for ri in range(max_rows):
        r = D + ri
        bg_dr = "FBE9E7" if ri % 2 == 1 else None
        bg_cr = "E8F5E9" if ri % 2 == 1 else None
        if ri < len(debits):
            tx = debits[ri]
            is_opening = "Opening" in str(tx.get("particulars",""))
            bg_dr = C["info_bg"] if is_opening else bg_dr
            wc(ws,r,1,tx["date"],     align="center", bg=bg_dr, bdr=thin("bottom"))
            wc(ws,r,2,tx["particulars"],align="left",  bg=bg_dr, bdr=thin("bottom"))
            wc(ws,r,3,tx["cu_invoice"],align="center", bg=bg_dr, bdr=thin("bottom"),
               fg="555555", sz=8)
            wc(ws,r,4,tx["amount"],   align="right",  bg=bg_dr, bdr=thin("bottom"),
               fmt=MFMT, fg=C["debit_fg"], bold=is_opening)
        else:
            for ci in range(1,5):
                ws.cell(row=r,column=ci).border = thin("bottom")
        ws.cell(row=r,column=5).fill = PatternFill("solid",start_color="EEEEEE")
        if ri < len(credits):
            tx = credits[ri]
            wc(ws,r,6,tx["date"],        align="center", bg=bg_cr, bdr=thin("bottom"))
            wc(ws,r,7,tx["particulars"], align="left",   bg=bg_cr, bdr=thin("bottom"))
            wc(ws,r,8,tx["amount"],      align="right",  bg=bg_cr, bdr=thin("bottom"),
               fmt=MFMT, fg=C["credit_fg"])
        else:
            for ci in range(6,9):
                ws.cell(row=r,column=ci).border = thin("bottom")
        ws.row_dimensions[r].height = 15
    T = D + max_rows
    ws.merge_cells(f"A{T}:C{T}")
    wc(ws,T,1,"TOTAL DEBITS",bold=True,fg="FFFFFF",bg=COA_DR,align="right",bdr=thin(),ind=1)
    c=ws.cell(row=T,column=4,value=f"=SUM(D{D}:D{T-1})")
    c.font=Font(name="Arial",bold=True,size=9,color="FFFFFF")
    c.fill=PatternFill("solid",start_color=COA_DR)
    c.number_format=MFMT; c.alignment=Alignment(horizontal="right",vertical="center"); c.border=thin()
    ws.cell(row=T,column=5).fill = PatternFill("solid",start_color="EEEEEE")
    ws.merge_cells(f"F{T}:G{T}")
    wc(ws,T,6,"TOTAL CREDITS",bold=True,fg="FFFFFF",bg=COA_CR,align="right",bdr=thin(),ind=1)
    c=ws.cell(row=T,column=8,value=f"=SUM(H{D}:H{T-1})")
    c.font=Font(name="Arial",bold=True,size=9,color="FFFFFF")
    c.fill=PatternFill("solid",start_color=COA_CR)
    c.number_format=MFMT; c.alignment=Alignment(horizontal="right",vertical="center"); c.border=thin()
    ws.row_dimensions[T].height=18
    SUM_R = T + 1
    ws.merge_cells(f"A{SUM_R}:C{SUM_R}")
    wc(ws,SUM_R,1,"CLOSING BALANCE",bold=True,fg="FFFFFF",bg=COA_COLOR,align="right",bdr=thin(),ind=1)
    c=ws.cell(row=SUM_R,column=4,value=hdr.get("closing_balance"))
    c.font=Font(name="Arial",bold=True,size=10,color="FFFFFF")
    c.fill=PatternFill("solid",start_color=COA_COLOR)
    c.number_format=MFMT; c.alignment=Alignment(horizontal="right",vertical="center"); c.border=thin()
    for ci in [5,6,7,8]:
        ws.cell(row=SUM_R,column=ci).fill=PatternFill("solid",start_color=COA_COLOR)
        ws.cell(row=SUM_R,column=ci).border=thin()
    ws.row_dimensions[SUM_R].height=20
    ws.freeze_panes = f"A{D}"
    ws.page_setup.orientation="landscape"; ws.page_setup.fitToPage=True; ws.page_setup.fitToWidth=1
    if ageing:
        ws2 = wb.create_sheet("Ageing Summary"); ws2.sheet_properties.tabColor=COA_COLOR
        for col,w in [("A",20),("B",18)]: ws2.column_dimensions[col].width=w
        ws2.merge_cells("A1:B1")
        wc(ws2,1,1,"AGEING SUMMARY",bold=True,fg="FFFFFF",bg=COA_COLOR,align="center",sz=12)
        ws2.row_dimensions[1].height=26
        for ci,h in enumerate(["Ageing Bucket","Amount (KES)"],1):
            wc(ws2,2,ci,h,bold=True,fg="FFFFFF",bg=COA_COLOR,align="center",bdr=thin())
        ws2.row_dimensions[2].height=20
        bucket_labels = {
            "120_days":"120+ Days","90_days":"90 Days","60_days":"60 Days",
            "30_days":"30 Days","current":"Current Balance","total_due":"Total Amount Due"
        }
        for ri,(key,lbl) in enumerate(bucket_labels.items(),start=3):
            val = ageing.get(key)
            bold = key == "total_due"
            bg2 = C["info_bg"] if bold else (C["alt_row"] if ri%2==0 else None)
            wc(ws2,ri,1,lbl,bold=bold,align="left",bdr=thin("bottom"),bg=bg2)
            wc(ws2,ri,2,val,align="right",fmt=MFMT,bdr=thin("bottom"),
               fg=C["debit_fg"],bg=bg2,bold=bold)
            ws2.row_dimensions[ri].height=16
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf
# ══════════════════════════════════════════════════════════════════════════════
# BASCO PRODUCTS (K) LTD PARSER  (multi-account merged statements: NE29, NE35 …)
# ══════════════════════════════════════════════════════════════════════════════
DATE_BSC  = re.compile(r"^\d{2}/\d{2}/\d{4}$")
MONEY_BSC = re.compile(r"^[\d,]+\.\d{2}-?$")
TYPES_BSC = {"Invoice", "Paymnt", "Adjmnt", "Credit", "Crnote", "Journal", "Dbnote"}
def parse_basco(lines):
    hdr = {"customer_name": "", "statement_date": "", "currency": "KES",
           "terms": "", "accounts": [], "aging": []}
    txs = []
    acct = ""
    seen_aging = set()
    i = 0
    while i < len(lines):
        ln = lines[i].strip()
        m = re.search(r"Customer #:\s*(\S+)", ln)
        if m:
            acct = m.group(1)
            if acct not in hdr["accounts"]: hdr["accounts"].append(acct)
        m = re.search(r"Customer Name:\s*(.+?)(?:\s+P\.O|\s+Bill To|$)", ln)
        if m and not hdr["customer_name"]: hdr["customer_name"] = m.group(1).strip()
        m = re.search(r"Statement Date:\s*(\d{2}/\d{2}/\d{4})", ln)
        if m: hdr["statement_date"] = m.group(1)
        m = re.search(r"Terms:\s*(.+?)(?:\s+Contact|$)", ln)
        if m and not hdr["terms"]: hdr["terms"] = m.group(1).strip()
        if "Current" in ln and "Amount Due" in ln:
            for k in range(i + 1, min(i + 4, len(lines))):
                nums = [t for t in lines[k].split() if MONEY_BSC.match(t)]
                if len(nums) == 6:
                    if acct not in seen_aging:
                        seen_aging.add(acct)
                        hdr["aging"].append({"acct": acct,
                            "current": pnf(nums[0]), "d30": pnf(nums[1]),
                            "d60": pnf(nums[2]), "d90": pnf(nums[3]),
                            "d120": pnf(nums[4]), "due": pnf(nums[5])})
                    i = k
                    break
            i += 1; continue
        toks = ln.split()
        if not toks or not DATE_BSC.match(toks[0]):
            i += 1; continue
        ms = []
        while len(toks) > 1 and MONEY_BSC.match(toks[-1]):
            ms.insert(0, toks.pop())
        date = toks[0]; rest = toks[1:]
        inv = ""; typ = ""; ref = ""
        if rest and rest[0] in TYPES_BSC:
            typ = rest[0]; ref = " ".join(rest[1:])
        elif len(rest) > 1 and rest[1] in TYPES_BSC:
            inv = rest[0]; typ = rest[1]; ref = " ".join(rest[2:])
        elif rest:
            typ = rest[0]; ref = " ".join(rest[1:])
        if not typ:
            i += 1; continue
        debit = credit = bal = None
        if len(ms) >= 2 and not ms[-1].endswith("-"):
            bal = pnf(ms[-1]); ms = ms[:-1]
        for t in ms:
            if t.endswith("-"): credit = pnf(t)
            else:               debit  = pnf(t)
        txs.append({"acct": acct, "date": date, "invoice": inv, "type": typ,
                    "ref": ref, "debit": debit, "credit": credit, "balance": bal})
        i += 1
    return hdr, txs
def excel_basco(hdr, txs):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Statement"
    ws.sheet_properties.tabColor = C["basco_blue"]
    for i, w in enumerate([8, 12, 10, 9, 40, 16, 16, 16], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    banner(ws, 1, 8, "BASCO PRODUCTS (K) LTD  |  CUSTOMER STATEMENT", C["basco_blue"])
    lf = Font(name="Arial", bold=True, size=9); vf = Font(name="Arial", size=9)
    lb = PatternFill("solid", start_color=C["info_bg"])
    al = Alignment(horizontal="left", vertical="center", indent=1)
    info = [("Customer Name", hdr["customer_name"], "Statement Date", hdr["statement_date"]),
            ("Account(s)", ", ".join(hdr["accounts"]), "Terms", hdr["terms"])]
    for r, (l1, v1, l2, v2) in enumerate(info, start=2):
        for ci, val, bold in [(1, l1, True), (3, v1, False), (5, l2, True), (7, v2, False)]:
            ws.merge_cells(start_row=r, start_column=ci, end_row=r, end_column=ci + 1)
            c = ws.cell(row=r, column=ci, value=val)
            c.font = lf if bold else vf; c.alignment = al
            if bold: c.fill = lb
        ws.row_dimensions[r].height = 16
    H = 5
    col_headers(ws, H, ["A/C", "Doc Date", "Invoice No", "Type",
                        "Reference / Order No", "Debit (KES)", "Credit (KES)",
                        "Balance (KES)"], C["basco_blue"])
    D = H + 1
    for ri, tx in enumerate(txs, start=D):
        cr = tx["type"] in ("Paymnt", "Credit")
        bg = C["credit_row"] if cr else (C["alt_row"] if (ri - D) % 2 == 1 else None)
        vals = [tx["acct"], tx["date"], tx["invoice"], tx["type"], tx["ref"],
                tx["debit"], tx["credit"], tx["balance"]]
        aligns = ["center", "center", "center", "center", "left", "right", "right", "right"]
        fmts = [None, None, None, None, None, MFMT, MFMT, MFMT]
        fgs = ["000000"] * 5 + [C["debit_fg"], C["credit_fg"], "000000"]
        for ci, (v, a, f, fg) in enumerate(zip(vals, aligns, fmts, fgs), 1):
            wc(ws, ri, ci, v, align=a, fmt=f, fg=fg, bg=bg, bdr=thin("bottom"))
        ws.row_dimensions[ri].height = 15
    T = D + len(txs)
    ws.merge_cells(f"A{T}:E{T}")
    wc(ws, T, 1, "TOTAL  (Debit + Credit = net movement)", bold=True, fg=C["col_hdr_fg"],
       bg=C["header_dark"], align="right", bdr=thin(), ind=1)
    for ci, col, formula in [(6, "F", f"=SUM(F{D}:F{T-1})"),
                             (7, "G", f"=SUM(G{D}:G{T-1})"),
                             (8, "H", f"=F{T}+G{T}")]:
        c = ws.cell(row=T, column=ci, value=formula)
        c.font = Font(name="Arial", bold=True, size=9, color=C["col_hdr_fg"])
        c.fill = PatternFill("solid", start_color=C["header_dark"])
        c.number_format = MFMT
        c.alignment = Alignment(horizontal="right", vertical="center"); c.border = thin()
    ws.row_dimensions[T].height = 18
    ws.freeze_panes = f"A{D}"; ws.auto_filter.ref = f"A{H}:H{T-1}"
    ws.page_setup.orientation = "landscape"; ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    if hdr["aging"]:
        ag = wb.create_sheet("Ageing Summary")
        ag.sheet_properties.tabColor = C["basco_blue"]
        for i, w in enumerate([10, 18, 18, 18, 18, 18, 20], 1):
            ag.column_dimensions[get_column_letter(i)].width = w
        banner(ag, 1, 7, "BASCO PRODUCTS  |  AGEING SUMMARY", C["basco_blue"])
        col_headers(ag, 3, ["A/C", "Current", "30 Days", "60 Days", "90 Days",
                            "120 Days", "Amount Due"], C["basco_blue"])
        r = 4
        for a in hdr["aging"]:
            vals = [a["acct"], a["current"], a["d30"], a["d60"], a["d90"], a["d120"], a["due"]]
            for ci, v in enumerate(vals, 1):
                wc(ag, r, ci, v, align="center" if ci == 1 else "right",
                   fmt=None if ci == 1 else MFMT, bdr=thin("bottom"))
            r += 1
        wc(ag, r, 1, "TOTAL", bold=True, fg=C["col_hdr_fg"], bg=C["header_dark"],
           align="right", bdr=thin(), ind=1)
        for ci in range(2, 8):
            L = get_column_letter(ci)
            c = ag.cell(row=r, column=ci, value=f"=SUM({L}4:{L}{r-1})")
            c.font = Font(name="Arial", bold=True, size=9, color=C["col_hdr_fg"])
            c.fill = PatternFill("solid", start_color=C["header_dark"])
            c.number_format = MFMT
            c.alignment = Alignment(horizontal="right", vertical="center"); c.border = thin()
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf
# ══════════════════════════════════════════════════════════════════════════════
# MAIN CONVERT
# ══════════════════════════════════════════════════════════════════════════════
def convert(pdf_bytes):
    lines=[]
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        pages=len(pdf.pages)
        for page in pdf.pages:
            t=page.extract_text()
            if t: lines.extend(t.split("\n"))
    fmt=detect(lines)
    if fmt=="dolphin":
        hdr,txs=parse_dolphin(lines)
        return fmt,len(txs),0,pages,excel_dolphin(hdr,txs)
    elif fmt=="mombasa":
        hdr,txs=parse_mombasa(lines)
        return fmt,len(txs),0,pages,excel_mombasa(hdr,txs)
    elif fmt=="karsan":
        hdr,txs=parse_karsan(lines)
        return fmt,len(txs),0,pages,excel_karsan(hdr,txs)
    elif fmt=="national":
        hdr,txs,chqs=parse_national(lines)
        return fmt,len(txs),len(chqs),pages,excel_national(hdr,txs,chqs)
    elif fmt=="savannah":
        hdr,txs=parse_savannah(lines)
        return fmt,len(txs),0,pages,excel_savannah(hdr,txs)
    elif fmt=="skylife":
        hdr,txs=parse_skylife(lines)
        return fmt,len(txs),0,pages,excel_skylife(hdr,txs)
    elif fmt=="springtech":
        hdr,txs=parse_springtech(lines)
        return fmt,len(txs),0,pages,excel_aging_statement(hdr,txs,C["springtech_indigo"],
                "SPRINGTECH (K) LTD  |  CUSTOMER STATEMENT")
    elif fmt=="transsolutions":
        hdr,txs=parse_transsolutions(lines)
        return fmt,len(txs),0,pages,excel_aging_statement(hdr,txs,C["transsolutions_cyan"],
                "TRANS-SOLUTIONS (E.A) LTD  |  CUSTOMER STATEMENT")
    elif fmt=="scania":
        hdr,txs=parse_scania(lines)
        return fmt,len(txs),0,pages,excel_scania(hdr,txs)
    elif fmt=="mapei":
        hdr,txs=parse_mapei(lines)
        return fmt,len(txs),0,pages,excel_mapei(hdr,txs)
    elif fmt=="autoxpress":
        hdr,txs=parse_autoxpress(lines)
        return fmt,len(txs),0,pages,excel_autoxpress(hdr,txs)
    elif fmt=="bamburi":
        hdr,txs=parse_bamburi(lines)
        return fmt,len(txs),0,pages,excel_bamburi(hdr,txs)
    elif fmt=="gathani":
        hdr,txs=parse_gathani(lines)
        return fmt,len(txs),0,pages,excel_gathani(hdr,txs)
    elif fmt=="anco":
        hdr,txs=parse_anco(lines)
        return fmt,len(txs),0,pages,excel_anco(hdr,txs)
    elif fmt=="coa":
        hdr,debits,credits,ageing=parse_coa(lines)
        return fmt,len(debits)+len(credits),0,pages,excel_coa(hdr,debits,credits,ageing)
    elif fmt=="crown":
        hdr,txs,products=parse_crown(lines)
        return fmt,len(txs),0,pages,excel_crown(hdr,txs,products)
    elif fmt=="basco":
        hdr,txs=parse_basco(lines)
        return fmt,len(txs),0,pages,excel_basco(hdr,txs)
    return "unknown",0,0,pages,None
# ══════════════════════════════════════════════════════════════════════════════
# UNIVERSAL AI PARSER  (Claude — works on any supplier statement)
# ══════════════════════════════════════════════════════════════════════════════
import urllib.request, urllib.error
def _get_api_key():
    try:
        key = st.secrets.get("ANTHROPIC_API_KEY", "")
        if key: return key.strip()
    except Exception:
        pass
    key = os.environ.get("ANTHROPIC_API_KEY", "")
    if key: return key.strip()
    return st.session_state.get("_anthropic_key", "").strip() or None
_AI_SYSTEM = """You are a financial document parser. Your only job is to extract data from a supplier statement PDF (provided as raw text) and return it as clean JSON.
Return ONLY a single JSON object with this exact structure — no markdown, no explanation, no preamble:
{
  "supplier": "string — supplier/company name",
  "customer": "string — customer name",
  "statement_date": "string — as at / statement date",
  "period": "string — date range if shown, else empty string",
  "currency": "string — e.g. KES, USD",
  "opening_balance": number or null,
  "closing_balance": number or null,
  "total_debit": number or null,
  "total_credit": number or null,
  "transactions": [
    {
      "date": "string",
      "description": "string",
      "reference": "string — invoice/doc/ref number",
      "debit": number or null,
      "credit": number or null,
      "balance": number or null,
      "note": "string — any extra info e.g. vehicle, LPO, type"
    }
  ],
  "ageing": {
    "0_30": number or null,
    "31_60": number or null,
    "61_90": number or null,
    "91_120": number or null,
    "above_120": number or null,
    "total": number or null
  },
  "extra_fields": {}
}
Rules:
- All monetary values must be plain numbers (no commas, no currency symbols).
- If a field is not present in the document, use null.
- Include EVERY line item in transactions — opening balance, invoices, receipts, credit notes, journals.
- For debit/credit: use positive numbers. Receipts/payments go in credit; invoices go in debit.
- If the statement has a single amount column with DR/CR indicators, split into debit or credit.
- Dates: preserve exactly as shown in the document.
- extra_fields: capture any columns that don't fit the standard schema.
"""
def _llm_extract(raw_text, api_key):
    payload = json.dumps({
        "model": "claude-sonnet-4-6",
        "max_tokens": 4000,
        "system": _AI_SYSTEM,
        "messages": [{"role": "user",
                       "content": f"Extract the statement data from this text:\n\n{raw_text[:12000]}"}]
    }).encode()
    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=payload,
        headers={
            "Content-Type":      "application/json",
            "anthropic-version": "2023-06-01",
            "x-api-key":         api_key,
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            result = json.loads(resp.read())
    except urllib.error.HTTPError as e:
        body = e.read().decode()
        if e.code == 401:
            raise ValueError("Invalid API key — check your Anthropic API key and try again.") from e
        if e.code == 429:
            raise ValueError("Rate limit reached — wait a moment and try again.") from e
        raise ValueError(f"API error {e.code}: {body[:200]}") from e
    raw = result["content"][0]["text"].strip()
    if raw.startswith("```"):
        raw = re.sub(r"^```[a-z]*\n?", "", raw)
        raw = re.sub(r"\n?```$",       "", raw)
    return json.loads(raw)
def _fmt_num(v):
    if v is None: return ""
    try: return f"{float(v):,.2f}"
    except: return str(v)
def excel_universal(data, filename):
    supplier  = data.get("supplier","Unknown Supplier") or "Unknown Supplier"
    customer  = data.get("customer","") or ""
    stmt_date = data.get("statement_date","") or ""
    period    = data.get("period","") or ""
    currency  = data.get("currency","KES") or "KES"
    txs       = data.get("transactions") or []
    ageing    = data.get("ageing") or {}
    AI_COLOR  = "4A235A"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Statement"
    ws.sheet_properties.tabColor = AI_COLOR
    col_widths = [14, 32, 20, 16, 16, 18, 28]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.merge_cells("A1:G1")
    c = ws["A1"]
    c.value = f"{supplier.upper()}  |  CUSTOMER STATEMENT  |  ✨ AI Extracted"
    c.font  = Font(name="Arial", bold=True, size=13, color="FFFFFF")
    c.fill  = PatternFill("solid", start_color=AI_COLOR)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30
    ws.merge_cells("A2:G2")
    c = ws["A2"]
    c.value = f"Converted by Claude AI  ·  File: {filename}"
    c.font  = Font(name="Arial", italic=True, size=9, color="FFFFFF")
    c.fill  = PatternFill("solid", start_color="6A1B7A")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 16
    lf = Font(name="Arial", bold=True, size=9)
    vf = Font(name="Arial", size=9)
    lb = PatternFill("solid", start_color=C["info_bg"])
    al = Alignment(horizontal="left", vertical="center", indent=1)
    info_rows = [
        ("Supplier",  supplier,  "Customer",       customer),
        ("As At",     stmt_date, "Period",          period),
        ("Currency",  currency,  "Opening Balance", _fmt_num(data.get("opening_balance"))),
        ("Total Debit", _fmt_num(data.get("total_debit")),
         "Total Credit", _fmt_num(data.get("total_credit"))),
    ]
    for r, (l1,v1,l2,v2) in enumerate(info_rows, start=3):
        for ci,lbl,val in [(1,l1,v1),(5,l2,v2)]:
            ws.merge_cells(start_row=r,start_column=ci,end_row=r,end_column=ci+1)
            lc=ws.cell(row=r,column=ci,value=lbl); lc.font=lf; lc.fill=lb; lc.alignment=al
            vc=ws.cell(row=r,column=ci+2,value=val); vc.font=vf; vc.alignment=al
        ws.row_dimensions[r].height=16
    H = 8
    hdrs = ["Date","Description","Reference",f"Debit ({currency})",f"Credit ({currency})",f"Balance ({currency})","Note"]
    for ci,h in enumerate(hdrs,1):
        c=ws.cell(row=H,column=ci,value=h)
        c.font=Font(name="Arial",bold=True,color="FFFFFF",size=10)
        c.fill=PatternFill("solid",start_color=AI_COLOR)
        c.alignment=Alignment(horizontal="center",vertical="center")
        c.border=thin()
    ws.row_dimensions[H].height=22
    D = H + 1
    for ri, tx in enumerate(txs, start=D):
        is_cr  = (tx.get("credit") or 0) > 0 and not (tx.get("debit") or 0)
        is_opn = "opening" in str(tx.get("description","")).lower() or \
                 "balance b" in str(tx.get("description","")).lower()
        bg = (C["info_bg"]    if is_opn else
              C["credit_row"] if is_cr  else
              (C["alt_row"]   if (ri-D)%2==1 else None))
        vals   = [tx.get("date"), tx.get("description"), tx.get("reference"),
                  tx.get("debit"), tx.get("credit"), tx.get("balance"), tx.get("note")]
        aligns = ["center","left","center","right","right","right","left"]
        fmts   = [None,None,None,MFMT,MFMT,MFMT,None]
        fgs    = ["000000","000000","555555",
                  C["debit_fg"], C["credit_fg"],
                  C["neg_bal"] if (tx.get("balance") or 0)<0 else C["credit_fg"],
                  "777777"]
        for ci,(v,a,f,fg) in enumerate(zip(vals,aligns,fmts,fgs),1):
            wc(ws,ri,ci,v,align=a,fmt=f,fg=fg,bg=bg,bdr=thin("bottom"))
        ws.row_dimensions[ri].height=15
    T = D + len(txs)
    ws.merge_cells(f"A{T}:C{T}")
    c=ws.cell(row=T,column=1,value="TOTALS")
    c.font=Font(name="Arial",bold=True,size=9,color="FFFFFF")
    c.fill=PatternFill("solid",start_color=AI_COLOR)
    c.alignment=Alignment(horizontal="right",vertical="center",indent=1)
    c.border=thin()
    for ci,col,label in [(4,"D","debit"),(5,"E","credit")]:
        formula = f"=SUM({col}{D}:{col}{T-1})"
        c=ws.cell(row=T,column=ci,value=formula)
        c.font=Font(name="Arial",bold=True,size=9,color="FFFFFF")
        c.fill=PatternFill("solid",start_color=AI_COLOR)
        c.number_format=MFMT
        c.alignment=Alignment(horizontal="right",vertical="center")
        c.border=thin()
    cb=data.get("closing_balance")
    c=ws.cell(row=T,column=6,value=cb)
    c.font=Font(name="Arial",bold=True,size=9,color="FFFFFF")
    c.fill=PatternFill("solid",start_color=AI_COLOR)
    c.number_format=MFMT
    c.alignment=Alignment(horizontal="right",vertical="center")
    c.border=thin()
    ws.cell(row=T,column=7).fill=PatternFill("solid",start_color=AI_COLOR)
    ws.row_dimensions[T].height=18
    ws.freeze_panes=f"A{D}"
    ws.auto_filter.ref=f"A{H}:G{T-1}"
    ws.page_setup.orientation="landscape"
    ws.page_setup.fitToPage=True
    ws.page_setup.fitToWidth=1
    if any(v is not None for v in ageing.values()):
        ws2=wb.create_sheet("Ageing Summary")
        ws2.sheet_properties.tabColor=AI_COLOR
        for col,w in [("A",24),("B",18)]: ws2.column_dimensions[col].width=w
        ws2.merge_cells("A1:B1")
        c=ws2["A1"]
        c.value="AGEING SUMMARY"
        c.font=Font(name="Arial",bold=True,size=12,color="FFFFFF")
        c.fill=PatternFill("solid",start_color=AI_COLOR)
        c.alignment=Alignment(horizontal="center",vertical="center")
        ws2.row_dimensions[1].height=26
        bucket_rows=[
            ("0–30 Days",    ageing.get("0_30")),
            ("31–60 Days",   ageing.get("31_60")),
            ("61–90 Days",   ageing.get("61_90")),
            ("91–120 Days",  ageing.get("91_120")),
            ("Above 120 Days",ageing.get("above_120")),
            ("Total Outstanding", ageing.get("total") or data.get("closing_balance")),
        ]
        for ri,(lbl,val) in enumerate(bucket_rows,start=2):
            bold=lbl.startswith("Total")
            bg2=PatternFill("solid",start_color=C["info_bg"]) if bold else None
            lc=ws2.cell(row=ri,column=1,value=lbl)
            lc.font=Font(name="Arial",bold=bold,size=9)
            if bg2: lc.fill=bg2
            lc.border=thin("bottom")
            vc=ws2.cell(row=ri,column=2,value=val)
            vc.font=Font(name="Arial",bold=bold,size=9,
                         color=C["neg_bal"] if (val or 0)<0 else "000000")
            vc.number_format=MFMT
            vc.alignment=Alignment(horizontal="right",vertical="center")
            if bg2: vc.fill=bg2
            vc.border=thin("bottom")
            ws2.row_dimensions[ri].height=16
    ws3=wb.create_sheet("AI Extraction Info")
    ws3.sheet_properties.tabColor="AB47BC"
    ws3.column_dimensions["A"].width=28
    ws3.column_dimensions["B"].width=50
    ws3.merge_cells("A1:B1")
    c=ws3["A1"]
    c.value="AI EXTRACTION DETAILS"
    c.font=Font(name="Arial",bold=True,size=11,color="FFFFFF")
    c.fill=PatternFill("solid",start_color=AI_COLOR)
    c.alignment=Alignment(horizontal="center",vertical="center")
    ws3.row_dimensions[1].height=24
    meta=[
        ("Model Used","claude-sonnet-4-6"),
        ("Extraction Method","Universal AI (no template required)"),
        ("Transactions Found",str(len(txs))),
        ("Supplier Detected",supplier),
        ("Customer Detected",customer),
        ("Statement Date",stmt_date),
        ("Currency",currency),
        ("",""),
        ("⚠️ Tip","Always verify key figures (totals, closing balance) against the original PDF"),
        ("✅ Accuracy","Monetary values are extracted exactly as shown in the source document"),
        ("🔒 Privacy","PDF text is sent to Claude API for processing — do not upload confidential statements"),
    ]
    for ri,(lbl,val) in enumerate(meta,start=2):
        lc=ws3.cell(row=ri,column=1,value=lbl)
        lc.font=Font(name="Arial",bold=True,size=9)
        lc.fill=PatternFill("solid",start_color=C["info_bg"])
        lc.alignment=Alignment(horizontal="left",vertical="center",indent=1)
        lc.border=thin("bottom")
        vc=ws3.cell(row=ri,column=2,value=val)
        vc.font=Font(name="Arial",size=9)
        vc.alignment=Alignment(horizontal="left",vertical="center",indent=1)
        vc.border=thin("bottom")
        ws3.row_dimensions[ri].height=16
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf
def convert_ai(pdf_bytes, filename, api_key):
    lines=[]
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        pages=len(pdf.pages)
        for page in pdf.pages:
            t=page.extract_text()
            if t: lines.extend(t.split("\n"))
    raw_text="\n".join(lines)
    data=_llm_extract(raw_text, api_key)
    txs=data.get("transactions") or []
    buf=excel_universal(data, filename)
    return "ai", len(txs), 0, pages, buf, data.get("supplier","Unknown")
# ══════════════════════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
.ai-banner{
  background:linear-gradient(120deg,#6a1b9a,#4a148c,#1a237e);
  background-size:200% 200%;animation:gradientshift 6s ease infinite;
  border-radius:18px;padding:20px 24px;color:white;margin:14px 0;
  box-shadow:0 10px 28px -8px rgba(106,27,154,.55);
}
.ai-banner h3{margin:0 0 6px;font-size:1.2rem;font-weight:800}
.ai-banner p{margin:0;font-size:.9rem;opacity:.9}
.ai-ok{
  background:linear-gradient(120deg,#ede7f6,#e8eaf6);
  border:1.5px solid #9c27b0;border-radius:16px;
  padding:18px 22px;color:#4a148c;font-weight:600;
  margin-top:14px;box-shadow:0 6px 18px -8px rgba(156,39,176,.3);
}
.key-box{
  background:#fafafa;border:1.5px solid #e0e0e0;border-radius:12px;
  padding:14px 16px;margin:10px 0;font-size:.85rem;color:#444;
}
.key-ok  {border-color:#43a047;background:#f1f8e9;color:#2e7d32}
.key-warn{border-color:#fb8c00;background:#fff8e1;color:#e65100}
</style>
""", unsafe_allow_html=True)
with st.sidebar:
    st.markdown("## ⚙️ Settings")
    st.markdown("---")
    st.markdown("### 🤖 AI Mode — API Key")
    st.markdown(
        "AI Mode uses the **Claude API**. "
        "The key is resolved in this order:\n"
        "1. Streamlit Cloud **Secrets** (`ANTHROPIC_API_KEY`)\n"
        "2. Environment variable `ANTHROPIC_API_KEY`\n"
        "3. Entered below (session only — not stored)"
    )
    _auto_key = ""
    try:
        _auto_key = st.secrets.get("ANTHROPIC_API_KEY","") or ""
    except Exception:
        pass
    if not _auto_key:
        _auto_key = os.environ.get("ANTHROPIC_API_KEY","")
    if _auto_key:
        masked = _auto_key[:8] + "••••••••" + _auto_key[-4:]
        st.markdown(f'<div class="key-box key-ok">✅ API key loaded automatically<br><code>{masked}</code></div>',
                    unsafe_allow_html=True)
    else:
        st.markdown('<div class="key-box key-warn">⚠️ No key found in secrets/env — enter below for this session</div>',
                    unsafe_allow_html=True)
        manual_key = st.text_input(
            "Anthropic API Key",
            type="password",
            placeholder="sk-ant-api03-...",
            help="Get your key at console.anthropic.com — it is used only for this session and never stored.",
            key="manual_api_key_input",
        )
        if manual_key:
            st.session_state["_anthropic_key"] = manual_key
            st.success("Key saved for this session ✅")
    st.markdown("---")
    st.markdown("### 📖 How to add a permanent key")
    with st.expander("Streamlit Cloud (recommended)"):
        st.markdown('''
1. Open your app in Streamlit Cloud
2. Go to **Settings → Secrets**
3. Add:
```toml
ANTHROPIC_API_KEY = "sk-ant-api03-..."
```
4. Save — your app will restart with the key available
        ''')
    with st.expander("Local / Anaconda"):
        st.markdown('''
Set the environment variable before running:
```bash
# Windows (Anaconda Prompt)
set ANTHROPIC_API_KEY=sk-ant-api03-...
streamlit run app.py
# Mac / Linux
export ANTHROPIC_API_KEY=sk-ant-api03-...
streamlit run app.py
```
        ''')
    with st.expander("Get an API key"):
        st.markdown('''
1. Go to **console.anthropic.com**
2. Sign in / create an account
3. Navigate to **API Keys**
4. Click **Create Key** — copy and save it securely
        ''')
    st.markdown("---")
    st.caption("Statement PDF → Excel · Files never stored")
mode = st.radio(
    "Conversion mode:",
    ["🔵  Rule-Based  (fast · free · known suppliers)",
     "🤖  AI Mode  (any supplier · requires API key)"],
    horizontal=True,
)
use_ai = mode.startswith("🤖")
if use_ai:
    api_key = _get_api_key()
    if not api_key:
        st.markdown("""
<div class="ai-banner">
  <h3>🔑 API Key Required</h3>
  <p>AI Mode needs an Anthropic API key. Enter it in the <b>sidebar → AI Mode — API Key</b>,
  or add it to Streamlit Secrets / your environment variable.<br>
  No key? Get one free at <b>console.anthropic.com</b></p>
</div>""", unsafe_allow_html=True)
    else:
        st.markdown("""
<div class="ai-banner">
  <h3>🤖 AI Universal Parser — powered by Claude</h3>
  <p>Upload <b>any</b> supplier statement PDF. Claude reads the document, understands its
  structure, and extracts all transactions into a formatted Excel — no template required.</p>
  <p style="margin-top:8px;font-size:.8rem;opacity:.75">
  ⚠️ PDF text is sent to the Claude API &nbsp;·&nbsp;
  Do not upload confidential documents &nbsp;·&nbsp;
  Always verify totals against the original PDF</p>
</div>""", unsafe_allow_html=True)
else:
    api_key = None
uploaded = st.file_uploader(
    "Drop your PDF here",
    type=["pdf","PDF"],
    label_visibility="collapsed",
    key="pdf_upload",
)
if uploaded:
    st.markdown(f"📎 **Uploaded:** `{uploaded.name}` &nbsp;({uploaded.size/1024:.1f} KB)")
    pdf_bytes = uploaded.read()
    if use_ai:
        api_key = _get_api_key()
        if not api_key:
            st.error("❌ No API key found. Enter your Anthropic API key in the sidebar first.")
            st.stop()
        with st.spinner("🤖 Claude is reading your statement… (10–20 seconds)"):
            try:
                fmt,tx_count,_,pages,buf,supplier = convert_ai(pdf_bytes, uploaded.name, api_key)
            except ValueError as e:
                st.error(f"❌ {e}")
                st.stop()
            except json.JSONDecodeError as e:
                st.error(f"❌ AI returned unexpected output: {e}. Try again.")
                st.stop()
            except Exception as e:
                st.error(f"❌ Unexpected error: {e}")
                st.stop()
        st.markdown(f"""<div class="ai-ok">
        🤖 &nbsp;<b>AI extracted successfully!</b> &nbsp;·&nbsp;
        📄 {pages} page(s) &nbsp;·&nbsp;
        🧾 {tx_count} transactions &nbsp;·&nbsp;
        🏢 {supplier}
        </div>""", unsafe_allow_html=True)
        st.caption("✅ Check the **AI Extraction Info** sheet in the Excel file for verification tips.")
        st.write("")
        st.download_button(
            "⬇️  Download AI-Extracted Excel",
            data=buf,
            file_name=Path(uploaded.name).stem+"_AI_extracted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        with st.spinner("✨ Converting your statement…"):
            try:
                fmt,tx_count,chq_count,pages,buf = convert(pdf_bytes)
            except Exception as e:
                st.error(f"❌ Error: {e}"); st.stop()
        if fmt == "unknown":
            st.warning("⚠️ Format not recognised by rule-based parser.")
            st.markdown("""
<div class="ai-banner" style="margin-top:12px">
  <h3>💡 Try AI Mode</h3>
  <p>Select <b>🤖 AI Mode</b> in the mode selector above and re-upload your PDF.
  Claude will automatically read and extract any statement format — no setup needed.</p>
</div>""", unsafe_allow_html=True)
        else:
            labels = {
                "savannah":"Savannah Cement","national":"National Cement",
                "mombasa":"Mombasa Cement","karsan":"Karsan Ramji (Ndovu)",
                "skylife":"Skylife Kenya","springtech":"Springtech (K) Ltd",
                "scania":"Scania East Africa","transsolutions":"Trans-Solutions (E.A)",
                "autoxpress":"AutoXpress Limited","gathani":"Gathani Limited",
                "anco":"Anco Trading Ltd",
                "coa":"Confirmation of Accounts",
                "bamburi":"Bamburi Cement PLC",
                "crown":"Crown Paints Kenya PLC",
                "mapei":"Mapei East Africa Ltd",
                "basco":"Basco Products (K) Ltd",
                "dolphin":"Dolphin Ricambi Ltd",
            }
            label   = labels.get(fmt, fmt)
            chq_txt = f" · {chq_count} cheques" if chq_count else ""
            st.markdown(
                f"""<div class="ok">🎉 &nbsp;<b>{label}</b> converted successfully! &nbsp;·&nbsp;
                📄 {pages} page(s) &nbsp;·&nbsp; 🧾 {tx_count} transactions{chq_txt}</div>""",
                unsafe_allow_html=True,
            )
            st.write("")
            st.download_button(
                "⬇️  Download Excel File",
                data=buf,
                file_name=Path(uploaded.name).stem+"_converted.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
st.markdown("---")
st.markdown(
    '<div class="footer-note">✨ Files processed in memory · never stored · free to use ✨</div>',
    unsafe_allow_html=True,
)
