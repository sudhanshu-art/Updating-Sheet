#!/usr/bin/env python3
"""
Shipment Bill Generator  (macOS)
--------------------------------
Reads a ParcelX shipment export and produces a formatted Excel bill:
  - Invoice  : flat RATE per AWB (cancelled shipments excluded)
  - AWB Details : one row per shipment
  - COD by Email : COD collected per mail id, counted only on DELIVERED shipments

USAGE
  1. Put your shipment .xlsx in ~/Downloads (e.g. Sidhhart.xlsx)
  2. In Terminal:   python3 generate_bill.py
     (optionally:   python3 generate_bill.py "/path/to/file.xlsx")
  3. The bill is written next to the source as <name>_Bill.xlsx and opened.

It auto-installs pandas/openpyxl on first run if missing.
Open the result in Excel/Numbers — formulas recalculate automatically on open.
"""

import os, sys, re, glob, subprocess
from datetime import datetime

# ----------------------- CONFIG -----------------------
RATE          = 66                 # Rs charged per AWB
FILENAME_HINT = "sidhhart"         # substring to find the file in Downloads (case-insensitive)
DOWNLOADS     = os.path.expanduser("~/Downloads")
# Source column names (edit here if your export uses different headers)
COL_AWB     = "Parcelx Tracking ID"
COL_STATUS  = "Current User Status"
COL_DATE    = "Placed Date (D-M-Y)"
COL_EMAIL   = "User Email"
COL_COD     = "COD Amount"
COL_CUST    = "Drop Customer Name"
COL_CITY    = "Drop City"
COL_STATE   = "Drop State"
COL_COURIER = "Courier Used"
DELIVERED_LABEL = "Delivered"
# ------------------------------------------------------

def ensure_deps():
    try:
        import pandas, openpyxl  # noqa
    except ImportError:
        print("Installing required packages (pandas, openpyxl)...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--user", "pandas", "openpyxl"])

def find_source():
    if len(sys.argv) > 1 and os.path.isfile(sys.argv[1]):
        return sys.argv[1]
    cands = [f for f in glob.glob(os.path.join(DOWNLOADS, "*.xls*"))
             if FILENAME_HINT in os.path.basename(f).lower()]
    if not cands:  # fall back to newest spreadsheet in Downloads
        cands = glob.glob(os.path.join(DOWNLOADS, "*.xls*"))
    if not cands:
        sys.exit(f"No .xlsx found in {DOWNLOADS}. Pass a path: python3 generate_bill.py /path/file.xlsx")
    return max(cands, key=os.path.getmtime)  # most recently modified

def build(src):
    import pandas as pd
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    df = pd.read_excel(src, dtype=str)
    df["_status"] = df[COL_STATUS].fillna("").str.strip()
    before = len(df)
    df = df[~df["_status"].str.contains("cancel", case=False, na=False)].copy()  # drop cancelled
    excluded = before - len(df)
    df["_date"] = pd.to_datetime(df[COL_DATE], errors="coerce").dt.strftime("%d-%m-%Y")
    df["_cod"]  = pd.to_numeric(df[COL_COD], errors="coerce").fillna(0)

    ERR = re.compile(r"#(REF!|DIV/0!|VALUE!|NAME\?|N/A|NULL!|NUM!|GETTING_DATA)")
    def clean(x):
        s = "" if x is None else str(x)
        s = ERR.sub("", s).strip()
        return (" " + s) if s[:1] in ("=", "+", "-", "@") else s

    data = pd.DataFrame({
        "awb":     df[COL_AWB].fillna("").map(clean),
        "date":    df["_date"].fillna(""),
        "status":  df["_status"],
        "email":   df[COL_EMAIL].fillna(""),
        "cust":    df[COL_CUST].fillna("").map(clean),
        "city":    df[COL_CITY].fillna("").map(clean),
        "state":   df[COL_STATE].fillna("").map(clean),
        "courier": df[COL_COURIER].fillna("").map(clean),
        "cod":     df["_cod"].astype(float),
    })
    n = len(data)
    emails = sorted(df[COL_EMAIL].fillna("(blank)").unique())

    FONT="Arial"; navy="1F3864"; blue="2E5496"; grey="F2F2F2"; green="375623"
    thin=Side(style="thin",color="BFBFBF"); border=Border(left=thin,right=thin,top=thin,bottom=thin)
    def f(**k): return Font(name=FONT, **k)
    hf=PatternFill("solid",fgColor=navy)
    wb=Workbook()

    # Invoice
    inv=wb.active; inv.title="Invoice"; inv.sheet_view.showGridLines=False
    inv["B2"]="SHIPMENT BILL"; inv["B2"].font=f(size=22,bold=True,color=navy)
    inv["B3"]="AWB Charges Statement"; inv["B3"].font=f(size=11,italic=True,color="808080")
    meta=[("Bill Date",datetime.now().strftime("%d-%m-%Y")),
          ("Source File",os.path.basename(src)),
          ("Billing Basis",f"Flat Rs {RATE} per AWB (one shipment = one charge)"),
          ("Cancelled shipments",f"Excluded ({excluded} found)")]
    r=5
    for k,v in meta:
        inv[f"B{r}"]=k; inv[f"B{r}"].font=f(bold=True,color="404040")
        inv[f"C{r}"]=v; inv[f"C{r}"].font=f(color="404040"); r+=1
    r+=1; hdr=r
    for c in ("B","C"):
        cc=inv[f"{c}{hdr}"]; cc.value="Description" if c=="B" else "Value"
        cc.font=f(bold=True,color="FFFFFF",size=11); cc.fill=hf; cc.border=border
    inv.row_dimensions[hdr].height=20
    rate_row=hdr+1; qty_row=hdr+2
    rows=[("Rate per AWB (Rs)",RATE,"rate"),
          ("Billable Shipments (AWBs)",f"=COUNTA('AWB Details'!B2:B{n+1})","num"),
          ("TOTAL PAYABLE (Rs)",f"=C{rate_row}*C{qty_row}","total")]
    for i,(lab,val,kind) in enumerate(rows):
        rr=hdr+1+i; lc=inv[f"B{rr}"]; vc=inv[f"C{rr}"]
        lc.value=lab; vc.value=val; lc.border=border; vc.border=border
        vc.alignment=Alignment(horizontal="right")
        if kind=="total":
            lc.font=vc.font=f(bold=True,size=12,color="FFFFFF")
            lc.fill=vc.fill=PatternFill("solid",fgColor=navy); vc.number_format='"Rs "#,##0'
        elif kind=="rate":
            lc.font=f(color="404040"); vc.font=f(bold=True,color="0000FF")
            lc.fill=vc.fill=PatternFill("solid",fgColor=grey); vc.number_format='"Rs "#,##0'
        else:
            lc.font=f(color="404040"); vc.font=f(color="000000")
            lc.fill=vc.fill=PatternFill("solid",fgColor=grey); vc.number_format="#,##0"
    inv.column_dimensions["A"].width=3; inv.column_dimensions["B"].width=34; inv.column_dimensions["C"].width=24

    # AWB Details
    det=wb.create_sheet("AWB Details"); det.sheet_view.showGridLines=False
    H=["S.No","AWB (Tracking ID)","Placed Date","Status","Email","Drop Customer",
       "Drop City","Drop State","Courier","COD Amount","Rate (Rs)","Amount (Rs)"]
    det.append(H)
    for ci,_ in enumerate(H,1):
        c=det.cell(row=1,column=ci); c.font=f(bold=True,color="FFFFFF",size=10)
        c.fill=hf; c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); c.border=border
    det.row_dimensions[1].height=26
    rref=f"'Invoice'!$C${rate_row}"
    for i,row in enumerate(data.itertuples(index=False),start=2):
        a,d,s,e,cu,ci_,st,co,cod=row
        for col,v in enumerate([i-1,str(a),d,s,e,cu,ci_,st,co,float(cod),f"={rref}",f"={rref}"],1):
            det.cell(row=i,column=col,value=v)
    last=n+1; totr=last+1
    det.cell(row=totr,column=9,value="TOTAL")
    det.cell(row=totr,column=10,value=f"=SUM(J2:J{last})")
    det.cell(row=totr,column=11,value=f"=SUM(K2:K{last})")
    det.cell(row=totr,column=12,value=f"=SUM(L2:L{last})")
    for i in range(2,last+1):
        for ci in range(1,13):
            c=det.cell(row=i,column=ci); c.font=f(size=9,color="333333"); c.border=border
            if i%2==0: c.fill=PatternFill("solid",fgColor=grey)
        det.cell(row=i,column=1).alignment=Alignment(horizontal="center")
        for ci in (10,11,12): det.cell(row=i,column=ci).number_format='"Rs "#,##0'
    for ci in (9,10,11,12):
        c=det.cell(row=totr,column=ci); c.font=f(bold=True,color="FFFFFF")
        c.fill=PatternFill("solid",fgColor=blue); c.border=border
        if ci>=10: c.number_format='"Rs "#,##0'; c.alignment=Alignment(horizontal="right")
    for ci,w in enumerate([7,22,13,12,30,24,15,15,18,13,11,13],1):
        det.column_dimensions[get_column_letter(ci)].width=w
    det.freeze_panes="A2"; det.auto_filter.ref=f"A1:L{last}"

    # COD by Email
    cod=wb.create_sheet("COD by Email"); cod.sheet_view.showGridLines=False
    cod["B2"]="COD COLLECTED BY MAIL ID"; cod["B2"].font=f(size=16,bold=True,color=green)
    cod["B3"]="COD counted only on Delivered shipments"; cod["B3"].font=f(size=10,italic=True,color="808080")
    ch=5
    for ci,h in enumerate(["Mail ID","Total AWBs","AWB Charges (Rs)","Delivered Shipments","COD Collected (Rs)"],2):
        c=cod.cell(row=ch,column=ci,value=h); c.font=f(bold=True,color="FFFFFF",size=10)
        c.fill=hf; c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); c.border=border
    cod.row_dimensions[ch].height=28
    em_r=f"'AWB Details'!$E$2:$E${last}"; st_r=f"'AWB Details'!$D$2:$D${last}"
    cod_r=f"'AWB Details'!$J$2:$J${last}"; amt_r=f"'AWB Details'!$L$2:$L${last}"
    rr=ch+1
    for em in emails:
        cod.cell(row=rr,column=2,value=em)
        cod.cell(row=rr,column=3,value=f'=COUNTIF({em_r},$B{rr})')
        cod.cell(row=rr,column=4,value=f'=SUMIF({em_r},$B{rr},{amt_r})')
        cod.cell(row=rr,column=5,value=f'=COUNTIFS({em_r},$B{rr},{st_r},"{DELIVERED_LABEL}")')
        cod.cell(row=rr,column=6,value=f'=SUMIFS({cod_r},{em_r},$B{rr},{st_r},"{DELIVERED_LABEL}")')
        for ci in range(2,7):
            c=cod.cell(row=rr,column=ci); c.font=f(size=10,color="333333"); c.border=border
            if rr%2==1: c.fill=PatternFill("solid",fgColor=grey)
        cod.cell(row=rr,column=3).number_format="#,##0"; cod.cell(row=rr,column=4).number_format='"Rs "#,##0'
        cod.cell(row=rr,column=5).number_format="#,##0"; cod.cell(row=rr,column=6).number_format='"Rs "#,##0'
        rr+=1
    tr=rr
    cod.cell(row=tr,column=2,value="TOTAL")
    for ci,col in zip(range(3,7),("C","D","E","F")):
        cod.cell(row=tr,column=ci,value=f"=SUM({col}{ch+1}:{col}{tr-1})")
    for ci in range(2,7):
        c=cod.cell(row=tr,column=ci); c.font=f(bold=True,color="FFFFFF"); c.fill=PatternFill("solid",fgColor=green); c.border=border
    cod.cell(row=tr,column=3).number_format="#,##0"; cod.cell(row=tr,column=4).number_format='"Rs "#,##0'
    cod.cell(row=tr,column=5).number_format="#,##0"; cod.cell(row=tr,column=6).number_format='"Rs "#,##0'
    for ci,w in enumerate([3,34,14,18,20,20],1):
        cod.column_dimensions[get_column_letter(ci)].width=w

    out=os.path.join(os.path.dirname(src), os.path.splitext(os.path.basename(src))[0]+"_Bill.xlsx")
    wb.save(out)
    return out, n, excluded

def main():
    ensure_deps()
    src=find_source()
    print("Reading:", src)
    out,n,exc=build(src)
    print(f"Billable AWBs : {n:,}  (cancelled excluded: {exc})")
    print(f"Total payable : Rs {n*RATE:,}")
    print("Saved bill    :", out)
    try: subprocess.run(["open", out])   # macOS: open in Excel/Numbers
    except Exception: pass

if __name__ == "__main__":
    main()
