
# app.py
import re
from difflib import SequenceMatcher
from io import BytesIO

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Product Price Comparison", layout="wide", page_icon="🚀")

st.markdown(
    """
    <div style='background-color:#002e5b;padding:12px 16px;border-radius:12px;'>
      <h1 style='text-align:center;color:#fff;margin:0;'>🚀 Product Price Comparison Dashboard</h1>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    "<p style='text-align:center;font-size:16px;color:#333;margin-top:8px;'>"
    "Upload Marlin, Website, and Master price files to generate a comparison report."
    "</p>",
    unsafe_allow_html=True,
)

def make_template_bytes() -> bytes:
    df = pd.DataFrame({"Variant Code": [], "Variant Price": []})
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Prices", index=False)
    return bio.getvalue()

CODE_CANDIDATES = ["variantcode","sku","code","itemcode","productcode"]
PRICE_CANDIDATES = ["price","rrp","retail","sellprice"]
INC_HINTS = ["incgst","inclgst","inc_gst"]
EXC_HINTS = ["exgst","exclgst","exc_gst"]

def norm(s):
    return re.sub(r"[^a-z0-9]+", "", str(s).lower())

def best_match_column(columns, candidates):
    scores = {}
    for col in columns:
        n = norm(col)
        score = sum(1 for c in candidates if c in n)
        score += max(SequenceMatcher(None, n, c).ratio() for c in candidates)
        scores[col] = score
    best = max(scores, key=scores.get)
    return best

def pick_sheet_and_columns(xls):
    book = pd.read_excel(xls, sheet_name=None, dtype=str)
    for sheet, df in book.items():
        if df is None or df.empty:
            continue
        df = df.loc[:, ~df.columns.astype(str).str.contains("^Unnamed")]
        code = best_match_column(df.columns, CODE_CANDIDATES)
        price = best_match_column(df.columns, PRICE_CANDIDATES)
        sub = df[[code, price]].copy()
        sub.columns = ["Variant Code", "Variant Price"]
        return sub, {"sheet": sheet, "code_col": code, "price_col": price}
    raise ValueError("No valid sheet found")

def coerce_price(x):
    try:
        return float(re.sub(r"[^\d.-]", "", str(x)))
    except:
        return pd.NA

def clean(df):
    df["Variant Price"] = df["Variant Price"].apply(coerce_price)
    return df.drop_duplicates("Variant Code")

def make_report(marlin, website, master):
    merged = master.merge(website, on="Variant Code", how="outer", suffixes=("_Master","_Website"))
    merged = merged.merge(marlin, on="Variant Code", how="outer")
    merged.rename(columns={"Variant Price":"Variant Price_Marlin"}, inplace=True)

    def cmp(a,b):
        if pd.isna(a) or pd.isna(b): return "N/A"
        return "Match" if abs(a-b)<=0.01 else "Mismatch"

    merged["Website vs Master"] = merged.apply(lambda r: cmp(r["Variant Price_Website"], r["Variant Price_Master"]), axis=1)
    merged["Marlin vs Master"] = merged.apply(lambda r: cmp(r["Variant Price_Marlin"], r["Variant Price_Master"]), axis=1)

    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        merged.to_excel(w, index=False)
    return out.getvalue()

st.subheader("Upload files")
c1,c2,c3 = st.columns(3)
with c1: marlin_file = st.file_uploader("Marlin", type=["xlsx"])
with c2: website_file = st.file_uploader("Website", type=["xlsx"])
with c3: master_file = st.file_uploader("Master", type=["xlsx"])

if st.button("Run Comparison"):
    if not all([marlin_file, website_file, master_file]):
        st.error("Upload all three files")
    else:
        m,_ = pick_sheet_and_columns(marlin_file)
        w,_ = pick_sheet_and_columns(website_file)
        master,_ = pick_sheet_and_columns(master_file)

        report = make_report(clean(m), clean(w), clean(master))
        st.download_button("Download Report", report, "comparison.xlsx")
