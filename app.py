# app.py
import re
from difflib import SequenceMatcher
from io import BytesIO

import pandas as pd
import streamlit as st

# --------------------------
# Page configuration & Header
# --------------------------
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
    "Upload a Marlin price file, a Website price file, and a Master (source of truth) file — "
    "or download our templates below — to generate a detailed comparison report."
    "</p>",
    unsafe_allow_html=True,
)

# --------------------------
# Sidebar settings
# --------------------------
st.sidebar.header("Comparison Settings")
COMPARE_DECIMALS = st.sidebar.selectbox(
    "Round prices to decimals before comparing",
    options=[0, 1, 2],
    index=0,  # default: whole dollars
    help="Fixes false mismatches when Master has decimals but Website/Marlin are rounded.",
)
TOLERANCE = st.sidebar.number_input(
    "Tolerance after rounding",
    min_value=0.0,
    value=0.0,
    step=0.01,
    help="Optional extra tolerance after rounding (usually 0.00 is best).",
)

# ---------------------------------
# Helpers: Template + Column Finding
# ---------------------------------
def make_template_bytes(template_name: str) -> bytes:
    """Create a 2-column Excel file: Variant Code, Variant Price."""
    df = pd.DataFrame({"Variant Code": [], "Variant Price": []})
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Prices", index=False)
        # Add a small 'README' sheet with directions
        pd.DataFrame(
            {
                "Instructions": [
                    "Fill only these two columns.",
                    "Variant Code should be the unique SKU/variant identifier.",
                    "Variant Price should be a number (Inc GST preferred).",
                ]
            }
        ).to_excel(writer, sheet_name="README", index=False)
    return bio.getvalue()

# Candidate text for variant code columns (normalized)
CODE_CANDIDATES = [
    "variantcode",
    "variant_code",
    "variant sku",
    "variantsku",
    "sku",
    "productcode",
    "product_code",
    "itemcode",
    "item_code",
    "code",
    "partnumber",
    "partno",
]

# Candidate text for price columns (normalized)
PRICE_CANDIDATES_PRIMARY = [
    "variantprice",
    "price",
    "webprice",
    "websiteprice",
    "retail",
    "rrp",
    "sellprice",
    "sellingprice",
    "listprice",
]

# Hints to break ties toward INC over EXC GST when both exist
INC_HINTS = ["incgst", "inclgst", "incl_gst", "inc_gst"]
EXC_HINTS = ["exgst", "exc_gst", "exclgst", "excl_gst"]


def norm(s: str) -> str:
    s = str(s).strip().lower()
    s = re.sub(r"[^a-z0-9]+", "", s)  # keep alnum only
    return s


def best_match_column(columns, candidates, extra_bias_inc=None, extra_bias_exc=None):
    """
    Pick the best matching column from a list of names using:
    1) direct containment against candidate tokens
    2) simple fuzzy score
    Bias toward INC GST columns when both INC/EXC are present.
    """
    if not columns:
        return None

    normalized = {c: norm(c) for c in columns}

    scores = {c: 0.0 for c in columns}
    for col, ncol in normalized.items():
        # base: candidate containment
        for cand in candidates:
            if norm(cand) in ncol:
                scores[col] += 1.0

        # fuzzy fallback vs each candidate
        fuzz = max(SequenceMatcher(None, ncol, norm(cand)).ratio() for cand in candidates)
        scores[col] += 0.4 * fuzz  # small fuzzy contribution

        # bias toward INC or EXC if present
        if extra_bias_inc and any(h in ncol for h in extra_bias_inc):
            scores[col] += 0.25
        if extra_bias_exc and any(h in ncol for h in extra_bias_exc):
            scores[col] += 0.10

    best = max(scores, key=lambda c: scores[c])
    return best if scores[best] >= 0.6 else None


def pick_sheet_and_columns(xls_file):
    """
    Read ALL sheets, find the sheet that most confidently contains one
    code column and one price column. Return (df_subset, meta).
    """
    book = pd.read_excel(xls_file, sheet_name=None, dtype=str)

    best = None
    best_score = -1

    for sheet_name, df in book.items():
        if df is None or df.empty:
            continue

        # drop empty columns that pandas might create
        df = df.loc[:, ~df.columns.astype(str).str.contains("^Unnamed", na=False)]

        code_col = best_match_column(list(df.columns), CODE_CANDIDATES)
        price_col = best_match_column(
            list(df.columns),
            PRICE_CANDIDATES_PRIMARY,
            extra_bias_inc=INC_HINTS,
            extra_bias_exc=EXC_HINTS,
        )

        score = (1 if code_col else 0) + (1 if price_col else 0)
        if code_col and price_col:
            score += 0.5

        if score > best_score:
            best = (df, code_col, price_col, sheet_name)
            best_score = score

    if not best or best_score <= 0:
        raise ValueError("Could not detect suitable columns in any sheet.")

    df, code_col, price_col, sheet_name = best
    if not code_col or not price_col:
        raise ValueError(
            f"Auto-detection incomplete in sheet '{sheet_name}'. "
            f"Found code column: {code_col}, price column: {price_col}"
        )

    sub = df[[code_col, price_col]].copy()
    sub.rename(columns={code_col: "Variant Code", price_col: "Variant Price"}, inplace=True)

    return sub, {"sheet": sheet_name, "code_col": code_col, "price_col": price_col}


def coerce_price(s):
    if pd.isna(s):
        return pd.NA
    if isinstance(s, (int, float)):
        return float(s)

    txt = str(s)
    neg = False
    if "(" in txt and ")" in txt:
        neg = True

    # strip currency and commas etc
    txt = re.sub(r"[^\d\.\-]", "", txt)  # keep digits, dot, minus

    # handle weird strings like "1.234.56" -> "1234.56"
    if txt.count(".") > 1:
        left, right = txt.rsplit(".", 1)
        txt = re.sub(r"\.", "", left) + "." + right

    try:
        val = float(txt)
        return -val if neg else val
    except Exception:
        return pd.NA


def clean_prices(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["Variant Code"] = out["Variant Code"].astype(str).str.strip()
    out["Variant Price"] = out["Variant Price"].apply(coerce_price)
    out = out.dropna(subset=["Variant Code"]).drop_duplicates(subset=["Variant Code"], keep="last")
    return out


def make_report(
    marlin_df: pd.DataFrame,
    website_df: pd.DataFrame,
    master_df: pd.DataFrame,
    meta_m,
    meta_w,
    meta_master,
    decimals: int = 0,
    tolerance: float = 0.0,
) -> bytes:
    # Merge all sources (keep everything)
    merged = master_df.merge(website_df, on="Variant Code", how="outer", suffixes=("_Master", "_Website"))
    merged = merged.merge(marlin_df, on="Variant Code", how="outer")
    merged.rename(columns={"Variant Price": "Variant Price_Marlin"}, inplace=True)

    def match_rounded(source_price, master_price):
        """
        Compare after rounding BOTH sides to the same decimals.
        This fixes false mismatches when master has more precision than sources.
        """
        if pd.isna(source_price) or pd.isna(master_price):
            return "N/A"
        a = round(float(source_price), decimals)
        b = round(float(master_price), decimals)
        return "Match" if abs(a - b) <= tolerance else "Mismatch"

    # Status columns
    merged["Website vs Master"] = merged.apply(
        lambda r: match_rounded(r.get("Variant Price_Website"), r.get("Variant Price_Master")),
        axis=1,
    )
    merged["Marlin vs Master"] = merged.apply(
        lambda r: match_rounded(r.get("Variant Price_Marlin"), r.get("Variant Price_Master")),
        axis=1,
    )

    # Sentence columns (as you requested)
    def website_sentence(row):
        if row["Website vs Master"] == "Match":
            return "Website Price matches with Master File."
        if row["Website vs Master"] == "Mismatch":
            return "Website Price does NOT match Master File."
        return "Website Price N/A (missing Website or Master)."

    def marlin_sentence(row):
        if row["Marlin vs Master"] == "Match":
            return "Marlin Price matches with Master File."
        if row["Marlin vs Master"] == "Mismatch":
            return "Marlin Price does NOT match Master File."
        return "Marlin Price N/A (missing Marlin or Master)."

    merged["Website Result"] = merged.apply(website_sentence, axis=1)
    merged["Marlin Result"] = merged.apply(marlin_sentence, axis=1)

    # Helpful diffs (raw, not rounded, so you can see the real underlying delta)
    merged["Website - Master Diff"] = merged["Variant Price_Website"] - merged["Variant Price_Master"]
    merged["Marlin - Master Diff"] = merged["Variant Price_Marlin"] - merged["Variant Price_Master"]

    # Which ones are not matching (single label)
    def mismatch_summary(row):
        problems = []
        if row["Website vs Master"] == "Mismatch":
            problems.append("Website")
        if row["Marlin vs Master"] == "Mismatch":
            problems.append("Marlin")
        if problems:
            return "Not matching Master: " + " & ".join(problems)

        if row["Website vs Master"] == "N/A" and row["Marlin vs Master"] == "N/A":
            return "N/A (insufficient data)"
        return "All matching (where comparable)"

    merged["Mismatch Summary"] = merged.apply(mismatch_summary, axis=1)

    # New: overall matched / mismatched groups
    # "All Matched" means: any comparison that exists must be Match, and at least one comparison exists.
    comparable = (merged["Website vs Master"] != "N/A") | (merged["Marlin vs Master"] != "N/A")
    all_matched = comparable & (merged["Website vs Master"].isin(["Match", "N/A"])) & (merged["Marlin vs Master"].isin(["Match", "N/A"])) \
                  & ((merged["Website vs Master"] == "Match") | (merged["Marlin vs Master"] == "Match"))
    any_mismatched = (merged["Website vs Master"] == "Mismatch") | (merged["Marlin vs Master"] == "Mismatch")

    # Export
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        col_order = [
            "Variant Code",
            "Variant Price_Master",
            "Variant Price_Website",
            "Variant Price_Marlin",
            "Website vs Master",
            "Marlin vs Master",
            "Website Result",
            "Marlin Result",
            "Mismatch Summary",
            "Website - Master Diff",
            "Marlin - Master Diff",
        ]

        merged[col_order].to_excel(writer, sheet_name="Full Data", index=False)

        # Requested tabs:
        merged[all_matched][col_order].to_excel(writer, sheet_name="All Matched", index=False)
        merged[any_mismatched][col_order].to_excel(writer, sheet_name="All Mismatched", index=False)

        # Extra helpful tabs (optional but useful)
        merged[merged["Website vs Master"] == "Mismatch"][col_order].to_excel(
            writer, sheet_name="Website Mismatched", index=False
        )
        merged[merged["Marlin vs Master"] == "Mismatch"][col_order].to_excel(
            writer, sheet_name="Marlin Mismatched", index=False
        )
        merged[(merged["Website vs Master"] == "Mismatch") & (merged["Marlin vs Master"] == "Mismatch")][
            col_order
        ].to_excel(writer, sheet_name="Both Mismatched", index=False)

        # Summary sheet
        summary = pd.DataFrame(
            [
                ["Detected Master sheet", meta_master["sheet"]],
                ["Master code column", meta_master["code_col"]],
                ["Master price column", meta_master["price_col"]],
                ["Detected Website sheet", meta_w["sheet"]],
                ["Website code column", meta_w["code_col"]],
                ["Website price column", meta_w["price_col"]],
                ["Detected Marlin sheet", meta_m["sheet"]],
                ["Marlin code column", meta_m["code_col"]],
                ["Marlin price column", meta_m["price_col"]],
                ["Compare decimals", decimals],
                ["Tolerance after rounding", tolerance],
                ["Total Master rows", len(master_df)],
                ["Total Website rows", len(website_df)],
                ["Total Marlin rows", len(marlin_df)],
                ["Website matches (vs Master)", int((merged["Website vs Master"] == "Match").sum())],
                ["Website mismatches (vs Master)", int((merged["Website vs Master"] == "Mismatch").sum())],
                ["Marlin matches (vs Master)", int((merged["Marlin vs Master"] == "Match").sum())],
                ["Marlin mismatches (vs Master)", int((merged["Marlin vs Master"] == "Mismatch").sum())],
                ["All Matched rows", int(all_matched.sum())],
                ["All Mismatched rows", int(any_mismatched.sum())],
            ],
            columns=["Metric", "Value"],
        )
        summary.to_excel(writer, sheet_name="Summary", index=False)

    return output.getvalue()


# ----------------
# Template section
# ----------------
with st.expander("📥 Download blank templates (fill & re-upload)"):
    tcol1, tcol2, tcol3 = st.columns(3)
    with tcol1:
        st.download_button(
            "Download Marlin Template",
            data=make_template_bytes("Marlin"),
            file_name="Marlin_Price_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with tcol2:
        st.download_button(
            "Download Website Template",
            data=make_template_bytes("Website"),
            file_name="Website_Price_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with tcol3:
        st.download_button(
            "Download Master Template",
            data=make_template_bytes("Master"),
            file_name="Master_Price_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

# ---------------------
# File uploads + action
# ---------------------
st.subheader("Upload files")

col1, col2, col3 = st.columns(3)
with col1:
    marlin_file = st.file_uploader("Upload Marlin Price File (.xlsx)", type=["xlsx"], key="marlin")
with col2:
    website_file = st.file_uploader("Upload Website Price File (.xlsx)", type=["xlsx"], key="website")
with col3:
    master_file = st.file_uploader("Upload Master Price File (.xlsx) (Source of Truth)", type=["xlsx"], key="master")

run = st.button("Run Comparison", use_container_width=True)

if run:
    if marlin_file is None or website_file is None or master_file is None:
        st.error("Please upload Marlin, Website, and Master Excel files to proceed.")
    else:
        try:
            with st.spinner("Auto-detecting columns and comparing to Master…"):
                m_raw, meta_m = pick_sheet_and_columns(marlin_file)
                w_raw, meta_w = pick_sheet_and_columns(website_file)
                master_raw, meta_master = pick_sheet_and_columns(master_file)

                m = clean_prices(m_raw)
                w = clean_prices(w_raw)
                master = clean_prices(master_raw)

                report_bytes = make_report(
                    marlin_df=m,
                    website_df=w,
                    master_df=master,
                    meta_m=meta_m,
                    meta_w=meta_w,
                    meta_master=meta_master,
                    decimals=int(COMPARE_DECIMALS),
                    tolerance=float(TOLERANCE),
                )

            st.success("Report ready!")
            st.download_button(
                label="Download Comparison Report",
                data=report_bytes,
                file_name="Price_Comparison_Report_MasterBased.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

            with st.expander("🔎 Auto-detection details"):
                c1, c2, c3 = st.columns(3)
                with c1:
                    st.markdown("**Marlin detection**")
                    st.write(meta_m)
                    st.dataframe(m.head(10), use_container_width=True)
                with c2:
                    st.markdown("**Website detection**")
                    st.write(meta_w)
                    st.dataframe(w.head(10), use_container_width=True)
                with c3:
                    st.markdown("**Master detection**")
                    st.write(meta_master)
                    st.dataframe(master.head(10), use_container_width=True)

        except Exception as e:
            st.error(f"Comparison failed: {e}")
