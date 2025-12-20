import streamlit as st
import pandas as pd
import plotly.express as px
import warnings

warnings.filterwarnings("ignore", message="Workbook contains no default style*")

import io

def download_excel(df, filename, label):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.reset_index().to_excel(writer, index=False)
    st.download_button(
        label=label,
        data=buffer.getvalue(),
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ---------- PAGE CONFIG ----------
st.set_page_config(page_title="Sales Analysis Dashboard", layout="wide")

# ---------- SIDEBAR ----------
with st.sidebar:
    st.title("⚙️ Controls")
    st.markdown("""
        **Steps:**
        1. Flipkart PM Excel upload karein  
        2. Top Products file upload karein  
        3. Generate Analysis button click karein  
    """)

# ---------- TITLE ----------
st.title("📊 Sales Analysis Dashboard")

# ---------- FILE UPLOAD ----------
c1, c2 = st.columns(2)
with c1:
    flipkart_file = st.file_uploader("Upload Flipkart PM Excel", type=["xlsx", "xls"])
with c2:
    top_products_file = st.file_uploader("Upload Top Products (CSV / Excel)", type=["csv", "xlsx", "xls"])

# ---------- BUTTON ----------
generate = st.button("🚀 Generate Analysis", use_container_width=True)

# ---------- MAIN ----------
if generate:
    if not flipkart_file or not top_products_file:
        st.error("❌ Please upload both files")
    else:
        try:
            # ---------- LOAD ----------
            flipkart = pd.read_excel(flipkart_file)
            top = pd.read_csv(top_products_file) if top_products_file.name.endswith(".csv") else pd.read_excel(top_products_file)

            if "Brand" in top.columns:
                top.rename(columns={"Brand": "Brand1"}, inplace=True)

            # ---------- VALIDATION ----------
            required_pm = ["FNS", "Product Name", "Vendor SKU Codes", "Brand", "Brand Manager"]
            for c in required_pm:
                if c not in flipkart.columns:
                    raise KeyError(f"Flipkart PM me '{c}' missing")

            if "Product Id" not in top.columns:
                raise KeyError("Top Products me 'Product Id' missing")
            if "Final Sale Units" not in top.columns:
                raise KeyError("Final Sale Units missing")

            # ---------- CLEAN KEYS ----------
            top["Product Id"] = top["Product Id"].astype(str).str.strip().str.upper()
            flipkart["FNS"] = flipkart["FNS"].astype(str).str.strip().str.upper()

            flipkart = flipkart.drop_duplicates("FNS")

            cp_col = [c for c in flipkart.columns if str(c).lower().startswith("cp")][0]

            # ---------- MERGE ----------
            final_df = top.merge(
                flipkart[
                    ["FNS", "Product Name", "Vendor SKU Codes", "Brand", "Brand Manager", cp_col]
                ],
                left_on="Product Id",
                right_on="FNS",
                how="left"
            ).rename(columns={
                "Brand Manager": "Manager",
                cp_col: "cost"
            })

            # ---------- SAFE TEXT CLEAN ----------
            def clean_text(x):
                try:
                    if pd.isna(x):
                        return "Unknown"
                    return str(x).strip()
                except Exception:
                    return "Unknown"

            final_df["Product Name"] = final_df["Product Name"].apply(clean_text)
            final_df["Vendor SKU Codes"] = final_df["Vendor SKU Codes"].apply(clean_text)

            # ---------- BRAND STANDARDIZATION ----------
            final_df["Brand"] = (
                final_df["Brand"]
                .astype(str)
                .str.strip()
                .str.lower()
                .str.title()
                .replace("Nan", "Unknown")
            )

            # ---------- NUMERIC CLEAN ----------
            final_df["Final Sale Units"] = pd.to_numeric(
                final_df["Final Sale Units"], errors="coerce"
            ).fillna(0).clip(lower=0)

            if "Final Sale Amount" in final_df.columns:
                final_df.rename(columns={"Final Sale Amount": "Sales"}, inplace=True)

            final_df["Sales"] = pd.to_numeric(final_df["Sales"], errors="coerce").fillna(0)
            final_df["cost"] = pd.to_numeric(final_df["cost"], errors="coerce").fillna(0)
            final_df["Manager"] = final_df["Manager"].fillna("Unknown")
            final_df["FNS"] = final_df["FNS"].fillna("Unknown")

            # ---------- METRICS ----------
            st.markdown("---")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Units", int(final_df["Final Sale Units"].sum()))
            m2.metric("Total Sales", f"₹{final_df['Sales'].sum():,.0f}")
            m3.metric("Brands", final_df["Brand"].nunique())
            m4.metric("Managers", final_df["Manager"].nunique())

            # ---------- TABS ----------
            tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
                "📊 Brand Analysis",
                "👥 Manager Analysis",
                "📋 Raw Data",
                "📉 Charts",
                "📦 Brand / FNS Pivot",
                "🏷 Manager / Brand / FNS Pivot",
            ])

            # ---------- TAB 1 (BRAND + GRAND TOTAL) ----------
            with tab1:
                brand_tbl = final_df.groupby("Brand")[["Final Sale Units", "Sales"]].sum()
                brand_tbl.loc["Grand Total"] = brand_tbl.sum()
                st.dataframe(brand_tbl, use_container_width=True)
                download_excel(brand_tbl, "brand_analysis.xlsx", "⬇️ Download Brand Report")

            # ---------- TAB 2 (MANAGER + GRAND TOTAL) ----------
            with tab2:
                mgr_tbl = final_df.groupby("Manager")[["Final Sale Units", "Sales"]].sum()
                mgr_tbl.loc["Grand Total"] = mgr_tbl.sum()
                st.dataframe(mgr_tbl, use_container_width=True)
                download_excel(mgr_tbl, "manager_analysis.xlsx", "⬇️ Download Manager Report")

            # ---------- TAB 3 ----------
            with tab3:
                st.dataframe(final_df, use_container_width=True)
                download_excel(final_df, "raw_data.xlsx", "⬇️ Download Raw Data")

            # ---------- TAB 4 ----------
            with tab4:
                fig = px.bar(
                    final_df.groupby("Brand")["Sales"].sum().reset_index(),
                    x="Brand",
                    y="Sales",
                    title="Sales by Brand"
                )
                st.plotly_chart(fig, use_container_width=True)

            # ---------- TAB 5 ----------
            with tab5:
                base = final_df.groupby(
                    ["FNS", "Brand", "Product Name", "Vendor SKU Codes"]
                )[["Final Sale Units", "Sales", "cost"]].sum()

                grand = pd.DataFrame(
                    [[base["Final Sale Units"].sum(), base["Sales"].sum(), base["cost"].sum()]],
                    index=pd.MultiIndex.from_tuples(
                        [("Grand Total", "", "", "")],
                        names=base.index.names
                    ),
                    columns=base.columns
                )

                st.dataframe(pd.concat([base, grand]), use_container_width=True)
                download_excel(pivot_df, "brand_fns_pivot.xlsx", "⬇️ Download Brand/FNS Pivot")

            # ---------- TAB 6 ----------
            with tab6:
                base = final_df.groupby(
                    ["FNS", "Manager", "Brand", "Product Name", "Vendor SKU Codes"]
                )[["Final Sale Units", "Sales", "cost"]].sum()

                grand = pd.DataFrame(
                    [[base["Final Sale Units"].sum(), base["Sales"].sum(), base["cost"].sum()]],
                    index=pd.MultiIndex.from_tuples(
                        [("Grand Total", "", "", "", "")],
                        names=base.index.names
                    ),
                    columns=base.columns
                )

                st.dataframe(pd.concat([base, grand]), use_container_width=True)
                download_excel(
                    pivot_df,
                    "manager_brand_fns_pivot.xlsx",
                    "⬇️ Download Manager/Brand/FNS Pivot"
                )

            st.success("✅ Analysis generated successfully!")

        except Exception as e:
            st.error(f"❌ Error: {e}")

else:
    st.info("👆 Upload files and click **Generate Analysis**")

