import streamlit as st
import pandas as pd
import plotly.express as px
import warnings

warnings.filterwarnings("ignore", message="Workbook contains no default style*")

# ---------- PAGE CONFIG ----------
st.set_page_config(page_title="Sales Analysis Dashboard", layout="wide")

# ---------- SIDEBAR ----------
with st.sidebar:
    st.title("⚙️ Controls")
    st.markdown("""
        **Steps:**
        1. Flipkart PM Excel upload karein  
        2. Top Products file upload karein (CSV / Excel)  
        3. Generate Analysis button click karein  
    """)

# ---------- TITLE ----------
st.title("📊 Sales Analysis Dashboard")

# ---------- FILE UPLOAD ----------
c1, c2 = st.columns(2)
with c1:
    flipkart_file = st.file_uploader(
        "Upload Flipkart PM Excel",
        type=["xlsx", "xls"]
    )
with c2:
    top_products_file = st.file_uploader(
        "Upload Top Products (CSV / Excel)",
        type=["csv", "xlsx", "xls"]
    )

# ---------- GENERATE BUTTON ----------
generate = st.button("🚀 Generate Analysis", use_container_width=True)

# ---------- MAIN ----------
if generate:
    if not flipkart_file or not top_products_file:
        st.error("❌ Please upload both files before generating analysis")
    else:
        try:
            # ---------- LOAD FILES ----------
            flipkart = pd.read_excel(flipkart_file)
            top = (
                pd.read_csv(top_products_file)
                if top_products_file.name.lower().endswith(".csv")
                else pd.read_excel(top_products_file)
            )
            
            # 🔁 If Brand already exists in Top Products, rename it
            if "Brand" in top.columns:
                top = top.rename(columns={"Brand": "Brand1"})

            # ---------- VALIDATION ----------
            if "SKU ID" not in top.columns:
                raise KeyError("Top Products file me 'SKU ID' column missing hai")
            if "Flipkart Sku Name" not in flipkart.columns:
                raise KeyError("Flipkart PM me 'Flipkart Sku Name' missing hai")

            flipkart["Flipkart Sku Name"] = flipkart["Flipkart Sku Name"].astype(str).str.strip()
            top["SKU ID"] = top["SKU ID"].astype(str).str.strip()

            flipkart_unique = flipkart.drop_duplicates("Flipkart Sku Name")

            # ---------- MERGE ----------
            final_df = top.merge(
                flipkart_unique[["Flipkart Sku Name", "Brand Manager", "Brand"]],
                left_on="SKU ID",
                right_on="Flipkart Sku Name",
                how="left"
            ).rename(columns={"Brand Manager": "Manager"}).drop(columns=["Flipkart Sku Name"])

            # ---------- COST ----------
            cp_col = [c for c in flipkart_unique.columns if str(c).lower().startswith("cp")][0]
            final_df["cost"] = final_df["SKU ID"].map(
                flipkart_unique.set_index("Flipkart Sku Name")[cp_col]
            ).fillna(0)

            # ---------- FNS ----------
            fns_col = [c for c in flipkart_unique.columns if str(c).lower().startswith("fns")][0]
            final_df["FNS"] = final_df["SKU ID"].map(
                flipkart_unique.set_index("Flipkart Sku Name")[fns_col]
            ).fillna("")

            # ---------- SALES ----------
            if "Final Sale Amount" in final_df.columns:
                final_df.rename(columns={"Final Sale Amount": "Sales"}, inplace=True)

            if "Gross Units" not in final_df.columns:
                raise KeyError("Gross Units column required")
            
            # ---------- FIX NEGATIVE UNITS ----------
            for col in final_df.columns:
                if col.strip().lower() == "final sale units":
                    final_df[col] = pd.to_numeric(final_df[col], errors="coerce").fillna(0)
                    final_df[col] = final_df[col].clip(lower=0)

            if "Gross Units" in final_df.columns:
                final_df["Gross Units"] = final_df["Gross Units"].clip(lower=0)

            # ---------- METRICS ----------
            st.markdown("---")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Units", int(final_df["Gross Units"].sum()))
            m2.metric("Total Sales", f"₹{final_df['Sales'].sum():,.0f}")
            m3.metric("Brands", final_df["Brand"].nunique())
            m4.metric("Managers", final_df["Manager"].nunique())

            # ---------- BASIC PIVOTS ----------
            pivot_brand = final_df.pivot_table(
                index="Brand",
                values=["Gross Units", "Sales"],
                aggfunc="sum"
            )
            pivot_brand.loc["Grand Total"] = pivot_brand.sum()

            pivot_manager = final_df.pivot_table(
                index="Manager",
                values=["Gross Units", "Sales"],
                aggfunc="sum"
            )
            pivot_manager.loc["Grand Total"] = pivot_manager.sum()

            # ---------- TABS ----------
            tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
                "📊 Brand Analysis",
                "👥 Manager Analysis",
                "📋 Raw Data",
                "📉 Charts",
                "📦 Brand / FNS Pivot",
                "🏷 Manager / Brand / FNS Pivot",
            ])

            # ---------- TAB 1 ----------
            with tab1:
                st.dataframe(pivot_brand, use_container_width=True)
                st.download_button(
                    "📥 Download Brand Pivot",
                    pivot_brand.to_csv(),
                    "pivot_brand.csv"
                )

            # ---------- TAB 2 ----------
            with tab2:
                st.dataframe(pivot_manager, use_container_width=True)
                st.download_button(
                    "📥 Download Manager Pivot",
                    pivot_manager.to_csv(),
                    "pivot_manager.csv"
                )

            # ---------- TAB 3 ----------
            with tab3:
                st.dataframe(final_df, use_container_width=True)
                st.download_button(
                    "📥 Download Raw Data",
                    final_df.to_csv(index=False),
                    "raw_data.csv"
                )

            # ---------- TAB 4 ----------
            with tab4:
                fig = px.bar(
                    pivot_brand.reset_index()[:-1],
                    x="Brand",
                    y="Sales",
                    title="Sales by Brand"
                )
                st.plotly_chart(fig, use_container_width=True)

            # ==================================================
            # TAB 5 – BRAND / FNS (CLEAN)
            # ==================================================
            with tab5:
                agg = ["Gross Units", "Sales", "cost"]

                base = final_df.groupby(["FNS", "Brand"])[agg].sum().reset_index()

                grand = pd.DataFrame({
                    "FNS": ["Grand Total"],
                    "Brand": [""],
                    "Gross Units": [base["Gross Units"].sum()],
                    "Sales": [base["Sales"].sum()],
                    "cost": [base["cost"].sum()],
                })

                combined = pd.concat([base, grand], ignore_index=True)

                combined = combined.rename(columns={
                    "Gross Units": "Sum of Units",
                    "Sales": "Sum of Sales",
                    "cost": "Sum of Cost",
                })

                display = combined.set_index(["FNS", "Brand"])

                st.dataframe(display, use_container_width=True)
                st.download_button(
                    "📥 Download Brand / FNS Pivot",
                    display.to_csv(),
                    "pivot_brand_fns.csv"
                )

            # ==================================================
            # TAB 6 – MANAGER / BRAND / FNS (CLEAN)
            # ==================================================
            with tab6:
                agg = ["Gross Units", "Sales", "cost"]

                base = (
                    final_df
                    .groupby(["FNS", "Manager", "Brand"])[agg]
                    .sum()
                    .reset_index()
                )

                grand = pd.DataFrame({
                    "FNS": ["Grand Total"],
                    "Manager": [""],
                    "Brand": [""],
                    "Gross Units": [base["Gross Units"].sum()],
                    "Sales": [base["Sales"].sum()],
                    "cost": [base["cost"].sum()],
                })

                combined = pd.concat([base, grand], ignore_index=True)

                combined = combined.rename(columns={
                    "Gross Units": "Sum of Units",
                    "Sales": "Sum of Sales",
                    "cost": "Sum of Cost",
                })

                display = combined.set_index(["FNS", "Manager", "Brand"])

                st.dataframe(display, use_container_width=True)
                st.download_button(
                    "📥 Download Manager / Brand / FNS Pivot",
                    display.to_csv(),
                    "pivot_manager_brand_fns.csv"
                )

            st.success("✅ Analysis generated successfully!")

        except Exception as e:
            st.error(f"❌ Error: {e}")

else:
    st.info("👆 Upload files and click **Generate Analysis**")
