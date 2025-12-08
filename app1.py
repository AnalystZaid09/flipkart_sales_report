import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="Sales Analysis Dashboard", layout="wide")

st.title("📊 Sales Analysis Dashboard")
st.markdown("Upload your Flipkart PM Excel file and Top Products CSV to analyze sales by brand and manager.")

# File uploaders
col1, col2 = st.columns(2)

with col1:
    flipkart_file = st.file_uploader("Upload Flipkart PM Excel File", type=['xlsx'])

with col2:
    top_products_file = st.file_uploader("Upload Top Products File (CSV / Excel)", type=['csv','xlsx','xls'])

if flipkart_file and top_products_file:
    try:
        # Load Flipkart Excel file
        flipkart = pd.read_excel(flipkart_file)

        # Detect file format & load Top Products file accordingly
        if top_products_file.name.endswith('.csv'):
            top = pd.read_csv(top_products_file)
        else:
            top = pd.read_excel(top_products_file)

        # Clean keys
        flipkart['Flipkart Sku Name'] = flipkart['Flipkart Sku Name'].astype(str).str.strip()
        top['SKU ID'] = top['SKU ID'].astype(str).str.strip()
        
        # Drop duplicates from flipkart (like VLOOKUP first match)
        flipkart_unique = flipkart.drop_duplicates(subset=['Flipkart Sku Name'], keep='first')

        # Detect CP column
        cp_candidates = [c for c in flipkart_unique.columns if isinstance(c, str) and c.lower().startswith("cp")]
        if cp_candidates:
            cp_col = cp_candidates[0]
        else:
            cp_col = flipkart_unique.columns[8]  # fallback to column I

        cp_map = flipkart_unique.set_index("Flipkart Sku Name")[cp_col]

        # Detect FNS column
        fns_candidates = [c for c in flipkart_unique.columns if isinstance(c, str) and c.lower().startswith("fns")]
        if fns_candidates:
            fns_col = fns_candidates[0]
        else:
            raise KeyError("❌ FNS column not found in Flipkart PM file. Ensure name begins with 'FNS'.")

        fns_map = flipkart_unique.set_index("Flipkart Sku Name")[fns_col]

        # Merge Top + Brand / Manager lookup
        final_df = top.merge(
            flipkart_unique[['Flipkart Sku Name', 'Brand Manager', 'Brand']],
            left_on='SKU ID',
            right_on='Flipkart Sku Name',
            how='left'
        ).rename(columns={'Brand Manager': 'Manager'})

        final_df = final_df.drop(columns=['Flipkart Sku Name'], errors='ignore')

        # Add cost column
        final_df["cost"] = final_df["SKU ID"].map(cp_map)
        final_df["cost"] = pd.to_numeric(final_df["cost"], errors="coerce").fillna(0)

        # Add FNS column
        final_df["FNS"] = final_df["SKU ID"].map(fns_map)
        final_df["FNS"] = final_df["FNS"].fillna("")

        # Add Vendor SKU Code column (VLOOKUP-style using SKU ID)
        vendor_candidates = [
            c for c in flipkart_unique.columns
            if isinstance(c, str) and "vendor" in c.lower() and "sku" in c.lower()
        ]
        if vendor_candidates:
            vendor_col = vendor_candidates[0]
        else:
            # approximate: column D (index 3) from B:D as per your Excel formula
            if len(flipkart_unique.columns) > 3:
                vendor_col = flipkart_unique.columns[3]
            else:
                vendor_col = flipkart_unique.columns[-1]

        # Assume column B (index 1) in PM corresponds to SKU Id for lookup (as in your VLOOKUP range B:D)
        key_col = flipkart_unique.columns[1]
        vendor_map = flipkart_unique.set_index(key_col)[vendor_col]

        final_df["Vendor SKU Code"] = final_df["SKU ID"].map(vendor_map)
        final_df["Vendor SKU Code"] = final_df["Vendor SKU Code"].astype(str).fillna("")

        # Insert cost after Gross Units and FNS after cost
        cols = list(final_df.columns)
        if "Gross Units" in cols and "cost" in cols:
            idx = cols.index("Gross Units")
            cols.insert(idx + 1, cols.pop(cols.index("cost")))
        if "cost" in cols and "FNS" in cols:
            idx = cols.index("cost")
            cols.insert(idx + 1, cols.pop(cols.index("FNS")))
        final_df = final_df[cols]

        # Insert Vendor SKU Code between FNS and FBF Units Percentage (i.e. immediately after FNS)
        cols = list(final_df.columns)
        if "FNS" in cols and "Vendor SKU Code" in cols:
            cols.remove("Vendor SKU Code")
            insert_pos = cols.index("FNS") + 1
            cols.insert(insert_pos, "Vendor SKU Code")
            final_df = final_df[cols]

        # Simple pivots (Brand, Manager)
        pivot_brand = final_df.pivot_table(
            index='Brand',
            values=['Gross Units', 'Sales'],
            aggfunc='sum',
            fill_value=0
        )
        pivot_brand.loc['Grand Total'] = pivot_brand.sum()
        
        pivot_manager = final_df.pivot_table(
            index='Manager',
            values=['Gross Units', 'Sales'],
            aggfunc='sum',
            fill_value=0
        )
        pivot_manager.loc['Grand Total'] = pivot_manager.sum()

        # Summary Metrics
        st.header("📈 Summary Metrics")
        m1, m2, m3, m4 = st.columns(4)

        m1.metric("Total Gross Units", f"{int(final_df['Gross Units'].sum()):,}")
        m2.metric("Total Sales", f"₹{final_df['Sales'].sum():,.0f}")
        m3.metric("Number of Brands", len(pivot_brand) - 1)
        m4.metric("Number of Managers", len(pivot_manager) - 1)

        # Tabs
        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
            "📊 Brand Analysis",
            "👥 Manager Analysis",
            "📋 Raw Data",
            "📉 Charts",
            "📦 Brand / FNS Pivot",
            "🏷 Manager / Brand / FNS Pivot"
        ])
        
        # BRAND TAB
        with tab1:
            st.subheader("Sales by Brand")
            pivot_brand_chart = pivot_brand[:-1].reset_index()
            c1, c2 = st.columns(2)

            with c1:
                st.dataframe(pivot_brand, width="stretch")
                st.download_button("📥 Download Brand Pivot", pivot_brand.to_csv(), "pivot_brand.csv")

            with c2:
                fig = px.bar(pivot_brand_chart, x='Brand', y='Sales',
                             title='Sales by Brand', color='Sales')
                st.plotly_chart(fig, width="stretch")

        # MANAGER TAB
        with tab2:
            st.subheader("Sales by Manager")
            pivot_manager_chart = pivot_manager[:-1].reset_index()
            c1, c2 = st.columns(2)

            with c1:
                st.dataframe(pivot_manager, width="stretch")
                st.download_button("📥 Download Manager Pivot", pivot_manager.to_csv(), "pivot_manager.csv")

            with c2:
                fig = px.bar(pivot_manager_chart, x='Manager', y='Sales',
                             title='Sales by Manager', color='Sales')
                st.plotly_chart(fig, width="stretch")

        # RAW DATA TAB
        with tab3:
            st.subheader("Complete Dataset (with cost, FNS & Vendor SKU Code)")
            st.dataframe(final_df, width="stretch")
            st.download_button("📥 Download Complete CSV", final_df.to_csv(index=False), "final_dataset.csv")

        # CHART TAB
        with tab4:
            st.subheader("Visual Analytics")
            c1, c2 = st.columns(2)

            with c1:
                fig = px.pie(pivot_brand_chart, values='Sales', names='Brand',
                             title='Sales Distribution by Brand')
                st.plotly_chart(fig, width="stretch")

            with c2:
                fig = px.pie(pivot_manager_chart, values='Sales', names='Manager',
                             title='Sales Distribution by Manager')
                st.plotly_chart(fig, width="stretch")

        # BRAND / FNS TAB WITH SUBTOTALS & GRAND TOTAL
        with tab5:
            st.subheader("Brand / FNS Pivot (with totals)")

            agg_cols = ["Gross Units", "Sales", "cost"]

            # Base aggregation at Brand + FNS level
            base_bf = (
                final_df
                .groupby(["Brand", "FNS"], dropna=False)[agg_cols]
                .sum()
                .reset_index()
            )
            base_bf["order"] = 0  # detail rows

            # Brand totals
            brand_totals_bf = (
                base_bf
                .groupby(["Brand"], as_index=False)[agg_cols]
                .sum()
            )
            brand_totals_bf["FNS"] = brand_totals_bf["Brand"] + " Total"
            brand_totals_bf["order"] = 1

            # Grand total
            grand_total_bf = pd.DataFrame({
                "Brand": [""],
                "FNS": ["Grand Total"],
                "Gross Units": [base_bf["Gross Units"].sum()],
                "Sales": [base_bf["Sales"].sum()],
                "cost": [base_bf["cost"].sum()],
                "order": [2],
            })

            # Combine
            combined_bf = pd.concat(
                [base_bf, brand_totals_bf, grand_total_bf],
                ignore_index=True
            )

            combined_bf["is_grand"] = (combined_bf["FNS"] == "Grand Total").astype(int)

            # Sort: Brand → order (detail, brand total, grand) → Gross Units desc
            combined_bf = combined_bf.sort_values(
                by=["is_grand", "Brand", "order", "Gross Units"],
                ascending=[True, True, True, False]
            )

            # Rename columns to "Sum of ..."
            combined_bf = combined_bf.rename(columns={
                "Gross Units": "Sum of Gross Units",
                "Sales": "Sum of Sales",
                "cost": "Sum of Cost"
            })

            # Set MultiIndex Brand → FNS
            display_bf = combined_bf.drop(columns=["order", "is_grand"]).set_index(
                ["Brand", "FNS"]
            )

            st.dataframe(display_bf, width="stretch")
            st.download_button(
                "📥 Download Brand / FNS Pivot",
                display_bf.to_csv(),
                "pivot_brand_fns.csv"
            )

        # MANAGER / BRAND / FNS TAB WITH SUBTOTALS & GRAND TOTAL
        with tab6:
            st.subheader("Manager / Brand / FNS Pivot (with totals)")

            agg_cols = ["Gross Units", "Sales", "cost"]

            # Base aggregation at Manager + Brand + FNS level
            base = (
                final_df
                .groupby(["Manager", "Brand", "FNS"], dropna=False)[agg_cols]
                .sum()
                .reset_index()
            )

            base["order"] = 0  # detail rows

            # Brand totals within each Manager
            brand_totals = (
                base
                .groupby(["Manager", "Brand"], as_index=False)[agg_cols]
                .sum()
            )
            brand_totals["FNS"] = brand_totals["Brand"] + " Total"
            brand_totals["order"] = 1

            # Manager totals
            manager_totals = (
                base
                .groupby(["Manager"], as_index=False)[agg_cols]
                .sum()
            )
            manager_totals["Brand"] = ""
            manager_totals["FNS"] = manager_totals["Manager"] + " Total"
            manager_totals["order"] = 2

            # Grand total
            grand_total = pd.DataFrame({
                "Manager": [""],
                "Brand": [""],
                "FNS": ["Grand Total"],
                "Gross Units": [base["Gross Units"].sum()],
                "Sales": [base["Sales"].sum()],
                "cost": [base["cost"].sum()],
                "order": [3],
            })

            # Combine all
            combined = pd.concat(
                [base, brand_totals, manager_totals, grand_total],
                ignore_index=True
            )

            combined["is_grand"] = (combined["FNS"] == "Grand Total").astype(int)

            # Sort: Manager → Brand → order → Gross Units desc; Grand Total last
            combined = combined.sort_values(
                by=["is_grand", "Manager", "Brand", "order", "Gross Units"],
                ascending=[True, True, True, True, False]
            )

            # Rename columns
            combined = combined.rename(columns={
                "Gross Units": "Sum of Gross Units",
                "Sales": "Sum of Sales",
                "cost": "Sum of Cost"
            })

            # Set index: Manager → Brand → FNS
            display_df = combined.drop(columns=["order", "is_grand"]).set_index(
                ["Manager", "Brand", "FNS"]
            )

            st.dataframe(display_df, width="stretch")
            st.download_button(
                "📥 Download Manager / Brand / FNS Pivot",
                display_df.to_csv(),
                "pivot_manager_brand_fns.csv"
            )

        st.success("✅ Data processed successfully!")

    except Exception as e:
        st.error(f"❌ Error processing files: {str(e)}")
        st.info("Please verify column names and file format.")

else:
    st.info("👆 Please upload both files to begin.")

