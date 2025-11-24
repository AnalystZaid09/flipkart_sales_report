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
    top_products_file = st.file_uploader("Upload Top Products CSV File", type=['csv'])

if flipkart_file and top_products_file:
    try:
        # Load files
        flipkart = pd.read_excel(flipkart_file)
        top = pd.read_csv(top_products_file)
        
        # Clean keys
        flipkart['Flipkart Sku Name'] = flipkart['Flipkart Sku Name'].astype(str).str.strip()
        top['SKU ID'] = top['SKU ID'].astype(str).str.strip()
        
        # Drop duplicates from flipkart (VLOOKUP behavior)
        flipkart_unique = flipkart.drop_duplicates(subset=['Flipkart Sku Name'], keep='first')
        
        # Merge data
        final_df = top.merge(
            flipkart_unique[['Flipkart Sku Name', 'Brand Manager', 'Brand']],
            left_on='SKU ID',
            right_on='Flipkart Sku Name',
            how='left'
        ).rename(columns={'Brand Manager': 'Manager'})
        
        final_df = final_df.drop(columns=['Flipkart Sku Name'], errors='ignore')
        
        # Create pivots
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
        
        # Display summary metrics
        st.header("📈 Summary Metrics")
        metric_col1, metric_col2, metric_col3, metric_col4 = st.columns(4)
        
        with metric_col1:
            st.metric("Total Gross Units", f"{int(final_df['Gross Units'].sum())}")
        with metric_col2:
            st.metric("Total Sales", f"₹{final_df['Sales'].sum():,.0f}")
        with metric_col3:
            st.metric("Number of Brands", len(pivot_brand) - 1)
        with metric_col4:
            st.metric("Number of Managers", len(pivot_manager) - 1)
        
        # Tabs for different views
        tab1, tab2, tab3, tab4 = st.tabs(["📊 Brand Analysis", "👥 Manager Analysis", "📋 Raw Data", "📉 Charts"])
        
        with tab1:
            st.subheader("Sales by Brand")
            # Remove Grand Total for chart
            pivot_brand_chart = pivot_brand[:-1].reset_index()
            
            col1, col2 = st.columns(2)
            with col1:
                st.dataframe(pivot_brand, use_container_width=True)
                # Download button for brand pivot
                csv_brand = pivot_brand.to_csv()
                st.download_button(
                    label="📥 Download Brand Pivot Table",
                    data=csv_brand,
                    file_name="pivot_by_brand.csv",
                    mime="text/csv"
                )
            with col2:
                fig = px.bar(pivot_brand_chart, x='Brand', y='Sales', 
                            title='Sales by Brand',
                            labels={'Sales': 'Sales (₹)'},
                            color='Sales',
                            color_continuous_scale='Blues')
                st.plotly_chart(fig, use_container_width=True)
        
        with tab2:
            st.subheader("Sales by Manager")
            pivot_manager_chart = pivot_manager[:-1].reset_index()
            
            col1, col2 = st.columns(2)
            with col1:
                st.dataframe(pivot_manager, use_container_width=True)
                # Download button for manager pivot
                csv_manager = pivot_manager.to_csv()
                st.download_button(
                    label="📥 Download Manager Pivot Table",
                    data=csv_manager,
                    file_name="pivot_by_manager.csv",
                    mime="text/csv"
                )
            with col2:
                fig = px.bar(pivot_manager_chart, x='Manager', y='Sales',
                            title='Sales by Manager',
                            labels={'Sales': 'Sales (₹)'},
                            color='Sales',
                            color_continuous_scale='Greens')
                st.plotly_chart(fig, use_container_width=True)
        
        with tab3:
            st.subheader("Complete Dataset")
            st.dataframe(final_df, use_container_width=True)
            
            # Download button
            csv = final_df.to_csv(index=False)
            st.download_button(
                label="📥 Download Merged Data as CSV",
                data=csv,
                file_name="merged_sales_data.csv",
                mime="text/csv"
            )
        
        with tab4:
            st.subheader("Visual Analytics")
            
            # Pie chart for brand distribution
            col1, col2 = st.columns(2)
            
            with col1:
                fig_pie = px.pie(pivot_brand_chart, values='Sales', names='Brand',
                                title='Sales Distribution by Brand')
                st.plotly_chart(fig_pie, use_container_width=True)
            
            with col2:
                fig_pie_manager = px.pie(pivot_manager_chart, values='Sales', names='Manager',
                                        title='Sales Distribution by Manager')
                st.plotly_chart(fig_pie_manager, use_container_width=True)
            
            # Combined comparison
            st.subheader("Gross Units vs Sales Comparison")
            comparison_option = st.radio("Select View:", ["By Brand", "By Manager"])
            
            if comparison_option == "By Brand":
                data_for_chart = pivot_brand_chart
                x_col = 'Brand'
            else:
                data_for_chart = pivot_manager_chart
                x_col = 'Manager'
            
            fig = go.Figure()
            fig.add_trace(go.Bar(name='Gross Units', x=data_for_chart[x_col], 
                                y=data_for_chart['Gross Units'], yaxis='y', offsetgroup=1))
            fig.add_trace(go.Bar(name='Sales', x=data_for_chart[x_col], 
                                y=data_for_chart['Sales'], yaxis='y2', offsetgroup=2))
            
            fig.update_layout(
                title='Gross Units and Sales Comparison',
                xaxis=dict(title=x_col),
                yaxis=dict(title='Gross Units', side='left'),
                yaxis2=dict(title='Sales (₹)', overlaying='y', side='right'),
                barmode='group'
            )
            st.plotly_chart(fig, use_container_width=True)
        
        st.success("✅ Data processed successfully!")
        
    except Exception as e:
        st.error(f"❌ Error processing files: {str(e)}")
        st.info("Please ensure your files have the correct format and column names.")

else:
    st.info("👆 Please upload both files to begin the analysis.")
    
    with st.expander("ℹ️ Required File Format"):
        st.markdown("""
        **Flipkart PM Excel File should contain:**
        - `Flipkart Sku Name`
        - `Brand Manager`
        - `Brand`
        
        **Top Products CSV File should contain:**
        - `SKU ID`
        - `Gross Units`
        - `Sales`
        """)