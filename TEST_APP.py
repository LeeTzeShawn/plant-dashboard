import streamlit as st
import sys
# st.write("🚨 THIS IS TEST FILE 'Z-DRIVE' 🚨")
st.write("Python executable:", sys.executable)
import pandas as pd
import plotly.express as px
import tempfile
import os
from TEST_plant_cleaner import process_plants_to_excel, get_Total_AllocatedCost_perProduct

st.set_page_config(page_title="Plant Data Dashboard", layout="wide")

st.title("Plant Data Processing Dashboard")
st.markdown("Upload an Excel file to process plant data")

### ADD CACHING FUNCTION ###
@st.cache_data
def process_and_cache(upload_file):
    """
    Process the file ONCE and cache the results
    """
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
        tmp_file.write(upload_file.getvalue())
        tmp_path = tmp_file.name
    output_path = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx').name
    process_plants_to_excel(tmp_path, output_path)

    #Clean up input temp file
    os.unlink(tmp_path)

    #Read and return the processed data
    output_data = pd.read_excel(output_path, sheet_name=None)

    #Store output path for later download
    st.session_state['output_path'] = output_path
    return output_data

uploaded_file = st.file_uploader(
    "Choose an Excel file",
    type=['xlsx', 'xlsm'],
    help="Upload your plant data Excel file"
)

if uploaded_file is not None:
    #Process ONLY ONCE and cache
    output_data = process_and_cache(uploaded_file)
    st.success("✅ Processing complete!")
    # with st.spinner("Processing plant data..."):
        # try:
            #Save uploaded file to temporary location
            # with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                # tmp_file.write(uploaded_file.getvalue())
                # tmp_path = tmp_file.name
            
            #Process the file usig your working function
            # output_path = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx').name
            # process_plants_to_excel(tmp_path, output_path)

            #Clean up temp files
            # os.unlink(tmp_path)

            # st.success("✅ Processing complete!")

            #Show what was processed
    st.subheader("📊 Processing Summary")

    #Read the output file to show preview
    # output_data = pd.read_excel(output_path, sheet_name=None)
    st.metric("Sheets Created", len(output_data))

    #Let user preview data
    if st.checkbox("Show data preview"):
        sheet_names = list(output_data.keys())
        selected_sheet = st.selectbox("Select sheet to preview", sheet_names)
        if selected_sheet:
            st.dataframe(output_data[selected_sheet].head(100))

    ####### ADD NEW COST ANALYSIS SECTION HERE
    st.subheader(" 💰 Cost Analysis Dashboard")
    #Get list of unique plant names from the sheet names
    plant_names = set()
    for sheet_name in output_data.keys():
        if sheet_name.endswith('_Inside'):
            plant_names.add(sheet_name.replace('_Inside', ''))
        elif sheet_name.endswith('_NotInside'):
            plant_names.add(sheet_name.replace('_NotInside', ''))
        elif sheet_name.endswith('_LC'):
            plant_names.add(sheet_name.replace('_LC', ''))
    plant_names = sorted(list(plant_names))
    if plant_names:
        selected_plant = st.selectbox("Select Plant for Cost Analysis", plant_names)

        #Get ALL THREE sheets for this plant
        inside_sheet = f"{selected_plant}_Inside"
        notinside_sheet = f"{selected_plant}_NotInside"
        lc_sheet = f"{selected_plant}_LC"
        
        #Check if all three sheets exist
        if inside_sheet in output_data and notinside_sheet in output_data and lc_sheet in output_data:
            #Get the dataframes
            df_inside = output_data[inside_sheet]
            df_notinside = output_data[notinside_sheet]
            df_lc = output_data[lc_sheet]

            #Call the get_Total_AllocatedCost_perProduct funtion imported from plant_cleaner.py
            result_df = get_Total_AllocatedCost_perProduct(df_inside, df_notinside, df_lc, selected_plant)
            if result_df is not None and not result_df.empty:
                st.subheader(f" Cost Analysis for {selected_plant}")
                #Display metrics
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Total Products", len(result_df))
                with col2:
                    st.metric("Total Cost", f"RM {result_df['Total_Allocated_Cost'].sum():,.2f}")
                with col3:
                    st.metric("Average Cost", f"RM {result_df['Total_Allocated_Cost'].mean():,.2f}")
                # with col4:
            # st.metric("Highest Cost", f"RM {result_df['Total_Allocated_Cost'].max():,.2f}")

                #Top N filter
                top_n = st.slider("Show Top N Products", 5, 30, 10, key="top_n_slider")
                top_products = result_df.head(top_n)

                #Bar chart
                fig = px.bar(
                    top_products,
                    x='Product_ID',
                    y='Total_Allocated_Cost',
                    title=f"Top {top_n} Products by Total Allocated Cost - {selected_plant}",
                    color='Total_Allocated_Cost',
                    color_continuous_scale='Viridis',
                    height=500,
                    text=top_products['Total_Allocated_Cost'].apply(lambda x: f'RM{x:.0f}')
                )
                fig.update_traces(textposition='outside')
                fig.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)

                #Pie chart
                if len(result_df)>10:
                    top_10_sum = result_df.head(10)['Total_Allocated_Cost'].sum()
                    others_sum=result_df.iloc[10:]['Total_Allocated_Cost'].sum()
                    pie_data = pd.DataFrame({
                        'Category':list(result_df.head(10)['Product_ID']) + ['Others'],
                        'Cost': list(result_df.head(10)['Total_Allocated_Cost']) + [others_sum]
                    })
                else:
                    pie_data = result_df.rename(columns={'Product_ID':'Category', 'Total_Allocated_Cost':'Cost'})
                
                fig_pie = px.pie(
                    pie_data,
                    values='Cost',
                    names='Category',
                    title=f"Cost Distribution - {selected_plant}",
                    height=450
                )
                st.plotly_chart(fig_pie, use_container_width=True)
                ### NEW ADDITION ###
                # Unit_Rate per Product_ID graph (what graph to use?)
                # DCOH_Inside = getDC_OH_Inside_ConsumptionList(plant_name, plant_key)
                st.write("DEBUG: before NEW graph")
                selected_products = st.multiselect(
                "Select Products",
                df_inside['Product_ID'].unique(),
                default=list(df_inside['Product_ID'].unique()[:5])
                )
                selected_consumptions = st.multiselect(
                    "Select Consumption",
                    df_inside['Consumption'].unique(),
                    default=list(df_inside['Consumption'].unique()[:3])
                )
                filtered_df_inside = df_inside[
                    (df_inside["Product_ID"].isin(selected_products)) &
                    (df_inside["Consumption"].isin(selected_consumptions))
                ]
                    
                fig_bar = px.bar(
                    filtered_df_inside,
                    x='Unit_Rate',
                    y='Product_ID', 
                    color='Consumption',
                    orientation='h'
                )
                if filtered_df_inside.empty:
                    st.warning("No data for selected filters")
                else:
                    fig_bar = px.bar(
                        filtered_df_inside,
                        x='Unit_Rate',
                        y='Product_ID', 
                        color='Consumption',
                        orientation='h'
                    )
                    st.plotly_chart(fig_bar, use_container_width=True)
                #Data table
                with st.expander("View Complete Cost Data"):
                    result_df = result_df.reset_index(drop=True)
                    st.dataframe(result_df)
                
                ##### Download Section for Total Allocated Cost ####
                st.subheader("📊 Download Total Allocated Cost Analysis")
                csv_total = result_df.to_csv(index=False)
                st.download_button(
                    "📥 Download Total Allocated Cost CSV",
                    csv_total,
                    f"{selected_plant}_total_allocated_cost.csv",
                    key="download_total_cost"
                )

                #Search functionality
                st.subheader("🔍 Search Products")
                product_list = result_df['Product_ID'].tolist()
                selected_product = st.selectbox("Select a Product to view details", product_list)

                #Get the selected product's data
                selected_data = result_df[result_df['Product_ID'] == selected_product]

                if not selected_data.empty:
                    # col1, col2, st.columns(2)
                    with col1:
                        st.metric("Product ID", selected_product)
                    with col2:
                        st.metric("Total Allocated Cost", f"RM {selected_data['Total_Allocated_Cost'].iloc[0]:,.2f}")
                    #Also show rank
                    rank = result_df[result_df['Product_ID'] == selected_product].index[0] + 1
                    st.metric("Rank", f"#{rank} out of {len(result_df)} products")

    #Download button for original 3 Excel sheets per plant
    
    if 'output_path' in st.session_state:
        with open(st.session_state['output_path'], 'rb') as f:
            st.download_button(
                label="📥 Download Cleaned Excel File",
                data=f,
                file_name="cleaned_plant_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    
    #Clean up output file
    # os.unlink(output_path)

        # except Exception as e:
            # st.error(f"Error: {e}")
            # st.exception(e)
    
else:
    st.info("👈 Please upload an Excel file to get started")


