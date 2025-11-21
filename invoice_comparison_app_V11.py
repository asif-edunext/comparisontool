import pandas as pd
import streamlit as st
import io 
import re 
import altair as alt

# --- CONFIGURATION ---
st.set_page_config(
    page_title="Invoice Comparison Tool", 
    layout="wide",
    initial_sidebar_state="expanded" 
)

# Custom function to apply conditional formatting based on the status columns
def color_summary_table(s):
    if s['MissingInSheets'] > 0:
        return ['background-color: #FFF3CD'] * len(s)
    elif s['IsDuplicate']:
        return ['background-color: #F8D7DA'] * len(s)
    elif s['AppearsInAllSheets']:
        return ['background-color: #D4EDDA'] * len(s)
    else:
        return [''] * len(s)

# Helper function to convert dataframe to Excel
def to_excel(df, sheet_name='Sheet1', engine='openpyxl'):
    """Converts a DataFrame to an in-memory Excel file buffer."""
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl') 
    
    if isinstance(df, pd.io.formats.style.Styler):
        df.data.to_excel(writer, index=False, sheet_name=sheet_name)
    else:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        
    writer.close() 
    processed_data = output.getvalue()
    return processed_data

# Initialize session state
if 'current_filter' not in st.session_state:
    st.session_state['current_filter'] = None
if 'processed' not in st.session_state:
    st.session_state['processed'] = False
if 'amount_cols_to_process' not in st.session_state:
    st.session_state['amount_cols_to_process'] = []
if 'invoice_col' not in st.session_state:
    st.session_state['invoice_col'] = 'InvoiceNumber'
if 'active_sheets' not in st.session_state:
    st.session_state['active_sheets'] = []


# Helper function to generate filtered data based on click
def filter_invoices(filter_type):
    df = st.session_state['final_summary']
    combined_df = st.session_state['combined']
    invoice_col_name = st.session_state.get('invoice_col', 'InvoiceNumber')

    if filter_type == 'all_sheets':
        filtered_df = df[df['AppearsInAllSheets'] == True].copy()
        title = "Invoices Available in ALL Sheets (Summary)"
        
    elif filter_type == 'missing':
        filtered_df = df[df['MissingInSheets'] > 0].copy()
        title = f"Missing Invoices (in 1 or more sheets) (Summary)"
        
    elif filter_type == 'duplicates':
        duplicate_invoice_list = df[df['IsDuplicate'] == True][invoice_col_name].tolist()
        
        if duplicate_invoice_list:
            filtered_df = combined_df[combined_df[invoice_col_name].isin(duplicate_invoice_list)].copy()
            filtered_df.sort_values([invoice_col_name, 'Sheet'], inplace=True)
            filtered_df.insert(0, 'S. No.', range(1, 1 + len(filtered_df)))
            title = "CROSS-SHEET Duplicates (All Rows)"
        else:
            filtered_df = pd.DataFrame()
            title = "CROSS-SHEET Duplicates (All Rows)"
            
    elif filter_type == 'total':
        filtered_df = df.copy()
        title = "TOTAL Unique Invoices (Summary)"
        
    else:
        return None, None
    
    if 'S. No.' in filtered_df.columns and filter_type != 'duplicates':
        cols = ['S. No.'] + [col for col in filtered_df if col != 'S. No.']
        filtered_df = filtered_df[cols]
        
    return filtered_df, title


# --- SIDEBAR FOR INPUTS ---
uploaded_file = None
with st.sidebar:
    st.image("https://webtel.in/Images/webtel-logo.png", width=250)
    st.markdown("---")

    uploaded_file = st.file_uploader("**Upload Excel file (.xlsx)**", type=["xlsx"])

    if uploaded_file:
        try:
            excel_file = pd.ExcelFile(uploaded_file)
            sheet_names = excel_file.sheet_names
            st.success(f"✅ Sheets detected: {len(sheet_names)}")

            st.markdown("##### 📝 Select Sheets")
            
            # --- DYNAMIC SHEET SELECTION ---
            sheet_options_with_none = ["---"] + sheet_names
            
            default_index_1 = 0 if len(sheet_names) > 0 else 0
            default_index_2 = 1 if len(sheet_names) > 1 else default_index_1
            default_index_3 = 3 if len(sheet_names) > 2 else 0

            sheet1_name = st.selectbox("Sheet 1 (Reference)", sheet_names, index=default_index_1, key="sheet_select_1")
            sheet2_name = st.selectbox("Sheet 2", sheet_names, index=default_index_2, key="sheet_select_2")
            sheet3_name = st.selectbox("Sheet 3 (Optional)", sheet_options_with_none, index=default_index_3, key="sheet_select_3")
            # --- END DYNAMIC SHEET SELECTION ---

            st.markdown("---")

            st.markdown("##### 🔑 Column Names")
            invoice_col = st.text_input("Invoice Column Name", "InvoiceNumber")
            st.session_state['invoice_col'] = invoice_col 
            
            st.markdown("##### 💰 Columns for Amount/Value Difference (Max 3)")
            amount_col1 = st.text_input("Column 1 (e.g., Amount)", "Amount")
            amount_col2 = st.text_input("Column 2 (optional, e.g., Tax)", "")
            amount_col3 = st.text_input("Column 3 (optional, e.g., Total)", "")
            
            amount_cols_input = [col.strip() for col in [amount_col1, amount_col2, amount_col3] if col.strip()]
            st.session_state['amount_cols_to_process'] = amount_cols_input
            
            st.markdown("---")

            if st.button("▶️ Run Comparison"):
                st.session_state['current_filter'] = None
                
                # --- DYNAMIC PROCESSING LOGIC ---
                try:
                    active_sheet_names = [sheet1_name, sheet2_name]
                    if sheet3_name != "---" and sheet3_name:
                        active_sheet_names.append(sheet3_name)
                    
                    st.session_state['active_sheets'] = active_sheet_names
                    total_sheets = len(active_sheet_names)

                    cols_to_read = [invoice_col] + amount_cols_input
                    
                    sheet_data = {}
                    for name in active_sheet_names:
                        sheet_data[name] = pd.read_excel(
                            uploaded_file, 
                            sheet_name=name, 
                            dtype={invoice_col: str}, 
                            usecols=lambda x: x in cols_to_read
                        )

                    def clean_cols(df):
                        df.columns = [c.strip() for c in df.columns]
                        return df
                    
                    for name in active_sheet_names:
                        sheet_data[name] = clean_cols(sheet_data[name])

                    all_cols_found = True
                    for name in active_sheet_names:
                        if invoice_col not in sheet_data[name].columns:
                            st.error(f"❌ Column '{invoice_col}' not found in sheet: '{name}'.")
                            all_cols_found = False
                    
                    if all_cols_found:
                        def prepare(df, name):
                            cols_to_keep = [invoice_col] + [col for col in amount_cols_input if col in df.columns]
                            
                            temp = df[cols_to_keep].dropna(subset=[invoice_col]).copy()
                            temp['Sheet'] = name 
                            temp[invoice_col] = temp[invoice_col].astype(str).str.strip()
                            temp[invoice_col] = temp[invoice_col].str.replace(r'\.0$', '', regex=True) 
                            temp[invoice_col] = temp[invoice_col].str.replace('\xa0', '').str.replace('\u200b', '')
                            temp = temp[temp[invoice_col].str.lower() != 'nan']
                            temp = temp[temp[invoice_col] != '']

                            for col in amount_cols_input:
                                if col in temp.columns:
                                    temp[col] = pd.to_numeric(temp[col], errors='coerce')
                            return temp 

                        prepared_dfs = []
                        for name in active_sheet_names:
                            prepared_dfs.append(prepare(sheet_data[name], name))

                        combined = pd.concat(prepared_dfs, ignore_index=True)
                        
                        summary = (
                            combined.groupby(invoice_col)
                            .agg(
                                SheetsAvailableIn=('Sheet', lambda x: ', '.join(sorted(set(x)))),
                                TotalCount=('Sheet', 'count')
                            )
                            .reset_index()
                        )
                        
                        summary['MissingInSheets'] = summary['TotalCount'].apply(lambda x: total_sheets - x if x < total_sheets else 0)
                        summary['IsDuplicate'] = summary['TotalCount'] > total_sheets
                        summary['AppearsInAllSheets'] = (summary['TotalCount'] >= total_sheets) & (summary['IsDuplicate'] == False)

                        final_summary = summary
                        
                        if amount_cols_input:
                            for col in amount_cols_input:
                                if col in combined.columns:
                                    pivot_amounts = combined.pivot_table(
                                        index=invoice_col, 
                                        columns='Sheet', 
                                        values=col, 
                                        aggfunc='first'
                                    ).reset_index()
                                    
                                    diff_col_name = f"Difference_{col}"
                                    pivot_amounts[diff_col_name] = ( 
                                        pivot_amounts.drop(columns=[invoice_col]) 
                                        .apply(lambda x: x.max() - x.min() if x.count() > 1 else 0, axis=1) 
                                    ) 
                                    
                                    pivot_amounts.columns = [
                                        f"{sheet_name.strip()}_{col}" if sheet_name in active_sheet_names else sheet_name 
                                        for sheet_name in pivot_amounts.columns
                                    ]
                                    
                                    final_summary = pd.merge(
                                        final_summary, 
                                        pivot_amounts.drop(columns=col, errors='ignore'), 
                                        on=invoice_col, 
                                        how='outer'
                                    )

                        final_summary.insert(0, 'S. No.', range(1, 1 + len(final_summary)))
                        
                        st.session_state['final_summary'] = final_summary
                        st.session_state['combined'] = combined
                        
                        duplicates_within_sheets = (
                            combined.groupby(['Sheet', invoice_col])
                            .size()
                            .reset_index(name='Count')
                            .query('Count > 1')
                            .reset_index(drop=True)
                        )
                        duplicates_within_sheets.insert(0, 'S. No.', range(1, 1 + len(duplicates_within_sheets)))
                        st.session_state['duplicates'] = duplicates_within_sheets
                        st.session_state['processed'] = True 
                        st.rerun() 

                except Exception as e:
                    st.error(f"⚠️ Error during comparison logic: {e}")

        except Exception as e:
            st.error(f"❌ Could not read Excel file/sheets: {e}")

    st.markdown("---")
    st.caption("Developed with Streamlit.")
# --- END OF SIDEBAR ---

# === MAIN CONTENT START ===
if st.session_state['processed']:
    final_summary = st.session_state['final_summary']
    duplicates_within_sheets = st.session_state['duplicates']
    combined = st.session_state['combined']
    
    
    
    # =======================================================
    # DYNAMIC GRAPHICAL STATISTICS (4. Visuals)
    # =======================================================
    st.subheader("📊 Visual Breakdown")
    
    chart_data = pd.DataFrame({
        'Status': ['In All Sheets', 'Missing in Some', 'Has Duplicates'],
        'Count': [
            len(final_summary[final_summary['AppearsInAllSheets'] == True]),
            len(final_summary[final_summary['MissingInSheets'] > 0]),
            len(final_summary[final_summary['IsDuplicate'] == True])
        ]
    })
    
    sheet_coverage = combined.groupby('Sheet')[st.session_state['invoice_col']].nunique().reset_index(name='Unique_Invoice_Count')
    
    # Data for the Internal Duplicates chart
    internal_dupe_summary = duplicates_within_sheets.groupby('Sheet').size().reset_index(name='Duplicate_Invoice_Count')

    # Changed from st.columns(2) to st.columns(3)
    chart_col1, chart_col2, chart_col3 = st.columns(3)
    
    with chart_col1:
        st.markdown("##### Invoice Status Breakdown")
        st.bar_chart(
            chart_data, 
            x='Status', 
            y='Count', 
            color='Status',
            height=350,
            use_container_width=True
        )

    with chart_col2:
        st.markdown("##### Unique Invoices per Sheet")
        bar_chart = alt.Chart(sheet_coverage).mark_bar().encode(
            x=alt.X('Sheet', sort='-y'),
            y=alt.Y('Unique_Invoice_Count', title='Unique Invoice Count'),
            tooltip=['Sheet', 'Unique_Invoice_Count'],
            color=alt.Color('Sheet')
        ).properties(
            title="Unique Invoices by Sheet"
        )
        st.altair_chart(bar_chart, use_container_width=True)
        
    # Third column for the new graph
    with chart_col3:
        st.markdown("##### Internal Duplicates per Sheet")
        dupe_bar_chart = alt.Chart(internal_dupe_summary).mark_bar().encode(
            x=alt.X('Sheet', sort='-y'),
            y=alt.Y('Duplicate_Invoice_Count', title='Duplicate Invoice Count'),
            tooltip=['Sheet', 'Duplicate_Invoice_Count'],
            color=alt.Color('Sheet', legend=None) 
        ).properties(
            title="Invoices with Duplicates (Internal)"
        )
        st.altair_chart(dupe_bar_chart, use_container_width=True)
        
    # --- Difference Chart ---
    if st.session_state['amount_cols_to_process']:
        st.markdown("---")
        st.subheader("📉 Amount Differences Overview")
        
        diff_chart_data = []
        for col in st.session_state['amount_cols_to_process']:
            diff_col = f"Difference_{col}"
            if diff_col in final_summary.columns:
                non_zero_diff_count = len(final_summary[final_summary[diff_col].abs() > 0.01]) 
                diff_chart_data.append({
                    'Column': col,
                    'Invoices with Difference': non_zero_diff_count
                })

        if diff_chart_data:
            df_diff_chart = pd.DataFrame(diff_chart_data)
            st.markdown("##### Invoices with a Non-Zero Difference")
            st.bar_chart(
                df_diff_chart,
                x='Column',
                y='Invoices with Difference',
                color='Column',
                height=350,
                use_container_width=True
            )

    st.markdown("---") 
    # =======================================================
    # INTERACTIVE METRICS (1. Key Invoice Statistics)
    # =======================================================
    st.subheader("💡 Key Invoice Statistics")
    
    col_stat1, col_stat2, col_stat3, col_stat4 = st.columns(4)
    
    with col_stat1:
        total_unique = len(final_summary)
        if st.button(f"**Total Invoices**\n{total_unique}", key='btn_total'): 
            st.session_state['current_filter'] = 'total'
    
    with col_stat2:
        count_all = len(final_summary[final_summary['AppearsInAllSheets'] == True])
        if st.button(f"**In All Sheets**\n{count_all}", key='btn_all_sheets'): 
            st.session_state['current_filter'] = 'all_sheets'
            
    with col_stat3:
        count_missing = len(final_summary[final_summary['MissingInSheets'] > 0])
        if st.button(f"**Missing**\n{count_missing} ⚠️", key='btn_missing'): 
            st.session_state['current_filter'] = 'missing'

    with col_stat4:
        count_duplicates = len(final_summary[final_summary['IsDuplicate'] == True])
        if st.button(f"**Duplicates**\n{count_duplicates} 🚩", key='btn_duplicates'): 
            st.session_state['current_filter'] = 'duplicates'
        
        # --- Download button for CROSS-SHEET Duplicates (visible at all times) ---
        duplicate_df, duplicate_title = filter_invoices('duplicates')
        
        if duplicate_df is not None and not duplicate_df.empty:
            excel_data_dupes = to_excel(duplicate_df, sheet_name="Cross_Sheet_Duplicates")
            
            st.download_button(
                label="📥 Download Cross-Sheet Duplicates",
                data=excel_data_dupes,
                file_name="Cross_Sheet_Duplicates_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key='download_dupes_metric'
            )
        # --- END Download button ---

    st.markdown("---") 

    # =======================================================
    # DYNAMIC FILTERED LIST (2. Displays filtered view on metric click)
    # =======================================================
    if st.session_state['current_filter']:
        st.subheader("🔍 Filtered Invoice List")
        
        filtered_df, title = filter_invoices(st.session_state['current_filter'])

        if filtered_df is not None and not filtered_df.empty:
            st.success(f"Showing {len(filtered_df)} rows for filter: {title}")
            
            st.dataframe( 
                filtered_df, 
                key=f"filtered_list_{st.session_state['current_filter']}", 
                hide_index=True,
                use_container_width=True 
            )
            
            col_down_filt, col_clear = st.columns([3, 1])
            with col_down_filt:
                excel_data = to_excel(filtered_df, sheet_name=title.replace(" ", "_").replace("(", "").replace(")", ""))
                st.download_button(
                    label=f"📥 Download {title} ({len(filtered_df)})",
                    data=excel_data,
                    file_name=f"{title.replace(' ', '_').replace('(', '').replace(')', '')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            
            with col_clear:
                if st.button("❌ Clear Filter", key='clear_filter_btn'): 
                    st.session_state['current_filter'] = None
                    st.rerun()

        elif filtered_df is not None and filtered_df.empty:
              st.info(f"No invoices found for the filter: {title}")
        
        st.markdown("---") 
    
    # =======================================================
    # MAIN SUMMARY TABLE (3. Full Summary)
    # =======================================================
    st.subheader("📝 Full Comparison Summary")
    
    with st.expander("Click to view Full Summary Table", expanded=True):
        filtered_summary = final_summary.copy()
        
        if not filtered_summary.empty:
            st.info(f"The table below shows all {len(final_summary)} unique invoices. **Hover over column headers to sort and filter.**")
            
            # Apply color styling
            styled_df = filtered_summary.style.apply(color_summary_table, axis=1)

            column_config = {
                col: st.column_config.TextColumn(label=col, disabled=True) 
                for col in filtered_summary.columns
            }

            st.data_editor( 
                styled_df, 
                key='full_summary_table', 
                hide_index=True,
                disabled=False, 
                column_config=column_config, 
                use_container_width=True
            ) 
        else:
            st.warning("⚠️ No unique invoice data found in the combined sheets.")
    
    st.markdown("---")


    # --- Internal Duplicates (5. Duplicates within single sheets) ---
    st.subheader("🔁 Duplicates Within Single Sheets")
    with st.expander("Click to view Internal Duplicates"):
        if not duplicates_within_sheets.empty:
            st.warning("The list below shows invoices that were duplicates *within* Sheet 1, Sheet 2, or Sheet 3 individually. **Use the buttons below to download the duplicates for each specific file.**")
            
            # --- Download Buttons for Internal Duplicates per Sheet (The requested feature) ---
            st.markdown("##### 📥 Download Internal Duplicates (Per Sheet)")
            
            sheets_with_duplicates = duplicates_within_sheets['Sheet'].unique()
            
            # Create columns based on the number of sheets with duplicates (max 4 to avoid overly narrow columns)
            cols = st.columns(min(len(sheets_with_duplicates), 4))
            
            for i, sheet_name in enumerate(sheets_with_duplicates):
                sheet_dupes_df = duplicates_within_sheets[duplicates_within_sheets['Sheet'] == sheet_name].copy()
                
                # Prepare data for download (excluding summary columns)
                df_to_download = sheet_dupes_df.drop(columns=['S. No.', 'Count'], errors='ignore')
                
                excel_data = to_excel(df_to_download, sheet_name=f"{sheet_name}_Internal_Dupes")
                
                with cols[i % 4]: # Use modulo for wrapping columns if more than 4 sheets
                    st.download_button(
                        label=f"📥 **{sheet_name}** ({len(sheet_dupes_df)})",
                        data=excel_data,
                        file_name=f"{sheet_name}_Internal_Duplicates_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f'download_internal_dupes_{sheet_name}'
                    )
            # --- END Download Buttons for Internal Duplicates per Sheet ---

            # --- Display the combined table ---
            column_config_dupes = {
                col: st.column_config.TextColumn(label=col, disabled=True) 
                for col in duplicates_within_sheets.columns
            }

            st.data_editor(
                duplicates_within_sheets, 
                key='internal_duplicates_table', 
                hide_index=True,
                disabled=False, 
                column_config=column_config_dupes, 
                use_container_width=True
            )
        else:
            st.info("No internal duplicates found.")

    st.markdown("---")
    
    # --- Download Options (6. Position Adjusted) ---
    st.subheader("📥 Download All Data")
    
    col_down1, col_down2 = st.columns(2)
    
    with col_down1:
        unfiltered_data = to_excel(final_summary, sheet_name='Summary')
        st.download_button(
            label="Download Full Summary Table (.xlsx)",
            data=unfiltered_data,
            file_name="Invoice_Summary_Table.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    
    with col_down2:
        output = io.BytesIO()
        writer = pd.ExcelWriter(output, engine='openpyxl') 
        final_summary.to_excel(writer, sheet_name='Summary', index=False)
        duplicates_within_sheets.to_excel(writer, sheet_name='Internal_Duplicates', index=False)
        combined.to_excel(writer, sheet_name='Combined_Data', index=False)
        writer.close()
        complete_data = output.getvalue()

        st.download_button(
            label="Download Complete Report (.xlsx)",
            data=complete_data,
            file_name="Invoice_Comparison_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# === Upload Prompt ===
else:
    if not uploaded_file:
          st.info("⬆️ Please upload an Excel file and click 'Run Comparison' in the sidebar to begin processing.")