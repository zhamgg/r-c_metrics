import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from datetime import datetime, timedelta
import re

# Set page config
st.set_page_config(
    page_title="Compliance Marketing Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main > div {
        padding-top: 2rem;
    }
    .stMetric {
        background-color: #f0f2f6;
        border: 1px solid #e0e0e0;
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 0.5rem 0;
    }
    .metric-container {
        display: flex;
        justify-content: space-around;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

@st.cache_data
def load_and_clean_data():
    """Load and clean the Excel data"""
    try:
        # Load all sheets
        third_party = pd.read_excel('compliance_marketing_master.xlsx', sheet_name='3rd Party Marketing')
        corporate = pd.read_excel('compliance_marketing_master.xlsx', sheet_name='Corporate Marketing')
        rfi = pd.read_excel('compliance_marketing_master.xlsx', sheet_name='RFI')
        
        # Validate minimum required columns
        if len(third_party.columns) < 5:
            raise ValueError(f"3rd Party Marketing sheet has {len(third_party.columns)} columns, expected at least 5")
        if len(corporate.columns) < 4:
            raise ValueError(f"Corporate Marketing sheet has {len(corporate.columns)} columns, expected at least 4")
        if len(rfi.columns) < 6:
            raise ValueError(f"RFI sheet has {len(rfi.columns)} columns, expected at least 6")
        
        # Data cleaning function
        def clean_text(text):
            if pd.isna(text) or text == '':
                return text
            text = str(text).strip()
            # Fix common typos
            typo_fixes = {
                'Presnetation': 'Presentation',
                'Presntation': 'Presentation',
                'Product Sheey': 'Product Sheet',
                'Exisiting': 'Existing',
                'Fact Sheet': 'Factsheet'
            }
            return typo_fixes.get(text, text)
        
        # Clean 3rd Party Marketing data - handle variable column count
        if len(third_party.columns) >= 5:
            third_party.columns = ['Project_Name', 'Document_Type', 'Company', 'Date_Received', 'Pages_Minutes'] + [f'Extra_{i}' for i in range(len(third_party.columns) - 5)]
        third_party['Document_Type'] = third_party['Document_Type'].apply(clean_text)
        third_party['Company'] = third_party['Company'].apply(clean_text)
        third_party['Date_Received'] = pd.to_datetime(third_party['Date_Received'], errors='coerce')
        
        # Separate pages and minutes for 3rd party
        def extract_pages_minutes(row):
            doc_type = str(row['Document_Type']).lower()
            value = row['Pages_Minutes']
            
            if pd.isna(value) or value == '':
                return pd.Series([np.nan, np.nan])
            
            # Convert to numeric, handling any errors
            try:
                value = float(value)
                # Filter out obvious errors (negative numbers, extremely large numbers)
                if value < 0 or value > 10000:
                    return pd.Series([np.nan, np.nan])
            except:
                return pd.Series([np.nan, np.nan])
            
            # Determine if it's minutes or pages based on document type
            if 'video' in doc_type or 'audio' in doc_type:
                return pd.Series([np.nan, value])  # [Pages, Minutes]
            else:
                return pd.Series([value, np.nan])  # [Pages, Minutes]
        
        third_party[['Pages', 'Minutes']] = third_party.apply(extract_pages_minutes, axis=1)
        third_party['Area'] = '3rd Party Marketing'
        third_party['Subadvisor'] = third_party['Company']  # Treat company as subadvisor
        # Drop any extra columns
        extra_cols = [col for col in third_party.columns if col.startswith('Extra_')]
        if extra_cols:
            third_party = third_party.drop(extra_cols, axis=1)
        
        # Clean Corporate Marketing data - handle variable column count
        if len(corporate.columns) >= 4:
            new_cols = ['Project_Name', 'Document_Type', 'Pages', 'Date_Received'] + [f'Extra_{i}' for i in range(len(corporate.columns) - 4)]
            corporate.columns = new_cols
        corporate['Document_Type'] = corporate['Document_Type'].apply(clean_text)
        corporate['Date_Received'] = pd.to_datetime(corporate['Date_Received'], errors='coerce')
        corporate['Pages'] = pd.to_numeric(corporate['Pages'], errors='coerce')
        corporate['Minutes'] = np.nan  # No minutes data for corporate
        corporate['Area'] = 'Corporate Marketing'
        # Drop any extra columns
        extra_cols = [col for col in corporate.columns if col.startswith('Extra_')]
        if extra_cols:
            corporate = corporate.drop(extra_cols, axis=1)
        
        # Clean RFI data - handle variable column count
        if len(rfi.columns) >= 6:
            new_cols = ['Project_Name', 'Document_Type', 'Current_Relationship', 'Services_Seeking', 'Pages', 'Date_Received'] + [f'Extra_{i}' for i in range(len(rfi.columns) - 6)]
            rfi.columns = new_cols
        rfi['Document_Type'] = rfi['Document_Type'].apply(clean_text)
        rfi['Current_Relationship'] = rfi['Current_Relationship'].apply(clean_text)
        rfi['Services_Seeking'] = rfi['Services_Seeking'].apply(clean_text)
        rfi['Date_Received'] = pd.to_datetime(rfi['Date_Received'], errors='coerce')
        rfi['Pages'] = pd.to_numeric(rfi['Pages'], errors='coerce')
        rfi['Minutes'] = np.nan  # No minutes data for RFI
        rfi['Area'] = 'RFI'
        # Drop any extra columns
        extra_cols = [col for col in rfi.columns if col.startswith('Extra_')]
        if extra_cols:
            rfi = rfi.drop(extra_cols, axis=1)
        
        # Add time-based columns to all datasets
        for df in [third_party, corporate, rfi]:
            df['Year'] = df['Date_Received'].dt.year
            df['Month'] = df['Date_Received'].dt.month
            df['Quarter'] = df['Date_Received'].dt.quarter
            df['Year_Quarter'] = df['Year'].astype(str) + '-Q' + df['Quarter'].astype(str)
            df['Year_Month'] = df['Date_Received'].dt.strftime('%Y-%m')
        
        return third_party, corporate, rfi
    
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        
        # Try to provide more helpful error information
        try:
            third_party_test = pd.read_excel('compliance_marketing_master.xlsx', sheet_name='3rd Party Marketing', nrows=0)
            corporate_test = pd.read_excel('compliance_marketing_master.xlsx', sheet_name='Corporate Marketing', nrows=0)
            rfi_test = pd.read_excel('compliance_marketing_master.xlsx', sheet_name='RFI', nrows=0)
            
            st.write("**Debugging Information:**")
            st.write(f"3rd Party Marketing columns ({len(third_party_test.columns)}): {list(third_party_test.columns)}")
            st.write(f"Corporate Marketing columns ({len(corporate_test.columns)}): {list(corporate_test.columns)}")
            st.write(f"RFI columns ({len(rfi_test.columns)}): {list(rfi_test.columns)}")
        except Exception as debug_error:
            st.write(f"Could not load file for debugging: {debug_error}")
        
        return None, None, None

def show_data_quality_issues(third_party_data, corporate_data, rfi_data):
    """Show specific data quality issues for user to fix"""
    today = pd.Timestamp.now().normalize()
    issues_found = False
    
    st.subheader("🔍 Data Quality Issues Found")
    st.write("Here are the specific records with data quality issues that you can correct in your Excel file:")
    
    # Check for future dates
    future_issues = []
    
    # Check 3rd Party Marketing
    if not third_party_data.empty:
        future_3rd = third_party_data[third_party_data['Date_Received'] > today].copy()
        if not future_3rd.empty:
            future_3rd['Sheet'] = '3rd Party Marketing'
            future_3rd['Issue'] = 'Future Date'
            future_issues.append(future_3rd[['Sheet', 'Project_Name', 'Date_Received', 'Issue']])
    
    # Check Corporate Marketing  
    if not corporate_data.empty:
        future_corp = corporate_data[corporate_data['Date_Received'] > today].copy()
        if not future_corp.empty:
            future_corp['Sheet'] = 'Corporate Marketing'
            future_corp['Issue'] = 'Future Date'
            future_issues.append(future_corp[['Sheet', 'Project_Name', 'Date_Received', 'Issue']])
    
    # Check RFI
    if not rfi_data.empty:
        future_rfi = rfi_data[rfi_data['Date_Received'] > today].copy()
        if not future_rfi.empty:
            future_rfi['Sheet'] = 'RFI'
            future_rfi['Issue'] = 'Future Date'
            future_issues.append(future_rfi[['Sheet', 'Project_Name', 'Date_Received', 'Issue']])
    
    if future_issues:
        issues_found = True
        st.error("⚠️ **Future Dates Found:**")
        all_future = pd.concat(future_issues, ignore_index=True)
        all_future['Date_Received'] = all_future['Date_Received'].dt.strftime('%Y-%m-%d %H:%M:%S')
        st.dataframe(all_future, use_container_width=True, hide_index=True)
        
        st.write("**To fix:** Open your Excel file and correct these dates to valid historical dates.")
    
    # Check for missing critical data
    missing_issues = []
    
    # Missing dates
    for name, data in [('3rd Party Marketing', third_party_data), ('Corporate Marketing', corporate_data), ('RFI', rfi_data)]:
        if not data.empty:
            missing_dates = data[data['Date_Received'].isna()].copy()
            if not missing_dates.empty:
                missing_dates['Sheet'] = name
                missing_dates['Issue'] = 'Missing Date'
                missing_issues.append(missing_dates[['Sheet', 'Project_Name', 'Issue']])
    
    if missing_issues:
        issues_found = True
        st.warning("⚠️ **Missing Dates Found:**")
        all_missing = pd.concat(missing_issues, ignore_index=True)
        st.dataframe(all_missing, use_container_width=True, hide_index=True)
        
        st.write("**To fix:** Open your Excel file and add valid dates for these records.")
    
    # Check for extremely high/low page counts (potential data entry errors)
    page_issues = []
    
    for name, data in [('3rd Party Marketing', third_party_data), ('Corporate Marketing', corporate_data), ('RFI', rfi_data)]:
        if not data.empty and 'Pages' in data.columns:
            # Flag pages > 500 or < 0 as potential errors
            problematic_pages = data[
                (data['Pages'] > 500) | (data['Pages'] < 0)
            ].copy()
            
            if not problematic_pages.empty:
                problematic_pages['Sheet'] = name
                problematic_pages['Issue'] = 'Unusual Page Count'
                page_issues.append(problematic_pages[['Sheet', 'Project_Name', 'Pages', 'Issue']])
    
    if page_issues:
        issues_found = True
        st.warning("⚠️ **Unusual Page Counts Found:**")
        all_page_issues = pd.concat(page_issues, ignore_index=True)
        st.dataframe(all_page_issues, use_container_width=True, hide_index=True)
        
        st.write("**To fix:** Review these page counts - they may be data entry errors or need special handling.")
    
    if not issues_found:
        st.success("✅ **Great! No obvious data quality issues found.**")
        st.write("Your data looks clean and ready for analysis.")
    
    return issues_found

def create_summary_metrics(data, area_name):
    """Create summary metrics for an area"""
    total_submissions = len(data)
    total_pages = data['Pages'].sum() if 'Pages' in data.columns else 0
    total_minutes = data['Minutes'].sum() if 'Minutes' in data.columns else 0
    
    # Handle NaN values
    total_pages = 0 if pd.isna(total_pages) else int(total_pages)
    total_minutes = 0 if pd.isna(total_minutes) else int(total_minutes)
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Submissions", f"{total_submissions:,}")
    with col2:
        st.metric("Pages", f"{total_pages:,}")
    with col3:
        st.metric("Minutes", f"{total_minutes:,}")
    with col4:
        # Calculate average pages per submission (excluding NaN)
        avg_pages = data['Pages'].mean()
        avg_pages = 0 if pd.isna(avg_pages) else avg_pages
        st.metric("Avg Pages", f"{avg_pages:.1f}")

def create_submission_timeline(data, title, time_grouping='Monthly'):
    """Create a timeline chart of submissions"""
    if data.empty:
        st.write("No data available for timeline")
        return
    
    # Filter out future dates (beyond today)
    today = pd.Timestamp.now().normalize()
    data_filtered = data[data['Date_Received'] <= today].copy()
    
    if data_filtered.empty:
        st.write("No historical data available for timeline")
        return
    
    # Group by selected time period
    if time_grouping == 'Quarterly':
        grouped_data = data_filtered.groupby('Year_Quarter').size().reset_index(name='Submissions')
        grouped_data['Period'] = grouped_data['Year_Quarter']
        x_label = 'Quarter'
    elif time_grouping == 'Yearly':
        grouped_data = data_filtered.groupby('Year').size().reset_index(name='Submissions')
        grouped_data['Period'] = grouped_data['Year'].astype(str)
        x_label = 'Year'
    else:  # Monthly
        grouped_data = data_filtered.groupby('Year_Month').size().reset_index(name='Submissions')
        grouped_data['Period'] = grouped_data['Year_Month']
        x_label = 'Month'
    
    # Add data quality indicator
    future_records = len(data) - len(data_filtered)
    if future_records > 0:
        st.warning(f"⚠️ Filtered out {future_records} records with future dates")
    
    # Create bar chart with width adjusted based on time grouping
    fig = px.bar(grouped_data, x='Period', y='Submissions', 
                 title=f"{title} - {time_grouping} Submissions (Historical Data Only)")
    
    # Adjust bar gaps based on time grouping (fewer periods = need bigger gaps to keep bars reasonable width)
    if time_grouping == 'Yearly':
        bargap_value = 0.7  # Much bigger gaps for yearly (fewer bars)
    elif time_grouping == 'Quarterly':
        bargap_value = 0.3  # Medium gaps for quarterly
    else:  # Monthly
        bargap_value = 0.15  # Smaller gaps for monthly (more bars)
    
    fig.update_layout(
        height=400, 
        xaxis_title=x_label, 
        yaxis_title='Number of Submissions',
        bargap=bargap_value
    )
    
    # Fix x-axis ticks for yearly view to show only whole years
    if time_grouping == 'Yearly':
        # Get the actual years from the data
        years = sorted([int(year) for year in grouped_data['Period'].unique()])
        fig.update_xaxes(
            tickmode='array',
            tickvals=years,
            ticktext=[str(year) for year in years]
        )
    
    st.plotly_chart(fig, use_container_width=True)

def create_document_type_charts(data, title):
    """Create document type visualization"""
    if data.empty:
        st.write("No data available for document types")
        return
    
    doc_counts = data['Document_Type'].value_counts()
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Pie chart
        fig_pie = px.pie(values=doc_counts.values, names=doc_counts.index, 
                        title=f"{title} - Document Types (Pie Chart)")
        fig_pie.update_traces(textposition='inside', textinfo='percent+label')
        st.plotly_chart(fig_pie, use_container_width=True)
    
    with col2:
        # Bar chart
        fig_bar = px.bar(x=doc_counts.values, y=doc_counts.index, 
                        orientation='h',
                        title=f"{title} - Document Types (Bar Chart)")
        fig_bar.update_layout(height=400, yaxis={'categoryorder':'total ascending'})
        st.plotly_chart(fig_bar, use_container_width=True)

def create_top_subadvisors_chart(data, title, top_n=15):
    """Create top subadvisors table for 3rd party data"""
    if 'Subadvisor' not in data.columns or data.empty:
        st.write("No subadvisor data available")
        return
    
    # Filter out future dates for cleaner analysis
    today = pd.Timestamp.now().normalize()
    data_filtered = data[data['Date_Received'] <= today].copy()
    
    if data_filtered.empty:
        st.write("No historical subadvisor data available")
        return
    
    # Get top subadvisors by submission count
    top_subs_submissions = data_filtered['Subadvisor'].value_counts().head(top_n)
    
    # Get page counts for the same subadvisors
    subadvisor_pages = data_filtered.groupby('Subadvisor')['Pages'].sum()
    
    if top_subs_submissions.empty:
        st.write("No subadvisor data to display")
        return
    
    # Create summary table like the example provided
    summary_data = []
    for i, (subadvisor, submission_count) in enumerate(top_subs_submissions.items(), 1):
        page_count = int(subadvisor_pages.get(subadvisor, 0)) if not pd.isna(subadvisor_pages.get(subadvisor, 0)) else 0
        summary_data.append({
            'NO': i,
            'NAME': subadvisor,
            'SUBMISSIONS': submission_count,
            'PAGE COUNT': page_count
        })
    
    # Add "Others" row for remaining subadvisors
    if len(data_filtered['Subadvisor'].unique()) > top_n:
        remaining_subs = data_filtered[~data_filtered['Subadvisor'].isin(top_subs_submissions.index)]
        others_submissions = len(remaining_subs)
        others_pages = int(remaining_subs['Pages'].sum()) if not pd.isna(remaining_subs['Pages'].sum()) else 0
        others_count = len(data_filtered['Subadvisor'].unique()) - top_n
        
        summary_data.append({
            'NO': top_n + 1,
            'NAME': f'Others ({others_count})',
            'SUBMISSIONS': others_submissions,
            'PAGE COUNT': others_pages
        })
    
    # Add total row
    total_submissions = len(data_filtered)
    total_pages = int(data_filtered['Pages'].sum()) if not pd.isna(data_filtered['Pages'].sum()) else 0
    summary_data.append({
        'NO': '',
        'NAME': 'TOTAL',
        'SUBMISSIONS': total_submissions,
        'PAGE COUNT': total_pages
    })
    
    summary_df = pd.DataFrame(summary_data)
    
    st.write(f"**{title} - Top {len(top_subs_submissions)} Subadvisors Summary**")
    st.dataframe(summary_df, use_container_width=True, hide_index=True)

def create_pages_analysis(data, title, time_grouping='Monthly'):
    """Create pages analysis with time-based aggregation"""
    if data.empty:
        st.write("No data available for pages analysis")
        return
    
    # Filter out future dates and missing pages
    today = pd.Timestamp.now().normalize()
    data_filtered = data[(data['Date_Received'] <= today) & (data['Pages'].notna())].copy()
    
    if data_filtered.empty:
        st.write("No pages data available for analysis")
        return
    
    # Group by selected time period
    if time_grouping == 'Quarterly':
        grouped_pages = data_filtered.groupby('Year_Quarter')['Pages'].sum().reset_index()
        grouped_pages['Period'] = grouped_pages['Year_Quarter']
        x_label = 'Quarter'
    elif time_grouping == 'Yearly':
        grouped_pages = data_filtered.groupby('Year')['Pages'].sum().reset_index()
        grouped_pages['Period'] = grouped_pages['Year'].astype(str)
        x_label = 'Year'
    else:  # Monthly
        grouped_pages = data_filtered.groupby('Year_Month')['Pages'].sum().reset_index()
        grouped_pages['Period'] = grouped_pages['Year_Month']
        x_label = 'Month'
    
    # Create bar chart with width adjusted based on time grouping
    fig = px.bar(grouped_pages, x='Period', y='Pages', 
                 title=f"{title} - {time_grouping} Page Count")
    
    # Adjust bar gaps based on time grouping
    if time_grouping == 'Yearly':
        bargap_value = 0.7  # Much bigger gaps for yearly
    elif time_grouping == 'Quarterly':
        bargap_value = 0.3  # Medium gaps for quarterly
    else:  # Monthly
        bargap_value = 0.15  # Smaller gaps for monthly
    
    fig.update_layout(
        height=400, 
        xaxis_title=x_label, 
        yaxis_title='Total Pages',
        bargap=bargap_value
    )
    
    # Fix x-axis ticks for yearly view to show only whole years
    if time_grouping == 'Yearly':
        years = sorted([int(year) for year in grouped_pages['Period'].unique()])
        fig.update_xaxes(
            tickmode='array',
            tickvals=years,
            ticktext=[str(year) for year in years]
        )
    
    st.plotly_chart(fig, use_container_width=True)
    
    # Show minutes separately if available
    if 'Minutes' in data.columns:
        minutes_data = data_filtered[data_filtered['Minutes'].notna()]
        if not minutes_data.empty:
            if time_grouping == 'Quarterly':
                grouped_minutes = minutes_data.groupby('Year_Quarter')['Minutes'].sum().reset_index()
                grouped_minutes['Period'] = grouped_minutes['Year_Quarter']
            elif time_grouping == 'Yearly':
                grouped_minutes = minutes_data.groupby('Year')['Minutes'].sum().reset_index()
                grouped_minutes['Period'] = grouped_minutes['Year'].astype(str)
            else:  # Monthly
                grouped_minutes = minutes_data.groupby('Year_Month')['Minutes'].sum().reset_index()
                grouped_minutes['Period'] = grouped_minutes['Year_Month']
            
            fig_minutes = px.bar(grouped_minutes, x='Period', y='Minutes', 
                               title=f"{title} - {time_grouping} Minutes Count")
            
            # Use same gap logic for minutes
            fig_minutes.update_layout(
                height=400, 
                xaxis_title=x_label, 
                yaxis_title='Total Minutes',
                bargap=bargap_value
            )
            
            # Fix x-axis ticks for yearly minutes chart
            if time_grouping == 'Yearly':
                years = sorted([int(year) for year in grouped_minutes['Period'].unique()])
                fig_minutes.update_xaxes(
                    tickmode='array',
                    tickvals=years,
                    ticktext=[str(year) for year in years]
                )
            
            st.plotly_chart(fig_minutes, use_container_width=True)

def create_rfi_specific_charts(rfi_data):
    """Create RFI-specific charts"""
    if rfi_data.empty:
        st.write("No RFI data available")
        return
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Current Relationship chart
        rel_counts = rfi_data['Current_Relationship'].value_counts()
        fig_rel = px.bar(x=rel_counts.values, y=rel_counts.index,
                        orientation='h',
                        title="RFI - Current Relationships")
        fig_rel.update_layout(height=300, yaxis={'categoryorder':'total ascending'})
        st.plotly_chart(fig_rel, use_container_width=True)
    
    with col2:
        # Services Seeking chart
        services_counts = rfi_data['Services_Seeking'].value_counts()
        fig_services = px.bar(x=services_counts.values, y=services_counts.index,
                             orientation='h',
                             title="RFI - Services Seeking")
        fig_services.update_layout(height=300, yaxis={'categoryorder':'total ascending'})
        st.plotly_chart(fig_services, use_container_width=True)

def main():
    st.title("📊 Compliance Marketing Dashboard")
    st.markdown("---")
    
    # Clear cache button for development
    if st.sidebar.button("🔄 Refresh Data & Charts"):
        st.cache_data.clear()
        st.rerun()
    
    # Load data
    with st.spinner("Loading and processing data..."):
        third_party_orig, corporate_orig, rfi_orig = load_and_clean_data()
    
    if third_party_orig is None:
        st.error("Failed to load data. Please check the file structure and try again.")
        st.write("**Expected file structure:**")
        st.write("- File name: `compliance_marketing_master.xlsx`")
        st.write("- Sheet 1: '3rd Party Marketing' with at least 5 columns")
        st.write("- Sheet 2: 'Corporate Marketing' with at least 4 columns") 
        st.write("- Sheet 3: 'RFI' with at least 6 columns")
        return
    
    # Make copies for filtering
    third_party = third_party_orig.copy()
    corporate = corporate_orig.copy()
    rfi = rfi_orig.copy()
    
    # Sidebar filters
    st.sidebar.header("🔍 Filters")
    
    # Date range filter
    all_data_combined = pd.concat([third_party_orig, corporate_orig, rfi_orig], ignore_index=True)
    min_date = all_data_combined['Date_Received'].min()
    max_date = all_data_combined['Date_Received'].max()
    today = pd.Timestamp.now().normalize()
    
    # Cap max date to today to avoid future dates by default
    default_max_date = min(max_date, today) if pd.notna(max_date) else today
    
    if pd.notna(min_date) and pd.notna(max_date):
        # Convert to Python dates for the date input widget
        min_date_py = min_date.date() if hasattr(min_date, 'date') else min_date
        max_date_py = max_date.date() if hasattr(max_date, 'date') else max_date
        default_max_date_py = default_max_date.date() if hasattr(default_max_date, 'date') else default_max_date
        
        date_range = st.sidebar.date_input(
            "Select Date Range",
            value=[min_date_py, default_max_date_py],
            min_value=min_date_py,
            max_value=max_date_py,
            help="Timeline charts automatically exclude future dates"
        )
        
        if len(date_range) == 2:
            start_date, end_date = date_range
            # Convert back to pandas timestamps for filtering
            start_timestamp = pd.Timestamp(start_date)
            end_timestamp = pd.Timestamp(end_date)
            
            # Filter all datasets
            mask_tp = (third_party['Date_Received'] >= start_timestamp) & (third_party['Date_Received'] <= end_timestamp)
            mask_corp = (corporate['Date_Received'] >= start_timestamp) & (corporate['Date_Received'] <= end_timestamp)
            mask_rfi = (rfi['Date_Received'] >= start_timestamp) & (rfi['Date_Received'] <= end_timestamp)
            
            third_party = third_party[mask_tp]
            corporate = corporate[mask_corp]
            rfi = rfi[mask_rfi]
    
    # Area selection
    area_options = ['All Areas', '3rd Party Marketing', 'Corporate Marketing', 'RFI']
    selected_area = st.sidebar.selectbox("Select Area", area_options)
    
    # Time grouping
    time_grouping = st.sidebar.selectbox("Time Grouping", ['Monthly', 'Quarterly', 'Yearly'])
    
    # Main content area
    if selected_area == 'All Areas':
        st.header("📈 Overall Summary")
        
        # Create better layout for combined metrics
        tab1, tab2, tab3 = st.tabs(["🏢 3rd Party Marketing", "🏢 Corporate Marketing", "📝 RFI"])
        
        with tab1:
            create_summary_metrics(third_party, "3rd Party")
            
        with tab2:
            create_summary_metrics(corporate, "Corporate")
            
        with tab3:
            create_summary_metrics(rfi, "RFI")
        
        st.markdown("---")
        
        # Add a summary table view
        st.subheader("📊 Quick Comparison")
        
        summary_data = {
            'Area': ['3rd Party Marketing', 'Corporate Marketing', 'RFI'],
            'Submissions': [len(third_party), len(corporate), len(rfi)],
            'Total Pages': [
                int(third_party['Pages'].sum()) if not pd.isna(third_party['Pages'].sum()) else 0,
                int(corporate['Pages'].sum()) if not pd.isna(corporate['Pages'].sum()) else 0,
                int(rfi['Pages'].sum()) if not pd.isna(rfi['Pages'].sum()) else 0
            ],
            'Avg Pages': [
                round(third_party['Pages'].mean(), 1) if not pd.isna(third_party['Pages'].mean()) else 0,
                round(corporate['Pages'].mean(), 1) if not pd.isna(corporate['Pages'].mean()) else 0,
                round(rfi['Pages'].mean(), 1) if not pd.isna(rfi['Pages'].mean()) else 0
            ]
        }
        
        summary_df = pd.DataFrame(summary_data)
        st.dataframe(summary_df, use_container_width=True, hide_index=True)
        
        st.markdown("---")
        
        # Combined timeline
        st.subheader("📅 Submission Timeline Comparison")
        
        # Filter out future dates for all areas
        today = pd.Timestamp.now().normalize()
        combined_timeline_data = []
        
        for area, data in [("3rd Party", third_party), ("Corporate", corporate), ("RFI", rfi)]:
            # Filter historical data only
            data_filtered = data[data['Date_Received'] <= today].copy()
            
            if not data_filtered.empty:
                # Group by selected time period
                if time_grouping == 'Quarterly':
                    grouped_data = data_filtered.groupby('Year_Quarter').size().reset_index(name='Submissions')
                    grouped_data['Period'] = grouped_data['Year_Quarter']
                elif time_grouping == 'Yearly':
                    grouped_data = data_filtered.groupby('Year').size().reset_index(name='Submissions')
                    grouped_data['Period'] = grouped_data['Year'].astype(str)
                else:  # Monthly
                    grouped_data = data_filtered.groupby('Year_Month').size().reset_index(name='Submissions')
                    grouped_data['Period'] = grouped_data['Year_Month']
                
                grouped_data['Area'] = area
                combined_timeline_data.append(grouped_data)
        
        if combined_timeline_data:
            combined_df = pd.concat(combined_timeline_data, ignore_index=True)
            
            # Create grouped bar chart with color coding and responsive bar width
            fig_combined = px.bar(combined_df, x='Period', y='Submissions', 
                                 color='Area', 
                                 title=f"{time_grouping} Submissions by Area (Historical Data Only)",
                                 barmode='group')
            
            # Adjust gaps based on time grouping for consistent bar width
            if time_grouping == 'Yearly':
                bargap_value = 0.6  # Much bigger gaps for yearly
                bargroupgap_value = 0.15
            elif time_grouping == 'Quarterly':
                bargap_value = 0.25  # Medium gaps for quarterly
                bargroupgap_value = 0.05
            else:  # Monthly
                bargap_value = 0.15  # Smaller gaps for monthly
                bargroupgap_value = 0.02
            
            fig_combined.update_layout(
                height=500,
                xaxis_title=time_grouping.replace('ly', ''),
                yaxis_title='Number of Submissions',
                bargap=bargap_value,
                bargroupgap=bargroupgap_value
            )
            
            # Fix x-axis ticks for yearly view in combined chart
            if time_grouping == 'Yearly':
                years = sorted([int(year) for year in combined_df['Period'].unique()])
                fig_combined.update_xaxes(
                    tickmode='array',
                    tickvals=years,
                    ticktext=[str(year) for year in years]
                )
            
            st.plotly_chart(fig_combined, use_container_width=True)
            
            # Add data quality notes
            total_future_records = 0
            for area, data in [("3rd Party", third_party), ("Corporate", corporate), ("RFI", rfi)]:
                future_count = len(data[data['Date_Received'] > today])
                total_future_records += future_count
            
            if total_future_records > 0:
                st.info(f"ℹ️ Note: {total_future_records} records with future dates were excluded from timeline charts")
        else:
            st.write("No historical data available for timeline comparison")
    
    elif selected_area == '3rd Party Marketing':
        st.header("🏢 3rd Party Marketing Analysis")
        create_summary_metrics(third_party, "3rd Party Marketing")
        
        st.markdown("---")
        
        # Top subadvisors
        st.subheader("🏆 Top Subadvisors")
        create_top_subadvisors_chart(third_party, "3rd Party Marketing")
        
        # Document types
        st.subheader("📄 Document Types")
        create_document_type_charts(third_party, "3rd Party Marketing")
        
        # Timeline
        st.subheader("📅 Submission Timeline")
        create_submission_timeline(third_party, "3rd Party Marketing", time_grouping)
        
        # Pages/Minutes analysis
        st.subheader("📊 Pages & Minutes Analysis")
        create_pages_analysis(third_party, "3rd Party Marketing", time_grouping)
    
    elif selected_area == 'Corporate Marketing':
        st.header("🏢 Corporate Marketing Analysis")
        create_summary_metrics(corporate, "Corporate Marketing")
        
        st.markdown("---")
        
        # Document types
        st.subheader("📄 Document Types")
        create_document_type_charts(corporate, "Corporate Marketing")
        
        # Timeline
        st.subheader("📅 Submission Timeline")
        create_submission_timeline(corporate, "Corporate Marketing", time_grouping)
        
        # Pages analysis
        st.subheader("📊 Pages Analysis")
        create_pages_analysis(corporate, "Corporate Marketing", time_grouping)
    
    elif selected_area == 'RFI':
        st.header("📝 RFI Analysis")
        create_summary_metrics(rfi, "RFI")
        
        st.markdown("---")
        
        # Document types
        st.subheader("📄 Document Types")
        create_document_type_charts(rfi, "RFI")
        
        # RFI-specific charts
        st.subheader("🔍 RFI-Specific Analysis")
        create_rfi_specific_charts(rfi)
        
        # Timeline
        st.subheader("📅 Submission Timeline")
        create_submission_timeline(rfi, "RFI", time_grouping)
        
        # Pages analysis
        st.subheader("📊 Pages Analysis")
        create_pages_analysis(rfi, "RFI", time_grouping)
    
    # Raw data view (optional)
    with st.expander("🔍 View Raw Data"):
        if selected_area == '3rd Party Marketing':
            st.dataframe(third_party)
        elif selected_area == 'Corporate Marketing':
            st.dataframe(corporate)
        elif selected_area == 'RFI':
            st.dataframe(rfi)
        else:
            tab1, tab2, tab3 = st.tabs(["3rd Party", "Corporate", "RFI"])
            with tab1:
                st.dataframe(third_party)
            with tab2:
                st.dataframe(corporate)
            with tab3:
                st.dataframe(rfi)

if __name__ == "__main__":
    main()
