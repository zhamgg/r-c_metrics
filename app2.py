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

        # Debug: Print actual column information (remove these lines once working)
        # st.write("Debug - 3rd Party columns:", list(third_party.columns), f"Count: {len(third_party.columns)}")
        # st.write("Debug - Corporate columns:", list(corporate.columns), f"Count: {len(corporate.columns)}")
        # st.write("Debug - RFI columns:", list(rfi.columns), f"Count: {len(rfi.columns)}")

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
            third_party.columns = ['Project_Name', 'Document_Type', 'Company', 'Date_Received', 'Pages_Minutes'] + [
                f'Extra_{i}' for i in range(len(third_party.columns) - 5)]
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

        # Clean Corporate Marketing data - handle variable column count
        if len(corporate.columns) >= 4:
            new_cols = ['Project_Name', 'Document_Type', 'Pages', 'Date_Received'] + [f'Extra_{i}' for i in
                                                                                      range(len(corporate.columns) - 4)]
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
            new_cols = ['Project_Name', 'Document_Type', 'Current_Relationship', 'Services_Seeking', 'Pages',
                        'Date_Received'] + [f'Extra_{i}' for i in range(len(rfi.columns) - 6)]
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
            third_party_test = pd.read_excel('compliance_marketing_master.xlsx', sheet_name='3rd Party Marketing',
                                             nrows=0)
            corporate_test = pd.read_excel('compliance_marketing_master.xlsx', sheet_name='Corporate Marketing',
                                           nrows=0)
            rfi_test = pd.read_excel('compliance_marketing_master.xlsx', sheet_name='RFI', nrows=0)

            st.write("**Debugging Information:**")
            st.write(f"3rd Party Marketing columns ({len(third_party_test.columns)}): {list(third_party_test.columns)}")
            st.write(f"Corporate Marketing columns ({len(corporate_test.columns)}): {list(corporate_test.columns)}")
            st.write(f"RFI columns ({len(rfi_test.columns)}): {list(rfi_test.columns)}")
        except Exception as debug_error:
            st.write(f"Could not load file for debugging: {debug_error}")

        return None, None, None


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


def create_submission_timeline(data, title):
    """Create a timeline chart of submissions"""
    if data.empty:
        st.write("No data available for timeline")
        return

    # Group by month
    monthly_data = data.groupby('Year_Month').size().reset_index(name='Submissions')
    monthly_data['Date'] = pd.to_datetime(monthly_data['Year_Month'])

    fig = px.line(monthly_data, x='Date', y='Submissions',
                  title=f"{title} - Monthly Submissions",
                  markers=True)
    fig.update_layout(height=400)
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
        fig_bar.update_layout(height=400, yaxis={'categoryorder': 'total ascending'})
        st.plotly_chart(fig_bar, use_container_width=True)


def create_top_subadvisors_chart(data, title, top_n=10):
    """Create top subadvisors chart for 3rd party data"""
    if 'Subadvisor' not in data.columns or data.empty:
        st.write("No subadvisor data available")
        return

    # Get top subadvisors by submission count
    top_subs = data['Subadvisor'].value_counts().head(top_n)

    # Create quarterly breakdown for top subadvisors
    quarterly_data = data[data['Subadvisor'].isin(top_subs.index)].groupby(
        ['Subadvisor', 'Year_Quarter']).size().reset_index(name='Submissions')

    fig = px.bar(quarterly_data, x='Year_Quarter', y='Submissions',
                 color='Subadvisor',
                 title=f"{title} - Top {top_n} Subadvisors by Quarter")
    fig.update_layout(height=500)
    st.plotly_chart(fig, use_container_width=True)


def create_pages_analysis(data, title):
    """Create pages/minutes analysis"""
    if data.empty:
        st.write("No data available for pages analysis")
        return

    col1, col2 = st.columns(2)

    with col1:
        # Pages distribution
        pages_data = data.dropna(subset=['Pages'])
        if not pages_data.empty:
            fig_pages = px.histogram(pages_data, x='Pages', nbins=20,
                                     title=f"{title} - Pages Distribution")
            st.plotly_chart(fig_pages, use_container_width=True)

    with col2:
        # Minutes distribution (if available)
        minutes_data = data.dropna(subset=['Minutes'])
        if not minutes_data.empty:
            fig_minutes = px.histogram(minutes_data, x='Minutes', nbins=20,
                                       title=f"{title} - Minutes Distribution")
            st.plotly_chart(fig_minutes, use_container_width=True)
        else:
            st.write("No minutes data available")


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
        fig_rel.update_layout(height=300, yaxis={'categoryorder': 'total ascending'})
        st.plotly_chart(fig_rel, use_container_width=True)

    with col2:
        # Services Seeking chart
        services_counts = rfi_data['Services_Seeking'].value_counts()
        fig_services = px.bar(x=services_counts.values, y=services_counts.index,
                              orientation='h',
                              title="RFI - Services Seeking")
        fig_services.update_layout(height=300, yaxis={'categoryorder': 'total ascending'})
        st.plotly_chart(fig_services, use_container_width=True)


def main():
    st.title("📊 Compliance Marketing Dashboard")
    st.markdown("---")

    # Load data
    with st.spinner("Loading and processing data..."):
        third_party, corporate, rfi = load_and_clean_data()

    if third_party is None:
        st.error("Failed to load data. Please check the file structure and try again.")
        st.write("**Expected file structure:**")
        st.write("- File name: `compliance_marketing_master.xlsx`")
        st.write("- Sheet 1: '3rd Party Marketing' with at least 5 columns")
        st.write("- Sheet 2: 'Corporate Marketing' with at least 4 columns")
        st.write("- Sheet 3: 'RFI' with at least 6 columns")
        return

    # Sidebar filters
    st.sidebar.header("🔍 Filters")

    # Date range filter
    all_data = pd.concat([third_party, corporate, rfi], ignore_index=True)
    min_date = all_data['Date_Received'].min()
    max_date = all_data['Date_Received'].max()

    if pd.notna(min_date) and pd.notna(max_date):
        date_range = st.sidebar.date_input(
            "Select Date Range",
            value=[min_date.date(), max_date.date()],
            min_value=min_date.date(),
            max_value=max_date.date()
        )

        if len(date_range) == 2:
            start_date, end_date = date_range
            # Filter all datasets
            mask_tp = (third_party['Date_Received'].dt.date >= start_date) & (
                        third_party['Date_Received'].dt.date <= end_date)
            mask_corp = (corporate['Date_Received'].dt.date >= start_date) & (
                        corporate['Date_Received'].dt.date <= end_date)
            mask_rfi = (rfi['Date_Received'].dt.date >= start_date) & (rfi['Date_Received'].dt.date <= end_date)

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
        combined_timeline_data = []

        for area, data in [("3rd Party", third_party), ("Corporate", corporate), ("RFI", rfi)]:
            monthly_counts = data.groupby('Year_Month').size().reset_index(name='Submissions')
            monthly_counts['Area'] = area
            monthly_counts['Date'] = pd.to_datetime(monthly_counts['Year_Month'])
            combined_timeline_data.append(monthly_counts)

        if combined_timeline_data:
            combined_df = pd.concat(combined_timeline_data, ignore_index=True)
            fig_combined = px.line(combined_df, x='Date', y='Submissions',
                                   color='Area', markers=True,
                                   title="Monthly Submissions by Area")
            fig_combined.update_layout(height=500)
            st.plotly_chart(fig_combined, use_container_width=True)

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
        create_submission_timeline(third_party, "3rd Party Marketing")

        # Pages/Minutes analysis
        st.subheader("📊 Pages & Minutes Analysis")
        create_pages_analysis(third_party, "3rd Party Marketing")

    elif selected_area == 'Corporate Marketing':
        st.header("🏢 Corporate Marketing Analysis")
        create_summary_metrics(corporate, "Corporate Marketing")

        st.markdown("---")

        # Document types
        st.subheader("📄 Document Types")
        create_document_type_charts(corporate, "Corporate Marketing")

        # Timeline
        st.subheader("📅 Submission Timeline")
        create_submission_timeline(corporate, "Corporate Marketing")

        # Pages analysis
        st.subheader("📊 Pages Analysis")
        create_pages_analysis(corporate, "Corporate Marketing")

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
        create_submission_timeline(rfi, "RFI")

        # Pages analysis
        st.subheader("📊 Pages Analysis")
        create_pages_analysis(rfi, "RFI")

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
