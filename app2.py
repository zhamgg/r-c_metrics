import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from datetime import datetime, timedelta
import warnings

warnings.filterwarnings('ignore')

# Page configuration
st.set_page_config(
    page_title="Marketing Compliance Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 0.5rem 0;
    }
    .stSelectbox label {
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)


@st.cache_data
def load_data():
    """Load and process data from local Excel file"""
    try:
        file_path = "compliance_marketing_master.xlsx"

        # Read both sheets
        third_party_df = pd.read_excel(file_path, sheet_name='3rd Party Marketing')
        corporate_df = pd.read_excel(file_path, sheet_name='Corporate Marketing')

        # Clean column names
        third_party_df.columns = third_party_df.columns.str.strip()
        corporate_df.columns = corporate_df.columns.str.strip()

        # Remove empty rows
        third_party_df = third_party_df.dropna(how='all')
        corporate_df = corporate_df.dropna(how='all')

        # Add marketing type identifier
        third_party_df['Marketing_Type'] = '3rd Party'
        corporate_df['Marketing_Type'] = 'Corporate'

        return third_party_df, corporate_df

    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        st.error("Make sure 'compliance_marketing_master.xlsx' is in the same directory as app.py")
        return pd.DataFrame(), pd.DataFrame()


def convert_excel_date(date_val):
    """Convert Excel date serial number to datetime"""
    if pd.isna(date_val) or date_val == '':
        return None
    if isinstance(date_val, (int, float)):
        try:
            # Excel date serial number to datetime
            return pd.to_datetime('1899-12-30') + pd.to_timedelta(date_val, 'D')
        except:
            return None
    return pd.to_datetime(date_val, errors='coerce')


def get_quarter(date):
    """Extract quarter from date"""
    if pd.isna(date):
        return 'Unknown'
    quarter = f"Q{((date.month - 1) // 3) + 1}"
    return f"{quarter} {date.year}"


def process_data(third_party_df, corporate_df):
    """Process and standardize both datasets"""

    # Process 3rd Party data
    if not third_party_df.empty:
        third_party_processed = third_party_df.copy()
        third_party_processed['Date_Received'] = third_party_processed['Date Received'].apply(convert_excel_date)
        third_party_processed['Quarter'] = third_party_processed['Date_Received'].apply(get_quarter)
        third_party_processed['Pages_Minutes'] = pd.to_numeric(
            third_party_processed['# of pages / # of minutes for videos'],
            errors='coerce'
        ).fillna(0)
        third_party_processed['Document_Type'] = third_party_processed['Document Type'].fillna('Unknown')
        third_party_processed['Project_Name'] = third_party_processed['Project Name'].fillna('Unknown')
    else:
        third_party_processed = pd.DataFrame()

    # Process Corporate data
    if not corporate_df.empty:
        corporate_processed = corporate_df.copy()
        corporate_processed['Date_Received'] = corporate_processed['Date Received'].apply(convert_excel_date)
        corporate_processed['Date_Sent_Compliance'] = corporate_processed['Date Sent To Compliance'].apply(
            convert_excel_date)

        corporate_processed['Quarter'] = corporate_processed['Date_Received'].apply(get_quarter)
        corporate_processed['Pages'] = pd.to_numeric(corporate_processed['# of Pages'], errors='coerce').fillna(0)
        corporate_processed['Document_Type'] = corporate_processed['Document Type'].fillna('Unknown')
        corporate_processed['Project_Name'] = corporate_processed['Project Name'].fillna('Unknown')

        # Calculate workflow metrics (simplified to 2-stage process)
        corporate_processed['Days_To_Compliance'] = (
                corporate_processed['Date_Sent_Compliance'] - corporate_processed['Date_Received']
        ).dt.days
    else:
        corporate_processed = pd.DataFrame()

    return third_party_processed, corporate_processed


def create_quarterly_summary(third_party_df, corporate_df):
    """Create quarterly aggregated metrics"""
    summaries = []

    # 3rd Party quarterly summary
    if not third_party_df.empty:
        tp_summary = third_party_df.groupby('Quarter').agg({
            'Project_Name': 'count',
            'Pages_Minutes': 'sum',
            'Document_Type': lambda x: x.nunique()
        }).round(2)
        tp_summary.columns = ['Projects', 'Total_Pages_Minutes', 'Unique_Doc_Types']
        tp_summary['Marketing_Type'] = '3rd Party'
        tp_summary['Quarter'] = tp_summary.index
        summaries.append(tp_summary.reset_index(drop=True))

    # Corporate quarterly summary
    if not corporate_df.empty:
        corp_summary = corporate_df.groupby('Quarter').agg({
            'Project_Name': 'count',
            'Pages': 'sum',
            'Document_Type': lambda x: x.nunique(),
            'Days_To_Compliance': 'mean'
        }).round(2)
        corp_summary.columns = ['Projects', 'Total_Pages', 'Unique_Doc_Types', 'Avg_Days_To_Compliance']
        corp_summary['Marketing_Type'] = 'Corporate'
        corp_summary['Quarter'] = corp_summary.index
        summaries.append(corp_summary.reset_index(drop=True))

    return summaries


def create_document_type_chart(df, marketing_type, chart_type='bar'):
    """Create document type distribution chart"""
    if df.empty:
        return go.Figure()

    doc_counts = df['Document_Type'].value_counts().head(15)

    if chart_type == 'pie':
        fig = px.pie(
            values=doc_counts.values,
            names=doc_counts.index,
            title=f"{marketing_type} - Document Type Distribution"
        )
    else:
        fig = px.bar(
            x=doc_counts.values,
            y=doc_counts.index,
            orientation='h',
            title=f"{marketing_type} - Top Document Types",
            labels={'x': 'Count', 'y': 'Document Type'}
        )
        fig.update_layout(yaxis={'categoryorder': 'total ascending'})

    fig.update_layout(height=400)
    return fig


def create_quarterly_trend_chart(summaries):
    """Create quarterly trends comparison chart"""
    fig = make_subplots(
        rows=2, cols=2,
        subplot_titles=('Projects by Quarter', 'Pages/Minutes by Quarter',
                        'Document Types by Quarter', 'Workflow Metrics'),
        specs=[[{"secondary_y": False}, {"secondary_y": False}],
               [{"secondary_y": False}, {"secondary_y": False}]]
    )

    colors = {'3rd Party': '#1f77b4', 'Corporate': '#ff7f0e'}

    for summary in summaries:
        if summary.empty:
            continue

        marketing_type = summary['Marketing_Type'].iloc[0]
        color = colors.get(marketing_type, '#2ca02c')

        # Projects trend
        fig.add_trace(
            go.Scatter(x=summary['Quarter'], y=summary['Projects'],
                       mode='lines+markers', name=f'{marketing_type} Projects',
                       line=dict(color=color), marker=dict(size=8)),
            row=1, col=1
        )

        # Pages/Minutes trend
        pages_col = 'Total_Pages_Minutes' if marketing_type == '3rd Party' else 'Total_Pages'
        if pages_col in summary.columns:
            fig.add_trace(
                go.Scatter(x=summary['Quarter'], y=summary[pages_col],
                           mode='lines+markers', name=f'{marketing_type} Pages/Min',
                           line=dict(color=color, dash='dash'), marker=dict(size=8)),
                row=1, col=2
            )

        # Document types trend
        fig.add_trace(
            go.Scatter(x=summary['Quarter'], y=summary['Unique_Doc_Types'],
                       mode='lines+markers', name=f'{marketing_type} Doc Types',
                       line=dict(color=color, dash='dot'), marker=dict(size=8)),
            row=2, col=1
        )

        # Workflow metrics (Corporate only)
        if marketing_type == 'Corporate' and 'Avg_Days_To_Compliance' in summary.columns:
            fig.add_trace(
                go.Scatter(x=summary['Quarter'], y=summary['Avg_Days_To_Compliance'],
                           mode='lines+markers', name='Avg Days to Compliance',
                           line=dict(color='red'), marker=dict(size=8)),
                row=2, col=2
            )

    fig.update_layout(height=600, showlegend=True)
    return fig


def create_workflow_funnel(corporate_df):
    """Create workflow stage funnel for Corporate marketing"""
    if corporate_df.empty:
        return go.Figure()

    total_received = len(corporate_df[corporate_df['Date_Received'].notna()])
    total_compliance = len(corporate_df[corporate_df['Date_Sent_Compliance'].notna()])

    fig = go.Figure(go.Funnel(
        y=["Received", "Sent to Compliance"],
        x=[total_received, total_compliance],
        textinfo="value+percent initial",
        marker=dict(color=["lightblue", "lightgreen"])
    ))

    fig.update_layout(
        title="Corporate Marketing Workflow Funnel",
        height=400
    )

    return fig


def main():
    st.markdown('<h1 class="main-header">📊 Marketing Compliance Dashboard</h1>', unsafe_allow_html=True)

    # Load data automatically
    third_party_df, corporate_df = load_data()

    if not third_party_df.empty or not corporate_df.empty:
        # Process data
        third_party_processed, corporate_processed = process_data(third_party_df, corporate_df)

        # Sidebar filters
        st.sidebar.header("🔍 Filters")

        # Add refresh button
        if st.sidebar.button("🔄 Refresh Data"):
            st.cache_data.clear()
            st.rerun()

        # Marketing type filter
        marketing_types = []
        if not third_party_processed.empty:
            marketing_types.append('3rd Party')
        if not corporate_processed.empty:
            marketing_types.append('Corporate')
        marketing_types.append('Both')

        selected_marketing_type = st.sidebar.selectbox(
            "Marketing Type",
            options=marketing_types,
            index=len(marketing_types) - 1  # Default to 'Both'
        )

        # Quarter filter
        all_quarters = set()
        if not third_party_processed.empty:
            all_quarters.update(third_party_processed['Quarter'].unique())
        if not corporate_processed.empty:
            all_quarters.update(corporate_processed['Quarter'].unique())
        all_quarters = sorted([q for q in all_quarters if q != 'Unknown'])

        selected_quarters = st.sidebar.multiselect(
            "Select Quarters",
            options=all_quarters,
            default=all_quarters
        )

        # Document type filter
        all_doc_types = set()
        if selected_marketing_type in ['3rd Party', 'Both'] and not third_party_processed.empty:
            all_doc_types.update(third_party_processed['Document_Type'].unique())
        if selected_marketing_type in ['Corporate', 'Both'] and not corporate_processed.empty:
            all_doc_types.update(corporate_processed['Document_Type'].unique())

        selected_doc_types = st.sidebar.multiselect(
            "Document Types",
            options=sorted(list(all_doc_types)),
            default=sorted(list(all_doc_types))
        )

        # Filter data based on selections
        filtered_tp = third_party_processed[
            (third_party_processed['Quarter'].isin(selected_quarters)) &
            (third_party_processed['Document_Type'].isin(selected_doc_types))
            ] if not third_party_processed.empty else pd.DataFrame()

        filtered_corp = corporate_processed[
            (corporate_processed['Quarter'].isin(selected_quarters)) &
            (corporate_processed['Document_Type'].isin(selected_doc_types))
            ] if not corporate_processed.empty else pd.DataFrame()

        # Create tabs
        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "📊 Combined Overview",
            "🏢 Corporate Deep Dive",
            "🤝 3rd Party Analysis",
            "📈 Quarterly Trends",
            "📋 Detailed Data"
        ])

        with tab1:
            st.header("Combined Marketing Overview")

            # Key metrics
            col1, col2, col3, col4 = st.columns(4)

            with col1:
                total_projects = len(filtered_tp) + len(filtered_corp)
                st.metric("Total Projects", f"{total_projects:,}")

            with col2:
                tp_pages = filtered_tp['Pages_Minutes'].sum() if not filtered_tp.empty else 0
                corp_pages = filtered_corp['Pages'].sum() if not filtered_corp.empty else 0
                st.metric("Total Pages/Minutes", f"{(tp_pages + corp_pages):,.2f}")

            with col3:
                tp_docs = len(filtered_tp) if not filtered_tp.empty else 0
                st.metric("3rd Party Projects", f"{tp_docs:,}")

            with col4:
                corp_docs = len(filtered_corp) if not filtered_corp.empty else 0
                st.metric("Corporate Projects", f"{corp_docs:,}")

            # Side-by-side document type distributions
            col1, col2 = st.columns(2)

            with col1:
                if not filtered_tp.empty and selected_marketing_type in ['3rd Party', 'Both']:
                    tp_chart = create_document_type_chart(filtered_tp, '3rd Party', 'bar')
                    st.plotly_chart(tp_chart, use_container_width=True, key="overview_tp_chart")
                else:
                    st.info("No 3rd Party data to display")

            with col2:
                if not filtered_corp.empty and selected_marketing_type in ['Corporate', 'Both']:
                    corp_chart = create_document_type_chart(filtered_corp, 'Corporate', 'bar')
                    st.plotly_chart(corp_chart, use_container_width=True, key="overview_corp_chart")
                else:
                    st.info("No Corporate data to display")

            # Combined quarterly comparison
            if selected_marketing_type == 'Both':
                summaries = create_quarterly_summary(filtered_tp, filtered_corp)
                if summaries:
                    quarterly_chart = create_quarterly_trend_chart(summaries)
                    st.plotly_chart(quarterly_chart, use_container_width=True, key="overview_quarterly_chart")

        with tab2:
            st.header("Corporate Marketing Deep Dive")

            if not filtered_corp.empty:
                # Workflow funnel
                col1, col2 = st.columns([2, 1])

                with col1:
                    funnel_chart = create_workflow_funnel(filtered_corp)
                    st.plotly_chart(funnel_chart, use_container_width=True, key="corp_funnel_chart")

                with col2:
                    # Workflow metrics
                    avg_to_compliance = filtered_corp['Days_To_Compliance'].mean()
                    completion_rate = len(filtered_corp[filtered_corp['Date_Sent_Compliance'].notna()]) / len(
                        filtered_corp) * 100 if len(filtered_corp) > 0 else 0

                    st.metric("Avg Days to Compliance",
                              f"{avg_to_compliance:.1f}" if not pd.isna(avg_to_compliance) else "N/A")
                    st.metric("Completion Rate", f"{completion_rate:.1f}%")

                # Document type analysis
                st.subheader("Document Type Analysis")

                col1, col2 = st.columns(2)

                with col1:
                    doc_bar = create_document_type_chart(filtered_corp, 'Corporate', 'bar')
                    st.plotly_chart(doc_bar, use_container_width=True, key="corp_doc_bar_chart")

                with col2:
                    doc_pie = create_document_type_chart(filtered_corp, 'Corporate', 'pie')
                    st.plotly_chart(doc_pie, use_container_width=True, key="corp_doc_pie_chart")

                # Quarterly trends
                st.subheader("Corporate Quarterly Trends")
                corp_quarterly = filtered_corp.groupby('Quarter').agg({
                    'Project_Name': 'count',
                    'Pages': 'sum',
                    'Days_To_Compliance': 'mean'
                }).round(2)

                st.dataframe(corp_quarterly, use_container_width=True)

            else:
                st.info("No Corporate marketing data available for selected filters")

        with tab3:
            st.header("3rd Party Marketing Analysis")

            if not filtered_tp.empty:
                # Key metrics
                col1, col2, col3 = st.columns(3)

                with col1:
                    st.metric("Total Projects", f"{len(filtered_tp):,}")

                with col2:
                    total_pages_minutes = filtered_tp['Pages_Minutes'].sum()
                    st.metric("Total Pages/Minutes", f"{total_pages_minutes:,.2f}")

                with col3:
                    avg_pages_minutes = filtered_tp['Pages_Minutes'].mean()
                    st.metric("Avg Pages/Minutes",
                              f"{avg_pages_minutes:.2f}" if not pd.isna(avg_pages_minutes) else "N/A")

                # Document type distribution
                col1, col2 = st.columns(2)

                with col1:
                    tp_bar = create_document_type_chart(filtered_tp, '3rd Party', 'bar')
                    st.plotly_chart(tp_bar, use_container_width=True, key="tp_doc_bar_chart")

                with col2:
                    # Pages/Minutes distribution
                    if not filtered_tp.empty:
                        fig = px.histogram(
                            filtered_tp[filtered_tp['Pages_Minutes'] > 0],
                            x='Pages_Minutes',
                            title='Distribution of Pages/Minutes',
                            nbins=20
                        )
                        st.plotly_chart(fig, use_container_width=True, key="tp_pages_histogram")

                # Quarterly summary
                st.subheader("3rd Party Quarterly Summary")
                tp_quarterly = filtered_tp.groupby('Quarter').agg({
                    'Project_Name': 'count',
                    'Pages_Minutes': ['sum', 'mean'],
                    'Document_Type': 'nunique'
                }).round(2)

                tp_quarterly.columns = ['Projects', 'Total_Pages_Minutes', 'Avg_Pages_Minutes', 'Unique_Doc_Types']
                st.dataframe(tp_quarterly, use_container_width=True)

            else:
                st.info("No 3rd Party marketing data available for selected filters")

        with tab4:
            st.header("Quarterly Trends & Comparisons")

            # Create quarterly summaries
            summaries = create_quarterly_summary(filtered_tp, filtered_corp)

            if summaries:
                # Combined trends chart
                quarterly_chart = create_quarterly_trend_chart(summaries)
                st.plotly_chart(quarterly_chart, use_container_width=True, key="trends_quarterly_chart")

                # Summary tables
                for i, summary in enumerate(summaries):
                    if not summary.empty:
                        marketing_type = summary['Marketing_Type'].iloc[0]
                        st.subheader(f"{marketing_type} Quarterly Summary")
                        st.dataframe(summary.drop('Marketing_Type', axis=1), use_container_width=True)

            else:
                st.info("No data available for quarterly analysis")

        with tab5:
            st.header("Detailed Data View")

            # Data type selector
            data_view = st.selectbox(
                "Select Data View",
                options=['3rd Party Marketing', 'Corporate Marketing', 'Combined']
            )

            if data_view == '3rd Party Marketing' and not filtered_tp.empty:
                st.subheader("3rd Party Marketing Details")

                # Search functionality
                search_term = st.text_input("Search project names...")
                if search_term:
                    display_data = filtered_tp[
                        filtered_tp['Project_Name'].str.contains(search_term, case=False, na=False)
                    ]
                else:
                    display_data = filtered_tp

                # Select columns to display
                columns_to_show = ['Project_Name', 'Document_Type', 'Date_Received', 'Quarter', 'Pages_Minutes']
                st.dataframe(display_data[columns_to_show], use_container_width=True)

                # Download option
                csv = display_data.to_csv(index=False)
                st.download_button(
                    label="Download 3rd Party data as CSV",
                    data=csv,
                    file_name="third_party_marketing_data.csv",
                    mime="text/csv"
                )

            elif data_view == 'Corporate Marketing' and not filtered_corp.empty:
                st.subheader("Corporate Marketing Details")

                # Search functionality
                search_term = st.text_input("Search project names...")
                if search_term:
                    display_data = filtered_corp[
                        filtered_corp['Project_Name'].str.contains(search_term, case=False, na=False)
                    ]
                else:
                    display_data = filtered_corp

                # Select columns to display
                columns_to_show = ['Project_Name', 'Document_Type', 'Date_Received', 'Quarter',
                                   'Pages', 'Date_Sent_Compliance', 'Days_To_Compliance']
                st.dataframe(display_data[columns_to_show], use_container_width=True)

                # Download option
                csv = display_data.to_csv(index=False)
                st.download_button(
                    label="Download Corporate data as CSV",
                    data=csv,
                    file_name="corporate_marketing_data.csv",
                    mime="text/csv"
                )

            elif data_view == 'Combined':
                # Combine datasets for display
                combined_data = []

                if not filtered_tp.empty:
                    tp_display = filtered_tp[
                        ['Project_Name', 'Document_Type', 'Date_Received', 'Quarter', 'Marketing_Type']].copy()
                    tp_display['Pages'] = filtered_tp['Pages_Minutes']
                    combined_data.append(tp_display)

                if not filtered_corp.empty:
                    corp_display = filtered_corp[
                        ['Project_Name', 'Document_Type', 'Date_Received', 'Quarter', 'Marketing_Type', 'Pages']].copy()
                    combined_data.append(corp_display)

                if combined_data:
                    combined_df = pd.concat(combined_data, ignore_index=True)

                    # Search functionality
                    search_term = st.text_input("Search project names...")
                    if search_term:
                        combined_df = combined_df[
                            combined_df['Project_Name'].str.contains(search_term, case=False, na=False)
                        ]

                    st.dataframe(combined_df, use_container_width=True)

                    # Download option
                    csv = combined_df.to_csv(index=False)
                    st.download_button(
                        label="Download combined data as CSV",
                        data=csv,
                        file_name="combined_marketing_data.csv",
                        mime="text/csv"
                    )
                else:
                    st.info("No data available")

            else:
                st.info("No data available for the selected view")

    else:
        st.error(
            "Could not load data. Please ensure 'compliance_marketing_master.xlsx' is in the same directory as app.py")


if __name__ == "__main__":
    main()
