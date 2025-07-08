import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# Page configuration
st.set_page_config(
    page_title="3rd Party Marketing Metrics Dashboard",
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
def load_data(file_path):
    """Load and process data from Excel file"""
    try:
        # Read all sheets
        excel_file = pd.ExcelFile(file_path)
        
        # Initialize data containers
        quarterly_data = {}
        monthly_data = {}
        
        # Process quarterly sheets
        quarterly_sheets = ['Q3 2024', 'Q4 2024', 'Q1 2025']
        for sheet in quarterly_sheets:
            if sheet in excel_file.sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet)
                # Clean column names
                df.columns = df.columns.str.strip()
                # Remove completely empty rows
                df = df.dropna(how='all')
                # Remove rows where the first column (content name) is empty
                if len(df.columns) > 0:
                    df = df.dropna(subset=[df.columns[0]])
                # Add quarter column
                df['Quarter'] = sheet
                quarterly_data[sheet] = df
        
        # Process monthly sheets
        monthly_sheets = ['January 2025', 'February 2025', 'March 2025', 'April 2025']
        for sheet in monthly_sheets:
            if sheet in excel_file.sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet)
                df.columns = df.columns.str.strip()
                df = df.dropna(how='all')
                if len(df.columns) > 0:
                    df = df.dropna(subset=[df.columns[0]])
                df['Month'] = sheet
                monthly_data[sheet] = df
        
        return quarterly_data, monthly_data
    
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        return {}, {}

def standardize_columns(df, quarter):
    """Standardize column names and add missing columns with zeros"""
    # Standard column mapping
    column_mapping = {
        'Q3 2024 3rd Party Marketing': 'Content_Name',
        'Q4 2024 3rd Party Marketing': 'Content_Name',
        'Q1 2025 3rd Party Marketing': 'Content_Name',
        'Submitter': 'Submitter',
        '# of pieces': 'Pieces',
        '# of pages': 'Pages',
        'Video Length (minutes)': 'Video_Length_Minutes',
        '# of Videos': 'Videos',
        'Aggregate Video length (minutes)': 'Aggregate_Video_Length',
        'Appx extra touches': 'Extra_Touches',
        'Equivalent Total': 'Equivalent_Total'
    }
    
    # Rename columns
    df_renamed = df.rename(columns=column_mapping)
    
    # Add missing columns with default values
    standard_columns = ['Content_Name', 'Submitter', 'Pieces', 'Pages', 
                       'Videos', 'Video_Length_Minutes', 'Aggregate_Video_Length',
                       'Extra_Touches', 'Equivalent_Total']
    
    for col in standard_columns:
        if col not in df_renamed.columns:
            df_renamed[col] = 0
    
    # Clean submitter column - handle NaN and convert to string
    if 'Submitter' in df_renamed.columns:
        df_renamed['Submitter'] = df_renamed['Submitter'].fillna('Unknown').astype(str)
        df_renamed['Submitter'] = df_renamed['Submitter'].str.strip()
        # Replace empty strings with 'Unknown'
        df_renamed['Submitter'] = df_renamed['Submitter'].replace('', 'Unknown')
    
    # Convert numeric columns
    numeric_columns = ['Pieces', 'Pages', 'Videos', 'Video_Length_Minutes', 
                      'Aggregate_Video_Length', 'Extra_Touches', 'Equivalent_Total']
    
    for col in numeric_columns:
        df_renamed[col] = pd.to_numeric(df_renamed[col], errors='coerce').fillna(0)
    
    return df_renamed[standard_columns + ['Quarter']]

def create_aggregate_summary(quarterly_data):
    """Create aggregate summary similar to Agg Trend 2 sheet"""
    all_data = []
    
    for quarter, df in quarterly_data.items():
        df_std = standardize_columns(df, quarter)
        
        # Calculate aggregates
        summary = {
            'Quarter': quarter,
            'Pieces': df_std['Pieces'].sum(),
            'Pages': df_std['Pages'].sum(),
            'Videos': df_std['Videos'].sum(),
            'Video_Length_Minutes': df_std['Video_Length_Minutes'].sum(),
            'Aggregate_Video_Length': df_std['Aggregate_Video_Length'].sum(),
            'Extra_Touches': df_std['Extra_Touches'].sum(),
            'Equivalent_Total': df_std['Equivalent_Total'].sum()
        }
        all_data.append(summary)
    
    return pd.DataFrame(all_data)

def create_trend_chart(df, metrics, title):
    """Create trend line chart"""
    fig = go.Figure()
    
    colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2']
    
    for i, metric in enumerate(metrics):
        if metric in df.columns:
            fig.add_trace(go.Scatter(
                x=df['Quarter'],
                y=df[metric],
                mode='lines+markers',
                name=metric.replace('_', ' ').title(),
                line=dict(color=colors[i % len(colors)], width=3),
                marker=dict(size=8)
            ))
    
    fig.update_layout(
        title=title,
        xaxis_title="Quarter",
        yaxis_title="Count",
        hovermode='x unified',
        height=400,
        showlegend=True
    )
    
    return fig

def create_bar_chart(df, x_col, y_col, title, color_col=None):
    """Create bar chart"""
    if color_col:
        fig = px.bar(df, x=x_col, y=y_col, color=color_col, title=title)
    else:
        fig = px.bar(df, x=x_col, y=y_col, title=title)
    
    fig.update_layout(height=400)
    return fig

def create_submitter_analysis(quarterly_data):
    """Create submitter-specific analysis"""
    all_data = []
    
    for quarter, df in quarterly_data.items():
        df_std = standardize_columns(df, quarter)
        df_std['Quarter'] = quarter
        all_data.append(df_std)
    
    combined_df = pd.concat(all_data, ignore_index=True)
    
    # Group by submitter and quarter
    submitter_summary = combined_df.groupby(['Submitter', 'Quarter']).agg({
        'Pieces': 'sum',
        'Pages': 'sum',
        'Videos': 'sum',
        'Video_Length_Minutes': 'sum',
        'Aggregate_Video_Length': 'sum',
        'Extra_Touches': 'sum',
        'Equivalent_Total': 'sum'
    }).reset_index()
    
    return submitter_summary, combined_df

def main():
    st.markdown('<h1 class="main-header">📊 3rd Party Marketing Metrics Dashboard</h1>', unsafe_allow_html=True)
    
    # File upload
    uploaded_file = st.file_uploader("Upload Excel file", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        quarterly_data, monthly_data = load_data(uploaded_file)
        
        if quarterly_data:
            # Sidebar filters
            st.sidebar.header("🔍 Filters")
            
            # Get all submitters
            all_submitters = set()
            for df in quarterly_data.values():
                df_std = standardize_columns(df, "")
                # Filter out NaN values and convert to string
                valid_submitters = df_std['Submitter'].dropna().astype(str).unique()
                # Remove empty strings and 'nan' strings
                valid_submitters = [s for s in valid_submitters if s.strip() and s.lower() != 'nan']
                all_submitters.update(valid_submitters)
            all_submitters = sorted(list(all_submitters))
            
            # Submitter filter
            selected_submitters = st.sidebar.multiselect(
                "Select Submitters",
                options=all_submitters,
                default=all_submitters
            )
            
            # Quarter filter
            available_quarters = list(quarterly_data.keys())
            selected_quarters = st.sidebar.multiselect(
                "Select Quarters",
                options=available_quarters,
                default=available_quarters
            )
            
            # Metric filter for trends
            metric_options = ['Pieces', 'Pages', 'Videos', 'Video_Length_Minutes', 
                            'Aggregate_Video_Length', 'Extra_Touches', 'Equivalent_Total']
            selected_metrics = st.sidebar.multiselect(
                "Select Metrics for Trend Analysis",
                options=metric_options,
                default=['Pieces', 'Pages', 'Videos']
            )
            
            # Create tabs
            tab1, tab2, tab3, tab4 = st.tabs(["📈 Aggregate Trends", "👥 Submitter Analysis", "📋 Detailed Data", "📊 Monthly Breakdown"])
            
            with tab1:
                st.header("Aggregate Trends Analysis")
                
                # Create aggregate summary
                agg_summary = create_aggregate_summary(quarterly_data)
                
                # Filter by selected quarters
                agg_summary_filtered = agg_summary[agg_summary['Quarter'].isin(selected_quarters)]
                
                # Display key metrics
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    total_pieces = agg_summary_filtered['Pieces'].sum()
                    st.metric("Total Pieces", f"{total_pieces:,}")
                
                with col2:
                    total_pages = agg_summary_filtered['Pages'].sum()
                    st.metric("Total Pages", f"{total_pages:,}")
                
                with col3:
                    total_videos = agg_summary_filtered['Videos'].sum()
                    st.metric("Total Videos", f"{total_videos:,}")
                
                with col4:
                    total_touches = agg_summary_filtered['Extra_Touches'].sum()
                    st.metric("Total Extra Touches", f"{total_touches:,}")
                
                # Trend charts
                if selected_metrics:
                    trend_chart = create_trend_chart(
                        agg_summary_filtered, 
                        selected_metrics, 
                        "Quarterly Trends"
                    )
                    st.plotly_chart(trend_chart, use_container_width=True)
                
                # Bar chart comparison
                col1, col2 = st.columns(2)
                
                with col1:
                    pieces_bar = create_bar_chart(
                        agg_summary_filtered, 
                        'Quarter', 
                        'Pieces', 
                        'Pieces by Quarter'
                    )
                    st.plotly_chart(pieces_bar, use_container_width=True)
                
                with col2:
                    pages_bar = create_bar_chart(
                        agg_summary_filtered, 
                        'Quarter', 
                        'Pages', 
                        'Pages by Quarter'
                    )
                    st.plotly_chart(pages_bar, use_container_width=True)
                
                # Summary table
                st.subheader("Quarterly Summary Table")
                st.dataframe(agg_summary_filtered, use_container_width=True)
            
            with tab2:
                st.header("Submitter Analysis")
                
                submitter_summary, combined_df = create_submitter_analysis(quarterly_data)
                
                # Filter data
                filtered_submitter_data = submitter_summary[
                    (submitter_summary['Submitter'].isin(selected_submitters)) &
                    (submitter_summary['Quarter'].isin(selected_quarters))
                ]
                
                # Top submitters by pieces
                top_submitters = filtered_submitter_data.groupby('Submitter')['Pieces'].sum().sort_values(ascending=False).head(10)
                
                col1, col2 = st.columns(2)
                
                with col1:
                    fig_top = px.bar(
                        x=top_submitters.values,
                        y=top_submitters.index,
                        orientation='h',
                        title="Top 10 Submitters by Total Pieces",
                        labels={'x': 'Total Pieces', 'y': 'Submitter'}
                    )
                    st.plotly_chart(fig_top, use_container_width=True)
                
                with col2:
                    # Submitter trend over time
                    if len(selected_submitters) <= 10:  # Limit for readability
                        submitter_trend = px.line(
                            filtered_submitter_data,
                            x='Quarter',
                            y='Pieces',
                            color='Submitter',
                            title='Submitter Trends Over Time'
                        )
                        st.plotly_chart(submitter_trend, use_container_width=True)
                    else:
                        st.info("Select 10 or fewer submitters to view trend chart")
                
                # Submitter details table
                st.subheader("Submitter Summary")
                submitter_totals = filtered_submitter_data.groupby('Submitter').agg({
                    'Pieces': 'sum',
                    'Pages': 'sum',
                    'Videos': 'sum',
                    'Extra_Touches': 'sum'
                }).round(2).sort_values('Pieces', ascending=False)
                
                st.dataframe(submitter_totals, use_container_width=True)
            
            with tab3:
                st.header("Detailed Data View")
                
                # Quarter selector for detailed view
                selected_quarter_detail = st.selectbox(
                    "Select Quarter for Detailed View",
                    options=available_quarters
                )
                
                if selected_quarter_detail in quarterly_data:
                    df_detail = standardize_columns(quarterly_data[selected_quarter_detail], selected_quarter_detail)
                    
                    # Filter by selected submitters
                    df_filtered = df_detail[df_detail['Submitter'].isin(selected_submitters)]
                    
                    # Display metrics
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Records", len(df_filtered))
                    with col2:
                        st.metric("Unique Submitters", df_filtered['Submitter'].nunique())
                    with col3:
                        st.metric("Total Pieces", df_filtered['Pieces'].sum())
                    
                    # Search functionality
                    search_term = st.text_input("Search content names...")
                    if search_term:
                        df_filtered = df_filtered[df_filtered['Content_Name'].str.contains(search_term, case=False, na=False)]
                    
                    # Display data
                    st.dataframe(df_filtered, use_container_width=True)
                    
                    # Download option
                    csv = df_filtered.to_csv(index=False)
                    st.download_button(
                        label="Download filtered data as CSV",
                        data=csv,
                        file_name=f"{selected_quarter_detail}_filtered_data.csv",
                        mime="text/csv"
                    )
            
            with tab4:
                st.header("Monthly Breakdown")
                
                if monthly_data:
                    st.info("Monthly data processing - this would show detailed monthly breakdowns")
                    
                    # Display available monthly sheets
                    for month, df in monthly_data.items():
                        with st.expander(f"📅 {month}"):
                            st.write(f"Records: {len(df)}")
                            st.dataframe(df.head(), use_container_width=True)
                else:
                    st.info("No monthly data available in the uploaded file")
        
        else:
            st.error("No quarterly data found in the uploaded file")
    
    else:
        st.info("👆 Please upload an Excel file to begin analysis")
        
        # Show sample data structure
        st.subheader("Expected Data Structure")
        st.write("""
        The app expects an Excel file with the following sheets:
        - **Quarterly sheets**: Q3 2024, Q4 2024, Q1 2025, etc.
        - **Monthly sheets**: January 2025, February 2025, etc.
        - Each sheet should contain columns for content name, submitter, and various metrics
        """)

if __name__ == "__main__":
    main()
