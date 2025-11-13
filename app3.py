import streamlit as st
import pandas as pd
import plotly.express as px

# --- Page Configuration ---
st.set_page_config(
    page_title="Global Diplomacy Dashboard",
    page_icon="🌐",
    layout="wide",
)

# (Create the .streamlit/config.toml file for dark theme as discussed)

# --- Data Loading and Transformation ---
@st.cache_data
def load_and_prepare_data(file_path, sheet_name):
    try:
        df_detail = pd.read_excel(file_path, sheet_name=sheet_name)
    except (FileNotFoundError, ValueError) as e:
        st.error(f"FATAL ERROR loading '{file_path}': {e}")
        st.stop()

    required_cols = ['COUNTRY', 'Year', 'OVERALL RANK']
    if not all(col in df_detail.columns for col in required_cols):
        st.error(f"FATAL ERROR: Your Excel file must contain the columns: {required_cols}")
        st.stop()

    df_long = df_detail.groupby(['COUNTRY', 'Year']).size().reset_index(name='Posts')
    df_long = df_long.rename(columns={'COUNTRY': 'Name'})
    df_long['Posts'] = pd.to_numeric(df_long['Posts'], errors='coerce').dropna()

    df_wide = df_long.pivot_table(index='Name', columns='Year', values='Posts').reset_index()
    df_wide.columns = [str(col) for col in df_wide.columns]
    
    return df_wide, df_long, df_detail

# (The calculate_rank_analysis function is unchanged)
@st.cache_data
def calculate_rank_analysis(df, start_year, end_year):
    # ... (code is the same)
    df_start = df[df['Year'] == start_year].copy()
    df_end = df[df['Year'] == end_year].copy()
    df_start['Rank'] = df_start['Posts'].rank(ascending=False, method='min')
    df_end['Rank'] = df_end['Posts'].rank(ascending=False, method='min')
    df_merged = pd.merge(df_start[['Name', 'Rank']], df_end[['Name', 'Rank']], on='Name', suffixes=(f'_{start_year}', f'_{end_year}'))
    df_merged['Rank_Change'] = df_merged[f'Rank_{start_year}'] - df_merged[f'Rank_{end_year}']
    most_improved = df_merged.sort_values(by='Rank_Change', ascending=False)
    biggest_fall = df_merged.sort_values(by='Rank_Change', ascending=True)
    return most_improved.head(10), biggest_fall.head(10)

# --- Main App Layout ---
st.title("🌐 Global Diplomacy Index Analysis")
st.markdown("An interactive dashboard inspired by the Lowy Institute's Global Diplomacy Index.")

df_wide, df_long, df_detail = load_and_prepare_data('globaldiplomacyindex.xlsx', 'Sheet1')

tab1, tab2, tab3, tab4 = st.tabs(["🔑 Key Findings", "🏆 Rankings", "🗺️ Network Map", "📈 Comparison"])

# (Tab 1 is unchanged)
with tab1:
    # ... [Code from previous version] ...
    st.header("Key Findings from the Global Diplomacy Index")
    latest_year = df_long['Year'].max()
    df_long_latest = df_long[df_long['Year'] == latest_year]
    df_detail_latest = df_detail[df_detail['Year'] == latest_year]
    top_performers = df_long_latest.sort_values(by='Posts', ascending=False)
    top_country_name = top_performers.iloc[0]['Name']
    second_country_name = top_performers.iloc[1]['Name']
    top_country_details = df_detail_latest[df_detail_latest['COUNTRY'] == top_country_name].iloc[0]
    second_country_details = df_detail_latest[df_detail_latest['COUNTRY'] == second_country_name].iloc[0]
    top_country_posts = top_performers.iloc[0]['Posts']
    second_country_posts = top_performers.iloc[1]['Posts']
    col1, col2, col3 = st.columns(3)
    col1.metric(label=f"Top Diplomatic Power ({latest_year})", value=top_country_details['COUNTRY'], delta=f"Rank #{top_country_details['OVERALL RANK']}")
    col2.metric(label=f"{top_country_details['COUNTRY']}'s Total Posts", value=int(top_country_posts))
    col3.metric(label=f"Runner Up: {second_country_details['COUNTRY']}", value=f"Rank #{second_country_details['OVERALL RANK']}", delta=f"{int(second_country_posts)} Posts")

# --- TAB 2: RANKINGS (REBUILT AND CORRECTED) ---
with tab2:
    st.header(f"Full Global Diplomacy Ranking")
    
    all_years = sorted(df_detail['Year'].unique())
    year_to_display = st.selectbox("Select a year:", options=all_years, index=len(all_years)-1, key="ranking_year_select")

    # Filter the detailed dataframe for the selected year
    df_ranking = df_detail[df_detail['Year'] == year_to_display].copy()

    # --- THE FIX IS HERE ---
    # 1. Get total posts from our reliable df_long
    df_total_posts = df_long[df_long['Year'] == year_to_display][['Name', 'Posts']]
    df_total_posts = df_total_posts.rename(columns={'Name': 'COUNTRY', 'Posts': 'TOTAL POSTS'})

    # 2. Calculate the counts for each post type for each country from the detailed data
    post_type_counts = df_ranking.groupby('COUNTRY')['POST TYPE TITLE'].value_counts().unstack(fill_value=0)

    # 3. Combine different consulate and embassy types into single columns
    post_type_counts['Consulates'] = post_type_counts.get('Consulate', 0) + post_type_counts.get('Consulate General', 0)
    post_type_counts['Embassies'] = post_type_counts.get('Embassy', 0) + post_type_counts.get('High Commission', 0)

    # 4. Get the unique metadata (POP, GDP, RANK) for each country
    df_meta = df_ranking[['COUNTRY', 'OVERALL RANK', 'POPULATION (M)', 'GDP (B, USD)']].drop_duplicates(subset=['COUNTRY'])

    # 5. Merge all our calculated and metadata tables together
    df_merged = pd.merge(df_meta, df_total_posts, on='COUNTRY')
    # We only merge the columns we just created to avoid errors
    df_final_display = pd.merge(df_merged, post_type_counts[['Embassies', 'Consulates']], on='COUNTRY', how='left')
    
    # Define columns to show and rename them for a professional look
    columns_to_show = ['OVERALL RANK', 'COUNTRY', 'POPULATION (M)', 'GDP (B, USD)', 'TOTAL POSTS', 'Embassies', 'Consulates']
    rename_map = {'OVERALL RANK': 'Rank', 'COUNTRY': 'Country'}
    
    # Ensure all required columns exist before trying to select them
    final_cols_exist = [col for col in columns_to_show if col in df_final_display.columns]
    df_display = df_final_display[final_cols_exist].rename(columns=rename_map)
    df_display = df_display.sort_values(by='Rank').set_index('Rank')
    
    st.dataframe(df_display, use_container_width=True)
    # --- END OF FIX ---

# (The other tabs are unchanged)
with tab3:
    st.header("Visualize a Country's Global Network")
    map_year = df_detail['Year'].max()
    df_map_data = df_detail[df_detail['Year'] == map_year]
    country_list_map = sorted(df_map_data['COUNTRY'].unique())
    selected_country_map = st.selectbox("Select a 'Home' Country:", options=country_list_map, index=country_list_map.index("China") if "China" in country_list_map else 0)
    country_posts = df_map_data[df_map_data['COUNTRY'] == selected_country_map]
    post_locations = country_posts['POST COUNTRY'].value_counts().reset_index()
    post_locations.columns = ['Host Country', 'Number of Posts']
    fig_map = px.scatter_geo(post_locations, locations="Host Country", locationmode='country names', size="Number of Posts")
    fig_map.update_layout(title_text=f"Global Footprint of {selected_country_map} ({map_year})")
    st.plotly_chart(fig_map, use_container_width=True)

with tab4:
    st.header("Compare Diplomatic Networks Over Time")
    all_countries_list_comp = sorted(df_long['Name'].unique())
    default_countries_comp = ['China', 'United States', 'Turkey', 'Japan', 'France', 'Russia']
    valid_defaults_comp = [c for c in default_countries_comp if c in all_countries_list_comp]
    selected_countries_comp = st.multiselect("Select countries to compare:", options=all_countries_list_comp, default=valid_defaults_comp, key="country_comparison")
    if selected_countries_comp:
        trend_df = df_long[df_long['Name'].isin(selected_countries_comp)]
        trend_df = trend_df[trend_df['Year'] != 2019]
        trend_pivot = trend_df.pivot_table(index='Year', columns='Name', values='Posts')
        st.line_chart(trend_pivot, height=500)