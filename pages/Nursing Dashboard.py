import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import altair as alt
import plotly.express as px
from datetime import datetime
import os

st.set_page_config(page_title="Nurse Call Dashboard II",page_icon=":bar_chart:",)

# Load and display an image from a local file path
image_path = "medkicklogo.png"
st.image(image_path, caption="", use_column_width=True)

# Custom color palette using different shades of #74E5CB
custom_color = "#74E5CB"
lighter_color = "#BDF2E9"
darker_color = "#2FA48E"

# Set the page background color and font style
st.markdown(
    f"""
    <style>
    .reportview-container {{
        background-color: {lighter_color};
        font-family: Arial, sans-serif;
    }}
    </style>
    """,
    unsafe_allow_html=True
)

# Specify the folder where your Excel files are located
folder_path = 'Sasi'

# Initialize an empty list to store DataFrames
dfs = []

# Loop through all files in the current directory
for filename in os.listdir(folder_path):
    if filename.endswith(('.xlsx', '.xls', '.csv')):
        file_path = os.path.join(folder_path, filename)
        if filename.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(file_path)
        elif filename.endswith('.csv'):
            df = pd.read_csv(file_path)
        dfs.append(df)

# Concatenate or merge the DataFrames into one

data = pd.concat(dfs, ignore_index=True)

# Set the title and header with custom text color and background
st.title("Call Data Dashboard")

# Create a sidebar for filtering with a custom background and rounded corners
st.sidebar.title("Filter Data")

# Create a "Select All" option for the caller filter
select_all_caller = st.sidebar.checkbox("Select All Callers")

# Create a "Select All" option for the callee filter
select_all_callee = st.sidebar.checkbox("Select All Callees")

# Create a "Select All" option for the direction filter
select_all_direction = st.sidebar.checkbox("Select All Directions")

# Populate the caller filter based on DataFrame unique values
if select_all_caller:
    selected_caller = data['From'].unique()
else:
    selected_caller = st.sidebar.multiselect("Select Caller:", data['From'].unique())

# Populate the callee filter based on DataFrame unique values
if select_all_callee:
    selected_callee = data['To'].unique()
else:
    selected_callee = st.sidebar.multiselect("Select Callee:", data['To'].unique())

# Populate the direction filter based on DataFrame unique values
if select_all_direction:
    selected_direction = data['Direction'].unique()
else:
    selected_direction = st.sidebar.multiselect("Select Direction:", data['Direction'].unique())

# Add a date filter to the sidebar
start_date = st.sidebar.date_input("Start Date", datetime(2023, 1, 1))
end_date = st.sidebar.date_input("End Date", datetime(2023, 12, 31))

st.sidebar.markdown(
    f'<div style="background-color:{custom_color};padding:10px;border-radius:10px;">'
    f'<p style="color:white;">Filter Data</p></div>',
    unsafe_allow_html=True
)

# Convert user input dates to the desired format
start_date_str = start_date.strftime("%Y-%m-%dT%H:%M:%S.%fZ")
end_date_str = end_date.strftime("%Y-%m-%dT%H:%M:%S.%fZ")

# Filter data based on user selections
filtered_data = data[(data['From'].isin(selected_caller)) &
                     (data['To'].isin(selected_callee)) &
                     (data['Direction'].isin(selected_direction)) &
                     (data['Start Time'] >= start_date_str) &
                     (data['End Time'] <= end_date_str)]

# Filter out data with negative duration (erroneous)
filtered_data = filtered_data[filtered_data['Duration'] >= 0]

# Recalculate duration using Start Time and End Time
filtered_data['Start Time'] = pd.to_datetime(filtered_data['Start Time'], format='%Y-%m-%dT%H:%M:%S.%fZ', errors='coerce')
filtered_data['End Time'] = pd.to_datetime(filtered_data['End Time'], format='%Y-%m-%dT%H:%M:%S.%fZ', errors='coerce')
filtered_data['Duration'] = (filtered_data['End Time'] - filtered_data['Start Time']).dt.total_seconds()

# Calculate call volume and durations
total_call_volume = len(filtered_data)
inbound_volume = (filtered_data['Direction'] == 'INBOUND').sum()
outbound_volume = (filtered_data['Direction'] == 'OUTBOUND').sum()
total_duration = filtered_data['Duration'].sum()
inbound_duration = filtered_data.loc[filtered_data['Direction'] == 'INBOUND', 'Duration'].sum()
outbound_duration = filtered_data.loc[filtered_data['Direction'] == 'OUTBOUND', 'Duration'].sum()
average_duration = total_duration / total_call_volume if total_call_volume > 0 else 0

# Display call volume and duration metrics at the top of the dashboard
st.write("## Call Metrics")
st.markdown(
    f'<div style="background-color:{custom_color};padding:10px;border-radius:10px;">'
    f'<p style="color:white;">Total Call Volume: {total_call_volume}</p>'
    f'<p style="color:white;">Inbound Volume: {inbound_volume}</p>'
    f'<p style="color:white;">Outbound Volume: {outbound_volume}</p>'
    f'<p style="color:white;">Total Duration: {total_duration:.2f} seconds</p>'
    f'<p style="color:white;">Inbound Duration: {inbound_duration:.2f} seconds</p>'
    f'<p style="color:white;">Outbound Duration: {outbound_duration:.2f} seconds</p>'
    f'<p style="color:white;">Average Duration: {average_duration:.2f} seconds</p>'
    f'</div>',
    unsafe_allow_html=True
)

# Display filtered data with a lighter background and custom styling
st.subheader("Filtered Data")
st.dataframe(filtered_data.style.set_properties(**{'background-color': lighter_color, 'color': 'black'}))

# Check for and handle null values with custom styling
if filtered_data.isnull().sum().any():
    st.subheader("Null Values in Filtered Data")
    st.dataframe(filtered_data.isnull().sum().to_frame(name='Null Count').style.set_properties(
        **{'background-color': lighter_color, 'color': 'black'}))

# Create charts

# Explanation for the Call Duration Distribution
st.write("#### Call Duration Distribution")
st.write("Distribution of call durations for selected nurse and direction.")
st.markdown(
    f'<div style="background-color:{lighter_color};padding:10px;border-radius:10px;border: 1px solid {custom_color};">'
    f'<p style="color:black;">Call Duration Distribution</p></div>',
    unsafe_allow_html=True
)

# Create a Matplotlib figure and axis with custom colors
fig, ax = plt.subplots()
sns.histplot(filtered_data['Duration'], bins=20, kde=True, color=custom_color, edgecolor='k')
ax.set_xlabel("Duration (seconds)")  # Label for the x-axis
ax.set_ylabel("Frequency")           # Label for the y-axis
st.pyplot(fig)  # Display the Matplotlib figure in Streamlit

# Explanation for the Disposition Counts
st.write("#### Disposition Counts")
st.write("Counts of different dispositions for selected nurse and direction.")
st.markdown(
    f'<div style="background-color:{lighter_color};padding:10px;border-radius:10px;border: 1px solid {custom_color};">'
    f'<p style="color:black;">Disposition Counts</p></div>',
    unsafe_allow_html=True
)

disposition_counts = filtered_data['Disposition'].value_counts()
st.bar_chart(disposition_counts, use_container_width=True)  # Use container width for better layout

# Explanation for the Time Series Analysis
st.write("### Time Series Analysis")
st.write("Call duration over time for selected nurse and direction.")
st.markdown(
    f'<div style="background-color:{lighter_color};padding:10px;border-radius:10px;border: 1px solid {custom_color};">'
    f'<p style="color:black;">Time Series Analysis</p></div>',
    unsafe_allow_html=True
)

# Create a time series line chart with Altair and add labels
line_chart = alt.Chart(filtered_data.reset_index()).mark_line().encode(
    x='Start Time:T',  # Time-based x-axis
    y=alt.Y('Duration:Q', title='Call Duration (seconds)'),  # Label for y-axis
    tooltip=['Start Time:T', alt.Tooltip('Duration:Q', title='Call Duration (seconds)')]
).properties(width=700, height=400)  # Adjust chart dimensions

st.altair_chart(line_chart, use_container_width=True)  # Use container width for better layout

# Explanation for the Distribution of Call Directions
st.write("#### Distribution of Call Directions")
st.write("Distribution of call directions for selected nurse and direction.")
st.markdown(
    f'<div style="background-color:{lighter_color};padding:10px;border-radius:10px;border: 1px solid {custom_color};">'
    f'<p style="color:black;">Distribution of Call Directions</p></div>',
    unsafe_allow_html=True
)

# Calculate the percentage of inbound and outbound calls
direction_counts = filtered_data['Direction'].value_counts()
total_calls = len(filtered_data)
inbound_percentage = (direction_counts.get("INBOUND", 0) / total_calls) * 100
outbound_percentage = (direction_counts.get("OUTBOUND", 0) / total_calls) * 100

# Create a custom Pie chart
fig_pie = px.pie(
    names=["INBOUND", "OUTBOUND"],
    values=[inbound_percentage, outbound_percentage],
    title='Call Directions'
)
fig_pie.update_traces(marker=dict(colors=[custom_color, darker_color]))
st.plotly_chart(fig_pie)

# Explanation for the Scatter Plot
st.write("#### Scatter Plot: Call Duration vs. Disposition")
st.write("Relationship between call duration and disposition.")
st.markdown(
    f'<div style="background-color:{lighter_color};padding:10px;border-radius:10px;border: 1px solid {custom_color};">'
    f'<p style="color:black;">Scatter Plot: Call Duration vs. Disposition</p></div>',
    unsafe_allow_html=True
)

scatter_chart = alt.Chart(filtered_data).mark_circle(size=60).encode(
    x='Duration:Q',
    y='Disposition:N',
    color=alt.Color('Disposition:N', scale=alt.Scale(domain=list(disposition_counts.index), range=[custom_color, darker_color])),
    tooltip=['Duration:Q', 'Disposition:N']
).properties(width=700, height=300)

st.altair_chart(scatter_chart, use_container_width=True)
