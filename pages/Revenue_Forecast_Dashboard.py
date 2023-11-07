import pandas as pd
#pip install plotly_express==0.4.0 
import plotly.express as px
import streamlit as st
import matplotlib.pyplot as plt
import numpy as np
from datetime import datetime
import calendar

st.set_page_config(page_title="Revenue_Forecast",page_icon=":bar_chart:",)

excel_file_path = 'finance.xlsx'
excel_file = pd.ExcelFile(excel_file_path)
sheet_names = excel_file.sheet_names

dfs = []

for year_sheet in sheet_names:
    if year_sheet.isdigit():  # Check if the sheet name is a year (numeric)
        df1 = pd.read_excel(io=excel_file, sheet_name=year_sheet, skiprows=[], usecols="A:M", nrows=100)
        new_columns = ['KPI'] + [f'{year_sheet}_{col}' for col in df1.columns[1:]]
        df1 = df1.rename(columns=dict(zip(df1.columns, new_columns)))
        dfs.append(df1)

# Merging
df = dfs[0]  # Start with the first DataFrame
for df1 in dfs[1:]:
    df = pd.merge(df, df1, on=['KPI'], how='outer')

# Extract unique years from column names
years = list(set([col.split('_')[0] for col in df.columns[1:]]))
years.sort()

# Load and display an image from a local file path
image_path = "medkicklogo.png"
st.image(image_path, caption="", use_column_width=True)

# Streamlit app
st.title(':green[Revenue Forecast Dashboard]  :chart_with_upwards_trend:')

# Extract the earliest and latest dates from the DataFrame
earliest_date = min(df.columns[1:], key=lambda x: datetime.strptime(x, '%Y_%B'))
latest_date = max(df.columns[1:], key=lambda x: datetime.strptime(x, '%Y_%B'))

# Extract the year and month from the earliest and latest dates
start_year, start_month = earliest_date.split('_')
end_year, end_month = latest_date.split('_')

# Convert the year and month to integers
start_year = int(start_year)
start_month = list(calendar.month_name).index(start_month)
end_year = int(end_year)
end_month = list(calendar.month_name).index(end_month)

# Select years and months in the sidebar with default selection
start_year = st.sidebar.number_input('Start Year :date:', value=start_year, step=1)
start_month = st.sidebar.number_input('Start Month :date:', value=start_month, min_value=1, max_value=12, step=1)
end_year = st.sidebar.number_input('End Year :date:', value=end_year, step=1)
end_month = st.sidebar.number_input('End Month :date:', value=end_month, min_value=1, max_value=12, step=1)

# Define callback function to get month options based on the selected year and month
def get_month_year_options(start_year, start_month, end_year, end_month):
    selected_months = []
    for year in range(start_year, end_year + 1):
        if year == start_year:
            start_m = start_month
        else:
            start_m = 1
        if year == end_year:
            end_m = end_month
        else:
            end_m = 12
        for month in range(start_m, end_m + 1):
            selected_months.append(f"{year}_{calendar.month_name[month]}")

    return selected_months

# Select months in the sidebar with default selection based on the selected years
selected_months = st.sidebar.multiselect(
    'Selected Month(s) :calendar:', 
    options=get_month_year_options(start_year, start_month, end_year, end_month), 
    default=get_month_year_options(start_year, start_month, end_year, end_month)
    )


# Rerun the app if the selected years or months change
if st.sidebar.button('Apply'):
    st.experimental_rerun()

# Filter the data based on selected years and months
selected_cols = ['KPI'] + [col for col in df.columns if any(col.startswith(month) for month in selected_months)]
dff = df[selected_cols]

# Check if the filtered DataFrame is empty (contains only null values)
if dff.drop('KPI', axis=1).isnull().all().all():
    st.write("No data available for the selected range.")
else:
    
    # Format the numeric data in the DataFrame as dollar currency
    dollar_format = '${:,.2f}'

    # Iterate through each KPI
    for kpi_name in df['KPI']:
        if kpi_name in ['Total Revenue','Revenue Growth Rate','Revenue forecast']:
            markdown_text = f'<h1 style="color: green;">KPI: {kpi_name}</h1>'
            st.markdown(markdown_text, unsafe_allow_html=True)
            # st.markdown(f'# KPI: {kpi_name}')

            # Filtered DataFrame for the specific KPI
            filtered_df = df[df['KPI'] == kpi_name]

            # Filter the data based on selected years and months
            selected_cols = ['KPI'] + [col for col in df.columns if any(col.startswith(month) for month in selected_months)]
            filtered_data = filtered_df[selected_cols]
    
            # Format the numeric data in the DataFrame as dollar currency
            dollar_format = '${:,.2f}'
    
            # Calculate important metrics
            total = filtered_data.iloc[:, 1:].sum().sum()
            average = filtered_data.iloc[:, 1:].mean().mean()
            median = filtered_data.iloc[:, 1:].median().median()
            # Calculate highest and lowest costs for a month (excluding 0 values)
            monthly_sums = filtered_data.iloc[:, 1:].sum(axis=0)
            monthly_sums_nonzero = monthly_sums[monthly_sums > 0]  # Filter out 0 value months
            highest_value = monthly_sums.max()
            lowest_value = monthly_sums_nonzero.min()  # Use the minimum value after excluding 0 value months
    
            # Find the months associated with the highest and lowest costs
            highest_month = filtered_data.columns[1:][monthly_sums.argmax()]
            lowest_month = filtered_data.columns[1:][monthly_sums==lowest_value]  # Find the month for the lowest value after excluding 0 value months
    
            # Convert lowest_month from Index to string
            lowest_month_str = lowest_month.item()  # or lowest_month.tolist()[0]
    
            # Define the format for bold text
            bold_font = "<b>{}</b>"
        
            if kpi_name in ['Revenue Growth Rate']:
                percentage_format = '{:.2%}'
                st.write(bold_font.format('Total : ' + percentage_format.format(total/18)), unsafe_allow_html=True)
                st.write(bold_font.format('Average : ' + percentage_format.format(average)), unsafe_allow_html=True)
                st.write(bold_font.format('Median : ' + percentage_format.format(median)), unsafe_allow_html=True)
                st.write(bold_font.format('Highest ({0}): '.format(highest_month) + percentage_format.format(highest_value)), unsafe_allow_html=True)
                st.write(bold_font.format('Lowest ({0}): '.format(lowest_month_str) + percentage_format.format(lowest_value)), unsafe_allow_html=True)
            else:
                dollar_format = '${:,.2f}'
                st.write(bold_font.format('Total : ' + dollar_format.format(total)), unsafe_allow_html=True)
                st.write(bold_font.format('Average : ' + dollar_format.format(average)), unsafe_allow_html=True)
                st.write(bold_font.format('Median : ' + dollar_format.format(median)), unsafe_allow_html=True)
                st.write(bold_font.format('Highest ({0}): '.format(highest_month) + dollar_format.format(highest_value)), unsafe_allow_html=True)
                st.write(bold_font.format('Lowest ({0}): '.format(lowest_month_str) + dollar_format.format(lowest_value)), unsafe_allow_html=True)
    
            # Display the filtered DataFrame as a table
            st.subheader('Data')
            st.dataframe(filtered_data, height=3, hide_index=True)

            # Reshape data for Plotly Express
            reshaped_data = pd.melt(filtered_data, id_vars='KPI', var_name='Month', value_name='Value')
    
            # Prepare the data for Plotly Express
            fig = px.bar(
                data_frame=reshaped_data,
                x='Month',
                y='Value',
                color='Month',  # Assign color based on the month
                color_discrete_sequence=['#d5f7ef','#c7f4ea','#b9f2e5','#abefdf','#9decda','#8fead5','#74e5cb','#68ceb6','#5cb7a2','#51a08e','#458979','#3a7265','#2e5b51','#22443c'],  # Set the desired color sequence
                labels={'x': 'Month</b>', 'y': 'Cost</b>'},  # Set the x-axis and y-axis labels
                title=f'Bar Chart for {kpi_name}',
            )

            if kpi_name in ['Revenue Growth Rate']:
                # Add value labels on top of the bars as percentage strings
                fig.update_traces(texttemplate='%{y:.2%}', textposition='outside')
            else:
                # Add value labels on top of the bars and Round the values to whole numbers
                fig.update_traces(texttemplate='<b>$%{y:.0f}</b>', textposition='outside')
    
            # Make the bar chart title big
            fig.update_layout(
                title_font=dict(size=30),
            )

            # Adjust the size of the chart
            fig.update_layout(
                autosize=False,
                width=1100,
                height=500,
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
            )

            # Set the font color to dark
            fig.update_layout(
                font=dict(color='black')
            )

            # Update x and y labels to bold
            fig.update_layout(
                xaxis_title_font=dict(family='Arial', size=14, color='black'),
                yaxis_title_font=dict(family='Arial', size=14, color='black'),
                xaxis_tickfont=dict(family='Arial', size=12, color='black'),
                yaxis_tickfont=dict(family='Arial', size=12, color='black'),
            )
    
            # Display the bar chart
            st.plotly_chart(fig)
    
            # Add space between KPIs
            st.markdown('<hr style="margin-top: 50px; margin-bottom: 50px; border-width: 0; border-top: 2px solid #74e5cb">', unsafe_allow_html=True)


