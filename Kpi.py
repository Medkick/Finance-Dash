#!/usr/bin/env python
# coding: utf-8

# In[25]:


import pandas as pd
#pip install plotly_express==0.4.0 
import plotly.express as px
import streamlit as st
import matplotlib.pyplot as plt
import numpy as np
from datetime import datetime
import calendar

st.set_page_config(page_title="Financial_Dashboard",page_icon=":bar_chart:",)

# Custom CSS styles
custom_css = """
<style>
h1 {
    color: #FF9800;
    font-size: 36px;
    margin-top: 40px;
    margin-bottom: 30px;
}

h2 {
    color: #607D8B;
    font-size: 24px;
    margin-top: 30px;
    margin-bottom: 20px;
}

table.dataframe {
    font-size: 14px;
    border-collapse: collapse;
    margin: 20px 0;
}

table.dataframe th {
    background-color: #FF9800;
    color: white;
}

table.dataframe td, table.dataframe th {
    border: 1px solid #FF9800;
    padding: 8px;
}

.stPlotlyChart {
    max-width: 800px;
    margin-top: 30px;
    margin-bottom: 50px;
}
</style>
"""
# Apply custom CSS styles
st.markdown(custom_css, unsafe_allow_html=True)

# In[26]:


#usecols="A:E",
#Sheet 2021
df1=pd.read_excel(io='Med2.xlsx',engine="openpyxl",sheet_name="2021",skiprows=[],usecols="A:M",nrows=100)
new_columns1 = ['KPI'] + [f'2021_{col}' for col in df1.columns[1:]]
df1 = df1.rename(columns=dict(zip(df1.columns, new_columns1)))


# In[27]:


#Sheet 2022
df2=pd.read_excel(io='Med2.xlsx',engine="openpyxl",sheet_name="2022",skiprows=[],usecols="A:M",nrows=100)
new_columns2 = ['KPI'] + [f'2022_{col}' for col in df2.columns[1:]]
df2 = df2.rename(columns=dict(zip(df2.columns, new_columns2)))


# In[28]:


#Sheet 2023
df3=pd.read_excel(io='Med2.xlsx',engine="openpyxl",sheet_name="2023",skiprows=[],usecols="A:M",nrows=100)
new_columns3 = ['KPI'] + [f'2023_{col}' for col in df3.columns[1:]]
df3 = df3.rename(columns=dict(zip(df3.columns, new_columns3)))


# In[29]:


#Merging
df4 = pd.merge(df1, df2, on=['KPI'], how ='outer')
df = pd.merge(df4, df3, on=['KPI'], how ='outer')

# Extract unique years from column names
years = list(set([col.split('_')[0] for col in df.columns[1:]]))
years.sort()

# In[33]:

# Extract unique years from column names
years = list(set([col.split('_')[0] for col in df.columns[1:]]))
years.sort()

# Streamlit app
st.title('KPI Analysis')

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
start_year = st.sidebar.number_input('Start Year', value=start_year, step=1)
start_month = st.sidebar.number_input('Start Month', value=start_month, min_value=1, max_value=12, step=1)
end_year = st.sidebar.number_input('End Year', value=end_year, step=1)
end_month = st.sidebar.number_input('End Month', value=end_month, min_value=1, max_value=12, step=1)

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
    'Select Month(s)', 
    options=get_month_year_options(start_year, start_month, end_year, end_month), 
    default=get_month_year_options(start_year, start_month, end_year, end_month)
    )


# Rerun the app if the selected years or months change
if st.sidebar.button('Apply'):
    st.experimental_rerun()

# Filter the data based on selected years and months
selected_cols = ['KPI'] + [col for col in df.columns if any(col.startswith(month) for month in selected_months)]
df6 = df[selected_cols]

# Check if the filtered DataFrame is empty (contains only null values)
if df6.drop('KPI', axis=1).isnull().all().all():
    st.write("No data available for the selected range.")
else:
    # Calculate combined metrics across all KPIs
    combined_total = df6.iloc[:, 1:].sum().sum()
    # Calculate highest and lowest costs for a month
    monthly_sums = df6.iloc[:, 1:].sum(axis=0)
    monthly_sums_nonzero = monthly_sums[monthly_sums > 0]  # Filter out 0 value months
    combined_average = monthly_sums_nonzero.mean()
    combined_median = monthly_sums_nonzero.median()
    combined_highest = monthly_sums.max()
    combined_lowest = monthly_sums_nonzero.min()  # Use the minimum value after excluding 0 value months
    # Find the months associated with the highest and lowest costs
    combined_highest_month = df6.columns[1:][monthly_sums.argmax()]
    combined_lowest_month = df6.columns[1:][monthly_sums==combined_lowest]  # Find the month for the lowest value after excluding 0 value months
    # Convert combined_lowest_month from Index to string
    combined_lowest_month_str = combined_lowest_month.item()  # or combined_lowest_month.tolist()[0]

    # Format the numeric data in the DataFrame as dollar currency
    #dollar_format = '${:,.2f}'

    # Display combined metrics with large font
    #st.markdown(f'### Total Expenditure: {dollar_format.format(combined_total)}', unsafe_allow_html=True)
    #st.markdown(f'### Average Expenditure: {dollar_format.format(combined_average)}', unsafe_allow_html=True)
    #st.markdown(f'### Median Expenditure: {dollar_format.format(combined_median)}', unsafe_allow_html=True)
    #st.markdown(f'### Highest Expenditure ({combined_highest_month}): {dollar_format.format(combined_highest)}', unsafe_allow_html=True)
    #st.markdown(f'### Lowest Expenditure ({combined_lowest_month_str}): {dollar_format.format(combined_lowest)}', unsafe_allow_html=True)
    #st.markdown('---')

    # Sum the values in each column (except the first column)
    sum_values = df6.iloc[:, 1:].sum()

    # Create a new data frame with a single row and assign the calculated values
    new_data = {
        'KPI': ['Total Expenditure for the Month'],
        **sum_values.to_dict()  # Using dictionary unpacking to merge the 'x' and sum_values dictionaries
    }

    df7 = pd.DataFrame(new_data)

    # Display the filtered DataFrame as a table
    #st.subheader('Data')
    #st.dataframe(df7, height=3, hide_index=True)

    # Reshape data for Plotly Express
    reshaped_data = pd.melt(df7, id_vars='KPI', var_name='Month', value_name='Total Expenditure for the Month')

    # Prepare the data for Plotly Express
    fig1 = px.bar(
        data_frame=reshaped_data,
        x='Month',
        y='Total Expenditure for the Month',
        color='Month',  # Assign color based on the month
        labels={'x': 'Month</b>', 'y': 'Total Expenditure for the Month</b>'},  # Set the x-axis and y-axis labels
        title=f'Bar Chart for Total Expenditure for the Month',
    )

    # Add value labels on top of the bars and Round the values to whole numbers
    fig1.update_traces(texttemplate='<b>%{y:.0f}</b>', textposition='outside')

    # Make the bar chart title big
    fig1.update_layout(
        title_font=dict(size=30),
    )

    # Adjust the size of the chart
    fig1.update_layout(
        autosize=False,
        width=1100,
        height=500,
        plot_bgcolor='rgba(0,0,0,0)',
       paper_bgcolor='rgba(0,0,0,0)',
    )

    # Set the font color to dark
    fig1.update_layout(
        font=dict(color='black')
    )

    # Update x and y labels to bold
    fig1.update_layout(
        xaxis_title_font=dict(family='Arial', size=14, color='black'),
        yaxis_title_font=dict(family='Arial', size=14, color='black'),
        xaxis_tickfont=dict(family='Arial', size=12, color='black'),
        yaxis_tickfont=dict(family='Arial', size=12, color='black'),
    )

    # Display the bar chart
    #st.plotly_chart(fig1)

 
  


    # Iterate through each KPI
    for kpi_name in df['KPI']:
        st.markdown(f'# KPI: {kpi_name}')

        # Filtered DataFrame for the specific KPI
        filtered_df = df[df['KPI'] == kpi_name]

        # Filter the data based on selected years and months
        selected_cols = ['KPI'] + [col for col in df.columns if any(col.startswith(month) for month in selected_months)]
        filtered_data = filtered_df[selected_cols]

        # Format the numeric data in the DataFrame as dollar currency
        dollar_format = '{:,.2f}'
# Handle the first two KPIs with percentages
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

        # Convert lowest_month to a comma-separated string
        lowest_month_str = ', '.join(lowest_month.tolist())

        # Define the format for bold text
        bold_font = "<b>{}</b>"
    
        st.write(bold_font.format('Total Value: ' + dollar_format.format(total)), unsafe_allow_html=True)
        st.write(bold_font.format('Average Value: ' + dollar_format.format(average)), unsafe_allow_html=True)
        st.write(bold_font.format('Median Value: ' + dollar_format.format(median)), unsafe_allow_html=True)
        st.write(bold_font.format('Highest Value ({0}): '.format(highest_month) + dollar_format.format(highest_value)), unsafe_allow_html=True)
        st.write(bold_font.format(f'Lowest Value ({lowest_month_str}): ' + dollar_format.format(lowest_value)), unsafe_allow_html=True)
    
        # Display the filtered DataFrame as a table
        st.subheader('Data')
        st.dataframe(filtered_data, height=3, hide_index=True)

        # Reshape data for Plotly Express
        reshaped_data = pd.melt(filtered_data, id_vars='KPI', var_name='Month', value_name='Cost')

        # Prepare the data for Plotly Express
        fig = px.bar(
            data_frame=reshaped_data,
            x='Month',
            y='Cost',
            color='Month',  # Assign color based on the month
            labels={'x': 'Month</b>', 'y': 'Cost</b>'},  # Set the x-axis and y-axis labels
            title=f'Bar Chart for {kpi_name}',
        )
        
        # Add value labels on top of the bars and Round the values to whole numbers
        fig.update_traces(texttemplate='<b>%{y:.0f}</b>', textposition='outside')

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
        st.markdown('<hr style="margin-top: 50px; margin-bottom: 50px; border-width: 0; border-top: 2px solid #FF9800">', unsafe_allow_html=True)


