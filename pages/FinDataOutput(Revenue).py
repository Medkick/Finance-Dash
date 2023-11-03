from pathlib import Path  # Python Standard Library
import pandas as pd  # pip install pandas openpyxl
from pathlib import Path  # Python Standard Library
import plotly.express as px
import streamlit as st
import matplotlib.pyplot as plt
import plotly.figure_factory as ff
#import openpyxl
#from openpyxl import Workbook,load_workbook





st.set_page_config(page_title = "Med-Kick Finance Dashboard",
                   layout ="wide"
                  )

@st.cache_data()
def rev_dash():
    data = FinDataOutput(Revenue).csv
    df = pd.read_csv(data)
    df.head(5)
    return df 
   

df = rev_dash()

df.rename(columns={'Revenue Stream by Service': 'Service'}, inplace=True)

# sidebars#
st.sidebar.header("Period of Operation" )
Year = st.sidebar.multiselect('Year',
                                options= df['Year'].unique(),
                                default=df['Year'].unique()
                                )


Service = st.sidebar.multiselect('Service',
                                options= df['Service'].unique(),
                                default=df['Service'].unique()
                                )


df_filter= df.query(
    "Year ==@Year & Service ==@Service"
    )

st.markdown('##')
st.sidebar.title("Helpful Definitions")
st.sidebar.write('99490- CCM 1st 20 min')
st.sidebar.write('99439- CCM 2nd & 3rd 20 min')
st.sidebar.write('99453- RPM setup')
st.sidebar.write('99454- RPM data')
st.sidebar.write('99457- RPM 1st 20 min')
st.sidebar.write('99458- RPM 2nd & 3rd 20 min')
st.sidebar.write('99426 – PCM 30mins intervention')
st.sidebar.write('99427 – PCM additional 30mins')
st.sidebar.write('99484 – BHI 20mins monthly intervention')
st.sidebar.write('G0511 – CCM and BHI for Federally qualified health centers')



#--Mainpage--#

st.title('Revenue Streams Dashboard')
st.markdown('##')



#--KPI'S--#
no_of_services = df_filter['Service'].nunique()
total_rev = df_filter[['January','February','March','April','May','June','July','August','September','October','November','December']].sum()
ttotal_rev= total_rev.sum()
round_total_rev= round(ttotal_rev,2)
rev_by_month = df_filter.groupby(by=['Service']).sum()[['January','February','March','April','May','June','July','August','September','October','November','December']]
trans_rev_by_month = rev_by_month.transpose()
trans_rev_by_month_sum = trans_rev_by_month.sum()
min_serv = min(trans_rev_by_month_sum)
rounded_min_serv= round(min_serv,2)
max_serv = max(trans_rev_by_month_sum)
rounded_max_serv= round(max_serv,2)

left_column,middle_column,right_column = st.columns(3)
with left_column:
    st.subheader('Number of services')
    st.subheader(no_of_services)
    
with right_column:
    st.subheader('Min Revenue service')
    st.subheader(f"US $ {rounded_min_serv:,}")
    
with middle_column:
    st.subheader('Total Revenue')
    st.subheader(f"US $ {round_total_rev:,}")
st.markdown('---')
with left_column:
    st.subheader('Max Revenue service')
    st.subheader(f"US $ {rounded_max_serv:,}")

st.markdown('---')

rev_by_serv = (
    df_filter[['January','February','March','April','May','June','July','August','September','October','November','December']].sum()
    )


rev_by_serv_month = (
    df_filter.groupby(by=['Service']).sum()[['January','February','March','April','May','June','July','August','September','October','November','December']]
   )

rev_by_serv_month_transpose = rev_by_serv_month.transpose()

rev_by_serv_month_transpose_sum = rev_by_serv_month_transpose.sum()

img_one_revenue= px.bar(
    rev_by_serv,
    x =rev_by_serv.index,
    y = rev_by_serv.values ,
    #orientation = 'h',
    title = "<b>Revenue distribution by Month</b>",
    template = 'plotly_white',
    )

img_three_revenue =px.pie(
    rev_by_serv_month_transpose_sum ,
    names=rev_by_serv_month_transpose_sum .index,
    values=rev_by_serv_month_transpose_sum.values,
    title ="<b>Services Proportion of Revenue</b>"
    )

img_two_revenue= px.line(
    rev_by_serv_month_transpose,
    #color=('green'),
    title = "<b>Revenue Trend of Service</b>",
    template = 'plotly_white',
    )
 


left_column,right_column = st.columns(2)
left_column.plotly_chart(img_two_revenue, use_container_width=True)
right_column.plotly_chart(img_three_revenue, use_container_width=True)

st.markdown("##")
st.plotly_chart(img_one_revenue)



hide_st_style = """
                <style>
                #MainMenu {visibility:hidden;}
                footer {visibility:hidden;}
                header {visibility:hidden;}
                </style>
                """
st.markdown(hide_st_style, unsafe_allow_html=True)
#st.plotly_chart(img_one_revenue)
#st.plotly_chart(img_two_revenue)
#st.plotly_chart(img_three_revenue)
