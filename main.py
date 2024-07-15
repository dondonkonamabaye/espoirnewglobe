import time
from tracemalloc import start

import pandas as pd
import numpy as np

import datetime as dt

from matplotlib.figure import Figure
import matplotlib.pyplot as plt
from matplotlib import cm
from matplotlib.backends.backend_agg import FigureCanvas  # not needed for mpl >= 3.1

import mplleaflet
from IPython.display import IFrame

import base64
from PIL import Image
import io

import hvplot.pandas

import requests # Pour effectuer la requête
import pandas as pd # Pour manipuler les données
import datetime as dt

import param
import panel as pn

import holoviews as hv
import holoviews.plotting.bokeh

import seaborn as sns

import altair as alt


df_espoir = pd.read_excel("espoir.xlsx")

df_espoir['Date Submitted'] = pd.to_datetime(df_espoir['Date Submitted'])
df_espoir['Year'] = df_espoir['Date Submitted'].dt.year
df_espoir['Month'] = df_espoir['Date Submitted'].dt.month
df_espoir['Day'] = df_espoir['Date Submitted'].dt.day

df_espoir['Year'] = df_espoir['Year'].astype("str")


df_espoir['Resolved date'] = pd.to_datetime(df_espoir['Last Updated'])
df_espoir['Year Resolved'] = df_espoir['Resolved date'].dt.year
df_espoir['Month Resolved'] = df_espoir['Resolved date'].dt.month
df_espoir['Day Resolved'] = df_espoir['Resolved date'].dt.day


df_espoir['Last Updated'] = pd.to_datetime(df_espoir['Last Updated'])
df_espoir['Year Updated'] = df_espoir['Last Updated'].dt.year
df_espoir['Month Updated'] = df_espoir['Last Updated'].dt.month
df_espoir['Day Updated'] = df_espoir['Last Updated'].dt.day

df_espoir['Year'] = df_espoir['Year'].astype("str")

import datetime

full_month_name = []

for date_month_name in df_espoir['Month']:
    datetime_object = datetime.datetime.strptime(str(date_month_name), "%m")
    
    full_month_name.append(datetime_object.strftime("%B"))
                           
df_espoir['Month Name'] = full_month_name

BRAND_COLOR = "teal"
BRAND_TEXT_ON_COLOR = "white"

CARD_STYLE = {
  "box-shadow": "rgba(50, 50, 93, 0.25) 0px 6px 12px -2px, rgba(0, 0, 0, 0.3) 0px 3px 7px -3px",
  "padding": "4px",
  "border-radius": "1px"
}

header = pn.Row(
    pn.pane.Markdown(
        "# Wind Turbine Report", styles={"color": BRAND_TEXT_ON_COLOR}, margin=(5, 20)
    ),
    styles={"background": BRAND_COLOR},
)



df = df_espoir
count = len(df)
min_capacity = df['Month'].min()
max_capacity = df['Month'].max()
max_capacity = int(max_capacity)
avg_rotor_diameter = df['Month'].mean()
avg_rotor_diameter = int(avg_rotor_diameter)
std_rotor_diameter = df['Month'].std()
std_rotor_diameter = int(std_rotor_diameter)

count_day = len(df)
min_day = df['Day'].min()
max_day = df['Day'].max()
max_day = int(max_day)
avg_day = df['Day'].mean()
avg_day = int(avg_day)
std_day = df['Day'].std()
std_day = int(std_day)

top_manufacturers = (
    df.groupby("Status")
)
df_espoir_example = df_espoir.dropna().head(5).iloc[:5,:13].reset_index(drop=True)

#########################################################################################################################

df_espoir_it_operations_2018_closed_open = df_espoir.loc[(df_espoir['Date Submitted'] == '01-Dec-2018') |
                                                                       ((df_espoir['Status'] != 'Closed') & (df_espoir['Status'] != 'Open'))]

df_espoir_it_operations_2018_closed_open['Date Submitted'] = pd.to_datetime(df_espoir_it_operations_2018_closed_open['Date Submitted'])

df_espoir_it_operations_2018_closed_open['Day Submitted'] = df_espoir_it_operations_2018_closed_open['Date Submitted'].dt.day
df_espoir_it_operations_2018_closed_open['Day Solved'] = df_espoir_it_operations_2018_closed_open['Day Resolved'] - df_espoir_it_operations_2018_closed_open['Day Submitted']
df_espoir_it_operations_2018_closed_open['Day Solved'] = df_espoir_it_operations_2018_closed_open['Day Resolved']

espoir_avg_day = df_espoir_it_operations_2018_closed_open['Day Solved'].mean()
espoir_avg_day = int(espoir_avg_day)

max_day_to_resolved = f'The average time of all tickets that are neither resolved nor closed is {int(espoir_avg_day)} days'


#########################################################################################################################

# df = df[df['Day Number'].isin(top_manufacturers)]
fig = (
    alt.Chart(
        df,
        title="Capacity by Manufacturer",
    )
    .mark_circle(size=8)
    .encode(
        y="Month:N",
        x="Month:Q",
        yOffset="jitter:Q",
        color=alt.Color("Month:N").legend(None),
        tooltip=["Month", "Month"],
    )
    .transform_calculate(jitter="sqrt(-2*log(random()))*cos(2*PI*random())")
    .properties(
        height=400,
        width="container",
    )
)


#########################################################################################################################

indicators_day = pn.FlexBox(

    pn.indicators.Number(
        value=max_day,
        name="Max Day",
        format="{value:,.1f}",
        styles=CARD_STYLE,
    ),

    pn.indicators.Number(
        value=min_day,
        name="Min Day",
        format="{value:,.0f}",
        styles=CARD_STYLE,
    ),

    pn.indicators.Number(
        value=avg_day,
        name="Avg Day",
        format="{value:,.0f}",
        styles=CARD_STYLE,
    ),

    pn.indicators.Number(
        value=std_day,
        name="Std Day",
        format="{value:,.0f}",
        styles=CARD_STYLE,
    ),
  
    pn.indicators.Number(
        value=count_day,
        name="Number of Rows",
        format="{value:,.0f}",
        styles=CARD_STYLE,
    ),
    margin=(20, 5),
)

#########################################################################################################################

indicators_month = pn.FlexBox(

    pn.indicators.Number(
        value=max_capacity,
        name="Max Month",
        format="{value:,.1f}",
        styles=CARD_STYLE,
    ),

    pn.indicators.Number(
        value=min_capacity,
        name="Min Month",
        format="{value:,.0f}",
        styles=CARD_STYLE,
    ),

    pn.indicators.Number(
        value=avg_rotor_diameter,
        name="Avg Month",
        format="{value:,.0f}",
        styles=CARD_STYLE,
    ),

    pn.indicators.Number(
        value=std_rotor_diameter,
        name="Std Month",
        format="{value:,.0f}",
        styles=CARD_STYLE,
    ),
  
    pn.indicators.Number(
        value=count,
        name="Number of Rows",
        format="{value:,.0f}",
        styles=CARD_STYLE,
    ),
    margin=(20, 5),
)

#########################################################################################################################

indicators_month_top = pn.FlexBox(

     pn.indicators.Number(
        value=count,
        name="Number of Rows",
        format="{value:,.0f}",
        styles=CARD_STYLE,
    ),

    pn.indicators.Number(
        value=max_day,
        name="Max ID of Day",
        format="{value:,.1f}",
        styles=CARD_STYLE,
    ),

    pn.indicators.Number(
        value=min_day,
        name="Min ID of Day",
        format="{value:,.0f}",
        styles=CARD_STYLE,
    ),

    pn.indicators.Number(
        value=max_capacity,
        name="Max ID of Month",
        format="{value:,.1f}",
        styles=CARD_STYLE,
    ),

    pn.indicators.Number(
        value=min_capacity,
        name="Min ID of Month",
        format="{value:,.0f}",
        styles=CARD_STYLE,
    ),

     pn.indicators.Number(
        value=espoir_avg_day,
        name="Avg Day Resolved",
        format="{value:,.1f}",
        styles=CARD_STYLE,
    ),

    margin=(20, 5),
)

#########################################################################################################################

indicators_day_top = pn.FlexBox(

    pn.indicators.Number(
        value=max_day,
        name="Max Day",
        format="{value:,.1f}",
        styles=CARD_STYLE,
    ),

    pn.indicators.Number(
        value=min_day,
        name="Min Day",
        format="{value:,.0f}",
        styles=CARD_STYLE,
    ),
    margin=(20, 5),
)

#########################################################################################################################

df_espoir = df_espoir.sort_values(by='Status', ascending=True)

unique_espoir_status = list(df_espoir['Status'].unique())
select_espoir_status_month = pn.widgets.Select(name = 'Status', options = unique_espoir_status)
select_espoir_status_day = pn.widgets.Select(name = 'Status', options = unique_espoir_status)

df_espoir = df_espoir.sort_values(by='Month Name', ascending=True)

unique_espoir_month_name = list(df_espoir['Month Name'].unique())
select_espoir_month_name = pn.widgets.Select(name = 'Month Name', options = unique_espoir_month_name)
select_espoir_month_name_status = pn.widgets.Select(name = 'Month Name', options = unique_espoir_month_name)

df_espoir = df_espoir.sort_values(by='Department', ascending=True)

unique_espoir_department = list(df_espoir['Department'].unique())
select_espoir_department = pn.widgets.Select(name = 'Department', options = unique_espoir_department)
select_espoir_department_all = pn.widgets.Select(name = 'Department', options = unique_espoir_department)
# select_espoir_month_name_status = pn.widgets.Select(name = 'Department', options = unique_espoir_month_name)

sum_day_espoir = df_espoir['Day'].sum()
max_day_espoir = df_espoir['Day'].max()
min_day_espoir = df_espoir['Day'].min()
avg_day_espoir = df_espoir['Day'].mean()
std_day_espoir = df_espoir['Day'].std()

int_slider_submitted_day = pn.widgets.IntSlider(name='Day ID', start=int(min_day_espoir), end=int(max_day_espoir), step=1, value=int(max_day_espoir))
int_slider_submitted_day_all = pn.widgets.IntSlider(name='Day ID', start=int(min_day_espoir), end=int(max_day_espoir), step=1, value=int(max_day_espoir))

max_month_espoir = df_espoir['Month'].max()
min_month_espoir = df_espoir['Month'].min()
avg_month_espoir = df_espoir['Month'].mean()
std_month_espoir = df_espoir['Month'].std()

int_slider_submitted_month = pn.widgets.IntSlider(name='Month ID', start=int(min_month_espoir), end=int(max_month_espoir), step=1, value=int(max_month_espoir))
int_slider_submitted_month_all = pn.widgets.IntSlider(name='Month ID', start=int(min_month_espoir), end=int(max_month_espoir), step=1, value=int(max_month_espoir))

idf_espoir_month = df_espoir.interactive()

idf_espoir_month = (
    idf_espoir_month[
            (idf_espoir_month['Status'] == select_espoir_status_month)
        ]
        .groupby(['Status', 'Date Submitted', 'Caller Type', 'Department', 'Day', 'Month Name', 'Issue Priority', 'Contact Person', 'Day Resolved'])['Month'].max()
        .to_frame()
        .reset_index()
        .sort_values(by='Date Submitted')
        .reset_index(drop=True)
)

idf_espoir_day = df_espoir.interactive()

idf_espoir_day = (
    idf_espoir_day[
            (idf_espoir_day['Status'] == select_espoir_status_day)
        ]
        .groupby(['Status', 'Date Submitted', 'Caller Type', 'Department', 'Day',  'Month Name', 'Issue Priority', 'Contact Person', 'Day Resolved'])['Month'].max()
        .to_frame()
        .reset_index()
        .sort_values(by='Date Submitted')
        .reset_index(drop=True)
)


idf_espoir = df_espoir.interactive()

idf_espoir = (
    idf_espoir[
            (idf_espoir['Department'] == select_espoir_department)
        ]
        .groupby(['Status', 'Date Submitted', 'Caller Type', 'Department', 'Day', 'Month Name', 'Year', 'Issue Priority', 'Reporter', 'Contact Person', 'Day Resolved'])['Month'].max()
        .to_frame()
        .reset_index()
        .sort_values(by='Date Submitted')
        .reset_index(drop=True)
)


idf_espoir_all = df_espoir.interactive()

idf_espoir_all = (
    idf_espoir_all[
            (idf_espoir_all['Department'] == select_espoir_department_all)
        ]
        .groupby(['Status', 'Date Submitted', 'Caller Type', 'Department', 'Day', 'Month Name', 'Year', 'Issue Priority', 'Reporter', 'Contact Person', 'Day Resolved'])['Month'].max()
        .to_frame()
        .reset_index()
        .sort_values(by='Date Submitted')
        .reset_index(drop=True)
)

########################################################################################################################

image = pn.pane.JPG("dondonedmond.jpg")

########################################################################################################################

table = pn.pane.DataFrame(df_espoir_example, styles=CARD_STYLE)

table_month = idf_espoir_month.pipe(pn.widgets.Tabulator, pagination='remote', page_size = 7, width=1380)
table_day = idf_espoir_day.pipe(pn.widgets.Tabulator, pagination='remote', page_size = 7, width=1380)   

location_day_bar_status  = idf_espoir.hvplot(x = 'Month Name', y = "Month", 
kind = "bar", by='Status', rot = 45, height = 480,  width = 660,  title='Bar Graph of Month Name vs Month by Status',  stacked = True)

location_month_bar_status  = idf_espoir.hvplot(x = 'Month Name', y = "Day Resolved", 
kind = "bar", by='Status', rot = 45, height = 480,  width = 660,  title='Bar Graph of Month Name vs Day by Status',  stacked = True)

location_day_scatter  = idf_espoir_all.hvplot(x = 'Day', y = "Month", 
kind = "scatter", by='Status', rot = 45, height = 480,  width = 660,  title='Scatter Graph Day vs Month by Status',  stacked = True)

location_month_scatter = idf_espoir_all.hvplot(x = 'Month', y = "Day", 
kind = "scatter", by='Status', rot = 45, height = 480,  width = 660,  title='Scatter Graph Month vs Day by Status',  stacked = True)

location_day_bar_priority  = idf_espoir_all.hvplot(x = 'Month Name', y = "Day Resolved", 
kind = "bar", by='Issue Priority',  rot = 45, height = 480,  width = 660,  title='Bar Graph Month Name vs Day by Issue Priority sorted with Department',  stacked = True)

location_month_bar_priority = idf_espoir_all.hvplot(x = 'Month Name', y = "Month", 
kind = "bar", by='Issue Priority', rot = 45, height = 480,  width = 660,  title='Bar Graph Month Name vs Month by Issue Priority sorted with Department',  stacked = True)

location_day_scatter_year  = idf_espoir_all.hvplot(x = 'Month Name', y = "Day", 
kind = "scatter", by='Issue Priority',  rot = 45, height = 480,  width = 660,  title='Bar Graph of Issue Priority vs Day by Status sorted with Department',  stacked = True)

location_month_scatter_year = idf_espoir_all.hvplot(x = 'Month Name', y = "Month", 
kind = "scatter", by='Issue Priority', rot = 45, height = 480,  width = 660,  title='Bar Graph of Issue Priority vs Month by Status sorted with Department',  stacked = True)

########################################################################################################################

main_container = pn.Row(
    max_width=1024,
    styles={"margin-right": "auto", "margin-left": "auto", "margin-top": "10px", "margin-bottom": "20px"},
)
report = pn.Column(header, main_container)

########################################################################################################################

report.save("index.html")

########################################################################################################################


Login_Template = pn.template.FastListTemplate(
title = f'{max_day_to_resolved}',
sidebar =  [
            "# Tickets 2018",
            image,
           ],   
main    =  [

        pn.Column
        (
            "## Tickets 2018 Day and Month (Count) & (Max Day ID, Min Day ID) & (Max Month ID, Min Month ID) & (Max Day Resolved)", indicators_month_top, name='Tickets 2018 Day and Month Number ID'
        ),

        pn.Tabs
        (
            pn.Column
            (
                "### Tickets 2018 Day and Month Name Details", table_day, name='Tickets 2018 Day and Month Name Table by Status Table'
            ),

            # pn.Column
            # (
            #     "### Tickets 2018 Month Details", table_month, name='Tickets 2018 Day and Month Table by Status'
            # ),

            pn.Column
            (
                "### Tickets 2018 Day and Month Name Details", 
                pn.Row(pn.Column(location_day_bar_status),
                pn.Column(location_month_bar_status)), name='Tickets 2018 Month Name and Day Bar by Status'
            ),

            # pn.Column
            # (
            #     "## Tickets 2018 Month Details", indicators_month,
            #     pn.Row(pn.Column(location_day_scatter),
            #     pn.Column(location_month_scatter)), name='Tickets 2018 Month Scatter'
            # ),


            pn.Column
            (
                "### Tickets 2018 Day and Month Name Details", 
                pn.Row(pn.Column(location_month_bar_priority),
                pn.Column(location_day_bar_priority)), name='Tickets 2018 Month Name and Day Bar by Priority'
            ),

            # pn.Column
            # (
            #     "## Tickets 2018 Day by Year Details", indicators_month,
            #     pn.Row(pn.Column(location_day_scatter_year),
            #     pn.Column(location_month_scatter_year)), name='Tickets 2018 Day and Month Scatter by Status'
            # ),
        ),
    ],
)

Login_Template.servable();