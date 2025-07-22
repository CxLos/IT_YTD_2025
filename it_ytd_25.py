# =================================== IMPORTS ================================= #
import csv, sqlite3
import numpy as np 
import pandas as pd 
import seaborn as sns 
import matplotlib.pyplot as plt 
import plotly.figure_factory as ff
import plotly.graph_objects as go
from geopy.geocoders import Nominatim
from folium.plugins import MousePosition
import plotly.express as px
from datetime import datetime
import folium
import os
import sys
# ------
import json
import base64
import gspread
from oauth2client.service_account import ServiceAccountCredentials
# ------
import dash
from dash import dcc, html
from dash.dependencies import Input, Output, State
from dash.development.base_component import Component

# 'data/~$bmhc_data_2024_cleaned.xlsx'
# print('System Version:', sys.version)
# -------------------------------------- DATA ------------------------------------------- #

current_dir = os.getcwd()
current_file = os.path.basename(__file__)
script_dir = os.path.dirname(os.path.abspath(__file__))
# data_path = 'data/IT_Responses.xlsx'
# file_path = os.path.join(script_dir, data_path)
# data = pd.read_excel(file_path)
# df = data.copy()

# Define the Google Sheets URL
sheet_url = "https://docs.google.com/spreadsheets/d/1wNUS59k4D6mSq-ciF6PDkcIcrmr9uu867lqP4fM6VyA/edit#gid=1758812507"

# Define the scope
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# Load credentials
encoded_key = os.getenv("GOOGLE_CREDENTIALS")

if encoded_key:
    json_key = json.loads(base64.b64decode(encoded_key).decode("utf-8"))
    creds = ServiceAccountCredentials.from_json_keyfile_dict(json_key, scope)
else:
    creds_path = r"C:\Users\CxLos\OneDrive\Documents\BMHC\Data\bmhc-timesheet-4808d1347240.json"
    if os.path.exists(creds_path):
        creds = ServiceAccountCredentials.from_json_keyfile_name(creds_path, scope)
    else:
        raise FileNotFoundError("Service account JSON file not found and GOOGLE_CREDENTIALS is not set.")

# Authorize and load the sheet
client = gspread.authorize(creds)
sheet = client.open_by_url(sheet_url)
data = pd.DataFrame(client.open_by_url(sheet_url).sheet1.get_all_records())
df = data.copy()

# Trim leading and trailing whitespaces from column names
df.columns = df.columns.str.strip()
df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

# Filtered df where 'Date of Activity:' is between October 1, 2024 to September 30, 2025
df['Date of Activity'] = pd.to_datetime(df['Date of Activity'], format='%m/%d/%Y', errors='coerce')
# df["Date of Activity"] = pd.to_datetime(df["Date of Activity"])
df = df[(df['Date of Activity'] >= '2024-10-01') & (df['Date of Activity'] <= '2025-09-30')]

# Get the reporting year:
report_year = datetime(2025, 1, 1).strftime("%Y")

# Define a discrete color sequence
# color_sequence = px.colors.qualitative.Plotly

# -----------------------------------------------------------

# print(df_m.head())
# print('Total Marketing Events: ', len(df))
# print('Column Names: \n', df.columns.to_list())
# print('DF Shape:', df.shape)
# print('Dtypes: \n', df.dtypes)
# print('Info:', df.info())
# print("Amount of duplicate rows:", df.duplicated().sum())

# print('Current Directory:', current_dir)
# print('Script Directory:', script_dir)
# print('Path to data:',file_path)

# ================================= Columns ================================= #

# Column Names: 

columns = [
    'Timestamp', 
    'Date of Activity',
    'Which form are you filling out?',
    'Person submitting this form:',
    'Activity Duration (hours):', 
    'How much time did you spend on these tasks? (minutes)',
    'Please provide brief description of activity',
    'Email Address'
  ]

# =============================== Missing Values ============================ #

# missing = df.isnull().sum()
# print('Columns with missing values before fillna: \n', missing[missing > 0])

# ============================== Data Preprocessing ========================== #



# ========================= Filtered DataFrames ========================== #

# # List of columns to include
# columns_to_include = [
#     'Person completing this form:.4',
#     'Were all Weekly and Monthly reports completed and submitted on time?',
#     'Was data collected accurately and reviewed for quality?',
#     "Did you identify any actionable insights from this month's data?",
#     'Please note any challenges, reasons for delays, or noteworthy outcomes from completed tasks:.5',
#     'Date:',
#     'Briefly describe what tasks you worked on:',
#     'How much time did you spend on these tasks? (minutes)'
# ]

# columns_to_include_2 = [
#     'Date:',
#     'Person completing this form:.4',
#     'Briefly describe what tasks you worked on:',
#     'How much time did you spend on these tasks? (minutes)'
# ]

# # Create a new DataFrame with only the specified columns
# df_data = df[columns_to_include]
# df_data2 = df[columns_to_include_2]
# # print(df_data.head(10))

# # Total data events
# data_events = len(df_data)

# # Total Data Hours
# total_data_hours = df_data['How much time did you spend on these tasks? (minutes)'].sum()/60

# # "Person completing this form:" dataframe:
# df['Person completing this form:.4'] = df['Person completing this form:.4'].str.strip()
# df_person = df.groupby('Person completing this form:.4').size().reset_index(name='Count')
# # print(df_person.value_counts())

# # Data Table 2
# columns_to_include_2 = [
#     'Date:',
#     'Person completing this form:.4',
#     'Briefly describe what tasks you worked on:',
#     'How much time did you spend on these tasks? (minutes)'
# ]

# df_data2 = df[columns_to_include_2]

# ========================== Total Events ========================== #

# Total IT Events:
it_events = len(df)

# ========================== Total IT Hours ========================== #

# Total IT Hours:
it_hours = df['Activity Duration (minutes):'].sum()/60
it_hours = round(it_hours)

# ======================= Person Submitting Form ======================= #

# Person submitting this form:
df_person = df.groupby('Person submitting this form:').size().reset_index(name='Count')

person_bar=px.bar(
    df_person,
    x='Person submitting this form:',
    y='Count',
    color='Person submitting this form:',
    text='Count',
).update_layout(
    height=440, 
    width=780,
    title=dict(
        text='People Submitting Forms',
        x=0.5, 
        font=dict(
            size=25,
            family='Calibri',
            color='black',
        )
    ),
    font=dict(
        family='Calibri',
        size=18,
        color='black'
    ),
    xaxis=dict(
        tickangle=0,  # Rotate x-axis labels for better readability
        tickfont=dict(size=18),  # Adjust font size for the tick labels
        title=dict(
            text=None,
            # text="Name",
            font=dict(size=20),  # Font size for the title
        ),
        # showticklabels=False  # Hide x-tick labels
        showticklabels=True  # Hide x-tick labels
    ),
    yaxis=dict(
        title=dict(
            text='Count',
            font=dict(size=20),  # Font size for the title
        ),
    ),
    legend=dict(
        # title='Support',
        title_text='',
        orientation="v",  # Vertical legend
        x=1.05,  # Position legend to the right
        y=1,  # Position legend at the top
        xanchor="left",  # Anchor legend to the left
        yanchor="top",  # Anchor legend to the top
        visible=False
        # visible=True
    ),
    hovermode='closest', # Display only one hover label per trace
    bargap=0.08,  # Reduce the space between bars
    bargroupgap=0,  # Reduce space between individual bars in groups
).update_traces(
    textposition='auto',
    hovertemplate='<b>Name:</b> %{label}<br><b>Count</b>: %{y}<extra></extra>'
)

# Person Pie Chart
person_pie=px.pie(
    df_person,
    names="Person submitting this form:",
    values='Count'  # Specify the values parameter
).update_layout(
    title='Ratio of People Filling Out Forms',
    title_x=0.5,
    font=dict(
        family='Calibri',
        size=17,
        color='black'
    )
).update_traces(
    rotation=0,
    textposition='auto',
    textinfo='value+percent',
    hovertemplate='<b>%{label} Status</b>: %{value}<extra></extra>',
)

# =================== Which Form are you filling out =================== #

# Which form are you filling out?
df_form = df.groupby('Which form are you filling out?').size().reset_index(name='Count')

form_bar=px.bar(
    df_form,
    x='Which form are you filling out?',
    y='Count',
    color='Which form are you filling out?',
    text='Count',
).update_layout(
    height=440,
    width=780,
    title=dict(
        text='Forms Filled Out',
        x=0.5,
        font=dict(
            size=25,
            family='Calibri',
            color='black',
        )
    ),
    font=dict(
        family='Calibri',
        size=18,
        color='black'
    ),
    xaxis=dict(
        tickangle=0,  # Rotate x-axis labels for better readability
        tickfont=dict(size=18),  # Adjust font size for the tick labels
        title=dict(
            text=None,
            # text="Name",
            font=dict(size=20),  # Font size for the title
        ),
        # showticklabels=False  # Hide x-tick labels
        showticklabels=True  # Hide x-tick labels
    ),
    yaxis=dict(
        title=dict(
            text='Count',
            font=dict(size=20),  # Font size for the title
        ),
    ),
    legend=dict(
        # title='Support',
        title_text='',
        orientation="v",  # Vertical legend
        x=1.05,  # Position legend to the right
        y=1,  # Position legend at the top
        xanchor="left",  # Anchor legend to the left
        yanchor="top",  # Anchor legend to the top
        visible=False
        # visible=True
    ),
    hovermode='closest', # Display only one hover label per trace
    bargap=0.08,  # Reduce the space between bars
    bargroupgap=0,  # Reduce space between individual bars in groups
).update_traces(
    textposition='auto',
    hovertemplate='<b>Form:</b> %{label}<br><b>Count</b>: %{y}<extra></extra>'
)

# Form Pie Chart
form_pie=px.pie(
    df_form,
    names="Which form are you filling out?",
    values='Count'  # Specify the values parameter
).update_layout(
    title='Ratio of Forms Filled Out',
    title_x=0.5,
    font=dict(
        family='Calibri',
        size=17,
        color='black'
    )
).update_traces(
    rotation=0,
    textposition='auto',
    textinfo='value+percent',
    hovertemplate='<b>%{label} Status</b>: %{value}<extra></extra>',
)

# # ========================== DataFrame Table ========================== #

# df_data2 data table
data_table = go.Figure(data=[go.Table(
    # columnwidth=[50, 50, 50],  # Adjust the width of the columns
    header=dict(
        values=list(df.columns),
        fill_color='paleturquoise',
        align='center',
        height=30,  # Adjust the height of the header cells
        # line=dict(color='black', width=1),  # Add border to header cells
        font=dict(size=12)  # Adjust font size
    ),
    cells=dict(
        values=[df[col] for col in df.columns],
        fill_color='lavender',
        align='left',
        height=25,  # Adjust the height of the cells
        # line=dict(color='black', width=1),  # Add border to cells
        font=dict(size=12)  # Adjust font size
    )
)])

data_table.update_layout(
    margin=dict(l=50, r=50, t=30, b=40),  # Remove margins
    height=400,
    # width=1500,  # Set a smaller width to make columns thinner
    paper_bgcolor='rgba(0,0,0,0)',  # Transparent background
    plot_bgcolor='rgba(0,0,0,0)'  # Transparent plot area
)

# ============================== Dash Application ========================== #

app = dash.Dash(__name__)
server= app.server   

app.layout = html.Div(
    children=[ 
    html.Div(
        className='divv', 
        children=[ 
        html.H1(
            'IT Report', 
            className='title'),
        html.H1(
            f'{report_year}', 
            className='title2'),
    html.Div(
        className='btn-box', 
        children=[
        html.A(
            'Repo',
            href=f'https://github.com/CxLos/IT_YTD_{report_year}',
            className='btn'),
        ]),
    ]),    

# Data Table
# html.Div(
#     className='row0',
#     children=[
#         html.Div(
#             className='table',
#             children=[
#                 html.H1(
#                     className='table-title',
#                     children='Data Table'
#                 )
#             ]
#         ),
#         html.Div(
#             className='table2', 
#             children=[
#                 dcc.Graph(
#                     className='data',
#                     figure=data_table
#                 )
#             ]
#         )
#     ]
# ),

html.Div(
    className='row1',
    children=[
        html.Div(
            className='graph11',
            children=[
            html.Div(
                className='high1',
                children=[f'{report_year} IT Events:']
            ),
            html.Div(
                className='circle',
                children=[
                    html.Div(
                        className='hilite',
                        children=[
                            html.H1(
                            className='high2',
                            children=[it_events]
                    ),
                        ]
                    )
 
                ],
            ),
            ]
        ),
        html.Div(
            className='graph22',
            children=[
            html.Div(
                className='high3',
                children=[f'{report_year} IT Hours:']
            ),
            html.Div(
                className='circle2',
                children=[
                    html.Div(
                        className='hilite',
                        children=[
                            html.H1(
                            className='high4',
                            children=[it_hours]
                    ),
                        ]
                    )
                ],
            ),
            ]
        ),
    ]
),

# ROW 2
html.Div(
    className='row2',
    children=[
        html.Div(
            className='graph3',
            children=[
                dcc.Graph(
                figure=person_bar
                )
            ]
        ),
        html.Div(
            className='graph4',
            children=[
                dcc.Graph(
                figure=person_pie
                )
            ]
        )
    ]
),

# ROW 2
# html.Div(
#     className='row2',
#     children=[
#         html.Div(
#             className='graph3',
#             children=[
#                 dcc.Graph(
#                 figure=form_bar
#                 )
#             ]
#         ),
#         html.Div(
#             className='graph4',
#             children=[
#                 dcc.Graph(
#                 figure=form_pie
#                 )
#             ]
#         )
#     ]
# ),

])

print(f"Serving Flask app '{current_file}'! 🚀")

if __name__ == '__main__':
    app.run_server(debug=
                   True)
                #    False)
# =================================== Updated Database ================================= #)

# updated_path1 = 'data/service_tracker_q4_2024_cleaned.csv'
# data_path1 = os.path.join(script_dir, updated_path1)
# df.to_csv(data_path1, index=False)
# print(f"DataFrame saved to {data_path1}")

# updated_path = f'data/IT_{current_month}_{report_year}.xlsx'
# data_path = os.path.join(script_dir, updated_path)

# with pd.ExcelWriter(data_path, engine='xlsxwriter') as writer:
#     df.to_excel(
#             writer, 
#             sheet_name=f'IT {current_month} {report_year}', 
#             startrow=1, 
#             index=False
#         )

#     # Access the workbook and each worksheet
#     workbook = writer.book
#     sheet1 = writer.sheets['IT April 2025']
    
#     # Define the header format
#     header_format = workbook.add_format({
#         'bold': True, 
#         'font_size': 13, 
#         'align': 'center', 
#         'valign': 'vcenter',
#         'border': 1, 
#         'font_color': 'black', 
#         'bg_color': '#B7B7B7',
#     })
    
#     # Set column A (Name) to be left-aligned, and B-E to be right-aligned
#     left_align_format = workbook.add_format({
#         'align': 'left',  # Left-align for column A
#         'valign': 'vcenter',  # Vertically center
#         'border': 0  # No border for individual cells
#     })

#     right_align_format = workbook.add_format({
#         'align': 'right',  # Right-align for columns B-E
#         'valign': 'vcenter',  # Vertically center
#         'border': 0  # No border for individual cells
#     })
    
#     # Create border around the entire table
#     border_format = workbook.add_format({
#         'border': 1,  # Add border to all sides
#         'border_color': 'black',  # Set border color to black
#         'align': 'center',  # Center-align text
#         'valign': 'vcenter',  # Vertically center text
#         'font_size': 12,  # Set font size
#         'font_color': 'black',  # Set font color to black
#         'bg_color': '#FFFFFF'  # Set background color to white
#     })

#     # Merge and format the first row (A1:E1) for each sheet
#     sheet1.merge_range('A1:M1', f'IT Report {current_month} {report_year}', header_format)

#     # Set column alignment and width
#     # sheet1.set_column('A:A', 20, left_align_format)   

#     print(f"IT Excel file saved to {data_path}")

# -------------------------------------------- KILL PORT ---------------------------------------------------

# netstat -ano | findstr :8050
# taskkill /PID 24772 /F
# npx kill-port 8050

# ---------------------------------------------- Host Application -------------------------------------------

# 1. pip freeze > requirements.txt
# 2. add this to procfile: 'web: gunicorn impact_11_2024:server'
# 3. heroku login
# 4. heroku create
# 5. git push heroku main

# Create venv 
# virtualenv venv 
# source venv/bin/activate # uses the virtualenv

# Update PIP Setup Tools:
# pip install --upgrade pip setuptools

# Install all dependencies in the requirements file:
# pip install -r requirements.txt

# Check dependency tree:
# pipdeptree
# pip show package-name

# Remove
# pypiwin32
# pywin32
# jupytercore

# ----------------------------------------------------

# Name must start with a letter, end with a letter or digit and can only contain lowercase letters, digits, and dashes.

# Heroku Setup:
# heroku login
# heroku create mc-impact-11-2024
# heroku git:remote -a mc-impact-11-2024
# git push heroku main

# Clear Heroku Cache:
# heroku plugins:install heroku-repo
# heroku repo:purge_cache -a mc-impact-11-2024

# Set buildpack for heroku
# heroku buildpacks:set heroku/python

# Heatmap Colorscale colors -----------------------------------------------------------------------------

#   ['aggrnyl', 'agsunset', 'algae', 'amp', 'armyrose', 'balance',
            #  'blackbody', 'bluered', 'blues', 'blugrn', 'bluyl', 'brbg',
            #  'brwnyl', 'bugn', 'bupu', 'burg', 'burgyl', 'cividis', 'curl',
            #  'darkmint', 'deep', 'delta', 'dense', 'earth', 'edge', 'electric',
            #  'emrld', 'fall', 'geyser', 'gnbu', 'gray', 'greens', 'greys',
            #  'haline', 'hot', 'hsv', 'ice', 'icefire', 'inferno', 'jet',
            #  'magenta', 'magma', 'matter', 'mint', 'mrybm', 'mygbm', 'oranges',
            #  'orrd', 'oryel', 'oxy', 'peach', 'phase', 'picnic', 'pinkyl',
            #  'piyg', 'plasma', 'plotly3', 'portland', 'prgn', 'pubu', 'pubugn',
            #  'puor', 'purd', 'purp', 'purples', 'purpor', 'rainbow', 'rdbu',
            #  'rdgy', 'rdpu', 'rdylbu', 'rdylgn', 'redor', 'reds', 'solar',
            #  'spectral', 'speed', 'sunset', 'sunsetdark', 'teal', 'tealgrn',
            #  'tealrose', 'tempo', 'temps', 'thermal', 'tropic', 'turbid',
            #  'turbo', 'twilight', 'viridis', 'ylgn', 'ylgnbu', 'ylorbr',
            #  'ylorrd'].

# rm -rf ~$bmhc_data_2024_cleaned.xlsx
# rm -rf ~$bmhc_data_2024.xlsx
# rm -rf ~$bmhc_q4_2024_cleaned2.xlsx