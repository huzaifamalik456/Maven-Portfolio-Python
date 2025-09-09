#!/usr/bin/env python
# coding: utf-8

# In[2]:


import pandas as pd
import numpy as np
import seaborn as sns
import plotly.express as px
from dash import Dash, html, dcc
from dash.dependencies import Output, Input
from jupyter_dash import JupyterDash


# In[3]:


from dash_bootstrap_templates import load_figure_template
import dash_bootstrap_components as dbc


# In[4]:


#Objective 1 (Profile & QA the data)

#Import the vgchartz-2024.csv file
#Check that each column has an appropriate data type
#Rename the "title", "genre", "publisher", and "developer" columns to start with capital letters
#Create a "release_year" column based on the "release_date" column


# In[5]:


games = pd.read_csv(
    "/Users/huzaifamalik/Downloads/Python for Data Analytics/Interactive Dashboards with Plotly & Dash/Course_Materials/Video+Game+Sales/vgchartz-2024.csv"
)


# In[6]:


games.shape


# In[7]:


games.head()


# In[8]:


games.info()


# In[9]:


games["title"].nunique()


# In[10]:


#Changing the datatype to datetime of date columns
games["release_date"] = pd.to_datetime(games["release_date"])
games["last_update"] = pd.to_datetime(games["last_update"])


# In[11]:


#Renaming the "title", "genre", "publisher", and "developer" columns to start with capital letters
games.rename(
    columns = {"title": "Title", "genre": "Genre", "publisher": "Publisher", "developer": "Developer"}, 
    inplace = True
)


# In[12]:


#Creating a "release_year" column based on the "release_date" column
games["release_year"] = games["release_date"].dt.year.astype("Int64")


# In[ ]:





# In[13]:


#Objective 2 (Prepare the data for visualization)

#Create an "annual sales" table that calculates the sum of total sales by year
#Create "top 10 titles" table which contains 10 highest selling titles by "total_sales", ranked from highest to lowest


# In[14]:


games.head()


# In[ ]:





# In[15]:


annual_sales_table = games.pivot_table(
    index = "release_year",
    values = "total_sales",
    aggfunc = "sum"
)


# In[16]:


annual_sales_table.head()


# In[ ]:





# In[17]:


top_10_titles = (
    games.groupby("Title").agg({"total_sales": "sum"})
    .sort_values("total_sales", ascending = False)
    .head(10)
)


# In[18]:


top_10_titles


# In[ ]:





# In[19]:


#Objective 3 (Build an interactive dashboard)

#Create a line chart to plot total sales by year
#Create a bar chart of the top ten selling titles of all time

#Dashboard should allow user to select "title", "genre", "publisher", "developer", and "console" with a Dropdown, 
#And "total_sales", "jp_sales", "na_sales", "pal_sales" and "other_sales" with Radio buttons

#Use Dropdown menu to control category labels in the bar chart 
#And the Radio buttons to control the sales figures shown both charts


# In[ ]:





# In[20]:


dropdown_options = []

for i in list(games.loc[:, ["Title", "Genre", "Publisher", "Developer", "console"]].columns):
    dropdown_options.append({"label": i.title(), "value": i})
                                 
                                 
metric_options = []

for i in list(games.loc[:, ["total_sales", "na_sales", "jp_sales", "pal_sales", "other_sales"]].columns):
    
    if i == "na_sales":
        metric_options.append({"label": "North American Sales" , "value": i})
        
    elif i == "jp_sales":
        metric_options.append({"label": "Japan Sales" , "value": i})
        
    elif i == "pal_sales":
        metric_options.append({"label": "European & African Sales" , "value": i})
    
    elif i == "other_sales":
        metric_options.append({"label": "Rest of World Sales" , "value": i})
    
    else:
        metric_options.append({"label": i.replace("_", " ").title(), "value": i})
                                 

dbc_css = "https://cdn.jsdelivr.net/gh/AnnMarieW/dash-bootstrap-templates/dbc.min.css"
            
app = Dash(__name__, external_stylesheets=[dbc.themes.CYBORG, dbc_css])

load_figure_template("CYBORG")

app.layout = dbc.Container(
    [
        dbc.Row(html.H2("Video Game Explorer", style={"textAlign": "center"})),
        #html.Br(),
        
        dbc.Row(
            [
                dbc.Col(html.H5("Please select from the Drop Down Options below: "), 
                        width = 6, className = "text-center"),
                dbc.Col(html.H5("Please select a Metric from the Radio Options below: "), 
                        width = 6, className = "text-center")
            ]
        ),
        
        dbc.Row(
            [
                dbc.Col(
                    dcc.Dropdown(
                        id = "dropdown_menu", options = dropdown_options, 
                        value = "Title", className = "dbc",  style={"textAlign": "center"}), width = 6
                ),
                
                dbc.Col(
                    dcc.RadioItems(
                        id = "metric_button", options = metric_options, 
                        labelStyle={"margin-right": "25px"}, style = {"textAlign": "center"},
                        value = "total_sales", className = "dbc", inline = True), width = 6
                )
            ]
        ),
        
        #html.Br(),
        
        dbc.Row(dcc.Graph(id = "line_chart")),
        #html.Br(),
        dbc.Row(dcc.Graph(id = "bar_chart"))
    
    ],
    
    fluid = True

)
                            
@app.callback(
    [Output("line_chart", "figure"), Output("bar_chart", "figure")], 
    [Input("metric_button", "value"), Input("dropdown_menu", "value")]
)

def produce_charts(metric_selection, dropdown_selection):
    
    for index, element in enumerate(metric_options):
    
        if metric_selection == metric_options[index]["value"]:
            text_for_graphs = metric_options[index]["label"]

            line_chart_title = "Trendline showing " + text_for_graphs + " for Video Games Over Time"
            #bar_chart_title = text_for_graphs + " by " + dropdown_selection + " of Video Games"
            break
            
    
    agg_table_for_line_graph = games.pivot_table(
        index = "release_year",
        values = metric_selection,
        aggfunc = "sum"
    )
    
    line_graph = px.line(
        
        agg_table_for_line_graph.reset_index().sort_values("release_year", ascending = True),
        x = "release_year",
        y = metric_selection,
        title = line_chart_title
    )
    
    line_graph.update_xaxes(title_text = "Year")
    line_graph.update_yaxes(title_text = text_for_graphs)
    
    line_graph.update_layout(
            title = {"x": 0.5, "xanchor": "center", "yanchor": "top"},
            title_font=dict(size=24) 
        )
    
    
    
    if not dropdown_selection or dropdown_selection is None:
        bar_graph = px.bar()
    
    else:
        agg_table_for_bar_chart = (
            games.groupby(dropdown_selection).agg({metric_selection: "sum"})
            .sort_values(metric_selection, ascending = False).head(10)
        )
        
        bar_chart_title = text_for_graphs + " by " + dropdown_selection.title() + " of Video Games"

        bar_graph = px.bar(
            agg_table_for_bar_chart,
            x = agg_table_for_bar_chart.index,
            y = metric_selection,
            title = bar_chart_title
        )
        
        
        bar_graph.update_xaxes(title_text = None)
        bar_graph.update_yaxes(title_text = text_for_graphs)
        
        bar_graph.update_layout(
            title = {"x": 0.5, "xanchor": "center", "yanchor": "top"},
            title_font=dict(size=24) 
        )


    return line_graph, bar_graph
                            
                            
#app.run(port = 8082)

import webbrowser
from threading import Timer

port = 8082
url = f"http://127.0.0.1:{port}/"

def open_browser():
    webbrowser.open_new(url)   # always new window

Timer(1, open_browser).start()
app.run(port=port, debug=True)
                            


# In[ ]:





# In[21]:


dropdown_options = []

for i in list(games.loc[:, ["Title", "Genre", "Publisher", "Developer", "console"]].columns):
    dropdown_options.append({"label": i.title(), "value": i})
    
dropdown_options


# In[ ]:





# In[22]:


metric_options = []

for i in list(games.loc[:, ["total_sales", "na_sales", "jp_sales", "pal_sales", "other_sales"]].columns):
    
    if i == "na_sales":
        metric_options.append({"label": "North American Sales" , "value": i})
        
    elif i == "jp_sales":
        metric_options.append({"label": "Japan Sales" , "value": i})
        
    elif i == "pal_sales":
        metric_options.append({"label": "European & African Sales" , "value": i})
    
    else:
        metric_options.append({"label": i.replace("_", " ").title(), "value": i})
    
metric_options


# In[ ]:





# In[23]:


for i in metric_options:
    print(list(i.items()))


# In[ ]:





# In[24]:


metric_options


# In[25]:


metric_selection = "pal_sales"
dropdown_selection = "Genre"

for index, element in enumerate(metric_options):
    
    if metric_selection == metric_options[index]["value"]:
        
        line_chart_title = "Trendline showing " + metric_options[index]["label"] + " Over Time"
        bar_chart_title = metric_options[index]["label"] + " by " + dropdown_selection
        
        break
        
print(line_chart_title)
print(bar_chart_title)
    


# In[26]:


metric_options[1]["value"]


# In[27]:


metric_options[1]["label"]


# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[28]:


games.head()


# In[ ]:





# In[29]:


metric_selection = "na_sales"
dropdown_selection = "Genre"

agg_table_for_line_graph = games.pivot_table(
        index = "release_year",
        values = metric_selection,
        aggfunc = "sum"
    )
    
line_graph = px.line(
        
    agg_table_for_line_graph.reset_index().sort_values("release_year", ascending = True),
    x = "release_year",
    y = metric_selection
)

agg_table_for_bar_chart = (
    games.groupby(dropdown_selection).agg({metric_selection: "sum"})
    .sort_values(metric_selection, ascending = False).head(10)
)

bar_graph = px.bar(
    agg_table_for_bar_chart,
    x = agg_table_for_bar_chart.index,
    y = metric_selection
)


# In[30]:


line_graph.show()


# In[ ]:





# In[31]:


bar_graph.show()


# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[32]:


import warnings

warnings.filterwarnings(
    "ignore",
    message="Plotly version >= 6 requires Jupyter Notebook >= 7"
)


# In[33]:


px.line(
    annual_sales_table,
    x = annual_sales_table.index,
    y = "total_sales"
)


# In[ ]:





# In[34]:


px.bar(
    top_10_titles,
    x = top_10_titles.index,
    y = "total_sales",
    #color = top_10_titles.index
)


# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:




