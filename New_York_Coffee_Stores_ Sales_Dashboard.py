#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import numpy as np
import seaborn as sns


# In[ ]:


#Objective 1 (Prepare the data for analysis)


# In[ ]:


#How many transactions were recorded, over what period of time? What products and product categories were sold?


# In[14]:


coffee = pd.read_excel(
    "/Users/huzaifamalik/Downloads/Python for Data Analytics/Interactive Dashboards with Plotly & Dash/Course_Materials/Coffee Shop Sales Guided Project/Coffee Shop Sales.xlsx",
    dtype = {"transaction_id": "Int32", "transaction_qty": "Int8", "store_id": "Int8", "product_id": "Int8"}
)


# In[16]:


coffee.shape


# In[17]:


coffee.head()


# In[18]:


coffee.info()


# In[24]:


print(coffee["store_location"].nunique())
print(coffee["product_category"].nunique())
print(coffee["product_detail"].nunique())


# In[ ]:





# In[ ]:


#Add a new column to calculate Revenue (price * quantity)


# In[26]:


coffee["Revenue"] = coffee["transaction_qty"] * coffee["unit_price"]


# In[ ]:





# In[ ]:


#Add new columns to calculate Month and Day of Week based on the transaction date 
#(BONUS: display them as text (ie “Jan”, “Feb”, “Sun”, “Mon”), instead of numerical values)


# In[33]:


coffee["transaction_date"].dt.month_name()


# In[57]:


month_names = []
for i in list(coffee["transaction_date"].dt.month_name()):
    month_names.append(i[0:3])

print(len(month_names))

print(month_names[0:5])
print("\n")

day_of_week = []

for j in list(coffee["transaction_date"].dt.day_name()):
    day_of_week.append(j[0:3])

print(len(day_of_week))
print(day_of_week[0:5])


# In[54]:


coffee["transaction_date"].dt.day_name()


# In[58]:


coffee["Month"] = month_names
coffee["Day of Week"] = day_of_week


# In[ ]:





# In[ ]:


#Add a new column to extract Hour from the transaction time


# In[66]:


coffee["transaction_time"]


# In[76]:


hours_list = []

for k in list(coffee["transaction_time"]):
    #print(k.hour)
    hours_list.append(k.hour)
    
print(len(hours_list))


# In[77]:


coffee["Hour"] = hours_list


# In[86]:


month_numbers = []
day_of_week_numbers = []

for x, y in zip(list(coffee["transaction_date"].dt.month), list(coffee["transaction_date"].dt.day_of_week)):
    month_numbers.append(x)
    day_of_week_numbers.append(y)

print(len(month_numbers))
print(len(day_of_week_numbers))


# In[87]:


coffee["month_number"] = month_numbers
coffee["day_of_week_number"] = day_of_week_numbers


# In[88]:


coffee.head()


# In[ ]:





# In[ ]:


#Objective 2 (Explore the data with Pivot Tables)


# In[ ]:


#Insert a PivotTable to show revenue by month


# In[92]:


coffee.pivot_table(
    index = ["Month", "month_number"],
    values = "Revenue",
    aggfunc = "sum"
).sort_values("month_number", ascending = True)


# In[ ]:





# In[ ]:


#Add two more PivotTables to show the number of transactions by day of week and by hour of day


# In[115]:


coffee.pivot_table(
    index = ["Day of Week", "day_of_week_number"],
    values = "transaction_id",
    aggfunc = "count"
).sort_values("day_of_week_number", ascending = True).reset_index().drop("day_of_week_number", axis = 1)


# In[ ]:





# In[101]:


coffee.pivot_table(
    index = ["Hour"],
    values = "transaction_id",
    aggfunc = "count"
).reset_index().sort_values("Hour", ascending = True).set_index("Hour")


# In[ ]:





# In[ ]:


#Add a PivotTable to show the number of transactions by product category, sorted descending by transactions


# In[119]:


(
    coffee.groupby("product_category").agg({"transaction_id": "count"})
    .sort_values("transaction_id", ascending = False)
    .rename(columns = {"transaction_id": "number_of_transactions"})
)


# In[ ]:





# In[ ]:


#Add a PivotTable to show the number of transactions and revenue by product type, 
#Sorted descending and filtered to the Top 15 (by transactions)


# In[125]:


(
    coffee.pivot_table(
        index = "product_type",
        values = ["transaction_id", "Revenue"],
        aggfunc = {"transaction_id": "count", "Revenue": "sum"}
    )
    .sort_values("transaction_id", ascending = False)
    .rename(columns = {"transaction_id": "number_of_transactions"})
    .iloc[0:15]
)


# In[ ]:





# In[ ]:


#Objective 3 Build a dynamic dashboard


# In[ ]:


#Add Pivot Charts to show revenue by month as a line chart, 
#Transactions by day of week and hour of day as column charts, 
#And transactions by product category as a bar chart

#Include space for the PivotTable showing Top 15 product types
#Add a slicer for store location, and connect it to all of the PivotTables on the sheet


# In[ ]:





# In[194]:


import plotly.express as px
from dash import Dash, html, dcc
from dash.dependencies import Output, Input
import dash_bootstrap_components as dbc
from dash_bootstrap_templates import load_figure_template
from dash import dash_table


# In[301]:


store_locations_user_options = []

for i in list(coffee["store_location"].unique()):
    store_locations_user_options.append({"label": i.upper(), "value": i})

store_locations_user_options

dbc_css = "https://cdn.jsdelivr.net/gh/AnnMarieW/dash-bootstrap-templates/dbc.min.css"

app = Dash(__name__, external_stylesheets=[dbc.themes.SUPERHERO, dbc_css])
server = app.server

load_figure_template("SUPERHERO")


app.layout = dbc.Container(
    [
        dbc.Row(html.H1("NYC Stores Coffee Sales Interactive Dashboard (Jan-Jun 2023)", style={"textAlign": "center"})),

        dbc.Row(
            children = [
                dbc.Col(
                    dcc.Dropdown(
                        id = "location_menu", options = store_locations_user_options, 
                        placeholder = "Select Store Location", className = "dbc"
                    ), width = 4
                ),

                dbc.Col(dcc.Graph(
                    id = "revenue_chart", style={"height": "290px"}
                ), width = 8)
            ]    
        ),
        html.Br(),

        dbc.Row(
            children = [
                dbc.Col(dcc.Graph(id = "day_of_week_transactions_chart", style={"height": "290px"}), width = 6),
                dbc.Col(dcc.Graph(id = "hour_transactions_chart", style={"height": "290px"}), width = 6)
            ]
        ),
        html.Br(),

        dbc.Row(
            children = [

                dbc.Col(dcc.Graph(id = "prod_category_transactions_chart", style={"height": "290px"}), width = 6),
                
                dbc.Col(
                    dash_table.DataTable(

                        id = "data_table",
                        style_table={
                            "overflowX": "auto",
                            "backgroundColor": "var(--bs-body-bg)",  # dark background
                        },
                        style_header={
                            #"backgroundColor": "var(--bs-secondary-bg)",
                            "color": "var(--bs-body-color)",
                            "fontWeight": "bold",
                            "border": "1px solid var(--bs-border-color)",
                            "fontSize": "12px"
                        },
                        style_data={
                            "backgroundColor": "var(--bs-body-bg)",
                            "color": "white",
                            "border": "1px solid var(--bs-border-color)",
                        },
                        style_cell={
                            "textAlign": "center",
                            "padding": "8px",
                            "fontSize": "12px",
                            "backgroundColor": "var(--bs-body-bg)",
                            "color": "var(--bs-body-color)",
                            "border": "1px solid var(--bs-border-color)",
                        },
                    ),
                    width = 6
                )
            ]

        )
    ],
    fluid = True
    
)

@app.callback(
    [
        Output("revenue_chart", "figure"), Output("day_of_week_transactions_chart", "figure"),
        Output("hour_transactions_chart", "figure"), Output("prod_category_transactions_chart", "figure"),
        Output("data_table", "columns"), Output("data_table", "data")
    ], 
    Input("location_menu", "value")
)

def produce_graphs(location):
    
    if not location or location is None:
        return px.line(), px.bar(), px.bar(), px.bar(), [], []
       
    
    filtered_df = coffee.query("store_location in @location")
    
    
    #Revenue Line Chart
    revenue_by_month = (
        filtered_df.pivot_table(
            index = ["Month", "month_number"],
            values = "Revenue",
            aggfunc = "sum"
        ).sort_values("month_number", ascending = True)
    )

    revenue_fig = px.line(
        revenue_by_month.reset_index(),
        x = "Month",
        y = "Revenue"
    )

    revenue_fig.update_yaxes(
        title_text = "Revenue ($USD)"
    )
    
    revenue_fig.update_layout(
        title=f"Total Revenue by Month for {location}",
        title_x=0.5,
        plot_bgcolor="rgba(0,0,0,0)",   # plot area
        paper_bgcolor="rgba(0,0,0,0)",  # outside plot
        margin=dict(l=0, r=0, t=40, b=0),
        xaxis_title=None
    )
    
    revenue_fig.update_traces(
        mode="lines+markers+text",  # show line, markers, and text
        text=revenue_by_month["Revenue"],
        texttemplate="$%{text:,.0f}",
        textposition="top center"
    )
    
    
    #Transactions Day of Week Bar Chart
    transactions_day_of_week = (
        filtered_df.pivot_table(
            index = ["Day of Week", "day_of_week_number"],
            values = "transaction_id",
            aggfunc = "count"
        )
        .sort_values("day_of_week_number", ascending = True)
        .reset_index()
        .drop("day_of_week_number", axis = 1)
        .rename(columns = {"transaction_id": "Transactions"})
    )

    transactions_day_fig = px.bar(
        transactions_day_of_week,
        x = "Day of Week",
        y = "Transactions"
    )

    transactions_day_fig.update_layout(
        title=f"Total Transactions by Day of Week for {location}",
        title_x=0.5,
        plot_bgcolor="rgba(0,0,0,0)",   # plot area
        paper_bgcolor="rgba(0,0,0,0)",  # outside plot
        margin=dict(l=0, r=0, t=40, b=0),
        xaxis_title=None
    )
    
    
    #Transactions Hour of Day Bar Chart
    
    transactions_hour_of_day = (
        filtered_df.pivot_table(
            index = ["Hour"],
            values = "transaction_id",
            aggfunc = "count"
        )
        .reset_index()
        .sort_values("Hour", ascending = True)
        .set_index("Hour")
        .rename(columns = {"transaction_id": "Transactions"})
    )

    transactions_hour_fig = px.bar(
        transactions_hour_of_day,
        x = transactions_hour_of_day.index,
        y = "Transactions"
    )

    transactions_hour_fig.update_layout(
        title=f"Total Transactions by Hour for {location}",
        title_x=0.5,
        plot_bgcolor="rgba(0,0,0,0)",   # plot area
        paper_bgcolor="rgba(0,0,0,0)",  # outside plot
        margin=dict(l=0, r=0, t=40, b=0),
        xaxis_title=None
    )
    
    
    #Transactions Product Category Bar Chart
    transactions_product_category = (
        filtered_df.groupby("product_category").agg({"transaction_id": "count"})
        .sort_values("transaction_id", ascending = True)
        .rename(columns = {"transaction_id": "Transactions"})
    )

    prod_category_fig = px.bar(
        transactions_product_category,
        y = transactions_product_category.index,
        x = "Transactions"
    )

    prod_category_fig.update_yaxes(title_text = "Product Category")

    prod_category_fig.update_layout(
        title=f"Total Transactions by Product Category for {location}",
        title_x=0.5,
        plot_bgcolor="rgba(0,0,0,0)",   # plot area
        paper_bgcolor="rgba(0,0,0,0)",  # outside plot
        margin=dict(l=0, r=0, t=40, b=0),
        yaxis_title=None
    )
    
    
    
    #Transactions Top 15 Product Categorys Table
    top_15_product_types = (
        filtered_df.pivot_table(
            index = "product_type",
            values = ["transaction_id", "Revenue"],
            aggfunc = {"transaction_id": "count", "Revenue": "sum"}
        )
        .sort_values("transaction_id", ascending = False)
        .rename(columns = {"transaction_id": "Transactions", "Revenue": "Revenue ($USD)"})
        .iloc[0:15]
    ).reset_index().rename(columns = {"product_type": "Top 15 Products"}).loc[:, ["Top 15 Products", "Transactions", "Revenue ($USD)"]]
    
    
    top_15_product_types = (
        top_15_product_types.assign(
            rounded_rev = round(top_15_product_types["Revenue ($USD)"],0).astype("int")
        )
        .drop("Revenue ($USD)", axis = 1)
        .rename(columns = {"rounded_rev": "Revenue ($USD)"})
    )
    
    table_columns = [{"name": col, "id": col} for col in top_15_product_types.columns]
    table_records = top_15_product_types.to_dict("records")
    
    return revenue_fig, transactions_day_fig, transactions_hour_fig, prod_category_fig, table_columns, table_records

    
    
#app.run(port = 8080)

# import webbrowser
# from threading import Timer

# port = 8080
# url = f"http://127.0.0.1:{port}/"

# def open_browser():
#     webbrowser.open_new(url)   # always new window

# Timer(1, open_browser).start()
# app.run(port=port, debug=True)






# In[ ]:





# In[130]:


coffee["store_location"].unique()


# In[135]:


store_locations_user_options = []

for i in list(coffee["store_location"].unique()):
    store_locations_user_options.append({"label": i.upper(), "value": i})

store_locations_user_options


# In[133]:


for i in list(coffee["store_location"].unique()):
    print(i.upper())


# In[145]:


import warnings
warnings.filterwarnings("ignore", message="Plotly version >= 6 requires Jupyter Notebook >= 7")


# In[292]:


location = "Lower Manhattan"

filtered_df = coffee.query("store_location in @location")

revenue_by_month = (
    filtered_df.pivot_table(
        index = ["Month", "month_number"],
        values = "Revenue",
        aggfunc = "sum"
    ).sort_values("month_number", ascending = True)
)

revenue_fig = px.line(
    revenue_by_month.reset_index(),
    x = "Month",
    y = "Revenue"
)

revenue_fig.update_yaxes(
    title_text = "Revenue ($USD)"
)

revenue_fig.update_layout(
    title=f"Total Monthly Revenue for {location} location from Jan-Jun 2023",
    title_x=0.5
)


# In[ ]:





# In[293]:


transactions_day_of_week = (
    filtered_df.pivot_table(
        index = ["Day of Week", "day_of_week_number"],
        values = "transaction_id",
        aggfunc = "count"
    )
    .sort_values("day_of_week_number", ascending = True)
    .reset_index()
    .drop("day_of_week_number", axis = 1)
    .rename(columns = {"transaction_id": "Transactions"})
)

transactions_day_fig = px.bar(
    transactions_day_of_week,
    x = "Day of Week",
    y = "Transactions"
)

transactions_day_fig.update_layout(
    title=f"Total Transactions per Day of Week for {location} location from Jan-Jun 2023",
    title_x=0.5
)


# In[ ]:





# In[294]:


transactions_hour_of_day = (
    filtered_df.pivot_table(
        index = ["Hour"],
        values = "transaction_id",
        aggfunc = "count"
    )
    .reset_index()
    .sort_values("Hour", ascending = True)
    .set_index("Hour")
    .rename(columns = {"transaction_id": "Transactions"})
)

transactions_hour_fig = px.bar(
    transactions_hour_of_day,
    x = transactions_hour_of_day.index,
    y = "Transactions"
)

transactions_hour_fig.update_layout(
    title=f"Total Transactions per Hour for Store Locations {location} from Jan-Jun 2023",
    title_x=0.5
)


# In[ ]:





# In[295]:


transactions_product_category = (
    filtered_df.groupby("product_category").agg({"transaction_id": "count"})
    .sort_values("transaction_id", ascending = True)
    .rename(columns = {"transaction_id": "Transactions"})
)

prod_category_fig = px.bar(
    transactions_product_category,
    y = transactions_product_category.index,
    x = "Transactions"
)

prod_category_fig.update_yaxes(title_text = "Product Category")

prod_category_fig.update_layout(
    title=f"Total Transactions per Product Category for Store Locations {location} from Jan-Jun 2023",
    title_x=0.5
)


# In[ ]:





# In[296]:


top_15_product_types = (
    filtered_df.pivot_table(
        index = "product_type",
        values = ["transaction_id", "Revenue"],
        aggfunc = {"transaction_id": "count", "Revenue": "sum"}
    )
    .sort_values("transaction_id", ascending = False)
    .rename(columns = {"transaction_id": "Transactions"})
    .iloc[0:15]
).reset_index().rename(columns = {"product_type": "Top 15 Products"}).loc[:, ["Top 15 Products", "Transactions", "Revenue"]]


top_15_product_types = (
    top_15_product_types.assign(
        rounded_rev = round(top_15_product_types["Revenue"],0).astype("int")
    )
    .drop("Revenue", axis = 1)
    .rename(columns = {"rounded_rev": "Revenue"})
)

top_15_product_types


# In[ ]:





# In[213]:


[{"name": col, "id": col} for col in top_15_product_types.columns]


# In[214]:


top_15_product_types.to_dict("records")


# In[ ]:





# In[ ]:




