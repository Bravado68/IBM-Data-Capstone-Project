# -*- coding: utf-8 -*-
"""
Created on Thu Nov 18 17:03:09 2021

@author: saf167687
"""

import dash
import dash_bootstrap_components as dbc
import dash_html_components as html
import requests
import pandas as pd
import dash_core_components as dcc
import plotly.express as px
import numpy as np
from dash.dependencies import Input,Output
import dash_table


app = dash.Dash(external_stylesheets = [ dbc.themes.FLATLY],)

df=pd.read_csv("https://cf-courses-data.s3.us.cloud-object-storage.appdomain.cloud/IBM-DS0321EN-SkillsNetwork/datasets/spacex_launch_dash.csv")

df_group=df[['Launch Site',"class"]]
df_dash=df_group.groupby(['Launch Site']).sum()
df_dash["percentage"]=round(100*df_dash["class"]/(df_dash["class"].sum()),1)
df_dash
df_dash.reset_index(level=0, inplace=True)


def figure(df):
    fig = px.pie(df_dash, values="percentage", names='Launch Site', title='Total Success Launches By Site')
    
    return fig


body_app=dbc.Container([
    
    dbc.Row(html.H1(id="H1",children="SpaceX Launcg Records Dashboard",style={"textAlign":"Center"}),
            ),
    
    dbc.Row([
        dcc.Dropdown(id="Launches-Dropdown", options=[
            {"label":i,"value":i} for i in ["All","Site1","Site2","Site3"]],
            value="Select the site",
            placeholder="Select the site"
            )],style = {'width':'100%', 'display':'inline-block'}
        ),
    html.Br(),
    dbc.Row([
        html.Div(id = 'pie_div', 
                 children = [dcc.Graph(id = 'pie_plot', 
                        figure = figure(df_dash))],style = {'width':'100%','display':'inline-block'})        
        ]),
    html.Br(),
     dbc.Row([
        
        dcc.Slider( id = 'slider', min = 0,max = 10000,
       step = 2500,
       
       marks={int(i):"{}".format(i)for i in np.linspace(0,10000,5,dtype=int)},
        value = 0)
        
        ]),
     html.Br(),
     dbc.Row([
         html.H3(id="H3", children="Pie Chart for is selected", style={"textAlign":"Left"})]),
     dbc.Row([
         
        dcc.Dropdown(id="Output-Dropdown", options=[
            {"label":i,"value":i} for i in list(df_dash["Launch Site"])],
            value='CCAFS LC-40',
            placeholder="Select the site"
            )],style = {'width':'100%', 'display':'inline-block'} ),

     dbc.Row([
                     
             dcc.Graph(id="pie_output")
             ])

         ])
         


app.layout = html.Div(id = 'parent', children = [body_app])

@app.callback(Output(component_id='pie_output',component_property='figure'),
              [Input(component_id='Output-Dropdown',component_property='value')])


def figure_pie(value):
        
        #A=list(df_dash["Launch Site"])
        df1=df[df["Launch Site"]==value][["Launch Site","class"]]
        df1=df1["class"].value_counts().to_frame()
        df1["percentage"]=round(100*df1["class"]/df1["class"].sum())
        
        return px.pie(df1, values="percentage", names=df1.index, title=value)
         
    
        
    
if __name__ == "__main__":
    app.run_server(debug=False,port=8053)
