from openpyxl import load_workbook
import pandas as pd
import os
from datetime import datetime

import plotly.express as px
import plotly.graph_objects as go

def recalculate_excel(file_path):
    os.system(f"libreoffice --headless --convert-to xls {file_path}")
    os.system(f"rm {file_path}")
    name = file_path.split('.')
    print(f"{name[0]}.xls")
    os.system(f"libreoffice --headless --convert-to xlsx {name[0]}.xls")
    os.system(f"rm {name[0]}.xls")
    
def get_new_excel(input_file_name, output_file_name, feuille, *args):
    #load excel file
    workbook = load_workbook(filename=input_file_name)
    #open workbook
    sheet = workbook.worksheets[0]
    sheet["B1"] = args[0]["nb_velo"]
    sheet["B2"] = args[0]["prix_exp"]
    sheet["B3"] = args[0]["dist_moy_exp"]
    sheet["B4"] = args[0]["nb_exp_sem"]
    sheet["B5"] = args[0]["nb_exp_we"]
    sheet["B6"] = args[0]["cout_chauffeur"]
    sheet["B7"] = args[0]["cout_stock"]
    sheet = workbook.worksheets[2]
    sheet["B1"] = args[0]["begin_date"]
    
    #save the file
    workbook.save(filename=output_file_name)

    # Recalculer et sauvegarder le fichier
    recalculate_excel(output_file_name)

    # Le fichier recalculé sera sauvegardé dans le même répertoire avec le même nom
    print(f"Le fichier recalculé est sauvegardé")
    
    df = pd.read_excel(output_file_name, sheet_name=feuille, header=0)#, index_col=0)
    
    os.system(f"rm {output_file_name}")
    df = df.dropna(how="any")
    df = df.dropna(axis=1, how="all")
    df = df.reset_index(drop=True)
    df = df.round(2)
    dates = [d for d in df.columns[2:]]
    dates = [col.strftime('%d/%m/%Y') for col in dates]
    df.columns = df.columns[:2].to_list()+dates
    #df.columns = [col.strftime('%d/%m/%Y') for col in df.columns]
    return df

def get_chart(title, x, y1, label='Y1', *args, **kwargs):
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=x, y=y1, mode='lines+markers', name=label))
    for i in range(len(args)):
        lab = f'Y{i+2}'
        if "names" in kwargs.keys():
            if i <= len(kwargs['names']):
                lab = kwargs['names'][i]
        fig.add_trace(go.Scatter(x=x, y=args[i], mode='lines+markers', name=lab))
    
    fig.update_layout(
        title=title,
        plot_bgcolor='white',
        paper_bgcolor='white',
        font=dict(color='black'),
        height=900,
        width=1200,
        xaxis=dict(showgrid=True),
        yaxis=dict(showgrid=True)
    )
    return fig.to_html(full_html=False)

def make_important(x):
    df = pd.DataFrame('', index=x.index, columns=x.columns)
    df.iloc[1] = 'background-color: orange'  # Styliser la deuxième ligne (index 1)
    return df