print("Dashboard is booting up.... ")

from dotenv import load_dotenv
import datetime
from dateutil import parser
from time import sleep
import os
import io
import sys
import warnings
from matplotlib.pyplot import subplots
import pandas as pd
#from sqlalchemy import create_engine
import pymssql
from seaborn import barplot
import PySimpleGUI as sg


if getattr(sys, 'frozen', False):
    output_file_path = os.path.dirname(sys.executable)
    os.chdir(output_file_path)


load_dotenv('.env')
db_host = os.getenv('DB_HOST')
db_user = os.getenv('DB_USER')
db_password = os.getenv('DB_PASSWORD')
db_database = os.getenv('DB_DATABASE')
db_port = os.getenv('DB_PORT')



warnings.filterwarnings('ignore', module='seaborn')
warnings.filterwarnings('ignore', category=UserWarning)


month_list = ["Januari", "Febuari", "Maret", "April", "Mei", "Juni","Juli","Agustus","September","Oktober","Nopember","Desember"]

"""
month_num = date.today().month
year = date.today().year
month_name = month_list[month_num-1]

if month_num >=10:
    date_string = f"{year}-{month_num}-01 00:00:00"
else:
    date_string = f"{year}-0{month_num}-01 00:00:00"

next_string = ""

if month_num == 12:
    next_string = f"{year+1}-01-01 00:00:00"
elif month_num+1 >= 10:
    next_string = f"{year}-{month_num+1}-01 00:00:00"
else:
    next_string = f"{year}-0{month_num+1}-01 00:00:00"
"""


def main():

    
    #conn = create_engine(conn_str)
    
    
    conn = pymssql.connect(server=db_host, port=db_port,user=db_user, password=db_password, database=db_database)
    
    client_list = pd.read_sql("SELECT ClientID,Name from Client", conn) 
    
   
    
    
    layout = [    [sg.Column(        [            [sg.Text("Nama Client", size=(20, 1), font="Lucida", justification='left', background_color="#4E6E81")],
                [sg.Combo(client_list["Name"].tolist(), key="client")],
                [sg.Text("Start Date", size=(20, 1), font="Lucida", justification='left', background_color="#4E6E81")],
                [sg.Input(key="start_date", size=(20,1)), sg.CalendarButton("Choose", target="start_date", format="%Y-%m-%d", )],  # Add a calendar button for start date
                [sg.Text("End Date", size=(20, 1), font="Lucida", justification='left', background_color="#4E6E81")],
                [sg.Input(key="end_date", size=(20,1)), sg.CalendarButton("Choose", target="end_date", format="%Y-%m-%d")],  # Add a calendar button for end date
                [sg.Button('Bikin Dokumen')]
            ], background_color="#4E6E81"
        )]
    ]
    #conn.dispose()
    conn.close()
    
    window = sg.Window("WBK Dashboard", layout,background_color="#4E6E81")
    print("Dashboard has started\n")
    while True:
        event, values = window.read()
        
        if event == sg.WIN_CLOSED:
            print("Dashboard is closing.... ")
            sleep(1)
            break
        
        if event == "Bikin Dokumen":
            print("Creating documents... ")
            client_name = values["client"]
            start_date = values["start_date"]
            end_date = values["end_date"]
            
            date_object = parser.parse(start_date)
            month_num = date_object.month
            year = date_object.year
            month_name = month_list[month_num-1]

            search = client_list.loc[client_list["Name"] == client_name].index
            client_id = client_list["ClientID"].values[search][0]
            query = f"SELECT CreateDateTime,RecordNo,lupdate,Subject,GType,Type,Status,Product FROM dbo.Support WHERE ClientID='{client_id}'"
            #conn = create_engine(conn_str)
            conn = pymssql.connect(server=db_host, port=db_port, user=db_user, password=db_password, database=db_database)
            df_select = pd.read_sql(query,conn)
            #conn.dispose()
            conn.close()
            processData(df_select, client_name,start_date,end_date,month_name,year)
            print(f"Success. Documents created for {client_name} {month_name} {year}. \n")
    window.close()

def processData(df_select,client_name,start,end,month_name,year):
    


    
    # Getting all the required libraries
    df_select["CreateDateTime"] = pd.to_datetime(df_select["CreateDateTime"]) # Converting date values to pandas Datetime
    df_select["lupdate"] = pd.to_datetime(df_select["lupdate"])
    df_sorted = df_select.sort_values(by="CreateDateTime", ascending=False) # Sorting by time ticket was created
    df_sorted.drop(df_sorted.loc[df_sorted.Subject == ''].index,inplace=True) # Removing any entries with an empty subject (e.g. replies to threads)
    
    df_support = df_sorted.loc[df_sorted["GType"] == "S"]
    df_project = df_sorted.loc[df_sorted["GType"] == "P"]
    
    close = df_sorted.loc[df_sorted["Status"] == "C"]
    close_support = close.loc[df_sorted["GType"] == "S"]
    close_project = close.loc[df_sorted["GType"] == "P"]
    
    date_string = parser.parse(start)
    next_string = parser.parse(end)
    
    
    month = df_sorted.loc[(df_sorted["CreateDateTime"] >= date_string) & (df_sorted["CreateDateTime"] < next_string)]
    
    month_support = month.loc[month["GType"] == "S"]
    month_project = month.loc[month["GType"] == "P"]    
    
    # Creating an excel with multiple worksheets
    writer = pd.ExcelWriter(f'{client_name}_{month_name}_{year}.xlsx', engine='xlsxwriter')
    workbook = writer.book
    
    


    # Pelaporan Project & Support
    
    s_count = month_support["Type"].count()
    p_count = month_project["Type"].count()
    s_out = df_support["Type"].count()
    p_out = df_project["Type"].count()
    s_close = close_support["Type"].count()
    p_close = close_project["Type"].count()
    t_close = close["Type"].count()
    
    df_info = pd.DataFrame()
    df_info["Type"] = ["Project", "Support", "Total"]
    df_info[f"{month_name} {year}"] = [p_count, s_count, p_count+s_count]
    df_info[f"Outstanding s.d. {month_name} {year}"] = [p_out, s_out, p_out+s_out]
    df_info["Close"] = [p_close, s_close, t_close]
    
    worksheet = workbook.add_worksheet("Pelaporan Project & Support")
    writer.sheets["Pelaporan Project & Support"] = worksheet
    df_info.to_excel(writer, sheet_name="Pelaporan Project & Support", startcol=0, startrow=4) # Data
    title = writer.sheets['Pelaporan Project & Support']
    title.write_string(0,0,f'Pelaporan Project & Support {client_name} {month_name} {year}') # Title
    
    


    
    # Outstanding Support
    s_ana = df_support.loc[df_support["Status"] == "A"]["CreateDateTime"].count() # In Analyzed
    s_ud = df_support.loc[df_support["Status"] == "D"]["CreateDateTime"].count() # Under development
    s_nsd = df_support.loc[df_support["Status"] == "W"]["CreateDateTime"].count() # Need support data
    s_conf = df_support.loc[df_support["Status"] == "F"]["CreateDateTime"].count() # Wating confirmation
    s_uat = df_support.loc[df_support["Status"] == "U"]["CreateDateTime"].count() # UAT
    s_close = df_support.loc[df_support["Status"] == "C"]["CreateDateTime"].count() # Closed
    s_total_out = df_support["CreateDateTime"].count()
    
    x_out = ["In Analyzed", "Under development", "Need support data", "Waiting confirmation", "UAT", "Closed", "Total"]
    x_out_graph = ["In Analyzed", "Under development", "Need support data", "Waiting confirmation", "UAT", "Closed"]
    y_s_out = [s_ana, s_ud, s_nsd, s_conf, s_uat, s_close, s_total_out]
    
    df_s_out = pd.DataFrame()
    
    df_s_out["Status"] = x_out
    df_s_out["Amount"] = y_s_out
    
    worksheet = workbook.add_worksheet("Outstanding Support")
    writer.sheets["Outstanding Support"] = worksheet
    df_s_out.to_excel(writer, sheet_name="Outstanding Support", startcol=0, startrow=3) # Data
    title = writer.sheets['Outstanding Support']
    title.write_string(0,0,f'Outstanding Support periode {month_name} {year}') # Title
    
    fig, ax = subplots(figsize=(11,11)) # Graph
    df_s_out.drop(df_s_out[df_s_out["Status"] == "Total"].index, inplace=True)
    barplot(df_s_out["Status"],df_s_out["Amount"], ax=ax)
    ax.set_xticklabels(x_out_graph,rotation=45)
    ax.set_ylim(bottom=0)
    for i, value in enumerate(df_s_out["Amount"]):
        ax.text(i, value, str(value), ha='center', va='bottom', fontsize=15, weight='bold')
        
    ax.set_title(f"Outstanding Project Periode {month_name} {year}")
    imgdata=io.BytesIO()
    fig.savefig(imgdata, format='png')
    worksheet.insert_image(2,6, '', {'image_data': imgdata, 'x_scale': 0.5, 'y_scale': 0.5})
    
    
    # Status Close Support
    s_ep = close_support.loc[close_support["Type"] == "E"]["CreateDateTime"].count() # Error Program
    s_mod = close_support.loc[close_support["Type"] == "M"]["CreateDateTime"].count() # Modification
    s_cons = close_support.loc[close_support["Type"] == "C"]["CreateDateTime"].count() # Consultancy
    s_ue = close_support.loc[close_support["Type"] == "U"]["CreateDateTime"].count() # User Error
    s_enh = close_support.loc[close_support["Type"] == "H"]["CreateDateTime"].count() # Enhancement

    x_close = ["Error Program", "Modification", "Consultancy", "User Error", "Enhancement"]
    y_s_close = [s_ep,s_mod,s_cons,s_ue,s_enh]
    df_s_close = pd.DataFrame()
    
    df_s_close["Status"] = x_close
    df_s_close["Amount"] = y_s_close
    
    worksheet = workbook.add_worksheet ("Status Support Close")
    writer.sheets["Status Support Close"] = worksheet
    df_s_close.to_excel(writer, sheet_name="Status Support Close", startcol=0, startrow=3) # Data
    title = writer.sheets['Status Support Close']
    title.write_string(0,0,f'Status Support CLOSE Periode {month_name} {year}') # Title
    
    fig, ax = subplots(figsize=(11,11)) # Graph
    barplot(df_s_close["Status"], df_s_close["Amount"], ax=ax)
    ax.set_xticklabels(x_close,rotation=45)
    for i, value in enumerate(df_s_close["Amount"]):
        ax.text(i, value, str(value), ha='center', va='bottom', fontsize=15, weight='bold')
    ax.set_ylim(bottom=0)
        
    ax.set_title(f"Status Close Support Periode {month_name} {year}")
    imgdata=io.BytesIO()
    fig.savefig(imgdata, format='png')
    worksheet.insert_image(2,6, '', {'image_data': imgdata, 'x_scale': 0.5, 'y_scale': 0.5})




    # Outstanding Project (Request)
    p_ana = df_project.loc[df_project["Status"] == "A"]["CreateDateTime"].count() # In Analyzed
    p_ud = df_project.loc[df_project["Status"] == "D"]["CreateDateTime"].count() # Under development
    p_nsd = df_project.loc[df_project["Status"] == "W"]["CreateDateTime"].count() # Need support data
    p_conf = df_project.loc[df_project["Status"] == "F"]["CreateDateTime"].count() # Wating confirmation
    p_uat = df_project.loc[df_project["Status"] == "U"]["CreateDateTime"].count() # UAT
    p_close = df_project.loc[df_project["Status"] == "C"]["CreateDateTime"].count() # Closed
    p_total_out = df_project["CreateDateTime"].count()
    
    y_p_out = [p_ana, p_ud, p_nsd, p_conf, p_uat, p_close, p_total_out]
    
    df_p_out = pd.DataFrame()
    
    df_p_out["Status"] = x_out
    df_p_out["Amount"] = y_p_out
    
    worksheet = workbook.add_worksheet("Outstanding Project")
    writer.sheets["Outstanding Project"] = worksheet
    df_p_out.to_excel(writer, sheet_name="Outstanding Project", startcol=0, startrow=3) # Data
    title = writer.sheets["Outstanding Project"]
    title.write_string(0,0,f'Outstanding Project periode {month_name} {year}')
    
    fig, ax = subplots(figsize=(11,11)) # Graph
    df_p_out.drop(df_p_out[df_p_out["Status"] == "Total"].index, inplace=True)
    barplot(df_p_out["Status"],df_p_out["Amount"], ax=ax)
    ax.set_xticklabels(x_out_graph,rotation=45)
    for i, value in enumerate(df_p_out["Amount"]):
        ax.text(i, value, str(value), ha='center', va='bottom', fontsize=15, weight='bold')
    ax.set_ylim(bottom=0)

    ax.set_title(f"Outstanding Project Periode {month_name} {year}")
    imgdata=io.BytesIO()
    fig.savefig(imgdata, format='png')
    worksheet.insert_image(2,6, '', {'image_data': imgdata, 'x_scale': 0.5, 'y_scale': 0.5})
       

    # Status close project
    
    p_ep = close_project.loc[close_project["Type"] == "E"]["CreateDateTime"].count() # Error Program
    p_mod = close_project.loc[close_project["Type"] == "M"]["CreateDateTime"].count() # Modification
    p_cons = close_project.loc[close_project["Type"] == "C"]["CreateDateTime"].count() # Consultancy
    p_ue = close_project.loc[close_project["Type"] == "U"]["CreateDateTime"].count() # User Error
    p_enh = close_project.loc[close_project["Type"] == "H"]["CreateDateTime"].count() # Enhancement

    y_p_close = [p_ep,p_mod,p_cons,p_ue,p_enh]
    
    df_p_close = pd.DataFrame()
    
    df_p_close["Status"] = x_close
    df_p_close["Amount"] = y_p_close
    
    worksheet = workbook.add_worksheet("Status Close Project")
    writer.sheets["Status Close Project"] = worksheet
    df_p_close.to_excel(writer, sheet_name="Status Close Project", startcol=0, startrow=3) # Data
    title = writer.sheets["Status Close Project"]
    title.write_string(0,0,f'Status Project(Request) Close s.d. {month_name} {year}')
    
    fig, ax = subplots(figsize=(11,11)) # Graph
    barplot(df_p_close["Status"], df_p_close["Amount"], ax=ax)
    ax.set_xticklabels(x_close,rotation=45)
    for i, value in enumerate(df_p_close["Amount"]):
        ax.text(i, value, str(value), ha='center', va='bottom', fontsize=15, weight='bold')
    ax.set_ylim(bottom=0)
    
    ax.set_title(f"Status Close Project Periode {month_name} {year}")
    imgdata=io.BytesIO()
    fig.savefig(imgdata, format='png')
    worksheet.insert_image(2,6, '', {'image_data': imgdata, 'x_scale': 0.5, 'y_scale': 0.5})




    # Trend Tiket Bedasarkan Waktu (Support)
    month_support_copy = month_support.copy(deep=True)
    month_support_copy["CreateDateTime"] = pd.to_datetime(month_support_copy["CreateDateTime"]).dt.date   
    count_s = month_support_copy["CreateDateTime"].value_counts().reset_index()
    count_s.rename(columns={"index": "Tanggal Pelaporan", "CreateDateTime": "Support"}, inplace=True)
    count_s = count_s.sort_values(by="Tanggal Pelaporan", ascending=True) # Sorting by time ticket was created

    
    worksheet =  workbook.add_worksheet("Trend Ticket Support")
    writer.sheets["Trend Tiket Support"] = worksheet
    count_s.to_excel(writer,sheet_name="Trend Tiket Support", startcol=0, startrow=3) # Data
    title = writer.sheets["Trend Tiket Support"]
    title.write_string(0,0,f'Trend Tiket Support Berdasarkan Waktu Pengajuan Periode {month_name} {year}')
    
    
    # Trend Tiket Berdasarkan Waktu (Project)
    month_project_copy = month_project.copy(deep=True)
    month_project_copy["CreateDateTime"] = pd.to_datetime(month_project_copy["CreateDateTime"]).dt.date
    count_p = month_project_copy["CreateDateTime"].value_counts().reset_index()
    count_p.rename(columns={"index": "Tanggal Pelaporan", "CreateDateTime": "Project"}, inplace=True)
    count_p = count_p.sort_values(by="Tanggal Pelaporan", ascending=True) # Sorting by time ticket was created
    
    worksheet =  workbook.add_worksheet("Trend Tiket Project")
    writer.sheets["Trend Tiket Project"] = worksheet
    count_p.to_excel(writer,sheet_name="Trend Tiket Project", startcol=0, startrow=3) # Data
    title = writer.sheets["Trend Tiket Project"]
    title.write_string(0,0,f'Trend Tiket Project Berdasarkan Waktu Pengajuan Periode {month_name} {year}')
    
    

    # All support sheet
    df_support_show = df_support.copy(deep=True)
    df_support_show = convert_values(df_support_show)
    
    worksheet = workbook.add_worksheet("All Support")
    writer.sheets["All Support"] = worksheet
    df_support_show.to_excel(writer, sheet_name="All Support", startcol=0, startrow=3) # Data
    title = writer.sheets['All Support']
    title.write_string(0,0,f'All status Support Periode {month_name} {year}') # Title


    # All project
    df_project_show = df_project.copy(deep=True)
    df_project_show = convert_values(df_project_show)
    
    worksheet = workbook.add_worksheet("All Project")
    writer.sheets["All Project"] = worksheet
    df_project_show.to_excel(writer, sheet_name="All Project", startcol=0, startrow=3) # Data
    title = writer.sheets['All Project']
    title.write_string(0,0,f'All status Project Periode {month_name} {year}') # Title
    
    
    
    # Saving the excel sheet
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", category=FutureWarning)
        writer.save()
    


def convert_values(df): # Need to add more rules if more parameters are added.
    df["GType"] = df["GType"].replace(["S"], "Support")
    df["GType"] = df["GType"].replace(["P"], "Project")
    df["Product"] = df["Product"].replace(["1"], "WINCore")
    df["Product"] = df["Product"].replace(["2"], "BI Report")
    df["Product"] = df["Product"].replace(["3"], "ATM")
    df["Product"] = df["Product"].replace(["4"], "MAXAVA")
    df["Product"] = df["Product"].replace(["9"], "Jasa Lain")
    df["Status"] = df["Status"].replace(["A"], "In Analyzed")
    df["Status"] = df["Status"].replace(["C"], "Closed")
    df["Status"] = df["Status"].replace(["D"], "Under Development")
    df["Status"] = df["Status"].replace(["F"], "Waiting Confirmation")
    df["Status"] = df["Status"].replace(["N"], "New")
    df["Status"] = df["Status"].replace(["U"], "UAT")
    df["Status"] = df["Status"].replace(["W"], "Need Support Data")
    df["Type"] = df["Type"].replace(["C"], "Consultancy")
    df["Type"] = df["Type"].replace(["E"], "Error Program")
    df["Type"] = df["Type"].replace(["H"], "Enhancement")
    df["Type"] = df["Type"].replace(["M"], "Modification")
    df["Type"] = df["Type"].replace(["U"], "User Error")
    df = df.rename({"CreateDateTime": "Date Created", "RecordNo": "Record ID", "lupdate": "Latest Update", "GType": "Type", "Type": "Sub-Type"},axis='columns')
    #df = df.rename({"HName": "Helper Name"})
    return df

if __name__ == '__main__':
    main()