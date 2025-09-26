"""
Author: Phillip Kusinski
GUI tool for analyzing and exporting screw assembly data for Audi B10/C9 production reports
"""

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
import os
from io import BytesIO

#initialize global variables
file_paths = []
save_path = 0
calendarweek = 0
year = 0
rob_nums = ["Rob_8_1", "Rob_8_2", "Rob_8_3", "Rob_9_1", "Rob_9_2", "Rob_9_3"]
variant = 0

#function definitions
def open_xlsx_files():
    global file_paths
    #open folder with .askdirectory
    folder_paths = filedialog.askdirectory(
        title="Ordner auswählen mit XLSX-Dateien"
    )
    #return if no folder was selected
    if not folder_paths:
        return

    #Select all files in the folders
    file_paths = []
    for root, dirs, files in os.walk(folder_paths):
        for file in files:
            if file.endswith(".xlsx"):
                file_paths.append(os.path.join(root, file))
    #failure message if more than one whole possible week was selected
    #5 robs * 7 days = 35 possible rawdata files
    if len(file_paths) > 35:
        messagebox.showwarning("Zu viele Dateien", "Bitte wählen Sie maximal 32 .xlsx-Dateien aus")
        return
    #update status
    lbl_status.config(text=f"{len(file_paths)} Datei(en) gefunden")

def build_dataframe():
    global rob_nums
    global df
    global calendarweek
    global front_back
    list_of_df = []
    calendarweek_status = 0
    expected_columns = 10
    
    #errormessage if no rawdata was selected
    if len(file_paths) == 0:
        messagebox.showerror("Keine Daten ausgewählt", "Es wurden keine Daten zur Auswertung ausgewählt!")
        return
    else:
        #check if its "Hintertür" or "Vordertür" 
        front_back = front_back_check()
        for file in file_paths:
            try:
                df = pd.read_excel(file, usecols = [0, 2, 3, 4, 14, 15, 16, 17, 18, 19], header = None, skiprows = 1)
                #loaded files must content 10 columns
                if df.shape[1] != expected_columns:
                    raise ValueError(f"Datei '{os.path.basename(file)}' hat {df.shape[1]} Spalten, erwartet wurden {expected_columns}.")
                path_parts = os.path.normpath(file).split(os.sep)
                rob_num_extracted = next((part for part in path_parts if part.startswith("Rob_")), "Unbekannt")
                df["Roboternummer"] = rob_num_extracted
                list_of_df.append(df)
                
            #Ffailure message if the readed file is not in correct shape
            except Exception as e:
                messagebox.showerror("Fehler beim Laden", f"❌ Datei konnte nicht verarbeitet")
                return 
        #concat all dfs with all stations
        df = pd.concat(list_of_df, ignore_index=True)
        header = ["Datum", "Programmnummer", "Fehlernummer", "Gesamtlaufzeit",
                "Schritt 3", "Drehmoment 3", "Drehwinkel 3", "Schritt NOK", 
                "Drehmoment NOK", "Drehwinkel NOK", "Roboternummer"]
        #set correct headers
        df.columns = header
        #check selected calendarweek
        #calendarweek_status == 1: First till last day is in the same cw
        #calendarweek_stauts != 0: Data is not consitently in the same cw
        calendarweek_status = calendarweek_check()
        
        if calendarweek_status == 1:
            messagebox.showinfo("Datenstruktur erfolgreich", f"Es wurde erfolgreich die Datenstruktur der Variante {front_back} der KW{calendarweek} aufgebaut")
        else:
            messagebox.showerror("Fehler beim Aufbau der Datenstruktur", "Es konnte keine Datenstruktur aufgebaut werden, da die Datensätze nicht aus der selben Kalenderwoche sind!")
            df = 0
            calendarweek = 0

def calendarweek_check():
    global df
    global calendarweek
    global year
    #select col Datum from df and set it to a datetime object
    df['Datum'] = pd.to_datetime(df['Datum'])
    #split df["Datum"] into iso with week and year
    iso = df["Datum"].dt.isocalendar()
    if iso['week'].nunique() == 1 and iso['year'].nunique() == 1:
        calendarweek = iso['week'].iloc[0]
        year = iso['year'].iloc[0]
        return 1
    else:
        return 0

def front_back_check():
    #select first file within file_paths (in file_paths must always be minimum of one file)
    path_parts = file_paths[0]
    path_parts = os.path.normpath(path_parts).split(os.sep)
    front_back_keywords = ["Hintertür", "Vordertür"]
    #check for keyword in the path
    front_back = next((part for part in path_parts if part in front_back_keywords), "Unbekannt")
    return front_back

def select_save_path():
    global save_path
    #get saveing directory from user input
    save_path = filedialog.askdirectory(
        title = "Ordner zur Abspeicherung der Prüfergebnisse auswählen."
    )
    #return if no save_path was selected
    if not save_path:
        return
    messagebox.showinfo("Ordnerwahl erfolgreich", "Es wurde erfolgreich ein Ordner zur Abspeicherung ausgewählt.")

def main_filter_func():
    #initialize needed lists
    list_of_df_daily = []
    list_of_df_weekly = []
    list_of_plots = []
    list_of_variants = ["B10", "C9"]
    #check if all data was set as needed
    if save_path and calendarweek and front_back != 0:
        
        #check for the front_back information that is in the folder path
        if front_back == "Vordertür":
            #iterate through the two production models
            for variant in list_of_variants:
                if variant == "B10":
                    #programnumbers specified by technician
                    df_filtered = df[df["Programmnummer"] < 111]
                    #create plots and dataframes and .append to correct list
                    fig = create_failure_plot(df_filtered, variant)        
                    df_grouped_detailed = create_detailed_dataframe(df_filtered)
                    df_grouped_detailed_weekly = create_detailed_dataframe_weekly(df_filtered)
                    list_of_plots.append(fig)
                    list_of_df_daily.append(df_grouped_detailed)
                    list_of_df_weekly.append(df_grouped_detailed_weekly)
                else:
                    #programnumbers specified by technician
                    df_filtered = df[df["Programmnummer"] >= 111]
                    #create plots and dataframes and .append to correct list
                    fig = create_failure_plot(df_filtered, variant)        
                    df_grouped_detailed = create_detailed_dataframe(df_filtered)
                    df_grouped_detailed_weekly = create_detailed_dataframe_weekly(df_filtered)
                    list_of_plots.append(fig)
                    list_of_df_daily.append(df_grouped_detailed)
                    list_of_df_weekly.append(df_grouped_detailed_weekly)
        elif front_back == "Hintertür":
            for variant in list_of_variants:
                #iterate through the two production models
                if variant == "B10":
                    #programnumbers specified by technician
                    #split df into two dfs to make varying filtering 
                    df_rob_8_2 = df[df["Roboternummer"] == "Rob_8_2"]
                    df_other_robs = df[df["Roboternummer"] != "Rob_8_2"]
                    df_rob_8_2 = df_rob_8_2[df_rob_8_2["Programmnummer"] < 20]
                    df_other_robs = df_other_robs[df_other_robs["Programmnummer"]< 111]
                    #concatenate filtered dataframes into one
                    df_filtered = pd.concat([df_rob_8_2, df_other_robs], ignore_index = True)
                    #create plots and dataframes and .append to correct list
                    fig = create_failure_plot(df_filtered, variant)        
                    df_grouped_detailed = create_detailed_dataframe(df_filtered)
                    df_grouped_detailed_weekly = create_detailed_dataframe_weekly(df_filtered)
                    list_of_plots.append(fig)
                    list_of_df_daily.append(df_grouped_detailed)
                    list_of_df_weekly.append(df_grouped_detailed_weekly)
                else:
                    #programnumbers specified by technician
                    #split df into two dfs to make varying filtering 
                    df_rob_8_2 = df[df["Roboternummer"] == "Rob_8_2"]
                    df_other_robs = df[df["Roboternummer"] != "Rob_8_2"]
                    df_rob_8_2 = df_rob_8_2[df_rob_8_2["Programmnummer"] > 20]
                    df_other_robs = df_other_robs[df_other_robs["Programmnummer"] > 111]
                    #concatenate filtered dataframes into one
                    df_filtered = pd.concat([df_rob_8_2, df_other_robs], ignore_index = True)
                    #create plots and dataframes and .append to correct list
                    fig = create_failure_plot(df_filtered, variant)        
                    df_grouped_detailed = create_detailed_dataframe(df_filtered)
                    df_grouped_detailed_weekly = create_detailed_dataframe_weekly(df_filtered)
                    list_of_plots.append(fig)
                    list_of_df_daily.append(df_grouped_detailed)
                    list_of_df_weekly.append(df_grouped_detailed_weekly)
        #plot and dataframe export
        create_export(list_of_df_daily, list_of_df_weekly, list_of_plots)
        messagebox.showinfo("Export erfolgreich", "Der Export wurde erfolgreich durchgeführt.")
    else:
        messagebox.showerror("Ungültige Angabe", "Es wurden nicht alle Parameter korrekt gesetzt um den Prozess zu starten.")
        
def create_failure_plot(df_filtered, variant):
    #1.         .groupby "Datum", "Roboternummer"; group_keys = False for next .apply func
    #2.         .apply: lambda func that iterates through set groupby filters and calculates the "Fehler in %" value of all days seperately per robot
    #3.         .reset_index: set index of new col "Fehleranteil in %"
    df_failure = (df_filtered.groupby(["Datum", "Roboternummer"], group_keys = False)
    .apply(lambda df_lambda: (df_lambda["Fehlernummer"] != 0).sum() / len(df_lambda) * 100)
    .reset_index(name="Fehleranteil in %")
    )

    #set date without timestamps
    df_failure["Datum"] = df_failure["Datum"].dt.date

    #pivot df into correct form
    pivot_df = df_failure.pivot(index="Datum", columns="Roboternummer", values="Fehleranteil in %")

    #calculate weekly failure
    #1.         .groupby: "Roboternummer": all dates will be set together per robot 
    #2.         .apply: lambda func that iterates through set groupby filters and calculates the "Fehler in %" value of the sum of all days per robot
    #3.         .round(2): for better visualization
    weekly_failure = (
        df_filtered.groupby("Roboternummer")
        .apply(lambda x: (x["Fehlernummer"] != 0).sum() / len(x) * 100)
        .round(2)
    )

    #set data for plot df
    pivot_df.loc["Ø Woche"] = weekly_failure

    #plot data
    ax = pivot_df.plot(kind="bar", figsize=(12, 6))
    plt.axhline(0.2, color='red', linestyle='--', linewidth = 2)
    plt.ylabel("Fehleranteil in %")
    plt.title(f"Variante = {variant} {front_back}, Kalenderwoche = {calendarweek}, Absoluter Fehleranteil in % pro Roboter")
    plt.xticks(rotation=0)
    plt.legend(title="Roboternummer", framealpha = 1)

    sep_index = len(pivot_df) - 2
    plt.axvline(x=sep_index + 0.5, color="gray", linestyle="--", linewidth=1)

    plt.tight_layout()
    fig = ax.figure
    return fig

def create_detailed_dataframe(df_filtered):
    #.groupby:group filtered dataframe into "Datum", "Roboternummer", "Fehlernummer"
    #.size(): counts number of entrys i nevery group
    #.unstack(): "Fehlernummer" will be changed to the different failure nums as a own col
    df_grouped_detailed = df_filtered.groupby([df_filtered["Datum"].dt.date, "Roboternummer", "Fehlernummer"]).size().unstack(fill_value=0)
    #sum num of all entrys of axis 1
    df_grouped_detailed["Gesamtverschraubungen"] = df_grouped_detailed.sum(axis=1)
    #get all the cols with failures except 0 and "Gesamtverschraubungen"
    fail_cols = [col for col in df_grouped_detailed.columns if col not in [0, "Gesamtverschraubungen"]]
    #Calculate "Fehler in %"
    df_grouped_detailed["Fehler in %"] = (df_grouped_detailed[fail_cols].sum(axis=1) / df_grouped_detailed["Gesamtverschraubungen"] * 100).round(2)
    return df_grouped_detailed

def create_detailed_dataframe_weekly(df_filtered):
    #.groupby:group filtered dataframe into "Roboternummer", "Fehlernummer"
    #.size(): counts number of entrys i nevery group
    #.unstack(): "Fehlernummer" will be changed to the different failure nums as a own col
    df_grouped_detailed_weekly = df_filtered.groupby(["Roboternummer", "Fehlernummer"]).size().unstack(fill_value=0)
    #sum num of all entrys of axis 1
    df_grouped_detailed_weekly["Gesamtverschraubungen"] = df_grouped_detailed_weekly.sum(axis=1)
    #get all the cols with failures except 0 and "Gesamtverschraubungen"
    fail_cols = [col for col in df_grouped_detailed_weekly.columns if col not in [0, "Gesamtverschraubungen"]]
    #Calculate "Fehler in %"
    df_grouped_detailed_weekly["Fehler in %"] = (df_grouped_detailed_weekly[fail_cols].sum(axis=1) / df_grouped_detailed_weekly["Gesamtverschraubungen"] * 100).round(2)
    return df_grouped_detailed_weekly

def create_export(list_of_df_daily, list_of_df_weekly, list_of_plots):
    #set sheet_names
    sheet_names_weekly = ["B10 weekly", "C9 weekly"]
    sheet_names_daily = ["B10 daily", "C9 daily"]
    
    #isolate data from lists
    df_daily_b10 = list_of_df_daily[0]
    df_daily_c9 = list_of_df_daily[1]
    df_weekly_b10 = list_of_df_weekly[0]
    df_weekly_c9 = list_of_df_weekly[1]
    
    #select save name for the file with correct path settings
    save_name = f"{save_path}/Schraubreport_{front_back}_KW{calendarweek}_{year}.xlsx"
    #start ExcelWriter engine
    with pd.ExcelWriter(save_name) as writer:
        df_daily_b10.to_excel(writer, sheet_name = sheet_names_daily[0])
        df_weekly_b10.to_excel(writer, sheet_name = sheet_names_weekly[0])
        df_daily_c9.to_excel(writer, sheet_name = sheet_names_daily[1])
        df_weekly_c9.to_excel(writer, sheet_name = sheet_names_weekly[1])

        #format seetings for failure percent coloring
        workbook = writer.book
        green_format = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
        red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
        yellow_format = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500'})

        #loop aggregation for coloring cells in col "Fehler in %" weekly
        for sheet_name, df in zip(sheet_names_weekly, list_of_df_weekly):
            worksheet = writer.sheets[sheet_name]  
            try:
                #code snippet created with AI
                col_index = df.columns.get_loc("Fehler in %")  
                excel_col_letter = chr(ord('A') + col_index + 1)   
                cell_range = f"{excel_col_letter}2:{excel_col_letter}6"
                worksheet.conditional_format(cell_range, {
                    'type': 'cell',
                    'criteria': '>=',
                    'value': 0.5,
                    'format': red_format
                })
                worksheet.conditional_format(cell_range, {
                    'type': 'cell',
                    'criteria': 'between',
                    'minimum': 0.2001,
                    'maximum': 0.4999,
                    'format': yellow_format
                })
                worksheet.conditional_format(cell_range, {
                    'type': 'cell',
                    'criteria': '<=',
                    'value': 0.2,
                    'format': green_format
                })
            except KeyError:
                print(f"Spalte 'Fehler in %' nicht gefunden in {sheet_name}")
            
        #loop aggregation for coloring cells in col "Fehler in %" daily
        for sheet_name, df in zip(sheet_names_daily, list_of_df_daily):
            worksheet = writer.sheets[sheet_name]
                
            try:
                #code snippet created with AI
                col_index = df.columns.get_loc("Fehler in %")  
                excel_col_letter = chr(ord('A') + col_index + 2)   
                row_start = 2
                row_end = len(df) + 1
                cell_range = f"{excel_col_letter}{row_start}:{excel_col_letter}{row_end}"
                worksheet.conditional_format(cell_range, {
                    'type': 'cell',
                    'criteria': '>=',
                    'value': 0.5,
                    'format': red_format
                })
                worksheet.conditional_format(cell_range, {
                    'type': 'cell',
                    'criteria': 'between',
                    'minimum': 0.2001,
                    'maximum': 0.4999,
                    'format': yellow_format
                })
                worksheet.conditional_format(cell_range, {
                    'type': 'cell',
                    'criteria': '<=',
                    'value': 0.2,
                    'format': green_format
                })
            except KeyError:
                print(f"Spalte 'Fehler in %' nicht gefunden in {sheet_name}")   
        #set target cells to save plot into the excel       
        target_cells_plot = ["A7", "A7"]  
        for fig, sheet, cell in zip(list_of_plots, sheet_names_weekly, target_cells_plot):
            #BytesIO that the image is buffered in the RAM rather than saved on the desktop
            image_stream = BytesIO()
            fig.savefig(image_stream, format='png', dpi=300, bbox_inches='tight')
            #reset RAM buffer
            image_stream.seek(0)

            #insert fig into selected worksheet
            worksheet = writer.sheets[sheet]
            worksheet.insert_image(cell, "", {
                "image_data": image_stream,
                "x_offset": 5,
                "y_offset": 5,
                "x_scale": 0.5,
                "y_scale": 0.5
            })
            
if __name__ == "__main__":  
    #Setup Main Window
    root = tk.Tk()
    root.title("B10/C9 Schraubauswertung")
    root.geometry("350x320")
    root.resizable(False, False)
    #iconbitmap does not work with .exe build without bigger changes
    #root.iconbitmap("ressources/logo_yf.ico")

    #global Padding
    root.configure(padx=20, pady=20, bg="#f0f0f0")

    #style config
    style = ttk.Style(root)
    style.theme_use("clam")
    style.configure("TFrame", background="#f0f0f0")
    style.configure("TLabel", background="#f0f0f0")
    style.configure("Export.TButton",
                    font=("Arial", 16, "bold"),
                    foreground="white",
                    background="#28a745")
    style.map("Export.TButton",
            background=[("active", "#1e7e34")],
            foreground=[("active", "white")])

    #xlsx import
    frame_xlsx = ttk.Frame(root)
    frame_xlsx.grid(row=0, column=0, sticky="ew")
    root.columnconfigure(0, weight=1)
    frame_xlsx.columnconfigure(0, weight=1)
    frame_xlsx.columnconfigure(1, weight=1)

    #load button for raw .xlsx screwdata
    btn_load_xlsx = ttk.Button(frame_xlsx,
                            text="📂 xlsx-Datei öffnen",
                            command=open_xlsx_files)
    btn_load_xlsx.grid(row=0, column=0, sticky="ew")

    #file status of how many where selected
    lbl_status = ttk.Label(frame_xlsx,
                        text="0 Dateien ausgewählt")
    lbl_status.grid(row=0, column=1, sticky="w", padx=(20, 0))

    #button for submitting the selected data
    btn_submit_xlsx = ttk.Button(
        frame_xlsx,
        text="Erstelle Datenstruktur",
        command= build_dataframe       
    )
    btn_submit_xlsx.grid(row=1, column=0, columnspan = 2, sticky="ew", pady=10)
    
    #Separator
    ttk.Separator(frame_xlsx, orient="horizontal") \
        .grid(row=5, column=0, sticky="ew", pady=15, columnspan = 2)

    #button for selecting the correct saving path of the report
    btn_select_path = ttk.Button(frame_xlsx,
                            text="📂 Speicherpfad auswählen",
                            command=select_save_path)
    btn_select_path.grid(row=6, column=0, sticky="ew", columnspan = 2)

    #Button for selecting the Export process 
    btn_export = ttk.Button(frame_xlsx,
                            text="Export starten",
                            command=main_filter_func,
                            style="Export.TButton")
    btn_export.grid(row=7, column=0, pady=20, sticky="ew", columnspan = 2)

    #Separator
    ttk.Separator(frame_xlsx, orient="horizontal") \
        .grid(row=8, column=0, sticky="ew", pady=15, columnspan = 2)

    #Author + Version
    lbl_version = ttk.Label(frame_xlsx,
                            text="Phillip Kusinski, V1.1",
                            style="TLabel") 
    lbl_version.grid(row=9, column=1, sticky="e")

    root.mainloop()