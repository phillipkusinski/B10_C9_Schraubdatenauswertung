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

#global variables
file_paths = []
save_path = 0
calendarweek = 0
year = 0
rob_nums = ["Rob_8_1", "Rob_8_2", "Rob_8_3", "Rob_9_1", "Rob_9_2", "Rob_9_3"]

#function definitions
def open_xlsx_files():
    global file_paths
    folder_paths = filedialog.askdirectory(
        title="Ordner auswählen mit XLSX-Dateien"
    )
    if not folder_paths:
        return

    #Select all files in the folders
    file_paths = []
    for root, dirs, files in os.walk(folder_paths):
        for file in files:
            if file.endswith(".xlsx"):
                file_paths.append(os.path.join(root, file))
    #failure message 
    if len(file_paths) > 32:
        messagebox.showwarning("Zu viele Dateien", "Bitte wählen Sie maximal 32 .xlsx-Dateien aus")
        return

    lbl_status.config(text=f"{len(file_paths)} Datei(en) gefunden")

def build_dataframe():
    global rob_nums
    global df
    global calendarweek
    list_of_df = []
    calendarweek_status = 0
    expected_columns = 10
    
    if len(file_paths) == 0:
        messagebox.showerror("Keine Daten ausgewählt", "Es wurden keine Daten zur Auswertung ausgewählt!")
        return
    else:
        for file in file_paths:
            try:
                df = pd.read_excel(file, usecols = [0, 2, 3, 4, 14, 15, 16, 17, 18, 19], header = None, skiprows = 2)
                
                if df.shape[1] != expected_columns:
                    raise ValueError(f"Datei '{os.path.basename(file)}' hat {df.shape[1]} Spalten, erwartet wurden {expected_columns}.")
                path_parts = os.path.normpath(file).split(os.sep)
                rob_num_extracted = next((part for part in path_parts if part.startswith("Rob_")), "Unbekannt")
                df["Roboternummer"] = rob_num_extracted
                list_of_df.append(df)
            
            except Exception as e:
                messagebox.showerror("Fehler beim Laden", f"❌ Datei konnte nicht verarbeitet")
                return 
        #concat all dfs with all stations
        df = pd.concat(list_of_df, ignore_index=True)
        header = ["Datum", "Programmnummer", "Fehlernummer", "Gesamtlaufzeit",
                "Schritt 3", "Drehmoment 3", "Drehwinkel 3", "Schritt NOK", 
                "Drehmoment NOK", "Drehwinkel NOK", "Roboternummer"]
        df.columns = header
        calendarweek_status = calendarweek_check()
        if calendarweek_status == 1:
            messagebox.showinfo("Datenstruktur erfolgreich", f"Es wurde erfolgreich die Datenstruktur der KW{calendarweek} aufgebaut")
        else:
            messagebox.showerror("Fehler beim Aufbau der Datenstruktur", "Es konnte keine Datenstruktur aufgebaut werden, da die Datensätze nicht aus der selben Kalenderwoche sind!")
            df = 0
            calendarweek = 0

def calendarweek_check():
    global df
    global calendarweek
    global year
    df['Datum'] = pd.to_datetime(df['Datum'])
    iso = df["Datum"].dt.isocalendar()
    if iso['week'].nunique() == 1 and iso['year'].nunique() == 1:
        calendarweek = iso['week'].iloc[0]
        year = iso['year'].iloc[0]
        return 1
    else:
        return 0
    
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
    list_of_df_daily = []
    list_of_df_weekly = []
    list_of_plots = []
    list_of_variants = ["B10", "C9"]
    if save_path and calendarweek != 0:
        for variant in list_of_variants:
            if variant == "B10":
                df_filtered = df[df["Programmnummer"] < 100]
                fig = create_failure_plot(df_filtered, variant)        
                df_grouped_detailed = create_detailed_dataframe(df_filtered)
                df_grouped_detailed_weekly = create_detailed_dataframe_weekly(df_filtered)
                list_of_plots.append(fig)
                list_of_df_daily.append(df_grouped_detailed)
                list_of_df_weekly.append(df_grouped_detailed_weekly)
            else:
                df_filtered = df[df["Programmnummer"] >= 100]
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
    df_failure_plot = (df_filtered.groupby(["Datum", "Roboternummer"], group_keys = False)
    .apply(lambda df_lambda: (df_lambda["Fehlernummer"] != 0).sum() / len(df_lambda) * 100)
    .reset_index(name="Fehleranteil in %")
    )

    df_failure_plot["Datum"] = df_failure_plot["Datum"].dt.date

    pivot_df = df_failure_plot.pivot(index="Datum", columns="Roboternummer", values="Fehleranteil in %")

    weekly_failure = (
        df_filtered.groupby("Roboternummer")
        .apply(lambda x: (x["Fehlernummer"] != 0).sum() / len(x) * 100)
        .round(2)
    )

    pivot_df.loc["Ø Woche"] = weekly_failure

    ax = pivot_df.plot(kind="bar", figsize=(12, 6))
    plt.axhline(0.2, color='red', linestyle='--', linewidth = 2)
    plt.ylabel("Fehleranteil in %")
    plt.title(f"Variante = {variant}, Kalenderwoche = {calendarweek}, Absoluter Fehleranteil in % pro Roboter")
    plt.xticks(rotation=0)
    plt.legend(title="Roboternummer", framealpha = 1)

    sep_index = len(pivot_df) - 2
    plt.axvline(x=sep_index + 0.5, color="gray", linestyle="--", linewidth=1)

    plt.tight_layout()
    fig = ax.figure
    return fig

def create_detailed_dataframe(df_filtered):
    df_grouped_detailed = df_filtered.groupby([df_filtered["Datum"].dt.date, "Roboternummer", "Fehlernummer"]).size().unstack(fill_value=0)
    df_grouped_detailed["Gesamtverschraubungen"] = df_grouped_detailed.sum(axis=1)
    fail_cols = [col for col in df_grouped_detailed.columns if col not in [0, "Gesamtverschraubungen"]]
    df_grouped_detailed["Fehler in %"] = (df_grouped_detailed[fail_cols].sum(axis=1) / df_grouped_detailed["Gesamtverschraubungen"] * 100).round(2)
    return df_grouped_detailed

def create_detailed_dataframe_weekly(df_filtered):
    df_grouped_detailed_weekly = df_filtered.groupby(["Roboternummer", "Fehlernummer"]).size().unstack(fill_value=0)
    df_grouped_detailed_weekly["Gesamtverschraubungen"] = df_grouped_detailed_weekly.sum(axis=1)
    fail_cols = [col for col in df_grouped_detailed_weekly.columns if col not in [0, "Gesamtverschraubungen"]]
    df_grouped_detailed_weekly["Fehler in %"] = (df_grouped_detailed_weekly[fail_cols].sum(axis=1) / df_grouped_detailed_weekly["Gesamtverschraubungen"] * 100).round(2)
    return df_grouped_detailed_weekly

def create_export(list_of_df_daily, list_of_df_weekly, list_of_plots):
    df_daily_b10 = list_of_df_daily[0]
    df_daily_c9 = list_of_df_daily[1]
    df_weekly_b10 = list_of_df_weekly[0]
    df_weekly_c9 = list_of_df_weekly[1]
    
    save_name = f"{save_path}/Schraubreport_KW{calendarweek}_{year}.xlsx"
    with pd.ExcelWriter(save_name) as writer:
        df_daily_b10.to_excel(writer, sheet_name = "B10 daily")
        df_weekly_b10.to_excel(writer, sheet_name = "B10 weekly")
        df_daily_c9.to_excel(writer, sheet_name = "C9 daily")
        df_weekly_c9.to_excel(writer, sheet_name = "C9 weekly")

        workbook = writer.book
        sheet_names = ["B10 weekly", "C9 weekly"]
        target_cells = ["A7", "A7"]  
        for fig, sheet, cell in zip(list_of_plots, sheet_names, target_cells):
            image_stream = BytesIO()
            fig.savefig(image_stream, format='png', dpi=300, bbox_inches='tight')
            image_stream.seek(0)

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
    root.geometry("320x300")
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

    btn_load_xlsx = ttk.Button(frame_xlsx,
                            text="📂 xlsx-Datei öffnen",
                            command=open_xlsx_files)
    btn_load_xlsx.grid(row=0, column=0, sticky="ew")

    lbl_status = ttk.Label(frame_xlsx,
                        text="0 Dateien ausgewählt")
    lbl_status.grid(row=0, column=1, sticky="w", padx=(20, 0))

    btn_submit_xlsx = ttk.Button(
        frame_xlsx,
        text="Erstelle Datenstruktur",
        command= build_dataframe       #build_dataframe
    )
    btn_submit_xlsx.grid(row=1, column=0, columnspan = 2, sticky="ew", pady=10)

    #Separator
    ttk.Separator(root, orient="horizontal") \
        .grid(row=1, column=0, sticky="ew", pady=15)

    btn_select_path = ttk.Button(root,
                            text="📂 Speicherpfad auswählen",
                            command=select_save_path)
    btn_select_path.grid(row=6, column=0, sticky="ew")

    #Export
    btn_export = ttk.Button(root,
                            text="Export starten",
                            command=main_filter_func,
                            style="Export.TButton")
    btn_export.grid(row=7, column=0, pady=20, sticky="ew")

    #Separator
    ttk.Separator(root, orient="horizontal") \
        .grid(row=8, column=0, sticky="ew", pady=15)

    #Author + Version
    lbl_version = ttk.Label(root,
                            text="Phillip Kusinski, V1.0",
                            style="TLabel") 
    lbl_version.grid(row=9, column=0, sticky="e")

    root.mainloop()