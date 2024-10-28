#! python3
# r: pandas, openpyxl

import Rhino
import rhinoscriptsyntax as rs
import scriptcontext as sc
import math
import System

import locale

locale.setlocale(locale.LC_ALL, 'nb_NO')

import pandas as pd
import openpyxl
import os


def excel_import(excel):
    berg = ["F1", "F2", "F3", "F4", "F5", "F6"]
    spr = ["S1", "S2", "S3", "S4", "S5", "S6"]
    bolt = ['B1', 'B2', 'B3', 'B4', 'B1A', 'B1B', 'B1C', 'B1D', 'B1E']
    mangler = ['M1', 'M2', 'M3', 'M4', 'M5']
    fremkom = ["I/X", "X", "L", "D"]

    if os.path.isfile(excel):
        df_obs = pd.read_excel(excel, header=0)

        print(df_obs['Kommentar'])
        T_scale = 200
        vl_max = 110 / 500 * T_scale

    for index, row in df_obs.iterrows():

        try:
            pel = float(row['Pel'])
            pos = row['Pos']

            
            if pos > -100 or pos < 100:
                pos = float(pos) * vl_max / 100
            else: 
                print('Pos > 100 eller <-100:', pos)
                pos = vl_max + 10
                
            if row['Symbol'] in berg:
                rs.CurrentLayer('Berg')
                rs.AddText(row['Symbol'], (pos, pel), height=2, justification=2)
                rs.AddText(row['Kommentar'], (pos, pel - 2), height=1.5, justification=2)

            elif row['Symbol'] in spr:
                rs.CurrentLayer('SPR')
                rs.AddText(row['Symbol'], (pos, pel), height=2, justification=2)
                rs.AddText(row['Kommentar'], (pos, pel - 2), height=1.5, justification=2)

            elif row['Symbol'] in bolt:
                rs.CurrentLayer('Bolt')
                rs.AddText(row['Symbol'], (pos, pel), height=2, justification=2)
                rs.AddText(row['Kommentar'], (pos, pel - 2), height=1.5, justification=2)

            elif row['Symbol'] in mangler:
                rs.CurrentLayer('Manglende utført bergsikring')
                rs.AddText(row['Symbol'], (pos, pel), height=2, justification=2)
                rs.AddText(row['Kommentar'], (pos, pel - 2), height=1.5, justification=2)

            elif row['Symbol'] in fremkom:
                rs.CurrentLayer('Fremkommlighet')
                rs.AddText(row['Symbol'], (pos, pel), height=2, justification=2)
                rs.AddText(row['Kommentar'], (pos, pel - 2), height=1.5, justification=2)

            else:
                rs.CurrentLayer('Plasser manuelt')
                rs.AddText(row['Symbol'], (pos, pel), height=2, justification=2)
                rs.AddText(row['Kommentar'], (pos, pel - 2), height=1.5, justification=2)

        except:
            print("Feilmelding - Ugyldig verdi: ", row['Pel'], row['Kommentar'])


if __name__ == "__main__":
    excel = rs.GetString('PATH til Excel-fil')
    excel_import(excel)
    rs.CurrentLayer("Default")
    view = rs.CurrentView('TOP')
    rs.IsViewMaximized(view)
    rs.ZoomExtents()
