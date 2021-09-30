import time
import os
import re
from tkinter import *
from tkinter import ttk

from pathlib import Path
import win32com.client as win32
import openpyxl
from openpyxl import Workbook, load_workbook
import pandas as pd
import xlwings as xw
import tempfile
import shutil
from helper import next_available_name
from last_modife import LastModifeFinder
from facture_to_table import AnalyzeForDataFrame, AnalyzeMaterialTable, InfoFetcher

# from xl_to_pdf_classes import AnalyzeWB, AnalyzeFactureSheet, AnalyzeAnnexeSheet

class TabXltoTable:
    def __init__(self, *args, **kwargs):
        self.users_facture_path = ''
        self.users_material_table_path = ''
        try:
            with open(Path('users_selections') / 'facture_path.txt', 'r') as f:
                self.users_facture_path = f.readline()
            with open(Path('users_selections') / 'material_table_path.txt', 'r') as f:
                self.users_material_table_path = f.readline()
        except:
            pass
        
        self.tab = kwargs['tab']

        # Address Selection selection
        self.all_factures_path_value=StringVar()
        self.material_table_path_value=StringVar()

        # Selection of All Factures Path
        self.all_factures_path = ttk.Entry(
            self.tab, width=50, font='Helvetica 12', textvariable=self.all_factures_path_value)
        self.all_factures_path.insert(INSERT, Path(self.users_facture_path))
        self.all_factures_path.grid(
            row=1, column=1, sticky=W, padx=60, pady=20, ipadx=20, ipady=10)
        self.browse_all_factures_path = ttk.Button(
            self.tab, width=20, text="Select the path of factures you want to add.", style='XltopdfBrowse.TButton')
        self.browse_all_factures_path.grid(
            row=1, column=2, columnspan=2, padx=5, ipadx=10, ipady=5)

        # Material Table Selection
        self.material_table_path = ttk.Entry(self.tab, width=50, font='Helvetica 12', textvariable=self.material_table_path_value)
        self.material_table_path.insert(INSERT, Path(self.users_material_table_path))
        self.material_table_path.grid(row=2, column=1, sticky=W,
                                padx=60, pady=20, ipadx=20, ipady=10)
        self.browse_material_table_path = ttk.Button(
            self.tab, width=20, text="Please Select the material table that you want to update.", style='XltopdfBrowse.TButton')
        self.browse_material_table_path.grid(
            row=2, column=2, columnspan=2, padx=5, ipadx=10, ipady=5)

        # Transfer Button
        self.transfer_button = ttk.Button(
            self.tab, text='Transfer', style='XltopdfConvertButton.TButton')
        self.transfer_button.grid(row=int(self.material_table_path.grid_info()[
                                 'row'])+1, column=2, columnspan=2, sticky=E, padx=10, pady=10, ipadx=50, ipady=20)
        self.transfer_button.config(command=self.transfer_files)

    def transfer_files(self):
        # Save selected paths to user_selection
        with open(Path('users_selections') / 'facture_path.txt', 'w') as f:
            f.write(self.all_factures_path_value.get())
        with open(Path('users_selections') / 'material_table_path.txt', 'w') as f:
            f.write(self.material_table_path_value.get())
        start=time.perf_counter()

        all_factures_path=Path(self.all_factures_path_value.get())
        material_table_path=Path(self.material_table_path_value.get())

        hungry_man = InfoFetcher(material_table=str(material_table_path), all_factures_path=all_factures_path)
        hungry_man.create_result_wb_for_result()

        for excel in hungry_man.list_combined_path_of_factures_to_transfer:
            if (excel.suffix).lower == 'xls':
                print('Skipped xls file: ', excel)
                continue

            try:
                testanalyze = AnalyzeForDataFrame(
                    type='df', wb=pd.ExcelFile(excel), extension=excel.suffix)
            except:
                print('testanalyze failed: ', excel)
                pass

            if testanalyze.is_facture:
                print("Started memorizing: ", excel.name)
                hungry_man.eat_facture(df_facture=testanalyze.get_facture_df(), df_annexe=testanalyze.get_annexe_df(), current_excel_file_name=excel.name)

        hungry_man.save_result_wb_after_done()
        hungry_man.transfer_to_material_table()
        finish=time.perf_counter()
        print('Total time spent {} seconds'.format(round(finish-start, 3)))
