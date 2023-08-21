# -*- coding: utf-8 -*-
"""
Created on Fri Aug 04 16:16:14 2023

@author: Long_S
"""

import tkinter as tk
from tkinter import ttk
import pandas as pd
import time
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg, NavigationToolbar2Tk)
import numpy as np
import table_functions as tf
import sqlalchemy as sa
from docx import Document
import os


class Tables(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        
        # Create industry group selection combobox
        self.group_label = ttk.Label(self, text = 'Industry Group: ')
        self.group_label.grid(row = 0, column = 0, pady = (50, 10), padx = (50, 10))

        self.group_selection = tk.StringVar(value = '')
        
        self.group_combobox = ttk.Combobox(self, 
                                           textvariable = self.group_selection, 
                                           width = 50,
                                           value = controller.industry_groups,
                                           state = 'readonly')
        self.group_combobox.grid(row = 0, column = 1, columnspan = 2, pady = (50, 10), padx = (20, 10), sticky = 'w')
        self.group_combobox.set(self.group_selection.get())
        
        # Digit options combobox
        self.digit_label = ttk.Label(self, text = 'Industry Digit: ')
        self.digit_label.grid(row = 5, column = 0, pady = 10, padx = (50, 10))

        self.digit_selection = tk.StringVar(value = '')
        
        self.digit_combobox = ttk.Combobox(self, 
                                           textvariable = self.digit_selection, 
                                           width = 50,
                                           value = controller.digit_options,
                                           state = 'readonly')
        self.digit_combobox.grid(row = 5, column = 1, columnspan = 2, pady = 10, padx = (20, 10), sticky = 'w')
        self.digit_combobox.set(self.digit_selection.get())

        # Data year
        self.start_label = ttk.Label(self, text = 'Year: ')
        self.start_label.grid(row = 2, column = 0, pady = 10, padx = (50, 10))
        
        self.start = tk.IntVar(value = controller.end_year)
        self.start_spinbox = ttk.Spinbox(self,
                                         from_= controller.start_year,
                                         to = controller.end_year, 
                                         textvariable = self.start, 
                                         wrap = False,
                                         state = 'readonly')
        
        self.start_spinbox.grid(row = 2, column = 1, padx = (20,10), pady= 10, sticky = 'w')
        self.start_spinbox.set(self.start.get())

        # Abbreviation
        self.name_label = ttk.Label(self, text = 'Abbreviated variable names: ')
        self.name_label.grid(row = 1, column = 0, pady = 10, padx = (50, 10))
        
        self.abbreviated = tk.BooleanVar(value = False)
        
        self.noabbrev_radiobutton = ttk.Radiobutton(self,
                                                   text = 'Non-abbreviated',
                                                   variable = self.abbreviated,
                                                   value = False,
                                                   command = self.grey_out)
        self.noabbrev_radiobutton.grid(row = 1, column = 1, padx = (20,0), pady= 10, sticky = 'w')
        
        self.abbrev_radiobutton = ttk.Radiobutton(self,
                                                   text = 'Abbreviated',
                                                   variable = self.abbreviated,
                                                   value = True,
                                                   command = self.grey_out)
        self.abbrev_radiobutton.grid(row = 1, column = 2, padx = (10,0), pady= 10, sticky = 'w')
        
        # File path insert
        self.path_label = ttk.Label(self, text = 'Enter filepath for export: ')
        self.path_label.grid(row = 7, column = 0, pady = 10, padx = (50,10))
        
        self.filepath = tk.StringVar()
        self.filepath_entry = ttk.Entry(self, textvariable = self.filepath, width = 100)
        self.filepath_entry.grid(row = 7, column = 1, columnspan = 3, padx = (20,10), pady= 10, sticky = 'w')
        
        # Create graphs button
        self.create_button = ttk.Button(self, text = 'Create Table', width = 20, command = lambda: self.create_graphs(self.group_selection.get(), \
                                                                                                                      self.digit_selection.get(),\
                                                                                                                      self.start.get(),\
                                                                                                                      self.abbreviated.get(),\
                                                                                                                      self.filepath.get()))
        self.create_button.grid(row = 8, column = 1, pady = 30)
        
        # Progress label
        self.progress = ttk.Label(self, text = '')
        self.progress.grid(row = 9, column = 1, padx = (20,10), pady = 10, sticky = 'w')

        self.explanation = ttk.Label(self, text = 'In order to get the filepath for export, open File Explorer and copy/paste the path. Be sure to put it in quotation marks.')
        self.explanation.grid(row = 10, column = 1, padx = (20,10), pady = 10, sticky = 'w')
        
        
    def create_graphs(self, group, digit, yr, abbreviated, export):

        self.progress['text'] = 'Exporting file... '
        self.update_idletasks()
        tf.data_table(group, digit, yr, abbreviated, export)
        self.progress['text'] = 'Export complete'
        self.update_idletasks()

    def grey_out(self):
        
        if self.periods.get():
            self.start_spinbox.configure(state = 'disabled')
            self.end_spinbox.configure(state = 'disabled')
            
        else:
            self.start_spinbox.configure(state = 'readonly')
            self.end_spinbox.configure(state = 'readonly')