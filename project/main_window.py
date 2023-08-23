# -*- coding: utf-8 -*-
"""
Created on Wed May 31 09:40:32 2023

@author: McGuigan_M
Modified by: Long_S
"""

import tkinter as tk
from tkinter import ttk
import pandas as pd
import time
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg, NavigationToolbar2Tk)
import datetime
import numpy as np

import source_data as sd
from lp_graph import LPGraphs
from table import Tables


class MainWindow(ttk.Frame):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.tabControl = ttk.Notebook(self)
        
        # Create tabs
        self.tab1 = ttk.Frame(self.tabControl)
        self.tab2 = ttk.Frame(self.tabControl)
        self.tab3 = ttk.Frame(self.tabControl)
        self.tab4 = ttk.Frame(self.tabControl)
        self.tab5 = ttk.Frame(self.tabControl)
        self.tab6 = ttk.Frame(self.tabControl)
        
        # Define tabs
        self.tabControl.add(self.tab1, text ='  LP graphs ')
        self.tabControl.add(self.tab2, text = ' Data table ')
        
        self.tabControl.pack(expand = 1, fill ="both")
        
        # Define helpful variables
        self.industry_groups = list(pd.unique(sd.industry_groups['IndustryGroupID']))
        self.start_year = sd.start_year
        self.end_year = sd.end_year
        self.sort_options = sd.sort_options
        self.digit_options = sd.digit_options
        self.graph_limit = sd.graph_limit
        

        # Define summary graphs tab
        self.lp_graphs = LPGraphs(self.tab1, self)
        self.lp_graphs.grid()

        self.tables = Tables(self.tab2, self)
        self.tables.grid()
        


if __name__ == '__main__':
    root = tk.Tk()
    root.title('Analysis app')
    root.geometry("1500x750")
    
    frame_main = MainWindow(root)
    frame_main.pack(fill = 'both', expand = True)
    
    
    root.mainloop()

