# -*- coding: utf-8 -*-
"""
Created on Wed May 31 08:42:33 2023

@author: McGuigan_M
Modified by: Long_S
"""

import pandas as pd
import os, sys
from pathlib import Path
import datetime
import sqlalchemy as sa

update_year = datetime.date.today().year - 1

# Add parent directory access
sys.path.append(os.path.dirname(os.path.dirname(os.path.realpath(__file__))))

# source_dir = Path.cwd().parents[0] / 'source_data'
# state_abbrev_dir = Path.cwd().parents[0]

red = '#D2232A'
blue = '#0F05A5'
light_blue = '#AFD2FF'
gold = '#DDB139'
light_gray = '#BEBEBE'
dark_gray = '#7D7D7D'
salmon = '#E19696'
medium_blue = '#5A96FF'

# NOTE FROM NATE DURING 2022 update: Best to refer to published levels of detail: three decimal points for indexes, 
# one for rates. It's all in the XLSX files except for the national-level series.

# I didn't use the file above that Nate is talking about for this program. Need to edit it next time

try:
    engine = sa.create_engine(
    "mssql+pyodbc://OPTSRV2/IPS2DB_SAS?driver=ODBC+Driver+17+for+SQL+Server")
    driver_test = pd.read_sql("""SELECT TOP 1 * FROM dbo.ReportBuilder_Preliminary""", engine)
except:
    engine = sa.create_engine(
        "mssql+pyodbc://OPTSRV2/IPS2DB_SAS?driver=ODBC Driver 18 for SQL Server&ssl_verify_cert=false&TrustServerCertificate=yes")
    
## Get industry groups
industry_groups = pd.read_sql("""SELECT IndustryID, IndustryGroupID FROM dbo.IndustryGroupSets""", engine)
industry_groups = industry_groups[~industry_groups['IndustryGroupID'].isin(['Labor', 'BEANIPA', 'Sector63', 'Hospitals'])]

start_year = 1987
end_year = update_year

sort_options = ['extreme oph change first', 'highest relative oph change first', 'highest interest score first']

digit_options = ['3-Digit', '4-Digit']

graph_limit = 5