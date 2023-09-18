# Bureau of Labor Statistics Project
Updated by Summer Long, 08/21/2023

[[_TOC_]]

## Project Overview
This program is intended to aid analysts in the identification of topics for internal releases. It does so in two ways:
- Generates labor productivity charts with user-inputted parameters and sorts them based on a user-selected criteria
- Generates tables with user-inputted parameters
The data is accessed through a Microsoft SQL Server. 
 
## How it Works
The user runs the main_window.py file on an IDE. (In the office, it was packaged in an .exe)
An interface comes up (created with Tkinter) with two tabs: one to produce labor productivity charts with a sorting criteria, and one to produce tables. 
The database is selected based on user-inputted parameters to avoid extraneous data from being stored in the user's memory using SQLAlchemy and Pandas.
The appropriate measures are calculated and charts are generated on an as-needed basis. Then, they are exported to a Word document.

## Notice of Redacted Code
As the database was accessed through a governmental server, the method of accessing the server has been redacted. In its place, the address has been set as "placeholder."
Due to this, sample outputs and screenshots of the interface have been shared in the outputs and interface folder respectively.  