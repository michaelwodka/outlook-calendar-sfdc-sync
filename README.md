# python-outlook-calendar-sfdc-sync
This project allows you to download you calendar/appointments from Microsoft Outlook into an Excel spreadsheet and then
upload the appointment data to contacts and opportunties in Salesforce. The script is written in Python. The calendar download/upload
processes are executed via a Python Ttkinter GUI application and simple-salesforce script. This script can only be run on a Windows machine.

# Getting Started
You will need to install several python libraries to get this project running on your local machine.
```
from tkinter import *                                   # for Tkinter GUI application
import tkinter as tk                                    # for Tkinter GUI application
import os                                               # for opening and closing files
import datetime as dt                                   # for using calendar dates as values
from dateutil.relativedelta import relativedelta        # for adding/subtracting dates from each other
import pytz                                             # for setting timezones for dates
from openpyxl import load_workbook                      # for downloading Pandas table into Excel workbook
import win32com.client                                  # for running Excel and Outlook applications from Python
import win32api                                         # for creating Windows pop-up messages
import xlrd                                             # for retrieving data from Excel cells
from simple_salesforce import Salesforce                # for running API calls to Salesforce
import pandas as pd                                     # for creating Pandas tables in Python
from openpyxl.styles import Font, Color, PatternFill    # for making design edits (color, font) in Excel workbook
import win32timezone                                    # for setting timezones for dates
```
# License
See the LICENSE file for license rights and limitations (MIT).
