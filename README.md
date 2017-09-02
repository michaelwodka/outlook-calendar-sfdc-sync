# python-outlook-calendar-sfdc-sync
This project allows you to download you calendar/appointments from Microsoft Outlook into an Excel spreadsheet and then
upload the appointment data to contacts and opportunties in Salesforce. The script is written in Python. The calendar download/upload
processes are executed via a Python Ttkinter GUI application and simple-salesforce script.

# Getting Started
You will need to install several python libraries to get this project running on your local machine.
```
from tkinter import *
import tkinter as tk
import os 
import datetime as dt
from dateutil.relativedelta import relativedelta
import pytz
from openpyxl import load_workbook
import win32com.client
import win32api
import xlrd
from simple_salesforce import Salesforce
import pandas as pd #imports Pandas to create a table in Python
from openpyxl.styles import Font, Color, PatternFill
import win32timezone
```
# License
See the LICENSE file for license rights and limitations (MIT).
