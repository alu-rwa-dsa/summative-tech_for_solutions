# import openpyxl and tkinter modules
# importing Mapping to avoid any errors
from collections.abc import Mapping
from openpyxl import *
from tkinter import *
import queue
import random

# create the document variable
applications_document = load_workbook('C:NSO_applications.xlsx')
members_document = load_workbook('C:NSO_members.xlsx')
#an excel function to handle the excel document
applications = []
members = []
