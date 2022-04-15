# import openpyxl and tkinter modules
# importing Mapping to avoid any errors
from collections.abc import Mapping
from openpyxl import *
from tkinter import *
from tkinter import messagebox
import queue
import random
admin_passcode = 9043
# create the document variable
applications_document = load_workbook('C:NSO_assets\\NSO_applications.xlsx')
members_document = load_workbook('C:NSO_assets\\NSO_members.xlsx')
#an excel function to handle the excel document
applications = [{'form number': 453, 'first name': 'Alec', 'Middle name': '', 'Last name': 'Dumbuya', 'Gender': 'male',
                 'Email ID': 'gd.gmail.com', 'Phone number': '0785963843', 'Address': 'Kicukiro, Kigali', 'Occupation': 'student'},
                {'form number': 76, 'first name': 'Gideon', 'Middle name': 'Agnes', 'Last name': 'Luthor', 'Gender': 'female',
                 'Email ID': 'Gideondf3@gmail.com', 'Phone number': '0758695484', 'Address': 'Gicumbi, Rwanda', 'Occupation': 'farmer'}]
members = [{'form number': 22, 'first name': 'Queen', 'Middle name': 'Keza', 'Last name': 'Umunyana', 'Gender': 'female',
           'Email ID': 'queenumunyana29@gmail.com', 'Phone number': '0784828258', 'Address': 'Kimironko, Kigali', 'Occupation': 'student'},
           {'form number': 30, 'first name': 'Becks', 'Middle name': 'Furaha', 'Last name': 'Nishinda', 'Gender': 'non binary',
            'Email ID': 'becksfinda@gmail.com', 'Phone number': '0784953735', 'Address': 'Remera, Kigali', 'Occupation': 'photographer'}]
rejected_applications = []
sheet = applications_document.active
sheet_members = members_document.active
details = {'form number': '', 'first name': '', 'Middle name': '', 'Last name': '', 'Gender': '',
           'Email ID': '', 'Phone number': '', 'Address': '', 'Occupation': ''}
def fetch_app_data():
    row_view = 'NOT EMPTY'
    row_ch = 2
    cell = 9
    while row_view == 'NOT EMPTY':
        ch_d = []
        col_ch = 1
        value = sheet.cell(row=row_ch, column=col_ch).value
        ch_d.append(value)
        for val in range(len(ch_d)):
            if ch_d[val] == '':
                cell = cell - 1
                if cell == 0:
                    row_view = 'EMPTY'
        if row_view == 'NOT EMPTY':
            cl = 0
            for i in details:
                details[i] = sheet.cell(row=row_ch, column=cl+1).value
            applications.append(details)

def fetch_member_data():
    row_view = 'NOT EMPTY'
    row_ch = 2
    cell = 9
    while row_view == 'NOT EMPTY':
        ch_d = []
        col_ch = 1
        value = sheet_members.cell(row=row_ch, column=col_ch).value
        ch_d.append(value)
        for val in range(len(ch_d)):
            if ch_d[val] == '':
                cell = cell - 1
            if cell == 0:
                row_view = 'EMPTY'
        if row_view == 'NOT EMPTY':
            cl = 0
            for i in details:
                details[i] = sheet_members.cell(row=row_ch, column=cl+1).value
            applications.append(details)