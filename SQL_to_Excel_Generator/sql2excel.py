#!/usr/bin/python3

import PySimpleGUI as sg
import mysql.connector
import pandas as pd
import os


#Define the GUI window

sg.theme('SandyBeach')
layout = [
    [sg.Text('Enter SQL')],
    [sg.Multiline(size=(50,5), key='textbox')],
    [sg.Submit(), sg.Cancel()]
]
window = sg.Window('sql2excel', layout)
event, values = window.read()
query = values['textbox']

#SQL Connection Query

userl = 'your_username'
password = 'your_password'
host1 = 'your_host'
database1 = 'your_database'

cnx = mysql.connector.connect(user = userl, password = password, host = host1,
                              database = database1)

#Now create and open report
df = pd.read_sql(query, cnx)
df.to_excel('your_report.xlsx', index=False)
os.startFile('your_report.xlsx')