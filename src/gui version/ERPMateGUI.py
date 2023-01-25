#importing packages
import PySimpleGUI as sg
import pandas as pd

#adding a colour theme
sg.theme("DarkTeal9")

#declaring and reading the excel file
EXCEL_FILE="ERPSessionData.xlsx"
df=pd.read_excel(EXCEL_FILE)

#layout
layout= [
    [sg.Text("Please enter the data of ERPsession in the following fields: ")],#title of the app
    [sg.Text("Date",size=(15,1)),sg.InputText(key='date')],#date
    [sg.Text("Time",size=(15,1)),sg.InputText(key="time")],#time
    [sg.Text("Obsession",size=(15,3)),sg.InputText(key="obsession")],#obsession
    [sg.Text("Compulsion",size=(15,3)),sg.InputText(key="compulsion")],#compulsion
    [sg.Text("Rate your anxiety levels out of 10: ")],#info
    [sg.Text("Anxiety at Exposure",size=(15,1)),sg.InputText(key="anxietyatexposure")],#anxiety at exposure
    [sg.Text("Anxiety after 5 min",size=(15,1)),sg.InputText(key="anxietyat5min")],#anxiety after 5 min
    [sg.Text("Anxiety after 10 min",size=(15,1)),sg.InputText(key="anxietyat10min")],#anxiety after 10 min
    [sg.Text("Anxiety after 30 min",size=(15,1)),sg.InputText(key="anxietyat30min")],#anxiety after 30 min
    [sg.Text("Anxiety after 1 hour",size=(15,1)),sg.InputText(key="anxietyat1hr")], #anxiety after 1 hour
    [sg.Submit(),sg.Button("clear"),sg.Exit()]
]

#passing the layout to window
window=sg.Window("ERPMateGUI v2.0",layout)

#for the clear button
def clear_input():
    for key in values:
        window[key]('')
    return None


#handling a close app event
while True:
    event,values=window.read()
    if event==sg.WIN_CLOSED or event=="Exit":
        break
    
    #handling clear button event
    if event=="Clear":
        clear_input()
    if event=="Submit":
        #appending the data to excel sheet using pandas
        df=df.append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE,index=False)
        sg.popup("Session Data Saved")
        clear_input()

#closing window
window.close()
