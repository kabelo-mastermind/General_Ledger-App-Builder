import PySimpleGUI as sg
import pandas as pd

# Add some color to the window
sg.theme('HotDogStand')

EXCEL_FILE = 'G-Ledger.xlsx'
df = pd.read_excel(EXCEL_FILE)

layout = [
    [sg.Text('Please fill out the following fields:')],
    [sg.Text('Date(dd/mm/yy)', size=(15,1)), sg.InputText(key='DATE')],
    [sg.Text('Description', size=(15,1)), sg.InputText(key='DESCRIPTION')],
    [sg.Text('Post Ref:', size=(15,1)), sg.InputText(key='POST REF')],
    [sg.Text('Debit', size=(15,1)), sg.InputText(key='DEBIT')],
    [sg.Text('Credit', size=(15,1)), sg.InputText(key='CREDIT')],
    [sg.Text('Balance', size=(15,1)), sg.InputText(key='BALANCE')],
    
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

window = sg.Window('Ledger entry form', layout)

def clear_input():
    for key in values:
        window[key]('')
    return None


while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
        clear_input()
    if event == 'Submit':
        new_record = pd.DataFrame(values, index=[0])
        df = pd.concat([df, new_record], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Data saved!')
        clear_input()
window.close()
