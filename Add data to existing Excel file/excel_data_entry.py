import PySimpleGUI as sg
import pandas as pd

# add some color to the window
sg.theme("DarkTeal9")

EXCEL_FILE = "excel_data_entry.xlsx"
df = pd.read_excel(EXCEL_FILE)

layout = [
    [sg.Text("Please fill out the following fields:")],
    [sg.Text("Date", size=(10, 1)), sg.InputText(key="DATE")],
    [sg.Text("Exchange", size=(10, 1)), sg.Combo(["Kraken", "eToro", "Binance"], key="EXCHANGE")],
    [sg.Text("Asset", size=(10, 1)), sg.InputText(key="ASSET")],
    [sg.Text("Units", size=(10, 1)), sg.InputText(key="UNITS")],
    [sg.Text("Price", size=(10, 1)), sg.InputText(key="PRICE")],
    [sg.Text("Type", size=(10, 1)), sg.Combo(["Reward", "Buy", "Sell"], key="TYPE")],
    [sg.Submit(), sg.Button("Clear"), sg.Exit()]
]

window = sg.Window("Simple data entry form", layout)


def clear_input():
    for key in values:
        window[key]("")
    return None


while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == "Exit":
        break

    if event == "Clear":
        clear_input()

    if event == "Submit":
        df = df.append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup("Data saved!")
        clear_input()
window.close()
