import tkinter.messagebox as mb

from gspread import WorksheetNotFound

import HelperFunctions as helperModule
import customtkinter as ct
import gspread
import pandas as pd
from ConstantVars.Dictionarys import *
from oauth2client.service_account import ServiceAccountCredentials
from Classes.LoadsClass import Load
from ConstantVars.Dictionarys import directory, loadStatuses
from ConstantVars.ConstantAndVars import *
from datetime import datetime

myData = pd.read_excel(PATH_TO_DATA).fillna("")


def connect_to_sheet():
    currentDay = datetime.now().day
    currentMonth = datetime.now().month
    currentYear = datetime.now().year

    currentDay = '0' + str(currentDay) if currentDay < 10 else str(currentDay)
    currentMonth = '0' + str(currentMonth) if currentMonth < 10 else str(currentMonth)
    currentYear = str(currentYear)
    currentTime = currentMonth + "." + currentDay + "." + currentYear + "."

    creds = ServiceAccountCredentials.from_json_keyfile_name(PATH_TO_SERVICE_FILE, SCOPE)
    client = gspread.authorize(creds)
    try:
        connected_sheet = client.open("Weekend Track and Trace").worksheet(currentTime)
        return connected_sheet
    except WorksheetNotFound:
        mb.showerror('Error', 'Worksheet not found. Create the worksheet.')
        exit()


def CreateObjectsFromExcel():
    global myData
    loadObject = helperModule.process_data(myData)
    return loadObject


def CreateTheMainScreen(loadObjects, connected_sheet):
    # create ctk window
    ct.set_appearance_mode("System")
    ct.set_default_color_theme("blue")
    app = ct.CTk()
    app.geometry(APP_WINDOW_SIZE)
    ct.CTkLabel(master=app, text="", height=10).pack()
    mb.showinfo('Connection', 'Successfully connected')
    frame = ct.CTkFrame(master=app, fg_color="transparent")
    create_initial_window(frame)
    for loadObject in loadObjects:
        LoadDisplay(loadObject, app, connected_sheet)
    app.mainloop()


def LoadDisplay(loadObject, app, connected_sheet):
    frame = ct.CTkFrame(master=app, fg_color="transparent")
    create_a_frame(frame, loadObject, connected_sheet)
    frame.pack()


def create_a_frame(frame, *args):
    create_a_load_frame(frame, args[0], args[1]) if args else create_initial_window(frame)


def create_a_load_frame(frame, loadObject, connected_sheet):
    routeId = ct.CTkTextbox(frame)
    routeId.insert("0.0", loadObject.returnId())
    routeId.configure(state="disabled", height=5, width=70, fg_color="transparent")
    route = loadObject.returnRoute()
    routeToDisplay = directory[route[0]] + " - " + directory[route[len(route) - 1]]
    labelDispatcher = ct.CTkLabel(master=frame, text=loadObject.returnDispatcher(), fg_color="transparent",
                                  anchor="w", width=120)
    labelRoute = ct.CTkLabel(master=frame, text=routeToDisplay, fg_color="transparent", anchor="w", width=250)
    comboBox = ct.CTkComboBox(master=frame, values=loadStatuses, width=150)
    button = ct.CTkButton(master=frame, text="Insert",
                          command=lambda: button_function(frame, loadObject, comboBox, connected_sheet), width=150)
    routeId.pack(**PACKING_ARGUMENTS)
    labelRoute.pack(**PACKING_ARGUMENTS)
    labelDispatcher.pack(**PACKING_ARGUMENTS)
    comboBox.pack(**PACKING_ARGUMENTS)
    button.pack(**PACKING_ARGUMENTS)
    frame.pack()


def create_initial_window(frame):
    # routeId
    ct.CTkLabel(master=frame, text="Invoice", fg_color="transparent", anchor="w", width=70,
                font=("Segue UI", 14, "bold")).pack(**PACKING_ARGUMENTS)
    # labelRoute
    ct.CTkLabel(master=frame, text="Route", fg_color="transparent", anchor="w", width=250,
                font=("Segue UI", 14, "bold")).pack(**PACKING_ARGUMENTS)
    # labelDispatcher
    ct.CTkLabel(master=frame, text="Dispatcher", fg_color="transparent", anchor="w", width=120,
                font=("Segue UI", 14, "bold")).pack(
        **PACKING_ARGUMENTS)
    # comboBox
    ct.CTkLabel(master=frame, text="Options", fg_color="transparent", anchor="w", width=150,
                font=("Segue UI", 14, "bold")).pack(**PACKING_ARGUMENTS)
    # button
    ct.CTkLabel(master=frame, text="Button", fg_color="transparent", anchor="w", width=150,
                font=("Segue UI", 14, "bold")).pack(**PACKING_ARGUMENTS)
    frame.pack()


# !!! Multithread !!!
def button_function(f: ct.CTkFrame, data: Load, combo: ct.CTkComboBox, sheet, dataFrame=myData):
    comboBoxValue = combo.get()
    if comboBoxValue == "Default":
        mb.showwarning("Warning", "ComboBox value is 'Default'")
    else:
        dataFrame.loc[dataFrame['Invoice'] == data.returnId(), 'Processed'] = 1
        try:
            dataFrame.to_excel('./Data/Data.xlsx', columns=None, index=False)
            dispatcherName = process_dispatcher(data.returnDispatcher())
            if comboBoxValue == READY and dispatcherName:
                UpdateCellValueByDispatcher(sheet, dispatcherName, data.loadFromTo(), "green", 'Something')
                f.destroy()
                mb.showinfo('Notification', 'Load is inserted')
        except FileNotFoundError:
            mb.showerror('Error', 'No file')


def UpdateCellValueByDispatcher(sheet, dispatcherName, value, color, note):
    if color == "black":
        color = BLACK_COLOR
    elif color == "red":
        color = RED_COLOR
    elif color == "green":
        color = GREEN_COLOR
    else:
        print("Bad color")
        exit()
    cell = sheet.find(dispatcherName)
    if cell is None:
        mb.showerror("Error", "Cell is not found and the program can not run successfully.")
        exit()
    else:
        cellRow = cell.row
        cellCol = cell.col
        while sheet.cell(cellRow, cellCol).value is not None:
            cellCol += 1
        sheet.update_cell(cellRow, cellCol, value)
        cellRow += 1
        cell = ["A"]
        for x in range(cellCol // 26):
            cell.append("A")
        cell[len(cell) - 1] = chr((ord(cell[len(cell) - 1]) + cellCol % 26) - 1)
        cellVal = ""
        for x in cell:
            cellVal += x
        cellVal = cellVal + str(cellRow)
        sheet.format(cellVal, {"backgroundColor": color})
        sheet.insert_note(cellVal, note)


def process_dispatcher(dispatcherName):
    return dispatcherNamesDictionary[dispatcherName] if dispatcherName in dispatcherNamesDictionary else False
