import tkinter.messagebox as mb
from gspread import WorksheetNotFound
import HelperFunctions as helperModule
import customtkinter as ct
import gspread
import pandas as pd
from ConstantVars.Dictionarys import *
from oauth2client.service_account import ServiceAccountCredentials
from Classes.LoadsClass import Load
from ConstantVars.Dictionarys import directory
from ConstantVars.ConstantAndVars import *

myData = pd.read_excel(PATH_TO_DATA).fillna("")


def connect_to_sheet():
    sheetID = helperModule.today_date_for_sheet()
    creds = ServiceAccountCredentials.from_json_keyfile_name(PATH_TO_SERVICE_FILE, SCOPE)
    client = gspread.authorize(creds)
    try:
        connected_sheet = client.open(SHEET_NAME).worksheet(sheetID)
        return connected_sheet
    except WorksheetNotFound:
        helperModule.worksheet_not_found_error()
        exit()


def create_objects_from_excel():
    global myData
    try:
        loadObject = helperModule.process_data(myData)
    except WorksheetNotFound:
        helperModule.worksheet_not_found_error()
        exit()
    return loadObject


def create_the_main_screen(loadObjects, connected_sheet):
    app = helperModule.create_main_window()
    helperModule.successful_connection_to_sheet_notification()
    scrollableFrame = helperModule.create_scrollable_frame(app)
    create_initial_window(scrollableFrame)
    for loadObject in loadObjects:
        load_display(loadObject, scrollableFrame, connected_sheet)
    scrollableFrame.pack()
    app.mainloop()


def load_display(loadObject, app, connected_sheet):
    frame = ct.CTkFrame(master=app, fg_color="transparent")
    create_a_frame(frame, loadObject, connected_sheet)
    frame.pack()


def create_a_frame(frame, *args):
    create_a_load_frame(frame, args[0], args[1]) if args else create_initial_window(frame)


def create_a_load_frame(frame, loadObject: Load, connected_sheet):
    routeId = ct.CTkTextbox(frame)
    routeId.insert("0.0", loadObject.returnId())
    routeId.configure(**ROUT_ID_CONFIG)
    routeToDisplay = loadObject.returnRoute()
    labelDispatcher = ct.CTkLabel(master=frame, text=loadObject.returnDispatcher(), fg_color="transparent",
                                  anchor="w", width=120)
    labelRoute = ct.CTkLabel(master=frame, text=routeToDisplay, fg_color="transparent", anchor="w", width=250)
    comboBox = ct.CTkComboBox(master=frame, values=LOAD_STATUSES, width=150)
    button = ct.CTkButton(master=frame, text="Insert",
                          command=lambda: button_function(frame, loadObject, comboBox, connected_sheet), width=150)
    routeId.pack(**PACK_ARGS)
    labelRoute.pack(**PACK_ARGS)
    labelDispatcher.pack(**PACK_ARGS)
    comboBox.pack(**PACK_ARGS)
    button.pack(**PACK_ARGS)
    frame.pack()


def create_initial_window(app):
    frame = ct.CTkFrame(master=app, fg_color="transparent")
    # routeId
    ct.CTkLabel(master=frame, text="Invoice", fg_color="transparent", anchor="w", width=70,
                font=("Segue UI", 14, "bold")).pack(**PACK_ARGS)
    # labelRoute
    ct.CTkLabel(master=frame, text="Route", fg_color="transparent", anchor="w", width=250,
                font=("Segue UI", 14, "bold")).pack(**PACK_ARGS)
    # labelDispatcher
    ct.CTkLabel(master=frame, text="Dispatcher", fg_color="transparent", anchor="w", width=120,
                font=("Segue UI", 14, "bold")).pack(
        **PACK_ARGS)
    # comboBox
    ct.CTkLabel(master=frame, text="Options", fg_color="transparent", anchor="w", width=150,
                font=("Segue UI", 14, "bold")).pack(**PACK_ARGS)
    # button
    ct.CTkLabel(master=frame, text="Button", fg_color="transparent", anchor="w", width=150,
                font=("Segue UI", 14, "bold")).pack(**PACK_ARGS)
    frame.pack()


# !!! Multithread !!!
def button_function(f: ct.CTkFrame, data: Load, combo: ct.CTkComboBox, sheet, dataFrame=myData):
    comboBoxValue = combo.get()
    if comboBoxValue == "Default":
        helperModule.combo_box_value_default_error()
    else:
        dispatcherName = process_dispatcher(data.returnDispatcher())
        if dispatcherName:
            try:
                dataFrame.loc[dataFrame['Invoice'] == data.returnId(), 'Processed'] = 1
                dataFrame.to_excel(PATH_TO_DATA, columns=None, index=False)
            except FileNotFoundError:
                helperModule.file_not_found_error()
                return
            if comboBoxValue == READY:
                update_cell_value_by_dispatcher(sheet, dispatcherName, data.loadFromTo(), GREEN_COLOR,
                                                comboBoxValue)
            elif comboBoxValue == NOT_READY:
                update_cell_value_by_dispatcher(sheet, dispatcherName, data.loadFromTo(), RED_COLOR, comboBoxValue)
            else:
                update_cell_value_by_dispatcher(sheet, dispatcherName, data.loadFromTo(), BLACK_COLOR,
                                                comboBoxValue)
            f.destroy()
            helperModule.successful_insertion_notification()
        else:
            helperModule.dispatcher_not_found_error()


def update_cell_value_by_dispatcher(sheet, dispatcherName, value, color, note):
    if color == "black":
        color = BLACK_COLOR_VALUE
    elif color == "red":
        color = RED_COLOR_VALUE
    elif color == "green":
        color = GREEN_COLOR_VALUE
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
