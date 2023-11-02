from gspread import WorksheetNotFound
import HelperFunctions as helperModule
import customtkinter as ct
import gspread
import pandas as pd
from ConstantVars.Dictionarys import *
from oauth2client.service_account import ServiceAccountCredentials
from Classes.LoadsClass import Load
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
    labelDate = ct.CTkLabel(**helperModule.configuration_for_label_date(frame, loadObject))
    labelRoute = ct.CTkLabel(**helperModule.configuration_for_label_route(frame, loadObject))
    comboBox = ct.CTkComboBox(**helperModule.configuration_for_combo(frame, LOAD_STATUSES))
    button = ct.CTkButton(**helperModule.configuration_for_button(frame, loadObject),
                          command=lambda: button_function(frame, loadObject, comboBox, connected_sheet))
    helperModule.pack_widgets(routeId, labelRoute, labelDate, comboBox, button)
    frame.pack()


def create_initial_window(app):
    frame = ct.CTkFrame(master=app, fg_color="transparent")
    routeId = ct.CTkLabel(**helperModule.configuration_for_label_id(frame))
    labelRoute = ct.CTkLabel(**helperModule.configuration_for_label_route(frame))
    labelDate = ct.CTkLabel(**helperModule.configuration_for_label_date(frame))
    comboBox = ct.CTkLabel(**helperModule.configuration_for_combo(frame))
    button = ct.CTkLabel(**helperModule.configuration_for_button(frame))
    helperModule.pack_widgets(routeId, labelRoute, labelDate, comboBox, button)
    frame.pack()


def button_function(f: ct.CTkFrame, data: Load, combo: ct.CTkComboBox, sheet, dataFrame=myData):
    comboBoxValue = combo.get()
    dispatcher = data.returnDispatcher()
    loadId = data.returnId()
    if comboBoxValue == DEFAULT:
        helperModule.combo_box_value_default_error()
        return
    dispatcherName = process_dispatcher(dispatcher)
    if dispatcherName:
        try:
            helperModule.register_process(dataFrame, loadId)
        except FileNotFoundError:
            helperModule.file_not_found_error()
            return
        if comboBoxValue == READY:
            update_cell(sheet, dispatcherName, data.acronymFromTo(), GREEN_COLOR, comboBoxValue)
        elif comboBoxValue == NOT_READY:
            update_cell(sheet, dispatcherName, data.acronymFromTo(), RED_COLOR, comboBoxValue)
        else:
            update_cell(sheet, dispatcherName, data.acronymFromTo(), BLACK_COLOR, comboBoxValue)
        f.destroy()
        helperModule.successful_insertion_notification()
    else:
        helperModule.dispatcher_not_found_error()


def update_cell(sheet, dispatcherName, value, color, note):
    color = helperModule.set_color(color)
    cell = sheet.find(dispatcherName)
    if cell is None:
        helperModule.cell_not_found_error()
    else:
        helperModule.insert_load_into_cell(sheet, cell, value, color, note)


def process_dispatcher(dispatcherName):
    return dispatcherNamesDictionary[dispatcherName] if dispatcherName in dispatcherNamesDictionary else False
