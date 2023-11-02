import pandas
from Classes.LoadsClass import Load
from ConstantVars.ConstantAndVars import *
import tkinter.messagebox as mb
from datetime import datetime
import customtkinter as ct


def process_data(myData: pandas.DataFrame):
    """
    :param myData: Data in a form of Pandas DataFrame
    :return: List of Load objects
    """
    if PROCESS_COLUMN not in myData.columns:
        myData[PROCESS_COLUMN] = 0
    myData[PROCESS_COLUMN] = myData[PROCESS_COLUMN].apply(lambda x: 0 if x != 1 else x)
    if myData.to_excel(PATH_TO_DATA, columns=None, index=False):
        raise Exception(READING_DATA_ERROR)
    myData = myData[myData[PROCESS_COLUMN] != 1]
    loadObject = [Load(data) for data in myData.itertuples()]
    return loadObject


def successful_connection_to_sheet_notification():
    mb.showinfo('Connection', 'Successfully connected')


def successful_insertion_notification():
    mb.showinfo('Notification', 'Load is inserted')


def file_not_found_error():
    mb.showerror('Error', 'No file')


def combo_box_value_default_error():
    mb.showwarning("Warning", "ComboBox value is 'Default'")


def dispatcher_not_found_error():
    mb.showerror('Error', 'Dispatcher not found.')


def updated_data_frame(data, dataFrame):
    dataFrame.loc[dataFrame['Invoice'] == data.returnId(), 'Processed'] = 1


def today_date_for_sheet():
    return datetime.now().strftime("%m.%d.%Y.")


def worksheet_not_found_error():
    mb.showerror('Error', 'Worksheet not found. Create the worksheet.')


def create_main_window():
    ct.set_appearance_mode("System")
    ct.set_default_color_theme("blue")
    app = ct.CTk()
    app.geometry(APP_WINDOW_SIZE)
    ct.CTkLabel(master=app, text="", height=10).pack()
    return app


def create_scrollable_frame(app):
    return ct.CTkScrollableFrame(master=app, fg_color="transparent", width=890, height=440)


def pack_widgets(*args):
    for arg in args:
        arg.pack(**PACK_ARGS)


def configuration_for_label_date(*args):
    if len(args) == 2:
        return {"master": args[0], "text": args[1].returnDate(), "fg_color": "transparent",
                "anchor": "w", "width": 120}
    else:
        return {"master": args[0], "text": "Date", "fg_color": "transparent", "anchor": "w", "width": 120,
                "font": ("Segue UI", 14, "bold")}


def configuration_for_label_route(*args):
    if len(args) == 2:
        return {"master": args[0], "text": args[1].returnRoute(), "fg_color": "transparent", "anchor": "w",
                "width": 250}
    else:
        return {"master": args[0], "text": "Route", "fg_color": "transparent", "anchor": "w", "width": 250,
                "font": ("Segue UI", 14, "bold")}


def configuration_for_combo(*args):
    if len(args) == 2:
        return {"master": args[0], "values": args[1], "width": 150}
    else:
        return {"master": args[0], "text": "Options", "fg_color": "transparent", "anchor": "w", "width": 150,
                "font": ("Segue UI", 14, "bold")}


def configuration_for_button(*args):
    if len(args) == 2:
        return {"master": args[0], "text": "Insert", "width": 150}
    else:
        return {"master": args[0], "text": "Button", "fg_color": "transparent", "anchor": "w", "width": 150,
                "font": ("Segue UI", 14, "bold")}


def configuration_for_label_id(frame):
    return {"master": frame, "text": "Invoice", "fg_color": "transparent", "anchor": "w", "width": 70,
            "font": ("Segue UI", 14, "bold")}


def register_process(dataFrame, loadId):
    dataFrame.loc[dataFrame['Invoice'] == loadId, 'Processed'] = 1
    dataFrame.to_excel(PATH_TO_DATA, columns=None, index=False)


def set_color(color):
    if color == "black":
        color = BLACK_COLOR_VALUE
    elif color == "red":
        color = RED_COLOR_VALUE
    elif color == "green":
        color = GREEN_COLOR_VALUE
    else:
        mb.showerror('Error', 'Bad color.')
        exit()
    return color


def cell_not_found_error():
    mb.showerror("Error", "Cell is not found and the program can not run successfully.")
    exit()


def insert_load_into_cell(sheet, cell, value, color, note):
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
