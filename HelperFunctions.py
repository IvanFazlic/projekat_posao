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
