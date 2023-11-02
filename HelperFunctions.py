import pandas
from Classes.LoadsClass import Load
from ConstantVars.ConstantAndVars import *
import tkinter.messagebox as mb


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
