import pandas
from Classes.LoadsClass import Load
from ConstantVars.ConstantAndVars import *


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
