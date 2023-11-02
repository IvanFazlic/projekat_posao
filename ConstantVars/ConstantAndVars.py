PACK_ARGS = {'side': 'left', 'padx': 15, 'pady': 15}
PROCESS_COLUMN = 'Processed'
PATH_TO_DATA = './Data/Data.xlsx'
PATH_TO_SERVICE_FILE = './Functions/ClientSecret.json'
READING_DATA_ERROR = "Do not succeed"
APP_WINDOW_SIZE = "890x440"
SCOPE = ['https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive.file",
         "https://www.googleapis.com/auth/drive"]
BLACK_COLOR_VALUE = {
    "red": 0,
    "green": 0,
    "blue": 0
}
RED_COLOR_VALUE = {
    "red": 224 / 255,
    "green": 102 / 255,
    "blue": 102 / 255
}
GREEN_COLOR_VALUE = {
    "red": 52 / 255,
    "green": 168 / 255,
    "blue": 83 / 255
}
BLACK_COLOR = 'black'
RED_COLOR = 'red'
GREEN_COLOR = 'green'
READY = "Spremno"
NOT_READY = "Load nije spreman"
LOAD_STATUSES = ["Default", "Spremno", "Load nije spreman", "Nisam mogao da dobijem shipping",
                 "Nisam nasao broj", "VoiceM",
                 "Nema adresa u dokumentu",
                 "Pogresan broj na internetu",
                 "Broj stalno nedostupan",
                 "Nije okacen dokument", "Nema odgovarajucih informacija u dokumentu", "Prerano za informacije"]
SHEET_NAME = "Weekend Track and Trace"
ROUT_ID_CONFIG = {"state": "disabled", "height": 5, "width": 70, "fg_color": "transparent"}
DEFAULT = "Default"
