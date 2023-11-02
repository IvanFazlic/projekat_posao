from pandas import DataFrame as dataFrame

from ConstantVars.Dictionarys import *


class Load:
    def __init__(self, data: dataFrame):
        self.id = data[1]
        self.route = data[2]
        self.date = data[3]
        self.dispatcher = data[4]

    def returnDispatcher(self):
        return self.dispatcher

    def returnRoute(self) -> str:
        lista = []
        for slovo in range(len(self.route) - 1):
            if self.route[slovo:slovo + 2] in directory:
                lista.append(self.route[slovo:slovo + 2])
        returnString = directory[lista[0]] + " - " + directory[lista[len(lista) - 1]]
        return returnString

    def returnDate(self):
        return self.date.strftime("%d.%m.%Y.")

    def returnId(self):
        return self.id

    def acronymFromTo(self):
        lista = []
        for slovo in range(len(self.route) - 1):
            if self.route[slovo:slovo + 2] in directory:
                lista.append(self.route[slovo:slovo + 2])
        return lista[0] + "-" + lista[len(lista) - 1]
