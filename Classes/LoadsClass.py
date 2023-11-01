from ConstantVars.Dictionarys import *


class Load:
    def __init__(self, data):
        self.id = data[1]
        self.route = data[2]
        self.date = data[3]
        self.dispatcher = data[4]

    def returnDispatcher(self):
        return self.dispatcher

    def returnRoute(self) -> list:
        lista = []
        for slovo in range(len(self.route) - 1):
            if self.route[slovo:slovo + 2] in directory:
                lista.append(self.route[slovo:slovo + 2])
        return lista

    def returnDate(self):
        return self.date

    def returnId(self):
        return self.id

    def returnEverything(self):
        print(self.id)
        print(self.date)
        print(self.route)
        print(self.dispatcher)

    def loadFromTo(self) -> str:
        myList = self.returnRoute()
        return str(myList[0]) + "-" + str(myList[len(myList) - 1])
