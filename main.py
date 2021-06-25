# -*- coding: utf-8 -*-
import time, json, pyodbc, pandas as pd
from colorama import init, Fore, Style
init()

class Sqller:
    def __init__(self):
        cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.70.70\SPOLKA;DATABASE=SPK;UID=lajkonik;PWD=Silcare01')
        self.cursor = cnxn.cursor()
        self.query = ''
        self.data = None

    def get_data(self):
        self.cursor.execute(self.query)
        self.data = []
        columns = [column[0] for column in self.cursor.description]
        for row in self.cursor.fetchall():
            self.data.append(dict(zip(columns, row)))

    def create_xlsx(self):
        df = pd.DataFrame(data=self.data)
        df.to_excel(str(time.time() * 1000) + '.xlsx')
        print(Fore.CYAN + "Plik został wygenerowany\n" + Style.RESET_ALL)

    def run(self):
        while True:
            print(Fore.CYAN + "\nPodaj zapytanie SQL: " + Style.RESET_ALL)
            self.query = input()
            self.get_data()
            self.create_xlsx()

if __name__ == '__main__':
    Sqller().run()
