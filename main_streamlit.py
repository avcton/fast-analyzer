import numpy as np
import pandas as pd
from taipy.gui import Gui


class Course():
    def __init__(self, code: str, title: str) -> None:
        self.Code = code
        self.Title = title
        self.View = code + ' - ' + title


class DateSheet():
    def __init__(self, path: str = None, title: str = None, semester: str = None,) -> None:
        self.title = title
        self.semester = semester
        self.excel_file_path = path
        self.loaded = False
        self.dateSheet_columns = ['Date', 'Day', 'Time', 'Code', 'Course']
        self.dateSheet = pd.DataFrame(columns=self.dateSheet_columns)
        self.courses = []

        if self.excel_file_path != None:
            self.load(self.excel_file_path)
            self.loaded = True

    def load(self, path: str) -> None:

        if self.excel_file_path == None:
            self.excel_file_path = path

        if self.title == None:
            self.title = pd.read_excel(
                self.excel_file_path, header=None, nrows=1).iloc[0, 0]

        if self.semester == None:
            self.semester = pd.read_excel(self.excel_file_path, header=None,
                                          skiprows=1, nrows=1).iloc[0, 0]

        dateSheet = pd.read_excel(self.excel_file_path, skiprows=2)

        dateSheet_refined = dateSheet.melt(id_vars=['Day', 'Date', 'Code'], value_vars=[
                                           dateSheet.columns[3]], var_name='Time', value_name='Course')
        dateSheet_refined.dropna(inplace=True)

        i = 0
        for index, col in enumerate(dateSheet.columns[4:]):
            if index % 2 == 0:
                continue
            i += 1
            fragment = dateSheet.melt(id_vars=['Day', 'Date', f'Code.{i}'], value_vars=[
                                      dateSheet.columns[4 + index]], var_name='Time', value_name='Course')
            fragment.columns = ['Day', 'Date', 'Code', 'Time', 'Course']
            fragment.dropna(inplace=True)

            dateSheet_refined = pd.concat(
                [dateSheet_refined, fragment],  ignore_index=False)

        dateSheet_refined['Date'] = pd.to_datetime(
            dateSheet_refined['Date'], format='%d-%b-%Y')

        dateSheet_refined = dateSheet_refined[self.dateSheet_columns]
        dateSheet_refined.sort_values(
            by=['Date', 'Day', 'Time', 'Code'], ascending=True, inplace=True, ignore_index=True)

        dateSheet_refined['Date'] = dateSheet_refined['Date'].dt.strftime(
            "%d %b %Y")

        self.dateSheet = dateSheet_refined

        for indx, row in self.dateSheet[['Code', 'Course']].iterrows():
            self.courses.append(Course(row.Code, row.Course))

    def readByCode(self, codes: list):
        mask = np.zeros(self.dateSheet.shape[0], dtype=bool)

        for code in codes:
            mask = mask | (self.dateSheet['Code'] == code)

        return self.dateSheet[mask]
