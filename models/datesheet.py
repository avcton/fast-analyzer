import numpy as np
import pandas as pd


class DateSheet():
    def __init__(self, title: str = None, semester: str = None,) -> None:
        self.title = title
        self.semester = semester
        self.dateSheet_columns = ['Date', 'Day', 'Time', 'Code', 'Course']
        self.dateSheet = pd.DataFrame(columns=self.dateSheet_columns)
        self.loaded = False
        self.courses = []

    def load(self, file: any) -> bool:
        try:
            self.file = file

            if self.title == None:
                # If custom title string was not providede
                self.title = pd.read_excel(
                    self.file, header=None, nrows=1).iloc[0, 0]

            if self.semester == None:
                # If custom semester string was not provided
                self.semester = pd.read_excel(self.file, header=None,
                                              skiprows=1, nrows=1).iloc[0, 0]

            dateSheet = pd.read_excel(self.file, skiprows=2)

            dateSheet_refined = dateSheet.melt(id_vars=['Day', 'Date', 'Code'], value_vars=[
                dateSheet.columns[3]], var_name='Time', value_name='Course')
            dateSheet_refined.dropna(inplace=True)

            i = 0
            for index, _ in enumerate(dateSheet.columns[4:]):
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
                self.courses.append({
                    'Index': indx,
                    'Code': row.Code,
                    'Title': row.Course,
                    'View': row.Code + ' - ' + row.Course
                })

            self.loaded = True

        except Exception as e:
            print(
                f"Caught an error: {e}")
            self.loaded = False

        finally:
            return self.loaded

    def readByCode(self, codes: list[str]):
        assert(self.loaded == True)
        mask = np.zeros(self.dateSheet.shape[0], dtype=bool)

        for code in codes:
            mask = mask | (self.dateSheet['Code'] == code)

        return self.dateSheet[mask]

    def readByCourse(self, courses: list[dict]):
        assert (self.loaded == True)
        indexes = [course['Index'] for course in courses]
        indexes.sort()

        return self.dateSheet.iloc[indexes]
