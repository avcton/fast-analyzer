import numpy as np
import pandas as pd
import streamlit as st


class DateSheet():
    def __init__(self, file: any = None, title: str = None, semester: str = None,) -> None:
        self.title = title
        self.semester = semester
        self.file = file
        self.loaded = False
        self.dateSheet_columns = ['Date', 'Day', 'Time', 'Code', 'Course']
        self.dateSheet = pd.DataFrame(columns=self.dateSheet_columns)
        self.courses = []

        if self.file != None:
            self.load(self.file)
            self.loaded = True

    def load(self, file: any) -> None:

        if self.file == None:
            self.file = file

        if self.title == None:
            self.title = pd.read_excel(
                self.file, header=None, nrows=1).iloc[0, 0]

        if self.semester == None:
            self.semester = pd.read_excel(self.file, header=None,
                                          skiprows=1, nrows=1).iloc[0, 0]

        dateSheet = pd.read_excel(self.file, skiprows=2)

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
            self.courses.append({
                'Index': indx,
                'Code': row.Code,
                'Title': row.Course,
                'View': row.Code + ' - ' + row.Course
            })

        self.loaded = True

    def readByCode(self, codes: list[str]):
        mask = np.zeros(self.dateSheet.shape[0], dtype=bool)

        for code in codes:
            mask = mask | (self.dateSheet['Code'] == code)

        return self.dateSheet[mask]

    def readByCourse(self, courses: list[dict]):
        indexes = [course['Index'] for course in courses]

        return self.dateSheet.iloc[indexes]


@st.cache_data
def instantiate(file: any):
    return DateSheet(file)


# Global DateSheet Currently Loaded

dateSheet = DateSheet()
result = None

# Initializing Page Settings

st.set_page_config(
    page_title="Datesheet Analyzer",
    page_icon="icons/favicon.svg",
    menu_items={
        "Get Help": None,
        'Report a bug': 'mailto:avcton@gmail.com',
        'About': "This app is built to serve as a utility tool. Cheers ;)"
    },
    layout="wide")

st.title(":rainbow[FAST] Datesheet Analyzer")

# Fetching File from User

uploaded_file = st.file_uploader(
    "Upload the Latest Excel File", type=['xls', 'xlsx', 'xlsm'])

if uploaded_file is not None:
    dateSheet = instantiate(uploaded_file)
    result = dateSheet.dateSheet

# Displaying the Interaction UI + DateSheet Uploaded

filters_col, showcase_col = st.columns(2)

with filters_col:

    filters = st.multiselect(
        "Enter your Subjects", dateSheet.courses,
        format_func=lambda course: course['View'],
        key="filters")

    if len(filters) > 0:
        result = dateSheet.readByCourse(filters)


with showcase_col:
    if result is not None:
        text_col1, text_col2 = st.columns(2)
        text_col1.caption(dateSheet.title)
        text_col2.caption(dateSheet.semester)
        st.dataframe(result, hide_index=True)

st.divider()
st.caption("<center>made with <3 by avcton</center>",
           unsafe_allow_html=True)
