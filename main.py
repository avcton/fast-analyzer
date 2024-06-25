import streamlit as st
from models.datesheet import DateSheet


@st.cache_data(persist="disk")
def instantiate(file: any, title: str = None, semester: str = None):
    dateSheet = DateSheet(title, semester)
    return dateSheet, dateSheet.load(file)


# Global DateSheet / Result Currently Loaded
dateSheet = result = None

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
    dateSheet, loaded = instantiate(uploaded_file)
    if not loaded:
        st.error('The provided sheet is not in a format as that of a typical FAST LHR date sheet. Kindly verify the integrity of the date sheet.',
                 icon=":material/sd_card_alert:")
    else:
        result = dateSheet.dateSheet

# Displaying the Interaction UI + DateSheet Uploaded
filters_col, showcase_col = st.columns(2)

with filters_col:
    filters = st.multiselect(
        "Enter your Subjects", dateSheet.courses if dateSheet else [],
        format_func=lambda course: course['View'],
        key="filters")

    if dateSheet and len(filters) > 0:
        result = dateSheet.readByCourse(filters)

with showcase_col:
    if result is not None:
        text_col1, text_col2 = st.columns(2)
        text_col1.caption(dateSheet.title)
        text_col2.caption(dateSheet.semester)
        st.dataframe(result, hide_index=True)

st.divider()
st.caption("<center>made with <3 by <a href=\"https://avcton.github.io/\">avcton</a></center>",
           unsafe_allow_html=True)
