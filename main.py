import time
import streamlit as st
from models.timetable import TimeTable
from models.datesheet import DateSheet

# Instantiating Function


@st.cache_data(persist="disk", show_spinner=False)
def instantiate(file: any, model_type: str, title: str = None):
    model_instance = TimeTable(
        title) if model_type == 'Timetable' else DateSheet(title)
    return model_instance, model_instance.load(file)


# Model Type / Instance / Result Currently Loaded
model_type = model_instance = result = None

# Initializing Page Settings
st.set_page_config(
    page_title="FAST Analyzer",
    page_icon="icons/favicon.svg",
    menu_items={
        "Get Help": None,
        'Report a bug': 'mailto:avcton@gmail.com',
        'About': "This app is built to serve as a utility tool. Cheers ;)"
    },
    layout="wide")

st.title(":rainbow[FAST] Analyzer")
st.logo(image='icons/favicon.svg')

# ------------------------------------------------------------------------ #

# Writing Disclaimer + Creator Info
disclaimer_text = """Please be aware that this app is meant to :rainbow[**break**] someday. This is because, we only have the spreadsheets shared by the university as the scarce resource to account for this information and thus the structure to scrap and read through that data highly realies on the format University uses, which is not really that *consistant*. Thus this might be the only reason why you might find this app to not work for specific sheets. One such **example** is the **Ramazan Timetable**, which is conditionally released by the University and on an unreliable format (no periods). Thus, this app can not accomodate in going through that sheet. Sincere appoligies that I cannot help in that :("""


creater_info_text = """If you **loved** this work, want to **support** me and possibiliy enjoy more such crafted projects, you might be interested in checking out my [**technical notes**](https://avcton.github.io/Literature/) which I have been crafting for the past 3 Semesters. You even are always welcome to **contact** me by any means and I will be always up for any kind of **collaboration** too. You can **learn** more about me through [:rainbow[**my portfolio**]](https://avcton.github.io/)."""

creater_info_col, disclaimer_col = st.columns(2)

disclaimer_col.warning(disclaimer_text, icon=':material/priority_high:')


def stream_creator_info():
    for word in creater_info_text.split(" "):
        yield word + " "
        time.sleep(0.02)


creater_info_col.html("<center><h5></h3><center>")

if "streamed_creator_text" not in st.session_state or st.session_state['streamed_creator_text'] == False:
    creater_info_col.write_stream(stream_creator_info)
    st.session_state['streamed_creator_text'] = True
else:
    creater_info_col.write(creater_info_text)

# ------------------------------------------------------------------------ #

# Model Selection / File Upload UI
model_selection_col, file_upload_col = st.columns(2)

if "model_selection_disabled" not in st.session_state:
    st.session_state["model_selection_disabled"] = False

with model_selection_col:
    model_type = st.radio(
        "Select your desired **analysis** from the given options",
        ['Timetable', 'Date Sheet'],
        format_func=lambda option: "Analyze LHR Campus " + option,
        captions=['Deduce your own timetable from the given excel file',
                  'Generate your date sheet for the upcoming exams'], disabled=st.session_state["model_selection_disabled"], help='You can only select an option when no file is uploaded')


# Callback that take cares of disabling model selection if a file is uploaded
def toggle_model_selection():
    if st.session_state["file_uploader"] == None:
        st.session_state["model_selection_disabled"] = False
    else:
        st.session_state["model_selection_disabled"] = True


with file_upload_col:
    # Fetching File from User
    uploaded_file = st.file_uploader(
        "Upload the latest **Excel File** for your :rainbow[**selected analysis**]", type=['xls', 'xlsx', 'xlsm'], key='file_uploader', on_change=toggle_model_selection)

# Course Selection / Result UI
course_filter_col, result_display_col = st.columns(2)

with result_display_col:
    if uploaded_file is not None:
        with st.spinner('Hold tight!.. We are indexing your provided data...'):
            model_instance, loaded = instantiate(uploaded_file, model_type)
            if not loaded:
                st.error(f'The provided sheet is not in a format as that of a typical FAST LHR {model_type.lower()}. Kindly verify its integrity',
                         icon=":material/sd_card_alert:")
            else:
                result = model_instance.data

with course_filter_col:
    filters = st.multiselect(
        "Enter your Courses", model_instance.options if model_instance else [],
        key="filters")

if model_instance and len(filters) > 0:
    result = model_instance.readByCourse(filters)

with result_display_col:
    if result is not None:
        st.html(f'<center>{model_instance.title}<center>')
        st.dataframe(result, hide_index=True, use_container_width=True)

st.divider()
st.caption("<center>made with <3 by <a href=\"https://avcton.github.io/\">avcton</a></center>",
           unsafe_allow_html=True)
