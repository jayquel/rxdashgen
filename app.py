# @Email:  contact@pythonandvba.com
# @Website:  https://pythonandvba.com
# @YouTube:  https://youtube.com/c/CodingIsFun
# @Project:  Sales Dashboard w/ Streamlit
# youtube guide: https://www.youtube.com/watch?v=nJHrSvYxzjE&t=0s


import pandas as pd  # pip install pandas openpyxl
import plotly.express as px  # pip install plotly-express
import streamlit as st  # pip install streamlit
import base64 # Standard Python Module
from io import StringIO, BytesIO # Standard Python Module

def generate_excel_download_link(df):
    # Credit Excel: https://discuss.streamlit.io/t/how-to-add-a-download-excel-csv-function-to-a-button/4474/5
    towrite = BytesIO()
    df.to_excel(towrite, encoding="utf-8", index=False, header=True)  # write to BytesIO buffer
    towrite.seek(0)  # reset pointer
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="Recommended_Dashboard.xlsx">Download Excel File</a>'
    return st.markdown(href, unsafe_allow_html=True)

# emojis: https://www.webfx.com/tools/emoji-cheat-sheet/
st.set_page_config(page_title="PGS Dashboard", page_icon=":bar_chart:", layout="wide")



# ---- MAINPAGE ----
st.title(":clipboard: PGS Dashboard")
st.markdown("##")



# ---- READ EXCEL ----
@st.cache
def get_data_from_excel():
    df = pd.read_excel(
        io="rxpgs_jobfamilies.xlsx",
        engine="openpyxl",
        sheet_name="Sales",
        skiprows=3,
        usecols="B:R",
        nrows=1000,
    )
    # Add 'hour' column to dataframe
    df["hour"] = pd.to_datetime(df["Time"], format="%H:%M:%S").dt.hour
    return df

df = get_data_from_excel()

# ---- SIDEBAR ----
st.sidebar.header("Please Filter Here:")
city = st.sidebar.multiselect(
    "Select the City:",
    options=df["City"].unique(),
    default=df["City"].unique()
)

customer_type = st.sidebar.multiselect(
    "Select the Customer Type:",
    options=df["Customer_type"].unique(),
    default=df["Customer_type"].unique(),
)

gender = st.sidebar.multiselect(
    "Select the Gender:",
    options=df["Gender"].unique(),
    default=df["Gender"].unique()
)

df_selection = df.query(
    "City == @city & Customer_type ==@customer_type & Gender == @gender"
)


st.dataframe(df_selection)


# TOP KPI's
total_measures = int(df_selection["Invoice ID"].count())
average_rating = round(df_selection["Rating"].mean(), 1)
star_rating = ":star:" * int(round(average_rating, 0))
average_sale_by_transaction = round(df_selection["Total"].mean(), 2)

left_column, right_column = st.columns(2)
with left_column:
    st.subheader("Total No. of Measures:")
    st.subheader(f"{total_measures:,}") # use this if you want to include comma in the figures (f"{total_measures:,}")

with right_column:
    st.subheader("Priority Matrix Rating:")
    st.subheader(f"US $ {average_sale_by_transaction}")

st.markdown("""---""")


# -- DOWNLOAD EXCEL FILE

st.subheader('Download:')
generate_excel_download_link(df_selection)

# ---- HIDE STREAMLIT STYLE ----
hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)
