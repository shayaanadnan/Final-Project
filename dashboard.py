import pandas as pd  # pip install pandas openpyxl
import plotly.express as px  # pip install plotly-express
import streamlit as st  # pip install streamlit

st.set_page_config(page_title="Dashboard", 
                   page_icon=":bar_chart:", 
                   layout="wide"
)

@st.cache_data
def get_data_from_excel():
    df = pd.read_excel(
        io="people.xlsx",
        engine="openpyxl",
        sheet_name="Sheet1",
        skiprows=0,
        usecols="A:G",
        nrows=9,
    )
    # Add 'hour' column to dataframe
    # df["hour"] = pd.to_datetime(df["Time"], format="%H:%M:%S").dt.hour
    return df

# st.dataframe(df)
df = get_data_from_excel()

# Side Bar.
st.sidebar.header("Filter your requirment:")
age = st.sidebar.multiselect(
    "Select the age",
    options=df["Age"].unique(),
    default=df["Age"].unique()

)

gender = st.sidebar.multiselect(
    "Select the Gender",
    options=df["Gender"].unique(),
    default=df["Gender"].unique()

)

coursename = st.sidebar.multiselect(
    "Select the Course Name",
    options=df["CourseName"].unique(),
    default=df["CourseName"].unique()

)

df_selection = df.query(
    "Age == @age & Gender == @gender & CourseName == @coursename"
)

st.title(":bar_chart: Bano Qabil 3.0 Dashboard")
st.markdown("##")

total_students = len(df_selection["Name"])
total_male = len(df_selection[df_selection["Gender"]=="Male"])
total_female = len(df_selection[df_selection["Gender"]=="female"])
total_course = len(df_selection["CourseName"])

left_column, middle_column, middle_column, right_column = st.columns(4)
with left_column:
    st.subheader("Total # of Students")
    st.subheader(f"{total_students}")
with middle_column:
    st.subheader("Total Male Students")
    st.subheader(f"{total_male}")
with middle_column:
    st.subheader("Total female Students")
    st.subheader(f"{total_female}")
with right_column:
    st.subheader("Total # of Courses")
    st.subheader(f"{total_course}")

st.markdown("---")


total_students = len(df_selection["Name"])
total_male = len(df_selection[df_selection["Gender"] == "Male"])
total_female = len(df_selection[df_selection["Gender"] == "female"])
total_course = len(df_selection["CourseName"])

data = {
    "Category": ["Total # of Students", "Total Male Students", "Total Female Students", "Total # of Courses"],
    "Count": [total_students, total_male, total_female, total_course]
}

df_chart = pd.DataFrame(data)

fig = px.bar(df_chart, x="Category", y="Count", color="Category",
             labels={"Count": "Count", "Category": "Category"},
             title="Summary Data")

st.plotly_chart(fig)

st.markdown("---")