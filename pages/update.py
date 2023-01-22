# data manipulation
import pandas as pd
# the bar charts
import plotly.express as px
# web interface
import streamlit as st
# current date
import datetime
# To find excel files
import os
# For reset button
import pyautogui


# Set page configuration
st.set_page_config(layout="wide", page_title="JC", page_icon="bar_chart")
st.title("Update Accounting")


# list to store files
res = []

# Directory of .xlsx files
dir_path = os.getcwd()

# Find all .xlsx files in the current working directory
for file in os.listdir(dir_path):
    if file.endswith('.xlsx'):
        res.append(file)

# removing main file
res.remove("jc.xlsx")

# Select the most recent project file
latest = len(res) - 1
newdf = st.selectbox("Select the Project", res, index=latest)

# newdf = st.selectbox("Select the Project",options = res)


# Read data from selected file
df = pd.read_excel(newdf,sheet_name =0)
df1 = pd.read_excel(newdf,sheet_name =1)


# Display data in table
st.table(df)

# Create options form in sidebar
st.sidebar.header("Options")
options_form = st.sidebar.form("options_form",clear_on_submit=True)
date = str(datetime.date.today()) # current date
deposit =options_form.text_input("Deposit")
worker = options_form.multiselect("Select workers",["Elliott","Jake","Braden","Mason","Ryan","Scott ","Chris ","Caleb ","Garrett"])
material =options_form.text_input("Description of Material Cost")
other =options_form.text_input("Description of Other Cost")
money =int(options_form.number_input("$IN/OUT",value=0))
add_data = options_form.form_submit_button()


# Convert worker list to string
worker=str(worker)
worker = worker.replace("[","")
worker = worker.replace("]","")

# Add data to dataframe when form is submitted
if add_data:
    new_data = {
        "Date": date,
        "Deposit": deposit,
        "Worker": worker,
        "Material Cost": material,
        "Other Cost": other,
        "$IN/OUT": money,
    }
    df1 = df1.append(new_data, ignore_index=True)

# Add undo button
if st.button("Undo Last Update"):
    df1.drop(df1.tail(1).index, inplace=True)

# Display updated data in table
st.table(df1)


# Create bar charts https://plotly.com/python/bar-charts/
fig = px.bar(
    df1,
    x="Deposit",
    y="$IN/OUT",
    color="$IN/OUT",
    text="$IN/OUT",
    range_y=[0, 10000],
    hover_name="Date",
    range_color=[0, 10000],
    labels={"$IN/OUT": "$IN"},
)
fig1 = px.bar(
    df1,
    x="Worker",
    y="$IN/OUT",
    color="$IN/OUT",
    text="$IN/OUT",
    range_y=[0, 10000],
    hover_name="Date",
    range_color=[0, 10000],
    labels={"$IN/OUT": "$OUT"},
)
fig2 = px.bar(
    df1,
    x="Material Cost",
    y="$IN/OUT",
    color="$IN/OUT",
    text="$IN/OUT",
    range_y=[0, 10000],
    hover_name="Date",
    range_color=[0, 10000],
    labels={"$IN/OUT": "$OUT"},
)
fig3 = px.bar(
    df1,
    x="Other Cost",
    y="$IN/OUT",
    color="$IN/OUT",
    text="$IN/OUT",
    range_y=[0, 10000],
    hover_name="Date",
    range_color=[0, 10000],
    labels={"$IN/OUT": "$OUT"},
)

# Set page layout for bar charts
col1, col2, col3, col4 = st.columns((1, 1, 1, 1))


# Display bar charts
with col1:
    st.plotly_chart(fig, use_container_width=True)
with col2:
    st.plotly_chart(fig1, use_container_width=True)
with col3:
    st.plotly_chart(fig2, use_container_width=True)
with col4:
    st.plotly_chart(fig3, use_container_width=True)

# Use ExcelWriter to save changes to file
with pd.ExcelWriter(newdf) as writer:
    df.to_excel(writer, sheet_name="jc" , index=False)
    df1.to_excel(writer, sheet_name="accounting" , index=False)
  
# Add reset button to refresh the charts
if st.button("Reset Charts"):
    pyautogui.hotkey("r")
