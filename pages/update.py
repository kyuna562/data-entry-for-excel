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
# # For reset button
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


# Read data from selected file
df = pd.read_excel(newdf, sheet_name=0)
df1 = pd.read_excel(newdf, sheet_name=1)


# Display data in table
st.table(df)

# Create options form in sidebar
st.sidebar.header("Options")
options_form = st.sidebar.form("options_form", clear_on_submit=True)
date = str(datetime.date.today())  # current date
types = options_form.selectbox(
    "Pick one", ("Cash", "Materials", "Labor", "Other"))
description = options_form.text_input("Description")
worker = options_form.multiselect("Select workers", [
                                  "Elliott", "Jake", "Branden", "Mason", "Ryan", "Scott ", "Chris ", "Caleb ", "Garrett"])
income = int(options_form.number_input("INCOME", value=0))
expense = int(options_form.number_input("EXPENSE", value=0))
add_data = options_form.form_submit_button()


worker = str(worker).replace("[", "").replace("]", "")
total_income = df1["Income/Debit"].sum() + income
total_expense = df1['Expense/Credit'].sum() + expense
net_income = total_income - total_expense
if total_income != 0:
    profit_margin = round((net_income / total_income) * 100)
else:
    profit_margin = 0

# Add data to dataframe when form is submitted
if add_data and net_income:
    new_data = {
        "Date": date,
        "Type": types,
        "Description": description,
        "Worker": worker,
        "Income/Debit": income,
        "Expense/Credit": expense,
        "Balance": net_income,
    }
    df1 = df1.append(new_data, ignore_index=True)

# Add undo button
if st.button("Undo Last Update"):
    df1.drop(df1.tail(1).index, inplace=True)
    pyautogui.hotkey("f5")

# Drop specific rows
drop_row = 0
drop_row = int(st.number_input("Delete Specific Rows", value=0))
if drop_row != 0:
    df1.drop(drop_row, inplace=True)
    pyautogui.hotkey("f5")

# Display updated data in table
st.table(df1)


# Create bar charts https://plotly.com/python/bar-charts/
fig = px.pie(df1, values='Expense/Credit', names='Type',
             labels='Type', title="Expenses Type")
fig.update_traces(textposition='inside', textinfo='percent+label')

fig1 = px.bar(
    df1,
    x="Worker",
    y="Expense/Credit",
    color="Expense/Credit",
    text="Description",
    range_y=[0, 10000],
    hover_name="Date",
    range_color=[0, 10000],
    title="Labor Expense",
)


# # Set page layout for charts
col1, col2, col3 = st.columns((1, 1, 1))


# # Display charts
with col1:
    st.plotly_chart(fig, use_container_width=True)
with col2:
    st.plotly_chart(fig1, use_container_width=True)
with col3:
    st.markdown(
        f" ## Total Income ${total_income}")
    st.markdown(
        f" ## Total Expense ${total_expense}")
    st.markdown(
        f" ## Net Income ${net_income}")
    st.markdown(
        f" # Profit Margin {profit_margin}%")


# Use ExcelWriter to save changes to file
with pd.ExcelWriter(newdf) as writer:
    df.to_excel(writer, sheet_name="jc", index=False)
    df1.to_excel(writer, sheet_name="accounting", index=False)
