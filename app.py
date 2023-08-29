# pip install openpyxl
import pandas as pd
# the bar charts
import plotly.express as px
# web interface
import streamlit as st
# current date
import datetime


# Set page configuration
st.set_page_config(layout="wide", page_title="JC",
                   page_icon="bar_chart")
st.title("New Project")

# Read data from Excel file
excel_file = "jc.xlsx"
df = pd.read_excel(excel_file, sheet_name=0)
df1 = pd.read_excel(excel_file, sheet_name=1)


# Create options form in sidebar
st.sidebar.header("Options")
options_form = st.sidebar.form("options_form", clear_on_submit=True)
date = str(datetime.date.today())  # current date
customer = options_form.text_input("Customer Name")
phone_number = options_form.text_input("Phone Number")
address = options_form.text_input("Address")
city = options_form.text_input("City")
material = options_form.text_input("Materials")
project = options_form.text_input("Contract Value")
scope = options_form.text_area("Scope of Work")
worker = options_form.multiselect("Select Workers:", [
                                  "Elliott", "Jake", "Branden", "Mason", "Ryan", "Scott ", "Chris ", "Caleb ", "Garrett"])
add_data = options_form.form_submit_button()

# mapping the updated values by creating a dictionary.
if add_data:
    new_data = {
        "Date": date,
        "Customer Name": customer,
        "Phone Number": phone_number,
        "Address": address,
        "City": city,
        "Materials": material,
        "Contract Value": project,
        "Scope of Work": scope,
        "Workers": worker,
    }
    df = df.append(new_data, ignore_index=True)

    file_name = (
        f"JC_{df.iloc[0, 1]}_{df.iloc[0, 2]}_{df.iloc[0, 3]}_{df.iloc[0, 4]}")

    with pd.ExcelWriter(file_name + ".xlsx") as writer:
        df.to_excel(writer, sheet_name="jc", index=False)
        df1.to_excel(writer, sheet_name="accounting", index=False)

# Display updated data in table
st.table(df)
