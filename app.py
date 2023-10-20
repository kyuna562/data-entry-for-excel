# pip install openpyxl
import pandas as pd
# the bar charts
import plotly.express as px
# web interface
import streamlit as st
# current date
import datetime

# List of worker names
worker_list = ["Elliott", "Jake", "Branden", "Mason",
               "Ryan", "Scott ", "Chris ", "Caleb ", "Garrett"]


# Function to set up the Streamlit page's basic attributes
def set_page_config():
    st.set_page_config(layout="wide", page_title="PEW", page_icon="bar_chart")
    st.title("New Project")


# Function to read data from two sheets of an Excel file
def read_file(file_name):
    return pd.read_excel(file_name, sheet_name=0), pd.read_excel(file_name, sheet_name=1)

# Main function to run the web app
def main():
    set_page_config()  # Set the page configuration
    df, df1 = read_file("pew.xlsx")  # Read data from the Excel file

    # Sidebar for Project input options
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
    worker = str(options_form.multiselect("Select workers", worker_list)).replace(
        "[", "").replace("]", "").replace("'", "")
    add_data = options_form.form_submit_button()

# When the user submits data, process and save it
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
        df = pd.concat([df, pd.DataFrame([new_data])], ignore_index=True)

        # Custom file name
        file_name = (
            f"PEW_{df.iloc[0, 1]}_{df.iloc[0, 2]}_{df.iloc[0, 3]}_{df.iloc[0, 4]}")

        # Saving it to excel
        with pd.ExcelWriter(file_name + ".xlsx") as writer:
            df.to_excel(writer, sheet_name="pew", index=False)
            df1.to_excel(writer, sheet_name="accounting", index=False)

    # Display updated data in table
    st.table(df)


if __name__ == "__main__":
    main()
