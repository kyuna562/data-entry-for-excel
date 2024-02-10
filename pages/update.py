# To find excel files
import os
# data manipulation
import pandas as pd
# the charts
import plotly.express as px
import plotly.graph_objects as go
# web interface
import streamlit as st
# current date
import datetime

worker_list = ["Elliott", "Jake", "Branden", "Mason",
               "Ryan", "Scott ", "Chris ", "Caleb ", "Garrett"]

color_map = {"Cash":"#ffffff","Materials":"#b491ff","Labor":"#82afff", "Equipment":"#78dfb9","Overhead/Other":"#ffb6de"}


# Function to set up the Streamlit page's basic attributes
def set_page_config():
    st.set_page_config(layout="wide", page_title="PEW", page_icon="bar_chart")
    st.title("Accounting")


# Function to add new accounting data to a DataFrame
def add_new_data(df, date, types, description, worker, income, expense, net_income):
    new_data = {
        "Date": date,
        "Type": types,
        "Description": description,
        "Worker": worker,
        "Income/Debit": income,
        "Expense/Credit": expense,
        "Balance": net_income,
    }
    return pd.concat([df, pd.DataFrame([new_data])], ignore_index=True)



# Metric and other utility calculations
def total_income(df1):
    total_income = df1["Income/Debit"].sum()
    return total_income


def total_expense(df1):
    total_expense = df1['Expense/Credit'].sum()
    return total_expense


def net_income(total_income, total_expense):
    net_income = total_income - total_expense
    return net_income


def profit_margin(net_income, total_income):
    profit_margin = round((net_income / total_income) *100) if total_income != 0 else 0
    return profit_margin

# Aggregating the data("See the summary")
def group_expense(df1):
    return df1.groupby('Type')[["Income/Debit","Expense/Credit"]].sum()

# Bar chart showing expenses over time
def fig2_bar_chart(df1):
    fig2 = px.bar(df1, x="Date", y="Expense/Credit", color='Type',
                    color_discrete_map=color_map, 
                    title="Expenses Over Time", text='Type', hover_data='Description', hover_name="Worker")
    
    # Summarize the total amount by day
    dft = df1.groupby('Date')[["Income/Debit","Expense/Credit"]].sum()
    fig2.add_trace(go.Scatter(
    x=dft.index, 
    y= dft['Expense/Credit'],
    text=dft['Expense/Credit'],
    mode='text',
    textposition='top center'))

    # For 1m/6m/YTD/all button
    fig2.update_xaxes(
        #rangeslider_visible=True,
        rangeselector=dict(
        buttons=list([
            dict(count=1, label="1m", step="month", stepmode="backward"),
            dict(count=6, label="6m", step="month", stepmode="backward"),
            dict(count=1, label="YTD", step="year", stepmode="todate"),
            dict(step="all")
            ])))
    return fig2




# Function to retrieve a list of Excel files in the directory, excluding 'pew.xlsx'
def excel_list():
    return [file for file in os.listdir(os.getcwd()) if file.endswith('.xlsx') and file != "pew.xlsx"]

# Function to get the index of the current working Excel file
def current_index(res):
    with open('text.txt', 'r') as r:
        content = r.read()
        if content in res:
            latest = res.index(content)
        else:
            latest = 0
    return latest

# Creating a session state for file_names
new_list = excel_list()
if 'new_index' not in st.session_state:
    st.session_state['new_index'] = current_index(new_list)

# Function to read data from two sheets of an Excel file
def read_file(file_name):
    return pd.read_excel(file_name, sheet_name=0), pd.read_excel(file_name, sheet_name=1)


def file_names(new_list):
    file_name = st.selectbox("Select the Project",  new_list,
                             index=st.session_state['new_index'])
    return file_name

# Saving selected excel file to .txt
def save_excel_name(file_name):
    with open('text.txt', 'w') as text_file:
        text_file.write(file_name)




# Use ExcelWriter to save changes to excel file
def xls_writer(df, df1, file_name):
    with pd.ExcelWriter(file_name) as writer:
        df.to_excel(writer, sheet_name="pew", index=False)
        df1.to_excel(writer, sheet_name="accounting", index=False)

# Generates and displays interactive charts related to project expenses.
def create_and_display_charts(df1): 
    if len(df1) != 0: # Only run if df1 is not empty
        # Pie chart for expense types
        fig = px.pie(df1, values='Expense/Credit', names=df1['Type'],color='Type', color_discrete_map=color_map,
                    labels='Type', title="Expenses Type")
        fig.update_traces(textposition='inside', textinfo='percent+label')

        # Bar chart for labor expenses by worker
        fig1 = px.bar(df1,
                    x="Worker",
                    y="Expense/Credit",
                    #color="Expense/Credit",
                    hover_name="Date",
                    title="Labor Expense")

    
        # Displaying charts and metrics on Streamlit
        col1, col2, col3 = st.columns((1, 1, 1))
    
        with col1:
            st.plotly_chart(fig, use_container_width=True)
        with col2:
            st.plotly_chart(fig1, use_container_width=True)
        with col3:
            st.markdown(f" ## Total Income ${total_income(df1)}")
            st.markdown(f" ## Total Expense ${total_expense(df1)}")
            st.markdown(
                f" ## Net Income ${net_income(total_income(df1),total_expense(df1))}")
            st.markdown(
                f" # Profit Margin {profit_margin(net_income(total_income(df1),total_expense(df1)), total_income(df1))}%")
            with st.expander("See the summary"):
                st.table(group_expense(df1))
        st.plotly_chart(fig2_bar_chart(df1), use_container_width=True)

 
# Function to get the dictionary of data_editor() session state "editable_df" and update the values into df1
def update_df_with_data(df1, dic):
    for df_row, values in dic:
        for column_name, new_value in values.items():
            df1.at[int(df_row), column_name] = new_value
    return df1


def main():
    set_page_config()

    # File name selection
    file_name = file_names(excel_list())
    
    # Save the file_name to a .txt file
    if file_name:
        save_excel_name(file_name)
        # Update the session state to the currently selected Excel file when the 'file_name' is chosen
        # if st.session_state['new_index'] != new_list.index(file_name):
        #     st.session_state['new_index'] = current_index(file_name)

    df, df1 = read_file(file_name)
    
    # Display first excel sheet in table
    st.table(df)

    # Sidebar form for accounting entries
    options_form = st.sidebar.form("options_form", clear_on_submit=True)
    date = pd.Timestamp(datetime.date.today())
    types = options_form.selectbox(
        "Pick one 📑", ("Cash", "Materials", "Labor", "Equipment", "Overhead/Other"))
    description = options_form.text_input("Description 🖹")
    worker = str(options_form.multiselect("Select workers 👷🏻", worker_list)).replace(
        "[", "").replace("]", "").replace("'", "") 
    income = options_form.number_input(
        "Transaction value 💲", value=None, format="%d", min_value=0)
    add_data = options_form.form_submit_button()

    # Add data to dataframe when form is submitted
    if add_data:
        expense = 0
        if types !="Cash":
            expense,income = income,0
        

        df1 = add_new_data(df1, date, types, description,
                           worker, income, expense, net_income=(total_income(df1) + income) - (total_expense(df1) + expense))


    # Add undo button
    if st.button("Undo Last Update"):
        df1.drop(df1.tail(1).index, inplace=True) 
    
    # Drop specific rows
    drop_row = st.number_input(label='Delete Specific Row',
                               min_value=0,  # don't have negative row indices
                               max_value=len(df1) - 1 if len(df1) > 1 else 0,  # prevent out-of-range indices , else default value is 0
                               value=None,  # Default value
                               format="%d")  # input is treated as an integer
    if  drop_row != None:
        df1.drop(drop_row, inplace=True)
        st.success(f'Row {drop_row} deleted successfully!')


    # Creating data_editor and "editable_df" session state
    st.data_editor(df1, use_container_width=True, key="editable_df")
    


    # Apply the changes from the table to the df1
    if st.button("Submit Table Changes"):
        update_df_with_data(
            df1, st.session_state["editable_df"]["edited_rows"].items())
        st.success('Tables updated successfully!')

    # Display charts and resets every time when selected/undo/drop/added/submit
    if file_name or add_data or drop_row or st.button(
        "Submit Table Changes") or st.button("Undo Last Update") :

        create_and_display_charts(df1)

        # Use ExcelWriter to save changes to excel file
        xls_writer(df, df1, file_name)


        
   

if __name__ == "__main__":
    main()
