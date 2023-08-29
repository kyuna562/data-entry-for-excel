import pandas as pd
import plotly.express as px
import streamlit as st
import datetime
import os
import pyautogui


def setup_page():
    st.set_page_config(layout="wide", page_title="JC", page_icon="bar_chart")
    st.title("Update Accounting")


def get_files():
    res = []
    dir_path = os.getcwd()
    for file in os.listdir(dir_path):
        if file.endswith('.xlsx'):
            res.append(file)
    res.remove("jc.xlsx")
    return res


def read_data(file):
    df = pd.read_excel(file, sheet_name=0)
    df1 = pd.read_excel(file, sheet_name=1)
    return df, df1


def display_table(df):
    st.table(df)


def get_form_input():
    st.sidebar.header("Options")
    options_form = st.sidebar.form("options_form", clear_on_submit=True)
    date = str(datetime.date.today())
    types = options_form.selectbox(
        "Pick one", ("Cash", "Materials", "Labor", "Other"))
    description = options_form.text_input("Description")
    worker = options_form.multiselect("Select workers", [
                                      "Elliott", "Jake", "Branden", "Mason", "Ryan", "Scott ", "Chris ", "Caleb ", "Garrett"])
    income = int(options_form.number_input("INCOME", value=0))
    expense = int(options_form.number_input("EXPENSE", value=0))
    add_data = options_form.form_submit_button()

    return date, types, description, worker, income, expense, add_data


def update_data(df1, date, types, description, worker, income, expense, add_data):
    worker = str(worker).replace("[", "").replace("]", "")
    total_income = df1["Income/Debit"].sum() + income
    total_expense = df1['Expense/Credit'].sum() + expense
    net_income = total_income - total_expense
    if total_income != 0:
        profit_margin = round((net_income / total_income) * 100)
    else:
        profit_margin = 0
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
    return df1, total_income, total_expense, net_income, profit_margin


def undo_update(df1):
    if st.button("Undo Last Update"):
        df1.drop(df1.tail(1).index, inplace=True)
        pyautogui.hotkey("f5")
    return df1


def delete_row(df1):
    drop_row = 0
    drop_row = int(st.number_input("Delete Specific Rows", value=0))
    if drop_row != 0:
        df1.drop(drop_row, inplace=True)
        pyautogui.hotkey("f5")
    return df1


def create_charts(df1):
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
    return fig, fig1


def display_charts(fig, fig1, total_income, total_expense, net_income, profit_margin):
    col1, col2, col3 = st.columns((1, 1, 1))
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


def save_data(df, df1, file):
    with pd.ExcelWriter(file) as writer:
        df.to_excel(writer, sheet_name="jc", index=False)
        df1.to_excel(writer, sheet_name="accounting", index=False)


def main():
    setup_page()
    files = get_files()
    latest = len(files) - 1
    selected_file = st.selectbox("Select the Project", files, index=latest)
    df, df1 = read_data(selected_file)
    display_table(df)
    date, types, description, worker, income, expense, add_data = get_form_input()
    df1, total_income, total_expense, net_income, profit_margin = update_data(
        df1, date, types, description, worker, income, expense, add_data)
    df1 = undo_update(df1)
    df1 = delete_row(df1)
    display_table(df1)
    fig, fig1 = create_charts(df1)
    display_charts(fig, fig1, total_income, total_expense,
                   net_income, profit_margin)
    save_data(df, df1, selected_file)


if __name__ == "__main__":
    main()
