# PEW (Python Excel Writer) - A Web Application for Project Management and Accounting.
![Capture2](https://github.com/kyuna562/data-entry-for-excel/assets/47129687/a86140c5-c89f-42ff-a8fd-927374da314a)

* This web application offers users an interface to generate new Excel files for tracking projects with an emphasis on accounting. Developed using Streamlit, the app facilitates financial data entry, enables data manipulation, and presents visualizations that highlight key project accounting metrics

### Getting Started:
1. **Clone/Download the Repository**: 
   ```bash
   git clone https://github.com/kyuna562/data-entry-for-excel.git
   ```
   Alternatively, you can directly download it from the repository page.

### Using the Start File:
2. **Automatic Setup**:
   Simply launch the **start+requirements** file. This action will install the necessary libraries from the `requirements.txt` and initiate the app.

### Manual Installation and Running:
3. **Navigate to the Project Directory**:
   Open your terminal or command prompt and change the directory to where you've cloned/downloaded the project.
   
4. **Install Dependencies**: 
   ```bash
   pip install -r requirements.txt
   ```

5. **Start the App**: 
   ```bash
   streamlit run app.py
   ```

6. **Access the App**: 
   Once started, the app will automatically launch in your default web browser.


## Features
- **Project Data Management**: Users can add, edit, and view project-related data.
- **Accounting**: View and manage financial data, including income, expenses, and profit margins.
- **Data Visualization**: Interactive charts display financial metrics and expense types.
- **Excel Integration**: Read data from and write to Excel files.
- **User-Friendly UI**: The application offers functionalities such as undoing updates, deleting specific rows, and editing data directly in tables.


## Libraries Used
- **pandas**: used for data manipulation and analysis.
- **plotly**: used for creating interactive charts.
- **streamlit**: used for creating the web interface.
- **Openpyxl**: To read from and write to Excel files.
- **os**: used for interacting with the operating system, in this case to find Excel files in the current directory.
- **datetime**: used to get the current date.

## Functions
- `set_page_config()`: sets the page configuration.
- `read_file(file_name)`: reads the first two sheets of an Excel file.
- `add_new_data(df, date, types, description, worker, income, expense, net_income)`: adds new data to a DataFrame.
- `excel_list()`: returns a list of all Excel files in the current directory, except 'pew.xlsx'.
- `current_index(res)`: returns the index of the current Excel file.
- `save_excel_name(file_name)`: saves the name of the current Excel file to a .txt file.
- `total_income(df1)`, `total_expense(df1)`, `net_income(total_income, total_expense)`, `profit_margin(net_income, total_income)`: calculate metrics.
- `file_names(new_list)`: allows the user to select an Excel file.
- `main()`: main function that runs the script.

## Notes
- You cannot concurrently open an Excel file that is already active in the application when making edits; otherwise, an error will occur.
- The app reads and writes data to and from an Excel file named "pew.xlsx". This file must be present in the project directory in order for the app to function properly.
- The app also creates new Excel files for each new project, using the customer's name, phone number, address, and city as the file name. These files will be stored in the project directory.
- The app is designed to be run locally and is not intended to be deployed on a public server