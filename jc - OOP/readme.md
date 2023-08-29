# JC

* This is a web application built using Python and the Streamlit library, designed to help a small painting company manage their projects and accounting. 

Features
- Add new projects and input customer information, materials, and scope of work
- View and update a list of current projects
- Add and view accounting data for each project, including deposits, worker payments, material costs, and other expenses
- Create and view bar charts displaying the financial data for each project
- Undo the last update or reset the form

Requirements
- Python 3.6 or higher
- Pandas
- Plotly
- Streamlit

Running the app
1. Clone or download the repository
2. Navigate to the project directory in your terminal
3. Run the command pip install -r requirements.txt to install the required libraries
4. Run the command streamlit run app.py to start the app
5. The app will open in your default web browser

Notes
- The app reads and writes data to and from an Excel file named "jc.xlsx". This file must be present in the project directory in order for the app to function properly.
- The app also creates new Excel files for each new project, using the customer's name, phone number, address, and city as the file name. These files will be stored in the project directory.
- The app is designed to be run locally and is not intended to be deployed on a public server.