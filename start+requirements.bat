@echo off

cd %CD%
pip install -r requirements.txt
streamlit run app.py

pause
