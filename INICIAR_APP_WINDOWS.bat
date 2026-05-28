@echo off
cd /d %~dp0
python -m pip install --upgrade pip
pip install -r requirements.txt
streamlit run app.py
pause
