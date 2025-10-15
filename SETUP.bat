@echo off
echo ===============================
echo  Setup pentru Generator N-ERP
echo ===============================

:: Creează mediul virtual (dacă nu există)
if not exist .venv (
    echo Creare mediu virtual...
    python -m venv .venv
)

:: Activează mediul virtual
call .venv\Scripts\activate

:: Instalează pachetele
echo Instalare pachete din requirements.txt...
pip install --upgrade pip
pip install -r requirements.txt

:: Pornește aplicația Streamlit
echo Lansare aplicație Streamlit...
streamlit run app.py

pause
