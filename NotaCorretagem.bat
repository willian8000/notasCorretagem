@echo off

IF EXIST .venv\ (
     .venv\Scripts\activate.bat;
    streamlit run .\NotaCorretagem.py

) ELSE ( 
    py -m venv .venv
    CALL .venv\Scripts\activate.bat 
    pip install -r requirements.txt 
    streamlit run .\NotaCorretagem.py
)