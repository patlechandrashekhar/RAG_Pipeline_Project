@echo off
echo Cleaning Python cache...
rmdir /S /Q src\page_indexing_rag\__pycache__ 2>NUL
rmdir /S /Q app\__pycache__ 2>NUL
rmdir /S /Q src\__pycache__ 2>NUL
del /S /Q *.pyc 2>NUL

echo.
echo Starting Streamlit app...
streamlit run app/streamlit_app.py