@echo off
setlocal
cd /d "%~dp0"
python -m streamlit run app.py
if errorlevel 1 (
  echo.
  echo Procell failed to start. Confirm Python and requirements are installed.
  pause
)
endlocal
