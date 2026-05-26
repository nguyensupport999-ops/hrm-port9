@echo off
title HRM-Port

cd /d D:\HRM_Port
start "" /B streamlit run app.py --server.headless true 2>nul

timeout /t 5 >nul
start "" http://localhost:8501

exit