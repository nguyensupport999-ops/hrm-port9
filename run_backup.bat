@echo off
cd /d D:\HRM_Port
C:\Users\Admin\AppData\Local\Programs\Python\Python310\python.exe backup_nv.py
echo Backup completed at %date% %time% >> backup_log.txt