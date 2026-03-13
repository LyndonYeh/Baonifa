﻿@echo off
echo 安裝 PyInstaller...
python -m pip install pyinstaller

echo.
echo 開始打包...
python -m PyInstaller --onefile --noconsole ^
  --name Baonifa ^
  --icon robot.ico ^
  --add-data "robot.ico;." ^
  --hidden-import patch_to_excel ^
  --hidden-import excel_to_word ^
  --hidden-import win32com ^
  --hidden-import win32com.client ^
  --hidden-import pythoncom ^
  --hidden-import pywintypes ^
  go.py

echo.
echo 完成！exe 在 dist\Baonifa.exe
pause
