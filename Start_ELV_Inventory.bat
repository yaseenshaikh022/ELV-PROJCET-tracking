@echo off
title ELV Project Inventory Manager
cd /d "%~dp0"
python ELV_Inventory.py
if errorlevel 1 (
 echo.
 echo ELV Inventory stopped with an error.
 pause
)
