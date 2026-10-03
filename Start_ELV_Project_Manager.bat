@echo off
title ELV Project Manager
cd /d "%~dp0"
python ELV_Project_Manager.py
if errorlevel 1 (
    echo.
    echo ELV Project Manager stopped with an error.
    pause
)
