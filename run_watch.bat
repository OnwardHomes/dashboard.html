@echo off
setlocal
REM Launch the watcher and keep the PowerShell window open.
REM Double-click this file to start watching Vacancy_Data/ and Arrears_Data/.
powershell.exe -NoLogo -NoExit -ExecutionPolicy Bypass -File "%~dp0build_watch.ps1"
