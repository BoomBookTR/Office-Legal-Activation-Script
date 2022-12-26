:: Bu da admin yetkili alýyor.
:: Admin Yetki kodu BAÞI
:: @pushd %~dp0 & fltmc | find "." && (powershell start '%~f0' ' %*' -verb runas 2>nul && exit /b)
:: Admin Yetki kodu SONU
::---------------------------------------------------
:: Admin Yetki kodu BAÞI
@echo off
if not "%1"=="am_admin" (powershell start -verb runas '%0' am_admin & exit /b)
:: Admin Yetki kodu SONU
::---------------------------------------------------

:: .... your code start ....

for %%a in (4,5,6) do (if exist "%ProgramFiles%\Microsoft Office\Office1%%a\ospp.vbs" (cd /d "%ProgramFiles%\Microsoft Office\Office1%%a")
if exist "%ProgramFiles(x86)%\Microsoft Office\Office1%%a\ospp.vbs" (cd /d "%ProgramFiles(x86)%\Microsoft Office\Office1%%a")) & cls
cscript ospp.vbs /dinstid > "%~dp0"\IID_office.txt
pushd "%~dp0"
start IID_office.txt
exit