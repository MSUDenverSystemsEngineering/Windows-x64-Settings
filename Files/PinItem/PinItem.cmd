:: ********************************************************************************
:: Sample command shell script to demonstrate the usage of PinItem.vbs
:: ********************************************************************************

:: ********************************************************************************
:: Clear all environment variables
:: ********************************************************************************
@echo off
set ITEM=
set TASKBAR=
set USAGE=


:: ********************************************************************************
:: Set environment variables for PinItem.vbs switch values
:: Commemt out a line to not use a switch
:: ********************************************************************************
:: set ITEM=/item:"%windir%\System32\calc.exe"
set ITEM=/item:"%%CSIDL_COMMON_PROGRAMS%%\Accessories\Calculator.lnk"
set TASKBAR=/taskbar
:: set USAGE=/?


:: ********************************************************************************
:: Execute PinToStartMenu.vbs
:: ********************************************************************************
echo on
:: Pin to Start Menu
cscript //nologo PinItem.wsf %ITEM% %USAGE%

:: Pin to Taskbar
cscript //nologo PinItem.wsf %ITEM% %TASKBAR% %USAGE%
