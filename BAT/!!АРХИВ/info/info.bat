@echo off
rem -------------------------------------------------------------------
if "%1" == "" set LogFile=%TEMP%\info.log
goto Kix 
set LogFile=%1
:Kix

rem -------------------------------------------------------------------
if "%KXLDIR%" == "" goto Begin
goto Stop
:Begin
echo Значение переменной среды KXLDIR не установлено
if "%COMPUTERNAME%" == "%USERDOMAIN%" goto Local
:Network
set KXLDIR=\\S73FS01\APPInfo\tools
goto Stop
:Local
set KXLDIR=c:\tools
goto Stop
:Stop
echo Значение переменной среды KXLDIR=%KXLDIR%
rem -------------------------------------------------------------------

kix32.exe %KXLDIR%\info.kix "$LogFile=%LogFile%" "$KxlDir=%KXLDIR%"

if "%1" == "none" goto end
if exist %LogFile% start %LogFile%
:end
