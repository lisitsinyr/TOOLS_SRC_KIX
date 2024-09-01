@echo on
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
set KXLDIR=d:\tools
goto Stop
:Stop
echo Значение переменной среды KXLDIR=%KXLDIR%
rem -------------------------------------------------------------------

kix32.exe %KXLDIR%\ListDir2.kix "$Dir=%1" "$ListDirFile=%2" "$Option=%3" "$KxlDir=%KXLDIR%" 
