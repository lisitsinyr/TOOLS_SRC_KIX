@echo on
rem -------------------------------------------------------------------
if "%KXLDIR%" == "" goto Begin
goto Stop

:Begin
echo «начение переменной среды KXLDIR не установлено
if "%COMPUTERNAME%" == "%USERDOMAIN%" goto Local

:Network
set KXLDIR=\\S73FS01\APPInfo\tools
goto Stop

:Local
set KXLDIR=D:\PROJECTS\я«џ »\Python\TOOLS\UDF
goto Stop

:Stop
echo «начение переменной среды KXLDIR=%KXLDIR%
rem -------------------------------------------------------------------

kix32.exe ListDir.kix "$KxlDir=%KXLDIR%" "$Format=%1" "$NLevel=%2"
