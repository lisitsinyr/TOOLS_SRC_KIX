@echo on
rem -------------------------------------------------------------------
if "%KXLDIR%" == "" goto Begin
goto Stop

:Begin
echo �������� ���������� ����� KXLDIR �� �����������
if "%COMPUTERNAME%" == "%USERDOMAIN%" goto Local

:Network
set KXLDIR=\\S73FS01\APPInfo\tools
goto Stop

:Local
set KXLDIR=D:\PROJECTS\�����\Python\TOOLS\UDF
goto Stop

:Stop
echo �������� ���������� ����� KXLDIR=%KXLDIR%
rem -------------------------------------------------------------------

kix32.exe ListDir.kix "$KxlDir=%KXLDIR%" "$Format=%1" "$NLevel=%2"
