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
set KXLDIR=C:\E\TOOLS
goto Stop
:Stop
echo �������� ���������� ����� KXLDIR=%KXLDIR%
rem -------------------------------------------------------------------

kix32.exe %KXLDIR%\List.kix "$KxlDir=%KXLDIR%" "$Format=0" "$log=a.bat" %1
