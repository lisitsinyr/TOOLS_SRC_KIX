@echo off
rem -------------------------------------------------------------------
rem ListFile_kix.bat
rem -------------------------------------------------------------------
chcp 1251>NUL

setlocal enabledelayedexpansion

rem -------------------------------------------------------------------
rem SCRIPTS_DIR - Каталог скриптов
rem -------------------------------------------------------------------
if not defined SCRIPTS_DIR (
    set SCRIPTS_DIR=D:\TOOLS\TOOLS_BAT
    set SCRIPTS_DIR=D:\PROJECTS_LYR\CHECK_LIST\SCRIPT\04_BAT\PROJECTS_BAT\TOOLS_SRC_BAT
)
rem echo SCRIPTS_DIR: %SCRIPTS_DIR%
rem -------------------------------------------------------------------
rem LIB_BAT - каталог библиотеки скриптов
rem -------------------------------------------------------------------
if not defined LIB_BAT (
    set LIB_BAT=!SCRIPTS_DIR!\SRC\LIB
    rem echo LIB_BAT: !LIB_BAT!
)
if not exist !LIB_BAT!\ (
    echo ERROR: Каталог библиотеки LYR !LIB_BAT! не существует...
    exit /b 0
)
rem -------------------------------------------------------------------
rem SCRIPTS_DIR_KIX - Каталог скриптов KIX
rem -------------------------------------------------------------------
if not defined SCRIPTS_DIR_KIX (
    set SCRIPTS_DIR_KIX=D:\TOOLS\TOOLS_KIX
    set SCRIPTS_DIR_KIX=D:\PROJECTS_LYR\CHECK_LIST\SCRIPT\01_KIX\PROJECTS_KIX\TOOLS_SRC_KIX\SCRIPTS
                        
)
rem echo SCRIPTS_DIR_KIX: !SCRIPTS_DIR_KIX!

rem --------------------------------------------------------------------------------
rem 
rem --------------------------------------------------------------------------------
:begin
    set BATNAME=%~nx0
    echo Start !BATNAME! ...

    rem Количество аргументов
    call :Read_N %* || exit /b 1
    call :SET_LIB %0 || exit /b 1

    rem -------------------------------------
    rem OPTION
    rem -------------------------------------
    set O1=O1
    set PN_CAPTION=O1
    call :Read_P O1 O1 || exit /b 1
    rem echo O1:!O1!
    if defined O1 (
        set OPTION=!OPTION! --O1 !O1!
    ) else (
        echo INFO: O1 not defined ...
    )

    rem -------------------------------------
    rem ARGS
    rem -------------------------------------
    rem Проверка на обязательные аргументы
    set Dir=.
    set PN_CAPTION=Каталог
    call :Read_P Dir %1 || exit /b 1
    rem echo Dir: !Dir!
    if defined Dir (
        set ARGS=!ARGS! !Dir!
    ) else (
        echo ERROR: Dir not defined ...
        set OK=
    )

    rem set APP_KIX_DIR=D:\PROJECTS_LYR\CHECK_LIST\SCRIPT\01_KIX\PROJECTS_KIX\TOOLS_KIX\KIX
    set APP_KIX=ListFile

    call :SET_KIX || exit /b 1

    if exist !APP_KIX_DIR!\!APP_KIX!.kix (
        echo START !APP_KIX_DIR!\!APP_KIX!.kix ... 
        kix32.exe !APP_KIX_DIR!\!APP_KIX!.kix "$Dir=!Dir!"
    )

    rem call :PressAnyKey || exit /b 1

    exit /b 0
:end
rem --------------------------------------------------------------------------------

rem =================================================
rem ФУНКЦИИ LIB
rem =================================================
rem =================================================
rem LYRConst.bat
rem =================================================
:SET_LIB
%LIB_BAT%\LYRConst.bat %*
exit /b 0
:SET_KIX
%LIB_BAT%\LYRConst.bat %*
exit /b 0
rem =================================================
rem LYRDateTime.bat
rem =================================================
rem =================================================
rem LYRFileUtils.bat
rem =================================================
rem =================================================
rem LYRLog.bat
rem =================================================
rem =================================================
rem LYRStrUtils.bat
rem =================================================
rem =================================================
rem LYRSupport.bat
rem =================================================
:Read_N
%LIB_BAT%\LYRSupport.bat %*
exit /b 0
:Read_P
%LIB_BAT%\LYRSupport.bat %*
exit /b 0
rem =================================================
