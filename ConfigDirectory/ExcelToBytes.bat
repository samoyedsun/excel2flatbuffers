@echo off
setlocal

set "WORK_DIR=%~dp0"
set "EXCEL_DIR=%WORK_DIR%"
set "TARGET_DIR=%WORK_DIR%\Temp"
set "PROGRAM_DIR=.\Tools"

if not exist "%TARGET_DIR%" md "%TARGET_DIR%"

if not "%1"=="" (
	:: 单个文件导出（有参数时）
    echo 导出文件 %~nx1 为 %~n1.bytes 通过 %~n1.bfbs
    %PROGRAM_DIR%\Tools.exe "%PROGRAM_DIR%\metadata.json" "%PROGRAM_DIR%\%~n1.bfbs" "%EXCEL_DIR%\%~nx1" "%TARGET_DIR%\%~n1.bytes"
) else (
	:: 批量导出模式（无参数时）
	for /f %%i in ('dir /b "%EXCEL_DIR%\*.xlsx"') do (
		echo 导出文件 %%~nxi 为 %%~ni.bytes 通过 %%~ni.bfbs
		%PROGRAM_DIR%\Tools.exe "%PROGRAM_DIR%\metadata.json" "%PROGRAM_DIR%\%%~ni.bfbs" "%EXCEL_DIR%\%%~nxi" "%TARGET_DIR%\%%~ni.bytes"
	)
)

pause