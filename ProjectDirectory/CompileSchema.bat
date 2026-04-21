@echo off
setlocal

set PROGRAM_DIR=%SVR_Header_2D%\sdk\flatbuffers-25.12.19\bin
set TARGET_DIR=.\config\schema
set CPPOUT_DIR=.\config\reader
set BINOUT_DIR=./

REM 生成C++文件 fbs映射文件
for /f %%i in ('dir /b "%TARGET_DIR%\*.fbs"') do (
	echo 编译文件 "%%~nxi"
	%PROGRAM_DIR%\flatc.exe --binary --schema -o "%BINOUT_DIR%" "%TARGET_DIR%\%%~nxi"
	%PROGRAM_DIR%\flatc.exe --cpp -o "%CPPOUT_DIR%" "%TARGET_DIR%\%%~nxi"
)

pause