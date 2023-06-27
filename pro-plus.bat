@echo off
color 2
set srv=1
for /F "tokens=1,2 delims=#" %%a in ('"prompt #$H#$E# & echo on & for %%b in (1) do     rem"') do (
  set "DEL=%%a"
)
:check_Permissions
    net session >nul 2>&1
    if %errorLevel% == 0 (
        cls
    ) else (
        color 4
        echo Run this tool as administrator.
        pause
        exit
    )
:top
cls
if %PROCESSOR_ARCHITECTURE%==AMD64 cd /d %ProgramFiles%\Microsoft Office\Office16
if %PROCESSOR_ARCHITECTURE%==x86 cd /d %ProgramFiles(x86)%\Microsoft Office\Office16
cscript ospp.vbs /dstatus | find "LICENSED" >>nul
if %errorLevel% == 0 (
echo -------------------------------------
echo   MS OFFICE IS ALREADY ACTIVATED
echo -------------------------------------
pause
exit
)
echo ----------------------------------
echo      ACTIVATING MS OFFICE
echo ----------------------------------
echo.
timeout 3 >nul
for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >>nul
:loop
echo -----------------------------
echo CONNECTING TO SERVER %srv%
echo -----------------------------
echo.
if %srv% == 1 set host=s8.uk.to
if %srv% == 2 set host=kms8.MSGuides.com
if %srv% == 3 set host=kms.lotro.cc
if %srv% == 4 set host=kms9.MSGuides.com
if %srv% == 5 set host=kms.guowaifuli.com
cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH >nul
cscript ospp.vbs /unpkey:BTDRB >nul
cscript ospp.vbs /unpkey:KHGM9 >nul
cscript ospp.vbs /unpkey:CPQVG >nul
cscript ospp.vbs /sethst:%host% >nul
cscript ospp.vbs /setprt:1688 >nul
cscript ospp.vbs /act | find /I "activation successful" >>nul
if %errorLevel% == 0 (
cls
echo ----------------------------------
echo  PRODUCT ACTIVATED SUCCESSFULLY
echo ----------------------------------
pause
exit
)
call :colorEcho 04 "-------------------------- "
echo.
call :colorEcho 04 " CONNECTION FAILLED!"
echo.
call :colorEcho 04 "--------------------------"
echo.
set /a srv=%srv% + 1
if %srv% GTR 5 (
cls
call :colorEcho 04 "-------------------------- "
echo.
call :colorEcho 04 "   FAILED TO ACTIVATE!"
echo.
call :colorEcho 04 "--------------------------"
echo.
pause
exit
)
goto loop

:colorEcho
echo off
<nul set /p ".=%DEL%" > "%~2"
findstr /v /a:%1 /R "^$" "%~2" nul
del "%~2" > nul 2>&1i