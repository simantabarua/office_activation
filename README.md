```sh
@echo off
echo Starting Microsoft Office activation process...
(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")
(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")
echo Installing license files...
for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
echo Setting up KMS parameters...
cscript //nologo slmgr.vbs /ckms >nul
cscript //nologo ospp.vbs /setprt:1688 >nul
cscript //nologo ospp.vbs /unpkey:6F7TH >nul
set i=1
cscript //nologo ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH >nul || goto notsupported

:skms
if %i% GTR 10 goto busy
if %i% EQU 1 set KMS=kms7.office.com
if %i% EQU 2 set KMS=e8.us.to
if %i% EQU 3 set KMS=e9.us.to
if %i% GTR 3 goto ato
echo Attempting to set KMS host to: %KMS%
cscript //nologo ospp.vbs /sethst:%KMS% >nul

:ato
cscript //nologo ospp.vbs /act | find /i "successful" && (
    echo Activation successful!
        goto halt
) || (
    echo Connection to KMS server failed. Retrying...
    set /a i+=1
    goto skms
)

:notsupported
echo Sorry, your version is not supported.
goto halt

:busy
echo Server is busy. Please try again later.
goto halt

:halt
pause >nul
exit

```
