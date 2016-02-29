@echo off
:START
cls
set InstallDIR=C:\Refresh
echo.
echo //////////////////////////////////
echo //                              //
echo //      REFRESH SCRIPT by JC    //
echo //                              //
echo //////////////////////////////////
echo.
echo Enter Account (1. ACCEPT 2. TRAIN)
set /p Choice=Enter Choice
if %Choice%==1 (
set Account=ACCEPT
) else (
if %Choice%==2 (
set Account=TRAIN
) else (
echo Wrong Entry
pause
GOTO START
))
echo %Account%
set /p Location=Enter location (E:\JACDatabasesACCEPT)
IF %Location%=="" (
echo Invalid entry
pause
GOTO START )
set Drive=%Location:~0,2%
echo.
set /p Reference=Enter JAC Reference(Optional)
echo.
set Month=%date:~3,2%
set Day=%date:~0,2%
set Year=%date:~8,2%
set RefreshDIR=Refresh-%Day%.%Month%.%Year%
md %Location%\%RefreshDIR%
IF EXIST %Reference% (
type NUL > %Reference%.txt
)
cscript "%InstallDIR%\replace.vbs" "%InstallDIR%\Backup.txt" "%Location%\%RefreshDIR%"
cscript "%InstallDIR%\replace.vbs" "%InstallDIR%\RestoreSettings.txt" "%Location%\%RefreshDIR%"
c:\intersystems\%Account%\bin\cterm.exe /CONSOLE=cn_ap:%Account%[PHM] %InstallDIR%\Backup.txt
echo.
echo DONE
echo.
echo Now restore account
c:\intersystems\%Account%\Bin\CControl.exe CTERMINAL %Account% ^^BACKUP %%SYS
pause
echo.
echo Restoring Settings....
c:\intersystems\%Account%\bin\cterm.exe /CONSOLE=cn_ap:%Account%[PHM] %InstallDIR%\RestoreSettings.txt
echo.
echo Stopping %Account% Account
c:\intersystems\%Account%\Bin\CSS stop %Account%
echo Starting %Account% Account
c:\intersystems\%Account%\Bin\CSS start %Account%
IF EXIST %Reference% (ren %Location%\%RefreshDIR%\Refresh.cbk %Location%\%RefreshDIR%\"Referesh %Reference%".cbk )
echo DONE
pause
