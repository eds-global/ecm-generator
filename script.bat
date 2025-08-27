@ECHO off
setlocal enabledelayedexpansion

:: Arguments:
:: %1 = full path to INP file (without extension)
:: %2 = weather file path
set inp_file=%1
set weather=%2
set doe_cmd=C:\doe22\doe22.bat exe48z

echo Running file: %inp_file%.inp
call %doe_cmd% %inp_file% %weather%

endlocal
