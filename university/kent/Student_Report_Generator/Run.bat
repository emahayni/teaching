@echo off
setlocal

:: Define source file and destination folder
set "SOURCE_FILE_1=marking_automation.py"
set "SOURCE_FILE_2=marking_utils.py"
set "DEST_FOLDER=examples\COMP5590.A3"
set "DEST_FILE=%DEST_FOLDER%\%SOURCE_FILE%"

:: Copy file to destination
copy "%SOURCE_FILE_1%" "%DEST_FOLDER%" >nul
copy "%SOURCE_FILE_2%" "%DEST_FOLDER%" >nul

:: Change directory
cd %DEST_FOLDER%

:: Run the copied file and wait for it to finish
python %SOURCE_FILE_1%

:: Delete the copied file
del "%SOURCE_FILE_1%"
del "%SOURCE_FILE_2%"

echo Done.
endlocal

pause