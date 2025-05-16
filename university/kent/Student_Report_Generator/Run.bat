@echo off
setlocal

:: Define source file and destination folder
set "SOURCE_FILE=reports_generator.py"
set "DEST_FOLDER=examples\COMP5590.A2"
set "DEST_FILE=%DEST_FOLDER%\%SOURCE_FILE%"

:: Copy file to destination
copy "%SOURCE_FILE%" "%DEST_FOLDER%" >nul

:: Change directory
cd %DEST_FOLDER%

:: Run the copied file and wait for it to finish
python %SOURCE_FILE%

:: Delete the copied file
del "%DEST_FILE%"

echo Done.
endlocal