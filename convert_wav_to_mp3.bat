@echo off
setlocal enabledelayedexpansion

:: Use PowerShell to select the source folder
for /f "usebackq delims=" %%i in (`powershell -noprofile -command ^
    "Add-Type -AssemblyName System.windows.forms;" ^
    "[System.Windows.Forms.Application]::EnableVisualStyles();" ^
    "$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog;" ^
    "$folderBrowser.Description = 'Select the SOURCE directory';" ^
    "if ($folderBrowser.ShowDialog() -eq 'OK') { write-output $folderBrowser.SelectedPath }"`) do (
    set "source=%%i"
)

if not defined source (
    echo No source folder selected. Exiting.
    pause
    exit /b 1
)

:: Use PowerShell to select the destination folder
for /f "usebackq delims=" %%i in (`powershell -noprofile -command ^
    "Add-Type -AssemblyName System.windows.forms;" ^
    "[System.Windows.Forms.Application]::EnableVisualStyles();" ^
    "$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog;" ^
    "$folderBrowser.Description = 'Select the DESTINATION directory';" ^
    "if ($folderBrowser.ShowDialog() -eq 'OK') { write-output $folderBrowser.SelectedPath }"`) do (
    set "dest=%%i"
)

if not defined dest (
    echo No destination folder selected. Exiting.
    pause
    exit /b 1
)

echo Source folder: "%source%"
echo Destination folder: "%dest%"
echo.

:: Make sure destination folder exists
if not exist "%dest%" (
    mkdir "%dest%"
)

:: Process each .wav file in source folder
pushd "%source%"
for %%a in (*.wav) do (
    echo Processing "%%a" ...
    ffmpeg -i "%%a" -vn -c:a libmp3lame -ar 44100 "%dest%\%%~na.mp3"
)
popd

echo.
echo Conversion complete!
pause