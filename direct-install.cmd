@echo off
setlocal enabledelayedexpansion

:: Color codes
for /F "tokens=1" %%e in ('echo prompt $E ^| cmd') do set "ESC=%%e"
set "RED=!ESC![31m"
set "GREEN=!ESC![32m"
set "YELLOW=!ESC![33m"
set "BLUE=!ESC![34m"
set "MAGENTA=!ESC![35m"
set "CYAN=!ESC![36m"
set "WHITE=!ESC![37m"
set "RESET=!ESC![0m"

:: Title and welcome message
title ok-garmin - simple setup
echo.
echo    _____ _                 _       _____      _               
echo   / ____(_)               ^| ^|     / ____^|    ^| ^|              
echo  ^| (___  _ _ __ ___  _ __ ^| ^| ___^| (___   ___^| ^|_ _   _ _ __  
echo   \___ \^| ^| '_ ` _ \^| '_ \^| ^|/ _ \\___ \ / _ \ __^| ^| ^| ^| '_ \ 
echo   ____) ^| ^| ^| ^| ^| ^| ^| ^|_) ^| ^|  __/____) ^|  __/ ^|_^| ^|_^| ^| ^|_) ^|
echo  ^|_____/^|_^|_^| ^|_^| ^|_^| .__/^|_^|\___^|_____/ \___^|\__^|\__,_^| .__/   for Python voice assistant
echo                     ^| ^|                                ^| ^|    
echo                     ^|_^|                                ^|_^|    
echo.
echo %WHITE%[i] Please wait while we are installing the required programs to make the voice assistant work%RESET%
echo.

:: Check for python installation
echo %WHITE%[*] Checking for Python installation...%RESET%
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo %YELLOW%[^^!] Python is not installed%RESET%
    echo %WHITE%[*] Starting installation via winget...%RESET%
    winget install -e --id Python.Python.3.11
    if %errorlevel% neq 0 (
        echo %RED%[^^!] Failed to install Python. Please install it manually from https://www.python.org/downloads/ %RESET%
        timeout -t 2 > NUL
        exit /b 1
    ) else (
        echo %GREEN%[+] Python installed successfully%RESET%
    )
) else (
    echo %GREEN%[+] Python is installed%RESET%
)

:: Checking for microphones
echo.
echo %WHITE%[*] Checking for Microphones...%RESET%
for /f "tokens=*" %%a in ('wmic path Win32_SoundDevice get Name /format:list ^| findstr /i "microphone mic"') do (
    set "mic_found=true"
)

if "%mic_found%"=="true" (
    echo %GREEN%[+] Microphones were found%RESET%

) else (
    echo %RED%[^^!] No microphones found%RESET%
    set /p no-mic="%RED%[^^!] Continue anyway (y/N): %RESET%"
)

if /i "%no-mic%" == "y" (
    echo %RED%[^^!] Exiting installation due to no microphones found%RESET%
    exit /b 1
)

:: Installing requirements
echo.
echo %WHITE%[*] Installing requirements for python program...%RESET%
set "packages[0]=SpeechRecognition"
set "packages[1]=Pillow"
set "packages[2]=fuzzywuzzy"
set "packages[3]=python-Levenshtein"
set "packages[4]=colorama"
set "packages[5]=pyautogui"
set "packages[6]=pyaudio"
set "all_success=true"

for /L %%i in (0,1,6) do (
    call :install_package "%%packages[%%i]%%"
    if not "!ERRORLEVEL!"=="0" (
        set "all_success=false"
    )
)
echo.
if "%all_success%"=="true" (
    echo %GREEN%[+] All pip installations finished successfully%RESET%
    echo.
) else (
    echo %RED%[^^!] Some installations failed. Please check the log above for details.%RESET%
)

:: Download voice assistant files
echo %WHITE%[*] Installing the program on your pc...%RESET%
if not exist "ok-garmin" mkdir ok-garmin
cd ok-garmin
curl -s https://raw.githubusercontent.com/jqvxz/ok-garmin/refs/heads/main/listener.py > listener.py
curl -s https://raw.githubusercontent.com/jqvxz/ok-garmin/refs/heads/main/blend.png > blend.png
if exist listener.py (
    if exist blend.png (
        echo %GREEN%[*] ok-garmin files downloaded successfully%RESET%
    ) else (
        echo %RED%[^^!] Failed to download ok-garmin files%RESET%
        timeout -t 2 > NUL
        exit /b 1
    )
) else (
    echo %RED%[^^!] Failed to download ok-garmin files%RESET%
    timeout -t 2 > NUL
    exit /b 1
)
echo.
echo %WHITE%[i] After hearing the beep 2 times, you can start the program by saying "ok garmin"%RESET%
echo %WHITE%[i] Starting listener in 3 seconds%RESET%
timeout -t 3 > NUL
cls
python listener.py

goto :eof

:install_package
set "package_name=%~1"
echo %WHITE%[*] Installing %package_name%:%RESET%
pip install %package_name% >NUL 2>&1
if "%ERRORLEVEL%"=="0" (
    echo %GREEN%[+] %package_name% installed successfully.%RESET%
) else (
    echo %RED%[^^!] Failed to install %package_name%.%RESET%
)
exit /b %ERRORLEVEL%