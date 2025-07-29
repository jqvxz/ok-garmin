@echo off
setlocal enabledelayedexpansion
title ok-garmin - simple setup
echo.
echo [i] Please wait while we are installing the required programs to make the voice assistant work
echo.

echo [*] Checking for Python installation...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [!] Python is not installed
    echo [*] Starting installation via winget...
    winget install --id Python.Python.3 -e
) else (
    echo [+] Python is installed
)

echo.
echo [*] Checking for Microphones...
for /f "tokens=*" %%a in ('wmic path Win32_SoundDevice get Name /format:list ^| findstr /i "microphone mic"') do (
    set "mic_found=true"
)

if "%mic_found%"=="true" (
    echo [+] Microphones were found

) else (
    echo [!] No microphones found
    set /p no-mic=" [!] Continue anyway (y/N): "
)

echo.
echo [*] Installing requirements for python program...
set "packages[0]=SpeechRecognition"
set "packages[1]=Pillow"
set "packages[2]=fuzzywuzzy"
set "packages[3]=python-Levenshtein"
set "packages[4]=colorama"
set "all_success=pyautogui"
set "all_success=pyaudio"
set "all_success=true"
for /L %%i in (0,1,4) do (
    call :install_package "%%packages[%%i]%%"
    if not "%ERRORLEVEL%"=="0" (
        set "all_success=false"
    )
)
echo.
if "%all_success%"=="true" (
    echo [+] All pip installations finished successfully
    echo.
) else (
    echo [!] Some installations failed. Please check the log above for details.
)

echo [*] Installing the program on your pc...
mkdir ok-garmin
cd ok-garmin
curl -s https://raw.githubusercontent.com/jqvxz/ok-garmin/refs/heads/main/listener.py > listener.py
curl -s https://raw.githubusercontent.com/jqvxz/ok-garmin/refs/heads/main/blend.png > blend.png
echo.
echo [i] The installation of ok-garmin has been finished
echo [i] After hearing the beep 2 times, you can start the program by saying "ok garmin" to start the voice assistant
echo [i] Starting listener in 3 seconds 
timeout -t 3 > NUL
python listener.py

goto :eof

:install_package
set "package_name=%~1"
echo [*] Installing %package_name%:
pip install %package_name% >NUL 2>&1
if "%ERRORLEVEL%"=="0" (
    echo [+] %package_name% installed successfully.
) else (
    echo [!] Failed to install %package_name%.
)
exit /b %ERRORLEVEL%
