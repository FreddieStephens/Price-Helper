@echo off
setlocal enabledelayedexpansion

:: -------------------------------
:: Setup logging (local only)
:: -------------------------------
set "logfile=%LOCALAPPDATA%\price_helper_setup.log"
echo ===== Setup started at %date% %time% ===== > "%logfile%"

:: Save script folder + venv path
set "scriptdir=%~dp0"
set "venvdir=%LOCALAPPDATA%\price_helper_venv"
set "startup=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
set "shortcut=%startup%\PriceHelper.lnk"

:: -------------------------------
:: Modes: reset / debug
:: -------------------------------
set "DEBUGMODE=0"
if /i "%1"=="reset" set "RESETMODE=1"
if /i "%1"=="debug" set "DEBUGMODE=1"
if /i "%1"=="resetdebug" (set "RESETMODE=1" & set "DEBUGMODE=1")

:: -------------------------------
:: Detect Python (prefer py launcher first)
:: -------------------------------
where py >nul 2>&1
if errorlevel 1 (
    where python >nul 2>&1
    if errorlevel 1 (
        echo [ERROR] Python not found! Please install 64-bit Python 3.10+ and check "Add to PATH".
        pause
        exit /b 1
    ) else (
        set "PYTHON=python"
    )
) else (
    set "PYTHON=py -3"
)

echo [DEBUG] Using Python command: %PYTHON%

:: -------------------------------
:: Reset mode
:: -------------------------------
if defined RESETMODE (
    echo [DEBUG] Reset mode enabled

    :: Kill any running python/pythonw first
    taskkill /IM python.exe /F >nul 2>&1
    taskkill /IM pythonw.exe /F >nul 2>&1

    if exist "%venvdir%" (
        echo [DEBUG] Deleting venv folder...
        rmdir /s /q "%venvdir%"
    )
    if exist "%shortcut%" (
        echo [DEBUG] Deleting startup shortcut...
        del "%shortcut%"
    )
)

:: -------------------------------
:: Kill any process on port 54007
:: -------------------------------
echo [DEBUG] Checking for processes on port 54007...
for /f "tokens=5" %%a in ('netstat -ano ^| findstr :54007') do (
    echo [DEBUG] Killing process PID %%a on port 54007...
    taskkill /PID %%a /F
)

:: -------------------------------
:: Ensure venv exists
:: -------------------------------
if not exist "%venvdir%" (
    echo [DEBUG] Creating venv...
    %PYTHON% -m venv "%venvdir%"
    if errorlevel 1 (
        echo [ERROR] Failed to create venv at %venvdir%
        pause
        exit /b 1
    )
)

:: -------------------------------
:: Upgrade pip
:: -------------------------------
echo [DEBUG] Upgrading pip...
if %DEBUGMODE%==1 (
    "%venvdir%\Scripts\python.exe" -m pip install --upgrade pip
) else (
    "%venvdir%\Scripts\python.exe" -m pip install --upgrade pip >> "%logfile%" 2>&1
)

if errorlevel 1 (
    echo [ERROR] pip upgrade failed. See %logfile%
    pause
    exit /b 1
)

:: -------------------------------
:: Install requirements
:: -------------------------------
if exist "%scriptdir%requirements.txt" (
    echo [DEBUG] Installing requirements.txt...
    if %DEBUGMODE%==1 (
        "%venvdir%\Scripts\pip.exe" install -r "%scriptdir%requirements.txt"
    ) else (
        "%venvdir%\Scripts\pip.exe" install -r "%scriptdir%requirements.txt" >> "%logfile%" 2>&1
    )
) else (
    echo [DEBUG] No requirements.txt found, installing core deps...
    if %DEBUGMODE%==1 (
        "%venvdir%\Scripts\pip.exe" install flask psutil pywin32
    ) else (
        "%venvdir%\Scripts\pip.exe" install flask psutil pywin32 >> "%logfile%" 2>&1
    )
)

if errorlevel 1 (
    echo [ERROR] Dependency install failed. See %logfile%
    pause
    exit /b 1
)

:: -------------------------------
:: Launch helper
:: -------------------------------
echo [DEBUG] Launching Price Helper...
start "" "%venvdir%\Scripts\pythonw.exe" "%scriptdir%price_helper.pyw"

:: -------------------------------
:: Create Startup shortcut
:: -------------------------------
echo [DEBUG] Creating startup shortcut...
if exist "%shortcut%" del "%shortcut%"

(
echo Set oWS = CreateObject("WScript.Shell"^)
echo sLinkFile = "%shortcut%"
echo Set oLink = oWS.CreateShortcut(sLinkFile^)
echo oLink.TargetPath = "%venvdir%\Scripts\pythonw.exe"
echo oLink.Arguments = """" ^& "%scriptdir%price_helper.pyw" ^& """"
echo oLink.WorkingDirectory = "%scriptdir%"
echo oLink.Save
) > "%temp%\makeshortcut.vbs"

cscript //nologo "%temp%\makeshortcut.vbs"
del "%temp%\makeshortcut.vbs"

:: -------------------------------
:: Health check (external ps1)
:: -------------------------------
echo [DEBUG] Waiting 5 seconds for health check...
timeout /t 5 >nul

echo [DEBUG] Running healthcheck.ps1...
powershell -NoProfile -ExecutionPolicy Bypass -File "%scriptdir%healthcheck.ps1" -LogFile "%logfile%"

echo.
echo [DEBUG] Health check finished. Press Enter to continue...
pause >nul

echo.
echo ✅ Setup complete! Log file (with details) is at %logfile%
echo - Price Helper runs at: http://127.0.0.1:54007/health
echo - It will auto-start each time you log in.
echo.
pause
endlocal
exit /b 0
