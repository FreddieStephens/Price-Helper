@echo off
setlocal enabledelayedexpansion

:: Ensure script is running with admin rights
>nul 2>&1 net session
if %errorLevel% neq 0 (
    echo [INFO] Requesting administrator privileges...
    powershell -Command "Start-Process cmd -ArgumentList '/k \"\"%~f0\"\"' -Verb RunAs"
    exit /b
)

echo ===== FULL Resetting Price Helper at %date% %time% =====

:: -------------------------------
:: Paths & Variables
:: -------------------------------
set "venvdir=%LOCALAPPDATA%\price_helper_venv"
set "startup=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
set "shortcut=%startup%\PriceHelper.lnk"
set "cfshortcut=%startup%\FreddieStartCloudflared.lnk"
set "cfbat=%~dp0FreddieStartCloudflared.bat"
set "CFDIR=%USERPROFILE%\.cloudflared"
set "CONFIG=%CFDIR%\config.yml"
set "SYS_CFDIR=C:\Windows\System32\config\systemprofile\.cloudflared"
set "logfile=%LOCALAPPDATA%\price_helper_setup.log"
set "TMP_JSON=%TEMP%\tunnels.json"
set "TUNNELNAME=freddie-price"
set "PATHFILE=%LOCALAPPDATA%\price_helper_freddie_path.txt"

:: Cloudflare API
set "CF_API=https://api.cloudflare.com/client/v4"
set "ACCOUNT_ID=f368ee963961e43860bafc8c02801881"
set "CF_TOKEN=-DamjUHreGAJhkikDOLsxAMueWG4kHNOmDO1tSD7"

:: -------------------------------
:: Kill processes
:: -------------------------------
echo [INFO] Killing Python + Cloudflare processes...
taskkill /F /IM python.exe /T >nul 2>&1
taskkill /F /IM pythonw.exe /T >nul 2>&1
taskkill /F /IM cloudflared.exe /T >nul 2>&1

:: -------------------------------
:: Stop + delete Cloudflare service
:: -------------------------------
sc stop Cloudflared >nul 2>&1
sc delete Cloudflared >nul 2>&1
echo [INFO] Cloudflare service removed (if existed).

:: -------------------------------
:: Delete venv
:: -------------------------------
if exist "%venvdir%" (
    echo [INFO] Removing venv folder...
    rmdir /s /q "%venvdir%"
)

:: -------------------------------
:: Delete startup shortcuts + wrapper
:: -------------------------------
if exist "%shortcut%" del "%shortcut%"
if exist "%cfshortcut%" del "%cfshortcut%"
if exist "%cfbat%" del "%cfbat%"

:: -------------------------------
:: Delete tunnel configs (but keep cert.pem)
:: -------------------------------
if exist "%CONFIG%" del "%CONFIG%"
if exist "%SYS_CFDIR%\config.yml" del "%SYS_CFDIR%\config.yml"

echo [INFO] Cleaning tunnel JSON files but preserving cert.pem...
if exist "%CFDIR%" del "%CFDIR%\*.json" >nul 2>&1
if exist "%SYS_CFDIR%" del "%SYS_CFDIR%\*.json" >nul 2>&1

:: Preserve cert.pem explicitly
if exist "%CFDIR%\cert.pem" (
    echo [INFO] Preserving user cert.pem
)
if exist "%SYS_CFDIR%\cert.pem" (
    echo [INFO] Preserving system cert.pem
)

:: -------------------------------
:: Delete ALL tunnels via Cloudflare API
:: -------------------------------
echo [INFO] Fetching all tunnels "%TUNNELNAME%" from Cloudflare API...
curl -s -X GET "%CF_API%/accounts/%ACCOUNT_ID%/cfd_tunnel" ^
  -H "Authorization: Bearer %CF_TOKEN%" ^
  -H "Content-Type: application/json" > "%TMP_JSON%"

for /f "usebackq delims=" %%i in (`
    powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$json = Get-Content -Raw '%TMP_JSON%' | ConvertFrom-Json; foreach ($t in $json.result) { if ($t.name -eq '%TUNNELNAME%') { $t.id } }"
`) do (
    echo [INFO] Deleting tunnel via API with id %%i...
    curl -s -X DELETE "%CF_API%/accounts/%ACCOUNT_ID%/cfd_tunnel/%%i" ^
      -H "Authorization: Bearer %CF_TOKEN%" ^
      -H "Content-Type: application/json" >nul
)

:: -------------------------------
:: Delete setup log + path file
:: -------------------------------
if exist "%logfile%" del "%logfile%"
if exist "%PATHFILE%" (
    del "%PATHFILE%"
    echo [INFO] Deleted path file: %PATHFILE%
)

echo.
echo ✅ FULL Reset complete. Tunnel "%TUNNELNAME%" removed via API.
echo cert.pem preserved in user + systemprofile – no need to re-login to Cloudflare.
echo.
pause
endlocal
exit /b 0
