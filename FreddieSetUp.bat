@echo off
setlocal enabledelayedexpansion

:: -------------------------------
:: Setup logging
:: -------------------------------
set "logfile=%LOCALAPPDATA%\price_helper_setup.log"
echo ===== Setup started at %date% %time% ===== > "%logfile%"

set "scriptdir=%~dp0"
set "venvdir=%LOCALAPPDATA%\price_helper_venv"
set "startup=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
set "shortcut=%startup%\PriceHelper.lnk"

:: -------------------------------
:: USER VARIABLES
:: -------------------------------
set "TUNNELNAME=freddie-price"
set "HOSTNAME=freddie-price.fritzcloudflare.uk"
set "MACHINE=freddie"
set "CFDIR=%USERPROFILE%\.cloudflared"
set "CONFIG=%CFDIR%\config.yml"

:: 🔑 Cloudflare API
set "CF_API=https://api.cloudflare.com/client/v4"
set "CF_TOKEN=t2V0VSzkHSw0l0T3ZRjsCuBuO5NT2fadyZbllTBu"
set "CF_ZONEID=0070b13784f4fff064c52eac008672be"

:: -------------------------------
:: Detect Python
:: -------------------------------
where py >nul 2>&1 && (set "PYTHON=py -3") || (set "PYTHON=python")
echo [DEBUG] Using Python command: %PYTHON%

:: -------------------------------
:: Kill anything on port 54007
:: -------------------------------
for /f "tokens=5" %%a in ('netstat -ano ^| findstr :54007') do (
    echo [DEBUG] Killing process PID %%a on port 54007...
    taskkill /PID %%a /F >nul 2>&1
)
timeout /t 2 >nul

:: -------------------------------
:: Ensure venv exists
:: -------------------------------
if not exist "%venvdir%" (
    echo [DEBUG] Creating venv...
    %PYTHON% -m venv "%venvdir%"
)

:: -------------------------------
:: Upgrade pip + install deps
:: -------------------------------
"%venvdir%\Scripts\python.exe" -m pip install --upgrade pip >> "%logfile%" 2>&1
if exist "%scriptdir%requirements.txt" (
    "%venvdir%\Scripts\pip.exe" install -r "%scriptdir%requirements.txt" >> "%logfile%" 2>&1
) else (
    "%venvdir%\Scripts\pip.exe" install flask psutil pywin32 >> "%logfile%" 2>&1
)

:: -------------------------------
:: Launch Python helper
:: -------------------------------
echo [DEBUG] Launching Price Helper...
start "" "%venvdir%\Scripts\pythonw.exe" "%scriptdir%price_helper.pyw"

:: -------------------------------
:: Create Startup shortcut
:: -------------------------------
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
cscript //nologo "%temp%\makeshortcut.vbs" >nul
del "%temp%\makeshortcut.vbs"

:: -------------------------------
:: Detect or create tunnel UUID
:: -------------------------------
set "UUID="
set "CRED="

for /f "delims=" %%i in ('powershell -NoProfile -Command ^
  "$list = cloudflared tunnel list --output json | ConvertFrom-Json; ($list | Where-Object { $_.name -eq '%TUNNELNAME%' }).id"') do (
    set "UUID=%%i"
)

if not defined UUID (
    echo [INFO] No existing tunnel found, creating %TUNNELNAME%...
    cloudflared tunnel login
    cloudflared tunnel create %TUNNELNAME%
    for /f "delims=" %%i in ('powershell -NoProfile -Command ^
      "$list = cloudflared tunnel list --output json | ConvertFrom-Json; ($list | Where-Object { $_.name -eq '%TUNNELNAME%' }).id"') do (
        set "UUID=%%i"
    )
)

if not defined UUID (
    echo [ERROR] Could not determine tunnel UUID after creation. Aborting.
    pause
    exit /b 1
)

set "CRED=%CFDIR%\%UUID%.json"
echo [DEBUG] Using credentials: %CRED%
echo [DEBUG] Tunnel UUID = %UUID%

:: -------------------------------
:: Manage DNS record via Cloudflare API
:: -------------------------------
echo [INFO] Ensuring DNS record for %HOSTNAME% points to tunnel...

for /f "delims=" %%r in ('powershell -NoProfile -Command ^
  "(Invoke-RestMethod -Method GET -Uri '%CF_API%/zones/%CF_ZONEID%/dns_records?name=%HOSTNAME%' -Headers @{Authorization='Bearer %CF_TOKEN%';'Content-Type'='application/json'}).result[0].id"') do (
    if not "%%r"=="" (
        echo [INFO] Deleting old DNS record %%r...
        curl -s -X DELETE "%CF_API%/zones/%CF_ZONEID%/dns_records/%%r" -H "Authorization: Bearer %CF_TOKEN%" -H "Content-Type: application/json" >nul
    )
)

curl -s -X POST "%CF_API%/zones/%CF_ZONEID%/dns_records" ^
 -H "Authorization: Bearer %CF_TOKEN%" ^
 -H "Content-Type: application/json" ^
 --data "{\"type\":\"CNAME\",\"name\":\"%HOSTNAME%\",\"content\":\"%UUID%.cfargotunnel.com\",\"ttl\":120,\"proxied\":true}" >nul

:: -------------------------------
:: Write config.yml
:: -------------------------------
echo [INFO] Writing config.yml...
(
echo tunnel: %UUID%
echo credentials-file: %CRED%
echo.
echo ingress:
echo   - hostname: %HOSTNAME%
echo     service: http://localhost:54007
echo   - service: http_status:404
) > "%CONFIG%"

:: -------------------------------
:: Run the tunnel
:: -------------------------------
echo [DEBUG] Starting Cloudflare tunnel for %UUID%...
start "" cloudflared tunnel run %UUID%

:: -------------------------------
:: Set Salesforce cookie
:: -------------------------------
echo [DEBUG] Setting cookie sheethelper_machine=%MACHINE% for Salesforce domain...
start msedge --new-window --app="data:text/html,<script>document.cookie='sheethelper_machine=%MACHINE%; domain=.lightning.force.com; path=/; SameSite=Lax';window.close();</script>"

echo.
echo ✅ Setup complete! Log file: %logfile%
echo - Price Helper: http://127.0.0.1:54007/health
echo - Cloudflare tunnel: https://%HOSTNAME%
echo - Machine registered as: %MACHINE%
echo - Auto-start enabled.
echo.
pause
endlocal
exit /b 0
