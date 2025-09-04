@echo off
setlocal enabledelayedexpansion

:: Ensure script is running with admin rights
>nul 2>&1 net session
if %errorLevel% neq 0 (
    echo [INFO] Requesting administrator privileges...
    powershell -Command "Start-Process cmd -ArgumentList '/k \"\"%~f0\"\"' -Verb RunAs"
    exit /b
)

:: -------------------------------
:: User / API settings
:: -------------------------------
set "CF_API=https://api.cloudflare.com/client/v4"
set "ACCOUNT_ID=f368ee963961e43860bafc8c02801881"
set "CF_ZONEID=713e672a3709ea5c4ae81bd674e0fe76"
set "CF_TOKEN=-DamjUHreGAJhkikDOLsxAMueWG4kHNOmDO1tSD7"
set "TUNNELNAME=freddie-price"
set "HOSTNAME=freddie-price.fritzsynergy.uk"
set "MACHINE=freddie"

set "CFDIR=%USERPROFILE%\.cloudflared"
set "CONFIG=%CFDIR%\config.yml"
set "venvdir=%LOCALAPPDATA%\price_helper_venv"
set "startup=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
set "shortcut=%startup%\PriceHelper.lnk"
set "logfile=%LOCALAPPDATA%\price_helper_setup.log"
set "SYS_CFDIR=C:\Windows\System32\config\systemprofile\.cloudflared"
set "cfshortcut=%startup%\FreddieStartCloudflared.lnk"
set "cfbat=%~dp0FreddieStartCloudflared.bat"

echo ===== Setup started at %date% %time% ===== > "%logfile%"

:: -------------------------------
:: Detect Python
:: -------------------------------
where py >nul 2>&1 && (set "PYTHON=py -3") || (set "PYTHON=python")
echo [DEBUG] Using Python: %PYTHON%

:: -------------------------------
:: Ensure venv
:: -------------------------------
if not exist "%venvdir%" (
    echo [DEBUG] Creating venv...
    %PYTHON% -m venv "%venvdir%"
)

:: -------------------------------
:: Install deps
:: -------------------------------
"%venvdir%\Scripts\python.exe" -m pip install --upgrade pip >> "%logfile%" 2>&1
if exist "%~dp0requirements.txt" (
    "%venvdir%\Scripts\pip.exe" install -r "%~dp0requirements.txt" >> "%logfile%" 2>&1
) else (
    "%venvdir%\Scripts\pip.exe" install flask psutil pywin32 >> "%logfile%" 2>&1
)

:: -------------------------------
:: Startup shortcut for PriceHelper
:: -------------------------------
echo [DEBUG] Creating startup shortcut...
if exist "%shortcut%" del "%shortcut%"
(
echo Set oWS = CreateObject("WScript.Shell"^)
echo sLinkFile = "%shortcut%"
echo Set oLink = oWS.CreateShortcut(sLinkFile^)
echo oLink.TargetPath = "%venvdir%\Scripts\pythonw.exe"
echo oLink.Arguments = """" ^& "%~dp0price_helper.pyw" ^& """"
echo oLink.WorkingDirectory = "%~dp0"
echo oLink.Save
) > "%temp%\makeshortcut.vbs"
cscript //nologo "%temp%\makeshortcut.vbs" >nul
del "%temp%\makeshortcut.vbs"

:: -------------------------------
:: Cloudflare login (first-time only)
:: -------------------------------
if exist "%CFDIR%\cert.pem" (
    echo [INFO] cert.pem already exists, skipping login.
) else (
    echo [ACTION REQUIRED] Cloudflare authentication needed!
    pause
    cloudflared tunnel login
    if exist "%CFDIR%\cert.pem" (
        echo [INFO] Login successful – cert.pem created.
    ) else (
        echo [ERROR] Login failed or cert.pem missing. Aborting setup.
        pause
        exit /b 1
    )
)

:: -------------------------------
:: Tunnel check or create
:: -------------------------------
echo [INFO] Checking for existing tunnel "%TUNNELNAME%" via API...
curl -s -X GET "%CF_API%/accounts/%ACCOUNT_ID%/cfd_tunnel" ^
  -H "Authorization: Bearer %CF_TOKEN%" ^
  -H "Content-Type: application/json" > "%TEMP%\tunnels.json"

for /f "usebackq delims=" %%i in (`
    powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$j = Get-Content -Raw '%TEMP%\tunnels.json' | ConvertFrom-Json; $m = $j.result | Where-Object { $_.name -eq '%TUNNELNAME%' -and -not $_.deleted_at }; if ($m) { $m[0].id }"
`) do (
    set "TUNNEL_ID=%%i"
)

if not defined TUNNEL_ID (
    echo [INFO] Creating tunnel %TUNNELNAME%...
    for /f "tokens=6" %%i in ('cloudflared tunnel create %TUNNELNAME% ^| findstr /i "Created tunnel"') do (
        set "TUNNEL_ID=%%i"
    )
)

if not defined TUNNEL_ID (
    echo [ERROR] Could not determine tunnel UUID! Aborting.
    pause
    exit /b 1
)

echo [DEBUG] Tunnel UUID=!TUNNEL_ID!
set "CRED=%CFDIR%\!TUNNEL_ID!.json"

:: -------------------------------
:: Manage DNS record
:: -------------------------------
echo [INFO] Checking for existing DNS record "%HOSTNAME%"...
set "DNS_ID="
for /f "usebackq delims=" %%r in (`
  powershell -NoProfile -Command ^
  "(Invoke-RestMethod -Uri '%CF_API%/zones/%CF_ZONEID%/dns_records?name=%HOSTNAME%' -Headers @{Authorization='Bearer %CF_TOKEN%';'Content-Type'='application/json'}).result[0].id"
`) do (
    set "DNS_ID=%%r"
)

if defined DNS_ID (
    echo [INFO] Deleting old DNS record !DNS_ID!...
    curl -s -X DELETE "%CF_API%/zones/%CF_ZONEID%/dns_records/!DNS_ID!" ^
      -H "Authorization: Bearer %CF_TOKEN%" ^
      -H "Content-Type: application/json" >nul
)

echo [INFO] Creating fresh DNS record for %HOSTNAME%...
curl -s -X POST "%CF_API%/zones/%CF_ZONEID%/dns_records" ^
 -H "Authorization: Bearer %CF_TOKEN%" ^
 -H "Content-Type: application/json" ^
 --data "{\"type\":\"CNAME\",\"name\":\"%HOSTNAME%\",\"content\":\"!TUNNEL_ID!.cfargotunnel.com\",\"ttl\":120,\"proxied\":true}" >nul

:: -------------------------------
:: Write config.yml safely
:: -------------------------------
echo [INFO] Writing config.yml...
(
    echo tunnel: !TUNNEL_ID!
    echo credentials-file: %CRED%
    echo.
    echo ingress:
    echo   - hostname: %HOSTNAME%
    echo     service: http://localhost:54007
    echo   - service: http_status:404
) > "%CONFIG%"

:: -------------------------------
:: Copy config + creds for LocalSystem
:: -------------------------------
echo [INFO] Copying config + credentials to systemprofile...
if not exist "%SYS_CFDIR%" mkdir "%SYS_CFDIR%"
copy /Y "%CONFIG%" "%SYS_CFDIR%\config.yml" >nul
copy /Y "%CRED%" "%SYS_CFDIR%\" >nul
copy /Y "%CFDIR%\cert.pem" "%SYS_CFDIR%\" >nul

:: -------------------------------
:: Create Cloudflare startup wrapper (detached + retry loop)
:: -------------------------------
echo [INFO] Writing FreddieStartCloudflared.bat wrapper...
> "%cfbat%" (
    echo @echo off
    echo setlocal enabledelayedexpansion
    echo ^>nul 2^>^&1 net session
    echo if %%errorlevel%% neq 0 (
    echo     powershell -Command "Start-Process cmd -ArgumentList '/c \"\"%%~f0\"\"\"' -Verb RunAs"
    echo     exit /b
    echo ^)
    echo :mainloop
    echo echo [INFO] Waiting for Flask health endpoint...
    echo :waitloop
    echo curl -s http://127.0.0.1:54007/health ^| find "ok" ^>nul
    echo if errorlevel 1 (
    echo     timeout /t 5 ^>nul
    echo     goto waitloop
    echo ^)
    echo echo [INFO] Flask is ready – starting Cloudflared in background...
    echo powershell -WindowStyle Hidden -Command "Start-Process -WindowStyle Hidden -FilePath 'C:\Windows\System32\cloudflared.exe' -ArgumentList '--config','%SYS_CFDIR%\config.yml','tunnel','run','%TUNNELNAME%'"
    echo echo [INFO] Cloudflared launched, monitoring...
    echo :monitor
    echo timeout /t 20 ^>nul
    echo tasklist ^| find /i "cloudflared.exe" ^>nul
    echo if errorlevel 1 (
    echo     echo [WARN] Cloudflared not running – restarting...
    echo     goto mainloop
    echo ^)
    echo goto monitor
    echo endlocal
)

:: -------------------------------
:: Create startup shortcut for Cloudflared wrapper
:: -------------------------------
echo [INFO] Creating startup shortcut for Cloudflared wrapper...
if exist "%cfshortcut%" del "%cfshortcut%"
(
echo Set oWS = CreateObject("WScript.Shell"^)
echo sLinkFile = "%cfshortcut%"
echo Set oLink = oWS.CreateShortcut(sLinkFile^)
echo oLink.TargetPath = "%cfbat%"
echo oLink.WorkingDirectory = "%~dp0"
echo oLink.WindowStyle = 7
echo oLink.Save
) > "%temp%\makecfshortcut.vbs"
cscript //nologo "%temp%\makecfshortcut.vbs" >nul
del "%temp%\makecfshortcut.vbs"

:: -------------------------------
:: Start Price Helper immediately
:: -------------------------------
echo [INFO] Starting Price Helper app...
start "" "%venvdir%\Scripts\pythonw.exe" "%~dp0price_helper.pyw"

:: -------------------------------
:: Also trigger Cloudflare wrapper immediately (hidden)
:: -------------------------------
echo [INFO] Launching Cloudflare wrapper immediately (hidden)...
powershell -WindowStyle Hidden -Command "Start-Process -FilePath '%cfbat%' -WindowStyle Hidden"
:: -------------------------------
:: Finish
:: -------------------------------
echo.
echo ✅ Setup complete! Log: %logfile%
echo - Price Helper: http://127.0.0.1:54007/health
echo - Cloudflare tunnel: https://%HOSTNAME%
echo - Machine: %MACHINE%
echo.
pause
endlocal
exit /b 0
