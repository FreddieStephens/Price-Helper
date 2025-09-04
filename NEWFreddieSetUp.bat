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
set "logfile=%LOCALAPPDATA%\price_helper_setup.log"
set "SYS_CFDIR=C:\Windows\System32\config\systemprofile\.cloudflared"
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
:: Ensure PsExec (for SYSTEM debugging)
:: -------------------------------
set "PSEXE=C:\Windows\System32\psexec.exe"
set "PSTOOLS=%USERPROFILE%\AppData\Local\Temp\PSTools.zip"
set "PSTMP=%USERPROFILE%\AppData\Local\Temp\PSTools"

if exist "%PSEXE%" (
    echo [INFO] PsExec already installed at "%PSEXE%"
) else (
    echo [INFO] PsExec not found, downloading...
    if exist "%PSTOOLS%" del /f /q "%PSTOOLS%" >nul 2>&1
    if exist "%PSTMP%" rd /s /q "%PSTMP%" >nul 2>&1

    powershell -Command "Invoke-WebRequest -Uri 'https://download.sysinternals.com/files/PSTools.zip' -OutFile '%PSTOOLS%'" || (
        echo [ERROR] Failed to download PsExec.
        pause
        exit /b 1
    )

    echo [INFO] Extracting PsExec...
    echo [DEBUG] ZIP path: %PSTOOLS%
    echo [DEBUG] Extracting to: %PSTMP%
    powershell -Command "Expand-Archive -Path '%PSTOOLS%' -DestinationPath '%PSTMP%' -Force"

    if exist "%PSTMP%\PsExec.exe" (
        copy /Y "%PSTMP%\PsExec.exe" "%PSEXE%" >nul
        echo [INFO] PsExec installed to "%PSEXE%"
    ) else (
        echo [ERROR] PsExec.exe not found after extraction!
        pause
        exit /b 1
    )
)

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
    for /f "tokens=6" %%i in ('cloudflared tunnel create "%TUNNELNAME%" ^| findstr /i "Created tunnel"') do (
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
:: Create Cloudflare wrapper (immediate start + monitor + restart)
:: -------------------------------
echo [INFO] Writing FreddieStartCloudflared.bat wrapper...
> "%cfbat%" (
    echo @echo off
    echo setlocal enabledelayedexpansion
    echo set "LOGFILE=%%LOCALAPPDATA%%\freddie_cloudflare_health.log"
    echo set "FAILCOUNT=0"
    echo set "MAXFAILS=3"
    echo.
    echo echo [%%date%% %%time%%] [INFO] Starting Cloudflared immediately... ^>^> "%%LOGFILE%%"
    echo powershell -WindowStyle Hidden -Command "Start-Process -WindowStyle Hidden -FilePath 'C:\Windows\System32\cloudflared.exe' -ArgumentList '--config','%SYS_CFDIR%\config.yml','tunnel','run','%TUNNELNAME%'"
    echo.
    echo :monitor
    echo timeout /t 60 ^>nul
    echo curl -s http://127.0.0.1:54007/health ^| find "ok" ^>nul
    echo if errorlevel 1 ^(
    echo   echo [%%date%% %%time%%] [WARN] Local Flask not healthy, attempt !FAILCOUNT! ^>^> "%%LOGFILE%%"
    echo   set /a FAILCOUNT+=1
    echo ^) else ^(
    echo   curl -s https://%HOSTNAME%/health ^| find "ok" ^>nul
    echo   if errorlevel 1 ^(
    echo     echo [%%date%% %%time%%] [WARN] Tunnel health check failed, attempt !FAILCOUNT! ^>^> "%%LOGFILE%%"
    echo     set /a FAILCOUNT+=1
    echo   ^) else ^(
    echo     echo [%%date%% %%time%%] [INFO] Tunnel healthy ^>^> "%%LOGFILE%%"
    echo     set FAILCOUNT=0
    echo   ^)
    echo ^)
    echo.
    echo if !FAILCOUNT! geq !MAXFAILS! ^(
    echo   echo [%%date%% %%time%%] [ERROR] Too many failures – restarting Cloudflared ^>^> "%%LOGFILE%%"
    echo   taskkill /F /IM cloudflared.exe /T ^>nul 2^>^&1
    echo   set FAILCOUNT=0
    echo   powershell -WindowStyle Hidden -Command "Start-Process -WindowStyle Hidden -FilePath 'C:\Windows\System32\cloudflared.exe' -ArgumentList '--config','%SYS_CFDIR%\config.yml','tunnel','run','%TUNNELNAME%'"
    echo ^)
    echo goto monitor
    echo endlocal
)


:: -------------------------------
:: Register both tasks in Task Scheduler
:: -------------------------------
echo [INFO] Registering Cloudflare + PriceHelper tasks in Task Scheduler...

schtasks /delete /tn "FreddieCloudflareTunnel" /f >nul 2>&1
schtasks /delete /tn "FreddiePriceHelper" /f >nul 2>&1

set "SCRIPT_DIR=%~dp0"
set "PRICEHELPER=%SCRIPT_DIR%price_helper.pyw"
set "PYTHONW=%venvdir%\Scripts\pythonw.exe"

:: Resolve to short 8.3 paths to avoid issues with spaces
for %%A in ("%cfbat%") do set "CFBAT_SHORT=%%~sA"
for %%A in ("%PYTHONW%") do set "PYTHONW_SHORT=%%~sA"
for %%A in ("%PRICEHELPER%") do set "PH_SHORT=%%~sA"

:: Debug output
echo [DEBUG] schtasks /create /tn "FreddieCloudflareTunnel" /tr "%CFBAT_SHORT%" /sc onlogon /rl highest /f
echo [DEBUG] schtasks /create /tn "FreddiePriceHelper" /tr "\"%PYTHONW_SHORT%\" \"%PH_SHORT%\"" /sc onlogon /rl highest /f

:: Register tasks with short paths
schtasks /create /tn "FreddieCloudflareTunnel" ^
  /tr "%CFBAT_SHORT%" ^
  /sc onlogon /rl highest /f

schtasks /create /tn "FreddiePriceHelper" ^
  /tr "\"%PYTHONW_SHORT%\" \"%PH_SHORT%\"" ^
  /sc onlogon /rl highest /f

:: -------------------------------
:: Start them immediately
:: -------------------------------
for %%A in ("%cfbat%") do set "CFBAT_SHORT=%%~sA"
for %%A in ("%PYTHONW%") do set "PYTHONW_SHORT=%%~sA"
for %%A in ("%PRICEHELPER%") do set "PH_SHORT=%%~sA"

echo [INFO] Launching PriceHelper + Cloudflare immediately...
echo [DEBUG] Starting: "%PYTHONW_SHORT%" "%PH_SHORT%"
start "" "%PYTHONW_SHORT%" "%PH_SHORT%"
echo [DEBUG] Starting: "%CFBAT_SHORT%"
start "" "%CFBAT_SHORT%"

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
