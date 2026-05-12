@echo off
setlocal

set "CONFIG_DIR=%APPDATA%\Claude"
set "CONFIG_FILE=%CONFIG_DIR%\claude_desktop_config.json"
set "SCRIPT_PATH=%USERPROFILE%\Documents\JaniCanCodeGo\crr_mcp_server.py"

echo Creating Claude Desktop config...
echo.

:: Create Claude config directory if it doesn't exist
if not exist "%CONFIG_DIR%" mkdir "%CONFIG_DIR%"

:: Write the JSON config using PowerShell (handles special characters safely)
powershell -Command ^
  "$p = $env:USERPROFILE + '\Documents\JaniCanCodeGo\crr_mcp_server.py';" ^
  "$p = $p -replace '\\', '/';" ^
  "$json = '{""mcpServers"":{""crr-agent"":{""command"":""python3"",""args"":[""%1""]}}}' -replace '%1', $p;" ^
  "Set-Content -Path $env:APPDATA\Claude\claude_desktop_config.json -Value $json -Encoding UTF8"

if exist "%CONFIG_FILE%" (
    echo SUCCESS: Config file created at:
    echo   %CONFIG_FILE%
    echo.
    echo Script path set to:
    echo   %SCRIPT_PATH%
    echo.
    echo -----------------------------------------------
    echo NEXT STEP: Quit and restart Claude Desktop
    echo   1. Right-click Claude in the system tray
    echo   2. Click Quit
    echo   3. Reopen Claude Desktop
    echo   4. Look for the hammer icon in the chat box
    echo -----------------------------------------------
) else (
    echo ERROR: Config file was not created.
    echo Please run this script as Administrator or create the file manually.
)

echo.
pause
