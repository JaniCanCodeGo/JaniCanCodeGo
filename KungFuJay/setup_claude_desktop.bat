@echo off
setlocal

set "CONFIG_DIR=%APPDATA%\Claude"
set "CONFIG_FILE=%CONFIG_DIR%\claude_desktop_config.json"

echo.
echo Clear Enough To Lead — Claude Desktop Setup
echo ============================================
echo.

if not exist "%CONFIG_DIR%" mkdir "%CONFIG_DIR%"

set /p "API_KEY=Enter your Anthropic API key: "
set /p "SCRIPT_PATH=Enter full path to cetl_mcp_server.py (e.g. C:\Users\You\Documents\KungFuJay\cetl_mcp_server.py): "

powershell -Command ^
  "$key = '%API_KEY%';" ^
  "$path = '%SCRIPT_PATH%' -replace '\\', '/';" ^
  "$json = '{""mcpServers"":{""clear-enough-to-lead"":{""command"":""python3"",""args"":[""' + $path + '""],""env"":{""ANTHROPIC_API_KEY"":""' + $key + '""}}}}';" ^
  "Set-Content -Path '%CONFIG_FILE%' -Value $json -Encoding UTF8"

if exist "%CONFIG_FILE%" (
    echo.
    echo SUCCESS: Config written to:
    echo   %CONFIG_FILE%
    echo.
    echo NEXT: Quit and restart Claude Desktop
    echo   1. Right-click Claude in the system tray ^> Quit
    echo   2. Reopen Claude Desktop
    echo   3. Look for the hammer icon in the chat box
    echo   4. Ask: "Write me an email to my team about..."
) else (
    echo ERROR: Could not write config. Try running as Administrator.
)

echo.
pause
