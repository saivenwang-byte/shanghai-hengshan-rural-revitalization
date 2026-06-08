@echo off
echo ====================================
echo  Codex Proxy Switcher Launcher
echo  Proxy: http://127.0.0.1:51008
echo ====================================
echo.
echo  [1] Start Codex with Proxy (VPN Mode)
echo  [2] Start Codex without Proxy (Native Mode)
echo  [3] Open Codex-Proxy-Switcher GUI
echo  [4] Exit
echo.
choice /C 1234 /N /M "Select option (1-4): "
if errorlevel 4 exit /b
if errorlevel 3 start "" "C:\Users\Lenovo\.codex\skills\codex-proxy-switcher-win\Codex-Proxy-Switcher.exe" & exit /b
if errorlevel 2 start "" "C:\Users\Lenovo\.codex\skills\codex-proxy-switcher-win\Start-Codex-Native.bat" & exit /b
if errorlevel 1 start "" "C:\Users\Lenovo\.codex\skills\codex-proxy-switcher-win\Start-Codex-VPN.bat" & exit /b
