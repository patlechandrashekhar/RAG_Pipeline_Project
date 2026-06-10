@echo off
echo ============================================
echo KEEPING YOUR SCREEN AND SYSTEM AWAKE
echo ============================================
echo.
echo Your computer will stay awake while this window is open
echo Close this window to allow sleep again
echo.

REM Method 1: PowerShell keep-awake
powershell -ExecutionPolicy Bypass -Command "Write-Host 'Screen will stay on - Close window to stop' -ForegroundColor Green; while ($true) { [void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'); [System.Windows.Forms.SendKeys]::SendWait('{SCROLLLOCK}'); [System.Windows.Forms.SendKeys]::SendWait('{SCROLLLOCK}'); Write-Host '.' -NoNewline; Start-Sleep -Seconds 60 }"