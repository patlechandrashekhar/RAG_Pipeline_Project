# PowerShell script to keep Windows awake
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "KEEPING YOUR SCREEN AND SYSTEM AWAKE" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Your computer will stay awake while this window is open" -ForegroundColor Yellow
Write-Host "Press Ctrl+C or close window to allow sleep again" -ForegroundColor Yellow
Write-Host ""

# Method 1: Use Windows Presentation Settings API
Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    public class KeepAwake
    {
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern uint SetThreadExecutionState(uint esFlags);

        public const uint ES_CONTINUOUS = 0x80000000;
        public const uint ES_SYSTEM_REQUIRED = 0x00000001;
        public const uint ES_DISPLAY_REQUIRED = 0x00000002;
        public const uint ES_AWAYMODE_REQUIRED = 0x00000040;
    }
"@

# Prevent system and display from sleeping
$null = [KeepAwake]::SetThreadExecutionState(
    [KeepAwake]::ES_CONTINUOUS -bor
    [KeepAwake]::ES_SYSTEM_REQUIRED -bor
    [KeepAwake]::ES_DISPLAY_REQUIRED
)

Write-Host "[OK] System sleep disabled" -ForegroundColor Green
Write-Host "[OK] Display sleep disabled" -ForegroundColor Green

# Method 2: Also use SendKeys as backup to simulate activity
Add-Type -AssemblyName System.Windows.Forms

$startTime = Get-Date
$counter = 0

try {
    Write-Host ""
    Write-Host "Running... (activity indicators below)" -ForegroundColor Cyan

    while ($true) {
        # Calculate elapsed time
        $elapsed = (Get-Date) - $startTime
        $elapsedString = "{0:D2}:{1:D2}:{2:D2}" -f $elapsed.Hours, $elapsed.Minutes, $elapsed.Seconds

        # Simulate key press every 60 seconds (F15 key - won't interfere with work)
        [System.Windows.Forms.SendKeys]::SendWait('{F15}')

        # Update status
        $counter++
        Write-Host -NoNewline "`r[Active for $elapsedString] Heartbeat: $counter " -ForegroundColor Gray

        # Wait 60 seconds
        Start-Sleep -Seconds 60
    }
}
catch {
    Write-Host ""
    Write-Host "Keep-alive interrupted: $_" -ForegroundColor Yellow
}
finally {
    # Re-enable sleep when script ends
    $null = [KeepAwake]::SetThreadExecutionState([KeepAwake]::ES_CONTINUOUS)
    Write-Host ""
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host "System can now sleep normally" -ForegroundColor Yellow
    Write-Host "============================================" -ForegroundColor Cyan
}