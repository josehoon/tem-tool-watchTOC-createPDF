#Requires -RunAsAdministrator
<#
.SYNOPSIS
  Registers the tem-ord-cards monitor as a Windows Scheduled Task (every 15 min, 24x7).
  Run this script once as Administrator.
#>

$TASK_NAME   = 'tem-ord-cards PDF Monitor'
$SCRIPT_PATH = Join-Path $PSScriptRoot 'monitor.ps1'

$action = New-ScheduledTaskAction `
    -Execute 'powershell.exe' `
    -Argument "-NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$SCRIPT_PATH`""

# Run every 15 minutes indefinitely
$trigger = New-ScheduledTaskTrigger -RepetitionInterval (New-TimeSpan -Minutes 15) -Once `
    -At (Get-Date) -RepetitionDuration ([TimeSpan]::MaxValue)

$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Hours 3) `   # tests can run long
    -MultipleInstances IgnoreNew `                   # don't stack if previous run still going
    -RunOnlyIfNetworkAvailable `
    -StartWhenAvailable                              # catch up if machine was asleep

$principal = New-ScheduledTaskPrincipal `
    -UserId $env:USERNAME `
    -LogonType Interactive `                         # runs in user session (for VPN/creds)
    -RunLevel Highest

# Remove old task if it exists
Unregister-ScheduledTask -TaskName $TASK_NAME -Confirm:$false -ErrorAction SilentlyContinue

Register-ScheduledTask `
    -TaskName  $TASK_NAME `
    -Action    $action `
    -Trigger   $trigger `
    -Settings  $settings `
    -Principal $principal `
    -Description 'Polls GitHub for tem-ord-cards validate-pipeline completions and runs PDF tests against integration.' |
    Out-Null

Write-Host "Task '$TASK_NAME' registered successfully." -ForegroundColor Green
Write-Host "It will run every 15 minutes in your user session." -ForegroundColor Cyan
Write-Host ""
Write-Host "Useful commands:"
Write-Host "  Start-ScheduledTask -TaskName '$TASK_NAME'          # run now"
Write-Host "  Get-ScheduledTask   -TaskName '$TASK_NAME'          # check status"
Write-Host "  Unregister-ScheduledTask -TaskName '$TASK_NAME'     # remove"
