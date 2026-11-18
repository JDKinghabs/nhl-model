# setup_daily_task.ps1  -  sets up (or replaces) the Windows Task Scheduler task that runs daily_update.py every morning
# =====================================================================================================================
# WHAT IT DOES:
#   Creates a task called "NHL Model daily update" in Task Scheduler that starts scripts\daily_update.py every day,
#   without a window. If the computer is asleep at that time, Windows wakes it to run the task (when wake timers are
#   allowed in the power settings). If the computer was off, the task runs soon after you next log on.
#   The task only runs while you are logged on (a locked screen is fine), so no password is stored.
#
# HOW TO RUN IT (in PowerShell, from the "Model 2.0" folder):
#   powershell -ExecutionPolicy Bypass -File scripts\setup_daily_task.ps1              every day at 6:00 AM
#   powershell -ExecutionPolicy Bypass -File scripts\setup_daily_task.ps1 -Time 7:30AM a different time
#   powershell -ExecutionPolicy Bypass -File scripts\setup_daily_task.ps1 -Remove      delete the task
#   ("-ExecutionPolicy Bypass" only lets this one script run; it does not change any setting.)
#
# To pause it without deleting it: Task Scheduler > Task Scheduler Library > "NHL Model daily update" > Disable.

param([string]$Time = "6:00AM", [switch]$Remove)

$TaskName = "NHL Model daily update"
if ($Remove) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Output "Removed the task '$TaskName'."
    return
}

$project = Split-Path $PSScriptRoot -Parent
$python = (Get-Command python -ErrorAction Stop).Source
$pythonw = Join-Path (Split-Path $python) "pythonw.exe"          # the window-less Python, so nothing pops up at 6 AM
if (-not (Test-Path $pythonw)) { $pythonw = $python }
$script = Join-Path $project "scripts\daily_update.py"

$action = New-ScheduledTaskAction -Execute $pythonw -Argument "`"$script`"" -WorkingDirectory $project
$trigger = New-ScheduledTaskTrigger -Daily -At $Time
$settings = New-ScheduledTaskSettingsSet -WakeToRun -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Hours 2) -MultipleInstances IgnoreNew
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Limited

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Force `
    -Description "Refreshes the NHL Model website (standings, games, contender fit, breakouts) and publishes it. Logs: data\daily\" | Out-Null

$task = Get-ScheduledTask -TaskName $TaskName
$info = Get-ScheduledTaskInfo -TaskName $TaskName
Write-Output "Task '$TaskName' is set up: every day at $Time (state: $($task.State); next run: $($info.NextRunTime))."
Write-Output "It runs: $pythonw `"$script`""
