param(
    $App="indexed",
    $Name="ファイル監視", 
    $Root="c:\opt\mtb",
    [switch]$Clean, 
    [switch]$Pass
)

# configure
$ExecutePath = "${root}\bin\${app}.exe"
$ExecuteArgs = "-m1" 
$WorkingDirectory = "%USERPROFILE%\documents"
$Interval = 1

if (-Not (Test-Path $ExecutePath)) { $ExecutePath = "${root}\${app}.exe" }
if (-Not (Test-Path $ExecutePath)) { $ExecutePath = "..\bin\${app}.exe" }
if (-Not (Test-Path $ExecutePath)) { $ExecutePath = ".\bin\${app}.exe" }
if (-Not (Test-Path $ExecutePath)) { $ExecutePath = ".\${app}.exe" }
if (-Not (Test-Path $ExecutePath)) {
    Write-Host "Not found ${App}.exe" -ForegroundColor Yellow
    if (-not $Pass) { $host.UI.RawUI.ReadKey() | Out-Null }
    Exit
}
$ExecutePath = Resolve-Path $ExecutePath

# remove resistered task
if ($Clean) {
    if (Get-ScheduledTask | Where-Object {$_.TaskName -eq $Name}) {
        try {
            Write-Host "Stop ScheduleTask ${Name}." -ForegroundColor Yellow
            Stop-ScheduledTask  -TaskName $Name

            Write-Host "Unregister ScheduleTask ${Name}." -ForegroundColor Yellow
            Unregister-ScheduledTask -TaskName $Name -Confirm:$false
        }
        catch {}
    }
    if (-not $Pass) { $host.UI.RawUI.ReadKey() | Out-Null }
    Exit
}

# registration task schedule
$Trigger = @()
$Trigger += New-ScheduledTaskTrigger `
    -Once -at (Get-Date) `
    -RepetitionInterval (New-TimeSpan -Minutes $Interval)

$Action = New-ScheduledTaskAction `
    -Execute $ExecutePath `
    -Argument $ExecuteArgs `
    -WorkingDirectory $WorkingDirectory

$Settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Write-Host "Register ScheduleTask ${Name}." -ForegroundColor Yellow
Register-ScheduledTask -Force `
    -TaskName $Name `
    -Action $Action `
    -Settings $Settings `
    -Trigger $Trigger | Out-Null

# enable task scedule
Write-Host "Enable ScheduleTask ${Name}." -ForegroundColor Yellow
Get-ScheduledTask -TaskName $Name | Enable-ScheduledTask | Out-Null

# start task scedule
Write-Host "Start ScheduleTask ${Name}." -ForegroundColor Yellow
Get-ScheduledTask -TaskName $Name | Start-ScheduledTask

# display next scedule information
Write-Host "Display ScheduleTask ${Name}." -ForegroundColor Yellow
Get-ScheduledTaskInfo -TaskName $Name

# end of resister task
Write-Host "Registered ScheduleTask completed." -ForegroundColor Yellow
if (-not $Pass) { $host.UI.RawUI.ReadKey() | Out-Null }
