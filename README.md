# MissingComputerTracker

Track hostname reachability and IP address changes with **PowerShell only**.

## What this does

`Track-HostnameIPs.ps1`:
- Reads hostnames from an input CSV.
- Pings each hostname (with DNS fallback).
- Writes a tracking CSV that includes:
  - current IP,
  - previous IP,
  - last seen timestamp,
  - last IP change timestamp,
  - IP change count,
  - reachable/unreachable status.

This output CSV is designed to be reused on each run so IP changes are tracked over time.

## Files

- `Track-HostnameIPs.ps1` - main script.
- `Run-TrackerLoop.ps1` - optional long-running scheduler loop (08:00 and 20:00).
- `devices.csv` (you create this) - input list of hostnames.
- `ip-tracking.csv` (script creates/updates this) - tracked output.

## Input CSV format

Create a CSV with a `Hostname` column.

Example `devices.csv`:

```csv
Hostname
PC-001
LAPTOP-22
SERVER-FILE01
```

## Run manually

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\Track-HostnameIPs.ps1 -InputCsvPath .\devices.csv -OutputCsvPath .\ip-tracking.csv
```

If using Windows PowerShell 5.1:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Track-HostnameIPs.ps1 -InputCsvPath .\devices.csv -OutputCsvPath .\ip-tracking.csv
```

## Schedule twice daily (Task Scheduler)

### Option A: GUI
1. Open **Task Scheduler**.
2. Create Task.
3. On **General**, choose a user account with network access.
4. On **Triggers**, create two daily triggers (example: `08:00` and `20:00`).
5. On **Actions**, set:
   - Program/script: `powershell.exe`
   - Arguments:
     ```text
     -NoProfile -ExecutionPolicy Bypass -File "C:\Path\To\Track-HostnameIPs.ps1" -InputCsvPath "C:\Path\To\devices.csv" -OutputCsvPath "C:\Path\To\ip-tracking.csv"
     ```
6. Save and test-run the task.

### Option B: PowerShell command

```powershell
$scriptPath = "C:\Path\To\Track-HostnameIPs.ps1"
$inputPath = "C:\Path\To\devices.csv"
$outputPath = "C:\Path\To\ip-tracking.csv"

$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -InputCsvPath `"$inputPath`" -OutputCsvPath `"$outputPath`""
$trigger1 = New-ScheduledTaskTrigger -Daily -At 8:00AM
$trigger2 = New-ScheduledTaskTrigger -Daily -At 8:00PM

Register-ScheduledTask -TaskName 'MissingComputerTracker' -Action $action -Trigger @($trigger1, $trigger2) -Description 'Ping hostnames and track IP changes twice daily.'
```

## Run twice daily without Task Scheduler

If you do not want to use Task Scheduler, use `Run-TrackerLoop.ps1`.

This script runs continuously and executes `Track-HostnameIPs.ps1` at 08:00 and 20:00 each day.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Run-TrackerLoop.ps1 `
  -ScriptPath .\Track-HostnameIPs.ps1 `
  -InputCsvPath .\devices.csv `
  -OutputCsvPath .\ip-tracking.csv
```

### Keep it running in the background

- **Locked session:** works fine if the process is already running.
- **If user logs off / machine reboots:** restart the loop via one of these:
  - **Startup folder** (`shell:startup`) shortcut for login-based startup.
  - **Windows service wrapper** (for example NSSM) if you need it to run without an interactive session.

Example NSSM setup:

```powershell
nssm install MissingComputerTrackerLoop "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" "-NoProfile -ExecutionPolicy Bypass -File C:\Path\Run-TrackerLoop.ps1 -ScriptPath C:\Path\Track-HostnameIPs.ps1 -InputCsvPath C:\Path\devices.csv -OutputCsvPath C:\Path\ip-tracking.csv"
nssm start MissingComputerTrackerLoop
```

## Notes

- Keep the same output CSV path on every run so change history is preserved.
- If a device is unreachable, `CurrentIP` is blank and `Status` is `Unreachable`.
- `PreviousIP` reflects the prior run's `CurrentIP` value for that hostname.
