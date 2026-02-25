[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ScriptPath,

    [Parameter(Mandatory = $true)]
    [string]$InputCsvPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputCsvPath,

    [TimeSpan[]]$RunTimes = @([TimeSpan]::FromHours(8), [TimeSpan]::FromHours(20)),

    [int]$PollSeconds = 30
)

if (-not (Test-Path -LiteralPath $ScriptPath)) {
    throw "ScriptPath not found: $ScriptPath"
}

if ($RunTimes.Count -eq 0) {
    throw 'RunTimes must include at least one time of day.'
}

$normalizedRunTimes = $RunTimes | Sort-Object -Unique
$lastRunSlotKey = $null

Write-Host "MissingComputerTracker loop started at $(Get-Date -Format o)"
Write-Host "Run slots: $($normalizedRunTimes -join ', ')"

while ($true) {
    $now = Get-Date

    $currentSlot = $normalizedRunTimes |
        Where-Object {
            $slotStart = [DateTime]::Today.Add($_)
            $slotEnd = $slotStart.AddSeconds($PollSeconds)
            $now -ge $slotStart -and $now -lt $slotEnd
        } |
        Select-Object -First 1

    if ($null -ne $currentSlot) {
        $slotKey = "{0:yyyy-MM-dd}-{1}" -f $now.Date, $currentSlot

        if ($slotKey -ne $lastRunSlotKey) {
            $lastRunSlotKey = $slotKey

            Write-Host "Running tracker for slot $currentSlot at $(Get-Date -Format o)..."
            & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $ScriptPath -InputCsvPath $InputCsvPath -OutputCsvPath $OutputCsvPath

            if ($LASTEXITCODE -ne 0) {
                Write-Warning "Tracker exited with code $LASTEXITCODE"
            }
        }
    }

    Start-Sleep -Seconds $PollSeconds
}
