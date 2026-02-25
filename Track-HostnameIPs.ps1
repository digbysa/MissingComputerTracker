[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputCsvPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputCsvPath,

    [int]$PingCount = 1,

    [int]$TimeoutSeconds = 2
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-IpAddressForHostname {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Hostname,
        [int]$Count = 1,
        [int]$Timeout = 2
    )

    try {
        $testResult = Test-Connection -ComputerName $Hostname -Count $Count -TimeoutSeconds $Timeout -ErrorAction Stop |
            Select-Object -First 1

        if ($null -ne $testResult -and $null -ne $testResult.IPV4Address) {
            return $testResult.IPV4Address.IPAddressToString
        }

        if ($null -ne $testResult -and $null -ne $testResult.Address) {
            return [string]$testResult.Address
        }
    }
    catch {
        # Fall through to DNS resolution attempt.
    }

    try {
        $dnsResult = Resolve-DnsName -Name $Hostname -Type A -ErrorAction Stop |
            Where-Object { $_.IPAddress } |
            Select-Object -First 1

        if ($null -ne $dnsResult) {
            return [string]$dnsResult.IPAddress
        }
    }
    catch {
        return $null
    }

    return $null
}

if (-not (Test-Path -LiteralPath $InputCsvPath)) {
    throw "Input CSV not found: $InputCsvPath"
}

$inputRows = Import-Csv -LiteralPath $InputCsvPath
if (-not $inputRows) {
    throw "Input CSV is empty: $InputCsvPath"
}

$outputByHostname = @{}
if (Test-Path -LiteralPath $OutputCsvPath) {
    $existingRows = Import-Csv -LiteralPath $OutputCsvPath
    foreach ($row in $existingRows) {
        if ($row.Hostname) {
            $outputByHostname[$row.Hostname.ToLowerInvariant()] = $row
        }
    }
}

$nowUtc = (Get-Date).ToUniversalTime().ToString('s') + 'Z'
$newOutputRows = foreach ($row in $inputRows) {
    if (-not $row.Hostname) {
        continue
    }

    $hostname = [string]$row.Hostname
    $hostnameKey = $hostname.ToLowerInvariant()

    $newIp = Get-IpAddressForHostname -Hostname $hostname -Count $PingCount -Timeout $TimeoutSeconds
    $status = if ($newIp) { 'Reachable' } else { 'Unreachable' }

    $previousRecord = $null
    if ($outputByHostname.ContainsKey($hostnameKey)) {
        $previousRecord = $outputByHostname[$hostnameKey]
    }

    $previousIp = if ($previousRecord) { [string]$previousRecord.CurrentIP } else { '' }
    $changeCount = if ($previousRecord -and $previousRecord.IPChangeCount) { [int]$previousRecord.IPChangeCount } else { 0 }
    $lastIpChangeUtc = if ($previousRecord) { [string]$previousRecord.LastIPChangeUtc } else { '' }

    if ($newIp -and $previousIp -and $newIp -ne $previousIp) {
        $changeCount++
        $lastIpChangeUtc = $nowUtc
    }

    [pscustomobject]@{
        Hostname        = $hostname
        CurrentIP       = if ($newIp) { $newIp } else { '' }
        PreviousIP      = $previousIp
        LastSeenUtc     = $nowUtc
        LastIpChangeUtc = $lastIpChangeUtc
        IPChangeCount   = $changeCount
        Status          = $status
    }
}

if (-not $newOutputRows) {
    throw 'No valid hostnames were found in the input CSV. Ensure there is a Hostname column.'
}

$newOutputRows |
    Sort-Object Hostname |
    Export-Csv -LiteralPath $OutputCsvPath -NoTypeInformation -Force

Write-Host "IP tracking completed. Output written to: $OutputCsvPath"
