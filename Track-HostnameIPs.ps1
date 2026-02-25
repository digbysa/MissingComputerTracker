[CmdletBinding()]
param(
    [string]$SourceInputPath = 'C:\Users\da1701_sa\Desktop\New-Inventory-Tool\Output\MissingDeviceList.csv',
    [string]$LocalInputFileName = 'MissingDeviceList.csv',
    [string]$OutputFolderName = 'Output',
    [string]$OutputFileName = 'SearchedDeviceList.csv',
    [string]$SubnetFileName = 'SiteSubnets.csv',
    [int]$PingCount = 1,
    [int]$TimeoutSeconds = 2
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Convert-IPv4ToUInt32 {
    param([Parameter(Mandatory = $true)][string]$IpAddress)

    $ipObj = [System.Net.IPAddress]::Parse($IpAddress)
    $bytes = $ipObj.GetAddressBytes()

    if ($bytes.Length -ne 4) {
        throw "Only IPv4 addresses are supported: $IpAddress"
    }

    [Array]::Reverse($bytes)
    return [BitConverter]::ToUInt32($bytes, 0)
}

function Test-IpInCidr {
    param(
        [Parameter(Mandatory = $true)][string]$IpAddress,
        [Parameter(Mandatory = $true)][string]$Cidr
    )

    if ($Cidr -notmatch '^([^/]+)/([0-9]|[12][0-9]|3[0-2])$') {
        return $false
    }

    $networkIp = $Matches[1]
    $prefixLength = [int]$Matches[2]

    try {
        $ipInt = Convert-IPv4ToUInt32 -IpAddress $IpAddress
        $networkInt = Convert-IPv4ToUInt32 -IpAddress $networkIp
    }
    catch {
        return $false
    }

    $mask = if ($prefixLength -eq 0) {
        [uint32]0
    }
    else {
        [uint32](0xFFFFFFFF -shl (32 - $prefixLength))
    }

    return (($ipInt -band $mask) -eq ($networkInt -band $mask))
}

function Get-IpAddressForHostname {
    param(
        [Parameter(Mandatory = $true)][string]$Hostname,
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
        # Continue to DNS fallback.
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

function Get-LoggedOnUser {
    param([Parameter(Mandatory = $true)][string]$Hostname)

    try {
        $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $Hostname -ErrorAction Stop
        if ($computerSystem.UserName) {
            return [string]$computerSystem.UserName
        }
    }
    catch {
        return ''
    }

    return ''
}

function Get-SubnetLabel {
    param(
        [Parameter(Mandatory = $true)][string]$IpAddress,
        [Parameter(Mandatory = $true)][System.Collections.IEnumerable]$SubnetRows
    )

    foreach ($subnetRow in $SubnetRows) {
        if ($null -eq $subnetRow.Cidr -or [string]::IsNullOrWhiteSpace($subnetRow.Cidr)) {
            continue
        }

        if (Test-IpInCidr -IpAddress $IpAddress -Cidr $subnetRow.Cidr) {
            return [string]$subnetRow.Label
        }
    }

    return ''
}

function Get-HistoryValue {
    param(
        [Parameter(Mandatory = $true)]$Row,
        [Parameter(Mandatory = $true)][string]$PropertyName
    )

    $property = $Row.PSObject.Properties[$PropertyName]
    if ($null -eq $property) {
        return ''
    }

    return [string]$property.Value
}

$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$localInputPath = Join-Path -Path $scriptRoot -ChildPath $LocalInputFileName
$outputFolderPath = Join-Path -Path $scriptRoot -ChildPath $OutputFolderName
$outputCsvPath = Join-Path -Path $outputFolderPath -ChildPath $OutputFileName
$subnetPath = Join-Path -Path $scriptRoot -ChildPath $SubnetFileName

if (Test-Path -LiteralPath $SourceInputPath) {
    Copy-Item -LiteralPath $SourceInputPath -Destination $localInputPath -Force
    Write-Host "Copied input from source path: $SourceInputPath"
}

if (-not (Test-Path -LiteralPath $localInputPath)) {
    throw "Input CSV not found at source or local path: $SourceInputPath / $localInputPath"
}

if (-not (Test-Path -LiteralPath $outputFolderPath)) {
    New-Item -Path $outputFolderPath -ItemType Directory -Force | Out-Null
}

$inputRows = Import-Csv -LiteralPath $localInputPath
if (-not $inputRows) {
    throw "Input CSV is empty: $localInputPath"
}

$requiredColumns = @('Name', 'Asset Tag', 'Location')
foreach ($requiredColumn in $requiredColumns) {
    if ($null -eq $inputRows[0].PSObject.Properties[$requiredColumn]) {
        throw "Input CSV is missing required column '$requiredColumn': $localInputPath"
    }
}

$subnetRows = @()
if (Test-Path -LiteralPath $subnetPath) {
    $rawSubnetRows = Import-Csv -LiteralPath $subnetPath

    foreach ($rawSubnetRow in $rawSubnetRows) {
        $properties = $rawSubnetRow.PSObject.Properties
        if ($properties.Count -lt 2) {
            continue
        }

        $subnetRows += [pscustomobject]@{
            Cidr  = [string]$properties[0].Value
            Label = [string]$properties[1].Value
        }
    }
}

$existingRowsByAssetTag = @{}
$maxPreviousDataIndex = 0

if (Test-Path -LiteralPath $outputCsvPath) {
    $existingRows = Import-Csv -LiteralPath $outputCsvPath

    foreach ($existingRow in $existingRows) {
        $assetTag = [string](Get-HistoryValue -Row $existingRow -PropertyName 'Asset Tag')
        if (-not [string]::IsNullOrWhiteSpace($assetTag)) {
            $existingRowsByAssetTag[$assetTag] = $existingRow
        }

        foreach ($prop in $existingRow.PSObject.Properties) {
            if ($prop.Name -match '^Previous Data ([0-9]+) --> IP Date$') {
                $index = [int]$Matches[1]
                if ($index -gt $maxPreviousDataIndex) {
                    $maxPreviousDataIndex = $index
                }
            }
        }
    }
}

$inputRowsByAssetTag = @{}
foreach ($inputRow in $inputRows) {
    $assetTag = [string]$inputRow.'Asset Tag'
    if ([string]::IsNullOrWhiteSpace($assetTag)) {
        continue
    }

    if ($inputRowsByAssetTag.ContainsKey($assetTag)) {
        $existingTimestamp = $null
        $incomingTimestamp = $null

        if ($inputRowsByAssetTag[$assetTag].Timestamp) {
            [DateTime]::TryParse([string]$inputRowsByAssetTag[$assetTag].Timestamp, [ref]$existingTimestamp) | Out-Null
        }

        if ($inputRow.Timestamp) {
            [DateTime]::TryParse([string]$inputRow.Timestamp, [ref]$incomingTimestamp) | Out-Null
        }

        if ($null -eq $existingTimestamp -or ($null -ne $incomingTimestamp -and $incomingTimestamp -gt $existingTimestamp)) {
            $inputRowsByAssetTag[$assetTag] = $inputRow
        }
    }
    else {
        $inputRowsByAssetTag[$assetTag] = $inputRow
    }
}

$nowUtc = (Get-Date).ToUniversalTime().ToString('s') + 'Z'
$resultRows = @()

foreach ($assetTag in ($inputRowsByAssetTag.Keys | Sort-Object)) {
    $inputRow = $inputRowsByAssetTag[$assetTag]
    $hostname = [string]$inputRow.Name
    $location = [string]$inputRow.Location

    $existingRow = $null
    if ($existingRowsByAssetTag.ContainsKey($assetTag)) {
        $existingRow = $existingRowsByAssetTag[$assetTag]
    }

    $ipAddress = ''
    $pingSuccessful = $false

    if (-not [string]::IsNullOrWhiteSpace($hostname)) {
        try {
            $pingSuccessful = [bool](Test-Connection -ComputerName $hostname -Count $PingCount -Quiet -TimeoutSeconds $TimeoutSeconds -ErrorAction Stop)
        }
        catch {
            $pingSuccessful = $false
        }

        $resolvedIp = Get-IpAddressForHostname -Hostname $hostname -Count $PingCount -Timeout $TimeoutSeconds
        if ($resolvedIp) {
            $ipAddress = [string]$resolvedIp
        }
    }

    $loggedUser = if ($pingSuccessful -and -not [string]::IsNullOrWhiteSpace($hostname)) {
        Get-LoggedOnUser -Hostname $hostname
    }
    else {
        ''
    }

    $subnetLabel = if ($ipAddress -and $subnetRows.Count -gt 0) {
        Get-SubnetLabel -IpAddress $ipAddress -SubnetRows $subnetRows
    }
    else {
        ''
    }

    $currentHistory = @()
    if ($existingRow) {
        $currentHistory += [pscustomobject]@{
            IpDate    = Get-HistoryValue -Row $existingRow -PropertyName 'Latest Data --> IP Date'
            IpAddress = Get-HistoryValue -Row $existingRow -PropertyName 'Latest Data --> IP Address'
            Subnet    = Get-HistoryValue -Row $existingRow -PropertyName 'Latest Data --> Subnet'
            LoggedUser = Get-HistoryValue -Row $existingRow -PropertyName 'Latest Data --> Logged User'
        }

        for ($i = 1; $i -le $maxPreviousDataIndex; $i++) {
            $currentHistory += [pscustomobject]@{
                IpDate    = Get-HistoryValue -Row $existingRow -PropertyName "Previous Data $i --> IP Date"
                IpAddress = Get-HistoryValue -Row $existingRow -PropertyName "Previous Data $i --> IP Address"
                Subnet    = Get-HistoryValue -Row $existingRow -PropertyName "Previous Data $i --> Subnet"
                LoggedUser = Get-HistoryValue -Row $existingRow -PropertyName "Previous Data $i --> Logged User"
            }
        }
    }

    $latestIpFromExisting = if ($existingRow) { Get-HistoryValue -Row $existingRow -PropertyName 'Latest Data --> IP Address' } else { '' }

    if ($ipAddress -and $latestIpFromExisting -and $ipAddress -ne $latestIpFromExisting) {
        $currentHistory = @(
            [pscustomobject]@{
                IpDate    = $nowUtc
                IpAddress = $ipAddress
                Subnet    = $subnetLabel
                LoggedUser = $loggedUser
            }
        ) + $currentHistory

        if ($currentHistory.Count -gt $maxPreviousDataIndex + 2) {
            $currentHistory = $currentHistory[0..($maxPreviousDataIndex + 1)]
        }

        $maxPreviousDataIndex = [Math]::Max($maxPreviousDataIndex, $currentHistory.Count - 1)
    }
    elseif (-not $existingRow) {
        $currentHistory = @(
            [pscustomobject]@{
                IpDate    = if ($ipAddress) { $nowUtc } else { '' }
                IpAddress = $ipAddress
                Subnet    = $subnetLabel
                LoggedUser = $loggedUser
            }
        )
    }
    elseif ($ipAddress -and $latestIpFromExisting -eq $ipAddress) {
        if ($currentHistory.Count -eq 0) {
            $currentHistory = @([pscustomobject]@{ IpDate = $nowUtc; IpAddress = $ipAddress; Subnet = $subnetLabel; LoggedUser = $loggedUser })
        }
        else {
            $currentHistory[0].IpDate = $nowUtc
            $currentHistory[0].IpAddress = $ipAddress
            $currentHistory[0].Subnet = $subnetLabel
            $currentHistory[0].LoggedUser = $loggedUser
        }
    }

    $successValue = if ($pingSuccessful -or ($existingRow -and (Get-HistoryValue -Row $existingRow -PropertyName 'Successfully Pinged') -eq 'Yes')) {
        'Yes'
    }
    else {
        'No'
    }

    $row = [ordered]@{
        'Name'                = $hostname
        'Asset Tag'           = $assetTag
        'Location'            = $location
        'Successfully Pinged' = $successValue
        'Latest Data --> IP Date'     = ''
        'Latest Data --> IP Address'  = ''
        'Latest Data --> Subnet'      = ''
        'Latest Data --> Logged User' = ''
    }

    if ($currentHistory.Count -gt 0) {
        $row['Latest Data --> IP Date'] = [string]$currentHistory[0].IpDate
        $row['Latest Data --> IP Address'] = [string]$currentHistory[0].IpAddress
        $row['Latest Data --> Subnet'] = [string]$currentHistory[0].Subnet
        $row['Latest Data --> Logged User'] = [string]$currentHistory[0].LoggedUser
    }

    $historyIndex = 1
    while ($historyIndex -lt $currentHistory.Count) {
        $row["Previous Data $historyIndex --> IP Date"] = [string]$currentHistory[$historyIndex].IpDate
        $row["Previous Data $historyIndex --> IP Address"] = [string]$currentHistory[$historyIndex].IpAddress
        $row["Previous Data $historyIndex --> Subnet"] = [string]$currentHistory[$historyIndex].Subnet
        $row["Previous Data $historyIndex --> Logged User"] = [string]$currentHistory[$historyIndex].LoggedUser
        $historyIndex++
    }

    $resultRows += [pscustomobject]$row
}

if (-not $resultRows) {
    throw 'No valid rows were found with Asset Tag values in MissingDeviceList.csv.'
}

$globalMaxPrevious = 0
foreach ($row in $resultRows) {
    foreach ($prop in $row.PSObject.Properties) {
        if ($prop.Name -match '^Previous Data ([0-9]+) --> IP Date$') {
            $candidate = [int]$Matches[1]
            if ($candidate -gt $globalMaxPrevious) {
                $globalMaxPrevious = $candidate
            }
        }
    }
}

$finalRows = foreach ($row in ($resultRows | Sort-Object 'Asset Tag')) {
    $ordered = [ordered]@{
        'Name'                = [string]$row.Name
        'Asset Tag'           = [string]$row.'Asset Tag'
        'Location'            = [string]$row.Location
        'Successfully Pinged' = [string]$row.'Successfully Pinged'
        'Latest Data --> IP Date'     = [string]$row.'Latest Data --> IP Date'
        'Latest Data --> IP Address'  = [string]$row.'Latest Data --> IP Address'
        'Latest Data --> Subnet'      = [string]$row.'Latest Data --> Subnet'
        'Latest Data --> Logged User' = [string]$row.'Latest Data --> Logged User'
    }

    for ($i = 1; $i -le $globalMaxPrevious; $i++) {
        $ordered["Previous Data $i --> IP Date"] = [string](Get-HistoryValue -Row $row -PropertyName "Previous Data $i --> IP Date")
        $ordered["Previous Data $i --> IP Address"] = [string](Get-HistoryValue -Row $row -PropertyName "Previous Data $i --> IP Address")
        $ordered["Previous Data $i --> Subnet"] = [string](Get-HistoryValue -Row $row -PropertyName "Previous Data $i --> Subnet")
        $ordered["Previous Data $i --> Logged User"] = [string](Get-HistoryValue -Row $row -PropertyName "Previous Data $i --> Logged User")
    }

    [pscustomobject]$ordered
}

$finalRows | Export-Csv -LiteralPath $outputCsvPath -NoTypeInformation -Force

if (Test-Path -LiteralPath $localInputPath) {
    Remove-Item -LiteralPath $localInputPath -Force
}

Write-Host "Tracker run complete. Output written to: $outputCsvPath"
