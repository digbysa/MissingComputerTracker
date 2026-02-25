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
        $ipBytes = [System.Net.IPAddress]::Parse($IpAddress).GetAddressBytes()
        $networkBytes = [System.Net.IPAddress]::Parse($networkIp).GetAddressBytes()
    }
    catch {
        return $false
    }

    if ($ipBytes.Length -ne 4 -or $networkBytes.Length -ne 4) {
        return $false
    }

    $fullBytesToCompare = [int][Math]::Floor($prefixLength / 8)
    $remainingBits = $prefixLength % 8

    for ($i = 0; $i -lt $fullBytesToCompare; $i++) {
        if ($ipBytes[$i] -ne $networkBytes[$i]) {
            return $false
        }
    }

    if ($remainingBits -gt 0) {
        $partialMask = [byte](((0xFF00 -shr $remainingBits) -band 0xFF))
        if (($ipBytes[$fullBytesToCompare] -band $partialMask) -ne ($networkBytes[$fullBytesToCompare] -band $partialMask)) {
            return $false
        }
    }

    return $true
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
            return [pscustomobject]@{
                PingSucceeded = $true
                IpAddress     = $testResult.IPV4Address.IPAddressToString
            }
        }

        if ($null -ne $testResult -and $null -ne $testResult.Address) {
            return [pscustomobject]@{
                PingSucceeded = $true
                IpAddress     = [string]$testResult.Address
            }
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
            return [pscustomobject]@{
                PingSucceeded = $false
                IpAddress     = [string]$dnsResult.IPAddress
            }
        }
    }
    catch {
        return [pscustomobject]@{
            PingSucceeded = $false
            IpAddress     = ''
        }
    }

    return [pscustomobject]@{
        PingSucceeded = $false
        IpAddress     = ''
    }
}

function Get-LoggedOnUser {
    param([Parameter(Mandatory = $true)][string]$Hostname)

    try {
        $profiles = Get-WmiObject -Class Win32_UserProfile -ComputerName $Hostname -ErrorAction Stop |
            Where-Object {
                -not $_.Special -and
                -not [string]::IsNullOrWhiteSpace($_.LocalPath) -and
                $_.LocalPath -match '\\Users\\' -and
                $_.LocalPath -notmatch '\\Users\\(Default|Default User|Public|All Users)$'
            } |
            Sort-Object LastUseTime -Descending

        $latestProfile = $profiles | Select-Object -First 1
        if ($null -ne $latestProfile -and -not [string]::IsNullOrWhiteSpace($latestProfile.LocalPath)) {
            return [string](Split-Path -Leaf $latestProfile.LocalPath)
        }
    }
    catch {
        # Fall back to Win32_ComputerSystem lookup.
    }

    try {
        $computerSystem = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $Hostname -ErrorAction Stop
        if (-not [string]::IsNullOrWhiteSpace($computerSystem.UserName)) {
            return [string]$computerSystem.UserName
        }
    }
    catch {
        return ''
    }

    return ''
}

function Set-SuccessPingHighlight {
    param(
        [Parameter(Mandatory = $true)][string]$CsvPath,
        [Parameter(Mandatory = $true)][string]$XlsxPath
    )

    $excel = $null
    $workbook = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Open($CsvPath)
        $worksheet = $workbook.Worksheets.Item(1)

        $headerMap = @{}
        $usedRange = $worksheet.UsedRange
        for ($column = 1; $column -le $usedRange.Columns.Count; $column++) {
            $headerMap[[string]$worksheet.Cells.Item(1, $column).Text] = $column
        }

        if (-not $headerMap.ContainsKey('Successfully Pinged')) {
            return
        }

        $successColumn = [int]$headerMap['Successfully Pinged']
        $lastRow = [int]$usedRange.Rows.Count

        for ($rowIndex = 2; $rowIndex -le $lastRow; $rowIndex++) {
            $cell = $worksheet.Cells.Item($rowIndex, $successColumn)
            if ([string]$cell.Text -eq 'Yes') {
                $cell.Interior.Color = 13434828 # Light green
            }
        }

        $workbook.SaveAs($XlsxPath, 51)
    }
    catch {
        Write-Warning "Unable to create formatted workbook '$XlsxPath'. CSV output was still generated. Error: $($_.Exception.Message)"
    }
    finally {
        if ($workbook) {
            $workbook.Close($false)
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
        }

        if ($excel) {
            $excel.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
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

function Get-DeviceKey {
    param(
        [Parameter(Mandatory = $true)]$Row
    )

    $name = [string](Get-HistoryValue -Row $Row -PropertyName 'Name')
    if (-not [string]::IsNullOrWhiteSpace($name)) {
        return "NAME::$($name.Trim().ToUpperInvariant())"
    }

    $assetTag = [string](Get-HistoryValue -Row $Row -PropertyName 'Asset Tag')
    if (-not [string]::IsNullOrWhiteSpace($assetTag)) {
        return "ASSET::$($assetTag.Trim().ToUpperInvariant())"
    }

    return ''
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

$inputRows = @(Import-Csv -LiteralPath $localInputPath)
if (-not $inputRows) {
    throw "Input CSV is empty: $localInputPath"
}

$requiredColumns = @('Name', 'Asset Tag', 'Location', 'Room')
foreach ($requiredColumn in $requiredColumns) {
    if ($null -eq $inputRows[0].PSObject.Properties[$requiredColumn]) {
        throw "Input CSV is missing required column '$requiredColumn': $localInputPath"
    }
}

$subnetRows = @()
if (Test-Path -LiteralPath $subnetPath) {
    $rawSubnetRows = @(Import-Csv -LiteralPath $subnetPath -Header 'Cidr', 'Label', 'Notes')

    foreach ($rawSubnetRow in $rawSubnetRows) {
        $cidr = [string]$rawSubnetRow.Cidr
        $label = [string]$rawSubnetRow.Label

        if ([string]::IsNullOrWhiteSpace($cidr)) {
            continue
        }

        # Allow optional header rows in SiteSubnets.csv.
        if ($cidr -eq 'Cidr') {
            continue
        }

        $subnetRows += [pscustomobject]@{
            Cidr  = $cidr.Trim()
            Label = $label.Trim()
        }
    }
}

$existingRowsByDeviceKey = @{}
$maxPreviousDataIndex = 0

if (Test-Path -LiteralPath $outputCsvPath) {
    $existingRows = Import-Csv -LiteralPath $outputCsvPath

    foreach ($existingRow in $existingRows) {
        $deviceKey = Get-DeviceKey -Row $existingRow
        if (-not [string]::IsNullOrWhiteSpace($deviceKey)) {
            $existingRowsByDeviceKey[$deviceKey] = $existingRow
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

$inputRowsByDeviceKey = @{}
foreach ($inputRow in $inputRows) {
    $deviceKey = Get-DeviceKey -Row $inputRow
    if ([string]::IsNullOrWhiteSpace($deviceKey)) {
        continue
    }

    if ($inputRowsByDeviceKey.ContainsKey($deviceKey)) {
        $existingTimestamp = $null
        $incomingTimestamp = $null

        if ($inputRowsByDeviceKey[$deviceKey].Timestamp) {
            [DateTime]::TryParse([string]$inputRowsByDeviceKey[$deviceKey].Timestamp, [ref]$existingTimestamp) | Out-Null
        }

        if ($inputRow.Timestamp) {
            [DateTime]::TryParse([string]$inputRow.Timestamp, [ref]$incomingTimestamp) | Out-Null
        }

        if ($null -eq $existingTimestamp -or ($null -ne $incomingTimestamp -and $incomingTimestamp -gt $existingTimestamp)) {
            $inputRowsByDeviceKey[$deviceKey] = $inputRow
        }
    }
    else {
        $inputRowsByDeviceKey[$deviceKey] = $inputRow
    }
}

$nowLocalFormatted = Get-Date -Format 'MM-dd-yyyy - HH:mm'
$resultRows = @()

$allDeviceKeys = @($existingRowsByDeviceKey.Keys + $inputRowsByDeviceKey.Keys | Sort-Object -Unique)

foreach ($deviceKey in $allDeviceKeys) {
    $inputRow = if ($inputRowsByDeviceKey.ContainsKey($deviceKey)) { $inputRowsByDeviceKey[$deviceKey] } else { $null }
    $existingRow = if ($existingRowsByDeviceKey.ContainsKey($deviceKey)) { $existingRowsByDeviceKey[$deviceKey] } else { $null }

    $hostname = if ($inputRow) { [string]$inputRow.Name } else { [string](Get-HistoryValue -Row $existingRow -PropertyName 'Name') }
    $assetTag = if ($inputRow) { [string]$inputRow.'Asset Tag' } else { [string](Get-HistoryValue -Row $existingRow -PropertyName 'Asset Tag') }
    $location = if ($inputRow) { [string]$inputRow.Location } else { [string](Get-HistoryValue -Row $existingRow -PropertyName 'Location') }
    $room = if ($inputRow) { [string]$inputRow.Room } else { [string](Get-HistoryValue -Row $existingRow -PropertyName 'Room') }

    if ([string]::IsNullOrWhiteSpace($hostname) -and $existingRow) {
        $hostname = [string](Get-HistoryValue -Row $existingRow -PropertyName 'Name')
    }

    if ([string]::IsNullOrWhiteSpace($assetTag) -and $existingRow) {
        $assetTag = [string](Get-HistoryValue -Row $existingRow -PropertyName 'Asset Tag')
    }

    if ([string]::IsNullOrWhiteSpace($location) -and $existingRow) {
        $location = [string](Get-HistoryValue -Row $existingRow -PropertyName 'Location')
    }

    if ([string]::IsNullOrWhiteSpace($room) -and $existingRow) {
        $room = [string](Get-HistoryValue -Row $existingRow -PropertyName 'Room')
    }

    $ipAddress = ''
    $pingSuccessful = $false

    if (-not [string]::IsNullOrWhiteSpace($hostname)) {
        $connectionInfo = Get-IpAddressForHostname -Hostname $hostname -Count $PingCount -Timeout $TimeoutSeconds
        if ($connectionInfo) {
            $ipAddress = [string]$connectionInfo.IpAddress
            $pingSuccessful = [bool]$connectionInfo.PingSucceeded -and -not [string]::IsNullOrWhiteSpace($ipAddress)
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
                IpDate    = $nowLocalFormatted
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
                IpDate    = if ($ipAddress) { $nowLocalFormatted } else { '' }
                IpAddress = $ipAddress
                Subnet    = $subnetLabel
                LoggedUser = $loggedUser
            }
        )
    }
    elseif ($ipAddress -and $latestIpFromExisting -eq $ipAddress) {
        if ($currentHistory.Count -eq 0) {
            $currentHistory = @([pscustomobject]@{ IpDate = $nowLocalFormatted; IpAddress = $ipAddress; Subnet = $subnetLabel; LoggedUser = $loggedUser })
        }
        else {
            $currentHistory[0].IpDate = $nowLocalFormatted
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
        'Room'                = $room
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

$finalRows = foreach ($row in ($resultRows | Sort-Object 'Name', 'Asset Tag')) {
    $ordered = [ordered]@{
        'Name'                = [string]$row.Name
        'Asset Tag'           = [string]$row.'Asset Tag'
        'Location'            = [string]$row.Location
        'Room'                = [string]$row.Room
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

$outputWorkbookPath = [System.IO.Path]::ChangeExtension($outputCsvPath, '.xlsx')
Set-SuccessPingHighlight -CsvPath $outputCsvPath -XlsxPath $outputWorkbookPath

if (Test-Path -LiteralPath $localInputPath) {
    Remove-Item -LiteralPath $localInputPath -Force
}

Write-Host "Tracker run complete. Output written to: $outputCsvPath"
