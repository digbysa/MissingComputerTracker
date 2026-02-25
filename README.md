# MissingComputerTracker

PowerShell tracker for missing devices using `MissingDeviceList.csv` input and `SearchedDeviceList.csv` output.

## What it does

`Track-HostnameIPs.ps1` now performs the full workflow when run:

1. Checks for `C:\Users\da1701_sa\Desktop\New-Inventory-Tool\Output\MissingDeviceList.csv`.
2. If found, copies/replaces `MissingDeviceList.csv` in the same folder as the script.
3. Uses that local `MissingDeviceList.csv` as input.
4. Updates/creates `Output\SearchedDeviceList.csv` in the same script folder.
5. Removes local `MissingDeviceList.csv` after processing.

Input columns required in `MissingDeviceList.csv`:

- `Timestamp`
- `Name`
- `Asset Tag`
- `Location`

Output includes:

- `Name`
- `Asset Tag` (unique key)
- `Location` (updated from newest input row for each asset tag)
- `Successfully Pinged` (`Yes` if ping succeeded at least once across runs)
- `Latest Data --> IP Date / IP Address / Subnet / Logged User`
- `Previous Data N --> IP Date / IP Address / Subnet / Logged User` for IP history

> Note: CSV cannot create true merged Excel headers. The script uses grouped header names (for example, `Latest Data --> ...`) to preserve the same structure in CSV form.

Subnet lookup:

- Script reads `SiteSubnets.csv` from the same folder as the script.
- First column = subnet CIDR (for example `10.0.0.0/24`), second column = display label.

## Run manually

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Track-HostnameIPs.ps1
```

## Run at 10:00 AM and 3:00 PM daily

Use `Run-TrackerLoop.ps1` (defaults to `10:00` and `15:00`):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Run-TrackerLoop.ps1
```

You can still override run times:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Run-TrackerLoop.ps1 -RunTimes @([TimeSpan]::FromHours(10), [TimeSpan]::FromHours(15))
```
