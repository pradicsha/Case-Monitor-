<#
.SYNOPSIS
    Uploads the generated Salesforce status report to a SharePoint document library.

.DESCRIPTION
    Sample reference script invoked by app.py when POWERSHELL_SCRIPT_PATH is set.
    Adjust the SharePoint site URL, library name, and authentication method to
    match your tenant. Uses the PnP.PowerShell module.

.PARAMETER ReportPath
    Absolute path to the .xlsx file produced by app.py. Passed automatically.

.EXAMPLE
    pwsh -File ./scripts/upload_csv.ps1 -ReportPath "C:\reports\status_report.xlsx"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ReportPath
)

# ---- CONFIGURE THESE FOR YOUR TENANT ----
$SiteUrl       = "https://your-tenant.sharepoint.com/sites/YourSite"
$LibraryName   = "Documents"
$TargetFolder  = "Salesforce Reports"
# ------------------------------------------

if (-not (Test-Path $ReportPath)) {
    Write-Error "Report file not found: $ReportPath"
    exit 1
}

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Error "PnP.PowerShell module not installed. Run: Install-Module PnP.PowerShell -Scope CurrentUser"
    exit 1
}

try {
    Import-Module PnP.PowerShell -ErrorAction Stop
    Connect-PnPOnline -Url $SiteUrl -Interactive

    Add-PnPFile `
        -Path $ReportPath `
        -Folder "$LibraryName/$TargetFolder" `
        -ErrorAction Stop | Out-Null

    Write-Host "[+] Uploaded $(Split-Path $ReportPath -Leaf) to $SiteUrl/$LibraryName/$TargetFolder"
}
catch {
    Write-Error "Upload failed: $_"
    exit 1
}
finally {
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}
