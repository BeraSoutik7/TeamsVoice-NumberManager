<#
================================================================================
SCRIPT NAME : Manage-TeamsDirectRoutingNumbers.ps1
DESCRIPTION : Enterprise Direct Routing Number & Extension Management Engine.
              Handles 5-digit extensions, DDI assignment/revocation,
              extension availability discovery, duplicate detection,
              and full tenant inventory exports.
AUTHOR      : Soutik Bera
REPOSITORY  : TeamsVoice-NumberManager
MODULES REQ : MicrosoftTeams (v4.0.0+)
================================================================================
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
    [Parameter(Mandatory = $true, ParameterSetName = "InventoryReport")]
    [switch]$InventoryReport,

    [Parameter(Mandatory = $true, ParameterSetName = "FindExtension")]
    [switch]$FindAvailableExtension,

    [Parameter(Mandatory = $true, ParameterSetName = "AssignNumber")]
    [switch]$AssignNumber,

    [Parameter(Mandatory = $true, ParameterSetName = "RevokeNumber")]
    [switch]$RevokeNumber,

    # Parameters for AssignNumber
    [Parameter(Mandatory = $true, ParameterSetName = "AssignNumber")]
    [string]$Identity,

    [Parameter(Mandatory = $true, ParameterSetName = "AssignNumber")]
    [ValidatePattern('^\+[1-9]\d{6,14}$')]
    [string]$DDI,

    [Parameter(Mandatory = $true, ParameterSetName = "AssignNumber")]
    [ValidatePattern('^\d{5}$')]
    [string]$Extension,

    [Parameter(Mandatory = $false, ParameterSetName = "AssignNumber")]
    [string]$VoiceRoutingPolicy,

    [Parameter(Mandatory = $false, ParameterSetName = "AssignNumber")]
    [string]$TenantDialPlan,

    # Parameters for RevokeNumber
    [Parameter(Mandatory = $true, ParameterSetName = "RevokeNumber")]
    [string]$RevokeIdentity,

    # Parameters for FindAvailableExtension
    [Parameter(Mandatory = $true, ParameterSetName = "FindExtension")]
    [ValidateRange(10000, 99999)]
    [int]$ExtensionRangeStart,

    [Parameter(Mandatory = $true, ParameterSetName = "FindExtension")]
    [ValidateRange(10000, 99999)]
    [int]$ExtensionRangeEnd,

    # Common Parameters
    [Parameter(Mandatory = $false)]
    [string]$CsvOutputPath = ".\reports\TeamsVoice_NumberInventory_$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
)

function Assert-TeamsConnected {
    try {
        $null = Get-CsTenant -ErrorAction Stop
    }
    catch {
        Write-Host "[*] Connecting to Microsoft Teams..." -ForegroundColor Cyan
        Connect-MicrosoftTeams
    }
}

function Get-NormalizedNumberInfo {
    param ([string]$LineUriString)

    $result = [PSCustomObject]@{
        FullLineURI = $LineUriString
        DDI         = ""
        Extension   = ""
    }

    if ([string]::IsNullOrWhiteSpace($LineUriString)) {
        return $result
    }

    $clean = $LineUriString -replace '(?i)^tel:', ''

    if ($clean -match '^(\+[1-9]\d{6,14})(?:;ext=(\d{5}))?$') {
        $result.DDI = $matches[1]
        if ($matches.Count -ge 3 -and $matches[2]) {
            $result.Extension = $matches[2]
        }
    }
    elseif ($clean -match ';ext=(\d{5})') {
        $result.Extension = $matches[1]
    }

    return $result
}

function Get-TenantVoiceInventory {
    Write-Host "[*] Fetching all voice-enabled users and application instances..." -ForegroundColor Cyan

    $rawUsers = Get-CsOnlineUser -Filter "EnterpriseVoiceEnabled -eq 'True'" -ResultSize 5000 -ErrorAction Stop |
        Select-Object Identity, UserPrincipalName, DisplayName, LineURI, OnPremLineURI, VoiceRoutingPolicy, TenantDialPlan

    $rawAppInstances = Get-CsOnlineApplicationInstance -ResultSize 5000 -ErrorAction Stop |
        Select-Object ObjectId, UserPrincipalName, DisplayName, PhoneNumber

    $inventoryList = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($u in $rawUsers) {
        $effectiveUri = if ($u.LineURI) { $u.LineURI } else { $u.OnPremLineURI }
        $parsed = Get-NormalizedNumberInfo -LineUriString $effectiveUri

        $inventoryList.Add([PSCustomObject]@{
            AccountType        = "User"
            DisplayName        = $u.DisplayName
            UserPrincipalName  = $u.UserPrincipalName
            Identity           = $u.Identity
            FullLineURI        = $effectiveUri
            DDI                = $parsed.DDI
            Extension          = $parsed.Extension
            VoiceRoutingPolicy = $u.VoiceRoutingPolicy
            TenantDialPlan     = $u.TenantDialPlan
        })
    }

    foreach ($app in $rawAppInstances) {
        if (-not [string]::IsNullOrWhiteSpace($app.PhoneNumber)) {
            $parsed = Get-NormalizedNumberInfo -LineUriString $app.PhoneNumber
            $inventoryList.Add([PSCustomObject]@{
                AccountType        = "ResourceAccount"
                DisplayName        = $app.DisplayName
                UserPrincipalName  = $app.UserPrincipalName
                Identity           = $app.ObjectId
                FullLineURI        = $app.PhoneNumber
                DDI                = $parsed.DDI
                Extension          = $parsed.Extension
                VoiceRoutingPolicy = "N/A"
                TenantDialPlan     = "N/A"
            })
        }
    }

    return $inventoryList
}

function Ensure-DashboardServer {
    param(
        [int]$Port = 8000
    )

    try {
        $null = Invoke-WebRequest -Uri "http://127.0.0.1:$Port/" -UseBasicParsing -TimeoutSec 2 -ErrorAction Stop
        return
    }
    catch {
        $repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
        $pythonCmd = Get-Command python -ErrorAction SilentlyContinue
        if (-not $pythonCmd) {
            $pythonCmd = Get-Command py -ErrorAction SilentlyContinue
        }

        if (-not $pythonCmd) {
            throw "Python is required to serve the dashboard. Install Python or run the HTML file from a web server."
        }

        Start-Process -FilePath $pythonCmd.Source -ArgumentList @("-m", "http.server", $Port.ToString(), "--bind", "127.0.0.1") -WorkingDirectory $repoRoot -WindowStyle Hidden -PassThru | Out-Null
        Start-Sleep -Seconds 2
    }
}

function Open-TeamsVoiceDashboard {
    param(
        [string]$CsvOutputPath
    )

    $repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
    $dashboardPath = Join-Path $repoRoot "reports\TeamsVoice_ReportDashboard.html"
    if (-not (Test-Path $dashboardPath)) {
        throw "Dashboard file not found: $dashboardPath"
    }

    $reportRelativePath = [System.IO.Path]::GetRelativePath($repoRoot, (Resolve-Path $CsvOutputPath).Path).Replace('\\', '/')
    $dashboardUri = "http://127.0.0.1:8000/reports/TeamsVoice_ReportDashboard.html?csv=$([System.Uri]::EscapeDataString($reportRelativePath))"
    Start-Process $dashboardUri
}

Assert-TeamsConnected

switch ($PSCmdlet.ParameterSetName) {
    "InventoryReport" {
        $inventory = Get-TenantVoiceInventory

        $ddiCounts = $inventory | Where-Object { $_.DDI -ne "" } | Group-Object DDI
        $extCounts = $inventory | Where-Object { $_.Extension -ne "" } | Group-Object Extension

        $duplicateDDIs = $ddiCounts | Where-Object { $_.Count -gt 1 }
        $duplicateExts = $extCounts | Where-Object { $_.Count -gt 1 }

        Write-Host "`n================ INVENTORY SUMMARY ================" -ForegroundColor Cyan
        Write-Host "Total Voice Endpoints Audited : $($inventory.Count)"
        Write-Host "Unique DDIs Discovered        : $($ddiCounts.Count)"
        Write-Host "5-Digit Extensions Discovered : $($extCounts.Count)"

        if ($duplicateDDIs.Count -gt 0) {
            Write-Warning "[-] Found $($duplicateDDIs.Count) duplicate DDI assignment conflicts!"
            foreach ($dup in $duplicateDDIs) {
                Write-Warning "    DDI $($dup.Name) assigned to: $(($dup.Group.UserPrincipalName) -join ', ')"
            }
        }
        else {
            Write-Host "[+] Zero duplicate DDI assignments detected." -ForegroundColor Green
        }

        if ($duplicateExts.Count -gt 0) {
            Write-Warning "[-] Found $($duplicateExts.Count) duplicate 5-Digit Extension conflicts!"
            foreach ($dup in $duplicateExts) {
                Write-Warning "    Extension $($dup.Name) assigned to: $(($dup.Group.UserPrincipalName) -join ', ')"
            }
        }
        else {
            Write-Host "[+] Zero duplicate 5-Digit Extensions detected." -ForegroundColor Green
        }

        $reportDir = Split-Path -Path $CsvOutputPath -Parent
        if ($reportDir -and (-not (Test-Path -Path $reportDir))) {
            New-Item -ItemType Directory -Path $reportDir -Force | Out-Null
        }

        $inventory | Export-Csv -Path $CsvOutputPath -NoTypeInformation -Encoding utf8
        Write-Host "`n[SUCCESS] Full inventory snapshot saved to: $CsvOutputPath" -ForegroundColor Green

        Ensure-DashboardServer
        Open-TeamsVoiceDashboard -CsvOutputPath $CsvOutputPath
        Write-Host "[+] Opening dashboard with the latest inventory data..." -ForegroundColor Green
    }

    "FindExtension" {
        if ($ExtensionRangeStart -ge $ExtensionRangeEnd) {
            throw "ExtensionRangeStart ($ExtensionRangeStart) must be strictly less than ExtensionRangeEnd ($ExtensionRangeEnd)."
        }

        $inventory = Get-TenantVoiceInventory
        $assignedExtensions = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

        foreach ($item in $inventory) {
            if ($item.Extension) {
                $null = $assignedExtensions.Add($item.Extension)
            }
        }

        $availablePool = [System.Collections.Generic.List[string]]::new()

        for ($i = $ExtensionRangeStart; $i -le $ExtensionRangeEnd; $i++) {
            $candidate = $i.ToString("D5")
            if (-not $assignedExtensions.Contains($candidate)) {
                $availablePool.Add($candidate)
            }
        }

        Write-Host "`n================ EXTENSION POOL AUDIT ================" -ForegroundColor Cyan
        Write-Host "Scanned Range       : $ExtensionRangeStart - $ExtensionRangeEnd"
        Write-Host "Allocated in Range  : $(($ExtensionRangeEnd - $ExtensionRangeStart + 1) - $availablePool.Count)"
        Write-Host "Available Extensions: $($availablePool.Count)"

        if ($availablePool.Count -gt 0) {
            Write-Host "`nNext 10 Available 5-Digit Extensions:" -ForegroundColor Green
            $availablePool | Select-Object -First 10 | ForEach-Object { Write-Host "  -> $_" }
        }
        else {
            Write-Error "No unallocated 5-digit extensions left in the specified range!"
        }
    }

    "AssignNumber" {
        $formattedTargetUri = "$DDI;ext=$Extension"

        Write-Host "`n[*] Target User       : $Identity" -ForegroundColor Cyan
        Write-Host "[*] Target Direct URI : $formattedTargetUri" -ForegroundColor Cyan

        $inventory = Get-TenantVoiceInventory
        $existingDdi = $inventory | Where-Object { $_.DDI -eq $DDI -and $_.UserPrincipalName -ne $Identity }
        $existingExt = $inventory | Where-Object { $_.Extension -eq $Extension -and $_.UserPrincipalName -ne $Identity }

        if ($existingDdi) {
            throw "Pre-flight check failed: DDI $DDI is already assigned to $(($existingDdi.UserPrincipalName) -join ', ')."
        }
        if ($existingExt) {
            throw "Pre-flight check failed: Extension $Extension is already assigned to $(($existingExt.UserPrincipalName) -join ', ')."
        }

        if ($PSCmdlet.ShouldProcess("$Identity", "Assign Direct Routing URI [$formattedTargetUri]")) {
            try {
                Set-CsPhoneNumberAssignment `
                    -Identity $Identity `
                    -PhoneNumber $formattedTargetUri `
                    -PhoneNumberType DirectRouting `
                    -ErrorAction Stop

                if ($VoiceRoutingPolicy) {
                    Grant-CsOnlineVoiceRoutingPolicy -Identity $Identity -PolicyName $VoiceRoutingPolicy -ErrorAction Stop
                    Write-Host "[+] Voice Routing Policy '$VoiceRoutingPolicy' granted." -ForegroundColor Green
                }

                if ($TenantDialPlan) {
                    Grant-CsTenantDialPlan -Identity $Identity -PolicyName $TenantDialPlan -ErrorAction Stop
                    Write-Host "[+] Tenant Dial Plan '$TenantDialPlan' granted." -ForegroundColor Green
                }

                Write-Host "[SUCCESS] Successfully assigned $formattedTargetUri to $Identity." -ForegroundColor Green
            }
            catch {
                Write-Error "Failed to assign Direct Routing number to $Identity. Details: $_"
            }
        }
    }

    "RevokeNumber" {
        Write-Host "`n[*] Auditing current telephony assignment for: $RevokeIdentity" -ForegroundColor Cyan

        $targetUser = Get-CsOnlineUser -Identity $RevokeIdentity -ErrorAction SilentlyContinue
        if (-not $targetUser) {
            throw "User '$RevokeIdentity' not found in tenant."
        }

        $currentUri = if ($targetUser.LineURI) { $targetUser.LineURI } else { $targetUser.OnPremLineURI }
        if ([string]::IsNullOrWhiteSpace($currentUri)) {
            Write-Warning "User $RevokeIdentity does not have an active LineURI/DDI assignment."
            return
        }

        $parsed = Get-NormalizedNumberInfo -LineUriString $currentUri
        $numberToRevoke = if ($parsed.Extension) { "$($parsed.DDI);ext=$($parsed.Extension)" } else { $parsed.DDI }

        if ($PSCmdlet.ShouldProcess("$RevokeIdentity", "Remove Direct Routing number [$numberToRevoke]")) {
            try {
                Remove-CsPhoneNumberAssignment `
                    -Identity $RevokeIdentity `
                    -PhoneNumber $numberToRevoke `
                    -PhoneNumberType DirectRouting `
                    -ErrorAction Stop

                Write-Host "[SUCCESS] Successfully released $numberToRevoke from $RevokeIdentity." -ForegroundColor Green
            }
            catch {
                Write-Error "Failed to revoke number from $RevokeIdentity. Details: $_"
            }
        }
    }
}