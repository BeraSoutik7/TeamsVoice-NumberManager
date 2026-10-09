# Operational Guidance: Direct Routing 5-Digit Extension & Number Manager

**Author:** Soutik Bera  
**Repository:** TeamsVoice-NumberManager  
**Target Environment:** Microsoft Teams Phone System, Direct Routing (Cloud-Only)  
**Required Modules:** `MicrosoftTeams` (v4.0.0+)  

---

## 1. Environment Verification & Session Setup

Launch PowerShell 7+ in an elevated console and verify that the `MicrosoftTeams` module meets the required baseline version:

```powershell
# Install or update the MicrosoftTeams module
Install-Module -Name MicrosoftTeams -MinimumVersion 4.0.0 -Scope CurrentUser -Repository PSGallery -Force

# Run full inventory audit using default output path (.\reports\)
.\src\Manage-TeamsDirectRoutingNumbers.ps1 -InventoryReport

.\src\Manage-TeamsDirectRoutingNumbers.ps1 -InventoryReport -CsvOutputPath ".\reports\Teams_Voice_Inventory_Audit.csv"

# Find free extensions in the 51000 - 51999 block
.\src\Manage-TeamsDirectRoutingNumbers.ps1 `
    -FindAvailableExtension `
    -ExtensionRangeStart 51000 `
    -ExtensionRangeEnd 51999

# Provision a user with a DDI, 5-digit extension, Voice Routing Policy, and Dial Plan
.\src\Manage-TeamsDirectRoutingNumbers.ps1 -AssignNumber `
    -Identity "alex.wilber@contoso.com" `
    -DDI "+14255550100" `
    -Extension "54321" `
    -VoiceRoutingPolicy "US-National" `
    -TenantDialPlan "DP-5Digit-Internal"

# Preview assignment without applying state changes
.\src\Manage-TeamsDirectRoutingNumbers.ps1 -AssignNumber `
    -Identity "alex.wilber@contoso.com" `
    -DDI "+14255550100" `
    -Extension "54321" -WhatIf

# Release telephone number and extension from an offboarded employee
.\src\Manage-TeamsDirectRoutingNumbers.ps1 -RevokeNumber -RevokeIdentity "alex.wilber@contoso.com"


# 1. Verify single user telephony properties
Get-CsOnlineUser -Identity "alex.wilber@contoso.com" |
    Select-Object DisplayName, UserPrincipalName, LineURI, EnterpriseVoiceEnabled, VoiceRoutingPolicy, TenantDialPlan |
    Format-List

# 2. Inspect generated inventory CSV directly in terminal
Import-Csv (Get-ChildItem .\reports\*.csv | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName |
    Select-Object AccountType, DisplayName, UserPrincipalName, DDI, Extension, VoiceRoutingPolicy |
    Format-Table -AutoSize