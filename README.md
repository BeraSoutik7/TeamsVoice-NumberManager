# TeamsVoice-NumberManager

[![PowerShell](https://img.shields.io/badge/PowerShell-7.2+-blue.svg)](https://github.com/PowerShell/PowerShell)
[![Platform](https://img.shields.io/badge/Platform-Microsoft%20Teams%20Phone-0078D4.svg)](https://learn.microsoft.com/microsoftteams/teams-phone)
[![Architecture](https://img.shields.io/badge/Telephony-Direct%20Routing-orange.svg)](docs/ARCHITECTURE.md)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

An enterprise lifecycle automation engine for **Microsoft Teams Phone System Direct Routing** environments utilizing dual E.164 Direct Dial-In (DDI) and 5-digit extension routing architectures.

---

## 🎯 Operational Challenges Solved

In enterprise Direct Routing topologies:
* **Extension Collisions:** Native Teams admin tools lack automatic safeguards to prevent two users from claiming the same internal 5-digit extension suffix (`tel:+14255550100;ext=51234`).
* **Unassigned Number Blind Spots:** Direct Routing numbers are managed on the carrier/SBC side, meaning Teams does not natively track unallocated extensions inside internal ranges.
* **Deprovisioning Drift:** When users depart, unassigned numbers can stay bound or orphaned, leading to routing errors.

`TeamsVoice-NumberManager` audits tenant assignments, performs pre-flight collision checks, dynamically searches for free 5-digit extensions, and manages onboarding and offboarding.

---

## 🚀 Key Features

* **Dual Architecture Normalization:** Seamlessly separates full E.164 DDIs from internal 5-digit extensions.
* **Collision Detection:** Scans users and voice application resource accounts to surface duplicate DDI or extension assignments.
* **Next Available Extension Finder:** Scans designated numeric branch ranges (e.g., `51000` to `51999`) and discovers available extensions in RAM using hash-set indexing.
* **Zero-Collision Provisioning:** Validates availability before invoking `Set-CsPhoneNumberAssignment`.
* **Safe Deprovisioning:** Automatically isolates and strips assigned LineURIs and recycling numbers cleanly.
* **Non-Destructive Simulation:** Built-in PowerShell `-WhatIf` and `-Confirm` support (`SupportsShouldProcess`).

---

## 📋 Prerequisites

* **PowerShell 7.2+ Core** (Recommended) or Windows PowerShell 5.1
* **MicrosoftTeams** PowerShell Module (v4.0.0+)
* **Role:** Teams Administrator or Global Administrator

```powershell
Install-Module -Name MicrosoftTeams -MinimumVersion 4.0.0 -Scope CurrentUser -Repository PSGallery -Force