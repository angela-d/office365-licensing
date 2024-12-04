# Onsite Active Directory / Office 365 License Integration
Powershell scripts using Microsoft Graph to automate Office 365 license assignments based on security groups in Active Directory.

## Pre-requisites
- Already syncing your onsite Active Directory to Office 365 using Azure / Entra Active Directory Sync / Entra Cloud Sync
- Configure the two attached Powershell scripts with your environment's config

## Features of the Scripts
- Exclude organizational units / OUs from O365 license assignment
- Apply licenses based on user's security group
- License management / failover to different license when primary seats are exhausted
- Override exclusions; if a user is in an excluded OU but part of a special override group, they get assigned a license whereas peers in the same OU do not (unless they're also members of the special security group)
- Easily mass remove/delete unwanted or old / obsolete licenses -- anything not designated by the script gets replaced (so run in debug mode, first!)
- Failure detection; if syncing activity isn't detected within 3 days, notify an admin

### Config / Setup
***
**readyfor365.ps1**
***
***
This script is phase one and should run before **process_licenses.ps1**

It generates a .csv of users to process and their designated license.


Config options for readyfor365.ps1:


- **$exportedFor365** = Absolute (full) path to where the CSV for the preceeding script will go; if using nested directories, make sure you create them; this script will not do it for you.
- **$licenseOne**, **$licenseTwo**, **$licenseThree** = One of your license assignments' license SKU
- **$OUtoSearch** = Organizational unit of where the user accounts that need licenses are
- **$domain** = Your O365 domain
- **$excludeOUs** = OUs you want to exclude from being processed & subsequently assigned licenses
- **$licenseGroups** = ALL of your security groups that get any type of license
- **$overrideGroups** = Users in these group(s) get a license, even if they belong to an exluded OU; separate multiple group names by a pipe: |
- **$exportList** = Must match for *process_license.ps1* - don't touch

Optional / only seen when running via command line - name "one, two, etc" accordingly to match your environment but avoid touching the variables unless you also change those in your code:
```powershell
Write-Output "One licenses to issue: $totalOne"
Write-Output "Two licenses to issue: $totalTwo"
Write-Output "Three licenses to issue: $totalThree"
Write-Output "Special One licenses to issue: $totalSpecialOne"
Write-Output "Special One licenses to issue: $totalSpecialTwo"
```

Config options for process_license.ps1:
- **Install the certificate you'll be using** to the *Local Machine* where this script will be running
- **Make note of the SN** mmc.exe > File > Add/Remove Snap-In > Certificates > Add > Computer Account > Certificates > Personal > Certificates > locate your cert > Open it > Details > Subject; this goes in the **$certSN** variable
- **$To** = Email address for notifications if there's licensing problems; like a helpdesk email
- **$From** = Sender of notifications for licensing problems
- **$SMTPServer** = SMTP host of your mail server
- **$SMTPPort** = SMTP port of your mail server
- **$licenseOne**, **$licenseTwo**, **$licenseThree**, **$licenseSpecial** = Leave these blank; they're just initializing the variables
- **$licenseOneTotalSeats**, **$licenseTwoTotalSeats**, **$licenseThreeTotalSeats**, **$licenseSpecialSeats** = Seat values alloted to your origanization; see How-To section on how to obtain
- **$failoverLicense** = License to assign if desired is at capacity
- `function consumedSeats` and `function whichLicense` - Set user-readable names for your licenses in this conditional lookup -- does not need to match Microsoft's naming convention
- **$debug** = Set to `1` when you **do not want to make changes to O365**!  `0` will make immediate changes to your O365 tenant!

Obtain the following from [entra.microsoft.com](https://entra.microsoft.com):
- **$tenantID** = Tenant ID of your MS Office ([direct link to obtain this stuff](https://entra.microsoft.com/#home))
- **$appID** = same as above
- **$csvPath** = this must match **$exportedFor365** from the *readyfor365.ps1* script
- **$resultFile** = Absolute (full) path to log textual output from the process; is written over on each run.
- **$newlyLicensedFile** = Absolute (full) path to log textual output from the process for every new license assignee; is written over on each run.



***
**process_license.ps1**
***
***
This script is phase two and should only run after readyfor365.ps1 has completed.




***

### How To

Once you've got your Graph credentials, you can login via terminal.
It's recommended to do so for testing & debugging purposes and to make sure it works before setting your script loose.

By default, Powershell may try utilizing antiquated TLS.. bypass such by running the following in your Powershell ISE:
```powershell
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
```

If this is your first time using graph, install the module:
```powershell
install-module Microsoft.Graph.Users
```

Import it:
```powershell
Import-Module Microsoft.Graph.Users
```

- Now, copy/paste your graph variables from your **process_licenses.ps1** script, from `$tenantID` down to `$thumbprint`

- Run the following to wrap it all together:
    ```powershell
    Write-Host "Connecting to Graph App ID: $appID on Tenant ID: $tenantID with thumbprint: $thumbprint"
    Connect-MgGraph -ClientID $appID -TenantId $tenantID -CertificateThumbprint $thumbprint -NoWelcome
    Write-Output "=================="
    Write-Output "Connected to Graph!"
    Write-Output "==================`n`n"
    ```
    - Once you're connected to Graph via ISE/terminal, you can run sample queries or get licensing info:

- Get a list of available SKUs to put in readyfor365.ps1:
    ```powershell
    Get-MgSubscribedSku | Select SkuPartNumber, SkuId
    ```

- Set your `$debug` variable to `1` and test your implemenation
- Once everything is working smoothly, automate to run headless via Task Scheduler

**NOTE!!!** 
- If you make changes to your api privileges from the Azure / Entra portal, Graph will inherit from your last session and not pick up the changes until you disconnect; to do so:

    ```powershell
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    ```

### Overview

- readyfor365.ps1 = Does not make changes to Active Directory or Office 365; reads AD and prepares a .csv based on what's configured within the script
- process_licese.ps1 = Depending on debug value, will adjust/remove licenses based on configuration within the script.

### Credits
Huge thank you to [Salaudeen Rajack](https://www.sharepointdiary.com/2023/04/how-to-connect-to-microsoft-graph-api-from-powershell.html) for the well-written article on Microsoft Graph.

Salaudeen's concise writeup made the transition from MSOnline cmdlet's to Microsoft Graph far easier.

Microsoft's depreciation table for MSOnline: [https://learn.microsoft.com/en-us/powershell/microsoftgraph/azuread-msoline-cmdlet-map?view=graph-powershell-1.0](https://learn.microsoft.com/en-us/powershell/microsoftgraph/azuread-msoline-cmdlet-map?view=graph-powershell-1.0)
