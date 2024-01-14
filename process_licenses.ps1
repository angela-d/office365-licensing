# ty to this excellent article:
# https://www.sharepointdiary.com/2023/04/how-to-connect-to-microsoft-graph-api-from-powershell.html

# script written for the following environment:
# Scopes                 : {User.ReadWrite.All, LicenseAssignment.ReadWrite.All, User.ManageIdentities.All}
# AuthType               : AppOnly
# TokenCredentialType    : ClientCertificate
# PSHostVersion          : 5.1.14393.6343+

# CONFIG
	# To should be a single address
	$To = "you@example.com"
	$From = "noreply@example.com"
	$SMTPServer = "relayer.example.com"
	$SMTPPort = "25"

    # https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Overview/
    $tenantID = ""
    $appID = ""
    $thumbprint = ""
    $csvPath = "C:\Scripts\officescripts\export\exportedFor365.csv"

    # breakdown of completed data - used to gauge whether or not to trigger an email alert for inactivity (file is written to at end of script execution)
    $resultFile = "C:\Scripts\officescripts\export\results.txt"
    $newlyLicensedFile = "C:\Scripts\officescripts\export\newlylicensed.txt"
    # also config the whichLicense function
    # current licenses - used for sending notices when at capacity
    $licenseOne = ""
    $licenseOneTotalSeats = "11049"
    $licenseTwo = ""
    $licenseTwoTotalSeats = "10000"
    $licenseThree = ""
    $licenseThreeTotalSeats = "10000"
    $licenseSpecial = ""
    $licenseSpecialTotalSeats = "1120"
    # license to assign if desired is at capacity
    $failoverLicense = $licenseTwo

    # separators
    $OFS = "`r`n"
    $debug = 0;
# END CONFIG

function test-mguser {
# https://old.reddit.com/r/PowerShell/comments/w2ursu/check_for_existence_of_user_microsoft_graph/
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true,position=0)]
        [string]$upn
    )

    try {
        $usersID = get-mguser -userid $upn -ErrorAction stop > $null
        return ($true)
    } catch {
        return ($false)
    }
}

function consumedSeats($license,$user) {
    $licenseLookup = whichLicense $license
    if ($licenseLookup -eq "Dynamics 365 Customer Voice Additional Responses") {
        $seats = $licenseTwoTotalSeats
    } elseif ($licenseLookup -eq "Dynamics 365 for Sales and Customer Service Enterprise Edition") {
        $seats = $licenseOneTotalSeats
    } elseif ($licenseLookup -eq "Exchange Online (PLAN 2)") {
        $seats = $licenseThreeTotalSeats
    }elseif ($licenseLookup -eq "Microsoft 365 Apps for enterprise") {
        $seats = $licenseSpecialSeats
    } elseif ($licenseLookup -eq "Microsoft 365 Business Basic") {
        $seats = $licenseSpecialTotalSeats
    }

    # because we don't need an email everytime this script runs, only warn with 1 remaining license
    $consumptionCheck = (Get-MgSubscribedSku | where SkuId -eq $license)

    if ($user) {
        if ($consumptionCheck.ConsumedUnits -ge $seats) {

            # set this to true so we don't waste time trying to apply a license we don't have
            return $true
        }
    } else {
        # send a warning when a license is maxed out
        if ($consumptionCheck.ConsumedUnits -eq $seats -and $licenseLookup -ne "some special license") {
            $sendWarning = "Seats fully consumed for $($consumptionCheck.SkuPartNumber) - $license`nAll $seats licenses are in use."
            sendEmail "Office 365 License capacity for $($consumptionCheck.SkuPartNumber) is maxed out" $sendWarning

            return $consumptionCheck.ConsumedUnits
        }
    }
}


function sendEmail($subject,$messageBody) {
    # send an email on exception
	Send-MailMessage -From $From -To $To -Subject $subject -Body $messageBody -SmtpServer $SMTPServer -port $SMTPPort
	Write-Warning "Email sent to $To regarding $Subject"
}

# for a failsafe, if the $resultFile hasn't been written in over 3 days, send an alert
if ($resultFile.LastWriteTime -gt (Get-Date).AddHours(-72)) {
    sendEmail "Possible 365 License Sync Issue" "Office 365 breakdown: $resultFile has not been modified in more than 3 days, better check it.`n`nLicense syncing could be jagged up.`n`n=====`nThis is an automated messge from the merge script."
}

function whichLicense($license) {
    # convert for readability during debug

    if ($license -eq $licenseOne) {
        return "Dynamics 365 Customer Voice Additional Responses"
    } elseif ($license -eq $licenseTwo) {
        return "Dynamics 365 for Sales and Customer Service Enterprise Edition"
    } elseif ($license -eq $licenseThree) {
        return "Exchange Online (PLAN 2)"
    } elseif ($license -eq $licenseSpecial) {
        return "Microsoft 365 Business Basic"
    }
}


# kinda spaghetti-y & redundant with the main processing block.. :/
function removeOldLicense($userId,$preferredLicense,$failover) {
    $unwantedLicense = (Get-MgUserLicenseDetail -UserId $userId | ? {$_.SkuId -ne $preferredLicense}).SkuId
    Write-Output " > Preferred: $preferredLicense"
    Write-Output " > In Use: $unwantedLicense"
    Write-Output " > Failover: $failover"
   
   if ($debug -eq 0 -and $unwantedLicense -and $unwantedLicense -ne $failover){
        Write-Output "Removing unwanted license: $unwantedLicense from $userId"
        Set-MgUserLicense -UserId $userId -RemoveLicenses $unwantedLicense -AddLicenses @()
        return $true
   } elseif (($debug -eq 1 -and $unwantedLicense)  -and ($unwantedLicense -ne $failover)) {
        Write-Warning "DEBUG mode on: if debug was off, SkuID $unwantedLicense would be purged from $userId"
        return $false
   } elseif ($unwantedLicense -eq $failover) {
        Write-Warning "Failover license in use by $userId; not removing"

        return $false
   }

}

# deprecation table
# https://learn.microsoft.com/en-us/powershell/microsoftgraph/azuread-msoline-cmdlet-map?view=graph-powershell-1.0

# ps keeps connecting on old tls.. to-do, for now:
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# install-module Microsoft.Graph.Users, if not installed yet
Import-Module Microsoft.Graph.Users

# if you make changes to your api privs, it'll inherit from your last session, soo:
Disconnect-MgGraph -ErrorAction SilentlyContinue
# will trigger: 'No application to sign out from.' without erroraction if there's no queued session, unsure how to error handle

Write-Host "Connecting to Graph App ID: $appID on Tenant ID: $tenantID with thumbprint: $thumbprint"
Connect-MgGraph -ClientID $appID -TenantId $tenantID -CertificateThumbprint $thumbprint -NoWelcome
Write-Output "=================="
Write-Output "Connected to Graph!"
Write-Output "==================`n`n"

if ($debug -eq 1){
    Write-Warning "In DEBUG mode, changes will not be made to Office 365 licenses or accounts!"
    $debugMsg = "Debug mode active during script run.$OFS"
} else {
    $debugMsg = "Debug mode inactive during script run.$OFS"
}

# write a log of newly licensed users for this round
if (Test-Path $newlyLicensedFile) {
   Remove-Item $newlyLicensedFile
   Write-Host "Removed old log $newlyLicensedFile"
}

New-Item -path "$newlyLicensedFile" -value "$debugMsg $OFS"
Write-Host "Created new $newlyLicensedFile"

$ExistingLicenses = 0;
$NewLicenses = 0;
$NotIn365 = @();
$missingUser = 0;

Write-Output "Importing $csvPath to the script...`n`n"

Import-Csv $csvPath | forEach {

    # make sure the user exists, first
    $userExists = (test-mguser $_.UserPrincipalName)

    # main processing block - loop through the users
	if ($userExists -eq $true) {

        # compare licenses
        $usersLicenses = (Get-MgUserLicenseDetail -UserId $_.UserPrincipalName | Where SkuId -eq $_.LicenseSkuId)

        # remove any license not controlled by this script
        $needToRemove = removeOldLicense $_.UserPrincipalName $_.LicenseSkuId $failoverLicense

        # sku name instead of id
        $friendlyLicenseName = whichLicense $_.LicenseSkuId 

        # user already has a license
		if ($usersLicenses) {
			'Existing license: ' + $_.UserPrincipalName +" " + $friendlyLicenseName;
			$ExistingLicenses++;
        }elseif (Get-MgUserLicenseDetail -UserId $_.UserPrincipalName | Where SkuId -eq $failoverLicense) {
			'Existing FAILOVER license: ' + $_.UserPrincipalName +" " + $friendlyLicenseName;
			$ExistingLicenses++;
		} else {
            # user's object id is req for update-mguser
            $userObjectID = (get-mguser -UserId $_.UserPrincipalName).Id
			' >> NEW LICENSE: ' + $_.UserPrincipalName +" " + $friendlyLicenseName + " User ID: " + $userObjectID;

            $unavailableLicense = consumedSeats $_.LicenseSkuId $_.UserPrincipalName

            if ($debug -eq 0){
                # You can only assign licenses to user accounts that have the UsageLocation property set to a valid ISO 3166-1 alpha-2 country code.
                # https://learn.microsoft.com/en-us/microsoft-365/enterprise/assign-licenses-to-user-accounts-with-microsoft-365-powershell?view=o365-worldwide
                if ($unavailableLicense -ne $true) {
                    Update-MgUser -UserId $userObjectID -UserPrincipalName $_.UserPrincipalName -UsageLocation "US"
                    # addlicenses must be a hashtable
                    Set-MgUserLicense -UserId $_.UserPrincipalName -AddLicenses @{SkuId = $_.LicenseSkuId} -RemoveLicenses @()
                } elseif ($needToRemove -eq $true -or $usersLicenses -eq $null) {
                    Write-Warning "At capacity for the license!!! Will attempt to assign the failover license."
                    Update-MgUser -UserId $userObjectID -UserPrincipalName $_.UserPrincipalName -UsageLocation "US"
                    # addlicenses must be a hashtable
                    Set-MgUserLicense -UserId $_.UserPrincipalName -AddLicenses @{SkuId = $failoverLicense} -RemoveLicenses @()
                }
            } else {
                Write-Warning "$($_.UserPrincipalName) UserId: $userObjectID (used for Update-MgUser)"
                Write-Warning "In DEBUG mode, so licenses and changes to $($_.UserPrincipalName) were NOT committed!"

                if ($unavailableLicense -eq $true) {
                    Write-Warning "At capacity for the license!!! If this was live, would attempt to assign the license!"
                }
            }
			$NewLicenses++;

            # log data to file irregardless of debug mode
            Add-Content -Path $newlyLicensedFile -Value "$($_.UserPrincipalName) -UserId $($_.UserPrincipalName) License type: $friendlyLicenseName"
		}
	}else{
		$NotIn365 += $_.UserPrincipalName;
        Write-Warning "`n$($_.UserPrincipalName) does not appear to be in 365...`n"
        sendEmail "Account Not in Office 365" "$($_.UserPrincipalName) is in AD, but not 365.`n`nPlease validate AD attributes for a mismatch between the two, if the user is in 365.`n`nThis is an automated message triggered by the sync script for 365."
	}
}

# license capacity check
consumedSeats $licenseOne
consumedSeats $licenseSpecial
consumedSeats $licenseTwo

$NotIn365List = $NotIn365 | ft | out-string
$OutputText = "Already have licenses: " + $ExistingLicenses + $OFS +  "New Licenses: " + $NewLicenses + $OFS + $OFS + "Users not in 365:" + $OFS + $NotIn365List + $OFS + $debugMsg;
$OutputText | Out-File -filepath $resultFile


Disconnect-MgGraph
Write-Output "Disconnected from Graph"