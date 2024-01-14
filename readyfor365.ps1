$exportedFor365 = "C:\Scripts\officescripts\export\exportedFor365.csv"
$licenseOne = ""
$licenseTwo = ""
$licenseThree = ""
$OUtoSearch = "OU=Staff,DC=example,DC=com"
$domain = "@example.com"
# get skus: Get-MgSubscribedSku | Select SkuPartNumber, SkuId

#Requires -RunAsAdministrator

Import-Module ActiveDirectory

# technique for exclusions came from these fabulous folks:
# https://stackoverflow.com/questions/58334659/powershell-get-aduser-filter-to-exclude-specific-ou-in-the-list
$excludeOUs =
            "OU=DepartedStaff,Staff,DC=example,DC=com",
            "OU=TemporaryWorkers,OU=Staff,DC=example,DC=com"

$reExcludeOUs = '(?:{0})$' -f ($excludeOUs -join '|')

# license groups
$licenseGroups = @("ftstaff", "ptstaff")
$licenseOneGroups = "contractors"
$licenseTwoGroups = "execs"
$licenseThreeGroups = ""
# even if a user is excluded via ou, presence in one of these groups allows a license
$overrideGroups = "SpecialOfficeLicOne|SpecialOfficeLicTwo"

#build the header row of the csv
$exportList = "UserPrincipalName,LicenseSkuId`n"

# output counts for cli
$totalOne = 0
$totalTwo = 0
$totalThree = 0
$totalSpecialOne = 0
$totalSpecialTwo = 0

# ignore user presets
$excludeSpecial = $false;
$excluded = $true;

$buildList = foreach ($group in $licenseGroups) {
   $userList = Get-ADGroupMember $group | select samaccountname, name, @{n='GroupName';e={$group}}

   $userList | forEach-Object {

       # exclude users from specific ous
       $excludeUser = ""
       $specialGroupLookup = ""
       $exclusionException = ""
       $excludeSpecial = $false;
       $userLookup = Get-ADUser -Identity $_.sAMAccountName -Properties SamAccountName | Where-Object { ($_.DistinguishedName -notmatch $reExcludeOUs) }

       # member of cannot be pulled without elevation
       $specialGroupLookup = (Get-ADUser -Identity $_.sAMAccountName -properties memberof | where memberof -match "group123").samaccountname
       $overrideGroupLookup = Get-ADUser -Identity $_.sAMAccountName -properties memberof | Where-Object {$_.memberof -match $overrideGroups}

       # ou exclusion
       if ($userLookup) {
          $excluded = $false;
       }

       # ou exclusion but are in a special group to override ou exclusion
       if ($overrideGroupLookup) {
          $excluded = $false;
          $excludeSpecial = $true;
       }

       # loop through each group & decide whether or not it's configured for a license
       # without  the writeRecord var, samaccount name in the append loop will unintentionally duplicate numerous matches
       $writeRecord = 0

       if ($excluded -eq $false) {
            if ($_.GroupName -contains $licenseThreeGroups) {
                $license = $licenseThree
                $totalThree++;
                $writeRecord = 1
            } elseif ($specialGroupLookup -ne $_.sAMAccountName -and $excludeSpecial -eq $true -and $_.GroupName -eq "SpecialOfficeLicOne") {
                $license = $licenseTwo
                $totalSpecialOne++;
                $writeRecord = 1
            } elseif ($specialGroupLookup -ne $_.sAMAccountName-and $excludeSpecial -eq $true -and $_.GroupName -eq "SpecialOfficeLicTwo") {
                $license = $licenseOne
                $totalSpecialTwo++;
                $writeRecord = 1
            }elseif ($specialGroupLookup -ne $_.sAMAccountName -and $excludeSpecial -eq $false -and $licenseOneGroups -contains $_.GroupName) {
                $license = $licenseOne
                $totalOne++;
                $writeRecord = 1
            } elseif ($specialGroupLookup -ne $_.sAMAccountName -and $excludeSpecial -eq $false -and $licenseTwoGroups -contains $_.GroupName) {
                $license = $licenseTwo
                $totalTwo++;
                $writeRecord = 1
      
            }
            

        # ugly exclusions to avoid duplicates from users in several groups
        if ($writeRecord -eq 1 -and $excluded -eq $false) {
            $userRow = $_.sAMAccountName+="$domain,$license"
            $newLine = ($userRow -join ",") + "`n"
            $exportList += $newLine

            # reset exclusion, otherwise are carried to the next loop user
            $excluded = $true;
        }

      }
 
   }

    #NOTE Needs full path to run from Task Scheduler
    $exportList | Out-File $exportedFor365
}


Write-Output "One licenses to issue: $totalOne"
Write-Output "Two licenses to issue: $totalTwo"
Write-Output "Three licenses to issue: $totalThree"
Write-Output "Special One licenses to issue: $totalSpecialOne"
Write-Output "Special One licenses to issue: $totalSpecialTwo"
