clear-host

write-output "`n"
write-output "`t`t  *$*$*  Stale Password Checker *$*$*`n"

$checkDate = read-host -prompt "Check for compliance as of [Default=Today] (Use format mm/dd/yyyy)"
if ($checkDate.ToUpper() -eq "Q") { exit }
if (!$checkDate) { $checkDate = Get-Date }
else { 
    Try { $checkDate = Get-Date($checkDate) -ErrorAction stop }
    Catch { write-output "`n$($_.Exception.Message)" ; exit }
}

$saveToFile = $false

do { $outputMode = read-host -prompt "`nSave To File (1), Console Output (2), or Both (3)" }
while ($outputMode -ne 1 -and $outputMode -ne 2 -and $outputMode -ne 3 -and $outputMode.ToUpper() -ne "Q")

if ($outputMode.ToUpper() -eq "Q") { exit }

$defaultOutFileName = "StalePasswordCheckerOut-$(Get-Date -Format MMddyyyy_HHmmss)"

if ($outputMode -eq 1 -or $outputMode -eq 3) {

    $saveToFile = $true
                
    write-output "`n* To save to any directory other than the current, enter fully qualified path name. *"
    write-output   "*              Leave this entry blank to use the default file name of               *"
    write-output   "*                   '$defaultOutFileName.csv',                  *"
    write-output   "*                 which will save to the current working directory.                 *"
    write-output   "*                                                                                   *"
    write-output   "*  THE '.csv' EXTENSION WILL BE APPENDED AUTOMATICALLY TO THE FILENAME SPECIFIED.   *`n"

    do { 
        $fileName = read-host -prompt "Save As [Default=$defaultOutFileName.csv]"

        if ($fileName -and $fileName.ToUpper() -eq "Q") { exit }

        $pathIsValid = $true
        $overwriteConfirmed = "Y"

        if (![string]::IsNullOrEmpty($fileName) -and $fileName.ToUpper() -ne "B") {

            $fileName += ".csv"
                                        
            $pathIsValid = Test-Path -Path $fileName -IsValid

            if ($pathIsValid) {
                        
                $fileAlreadyExists = Test-Path -Path $fileName

                if ($fileAlreadyExists) {

                    do {

                        $overWriteConfirmed = read-host -prompt "File '$fileName' Already Exists. Overwrite (Y) or Cancel (N)"
                                    
                        if ($overWriteConfirmed.ToUpper() -eq "Q") { exit }

                    } while ($overWriteConfirmed.ToUpper() -ne "Y" -and $overWriteConfirmed.ToUpper() -ne "N" -and 
                                $overWriteConfirmed.ToUpper() -ne "B")
                }
            }

            else { write-output "* Path is not valid. Try again. ('b' to return to main, 'q' to quit.) *" }
        }
    }
    while (!$pathIsValid -or $overWriteConfirmed.ToUpper() -eq "N")
}

do { 
    $OUSpecific = read-host -prompt "`nSearch a specific OU? [Y/N, Default=N] (If no OU specified, all users are queried)" 
    if (!$OUSpecific) { $OUSpecific = "N" }
}
while ($OUSpecific.ToUpper() -ne "Y" -and $OUSpecific.ToUpper() -ne "N" -and $OUSpecific.ToUpper() -ne "Q")

if ($OUSpecific -eq "Q") { exit }

if ($OUSpecific -eq "Y") { 
    $OUToSearch = read-host -prompt "`nEnter OU Directory Path (Ex: OU=0_Administrators,OU=Accounts,DC=wmgpcn,DC=local)"
    if ($OUToSearch.ToUpper() -eq "Q") { exit }
    do { 
        $searchLevel = read-host -prompt "`nSearch Depth? (0=Base, 1=One Level Down, 2=All Children) [Default=2]"
        if (!$searchLevel) { $searchLevel = 2 }
    }
    while ($searchLevel -ne 0 -and $searchLevel -ne 1 -and $searchLevel -ne 2 -and $searchLevel.ToUpper() -ne "Q")
}
if ($searchLevel -eq "Q") { exit }

write-output "`nRunning...Please wait..."

if ($outputMode -eq 1) { 
    if ($OUSpecific -eq "N") { get-aduser -Filter "Enabled -eq 'True' -and Description -notlike '*Exception*'" -Property PasswordLastSet, Description -SearchScope Subtree | Tee-Object -Variable userInfo | Measure -OutVariable userCount >$null }
    else { get-aduser -Filter "Enabled -eq 'True' -and Description -notlike '*Exception*'" -SearchBase $OUToSearch -Property PasswordLastSet, Description -SearchScope $searchLevel | Tee-Object -Variable userInfo | Measure -OutVariable userCount >$null }
    $userInfo | select-object @{n='Name';e={$_.Name}}, PasswordLastSet, 
                    @{n ='Is_Admin' ; e = { $(if ($_.description -like "*administrator*") {'Yes'} else {'No'})}}, 
                    @{n = 'Is_Compliant' ; e = { $(if ($_.description -like "*administrator*") {
                                                      if((New-TimeSpan -Start $_.PasswordLastSet -End $checkDate.AddDays(-90)).Days -lt 90) {
                                                          "YES"
                                                      }
                                                      else { "NO" }
                                                    } 
                                                    else {
                                                        if((New-TimeSpan -Start $_.PasswordLastSet -End $checkDate.AddDays(-90)).Days -lt 365) {
                                                            "YES"
                                                        }
                                                        else { "NO" }
                                                    }
                                                  )
                                               } 
                    },
                    @{n = 'Number of Days Expired (+) or Until Expiration(-)' ; e = { $(if ($_.description -like "*administrator*") {
                                                                if($_.PasswordLastSet) {
                                                                    (New-TimeSpan -Start $_.PasswordLastSet -End $checkDate.AddDays(-90)).Days
                                                                }
                                                                else { "No 'PasswordLastSet' Available" }
                                                            } 
                                                            else {
                                                                if($_.PasswordLastSet) {
                                                                    (New-TimeSpan -Start $_.PasswordLastSet -End $checkDate.AddDays(-365)).Days
                                                                }
                                                                else { "No 'PasswordLastSet' Available" }
                                                              }
                                                            )
                                                          } 
                    } | 
    # where-object {($_.IsAdmin -eq "Yes" -and $_.PasswordLastSet -lt $checkDate.adddays(-90)) -or 
    #                 ($_.IsAdmin -eq "No" -and $_.PasswordLastSet -lt $checkDate.adddays(-365))} |
    sort-object -Descending 'Number of Days Expired (+) or Until Expiration(-)' -OutVariable Export *>$null
}
else {
    if ($OUSpecific -eq "N") { get-aduser -Filter "Enabled -eq 'True' -and Description -notlike '*Exception*'" -Property PasswordLastSet, Description -SearchScope Subtree | Tee-Object -Variable userInfo | Measure -OutVariable userCount >$null }
    else { get-aduser -Filter "Enabled -eq 'True' -and Description -notlike '*Exception*'" -SearchBase $OUToSearch -Property PasswordLastSet, Description -SearchScope $searchLevel | Tee-Object -Variable userInfo | Measure -OutVariable userCount >$null } 
    $userInfo | select-object @{n='Name';e={$_.Name}}, PasswordLastSet, 
                    @{n ='Is_Admin' ; e = { $(if ($_.description -like "*administrator*") {'Yes'} else {'No'})}}, 
                    @{n = 'Is_Compliant' ; e = { $(if ($_.description -like "*administrator*") {
                                                      if((New-TimeSpan -Start $_.PasswordLastSet -End $checkDate.AddDays(-90)).Days -lt 90) {
                                                          "YES"
                                                      }
                                                      else { "NO" }
                                                    } 
                                                    else {
                                                        if((New-TimeSpan -Start $_.PasswordLastSet -End $checkDate.AddDays(-90)).Days -lt 365) {
                                                            "YES"
                                                        }
                                                        else { "NO" }
                                                    }
                                                  )
                                               } 
                    },
                    @{n = 'Number of Days Expired (+) or Until Expiration(-)' ; e = { $(if ($_.description -like "*administrator*") {
                                                                if($_.PasswordLastSet) {
                                                                    (New-TimeSpan -Start $_.PasswordLastSet -End $checkDate.AddDays(-90)).Days
                                                                }
                                                                else { "No 'PasswordLastSet' Available" }
                                                            } 
                                                            else {
                                                                if($_.PasswordLastSet) {
                                                                    (New-TimeSpan -Start $_.PasswordLastSet -End $checkDate.AddDays(-365)).Days
                                                                }
                                                                else { "No 'PasswordLastSet' Available" }
                                                              }
                                                            )
                                                          } 
                    } |
    # where-object {($_.IsAdmin -eq "Yes" -and $_.PasswordLastSet -lt $checkDate.adddays(-90)) -or 
    #                 ($_.IsAdmin -eq "No" -and $_.PasswordLastSet -lt $checkDate.adddays(-365))} |
    sort-object -Descending 'Number of Days Expired (+) or Until Expiration(-)' -OutVariable Export | Format-Table
}

if ($saveToFile -and !$fileName) { $Export | export-CSV -Path $($defaultOutFileName + ".csv") -NoTypeInformation }
elseif ($saveToFile) { $Export | export-CSV -Path $fileName -NoTypeInformation }

write-output "`n** Finished **"

$userCount | Format-Table @{n="Users Checked";e={$_.Count};Alignment="left"}

write-output "`n`n** Note: The 'IsAdmin' attribute in the result set is based on the 'adminCount' property of Get-Aduser cmdlet."
write-output "**       If flagged as admin, a user is listed above if the accound password is older than 90 days."
write-output "**       Non-admins are listed above if the account password is older than 365 days."
write-output "**       Only enabled accounts are checked. A blank output indicates all passwords are in compliance."

Write-Host "`nPress enter to exit..." -NoNewLine
$Host.UI.ReadLine()

# References
# https://stackoverflow.com/questions/44151502/getting-the-no-of-days-difference-from-the-two-dates-in-powershell/44151764
# https://docs.microsoft.com/en-us/powershell/module/addsadministration/get-aduser?view=win10-ps
# https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/select-object?view=powershell-7
# https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/where-object?view=powershell-7
# https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/get-date?view=powershell-7
# https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/tee-object?view=powershell-7#:~:text=The%20Tee%2DObject%20cmdlet%20redirects,is%20displayed%20at%20the%20prompt.
# https://stackoverflow.com/questions/50788152/read-host-always-ends-in-a-colon
# https://stackoverflow.com/questions/20621104/left-alignment-the-output-of-a-powershell-command