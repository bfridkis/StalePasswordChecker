clear-host

$checkDate = $outputMode = $fileName = $OUSpecific = $OUToSearch = $searchLevel = $helpRequest = $userInfo = $null

([string]$args).split('-') | %{ 
                                if ($_.Split(' ')[0] -eq "CheckComplianceAsOfDate") { $checkDate = $_.Split(' ')[1] } 
                                elseif ($_.Split(' ')[0] -eq "OutputMode") { $outputMode = $_.Split(' ')[1] }
                                elseif ($_.Split(' ')[0] -eq "OutputFile") { $fileName = $_.Split(' ')[1] }
                                elseif ($_.Split(' ')[0] -eq "OUSpecific") { $OUSpecific = $_.Split(' ')[1] }
                                elseif ($_.Split(' ')[0] -eq "OUToSearch") { $OUToSearch = $_.Split(' ')[1] }
                                elseif ($_.Split(' ')[0] -eq "SearchLevel") { $searchLevel = $_.Split(' ')[1] }
                                elseif ($_.Split(' ')[0] -eq "Help") { $helpRequest = $true }
                              }
if ($checkDate -and ($checkDate -eq "DEFAULT" -or $checkDate -eq "TODAY")) { $checkDate = Get-Date }
if ($outputMode) {
    if ($outputMode -eq "DEFAULT") { $outputMode = 3 }
    if ($outputMode -ne 1 -and $outputMode -ne 2 -and $outputMode -ne 3) { $outputMode = $null }
    if ($outputMode -eq 1) { $CLOutput1 = $true }
}
if ($fileName) { $usePassedFileName = $true } else { $usePassedFileName = $false }
if ($OUToSearch -and !$searchLevel) { $searchLevel = 2 }
if ($OUSpecific -and $OUSpecific -eq "TRUE") { $OUSpecific = "Y" }
if ($OUSpecific -and $OUSpecific -eq "FALSE") { $OUSpecific = "N" }
if ($OUSpecific -and $OUSpecific -ne "Y" -and $OUSpecific -ne "N") { $OUSpecific = $null }
if ($searchLevel -and $searchLevel -eq "DEFAULT") { $searchLevel = 2 }
if ($searchLevel -and $searchLevel -notin 0..2) { $searchLevel = $null }

if (!$helpRequest) {
    write-output "`n"
    write-output "`t`t  *$*$*  Stale Password Checker *$*$*`n"

    if(!$checkDate) { 
        $checkDate = read-host -prompt "Check for compliance as of [Default=Today] (Use format mm/dd/yyyy)"
        if (!$checkDate) { $checkDate = Get-Date }
    }

    if ($checkDate -eq "Q") { exit }
    else { 
        Try { $checkDate = Get-Date($checkDate) -ErrorAction stop }
        Catch { write-output "`n$($_.Exception.Message)" ; exit }
    }

    $saveToFile = $false

    if (!$outputMode) {
        do { 
                $outputMode = read-host -prompt "`nSave To File (1), Console Output (2), or Both (3) [Default=3]" 
                if (!$outputMode) { $outputMode = 3 }
            }
        while ($outputMode -ne 1 -and $outputMode -ne 2 -and $outputMode -ne 3 -and $outputMode.ToUpper() -ne "Q")
    }

    if ($outputMode -eq "Q") { exit }

    $defaultOutFileName = "StalePasswordCheckerOut-$(Get-Date -Format MMddyyyy_HHmmss)"
    if($fileName -and $fileName -eq "DEFAULT") { $fileName = $defaultOutFileName }

    if ($outputMode -eq 1 -or $outputMode -eq 3) {

        $saveToFile = $true
    
        if (!$usePassedFileName) {            
            write-output "`n* To save to any directory other than the current, enter fully qualified path name. *"
            write-output   "*              Leave this entry blank to use the default file name of               *"
            write-output   "*                   '$defaultOutFileName.csv',                  *"
            write-output   "*                 which will save to the current working directory.                 *"
            write-output   "*                                                                                   *"
            write-output   "*  THE '.csv' EXTENSION WILL BE APPENDED AUTOMATICALLY TO THE FILENAME SPECIFIED.   *`n"
        }

        do { 
            if (!$usePassedFileName) { $fileName = read-host -prompt "Save As [Default=$defaultOutFileName.csv]" }

            if ($fileName -and $fileName -eq "Q") { exit }

            $pathIsValid = $true
            $overwriteConfirmed = "Y"

            if (![string]::IsNullOrEmpty($fileName) -and $fileName -ne "B") {

                $fileName += ".csv"
                                        
                $pathIsValid = Test-Path -Path $fileName -IsValid

                if ($pathIsValid) {
                        
                    $fileAlreadyExists = Test-Path -Path $fileName

                    if ($fileAlreadyExists) {

                        do {

                            $overWriteConfirmed = read-host -prompt "File '$fileName' Already Exists. Overwrite (Y) or Cancel (N)"
                                    
                            if ($overWriteConfirmed -eq "Q") { exit }
                            if ($overWriteConfirmed -eq "N") { $usePassedFileName = $false }

                        } while ($overWriteConfirmed -ne "Y" -and $overWriteConfirmed -ne "N" -and 
                                    $overWriteConfirmed -ne "B")
                    }
                }

                else { 
                    write-output "* Path is not valid. Try again. ('b' to return to main, 'q' to quit.) *"
                    $usePassedFileName = $false 
                }
            }
        }
        while (!$pathIsValid -or $overWriteConfirmed -eq "N")
    }

    if (!$OUSpecific -and !$OUToSearch) {
        do { 
            $OUSpecific = read-host -prompt "`nSearch a specific OU? [Y/N, Default=N] (If no OU specified, all users are queried)" 
            if (!$OUSpecific) { $OUSpecific = "N" }
        }
        while ($OUSpecific -ne "Y" -and $OUSpecific -ne "N" -and $OUSpecific -ne "Q")
    }

    if ($OUSpecific -eq "Q") { exit }

    if ($OUSpecific -eq "Y" -and !$OUToSearch) { 
        $OUToSearch = read-host -prompt "`nEnter OU Directory Path (Ex: OU=0_Administrators,OU=Accounts,DC=wmgpcn,DC=local)"
        if ($OUToSearch -eq "Q") { exit }
        if (!$searchLevel) {    
            do { 
                $searchLevel = read-host -prompt "`nSearch Depth? (0=Base, 1=One Level Down, 2=All Children) [Default=2]"
                if (!$searchLevel) { $searchLevel = 2 }
            }
            while ($searchLevel -ne 0 -and $searchLevel -ne 1 -and $searchLevel -ne 2 -and $searchLevel -ne "Q")
        }
    }
    if ($searchLevel -eq "Q") { exit }

    write-output "`nRunning...Please wait..."

    if ($outputMode -eq 1) { 
        if ($OUSpecific -eq "N") { get-aduser -Filter "Enabled -eq 'True'" -Property SamAccountName, PasswordLastSet, Description, MemberOf, Enabled, LastLogonDate -SearchScope Subtree | Tee-Object -Variable userInfo | Measure -OutVariable userCount >$null }
        else { get-aduser -Filter "Enabled -eq 'True'" -Property SamAccountName, PasswordLastSet, Description, MemberOf, Enabled, LastLogonDate -SearchScope $searchLevel | Tee-Object -Variable userInfo | Measure -OutVariable userCount >$null }
        $userInfo | select-object @{n='Name';e={$_.Name}}, SamAccountName, PasswordLastSet, LastLogonDate,
                        @{n ='Is_Admin' ; e = { $(if ($_.memberof -like "*Domain Admins*") {'Yes'} else {'No'})}}, 
                        @{n = 'Is_Compliant' ; e = { $(if ($_.memberof -like "*Domain Admins*") {
                                                          if((New-TimeSpan -Start $_.PasswordLastSet -End $checkDate.AddDays(-90)).Days -lt 0) {
                                                              "YES"
                                                          }
                                                          else { "NO" }
                                                        } 
                                                        else {
                                                            if((New-TimeSpan -Start $_.PasswordLastSet -End $checkDate.AddDays(-365)).Days -lt 0) {
                                                                "YES"
                                                            }
                                                            else { "NO" }
                                                        }
                                                      )
                                                   } 
                        },
                        @{n = 'Number of Days Expired (+) or Until Expiration(-)' ; e = { $(if ($_.memberof -like "*Domain Admins*") {
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
                        },
                        @{n = 'Group Membership' ; e = { Get-ADPrincipalGroupMembership -Identity $_.SamAccountName -ResourceContextServer "process.local" | 
                            Select-Object -ExpandProperty Name | ForEach-Object { "'$_'" } } } |
        # where-object {($_.IsAdmin -eq "Yes" -and $_.PasswordLastSet -lt $checkDate.adddays(-90)) -or 
        #                 ($_.IsAdmin -eq "No" -and $_.PasswordLastSet -lt $checkDate.adddays(-365))} |
        sort-object -Descending 'Number of Days Expired (+) or Until Expiration(-)' -OutVariable Export *>$null
    }
    else {
        if ($OUSpecific -eq "N") { get-aduser -Filter "Enabled -eq 'True'" -Property SamAccountName, PasswordLastSet, Description, MemberOf, Enabled, LastLogonDate -SearchScope Subtree | Where-Object { $_.Name -ne "Roy Lapeyronnie" } | Tee-Object -Variable userInfo | Measure -OutVariable userCount >$null }
        else { get-aduser -Filter "Enabled -eq 'True'" -Property SamAccountName, PasswordLastSet, Description, MemberOf, Enabled, LastLogonDate -SearchScope $searchLevel | Where-Object { $_.Name -ne "Roy Lapeyronnie" } | Tee-Object -Variable userInfo | Measure -OutVariable userCount >$null } 
        $userInfo | select-object @{n='Name';e={$_.Name}}, SamAccountName, PasswordLastSet, LastLogonDate,
                        @{n ='Is_Admin' ; e = { $(if ($_.memberof -like "*Domain Admins*") {'Yes'} else {'No'})}}, 
                        @{n = 'Is_Compliant' ; e = { $(if ($_.memberof -like "*Domain Admins*") {
                                                          if((New-TimeSpan -Start $_.PasswordLastSet -End $checkDate.AddDays(-90)).Days -lt 0) {
                                                              "YES"
                                                          }
                                                          else { "NO" }
                                                        } 
                                                        else {
                                                            if((New-TimeSpan -Start $_.PasswordLastSet -End $checkDate.AddDays(-365)).Days -lt 0) {
                                                                "YES"
                                                            }
                                                            else { "NO" }
                                                        }
                                                      )
                                                   } 
                        },
                        @{n = 'Number of Days Expired (+) or Until Expiration(-)' ; e = { $(if ($_.memberof -like "*Domain Admins*") {
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
                        },
                        @{n = 'Group Membership' ; e = { Get-ADPrincipalGroupMembership -Identity $_.SamAccountName -ResourceContextServer "process.local" | 
                            Select-Object -ExpandProperty Name | ForEach-Object { "'$_'" } } } |
        # where-object {($_.IsAdmin -eq "Yes" -and $_.PasswordLastSet -lt $checkDate.adddays(-90)) -or 
        #                 ($_.IsAdmin -eq "No" -and $_.PasswordLastSet -lt $checkDate.adddays(-365))} |
        sort-object -Descending 'Number of Days Expired (+) or Until Expiration(-)' -OutVariable Export | Format-Table
    }

    if ($saveToFile -and !$fileName) { $Export | export-CSV -Path $($defaultOutFileName + ".csv") -NoTypeInformation }
    elseif ($saveToFile) { $Export | export-CSV -Path $fileName -NoTypeInformation }

    write-output "`n** Finished **"

    $userCount | Format-Table @{n="Users Checked";e={$_.Count};Alignment="left"}
    if ($saveToFile -and !$fileName) { Add-Content -Path $($defaultOutFileName + ".csv") -Value "`r`nUsers Checked: $($userCount[0].Count)" }
    elseif ($saveToFile) { Add-Content -Path $fileName -Value "`r`nUsers Checked: $($userCount[0].Count)" }

    write-output "`n`n** Note: The 'IsAdmin' attribute in the result set is based on the 'MemberOf' property of Get-Aduser cmdlet."
    write-output "**       (More specifically, a user is flagged as admin if his or her account is a member of the 'Domain Admins' group.)"
    write-output "**       If flagged as admin, a user is listed above if the accound password is older than 90 days."
    write-output "**       Non-admins are listed above if the account password is older than 365 days."
    write-output "**       Only enabled accounts are checked. A blank output indicates all passwords are in compliance."

    if (!$CLOutput1) {
        Write-Host "`nPress enter to exit..." -NoNewLine
        $Host.UI.ReadLine()
    }
}

else {
    clear-host

    write-output "`n"
    write-output "`t`t`t`t`t`t`t`t`t*!*!* Stale Password Checker - Help Page*!*!*"

    write-output "SYNTAX"
    write-output "`tStalePasswordCheckerV1.3.ps1 [-CheckComplianceAsOfDate <date> {format: mm/dd/yyyy}]"
    write-output "`t                             [-OutputMode <int32> {1 = Save to File; 2 = Console Output; 3 = Save to File and Console Output}]"
    write-output "`t                             [-OutputFile <string[]> (Note: '.csv' is automatically appended to specified file name. Use `"Default`" for pre-assigned filename.)])]"
    write-output "`t                             [-OUSpecific <string[]> {'TRUE'/'Y' or 'FALSE'/'N'}]"
    write-output "`t                             [-OUToSearch <string[]> {Ex: OU=0_Administrators,OU=Accounts,DC=wmgpcn,DC=local}]"
    write-output "`t                             [-SearchLevel <int32> {0=Base, 1=One Level Down, 2=All Children}]"
    write-output "`n"

    write-host "Press enter to exit..." -NoNewLine
    $Host.UI.ReadLine()
}

# References
# https://stackoverflow.com/questions/44151502/getting-the-no-of-days-difference-from-the-two-dates-in-powershell/44151764
# https://docs.microsoft.com/en-us/powershell/module/addsadministration/get-aduser?view=win10-ps
# https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/select-object?view=powershell-7
# https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/where-object?view=powershell-7
# https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/get-date?view=powershell-7
# https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/tee-object?view=powershell-7#:~:text=The%20Tee%2DObject%20cmdlet%20redirects,is%20displayed%20at%20the%20prompt.
# https://stackoverflow.com/questions/50788152/read-host-always-ends-in-a-colon
# https://stackoverflow.com/questions/20621104/left-alignment-the-output-of-a-powershell-command