clear-host

write-output "`n"
write-output "`t`t  *$*$*  Stale Password Checker *$*$*`n"

$checkDate = read-host -prompt "Check for compliance as of [Default=Today] (Use format mm/dd/yyyy)"
if ($checkDate.ToUpper() -eq "Q") { exit }
if (!$checkDate) { $checkDate = Get-Date }
else { $checkDate = Get-Date($checkDate) }

$saveToFile = $false

do { $outputMode = read-host -prompt "`nSave To File (1), Console Output (2), or Both (3)" }
while ($outputMode -ne 1 -and $outputMode -ne 2 -and $outputMode -ne 3)

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

write-output "`nRunning...Please wait..."

if ($outputMode -eq 1) { 
    get-aduser -Filter "Enabled -eq 'True'" -Property PasswordLastSet, description | 
        select-object @{n='Name';e={$_.Name}}, PasswordLastSet, 
                      @{n ='IsAdmin' ; e = { $(if ($_.description -like "*administrator*") {'Yes'} else {'No'})}},  
                      @{n = 'Number of Days Expired' ; e = { $(if ($_.description -like "*administrator*") {
                                                                if($_.PasswordLastSet) {
                                                                    (New-TimeSpan -Start $_.PasswordLastSet -End $checkDate.AddDays(-90)).Days
                                                                }
                                                                else { "Password Never Set" }
                                                               } 
                                                               else {
                                                                 if($_.PasswordLastSet) {
                                                                     (New-TimeSpan -Start $_.PasswordLastSet -End $checkDate.AddDays(-365)).Days
                                                                 }
                                                                 else { "Password Never Set" }
                                                              }
                                                              )
                                                             } 
                       } | 
        where-object {($_.IsAdmin -eq "Yes" -and $_.PasswordLastSet -lt $checkDate.adddays(-90)) -or 
                      ($_.IsAdmin -eq "No" -and $_.PasswordLastSet -lt $checkDate.adddays(-365))} |
        sort-object -Descending 'Number of Days Expired' -OutVariable Export *>$null
}
else {
    get-aduser -Filter "Enabled -eq 'True'" -Property PasswordLastSet, description | 
        select-object @{n='Name';e={$_.Name}}, PasswordLastSet, 
                      @{n ='IsAdmin' ; e = { $(if ($_.description -like "*administrator*") {'Yes'} else {'No'})}},  
                      @{n = 'Number of Days Expired' ; e = { $(if ($_.description -like "*administrator*") {
                                                                if($_.PasswordLastSet) {
                                                                    (New-TimeSpan -Start $_.PasswordLastSet -End $checkDate.AddDays(-90)).Days
                                                                }
                                                                else { "Password Never Set" }
                                                               } 
                                                               else {
                                                                 if($_.PasswordLastSet) {
                                                                     (New-TimeSpan -Start $_.PasswordLastSet -End $checkDate.AddDays(-365)).Days
                                                                 }
                                                                 else { "Password Never Set" }
                                                              }
                                                              )
                                                             } 
                       } | 
        where-object {($_.IsAdmin -eq "Yes" -and $_.PasswordLastSet -lt $checkDate.adddays(-90)) -or 
                      ($_.IsAdmin -eq "No" -and $_.PasswordLastSet -lt $checkDate.adddays(-365))} |
        sort-object -Descending 'Number of Days Expired' -OutVariable Export
}

if ($saveToFile -and !$fileName) { $Export | export-CSV -Path $($defaultOutFileName + ".csv") -NoTypeInformation }
elseif ($saveToFile) { $Export | export-CSV -Path $fileName -NoTypeInformation }

write-output "`n`n** Finished **"

write-output "`n`n** Note: The 'IsAdmin' attribute in the result set is based on the 'adminCount' property of Get-Aduser cmdlet."
write-output "**       If flagged as admin, a user is listed above if the accound password is older than 90 days."
write-output "**       Non-admins are listed above if the account password is older than 365 days."
write-output "**       Only enabled accounts are checked. A blank output indicates all passwords are in compliance."

# References
# https://stackoverflow.com/questions/44151502/getting-the-no-of-days-difference-from-the-two-dates-in-powershell/44151764
# https://docs.microsoft.com/en-us/powershell/module/addsadministration/get-aduser?view=win10-ps
# https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/select-object?view=powershell-7
# https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/where-object?view=powershell-7
# https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/get-date?view=powershell-7