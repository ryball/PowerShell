#Break
#region do not touch !
$ProdRun = $false
#endregion do not touch !

#region Load Helper files
. ".\Config.ps1"
. ".\Credential.ps1"

#endregion Load Helper files

$LogFile = $LogFilePath + "Logfile-"+(Get-Date -Format "ddMMyyyy-HHmm")+".txt"
$ExceptionFile = "$ExceptionFile"

#region Loggfile setup
$TimeStamp = Get-Date -Format R
LogWrite -Logfile $LogFile -LogString "Script startet : $($TimeStamp)"
LogWrite -Logfile $LogFile -LogString "*****************************************************************************"
LogWrite -Logfile $LogFile -LogString "*         Script Created by : Raymond Mikalsen                              *"
LogWrite -Logfile $LogFile -LogString "*         Email : raymond.mikalsen@atea.no                                  *"
LogWrite -Logfile $LogFile -LogString "*                                                                           *"
LogWrite -Logfile $LogFile -LogString "* This script will assign, change or remove o365 lisence based on ADgroup.  *"
LogWrite -Logfile $LogFile -LogString "* Will assign each lisense spesific service plans based on AdGroup          *"
LogWrite -Logfile $LogFile -LogString "*****************************************************************************"
LogWrite -Logfile $LogFile -LogString "`n"
#endregion Loggfile setup

#region Setup
#Import Exception file if $Exceptionfile is not set to $null throw error if file not found. log to file.
if ($ExceptionFile)
{
    try
    {
        $Exceptions = Import-Csv $ExceptionFile
        Write-Verbose "Exception File Imported : $($ExceptionFile)"
    } catch
        {
            $ErrorMessage = $_.Exception.Message
            LogWrite -Logfile $LogFile -LogString $ErrorMessage
            Write-Error "Failed to import Exception file $($LogFile)"
        }
} else {
        Write-Verbose "Exception file not used"
}

#Connect to MsolService throw and error if it fails, log to file.
try
{
    Connect-MsolService -Credential (get-mycredential -User $MsolUserName -File $CredFile)
} catch 
    {
        $ErrorMessage = $_.Exception.Message
        LogWrite -Logfile $LogFile -LogString $ErrorMessage
        Write-Error "Failed to connect to Microsoft Online check $($LogFile)"
        Break
    }

#Check if module is loaded, if not throw error, log to file
if (!(Get-Module MsOnline))
{
    try 
    {
        Import-Module MsOnline
    } catch 
        {
          $ErrorMessage = $_.Exception.Message
          LogWrite -Logfile $LogFile -LogString $ErrorMessage
          Write-Error "Failed to load module check $($LogFile)"  
          Break
        }
}

if (!(Get-Module ActiveDirectory))
{
    try
    {
        Import-Module ActiveDirectory
    } catch
        {
            $ErrorMessage = $_.Exception.Message
            LogWrite -Logfile $LogFile -LogString $ErrorMessage
            Write-Error "Failed to load module check $($LogFile)"
            Break
        }
}

#endregion Setup

#region Get users from O365 - Remove users in Exception List
Write-Verbose "Grabbing users from O365"

$MsolUsers = New-Object System.Collections.ArrayList
[System.Collections.ArrayList]$MsolUsers = Get-MsolUser -All 

$TimeStamp = Get-Date -Format R
LogWrite -Logfile $LogFile -LogString $TimeStamp
LogWrite -Logfile $LogFile -LogString "$($MsolUsers.count) users grabbed and stored in an Arraylist"
LogWrite -Logfile $LogFile -LogString "`n"
Write-Verbose "$($MsolUsers.count) users grabbed and stored in an Arraylist..."

#Remove users spesified in the exception file
if ($ExceptionFile)
{
    LogWrite -Logfile $LogFile -LogString "Exception List:`r`n"
    #Remove users that is spesified in the Exception List
    foreach ($User in $Exceptions)
    {
        try 
        {
            $MsolUsers.RemoveAt($MsolUsers.UserPrincipalName.IndexOf($User.UserPrincipalName)) 
            LogWrite -Logfile $LogFile -LogString "$($User.UserPrincipalName) removed from array, exists in Exceptionlist $($ExceptionFile)"
            Write-Verbose "$($User.UserPrincipalName) removed from array"
        } catch
            {
                LogWrite -Logfile $LogFile -LogString "$($User.UserPrincipalName) is not a MsolUser but can still be part of an AD group."
                Write-Warning "$($User.UserPrincipalName) is not a MsolUser but can still be part of an AD group."
            }
    }
    LogWrite -Logfile $LogFile -LogString "`n"
}

#endregion Get users from O365 - Remove users in Exception List

#region Get all licensed users and put them in a custom object
LogWrite -Logfile $LogFile -LogString "Grabbing license details for each users in array"
Write-Verbose "Grabbing license details for each users in array..."

$LicensedUserDetails = $MsolUsers | Where-Object {$_.IsLicensed -eq $true} | ForEach-Object {
    [PsCustomObject]@{
                        UserPrincipalName = $_.UserPrincipalName
                        License = $_.Licenses.AccountSkuId
                        } 
}

LogWrite -Logfile $LogFile -LogString ($($LicensedUserDetails) | Out-String)

#endregion Get all licensed users and put them in a custom object
 
#region Get Userdefined AccountSku
LogWrite -Logfile $LogFile -LogString "Grabbing Userdefined Account Sku from : $($UserSkuFile)"

try 
{
    $UserSku = Import-Csv -Delimiter ";" -Path $UserSkuFile
    Write-Verbose "Userdefined Sku Imported from : $($UserSkuFile)"
} catch
    {
            $ErrorMessage = $_.Exception.Message
            LogWrite -Logfile $LogFile -LogString $ErrorMessage
            Write-Error "Failed to import Exception file $($LogFile)"
            Break
    }

$hashtable = @{}
$hashtable.Clear()

$Licenses =  $UserSku | ForEach-Object{
        $key = $_.Key
        $LicenseSku = $_.LicenseSku
        $Adgroup = $_.AdGroup
        $Plans = $_.Plans

        $hashtable = @{
            Key = $key
            LicenseSku = $LicenseSku
            AdGroup = $Adgroup
            EnabledPlans = $Plans
        }

        New-Object -TypeName psobject -Property $hashtable
    } | Select-Object Key, LicenseSku, AdGroup, EnabledPlans

LogWrite -Logfile $LogFile -LogString ($($Licenses) | Out-String)
#endregion Get Userdefined AccountSku
                
#region Clear out som Variabels before we start
#Array for users to change or remove
$UsersToChangeorRemove = @()
$UsersToChangeorRemove.Clear()
$License = $null
[int]$addedCount = 0
[int]$RemoveCount = 0
[int]$ChangeCount = 0
$dirsyncIssue = @()
$couldnotlicense = @() # Array of all users that threw Exception when licensing
#endregion Clear out som Variabels before we start

#region This is where the Magic Happens
foreach ($License in $Licenses.Key)
{
    $GroupName = ($Licenses[$Licenses.key.IndexOf($License)]).adgroup
    $GroupID = (Get-MsolGroup -All | Where-Object {$_.DisplayName -eq $GroupName}).ObjectId
    $AccountSKU = Get-MsolAccountSku | Where-Object {$_.AccountSKUID -eq $Licenses[$Licenses.key.IndexOf($License)].LicenseSKU}
    
    Write-Verbose "Checking for unlicensed $License users in group $GroupName...." 
    LogWrite -Logfile $LogFile -LogString (get-date -Format R)
    LogWrite -Logfile $LogFile -LogString "Checking for unlicensed $License users in group $GroupName"

    #Disable plans
    $EnabledPlans = $Licenses[$Licenses.key.IndexOf($License)].EnabledPlans
    if ($EnabledPlans)
    {
        $LicenseOptionsHt = @{
            AccountSkuId = $AccountSKU.AccountSkuId
            DisabledPlans = (Compare-Object -ReferenceObject $AccountSKU.ServiceStatus.ServicePlan.ServiceName -DifferenceObject $EnabledPlans).InputObject
            }
            $LicenseOptions = New-MsolLicenseOptions @LicenseOptionsHt
        LogWrite -Logfile $LogFile "Service plans for this sku : $EnabledPlans `r`n"
    }

    #All Members of the group 
    Write-Verbose "Grabbing users from adgroup : $($GroupName)"
    [System.Collections.ArrayList]$AdGroupMembers = Get-AdGroupMember -Identity $GroupName | Get-Aduser -Properties EmailAddress, TargetAddress, UserPrincipalName
    Write-Verbose "$($AdGroupMembers.count) members found in $($GroupName)"
    LogWrite -Logfile $LogFile -LogString "$($AdGroupMembers.count) members found in Adgroup: $($GroupName)"

    foreach ($AdUser in $Exceptions)
    {
        try
        {
            $AdGroupMembers.RemoveAt($AdGroupMembers.UserPrincipalName.IndexOf($AdUser.UserPrincipalName))
            Write-Verbose "User found in Exception list : $($AdUser.UserPrincipalName) will not be processed."
            LogWrite -Logfile $LogFile -LogString "User found in Exception list : $($AdUser.UserPrincipalName) will not be processed. `r`n"
        } catch
            {
               # Do Nothing
            }
        $AddUser = $null
    }

    write-verbose "Get all users in $($GroupName) that has $($TargetAddress) set"
    $MembersWithExchange = $AdGroupMembers | Where-Object {($_.EmailAddress -like "*") -and ($_.targetaddress -like "$TargetAddress")}
    Write-Verbose "$($MembersWithExchange.Count) with O365 mailbox"
    LogWrite -Logfile $LogFile -LogString "$($MembersWithExchange.Count) with Targetaddress : $($TargetAddress) `r`n"

    #All Members that allready have license assigned
    $ActiveUsers = ($LicensedUserDetails | Where-Object {$_.License -eq $Licenses[$Licenses.key.IndexOf($License)].LicenseSKU}).UserPrincipalName
    
    $UsersToChange = $null
    $UsersToChangeorRemove = $null
    $UsersToAdd = $null
    
    if ($AdGroupMembers)
    {
        if ($ActiveUsers)
        {
            #Compare $GroupMembers and $ActiveUsers
            #Users in the group but not licensed, will be added
            #Licensed users not in group will be evaluated for deletion or change of license
            Write-Verbose "Checking what to do with the user : Change, Add, Remove?"
            $UsersToChange = Compare-Object -ReferenceObject $AdGroupMembers.UserPrincipalName -DifferenceObject $ActiveUsers -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            $UsersToAdd = ($UsersToChange | Where-Object {($_.SideIndicator -eq '<=')}).InputObject
            $UsersToChangeorRemove += ($UsersToChange | Where-Object {($_.SideIndicator -eq '=>')} ).InputObject
        } else {
            #Licenses currently not assigned for the license in scope, assign licenses to all group members.
            $UsersToAdd = $AdGroupMembers.UserPrincipalName
        }
    } else {
        Write-Warning "Group $GroupName is Empty - Will process remove or move of all users with license $($AccountSKU.AccountSkuId)"
        #If no users are a member in the group, add them for deletion or change of licenses.
        $UsersToChangeorRemove += $ActiveUsers
    }
    LogWrite -Logfile $LogFile -LogString "Users to change in group : $($GroupName)"
    LogWrite -Logfile $LogFile -LogString ($UsersToChange | Out-String)


    if ($UsersToAdd)
    {
        $UsersToAdd = $UsersToAdd -split " "
     
        LogWrite -Logfile $LogFile -LogString "Users that will be assigned $($License) in Group $($GroupName):"
        LogWrite -Logfile $LogFile -LogString ($UsersToAdd | Out-String)
     
        foreach ($AddUser in $UsersToAdd)
        {
            Write-Verbose "Processing ADuser : $($AddUser)"
            LogWrite -Logfile $LogFile -LogString "Processing Aduser : $($AddUser)"
            
            if ($AddUser) 
            {
                
                try 
                {
                    $Msoluser = Get-MsolUser -UserPrincipalName $AddUser -ErrorAction Stop -WarningAction Stop
                } catch {
                    $ErrorMessage = $_.Exception.Message
                    $Usernotfound = "User not Found in Office 365. User : $($AddUser)`n"
                    Write-Warning $Usernotfound
                    LogWrite -Logfile $LogFile -LogString $Usernotfound    
                    $dirsyncIssue += @($AddUser, $Error[0])
                }

                #Assign License, if not already licensed with the SKU.  
                if ($Msoluser.Licenses.AccountSkuID -notcontains $AccountSKU.AccountSkuId)
                {
                    $Error.Clear()
                    try 
                    {
                        #Location and License and Options.
                        if($ProdRun)
                        {
#                            Set-MsolUser -UserPrincipalName $AddUser -UsageLocation $UsageLocation -ErrorAction Stop -WarningAction Stop
#                            Set-MsolUserLicense -UserPrincipalName $AddUser -AddLicenses $AccountSKU.AccountSkuId -LicenseOptions $LicenseOptions -ErrorAction Stop -WarningAction Stop
                                $addedCount = $addedCount+1
                        } else {
                            Write-Warning "This is a test run no changes will be made to $($AddUser)" 
                            LogWrite -Logfile $LogFile -LogString "This is a test run no changes will be made to $($AddUser)" 
                            $addedCount = $addedCount+1
                        }

                        Write-Verbose "SUCCESS: Licensed $($AddUser) with $License"
                        LogWrite -Logfile $LogFile -LogString "SUCCESS: Licensed $($AddUser) with $License"
                    } 
                    catch
                    {
                        Write-Warning "Error when licensing $($AddUser)"
                        $ErrorMessage = $_.Exception.Message
                        LogWrite -Logfile $LogFile -LogString "Error when licensing $($AddUser)"
                        LogWrite -Logfile $LogFile -LogString $ErrorMessage
                        $couldnotlicense += @($AddUser, $Error[0])
                    }
                }
            }
        }
    }

    #Evaluate Users to change or Delete
    $UsersToChangeorRemove = $UsersToChangeorRemove -split " "

    LogWrite -Logfile $LogFile -LogString "`r`n"
    LogWrite -Logfile $LogFile -LogString "Users that will be changed or removed in Group $($GroupName):"
    LogWrite -Logfile $LogFile -LogString ($UsersToChangeorRemove | Out-String)

    if (($UsersToChangeorRemove -ne $null) -or ($UsersToChangeorRemove -notlike ""))
    {
        foreach ($User in $UsersToChangeorRemove)
        {
            Write-Verbose "Processing user : $($User)"
            LogWrite -Logfile $LogFile -LogString "Processing user : $($User)"

            if ($User -ne $null)
            {
                #Get users old license

                $OldLic = ($LicensedUserDetails | Where-Object {$_.UserPrincipalName -eq $User}).License

                #Loop to check if the user group assignment has been changed, Custom object for new license
                $ChangeLicense = $Licenses.Keys | ForEach-Object {
                                                                    $GroupName = $Licenses[$Licenses.key.IndexOf($_)].AdGroup
                                                                    if (Get-ADGroupMember -Identity $GroupName | Get-ADUser | Where-Object {$_.UserPrincipalName -eq $User})   
                                                                    {
                                                                        [pscustomobject]@{
                                                                                        NewLicense = $Licenses[$_].LicenseSKU
                                                                                        Options = $Licenses[$_].EnabledPlans
                                                                        }
                                                                    }  
                                                                  }
    
                try
                {
                    $compareLic = Compare-Object -ReferenceObject $OldLic -DifferenceObject $ChangeLicense.NewLicense 
                    $LicenseToRemove = $compareLic | Where-Object {($_.SideIndicator -eq '<=')}
                } catch
                    {
                        #Do Nothing
                    }
                
                $ChangeRange = New-Object System.Collections.ArrayList($null)
                $ChangeRange.Add($ChangeLicense)

                if ($ChangeLicense)
                { 
                    $allreadyEnabled = (Get-MsolUser -UserPrincipalName $User).Licenses.accountskuid
                    
                    if ($allreadyEnabled)
                    {
                        $CompareChange = Compare-Object -ReferenceObject $AllreadyEnabled -DifferenceObject $ChangeLicense.Newlicense -IncludeEqual   
                        $EqualLicense = $CompareChange | Where-Object {$_.SideIndicator -eq '=='}
                    }
       
                                                                              
                    if ($EqualLicense)
                    {
                        $count =  $EqualLicense.InputObject.count   

                        if ($count -gt 1)
                        {
                            for ($i = 0; $i -lt $count; $i++)
                            { 
                               $ChangeRange.RemoveAt($ChangeRange.newlicense.IndexOf($EqualLicense[$i].InputObject))
                            }
                        } else 
                            {
                                $ChangeRange.RemoveAt($ChangeRange.newlicense.IndexOf($EqualLicense.InputObject))
                            }
                    } 

                    foreach ($lic in $LicenseToRemove)
                    {
                        try
                        {
                            if ($ProdRun)
                            {                  
#                             Set-MsolUserLicense -UserPrincipalName $User -RemoveLicenses $lic.InputObject -ErrorAction Stop -WarningAction Stop
                              $RemoveCount = $RemoveCount+1
                            } else {
                                Write-Warning "This is a test run no License will be removed from : $($User.UserPrincipalName)"
                            }
                                Write-Verbose "Removing license $($lic.InputObject) from $($User.UserPrincipalName)"
                                LogWrite -Logfile $LogFile -LogString "Removing license $($lic.InputObject) from $($User.UserPrincipalName)"
                            
                        }
                        catch
                        {
                            #Write-warning "$($OldLic) Lisense allready removed from $($User)"
                        }
                    } 

                    for ($i = 0; $i -lt $ChangeRange.NewLicense.Count; $i++)
                    { 
                        
                        $Error.Clear()
                        try
                        {
                            $SKU = Get-MsolAccountSku | Where-Object {$_.AccountSKUID -eq $ChangeRange[$i].NewLicense}
                            $Options = (Compare-Object -ReferenceObject $SKU.ServiceStatus.ServicePlan.ServiceName -DifferenceObject $ChangeRange[$i].Options).InputObject
                            $LicenseOptions = New-MsolLicenseOptions -AccountSkuId $ChangeRange[$i].NewLicense -DisabledPlans $Options -ErrorAction Stop -WarningAction Stop
                            
                            if ($ProdRun)
                            {
#                                Set-MsolUserLicense -UserPrincipalName $User -AddLicenses $ChangeRange[$i].NewLicense -LicenseOptions $LicenseOptions -ErrorAction Stop -WarningAction Stop
                                 $ChangeCount = $ChangeCount+1
                            } else {
                                  Write-Warning "This is a test run no License will be changed on : $($User.UserPrincipalName)"
                                  LogWrite -Logfile $LogFile -LogString "This is a test run no License will be changed on : $($User.UserPrincipalName)"
                                  $ChangeCount = $ChangeCount+1
                            }
                            Write-Verbose "SUCCESS: Changed license for users $User from $($lic.InputObject) to $($ChangeRange[$i].NewLicense)" 
                            LogWrite -Logfile $LogFile -LogString "SUCCESS: Changed license for users $User from $($lic.InputObject) to $($ChangeRange[$i].NewLicense)"
                        }
                        catch
                        {
                            Write-Warning "Error when changing license on $User`r`n$_"
                            LogWrite -Logfile $LogFile -LogString "Error when changing license on $User`r`n$_"
                            $couldnotlicense += @($User, $Error[0])
                        }
                    }

                } else {
                    #User is no longer a member of any license group, remove license
                    Write-Warning "$User is not a member of any group, license will be removed"
                    LogWrite -Logfile $LogFile -LogString "$User is not a member of any group, license will be removed"
                    try
                    {
                        if ($ProdRun)
                        {
#                            Set-MsolUserLicense -UserPrincipalName $User -RemoveLicenses $OldLic -ErrorAction Stop -WarningAction Stop
                             $RemoveCount = $RemoveCount+1
                        } else 
                        {
                             Write-Warning "This is a test run no License will be removed from : $($User)"
                             LogWrite -Logfile $LogFile -LogString "This is a test run no License will be removed from : $($User)"
                             $RemoveCount = $RemoveCount+1
                        }
                        Write-Verbose "SUCCESS: Removed $OldLic for $User" 
                        LogWrite -Logfile $LogFile -LogString "SUCCESS: Removed $OldLic for $User"
                    }
                    catch
                    {
                        $WarningString = "Error when removing license on user'r'n$_"                           
                        Write-Warning $WarningString
                        LogWrite -Logfile $LogFile -LogString $WarningString
                        $RemoveError += $WarningString
                    }
                } 
            }
            LogWrite -Logfile $LogFile -LogString "`r`n"
       }
    }

    if ($MembersWithExchange)
    {
        $MembersWithExchange = $MembersWithExchange -split " "

        foreach ($usr in $MembersWithExchange.UserPrincipalName)
        {

                $MsoUsr = Get-MsolUser -UserPrincipalName $usr
                # Check if EXCHANGE_S_STANDARD should be activated on the account
                $i = 0
                foreach ($item in $MsoUsr.Licenses.ServiceStatus.ServicePlan.ServiceName)
                {
                    if (($item -like "EXCHANGE_S_STANDARD") -or ($item -like "EXCHANGE_S_ENTERPRISE") )
                    {
                        break    
                    }
                    $i++
                }

                if ($MsoUsr.Licenses.ServiceStatus[$i].ProvisioningStatus -like "Disabled") 
                {
                    $EnablePlansWithExchange = $EnabledPlans  
                    $EnablePlansWithExchange += $MsoUsr.Licenses.ServiceStatus.ServicePlan.ServiceName[$i]
                        
                    $LicenseOptionsWithExchangeHt = @{
                            AccountSkuId = $AccountSKU.AccountSkuId
                            DisabledPlans = (Compare-Object -ReferenceObject $AccountSKU.ServiceStatus.ServicePlan.ServiceName -DifferenceObject $EnablePlansWithExchange).InputObject
                    }
                    $LicenseOptionsWithExchange = New-MsolLicenseOptions @LicenseOptionsWithExchangeHt

                    $OldLic = ($LicensedUserDetails | Where-Object {$_.UserPrincipalName -eq $usr}).License
                    try 
                    {
                        #Location and License and Options.

                        if ($ProdRun)
                        {
#                            Set-MsolUser -UserPrincipalName $usr -UsageLocation $UsageLocation -ErrorAction Stop -WarningAction Stop
#                            Set-MsolUserLicense -UserPrincipalName $usr -RemoveLicenses $OldLic 
#                            Set-MsolUserLicense -UserPrincipalName $usr -AddLicenses $AccountSKU.AccountSkuId -LicenseOptions $LicenseOptionsWithExchange -ErrorAction Stop -WarningAction Stop
                        } else
                        {
                             Write-Warning "This is a test run, no exchange lisense will be added to : $($User.UserPrincipalName)"
                             LogWrite -Logfile $LogFile -LogString "This is a test run, no exchange lisense will be added to : $($User.UserPrincipalName)"
                        }                   
                        Write-Verbose "SUCCESS: Licensed $usr with $License , Enabling $($MsoUsr.Licenses.ServiceStatus.ServicePlan.ServiceName[$i])" 
                        LogWrite -Logfile $LogFile -LogString "SUCCESS: Licensed $usr with $License , Enabling $($MsoUsr.Licenses.ServiceStatus.ServicePlan.ServiceName[$i])"
                    } 
                    catch
                    {
                        Write-Warning "Error when setting new license with $($MsoUsr.Licenses.ServiceStatus.ServicePlan.ServiceName[$i]) Enabled for : $usr"
                        LogWrite -Logfile $LogFile -LogString "Error when setting new license with $($MsoUsr.Licenses.ServiceStatus.ServicePlan.ServiceName[$i]) Enabled for : $usr"
                    }
                
                } 
             }
        }
    LogWrite -Logfile $LogFile -LogString "***************************************************************************** `r`n"

    #region Clearing som Variabels 
    $GroupName = $null
    $GroupID = $null
    $AccountSKU = $null
    $EnabledPlans = $null  
    $LicenseOptions = $null
    $AdGroupMembers = New-Object System.Collections.ArrayList
    $AdGroupMembers.Clear()
    $MembersWithExchange = $null
    $ActiveUsers = $null

    $UsersToChange = $null
    $UsersToChangeorRemove = $null
    $ChangeLicense = $null
    $compareLic = $null
    $LicenseToRemove = $null
    $ChangeRange.Clear()
    $allreadyEnabled = $null
    $CompareChange = $null
    $EqualLicense = $null
    $lic = $null
    $Options = $null
    $usr = $null
    $MsoUsr = $null
    $EnablePlansWithExchange = $null
    $LicenseOptionsWithExchange = $null
    $OldLic = $null
    #endregion Clearing som Variabels
 }
#endregion This is where the Magic Happens


if($usereport)
{
    . ".\repport.ps1" -Path $ReportPath
}