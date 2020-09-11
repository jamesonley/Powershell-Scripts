# Change to your AAD domain name
$domainName = "brockport.edu"; 

# Add Users
$users = ( 
    <# Add user to be added as local admins

       Example:
            "user",
            "user2"
    #>

    "jonley_sa@brockport.edu",
     "smaher_sa@brockport.edu",
    "tconner_sa@brockport.edu",
    "fmancini_sa@brockport.edu",
    "gelyjiw_sa@brockport.edu",
    "jperevic_sa@brockport.edu",
    "kgaylor_sa@brockport.edu",
    "pgattie_sa@brockport.edu",
    "shardenb_sa@brockport.edu",
    "aauriemm_sa@brockport.edu",
    "dmorriso_sa@brockport.edu",
    "kpijanow_sa@brockport.edu",
    "rcushman_sa@brockport.edu",
    "rheywood_sa@brockport.edu",
    "npalm4@brockport.edu",
    "tconn4_sa@brockport.edu",
    "kperlman_sa@brockportsuny.onmicrosoft.com",
    "kmontano_sa@brockportsuny.onmicrosoft.com",
    "mfook1_sa@brockportsuny.onmicrosoft.com",
    "sdarl1_sa@brockportsuny.onmicrosoft.com",
    "dle1_sa@brockportsuny.onmicrosoft.com",
    "pgard1_sa@brockportsuny.onmicrosoft.com",
    "bteslinger_sa@brockportsuny.onmicrosoft.com",
    "edoyl4_sa@brockportsuny.onmicrosoft.com"
    # "jlitt6_sa@brockportsuny.onmicrosoft.com",
    # "mdela3_sa@brockportsuny.onmicrosoft.com"
) 

# Remove Users
$removeUsers = ( 
    <# Remove users from local admins

       Example:
            "user",
            "user"
    #>
      ""
    
)


# Creates Windows Event Application Log Entry
# Source Name: PS Script
# Event Level: Information

$logName = "Add-Remove Admins - PS Script";
$scriptName = "Add-Remove Admins";
if($null -eq (Get-EventLog -LogName Application -Source $logName -ErrorAction SilentlyContinue)){
    New-EventLog -LogName Application -Source $logName;
    Write-EventLog -LogName Application -Source $logName -EventId 1 -Message "EventLog Created";
}
$logString = "";   
$logString += @"
Script Author: James Onley
Created Date: July 21, 2020
Running From: Intune
Scope: Campus 
Purpose: This script adds and removes administrators from the local administrator group on the device. It allows admins to elevate while assisting users without having to login first.

"@
$logString += "`n$scriptName started. `n`n"


# Gets current local admins from device
$existingAdmins = net localgroup administrators


# Start Adding Users to Local Admininstrator Group
if($null -ne $users -and $users -ne ""){
    foreach($user in $users){
        $name = $user.Split("@")[0];
        $domainUser = "WIN\" + $name;
        if(!$existingAdmins.Contains($domainUser)){
            $AAdUserToAdd = "azuread\$user"
            Write-Host $AAdUserToAdd;
            net localgroup administrators /add $AAdUsertoAdd
            $logString += "+++ $user added to local administrators group. `n"
        }
    }
}


# Start Removing Users from Local Admins Group
if($null -ne $removeUsers -and $removeUsers -ne ""){
    foreach($removeUser in $removeUsers){
        $name = $removeUser.Split("@")[0];
        $domainRemoveUser = "WIN\$name";
        if($existingAdmins.Contains($domainRemoveUser)){
            $AAdUserToRemove = "azuread\$removeUser"
            Write-Host $AAdUserToRemove;
            net localgroup administrators /delete $AAdUserToRemove
            $logString += "--- $removeUser removed from local administrators group. `n"
        }
    }
}

$logString += "`n$scriptName completed."

# Send Log update to Windows Event Viewe
Write-EventLog -LogName Application -Source $logName -EventId 1 -Message $logString