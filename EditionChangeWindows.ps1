# Creates Windows Event Application Log Entry
# Source Name: PS Script
# Event Level: Information

$logName = "Change Windows Edition - PS Script";
$scriptName = "Change Windows Edition";
if($null -eq (Get-EventLog -LogName Application -Source $logName -ErrorAction SilentlyContinue)){
    New-EventLog -LogName Application -Source $logName;
    Write-EventLog -LogName Application -Source $logName -EventId 1 -Message "EventLog Created";
}
$logString = "";   
$logString += @"
Script Author: James Onley
Created Date: October 14, 2020
Running From: Intune
Scope: Campus 
Purpose: This script changes the Windows Edition to Windows 10 Education and activates with the campus MAK key. 

"@
$logString += "`n$scriptName started. `n`n"


# Check the current Windows Edition / Activation Status
$logString += "Checking Windows 10 activation status. `n"
$activationStatus = cscript "c:\Windows\System32\slmgr.vbs" /xpr

$isActivated = $false;
$isEnterpriseOrEducation = $false;
foreach($status in $activationStatus){
    if($status.contains("Education") -or $status.contains("Enterprise")){
        $isEnterpriseOrEducation = $true;
    }
    if($status.contains("activated")){
        $isActivated = $true;
    }
    if($isActivated -and $isEnterpriseOrEducation){
        $logString += "Windows 10 is already activated. `n"
        $logString += "`n$scriptName completed."

        # Send Log update to Windows Event Viewe
        Write-EventLog -LogName Application -Source $logName -EventId 1 -Message $logString
        exit;
    }
}

# If the Windows Edition is not Enterprise or Education then activate with MAK key
$logString += "Activating Windows with Education Edition license. `n"
$logString += "`nDevice may require restart to complete the Windows Activation."

$logString += "`n$scriptName completed."

# Send Log update to Windows Event Viewe
Write-EventLog -LogName Application -Source $logName -EventId 1 -Message $logString

cscript "c:\Windows\System32\slmgr.vbs" /ipk YKVKG-BN6QD-MFMGT-Y4CCG-9KXRQ