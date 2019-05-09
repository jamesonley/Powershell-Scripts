# Title: Search for Bad Rules
# Author: James Onley
# Purpose: To search Exchange Online for malicious inbox rules. 
# Date: May 7, 2019 5:20pm

# Note: Connections to Exchange Online are metered and sometimes fail. For this reason the script stores 
# data from big queries in CSV files. This prevents the query from having to run again if the connection 
# fails. To get new data simply delete the mail_users and rules csv files. 

$script_location = split-path -parent $MyInvocation.MyCommand.Definition
if($script_location -ne $null){
    $script_location = $script_location + "\"
}
else{
    $script_location = "C:\scripts\get_bad_inbox_rules\"
}
$mail_users_csv = "mail_users.csv"
$mail_rules_csv = "rules.csv"




$mail_users_file = $script_location + $mail_users_csv
$mail_rules_file = $script_location + $mail_rules_csv
cd $script_location
$global:log_file = "$script_location\logs\bad_rules-" + $(get-date -f yyyy-MM-dd-hhmm) + ".log"
$global:show_logs = $true

if([System.IO.File]::Exists($log_file) -eq $false){
    md "$script_location\logs"
}

function log_action{
	Param(
		[string]$event
	)
	$timestamp = Get-Date -Format g
	$message = $timestamp + " - " + $event
	if($show_logs -eq $true){
		Add-Content -Path $log_file -Value $message -force  
		write-host $message
	}
	else{
		Add-Content -Path $log_file -Value $message -force -erroraction silentlycontinue
	}
}

function show_action{
	Param(
		[string]$event
	)
	$timestamp = Get-Date -Format g
	$message = $timestamp + " - " + $event
	write-host $message
}

cls
log_action "Starting Get Bad Inbox Rules Script"
$script_location

# Connect to Exchange and Office 365

$UserCredential = Get-Credential

try{
    $EmailSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
    log_action "Connected to Exchange Online."
}
catch{
    log_action "Connection to Exchange Online couldn't be made. Ending script."
    PAUSE
    break
}

Import-PSSession $EmailSession –DisableNameChecking 





# Get mail users if a mail_users.csv doesn't exist
if([System.IO.File]::Exists($mail_users_file) -eq $false){
    log_action ""
    log_action "Connected to Exchange Online."

    # Run the Get-mailbox command on Microsoft's server and return the result. Works faster. 
    log_action "Loading mail user list from Exchange Online"
    try{
        Invoke-Command -Session $EmailSession -ScriptBlock {Get-Mailbox -ResultSize unlimited | select-object -Property DisplayName, Identity, PrimarySMTPAddress, isMailboxEnabled,accountDisabled, recipientType,RecipientTypeDetails   } | sort-object -property PrimarySMTPAddress |  Export-CSV $mail_users_file
        log_action "Success. Mailbox list complete."
    }
    catch{
        log_action "Failure. Getting mailbox list from Exchange Online failed. Script ending."
        PAUSE
        break
    }
}
try{
    # Load the mail users from the CSV file. 
    log_action "Loading mail user list from file."
    $mail_users = Import-csv $mail_users_file 
    log_action "Success. Mail user list loaded."
}
catch{
    log_action "Failure. Mail user not list loaded. Script ending."
    PAUSE
    break
}

$mail_users = $mail_users |  where-object {$_.isMailboxEnabled -eq $true -and $_.accountDisabled -eq $false -and $_.recipientType -eq "UserMailbox" -and $_.RecipientTypeDetails -eq "UserMailbox"} | Sort-Object -Property PrimarySMTPAddress
$i = 1
$index_last_user = 0
$rule_Objs = @()

# if the mail rules csv exists import it. See where the script left off on the user list.
if([System.IO.File]::Exists($mail_rules_file) -eq $true){
    $rule_Objs = Import-csv $mail_rules_file
    if($rule_Objs -ne $null){
        $index_from_file = $rule_Objs.length -1
        $last_processed_user = $rule_Objs[$index_from_file].primarysmtpaddress
        $index_last_user = 0
        foreach( $mail_user in $mail_users){
       
            $last_processed_user.PrimarySmtpAddress
            if($mail_user.PrimarySmtpAddress -eq $last_processed_user){
                break
            }
            else{
                $index_last_user++       
            }
        }
        $i = $index_last_user
    }
}

$total = $mail_users.count
log_action "Total mail users to process: $total"

# Process mail users. Skip past any users loaded from existin mail rule csv file
foreach ($mail_user in ($mail_users | select -skip $index_last_user)){
    
    # Show progress bar
    $percent_complete = [int](($i / $total) * 100)
    Write-Progress -Activity "Processing Accounts - Slow Progress Bar" -Status "$percent_complete% Complete - $i / $total" -PercentComplete $percent_complete;
    $i++

    $rule = $null
    $rule_obj = $null
    $inbox_rules = $null
    $user_address = $null

    $user_address = $mail_user.PrimarySMTPAddress
    $inbox_rules = Get-InboxRule -Mailbox $user_address 
   
 
    # Process all mail rules associated with a user. Builds and object to load into the csv file.
    foreach ($rule in $inbox_rules){
        $rule_description = $rule.Description
        $rule_identity = $rule.RuleIdentity
        $rule_delete_message = $rule.DeleteMessage
        $rule_forward_attachment = $rule.ForwardAsAttachmentTo
        $rule_move_to_folder = $rule.MoveToFolder
        $rule_redirect_to = $rule.RedirectTo
        $rule_enabled = $rule.Enabled
                
        $rule_obj = New-Object System.Object
        $rule_Obj | Add-Member -MemberType NoteProperty -Name "PrimarySMTPAddress" -Value $user_address
        $rule_Obj | Add-Member -MemberType NoteProperty -Name "Description" -Value $rule_description
        $rule_Obj | Add-Member -MemberType NoteProperty -Name "RuleIdentity" -Value $rule_identity
        $rule_Obj | Add-Member -MemberType NoteProperty -Name "DeleteMessage" -Value $rule_delete_message
        $rule_Obj | Add-Member -MemberType NoteProperty -Name "ForwardAsAttachmentTo" -Value $rule_forward_attachment
        $rule_Obj | Add-Member -MemberType NoteProperty -Name "MoveToFolder" -Value $rule_move_to_folder
        $rule_Obj | Add-Member -MemberType NoteProperty -Name "RedirectTo" -Value $rule_redirect_to
        $rule_Obj | Add-Member -MemberType NoteProperty -Name "Enabled" -Value $rule_enabled
        log_action "Rule $rule_description discovered." 
                                 
        $rule_Objs += $rule_obj
        $rule = $null
        $rule_obj = $null
    }

    # Periodically save the CSV file in case the script crashes or the server connection is lost
    if($i % 100 -eq 0){
        $rule_Objs | export-csv $mail_rules_file
        log_action "Saved mail rules to CSV file"
    }
    
}
$rule_Objs | export-csv $mail_rules_file

try{
    log_action "Loading mail rules from file."
    $rule_Objs = Import-csv $mail_rules_file
}
catch{
    log_action "Failed. Couldn't load mail rules from file."
    PAUSE
    break
}

# Queries for the existing data. Once the mail rules CSV is loaded, you can quickly run different queries against the $rule_objs such as the one below. 
    #$rule_objs | where-object {$_.description -match "phishing" -or $_.description -match "helpdesk" -or $_.description -match "attack" -or $_.description -match "scam" -or $_.description -match "password" -or $_.description -match "spam" -and $_.enabled -eq $true} | select primarysmtpaddress, enabled, description | ft -wrap
$rule_objs | out-gridview

remove-PSSession $EmailSession

PAUSE
