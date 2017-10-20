# Initialize variables 
[Int] $intTime = 15 
[Int] $intNumEmails = 25 
[Bool] $boolSendEmail = $False 
 
# Get the date and time for 15 minutes ago. This will be the starting point to search the transport logs 
[DateTime] $dateStartFrom = (Get-Date).AddMinutes(-$intTime) 
 
$hashEmails = @{} 
 
[String] $strEmailBody = "" 
$strEmailBody += "`n**********************************" 
$strEmailBody += "`n*                                *" 
$strEmailBody += "`n*    WARNING: Storm Detected!    *" 
$strEmailBody += "`n*                                *" 
$strEmailBody += "`n**********************************" 
 
 
# Check if we previously created any Transport Rule to prevent a storm. If yes, delete it 
Get-TransportRule | ? {$_.Name -match "Prevent Storm"} | Remove-TransportRule -Confirm:$False 
 
 
# Get all the users who sent 25 or more e-mails in the last 15 minutes 
# If you have multiple AD sites or are running in a co-existance scenario with multiple Exchange versions, you might want to include -and $_.Recipients -notmatch "letsexchange.com 
$logEntries = Get-TransportServer | Get-MessageTrackingLog -ResultSize Unlimited -Start $dateStartFrom -EventId SEND | ? {$_.Source -eq "SMTP" -and $_.Recipients -notmatch "letsexchange.com"} | Group Sender | ? {$_.Count -ge $intNumEmails} 
If ($logEntries -eq $null) { Exit } 
 
 
# For each sender, analyze all the e-mails they sent and put them in the HashTable 
ForEach ($logEntry in $logEntries) 
{ 
    ForEach ($email in (Get-TransportServer | Get-MessageTrackingLog -ResultSize Unlimited -Start $dateStartFrom -Sender $logEntry.Name -EventId SEND | ? {$_.Source -eq "SMTP" -and $_.Recipients -notmatch "letsexchange.com"} | Select MessageSubject, Sender, Recipients)) 
    { 
        $hashEmails["$($email.Sender) ô $($email.Recipients) ô $($email.MessageSubject)"] += 1 
    } 
} 
 
 
# To sort/filter a Hash Table, we need to transform its content into individual objects by using the GetEnumerator method 
# If we find any key with a value of at least 25, then we have detected a storm 
ForEach ($storm in ($hashEmails.GetEnumerator() | ? {$_.Value -ge $intNumEmails} | Sort Name)) 
{ 
    $boolSendEmail = $True 
    $arrDetails = [Regex]::Split($storm.Name, " ô ") 
 
    $strEmailBody += "`n`nSender:     ", $arrDetails[0] 
    $strEmailBody += "`nRecipient:  ", $arrDetails[1] 
    $strEmailBody += "`nSubject:    ", $arrDetails[2] 
    $strEmailBody += "`n# e-mails:  ", $storm.Value 
 
    # Get the local part of the e-mail address to use as part of the name for the transport rule 
    $strName = [regex]::split($arrDetails[0], "@")[0] 
 
    $ruleResult = New-TransportRule -Name "Prevent Storm - $strName" -Comments "Prevent Outlook AutoReply Storm" -From $arrDetails[0] -SentToScope "NotInOrganization" -RecipientAddressContainsWords $arrDetails[1] -SubjectContainsWords $arrDetails[2] -RedirectMessageTo "quarantine@letsexchange.com" -Enabled $True -Priority 0 
    If (!$ruleResult) 
    { 
        $strEmailBody += "`nRule Created? NO" 
    } Else { 
        $strEmailBody += "`nRule Created? Yes" 
    } 
     
#    # For every e-mail sent by that user 
#    $condition1 = Get-TransportRulePredicate From 
#    $condition1.Addresses = @(Get-Mailbox $arrDetails[0]) 
# 
#    # only when the e-mail is going Outside the organization 
#    $condition2 = Get-TransportRulePredicate SentToScope 
#    $condition2.Scope = @("NotInOrganization") 
# 
#    # only for e-mails that contain the subject we want 
#    $condition3 = Get-TransportRulePredicate SubjectContains 
#    $condition3.Words = @($arrDetails[2]) 
# 
#    # Redirect the e-mails e-mail to the Quarantine mailbox 
#    $action = Get-TransportRuleAction RedirectMessage 
#    $action.Addresses = @(Get-Mailbox Quarantine) 
# 
#    # Get the local part of the e-mail address to use as part of the name for the transport rule 
#    $strName = [regex]::split($email.Sender, "@")[0] 
# 
#    # Create the Transport Rule itself 
#    New-TransportRule -Name "Prevent Storm - $strName" -Comments "Prevent Outlook AutoReply Storm" -Conditions @($condition1, $condition2, $condition3) -Actions @($action) -Enabled $True -Priority 0 
} 
 
# Send an e-mail to the administrator(s) 
If ($boolSendEmail) { Send-MailMessage -From "AdminMotaN@letsexchange.com" -To "motan@letsexchange.com", "khanmr@letsexchange.com" -Subject "Storm Detected!" -Body $strEmailBody -SMTPserver "smtp.letsexchange.com" -DeliveryNotificationOption onFailure }
