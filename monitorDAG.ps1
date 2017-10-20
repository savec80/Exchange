Function Get-ExchangeServerADSite ([String] $excServer) 
{ 
    # We could use WMI to check for the domain, but I think this method is better 
    # Get-WmiObject Win32_NTDomain -ComputerName $excServer 
 
    $configNC =([ADSI]"LDAP://RootDse").configurationNamingContext 
    $search = new-object DirectoryServices.DirectorySearcher([ADSI]"LDAP://$configNC") 
    $search.Filter = "(&(objectClass=msExchExchangeServer)(name=$excServer))" 
    $search.PageSize = 1000 
    [Void] $search.PropertiesToLoad.Add("msExchServerSite") 
 
    Try { 
        $adSite = [String] ($search.FindOne()).Properties.Item("msExchServerSite") 
        Return ($adSite.Split(",")[0]).Substring(3) 
    } Catch { 
        Return $null 
    } 
} 
 
 
 
[Bool] $bolFailover = $False 
[String] $errMessage = $null 
 
Get-MailboxDatabase | Sort Name | ForEach { 
    $db = $_.Name 
    $curServer = $_.Server.Name 
    $ownServer = $_.ActivationPreference | ? {$_.Value -eq 1} 
 
    # Compare the server where the DB is currently active to the server where it should be 
    If ($curServer -ne ($ownServer.Key).Name) 
    { 
        # Compare the AD sites of both servers 
        $siteCur = Get-ExchangeServerADSite $curServer 
        $siteOwn = Get-ExchangeServerADSite $ownServer.Key 
         
        If ($siteCur -ne $null -and $siteOwn -ne $null -and $siteCur -ne $siteOwn) { 
            $errMessage += "`n$db on $curServer should be on $($ownServer.Key) (DIFFERENT AD SITE: $siteCur)!"     
        } Else { 
            $errMessage += "`n$db on $curServer should be on $($ownServer.Key)!" 
        } 
 
        $bolFailover = $True 
    } 
} 
 
$errMessage += "`n`n" 
 
#Get-MailboxDatabase -Status | ? {$_.Recovery -eq $False -and $_.Mounted -eq $False} | Sort Name (...) 
Get-MailboxDatabase | Sort Name | Get-MailboxDatabaseCopyStatus | ForEach { 
    If ($_.Status -notmatch "Mounted" -and $_.Status -notmatch "Healthy" -or $_.ContentIndexState -notmatch "Healthy" -or $_.CopyQueueLength -gt 300 -or $_.ReplayQueueLength -gt 300) 
    { 
        $errMessage += "`n`n$($_.Name) - Status: $($_.Status) - Copy QL: $($_.CopyQueueLength) - Replay QL: $($_.ReplayQueueLength) - Index: $($_.ContentIndexState)" 
        $bolFailover = $True 
    } 
} 
 
If ($bolFailover) { 
    Schtasks.exe /Change /TN "MonitorDAG" /DISABLE 
     
    #$errMessage  
    Send-MailMessage -From "admin@letsexchange.com" -To "nuno.mota@letsexchange.com", "user2@letsexchange.com" -Subject "DAG NOT Healthy!" -Body $errMessage -Priority High -SMTPserver "smtp.letsexchange.com" -DeliveryNotificationOption onFailure 
}
