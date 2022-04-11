
# A script to check the activity of Microsoft 365 Groups and Teams and report the groups and teams that might be deleted because they're not used.
# We check the group mailbox to see what the last time a conversation item was added to the Inbox folder. 
# Another check sees whether a low number of items exist in the mailbox, which would show that it's not being used.
# We also check the group document library in SharePoint Online to see whether it exists or has been used in the last 90 days.
# And we check Teams compliance items to figure out if any chatting is happening.

Function Get-SavedCredential([string]$KeyPath)
{
   If (Test-Path $KeyPath) 
   {
      $UserName = Get-Content "$($KeyPath)\UserName.txt"
      $SecureString = Get-Content "$($KeyPath)\securePassword.cred" | ConvertTo-SecureString
      return $UserName, $SecureString
   }
   Else
   {
      $Credential = Get-Credential -Message "Enter the Credentials:"
      if ($Null -eq $Credential.Password -Or "" -eq $Credential.password){
         Exit
      }
      New-Item -ItemType Directory -Path $KeyPath -ErrorAction STOP | Out-Null
      New-Item -Path $KeyPath -Name "UserName.txt" -ItemType "file" -Value $Credential.UserName -ErrorAction STOP | Out-Null
      $Credential.Password | ConvertFrom-SecureString | Out-File "$($KeyPath)\securePassword.cred" -Force
      Get-SavedCredential -KeyPath $KeyPath
   }
}

$UserName, $SecureString = Get-SavedCredential  -KeyPath "Credentials"

#Set-Location -Path "C:\Users\codem\Documents\Code\OIT DEV"
#.\GroupActivityReport.ps1
./ConnectO365Services.ps1 -Services AzureAD, ExchangeOnline, SharePoint, Teams -SharePointHostName wmutest1 -UserName $UserName -Password $SecureString

#z:\ConnectO365Services.ps1 -Services AzureAD, ExchangeOnline, SharePoint, Teams -SharePointHostName wmich -UserName $UserName -Password $SecureString

# CLS #Clears text printed above
 
$OrgName = (Get-OrganizationConfig).Name  
       
# OK, we seem to be fully connected to both Exchange Online and SharePoint Online...

Write-Host "Checking Microsoft 365 Groups and Teams in the tenant:" $OrgName

# Setup some stuff we use
$WarningDate = (Get-Date).AddDays(-90);
$WarningEmailDate = (Get-Date).AddDays(-365);
$Today = (Get-Date);
$Date = $Today.ToShortDateString()

$TeamsGroups = 0;  $TeamsEnabled = $False; $ObsoleteSPOGroups = 0; $ObsoleteEmailGroups = 0

$startTime = Get-Date
$startTimeISO = Get-Date($StartTime) -format "yyyyMMdd HHmmss"

$TextFileName = "C:\temp\log\" + $startTimeISO + "_log.text"
Start-Transcript -Path $TextFileName

$Report = [System.Collections.Generic.List[Object]]::new()
$ReportFile = "C:\temp\html\" + $startTimeISO + "_GroupsActivityReport.html"
$CSVFile = "C:\temp\csv\" + $startTimeISO + "_GroupsActivityReport.csv"

$htmlhead="<html>
	   <style>
	   BODY{font-family: Arial; font-size: 8pt;}
	   H1{font-size: 22px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H2{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
	   th{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
	   td{border: 1px solid #969595; padding: 5px; }
	   td.pass{background: #B7EB83;}
	   td.warn{background: #FFF275;}
	   td.fail{background: #FF2626; color: #ffffff;}
	   td.info{background: #85D4FF;}
	   </style>
	   <body>
           <div align=center>
           <p><h1>Microsoft 365 Groups and Teams Activity Report</h1></p>
           <p><h3>Generated: " + $date + "</h3></p></div>"


Write-Host "`n Script Start Time: $startTimeISO `n"

$GroupCheckStartTime = Get-Date
Write-Host "Start fetching Microsoft 365 Groups for checking at $($GroupCheckStartTime.ToString("yyyy-MM-dd HH:mm:ss"))"

[Int]$GroupsCount = 0;
[int]$TeamsCount = 0;
$TeamsList = @{};
$UsedGroups = $False

if (Test-Path "C:\Temp\Input.txt"){
   $Content = Get-Content -Path C:\Temp\Input.txt 
}
else{
   $Groups = $NULL
}

if($Groups -eq $NULL){
   # Get a list of Groups in the tenant
   #Start time for Groups Fetch

   $Groups = Get-Recipient -RecipientTypeDetails GroupMailbox -ResultSize Unlimited | Sort-Object DisplayName

   $GroupsCount = $Groups.Count

   # If we don't find any groups (possible with Get-Recipient on a bad day), try to find them with Get-UnifiedGroup before giving up.
   If ($GroupsCount -eq 0) { # 
      Write-Host "Fetching Groups using Get-UnifiedGroup"
      $Groups = Get-UnifiedGroup -ResultSize Unlimited | Sort-Object DisplayName 
      $GroupsCount = $Groups.Count; $UsedGroups = $True

      If ($GroupsCount -eq 0) {
      Write-Host "No Microsoft 365 Groups found; script exiting" ; break} 
   } # End If

   #End time for Groups Fetch
   $GroupCheckEndTime = Get-Date
   Write-Host "`n Group Fetch Time:   $($GroupCheckEndTime.ToString("yyyy-MM-dd HH:mm:ss")) `n"
   $GroupRunTime = New-TimeSpan -End $GroupCheckEndTime -Start $GroupCheckStartTime
   Write-Host "Group Fetch ran for" $GroupRunTime.toString("mm' minutes 'ss' seconds'") "`n"

   #Start time for Teams Fetch
   $TeamsCheckStartTime = Get-Date
   Write-Host "Populating list of Teams at $TeamsCheckStartTime"

   If ($UsedGroups -eq $False) { # Populate the Teams hash table with a call to Get-UnifiedGroup
      Get-UnifiedGroup -Filter {ResourceProvisioningOptions -eq "Team"} -ResultSize Unlimited | ForEach { $TeamsList.Add($_.ExternalDirectoryObjectId, $_.DisplayName) } }
   Else { # We already have the $Groups variable populated with data, so extract the Teams from that data
      $Groups | ? {$_.ResourceProvisioningOptions -eq "Team"} | ForEach { $TeamsList.Add($_.ExternalDirectoryObjectId, $_.DisplayName) } }

   $TeamsCount = $TeamsList.Count

   #End time for Teams Fetch
   $TeamsCheckEndTime = Get-Date
   Write-Host "`n Teams Fetch Time:   $($TeamsCheckEndTime.ToString("yyyy-MM-dd HH:mm:ss")) `n"
   $TeamsRunTime = New-TimeSpan -End $TeamsCheckEndTime -Start $TeamsCheckStartTime
   Write-Host "Teams Fetch ran for" $TeamsRunTime.toString("mm' minutes 'ss' seconds'") "`n"
}
else{
   #Read groups and total number from input file

   #do something with the $content var created to make the groups alias identity the group object in the cloud

   #Get-UnifiedGroup -Identity "adamnewsted" | Format-List
   #$Groups = Get-Recipient -RecipientTypeDetails GroupMailbox | Sort-Object DisplayName

   $GroupsCount = $Groups.Count

   If ($UsedGroups -eq $False) { # Populate the Teams hash table with a call to Get-UnifiedGroup
      Get-UnifiedGroup -Filter {ResourceProvisioningOptions -eq "Team"} -ResultSize Unlimited | ForEach { $TeamsList.Add($_.ExternalDirectoryObjectId, $_.DisplayName) } }
   Else { # We already have the $Groups variable populated with data, so extract the Teams from that data
      $Groups | ? {$_.ResourceProvisioningOptions -eq "Team"} | ForEach { $TeamsList.Add($_.ExternalDirectoryObjectId, $_.DisplayName) } }

   $TeamsCount = $TeamsList.Count
}
#Write-Output $Groups   

# Set up progress bar
$ProgDelta = 100/($GroupsCount);
$CheckCount = 0;
$GroupNumber = 0;
Write-Host "`n-----------------------------------------------------------------------------";
#Start time for processing groups
$MainStartTime = Get-Date;
Write-Host "`nStart processing of $GroupsCount groups at $($MainStartTime.ToString("yyyy-MM-dd HH:mm:ss"))"


# Main loop
ForEach ($Group in $Groups) { #Because we fetched the list of groups with Get-Recipient, the first thing is to get the group properties
   try {
         <#if([String]::IsNullOrWhiteSpace((Get-content C:\Temp\Input.txt))){
            $G = Get-UnifiedGroup -Identity $Group.DistinguishedName
         }
         else{
            $G = Get-UnifiedGroup -Identity "Deletable Test Team" $Group.DistinguishedName
         }#>
         $G = Get-UnifiedGroup -Identity $Group.DistinguishedName
         $GroupNumber++
         $GroupStatus = $G.DisplayName + " ["+ $GroupNumber +"/" + $GroupsCount + "]"
         Write-Progress -Activity "Checking group" -Status $GroupStatus -PercentComplete $CheckCount
         
         $GroupTime = Get-Date -Format "yyyyMMddTHHmmss" 
         Write-Host "`n$GroupTime Processing group ($GroupNumber/$GroupsCount): " $G.DistinguishedName $G.objectID
         
         $CheckCount += $ProgDelta;
         $ObsoleteReportLine = $G.DisplayName;
         $SPOStatus = "Normal"
         $SPOActivity = "Document library in use";
         $SPOStorage = 0
         $NumberWarnings = 0;
         $NumberofChats = 0;
         $TeamsChatData = $Null;
         $TeamsEnabled = $False;
         $LastItemAddedtoTeams = "N/A";
         $MailboxStatus = $Null;
         $ObsoleteReportLine = "  "
   
         # Check who manages the group
         $ManagedBy = $G.ManagedBy
         If ([string]::IsNullOrWhiteSpace($ManagedBy) -and [string]::IsNullOrEmpty($ManagedBy)) {
            $ManagedBy = "***NO OWNERS!***"
            Write-Host "  $G.DisplayName has no group owners!" -ForegroundColor Red }
         Else {
            $ManagedBy = $G.ManagedBy -join ", "
         }
         
         # Group Age
         $GroupAge = (New-TimeSpan -Start $G.WhenCreated -End $Today).Days
           
         # Fetch information about activity in the Inbox folder of the group mailbox  
         $Data = (Get-ExoMailboxFolderStatistics -Identity $G.ExternalDirectoryObjectId -IncludeOldestAndNewestITems -FolderScope Inbox)
         If ([string]::IsNullOrEmpty($Data.NewestItemReceivedDate)) {
            $LastConversation = "No items found"}
         Else {
            $LastConversation = Get-Date ($Data.NewestItemReceivedDate) -Format g
         }
         $NumberConversations = $Data.ItemsInFolder
         $MailboxStatus = "Normal"
           
         If ($Data.NewestItemReceivedDate -le $WarningEmailDate) {
            Write-Host "Last conversation item created in" $G.DisplayName "was" $Data.NewestItemReceivedDate "-> Obsolete?"
            $ObsoleteReportLine = $ObsoleteReportLine + "Last Outlook conversation dated: " + $LastConversation + ". "
            $MailboxStatus = "Group Inbox Not Recently Used"
            $ObsoleteEmailGroups++
            $NumberWarnings++ }
         Else
            {# Some conversations exist - but if there are fewer than 20, we should flag this...
            If ($Data.ItemsInFolder -lt 20) {
               $ObsoleteReportLine = $ObsoleteReportLine + "Only " + $Data.ItemsInFolder + " Outlook conversation item(s) found. "
               $MailboxStatus = "Low number of conversations"
               $NumberWarnings++}
            }
         
         # Loop to check audit records for activity in the group's SharePoint document library
         If ($G.SharePointSiteURL -ne $Null) {
            $SPOStorage = (Get-SPOSite -Identity $G.SharePointSiteUrl).StorageUsageCurrent
            $SPOStorage = [Math]::Round($SpoStorage/1024,2) # SharePoint site storage in GB
            $AuditCheck = $G.SharePointDocumentsUrl + "/*"
            $AuditRecs = $Null
            $AuditRecs = (Search-UnifiedAuditLog -RecordType SharePointFileOperation -StartDate $WarningDate -EndDate $Today -ObjectId $AuditCheck -ResultSize 1)
            If ($AuditRecs -eq $Null) {
               Write-Host "No audit records found for" $SPOSite.Title "-> Potentially obsolete!"
               $ObsoleteSPOGroups++   
               $ObsoleteReportLine = $ObsoleteReportLine + "`n  No SPO activity detected in the last 90 days." }          
            }
            Else
            {
               # The SharePoint document library URL is blank, so the document library was never created for this group
               Write-Host "SharePoint team site never created for the group" $G.DisplayName 
               $ObsoleteSPOGroups++  
               $AuditRecs = $Null
               $ObsoleteReportLine = $ObsoleteReportLine + "`n  SPO document library never created." 
            }
   
            # Report to the screen what we found - but only if something was found...
            If ($ObsoleteReportLine -ne $G.DisplayName)
            {
               Write-Host $ObsoleteReportLine
            }
   
            # Generate the number of warnings to decide how obsolete the group might be...   
            If ($AuditRecs -eq $Null) {
               $SPOActivity = "No SPO activity detected in the last 90 days"
               $NumberWarnings++ }
            If ($G.SharePointDocumentsUrl -eq $Null) {
               $SPOStatus = "Document library never created"
               $NumberWarnings++ }
   
            $Status = "Pass"
            If ($NumberWarnings -eq 1){
               $Status = "Warning"
            }
            If ($NumberWarnings -eq 2){
               $Status = "Fail"
            } 
            If ($NumberWarnings -eq 3){
               $Status = "Severe"
            } 
            If ($NumberWarnings -gt 3){
               $Status = "Severe"
            } 
         
         # If the group is team-enabled, find the date of the last Teams conversation compliance record
         If ($TeamsList.ContainsKey($G.ExternalDirectoryObjectId) -eq $True) {
             $TeamsEnabled = $True
             [datetime]$DateOldTeams = "1-Jun-2021" # After this date, Microsoft should have moved the old Teams data to the new location
             $CountOldTeamsData = $False
         
            # Start by looking in the new location (TeamsMessagesData in Non-IPMRoot)
            $TeamsChatData = (Get-ExoMailboxFolderStatistics -Identity $G.ExternalDirectoryObjectId -IncludeOldestAndNewestItems -FolderScope NonIPMRoot | ? {$_.FolderType -eq "TeamsMessagesData" })
            If ($TeamsChatData.ItemsInFolder -gt 0) {
               $LastItemAddedtoTeams = Get-Date ($TeamsChatData.NewestItemReceivedDate) -Format g}
            $NumberOfChats = $TeamsChatData.ItemsInFolder
             
            # If the script is running before 1-Jun-2021, we need to check the old location of the Teams compliance records
            If ($Today -lt $DateOldTeams) {
               $CountOldTeamsData = $True
               $OldTeamsChatData = (Get-ExoMailboxFolderStatistics -Identity $G.ExternalDirectoryObjectId -IncludeOldestAndNewestItems -FolderScope ConversationHistory)
               ForEach ($T in $OldTeamsChatData) {
                  # We might have one or two subfolders in Conversation History; find the one for Teams
                  If ($T.FolderType -eq "TeamChat") {
                     If ($T.ItemsInFolder -gt 0) {
                        $OldLastItemAddedtoTeams = Get-Date ($T.NewestItemReceivedDate) -Format g
                     }
                     $OldNumberofChats = $T.ItemsInFolder
                  }
               }
            }
         
            If ($CountOldTeamsData -eq $True) { # We have counted the old date, so let's put the two sets together
               $NumberOfChats = $NumberOfChats + $OldNumberOfChats
               If (!$LastItemAddedToTeams) {
                 $LastItemAddedToTeams = $OldLastItemAddedToTeams
               }
            } # End if
         
            If (($TeamsEnabled -eq $True) -and ($NumberOfChats -le 100)) {
               Write-Host "  Team-enabled group" $G.DisplayName "has only" $NumberOfChats "compliance record(s)"
            }
   
         } # End if Processing Teams data
         
         # Generate a line for this group and store it in the report
         $ReportLine = [PSCustomObject][Ordered]@{
            GroupName                 = $G.DisplayName
            GroupOwners               = $ManagedBy
            Members                   = $G.GroupMemberCount
            ExternalGuests            = $G.GroupExternalMemberCount
            GroupDescription          = $G.Notes
            MailboxRecentActivity     = $MailboxStatus
            LastMailboxConversation   = $LastConversation
            GroupMailboxConversations = $NumberConversations
            TeamsEnabled              = $TeamsEnabled
            LastTeamsChat             = $LastItemAddedtoTeams
            TeamsChats                = $NumberofChats
            SharePointActivity        = $SPOActivity
            SharePointStorageGB       = $SPOStorage
            SharePointStatus          = $SPOStatus
            GroupCreationDate         = Get-Date ($G.WhenCreated) -Format g
            GroupAge_Days             = $GroupAge
            NumberWarnings            = $NumberWarnings
            Status                    = $Status}
         $Report.Add($ReportLine)
   
      #End of main loop
      }
      catch [System.Net.WebException],[System.IO.IOException] {
         Write-Host "I/O operation has been aborted."
         Receive-PSSession
      }
   }
   
   If ($TeamsCount -gt 0) { # We have some teams, so we can calculate a percentage of Team-enabled groups
       $PercentTeams = ($TeamsCount/$GroupsCount)
       $PercentTeams = ($PercentTeams).tostring("P") }
   Else {
       $PercentTeams = "No teams found"
   }
       
   # Create the HTML report
   $htmlbody = $Report | ConvertTo-Html -Fragment
   $htmltail = "<p>Report created for: " + $OrgName + "
                </p>
                <p>Number of groups scanned: " + $GroupsCount + "</p>" +
                "<p>Number of potentially obsolete groups (based on document library activity): " + $ObsoleteSPOGroups + "</p>" +
                "<p>Number of potentially obsolete groups (based on conversation activity): " + $ObsoleteEmailGroups + "<p>"+
                "<p>Number of Teams-enabled groups    : " + $TeamsCount + "</p>" +
                "<p>Percentage of Teams-enabled groups: " + $PercentTeams + "</body></html>" +
                "<p>-----------------------------------------------------------------------------------------------------------------------------"+
                "<p>Microsoft 365 Groups and Teams Activity Report"	
   $htmlreport = $htmlhead + $htmlbody + $htmltail
   $htmlreport | Out-File $ReportFile  -Encoding UTF8
   
   $Report | Export-CSV -NoTypeInformation $CSVFile
   $Report | Out-GridView
   
   
   # Summary note
   Write-Host "`n-----------------------------------------------------------------------------`n";
   Write-Host "Results"
   Write-Host "-------"
   Write-Host "Number of Microsoft 365 Groups scanned                          :" $GroupsCount
   Write-Host "Potentially obsolete groups (based on document library activity):" $ObsoleteSPOGroups
   Write-Host "Potentially obsolete groups (based on conversation activity)    :" $ObsoleteEmailGroups
   Write-Host "Number of Teams-enabled groups                                  :" $TeamsList.Count
   Write-Host "Percentage of Teams-enabled groups                              :" $PercentTeams
   Write-Host " "
   Write-Host "Summary CSV report in " $CSVFile
   Write-Host "Summary HTML report in " $ReportFile
   Write-Host "PowerShell log in" $TextFileName
   
   $endTime = Get-Date;
   Write-Host "`nScript End Time: $($endTime.ToString("yyyy-MM-dd HH:mm:ss"))"
   $runTime = New-TimeSpan -End $endTime -Start $startTime
   Write-Host "Script ran for" $runTime.ToString("d' day(s) 'hh' hours 'mm' minutes 'ss' seconds'") "`n"
   
   Stop-Transcript
   
   #Disconnect Exchange Online, Skype and Security & Compliance center session
   Get-PSSession | Remove-PSSession
   #Disconnect Teams connection
   Disconnect-MicrosoftTeams
   #Disconnect SharePoint connection
   Disconnect-SPOService
   Write-Host "All sessions in the current window has been removed." -ForegroundColor Yellow
