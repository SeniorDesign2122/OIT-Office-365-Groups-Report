##########################################################################
<#                 Connection Mods

Installs AD connection Mod (Azure AD V1 & AzureAD V2) (Run in PowerShell as Admin)
Install-Module -Name MSOnline (Azure AD V1)
Install-Module -Name AzureAD (Azure AD V2)

Installs 
Exchange connection Mod (Run in PowerShell as Admin)
Install-Module ExchangeOnlineManagement

Download & install Sharepoint Mod (Link below)
Install-Module Microsoft.Online.SharePoint.PowerShell
& Install-Module SharePointPnPPowerShellOnline

Installs Teams connection Mod (Run in PowerShell as Admin)
Install-Module -Name MicrosoftTeams -RequiredVersion 3.0.0	


#>
##########################################################################

#./ConnectO365Services.ps1 -UserName ctv8426@WMICH.EDU -Password "???" -Services AzureAD,ExchangeOnline,MSOnline,Teams,SharePoint -SharePointHostName wexchange.wmich.edu

#Connects to these services but wont reconize "Get-UnifiedGroups" call, issue with user account? (Premissions not available on our Student Accounts)
#./ConnectO365Services.ps1 -Services AzureAD,ExchangeOnline,MSOnline -SharePointHostName wexchange.wmich.edu -MFA

#Connects to Teams cant get powershell to work as of now / SharePoint still not connecting with MFA
#./ConnectO365Services.ps1 -Services AzureAD, SharePoint, ExchangeOnline,Teams -SharePointHostName wexchange.wmich.edu -MFA

#Connects to Teams cant get powershell to work as of now / SharePoint still not connecting with MFA
./service_connect.ps1 -Services AzureAD, SharePoint, ExchangeOnline,Teams -SharePointHostName wmutest1 -UserName cs4900_admin@wmutest1.onmicrosoft.com -Password "???"



##################### Old Code #####################################
#Connects to all resources on Office 365 in one PowerShell Window
<#
$orgName="<OITteam1.onmicrosoft.com>"
$acctName="<ctv8426_WMICH.EDU#EXT#@OITteam1.onmicrosoft.com>"
$credential = Get-Credential -UserName $acctName -Message ""
#Azure Active Directory
Connect-MsolService -Credential $credential
#SharePoint Online
Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
Connect-SPOService -Url https://$orgName-admin.sharepoint.com -credential $credential
#Exchange Online
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -ShowProgress $true
#Security & Compliance Center
Connect-IPPSSession -UserPrincipalName $acctName
#Teams Online
Import-Module MicrosoftTeams
Connect-MicrosoftTeams -Credential $credential

#$credential = 
#Connect-AzureAD -TenantId 565e27df-64a8-4764-901a-cc6374679c15 -Credential $credential

#get-azureadgroup
#>
##################### Old Code #####################################

#https://petri.com/identifying-obsolete-office-365-groups-powershell

#Checking Groups for Low SharePoint Activity
<#
$WarningDate = (Get-Date).AddDays(-90)
$Today = (Get-Date)
$Groups = Get-UnifiedGroup 
$ObsoleteGroups = 0
ForEach ($G in $Groups) {
   If ( $Null -ne $G.SharePointDocumentsUrl)
      {
      $SPOSite = (Get-SPOSite -Identity $G.SharePointDocumentsUrl.replace("/Shared Documents", ""))
      Write-Host "Checking" $SPOSite.Title "..."
      $AuditCheck = $G.SharePointDocumentsUrl + "/*"
      $AuditRecs = 0
      $AuditRecs = (Search-UnifiedAuditLog -RecordType SharePointFileOperation -StartDate $WarningDate -EndDate $Today -ObjectId $AuditCheck -SessionCommand ReturnNextPreviewPage)
      If ($null -eq $AuditRecs) 
         {
         Write-Host "No audit records found for" $SPOSite.Title "-> It is potentially obsolete!"
         $ObsoleteGroups++   
         }
      Else 
         {
         Write-Host $AuditRecs.Count "audit records found for " $SPOSite.Title "the last is dated" $AuditRecs.CreationDate[0]
       }}
   Else
         {
         Write-Host "SharePoint has never been used for the group" $G.DisplayName 
         $ObsoleteGroups++   
         }
    }
Write-Host $ObsoleteGroups "obsolete group document libraries found out of" $Groups.Count "checked"


#>
#Checking Group Mailboxes for Low Conversation Activity
$Groups = Get-UnifiedGroup
$BadGroups = 0
$WarningDate = (Get-Date).AddDays(-2)
ForEach ($G in $Groups) {
  Write-Host "Checking Inbox traffic for" $G.DisplayName
  $CheckDate = (Get-MailboxFolderStatistics -Identity $G.Alias -IncludeOldestAndNewestItems -FolderScope Inbox).NewestItemReceivedDate
  If ($CheckDate -le $WarningDate)
  {
    Write-Host "Last conversation item created in" $G.DisplayName "was" `
    $Data.NewestItemReceivedDate "-> Could be Obsolete?"
    $BadGroups++
  }
  Else
  {
    Write-Host $G.DisplayName "has" $Data.ItemsInFolder "conversation items amounting to" $Data.FolderSize
  }
}
Write-Host $BadGroups "Obsolete Groups found out of" $Groups.Count

<#
#Checking Teams Activity
$TData = (Get-MailboxFolderStatistics -Identity $G.Alias -IncludeOldestAndNewestItems -FolderScope ConversationHistory).NewestItemReceivedDate
If ($TData -le $WarningDate)
      {
      Write-Host "Last Teams chat item created in" $G.DisplayName "was" `
             $Data.NewestItemReceivedDate "-> Could be Obsolete?"
      }
#>