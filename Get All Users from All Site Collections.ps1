Import-Module Microsoft.Online.Sharepoint.PowerShell -DisableNameChecking
 
$AdminSiteURL="https://xxx-admin.sharepoint.com/"
 
#Connect to SharePoint Online Admin
Write-host "Connecting to Admin Center..." -f Yellow
Connect-SPOService -url $AdminSiteURL -Credential (Get-Credential)
 
Write-host "Getting All Site collections..." -f Yellow
#Get each site collection and users
$Sites = Get-SPOSite -Limit ALL
 
Foreach($Site in $Sites)
{
    Write-host "Getting Users from Site collection:"$Site.Url -f Yellow
    Get-SPOUser -Limit ALL -Site $Site.Url | Select DisplayName, LoginName
}

