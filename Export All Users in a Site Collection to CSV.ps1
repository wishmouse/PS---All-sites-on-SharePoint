#Config Variables
$SiteURL = "https://xxx.sharepoint.com/sites/xxxx"
$CSVFile = "C:\users\xx\downloads\UserData.csv"
   
#Connect to PnP Online
Connect-PnPOnline -Url $SiteURL -Credentials (Get-credential)  #-UseWebLogin
  
#Get All users of the site collection
$Users = Get-PnPUser
$UserData=@()
 
#Loop through Users and get properties
ForEach ($User in $Users)
{
    $UserData += New-Object PSObject -Property @{
        "User Name" = $User.Title
        "Login ID" = $User.LoginName
        "E-mail" = $User.Email
        "User Type" = $User.PrincipalType
    }
}
$UserData
#Export Users data to CSV file
$UserData | Export-Csv -NoTypeInformation $CSVFile

