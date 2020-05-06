#Connect to Exchange Online to access the unified audit log
#clear runspace
Get-PSSession | Remove-PSSession


$username = "== your name ==@nrcl.onmicrosoft.com"
$password = ConvertTo-SecureString "===Password===" -AsPlainText -Force
$psCred = New-Object System.Management.Automation.PSCredential -ArgumentList ($username, $password)


   #Upload Variables  - Admin Heidi
    [string]$AdminName = "== your name ==@nrcl.onmicrosoft.com"
    [string]$AdminPassword = "=== Password ==="

    #Upload Credentials 
    [SecureString]$SecurePass = ConvertTo-SecureString $AdminPassword -AsPlainText -Force 
    $credentials = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $AdminName, $(convertto-securestring $AdminPassword -asplaintext -force)
    $AdminSiteURL="https://nrcl.sharepoint.com"

    
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $psCred -Authentication Basic -AllowRedirection 
Import-PSSession $Session -AllowClobber -DisableNameChecking


$OutputPath = "c:\== your location ==\"
$Today  = (Get-Date).Date   # a REAL DateTimeObject for midnight today, not a formatted string
$FileAccessOperations = 'FileAccessed', 'FileDownloaded', 'PageViewed', 'FileModified', 'FileUploaded', 'SharingSet'
$intDays = 1

# loop through the days and hours to collect audit data
$result = for ($days = 0; $days -le $intDays; $days++) {
    for ($hours = 1; $hours -ge 0; $hours--){
        $StartDate = ($Today.AddDays(-$days)).AddHours($hours)
        $EndDate   = ($Today.AddDays(-$days)).AddHours($hours + 1)

        $Audit = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -Operations $FileAccessOperations -ResultSize 5000 

        $ConvertAudit = $Audit | Select-Object -ExpandProperty AuditData | ConvertFrom-Json
        # output toget collected in $result
        $ConvertAudit | Select-Object CreationTime,UserId,Operation,Workload,ObjectID,SiteUrl,SourceFileName,ClientIP,UserAgent
                Write-Host $StartDate `t $Audit.Count
    }
}

# group the resulting data on the UserId property and output each group as separate csv.
$result | Group-Object UserId | ForEach-Object {
    $OutputFile = Join-Path -Path $OutputPath -ChildPath ('Audit {0}.csv' -f $_.Name)
    $_.Group | Export-Csv $OutputFile -NoTypeInformation


   Get-ChildItem $OutputPath -Filter *.csv | % {


        $eachFile =Get-ChildItem $Outputfile -Filter *.csv

        $eachFile | ForEach-Object{
        Write-host 'lets see what happnes'

         #Setup Credentials to connect
        $Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($AdminUser,(ConvertTo-SecureString $AdminPassword -AsPlainText -Force))

        #Site report will be stored on SharePoint
        $SiteURL="https://nrcl.sharepoint.com/sites/nrg"
        $LibraryName="Audit History"

         #Set up the context
        $Context = New-Object Microsoft.SharePoint.Client.ClientContext($WebUrl)
        $Context.Credentials = $Credentials
 
        #Get the Library
        $Library =  $Context.Web.Lists.GetByTitle($LibraryName)

 
        #Get the file from disk
        $FileStream = ([System.IO.FileInfo] (Get-Item $OutputFile)).OpenRead()
        #Get File Name from source file path
        $SourceFileName = Split-path $eachFile -leaf

         #sharepoint online upload file powershell
        $FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
        $FileCreationInfo.Overwrite = $true
        $FileCreationInfo.ContentStream = $FileStream
        $FileCreationInfo.URL = $SourceFileName
        $FileUploaded = $Library.RootFolder.Files.Add($FileCreationInfo)
  
        #powershell upload single file to sharepoint online
        $Context.Load($FileUploaded)
        $Context.ExecuteQuery()
 
        #Close file stream
        $FileStream.Close()
  
        write-host "File has been uploaded!"

        }
 }

}

       