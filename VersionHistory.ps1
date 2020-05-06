powershell -ExecutionPolicy ByPass {

   #Load SharePoint CSOM Assemblies
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

#Set Parameters
$SiteURL="https://mdanz.sharepoint.com/sites/CRMDevelopment"
$LibraryName="Documents"

#Outputs
$ReportOutput = "C:\Windows\Temp\Script\VersionHistory.csv" #Where report is stored on local machine.
$WebUrl = "https://mdanz.sharepoint.com/Sites/CRMDevelopment/"  #Site report will be stored on SharePoint

# Define Credentials - heidi
[string]$userName = "heidi.goosen@mda.nz"
[string]$userPassword = "xxxxx"

   #Upload Variables  - Admin Heidi
    [string]$AdminName = "heidi.goosen@mda.nz"
    [string]$AdminPassword = "xxxxxx"

    #Upload Credentials 
    [SecureString]$SecurePass = ConvertTo-SecureString $AdminPassword -AsPlainText -Force 
    $credentials = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $AdminName, $(convertto-securestring $AdminPassword -asplaintext -force)
    $AdminSiteURL="https://mdanz-admin.sharepoint.com"

    #Connect to SharePoint Online Admin Center
    Connect-SPOService -Url $AdminSiteURL –Credential $credentials

    #Get All site collections
    $SiteCollections = Get-SPOSite -Limit All
    Write-Host "Total Number of Site collections Found:"$SiteCollections.count -f Yellow
    Write-Host "---------------alrighty-----------------"

Try {
   Write-Host "---------------4-----------------" 
    #Setup Credentials to connect
   # $Cred= Get-Credential
    #$Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Cred.Username, $Cred.Password)

        #Upload Credentials 
    [SecureString]$SecureUserPass = ConvertTo-SecureString $userPassword -AsPlainText -Force 
    $Credentials = New-Object -TypeName Microsoft.SharePoint.Client.SharePointOnlineCredentials -argumentlist $userName, $(convertto-securestring $userPassword -asplaintext -force)
 


           Write-Host "---------------5-----------------"  
    #Setup the context
    $Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
    $Ctx.Credentials = $Credentials
               Write-Host "---------------6-----------------" 


#Read more: https://www.sharepointdiary.com/2016/12/sharepoint-online-version-history-report-using-powershell.html#ixzz6KURC5vlh

   ## Crete credential Object
    # [SecureString]$secureString =  ConvertTo-SecureString $userPassword -AsPlainText -Force 
 #[PSCredential]$Credentials_2 = New-Object System.Management.Automation.PSCredential -ArgumentList $userName,$(convertto-securestring $userPassword -asplaintext -force)
          Write-Host "---------------4-----------------" 

   # #Setup Credentials to connect
 #$Cred = Get-Credential
 #$Credentials_2 = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Cred.Username, $Cred.Password)

         Write-Host "---------------5-----------------"
 ##Setup the context
 #$Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
 #$Ctx.Credentials = $Credentials_2
    
    #Get the web & Library
    $Web=$Ctx.Web
    $Ctx.Load($Web)
    $List = $Web.Lists.GetByTitle($LibraryName)
    Write-Host "--------------------Total Number 65 ---------------------------" 
    $Ctx.ExecuteQuery()
   
         
    #Get All Files of from the document library - Excluding Folders
    $Query =  New-Object Microsoft.SharePoint.Client.CamlQuery
    $Query.ViewXml = "<View Scope='RecursiveAll'><Query><Where><Eq><FieldRef Name='FSObjType' /><Value Type='Integer'>0</Value></Eq></Where><OrderBy><FieldRef Name='ID' /></OrderBy></Query></View>"
    $ListItems=$List.GetItems($Query)
    $Ctx.Load($ListItems)
    $Ctx.ExecuteQuery()
 

    $VersionHistoryData = @()
    #Iterate throgh each version of file
    Foreach ($Item in $ListItems)
    {
        $File = $Web.GetFileByServerRelativeUrl($Item["FileRef"])
        $Ctx.Load($File)
        $Ctx.Load($File.ListItemAllFields)
        $Ctx.Load($File.Versions)

        $Ctx.ExecuteQuery()

        $today = Get-Date 
        $modifiedDate = $File.TimeLastModified
        $daysSinceModified = (New-TimeSpan -Start $modifiedDate -End $today).Days
        

        Write-host -f Yellow "Processing File:"$File.Name
       # If($File.Versions.Count -ge 1 -AND $daysSinceModified  -eq 172)
       # If($File.Versions.Count -ge 1 -AND $daysSinceModified  -gt 0 -AND $daysSinceModified -lt 173)
       If($File.Versions.Count -ge 1)
        {
            #Calculate Version Size
            $VersionSize = $File.Versions | Measure-Object -Property Size -Sum | Select-Object -expand Sum
            If($Web.ServerRelativeUrl -eq "/")
            {
                $FileURL = $("{0}{1}" -f $Web.Url, $File.ServerRelativeUrl)
            }
            Else
            {
                $FileURL = $("{0}{1}" -f $Web.Url.Replace($Web.ServerRelativeUrl,''), $File.ServerRelativeUrl)
            }
 
            #Send Data to object array
            $VersionHistoryData += New-Object PSObject -Property @{
            'daysSinceModified' =$daysSinceModified
            'File Name' = $File.Name
            'Versions Count' = $File.Versions.count
            'Modified On' = $File.TimeLastModified
            'File Size' = ($File.Length/1KB)
            'Version Size' = ($VersionSize/1KB)
            'URL' = $FileURL
            'Modified by Option 2' = $Item["Editor"].LookupValue

            }
        }
    }

         #Export the data to CSV
         $VersionHistoryData | Export-Csv $ReportOutput -NoTypeInformation
 
        #Setup Credentials to connect
        $Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($userName,(ConvertTo-SecureString $userPassword -AsPlainText -Force))
  
        #Set up the context
        $Context = New-Object Microsoft.SharePoint.Client.ClientContext($WebUrl)
        $Context.Credentials = $Credentials
 
        #Get the Library
        $Library =  $Context.Web.Lists.GetByTitle($LibraryName)
 
        #Get the file from disk
        $FileStream = ([System.IO.FileInfo] (Get-Item $ReportOutput)).OpenRead()
        #Get File Name from source file path
        $SourceFileName = Split-path $ReportOutput -leaf
   
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

             
        Write-host -f Green "Versioning History Report has been Generated Successfully!"
}
Catch {
    write-host -f Red "Error Generating Version History Report!" $_.Exception.Message
}

}
