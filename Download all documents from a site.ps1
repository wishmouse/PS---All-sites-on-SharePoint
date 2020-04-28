#Load SharePoint CSOM Assemblies
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
 
Function Download-FilesFromLibrary()
{
    param
    (
        [Parameter(Mandatory=$true)] [string] $SiteURL,
        [Parameter(Mandatory=$true)] [string] $LibraryName,
        [Parameter(Mandatory=$true)] [string] $TargetFolder
    )
 
    Try {
        #Setup Credentials to connect
        $Cred= Get-Credential
        $Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Cred.Username, $Cred.Password)
 
        #Setup the context
        $Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
        $Ctx.Credentials = $Credentials
      
        #Get all files from the Library
        $List = $Ctx.Web.Lists.GetByTitle($LibraryName)
        $Ctx.Load($List)
        $Ctx.ExecuteQuery()
         
        #Get Files from the Folder
        $Folder = $List.RootFolder
        $FilesColl = $Folder.Files
        $Ctx.Load($FilesColl)
        $Ctx.ExecuteQuery()
 
        Foreach($File in $FilesColl)
        {
            $TargetFile = $TargetFolder+$File.Name
            #Download the file
            $FileInfo = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($Ctx,$File.ServerRelativeURL)
            $WriteStream = [System.IO.File]::Open($TargetFile,[System.IO.FileMode]::Create)
            $FileInfo.Stream.CopyTo($WriteStream)
            $WriteStream.Close()
        }
        Write-host -f Green "All Files from Library '$LibraryName' Downloaded to Folder '$TargetFolder' Successfully!" $_.Exception.Message
  }
    Catch {
        write-host -f Red "Error Downloading Files from Library!" $_.Exception.Message
    }
}
 
#Set parameter values
$SiteURL="https://xxx.sharepoint.com/sites/xxx/"
$LibraryName="Documents"
$TargetFolder="C:\users\xxxxx\Downloads"
 
#Call the function to download file
Download-FilesFromLibrary -SiteURL $SiteURL -LibraryName $LibraryName -TargetFolder $TargetFolder


