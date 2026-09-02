# Report-ArchivedSPOFiles.PS1

# A script to report on SharePoint Online files that have been archived (moved to Microsoft 365 Archive), either by
# explict user action or by retention policy or label. The script reports on the files that have been archived, 
# including the file name, path, and the date it was archived.

# Needs to be run with an account or Azure Automation account that has access to the target SharePoint Online 
# sites.
Connect-MgGraph -Scopes "Sites.Read.All" -Nowelcome

function Get-SPOSiteDetails {
    param (
        [string]$SiteUrl
    )

    $SiteId = $SiteUrl.Split('//')[1].split("/")[0] + ":/sites/" + $SiteUrl.Split('//')[1].split("/")[2]
    $Site = Get-MgSite -SiteId $SiteId

    # Get site
    Try {
        $Site = Get-MgSite -SiteId $SiteId -ErrorAction Stop
    } Catch {
        Write-Error "Failed to retrieve site information: $_"
        Return
    }
    # Get the default document library for the site
    Try {
        $Drive = Get-MgSiteDefaultDrive -SiteId $Site.Id -ErrorAction Stop
    } Catch {
        Write-Error "Failed to retrieve default document library: $_"
        Return
    }    
    return [PSCustomObject]@{
        Site  = $Site
        Drive = $Drive
    }

}

function FormatFileSize {
    # Format File Size nicely
    param (
            [parameter(Mandatory = $true)]
            $InFileSize
        ) 
    
    If ($InFileSize -lt 1KB) { # Format the size of a document
        $FileSize = $InFileSize.ToString() + " B" 
    } 
    ElseIf ($InFileSize -lt 1MB) {
        $FileSize = $InFileSize / 1KB
        $FileSize = ("{0:n2}" -f $FileSize) + " KB"
    } 
    Elseif ($InFileSize -lt 1GB) {
        $FileSize = $InFileSize / 1MB
        $FileSize = ("{0:n2}" -f $FileSize) + " MB" 
    }
    Elseif ($InFileSize -ge 1GB) {
        $FileSize = $InFileSize / 1GB
        $FileSize = ("{0:n2}" -f $FileSize) + " GB" 
    }
    Return $FileSize
} 

function Get-DriveItems {
    [CmdletBinding()]
    param (
        [Parameter()]
        $DriveId,
        [Parameter()]
        $FolderId
    )

    # Return details of archived files found in a document library. The function is recursive and will process 
    # all folders in the library. The function returns a list of archived files found in the library.

    # Get data for a folder and its children - Get-MgDriveItemChild doesn't return the archive status currently
    $Uri = ("https://graph.microsoft.com/V1.0/drives/{0}/items/{1}/children" -f $DriveId, $FolderId)
    [array]$Data = Invoke-MgGraphRequest -Uri $Uri -Method Get -OutputType PsObject | Select-Object -ExpandProperty Value
    [array]$Folders = $Data | Where-Object {$_.folder.childcount -gt 0} | Sort-Object Name
    $Global:TotalFolders = $TotalFolders + $Folders.Count
    [array]$ArchivedFiles = $Data | Where-Object {$null -ne $_.file.mimetype -and $null -ne $_.file.archiveStatus}
    #[array]$Data = Get-MgDriveItemChild -DriveId $Drive -DriveItemId $FolderId -All

    # Process the files
    ForEach ($File in $ArchivedFiles) {  

        If ($File.LastModifiedDateTime) {
            $LastModifiedDateTime = Get-Date $File.LastModifiedDateTime -format 'dd-MMM-yyyy HH:mm'
        } Else {
            $LastModifiedDateTime = $null
        }
        If ($File.CreatedDateTime) {
            $FileCreatedDateTime = Get-Date $File.CreatedDateTime -format 'dd-MMM-yyyy HH:mm'
        }
        
        $ReportLine = [PSCustomObject]@{
            FileName                = $File.Name
            Folder                  = $File.parentreference.name
            Size                    = (FormatFileSize $File.Size)
            Created                 = $FileCreatedDateTime
            Author                  = $File.CreatedBy.User.DisplayName
            LastModified            = $LastModifiedDateTime
            'Last modified by'      = $File.LastModifiedBy.User.DisplayName
            SizeInBytes             = $File.Size
            ArchiveStatus           = $File.File.ArchiveStatus
            WebURL                  = $File.WebUrl
        }
        $ReportData.Add($ReportLine)
    }

    # Process the child folders found in the root
    ForEach ($Folder in $Folders) {
        Write-Host ("Processing folder {0} ({1} files/size {2})" -f $Folder.Name, $Folder.folder.childcount, (FormatFileSize $Folder.Size))
        Get-DriveItems -Drive $DriveId -FolderId $Folder.Id
    }
}


$Uri = "https://redmondassociates.sharepoint.com/sites/O365ExchPro"
$SiteDetails = Get-SPOSiteDetails -SiteUrl $Uri

Write-Output ("Processing site: {0}..." -f $SiteDetails.Site.DisplayName)
$SiteId = $SiteDetails.Site.Id
$DriveId = $SiteDetails.Drive.Id
$RootId = (Get-MgDriveRoot -DriveId $DriveId).Id

$Global:ReportData = [System.Collections.Generic.List[Object]]::new()
[array]$ArchivedFiles = Get-DriveItems -Drive $DriveId -FolderId $RootId


# An example script used to illustrate a concept. More information about the topic can be found in the Microsoft 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 