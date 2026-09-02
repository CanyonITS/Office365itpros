# Report-ArchivedSPOFiles.PS1

# A script to report on SharePoint Online files that have been archived (moved to Microsoft 365 Archive), either by
# explict user action or by retention policy or label. The script reports on the files that have been archived, 
# including the file name, path, and the date it was archived.

# Needs to be run with an account or Azure Automation account that has access to the target SharePoint Online 
# sites. Given the number of files that might be checked, best to use an Azure Automation runbook with a managed identity that has access to the sites. The script uses Microsoft Graph PowerShell SDK to connect and retrieve data.

# V1.0 2-Sept-2026
# GitHub Link:

function Get-SPOSiteDetails {
    param (
        [string]$SiteUrl
    )

    # Function to take a SharePoint Online site URL and return the site and its default document library (drive) details. The function uses Microsoft Graph to retrieve the information.

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
    Try {
        [array]$Data = Invoke-MgGraphRequest -Uri $Uri -Method Get -OutputType PsObject 
    } Catch {
        Write-Error "Failed to retrieve drive items: $_"
        Return
    }

    If ($Data.'@odata.nextlink') {
        # More data needs to be retrieved, so get the next page of data
        $NextLink = $Data.'@odata.nextlink'
        $FolderData = $Data.Value
        While ($NextLink) {
            Try {
                [array]$NextData = Invoke-MgGraphRequest -Uri $NextLink -Method Get -OutputType PsObject 
            } Catch {
                Write-Error "Failed to retrieve next page of drive items: $_"
                Return
            }
            $FolderData += $NextData.Value
            $NextLink = $NextData.'@odata.nextlink'
        }
    } Else {
        $FolderData = $Data.Value
    }

    [array]$Folders = $FolderData | Where-Object {$_.folder.childcount -gt 0} | Sort-Object Name
    $Global:TotalFolders = $TotalFolders + $Folders.Count
    [array]$ArchivedFiles = $FolderData | Where-Object {$null -ne $_.file.mimetype -and $null -ne $_.file.archiveStatus}
    
    # Equivalent Graph SDK cmdlet when the SDK is updated (last checked with V2.39)
    # [array]$Data = Get-MgDriveItemChild -DriveId $Drive -DriveItemId $FolderId -All

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
            Site                    = $SiteName
            FileName                = $File.Name
            Folder                  = $File.parentreference.name
            Size                    = (FormatFileSize $File.Size)
            Created                 = $FileCreatedDateTime
            Author                  = $File.CreatedBy.User.DisplayName
            LastModified            = $LastModifiedDateTime
            'Last modified by'      = $File.LastModifiedBy.User.DisplayName
            SizeInBytes             = $File.Size
            ArchiveStatus           = $File.File.ArchiveStatus
        }
        $ReportData.Add($ReportLine)
    }

    # Process the child folders found in the root
    ForEach ($Folder in $Folders) {
        Write-Host ("Processing folder {0} ({1} files/size {2})" -f $Folder.Name, $Folder.folder.childcount, (FormatFileSize $Folder.Size))
        Get-DriveItems -DriveId $DriveId -FolderId $Folder.Id
    }
}

# Start of processing
# This is the code for app-only authentication. The permissions needed are
# Sites.Read.All and Group.Read.All. The app must be registered in Entra ID and the certificate uploaded to the app.
# registration. The script uses the certificate to authenticate to Microsoft Graph. Amend the values for $AppId
# $TenantId and $CertificateThumbprint to match your app registration.

$AppId = 'd4c18027-cb9c-4983-b2d8-8ba6a5e753e7'
$TenantId = 'b662313f-14fc-43a2-9a7a-d2e27f4f3478'
$CertificateThumbprint = '8D225839321765FE6D224D393CDE5A61A45CD846'
Connect-MgGraph -NoWelcome -AppId $AppId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint

# Find the Microsoft 365 groups in the tenant. We use the groups to find the associated SharePoint Online sites.
# The script will then check each site for archived files.

$Global:ReportData = [System.Collections.Generic.List[Object]]::new()
Try {
    [array]$Groups = Get-MgGroup -Filter "groupTypes/any(c:c eq 'unified')" -All -ErrorAction Stop
} Catch {
    Write-Error "Failed to retrieve Microsoft 365 groups: $_"
    Return
}

$Groups = $Groups | Sort-Object DisplayName
Write-Output ("{0} Microsoft 365 groups found" -f $Groups.Count)

ForEach ($Group in $Groups) {
    Write-Output ("Processing group: {0}..." -f $Group.DisplayName)
    $Uri = ('https://graph.microsoft.com/v1.0/groups/{0}/sites/root' -f $Group.Id)
    Try {
        $GroupSite = Invoke-MgGraphRequest -Uri $Uri -Method Get -OutputType PsObject -ErrorAction Stop
    } Catch {
        Write-Error "Failed to retrieve site for group $($Group.DisplayName): $_"
        Continue
    }  
    $SiteId = $GroupSite.Id
    Try {
        # Get default document library (drive) for the site.
        $Drive = Get-MgSiteDefaultDrive -SiteId $SiteId -ErrorAction Stop
    } Catch {
        Write-Error "Failed to retrieve default drive for site $($GroupSite.DisplayName): $_"
        Continue
    }
    $DriveId = $Drive.Id
    $Global:SiteName = $GroupSite.displayName
    Try {
        $RootId = Get-MgDriveRoot -DriveId $DriveId -ErrorAction Stop
    } Catch {
        Write-Error "Failed to retrieve root folder for drive $($Drive.Name): $_"
        Continue
    }
    $RootId = $RootId.Id
    [array]$ArchivedFiles = Get-DriveItems -DriveId $DriveId -FolderId $RootId
}

# Now analyze the results and output a report of the archived files found in the Microsoft 365 groups.

# An example script used to illustrate a concept. More information about the topic can be found in the Microsoft 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 