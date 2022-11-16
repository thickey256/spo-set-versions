##############################################
# Global
##############################################
$global:versionSize = 0

##############################################
# Functions
##############################################


function ConnectToMSGraph ($parameters) {
  
    Connect-MgGraph -ClientId $parameters.clientId.Value -TenantId $parameters.tenantId.Value -CertificateThumbprint $parameters.thumbprint.Value
}
function ProcessSites($sites) {

    foreach($site in $sites)
    {
        ## Connect to Site
        Connect-PnPOnline -Url $site.SiteUrl -ClientId $parameters.clientId.Value -Tenant $parameters.tenantId.Value -Thumbprint $parameters.thumbprint.Value       

        try {
            $siteId = (Get-PnPSite -Includes Id -ErrorAction Stop).Id.Guid
            Write-Host "Processing Site: $($site.SiteUrl)" -ForegroundColor DarkGray
        }
        catch {
            Write-Host "No Site Found with Url: $($site.SiteUrl)" -ForegroundColor Yellow
            continue
        }
        
        $docLibs = GetLibrariesInSite $site$site.SiteUrl

        foreach($lib in $docLibs)
        {
            SetLibraryVersionConfig $site $lib $siteId
        }

        ## Process Subs
        $subs = Get-PnPSubWeb -Recurse

        foreach($sub in $subs)
        {
            # Switch the context to the sub (no need to catch as subsite should always exist)
            Connect-PnPOnline -Url $sub.Url -ClientId $parameters.clientId.Value -Tenant $parameters.tenantId.Value -Thumbprint $parameters.thumbprint.Value
            Write-Host " Processing Site: $($sub.Url)" -ForegroundColor DarkGray

            $docLibs = GetLibrariesInSite $sub

            foreach($lib in $docLibs)
            {
                # Still pass in the site object as this has the version info (if applicable)#
                # Site id for sub site is "SiteCollectionId,SubSiteId" (used for MSGraph)
                SetLibraryVersionConfig $site $lib "$siteId,$($sub.Id)"
            }
        }

        ## Update Site Saved MB
        $site.SizeSavedMB = $global:versionSize / 1024 / 1024

        Write-Host "Saved: $($site.SizeSavedMB) MB from Site $($site.Url)" -ForegroundColor Green

        ## Rest for next site
        $global:versionSize = 0

    }
    
}


function GetLibrariesInSite($site) {

    ## Get doc libs (that aren't system lists)
    $docLibs = Get-PnPList -Includes IsSystemList, Fields | Where { ($_.BaseType -eq "DocumentLibrary" ) -and !($_.IsSystemList) -and ($_.Title -ne "Site Pages")}
    Write-Host " Found $($docLibs.Length) document libraries" -ForegroundColor DarkGreen
    return $docLibs
}

function SetLibraryVersionConfig($site, $lib, $siteId) {

    # Use global my default
    $majorVersionCount = $parameters.majorVersionCount.Value
    $minorVersionCount = $parameters.minorVersionCount.Value

    ## Have we got specific Site Version setting
    if ($site.SpecificVersionSetting)
    {
        $majorVersionCount = $site.MajorVersionCount
        $minorVersionCount = $site.MinorVersionCount
    }
    
    $enableMajorVersions = $majorVersionCount -gt 0
    $enableMinorVersions = $minorVersionCount -gt 0

    if(!$parameters.whatIfMode.Value)
    {
        try {
            # Set Versions
            Set-PnPList -Identity $lib.Id -EnableVersioning $enableMajorVersions -MajorVersions $majorVersionCount -EnableMinorVersions $enableMinorVersions -MinorVersions $minorVersionCount -ErrorAction Stop | Out-Null 
            Write-Host "  Successfully configured (Major: $majorVersionCount, Minor: $minorVersionCount) for library: $($lib.Title)" -ForegroundColor Green
        }
        catch {
            Write-Host "  Unsuccessfully configuraiton attempt (Major: $majorVersionCount, Minor: $minorVersionCount) for library: $($lib.Title)" -ForegroundColor DarkRed
        }
    }

    if($parameters.deleteOldMajorVersions.Value)
    {
        # Delete versions PnP
        #DeleteOldMajorVersions $lib $majorVersionCount

        # Delete Versions MSGraph
        DeleteOldMajorVersionsGraph $lib $majorVersionCount $siteId
    }
}

function DeleteOldMajorVersions($docLib,$majorVersionCount)
{
    $items = Get-PnPListItem -List $docLib.Id -PageSize 500

    foreach ($item in $items)
    {
        try {
            $versionHistory = Get-PnPFileVersion -Url $item.FieldValues.FileRef -ErrorAction Stop
        } catch {
            $versionHistory = @()
        }

        if (@($versionHistory).Length -gt $majorVersionCount)
        {
            $outdatedVersions = $versionHistory | sort VersionLabel -Descending | select -Skip $majorVersionCount
            Write-Host "   $(@($outdatedVersions).Length) versions to remove, from $($item.FieldValues.FileRef)" -ForegroundColor Magenta

            if($parameters.whatIfMode.Value)
            {
                $outdatedVersions | % { $global:versionSize += $_.Size }
            } 
            else {

                ## remove version
                foreach ($version in $outdatedVersions) 
                {
                    $global:versionSize += $version.Size
                    Remove-PnPFileVersion -Url $item.FieldValues.FileRef -Identity $version.Id -Force
                    #Write-Host "    Removing version, $($version.VersionLabel)" -ForegroundColor DarkMagenta
                }
                Write-Host "    Removed $(@($outdatedVersions).Length) versions" -ForegroundColor DarkMagenta
            }
        }
    }
}

function DeleteOldMajorVersionsGraph($docLib,$majorVersionCount,$siteId)
{
    $graphList = Get-MgSiteList -ListId $docLib.Id -SiteId $siteId -Property "Drive" -ExpandProperty Drive

    # Get root children
    $children = Get-MgDriveRootChild -DriveId $graphList.Drive.Id -All -Property "id, Folder"

    GetChildrensChildren $children $graphList.Drive.Id $majorVersionCount
}

function GetChildrensChildren($children, $driveId, $majorVersionCount)
{
    foreach($child in $children)
    {
        if ($child.Folder.ChildCount) 
        { 
            $grandchildren = Get-MgDriveItemChild -All -DriveId $driveId -DriveItemId $child.Id -Property "id, Folder"
            GetChildrensChildren $grandchildren $driveId $majorVersionCount
        }
        else {

            # don't forget versions
            $versionHistory = Get-MgDriveItemVersion -DriveId $driveId -DriveItemId $child.Id -Property "Id, LastModifiedDateTime, Size" -All

            $majorVersionPlus1 = [int]($majorVersionCount) + 1
            if (@($versionHistory).Length -gt $majorVersionPlus1)
            {
                $outdatedVersions = $versionHistory | sort LastModifiedDateTime -Descending | select -Skip $majorVersionPlus1
                Write-Host "   $(@($outdatedVersions).Length) versions to remove, from $($child.Id)" -ForegroundColor Magenta

                if($parameters.whatIfMode.Value)
                {
                    $outdatedVersions | % { $global:versionSize += $_.Size }
                }  
                else {
                    ## remove version
                    foreach ($version in $outdatedVersions) 
                    {
                        $global:versionSize += $version.Size
                        Remove-MgDriveItemVersion -DriveId $driveId -DriveItemId $child.Id -DriveItemVersionId $version.Id
                        #Write-Host "    Removing version, $($version.Id)" -ForegroundColor DarkMagenta
                    }
                    Write-Host "    Removed $(@($outdatedVersions).Length) versions" -ForegroundColor DarkMagenta
                }
            }
        }
    }
}


##############################################
# Main
##############################################

# Install required PS Modules
# Write-Host "Installing required PowerShell Modules..." -ForegroundColor Yellow
# Install-Module "PnP.PowerShell" -Scope CurrentUser
# Install-Module Microsoft.Graph

# Load Parameters from json file
$parametersListContent = Get-Content '.\parametersVersions.json' -ErrorAction Stop

$parameters = $parametersListContent | ConvertFrom-Json

if($parameters.whatIfMode.Value)
{
    Write-Host "RUNNING IN WHATIF MODE - Nothing will be delted, library versions won't be configured" -ForegroundColor Cyan
    Write-Host "Press [Y]es to continue" -ForegroundColor Cyan
} 
else {
    Write-Host "RUNNING IN CONFIGURE & DELETE MODE - Library versions will be configured, and previous versions will be deleted" -ForegroundColor Red
    Write-Host "Press [Y]es to continue" -ForegroundColor Cyan
}

$continue = Read-Host

if($continue.ToLowerInvariant() -eq "y")
{
    $stopwatch = [System.Diagnostics.Stopwatch]::new()

    $stopwatch.Start()
    
    if ($parameters.deleteOldMajorVersions.Value)
    {
        ## Connect to graph as deleting old verison
        ConnectToMSGraph $parameters
    }
    
    $sites = Import-Csv $parameters.inputCSV.Value
    
    ProcessSites $sites
    
    $sites | Export-CSV $parameters.inputCSV.Value -NoTypeInformation -Force
    
    $stopwatch.Stop()
    
    $stopwatch.Elapsed.TotalMinutes
}