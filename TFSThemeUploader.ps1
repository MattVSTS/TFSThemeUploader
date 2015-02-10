# TFSThemeUploader.ps1 - script for automatically creating Theme Work Item Types in an existing Team Project
# matteo.emili@live.com || http://mattvsts.blogspot.com || @MattVSTS
# Basic parameters
Param(
[String]$CollectionURL,
[String]$TeamProject
)

# Execution path and backup folder
$data = Split-path $MyInvocation.MyCommand.Path
$backupdata = New-Item -Force -ItemType Directory -Path "$data\Backup"


# WITAdmin location (VS 2013)
$WitAdmin = "${env:ProgramFiles(x86)}\Microsoft Visual Studio 12.0\Common7\IDE\witadmin.exe"

$logline = "{0} - TFSThemeUploader started." -f (Get-Date).ToString("h:m:s")
Add-Content -Path "$data\log.txt" -Value $logline -Force
$logline = "{0} - Team Foundation Server: $CollectionURL" -f (Get-Date).ToString("h:m:s")
Add-Content -Path "$data\log.txt" -Value $logline -Force
$logline = "{0} - Team Project: $TeamProject" -f (Get-Date).ToString("h:m:s")
Add-Content -Path "$data\log.txt" -Value $logline -Force

try
{
    # Get files to edit
    & $WitAdmin exportwitd /collection:$CollectionUrl /p:$TeamProject /n:Feature /f:"$backupdata\bck_Feature.xml"
    & $WitAdmin exportcategories /collection:$CollectionUrl /p:$TeamProject /f:"$backupdata\bck_Categories.xml"
    & $WitAdmin exportprocessconfig /collection:$CollectionUrl /p:$TeamProject /f:"$backupdata\bck_ProcessConfiguration.xml"
    $logline = "{0} - Original Work Item Type, Categories List and Process Configuration exported." -f (Get-Date).ToString("h:m:s")
    Add-Content -Path "$data\log.txt" -Value $logline -Force
    Write-Output "Original files downloaded."
}
catch
{
    $logline = "{0} - Exception while backing up the original files." -f (Get-Date).ToString("h:m:s")
    Add-Content -Path "$data\log.txt" -Value $logline -Force
}



# Modify the Feature Work Item Type to create a Theme Work Item Type
# Change name, description and child filter
try
{
    $themeWIT = New-Object XML
    $themeWIT.Load("$backupdata\bck_Feature.xml")
    $themeWIT.WITD.WORKITEMTYPE.name = "Theme"
    $themeWIT.WITD.WORKITEMTYPE.DESCRIPTION = "Tracks a theme that is part of the overall strategy"

    foreach ($x in $themeWIT.WITD.WORKITEMTYPE.FORM.Layout.Group.Column.TabGroup.Tab.Where({$_.Label -eq "Implementation"}))
    {
        if ($x.Control.LinksControlOptions.WorkItemTypeFilters.Filter.WorkItemType -eq "User Story")
        {
            $x.Control.LinksControlOptions.WorkItemTypeFilters.Filter.WorkItemType = "Feature"
        }
        elseif ($x.Control.LinksControlOptions.WorkItemTypeFilters.Filter.WorkItemType -eq "Product Backlog Item")
        {
            $x.Control.LinksControlOptions.WorkItemTypeFilters.Filter.WorkItemType = "Feature"
        }
    }
    $themeWIT.Save("$data\Theme.xml")
    $logline = "{0} - Theme Work Item Type created." -f (Get-Date).ToString("h:m:s")
    Add-Content -Path "$data\log.txt" -Value $logline -Force
    Write-Output "Theme created."
}
catch
{
    $logline = "{0} - Exception while creating the Theme Work Item Type." -f (Get-Date).ToString("h:m:s")
    Add-Content -Path "$data\log.txt" -Value $logline -Force
}



try
{
    # Add a new Category for Themes
    $categories = New-Object XML
    $categories.Load("$backupdata\bck_Categories.xml")

    $themeCat = $categories.CATEGORIES.CATEGORY[0].Clone()
    $themeCat.refname = "Custom.ThemeCategory"
    $themeCat.name = "Theme Category"
    $themeCat.DEFAULTWORKITEMTYPE.name = "Theme"
    
    $categories.DocumentElement.AppendChild($themeCat)
    
    $categories.Save("$data\Categories.xml")
    $logline = "{0} - Custom.ThemeCategory Category created." -f (Get-Date).ToString("h:m:s")
    Add-Content -Path "$data\log.txt" -Value $logline -Force
    Write-Output "Category created."
}
catch
{
    $logline = "{0} - Exception while creating the Custom.ThemeCategory Category." -f (Get-Date).ToString("h:m:s")
    Add-Content -Path "$data\log.txt" -Value $logline -Force
}

try
{
    # Add a new Portfolio Backlog
    # Set the Feature's PB parent attribute as it is no longer the topmost one
    # Set the Theme's Portfolio Backlog attributes
    # Create a new Work Item Color visualisation for Themes
    $processconfig = New-Object XML  
    $processconfig.Load("$backupdata\bck_ProcessConfiguration.xml")  

    $themePB = $processconfig.ProjectProcessConfiguration.PortfolioBacklogs.PortfolioBacklog | Where-Object {$_.category -eq 'Microsoft.FeatureCategory'} 
    $featurePB = $themePB.Clone() 
    $featurePB.SetAttribute("parent", "Custom.ThemeCategory") 
    $processconfig.ProjectProcessConfiguration.PortfolioBacklogs.AppendChild($featurePB) 
    $themePB.category = "Custom.ThemeCategory"  
    $themePB.pluralName = "Themes"  
    $themePB.singularName = "Theme"

    $basecolour = $processconfig.ProjectProcessConfiguration.WorkItemColors.WorkItemColor | Where-Object {$_.name -eq 'Feature'} 
    $themecolour = $basecolour.Clone() 
    $themecolour.SetAttribute("primary", "33CC33FF") 
    $themecolour.SetAttribute("secondary", "009900FF")
    $themecolour.SetAttribute("name", "Theme")
    $processconfig.ProjectProcessConfiguration.WorkItemColors.AppendChild($themecolour)

    $processconfig.Save("$data\ProcessConfiguration.xml") 
    $logline = "{0} - Theme Portfolio Backlog created." -f (Get-Date).ToString("h:m:s")
    Add-Content -Path "$data\log.txt" -Value $logline -Force
    Write-Output "Process Configuration updated."
}
catch
{
    $logline = "{0} - Exception while creating the Theme Portfolio Backlog." -f (Get-Date).ToString("h:m:s")
    Add-Content -Path "$data\log.txt" -Value $logline -Force
}

try
{
    # Import the new files into Team Foundation Server
    & $WitAdmin importwitd /collection:$CollectionUrl /p:$TeamProject /f:"$data\Theme.xml"
    & $WitAdmin importcategories /collection:$CollectionUrl /p:$TeamProject /f:"$data\Categories.xml"
    & $WitAdmin importprocessconfig /collection:$CollectionUrl /p:$TeamProject /f:"$data\ProcessConfiguration.xml"
    $logline = "{0} - New Work Item Type, Categories List and Process Configuration imported." -f (Get-Date).ToString("h:m:s")
    Add-Content -Path "$data\log.txt" -Value $logline -Force
    Write-Output "New files uploaded."
}
catch
{
    $logline = "{0} - Exception while uploading the new files." -f (Get-Date).ToString("h:m:s")
    Add-Content -Path "$data\log.txt" -Value $logline -Force
}